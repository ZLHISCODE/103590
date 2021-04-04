Create Or Replace Procedure Zl_Cissvr_Addadviceannex
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ�����ҽ��������Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --    bill_no                 C 1 ����:No
  --    bill_prop               N 1 ����:��¼����
  --    advice_id               N 1 ҽ��ID
  --    send_no                 N 1 ���ͺ�

  --����: Json_Out,��ʽ����
  --  output
  --    code                      N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                   C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Json     Pljson;
  j_In       Pljson;
  v_No       ����ҽ������.No%Type;
  n_��¼���� ����ҽ������.��¼����%Type;
  n_���ͺ�   ����ҽ������.���ͺ�%Type;
  n_ҽ��id   ����ҽ������.ҽ��id%Type;

Begin
  --�������
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  n_ҽ��id   := j_Json.Get_Number('advice_id');
  n_���ͺ�   := j_Json.Get_Number('send_no');
  v_No       := j_Json.Get_String('bill_no');
  n_��¼���� := j_Json.Get_Number('bill_prop');

  If Nvl(n_ҽ��id, 0) = 0 Or Nvl(n_���ͺ�, 0) = 0 Then
    Json_Out := Zljsonout('δ����ҽ��id���ͺţ����飡');
    Return;
  End If;

  If Nvl(n_��¼����, 0) = 0 Then
    Json_Out := Zljsonout('δ����¼���ʣ����飡');
    Return;
  End If;

  If Nvl(v_No, '-') = '-' Then
    Json_Out := Zljsonout('δ����NO�����飡');
    Return;
  End If;

  Zl_����ҽ������_Insert(n_ҽ��id, n_���ͺ�, n_��¼����, v_No);

  Json_Out := Zljsonout('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Addadviceannex;
/
Create Or Replace Procedure Zl_Cissvr_Adviceannex_Add
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:����ҽ��������Ϣ
  --��Σ�Json_In:��ʽ
  --    input
  --      advice_send_no  C  1  ����:No
  --      advice_send_properties  N  1  ����:��¼����
  --      advice_id  N  1  ҽ��ID
  --      advice_send_number  N  1  ���ͺ�
  --����: Json_Out,��ʽ����
  --  output
  --    code              N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message           C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------

  j_Json Pljson;
  j_In   Pljson;
  v_No   ����ҽ������.No%Type;

  n_��¼���� ����ҽ������.��¼����%Type;
  n_ҽ��id   ����ҽ������.ҽ��id%Type;

  n_���ͺ� ����ҽ������.���ͺ�%Type;

Begin
  --�������
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');

  v_No       := j_Json.Get_String('advice_send_no');
  n_��¼���� := j_Json.Get_Number('advice_send_properties');
  n_ҽ��id   := j_Json.Get_Number('advice_id');
  n_���ͺ�   := j_Json.Get_Number('advice_send_number');

  If v_No Is Null Or Nvl(n_��¼����, 0) = 0 Or Nvl(n_ҽ��id, 0) = 0 Or Nvl(n_���ͺ�, 0) = 0 Then
    Json_Out := Zljsonout('���븽����Ϣ��������');
    Return;
  End If;

  Zl_����ҽ������_Insert(n_ҽ��id, n_���ͺ�, n_��¼����, v_No);

  Json_Out := Zljsonout('�ɹ�', 1);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Adviceannex_Add;
/
Create Or Replace Procedure Zl_Cissvr_Adviceexecuting
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ָ����ҽ���Ƿ�����ִ��
  --��Σ�Json_In:��ʽ
  --input
  --    item_list
  --      advice_id               C  1  ҽ��ID
  --      advice_send_no          C  1  ���͵���
  --      advice_send_properties  N  1  ��¼����
  --����: Json_Out,��ʽ����
  --    output
  --        code                  N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --        message               C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --        isexist               N   0   �Ƿ���ڣ�1-����;0-������
  ---------------------------------------------------------------------------
  j_Json        Pljson;
  j_Jsonlist_In Pljson_List;

  j_Json_Tmp Pljson;

  v_No       ����ҽ������.No%Type;
  n_��¼���� ����ҽ������.��¼����%Type;
  n_ҽ��id   ����ҽ������.ҽ��id%Type;
  n_����     Number(2);

  n_Count Number(18);
Begin
  --�������
  j_Json_Tmp    := Pljson(Json_In);
  j_Json        := j_Json_Tmp.Get_Pljson('input');
  j_Jsonlist_In := j_Json.Get_Pljson_List('item_list');
  n_����        := 0;
  If Not j_Jsonlist_In Is Null Then
    n_Count := j_Jsonlist_In.Count;
    For I In 1 .. n_Count Loop
      j_Json_Tmp := Pljson();
      j_Json_Tmp := Pljson(j_Jsonlist_In.Get(I));
      n_ҽ��id   := j_Json_Tmp.Get_Number('advice_id');
      v_No       := j_Json_Tmp.Get_String('advice_send_no');
      n_��¼���� := j_Json_Tmp.Get_Number('advice_send_properties');
    
      --�������������̵ģ������ҽ��ִ��״̬
      Select Nvl(Count(1), 0)
      Into n_Count
      From ����ҽ������
      Where ִ��״̬ = 3 And NO = v_No And ��¼���� = Nvl(n_��¼����, 0) And ҽ��id = Nvl(n_ҽ��id, 0);
      If n_Count > 0 Then
        n_���� := 1;
        Exit;
      End If;
    End Loop;
  End If;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","isexist":' || n_���� || '}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Adviceexecuting;
/
Create Or Replace Procedure Zl_Cissvr_Adviceexistitem
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  ---------------------------------------------------------------------------
  --���ܣ���ȡ����ҽ����Ϣ���Ƿ����ָ�����շ���Ŀ
  --��Σ�Json_In:��ʽ
  --input   
  --       advice_item_id        N  1  �շ���Ŀid
  --����: Json_Out,��ʽ����
  --output
  --       code                  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --       message               C  1  Ӧ����Ϣ��
  --       item_exits            N  1  �Ƿ���ڣ�0-�����ڣ�1-����
  ---------------------------------------------------------------------------
  j_Input      Pljson;
  j_Json       Pljson;
  n_�շ���Ŀid Number(18);
  n_Count      Number(18);
Begin
  j_Input      := Pljson(Json_In);
  j_Json       := j_Input.Get_Pljson('input');
  n_�շ���Ŀid := j_Json.Get_Number('advice_item_id');
  Select Count(1) Into n_Count From ����ҽ����¼ Where �շ�ϸĿid = n_�շ���Ŀid And Rownum < 2;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_exits":' || Nvl(n_Count, 0) || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Adviceexistitem;
/
Create Or Replace Procedure Zl_Cissvr_Adviceishistory
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:���ҽ���Ƿ��Ѿ�������ʷ���ݱ���
  --��Σ�Json_In:��ʽ
  --    input
  --       advice_list[]     ����
  --             advice_id           N 1 ҽ��ID
  --             send_no             N 1 ���ͺ�

  --����: Json_Out,��ʽ����
  --  output
  --    code                  N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message               C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    is_history            N   1   �Ƿ����:1-����;0-������
  ---------------------------------------------------------------------------

  j_Json    Pljson;
  o_Json    Pljson;
  j_List    Pljson_List := Pljson_List();
  n_ҽ��id  ����ҽ������.ҽ��id%Type;
  n_���ͺ�  ����ҽ������.���ͺ�%Type;
  n_Count   Number(18);
  n_Exist   Number(1);
  n_Exist_h Number(1);

Begin
  --�������
  o_Json    := Pljson(Json_In);
  j_Json    := o_Json.Get_Pljson('input');
  j_List    := j_Json.Get_Pljson_List('advice_list');
  n_Exist   := 0;
  n_Exist_h := 0;
  If Not j_List Is Null Then
    n_Count := j_List.Count;
    For I In 1 .. n_Count Loop
      o_Json   := Pljson();
      o_Json   := Pljson(j_List.Get(I));
      n_ҽ��id := o_Json.Get_Number('advice_id');
      n_���ͺ� := o_Json.Get_Number('send_no');
      If Nvl(n_ҽ��id, 0) = 0 Or Nvl(n_���ͺ�, 0) = 0 Then
        Json_Out := Zljsonout('δ����ҽ��id���ͺţ����飡');
        Return;
      End If;
      Select Max(1) Into n_Exist From ����ҽ������ Where ҽ��id = n_ҽ��id And ���ͺ� = n_���ͺ�;
      If Nvl(n_Exist, 0) = 0 Then
        Select Max(1) Into n_Exist_h From H����ҽ������ Where ҽ��id = n_ҽ��id And ���ͺ� = n_���ͺ�;
        If Nvl(n_Exist_h, 0) = 1 Then
          Exit;
        End If;
      End If;
    End Loop;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","is_history":' || Nvl(n_Exist_h, 0) || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Adviceishistory;
/
CREATE OR REPLACE Procedure Zl_Cissvr_Adviceisinvalid
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ�����ҽ��ID��ѯҽ��״̬
  --��Σ�Json_In:��ʽ
  --input 
  --   advice_ids           C  1  ���ҽ��ID����,�ָ�
  --����: Json_Out,��ʽ����
  --output
  --    code                 N  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message              C  1  Ӧ����Ϣ
  --    advice_ids           C  1  ҽ��ID�������ϵģ�
  ---------------------------------------------------------------------------
  v_ҽ��id Clob; --��¼ҽ��id
  v_Tmp    Varchar2(32767); --��Ϊ�м����
  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_ҽ��id Collection_Type;
  I          Number;
  j_In       Pljson;
  j_Json     Pljson;
Begin
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  v_ҽ��id := j_Json.Get_Clob('advice_ids');
  --�� v_ҽ��id ����װ�ɲ�����4000 �ļ��ϴ�����ֹʹ�� f_Num2list ��������
  I := 0;
  While v_ҽ��id Is Not Null Loop
    If Length(v_ҽ��id) <= 4000 Then
      Col_ҽ��id(I) := v_ҽ��id;
      v_ҽ��id := Null;
    Else
      Col_ҽ��id(I) := Substr(v_ҽ��id, 1, Instr(v_ҽ��id, ',', 3980) - 1);
      v_ҽ��id := Substr(v_ҽ��id, Instr(v_ҽ��id, ',', 3980) + 1);
    End If;
    I := I + 1;
  End Loop;
  I     := 0;
  v_Tmp := Null;
  For I In 0 .. Col_ҽ��id.Count - 1 Loop
    For v_ҽ������ In (Select /*+cardinality(b,10)*/
                   Distinct ID
                   From ����ҽ����¼ A, Table(f_Num2list(Col_ҽ��id(I))) B
                   Where Nvl(ҽ��״̬, 0) = 4 And ID = Column_Value) Loop
    
      v_Tmp := v_Tmp || ',' || v_ҽ������.Id;
    End Loop;
  End Loop;
  Json_Out := '{"output":{"code":1,"message": "�ɹ�","advice_ids":"' || Substr(v_Tmp, 2) || '"}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Adviceisinvalid;
/
Create Or Replace Procedure Zl_Cissvr_Auditadvicecharge
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:��ָ��ҽ�����з���������
  --��Σ�Json_In:��ʽ
  -- input
  --   advice_id            N 1 ҽ��id
  --   verfy_statu          N 1 ���״̬:1-���;0-ȡ�����
  --����: Json_Out,��ʽ����
  --  output
  --    code               N  1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message            C  1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Json     Pljson;
  j_In       Pljson;
  n_ҽ��id   ����ҽ����¼.Id%Type;
  n_������� ����ҽ����¼.�Ƿ�������%Type;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --�������
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  n_ҽ��id   := j_Json.Get_Number('advice_id');
  n_������� := Nvl(j_Json.Get_Number('verfy_statu'), 0);

  If Nvl(n_ҽ��id, 0) = 0 Then
    v_Error := '���봫��ҽ��id��';
    Raise Err_Custom;
  End If;

  Zl_����ҽ����¼_�������(n_ҽ��id, n_�������);
  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Auditadvicecharge;
/
Create Or Replace Procedure Zl_Cissvr_AuditDrugOrder
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:����ҽ�����
  --��Σ�Json_In:��ʽ 
  --input     ����ҽ�����
  --  auditor        C  1  �����
  --  audit_content  C  1  ҽ��������ݣ���ʽ������ID1,����1,˵��1||ID2,����2,˵��2��

  --����: Json_Out,��ʽ����
  --output
  --  code           N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message        C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Json     Pljson;
  j_In       Pljson;
  v_�����   Varchar2(20);
  v_������� Varchar2(32767);
  v_Field    Varchar2(32767);
  v_Tmp      Varchar2(32767);
  n_ҽ��id   Number(18);
  n_����     Number(2);
  v_���ԭ�� Varchar2(100);
  Err_Custom Exception;
  v_Err Varchar2(255);
Begin
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');

  v_�����   := j_Json.Get_String('auditor');
  v_������� := j_Json.Get_String('audit_content');

  If v_������� Is Null Then
    v_Err := 'δ������ϸ��Ϣ��audit_content���ڵ�';
    Raise Err_Custom;
    Return;
  End If;

  v_Tmp := v_������� || '||';
  While v_Tmp Is Not Null Loop
    v_Field := Substr(v_Tmp, 1, Instr(v_Tmp, '||') - 1);
    v_Tmp   := Replace('||' || v_Tmp, '||' || v_Field || '||');
  
    n_ҽ��id := To_Number(Substr(v_Field, 1, Instr(v_Field, ',') - 1));
    v_Field  := Substr(v_Field, Instr(v_Field, ',') + 1);
  
    n_����     := To_Number(Substr(v_Field, 1, Instr(v_Field, ',') - 1));
    v_���ԭ�� := Substr(v_Field, Instr(v_Field, ',') + 1);
  
    If v_����� Is Null And n_���� <> 0 Then
      Update ����ҽ����¼
      Set ҩʦ��˱�־ = n_����, ҩʦ���ʱ�� = Sysdate, ҩʦ���ԭ�� = v_���ԭ��
      Where ID = n_ҽ��id And Nvl(ҩʦ��˱�־, 0) = 0;
    Else
      If n_���� = 0 Then
        Update ����ҽ����¼
        Set ҩʦ��˱�־ = n_����, ҩʦ���ʱ�� = Null, ���ҩʦ = Null, ҩʦ���ԭ�� = v_���ԭ��
        Where ID = n_ҽ��id;
      Elsif n_���� = 3 Then
        --���ŷ�ҩ���ҽ��������˹��Ĳ��������
        Update ����ҽ����¼
        Set ҩʦ��˱�־ = n_����, ҩʦ���ʱ�� = Sysdate, ���ҩʦ = v_�����, ҩʦ���ԭ�� = v_���ԭ��
        Where Nvl(ҩʦ��˱�־, 0) = 0 And ID = n_ҽ��id;
      Else
        Update ����ҽ����¼
        Set ҩʦ��˱�־ = n_����, ҩʦ���ʱ�� = Sysdate, ���ҩʦ = v_�����, ҩʦ���ԭ�� = v_���ԭ��
        Where ID = n_ҽ��id;
      End If;
    End If;
  End Loop;

  Json_Out := '{"output":{"code": 1,"message":"�ɹ�"}}';
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_AuditDrugOrder;
/

Create Or Replace Procedure Zl_Cissvr_Buildadviceexecharge
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:��������ҽ��ִ�мƼ�����
  --��Σ�Json_In:��ʽ
  --    input
  --    pati_id             N 1 ����id
  --    pati_pageid         N 1 ��ҳid
  --    bill_no             C 0 ���ݺ�
  --    bill_prop           N 0 ��¼����
  --����: Json_Out,��ʽ����
  --  output
  --    code                N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message             C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------

  j_In       Pljson;
  j_Json     Pljson;
  n_����id   ����ҽ����¼.����id%Type;
  n_��ҳid   ����ҽ����¼.��ҳid%Type;
  v_No       ����ҽ������.No%Type;
  n_��¼���� ����ҽ������.��¼����%Type;

Begin
  --�������
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');
  n_��ҳid := j_Json.Get_Number('pati_pageid');

  v_No       := j_Json.Get_String('bill_no');
  n_��¼���� := j_Json.Get_Number('bill_prop');

  If v_No Is Null And Nvl(n_����id, 0) <> 0 Then
    Json_Out := Zljsonout('δ������Ҫ���µĵ�����Ϣ������Ϣ��');
    Return;
  End If;

  If Nvl(n_����id, 0) <> 0 Then
    --������ID���д���
  
    For c_��¼ In (Select Distinct b.ҽ��id, b.No, b.��¼����
                 From ����ҽ����¼ A, ����ҽ������ B, ҽ��ִ�мƼ� C
                 Where a.����id = n_����id And (a.��ҳid = Nvl(n_��ҳid, 0) Or n_��ҳid Is Null) And a.Id = b.ҽ��id And
                       b.ҽ��id = c.ҽ��id And b.���ͺ� = c.���ͺ� And
                       ((b.No = v_No And b.��¼���� = Nvl(n_��¼����, 0) Or n_��¼���� Is Null Or v_No Is Null)) And c.ִ��״̬ Is Null
                 Order By b.No) Loop
    
      Zl_ҽ��ִ�мƼ�_����(c_��¼.ҽ��id, c_��¼.No, c_��¼.��¼����);
    End Loop;
  End If;

  For c_��¼ In (Select Distinct b.ҽ��id, b.No
               From ����ҽ������ B, ҽ��ִ�мƼ� C
               Where b.No = v_No And b.��¼���� = n_��¼���� And c.ִ��״̬ Is Null
               Order By b.No) Loop
  
    Zl_ҽ��ִ�мƼ�_����(c_��¼.ҽ��id, c_��¼.No, n_��¼����);
  End Loop;

  Json_Out := Zljsonout('�ɹ�', 1);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Buildadviceexecharge;
/
Create Or Replace Procedure Zl_Cissvr_Checkbabylimit
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ת��Ӥ��Ӥ���Ƿ�����¼����
  --��Σ�Json_In:��ʽ
  --  input
  --    pati_id             N 1 ����id
  --    pati_pageid         N 1 ��ҳid
  --    baby_num            N 1 Ӥ�����
  --����: Json_Out,��ʽ����
  --  output
  --    code                N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    in_time             C 1 ��Ժʱ��
  --    baby_wardarea_id    N 1 Ӥ������id
  --    have_data           N 1 �Ƿ��м�¼
  ---------------------------------------------------------------------------
  j_Json       Pljson;
  j_In         Pljson;
  v_��Ժ����   Varchar2(200);
  n_Ӥ������id Number;
  n_Havedata   Number := 0;
Begin
  --�������
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');
  For R In (Select b.��Ժ����, Nvl(a.Ӥ������id, 0) As Ӥ������id
            From ������������¼ A, ������ҳ B
            Where a.Ӥ������id = b.����id(+) And a.Ӥ����ҳid = b.��ҳid(+) And a.����id = j_Json.Get_Number('pati_id') And
                  a.��ҳid = j_Json.Get_Number('pati_pageid') And a.��� = j_Json.Get_Number('baby_num')) Loop
    n_Havedata   := 1;
    v_��Ժ����   := To_Char(r.��Ժ����, 'yyyy-mm-dd hh24:mi:ss');
    n_Ӥ������id := r.Ӥ������id;
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","in_time":"' || v_��Ժ���� || '","baby_wardarea_id":' || Nvl(n_Ӥ������id, 0) ||
              ',"have_data":' || n_Havedata || '}}';

Exception
  When Others Then 
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Checkbabylimit;
/
Create Or Replace Procedure Zl_Cissvr_Checkdepositerrorno
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ݵ��ݺŻ�ȡ���ڲ��˽����쳣��¼�е�NO
  --��Σ�Json_In:��ʽ
  --  input
  --   pati_id            N 1 ����id
  --   bill_nos           C 1 ����Ԥ����¼.NO,����ö��ŷָ�
  --����: Json_Out,��ʽ����
  --  output
  --    code              N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message           C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    bill_nos          C 1 ��Ч��Nos,����ö��ŷָ�
  ---------------------------------------------------------------------------
  n_����id  ���˽����쳣��¼.����id%Type;
  v_Nos     Varchar2(32767);
  j_Json    Pljson;
  j_In      Pljson;
  v_Nos_Out Varchar2(32767);

Begin
  --�������
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_����id := Nvl(j_Json.Get_Number('pati_id'), 0);
  v_Nos    := j_Json.Get_String('bill_nos');

  If Nvl(v_Nos, '-') = '-' Then
    Json_Out := Zljsonout('δ����NO�����飡');
    Return;
  End If;

  Select /*+cardinality(B,10)*/
   f_List2str(Cast(Collect(b.Column_Value) As t_Strlist))
  Into v_Nos_Out
  From ���˽����쳣��¼ A, Table(f_Str2list(v_Nos)) B
  Where a.�������� = 3 And (a.Ԥ������ = b.Column_Value Or a.ҽ�ƿ����� = b.Column_Value) And Decode(n_����id, 0, 0, a.����id) = n_����id;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","bill_nos":"' || v_Nos_Out || '"}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Checkdepositerrorno;
/
Create Or Replace Procedure Zl_Cissvr_Checkpatexist
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ�����Ƿ���ڲ���
  --��Σ�JSON��ʽ
  --input
  --   pati_id       N 1 ����id
  --   visit_id   N 1 ����id,���ﲡ��Ϊ�Һ�ID;סԺ����Ϊ��ҳID;Ϊ0˵����������ǹҺž���Ĳ���
  --   occasion      N 1 ����,1-����;2-סԺ
  --   pati_name     C 1 ����
  --   pati_sex      C 1 �Ա�
  --   pati_age      C 1 ����
  --   pati_birthdate C 1 ��������
  --���Σ�JSON��ʽ
  --output
  --  code            N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message        C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --   pati_name     C 1 ����
  --   pati_sex      C 1 �Ա�
  --   pati_age      C 1 ����
  --   pati_birthdate C 1 ��������
  ---------------------------------------------------------------------------
  n_����id   ������ҳ.����id%Type;
  n_����id   Number;
  n_����     Number;
  v_����     ������ҳ.����%Type;
  v_�Ա�     ������ҳ.�Ա�%Type;
  v_����     ������ҳ.����%Type;
  d_�������� Date;
  j_In       Pljson;
  v_Error    Varchar2(2000);
  j_Json     Pljson;
Begin
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  n_����id   := j_Json.Get_Number('pati_id');
  n_����id   := j_Json.Get_Number('visit_id');
  n_����     := j_Json.Get_Number('occasion');
  v_����     := j_Json.Get_String('pati_name');
  v_�Ա�     := j_Json.Get_String('pati_sex');
  v_����     := j_Json.Get_String('pati_age');
  d_�������� := To_Date(j_Json.Get_String('pati_birthdate'), 'yyyy-mm-dd hh24:mi:ss');
  Case n_����
    When 1 Then
      Begin
        Select b.����, b.�Ա�, b.����, a.��������
        Into v_����, v_�Ա�, v_����, d_��������
        From (Select n_����id As ����id, v_���� As ����, v_�Ա� As �Ա�, v_���� As ����, d_�������� As ��������
               From Dual) A, ���˹Һż�¼ B
        Where a.����id = b.����id And a.����id = n_����id And b.Id = n_����id;
      Exception
        When Others Then
          v_Error := '����ID[' || n_����id || ']���Һ�ID[' || n_����id || ']�ڲ��˹Һż�¼�в�����,���ܼ������в�����Ϣ�������!';
      End;
    When 2 Then
      Begin
        Select Nvl(b.����, a.����), Nvl(b.�Ա�, a.�Ա�), b.����, a.��������
        Into v_����, v_�Ա�, v_����, d_��������
        From (Select n_����id As ����id, v_���� As ����, v_�Ա� As �Ա�, v_���� As ����, d_�������� As ��������
               From Dual) A, ������ҳ B
        Where a.����id = b.����id And a.����id = n_����id And b.��ҳid = n_����id;
      Exception
        When Others Then
          v_Error := '����ID[' || n_����id || ']����ҳID[' || n_����id || ']�ڲ�����ҳ�в�����,���ܼ������в�����Ϣ�������!';
      End;
    Else
      v_Error := '���̲���[����]ֻ��Ϊ1��2,���ܼ������в�����Ϣ�������!';
  End Case;
  If v_Error Is Not Null Then
    Json_Out := Zljsonout(v_Error, 0);
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�"';
    Json_Out := Json_Out || ',"pati_name":"' || Zljsonstr(v_����, 0) || '"';
    Json_Out := Json_Out || ',"pati_age":"' || Zljsonstr(v_����, 0) || '"';
    Json_Out := Json_Out || ',"pati_sex":"' || Zljsonstr(v_�Ա�, 0) || '"';
    Json_Out := Json_Out || ',"pati_birthdate":"' || Zljsonstr(To_Char(d_��������, 'yyyy-mm-dd hh24:mi:ss'), 0) || '"' || '}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Checkpatexist;
/
Create Or Replace Procedure Zl_Cissvr_Checkpaticatalogue
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --���ܣ��жϲ����Ƿ��Ŀ
  --��Σ�Json_In:��ʽ
  --input
  --  pati_id           N    1 ����id
  --  pati_pageid       N    1��ҳid
  --���Σ�JSON��ʽ
  --output
  --    code                N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    isexist             N  1 �Ƿ��Ŀ
  -------------------------------------------
  j_Json      Pljson;
  j_In        Pljson;
  n_����id    Number(18);
  n_��ҳid    Number;
  d_Catalogue Date;
  n_Isexist   Number;
Begin
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');
  n_��ҳid := j_Json.Get_Number('pati_pageid');

  Begin
    Select ��Ŀ���� Into d_Catalogue From ������ҳ Where ����id = n_����id And ��ҳid = n_��ҳid;
  Exception
    When Others Then
      Null;
  End;
  If d_Catalogue Is Null Then
    n_Isexist := 0;
  Else
    n_Isexist := 1;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","isexist":' || n_Isexist || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Checkpaticatalogue;
/
Create Or Replace Procedure Zl_Cissvr_Checkpatiexecute
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ݲ�����Ϣ��ȡҽ��δִ�е���Ŀ
  --��Σ�Json_In:��ʽ
  --  input
  --     pati_id              N 1 ����ID
  --     pati_pageid          N 1 ��ҳID
  --     baby_num             N 0 Ӥ�����:-1��ʾ������;0-ĸ�׵�;>0����Ӥ������
  --     fee_source           N 1 ������Դ:1-����;2-סԺ;4-���
  --     check_type           N 0 ������ͣ�null/0-��ʾ���������ִ����Ŀ��1-��ʾ���δ��Ѫ
  --����: Json_Out,��ʽ����
  --  output
  --    code                  N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message               C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    isexist               N 1 �Ƿ����: 1-����;0-������
  --    notexecute_infor      C 1 δִ�е���Ŀ��Ϣ,isexist=1ʱ����
  ---------------------------------------------------------------------------
  j_Json     Pljson;
  j_In       Pljson;
  n_����id   Number;
  n_��ҳid   Number;
  n_Ӥ����� Number;
  n_������Դ Number;
  n_��鷽ʽ Number;
  n_����     Number;
  Type t_Bool Is Ref Cursor;
  c_Bool t_Bool;
  v_��Ŀ Varchar2(32767);
  v_���� Varchar2(32767);
  v_���� Varchar2(100);

  v_Sql   Varchar2(32767);
  v_Pars  Varchar2(4000);
  v_Text  Varchar2(4000);
  n_Count Number(18);
Begin
  --�������
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  n_����id   := j_Json.Get_Number('pati_id');
  n_��ҳid   := j_Json.Get_Number('pati_pageid');
  n_Ӥ����� := j_Json.Get_Number('baby_num');
  n_������Դ := j_Json.Get_Number('fee_source');
  n_��鷽ʽ := j_Json.Get_Number('check_type');
  If Nvl(n_����id, 0) = 0 Then
    Json_Out := '{"output":{"code":0,"message":"δ���벡����Ϣ��"}}';
    Return;
  End If;

  n_Count := 0;
  If Nvl(n_��鷽ʽ, 0) = 0 Then
    --1.ҽ������ִ�е���Ŀ,�ٴ�����
    --2.������������Ŀ�����ִ��
    --3.PACS�ѱ�����(ִ�й���Ϊ">=2-�����"����Ϊδִ����ɵ���Ŀ
    If Nvl(n_������Դ, 0) = 2 Then
      Select zl_GetSysParameter(234) Into v_Pars From Dual;
      v_Pars := Replace(v_Pars, '|', ',');
      For r_Info In (Select Distinct b.No, c.���� As ��Ŀ, d.���� As ����, Decode(Nvl(b.ִ��״̬, 0), 0, 'δִ��', 3, '����ִ��') As ִ��״̬
                     From ����ҽ����¼ A, ����ҽ������ B, ������ĿĿ¼ C, ���ű� D
                     Where a.����id = n_����id And Nvl(a.��ҳid, 0) = n_��ҳid And (Nvl(a.Ӥ��, 0) = n_Ӥ����� Or n_Ӥ����� = -1) And
                           a.Id = b.ҽ��id And b.ִ��״̬ In (0, 3) And a.������Ŀid = c.Id And b.ִ�в���id + 0 = d.Id And
                           a.������� Not In ('4', '5', '6', '7') And Not (a.������� In ('F', 'D') And a.���id Is Not Null) And
                           (d.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or d.����ʱ�� Is Null) And
                           Not (a.������� = 'D' And Nvl(b.ִ�й���, 0) >= 2) And
                           (Not (a.������� = 'Z' And Nvl(c.��������, '0') <> '0') Or a.������� = 'Z' And c.�������� = '7') And
                           c.Id Not In (Select /*+cardinality(j,10) */
                                         Column_Value
                                        From Table(Cast(f_Num2list(v_Pars) As Zltools.t_Numlist)) J)) Loop
        If Lengthb(v_Text || Chr(13) || Chr(10) || '����[' || Nvl(r_Info.No, '') || ']�е�' || Nvl(r_Info.��Ŀ, '') || '����' ||
                   Nvl(r_Info.����, '[δ������]') || r_Info.ִ��״̬) > 1000 Then
          v_Text := v_Text || Chr(13) || Chr(10) || '... ...';
          Exit;
        Else
          v_Text := v_Text || Chr(13) || Chr(10) || '����[' || Nvl(r_Info.No, '') || ']�е�' || Nvl(r_Info.��Ŀ, '') || '����' ||
                    Nvl(r_Info.����, '[δ������]') || r_Info.ִ��״̬;
        End If;
      
        n_Count := n_Count + 1;
      End Loop;
    
      v_Text := Substr(v_Text, 3);
    Else
      For r_Info In (Select Distinct b.No, c.���� As ��Ŀ, d.���� As ����, Decode(Nvl(b.ִ��״̬, 0), 0, 'δִ��', 3, '����ִ��') As ִ��״̬
                     From ����ҽ����¼ A, ����ҽ������ B, ������ĿĿ¼ C, ���ű� D
                     Where a.����id = n_����id And a.��ҳid Is Null And (Nvl(a.Ӥ��, 0) = n_Ӥ����� Or n_Ӥ����� = -1) And
                           a.Id = b.ҽ��id And b.ִ��״̬ In (0, 3) And a.������Ŀid = c.Id And b.ִ�в���id + 0 = d.Id And
                           a.������� Not In ('4', '5', '6', '7') And Not (a.������� In ('F', 'D') And a.���id Is Not Null) And
                           (d.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or d.����ʱ�� Is Null) And
                           Not (a.������� = 'D' And Nvl(b.ִ�й���, 0) >= 2) And
                           (Not (a.������� = 'Z' And Nvl(c.��������, '0') <> '0') Or a.������� = 'Z' And c.�������� = '7')) Loop
        If Lengthb(v_Text || Chr(13) || Chr(10) || '����[' || Nvl(r_Info.No, '') || ']�е�' || Nvl(r_Info.��Ŀ, '') || '����' ||
                   Nvl(r_Info.����, '[δ������]') || r_Info.ִ��״̬) > 1000 Then
          v_Text := v_Text || Chr(13) || Chr(10) || '... ...';
          Exit;
        Else
          v_Text := v_Text || Chr(13) || Chr(10) || '����[' || Nvl(r_Info.No, '') || ']�е�' || Nvl(r_Info.��Ŀ, '') || '����' ||
                    Nvl(r_Info.����, '[δ������]') || r_Info.ִ��״̬;
        End If;
      
        n_Count := n_Count + 1;
      End Loop;
    
      v_Text := Substr(v_Text, 3);
    End If;
  
    If n_Count > 0 Then
      n_Count := 1;
    End If;
  Elsif n_��鷽ʽ = 1 Then
    v_Text := '';
    --����Ƿ�װ��Ѫ��
    Begin
      Select 1
      Into n_����
      From zlSystems
      Where Trunc(��� / 100) = 22 And ������ = Sys_Context('USERENV', 'CURRENT_SCHEMA');
    Exception
      When Others Then
        n_���� := 0;
    End;
    v_Sql := 'select e.���� ��Ŀ, c.���� As ����, To_Char(a.����) As ���� ';
    v_Sql := v_Sql || ' from  ���ű� c,ѪҺ�շ���¼ a, �շ���ĿĿ¼ e, ѪҺ��Ѫ��¼ b, ����ҽ����¼ d ';
    v_Sql := v_Sql || ' where b.����id = d.id  And d.������� = :1 and d.ҽ��״̬<>4 and c.Id = a.�ⷿid and Nvl(a.��д����, 0) <> 0 And a.���� = 6 And Mod(a.��¼״̬, 3) = 1 ';
    v_Sql := v_Sql || ' And a.��Ѫ״̬ = 1 And a.����� Is Null ';
    v_Sql := v_Sql || ' and b.����id = :2 and b.��ҳid = :3 And a.ѪҺid = e.Id And a.�䷢id = b.Id and b.��¼���� + 0 = 1 And b.��¼״̬ in (1,2)';
    --����װ������δ��ѪҺ�ļ��
    If n_���� = 1 Then
      Open c_Bool For v_Sql
        Using 'K',n_����id, n_��ҳid;
      Loop
        Fetch c_Bool
          Into  v_��Ŀ, v_����, v_����;
        Exit When c_Bool%NotFound;

        If v_Text Is Not Null Then
          If Instr(Chr(13) || Chr(10) || v_Text || Chr(13) || Chr(10),
                   Chr(13) || Chr(10) || Nvl(v_��Ŀ, '') || '����' || Nvl(v_����, '[δ��Ѫ��]') ||
                    'δ��Ѫ' || Chr(13) || Chr(10), 1) = 0 Then
            If Lengthb(v_Text || Chr(13) || Chr(10) ||  Nvl(v_��Ŀ, '') || '����' ||
                       Nvl(v_����, '[δ��Ѫ��]') || 'δ��Ѫ') <= 1000 Then
              v_Text := v_Text || Chr(13) || Chr(10) ||  Nvl(v_��Ŀ, '') || '����' ||
                        Nvl(v_����, '[δ��Ѫ��]') || 'δ��Ѫ';
            Else
              v_Text := v_Text || Chr(13) || Chr(10) || '... ...';
            End If;
          End If;
        Else
          v_Text :=  Nvl(v_��Ŀ, '') || '����' || Nvl(v_����, '[δ��Ѫ��]') || 'δ��Ѫ';
        End If;
        n_Count := n_Count + 1;
      End Loop;
      Close c_Bool;
      If v_Text Is Not Null Then
        v_Text := '����δ���ŵ�ѪҺ��' || Chr(13) || Chr(10) || Chr(13) || Chr(10) || v_Text;
      End If;
      v_Text := v_Text;
    End If;
    If n_Count > 0 Then
      n_Count := 1;
    End If;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","isexist":' || n_Count || ',"notexecute_infor":"' ||
              Zljsonstr(v_Text) || '"}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Checkpatiexecute;
/
Create Or Replace Procedure Zl_Cissvr_Checkpativisitorin
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ݲ���ID�����֤�ż��ͬһ���ֻ֤�ܶ�Ӧһ����������
  --��Σ�Json_In:��ʽ
  --  input
  --    pati_id              N   1  ����ID
  --    pati_pageid          N   1  ��ҳid
  --����: Json_Out,��ʽ����
  --  output
  --    code                 N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message              C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    isexist              N   1   �Ƿ���ڣ�0-������ 1-���ڣ�
  ---------------------------------------------------------------------------
  n_����id   ������ҳ.����id%Type;
  n_��ҳid   ������ҳ.��ҳid%Type;
  d_��Ժ���� ������ҳ.��Ժ����%Type;
  n_Count    Number;
  j_Json     Pljson;
  j_In       Pljson;
  n_Isexist  Number(1);
Begin
  --�������
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');
  n_��ҳid := j_Json.Get_Number('pati_pageid');
  If Not n_��ҳid Is Null Then
    Select ��Ժ���� Into d_��Ժ���� From ������ҳ Where ����id = n_����id And ��ҳid = n_��ҳid;
    --���ڳ�Ժʱ�䣬���жϸó�Ժ���Ƿ���ھ����סԺ����
    If Not d_��Ժ���� Is Null Then
      --���ж�סԺ
      Select Count(1) Into n_Count From ������ҳ Where ����id = n_����id And ��Ժ���� >= d_��Ժ����;
      If n_Count = 0 Then
        Begin
          --�ù��̲�������׼����С�����ϵͳ��������װû�в��˹Һż�¼
          Execute Immediate 'Select Count(1) From ���˹Һż�¼ Where ����id =:1  And �Ǽ�ʱ�� >=:2 '
            Into n_Count
            Using n_����id, d_��Ժ����;
        Exception
          When Others Then
            Null;
        End;
      End If;
    End If;
  End If;
  If Nvl(n_Count, 0) > 0 Then
    n_Isexist := 1;
  Else
    n_Isexist := 0;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","isexist":' || n_Isexist || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Checkpativisitorin;
/
Create Or Replace Procedure Zl_Cissvr_Checkskinresult
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ��ж�ҽ���Ƿ�����Ƥ�Խ��
  --��Σ�Json_In:��ʽ
  --  input
  --     advice_id          N 1 ҽ��id
  --����: Json_Out,��ʽ����
  --  output
  --    code                N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    skintest_info       N 1 Ƥ�Խ�� -1��ʾ����Ƥ�Ի����ԣ�0��ʾ��δ���Ƥ�Խ����δ�´�Ƥ��ҽ����1��ʾ���ԣ�2��ʾ����
  ---------------------------------------------------------------------------
  j_Json     Pljson;
  j_In       Pljson;
  n_Skintest Number;
  v_Test     Varchar2(3000);
  v_�Һŵ�   Varchar2(30);
  n_����id   Number;
  n_��ҳid   Number;
  v_��Ŀids  Varchar2(32767);
  n_��Ч���� Number;
  n_ҽ��id   Number;
  v_Err_Msg  Varchar(2000);
  Err_Item Exception;
Begin
  --�������
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_ҽ��id := j_Json.Get_Number('advice_id');

  n_��Ч���� := zl_GetSysParameter(70);
  n_Skintest := -1;

  For R In (Select b.�÷�id, a.����id, a.��ҳid, a.�Һŵ�
            From ����ҽ����¼ A, �����÷����� B
            Where a.������Ŀid = b.��Ŀid And b.���� = 0 And a.Id = n_ҽ��id And a.������� In ('5', '6')) Loop
    v_��Ŀids := v_��Ŀids || ',' || r.�÷�id;
    n_����id  := r.����id;
    n_��ҳid  := r.��ҳid;
    v_�Һŵ�  := r.�Һŵ�;
  End Loop;

  If v_��Ŀids Is Not Null Then
    v_��Ŀids := v_��Ŀids || ',';
    If v_�Һŵ� Is Null Then
      For X In (Select a.Ƥ�Խ��, Nvl(b.�걾��λ, '����(+);����(-)') As �걾��λ
                From ����ҽ����¼ A, ������ĿĿ¼ B
                Where a.������Ŀid = b.Id And a.����id = n_����id And a.��ҳid = n_��ҳid And
                      Instr(v_��Ŀids, ',' || a.������Ŀid || ',') > 0 And a.��ʼִ��ʱ�� >= Trunc(Sysdate) - n_��Ч����
                Order By a.��ʼִ��ʱ�� Desc) Loop
        --ֻѭ��һ��
        v_Test := x.Ƥ�Խ��;
        Exit;
      End Loop;
    Else
      For X In (Select a.Ƥ�Խ��, Nvl(b.�걾��λ, '����(+);����(-)') As �걾��λ
                From ����ҽ����¼ A, ������ĿĿ¼ B
                Where a.������Ŀid = b.Id And a.�Һŵ� = v_�Һŵ� And Instr(v_��Ŀids, ',' || a.������Ŀid || ',') > 0 And
                      a.��ʼִ��ʱ�� >= Trunc(Sysdate) - n_��Ч����
                Order By a.��ʼִ��ʱ�� Desc) Loop
        --ֻѭ��һ��
        v_Test := x.Ƥ�Խ��;
        Exit;
      End Loop;
    End If;
    If v_Test Is Null Then
      n_Skintest := 0;
    Elsif v_Test = '����' Then
      n_Skintest := -1;
    Elsif Instr(v_Test, '-') > 0 Then
      n_Skintest := 1;
    Elsif Instr(v_Test, '+') > 0 Then
      n_Skintest := 2;
    End If;
  End If;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","skintest_info":' || Nvl(n_Skintest, 0) || '}}';

Exception
  When Err_Item Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr('-20101:[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]') || '"}}';
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Checkskinresult;
/
Create Or Replace Procedure Zl_Cissvr_Chkoutpatiexistorder
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:�Һ�ʱ��鲡���ڹҺ���Ч�������Ƿ����ҽ��
  --��Σ�Json_In:��ʽ
  --    input
  --      pati_id                 N 1 ����id
  --      rgst_expidate           N 1 �Һ���Ч����
  --����: Json_Out,��ʽ����
  --  output
  --    code              N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message           C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    exist_order       N   1   �Ƿ����ҽ��
  ---------------------------------------------------------------------------

  j_Json Pljson;
  j_In   Pljson;

  n_����id   ����ҽ����¼.����id%Type;
  n_��Ч���� Number(10);
  n_Count    Number(1);

Begin
  --�������
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  n_����id   := j_Json.Get_Number('pati_id');
  n_��Ч���� := j_Json.Get_Number('rgst_expidate');

  If Nvl(n_����id, 0) = 0 Then
    Json_Out := Zljsonout('δ���벡��id�����飡');
    Return;
  End If;

  Select Count(1)
  Into n_Count
  From ���˹Һż�¼ A, ����ҽ����¼ B
  Where a.����id + 0 = b.����id And a.No || '' = b.�Һŵ� And a.��¼״̬ = 1 And a.��¼���� = 1 And
        a.�Ǽ�ʱ�� - 0 >= Trunc(Sysdate) - n_��Ч���� And a.����id = n_����id And Rownum < 2;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","exist_order":' || n_Count || '}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Chkoutpatiexistorder;
/
Create Or Replace Procedure Zl_Cissvr_Delerrbillinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------------------------------------------------------------
  --���ܣ�ɾ��ҽ��վԤԼ���ղ������쳣��¼����Ϊ���ڴ���
  --��Σ�json��ʽ
  --Input
  --   rgst_no               C  1 �Һŵ�
  --���Σ�json��ʽ
  --Json_Out
  --   code                  N  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --   message               C  1  Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  -------------------------------------------------------------------------------------------------
  n_�쳣id ���˽����쳣��¼.Id%Type;
  v_�Һŵ� ���˽����쳣��¼.Ԥ������%Type;
  j_Json   Pljson;
  j_In     Pljson;
Begin
  --�������
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  v_�Һŵ� := j_Json.Get_String('rgst_no');

  Select ID Into n_�쳣id From ���˽����쳣��¼ Where Ԥ������ = v_�Һŵ� And �������� = 4 And Rownum < 2;
  Zl_���˽����쳣��¼_Modify(n_�쳣id, 2, Null, Null, Null);
  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Delerrbillinfo;
/
Create Or Replace Procedure Zl_Cissvr_Deloutpativisitrec
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����: �˺ųɹ��������ٴ�����ǼǼ�¼,Ŀǰ�ٴ��ľ���ǼǼ�¼���ǲ��˹Һż�¼�����Բ��ô���
  --��Σ�Json_In:��ʽ
  --input
  --  rgst_no     C  1  �Һŵ���,����ȡ��ԤԼʱ�ᴫ����,�磺U0000001,U0000002
  --����: Json_Out,��ʽ����
  --output
  --  code        N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message     C 1 Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ,ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ----------------------------------------------------------------------------
  j_Json     Pljson;
  v_�Һŵ��� varchar2(4000);
  j_In       Pljson;
Begin
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  v_�Һŵ��� := j_Json.Get_String('rgst_no');

  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Deloutpativisitrec;
/
Create Or Replace Procedure Zl_Cissvr_Existadvice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:�ж��Ƿ����ҽ�����ݻ��ж�ָ���Һŵ��Ƿ��Ѿ���ҽ��
  --��Σ�Json_In:��ʽ
  -- input
  --   pati_id              N 1 ����ID
  --   pati_pageid          N   ��ҳId
  --   rgst_no              C 1 �Һŵ�������ö��ŷָ�
  --   only_valid           N   ֻ���û�����ϵ�ҽ��
  --����: Json_Out,��ʽ����
  --  output
  --    code               C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message            C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    exist              N 1 �Ƿ���ڣ�1-����;0-������
  ---------------------------------------------------------------------------
  j_Json       PLJson;
  j_Json_Tmp   PLJson;
  n_����id     Number(18);
  n_��ҳid     Number(18);
  v_�Һŵ�     Varchar2(3000);
  n_Exist      Number(2);
  n_��������� Number(1);

Begin
  --�������
  j_Json     := PLJson(Json_In);
  j_Json_Tmp := j_Json.Get_Pljson('input');

  n_����id     := j_Json_Tmp.Get_Number('pati_id');
  n_��ҳid     := j_Json_Tmp.Get_Number('pati_pageid');
  v_�Һŵ�     := j_Json_Tmp.Get_String('rgst_no');
  n_��������� := Nvl(j_Json_Tmp.Get_Number('only_valid'), 0);

  If v_�Һŵ� Is Not Null Then
  
    Select Max(1)
    Into n_Exist
    From ����ҽ����¼ A
    Where (����id + 0 = Nvl(n_����id, 0) Or Nvl(n_����id, 0) = 0) And
          �Һŵ� In (Select Column_Value As �Һŵ� From Table(f_Str2List(v_�Һŵ�))) And
          (n_��������� = 0 Or n_��������� = 1 And ҽ��״̬ <> 4);
  
  Elsif j_Json_Tmp.Exist('pati_pageid') Then
    Select Max(����)
    Into n_Exist
    From (Select 1 As ����
           From ����ҽ����¼ A
           Where ����id = n_����id And Nvl(��ҳid, 0) = Nvl(n_��ҳid, 0) And (n_��������� = 0 Or n_��������� = 1 And ҽ��״̬ <> 4) And
                 Rownum < 2);
  Else
    Select Max(����)
    Into n_Exist
    From (Select 1 As ����
           From ����ҽ����¼ A
           Where ����id = n_����id And (n_��������� = 0 Or n_��������� = 1 And ҽ��״̬ <> 4) And Rownum < 2);
  End If;

  Json_Out := '{"output":{"code":1,"message":"' || '�ɹ�' || '","exist":' || Nvl(n_Exist, 0) || '}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || zlJsonStr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Existadvice;
/
Create Or Replace Procedure Zl_Cissvr_Existadvicesend
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ݷ��õ��ݺż����ʣ��ж��Ƿ����ҽ����������
  --��Σ�Json_In:��ʽ
  --  input
  --    fee_no              C 1 ���ݺ�
  --    send_no             N 1 ���ͺ�

  --����: Json_Out,��ʽ����
  --  output
  --    code                N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    exsit               N 1 �Ƿ����:1-����;0-������
  ---------------------------------------------------------------------------
  j_Json   Pljson;
  j_In     Pljson;
  v_���ݺ� ����ҽ������.No%Type;
  n_���ͺ� ����ҽ������.���ͺ�%Type;
  n_Exist  Number(1);

Begin
  --�������
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_���ͺ� := j_Json.Get_Number('send_no');
  v_���ݺ� := j_Json.Get_String('fee_no');

  Select Count(1) Into n_Exist From ����ҽ������ A Where a.No = v_���ݺ� And a.���ͺ� = n_���ͺ�;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","exsit":' || n_Exist || '}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Existadvicesend;
/
Create Or Replace Procedure Zl_Cissvr_Existoutadvice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:ҽ���Ƿ��´��˳�Ժҽ��
  --��Σ�Json_In:��ʽ
  -- input
  --   pati_id              N 1 ����ID
  --   pati_pageid          N 1 ��ҳId
  --����: Json_Out,��ʽ����
  --  output
  --    code               N  1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message            C  1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    is_out             N  1 �Ƿ��ѿ���Ժҽ�� ��1-�ѿ���Ժ;0-δ����Ժ
  --    out_advice_id      N  1 �Ѿ����˳�Ժҽ���ģ�����ҽ����ID
  ---------------------------------------------------------------------------
  j_In     Pljson;
  j_Json   Pljson;
  n_����id ����ҽ����¼.����id%Type;
  n_��ҳid ����ҽ����¼.��ҳid%Type;

  n_Tmp    Number(1);
  n_ҽ��id ����ҽ����¼.Id%Type;

Begin
  --�������
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');
  n_��ҳid := j_Json.Get_Number('pati_pageid');

  If Nvl(n_����id, 0) = 0 Or Nvl(n_��ҳid, 0) = 0 Then
    Json_Out := Zljsonout('���봫��ҽ��id����ҳid��');
    Return;
  End If;

  Select Max(a.Id), Count(1)
  Into n_ҽ��id, n_Tmp
  From ����ҽ����¼ A, ���˱䶯��¼ B, ������ҳ C, ������ĿĿ¼ D
  Where a.����id = n_����id And a.��ҳid = n_��ҳid And a.ҽ��״̬ = 8 And a.����id = b.����id And a.��ҳid = b.��ҳid And
        a.��ʼִ��ʱ�� = b.��ʼʱ�� + 0 And b.��ʼԭ�� = 10 And b.����id = c.����id And b.��ҳid = c.��ҳid And c.״̬ = 3 And d.��� = 'Z' And
        d.�������� In ('5', '6', '11') And a.������Ŀid = d.Id;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","is_out":' || n_Tmp || ',"out_advice_id":' || Nvl(n_ҽ��id, 0) || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Existoutadvice;
/
Create Or Replace Procedure Zl_Cissvr_Getadviceannexinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡҽ��������Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --    advice_ids                C 1 ҽ��ID,�����','�ָ�
  --    send_no                   N 0 ���ͺ�
  --    bill_no                   C 0 NO
  --    bill_prop                 N 1 ��¼����
  --����: Json_Out,��ʽ����
  --  output
  --    code                      N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                   C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    advice_annex_list             [����]ÿ��ҽ��id��Ӧ�ĸ��ѵ�����Ϣ
  --      advice_id               N 1 ҽ��ID
  --      send_no                 N 1 ���ͺ�
  --      bill_no                 C   No
  --      bill_prop               N   ��¼����
  ---------------------------------------------------------------------------
  j_Json Pljson;
  j_In   Pljson;

  v_No       ����ҽ������.No%Type;
  n_��¼���� ����ҽ������.��¼����%Type;
  n_���ͺ�   ����ҽ������.���ͺ�%Type;

  v_ҽ��ids Clob; --��¼ҽ��id
  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_ҽ��id Collection_Type;
  I          Number;
  v_Jtmp     Varchar2(32767);
  c_Jtmp     Clob;
Begin
  --�������
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  v_ҽ��ids  := j_Json.Get_String('advice_ids');
  n_���ͺ�   := j_Json.Get_Number('send_no');
  v_No       := j_Json.Get_String('bill_no');
  n_��¼���� := j_Json.Get_Number('bill_prop');

  I := 0;
  While v_ҽ��ids Is Not Null Loop
    If Length(v_ҽ��ids) <= 4000 Then
      Col_ҽ��id(I) := v_ҽ��ids;
      v_ҽ��ids := Null;
    Else
      Col_ҽ��id(I) := Substr(v_ҽ��ids, 1, Instr(v_ҽ��ids, ',', 3980) - 1);
      v_ҽ��ids := Substr(v_ҽ��ids, Instr(v_ҽ��ids, ',', 3980) + 1);
    End If;
    I := I + 1;
  End Loop;
  I := 0;
  For I In 0 .. Col_ҽ��id.Count - 1 Loop
    For r_ҽ�� In (Select a.��¼����, a.ҽ��id, a.���ͺ�, a.No
                 From ����ҽ������ A
                 Where a.ҽ��id In (Select /*+cardinality(B,10) */
                                   Column_Value As ҽ��id
                                  From Table(f_Num2list(Col_ҽ��id(I))) B) And (a.���ͺ� = n_���ͺ� Or n_���ͺ� Is Null) And
                       a.��¼���� = n_��¼���� And (a.No = v_No Or v_No Is Null)) Loop
    
      v_Jtmp := v_Jtmp || ',{"advice_id":' || r_ҽ��.ҽ��id;
      v_Jtmp := v_Jtmp || ',"send_no":' || r_ҽ��.���ͺ�;
      v_Jtmp := v_Jtmp || ',"bill_no":"' || r_ҽ��.No || '"';
      v_Jtmp := v_Jtmp || ',"bill_prop":' || r_ҽ��.��¼����;
      v_Jtmp := v_Jtmp || '}';
    
      If Length(v_Jtmp) > 30000 Then
        If c_Jtmp Is Null Then
          c_Jtmp := Substr(v_Jtmp, 2);
        Else
          c_Jtmp := c_Jtmp || v_Jtmp;
        End If;
        v_Jtmp := Null;
      End If;
    End Loop;
  End Loop;
  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","advice_annex_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","advice_annex_list":[' || c_Jtmp || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getadviceannexinfo;
/
Create Or Replace Procedure Zl_Cissvr_Getadviceannexnote
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ�����ҽ��ID��Ԫ�����ƻ�ȡҽ����������
  --��Σ�Json_In:��ʽ
  --  input
  --    advice_id           N 1 ҽ��ID
  --    chn_name            C 1 ����������Ŀ.������,���磺����ҽ������
  --����: Json_Out,��ʽ����
  --  output
  --    code                      N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                   C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    note_list[]������
  --       annex_note               C 1 ҽ����������:����ҽ������.����
  ---------------------------------------------------------------------------
  j_In       Pljson;
  j_Json     Pljson;
 
  n_ҽ��id   ����ҽ������.ҽ��id%Type;
  v_������Ŀ ����������Ŀ.������%Type;
  v_Jtmp     Varchar2(32767);
 
Begin
  --�������
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');

  n_ҽ��id   := j_Json.Get_Number('advice_id');
  v_������Ŀ := j_Json.Get_String('chn_name');

  If Nvl(n_ҽ��id, 0) = 0 Then
    Json_Out := Zljsonout('δ����ҽ��id�����飡');
    Return;
  End If;

  If Nvl(v_������Ŀ, '-') = '-' Then
    Json_Out := Zljsonout('δ��������������Ŀ���ƣ����飡');
    Return;
  End If;

  For r_���� In (Select a.����
               From ����ҽ������ A, ����������Ŀ B
               Where a.Ҫ��id = b.Id And a.ҽ��id = n_ҽ��id And b.������ = v_������Ŀ) Loop
  
    v_Jtmp := v_Jtmp || ',{"annex_note":"' || Zljsonstr(r_����.����) || '"}';
  
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","note_list":[' || Substr(v_Jtmp, 2) || ']}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getadviceannexnote;
/
Create Or Replace Procedure Zl_Cissvr_Getadvicedefinedinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡҽ�����ݶ���������Ϣ
  --��Σ�Json_In:��ʽ
  --��
  --����: Json_Out,��ʽ����
  --  output
  --    code                      N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                   C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    item_list[]
  --      clinic_type             C 1 �������
  --      advice_note             C 1 ҽ������
  ---------------------------------------------------------------------------

  v_Jtmp Varchar2(32767);

Begin
  For r_ҽ������ In (Select �������, ҽ������ From ҽ�����ݶ��� Order By �������) Loop
  
    v_Jtmp := v_Jtmp || ',{"clinic_type":"' || r_ҽ������.������� || '"';
    v_Jtmp := v_Jtmp || ',"advice_note":"' || Zljsonstr(r_ҽ������.ҽ������) || '"';
    v_Jtmp := v_Jtmp || '}';
  
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getadvicedefinedinfo;
/
Create Or Replace Procedure Zl_Cissvr_Getadviceexcutnums
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡҽ�����͵���ִ������������ҽ��ID��ѯҽ��������Ϣ
  --���      json
  --input     
  --  item_list                 ������
  --    advice_id               N 1 ҽ��ID
  --    bill_no                 C 1 NO
  --    bill_prop               N 1 ��¼����
  --����      json
  --output
  --  code                      C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message                   C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  item_list[]������
  --    advice_id               N 1 ҽ��ID
  --    bill_no                 C 1 No
  --    fee_item_id             N 1 �շ�ϸĿID
  --    execute_num             N 1 ��ִ����
  --˵����ע�⣬��ʹ��ִ����Ϊ0ҲҪ����
  ---------------------------------------------------------------------------
  j_Json     Pljson;
  j_List     Pljson_List;
  j_Json_Tmp Pljson;

  n_ҽ��id   ����ҽ������.ҽ��id%Type;
  v_No       ����ҽ������.No%Type;
  n_��¼���� ����ҽ������.��¼����%Type;
  v_Jtmp     Varchar2(32767);
  c_Jtmp     Clob;
Begin
  j_Json_Tmp := Pljson(Json_In);
  j_Json     := j_Json_Tmp.Get_Pljson('input');
  j_List     := j_Json.Get_Pljson_List('item_list');
  
  For I In 1 .. j_List.Count Loop
    j_Json_Tmp := Pljson();
    j_Json_Tmp := Pljson(j_List.Get(I));
    n_ҽ��id   := j_Json_Tmp.Get_Number('advice_id');
    v_No       := j_Json_Tmp.Get_String('bill_no');
    n_��¼���� := j_Json_Tmp.Get_Number('bill_prop');
  
    --Zl_ҽ��ִ�мƼ�_����(
    --  ҽ��id_In   ����ҽ��ִ��.ҽ��id%Type,
    --  No_In       ����ҽ������.No%Type,
    --  ��¼����_In ����ҽ������.��¼����%Type
    Zl_ҽ��ִ�мƼ�_����(n_ҽ��id, v_No, n_��¼����);
  
    --1.����ҽ��ִ�мƼ۵�,����ҽ��ִ�мƼ�Ϊ׼(�����ܰ���:���;����;����;������Ѫ)
    --2.����ҽ������.ִ��״̬=1�����ִ�У�ʱ��׼����Ϊ0�����ٸ���ҽ��ִ�мƼ���ͳ��
    For c_ҽ�� In (Select ҽ��id, NO, �շ�ϸĿid, Sum(��ִ����) As ��ִ����
                 From (Select b.ҽ��id, b.No, c.�շ�ϸĿid, Decode(b.ִ��״̬, 1, 1, Decode(c.ִ��״̬, 1, 1, 0)) * c.���� As ��ִ����
                        From ����ҽ������ B, ҽ��ִ�мƼ� C, ����ҽ����¼ M
                        Where b.ҽ��id = c.ҽ��id And b.���ͺ� = c.���ͺ� And b.ҽ��id = m.Id And
                              Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And b.ҽ��id = n_ҽ��id And b.No = v_No And
                              b.��¼���� = n_��¼����)
                 Group By ҽ��id, NO, �շ�ϸĿid) Loop
    
      v_Jtmp := v_Jtmp || ',{"advice_id":' || c_ҽ��.ҽ��id;
      v_Jtmp := v_Jtmp || ',"bill_no":"' || c_ҽ��.No || '"';
      v_Jtmp := v_Jtmp || ',"fee_item_id":' || c_ҽ��.�շ�ϸĿid;
      v_Jtmp := v_Jtmp || ',"execute_num":' || Zljsonstr(c_ҽ��.��ִ����, 1);
      v_Jtmp := v_Jtmp || '}';
    
      If Length(v_Jtmp) > 30000 Then
        If c_Jtmp Is Null Then
          c_Jtmp := Substr(v_Jtmp, 2);
        Else
          c_Jtmp := c_Jtmp || v_Jtmp;
        End If;
        v_Jtmp := Null;
      End If;
    
    End Loop;
  End Loop;
  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || c_Jtmp || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getadviceexcutnums;
/
Create Or Replace Procedure Zl_Cissvr_Getadviceexestatus
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡҽ�����͵�ִ��״̬
  --��Σ�Json_In:��ʽ
  --  input
  --    advice_id               N 1 ҽ��ID
  --    send_no                 N 1 ���ͺ�

  --����: Json_Out,��ʽ����
  --  output
  --    code                      N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                   C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    exe_status                N 1 ִ��״̬
  ---------------------------------------------------------------------------
  j_In       Pljson;
  j_Json     Pljson;
  n_���ͺ�   ����ҽ������.���ͺ�%Type;
  n_ҽ��id   ����ҽ������.ҽ��id%Type;
  n_ִ��״̬ ����ҽ������.ִ��״̬%Type;

Begin
  --�������
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_ҽ��id := j_Json.Get_Number('advice_id');
  n_���ͺ� := j_Json.Get_Number('send_no');

  If Nvl(n_ҽ��id, 0) = 0 Or Nvl(n_���ͺ�, 0) = 0 Then
    Json_Out := Zljsonout('δ����ҽ��id���ͺţ����飡');
    Return;
  End If;

  Select Max(ִ��״̬)
  Into n_ִ��״̬
  From ����ҽ������
  Where ���ͺ� = n_���ͺ� And
        ҽ��id In (Select ID From ����ҽ����¼ Where (ID = n_ҽ��id Or ���id = n_ҽ��id) And ������� In ('C', 'D'));
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","exe_status":' || Nvl(n_ִ��״̬, 0) || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getadviceexestatus;
/
Create Or Replace Procedure Zl_Cissvr_Getadvicefeestatus
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡҽ�����͵ļƷ�״̬
  --��Σ�Json_In:��ʽ
  --  input
  --   advice_id            N 1 ҽ��ID
  --   send_no              N 1 ���ͺ�
  --   isalone_exe          N 1 �Ƿ����ִ��:1-����ִ��;0-�Ƕ���ִ��

  --����: Json_Out,��ʽ����
  --  output
  --    code                      N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                   C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    charge_status             C 1 �Ʒ�״̬:����ö��ŷ�������-1=����Ʒ�,1=�ѼƷ�,0=δ�Ʒ�,�������ﵥ�ݣ�2=�����շ�,3=ȫ���շ�
  ---------------------------------------------------------------------------
  j_In   Pljson;
  j_Json Pljson;

  n_ҽ��id   ����ҽ����¼.Id%Type;
  n_���id   ����ҽ����¼.���id%Type;
  n_����ִ�� Number(1);
  n_���ͺ�   ����ҽ������.���ͺ�%Type;
  v_�Ƿ�״̬ Varchar2(100);
  v_������� ����ҽ����¼.�������%Type;

Begin
  --�������
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');

  n_ҽ��id   := j_Json.Get_Number('advice_id');
  n_���ͺ�   := j_Json.Get_Number('send_no');
  n_����ִ�� := Nvl(j_Json.Get_Number('isalone_exe'), 0);

  If Nvl(n_ҽ��id, 0) = 0 Or Nvl(n_���ͺ�, 0) = 0 Then
    Json_Out := Zljsonout('δ����ҽ��id���ͺţ����飡');
    Return;
  End If;

  Select Decode(a.�������, 'D', Nvl(a.���id, a.Id), a.���id), �������
  Into n_���id, v_�������
  From ����ҽ����¼ A
  Where a.Id = n_ҽ��id;

  If n_����ִ�� = 1 Then
    Select Distinct �Ʒ�״̬ Into v_�Ƿ�״̬ From ����ҽ������ Where ҽ��id = n_ҽ��id And ���ͺ� = n_���ͺ�;
  Else
  
    Select Distinct f_List2str(Cast(Collect(To_Char(�Ʒ�״̬)) As t_Strlist))
    Into v_�Ƿ�״̬
    From ����ҽ������
    Where ҽ��id In (Select ID From ����ҽ����¼ Where (ID = n_���id Or ���id = n_���id) And ������� = v_�������) And ���ͺ� = n_���ͺ�;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","charge_status":"' || v_�Ƿ�״̬ || '"}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getadvicefeestatus;
/
Create Or Replace Procedure Zl_Cissvr_Getadviceids
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:��ȡ���˵�ҽ��IDs
  --��Σ�Json_In:��ʽ
  -- input
  --   advice_register_no   C 1 �Һŵ���:�Һŵ����˱ش�����һ������
  --   pati_id              N 1 ����ID:�Һŵ����˱ش�����һ������
  --   advice_starttime     C 0 ��ʼ�Ŀ���ʱ��,��ʽ:yyyy-mm-dd hh24:mi:ss
  --   advice_endtime       C 0 �����Ŀ���ʱ��,��ʽ:yyyy-mm-dd hh24:mi:ss
  --   isgetlast_adviceid   N 0 �Ƿ��ȡ���һ��ҽ��id,1-��ȡ���һ��ҽ��id;0-ȫ����ȡ

  --����: Json_Out,��ʽ����
  --  output
  --    code               C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message            C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    advice_ids         C 1 ���:isgetlast_adviceid=1ʱ���������һ��ҽ��id,���򷵻���������������ҽ��id
  ---------------------------------------------------------------------------
  j_In       Pljson;
  j_Json     Pljson;
  n_����id   Number(18);
  v_�Һŵ�   Varchar2(100);
  d_��ʼʱ�� Date;
  d_����ʱ�� Date;
  n_���һ�� Number(1);
  v_ҽ��ids  Varchar2(32767);
Begin
  --�������
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  n_����id   := j_Json.Get_Number('pati_id');
  v_�Һŵ�   := j_Json.Get_String('advice_register_no');
  d_��ʼʱ�� := To_Date(j_Json.Get_String('advice_starttime'), 'yyyy-mm-dd hh24:mi:ss');
  d_����ʱ�� := To_Date(j_Json.Get_String('advice_endtime'), 'yyyy-mm-dd hh24:mi:ss');
  n_���һ�� := Nvl(j_Json.Get_Number('isgetlast_adviceid'), 0);

  If Nvl(n_����id, 0) = 0 And Nvl(v_�Һŵ�, '-') = '-' Then
    Json_Out := Zljsonout('δ����Һŵ�����id��');
    Return;
  End If;

  If Nvl(n_����id, 0) <> 0 Then
    --���ݲ���id��ѯҽ��id
    If n_���һ�� = 1 Then
      Select Max(ID)
      Into v_ҽ��ids
      From ����ҽ����¼
      Where ����id = n_����id And ����ʱ�� Between Nvl(d_��ʼʱ��, ����ʱ��) And Nvl(d_����ʱ��, ����ʱ��);
    Else
      For r_ҽ�� In (Select ID
                   From ����ҽ����¼
                   Where ����id = n_����id And ����ʱ�� Between Nvl(d_��ʼʱ��, ����ʱ��) And Nvl(d_����ʱ��, ����ʱ��)) Loop
        v_ҽ��ids := v_ҽ��ids || ',' || r_ҽ��.Id;
      End Loop;
    End If;
  Else
    --���ݹҺŵ���ѯҽ��id
    If n_���һ�� = 1 Then
      Select Max(ID)
      Into v_ҽ��ids
      From ����ҽ����¼ M
      Where �Һŵ� = v_�Һŵ� And ����ʱ�� Between Nvl(d_��ʼʱ��, ����ʱ��) And Nvl(d_����ʱ��, ����ʱ��);
    Else
      For r_ҽ�� In (Select ID
                   From ����ҽ����¼
                   Where �Һŵ� = v_�Һŵ� And ����ʱ�� Between Nvl(d_��ʼʱ��, ����ʱ��) And Nvl(d_����ʱ��, ����ʱ��)) Loop
        v_ҽ��ids := v_ҽ��ids || ',' || r_ҽ��.Id;
      End Loop;
    End If;
  End If;
  v_ҽ��ids := Substr(v_ҽ��ids, 2);
  Json_Out  := '{"output":{"code":1,"message":"�ɹ�","advice_ids":"' || v_ҽ��ids || '"}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getadviceids;
/
Create Or Replace Procedure Zl_Cissvr_Getadviceidsfromdiag
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ��������ID��ȡ��Ӧ��ҽ��Ids
  --��Σ�Json_In:��ʽ
  --  input
  --    diag_ids                  C 1 ���id,����ö��ŷ���
  --����: Json_Out,��ʽ����
  --  output
  --    code                      N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                   C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    advice_ids                C 1 ҽ��ids,����ö��ŷ���
  ---------------------------------------------------------------------------
  j_In      Pljson;
  j_Json    Pljson;
  I         Number;
  c_���ids Clob;
  c_ҽ��ids Varchar2(32767);
  v_ҽ��ids Varchar2(32767);
  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_���id Collection_Type;

Begin
  --�������
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');
  Begin
    c_���ids := j_Json.Get_Clob('diag_ids');
  Exception
    When Others Then
      Json_Out := Zljsonout('δ�������id�����飡');
      Return;
  End;

  If Nvl(c_���ids, '-') = '-' Then
    Json_Out := Zljsonout('δ�������id�����飡');
    Return;
  End If;

  I := 0;
  While c_���ids Is Not Null Loop
    If Length(c_���ids) <= 4000 Then
      Col_���id(I) := c_���ids;
      c_���ids := Null;
    Else
      Col_���id(I) := Substr(c_���ids, 1, Instr(c_���ids, ',', 3980) - 1);
      c_���ids := Substr(c_���ids, Instr(c_���ids, ',', 3980) + 1);
    End If;
    I := I + 1;
  End Loop;
  I := 0;
  For I In 0 .. Col_���id.Count - 1 Loop
    Select f_List2str(Cast(Collect(To_Char(ҽ��id)) As t_Strlist))
    Into v_ҽ��ids
    From (With ҽ������ As (Select /*+cardinality(b,10)*/
                         a.ҽ��id
                        From �������ҽ�� A, Table(f_Num2list(Col_���id(I))) B
                        Where a.���id = b.Column_Value)
           Select a.Id As ҽ��id
           From ����ҽ����¼ A, ҽ������ B
           Where a.Id = b.ҽ��id
           Union
           Select a.Id
           From ����ҽ����¼ A, ҽ������ B
           Where a.���id = b.ҽ��id);
  
  
    c_ҽ��ids := c_ҽ��ids || ',' || v_ҽ��ids;
  End Loop;
  c_ҽ��ids := Substr(c_ҽ��ids, 2);

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","advice_ids":"' || c_ҽ��ids || '"}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getadviceidsfromdiag;
/
Create Or Replace Procedure Zl_Cissvr_Getadviceinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:��ȡҽ����Ϣ
  --��Σ�Json_In:��ʽ
  -- input
  --   query_type                   N 0 ��ѯ���ͣ�0:��ѯ������Ϣ��1:��ѯ������Ϣ+��չ��Ϣ
  --   advice_ids                   C 0 ���ҽ��ID��������ҩ����Ҳ��������ҽ������ҩ;����,�á�,���ָ�
  --   rgst_no                      C 0 �Һŵ���:�Һŵ�����ID��ҽ��ID�ش�����һ������
  --   pati_id                      N 0 ����ID:�Һŵ�����ID��ҽ��ID�ش�����һ������
  --   pati_pageid                  N 0 ��ҳId
  --����: Json_Out,��ʽ����
  --  output
  --    code                        C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message                     C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    advice_list      [����]ÿ��ҽ����Ϣ
  --      advice_id                 N    id
  --      advice_related_id         N    ���id
  --      pati_id                   N    ����id
  --      pati_pageid               N    ��ҳid
  --      pati_source               N    ������Դ
  --      advice_statu              N    ҽ��״̬:-1-δ��Ч���ݴ�ҽ����1-�¿���2-У�����ʣ�3-��У�ԣ�4-�����ϣ�5-��������6-����ͣ��7-�����ã�8-��ֹͣ��9-��ȷ��ֹͣ
  --      serial_num                N    ���
  --      advice_day                N    ����
  --      advice_dosage             N    ��������
  --      oper_type                 C    ��������:������ĿĿ¼.��������
  --      clinic_type               C    �������
  --      advice_exe_properties     N    ִ������
  --      advice_exe_sign           N    ִ�б��
  --      effective_time            N    ҽ����Ч
  --      advice_record_time        D    ����ʱ��
  --      advice_doctor             C    ����ҽ��
  --      advice_purpose            C    ��ҩĿ��
  --      advice_reason             C    ��ҩ����
  --      advice_taboonote          C    ����ҩƷ˵��
  --      advice_doctor_note        C    ҽ������
  --      rcpdtl_excs_desc          C    ����˵��
  --      advice_audit_result       N    �����
  --      advice_audit_sign         N    ҩʦ��˱�־
  --      advice_audit_time         D    ҩʦ���ʱ��
  --      advice_interval_unit      C    �����λ
  --      advice_frequency          C    ִ��Ƶ��
  --      advice_frequency_times    N    Ƶ�ʴ���
  --      advice_frequency_interval N    Ƶ�ʼ��
  --      advice_exetime_plane      C    ִ��ʱ�䷽��
  --      advice_begintime          D    ��ʼִ��ʱ��
  --      advice_endtime            D    ִ����ֹʱ��
  --      rgst_no                   C   �Һŵ���
  --      advice_receipt_name       C   �䷽����
  --      advice_receipt_issecret   N   �Ƿ���
  --      advice_note               C   ҽ������
  --      advice_cisitem_id         N   ������Ŀid
  --      advice_item_id            N   �շ�ϸĿid
  --      pati_deptid               N   ���˿���id
  --      pati_name                 C   ����
  --      pati_sex                  C   �Ա�
  --      pati_age                  C   ����
  --      advice_audit_reason       C   ҩʦ���ԭ��
  --      skintest_info             C   Ƥ�Խ��

  --      total_qunt                N    �ܸ�����
  --      Total_qunt_unit           C    ����:����λ������(�ܸ�����+��λ)
  --      single                    C    ����:��������+��λ
  --      toxicity_type             C    �������:ҩƷ��Ч��ҩƷ����.�������
  --      advice_dept_id            N    ��������ID
  --      advice_stop_doctor        C    ͣ��ҽ��
  --      advice_stop_nurse         C    ͣ����ʿ
  --      advice_stoptime           C    ͣ��ʱ��,��ʽ��yyyy-mm-dd hh24:mi:ss
  --      advice_stoptime_confirm   C    ȷ��ͣ��ʱ��:��ʽ:yyyy-mm-dd hh24:mi:ss
  --      order_chk_nurse           C    У�Ի�ʿ
  --      order_chk_time            C    У��ʱ��:yyyy-mm-dd hh24:mi;ss
  --      lastexe_time              D    �ϴ�ִ��ʱ��
  --      usage                     C    �÷�
  --      emergency_tag             N    ������־:0-��ͨ;1-����;2-��¼(��������Ч)
  --      is_charge_verfy           N   �Ƿ�������:1-���;0-δ���
  --      baby_num                  N   Ӥ�����
  --      valuation_nature          N   �Ƽ�����:0-�����Ƽۣ�1-���Ƽۣ�2-�ֹ��Ƽ�
  --      advice_exedept_id         N   ִ�п���ID:
  --      advice_exedept_name       C   ִ�п�������
  --      testtube_code             C   �Թܱ���:������ĿĿ¼.�Թܱ���
  --      hide_print                N   ���δ�ӡ
  --      prerequisite_id           N   ǰ��ID
  --      is_staff_sig              N   �Ƿ�ǩ��:1-ǩ��;0-δǩ��
  --      rpt_id                    N   ��������id
  --      consult_statu             N   ����״̬:0-δ����,1-�Ѳ���
  --      advice_verfy_statu        N   ���״̬:Null-������ˣ�1-����ˣ�2-���ͨ����3-���δͨ����4��Ѫ������գ�5��Ѫ����Ѫ�У�6-Ѫ��ֹͣ��Ѫ��7-��Ѫ����˴�ǩ��
  --      apply_num                 N   �������
  --      cisitem_unit              C   ������Ŀ���㵥λ
  --      pati_wardarea_id          N   ��ǰ����ID
  --      decoction_method          C   �巨
  ---------------------------------------------------------------------------
  j_In     Pljson;
  j_Json   Pljson;
  v_ҽ��id Clob; --��¼ҽ��id
  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_ҽ��id Collection_Type;
  I          Number;

  v_�Һŵ�    Varchar2(50);
  n_����id    Number;
  n_��ҳid    Number;
  n_Ƥ�Ի�ȡ  Number; --null δ��ȡ�α�1-����ȡ��2-����ȡ��������
  v_Ƥ�Խ��  Varchar2(200);
  v_Pre����   Varchar2(600);
  n_��ѯ����  Number(1);
  v_Adviceout Varchar2(32767);
  v_Temp      Varchar2(32767);
  c_Jtmp      Clob;
  n_���id    ����ҽ����¼.Id%Type;
  v_�巨      Varchar2(50);

  Cursor Ctest�巨(���id_In ����ҽ����¼.Id%Type) Is
    Select b.����
    From ����ҽ����¼ A, ������ĿĿ¼ B
    Where a.������Ŀid = b.Id And a.���id = ���id_In And a.������� = 'E';
  r_�巨 Ctest�巨%RowType;

  Cursor Ctest(�Һŵ�_In ����ҽ����¼.�Һŵ�%Type) Is
    Select a.Ƥ�Խ��, a.ҩ��id, a.ҩƷid
    From (Select a.Ƥ�Խ��, c.��Ŀid As ҩ��id, 0 As ҩƷid, a.��ʼִ��ʱ��
           From ����ҽ����¼ A, �����÷����� C
           Where a.������Ŀid = c.�÷�id And Nvl(c.����, 0) = 0 And Nvl(a.ҽ��״̬, 0) = 8 And a.Ƥ�Խ�� Is Not Null And a.�Һŵ� = �Һŵ�_In
           Union All
           Select a.Ƥ�Խ��, b.ҩ��id, b.ҩƷid, a.��ʼִ��ʱ��
           From ����ҽ����¼ A, ҩƷ��� B, ҩƷ�÷����� C
           Where a.������Ŀid = c.�÷�id And b.ҩƷid = c.ҩƷid And Nvl(c.����, 0) = 0 And Nvl(a.ҽ��״̬, 0) = 8 And a.Ƥ�Խ�� <> '����' And
                 a.�Һŵ� = �Һŵ�_In
           Union All
           Select a.Ƥ�Խ��, c.��Ŀid As ҩ��id, 0 As ҩƷid, a.��ʼִ��ʱ��
           From ����ҽ����¼ A, �����÷����� C
           Where a.������Ŀid = c.�÷�id And Nvl(c.����, 0) = 0 And a.Ƥ�Խ�� = '����' And a.�Һŵ� = �Һŵ�_In) A
    Order By a.��ʼִ��ʱ�� Desc;
  Type t_Test Is Table Of Ctest%RowType;
  Rs_Test t_Test;

  Cursor CtestסԺ
  (
    ����id_In ����ҽ����¼.����id%Type,
    ��ҳid_In ����ҽ����¼.��ҳid%Type
  ) Is
    Select a.Ƥ�Խ��, a.ҩ��id, a.ҩƷid
    From (Select a.Ƥ�Խ��, c.��Ŀid As ҩ��id, 0 As ҩƷid, a.��ʼִ��ʱ��
           From ����ҽ����¼ A, �����÷����� C
           Where a.������Ŀid = c.�÷�id And Nvl(c.����, 0) = 0 And Nvl(a.ҽ��״̬, 0) = 8 And a.Ƥ�Խ�� Is Not Null And
                 a.����id = ����id_In And a.��ҳid = ��ҳid_In
           Union All
           Select a.Ƥ�Խ��, b.ҩ��id, b.ҩƷid, a.��ʼִ��ʱ��
           From ����ҽ����¼ A, ҩƷ��� B, ҩƷ�÷����� C
           Where a.������Ŀid = c.�÷�id And b.ҩƷid = c.ҩƷid And Nvl(c.����, 0) = 0 And Nvl(a.ҽ��״̬, 0) = 8 And a.Ƥ�Խ�� <> '����' And
                 a.����id = ����id_In And a.��ҳid = ��ҳid_In
           Union All
           Select a.Ƥ�Խ��, c.��Ŀid As ҩ��id, 0 As ҩƷid, a.��ʼִ��ʱ��
           From ����ҽ����¼ A, �����÷����� C
           Where a.������Ŀid = c.�÷�id And Nvl(c.����, 0) = 0 And a.Ƥ�Խ�� = '����' And a.����id = ����id_In And a.��ҳid = ��ҳid_In) A
    Order By a.��ʼִ��ʱ�� Desc;

  Procedure Pp_Test
  (
    Nҩ��id    Number,
    v_Ƥ�Խ�� Out Varchar2
  ) Is
  Begin
    v_Ƥ�Խ�� := Null;
    For I In 1 .. Rs_Test.Count Loop
      If Rs_Test(I).ҩ��id = Nҩ��id Then
        v_Ƥ�Խ�� := Rs_Test(I).Ƥ�Խ��;
      End If;
    End Loop;
  End Pp_Test;
Begin
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');
  If j_Json.Exist('advice_ids') Then
    v_ҽ��id := j_Json.Get_Clob('advice_ids');
  End If;
  n_����id   := j_Json.Get_Number('pati_id');
  n_��ҳid   := Nvl(j_Json.Get_Number('pati_pageid'), 0);
  v_�Һŵ�   := j_Json.Get_String('rgst_no');
  n_��ѯ���� := Nvl(j_Json.Get_Number('query_type'), 0);

  If v_ҽ��id Is Null Then
    If Nvl(n_����id, 0) = 0 And Nvl(v_�Һŵ�, '-') = '-' Then
      Json_Out := '{"output":{"code":0,"message":"δ�����κβ�ѯ���������飡"}}';
      Return;
    End If;
  
    If Nvl(n_����id, 0) <> 0 Then
      Select f_List2str(Cast(Collect(To_Char(ID)) As t_Strlist))
      Into v_ҽ��id
      From ����ҽ����¼
      Where ����id = n_����id And Decode(n_��ҳid, 0, 0, ��ҳid) = Decode(n_��ҳid, 0, 0, n_��ҳid);
    Else
      Select f_List2str(Cast(Collect(To_Char(ID)) As t_Strlist))
      Into v_ҽ��id
      From ����ҽ����¼
      Where �Һŵ� = v_�Һŵ�;
    End If;
  End If;

  --�� v_ҽ��id ����װ�ɲ�����4000 �ļ��ϴ�����ֹʹ�� f_Num2list ��������
  I := 0;
  While v_ҽ��id Is Not Null Loop
    If Length(v_ҽ��id) <= 4000 Then
      Col_ҽ��id(I) := v_ҽ��id;
      v_ҽ��id := Null;
    Else
      Col_ҽ��id(I) := Substr(v_ҽ��id, 1, Instr(v_ҽ��id, ',', 3980) - 1);
      v_ҽ��id := Substr(v_ҽ��id, Instr(v_ҽ��id, ',', 3980) + 1);
    End If;
    I := I + 1;
  End Loop;

  v_Temp := Null;

  For I In 0 .. Col_ҽ��id.Count - 1 Loop
    For r_ҽ�� In (Select /*+cardinality(j,10)*/
                 Distinct a.Id, a.���id, a.���, g.��� As ҩƷ���, a.����id, a.��ҳid, a.������Դ, a.����, a.��������, a.ִ������, a.ִ�б��, a.ҽ����Ч,
                          a.����ʱ��, a.����ҽ��, a.��ҩĿ��, a.��ҩ����, a.����ҩƷ˵��, a.ҽ������, a.����˵��, a.�����, a.ҩʦ��˱�־, a.ҩʦ���ʱ��, a.�����λ,
                          a.ִ��Ƶ��, a.Ƶ�ʴ���, a.Ƶ�ʼ��, a.ִ��ʱ�䷽��, a.��ʼִ��ʱ��, a.ִ����ֹʱ��, a.�Һŵ�, a.ҽ������, a.������Ŀid, a.�շ�ϸĿid,
                          a.���˿���id, a.����, a.�Ա�, a.����, a.ҩʦ���ԭ��, a.Ƥ�Խ��,
                          Decode(c.���, '7',
                                  Decode(Nvl(g.�䷽id, 0), 0,
                                          Decode(Nvl(v.���id, 0), 0, '', Substr(g.ҽ������, 1, Instr(g.ҽ������, ':') - 1)),
                                          Nvl(n.����, '')), '') As �䷽����,
                          Decode(c.���, '7',
                                  Decode(Nvl(g.�䷽id, 0), 0,
                                          Decode(Nvl(v.���id, 0), 0, 0, Decode(Instr(g.ҽ������, '(�����䷽)'), 0, 0, 1)),
                                          Nvl(n.�Ƿ���, 0)), 0) As �Ƿ���, a.�ܸ�����,
                          Decode(a.�ܸ�����, Null, Null,
                                  Decode(a.�������, 'E', Decode(p.��������, '4', a.�ܸ����� || '��', a.�ܸ����� || p.���㵥λ), '4',
                                          a.�ܸ����� || c.���㵥λ, '5', Round(a.�ܸ����� / d.סԺ��װ, 5) || d.סԺ��λ, '6',
                                          Round(a.�ܸ����� / d.סԺ��װ, 5) || d.סԺ��λ, a.�ܸ����� || p.���㵥λ)) As ����,
                          Decode(a.��������, Null, Null, a.�������� || Decode(a.�������, '4', c.���㵥λ, p.���㵥λ)) As ����, a.�������, p.��������,
                          e.�������, a.��������id, a.ͣ��ҽ��, a.ȷ��ͣ����ʿ, a.ͣ��ʱ��, a.ȷ��ͣ��ʱ��, a.У�Ի�ʿ, a.У��ʱ��, a.�ϴ�ִ��ʱ��, a.������־,
                          a.�Ƿ�������, a.Ӥ�� As Ӥ�����, a.�Ƽ�����, a.ִ�п���id, b.���� As ִ�п�������, p.�Թܱ���, a.���δ�ӡ, a.ǰ��id,
                          Decode(f.ǩ��id, Null, 0, 1) As �Ƿ�ǩ��, h.����״̬, a.���״̬, a.�������, p.���㵥λ As ������Ŀ���㵥λ, m.��ǰ����id,
                          h.����id As ����id, a.ҽ��״̬, Nvl(P1.����, p.����) As �÷�
                 From ����ҽ����¼ A, ���ű� B, �շ���ĿĿ¼ C, ҩƷ��� D, ҩƷ���� E, ����ҽ��״̬ F, ����ҽ����¼ G, ����ҽ������ H, ������ҽ��ϼ�¼ V, ������ĿĿ¼ N,
                      Table(f_Num2list(Col_ҽ��id(I))) J, ������ҳ M, ������ĿĿ¼ P, ����ҽ����¼ A1, ������ĿĿ¼ P1
                 Where Nvl(a.���id, a.Id) = g.Id And a.�շ�ϸĿid = c.Id(+) And g.�䷽id = n.Id(+) And a.ִ�п���id = b.Id(+) And
                       a.������Ŀid = e.ҩ��id(+) And a.�շ�ϸĿid = d.ҩƷid(+) And a.Id = h.ҽ��id(+) And a.Id = f.ҽ��id(+) And
                       g.Id = v.Hisҽ��id(+) And a.����id = m.����id(+) And a.��ҳid = m.��ҳid(+) And a.������Ŀid = p.Id(+) And
                       (a.Id = j.Column_Value Or a.���id = j.Column_Value) And f.�������� = 1 And a.��ʼִ��ʱ�� Is Not Null And
                       Nvl(a.ҽ��״̬, 0) <> -1 And a.���id = A1.Id(+) And A1.������Ŀid = P1.Id(+)
                 Order By a.���id, a.Id) Loop
    
      --���ظ���ҽ��ID�ż��뵽����
      If v_Adviceout Is Null Or Instr(',' || v_Adviceout || ',', ',' || r_ҽ��.Id || ',') = 0 Then
        If v_Adviceout Is Null Then
          v_Adviceout := r_ҽ��.Id;
        Else
          v_Adviceout := v_Adviceout || ',' || r_ҽ��.Id;
        End If;
        --��ȡ����Ƥ����Ϣ�������Ƕಡ��
        If Nvl(n_Ƥ�Ի�ȡ, 0) = 0 Or
           (Nvl(v_Pre����, '*') <> r_ҽ��.����id || '_' || r_ҽ��.��ҳid || '_' || r_ҽ��.�Һŵ� And v_Pre���� Is Not Null) Then
        
          n_Ƥ�Ի�ȡ := 1;
          If r_ҽ��.�Һŵ� Is Not Null Then
            Open Ctest(r_ҽ��.�Һŵ�);
            Fetch Ctest Bulk Collect
              Into Rs_Test;
            Close Ctest;
          
            If Rs_Test.Count > 0 Then
              n_Ƥ�Ի�ȡ := 2;
            End If;
          Elsif r_ҽ��.��ҳid Is Not Null Then
            Open CtestסԺ(r_ҽ��.����id, r_ҽ��.��ҳid);
            Fetch CtestסԺ Bulk Collect
              Into Rs_Test;
            Close CtestסԺ;
          
            If Rs_Test.Count > 0 Then
              n_Ƥ�Ի�ȡ := 2;
            End If;
          End If;
        
          v_Pre���� := r_ҽ��.����id || '_' || r_ҽ��.��ҳid || '_' || '_' || r_ҽ��.�Һŵ�;
        End If;
      
        v_Temp := v_Temp || ',{"advice_id":' || r_ҽ��.Id;
        v_Temp := v_Temp || ',"advice_related_id":' || Nvl(r_ҽ��.���id, 0);
        v_Temp := v_Temp || ',"pati_id":' || Nvl(r_ҽ��.����id, 0);
        v_Temp := v_Temp || ',"pati_pageid":' || Nvl(r_ҽ��.��ҳid, 0);
        v_Temp := v_Temp || ',"pati_source":' || Nvl(r_ҽ��.������Դ, 0);
        v_Temp := v_Temp || ',"advice_statu":' || Nvl(r_ҽ��.ҽ��״̬, 0);
      
        v_Temp := v_Temp || ',"serial_num":' || Nvl(r_ҽ��.���, 0);
        v_Temp := v_Temp || ',"advice_day":' || Nvl(r_ҽ��.����, 0);
        v_Temp := v_Temp || ',"advice_dosage":' || Zljsonstr(r_ҽ��.��������, 1);
        v_Temp := v_Temp || ',"oper_type":"' || Zljsonstr(r_ҽ��.��������) || '"';
        v_Temp := v_Temp || ',"clinic_type":"' || Zljsonstr(r_ҽ��.�������) || '"';
      
        v_Temp := v_Temp || ',"advice_exe_properties":' || Nvl(r_ҽ��.ִ������, 0);
        v_Temp := v_Temp || ',"advice_exe_sign":' || Nvl(r_ҽ��.ִ�б��, 0);
        v_Temp := v_Temp || ',"effective_time":' || Nvl(r_ҽ��.ҽ����Ч, 0);
        v_Temp := v_Temp || ',"advice_record_time":"' || Zljsonstr(To_Char(r_ҽ��.����ʱ��, 'yyyy-mm-dd hh24:mi:ss')) || '"';
        v_Temp := v_Temp || ',"advice_doctor":"' || Zljsonstr(r_ҽ��.����ҽ��) || '"';
      
        v_Temp := v_Temp || ',"advice_purpose":"' || Zljsonstr(r_ҽ��.��ҩĿ��) || '"';
        v_Temp := v_Temp || ',"advice_reason":"' || Zljsonstr(r_ҽ��.��ҩ����) || '"';
        v_Temp := v_Temp || ',"advice_taboonote":"' || Zljsonstr(r_ҽ��.����ҩƷ˵��) || '"';
        v_Temp := v_Temp || ',"advice_doctor_note":"' || Zljsonstr(r_ҽ��.ҽ������) || '"';
        v_Temp := v_Temp || ',"rcpdtl_excs_desc":"' || Zljsonstr(r_ҽ��.����˵��) || '"';
      
        v_Temp := v_Temp || ',"advice_audit_result":' || Nvl(r_ҽ��.�����, 0);
        v_Temp := v_Temp || ',"advice_audit_sign":' || Nvl(r_ҽ��.ҩʦ��˱�־, 0);
        v_Temp := v_Temp || ',"advice_audit_time":"' || Zljsonstr(To_Char(r_ҽ��.ҩʦ���ʱ��, 'yyyy-mm-dd hh24:mi:ss')) || '"';
        v_Temp := v_Temp || ',"advice_interval_unit":"' || Zljsonstr(r_ҽ��.�����λ) || '"';
        v_Temp := v_Temp || ',"advice_frequency":"' || Zljsonstr(r_ҽ��.ִ��Ƶ��) || '"';
      
        v_Temp := v_Temp || ',"advice_frequency_times":' || Zljsonstr(r_ҽ��.Ƶ�ʴ���, 1);
        v_Temp := v_Temp || ',"advice_frequency_interval":' || Zljsonstr(r_ҽ��.Ƶ�ʼ��, 1);
        v_Temp := v_Temp || ',"advice_exetime_plane":"' || Zljsonstr(r_ҽ��.ִ��ʱ�䷽��) || '"';
        v_Temp := v_Temp || ',"advice_begintime":"' || Zljsonstr(To_Char(r_ҽ��.��ʼִ��ʱ��, 'yyyy-mm-dd hh24:mi:ss')) || '"';
        v_Temp := v_Temp || ',"advice_endtime":"' || Zljsonstr(To_Char(r_ҽ��.ִ����ֹʱ��, 'yyyy-mm-dd hh24:mi:ss')) || '"';
      
        v_Temp := v_Temp || ',"rgst_no":"' || Zljsonstr(r_ҽ��.�Һŵ�) || '"';
        v_Temp := v_Temp || ',"advice_receipt_name":"' || Zljsonstr(r_ҽ��.�䷽����) || '"';
        v_Temp := v_Temp || ',"advice_receipt_issecret":' || Nvl(r_ҽ��.�Ƿ���, 0);
        v_Temp := v_Temp || ',"advice_note":"' || Zljsonstr(r_ҽ��.ҽ������) || '"';
        v_Temp := v_Temp || ',"advice_cisitem_id":' || Nvl(r_ҽ��.������Ŀid, 0);
      
        v_Temp := v_Temp || ',"advice_item_id":' || Nvl(r_ҽ��.�շ�ϸĿid, 0);
        v_Temp := v_Temp || ',"pati_deptid":' || Nvl(r_ҽ��.���˿���id, 0);
        v_Temp := v_Temp || ',"pati_name":"' || Zljsonstr(r_ҽ��.����) || '"';
        v_Temp := v_Temp || ',"pati_sex":"' || Zljsonstr(r_ҽ��.�Ա�) || '"';
        v_Temp := v_Temp || ',"pati_age":"' || Zljsonstr(r_ҽ��.����) || '"';
      
        v_Temp := v_Temp || ',"advice_audit_reason":"' || Zljsonstr(r_ҽ��.ҩʦ���ԭ��) || '"';
        If n_Ƥ�Ի�ȡ = 2 And r_ҽ��.Ƥ�Խ�� Is Null Then
          Pp_Test(r_ҽ��.������Ŀid, v_Ƥ�Խ��);
        Else
          v_Ƥ�Խ�� := r_ҽ��.Ƥ�Խ��;
        End If;
        v_Temp := v_Temp || ',"skintest_info":"' || Zljsonstr(v_Ƥ�Խ��) || '"';
      
        If n_��ѯ���� = 1 Then
          v_Temp := v_Temp || ',"total_qunt":' || Zljsonstr(r_ҽ��.�ܸ�����, 1);
          v_Temp := v_Temp || ',"Total_qunt_unit":"' || Zljsonstr(r_ҽ��.����) || '"';
          v_Temp := v_Temp || ',"single":"' || Zljsonstr(r_ҽ��.����) || '"';
          v_Temp := v_Temp || ',"toxicity_type":"' || Zljsonstr(r_ҽ��.�������) || '"';
          v_Temp := v_Temp || ',"advice_dept_id":' || Nvl(r_ҽ��.��������id, 0);
        
          v_Temp := v_Temp || ',"advice_stop_doctor":"' || Zljsonstr(r_ҽ��.ͣ��ҽ��) || '"';
          v_Temp := v_Temp || ',"advice_stop_nurse":"' || Zljsonstr(r_ҽ��.ȷ��ͣ����ʿ) || '"';
          v_Temp := v_Temp || ',"advice_stoptime":"' || Zljsonstr(To_Char(r_ҽ��.ͣ��ʱ��, 'yyyy-mm-dd hh24:mi:ss')) || '"';
          v_Temp := v_Temp || ',"advice_stoptime_confirm":"' ||
                    Zljsonstr(To_Char(r_ҽ��.ȷ��ͣ��ʱ��, 'yyyy-mm-dd hh24:mi:ss')) || '"';
          v_Temp := v_Temp || ',"order_chk_nurse":"' || Zljsonstr(r_ҽ��.У�Ի�ʿ) || '"';
        
          v_Temp := v_Temp || ',"order_chk_time":"' || Zljsonstr(To_Char(r_ҽ��.У��ʱ��, 'yyyy-mm-dd hh24:mi:ss')) || '"';
          v_Temp := v_Temp || ',"lastexe_time":"' || Zljsonstr(To_Char(r_ҽ��.�ϴ�ִ��ʱ��, 'yyyy-mm-dd hh24:mi:ss')) || '"';
          v_Temp := v_Temp || ',"usage":"' || Zljsonstr(r_ҽ��.�÷�) || '"';
          v_Temp := v_Temp || ',"emergency_tag":' || Nvl(r_ҽ��.������־, 0);
          v_Temp := v_Temp || ',"is_charge_verfy":' || Nvl(r_ҽ��.�Ƿ�������, 0);
        
          v_Temp := v_Temp || ',"baby_num":' || Nvl(r_ҽ��.Ӥ�����, 0);
          v_Temp := v_Temp || ',"valuation_nature":' || Nvl(r_ҽ��.�Ƽ�����, 0);
          v_Temp := v_Temp || ',"advice_exedept_id":' || Nvl(r_ҽ��.ִ�п���id, 0);
          v_Temp := v_Temp || ',"advice_exedept_name":"' || Zljsonstr(r_ҽ��.ִ�п�������) || '"';
          v_Temp := v_Temp || ',"testtube_code":"' || Zljsonstr(r_ҽ��.�Թܱ���) || '"';
        
          v_Temp := v_Temp || ',"hide_print":' || Nvl(r_ҽ��.���δ�ӡ, 0);
          v_Temp := v_Temp || ',"Prerequisite_id":' || Nvl(r_ҽ��.ǰ��id, 0);
          v_Temp := v_Temp || ',"is_staff_sig":' || Nvl(r_ҽ��.�Ƿ�ǩ��, 0);
          v_Temp := v_Temp || ',"rpt_id":' || Nvl(r_ҽ��.����id, 0);
          v_Temp := v_Temp || ',"consult_statu":' || Nvl(r_ҽ��.����״̬, 0);
        
          v_Temp := v_Temp || ',"advice_verfy_statu":' || Nvl(r_ҽ��.���״̬, 0);
          v_Temp := v_Temp || ',"apply_num":' || Nvl(r_ҽ��.�������, 0);
          v_Temp := v_Temp || ',"cisitem_unit":"' || Zljsonstr(r_ҽ��.������Ŀ���㵥λ) || '"';
          v_Temp := v_Temp || ',"pati_wardarea_id":' || Nvl(r_ҽ��.��ǰ����id, 0);
        
          If Nvl(r_ҽ��.���id, 0) <> 0 And (n_���id Is Null Or n_���id <> Nvl(r_ҽ��.���id, 0)) Then
            v_�巨 := Null;
            For r_�巨 In Ctest�巨(r_ҽ��.���id) Loop
              v_�巨 := r_�巨.����;
            End Loop;
          
            n_���id := r_ҽ��.���id;
          End If;
          v_Temp := v_Temp || ',"decoction_method":"' || Zljsonstr(v_�巨) || '"';
        End If;
        v_Temp := v_Temp || '}';
      
        If Length(v_Temp) > 25000 Then
          If c_Jtmp Is Null Then
            c_Jtmp := Substr(v_Temp, 2);
          Else
            c_Jtmp := c_Jtmp || v_Temp;
          End If;
          v_Temp := Null;
        End If;
      End If;
    End Loop;
  End Loop;

  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","advice_list":[' || Substr(v_Temp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Temp;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","advice_list":[' || c_Jtmp || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getadviceinfo;
/
Create Or Replace Procedure Zl_Cissvr_Getadviceoperinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:��ȡҽ����صĲ���˵��
  --��Σ�Json_In:��ʽ
  -- input
  --   advice_id            C 1 ҽ��id,����ö��ŷָ�
  --   oper_type            N 1 ��������:1-�¿���2-У�����ʣ�3-У��ͨ����4-���ϣ�5-������
  --                                     6-��ͣ��7-���ã�8-ֹͣ��9-ȷ��ֹͣ��10-Ƥ�Խ����
  --                                     11-���ͨ����12-���δͨ����13-ʵϰҽʦͣ�������ˣ�14-Ѫ����գ�15-Ѫ�����ͨ����
  --                                     16-Ѫ����Ѫ�ܾ���17-Ѫ��ֹͣ��Ѫ��18-��Ѫ����ͨ����ǩ����9-��Ѫ������ˣ�20-��Ѫҽ�����δ��



  --   oper_last            N 1 �Ƿ�ȡ���һ�β�����1-ȡ���һ��,0-ȡ����
  --����: Json_Out,��ʽ����
  --  output
  --    code               N  1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message            C  1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    oper_list[]����
  --      advice_id        N  1 ҽ��ID
  --      oper_time        C  1 ����ʱ��:yyyy-mm-dd hh24:mi:ss
  --      oper_note        C  1 ����˵��
  ---------------------------------------------------------------------------
  j_Json   Pljson;
  j_In     Pljson;
  c_ҽ��id Clob;
  l_ҽ��id t_Strlist := t_Strlist();

  n_�������� ����ҽ��״̬.��������%Type;
  n_Last     Number(2);
  v_Temp     Varchar2(32767);
  c_Jtmp     Clob;
Begin
  --�������
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');
  If j_Json.Exist('advice_id') Then
    c_ҽ��id := j_Json.Get_Clob('advice_id');
  End If;
  n_�������� := j_Json.Get_Number('oper_type');
  n_Last     := j_Json.Get_Number('oper_last');

  While c_ҽ��id Is Not Null Loop
    If Length(c_ҽ��id) <= 4000 Then
      l_ҽ��id.Extend;
      l_ҽ��id(l_ҽ��id.Count) := c_ҽ��id;
      c_ҽ��id := Null;
    Else
      l_ҽ��id.Extend;
      l_ҽ��id(l_ҽ��id.Count) := Substr(c_ҽ��id, 1, Instr(c_ҽ��id, ',', 3950) - 1);
      c_ҽ��id := Substr(c_ҽ��id, Instr(c_ҽ��id, ',', 3950) + 1);
    End If;
  End Loop;

  v_Temp := Null;
  For I In 1 .. l_ҽ��id.Count Loop
    For r_ҽ�� In (Select /*+cardinality(b,10)*/
                  a.ҽ��id, Nvl(a.����˵��, '��') As ����˵��, To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��
                 From ����ҽ��״̬ A, Table(f_Num2list(l_ҽ��id(I))) B
                 Where a.ҽ��id = b.Column_Value And a.�������� = n_��������
                 Order By ����ʱ�� Desc) Loop
    
      v_Temp := v_Temp || ',{"advice_id":' || r_ҽ��.ҽ��id;
      v_Temp := v_Temp || ',"oper_time":"' || r_ҽ��.����ʱ�� || '"';
      v_Temp := v_Temp || ',"oper_note":"' || Zljsonstr(r_ҽ��.����˵��) || '"';
      v_Temp := v_Temp || '}';
    
      If Nvl(n_Last, 0) = 1 Then
        Exit; --ֻȡ���һ��
      End If;
    
      If Length(v_Temp) > 30000 Then
        If c_Jtmp Is Null Then
          c_Jtmp := Substr(v_Temp, 2);
        Else
          c_Jtmp := c_Jtmp || v_Temp;
        End If;
        v_Temp := Null;
      End If;
    End Loop;
    If Nvl(n_Last, 0) = 1 Then
      Exit; --ֻȡ���һ��
    End If;
  End Loop;
  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || Substr(v_Temp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Temp;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getadviceoperinfo;
/
Create Or Replace Procedure Zl_Cissvr_Getadvicesendinfo
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡҽ��������Ϣ
  --���      json
  --input
  --  query_type                N 1 ��ѯ���ͣ�0-��ҽ��id����advice_ids����ѯ;1-��ҽ��ID+ҽ�����ͺŲ�ѯ;2-��ҽ��ID+��¼����+NO��ѯ;3-����ҽ�����ͺŲ�ѯ
  --  return_type               N 1 �������ͣ�0-���ػ�����Ϣ��1-���ػ�����Ϣ+��չ��Ϣ
  --  contain_related           N 1 �Ƿ�������ҽ����¼
  --  advice_ids                C 1 ҽ��ID����֧�ֶ��ҽ��ID���á������ָ�
  --  send_nos                  C 1 ҽ�����ͺŴ���֧�ֶ�����á������ָ�,����ѯ����Ϊ3ʱ��Ч
  --  item_list[]
  --    advice_id               N 1 ҽ��ID
  --    send_no                 N 1 ҽ�����ͺ�
  --    bill_no                 C 1 NO
  --    bill_prop               N 1 ��¼����
  --����      json
  --output
  --  code                      C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message                   C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  advice_send_list[]
  --    advice_id               N   ҽ��id
  --    send_no                 N   ���ͺ�
  --    register_no             C   �Һŵ���
  --    pati_id                 N   ����id
  --    pati_pageid             N   ��ҳid
  --    pati_deptid             N   ���˿���id
  --    advice_dept_id          N   ��������ID
  --    pati_source             N   ������Դ
  --    clinic_type             C   �������
  --    valuation_nature        N   �Ƽ�����:0-�����Ƽۣ�1-���Ƽۣ�2-�ֹ��Ƽ�
  --    advice_related_id       N   ���id
  --    outpati_account         N   �Ƿ��������
  --    advice_note             C   ҽ������
  --    sample_barcode          C   ��������
  --    bill_no                 C   No
  --    bill_prop               N   ��¼����
  --    advice_send_firsttime   D   �״�ʱ��
  --    advice_send_endtime     D   ĩ��ʱ��
  --    advice_send_exestatus   N   ִ��״̬
  --    advice_send_sendtime    D   ����ʱ��
  --    effective_time          N   ҽ����Ч
  ---------------------------------------------------------------------------
  v_ҽ��id Clob; --��¼ҽ��id
  v_���ͺ� Clob;
  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_ҽ��id Collection_Type;
  Col_���ͺ� Collection_Type;

  I            Number;
  n_��ѯ����   Number(1);
  n_��������   Number(1);
  n_�����ҽ�� Number(1);
  j_In         Pljson;
  j_Json       Pljson;
  j_List       Pljson_List := Pljson_List();
  j_Jsonlist   Pljson_List := Pljson_List();
  j_Json_Tmp   Pljson;

  v_No       ����ҽ������.No%Type;
  n_��¼���� ����ҽ������.��¼����%Type;
  n_ҽ��id   ����ҽ������.ҽ��id%Type;
  n_���ͺ�   ����ҽ������.���ͺ�%Type;
  v_Tmp      Varchar2(32767);
  v_Tmp2     Varchar2(32767);
  c_Temp     Clob; --����json�ڵ�

  Cursor c_ҽ��������Ϣ Is
    Select a.ҽ��id, a.���ͺ�, a.�������, a.No, a.��¼����, a.�״�ʱ��, a.ĩ��ʱ��, a.ִ��״̬, a.����ʱ��, a.��������, b.�Һŵ�, b.����id, b.��ҳid, b.���˿���id,
           b.��������id, b.������Դ, b.�������, b.���id, b.ҽ������, c.�Ƽ�����, b.ҽ����Ч
    From ����ҽ������ A, ����ҽ����¼ B, ������ĿĿ¼ C
    Where a.ҽ��id = b.Id And b.������Ŀid = c.Id And Rownum < 1;
  r_Advice_Send c_ҽ��������Ϣ%RowType;

  Type Ty_Advice_Send Is Ref Cursor;
  c_Advice_Send Ty_Advice_Send; --��̬�α����

  --��װʧ��ʱ���ص�����
  Function Get_Err_Message(Message_In Varchar2) Return Clob Is
  Begin
    Return '{"output":{"code":0,"message":"' || Zltools.Zljsonstr(Message_In) || '"}}';
  End Get_Err_Message;
Begin
  j_In         := Pljson(Json_In);
  j_Json       := j_In.Get_Pljson('input');
  n_��ѯ����   := Nvl(j_Json.Get_Number('query_type'), 0);
  n_��������   := Nvl(j_Json.Get_Number('return_type'), 0);
  n_�����ҽ�� := Nvl(j_Json.Get_Number('contain_related'), 0);

  If n_��ѯ���� = 0 Then
    Begin
      v_ҽ��id := j_Json.Get_Clob('advice_ids');
    Exception
      When Others Then
        Json_Out := Get_Err_Message('δ����ҽ��id�����飡');
        Return;
    End;
  Elsif n_��ѯ���� = 3 Then
    Begin
      v_���ͺ� := j_Json.Get_Clob('send_nos');
    Exception
      When Others Then
        Json_Out := Get_Err_Message('δ����ҽ�����ͺţ����飡');
        Return;
    End;
  Else
    j_List := j_Json.Get_Pljson_List('item_list');
    If j_List Is Null Then
      Json_Out := Get_Err_Message('δ����ҽ��id����Ϣ�����飡');
      Return;
    End If;
  End If;

  If n_��ѯ���� = 0 Then
    --�� v_ҽ��id ����װ�ɲ�����4000 �ļ��ϴ�����ֹʹ�� f_Num2list ��������
    I := 0;
    While v_ҽ��id Is Not Null Loop
      If Length(v_ҽ��id) <= 4000 Then
        Col_ҽ��id(I) := v_ҽ��id;
        v_ҽ��id := Null;
      Else
        Col_ҽ��id(I) := Substr(v_ҽ��id, 1, Instr(v_ҽ��id, ',', 3980) - 1);
        v_ҽ��id := Substr(v_ҽ��id, Instr(v_ҽ��id, ',', 3980) + 1);
      End If;
      I := I + 1;
    End Loop;
    I := 0;
  
    For I In 0 .. Col_ҽ��id.Count - 1 Loop
      Open c_Advice_Send For
        Select /*+cardinality(j,10)*/
         a.ҽ��id, a.���ͺ�, a.�������, a.No, a.��¼����, a.�״�ʱ��, a.ĩ��ʱ��, a.ִ��״̬, a.����ʱ��, a.��������, b.�Һŵ�, b.����id, b.��ҳid, b.���˿���id,
         b.��������id, b.������Դ, b.�������, b.���id, b.ҽ������, c.�Ƽ�����, b.ҽ����Ч
        From ����ҽ������ A, ����ҽ����¼ B, ������ĿĿ¼ C, Table(f_Num2list(Col_ҽ��id(I))) J
        Where a.ҽ��id = b.Id And b.������Ŀid = c.Id And b.Id = j.Column_Value
        Union All
        Select /*+cardinality(j,10)*/
         a.ҽ��id, a.���ͺ�, a.�������, a.No, a.��¼����, a.�״�ʱ��, a.ĩ��ʱ��, a.ִ��״̬, a.����ʱ��, a.��������, b.�Һŵ�, b.����id, b.��ҳid, b.���˿���id,
         b.��������id, b.������Դ, b.�������, b.���id, b.ҽ������, c.�Ƽ�����, b.ҽ����Ч
        From ����ҽ������ A, ����ҽ����¼ B, ������ĿĿ¼ C, Table(f_Num2list(Col_ҽ��id(I))) J
        Where a.ҽ��id = b.Id And b.������Ŀid = c.Id And n_�����ҽ�� = 1 And b.���id = j.Column_Value;
    
      Loop
        Fetch c_Advice_Send
          Into r_Advice_Send;
        Exit When c_Advice_Send %NotFound;
      
        --������Ϣ
        v_Tmp := '';
        Zljsonputvalue(v_Tmp, 'advice_id', r_Advice_Send.ҽ��id, 1, 1);
        Zljsonputvalue(v_Tmp, 'send_no', r_Advice_Send.���ͺ�, 1);
        Zljsonputvalue(v_Tmp, 'bill_no', r_Advice_Send.No);
        Zljsonputvalue(v_Tmp, 'bill_prop', r_Advice_Send.��¼����, 1);
        Zljsonputvalue(v_Tmp, 'advice_send_firsttime', r_Advice_Send.�״�ʱ��);
        Zljsonputvalue(v_Tmp, 'advice_send_endtime', r_Advice_Send.ĩ��ʱ��);
        Zljsonputvalue(v_Tmp, 'advice_send_exestatus', r_Advice_Send.ִ��״̬, 1);
      
        If n_�������� = 1 Then
          Zljsonputvalue(v_Tmp, 'advice_send_sendtime', r_Advice_Send.����ʱ��);
        
          --��չ��Ϣ
          Zljsonputvalue(v_Tmp, 'register_no', r_Advice_Send.�Һŵ�);
          Zljsonputvalue(v_Tmp, 'pati_id', r_Advice_Send.����id, 1);
          Zljsonputvalue(v_Tmp, 'pati_pageid', r_Advice_Send.��ҳid, 1);
          Zljsonputvalue(v_Tmp, 'pati_deptid', r_Advice_Send.���˿���id, 1);
          Zljsonputvalue(v_Tmp, 'advice_dept_id', r_Advice_Send.��������id, 1);
        
          Zljsonputvalue(v_Tmp, 'pati_source', r_Advice_Send.������Դ, 1);
          Zljsonputvalue(v_Tmp, 'clinic_type', r_Advice_Send.�������);
          Zljsonputvalue(v_Tmp, 'valuation_nature', r_Advice_Send.�Ƽ�����, 1);
          Zljsonputvalue(v_Tmp, 'advice_related_id', r_Advice_Send.���id, 1);
          Zljsonputvalue(v_Tmp, 'outpati_account', r_Advice_Send.�������, 1);
        
          Zljsonputvalue(v_Tmp, 'advice_note', Nvl(r_Advice_Send.ҽ������, ''));
          Zljsonputvalue(v_Tmp, 'sample_barcode', Nvl(r_Advice_Send.��������, ''));
          Zljsonputvalue(v_Tmp, 'effective_time', r_Advice_Send.ҽ����Ч, 1, 2);
        Else
          Zljsonputvalue(v_Tmp, 'advice_send_sendtime', r_Advice_Send.����ʱ��, 0, 2);
        End If;
      
        If (Length(v_Tmp) + Length(v_Tmp2)) > 32000 Then
          c_Temp := c_Temp || v_Tmp2 || ',' || v_Tmp;
          v_Tmp2 := Null;
        Else
          v_Tmp2 := v_Tmp2 || ',' || v_Tmp;
        End If;
      End Loop;
    End Loop;
  
  Elsif n_��ѯ���� = 3 Then
    --�� v_���ͺ� ����װ�ɲ�����4000 �ļ��ϴ�����ֹʹ�� f_Num2list ��������
    I := 0;
    While v_���ͺ� Is Not Null Loop
      If Length(v_���ͺ�) <= 4000 Then
        Col_���ͺ�(I) := v_���ͺ�;
        v_���ͺ� := Null;
      Else
        Col_���ͺ�(I) := Substr(v_���ͺ�, 1, Instr(v_���ͺ�, ',', 3980) - 1);
        v_���ͺ� := Substr(v_���ͺ�, Instr(v_���ͺ�, ',', 3980) + 1);
      End If;
      I := I + 1;
    End Loop;
    I := 0;
  
    For I In 0 .. Col_���ͺ�.Count - 1 Loop
      Open c_Advice_Send For
        Select /*+cardinality(j,10)*/
         a.ҽ��id, a.���ͺ�, a.�������, a.No, a.��¼����, a.�״�ʱ��, a.ĩ��ʱ��, a.ִ��״̬, a.����ʱ��, a.��������, b.�Һŵ�, b.����id, b.��ҳid, b.���˿���id,
         b.��������id, b.������Դ, b.�������, b.���id, b.ҽ������, c.�Ƽ�����, b.ҽ����Ч
        From ����ҽ������ A, ����ҽ����¼ B, ������ĿĿ¼ C, Table(f_Num2list(Col_���ͺ�(I))) J
        Where a.ҽ��id = b.Id And b.������Ŀid = c.Id And a.���ͺ� = j.Column_Value;
    
      Loop
        Fetch c_Advice_Send
          Into r_Advice_Send;
        Exit When c_Advice_Send %NotFound;
      
        --������Ϣ
        v_Tmp := '';
        Zljsonputvalue(v_Tmp, 'advice_id', r_Advice_Send.ҽ��id, 1, 1);
        Zljsonputvalue(v_Tmp, 'send_no', r_Advice_Send.���ͺ�, 1);
        Zljsonputvalue(v_Tmp, 'bill_no', r_Advice_Send.No);
        Zljsonputvalue(v_Tmp, 'bill_prop', r_Advice_Send.��¼����, 1);
        Zljsonputvalue(v_Tmp, 'advice_send_firsttime', r_Advice_Send.�״�ʱ��);
        Zljsonputvalue(v_Tmp, 'advice_send_endtime', r_Advice_Send.ĩ��ʱ��);
        Zljsonputvalue(v_Tmp, 'advice_send_exestatus', r_Advice_Send.ִ��״̬, 1);
      
        If n_�������� = 1 Then
          Zljsonputvalue(v_Tmp, 'advice_send_sendtime', r_Advice_Send.����ʱ��);
        
          --��չ��Ϣ
          Zljsonputvalue(v_Tmp, 'register_no', r_Advice_Send.�Һŵ�);
          Zljsonputvalue(v_Tmp, 'pati_id', r_Advice_Send.����id, 1);
          Zljsonputvalue(v_Tmp, 'pati_pageid', r_Advice_Send.��ҳid, 1);
          Zljsonputvalue(v_Tmp, 'pati_deptid', r_Advice_Send.���˿���id, 1);
          Zljsonputvalue(v_Tmp, 'advice_dept_id', r_Advice_Send.��������id, 1);
        
          Zljsonputvalue(v_Tmp, 'pati_source', r_Advice_Send.������Դ, 1);
          Zljsonputvalue(v_Tmp, 'clinic_type', r_Advice_Send.�������);
          Zljsonputvalue(v_Tmp, 'valuation_nature', r_Advice_Send.�Ƽ�����, 1);
          Zljsonputvalue(v_Tmp, 'advice_related_id', r_Advice_Send.���id, 1);
          Zljsonputvalue(v_Tmp, 'outpati_account', r_Advice_Send.�������, 1);
        
          Zljsonputvalue(v_Tmp, 'advice_note', Nvl(r_Advice_Send.ҽ������, ''));
          Zljsonputvalue(v_Tmp, 'sample_barcode', Nvl(r_Advice_Send.��������, ''));
          Zljsonputvalue(v_Tmp, 'effective_time', r_Advice_Send.ҽ����Ч, 1, 2);
        Else
          Zljsonputvalue(v_Tmp, 'advice_send_sendtime', r_Advice_Send.����ʱ��, 0, 2);
        End If;
      
        If (Length(v_Tmp) + Length(v_Tmp2)) > 32000 Then
          c_Temp := c_Temp || v_Tmp2 || ',' || v_Tmp;
          v_Tmp2 := Null;
        Else
          v_Tmp2 := v_Tmp2 || ',' || v_Tmp;
        End If;
      End Loop;
    End Loop;
  Else
    For I In 1 .. j_List.Count Loop
      j_Json_Tmp := Pljson();
      j_Json_Tmp := Pljson(j_List.Get(I));
    
      If n_��ѯ���� = 1 Then
        n_ҽ��id   := j_Json_Tmp.Get_Number('advice_id');
        n_���ͺ�   := j_Json_Tmp.Get_Number('send_no');
        n_��¼���� := Nvl(j_Json_Tmp.Get_Number('bill_prop'), 0);
      
        Open c_Advice_Send For
          Select a.ҽ��id, a.���ͺ�, a.�������, a.No, a.��¼����, a.�״�ʱ��, a.ĩ��ʱ��, a.ִ��״̬, a.����ʱ��, a.��������, b.�Һŵ�, b.����id, b.��ҳid,
                 b.���˿���id, b.��������id, b.������Դ, b.�������, b.���id, b.ҽ������, c.�Ƽ�����, b.ҽ����Ч
          From ����ҽ������ A, ����ҽ����¼ B, ������ĿĿ¼ C
          Where a.ҽ��id = b.Id And b.������Ŀid = c.Id And b.Id = n_ҽ��id And a.���ͺ� = n_���ͺ�
          Union All
          Select a.ҽ��id, a.���ͺ�, a.�������, a.No, a.��¼����, a.�״�ʱ��, a.ĩ��ʱ��, a.ִ��״̬, a.����ʱ��, a.��������, b.�Һŵ�, b.����id, b.��ҳid,
                 b.���˿���id, b.��������id, b.������Դ, b.�������, b.���id, b.ҽ������, c.�Ƽ�����, b.ҽ����Ч
          From ����ҽ������ A, ����ҽ����¼ B, ������ĿĿ¼ C
          Where a.ҽ��id = b.Id And b.������Ŀid = c.Id And n_�����ҽ�� = 1 And b.���id = n_ҽ��id And a.���ͺ� = n_���ͺ�;
      End If;
    
      If n_��ѯ���� = 2 Then
        n_ҽ��id   := j_Json_Tmp.Get_Number('advice_id');
        v_No       := j_Json_Tmp.Get_String('bill_no');
        n_��¼���� := j_Json_Tmp.Get_Number('bill_prop');
      
        Open c_Advice_Send For
          Select a.ҽ��id, a.���ͺ�, a.�������, a.No, a.��¼����, a.�״�ʱ��, a.ĩ��ʱ��, a.ִ��״̬, a.����ʱ��, a.��������, b.�Һŵ�, b.����id, b.��ҳid,
                 b.���˿���id, b.��������id, b.������Դ, b.�������, b.���id, b.ҽ������, c.�Ƽ�����, b.ҽ����Ч
          From ����ҽ������ A, ����ҽ����¼ B, ������ĿĿ¼ C
          Where a.ҽ��id = b.Id And b.������Ŀid = c.Id And b.Id = n_ҽ��id And a.No = v_No And a.��¼���� = n_��¼����
          Union All
          Select a.ҽ��id, a.���ͺ�, a.�������, a.No, a.��¼����, a.�״�ʱ��, a.ĩ��ʱ��, a.ִ��״̬, a.����ʱ��, a.��������, b.�Һŵ�, b.����id, b.��ҳid,
                 b.���˿���id, b.��������id, b.������Դ, b.�������, b.���id, b.ҽ������, c.�Ƽ�����, b.ҽ����Ч
          From ����ҽ������ A, ����ҽ����¼ B, ������ĿĿ¼ C
          Where a.ҽ��id = b.Id And b.������Ŀid = c.Id And n_�����ҽ�� = 1 And b.���id = n_ҽ��id And a.No = v_No And
                a.��¼���� = n_��¼����;
      End If;
    
      Loop
        Fetch c_Advice_Send
          Into r_Advice_Send;
        Exit When c_Advice_Send %NotFound;
      
        --������Ϣ
        v_Tmp := '';
        Zljsonputvalue(v_Tmp, 'advice_id', r_Advice_Send.ҽ��id, 1, 1);
        Zljsonputvalue(v_Tmp, 'send_no', r_Advice_Send.���ͺ�, 1);
        Zljsonputvalue(v_Tmp, 'bill_no', r_Advice_Send.No);
        Zljsonputvalue(v_Tmp, 'bill_prop', r_Advice_Send.��¼����, 1);
        Zljsonputvalue(v_Tmp, 'advice_send_firsttime', r_Advice_Send.�״�ʱ��);
        Zljsonputvalue(v_Tmp, 'advice_send_endtime', r_Advice_Send.ĩ��ʱ��);
        Zljsonputvalue(v_Tmp, 'advice_send_exestatus', r_Advice_Send.ִ��״̬, 1);
      
        If n_�������� = 1 Then
          Zljsonputvalue(v_Tmp, 'advice_send_sendtime', r_Advice_Send.����ʱ��);
        
          --��չ��Ϣ
          Zljsonputvalue(v_Tmp, 'register_no', r_Advice_Send.�Һŵ�);
          Zljsonputvalue(v_Tmp, 'pati_id', r_Advice_Send.����id, 1);
          Zljsonputvalue(v_Tmp, 'pati_pageid', r_Advice_Send.��ҳid, 1);
          Zljsonputvalue(v_Tmp, 'pati_deptid', r_Advice_Send.���˿���id, 1);
          Zljsonputvalue(v_Tmp, 'advice_dept_id', r_Advice_Send.��������id, 1);
        
          Zljsonputvalue(v_Tmp, 'pati_source', r_Advice_Send.������Դ, 1);
          Zljsonputvalue(v_Tmp, 'clinic_type', r_Advice_Send.�������);
          Zljsonputvalue(v_Tmp, 'valuation_nature', r_Advice_Send.�Ƽ�����, 1);
          Zljsonputvalue(v_Tmp, 'advice_related_id', r_Advice_Send.���id, 1);
          Zljsonputvalue(v_Tmp, 'outpati_account', r_Advice_Send.�������, 1);
        
          Zljsonputvalue(v_Tmp, 'advice_note', Nvl(r_Advice_Send.ҽ������, ''));
          Zljsonputvalue(v_Tmp, 'sample_barcode', Nvl(r_Advice_Send.��������, ''));
          Zljsonputvalue(v_Tmp, 'effective_time', r_Advice_Send.ҽ����Ч, 1, 2);
        Else
          Zljsonputvalue(v_Tmp, 'advice_send_sendtime', r_Advice_Send.����ʱ��, 0, 2);
        End If;
      
        If (Length(v_Tmp) + Length(v_Tmp2)) > 32000 Then
          c_Temp := c_Temp || v_Tmp2 || ',' || v_Tmp;
          v_Tmp2 := Null;
        Else
          v_Tmp2 := v_Tmp2 || ',' || v_Tmp;
        End If;
      End Loop;
    End Loop;
  End If;

  If v_Tmp2 Is Not Null Then
    c_Temp := c_Temp || v_Tmp2;
  End If;

  Json_Out := '{"output":{"code":1,"message": "�ɹ�","advice_send_list":[' || Substr(c_Temp, 2) || ']}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getadvicesendinfo;
/
Create Or Replace Procedure Zl_Cissvr_Getadvicesendnums
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:��ȡҽ�����͵�������Ϣ
  --��Σ�Json_In:��ʽ
  -- input
  --   advice_id            N 1 ҽ��id
  --   advice_sendno        N 1 ���ͺ�
  --   isalone_exe          N 1 �Ƿ����ִ��:1-����ִ��;0-�Ƕ���ִ��

  --����: Json_Out,��ʽ����
  --  output
  --    code                  N  1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message               C  1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    sendlist              C   ������
  --      advice_num          N 1 ���:����ҽ����¼.���
  --      advice_id           N 1 ҽ��id
  --      advice_related_id   N 1 ���id
  --      clinic_type         C 1 �������
  --      advice_cisitem_id   N 1 ������Ŀid
  --      advice_dept_id      N 1 ��������ID
  --      exe_deptid          N 1 ִ�в���ID
  --      nums                N 1 ����
  --      citem_spcm_parts    C 1 �걾��λ
  --      citem_exam_method   C 1 ��鷽��
  --      advice_exe_sign     N 1 ִ�б��
  --      pati_id             N 1 ����id
  --      pati_pageid         N 1 ��ҳid
  ---------------------------------------------------------------------------
  j_In       Pljson;
  j_Json     Pljson;
  n_ҽ��id   ����ҽ����¼.Id%Type;
  n_���id   ����ҽ����¼.���id%Type;
  n_����ִ�� Number(1);
  n_���ͺ�   ����ҽ������.���ͺ�%Type;

  v_Jtmp Varchar2(32767);

Begin
  --�������
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  n_ҽ��id   := j_Json.Get_Number('advice_id');
  n_����ִ�� := Nvl(j_Json.Get_Number('isalone_exe'), 0);
  n_���ͺ�   := j_Json.Get_Number('advice_sendno');

  If Nvl(n_ҽ��id, 0) = 0 And Nvl(n_���ͺ�, 0) = 0 Then
    Json_Out := Zljsonout('δ����ҽ��id���ͺţ����飡');
    Return;
  End If;

  Select Decode(a.�������, 'D', Nvl(a.���id, a.Id), a.���id)
  Into n_���id
  From ����ҽ����¼ A
  Where a.Id = n_ҽ��id;

  For r_ҽ������ In (Select b.���, a.ҽ��id, b.���id, b.�������, b.������Ŀid, b.��������id, a.ִ�в���id,
                        Nvl(a.��������, Sum(Nvl(c.��������, 0))) As ����, b.�걾��λ, b.��鷽��, b.ִ�б��, b.����id, b.��ҳid
                 From ����ҽ������ A, ����ҽ����¼ B, ����ҽ��ִ�� C
                 Where Nvl(a.�Ʒ�״̬, 0) = 0 And b.Id = n_ҽ��id And
                       (b.Id = n_ҽ��id Or
                        ((b.���id = n_ҽ��id And b.������� In ('F', 'D')) Or (b.���id = n_���id And b.������� = 'C') And n_����ִ�� = 0)) And
                       a.���ͺ� = n_���ͺ� And c.ҽ��id(+) = a.ҽ��id And c.���ͺ�(+) = a.���ͺ�
                 Group By b.���, a.ҽ��id, b.���id, b.�������, b.������Ŀid, b.��������id, a.ִ�в���id, a.��������, b.�걾��λ, b.��鷽��, b.ִ�б��,
                          b.����id, b.��ҳid
                 Having Nvl(a.��������, Sum(Nvl(c.��������, 0))) <> 0
                 Order By ���) Loop
  
    --      advice_num          N 1 ���:����ҽ����¼.���
    --      advice_id           N 1 ҽ��id
    --      advice_related_id   N 1 ���id
    --      clinic_type         C 1 �������
    --      advice_cisitem_id   N 1 ������Ŀid
    --      advice_dept_id      N 1 ��������ID
    --      exe_deptid          N 1 ִ�в���ID
    v_Jtmp := v_Jtmp || ',{"advice_num":' || r_ҽ������.���;
    v_Jtmp := v_Jtmp || ',"advice_id":' || r_ҽ������.ҽ��id;
    v_Jtmp := v_Jtmp || ',"advice_related_id":' || Nvl(r_ҽ������.���id || '', 'null');
    v_Jtmp := v_Jtmp || ',"clinic_type":"' || r_ҽ������.������� || '"';
    v_Jtmp := v_Jtmp || ',"advice_cisitem_id":' || Nvl(r_ҽ������.������Ŀid || '', 'null');
    v_Jtmp := v_Jtmp || ',"advice_dept_id":' || Nvl(r_ҽ������.��������id || '', 'null');
    v_Jtmp := v_Jtmp || ',"exe_deptid":' || Nvl(r_ҽ������.ִ�в���id || '', 'null');
  
    --      nums                N 1 ����
    --      citem_spcm_parts    C 1 �걾��λ
    --      citem_exam_method   C 1 ��鷽��
    --      advice_exe_sign     N 1 ִ�б��
    --      pati_id             N 1 ����id
    --      pati_pageid         N 1 ��ҳid
    v_Jtmp := v_Jtmp || ',"nums":' || Zljsonstr(r_ҽ������.����, 1);
    v_Jtmp := v_Jtmp || ',"citem_spcm_parts":"' || r_ҽ������.�걾��λ || '"';
    v_Jtmp := v_Jtmp || ',"citem_exam_method":"' || Zljsonstr(r_ҽ������.��鷽��) || '"';
    v_Jtmp := v_Jtmp || ',"advice_exe_sign":' || Nvl(r_ҽ������.ִ�б��, 0);
    v_Jtmp := v_Jtmp || ',"pati_id":' || r_ҽ������.����id;
    v_Jtmp := v_Jtmp || ',"pati_pageid":' || Nvl(r_ҽ������.��ҳid || '', 'null');
  
    v_Jtmp := v_Jtmp || '}';
  
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","sendlist":[' || Substr(v_Jtmp, 2) || ']}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getadvicesendnums;
/
Create Or Replace Procedure Zl_Cissvr_Getaffirmerrordata
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ��ٴ�ҽ��ִ�����ʱ�Զ���˷��ã��쳣��δ��ҩƷ�����Ľ����շ�ȷ�ϣ���Դ����쳣���ݻ�ȡ����
  --��Σ�Json_In:��ʽ
  --  input
  --      pati_list[]���˹ؼ���Ϣ�����ڻ�ȡҽ��
  --           pati_id                    N 1 ����id
  --           pati_pageid                N 1 ��ҳid��סԺ���˴��룬���ﴫ0
  --           rgst_id                    N 1 �Һ�id�����ﲡ�˴��룬סԺ���˴���
  --           rgst_no                    C 1 �Һŵ���
  --����: Json_Out,��ʽ����
  --   output:
  --    code                  N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message               C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --     pati_bill_list[]
  --         pati_id                      N 1 ����id
  --         pati_pageid                  N 0 ��ҳid��סԺ���˴��룬���ﴫ0
  --         rgst_id                      N 0 �Һ�id�����ﲡ�˴��룬סԺ���˴���
  --         rgst_no                      C 0 �Һŵ���
  --         order_ids                    C 1 ���쳣������ҽ��idƴ��
  --         fee_nos                      C 1 ���쳣�����е��ݺ�ƴ��
  --         order_list[]ҽ��������Ϣ�б�
  --             send_no                  N 1 ���ͺ�
  --             advice_id                N 1 ҽ��id
  --             fee_no                   C 1 ���ݺ�
  --             bill_prop                N 1 ��¼����
  --             outpati_account          N 1 �Ƿ�������� 0-����������ʣ�1-���������
  --             pati_source              N 1 ������Դ 1-����ҽ����2-סԺҽ��
  ----------------------------------------------------------------------------------

  j_Json        Pljson;
  v_Rgs_No      Varchar2(2000);
  Jl_All_In     Pljson_List := Pljson_List();
  j_Item_a      Pljson;
  n_Pati_Id     Number; --N   1����ID
  n_Pati_Pageid Number; --N   1��ҳID
  n_Rgst_Id     Number; --N   1�Һŵ�id
  v_ҽ��ids     Varchar2(32767);
  v_Nos         Varchar2(32767);
  v_Item_a      Varchar2(32767);
  c_Item_a      Clob;
  v_Item        Varchar2(32767);
  c_Item        Clob;
  Cursor c_Out Is
    Select b.ҽ��id, b.No, b.���ͺ�, b.��¼����, b.�������, 1 As ��Դ
    From ����ҽ����¼ A, ����ҽ������ B, ����ҽ���쳣��¼ C
    Where a.Id = b.ҽ��id And b.ҽ��id = c.ҽ��id And b.���ͺ� = c.���ͺ� And a.�Һŵ� = v_Rgs_No And c.�������� In (4, 5)
    Order By b.���ͺ�;

  Type t_Order Is Table Of c_Out%RowType;
  r_Odr t_Order;

  Cursor c_In Is
    Select b.ҽ��id, b.No, b.���ͺ�, b.��¼����, b.�������, 2 As ��Դ
    From ����ҽ����¼ A, ����ҽ������ B, ����ҽ���쳣��¼ C
    Where a.Id = b.ҽ��id And b.ҽ��id = c.ҽ��id And b.���ͺ� = c.���ͺ� And a.����id = n_Pati_Id And a.��ҳid = n_Pati_Pageid And
          c.�������� In (4, 5)
    Order By b.���ͺ�;

  Procedure p_Getbaseinfo As
  Begin
    If Nvl(n_Rgst_Id, 0) = 0 Then
      Open c_In;
      Fetch c_In Bulk Collect
        Into r_Odr;
      Close c_In;
    Else
      Open c_Out;
      Fetch c_Out Bulk Collect
        Into r_Odr;
      Close c_Out;
    End If;
  End p_Getbaseinfo;

Begin
  --�������
  j_Item_a  := Pljson(Json_In);
  j_Json    := j_Item_a.Get_Pljson('input');
  Jl_All_In := j_Json.Get_Pljson_List('pati_list');
  For Pi In 1 .. Jl_All_In.Count Loop
    j_Item_a      := Pljson();
    j_Item_a      := Pljson(Jl_All_In.Get(Pi));
    n_Pati_Id     := j_Item_a.Get_Number('pati_id');
    n_Pati_Pageid := j_Item_a.Get_Number('pati_pageid');
    n_Rgst_Id     := j_Item_a.Get_Number('rgst_id');
    v_Rgs_No      := j_Item_a.Get_String('rgst_no');
    p_Getbaseinfo;
    If r_Odr.Count > 0 Then
      For Ol In 1 .. r_Odr.Count Loop
      
        If Instr(',' || v_ҽ��ids || ',', ',' || r_Odr(Ol).ҽ��id || ',') = 0 Then
          v_ҽ��ids := v_ҽ��ids || ',' || r_Odr(Ol).ҽ��id;
        End If;
        If Instr(',' || v_Nos || ',', ',' || r_Odr(Ol).No || ',') = 0 Then
          v_Nos := v_Nos || ',' || r_Odr(Ol).No;
        End If;
      
        v_Item := v_Item || ',{"send_no":' || r_Odr(Ol).���ͺ�;
        v_Item := v_Item || ',"advice_id":' || r_Odr(Ol).ҽ��id;
        v_Item := v_Item || ',"fee_no":"' || r_Odr(Ol).No || '"';
        v_Item := v_Item || ',"bill_prop":' || r_Odr(Ol).��¼����;
        v_Item := v_Item || ',"outpati_account":' || Nvl(r_Odr(Ol).������� || '', 'null');
        v_Item := v_Item || ',"pati_source":' || r_Odr(Ol).��Դ;
        v_Item := v_Item || '}';
      
        If Length(v_Item) > 30000 Then
          If c_Item Is Null Then
            c_Item := Substr(v_Item, 2);
          Else
            c_Item := c_Item || v_Item;
          End If;
          v_Item := Null;
        End If;
      End Loop;
    
      v_Item_a := v_Item_a || ',{"pati_id":' || n_Pati_Id;
      v_Item_a := v_Item_a || ',"pati_pageid":' || Nvl(n_Pati_Pageid || '', 'null');
      v_Item_a := v_Item_a || ',"rgst_id":' || Nvl(n_Rgst_Id || '', 'null');
      v_Item_a := v_Item_a || ',"rgst_no":"' || v_Rgs_No || '"';
      v_Item_a := v_Item_a || ',"order_ids":"' || Substr(v_ҽ��ids, 2) || '"';
      v_Item_a := v_Item_a || ',"fee_nos":"' || Substr(v_Nos, 2) || '"';
    
      If c_Item_a Is Null Then
        If c_Item Is Null Then
          v_Item_a := v_Item_a || ',"order_list":[' || Substr(v_Item, 2) || ']';
          v_Item_a := v_Item_a || '}';
          c_Item_a := Substr(v_Item_a, 2);
        Else
          c_Item   := c_Item || v_Item;
          c_Item_a := Substr(v_Item_a, 2) || ',"order_list":[' || c_Item || ']';
          c_Item_a := c_Item_a || '}';
        End If;
      Else
        If c_Item Is Null Then
          v_Item_a := v_Item_a || ',"order_list":[' || Substr(v_Item, 2) || ']';
          v_Item_a := v_Item_a || '}';
          c_Item_a := c_Item_a || v_Item_a;
        Else
          c_Item   := c_Item || v_Item;
          c_Item_a := c_Item_a || v_Item_a || ',"order_list":[' || c_Item || ']';
          c_Item_a := c_Item_a || '}';
        End If;
      End If;
    
      v_Item    := Null;
      c_Item    := Null;
      v_Item_a  := Null;
      v_ҽ��ids := Null;
      v_Nos     := Null;
    End If;
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","pati_bill_list":[' || c_Item_a || ']}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getaffirmerrordata;
/
Create Or Replace Procedure Zl_Cissvr_Getallgroupadviceids
(
  Json_In  Varchar,
  Json_Out Out Varchar
) Is
  ---------------------------------------------------------------------------
  --����ҽ��ID��ѯһ��ҽ��������ҽ����Ϣ
  --���      json
  --input     ����ҽ��ID��ѯҽ����Ϣ
  --  advice_id                     C 1  ���ҽ��ID���ö��ŷָ�
  --����      json
  --output
  --  code                          C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message                       C  1  Ӧ����Ϣ��
  --  advice_list[]                 ÿ��ҽ����Ϣ
  --    advice_id                   N    id
  --    advice_related_id           N    ���id
  --    clinic_type                 C    �������
  ---------------------------------------------------------------------------
  j_Json    Pljson;
  j_In      Pljson;
  v_ҽ��ids Varchar2(32767);
  v_Temp Varchar2(32767); 
Begin
  j_In      := Pljson(Json_In);
  j_Json    := j_In.Get_Pljson('input');
  v_ҽ��ids := j_Json.Get_String('advice_id');

  For r_ҽ�� In (Select /*+cardinality(j,10)*/
                a.Id, a.�Һŵ�, a.ҽ������, a.���id, a.�������
               From ����ҽ����¼ A, Table(f_Num2list(v_ҽ��ids)) J
               Where a.Id = j.Column_Value Or a.���id = j.Column_Value
               Union All
               Select /*+cardinality(j,10)*/
                a.Id, a.�Һŵ�, a.ҽ������, a.���id, a.�������
               From ����ҽ����¼ A, ����ҽ����¼ B, Table(f_Num2list(v_ҽ��ids)) J
               Where a.Id = b.���id And b.Id = j.Column_Value) Loop
  
    v_Temp := v_Temp || ',{"advice_id":' || r_ҽ��.Id;
    v_Temp := v_Temp || ',"advice_related_id":' || Nvl(r_ҽ��.���id || '', 'null');
    v_Temp := v_Temp || ',"clinic_Type":"' || r_ҽ��.������� || '"';
    v_Temp := v_Temp || '}';
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","advice_list":[' || Substr(v_Temp, 2) || ']}}';
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Cissvr_Getallgroupadviceids;
/
Create Or Replace Procedure Zl_Cissvr_Getbabydata
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:���ݲ���id����ҳId��ȡ����������
  --��Σ�Json_In:��ʽ
  --    input
  --      pati_id           N 1 ����id
  --      pati_pageid       N 1 ��ҳid

  --����: Json_Out,��ʽ����
  --  output
  --    code                  N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message               C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    baby_list[]
  --      baby_num            N 1 Ӥ�����
  --      baby_name           C 1 Ӥ������
  ---------------------------------------------------------------------------
  j_In     Pljson;
  j_Json   Pljson;
  v_List   Varchar2(32767);
  n_����id ������������¼.����id%Type;
  n_��ҳid ������������¼.��ҳid%Type;

Begin
  --�������
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');
  n_��ҳid := j_Json.Get_Number('pati_pageid');

  If Nvl(n_����id, 0) = 0 Then
    Json_Out := Zljsonout('δ���벡��id�����飡');
    Return;
  End If;

  If Nvl(n_��ҳid, 0) = 0 Then
    Json_Out := Zljsonout('δ������ҳid�����飡');
    Return;
  End If;

  For r_������ In (Select ���, Ӥ������ From ������������¼ Where ����id = n_����id And ��ҳid = n_��ҳid) Loop
  
    v_List := v_List || ',{"baby_num":' || r_������.���;
    v_List := v_List || ',"baby_name":"' || Zljsonstr(r_������.Ӥ������) || '"';
    v_List := v_List || '}';
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","baby_list":[' || Substr(v_List, 2) || ']}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getbabydata;
/
Create Or Replace Procedure Zl_Cissvr_Getcriticalinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  -------------------------------------------------------------
  --���ܣ���ȡ���˵�Σ��ֵ��Ϣ

  --���      json
  --input
  --    use_type                            N  1    �������� ��1-ͨ��id�ַ�����ѯ,2-ͨ���Һŵ���ѯ,3-ͨ������id����ҳid��ѯ,4-����Σ��ֵ�б�(����ʱ��Ϊ����  ����Ϊ��)
  --    cvalue_ids                          N  1    Σ��ֵids
  --    rgst_no                             C  0    �Һŵ�
  --    pati_id                             N  0    ����id
  --    pati_pageid                         N  0    ��ҳid

  --    cvalue_time_begin                   C  0   ����ʱ�䷶Χ��ʼʱ��
  --    cvalue_time_end                     C  0   ����ʱ�䷶Χ����ʱ��
  --    pati_type                           N  0   �������� 0-ȫ�� 1-���� 2-סԺ 3-����
  --    rpt_deptid                          N  0   �������ID Ϊ��ʱ  ������
  --    cnfm_deptid                         N  0   ȷ�Ͽ���ID Ϊ��ʱ  ������
  --    cvalue_rec_status                   N  0   ȷ��״̬ 0-ȫ�� 1-δȷ�� 2-ȷ��Ϊ��Σ��ֵ 3-ȷ��Ϊ��Σ��ֵ

  --����      json
  --output
  --    code                N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    cvalue_list[]         Σ��ֵ�б�֧�ֶ����[����]
  --       cvalue_id               N   1  Σ��ֵid
  --       advice_id               N   1  ҽ��id
  --       pat_name                C   1  ��������
  --       pat_sex                 C   1  �����Ա�
  --       pat_age                 C   1  ��������
  --       cvalue_rec_create_time  C   1  ����ʱ��
  --       cvalue_rec_status       N   1  Σ��ֵ״̬
  --       cvitem_result           N   1  �Ƿ�Σ��ֵ
  --       cvalue_rec_desc         C   1  Σ��ֵ˵��
  --       cvitem_source           C   1  ������Դ
  --       pati_id                 N   1  ����id
  --       pati_pageid             N   1  ��ҳid
  --       rgst_no                 C   1  �Һŵ�
  --       baby_num                N   1  Ӥ��
  --       lspcm_id                N   1  �걾id
  --       rpt_deptid   N   1  �������id
  --       rec_rptor               C   1  ������
  --       proc_note          C   1  �������
  --       cvalue_cnfmtime       C   1  ȷ��ʱ��
  --       cvalue_cnfmer         C   1  ȷ����
  --       cnfm_deptid         N   1  ȷ�Ͽ���id
  --       cvalue_dept           C   1  ȷ�Ͽ���
  -------------------------------------------------------------
  j_Json      Pljson;
  j_In        Pljson;
  n_Type      Number; --�������� ��1-ͨ��id��ѯ,2-ͨ���Һŵ���ѯ,3-ͨ������id����ҳid��ѯ
  v_Σ��ֵids Varchar2(4000);
  v_�Һŵ�    Varchar2(20);
  n_����id    Number(18);
  n_��ҳid    Number(18);

  v_List Varchar2(32767);

  d_Begin      Date;
  d_End        Date;
  n_��������   Number;
  n_�������id Number;
  n_ȷ�Ͽ���id Number;
  n_ȷ��״̬   Number;

Begin
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');
  n_Type := j_Json.Get_Number('use_type');

  If n_Type = 1 Then
    v_Σ��ֵids := j_Json.Get_String('cvalue_ids');
  
    If v_Σ��ֵids Is Null Then
      Json_Out := Zljsonout('δ����Σ��ֵid');
      Return;
    End If;
  
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","cvalue_list":[';
  
    For c_Σ��ֵ In (Select a.Id, a.������Դ, a.����id, a.��ҳid, a.�Һŵ�, a.Ӥ��, a.����, a.�Ա�, a.����, a.ҽ��id, a.�걾id, a.Σ��ֵ����,
                         To_Char(a.����ʱ��, 'YYYY-MM-DD HH24:MI') As ����ʱ��, a.�������id, a.������, a.�������,
                         To_Char(a.ȷ��ʱ��, 'YYYY-MM-DD HH24:MI') As ȷ��ʱ��, a.ȷ����, a.ȷ�Ͽ���id, a.״̬, a.�Ƿ�Σ��ֵ, c.���� As ȷ�Ͽ���
                  From ����Σ��ֵ��¼ A,
                       (Select /*+cardinality(b,10)*/
                          Column_Value
                         From Table(f_Str2list(v_Σ��ֵids))) B, ���ű� C
                  Where a.Id = b.Column_Value And a.ȷ�Ͽ���id = c.Id(+)) Loop
      Zljsonputvalue(v_List, 'cvalue_id', c_Σ��ֵ.Id, 1, 1);
      Zljsonputvalue(v_List, 'advice_id', c_Σ��ֵ.ҽ��id, 1);
      Zljsonputvalue(v_List, 'pat_name', c_Σ��ֵ.����, 0);
      Zljsonputvalue(v_List, 'pat_sex', c_Σ��ֵ.�Ա�, 0);
      Zljsonputvalue(v_List, 'pat_age', c_Σ��ֵ.����, 0);
      Zljsonputvalue(v_List, 'cvalue_rec_create_time', c_Σ��ֵ.����ʱ��, 0);
      Zljsonputvalue(v_List, 'cvalue_rec_status', c_Σ��ֵ.״̬, 1);
      Zljsonputvalue(v_List, 'cvitem_result', c_Σ��ֵ.�Ƿ�Σ��ֵ, 1);
      Zljsonputvalue(v_List, 'cvalue_rec_desc', c_Σ��ֵ.Σ��ֵ����, 0);
      Zljsonputvalue(v_List, 'cvitem_source', c_Σ��ֵ.������Դ, 0);
      Zljsonputvalue(v_List, 'pati_id', c_Σ��ֵ.����id, 1);
      Zljsonputvalue(v_List, 'pati_pageid', c_Σ��ֵ.��ҳid, 1);
      Zljsonputvalue(v_List, 'rgst_no', c_Σ��ֵ.�Һŵ�, 0);
      Zljsonputvalue(v_List, 'baby_num', c_Σ��ֵ.Ӥ��, 1);
      Zljsonputvalue(v_List, 'lspcm_id', c_Σ��ֵ.�걾id, 1);
      Zljsonputvalue(v_List, 'rpt_deptid', c_Σ��ֵ.�������id, 1);
      Zljsonputvalue(v_List, 'rec_rptor', c_Σ��ֵ.������, 0);
      Zljsonputvalue(v_List, 'proc_note', c_Σ��ֵ.�������, 0);
      Zljsonputvalue(v_List, 'cvalue_cnfmtime', c_Σ��ֵ.ȷ��ʱ��, 0);
      Zljsonputvalue(v_List, 'cvalue_cnfmer', c_Σ��ֵ.ȷ����, 0);
      Zljsonputvalue(v_List, 'cnfm_deptid', c_Σ��ֵ.ȷ�Ͽ���id, 1);
      Zljsonputvalue(v_List, 'cvalue_dept', c_Σ��ֵ.ȷ�Ͽ���, 0, 2);
      If Length(v_List) > 20000 Then
        Json_Out := Json_Out || v_List;
        v_List   := ',';
      End If;
    End Loop;
  Elsif n_Type = 2 Then
    v_�Һŵ� := j_Json.Get_String('rgst_no');
  
    If v_�Һŵ� Is Null Then
      Json_Out := Zljsonout('δ����Һŵ�');
      Return;
    End If;
  
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","cvalue_list":[';
  
    For c_Σ��ֵ In (Select a.Id, a.������Դ, a.����id, a.��ҳid, a.�Һŵ�, a.Ӥ��, a.����, a.�Ա�, a.����, a.ҽ��id, a.�걾id, a.Σ��ֵ����,
                         To_Char(a.����ʱ��, 'YYYY-MM-DD HH24:MI') As ����ʱ��, a.�������id, a.������, a.�������,
                         To_Char(a.ȷ��ʱ��, 'YYYY-MM-DD HH24:MI') As ȷ��ʱ��, a.ȷ����, a.ȷ�Ͽ���id, a.״̬, a.�Ƿ�Σ��ֵ
                  From ����Σ��ֵ��¼ A
                  Where �Һŵ� = v_�Һŵ�) Loop
      Zljsonputvalue(v_List, 'cvalue_id', c_Σ��ֵ.Id, 1, 1);
      Zljsonputvalue(v_List, 'advice_id', c_Σ��ֵ.ҽ��id, 1);
      Zljsonputvalue(v_List, 'pat_name', c_Σ��ֵ.����, 0);
      Zljsonputvalue(v_List, 'pat_sex', c_Σ��ֵ.�Ա�, 0);
      Zljsonputvalue(v_List, 'pat_age', c_Σ��ֵ.����, 0);
      Zljsonputvalue(v_List, 'cvalue_rec_create_time', c_Σ��ֵ.����ʱ��, 0);
      Zljsonputvalue(v_List, 'cvalue_rec_status', c_Σ��ֵ.״̬, 1);
      Zljsonputvalue(v_List, 'cvitem_result', c_Σ��ֵ.�Ƿ�Σ��ֵ, 1);
      Zljsonputvalue(v_List, 'cvalue_rec_desc', c_Σ��ֵ.Σ��ֵ����, 0, 2);
      If Length(v_List) > 20000 Then
        Json_Out := Json_Out || v_List;
        v_List   := ',';
      End If;
    End Loop;
  
  Elsif n_Type = 3 Then
    n_����id := j_Json.Get_Number('pati_id');
    n_��ҳid := j_Json.Get_Number('pati_pageid');
  
    If n_����id Is Null Then
      Json_Out := Zljsonout('δ���벡��id');
      Return;
    End If;
  
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","cvalue_list":[';
  
    For c_Σ��ֵ In (Select a.Id, a.������Դ, a.����id, a.��ҳid, a.�Һŵ�, a.Ӥ��, a.����, a.�Ա�, a.����, a.ҽ��id, a.�걾id, a.Σ��ֵ����,
                         To_Char(a.����ʱ��, 'YYYY-MM-DD HH24:MI') As ����ʱ��, a.�������id, a.������, a.�������,
                         To_Char(a.ȷ��ʱ��, 'YYYY-MM-DD HH24:MI') As ȷ��ʱ��, a.ȷ����, a.ȷ�Ͽ���id, a.״̬, a.�Ƿ�Σ��ֵ
                  From ����Σ��ֵ��¼ A
                  Where ����id = n_����id And ��ҳid = n_��ҳid) Loop
      Zljsonputvalue(v_List, 'cvalue_id', c_Σ��ֵ.Id, 1, 1);
      Zljsonputvalue(v_List, 'advice_id', c_Σ��ֵ.ҽ��id, 1);
      Zljsonputvalue(v_List, 'pat_name', c_Σ��ֵ.����, 0);
      Zljsonputvalue(v_List, 'pat_sex', c_Σ��ֵ.�Ա�, 0);
      Zljsonputvalue(v_List, 'pat_age', c_Σ��ֵ.����, 0);
      Zljsonputvalue(v_List, 'cvalue_rec_create_time', c_Σ��ֵ.����ʱ��, 0);
      Zljsonputvalue(v_List, 'cvalue_rec_status', c_Σ��ֵ.״̬, 1);
      Zljsonputvalue(v_List, 'cvitem_result', c_Σ��ֵ.�Ƿ�Σ��ֵ, 1);
      Zljsonputvalue(v_List, 'cvalue_rec_desc', c_Σ��ֵ.Σ��ֵ����, 0, 2);
      If Length(v_List) > 20000 Then
        Json_Out := Json_Out || v_List;
        v_List   := ',';
      End If;
    End Loop;
  Elsif n_Type = 4 Then
    d_Begin      := To_Date(j_Json.Get_String('cvalue_time_begin'), 'YYYY-MM-DD HH24:MI:SS');
    d_End        := To_Date(j_Json.Get_String('cvalue_time_end'), 'YYYY-MM-DD HH24:MI:SS');
    n_��������   := Nvl(j_Json.Get_Number('pati_type'), 0);
    n_�������id := Nvl(j_Json.Get_Number('rpt_deptid'), 0);
    n_ȷ�Ͽ���id := Nvl(j_Json.Get_Number('cnfm_deptid'), 0);
    n_ȷ��״̬   := Nvl(j_Json.Get_Number('cvalue_rec_status'), 0);
  
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","cvalue_list":[';
  
    For c_Σ��ֵ In (Select a.Id, a.������Դ, a.����id, a.��ҳid, a.�Һŵ�, a.Ӥ��, a.����, a.�Ա�, a.����, a.ҽ��id, a.�걾id, a.Σ��ֵ����,
                         To_Char(a.����ʱ��, 'YYYY-MM-DD HH24:MI') As ����ʱ��, a.�������id, a.������, a.�������,
                         To_Char(a.ȷ��ʱ��, 'YYYY-MM-DD HH24:MI') As ȷ��ʱ��, a.ȷ����, a.ȷ�Ͽ���id, a.״̬, a.�Ƿ�Σ��ֵ, c.���� As ȷ�Ͽ���
                  From ����Σ��ֵ��¼ A, ���ű� C
                  Where a.����ʱ�� Between d_Begin And d_End And
                        ((Nvl(n_�������id, 0) = a.�������id Or Nvl(n_�������id, 0) = 0) And
                        (Nvl(n_��������, 0) = 0 Or (Nvl(n_��������, 0) = 1 And a.�Һŵ� Is Not Null) Or
                        (Nvl(n_��������, 0) = 2 And Nvl(a.��ҳid, 0) > 0) Or
                        (Nvl(n_��������, 0) = 3 And Nvl(a.��ҳid, 0) = 0 And a.�Һŵ� Is Null)) And
                        (Nvl(n_ȷ�Ͽ���id, 0) = a.ȷ�Ͽ���id Or Nvl(n_ȷ�Ͽ���id, 0) = 0) And
                        (Nvl(n_ȷ��״̬, 0) = 0 Or (Nvl(n_ȷ��״̬, 0) = 1 And a.״̬ = 1) Or
                        (Nvl(n_ȷ��״̬, 0) = 2 And a.״̬ = 2 And Nvl(a.�Ƿ�Σ��ֵ, 0) = 0) Or
                        (Nvl(n_ȷ��״̬, 0) = 3 And a.״̬ = 2 And Nvl(a.�Ƿ�Σ��ֵ, 0) = 1))) And a.ȷ�Ͽ���id = c.Id(+)
                  Order By a.����ʱ�� Desc) Loop
    
      Zljsonputvalue(v_List, 'cvalue_id', c_Σ��ֵ.Id, 1, 1);
      Zljsonputvalue(v_List, 'advice_id', c_Σ��ֵ.ҽ��id, 1);
      Zljsonputvalue(v_List, 'pat_name', c_Σ��ֵ.����, 0);
      Zljsonputvalue(v_List, 'pat_sex', c_Σ��ֵ.�Ա�, 0);
      Zljsonputvalue(v_List, 'pat_age', c_Σ��ֵ.����, 0);
      Zljsonputvalue(v_List, 'cvalue_rec_create_time', c_Σ��ֵ.����ʱ��, 0);
      Zljsonputvalue(v_List, 'cvalue_rec_status', c_Σ��ֵ.״̬, 1);
      Zljsonputvalue(v_List, 'cvitem_result', c_Σ��ֵ.�Ƿ�Σ��ֵ, 1);
      Zljsonputvalue(v_List, 'cvalue_rec_desc', c_Σ��ֵ.Σ��ֵ����, 0);
      Zljsonputvalue(v_List, 'cvitem_source', c_Σ��ֵ.������Դ, 0);
      Zljsonputvalue(v_List, 'pati_id', c_Σ��ֵ.����id, 1);
      Zljsonputvalue(v_List, 'pati_pageid', c_Σ��ֵ.��ҳid, 1);
      Zljsonputvalue(v_List, 'rgst_no', c_Σ��ֵ.�Һŵ�, 0);
      Zljsonputvalue(v_List, 'baby_num', c_Σ��ֵ.Ӥ��, 1);
      Zljsonputvalue(v_List, 'lspcm_id', c_Σ��ֵ.�걾id, 1);
      Zljsonputvalue(v_List, 'rpt_deptid', c_Σ��ֵ.�������id, 1);
      Zljsonputvalue(v_List, 'rec_rptor', c_Σ��ֵ.������, 0);
      Zljsonputvalue(v_List, 'proc_note', c_Σ��ֵ.�������, 0);
      Zljsonputvalue(v_List, 'cvalue_cnfmtime', c_Σ��ֵ.ȷ��ʱ��, 0);
      Zljsonputvalue(v_List, 'cvalue_cnfmer', c_Σ��ֵ.ȷ����, 0);
      Zljsonputvalue(v_List, 'cnfm_deptid', c_Σ��ֵ.ȷ�Ͽ���id, 1);
      Zljsonputvalue(v_List, 'cvalue_dept', c_Σ��ֵ.ȷ�Ͽ���, 0, 2);
      If Length(v_List) > 20000 Then
        Json_Out := Json_Out || v_List;
        v_List   := ',';
      End If;
    End Loop;
  End If;

  If v_List <> ',' Then
    Json_Out := Json_Out || v_List || ']}}';
  Else
    Json_Out := Json_Out || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getcriticalinfo;
/
Create Or Replace Procedure Zl_Cissvr_Getdiagfitsituation
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:��ȡ��Ϸ������
  --��Σ�Json_In:��ʽ
  -- input
  --   pati_id              N 1 ����id
  --   pati_pageid          N 1 ��ҳid
  --����: Json_Out,��ʽ����
  --  output
  --    code               C  1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message            C  1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    fit_list[]      [����]
  --      fit_type         N  1 1.�������Ժ��2.��Ժ���Ժ��3.�����벡��4.�ٴ��벡��5.�ٴ���ʬ�졢6.��ǰ������11.��ҽ�������Ժ��12.��ҽ��Ժ���Ժ��13.��ҽ��֤��14.��ҽ�η���15.��ҽ��ҩ
  --      diag_cnst        N  1 ��Ϸ������: 0.δ����1.���ϡ�2.�����ϡ�3.���϶���������ҽ׼ȷ�ȣ�1.׼ȷ��2.����׼ȷ��3.�ش�ȱ�ݡ�4.����

  ---------------------------------------------------------------------------
  j_In     Pljson;
  j_Json   Pljson;
  n_����id ��Ϸ������.����id%Type;
  n_��ҳid ��Ϸ������.��ҳid%Type;
  v_List   Varchar2(32767);

Begin
  --�������
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');
  n_��ҳid := j_Json.Get_Number('pati_pageid');

  If Nvl(n_����id, 0) = 0 Or Nvl(n_��ҳid, 0) = 0 Then
    Json_Out := Zljsonout('���봫�벡��id����ҳid��');
    Return;
  End If;

  For r_��� In (Select ��������, Nvl(�������, 0) As �������
               From ��Ϸ������
               Where ����id = n_����id And ��ҳid = n_��ҳid) Loop
    v_List := v_List || ',{"fit_type":' || Nvl(r_���.�������� || '', 'null');
    v_List := v_List || ',"diag_cnst":' || r_���.�������;
    v_List := v_List || '}';
  
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","fit_list":[' || Substr(v_List, 2) || ']}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getdiagfitsituation;
/
Create Or Replace Procedure Zl_Cissvr_Getdiaginfo
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:��ȡ���������Ϣ
  --��Σ�Json_In:��ʽ
  -- input
  --   advice_ids           C 1  ҽ��ids,ҽ��idƴ��
  --   query_type           N 1 ��ѯ��ʽ1-��ָ��������ѯ,2-��������id,��ҳid��ѯ���
  --   pati_info            C 0  ����id��������Ϣ
  --     pati_id            N 1 ����id
  --     pati_pageid        N 1 ��ҳid
  --     diag_types         C 1  �������:0-��������,1-��ҽ�������;2-��ҽ��Ժ���;3-��ҽ��Ժ���;5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���,8-��ǰ���;9-�������;10-����֢;11-��ҽ�������;12-��ҽ��Ժ���;13-��ҽ��Ժ���;21-��ԭѧ���;22-Ӱ��ѧ���
  --                            ����Ϊ���������ͣ��ö��ŷ���,��:2,12
  --     rec_source         N 1 ��¼��Դ:1-������2-��Ժ�Ǽǣ�3-��ҳ����;4-����;NULL-��������
  --     diag_num           N 1 ��ϴ���:NULL��ʾ��������
  --     code_type          C 1  �������:ICD-11�ı���������Ϊ'E',��ʱ��ʾ��ȡICD-10��
  --     input_num          C 1  ¼�����:������ICD-11����¼�����ϵ�¼�����
  --     rec_sources        C 1 ��¼��Դƴ��

  --����: Json_Out,��ʽ����
  --  output
  --    code                N  1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message             C  1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    diag_list     [����]
  --      diag_type         N 1 �������:1-��ҽ�������;2-��ҽ��Ժ���;3-��ҽ��Ժ���;5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���,8-��ǰ���;9-�������;10-����֢;11-��ҽ�������;
  --                                     12-��ҽ��Ժ���;13-��ҽ��Ժ���;21-��ԭѧ���;22-Ӱ��ѧ���
  --      diag_num          N 1 ������
  --      code_num          N 1 �������
  --      dz_id             N 1 ����ID
  --      dz_code           C 1 ��������
  --      diag_note         C 1 �������
  --      recoder           C 1 ��¼��
  --      rec_time          C 1 ��¼ʱ��:yyyy-mm-dd hh24:mi:ss
  --      adtd_rsn          C 1 ��Ժ���:��������ת��δ��������������
  --      diag_id           N 1 ���id
  --      diag_rec_id       N 1 ��ϼ�¼ID:������ϼ�¼.ID
  --      diag_doubt        N 1 �Ƿ�����
  --      advice_id         N   ҽ��id(����ҽ��ids��ѯʱ�ŷ���)
  --      advice_main_id    N   ��ҽ��id(����ҽ��ids��ѯʱ�ŷ���)
  --      advice_related_id N   ���id(����ҽ��ids��ѯʱ�ŷ���)
  --      rec_source        N   ��¼��Դ:1-������2-��Ժ�Ǽǣ�3-��ҳ����;4-����;NULL-��������
  ---------------------------------------------------------------------------
  j_In       Pljson;
  j_Json     Pljson;
  j_Patiinfo Pljson;
  c_ҽ��ids  Clob;
  v_ҽ��ids  Varchar2(32767);
  n_����id   ������ϼ�¼.����id%Type;
  n_��ҳid   ������ϼ�¼.��ҳid%Type;
  n_��¼��Դ ������ϼ�¼.��¼��Դ%Type;
  v_������� Varchar2(3000);
  n_��ϴ��� ������ϼ�¼.��ϴ���%Type;
  v_������� ������ϼ�¼.�������%Type;
  v_¼����� ������ϼ�¼.¼�����%Type;
  I          Number := 0;
  n_��ѯ���� Number := 0;
  v_�����Դ Varchar2(100);
  v_List     Varchar2(32767);

  --��װʧ��ʱ���ص�����
  Function Get_Err_Message(Message_In Varchar2) Return Varchar2 Is
  Begin
    Return '{"output":{"code":0,"message":"' || Zljsonstr(Message_In) || '"}}';
  End Get_Err_Message;
Begin
  --�������
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","diag_list":[';

  If j_Json.Get_Number('query_type') = 2 Then
    j_Patiinfo := j_Json.Get_Pljson('pati_info');
    If Not j_Patiinfo Is Null Then
      n_����id := j_Patiinfo.Get_Number('pati_id');
      n_��ҳid := j_Patiinfo.Get_Number('pati_pageid');
    End If;
    For r_��� In (Select �������, ��ϴ���, �������, ��Ժ���, ��¼��Դ
                 From ������ϼ�¼
                 Where ��¼��Դ In (2, 3) And Nvl(�������, 1) = 1 And ����id = n_����id And ��ҳid = n_��ҳid And
                       Nvl(¼�����, '01') = '01'
                 Order By ��¼��Դ Desc) Loop
    
      Zljsonputvalue(v_List, 'diag_type', r_���.�������, 1, 1);
      Zljsonputvalue(v_List, 'diag_num', r_���.��ϴ���, 1);
      Zljsonputvalue(v_List, 'rec_source', r_���.��¼��Դ, 1);
      Zljsonputvalue(v_List, 'diag_note', r_���.�������);
      Zljsonputvalue(v_List, 'adtd_rsn', r_���.��Ժ���, 0, 2);
      If Length(v_List) > 20000 Then
        Json_Out := Json_Out || v_List;
        v_List   := ',';
      End If;
    End Loop;
    If v_List <> ',' Then
      Json_Out := Json_Out || v_List || ']}}';
    Else
      Json_Out := Json_Out || ']}}';
    End If;
    Return;
  End If;
  Begin
    c_ҽ��ids := j_Json.Get_Clob('advice_ids');
  Exception
    When Others Then
      c_ҽ��ids := Null;
  End;
  If c_ҽ��ids Is Null Then
    j_Patiinfo := Pljson();

    j_Patiinfo := j_Json.Get_Pljson('pati_info');
    If Not j_Patiinfo Is Null Then
      n_����id   := j_Patiinfo.Get_Number('pati_id');
      n_��ҳid   := j_Patiinfo.Get_Number('pati_pageid');
      v_������� := j_Patiinfo.Get_String('diag_types');
      n_��¼��Դ := j_Patiinfo.Get_Number('rec_source');
      n_��ϴ��� := j_Patiinfo.Get_Number('diag_num');
      v_������� := j_Patiinfo.Get_String('code_type');
      v_¼����� := j_Patiinfo.Get_String('input_num');
      v_�����Դ := j_Patiinfo.Get_String('rec_sources');
      If Nvl(n_����id, 0) = 0 Or Nvl(n_��ҳid, 0) = 0 Then
        Json_Out := Get_Err_Message('δ���벡��id����ҳid�����飡');
        Return;
      End If;
    End If;
  End If;

  If c_ҽ��ids Is Null Then
    If n_��ѯ���� = 0 Then
      For r_��� In (Select a.�������, a.��ϴ���, a.�������, b.Id As ����id, b.���� As ��������, a.�������, a.��¼��,
                          To_Char(Nvl(a.��¼����, Sysdate), 'yyyy-mm-dd hh24:mi:ss') As ��¼����, a.��Ժ���, a.Id As ��ϼ�¼id, a.���id,
                          a.�Ƿ�����
                   From ������ϼ�¼ A, ��������Ŀ¼ B
                   Where a.����id = b.Id And (a.��¼��Դ = n_��¼��Դ Or n_��¼��Դ Is Null) And (a.��ϴ��� = n_��ϴ��� Or n_��ϴ��� Is Null) And
                         (a.������� In (Select Column_Value From Table(f_Str2list(v_�������)) C) Or Nvl(v_�������, 0) = 0) And
                         a.����id = n_����id And a.��ҳid = n_��ҳid And Nvl(a.¼�����, '01') = Nvl(v_¼�����, '01') And
                         Nvl(a.�������, 'E') = Nvl(v_�������, 'E')) Loop
        --      diag_type         N 1 �������:1-��ҽ�������;2-��ҽ��Ժ���;3-��ҽ��Ժ���;5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���,8-��ǰ���;9-�������;10-����֢;11-��ҽ�������;
        --                                     12-��ҽ��Ժ���;13-��ҽ��Ժ���;21-��ԭѧ���;22-Ӱ��ѧ���
        --      diag_num          N 1 ������
        --      code_num          N 1 �������
        --      dz_id             N 1 ����ID
        --      dz_code           C 1 ��������
        --      diag_note         C 1 �������
        --      recoder           C 1 ��¼��
        --      rec_time          C 1 ��¼ʱ��:yyyy-mm-dd hh24:mi:ss
        --      adtd_rsn          C 1 ��Ժ���:��������ת��δ��������������
        --      diag_id           N 1 ���id
        --      diag_rec_id       N 1 ��ϼ�¼ID:������ϼ�¼.ID
        --      diag_doubt        N 1 �Ƿ�����
        Zljsonputvalue(v_List, 'diag_type', r_���.�������, 1, 1);
        Zljsonputvalue(v_List, 'diag_num', r_���.��ϴ���, 1);
        Zljsonputvalue(v_List, 'code_num', r_���.�������);
        Zljsonputvalue(v_List, 'dz_id', r_���.����id, 1);
        Zljsonputvalue(v_List, 'dz_code', r_���.��������);
        Zljsonputvalue(v_List, 'diag_note', r_���.�������);
        Zljsonputvalue(v_List, 'recoder', r_���.��¼��);
        Zljsonputvalue(v_List, 'rec_time', r_���.��¼����);
        Zljsonputvalue(v_List, 'adtd_rsn', r_���.��Ժ���);
        Zljsonputvalue(v_List, 'diag_id', r_���.���id, 1);
        Zljsonputvalue(v_List, 'diag_rec_id', r_���.��ϼ�¼id, 1);
        Zljsonputvalue(v_List, 'diag_doubt', r_���.�Ƿ�����, 1, 2);
      
        If Length(v_List) > 20000 Then
          Json_Out := Json_Out || v_List;
          v_List   := ',';
        End If;
      End Loop;
      If v_List <> ',' Then
        Json_Out := Json_Out || v_List || ']}}';
      Else
        Json_Out := Json_Out || ']}}';
      End If;
    Else
      For r_��� In (Select �������, ����id, �������, ��Ժ���
                   From ������ϼ�¼
                   Where ((��ϴ��� = 1 And n_��ϴ��� = 0) Or (��ϴ��� > 1 And n_��ϴ��� = 1) Or Nvl(n_��ϴ���, 0) = 0) And
                         ��¼��Դ In (v_�����Դ) And Nvl(�������, 1) = 1 And ����id = n_����id And ��ҳid = n_��ҳid And
                         (������� In (v_�������) Or Nvl(v_�������, '-') = '-') And Nvl(¼�����, '01') = '01'
                   Order By ��¼��Դ Desc) Loop
        Zljsonputvalue(v_List, 'diag_type', r_���.�������, 1, 1);
        Zljsonputvalue(v_List, 'dz_id', r_���.����id, 1);
        Zljsonputvalue(v_List, 'diag_note', r_���.�������);
        Zljsonputvalue(v_List, 'adtd_rsn', r_���.��Ժ���, 0, 2);
      
        If Length(v_List) > 20000 Then
          Json_Out := Json_Out || v_List;
          v_List   := ',';
        End If;
      End Loop;
      If v_List <> ',' Then
        Json_Out := Json_Out || v_List || ']}}';
      Else
        Json_Out := Json_Out || ']}}';
      End If;
    End If;
  
  Else
    While c_ҽ��ids Is Not Null Loop
      If Length(c_ҽ��ids) <= 4000 Then
        v_ҽ��ids := c_ҽ��ids;
        c_ҽ��ids := Null;
      Else
        v_ҽ��ids := Substr(c_ҽ��ids, 1, Instr(c_ҽ��ids, ',', 3980) - 1);
        c_ҽ��ids := Substr(c_ҽ��ids, Instr(c_ҽ��ids, ',', 3980) + 1);
      End If;
      I := I + 1;
    
      For r_��� In (Select a.��id, a.ҽ��id, a.���id, c.�������, c.��ϴ���, c.�������, c.�������, c.��¼��,
                          To_Char(Nvl(c.��¼����, Sysdate), 'yyyy-mm-dd hh24:mi:ss') As ��¼����, c.��Ժ���, c.Id As ��ϼ�¼id, c.���id,
                          c.�Ƿ�����
                   From (Select Nvl(a.���id, a.Id) As ��id, a.Id As ҽ��id, a.���id
                          From ����ҽ����¼ A
                          Where a.Id In (Select /*+cardinality(x,10)*/
                                          x.Column_Value As ҽ��id
                                         From Table(Cast(f_Num2list(v_ҽ��ids) As Zltools.t_Numlist)) X)) A, �������ҽ�� B,
                        ������ϼ�¼ C
                   Where a.��id = b.ҽ��id And b.���id = c.Id
                   Order By c.Id) Loop
      
        --      diag_type         N 1 �������:1-��ҽ�������;2-��ҽ��Ժ���;3-��ҽ��Ժ���;5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���,8-��ǰ���;9-�������;10-����֢;11-��ҽ�������;
        --                                     12-��ҽ��Ժ���;13-��ҽ��Ժ���;21-��ԭѧ���;22-Ӱ��ѧ���
        --      diag_num          N 1 ������
        --      code_num          N 1 �������
        --      diag_note         C 1 �������
        --      recoder           C 1 ��¼��
        --      rec_time          C 1 ��¼ʱ��:yyyy-mm-dd hh24:mi:ss
        --      adtd_rsn          C 1 ��Ժ���:��������ת��δ��������������
        --      diag_id           N 1 ���id
        --      diag_rec_id       N 1 ��ϼ�¼ID:������ϼ�¼.ID
        --      diag_doubt        N 1 �Ƿ�����
        --      advice_id         N   ҽ��id(����ҽ��ids��ѯʱ�ŷ���)
        --      advice_main_id    N   ��ҽ��id(����ҽ��ids��ѯʱ�ŷ���)
        --      advice_related_id N   ���id(����ҽ��ids��ѯʱ�ŷ���)
        Zljsonputvalue(v_List, 'advice_id', r_���.ҽ��id, 1, 1);
        Zljsonputvalue(v_List, 'advice_main_id', r_���.��id, 1);
        Zljsonputvalue(v_List, 'advice_related_id', r_���.���id, 1);
        Zljsonputvalue(v_List, 'diag_type', r_���.�������, 1);
        Zljsonputvalue(v_List, 'diag_num', r_���.��ϴ���, 1);
        Zljsonputvalue(v_List, 'code_num', r_���.�������);
        Zljsonputvalue(v_List, 'diag_note', r_���.�������);
        Zljsonputvalue(v_List, 'recoder', r_���.��¼��);
        Zljsonputvalue(v_List, 'rec_time', r_���.��¼����);
        Zljsonputvalue(v_List, 'adtd_rsn', r_���.��Ժ���);
        Zljsonputvalue(v_List, 'diag_id', r_���.���id, 1);
        Zljsonputvalue(v_List, 'diag_rec_id', r_���.��ϼ�¼id, 1);
        Zljsonputvalue(v_List, 'diag_doubt', r_���.�Ƿ�����, 1, 2);
      
        If Length(v_List) > 20000 Then
          Json_Out := Json_Out || v_List;
          v_List   := ',';
        End If;
      
      End Loop;
    End Loop;
  
    If v_List <> ',' Then
      Json_Out := Json_Out || v_List || ']}}';
    Else
      Json_Out := Json_Out || ']}}';
    End If;
  
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getdiaginfo;
/
Create Or Replace Procedure Zl_Cissvr_Getdrugerrdata
(
  Json_In  In Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ��ٴ�ҽ���������ɴ���,����,����,����ͬ��
  --��Σ�Json_In:��ʽ
  --  input
  --      pati_ids                        C 1 ����ids����ƴ��  
  --����: Json_Out,��ʽ����
  --   output
  --     code                             N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --     message                          C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --     pati_bill_list[]                 ����ҽ�����õ�����Ϣ
  --         pati_id                      N 1 ����id
  --         pati_pageid                  N 0 ��ҳid��סԺ���˴��룬���ﴫ0
  --         rgst_id                      N 0 �Һ�id�����ﲡ�˴��룬סԺ���˴���
  --         rgst_no                      C 0 �Һŵ���
  --         send_no                      N 1 ���ͺ�
  --         operator_name                C 1 ������(����Ա����)
  --         operator_time                C 1 ����ʱ��
  --         pati_type                    C 0 �������ͣ���ΪסԺ����ʱ���Ի�ȡ����������ﲡ�����ⲿ������ȡ
  --         diag_list[]                  �ٴ������Ϣ��
  --             diag_rec_id              N 1 ��ϼ�¼id
  --             diag_type                N 1 ������� 1-��ҽ�������;2-��ҽ��Ժ���;3-��ҽ��Ժ���;5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���,8-��ǰ���;9-�������;10-����֢;11-��ҽ�������;12-��ҽ��Ժ���;13-��ҽ��Ժ���;21-��ԭѧ���;22-Ӱ��ѧ���
  --             diag_name                C 1 �������
  --         bill_list[]ҽ����Ϣ�б�
  --             advice_id                N 1 ҽ��id
  --             group_sno                N 0 ������� (�����洢)��1��2��3
  --             effectivetime            N  0 ҽ����Ч
  --             drug_method_id           N 1 ��ҩ;��id(������):������ĿID: 204,
  --             drug_method_name         C 1 ��ҩ;������: ��������,
  --             drug_method_class_code   C 1 ��ҩ;������(������):ִ�з�����: 001,
  --             drug_freq_id             N 1 ��ҩƵ��id(������):����Ƶ�ʱ���: 1,
  --             drug_freq_name           C 1 ��ҩƵ������(������):: ÿ�����,
  --             emergency_tag            N 1 ҽ����¼�еĽ�����־(0-��ͨ;1-����;2-��¼(��������Ч))
  --             fee_mode                 N 1 �Ƽ����ԣ�0-�����Ƽۣ�1-���Ƽۣ�2-�ֹ��Ƽ�
  --             use_mode                 N 1 ȡҩ���ԣ�0-������ʽ��1-��Ժ��ҩ��2-��ȡҩ
  --             frequency                N 0 Ƶ��: 2,
  --             single                   N 0 ����: 1,
  --             usage                    C 0 �÷�: ��������,
  --             rcpdtl_st_result         N 0 Ƥ�Խ��(������)1-���ԣ�2-���ԣ�3-���ԣ�4-������ҩ ��������ʱ��ȷ��������Ƥ�Խ����ZLHISĿǰ֧�ֲ�ȫ: ,
  --             rcpdtL_excs_desc         C 0 ����˵��(������): ,
  --             rcpdtL_drask             C 0 ʹ������(������): ,
  --             memo                     C 0 ժҪ: ҽ������,
  --             diag_name                C 0 ������ƣ�������)�����ﴫ�룬�������:
  --             take_no                  C 0 ��ҩ��
  --             advice_purpose           C 0 ��ҩĿ��
  --             fee_source               N 0 ������Դ��1-���2-סԺ
  --             fee_billtype             N 0 ���õ������ͣ�1-�շѴ�����2-���ʵ�����
  --             fee_no                   C 0 ���õ��ݺ�
  --         pivas_list[] ������Ϣ�б�ֻ��һ��Ԫ��
  --             pati_id                  N 1 ����id
  --             page_id                  N 1 ��ҳid
  --             pati_name                C 1 ����
  --             pati_sex                 C 1 �Ա�
  --             pati_age                 C 1 ����
  --             inpatient_num            C 1 סԺ��
  --             pati_bed                 C 1 ����
  --             pati_wardarea_id         N 1 ���˲���id
  --             pati_deptid              N 1 ���˿���id
  --             advice_list[]��ҽ��������
  --                 pivas_deptid         N 1 ��������id
  --                 advice_id            N 1 ��ҽ��ID(��ҩ;��)
  --                 effective_time       N 1 ҽ����Ч
  --                 drug_method_id       N 1 ��ҩ;��id
  --                 is_tpn               N 1 �Ƿ�tpn
  --                 advice_frequency     C 1 ִ��Ƶ��
  --                 advice_drug_list[]��ҩ;����Ӧ��ҩ��������
  --                     advice_id        N 1 ҩ��id
  --                     advice_rcpno     C 1 ҩ�����Ͳ����ķ���no
  --                 advice_exetime_list[]ҽ��ִ��ʱ�䣬��ҩ;��ҽ����ִ��ʱ�䣬��ʱ�ṩ��ҽ�����з��͵�ʱ�䣬�������η��͵�ִ��ʱ�䡣�����ͺŵ�����֯��������
  --                     advice_send_no   N 1 ��ҩ;��ҽ���ķ��ͺ�
  --                     advice_require_time  C 1 Ҫ��ʱ��: 2019-11-30 23:00:00
  ------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  b_Pivasout Clob;
  v_����Ա   Varchar2(300);
  v_����ʱ�� Varchar2(40);

  c_Outtmp   Clob;
  v_Jtmp     Varchar2(32767);
  v_Tmp      Varchar2(32767);
  n_Cnt      Number;
  v_Err      Varchar2(32767);
  v_Pҽ��ids Varchar2(32767);
  n_��id     Number;
  j_Json     Pljson;
  j_Tmp      Pljson;
  n_�к�     Number;
  n_Send_No  Number;
  v_Vals     Clob;
  l_Vals     t_Strlist;
  Err_Item Exception;

  n_Rgst_Id                Number; --N   1�Һŵ�id��������)
  v_Take_No                Varchar2(32767); --C 0 ��ҩ�� ��ҩ�ţ�δ��ҩƷ��¼.��ҩ�ţ�ҩƷ�շ���¼.��Ʒ�ϸ�֤��ҽ������ʱ����
  n_Group_Sno              Number; --N   �������(�����洢)��1��2��3
  n_Cadn_Id                Number; --N   1ҩƷͨ������id(ҩ��ID)(������)
  n_Advice_Id              Number; --N   0ҽ��ID
  n_Drug_Method_Id         Number; --N   1��ҩ;��id(������):������ĿID
  v_Drug_Method_Name       Varchar2(32767); --C   1��ҩ;������
  v_Drug_Method_Class_Code Varchar2(32767); --C   1��ҩ;������(������):ִ�з�����
  n_Drug_Freq_Id           Number; --N   1��ҩƵ��id(������):����Ƶ�ʱ���
  v_Drug_Freq_Name         Varchar2(32767); --C   1��ҩƵ������d(������):
  n_Emergency_Tag          Number; --N   ҽ����¼�еĽ�����־(0-��ͨ;1-����;2-��¼(��������Ч))
  n_Effectivetime          Number; --N   0ҽ����Ч
  n_Denominated            Number; --N   0�Ƽ����ԣ�0-2-�Ƽ�����,3-��Ժ��ҩ,4-��ȡҩ�Ƽ�����(0-�����Ƽۣ�1-���Ƽۣ�2-�ֹ��Ƽ�))
  n_Frequency              Number; --N   0Ƶ��
  n_Single                 Number; --N   0����
  v_Usage                  Varchar2(32767); --C   0�÷�
  v_Rcpdtl_St_Result       Varchar2(32767); --N   Ƥ�Խ��(������)1-���ԣ�2-���ԣ�3-���ԣ�4-������ҩ��������ʱ��ȷ��������Ƥ�Խ����ZLHISĿǰ֧�ֲ�ȫ
  v_Rcpdtl_Drask           Varchar2(32767); --C   ʹ������(������)
  v_Diag_Name              Varchar2(32767); --C  0 ������ƣ�������)�����ﴫ�룬�������
  n_Use_Mode               Number;
  n_������Դ               Number;
  n_��Һ���               Number(3);

  Vj_Iitem  Varchar2(32767);
  Cjl_Iitem Clob;
  --����ҽ������
  Cursor c_Out
  (
    �Һŵ�_In ����ҽ����¼.�Һŵ�%Type,
    ���ͺ�_In ����ҽ������.���ͺ�%Type
  ) Is
    Select a.��ҳid, a.Id, a.���id, b.No, a.������־, a.�Ƽ�����, a.��������, d.���� As �÷�, b.��¼����, b.�������, a.Ƥ�Խ��, 'ҽ������' As ժҪ, a.��ҩĿ��,
           a.����˵��, a.ҽ������, a.ҽ����Ч, a.Ƶ�ʴ���, a.�������, a.������Ŀid, c.������Ŀid As ��ҩ;��id, d.���� As ��ҩ;������, d.��� ��ҩ���, Null ��ҩ��������,
           d.ִ�з��� As ��ҩ;������, e.���� As ����Ƶ�ʱ���, c.ִ��Ƶ�� As ��ҩƵ������, b.��ҩ��, b.���ͺ�, a.ִ�б��, a.ִ������ As ҩ��ִ������, c.ִ������ As ��ҩִ������
    From ����ҽ����¼ A, ����ҽ������ B, ����ҽ����¼ C, ������ĿĿ¼ D, ����Ƶ����Ŀ E
    Where a.���id = c.Id(+) And c.������Ŀid = d.Id(+) And c.ִ��Ƶ�� = e.����(+) And a.Id = b.ҽ��id And a.�Һŵ� = �Һŵ�_In And
          b.���ͺ� = ���ͺ�_In
    Order By a.���, b.���ͺ�;

  Type t_Order Is Table Of c_Out%RowType;
  r_Odr t_Order;

  Cursor c_In
  (
    ����id_In ����ҽ����¼.����id%Type,
    ��ҳid_In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In ����ҽ������.���ͺ�%Type
  ) Is
    Select a.��ҳid, a.Id, a.���id, b.No, a.������־, a.�Ƽ�����, a.��������, d.���� As �÷�, b.��¼����, b.�������, a.Ƥ�Խ��, 'ҽ������' As ժҪ, a.��ҩĿ��,
           a.����˵��, a.ҽ������, a.ҽ����Ч, a.Ƶ�ʴ���, a.�������, a.������Ŀid, c.������Ŀid As ��ҩ;��id, d.���� As ��ҩ;������, d.��� ��ҩ���,
           d.�������� ��ҩ��������, d.ִ�з��� As ��ҩ;������, e.���� As ����Ƶ�ʱ���, c.ִ��Ƶ�� As ��ҩƵ������, b.��ҩ��, b.���ͺ�, a.ִ�б��, a.ִ������ As ҩ��ִ������,
           c.ִ������ As ��ҩִ������
    From ����ҽ����¼ A, ����ҽ������ B, ����ҽ����¼ C, ������ĿĿ¼ D, ����Ƶ����Ŀ E
    Where a.���id = c.Id(+) And c.������Ŀid = d.Id(+) And c.ִ��Ƶ�� = e.����(+) And a.Id = b.ҽ��id And a.����id = ����id_In And
          a.��ҳid = ��ҳid_In And b.���ͺ� = ���ͺ�_In
    Order By a.���, b.���ͺ�;
  -- �������
  Cursor c_Adv(P��id Number) Is
    Select a.Id From ����ҽ����¼ A Where a.Id = P��id Or a.���id = P��id Order By a.���;

  --����ҽ���������
  Cursor c_Diag
  (
    ����id_In ����ҽ����¼.����id%Type,
    �Һ�id_In ������ϼ�¼.��ҳid%Type
  ) Is
    Select a.ҽ��id, b.������� As ����
    From �������ҽ�� A, ������ϼ�¼ B
    Where a.���id = b.Id And b.����id = ����id_In And b.��ҳid = �Һ�id_In And Nvl(b.¼�����, '01') = '01'
    Order By a.ҽ��id;

  Type t_Diag Is Table Of c_Diag%RowType;
  Rs_Diag t_Diag;

  Cursor Ctest(�Һŵ�_In ����ҽ����¼.�Һŵ�%Type) Is
    Select a.Ƥ�Խ��, a.ҩ��id, a.ҩƷid
    From (Select a.Ƥ�Խ��, c.��Ŀid As ҩ��id, 0 As ҩƷid, a.��ʼִ��ʱ��
           From ����ҽ����¼ A, �����÷����� C
           Where a.������Ŀid = c.�÷�id And Nvl(c.����, 0) = 0 And Nvl(a.ҽ��״̬, 0) = 8 And a.Ƥ�Խ�� Is Not Null And a.�Һŵ� = �Һŵ�_In
           Union All
           Select a.Ƥ�Խ��, b.ҩ��id, b.ҩƷid, a.��ʼִ��ʱ��
           From ����ҽ����¼ A, ҩƷ��� B, ҩƷ�÷����� C
           Where a.������Ŀid = c.�÷�id And b.ҩƷid = c.ҩƷid And Nvl(c.����, 0) = 0 And Nvl(a.ҽ��״̬, 0) = 8 And a.Ƥ�Խ�� <> '����' And
                 a.�Һŵ� = �Һŵ�_In
           Union All
           Select a.Ƥ�Խ��, c.��Ŀid As ҩ��id, 0 As ҩƷid, a.��ʼִ��ʱ��
           From ����ҽ����¼ A, �����÷����� C
           Where a.������Ŀid = c.�÷�id And Nvl(c.����, 0) = 0 And a.Ƥ�Խ�� = '����' And a.�Һŵ� = �Һŵ�_In) A
    Order By a.��ʼִ��ʱ�� Desc;
  Type t_Test Is Table Of Ctest%RowType;
  Rs_Test t_Test;

  Cursor Ctestin
  (
    ����id_In ����ҽ����¼.����id%Type,
    ��ҳid_In ����ҽ����¼.��ҳid%Type
  ) Is
    Select a.Ƥ�Խ��, a.ҩ��id, a.ҩƷid
    From (Select a.Ƥ�Խ��, c.��Ŀid As ҩ��id, 0 As ҩƷid, a.��ʼִ��ʱ��
           From ����ҽ����¼ A, �����÷����� C
           Where a.������Ŀid = c.�÷�id And Nvl(c.����, 0) = 0 And Nvl(a.ҽ��״̬, 0) = 8 And a.Ƥ�Խ�� Is Not Null And
                 a.����id = ����id_In And a.��ҳid = ��ҳid_In
           Union All
           Select a.Ƥ�Խ��, b.ҩ��id, b.ҩƷid, a.��ʼִ��ʱ��
           From ����ҽ����¼ A, ҩƷ��� B, ҩƷ�÷����� C
           Where a.������Ŀid = c.�÷�id And b.ҩƷid = c.ҩƷid And Nvl(c.����, 0) = 0 And Nvl(a.ҽ��״̬, 0) = 8 And a.Ƥ�Խ�� <> '����' And
                 a.����id = ����id_In And a.��ҳid = ��ҳid_In
           Union All
           Select a.Ƥ�Խ��, c.��Ŀid As ҩ��id, 0 As ҩƷid, a.��ʼִ��ʱ��
           From ����ҽ����¼ A, �����÷����� C
           Where a.������Ŀid = c.�÷�id And Nvl(c.����, 0) = 0 And a.Ƥ�Խ�� = '����' And a.����id = ����id_In And a.��ҳid = ��ҳid_In) A
    Order By a.��ʼִ��ʱ�� Desc;

  --��ȡƤ�Խ��
  Procedure Pp_Test
  (
    ҩ��id_In    Number,
    Ƥ�Խ��_Out Out Varchar2
  ) Is
  Begin
    Ƥ�Խ��_Out := Null;
    For I In 1 .. Rs_Test.Count Loop
      If Rs_Test(I).ҩ��id = ҩ��id_In Then
        Ƥ�Խ��_Out := Rs_Test(I).Ƥ�Խ��;
		exit;
      End If;
    End Loop;
  End;

  --��ȡ���
  Procedure Pp_Diag
  (
    ҽ��id_In Number,
    ���_Out  Out Varchar2
  ) Is
  Begin
    ���_Out := Null;
    For I In 1 .. Rs_Diag.Count Loop
      If Rs_Diag(I).ҽ��id = ҽ��id_In Then
        ���_Out := ���_Out || ',' || Rs_Diag(I).����;
      End If;
    End Loop;
    ���_Out := Substr(���_Out, 2);
  End;

  Procedure p_Getbaseinfo
  (
    Rgstid_In Number,
    �Һŵ�_In ����ҽ����¼.�Һŵ�%Type,
    ����id_In ����ҽ����¼.����id%Type,
    ��ҳid_In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In ����ҽ������.���ͺ�%Type
  ) As
  Begin
    If Nvl(Rgstid_In, 0) = 0 Then
      Open c_In(����id_In, ��ҳid_In, ���ͺ�_In);
      Fetch c_In Bulk Collect
        Into r_Odr;
      Close c_In;
      --Ƥ�Խ��
      Open Ctestin(����id_In, ��ҳid_In);
      Fetch Ctestin Bulk Collect
        Into Rs_Test;
      Close Ctestin;
    Else
      Open c_Out(�Һŵ�_In, ���ͺ�_In);
      Fetch c_Out Bulk Collect
        Into r_Odr;
      Close c_Out;
      --Ƥ�Խ��
      Open Ctest(�Һŵ�_In);
      Fetch Ctest Bulk Collect
        Into Rs_Test;
      Close Ctest;
      --���
      Open c_Diag(����id_In, Rgstid_In);
      Fetch c_Diag Bulk Collect
        Into Rs_Diag;
      Close c_Diag;
    End If;
  End;

  Procedure p_Pivasbill_Get
  (
    ����id_In  Number,
    ��ҳid_In  Number,
    ���ͺ�_In  Number,
    ҽ��ids_In Varchar2,
    j_Out      Out Clob
  ) As
    -----------------------------------------------------------
    --����:��HIS���л�ȡ���Է��͵������ҽ����Ϣ
    -- ���
    -- input
    --   pati_id                   N 1 ����id
    --   pati_pageid               N 1 ��ҳid
    --   send_num                  N 1 ���ͺ�
    --   order_ids                 C 1 ������ҽ��id ƴ��
    -- ����
    --    pati_id                  N  1  ����id
    --    page_id                  N  1  ��ҳID
    --    pati_name                C  1  ����
    --    pati_sex                 C  1  �Ա�
    --    pati_age                 C  1  ����
    --    inpatient_num            N  1  סԺ��
    --    pati_bed                 C  1  ����
    --    pati_wardarea_id         N  1  ���˲���id
    --    pati_deptid              N  1  ���˿���id
    --    advice_list[]            ��ҽ���б�
    --      advice_id              N  1  ҽ��id --��ҽ��ID(��ҩ;��)
    --      advice_send_no         N  1  ���ͺ� 351,  --���ͺ�
    --      effective_time         N  1  ҽ����Ч
    --      drug_method_id         N  1  ��ҩ;��id
    --      is_tpn                 N  1  �Ƿ�tpn
    --      advice_frequency       C  1  ִ��Ƶ�Σ�һ������
    --      acvice_drug_list[]     ҩ����Ϣ
    --         advice_id           N  1  ��id
    --         advice_rcpno        C  1  ҩ�����Ͳ����ķ���no
    --      advice_exetime_list[]  ҽ��ִ��ʱ�䣬3����,���εķ�����Ϣ+��ʷ�ķ�����Ϣ, 3����
    --         advice_id           N  1  ��ҩ;��ҽ��
    --         advice_send_no      N  1  ���ͺ�
    --         advice_require_time C  1  2019-07-02 16:30:00    --Ҫ��ʱ��
    -----------------------------------------------------------
    v_Exetimes   Varchar2(32767);
    n_Preҽ��id  Number(18) := 0;
    n_�������id Number(18);
  
    Vj_Pati      Varchar2(32767);
    Vj_Last      Varchar2(32767);
    Vj_Jsonlist1 Varchar2(32767);
    Cj_Ad        Clob;
    Vj_Ad        Varchar2(32767);
  
    Cursor c_Ad Is
      Select Nvl(a.���id, a.Id) As ��ҽ��id, b.ҽ��id, a.����id, a.��ҳid, c.����, c.�Ա�, c.����, c.סԺ��, c.��Ժ���� As ����, a.���˿���id,
             c.��ǰ����id As ���˲���id, Nvl(b.ִ�в���id, 0) ִ�в���id, e.Id As ��ҩ;��id, d.ִ��Ƶ��, Decode(e.ִ�б��, 2, 1, 0) As Tpn, d.ҽ����Ч,
             b.���ͺ�, b.No, b.������, To_Char(b.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��
      From ����ҽ����¼ A, ����ҽ������ B, ������ҳ C, ����ҽ����¼ D, ������ĿĿ¼ E
      Where a.Id = b.ҽ��id And a.����id = c.����id And a.��ҳid = c.��ҳid And a.���id = d.Id And d.������Ŀid = e.Id And
            a.����id = ����id_In And a.��ҳid = ��ҳid_In And b.���ͺ� = ���ͺ�_In And a.������� In ('5', '6') And
            Instr(',' || ҽ��ids_In || ',', ',' || Nvl(a.���id, a.Id) || ',') > 0
      Order By Nvl(a.���id, a.Id), a.���;
  
    Cursor c_Ext(ҽ��id_In ҽ��ִ��ʱ��.ҽ��id%Type) Is
      Select /*+cardinality(j,10)*/
       a.���ͺ�, To_Char(a.Ҫ��ʱ��, 'YYYY-MM-DD HH24:MI:SS') As Ҫ��ʱ��
      From ҽ��ִ��ʱ�� A
      Where a.Ҫ��ʱ�� Between Sysdate - 3 And Sysdate + 3 And a.ҽ��id = ҽ��id_In;
  Begin
  
    Vj_Ad := Null;
    For R In c_Ad Loop
      If r.ִ�в���id <> 0 Then
        n_�������id := r.ִ�в���id;
      End If;
    
      If n_Preҽ��id = 0 Then
        Vj_Pati := Vj_Pati || '{"pati_id":' || r.����id;
        Vj_Pati := Vj_Pati || ',"page_id":' || r.��ҳid;
        Vj_Pati := Vj_Pati || ',"pati_name":"' || Zljsonstr(r.����) || '"';
        Vj_Pati := Vj_Pati || ',"pati_sex":"' || Zljsonstr(r.�Ա�) || '"';
        Vj_Pati := Vj_Pati || ',"pati_age":"' || Zljsonstr(r.����) || '"';
        Vj_Pati := Vj_Pati || ',"inpatient_num":"' || r.סԺ�� || '"';
        Vj_Pati := Vj_Pati || ',"pati_bed":"' || Zljsonstr(r.����) || '"';
        Vj_Pati := Vj_Pati || ',"pati_wardarea_id":' || Nvl(r.���˲���id || '', 'null');
        Vj_Pati := Vj_Pati || ',"pati_deptid":' || Nvl(r.���˿���id || '', 'null');
      End If;
    
      If n_Preҽ��id <> 0 And n_Preҽ��id <> r.��ҽ��id Then
        Vj_Ad := Vj_Ad || Vj_Last || ',"advice_drug_list":[' || Substr(Vj_Jsonlist1, 2) || ']';
      
        --���εķ���ʱ������εķ��͵�ִ��ʱ�����Ϣ
        v_Exetimes := Null;
        For r_Ext In c_Ext(n_Preҽ��id) Loop
          --ĳ��ҽ����ִ��ʱ���
          v_Exetimes := v_Exetimes || ',{"advice_send_no":' || r_Ext.���ͺ�;
          v_Exetimes := v_Exetimes || ',"advice_require_time":"' || r_Ext.Ҫ��ʱ�� || '"';
          v_Exetimes := v_Exetimes || '}';
        End Loop;
      
        Vj_Ad := Vj_Ad || ',"advice_exetime_list":[' || Substr(v_Exetimes, 2) || ']';
        Vj_Ad := Vj_Ad || '}';
      
        If Length(Vj_Ad) > 20000 Then
          Cj_Ad := Cj_Ad || Vj_Ad;
          Vj_Ad := Null;
        End If;
        Vj_Jsonlist1 := Null;
      End If;
    
      n_Preҽ��id := r.��ҽ��id;
    
      --ҩƷ��ҽ����Ϣ
      Vj_Jsonlist1 := Vj_Jsonlist1 || ',{"advice_id":' || r.ҽ��id;
      Vj_Jsonlist1 := Vj_Jsonlist1 || ',"advice_rcpno":"' || r.No || '"';
      Vj_Jsonlist1 := Vj_Jsonlist1 || '}';
    
      --�������һ�ε�jsonƴ��
      Vj_Last := ',{"pivas_deptid":' || n_�������id;
      Vj_Last := Vj_Last || ',"advice_id":' || r.��ҽ��id;
      Vj_Last := Vj_Last || ',"advice_send_no":' || r.���ͺ�;
      Vj_Last := Vj_Last || ',"effective_time":' || r.ҽ����Ч;
      Vj_Last := Vj_Last || ',"drug_method_id":' || r.��ҩ;��id;
      Vj_Last := Vj_Last || ',"is_tpn":' || Nvl(r.Tpn || '', 'null');
      Vj_Last := Vj_Last || ',"advice_frequency":"' || Zljsonstr(r.ִ��Ƶ��) || '"';
    End Loop;
  
    If n_Preҽ��id <> 0 Then
      Vj_Ad := Vj_Ad || Vj_Last || ',"advice_drug_list":[' || Substr(Vj_Jsonlist1, 2) || ']';
    
      --���εķ���ʱ������εķ��͵�ִ��ʱ�����Ϣ
      v_Exetimes := Null;
      For r_Ext In c_Ext(n_Preҽ��id) Loop
        --ĳ��ҽ����ִ��ʱ���
        v_Exetimes := v_Exetimes || ',{"advice_send_no":' || r_Ext.���ͺ�;
        v_Exetimes := v_Exetimes || ',"advice_require_time":"' || r_Ext.Ҫ��ʱ�� || '"';
        v_Exetimes := v_Exetimes || '}';
      End Loop;
    
      Vj_Ad := Vj_Ad || ',"advice_exetime_list":[' || Substr(v_Exetimes, 2) || ']';
      Vj_Ad := Vj_Ad || '}';
    End If;
  
    Cj_Ad := Cj_Ad || Vj_Ad;
    If Cj_Ad Is Not Null Then
      j_Out := Vj_Pati || ',"advice_list":[' || Substr(Cj_Ad, 2) || ']}';
    End If;
  End;
Begin
  --�������
  If Json_In Is Null Then
    Select f_List2str(Cast(Collect(a.����id || '') As t_Strlist), ',')
    Into v_Vals
    From (Select a.����id From ����ҽ���쳣��¼ A Where a.�������� = 1 Group By a.����id) A;
  Else
    j_Tmp  := Pljson(Json_In);
    j_Json := j_Tmp.Get_Pljson('input');
    v_Vals := j_Json.Get_Clob('pati_ids');
    If v_Vals Is Null Then
      Select f_List2str(Cast(Collect(a.����id || '') As t_Strlist), ',')
      Into v_Vals
      From (Select a.����id From ����ҽ���쳣��¼ A Where a.�������� = 1 Group By a.����id) A;
    End If;
  End If;

  l_Vals := t_Strlist();
  While v_Vals Is Not Null Loop
    If Length(v_Vals) <= 4000 Then
      l_Vals.Extend;
      l_Vals(l_Vals.Count) := v_Vals;
      v_Vals := Null;
    Else
      l_Vals.Extend;
      l_Vals(l_Vals.Count) := Substr(v_Vals, 1, Instr(v_Vals, ',', 3980) - 1);
      v_Vals := Substr(v_Vals, Instr(v_Vals, ',', 3980) + 1);
    End If;
  End Loop;

  n_�к� := 0;
  For Lp In 1 .. l_Vals.Count Loop
    For Cp In (Select /*+Cardinality(j,10)*/
                a.����id, a.��ҳid, a.�Һŵ�, b.���ͺ�
               From ����ҽ����¼ A, ����ҽ���쳣��¼ B, Table(f_Num2list(l_Vals(Lp))) J
               Where a.Id = b.ҽ��id And a.����id = j.Column_Value And b.�������� = 1
               Group By a.����id, a.��ҳid, a.�Һŵ�, b.���ͺ�) Loop
    
      n_Rgst_Id := 0;
      If Cp.�Һŵ� Is Not Null Then
        Select a.Id Into n_Rgst_Id From ���˹Һż�¼ A Where a.No = Cp.�Һŵ� And a.��¼���� = 1 And a.��¼״̬ = 1;
      End If;
    
      v_Jtmp := Null;
      v_Jtmp := v_Jtmp || ',{"pati_id":' || Cp.����id;
      v_Jtmp := v_Jtmp || ',"pati_pageid":' || Nvl(Cp.��ҳid || '', 'null');
      v_Jtmp := v_Jtmp || ',"rgst_id":' || Nvl(n_Rgst_Id || '', 'null');
      v_Jtmp := v_Jtmp || ',"rgst_no":"' || Cp.�Һŵ� || '"';
      v_Jtmp := v_Jtmp || ',"send_no":' || Cp.���ͺ�;
    
      If v_����Ա Is Null Then
        Select Max(a.������), To_Char(Max(a.����ʱ��), 'yyyy-mm-dd hh24:mi:ss')
        Into v_����Ա, v_����ʱ��
        From ����ҽ������ A
        Where a.���ͺ� = Cp.���ͺ�;
      End If;
      v_Jtmp := v_Jtmp || ',"operator_name":"' || Zljsonstr(v_����Ա) || '"';
      v_Jtmp := v_Jtmp || ',"operator_time":"' || v_����ʱ�� || '"';
    
      --��������
      If Cp.��ҳid Is Not Null Then
        Select a.�������� Into v_Tmp From ������ҳ A Where a.����id = Cp.����id And a.��ҳid = Cp.��ҳid;
        v_Jtmp := v_Jtmp || ',"pati_type":"' || v_Tmp || '"';
      End If;
    
      --S��ȡ���������Ϣ 
      v_Tmp := Null;
      For K In (Select ID, �������, �������
                From ������ϼ�¼
                Where ����id = Cp.����id And ��ҳid = Nvl(Cp.��ҳid, n_Rgst_Id) And ������� Is Not Null
                Order By �������, ��¼��Դ, ��¼���� Desc, ��ϴ��� Asc, Nvl(¼�����, '01'), Nvl(�������, 1)) Loop
        v_Tmp := v_Tmp || ',{"diag_rec_id":' || k.Id; -- : N ��ϼ�¼id
        v_Tmp := v_Tmp || ',"diag_type":' || k.�������; -- ��N ������� 
        v_Tmp := v_Tmp || ',"diag_name":"' || Zljsonstr(k.�������) || '"'; -- ��C ������ƣ����������
        v_Tmp := v_Tmp || '}';
      End Loop;
      If v_Tmp Is Not Null Then
        v_Jtmp := v_Jtmp || ',"diag_list":[' || Substr(v_Tmp, 2) || ']';
      End If;
      --E��ȡ���������Ϣ 
    
      n_�к� := n_�к� + 1;
      If n_�к� = 1 Then
        c_Outtmp := Substr(v_Jtmp, 2);
      Else
        c_Outtmp := c_Outtmp || v_Jtmp;
      End If;
    
      Cjl_Iitem := Null;
      Vj_Iitem  := Null;
      --��ȡҽ����ϸ�����Ϣ���浽 r_Odr ��
      p_Getbaseinfo(n_Rgst_Id, Cp.�Һŵ�, Cp.����id, Cp.��ҳid, Cp.���ͺ�);
      For Ol In 1 .. r_Odr.Count Loop
        n_Send_No                := r_Odr(Ol).���ͺ�;
        n_Advice_Id              := r_Odr(Ol).Id;
        n_Drug_Method_Id         := r_Odr(Ol).��ҩ;��id;
        v_Drug_Method_Name       := r_Odr(Ol).��ҩ;������;
        v_Drug_Method_Class_Code := r_Odr(Ol).����Ƶ�ʱ���;
        n_Drug_Freq_Id           := r_Odr(Ol).����Ƶ�ʱ���;
        v_Drug_Freq_Name         := r_Odr(Ol).��ҩƵ������;
        n_Emergency_Tag          := r_Odr(Ol).������־;
        n_Denominated            := r_Odr(Ol).�Ƽ�����;
        n_Frequency              := r_Odr(Ol).Ƶ�ʴ���;
        n_Single                 := r_Odr(Ol).��������;
        v_Usage                  := r_Odr(Ol).�÷�;
        v_Rcpdtl_Drask           := r_Odr(Ol).ҽ������;
        n_Effectivetime          := r_Odr(Ol).ҽ����Ч;
        n_Cadn_Id                := r_Odr(Ol).������Ŀid;
        n_��id                   := Nvl(r_Odr(Ol).���id, r_Odr(Ol).Id);
      
        n_Group_Sno := Null;
        n_Cnt       := 0;
        For Ir In c_Adv(n_��id) Loop
          n_Cnt := n_Cnt + 1;
          If Ir.Id = r_Odr(Ol).Id Then
            n_Group_Sno := n_Cnt;
            Exit;
          End If;
        End Loop;
      
        v_Take_No   := r_Odr(Ol).��ҩ��;
        v_Diag_Name := Null;
        --���ﲡ������
        If Cp.�Һŵ� Is Not Null Then
          Pp_Diag(n_��id, v_Diag_Name);
        End If;
      
        If r_Odr(Ol).������� In ('5', '6') Then
          Pp_Test(n_Cadn_Id, v_Rcpdtl_St_Result);
        Else
          v_Rcpdtl_St_Result := Null;
        End If;
      
        --n_use_mode--
        n_Use_Mode := 0;
        If r_Odr(Ol).ҩ��ִ������ = 4 And r_Odr(Ol).��ҩִ������ = 5 Then
          n_Use_Mode := 1;
        Elsif r_Odr(Ol).ִ�б�� = 1 And r_Odr(Ol).ҩ��ִ������ = 4 And r_Odr(Ol).��ҩִ������ = 2 Then
          n_Use_Mode := 2;
        End If;
      
        If r_Odr(Ol).��¼���� = 1 Or r_Odr(Ol).��¼���� = 2 And r_Odr(Ol).������� = 1 Then
          n_������Դ := 1;
        Else
          n_������Դ := 2;
        End If;
      
        Vj_Iitem := Vj_Iitem || ',{"advice_id":' || n_Advice_Id;
        Vj_Iitem := Vj_Iitem || ',"group_sno":' || Nvl(n_Group_Sno || '', 'null');
        Vj_Iitem := Vj_Iitem || ',"effectivetime":' || r_Odr(Ol).ҽ����Ч;
        Vj_Iitem := Vj_Iitem || ',"drug_method_id":' || Nvl(n_Drug_Method_Id || '', 'null');
        Vj_Iitem := Vj_Iitem || ',"drug_method_name":"' || Zljsonstr(v_Drug_Method_Name) || '"';
        Vj_Iitem := Vj_Iitem || ',"drug_method_class_code":"' || Zljsonstr(v_Drug_Method_Class_Code) || '"';
        Vj_Iitem := Vj_Iitem || ',"drug_freq_id":' || Nvl(n_Drug_Freq_Id || '', 'null');
        Vj_Iitem := Vj_Iitem || ',"drug_freq_name":"' || Zljsonstr(v_Drug_Freq_Name) || '"';
        Vj_Iitem := Vj_Iitem || ',"emergency_tag":' || Nvl(n_Emergency_Tag || '', 'null');
        Vj_Iitem := Vj_Iitem || ',"fee_mode":0';
        Vj_Iitem := Vj_Iitem || ',"use_mode":' || Zljsonstr(n_Use_Mode, 1);
        Vj_Iitem := Vj_Iitem || ',"frequency":' || Zljsonstr(n_Frequency, 1);
        Vj_Iitem := Vj_Iitem || ',"single":' || Zljsonstr(n_Single, 1);
        Vj_Iitem := Vj_Iitem || ',"usage":"' || Zljsonstr(v_Usage) || '"';
        Vj_Iitem := Vj_Iitem || ',"rcpdtl_st_result":"' || Zljsonstr(v_Rcpdtl_St_Result) || '"';
        Vj_Iitem := Vj_Iitem || ',"rcpdtL_excs_desc":"' || Zljsonstr(r_Odr(Ol).����˵��) || '"';
        Vj_Iitem := Vj_Iitem || ',"rcpdtL_drask":"' || Zljsonstr(v_Drug_Method_Name) || '"';
        Vj_Iitem := Vj_Iitem || ',"memo":"ҽ������"';
        Vj_Iitem := Vj_Iitem || ',"diag_name":"' || Zljsonstr(v_Diag_Name) || '"';
        Vj_Iitem := Vj_Iitem || ',"take_no":"' || Zljsonstr(r_Odr(Ol).��ҩ��) || '"';
        Vj_Iitem := Vj_Iitem || ',"advice_purpose":"' || Zljsonstr(r_Odr(Ol).��ҩĿ��) || '"';
        Vj_Iitem := Vj_Iitem || ',"fee_source":' || n_������Դ;
        Vj_Iitem := Vj_Iitem || ',"fee_billtype":' || r_Odr(Ol).��¼����;
        Vj_Iitem := Vj_Iitem || ',"fee_no":"' || r_Odr(Ol).No || '"';
        Vj_Iitem := Vj_Iitem || '}';
      
        If Length(Vj_Iitem) > 30000 Then
          If Cjl_Iitem Is Null Then
            Cjl_Iitem := Substr(Vj_Iitem, 2);
          Else
            Cjl_Iitem := Cjl_Iitem || Vj_Iitem;
          End If;
          Vj_Iitem := Null;
        End If;
      
        --a�������-----------------------------
        If r_Odr(Ol).��ҩ�������� = '2' And r_Odr(Ol).��ҩ��� = 'E' And r_Odr(Ol).��ҩ;������ = '1' Then
          Select Count(1)
          Into n_��Һ���
          From ����ҽ���쳣��¼ A
          Where a.ҽ��id = n_��id And a.���ͺ� = r_Odr(Ol).���ͺ� And a.�������� = 3;
          If n_��Һ��� > 0 Then
            --�ռ���������ͬ���쳣����ҽ��
            If Instr(',' || v_Pҽ��ids || ',', ',' || n_��id || ',') = 0 Then
              v_Pҽ��ids := v_Pҽ��ids || ',' || n_��id;
            End If;
          End If;
        End If;
        --e�������-----------------------------
      End Loop;
    
      If Cjl_Iitem Is Null Then
        c_Outtmp := c_Outtmp || ',"order_list":[' || Substr(Vj_Iitem, 2) || ']';
      Else
        c_Outtmp := c_Outtmp || ',"order_list":[' || Cjl_Iitem || Vj_Iitem || ']';
      End If;
    
      --a�������-----------------------------
      If v_Pҽ��ids Is Not Null Then
        p_Pivasbill_Get(Cp.����id, Cp.��ҳid, Cp.���ͺ�, Substr(v_Pҽ��ids, 2), b_Pivasout);
        If b_Pivasout Is Not Null Then
          c_Outtmp := c_Outtmp || ',"pivas_list":[' || b_Pivasout || ']';
        End If;
      End If;
      --e�������-----------------------------
    
      c_Outtmp := c_Outtmp || '}';
    End Loop;
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","pati_bill_list":[' || c_Outtmp || ']}}';
Exception
  When Err_Item Then
    Json_Out := Zljsonout(v_Err);
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Cissvr_Getdrugerrdata;
/
CREATE OR REPLACE Procedure Zl_Cissvr_Geterrbillinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) Is
  -------------------------------------------------------------------------------------------------
  --���ܣ���ȡҽ��վԤԼ�����쳣����
  --��Σ�json��ʽ
  --Input
  --   rgst_no               C  1 �Һŵ�
  --���Σ�json��ʽ
  --Json_Out
  --   code                  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --   message               C  1  Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --   snyc_status           N  1  ͬ��״̬��-1-����IDΪͬ����ҵ������;1-����֧��δУ��;2-У�����δ֧���ɹ�
  --   pati_id               N  1  ����ID
  --   outpno                C  1  �����
  --   rgst_balance          C  1  ������Ϣ:ͬ��״̬Ϊ2ʱ��������������Ϣ
  -------------------------------------------------------------------------------------------------

  v_�Һŵ�   ���˽����쳣��¼.Ԥ������%Type;
  n_����id   ���˽����쳣��¼.����id%Type;
  n_�����   ���˽����쳣��¼.�����%Type;
  n_ͬ��״̬ ���˽����쳣��¼.ͬ��״̬%Type;
  v_������Ϣ ���˽����쳣��¼.������Ϣ%Type;
  j_Json     Pljson;
  j_Jsontmp  Pljson;
  v_Temp     Varchar2(100);
Begin
  --�������
  j_Jsontmp := Pljson(Json_In);
  j_Json    := j_Jsontmp.Get_Pljson('input');
  v_�Һŵ�  := j_Json.Get_String('rgst_no');

  Begin
    Select ����id, �����, ͬ��״̬, ������Ϣ
    Into n_����id, n_�����, n_ͬ��״̬, v_������Ϣ
    From ���˽����쳣��¼
    Where Ԥ������ = v_�Һŵ� And �������� = 4 And Rownum < 2;
  Exception
    When Others Then
      Json_Out := '{"output":{"code":1,"message":"�ɹ�","snyc_status":0}}';
      Return;
  End;
  --ȥ��'input'
  If n_ͬ��״̬ = 2 Then
    j_Jsontmp := Pljson();
    j_Json    := Pljson(v_������Ϣ);
    j_Jsontmp := j_Json.Get_Pljson('input');

    v_������Ϣ := Empty_Clob();
    Dbms_Lob.Createtemporary(v_������Ϣ, True);
    j_Jsontmp.To_Clob(v_������Ϣ);
  Else
    v_������Ϣ := '""';
  End If;

  v_Temp := '{"output":{"code":1,"message":"�ɹ�"';
  v_Temp := v_Temp || ',"pati_id":' || Nvl(n_����id, 0);
  v_Temp := v_Temp || ',"outpno":"' || n_����� || '"';
  v_Temp := v_Temp || ',"snyc_status":' || Nvl(n_ͬ��״̬, 0);
  v_Temp := v_Temp || ',"rgst_balance":';

  Json_Out := v_Temp || v_������Ϣ || '}}';
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Cissvr_Geterrbillinfo;
/
Create Or Replace Procedure Zl_Cissvr_Getexecadvicerecord
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ�����ָ��ִ�е�ҽ��ID������ִ��ҽ����Ϣ��¼��
  --��Σ�Json_In:��ʽ
  --  input
  --    advice_send_ids             C 1 ҽ��ID�ͷ��ͺ��ַ�����ҽ��ID1:���ͺ�1,ҽ��ID2:���ͺ�2
  --����: Json_Out,��ʽ����
  --  output
  --    code                        N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                     C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    item_list
  --         advice_id              N 1 ҽ��id
  --         advice_related_id      N 1 ���id
  --         send_no                N 1 ���ͺ�
  --         pati_id                N 1 ����id
  --         pati_ageid             N 1 ��ҳid
  --         advice_begintime       C 1 ��ʼִ��ʱ��
  --         advice_note            C 1 ҽ������
  --         nums                   C 1 ����
  --         advice_doctor_note     C 1 ҽ������
  --         advice_doctor          C 1 ����ҽ��
  --         advice_record_time     C 1 ����ʱ��
  ---------------------------------------------------------------------------
  j_Tmp            Pljson;
  j_In             Pljson;
  v_List           Varchar2(32767);
  v_ҽ������       Varchar2(4000);
  v_Order_Send_Ids Varchar2(32767);

  Cursor c_Ad Is
    Select b.Id As ҽ��id, b.���id, a.���ͺ�, b.����id, b.��ҳid, To_Char(b.��ʼִ��ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ��ʼʱ��,
           b.ҽ������ As ҽ������, a.�������� || Nvl(d.���㵥λ, c.���㵥λ) As ����, b.ҽ������, b.����ҽ��,
           To_Char(b.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��, a.����ʱ��, a.������, Nvl(d.���㵥λ, c.���㵥λ) As ���㵥λ, a.No As ���ݺ�,
           a.��¼����, a.ִ�в���id, a.�����, a.���ʱ��, b.���, b.������Դ, b.�Һŵ�, b.Ӥ��, b.����, b.ҽ����Ч, b.�������, b.������Ŀid, b.�걾��λ, b.��鷽��,
           b.����, b.��������, b.�ܸ�����, b.ִ��Ƶ��, b.������־, b.�շ�ϸĿid, b.�״�����, b.ִ��ʱ�䷽��, b.Ƥ�Խ��, b.ҽ������ As ҽ������tmp, 0 As ����ת��
    From ����ҽ������ A, ����ҽ����¼ B, ������ĿĿ¼ C, �շ���ĿĿ¼ D
    Where a.ҽ��id = b.Id And b.������Ŀid = c.Id And b.�շ�ϸĿid = d.Id(+) And
          (a.ҽ��id, a.���ͺ�) In (Select /*+cardinality(x,10)*/
                               x.C1 As ҽ��id, x.C2 As ���ͺ�
                              From Table(f_Num2list2(v_Order_Send_Ids)) X)
    Order By b.���;

  --��ȡҽ������
  Function Getִ������
  (
    Pҽ��id Number,
    P���id Number,
    P���   Varchar2,
    P����   Varchar2
  ) Return Varchar2 Is
  
    v_����      Varchar2(32767);
    v_Tmp       Varchar2(32767);
    StrƤ�Խ�� Varchar2(32767);
    P����       Number := 0;
    P��¼��     Number := 0;
    Bln��ҩ;�� Boolean := False;
    Cursor c_Pad Is
      Select a.Id, a.���id, a.�������, a.ҽ������, a.Ƥ�Խ��, a.��������, b.���㵥λ, b.��������, a.ִ��Ƶ��, a.ִ��ʱ�䷽��, b.����
      From ����ҽ����¼ A, ������ĿĿ¼ B
      Where Not (a.������� = 'E' And ���id Is Not Null) And a.������Ŀid = b.Id And (a.���id = Pҽ��id Or a.Id = Pҽ��id)
      Order By a.���;
    Type t_Pad Is Table Of c_Pad%RowType;
    Rstmp t_Pad;
  
  Begin
  
    If (P��� = 'C' And Nvl(P���id, 0) <> 0) Or P��� = 'D' Then
      v_���� := P����;
    Elsif P��� <> 'E' Or Nvl(P���id, 0) <> 0 Then
      v_���� := P����;
      If P��� = 'E' Then
        Select a.ҽ������ Into v_���� From ����ҽ����¼ A Where a.Id = P���id;
      End If;
    Else
      --���ΪE,�����ID=0
      Open c_Pad;
      Fetch c_Pad Bulk Collect
        Into Rstmp;
      Close c_Pad;
      P��¼�� := Rstmp.Count;
      For I In 1 .. P��¼�� Loop
        If Nvl(Rstmp(I).���id, 0) = Pҽ��id Then
          P���� := P���� + 1;
          If Rstmp(I).������� In ('5', '6') Then
            Bln��ҩ;�� := True;
          End If;
        End If;
      End Loop;
      If Not Bln��ҩ;�� Then
        v_���� := P����;
        If Rstmp(1).������� = 'E' And Rstmp(1).�������� = '1' Then
          StrƤ�Խ�� := '��Ƥ�Խ����' || Rstmp(1).Ƥ�Խ��;
          For R In (Select b.������Ӧ, b.����ʱ��
                    From ����ҽ����¼ A, ���˹�����¼ B, ������ĿĿ¼ C, �����÷����� D
                    Where a.����id = b.����id And a.������Ŀid = d.�÷�id And d.��Ŀid = c.Id And c.��� In ('5', '6') And
                          d.��Ŀid = b.ҩ��id And Nvl(d.����, 0) = 0 And
                          b.��¼ʱ�� = (Select Max(����ʱ��) From ����ҽ��״̬ Where ҽ��id = a.Id And �������� = 10) And a.Id = Pҽ��id And
                          Rownum < 2) Loop
          
            StrƤ�Խ�� := StrƤ�Խ�� || ',����ʱ�䣺' || To_Char(r.����ʱ��, 'yyyy-mm-dd');
            If r.������Ӧ Is Not Null Then
              StrƤ�Խ�� := StrƤ�Խ�� || ',������Ӧ��' || r.������Ӧ;
            End If;
          End Loop;
        End If;
      Else
        --��ҩ;��
        v_���� := Null;
        For I In 1 .. P���� Loop
          If I = P���� Then
            v_���� := v_���� || Chr(13) || '��';
          Else
            v_���� := v_���� || Chr(13) || '��';
          End If;
          v_���� := v_���� || Rstmp(I).ҽ������;
          If Rstmp(I).�������� Is Not Null Then
            v_���� := v_���� || ' ' || Round(Rstmp(I).��������, 5) || Rstmp(I).���㵥λ;
          End If;
        End Loop;
        v_Tmp := Rstmp(P��¼��).���� || ',' || Rstmp(P��¼��).ִ��Ƶ��;
        If Rstmp(P��¼��).ִ��ʱ�䷽�� Is Not Null Then
          v_Tmp := v_Tmp || '(' || Rstmp(P��¼��).ִ��ʱ�䷽�� || ')';
        End If;
        v_Tmp  := v_Tmp || ':ÿ' || Rstmp(P��¼��).���㵥λ;
        v_���� := v_Tmp || ' ' || v_����;
      End If;
    End If;
    Return v_���� || StrƤ�Խ��;
  End Getִ������;

Begin
  --�������
  j_In             := Pljson(Json_In);
  j_Tmp            := j_In.Get_Pljson('input');
  v_Order_Send_Ids := j_Tmp.Get_String('advice_send_ids');

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[';

  For R In c_Ad Loop
  
    v_ҽ������ := Getִ������(r.ҽ��id, r.���id, r.�������, r.ҽ������);
  
    v_List := v_List || ',{';
    v_List := v_List || '"advice_id":' || r.ҽ��id;
    v_List := v_List || ',"advice_related_id":' || Nvl(r.���id || '', 'null');
    v_List := v_List || ',"send_no":' || r.���ͺ�;
    v_List := v_List || ',"pati_id":' || r.����id;
    v_List := v_List || ',"pati_ageid":' || Nvl(r.��ҳid || '', 'null');
    v_List := v_List || ',"advice_begintime":"' || r.��ʼʱ�� || '"';
    v_List := v_List || ',"advice_note":"' || Zljsonstr(v_ҽ������) || '"';
    v_List := v_List || ',"nums":"' || Zljsonstr(r.����) || '"';
    v_List := v_List || ',"advice_doctor_note":"' || Zljsonstr(r.ҽ������) || '"';
    v_List := v_List || ',"advice_doctor":"' || Zljsonstr(r.����ҽ��) || '"';
    v_List := v_List || ',"advice_record_time":"' || r.����ʱ�� || '"';
    v_List := v_List || '}';
  
  End Loop;

  Json_Out := Json_Out || Substr(v_List, 2) || ']}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getexecadvicerecord;
/
Create Or Replace Procedure Zl_Cissvr_Getgroupadviceinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:��ȡһ����ҩ��ҽ������
  --��Σ�Json_In:��ʽ
  -- input
  --   advice_ids           C   1 ҽ��ID�������Ӣ�ĵĶ��ŷָ�
  --����: Json_Out,��ʽ����
  --  output
  --    code                C  1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message             C  1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    advice_list[]              [����]ÿ��ҽ����Ϣ
  --      advice_id         N   ҽ��id������ҽ��ID
  --      advice_note       C   ҽ������
  ---------------------------------------------------------------------------
  j_In        Pljson;
  j_Json      Pljson;
  c_ҽ��ids   Clob;
  l_ҽ��id    t_Strlist := t_Strlist();
  n_Firstitem Number(1);
  v_Temp      Varchar2(32767);
Begin
  --�������
  j_In      := Pljson(Json_In);
  j_Json    := j_In.Get_Pljson('input');
  c_ҽ��ids := j_Json.Get_Clob('advice_ids');

  While c_ҽ��ids Is Not Null Loop
    If Length(c_ҽ��ids) <= 4000 Then
      l_ҽ��id.Extend;
      l_ҽ��id(l_ҽ��id.Count) := c_ҽ��ids;
      c_ҽ��ids := Null;
    Else
      l_ҽ��id.Extend;
      l_ҽ��id(l_ҽ��id.Count) := Substr(c_ҽ��ids, 1, Instr(c_ҽ��ids, ',', 3950) - 1);
      c_ҽ��ids := Substr(c_ҽ��ids, Instr(c_ҽ��ids, ',', 3950) + 1);
    End If;
  End Loop;

  If l_ҽ��id.Count = 0 Then
    Json_Out := '{"output":{"code":0,"message":"δ����ҽ��id�����飡"}}';
    Return;
  End If;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","advice_list":[';

  v_Temp      := '';
  n_Firstitem := 1;
  For I In 1 .. l_ҽ��id.Count Loop
    For r_ҽ�� In (Select /*+cardinality(j,10)*/
                  a.Id, b.ҽ������
                 From ����ҽ����¼ A, ����ҽ����¼ B, Table(f_Num2list(l_ҽ��id(I))) J
                 Where a.���id = b.���id And a.Id <> b.Id And a.Id = j.Column_Value) Loop
    
      If Nvl(n_Firstitem, 0) = 0 Then
        v_Temp := v_Temp || ',';
      Else
        n_Firstitem := 0;
      End If;
      v_Temp := v_Temp || '{';
      v_Temp := v_Temp || '"advice_id":' || Nvl(r_ҽ��.Id, 0);
      v_Temp := v_Temp || ',"advice_note":"' || Zljsonstr(r_ҽ��.ҽ������) || '"';
      v_Temp := v_Temp || '}';
    
      If Length(v_Temp) > 20000 Then
        Json_Out := Json_Out || v_Temp;
        v_Temp   := '';
      End If;
    End Loop;
  End Loop;

  Json_Out := Json_Out || v_Temp || ']}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getgroupadviceinfo;
/
Create Or Replace Procedure Zl_Cissvr_Getinfectinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ���ѯ���˵����Խ����������Ϣ
  --��Σ�Json_In:��ʽ
  --input
  --    query_type         N    1  �������� ��1-ͨ������id+��ҳid��ѯ   2-ͨ������id+�Һŵ���ѯ  3-ͨ���Ǽǿ��Ҳ�ѯ   4-ͨ������id����ѯ
  --    pati_id            N    1  ����id
  --    create_dept_id     N    1  �Ǽǿ���ID
  --    pati_pageid        N    0  ��ҳid
  --    reg_no             C    0  �Һŵ�
  --    create_time_begin  C    0  �Ǽǿ�ʼʱ��
  --    create_time_end    C    0  �Ǽǿ�ʼʱ��
  --    rec_ids            C    0  ����id��

  --���Σ�Json_Out:��ʽ
  --output
  --    code                N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    disease_list         ���Խ���������б�֧�ֶ����[����]
  --       rec_id                    N   1  ����id
  --       pati_source              C   1  ��Դ
  --       pati_id                  C   1  ����id
  --       pati_name                C   1  ��������
  --       pati_sex                 C   1  �����Ա�
  --       pati_age                 C   1  ��������
  --       pati_dept_name           C   1  ���˿���
  --       inpatient_num            C   1  סԺ��
  --       outpatient_num           C   1  �����
  --       spcm_send_time           C   1  �걾�ͼ�ʱ��
  --       spcm_send_dr             C   1  �ͼ�ҽ��
  --       spcm_send_dept           C   1  �ͼ����
  --       spcm_send_deptid         N   1  �ͼ����ID

  --       spcm_rec_status          N   1  ��¼״̬
  --       create_dept_name         C   1  �Ǽǿ���
  --       spcm_name                C   1  �걾����
  --       send_content             C   1  ��������
  --       infctdz_name             C   1  ���Ƽ���
  --       create_dr                C   1  �Ǽ�ҽ��
  --       create_time              C   1  �Ǽ�ʱ��
  --       spcm_procor              C   1  ���鴦����
  --       spcm_proctime            C   1  ���鴦��ʱ��
  --       spcm_procdesc            C   1  ���鴦��˵��

  --       pati_pageid              N   1  ��ҳid
  --       reg_no                   C   1  �Һŵ�
  --       advice_id                N   1  ҽ��ID
  --       eqpmtn_exetime           C   1  ���ʱ��
  --       create_dept_id           N   1  �Ǽǿ���ID
  --       clinic_type              C   1  �������
  --       advice_doctor            C   1  ����ҽ��
  ---------------------------------------------------------------------------
  j_Json   Pljson;
  j_In     Pljson;
  n_Type   Number; --�������� ��1-ͨ��id��ѯ,2-ͨ���Һŵ���ѯ,3-ͨ������id����ҳid��ѯ
  v_�Һŵ� Varchar2(20);
  n_����id Number(18);
  n_��ҳid Number(18);
  v_List   Varchar2(32765);

  d_�Ǽǿ�ʼʱ�� Date;
  d_�Ǽǽ���ʱ�� Date;

  n_�Ǽǿ���id Number;
  v_����ids    Varchar2(4000);
Begin
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');
  n_Type := Nvl(j_Json.Get_Number('query_type'), 0);

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","disease_list":[';

  If n_Type = 1 Then
    n_����id     := Nvl(j_Json.Get_Number('pati_id'), 0);
    n_�Ǽǿ���id := Nvl(j_Json.Get_Number('create_dept_id'), 0);
    n_��ҳid     := Nvl(j_Json.Get_Number('pati_pageid'), 0);
  
    For c_��Ⱦ�� In (Select a.Id, 'סԺ' As ��Դ, c.����id, c.����, c.�Ա�, c.����, e.���� As ����, c.סԺ�� As ��ʶ��, a.�ͼ�ʱ��, a.�ͼ�ҽ��, a.�ͼ����id,
                         g.���� As �ͼ����, a.��¼״̬, f.���� As �Ǽǿ���, a.�걾����, a.�������, a.��Ⱦ������ As ���Ƽ���, a.�Ǽ���, a.�Ǽ�ʱ��, a.������,
                         a.����ʱ��, a.�������˵��, a.��ҳid, a.�Һŵ�, a.ҽ��id, a.���ʱ��, a.�Ǽǿ���id, m.�������, m.����ҽ��


                  
                  From �������Լ�¼ A, ����ҽ����¼ M, ������ҳ C, ���ű� E, ���ű� F, ���ű� G
                  Where a.����id = c.����id And a.��ҳid = c.��ҳid And c.����id = n_����id And c.��ҳid = n_��ҳid And
                        a.�Ǽǿ���id = f.Id(+) And c.��Ժ����id = e.Id(+) And a.�ͼ����id = g.Id(+) And a.ҽ��id = m.Id(+) And
                        (n_�Ǽǿ���id = 0 Or (a.�Ǽǿ���id = n_�Ǽǿ���id))
                  Order By a.�Ǽ�ʱ�� Desc) Loop
      Zljsonputvalue(v_List, 'rec_id', c_��Ⱦ��.Id, 1, 1);
      Zljsonputvalue(v_List, 'pati_source', c_��Ⱦ��.��Դ, 0);
      Zljsonputvalue(v_List, 'pati_id', c_��Ⱦ��.����id, 1);
      Zljsonputvalue(v_List, 'pati_name', c_��Ⱦ��.����, 0);
      Zljsonputvalue(v_List, 'pati_sex', c_��Ⱦ��.�Ա�, 0);
      Zljsonputvalue(v_List, 'pati_age', c_��Ⱦ��.����, 0);
      Zljsonputvalue(v_List, 'pati_dept_name', c_��Ⱦ��.����, 0);
      Zljsonputvalue(v_List, 'inpatient_num', c_��Ⱦ��.��ʶ��, 0);
      Zljsonputvalue(v_List, 'spcm_send_time', To_Char(c_��Ⱦ��.�ͼ�ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'spcm_send_dr', c_��Ⱦ��.�ͼ�ҽ��, 0);
      Zljsonputvalue(v_List, 'spcm_rec_status', c_��Ⱦ��.��¼״̬, 1);
      Zljsonputvalue(v_List, 'create_dept_name', c_��Ⱦ��.�Ǽǿ���, 0);
      Zljsonputvalue(v_List, 'spcm_name', c_��Ⱦ��.�걾����, 0);
      Zljsonputvalue(v_List, 'send_content', c_��Ⱦ��.�������, 0);
      Zljsonputvalue(v_List, 'infctdz_name', c_��Ⱦ��.���Ƽ���, 0);
      Zljsonputvalue(v_List, 'create_dr', c_��Ⱦ��.�Ǽ���, 0);
      Zljsonputvalue(v_List, 'create_time', To_Char(c_��Ⱦ��.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'spcm_procor', c_��Ⱦ��.������, 0);
      Zljsonputvalue(v_List, 'spcm_proctime', To_Char(c_��Ⱦ��.����ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'spcm_procdesc', c_��Ⱦ��.�������˵��, 0);
    
      Zljsonputvalue(v_List, 'pati_pageid', c_��Ⱦ��.��ҳid, 1);
      Zljsonputvalue(v_List, 'reg_no', c_��Ⱦ��.�Һŵ�, 0);
      Zljsonputvalue(v_List, 'advice_id', c_��Ⱦ��.ҽ��id, 1);
      Zljsonputvalue(v_List, 'eqpmtn_exetime', To_Char(c_��Ⱦ��.���ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'create_dept_id', c_��Ⱦ��.�Ǽǿ���id, 1);
      Zljsonputvalue(v_List, 'clinic_type', c_��Ⱦ��.�������, 0);
      Zljsonputvalue(v_List, 'advice_doctor', c_��Ⱦ��.����ҽ��, 0);
    
      Zljsonputvalue(v_List, 'spcm_send_dept', c_��Ⱦ��.�ͼ����, 0);
      Zljsonputvalue(v_List, 'spcm_send_deptid', c_��Ⱦ��.�ͼ����id, 1, 2);
      If Length(v_List) > 20000 Then
        Json_Out := Json_Out || v_List;
        v_List   := ',';
      End If;
    
    End Loop;
  Elsif n_Type = 2 Then
    n_����id     := Nvl(j_Json.Get_Number('pati_id'), 0);
    n_�Ǽǿ���id := Nvl(j_Json.Get_Number('create_dept_id'), 0);
  
    v_�Һŵ� := j_Json.Get_String('reg_no');
  
    For c_��Ⱦ�� In (Select a.Id, '����' As ��Դ, b.����id, b.����, b.�Ա�, b.����, e.���� As ����, b.����� As ��ʶ��, a.�ͼ�ʱ��, a.�ͼ�ҽ��, a.�ͼ����id,
                         g.���� As �ͼ����, a.��¼״̬, f.���� As �Ǽǿ���, a.�걾����, a.�������, a.��Ⱦ������ As ���Ƽ���, a.�Ǽ���, a.�Ǽ�ʱ��, a.������,
                         a.����ʱ��, a.�������˵��, a.��ҳid, a.�Һŵ�, a.ҽ��id, a.���ʱ��, a.�Ǽǿ���id, m.�������, m.����ҽ��


                  
                  From �������Լ�¼ A, ���˹Һż�¼ B, ���ű� E, ���ű� F, ����ҽ����¼ M, ���ű� G
                  Where a.����id = b.����id And a.�Һŵ� = b.No And b.����id = n_����id And b.No = v_�Һŵ� And a.�Ǽǿ���id = f.Id(+) And
                        a.�ͼ����id = g.Id(+) And b.ִ�в���id = e.Id(+) And a.ҽ��id = m.Id(+) And
                        (n_�Ǽǿ���id = 0 Or (a.�Ǽǿ���id = n_�Ǽǿ���id))
                  Order By a.�Ǽ�ʱ�� Desc) Loop
      Zljsonputvalue(v_List, 'rec_id', c_��Ⱦ��.Id, 1, 1);
      Zljsonputvalue(v_List, 'pati_source', c_��Ⱦ��.��Դ, 0);
      Zljsonputvalue(v_List, 'pati_id', c_��Ⱦ��.����id, 1);
      Zljsonputvalue(v_List, 'pati_name', c_��Ⱦ��.����, 0);
      Zljsonputvalue(v_List, 'pati_sex', c_��Ⱦ��.�Ա�, 0);
      Zljsonputvalue(v_List, 'pati_age', c_��Ⱦ��.����, 0);
      Zljsonputvalue(v_List, 'pati_dept_name', c_��Ⱦ��.����, 0);
      Zljsonputvalue(v_List, 'outpatient_num', c_��Ⱦ��.��ʶ��, 0);
      Zljsonputvalue(v_List, 'spcm_send_time', To_Char(c_��Ⱦ��.�ͼ�ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'spcm_send_dr', c_��Ⱦ��.�ͼ�ҽ��, 0);
      Zljsonputvalue(v_List, 'spcm_rec_status', c_��Ⱦ��.��¼״̬, 1);
      Zljsonputvalue(v_List, 'create_dept_name', c_��Ⱦ��.�Ǽǿ���, 0);
      Zljsonputvalue(v_List, 'spcm_name', c_��Ⱦ��.�걾����, 0);
      Zljsonputvalue(v_List, 'send_content', c_��Ⱦ��.�������, 0);
      Zljsonputvalue(v_List, 'infctdz_name', c_��Ⱦ��.���Ƽ���, 0);
      Zljsonputvalue(v_List, 'create_dr', c_��Ⱦ��.�Ǽ���, 0);
      Zljsonputvalue(v_List, 'create_time', To_Char(c_��Ⱦ��.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'spcm_procor', c_��Ⱦ��.������, 0);
      Zljsonputvalue(v_List, 'spcm_proctime', To_Char(c_��Ⱦ��.����ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'spcm_procdesc', c_��Ⱦ��.�������˵��, 0);
    
      Zljsonputvalue(v_List, 'pati_pageid', c_��Ⱦ��.��ҳid, 1);
      Zljsonputvalue(v_List, 'reg_no', c_��Ⱦ��.�Һŵ�, 0);
      Zljsonputvalue(v_List, 'advice_id', c_��Ⱦ��.ҽ��id, 1);
      Zljsonputvalue(v_List, 'eqpmtn_exetime', To_Char(c_��Ⱦ��.���ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'create_dept_id', c_��Ⱦ��.�Ǽǿ���id, 1);
      Zljsonputvalue(v_List, 'clinic_type', c_��Ⱦ��.�������, 0);
      Zljsonputvalue(v_List, 'advice_doctor', c_��Ⱦ��.����ҽ��, 0);
    
      Zljsonputvalue(v_List, 'spcm_send_dept', c_��Ⱦ��.�ͼ����, 0);
      Zljsonputvalue(v_List, 'spcm_send_deptid', c_��Ⱦ��.�ͼ����id, 1, 2);
      If Length(v_List) > 20000 Then
        Json_Out := Json_Out || v_List;
        v_List   := ',';
      End If;
    End Loop;
  Elsif n_Type = 3 Then
    n_�Ǽǿ���id   := Nvl(j_Json.Get_Number('create_dept_id'), 0);
    d_�Ǽǿ�ʼʱ�� := To_Date(j_Json.Get_String('create_time_begin'), 'yyyy-mm-dd hh24:mi:ss');
    d_�Ǽǽ���ʱ�� := To_Date(j_Json.Get_String('create_time_end'), 'yyyy-mm-dd hh24:mi:ss');
    For c_��Ⱦ�� In (Select a.Id, a.��Դ, a.����id, a.����, a.�Ա�, a.����, e.���� As ����, a.��ʶ��, a.�ͼ�ʱ��, a.�ͼ�ҽ��, a.�ͼ����id, a.��¼״̬,
                         f.���� As �ͼ����, a.�걾����, a.�������, a.��Ⱦ������ As ���Ƽ���, a.�Ǽ���, a.�Ǽ�ʱ��, a.������, a.����ʱ��, a.�������˵��, a.��ҳid,
                         a.�Һŵ�, a.ҽ��id, a.���ʱ��, a.�Ǽǿ���id, g.���� As �Ǽǿ���, a.�������, a.����ҽ��
                  From (Select a.Id, '����' As ��Դ, a.����id, b.����, b.�Ա�, b.����, b.����� As ��ʶ��, a.�ͼ�ʱ��, a.�ͼ�ҽ��, a.�ͼ����id, a.��¼״̬,
                                a.�걾����, a.�������, a.��Ⱦ������, a.�Ǽ���, a.�Ǽ�ʱ��, a.������, a.����ʱ��, a.�������˵��, b.ִ�в���id As ����id, a.��ҳid,
                                a.�Һŵ�, a.ҽ��id, a.���ʱ��, a.�Ǽǿ���id, m.�������, m.����ҽ��
                         From �������Լ�¼ A, ���˹Һż�¼ B, ����ҽ����¼ M
                         Where a.����id = b.����id And a.�Һŵ� = b.No And a.ҽ��id = m.Id(+) And a.�Ǽǿ���id = n_�Ǽǿ���id And
                               a.�Ǽ�ʱ�� Between d_�Ǽǿ�ʼʱ�� And d_�Ǽǽ���ʱ��
                         Union All
                         Select a.Id, 'סԺ' As ��Դ, a.����id, c.����, c.�Ա�, c.����, c.סԺ�� As ��ʶ��, a.�ͼ�ʱ��, a.�ͼ�ҽ��, a.�ͼ����id, a.��¼״̬,
                                a.�걾����, a.�������, a.��Ⱦ������, a.�Ǽ���, a.�Ǽ�ʱ��, a.������, a.����ʱ��, a.�������˵��, c.��Ժ����id As ����id, a.��ҳid,
                                a.�Һŵ�, a.ҽ��id, a.���ʱ��, a.�Ǽǿ���id, m.�������, m.����ҽ��
                         From �������Լ�¼ A, ������ҳ C, ����ҽ����¼ M
                         Where a.����id = c.����id And a.��ҳid = c.��ҳid And a.ҽ��id = m.Id(+) And a.�Ǽǿ���id = n_�Ǽǿ���id And
                               a.�Ǽ�ʱ�� Between d_�Ǽǿ�ʼʱ�� And d_�Ǽǽ���ʱ��) A, ���ű� E, ���ű� F, ���ű� G
                  Where a.�ͼ����id = f.Id(+) And a.����id = e.Id(+) And a.�Ǽǿ���id = g.Id(+)
                  Order By a.�Ǽ�ʱ�� Desc) Loop
      Zljsonputvalue(v_List, 'rec_id', c_��Ⱦ��.Id, 1, 1);
      Zljsonputvalue(v_List, 'pati_source', c_��Ⱦ��.��Դ, 0);
      Zljsonputvalue(v_List, 'pati_id', c_��Ⱦ��.����id, 1);
      Zljsonputvalue(v_List, 'pati_name', c_��Ⱦ��.����, 0);
      Zljsonputvalue(v_List, 'pati_sex', c_��Ⱦ��.�Ա�, 0);
      Zljsonputvalue(v_List, 'pati_age', c_��Ⱦ��.����, 0);
      Zljsonputvalue(v_List, 'pati_dept_name', c_��Ⱦ��.����, 0);
      If c_��Ⱦ��.��Դ = '����' Then
        Zljsonputvalue(v_List, 'outpatient_num', c_��Ⱦ��.��ʶ��, 0);
      Else
        Zljsonputvalue(v_List, 'inpatient_num', c_��Ⱦ��.��ʶ��, 0);
      End If;
      Zljsonputvalue(v_List, 'spcm_send_time', To_Char(c_��Ⱦ��.�ͼ�ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'spcm_send_dr', c_��Ⱦ��.�ͼ�ҽ��, 0);
      Zljsonputvalue(v_List, 'spcm_rec_status', c_��Ⱦ��.��¼״̬, 1);
      Zljsonputvalue(v_List, 'create_dept_name', c_��Ⱦ��.�Ǽǿ���, 0);
      Zljsonputvalue(v_List, 'spcm_name', c_��Ⱦ��.�걾����, 0);
      Zljsonputvalue(v_List, 'send_content', c_��Ⱦ��.�������, 0);
      Zljsonputvalue(v_List, 'infctdz_name', c_��Ⱦ��.���Ƽ���, 0);
      Zljsonputvalue(v_List, 'create_dr', c_��Ⱦ��.�Ǽ���, 0);
      Zljsonputvalue(v_List, 'create_time', To_Char(c_��Ⱦ��.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'spcm_procor', c_��Ⱦ��.������, 0);
      Zljsonputvalue(v_List, 'spcm_proctime', To_Char(c_��Ⱦ��.����ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'spcm_procdesc', c_��Ⱦ��.�������˵��, 0);
    
      Zljsonputvalue(v_List, 'pati_pageid', c_��Ⱦ��.��ҳid, 1);
      Zljsonputvalue(v_List, 'reg_no', c_��Ⱦ��.�Һŵ�, 0);
      Zljsonputvalue(v_List, 'advice_id', c_��Ⱦ��.ҽ��id, 1);
      Zljsonputvalue(v_List, 'eqpmtn_exetime', To_Char(c_��Ⱦ��.���ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'create_dept_id', c_��Ⱦ��.�Ǽǿ���id, 1);
      Zljsonputvalue(v_List, 'clinic_type', c_��Ⱦ��.�������, 0);
      Zljsonputvalue(v_List, 'advice_doctor', c_��Ⱦ��.����ҽ��, 0);
    
      Zljsonputvalue(v_List, 'spcm_send_dept', c_��Ⱦ��.�ͼ����, 0);
      Zljsonputvalue(v_List, 'spcm_send_deptid', c_��Ⱦ��.�ͼ����id, 1, 2);
    
      If Length(v_List) > 20000 Then
        Json_Out := Json_Out || v_List;
        v_List   := ',';
      End If;
    
    End Loop;
  
  Elsif n_Type = 4 Then
    v_����ids := j_Json.Get_String('rec_ids');
    For c_��Ⱦ�� In (Select a.Id, a.��Դ, a.����id, a.����, a.�Ա�, a.����, e.���� As ����, a.��ʶ��, a.�ͼ�ʱ��, a.�ͼ�ҽ��, a.�ͼ����id, a.��¼״̬,
                         f.���� As �ͼ����, a.�걾����, a.�������, a.��Ⱦ������ As ���Ƽ���, a.�Ǽ���, a.�Ǽ�ʱ��, a.������, a.����ʱ��, a.�������˵��, a.��ҳid,
                         a.�Һŵ�, a.ҽ��id, a.���ʱ��, a.�Ǽǿ���id, g.���� As �Ǽǿ���, a.�������, a.����ҽ��
                  From (Select a.Id, '����' As ��Դ, a.����id, b.����, b.�Ա�, b.����, b.����� As ��ʶ��, a.�ͼ�ʱ��, a.�ͼ�ҽ��, a.�ͼ����id, a.��¼״̬,
                                a.�걾����, a.�������, a.��Ⱦ������, a.�Ǽ���, a.�Ǽ�ʱ��, a.������, a.����ʱ��, a.�������˵��, b.ִ�в���id As ����id, a.��ҳid,
                                a.�Һŵ�, a.ҽ��id, a.���ʱ��, a.�Ǽǿ���id, m.�������, m.����ҽ��
                         From �������Լ�¼ A, ���˹Һż�¼ B, ����ҽ����¼ M
                         
                         Where a.����id = b.����id And a.�Һŵ� = b.No And a.ҽ��id = m.Id(+) And
                               a.Id In
                               (Select Column_Value As ����id From Table(Cast(f_Str2list(v_����ids) As Zltools.t_Strlist)))
                         Union All
                         Select a.Id, 'סԺ' As ��Դ, a.����id, c.����, c.�Ա�, c.����, c.סԺ�� As ��ʶ��, a.�ͼ�ʱ��, a.�ͼ�ҽ��, a.�ͼ����id, a.��¼״̬,
                                a.�걾����, a.�������, a.��Ⱦ������, a.�Ǽ���, a.�Ǽ�ʱ��, a.������, a.����ʱ��, a.�������˵��, c.��Ժ����id As ����id, a.��ҳid,
                                a.�Һŵ�, a.ҽ��id, a.���ʱ��, a.�Ǽǿ���id, m.�������, m.����ҽ��
                         From �������Լ�¼ A, ������ҳ C, ����ҽ����¼ M
                         
                         Where a.����id = c.����id And a.��ҳid = c.��ҳid And a.ҽ��id = m.Id(+) And
                               a.Id In
                               (Select Column_Value As ����id From Table(Cast(f_Str2list(v_����ids) As Zltools.t_Strlist)))) A,
                       ���ű� E, ���ű� F, ���ű� G
                  Where a.�ͼ����id = f.Id(+) And a.����id = e.Id(+) And a.�Ǽǿ���id = g.Id(+)
                  Order By a.�Ǽ�ʱ�� Desc) Loop
    
      Zljsonputvalue(v_List, 'rec_id', c_��Ⱦ��.Id, 1, 1);
      Zljsonputvalue(v_List, 'pati_source', c_��Ⱦ��.��Դ, 0);
      Zljsonputvalue(v_List, 'pati_id', c_��Ⱦ��.����id, 1);
      Zljsonputvalue(v_List, 'pati_name', c_��Ⱦ��.����, 0);
      Zljsonputvalue(v_List, 'pati_sex', c_��Ⱦ��.�Ա�, 0);
      Zljsonputvalue(v_List, 'pati_age', c_��Ⱦ��.����, 0);
      Zljsonputvalue(v_List, 'pati_dept_name', c_��Ⱦ��.����, 0);
      If c_��Ⱦ��.��Դ = '����' Then
        Zljsonputvalue(v_List, 'outpatient_num', c_��Ⱦ��.��ʶ��, 0);
      Else
        Zljsonputvalue(v_List, 'inpatient_num', c_��Ⱦ��.��ʶ��, 0);
      End If;
      Zljsonputvalue(v_List, 'spcm_send_time', To_Char(c_��Ⱦ��.�ͼ�ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'spcm_send_dr', c_��Ⱦ��.�ͼ�ҽ��, 0);
      Zljsonputvalue(v_List, 'spcm_rec_status', c_��Ⱦ��.��¼״̬, 1);
      Zljsonputvalue(v_List, 'create_dept_name', c_��Ⱦ��.�Ǽǿ���, 0);
      Zljsonputvalue(v_List, 'spcm_name', c_��Ⱦ��.�걾����, 0);
      Zljsonputvalue(v_List, 'send_content', c_��Ⱦ��.�������, 0);
      Zljsonputvalue(v_List, 'infctdz_name', c_��Ⱦ��.���Ƽ���, 0);
      Zljsonputvalue(v_List, 'create_dr', c_��Ⱦ��.�Ǽ���, 0);
      Zljsonputvalue(v_List, 'create_time', To_Char(c_��Ⱦ��.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'spcm_procor', c_��Ⱦ��.������, 0);
      Zljsonputvalue(v_List, 'spcm_proctime', To_Char(c_��Ⱦ��.����ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'spcm_procdesc', c_��Ⱦ��.�������˵��, 0);
    
      Zljsonputvalue(v_List, 'pati_pageid', c_��Ⱦ��.��ҳid, 1);
      Zljsonputvalue(v_List, 'reg_no', c_��Ⱦ��.�Һŵ�, 0);
      Zljsonputvalue(v_List, 'advice_id', c_��Ⱦ��.ҽ��id, 1);
      Zljsonputvalue(v_List, 'eqpmtn_exetime', To_Char(c_��Ⱦ��.���ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'create_dept_id', c_��Ⱦ��.�Ǽǿ���id, 1);
      Zljsonputvalue(v_List, 'clinic_type', c_��Ⱦ��.�������, 0);
      Zljsonputvalue(v_List, 'advice_doctor', c_��Ⱦ��.����ҽ��, 0);
    
      Zljsonputvalue(v_List, 'spcm_send_dept', c_��Ⱦ��.�ͼ����, 0);
      Zljsonputvalue(v_List, 'spcm_send_deptid', c_��Ⱦ��.�ͼ����id, 1, 2);
    
      If Length(v_List) > 20000 Then
        Json_Out := Json_Out || v_List;
        v_List   := ',';
      End If;
    
    End Loop;
  End If;

  If v_List <> ',' Then
    Json_Out := Json_Out || v_List || ']}}';
  Else
    Json_Out := Json_Out || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getinfectinfo;
/
Create Or Replace Procedure Zl_Cissvr_Getinfectreport
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ------------------------------------------------
  --���ܣ���ѯ���Է����������ļ�������

  --���      json
  --input
  --    pati_id                     N    1  ����id
  --    infctdz_name                C    1  ���Ƽ���

  --����      json
  --output
  --    code                        N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                     C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    report_list[]        ���������б�֧�ֶ����[����]
  --       report_id                N   1  ����id
  --       rec_id                   N   1  ������ID
  --       create_time              C   1  ����ʱ��
  --       report_name              C   1  ��������
  --       infctdz_name             C   1  ��Ⱦ������
  ------------------------------------------------
  j_Json Pljson;
  j_In   Pljson;
  v_List Varchar2(32767);

  v_���Ƽ��� Varchar2(500);
  n_����id   Number(18);
Begin
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');

  n_����id   := Nvl(j_Json.Get_Number('pati_id'), 0);
  v_���Ƽ��� := j_Json.Get_String('infctdz_name');

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","report_list":[';
  For c_���� In (Select a.Id, b.Id As ������id, a.����ʱ��, a.��������, b.��Ⱦ������
               From ���Ӳ�����¼ A, �������Լ�¼ B
               Where a.Id = b.�ļ�id And a.����id = b.����id And b.����id = n_����id And b.��Ⱦ������ = v_���Ƽ���) Loop
    Zljsonputvalue(v_List, 'report_id', c_����.Id, 1, 1);
    Zljsonputvalue(v_List, 'rec_id', c_����.������id, 1);
    Zljsonputvalue(v_List, 'create_time', To_Char(c_����.����ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 0);
    Zljsonputvalue(v_List, 'report_name', c_����.��������, 0);
    Zljsonputvalue(v_List, 'infctdz_name', c_����.��Ⱦ������, 0, 2);
  
    If Length(v_List) > 20000 Then
      Json_Out := Json_Out || v_List;
      v_List   := ',';
    End If;
  End Loop;

  If v_List <> ',' Then
    Json_Out := Json_Out || v_List || ']}}';
  Else
    Json_Out := Json_Out || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getinfectreport;
/
Create Or Replace Procedure Zl_Cissvr_Getinpatistate
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡ����״̬
  --��Σ�Json_In:��ʽ
  --  input
  --   pati_id                  N  1  ����id
  --   pati_pageid              N  1  ��ҳid
  --   pati_type                N  1  �������� 0-��ͨסԺ���� 1-�������۲��� 2-סԺ���۲���
  --����: Json_Out,��ʽ����
  --  output
  --    code                    N 1  Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                 C 1  Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    pati_state              N 1  ����״̬
  --    out_time                C 1 ��Ժ����
  --    out_type                C 1 ��Ժ��ʽ
  ---------------------------------------------------------------------------
  n_����id   ������ҳ.����id%Type;
  n_��ҳid   ������ҳ.��ҳid%Type;
  n_�������� ������ҳ.��������%Type;
  d_��Ժ���� ������ҳ.��Ժ����%Type;
  v_��Ժ��ʽ ������ҳ.��Ժ��ʽ%Type;
  n_סԺ���� ������ҳ.��������%Type;
  n_State    Number;
  j_Json     Pljson;
  j_In       Pljson;
Begin

  --�������
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  n_��ҳid   := j_Json.Get_Number('pati_pageid');
  n_����id   := j_Json.Get_Number('pati_id');
  n_�������� := j_Json.Get_Number('pati_type');
  Begin
    Select Nvl(״̬, 0) ״̬, ��Ժ����, ��Ժ��ʽ, ��������
    Into n_State, d_��Ժ����, v_��Ժ��ʽ,n_סԺ����
    From ������ҳ
    Where ����id = n_����id And ��ҳid = n_��ҳid And (�������� = n_�������� Or Nvl(n_��������, 0) = 0);
  Exception
    When Others Then
      Null;
  End;
  If Sql%RowCount > 0 Then
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","pati_state":' || Zljsonstr(n_State, 1) || ',"pati_type":' || Zljsonstr(n_סԺ����, 1) || ',"out_time":"' ||
                Zljsonstr(To_Char(d_��Ժ����, 'YYYY-MM-DD HH24:MI:SS'), 0) || '","out_type":"' || Zljsonstr(v_��Ժ��ʽ, 0) ||
                '"}}';
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getinpatistate;
/
Create Or Replace Procedure Zl_Cissvr_Getinsdeptinfor
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:��ȡ�����в��˵�סԺ����
  --��Σ�Json_In:��ʽ
  -- input
  --   opr_fun             N  1 ִ�й��� 0-��ȡ�����в��˵�סԺ���� 1-ͨ������id/����id�������в��˵���Ժ���һ��߲��� 2-����վ��
  --   pati_source         N  1 ������Դ��1-���2-סԺ
  --   nodeno              C  1 վ����
  --   wararea_ids         C  1 ���˲���ids
  --   find_type           N  1 ���ҷ�ʽ 0-�����Ҳ��� 1-����������
  --   all_wararea         N  1 �Ƿ����в���
  --   pati_in             N  1 �Ƿ���Ժ
  --   dept_srvtype        C  1 �������,������ŷָ�,��:1,2,3
  --����: Json_Out,��ʽ����
  --  output
  --    code               C  1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message            C  1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    dept_list[]             �ٴ������б�
  --      dept_id          N  1 ����id
  --      dept_code        C  1 ���ұ���
  --      dept_name        C  1 ��������
  --      dept_spell       C  1 ���Ҽ���
  --      nodeno           C  1 վ��
  ---------------------------------------------------------------------------
  j_Json     Pljson;
  j_In       Pljson;
  n_������Դ Number(1);
  v_վ��     ���ű�.վ��%Type;
  v_����ids  Varchar2(32767);

  n_ִ�й��� Number;
  n_���ҷ�ʽ Number;
  n_���в��� Number;
  n_��Ժ     Number;
  v_������� Varchar(100);
  v_Jtmp     Varchar2(32767);
  c_Jtmp     Clob;

Begin
  --�������
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  n_������Դ := j_Json.Get_Number('pati_source');
  v_վ��     := j_Json.Get_String('nodeno');
  v_����ids  := j_Json.Get_String('wararea_ids');
  n_ִ�й��� := j_Json.Get_Number('opr_fun');
  n_���ҷ�ʽ := j_Json.Get_Number('find_type');
  n_���в��� := j_Json.Get_Number('all_wararea');
  n_��Ժ     := j_Json.Get_Number('pati_in');
  v_������� := Nvl(j_Json.Get_String('dept_srvtype'), '1,2,3');

  If Nvl(n_ִ�й���, 0) = 0 Then
    If Nvl(n_������Դ, 0) = 0 Then
      Json_Out := '{"output":{"code":0,"message":"δ���벡����Դ������!"}}';
      Return;
    End If;
  
    If n_������Դ = 1 Then
      For r_���� In (Select Distinct a.Id, a.����, a.����, a.����
                   From ���ű� A, ��������˵�� B
                   Where a.Id = b.����id And b.�������� = '�ٴ�' And Instr(',' || v_������� || ',', ',' || b.������� || ',') > 0 And
                         Exists (Select 1 From ��λ״����¼ Where ����id Is Not Null And ����id = a.Id) And
                         (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And
                         (a.վ�� = v_վ�� Or a.վ�� Is Null)) Loop
      
        v_Jtmp := v_Jtmp || ',{"dept_id":' || r_����.Id;
        v_Jtmp := v_Jtmp || ',"dept_code":"' || Zljsonstr(r_����.����) || '"';
        v_Jtmp := v_Jtmp || ',"dept_name":"' || Zljsonstr(r_����.����) || '"';
        v_Jtmp := v_Jtmp || ',"dept_spell":"' || Zljsonstr(r_����.����) || '"';
        v_Jtmp := v_Jtmp || '}';
      
        If Length(v_Jtmp) > 30000 Then
          If c_Jtmp Is Null Then
            c_Jtmp := Substr(v_Jtmp, 2);
          Else
            c_Jtmp := c_Jtmp || v_Jtmp;
          End If;
          v_Jtmp := Null;
        End If;
      
      End Loop;
    
    Else
      For r_���� In (Select Distinct a.Id, a.����, a.����, a.����
                   From ���ű� A, ��Ժ���� B, ��������˵�� C
                   Where a.Id = b.����id And a.Id = c.����id And Instr(',' || v_������� || ',', ',' || c.������� || ',') > 0 And
                         (a.Id In (Select Column_Value From Table(f_Str2list(v_����ids))) Or v_����ids Is Null)) Loop
      
        v_Jtmp := v_Jtmp || ',{"dept_id":' || r_����.Id;
        v_Jtmp := v_Jtmp || ',"dept_code":"' || Zljsonstr(r_����.����) || '"';
        v_Jtmp := v_Jtmp || ',"dept_name":"' || Zljsonstr(r_����.����) || '"';
        v_Jtmp := v_Jtmp || ',"dept_spell":"' || Zljsonstr(r_����.����) || '"';
        v_Jtmp := v_Jtmp || '}';
      
        If Length(v_Jtmp) > 30000 Then
          If c_Jtmp Is Null Then
            c_Jtmp := Substr(v_Jtmp, 2);
          Else
            c_Jtmp := c_Jtmp || v_Jtmp;
          End If;
          v_Jtmp := Null;
        End If;
      End Loop;
    End If;
  Elsif Nvl(n_ִ�й���, 0) = 1 Then
    If Nvl(n_���ҷ�ʽ, 0) = 0 Then
      --0-�����Ҳ���
      For r_���� In (Select a.Id, a.����, a.����, a.����
                   From ���ű� A, ��������˵�� B
                   Where a.Id = b.����id And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) And
                         b.�������� = '�ٴ�' And Instr(',' || v_������� || ',', ',' || b.������� || ',') > 0 And Exists
                    (Select 1
                          From ��λ״����¼
                          Where ����id = a.Id And
                                (Nvl(n_���в���, 0) = 1 Or ����id In (Select Column_Value From Table(f_Str2list(v_����ids)))) And
                                (Nvl(n_��Ժ, 0) = 0 Or ����id Is Not Null)) And (a.վ�� = v_վ�� Or a.վ�� Is Null)) Loop
        v_Jtmp := v_Jtmp || ',{"dept_id":' || r_����.Id;
        v_Jtmp := v_Jtmp || ',"dept_code":"' || Zljsonstr(r_����.����) || '"';
        v_Jtmp := v_Jtmp || ',"dept_name":"' || Zljsonstr(r_����.����) || '"';
        v_Jtmp := v_Jtmp || ',"dept_spell":"' || Zljsonstr(r_����.����) || '"';
        v_Jtmp := v_Jtmp || '}';
      
        If Length(v_Jtmp) > 30000 Then
          If c_Jtmp Is Null Then
            c_Jtmp := Substr(v_Jtmp, 2);
          Else
            c_Jtmp := c_Jtmp || v_Jtmp;
          End If;
          v_Jtmp := Null;
        End If;
      End Loop;
    
    Else
      --1-����������
      For r_���� In (Select a.Id, a.����, a.����, a.����
                   From ���ű� A, ��������˵�� B
                   Where a.Id = b.����id And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) And
                         b.�������� = '����' And Instr(',' || v_������� || ',', ',' || b.������� || ',') > 0 And Exists
                    (Select 1
                          From ��λ״����¼
                          Where ����id = a.Id And
                                (Nvl(n_���в���, 0) = 1 Or ����id In (Select Column_Value From Table(f_Str2list(v_����ids)))) And
                                (Nvl(n_��Ժ, 0) = 0 Or ����id Is Not Null)) And (a.վ�� = v_վ�� Or a.վ�� Is Null)
                   Order By a.����) Loop
        v_Jtmp := v_Jtmp || ',{"dept_id":' || r_����.Id;
        v_Jtmp := v_Jtmp || ',"dept_code":"' || Zljsonstr(r_����.����) || '"';
        v_Jtmp := v_Jtmp || ',"dept_name":"' || Zljsonstr(r_����.����) || '"';
        v_Jtmp := v_Jtmp || ',"dept_spell":"' || Zljsonstr(r_����.����) || '"';
        v_Jtmp := v_Jtmp || '}';
      
        If Length(v_Jtmp) > 30000 Then
          If c_Jtmp Is Null Then
            c_Jtmp := Substr(v_Jtmp, 2);
          Else
            c_Jtmp := c_Jtmp || v_Jtmp;
          End If;
          v_Jtmp := Null;
        End If;
      End Loop;
    
    End If;
  Elsif Nvl(n_ִ�й���, 0) = 2 Then
    For r_���� In (Select Distinct a.վ��, c.����
                 From ���ű� A, ��������˵�� B, Zlnodelist C
                 Where a.Id = b.����id And a.վ�� = c.��� And
                       (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) And
                       ((b.�������� = '�ٴ�' And Nvl(n_���ҷ�ʽ, 0) = 0) Or (b.�������� = '����' And Nvl(n_���ҷ�ʽ, 0) = 1)) And
                       Instr(',' || v_������� || ',', ',' || b.������� || ',') > 0 And
                       ((ID In (Select Distinct Decode(n_���ҷ�ʽ, 0, ����id, 1, ����id)
                                From ��λ״����¼
                                Where ����id Is Not Null) And Nvl(n_��Ժ, 0) = 1) Or Nvl(n_��Ժ, 0) = 0) And
                       ((a.Id In (Select Column_Value From Table(f_Str2list(v_����ids))) And Nvl(n_���в���, 0) = 1) Or
                       Nvl(n_���в���, 0) = 0)
                 Order By a.վ��) Loop
    
      v_Jtmp := v_Jtmp || ',{"dept_name":"' || Zljsonstr(r_����.����) || '"';
      v_Jtmp := v_Jtmp || ',"nodeno":"' || Zljsonstr(r_����.վ��) || '"';
      v_Jtmp := v_Jtmp || '}';
    
      If Length(v_Jtmp) > 30000 Then
        If c_Jtmp Is Null Then
          c_Jtmp := Substr(v_Jtmp, 2);
        Else
          c_Jtmp := c_Jtmp || v_Jtmp;
        End If;
        v_Jtmp := Null;
      End If;
    
    End Loop;
  
  End If;

  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","dept_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","dept_list":[' || c_Jtmp || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getinsdeptinfor;
/
Create Or Replace Procedure Zl_Cissvr_Getmaxbedlen
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:��ȡ�������ŵ�����ų���
  --��Σ�Json_In:��ʽ
  -- input            ����id�Ϳ���ID�ڵ����ֻ�ܴ�һ��
  --   wardarea_id    N  0  ����id
  --   dept_id        N  0 ����ID
  --����: Json_Out,��ʽ����
  --  output
  --    code               C  1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message            C  1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    max_len            N  1 ��󳤶�
  ---------------------------------------------------------------------------
  j_Json   Pljson;
  j_In     Pljson;
  n_����id ��λ״����¼.����id%Type;
  n_����id ��λ״����¼.����id%Type;
  n_����   Number(20);

  n_���벡��id Number(1);
  n_�������id Number(1);

Begin
  --�������
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');
  If j_Json.Exist('wardarea_id') Then
    n_����id     := j_Json.Get_Number('wardarea_id');
    n_���벡��id := 1;
  End If;

  If j_Json.Exist('dept_id') Then
    n_����id     := j_Json.Get_Number('dept_id');
    n_�������id := 1;
  End If;

  If Nvl(n_���벡��id, 0) = 0 And Nvl(n_�������id, 0) = 0 Then
    Json_Out := Zljsonout('����id�Ϳ���id���봫��һ��������');
    Return;
  End If;

  If Nvl(n_���벡��id, 0) = 1 Then
    If Nvl(n_����id, 0) = 0 Then
      Select Max(Length(����)) Into n_���� From ��λ״����¼ Where ״̬ = 'ռ��' And ����id Is Not Null;
    Else
      Select Max(Length(����)) Into n_���� From ��λ״����¼ Where ״̬ = 'ռ��' And ����id = n_����id;
    End If;
  Else
    If Nvl(n_����id, 0) = 0 Then
      Select Max(Length(����)) Into n_���� From ��λ״����¼ Where ״̬ = 'ռ��' And ����id Is Not Null;
    Else
      Select Max(Length(����)) Into n_���� From ��λ״����¼ Where ״̬ = 'ռ��' And ����id = n_����id;
    End If;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","max_len":' || Nvl(n_����, 0) || '}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getmaxbedlen;
/
Create Or Replace Procedure Zl_Cissvr_Getmedicalgroupid
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:���ݲ��˱䶯��¼����ȡ��Ӧ��ҽ��С��ID
  --��Σ�Json_In:��ʽ
  --    input
  --      pati_id           N 1 ����id
  --      pati_pageid       N 1 ��ҳid
  --      plcdept_id        N 1 ��������ID
  --      placer            C 1 ������
  --      occur_time        C 1 ����ʱ��:yyyy-mm-dd hh24:mi:ss
  --����: Json_Out,��ʽ����
  --  output
  --    code                  N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message               C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    group_id              N 0 ҽ��С��id
  ---------------------------------------------------------------------------
  j_Json   Pljson;
  j_In     Pljson;
  n_����id ���˱䶯��¼.����id%Type;
  n_��ҳid ���˱䶯��¼.��ҳid%Type;

  n_��������id ���˱䶯��¼.����id%Type;
  v_������     ��Ա��.����%Type;
  d_����ʱ��   ���˱䶯��¼.��ʼʱ��%Type;

  n_��id ���˱䶯��¼.ҽ��С��id%Type;
Begin
  --�������
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');
  n_��ҳid := j_Json.Get_Number('pati_pageid');

  n_��������id := j_Json.Get_Number('plcdept_id');
  v_������     := j_Json.Get_String('placer');
  d_����ʱ��   := To_Date(j_Json.Get_String('occur_time'), 'yyyy-mm-dd hh24:mi:ss');

  n_��id := Zl_ҽ��С��_Get(n_��������id, v_������, n_����id, n_��ҳid, d_����ʱ��);

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","group_id":' || Nvl(n_��id, 0) || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getmedicalgroupid;
/
Create Or Replace Procedure Zl_Cissvr_Getmedrecscoreresult
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:��ȡ�������ֽ��
  --��Σ�Json_In:��ʽ
  -- input
  --   pati_id              N 1 ����id
  --   pati_pageid          N 1 ��ҳid
  --����: Json_Out,��ʽ����
  --  output
  --    code                N  1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message             C  1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    result              C  1  ���ݽ����ȼ�:���ס�/���ҡ�/������/���񡱣������ϸ�,����ʱ��ȡ��һ����
  ---------------------------------------------------------------------------
  j_Json   Pljson;
  j_In     Pljson;
  n_����id �������ֽ��.����id%Type;
  n_��ҳid �������ֽ��.��ҳid%Type;
  v_�ȼ�   �������ֽ��.�ȼ�%Type;

Begin
  --�������
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');
  n_��ҳid := j_Json.Get_Number('pati_pageid');

  If Nvl(n_����id, 0) = 0 Or Nvl(n_��ҳid, 0) = 0 Then
    Json_Out := Zljsonout('δ���벡��id����ҳid�����飡');
    Return;
  End If;

  Select Max(�ȼ�) Into v_�ȼ� From �������ֽ�� Where ����id = n_����id And ��ҳid = n_��ҳid;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","result":"' || v_�ȼ� || '"}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getmedrecscoreresult;
/
Create Or Replace Procedure Zl_Cissvr_Getnextid
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --���ܣ���ȡָ��������Ӧ������(���淶������������Ϊ��������_id��)����һ��ֵ
  --��Σ�Json_In:��ʽ
  --input
  --  table_name    C  1 ����
  --  col_name      C  1 �ֶ���  �������Ʋ�һ����ID�������¼ID
  -- ����:
  --  output
  --  next_id      N   1  ����
  -------------------------------------------

  v_Table Varchar2(500);
  v_Col   Varchar2(500);

  n_Nextid Number(18);
  j_Json   Pljson;
  j_In     Pljson;
Begin
  --�������
  j_In    := Pljson(Json_In);
  j_Json  := j_In.Get_Pljson('input');
  v_Table := j_Json.Get_String('table_name');
  v_Col   := Nvl(j_Json.Get_String('col_name'), 'ID');

  Execute Immediate 'select ' || v_Table || '_' || Nvl(v_Col, 'ID') || '.nextval from dual'
    Into n_Nextid;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","next_id":' || Nvl(n_Nextid, 0) || '}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Getnextid;
/
Create Or Replace Procedure Zl_Cissvr_Getpatiallergyinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡ���˹�����Ϣ
  --input      ��ȡ���˹�����Ϣ
  --  pati_id               N  1  ����id
  --  visit_id              N  1  ��ʶ�ţ��Һ�id���������ҳid��סԺ��
  --output
  --  code                  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message               C  1  Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  allergy_list[]    ������Ϣ��[����]
  --     drug_name          C  1  ҩ������
  --     allergy_time       C  1  ����ʱ��
  --     allergy_info       C  1  ������Ӧ
  ---------------------------------------------------------------------------
  j_Input  Pljson;
  j_In     Pljson;
  n_����id Number(18);
  n_��ʶid Number(18); --����Һ�id ��סԺ����ҳid

  v_Jtmp Varchar2(32767);

Begin
  j_In     := Pljson(Json_In);
  j_Input  := j_In.Get_Pljson('input');
  n_����id := j_Input.Get_Number('pati_id');
  n_��ʶid := j_Input.Get_Number('visit_id');

  For v_������¼ In (Select Distinct a.ҩ����, Nvl(a.����ʱ��, a.��¼ʱ��) As ����ʱ��, a.������Ӧ
                 From ���˹�����¼ A, ���˹Һż�¼ B, ������ҳ C, ���ű� D, ���ű� E
                 Where a.����id = b.����id(+) And a.��ҳid = b.Id(+) And b.��¼����(+) = 1 And b.��¼״̬(+) = 1 And
                       a.����id = c.����id(+) And a.��ҳid = c.��ҳid(+) And b.ִ�в���id = d.Id(+) And c.��Ժ����id = e.Id(+) And
                       a.��� = 1 And ҩ���� Is Not Null And a.����id = n_����id And a.��ҳid = n_��ʶid And Not Exists
                  (Select ҩ��id
                        From ���˹�����¼
                        Where (Nvl(ҩ��id, 0) = Nvl(a.ҩ��id, 0) Or Nvl(ҩ����, 'Null') = Nvl(a.ҩ����, 'Null')) And Nvl(���, 0) = 0 And
                              ��¼ʱ�� > a.��¼ʱ�� And ����id = n_����id And ��ҳid = n_��ʶid)
                 Order By Nvl(a.����ʱ��, a.��¼ʱ��) Desc) Loop
  
    v_Jtmp := v_Jtmp || ',{"drug_name":"' || Zljsonstr(v_������¼.ҩ����) || '"';
    v_Jtmp := v_Jtmp || ',"allergy_time":"' || To_Char(v_������¼.����ʱ��, 'YYYY-MM-DD HH24:MI') || '"';
  
    v_Jtmp := v_Jtmp || ',"allergy_info":"' || Zljsonstr(v_������¼.������Ӧ) || '"';
    v_Jtmp := v_Jtmp || '}';
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","allergy_list":[' || Substr(v_Jtmp, 2) || ']}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zltools.Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getpatiallergyinfo;
/

Create Or Replace Procedure Zl_Cissvr_Getpatibaseinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡ���˻�����Ϣ��ȡ(�ٴ�)
  --��Σ�Json_In:��ʽ
  --  input
  --    query_type        N 1 ��ѯ��ʽ-- 1-ͨ������ID+��ҳID��ѯ������Ϣ,2-ͨ��ҽ��ID��ȡ���˻�����Ϣ ,3-ͨ���Һŵ���ȡ���˻�����Ϣ


  --    pati_id           N 1 ����id--
  --    page_id           N 1 ��ҳid--
  --    advice_id         N 1 ҽ��ID--
  --    pati_type         N 1 �������ͣ�0-סԺ���� 1-���ﲡ��

  --����: Json_Out,��ʽ����
  --  output
  --    code                 N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message              C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    page_list[]        1  ������
  --       pati_id           N 1 ����id
  --       page_id           N 1 ��ҳid
  --       pati_name         C 1 ��������
  --       pati_sex          C 1 �����Ա�
  --       pati_age          C 1 ��������
  --       dept_name         C 1 ��������
  --       inpatient_num     C 1 סԺ��
  --       pati_bed          C 1 ��ǰ����
  --       dept_id           N 1 ����id
  --       regist_no         C 1 �Һŵ�
  --       registration_time C 1 ����ʱ��
  --       adtd_time         C 1 ��Ժʱ��

  --       pati_content      C 1 ��ǰ����
  --       insurance_type    N 1 ����
  --       pati_wardarea_id  N 1 ��ǰ����ID

  --       reg_id            N 1 �Һ�id
  --       outpatient_num    C 1 �����
  --       return_visit      N 1 �����־
  --       outp_room_name    C 1 ������������
  ---------------------------------------------------------------------------
  j_In       Pljson;
  j_Json     Pljson;
  n_��ѯ��ʽ Number;
  n_����id   Number;
  n_��ҳid   Number;
  n_ҽ��id   Number;
  n_�������� Number;
  v_List     Varchar2(32767);

  v_�Һŵ� Varchar(50);

  v_Err_Msg Varchar(2000);
  Err_Item Exception;
Begin
  --�������
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  n_��ѯ��ʽ := j_Json.Get_Number('query_type');
  n_�������� := j_Json.Get_Number('pati_type');
  Json_Out   := '{"output":{"code":1,"message":"�ɹ�","page_list":[';
  If n_��ѯ��ʽ = 1 Then
    n_����id := j_Json.Get_Number('pati_id');
    n_��ҳid := j_Json.Get_Number('page_id');
    If Nvl(n_��������, 0) = 0 Then
      For R In (Select a.����, a.�Ա�, a.����, b.���� As ����, a.סԺ��, a.��Ժ���� As ����,
                       To_Char(a.��Ժ����, 'yyyy-MM-dd HH24:MI:SS') As ����ʱ��, To_Char(a.��Ժ����, 'yyyy-MM-dd HH24:MI:SS') As ��Ժʱ��,
                       b.Id As ����id, a.��ǰ����, a.����, a.��ǰ����id
                From ������ҳ A, ���ű� B
                Where a.��Ժ����id = b.Id And a.����id = n_����id And a.��ҳid = Decode(n_��ҳid, Null, a.��ҳid, n_��ҳid)
                Order By Nvl(a.��ҳid, 0)) Loop
        Zljsonputvalue(v_List, 'pati_id', n_����id, 1, 1);
        Zljsonputvalue(v_List, 'page_id', n_��ҳid, 1);
        Zljsonputvalue(v_List, 'pati_name', r.����);
        Zljsonputvalue(v_List, 'pati_sex', r.�Ա�);
        Zljsonputvalue(v_List, 'pati_age', r.����);
        Zljsonputvalue(v_List, 'dept_name', r.����);
        Zljsonputvalue(v_List, 'inpatient_num', r.סԺ��, 0);
        Zljsonputvalue(v_List, 'pati_bed', r.����);
        Zljsonputvalue(v_List, 'registration_time', r.����ʱ��);
        Zljsonputvalue(v_List, 'adtd_time', r.��Ժʱ��);
        Zljsonputvalue(v_List, 'dept_id', r.����id, 1);
      
        Zljsonputvalue(v_List, 'pati_content', r.��ǰ����);
        Zljsonputvalue(v_List, 'insurance_type', r.����, 1);
        Zljsonputvalue(v_List, 'pati_wardarea_id', r.��ǰ����id, 1, 2);
      
        If Length(v_List) > 20000 Then
          Json_Out := Json_Out || v_List;
          v_List   := ',';
        End If;
      End Loop;
      If v_List <> ',' Then
        Json_Out := Json_Out || v_List || ']}}';
      Else
        Json_Out := Json_Out || ']}}';
      End If;
    Else
      For R In (Select a.No, a.����id, a.����, a.�Ա�, a.����
                From ���˹Һż�¼ A
                Where a.����id = n_����id And a.Id = n_��ҳid) Loop
        Zljsonputvalue(v_List, 'pati_id', n_����id, 1, 1);
        Zljsonputvalue(v_List, 'pati_name', r.����);
        Zljsonputvalue(v_List, 'pati_sex', r.�Ա�);
        Zljsonputvalue(v_List, 'pati_age', r.����);
        Zljsonputvalue(v_List, 'regist_no', r.No, 0, 2);
        If Length(v_List) > 20000 Then
          Json_Out := Json_Out || v_List;
          v_List   := ',';
        End If;
      End Loop;
      If v_List <> ',' Then
        Json_Out := Json_Out || v_List || ']}}';
      Else
        Json_Out := Json_Out || ']}}';
      End If;
    End If;
  Elsif n_��ѯ��ʽ = 2 Then
    n_ҽ��id := j_Json.Get_Number('advice_id');
    For R In (Select a.����id, a.��ҳid, Nvl(q.Ӥ������, a.����) ����, Nvl(q.Ӥ���Ա�, a.�Ա�) �Ա�,
                     Decode(q.���, Null, a.����, Round(Decode(q.����ʱ��, Null, Sysdate, q.����ʱ��) - q.����ʱ��) || '��') ����,
                     b.���� As ����, Null As סԺ��, Null As ����, To_Char(a.��ʼִ��ʱ��, 'yyyy-MM-dd HH24:MI:SS') As ����ʱ��,
                     b.Id As ����id, Null As ��ǰ����, Null As ����, Null As ��ǰ����id
              From ����ҽ����¼ A, ���ű� B, ������������¼ Q
              Where a.���˿���id = b.Id And a.Id = n_ҽ��id And a.����id = q.����id(+) And a.��ҳid = q.��ҳid(+) And a.Ӥ�� = q.���(+)) Loop
      Zljsonputvalue(v_List, 'pati_id', r.����id, 1, 1);
      Zljsonputvalue(v_List, 'page_id', r.��ҳid, 1);
      Zljsonputvalue(v_List, 'pati_name', r.����);
      Zljsonputvalue(v_List, 'pati_sex', r.�Ա�);
      Zljsonputvalue(v_List, 'pati_age', r.����);
      Zljsonputvalue(v_List, 'dept_name', r.����);
      Zljsonputvalue(v_List, 'inpatient_num', r.סԺ��, 0);
      Zljsonputvalue(v_List, 'pati_bed', r.����);
      Zljsonputvalue(v_List, 'registration_time', r.����ʱ��);
      Zljsonputvalue(v_List, 'dept_id', r.����id, 1);
    
      Zljsonputvalue(v_List, 'pati_content', r.��ǰ����);
      Zljsonputvalue(v_List, 'insurance_type', r.����, 1);
      Zljsonputvalue(v_List, 'pati_wardarea_id', r.��ǰ����id, 1, 2);
    
      If Length(v_List) > 20000 Then
        Json_Out := Json_Out || v_List;
        v_List   := ',';
      End If;
    End Loop;
    If v_List <> ',' Then
      Json_Out := Json_Out || v_List || ']}}';
    Else
      Json_Out := Json_Out || ']}}';
    End If;
  Elsif n_��ѯ��ʽ = 3 Then
    v_�Һŵ� := j_Json.Get_String('reg_no');
    n_����id := Nvl(j_Json.Get_Number('pati_id'), 0);
  
    For R In (Select a.����, a.�Ա�, a.����, b.���� As ����, a.�����, Null As ����, a.����ʱ�� As ����ʱ��, b.Id As ����id, a.����, a.����id, a.Id,
                     Null As ��ǰ����, a.����, Null As ��ǰ����id, a.����
              From ���˹Һż�¼ A, ���ű� B
              Where a.ִ�в���id = b.Id And a.��¼���� = 1 And a.��¼״̬ = 1 And a.No = v_�Һŵ� And (n_����id = 0 Or (n_����id = a.����id))) Loop
      Zljsonputvalue(v_List, 'pati_id', r.����id, 1, 1);
      Zljsonputvalue(v_List, 'reg_id', r.Id, 1);
      Zljsonputvalue(v_List, 'pati_name', r.����);
      Zljsonputvalue(v_List, 'pati_sex', r.�Ա�);
      Zljsonputvalue(v_List, 'pati_age', r.����);
      Zljsonputvalue(v_List, 'dept_name', r.����);
      Zljsonputvalue(v_List, 'outpatient_num', r.�����, 0);
      Zljsonputvalue(v_List, 'pati_bed', r.����);
      Zljsonputvalue(v_List, 'registration_time', To_Date(r.����ʱ��, 'yyyy-MM-dd HH24:MI:SS'));
      Zljsonputvalue(v_List, 'dept_id', r.����id, 1);
      Zljsonputvalue(v_List, 'return_visit', r.����, 1);
      Zljsonputvalue(v_List, 'outp_room_name', r.����);
    
      Zljsonputvalue(v_List, 'pati_content', r.��ǰ����);
      Zljsonputvalue(v_List, 'insurance_type', r.����, 1);
      Zljsonputvalue(v_List, 'pati_wardarea_id', r.��ǰ����id, 1, 2);
      If Length(v_List) > 20000 Then
        Json_Out := Json_Out || v_List;
        v_List   := ',';
      End If;
    End Loop;
    If v_List <> ',' Then
      Json_Out := Json_Out || v_List || ']}}';
    Else
      Json_Out := Json_Out || ']}}';
    End If;
  End If;
Exception
  When Err_Item Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr('-20101:[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]') || '"}}';
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getpatibaseinfo;
/

Create Or Replace Procedure Zl_Cissvr_Getpatichangerec
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:����������ȡ���˱䶯��Ϣ
  --��Σ�Json_In:��ʽ
  -- input
  --   query_type           N 1 ��ѯ��ʽ��0-����ָ��������ѯ���˱䶯��Ϣ��1-����ѯ����ĳ��סԺ��ת����Ϣ
  --   pati_id              N 1 ����id
  --   pati_pageid          N 1 ��ҳid
  --   pati_wararea_id      N 0 ����id
  --   pati_dept_id         N 0 ����id
  --   start_reasons        C 0 ��ʼԭ��s:����ö��ŷ���,��:3,15,10,1
  --   stop_reasons         C 0 ��ֹԭ��s:����ö��ŷ���,��:3,15,10,1
  --����: Json_Out,��ʽ����
  --  output
  --    code               C  1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message            C  1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    change_list[]
  --        start_time         C 1 ��ʼʱ��:yyyy-mm-dd hh24:mi:ss
  --        start_reason       N 1 ��ʼԭ��
  --        dept_name          C 1 ��������
  --        stop_time          C 1 ��ֹʱ��:yyyy-mm-dd hh24:mi:ss
  --        stop_reason        N 1 ��ֹԭ��
  ---------------------------------------------------------------------------
  j_Json      Pljson;
  j_In        Pljson;
  n_����id    ���˱䶯��¼.����id%Type;
  n_��ҳid    ���˱䶯��¼.��ҳid%Type;
  n_����id    ���˱䶯��¼.����id%Type;
  n_����id    ���˱䶯��¼.����id%Type;
  v_��ֹԭ��  Varchar2(3000);
  v_��ʼԭ��  Varchar2(200);
  n_��ѯ��ʽ  Number(2);
  c_Jtmp      Clob; 
  v_Temp      Varchar2(32767);
Begin
  --�������
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  n_��ѯ��ʽ := j_Json.Get_Number('query_type');
  n_����id   := j_Json.Get_Number('pati_id');
  n_��ҳid   := j_Json.Get_Number('pati_pageid');

  n_����id   := j_Json.Get_Number('pati_wararea_id');
  n_����id   := j_Json.Get_Number('pati_dept_id');
  v_��ʼԭ�� := j_Json.Get_String('start_reasons');
  v_��ֹԭ�� := j_Json.Get_String('stop_reasons');

  If Nvl(n_����id, 0) = 0 Or Nvl(n_��ҳid, 0) = 0 Then
    Json_Out := '{"output":{"code":0,"message":"���봫�벡��id����ҳid��"}}';
    Return;
  End If; 
  
  If n_��ѯ��ʽ = 1 Then
    v_Temp := Null;
    For R In (Select Distinct 1 As ��ʼԭ��, To_Date('1900-01-01', 'yyyy-mm-dd') As ��ʼʱ��, b.����
              From ���˱䶯��¼ A, ���ű� B
              Where a.����id = b.Id And a.��ʼʱ�� Is Not Null And a.��ʼԭ�� In (1, 2) And a.����id = n_����id And ��ҳid = n_��ҳid
              Union All
              Select a.��ʼԭ��, a.��ʼʱ��, b.����
              From ���˱䶯��¼ A, ���ű� B
              Where a.����id = b.Id And a.��ʼʱ�� Is Not Null And a.��ʼԭ�� = 3 And a.����id = n_����id And ��ҳid = n_��ҳid
              Order By ��ʼʱ��) Loop
      v_Temp := v_Temp || ',{"start_reason":' || r.��ʼԭ�� || ',"dept_name":"' || Zljsonstr(r.����) || '"}';
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","change_list":[' || Substr(v_Temp, 2) || ']}}';
    Return;
  End If;

  v_Temp := Null;
  For r_�䶯 In (Select a.��ʼԭ��, To_Char(a.��ʼʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��, b.����,
                      To_Char(a.��ֹʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��ֹʱ��, a.��ֹԭ��
               From ���˱䶯��¼ A, ���ű� B
               Where a.����id = b.Id And ((a.��ʼʱ�� Is Not Null And Nvl(v_��ʼԭ��, '-') <> '-') Or Nvl(v_��ʼԭ��, '-') = '-') And
                     (Instr(',' || v_��ʼԭ�� || ',', ',' || a.��ʼԭ�� || ',') > 0 Or Nvl(v_��ʼԭ��, '-') = '-') And
                     a.����id = n_����id And a.��ҳid = n_��ҳid And (a.����id = n_����id Or Nvl(n_����id, 0) = 0) And
                     (a.����id = n_����id Or Nvl(n_����id, 0) = 0) And
                     (Instr(',' || v_��ֹԭ�� || ',', ',' || a.��ֹԭ�� || ',') > 0 Or Nvl(v_��ֹԭ��, '-') = '-')
               Order By ��ʼʱ��, ��ֹʱ��) Loop
  
    v_Temp := v_Temp || ',{"start_time":"' || Zljsonstr(r_�䶯.��ʼʱ��) || '"';
    v_Temp := v_Temp || ',"start_reason":' || Nvl(r_�䶯.��ʼԭ��, 0);
    v_Temp := v_Temp || ',"dept_name":"' || Zljsonstr(r_�䶯.����) || '"';
    v_Temp := v_Temp || ',"stop_time":"' || Zljsonstr(r_�䶯.��ֹʱ��) || '"';
    v_Temp := v_Temp || ',"stop_reason":' || Nvl(r_�䶯.��ֹԭ��, 0);
    v_Temp := v_Temp || '}';
  
    If Length(v_Temp) > 30000 Then
      If c_Jtmp Is Null Then
        c_Jtmp := Substr(v_Temp, 2);
      Else
        c_Jtmp := c_Jtmp || v_Temp;
      End If;
      v_Temp := Null;
    End If;
  
  End Loop;
  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","change_list":[' || Substr(v_Temp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Temp;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","change_list":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getpatichangerec;
/
Create Or Replace Procedure Zl_Cissvr_Getpatidiagnose
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡ���˻�ȡ���������Ϣ
  --input      ��ȡ���������Ϣ
  --  pati_id               N  1  ����ID
  --  visit_id              N  1  ����ID
  --output
  --  code                  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message               C  1  Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  diagnose_list[]      ������ݣ�[����]
  --     diag_origin        N  1  ��¼��Դ
  --     diag_type          N  1  �������
  --     diag_order         N  1 ��ϴ���
  --     diag_description   C  1  �������
  --     diag_distrustful   N  1  �Ƿ�����
  --     diag_record_time   C  1  ��¼����
  ---------------------------------------------------------------------------
  j_Input  Pljson;
  j_In     Pljson;
  n_����id Number(18);
  n_����id Number(18);
  v_Output Varchar2(32767);
Begin
  j_In     := Pljson(Json_In);
  j_Input  := j_In.Get_Pljson('input');
  n_����id := j_Input.Get_Number('pati_id');
  n_����id := j_Input.Get_Number('visit_id');

  For r_Diagnose In (Select ��¼��Դ, �������, ��ϴ���, �������, �Ƿ�����, ��¼����
                     From ������ϼ�¼
                     Where ����id = n_����id And ��ҳid = n_����id And Nvl(¼�����, '01') = '01' And Nvl(�������, 'E') = 'E'
                     Order By ��¼���� Desc, ������� Desc) Loop
  
    Zljsonputvalue(v_Output, 'diag_origin', Nvl(r_Diagnose.��¼��Դ, 0), 1, 1);
    Zljsonputvalue(v_Output, 'diag_type', Nvl(r_Diagnose.�������, 0), 1);
    Zljsonputvalue(v_Output, 'diag_order', r_Diagnose.��ϴ���, 1);
    Zljsonputvalue(v_Output, 'diag_description', Nvl(r_Diagnose.�������, ''));
    Zljsonputvalue(v_Output, 'diag_distrustful', Nvl(r_Diagnose.�Ƿ�����, 0), 1);
    Zljsonputvalue(v_Output, 'diag_record_time', Nvl(r_Diagnose.��¼����, ''), 0, 2);
  End Loop;

  Json_Out := '{"output":{"code":1,"message": "�ɹ�","diagnose_list":[' || v_Output || ']}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getpatidiagnose;
/
Create Or Replace Procedure Zl_Cissvr_Getpatiid
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ݴ��š�סԺ�ʻ�ȡ����ID����ҳID
  --input
  --   wardarea_id          N 1 ��ǰ����id
  --   pati_bed             C 1 ��ǰ����
  --   inpatient_num        C 1 סԺ��
  --   obsv_no              C 1 ���ۺ�
  --output
  --    code                N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message             C 1 Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    pati_id             N 1 ����ID:δ�ҵ�ʱҲ�ɹ�������0
  --    pati_pageid         N   ��ҳID
  ---------------------------------------------------------------------------
  j_In     PLJson;
  j_Input  PLJson;
  n_��ҳid ������ҳ.��ҳid%Type;
  n_����id ������ҳ.����id%Type;

  n_��ǰ����id ������ҳ.��ǰ����id%Type;
  v_��ǰ����   ������ҳ.��Ժ����%Type;

  n_סԺ�� ������ҳ.סԺ��%Type;
  n_���ۺ� ������ҳ.���ۺ�%Type;
Begin
  j_In         := PLJson(Json_In);
  j_Input      := j_In.Get_Pljson('input');
  n_��ǰ����id := j_Input.Get_Number('wardarea_id');
  v_��ǰ����   := j_Input.Get_String('pati_bed');
  n_סԺ��     := To_Number(j_Input.Get_String('inpatient_num'));
  n_���ۺ�     := To_Number(j_Input.Get_String('obsv_no'));

  If Nvl(n_���ۺ�, 0) <> 0 Then
    Select Max(����id), Max(��ҳid) Into n_����id, n_��ҳid From ������ҳ Where ���ۺ� = n_���ۺ�;
  Elsif Nvl(n_סԺ��, 0) <> 0 Then
    Select Max(����id), Max(��ҳid) Into n_����id, n_��ҳid From ������ҳ Where סԺ�� = n_סԺ��;
  Else
    Select Max(a.����id), Max(a.��ҳid)
    Into n_����id, n_��ҳid
    From ��Ժ���� A, ��λ״����¼ B
    Where a.����id = b.����id And b.����id = Nvl(n_��ǰ����id, 0) And b.���� = v_��ǰ����;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","pati_id":' || Nvl(n_����id, 0) || ',"pati_pageid":' || Nvl(n_��ҳid, 0) || '}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || zlJsonStr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getpatiid;
/
Create Or Replace Procedure Zl_Cissvr_Getpatiidbyinpno
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --���ܣ���ȡ�������һ�ιҺŵĹҺż�¼
  --��Σ�Json_In:��ʽ
  --input
  --  inpatient_num         C   1 סԺ��
  --����      json
  --output
  --    code                N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    pati_id             N   1 ����id
  -------------------------------------------
  j_Json   Pljson;
  j_In     Pljson;
  n_����id Number(18);
  n_סԺ�� Number;
Begin
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_סԺ�� := j_Json.Get_String('inpatient_num');
  If Nvl(n_סԺ��, 0) <> 0 Then
    Begin
      Select Nvl(Max(����id), 0) As ����id Into n_����id From ������ҳ Where סԺ�� = n_סԺ��;
    Exception
      When Others Then
        Null;
    End;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","pati_id":' || Nvl(n_����id, 0) || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getpatiidbyinpno;
/
Create Or Replace Procedure Zl_Cissvr_Getpatimaxpageid
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:��ȡָ�������������ҳID
  --��Σ�Json_In:��ʽ
  --    input
  --      pati_id           N 1 ����id
  --      pati_wararea_id   N 1 ��ǰ����iD
  --����: Json_Out,��ʽ����
  --  output
  --    code                  N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message               C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    pati_pageid           N 1 ��ҳid
  ---------------------------------------------------------------------------

  j_Json Pljson;
  j_In   Pljson;

  n_����id ������ҳ.��ǰ����id%Type;
  n_����id ������ҳ.����id%Type;
  n_��ҳid ������ҳ.��ҳid%Type;
Begin
  --�������
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');
  n_����id := j_Json.Get_Number('pati_wararea_id');

  Select Max(��ҳid)
  Into n_��ҳid
  From ������ҳ
  Where ����id = n_����id And Decode(n_����id, Null, 0, ��ǰ����id) = Decode(n_����id, Null, 0, n_����id);

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","pati_pageid":' || Nvl(n_��ҳid, 0) || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getpatimaxpageid;
/

Create Or Replace Procedure Zl_Cissvr_Getpatipageextinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:���ݲ���id����ҳid��ȡ������ҳ�ӱ���Ϣ
  --��Σ�Json_In:��ʽ
  --    input
  --      pati_id           N 1 ����id
  --      pati_pageid       N 1 ��ҳid
  --      info_names        C 1 ��Ϣ��������ö���

  --����: Json_Out,��ʽ����
  --  output
  --    code                  N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message               C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    pati_Name             C 1 ����
  --    ext_list[]             �����ӱ���Ϣ�б�
  --      info_name           C 1 ��Ϣ��
  --      info_value          C 1 ��Ϣֵ
  ---------------------------------------------------------------------------
  j_In     Pljson;
  j_Json   Pljson;
  n_����id ������ҳ.����id%Type;
  n_��ҳid ������ҳ.��ҳid%Type;
  v_��Ϣ�� Varchar2(32767);
  v_List   Varchar2(32767);
Begin
  --�������
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');
  n_��ҳid := j_Json.Get_Number('pati_pageid');
  v_��Ϣ�� := j_Json.Get_String('info_names');

  If Nvl(n_����id, 0) = 0 Then
    Json_Out := Zljsonout('δ���벡��id�����飡');
    Return;
  End If;

  If Nvl(n_��ҳid, 0) = 0 Then
    Json_Out := Zljsonout('δ������ҳid�����飡');
    Return;
  End If;

  If Nvl(v_��Ϣ��, '-') = '-' Then
    Json_Out := Zljsonout('δ������Ϣ�������飡');
    Return;
  End If;

  For r_��Ϣ In (Select a.��Ϣ��, a.��Ϣֵ
               From ������ҳ�ӱ� A,
                    (Select /*+cardinality(B,10) */
                       Column_Value As ��Ϣ��
                      From Table(f_Str2list(v_��Ϣ��))) B
               Where a.��Ϣ�� = b.��Ϣ�� And a.����id = n_����id And a.��ҳid = n_��ҳid) Loop
    v_List := v_List || ',{"info_name":"' || Zljsonstr(r_��Ϣ.��Ϣ��) || '"';
    v_List := v_List || ',"info_value":"' || Zljsonstr(r_��Ϣ.��Ϣֵ) || '"';
    v_List := v_List || '}';
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","ext_list":[' || Substr(v_List, 2) || ']}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getpatipageextinfo;
/
Create Or Replace Procedure Zl_Cissvr_Getpatipageinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:��ȡ������ҳ�����Ϣ
  --��Σ�Json_In:��ʽ
  --    input
  --      query_type          C 1 ��ѯ����:0-������Ϣ;1-������Ϣ��չ;2-��ȡ��ҳ
  --      pati_pageids        C 1 ������Ϣ,��ʽ����:һ����:����id:��ҳID,��;һ�֣�����id,��
  --      is_babyinfo         N 1 �Ƿ����Ӥ����Ϣ:1-����;0-������
  --      is_transdeptinfo    N 1 �Ƿ����ת����Ϣ:1-����;0-������
  --      is_lastpage         N 1 �Ƿ�ȡ���һ��סԺ
  --      pati_natures        C 1 ��������:0-��ͨסԺ����,1-�������۲���,2-סԺ���۲��ˣ�������ŷָ�������Ϊ����
  --      rgst_id             N 1 �Һ�ID,���ݹҺ�ID��ѯ
  --      is_badinfo          N 0 �Ƿ������λ��Ϣ:1-����;0-������
  --����: Json_Out,��ʽ����
  --  output
  --    code                  N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message               C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    pati_count            N 1 ��ѯ�Ĳ�����Ϣ����
  --    page_list[]             1 ������
  --      pati_id             N 1 ����id
  --      pati_pageid         N 1 ��ҳid
  --      pati_name           C 1 ����
  --      pati_sex            C 1 �Ա�
  --      pati_age            C 1 ����
  --      inpatient_num       C 1 סԺ��
  --      fee_category        C 1 �ѱ�
  --      mdlpay_mode_name    C 1 ҽ�Ƹ��ʽ����
  --      mdlpay_mode_code    C 1 ҽ�Ƹ��ʽ����
  --      pati_bed            C 1 ��ǰ����
  --      pati_type           C 1 ��������(��ͨ��ҽ��������)
  --      pati_show_color     N 1 ������ʾ��ɫ
  --      pati_education      C 1 ѧ��
  --      ocpt_name           C 1 ְҵ
  --      country_name        C 1 ����
  --      pati_marital_cstatus  C 1 ����״��
  --      pati_nature         N 1 ��������:0-��ͨסԺ����,1-�������۲���,2-סԺ���۲���
  --      audit_sign          N 1 ��˱�־:������ҳ.��˱�־
  --      si_inp_status       N 1 סԺ״̬:������ҳ.״̬(0-����סԺ��1-��δ��ƣ�2-����ת�ƻ�����ת������3-��Ԥ��Ժ)
  --      pati_wardarea_id    N 1 ��ǰ����id
  --      pati_wardarea_name  C 1 ��ǰ��������
  --      pati_dept_id        N 1 ��ǰ����id
  --      pati_dept_name      C 1 ��ǰ��������
  --      adta_time           C 1 ��Ժʱ��:yyyy-mm-dd hh24:mi:ss
  --      adtd_time           C 1 ��Ժʱ��:yyyy-mm-dd hh24:mi:ss
  --      insurance_type      N 1 ����
  --      rgst_id             N 1 �Һ�id
  --      catalog_date        C 1 ��Ŀ����:yyyy-mm-dd hh24:mi:ss
  --      in_objective        C 1 סԺĿ��
  --      reg_name            C 1 �Ǽ���
  --      reg_date            C 1 סԺ�Ǽ�ʱ��
  --      pat_rsdpscn         C 1 סԺҽʦ
  --      pati_desc           C 1 ���˱�ע
  --      insurance_num       C 1 ҽ����
  --      outpatient_doctor   C 1 ����ҽʦ
  --      responsible_nurse   C 1 ���λ�ʿ
  --      hospital_admissions C 1 ��Ժ����
  --      current_conditions  C 1 ��ǰ����
  --      hospital_days       N 1 סԺ����
  --      hospital_dept       C 1 ��Ժ����
  --      level_of_care       C 1 ����ȼ�
  --      level_of_bed        C 1 ��λ�ȼ�
  --      in_dept             N 1 �Ƿ������
  --      pati_home_addr      C 1 ��ͥ��ַ
  --      pati_house_addr     C 1 ���ڵ�ַ
  --      pati_contact_addr   C 1 ��ϵ�˵�ַ
  --      baby_list[]           1 Ӥ����Ϣ��[����]
  --        pati_id           N 1 ����id
  --        pati_pageid       N 1 ��ҳid
  --        baby_num          N 1 Ӥ�����
  --        baby_name         C 1 Ӥ������
  --        baby_sex          C 1 Ӥ���Ա�
  --        baby_date         C 1 ����ʱ��
  --      trans_list[]        C   ת���б���Ϣ
  --        start_reason      C 1 ��ʼԭ��
  --        start_time        C 1 ��ʼʱ��:yyyy-mm-dd hh24:mi:ss
  --        dept_name         C 1 ��������
  --      badinfo_list[]      ��λ��Ϣ��[����]
  --        wardarea_id       N 1 ����id
  --        wardarea_name     C 1 ��������
  --        bed_no            C 1 ����
  --        bed_class_code    C 1 �������
  --        bed_class_name    C 1 ��������
  ---------------------------------------------------------------------------
  j_In           Pljson;
  j_Json         Pljson;
  n_����         Number(1);
  v_������Ϣ     Varchar2(32767);
  n_����Ӥ����Ϣ Number(1);
  n_����ת����Ϣ Number(1);
  n_������λ��Ϣ Number(1);
  n_���һ��סԺ Number(1);
  n_��ѯ��ʽ     Number(1); --0��ͨ������id:��ҳid����ʽ��ȡ��1��ͨ������id��ȡ
  n_�����       Number(1);

  v_�������� Varchar2(32767);
  n_�Һ�id   ������ҳ.�Һ�id%Type;
  I          Number(6);
  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_������Ϣ Collection_Type;

  Json_Out_Tmp Clob; --ȫ�ֱ�����Setreturnjson �ӷ����и�ֵ ��������ʹ��
  v_List       Varchar2(32767); --ȫ�ֱ�����Setreturnjson �ӷ����и�ֵ ��������ʹ��

  --���е��α궼���õ�һ���̶��ṹ
  --���α�ֻΪ����RowType�ṹ��������Ҫ��ѯ���ݡ�
  Cursor c_������Ϣ Is
    Select a.����id, a.��ҳid, a.����, a.�Ա�, a.����, a.סԺ��, a.�ѱ�, a.��������, a.��˱�־, a.״̬ As סԺ״̬,
           To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.סԺĿ��,
           a.�Ǽ���, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd') As �Ǽ�ʱ��, a.סԺҽʦ, c.���� As ҽ�Ƹ��ʽ����, c.���� ҽ�Ƹ��ʽ����, a.��Ժ���� As ��ǰ����, a.��������,
           a.ѧ��, a.ְҵ, a.����, a.����״��, a.��ǰ����id, D1.���� As ��ǰ��������, a.��Ժ����id As ��ǰ����id, D2.���� As ��ǰ��������, a.����, a.�Һ�id,
           To_Char(a.��Ŀ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ŀ����, a.��ע, a.����ҽʦ, a.���λ�ʿ, a.��Ժ����, a.��ǰ����, a.סԺ����, D3.���� As ��Ժ����,
           e.���� As ����ȼ�, a.��ͥ��ַ, a.���ڵ�ַ, a.��ϵ�˵�ַ
    From ������ҳ A, ҽ�Ƹ��ʽ C, ���ű� D1, ���ű� D2, ���ű� D3, �շ���ĿĿ¼ E
    Where a.ҽ�Ƹ��ʽ = c.���� And a.��ǰ����id = D1.Id And a.��Ժ����id = D2.Id And a.��Ժ����id = D3.Id And a.����ȼ�id = e.Id(+) And
          a.����id = 0 And a.��ҳid = 0 And Rownum < 1;
  Type Ty_������Ϣ Is Ref Cursor;
  c_Pati Ty_������Ϣ; --��̬�α����

  r_������Ϣ c_������Ϣ%RowType; --ȫ�ֱ�����Setreturnjson �ӷ����и�ֵ ��������ʹ��

  -- c_Pati_0
  --��������id��ѯ
  Cursor c_Pati_0(P������Ϣ Varchar2) Is
    Select /*+cardinality(B,10) */
     a.����id, a.��ҳid, a.����, Null As �Ա�, Null As ����, Null As סԺ��, Null As �ѱ�, Null As ��������, Null As ��˱�־, Null As סԺ״̬,
     Null As ��Ժʱ��, Null As ��Ժʱ��, Null As סԺĿ��, Null As �Ǽ���, Null As �Ǽ�ʱ��, Null As סԺҽʦ, Null As ҽ�Ƹ��ʽ����,
     Null As ҽ�Ƹ��ʽ����, Null As ��ǰ����, Null As ��������, Null As ѧ��, Null As ְҵ, Null As ����, Null As ����״��, Null As ��ǰ����id,
     Null As ��ǰ��������, Null As ��ǰ����id, Null As ��ǰ��������, Null As ����, Null As �Һ�id, Null As ��Ŀ����, Null As ��ע, Null As ����ҽʦ,
     Null As ���λ�ʿ, Null As ��Ժ����, Null As ��ǰ����, Null As סԺ����, Null As ��Ժ����, Null As ����ȼ�, a.��ͥ��ַ, a.���ڵ�ַ, a.��ϵ�˵�ַ
    From ������ҳ A, (Select Column_Value As ����id From Table(f_Str2list(P������Ϣ))) B
    Where a.����id = b.����id And (n_���һ��סԺ = 0 Or ��ҳid = (Select Max(��ҳid) From ������ҳ Where ����id = b.����id)) And
          (v_�������� Is Null Or Instr(',' || v_�������� || ',', ',' || a.�������� || ',') > 0)
    Order By a.��ҳid Desc;

  -- c_Pati_1  c_Pati_1(P������Ϣ
  --������id����ҳid��ѯ
  Cursor c_Pati_1(P������Ϣ Varchar2) Is
    Select /*+cardinality(B,10) */
     a.����id, a.��ҳid, a.����, Null As �Ա�, Null As ����, Null As סԺ��, Null As �ѱ�, Null As ��������, Null As ��˱�־, Null As סԺ״̬,
     Null As ��Ժʱ��, Null As ��Ժʱ��, Null As סԺĿ��, Null As �Ǽ���, Null As �Ǽ�ʱ��, Null As סԺҽʦ, Null As ҽ�Ƹ��ʽ����,
     Null As ҽ�Ƹ��ʽ����, Null As ��ǰ����, Null As ��������, Null As ѧ��, Null As ְҵ, Null As ����, Null As ����״��, Null As ��ǰ����id,
     Null As ��ǰ��������, Null As ��ǰ����id, Null As ��ǰ��������, Null As ����, Null As �Һ�id, Null As ��Ŀ����, Null As ��ע, Null As ����ҽʦ,
     Null As ���λ�ʿ, Null As ��Ժ����, Null As ��ǰ����, Null As סԺ����, Null As ��Ժ����, Null As ����ȼ�, Null As ��ͥ��ַ, Null As ���ڵ�ַ,
     Null As ���ڵ�ַ
    From ������ҳ A, (Select C1 As ����id, C2 As ��ҳid From Table(f_Str2list2(P������Ϣ, ',', ':'))) B
    Where a.����id = b.����id And a.��ҳid = b.��ҳid;

  -- c_Pati_0_1  c_Pati_0_1(P������Ϣ
  --��������id��ѯ
  Cursor c_Pati_0_1(P������Ϣ Varchar2) Is
    Select /*+cardinality(B,10) */
     a.����id, a.��ҳid, a.����, a.�Ա�, a.����, a.סԺ��, a.�ѱ�, a.��������, a.��˱�־, a.״̬ As סԺ״̬,
     To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.סԺĿ��, a.�Ǽ���,
     To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd') As �Ǽ�ʱ��, a.סԺҽʦ, c.���� As ҽ�Ƹ��ʽ����, c.���� ҽ�Ƹ��ʽ����, a.��Ժ���� As ��ǰ����, a.��������, a.ѧ��, a.ְҵ,
     a.����, a.����״��, a.��ǰ����id, D1.���� As ��ǰ��������, a.��Ժ����id As ��ǰ����id, D2.���� As ��ǰ��������, a.����, a.�Һ�id,
     To_Char(a.��Ŀ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ŀ����, a.��ע, Null As ����ҽʦ, Null As ���λ�ʿ, Null As ��Ժ����, Null As ��ǰ����,
     Null As סԺ����, Null As ��Ժ����, Null As ����ȼ�, a.��ͥ��ַ, a.���ڵ�ַ, a.��ϵ�˵�ַ
    From ������ҳ A, (Select Column_Value As ����id From Table(f_Str2list(P������Ϣ))) B, ҽ�Ƹ��ʽ C, ���ű� D1, ���ű� D2
    Where a.����id = b.����id And (n_���һ��סԺ = 0 Or ��ҳid = (Select Max(��ҳid) From ������ҳ Where ����id = b.����id)) And
          a.ҽ�Ƹ��ʽ = c.���� And a.��ǰ����id = D1.Id(+) And a.��Ժ����id = D2.Id(+)
    Order By a.��ҳid Desc;

  -- c_Pati_1_1  c_Pati_1_1(P������Ϣ
  --������id����ҳid��ѯ
  Cursor c_Pati_1_1(P������Ϣ Varchar2) Is
    Select /*+cardinality(B,10) */
     a.����id, a.��ҳid, a.����, a.�Ա�, a.����, a.סԺ��, a.�ѱ�, a.��������, a.��˱�־, a.״̬ As סԺ״̬,
     To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.סԺĿ��, a.�Ǽ���,
     To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd') As �Ǽ�ʱ��, a.סԺҽʦ, c.���� As ҽ�Ƹ��ʽ����, c.���� ҽ�Ƹ��ʽ����, a.��Ժ���� As ��ǰ����, a.��������, a.ѧ��, a.ְҵ,
     a.����, a.����״��, a.��ǰ����id, D1.���� As ��ǰ��������, a.��Ժ����id As ��ǰ����id, D2.���� As ��ǰ��������, a.����, a.�Һ�id,
     To_Char(a.��Ŀ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ŀ����, a.��ע, a.����ҽʦ, a.���λ�ʿ, a.��Ժ����, a.��ǰ����, a.סԺ����, D3.���� As ��Ժ����,
     e.���� As ����ȼ�, a.��ͥ��ַ, a.���ڵ�ַ, a.��ϵ�˵�ַ
    From ������ҳ A, (Select C1 As ����id, C2 As ��ҳid From Table(f_Str2list2(P������Ϣ, ',', ':'))) B, ҽ�Ƹ��ʽ C, ���ű� D1, ���ű� D2,
         ���ű� D3, �շ���ĿĿ¼ E
    Where a.����id = b.����id And a.��ҳid = b.��ҳid And a.ҽ�Ƹ��ʽ = c.���� And a.��ǰ����id = D1.Id(+) And a.��Ժ����id = D2.Id(+) And
          a.��Ժ����id = D3.Id(+) And a.����ȼ�id = e.Id(+);

  --c_Pati_0_0  c_Pati_0_0(P������Ϣ
  --��������id��ѯ
  Cursor c_Pati_0_0(P������Ϣ Varchar2) Is
    Select /*+cardinality(B,10) */
     a.����id, a.��ҳid, a.����, a.�Ա�, a.����, a.סԺ��, a.�ѱ�, a.��������, a.��˱�־, a.״̬ As סԺ״̬,
     To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.סԺĿ��, a.�Ǽ���,
     To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd') As �Ǽ�ʱ��, a.סԺҽʦ, Null As ҽ�Ƹ��ʽ����, Null As ҽ�Ƹ��ʽ����, Null As ��ǰ����, Null As ��������,
     Null As ѧ��, Null As ְҵ, Null As ����, Null As ����״��, Null As ��ǰ����id, Null As ��ǰ��������, Null As ��ǰ����id, Null As ��ǰ��������,
     Null As ����, Null As �Һ�id, Null As ��Ŀ����, Null As ��ע, Null As ����ҽʦ, Null As ���λ�ʿ, Null As ��Ժ����, Null As ��ǰ����,
     Null As סԺ����, Null As ��Ժ����, Null As ����ȼ�, a.��ͥ��ַ, a.���ڵ�ַ, a.��ϵ�˵�ַ
    From ������ҳ A, (Select Column_Value As ����id From Table(f_Str2list(P������Ϣ))) B
    Where a.����id = b.����id And (n_���һ��סԺ = 0 Or ��ҳid = (Select Max(��ҳid) From ������ҳ Where ����id = b.����id))
    Order By a.��ҳid Desc;

  --c_Pati_1_0  c_Pati_1_0(P������Ϣ
  --������id����ҳid��ѯ
  Cursor c_Pati_1_0(P������Ϣ Varchar2) Is
    Select /*+cardinality(B,10) */
     a.����id, a.��ҳid, a.����, a.�Ա�, a.����, a.סԺ��, a.�ѱ�, a.��������, a.��˱�־, a.״̬ As סԺ״̬,
     To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.סԺĿ��, a.�Ǽ���,
     To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd') As �Ǽ�ʱ��, a.סԺҽʦ, Null As ҽ�Ƹ��ʽ����, Null As ҽ�Ƹ��ʽ����, Null As ��ǰ����, Null As ��������,
     Null As ѧ��, Null As ְҵ, Null As ����, Null As ����״��, Null As ��ǰ����id, Null As ��ǰ��������, Null As ��ǰ����id, Null As ��ǰ��������,
     Null As ����, Null As �Һ�id, Null As ��Ŀ����, Null As ��ע, Null As ����ҽʦ, Null As ���λ�ʿ, Null As ��Ժ����, Null As ��ǰ����,
     Null As סԺ����, Null As ��Ժ����, Null As ����ȼ�, a.��ͥ��ַ, a.���ڵ�ַ, a.��ϵ�˵�ַ
    From ������ҳ A, (Select C1 As ����id, C2 As ��ҳid From Table(f_Str2list2(P������Ϣ, ',', ':'))) B
    Where a.����id = b.����id And a.��ҳid = b.��ҳid;

  Procedure Setreturnjson
  (
    n_�Ƿ�Ӥ��   Number,
    n_ת��       Number,
    n_����λ��Ϣ Number
  ) Is
    --���ܣ����ݼ�¼����Ϣƴ��json��
    --      �α�һ�м�¼��һ��
    v_List_Baby Varchar2(32767);
    v_List_Tran Varchar2(32767);
    v_List_Bad  Varchar2(32767);
    v_ҽ����    ������ҳ�ӱ�.��Ϣֵ%Type;
    n_��ɫ      ��������.��ɫ%Type;
  Begin
    --��ҳ��Ϣ
    Zljsonputvalue(v_List, 'pati_id', r_������Ϣ.����id, 1, 1);
    Zljsonputvalue(v_List, 'pati_pageid', r_������Ϣ.��ҳid);
    Zljsonputvalue(v_List, 'pati_name', r_������Ϣ.����);
  
    --������Ϣ
    If n_���� < 2 Then
      Zljsonputvalue(v_List, 'pati_sex', r_������Ϣ.�Ա�);
      Zljsonputvalue(v_List, 'pati_age', r_������Ϣ.����);
      Zljsonputvalue(v_List, 'inpatient_num', r_������Ϣ.סԺ��, 0);
      Zljsonputvalue(v_List, 'fee_category', r_������Ϣ.�ѱ�);
      Zljsonputvalue(v_List, 'pati_nature', r_������Ϣ.��������, 1);
      Zljsonputvalue(v_List, 'audit_sign', r_������Ϣ.��˱�־, 1);
      Zljsonputvalue(v_List, 'si_inp_status', r_������Ϣ.סԺ״̬, 1);
      Zljsonputvalue(v_List, 'adta_time', r_������Ϣ.��Ժʱ��);
      Zljsonputvalue(v_List, 'adtd_time', r_������Ϣ.��Ժʱ��);
      Zljsonputvalue(v_List, 'in_objective', r_������Ϣ.סԺĿ��);
      Zljsonputvalue(v_List, 'reg_name', r_������Ϣ.�Ǽ���);
      Zljsonputvalue(v_List, 'reg_date', r_������Ϣ.�Ǽ�ʱ��);
      Zljsonputvalue(v_List, 'pat_rsdpscn', r_������Ϣ.סԺҽʦ);
      Zljsonputvalue(v_List, 'pati_home_addr', r_������Ϣ.��ͥ��ַ);
      Zljsonputvalue(v_List, 'pati_house_addr', r_������Ϣ.���ڵ�ַ);
      Zljsonputvalue(v_List, 'pati_contact_addr', r_������Ϣ.��ϵ�˵�ַ);
    End If;
  
    --��չ��Ϣ
    If n_���� = 1 Then
      --      mdlpay_mode_name    C 1 ҽ�Ƹ��ʽ����
      --      mdlpay_mode_code    C 1 ҽ�Ƹ��ʽ����
      Zljsonputvalue(v_List, 'mdlpay_mode_name', r_������Ϣ.ҽ�Ƹ��ʽ����);
      Zljsonputvalue(v_List, 'mdlpay_mode_code', r_������Ϣ.ҽ�Ƹ��ʽ����);
      --      pati_bed            C 1 ��ǰ����
      --      pati_type           C 1 ��������(��ͨ��ҽ��������)
      --      pati_show_color     N 1 ������ʾ��ɫ
      --      pati_education      C 1 ѧ��
      --      ocpt_name           C 1 ְҵ
      --      country_name        C 1 ����
      --      pati_marital_cstatus  C 1 ����״��
      Zljsonputvalue(v_List, 'pati_bed', r_������Ϣ.��ǰ����);
      Zljsonputvalue(v_List, 'pati_type', r_������Ϣ.��������);
    
      If r_������Ϣ.�������� Is Not Null Then
        Select Max(��ɫ) Into n_��ɫ From �������� Where ���� = Nvl(r_������Ϣ.��������, '');
      End If;
      Zljsonputvalue(v_List, 'pati_show_color', n_��ɫ, 1);
      Zljsonputvalue(v_List, 'pati_education', r_������Ϣ.ѧ��);
      Zljsonputvalue(v_List, 'ocpt_name', r_������Ϣ.ְҵ);
      Zljsonputvalue(v_List, 'country_name', r_������Ϣ.����);
      Zljsonputvalue(v_List, 'pati_marital_cstatus', r_������Ϣ.����״��);
      --      pati_wardarea_id    N 1 ��ǰ����id
      --      pati_wardarea_name  C 1 ��ǰ��������
      --      pati_dept_id        N 1 ��ǰ����id
      --      pati_dept_name      C 1 ��ǰ��������
      Zljsonputvalue(v_List, 'pati_wardarea_id', r_������Ϣ.��ǰ����id, 1);
      Zljsonputvalue(v_List, 'pati_wardarea_name', r_������Ϣ.��ǰ��������);
      Zljsonputvalue(v_List, 'pati_dept_id', r_������Ϣ.��ǰ����id, 1);
      Zljsonputvalue(v_List, 'pati_dept_name', r_������Ϣ.��ǰ��������);
      --      insurance_type      N 1 ����
      --      rgst_id             N 1 �Һ�id
      --      catalog date        C 1 ��Ŀ����:yyyy-mm-dd hh24:mi:ss
      Zljsonputvalue(v_List, 'insurance_type', r_������Ϣ.����, 1);
      Zljsonputvalue(v_List, 'rgst_id', r_������Ϣ.�Һ�id, 1);
      Zljsonputvalue(v_List, 'catalog_date', r_������Ϣ.��Ŀ����);
      Zljsonputvalue(v_List, 'pati_desc', r_������Ϣ.��ע);
      --      outpatient_doctor   C 1 ����ҽʦ
      --      responsible_nurse   C 1 ���λ�ʿ
      --      hospital_admissions C 1 ��Ժ����
      --      current_conditions  C 1 ��ǰ����
      --      hospital_days       N 1 סԺ����
      --      hospital_dept       C 1 ��Ժ����
      --      level_of_care       C 1 ����ȼ�
      Zljsonputvalue(v_List, 'outpatient_doctor', r_������Ϣ.����ҽʦ);
      Zljsonputvalue(v_List, 'responsible_nurse', r_������Ϣ.���λ�ʿ);
      Zljsonputvalue(v_List, 'hospital_admissions', r_������Ϣ.��Ժ����);
      Zljsonputvalue(v_List, 'current_conditions', r_������Ϣ.��ǰ����);
      Zljsonputvalue(v_List, 'hospital_days', r_������Ϣ.סԺ����, 1);
      Zljsonputvalue(v_List, 'hospital_dept', r_������Ϣ.��Ժ����);
      Zljsonputvalue(v_List, 'level_of_care', r_������Ϣ.����ȼ�);
    
      Select Max(Decode(a.��Ϣ��, 'ҽ����', a.��Ϣֵ, '')) As ҽ����
      Into v_ҽ����
      From ������ҳ�ӱ� A
      Where a.����id = r_������Ϣ.����id And a.��ҳid = r_������Ϣ.��ҳid And a.��Ϣ�� In ('ҽ����');
    
      Zljsonputvalue(v_List, 'insurance_num', v_ҽ����);
    
      For r_�ȼ� In (Select b.���� As ��λ�ȼ�
                   From ���˱䶯��¼ A, �շ���ĿĿ¼ B
                   Where r_������Ϣ.����id = a.����id And r_������Ϣ.��ҳid = a.��ҳid And
                         ((a.��ֹʱ�� Is Null And ���Ӵ�λ = 0) Or (a.��ֹʱ�� Is Not Null And ��ֹԭ�� = 1)) And a.��λ�ȼ�id = b.Id(+)) Loop
        Zljsonputvalue(v_List, 'level_of_bed', r_�ȼ�.��λ�ȼ�);
      End Loop;
      Select Count(1)
      Into n_�����
      From ���˱䶯��¼
      Where ����id = r_������Ϣ.����id And ��ҳid = r_������Ϣ.��ҳid And ��ʼԭ�� = 2 And ���� Is Not Null And Rownum < 2;
      Zljsonputvalue(v_List, 'in_dept', n_�����, 1);
    End If;
  
    If n_�Ƿ�Ӥ�� = 1 Then
      For r_Ӥ����Ϣ In (Select ����id, ��ҳid, ��� As Ӥ�����, Ӥ������, Ӥ���Ա�, To_Char(����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��
                     From ������������¼
                     Where ����id = r_������Ϣ.����id And ��ҳid = r_������Ϣ.��ҳid) Loop
        --      baby_list[]           1 Ӥ����Ϣ��[����]
        --        pati_id           N 1 ����id
        --        pati_pageid       N 1 ��ҳid
        --        baby_num          N 1 Ӥ�����
        --        baby_name         C 1 Ӥ������
        --        baby_sex          C 1 Ӥ���Ա�
        --        baby_date         C 1 ����ʱ��
        v_List_Baby := v_List_Baby || ',{';
        v_List_Baby := v_List_Baby || '"pati_id":' || r_Ӥ����Ϣ.����id || ',';
        v_List_Baby := v_List_Baby || '"pati_pageid":' || r_Ӥ����Ϣ.��ҳid || ',';
        v_List_Baby := v_List_Baby || '"baby_num":' || r_Ӥ����Ϣ.Ӥ����� || ',';
        v_List_Baby := v_List_Baby || '"baby_name":"' || Zljsonstr(r_Ӥ����Ϣ.Ӥ������) || '",';
        v_List_Baby := v_List_Baby || '"baby_sex":"' || r_Ӥ����Ϣ.Ӥ���Ա� || '",';
        v_List_Baby := v_List_Baby || '"baby_date":"' || r_Ӥ����Ϣ.����ʱ�� || '"';
        v_List_Baby := v_List_Baby || '}';
      End Loop;
      v_List := v_List || ',"baby_list":[' || Substr(v_List_Baby, 2) || ']';
    End If;
  
    If n_ת�� = 1 Then
      For r_ת�� In (Select a.��ʼԭ��, To_Char(a.��ʼʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��, b.����
                   From ���˱䶯��¼ A, ���ű� B
                   Where a.����id = b.Id And a.��ʼʱ�� Is Not Null And a.��ʼԭ�� = 3 And a.����id = r_������Ϣ.����id And
                         ��ҳid = r_������Ϣ.��ҳid) Loop
      
        --      trans_list[]        C   ת���б���Ϣ
        --        start_reason      C 1 ��ʼԭ��
        --        start_time        C 1 ��ʼʱ��:yyyy-mm-dd hh24:mi:ss
        --        dept_name         C 1 ��������
        v_List_Tran := v_List_Tran || ',{';
        v_List_Tran := v_List_Tran || '"start_reason":"' || r_ת��.��ʼԭ�� || '",';
        v_List_Tran := v_List_Tran || '"start_time":"' || r_ת��.��ʼʱ�� || '",';
        v_List_Tran := v_List_Tran || '"dept_name":"' || Zljsonstr(r_ת��.����) || '"';
        v_List_Tran := v_List_Tran || '}';
      End Loop;
      v_List := v_List || ',"trans_list":[' || Substr(v_List_Tran, 2) || ']';
    End If;
  
    If Nvl(n_����λ��Ϣ, 0) = 1 Then
      For r_��λ In (Select a.����id, c.���� As ��������, a.����, b.����, b.����
                   From ��λ״����¼ A, ��λ���Ʒ��� B, ���ű� C
                   Where a.��λ���� = b.����(+) And a.����id = c.Id And a.����id = r_������Ϣ.����id) Loop
      
        --      badinfo_list[]      ��λ��Ϣ��[����]
        --        wardarea_id       N 1 ����id
        --        wardarea_name     C 1 ��������
        --        bed_no            C 1 ����
        --        bed_class_code    C 1 �������
        --        bed_class_name    C 1 ��������
        v_List_Bad := v_List_Bad || ',{';
        v_List_Bad := v_List_Bad || '"wardarea_id":' || r_��λ.����id || ',';
        v_List_Bad := v_List_Bad || '"wardarea_name":"' || Zljsonstr(r_��λ.��������) || '",';
        v_List_Bad := v_List_Bad || '"bed_no":"' || Zljsonstr(r_��λ.����) || '",';
        v_List_Bad := v_List_Bad || '"bed_class_code":"' || Zljsonstr(r_��λ.����) || '",';
        v_List_Bad := v_List_Bad || '"bed_class_name":"' || Zljsonstr(r_��λ.����) || '"';
        v_List_Bad := v_List_Bad || '}';
      End Loop;
      v_List := v_List || ',"badinfo_list":[' || Substr(v_List_Bad, 2) || ']';
    End If;
    v_List := v_List || '}';
  
    If Length(v_List) > 20000 Then
      Json_Out_Tmp := Json_Out_Tmp || v_List;
      v_List       := ',';
    End If;
  End;
Begin
  --�������
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');
  --      query_type          C 1 ��ѯ����:0-������Ϣ;1-������Ϣ��չ
  --      pati_pageids        C 1 ������Ϣ,��ʽ����:һ����:����id:��ҳID,��;һ�֣�����id,��
  --      is_babyinfo         N 1 �Ƿ����Ӥ����Ϣ:1-����;0-������
  --      is_transdeptinfo    N 1 �Ƿ����ת����Ϣ:1-����;0-������
  --      is_lastpage         N 1 �Ƿ�ȡ���һ��סԺ
  --      pati_natures        C 1 ��������:0-��ͨסԺ����,1-�������۲���,2-סԺ���۲��ˣ�������ŷָ�������Ϊ����
  --      rgst_id             N 1 �Һ�ID,���ݹҺ�ID��ѯ
  --      is_badinfo          N 1 �Ƿ������λ��Ϣ:1-����;0-������
  n_����         := Nvl(j_Json.Get_Number('query_type'), 0);
  v_������Ϣ     := j_Json.Get_String('pati_pageids');
  n_����Ӥ����Ϣ := Nvl(j_Json.Get_Number('is_babyinfo'), 0);
  n_����ת����Ϣ := Nvl(j_Json.Get_Number('is_transdeptinfo'), 0);
  n_���һ��סԺ := Nvl(j_Json.Get_Number('is_lastpage'), 0);
  v_��������     := j_Json.Get_String('pati_natures');
  n_�Һ�id       := Nvl(j_Json.Get_Number('rgst_id'), 0);
  n_������λ��Ϣ := Nvl(j_Json.Get_Number('is_badinfo'), 0);

  If Nvl(v_������Ϣ, '-') = '-' Then
    Json_Out := Zljsonout('δ���벡����Ϣ�����飡');
    Return;
  End If;

  n_��ѯ��ʽ := 0;
  If Instr(v_������Ϣ, ':') > 0 Then
    n_��ѯ��ʽ := 1;
  End If;
  If Nvl(n_�Һ�id, 0) <> 0 Then
    n_��ѯ��ʽ := 2;
  End If;

  --�� v_������Ϣ ����װ�ɲ�����4000 �ļ��ϴ�����ֹʹ�ö�̬�ڴ���ѯʱ��������
  I := 0;
  While v_������Ϣ Is Not Null Loop
    If Length(v_������Ϣ) <= 4000 Then
      Col_������Ϣ(I) := v_������Ϣ;
      v_������Ϣ := Null;
    Else
      Col_������Ϣ(I) := Substr(v_������Ϣ, 1, Instr(v_������Ϣ, ',', 3980) - 1);
      v_������Ϣ := Substr(v_������Ϣ, Instr(v_������Ϣ, ',', 3980) + 1);
    End If;
    I := I + 1;
  End Loop;

  --���ο�ʼ����ƴ��
  Json_Out_Tmp := '{"output":{"code":1,"message":"�ɹ�","page_list":[';

  If n_���� = 2 Then
    --����ѯ��ҳ
    If n_��ѯ��ʽ = 0 Then
      -- c_Pati_0 
      --��������id��ѯ      
      For K In 0 .. Col_������Ϣ.Count - 1 Loop
        Open c_Pati_0(Col_������Ϣ(K));
        Loop
          Fetch c_Pati_0
            Into r_������Ϣ;
          Exit When c_Pati_0%NotFound;
          Setreturnjson(n_����Ӥ����Ϣ, n_����ת����Ϣ, n_������λ��Ϣ);
        End Loop;
        Close c_Pati_0;
      End Loop;
    Elsif n_��ѯ��ʽ = 1 Then
      -- c_Pati_1  
      --������id����ҳid��ѯ
      For K In 0 .. Col_������Ϣ.Count - 1 Loop
        Open c_Pati_1(Col_������Ϣ(K));
        Loop
          Fetch c_Pati_1
            Into r_������Ϣ;
          Exit When c_Pati_1%NotFound;
          Setreturnjson(n_����Ӥ����Ϣ, n_����ת����Ϣ, n_������λ��Ϣ);
        End Loop;
        Close c_Pati_1;
      End Loop;
    Elsif n_��ѯ��ʽ = 2 Then
      --���Һ�ID��ѯ
      Open c_Pati For
        Select a.����id, a.��ҳid, a.����, Null As �Ա�, Null As ����, Null As סԺ��, Null As �ѱ�, Null As ��������, Null As ��˱�־,
               Null As סԺ״̬, Null As ��Ժʱ��, Null As ��Ժʱ��, Null As סԺĿ��, Null As �Ǽ���, Null As �Ǽ�ʱ��, Null As סԺҽʦ,
               Null As ҽ�Ƹ��ʽ����, Null As ҽ�Ƹ��ʽ����, Null As ��ǰ����, Null As ��������, Null As ѧ��, Null As ְҵ, Null As ����,
               Null As ����״��, Null As ��ǰ����id, Null As ��ǰ��������, Null As ��ǰ����id, Null As ��ǰ��������, Null As ����, Null As �Һ�id,
               Null As ��Ŀ����, Null As ��ע, Null As ����ҽʦ, Null As ���λ�ʿ, Null As ��Ժ����, Null As ��ǰ����, Null As סԺ����,
               Null As ��Ժ����, Null As ����ȼ�, Null As ��ͥ��ַ, Null As ���ڵ�ַ, Null As ���ڵ�ַ
        From ������ҳ A
        Where a.�Һ�id = n_�Һ�id;
      Loop
        Fetch c_Pati
          Into r_������Ϣ;
        Exit When c_Pati%NotFound;
        Setreturnjson(n_����Ӥ����Ϣ, n_����ת����Ϣ, n_������λ��Ϣ);
      End Loop;
    End If;
  Elsif n_���� = 1 Then
    --��ѯ������Ϣ+��չ��Ϣ
    If n_��ѯ��ʽ = 0 Then
      -- c_Pati_0_1  
      --��������id��ѯ
      For K In 0 .. Col_������Ϣ.Count - 1 Loop
        Open c_Pati_0_1(Col_������Ϣ(K));
        Loop
          Fetch c_Pati_0_1
            Into r_������Ϣ;
          Exit When c_Pati_0_1%NotFound;
          Setreturnjson(n_����Ӥ����Ϣ, n_����ת����Ϣ, n_������λ��Ϣ);
        End Loop;
        Close c_Pati_0_1;
      End Loop;
    Elsif n_��ѯ��ʽ = 1 Then
      -- c_Pati_1_1  
      --������id����ҳid��ѯ
      For K In 0 .. Col_������Ϣ.Count - 1 Loop
        Open c_Pati_1_1(Col_������Ϣ(K));
        Loop
          Fetch c_Pati_1_1
            Into r_������Ϣ;
          Exit When c_Pati_1_1%NotFound;
          Setreturnjson(n_����Ӥ����Ϣ, n_����ת����Ϣ, n_������λ��Ϣ);
        End Loop;
        Close c_Pati_1_1;
      End Loop;
    Elsif n_��ѯ��ʽ = 2 Then
      --���Һ�ID��ѯ
      Open c_Pati For
        Select a.����id, a.��ҳid, a.����, a.�Ա�, a.����, a.סԺ��, a.�ѱ�, a.��������, a.��˱�־, a.״̬ As סԺ״̬,
               To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��,
               a.סԺĿ��, a.�Ǽ���, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd') As �Ǽ�ʱ��, a.סԺҽʦ, c.���� As ҽ�Ƹ��ʽ����, c.���� ҽ�Ƹ��ʽ����,
               a.��Ժ���� As ��ǰ����, a.��������, a.ѧ��, a.ְҵ, a.����, a.����״��, a.��ǰ����id, D1.���� As ��ǰ��������, a.��Ժ����id As ��ǰ����id,
               D2.���� As ��ǰ��������, a.����, a.�Һ�id, To_Char(a.��Ŀ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ŀ����, a.��ע, a.����ҽʦ, a.���λ�ʿ,
               a.��Ժ����, a.��ǰ����, a.סԺ����, D3.���� As ��Ժ����, e.���� As ����ȼ�, a.��ͥ��ַ, a.���ڵ�ַ, a.��ϵ�˵�ַ
        From ������ҳ A, ҽ�Ƹ��ʽ C, ���ű� D1, ���ű� D2, ���ű� D3, �շ���ĿĿ¼ E
        Where a.�Һ�id = n_�Һ�id And a.ҽ�Ƹ��ʽ = c.���� And a.��ǰ����id = D1.Id(+) And a.��Ժ����id = D2.Id(+) And
              a.��Ժ����id = D3.Id(+) And a.����ȼ�id = e.Id(+);
    
      Loop
        Fetch c_Pati
          Into r_������Ϣ;
        Exit When c_Pati%NotFound;
        Setreturnjson(n_����Ӥ����Ϣ, n_����ת����Ϣ, n_������λ��Ϣ);
      End Loop;
    End If;
  Else
    --ֻ��ѯ������Ϣ
    If n_��ѯ��ʽ = 0 Then
      --c_Pati_0_0 
      --��������id��ѯ
      For K In 0 .. Col_������Ϣ.Count - 1 Loop
        Open c_Pati_0_0(Col_������Ϣ(K));
        Loop
          Fetch c_Pati_0_0
            Into r_������Ϣ;
          Exit When c_Pati_0_0%NotFound;
          Setreturnjson(n_����Ӥ����Ϣ, n_����ת����Ϣ, n_������λ��Ϣ);
        End Loop;
        Close c_Pati_0_0;
      End Loop;
    Elsif n_��ѯ��ʽ = 1 Then
      --c_Pati_1_0 
      --������id����ҳid��ѯ
      For K In 0 .. Col_������Ϣ.Count - 1 Loop
        Open c_Pati_1_0(Col_������Ϣ(K));
        Loop
          Fetch c_Pati_1_0
            Into r_������Ϣ;
          Exit When c_Pati_1_0%NotFound;
          Setreturnjson(n_����Ӥ����Ϣ, n_����ת����Ϣ, n_������λ��Ϣ);
        End Loop;
        Close c_Pati_1_0;
      End Loop;
    Elsif n_��ѯ��ʽ = 2 Then
      --���Һ�ID��ѯ
      Open c_Pati For
        Select /*+cardinality(B,10) */
         a.����id, a.��ҳid, a.����, a.�Ա�, a.����, a.סԺ��, a.�ѱ�, a.��������, a.��˱�־, a.״̬ As סԺ״̬,
         To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.סԺĿ��,
         a.�Ǽ���, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd') As �Ǽ�ʱ��, a.סԺҽʦ, Null As ҽ�Ƹ��ʽ����, Null As ҽ�Ƹ��ʽ����, Null As ��ǰ����,
         Null As ��������, Null As ѧ��, Null As ְҵ, Null As ����, Null As ����״��, Null As ��ǰ����id, Null As ��ǰ��������, Null As ��ǰ����id,
         Null As ��ǰ��������, Null As ����, Null As �Һ�id, Null As ��Ŀ����, Null As ��ע, Null As ����ҽʦ, Null As ���λ�ʿ, Null As ��Ժ����,
         Null As ��ǰ����, Null As סԺ����, Null As ��Ժ����, Null As ����ȼ�, a.��ͥ��ַ, a.���ڵ�ַ, a.��ϵ�˵�ַ
        From ������ҳ A
        Where a.�Һ�id = n_�Һ�id;
      Loop
        Fetch c_Pati
          Into r_������Ϣ;
        Exit When c_Pati%NotFound;
        Setreturnjson(n_����Ӥ����Ϣ, n_����ת����Ϣ, n_������λ��Ϣ);
      End Loop;
    End If;
  End If;

  --���ν�������ƴ��
  If v_List <> ',' Then
    Json_Out_Tmp := Json_Out_Tmp || v_List || ']}}';
  Else
    Json_Out_Tmp := Json_Out_Tmp || ']}}';
  End If;

  Json_Out := Json_Out_Tmp;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getpatipageinfo;
/
Create Or Replace Procedure Zl_Cissvr_Getpativisitid
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --���ܣ���ȡ�������һ�ιҺŵĹҺż�¼
  --��Σ�Json_In:��ʽ
  --input
  --  query_type        N    1 ��ѯ��ʽ  0-������ID��ѯ���һ�εľ���ID(�����סԺ) 1-��ȡÿ��סԺ����ҳID
  --  pati_id           N    1 ����id
  --  occasion          N    1 ���ϣ�0-�����֣�1-���2-סԺ
  --����      json
  --output
  --    code                N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    visit_id            C  1 ����id
  --    occasion            N  1 ����
  --    visit_list[]           query_type=1 ʱ������ҳid�Ͳ�������
  --    visit_id             N  1 ����id
  --    pati_type           N  1 ��������
  -------------------------------------------
  j_In Pljson;

  j_Json Pljson;
  v_List Varchar2(32767);

  n_����id Number(18);
  n_����id Number;
  n_����   Number;
  n_��ҳid Number;
  n_Type   Number;
  n_Id     Number;
Begin
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_Type   := j_Json.Get_Number('query_type');
  n_����id := j_Json.Get_Number('pati_id');
  n_����   := j_Json.Get_Number('occasion');
  Json_Out := '{"output":{"code":1,"message":"�ɹ�"';
  If Nvl(n_����id, 0) <> 0 Then
    If Nvl(n_Type, 0) = 0 Then
      If Nvl(n_����, 0) = 0 Then
        Begin
          Select ID
          Into n_����id
          From (Select ID From ���˹Һż�¼ Where ����id = n_����id And Mod(��¼״̬, 2) <> 0 Order By �Ǽ�ʱ�� Desc)
          Where Rownum < 2;
        Exception
          When Others Then
            Null;
        End;
        Begin
          Select Max(a.��ҳid)
          Into n_��ҳid
          From ������ҳ A, ��Ժ���� B
          Where a.����id = n_����id And Nvl(a.��ҳid, 0) <> 0 And a.����id = b.����id And a.��ҳid = b.��ҳid;
        Exception
          When Others Then
            Null;
        End;
        If Nvl(n_��ҳid, 0) = 0 Then
          n_Id   := Nvl(n_����id, 0);
          n_���� := 1;
        Else
          n_Id   := n_��ҳid;
          n_���� := 2;
        End If;
        Json_Out := Json_Out || ',"visit_id":' || n_Id || ',"occasion":' || n_���� || '}}';
      Elsif Nvl(n_����, 0) = 1 Then
        Begin
          Select ID
          Into n_����id
          From (Select ID From ���˹Һż�¼ Where ����id = n_����id And Mod(��¼״̬, 2) <> 0 Order By �Ǽ�ʱ�� Desc)
          Where Rownum < 2;
        Exception
          When Others Then
            Null;
        End;
        n_Id     := Nvl(n_����id, 0);
        n_����   := 1;
        Json_Out := Json_Out || ',"visit_id":' || n_Id || ',"occasion":' || n_���� || '}}';
      Elsif Nvl(n_����, 0) = 2 Then
        Begin
          Select Max(a.��ҳid)
          Into n_��ҳid
          From ������ҳ A, ��Ժ���� B
          Where a.����id = n_����id And Nvl(a.��ҳid, 0) <> 0 And a.����id = b.����id And a.��ҳid = b.��ҳid;
        Exception
          When Others Then
            Null;
        End;
        n_Id     := Nvl(n_��ҳid, 0);
        n_����   := 2;
        Json_Out := Json_Out || ',"visit_id":' || n_Id || ',"occasion":' || n_���� || '}}';
      End If;
    Elsif Nvl(n_Type, 0) = 1 Then
      For R In (Select ��ҳid, �������� From ������ҳ Where ����id = n_����id And Nvl(��ҳid, 0) <> 0 Order By ��ҳid) Loop
        Zljsonputvalue(v_List, 'visit_id', r.��ҳid, 1, 1);
        Zljsonputvalue(v_List, 'pati_type', r.��������, 1, 2);
      End Loop;
      Json_Out := Json_Out || ',visit_list:[' || v_List || ']}}';
    End If;
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getpativisitid;
/
Create Or Replace Procedure Zl_Cissvr_Getpativisitinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  -------------------------------------------
  --���ܣ���ȡ���˵ľ����¼
  --��Σ�Json_In:��ʽ
  --input
  --  pati_id                 N    1 ����id
  --����      json
  --output
  --    code                N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    visit_list[]
  --      visit_id             N  1 ����id
  --      pati_nature          N  1 ��������
  --      occasion             N  1 ���� 1-���� 2-סԺ
  --      regist_no            C  1 �Һŵ�
  --      create_time          C  1 �Ǽ�ʱ��
  --      pati_type            C  1 ��������
  --      insurance_type       C  1 ����
  --      adta_time            C  1 ��Ժʱ��
  -------------------------------------------
  j_In   Pljson;
  j_Json Pljson;
  v_List Varchar2(32767);

  n_����id Number(18);
Begin
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","visit_list":[';
  For R In (Select *
            From (Select 1 ����, ID ID, NO, 0 ��������, To_Char(�Ǽ�ʱ��, 'YYYY-MM-DD hh24:mi:ss') �Ǽ�ʱ��, Null ��������, Null ����,
                          Null As ��Ժ����
                   From ���˹Һż�¼
                   Where ����id = n_����id And Mod(��¼״̬, 2) <> 0
                   Union All
                   Select 2 ����, ��ҳid ID, '' || ��ҳid NO, ��������, To_Char(�Ǽ�ʱ��, 'YYYY-MM-DD hh24:mi:ss') �Ǽ�ʱ��, ��������, ����, ��Ժ����
                   From ������ҳ
                   Where ����id = n_����id And Nvl(��ҳid, 0) <> 0)
            Order By NO Desc) Loop
    Zljsonputvalue(v_List, 'visit_id', r.Id, 1, 1);
    Zljsonputvalue(v_List, 'occasion', r.����, 1);
    Zljsonputvalue(v_List, 'pati_nature', r.��������, 1);
    Zljsonputvalue(v_List, 'regist_no', r.No);
    Zljsonputvalue(v_List, 'create_time', r.�Ǽ�ʱ��);
    Zljsonputvalue(v_List, 'pati_type', r.��������);
    Zljsonputvalue(v_List, 'insurance_type', r.����, 1);
    Zljsonputvalue(v_List, 'adta_time', r.��Ժ����, 0, 2);
    If Length(v_List) > 20000 Then
      Json_Out := Json_Out || v_List;
      v_List   := ',';
    End If;
  End Loop;
  If v_List <> ',' Then
    Json_Out := Json_Out || v_List || ']}}';
  Else
    Json_Out := Json_Out || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getpativisitinfo;
/
Create Or Replace Procedure Zl_Cissvr_Getpativitalsigns
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡ��������������Ϣ
  --input      ��ȡ��������������Ϣ
  --  pati_id               N  1  ����ID
  --  visit_id              N  1  ����id �����ﲡ��Ϊ�Һ�ID;סԺ����Ϊ��ҳID;
  --  outpati_flag          N    �����־��1-���2-סԺ
  --output
  --  code                  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message               C  1  Ӧ����Ϣ��
  --  pativital_list[]      ������Ϣ��������Ŀ����ֵ����λ��[����]
  --     pativital_item     C  1  ��Ŀ
  --     pativital_value    C  1  ֵ
  --     pativital_unit     C  1  ��λ
  ---------------------------------------------------------------------------
  j_Input    Pljson;
  n_����id   Number(18);
  n_����id   Number(18);
  n_�����־ Number(1) := 1; --1-���2-סԺ

  v_���� Varchar2(100);
  v_��� Varchar2(100);

  j_In     Pljson;
  v_Output Varchar2(32767);
Begin
  j_In       := Pljson(Json_In);
  j_Input    := j_In.Get_Pljson('input');
  n_����id   := j_Input.Get_Number('pati_id');
  n_����id   := j_Input.Get_Number('visit_id');
  n_�����־ := j_Input.Get_Number('outpati_flag');

  If n_�����־ = 1 Then
    For R In (Select a.������, a.������ || '<;>' || b.ֵ || '<;>' || a.��λ As ֵ, a.������ As ��Ŀ, b.ֵ As ��Ŀֵ, a.��λ
              From (Select ������, ��λ
                     From ����������Ŀ
                     Where ����id = 7 And ������ In ('����', '����', '����ѹ', '����ѹ', '����', '���', '����', 'Ѫ��')) A,
                   (Select b.��Ŀ��λ, b.��Ŀ���� As ������, b.��¼���� As ֵ
                     From ���˻����¼ A, ���˻������� B
                     Where a.Id = b.��¼id And a.����id = n_����id And a.��ҳid = n_����id And
                           b.��Ŀ���� In ('����', '����', '����ѹ', '����ѹ', '����', '���', '����', 'Ѫ��')) B
              Where a.������ = b.������(+)) Loop
    
      Zljsonputvalue(v_Output, 'pativital_item', Nvl(r.��Ŀ, ''), 0, 1);
      Zljsonputvalue(v_Output, 'pativital_value', Nvl(r.��Ŀֵ, ''));
      Zljsonputvalue(v_Output, 'pativital_unit', Nvl(r.��λ, ''), 0, 2);
    End Loop;
  Else
    --ȡ������Ŀ���һ�εļ�¼
    For R In (Select a.������, a.������ || '<;>' || b.ֵ || '<;>' || a.��λ As ֵ, b.ֵ As ��¼����, a.������ As ��Ŀ, b.ֵ As ��Ŀֵ, a.��λ
              From (Select ������, ��λ
                     From ����������Ŀ
                     Where ����id = 7 And ������ In ('����', '����', '����ѹ', '����ѹ', '����', '���', '����', 'Ѫ��')) A,
                   (Select ��Ŀ���� As ������, ��¼���� As ֵ
                     From (Select ��Ŀ����, ��¼����, Row_Number() Over(Partition By ��Ŀ���� Order By ��¼ʱ�� Desc) Rn
                            From (Select c.��Ŀ����, c.��¼����, c.��¼ʱ��
                                   From ���˻����ļ� A, ���˻������� B, ���˻�����ϸ C
                                   Where a.����id = n_����id And a.��ҳid = n_����id And Nvl(a.Ӥ��, 0) = 0 And a.Id = b.�ļ�id And
                                         b.Id = c.��¼id
                                   Union All
                                   Select b.��Ŀ����, b.��¼����, b.�޸�ʱ�� As ��¼ʱ��
                                   From ���˻����¼ A, ���˻������� B
                                   Where a.����id = n_����id And a.��ҳid = n_����id And Nvl(a.Ӥ��, 0) = 0 And b.��¼���� = 1 And
                                         a.Id = b.��¼id))
                     Where Rn = 1) B
              Where a.������ = b.������(+)) Loop
    
      If r.������ = '����' Then
        If r.��¼���� Is Null Then
          v_���� := '';
        Else
          v_���� := r.ֵ;
        End If;
      Elsif r.������ = '���' Then
        If r.��¼���� Is Null Then
          v_��� := '';
        Else
          v_��� := r.ֵ;
        End If;
      End If;
    
      Zljsonputvalue(v_Output, 'pativital_item', Nvl(r.��Ŀ, ''), 0, 1);
      Zljsonputvalue(v_Output, 'pativital_value', Nvl(r.��Ŀֵ, ''));
      Zljsonputvalue(v_Output, 'pativital_unit', Nvl(r.��λ, ''), 0, 2);
    End Loop;
  End If;

  Json_Out := '{"output":{"code":1,"message": "�ɹ�","pativital_list":[' || v_Output || ']}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getpativitalsigns;
/
Create Or Replace Procedure Zl_Cissvr_Getpatpageinfbyrange
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:��ȡ������Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --    query_type          N 1 ��ѯ����:0-����;1-������չ
  --    wararea_ids         C   ����ids:����ö��ŷָ�
  --    dept_ids            C   ����IDs:����ö��ŷָ�
  --    pati_ids            C   ����ids:����ö��ŷָ�
  --    pati_pageIds        C   ��ҳIDs:����id:��ҳid,��
  --    adta_start_time     C   ��Ժ��ʼʱ��:yyyy-mm-dd hh24:mi:ss
  --    adta_end_time       C   ��Ժ����ʱ��:yyyy-mm-dd hh24:mi:ss
  --    adtd_start_time     C   ��Ժ��ʼʱ��:yyyy-mm-dd hh24:mi:ss
  --    adtd_end_time       C   ��Ժ����ʱ��:yyyy-mm-dd hh24:mi:ss
  --    fee_category        C   �ѱ�:����ö��ŷָ�
  --    inp_status          N   סԺ״̬:0-��Ժ����;1-��Ժ����;2-��Ժ���Ժ
  --    pati_natures        C   �������ʣ�����ö��ŷ�0-��ͨסԺ����,1-�������۲���,2-סԺ���۲��ˣ�NULL-��ʾ������
  --    pati_name           C   ����:���Դ�%�ֺű������ƥ��
  --    dept_nodeno         C   ����վ����
  --    change_dept_pati    N   �Ƿ��ѯת�Ʋ���
  --    is_lastpage         N   �Ƿ�ȡ���һ��סԺ 0-ȡ���� 1-ȡ���һ��
  --    insurance_type      N   �����������:>0:ָ������ҽ������,0:ҽ������ͨ����,-1:��ͨ����,-2:ҽ������
  --    wararea_nodeno      C   ����վ����
  --    mdlpay_mode_name    C   ҽ�Ƹ��ʽ
  --    fee_type            C   �ѱ�
  --    is_babyinfo         N 1 �Ƿ����Ӥ����Ϣ:1-����;0-������
  --����      json
  --output
  -- code                   N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  -- message                C 1 Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --   page_list[]          ������  ��  ��
  --    pati_id             N    ����id  ��  ��
  --    pati_pageid         N    ��ҳid  ��  ��
  --    pati_name           C    ����  ��  ��
  --    pati_sex            C    �Ա�  ��  ��
  --    pati_age            C    ����  ��  ��
  --    inpatient_num       C    סԺ��  ��  ��
  --    pati_bed            C    ��Ժ����  ��  ��
  --    insurance_type      N    ����  ��  ��
  --    insurance_type_name    C 1 ��������
  --    fee_category        C    �ѱ�  ��  ��
  --    pati_type           C    ��������(��ͨ,ҽ��,����)  ��  ��
  --    adta_time           C    ��Ժʱ��:yyyy-mm-dd hh24:mi:ss  ��  ��
  --    adtd_time           C    ��Ժʱ��:yyyy-mm-dd hh24:mi:ss  ��  ��
  --    si_inp_status       N    סԺ״̬:������ҳ.״̬(0-����סԺ��1-��δ��ƣ�2-����ת�ƻ�����ת������3-��Ԥ��Ժ)  ��  ��
  --    pati_nature         N    ��������:0-��ͨסԺ����,1-�������۲���,2-סԺ���۲���
  --    pati_wardarea_id    N    ��ǰ����id
  --    pati_wardarea_name  C    ��ǰ��������
  --    pati_dept_id        N    ��ǰ����id
  --    pati_dept_name      C    ��ǰ��������
  --    mdlpay_mode_name    C    ҽ�Ƹ��ʽ����
  --    mdlpay_mode_code    C    ҽ�Ƹ��ʽ����
  --    pat_rsdpscn         C    סԺҽʦ
  --    pati_desc           C    ���˱�ע
  --    catalog_date        C    ��Ŀ����:yyyy-mm-dd hh24:mi:ss
  --    create_pati         C    �Ǽ���
  --    in_objective        C    סԺĿ��
  --    insurance_num       C    ҽ����
  --    level_of_care       C    ����ȼ�
  --    data_adto_sign      N    ����ת����־:0-δת����1-��ת��
  --    fee_auditor_sign    N    ������˱�־:0���-δ���,1-����˻�ʼ���;2-������
  --    fee_auditor         C    ���������
  --    pre_dstat_time      C    Ԥ��Ժʱ��
  --    last_press_money    N    �ϴδ߿���
  --    baby_list[]           1 Ӥ����Ϣ��[����]
  --      baby_num          N 1 Ӥ�����
  --      baby_name         C 1 Ӥ������
  --      baby_sex          C 1 Ӥ���Ա�
  --      baby_date         C 1 ����ʱ��
  ---------------------------------------------------------------------------
  j_In       Pljson;
  j_Json     Pljson;
  n_��ѯ���� Number(2);
  v_����ids  Varchar2(32680);
  v_����ids  Varchar2(32680);
  v_����     Varchar2(200);
  v_����վ�� Zlnodelist.���%Type;
  v_����վ�� Zlnodelist.���%Type;
  n_����     ������ҳ.����%Type;

  n_Like Number(2);

  c_����ids Clob;
  c_��ҳids Clob;

  n_��Ӥ����Ϣ   Number(2);
  n_ת�Ʋ���     Number(2);
  d_��Ժ��ʼʱ�� Date;
  d_��Ժ����ʱ�� Date;
  d_��Ժ��ʼʱ�� Date;
  d_��Ժ����ʱ�� Date;
  v_�ѱ�         Varchar2(32680);
  n_סԺ״̬     Number(2);
  v_��������     Varchar2(100);
  l_����id       t_Strlist := t_Strlist();
  n_Last         Number;
  v_ҽ�Ƹ��ʽ Varchar2(100);
  n_Firstitem    Number(1); --�Ƿ��ǵ�һ����Ŀ

  l_��ҳid t_Strlist := t_Strlist();

  Cursor c_����������Ϣ Is
    Select a.����id, a.��ҳid, a.סԺ��, a.��������, a.�ѱ�, To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.��ǰ����id, a.��Ժ����id,
           a.��Ժ����, To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.סԺҽʦ,
           To_Char(a.��Ŀ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ŀ����, a.״̬, a.����, a.�Ա�, a.����, a.����, a.�Ǽ���, a.�Ǽ�ʱ��, a.��ע, a.��������,
           To_Char(a.Ԥ��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As Ԥ��Ժ����, c.���� As ��ǰ��������, d.���� As ��ǰ��������, e.���� As ҽ�Ƹ��ʽ����,
           a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, a.סԺĿ��, f.���� As ����ȼ�, a.����ת��, a.��˱�־, a.�����, x.���� As ��������
    
    From ������ҳ A, ���ű� C, ���ű� D, ҽ�Ƹ��ʽ E, �շ���ĿĿ¼ F, ������� X
    Where ����id = 0 And ��ҳid = 0 And a.��Ժ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.ҽ�Ƹ��ʽ = e.����(+) And
          a.����ȼ�id = f.Id(+) And a.���� = x.���(+) And Rownum < 1;
  Type Ty_������Ϣ Is Ref Cursor;
  c_������Ϣ Ty_������Ϣ; --��̬�α����
  --��װ��������
  Procedure Get_Jsonliststr
  (
    ������Ϣ_In  In Ty_������Ϣ,
    ��ѯ����_In  In Number,
    ��Ӥ��_In    In Number,
    Pati_Out     In Out Clob,
    Firstitem_In In Out Number
  ) As
    r_Pati   c_����������Ϣ%RowType;
    v_ҽ���� ������ҳ�ӱ�.��Ϣֵ%Type;
  
    n_�ϴδ߿��� ������ҳ�ӱ�.��Ϣֵ%Type;
  
    n_Firstitem Number(1);
    v_Temp      Varchar2(32767);
  
    n_Firstsubitem Number(1);
    v_Tempsub      Varchar2(32767);
  Begin
    n_Firstitem := Firstitem_In;
  
    Loop
      Fetch ������Ϣ_In
        Into r_Pati;
      Exit When ������Ϣ_In%NotFound;
    
      If Nvl(n_Firstitem, 0) = 1 Then
        v_Temp := v_Temp || ',';
      Else
        n_Firstitem := 1;
      End If;
      v_Temp := v_Temp || '{';
      v_Temp := v_Temp || '"pati_id":' || r_Pati.����id;
      v_Temp := v_Temp || ',"pati_pageid":' || Nvl(r_Pati.��ҳid || '', 'null');
      v_Temp := v_Temp || ',"pati_name":"' || Zljsonstr(r_Pati.����) || '"';
      v_Temp := v_Temp || ',"pati_sex":"' || Zljsonstr(r_Pati.�Ա�) || '"';
      v_Temp := v_Temp || ',"pati_age":"' || Zljsonstr(r_Pati.����) || '"';
    
      v_Temp := v_Temp || ',"inpatient_num":"' || Zljsonstr(r_Pati.סԺ��) || '"';
      v_Temp := v_Temp || ',"pati_bed":"' || Zljsonstr(r_Pati.��Ժ����) || '"';
      v_Temp := v_Temp || ',"insurance_type":' || Nvl(r_Pati.����, 0);
      v_Temp := v_Temp || ',"insurance_type_name":"' || Zljsonstr(r_Pati.��������) || '"';
      v_Temp := v_Temp || ',"fee_category":"' || Zljsonstr(r_Pati.�ѱ�) || '"';
      v_Temp := v_Temp || ',"pati_type":"' || Zljsonstr(r_Pati.��������) || '"';
    
      v_Temp := v_Temp || ',"adta_time":"' || Zljsonstr(r_Pati.��Ժʱ��) || '"';
      v_Temp := v_Temp || ',"adtd_time":"' || Zljsonstr(r_Pati.��Ժʱ��) || '"';
      v_Temp := v_Temp || ',"create_pati":"' || Zljsonstr(r_Pati.�Ǽ���) || '"';
      v_Temp := v_Temp || ',"create_time":"' || Zljsonstr(r_Pati.�Ǽ�ʱ��) || '"';
      v_Temp := v_Temp || ',"si_inp_status":' || Nvl(r_Pati.״̬, 0);
    
      v_Temp := v_Temp || ',"pati_nature":' || Nvl(r_Pati.��������, 0);
      v_Temp := v_Temp || ',"in_objective":"' || Zljsonstr(r_Pati.סԺĿ��) || '"';
    
      If Nvl(��ѯ����_In, 0) = 1 Then
        v_Temp := v_Temp || ',"pati_wardarea_id":' || Nvl(r_Pati.��ǰ����id, 0);
        v_Temp := v_Temp || ',"pati_wardarea_name":"' || Zljsonstr(r_Pati.��ǰ��������) || '"';
        v_Temp := v_Temp || ',"pati_dept_id":' || Nvl(r_Pati.��Ժ����id, 0);
        v_Temp := v_Temp || ',"pati_dept_name":"' || Zljsonstr(r_Pati.��ǰ��������) || '"';
      
        v_Temp := v_Temp || ',"mdlpay_mode_name":"' || Zljsonstr(r_Pati.ҽ�Ƹ��ʽ����) || '"';
        v_Temp := v_Temp || ',"mdlpay_mode_code":"' || Zljsonstr(r_Pati.ҽ�Ƹ��ʽ����) || '"';
        v_Temp := v_Temp || ',"pat_rsdpscn":"' || Zljsonstr(r_Pati.סԺҽʦ) || '"';
        v_Temp := v_Temp || ',"pati_desc":"' || Zljsonstr(r_Pati.��ע) || '"';
        v_Temp := v_Temp || ',"catalog_date":"' || Zljsonstr(r_Pati.��Ŀ����) || '"';
      
        Select Max(Decode(a.��Ϣ��, 'ҽ����', a.��Ϣֵ, 0)), Max(Decode(a.��Ϣ��, '�ϴδ߿���', a.��Ϣֵ, 0))
        Into v_ҽ����, n_�ϴδ߿���
        From ������ҳ�ӱ� A
        Where a.����id = r_Pati.����id And a.��ҳid = r_Pati.��ҳid And a.��Ϣ�� In ('ҽ����', '�ϴδ߿���');
      
        v_Temp := v_Temp || ',"insurance_num":"' || Zljsonstr(v_ҽ����) || '"';
        v_Temp := v_Temp || ',"level_of_care":"' || Zljsonstr(r_Pati.����ȼ�) || '"';
        v_Temp := v_Temp || ',"data_adto_sign":' || Nvl(r_Pati.����ת��, 0);
        v_Temp := v_Temp || ',"fee_auditor_sign":' || Nvl(r_Pati.��˱�־, 0);
        v_Temp := v_Temp || ',"fee_auditor":"' || Zljsonstr(r_Pati.�����) || '"';
        v_Temp := v_Temp || ',"pre_dstat_time":"' || Zljsonstr(r_Pati.Ԥ��Ժ����) || '"';
        v_Temp := v_Temp || ',"last_press_money":' || Nvl(n_�ϴδ߿���, 0);
      End If;
    
      If Nvl(��Ӥ��_In, 0) = 1 Then
        v_Tempsub      := '';
        n_Firstsubitem := 1;
        For r_Ӥ�� In (Select ����id, ��ҳid, ��� As Ӥ�����, Ӥ������, Ӥ���Ա�, To_Char(����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��
                     From ������������¼
                     Where ����id = r_Pati.����id And ��ҳid = r_Pati.��ҳid) Loop
        
          If Nvl(n_Firstsubitem, 0) = 0 Then
            v_Tempsub := v_Tempsub || ',';
          Else
            n_Firstsubitem := 0;
          End If;
        
          v_Tempsub := v_Tempsub || '{';
          v_Tempsub := v_Tempsub || '"baby_num":' || Nvl(r_Ӥ��.Ӥ�����, 0);
          v_Tempsub := v_Tempsub || ',"baby_name":"' || Zljsonstr(r_Ӥ��.Ӥ������) || '"';
          v_Tempsub := v_Tempsub || ',"baby_sex":"' || Zljsonstr(r_Ӥ��.Ӥ���Ա�) || '"';
          v_Tempsub := v_Tempsub || ',"baby_date":"' || Zljsonstr(r_Ӥ��.����ʱ��) || '"';
          v_Tempsub := v_Tempsub || '}';
        End Loop;
        v_Temp := v_Temp || ',"baby_list":[' || v_Tempsub || ']';
      End If;
      v_Temp := v_Temp || '}';
    
      If Length(v_Temp) > 20000 Then
        Dbms_Lob.Append(Pati_Out, v_Temp);
        v_Temp := '';
      End If;
    End Loop;
    Firstitem_In := n_Firstitem;
    If v_Temp Is Not Null Then
      Dbms_Lob.Append(Pati_Out, v_Temp);
    End If;
  End Get_Jsonliststr;
Begin

  --�������
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  n_��ѯ���� := j_Json.Get_Number('query_type');
  v_����ids  := j_Json.Get_String('wararea_ids');
  v_����ids  := j_Json.Get_String('dept_ids');
  If j_Json.Exist('pati_ids') Is Not Null Then
    c_����ids := j_Json.Get_Clob('pati_ids');
  End If;
  If j_Json.Exist('pati_pageids') Is Not Null Then
    c_��ҳids := j_Json.Get_Clob('pati_pageids');
  End If;

  d_��Ժ��ʼʱ�� := To_Date(j_Json.Get_String('adta_start_time'), 'YYYY-MM-DD hh24:mi:ss');
  d_��Ժ����ʱ�� := To_Date(j_Json.Get_String('adta_end_time'), 'YYYY-MM-DD hh24:mi:ss');
  d_��Ժ��ʼʱ�� := To_Date(j_Json.Get_String('adtd_start_time'), 'YYYY-MM-DD hh24:mi:ss');
  d_��Ժ����ʱ�� := To_Date(j_Json.Get_String('adtd_end_time'), 'YYYY-MM-DD hh24:mi:ss');

  v_�ѱ�         := j_Json.Get_String('fee_category');
  n_סԺ״̬     := Nvl(j_Json.Get_Number('inp_status'), 0);
  v_��������     := j_Json.Get_String('pati_natures');
  v_����         := j_Json.Get_String('pati_name');
  v_����վ��     := j_Json.Get_String('dept_nodeno');
  n_ת�Ʋ���     := Nvl(j_Json.Get_Number('change_dept_pati'), 0);
  n_Last         := j_Json.Get_Number('is_lastpage');
  n_����         := j_Json.Get_Number('insurance_type');
  v_����վ��     := j_Json.Get_String('wararea_nodeno');
  v_ҽ�Ƹ��ʽ := j_Json.Get_String('mdlpay_mode_name');
  n_��Ӥ����Ϣ   := j_Json.Get_Number('is_babyinfo');

  If v_�������� Is Null Then
    v_�������� := ',0,1,2,';
  Else
    v_�������� := ',' || v_�������� || ',';
  End If;

  If v_����ids Is Not Null Then
    v_����ids := ',' || v_����ids || ',';
  End If;

  If v_�ѱ� Is Not Null Then
    v_�ѱ� := ',' || v_�ѱ� || ',';
  End If;

  If v_����ids Is Not Null Then
    v_����ids := ',' || v_����ids || ',';
  End If;

  n_Like := 0;
  If Instr(Nvl(v_����, '-'), '%') > 0 Then
    n_Like := 1;
  End If;

  While c_����ids Is Not Null Loop
    If Length(c_����ids) <= 4000 Then
      l_����id.Extend;
      l_����id(l_����id.Count) := c_����ids;
      c_����ids := Null;
    Else
      l_����id.Extend;
      l_����id(l_����id.Count) := Substr(c_����ids, 1, Instr(c_����ids, ',', 3950) - 1);
      c_����ids := Substr(c_����ids, Instr(c_����ids, ',', 3950) + 1);
    End If;
  End Loop;

  While c_��ҳids Is Not Null Loop
    If Length(c_��ҳids) <= 4000 Then
      l_��ҳid.Extend;
      l_��ҳid(l_��ҳid.Count) := c_��ҳids;
      c_��ҳids := Null;
    Else
      l_��ҳid.Extend;
      l_��ҳid(l_��ҳid.Count) := Substr(c_��ҳids, 1, Instr(c_��ҳids, ',', 3950) - 1);
      c_��ҳids := Substr(c_��ҳids, Instr(c_��ҳids, ',', 3950) + 1);
    End If;
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","page_list":[';
  If Nvl(n_ת�Ʋ���, 0) = 1 Then
    --ת�Ʋ���
    If l_����id.Count <> 0 Then
      For I In 1 .. l_����id.Count Loop
        --���ܴ���ͬһ����һ���Χ�ڵ����������ϵ�ת��,�������һ��Ϊ׼.
        Open c_������Ϣ For
          Select /*+cardinality(f,10)*/
           a.����id, a.��ҳid, a.סԺ��, a.��������, a.�ѱ�, To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.��ǰ����id, a.��Ժ����id,
           a.��Ժ����, To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.סԺҽʦ,
           To_Char(a.��Ŀ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ŀ����, a.״̬, a.����, a.�Ա�, a.����, a.����, a.�Ǽ���, a.�Ǽ�ʱ��, a.��ע, a.��������,
           To_Char(a.Ԥ��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As Ԥ��Ժ����, c.���� As ��ǰ��������, d.���� As ��ǰ��������, e.���� As ҽ�Ƹ��ʽ����,
           a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, a.סԺĿ��, Null As ����ȼ�, a.����ת��, a.��˱�־, a.�����, x.���� As ��������
          
          From ������ҳ A, ���˱䶯��¼ B, ���ű� C, ���ű� D, ҽ�Ƹ��ʽ E, Table(f_Num2list(l_����id(I))) F, ������� X
          Where a.����id = b.����id And a.��ҳid = b.��ҳid And b.����id = c.Id(+) And b.����id = d.Id(+) And a.ҽ�Ƹ��ʽ = e.����(+) And
                a.���� = x.���(+) And a.����id = f.Column_Value
               
                And Instr(v_��������, ',' || a.�������� || ',') > 0 And Nvl(a.��ҳid, 0) <> 0 And Nvl(a.״̬, 0) <> 2
               
                And (v_���� Is Null Or (n_Like = 1 And a.���� Like v_����) Or (n_Like = 0 And a.���� = v_����))
               
                And (v_����ids Is Null Or Instr(v_����ids, ',' || a.��ǰ����id || ',') = 0) And
                (v_����ids Is Null Or Instr(v_����ids, ',' || b.����id || ',') > 0)
               
                And (v_����ids Is Null Or Instr(v_����ids, ',' || a.��Ժ����id || ',') = 0) And
                (v_����ids Is Null Or Instr(v_����ids, ',' || b.����id || ',') > 0)
               
                And b.��ֹԭ�� = 3 And (b.��ֹʱ�� >= d_��Ժ��ʼʱ�� Or d_��Ժ��ʼʱ�� Is Null) And
                (b.��ֹʱ�� <= d_��Ժ����ʱ�� Or d_��Ժ����ʱ�� Is Null) And Nvl(b.���Ӵ�λ, 0) = 0 And Nvl(a.����״̬, 0) <> 5 And
                a.���ʱ�� Is Null And
                b.��ֹʱ�� = (Select Max(��ֹʱ��)
                          From ���˱䶯��¼
                          Where ����id = b.����id And ��ҳid = b.��ҳid And ��ֹԭ�� = 3 And Nvl(���Ӵ�λ, 0) = 0 And
                                (v_����ids Is Null Or Instr(v_����ids, ',' || ����id || ',') > 0) And
                                (v_����ids Is Null Or Instr(v_����ids, ',' || ����id || ',') > 0));
      
        Get_Jsonliststr(c_������Ϣ, n_��ѯ����, n_��Ӥ����Ϣ, Json_Out, n_Firstitem);
        If n_Firstitem = 1 Then
          n_Firstitem := 0;
        End If;
      End Loop;
    
    Else
      --���ܴ���ͬһ����һ���Χ�ڵ����������ϵ�ת��,�������һ��Ϊ׼.
      Open c_������Ϣ For
        Select a.����id, a.��ҳid, a.סԺ��, a.��������, a.�ѱ�, To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.��ǰ����id, a.��Ժ����id,
               a.��Ժ����, To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.סԺҽʦ,
               To_Char(a.��Ŀ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ŀ����, a.״̬, a.����, a.�Ա�, a.����, a.����, a.�Ǽ���, a.�Ǽ�ʱ��, a.��ע,
               a.��������, To_Char(a.Ԥ��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As Ԥ��Ժ����, c.���� As ��ǰ��������, d.���� As ��ǰ��������,
               e.���� As ҽ�Ƹ��ʽ����, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, a.סԺĿ��, Null As ����ȼ�, a.����ת��, a.��˱�־, a.�����, x.���� As ��������
        From ������ҳ A, ���˱䶯��¼ B, ���ű� C, ���ű� D, ҽ�Ƹ��ʽ E, ������� X
        Where a.����id = b.����id And a.��ҳid = b.��ҳid And b.����id = c.Id(+) And b.����id = d.Id(+) And a.ҽ�Ƹ��ʽ = e.����(+) And
              a.���� = x.���(+)
             
              And Instr(v_��������, ',' || a.�������� || ',') > 0 And Nvl(a.��ҳid, 0) <> 0 And Nvl(a.״̬, 0) <> 2
             
              And (v_���� Is Null Or (n_Like = 1 And a.���� Like v_����) Or (n_Like = 0 And a.���� = v_����))
             
              And (v_����ids Is Null Or Instr(v_����ids, ',' || a.��ǰ����id || ',') = 0) And
              (v_����ids Is Null Or Instr(v_����ids, ',' || b.����id || ',') > 0)
             
              And (v_����ids Is Null Or Instr(v_����ids, ',' || a.��Ժ����id || ',') = 0) And
              (v_����ids Is Null Or Instr(v_����ids, ',' || b.����id || ',') > 0)
             
              And b.��ֹԭ�� = 3 And (b.��ֹʱ�� >= d_��Ժ��ʼʱ�� Or d_��Ժ��ʼʱ�� Is Null) And (b.��ֹʱ�� <= d_��Ժ����ʱ�� Or d_��Ժ����ʱ�� Is Null) And
              Nvl(b.���Ӵ�λ, 0) = 0 And Nvl(a.����״̬, 0) <> 5 And a.���ʱ�� Is Null And
              b.��ֹʱ�� = (Select Max(��ֹʱ��)
                        From ���˱䶯��¼
                        Where ����id = b.����id And ��ҳid = b.��ҳid And ��ֹԭ�� = 3 And Nvl(���Ӵ�λ, 0) = 0 And
                              (v_����ids Is Null Or Instr(v_����ids, ',' || ����id || ',') > 0) And
                              (v_����ids Is Null Or Instr(v_����ids, ',' || ����id || ',') > 0));
      Get_Jsonliststr(c_������Ϣ, n_��ѯ����, n_��Ӥ����Ϣ, Json_Out, n_Firstitem);
    End If;
  Elsif l_����id.Count <> 0 Then
    For I In 1 .. l_����id.Count Loop
      Open c_������Ϣ For
        Select /*+cardinality(B,10)*/
         a.����id, a.��ҳid, a.סԺ��, a.��������, a.�ѱ�, To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.��ǰ����id, a.��Ժ����id,
         a.��Ժ����, To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.סԺҽʦ,
         To_Char(a.��Ŀ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ŀ����, a.״̬, a.����, a.�Ա�, a.����, a.����, a.�Ǽ���, a.�Ǽ�ʱ��, a.��ע, a.��������,
         To_Char(a.Ԥ��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As Ԥ��Ժ����, c.���� As ��ǰ��������, d.���� As ��ǰ��������, e.���� As ҽ�Ƹ��ʽ����,
         a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, a.סԺĿ��, Null As ����ȼ�, a.����ת��, a.��˱�־, a.�����, x.���� As ��������
        From ������ҳ A, (Select Column_Value As ����id From Table(f_Num2list(l_����id(I)))) B, ���ű� C, ���ű� D, ҽ�Ƹ��ʽ E, ��Ժ���� H,
             ������� X
        Where a.��Ժ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.����id = b.����id And a.���� = x.���(+) And
              (v_���� Is Null Or (n_Like = 1 And a.���� Like v_����) Or (n_Like = 0 And a.���� = v_����)) And
              (v_����ids Is Null Or Instr(v_����ids, ',' || a.��ǰ����id || ',') > 0) And
              (v_����ids Is Null Or Instr(v_����ids, ',' || a.��Ժ����id || ',') > 0) And
              (v_�ѱ� Is Null Or Instr(v_�ѱ�, ',' || a.�ѱ� || ',') > 0) And Instr(v_��������, ',' || a.�������� || ',') > 0 And
              a.����id = h.����id(+) And a.��ҳid = h.��ҳid(+) And
              (Nvl(v_ҽ�Ƹ��ʽ, '-') = '-' Or ((Nvl(n_סԺ״̬, 0) = 0 Or Nvl(n_סԺ״̬, 0) = 1) And a.ҽ�Ƹ��ʽ = v_ҽ�Ƹ��ʽ)) And
             
              (n_סԺ״̬ = 0 And h.����id Is Not Null And (a.��Ժ���� >= d_��Ժ��ʼʱ�� Or d_��Ժ��ʼʱ�� Is Null) And
              (a.��Ժ���� <= d_��Ժ����ʱ�� Or d_��Ժ����ʱ�� Is Null) Or
              
              n_סԺ״̬ = 1 And h.����id Is Null And (a.��Ժ���� >= d_��Ժ��ʼʱ�� Or d_��Ժ��ʼʱ�� Is Null) And
              (a.��Ժ���� <= d_��Ժ����ʱ�� Or d_��Ժ����ʱ�� Is Null) Or
              
              n_סԺ״̬ = 2 And
              (h.����id Is Not Null And (a.��Ժ���� >= d_��Ժ��ʼʱ�� Or d_��Ժ��ʼʱ�� Is Null) And
              (a.��Ժ���� <= d_��Ժ����ʱ�� Or d_��Ժ����ʱ�� Is Null) Or h.����id Is Null And (a.��Ժ���� >= d_��Ժ��ʼʱ�� Or d_��Ժ��ʼʱ�� Is Null) And
              (a.��Ժ���� <= d_��Ժ����ʱ�� Or d_��Ժ����ʱ�� Is Null)))
             
              And a.ҽ�Ƹ��ʽ = e.����(+) And
              (a.��ҳid = (Select Max(��ҳid) From ������ҳ Where ����id = a.����id) And Nvl(n_Last, 0) = 1 Or Nvl(n_Last, 0) = 0) And
              (Nvl(n_����, 0) = 0 Or n_���� > 0 And a.���� = n_���� Or n_���� = -1 And Nvl(a.����, 0) = 0 Or
              n_���� = -2 And Nvl(a.����, 0) <> 0) And (c.վ�� Is Null Or c.վ�� = v_����վ�� Or v_����վ�� Is Null) And
              (d.վ�� Is Null Or d.վ�� = v_����վ�� Or v_����վ�� Is Null);
    
      Get_Jsonliststr(c_������Ϣ, n_��ѯ����, n_��Ӥ����Ϣ, Json_Out, n_Firstitem);
    End Loop;
  Elsif l_��ҳid.Count <> 0 Then
    For I In 1 .. l_��ҳid.Count Loop
      Open c_������Ϣ For
        With c_���� As
         (Select Distinct C1 As ����id, C2 As ��ҳid From Table(f_Num2list2(l_��ҳid(I))))
        Select /*+cardinality(B,10)*/
         a.����id, a.��ҳid, a.סԺ��, a.��������, a.�ѱ�, To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.��ǰ����id, a.��Ժ����id,
         a.��Ժ����, To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.סԺҽʦ,
         To_Char(a.��Ŀ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ŀ����, a.״̬, a.����, a.�Ա�, a.����, a.����, a.�Ǽ���, a.�Ǽ�ʱ��, a.��ע, a.��������,
         To_Char(a.Ԥ��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As Ԥ��Ժ����, c.���� As ��ǰ��������, d.���� As ��ǰ��������, e.���� As ҽ�Ƹ��ʽ����,
         a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, a.סԺĿ��, Null As ����ȼ�, a.����ת��, a.��˱�־, a.�����, x.���� As ��������
        From ������ҳ A, c_���� B, ���ű� C, ���ű� D, ҽ�Ƹ��ʽ E, ��Ժ���� H, ������� X
        Where a.��Ժ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.����id = b.����id And a.��ҳid = b.��ҳid And a.���� = x.���(+) And
              (v_���� Is Null Or (n_Like = 1 And a.���� Like v_����) Or (n_Like = 0 And a.���� = v_����)) And
              (v_����ids Is Null Or Instr(v_����ids, ',' || a.��ǰ����id || ',') > 0) And
              (v_����ids Is Null Or Instr(v_����ids, ',' || a.��Ժ����id || ',') > 0) And
              (v_�ѱ� Is Null Or Instr(v_�ѱ�, ',' || a.�ѱ� || ',') > 0) And Instr(v_��������, ',' || a.�������� || ',') > 0 And
              a.����id = h.����id(+) And a.��ҳid = h.��ҳid(+) And
              (Nvl(v_ҽ�Ƹ��ʽ, '-') = '-' Or ((Nvl(n_סԺ״̬, 0) = 0 Or Nvl(n_סԺ״̬, 0) = 1) And a.ҽ�Ƹ��ʽ = v_ҽ�Ƹ��ʽ)) And
             
              (n_סԺ״̬ = 0 And h.����id Is Not Null And (a.��Ժ���� >= d_��Ժ��ʼʱ�� Or d_��Ժ��ʼʱ�� Is Null) And
              (a.��Ժ���� <= d_��Ժ����ʱ�� Or d_��Ժ����ʱ�� Is Null) Or
              
              n_סԺ״̬ = 1 And h.����id Is Null And (a.��Ժ���� >= d_��Ժ��ʼʱ�� Or d_��Ժ��ʼʱ�� Is Null) And
              (a.��Ժ���� <= d_��Ժ����ʱ�� Or d_��Ժ����ʱ�� Is Null) Or
              
              n_סԺ״̬ = 2 And
              (h.����id Is Not Null And (a.��Ժ���� >= d_��Ժ��ʼʱ�� Or d_��Ժ��ʼʱ�� Is Null) And
              (a.��Ժ���� <= d_��Ժ����ʱ�� Or d_��Ժ����ʱ�� Is Null) Or h.����id Is Null And (a.��Ժ���� >= d_��Ժ��ʼʱ�� Or d_��Ժ��ʼʱ�� Is Null) And
              (a.��Ժ���� <= d_��Ժ����ʱ�� Or d_��Ժ����ʱ�� Is Null))) And a.ҽ�Ƹ��ʽ = e.����(+) And
              (Nvl(n_����, 0) = 0 Or n_���� > 0 And a.���� = n_���� Or n_���� = -1 And Nvl(a.����, 0) = 0 Or
              n_���� = -2 And Nvl(a.����, 0) <> 0) And (c.վ�� Is Null Or c.վ�� = v_����վ�� Or v_����վ�� Is Null) And
              (d.վ�� Is Null Or d.վ�� = v_����վ�� Or v_����վ�� Is Null);
    
      Get_Jsonliststr(c_������Ϣ, n_��ѯ����, n_��Ӥ����Ϣ, Json_Out, n_Firstitem);
    End Loop;
  Elsif d_��Ժ��ʼʱ�� Is Not Null Then
    Open c_������Ϣ For
      Select a.����id, a.��ҳid, a.סԺ��, a.��������, a.�ѱ�, To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.��ǰ����id, a.��Ժ����id,
             a.��Ժ����, To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.סԺҽʦ,
             To_Char(a.��Ŀ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ŀ����, a.״̬, a.����, a.�Ա�, a.����, a.����, a.�Ǽ���, a.�Ǽ�ʱ��, a.��ע, a.��������,
             To_Char(a.Ԥ��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As Ԥ��Ժ����, c.���� As ��ǰ��������, d.���� As ��ǰ��������, e.���� As ҽ�Ƹ��ʽ����,
             a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, a.סԺĿ��, Null As ����ȼ�, a.����ת��, a.��˱�־, a.�����, x.���� As ��������
      
      From ������ҳ A, ���ű� C, ���ű� D, ҽ�Ƹ��ʽ E, ��Ժ���� H, ������� X
      Where a.��Ժ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And
            (v_���� Is Null Or (n_Like = 1 And a.���� Like v_����) Or (n_Like = 0 And a.���� = v_����)) And a.���� = x.���(+) And
            (v_����ids Is Null Or Instr(v_����ids, ',' || a.��ǰ����id || ',') > 0) And
            (v_����ids Is Null Or Instr(v_����ids, ',' || a.��Ժ����id || ',') > 0) And Instr(v_��������, ',' || a.�������� || ',') > 0 And
            (v_�ѱ� Is Null Or Instr(v_�ѱ�, ',' || a.�ѱ� || ',') > 0) And a.����id = h.����id(+) And a.��ҳid = h.��ҳid(+) And
            (Nvl(v_ҽ�Ƹ��ʽ, '-') = '-' Or ((Nvl(n_סԺ״̬, 0) = 0 Or Nvl(n_סԺ״̬, 0) = 1) And a.ҽ�Ƹ��ʽ = v_ҽ�Ƹ��ʽ)) And
            (n_סԺ״̬ = 0 And h.����id Is Not Null And (a.��Ժ���� >= d_��Ժ��ʼʱ�� Or d_��Ժ��ʼʱ�� Is Null) And
            (a.��Ժ���� <= d_��Ժ����ʱ�� Or d_��Ժ����ʱ�� Is Null) Or
            
            n_סԺ״̬ = 1 And h.����id Is Null And (a.��Ժ���� >= d_��Ժ��ʼʱ�� Or d_��Ժ��ʼʱ�� Is Null) And
            (a.��Ժ���� <= d_��Ժ����ʱ�� Or d_��Ժ����ʱ�� Is Null) Or
            
            n_סԺ״̬ = 2 And
            (h.����id Is Not Null And (a.��Ժ���� >= d_��Ժ��ʼʱ�� Or d_��Ժ��ʼʱ�� Is Null) And
            (a.��Ժ���� <= d_��Ժ����ʱ�� Or d_��Ժ����ʱ�� Is Null) Or
            h.����id Is Null And (a.��Ժ���� >= d_��Ժ��ʼʱ�� Or d_��Ժ��ʼʱ�� Is Null) And (a.��Ժ���� <= d_��Ժ����ʱ�� Or d_��Ժ����ʱ�� Is Null))) And
            a.ҽ�Ƹ��ʽ = e.����(+) And
            (a.��ҳid = (Select Max(��ҳid) From ������ҳ Where ����id = a.����id) And Nvl(n_Last, 0) = 1 Or Nvl(n_Last, 0) = 0) And
            (c.վ�� Is Null Or c.վ�� = v_����վ�� Or v_����վ�� Is Null) And
            (Nvl(n_����, 0) = 0 Or n_���� > 0 And a.���� = n_���� Or n_���� = -1 And Nvl(a.����, 0) = 0 Or
            n_���� = -2 And Nvl(a.����, 0) <> 0) And (d.վ�� Is Null Or d.վ�� = v_����վ�� Or v_����վ�� Is Null);
    Get_Jsonliststr(c_������Ϣ, n_��ѯ����, n_��Ӥ����Ϣ, Json_Out, n_Firstitem);
  
  Elsif d_��Ժ��ʼʱ�� Is Not Null Then
    Open c_������Ϣ For
      Select a.����id, a.��ҳid, a.סԺ��, a.��������, a.�ѱ�, To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.��ǰ����id, a.��Ժ����id,
             a.��Ժ����, To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.סԺҽʦ,
             To_Char(a.��Ŀ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ŀ����, a.״̬, a.����, a.�Ա�, a.����, a.����, a.�Ǽ���, a.�Ǽ�ʱ��, a.��ע, a.��������,
             To_Char(a.Ԥ��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As Ԥ��Ժ����, c.���� As ��ǰ��������, d.���� As ��ǰ��������, e.���� As ҽ�Ƹ��ʽ����,
             a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, a.סԺĿ��, '' As ����ȼ�, a.����ת��, a.��˱�־, a.�����, x.���� As ��������
      
      From ������ҳ A, ���ű� C, ���ű� D, ҽ�Ƹ��ʽ E, ��Ժ���� H, ������� X
      Where a.��Ժ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And
            (v_���� Is Null Or (n_Like = 1 And a.���� Like v_����) Or (n_Like = 0 And a.���� = v_����)) And a.���� = x.���(+) And
            (v_����ids Is Null Or Instr(v_����ids, ',' || a.��ǰ����id || ',') > 0) And
            (v_����ids Is Null Or Instr(v_����ids, ',' || a.��Ժ����id || ',') > 0) And Instr(v_��������, ',' || a.�������� || ',') > 0 And
            (v_�ѱ� Is Null Or Instr(v_�ѱ�, ',' || a.�ѱ� || ',') > 0) And a.����id = h.����id(+) And a.��ҳid = h.��ҳid(+) And
            (Nvl(v_ҽ�Ƹ��ʽ, '-') = '-' Or ((Nvl(n_סԺ״̬, 0) = 0 Or Nvl(n_סԺ״̬, 0) = 1) And a.ҽ�Ƹ��ʽ = v_ҽ�Ƹ��ʽ)) And
            (n_סԺ״̬ = 0 And h.����id Is Not Null And (a.��Ժ���� >= d_��Ժ��ʼʱ�� Or d_��Ժ��ʼʱ�� Is Null) And
            (a.��Ժ���� <= d_��Ժ����ʱ�� Or d_��Ժ����ʱ�� Is Null) Or
            
            n_סԺ״̬ = 1 And h.����id Is Null And (a.��Ժ���� >= d_��Ժ��ʼʱ�� Or d_��Ժ��ʼʱ�� Is Null) And
            (a.��Ժ���� <= d_��Ժ����ʱ�� Or d_��Ժ����ʱ�� Is Null) Or
            
            n_סԺ״̬ = 2 And
            (h.����id Is Not Null And (a.��Ժ���� >= d_��Ժ��ʼʱ�� Or d_��Ժ��ʼʱ�� Is Null) And
            (a.��Ժ���� <= d_��Ժ����ʱ�� Or d_��Ժ����ʱ�� Is Null) Or
            h.����id Is Null And (a.��Ժ���� >= d_��Ժ��ʼʱ�� Or d_��Ժ��ʼʱ�� Is Null) And (a.��Ժ���� <= d_��Ժ����ʱ�� Or d_��Ժ����ʱ�� Is Null))) And
            a.ҽ�Ƹ��ʽ = e.����(+) And
            (a.��ҳid = (Select Max(��ҳid) From ������ҳ Where ����id = a.����id) And Nvl(n_Last, 0) = 1 Or Nvl(n_Last, 0) = 0) And
            (c.վ�� Is Null Or c.վ�� = v_����վ�� Or v_����վ�� Is Null) And
            (Nvl(n_����, 0) = 0 Or n_���� > 0 And a.���� = n_���� Or n_���� = -1 And Nvl(a.����, 0) = 0 Or
            n_���� = -2 And Nvl(a.����, 0) <> 0) And (d.վ�� Is Null Or d.վ�� = v_����վ�� Or v_����վ�� Is Null);
    Get_Jsonliststr(c_������Ϣ, n_��ѯ����, n_��Ӥ����Ϣ, Json_Out, n_Firstitem);
  
  Elsif Nvl(n_סԺ״̬, 0) = 0 Then
    --ֻȡ��Ժ����
    Open c_������Ϣ For
      Select a.����id, a.��ҳid, a.סԺ��, a.��������, a.�ѱ�, To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.��ǰ����id, a.��Ժ����id,
             a.��Ժ����, To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.סԺҽʦ,
             To_Char(a.��Ŀ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ŀ����, a.״̬, a.����, a.�Ա�, a.����, a.����, a.�Ǽ���, a.�Ǽ�ʱ��, a.��ע, a.��������,
             To_Char(a.Ԥ��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As Ԥ��Ժ����, c.���� As ��ǰ��������, d.���� As ��ǰ��������, e.���� As ҽ�Ƹ��ʽ����,
             a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, a.סԺĿ��, f.���� As ����ȼ�, a.����ת��, a.��˱�־, a.�����, x.���� As ��������
      
      From ������ҳ A, ��Ժ���� B, ���ű� C, ���ű� D, ҽ�Ƹ��ʽ E, �շ���ĿĿ¼ F, ������� X
      Where a.��Ժ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.����id = b.����id And a.��ҳid = b.��ҳid And a.���� = x.���(+) And
            (v_���� Is Null Or (n_Like = 1 And a.���� Like v_����) Or (n_Like = 0 And a.���� = v_����)) And
            (v_�ѱ� Is Null Or Instr(v_�ѱ�, ',' || a.�ѱ� || ',') > 0) And Instr(v_��������, ',' || a.�������� || ',') > 0 And
            (Nvl(v_ҽ�Ƹ��ʽ, '-') = '-' Or ((Nvl(n_סԺ״̬, 0) = 0 Or Nvl(n_סԺ״̬, 0) = 1) And a.ҽ�Ƹ��ʽ = v_ҽ�Ƹ��ʽ)) And
            (v_����ids Is Null Or Instr(v_����ids, ',' || a.��ǰ����id || ',') > 0 Or
            Nvl(n_��Ӥ����Ϣ, 0) = 1 And Instr(v_����ids, ',' || a.Ӥ������id || ',') > 0) And
            (v_����ids Is Null Or Instr(v_����ids, ',' || a.��Ժ����id || ',') > 0) And a.ҽ�Ƹ��ʽ = e.����(+) And
            (a.��ҳid = (Select Max(��ҳid) From ������ҳ Where ����id = a.����id) And Nvl(n_Last, 0) = 1 Or Nvl(n_Last, 0) = 0) And
            (c.վ�� Is Null Or c.վ�� = v_����վ�� Or v_����վ�� Is Null) And
            (Nvl(n_����, 0) = 0 Or n_���� > 0 And a.���� = n_���� Or n_���� = -1 And Nvl(a.����, 0) = 0 Or
            n_���� = -2 And Nvl(a.����, 0) <> 0) And a.����ȼ�id = f.Id(+) And
            (d.վ�� Is Null Or d.վ�� = v_����վ�� Or v_����վ�� Is Null);
    Get_Jsonliststr(c_������Ϣ, n_��ѯ����, n_��Ӥ����Ϣ, Json_Out, n_Firstitem);
  
  Elsif v_����ids Is Not Null Then
    Open c_������Ϣ For
      With c_���� As
       (Select Column_Value As ����id From Table(f_Num2list(v_����ids)))
      Select /*+cardinality(B,10)*/
       a.����id, a.��ҳid, a.סԺ��, a.��������, a.�ѱ�, To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.��ǰ����id, a.��Ժ����id, a.��Ժ����,
       To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.סԺҽʦ, To_Char(a.��Ŀ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ŀ����, a.״̬,
       a.����, a.�Ա�, a.����, a.����, a.�Ǽ���, a.�Ǽ�ʱ��, a.��ע, a.��������, To_Char(a.Ԥ��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As Ԥ��Ժ����,
       c.���� As ��ǰ��������, d.���� As ��ǰ��������, e.���� As ҽ�Ƹ��ʽ����, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, a.סԺĿ��, Null As ����ȼ�, a.����ת��, a.��˱�־,
       a.�����, x.���� As ��������
      From ������ҳ A, c_���� B, ���ű� C, ���ű� D, ҽ�Ƹ��ʽ E, ��Ժ���� H, ������� X
      Where a.��Ժ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And
            (a.��ǰ����id = b.����id Or Nvl(n_��Ӥ����Ϣ, 0) = 1 And a.Ӥ������id = b.����id) And a.���� = x.���(+) And
            (v_���� Is Null Or (n_Like = 1 And a.���� Like v_����) Or (n_Like = 0 And a.���� = v_����)) And
            (v_����ids Is Null Or Instr(v_����ids, ',' || a.��Ժ����id || ',') > 0) And Instr(v_��������, ',' || a.�������� || ',') > 0 And
            (v_�ѱ� Is Null Or Instr(v_�ѱ�, ',' || a.�ѱ� || ',') > 0) And a.����id = h.����id(+) And a.��ҳid = h.��ҳid(+) And
            (Nvl(v_ҽ�Ƹ��ʽ, '-') = '-' Or ((Nvl(n_סԺ״̬, 0) = 0 Or Nvl(n_סԺ״̬, 0) = 1) And a.ҽ�Ƹ��ʽ = v_ҽ�Ƹ��ʽ)) And
            a.ҽ�Ƹ��ʽ = e.����(+) And
            (a.��ҳid = (Select Max(��ҳid) From ������ҳ Where ����id = a.����id) And Nvl(n_Last, 0) = 1 Or Nvl(n_Last, 0) = 0) And
            (c.վ�� Is Null Or c.վ�� = v_����վ�� Or v_����վ�� Is Null) And
            (Nvl(n_����, 0) = 0 Or n_���� > 0 And a.���� = n_���� Or n_���� = -1 And Nvl(a.����, 0) = 0 Or
            n_���� = -2 And Nvl(a.����, 0) <> 0) And (d.վ�� Is Null Or d.վ�� = v_����վ�� Or v_����վ�� Is Null);
    Get_Jsonliststr(c_������Ϣ, n_��ѯ����, n_��Ӥ����Ϣ, Json_Out, n_Firstitem);
  
  Else
    Open c_������Ϣ For
      Select a.����id, a.��ҳid, a.סԺ��, a.��������, a.�ѱ�, To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.��ǰ����id, a.��Ժ����id,
             a.��Ժ����, To_Char(a.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.סԺҽʦ,
             To_Char(a.��Ŀ����, 'yyyy-mm-dd hh24:mi:ss') As ��Ŀ����, a.״̬, a.����, a.�Ա�, a.����, a.����, a.�Ǽ���, a.�Ǽ�ʱ��, a.��ע, a.��������,
             To_Char(a.Ԥ��Ժ����, 'yyyy-mm-dd hh24:mi:ss') As Ԥ��Ժ����, c.���� As ��ǰ��������, d.���� As ��ǰ��������, e.���� As ҽ�Ƹ��ʽ����,
             a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, a.סԺĿ��, Null As ����ȼ�, a.����ת��, a.��˱�־, a.�����, x.���� As ��������
      
      From ������ҳ A, ���ű� C, ���ű� D, ҽ�Ƹ��ʽ E, ��Ժ���� H, ������� X
      Where a.��Ժ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And
            (v_���� Is Null Or (n_Like = 1 And a.���� Like v_����) Or (n_Like = 0 And a.���� = v_����)) And a.���� = x.���(+) And
            (v_����ids Is Null Or Instr(v_����ids, ',' || a.��Ժ����id || ',') > 0) And Instr(v_��������, ',' || a.�������� || ',') > 0 And
            (v_�ѱ� Is Null Or Instr(v_�ѱ�, ',' || a.�ѱ� || ',') > 0) And a.����id = h.����id(+) And a.��ҳid = h.��ҳid(+) And
            a.ҽ�Ƹ��ʽ = e.����(+) And
            (Nvl(v_ҽ�Ƹ��ʽ, '-') = '-' Or ((Nvl(n_סԺ״̬, 0) = 0 Or Nvl(n_סԺ״̬, 0) = 1) And a.ҽ�Ƹ��ʽ = v_ҽ�Ƹ��ʽ)) And
            (a.��ҳid = (Select Max(��ҳid) From ������ҳ Where ����id = a.����id) And Nvl(n_Last, 0) = 1 Or Nvl(n_Last, 0) = 0) And
            (c.վ�� Is Null Or c.վ�� = v_����վ�� Or v_����վ�� Is Null) And
            (Nvl(n_����, 0) = 0 Or n_���� > 0 And a.���� = n_���� Or n_���� = -1 And Nvl(a.����, 0) = 0 Or
            n_���� = -2 And Nvl(a.����, 0) <> 0) And (d.վ�� Is Null Or d.վ�� = v_����վ�� Or v_����վ�� Is Null);
    Get_Jsonliststr(c_������Ϣ, n_��ѯ����, n_��Ӥ����Ϣ, Json_Out, n_Firstitem);
  
  End If;

  Dbms_Lob.Append(Json_Out, ']}}');
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getpatpageinfbyrange;
/
Create Or Replace Procedure Zl_Cissvr_Getpatpagewarnscheme
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:��ȡ���˵���������
  --��Σ�Json_In:��ʽ
  --    input
  --      pati_id           N 1 ����id
  --      pati_pageid       N 1 ��ҳid
  --����: Json_Out,��ʽ����
  --  output
  --    code                  N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message               C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    pati_type             N 1 ���ò�������
  ---------------------------------------------------------------------------
  j_In       Pljson;
  j_Json     Pljson;
  n_����id   ���˱䶯��¼.����id%Type;
  n_��ҳid   ���˱䶯��¼.��ҳid%Type;
  v_�������� Varchar2(100);
Begin
  --�������
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');
  n_��ҳid := j_Json.Get_Number('pati_pageid');

  If Nvl(n_����id, 0) = 0 Then
    Json_Out := Zljsonout('δ���벡��id�����飡');
    Return;
  End If;

  v_�������� := Zl_Patiwarnscheme(n_����id, n_��ҳid);
  Json_Out   := '{"output":{"code":1,"message":"�ɹ�","pati_type":"' || Zljsonstr(v_��������) || '"}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getpatpagewarnscheme;
/
Create Or Replace Procedure Zl_Cissvr_Getstufferrdata
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ��ٴ�ҽ������������������ͬ��
  --��Σ�Json_In:��ʽ
  --  input
  --      pati_ids                        C 1 ����ids����ƴ��
  --����: Json_Out,��ʽ����
  --   output:
  --     code: 1,
  --     message: �ɹ�,
  --     pati_bill_list[]
  --         pati_id                      N 1 ����id
  --         pati_pageid                  N 0 ��ҳid��סԺ���˴��룬���ﴫ0
  --         rgst_id                      N 0 �Һ�id�����ﲡ�˴��룬סԺ���˴���
  --         rgst_no                      C 0 �Һŵ���
  --         send_no                      N 1 ���ͺ�
  --         order_list[]ҽ����Ϣ�б�
  --             advice_id                N 1 ҽ��id
  --             effectivetime            N 1 ҽ����Ч
  --             emergency_tag            N 1 ������־
  --             denominated              N 1 �Ƽ�����
  --             fee_source               N 0 ������Դ��1-���2-סԺ
  --             fee_billtype             N 0 ���õ������ͣ�1-�շѴ�����2-���ʵ�����
  --             fee_no                   C 0 ���õ��ݺ�
  --             freq_name                C 0 Ƶ������
  --             single                   N 0 ����
  ------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  n_��1��   Number;
  c_Jtmp    Clob;
  n_Rgst_Id Number;

  j_Json       Pljson;
  j_Tmp        Pljson;
  v_Jtmp       Varchar2(32767);
  v_Stuff_List Varchar2(32767);
  c_Stuff_List Clob;
  v_Vals       Clob;
  l_Vals       t_Strlist;
  n_������Դ   Number;

  --����ҽ������
  Cursor c_Out
  (
    �Һŵ�_In ����ҽ����¼.�Һŵ�%Type,
    ���ͺ�_In ����ҽ������.���ͺ�%Type
  ) Is
    Select b.ҽ��id As ID, b.No, b.���ͺ�, a.ҽ����Ч, a.������־, a.�Ƽ�����, b.��¼����, c.�������, a.ִ��Ƶ��, a.��������
    From ����ҽ����¼ A, ����ҽ���쳣��¼ B, ����ҽ������ C
    Where a.Id = b.ҽ��id And a.�Һŵ� = �Һŵ�_In And b.�������� = 2 And b.���ͺ� = ���ͺ�_In And b.ҽ��id = c.ҽ��id And b.���ͺ� = c.���ͺ�;

  Type t_Order Is Table Of c_Out%RowType;
  r_Odr t_Order;

  Cursor c_In
  (
    ����id_In ����ҽ����¼.����id%Type,
    ��ҳid_In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In ����ҽ������.���ͺ�%Type
  ) Is
    Select b.ҽ��id As ID, b.No, b.���ͺ�, a.ҽ����Ч, a.������־, a.�Ƽ�����, b.��¼����, c.�������, a.ִ��Ƶ��, a.��������
    From ����ҽ����¼ A, ����ҽ���쳣��¼ B, ����ҽ������ C
    Where a.Id = b.ҽ��id And a.����id = ����id_In And a.��ҳid = ��ҳid_In And b.�������� = 2 And b.���ͺ� = ���ͺ�_In And b.ҽ��id = c.ҽ��id And
          b.���ͺ� = c.���ͺ�;

Begin
  --�������
  If Json_In Is Null Then
    Select f_List2str(Cast(Collect(a.����id || '') As t_Strlist), ',')
    Into v_Vals
    From (Select a.����id From ����ҽ���쳣��¼ A Where a.�������� = 2 Group By a.����id) A;
  Else
    j_Tmp  := Pljson(Json_In);
    j_Json := j_Tmp.Get_Pljson('input');
    v_Vals := j_Json.Get_Clob('pati_ids');
    If v_Vals Is Null Then
      Select f_List2str(Cast(Collect(a.����id || '') As t_Strlist), ',')
      Into v_Vals
      From (Select a.����id From ����ҽ���쳣��¼ A Where a.�������� = 2 Group By a.����id) A;
    End If;
  End If;

  l_Vals := t_Strlist();
  While v_Vals Is Not Null Loop
    If Length(v_Vals) <= 4000 Then
      l_Vals.Extend;
      l_Vals(l_Vals.Count) := v_Vals;
      v_Vals := Null;
    Else
      l_Vals.Extend;
      l_Vals(l_Vals.Count) := Substr(v_Vals, 1, Instr(v_Vals, ',', 3980) - 1);
      v_Vals := Substr(v_Vals, Instr(v_Vals, ',', 3980) + 1);
    End If;
  End Loop;

  n_��1�� := 0;
  For Lp In 1 .. l_Vals.Count Loop
    For Cp In (Select /*+cardinality(j,10) */
                a.����id, a.��ҳid, a.�Һŵ�, b.���ͺ�
               From ����ҽ����¼ A, ����ҽ���쳣��¼ B, Table(f_Num2list(l_Vals(Lp))) J
               Where a.Id = b.ҽ��id And a.����id = j.Column_Value And b.�������� = 2
               Group By a.����id, a.��ҳid, a.�Һŵ�, b.���ͺ�) Loop
    
      n_Rgst_Id := 0;
      If Cp.�Һŵ� Is Not Null Then
        Select a.Id Into n_Rgst_Id From ���˹Һż�¼ A Where a.No = Cp.�Һŵ� And a.��¼���� = 1 And a.��¼״̬ = 1;
      End If;
    
      If Cp.�Һŵ� Is Null Then
        Open c_In(Cp.����id, Cp.��ҳid, Cp.���ͺ�);
        Fetch c_In Bulk Collect
          Into r_Odr;
        Close c_In;
      Else
        Open c_Out(Cp.�Һŵ�, Cp.���ͺ�);
        Fetch c_Out Bulk Collect
          Into r_Odr;
        Close c_Out;
      End If;
    
      n_��1�� := n_��1�� + 1;
      If r_Odr.Count > 0 Then
        v_Stuff_List := Null;
        c_Stuff_List := Null;
        For Ol In 1 .. r_Odr.Count Loop
          If r_Odr(Ol).��¼���� = 1 Or r_Odr(Ol).��¼���� = 2 And r_Odr(Ol).������� = 1 Then
            n_������Դ := 1;
          Else
            n_������Դ := 2;
          End If;
        
          v_Stuff_List := v_Stuff_List || ',{"advice_id":' || r_Odr(Ol).Id;
          v_Stuff_List := v_Stuff_List || ',"effectivetime":' || r_Odr(Ol).ҽ����Ч;
          v_Stuff_List := v_Stuff_List || ',"emergency_tag":' || Nvl(r_Odr(Ol).������־ || '', 'null');
          v_Stuff_List := v_Stuff_List || ',"denominated":' || Nvl(r_Odr(Ol).�Ƽ����� || '', 'null');
          v_Stuff_List := v_Stuff_List || ',"fee_source":' || n_������Դ;
          v_Stuff_List := v_Stuff_List || ',"fee_billtype":' || r_Odr(Ol).��¼����;
          v_Stuff_List := v_Stuff_List || ',"fee_no":"' || r_Odr(Ol).No || '"';
          v_Stuff_List := v_Stuff_List || ',"freq_name":"' || Zljsonstr(r_Odr(Ol).ִ��Ƶ��) || '"';
          v_Stuff_List := v_Stuff_List || ',"single":' || Zljsonstr(r_Odr(Ol).��������, 1);
          v_Stuff_List := v_Stuff_List || '}';
        
          If Length(v_Stuff_List) > 30000 Then
            If c_Stuff_List Is Null Then
              c_Stuff_List := Substr(v_Stuff_List, 2);
            Else
              c_Stuff_List := c_Stuff_List || v_Stuff_List;
            End If;
            v_Stuff_List := Null;
          End If;
        End Loop;
      
        v_Jtmp := Null;
        v_Jtmp := v_Jtmp || ',{"pati_id":' || Cp.����id;
        v_Jtmp := v_Jtmp || ',"pati_pageid":' || Nvl(Cp.��ҳid || '', 'null');
        v_Jtmp := v_Jtmp || ',"rgst_id":' || Nvl(n_Rgst_Id || '', 'null');
        v_Jtmp := v_Jtmp || ',"rgst_no":"' || Cp.�Һŵ� || '"';
        v_Jtmp := v_Jtmp || ',"send_no":' || Cp.���ͺ�;
      
        If n_��1�� = 1 Then
          c_Jtmp := Substr(v_Jtmp, 2);
        Else
          c_Jtmp := c_Jtmp || v_Jtmp;
        End If;
      
        --ҽ���б���
        If c_Stuff_List Is Null Then
          c_Jtmp := c_Jtmp || ',"order_list":[' || Substr(v_Stuff_List, 2) || ']';
        Else
          c_Jtmp := c_Jtmp || ',"order_list":[' || c_Stuff_List || v_Stuff_List || ']';
        End If;
      
        c_Jtmp := c_Jtmp || '}';
      End If;
    End Loop;
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","pati_bill_list":[' || c_Jtmp || ']}}';
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Cissvr_Getstufferrdata;
/
Create Or Replace Procedure Zl_Cissvr_Getsurgandaneinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:��ȡ���˵�������������Ϣ
  --��Σ�Json_In:��ʽ
  -- input
  --   pati_id              N 1 ����id
  --   pati_pageid          N 1 ��ҳid
  --����: Json_Out,��ʽ����
  --  output
  --    code                N  1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message             C  1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    surg_list[]                [����]
  --      dz_code           C 1 ��������
  --      dz_name           C 1 ��������
  --      sicstype          C 1 �����пڷ���
  --      icshlv            C 1 �����п����ϵȼ�
  --      oper_date         C 1 ��������:yyyy-mm-dd hh24:mi:ss
  --      aneitem_type      C 1 ��������
  --      surgeon_name      C 1 ����ҽ������
  --      first_assistant   C 1 ��һ����
  --      second_assistant  C 1 �ڶ�����
  --      surg_anst         C 1 ����ҽ��
  --      recoder           C   ��¼��
  --      rec_time          C   ��¼ʱ��:yyyy-mm-dd hh24:mi:ss
  ---------------------------------------------------------------------------
  j_In     Pljson;
  j_Json   Pljson;
  n_����id ���������¼.����id%Type;
  n_��ҳid ���������¼.��ҳid%Type;
  v_List   Varchar2(32767);
Begin
  --�������
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');
  n_��ҳid := j_Json.Get_Number('pati_pageid');

  If Nvl(n_����id, 0) = 0 Or Nvl(n_��ҳid, 0) = 0 Then
    Json_Out := Zljsonout('δ���벡��id����ҳid�����飡');
    Return;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","surg_list":[';
  For r_���� In (Select b.����, b.����, a.�п�, a.����, To_Char(a.��������, 'yyyy-mm-dd hh24:mi:ss') As ��������, a.��������, a.����ҽʦ, a.��һ����,
                      a.�ڶ�����, a.����ҽʦ, a.��¼��, To_Char(Nvl(a.��¼����, Sysdate), 'yyyy-mm-dd hh24:mi:ss') As ��¼����
               From ���������¼ A, ��������Ŀ¼ B
               Where a.��������id = b.Id And a.����id = n_����id And a.��ҳid = n_��ҳid) Loop
  
    --      dz_code           C 1 ��������
    --      dz_name           C 1 ��������
    --      sicstype          C 1 �����пڷ���
    --      icshlv            C 1 �����п����ϵȼ�
    --      oper_date         C 1 ��������:yyyy-mm-dd hh24:mi:ss
    --      aneitem_type      C 1 ��������
  
    Zljsonputvalue(v_List, 'dz_code', r_����.����, 0, 1);
    Zljsonputvalue(v_List, 'dz_name', r_����.����);
    Zljsonputvalue(v_List, 'sicstype', r_����.�п�);
    Zljsonputvalue(v_List, 'icshlv', r_����.����);
    Zljsonputvalue(v_List, 'oper_date', r_����.��������);
    Zljsonputvalue(v_List, 'aneitem_type', r_����.��������);
  
    --      surgeon_name      C 1 ����ҽ������
    --      first_assistant   C 1 ��һ����
    --      second_assistant  C 1 �ڶ�����
    --      surg_anst         C 1 ����ҽ��
    --      recoder           C   ��¼��
    --      rec_time          C   ��¼ʱ��:yyyy-mm-dd hh24:mi:ss
    Zljsonputvalue(v_List, 'surgeon_name', r_����.����ҽʦ);
    Zljsonputvalue(v_List, 'first_assistant', r_����.��һ����);
    Zljsonputvalue(v_List, 'second_assistant', r_����.�ڶ�����);
    Zljsonputvalue(v_List, 'surg_anst', r_����.����ҽʦ);
    Zljsonputvalue(v_List, 'recoder', r_����.��¼��);
    Zljsonputvalue(v_List, 'rec_time', r_����.��¼����, 0, 2);
    If Length(v_List) > 20000 Then
      Json_Out := Json_Out || v_List;
      v_List   := ',';
    End If;
  End Loop;
  If v_List <> ',' Then
    Json_Out := Json_Out || v_List || ']}}';
  Else
    Json_Out := Json_Out || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getsurgandaneinfo;
/
Create Or Replace Procedure Zl_Cissvr_Isouttakedrug
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ݲ���ID����ҳID,�жϸò����Ƿ��Ժ��ҩ
  --��Σ�Json_In:��ʽ
  --    input
  --        pati_id                 N   1   ����ID
  --        pati_pageid             N   1   ��ҳID
  --    ���� json
  --    output
  --        code                    N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --        message                 C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --        isexist                 N   1   �Ƿ����: 1-����;0-������
  ---------------------------------------------------------------------------
  j_In     Pljson;
  j_Json   Pljson;
  n_����id ����ҽ����¼.����id%Type;
  n_��ҳid ����ҽ����¼.��ҳid%Type;
  n_Count  Number(18);
Begin

  --�������
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');
  n_��ҳid := j_Json.Get_Number('pati_pageid');

  If n_����id Is Null Then
    Json_Out := Zljsonout('δ���벡����Ϣ');
    Return;
  End If;

  Select Count(1)
  Into n_Count
  From ����ҽ����¼ A, ����ҽ����¼ B
  Where a.���id = b.Id And Nvl(a.ִ������, 0) <> 5 And Nvl(b.ִ������, 0) = 5 And a.������� In ('5', '6', '7') And a.����id = n_����id And
        a.��ҳid = n_��ҳid And Rownum < 2;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","isexist":' || n_Count || '}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Isouttakedrug;
/
Create Or Replace Procedure Zl_Cissvr_Newoutpativisitrec
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����: �Һųɹ��������ٴ�����ǼǼ�¼��  Ŀǰ�ٴ��ľ���ǼǼ�¼���ǲ��˹Һż�¼�����Բ��ô���
  --input
  --  rgst_no             C 1 �Һŵ���
  --  rgst_appt_sign      N 1 ԤԼ��־��0-�Һż�¼,1-ԤԼ��¼
  --  rgst_code           C 1 �ű�
  --  rgst_rec_id         N 1 �����¼id
  --  outptyp_name        C 1 ����
  --  pati_id             N 1 ����ID
  --  outpno              C 1 �����
  --  pati_name           C 1 ����
  --  pati_sex            C 1 �Ա�
  --  pati_age            C 1 ����
  --  fee_category        C 1 �ѱ�
  --  revst_sign          N 1 ����:0-��1-��
  --  emg_sign            N 1 ����
  --  outproom_name       C 1 ����
  --  exe_deptid          N 1 ִ�в���ID
  --  rgst_exetr          C 1 ִ����
  --  happen_time         C 1 ����ʱ��
  --  close_account_type  N 1 ����ģʽ
  --����      json
  --output
  --  code                N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message             C 1 Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ,ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ----------------------------------------------------------------------------
  j_In         Pljson;
  j_Json       Pljson;
  v_�Һŵ���   ���˹Һż�¼.No%Type;
  n_ԤԼ��־   ���˹Һż�¼.��¼����%Type;
  v_�ű�       ���˹Һż�¼.�ű�%Type;
  n_�����¼id ���˹Һż�¼.�����¼id%Type;
  v_����       ���˹Һż�¼.����%Type;
  n_����id     ���˹Һż�¼.����id%Type;
  n_�����     ���˹Һż�¼.�����%Type;
  v_����       ���˹Һż�¼.����%Type;
  v_�Ա�       ���˹Һż�¼.�Ա�%Type;
  v_����       ���˹Һż�¼.����%Type;
  n_����       ���˹Һż�¼.����%Type;
  v_�ѱ�       ���˹Һż�¼.�ѱ�%Type;
  n_����       ���˹Һż�¼.����%Type;
  v_����       ���˹Һż�¼.����%Type;
  n_ִ�в���id ���˹Һż�¼.ִ�в���id%Type;
  v_ִ����     ���˹Һż�¼.ִ����%Type;
  d_����ʱ��   ���˹Һż�¼.����ʱ��%Type;
  n_����ģʽ   ���˹Һż�¼.����ģʽ%Type;

Begin
  j_In         := Pljson(Json_In);
  j_Json       := j_In.Get_Pljson('input');
  v_�Һŵ���   := j_Json.Get_String('rgst_no');
  n_ԤԼ��־   := j_Json.Get_Number('rgst_appt_sign');
  v_�ű�       := j_Json.Get_String('rgst_code');
  n_�����¼id := j_Json.Get_Number('rgst_rec_id');
  v_����       := j_Json.Get_String('outptyp_name');
  n_����id     := j_Json.Get_Number('pati_id');
  n_�����     := To_Number(j_Json.Get_String('outpno'));
  v_����       := j_Json.Get_String('pati_name');
  v_�Ա�       := j_Json.Get_String('pati_sex');
  v_����       := j_Json.Get_String('pati_age');
  n_����       := j_Json.Get_Number('revst_sign');
  v_�ѱ�       := j_Json.Get_String('fee_category');
  n_����       := j_Json.Get_Number('emg_sign');
  v_����       := j_Json.Get_String('outproom_name');
  n_ִ�в���id := j_Json.Get_Number('exe_deptid');
  v_ִ����     := j_Json.Get_String('rgst_exetr');
  d_����ʱ��   := To_Date(j_Json.Get_String('happen_time'), 'YYYY-MM-DD hh24:mi:ss');
  n_����ģʽ   := j_Json.Get_Number('close_account_type');

  Json_Out := Zljsonout('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Newoutpativisitrec;
/
Create Or Replace Procedure Zl_Cissvr_Patiexistmemo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:���ݲ���id����ҳid����Ƿ���ڱ�ע��Ϣ
  --��Σ�Json_In:��ʽ
  -- input
  --   pati_id              N 1 ����id
  --   pati_pageid          N 1 ��ҳid
  --����: Json_Out,��ʽ����
  --  output
  --    code               C  1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message            C  1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    exist              N  1 �Ƿ���ڱ�ע��1-�ǵ�;0-��
  ---------------------------------------------------------------------------
  j_In     Pljson;
  j_Json   Pljson;
  n_����id ���˱�ע��Ϣ.����id%Type;
  n_��ҳid ���˱�ע��Ϣ.��ҳid%Type;
  n_Tmp    Number(1);

Begin
  --�������
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');
  n_��ҳid := j_Json.Get_Number('pati_pageid');

  If Nvl(n_����id, 0) = 0 Then
    Json_Out := Zljsonout('���봫�벡��id��');
    Return;
  End If;

  Select Count(1)
  Into n_Tmp
  From ���˱�ע��Ϣ
  Where ����id = n_����id And Nvl(��ҳid, 0) = Nvl(n_��ҳid, 0) And Rownum <= 1;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","exist":' || n_Tmp || '}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Patiexistmemo;
/
Create Or Replace Procedure Zl_Cissvr_Patiisdead
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:���ݲ���id��鲡���Ƿ��Ѿ�����
  --��Σ�Json_In:��ʽ
  -- input
  --   pati_id              N 1 ����id
  --����: Json_Out,��ʽ����
  --  output
  --    code               C  1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message            C  1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    isdeath            N  1 �Ƿ��Ѿ�������1-�ǵ�;0-��
  ---------------------------------------------------------------------------
  j_In     Pljson;
  j_Json   Pljson;
  n_����id ���˱䶯��¼.����id%Type;
  n_Tmp    Number(1);

Begin
  --�������
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');

  If Nvl(n_����id, 0) = 0 Then
    Json_Out := Zljsonout('���봫�벡��id��');
    Return;
  End If;

  Select Count(1) Into n_Tmp From ������ҳ Where ����id = n_����id And ��Ժ��ʽ Like '%����%' And Rownum <= 1;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","isdeath":' || n_Tmp || '}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Patiisdead;
/
Create Or Replace Procedure Zl_Cissvr_Patiisinhospital
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  --------------------------------------------------------------------------------------------------------------------
  --����:�ж�ָ�������Ƿ��Ǵ�����Ժ��ҽ״̬
  --��Σ�Json_In,��ʽ����
  --  input
  --    pati_id            N 1 ����ID
  --    pati_pageid        N 1 ��ҳid
  --����: Json_Out,��ʽ����
  --  output
  --    code               N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message            C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    inhouspital        N 1 0-������Ժ��ҽ״̬��1-�Ǵ�����Ժ��ҽ״̬
  --------------------------------------------------------------------------------------------------------------------
  j_In     Pljson;
  j_Json   Pljson;
  n_����id ������ҳ.����id%Type;
  n_��ҳid ������ҳ.��ҳid%Type;
  n_״̬   Number;

Begin
  --�������
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');
  n_��ҳid := j_Json.Get_Number('pati_pageid');

  n_״̬ := Zl_Pati_Is_Inhospital(n_����id, n_��ҳid);

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","inhouspital":' || Nvl(n_״̬, 0) || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Patiisinhospital;
/
Create Or Replace Procedure Zl_Cissvr_Patiisout
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���ѯ�����Ƿ��Ѿ���Ժ
  --input      ��ѯ�����Ƿ��Ѿ���Ժ
  --  pati_id               N  1  ����id
  --  pati_pageid           N  1  ��ҳid
  --  query_type            N  1  ��ѯ���ͣ�0-�������˲�ѯ��1-�������������ѯ
  --  pati_pageids          C  1  ��ʽ������ID:��ҳID,����ID:��ҳID,...
  --output
  --  code                  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message               C  1  Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  pati_outsign          N  1  ��Ժ��ǣ�0-δ��Ժ��1-��Ժ��query_type=0ʱ����
  --  item_list[]           ��Ժ����б�query_type=1ʱ����
  --    pati_id             N  1  ����id
  --    pati_outsign        N  1  ��Ժ��ǣ�0-δ��Ժ��1-��Ժ
  ---------------------------------------------------------------------------
  j_In    PLJson;
  j_Input PLJson;

  n_��ѯ���� Number(1);
  n_����id   Number(18);
  n_��ҳid   Number(5);
  n_��Ժ��� Number(1); --0-δ��Ժ��1-��Ժ

  c_��ҳid Clob;
  l_��ҳid t_StrList := t_StrList();

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;
Begin
  j_In       := PLJson(Json_In);
  j_Input    := j_In.Get_Pljson('input');
  n_��ѯ���� := j_Input.Get_Number('query_type');

  --0-�������˲�ѯ
  If Nvl(n_��ѯ����, 0) = 0 Then
    n_����id := j_Input.Get_Number('pati_id');
    n_��ҳid := j_Input.Get_Number('pati_pageid');
  
    Select Count(1)
    Into n_��Ժ���
    From ������ҳ
    Where ����id = n_����id And ��ҳid = n_��ҳid And ��Ժ���� Is Not Null And Rownum < 2;
  
    Json_Out := '{"output":{"code":1,"message": "�ɹ�","pati_outsign":' || n_��Ժ��� || '}}';
    Return;
  End If;

  --1 - �������������ѯ 
  If Nvl(n_��ѯ����, 0) = 1 Then
    c_��ҳid := j_Input.Get_Clob('pati_pageids');
  
    While c_��ҳid Is Not Null Loop
      If Length(c_��ҳid) <= 4000 Then
        l_��ҳid.Extend;
        l_��ҳid(l_��ҳid.Count) := c_��ҳid;
        c_��ҳid := Null;
      Else
        l_��ҳid.Extend;
        l_��ҳid(l_��ҳid.Count) := Substr(c_��ҳid, 1, Instr(c_��ҳid, ',', 3950) - 1);
        c_��ҳid := Substr(c_��ҳid, Instr(c_��ҳid, ',', 3950) + 1);
      End If;
    End Loop;
  
    For I In 1 .. l_��ҳid.Count Loop
      For r_���� In (Select /*+Cardinality(j,10)*/
                    j.C1 As ����id, Decode(a.��Ժ����, Null, 0, 1) As ��Ժ��־
                   From ������ҳ A, Table(f_Num2List2(l_��ҳid(I), ',', ':')) J
                   Where a.����id(+) = j.C1 And a.��ҳid(+) = j.C2) Loop
      
        v_Jtmp := v_Jtmp || ',{"pati_id":' || r_����.����id;
        v_Jtmp := v_Jtmp || ',"pati_outsign":' || r_����.��Ժ��־;
        v_Jtmp := v_Jtmp || '}';
      
        If Length(v_Jtmp) > 30000 Then
          If c_Jtmp Is Null Then
            c_Jtmp := Substr(v_Jtmp, 2);
          Else
            c_Jtmp := c_Jtmp || v_Jtmp;
          End If;
          v_Jtmp := Null;
        End If;
      End Loop;
    End Loop;
  
    If c_Jtmp Is Null Then
      Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
    Else
      c_Jtmp   := c_Jtmp || v_Jtmp;
      Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || c_Jtmp || ']}}';
    End If;
    Return;
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || zlJsonStr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Patiisout;
/
Create Or Replace Procedure Zl_Cissvr_Patiobsvtoinhos
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���סԺ���۲���תΪסԺ����
  --���:JSON��ʽ
  --input
  --   pati_id              N 1 ����id
  --   pati_pageid          N 1 ��ҳid
  --   inpatient_num        C 1 סԺ��
  --���Σ�JSON��ʽ
  --output
  --    code                 N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message              C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  n_����id ������ҳ.����id%Type;
  n_��ҳid ������ҳ.��ҳid%Type;
  n_סԺ�� ������ҳ.סԺ��%Type;
  j_In     Pljson;
  j_Json   Pljson;
Begin
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');
  n_��ҳid := j_Json.Get_Number('pati_pageid');
  n_סԺ�� := To_Number(j_Json.Get_String('inpatient_num'));
  Zl_���˱䶯��¼_תסԺ_s(n_����id, n_��ҳid, n_סԺ��);
  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Patiobsvtoinhos;
/
Create Or Replace Procedure Zl_Cissvr_Patipageiscatalogue
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:���ָ����סԺ�Ƿ��Ѿ�����
  --��Σ�Json_In:��ʽ
  --    input
  --      pati_id           N 1 ����id
  --      pati_pageid       N 1 ��ҳid
  --����: Json_Out,��ʽ����
  --  output
  --    code                  N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message               C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    pati_Name             C 1 ����
  --    iscatalogueed         N 1 �Ƿ��Ѿ���Ŀ:1-�Ѿ���Ŀ;0-δ��Ŀ
  ---------------------------------------------------------------------------
  j_In     Pljson;
  j_Json   Pljson;
  n_����id ������ҳ.����id%Type;
  n_��ҳid ������ҳ.��ҳid%Type;
  v_����   ������ҳ.����%Type;
  n_Tmp    Number(1);
  --��װʧ��ʱ���ص�����
  Function Get_Err_Message(Message_In Varchar2) Return Clob Is
    j_Out Clob;
  Begin
    j_Out := '{"output":{"code":0,"message":"' || Zltools.Zljsonstr(Message_In) || '"}}';
  
    Return j_Out;
  End Get_Err_Message;

Begin
  --�������
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');
  n_��ҳid := j_Json.Get_Number('pati_pageid');

  If Nvl(n_����id, 0) = 0 Then
    Json_Out := Get_Err_Message('δ���벡��id�����飡');
    Return;
  End If;

  If Nvl(n_��ҳid, 0) = 0 Then
    Json_Out := Get_Err_Message('δ������ҳid�����飡');
    Return;
  End If;

  Select Count(1), Max(����)
  Into n_Tmp, v_����
  From ������ҳ
  Where ����id = n_����id And ��ҳid = n_��ҳid And ��Ŀ���� Is Not Null;

  Json_Out := '{"output":{"code":1,"message": "�ɹ�","pati_Name":"' || Zljsonstr(v_����) || '","iscatalogueed":' || n_Tmp || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Patipageiscatalogue;
/
Create Or Replace Procedure Zl_Cissvr_Patiregistblacklist
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  --------------------------------------------------------------------------------------------------
  --����:ʵ����֤ǰ�ļ�� 
  --��� JSOM��ʽ
  --input
  --  pati_id                N 1 ����id
  --  calcdate               C 1 ��������
  --  operat_name            C 1 ������Ա
  --  order_last_date        C 1 ���ԤԼʱ��
  --  black_info             C 1 ����id,���ӱ�־;����id,���ӱ�־....
  --���� JSON��ʽ
  --output
  --  code            N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message         C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  black_list[]
  --      operat_type     C 1 ��Ϊ���
  --      pati_id         N 1 ����id
  --      order_date      C 1 ԤԼʱ��
  --      in_reason       C 1 ����ԭ��
  --      in_explain      C 1 ����˵��
  --      sign            C 1 ���ӱ�־
  --      create_name     C 1 �Ǽ���
  --------------------------------------------------------------------------------------------------
  j_In           Pljson;
  j_Json         Pljson;
  n_����id       Number;
  d_��������     Date;
  v_������Ա     Varchar2(100);
  d_���ԤԼʱ�� Date;
  n_Count        Number(18);
  n_ԤԼ����Ч�� Number(18);
  n_ԤԼ�˺�Ч�� Number(18);
  n_ԤԼ����Ч�� Number(18);
  v_List         Varchar2(32676);
  v_Para         Varchar2(32767);
  c_������Ϣ     Clob;
  l_������Ϣ     t_Strlist := t_Strlist();
Begin
  j_In           := Pljson(Json_In);
  j_Json         := j_In.Get_Pljson('input');
  n_����id       := j_Json.Get_Number('pati_id');
  v_������Ա     := j_Json.Get_String('create_name');
  d_��������     := To_Date(j_Json.Get_String('calc_date'), 'yyyy-mm-dd hh24:mi:ss');
  d_���ԤԼʱ�� := To_Date(j_Json.Get_String('order_last_date'), 'yyyy-mm-dd hh24:mi:ss');
  c_������Ϣ     := j_Json.Get_Clob('black_info');

  While c_������Ϣ Is Not Null Loop
    If Length(c_������Ϣ) <= 4000 Then
      l_������Ϣ.Extend;
      l_������Ϣ(l_������Ϣ.Count) := c_������Ϣ;
      c_������Ϣ := Null;
    Else
      l_������Ϣ.Extend;
      l_������Ϣ(l_������Ϣ.Count) := Substr(c_������Ϣ, 1, Instr(c_������Ϣ, ';', 3980) - 1);
      c_������Ϣ := Substr(c_������Ϣ, Instr(c_������Ϣ, ';', 3980) + 1);
    End If;
  End Loop;

  If Nvl(n_����id, 0) = 0 Then
    d_�������� := Trunc(Sysdate) - 1 / 24 / 60 / 60; -- ȱʡ����ͷһ�������
  End If;
  v_Para  := Nvl(zl_GetSysParameter('ԤԼ��Ч��������', '1111'), '0|0|0');
  n_Count := Instr(v_Para, '|');
  If n_Count = 0 Then
    n_ԤԼ����Ч�� := To_Number(Nvl(v_Para, '0'));
    v_Para         := Null;
  Else
    n_ԤԼ����Ч�� := To_Number(Substr(v_Para, 1, n_Count - 1));
    v_Para         := Substr(v_Para, n_Count + 1);
  End If;

  n_ԤԼ����Ч�� := 0;
  If v_Para Is Not Null Then
    n_Count := Instr(v_Para, '|');
    If n_Count = 0 Then
      n_ԤԼ����Ч�� := To_Number(Nvl(v_Para, '0'));
      v_Para         := Null;
    Else
      n_ԤԼ����Ч�� := To_Number(Substr(v_Para, 1, n_Count - 1));
      v_Para         := Substr(v_Para, n_Count + 1);
    End If;
  End If;

  n_ԤԼ�˺�Ч�� := 0;
  If v_Para Is Not Null Then
    n_Count := Instr(v_Para, '|');
    If n_Count = 0 Then
      n_ԤԼ�˺�Ч�� := To_Number(Nvl(v_Para, '0'));
      v_Para         := Null;
    Else
      n_ԤԼ�˺�Ч�� := To_Number(Substr(v_Para, 1, n_Count - 1));
    End If;
  End If;

  n_ԤԼ����Ч�� := -1 * Nvl(n_ԤԼ����Ч��, 0);
  n_ԤԼ����Ч�� := Nvl(n_ԤԼ����Ч��, 0);
  n_ԤԼ�˺�Ч�� := Nvl(n_ԤԼ�˺�Ч��, 0);

  If n_ԤԼ����Ч�� = 0 And n_ԤԼ����Ч�� = 0 And n_ԤԼ�˺�Ч�� = 0 Then
    Return;
  End If;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","black_list":[';
  For I In 1 .. l_������Ϣ.Count Loop
    For c_ԤԼ In (Select Distinct a.No, a.����id, a.��¼����, a.��¼״̬, Nvl(a.ԤԼʱ��, a.����ʱ��) As ԤԼʱ��, c.���� As ��������, ִ����, ����ʱ��
                 From ���˹Һż�¼ A, ���ű� C
                 Where a.ִ�в���id = c.Id(+) And ((a.����id = n_����id And Nvl(n_����id, 0) <> 0) Or Nvl(n_����id, 0) = 0) And
                       a.ԤԼ = 1 And
                       ((a.��¼���� = 2 And Nvl(a.��¼״̬, 0) = 1 And
                       ((a.ԤԼʱ�� + n_ԤԼ����Ч�� * (1 / 24 / 60)) <= Sysdate And n_ԤԼ����Ч�� <> 0)) Or
                       (a.��¼���� = 1 And Nvl(a.��¼״̬, 0) = 1 And
                       ((Nvl(a.ִ��ʱ��, Sysdate) - Nvl(a.ԤԼʱ��, Sysdate)) * 24 * 60 >= n_ԤԼ����Ч��) And n_ԤԼ����Ч�� <> 0) Or
                       (a.��¼״̬ = 2 And ((a.�Ǽ�ʱ�� - Nvl(a.����ʱ��, Sysdate)) * 24 * 60 >= n_ԤԼ�˺�Ч��) And n_ԤԼ�˺�Ч�� <> 0)) And
                       ((a.����ʱ�� + 0 >= d_���ԤԼʱ�� And d_���ԤԼʱ�� Is Not Null) Or d_���ԤԼʱ�� Is Null) And Not Exists
                  (Select 1
                        From (Select Distinct To_Number(C1) As ����id, C2 As ���ӱ�־
                               From Table(f_Str2list2(l_������Ϣ(I)))) B
                        Where a.No = b.���ӱ�־ And a.����id = b.����id) And
                       ((a.ԤԼʱ�� >= Trunc(d_��������) And a.ԤԼʱ�� <= d_�������� And Nvl(n_����id, 0) = 0) Or Nvl(n_����id, 0) <> 0)) Loop
      If v_������Ա Is Null Then
        v_������Ա := Zl_Username;
      End If;
      v_Para := '��' || To_Char(c_ԤԼ.ԤԼʱ��, 'yyyy-mm-dd hh24:mi:ss');
      v_Para := v_Para || 'ԤԼ��"' || c_ԤԼ.�������� || '"����';
    
      If c_ԤԼ.ִ���� Is Not Null Then
        v_Para := v_Para || '��ҽ��Ϊ"' || c_ԤԼ.ִ���� || '"';
      End If;
      v_Para := v_Para || '(ԤԼ��:' || c_ԤԼ.No || Case
                  When c_ԤԼ.��¼״̬ = 2 Then
                   '�����˺�'
                  When c_ԤԼ.��¼���� = 1 Then
                   ' �������ڽ���'
                  Else
                   ''
                End || ')�ĺ�Դ��δ��ʱ���';
      Zljsonputvalue(v_List, 'pati_id', c_ԤԼ.����id, 1, 1);
      Zljsonputvalue(v_List, 'operat_type', 'ԤԼ�Һ�');
      Zljsonputvalue(v_List, 'order_date', To_Char(c_ԤԼ.ԤԼʱ��, 'yyyy-mm-dd hh24:mi:ss'));
      Zljsonputvalue(v_List, 'in_reason', 'ԤԼ����');
      Zljsonputvalue(v_List, 'in_explain', Zljsonstr(v_Para, 0), 0);
      Zljsonputvalue(v_List, 'sign', c_ԤԼ.No);
      Zljsonputvalue(v_List, 'create_name', v_������Ա, 0, 2);
      If Length(v_List) > 20000 Then
        Json_Out := Json_Out || v_List;
        v_List   := ',';
      End If;
    End Loop;
    If v_List <> ',' Then
      Json_Out := Json_Out || v_List || ']}}';
    Else
      Json_Out := Json_Out || ']}}';
    End If;
  End Loop;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Patiregistblacklist;
/

Create Or Replace Procedure Zl_Cissvr_Setautocalcsign
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  -------------------------------------------------------------------------------------------------
  --���ܣ������Զ����ʱ��
  --��Σ�json��ʽ
  --Input
  --   pati_id               N  1 ����id
  --   pati_pageids          C  1 ��ҳid,�����Ӣ�Ķ��ŷָ�
  --   auto_account_sign     N  1 �Զ����ʱ�־��0-�����Զ�����,1-��ֹ�Զ�����
  --���Σ�json��ʽ
  --Json_Out
  --   code                  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --   message               C  1  Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
  -------------------------------------------------------------------------------------------------
  j_In   Pljson;
  j_Json Pljson;

  n_����id       ������ҳ.����id%Type;
  v_��ҳids      Varchar2(32767);
  n_��ֹ�Զ����� ������ҳ.�Ƿ��ֹ�Զ�����%Type;
Begin
  --�������
  j_In           := Pljson(Json_In);
  j_Json         := j_In.Get_Pljson('input');
  n_����id       := j_Json.Get_Number('pati_id');
  v_��ҳids      := j_Json.Get_String('pati_pageids');
  n_��ֹ�Զ����� := j_Json.Get_Number('auto_account_sign');

  Update ������ҳ
  Set �Ƿ��ֹ�Զ����� = Nvl(n_��ֹ�Զ�����, 0)
  Where ����id = n_����id And ��ҳid In (Select Column_Value From Table(f_Num2list(v_��ҳids))) And �������� = 0;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Setautocalcsign;
/
Create Or Replace Procedure Zl_Cissvr_Setbedempty
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --���ܣ�ɾ�����˴�λ״����¼
  --��Σ�Json_In:��ʽ
  --input
  --  pati_id           N    1 ����id
  --����      json
  --output
  --    code                N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  -------------------------------------------
  j_In     Pljson;
  j_Json   Pljson;
  n_����id Number(18);
Begin
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');
  Update ��λ״����¼ Set ״̬ = '�մ�', ����id = Null, ����id = Decode(����, 1, Null, ����id) Where ����id = n_����id;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Setbedempty;
/
Create Or Replace Procedure Zl_Cissvr_Updadvicechargetag
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����Ĳ���ҽ�����ݵļƷ�״̬��ɾ����������
  --��Σ�Json_In:��ʽ
  --input
  --  item_list[]
  --    advice_id               N   1 ҽ��ID
  --    send_no                 N   1 ���ͺ�
  --    bill_no                 C   1 ���ݺ�
  --    bill_prop               N   1 ��¼����
  --    del_annex               N   0 �Ƿ�ɾ����Ӧ��ҽ����������:1-ɾ��;0-��ɾ��
  --    charge_status           N   1 ���µļƷ�״̬:-1-����Ʒ�(ͨ����ִ�к�Ժ��ִ�еĶ�����Ʒ�);0-δ�Ʒ�;1-�ѼƷ�(�����)��2-�����շ�/�˷�(����/����)��3-ȫ���շ�(�������շ��д���)��4-ȫ���˷�(����)
  --  fee_detail_list[]                ������ϸ���������˷�ʱ����
  --    advice_id               N   1 ҽ��ID
  --    bill_no                 C   1 ���ݺ�
  --    bill_prop               N   1 ��¼����
  --    fee_item_id             N   1 �շ�ϸĿID
  --����: Json_Out,��ʽ����
  --output
  --   code                     N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --   message                  C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Json    Pljson;
  j_Json_In Pljson;

  j_Jsonlist_In Pljson_List;
  j_Detail      Pljson_List;
  v_No          ����ҽ������.No%Type;
  n_���ͺ�      ����ҽ������.���ͺ�%Type;
  n_��¼����    ����ҽ������.��¼����%Type;
  n_ҽ��id      ����ҽ������.ҽ��id%Type;
  n_�Ʒ�״̬    ����ҽ������.�Ʒ�״̬%Type;
  n_ɾ������    Number;
  n_�շ�ϸĿid  ҽ��ִ�мƼ�.�շ�ϸĿid%Type;
  n_��ҽ��id    ����ҽ������.ҽ��id%Type;
Begin
  j_Json        := Pljson(Json_In);
  j_Json_In     := j_Json.Get_Pljson('input');
  j_Jsonlist_In := j_Json_In.Get_Pljson_List('item_list');
  j_Detail      := j_Json_In.Get_Pljson_List('fee_detail_list');

  For I In 1 .. j_Jsonlist_In.Count Loop
    j_Json     := Pljson();
    j_Json     := Pljson(j_Jsonlist_In.Get(I));
    n_ҽ��id   := j_Json.Get_Number('advice_id');
    n_���ͺ�   := j_Json.Get_Number('send_no');
    v_No       := j_Json.Get_String('bill_no');
    n_��¼���� := j_Json.Get_Number('bill_prop');
    n_ɾ������ := j_Json.Get_Number('del_annex');
    n_�Ʒ�״̬ := j_Json.Get_Number('charge_status');
  
    If Nvl(n_ɾ������, 0) = 1 Then
      Delete From ����ҽ������ Where ҽ��id = n_ҽ��id And ��¼���� = n_��¼���� And NO = v_No;
    End If;
  
    If Nvl(n_���ͺ�, 0) <> 0 And Nvl(n_ҽ��id, 0) <> 0 And v_No Is Null Then
      --��������ҽ����Ҫͬ��������ҽ���ļƷ�״̬
      Select Max(ID)
      Into n_��ҽ��id
      From ����ҽ����¼
      Where ID In (Select Max(Nvl(���id, 0)) From ����ҽ����¼ Where ID = n_ҽ��id) And Instr(',D,F,', ',' || ������� || ',') > 0;
    
      Update ����ҽ������
      Set �Ʒ�״̬ = n_�Ʒ�״̬
      Where (ҽ��id = n_ҽ��id Or ҽ��id = Nvl(n_��ҽ��id, 0)) And ���ͺ� = n_���ͺ�;
    Else
      Update ����ҽ������ A
      Set a.�Ʒ�״̬ = n_�Ʒ�״̬
      Where ҽ��id = n_ҽ��id And ��¼���� = n_��¼���� And NO = v_No;
    End If;
  End Loop;

  If j_Detail Is Not Null Then
    For I In 1 .. j_Detail.Count Loop
      j_Json       := Pljson();
      j_Json       := Pljson(j_Jsonlist_In.Get(I));
      n_ҽ��id     := j_Json.Get_Number('advice_id');
      v_No         := j_Json.Get_String('bill_no');
      n_��¼����   := j_Json.Get_Number('bill_prop');
      n_�շ�ϸĿid := j_Json.Get_Number('fee_item_id');
    
      Update ҽ��ִ�мƼ�
      Set ִ��״̬ = 2
      Where (ҽ��id, ���ͺ�) In (Select ҽ��id, ���ͺ� From ����ҽ������ Where NO = v_No And ��¼���� = Nvl(n_��¼����, 0)) And
            �շ�ϸĿid = n_�շ�ϸĿid And ִ��״̬ = 0;
    End Loop;
  End If;

  Json_Out := Zljsonout('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Updadvicechargetag;
/
Create Or Replace Procedure Zl_Cissvr_Updadviceexestatus
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ�����ҽ�����͵�ִ��״̬
  --��Σ�Json_In:��ʽ
  --  input
  --    item_list[]
  --      advice_id              N  1 ҽ��ID
  --      bill_no                C  1 ���ݺ�
  --      bill_prop              N  1 ��¼����
  --      exe_status_old         C    ִ��״̬:����ö���
  --      exe_status             N  1 ���µ�ִ��״̬
  --      exetr                  C  1 ���µ�ִ����
  --      exe_time               C  1 ���µ�ִ��ʱ��:yyyy-mm-dd hh24:mi:ss
  --����: Json_Out,��ʽ����
  --  output
  --    code                      N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                   C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_In   Pljson;
  j_Json Pljson;
  j_List Pljson_List;

  n_ҽ��id   ����ҽ������.ҽ��id%Type;
  n_ִ��״̬ ����ҽ������.ִ��״̬%Type;
  n_��¼���� ����ҽ������.��¼����%Type;
  v_ִ��״̬ Varchar2(100);
  v_ִ����   ����ҽ������.�����%Type;
  d_ִ��ʱ�� ����ҽ������.���ʱ��%Type;
  v_No       ����ҽ������.No%Type;
  n_Count    Number(18);

Begin
  --�������
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');
  j_List := j_Json.Get_Pljson_List('item_list');

  If Not j_List Is Null Then
    n_Count := j_List.Count;
    For I In 1 .. n_Count Loop
      j_Json     := Pljson();
      j_Json     := Pljson(j_List.Get(I));
      n_ҽ��id   := j_Json.Get_Number('advice_id');
      v_No       := j_Json.Get_String('bill_no');
      n_ִ��״̬ := j_Json.Get_Number('exe_status');
      n_��¼���� := j_Json.Get_Number('bill_prop');
      v_ִ��״̬ := j_Json.Get_String('exe_status_old');
      v_ִ����   := j_Json.Get_String('exetr');
      d_ִ��ʱ�� := To_Date(j_Json.Get_String('exe_time'), 'yyyy-mm-dd hh24:mi:ss');
    
      If Nvl(n_ҽ��id, 0) = 0 Then
        Json_Out := Zljsonout('δ����ҽ��id���ͺţ����飡');
        Return;
      End If;
    
      If Nvl(n_��¼����, 0) = 0 Then
        Json_Out := Zljsonout('δ�����¼���ʣ����飡');
        Return;
      End If;
    
      If Nvl(v_ִ��״̬, '-') = '-' Then
        Json_Out := Zljsonout('δ����ִ��״̬�����飡');
        Return;
      End If;
    
      If Nvl(v_No, '-') = '-' Then
        Json_Out := Zljsonout('δ����NO�����飡');
        Return;
      End If;
    
      Update ����ҽ������
      Set ִ��״̬ = Nvl(n_ִ��״̬, 0), ����� = v_ִ����, ���ʱ�� = d_ִ��ʱ��
      Where ҽ��id = n_ҽ��id And NO = v_No And ��¼���� = n_��¼���� And Instr(',' || v_ִ��״̬ || ',', ',' || ִ��״̬ || ',') > 0;
    End Loop;
  End If;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Updadviceexestatus;
/
Create Or Replace Procedure Zl_Cissvr_Updatecritical
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is

  --------------------------------------------------------------------------------------------------------------------
  --���ܣ�Σ��ֵ��ز������������޸ģ�����,ɾ��
  --��Σ�Json_In,��ʽ����
  --  input
  --       business_type           N  1   ҵ������:1-����,2-�޸�,3-����,4-ɾ��

  --       cvalue_id               N  1    Σ��ֵid
  --       cvitem_source           C  1    ������Դ
  --       pati_id                 N  1    ����id
  --       pati_pageid             N  1    ��ҳid
  --       rgst_no                 C  1    �Һŵ�
  --       baby_num                N  1    Ӥ��
  --       pat_name                C  1    ��������
  --       pat_sex                 C  1    �����Ա�
  --       pat_age                 C  1    ��������
  --       advice_id               N  1    ҽ��id
  --       lspcm_id                N  1    �걾id
  --       cvalue_rec_desc         C  1    Σ��ֵ˵��
  --       cvalue_rec_create_time  C  1    ����ʱ��
  --       rpt_deptid              N  1    �������id
  --       rec_rptor               C  1    ������

  --       proc_note               C  1    �������
  --       cvalue_cnfmtime         C  1    ȷ��ʱ��
  --       cvalue_cnfmer           C  1    ȷ����
  --       cvalue_deptid           N  1    ȷ�Ͽ���id
  --       cvitem_result           N  1    �Ƿ�Σ��ֵ

  --����: Json_Out,��ʽ����
  --  output
  --    code                            N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                         C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --------------------------------------------------------------------------------------------------------------------
  n_Type Number(5); --1-����,2-�޸�,3-����

  n_Id         ����Σ��ֵ��¼.Id%Type;
  v_������Դ   ����Σ��ֵ��¼.������Դ%Type;
  n_����id     ����Σ��ֵ��¼.����id%Type;
  n_��ҳid     ����Σ��ֵ��¼.��ҳid%Type;
  v_�Һŵ�     ����Σ��ֵ��¼.�Һŵ�%Type;
  n_Ӥ��       ����Σ��ֵ��¼.Ӥ��%Type;
  v_����       ����Σ��ֵ��¼.����%Type;
  v_�Ա�       ����Σ��ֵ��¼.�Ա�%Type;
  v_����       ����Σ��ֵ��¼.����%Type;
  n_ҽ��id     ����Σ��ֵ��¼.ҽ��id%Type;
  n_�걾id     ����Σ��ֵ��¼.�걾id%Type;
  v_Σ��ֵ���� ����Σ��ֵ��¼.Σ��ֵ����%Type;
  v_����ʱ��   ����Σ��ֵ��¼.����ʱ��%Type;
  n_�������id ����Σ��ֵ��¼.�������id%Type;
  v_������     ����Σ��ֵ��¼.������%Type;

  v_�������   ����Σ��ֵ��¼.�������%Type;
  v_ȷ��ʱ��   ����Σ��ֵ��¼.ȷ��ʱ��%Type;
  v_ȷ����     ����Σ��ֵ��¼.ȷ����%Type;
  n_ȷ�Ͽ���id ����Σ��ֵ��¼.ȷ�Ͽ���id%Type;
  n_�Ƿ�Σ��ֵ ����Σ��ֵ��¼.�Ƿ�Σ��ֵ%Type;

  j_In   Pljson;
  j_Json Pljson;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin

  --�������
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');

  n_Type := Nvl(j_Json.Get_Number('business_type'), 0);
  If n_Type = 1 Or n_Type = 2 Then
    n_Id         := j_Json.Get_Number('cvalue_id');
    v_������Դ   := j_Json.Get_String('cvitem_source');
    n_����id     := j_Json.Get_Number('pati_id');
    n_��ҳid     := j_Json.Get_Number('pati_pageid');
    v_�Һŵ�     := j_Json.Get_String('rgst_no');
    n_Ӥ��       := j_Json.Get_Number('baby_num');
    v_����       := j_Json.Get_String('pat_name');
    v_�Ա�       := j_Json.Get_String('pat_sex');
    v_����       := j_Json.Get_String('pat_age');
    n_ҽ��id     := j_Json.Get_Number('advice_id');
    n_�걾id     := j_Json.Get_Number('lspcm_id');
    v_Σ��ֵ���� := j_Json.Get_String('cvalue_rec_desc');
    v_����ʱ��   := To_Date(j_Json.Get_String('cvalue_rec_create_time'), 'yyyy-MM-dd HH24:MI:SS');
    n_�������id := j_Json.Get_Number('rpt_deptid');
    v_������     := j_Json.Get_String('rec_rptor');
  
    If n_Type = 1 Then
      Zl_����Σ��ֵ��¼_Insert(n_Id, v_������Դ, n_����id, n_��ҳid, v_�Һŵ�, n_Ӥ��, v_����, v_�Ա�, v_����, n_ҽ��id, n_�걾id, v_Σ��ֵ����, v_����ʱ��,
                        n_�������id, v_������);
    Else
      Zl_����Σ��ֵ��¼_Update(n_Id, v_������Դ, n_����id, n_��ҳid, v_�Һŵ�, n_Ӥ��, v_����, v_�Ա�, v_����, n_ҽ��id, n_�걾id, v_Σ��ֵ����, v_����ʱ��,
                        n_�������id, v_������);
    End If;
  Elsif n_Type = 3 Then
    n_Id         := j_Json.Get_Number('cvalue_id');
    v_�������   := j_Json.Get_String('proc_note');
    v_ȷ��ʱ��   := To_Date(j_Json.Get_String('cvalue_cnfmtime'), 'yyyy-MM-dd HH24:MI:SS');
    v_ȷ����     := j_Json.Get_String('cvalue_cnfmer');
    n_ȷ�Ͽ���id := j_Json.Get_Number('cvalue_deptid');
    n_�Ƿ�Σ��ֵ := j_Json.Get_Number('cvitem_result');
    Zl_����Σ��ֵ��¼_����(n_Id, v_�������, v_ȷ��ʱ��, v_ȷ����, n_ȷ�Ͽ���id, n_�Ƿ�Σ��ֵ);
  Elsif n_Type = 4 Then
    n_Id := j_Json.Get_Number('cvalue_id');
    Zl_����Σ��ֵ��¼_Delete(n_Id);
  End If;

  Json_Out := Zljsonout('�ɹ�', 1);
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Updatecritical;
/
Create Or Replace Procedure Zl_Cissvr_Updatediaginfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����²��������Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --     pati_id            N 1 ����id
  --     pati_pageid        N 1 ��ҳid
  --     diag_types         C 1 �������:0-��������,1-��ҽ�������;2-��ҽ��Ժ���;3-��ҽ��Ժ���;5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���,8-��ǰ���;9-�������;10-����֢;11-��ҽ�������;12-��ҽ��Ժ���;
  --                                     13-��ҽ��Ժ���;21-��ԭѧ���;22-Ӱ��ѧ���.����Ϊ���������ͣ��ö��ŷ���,��:2,12
  --     diag_num           N 1 ��ϴ���
  --     rec_source         N 1 ��¼��Դ:1-������2-��Ժ�Ǽǣ�3-��ҳ����;4-����
  --     diag_note          C 1 �������
  --����: Json_Out,��ʽ����
  --  output
  --    code          N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_In       Pljson;
  j_Json     Pljson;
  n_����id   ������ϼ�¼.����id%Type;
  n_��ҳid   ������ϼ�¼.��ҳid%Type;
  n_��¼��Դ ������ϼ�¼.��¼��Դ%Type;
  n_��ϴ��� ������ϼ�¼.��ϴ���%Type;
  v_������� ������ϼ�¼.�������%Type;
  v_������� ������ϼ�¼.�������%Type;

Begin
  --�������
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  n_����id   := j_Json.Get_Number('pati_id');
  n_��ҳid   := j_Json.Get_Number('pati_pageid');
  n_��ϴ��� := j_Json.Get_Number('diag_num');
  n_��¼��Դ := j_Json.Get_Number('rec_source');
  v_������� := j_Json.Get_String('diag_types');
  v_������� := j_Json.Get_String('diag_note');

  Update ������ϼ�¼
  Set ������� = v_�������
  Where ��¼��Դ = n_��¼��Դ And (Instr(',' || v_������� || ',', ',' || ������� || ',') > 0 Or Nvl(v_�������, 0) = 0) And ��ϴ��� = n_��ϴ��� And
        ����id = n_����id And ��ҳid = n_��ҳid;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Updatediaginfo;
/
Create Or Replace Procedure Zl_Cissvr_Updateinfectinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  --------------------------------------------------------------------------------------------------------------------
  --���ܣ��������Լ�¼��ز������������޸�,ɾ��
  --��Σ�Json_In,��ʽ����
  --  input
  --       business_type           N   1   ҵ������:1-����,2-�޸�,3-ɾ��

  --       rec_id                   N   1  ����id
  --       pati_id                 N   0  ����id
  --       pati_pageid             N   0  ��ҳid
  --       reg_no                  C   0  �Һŵ�
  --       advice_id               N   0  ҽ��id
  --       spcm_send_time          C   0  �ͼ�ʱ��
  --       spcm_send_deptid        N   0  �ͼ����ID
  --       spcm_send_dr            C   0  �ͼ�ҽ��
  --       spcm_name               C   0  �걾����
  --       send_content            C   0  �������
  --       infctdz_name            C   0  ��Ⱦ������
  --       eqpmtn_exetime          C   0  ���ʱ��
  --       create_time             C   0  �Ǽ�ʱ��
  --       create_dr               C   0  �Ǽ�ҽ��
  --       create_dept_id          N   0  �Ǽǿ���ID
  --       spcm_rec_status         N   0  ��¼״̬

  --       operate_type            N   0  �޸Ĳ������ͣ�1-���ô���˵�� ,2-�������浥�����Խ��������,3-ȡ�����浥�����Խ���������Ĺ���,4-�޸����Խ��������
  --       spcm_procor             C   0  ������
  --       spcm_proctime           C   0  ����ʱ��
  --       spcm_procdesc           C   0  �������˵��
  --       emr_doc_id              N   0  �ļ�ID

  --����: Json_Out,��ʽ����
  --  output
  --    code                            N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                         C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --------------------------------------------------------------------------------------------------------------------
  n_Type Number(5); --1-����,2-�޸�,3-ɾ��

  n_����id     �������Լ�¼.Id%Type;
  n_����id     �������Լ�¼.����id%Type;
  n_��ҳid     �������Լ�¼.��ҳid%Type;
  v_�Һŵ�     �������Լ�¼.�Һŵ�%Type;
  n_ҽ��id     �������Լ�¼.ҽ��id%Type;
  d_�ͼ�ʱ��   �������Լ�¼.�ͼ�ʱ��%Type;
  n_�ͼ����id �������Լ�¼.�ͼ����id%Type;
  v_�ͼ�ҽ��   �������Լ�¼.�ͼ�ҽ��%Type;
  v_�걾����   �������Լ�¼.�걾����%Type;
  v_�������   �������Լ�¼.�������%Type;
  v_��Ⱦ������ �������Լ�¼.��Ⱦ������%Type;
  d_���ʱ��   �������Լ�¼.���ʱ��%Type;
  d_�Ǽ�ʱ��   �������Լ�¼.�Ǽ�ʱ��%Type;
  v_�Ǽ�ҽ��   �������Լ�¼.�Ǽ���%Type;
  n_�Ǽǿ���id �������Լ�¼.�Ǽǿ���id%Type;
  n_��¼״̬   �������Լ�¼.��¼״̬%Type;

  n_�޸Ĳ������� Number(5);

  v_������       �������Լ�¼.������%Type;
  d_����ʱ��     �������Լ�¼.����ʱ��%Type;
  v_�������˵�� �������Լ�¼.�������˵��%Type;
  n_�ļ�id       �������Լ�¼.�ļ�id%Type;

  j_In    Pljson;
  j_Json  Pljson;
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin

  --�������
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');

  n_Type := Nvl(j_Json.Get_Number('business_type'), 0);
  If n_Type = 1 Then
    n_����id     := j_Json.Get_Number('rec_id');
    n_����id     := j_Json.Get_Number('pati_id');
    n_��ҳid     := j_Json.Get_Number('pati_pageid');
    v_�Һŵ�     := j_Json.Get_String('reg_no');
    n_ҽ��id     := j_Json.Get_Number('advice_id');
    d_�ͼ�ʱ��   := To_Date(j_Json.Get_String('spcm_send_time'), 'yyyy-MM-dd HH24:MI:SS');
    n_�ͼ����id := j_Json.Get_Number('spcm_send_deptid');
    v_�ͼ�ҽ��   := j_Json.Get_String('spcm_send_dr');
    v_�걾����   := j_Json.Get_String('spcm_name');
    v_�������   := j_Json.Get_String('send_content');
    v_��Ⱦ������ := j_Json.Get_String('infctdz_name');
    d_���ʱ��   := To_Date(j_Json.Get_String('eqpmtn_exetime'), 'yyyy-MM-dd HH24:MI:SS');
    d_�Ǽ�ʱ��   := To_Date(j_Json.Get_String('create_time'), 'yyyy-MM-dd HH24:MI:SS');
    v_�Ǽ�ҽ��   := j_Json.Get_String('create_dr');
    n_�Ǽǿ���id := j_Json.Get_Number('create_dept_id');
    n_��¼״̬   := j_Json.Get_Number('spcm_rec_status');
  
    Zl_�������Լ���¼_Insert(n_����id, n_����id, n_��ҳid, v_�Һŵ�, n_ҽ��id, d_�ͼ�ʱ��, n_�ͼ����id, v_�ͼ�ҽ��, v_�걾����, v_�������, v_��Ⱦ������, d_���ʱ��,
                       d_�Ǽ�ʱ��, v_�Ǽ�ҽ��, n_�Ǽǿ���id, n_��¼״̬);
  Elsif n_Type = 2 Then
    n_�޸Ĳ������� := Nvl(j_Json.Get_Number('operate_type'), 0);
    n_����id       := j_Json.Get_Number('rec_id');
    n_�ļ�id       := j_Json.Get_Number('emr_doc_id');
    n_��¼״̬     := j_Json.Get_Number('spcm_rec_status');
    v_������       := j_Json.Get_String('spcm_procor');
    d_����ʱ��     := To_Date(j_Json.Get_String('spcm_proctime'), 'yyyy-MM-dd HH24:MI:SS');
    v_�������˵�� := j_Json.Get_String('spcm_procdesc');
  
    d_�ͼ�ʱ��   := To_Date(j_Json.Get_String('spcm_send_time'), 'yyyy-MM-dd HH24:MI:SS');
    n_�ͼ����id := j_Json.Get_Number('spcm_send_deptid');
    v_�ͼ�ҽ��   := j_Json.Get_String('spcm_send_dr');
    v_�걾����   := j_Json.Get_String('spcm_name');
    v_�������   := j_Json.Get_String('send_content');
    v_��Ⱦ������ := j_Json.Get_String('infctdz_name');
    d_���ʱ��   := To_Date(j_Json.Get_String('eqpmtn_exetime'), 'yyyy-MM-dd HH24:MI:SS');
    d_�Ǽ�ʱ��   := To_Date(j_Json.Get_String('create_time'), 'yyyy-MM-dd HH24:MI:SS');
    v_�Ǽ�ҽ��   := j_Json.Get_String('create_dr');
    n_�Ǽǿ���id := j_Json.Get_Number('create_dept_id');
  
    Zl_�������Լ���¼_Update(n_�޸Ĳ�������, n_����id, n_�ļ�id, n_��¼״̬, v_������, d_����ʱ��, v_�������˵��, d_�ͼ�ʱ��, n_�ͼ����id, v_�ͼ�ҽ��, v_�걾����,
                       v_�������, v_��Ⱦ������, d_���ʱ��, d_�Ǽ�ʱ��, v_�Ǽ�ҽ��, n_�Ǽǿ���id);
  Elsif n_Type = 3 Then
    n_����id := j_Json.Get_Number('rec_id');
    Zl_�������Լ�¼_Delete(n_����id);
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Updateinfectinfo;
/
Create Or Replace Procedure Zl_Cissvr_Updateinpatiextinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:�޸Ĳ�����ҳ�ӱ������Ϣ
  --��Σ�Json_In:��ʽ
  --    input
  --    pati_id             N  1  ����id
  --    pati_pageid         N  1  ��ҳId
  --    item_list[]               �б�
  --      info_name         C  1  ��Ϣ��
  --      info_value        C  1  �޸ĵ���Ϣֵ
  --����: Json_Out,��ʽ����
  --  output
  --    code                N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message             C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------

  j_Json     Pljson;
  o_Json     Pljson;
  j_Jsonlist Pljson_List := Pljson_List();

  n_����id ������ҳ�ӱ�.����id%Type;
  n_��ҳid ������ҳ�ӱ�.��ҳid%Type;
  v_��Ϣ�� ������ҳ�ӱ�.��Ϣ��%Type;
  v_��Ϣֵ ������ҳ�ӱ�.��Ϣֵ%Type;
  Err_Item Exception;
Begin
  --�������
  o_Json     := Pljson(Json_In);
  j_Json     := o_Json.Get_Pljson('input');
  n_����id   := j_Json.Get_Number('pati_id');
  n_��ҳid   := j_Json.Get_Number('pati_pageid');
  j_Jsonlist := j_Json.Get_Pljson_List('item_list');
  If j_Jsonlist Is Null Then
    Json_Out := Zljsonout('��������Ҫ����Ĵ�����Ŀ��Ϣ', 0);
    Return;
  End If;
  For I In 1 .. j_Jsonlist.Count Loop
    o_Json   := Pljson();
    o_Json   := Pljson(j_Jsonlist.Get(I));
    v_��Ϣ�� := o_Json.Get_String('info_name');
    v_��Ϣֵ := o_Json.Get_String('info_value');
    Zl_������ҳ�ӱ�_��ҳ����(n_����id, n_��ҳid, v_��Ϣ��, v_��Ϣֵ);
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Updateinpatiextinfo;
/
Create Or Replace Procedure Zl_Cissvr_Updateinpatipageinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------------------------------------------------------------
  --���ܣ����²�����ҳ�Ͳ�����ҳ�ӱ���Ϣ
  --��Σ�json��ʽ
  --Input
  --   pati_id               N  1 ����id
  --   pati_pageid           N  1 ��ҳid
  --   inpatient_num         C  1 סԺ��
  --   fee_type              C  1 �ѱ�
  --   inpati_fee_type       C  1 סԺ�ѱ�
  --   mdlpay_mode_name      C  1 ҽ�Ƹ��ʽ
  --   pati_area             C  1 ����
  --   remarkes              C  1 ��ע
  --   pati_marital_cstatus  C  1 ����״��
  --   pati_education        C  1 ѧ��
  --   ocpt_name             C  1 ְҵ
  --   emp_phno              C  1 ��λ�绰
  --   emp_postcode          C  1 ��λ�ʱ�
  --   pati_home_addr         C  1 ��ͥ��ַ
  --   pati_home_phno         C  1 ��ͥ�绰
  --   pati_home_postcode     C  1 ��ͥ��ַ�ʱ�
  --   pati_hous_addr         C  1 ���ڵ�ַ
  --   pati_hous_postcode     C  1 ���ڵ�ַ�ʱ�
  --   contacts_name         C  1 ��ϵ������
  --   contacts_relation     C  1 ��ϵ�˹�ϵ
  --   contacts_addr         C  1 ��ϵ�˵�ַ
  --   contacts_phno         C  1 ��ϵ�˵绰
  --   pati_type             C  1 ��������
  --   insurance_num         C  1 ҽ����
  --   outpatient_num        N  1 �����
  --   item_list[]              1 ������ҳ�ӱ���Ϣ
  --   regist_time           C  1 �Ǽ�ʱ��
  --   opr_type              N  1 ִ�з�ʽ 0-���� 1-�޸�
  --���Σ�json��ʽ
  --Json_Out
  --   code                  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --   message               C  1  Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  -------------------------------------------------------------------------------------------------

  j_In       Pljson;
  j_Json     Pljson;
  j_Jsonlist Pljson_List := Pljson_List();
  o_Json     Pljson;

  n_����id       ������ҳ.����id%Type;
  n_��ҳid       ������ҳ.��ҳid%Type;
  n_סԺ��       ������ҳ.סԺ��%Type;
  v_�ѱ�         ������ҳ.�ѱ�%Type;
  v_סԺ�ѱ�     ������ҳ.�ѱ�%Type;
  v_ҽ�Ƹ��ʽ ������ҳ.ҽ�Ƹ��ʽ%Type;
  v_����         ������ҳ.����%Type;
  v_��ע         ������ҳ.��ע%Type;
  v_����״��     ������ҳ.����״��%Type;
  v_ѧ��         ������ҳ.ѧ��%Type;
  v_ְҵ         ������ҳ.ְҵ%Type;
  v_��λ�绰     ������ҳ.��λ�绰%Type;
  v_��λ�ʱ�     ������ҳ.��λ�ʱ�%Type;
  v_��ͥ��ַ     ������ҳ.��ͥ��ַ%Type;
  v_��ͥ�绰     ������ҳ.��ͥ�绰%Type;
  v_��ͥ��ַ�ʱ� ������ҳ.��ͥ��ַ�ʱ�%Type;
  v_���ڵ�ַ     ������ҳ.���ڵ�ַ%Type;
  v_���ڵ�ַ�ʱ� ������ҳ.���ڵ�ַ�ʱ�%Type;
  v_��ϵ������   ������ҳ.��ϵ������%Type;
  v_��ϵ�˹�ϵ   ������ҳ.��ϵ�˹�ϵ%Type;
  v_��ϵ�˵�ַ   ������ҳ.��ϵ�˵�ַ%Type;
  v_��ϵ�˵绰   ������ҳ.��ϵ�˵绰%Type;
  v_��������     ������ҳ.��������%Type;
  v_��Ժ����     ������ҳ.��Ժ����%Type;
  v_ҽ����       ������ҳ�ӱ�.��Ϣֵ%Type;
  n_�����       ���ﲡ����¼.������%Type;
  v_��Ϣ��       ������ҳ�ӱ�.��Ϣ��%Type;
  v_��Ϣֵ       ������ҳ�ӱ�.��Ϣֵ%Type;
  d_�Ǽ�ʱ��     Date;
  n_Type         Number;
Begin
  --�������
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');

  n_����id       := j_Json.Get_Number('pati_id');
  n_��ҳid       := j_Json.Get_Number('pati_pageid');
  n_סԺ��       := To_Number(j_Json.Get_String('inpatient_num'));
  v_�ѱ�         := j_Json.Get_String('fee_type');
  v_סԺ�ѱ�     := j_Json.Get_String('inpati_fee_type');
  v_ҽ�Ƹ��ʽ := j_Json.Get_String('mdlpay_mode_name');
  v_����         := j_Json.Get_String('pati_area');
  v_��ע         := j_Json.Get_String('remarkes');
  v_����״��     := j_Json.Get_String('pati_marital_cstatus');
  v_ѧ��         := j_Json.Get_String('pati_education');
  v_ְҵ         := j_Json.Get_String('ocpt_name');
  v_��λ�绰     := j_Json.Get_String('emp_phno');
  v_��λ�ʱ�     := j_Json.Get_String('emp_postcode');
  v_��ͥ��ַ     := j_Json.Get_String('pat_home_addr');
  v_��ͥ�绰     := j_Json.Get_String('pat_home_phno');
  v_��ͥ��ַ�ʱ� := j_Json.Get_String('pat_home_postcode');
  v_���ڵ�ַ     := j_Json.Get_String('pat_hous_addr');
  v_���ڵ�ַ�ʱ� := j_Json.Get_String('pat_hous_postcode');
  v_��ϵ������   := j_Json.Get_String('contacts_name');
  v_��ϵ�˹�ϵ   := j_Json.Get_String('contacts_relation');
  v_��ϵ�˵�ַ   := j_Json.Get_String('contacts_addr');
  v_��ϵ�˵绰   := j_Json.Get_String('contacts_phno');
  v_��������     := j_Json.Get_String('pati_type');
  v_ҽ����       := j_Json.Get_String('insurance_num');
  n_�����       := To_Number(j_Json.Get_Number('outpatient_num'));
  j_Jsonlist     := j_Json.Get_Pljson_List('item_list');
  d_�Ǽ�ʱ��     := To_Date(j_Json.Get_String('regist_time'), 'yyyy-mm-dd hh24:mi:ss');
  n_Type         := j_Json.Get_Number('opr_type');
  If Nvl(n_Type, 0) = 1 Then
    If n_����� Is Not Null Then
      Update ���ﲡ����¼ Set ������ = n_����� Where ����id = n_����id;
      If Sql%RowCount = 0 Then
        Insert Into ���ﲡ����¼
          (����id, ������, ��������, �������, �洢״̬, ���λ��)
        Values
          (n_����id, n_�����, Sysdate, 'һ��', '����', Null);
      End If;
    Else
      Delete From ���ﲡ����¼ Where ����id = n_����id;
    End If;
  Else
    If n_����� Is Not Null Then
      Insert Into ���ﲡ����¼
        (����id, ������, ��������, �������, �洢״̬, ���λ��)
      Values
        (n_����id, n_�����, d_�Ǽ�ʱ��, 'һ��', '����', Null);
    End If;
  End If;

  If n_��ҳid Is Not Null And n_��ҳid <> 0 Then
    If Nvl(n_Type, 0) = 1 Then
      Update ������ҳ
      Set סԺ�� = n_סԺ��, �ѱ� = Decode(Nvl(v_סԺ�ѱ�, 0), 1, v_�ѱ�, �ѱ�), ҽ�Ƹ��ʽ = v_ҽ�Ƹ��ʽ, ���� = Decode(v_����, Null, ����, v_����),
          ��ע = v_��ע
      Where ����id = n_����id And ��ҳid = n_��ҳid;
    
      --����Ժ���˸��²�����ҳ�е���Ϣ
      Begin
        Select ��Ժ���� Into v_��Ժ���� From ������ҳ Where ����id = n_����id And ��ҳid = n_��ҳid;
      Exception
        When Others Then
          Null;
      End;
      If v_��Ժ���� Is Null Then
        Update ������ҳ
        Set ����״�� = v_����״��, ѧ�� = v_ѧ��, ְҵ = v_ְҵ, ��λ�绰 = v_��λ�绰, ��λ�ʱ� = v_��λ�ʱ�, ��ͥ��ַ = v_��ͥ��ַ, ��ͥ�绰 = v_��ͥ�绰,
            ��ͥ��ַ�ʱ� = v_��ͥ��ַ�ʱ�, ���ڵ�ַ = Nvl(v_���ڵ�ַ, v_���ڵ�ַ), ���ڵ�ַ�ʱ� = Nvl(v_���ڵ�ַ�ʱ�, ���ڵ�ַ�ʱ�), ��ϵ������ = v_��ϵ������,
            ��ϵ�˹�ϵ = v_��ϵ�˹�ϵ, ��ϵ�˵�ַ = v_��ϵ�˵�ַ, ��ϵ�˵绰 = v_��ϵ�˵绰, �������� = v_��������, ��ע = v_��ע

        Where ����id = n_����id And ��ҳid = n_��ҳid;
      End If;
      If v_ҽ���� Is Not Null Then
        Update ������ҳ�ӱ� Set ��Ϣֵ = v_ҽ���� Where ����id = n_����id And ��ҳid = n_��ҳid And ��Ϣ�� = 'ҽ����';
        If Sql%RowCount = 0 Then
          Insert Into ������ҳ�ӱ� (����id, ��ҳid, ��Ϣ��, ��Ϣֵ) Values (n_����id, n_��ҳid, 'ҽ����', v_ҽ����);
        End If;
      Else
        Delete From ������ҳ�ӱ� Where ����id = n_����id And ��ҳid = n_��ҳid And ��Ϣ�� = 'ҽ����';
      End If;
    End If;
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json   := Pljson();
      o_Json   := Pljson(j_Jsonlist.Get(I));
      v_��Ϣ�� := o_Json.Get_String('info_name');
      v_��Ϣֵ := o_Json.Get_String('info_value');
      Zl_������ҳ�ӱ�_��ҳ����(n_����id, n_��ҳid, v_��Ϣ��, v_��Ϣֵ);
    End Loop;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Updateinpatipageinfo;
/
Create Or Replace Procedure Zl_Cissvr_Updateoutmedrecord
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  -------------------------------------------------------------------------------------------------
  --���ܣ��������ﵵ����Ϣ
  --��Σ�json��ʽ
  --Input
  --   pati_id               N  1 ����id
  --   mr_no                 C  1 ������
  --   outpatient_num        C  1 �����
  --   create_date           C  1 ��������
  --   mr_type               C  1 �������
  --   strg_status           C  1 �洢״̬
  --   strgloc               C  1 ���λ��
  --���Σ�json��ʽ
  --Json_Out
  --   code                  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --   message               C  1  Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  -------------------------------------------------------------------------------------------------

  n_����id   ���ﲡ����¼.����id%Type;
  n_������   ���ﲡ����¼.������%Type;
  n_�����   ���ﲡ����¼.������%Type;
  d_�������� ���ﲡ����¼.��������%Type;
  v_���λ�� ���ﲡ����¼.���λ��%Type;
  v_������� ���ﲡ����¼.�������%Type;
  v_�洢״̬ ���ﲡ����¼.�洢״̬%Type;
  j_Json     Pljson;
  j_In       Pljson;

Begin
  --�������
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');

  n_����id   := j_Json.Get_Number('pati_id');
  n_������   := To_Number(j_Json.Get_String('mr_no'));
  d_�������� := To_Date(j_Json.Get_String('create_date'), 'yyyy-mm-dd hh24:mi:ss');
  n_�����   := To_Number(j_Json.Get_String('outpatient_num'));
  v_���λ�� := j_Json.Get_String('strgloc');
  v_������� := j_Json.Get_String('mr_type');
  v_�洢״̬ := j_Json.Get_String('strg_status');
  If Nvl(n_����id, 0) = 0 Then
    Json_Out := Zljsonout('δ���벡��id�����飡', 0);
    Return;
  End If;
  If n_������ Is Not Null Then
    If Nvl(n_������, 0) = 0 Then
      Json_Out := Zljsonout('δ��������ţ����飡', 0);
      Return;
    End If;
  
    Update ���ﲡ����¼ Set ������ = n_������ Where ����id = n_����id;
    If Sql%RowCount = 0 Then
      Insert Into ���ﲡ����¼
        (����id, ������, ��������, �������, �洢״̬, ���λ��)
      Values
        (n_����id, n_������, d_��������, v_�������, v_�洢״̬, v_���λ��);
    End If;
  Else
    If n_����� Is Not Null Then
      Update ���ﲡ����¼ Set ������ = n_����� Where ����id = n_����id;
      If Sql%RowCount = 0 Then
        Insert Into ���ﲡ����¼
          (����id, ������, ��������, �������, �洢״̬, ���λ��)
        Values
          (n_����id, n_�����, Sysdate, 'һ��', '����', Null);
      End If;
    Else
      Delete From ���ﲡ����¼ Where ����id = n_����id;
    End If;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Updateoutmedrecord;
/
Create Or Replace Procedure Zl_Cissvr_Updateoutvitalsigns
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:�޸Ĳ��˵�����������Ϣ
  --��Σ�Json_In:��ʽ
  --    input
  --      pati_id           N  1 ����ID
  --      rgst_id           N  1 �Һ�ID
  --      pat_vsign         N  1 ��������ʽΪ����ĿID1|��Ŀֵ1|��λ1,��ĿID2|��Ŀֵ2|��λ2
  --      operator_name     C    ����Ա����

  --����: Json_Out,��ʽ����
  --  output
  --    code                N  1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message             C  1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------

  j_In   Pljson;
  j_Json Pljson;

  n_����id     ���˻����¼.����id%Type;
  n_�Һ�id     ���˻����¼.��ҳid%Type;
  v_����Ա���� ���˻����¼.������%Type;
  v_����       Varchar2(4000);

Begin
  --�������
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');

  v_����       := j_Json.Get_String('pat_vsign');
  n_����id     := j_Json.Get_Number('pati_id');
  n_�Һ�id     := j_Json.Get_Number('rgst_id');
  v_����Ա���� := j_Json.Get_String('operator_name');

  If Nvl(n_�Һ�id, 0) = 0 Or Nvl(n_����id, 0) = 0 Then
    Json_Out := Zljsonout('δȷ�����˵ľ�����Ϣ,����!');
    Return;
  End If;

  Zl_������������_Update(n_����id, n_�Һ�id, v_����, v_����Ա����);

  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Updateoutvitalsigns;
/
Create Or Replace Procedure Zl_Cissvr_Updatepatiauditinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------------------------------------------------------------
  --���ܣ����²��������Ϣ
  --��Σ�json��ʽ
  --Input
  --   pati_id               N  1 ����id
  --   pati_pageid           N  1 ��ҳid
  --   audit_sign            N  1 ��˱�ǣ�0���-δ���,1-����˻�ʼ���;2-������
  --   auditor               C  0 �����
  --   audit_desc            C  0 ���˵��
  --   cancel_audit          N  0 �Ƿ�ȡ����ˣ�1-ȡ�����,0-���
  --���Σ�json��ʽ
  --Json_Out
  --   code                  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --   message               C  1  Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
  -------------------------------------------------------------------------------------------------
  j_In   Pljson;
  j_Json Pljson;

  n_����id   ������ҳ.����id%Type;
  n_��ҳid   ������ҳ.��ҳid%Type;
  n_��˱�� ������ҳ.��˱�־%Type;
  v_�����   ������ҳ.�����%Type;
  v_���˵�� ������ҳ.���˵��%Type;
  n_ȡ����� Number(2);
Begin
  --�������
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  n_����id   := j_Json.Get_Number('pati_id');
  n_��ҳid   := j_Json.Get_Number('pati_pageid');
  n_��˱�� := j_Json.Get_Number('audit_sign');
  v_�����   := j_Json.Get_String('auditor');
  v_���˵�� := j_Json.Get_String('audit_desc');
  n_ȡ����� := j_Json.Get_Number('cancel_audit');

  Zl_�������_Execute(n_����id, n_��ҳid, n_��˱��, v_�����, Nvl(n_ȡ�����, 0), v_���˵��);

  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Updatepatiauditinfo;
/
Create Or Replace Procedure Zl_Cissvr_Updatepatibaseinfo
(
  Json_in    Varchar2,
  Json_out Out Varchar2
) Is
  --------------------------------------------------------------------------------------------------
  --����:���²����ٴ���صĲ��˻�����Ϣ
  --��� JSOM��ʽ
  --input
  --  pati_id               N 1 ����id
  --  visit_id              N 1 ����id
  --  occasion              N 1 ����
  --  update_info[]  ������Ϣ
  --      pati_name             C 1 ����
  --      pati_age              C 1 ����
  --      pati_sex              C 1 �Ա�
  --���� JSON��ʽ
  --output
  --  code                      N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message                   C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  adjust_explain            C 1 �޸�ԭ��
  --------------------------------------------------------------------------------------------------
  o_Json   Pljson;
  j_Json   Pljson;
  j_In     Pljson;
  n_����id ������ҳ.����id%Type;
  v_����   ������ҳ.����%Type;
  v_�Ա�   ������ҳ.�Ա�%Type;
  v_����   ������ҳ.����%Type; --����ǰ������
  n_����id Number;
  n_����   Number(1);
  v_˵��   Varchar2(32676);
  ˵��_Out Clob;

  Procedure p_����
  (
    ����id_In ���˹Һż�¼.����id%Type,
    ����id_In Number,
    ����_In   ���˹Һż�¼.����%Type,
    �Ա�_In   ���˹Һż�¼.�Ա�%Type,
    ����_In   ���˹Һż�¼.����%Type,
    ����_In   Number, --1-����;2-סԺ
    ˵��_Out  Out Varchar2
  ) As
    Err_Custom Exception;
    v_Error Varchar2(2000);
  Begin
    --��Ҫ�ط�ʽ�洢���������Ա�����
    For r_Rec In (Select /*+ RULE */
                  Distinct d.���� ����, a.Id, a.��������, a.��������, a.���ʱ��, a.������,
                           Nvl((Select Decode(a.�༭��ʽ, 0, Decode(Substr(b.��������, 1, 1), '2', 1, 0),
                                               Decode(Substr(b.��������, Instr(b.��������, '|'), 1), '2', 1, 0))
                                From ���Ӳ������� B
                                Where a.Id = b.�ļ�id And
                                      ((b.�������� = 8 And a.�༭��ʽ = 0) Or (b.�������� In (6, 7, 8) And a.�༭��ʽ = 1)) And
                                      Rownum < 2), 0) As ����ǩ��
                  From ���Ӳ�����¼ A, ���ű� D
                  Where a.����id = ����id_In And a.��ҳid = ����id_In And a.����id = d.Id And Exists
                   (Select 1 --����
                         From ���Ӳ������� C
                         Where a.Id = c.�ļ�id And ((c.�������� = 4 And c.Ҫ������ = '����') Or (c.�������� = 4 And c.Ҫ������ = '�Ա�') Or
                               (c.�������� = 4 And c.Ҫ������ = '����') Or (c.�������� = 2 And c.Ҫ������ = '����') Or
                               (c.�������� = 2 And c.Ҫ������ = '�Ա�') Or (c.�������� = 2 And c.Ҫ������ = '����'))) And Rownum < 2) Loop
    
      --��ȡ���а����������Ա�����Ҫ�صĲ���
      If r_Rec.����ǩ�� <> 1 Then
        --���ز������ƴ�
        ˵��_Out := r_Rec.���� || ':��д�Ĳ����а������˻�����Ϣ����Ҫ�ֹ�������';
        If r_Rec.�������� = 5 Then
          --���¼����걨��¼
          Update �����걨��¼ Set ���� = ����_In, �Ա� = �Ա�_In, ���� = ����_In Where �ļ�id = r_Rec.Id;
        End If;
      End If;
    End Loop;
  
    --������¼�������
    If Nvl(����_In, 0) = 1 Then
      For r_Rec In (Select /*+ RULE */
                    Distinct d.���� ����, a.Id, a.��������, a.��������, a.���ʱ��, a.������,
                             Nvl((Select Decode(a.�༭��ʽ, 0, Decode(Substr(b.��������, 1, 1), '2', 1, 0),
                                                 Decode(Substr(b.��������, Instr(b.��������, '|'), 1), '2', 1, 0))
                                  From ���Ӳ������� B
                                  Where a.Id = b.�ļ�id And
                                        ((b.�������� = 8 And �༭��ʽ = 0) Or (b.�������� In (6, 7, 8) And �༭��ʽ = 1)) And Rownum < 2),
                                  0) As ����ǩ��
                    From ���Ӳ�����¼ A, ���ű� D, ���˹Һż�¼ E
                    Where a.����id = ����id_In And a.��ҳid = ����id_In And a.����id = d.Id And a.����id = e.����id And Exists
                     (Select 1 --����
                           From ���Ӳ������� C
                           Where a.Id = c.�ļ�id And
                                 ((c.�������� = 2 And Instr(c.�����ı�, e.����) > 0) Or (c.�������� = 1 And Instr(c.�����ı�, e.����) > 0))) And
                          Rownum < 2) Loop
      
        --��ȡ���а��������Ĳ���
        If r_Rec.����ǩ�� <> 1 Then
          --���ز������ƴ�
          ˵��_Out := r_Rec.���� || ':��д�Ĳ����а������˻�����Ϣ����Ҫ�ֹ�������';
          If r_Rec.�������� = 5 Then
            --���¼����걨��¼
            Update �����걨��¼ Set ���� = ����_In, �Ա� = �Ա�_In, ���� = ����_In Where �ļ�id = r_Rec.Id;
          End If;
        End If;
      End Loop;
    Else
      For r_Rec In (Select /*+ RULE */
                    Distinct d.���� ����, a.Id, a.��������, a.��������, a.���ʱ��, a.������,
                             Nvl((Select Decode(a.�༭��ʽ, 0, Decode(Substr(b.��������, 1, 1), '2', 1, 0),
                                                 Decode(Substr(b.��������, Instr(b.��������, '|'), 1), '2', 1, 0))
                                  From ���Ӳ������� B
                                  Where a.Id = b.�ļ�id And
                                        ((b.�������� = 8 And �༭��ʽ = 0) Or (b.�������� In (6, 7, 8) And �༭��ʽ = 1)) And Rownum < 2),
                                  0) As ����ǩ��
                    From ���Ӳ�����¼ A, ���ű� D, ������ҳ E
                    Where a.����id = ����id_In And a.��ҳid = ����id_In And a.����id = d.Id And a.����id = e.����id And Exists
                     (Select 1 --����
                           From ���Ӳ������� C
                           Where a.Id = c.�ļ�id And
                                 ((c.�������� = 2 And Instr(c.�����ı�, e.����) > 0) Or (c.�������� = 1 And Instr(c.�����ı�, e.����) > 0))) And
                          Rownum < 2) Loop
      
        --��ȡ���а��������Ĳ���
        If r_Rec.����ǩ�� <> 1 Then
          --���ز������ƴ�
          ˵��_Out := r_Rec.���� || ':��д�Ĳ����а������˻�����Ϣ����Ҫ�ֹ�������';
          If r_Rec.�������� = 5 Then
            --���¼����걨��¼
            Update �����걨��¼ Set ���� = ����_In, �Ա� = �Ա�_In, ���� = ����_In Where �ļ�id = r_Rec.Id;
          End If;
        End If;
      End Loop;
    End If;
  
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_����;

  Procedure p_ҽ��
  (
    ����id_In ���˹Һż�¼.����id%Type,
    ����id_In Number,
    ����_In   ���˹Һż�¼.����%Type,
    �Ա�_In   ���˹Һż�¼.�Ա�%Type,
    ����_In   ���˹Һż�¼.����%Type,
    ����_In   Number, --1-����;2-סԺ
    ˵��_Out  Out Varchar2
  ) As
    ------------------------------------------------------------------------------------------
    --����:����ҽ�����ҵ�����ݵĲ��˻�����Ϣ
    --���:����id_In:����ID
    --     ����id_In:���ﲡ��Ϊ�Һ�ID;סԺ����Ϊ��ҳID;Ϊ0˵����������ǹҺž���Ĳ���(����id_InΪ��ʱ,���������ĸò��˵�����ҵ������)
    --     ����_In:��Ҫ���ĵĲ�������
    --     �Ա�_In:��Ҫ���ĵĲ����Ա�
    --     ����_In:��Ҫ���ĵĲ�������
    --     ����_In:1-����;2-סԺ
    --����:˵��_Out:������Ϣ�������˵����Ϣ��������ʾ����Ա������ز���
    ------------------------------------------------------------------------------------------
    Err_Custom Exception;
    v_Error Varchar2(2000);
    n_Count Number(3);
    v_No    ���˹Һż�¼.No%Type;
  Begin
    --������Ա��������
    If Nvl(����id_In, 0) = 0 Then
      Return;
    End If;
    --����ȡ�Һŵ�
    If Nvl(����_In, 0) = 1 Then
      --���²��˱��ξ����ҽ���еĲ��˻�����Ϣ
      Select NO Into v_No From ���˹Һż�¼ Where ID = ����id_In;
    
      Update ����ҽ����¼
      Set ���� = Nvl(����_In, ����), �Ա� = Nvl(�Ա�_In, �Ա�), ���� = Nvl(����_In, ����)
      Where ����id = ����id_In And �Һŵ� = v_No;
    
      ---���²���Σ��ֵ��¼
      Update ����Σ��ֵ��¼
      Set ���� = Nvl(����_In, ����), �Ա� = Nvl(�Ա�_In, �Ա�), ���� = Nvl(����_In, ����)
      Where ����id = ����id_In And �Һŵ� = v_No;
      Return;
    End If;
    --סԺ����
    If Nvl(����_In, 0) = 2 Then
      --�Ѿ���ӡ��ҽ���嵥����ʾ���´�ӡ
      Select Nvl(Count(1), 0)
      Into n_Count
      From ����ҽ����ӡ
      Where ����id = ����id_In And ��ҳid = ����id_In And Rownum < 2;
    
      If n_Count <> 0 Then
        If Not ˵��_Out Is Null Then
          ˵��_Out := ˵��_Out || Chr(13);
        End If;
        ˵��_Out := ˵��_Out || 'ҽ���嵥:�Ѿ���ӡ�����´�ӡ.';
      End If;
    
      --�Ѿ���ӡ����ҳ����ʾ���´�ӡ
      Select Nvl(Count(1), 0)
      Into n_Count
      From ���Ӳ�����ӡ
      Where ����id = ����id_In And ��ҳid = ����id_In And �ļ�id Is Null And ���� = 9 And Rownum < 2;
      If n_Count <> 0 Then
        If Not ˵��_Out Is Null Then
          ˵��_Out := ˵��_Out || Chr(13);
        End If;
        ˵��_Out := ˵��_Out || '������ҳ:�Ѿ���ӡ�����´�ӡ.';
      End If;
    
      Update ����ҽ����¼
      Set ���� = Nvl(����_In, ����), �Ա� = Nvl(�Ա�_In, �Ա�), ���� = Nvl(����_In, ����)
      Where ����id = ����id_In And ��ҳid = ����id_In;
    
      ---���²���Σ��ֵ��¼
      Update ����Σ��ֵ��¼
      Set ���� = Nvl(����_In, ����), �Ա� = Nvl(�Ա�_In, �Ա�), ���� = Nvl(����_In, ����)
      Where ����id = ����id_In And ��ҳid = ����id_In;
      Return;
    End If;
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_ҽ��;

Begin
  j_In     := Pljson(Json_in);
  j_Json   := j_In.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');
  n_����id := j_Json.Get_Number('visit_id');
  n_����   := j_Json.Get_Number('occasion');
  o_Json   := j_Json.Get_Pljson('update_info');
  If o_Json Is Null Then
    Json_out := Zljsonout('δ������Ҫ���µ���Ϣ�����飡', 0);
    Return;
  End If;
  v_���� := o_Json.Get_String('pati_name');
  v_�Ա� := o_Json.Get_String('pati_sex');
  v_���� := o_Json.Get_String('pati_age');

  --ҽ������
  p_ҽ��(n_����id, n_����id, v_����, v_�Ա�, v_����, n_����, v_˵��);
  If v_˵�� Is Not Null Then
    ˵��_Out := ˵��_Out || Chr(13) || 'ҽ������:' || Chr(13) || v_˵��;
  End If;
  --��������
  v_˵�� := '';
  p_����(n_����id, n_����id, v_����, v_�Ա�, v_����, n_����, v_˵��);
  If v_˵�� Is Not Null Then
    ˵��_Out := ˵��_Out || Chr(13) || '��������:' || Chr(13) || v_˵��;
  End If;

  --����������ĸ���(������ҳ�����˹Һż�¼��������Ϣ)
  If n_���� = 1 And Nvl(n_����id, 0) <> 0 Then
    --���ﲡ��
    Update ���˹Һż�¼ A Set ���� = v_����, �Ա� = v_�Ա�, ���� = v_���� Where ����id = n_����id And a.Id = n_����id;
  End If;
  If n_���� = 2 And Nvl(n_����id, 0) <> 0 Then
    --סԺ����
    Update ������ҳ Set ���� = v_����, �Ա� = v_�Ա�, ���� = v_���� Where ����id = n_����id And ��ҳid = n_����id;
  End If;
  Json_out := '{"output":{"code":1,"message":"�ɹ�","adjust_explain":"' || Zljsonstr(˵��_Out) || '"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Updatepatibaseinfo;
/
Create Or Replace Procedure Zl_Cissvr_Updatesyncstate
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ�ͬ�����¼����
  --��Σ�Json_In:��ʽ
  --  input
  --      order_list[]
  --          order_id          N 1 ҽ��id
  --          send_no           N 1 ���ͺ�
  --          sign_type         N 1 ���ñ��¼�����ͣ�
  --                                  ˵����1-���������¼
  --                                        2-��� ����ҩƷͬ�����
  --                                        3-��� ��������ͬ�����
  --����: Json_Out,��ʽ����
  --  output
  --    code          N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Json       Pljson;
  o_Json       Pljson;
  j_Order_List Pljson_List;
  n_ҽ��id     Number;
  n_���ͺ�     Number;
  n_����       Number;
  n_Count      Number;
Begin
  --�������
  o_Json       := Pljson(Json_In);
  j_Json       := o_Json.Get_Pljson('input');
  j_Order_List := j_Json.Get_Pljson_List('order_list');
  n_Count      := j_Order_List.Count;
  If n_Count > 0 Then
    For I In 1 .. n_Count Loop
      o_Json   := Pljson();
      o_Json   := Pljson(j_Order_List.Get(I));
      n_ҽ��id := o_Json.Get_Number('order_id');
      n_���ͺ� := o_Json.Get_Number('send_no');
      n_����   := o_Json.Get_Number('sign_type');
      --��������: 1-����ҩƷ��2-�������ģ�3-������Һ��4-�շ�ȷ��ҩƷ��5-�շ�ȷ�����ģ�6-����ҩƷ��7-��������
      If n_���� = 1 Then
        Delete ����ҽ���쳣��¼ Where ҽ��id = n_ҽ��id And ���ͺ� = n_���ͺ� And �������� In (3, 5);
      Elsif n_���� = 2 Then
        Delete ����ҽ���쳣��¼ Where ҽ��id = n_ҽ��id And ���ͺ� = n_���ͺ� And �������� In (1, 4);
      Elsif n_���� = 3 Then
        Delete ����ҽ���쳣��¼ Where ҽ��id = n_ҽ��id And ���ͺ� = n_���ͺ� And �������� In (2, 5);
      End If;
    End Loop;
  End If;
  Json_Out := '{"output":{"code":1,"message": "�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Updatesyncstate;
/
CREATE OR REPLACE Procedure Zl_Cissvr_Updoutpativisitrec
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------------------------------------------------------------
  --���ܣ����²��˾�����Ϣ
  --��Σ�json��ʽ
  --Input
  --   rgst_no	            C	1	�Һŵ���
  --   pati_id              N   ����ID
  --   outpatient_num       C   �����
  --   pati_name            C   ����
  --   pati_sex             C   �Ա�
  --   pati_age             C   ����
  --   fee_category         C   �ѱ�
  --   exetr	              C	 	ִ����
  --   outproom_name	      C	 	����
  --   exe_time	            C	 	ִ��ʱ��
  --   exe_status           N   ִ��״̬
  --   rgst_desc	          C	 	ժҪ
  --   pnurs_oprtr	        N	 	�Ƿ�ʿִ��
  --   outp_recv_time_end   C	 	���ʱ��

  --���Σ�json��ʽ
  --Json_Out
  --   code                  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --   message               C  1  Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  -------------------------------------------------------------------------------------------------

  v_�Һŵ��� ���˹Һż�¼.No%Type;
  v_ִ����   ���˹Һż�¼.ִ����%Type;
  v_����     ���˹Һż�¼.����%Type;
  d_ִ��ʱ�� ���˹Һż�¼.ִ��ʱ��%Type;
  v_ժҪ     ���˹Һż�¼.ժҪ%Type;
  n_��ʿִ�� Number;
  j_In       Pljson;
  j_Json     Pljson;

Begin
  --�������
  j_In := Pljson(Json_In);

  j_Json     := j_In.Get_Pljson('input');
  v_�Һŵ��� := j_Json.Get_String('rgst_no');
  v_ִ����   := j_Json.Get_String('exetr');
  v_����     := j_Json.Get_String('outproom_name');
  d_ִ��ʱ�� := To_Date(j_Json.Get_String('exe_time'), 'YYYY-MM-DD hh24:mi:ss');
  v_ժҪ     := j_Json.Get_String('rgst_desc');
  n_��ʿִ�� := j_Json.Get_Number('pnurs_oprtr');

  Json_Out := '{"output":{"code":1,"message": "�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Updoutpativisitrec;
/
Create Or Replace Procedure Zl_Cissvr_Updpatbaseinfocheck
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ------------------------------------------------------------------------------------------
  --����:����ҽ�����ҵ�����ݵĲ��˻�����Ϣ�ļ��
  --���:JSON��ʽ
  --input
  --   pati_id   N 1 ����id
  --   visit_id   N 1 ����id �����ﲡ��Ϊ�Һ�ID;סԺ����Ϊ��ҳID;Ϊ0˵����������ǹҺž���Ĳ���(����id_InΪ��ʱ,���������ĸò��˵�����ҵ������)
  --   occasion   N 1 ����,����_In:1-����;2-סԺ
  --����:JSON��ʽ
  --output
  --  code            N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message        C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ------------------------------------------------------------------------------------------
  j_In     Pljson;
  j_Json   Pljson;
  v_Error  Varchar2(2000);
  n_Count  Number(3);
  v_No     ���˹Һż�¼.No%Type;
  v_Tmp    Varchar2(100);
  n_����id Number;
  n_����id Number;
  n_����   Number;
Begin
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');
  n_����id := j_Json.Get_Number('visit_id');
  n_����   := j_Json.Get_Number('occasion');
  --����ȡ�Һŵ�
  If Nvl(n_����id, 0) = 0 Then
    Return;
  End If;
  If Nvl(n_����, 0) = 1 Then
    Select NO Into v_No From ���˹Һż�¼ Where ID = n_����id;
    If v_No Is Null Then
      v_Error  := 'δ�ҵ��ò��˵ĹҺż�¼,���ܸ��²��˻�����Ϣ.';
      Json_Out := Zljsonout(v_Error, 1);
      Return;
    End If;
    --����ҽ��ǩ��,�������޸Ĳ��˻�����Ϣ
    Select Nvl(Count(1), 0)
    Into n_Count
    From ����ҽ����¼
    Where ����id = n_����id And �Һŵ� = v_No And �¿�ǩ��id Is Not Null And Rownum < 2;
    If n_Count <> 0 Then
      v_Error  := '����ҽ���Ѿ�ǩ��,���ܸ��²��˻�����Ϣ.';
      Json_Out := Zljsonout(v_Error, 1);
      Return;
    End If;
  End If;
  --סԺ����
  If Nvl(n_����, 0) = 2 Then
    --סԺҽ��ǩ��,�������޸Ĳ��˻�����Ϣ
    Select Nvl(Count(1), 0)
    Into n_Count
    From ����ҽ����¼
    Where ����id = n_����id And ��ҳid = n_����id And �¿�ǩ��id Is Not Null And Rownum < 2;
  
    If n_Count <> 0 Then
      v_Error  := '�ò���ҽ���Ѿ�ǩ��,���ܸ��²��˻�����Ϣ.';
      Json_Out := Zljsonout(v_Error, 1);
      Return;
    End If;
    --������������״̬���������޸Ĳ��˻�����Ϣ
    Select Decode(����״̬, 1, '�ȴ������', 3, '���������', 5, '�Ѿ����鵵', 10, '���մ�����', Null)
    Into v_Tmp
    From ������ҳ
    Where ����id = n_����id And ��ҳid = n_����id;
  
    If Not v_Tmp Is Null Then
      v_Error  := '�ò��˵Ĳ���' || v_Tmp || ',���ܸ��²��˻�����Ϣ.';
      Json_Out := Zljsonout(v_Error, 1);
      Return;
    End If;
    --�������ڱ�Ŀ״̬���������޸Ĳ��˻�����Ϣ
    Select Nvl(Count(1), 0)
    Into n_Count
    From ������ҳ
    Where ����id = n_����id And ��ҳid = n_����id And ��Ŀ���� Is Not Null;
    If n_Count <> 0 Then
      v_Error  := '�ò��˵Ĳ����Ѿ���Ŀ,���ܸ��²��˻�����Ϣ.';
      Json_Out := Zljsonout(v_Error, 1);
      Return;
    End If;
  End If;
  --��Ҫ�ط�ʽ�洢���������Ա�����
  For r_Rec In (Select /*+ RULE */
                Distinct d.���� ����, a.Id, a.��������, a.��������, a.���ʱ��, a.������,
                         Nvl((Select Decode(a.�༭��ʽ, 0, Decode(Substr(b.��������, 1, 1), '2', 1, 0),
                                             Decode(Substr(b.��������, Instr(b.��������, '|'), 1), '2', 1, 0))
                              From ���Ӳ������� B
                              Where a.Id = b.�ļ�id And
                                    ((b.�������� = 8 And a.�༭��ʽ = 0) Or (b.�������� In (6, 7, 8) And a.�༭��ʽ = 1)) And Rownum < 2),
                              0) As ����ǩ��
                From ���Ӳ�����¼ A, ���ű� D
                Where a.����id = n_����id And a.��ҳid = n_����id And a.����id = d.Id And Exists
                 (Select 1 --����
                       From ���Ӳ������� C
                       Where a.Id = c.�ļ�id And ((c.�������� = 4 And c.Ҫ������ = '����') Or (c.�������� = 4 And c.Ҫ������ = '�Ա�') Or
                             (c.�������� = 4 And c.Ҫ������ = '����') Or (c.�������� = 2 And c.Ҫ������ = '����') Or
                             (c.�������� = 2 And c.Ҫ������ = '�Ա�') Or (c.�������� = 2 And c.Ҫ������ = '����'))) And Rownum < 2) Loop
  
    --��ȡ���а����������Ա�����Ҫ�صĲ���
    If r_Rec.����ǩ�� = 1 Then
      --��������ǩ����������
      v_Error  := '��д�Ĳ����Ѿ����й�����ǩ��,���ܽ��в�����Ϣ�޸Ĳ�����';
      Json_Out := Zljsonout(v_Error, 1);
      Return;
    End If;
  End Loop;

  --������¼�������
  If Nvl(n_����, 0) = 0 Then
    For r_Rec In (Select /*+ RULE */
                  Distinct d.���� ����, a.Id, a.��������, a.��������, a.���ʱ��, a.������,
                           Nvl((Select Decode(a.�༭��ʽ, 0, Decode(Substr(b.��������, 1, 1), '2', 1, 0),
                                               Decode(Substr(b.��������, Instr(b.��������, '|'), 1), '2', 1, 0))
                                From ���Ӳ������� B
                                Where a.Id = b.�ļ�id And ((b.�������� = 8 And �༭��ʽ = 0) Or (b.�������� In (6, 7, 8) And �༭��ʽ = 1)) And
                                      Rownum < 2), 0) As ����ǩ��
                  From ���Ӳ�����¼ A, ���ű� D, ���˹Һż�¼ E
                  Where a.����id = n_����id And a.��ҳid = n_����id And a.����id = d.Id And a.����id = e.����id And Exists
                   (Select 1 --����
                         From ���Ӳ������� C
                         Where a.Id = c.�ļ�id And
                               ((c.�������� = 2 And Instr(c.�����ı�, e.����) > 0) Or (c.�������� = 1 And Instr(c.�����ı�, e.����) > 0))) And
                        Rownum < 2) Loop
    
      --��ȡ���а��������Ĳ���
      If r_Rec.����ǩ�� = 1 Then
        --��������ǩ����������
        v_Error  := '��д�Ĳ����Ѿ����й�����ǩ��,���ܽ��в�����Ϣ�޸Ĳ�����';
        Json_Out := Zljsonout(v_Error, 1);
        Return;
      End If;
    End Loop;
  Else
    For r_Rec In (Select /*+ RULE */
                  Distinct d.���� ����, a.Id, a.��������, a.��������, a.���ʱ��, a.������,
                           Nvl((Select Decode(a.�༭��ʽ, 0, Decode(Substr(b.��������, 1, 1), '2', 1, 0),
                                               Decode(Substr(b.��������, Instr(b.��������, '|'), 1), '2', 1, 0))
                                From ���Ӳ������� B
                                Where a.Id = b.�ļ�id And ((b.�������� = 8 And �༭��ʽ = 0) Or (b.�������� In (6, 7, 8) And �༭��ʽ = 1)) And
                                      Rownum < 2), 0) As ����ǩ��
                  From ���Ӳ�����¼ A, ���ű� D, ������ҳ E
                  Where a.����id = n_����id And a.��ҳid = n_����id And a.����id = d.Id And a.����id = e.����id And Exists
                   (Select 1 --����
                         From ���Ӳ������� C
                         Where a.Id = c.�ļ�id And
                               ((c.�������� = 2 And Instr(c.�����ı�, e.����) > 0) Or (c.�������� = 1 And Instr(c.�����ı�, e.����) > 0))) And
                        Rownum < 2) Loop
    
      --��ȡ���а��������Ĳ���
      If r_Rec.����ǩ�� = 1 Then
        --��������ǩ����������
        v_Error  := '��д�Ĳ����Ѿ����й�����ǩ��,���ܽ��в�����Ϣ�޸Ĳ�����';
        Json_Out := Zljsonout(v_Error, 1);
        Return;
      End If;
    End Loop;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Updpatbaseinfocheck;
/

Create Or Replace Procedure Zl_Cissvr_Getfeeitem
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡ������Ŀ��Ӧ���շ�ϸĿ��ϸ
  --��Σ�Json_In:��ʽ
  --  input
  --      nodeno                                     C 0 վ�㣬�ɲ���
  --      plcdept_id                                 N 1 ��������id
  --      item_list[]                                ��Ŀ�б���������ĸ���������ϸ��Ӧ��������Ŀ��Ϣ��ִ�п�����Ϣ
  --          apply_id                               C 1 ������ţ������������ţ�Ψһ��ʶһ������
  --          apply_type                             N 1 �������1-��ҩ���ҩ��2-��ҩ��3-���飬4-����(����Ƥ�ԣ����ˣ���ҩ��)��5-���
  --          cure_info                              һ��������Ŀ��Ϣ�����Բ���
  --              cure_item_id                       N 1 ������Ŀid,��ͨ������Ŀ��Ӧ��������Ŀid
  --              cure_exedept_id                    N 1 ִ�п���id,��ͨ������Ŀ��Ӧ��������Ŀid
  --          lis_info                               ������Ŀ��Ϣ
  --              lis_items                          C 1 ������Ŀ��������Ŀid�����Դ��������ƴ������Ϊ���ʱ��ʾһ���ɼ�
  --              lis_exedept_id                     N 1 ������Ŀ��Ӧ��ִ�п���id
  --              lis_collect_item_id                N 1 ����ɼ���Ŀid
  --              lis_collect_exedept_id             N 1 ����ɼ���Ŀ��Ӧ�Ĳɼ�ִ�п���id
  --              lis_spcm                           C 1 ����ɼ��걾
  --              emergency_tag                      N 1 ������ʶ������ǰ�����Ƿ��ǽ���ִ��
  --          pacs_info                              �����Ŀ��Ϣ
  --              pacs_item_id                       N 1 �����Ŀid
  --              pacs_exedept_id                    N 1 �����Ŀִ�п���id
  --              pacs_part_list[]                   �����Ŀ�Ĳ�λ�����б��ɲ�����������ֻ��һ������
  --                  part_name                      C 1 ��鲿λ���ơ���ӦZLHIS�걾������
  --                  part_way                       C 1 ��鷽�����ơ���ӦZLHIS��鷽����
  --          drug_info                              ҩƷ����ҩ��Ŀ��Ϣ
  --              drug_use_item_id                   N 1 ҩƷ��ҩ;����Ŀid,�����䷽���������Ŀid
  --              drug_use_exedept_id                N 1 ҩƷ��ҩ;����Ŀ��Ӧ��ִ�п���id
  --              drug_decoction_id                  N 1 �巨��Ŀid����ҩ�䷽ʱ�Ŵ���
  --              drug_decoction_exedept_id          N 1 �巨��Ŀ��Ӧ��ִ�п���id���Բ���������ʱ��Ϊ������ִ�еĶ���

  --����: Json_Out,��ʽ����
  --  output
  --    code                      N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                   C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    item_list[]     ҽ���б���������ĸ���������ϸ
  --          apply_id            C 1 ������ţ������������ţ�Ψһ��ʶһ������������
  --          cisitem_id          N 1 ������Ŀid
  --          fee_item_id         N 1 �շ�ϸĿid
  --          exedept_id          N 0 ִ�в���id����������Ŀ��Ӧ����ҩƷ�����ķ���ʱ��Ҫ�ⲿ��ȡ������ʱ��ҽ��������Ŀ��ִ�п���idΪ׼����ΪҩƷ�����ķ���ʱ���������Ϊ��ʱ����ָ������ʱ�ɵ��÷����ж��μ���
  --          quantity            N 1 �շ������������շѹ�ϵ�еĶ�������
  --          part_name           C 0 ��鲿λ���ƣ�ZLHIS�з��ö��մ��ڰ���λ����Ŀ����
  --          part_way            C 0 ��鷽������
  --    reject_list[] ����Ŀ�ų��������Ҫ����Բɼ���Ŀ�ϵİ󶨷��ã���Ϊ�Ϲܵ�ԭ����������ɼ����ú��Թܷ��ò�����ȡ�����б��¼�Ϲ����
  --          apply_id            C 1 ������ţ�������ȡ���������
  --          reject_id           C 1 ���ϲ���������ţ������������ţ�Ψһ��ʶһ������������    
  ---------------------------------------------------------------------------

  v_Nodeno       Varchar2(2000);
  v_���ʽ���� Varchar2(2000);
  v_�۸�ȼ�     Varchar2(100);
  v_��ͨ�ȼ�     Varchar2(100);
  v_ҩƷ�ȼ�     Varchar2(100);
  v_���ĵȼ�     Varchar2(100);
  v_Pricegrade   Varchar2(500);

  --LIS�ɼ���Ŀ�շѶ���
  Type Rs_Collection Is Record(
    �����ʶ       Varchar2(4000),
    �����ų�       Varchar2(4000),
    ��Ŀid         Number(18),
    ִ�п���id     Number(18),
    ����           Varchar2(200),
    ��������id     Number(18),
    �ɼ���Ŀid     Number(18),
    �ɼ�ִ�п���id Number(18),
    �ɼ��걾       Varchar2(4000),
    �շ�ϸĿid     Number(18),
    �շ�����       Number(16, 5),
    ����           Number(16, 5),
    ��λ           Varchar2(4000),
    ����           Varchar2(4000),
    ������־       Number(1),
    �շ����       Varchar2(10),
    �ɼ�����       Varchar2(300),
    �շ���Ŀ����   Varchar2(300),
    �շѵ�λ       Varchar2(300),
    �շѷ�ʽ       Number(1));
  Type t_Col Is Table Of Rs_Collection;
  r_Lis      t_Col; --�ɼ���Ŀ�շѶ��ջ����б�
  r_Item     t_Col; --�����շѶ��� 
  r_�ɼ��ų� t_Col;

  --��ͨ��Ŀ
  Cursor c_��ͨ
  (
    P��Ŀid Number,
    P����id Number
  ) Is
    Select a.������Ŀid, a.�շ���Ŀid, a.�շ�����, a.�շѷ�ʽ, a.ִ�п���id, b.����id, c.��� �շ����, c.���� �շ���Ŀ����, c.���㵥λ �շѵ�λ
    From (Select c.������Ŀid, c.�շ���Ŀid, c.�շ�����, c.�շѷ�ʽ, c.ִ�п���id
           From (Select c.������Ŀid, c.�շ���Ŀid, c.��������, c.�շ�����, c.���ж���, c.������Ŀ, c.�շѷ�ʽ, c.���ÿ���id, P����id ִ�п���id,
                         Max(Nvl(c.���ÿ���id, 0)) Over(Partition By c.������Ŀid, c.��鲿λ, c.��鷽��, c.��������) As Top
                  From �����շѹ�ϵ C
                  Where c.������Ŀid = P��Ŀid And c.��鲿λ Is Null And c.��鷽�� Is Null And
                        (c.���ÿ���id Is Null And Nvl(c.������Դ, 0) = 0 Or c.���ÿ���id = P����id And c.������Դ = 1)) C
           Where Nvl(c.���ÿ���id, 0) = c.Top) A, �շ���ĿĿ¼ C, �������� B
    Where a.�շ���Ŀid = c.Id And c.������� In (1, 3) And (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null) And
          (c.վ�� = v_Nodeno Or c.վ�� Is Null) And c.Id = b.����id(+);

  --��鲿λ����
  Cursor c_���
  (
    P��Ŀid Number,
    P����id Number,
    P��λ   Varchar2,
    P����   Varchar2
  ) Is
    Select a.������Ŀid, a.�շ���Ŀid, a.�շ�����, a.�շѷ�ʽ, a.ִ�п���id, b.����id, c.��� �շ����, c.���� �շ���Ŀ����, c.���㵥λ �շѵ�λ
    From (Select c.������Ŀid, c.�շ���Ŀid, c.�շ�����, c.�շѷ�ʽ, c.ִ�п���id
           From (Select c.������Ŀid, c.�շ���Ŀid, c.��������, c.�շ�����, c.���ж���, c.������Ŀ, c.�շѷ�ʽ, c.���ÿ���id, P����id ִ�п���id,
                         Max(Nvl(c.���ÿ���id, 0)) Over(Partition By c.������Ŀid, c.��鲿λ, c.��鷽��, c.��������) As Top
                  From �����շѹ�ϵ C
                  Where c.������Ŀid = P��Ŀid And c.��鲿λ = P��λ And c.��鷽�� = P���� And
                        (c.���ÿ���id Is Null And Nvl(c.������Դ, 0) = 0 Or c.���ÿ���id = P����id And c.������Դ = 1)) C
           Where Nvl(c.���ÿ���id, 0) = c.Top) A, �շ���ĿĿ¼ C, �������� B
    Where a.�շ���Ŀid = c.Id And c.������� In (1, 3) And (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null) And
          (c.վ�� = v_Nodeno Or c.վ�� Is Null) And c.Id = b.����id(+);

  j_Json       Pljson;
  j_In         Pljson;
  j_Json_Tmp   Pljson;
  j_Jsonlist   Pljson_List;
  n_��������id Number(18);
  n_Count      Number(6);
  n_�������   Number(2); --�������1-��ҩ���ҩ��ҩ;����2-��ҩ��3-���飬4-����(����Ƥ�ԣ����ˣ���ҩ��)��5-���
  v_�����ʶ   Varchar2(4000);
  v_Out_Tmp    Varchar2(32767);

  Function Getitem_Price
  (
    P��Ŀid   Number,
    P�շ���� Varchar2
  ) Return Number As
    --��ȡ�շ���Ŀ�ĵ���
    n_���� Number(16, 5);
  Begin
    If Instr(',5,6,7,', ',' || P�շ���� || ',') > 0 Then
      v_�۸�ȼ� := v_ҩƷ�ȼ�;
    Elsif P�շ���� = '4' Then
      v_�۸�ȼ� := v_���ĵȼ�;
    Else
      v_�۸�ȼ� := v_��ͨ�ȼ�;
    End If;
    n_���� := 0;
    For r_������Ŀ In (Select a.Id As �շ�ϸĿid, b.������Ŀid, c.����, c.�վݷ�Ŀ, b.�ּ�, b.ԭ��, b.�Ӱ�Ӽ���, b.�����շ���, b.ȱʡ�۸�, a.���㵥λ, a.��������,
                          a.���ηѱ�, a.��� As �շ����
                   From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C
                   Where b.�շ�ϸĿid = a.Id And c.Id = b.������Ŀid And Sysdate Between b.ִ������ And
                         Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And a.Id = P��Ŀid And
                         ((b.�۸�ȼ� Is Null And Nvl(v_�۸�ȼ�, '-') = '-') Or
                         (b.�۸�ȼ� = v_�۸�ȼ� Or
                         (b.�۸�ȼ� Is Null And Not Exists
                          (Select 1
                             From �շѼ�Ŀ
                             Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = v_�۸�ȼ� And Sysdate Between ִ������ And
                                   Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))))))
                   Order By �շ�ϸĿid, ������Ŀid) Loop
      n_���� := n_���� + Nvl(r_������Ŀ.�ּ�, 0);
    End Loop;
    Return n_����;
  End;

  Function GetҩƷ����ȱʡִ�п���id
  (
    P��Ŀid   Number,
    P�շ���� Varchar2 --4,5,6,7
  ) Return Number As
    --���ܣ���ȡҩƷ����һ��ȱʡ��ִ�п���id
    v_��������   Varchar2(100);
    n_ִ�п���id Number(18);
  Begin
    If P�շ���� = '4' Then
      v_�������� := '���ϲ���';
    Elsif P�շ���� = '5' Then
      v_�������� := '��ҩ��';
    Elsif P�շ���� = '6' Then
      v_�������� := '��ҩ��';
    Elsif P�շ���� = '7' Then
      v_�������� := '��ҩ��';
    End If;
  
    For R In (Select a.ִ�п���id
              From �շ�ִ�п��� A, ��������˵�� B, ���ű� C
              Where a.ִ�п���id + 0 = b.����id And b.�������� = v_�������� And b.������� In (1, 3) And b.����id = c.Id And
                    (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null) And (a.������Դ Is Null Or a.������Դ = 1) And
                    Nvl(a.��������id, 0) = n_��������id And (c.վ�� = v_Nodeno Or c.վ�� Is Null) And a.�շ�ϸĿid = P��Ŀid
              Order By b.�������, c.����) Loop
      n_ִ�п���id := r.ִ�п���id;
      Exit;
    End Loop;
    Return n_ִ�п���id;
  End;

  Procedure Additem
  (
    P�����ʶ Varchar2,
    P��Ŀid   Number,
    P����id   Number,
    P��λ     Varchar2 := Null,
    P����     Varchar2 := Null
  ) As
    --˵����һ�������շѶ���
    N Number(3);
  Begin
    If P��λ Is Null Then
      For R In c_��ͨ(P��Ŀid, P����id) Loop
        r_Item.Extend;
        N := r_Item.Count;
        r_Item(N).�����ʶ := P�����ʶ;
        r_Item(N).��Ŀid := r.������Ŀid;
        r_Item(N).ִ�п���id := r.ִ�п���id;
        r_Item(N).���� := Null;
        r_Item(N).��������id := Null;
        r_Item(N).�ɼ���Ŀid := Null;
        r_Item(N).�ɼ�ִ�п���id := Null;
        r_Item(N).�շ�ϸĿid := r.�շ���Ŀid;
        r_Item(N).�շ����� := r.�շ�����;
        r_Item(N).��λ := Null;
        r_Item(N).���� := Null;
        r_Item(N).�շѷ�ʽ := r.�շѷ�ʽ;
      
        If r.����id Is Not Null Then
          r_Item(N).ִ�п���id := GetҩƷ����ȱʡִ�п���id(r.�շ���Ŀid, '4');
        Elsif r.�շ���� In ('5', '6', '7') Then
          r_Item(N).ִ�п���id := GetҩƷ����ȱʡִ�п���id(r.�շ���Ŀid, r.�շ����);
        End If;
        r_Item(N).�շ���Ŀ���� := r.�շ���Ŀ����;
        r_Item(N).�շѵ�λ := r.�շѵ�λ;
        r_Item(N).���� := Getitem_Price(r.�շ���Ŀid, r.�շ����);
        r_Item(N).�շ���� := r.�շ����;
      End Loop;
    Else
      For R In c_���(P��Ŀid, P����id, P��λ, P����) Loop
        r_Item.Extend;
        N := r_Item.Count;
        r_Item(N).�����ʶ := P�����ʶ;
        r_Item(N).��Ŀid := r.������Ŀid;
        r_Item(N).ִ�п���id := r.ִ�п���id;
        r_Item(N).���� := Null;
        r_Item(N).��������id := Null;
        r_Item(N).�ɼ���Ŀid := Null;
        r_Item(N).�ɼ�ִ�п���id := Null;
        r_Item(N).�շ�ϸĿid := r.�շ���Ŀid;
        r_Item(N).�շ����� := r.�շ�����;
        r_Item(N).��λ := P��λ;
        r_Item(N).���� := P����;
        r_Item(N).�շѷ�ʽ := r.�շѷ�ʽ;
      
        If r.����id Is Not Null Then
          r_Item(N).ִ�п���id := GetҩƷ����ȱʡִ�п���id(r.�շ���Ŀid, '4');
        Elsif r.�շ���� In ('5', '6', '7') Then
          r_Item(N).ִ�п���id := GetҩƷ����ȱʡִ�п���id(r.�շ���Ŀid, r.�շ����);
        End If;
        r_Item(N).�շ���Ŀ���� := r.�շ���Ŀ����;
        r_Item(N).�շѵ�λ := r.�շѵ�λ;
        r_Item(N).���� := Getitem_Price(r.�շ���Ŀid, r.�շ����);
        r_Item(N).�շ���� := r.�շ����;
      End Loop;
    End If;
  End;

  Procedure Additem_�ɼ�
  (
    P�����ʶ   Varchar2,
    P��Ŀid     Number,
    P����id     Number,
    P������Ŀid Number,
    P�������id Number,
    P����       Number,
    P�ɼ��걾   Varchar2
  ) As
    --�ɼ���Ŀ�շѶ���
    --˵���������շѷ�ʽ �����Թܷ��ã�Ҫ����������� �����Ŀ���ù��룬��Ѫ�ܰ����� ����Ч��������չ�ϵδ��Զ��ᵱ��������ȡ���������շѲ���ȷ
    N                  Number(3);
    v_����             Varchar2(3000);
    n_����id           Number(18);
    n_�����Թܷ������� Number(1) := 0; --�󶨷����� �շѷ�ʽΪ 1-�����Թܷ��� ���ʵķ�������ȡ����ʣ���������ȡ��ʽ�ķ�����Ҫ�������
    n_Ҫ�շ�           Number(1);
    v_�ɼ�����         Varchar2(300);
  Begin
    Select Max(a.����) Into v_�ɼ����� From ������ĿĿ¼ A Where a.Id = P��Ŀid;
    Select Max(a.�Թܱ���) Into v_���� From ������ĿĿ¼ A Where a.Id = P������Ŀid;
    --ֻȡһ��
    For R In (Select ����, ����id From ��Ѫ������ Where ����id Is Not Null And ���� = v_����) Loop
      n_����id := r.����id;
    End Loop;
  
    --�ȴ����з��ò��ң��ж��Ƿ����Ѿ��ɼ���һ���ˣ��ж��Ƿ���Ҫ�����շ�
    For N In 1 .. r_Lis.Count Loop
      If P������Ŀid <> r_Lis(N).��Ŀid And r_Lis(N).�ɼ���Ŀid = P��Ŀid And r_Lis(N).���� = v_���� And r_Lis(N).ִ�п���id = P�������id And r_Lis(N)
        .������־ = Nvl(P����, 0) And r_Lis(N).�ɼ��걾 = P�ɼ��걾 And r_Lis(N).�ɼ�ִ�п���id = P����id Then
        n_�����Թܷ������� := 1; --���ҵ� ���չ��� �����Թ������ ���ʵķ���
      
        r_�ɼ��ų�.Extend;
        r_�ɼ��ų�(r_�ɼ��ų�.Count).�����ʶ := r_Lis(N).�����ʶ;
        r_�ɼ��ų�(r_�ɼ��ų�.Count).�����ų� := P�����ʶ;
      
        Exit;
      End If;
    End Loop;
  
    For R In c_��ͨ(P��Ŀid, P����id) Loop
      n_Ҫ�շ� := 1;
    
      If n_�����Թܷ������� = 1 And r.�շѷ�ʽ = 1 Then
        --�����շѷ�ʽ �����Թܷ��� �ķ�����ϸ�ų���
        n_Ҫ�շ� := 0;
      End If;
    
      If n_Ҫ�շ� = 1 Then
        --���Ϸ�������
        If r.�շѷ�ʽ = 1 And n_����id <> r.����id Then
          n_Ҫ�շ� := 0;
        End If;
      End If;
    
      If n_Ҫ�շ� = 1 Then
      
        r_Lis.Extend;
        N := r_Lis.Count;
        r_Lis(N).�����ʶ := P�����ʶ;
        r_Lis(N).��Ŀid := P������Ŀid;
        r_Lis(N).ִ�п���id := P�������id;
        r_Lis(N).���� := v_����;
        r_Lis(N).��������id := n_����id;
        r_Lis(N).�ɼ���Ŀid := r.������Ŀid;
        r_Lis(N).�ɼ�ִ�п���id := r.ִ�п���id;
        r_Lis(N).�շ�ϸĿid := r.�շ���Ŀid;
        r_Lis(N).�շ����� := r.�շ�����;
        r_Lis(N).������־ := Nvl(P����, 0);
        r_Lis(N).�ɼ��걾 := P�ɼ��걾;
        r_Lis(N).��λ := Null;
        r_Lis(N).���� := Null;
        r_Lis(N).�շѷ�ʽ := r.�շѷ�ʽ;
      
        --׷��ͨ���շ���ϸ��Ŀ��
        r_Item.Extend;
        N := r_Item.Count;
        r_Item(N).�����ʶ := P�����ʶ;
        r_Item(N).��Ŀid := r.������Ŀid;
        r_Item(N).ִ�п���id := r.ִ�п���id;
        r_Item(N).�շ�ϸĿid := r.�շ���Ŀid;
        r_Item(N).�շ����� := r.�շ�����;
      
        If r.����id Is Not Null Then
          r_Item(N).ִ�п���id := GetҩƷ����ȱʡִ�п���id(r.�շ���Ŀid, '4');
        Elsif r.�շ���� In ('5', '6', '7') Then
          r_Item(N).ִ�п���id := GetҩƷ����ȱʡִ�п���id(r.�շ���Ŀid, r.�շ����);
        End If;
        r_Item(N).�ɼ����� := v_�ɼ�����;
        r_Item(N).�ɼ��걾 := P�ɼ��걾;
        r_Item(N).�շ���Ŀ���� := r.�շ���Ŀ����;
        r_Item(N).�շѵ�λ := r.�շѵ�λ;
        r_Item(N).���� := Getitem_Price(r.�շ���Ŀid, r.�շ����);
        r_Item(N).�շ���� := r.�շ����;
        r_Item(N).�ɼ���Ŀid := P��Ŀid;
      End If;
    End Loop;
  End;

  Procedure Getitem(j_Par_In Pljson) As
    --��json��Ŀ�����н���
    n_��Ŀid      Number(18);
    n_����id      Number(18);
    v_��λ        Varchar2(4000);
    v_����        Varchar2(4000);
    v_������Ŀids Varchar2(4000);
    j_List        Pljson_List;
    j_Json_Tmp    Pljson;
    j_Par         Pljson;
  Begin
    v_�����ʶ := j_Par_In.Get_String('apply_id');
    n_������� := j_Par_In.Get_Number('apply_type');
    If n_������� = 1 Then
      j_Par    := j_Par_In.Get_Pljson('drug_info');
      n_��Ŀid := j_Par.Get_Number('drug_use_item_id');
      n_����id := j_Par.Get_Number('drug_use_exedept_id');
      Additem(v_�����ʶ, n_��Ŀid, n_����id, Null, Null);
    Elsif n_������� = 2 Then
      j_Par    := j_Par_In.Get_Pljson('drug_info');
      n_��Ŀid := j_Par.Get_Number('drug_use_item_id');
      n_����id := j_Par.Get_Number('drug_use_exedept_id');
      Additem(v_�����ʶ, n_��Ŀid, n_����id, Null, Null);
      n_��Ŀid := j_Par.Get_Number('drug_decoction_id');
      n_����id := j_Par.Get_Number('drug_decoction_exedept_id');
      Additem(v_�����ʶ, n_��Ŀid, n_����id, Null, Null);
    Elsif n_������� = 3 Then
      j_Par         := j_Par_In.Get_Pljson('lis_info');
      v_������Ŀids := j_Par.Get_String('lis_items');
      n_����id      := j_Par.Get_Number('lis_exedept_id');
      For R In (Select /*+cardinality(j,10) */
                 j.Column_Value ������Ŀid
                From Table(Cast(f_Num2list(v_������Ŀids) As Zltools.t_Numlist)) J) Loop
        If n_��Ŀid Is Null Then
          --һ���ɼ������м�����Ŀ
          n_��Ŀid := r.������Ŀid;
        End If;
        Additem(v_�����ʶ, r.������Ŀid, n_����id, Null, Null);
      End Loop;
      Additem_�ɼ�(v_�����ʶ, j_Par.Get_Number('lis_collect_item_id'), j_Par.Get_Number('lis_collect_exedept_id'), n_��Ŀid,
                 n_����id, j_Par.Get_Number('emergency_tag'), j_Par.Get_String('lis_spcm'));
    Elsif n_������� = 4 Then
      j_Par    := j_Par_In.Get_Pljson('cure_info');
      j_Par    := j_Par_In.Get_Pljson('cure_info');
      n_��Ŀid := j_Par.Get_Number('cure_item_id');
      n_����id := j_Par.Get_Number('cure_exedept_id');
      Additem(v_�����ʶ, n_��Ŀid, n_����id, Null, Null);
    Elsif n_������� = 5 Then
      j_Par    := j_Par_In.Get_Pljson('pacs_info');
      n_��Ŀid := j_Par.Get_Number('pacs_item_id');
      n_����id := j_Par.Get_Number('pacs_exedept_id');
      Additem(v_�����ʶ, n_��Ŀid, n_����id, Null, Null);
      j_List := j_Par.Get_Pljson_List('pacs_part_list');
      If j_List Is Not Null Then
        For I In 1 .. j_List.Count Loop
          j_Json_Tmp := Pljson();
          j_Json_Tmp := Pljson(j_List.Get(I));
          v_��λ     := j_Json_Tmp.Get_String('part_name');
          v_����     := j_Json_Tmp.Get_String('part_way');
          Additem(v_�����ʶ, n_��Ŀid, n_����id, v_��λ, v_����);
        End Loop;
      End If;
    End If;
  End;

  Function Get�ɼ��ų� Return Varchar2 As
    --���ܣ����Ա��ϲ��ķ�����Ŀ����Ҫ��Բɼ���Ŀ���յķ��úϲ����
    --���θ�",reject_list":[{"apply_id":"444","reject_id":"2323"},{},{}...]  
    --          reject_list[] ����Ŀ�ų��������Ҫ����Բɼ���Ŀ�ϵİ󶨷���
    --               apply_id       C 1 ������ţ������������ţ�Ψһ��ʶһ������������
    --               reject_id      C 1 ���ϲ���������ţ������������ţ�Ψһ��ʶһ������������  
  
    v_Jtmp Varchar2(32767);
  Begin
    For I In 1 .. r_�ɼ��ų�.Count Loop
      v_Jtmp := v_Jtmp || ',{"apply_id":"' || r_�ɼ��ų�(I).�����ʶ || '"';
      v_Jtmp := v_Jtmp || ',"reject_id":"' || r_�ɼ��ų�(I).�����ų� || '"';
      v_Jtmp := v_Jtmp || '}';
    End Loop;
    If v_Jtmp Is Not Null Then
      v_Jtmp := ',"reject_list":[' || Substr(v_Jtmp, 2) || ']';
    End If;
    Return v_Jtmp;
  End;

Begin
  --�������
  j_In         := Pljson(Json_In);
  j_Json       := j_In.Get_Pljson('input');
  v_Nodeno     := j_Json.Get_String('nodeno');
  n_��������id := j_Json.Get_Number('plcdept_id');
  j_Jsonlist   := j_Json.Get_Pljson_List('item_list');

  v_Nodeno       := j_Json.Get_String('site_no');
  v_���ʽ���� := j_Json.Get_String('mdlpay_mode_name');
  If Nvl(v_Nodeno, '-') = '-' And Nvl(v_���ʽ����, '-') = '-' Then
    v_��ͨ�ȼ� := Null;
    v_ҩƷ�ȼ� := Null;
    v_���ĵȼ� := Null;
  Else
    v_Pricegrade := Zl_Get_Pricegrade_s(v_Nodeno, v_���ʽ����);
    For c_�۸�ȼ� In (Select Rownum As ���, Column_Value As �۸�ȼ� From Table(f_Str2list(v_Pricegrade, '|'))) Loop
      If c_�۸�ȼ�.��� = 1 Then
        v_��ͨ�ȼ� := c_�۸�ȼ�.�۸�ȼ�;
      End If;
      If c_�۸�ȼ�.��� = 2 Then
        v_ҩƷ�ȼ� := c_�۸�ȼ�.�۸�ȼ�;
      End If;
      If c_�۸�ȼ�.��� = 3 Then
        v_���ĵȼ� := c_�۸�ȼ�.�۸�ȼ�;
      End If;
    End Loop;
  End If;

  If Not j_Jsonlist Is Null Then
    n_Count    := j_Jsonlist.Count;
    r_Item     := t_Col();
    r_Lis      := t_Col();
    r_�ɼ��ų� := t_Col();
    For I In 1 .. n_Count Loop
      j_Json_Tmp := Pljson();
      j_Json_Tmp := Pljson(j_Jsonlist.Get(I));
      Getitem(j_Json_Tmp);
    End Loop;
  End If;

  v_Out_Tmp := Null;
  For I In 1 .. r_Item.Count Loop
    v_Out_Tmp := v_Out_Tmp || ',{"apply_id":"' || r_Item(I).�����ʶ || '"';
    v_Out_Tmp := v_Out_Tmp || ',"cisitem_id":' || Nvl(r_Item(I).��Ŀid, 0);
    v_Out_Tmp := v_Out_Tmp || ',"fee_item_id":' || Nvl(r_Item(I).�շ�ϸĿid, 0);
    v_Out_Tmp := v_Out_Tmp || ',"fee_item_type":"' || r_Item(I).�շ���� || '"';
    v_Out_Tmp := v_Out_Tmp || ',"fee_item_name":"' || r_Item(I).�շ���Ŀ���� || '"';
    v_Out_Tmp := v_Out_Tmp || ',"item_unit":"' || r_Item(I).�շѵ�λ || '"';
    v_Out_Tmp := v_Out_Tmp || ',"price":' || Zljsonstr(r_Item(I).����, 1);
    v_Out_Tmp := v_Out_Tmp || ',"exedept_id":' || Nvl(r_Item(I).ִ�п���id, 0);
    v_Out_Tmp := v_Out_Tmp || ',"quantity":' || Nvl(r_Item(I).�շ�����, 0);
    v_Out_Tmp := v_Out_Tmp || ',"part_name":"' || r_Item(I).��λ || '"';
    v_Out_Tmp := v_Out_Tmp || ',"part_way":"' || r_Item(I).���� || '"';
    v_Out_Tmp := v_Out_Tmp || ',"lis_spcm":"' || r_Item(I).�ɼ��걾 || '"';
    v_Out_Tmp := v_Out_Tmp || ',"lis_collect_way":"' || r_Item(I).�ɼ����� || '"';
    v_Out_Tmp := v_Out_Tmp || ',"lis_collect_item_id":' || Nvl(r_Item(I).�ɼ���Ŀid, 0);
    v_Out_Tmp := v_Out_Tmp || '}';
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || Substr(v_Out_Tmp, 2) || ']' || Get�ɼ��ų� || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getfeeitem;
/

Create Or Replace Procedure Zl_Cissvr_Outsendapply
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  ---------------------------------------------------------------------------
  --���ܣ�����������ҩƷ(�䷽)/����/һ��������Ŀ���룬���ɡ�ҽ��/�Ƽ�/����/ִ�С���ص��ٴ�������
  --��Σ�Json_In:��ʽ
  --  input  

  --����: Json_Out,��ʽ����
  --  output
  --    code                N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    advice_list[]       ҽ���б���������ĸ���������ϸ
  --      apply_id          C 1 ������ţ������������ţ�Ψһ��ʶһ������������
  --      fee_no            C 1 ���õ��ݺ�
  --      lis_bar_code      C 1 ������Ŀ���룬����������Ŀ���д˽��
  --      order_list[]      ������ҽ���б�
  --      order_id          N 1 ҽ��id,zlhis�����ɵ�ҽ��id
  --      order_related_id  N 1 ���id,zlhis�����ɵ����id
  --      cisitem_id        N 1 ������Ŀid
  ---------------------------------------------------------------------------
  --����б� 
  Type r_Diag Is Record(
    Diag_Num      Varchar2(200), --������     
    Csd_Code      Varchar2(200), --��ϱ��룬��ҽ��ϱ��� ��������Ŀ¼
    Icd10_Code    Varchar2(200), --�������� ICD10 ��ҽ
    Dz_Note       Varchar2(4000), --�������
    Syndrome      Varchar2(4000), --��ҽ֤��
    Disease_Time  Varchar2(30), --����ʱ�� ���ڸ�ʽ
    Diagnostician Varchar2(200), --���ҽ������¼��
    �Ƿ����      Number(1), --     ���������е���ϼ�¼,
    �������      Number(1), --������ͣ�1-��ҽ��ϣ�11-��ҽ���       
    �������      Varchar2(4000),
    ����id        Number(18),
    ��ϴ���      Number(2),
    ���ids       Varchar2(4000),
    ҽ��id        Number(18),
    ���id        Number(18));

  Type r_Apply Is Record(
    Apply_Id     Varchar2(200), --����Ψһ��ʶ�����˱�־�������͵�NO��
    Apply_Type   Number(1), --�������1-��ҩ���ҩ��2-��ҩ��3-���飬4-����(����Ƥ�ԣ����ˣ���ҩ��)��5-���
    Diag_Nums    Varchar2(4000), --������
    Zlhis��¼ids Varchar2(4000),
    Serial_Num   Number(18),
    Group_Sno    Number(18),
    ID           Number(18),
    ���id       Number(18),
    ǰ��id       Number(18),
    ������Դ     Number(1),
    ����id       Number(18),
    ��ҳid       Number(5),
    �Һŵ�       Varchar2(8),
    Ӥ��         Number(3),
    ����         Varchar2(100),
    �Ա�         Varchar2(4),
    ����         Varchar2(20),
    ���˿���id   Number(18),
    ���         Number(18),
    ҽ��״̬     Number(3),
    ҽ����Ч     Number(1),
    �������     Varchar2(1),
    ������Ŀid   Number(18),
    �걾��λ     Varchar2(60),
    ��鷽��     Varchar2(30),
    �շ�ϸĿid   Number(18),
    ����         Number(16, 5),
    ��������     Number(16, 5),
    �״�����     Number(16, 5),
    �ܸ�����     Number(16, 5),
    ҽ������     Varchar2(1000),
    ҽ������     Varchar2(200),
    ִ�п���id   Number(18),
    Ƥ�Խ��     Varchar2(10),
    ִ��Ƶ��     Varchar2(20),
    Ƶ�ʴ���     Number(3),
    Ƶ�ʼ��     Number(3),
    �����λ     Varchar2(4),
    ִ��ʱ�䷽�� Varchar2(100),
    �Ƽ�����     Number(1),
    ִ������     Number(1),
    ִ�б��     Number(1),
    ��˱��     Number(1),
    �ɷ����     Number(3),
    ������־     Number(1),
    ��ʼִ��ʱ�� Date,
    ��������id   Number(18),
    ����ҽ��     Varchar2(41),
    ����ʱ��     Date,
    ����ʱ��     Date,
    �Ƿ��ϴ�     Number(1),
    �����     Number(1),
    ���δ�ӡ     Number(1),
    ժҪ         Varchar2(1000),
    ��Ѽ���     Number(1),
    ��ҩĿ��     Number(1),
    ��ҩ����     Varchar2(1000),
    ����˵��     Varchar2(1000),
    ����         Varchar2(100),
    �䷽id       Number(18),
    ----������
    NO       Varchar2(60),
    ��¼��� Number(3),
    ���ͺ�   Number(18),
    �������� Number(16, 5), --�����ҩƷ�����ǰ����㵥λ����ģ���Ҫ����
    �״�ʱ�� Date,
    ĩ��ʱ�� Date,
    ����ʱ�� Date,
    �������� Varchar2(600),
    
    Firstrow Number(1) --ZLHIS��һ��ҽ���еĵ�һ��
    );

  Type r_Price Is Record(
    
    Apply_Id   Varchar2(200), --����Ψһ��ʶ�����˱�־�������͵�NO��
    ������Ŀid Number(18),
    �������id Number(18),
    �ɼ���Ŀid Number(18),
    �ɼ�����id Number(18),
    �ɼ��걾   Varchar(200),
    ������־   Number(1),
    Ӥ��       Number(3),
    ����       Varchar2(100), --���μӹ�
    ��������id Number(18), --���μӹ�
    ��������   Varchar(200), --���μӹ�    
    ҽ��id     Number(18),
    ������Ŀid Number(18),
    �շ�ϸĿid Number(18),
    ִ�п���id Number(18),
    �շѷ�ʽ   Number(1));

  Type t_Apply Is Table Of r_Apply;
  Type t_Diag Is Table Of r_Diag;
  Type t_Price Is Table Of r_Price;

  Rsdiag         t_Diag := t_Diag();
  Rs���ҽ��     t_Diag := t_Diag();
  Rsdiagnew      t_Diag := t_Diag();
  Rsap           t_Apply := t_Apply();
  r_Base         r_Apply;
  Rs����         t_Price := t_Price();
  n_����id       Number(18); --���˹Һż�¼�Һ�id
  v_Nodeno       Varchar2(1000);
  d_����ʱ��     Date;
  n_�Ƿ��������� Number; --������Ŀ�Ƿ����������־��null/0�����ɣ�1-Ҫ���ɣ���������ԴΪ��첡��ʱ�Ŵ��룬������������ZLHISϵͳ����Ϊ׼����������ҽ����������������

  --v_ִ��ʱ�䷽�� Varchar2(200); --�ɲ�����ԭ����Ҫ��advice_frequency���ʹ�ã���Ҫ���ڼ���ҽ��ִ��ʱ��㣬������ʱ�ɸ��� Ƶ�ʱ��� ȡ��ȱʡִ��ʱ�䷽��
  --                              ����ִ��  ÿ������ 1/8-3/8-5/8 �� 1/8:00-3/8:00-5/8:00 ��ʾ��ÿ������һ��8:00,��������8:00,�������8:00�⼸��ʱ��ִ��
  --                              ����ִ��  ÿ������ 8-12-16 �� 8:00-12:00-16:00 ��ʾ��ÿ��8:00,12:00,16:00�⼸��ʱ��ִ�� 
  --                                    ����һ�� 1/8 �� 1/8:00 ��ʾ��ÿ�����еĵ�1��8:00���ʱ��ִ��
  --                              ��ʱִ�� ÿСʱ���� 1:20-1:40 ��ʾ��ÿСʱ�ڵ�20��40����������ʱ��ִ�� 
  --                                    ��Сʱһ�� 2:30 �� 1:30 �� 1:00 ��ʾ��ÿ��Сʱ�ڵĵ�2�ĸ�Сʱ��30�������ʱ��ִ�� ����ÿ��Сʱ�ڵĵ�1�ĸ�Сʱ��30�������ʱ��ִ�� ����ÿ��Сʱ�ڵĵ�1�ĸ�Сʱ���ʱ��ִ��

  v_Input      Pljson;
  j_Tmp        Pljson;
  Jl_Tmp       Pljson_List;
  j_Advicelist Pljson_List;
  Jl_Diag      Pljson_List;
  Idx          Number(6);
  j_Adviceitem Pljson;
  l_ҽ��id     t_Numlist := t_Numlist();
  n_ҽ��id     Number(18);
  n_����id     Number(18) := 0; --ȫ�ֱ����� 
  n_���       Number;
  v_����       Varchar2(2000);
  v_���       Varchar2(2000);
  n_����ϵ��   Number(16, 5);
  n_���ͺ�     ����ҽ������.���ͺ�%Type;
  v_No         ����ҽ������.No%Type;

  n_Drug_Rows Number(6);
  n_Pre���   Number(6);

  v_Lisitems Varchar2(4000);

  v_Out  Varchar2(32767);
  v_Tmp1 Varchar2(32767);

  --��ͨ��Ŀ  P��Ŀid - ������ĿĿ¼ ��P����id - �������ң��������id��
  Cursor c_��ͨ
  (
    P��Ŀid Number,
    P����id Number
  ) Is
    Select a.������Ŀid, a.�շ���Ŀid, a.�շ�����, a.�շѷ�ʽ, a.ִ�п���id, b.����id, c.��� �շ����
    From (Select c.������Ŀid, c.�շ���Ŀid, c.�շ�����, c.�շѷ�ʽ, c.ִ�п���id
           From (Select c.������Ŀid, c.�շ���Ŀid, c.��������, c.�շ�����, c.���ж���, c.������Ŀ, c.�շѷ�ʽ, c.���ÿ���id, P����id ִ�п���id,
                         Max(Nvl(c.���ÿ���id, 0)) Over(Partition By c.������Ŀid, c.��鲿λ, c.��鷽��, c.��������) As Top
                  From �����շѹ�ϵ C
                  Where c.������Ŀid = P��Ŀid And c.��鲿λ Is Null And c.��鷽�� Is Null And
                        (c.���ÿ���id Is Null And Nvl(c.������Դ, 0) = 0 Or c.���ÿ���id = P����id And c.������Դ = 1)) C
           Where Nvl(c.���ÿ���id, 0) = c.Top) A, �շ���ĿĿ¼ C, �������� B
    Where a.�շ���Ŀid = c.Id And c.������� In (1, 3) And (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null) And
          (c.վ�� = v_Nodeno Or c.վ�� Is Null) And c.Id = b.����id(+);

  --�ֽ�ʱ�� 
  Function f_Calc�����ֽ�ʱ��
  (
    ����_In     In ����ҽ����¼.�ܸ�����%Type,
    ��ʼʱ��_In In Date,
    ��ֹʱ��_In In Date,
    ִ��ʱ��_In In ����ҽ����¼.ִ��ʱ�䷽��%Type,
    Ƶ�ʼ��_In In ����ҽ����¼.Ƶ�ʼ��%Type,
    �����λ_In In ����ҽ����¼.�����λ%Type
  ) Return Varchar2 Is
    v_Detailtime Varchar2(4000);
    n_First      Number(1);
    v_First      ����ҽ����¼.ִ��ʱ�䷽��%Type;
    v_Normal     ����ҽ����¼.ִ��ʱ�䷽��%Type;
    v_Mtime      Varchar(100);
    v_Rtime      Varchar(100);
    v_Curtime    Date;
    v_Tmptime    Date;
    n_Cnt        Number; --������ 
    --��ĳʱ�������ܵ�����һ������ 
    Function Getweekbase(v_Time Date) Return Date Is
      v_Week Number(1);
    Begin
      v_Week := To_Number(To_Char(v_Time, 'D'));
      v_Week := v_Week - 1;
      If v_Week = 0 Then
        v_Week := 7;
      End If;
      Return(Trunc(v_Time - (v_Week - 1)));
    End;
  Begin
  
    n_Cnt := 0;
  
    --ִ��ʱ�䷽����׼ 
    If Nvl(Instr(ִ��ʱ��_In, ','), 0) > 0 Then
      v_First  := Substr(ִ��ʱ��_In, 1, Instr(ִ��ʱ��_In, ',') - 1);
      v_Normal := Substr(ִ��ʱ��_In, Instr(ִ��ʱ��_In, ',') + 1);
    Else
      v_First  := Null;
      v_Normal := ִ��ʱ��_In;
    End If;
  
    If �����λ_In = '��' Then
      v_Curtime := Getweekbase(��ʼʱ��_In); --����ִ��ʱ��ҽ����ʼ���ܵ�����һ��Ϊ��׼ 
    Else
      v_Curtime := ��ʼʱ��_In;
    End If;
    If �����λ_In = '��' Then
      If v_First Is Not Null Then
        If v_Curtime = Getweekbase(��ʼʱ��_In) Then
          n_First := 1;
        End If;
      End If;
      While v_Curtime <= ��ֹʱ��_In And n_Cnt < ����_In Loop
        If Nvl(n_First, 0) = 1 Then
          v_Rtime := v_First || '-';
        Else
          v_Rtime := v_Normal || '-';
        End If;
        n_First := 0;
        --1/8:00-3/8:00-5/8:00 
        While v_Rtime Is Not Null Loop
          v_Mtime   := Substr(v_Rtime, 1, Instr(v_Rtime, '-') - 1);
          v_Tmptime := v_Curtime + To_Number(Substr(v_Mtime, 1, Instr(v_Mtime, '/') - 1)) - 1;
          v_Mtime   := Substr(v_Mtime, Instr(v_Mtime, '/') + 1);
          If Instr(v_Mtime, ':') = 0 Then
            v_Mtime := v_Mtime || ':00';
          End If;
          v_Tmptime := Trunc(v_Tmptime) + (To_Date(v_Mtime, 'HH24:MI:SS') - Trunc(To_Date(v_Mtime, 'HH24:MI:SS')));
          If v_Tmptime >= ��ʼʱ��_In And v_Tmptime <= ��ֹʱ��_In Then
            v_Detailtime := v_Detailtime || ',' || To_Char(v_Tmptime, 'YYYY-MM-DD HH24:MI:SS');
            n_Cnt        := n_Cnt + 1;
            If n_Cnt >= ����_In Then
              Exit;
            End If;
          Elsif v_Tmptime > ��ֹʱ��_In Then
            Exit;
          End If;
          v_Rtime := Substr(v_Rtime, Instr(v_Rtime, '-') + 1);
        End Loop;
        v_Curtime := Trunc(v_Curtime + 7);
      End Loop;
    Elsif �����λ_In = '��' Then
      If v_First Is Not Null Then
        If Trunc(��ʼʱ��_In) = Trunc(��ʼʱ��_In) Then
          n_First := 1;
        End If;
      End If;
      While v_Curtime <= ��ֹʱ��_In And n_Cnt < ����_In Loop
        If Nvl(n_First, 0) = 1 Then
          v_Rtime := v_First || '-';
        Else
          v_Rtime := v_Normal || '-';
        End If;
        n_First := 0;
        If Ƶ�ʼ��_In = 1 Then
          --8:00-12:00-14:00��8-12-14 
          While v_Rtime Is Not Null Loop
            v_Mtime := Substr(v_Rtime, 1, Instr(v_Rtime, '-') - 1);
            If Instr(v_Mtime, ':') = 0 Then
              v_Mtime := v_Mtime || ':00';
            End If;
            v_Tmptime := Trunc(v_Curtime) + (To_Date(v_Mtime, 'HH24:MI:SS') - Trunc(To_Date(v_Mtime, 'HH24:MI:SS')));
            If v_Tmptime >= ��ʼʱ��_In And v_Tmptime <= ��ֹʱ��_In Then
            
              v_Detailtime := v_Detailtime || ',' || To_Char(v_Tmptime, 'YYYY-MM-DD HH24:MI:SS');
              n_Cnt        := n_Cnt + 1;
              If n_Cnt >= ����_In Then
                Exit;
              
              End If;
            Elsif v_Tmptime > ��ֹʱ��_In Then
              Exit;
            End If;
            v_Rtime := Substr(v_Rtime, Instr(v_Rtime, '-') + 1);
          End Loop;
        Else
          --1/8:00-1/15:00-2/9:00 
          While v_Rtime Is Not Null Loop
            v_Mtime   := Substr(v_Rtime, 1, Instr(v_Rtime, '-') - 1);
            v_Tmptime := v_Curtime + To_Number(Substr(v_Mtime, 1, Instr(v_Mtime, '/') - 1)) - 1;
            v_Mtime   := Substr(v_Mtime, Instr(v_Mtime, '/') + 1);
            If Instr(v_Mtime, ':') = 0 Then
              v_Mtime := v_Mtime || ':00';
            End If;
            v_Tmptime := Trunc(v_Tmptime) + (To_Date(v_Mtime, 'HH24:MI:SS') - Trunc(To_Date(v_Mtime, 'HH24:MI:SS')));
            If v_Tmptime >= ��ʼʱ��_In And v_Tmptime <= ��ֹʱ��_In Then
            
              v_Detailtime := v_Detailtime || ',' || To_Char(v_Tmptime, 'YYYY-MM-DD HH24:MI:SS');
              n_Cnt        := n_Cnt + 1;
              If n_Cnt >= ����_In Then
                Exit;
              
              End If;
            Elsif v_Tmptime > ��ֹʱ��_In Then
              Exit;
            End If;
            v_Rtime := Substr(v_Rtime, Instr(v_Rtime, '-') + 1);
          End Loop;
        End If;
        v_Curtime := Trunc(v_Curtime + Ƶ�ʼ��_In); --��ΪLoop����ע��Ҫȡ�� 
      End Loop;
    Elsif �����λ_In = 'Сʱ' Then
      --10:00-20:00-40:00��10-20-40��02:30 
      While v_Curtime <= ��ֹʱ��_In And n_Cnt < ����_In Loop
      
        v_Rtime := ִ��ʱ��_In || '-';
        While v_Rtime Is Not Null Loop
          v_Mtime := Substr(v_Rtime, 1, Instr(v_Rtime, '-') - 1);
          If Instr(v_Mtime, ':') = 0 Then
            v_Tmptime := v_Curtime + (To_Number(v_Mtime) - 1) / 24;
          Else
            v_Tmptime := v_Curtime + (To_Number(Substr(v_Mtime, 1, Instr(v_Mtime, ':') - 1)) - 1) / 24 +
                         To_Number(Substr(v_Mtime, Instr(v_Mtime, ':') + 1)) / 60 / 24;
          End If;
        
          If v_Tmptime >= ��ʼʱ��_In And v_Tmptime <= ��ֹʱ��_In Then
          
            v_Detailtime := v_Detailtime || ',' || To_Char(v_Tmptime, 'YYYY-MM-DD HH24:MI:SS');
            n_Cnt        := n_Cnt + 1;
            If n_Cnt >= ����_In Then
              Exit;
            End If;
          
          Elsif v_Tmptime > ��ֹʱ��_In Then
            Exit;
          End If;
          v_Rtime := Substr(v_Rtime, Instr(v_Rtime, '-') + 1);
        End Loop;
        v_Curtime := v_Curtime + Ƶ�ʼ��_In / 24;
      End Loop;
    Elsif �����λ_In = '����' Then
      --��ִ��ʱ�� 
      While v_Curtime <= ��ֹʱ��_In And n_Cnt < ����_In Loop
        v_Tmptime := v_Curtime;
        If v_Tmptime >= ��ʼʱ��_In And v_Tmptime <= ��ֹʱ��_In Then
          v_Detailtime := v_Detailtime || ',' || To_Char(v_Tmptime, 'YYYY-MM-DD HH24:MI:SS');
          n_Cnt        := n_Cnt + 1;
          If n_Cnt >= ����_In Then
            Exit;
          End If;
        Elsif v_Tmptime > ��ֹʱ��_In Then
          Exit;
        End If;
        v_Curtime := v_Curtime + Ƶ�ʼ��_In / (24 * 60);
      End Loop;
    End If;
    If v_Detailtime Is Not Null Then
      v_Detailtime := Substr(v_Detailtime, 2);
    End If;
    Return(v_Detailtime);
  End;

  --����ָ���������� 
  Function f_Intex(Num_In In Number) Return Number Is
    n_Num Number;
  Begin
    Select Round(Num_In) Into n_Num From Dual;
    If Num_In > n_Num Then
      n_Num := n_Num + 1;
    End If;
    Return(n_Num);
  End;

  Procedure Getҽ����϶�Ӧ(Pd Number) As
    --  pd Ϊ��¼�� rsap ���±�����ֵ 
    v_��Ŵ� Varchar2(4000);
    v_��ϴ� Varchar2(4000);
  Begin
    v_��Ŵ� := Rsap(Pd).Diag_Nums;
    If v_��Ŵ� Is Not Null Then
      v_��Ŵ� := ',' || v_��Ŵ� || ',';
      For I In 1 .. Rsdiag.Count Loop
        If Instr(v_��Ŵ�, ',' || Rsdiag(I).Diag_Num || ',') > 0 Then
          v_��ϴ� := v_��ϴ� || ',' || Rsdiag(I).���id;
        End If;
      End Loop;
      If v_��ϴ� Is Not Null Then
        Rs���ҽ��.Extend;
        Rs���ҽ��(Rs���ҽ��.Count).���ids := Substr(v_��ϴ�, 2);
        Rs���ҽ��(Rs���ҽ��.Count).ҽ��id := Rsap(Pd).Id;
      End If;
    End If;
  End;

  Procedure Get��ʵ�����id(Pd Number) As
    --���� rsdiag ����������ݣ�pd ���Ϊ��¼���±�����ֵ
    --��������ڣ���ֱ�ӱ���һ�������
    n_��¼id Number(18);
  
  Begin
    If Rsdiag(Pd).������� = 1 And Rsdiag(Pd).Icd10_Code Is Not Null Then
      Select Max(a.Id)
      Into n_��¼id
      From ������ϼ�¼ A, ��������Ŀ¼ B
      Where a.����id = b.Id And a.����id = r_Base.����id And a.��ҳid = n_����id And a.������� = 1 And b.���� = Rsdiag(Pd).Icd10_Code And
            Nvl(a.¼�����, '01') = '01' And b.��� = 'D';
    Elsif Rsdiag(Pd).������� = 11 And Rsdiag(Pd).Csd_Code Is Not Null Then
      Select Max(a.Id)
      Into n_��¼id
      From ������ϼ�¼ A, ��������Ŀ¼ B
      Where a.����id = b.Id And a.����id = r_Base.����id And a.��ҳid = n_����id And a.������� = 11 And b.���� = Rsdiag(Pd).Csd_Code And
            Nvl(a.¼�����, '01') = '01' And b.��� = 'B';
    Else
      --���������¼�����ж��ı�������
      Select Max(a.Id)
      Into n_��¼id
      From ������ϼ�¼ A
      Where a.����id = r_Base.����id And a.��ҳid = n_����id And Instr(a.�������, Rsdiag(Pd).Dz_Note) > 0 And
            Nvl(a.¼�����, '01') = '01';
    End If;
  
    If n_��¼id Is Not Null Then
      Rsdiag(Pd).���id := n_��¼id;
      Rsdiag(Pd).�Ƿ���� := 1;
    Else
      --�������б���
      Select ������ϼ�¼_Id.Nextval Into Rsdiag(Pd).���id From Dual;
      Rsdiag(Pd).�Ƿ���� := 0;
    End If;
  End;

  Procedure Get������� As
    --���ܣ���ȡ��Ҫ�²������ϵ�׼�����ݣ����浽 Rsdiagnew ��¼����
    n_��ҽ���� Number(3);
    n_��ҽ���� Number(3);
  Begin
    For I In 1 .. Rsdiag.Count Loop
      If Nvl(Rsdiag(I).�Ƿ����, 0) = 0 Then
        If n_��ҽ���� Is Null Then
          n_��ҽ���� := 0;
          n_��ҽ���� := 0;
          For R In (Select a.�������, Nvl(Max(a.��ϴ���), 0) ��ϴ���
                    From ������ϼ�¼ A
                    Where a.����id = r_Base.����id And a.��ҳid = n_����id And a.��¼��Դ = 3 And a.������� In (1, 11) And
                          Nvl(a.¼�����, '01') = '01'
                    Group By a.�������) Loop
          
            If r.������� = 11 Then
              n_��ҽ���� := r.��ϴ���;
            Else
              n_��ҽ���� := r.��ϴ���;
            End If;
          End Loop;
        End If;
      
        If Rsdiag(I).������� = 1 And Rsdiag(I).Icd10_Code Is Not Null Then
          Select Max(b.Id)
          Into Rsdiag(I).����id
          From ��������Ŀ¼ B
          Where b.���� = Rsdiag(I).Icd10_Code And b.��� = 'D';
          n_��ҽ���� := n_��ҽ���� + 1;
          Rsdiag(I).��ϴ��� := n_��ҽ����;
          Rsdiag(I).������� := '(' || Rsdiag(I).Icd10_Code || ')' || Rsdiag(I).Dz_Note;
        Elsif Rsdiag(I).������� = 11 And Rsdiag(I).Csd_Code Is Not Null Then
          Select Max(b.Id)
          Into Rsdiag(I).����id
          From ��������Ŀ¼ B
          Where b.���� = Rsdiag(I).Csd_Code And b.��� = 'B';
          n_��ҽ���� := n_��ҽ���� + 1;
          Rsdiag(I).��ϴ��� := n_��ҽ����;
          Rsdiag(I).������� := '(' || Rsdiag(I).Csd_Code || ')' || Rsdiag(I).Dz_Note;
          If Rsdiag(I).Syndrome Is Not Null Then
            Rsdiag(I).������� := Rsdiag(I).������� || '(' || Rsdiag(I).Syndrome || ')';
          End If;
        Else
          --����¼�����
          If Rsdiag(I).������� = 1 Then
            n_��ҽ���� := n_��ҽ���� + 1;
            Rsdiag(I).��ϴ��� := n_��ҽ����;
          Else
            n_��ҽ���� := n_��ҽ���� + 1;
            Rsdiag(I).��ϴ��� := n_��ҽ����;
          End If;
          Rsdiag(I).������� := Rsdiag(I).Dz_Note;
        End If;
      
        -- Zl_������ϼ�¼_Insert(n_����id, n_����id, 3, Null, r_Dz(R).�������, r_Dz(R).����id, Null, Null, r_Dz(R).�������, Null, Null, 0,
        --r_Dz(R).��¼����, r_Dz(R).ҽ��ids, r_Dz(R).��ϴ���);
        Rsdiagnew.Extend;
        Rsdiagnew(Rsdiagnew.Count) := Rsdiag(I);
      End If;
    End Loop;
  End;

  Procedure Getmore��������(Pd Number) As
    --pd ��ǰ��ҩ�е����±�
    --�������id���ռ�������
    v_������s Varchar2(4000);
    v_������  Varchar2(4000);
  Begin
    For I In 1 .. Pd Loop
      If Rsap(Pd - I).Serial_Num = Rsap(Pd).Serial_Num Then
        Rsap(Pd - I).���id := Rsap(Pd).Id;
        If Rsap(Pd - I).Diag_Nums Is Not Null Then
          --������������ظ��ģ�Ҫȥ�ظ�
          v_������s := v_������s || ',' || Rsap(Pd - I).Diag_Nums;
        End If;
      Else
        Rsap(Pd - I + 1).Firstrow := 1;
        Exit;
      End If;
    End Loop;
  
    If v_������s Is Not Null Then
      Select f_List2str(Cast(Collect(a.������ || '') As t_Strlist), ',') ������
      Into v_������
      From (Select a.������
             From (Select /*+cardinality(b,10) */
                     b.Column_Value ������
                    From Table(f_Str2list(v_������s)) B) A
             Where a.������ Is Not Null
             Group By a.������) A;
      Rsap(Pd).Diag_Nums := v_������;
    End If;
  
    Rsap(Pd).Apply_Id := r_Base.Apply_Id;
    Rsap(Pd).Apply_Type := r_Base.Apply_Type;
    Rsap(Pd).ҽ������ := r_Base.ҽ������;
    Rsap(Pd).������Ŀid := r_Base.������Ŀid; -- N
    Rsap(Pd).ִ�п���id := r_Base.ִ�п���id; --  N
    Rsap(Pd).�ܸ����� := r_Base.�ܸ�����; -- N  
    Rsap(Pd).������־ := Rsap(Pd - 1).������־;
    Rsap(Pd).ִ��Ƶ�� := Rsap(Pd - 1).ִ��Ƶ��;
    Rsap(Pd).Ƶ�ʴ��� := Rsap(Pd - 1).Ƶ�ʴ���;
    Rsap(Pd).Ƶ�ʼ�� := Rsap(Pd - 1).Ƶ�ʼ��;
    Rsap(Pd).�����λ := Rsap(Pd - 1).�����λ;
    Rsap(Pd).ִ��ʱ�䷽�� := Rsap(Pd - 1).ִ��ʱ�䷽��;
    Rsap(Pd).���� := Rsap(Pd - 1).����;
  
  End;

  Function Getҽ������(Pd Number) Return Varchar2 As
    --��������Ŀ��֯ҽ������
    --Pd ��ҽ����Ӧ���������±�
    v_����    Varchar2(4000);
    v_��λ    Varchar2(4000);
    v_����    Varchar2(4000);
    v_��λ_ǰ Varchar2(4000);
    --n_����id  Number(18);
    --v_����    Varchar2(1000);
  Begin
    If Rsap(Pd).Apply_Type = 3 Then
      --����  
      v_���� := '(' || Rsap(Pd).�걾��λ || ')';
      For I In 1 .. Pd Loop
        If Rsap(Pd).Id = Rsap(Pd - I).���id Then
          If I = 1 Then
            v_���� := Rsap(Pd - I).ҽ������ || v_����;
          Else
            v_���� := Rsap(Pd - I).ҽ������ || ',' || v_����;
          End If;
        Else
          Exit;
        End If;
      End Loop;
      --Rsap(Pd).��������id := n_����id;
      --Rsap(Pd).���� := v_����;
    Elsif Rsap(Pd).Apply_Type = 5 Then
      --���      
      For I In Pd + 1 .. Rsap.Count Loop
        If Rsap(Pd).Id = Rsap(I).���id Then
          If v_��λ_ǰ <> Rsap(I).�걾��λ And v_��λ_ǰ Is Not Null Then
            v_��λ := v_��λ || ',' || v_��λ_ǰ || '(' || Substr(v_����, 2) || ')';
            v_���� := Null;
          End If;
          v_��λ_ǰ := Rsap(I).�걾��λ;
          v_����    := v_���� || ',' || Rsap(I).��鷽��;
        Else
          Exit;
        End If;
      End Loop;
      If v_��λ_ǰ Is Not Null Then
        v_��λ := v_��λ || ',' || v_��λ_ǰ || '(' || Substr(v_����, 2) || ')';
      End If;
      v_���� := Rsap(Pd).ҽ������ || ':' || Substr(v_��λ, 2);
    End If;
    Return v_����;
  End;

  Procedure Get�������� As
    --���� Rs���� ��Ϣ�������ɵ� ��������  
  
    Rs����tmp t_Price := t_Price();
    n_����    Number(1);
  
    Function ������������
    (
      P�����ʶ Varchar2,
      P��Ŀid   Number
    ) Return Varchar2 As
      n_Keyҽ��id Number(18);
      v_��������  Varchar2(1000);
    Begin
      For I In 1 .. Rsap.Count Loop
        If P�����ʶ = Rsap(I).Apply_Id Then
          n_Keyҽ��id := Nvl(Rsap(I).���id, Rsap(I).Id);
          Exit;
        End If;
      End Loop;
      v_�������� := Zl_Cis_Nextno('125', n_Keyҽ��id, P��Ŀid);
      Return v_��������;
    End;
  Begin
    If n_�Ƿ��������� = 1 Then
      For I In 1 .. Rs����.Count Loop
        Select Max(a.�Թܱ���) Into Rs����(I).���� From ������ĿĿ¼ A Where a.Id = Rs����(I).������Ŀid;
        If Rs����(I).���� Is Not Null Then
          For R In (Select ����id From ��Ѫ������ Where ����id Is Not Null And ���� = Rs����(I).����) Loop
            Rs����(I).��������id := r.����id;
            Exit;
          End Loop;
        End If;
        Rs����(I).������־ := Nvl(Rs����(I).������־, 0);
        Rs����(I).Ӥ�� := Nvl(Rs����(I).Ӥ��, 0);
      End Loop;
    
      --���ڼ�����Ŀһ������Ԫ�ؾͱ�ʾһ��ҽ��
      For I In 1 .. Rs����.Count Loop
        n_���� := 0;
        --�ȴ������ɵ������в��ң�δ�ҵ���������ȡ
        For J In 1 .. Rs����tmp.Count Loop
        
          If Rs����(I)
           .������Ŀid <> Rs����tmp(J).������Ŀid And Rs����(I).�������id = Rs����tmp(J).�������id And Rs����(I).�ɼ���Ŀid = Rs����tmp(J).�ɼ���Ŀid And Rs����(I)
             .�ɼ�����id = Rs����tmp(J).�ɼ�����id And Rs����(I).�ɼ��걾 = Rs����tmp(J).�ɼ��걾 And Rs����(I).������־ = Rs����tmp(J).������־ And Rs����(I)
             .Ӥ�� = Rs����tmp(J).Ӥ�� And Rs����(I).���� = Rs����tmp(J).���� Then
          
            n_���� := 1;
            Rs����(I).�������� := Rs����tmp(J).��������;
            Exit;
          End If;
        End Loop;
      
        If n_���� = 0 Then
          Rs����(I).�������� := ������������(Rs����(I).Apply_Id, Rs����(I).������Ŀid); ----��ʱ�Ǽ����룬��������ҽ��id����Ŀid,������������ʵ������
          Rs����tmp.Extend;
          Rs����tmp(Rs����tmp.Count) := Rs����(I);
        End If;
      End Loop;
    End If;
  End;

Begin
  --�����۵���������Ϣ����
  If '�����۵�' = '�����۵�' Then
    --�������
    j_Tmp             := Pljson(Json_In);
    v_Input           := j_Tmp.Get_Pljson('input');
    r_Base.����id     := v_Input.Get_Number('pati_id');
    r_Base.�Һŵ�     := v_Input.Get_String('visit_no');
    r_Base.������Դ   := v_Input.Get_Number('pati_source');
    r_Base.����       := v_Input.Get_String('pati_name');
    r_Base.�Ա�       := v_Input.Get_String('pati_sex');
    r_Base.����       := v_Input.Get_String('pati_age');
    r_Base.���˿���id := v_Input.Get_Number('pati_deptid');
    r_Base.��������id := v_Input.Get_Number('apply_dept_id');
    r_Base.����ҽ��   := v_Input.Get_String('apply_doctor');
    r_Base.����ʱ��   := To_Date(v_Input.Get_String('apply_time'), 'yyyy-MM-dd HH24:MI:SS');
    v_Nodeno          := v_Input.Get_String('nodeno');
  
    j_Advicelist := v_Input.Get_Pljson_List('apply_list');
    Jl_Diag      := v_Input.Get_Pljson_List('diag_info');
  
    If r_Base.������Դ = 4 Then
      n_�Ƿ��������� := Nvl(v_Input.Get_Number('lis_bar_code_tag'), 0);
    Else
      n_�Ƿ��������� := Nvl(zl_GetSysParameter(143), 0);
    End If;
    d_����ʱ�� := r_Base.����ʱ�� + 1 / 24 / 60 / 60;
  
    j_Tmp := Pljson();
  
    If Jl_Diag Is Not Null Then
      Rsdiag.Extend(Jl_Diag.Count);
      Select a.Id Into n_����id From ���˹Һż�¼ A Where a.No = r_Base.�Һŵ� And a.��¼״̬ = 1 And a.��¼���� = 1;
      For I In 1 .. Jl_Diag.Count Loop
        j_Tmp := Pljson(Jl_Diag.Get(I));
      
        Rsdiag(I).Diag_Num := j_Tmp.Get_String('diag_num'); --  C
        Rsdiag(I).������� := j_Tmp.Get_Number('diag_type'); -- N
        Rsdiag(I).Csd_Code := j_Tmp.Get_String('csd_code'); --  C
        Rsdiag(I).Icd10_Code := j_Tmp.Get_String('icd10_code'); --  C
        Rsdiag(I).Dz_Note := j_Tmp.Get_String('dz_note'); --  C
        Rsdiag(I).Syndrome := j_Tmp.Get_String('syndrome'); --  C
        Rsdiag(I).Disease_Time := j_Tmp.Get_String('disease_time'); --  C
        Rsdiag(I).Diagnostician := j_Tmp.Get_String('diagnostician'); --  C
      
        j_Tmp := Pljson();
      End Loop;
    End If;
  End If;

  For I In 1 .. j_Advicelist.Count Loop
    j_Adviceitem      := Pljson(j_Advicelist.Get(I));
    r_Base.Apply_Id   := j_Adviceitem.Get_String('apply_id');
    r_Base.Diag_Nums  := j_Adviceitem.Get_String('diag_nums');
    r_Base.Apply_Type := j_Adviceitem.Get_Number('apply_type');
    r_Base.������־   := j_Adviceitem.Get_Number('emergency_tag');
  
    --����
    If r_Base.Apply_Type = 4 Then
      Rsap.Extend;
      Idx := Rsap.Count;
      n_����id := n_����id + 1;
      Rsap(Idx).Id := n_����id;
      Rsap(Idx).Apply_Id := r_Base.Apply_Id;
      Rsap(Idx).Apply_Type := r_Base.Apply_Type;
      Rsap(Idx).Diag_Nums := r_Base.Diag_Nums;
      Rsap(Idx).������־ := r_Base.������־;
    
      j_Tmp := j_Adviceitem.Get_Pljson('cure_info');
      Rsap(Idx).ִ��Ƶ�� := j_Tmp.Get_String('frequency_name');
      Rsap(Idx).Ƶ�ʴ��� := j_Tmp.Get_Number('frequency_times');
      Rsap(Idx).Ƶ�ʼ�� := j_Tmp.Get_Number('frequency_interval');
      Rsap(Idx).�����λ := j_Tmp.Get_String('interval_unit');
      Rsap(Idx).ִ��ʱ�䷽�� := j_Tmp.Get_String('exetime_plane');
      Rsap(Idx).������Ŀid := j_Tmp.Get_Number('cure_item_id');
      Rsap(Idx).ִ�п���id := j_Tmp.Get_Number('cure_exedept_id');
      Rsap(Idx).ҽ������ := j_Tmp.Get_String('cure_doctor_note');
      Rsap(Idx).�������� := j_Tmp.Get_Number('cure_once_qunt');
      Rsap(Idx).�ܸ����� := j_Tmp.Get_Number('cure_total_qunt');
      Rsap(Idx).Firstrow := 1;
    
      j_Tmp := Pljson();
      --����
    Elsif r_Base.Apply_Type = 3 Then
    
      j_Tmp             := j_Adviceitem.Get_Pljson('lis_info');
      v_Lisitems        := j_Tmp.Get_String('lis_items');
      n_����id          := n_����id + 1;
      r_Base.Id         := n_����id;
      r_Base.ִ�п���id := j_Tmp.Get_Number('lis_exedept_id');
      r_Base.�걾��λ   := j_Tmp.Get_String('lis_spcm');
    
      --�������������׼��
      Rs����.Extend;
      Rs����(Rs����.Count).Apply_Id := r_Base.Apply_Id;
      Rs����(Rs����.Count).�������id := r_Base.ִ�п���id;
      Rs����(Rs����.Count).�ɼ��걾 := r_Base.�걾��λ;
      Rs����(Rs����.Count).Ӥ�� := 0;
      Rs����(Rs����.Count).������־ := r_Base.������־;
    
      For r_��Ŀ In (Select /*+cardinality(b,10)*/
                    b.Column_Value ��Ŀid
                   From Table(f_Str2list(v_Lisitems)) B) Loop
      
        Rsap.Extend;
        Idx := Rsap.Count;
        n_����id := n_����id + 1;
        Rsap(Idx).Id := n_����id;
        Rsap(Idx).���id := r_Base.Id;
        Rsap(Idx).������Ŀid := r_��Ŀ.��Ŀid;
        Rsap(Idx).ִ�п���id := r_Base.ִ�п���id;
        Rsap(Idx).�걾��λ := r_Base.�걾��λ;
        Rsap(Idx).Apply_Id := r_Base.Apply_Id;
        Rsap(Idx).Apply_Type := r_Base.Apply_Type;
        Rsap(Idx).Diag_Nums := r_Base.Diag_Nums;
        Rsap(Idx).������־ := r_Base.������־;
        Rsap(Idx).ִ��Ƶ�� := 'һ����';
      
        --һ���ɼ������еļ�����Ŀ
        If Rs����(Rs����.Count).������Ŀid Is Null Then
          Rs����(Rs����.Count).������Ŀid := r_��Ŀ.��Ŀid;
        
          Rsap(Idx).Firstrow := 1;
        End If;
      End Loop;
    
      Rsap.Extend;
      Idx := Rsap.Count;
      Rsap(Idx).Id := r_Base.Id;
      Rsap(Idx).������Ŀid := j_Tmp.Get_Number('lis_collect_item_id');
      Rsap(Idx).ִ�п���id := j_Tmp.Get_Number('lis_collect_exedept_id');
      Rsap(Idx).ҽ������ := j_Tmp.Get_String('lis_doctor_note');
      Rsap(Idx).�걾��λ := r_Base.�걾��λ;
      Rsap(Idx).Apply_Id := r_Base.Apply_Id;
      Rsap(Idx).Apply_Type := r_Base.Apply_Type;
      Rsap(Idx).Diag_Nums := r_Base.Diag_Nums;
      Rsap(Idx).������־ := r_Base.������־;
      Rsap(Idx).ִ��Ƶ�� := 'һ����';
    
      Rs����(Rs����.Count).�ɼ���Ŀid := Rsap(Idx).������Ŀid;
      Rs����(Rs����.Count).�ɼ�����id := Rsap(Idx).ִ�п���id;
    
      j_Tmp := Pljson();
      --���
    Elsif r_Base.Apply_Type = 5 Then
    
      j_Tmp := j_Adviceitem.Get_Pljson('pacs_info');
      Rsap.Extend;
      Idx := Rsap.Count;
      n_����id := n_����id + 1;
      Rsap(Idx).Id := n_����id;
      r_Base.Id := n_����id;
      Rsap(Idx).������Ŀid := j_Tmp.Get_Number('pacs_item_id');
      Rsap(Idx).ִ�п���id := j_Tmp.Get_Number('pacs_exedept_id');
      Rsap(Idx).ҽ������ := j_Tmp.Get_String('pacs_doctor_note');
    
      Rsap(Idx).Apply_Id := r_Base.Apply_Id;
      Rsap(Idx).Apply_Type := r_Base.Apply_Type;
      Rsap(Idx).Diag_Nums := r_Base.Diag_Nums;
      Rsap(Idx).������־ := r_Base.������־;
      Rsap(Idx).ִ��Ƶ�� := 'һ����';
      Rsap(Idx).Firstrow := 1;
      --��λ��Ϣ
    
      Jl_Tmp := j_Tmp.Get_Pljson_List('pacs_part_list');
    
      If Jl_Tmp Is Not Null Then
        For J In 1 .. Jl_Tmp.Count Loop
        
          j_Tmp := Pljson(Jl_Tmp.Get(J));
          Rsap.Extend;
          Idx := Rsap.Count;
          n_����id := n_����id + 1;
          Rsap(Idx).Id := n_����id;
          Rsap(Idx).���id := r_Base.Id;
          Rsap(Idx).������Ŀid := Rsap(Idx - 1).������Ŀid;
          Rsap(Idx).ִ�п���id := Rsap(Idx - 1).ִ�п���id;
          Rsap(Idx).�걾��λ := j_Tmp.Get_String('part_name');
          Rsap(Idx).��鷽�� := j_Tmp.Get_String('part_way');
        
          Rsap(Idx).Apply_Id := r_Base.Apply_Id;
          Rsap(Idx).Apply_Type := r_Base.Apply_Type;
          Rsap(Idx).Diag_Nums := r_Base.Diag_Nums;
          Rsap(Idx).������־ := r_Base.������־;
          Rsap(Idx).ִ��Ƶ�� := 'һ����';
        
          j_Tmp := Pljson();
        End Loop;
      End If;
    
      j_Tmp  := Pljson();
      Jl_Tmp := Pljson_List();
      --����
    Elsif r_Base.Apply_Type = 1 Then
      n_Pre���   := Null;
      Jl_Tmp      := j_Adviceitem.Get_Pljson_List('drug_info');
      n_Drug_Rows := Jl_Tmp.Count;
      For J In 1 .. n_Drug_Rows Loop
        j_Tmp  := Pljson(Jl_Tmp.Get(J));
        n_��� := j_Tmp.Get_Number('serial_num');
        If n_��� <> n_Pre��� And n_Pre��� Is Not Null Then
          --׷��һ�и�ҩ;���У�ͬʱ��Ҫ����ǰ���е�ҩƷ�е����id���ռ�������          
          Rsap.Extend;
          Idx := Rsap.Count;
          n_����id := n_����id + 1;
          Rsap(Idx).Id := n_����id;
          Rsap(Idx).Serial_Num := n_Pre���;
          Getmore��������(Idx);
        End If;
        n_Pre��� := n_���;
      
        Rsap.Extend;
        Idx := Rsap.Count;
        n_����id := n_����id + 1;
        Rsap(Idx).Id := n_����id;
        Rsap(Idx).Apply_Id := r_Base.Apply_Id;
        Rsap(Idx).Apply_Type := r_Base.Apply_Type;
        Rsap(Idx).Serial_Num := n_���;
        Rsap(Idx).Diag_Nums := j_Tmp.Get_String('diag_nums');
        Rsap(Idx).������־ := j_Tmp.Get_Number('emergency_tag');
        Rsap(Idx).ִ��Ƶ�� := j_Tmp.Get_String('frequency_name');
        Rsap(Idx).Ƶ�ʴ��� := j_Tmp.Get_Number('frequency_times');
        Rsap(Idx).Ƶ�ʼ�� := j_Tmp.Get_Number('frequency_interval');
        Rsap(Idx).�����λ := j_Tmp.Get_String('interval_unit');
        Rsap(Idx).ִ��ʱ�䷽�� := j_Tmp.Get_String('exetime_plane');
        Rsap(Idx).���� := j_Tmp.Get_Number('user_day');
        Rsap(Idx).�շ�ϸĿid := j_Tmp.Get_Number('drug_id'); -- N
        Rsap(Idx).ִ�п���id := j_Tmp.Get_Number('pharmacy_id'); -- N
        Rsap(Idx).�������� := j_Tmp.Get_Number('drug_once_qunt'); --  N
        Rsap(Idx).�ܸ����� := j_Tmp.Get_Number('drug_total_qunt'); -- N
        Rsap(Idx).ҽ������ := j_Tmp.Get_String('doctor_note'); -- C
        Rsap(Idx).��ҩĿ�� := j_Tmp.Get_Number('drug_purpose'); --  N
        Rsap(Idx).��ҩ���� := j_Tmp.Get_String('drug_reason'); -- C
        Rsap(Idx).����˵�� := j_Tmp.Get_String('excs_desc'); -- C      
      
        r_Base.ҽ������   := j_Tmp.Get_String('dripping_speed');
        r_Base.������Ŀid := j_Tmp.Get_Number('use_item_id'); -- N
        r_Base.ִ�п���id := j_Tmp.Get_Number('use_exedept_id'); --  N
        r_Base.�ܸ�����   := j_Tmp.Get_Number('use_count'); -- N  
      
        j_Tmp := Pljson();
      End Loop;
      If n_Pre��� Is Not Null Then
        --׷��һ�и�ҩ;���У�ͬʱ��Ҫ����ǰ���е�ҩƷ�е����id���ռ�������          
        Rsap.Extend;
        Idx := Rsap.Count;
        n_����id := n_����id + 1;
        Rsap(Idx).Id := n_����id;
        Rsap(Idx).Serial_Num := n_Pre���;
        Getmore��������(Idx);
      End If;
      Jl_Tmp := Pljson_List();
    End If;
    j_Adviceitem := Pljson();
  End Loop;

  --���±����б���������Ϣ
  Select Nvl(Max(���), 0) Into n_��� From ����ҽ����¼ Where �Һŵ� = r_Base.�Һŵ�;
  For I In 1 .. Rsap.Count Loop
    If Nvl(Rsap(I).�շ�ϸĿid, 0) <> 0 Then
      --ҩƷ���������䷽��
      Select a.����, a.Id, a.���, a.ִ�п���, c.����, c.���, b.����ϵ��
      Into Rsap(I).�걾��λ,Rsap(I).������Ŀid,Rsap(I).�������,Rsap(I).ִ������, v_����, v_���, n_����ϵ��
      From ������ĿĿ¼ A, ҩƷ��� B, �շ���ĿĿ¼ C
      Where a.Id = b.ҩ��id And b.ҩƷid = c.Id And c.Id = Rsap(I).�շ�ϸĿid;
      Rsap(I).ҽ������ := Rsap(I).�걾��λ;
      If v_���� Is Not Null Then
        Rsap(I).ҽ������ := Rsap(I).ҽ������ || '(' || v_���� || ')';
      End If;
      If v_��� Is Not Null Then
        Rsap(I).ҽ������ := Rsap(I).ҽ������ || ' ' || v_���;
      End If;
      Rsap(I).ִ������ := 4; --ҩƷͨ��Ϊָ������ִ�У��������������Ժ��ҩ���ᷢ����
      Rsap(I).�������� := n_����ϵ�� * Rsap(I).�ܸ�����;
    Else
      Select a.����, a.���, a.ִ�п���, a.�Թܱ���
      Into Rsap(I).ҽ������,Rsap(I).�������,Rsap(I).ִ������,Rsap(I).����
      From ������ĿĿ¼ A
      Where a.Id = Rsap(I).������Ŀid;
      Rsap(I).�������� := Nvl(Rsap(I).�ܸ�����, 1);
    End If;
  
    n_��� := n_��� + 1;
    Rsap(I).��� := n_���;
    Rsap(I).����id := r_Base.����id;
    Rsap(I).�Һŵ� := r_Base.�Һŵ�;
    Rsap(I).������Դ := r_Base.������Դ;
    Rsap(I).���� := r_Base.����;
    Rsap(I).�Ա� := r_Base.�Ա�;
    Rsap(I).���� := r_Base.����;
    Rsap(I).���˿���id := r_Base.���˿���id;
    Rsap(I).��������id := r_Base.��������id;
    Rsap(I).����ҽ�� := r_Base.����ҽ��;
    Rsap(I).����ʱ�� := r_Base.����ʱ��;
    Rsap(I).��ʼִ��ʱ�� := r_Base.����ʱ��;
    --Rsap(I).Ӥ�� := 0; Ϊnull ����ZLHIS����̨�п�������0Ŀǰ���ô������֣�Rsap(I).ִ�б�� Ҳ����Ϊnull
    Rsap(I).ҽ��״̬ := 1;
    Rsap(I).ҽ����Ч := 1;
    Rsap(I).�Ƽ����� := 0;
  
  End Loop;

  -------����Ϊ����������ȡ����---------
  -------���ݼӹ�-----------------------
  --����ҽ������
  l_ҽ��id.Extend(n_����id);
  --��飬������Ŀ����ҽ�����ݣ�����ҽ������
  For I In 1 .. Rsap.Count Loop
    If Rsap(I).Apply_Type = 3 And Rsap(I).���id Is Null Then
      Rsap(I).ҽ������ := Getҽ������(I);
    Elsif Rsap(I).Apply_Type = 5 And Rsap(I).���id Is Null Then
      Rsap(I).ҽ������ := Getҽ������(I);
    End If;
    --��������
    n_ҽ��id := Rsap(I).Id;
    Select ����ҽ����¼_Id.Nextval Into l_ҽ��id(n_ҽ��id) From Dual;
  End Loop;

  --�滻Ϊ��ʵ��ҽ��id
  For I In 1 .. Rsap.Count Loop
    n_ҽ��id := Rsap(I).Id;
    Rsap(I).Id := l_ҽ��id(n_ҽ��id);
    If Rsap(I).���id Is Not Null Then
      n_ҽ��id := Rsap(I).���id;
      Rsap(I).���id := l_ҽ��id(n_ҽ��id);
    End If;
  End Loop;

  --���ɼ�����Ŀ�����룬������ rs���� �����У��ڲ����õ���ʵ��ҽ��id�����������˳��
  Get��������;

  --������ݴ���
  For I In 1 .. Rsdiag.Count Loop
    Get��ʵ�����id(I);
  End Loop;

  --ҽ���������ݲ������ZLHIS�Ĺ���  
  --�������ݺţ����ͺţ���¼���
  Select Zl_Cis_Nextno('10') Into n_���ͺ� From Dual;
  For I In 1 .. Rsap.Count Loop
    If I = 1 Then
      Select Zl_Cis_Nextno('13') Into v_No From Dual;
      Rsap(I).No := v_No;
      n_��� := 1;
      Rsap(I).��¼��� := n_���;
    Else
      If Rsap(I).Apply_Id = Rsap(I - 1).Apply_Id Then
        Rsap(I).No := Rsap(I - 1).No;
        n_��� := n_��� + 1;
        Rsap(I).��¼��� := n_���;
      Else
        Select Zl_Cis_Nextno('13') Into v_No From Dual;
        Rsap(I).No := v_No;
        n_��� := 1;
        Rsap(I).��¼��� := n_���;
      End If;
    End If;
    Rsap(I).���ͺ� := n_���ͺ�;
  
    --���û��ִ��ʱ�䷽��������ִ��һ��
    If Rsap(I).ִ��ʱ�䷽�� Is Null Then
      Rsap(I).�״�ʱ�� := Rsap(I).��ʼִ��ʱ��;
      Rsap(I).ĩ��ʱ�� := Rsap(I).�״�ʱ��;
    End If;
    --�������봦��
    If Rsap(I).Apply_Type = 3 Then
      If Rsap(I).Firstrow = 1 Then
        For R In 1 .. Rs����.Count Loop
          If Rs����(R).Apply_Id = Rsap(I).Apply_Id And Rs����(R).�������� Is Not Null Then
            Rsap(I).�������� := Rs����(R).��������;
            Exit;
          End If;
        End Loop;
      Else
        Rsap(I).�������� := Rsap(I - 1).��������;
      End If;
    End If;
  
    --��ȡ��϶�Ӧ
    Getҽ����϶�Ӧ(I);
  End Loop;

  --��ȡ��Ҫ�²������ϵ�׼�����ݣ����浽 Rsdiagnew ��¼����
  Get�������;
  For I In 1 .. Rsdiagnew.Count Loop
    Zl_������ϼ�¼_Insert(r_Base.����id, n_����id, 3, Null, Rsdiagnew(I).�������, Rsdiagnew(I).����id, Null, Null, Rsdiagnew(I).�������,
                     Null, Null, 0, Sysdate, Null, Rsdiagnew(I).��ϴ���, Null, Null,
                     To_Date(Rsdiagnew(I).Disease_Time, 'yyyy-mm-dd hh24:mi:ss'), Rsdiagnew(I).Diagnostician,
                     Rsdiagnew(I).���id);
  End Loop;

  For I In 1 .. Rsap.Count Loop
    Zl_����ҽ����¼_Insert_s(Rsap(I).Id, Rsap(I).���id, Rsap(I).���, Rsap(I).������Դ, Rsap(I).����id, Null, 0, 1, 1, Rsap(I).�������,
                       Rsap(I).������Ŀid, Rsap(I).�շ�ϸĿid, Rsap(I).����, Rsap(I).��������, Rsap(I).�ܸ�����, Rsap(I).ҽ������,
                       Rsap(I).ҽ������, Rsap(I).�걾��λ, Rsap(I).ִ��Ƶ��, Rsap(I).Ƶ�ʴ���, Rsap(I).Ƶ�ʼ��, Rsap(I).�����λ,
                       Rsap(I).ִ��ʱ�䷽��, Rsap(I).�Ƽ�����, Rsap(I).ִ�п���id, Rsap(I).ִ������, Rsap(I).������־, Rsap(I).��ʼִ��ʱ��, Null,
                       Rsap(I).���˿���id, Rsap(I).��������id, Rsap(I).����ҽ��, Rsap(I).����ʱ��, Rsap(I).�Һŵ�, Null, Rsap(I).��鷽��,
                       Rsap(I).ִ�б��, Null, Null, Rsap(I).����ҽ��, Null, Rsap(I).��ҩĿ��, Rsap(I).��ҩ����, Null, Null,
                       Rsap(I).����˵��, Null, Rsap(I).�䷽id, Null, Null, Null, Null, Null, Rsap(I).����, Rsap(I).�Ա�,
                       Rsap(I).����);
  End Loop;

  For I In 1 .. Rs���ҽ��.Count Loop
    Zl_�������ҽ��_Insert(Rs���ҽ��(I).ҽ��id, Rs���ҽ��(I).���ids);
  End Loop;

  For I In 1 .. Rsap.Count Loop
    Zl_����ҽ������_Insert_s(Rsap(I).Id, n_���ͺ�, 1, Rsap(I).No, Rsap(I).��¼���, Rsap(I).��������, Rsap(I).�״�ʱ��, Rsap(I).ĩ��ʱ��, d_����ʱ��,
                       Rsap(I).ִ�п���id, 0, Rsap(I).����ҽ��, Rsap(I).Firstrow, Rsap(I).��������, Null, 0, 0);
  End Loop;

  --��ȡ��������Ϣ
  For I In 1 .. Rsap.Count Loop
    If I = 1 Then
      v_Out := v_Out || ',{"apply_id":"' || Rsap(I).Apply_Id || '"';
      v_Out := v_Out || ',"fee_no":"' || Rsap(I).No || '"';
      If Rsap(I).Apply_Type = 3 Then
        v_Out := v_Out || ',"lis_bar_code":"' || Rsap(I).�������� || '"';
      End If;
    Else
      If Rsap(I).Apply_Id <> Rsap(I - 1).Apply_Id Then
        v_Out  := v_Out || ',"order_list":[' || Substr(v_Tmp1, 2) || ']';
        v_Out  := v_Out || '}';
        v_Tmp1 := Null;
        v_Out  := v_Out || ',{"apply_id":"' || Rsap(I).Apply_Id || '"';
        v_Out  := v_Out || ',"fee_no":"' || Rsap(I).No || '"';
        If Rsap(I).Apply_Type = 3 Then
          v_Out := v_Out || ',"lis_bar_code":"' || Rsap(I).�������� || '"';
        End If;
      End If;
    End If;
    v_Tmp1 := v_Tmp1 || ',{"order_id":' || Rsap(I).Id;
    v_Tmp1 := v_Tmp1 || ',"order_related_id":' || Nvl(Rsap(I).���id || '', 'null');
    v_Tmp1 := v_Tmp1 || ',"cisitem_id":' || Rsap(I).������Ŀid;
    v_Tmp1 := v_Tmp1 || '}';
  End Loop;
  If v_Out Is Not Null Then
    v_Out := v_Out || ',"order_list":[' || Substr(v_Tmp1, 2) || ']';
    v_Out := v_Out || '}';
  End If;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","advice_list":[' || Substr(v_Out, 2) || ']}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Outsendapply;
/

Create Or Replace Procedure Zl_Cissvr_Revokeoutadvice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ������ﲡ��ȡ�����루��ע����Һ�ദ���ȣ�����ZLHIS����ҽ��
  --��Σ�Json_In:��ʽ
  --  input
  --          operator_name     C 1 ����Ա����
  --          operator_code     C 1 ����Ա���
  --          order_ids         C 1 ��ҽ��id,��ҽ��id������ƴ������֧��һ�δ���������ȡ��

  --����: Json_Out,��ʽ����
  --  output
  --    code                      N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                   C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Json       Pljson;
  j_In         Pljson;
  v_ҽ��ids    Varchar2(4000);
  v_����Ա���� Varchar2(1000);
  v_����Ա��� Varchar2(1000);
  v_����ʱ��   Varchar2(300);
  n_Count      Number(2);
Begin
  --�������
  j_In         := Pljson(Json_In);
  j_Json       := j_In.Get_Pljson('input');
  v_ҽ��ids    := j_Json.Get_String('order_ids');
  v_����Ա���� := j_Json.Get_String('operator_name');
  v_����Ա��� := j_Json.Get_String('operator_code');

  Select To_Char(Sysdate, 'yyyy-mm-dd hh24:mi:ss') Into v_����ʱ�� From Dual;

  If Instr(v_ҽ��ids, ',') > 0 Then
  
    Select Count(1)
    Into n_Count
    From ����ҽ������ A
    Where a.ҽ��id In (Select /*+cardinality(j,10)*/
                      x.Id
                     From ����ҽ����¼ X, Table(f_Num2list(v_ҽ��ids)) J
                     Where x.Id = j.Column_Value Or x.���id = j.Column_Value) And a.ִ��״̬ In (1, 3);
  
    If n_Count > 0 Then
      Json_Out := Zljsonout('��ҽ����Ŀ��ִ�л�����ִ�в������ϣ�');
      Return;
    End If;
  
    For R In (Select /*+cardinality(j,10) */
               j.Column_Value ҽ��id
              From Table(Cast(f_Num2list(v_ҽ��ids) As Zltools.t_Numlist)) J) Loop
      Zl_����ҽ����¼_����_s(Null, r.ҽ��id, Null, Null, v_����Ա����, v_����Ա���, v_����ʱ��);
    End Loop;
  
  Else
  
    Select Count(1)
    Into n_Count
    From ����ҽ������ A
    Where a.ҽ��id In (Select x.Id From ����ҽ����¼ X Where x.Id = v_ҽ��ids Or x.���id = v_ҽ��ids) And a.ִ��״̬ In (1, 3);
    If n_Count > 0 Then
      Json_Out := Zljsonout('��ҽ����Ŀ��ִ�л�����ִ�в������ϣ�');
      Return;
    End If;
  
    Zl_����ҽ����¼_����_s(Null, v_ҽ��ids, Null, Null, v_����Ա����, v_����Ա���, v_����ʱ��);
  End If;

  Json_Out := Zljsonout('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Revokeoutadvice;
/