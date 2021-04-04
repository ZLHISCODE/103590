Create Or Replace Procedure Zl_Pivassvr_Checkorderroll
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ�ҽ�����˷���ʱ������ҽ������ж�
  --��Σ�Json_In:��ʽ
  --  input
  --     order_id           N 1 ҽ��ID,��ҽ��id
  --     send_no            N 1 ���ͺ�
  --     item_list[]
  --            order_id           N 1 ҽ��ID,��ҽ��id
  --            send_no            N 1 ���ͺ�

  --����: Json_Out,��ʽ����
  --  output
  --    code                N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    pivas_ids           C 1 Ҫ���ʵ���Һ��¼id��
  ---------------------------------------------------------------------------
  n_ҽ��id   Number;
  n_���ͺ�   Number;
  j_Json     Pljson;
  j_Tmp      Pljson;
  n_Tmp      Number;
  v_��Һids  Varchar2(32767);
  n_List_Cnt Number;
  j_Jsonlist Pljson_List;
Begin
  --�������
  j_Tmp      := Pljson(Json_In);
  j_Json     := j_Tmp.Get_Pljson('input');
  j_Jsonlist := j_Json.Get_Pljson_List('item_list');

  If Not j_Jsonlist Is Null Then
    n_List_Cnt := j_Jsonlist.Count;
  Else
    n_List_Cnt := 1;
  End If;

  For I In 1 .. n_List_Cnt Loop
  
    If Not j_Jsonlist Is Null Then
      j_Json := Pljson();
      j_Json := Pljson(j_Jsonlist.Get(I));
    End If;
  
    n_ҽ��id := j_Json.Get_Number('order_id');
    n_���ͺ� := j_Json.Get_Number('send_no');
  
    --����Ƿ�����Һ��Һ��¼�����Ƿ��Ѿ�����
    Select Decode(Max(�Ƿ�����), 1, 1, 0) Into n_Tmp From ��Һ��ҩ��¼ Where ҽ��id = n_ҽ��id And ���ͺ� = n_���ͺ�;
    If n_Tmp = 1 Then
      Json_Out := Zljsonout('��ǰ���������ҺҩƷҽ�����Ѿ�����Һ�����������������ܻ��˷��͡�');
      Return;
    Elsif n_Tmp = 0 Then
      --ֻ��״̬=1(δ��ҩ)�ļ�¼��������Ѿ���ҩ�ˣ���ͨ�����˷�ʽ����
      Select Count(ID)
      Into n_Tmp
      From ��Һ��ҩ��¼
      Where ����״̬ In (1, 10) And ҽ��id = n_ҽ��id And ���ͺ� = n_���ͺ�;
      If n_Tmp > 0 Then
        For R In (Select ID From ��Һ��ҩ��¼ Where ҽ��id = n_ҽ��id And ���ͺ� = n_���ͺ�) Loop
          v_��Һids := v_��Һids || ',' || r.Id;
        End Loop;
      End If;
    End If;
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","pivas_ids":"' || Substr(v_��Һids, 2) || '"}}';
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Pivassvr_Checkorderroll;
/
Create Or Replace Procedure Zl_Pivassvr_Infusion_Update
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���Һ��ҩ��¼����Һ��ҩ״̬����
  --��Σ� Json_In:��ʽ
  --  input
  --    dispensing_id       C 0 ��ҩID���������գ���advice_list[]�ش�
  --    type                N 1 ����0-��������;1-��������ȡ��(ɾ��);2-�����������
  --    operator_name       C 1 ����Ա����
  --    operator_notes      C 0 ����˵��
  --    operator_time       C 1 ����ʱ�䣺yyyy-mm-dd hh24:mi:ss
  --    apply_time          C 1 ����ʱ��(�����������ʱ��Ч)��yyyy-mm-dd hh24:mi:ss

  --    advice_list[]ҽ����Ϣ�б�:��������ҩID�Ļ�type<>0ʱ�����б���Ч  ��ʾ�������������һ��ȡ�������
  --        advice_id        N 1 ҽ��Id
  --        send_no          N 1 ���ͺ�
  --����: Json_Out,��ʽ����
  --  output
  --    code                 N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message              C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------

  j_Json       Pljson;
  j_Jsonlist   Pljson_List;
  o_Json       Pljson;
  n_��ҩid     Number(18);
  v_Tmp        Varchar2(4000);
  n_Count      Number(18);
  n_Temp       Number(18);
  n_����״̬   Number(2);
  v_����Ա���� ��Һ��ҩ��¼.������Ա%Type;
  v_����˵��   ��Һ��ҩ״̬.����˵��%Type;
  d_Date       Date;
  n_��������   ��Һ��ҩ��¼.����״̬%Type;
  n_ҽ��id     ��Һ��ҩ��¼.ҽ��id%Type;
  n_���ͺ�     ��Һ��ҩ��¼.���ͺ�%Type;
  d_Apply      Date;
  j_Json_Tmp   Pljson;
Begin
  --�������
  j_Json_Tmp := Pljson(Json_In);
  j_Json     := j_Json_Tmp.Get_Pljson('input');
  v_Tmp      := j_Json.Get_String('dispensing_id');
  If Nvl(v_Tmp, '-') <> '-' Then
    n_��ҩid := To_Number(v_Tmp);
  End If;
  v_Tmp := j_Json.Get_String('operator_time');
  If Nvl(v_Tmp, '-') <> '-' Then
    d_Date := To_Date(v_Tmp, 'yyyy-mm-dd hh24:mi:ss');
  Else
    d_Date := Sysdate;
  End If;
  n_����״̬ := j_Json.Get_Number('type');

  If Nvl(n_����״̬, 0) = 0 Then
  
    v_����Ա���� := j_Json.Get_String('operator_name');
    v_����˵��   := j_Json.Get_String('operator_notes');
  
    --��������:
    If n_��ҩid = 0 Then
      --��������ʱ����ҩID���봫�� 
      Json_Out := '{"output":{"code":0,"message":"��ҩIDδ����"}}';
      Return;
    End If;
    Select Count(1) Into n_Count From ��Һ��ҩ״̬ Where ��ҩid = n_��ҩid And �������� = 9 And ����ʱ�� = d_Date;
    If n_Count = 0 Then
      Insert Into ��Һ��ҩ״̬
        (��ҩid, ��������, ������Ա, ����ʱ��, ����˵��)
      Values
        (n_��ҩid, 9, v_����Ա����, d_Date, v_����˵��);
    End If;
    Update ��Һ��ҩ��¼ Set ������Ա = v_����Ա����, ����ʱ�� = d_Date, ����״̬ = 9 Where ID = n_��ҩid;
  End If;

  If Nvl(n_����״̬, 0) = 1 Then
    --��������ȡ����ɾ��
    If n_��ҩid Is Not Null Then
      Select ������Ա, ����ʱ��, ��������
      Into v_����Ա����, d_Date, n_��������
      From (Select ������Ա, ����ʱ��, ��������
             From ��Һ��ҩ״̬
             Where ��ҩid = n_��ҩid And �������� <> 9
             Order By ����ʱ�� Desc, �������� Desc)
      Where Rownum < 2;
      Update ��Һ��ҩ��¼ Set ������Ա = v_����Ա����, ����ʱ�� = d_Date, ����״̬ = n_�������� Where ID = n_��ҩid;
    End If;
    j_Jsonlist := Pljson_List();
    j_Jsonlist := j_Json.Get_Pljson_List('advice_list');
    n_Count    := j_Jsonlist.Count;
    If n_Count = 0 Then
      Json_Out := '{"output":{"code":0,"message":"δ����ҽ����Ϣ���䷽ID"}}';
      Return;
    End If;
    --��δ�ṩ����ҩ����ȡ���Ĺ��ܣ����������������һ��ȡ��
    For I In 1 .. n_Count Loop
      o_Json   := Pljson();
      o_Json   := Pljson(j_Jsonlist.Get(I));
      n_ҽ��id := o_Json.Get_Number('advice_id');
      n_���ͺ� := o_Json.Get_Number('send_no');    
      For R In (Select d.Id From ��Һ��ҩ��¼ D Where d.ҽ��id = Nvl(n_ҽ��id, 0) And d.���ͺ� = Nvl(n_���ͺ�, 0)) Loop
        Select ������Ա, ����ʱ��, ��������
        Into v_����Ա����, d_Date, n_��������
        From (Select ������Ա, ����ʱ��, ��������
               From ��Һ��ҩ״̬
               Where ��ҩid = r.Id And �������� <> 9
               Order By ����ʱ�� Desc, �������� Desc)
        Where Rownum < 2;      
        Update ��Һ��ҩ��¼ Set ������Ա = v_����Ա����, ����ʱ�� = d_Date, ����״̬ = n_�������� Where ID = r.Id;
      End Loop;
    End Loop;
  End If;
  If Nvl(n_����״̬, 0) = 2 Then
    --�������
    v_Tmp := j_Json.Get_String('apply_time');
    If Nvl(v_Tmp, '-') <> '-' Then
      d_Apply := To_Date(v_Tmp, 'yyyy-mm-dd hh24:mi:ss');
    Else
      d_Apply := Sysdate;
    End If;
    j_Jsonlist := Pljson_List();
    j_Jsonlist := j_Json.Get_Pljson_List('advice_list');
    n_Count    := j_Jsonlist.Count;
    If n_Count = 0 Then
      Json_Out := '{"output":{"code":0,"message":"δ����ҽ����Ϣ"}}';
      Return;
    End If;
  
    --��δ�ṩ����ҩ����ȡ���Ĺ��ܣ����������������һ��ȡ��
    For I In 1 .. n_Count Loop
      o_Json   := Pljson();
      o_Json   := Pljson(j_Jsonlist.Get(I));
      n_ҽ��id := o_Json.Get_Number('advice_id');
      n_���ͺ� := o_Json.Get_Number('send_no');
      Select Nvl(Max(d.Id), 0)
      Into n_��ҩid
      From ��Һ��ҩ��¼ D
      Where d.Id = n_ҽ��id And d.���ͺ� = n_���ͺ� And d.����ʱ�� = d_Apply And d.����״̬ = 9;
    
      If n_��ҩid <> 0 Then
        Select Count(1) Into n_Temp From ��Һ��ҩ״̬ Where ��ҩid = n_��ҩid And �������� = 10 And ����ʱ�� = d_Date;
      
        If n_Temp = 0 Then
          Insert Into ��Һ��ҩ״̬ (��ҩid, ��������, ������Ա, ����ʱ��) Values (n_��ҩid, 10, v_����Ա����, d_Date);
        End If;
        Update ��Һ��ҩ��¼ Set ������Ա = v_����Ա����, ����ʱ�� = d_Date, ����״̬ = 10 Where ID = n_��ҩid;
      End If;
    End Loop;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Pivassvr_Infusion_Update;
/

Create Or Replace Procedure Zl_Pivassvr_Isexsitinfusion
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ݷ���ID,�ж϶�Ӧ��ҩƷ�Ƿ��Ѿ�������Һ��ҩ����
  --��Σ�Json_In:��ʽ
  --input
  --        advice_id   N 1 ҽ��ID
  --        rcpdtl_ids  C 1 ������ϸIDs������ķ���IDs��,����ö��ŷ��룬��������˸ýڵ�����ϸIDsΪ׼��������ҽ��IDΪ׼��
  --        is_return   N 1 �Ƿ񷵻� ������Һ�еķ���id��
  --����: Json_Out,��ʽ����
  --output
  --        code        N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --        message     C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --        isexist     N 1 1-����;0-������
  --        rcpdtl_ids  C 1 �Ѿ���������Һ���ĵķ���id�����ŷָ�
  ---------------------------------------------------------------------------
  j_Input      PLJson;
  j_Json       PLJson;
  n_Is_Return  Number(1);
  v_����ids    Varchar2(4000);
  n_ҽ��id     ��Һ��ҩ��¼.ҽ��id%Type;
  n_Count      Number(18);
  v_Rcpdtl_Ids Varchar2(32767);
Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_ҽ��id    := j_Json.Get_Number('advice_id');
  v_����ids   := j_Json.Get_String('rcpdtl_ids');
  n_Is_Return := j_Json.Get_Number('is_return');

  If v_����ids Is Not Null Then
    If Nvl(n_Is_Return, 0) = 0 Then
      Select Count(1)
      Into n_Count
      From ��Һ��ҩ���� A, ҩƷ�շ���¼ B
      Where a.�շ�id = b.Id And b.����id In (Select Column_Value From Table(f_Num2List(v_����ids))) And
            Instr(',8,9,10,', ',' || b.���� || ',') > 0;
    End If;
  
    If Nvl(n_Is_Return, 0) = 1 Then
      For R In (Select Distinct b.����id
                From ��Һ��ҩ���� A, ҩƷ�շ���¼ B
                Where a.�շ�id = b.Id And b.����id In (Select Column_Value From Table(f_Num2List(v_����ids))) And
                      Instr(',8,9,10,', ',' || b.���� || ',') > 0) Loop
        v_Rcpdtl_Ids := v_Rcpdtl_Ids || ',' || r.����id;
      End Loop;
      If v_Rcpdtl_Ids Is Not Null Then
        n_Count := 1;
      Else
        n_Count := 0;
      End If;
    End If;
  Else
    Select Count(1)
    Into n_Count
    From ��Һ��ҩ���� A, ҩƷ�շ���¼ B, ��Һ��ҩ��¼ C
    Where a.�շ�id = b.Id And b.ҽ��id = n_ҽ��id And a.��¼id = c.Id And Rownum < 2;
  End If;

  If n_Count <> 0 Then
    n_Count := 1;
  End If;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","rcpdtl_ids":"' || Substr(v_Rcpdtl_Ids, 2) || '","isexist":' ||
              n_Count || '}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Pivassvr_Isexsitinfusion;
/


Create Or Replace Procedure Zl_Pivassvr_Getinfusion_Record
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ����ݲ�����Ϣ�򴦷�����Ϣ����ȡ������Һ��ҩ��¼�еĴ���ids
  --��Σ�Json_In:��ʽ
  --    input
  --        pati_id                 N   1   ����ID
  --        pati_pageids            C   1   ��ҳID
  --        rcpdtl_ids              C   0   ����ö���
  --        rcp_nos                 C   0   ����ö��ţ�rcpdtl_ids����ʱ���˲�����Ч
  --����: Json_Out,��ʽ����
  --    output
  --        code                    N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --        message                 C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --        rcpdtl_ids              C   1   ������ϸid,Ŀǰ����ķ���ID
  ---------------------------------------------------------------------------

  j_Json    PLJson;
  n_����id  ҩƷ�շ���¼.����id%Type;
  v_��ҳids Varchar2(4000);
  n_����id  ҩƷ�շ���¼.����id%Type;
  Type t_��Һ���� Is Ref Cursor;
  c_��Һ���� t_��Һ����;

  j_Json_Tmp PLJson;

  v_������  Varchar2(4000);
  v_Temp    Varchar2(32680);
  v_����ids Varchar2(32680);
  c_����ids Clob;
  n_Count   Number(18);
Begin
  --�������
  j_Json_Tmp := PLJson(Json_In);
  j_Json     := j_Json_Tmp.Get_Pljson('input');
  n_����id   := j_Json.Get_Number('pati_id');
  v_��ҳids  := j_Json.Get_String('pati_pageids');
  c_����ids  := j_Json.Get_String('rcpdtl_ids');
  v_������   := j_Json.Get_Clob('rcp_nos');

  If v_������ Is Not Null Then
    If Instr(v_������, ',') > 0 Then
      Open c_��Һ���� For
        Select Distinct ����id
        From ҩƷ�շ���¼ B1, ��Һ��ҩ���� C1
        Where B1.Id = C1.�շ�id And B1.No In (Select /*+cardinality(j,10) */
                                             Column_Value
                                            From Table(f_Str2List(v_������)) J) And
              Instr(',9,10,', ',' || B1.���� || ',') > 0;
    Else
      Open c_��Һ���� For
        Select Distinct ����id
        From ҩƷ�շ���¼ B1, ��Һ��ҩ���� C1
        Where B1.Id = C1.�շ�id And B1.No = v_������ And Instr(',9,10,', ',' || B1.���� || ',') > 0;
    End If;
  Elsif c_����ids Is Not Null Then
  
    If Length(c_����ids) <= 4000 Then
    
      Open c_��Һ���� For
        Select /*+cardinality(a,10) */
        Distinct B1.����id
        From ҩƷ�շ���¼ B1, ��Һ��ҩ���� C1, (Select Column_Value As ����id From Table(f_Num2List(c_����ids)) J) A
        Where a.����id = B1.����id And B1.Id = C1.�շ�id And Instr(',9,10,', ',' || B1.���� || ',') > 0;
    Else
    
      v_����ids := Null;
      Loop
        Exit When c_����ids Is Null;
      
        If Length(c_����ids) <= 4000 Then
          v_Temp    := Substr(c_����ids, 1);
          c_����ids := Null;
        Else
          n_Count   := Instr(c_����ids, ',', 3900);
          v_Temp    := Substr(c_����ids, 1, n_Count - 1);
          c_����ids := Substr(c_����ids, n_Count + 1);
        End If;
      
        For c_����id In (Select /*+cardinality(a,10) */
                       Distinct B1.����id
                       From ҩƷ�շ���¼ B1, ��Һ��ҩ���� C1, (Select Column_Value As ����id From Table(f_Num2List(v_Temp)) J) A
                       Where a.����id = B1.����id And B1.Id = C1.�շ�id And Instr(',9,10,', ',' || B1.���� || ',') > 0) Loop
          If Instr(Nvl(v_����ids, '') || ',', ',' || c_����id.����id || ',') = 0 Then
            v_����ids := Nvl(v_����ids, '') || ',' || c_����id.����id;
          End If;
        End Loop;
      End Loop;
      If Not v_����ids Is Null Then
        v_����ids := Substr(v_����ids, 2);
      End If;
      Json_Out := '{"output":{"code":1,"message":"�ɹ�","rcpdtl_ids":"' || v_Temp || '"}}';
      Return;
    End If;
  Elsif Nvl(n_����id, 0) <> 0 Then
    Open c_��Һ���� For
      Select /*+cardinality(a,10) */
      Distinct B1.����id
      From ҩƷ�շ���¼ B1, ��Һ��ҩ���� C1
      Where B1.����id = n_����id And (Instr(',' || v_��ҳids || ',', ',' || Nvl(B1.��ҳid, 0) || ',') > 0 Or v_��ҳids Is Null) And
            B1.Id = C1.�շ�id And Instr(',9,10,', ',' || B1.���� || ',') > 0;
  Else
    Json_Out := zlJsonOut('����ȷ�����λ�ȡ���ݵ����������飡');
    Return;
  End If;
  v_����ids := Null;
  Loop
    Fetch c_��Һ����
      Into n_����id;
    Exit When c_��Һ����%NotFound;
  
    If v_����ids Is Null Then
      v_����ids := '' || n_����id;
    Else
      v_����ids := v_����ids || ',' || n_����id;
    End If;
  End Loop;
  Close c_��Һ����;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","rcpdtl_ids":"' || v_����ids || '"}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Pivassvr_Getinfusion_Record;
/

Create Or Replace Procedure Zl_Pivassvr_Newbill
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ�ҽ�����ͺ������Һ��ҩ��¼
  --��Σ�Json_In:��ʽ
  --input      ������Һ��ҩ��¼�������˵���
  --  operator_name                   C 1 ������(����Ա����)
  --  operator_time                   C 1 ����ʱ��
  --  pati_id                         N 1 ����id
  --  page_id                         N 1 ��ҳID
  --  pati_name                       C 1 ����
  --  pati_sex                        C 1 �Ա�
  --  pati_age                        C 1 ����
  --  inpatient_num                   C 1 סԺ��
  --  pati_bed                        C 1 ����
  --  pati_wardarea_id                N 1 ���˲���id
  --  pati_deptid                     N 1 ���˿���id
  --  advice_list[]��ҽ��������
  --    pivas_deptid                  N 1 ��������id
  --    advice_id                     N 1 ��ҽ��ID(��ҩ;��)
  --    advice_send_no                N 1 ���ͺ�
  --    effective_time                N 1 ҽ����Ч��0-������1-����
  --    drug_method_id                N 1 ��ҩ;��id
  --    is_tpn                        N 1 �Ƿ�tpn��0-���ǣ�1-��
  --    advice_frequency              C 1 ִ��Ƶ��
  --    advice_drug_list[]��ҩ;����Ӧ��ҩ��������
  --            advice_id             N 1 ҩ��id
  --            advice_rcpno          C 1 ҩ�����Ͳ����ķ���no
  --    advice_exetime_list[]ҽ��ִ��ʱ�䣬��ҩ;��ҽ����ִ��ʱ�䣬��ʱ�ṩ��ҽ�����з��͵�ʱ�䣬�������η��͵�ִ��ʱ�䡣�����ͺŵ�����֯��������


  --            advice_send_no        N 1 ��ҩ;��ҽ���ķ��ͺ�
  --            advice_require_time   C 1 Ҫ��ʱ��

  --����: Json_Out,��ʽ����
  --output
  --  code                          C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message                       C 1 Ӧ����Ϣ:ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------

  j_Pivas        PLJson;
  j_Input        PLJson;
  Jl_Pid_Main    Pljson_List;
  j_Pid_Main     PLJson;
  Jl_Pid_Exetime Pljson_List;
  j_Pid_Exetime  PLJson;
  Jl_Pid_Drug    Pljson_List;
  j_Pid_Drug     PLJson;

  n_���id_m     ��Һ��ҩ��¼.ҽ��id%Type;
  d_���ͺ�_m     ��Һ��ҩ��¼.���ͺ�%Type;
  v_����_m       ��Һ��ҩ��¼.����%Type;
  v_�Ա�_m       ��Һ��ҩ��¼.�Ա�%Type;
  v_����_m       ��Һ��ҩ��¼.����%Type;
  v_סԺ��_m     ��Һ��ҩ��¼.סԺ��%Type;
  v_����_m       ��Һ��ҩ��¼.����%Type;
  n_���˲���id_m ��Һ��ҩ��¼.���˲���id%Type;
  n_���˿���id_m ��Һ��ҩ��¼.���˿���id%Type;
  d_����ʱ��_m   ��Һ��ҩ��¼.����ʱ��%Type;
  n_ҽ������_m   Number(1);
  n_��ҩ;��id_m ������ĿĿ¼.Id%Type;
  n_����id_m     ��Һ��ҩ��¼.����id%Type;
  n_�Ƿ�tpn_m    Number(1);
  v_ִ��Ƶ��_m   Varchar2(100);
  n_��ҳid_m     ��Һ��ҩ��¼.��ҳid%Type;
  v_ִ��ʱ��s_m  Varchar2(4000);

  n_ҽ��id_d  ��Һ��ҩ��¼.ҽ��id%Type;
  v_����no_d  ҩƷ�շ���¼.No%Type;
  v_ҽ��ids_d Varchar2(32767);

  n_��������id ��Һ��ҩ��¼.����id%Type;
  v_�˲���     ��Һ��ҩ��¼.������Ա%Type;
  v_�˲�ʱ��   ��Һ��ҩ��¼.����ʱ��%Type;

  v_��ҽ��   Varchar2(32767);
  v_ҩ��     Varchar2(32767);
  v_ҽ����Ϣ Varchar2(32767);
  c_ҽ����Ϣ Clob;
Begin
  If Json_In Is Null Then
    Json_Out := zlJsonOut('δ�������ݣ�����');
    Return;
  End If;

  j_Pivas := PLJson(Json_In);
  If j_Pivas Is Not Null Then
    --�������ݵ���������
    j_Input := j_Pivas.Get_Pljson('input');
  
    v_�˲���       := j_Input.Get_String('operator_name');
    v_�˲�ʱ��     := To_Date(j_Input.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');
    d_����ʱ��_m   := To_Date(j_Input.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');
    n_����id_m     := j_Input.Get_Number('pati_id');
    n_��ҳid_m     := j_Input.Get_Number('page_id');
    v_����_m       := j_Input.Get_String('pati_name');
    v_�Ա�_m       := j_Input.Get_String('pati_sex');
    v_����_m       := j_Input.Get_String('pati_age');
    v_����_m       := j_Input.Get_String('pati_bed');
    v_סԺ��_m     := j_Input.Get_Number('inpatient_num');
    n_���˲���id_m := j_Input.Get_Number('pati_wardarea_id');
    n_���˿���id_m := j_Input.Get_Number('pati_deptid');
  
    --ҽ����Ϣ����ҽ�����������������ݽڵ�
    Jl_Pid_Main := Pljson_List();
    Jl_Pid_Main := j_Input.Get_Pljson_List('advice_list');
    For I In 1 .. Jl_Pid_Main.Count Loop
    
      j_Pid_Main := PLJson();
      j_Pid_Main := PLJson(Jl_Pid_Main.Get(I));
    
      v_ҩ��        := Null;
      v_ִ��ʱ��s_m := Null;
    
      --ҽ�������Ϣ
      n_��������id   := j_Pid_Main.Get_Number('pivas_deptid');
      n_���id_m     := j_Pid_Main.Get_Number('advice_id');
      d_���ͺ�_m     := j_Pid_Main.Get_Number('advice_send_no');
      n_ҽ������_m   := j_Pid_Main.Get_Number('effective_time');
      n_��ҩ;��id_m := j_Pid_Main.Get_Number('drug_method_id');
      n_�Ƿ�tpn_m    := j_Pid_Main.Get_Number('is_tpn');
      v_ִ��Ƶ��_m   := j_Pid_Main.Get_String('advice_frequency');
    
      v_��ҽ�� := n_��������id || ',' || n_���id_m || ',' || d_���ͺ�_m || ',' || n_ҽ������_m || ',' || n_��ҩ;��id_m || ',' ||
               n_�Ƿ�tpn_m || ',' || v_ִ��Ƶ��_m;
    
      --ҽ��ִ��ʱ��ֽ⣬������ִ��������ѯ�õ�
      Jl_Pid_Exetime := Pljson_List();
      Jl_Pid_Exetime := j_Pid_Main.Get_Pljson_List('advice_exetime_list');
      For N In 1 .. Jl_Pid_Exetime.Count Loop
        j_Pid_Exetime := PLJson();
        j_Pid_Exetime := PLJson(Jl_Pid_Exetime.Get(N));
      
        --��ʽ��ִ��ʱ�䴮��Ҫ��ʱ��,���ͺ�|...
        If v_ִ��ʱ��s_m Is Null Then
          v_ִ��ʱ��s_m := j_Pid_Exetime.Get_String('advice_require_time') || ',' ||
                       j_Pid_Exetime.Get_Number('advice_send_no');
        Else
          If Length(v_ִ��ʱ��s_m || '|' || j_Pid_Exetime.Get_String('advice_require_time') || ',') > 4000 Then
            --����4K�Ͳ�Ҫ��������ݣ�������ǰ����㹻��
            Exit;
          Else
            v_ִ��ʱ��s_m := v_ִ��ʱ��s_m || '|' || j_Pid_Exetime.Get_String('advice_require_time') || ',' ||
                         j_Pid_Exetime.Get_Number('advice_send_no');
          End If;
        End If;
      End Loop;
    
      --�ֽ�ҩ�����Ȳ���ҩ������ҩ����Ӧ�ķ��ͷ���NO
      v_ҽ��ids_d := Null;
      Jl_Pid_Drug := Pljson_List();
      Jl_Pid_Drug := j_Pid_Main.Get_Pljson_List('advice_drug_list');
      For M In 1 .. Jl_Pid_Drug.Count Loop
        j_Pid_Drug := PLJson();
        j_Pid_Drug := PLJson(Jl_Pid_Drug.Get(M));
        n_ҽ��id_d := j_Pid_Drug.Get_Number('advice_id');
      
        If v_ҽ��ids_d Is Null Then
          v_ҽ��ids_d := n_ҽ��id_d;
        Else
          v_ҽ��ids_d := v_ҽ��ids_d || ',' || n_ҽ��id_d;
        End If;
      
        v_����no_d := j_Pid_Drug.Get_String('advice_rcpno');
      End Loop;
      v_ҩ�� := v_����no_d || '|' || v_ҽ��ids_d;
    
      If v_ҽ����Ϣ Is Null Then
        v_ҽ����Ϣ := v_��ҽ�� || ';' || v_ҩ�� || ';' || v_ִ��ʱ��s_m;
      Elsif Length(v_ҽ����Ϣ || '||' || v_��ҽ�� || ';' || v_ҩ�� || ';' || v_ִ��ʱ��s_m) > 4000 Then
      
        If c_ҽ����Ϣ Is Null Then
          c_ҽ����Ϣ := v_ҽ����Ϣ;
        Else
          c_ҽ����Ϣ := c_ҽ����Ϣ || '||' || v_ҽ����Ϣ;
        End If;
        v_ҽ����Ϣ := v_��ҽ�� || ';' || v_ҩ�� || ';' || v_ִ��ʱ��s_m;
      
      Else
        v_ҽ����Ϣ := v_ҽ����Ϣ || '||' || v_��ҽ�� || ';' || v_ҩ�� || ';' || v_ִ��ʱ��s_m;
      End If;
    End Loop;
  
    If c_ҽ����Ϣ Is Not Null Then
      c_ҽ����Ϣ := c_ҽ����Ϣ || '||' || v_ҽ����Ϣ;
      v_ҽ����Ϣ := Null;
    End If;
  
    --��ҽ��(��������ID,��ҽ��ID,���ͺ�,ҽ����Ч,��ҩ;��id,�Ƿ�tpn,ִ��Ƶ��);ҩ��(����NO|ҩ��id,...);ҽ������(����ʱ��,���ͺ�|...)||...
    Zl_��Һ��ҩ��¼_Insert_s(v_�˲���, v_�˲�ʱ��, d_����ʱ��_m, n_����id_m, n_��ҳid_m, v_����_m, v_�Ա�_m, v_����_m, v_����_m, v_סԺ��_m, n_���˲���id_m,
                       n_���˿���id_m, v_ҽ����Ϣ, c_ҽ����Ϣ);
  End If;

  Json_Out := '{"output":{"code":1,"message": "�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Pivassvr_Newbill;
/


Create Or Replace Procedure Zl_Pivassvr_Locked
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡ��Һ��Һ��¼�Ƿ�����
  --��Σ�Json_In:��ʽ
  --  input
  --     pivas_ids      C  0   ��Һidƴ�������ŷָ�
  --     advice_ids     C  0   ҽ��idƴ�������ŷָ�
  --     send_no        C  0   ���ͺ�ƴ��,���ŷָ�
  --     advice_info    C  0   ��ʽ:ҽ��ID1,ִ����ֹʱ��1;ҽ��ID2,ִ����ֹʱ��2;.......... ִ����ֹʱ��Ϊ���ڸ�ʽ
  --     query_type     N  1   ��ѯ��ʽ��0-ֻ�ж��Ƿ�棬������Һid��1-�ж��Ƿ�棬��������Һid ��2-ֻ�ж��Ƿ�棬������Һid(����ҽ��ID��ѯ)��3-�ж��Ƿ�棬��������Һid(����ҽ��ID��ѯ)
  --                                     4-��ҽ��id�ͷ��ͺŲ�ѯ��Һ��Ϣ,���ַ�ʽʱ,ҽ��id�ͷ��ͺ�ֻ��һ��ֵ�����ж���,����ֵ�б�ֵ
  --                                     5-�ж��Ƿ������������Һҽ��,������Ϣ  advice_info �����ж��Ƿ�棬���ҷ����б� ��Һid,����״̬���Ƿ���

  --����: Json_Out,��ʽ����

  --  output
  --    code            N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message         C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    isexist         N 1 �Ƿ�������1-������0-δ����
  --    pivas_ids       C 1 ��Һidƴ�������ŷָ�
  --    item_list
  --       pivas_id     N 1 ��ҺID
  --       status       N 1 ����״̬
  --       is_package   N 0 �Ƿ���
  --       order_id     N 0 ҽ��id

  ---------------------------------------------------------------------------

  j_Json       PLJson;
  j_Json_Tmp   PLJson;
  n_Cnt        Number(2) := 0;
  v_��Һids    Varchar2(32767);
  v_ҽ��ids    Varchar2(32767);
  v_���ͺ�     Varchar2(32767);
  n_��ѯ��ʽ   Number(1);
  v_��Һoutids Varchar2(32767);
  v_ҽ����Ϣ   Varchar2(32767);
  v_Jtmp       Varchar2(32767);
  c_Jtmp       Clob;

Begin
  --�������
  j_Json_Tmp := PLJson(Json_In);
  j_Json     := j_Json_Tmp.Get_Pljson('input');
  n_��ѯ��ʽ := Nvl(j_Json.Get_Number('query_type'), 0);
  If n_��ѯ��ʽ > 1 Then
    v_ҽ��ids := j_Json.Get_String('advice_ids');
    v_���ͺ�  := j_Json.Get_String('send_no');
  Else
    v_��Һids := j_Json.Get_String('pivas_ids');
  End If;

  If n_��ѯ��ʽ = 0 Then
    Select Count(1)
    Into n_Cnt
    From ��Һ��ҩ��¼ A
    Where a.Id In (Select /*+cardinality(E,10)*/
                    e.Column_Value
                   From Table(f_Num2List(v_��Һids)) E) And a.�Ƿ����� = 1;
  
    If n_Cnt > 0 Then
      n_Cnt := 1;
    End If;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","isexist":' || n_Cnt || '}}';
  Elsif n_��ѯ��ʽ = 1 Then
  
    For R In (Select a.Id
              From ��Һ��ҩ��¼ A
              Where a.Id In (Select /*+cardinality(E,10)*/
                              e.Column_Value
                             From Table(f_Num2List(v_��Һids)) E) And a.�Ƿ����� = 1) Loop
    
      v_��Һoutids := v_��Һoutids || ',' || r.Id;
    
    End Loop;
    v_��Һoutids := Substr(v_��Һoutids, 2);
  
    If v_��Һoutids Is Not Null Then
      n_Cnt := 1;
    Else
      n_Cnt := 0;
    End If;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","isexist":' || n_Cnt || ',"pivas_ids":"' || v_��Һoutids || '"}}';
  Elsif n_��ѯ��ʽ = 2 Then
    Select Count(1)
    Into n_Cnt
    From ��Һ��ҩ��¼ A
    Where a.ҽ��id In (Select /*+cardinality(E,10)*/
                      e.Column_Value
                     From Table(f_Num2List(v_ҽ��ids)) E) And
          (a.���ͺ� In (Select /*+cardinality(E,10)*/
                      g.Column_Value
                     From Table(f_Num2List(v_���ͺ�)) G) Or Nvl(v_���ͺ�, '��') = '��') And a.�Ƿ����� = 1;
  
    If n_Cnt > 0 Then
      n_Cnt := 1;
    End If;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","isexist":' || n_Cnt || '}}';
  
  Elsif n_��ѯ��ʽ = 3 Then
  
    For R In (Select a.Id
              From ��Һ��ҩ��¼ A
              Where a.ҽ��id In (Select /*+cardinality(E,10)*/
                                e.Column_Value
                               From Table(f_Num2List(v_ҽ��ids)) E) And
                    (a.���ͺ� In (Select /*+cardinality(E,10)*/
                                g.Column_Value
                               From Table(f_Num2List(v_���ͺ�)) G) Or Nvl(v_���ͺ�, '��') = '��') And a.�Ƿ����� = 1) Loop
    
      v_��Һoutids := v_��Һoutids || ',' || r.Id;
    
    End Loop;
    v_��Һoutids := Substr(v_��Һoutids, 2);
  
    If v_��Һoutids Is Not Null Then
      n_Cnt := 1;
    Else
      n_Cnt := 0;
    End If;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","isexist":' || n_Cnt || ',"pivas_ids":"' || v_��Һoutids || '"}}';
  Elsif n_��ѯ��ʽ = 4 Then
    For R In (Select a.Id, a.�Ƿ�����, a.����״̬
              From ��Һ��ҩ��¼ A
              Where a.ҽ��id = To_Number(v_ҽ��ids) And a.���ͺ� = To_Number(v_���ͺ�)) Loop
      If Nvl(r.�Ƿ�����, 0) = 1 Then
        n_Cnt := 1;
      End If;
    
      v_Jtmp := v_Jtmp || ',{"pivas_id":' || r.Id;
      v_Jtmp := v_Jtmp || ',"status":' || r.����״̬;
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
  
    If n_Cnt > 0 Then
      n_Cnt := 1;
    End If;
    If c_Jtmp Is Null Then
      Json_Out := '{"output":{"code":1,"message":"�ɹ�","isexist":' || n_Cnt || ',"item_list":[' || Substr(v_Jtmp, 2) ||
                  ']}}';
    Else
      c_Jtmp   := c_Jtmp || v_Jtmp;
      Json_Out := '{"output":{"code":1,"message":"�ɹ�","isexist":' || n_Cnt || ',"item_list":[' || c_Jtmp || ']}}';
    End If;
  
  Elsif n_��ѯ��ʽ = 5 Then
    v_ҽ����Ϣ := j_Json.Get_String('advice_info');
    For R In (Select a.ҽ��id, a.Id, a.�Ƿ�����, a.����״̬, a.�Ƿ���, a.���ͺ�
              From ��Һ��ҩ��¼ A,
                   (Select /*+cardinality(b,10)*/
                      To_Number(C1) As ҽ��id, To_Date(C2, 'yyyy-mm-dd hh24:mi:ss') As ִ����ֹʱ��
                     From Table(Cast(f_Str2List2(v_ҽ����Ϣ, ';', ',') As t_StrList2)) B) X
              Where a.ҽ��id = x.ҽ��id And a.ִ��ʱ�� > x.ִ����ֹʱ��) Loop
      If Nvl(r.�Ƿ�����, 0) = 1 Then
        n_Cnt := 1;
      End If;
      --��ҽ��ҽ���� ��ҩidһ�𷵻س�ȥ������һ��ʹ��    
      v_Jtmp := v_Jtmp || ',{"pivas_id":' || r.Id;
      v_Jtmp := v_Jtmp || ',"status":' || r.����״̬;
      v_Jtmp := v_Jtmp || ',"is_package":' || Nvl(r.�Ƿ��� || '', 'null');
      v_Jtmp := v_Jtmp || ',"order_id":' || r.ҽ��id;
      v_Jtmp := v_Jtmp || ',"send_no":' || r.���ͺ�;
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
    If n_Cnt > 0 Then
      n_Cnt := 1;
    End If;
    If c_Jtmp Is Null Then
      Json_Out := '{"output":{"code":1,"message":"�ɹ�","isexist":' || n_Cnt || ',"item_list":[' || Substr(v_Jtmp, 2) ||
                  ']}}';
    Else
      c_Jtmp   := c_Jtmp || v_Jtmp;
      Json_Out := '{"output":{"code":1,"message":"�ɹ�","isexist":' || n_Cnt || ',"item_list":[' || c_Jtmp || ']}}';
    End If;
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Pivassvr_Locked;
/

Create Or Replace Procedure Zl_Pivassvr_Getinfo_Batch
(
  Json_In  In Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡ��Һ��Һ��¼�嵥
  --��Σ�Json_In:��ʽ
  --  input
  --    query_type                          N  1  ��ѯ��ʽ   0-����Һ��ҩ��¼.id��ѯ by_pivas_id���봫ֵ
  --                                                         1-��ҽ��id���ͺŲ�ѯ��ҽ��id�ͷ��ͺ� ��ѯ
  --                                                         2-����Һʱ�䡢���״̬������״̬��ѯ
  --                                                         3-��ҽ��IDƴ����ѯ�����ڳ����ջ�ʱҽ���������Һ��ҽ�������ջ� order_ids ��Ҫ��ֵ
  --                                                         4-��ҽ��id��ѯ�ж�ҽ���Ƿ����������Һ��ҩ��¼
  --                                                         5-��ҽ��id��ѯ�ж�ҽ����״̬�ж��Ƿ���ҽ��վ�·���ʾ��Һ��Ϣҳ��
  --                                                         6-��ҽ��ID ��ѯ��ǰ����ҽ������ҩ��¼��Ϣ��order_id ��ʾ���Ǹ�ҩ;���е�ҽ��id
  --                                                                    ����Ϊorder_and_no ҽ��+���ͺ�ƴ�������ź�ð�ŷָ�
  --    by_pivas_id                         N  1   ��Һ��ҩ��¼.id ���ý�㴫ֵ�󰴵�����ѯ
  --    order_id                            N  1   ҽ��id
  --    send_no                             N  1   ���ͺ�
  --    begin_time                          C  1   ��ʼʱ��
  --    end_time                            C  1   ����ʱ��
  --    is_package                          N  1   �Ƿ���
  --    operator_status                     C  1   ����״̬���ַ���������״̬ƴ��
  --    order_ids                           C  1   ҽ��ids

  --����: Json_Out,��ʽ����

  --  output
  --    code                                N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                             C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    isexist                             N 1 ��ѯ��ʽΪ4ʱ�Ƿ���ڼ�¼
  --    order_and_no                        C 0 ҽ��+���ͺ�ƴ�������ź�ð�ŷָ�
  --    item_list
  --       pivas_id                         N 1 id
  --       order_id                         N 1 ҽ��id
  --       dept_id                          N 1 ��Һ����ID
  --       exe_time                         C 1 ִ��ʱ��
  --       is_package                       N 1 �Ƿ���
  --       is_locked                        N 1 �Ƿ�����
  --       bottle_label                     C 1 ƿǩ��
  --       status                           C 1 ״̬����������
  --       name                             C 1 ��������
  --       inpatientnum                     C 1 סԺ��
  --       pati_bed                         C 1 ����
  --       pati_deptid                      N 1 ���˿���id
  --       send_no                          N 1 ���ͺ�
  --       pivas_batchno                    N 1 ��ҩ����
  --       pivas_work_time                  C 1 ��ҩ����ʱ��
  --       package_time                     C 1 ���ʱ��
  --       oper_type                        N 1 ����״̬����������
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  n_ҽ��id       Number(18);
  n_���ͺ�       Number(18);
  n_��ҩid       Number(18);
  n_��ѯ��ʽ     Number(3);
  d_��ʼʱ��     Date;
  d_����ʱ��     Date;
  n_�Ƿ���     Number(1);
  v_����״̬     Varchar2(30);
  v_ҽ��ids      Varchar2(32767);
  v_Order_And_No Varchar2(32767);
  n_Cnt          Number;
  v_Jtmp         Varchar2(32767);
  c_Jtmp         Clob;
  Cursor c_Pivas Is
    Select b.Id As ��ҩid, b.���ͺ�, b.ҽ��id, To_Char(b.ִ��ʱ��, 'YYYY-MM-DD HH24:MI') As ִ��ʱ��, b.��ҩ����, g.��ҩʱ�� As ��ҩ����ʱ��, b.ƿǩ��,
           Decode(b.����״̬, 1, '����ҩ', 2, '����ҩ', 3, '����ҩ', 4, '����ҩ', '�ѷ���') As ״̬,
           To_Char(b.���ʱ��, 'YYYY-MM-DD HH24:MI') As ���ʱ��
    From ��Һ��ҩ��¼ B, ��ҩ�������� G
    Where b.��ҩ���� = g.����(+) And b.����id = g.��������id(+) And b.ִ��ʱ�� Between d_��ʼʱ�� And d_����ʱ�� And Nvl(b.�Ƿ���, 0) = n_�Ƿ��� And
          (Instr(',' || v_����״̬ || ',', ',' || b.����״̬ || ',') > 0 Or v_����״̬ Is Null);

Begin
  --�������
  j_Jsonin := PLJson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');

  n_��ѯ��ʽ := j_Json.Get_Number('query_type');
  n_��ѯ��ʽ := Nvl(n_��ѯ��ʽ, 0);
  If n_��ѯ��ʽ = 0 Then
    --ֻ����һ������
    n_��ҩid := j_Json.Get_Number('by_pivas_id');
    For R In (Select a.Id, a.ҽ��id, a.����id As ��Һ����id, To_Char(a.ִ��ʱ��, 'YYYY-MM-DD HH24:MI') As ִ��ʱ��, a.�Ƿ���, a.��ҩ����, a.ƿǩ��,
                     Decode(a.����״̬, 1, '����ҩ', 2, '����ҩ', 3, '����ҩ', 4, '����ҩ', 5, '�ѷ���', 6, '��ǩ��', 7, '�Ѿܾ�ǩ��', 8, '��ȷ�Ͼ���', 9,
                             '����������', 10, '���������', 11, '���ʾܾ�', '�ѷ���') As ״̬, a.����, a.סԺ��, a.����, a.���˿���id, a.�Ƿ�����, a.����״̬
              From ��Һ��ҩ��¼ A
              Where a.Id = n_��ҩid) Loop
    
      v_Jtmp := v_Jtmp || '{"pivas_id":' || r.Id;
      v_Jtmp := v_Jtmp || ',"order_id":' || r.ҽ��id;
      v_Jtmp := v_Jtmp || ',"dept_id":' || r.��Һ����id;
      v_Jtmp := v_Jtmp || ',"exe_time":"' || r.ִ��ʱ�� || '"';
      v_Jtmp := v_Jtmp || ',"is_package":' || Nvl(r.�Ƿ���, 0);
      v_Jtmp := v_Jtmp || ',"is_locked":' || Nvl(r.�Ƿ�����, 0);
      v_Jtmp := v_Jtmp || ',"pivas_batchno":' || Nvl(r.��ҩ���� || '', 'null');
      v_Jtmp := v_Jtmp || ',"bottle_label":"' || zlJsonStr(r.ƿǩ��) || '"';
      v_Jtmp := v_Jtmp || ',"status":"' || r.״̬ || '"';
      v_Jtmp := v_Jtmp || ',"oper_type":' || Nvl(r.����״̬, 0);
      v_Jtmp := v_Jtmp || ',"name":"' || zlJsonStr(r.����) || '"';
      v_Jtmp := v_Jtmp || ',"inpatientnum":"' || r.סԺ�� || '"';
      v_Jtmp := v_Jtmp || ',"pati_bed":"' || zlJsonStr(r.����) || '"';
      v_Jtmp := v_Jtmp || ',"pati_deptid":' || Nvl(r.���˿���id || '', 'null');
      v_Jtmp := v_Jtmp || '}';
    
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || v_Jtmp || ']}}';
  
  Elsif n_��ѯ��ʽ = 1 Then
    n_ҽ��id := j_Json.Get_Number('order_id');
    n_���ͺ� := j_Json.Get_Number('send_no');
    For R In (Select a.Id, a.ҽ��id, a.����id As ��Һ����id, To_Char(a.ִ��ʱ��, 'YYYY-MM-DD HH24:MI') As ִ��ʱ��, a.�Ƿ���, a.��ҩ����, a.ƿǩ��,
                     Decode(����״̬, 1, '����ҩ', 2, '����ҩ', 3, '����ҩ', 4, '����ҩ', 5, '�ѷ���', 6, '��ǩ��', 7, '�Ѿܾ�ǩ��', 8, '��ȷ�Ͼ���', 9,
                             '����������', 10, '���������', 11, '���ʾܾ�', '�ѷ���') As ״̬, a.����, a.סԺ��, a.����, a.���˿���id, a.�Ƿ�����, a.����״̬
              From ��Һ��ҩ��¼ A
              Where a.ҽ��id = n_ҽ��id And (a.���ͺ� = n_���ͺ� Or n_���ͺ� = 0) And a.����״̬ <> 8
              Order By a.ִ��ʱ��) Loop
    
      v_Jtmp := v_Jtmp || ',{"pivas_id":' || r.Id;
      v_Jtmp := v_Jtmp || ',"order_id":' || r.ҽ��id;
      v_Jtmp := v_Jtmp || ',"dept_id":' || r.��Һ����id;
      v_Jtmp := v_Jtmp || ',"exe_time":"' || r.ִ��ʱ�� || '"';
      v_Jtmp := v_Jtmp || ',"is_package":' || Nvl(r.�Ƿ���, 0);
      v_Jtmp := v_Jtmp || ',"is_locked":' || Nvl(r.�Ƿ�����, 0);
      v_Jtmp := v_Jtmp || ',"pivas_batchno":' || Nvl(r.��ҩ���� || '', 'null');
      v_Jtmp := v_Jtmp || ',"bottle_label":"' || zlJsonStr(r.ƿǩ��) || '"';
      v_Jtmp := v_Jtmp || ',"status":"' || r.״̬ || '"';
      v_Jtmp := v_Jtmp || ',"oper_type":' || Nvl(r.����״̬, 0);
      v_Jtmp := v_Jtmp || ',"name":"' || zlJsonStr(r.����) || '"';
      v_Jtmp := v_Jtmp || ',"inpatientnum":"' || r.סԺ�� || '"';
      v_Jtmp := v_Jtmp || ',"pati_bed":"' || zlJsonStr(r.����) || '"';
      v_Jtmp := v_Jtmp || ',"pati_deptid":' || Nvl(r.���˿���id || '', 'null');
      v_Jtmp := v_Jtmp || '}';
    
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Elsif n_��ѯ��ʽ = 2 Then
  
    d_��ʼʱ�� := To_Date(j_Json.Get_String('begin_time'), 'YYYY-MM-DD HH24:MI:SS');
    d_����ʱ�� := To_Date(j_Json.Get_String('end_time'), 'YYYY-MM-DD HH24:MI:SS');
    n_�Ƿ��� := j_Json.Get_Number('is_package');
    v_����״̬ := j_Json.Get_String('operator_status');
  
    For R In c_Pivas Loop
    
      v_Jtmp := v_Jtmp || '{"pivas_id":' || r.��ҩid;
      v_Jtmp := v_Jtmp || ',"order_id":' || r.ҽ��id;
      v_Jtmp := v_Jtmp || ',"send_no":' || r.���ͺ�;
      v_Jtmp := v_Jtmp || ',"exe_time":"' || r.ִ��ʱ�� || '"';
      v_Jtmp := v_Jtmp || ',"pivas_work_time":"' || r.��ҩ����ʱ�� || '"';
      v_Jtmp := v_Jtmp || ',"pivas_batchno":' || Nvl(r.��ҩ���� || '', 'null');
      v_Jtmp := v_Jtmp || ',"bottle_label":"' || zlJsonStr(r.ƿǩ��) || '"';
      v_Jtmp := v_Jtmp || ',"status":"' || r.״̬ || '"';
      v_Jtmp := v_Jtmp || ',"package_time":"' || r.���ʱ�� || '"';
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
    If c_Jtmp Is Null Then
      Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
    Else
      c_Jtmp   := c_Jtmp || v_Jtmp;
      Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || c_Jtmp || ']}}';
    End If;
  Elsif n_��ѯ��ʽ = 3 Then
    v_ҽ��ids := j_Json.Get_String('order_ids');
    For R In (Select b.ҽ��id, To_Char(Max(b.ִ��ʱ��), 'YYYY-MM-DD HH24:MI:SS') As ִ��ʱ��
              From ��Һ��ҩ��¼ B
              Where (b.����״̬ In (4, 5, 6, 7, 8) And Nvl(b.�Ƿ���, 0) = 0) And
                    b.ҽ��id In (Select /*+cardinality(x,10)*/
                                x.Column_Value
                               From Table(Cast(f_Num2List(v_ҽ��ids) As Zltools.t_Numlist)) X)
              Group By b.ҽ��id) Loop
    
      v_Jtmp := v_Jtmp || ',{"order_id":' || r.ҽ��id;
      v_Jtmp := v_Jtmp || ',"exe_time":"' || r.ִ��ʱ�� || '"';
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
  
    If c_Jtmp Is Null Then
      Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
    Else
      c_Jtmp   := c_Jtmp || v_Jtmp;
      Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || c_Jtmp || ']}}';
    End If;
  
  Elsif n_��ѯ��ʽ = 4 Then
    n_ҽ��id := j_Json.Get_Number('order_id');
    Select Count(1) Into n_Cnt From ��Һ��ҩ��¼ A Where a.ҽ��id = n_ҽ��id;
    If n_Cnt > 0 Then
      n_Cnt := 1;
    End If;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","isexist":' || n_Cnt || '}}';
  Elsif n_��ѯ��ʽ = 5 Then
    n_ҽ��id   := j_Json.Get_Number('order_id');
    v_����״̬ := j_Json.Get_String('operator_status');
    Select Count(1)
    Into n_Cnt
    From ��Һ��ҩ��¼ A
    Where a.ҽ��id = n_ҽ��id And Instr(',' || v_����״̬ || ',', ',' || a.����״̬ || ',') = 0;
    If n_Cnt > 0 Then
      n_Cnt := 1;
    End If;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","isexist":' || n_Cnt || '}}';
  Elsif n_��ѯ��ʽ = 6 Then
    n_ҽ��id := j_Json.Get_Number('order_id');
    For R In (Select a.ҽ��id || ':' || a.���ͺ� As ֵ From ��Һ��ҩ��¼ A Where a.ҽ��id = n_ҽ��id) Loop
      v_Order_And_No := v_Order_And_No || ',' || r.ֵ;
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","order_and_no":"' || Substr(v_Order_And_No, 2) || '"}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Pivassvr_Getinfo_Batch;
/

Create Or Replace Procedure Zl_Pivassvr_Pivasupdate
(
  Json_In  Clob,
  Json_Out Out Clob
) Is

  --------------------------------------------------------------------------------------------------------------------
  --���ܣ���Һ��ҩ��¼�޸�
  --��Σ�Json_In,��ʽ����
  --  input
  --    item_list
  --         pivas_id                   N   1   ��ҩid
  --         is_package                 N   1   �Ƿ���
  --         batch_no                   N   1   ��ҩ����
  --         checker                    C   1   �˲���
  --         check_time                 D   1   �˲�ʱ��
  --����: Json_Out,��ʽ����
  --  output
  --    code                            N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                         C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --------------------------------------------------------------------------------------------------------------------

  n_��ҩid   Number(18);
  n_�Ƿ��� ��Һ��ҩ��¼.�Ƿ���%Type;
  n_��ҩ���� ��Һ��ҩ��¼.��ҩ����%Type;
  v_�˲���   ��Һ��ҩ״̬.������Ա%Type;
  d_�˲�ʱ�� ��Һ��ҩ״̬.����ʱ��%Type;

  j_Json  Pljson;
  Jl_Item Pljson_List;
  j_Item  Pljson;

  n_Ct    Number;
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --�������
  j_Item  := Pljson(Json_In);
  j_Json  := j_Item.Get_Pljson('input');
  Jl_Item := j_Json.Get_Pljson_List('item_list');
  For I In 1 .. Jl_Item.Count Loop
    j_Item     := Pljson();
    j_Item     := Pljson(Jl_Item.Get(I));
    n_��ҩid   := j_Item.Get_Number('pivas_id');
    n_�Ƿ��� := j_Item.Get_Number('is_package');
    n_��ҩ���� := j_Item.Get_Number('batch_no');
    v_�˲���   := j_Item.Get_String('checker');
    d_�˲�ʱ�� := To_Date(j_Item.Get_String('check_time'), 'YYYY-MM-DD HH24:MI:SS');
    Select Count(1)
    Into n_Ct
    From ��Һ��ҩ��¼
    Where ID = n_��ҩid And ����״̬ In (1, 2, 3) And Nvl(��ҩ����, 0) <> Nvl(n_��ҩ����, 0);
    Update ��Һ��ҩ��¼
    Set �Ƿ��� = n_�Ƿ���, ��ҩ���� = n_��ҩ����, ���ʱ�� = Decode(n_�Ƿ���, 1, d_�˲�ʱ��, Null), ���α�� = Decode(n_Ct, 1, 2, ���α��)
    Where ID = n_��ҩid And ����״̬ In (1, 2, 3);
    If Sql%RowCount = 0 Then
      v_Error := '���ڲ�������,��ǰ�޸ĵ���ҩ��¼����ҩ,����ʧ��.';
      Raise Err_Custom;
    End If;
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Pivassvr_Pivasupdate;
/

Create Or Replace Procedure Zl_Pivassvr_Getstatus_Info
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡ��Һ��Һ��¼״̬��Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --     pivas_ids    C  1   ��Һids������ƴ��
  --����: Json_Out,��ʽ����
  --  output
  --    code            N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message         C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    item_list
  --       pivas_id      N 1 ��ҩID
  --       oper_type     N 1 ��������
  --       operator_name C 1 ����Ա
  --       operator_time C 1 ����ʱ��
  --       operator_notes C 1 ����˵��
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  v_��Һids Varchar2(32767);
  v_Jtmp    Varchar2(32767);
  c_Jtmp    Clob;

  v_Vals Clob;
  l_Vals t_StrList;
Begin
  --�������
  j_Jsonin := PLJson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  v_Vals   := j_Json.Get_Clob('pivas_ids');

  l_Vals := t_StrList();
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

  For Lp In 1 .. l_Vals.Count Loop
    v_��Һids := l_Vals(Lp);
    For R In (Select a.��ҩid, a.��������, a.������Ա, To_Char(a.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��, a.����˵��
              From ��Һ��ҩ״̬ A
              Where a.��ҩid In (Select /*+cardinality(x,10)*/
                                x.Column_Value
                               From Table(f_Num2List(v_��Һids)) X)) Loop
    
      v_Jtmp := v_Jtmp || ',{"pivas_id":' || r.��ҩid;
      v_Jtmp := v_Jtmp || ',"oper_type":' || r.��������;
      v_Jtmp := v_Jtmp || ',"operator_name":"' || zlJsonStr(r.������Ա) || '"';
      v_Jtmp := v_Jtmp || ',"operator_time":"' || r.����ʱ�� || '"';
      v_Jtmp := v_Jtmp || ',"operator_notes":"' || zlJsonStr(r.����˵��) || '"';
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
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Pivassvr_Getstatus_Info;
/

CREATE OR REPLACE Procedure Zl_Pivassvr_Pivascanstop
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ������Һ��Һ��¼�Ƿ�����ֹͣ
  --��Σ�Json_In:��ʽ
  --  input
  --     order_id    N  1   ��Һid
  --     stop_time   C  1   ͣ��ʱ��

  --����: Json_Out,��ʽ����

  --  output
  --    code            N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message         C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    can_stop_time   C 1 ����ֹͣ��ʱ��
  --    tip_time        C 1 ʵ��ִ��ʱ��
  --    is_package      N 1 �Ƿ���

  ---------------------------------------------------------------------------

  j_Json         PLJson;
  j_Json_Tmp     PLJson;
  j_Item         PLJson;
  n_ҽ��id       Number(18);
  n_���         Number(2);
  d_Stoptime     Date;
  v_����ʱ��     Varchar2(30);
  v_��ʾִ��ʱ�� Varchar2(30);

Begin
  --�������
  j_Item     := PLJson(Json_In);
  j_Json     := j_Item.Get_Pljson('input');
  n_ҽ��id   := j_Json.Get_Number('order_id');
  d_Stoptime := To_Date(j_Json.Get_String('stop_time'), 'YYYY-MM-DD HH24:MI:SS');

  For R In (Select Min(Decode(Instr(',4,5,6,7,8,', ',' || b.�������� || ','), 0, Null, To_Char(a.ִ��ʱ��, 'yyyy-MM-dd HH24:MI'))) As ����ִ��ʱ��,
                   Min(Decode(a.����״̬, 1, Null, To_Char(a.ִ��ʱ��, 'yyyy-MM-dd HH24:MI'))) As ��ʾִ��ʱ��, Min(a.�Ƿ���) As ���
            From ��Һ��ҩ��¼ A, ��Һ��ҩ״̬ B
            Where a.ҽ��id = n_ҽ��id And a.Id = b.��ҩid And a.ִ��ʱ�� > d_Stoptime And a.����״̬ <> 10) Loop
    v_����ʱ��     := r.����ִ��ʱ��;
    v_��ʾִ��ʱ�� := r.��ʾִ��ʱ��;
    n_���         := r.���;
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�"';
  Json_Out := Json_Out || ',"can_stop_time":"' || v_����ʱ�� || '"';
  Json_Out := Json_Out || ',"tip_time":"' || v_��ʾִ��ʱ�� || '"';
  Json_Out := Json_Out || ',"is_package":' || Nvl(n_���, 0);
  Json_Out := Json_Out || '}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Pivassvr_Pivascanstop;
/

Create Or Replace Procedure Zl_Pivassvr_Isselfdrug
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡ��Һ�Ա�ҩ�嵥
  --��Σ�Json_In:��ʽ
  --  input
  --    drug_ids    C   1   ҩƷID�������Ӣ�ĵĶ��ŷָ�
  --����: Json_Out,��ʽ����
  --  output
  --    code          N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    drug_ids      C 1 ҩƷID�������Ӣ�ĵĶ��ŷָ�
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  v_Drugids Varchar2(32767);

  v_Tmp Varchar2(32767);
Begin
  --�������
  j_Jsonin  := PLJson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  v_Drugids := j_Json.Get_String('drug_ids');

  For R In (Select a.ҩƷid
            From ��Һ�Ա�ҩ�嵥 A
            Where a.ҩƷid In (Select /*+cardinality(x,10)*/
                              x.Column_Value
                             From Table(Cast(f_Num2List(v_Drugids) As Zltools.t_Numlist)) X)) Loop
    v_Tmp := v_Tmp || ',' || r.ҩƷid;
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","drug_ids":"' || Substr(v_Tmp, 2) || '"}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Pivassvr_Isselfdrug;
/

Create Or Replace Procedure Zl_Pivassvr_Updatebatch
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  --------------------------------------------------------------------------------------------------------------------
  --���ܣ���Һҽ�����ε���
  --��Σ�Json_In,��ʽ����
  --  input
  --    item_list
  --         order_id                   N   1   ҽ��id,ҩƷҽ������ҽ��id
  --         operator_time              C   1   ����ʱ�䣬����ҽ��״̬�в�������=8�ļ�¼��������ʱ��
  --����: Json_Out,��ʽ����
  --  output
  --    code                            N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                         C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --------------------------------------------------------------------------------------------------------------------
  n_Count    Number(8);
  n_ҽ��id   Varchar2(4000);
  d_����ʱ�� Date;
  v_����     Varchar2(20);

  j_Json     Pljson; 
  Jl_Item    Pljson_List;
  j_Item     Pljson;

  Cursor c_��Һ��¼ Is
    Select Distinct a.Id ��ҩid, a.ִ��ʱ�� ʱ��, a.����id
    From ��Һ��ҩ��¼ A
    Where a.ҽ��id = n_ҽ��id And a.����״̬ = 1 And Trunc(d_����ʱ��) = Trunc(a.ִ��ʱ��) And a.ִ��ʱ�� < d_����ʱ��;

  v_��Һ��¼ c_��Һ��¼%RowType;

  Function Zl_Getpivaworkbatch
  (
    ִ��ʱ��_In   In Date,
    ��������id_In In ��Һ��ҩ��¼.����id%Type
  ) Return Number As
    v_Exetime   Date;
    v_Starttime Date;
    v_Endtime   Date;
    v_Maxbatch  Number(2);
    v_Batch     Number;
    Cursor c_��ҩ���� Is
      Select ����, ��ҩʱ��, ��ҩʱ�� From ��ҩ�������� Where ���� = 1 And ��������id = ��������id_In Order By ����;
  
    v_��ҩ���� c_��ҩ����%RowType;
  Begin
    v_Exetime := To_Date(Substr(To_Char(ִ��ʱ��_In, 'yyyy-mm-dd hh24:mi'), 12), 'hh24:mi');
  
    Select Max(Nvl(����, 0)) + 1 Into v_Maxbatch From ��ҩ�������� Where ���� = 1 And ��������id = ��������id_In;
  
    For v_��ҩ���� In c_��ҩ���� Loop
      v_Batch     := 0;
      v_Starttime := To_Date(Substr(v_��ҩ����.��ҩʱ��, 1, Instr(v_��ҩ����.��ҩʱ��, '-') - 1), 'hh24:mi');
      v_Endtime   := To_Date(Substr(v_��ҩ����.��ҩʱ��, Instr(v_��ҩ����.��ҩʱ��, '-') + 1), 'hh24:mi');
    
      If v_Exetime >= v_Starttime And v_Exetime <= v_Endtime Then
        v_Batch := v_��ҩ����.����;
      
        Exit When v_Batch > 0;
      End If;
    End Loop;
  
    If v_Batch = 0 Then
      v_Batch := v_Maxbatch;
    End If;
    Return(v_Batch);
  End;
Begin
  --�������
  j_Item := Pljson(Json_In);
  j_Json := j_Item.Get_Pljson('input');
  Jl_Item := j_Json.Get_Pljson_List('item_list');
  
  For I In 1 .. Jl_Item.Count Loop
    j_Item := Pljson();
    j_Item := Pljson(Jl_Item.Get(I));
  
    n_ҽ��id   := j_Item.Get_Number('order_id');
    d_����ʱ�� := To_Date(j_Item.Get_String('operator_time'), 'YYYY-MM-DD HH24:MI:SS');
  
    Select Count(a.Id)
    Into n_Count
    From ��Һ��ҩ��¼ A
    Where a.ҽ��id = n_ҽ��id And a.����״̬ = 1 And a.ִ��ʱ�� > d_����ʱ��;
  
    If n_Count > 0 Then
      For v_��Һ��¼ In c_��Һ��¼ Loop
        v_���� := Zl_Getpivaworkbatch(v_��Һ��¼.ʱ��, v_��Һ��¼.����id);
        Update ��Һ��ҩ��¼ Set ��ҩ���� = v_����, �Ƿ�������� = 1 Where ID = v_��Һ��¼.��ҩid;
      End Loop;
    End If;
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Pivassvr_Updatebatch;
/

Create Or Replace Procedure Zl_Pivassvr_Getpivascontent
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  -----------------------------------------------------------------------------
  --����:��ȡ��Һ��ҩ����
  --��Σ�Json_In:��ʽ
  --  input
  --    pivas_id                    N      1   ��Һid
  --    order_id                    N      0   ҽ��ID
  --    advice_endtime              C      0   ҽ����ִ����ֹʱ��
  --    auto_aduit                  N      0   �Ƿ�����Զ���ˣ�0-�����Զ���ˣ�1-�����Զ���ˣ���Ҫ���������Ƿ��Ѿ���ҩ
  --    query_type                  N      1   ��ѯ��ʽ
  --                                           0-����pivas_id���������в�ѯ,Ӧ�ó�����ʿ����վ����Һҽ����������
  --                                           1-����pivas_id+order_id+advice_endtime���в�ѯ��ȡ��ҩ��ϸ��Ϣ
  --                                           2-ȡ����������ʱ��ȡ����id��ϸ
  --����: Json_Out,��ʽ����
  --  output
  --    code                        N      1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                     C      1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    fee_ids                     C      0   ������ϸid��query_type=2ʱ�д˽��
  --    oper_time                   C      0   ����ʱ�䣬���������ʱ��query_type=2ʱ�д˽��
  --    item_list[]��Һ��¼��ϸ
  --       pivas_id                 N      1   ��Һid
  --       rcp_no                   C      1   ���ݺ�
  --       rcpdtl_id                N      1   ҩƷ�շ���¼��ϸid,������ϸid/����id
  --       send_num                 N      1   ��ҩƷ����
  --       drug_id                  N      1   ҩƷid
  --       si_drug_packg_qunt       N      1   סԺ��װ,��װ����
  --       si_drug_packg_unit       C      1   ��װ��λ
  --       drug_name                C      1   ҩƷ����
  --       is_sended                N      1   �Ƿ�ҩ
  --       drugstore_id             N      1   ��Һ��ҩ��¼�еĲ��ţ����غ�������ʱ����˲���
  --       status                   N      1   ����״̬  ����Һ��ҩ��¼��
  --       order_id                 N      0   ҽ��id
  --       send_no                  N      0   ҽ�����ͺ�
  -----------------------------------------------------------------------------
  n_��Һid     Number(18);
  n_ҽ��id     Number(18);
  d_��ֹʱ��   Date;
  n_Query_Type Number;
  v_Fee_Ids    Varchar2(32767);
  j_Json       Pljson;
  j_Json_Tmp   Pljson;
  n_Auto_Aduit Number;
  v_Jtmp       Varchar2(32767);
  c_Jtmp       Clob;

  Cursor c_Pivas_Ex Is
    Select b.����id, b.ҩƷid As �շ�ϸĿid, Sum(a.����) As ����, c.סԺ��װ, c.סԺ��λ, d.����, b.No, e.ҽ��id, e.���ͺ�
    From ��Һ��ҩ���� A, ҩƷ�շ���¼ B, ҩƷ��� C, �շ���ĿĿ¼ D, ��Һ��ҩ��¼ E
    Where a.��¼id = n_��Һid And a.�շ�id = b.Id And b.ҩƷid = c.ҩƷid And c.ҩƷid = d.Id And a.��¼id = e.Id And
          Instr(',8,9,10,21,24,25,26,', ',' || b.���� || ',') > 0 And
          (n_Auto_Aduit = 1 And b.����� Is Null Or n_Auto_Aduit = 0 And b.����� Is Not Null)
    Group By b.����id, b.ҩƷid, c.סԺ��װ, c.סԺ��λ, d.����, b.No, e.ҽ��id, e.���ͺ�;

  Cursor c_Pivas2
  (
    ���id_In       Number,
    ִ����ֹʱ��_In Date,
    ��ҩid_In       Number
  ) Is
    Select a.��¼id, a.�շ�id, Sum(a.����) As ����, b.����id, b.ҩƷid, 0 As �Ƿ�ҩ, c.סԺ��װ, c.סԺ��λ, d.����, b.No, e.���˲���id As ����id,
           e.����״̬
    From ��Һ��ҩ���� A, ҩƷ�շ���¼ B, ҩƷ��� C, �շ���ĿĿ¼ D, ��Һ��ҩ��¼ E
    Where a.�շ�id = b.Id And b.ҩƷid = c.ҩƷid And c.ҩƷid = d.Id And e.Id = a.��¼id And e.ҽ��id = ���id_In And
          e.ִ��ʱ�� > ִ����ֹʱ��_In And e.Id = ��ҩid_In
    Group By b.����id, b.ҩƷid, c.סԺ��װ, c.סԺ��λ, d.����, e.���˲���id, e.����״̬, e.Id, a.�շ�id, a.��¼id, e.���˲���id, d.����, b.No, e.����״̬;

  Type t_Pivas1 Is Table Of c_Pivas2%RowType;
  r_p t_Pivas1; --�α���丳ֵ�󣬿����޸������е�ֵ

Begin
  --�������
  j_Json_Tmp   := Pljson(Json_In);
  j_Json       := j_Json_Tmp.Get_Pljson('input');
  n_Query_Type := j_Json.Get_Number('query_type');
  n_��Һid     := j_Json.Get_Number('pivas_id');
  n_Auto_Aduit := j_Json.Get_Number('auto_aduit');
  n_Auto_Aduit := Nvl(n_Auto_Aduit, 0);

  If Nvl(n_Query_Type, 0) = 0 Then
    For R In c_Pivas_Ex Loop
      v_Jtmp := v_Jtmp || ',{"rcp_no":"' || r.No || '"';
      v_Jtmp := v_Jtmp || ',"rcpdtl_id":' || r.����id;
      v_Jtmp := v_Jtmp || ',"send_num":' || Zljsonstr(r.����, 1);
      v_Jtmp := v_Jtmp || ',"drug_id":' || r.�շ�ϸĿid;
      v_Jtmp := v_Jtmp || ',"si_drug_packg_qunt":' || Zljsonstr(r.סԺ��װ, 1);
      v_Jtmp := v_Jtmp || ',"si_drug_packg_unit":"' || Zljsonstr(r.סԺ��λ) || '"';
      v_Jtmp := v_Jtmp || ',"drug_name":"' || Zljsonstr(r.����) || '"';
      v_Jtmp := v_Jtmp || ',"order_id":' || r.ҽ��id;
      v_Jtmp := v_Jtmp || ',"send_no":' || r.���ͺ�;
      v_Jtmp := v_Jtmp || '}';
      If Length(v_Jtmp) > 30000 Then
        If c_Jtmp Is Null Then
          c_Jtmp := Substr(v_Jtmp, 2);
        Else
          c_Jtmp := c_Jtmp || v_Jtmp;
        End If;
        v_Jtmp := Null;
      End If;
      If Instr(',' || v_Fee_Ids || ',', ',' || r.����id || ',') = 0 Then
        v_Fee_Ids := v_Fee_Ids || ',' || r.����id;
      End If;
    End Loop;
    v_Fee_Ids := Substr(v_Fee_Ids, 2);
  
    If c_Jtmp Is Null Then
      Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_ids":"' || v_Fee_Ids || '","item_list":[' ||
                  Substr(v_Jtmp, 2) || ']}}';
    Else
      c_Jtmp   := c_Jtmp || v_Jtmp;
      Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_ids":"' || v_Fee_Ids || '","item_list":[' || c_Jtmp || ']}}';
    End If;
  
  Elsif n_Query_Type = 1 Then
    n_ҽ��id   := j_Json.Get_Number('order_id');
    d_��ֹʱ�� := To_Date(j_Json.Get_String('advice_endtime'), 'yyyy-mm-dd hh24:mi:ss');
    Open c_Pivas2(n_ҽ��id, d_��ֹʱ��, n_��Һid);
    Fetch c_Pivas2 Bulk Collect
      Into r_p;
    Close c_Pivas2;
  
    For I In 1 .. r_p.Count Loop
      v_Jtmp := v_Jtmp || ',{"pivas_id":' || r_p(I).��¼id;
      v_Jtmp := v_Jtmp || ',"rcp_no":"' || r_p(I).No || '"';
      v_Jtmp := v_Jtmp || ',"rcpdtl_id":' || r_p(I).����id;
      v_Jtmp := v_Jtmp || ',"send_num":' || Zljsonstr(r_p(I).����, 1);
      v_Jtmp := v_Jtmp || ',"drug_id":' || r_p(I).ҩƷid;
      v_Jtmp := v_Jtmp || ',"si_drug_packg_qunt":' || Zljsonstr(r_p(I).סԺ��װ, 1);
      v_Jtmp := v_Jtmp || ',"si_drug_packg_unit":"' || Zljsonstr(r_p(I).סԺ��λ) || '"';
      v_Jtmp := v_Jtmp || ',"drug_name":"' || Zljsonstr(r_p(I).����) || '"';
      v_Jtmp := v_Jtmp || ',"is_sended":' || Nvl(r_p(I).�Ƿ�ҩ, 0);
      v_Jtmp := v_Jtmp || ',"drugstore_id":' || Nvl(r_p(I).����id || '', 'null');
      v_Jtmp := v_Jtmp || ',"status":' || Nvl(r_p(I).����״̬ || '', 'null');
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
    If c_Jtmp Is Null Then
      Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
    Else
      c_Jtmp   := c_Jtmp || v_Jtmp;
      Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || c_Jtmp || ']}}';
    End If;
  Elsif n_Query_Type = 2 Then
    v_Fee_Ids  := Null;
    d_��ֹʱ�� := Null;
    For R In (Select Distinct b.����id From ��Һ��ҩ���� A, ҩƷ�շ���¼ B Where a.�շ�id = b.Id And a.��¼id = n_��Һid) Loop
      v_Fee_Ids := v_Fee_Ids || ',' || r.����id;
    End Loop;
    --�����²���ʱ�䣬������ʱ��ɾ�������¼��׼ȷЩ
    For R In (Select a.����ʱ��
              From (Select ����ʱ��
                     From ��Һ��ҩ״̬
                     Where ��ҩid = n_��Һid And �������� = 9
                     Order By ����ʱ�� Desc, �������� Desc) A) Loop
      d_��ֹʱ�� := r.����ʱ��;
      Exit;
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_ids":"' || Substr(v_Fee_Ids, 2) || '","oper_time":"' ||
                To_Char(d_��ֹʱ��, 'yyyy-mm-dd hh24:mi:ss') || '"}}';
  End If;
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Pivassvr_Getpivascontent;
/

Create Or Replace Procedure Zl_Pivassvr_Statusupdate
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------------
  --���ܣ���Һ��������״̬����
  --���
  -- input
  --   pivas_ids                               C   1   ��ҩ��¼id����ƴ����ʽ,Щ��㲻�����б�ֵ���и���
  --   operator_status                         N   1   ����״̬, ������״̬Ϊ -1 ʱ��ʾΪȡ����������
  --   operator_name                           C   1   ����Ա����
  --   operator_notes                          C   0   ����˵��
  --   operator_time                           C   1   ����ʱ��
  --   auto_aduit                              N   0   �Զ�����������룬�Զ����״̬Ϊ10
  --   item_list[]���б�����Բ�������������б�ʽ����
  --      pivas_id                                N   1   ��ҩ��¼id
  --      operator_status                         N   1   ����״̬, ������״̬Ϊ -1 ʱ��ʾΪȡ����������
  --      operator_name                           C   1   ����Ա����
  --      operator_notes                          C   0   ����˵��
  --      operator_time                           C   1   ����ʱ��
  --����
  -- output
  --   code                                    N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --   message                                 C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------------

  j_Json       Pljson;
  j_List       Pljson_List;
  j_Tmp        Pljson;
  v_��Һids    Varchar2(32767);
  v_����Ա���� Varchar2(40);
  d_����ʱ��   Date;
  d_���ʱ��   Date;
  n_����״̬   Number(2);
  v_����˵��   Varchar2(4000);
  n_��Һid     Number(18);
  n_��������   Number(2);
  n_Auto_Aduit Number;
  v_Vals       Clob;
  l_Vals       t_Strlist;
Begin
  --�������
  j_Tmp  := Pljson(Json_In);
  j_Json := j_Tmp.Get_Pljson('input');
  v_Vals := j_Json.Get_Clob('pivas_ids');
  If v_Vals Is Not Null Then
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
  
    n_����״̬   := j_Json.Get_Number('operator_status');
    v_����Ա���� := j_Json.Get_String('operator_name');
    d_����ʱ��   := To_Date(j_Json.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');
    v_����˵��   := j_Json.Get_String('operator_notes');
    n_Auto_Aduit := j_Json.Get_Number('auto_aduit');
  
    If Nvl(n_����״̬, 0) <> -1 Then
      If n_Auto_Aduit = 1 Then
        d_����ʱ�� := d_����ʱ�� + 1 / 24 / 60 / 60;
      End If;
      For Lp In 1 .. l_Vals.Count Loop
        v_��Һids := l_Vals(Lp);
        Insert Into ��Һ��ҩ״̬
          (��ҩid, ��������, ������Ա, ����ʱ��, ����˵��)
          Select /*+cardinality(y,10)*/
           To_Number(y.Column_Value) As ��ҩid, n_����״̬, v_����Ա����, d_����ʱ��, v_����˵��
          From Table(f_Num2list(v_��Һids)) Y;
      
        Update ��Һ��ҩ��¼
        Set ������Ա = v_����Ա����, ����ʱ�� = d_����ʱ��, ����״̬ = n_����״̬
        Where ID In (Select /*+cardinality(y,10)*/
                      y.Column_Value
                     From Table(f_Num2list(v_��Һids)) Y);
      
        If n_Auto_Aduit = 1 Then
          --���������������״̬��ʱ�����ֿ�����һ��
          Insert Into ��Һ��ҩ״̬
            (��ҩid, ��������, ������Ա, ����ʱ��, ����˵��)
            Select /*+cardinality(y,10)*/
             To_Number(y.Column_Value) As ��ҩid, 10, v_����Ա����, d_���ʱ��, v_����˵��
            From Table(f_Num2list(v_��Һids)) Y;
        
          Update ��Һ��ҩ��¼
          Set ������Ա = v_����Ա����, ����ʱ�� = d_���ʱ��, ����״̬ = 10
          Where ID In (Select /*+cardinality(y,10)*/
                        y.Column_Value
                       From Table(f_Num2list(v_��Һids)) Y);
        
        End If;
      End Loop;
    Else
      --�ر�ע�⣬ȡ������ֻ����һ��һ����ȡ��
      v_��Һids := l_Vals(1);
      n_��Һid  := To_Number(v_��Һids);
      Select ������Ա, ����ʱ��, ��������
      Into v_����Ա����, d_����ʱ��, n_��������
      From (Select ������Ա, ����ʱ��, ��������
             From ��Һ��ҩ״̬
             Where ��ҩid = n_��Һid And �������� <> 9
             Order By ����ʱ�� Desc, �������� Desc)
      Where Rownum = 1;
    
      Update ��Һ��ҩ��¼
      Set ������Ա = v_����Ա����, ����ʱ�� = d_����ʱ��, ����״̬ = n_��������
      Where ID = n_��Һid;
    
    End If;
  Else
    j_List := j_Json.Get_Pljson_List('item_list');
    For I In 1 .. j_List.Count Loop
      j_Tmp        := Pljson();
      j_Tmp        := Pljson(j_List.Get(I));
      n_��Һid     := j_Tmp.Get_Number('pivas_id');
      n_����״̬   := j_Tmp.Get_Number('operator_status');
      v_����Ա���� := j_Tmp.Get_String('operator_name');
      d_����ʱ��   := To_Date(j_Tmp.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');
      v_����˵��   := j_Tmp.Get_String('operator_notes');
    
      Insert Into ��Һ��ҩ״̬
        (��ҩid, ��������, ������Ա, ����ʱ��, ����˵��)
      Values
        (n_��Һid, n_����״̬, v_����Ա����, d_����ʱ��, v_����˵��);
    
      Update ��Һ��ҩ��¼
      Set ������Ա = v_����Ա����, ����ʱ�� = d_����ʱ��, ����״̬ = n_����״̬
      Where ID = n_��Һid;
    
    End Loop;
  End If;
  Json_Out := '{"output":{"code":1,"message": "�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Pivassvr_Statusupdate;
/

Create Or Replace Procedure Zl_Pivassvr_Patiinfoupdate
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ�סԺ������Һ��ҩ��¼������Ϣ�޸�
  --��Σ�Json_In:��ʽ
  --  input
  --   pati_id       N   1   ����id
  --   pati_name      C   1   ��������
  --   pati_sex       C   1   �����Ա�
  --   pati_age       C   1   ��������
  --   visit_id      N   1   ����id

  --����: Json_Out,��ʽ����
  --  output
  --    code                 N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message              C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------

  j_Json Pljson;

  v_����     Varchar2(100);
  v_�Ա�     Varchar2(100);
  v_����     Varchar2(100);
  n_����id   Number;
  n_����id   Number;
  j_Json_Tmp Pljson;
Begin
  --�������
  j_Json_Tmp := Pljson(Json_In);
  j_Json     := j_Json_Tmp.Get_Pljson('input');
  v_����     := j_Json.Get_String('pati_name');
  v_�Ա�     := j_Json.Get_String('pati_sex');
  v_����     := j_Json.Get_String('pati_age');
  n_����id   := j_Json.Get_Number('visit_id');
  n_����id   := j_Json.Get_Number('pati_id');
  Update ��Һ��ҩ��¼
  Set ���� = Nvl(v_����, ����), �Ա� = Nvl(v_�Ա�, �Ա�), ���� = Nvl(v_����, ����)
  Where ����id = n_����id And ��ҳid = n_����id;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Pivassvr_Patiinfoupdate;
/

Create Or Replace Procedure Zl_Pivassvr_Getworkbatch
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------------------------------------------------------------
  --���ܣ���ȡ��ҩ��������
  --��Σ�json
  --      ���Բ����룬��ʱΪ��
  --���Σ�json
  --output
  --  code                      N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message                   C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  item_list
  --        pivas_deptid        N 1     ��������id
  --        batch               N 1     ����
  --        pivas_time          C 1     ��ҩʱ��
  -------------------------------------------------------------------------------------------------
  v_Jtmp Varchar2(32767);
Begin
  For R In (Select ��������id, ����, ��ҩʱ�� From ��ҩ�������� Order By ����) Loop
    v_Jtmp := v_Jtmp || ',{"pivas_deptid":' || r.��������id;
    v_Jtmp := v_Jtmp || ',"batch":' || r.����;
    v_Jtmp := v_Jtmp || ',"pivas_time":"' || r.��ҩʱ�� || '"';
    v_Jtmp := v_Jtmp || '}';
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Pivassvr_Getworkbatch;
/

Create Or Replace Procedure Zl_Pivassvr_Checkordersendroll
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ�ҽ�����˷���ʱ������ҽ������ж�
  --��Σ�Json_In:��ʽ
  --  input
  --     order_id           N 1 ҽ��ID,��ҽ��id
  --     send_no            N 1 ���ͺ�
  --����: Json_Out,��ʽ����
  --  output
  --    code                N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    pivas_ids           C 1 Ҫ���ʵ���Һ��¼id��
  ---------------------------------------------------------------------------
  ҽ��id_In Number;
  n_���ͺ�  Number;
  j_Json    PLJson;
  j_Tmp     PLJson;
  n_Tmp     Number;
  v_��Һids Varchar2(32767);
Begin
  --�������
  j_Tmp     := PLJson(Json_In);
  j_Json    := j_Tmp.Get_Pljson('input');
  ҽ��id_In := j_Json.Get_Number('order_id');
  n_���ͺ�  := j_Json.Get_Number('send_no');

  --����Ƿ�����Һ��Һ��¼�����Ƿ��Ѿ�����
  Select Decode(Max(�Ƿ�����), 1, 1, 0) Into n_Tmp From ��Һ��ҩ��¼ Where ҽ��id = ҽ��id_In And ���ͺ� = n_���ͺ�;
  If n_Tmp = 1 Then
    Json_Out := zlJsonOut('��ǰ���������ҺҩƷҽ�����Ѿ�����Һ�����������������ܻ��˷��͡�');
    Return;
  Elsif n_Tmp = 0 Then
    --Zl_��Һ��ҩ��¼_ҽ������(ҽ��id_In, n_���ͺ�, Null, Null);
    --ֻ��״̬=1(δ��ҩ)�ļ�¼��������Ѿ���ҩ�ˣ���ͨ�����˷�ʽ����
    Select Count(ID)
    Into n_Tmp
    From ��Һ��ҩ��¼
    Where ����״̬ In (1, 10) And ҽ��id = ҽ��id_In And ���ͺ� = n_���ͺ�;
    If n_Tmp > 0 Then
      For R In (Select ID From ��Һ��ҩ��¼ Where ҽ��id = ҽ��id_In And ���ͺ� = n_���ͺ�) Loop
        v_��Һids := v_��Һids || ',' || r.Id;
      End Loop;
    End If;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","pivas_ids":"' || Substr(v_��Һids, 2) || '"}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Pivassvr_Checkordersendroll;
/

Create Or Replace Procedure Zl_Pivassvr_Odr_Check
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ڷ����ջؾ������ȡ���ͼ��
  --��Σ�Json_In:��ʽ
  --  input
  --     chk_type                                  N 1 ��鷽ʽ��0-��ȡ�б�1-��Һ������ҩ��ʾ
  --     item_list[]Ҫ�ջص�ҽ���б�
  --               order_id                        N 1 ҽ��ID,��ҽ��id
  --               exe_end_time                    C 1 ִ����ֹʱ��
  --               advice_note                     C 0 ҽ�����ݣ���鷽ʽΪ1ʱ�˽����Բ���
  --����: Json_Out,��ʽ����
  --  output
  --    code                          N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    isexist                       N 1 �Ƿ�����Ѿ���Һ�ļ�¼��1-����;0-������
  --    pivas_list[]�б�
  --               pivas_id           N 1 ��Һid
  --               order_id           N 1 ҽ��id
  --               fee_id             N 1 ����id
  --               fee_item_id        N 1 �շ�ϸĿid
  --               operator_status    N 1 ����״̬
  --               quantity           N 1 ����
  ---------------------------------------------------------------------------
  j_Input          Pljson;
  j_Item           Pljson;
  j_List           Pljson_List := Pljson_List();
  n_���id         Number(18); --��ҽ��ID
  d_ִ����ֹʱ��   Date;
  v_ҽ������       Varchar2(30000);
  n_Count          Number;
  v_����Һ��¼     Varchar2(1000);
  v_��Һҩ�������� Varchar2(4000);
  v_Charge_List    Varchar2(32767);
  v_Json_Out       Varchar2(32767);
  v_ҽ��ids        Varchar2(32767);
  n_��鷽ʽ       Number;
  v_Error          Varchar2(255);
  Err_Custom Exception;
Begin
  j_Item     := Pljson(Json_In);
  j_Input    := j_Item.Get_Pljson('input');
  n_��鷽ʽ := j_Input.Get_Number('chk_type');
  j_List     := j_Input.Get_Pljson_List('item_list');
  If j_List Is Not Null Then
    For I In 1 .. j_List.Count Loop
      j_Item := Pljson();
      j_Item := Pljson(j_List.Get(I));
    
      n_���id       := j_Item.Get_Number('order_id');
      d_ִ����ֹʱ�� := To_Date(j_Item.Get_String('exe_end_time'), 'yyyy-mm-dd hh24:mi:ss');
      If d_ִ����ֹʱ�� Is Null Then
        d_ִ����ֹʱ�� := To_Date('1900-01-01', 'yyyy-mm-dd');
      End If;
    
      If n_��鷽ʽ = 1 Then
        --�ж��Ƿ��Ѿ���Һ�ļ�¼�������л��и�ѯ��ʾ�Ľ�������
        Select Count(1)
        Into n_Count
        From ��Һ��ҩ��¼ B
        Where (b.����״̬ In (4, 5, 6, 7, 8) And Nvl(b.�Ƿ���, 0) = 0) And b.ҽ��id = n_���id And b.ִ��ʱ�� > d_ִ����ֹʱ��;
        If n_Count > 0 Then
          v_����Һ��¼ := ',"isexist":1';
          Exit;
        End If;
      Else
        v_��Һҩ�������� := zl_GetSysParameter('��Һ��Һ����ҩ��������������', 1345);
        v_ҽ������       := j_Item.Get_String('advice_note');
        --����Ƿ�����Һ��Һ��¼�����Ƿ��Ѿ�����
        Select Count(1)
        Into n_Count
        From ��Һ��ҩ��¼ A
        Where a.ҽ��id = n_���id And a.ִ��ʱ�� > d_ִ����ֹʱ�� And a.�Ƿ����� = 1;
      
        If n_Count > 0 Then
          v_Error := 'ҽ��"' || v_ҽ������ || '"����ҺҩƷ���Ѿ�����Һ�����������������ܳ����ջء�';
          Raise Err_Custom;
        End If;
      
        For R In (Select b.����id, b.ҩƷid As �շ�ϸĿid, c.����״̬, b.ҽ��id, c.Id ��ҩid, Sum(a.����) As ����
                  From ��Һ��ҩ���� A, ҩƷ�շ���¼ B, ��Һ��ҩ��¼ C
                  Where a.�շ�id = b.Id And a.��¼id = c.Id And c.ҽ��id = n_���id And c.ִ��ʱ�� > d_ִ����ֹʱ�� And
                        Nvl(c.����״̬, 0) In (1, 2, 3, 4, 5, 6, 7, 8) And
                        Not (c.����״̬ In (4, 5, 6, 7, 8) And Nvl(c.�Ƿ���, 0) = 0 And Nvl(v_��Һҩ��������, '0') = '0')
                  Group By b.����id, b.ҩƷid, b.ҽ��id, c.����״̬, b.ҽ��id, c.Id, c.���ͺ�, c.ִ��ʱ��
                  Order By c.���ͺ�, b.ҽ��id, c.ִ��ʱ��) Loop
        
          v_Charge_List := v_Charge_List || ',{"pivas_id":' || r.��ҩid;
          v_Charge_List := v_Charge_List || ',"order_id":' || r.ҽ��id;
          v_Charge_List := v_Charge_List || ',"fee_id":' || r.����id;
          v_Charge_List := v_Charge_List || ',"fee_item_id":' || r.�շ�ϸĿid;
          v_Charge_List := v_Charge_List || ',"operator_status":' || r.����״̬;
          v_Charge_List := v_Charge_List || ',"quantity":' || Zljsonstr(r.����, 1); --N 1 ��������  
          v_Charge_List := v_Charge_List || '}';
        End Loop;
      End If;
    
    End Loop;
    If v_Charge_List Is Not Null Then
      v_Charge_List := ',"pivas_list":[' || Substr(v_Charge_List, 2) || ']';
    End If;
  End If;

  v_Json_Out := '{"code":1,"message":"�ɹ�"';
  v_Json_Out := v_Json_Out || v_����Һ��¼;
  v_Json_Out := v_Json_Out || v_Charge_List;
  v_Json_Out := v_Json_Out || '}';
  Json_Out   := '{"output":' || v_Json_Out || '}';

Exception
  When Err_Custom Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(v_Error) || '"}}';
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Pivassvr_Odr_Check;
/

CREATE OR REPLACE Procedure Zl_Pivassvr_Overdue_Recovery
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ڷ����վ������ش���
  --��Σ�Json_In:��ʽ
  --  input
  --     operator_name                        C 1 ����Ա����
  --     pivas_list[]���������б�
  --                  pivas_id               N 1 ��Һid
  --                  auto_aduit             N 1 �Ƿ��Զ���� 0-�����,1-Ҫ�Զ����
  --                  request_time           C 1 ����ʱ��
  --                  reason                 C 1 ����ԭ��
  --����: Json_Out,��ʽ����
  --  output
  --    code                          N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Input      Pljson;
  j_Item       Pljson;
  j_List       Pljson_List := Pljson_List();
  v_����Ա���� Varchar2(300);
  n_��Һid     Number;
  n_�Զ����   Number;
  d_����ʱ��   Date;
  v_����ԭ��   Varchar2(4000);
Begin
  j_Item  := Pljson(Json_In);
  j_Input := j_Item.Get_Pljson('input');

  v_����Ա���� := j_Input.Get_String('operator_name');

  j_List := j_Input.Get_Pljson_List('pivas_list');
  If j_List Is Not Null Then
    For I In 1 .. j_List.Count Loop
      j_Item     := Pljson();
      j_Item     := Pljson(j_List.Get(I));
      n_��Һid   := j_Item.Get_Number('pivas_id');
      n_�Զ���� := j_Item.Get_Number('auto_aduit');
      d_����ʱ�� := To_Date(j_Item.Get_String('request_time'), 'yyyy-mm-dd hh24:mi:ss');
      v_����ԭ�� := j_Item.Get_String('reason');
    
      Insert Into ��Һ��ҩ״̬
        (��ҩid, ��������, ������Ա, ����ʱ��, ����˵��)
      Values
        (n_��Һid, 9, v_����Ա����, d_����ʱ��, v_����ԭ��);
    
      Update ��Һ��ҩ��¼ Set ������Ա = v_����Ա����, ����ʱ�� = d_����ʱ��, ����״̬ = 9 Where ID = n_��Һid;
    
      If n_�Զ���� = 1 Then
        --��һ��,������ܺͷ���������������ʱ������һ��,Ӱ�첻��
        d_����ʱ�� := d_����ʱ�� + 1 / 24 / 60 / 60;
        Insert Into ��Һ��ҩ״̬
          (��ҩid, ��������, ������Ա, ����ʱ��,����˵��)
        Values
          (n_��Һid, 10, v_����Ա����, d_����ʱ��,v_����ԭ��);
        Update ��Һ��ҩ��¼ Set ������Ա = v_����Ա����, ����ʱ�� = d_����ʱ��, ����״̬ = 10 Where ID = n_��Һid;
      End If;
    End Loop;
  End If;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Pivassvr_Overdue_Recovery;
/

Create Or Replace Procedure Zl_Pivassvr_Getchargeerrdata
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ��������������Զ����ʱ��������ʧ�ܺ��޸�ʱ��ȡ��������
  --��Σ�Json_In:��ʽ
  --  input
  --     order_id                  N 1 ҽ��ID,��ҽ��id
  --     send_no                   N 1 ���ͺ�
  --     pivas_id                  N 1 ��Һ��¼id�����������Զ���˺�����쳣����Һ��¼id
  --     item_list[]
  --            rcpdtl_id          N 1 ������ϸid
  --            rcp_no             C 1 ��������
  --            drug_id            N 1 ҩƷid
  --            quantity           N 1 ҽ�����ͺ��������,������λ����������Դ��ҽ�����ͼ�¼������˵Ӧ����ҽ�����ͼ�¼������һ��ҩƷid�ֶ���Ϊ������Ʒ���´�ʱȡ����
  --            order_id           N 1 ҽ��ID
  --����: Json_Out,��ʽ����
  --  output
  --    code                N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    item_list[]
  --            rcpdtl_id          N 1 ������ϸid
  --            rcp_no             C 1 ��������
  --            drug_id            N 1 ҩƷid
  --            quantity           N 1 ����������������
  --            order_id           N 1 ҽ��ID  
  --            request_time       C 1 ����ʱ��
  --            audit_time         C 1 ���ʱ��
  --            reason             C 1 ����ԭ��
  --            request_operator   C 1 ���ʲ���Ա
  --            request_code       C 1 ���ʲ���Ա����
  ---------------------------------------------------------------------------
  n_ҽ��id     Number(18);
  n_���ͺ�     Number(18);
  n_��ҩid     Number(18);
  j_Json       Pljson;
  j_Tmp        Pljson;
  j_Item       Pljson;
  n_Tmp        Number;
  v_Jtmp       Varchar2(32767);
  n_List_Cnt   Number;
  j_Jsonlist   Pljson_List;
  v_����ʱ��   Varchar2(30);
  v_���ʱ��   Varchar2(30);
  n_����       Number(3);
  n_���һ��   Number(3);
  n_����       Number(16, 5);
  n_ҩƷid     Number(18);
  n_��������   Number(16, 5);
  v_����ԭ��   Varchar2(1000);
  v_������Ա   Varchar2(1000);
  v_����Ա��� Varchar2(1000);
Begin
  --�������
  j_Tmp      := Pljson(Json_In);
  j_Json     := j_Tmp.Get_Pljson('input');
  n_ҽ��id   := j_Json.Get_Number('order_id');
  n_���ͺ�   := j_Json.Get_Number('send_no');
  n_��ҩid   := j_Json.Get_Number('pivas_id');
  j_Jsonlist := j_Json.Get_Pljson_List('item_list');
  If Not j_Jsonlist Is Null Then
    n_List_Cnt := j_Jsonlist.Count;
  Else
    n_List_Cnt := 1;
  End If;

  For R In (Select To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') ����ʱ��, a.��������, a.����˵��, a.������Ա
            From ��Һ��ҩ״̬ A
            Where a.�������� In (9, 10) And a.��ҩid = n_��ҩid) Loop
    If r.�������� = 9 Then
      v_����ʱ�� := r.����ʱ��;
      v_����ԭ�� := r.����˵��;
      v_������Ա := r.������Ա;
      Select a.��� Into v_����Ա��� From ��Ա�� A Where a.���� = v_������Ա;
    Else
      v_���ʱ�� := r.����ʱ��;
    End If;
  End Loop;
  n_���� := 0;
  For R In (Select a.Id, a.ִ��ʱ��
            From ��Һ��ҩ��¼ A
            Where a.ҽ��id = n_ҽ��id And a.���ͺ� = n_���ͺ�
            Order By a.ִ��ʱ��) Loop
    n_���� := n_���� + 1;
    If r.Id = n_��ҩid Then
      n_���һ�� := n_����;
    End If;
  End Loop;

  If n_���һ�� = n_���� Then
    n_���һ�� := 1;
  Else
    n_���һ�� := 0;
  End If;

  For I In 1 .. n_List_Cnt Loop
    j_Item     := Pljson(j_Jsonlist.Get(I));
    n_����     := j_Item.Get_Number('quantity');
    n_ҩƷid   := j_Item.Get_Number('drug_id');
    n_�������� := n_���� / n_����;
    If n_���һ�� = 1 Then
      n_�������� := n_���� - (n_���� / n_����) * (n_���� - 1);
    End If;
    Select (n_�������� / a.����ϵ��) Into n_�������� From ҩƷ��� A Where a.ҩƷid = n_ҩƷid;
  
    v_Jtmp := v_Jtmp || ',{"rcpdtl_id":' || j_Item.Get_Number('rcpdtl_id');
    v_Jtmp := v_Jtmp || ',"rcp_no":"' || j_Item.Get_String('rcp_no') || '"';
    v_Jtmp := v_Jtmp || ',"drug_id":' || j_Item.Get_Number('drug_id');
    v_Jtmp := v_Jtmp || ',"quantity":' || Zljsonstr(n_��������, 1);
    v_Jtmp := v_Jtmp || ',"order_id":' || j_Item.Get_Number('order_id');
    v_Jtmp := v_Jtmp || ',"request_time":"' || v_����ʱ�� || '"';
    v_Jtmp := v_Jtmp || ',"audit_time":"' || v_���ʱ�� || '"';
    v_Jtmp := v_Jtmp || ',"reason":"' || Zljsonstr(v_����ԭ��) || '"';
    v_Jtmp := v_Jtmp || ',"request_operator":"' || Zljsonstr(v_������Ա) || '"';
    v_Jtmp := v_Jtmp || ',"request_code":"' || Zljsonstr(v_����Ա���) || '"';
    v_Jtmp := v_Jtmp || '}';
    j_Item := Pljson();
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Pivassvr_Getchargeerrdata;
/