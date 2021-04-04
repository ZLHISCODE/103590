Create Or Replace Procedure Zl_Drugsvr_Getstockshow
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡָ���ⷿ�Ŀ�����ݣ�������ʾ
  --��Σ�Json_In:��ʽ
  --  input
  --    pharmacy_ids        C   1   �ⷿID
  --����: Json_Out,��ʽ����
  --  output
  --    code                N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    item_list
  --      drug_id              N   1   ҩƷID
  --      pharmacy_id          N   1   �ⷿID
  --      stock                N   1   ��������
  --      real_stock          N  1 ʵ�ʿ��
  --      avg_price           N  1 ƽ���ۼ�
  ---------------------------------------------------------------------------

  v_�ⷿids Varchar2(32767);

  j_Jsonin Pljson;
  j_Json   Pljson;

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;
Begin
  --�������
  j_Jsonin  := Pljson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  v_�ⷿids := j_Json.Get_String('pharmacy_ids');

  If Nvl(v_�ⷿids, 0) = 0 Then
    Json_Out := Zljsonout('δ������ؿⷿ��Ϣ');
    Return;
  End If;

  For c_��� In (Select a.�ⷿid, a.ҩƷid, Nvl(Sum(a.��������), 0) As ��������, Nvl(Sum(a.ʵ������), 0) As ʵ������,
                      Decode(Nvl(Sum(a.ʵ������), 0), 0, Max(a.���ۼ�), Nvl(Sum(a.ʵ�ʽ��), 0) / Nvl(Sum(a.ʵ������), 0)) As ƽ���ۼ�
               From ҩƷ��� A
               Where a.���� = 1 And Instr(',' || v_�ⷿids || ',', ',' || a.�ⷿid || ',') > 0
               Group By a.�ⷿid, a.ҩƷid) Loop
  
    v_Jtmp := v_Jtmp || ',';
    Zljsonputvalue(v_Jtmp, 'drug_id', c_���.ҩƷid, 1, 1);
    Zljsonputvalue(v_Jtmp, 'pharmacy_id', c_���.�ⷿid, 1);
    Zljsonputvalue(v_Jtmp, 'stock', c_���.��������, 1);
    Zljsonputvalue(v_Jtmp, 'real_stock', c_���.ʵ������, 1);
    Zljsonputvalue(v_Jtmp, 'avg_price', c_���.ƽ���ۼ�, 1, 2);
  
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

Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getstockshow;
/

Create Or Replace Procedure Zl_Drugsvr_CheckDrugExistStock
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  ---------------------------------------------------------------------------
  --input      �ж�ҩƷ�Ƿ���ڿ�����
  --  drug_id      N  1  ҩƷid
  --  is_item      N  1  �Ƿ�Ʒ�ֲ�ѯ��0-������ѯ��1-��Ʒ�ֲ�ѯ
  --output      
  --  code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message  C  1  Ӧ����Ϣ��
  --  isexist  N 1 �Ƿ����: 1-����;0-������
  ---------------------------------------------------------------------------
  j_Jsonin Pljson;
  j_Json   Pljson;
  n_ҩƷid ҩƷ���.ҩƷid%Type;
  n_Ʒ��   Number(1);
  n_Exist  Number(1);
Begin
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_ҩƷid := j_Json.Get_Number('drug_id');
  n_Ʒ��   := j_Json.Get_Number('is_item');

  If n_Ʒ�� = 0 Then
    Select Count(1) Into n_Exist From ҩƷ��� Where ҩƷid = n_ҩƷid And Rownum < 2;
  Else
    Select Count(1)
    Into n_Exist
    From ҩƷ��� A, ҩƷ��� B
    Where a.ҩƷid = b.ҩƷid And a.ҩ��id = n_ҩƷid And Rownum < 2;
  End If;

  Json_Out := '{"output":{"code": 1,"message":"�ɹ�","isexist":' || n_Exist || '}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugsvr_CheckDrugExistStock;
/

Create Or Replace Procedure Zl_DrugSvr_GetCostPriceAdjust
(
  Json_In  Varchar2,
  Json_Out Out Clob
) Is
  ---------------------------------------------------------------------------
  --���ܣ���ȡҩƷ�ɱ��۵��ۼ�¼
  --input      
  --  drug_id      N   1 ҩƷid
  --  show_unit    N   1   ��ʾ��λ:0-�ۼ۵�λ;3-ҩ�ⵥλ
  --output      
  --  code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message  C  1  Ӧ����Ϣ��  
  --  price_list[]  ҩƷ�ɱ��۵��ۼ�¼
  --     drug_id   N 1  ҩƷID
  --     drug_name   C 1  ҩƷ��Ϣ
  --     stock_name   C 1  �ⷿ
  --     batch_number   C 1  ����
  --     effective_time   C 1  Ч��
  --     place_name   C 1  ����
  --     unit_name   C 1  ��λ
  --     cost_old   N 1  ԭ�ɱ���
  --     cost_new    N 1  �ֳɱ���
  --     adjust_time   C 1  ����ʱ��
  --     adjust_reson   C 1  ����˵��
  --     adjust_no   C 1  ���۵��ݺ�
  --     drug_revoke_time  C 1 ����ʱ��
  --     node_no      C    0  վ�����   
  --     is_stock    N   1 �Ƿ��п������  0-��1-��
  ---------------------------------------------------------------------------
  j_Jsonin Pljson;
  j_Json   Pljson;
  n_ҩƷid ҩƷ���.ҩƷid%Type;
  n_��λ   Number(1);

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;
Begin
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_ҩƷid := j_Json.Get_Number('drug_id');
  n_��λ   := j_Json.Get_Number('show_unit');

  v_Jtmp := Null;
  For r_Costprice In (Select Distinct b.No, i.Id As ҩƷid, '[' || i.���� || ']' || i.���� || ' ' || i.��� || ' ' || i.���� As ҩƷ,
                                      p.���� As �ⷿ, a.����, a.Ч��, a.����, Decode(n_��λ, 0, i.���㵥λ, s.ҩ�ⵥλ) As ��λ,
                                      Decode(n_��λ, 0, a.ԭ��, a.ԭ�� * Nvl(s.ҩ���װ, 1)) As ԭ�ɱ���,
                                      Decode(n_��λ, 0, a.�ּ�, a.�ּ� * Nvl(s.ҩ���װ, 1)) As �ɱ���, a.ִ������, a.����˵��, i.����ʱ��, i.վ��,
                                      Decode(k.�ⷿid, Null, 0, 1) As ���
                      From ҩƷ�շ���¼ B, �շ���ĿĿ¼ I, ҩƷ��� S, ���ű� P, ҩƷ�۸��¼ A, ҩƷ��� K
                      Where a.�۸����� = 2 And a.�շ�id = b.Id(+) And a.ҩƷid = i.Id And i.Id = s.ҩƷid And a.�ⷿid = p.Id(+) And
                            s.ҩ��id = n_ҩƷid And k.����(+) = 1 And k.�ⷿid(+) = a.�ⷿid And k.ҩƷid(+) = a.ҩƷid And
                            k.����(+) = a.����
                      Order By '[' || i.���� || ']' || i.���� || ' ' || i.��� || ' ' || i.����, p.����, a.����, a.ִ������ Desc, NO Desc) Loop
  
    v_Jtmp := v_Jtmp || ',';
    Zljsonputvalue(v_Jtmp, 'drug_id', r_Costprice.ҩƷid, 1, 1);
    Zljsonputvalue(v_Jtmp, 'drug_name', r_Costprice.ҩƷ, 0);
    Zljsonputvalue(v_Jtmp, 'stock_name', r_Costprice.�ⷿ, 0);
    Zljsonputvalue(v_Jtmp, 'batch_number', r_Costprice.����, 0);
    Zljsonputvalue(v_Jtmp, 'effective_time', r_Costprice.Ч��, 0);
  
    Zljsonputvalue(v_Jtmp, 'place_name', r_Costprice.����, 0);
    Zljsonputvalue(v_Jtmp, 'unit_name', r_Costprice.��λ, 0);
    Zljsonputvalue(v_Jtmp, 'cost_old', r_Costprice.ԭ�ɱ���, 1);
    Zljsonputvalue(v_Jtmp, 'cost_new', r_Costprice.�ɱ���, 1);
    Zljsonputvalue(v_Jtmp, 'adjust_time', r_Costprice.ִ������, 0);
  
    Zljsonputvalue(v_Jtmp, 'adjust_reson', r_Costprice.����˵��, 0);
    Zljsonputvalue(v_Jtmp, 'adjust_no', r_Costprice.No, 0);
    Zljsonputvalue(v_Jtmp, 'drug_revoke_time', r_Costprice.����ʱ��, 0);
    Zljsonputvalue(v_Jtmp, 'node_no', r_Costprice.վ��, 0);
    Zljsonputvalue(v_Jtmp, 'is_stock', r_Costprice.���, 1, 2);
  
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
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","price_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","price_list":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_DrugSvr_GetCostPriceAdjust;
/

Create Or Replace Procedure Zl_DrugSvr_AdjustPriceType
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --ִ�е���
  ---------------------------------------------------------------------------
  --input      ҩƷ�۸����Ե���ʱ�����ĵ���ӯ���Ϳ��仯���ݴ���
  --    drug_list[]
  --       drug_id      N    ҩƷid
  --       price_type_old    N    ԭ�۸����ͣ�0-���ۣ�1-ʱ��
  --       price_type_new    N    �¼۸����ͣ�0-���ۣ�1-ʱ��
  --output      
  --  code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message  C  1  Ӧ����Ϣ��
  ---------------------------------------------------------------------------
  j_Jsonin   Pljson;
  j_Json     Pljson;
  j_Jsonlist Pljson_List;
  n_Count    Number;

  n_ҩƷid     ҩƷ���.ҩƷid%Type;
  n_ԭ�۸����� Number(1); --0-���ۣ�1-ʱ��
  n_�¼۸����� Number(1); --0-���ۣ�1-ʱ��

  n_���۽��     ҩƷ���.ʵ�ʽ��%Type;
  n_�շ�id       ҩƷ�շ���¼.Id%Type;
  n_��ͨ���С�� Number;
  n_���         Number(8);
  n_������id   Number(18); --������
  v_Billno       ҩƷ�շ���¼.No%Type; --���۵���
  n_�۸�id       �շѼ�Ŀ.Id%Type;
  n_�շѼ�Ŀ�ּ� �շѼ�Ŀ.�ּ�%Type;
  n_�շѼ�Ŀԭ�� �շѼ�Ŀ.ԭ��%Type;
  n_ԭ��         ҩƷ�۸��¼.ԭ��%Type;
  n_ҩƷ�۸��¼ Number(1);
  v_���         �շ���ĿĿ¼.���%Type;

  --����->ʱ�ۺ����ҩƷ�۸��¼��ֵ
  Cursor c_Priceadjust Is
    Select s.ҩƷid, s.�ⷿid, Nvl(s.����, 0) As ����, s.�ϴι�Ӧ��id As ��Ӧ��id, s.�ϴ����� As ����, s.Ч��, s.�ϴβ��� As ����,
           Nvl(s.ʵ������, 0) As ʵ������, s.�ϴο��� As ����, Nvl(s.ʵ�ʽ��, 0) As ʵ�ʽ��, Nvl(s.ʵ�ʲ��, 0) As ʵ�ʲ��, Nvl(s.���ۼ�, 0) As ���ۼ�,
           s.ƽ���ɱ���, s.���Ч��, s.��׼�ĺ�, s.�ϴ��������� As ��������
    From ҩƷ��� S
    Where s.ҩƷid = n_ҩƷid And s.���� = 1
    Order By s.ҩƷid, s.����, s.�ⷿid;

  r_Priceadjust c_Priceadjust%RowType;
Begin
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  j_Jsonlist := j_Json.Get_Pljson_List('drug_list');

  n_Count := j_Jsonlist.Count;
  If j_Jsonlist.Count = 0 Then
    Json_Out := Zljsonout('δ����ҩƷ��Ϣ��');
    Return;
  End If;

  For I In 1 .. j_Jsonlist.Count Loop
    j_Json := Pljson();
    j_Json := Pljson(j_Jsonlist.Get(I));
  
    n_ҩƷid     := j_Json.Get_Number('drug_id');
    n_ԭ�۸����� := j_Json.Get_Number('price_type_old');
    n_�¼۸����� := j_Json.Get_Number('price_type_new');
  
    If n_ԭ�۸����� <> n_�¼۸����� Then
      --ȡԭ�ۺ�ԭ��id(���øù���ǰ�Ѿ��������¼۸�)
      Begin
        Select ԭ��, �ּ�, ԭ��id As �۸�id
        Into n_�շѼ�Ŀԭ��, n_�շѼ�Ŀ�ּ�, n_�۸�id
        From �շѼ�Ŀ
        Where �շ�ϸĿid = n_ҩƷid And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And �䶯ԭ�� = 1;
      Exception
        When Others Then
          n_�շѼ�Ŀԭ�� := Null;
          n_�շѼ�Ŀ�ּ� := Null;
          n_�۸�id       := Null;
      End;
    
      --ʱ��->����
      If n_ԭ�۸����� = 1 And n_�¼۸����� = 0 Then
        Select ���� Into n_��ͨ���С�� From ҩƷ���ľ��� Where ��� = 1 And ���� = 4 And ��λ = 5;
      
        --ȡ������ID
        Select ���id Into n_������id From ҩƷ�������� Where ���� = 13;
      
        n_���   := 0;
        v_Billno := Null;
      
        For r_Priceadjust In c_Priceadjust Loop
          If n_�շѼ�Ŀ�ּ� <> r_Priceadjust.���ۼ� Then
            If v_Billno Is Null Then
              Select Nextno(147) Into v_Billno From Dual;
            End If;
            n_��� := n_��� + 1;
            Select ҩƷ�շ���¼_Id.Nextval Into n_�շ�id From Dual;
            n_���۽�� := Round(n_�շѼ�Ŀ�ּ� * r_Priceadjust.ʵ������, n_��ͨ���С��) -
                      Round(r_Priceadjust.���ۼ� * r_Priceadjust.ʵ������, n_��ͨ���С��);
            --��������Ӱ���¼
            Insert Into ҩƷ�շ���¼
              (ID, ��¼״̬, ����, NO, ���, ������id, ҩƷid, ����, ����, Ч��, ����, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ����, ���۽��, ���, ժҪ,
               ������, ��������, �ⷿid, ���ϵ��, �۸�id, �����, �������, ����, Ƶ��, ��ҩ��λid)
            Values
              (n_�շ�id, 1, 13, v_Billno, n_���, n_������id, r_Priceadjust.ҩƷid, r_Priceadjust.����, r_Priceadjust.����,
               r_Priceadjust.Ч��, r_Priceadjust.����, 1, r_Priceadjust.ʵ������, 0, r_Priceadjust.���ۼ�, 0, n_�շѼ�Ŀ�ּ�,
               r_Priceadjust.����, n_���۽��, n_���۽��, 'ʱ��ת����', zl_UserName, Sysdate, r_Priceadjust.�ⷿid, 1, n_�۸�id,
               zl_UserName, Sysdate, r_Priceadjust.ʵ�ʽ��, r_Priceadjust.ʵ�ʲ��, r_Priceadjust.��Ӧ��id);
          
            Zl_ҩƷ���_Update(n_�շ�id, 2, 0);
          End If;
        End Loop;
      
        --����->ʱ��
      Elsif n_ԭ�۸����� = 0 And n_�¼۸����� = 1 Then
        For r_Priceadjust In c_Priceadjust Loop
          n_ҩƷ�۸��¼ := 0;
          Begin
            Select 1, �ּ�
            Into n_ҩƷ�۸��¼, n_ԭ��
            From ҩƷ�۸��¼
            Where ҩƷid = r_Priceadjust.ҩƷid And �ⷿid = r_Priceadjust.�ⷿid And Nvl(����, 0) = r_Priceadjust.���� And
                  ��¼״̬ = 1 And �۸����� = 1;
          Exception
            When Others Then
              n_ҩƷ�۸��¼ := 0;
              n_ԭ��         := n_�շѼ�Ŀԭ��;
          End;
        
          If n_ҩƷ�۸��¼ = 1 Then
            Zl_ҩƷ�۸��¼_Stop(1, r_Priceadjust.�ⷿid, r_Priceadjust.ҩƷid, r_Priceadjust.����, Sysdate - 1 / 24 / 60 / 60, 2);
          End If;
          Zl_ҩƷ�۸��¼_Insert(0, 1, r_Priceadjust.�ⷿid, r_Priceadjust.ҩƷid, r_Priceadjust.����, n_ԭ��, n_�շѼ�Ŀ�ּ�, Sysdate,
                           '����תʱ��', zl_UserName, Null, r_Priceadjust.��Ӧ��id, r_Priceadjust.����, r_Priceadjust.Ч��,
                           r_Priceadjust.����, r_Priceadjust.���Ч��, Null, Null, Null, Null, 1);
        
          Update ҩƷ���
          Set ���ۼ� = n_�շѼ�Ŀ�ּ�
          Where ���� = 1 And �ⷿid = r_Priceadjust.�ⷿid And ҩƷid = r_Priceadjust.ҩƷid And Nvl(����, 0) = r_Priceadjust.����;
        
        End Loop;
      End If;
    End If;
  End Loop;

  Json_Out := '{"output":{"code": 1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_DrugSvr_AdjustPriceType;
/


Create Or Replace Procedure Zl_DrugSvr_CheckDrugExistRec
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --ִ�е���
  ---------------------------------------------------------------------------
  --input      �ж�ҩƷ�Ƿ�����շ���¼
  --  drug_id      N    ҩƷid
  --output      
  --  code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message  C  1  Ӧ����Ϣ��
  --  isexist  N 1 �Ƿ����: 1-����;0-������
  ---------------------------------------------------------------------------
  j_Jsonin Pljson;
  j_Json   Pljson;
  n_ҩƷid ҩƷ���.ҩƷid%Type;
  n_Exist  Number(1);
Begin
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_ҩƷid := j_Json.Get_Number('drug_id');

  Select Count(1) Into n_Exist From ҩƷ�շ���¼ Where ҩƷid = n_ҩƷid And Rownum < 2;

  Json_Out := '{"output":{"code": 1,"message":"�ɹ�","isexist":' || n_Exist || '}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_DrugSvr_CheckDrugExistRec;
/


Create Or Replace Procedure Zl_Drugsvr_Checkpriceadjust
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --���۹���ģʽʱ���жϼ۸��Ƿ��������۹���Ҫ���ɱ��ۺ��ۼ�һ�£�
  ---------------------------------------------------------------------------
  --input      ���ҩƷ���ۿ���
  --  pharmacy_drug_ids     C    �ⷿҩƷid�����ⷿid,ҩƷid;...
  --  is_ignore    N    �������۹��� 0-�����ԣ�1-����
  --output      
  --  code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message  C  1  Ӧ����Ϣ��
  --  price_list[]  ҩƷ�ɱ��۵��ۼ�¼
  --     drug_id   N 1  ҩƷID
  --     drug_name   C 1  ҩƷ��Ϣ
  --     pharmacy_id   id 1  �ⷿid
  --     pharmacy_name   C 1  �ⷿ����
  --     isstock   N 1 ��ʾ���ͣ�0-û�п�棬1-�п��
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  n_ҩƷid         ҩƷ���.ҩƷid%Type;
  n_�ⷿid         ҩƷ���.�ⷿid%Type;
  n_���۹���     Number(1);
  n_�ɱ���         ҩƷ���.�ɱ���%Type;
  n_�ۼ�           ҩƷ���.�ϴ��ۼ�%Type;
  v_ͨ����         Varchar2(32767);
  n_���           Number(18);
  n_�������۹��� Number(1);
  v_ҩƷ��         Varchar2(4000);
  v_Fields         Varchar2(4000);

  v_Jtmp       Varchar2(32767);
  n_Checkvalue Number(1);
Begin
  Select zl_To_Number(Nvl(zl_GetSysParameter(275), '0')) Into n_���۹��� From Dual;

  If n_���۹��� = 0 Then
    --û�������۹���ʱ�˳�
    Json_Out := zlJsonOut('�ɹ�', 1);
    Return;
  End If;

  j_Jsonin         := PLJson(Json_In);
  j_Json           := j_Jsonin.Get_Pljson('input');
  v_ҩƷ��         := j_Json.Get_String('pharmacy_drug_ids');
  n_�������۹��� := Nvl(j_Json.Get_Number('is_ignore'), 0);

  If v_ҩƷ�� Is Null Then
    Json_Out := zlJsonOut('δ����ҩƷ��Ϣ������!');
    Return;
  End If;

  v_ҩƷ�� := v_ҩƷ�� || ';';

  While v_ҩƷ�� Is Not Null Loop
    v_Fields := Substr(v_ҩƷ��, 1, Instr(v_ҩƷ��, ';') - 1);
    n_�ⷿid := To_Number(Substr(v_Fields, 1, Instr(v_Fields, ',') - 1));
    n_ҩƷid := To_Number(Substr(v_Fields, Instr(v_Fields, ',') + 1));
    v_ҩƷ�� := Replace(';' || v_ҩƷ��, ';' || v_Fields || ';');
  
    If n_ҩƷid = 0 Then
      Json_Out := zlJsonOut('δ����ҩƷID��Ϣ������');
      Return;
    End If;
  
    n_�ɱ��� := Null;
    n_�ۼ�   := Null;
    v_ͨ���� := Null;
  
    Select a.�ɱ���, b.�ּ� As �ۼ�, '[' || c.���� || ']' || c.���� || Decode(c.����, Null, Null, '(' || c.���� || ')') || c.��� As ͨ����,
           Nvl(a.�Ƿ����۹���, 0) As ���۹���
    Into n_�ɱ���, n_�ۼ�, v_ͨ����, n_���۹���
    From ҩƷ��� A, �շѼ�Ŀ B, �շ���ĿĿ¼ C
    Where a.ҩƷid = b.�շ�ϸĿid And a.ҩƷid = c.Id And (Sysdate Between b.ִ������ And b.��ֹ����) And a.ҩƷid = n_ҩƷid;
  
    --������޿��
    If n_�ⷿid > 0 Then
      Select Count(*)
      Into n_���
      From ҩƷ���
      Where ���� = 1 And ҩƷid = n_ҩƷid And �ⷿid = n_�ⷿid And
            Not (���� = 0 And �������� < 0 And ʵ������ = 0 And ʵ�ʽ�� = 0 And ʵ�ʲ�� = 0);
    Else
      Select Count(*)
      Into n_���
      From ҩƷ���
      Where ���� = 1 And ҩƷid = n_ҩƷid And Not (���� = 0 And �������� < 0 And ʵ������ = 0 And ʵ�ʽ�� = 0 And ʵ�ʲ�� = 0);
    End If;
  
    If n_��� = 0 Then
      --�޿��ʱ�����շѼ�Ŀȡ�ۼۣ���ҩƷ���ȡ�ɱ��ۣ����Ƚϼ۸�
      If n_���۹��� = 0 And n_�������۹��� = 0 Then
        --���������۹���
        n_Checkvalue := 0;
      Else
        If n_�ɱ��� = n_�ۼ� Then
          --�ۼۺͳɱ���һ��ʱ
          n_Checkvalue := 0;
        Else
          --�ۼۺͳɱ��۲�һ��ʱ
          n_Checkvalue := 1;
        End If;
      End If;
    
      If n_Checkvalue = 1 Then
        v_Jtmp := v_Jtmp || ',';
        zlJsonPutValue(v_Jtmp, 'drug_id', n_ҩƷid, 1, 1);
        zlJsonPutValue(v_Jtmp, 'drug_name', v_ͨ����, 0);
        zlJsonPutValue(v_Jtmp, 'pharmacy_id', n_�ⷿid, 1);
        zlJsonPutValue(v_Jtmp, 'pharmacy_name', '', 0);
        zlJsonPutValue(v_Jtmp, 'isstock', 0, 1, 2);
      End If;
    Else
      --�п������ʱ
      n_Checkvalue := 0;
      For r_�۸� In (Select ҩƷid, ͨ����, ���, �ⷿid, �ⷿ, ������, '' As ����, ����, ��λ, ҩ���װ, �ۼ�, Sum(�ɱ��� * ʵ������) / Sum(ʵ������) As �ɱ���,
                          �Ƿ�ʱ��
                   From (Select a.ҩƷid,
                                 '[' || c.���� || ']' || c.���� || Decode(c.����, Null, Null, '(' || c.���� || ')') || c.��� As ͨ����,
                                 c.���, c.���� As ������, Null As ����, a.ҩ�ⵥλ As ��λ, a.ҩ���װ, b.�ּ� As �ۼ�, d.ƽ���ɱ��� As �ɱ���, 0 As �Ƿ�ʱ��,
                                 d.ʵ������, d.�ⷿid, e.���� As �ⷿ
                          From ҩƷ��� A, �շѼ�Ŀ B, �շ���ĿĿ¼ C, ҩƷ��� D, ���ű� E
                          Where a.ҩƷid = b.�շ�ϸĿid And a.ҩƷid = c.Id And a.ҩƷid = d.ҩƷid And d.���� = 1 And
                                (Sysdate Between b.ִ������ And b.��ֹ����) And
                                (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And c.�Ƿ��� = 0 And
                                Nvl(a.�Ƿ����۹���, 0) = 1 And b.�ּ� <> d.ƽ���ɱ��� And
                                d.�ⷿid In (Select Distinct ����id From ��������˵�� Where �������� Like '%ҩ��') And
                                a.ҩƷid = Decode(n_ҩƷid, 0, a.ҩƷid, n_ҩƷid) And d.�ⷿid = Decode(n_�ⷿid, 0, d.�ⷿid, n_�ⷿid) And
                                Not (d.���� = 0 And d.�������� < 0 And d.ʵ������ = 0 And d.ʵ�ʽ�� = 0 And d.ʵ�ʲ�� = 0) And
                                d.�ⷿid = e.Id)
                   Group By ҩƷid, ͨ����, ���, �ⷿid, �ⷿ, ������, ����, ��λ, ҩ���װ, �ۼ�, �Ƿ�ʱ��
                   Having Sum(ʵ������) <> 0
                   Union All
                   Select a.ҩƷid,
                          '[' || c.���� || ']' || c.���� || Decode(c.����, Null, Null, '(' || c.���� || ')') || c.��� As ͨ����, c.���,
                          d.�ⷿid, e.���� As �ⷿ, d.�ϴβ��� As ������, d.�ϴ����� As ����, d.����, a.ҩ�ⵥλ As ��λ, a.ҩ���װ,
                          Nvl(d.���ۼ�, 0) As �ۼ�, d.ƽ���ɱ��� As �ɱ���, 1 As �Ƿ�ʱ��
                   From ҩƷ��� A, �շ���ĿĿ¼ C, ҩƷ��� D, ���ű� E
                   Where a.ҩƷid = c.Id And a.ҩƷid = d.ҩƷid And d.�ⷿid = e.Id And d.���� = 1 And c.�Ƿ��� = 1 And
                         (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And Nvl(a.�Ƿ����۹���, 0) = 1 And
                         Nvl(d.���ۼ�, 0) <> d.ƽ���ɱ��� And
                         d.�ⷿid In (Select Distinct ����id From ��������˵�� Where �������� Like '%ҩ��') And
                         a.ҩƷid = Decode(n_ҩƷid, 0, a.ҩƷid, n_ҩƷid) And d.�ⷿid = Decode(n_�ⷿid, 0, d.�ⷿid, n_�ⷿid) And
                         Not (d.���� = 0 And d.�������� < 0 And d.ʵ������ = 0 And d.ʵ�ʽ�� = 0 And d.ʵ�ʲ�� = 0)
                   Order By ͨ����, �ⷿid, ����) Loop
      
        --�ҵ�����ʱ
        n_Checkvalue := 1;
        v_Jtmp       := v_Jtmp || ',';
        zlJsonPutValue(v_Jtmp, 'drug_id', r_�۸�.ҩƷid, 1, 1);
        zlJsonPutValue(v_Jtmp, 'drug_name', r_�۸�.ͨ����, 0);
        zlJsonPutValue(v_Jtmp, 'pharmacy_id', r_�۸�.�ⷿid, 1);
        zlJsonPutValue(v_Jtmp, 'pharmacy_name', r_�۸�.�ⷿ, 0);
        zlJsonPutValue(v_Jtmp, 'isstock', 1, 1, 2);
      End Loop;
    End If;
  End Loop;

  If v_Jtmp Is Not Null Then
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","price_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    Json_Out := zlJsonOut('�ɹ�', 1);
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Checkpriceadjust;
/




Create Or Replace Procedure Zl_Drugsvr_Executeprice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --ִ�е���
  ---------------------------------------------------------------------------
  --input      ���ҩƷ�ۼۣ��ɱ����Ƿ��������Ч��δִ�еļ۸����������ִ�е���
  --  drug_id      N    ҩƷid
  --output      
  --  code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message  C  1  Ӧ����Ϣ��
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  n_ҩƷid ҩƷ���.ҩƷid%Type;

  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin
  j_Jsonin := PLJson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_ҩƷid := j_Json.Get_Number('drug_id');

  If n_ҩƷid = 0 Then
    v_Err_Msg := 'δ����ҩƷID��Ϣ��';
    Raise Err_Item;
  End If;

  For r_���� In (Select Distinct b.ҩƷid As ҩƷid
               From �շ���ĿĿ¼ I, �շѼ�Ŀ N, ҩƷ��� B
               Where i.Id = n.�շ�ϸĿid And i.Id = b.ҩƷid And
                     (i.����ʱ�� Is Null Or i.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) And n.�䶯ԭ�� = 0 And
                     Sysdate > n.ִ������ And n.�۸�ȼ� Is Null And b.ҩƷid = n_ҩƷid
               Union
               Select Distinct a.ҩƷid
               From ҩƷ�۸��¼ A, ҩƷ��� B
               Where a.ҩƷid = b.ҩƷid And a.��¼״̬ = 0 And a.ִ������ <= Sysdate And b.ҩƷid = n_ҩƷid) Loop
  
    n_ҩƷid := r_����.ҩƷid;
    Exit;
  End Loop;

  If n_ҩƷid > 0 Then
    Zl_ҩƷ�շ���¼_Adjust(n_ҩƷid);
  End If;

  Json_Out := '{"output":{"code": 1,"message":"�ɹ�"}}';
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugsvr_Executeprice;
/
Create Or Replace Procedure Zl_Drugsvr_Patiinfoupdate
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ�סԺ����ҩƷ�շ���¼������Ϣ�޸�
  --��Σ�Json_In:��ʽ
  --  input
  --   pati_name      C   1   ��������
  --   pati_sex       C   1   �����Ա�
  --   pati_age       C   1   ��������
  --   visit_id       N   1   ����id
  --   pati_id        N   1   ����id

  --����: Json_Out,��ʽ����
  --  output
  --    code                 N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message              C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  v_����   Varchar2(100);
  v_�Ա�   Varchar2(100);
  v_����   Varchar2(100);
  n_����id Number;
  n_����id Number;
Begin
  --�������
  j_Jsonin := PLJson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');
  n_����id := j_Json.Get_Number('visit_id');
  v_����   := j_Json.Get_String('pati_name');
  v_�Ա�   := j_Json.Get_String('pati_sex');
  v_����   := j_Json.Get_String('pati_age');

  Update ҩƷ�շ���¼
  Set ���� = Nvl(v_����, ����), �Ա� = Nvl(v_�Ա�, �Ա�), ���� = Nvl(v_����, ����)
  Where ����id = n_����id And ��ҳid = n_����id;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugsvr_Patiinfoupdate;
/

Create Or Replace Procedure Zl_Drugsvr_Getretained
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  --------------------------------------------------------------------
  --���ܣ�סԺҽ��ҽ�����˷���ʱҩƷҽ�������������
  --��Σ�JOSN��ʽ
  --input
  --     order_ids                     C 1 ҽ��IDƴ��
  --���Σ�JSON
  --output
  --     code                          N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --     message                       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --     data[]�б�
  --         warehouse                 C 1 �ⷿ����
  --         drug_name                 C 1 ҩƷ���ƣ�������ĿĿ¼.����
  --         inp_unit                  C 1 סԺ��λ
  --         re_quantity               N 1 ��������
  --         quantity                  N 1 ��������
  --------------------------------------------------------------------
  l_Vals    t_Strlist;
  P         Number;
  c_ҽ��ids Clob;

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;

  j_Jsonin Pljson;
  j_Json   Pljson;
Begin
  --�������
  j_Jsonin  := Pljson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  c_ҽ��ids := j_Json.Get_Clob('order_ids');

  l_Vals    := t_Strlist();
  c_ҽ��ids := c_ҽ��ids || ',';
  Loop
    P := Instr(c_ҽ��ids, ',');
    Exit When(Nvl(P, 0) = 0);
    l_Vals.Extend;
    l_Vals(l_Vals.Count) := (Substr(c_ҽ��ids, 1, P - 1));
    c_ҽ��ids := Substr(c_ҽ��ids, P + 1);
  End Loop;

  v_Jtmp := Null;
  For r_Data In (Select d.���� As �ⷿ, e.���� As ҩƷ, (a.���� / Nvl(b.סԺ��װ, 1)) As ��������, (a.���� / Nvl(b.סԺ��װ, 1)) As ��������, b.סԺ��λ
                 From (Select a.�ⷿid, a.ҩƷid, a.�������� As ����, b.�������� As ����
                        From (Select a.�ⷿid, a.���˲���id As ����id, a.ҩƷid, Sum(a.ʵ������) As ��������
                               From ҩƷ�շ���¼ A
                               Where a.ҽ��id In (Select /*+cardinality(f,10)*/
                                                 To_Number(f.Column_Value) As ҽ��id
                                                From Table(l_Vals) F)
                               Group By a.�ⷿid, a.���˲���id, a.ҩƷid) A, ҩƷ����ƻ� B
                        Where a.����id = b.����id(+) And a.�ⷿid = b.�ⷿid(+) And a.ҩƷid = b.ҩƷid(+) And b.״̬(+) = 0) A, ҩƷ��� B,
                      ���ű� D, ������ĿĿ¼ E
                 Where a.�ⷿid = d.Id And a.ҩƷid = b.ҩƷid And b.ҩ��id = e.Id And a.���� <> 0 And a.���� > a.����) Loop
  
    v_Jtmp := v_Jtmp || ',';
    Zljsonputvalue(v_Jtmp, 'warehouse', r_Data.�ⷿ, 0, 1);
    Zljsonputvalue(v_Jtmp, 'drug_name', r_Data.ҩƷ, 0);
    Zljsonputvalue(v_Jtmp, 'inp_unit', r_Data.סԺ��λ, 0);
    Zljsonputvalue(v_Jtmp, 're_quantity', r_Data.��������, 1);
    Zljsonputvalue(v_Jtmp, 'quantity', r_Data.��������, 1, 2);
  
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
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getretained;
/

Create Or Replace Procedure Zl_Drugsvr_Newrecbill_Check
(
  Json_In  Clob,
  Json_Out Out Varchar2
) Is
  --����Zl_ҩƷ���۳���_Check��ɹ���
  --�������̣�����Ҫ�������
  --�������ݣ�
  --1.��ֵ������������ⷿ���ü��
  --2.ҩƷ�����ۼۼ�飬�շѽ���ͱ���ʱ���ܷ����仯��ʱ�۷�����
  --3.����飬���ݲ��������ʵ�����
  --4.�������Լ�飬�������Ա仯
  -------------------------------------------------------------------------------------------------
  --���      json
  --input     ����������Ҫ�����Ĵ������м��
  --  fee_list      �շ���ϸ��Ϣ��֧�ֶ����[����]
  --    drug_id   N 1 ҩƷid
  --    send_num  N 1 ��ҩ����
  --    pharmacy_id N 1 ҩ��id
  --    price           N       1       �ۼ�
  --����      json
  --output      
  --  code  C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message C 1 Ӧ����Ϣ��
  -------------------------------------------------------------------------------------------------
Begin

  Zl_ҩƷ���۳���_Check(Json_In, Json_Out);

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Newrecbill_Check;
/
Create Or Replace Procedure Zl_Drugsvr_Skintest_Update
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  ---------------------------------------------------------------------------
  --���ܣ����´�������Ƥ�Խ��
  ---------------------------------------------------------------------------
  --��Σ�Json_In:��ʽ
  --input  
  --   order_ids              C 1 ҽ��ids������ƴ��
  --   skintest_info          C 1 (+)����,(-)����,��
  --����: Json_Out,��ʽ����
  --output
  --    code                   C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message                C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  v_Ƥ�Խ��  Varchar2(60);
  j_Jsonin    Pljson;
  j_Json      Pljson;
  v_Order_Ids Varchar2(32767);
Begin
  --�������
  j_Jsonin    := Pljson(Json_In);
  j_Json      := j_Jsonin.Get_Pljson('input');
  v_Order_Ids := j_Json.Get_String('order_ids');
  v_Ƥ�Խ��  := j_Json.Get_String('skintest_info');

  Update ҩƷ�շ���¼
  Set Ƥ�Խ�� = v_Ƥ�Խ��
  Where ���� In (8, 9) And ������� Is Not Null And
        ҽ��id In (Select /*+cardinality(x,10)*/
                  x.Column_Value
                 From Table(Cast(f_Num2list(v_Order_Ids) As Zltools.t_Numlist)) X);

  Json_Out := '{"output":{"code": 1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugsvr_Skintest_Update;
/

Create Or Replace Procedure Zl_Drugsvr_Patibase_Upate
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --���´����еĲ��˻�����Ϣ
  ---------------------------------------------------------------------------
  --���      json
  --input     ���´����еĲ��˻�����Ϣ
  --  pati_id N 1 ����id
  --  register_id N   �Һ�id
  --  pati_pageid N   ��ҳid
  --  pati_name C   ����
  --  pati_sex  C   �Ա�
  --  pati_age  C   ����
  --  pati_birthdate  D   ��������
  --  pati_Idcard C   ���֤��
  --����      json
  --output      
  --code  C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --message C 1 Ӧ����Ϣ��
  ---------------------------------------------------------------------------
  n_����id   Number(18);
  v_����     Varchar2(100);
  v_�Ա�     Varchar2(4);
  v_����     Varchar2(20);
  d_�������� Date;
  v_���֤�� Varchar2(18);

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --�������
  j_Jsonin   := PLJson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_����id   := j_Json.Get_Number('pati_id');
  v_����     := j_Json.Get_String('pati_name');
  v_�Ա�     := j_Json.Get_String('pati_sex');
  v_����     := j_Json.Get_String('pati_age');
  d_�������� := To_Date(j_Json.Get_String('pati_birthdate'), 'yyyy-mm-dd');
  v_���֤�� := j_Json.Get_String('pati_Idcard');

  --�޸�δ��ҩƷ��¼
  Update δ��ҩƷ��¼ Set ���� = v_���� Where ����id = n_����id;

  --�޸�δ��˵�ҩƷ�շ���¼
  Update ҩƷ�շ���¼
  Set ���� = v_����, �Ա� = v_�Ա�, ���� = v_����, �������� = d_��������, ���֤�� = v_���֤��
  Where ������� Is Null And ����id = n_����id;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugsvr_Patibase_Upate;
/
Create Or Replace Procedure Zl_Drugsvr_Autosenddrug
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ��շѻ���ʺ��Զ���ҩ����NO�򴦷���ϸ��
  --��Σ�Json_In:��ʽ
  --  input
  --    billtype             N 1 ��������:1 -�շѴ�����ҩ  ;2- ���ʵ�������ҩ;3- ���ʱ�����ҩ;
  --    operator_code        C 1 ����Ա���
  --    operator_name        C 1 ����Ա����
  --    rcp_nos              C 1 ����NO����NO1,NO2...
  --    rcpdtl_ids           C 1 ������ϸid��,Ŀǰ����ķ���ID�����ö��ŷָ� ��1,2,3,4
  --    send_type            N 0 ��ҩ����,0-�� �����no�������ͷ�ҩ,1-ֻ�� ������ϸid����ҩ
  --����: Json_Out,��ʽ����
  --  output
  --    code                 N 1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message              C 1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  n_����       ҩƷ�շ���¼.����%Type;
  v_����Ա���� ��Ա��.����%Type := Null;
  v_����Ա��� ��Ա��.���%Type := Null;
  v_Err        Varchar2(255);
  v_Nos        Varchar2(32767);
  Err_Custom Exception;
  v_Ids       Varchar2(32767);
  n_Send_Type Number;
Begin
  --�������
  j_Jsonin    := PLJson(Json_In);
  j_Json      := j_Jsonin.Get_Pljson('input');
  n_Send_Type := j_Json.Get_Number('send_type');

  If Nvl(n_Send_Type, 0) = 0 Then
    n_���� := j_Json.Get_Number('billtype');
    If n_���� = 1 Then
      n_���� := 8;
    Elsif n_���� = 2 Then
      n_���� := 9;
    Elsif n_���� = 3 Then
      n_���� := 10;
    Else
      v_Err := '����ڵ㡾billtype���������飡';
      Raise Err_Custom;
    End If;
  
    v_Nos := j_Json.Get_String('rcp_nos');
    v_Ids := j_Json.Get_String('rcpdtl_ids');
  
    If v_Ids Is Null And v_Nos Is Null Then
      v_Err := 'δ����ҩƷ���ݡ�rcp_nos���ڵ����ϸ��Ϣ��rcpdtl_ids���ڵ�';
      Raise Err_Custom;
      Return;
    End If;
  End If;

  If Nvl(n_Send_Type, 0) = 1 Then
    v_Ids := j_Json.Get_String('rcpdtl_ids');
    If v_Ids Is Null Then
      v_Err := 'δ����ҩƷ��ϸ��Ϣ��rcpdtl_ids���ڵ�';
      Raise Err_Custom;
      Return;
    End If;
  End If;

  v_����Ա���� := j_Json.Get_String('operator_name');
  v_����Ա��� := j_Json.Get_String('operator_code');

  --�������ŷ�ҩ
  If v_Nos Is Not Null Then
    Zl_ҩƷ�շ���¼_�Զ���ҩ_s(n_����, v_����Ա����, v_����Ա���, v_Nos, 0);
  End If;

  --������ID��ҩ
  If v_Ids Is Not Null Then
    Zl_ҩƷ�շ���¼_�Զ���ҩ_s(n_����, v_����Ա����, v_����Ա���, v_Ids, 1);
  End If;

  Json_Out := '{"output":{"code":1,"message": "�ɹ�"}}';
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugsvr_Autosenddrug;
/

Create Or Replace Procedure Zl_Drugsvr_Autoreturndrug
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ��Զ���ҩ����������ϸ������ID��ҩ��Ĭ����ȫ�ˣ�
  --��Σ�Json_In:��ʽ
  --  input
  --    audit_operator        C 1 �����
  --    rcpdtl_list[]   ��ҩ�б���Ϣ
  --         rcpdtl_id        N 1 ����id
  --         re_quantity      N 0 ��ҩ���������۵����շ���ĿĿ¼.���㵥λ�����ɲ����˽����ߴ��գ���ʾȫ�ˣ�����ָ��������ҩ
  --����: Json_Out,��ʽ����
  --  output
  --    code                 N 1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message              C 1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Jsonin     Pljson;
  j_Json       Pljson;
  j_Jsonlist   Pljson_List;
  d_����ʱ��   ҩƷ�շ���¼.�������%Type;
  v_����Ա���� ��Ա��.����%Type;
  n_����       ҩƷ�շ���¼.ʵ������%Type;
  n_��ҩ����   ҩƷ�շ���¼.ʵ������%Type;
  n_С��       Number(6);
  n_����id     Number(18);
Begin
  --�������
  j_Jsonin     := Pljson(Json_In);
  j_Json       := j_Jsonin.Get_Pljson('input');
  v_����Ա���� := j_Json.Get_String('audit_operator');
  j_Jsonlist   := j_Json.Get_Pljson_List('rcpdtl_list');

  --ȡ��ͨҵ�񾫶�λ��
  --���:1-ҩƷ 2-����
  --���ݣ�2-���ۼ� 4-���
  --��λ��ҩƷ:1-�ۼ� 5-��λ
  Select ���� Into n_С�� From ҩƷ���ľ��� Where ��� = 1 And ���� = 4 And ��λ = 5;
  Select Sysdate Into d_����ʱ�� From Dual;

  j_Json := Pljson();
  For I In 1 .. j_Jsonlist.Count Loop
    j_Json   := Pljson(j_Jsonlist.Get(I));
    n_����id := j_Json.Get_Number('rcpdtl_id');
    n_����   := j_Json.Get_Number('re_quantity');
    j_Json   := Pljson();
    If n_���� Is Not Null Then
      n_��ҩ���� := n_����;
    End If;
    --�ֽ���ҩ����
    For r_������ϸ In (Select a.�ⷿid, a.Id, Nvl(a.����, 1) * a.ʵ������ As ����
                   From ҩƷ�շ���¼ A
                   Where a.���� In (8, 9, 10) And (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0) And ����� Is Not Null And
                         a.����id = n_����id
                   Order By a.�ⷿid, a.ҩƷid, a.����) Loop
      If n_���� Is Null Then
        --���������Ϊ�ձ�ʾȫ��      
        --������ҩ��������ϸ��
        Zl_ҩƷ�շ���¼_������ҩ_s(r_������ϸ.Id, v_����Ա����, d_����ʱ��, Null, Null, Null, Null, Null, Null, n_С��);
      Else
        If n_��ҩ���� > 0 Then
          If n_��ҩ���� > r_������ϸ.���� Then
            --������ҩ��������ϸ��
            Zl_ҩƷ�շ���¼_������ҩ_s(r_������ϸ.Id, v_����Ա����, d_����ʱ��, Null, Null, Null, r_������ϸ.����, Null, Null, n_С��);
          
            n_��ҩ���� := n_��ҩ���� - r_������ϸ.����;
          Else
            --������ҩ��������ϸ��
            Zl_ҩƷ�շ���¼_������ҩ_s(r_������ϸ.Id, v_����Ա����, d_����ʱ��, Null, Null, Null, n_��ҩ����, Null, Null, n_С��);
            n_��ҩ���� := 0;
          End If;
        End If;
      End If;
    End Loop;
  End Loop;
  Json_Out := '{"output":{"code":1,"message": "�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugsvr_Autoreturndrug;
/
Create Or Replace Procedure Zl_Drugsvr_Checkpatiexecute
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ݲ�����Ϣ��ȡδ��ҩƷ����
  --��Σ�Json_In:��ʽ
  --  input
  --     pati_id                N 1 ����ID
  --     pati_pageid            N 1 ��ҳID
  --     baby_num               N 0 Ӥ�����:-1��ʾ������;0-ĸ�׵�;>0����Ӥ������
  --     check_excutenature     N 0 ���Ժ��ҩ��1-��Ҫ������Ժ��ҩ���;0-����Ҫ��Ժ��ҩ���
  --     fee_source             N 1 ������Դ:1-����;2-סԺ;4-���
  --     rcp_nos            �������ݺţ������磺["A0001","A0002"]
  --����: Json_Out,��ʽ����
  --  output
  --    code                    N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                 C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    data{}
  --       isexist              N 1 �Ƿ����: 1-����;0-������
  --       drug_notsend_infor   C 1 δ��ҩ��Ϣ,isexist=1ʱ����
  ---------------------------------------------------------------------------
  v_Drug     Varchar2(4000);
  n_����id   ҩƷ�շ���¼.����id%Type;
  n_��ҳid   ҩƷ�շ���¼.��ҳid%Type;
  n_Ӥ����� ҩƷ�շ���¼.Ӥ�����%Type;
  v_No       ҩƷ�շ���¼.No%Type;
  n_������Դ ҩƷ�շ���¼.������Դ%Type;
  n_Ժ���ҩ Number(2);
  n_Add      Number(2);
  j_Jsonlist Pljson_List := Pljson_List();
  n_Count    Number(18);
  l_Nos      t_Strlist := t_Strlist();

  v_��Ŀ �շ���ĿĿ¼.����%Type;
  v_���� ���ű�.����%Type;
  v_���� Varchar2(100);

  Type t_δ��ҩƷ Is Ref Cursor;
  c_δ��ҩƷ t_δ��ҩƷ;

  j_Jsonin Pljson;
  j_Json   Pljson;
Begin
  --�������
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_����id   := j_Json.Get_Number('pati_id');
  n_��ҳid   := j_Json.Get_Number('pati_pageid');
  n_Ӥ����� := j_Json.Get_Number('baby_num');
  n_������Դ := j_Json.Get_Number('fee_source');
  n_Ժ���ҩ := j_Json.Get_Number('check_excutenature');

  j_Jsonlist := j_Json.Get_Pljson_List('rcp_nos');
  n_Count    := 0;
  If j_Jsonlist Is Not Null Then
    n_Count := j_Jsonlist.Count;
  End If;
  If n_Count <> 0 Then
    For I In 1 .. n_Count Loop
      v_No := j_Jsonlist.Get_String(I);
      l_Nos.Extend();
      l_Nos(l_Nos.Count) := v_No;
    End Loop;
  
    Open c_δ��ҩƷ For
      Select Distinct b.No, d.���� ��Ŀ, c.���� As ����, To_Char(b.����) As ����
      From ҩƷ�շ���¼ B, ���ű� C, �շ���ĿĿ¼ D
      Where b.ҩƷid = d.Id And b.�ⷿid + 0 = c.Id(+) And b.���� In (9, 10) And Mod(b.��¼״̬, 3) = 1 And b.����� Is Null And
            b.����id = n_����id And (Nvl(b.��ҳid, 0) = n_��ҳid And n_������Դ = 2 Or n_������Դ <> 2) And
            (Nvl(b.Ӥ�����, 0) = Nvl(n_Ӥ�����, 0) Or Nvl(n_Ӥ�����, 0) = -1) And Nvl(b.ժҪ, '��ҽ') <> '�ܷ�' And
            b.No In (Select Column_Value From Table(l_Nos)) And Nvl(b.������Դ, 1) = n_������Դ;
  Else
  
    Open c_δ��ҩƷ For
      Select Distinct b.No, d.���� ��Ŀ, c.���� As ����, To_Char(b.����) As ����
      From ҩƷ�շ���¼ B, ���ű� C, �շ���ĿĿ¼ D
      Where b.ҩƷid = d.Id And b.�ⷿid + 0 = c.Id(+) And b.���� In (9, 10) And Mod(b.��¼״̬, 3) = 1 And b.����� Is Null And
            b.����id = n_����id And (Nvl(b.��ҳid, 0) = n_��ҳid And n_������Դ = 2 Or n_������Դ <> 2) And
            (Nvl(b.Ӥ�����, 0) = Nvl(n_Ӥ�����, 0) Or n_Ӥ����� = -1) And Nvl(b.ժҪ, '��ҽ') <> '�ܷ�' And b.������Դ = n_������Դ;
  
  End If;

  Loop
    Fetch c_δ��ҩƷ
      Into v_No, v_��Ŀ, v_����, v_����;
    Exit When c_δ��ҩƷ%NotFound;
  
    n_Add := 1;
    If Substr(v_����, 2) = '3' Then
      n_Add := Nvl(n_Ժ���ҩ, 0);
    End If;
  
    If Nvl(n_Add, 0) = 1 Then
    
      If v_Drug Is Not Null Then
        If Instr(Chr(13) || Chr(10) || v_Drug || Chr(13) || Chr(10),
                 Chr(13) || Chr(10) || '����[' || Nvl(v_No, '') || ']�е�' || Nvl(v_��Ŀ, '') || '����' || Nvl(v_����, '[δ��ҩ��]') ||
                  'δ��ҩ' || Chr(13) || Chr(10), 1) = 0 Then
          If Lengthb(v_Drug || Chr(13) || Chr(10) || '����[' || Nvl(v_No, '') || ']�е�' || Nvl(v_��Ŀ, '') || '����' ||
                     Nvl(v_����, '[δ��ҩ��]') || 'δ��ҩ') <= 1000 Then
            v_Drug := v_Drug || Chr(13) || Chr(10) || '����[' || Nvl(v_No, '') || ']�е�' || Nvl(v_��Ŀ, '') || '����' ||
                      Nvl(v_����, '[δ��ҩ��]') || 'δ��ҩ';
          Else
            v_Drug := v_Drug || Chr(13) || Chr(10) || '... ...';
          End If;
        End If;
      Else
        v_Drug := '����[' || Nvl(v_No, '') || ']�е�' || Nvl(v_��Ŀ, '') || '����' || Nvl(v_����, '[δ��ҩ��]') || 'δ��ҩ';
      End If;
    End If;
  End Loop;

  n_Count := 0;
  If v_Drug Is Not Null Then
    v_Drug  := '����δ��ҩƷ��' || Chr(13) || Chr(10) || Chr(13) || Chr(10) || v_Drug;
    n_Count := 1;
  End If;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":{"isexist":' || n_Count || ',"drug_notsend_infor":"' ||
              Zljsonstr(v_Drug) || '"}}}';
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Checkpatiexecute;
/

Create Or Replace Procedure Zl_Drugsvr_Getrefusesendlist
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡ�ܷ�ҩ�嵥
  --��Σ�Json_In:��ʽ
  --  input
  --     pati_id            N 1 ����Id
  --     pati_pageids       C 1 ��ҳIDs:���סԺ ���ö��ŷ��� 
  --����: Json_Out,��ʽ����
  --  output
  --    code                N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    data []
  --        rcp_no          C   1   ���õ��ݺ�
  --        rcpdtl_id       C   1   ������ϸID,Ŀǰ������Ƿ���ID
  ---------------------------------------------------------------------------
  v_��ҳids Varchar2(4000);
  n_����id  ҩƷ�շ���¼.����id%Type;

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --�������
  j_Jsonin  := PLJson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  n_����id  := j_Json.Get_Number('pati_id');
  v_��ҳids := j_Json.Get_String('pati_pageids');

  If Nvl(n_����id, 0) = 0 Then
    Json_Out := zlJsonOut('δ������ز���id��Ϣ');
    Return;
  End If;

  For r_Info In (Select NO As Rcp_No, ����id As Rcpdtl_Id
                 From ҩƷ�շ���¼
                 Where ����id = n_����id And
                       (Instr(',' || Nvl(v_��ҳids, '-') || ',', ',' || Nvl(��ҳid, 0) || ',') > 0 Or v_��ҳids Is Null) And
                       Mod(��¼״̬, 3) = 1 And Nvl(ժҪ, '��һ') = '�ܷ�' And Instr(',8,9,10,', ',' || ���� || ',') > 0
                 Order By NO, ����id) Loop
  
    v_Jtmp := v_Jtmp || ',';
    zlJsonPutValue(v_Jtmp, 'rcp_no', r_Info.Rcp_No, 0, 1);
    zlJsonPutValue(v_Jtmp, 'rcpdtl_id', r_Info.Rcpdtl_Id, 1, 2);
  
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
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getrefusesendlist;
/

Create Or Replace Procedure Zl_Drugsvr_Getexecutednum
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡҩƷ�ѷ�������
  --��Σ�Json_In:��ʽ
  --  input
  --    billtype            N   1 ��������:1-�շѴ�����ҩ;2-���ʵ�������ҩ;3-���ʱ�����ҩ
  --    rcp_nos             C   1 ���ݺ�:���Դ�����ŵ���
  --    notcontain_zero     N   1 �Ƿ񲻰����ѷ�����Ϊ0�ģ�1-��������0-����
  --    rcpdtl_ids          C   0 ������ϸids�������Ӣ�ĵĶ��ŷָ�,δ����ʱ�����ݺŲ���,����ʱ����ϸid���в���
  --    order_ids           C   0 ҽ��id�������δ����һ��ҽ��id���ŷָ�
  --����: Json_Out,��ʽ����
  --  output
  --    code                 N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message              C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    data[]
  --       rcp_no            C   1   ��������(���õ��ݺ�)
  --       rcpdtl_id         N   1   ������ϸID(����ID)
  --       order_id          N   0   ҽ��id
  --       drug_id           N   1   ҩƷID
  --       sended_num        N   1   �ѷ�����
  ---------------------------------------------------------------------------
  n_������  Number(1);
  n_����    ҩƷ�շ���¼.����%Type;
  v_Nos     Varchar2(4000);
  v_��ϸids Clob;

  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_��ϸids Collection_Type;
  I           Number;

  v_Jtmp      Varchar2(32767);
  c_Order_Ids Clob;
  c_Jtmp      Clob;

  j_Jsonin Pljson;
  j_Json   Pljson;
Begin
  --�������
  j_Jsonin    := Pljson(Json_In);
  j_Json      := j_Jsonin.Get_Pljson('input');
  n_����      := j_Json.Get_Number('billtype');
  v_Nos       := j_Json.Get_String('rcp_nos');
  n_������    := j_Json.Get_Number('notcontain_zero');
  v_��ϸids   := j_Json.Get_Clob('rcpdtl_ids');
  c_Order_Ids := j_Json.Get_Clob('order_ids');

  If n_���� = 1 Then
    n_���� := 8;
  Elsif n_���� = 2 Then
    n_���� := 9;
  Elsif n_���� = 3 Then
    n_���� := 10;
  Elsif c_Order_Ids Is Null Then
    If v_��ϸids Is Null Then
      Json_Out := Zljsonout('����ڵ㡾billtype���������飡');
      Return;
    End If;
  End If;

  If v_��ϸids Is Null Then
    For c_ҩƷ In (Select /*+cardinality(j,10)*/
                  a.No, a.����id, a.ҽ��id, a.ҩƷid, Sum(Nvl(a.����, 1) * Decode(a.�����, Null, 0, a.ʵ������)) As �ѷ�����
                 From ҩƷ�շ���¼ A, Table(f_Str2list(v_Nos)) J
                 Where a.No = j.Column_Value And
                       (a.���� = 8 And n_���� = 8 Or n_���� <> 8 And Instr(',9,10,', ',' || a.���� || ',') > 0 Or n_���� Is Null) And
                       (c_Order_Ids Is Null Or Instr(',' || c_Order_Ids || ',', ',' || a.ҽ��id || ',') > 0)
                 Group By a.No, a.����id, a.ҩƷid, a.ҽ��id) Loop
    
      If Not (Nvl(n_������, 0) = 1 And Nvl(c_ҩƷ.�ѷ�����, 0) = 0) Then
        v_Jtmp := v_Jtmp || ',';
        Zljsonputvalue(v_Jtmp, 'rcp_no', c_ҩƷ.No, 0, 1);
        Zljsonputvalue(v_Jtmp, 'rcpdtl_id', c_ҩƷ.����id, 1);
        Zljsonputvalue(v_Jtmp, 'order_id', c_ҩƷ.ҽ��id, 1);
        Zljsonputvalue(v_Jtmp, 'drug_id', c_ҩƷ.ҩƷid, 1);
        Zljsonputvalue(v_Jtmp, 'sended_num', Nvl(c_ҩƷ.�ѷ�����, 0), 1, 2);
      
        If Length(v_Jtmp) > 30000 Then
          If c_Jtmp Is Null Then
            c_Jtmp := Substr(v_Jtmp, 2);
          Else
            c_Jtmp := c_Jtmp || v_Jtmp;
          End If;
          v_Jtmp := Null;
        End If;
      End If;
    End Loop;
  Else
    I := 0;
    While v_��ϸids Is Not Null Loop
      If Length(v_��ϸids) <= 4000 Then
        Col_��ϸids(I) := v_��ϸids;
        v_��ϸids := Null;
      Else
        Col_��ϸids(I) := Substr(v_��ϸids, 1, Instr(v_��ϸids, ',', 3980) - 1);
        v_��ϸids := Substr(v_��ϸids, Instr(v_��ϸids, ',', 3980) + 1);
      End If;
      I := I + 1;
    End Loop;
  
    I := 0;
    For I In 0 .. Col_��ϸids.Count - 1 Loop
      For c_ҩƷ In (Select /*+cardinality(j,10)*/
                    a.No, a.����id, a.ҩƷid, Sum(Nvl(a.����, 1) * Decode(a.�����, Null, 0, a.ʵ������)) As �ѷ�����
                   From ҩƷ�շ���¼ A, Table(f_Num2list(Col_��ϸids(I))) J
                   Where a.����id = j.Column_Value
                   Group By a.No, a.����id, a.ҩƷid) Loop
      
        If Not (Nvl(n_������, 0) = 1 And Nvl(c_ҩƷ.�ѷ�����, 0) = 0) Then
          v_Jtmp := v_Jtmp || ',';
          Zljsonputvalue(v_Jtmp, 'rcp_no', c_ҩƷ.No, 0, 1);
          Zljsonputvalue(v_Jtmp, 'rcpdtl_id', c_ҩƷ.����id, 1);
          Zljsonputvalue(v_Jtmp, 'drug_id', c_ҩƷ.ҩƷid, 1);
          Zljsonputvalue(v_Jtmp, 'sended_num', Nvl(c_ҩƷ.�ѷ�����, 0), 1, 2);
        
          If Length(v_Jtmp) > 30000 Then
            If c_Jtmp Is Null Then
              c_Jtmp := Substr(v_Jtmp, 2);
            Else
              c_Jtmp := c_Jtmp || v_Jtmp;
            End If;
            v_Jtmp := Null;
          End If;
        End If;
      End Loop;
    End Loop;
  End If;

  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getexecutednum;
/

Create Or Replace Procedure Zl_Drugsvr_Getsendwindows
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡҩƷ�ķ�ҩ����
  --��Σ�Json_In:��ʽ
  --  input
  --    item_list[]
  --        billtype        N   1   ��������: 1 -�շѴ�����ҩ  ;2- ���ʵ�������ҩ;3- ���ʱ�����ҩ
  --        pharmacy_id     N   1   ҩ��id
  --        pati_id         N   1   ����id
  --        valid_days      N       δ��ҩƷ��¼��ѯ��Χ����Ч����
  --        defaultwindow   C       ���ݸ�ҵ��ģ����������д����ȱʡ����
  --����: Json_Out,��ʽ����
  --  output
  --    code                N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    item_list                   [����]ÿ���ⷿID��Ӧ�ķ�ҩ����
  --        pharmacy_id     N   1   �ⷿID
  --        pharmacy_window C   1   �ⷿID��Ӧ�ķ�ҩ����
  --���ڻ�ȡ����:
  --  �ж�ָ��������ָ��ҩ����δ��ҩƷ��¼���Ƿ���������ϰ�ķ�ҩ����
  --  a.��ҩ���ڴ��ڣ�����������������ķ�ҩ����
  --  b.��ҩ���ڲ����ڣ�
  --    i:�������ȱʡ�ķ�ҩ���ڣ��������ϰ࣬�򷵻�ȱʡ�ķ�ҩ���ڣ��������δ�ϰ��򷵻�null
  --    ii:���������ȱʡ�ķ�ҩ���ڣ�����ݶ�̬�������0-��æ;1-ƽ������ȡ��ר�ҵķ�ҩ����
  ---------------------------------------------------------------------------
  j_Jsonlist Pljson_List;

  n_����     ҩƷ�շ���¼.����%Type;
  n_�ⷿid   ҩƷ�շ���¼.�ⷿid%Type;
  v_ȱʡ���� δ��ҩƷ��¼.��ҩ����%Type;
  v_��ҩ���� δ��ҩƷ��¼.��ҩ����%Type;
  n_����id   Number(18);
  n_��Ч���� Number(10);

  Type t_Record Is Record(
    ҩ��id   Number(18),
    ��ҩ���� Varchar2(10));

  Type t_��ҩ���� Is Table Of t_Record;
  c_��ҩ���� t_��ҩ���� := t_��ҩ����();

  v_List  Varchar2(32767);
  n_Count Number;

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --�������
  j_Jsonin   := PLJson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  j_Jsonlist := j_Json.Get_Pljson_List('item_list');

  n_Count := j_Jsonlist.Count;
  If j_Jsonlist.Count = 0 Then
    Json_Out := zlJsonOut('δ����ⷿ��Ϣ��');
    Return;
  End If;

  For I In 1 .. j_Jsonlist.Count Loop
    j_Json     := PLJson();
    j_Json     := PLJson(j_Jsonlist.Get(I));
    n_����     := j_Json.Get_Number('billtype');
    n_�ⷿid   := j_Json.Get_Number('pharmacy_id');
    n_����id   := j_Json.Get_Number('pati_id');
    n_��Ч���� := j_Json.Get_Number('valid_days');
    v_ȱʡ���� := j_Json.Get_String('defaultwindow');
    If Nvl(n_��Ч����, 0) = 0 Then
      n_��Ч���� := 7;
    End If;
  
    --0)�жϸ�ҩ���Ƿ��ѷ����˷�ҩ����
    v_��ҩ���� := Null;
    For I In 1 .. c_��ҩ����.Count Loop
      If c_��ҩ����(I).ҩ��id = Nvl(n_�ⷿid, 0) Then
        v_��ҩ���� := c_��ҩ����(I).��ҩ����;
        Exit;
      End If;
    End Loop;
  
    If v_��ҩ���� Is Null Then
      --1)����ָ��������ָ��ҩ���з����˷�ҩ�����Ҹô��ڴ����ϰ�����һ��δ��ҩƷ�Ĵ��ڣ�������ȡ�÷�ҩ����
      Select Max(��ҩ����)
      Into v_��ҩ����
      From (Select a.��ҩ����
             From δ��ҩƷ��¼ A
             Where a.���� = Decode(n_����, 1, 8, 2, 9, 10) And a.����id = n_����id And a.�������� Between Trunc(Sysdate) - n_��Ч���� - 1 And
                   Sysdate And a.�ⷿid = n_�ⷿid And a.��ҩ���� Is Not Null And Exists
              (Select 1 From ��ҩ���� Where Nvl(�ϰ��, 0) = 1 And ���� = a.��ҩ���� And ҩ��id = a.�ⷿid)
             Order By a.�������� Desc)
      Where Rownum < 2;
    
      If v_��ҩ���� Is Null Then
        --2)���ȱʡ�����Ƿ����ϰ�ģ����ϰ�ģ�ȡ�÷�ҩ����
        If v_ȱʡ���� Is Not Null Then
          Select Count(1)
          Into n_Count
          From ��ҩ����
          Where Nvl(�ϰ��, 0) = 1 And ���� = v_ȱʡ���� And ҩ��id = n_�ⷿid;
          If n_Count <> 0 Then
            v_��ҩ���� := v_ȱʡ����;
          End If;
        Else
          --3)����ҩ�����ڵķ������æ��/ƽ�� ����ȡ��ҩ����
          v_��ҩ���� := Zl_Get��ҩ����(n_�ⷿid);
        End If;
      End If;
    
      If v_��ҩ���� Is Not Null Then
        c_��ҩ����.Extend;
        c_��ҩ����(c_��ҩ����.Count).ҩ��id := n_�ⷿid;
        c_��ҩ����(c_��ҩ����.Count).��ҩ���� := v_��ҩ����;
      End If;
    End If;
  End Loop;

  For I In 1 .. c_��ҩ����.Count Loop
    zlJsonPutValue(v_List, 'pharmacy_id', c_��ҩ����(I).ҩ��id, 1, 1);
    zlJsonPutValue(v_List, 'pharmacy_window', c_��ҩ����(I).��ҩ����, 0, 2);
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || v_List || ']}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getsendwindows;
/

Create Or Replace Procedure Zl_Drugsvr_Getpharmacywindows
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡҩ�����漰�ķ�ҩ����
  --��Σ�Json_In:��ʽ
  --  input
  --    pharmacy_ids            C   1  ҩ��ID1,ҩ��ID2��
  --����: Json_Out,��ʽ����
  --  output
  --    code                    N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message                 C   1   ÿ��ҩ��id��Ӧ�ķ�ҩ����[����]
  --    window_list[]    ���������б�[����]
  --        pharmacy_id             N 1 ҩ��ID
  --        pharmacy_window         C 1 ��ҩ����
  --        expert_window           N 1 �Ƿ�ר�Ҵ��ڣ�1-�ǣ�0-����
  ---------------------------------------------------------------------------
  v_ҩ��ids Varchar2(32767);
  v_Temp    Varchar2(32767);

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --�������
  j_Jsonin  := PLJson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  v_ҩ��ids := j_Json.Get_String('pharmacy_ids');

  If v_ҩ��ids Is Null Then
    Json_Out := zlJsonOut('δ����ҩ����Ϣ');
    Return;
  End If;

  For c_ҩƷ In (Select /*+cardinality(b,10)*/
                a.ҩ��id, a.����, Nvl(a.ר��, 0) As ר��
               From ��ҩ���� A, Table(f_Num2List(v_ҩ��ids)) B
               Where a.ҩ��id = b.Column_Value
               Order By a.ҩ��id, a.����) Loop
  
    zlJsonPutValue(v_Temp, 'pharmacy_id', c_ҩƷ.ҩ��id, 1, 1);
    zlJsonPutValue(v_Temp, 'pharmacy_window', c_ҩƷ.����, 0);
    zlJsonPutValue(v_Temp, 'expert_window', c_ҩƷ.ר��, 1, 2);
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","window_list":[' || v_Temp || ']}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getpharmacywindows;
/

Create Or Replace Procedure Zl_Drugsvr_Getstockcheck
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡҩƷ����鷽ʽ
  --��Σ�Json_In:��ʽ
  --  input
  --����: Json_Out,��ʽ����
  --  output
  --    code               N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message            C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    item_list                   [����]
  --       pharmacy_id     N   1   ҩ��ID
  --       check_type      N   1   ��鷽ʽ��0-����飬1-�����ʾ����2-����ֹ
  --------------------------------------------------------------------------- 
  v_Output Varchar2(32767);
Begin

  For r_Data In (Select Distinct b.����id, Nvl(c.��鷽ʽ, 0) As ��鷽ʽ
                 From ��������˵�� B, ҩƷ������ C
                 Where b.����id = c.�ⷿid(+) And b.������� In (1, 2, 3) And b.�������� In ('��ҩ��', '��ҩ��', '��ҩ��')) Loop
  
    zlJsonPutValue(v_Output, 'pharmacy_id', r_Data.����id, 1, 1);
    zlJsonPutValue(v_Output, 'check_type', r_Data.��鷽ʽ, 1, 2);
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || v_Output || ']}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getstockcheck;
/

Create Or Replace Procedure Zl_Drugsvr_Getstock
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡָ��ҩƷ��ָ���ⷿ�Ŀ��ÿ����
  --��Σ�Json_In:��ʽ
  --  input
  --    drug_id             N   1   ҩƷID
  --    pharmacy_ids        C   1   �ⷿID
  --    batch               N       ���Σ�<=0-���������Σ�>0ֻ��ĳ����
  --����: Json_Out,��ʽ����
  --  output
  --    code                N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    data                N  1  ���ÿ��
  ---------------------------------------------------------------------------
  n_ҩƷid   ҩƷ�շ���¼.ҩƷid%Type;
  n_������� ҩƷ���.�������� %Type;
  v_�ⷿids  Varchar2(32767);

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --�������
  j_Jsonin  := PLJson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  n_ҩƷid  := j_Json.Get_Number('drug_id');
  v_�ⷿids := j_Json.Get_String('pharmacy_ids');

  If Nvl(n_ҩƷid, 0) = 0 Then
    Json_Out := zlJsonOut('δ�������ҩƷ��Ϣ');
    Return;
  End If;

  --��ȡ���(�����������),ҩ��������(����=0,����Ϊҩ��)����Ч��
  Select Nvl(Sum(a.��������), 0)
  Into n_�������
  From ҩƷ��� A
  Where (Nvl(a.����, 0) = 0 Or a.Ч�� Is Null Or a.Ч�� > Trunc(Sysdate)) And a.���� = 1 And a.ҩƷid = n_ҩƷid And
        Instr(',' || v_�ⷿids || ',', ',' || a.�ⷿid || ',') > 0;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":' || zlJsonStr(n_�������, 1) || '}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getstock;
/

Create Or Replace Procedure Zl_Drugsvr_Getstockbydept
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ�����ָ��ҩƷ���ⷿ���ʻ�ȡ���ⷿ�Ŀ����Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --    drug_id             N   1   ҩƷID
  --    pharmacy_nature     C   1  �ⷿ���ʣ���ҩ������ҩ������ҩ����
  --����: Json_Out,��ʽ����
  --  output
  --    code                N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    data
  --        pharmacy_id     N   1   ҩ��ID
  --        pharmacy_code   C   1   �ⷿ����
  --        pharmacy_name   C   1   �ⷿ����
  --        stock           N   1   ��������
  ---------------------------------------------------------------------------
  n_ҩƷid   ҩƷ�շ���¼.ҩƷid%Type;
  v_�ⷿ���� Varchar2(50);
  v_Output   Varchar2(32767);

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --�������
  j_Jsonin   := PLJson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_ҩƷid   := j_Json.Get_Number('drug_id');
  v_�ⷿ���� := j_Json.Get_String('pharmacy_nature');

  If Nvl(n_ҩƷid, 0) = 0 Or Nvl(v_�ⷿ����, '-') = '-' Then
    Json_Out := zlJsonOut('δ�������ҩƷ��Ϣ');
    Return;
  End If;

  For c_��� In (Select b.����, b.����, a.�ⷿid, Nvl(Sum(a.��������), 0) As ���
               From ҩƷ��� A,
                    (Select Distinct a.Id, a.����, a.����
                      From ���ű� A, ��������˵�� B
                      Where a.Id = b.����id And Instr(',' || v_�ⷿ���� || ',', ',' || b.�������� || ',') > 0) B
               Where a.�ⷿid = b.Id And ((a.Ч�� Is Null Or Ч�� > Trunc(Sysdate)) Or Nvl(a.����, 0) = 0) And a.���� = 1 And
                     a.ҩƷid = n_ҩƷid
               Group By b.����, b.����, a.�ⷿid
               Having Sum(Nvl(a.��������, 0)) <> 0
               Order By b.����) Loop
  
    zlJsonPutValue(v_Output, 'pharmacy_id', c_���.�ⷿid, 1, 1);
    zlJsonPutValue(v_Output, 'pharmacy_code', c_���.����);
    zlJsonPutValue(v_Output, 'pharmacy_name', c_���.����);
    zlJsonPutValue(v_Output, 'stock', c_���.���, 1, 2);
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":[' || v_Output || ']}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getstockbydept;
/

CREATE OR REPLACE Procedure Zl_Drugsvr_Getstockbatch
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ�������ȡ���ҩƷ��漰�۸���Ϣ:����Ŀѡ������չʾ��漰�۸���Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --   drug_ids                   C 1 ҩƷID�������Ӣ�ĵĶ��ŷָ�
  --   pharmacy_ids               C 0 �ⷿID�������Ӣ�ĵĶ��ŷָ�;���ַ���,��ѯ���пⷿ
  --   return_price               N 0 �Ƿ񷵻��ۼۣ�1-���ؼ۸���Ϣ(�ۼ�);0-������
  --   return_dept                N 0 �����ҷ��ؿ�棺1-�����ҷ��ؿ��;0-��ҩƷ���ؿ��;2-���ؿ�������ҩƷ�Ŀ��
  --   query_type                 N 1 ��ѯ����:�磺0-��ѯ��治����0,1-��ѯ���С�ڵ���0
  --����: Json_Out,��ʽ����
  --  output
  --    code                      N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                   C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    data[]
  --      drug_id                 N 1 ҩƷID
  --      pharmacy_id             N 1 �ⷿID(�����ҷ��ؿ����д���)
  --      stock                   N 1 ��������
  --      price                   N 1 ���ۼ�(���ؼ۸�ʱ���д���)
  ---------------------------------------------------------------------------
  v_ҩƷids  Clob;
  v_�ⷿids  Varchar2(32767);
  n_���ؼ۸� Number(2);
  n_���ҷ��� Number(2);
  n_��ѯ���� Number(2);

  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_ҩƷids Collection_Type;
  I           Number;

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --�������
  j_Jsonin   := PLJson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  v_ҩƷids  := j_Json.Get_Clob('drug_ids');
  v_�ⷿids  := j_Json.Get_String('pharmacy_ids');
  n_���ؼ۸� := Nvl(j_Json.Get_Number('return_price'), 0);
  n_���ҷ��� := j_Json.Get_Number('return_dept');
  n_��ѯ���� := Nvl(j_Json.Get_Number('query_type'), 0);

  I := 0;
  While v_ҩƷids Is Not Null Loop
    If Length(v_ҩƷids) <= 4000 Then
      Col_ҩƷids(I) := v_ҩƷids;
      v_ҩƷids := Null;
    Else
      Col_ҩƷids(I) := Substr(v_ҩƷids, 1, Instr(v_ҩƷids, ',', 3980) - 1);
      v_ҩƷids := Substr(v_ҩƷids, Instr(v_ҩƷids, ',', 3980) + 1);
    End If;
    I := I + 1;
  End Loop;

  I := 0;
  If Nvl(n_���ؼ۸�, 0) = 0 Then
    If Nvl(n_���ҷ���, 0) = 0 Then
      For I In 0 .. Col_ҩƷids.Count - 1 Loop
        If n_��ѯ���� = 0 Then
          For c_��� In (With c_ҩƷ��Ϣ As
                          (Select Column_Value As ҩƷid From Table(f_Num2List(Col_ҩƷids(I))))
                         Select /*+cardinality(b,10)*/
                          a.ҩƷid, Sum(Nvl(a.��������, 0)) As ���
                         From ҩƷ��� A, c_ҩƷ��Ϣ B
                         Where a.ҩƷid = b.ҩƷid And a.���� = 1 And
                               (Instr(',' || v_�ⷿids || ',', ',' || a.�ⷿid || ',') > 0 Or v_�ⷿids Is Null) And
                               (Nvl(a.����, 0) = 0 Or a.Ч�� Is Null Or a.Ч�� > Trunc(Sysdate))
                         Group By a.ҩƷid
                         Having Sum (Nvl(a.��������, 0)) <> 0) Loop

            v_Jtmp := v_Jtmp || ',';
            zlJsonPutValue(v_Jtmp, 'drug_id', c_���.ҩƷid, 1, 1);
            zlJsonPutValue(v_Jtmp, 'stock', c_���.���, 1, 2);

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
          For c_��� In (With c_ҩƷ��Ϣ As
                          (Select Column_Value As ҩƷid From Table(f_Num2List(Col_ҩƷids(I))))
                         Select /*+cardinality(b,10)*/
                          a.ҩƷid, Sum(Nvl(a.��������, 0)) As ���
                         From ҩƷ��� A, c_ҩƷ��Ϣ B
                         Where a.ҩƷid = b.ҩƷid And a.���� = 1 And
                               (Instr(',' || v_�ⷿids || ',', ',' || a.�ⷿid || ',') > 0 Or v_�ⷿids Is Null) And
                               (Nvl(a.����, 0) = 0 Or a.Ч�� Is Null Or a.Ч�� > Trunc(Sysdate))
                         Group By a.ҩƷid
                         Having Sum (Nvl(a.��������, 0)) <= 0) Loop

            v_Jtmp := v_Jtmp || ',';
            zlJsonPutValue(v_Jtmp, 'drug_id', c_���.ҩƷid, 1, 1);
            zlJsonPutValue(v_Jtmp, 'stock', c_���.���, 1, 2);

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
      End Loop;
    Elsif Nvl(n_���ҷ���, 0) = 1 Then
      For I In 0 .. Col_ҩƷids.Count - 1 Loop
        For c_��� In (Select /*+cardinality(b,10)*/
                      a.ҩƷid, Sum(Nvl(a.��������, 0)) As ���, a.�ⷿid
                     From ҩƷ��� A, Table(f_Num2List(Col_ҩƷids(I))) B
                     Where a.ҩƷid = b.Column_Value And a.���� = 1 And
                           (Instr(',' || v_�ⷿids || ',', ',' || a.�ⷿid || ',') > 0 Or v_�ⷿids Is Null) And
                           (Nvl(a.����, 0) = 0 Or a.Ч�� Is Null Or a.Ч�� > Trunc(Sysdate))
                     Group By a.ҩƷid, a.�ⷿid
                     Having Sum(Nvl(a.��������, 0)) <> 0) Loop

          v_Jtmp := v_Jtmp || ',';
          zlJsonPutValue(v_Jtmp, 'drug_id', c_���.ҩƷid, 1, 1);
          zlJsonPutValue(v_Jtmp, 'pharmacy_id', c_���.�ⷿid, 1);
          zlJsonPutValue(v_Jtmp, 'stock', c_���.���, 1, 2);

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
    Elsif Nvl(n_���ҷ���, 0) = 2 Then
      For c_��� In (Select a.ҩƷid, Sum(Nvl(a.��������, 0)) As ���, a.�ⷿid
                   From ҩƷ��� A,
                        (Select /*+cardinality(c,10)*/
                           Column_Value As �ⷿid
                          From Table(f_Num2List(Nvl(v_�ⷿids, 0)))) C
                   Where (Nvl(����, 0) = 0 Or Ч�� Is Null Or Ч�� > Trunc(Sysdate)) And ���� = 1 And a.�ⷿid = c.�ⷿid
                   Group By a.ҩƷid, a.�ⷿid
                   Having Sum(Nvl(a.��������, 0)) <> 0) Loop

        v_Jtmp := v_Jtmp || ',';
        zlJsonPutValue(v_Jtmp, 'drug_id', c_���.ҩƷid, 1, 1);
        zlJsonPutValue(v_Jtmp, 'pharmacy_id', c_���.�ⷿid, 1);
        zlJsonPutValue(v_Jtmp, 'stock', c_���.���, 1, 2);

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
      Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":[' || Substr(v_Jtmp, 2) || ']}}';
    Else
      c_Jtmp   := c_Jtmp || v_Jtmp;
      Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":[' || c_Jtmp || ']}}';
    End If;
    Return;
  End If;

  --�����۸�
  If Nvl(n_���ҷ���, 0) = 0 Then
    For I In 0 .. Col_ҩƷids.Count - 1 Loop
      For c_��� In (Select a.ҩƷid, Nvl(a.���, 0) As ���, Decode(Nvl(b.�Ƿ���, 0), 1, 0, Nvl(c.�ּ�, 0)) As �۸�
                   From (With c_ҩƷ��Ϣ As (Select Column_Value As ҩƷid From Table(f_Num2List(Col_ҩƷids(I))))
                          Select /*+cardinality(b,10)*/
                           a.ҩƷid, Sum(Nvl(a.��������, 0)) As ���
                          From ҩƷ��� A, c_ҩƷ��Ϣ B
                          Where a.ҩƷid = b.ҩƷid And a.���� = 1 And
                                (Instr(',' || v_�ⷿids || ',', ',' || a.�ⷿid || ',') > 0 Or v_�ⷿids Is Null) And
                                (Nvl(a.����, 0) = 0 Or a.Ч�� Is Null Or a.Ч�� > Trunc(Sysdate))
                          Group By a.ҩƷid
                          Having Sum(Nvl(a.��������, 0)) <> 0) A, �շ���ĿĿ¼ B, �շѼ�Ŀ C
                          Where a.ҩƷid = c.�շ�ϸĿid And a.ҩƷid = b.Id And c.�۸�ȼ� Is Null And Sysdate Between c.ִ������ And
                                Nvl(c.��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))
                   ) Loop

        v_Jtmp := v_Jtmp || ',';
        zlJsonPutValue(v_Jtmp, 'drug_id', c_���.ҩƷid, 1, 1);
        zlJsonPutValue(v_Jtmp, 'pharmacy_id', 0, 1);
        zlJsonPutValue(v_Jtmp, 'stock', c_���.���, 1);
        zlJsonPutValue(v_Jtmp, 'price', c_���.�۸�, 1, 2);

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
  Elsif n_���ҷ��� = 1 Then

    For I In 0 .. Col_ҩƷids.Count - 1 Loop
      For c_��� In (Select a.ҩƷid, Nvl(a.���, 0) As ���, Decode(Nvl(b.�Ƿ���, 0), 1, 0, Nvl(c.�ּ�, 0)) As �۸�, a.�ⷿid
                   From (With c_ҩƷ��Ϣ As (Select Column_Value As ҩƷid From Table(f_Num2List(Col_ҩƷids(I))))
                          Select /*+cardinality(b,10)*/
                           a.ҩƷid, Sum(Nvl(a.��������, 0)) As ���, a.�ⷿid
                          From ҩƷ��� A, c_ҩƷ��Ϣ B
                          Where a.ҩƷid = b.ҩƷid And a.���� = 1 And
                                (Instr(',' || v_�ⷿids || ',', ',' || a.�ⷿid || ',') > 0 Or v_�ⷿids Is Null) And
                                (Nvl(a.����, 0) = 0 Or a.Ч�� Is Null Or a.Ч�� > Trunc(Sysdate))
                          Group By a.ҩƷid, a.�ⷿid
                          Having Sum(Nvl(a.��������, 0)) <> 0) A, �շ���ĿĿ¼ B, �շѼ�Ŀ C
                          Where a.ҩƷid = c.�շ�ϸĿid And a.ҩƷid = b.Id And c.�۸�ȼ� Is Null And Sysdate Between c.ִ������ And
                                Nvl(c.��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))
                   ) Loop

        v_Jtmp := v_Jtmp || ',';
        zlJsonPutValue(v_Jtmp, 'drug_id', c_���.ҩƷid, 1, 1);
        zlJsonPutValue(v_Jtmp, 'pharmacy_id', c_���.�ⷿid, 1);
        zlJsonPutValue(v_Jtmp, 'stock', c_���.���, 1);
        zlJsonPutValue(v_Jtmp, 'price', c_���.�۸�, 1, 2);

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

  Elsif n_���ҷ��� = 2 Then
    For c_��� In (Select a.ҩƷid, Nvl(a.���, 0) As ���, Decode(Nvl(b.�Ƿ���, 0), 1, 0, Nvl(c.�ּ�, 0)) As �۸�, a.�ⷿid
                 From (Select a.ҩƷid, Sum(Nvl(a.��������, 0)) As ���, a.�ⷿid
                        From ҩƷ��� A,
                             (Select /*+cardinality(c,10)*/
                                Column_Value As �ⷿid
                               From Table(f_Num2List(v_�ⷿids))) C
                        Where (Nvl(����, 0) = 0 Or Ч�� Is Null Or Ч�� > Trunc(Sysdate)) And ���� = 1 And a.�ⷿid = c.�ⷿid
                        Group By a.ҩƷid, a.�ⷿid
                        Having Sum(Nvl(a.��������, 0)) <> 0) A, �շ���ĿĿ¼ B, �շѼ�Ŀ C
                 Where a.ҩƷid = c.�շ�ϸĿid And a.ҩƷid = b.Id And c.�۸�ȼ� Is Null And Sysdate Between c.ִ������ And
                       Nvl(c.��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))) Loop

      v_Jtmp := v_Jtmp || ',';
      zlJsonPutValue(v_Jtmp, 'drug_id', c_���.ҩƷid, 1, 1);
      zlJsonPutValue(v_Jtmp, 'pharmacy_id', c_���.�ⷿid, 1);
      zlJsonPutValue(v_Jtmp, 'stock', c_���.���, 1);
      zlJsonPutValue(v_Jtmp, 'price', c_���.�۸�, 1, 2);

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
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getstockbatch;
/

Create Or Replace Procedure Zl_Drugsvr_Getstockbydrugname
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ�����ҩ����Ʒ�֣�����ȡ�����Ŀ��ÿ��
  --��Σ�Json_In:��ʽ
  --  input 
  --    clinicdrug_id   N   1   ҩ��ID
  --    pharmacy_id     N   1   ҩ��ID��ҩ��ID=0����ʾ���е�
  --    occasion        N   1   ���ϣ�1-���� ��2-סԺ
  --    show_unit       N   1   ��ʾ��λ:0-�ۼ۵�λ;1-סԺ��λ;2-���ﵥλ
  --    site_no         C   1   վ���
  --����: Json_Out,��ʽ����
  --  output
  --    code            N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message         C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    data[]
  --      drug_code               ҩƷ����
  --      drug_name               ҩƷ����
  --      drug_spec               ���
  --      unit                    ��λ
  --      drug_id         N   1   ҩƷID
  --      pharmacy_name           ҩ������
  --      pharmacy_id     N   1   ҩ��ID
  --      stock           N   1   ��������
  ---------------------------------------------------------------------------
  n_ҩ��id   ҩƷ���.ҩƷid%Type;
  n_�ⷿid   ҩƷ���.�ⷿid%Type;
  n_����     Number(2);
  n_��ʾ��λ Number(2);
  v_վ���   Varchar2(6);

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --�������
  j_Jsonin   := PLJson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_ҩ��id   := j_Json.Get_Number('clinicdrug_id');
  n_�ⷿid   := Nvl(j_Json.Get_Number('pharmacy_id'), 0);
  n_����     := j_Json.Get_Number('occasion');
  n_��ʾ��λ := j_Json.Get_Number('show_unit');
  v_վ���   := j_Json.Get_String('site_no');

  If Nvl(n_����, 0) = 0 Then
    n_���� := 1;
  End If;
  If n_��ʾ��λ Is Null Then
  
    n_��ʾ��λ := 0;
  End If;

  v_Jtmp := Null;
  For c_��� In (Select d.����, d.���, d.����, e.���� As ҩ��, Max(Decode(n_��ʾ��λ, 0, d.���㵥λ, 1, a.���ﵥλ, a.סԺ��λ)) As ��λ, 1 As סԺ��װ,
                      Sum(Nvl(m.��������, 0) / Decode(n_��ʾ��λ, 0, 1, 1, a.�����װ, a.סԺ��װ)) As ��������, m.�ⷿid, m.ҩƷid
               From ҩƷ��� M, ҩƷ��� A, �շ���ĿĿ¼ D, ���ű� E
               Where m.ҩƷid = d.Id And m.ҩƷid = a.ҩƷid And m.�ⷿid = e.Id And (m.�ⷿid = n_�ⷿid Or n_�ⷿid = 0) And
                     (Nvl(m.����, 0) = 0 Or m.Ч�� Is Null Or m.Ч�� > Trunc(Sysdate)) And a.ҩ��id = n_ҩ��id And
                     (d.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or d.����ʱ�� Is Null) And d.������� In (n_����, 3) And
                     (d.վ�� = v_վ��� Or d.վ�� Is Null)
               Group By e.����, d.����, d.���, d.����, d.���㵥λ, m.�ⷿid, m.ҩƷid
               Having Sum(Nvl(m.��������, 0)) > 0
               Order By d.����) Loop
  
    v_Jtmp := v_Jtmp || ',';
    zlJsonPutValue(v_Jtmp, 'drug_code', c_���.����, 0, 1);
    zlJsonPutValue(v_Jtmp, 'drug_name', c_���.����);
    zlJsonPutValue(v_Jtmp, 'drug_spec', c_���.���);
    zlJsonPutValue(v_Jtmp, 'unit', c_���.��λ);
    zlJsonPutValue(v_Jtmp, 'drug_ide', c_���.ҩƷid, 1);
    zlJsonPutValue(v_Jtmp, 'pharmacy_name', c_���.ҩ��);
    zlJsonPutValue(v_Jtmp, 'pharmacy_id', c_���.�ⷿid, 1);
    zlJsonPutValue(v_Jtmp, 'stock', Round(c_���.��������, 5), 1, 2);
  
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
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getstockbydrugname;
/

Create Or Replace Procedure Zl_Drugsvr_Batchgetprice
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ�������ȡҩƷ�ۼ�(������ʹ��)
  --��Σ�Json_In:��ʽ
  --  input
  --   drug_ids    C   1   ҩƷID�������Ӣ�ĵĶ��ŷָ�
  --����: Json_Out,��ʽ����
  --  output
  --    code          N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message         C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    data
  --      drug_id N   1   ҩƷID
  --      price   N   1   ���ۼ�(���ؼ۸�ʱ���д���)
  ---------------------------------------------------------------------------
  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_ҩƷids Collection_Type;
  I           Integer;
  c_ҩƷids   Clob;

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --�������
  j_Jsonin  := PLJson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  c_ҩƷids := j_Json.Get_Clob('drug_ids');
  If c_ҩƷids Is Null Then
    Json_Out := zlJsonOut('δ������Ч��ҩƷid,����!');
  End If;

  I := 0;
  While c_ҩƷids Is Not Null Loop
    If Length(c_ҩƷids) <= 4000 Then
      Col_ҩƷids(I) := c_ҩƷids;
      c_ҩƷids := Null;
    Else
      Col_ҩƷids(I) := Substr(c_ҩƷids, 1, Instr(c_ҩƷids, ',', 3980) - 1);
      c_ҩƷids := Substr(c_ҩƷids, Instr(c_ҩƷids, ',', 3980) + 1);
    End If;
    I := I + 1;
  End Loop;

  v_Jtmp := Null;
  For I In 0 .. Col_ҩƷids.Count - 1 Loop
    --�����۸�
    For c_��� In (With c_ҩƷ��Ϣ As
                    (Select /*+cardinality(D,10)*/
                     d.Column_Value As ҩƷid
                    From Table(f_Num2List(Col_ҩƷids(I))) D)
                   Select a.ҩƷid, Sum(a.ʵ�ʽ��) / Sum(a.ʵ������) As �۸�
                   From ҩƷ��� A, c_ҩƷ��Ϣ B
                   Where a.ҩƷid = b.ҩƷid And a.���� = 1 And (Nvl(a.����, 0) = 0 Or a.Ч�� Is Null Or a.Ч�� > Trunc(Sysdate))
                   Group By a.ҩƷid
                   Having Sum (Nvl(a.ʵ������, 0)) <> 0) Loop
    
      v_Jtmp := v_Jtmp || ',';
      zlJsonPutValue(v_Jtmp, 'drug_id', c_���.ҩƷid, 1, 1);
      zlJsonPutValue(v_Jtmp, 'price', c_���.�۸�, 1, 2);
    
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
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Batchgetprice;
/

Create Or Replace Procedure Zl_Drugsvr_Checkmedicinesended
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ָ��ҽ�������һ�η����Ƿ��ѷ�ҩ
  --��Σ�Json_In:��ʽ
  --  input
  --    item_list[]     ���͵�ҽ���б�
  --       order_id       N 1 ҽ��ID
  --       rcpno          C 1 ���ݺ�
  --����: Json_Out,��ʽ����
  --  output
  --    code              N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message           C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    data              N 1 �Ƿ��ѷ�ҩ 0-δ��ҩ��1-�ѷ�ҩ
  ---------------------------------------------------------------------------
  j_Json_Tmp Pljson;
  j_Jsonlist Pljson_List;
  n_Tmp      Number;
  v_��ϸ     Varchar2(32767);

  j_Jsonin Pljson;
  j_Json   Pljson;
Begin
  --�������
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  j_Jsonlist := j_Json.Get_Pljson_List('item_list');

  For I In 1 .. j_Jsonlist.Count Loop
    j_Json_Tmp := Pljson(j_Jsonlist.Get(I));
    v_��ϸ     := v_��ϸ || ',' || j_Json_Tmp.Get_Number('order_id');
    v_��ϸ     := v_��ϸ || ':' || j_Json_Tmp.Get_String('rcpno');
  End Loop;

  If v_��ϸ Is Not Null Then
    Select Count(1)
    Into n_Tmp
    From ҩƷ�շ���¼ A,
         (Select /*+cardinality(b,10)*/
            To_Number(C1) As ҽ��id, C2 As NO
           From Table(Cast(f_Str2list2(Substr(v_��ϸ, 2)) As t_Strlist2)) B) B
    Where a.No = b.No And a.ҽ��id = b.ҽ��id And a.���� In (9, 10) And Mod(a.��¼״̬, 3) = 1 And a.����� Is Null And Rownum < 2;
  End If;

  If Nvl(n_Tmp, 0) > 0 Then
    n_Tmp := 0;
  Else
    n_Tmp := 1;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":' || n_Tmp || '}}';
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Checkmedicinesended;
/

Create Or Replace Procedure Zl_Drugsvr_Autosplitspeci
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ�����в�ҩ����ҩ����Ʒ�֣����Զ�����ҩƷ(�Զ��ֽ������)
  --��Σ�Json_In:��ʽ
  --  input
  --    clinic_drug_id    N 1 ҩ��ID
  --    form              N 1 ��̬��0-ɢװ;1-��ҩ��Ƭ;2-����
  --    quantity          N 1 ��������������λ����
  --    packages_num      N 1 ����
  --    pharmacy_id       N 1 ҩ��ID
  --    occasion          N 1 ���ϣ�1-���� ��2-סԺ
  --    drug_ids          C 0 ָ��ҩƷ�ķ���
  --����: Json_Out,��ʽ����
  --  output
  --    code              N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message           C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    data{}
  --       quantity_remain   N   0   ʣ������
  --       item_list[] ���ܷ���ʱ�ýڵ㷵�ؿ�
  --           drug_id           N   1   ҩƷID
  --           quantity          N   1   ����
  ---------------------------------------------------------------------------
  n_ҩ��id   ҩƷ���.ҩ��id%Type;
  n_��̬     ҩƷ���.��ҩ��̬%Type;
  n_����     ҩƷ���.��������%Type;
  n_����     ҩƷ���.��������%Type;
  n_ҩ��id   ҩƷ���.�ⷿid%Type;
  n_����     Number(1);
  n_ҩƷid   ҩƷ���.ҩƷid%Type;
  v_ҩƷids  Varchar2(4000);
  n_ʣ������ ҩƷ���.��������%Type;
  v_Result   Varchar2(4000);
  v_ҩƷ��Ϣ Varchar2(200);
  v_Json     Varchar2(32767);

  j_Jsonin Pljson;
  j_Json   Pljson;
Begin
  --�������
  j_Jsonin  := Pljson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  n_ҩ��id  := j_Json.Get_Number('clinic_drug_id');
  n_��̬    := j_Json.Get_Number('form');
  n_����    := j_Json.Get_Number('quantity');
  n_����    := j_Json.Get_Number('packages_num');
  n_ҩ��id  := j_Json.Get_Number('pharmacy_id');
  n_����    := j_Json.Get_Number('occasion');
  v_ҩƷids := j_Json.Get_String('drug_ids');

  v_Result := Zl_Dispensechspecs(n_ҩ��id, n_��̬, n_����, n_����, n_ҩ��id, 0, n_����, v_ҩƷids);
  If v_Result Is Null Then
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":{"quantity_remain":0,"item_list":[]}}}';
  Else
    If Instr(v_Result, '|') > 0 Then
      n_ʣ������ := To_Number(Substr(v_Result, Instr(v_Result, '|') + 1, Length(v_Result)));
      v_Result   := Substr(v_Result, 1, Instr(v_Result, '|') - 1);
    End If;
    For r_Row In (Select /*+cardinality(x,10)*/
                   x.Column_Value As ҩƷ��Ϣ
                  From Table(Cast(f_Str2list(v_Result, ';') As t_Strlist)) X) Loop
      v_ҩƷ��Ϣ := r_Row.ҩƷ��Ϣ;
      --�ֽ�
      n_ҩƷid   := To_Number(Substr(v_ҩƷ��Ϣ, 1, Instr(v_ҩƷ��Ϣ, ',') - 1));
      v_ҩƷ��Ϣ := Substr(v_ҩƷ��Ϣ, Instr(v_ҩƷ��Ϣ, ',') + 1);
      n_����     := To_Number(v_ҩƷ��Ϣ);
      v_Json     := v_Json || ',{"drug_id":' || n_ҩƷid;
      v_Json     := v_Json || ',"quantity":' || Nvl(n_����, 0);
      v_Json     := v_Json || '}';
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":{';
    Json_Out := Json_Out || '"quantity_remain":' || Nvl(n_ʣ������, 0);
    Json_Out := Json_Out || ',"item_list":[' || Substr(v_Json, 2) || ']}}}';
  End If;
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Autosplitspeci;
/

Create Or Replace Procedure Zl_Drugsvr_Getprice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡָ��ҩƷ���ۼۡ��ɱ���
  --��Σ�Json_In:��ʽ
  --  input
  --    item_list[]�б�
  --          drug_id          N 1 ҩƷID
  --          pharmacy_id      N 1 ҩ��ID
  --          quantity         N 1 ����
  --����: Json_Out,��ʽ����
  --  output
  --    code              N 1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message           C 1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    data[]�б�
  --          price             N 1 �ۼ�
  --          price_cost        N 1 �ɱ���
  --          quantity_remain   N 1 ʣ������������0��ʾ�����㹻������0���ʾ��������
  ---------------------------------------------------------------------------
  n_ҩƷid   ҩƷ���.ҩƷid%Type;
  n_ҩ��id   ҩƷ���.�ⷿid%Type;
  n_����     ҩƷ���.ʵ������%Type;
  v_Temp     Varchar2(4000);
  n_����     ҩƷ���.�ɱ���%Type;
  n_�ɱ���   ҩƷ���.�ɱ���%Type;
  n_ʣ������ ҩƷ���.ʵ������%Type;
  j_List     Pljson_List := Pljson_List();
  j_Tmpout   Varchar2(32767);
  j_Jsonin   Pljson;
  j_Json     Pljson;

Begin
  --�������
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  j_List   := j_Json.Get_Pljson_List('item_list');
  j_Json   := Pljson();
  If j_List Is Not Null Then
    For I In 1 .. j_List.Count Loop
      j_Json   := Pljson(j_List.Get(I));
      n_ҩƷid := j_Json.Get_Number('drug_id');
      n_ҩ��id := j_Json.Get_Number('pharmacy_id');
      n_����   := j_Json.Get_Number('quantity');
    
      v_Temp := Zl_Fun_Getprice(n_ҩƷid, n_ҩ��id, n_����, 0);
      --�ֽ�
      n_����     := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1));
      v_Temp     := Substr(v_Temp, Instr(v_Temp, '|') + 1);
      n_�ɱ���   := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1));
      v_Temp     := Substr(v_Temp, Instr(v_Temp, '|') + 1);
      n_ʣ������ := To_Number(v_Temp);
      j_Tmpout   := j_Tmpout || ',{"price":' || Zljsonstr(Nvl(n_����, 0), 1);
      j_Tmpout   := j_Tmpout || ',"price_cost": ' || Zljsonstr(Nvl(n_�ɱ���, 0), 1);
      j_Tmpout   := j_Tmpout || ',"quantity_remain":' || Zljsonstr(Nvl(n_ʣ������, 0), 1);
      j_Tmpout   := j_Tmpout || '}';
    
      j_Json := Pljson();
    End Loop;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":[' || Substr(j_Tmpout, 2) || ']}}';

Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getprice;
/

Create Or Replace Procedure Zl_Drugsvr_Getcargospace
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡָ��ҩƷ��ָ���ⷿ�Ļ�λ��Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --   drug_id            N   1   ҩƷID
  --   pharmacy_id        N   1   �ⷿID
  --����: Json_Out,��ʽ����
  --  output
  --    code              N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message           C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    data              C   1   �ⷿ��λ
  ---------------------------------------------------------------------------
  n_ҩƷid ҩƷ���.ҩƷid%Type;
  n_�ⷿid ҩƷ���.�ⷿid%Type;
  v_��λ   ҩƷ�����޶�.�ⷿ��λ%Type;

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --�������
  j_Jsonin := PLJson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_ҩƷid := j_Json.Get_Number('drug_id');
  n_�ⷿid := j_Json.Get_Number('pharmacy_id');

  --��ȡ�ⷿ��λ
  Select Max(�ⷿ��λ) Into v_��λ From ҩƷ�����޶� Where ҩƷid = n_ҩƷid And �ⷿid = n_�ⷿid;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":"' || v_��λ || '"}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getcargospace;
/

Create Or Replace Procedure Zl_Drugsvr_Checkstorelimit
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ָ��ҩƷ��ָ���ⷿ�Ŀ���Ƿ���ڴ�������
  --��Σ�Json_In:��ʽ
  --  input
  --   drug_id            N   1   ҩƷID
  --   pharmacy_id        N   1   �ⷿID
  --   stock              N   1   �������
  --����: Json_Out,��ʽ����
  --  output
  --    code              N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message           C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    below_limit_lower N 1 1-���ڴ������ޣ�0-���ڵ��ڴ�������
  ---------------------------------------------------------------------------
  n_ҩƷid ҩƷ���.ҩƷid%Type;
  n_�ⷿid ҩƷ���.�ⷿid%Type;
  n_���   ҩƷ���.ʵ������%Type;
  n_Count  Number(1);

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --�������
  j_Jsonin := PLJson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_ҩƷid := j_Json.Get_Number('drug_id');
  n_�ⷿid := j_Json.Get_Number('pharmacy_id');
  n_���   := j_Json.Get_Number('stock');

  --��ȡҩƷ�����޶�
  Select Count(1)
  Into n_Count
  From ҩƷ�����޶�
  Where ҩƷid = n_ҩƷid And �ⷿid = n_�ⷿid And Nvl(����, 0) <> 0 And ���� > n_��� And Rownum < 2;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","below_limit_lower":' || n_Count || '}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Checkstorelimit;
/

Create Or Replace Procedure Zl_Drugsvr_Checkisputdrug
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ҩƷ�Ƿ��ѽ�����ҩ
  --��Σ�Json_In:��ʽ
  --  input
  --   rcp_nos            C  1  ҩƷ�շ���¼.no�������Ӣ�Ķ��ŷָ�
  --   billtype           N  1  1-�շѴ�����2-���ʵ�������3-���ʱ���
  --����: Json_Out,��ʽ����
  --  output
  --    code              N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message           C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    isputdrug         N   1   �Ƿ�����ҩ��1-����ҩ,0-δ��ҩ
  ---------------------------------------------------------------------------
  v_Rcp_Nos  Varchar2(4000);
  n_Billtype Number(1);
  n_Count    Number(1);

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --�������
  j_Jsonin   := PLJson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  v_Rcp_Nos  := j_Json.Get_String('rcp_nos');
  n_Billtype := j_Json.Get_Number('billtype');

  Select /*+cardinality(b,10) */
   Count(1)
  Into n_Count
  From δ��ҩƷ��¼ A, Table(f_Str2List(v_Rcp_Nos)) J
  Where a.No = j.Column_Value And a.���� = Decode(n_Billtype, 1, 8, 2, 9, 10) And a.��ҩ�� Is Not Null And Rownum <= 1;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","isputdrug":' || n_Count || '}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Checkisputdrug;
/

Create Or Replace Procedure Zl_Drugsvr_Adjustdata
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ��������תסԺʱ����ҩƷ��������
  --��Σ�Json_In:��ʽ
  --  input
  --    pati_id          N  1 ����ID
  --    pati_pageid      N  1 ��ҳID
  --    billtype         N  1 �������ͣ�1-�շѵ�;2-���ʵ�
  --    item_list
  --      rcp_no_old       C  1 ԭ���ݺ�
  --      rcpdtl_id_old    N  1 ԭ������ϸID(Ŀǰ������Ƿ���id)
  --      rcp_no_new       C  1 �µ��ݺ�
  --      rcpdtl_id_new    N  1 �´�����ϸID(Ŀǰ������Ƿ���id)
  --����: Json_Out,��ʽ����
  --  output
  --    code              N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message           C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Jsonlist Pljson_List;
  n_����id   ҩƷ�շ���¼.����id%Type;
  n_��ҳid   ҩƷ�շ���¼.��ҳid%Type;
  n_����     ҩƷ�շ���¼.����%Type;

  v_ԭ���ݺ� δ��ҩƷ��¼.No%Type;
  n_ԭ��ϸid ҩƷ�շ���¼.����id%Type;
  v_�µ��ݺ� ҩƷ�շ���¼.No%Type;
  n_����ϸid ҩƷ�շ���¼.����id%Type;

  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --�������
  j_Jsonin   := PLJson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_����id   := j_Json.Get_Number('pati_id');
  n_��ҳid   := j_Json.Get_Number('pati_pageid');
  n_����     := j_Json.Get_Number('billtype');
  j_Jsonlist := j_Json.Get_Pljson_List('item_list');

  If j_Jsonlist.Count = 0 Then
    Json_Out := zlJsonOut('δ����ҩƷ������Ϣ');
    Return;
  End If;

  For I In 1 .. j_Jsonlist.Count Loop
    j_Json := PLJson();
  
    j_Json     := PLJson(j_Jsonlist.Get(I));
    v_ԭ���ݺ� := j_Json.Get_String('rcp_no_old');
    n_ԭ��ϸid := j_Json.Get_Number('rcpdtl_id_old');
    v_�µ��ݺ� := j_Json.Get_String('rcp_no_new');
    n_����ϸid := j_Json.Get_Number('rcpdtl_id_new');
  
    Update δ��ҩƷ��¼
    Set ���� = Decode(����, 8, 9, ����), ��ҳid = n_��ҳid, NO = v_�µ��ݺ�
    Where NO = v_ԭ���ݺ� And ���� = Decode(n_����, 1, 8, 2, 9, 10) And ����id = n_����id;
  
    Update ҩƷ�շ���¼
    Set ���� = Decode(����, 8, 9, ����), ����id = n_����ϸid, NO = v_�µ��ݺ�, ��ҳid = n_��ҳid, ������Դ = 2, ������Դ = 2
    Where NO = v_ԭ���ݺ� And ���� = Decode(n_����, 1, 8, 2, 9, 10) And ����id = n_ԭ��ϸid;
  End Loop;

  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugsvr_Adjustdata;
/

Create Or Replace Procedure Zl_Drugsvr_Recipeaffirm
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ�ҩƷ��������ȷ�ϻ�ҩƷ�����շ�ȷ��
  --��Σ�Json_In:��ʽ
  --  input
  --    pati_id                  N  0 ����id������շ�ʱ��Ч
  --    pati_name                C  0 ����������շ�ʱ��Ч
  --    pati_sex                 C  0 �Ա�����շ�ʱ��Ч
  --    pati_age                 C  0 ���䣺����շ�ʱ��Ч
  --    auditor                  C  1 �����
  --    auditor_code             C  1 ����˱��
  --    audit_time               C  1 ���ʱ�䣺yyyy-mm-dd hh24:mi:ss
  --    item_list[]                   ���������б�[����]
  --      billtype               N  1 ��������:1 -�շѴ�����ҩ  ;2- ���ʵ�������ҩ;3- ���ʱ�����ҩ
  --      rcp_no                 C  1 ���ݺ�
  --      rcpdtl_ids             C  0 ����ID,���Դ�����,�ö��ŷ���
  --      pharmacy_window        C  0 ��ҩ����:��ҩ����1:ҩ��ID1| ��|��ҩ����n:ҩ��Idn
  --      drug_auto_send         N  0 �Ƿ��Զ�����ҩƷ:0-���Զ���ҩ,1-�Զ���ҩ
  --      auto_send_ids          C  0 �Զ���ҩ����ϸid����,����ö��ŷָ�
  --����: Json_Out,��ʽ����
  --  output
  --    code              N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message           C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Jsonlist    Pljson_List;
  n_����id      ҩƷ�շ���¼.����id%Type;
  v_����        ҩƷ�շ���¼.����%Type;
  v_�Ա�        ҩƷ�շ���¼.�Ա�%Type;
  v_����        ҩƷ�շ���¼.����%Type;
  v_���ݺ�      ҩƷ�շ���¼.No%Type;
  v_��ϸids     Varchar2(4000);
  v_��ҩ����s   Varchar2(4000);
  n_����        ҩƷ�շ���¼.����%Type;
  n_����_In     ҩƷ�շ���¼.����%Type;
  d_���ʱ��    Date;
  n_�Զ���ҩ    Number(1);
  v_Err         Varchar2(255);
  v_��ҩ��ϸid  Varchar2(400);
  v_��ҩ��ϸids Varchar2(4000);
  v_Nos         Varchar2(32767);
  v_�����      ��Ա��.����%Type;
  v_����˱��  ��Ա��.���%Type;
  Err_Custom Exception;

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --�������
  j_Jsonin     := PLJson(Json_In);
  j_Json       := j_Jsonin.Get_Pljson('input');
  n_����id     := j_Json.Get_Number('pati_id');
  v_����       := j_Json.Get_String('pati_name');
  v_�Ա�       := j_Json.Get_String('pati_sex');
  v_����       := j_Json.Get_String('pati_age');
  v_�����     := j_Json.Get_String('auditor');
  v_����˱�� := j_Json.Get_String('auditor_code');
  d_���ʱ��   := To_Date(j_Json.Get_String('audit_time'), 'yyyy-mm-dd hh24:mi:ss');
  j_Jsonlist   := j_Json.Get_Pljson_List('item_list');

  If d_���ʱ�� Is Null Then
    d_���ʱ�� := Sysdate;
  End If;

  If j_Jsonlist.Count = 0 Then
    v_Err := 'δ����ҩƷ������Ϣ��';
    Raise Err_Custom;
  End If;

  For I In 1 .. j_Jsonlist.Count Loop
    j_Json       := PLJson();
    j_Json       := PLJson(j_Jsonlist.Get(I));
    n_����       := j_Json.Get_Number('billtype');
    v_���ݺ�     := j_Json.Get_String('rcp_no');
    v_��ϸids    := j_Json.Get_String('rcpdtl_ids');
    v_��ҩ����s  := j_Json.Get_String('pharmacy_window');
    n_�Զ���ҩ   := j_Json.Get_Number('drug_auto_send');
    v_��ҩ��ϸid := j_Json.Get_String('auto_send_ids');
  
    n_����_In := n_����;
    If n_���� = 1 Then
      n_���� := 8;
    Elsif n_���� = 2 Then
      n_���� := 9;
    Elsif n_���� = 3 Then
      n_���� := 10;
    Else
      v_Err := '���뵥��������Ч�����飡';
      Raise Err_Custom;
    End If;
  
    If Nvl(v_���ݺ�, '-') = '-' Then
      v_Err := 'δ���봦�����ţ����飡';
      Raise Err_Custom;
    End If;
  
    If Nvl(n_�Զ���ҩ, 0) = 1 Then
      If v_��ҩ��ϸid Is Null Then
        v_Nos := v_Nos || ',' || v_���ݺ�;
      Else
        v_��ҩ��ϸids := v_��ҩ��ϸids || ',' || v_��ҩ��ϸid;
      End If;
    End If;
  
    If Nvl(v_��ϸids, '-') = '-' Then
      v_Err := 'δ���봦����ϸID�����飡';
      Raise Err_Custom;
    End If;
  
    Zl_ҩƷ�շ���¼_�������(n_����, v_���ݺ�, v_��ϸids, d_���ʱ��, v_��ҩ����s, n_����id, v_����, v_�Ա�, v_����);
  End Loop;

  --�������ŷ�ҩ
  If Nvl(v_Nos, '-') <> '-' Then
    v_Nos := Substr(v_Nos, 2);
    Zl_ҩƷ�շ���¼_�Զ���ҩ_s(n_����, v_�����, v_����˱��, v_Nos, 0);
  End If;

  --������ID��ҩ
  If Nvl(v_��ҩ��ϸids, '-') <> '-' Then
    v_��ҩ��ϸids := Substr(v_��ҩ��ϸids, 2);
    Zl_ҩƷ�շ���¼_�Զ���ҩ_s(n_����, v_�����, v_����˱��, v_��ҩ��ϸids, 1);
  End If;

  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugsvr_Recipeaffirm;
/

Create Or Replace Procedure Zl_Drugsvr_Gettakeno
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  --���ܣ���ȡ�Ѿ����ɵ���ҩ��
  --��Σ�Json_In
  --input
  --       dept_id  N 1 ����ID����ID
  --���Σ�Json_Out
  --output
  --       code     N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --       message  C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --       data     C 1 ��ҩ��
  n_����id  δ��ҩƷ��¼.�Է�����id%Type;
  v_Takenos Varchar2(5000);

  j_Jsonin Pljson;
  j_Json   Pljson;
Begin
  --�������
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_����id := j_Json.Get_Number('dept_id');

  For R In (Select Distinct a.��ҩ��
            From δ��ҩƷ��¼ A
            Where a.�������� >= Trunc(Sysdate) And a.���� = 9 And a.�Է�����id = n_����id And a.��ҩ�� Is Not Null
            Order By a.��ҩ�� Desc) Loop
    v_Takenos := v_Takenos || ',' || r.��ҩ��;
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":"' || Substr(v_Takenos, 2) || '"}}';
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Gettakeno;
/

Create Or Replace Procedure Zl_Drugsvr_Newrecipebill
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���Ҫ���ڼ��ʣ������ۣ��� �շ�(������)������µĴ�����ҩ����¼
  --��Σ�Json_In:��ʽ
  --  input
  --     billtype             N   1 ��������: 1 -�շѴ���  ;2- ���ʵ�����;3- ���ʱ���
  --     pati_source          N   1 ������Դ:1-����;2-סԺ;4-���

  ---------------------------billtype = 3,���ʱ�����ҩƷ�շ���¼.����=10��ʱ�������½ڵ�--------------------------------------
  --     pati_id                    N   1 ����ID
  --     pati_pageid                N   1 ��ҳID
  --     pati_name                  C   1 ��������
  --     pati_sex_code              C   1 �Ա��ţ�������)
  --     pati_sex                   C   1 �Ա�
  --     pati_age                   C   1 ����
  --     pati_identity              C     ���
  --     pati_birthdate             C     ��������:yyyy-mm-dd hh:mi:ss
  --     pati_idcard                C     ���֤��
  --     pati_deptid                N   1 ���˿���ID
  --     pati_wardarea_id           N     ���˲���ID
  --     pati_type					        C	  1	�������ͣ���ͨ����,ҽ������,�������,����ְ��...
  ---------------------------billtype = 3,���ʱ�����ҩƷ�շ���¼.����=10��ʱ�������Ͻڵ�-----------------------------------------

  --     si_type_id                 N     ��������id(������):ZLHIS������(���)
  --     si_type_name               C     ������������(������)
  --     rgst_id                    N   1 �Һŵ�id��������)
  --     recipe_proxy_name          C     ������������������)
  --     recipe_proxy_idno          C     ���������֤�ţ�������)
  --     recipe_pat_bodywt          C     �������أ�������)
  --     recipe_pat_bodywt_unit     C     �������ص�λ��������)  
  --     diag_list[]                      �����ٴ�����б�[����]�����²��ŷ�ҩ��  
  --        diag_rec_id             N     ��ϼ�¼id �����²��ŷ�ҩ�� 
  --        diag_type               N     ������� 1-��ҽ�������;2-��ҽ��Ժ���;3-��ҽ��Ժ���;5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���,8-��ǰ���;9-�������;10-����֢;11-��ҽ�������;12-��ҽ��Ժ���;13-��ҽ��Ժ���;21-��ԭѧ���;22-Ӱ��ѧ��ϣ����²��ŷ�ҩ��  
  --        diag_name               C     �ٴ�������ƣ����²��ŷ�ҩ��  
  
  --     pivas_info                 C   0 ������������������Σ����Բ��� ���Ϊһ��json��ʽ����ϸ��ʽͬ��Zl_Pivassvr_Newbill ��������

  --     bill_list[]                      ���������б�[����]
  --        recipe_id                 N  1 ����id(������):ZLHIS�ޣ�����NOת����(��ĸ��Asci����)+������
  --        rcp_no                    C  1 NO
  --        recipe_type               N  0 ��������:0�Ϳ�-��ͨ,1-����,2-����,3-����,4-��һ,5-����
  --        charge_tag                N  1 �շѱ�־:0-δ�շѻ���ʻ���;1-���շѻ����
  --        fee_acnter                C    ������
  --        recipe_plcdept_id         C    ��������id��������)
  --        recipe_plcdept            C    �����������ƣ�������)
  --        recipe_placer_id          C    ����ҽʦid��������)
  --        recipe_placer             C    ����ҽʦ��������) ����
  --        apply_fee_category_code   C    ���뵥�ѱ����(ҽ�Ƹ��ʽ����)(������)  ���ӣ�
  --        apply_fee_category_name   C    ���뵥�ѱ����ƣ�ҽ�Ƹ��ʽ���ƣ�(������)  ���ӣ�
  --        operator_name             C  1 ����Ա����
  --        operator_code             C  1 ����Ա���
  --        create_time               C  1 �Ǽ�ʱ��:yyyy-mm-dd hh:mi:ss
  --        take_no                   C    ��ҩ�� ��ҩ�ţ�δ��ҩƷ��¼.��ҩ�ţ�ҩƷ�շ���¼.��Ʒ�ϸ�֤��ҽ������ʱ����
  --        item_list[]                    ���������б�[����]

  ---------------------------billtype = 3,���ʱ�����ҩƷ�շ���¼.����=10��ʱ�������½ڵ�----------------------------------------
  --           pati_id                 N  1 ����ID
  --           pati_pageid             N    ��ҳID
  --           pati_name               C  1 ��������
  --           pati_sex_code           C  1 �Ա��ţ�������)
  --           pati_sex                C  1 �Ա�
  --           pati_age                C  1 ����
  --           pati_identity           C    ���
  --           pati_birthdate          C    ��������:yyyy-mm-dd hh:mi:ss
  --           pati_idcard             C    ���֤��
  --           pati_wardarea_id        N    ���˲���ID
  --           pati_deptid             N  1 ���˿���ID
  ---------------------------billtype = 3,���ʱ�����ҩƷ�շ���¼.����=10��ʱ�������Ͻڵ�-----------------------------------------

  --           rcpdtl_id               N  1 ������ϸID
  --           serial_num              N  1 ���:(���(�����洢)����ź���ţ�1��2��3��3��3��4��)
  --           group_sno               N    ������� (�����洢)��1��2��3
  --           pharmacy_id             N  1 ҩ��ID
  --           pharmacy_name           C  1 ҩ������(������)
  --           takedept_id             N  1 ��ҩ����ID:���סԺ�Ŵ���
  --           cadn_id                 N  1 ҩƷͨ������id(ҩ��ID)(������)
  --           drug_id                 N  1 ҩƷID
  --           drug_type			         N	1	ҩƷ���ͣ�5-��ҩ��6-��ҩ��7-��ҩ�����²��ŷ�ҩ��
  --           baby_num                N    Ӥ�����

  --           advice_id               N  0 ҽ��ID
  --           drug_method_id          N  1 ��ҩ;��id(������):������ĿID
  --           drug_method_name        C  1 ��ҩ;������
  --           drug_method_class_code  C  1 ��ҩ;������(������):ִ�з�����
  --           drug_freq_id            N  1 ��ҩƵ��id(������):����Ƶ�ʱ���
  --           drug_freq_name          C  1 ��ҩƵ������d(������):

  ---------------------------���½ڵ�Ϊ��ѡ������ҽ����¼����-----------------------------------------------
  --           emergency_tag           N    ҽ����¼�еĽ�����־(0-��ͨ;1-����;2-��¼(��������Ч))
  --           effectivetime           N  0 ҽ����Ч
  --           fee_mode                N  0 �Ƽ����ԣ�0-�����Ƽۣ�1-���Ƽۣ�2-�ֹ��Ƽ�
  --           use_mode                N  0 ȡҩ���ԣ�0-������ʽ��1-��Ժ��ҩ��2-��ȡҩ
  --           frequency               N  0 Ƶ��
  --           single                  N  0 ����
  --           usage                   C  0 �÷�
  --           rcpdtl_st_result        N    Ƥ�Խ��(������)1-���ԣ�2-���ԣ�3-���ԣ�4-������ҩ ��������ʱ��ȷ��������Ƥ�Խ����ZLHISĿǰ֧�ֲ�ȫ
  --           rcpdtL_excs_desc        C    ����˵��(������)
  --           rcpdtL_drask            C    ʹ������(������)
  --           disps_mode_code         C  1 ��ҩ��ʽ(������)1-�������ţ�2-������ҩ��3-�Ա�ҩ��4-����ҩ ZLHISĿǰ֧�ֲ�ȫ(2,4)
  --           drug_content            N    ҩƷ����������ϵ����(������)��
  --           rcpdtl_outp_drugdays    N    ��Ժ����ִ������(������)��ZLHIS�Ǹ�ҩִ�д�����Ҫת��Ϊ������
  --           decoction_method        C  0 �巨
  --           advice_purpose			     C		��ҩĿ�ģ����²��ŷ�ҩ��
  ---------------------------���Ͻڵ�Ϊ��ѡ������ҽ����¼����-----------------------------------------------

  --           packages_num            N  1 ��ҩ����
  --           send_num                N  1 ��ҩ����
  --           send_unit               C  1 ��ҩ��λ��zlhis���۵�λ
  --           price                   N    �ۼ�
  --           money                   N    ���۽��(������)
  --           pharmacy_window         C  0 ��ҩ����
  --           memo                    C  0 ժҪ
  --           fee_source              N  0 ������Դ
  --           drug_auto_send          N  0 �Ƿ��Զ�����ҩƷ:0-���Զ���ҩ,1-�Զ���ҩ
  --           diag_name               C  0 ������ƣ�������)�����ﴫ�룬�������

  --����: Json_Out,��ʽ����
  --  output
  --    code                 N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message              C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------  
Begin
  --ֱ�ӵ���ҩƷҵ����̣���θ�ʽһ�£�
  Zl_ҩƷ�շ���¼_Newdrugbill(Json_In, Json_Out);

  Json_Out := Zljsonout('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugsvr_Newrecipebill;
/

Create Or Replace Procedure Zl_Drugsvr_Surplus
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  --���ܣ�ҩƷ����Ǽǹ��ܶ�ȡҩƷδ������
  --���: Json��ʽ
  --input
  --    dept_id                 N 1 ҩ��id
  --    regbegin_time           C 1 ����������ʼ
  --    regend_time             C 1 �������ڽ���
  --    ward_id                 N 1 ����ID
  --    drugname_show_type      N 1 ҩƷ������ʾ��ʽ
  --    drug_method_ids         C 1 ��ҩ;��
  --    site_no                 C 1 վ��
  --���� Json ��ʽ
  --output
  --    code                    N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message                 C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    data[]�б���Ϣ
  --       drug_id              N 1 ҩƷid
  --       rcpdtl_code          C 1 ����
  --       drug_name            C 1 ҩƷ����
  --       unit                 C 1 ��λ
  --       spec                 C 1 ���
  --       place_name           C 1 ����
  --       quantity             N 1 ��д����
  --       category             C 1 ���
  --       inpack               N 1 סԺ��װ
  n_Deptid       Number(18);
  n_Ward_Id      Number(18);
  n_Show_Type    Number(1);
  v_Method_Ids   Varchar2(4000);
  v_Site_No      Varchar2(50);
  d_Begindate    Date;
  d_Enddate      Date;
  v_Method_Names Varchar2(4000);

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;

  j_Jsonin Pljson;
  j_Json   Pljson;
Begin
  --�������
  j_Jsonin     := Pljson(Json_In);
  j_Json       := j_Jsonin.Get_Pljson('input');
  n_Deptid     := j_Json.Get_Number('dept_id');
  d_Begindate  := To_Date(j_Json.Get_String('regbegin_time'), 'YYYY-MM-DD HH24:MI:SS');
  d_Enddate    := To_Date(j_Json.Get_String('regend_time'), 'YYYY-MM-DD HH24:MI:SS');
  n_Ward_Id    := j_Json.Get_Number('ward_id');
  n_Show_Type  := j_Json.Get_Number('drugname_show_type');
  v_Method_Ids := j_Json.Get_String('drug_method_ids');
  v_Site_No    := j_Json.Get_String('site_no');

  For I In (Select ����
            From ������ĿĿ¼
            Where ��� = 'E' And �������� = '2' And ������� In (2, 3) And (վ�� = v_Site_No Or վ�� Is Null) And
                  ID In (Select Column_Value From Table(f_Str2list(v_Method_Ids)))
            Order By ����) Loop
    v_Method_Names := v_Method_Names || ',' || i.����;
  End Loop;
  If Not v_Method_Names Is Null Then
    v_Method_Names := Substr(v_Method_Names, 2);
  End If;

  v_Jtmp := Null;
  For r_Data In (Select /*+ Rule*/
                  a.ҩƷid, d.����, Nvl(d.����, e.����) As ����, c.סԺ��λ As ��λ, d.���, d.����,
                  Decode(d.���, '7', Sum(a.��д���� / Nvl(c.סԺ��װ, 1) * Nvl(a.����, 1)), Sum(a.��д���� / Nvl(c.סԺ��װ, 1))) As ����,
                  d.���, c.סԺ��װ
                 From ҩƷ�շ���¼ A, ҩƷ��� C, �շ���ĿĿ¼ D, �շ���Ŀ���� E
                 Where a.���� = 9 And a.ҩƷid = c.ҩƷid And c.ҩƷid = d.Id And Mod(a.��¼״̬, 3) = 1 And a.����� Is Null And
                       a.Id = e.�շ�ϸĿid(+) And e.����(+) = 1 And e.����(+) = n_Show_Type And
                       Decode(v_Method_Ids, '', '', Nvl(a.�÷�, 'Null')) Not In
                       (Select Column_Value From Table(f_Str2list(v_Method_Names))) And a.�������� Between d_Begindate And
                       d_Enddate And a.�ⷿid = n_Deptid And a.���˲���id = n_Ward_Id
                 Group By a.ҩƷid, d.����, Nvl(d.����, e.����), c.סԺ��λ, d.���, d.����, d.���, c.סԺ��װ
                 Order By ����) Loop
  
    v_Jtmp := v_Jtmp || ',';
    Zljsonputvalue(v_Jtmp, 'drug_id', r_Data.ҩƷid, 1, 1);
    Zljsonputvalue(v_Jtmp, 'rcpdtl_code', r_Data.����, 0);
    Zljsonputvalue(v_Jtmp, 'drug_name', r_Data.����, 0);
    Zljsonputvalue(v_Jtmp, 'unit', r_Data.��λ, 0);
    Zljsonputvalue(v_Jtmp, 'spec', r_Data.���, 0);
    Zljsonputvalue(v_Jtmp, 'place_name', r_Data.����, 0);
    Zljsonputvalue(v_Jtmp, 'quantity', r_Data.����, 1);
    Zljsonputvalue(v_Jtmp, 'category', r_Data.���, 0);
    Zljsonputvalue(v_Jtmp, 'inpack', r_Data.סԺ��װ, 1, 2);
  
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
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Surplus;
/

Create Or Replace Procedure Zl_Drugsvr_Getrecipeaudit
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ��鿴���������
  --��Σ�Json_In:��ʽ
  --  input
  --   pati_id              N   1   ����ID
  --   pvid                 N   1   ���߾���id     ����Һ�ID   סԺ����ҳID
  --   pat_source           N   1   ������Դ       ������Դ:1-����;2-סԺ
  --����: Json_Out,��ʽ����
  --  output
  --    code                 N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message              C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    item_list
  --     advice_id            N   1   ҽ��ID
  --     recipe_audit_status  N   1   �������״̬      0-����1-����2-��ʱ����(����ҩʦһֱδ������)�� 11-���󱻳�����
  --     recipe_audit_result  N   1   ���������      1-�ϸ�2-���ϸ�
  ---------------------------------------------------------------------------
  n_����id Number(18);
  n_����id Number(18);
  n_��Դ   Number(2);
  v_Jtmp   Varchar2(32767);

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --�������
  j_Jsonin := PLJson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');
  n_����id := j_Json.Get_Number('pvid');
  n_��Դ   := j_Json.Get_Number('pat_source');

  If Nvl(n_����id, 0) <> 0 Then
    If Nvl(n_��Դ, 0) = 1 Then
      For c_��� In (Select i.ҽ��id, j.״̬ As �������״̬, j.����� As ���������
                   From ��������¼ J, ���������ϸ I
                   Where j.����id = n_����id And j.�Һ�id = n_����id And j.Id = i.��id(+) And (i.����ύ = 1 Or i.��id Is Null)) Loop
      
        v_Jtmp := v_Jtmp || ',{"advice_id":' || c_���.ҽ��id;
        v_Jtmp := v_Jtmp || ',"recipe_audit_status":' || Nvl(c_���.�������״̬, 0);
        v_Jtmp := v_Jtmp || ',"recipe_audit_result":' || Nvl(c_���.���������, 0);
        v_Jtmp := v_Jtmp || '}';
      End Loop;
    Else
      For c_��� In (Select i.ҽ��id, j.״̬ As �������״̬, j.����� As ���������
                   From ��������¼ J, ���������ϸ I
                   Where j.����id = n_����id And j.��ҳid = n_����id And j.Id = i.��id(+) And (i.����ύ = 1 Or i.��id Is Null)) Loop
      
        v_Jtmp := v_Jtmp || ',{"advice_id":' || c_���.ҽ��id;
        v_Jtmp := v_Jtmp || ',"recipe_audit_status":' || Nvl(c_���.�������״̬, 0);
        v_Jtmp := v_Jtmp || ',"recipe_audit_result":' || Nvl(c_���.���������, 0);
        v_Jtmp := v_Jtmp || '}';
      
      End Loop;
    End If;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getrecipeaudit;
/

Create Or Replace Procedure Zl_Drugsvr_Retainedrecords
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  --���ܣ���ȡ�Ѿ���д��ҩƷ�����¼
  --��Σ�JOSN��ʽ
  --input
  --  ward_id                  N 1 ����ID
  --  dept_id                  N 1 �ⷿID
  --  drugname_show_type       N 1 ҩƷ������ʾ��ʽ 
  --���Σ�JSON
  --output
  --  code                     N Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message                  C Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  data[]�б�
  --     drug_id               N 1 ҩƷid
  --     rcpdtl_code           C 1 ����
  --     drug_name             C 1 ҩƷ����
  --     unit                  C 1 ��λ
  --     spec                  C 1 ���
  --     place_name            C 1 ����
  --     quantity              N 1 ��������
  --     category              C 1 ���
  --     inpack                N 1 סԺ��װ 
  n_Dept_Id Number(18);
  n_Ward_Id Number(18);
  n_Type    Number(3);

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;

  j_Jsonin Pljson;
  j_Json   Pljson;
Begin
  --�������
  j_Jsonin  := Pljson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  n_Ward_Id := j_Json.Get_Number('ward_id');
  n_Dept_Id := j_Json.Get_Number('dept_id');
  n_Type    := j_Json.Get_Number('drugname_show_type');

  If Nvl(n_Ward_Id, 0) <> 0 Then
    For r_Data In (Select a.ҩƷid, c.����, Nvl(d.����, c.����) As ����, c.���, c.����, b.סԺ��λ As ��λ, a.�������� / Nvl(b.סԺ��װ, 1) As ��������,
                          c.���, b.סԺ��װ
                   From ҩƷ����ƻ� A, ҩƷ��� B, �շ���ĿĿ¼ C, �շ���Ŀ���� D
                   Where a.ҩƷid = b.ҩƷid And a.ҩƷid = c.Id And a.����id = n_Ward_Id And a.�ⷿid = n_Dept_Id And a.״̬ = 0 And
                         c.Id = d.�շ�ϸĿid(+) And d.����(+) = 1 And d.����(+) = n_Type
                   Order By c.����) Loop
    
      v_Jtmp := v_Jtmp || ',';
      Zljsonputvalue(v_Jtmp, 'drug_id', r_Data.ҩƷid, 1, 1);
      Zljsonputvalue(v_Jtmp, 'rcpdtl_code', r_Data.����, 0);
      Zljsonputvalue(v_Jtmp, 'drug_name', r_Data.����, 0);
      Zljsonputvalue(v_Jtmp, 'unit', r_Data.��λ, 0);
      Zljsonputvalue(v_Jtmp, 'spec', r_Data.���, 0);
      Zljsonputvalue(v_Jtmp, 'place_name', r_Data.����, 0);
      Zljsonputvalue(v_Jtmp, 'quantity', r_Data.��������, 1);
      Zljsonputvalue(v_Jtmp, 'category', r_Data.���, 0);
      Zljsonputvalue(v_Jtmp, 'inpack', r_Data.סԺ��װ, 1, 2);
    
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
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Retainedrecords;
/

Create Or Replace Procedure Zl_Drugsvr_Getretainbyorder
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  --���ܣ���ȡ�Ѿ���д��ҩƷ�����¼����ҽ��id���в�ѯ
  --��Σ�JOSN��ʽ
  --input
  --  order_ids                C 1 ҽ��IDs������ҽ��id��ȡҩƷ��ϸ������ƴ��
  --���Σ�JSON
  --output
  --  code                     N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message                  C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ 
  --  data[]�б�
  --     order_id              N 1 ҽ��id
  --     drug_id               N 1 ҩƷid
  --     drug_content          N 1 ����ϵ��
  --     si_drug_packg_qunt    N 1 סԺ��װ
  --     is_part               N 1 סԺ�ɷ����

  v_Order_Ids Varchar2(32767);
  v_Jtmp      Varchar2(32767);
  c_Jtmp      Clob;
  j_Jsonin    Pljson;
  j_Json      Pljson;
Begin
  --�������
  j_Jsonin    := Pljson(Json_In);
  j_Json      := j_Jsonin.Get_Pljson('input');
  v_Order_Ids := j_Json.Get_String('order_ids');
  For r_Data In (Select a.ҽ��id, a.ҩƷid, b.����ϵ��, b.סԺ��װ, b.סԺ�ɷ����
                 From ҩƷ�շ���¼ A, ҩƷ��� B
                 Where a.ҩƷid = b.ҩƷid And
                       a.ҽ��id In (Select /*+cardinality(x,10)*/
                                   x.Column_Value
                                  From Table(Cast(f_Num2list(v_Order_Ids) As Zltools.t_Numlist)) X)
                 Order By a.ҽ��id) Loop
  
    v_Jtmp := v_Jtmp || ',';
    Zljsonputvalue(v_Jtmp, 'order_id', r_Data.ҽ��id, 1, 1);
    Zljsonputvalue(v_Jtmp, 'drug_id', r_Data.ҩƷid, 1);
    Zljsonputvalue(v_Jtmp, 'drug_content', r_Data.����ϵ��, 1);
    Zljsonputvalue(v_Jtmp, 'si_drug_packg_qunt', r_Data.סԺ��װ, 1);
    Zljsonputvalue(v_Jtmp, 'is_part', r_Data.סԺ�ɷ����, 1, 2);
  
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
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":[' || c_Jtmp || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getretainbyorder;
/

CREATE OR REPLACE Procedure Zl_Drugsvr_Retainplan_Insert
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  --���ܣ� ����ҩƷ����ƻ�
  --���:JSON��ʽ
  --input
  --       ward_id              N 1 ����id
  --       dept_id              N 1 �ⷿid
  --       operator_name        C 1 ����Ա����
  --       drug_list[]ҩƷ�����б�
  --           drug_id          N 1 ҩƷid
  --           quantity         N 1 ����
  --����:JSON��ʽ
  --output
  --       code                 N 1 Ӧ���룺0-ʧ�� 1-�ɹ�
  --       message              C 1 �ɹ���ʧ�ܺ󷵻ص���Ϣ
  j_Jsonlist      Pljson_List;
  d_����ʱ��      Date;
  n_Dept_Id       Number(18); --�ⷿid
  n_Warda_Id      Number(18); --����id
  v_Operator_Name Varchar2(50);
  j_Jsonin        Pljson;
  j_Json          Pljson;
Begin
  --�������
  j_Jsonin        := Pljson(Json_In);
  j_Json          := j_Jsonin.Get_Pljson('input');
  n_Dept_Id       := j_Json.Get_Number('dept_id');
  n_Warda_Id      := j_Json.Get_Number('ward_id');
  v_Operator_Name := j_Json.Get_String('operator_name');
  j_Jsonlist      := j_Json.Get_Pljson_List('drug_list');
  Select Sysdate Into d_����ʱ�� From Dual;
  j_Json := Pljson();
  For I In 1 .. j_Jsonlist.Count Loop
    j_Json := Pljson(j_Jsonlist.Get(I));
    Insert Into ҩƷ����ƻ�
      (����id, �ⷿid, ҩƷid, ��������, ״̬, �Ǽ���, �Ǽ�ʱ��)
    Values
      (n_Warda_Id, n_Dept_Id, j_Json.Get_Number('drug_id'), j_Json.Get_Number('quantity'), 0, v_Operator_Name, d_����ʱ��);
    j_Json := Pljson();
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugsvr_Retainplan_Insert;
/

Create Or Replace Procedure Zl_Drugsvr_Retainplan_Delete
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  --���ܣ� ����ҩƷ����ƻ�
  --���:JSON��ʽ
  --input
  --     ward_id            N 1 ����id
  --     dept_id            N 1 �ⷿid
  --     drug_id            N 0 ҩƷid
  --����:JSON��ʽ
  --output
  --     code               N 1 Ӧ���룺0-ʧ�� 1-�ɹ�
  --     message            C 1 �ɹ���ʧ�ܺ󷵻ص���Ϣ
  n_Dept_Id  Number;
  n_Warda_Id Number;
  n_Drug_Id  Number;

  j_Jsonin Pljson;
  j_Json   Pljson;
Begin
  --�������
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_Dept_Id  := j_Json.Get_Number('dept_id');
  n_Warda_Id := j_Json.Get_Number('ward_id');
  n_Drug_Id  := j_Json.Get_Number('drug_id');

  If n_Drug_Id Is Null Then
    Delete ҩƷ����ƻ� Where ����id = n_Warda_Id And �ⷿid = n_Dept_Id And ״̬ = 0 And ����id Is Null;
  Else
    Delete ҩƷ����ƻ�
    Where ����id = n_Warda_Id And �ⷿid = n_Dept_Id And ҩƷid = n_Drug_Id And ״̬ = 0 And ����id Is Null;
  End If;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugsvr_Retainplan_Delete;
/

Create Or Replace Procedure Zl_Drugsvr_Merage
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  --���ܣ�������Ϣ�ϲ�ʱ��ҩƷ�շ���¼�ϲ�
  --��Σ�JSON��ʽ
  --input
  ----retain_id ����id
  ----merge_id �ϲ�id
  ----pati_name C ����
  ----pati_sex C �Ա�
  ----pati_age C ����
  ----pati_borth_time C ��������
  ----pati_identity C ���֤��;
  ----item_list
  ------page_id_new N ����ҳid 
  ------pati_id_befor N  ԭ����id 
  ------page_id_befor N  ԭ��ҳid 
  --���Σ�JSON��ʽ
  --output
  ----code N Ӧ���룺0-ʧ�ܣ�1-�ɹ�  
  ----message C Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ 
  j_Json_List  Pljson_List;
  n_Retain_Id  Number;
  n_Merge_Id   Number;
  v_Name       Varchar2(100);
  v_Sex        Varchar2(10);
  v_Age        Varchar2(20);
  d_Borth_Time Date;
  v_Identity   Varchar2(20);
  n_Page_New   Number;
  n_Pati_Befor Number;
  n_Page_Befor Number;
  I            Number;

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --�������
  j_Jsonin     := PLJson(Json_In);
  j_Json       := j_Jsonin.Get_Pljson('input');
  n_Retain_Id  := j_Json.Get_Number('retain_id');
  n_Merge_Id   := j_Json.Get_Number('merge_id');
  v_Name       := j_Json.Get_String('pati_name');
  v_Sex        := j_Json.Get_String('pati_sex');
  v_Age        := j_Json.Get_String('pati_age');
  d_Borth_Time := To_Date(j_Json.Get_String('pati_borth_time'), 'YYYY-MM-DD HH24:MI:SS');
  v_Identity   := j_Json.Get_String('pati_identity');
  j_Json_List  := j_Json.Get_Pljson_List('item_list');

  For I In 1 .. j_Json_List.Count Loop
    j_Json       := PLJson();
    j_Json       := PLJson(j_Json_List.Get(I));
    n_Page_New   := j_Json.Get_Number('page_id_new');
    n_Pati_Befor := j_Json.Get_Number('pati_id_befor');
    n_Page_Befor := j_Json.Get_Number('page_id_befor');
  
    Update δ��ҩƷ��¼
    Set ����id = n_Retain_Id, ��ҳid = n_Page_New, ���� = v_Name
    Where ����id = n_Pati_Befor And ��ҳid = n_Page_Befor;
  
    Update ҩƷ�շ���¼
    Set ����id = n_Retain_Id, ��ҳid = n_Page_New, ���� = v_Name, �Ա� = v_Sex, ���� = v_Age, �������� = d_Borth_Time,
        ���֤�� = v_Identity
    Where ����id = n_Pati_Befor And ��ҳid = n_Page_Befor;
  End Loop;

  Update δ��ҩƷ��¼ Set ����id = n_Retain_Id, ���� = v_Name Where ����id = n_Merge_Id And ��ҳid Is Null;

  Update ҩƷ�շ���¼
  Set ����id = n_Retain_Id, ���� = v_Name, �Ա� = v_Sex, ���� = v_Age, �������� = d_Borth_Time, ���֤�� = v_Identity
  Where ����id = n_Merge_Id And ��ҳid Is Null;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugsvr_Merage;
/

Create Or Replace Procedure Zl_Drugsvr_Delrecipebill
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ�ҩƷ��������(���˷�)
  --��Σ�Json_In:��ʽ
  -- input
  --     billtype                 N   1   ��������:1 -�շѴ�����ҩ  ;2- ���ʵ�������ҩ
  --     rcp_no                   C   1   ���ݺ�,�иýڵ��ǰ�����NO�������ʣ���ʱֻ������������ϵͳ�ӿڣ�
  --     item_list[]���������б�[����]�����б�
  --          rcpdtl_id           N 1 ������ϸid,Ŀǰ����ķ���ID
  --          chargeoffs_num      N 1 ��������
  --          dispensing_ids      N 1 ��ҩIDs :�������ʱ��Һ����������Ҫ���ݵļ�¼id�����ַ������ݣ��ö��ŷָ�磺1001,1002,1003
  --     pivas_list[]���ھ���������Ҫ�Զ���˵��������ɾ��ҩƷ��ͬʱ���������������,��ͨ����ɾ�����ô�����б�
  --          pivas_ids           C 1 ��ҩIDs����Ҫ���ʲ�ͬʱ��˵���Һids��
  --          operator_name       C 1 ����Ա��
  --          operator_time       C 1 ����ʱ��
  --          reason              C 1 ����ԭ��
  --      return_list[]�Զ���ҩ�б�
  --           audit_operator     C 1 �����
  --           operator_time      C 1 ����ʱ��
  --           rcpdtl_list[]��ҩ�б���Ϣ
  --               rcpdtl_id      N 1 ����id
  --               re_quantity    N 0 ��ҩ���������۵����շ���ĿĿ¼.���㵥λ�����ɲ����˽����ߴ��գ���ʾȫ�ˣ�����ָ��������ҩ

  --����: Json_Out,��ʽ����
  -- output
  --    code                      N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                   C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  n_������ϸid ҩƷ�շ���¼.����id%Type;
  n_��������   ҩƷ�շ���¼.ʵ������%Type;
  v_����Ա���� Varchar2(4000);
  d_����ʱ��   Date;
  v_����˵��   Varchar2(4000);
  v_��Һids    Varchar2(32767);
  n_С��       Number(6);
  n_����id     Number(18);
  n_����       ҩƷ�շ���¼.ʵ������%Type;
  n_��ҩ����   ҩƷ�շ���¼.ʵ������%Type;
  n_����״̬   Number(3);
  n_����       Number(1);
  v_No         ҩƷ�շ���¼.No%Type;

  v_Err Varchar2(255);
  Err_Custom Exception;

  j_Jsonin   Pljson;
  j_Json     Pljson;
  j_Jsonlist Pljson_List;
  Jl_Re      Pljson_List;
  j_Re       Pljson;
  j_Item     Pljson;
Begin
  --�������
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');

  --��NO���ˣ�Ŀǰ����������ӿ�
  v_No := j_Json.Get_String('rcp_no');
  If v_No Is Not Null Then
    n_���� := j_Json.Get_Number('billtype');
  
    For r_������ϸ In (Select ����id, Sum(Nvl(����, 1) * ʵ������) As ��ҩ����
                   From ҩƷ�շ���¼
                   Where ���� = Decode(n_����, 1, 8, 2, 9, 10) And NO = v_No And ������� Is Null
                   Group By ����id
                   Order By ����id) Loop
    
      --������id��������
      Zl_ҩƷ�շ���¼_�����˷�_s(r_������ϸ.����id, r_������ϸ.��ҩ����, Null, 1);
    End Loop;
  End If;

  --�Զ���ҩ 
  j_Jsonlist := j_Json.Get_Pljson_List('return_list');
  If j_Jsonlist Is Not Null Then
    Select ���� Into n_С�� From ҩƷ���ľ��� Where ��� = 1 And ���� = 4 And ��λ = 5;
    For I In 1 .. j_Jsonlist.Count Loop
      j_Item       := Pljson(); --����м���� 
      j_Item       := Pljson(j_Jsonlist.Get(I));
      v_����Ա���� := j_Item.Get_String('audit_operator');
      d_����ʱ��   := To_Date(j_Item.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');
      Jl_Re        := j_Item.Get_Pljson_List('rcpdtl_list');
      For K In 1 .. Jl_Re.Count Loop
        j_Re     := Pljson(Jl_Re.Get(K));
        n_����id := j_Re.Get_Number('rcpdtl_id');
        n_����   := j_Re.Get_Number('re_quantity');
        j_Re     := Pljson();
        If n_���� Is Not Null Then
          n_��ҩ���� := n_����;
        End If;
        --�ֽ���ҩ����
        For r_������ϸ In (Select a.�ⷿid, a.Id, Nvl(a.����, 1) * a.ʵ������ As ����
                       From ҩƷ�շ���¼ A
                       Where a.���� In (8, 9, 10) And (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0) And ����� Is Not Null And
                             a.����id = n_����id
                       Order By a.�ⷿid, a.ҩƷid, a.����) Loop
          If n_���� Is Null Then
            --���������Ϊ�ձ�ʾȫ��
            --������ҩ��������ϸ��
            Zl_ҩƷ�շ���¼_������ҩ_s(r_������ϸ.Id, v_����Ա����, d_����ʱ��, Null, Null, Null, Null, Null, Null, n_С��);
          Else
            If n_��ҩ���� > 0 Then
              If n_��ҩ���� > r_������ϸ.���� Then
                --������ҩ��������ϸ��
                Zl_ҩƷ�շ���¼_������ҩ_s(r_������ϸ.Id, v_����Ա����, d_����ʱ��, Null, Null, Null, r_������ϸ.����, Null, Null, n_С��);
                n_��ҩ���� := n_��ҩ���� - r_������ϸ.����;
              Else
                --������ҩ��������ϸ��
                Zl_ҩƷ�շ���¼_������ҩ_s(r_������ϸ.Id, v_����Ա����, d_����ʱ��, Null, Null, Null, n_��ҩ����, Null, Null, n_С��);
                n_��ҩ���� := 0;
              End If;
            End If;
          End If;
        End Loop;
      End Loop;
      Jl_Re := Pljson_List();
    End Loop;
  End If;

  --ɾ������
  j_Jsonlist := Pljson_List();
  j_Jsonlist := j_Json.Get_Pljson_List('item_list');
  If j_Jsonlist Is Not Null Then
    For I In 1 .. j_Jsonlist.Count Loop
      j_Item       := Pljson();
      j_Item       := Pljson(j_Jsonlist.Get(I));
      n_������ϸid := j_Item.Get_Number('rcpdtl_id');
      n_��������   := j_Item.Get_Number('chargeoffs_num');
      v_��Һids    := j_Item.Get_String('dispensing_ids');
      If n_������ϸid Is Null Then
        v_Err := '����ڵ㡾rcpdtl_id���������飡';
        Raise Err_Custom;
      End If;
      If n_�������� Is Null Then
        v_Err := '����ڵ㡾chargeoffs_num���������飡';
        Raise Err_Custom;
      End If;
      Zl_ҩƷ�շ���¼_�����˷�_s(n_������ϸid, n_��������, v_��Һids, 1);
    End Loop;
  End If;

  --������ش���
  j_Jsonlist := Pljson_List();
  j_Jsonlist := j_Json.Get_Pljson_List('pivas_list');
  If j_Jsonlist Is Not Null Then
    For I In 1 .. j_Jsonlist.Count Loop
      j_Item       := Pljson();
      j_Item       := Pljson(j_Jsonlist.Get(I));
      v_����Ա���� := j_Item.Get_String('operator_name');
      d_����ʱ��   := To_Date(j_Item.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');
      v_����˵��   := j_Item.Get_String('reason');
      v_��Һids    := j_Item.Get_String('pivas_ids');
      n_����״̬   := j_Item.Get_Number('operator_status');
      Zl_��Һ��ҩ��¼_���ʸ���״̬_s(v_����Ա����, d_����ʱ��, v_����˵��, v_��Һids, n_����״̬);
    End Loop;
  End If;
  Json_Out := Zljsonout('�ɹ�', 1);
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugsvr_Delrecipebill;
/

Create Or Replace Procedure Zl_Drugsvr_Sendquery
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  -----------------------------------------------------------------------------------------------------------------------
  --���ܣ�ҩ���շ���ѯ
  --��Σ�JSON��ʽ
  --  input
  --      type_query              N 0 ��ѯ����:0-��ҩ��ϸ�嵥 1-��ҩ�����嵥 2-��ҩ��ϸ�嵥 3-��ҩ�����嵥[�ɷ��������ֱ�ӻ�ȡ] 4-�˷������嵥 
  --      dept_id                 N 1 �ⷿid
  --      otherdept_id            N 0 �Է�����id,��ҩ����id
  --      verify_begin_time       C 0 �����ʼ����
  --      verify_end_time         C 1 �����ʼ���� 
  --      rcp_no                  C 1 NO��ҩƷ�շ���¼.NO
  --      group_no                N 1 ���ܷ�ҩ��
  --      record_state            C 1 ��¼״̬,��λ������ƴ����01-��ѯ�ѷ�ҩ��10-ֻ��ѯδ��ҩ������-������ͬʱ��ѯ�ѷ�ҩ��δ��ҩ
  --      effective_time          N 1 ��Ч 0-������1-������2-����������
  --      ward_id                 N 1 ����ID
  --      usages                  C 1 �÷�����ҩ;�� ����ƴ�������ŷָ�
  --      drugname_show_type      N 0 ҩƷ������ʾ��ʽ���ɲ���ȱʡΪ1;��Ӧ�ڡ��շ���Ŀ����.���ʡ���1-����(��Ӧ��Ŀ�е�����);2-Ӣ����;3-��Ʒ��;9-����������������ҩ����Ӣ��������Ʒ�������Ĵ�����Ʒ�����������ֻ�������ͱ���
  --      pati_ids                C 0 ����ID����ƴ������������ֲ��ˣ�����
  --      rcpdtl_ids              C 0 ��ҩ��ѯʱ����  ���ܷ�ҩ�� ��ȡ������ϸid����ƴ��
  --      rcp_nos                 C 0 ���ݺţ�����ƴ���������水����ʱ���ѯʱ����
  --      return_list[]��ҩ��Ϣ�б�˵����type_query=2��ѯʱ���루quantity+rcpdtl_id+serial_num����type_query=(3��4)��ѯʱ���루drug_id+quantity+re_money��
  --         drug_id              N 1 ҩƷid
  --         quantity             N 1 ��ҩ����
  --         re_money             N 1 ���
  --         rcpdtl_id            N 0 ����id,������ϸ 
  --         serial_num           N 0 ��ţ�Ψһ��ʶһ������

  --����
  --output
  --  code                        N 1 Ӧ���룺0-ʧ�� 1-�ɹ�
  --  message                     C 1 ʧ�ܺ���Ҫ���صĴ�����Ϣ
  --  data[]  �б���Ϣ
  --      record_state            N 1 ��¼״̬
  --      rcpdtl_id               N 1 ����id
  --      rcp_no                  C 1 NO��ҩƷ�շ���¼.no������ҽ������.no��
  --      order_number            N 1 ���
  --      drugstore_name          C 1 ҩ����
  --      advice_dept_name        C 1 ��������
  --      rcp_info                C 1 ҩƷ��Ϣ
  --      in_unit                 C 1 סԺ��λ
  --      dosage_unit             C 1 ������λ
  --      quantity                N 1 ����
  --      uint_price              N 1 ����
  --      effective               N 1 Ч��
  --      singular_quantity       N 1 ����
  --      money                   N 1 ���
  --      category                N 1 ���
  --      frequency               N 1 Ƶ��
  --      usage                   C 1 �÷�
  --      payment                 N 1 ����
  --      advice_id               N 1 ҽ��ID
  --      pati_name               C 1 ����
  --      pati_id                 N 1 ����id
  --      page_id                 N 1 ��ҳid
  --      drug_code               C 1 ҩƷ����   n_type=1��n_type=3��n_type=4ʱ�������ֵ 
  --      back_number             N 1 ��ҩ�� n_type=4ʱ�������ֵ
  --      reality_number          N 1 ʵ���� n_type=4ʱ�������ֵ
  -----------------------------------------------------------------------------------------------------------------------

  n_Type           Number(6);
  n_�ⷿid         Number(18);
  n_�Է�����id     Number(18);
  n_����id         Number(18);
  d_�����ʼ����   Date;
  d_��˽�������   Date;
  v_No             Varchar2(30);
  n_���ܷ�ҩ��     Number(18);
  v_��¼״̬       Varchar2(30);
  n_Ч��           Number(3);
  v_�÷�s          Varchar2(3980);
  v_����ids        Varchar2(3980);
  n_Showtype       Number(3);
  d_��ҩ����ʱ���� Date;
  d_��ҩ����ʱ��ֹ Date;
  l_��ҩ��Ϣ       t_Strlist2 := t_Strlist2(); --�����ָ�ʽ��һ������ϸ������id+��ҩ����+��� ��һ���ǻ��ܣ�ҩƷid+��ҩ����+���
  n_����id         Number(18);
  v_Nos            Varchar2(32767);
  c_No             t_Strlist := t_Strlist();

  Cursor c_List_Type Is
    Select a.Id ״̬, a.No, a.���, a.ժҪ ҩ��, a.ժҪ ��������, a.����, a.ժҪ ҩƷ��Ϣ, a.���� ����, a.ժҪ סԺ��λ, a.���� ����, a.���� ���, a.ժҪ ��Ч, a.����,
           a.ժҪ ������λ, a.Ƶ��, a.�÷�, a.����, a.ժҪ ���, a.����id, a.��ҳid, a.����id ��ҩ���
    From ҩƷ�շ���¼ A
    Where 0 = 1;
  r_Detail c_List_Type%RowType;

  Cursor c_Group_Type Is
    Select a.ժҪ ҩƷ����, a.ժҪ ҩƷ��Ϣ, a.ժҪ סԺ��λ, a.���� ����, a.���� ��ҩ��, a.���� ʵ����, a.���� ���
    From ҩƷ�շ���¼ A
    Where 0 = 1;
  r_Grp c_Group_Type%RowType;

  v_Jtmp     Varchar2(32767); --��Ҫ���ʹ�ô˱���
  c_Jtmp     Clob; --��Ҫ���ʹ�ô˱���
  j_Jsonlist Pljson_List;
  j_Tmp      Pljson;
  j_Jsonin   Pljson;
  j_Json     Pljson;

  Procedure Get����ƴ����ϸ As
  Begin
    v_Jtmp := v_Jtmp || ',';
    Zljsonputvalue(v_Jtmp, 'record_state', r_Detail.״̬, 1, 1);
    Zljsonputvalue(v_Jtmp, 'rcp_no', r_Detail.No, 0);
    Zljsonputvalue(v_Jtmp, 'order_number', r_Detail.���, 1);
    Zljsonputvalue(v_Jtmp, 'drugstore_name', r_Detail.ҩ��, 0);
    Zljsonputvalue(v_Jtmp, 'advice_dept_name', r_Detail.��������, 0);
    Zljsonputvalue(v_Jtmp, 'rcp_info', r_Detail.ҩƷ��Ϣ, 0);
    Zljsonputvalue(v_Jtmp, 'in_unit', r_Detail.סԺ��λ, 0);
    Zljsonputvalue(v_Jtmp, 'dosage_unit', r_Detail.������λ, 0);
    Zljsonputvalue(v_Jtmp, 'uint_price', r_Detail.����, 1);
    Zljsonputvalue(v_Jtmp, 'money', r_Detail.���, 1);
    Zljsonputvalue(v_Jtmp, 'effective', r_Detail.��Ч, 0);
    Zljsonputvalue(v_Jtmp, 'singular_quantity', r_Detail.����, 1);
    Zljsonputvalue(v_Jtmp, 'quantity', r_Detail.����, 1);
    Zljsonputvalue(v_Jtmp, 'category', r_Detail.���, 0);
    Zljsonputvalue(v_Jtmp, 'frequency', r_Detail.Ƶ��, 0);
    Zljsonputvalue(v_Jtmp, 'usage', r_Detail.�÷�, 0);
    Zljsonputvalue(v_Jtmp, 'payment', r_Detail.����, 1);
    Zljsonputvalue(v_Jtmp, 'pati_name', r_Detail.����, 0);
    Zljsonputvalue(v_Jtmp, 'pati_id', r_Detail.����id, 1);
    Zljsonputvalue(v_Jtmp, 'rcpdtl_id', r_Detail.��ҩ���, 1);
    Zljsonputvalue(v_Jtmp, 'page_id', r_Detail.��ҳid, 1, 2);
  
    If Length(v_Jtmp) > 30000 Then
      If c_Jtmp Is Null Then
        c_Jtmp := Substr(v_Jtmp, 2);
      Else
        c_Jtmp := c_Jtmp || v_Jtmp;
      End If;
      v_Jtmp := Null;
    End If;
  End;

  Procedure Get����ƴ������ As
  Begin
    v_Jtmp := v_Jtmp || ',';
    Zljsonputvalue(v_Jtmp, 'rcp_info', r_Grp.ҩƷ��Ϣ, 0, 1);
    Zljsonputvalue(v_Jtmp, 'drug_code', r_Grp.ҩƷ����, 0);
    Zljsonputvalue(v_Jtmp, 'in_unit', r_Grp.סԺ��λ, 0);
    Zljsonputvalue(v_Jtmp, 'quantity', r_Grp.����, 1);
    Zljsonputvalue(v_Jtmp, 'money', r_Grp.���, 1);
    Zljsonputvalue(v_Jtmp, 'back_number', r_Grp.��ҩ��, 1);
    Zljsonputvalue(v_Jtmp, 'reality_number', r_Grp.ʵ����, 1, 2);
  
    If Length(v_Jtmp) > 30000 Then
      If c_Jtmp Is Null Then
        c_Jtmp := Substr(v_Jtmp, 2);
      Else
        c_Jtmp := c_Jtmp || v_Jtmp;
      End If;
      v_Jtmp := Null;
    End If;
  End;

Begin
  --�������
  j_Jsonin     := Pljson(Json_In);
  j_Json       := j_Jsonin.Get_Pljson('input');
  n_Type       := Nvl(j_Json.Get_Number('type_query'), 0);
  n_�ⷿid     := j_Json.Get_Number('dept_id');
  n_�Է�����id := j_Json.Get_Number('otherdept_id');
  n_����id     := j_Json.Get_Number('ward_id');
  Select zl_GetSysParameter('ҩƷ������ʾ', Null, 100) Into n_Showtype From Dual;
  If Nvl(n_Showtype, 0) = 0 Then
    n_Showtype := 1;
  End If;
  d_�����ʼ���� := To_Date(j_Json.Get_String('verify_begin_time'), 'YYYY-MM-DD HH24:MI:SS');
  d_��˽������� := To_Date(j_Json.Get_String('verify_end_time'), 'YYYY-MM-DD HH24:MI:SS');
  v_No           := j_Json.Get_String('rcp_no');
  n_���ܷ�ҩ��   := j_Json.Get_String('group_no');
  v_��¼״̬     := j_Json.Get_String('record_state');
  If Not (v_��¼״̬ = '01' Or v_��¼״̬ = '10') Then
    v_��¼״̬ := Null;
  End If;
  n_Ч��    := j_Json.Get_Number('effective_time');
  v_�÷�s   := j_Json.Get_String('usages');
  v_����ids := j_Json.Get_String('pati_ids');
  v_Nos     := j_Json.Get_String('rcp_nos');

  If v_Nos Is Not Null Then
    v_Nos := v_Nos || ',';
    While v_Nos Is Not Null Loop
      c_No.Extend;
      c_No(c_No.Count) := Substr(v_Nos, 1, Instr(v_Nos, ',') - 1);
      v_Nos := Substr(v_Nos, Instr(v_Nos, ',') + 1);
    End Loop;
  End If;

  j_Jsonlist := j_Json.Get_Pljson_List('return_list');
  If j_Jsonlist Is Not Null Then
    v_Jtmp := Null;
    For K In 1 .. j_Jsonlist.Count Loop
      j_Tmp    := Pljson(j_Jsonlist.Get(K));
      n_����id := j_Tmp.Get_Number('rcpdtl_id');
      l_��ҩ��Ϣ.Extend();
      If Nvl(n_����id, 0) = 0 Then
        l_��ҩ��Ϣ(l_��ҩ��Ϣ.Count) := t_Strobj2(j_Tmp.Get_Number('drug_id'),
                                          j_Tmp.Get_Number('quantity') || '_' || j_Tmp.Get_Number('re_money'));
      Else
        l_��ҩ��Ϣ(l_��ҩ��Ϣ.Count) := t_Strobj2(n_����id, j_Tmp.Get_Number('quantity') || '_' || j_Tmp.Get_Number('serial_num'));
      End If;
      j_Tmp := Pljson();
    End Loop;
  End If;

  -----------------------------------------------
  --�Է�ҩʱ��Ϊ׼�Ĳ�ѯ,�� ������� Ϊ������
  If n_Type = 0 Then
    --��ҩ��ϸ�嵥 tbcQuery.Selected.Index = 0
    If d_�����ʼ���� Is Not Null Then
      For R In (Select a.״̬, a.No, a.���, i.���� As ҩ��, h.���� As ��������, a.����,
                       Nvl(x.����, f.����) || Decode(f.����, Null, Null, '(' || f.���� || ')') ||
                        Decode(f.���, Null, Null, ' ' || f.���) As ҩƷ��Ϣ, a.���� / Nvl(e.סԺ��װ, 1) As ����, e.סԺ��λ,
                       a.���� * Nvl(e.סԺ��װ, 1) As ����, a.���, a.��Ч, a.����, g.���㵥λ As ������λ, a.Ƶ��, a.�÷�, a.����, g.���, a.����id,
                       a.��ҳid, 0 ��ҩid
                From (Select a.״̬, a.No, a.���, a.�ⷿid, a.�Է�����id, a.����, a.סԺ��, a.����, a.ҩƷid, a.����, a.��Ч, a.����, a.Ƶ��, a.�÷�,
                              a.ʱ��, a.��Ա, a.����, a.����id, a.��ҳid, Sum(a.����) As ����, Sum(a.���) As ���
                       From (Select b.״̬, c.No, c.���, c.�ⷿid, c.�Է�����id, c.����, '�貹��' סԺ��, '�貹��' ����, c.ҩƷid, b.����, c.���ۼ� As ����,
                                     b.���, Decode(Nvl(Substr(c.����, 1, 1), 0), 0, '����', '����') As ��Ч, c.����, c.Ƶ��, c.�÷�,
                                     'A.����ʱ��' As ʱ��, 'A.������' As ��Ա, c.����, c.����id, c.��ҳid
                              From (Select Decode(a.�����, Null, 0, 1) As ״̬, a.No, a.���, Sum(a.��д���� * a.����) As ����,
                                            Sum(a.���۽��) As ���
                                     From ҩƷ�շ���¼ A
                                     Where a.������� Between d_�����ʼ���� And d_��˽������� And a.���� + 0 = 9 And a.ҽ��id Is Not Null And
                                           (a.�Է�����id = n_�Է�����id Or Nvl(n_�Է�����id, 0) = 0) And
                                           (a.�ⷿid = n_�ⷿid Or Nvl(n_�ⷿid, 0) = 0) And
                                           (a.No = v_No Or Nvl(v_No, 'NONE') = 'NONE') And
                                           (a.���ܷ�ҩ�� = n_���ܷ�ҩ�� Or Nvl(n_���ܷ�ҩ��, 0) = 0) And
                                           (Nvl(Substr(a.����, 1, 1), 0) = n_Ч�� Or Nvl(n_Ч��, 2) = 2) And
                                           (v_��¼״̬ = '01' And a.����� Is Not Null Or
                                           v_��¼״̬ = '10' And Mod(a.��¼״̬, 3) = 1 And a.����� Is Null Or
                                           Nvl(v_��¼״̬, 'NONE') = 'NONE') And
                                           (a.�÷� In (Select /*+cardinality(x,10)*/
                                                      x.Column_Value
                                                     From Table(f_Str2list(v_�÷�s)) X) Or Nvl(v_�÷�s, 'NONE') = 'NONE') And
                                           (a.����id + 0 In (Select /*+cardinality(x,10)*/
                                                            x.Column_Value
                                                           From Table(f_Str2list(v_����ids)) X) Or
                                           Nvl(v_����ids, 'NONE') = 'NONE')
                                     Group By Decode(a.�����, Null, 0, 1), a.No, a.���
                                     Having Nvl(Sum(a.��д����), 0) <> 0 Or Nvl(Sum(a.���۽��), 0) <> 0) B, ҩƷ�շ���¼ C
                              Where b.No = c.No And b.��� = c.��� And (c.��¼״̬ = 1 Or Mod(c.��¼״̬, 3) = 0) And
                                    (c.���˲���id = n_����id Or Nvl(n_����id, 0) = 0)) A
                       Group By a.״̬, a.No, a.���, a.�ⷿid, a.�Է�����id, a.����, a.סԺ��, a.����, a.ҩƷid, a.����, a.��Ч, a.����, a.Ƶ��, a.�÷�,
                                a.ʱ��, a.��Ա, a.����, a.����id, a.��ҳid) A, ҩƷ��� E, �շ���ĿĿ¼ F, ������ĿĿ¼ G, ���ű� H, ���ű� I, �շ���Ŀ���� X
                Where a.ҩƷid = e.ҩƷid And a.ҩƷid = f.Id And e.ҩ��id = g.Id And a.�Է�����id = h.Id And a.�ⷿid = i.Id And
                      f.Id = x.�շ�ϸĿid(+) And x.����(+) = 1 And x.����(+) = n_Showtype
                Order By a.No, a.���) Loop
        --�����ⲿ����Ҫ�� ���� ���������ⲿ׷��α�б�
        r_Detail := R;
        Get����ƴ����ϸ;
      End Loop;
    Else
      For R In (Select a.״̬, a.No, a.���, i.���� As ҩ��, h.���� As ��������, a.����,
                       Nvl(x.����, f.����) || Decode(f.����, Null, Null, '(' || f.���� || ')') ||
                        Decode(f.���, Null, Null, ' ' || f.���) As ҩƷ��Ϣ, a.���� / Nvl(e.סԺ��װ, 1) As ����, e.סԺ��λ,
                       a.���� * Nvl(e.סԺ��װ, 1) As ����, a.���, a.��Ч, a.����, g.���㵥λ As ������λ, a.Ƶ��, a.�÷�, a.����, g.���, a.����id,
                       a.��ҳid, 0 ��ҩid
                From (Select a.״̬, a.No, a.���, a.�ⷿid, a.�Է�����id, a.����, a.סԺ��, a.����, a.ҩƷid, a.����, a.��Ч, a.����, a.Ƶ��, a.�÷�,
                              a.ʱ��, a.��Ա, a.����, a.����id, a.��ҳid, Sum(a.����) As ����, Sum(a.���) As ���
                       From (Select b.״̬, c.No, c.���, c.�ⷿid, c.�Է�����id, c.����, '�貹��' סԺ��, '�貹��' ����, c.ҩƷid, b.����, c.���ۼ� As ����,
                                     b.���, Decode(Nvl(Substr(c.����, 1, 1), 0), 0, '����', '����') As ��Ч, c.����, c.Ƶ��, c.�÷�,
                                     'A.����ʱ��' As ʱ��, 'A.������' As ��Ա, c.����, c.����id, c.��ҳid
                              From (Select Decode(a.�����, Null, 0, 1) As ״̬, a.No, a.���, Sum(a.��д���� * a.����) As ����,
                                            Sum(a.���۽��) As ���
                                     From ҩƷ�շ���¼ A
                                     Where a.No In (Select /*+cardinality(x,10)*/
                                                     Column_Value
                                                    From Table(c_No) X) And a.���� + 0 = 9 And a.ҽ��id Is Not Null And
                                           (a.�Է�����id = n_�Է�����id Or Nvl(n_�Է�����id, 0) = 0) And
                                           (a.�ⷿid = n_�ⷿid Or Nvl(n_�ⷿid, 0) = 0) And
                                           (a.No = v_No Or Nvl(v_No, 'NONE') = 'NONE') And
                                           (a.���ܷ�ҩ�� = n_���ܷ�ҩ�� Or Nvl(n_���ܷ�ҩ��, 0) = 0) And
                                           (Nvl(Substr(a.����, 1, 1), 0) = n_Ч�� Or Nvl(n_Ч��, 2) = 2) And
                                           (v_��¼״̬ = '01' And a.����� Is Not Null Or
                                           v_��¼״̬ = '10' And Mod(a.��¼״̬, 3) = 1 And a.����� Is Null Or
                                           Nvl(v_��¼״̬, 'NONE') = 'NONE') And
                                           (a.�÷� In (Select /*+cardinality(x,10)*/
                                                      x.Column_Value
                                                     From Table(f_Str2list(v_�÷�s)) X) Or Nvl(v_�÷�s, 'NONE') = 'NONE') And
                                           (a.����id + 0 In (Select /*+cardinality(x,10)*/
                                                            x.Column_Value
                                                           From Table(f_Str2list(v_����ids)) X) Or
                                           Nvl(v_����ids, 'NONE') = 'NONE')
                                     Group By Decode(a.�����, Null, 0, 1), a.No, a.���
                                     Having Nvl(Sum(a.��д����), 0) <> 0 Or Nvl(Sum(a.���۽��), 0) <> 0) B, ҩƷ�շ���¼ C
                              Where b.No = c.No And b.��� = c.��� And (c.��¼״̬ = 1 Or Mod(c.��¼״̬, 3) = 0) And
                                    (c.���˲���id = n_����id Or Nvl(n_����id, 0) = 0)) A
                       Group By a.״̬, a.No, a.���, a.�ⷿid, a.�Է�����id, a.����, a.סԺ��, a.����, a.ҩƷid, a.����, a.��Ч, a.����, a.Ƶ��, a.�÷�,
                                a.ʱ��, a.��Ա, a.����, a.����id, a.��ҳid) A, ҩƷ��� E, �շ���ĿĿ¼ F, ������ĿĿ¼ G, ���ű� H, ���ű� I, �շ���Ŀ���� X
                Where a.ҩƷid = e.ҩƷid And a.ҩƷid = f.Id And e.ҩ��id = g.Id And a.�Է�����id = h.Id And a.�ⷿid = i.Id And
                      f.Id = x.�շ�ϸĿid(+) And x.����(+) = 1 And x.����(+) = n_Showtype
                Order By a.No, a.���) Loop
        --�����ⲿ����Ҫ�� ���� ���������ⲿ׷��α�б�
        r_Detail := R;
        Get����ƴ����ϸ;
      End Loop;
    End If;
  
  Elsif n_Type = 1 Then
    --��ҩ�����嵥 tbcQuery.Selected.Index = 1    
    If d_�����ʼ���� Is Not Null Then
      For R In (Select a.ҩƷ����,
                       Nvl(b.����, a.����) || Decode(a.����, Null, Null, '(' || a.���� || ')') ||
                        Decode(a.���, Null, Null, ' ' || a.���) As ҩƷ��Ϣ, a.סԺ��λ, a.����, 0 ��ҩ��, 0 ʵ����, a.���
                From (Select b.ҩƷid, c.���� As ҩƷ����, c.����, c.����, c.���, b.סԺ��λ, Sum(a.���� / Nvl(b.סԺ��װ, 1)) As ����,
                              Sum(a.���) As ���
                       From (Select b.״̬, c.No, c.���, c.�ⷿid, c.�Է�����id, c.����, '�貹��' סԺ��, '�貹��' ����, c.ҩƷid, b.����, c.���ۼ� As ����,
                                     b.���, Decode(Nvl(Substr(c.����, 1, 1), 0), 0, '����', '����') As ��Ч, c.����, c.Ƶ��, c.�÷�,
                                     '����ʱ���貹��' As ʱ��, '�������貹��' As ��Ա, c.����, c.����id, c.��ҳid
                              From (Select Decode(a.�����, Null, 0, 1) As ״̬, a.No, a.���, Sum(a.��д���� * a.����) As ����,
                                            Sum(a.���۽��) As ���
                                     From ҩƷ�շ���¼ A
                                     Where a.������� Between d_�����ʼ���� And d_��˽������� And a.���� + 0 = 9 And a.ҽ��id Is Not Null And
                                           (a.�Է�����id = n_�Է�����id Or Nvl(n_�Է�����id, 0) = 0) And
                                           (a.�ⷿid = n_�ⷿid Or Nvl(n_�ⷿid, 0) = 0) And
                                           (a.No = v_No Or Nvl(v_No, 'NONE') = 'NONE') And
                                           (a.���ܷ�ҩ�� = n_���ܷ�ҩ�� Or Nvl(n_���ܷ�ҩ��, 0) = 0) And
                                           (Nvl(Substr(a.����, 1, 1), 0) = n_Ч�� Or Nvl(n_Ч��, 2) = 2) And
                                           (v_��¼״̬ = '01' And a.����� Is Not Null Or
                                           v_��¼״̬ = '10' And Mod(a.��¼״̬, 3) = 1 And a.����� Is Null Or
                                           Nvl(v_��¼״̬, 'NONE') = 'NONE') And
                                           (a.�÷� In (Select /*+cardinality(x,10)*/
                                                      x.Column_Value
                                                     From Table(f_Str2list(v_�÷�s)) X) Or Nvl(v_�÷�s, 'NONE') = 'NONE') And
                                           (a.����id + 0 In (Select /*+cardinality(x,10)*/
                                                            x.Column_Value
                                                           From Table(f_Str2list(v_����ids)) X) Or
                                           Nvl(v_����ids, 'NONE') = 'NONE')
                                     Group By Decode(a.�����, Null, 0, 1), a.No, a.���
                                     Having Nvl(Sum(a.��д����), 0) <> 0 Or Nvl(Sum(a.���۽��), 0) <> 0) B, ҩƷ�շ���¼ C
                              Where b.No = c.No And b.��� = c.��� And (c.��¼״̬ = 1 Or Mod(c.��¼״̬, 3) = 0) And
                                    (c.���˲���id = n_����id Or Nvl(n_����id, 0) = 0)) A, ҩƷ��� B, �շ���ĿĿ¼ C
                       Where a.ҩƷid = b.ҩƷid And a.ҩƷid = c.Id
                       Group By b.ҩƷid, c.����, c.����, c.����, c.���, b.סԺ��λ
                       Having Sum(a.���� / Nvl(b.סԺ��װ, 1)) <> 0 Or Sum(a.���) <> 0) A, �շ���Ŀ���� B
                Where a.ҩƷid = b.�շ�ϸĿid(+) And b.����(+) = 1 And b.����(+) = n_Showtype
                Order By a.ҩƷ����) Loop
        r_Grp := R;
        Get����ƴ������;
      End Loop;
    Else
      For R In (Select a.ҩƷ����,
                       Nvl(b.����, a.����) || Decode(a.����, Null, Null, '(' || a.���� || ')') ||
                        Decode(a.���, Null, Null, ' ' || a.���) As ҩƷ��Ϣ, a.סԺ��λ, a.����, 0 ��ҩ��, 0 ʵ����, a.���
                From (Select b.ҩƷid, c.���� As ҩƷ����, c.����, c.����, c.���, b.סԺ��λ, Sum(a.���� / Nvl(b.סԺ��װ, 1)) As ����,
                              Sum(a.���) As ���
                       From (Select b.״̬, c.No, c.���, c.�ⷿid, c.�Է�����id, c.����, '�貹��' סԺ��, '�貹��' ����, c.ҩƷid, b.����, c.���ۼ� As ����,
                                     b.���, Decode(Nvl(Substr(c.����, 1, 1), 0), 0, '����', '����') As ��Ч, c.����, c.Ƶ��, c.�÷�,
                                     '����ʱ���貹��' As ʱ��, '�������貹��' As ��Ա, c.����, c.����id, c.��ҳid
                              From (Select Decode(a.�����, Null, 0, 1) As ״̬, a.No, a.���, Sum(a.��д���� * a.����) As ����,
                                            Sum(a.���۽��) As ���
                                     From ҩƷ�շ���¼ A
                                     Where a.No In (Select /*+cardinality(x,10)*/
                                                     Column_Value
                                                    From Table(c_No) X) And a.���� + 0 = 9 And a.ҽ��id Is Not Null And
                                           (a.�Է�����id = n_�Է�����id Or Nvl(n_�Է�����id, 0) = 0) And
                                           (a.�ⷿid = n_�ⷿid Or Nvl(n_�ⷿid, 0) = 0) And
                                           (a.No = v_No Or Nvl(v_No, 'NONE') = 'NONE') And
                                           (a.���ܷ�ҩ�� = n_���ܷ�ҩ�� Or Nvl(n_���ܷ�ҩ��, 0) = 0) And
                                           (Nvl(Substr(a.����, 1, 1), 0) = n_Ч�� Or Nvl(n_Ч��, 2) = 2) And
                                           (v_��¼״̬ = '01' And a.����� Is Not Null Or
                                           v_��¼״̬ = '10' And Mod(a.��¼״̬, 3) = 1 And a.����� Is Null Or
                                           Nvl(v_��¼״̬, 'NONE') = 'NONE') And
                                           (a.�÷� In (Select /*+cardinality(x,10)*/
                                                      x.Column_Value
                                                     From Table(f_Str2list(v_�÷�s)) X) Or Nvl(v_�÷�s, 'NONE') = 'NONE') And
                                           (a.����id + 0 In (Select /*+cardinality(x,10)*/
                                                            x.Column_Value
                                                           From Table(f_Str2list(v_����ids)) X) Or
                                           Nvl(v_����ids, 'NONE') = 'NONE')
                                     Group By Decode(a.�����, Null, 0, 1), a.No, a.���
                                     Having Nvl(Sum(a.��д����), 0) <> 0 Or Nvl(Sum(a.���۽��), 0) <> 0) B, ҩƷ�շ���¼ C
                              Where b.No = c.No And b.��� = c.��� And (c.��¼״̬ = 1 Or Mod(c.��¼״̬, 3) = 0) And
                                    (c.���˲���id = n_����id Or Nvl(n_����id, 0) = 0)) A, ҩƷ��� B, �շ���ĿĿ¼ C
                       Where a.ҩƷid = b.ҩƷid And a.ҩƷid = c.Id
                       Group By b.ҩƷid, c.����, c.����, c.����, c.���, b.סԺ��λ
                       Having Sum(a.���� / Nvl(b.סԺ��װ, 1)) <> 0 Or Sum(a.���) <> 0) A, �շ���Ŀ���� B
                Where a.ҩƷid = b.�շ�ϸĿid(+) And b.����(+) = 1 And b.����(+) = n_Showtype
                Order By a.ҩƷ����) Loop
        r_Grp := R;
        Get����ƴ������;
      End Loop;
    End If;
  Elsif n_Type = 2 Then
    --tbcQuery.Selected.Index = 2 ��ҩ��ϸ 
    For R In (Select a.״̬, a.No, a.���, i.���� As ҩ��, h.���� As ��������, a.����,
                     Nvl(x.����, f.����) || Decode(f.����, Null, Null, '(' || f.���� || ')') ||
                      Decode(f.���, Null, Null, ' ' || f.���) As ҩƷ��Ϣ, a.���� / Nvl(e.סԺ��װ, 1) As ����, e.סԺ��λ,
                     a.���� * Nvl(e.סԺ��װ, 1) As ����, a.���, a.��Ч, a.����, g.���㵥λ As ������λ, a.Ƶ��, a.�÷�, a.����, g.���, a.����id,
                     a.��ҳid, a.�˷�id
              From (Select Distinct -1 As ״̬, d.No, d.���, d.�ⷿid, d.�Է�����id, d.����, '�貹��' סԺ��, '�貹��' ����, d.ҩƷid, a.����,
                                     d.���ۼ� As ����, a.���� * d.���ۼ� ���, Decode(Nvl(Substr(d.����, 1, 1), 0), 0, '����', '����') ��Ч,
                                     d.����, d.Ƶ��, d.�÷�, a.��ҩ������� As ʱ��, a.��ҩ������� As ��Ա, d.����, d.����id, d.��ҳid, a.��ҩ������� �˷�id
                     From (Select /*+cardinality(x,10)*/
                             To_Number(x.C1) ����id, To_Number(Substr(x.C2, 1, Instr(x.C2, '_') - 1)) ����,
                             To_Number(Substr(x.C2, Instr(x.C2, '_') + 1)) ��ҩ�������
                            From Table(l_��ҩ��Ϣ) X) A, ҩƷ�շ���¼ D
                     Where a.����id = d.����id) A, ҩƷ��� E, �շ���ĿĿ¼ F, ������ĿĿ¼ G, ���ű� H, ���ű� I, �շ���Ŀ���� X
              Where a.ҩƷid = e.ҩƷid And a.ҩƷid = f.Id And e.ҩ��id = g.Id And a.�Է�����id = h.Id And a.�ⷿid = i.Id And
                    f.Id = x.�շ�ϸĿid(+) And x.����(+) = 1 And x.����(+) = n_Showtype
              Order By a.No, a.���) Loop
      r_Detail := R;
      Get����ƴ����ϸ;
    End Loop;
  Elsif n_Type = 3 Then
    ---tbcQuery.Selected.Index = 3 ��ҩ���� 
    For R In (Select a.ҩƷ����,
                     Nvl(b.����, a.����) || Decode(a.����, Null, Null, '(' || a.���� || ')') ||
                      Decode(a.���, Null, Null, ' ' || a.���) As ҩƷ��Ϣ, a.סԺ��λ, a.����, 0 ��ҩ��, 0 ʵ����, a.���
              From (Select b.ҩƷid, c.���� As ҩƷ����, c.����, c.����, c.���, b.סԺ��λ, Sum(a.���� / Nvl(b.סԺ��װ, 1)) As ����,
                            Sum(a.���) As ���
                     From (Select a.����id, a.����, a.���� * b.��׼���� ���, b.�շ�ϸĿid ҩƷid
                            From ���˷������� A, סԺ���ü�¼ B
                            Where a.����id = b.Id And a.����ʱ�� Between d_��ҩ����ʱ���� And d_��ҩ����ʱ��ֹ And b.ҽ����� Is Not Null And
                                  a.��˲���id = n_�ⷿid And Nvl(a.״̬, 0) = 0 And b.�շ���� In ('5', '6', '7') And
                                  (b.��ҩ����id = n_�Է�����id Or Nvl(n_�Է�����id, 0) = 0) And a.���벿��id = n_����id And
                                  (b.����id + 0 In (Select /*+cardinality(x,10)*/
                                                   x.Column_Value
                                                  From Table(f_Str2list(v_����ids)) X) Or Nvl(v_����ids, 'NONE') = 'NONE') And
                                  (b.ҽ����Ч = n_Ч�� Or Nvl(n_Ч��, 2) = 2)) A, ҩƷ��� B, �շ���ĿĿ¼ C
                     Where a.ҩƷid = b.ҩƷid And a.ҩƷid = c.Id
                     Group By b.ҩƷid, c.����, c.����, c.����, c.���, b.סԺ��λ
                     Having Sum(a.���� / Nvl(b.סԺ��װ, 1)) <> 0 Or Sum(a.���) <> 0) A, �շ���Ŀ���� B
              Where a.ҩƷid = b.�շ�ϸĿid(+) And b.����(+) = 1 And b.����(+) = n_Showtype
              Order By a.ҩƷ����
              -- ���ܷ�ҩ�ţ�ת���ɷ���IDƴ���ٴ��������Ǳ�ȥ��
              -- ��ҩ��ѯ��ʱ���� ��ҩ;����������ʧЧ���ڲ�����������ֻ�����������
              ) Loop
      r_Grp := R;
      Get����ƴ������;
    End Loop;
  Elsif n_Type = 4 Then
    --tbcQuery.Selected.Index = 4 ����ҩ����ҳ��
    For R In (Select a.ҩƷ����,
                     Nvl(b.����, a.����) || Decode(a.����, Null, Null, '(' || a.���� || ')') ||
                      Decode(a.���, Null, Null, ' ' || a.���) As ҩƷ��Ϣ, a.סԺ��λ, a.Ӧ����, a.��ҩ��, a.ʵ����, a.���
              From (Select b.ҩƷid, c.���� As ҩƷ����, c.����, c.����, c.���, b.סԺ��λ, Sum(a.Ӧ���� / Nvl(b.סԺ��װ, 1)) As Ӧ����,
                            Sum(a.��ҩ�� / Nvl(b.סԺ��װ, 1)) As ��ҩ��, (Sum(a.Ӧ����) - Sum(a.��ҩ��)) / Nvl(b.סԺ��װ, 1) As ʵ����,
                            Sum(a.���) As ���
                     From (Select a.ҩƷid, a.���� As Ӧ����, 0 As ��ҩ��, a.���
                            From (Select a.״̬, a.No, a.���, a.�ⷿid, a.�Է�����id, a.����, a.סԺ��, a.����, a.ҩƷid, a.����, a.��Ч, a.����, a.Ƶ��,
                                          a.�÷�, a.ʱ��, a.��Ա, a.����, a.����id, a.��ҳid, Sum(a.����) As ����, Sum(a.���) As ���
                                   From (Select b.״̬, c.No, c.���, c.�ⷿid, c.�Է�����id, c.����, '�貹��' סԺ��, '�貹��' ����, c.ҩƷid, b.����,
                                                 c.���ۼ� As ����, b.���, Decode(Nvl(Substr(c.����, 1, 1), 0), 0, '����', '����') As ��Ч, c.����,
                                                 c.Ƶ��, c.�÷�, 'A.����ʱ��' As ʱ��, 'A.������' As ��Ա, c.����, c.����id, c.��ҳid
                                          From (Select Decode(a.�����, Null, 0, 1) As ״̬, a.No, a.���, Sum(a.��д���� * a.����) As ����,
                                                        Sum(a.���۽��) As ���
                                                 From ҩƷ�շ���¼ A
                                                 Where a.������� Between d_�����ʼ���� And d_��˽������� And a.���� + 0 = 9 And
                                                       a.ҽ��id Is Not Null And (a.�Է�����id = n_�Է�����id Or Nvl(n_�Է�����id, 0) = 0) And
                                                       (a.�ⷿid = n_�ⷿid Or Nvl(n_�ⷿid, 0) = 0) And
                                                       (a.No = v_No Or Nvl(v_No, 'NONE') = 'NONE') And
                                                       (a.���ܷ�ҩ�� = n_���ܷ�ҩ�� Or Nvl(n_���ܷ�ҩ��, 0) = 0) And
                                                       (Nvl(Substr(a.����, 1, 1), 0) = n_Ч�� Or Nvl(n_Ч��, 2) = 2) And
                                                       (v_��¼״̬ = '01' And a.����� Is Not Null Or
                                                       v_��¼״̬ = '10' And Mod(a.��¼״̬, 3) = 1 And a.����� Is Null Or
                                                       Nvl(v_��¼״̬, 'NONE') = 'NONE') And
                                                       (a.�÷� In (Select /*+cardinality(x,10)*/
                                                                  x.Column_Value
                                                                 From Table(f_Str2list(v_�÷�s)) X) Or Nvl(v_�÷�s, 'NONE') = 'NONE') And
                                                       (a.����id + 0 In (Select /*+cardinality(x,10)*/
                                                                        x.Column_Value
                                                                       From Table(f_Str2list(v_����ids)) X) Or
                                                       Nvl(v_����ids, 'NONE') = 'NONE')
                                                 Group By Decode(a.�����, Null, 0, 1), a.No, a.���
                                                 Having Nvl(Sum(a.��д����), 0) <> 0 Or Nvl(Sum(a.���۽��), 0) <> 0) B, ҩƷ�շ���¼ C
                                          Where b.No = c.No And b.��� = c.��� And (c.��¼״̬ = 1 Or Mod(c.��¼״̬, 3) = 0) And
                                                (c.���˲���id = n_����id Or Nvl(n_����id, 0) = 0)) A
                                   Group By a.״̬, a.No, a.���, a.�ⷿid, a.�Է�����id, a.����, a.סԺ��, a.����, a.ҩƷid, a.����, a.��Ч, a.����,
                                            a.Ƶ��, a.�÷�, a.ʱ��, a.��Ա, a.����, a.����id, a.��ҳid) A
                            Union All
                            --��ҩ��ϸ l_��ҩ��Ϣ ��ҩƷ:����_��  123123:13_33                             
                            Select /*+cardinality(x,10)*/
                             To_Number(x.C1) ҩƷid, 0 Ӧ����, To_Number(Substr(x.C2, 1, Instr(x.C2, '_') - 1)) ��ҩ��,
                             -1 * To_Number(Substr(x.C2, Instr(x.C2, '_') + 1)) ���
                            From Table(l_��ҩ��Ϣ) X) A, ҩƷ��� B, �շ���ĿĿ¼ C
                     Where a.ҩƷid = b.ҩƷid And a.ҩƷid = c.Id
                     Group By b.ҩƷid, c.����, c.����, c.����, c.���, b.סԺ��λ, Nvl(b.סԺ��װ, 1)
                     Having Sum(a.Ӧ���� / Nvl(b.סԺ��װ, 1)) <> 0 Or Sum(a.��ҩ�� / Nvl(b.סԺ��װ, 1)) <> 0 Or Sum(a.���) <> 0) A,
                   �շ���Ŀ���� B
              Where a.ҩƷid = b.�շ�ϸĿid(+) And b.����(+) = 1 And b.����(+) = n_Showtype
              Order By a.ҩƷ����) Loop
      --�����ⲿ����Ҫ�� ���� ���������ⲿ׷��α�б�    
      r_Grp := R;
      Get����ƴ������;
    End Loop;
    --�����嵥���
    --ҩƷ���룬ҩƷ��Ϣ��סԺ��λ����������ҩ����ʵ���������
    --��ϸ�嵥���
    --״̬��NO����ţ�ҩ�����������ң�������סԺ�ţ����ţ�סԺ��λ�����������ۣ�����Ч��������������λ��Ƶ�Σ��÷���ʱ�䣬��Ա������������˷�id      
  End If;

  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":[' || c_Jtmp || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Sendquery;
/


Create Or Replace Procedure Zl_Drugsvr_Getadditional_Infor
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡҩƷ��һЩ��չ�򸽼ӵ���Ϣ���������÷���������Ƶ�Σ����͵�
  --��Σ�Json_In:��ʽ
  --    input
  --        billtype                    N   1   ��������:1 -�շѴ�����ҩ  ;2- ���ʵ�������ҩ
  --        rcp_no                  C   1   ���ݺ�
  --        rcpdtl_ids                  C       ������ϸids,Ŀǰ����ķ���ID
  --    ���� json
  --    output
  --        code                    N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --        message                 C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --        data[]                         ���������б�[����]
  --            rcp_no              C   1   NO
  --            rcpdtl_id               N   1   ������ϸid,Ŀǰ����ķ���ID
  --            frequency               C   1   Ƶ��
  --            usage               C   1   �÷�
  --            si_drug_form                C   1   ����
  --            loitem_detail_measunit              C   1   ������λ
  --            advice_exe_properties               N   1   ִ������:0~2-�Ƽ�����,3-��Ժ��ҩ,4-��ȡҩ
  ---------------------------------------------------------------------------
  v_No       ҩƷ�շ���¼.No%Type;
  n_�������� Number(2);
  v_�������� Varchar2(10);
  n_ִ������ Number(2);
  v_����ids  Varchar2(32767);

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --�������
  j_Jsonin   := PLJson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_�������� := j_Json.Get_Number('billtype');
  v_No       := j_Json.Get_String('rcp_no');
  v_����ids  := ',' || Nvl(j_Json.Get_String('rcpdtl_ids'), '') || ',';

  If v_No Is Null Then
    Json_Out := zlJsonOut('δ��ȷ����');
    Return;
  End If;

  If Nvl(n_��������, 0) = 2 Then
    v_�������� := ',9,10,';
  Else
    v_�������� := ',8,';
  End If;

  For c_ҩƷ In (Select a.����id, a.No, a.����, a.Ƶ��, a.�÷�, d.���� As ����, c.������λ
               From (Select a.����id, a.ҩƷid, a.No, Max(a.Ƶ��) As Ƶ��, Max(a.�÷�) As �÷�, Max(����) As ����
                      From ҩƷ�շ���¼ A
                      Where Instr(v_��������, ',' || ���� || ',') > 0 And NO = v_No
                      Group By a.No, a.����id, a.ҩƷid) A, ҩƷĿ¼ B, ҩƷ��Ϣ C, ҩƷ���� D
               Where Instr(v_����ids, ',' || a.����id || ',') > 0 And a.ҩƷid = b.ҩƷid And b.ҩ��id = c.ҩ��id And c.���� = d.����) Loop
  
    n_ִ������ := To_Number(Substr(Nvl(c_ҩƷ.����, '00'), 2, 1));
  
    v_Jtmp := v_Jtmp || ',';
    zlJsonPutValue(v_Jtmp, 'rcp_no', c_ҩƷ.No, 0, 1);
    zlJsonPutValue(v_Jtmp, 'rcpdtl_id', c_ҩƷ.����id, 1);
    zlJsonPutValue(v_Jtmp, 'frequency', c_ҩƷ.Ƶ��);
    zlJsonPutValue(v_Jtmp, 'usage', c_ҩƷ.�÷�);
    zlJsonPutValue(v_Jtmp, 'si_drug_form', c_ҩƷ.����);
    zlJsonPutValue(v_Jtmp, 'loitem_detail_measunit', c_ҩƷ.������λ);
    zlJsonPutValue(v_Jtmp, 'advice_exe_properties', n_ִ������, 1, 2);
  
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
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getadditional_Infor;
/


Create Or Replace Procedure Zl_Drugsvr_Drugcalcprice
(
  Json_In  Clob,
  Json_Out Out Clob
) Is
  --������ȡ���ҩƷ��ʱ�� 
  --���      json
  --input              ����������Ҫ�����Ĵ������м��
  --   price_ddigits  N   1  ���С��λ��  
  --  drug_list       ҩƷ��ϸ��Ϣ��֧�ֶ����[����]
  --    drug_id       N   1  ҩƷid
  --    pharmacy_id   N   1  ҩ��id
  --    send_num      N   1  ������Ҫ���ͻ������¿�ҩƷ������  
  --����      json
  --output
  --  code     C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message  C 1 Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  fee_list       ʱ����ϸ��Ϣ��֧�ֶ����[����]
  --    drug_id      N   1  ҩƷid
  --    pharmacy_id  N   1  ҩ��id
  --    send_num     N   1  ������Ҫ���ͻ������¿�ҩƷ������ 
  --    price        N   1  ʱ��
  j_List Pljson_List;
  j_Temp PLJson;

  n_Medioutmode Number; --����ҩƷ���ⷽʽ
  n_Decprice    Number; --���С��λ��  

  n_ҩƷid Number;
  n_ҩ��id Number;
  n_����   Number(16, 5);
  n_ʱ��   Number(16, 5);

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;

  j_Jsonin PLJson;
  j_Json   PLJson;

  Function Calcdrugprice
  (
    ҩ��id_In   ҩƷ���.�ⷿid%Type,
    ҩƷid_In   ҩƷ���.ҩƷid%Type,
    ����_In     ҩƷ���.��������%Type,
    �����㷨_In Number,
    Decprice_In Number,
    Dec_In      Number
  ) Return Number Is
    --����:���ر��ҩƷ��ʱ�� 
    --����: 
    --     ����_In ������Ҫ���ͻ������¿�ҩƷ������ 
    --     �����㷨_In 0-�������Ƚ��ȳ���1-��Ч������ȳ� 
    --     Decprice_In ���С��λ�� 
    --     Dec_In �۸�С��λ 
    n_ʱ��     Number(16, 5);
    n_����ʱ�� Number(16, 5);
    n_�ܽ��   Number;
    n_������   Number;
    n_��ǰ���� Number;
    n_Cnt      Number;
  Begin
    If ����_In <= 0 Then
      Return 0;
    End If;
  
    n_�ܽ�� := 0;
    n_������ := ����_In;
    For Rs In (Select Nvl(����, 0) As ����, Nvl(��������, 0) As ���,
                      Nvl(���ۼ�, Nvl(Decode(Nvl(ʵ������, 0), 0, 0, ʵ�ʽ�� / ʵ������), 0)) As ʱ��
               From ҩƷ���
               Where �ⷿid = ҩ��id_In And ҩƷid = ҩƷid_In And Nvl(��������, 0) > 0 And ���� = 1 And
                     (Nvl(����, 0) = 0 Or Ч�� Is Null Or Ч�� > Trunc(Sysdate))
               Order By Decode(�����㷨_In, 1, Ч��, To_Date('2008-08-08', 'yyyy-mm-dd')), Decode(�����㷨_In, 2, �ϴ�����, Null),
                        Nvl(����, 0)) Loop
      --��һ�����ε�ʱ�� 
      n_Cnt := n_Cnt + 1;
      If n_Cnt = 1 Then
        n_����ʱ�� := Round(Rs.ʱ��, Decprice_In);
      End If;
    
      If n_������ = 0 Then
        Exit;
      End If;
    
      If n_������ <= Rs.��� Then
        n_��ǰ���� := n_������;
      Else
        n_��ǰ���� := Rs.���;
      End If;
    
      n_�ܽ�� := n_�ܽ�� + Round(n_��ǰ���� * Round(Rs.ʱ��, Decprice_In), Dec_In);
      n_������ := n_������ - n_��ǰ����;
    
      If n_������ = 0 Then
        Exit;
      End If;
    End Loop;
  
    If n_������ <> 0 Then
      -- ��治��,ֻ�漰һ������ʱ������ʱ��Ϊ׼�������Ե�һ������ƽ���۶������� 
      If n_Cnt = 1 Then
        n_ʱ�� := n_����ʱ��;
      Else
        n_ʱ�� := 0;
      End If;
    Else
      If n_Cnt = 1 Then
        n_ʱ�� := n_����ʱ��;
      Else
        n_ʱ�� := Round(n_�ܽ�� / ����_In, Decprice_In);
      End If;
    End If;
  
    Return n_ʱ��;
  End Calcdrugprice;
Begin
  --�������
  j_Jsonin   := PLJson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_Decprice := j_Json.Get_Number('price_ddigits');
  j_List     := j_Json.Get_Pljson_List('drug_list');

  --��ȡ����
  n_Medioutmode := Nvl(zl_GetSysParameter(150), 0);

  --ѭ����ȡʱ��
  v_Jtmp := Null;
  For I In 1 .. j_List.Count Loop
    j_Temp   := PLJson(j_List.Get(I));
    n_ҩƷid := j_Temp.Get_Number('drug_id');
    n_����   := Nvl(j_Temp.Get_Number('send_num'), 0);
    n_ҩ��id := j_Temp.Get_Number('pharmacy_id');
  
    n_ʱ�� := Nvl(Calcdrugprice(n_ҩ��id, n_ҩƷid, n_����, n_Medioutmode, n_Decprice, n_Decprice), 0);
  
    v_Jtmp := v_Jtmp || ',';
    zlJsonPutValue(v_Jtmp, 'drug_id', n_ҩƷid, 1, 1);
    zlJsonPutValue(v_Jtmp, 'pharmacy_id', n_ҩ��id, 1);
    zlJsonPutValue(v_Jtmp, 'send_num', n_����, 1);
    zlJsonPutValue(v_Jtmp, 'price', n_ʱ��, 1, 2);
  
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
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_list":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Drugcalcprice;
/
Create Or Replace Procedure Zl_Drugsvr_Rcplock
(
  Json_In  Clob,
  Json_Out Out Clob
) Is
  --���ܣ�������鳷�� 
  --��Σ�json��ʽ
  --input
  ----rcp_ids C 1 ����id
  --���Σ�json��ʽ
  --output
  ----code 0-ʧ�� 1-�ɹ�
  ----message �ɹ���ʧ�ܺ󷵻ص���Ϣ
  ----lockadvice_ids ������ҽ��id
  c_Rcpids Clob;
  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_Rcpid Collection_Type;
  I         Number;

  n_Count Number;
  v_Error Varchar2(255);
  Err_Custom Exception;

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --�������
  j_Jsonin := PLJson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  c_Rcpids := j_Json.Get_Clob('rcp_ids');

  I := 0;
  While c_Rcpids Is Not Null Loop
    If Length(c_Rcpids) <= 4000 Then
      Col_Rcpid(I) := c_Rcpids;
      c_Rcpids := Null;
    Else
      Col_Rcpid(I) := Substr(c_Rcpids, 1, Instr(c_Rcpids, ',', 3980) - 1);
      c_Rcpids := Substr(c_Rcpids, Instr(c_Rcpids, ',', 3980) + 1);
    End If;
    I := I + 1;
  End Loop;

  For I In 0 .. Col_Rcpid.Count - 1 Loop
    For r_Info In (Select /* +RULE*/
                   Distinct b.Id, b.״̬
                   From ���������ϸ A, ��������¼ B, Table(f_Num2List(Col_Rcpid(I), ',')) C
                   Where a.��id = b.Id And a.ҽ��id = c.Column_Value And a.����ύ = 1 And
                         (b.״̬ Between 0 And 1 Or b.״̬ Is Null) And b.����� Is Null) Loop
    
      Select Count(1) Into n_Count From ��������¼ Where ID = r_Info.Id And �����û� Is Not Null;
    
      If n_Count = 0 Then
        --δ���� 
        If Nvl(r_Info.״̬, 0) = 0 Then
          --δ��飬ֱ��ɾ����¼ 
          Delete ��������¼ Where ID = r_Info.Id And (״̬ = 0 Or ״̬ Is Null);
        Elsif r_Info.״̬ = 1 Then
          --����飬����״̬ 
          Update ��������¼ Set ״̬ = ״̬ + 10 Where ID = r_Info.Id And ״̬ = 1;
        End If;
      
      Else
        --������  
        v_Jtmp := v_Jtmp || ',' || r_Info.Id;
      
        If Length(v_Jtmp) > 30000 Then
          If c_Jtmp Is Null Then
            c_Jtmp := Substr(v_Jtmp, 2);
          Else
            c_Jtmp := c_Jtmp || v_Jtmp;
          End If;
          v_Jtmp := Null;
        End If;
      End If;
    End Loop;
  End Loop;

  If c_Jtmp Is Null Then
    c_Jtmp := Substr(v_Jtmp, 2);
  Else
    c_Jtmp := c_Jtmp || v_Jtmp;
  End If;

  If c_Jtmp Is Not Null Then
    v_Error := '������ҽ���д�����������ҽ�������ڽ��д�����飬������ɾ����';
    Raise Err_Custom;
  End If;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","lockadvice_ids":"' || c_Jtmp || '"}}';
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugsvr_Rcplock;
/
Create Or Replace Procedure Zl_Drugsvr_Getnotsendrec
(
  Json_In  Varchar,
  Json_Out Out Varchar
) Is
  --------------------------------------------------------------------------- 
  --���ܣ���ȡδ��ҩƷ��¼
  --��Σ�JSON��ʽ
  --input
  --  billtypes             C  1 �������ͣ������Ӣ�Ķ��ŷָ�: 1-�շѴ���;2-���ʵ�����;3-���ʱ���
  --  charge_tag            N  1 �շѱ�־:0-δ�շѻ���ʻ���;1-���շѻ����
  --  fee_source            C  1 ������Դ�������Ӣ�Ķ��ŷָ�:1-����,2-סԺ,4-���
  --  start_time            C  0 ��ʼʱ��:yyyy-mm-dd hh:mi:ss
  --  end_time              C  0 ����ʱ��:yyyy-mm-dd hh:mi:ss
  --���Σ�JSON��ʽ
  --output
  --  code                  N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --  message               C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  data[]
  --    billtype            N  1 ��������:1-�շѴ���,2-���ʵ�����,3-���ʱ���
  --    rcp_no              C  1 ��������
  --    pharmacy_id         N  1 ҩ��ID
  ---------------------------------------------------------------------------
  v_�������� Varchar2(100);
  n_�շѱ�־ Number(1);
  v_������Դ Varchar2(100);
  d_��ʼʱ�� Date;
  d_����ʱ�� Date;

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --�������
  j_Jsonin   := PLJson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  v_�������� := j_Json.Get_String('billtypes');
  n_�շѱ�־ := j_Json.Get_Number('charge_tag');
  v_������Դ := j_Json.Get_String('fee_source');
  d_��ʼʱ�� := To_Date(j_Json.Get_String('start_time'), 'yyyy-mm-dd hh24:mi:ss');
  d_����ʱ�� := To_Date(j_Json.Get_String('end_time'), 'yyyy-mm-dd hh24:mi:ss');

  v_�������� := ',' || v_�������� || ',';
  v_�������� := Replace(v_��������, ',1,', ',8,');
  v_�������� := Replace(v_��������, ',2,', ',9,');
  v_�������� := Replace(v_��������, ',3,', ',10,');

  For r_ҩƷ In (Select Decode(b.����, 8, 1, 9, 2, 10, 3) As ��������, b.No, b.�ⷿid
               From δ��ҩƷ��¼ B
               Where Instr(v_��������, ',' || b.���� || ',') > 0 And Nvl(b.���շ�, 0) = n_�շѱ�־ And
                     b.�������� Between Nvl(d_��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And Nvl(d_����ʱ��, Sysdate) And Exists
                (Select 1
                      From ҩƷ�շ���¼
                      Where ���� = b.���� And NO = b.No And
                            (Instr(',' || v_������Դ || ',', ',' || ������Դ || ',') > 0 Or ������Դ Is Null))) Loop
  
    v_Jtmp := v_Jtmp || ',{"billtype":' || r_ҩƷ.��������;
    v_Jtmp := v_Jtmp || ',"rcp_no":"' || r_ҩƷ.No || '"';
    v_Jtmp := v_Jtmp || ',"pharmacy_id":' || Nvl(r_ҩƷ.�ⷿid, 0);
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
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getnotsendrec;
/

Create Or Replace Procedure Zl_Drugsvr_Odr_Check
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ڷ����ջ�ҩƷ���ȡ���ͼ��
  --��Σ�Json_In:��ʽ
  --     chk_type                    N 1 ��鷽ʽ��0-��ȡ�б�1-�ж���ҩ�Ƿ��Ѿ���ҩ
  --     item_list[]�б�
  --               order_id          N 1 ҽ��ID
  --               rcp_nos           C 1 ���ݺ�ƴ��
  --����: Json_Out,��ʽ����
  --  output
  --    code                          N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    data{}
  --         isexist                        N 1 �Ƿ���ѷ�ҩ����ҩ
  --         item_list[]
  --             rcpdtl_id                  N 1 ������ϸid
  --             sended_num                 N 1 �ѷ�ҩƷ����
  --             order_id                   N 1 ҽ��id
  --             drug_id                    N 1 ҩƷid
  ---------------------------------------------------------------------------
  j_Input        Pljson;
  j_Item         Pljson;
  j_List         Pljson_List := Pljson_List();
  n_ҽ��id       Number(18);
  n_��鷽ʽ     Number;
  v_��ҩ���ڷ�ҩ Varchar2(3000);
  n_Count        Number;
  v_Nos          Varchar2(30000);
  v_List         Varchar2(32767);
  v_Json_Out     Varchar2(32767);
  v_Data_Out     Varchar2(32767);
  v_Error        Varchar2(255);
  Err_Custom Exception;
Begin
  j_Item     := Pljson(Json_In);
  j_Input    := j_Item.Get_Pljson('input');
  n_��鷽ʽ := j_Input.Get_Number('chk_type');
  j_List     := j_Input.Get_Pljson_List('item_list');
  If j_List Is Not Null Then
    j_Item := Pljson();
    For I In 1 .. j_List.Count Loop
      j_Item   := Pljson(j_List.Get(I));
      n_ҽ��id := j_Item.Get_Number('order_id');
      v_Nos    := j_Item.Get_String('rcp_nos');
      j_Item   := Pljson();
      If n_��鷽ʽ = 1 Then
        Select Count(1)
        Into n_Count
        From (Select /*+cardinality(j,10)*/
                a.No, a.����id, a.ҩƷid, Sum(Nvl(a.����, 1) * Decode(a.�����, Null, 0, a.ʵ������)) As �ѷ�����
               From ҩƷ�շ���¼ A, Table(f_Str2list(v_Nos)) J
               Where a.No = j.Column_Value And a.ҽ��id = n_ҽ��id
               Group By a.No, a.����id, a.ҩƷid) A
        Where a.�ѷ����� > 0;
        If n_Count > 0 Then
          v_��ҩ���ڷ�ҩ := ',"isexist":1';
          Exit;
        End If;
      Else
        For c_ҩƷ In (Select /*+cardinality(j,10)*/
                      a.No, a.����id, a.ҩƷid, Sum(Nvl(a.����, 1) * Decode(a.�����, Null, 0, a.ʵ������)) As �ѷ�����
                     From ҩƷ�շ���¼ A, Table(f_Str2list(v_Nos)) J
                     Where a.No = j.Column_Value And a.ҽ��id = n_ҽ��id
                     Group By a.No, a.����id, a.ҩƷid) Loop
        
          v_List := v_List || ',{"rcpdtl_id":' || c_ҩƷ.����id;
          v_List := v_List || ',"sended_num":' || Zljsonstr(c_ҩƷ.�ѷ�����, 1);
          v_List := v_List || ',"order_id":' || n_ҽ��id;
          v_List := v_List || ',"drug_id":' || c_ҩƷ.ҩƷid;
          v_List := v_List || '}';
        
        End Loop;
      End If;
    End Loop;
  
    If v_List Is Not Null Then
      v_List := ',"item_list":[' || Substr(v_List, 2) || ']';
    End If;
  
    v_Data_Out := v_��ҩ���ڷ�ҩ || v_List;
    v_Json_Out := '{"code":1,"message":"�ɹ�","data":{';
    v_Json_Out := v_Json_Out || Substr(v_Data_Out, 2);
    v_Json_Out := v_Json_Out || '}}';
    Json_Out   := '{"output":' || v_Json_Out || '}';
  
  End If;
Exception
  When Err_Custom Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(v_Error) || '"}}';
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Odr_Check;
/
Create Or Replace Procedure Zl_Drugsvr_Overdue_Recovery
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ڷ����ջ�ҩƷ��ش���
  --��Σ�Json_In:��ʽ
  --  input
  --     operator_name                      C 1 ����Ա����
  --     operator_time                      C 1 ����ʱ��
  --     item_list[]ҩƷɾ���б�
  --                  rcpdtl_id              N 1 ������ϸid,Ŀǰ����ķ���ID
  --                  chargeoffs_num         N 1 ��������

  --     roll_list[]�����ջ��б�
  --                  clinic_type            C 1 ҽ���������
  --                  rcp_no                 C 1 ������,���õ���
  --                  rcpdtl_id              N 1 ������ϸID
  --                  rcpdtl_id_old          N 1 ������ϸID,ԭʼ��ϸid
  --                  packages_num           N 1 ��ҩ����
  --                  send_num               N 1 ��ҩ����
  --                  item_type              C 1 �շ���Ŀ���

  --����: Json_Out,��ʽ����
  --  output
  --    code                          N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Input      Pljson;
  j_Item       Pljson;
  j_List       Pljson_List := Pljson_List();
  n_��Һid     ҩƷ�շ���¼.Id%Type;
  n_������ϸid ҩƷ�շ���¼.Id%Type;
  n_��������   ҩƷ�շ���¼.��д����%Type;

  v_��Ա����  Varchar2(3000);
  �ջ�ʱ��_In Date;
  v_Dec       Number;
  v_�շ����  Number;
  No_In       Varchar2(3000);
  v_����id    Number;
  n_�շѱ�־  Number;
  Old_����id  Number;
  v_��ǰ����  Number;
  v_��ǰ����  Number;
  v_�������  Varchar2(3000);
  n_����      Number;
  n_Count     Number;
  Cursor c_Drug Is
    Select b.����, Nvl(x.ҩ������, 0) As ����, b.����, b.Ч��, x.���Ч��, b.Id As �շ�id, b.����id, b.��ҳid, b.�ⷿid, b.����, b.����, b.�Է�����id,
           b.���
    From ҩƷ�շ���¼ B, ҩƷ��� X
    Where b.����id = Old_����id And b.ҩƷid = x.ҩƷid;

  Procedure �����շ���¼_Insert
  (
    ����id_In     Number,
    ����_In       ҩƷ�շ���¼.����%Type,
    ����_In       ҩƷ���.ҩ������%Type,
    ����_In       ҩƷ�շ���¼.����%Type,
    Ч��_In       ҩƷ�շ���¼.Ч��%Type,
    ���Ч��_In   ҩƷ���.���Ч��%Type,
    �շ�id_In     ҩƷ�շ���¼.Id%Type,
    ����id_In     ҩƷ�շ���¼.�Է�����id%Type,
    ��ҳid_In     ҩƷ�շ���¼.�Է�����id%Type,
    �ⷿid_In     ҩƷ�շ���¼.�ⷿid%Type,
    ����_In       ҩƷ�շ���¼.����%Type,
    ����_In       Varchar2,
    �Է�����id_In ҩƷ�շ���¼.�Է�����id%Type,
    P����         ҩƷ�շ���¼.����%Type,
    P����         ҩƷ�շ���¼.��д����%Type,
    P���         ҩƷ�շ���¼.���%Type
  ) Is
    v_����   ҩƷ�շ���¼.����%Type;
    v_Ч��   ҩƷ�շ���¼.Ч��%Type;
    v_����   ҩƷ�շ���¼.����%Type;
    v_Lngid  Number;
    v_���ȼ� ���.���ȼ�%Type; --ȡ������ȼ�,���ⲿ����
  Begin
    --ȷ������
    If Nvl(����_In, 0) <> 0 And ����_In = 0 Then
      --ԭ����,�ֲ�����
      v_���� := Null;
      v_���� := ����_In;
      v_Ч�� := Ч��_In;
    Elsif Nvl(����_In, 0) = 0 And ����_In = 1 Then
      --ԭ������,�ַ���
      Select ҩƷ�շ���¼_Id.Nextval Into v_���� From Dual;
      Select To_Char(Sysdate, 'YYYYMMDD') Into v_���� From Dual;
      If ���Ч��_In Is Not Null Then
        v_Ч�� := Trunc(Sysdate + ���Ч��_In * 30);
      Else
        v_Ч�� := Null;
      End If;
    Else
      v_���� := ����_In;
      v_���� := ����_In;
      v_Ч�� := Ч��_In;
    End If;
  
    --����,�Ա�,����,��������,���֤��,����ID,��ҳID,
    --���˿���id,���˲���id,Ӥ�����
    --,������Դ,ҽ��id,���,��������,Ƥ�Խ��,�������,���շ�,������Դ
    Select ҩƷ�շ���¼_Id.Nextval Into v_Lngid From Dual;
    Insert Into ҩƷ�շ���¼
      (ID, ��¼״̬, ����, NO, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ����, ��д����, ʵ������, ���ۼ�, ���۽��, ժҪ, ������, ��������,
       ����id, ����, Ƶ��, �÷�, ��ҩ��λid, ��������, ��׼�ĺ�, ���Ч��, ����, �Ա�, ����, ��������, ���֤��, ����id, ��ҳid, ���˿���id, ���˲���id, Ӥ�����, ������Դ, ҽ��id,
       ���, ��������, Ƥ�Խ��, �������, ���շ�, ������Դ)
      Select v_Lngid, 1, ����, No_In, v_�շ����, �ⷿid, �Է�����id, ������id, -1, ҩƷid, Nvl(v_����, 0), ����, v_����, v_Ч��, P����, -1 * P����,
             -1 * P����, ���ۼ�, Round(-1 * P���� * P���� * ���ۼ�, v_Dec), '���ڷ����ջ�', v_��Ա����, �ջ�ʱ��_In, ����id_In, ����, Ƶ��, �÷�, ��ҩ��λid,
             ��������, ��׼�ĺ�, ���Ч��, ����, �Ա�, ����, ��������, ���֤��, ����id, ��ҳid, ���˿���id, ���˲���id, Ӥ�����, ������Դ, ҽ��id, ���, ��������, Ƥ�Խ��,
             �������, n_�շѱ�־, ������Դ
      From ҩƷ�շ���¼
      Where ID = �շ�id_In;
  
    Zl_δ��ҩƷ��¼_Insert(v_Lngid);
  
    Zl_ҩƷ���_Update(v_Lngid, 0, 1);
  
    --δ��ҩƷ��¼
    Update δ��ҩƷ��¼
    Set ����id = ����id_In, ��ҳid = ��ҳid_In, ���� = ����_In
    Where ���� = ����_In And NO = No_In And �ⷿid + 0 = �ⷿid_In;
  
    If Sql%RowCount = 0 Then
    
      If P��� Is Not Null Then
        Select Max(b.���ȼ�) Into v_���ȼ� From ��� B Where b.���� = P���;
      End If;
    
      Insert Into δ��ҩƷ��¼
        (����, NO, ����id, ��ҳid, ����, ���ȼ�, �Է�����id, �ⷿid, ��������, ���շ�, ��ӡ״̬)
      Values
        (����_In, No_In, ����id_In, ��ҳid_In, ����_In, v_���ȼ�, �Է�����id_In, �ⷿid_In, �ջ�ʱ��_In, n_�շѱ�־, 0);
    End If;
  
  End �����շ���¼_Insert;

Begin
  j_Item  := Pljson(Json_In);
  j_Input := j_Item.Get_Pljson('input');

  --ҩƷɾ���б�
  j_List := j_Input.Get_Pljson_List('item_list');
  If j_List Is Not Null Then
    For I In 1 .. j_List.Count Loop
      j_Item       := Pljson();
      j_Item       := Pljson(j_List.Get(I));
      n_������ϸid := j_Item.Get_Number('rcpdtl_id');
      n_��������   := j_Item.Get_Number('chargeoffs_num');
      n_��Һid     := j_Item.Get_Number('pivas_id');
      Zl_ҩƷ�շ���¼_�����˷�_s(n_������ϸid, n_��������, n_��Һid, 1);
    End Loop;
  End If;

  j_List := Pljson_List();
  j_List := j_Input.Get_Pljson_List('roll_list');
  If j_List Is Not Null Then
  
    v_��Ա����  := j_Input.Get_String('operator_name');
    �ջ�ʱ��_In := To_Date(j_Input.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');
  
    --���С��λ��
    Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into v_Dec From Dual;
  
    For I In 1 .. j_List.Count Loop
      j_Item := Pljson();
      j_Item := Pljson(j_List.Get(I));
    
      Select Nvl(Max(���), 0) + 1 Into v_�շ���� From ҩƷ�շ���¼ Where ���� = 9 And ��¼״̬ = 1 And NO = No_In;
    
      v_����id   := j_Item.Get_Number('rcpdtl_id');
      No_In      := j_Item.Get_String('rcp_no');
      v_��ǰ���� := j_Item.Get_Number('packages_num');
      v_��ǰ���� := j_Item.Get_Number('send_num');
      v_������� := j_Item.Get_String('clinic_type'); --ҽ����¼�е��������
      Old_����id := j_Item.Get_Number('rcpdtl_id_old');
      n_�շѱ�־ := j_Item.Get_Number('charge_tag');
    
      For r_Drug In c_Drug Loop
      
        If v_������� In ('5', '6', '7') Then
          �����շ���¼_Insert(v_����id, r_Drug.����, r_Drug.����, r_Drug.����, r_Drug.Ч��, r_Drug.���Ч��, r_Drug.�շ�id, r_Drug.����id,
                        r_Drug.��ҳid, r_Drug.�ⷿid, r_Drug.����, r_Drug.����, r_Drug.�Է�����id, v_��ǰ����, v_��ǰ����, r_Drug.���);
        Else
          n_���� := v_��ǰ����;
          For r_Otherdrug In (Select b.����, Nvl(x.ҩ������, 0) As ����, Nvl(b.����, 1) * b.ʵ������ As ����, b.����, b.Ч��, x.���Ч��,
                                     b.Id As �շ�id, b.����id, b.��ҳid, b.�ⷿid, b.����, b.����, b.�Է�����id, b.���
                              From ҩƷ�շ���¼ B, ҩƷ��� X
                              Where b.����id = Old_����id And (b.��¼״̬ = 1 Or Mod(b.��¼״̬, 3) = 0) And b.ҩƷid = x.ҩƷid
                              Order By b.Id Desc) Loop
            If n_���� > 0 Then
              n_Count := r_Otherdrug.����;
              If n_���� < n_Count Then
                n_Count := n_����;
              End If;
              �����շ���¼_Insert(v_����id, r_Otherdrug.����, r_Otherdrug.����, r_Otherdrug.����, r_Otherdrug.Ч��, r_Otherdrug.���Ч��,
                            r_Otherdrug.�շ�id, r_Otherdrug.����id, r_Otherdrug.��ҳid, r_Otherdrug.�ⷿid, r_Otherdrug.����,
                            r_Otherdrug.����, r_Otherdrug.�Է�����id, 1, n_Count, r_Otherdrug.���);
              n_���� := n_���� - r_Otherdrug.����;
            End If;
          End Loop;
        End If;
      End Loop;
    End Loop;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugsvr_Overdue_Recovery;
/