Create Or Replace Procedure Zl_StuffSvr_ExecutePrice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --ִ�е���
  ---------------------------------------------------------------------------
  --input      ��������ۼۣ��ɱ����Ƿ��������Ч��δִ�еļ۸����������ִ�е���
  --  stuff_id      N    ����id
  --output      
  --  code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message  C  1  Ӧ����Ϣ��
  ---------------------------------------------------------------------------
  j_Jsonin Pljson;
  j_Json   Pljson;

  n_����id ��������.����id%Type;

  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_����id := j_Json.Get_Number('stuff_id');

  If n_����id = 0 Then
    v_Err_Msg := 'δ��������ID��Ϣ��';
    Raise Err_Item;
  End If;

  For r_���� In (Select Distinct i.Id As ����id
               From �շ���ĿĿ¼ I, �շѼ�Ŀ N, �������� P
               Where i.Id = n.�շ�ϸĿid And i.Id = p.����id And
                     (i.����ʱ�� Is Null Or i.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) And n.�䶯ԭ�� = 0 And p.����id = n_����id And
                     Sysdate > n.ִ������
               Union
               Select Distinct a.ҩƷid As ����id
               From ҩƷ�۸��¼ A
               Where a.��¼״̬ = 0 And a.ҩƷid = n_����id And a.ִ������ <= Sysdate
               Order By ����id) Loop
  
    n_����id := r_����.����id;
    Exit;
  End Loop;

  If n_����id > 0 Then
    Zl_�����շ���¼_Adjust(n_����id);
  End If;

  Json_Out := '{"output":{"code": 1,"message":"�ɹ�"}}';
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_StuffSvr_ExecutePrice;
/

Create Or Replace Procedure Zl_StuffSvr_CheckHCostExistRec
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --�жϸ�ֵ�����Ƿ����ʹ�ü�¼
  ---------------------------------------------------------------------------
  --input      �жϸ�ֵ�����Ƿ����ʹ�ü�¼
  --  stuff_id      N    ����id
  --output      
  --  code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message  C  1  Ӧ����Ϣ��
  --  isexist  N 1 �Ƿ����: 1-����;0-������
  ---------------------------------------------------------------------------
  j_Jsonin Pljson;
  j_Json   Pljson;
  n_����id ��������.����id%Type;
  n_Exist  Number(1);
Begin
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_����id := j_Json.Get_Number('stuff_id');

  Select Count(1)
  Into n_Exist
  From ҩƷ�շ���¼ A, �շ���¼������Ϣ B
  Where a.ҩƷid = n_����id And a.Id = b.�շ�id And Rownum < 2;

  Json_Out := '{"output":{"code": 1,"message":"�ɹ�","isexist":' || n_Exist || '}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_StuffSvr_CheckHCostExistRec;
/

Create Or Replace Procedure Zl_StuffSvr_GetStockShow
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡָ���ⷿ�Ŀ�����ݣ�������ʾ
  --��Σ�Json_In:��ʽ
  --  input
  --    warehouse_ids        C   1   �ⷿID��
  --����: Json_Out,��ʽ����
  --  output
  --    code                N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    item_list
  --      stuff_id              N   1   ����ID
  --      warehouse_id          N   1   �ⷿID
  --      stock                N   1   ��������
  --      real_stock          N  1 ʵ�ʿ��
  --      avg_price           N  1 ƽ���ۼ�
  --      avg_cost            N  1 ƽ���ɱ���
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
  v_�ⷿids := j_Json.Get_String('warehouse_ids');

  If Nvl(v_�ⷿids, 0) = 0 Then
    Json_Out := Zljsonout('δ������ؿⷿ��Ϣ');
    Return;
  End If;

  For c_��� In (Select a.�ⷿid, a.ҩƷid, Sum(Nvl(a.��������, 0)) As ��������, Sum(Nvl(a.ʵ������, 0)) As ʵ������,
                      Decode(Sum(Nvl(a.ʵ������, 0)), 0, Max(a.���ۼ�), Sum(Nvl(a.ʵ�ʽ��, 0)) / Sum(Nvl(a.ʵ������, 0))) As ƽ���ۼ�,
                      Decode(Sum(Nvl(a.ʵ������, 0)), 0, Max(a.ƽ���ɱ���),
                              (Sum(Nvl(a.ʵ�ʽ��, 0)) - Sum(Nvl(a.ʵ�ʲ��, 0))) / Sum(Nvl(a.ʵ������, 0))) As ƽ���ɱ���
               From ҩƷ��� A
               Where a.���� = 1 And Instr(',' || v_�ⷿids || ',', ',' || a.�ⷿid || ',') > 0
               Group By a.�ⷿid, a.ҩƷid) Loop
  
    v_Jtmp := v_Jtmp || ',';
    Zljsonputvalue(v_Jtmp, 'stuff_id', c_���.ҩƷid, 1, 1);
    Zljsonputvalue(v_Jtmp, 'warehouse_id', c_���.�ⷿid, 1);
    Zljsonputvalue(v_Jtmp, 'stock', c_���.��������, 1);
    Zljsonputvalue(v_Jtmp, 'real_stock', c_���.ʵ������, 1);
    Zljsonputvalue(v_Jtmp, 'avg_price', c_���.ƽ���ۼ�, 1);
    Zljsonputvalue(v_Jtmp, 'avg_cost', c_���.ƽ���ɱ���, 1, 2);
  
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
End Zl_StuffSvr_GetStockShow;
/

Create Or Replace Procedure Zl_StuffSvr_GetCostPriceAdjust
(
  Json_In  Varchar2,
  Json_Out Out Clob
) Is
  ---------------------------------------------------------------------------
  --���ܣ���ȡ���ĳɱ��۵��ۼ�¼
  --input      
  --  stuff_id      N   1 ����id
  --  show_unit    N   1   ��ʾ��λ:0-ɢװ��λ;1-�ⷿ��λ
  --output      
  --  code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message  C  1  Ӧ����Ϣ��  
  --  price_list[]  ���ĳɱ��۵��ۼ�¼
  --     stuff_id   N 1  ����ID
  --     stuff_name   C 1  ������Ϣ
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
  --     stuff_revoke_time  C 1 ����ʱ��
  --     node_no      C    0  վ�����   
  --     is_stock    N   1 �Ƿ��п������  0-��1-��
  ---------------------------------------------------------------------------
  j_Jsonin Pljson;
  j_Json   Pljson;
  n_����id ��������.����ID%Type;
  n_��λ   Number(1);

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;
Begin
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_����id := j_Json.Get_Number('stuff_id');
  n_��λ   := j_Json.Get_Number('show_unit');

  v_Jtmp := Null;
  For r_Costprice In (Select distinct b.No, i.Id As ����id, '[' || i.���� || ']' || i.���� || ' ' || i.��� || ' ' || i.���� As ����,
                             p.���� As �ⷿ, a.����, a.Ч��, a.����, Decode(n_��λ, 0, i.���㵥λ, s.��װ��λ) As ��λ,
                             Decode(n_��λ, 0, a.ԭ��, a.ԭ�� * Nvl(s.����ϵ��, 1)) As ԭ�ɱ���,
                             Decode(n_��λ, 0, a.�ּ�, a.�ּ� * Nvl(s.����ϵ��, 1)) As �ɱ���, a.ִ������, a.����˵��, i.����ʱ��, i.վ��,
                             Decode(k.�ⷿid, Null, 0, 1) As ���
                      From ҩƷ�շ���¼ B, �շ���ĿĿ¼ I, �������� S, ���ű� P, ҩƷ�۸��¼ A, ҩƷ��� K
                      Where a.�۸����� = 2 And a.�շ�id = b.Id(+) And a.ҩƷid = i.Id And i.Id = s.����id And a.�ⷿid = p.Id(+) And
                            s.����id = n_����id And k.����(+) = 1 And k.�ⷿid(+) = a.�ⷿid And k.ҩƷid(+) = a.ҩƷid And
                            k.����(+) = a.����
                      Order By '[' || i.���� || ']' || i.���� || ' ' || i.��� || ' ' || i.����, p.����, a.����, a.ִ������ Desc, b.No Desc) Loop
  
    v_Jtmp := v_Jtmp || ',';
    Zljsonputvalue(v_Jtmp, 'stuff_id', r_Costprice.����id, 1, 1);
    Zljsonputvalue(v_Jtmp, 'stuff_name', r_Costprice.����, 0);
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
    Zljsonputvalue(v_Jtmp, 'stuff_revoke_time', r_Costprice.����ʱ��, 0);
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
End Zl_StuffSvr_GetCostPriceAdjust;
/

Create Or Replace Procedure Zl_StuffSvr_AdjustPriceType
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --ִ�е���
  ---------------------------------------------------------------------------
  --input      ���ļ۸����Ե���ʱ�����ĵ���ӯ���Ϳ��仯���ݴ���
  --    item_list[]         �����б�
  --       stuff_id      N    ����id
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

  n_����id     ��������.����id%Type;
  n_����id     �շѼ�Ŀ.������Ŀid%Type;
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
  n_�۸��¼     Number(1);
  v_���         �շ���ĿĿ¼.���%Type;

  --����->ʱ�ۺ���¼۸��¼��ֵ
  Cursor c_Priceadjust Is
    Select s.ҩƷid, s.�ⷿid, Nvl(s.����, 0) As ����, s.�ϴι�Ӧ��id As ��Ӧ��id, s.�ϴ����� As ����, s.Ч��, s.�ϴβ��� As ����,
           Nvl(s.ʵ������, 0) As ʵ������, s.�ϴο��� As ����, Nvl(s.ʵ�ʽ��, 0) As ʵ�ʽ��, Nvl(s.ʵ�ʲ��, 0) As ʵ�ʲ��, Nvl(s.���ۼ�, 0) As ���ۼ�,
           s.ƽ���ɱ���, s.���Ч��, s.��׼�ĺ�, s.�ϴ��������� As ��������
    From ҩƷ��� S
    Where s.ҩƷid = n_����id And s.���� = 1
    Order By s.ҩƷid, s.����, s.�ⷿid;

  r_Priceadjust c_Priceadjust%RowType;
Begin
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  j_Jsonlist := j_Json.Get_Pljson_List('item_list');

  n_Count := j_Jsonlist.Count;
  If j_Jsonlist.Count = 0 Then
    Json_Out := Zljsonout('δ����������Ϣ��');
    Return;
  End If;

  For I In 1 .. j_Jsonlist.Count Loop
    j_Json := Pljson();
    j_Json := Pljson(j_Jsonlist.Get(I));
  
    n_����id     := j_Json.Get_Number('stuff_id');
    n_ԭ�۸����� := j_Json.Get_Number('price_type_old');
    n_�¼۸����� := j_Json.Get_Number('price_type_new');
  
    If n_ԭ�۸����� <> n_�¼۸����� Then
      --ȡԭ�ۺ�ԭ��id(���øù���ǰ�Ѿ��������¼۸�)
      Begin
        Select ԭ��, �ּ�, ԭ��id As �۸�id
        Into n_�շѼ�Ŀԭ��, n_�շѼ�Ŀ�ּ�, n_�۸�id
        From �շѼ�Ŀ
        Where �շ�ϸĿid = n_����id And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And �䶯ԭ�� = 1;
      Exception
        When Others Then
          n_�շѼ�Ŀԭ�� := Null;
          n_�շѼ�Ŀ�ּ� := Null;
          n_�۸�id       := Null;
      End;
      
      --ʱ��->����
      If n_ԭ�۸����� = 1 And n_�¼۸����� = 0 Then
        Select ���� Into n_��ͨ���С�� From ҩƷ���ľ��� Where ��� = 2 And ���� = 4 And ��λ = 5;
      
        --ȡ������ID
        Select ���id Into n_������id From ҩƷ�������� Where ���� = 38;        
              
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
               r_Priceadjust.����, n_���۽��, n_���۽��, '����ʱ��ת����', zl_UserName, Sysdate, r_Priceadjust.�ⷿid, 1, n_�۸�id,
               zl_UserName, Sysdate, r_Priceadjust.ʵ�ʽ��, r_Priceadjust.ʵ�ʲ��, r_Priceadjust.��Ӧ��id);
          
            Zl_ҩƷ���_Update(n_�շ�id, 2, 0);
          End If;
        End Loop;
      
        --����->ʱ��
      Elsif n_ԭ�۸����� = 0 And n_�¼۸����� = 1 Then
        For r_Priceadjust In c_Priceadjust Loop
          n_�۸��¼ := 0;
          Begin
            Select 1, �ּ�
            Into n_�۸��¼, n_ԭ��
            From ҩƷ�۸��¼
            Where ҩƷid = r_Priceadjust.ҩƷid And �ⷿid = r_Priceadjust.�ⷿid And Nvl(����, 0) = r_Priceadjust.���� And
                  ��¼״̬ = 1 And �۸����� = 1;
          Exception
            When Others Then
              n_�۸��¼ := 0;
              n_ԭ��     := n_�շѼ�Ŀԭ��;
          End;
        
          If n_�۸��¼ = 1 Then
            Zl_ҩƷ�۸��¼_Stop(1, r_Priceadjust.�ⷿid, r_Priceadjust.ҩƷid, r_Priceadjust.����, Sysdate - 1 / 24 / 60 / 60, 2);
          End If;
          Zl_ҩƷ�۸��¼_Insert(0, 1, r_Priceadjust.�ⷿid, r_Priceadjust.ҩƷid, r_Priceadjust.����, n_ԭ��, n_�շѼ�Ŀ�ּ�, Sysdate,
                           '���Ķ���תʱ��', zl_UserName, Null, r_Priceadjust.��Ӧ��id, r_Priceadjust.����, r_Priceadjust.Ч��,
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
End Zl_StuffSvr_AdjustPriceType;
/


Create Or Replace Procedure Zl_StuffSvr_CheckExistStock
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  ---------------------------------------------------------------------------
  --input      �ж������Ƿ���ڿ�����
  --  stuff_id      N  1  ����id
  --  is_item      N  1  �Ƿ�Ʒ�ֲ�ѯ��0-������ѯ��1-��Ʒ�ֲ�ѯ
  --output      
  --  code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message  C  1  Ӧ����Ϣ��
  --  isexist  N 1 �Ƿ����: 1-����;0-������
  ---------------------------------------------------------------------------
  j_Jsonin Pljson;
  j_Json   Pljson;
  n_����id ��������.����ID%Type;
  n_Ʒ��   Number(1);
  n_Exist  Number(1);
Begin
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_����id := j_Json.Get_Number('stuff_id');
  n_Ʒ��   := j_Json.Get_Number('is_item');

  If n_Ʒ�� = 0 Then
    Select Count(1) Into n_Exist From ҩƷ��� Where ҩƷid = n_����id And Rownum < 2;
  Else
    Select Count(1)
    Into n_Exist
    From �������� A, ҩƷ��� B
    Where a.����id = b.ҩƷid And a.����id = n_����id And Rownum < 2;
  End If;

  Json_Out := '{"output":{"code": 1,"message":"�ɹ�","isexist":' || n_Exist || '}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_StuffSvr_CheckExistStock;
/

Create Or Replace Procedure Zl_StuffSvr_CheckStuffExistRec
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --�ж������Ƿ�����շ���¼
  ---------------------------------------------------------------------------
  --input      �ж������Ƿ�����շ���¼
  --  stuff_id      N    ����id
  --output      
  --  code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message  C  1  Ӧ����Ϣ��
  --  isexist  N 1 �Ƿ����: 1-����;0-������
  ---------------------------------------------------------------------------
  j_Jsonin Pljson;
  j_Json   Pljson;
  n_����id ��������.����ID%Type;
  n_Exist  Number(1);
Begin
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_����id := j_Json.Get_Number('stuff_id');

  Select Count(1) Into n_Exist From ҩƷ�շ���¼ Where ҩƷid = n_����id And Rownum < 2;

  Json_Out := '{"output":{"code": 1,"message":"�ɹ�","isexist":' || n_Exist || '}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_StuffSvr_CheckStuffExistRec;
/

Create Or Replace Procedure Zl_Stuffsvr_Patiinfoupdate
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ�סԺ����ҩƷ�շ���¼������Ϣ�޸�
  --��Σ�Json_In:��ʽ
  --  input
  --   pati_id        N   1   ����id
  --   pati_name      C   1   ��������
  --   pati_sex       C   1   �����Ա�
  --   pati_age       C   1   ��������
  --   visit_id       N   1   ����id
  --����: Json_Out,��ʽ����
  --  output
  --    code                 N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message              C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  n_����id Number;
  n_����id Number;
  v_����   Varchar2(100);
  v_�Ա�   Varchar2(100);
  v_����   Varchar2(100);
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
End Zl_Stuffsvr_Patiinfoupdate;
/

Create Or Replace Procedure Zl_Stuffsvr_Newbill_Check
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
  --    stuff_id  N 1 ����id
  --    send_num  N 1 ��ҩ����
  --    warehouse_id  N 1 �ⷿid
  --    price           N       1       �ۼ�
  --    is_bakstuff N   �Ƿ񱸻�����:�и�ֵ���Ĳ���Ҫ���룬��0��ʾ�Ǹ�ֵ����ģʽ(��ɨ��ʱʹ��)
  --    bakstuff_batch  N   ������������
  --����      json
  --output      
  --  code  C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message C 1 Ӧ����Ϣ��
  -------------------------------------------------------------------------------------------------
Begin

  Zl_�������۳���_Check(Json_In, Json_Out);

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Newbill_Check;
/

Create Or Replace Procedure Zl_Stuffsvr_Send_Check
(
  Json_In  Clob,
  Json_Out Out Varchar2
) Is
  -------------------------------------------------------------------------------------------------
  --���ܣ�����������ǰ��鷢�������Ƿ���ڣ��Ƿ��Ѿ����ϣ��Ƿ��Ѿ��ܷ����Ƿ�����δ���/δ�շѵļ�¼���ϵ�
  --��Σ�json��ʽ
  --Input
  --  stuff_rec_id C   ҩƷ�շ���¼ID��,֧�ֶ��id���á�,���ָ�
  --���Σ�json��ʽ
  --Json_Out
  --  code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message  C  1  "Ӧ����Ϣ��
  --  �ɹ�ʱ���سɹ���Ϣ
  --  ʧ��ʱ���ؾ���Ĵ�����Ϣ"
  -------------------------------------------------------------------------------------------------

Begin

  Zl_������������_Check(Json_In, Json_Out);

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Send_Check;
/

Create Or Replace Procedure Zl_Stuffsvr_Checkpatiexecute
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ݲ�����Ϣ���ȡδ��������
  --��Σ�Json_In:��ʽ
  --  input
  --     check_type         N 1 ��鷽ʽ:0-�����²���ֵ���м�飻1-��������ID�ͷ��Ͽⷿ���м��
  --     pati_id            N 1 ����ID
  --     pati_pageid        N 1 ��ҳID
  --     baby_num           N 1 Ӥ�����:-1��ʾ������;0-ĸ�׵�;>0����Ӥ������
  --     fee_source         N 1 ������Դ:1-����;2-סԺ;4-���
  --     stuff_nos             �������ݺţ������磺["A0001","A0002"]
  --     warehouse_id       N 0 �ⷿID
  --����: Json_Out,��ʽ����
  --  output
  --    code                    N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                 C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    isexist                 N 1 �Ƿ����: 1-����;0-������
  --    stuff_notsend_infor     C 1 δ������Ϣ,isexist=1ʱ����
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  n_��鷽ʽ Number(1);
  v_Stuff    Varchar2(32767);
  n_����id   ҩƷ�շ���¼.����id%Type;
  n_��ҳid   ҩƷ�շ���¼.��ҳid%Type;
  n_Ӥ����� ҩƷ�շ���¼.Ӥ�����%Type;
  v_No       ҩƷ�շ���¼.No%Type;
  n_������Դ ҩƷ�շ���¼.������Դ%Type;
  j_Jsonlist Pljson_List := Pljson_List();
  n_Count    Number(18);
  l_Nos      t_StrList := t_StrList();
  v_��Ŀ     �շ���ĿĿ¼.����%Type;
  v_����     ���ű�.����%Type;
  n_�ⷿid   ҩƷ�շ���¼.�ⷿid%Type;

  Type t_δ���� Is Ref Cursor;
  c_δ���� t_δ����;

Begin
  --�������
  j_Jsonin   := PLJson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_��鷽ʽ := j_Json.Get_Number('check_type');
  n_����id   := j_Json.Get_Number('pati_id');
  n_��ҳid   := j_Json.Get_Number('pati_pageid');
  n_Ӥ����� := j_Json.Get_Number('baby_num');
  n_������Դ := j_Json.Get_Number('fee_source');
  j_Jsonlist := j_Json.Get_Pljson_List('stuff_nos');
  n_�ⷿid   := j_Json.Get_Number('warehouse_id');

  If Nvl(n_��鷽ʽ, 0) = 0 Then
    n_Count := 0;
    If j_Jsonlist Is Not Null Then
      n_Count := j_Jsonlist.Count;
    End If;
    If n_Count <> 0 Then
      For I In 1 .. n_Count Loop
        v_No := j_Jsonlist.Get_String(I);
        l_Nos.Extend();
        l_Nos(l_Nos.Count) := v_No;
      End Loop;
    
      Open c_δ���� For
        Select Distinct b.No, d.���� ��Ŀ, c.���� As ����
        From ҩƷ�շ���¼ B, ���ű� C, �շ���ĿĿ¼ D
        Where b.ҩƷid = d.Id And b.�ⷿid + 0 = c.Id(+) And b.���� In (25, 26) And Mod(b.��¼״̬, 3) = 1 And b.����� Is Null And
              b.����id = n_����id And (Nvl(b.��ҳid, 0) = Nvl(n_��ҳid, 0) Or n_������Դ <> 2) And
              (Nvl(b.Ӥ�����, 0) = Nvl(n_Ӥ�����, 0) Or Nvl(n_Ӥ�����, 0) = -1) And Nvl(b.ժҪ, '��ҽ') <> '�ܷ�' And
              b.No In (Select Column_Value From Table(l_Nos)) And Nvl(b.������Դ, 1) = n_������Դ;
    
    Else
      Open c_δ���� For
        Select Distinct b.No, d.���� ��Ŀ, c.���� As ����
        From ҩƷ�շ���¼ B, ���ű� C, �շ���ĿĿ¼ D
        Where b.ҩƷid = d.Id And b.�ⷿid + 0 = c.Id(+) And b.���� In (25, 26) And Mod(b.��¼״̬, 3) = 1 And b.����� Is Null And
              b.����id = n_����id And (Nvl(b.��ҳid, 0) = Nvl(n_��ҳid, 0) Or n_������Դ <> 2) And
              (Nvl(b.Ӥ�����, 0) = Nvl(n_Ӥ�����, 0) Or n_Ӥ����� = -1) And Nvl(b.ժҪ, '��ҽ') <> '�ܷ�' And b.������Դ = n_������Դ;
    
    End If;
  
  Elsif n_��鷽ʽ = 1 Then
    Open c_δ���� For
      Select Distinct b.No, d.���� ��Ŀ, c.���� As ����
      From ҩƷ�շ���¼ B, ���ű� C, �շ���ĿĿ¼ D
      Where b.ҩƷid = d.Id And b.�ⷿid + 0 = c.Id(+) And b.���� In (24, 25) And Mod(b.��¼״̬, 3) = 1 And b.����� Is Null And
            b.����id = n_����id And b.�ⷿid = n_�ⷿid;
  End If;

  Loop
    Fetch c_δ����
      Into v_No, v_��Ŀ, v_����;
    Exit When c_δ����%NotFound;
  
    If v_Stuff Is Not Null Then
      If Instr(Chr(13) || Chr(10) || v_Stuff || Chr(13) || Chr(10),
               Chr(13) || Chr(10) || '����[' || Nvl(v_No, '') || ']�е�' || Nvl(v_��Ŀ, '') || '����' || Nvl(v_����, '[δ������]') ||
                'δ����' || Chr(13) || Chr(10), 1) = 0 Then
        If Lengthb(v_Stuff || Chr(13) || Chr(10) || '����[' || Nvl(v_No, '') || ']�е�' || Nvl(v_��Ŀ, '') || '����' ||
                   Nvl(v_����, '[δ������]') || 'δ����') <= 1000 Then
          v_Stuff := v_Stuff || Chr(13) || Chr(10) || '����[' || Nvl(v_No, '') || ']�е�' || Nvl(v_��Ŀ, '') || '����' ||
                     Nvl(v_����, '[δ������]') || 'δ����';
        Else
          v_Stuff := v_Stuff || Chr(13) || Chr(10) || '... ...';
        End If;
      End If;
    Else
      v_Stuff := '����[' || Nvl(v_No, '') || ']�е�' || Nvl(v_��Ŀ, '') || '����' || Nvl(v_����, '[δ������]') || 'δ����';
    End If;
  End Loop;

  n_Count := 0;
  If v_Stuff Is Not Null Then
    v_Stuff := '����δ���������ϣ�' || Chr(13) || Chr(10) || Chr(13) || Chr(10) || v_Stuff;
    n_Count := 1;
  End If;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","isexist":' || n_Count || ',"stuff_notsend_infor":"' ||
              Zltools.Zljsonstr(v_Stuff) || '"}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Checkpatiexecute;
/

Create Or Replace Procedure Zl_Stuffsvr_Getrefusesendlist
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡ�ܷ�ҩ�嵥
  --��Σ�Json_In:��ʽ
  --  input
  --     pati_id  N 1 ����Id
  --     pati_pageids C 1 ��ҳIDs:���סԺ ���ö��ŷ���
  --����: Json_Out,��ʽ����
  --  output
  --    code                    N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                 C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    item_list []
  --        stuff_no          C   1   ���õ��ݺ�
  --        stuffdtl_id          C   1   ����ID
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  v_��ҳids Varchar2(4000);
  n_����id  ҩƷ�շ���¼.����id%Type;

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;
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

  v_Jtmp := Null;
  For r_Info In (Select NO As Stuff_No, ����id As Stuffdtl_Id
                 From ҩƷ�շ���¼
                 Where ����id = n_����id And
                       (Instr(',' || Nvl(v_��ҳids, '-') || ',', ',' || Nvl(��ҳid, 0) || ',') > 0 Or v_��ҳids Is Null) And
                       Mod(��¼״̬, 3) = 1 And Nvl(ժҪ, '��һ') = '�ܷ�' And Instr(',21,24,25,26,', ',' || ���� || ',') > 0
                 Order By NO, ����id) Loop
  
    v_Jtmp := v_Jtmp || ',{"stuff_no":"' || r_Info.Stuff_No || '"';
    v_Jtmp := v_Jtmp || ',"stuffdtl_id":' || r_Info.Stuffdtl_Id;
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
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Getrefusesendlist;
/

Create Or Replace Procedure Zl_Stuffsvr_Getexecutednum
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡ���������ѷ�������
  --��Σ�Json_In:��ʽ
  --  input
  --     billtype               N 1 ��������:1-�շѴ������ϣ�2-���ʵ��������ϣ�3-���ʱ�������
  --     stuff_nos              C 1 ���ݺ�:���Դ�����ŵ���
  --     notcontain_zero        N 1 �Ƿ񲻰����ѷ�����Ϊ0�ģ�1-��������0-����
  --     stuffdtl_ids           C 0 ������ϸids�������Ӣ�ĵĶ��ŷָ�,δ����ʱ�����ݺŲ���,����ʱ����ϸid���в���
  --     order_ids              C 0 ҽ��id�������δ����һ��ҽ��id���ŷָ�
  --����: Json_Out,��ʽ����
  --  output
  --    code                    N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                 C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    item_list[]
  --       stuff_no             C   1   NO
  --       stuffdtl_id          N   1   ����ID
  --       order_id             N   0   ҽ��id
  --       stuff_id             N   1   ����ID
  --       sended_num           N   1   �ѷ�����
  --       barcode_goods        C       ��Ʒ����
  --       barcode_inside       C       �ڲ�����
  ---------------------------------------------------------------------------
  j_Jsonin Pljson;
  j_Json   Pljson;

  n_������    Number(1);
  n_����      ҩƷ�շ���¼.����%Type;
  v_Nos       Varchar2(4000);
  c_Order_Ids Clob;

  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_��ϸids Collection_Type;
  I           Number;
  v_��ϸids   Clob;

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;
Begin
  --�������
  j_Jsonin    := Pljson(Json_In);
  j_Json      := j_Jsonin.Get_Pljson('input');
  n_����      := j_Json.Get_Number('billtype');
  v_Nos       := j_Json.Get_String('stuff_nos');
  n_������    := j_Json.Get_Number('notcontain_zero');
  c_Order_Ids := j_Json.Get_Clob('order_ids');

  If j_Json.Exist('stuffdtl_ids') Then
    v_��ϸids := j_Json.Get_Clob('stuffdtl_ids');
  End If;

  If n_���� = 1 Then
    n_���� := 24;
  Elsif n_���� = 2 Then
    n_���� := 25;
  Elsif n_���� = 3 Then
    n_���� := 26;
  Elsif c_Order_Ids Is Null Then
    If v_��ϸids Is Null Then
      Json_Out := Zljsonout('����ڵ㡾billtype���������飡');
      Return;
    End If;
  End If;

  v_Jtmp := Null;
  If v_��ϸids Is Null Then
    For r_���� In (Select /*+cardinality(j,10)*/
                  a.No, a.����id, a.ҽ��id, a.ҩƷid, Sum(Nvl(a.����, 1) * Decode(a.�����, Null, 0, a.ʵ������)) As �ѷ�����,
                  Max(��Ʒ����) As ��Ʒ����, Max(�ڲ�����) As �ڲ�����
                 From ҩƷ�շ���¼ A, Table(f_Str2list(v_Nos)) J
                 Where a.No = j.Column_Value And
                       (a.���� = 24 And n_���� = 24 Or n_���� <> 24 And Instr(',25,26,', ',' || a.���� || ',') > 0 Or
                       n_���� Is Null) And
                       (c_Order_Ids Is Null Or Instr(',' || c_Order_Ids || ',', ',' || a.ҽ��id || ',') > 0)
                 Group By a.No, a.����id, a.ҩƷid, a.ҽ��id) Loop
    
      If Not (Nvl(n_������, 0) = 1 And Nvl(r_����.�ѷ�����, 0) = 0) Then
        v_Jtmp := v_Jtmp || ',';
        Zljsonputvalue(v_Jtmp, 'stuff_no', r_����.No, 0, 1);
        Zljsonputvalue(v_Jtmp, 'stuffdtl_id', r_����.����id, 1);
        Zljsonputvalue(v_Jtmp, 'order_id', r_����.ҽ��id, 1);
        Zljsonputvalue(v_Jtmp, 'stuff_id', r_����.ҩƷid, 1);
        Zljsonputvalue(v_Jtmp, 'sended_num', Nvl(r_����.�ѷ�����, 0), 1);
        Zljsonputvalue(v_Jtmp, 'barcode_goods', r_����.��Ʒ����, 0);
        Zljsonputvalue(v_Jtmp, 'barcode_inside', r_����.�ڲ�����, 0, 2);
      
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
      For r_���� In (Select /*+cardinality(j,10)*/
                    a.No, a.����id, a.ҩƷid, Sum(Nvl(a.����, 1) * Decode(a.�����, Null, 0, a.ʵ������)) As �ѷ�����, Max(��Ʒ����) As ��Ʒ����,
                    Max(�ڲ�����) As �ڲ�����
                   From ҩƷ�շ���¼ A, Table(f_Num2list(Col_��ϸids(I))) J
                   Where a.����id = j.Column_Value
                   Group By a.No, a.����id, a.ҩƷid) Loop
      
        If Not (Nvl(n_������, 0) = 1 And Nvl(r_����.�ѷ�����, 0) = 0) Then
          v_Jtmp := v_Jtmp || ',';
          Zljsonputvalue(v_Jtmp, 'stuff_no', r_����.No, 0, 1);
          Zljsonputvalue(v_Jtmp, 'stuffdtl_id', r_����.����id, 1);
          Zljsonputvalue(v_Jtmp, 'stuff_id', r_����.ҩƷid, 1);
          Zljsonputvalue(v_Jtmp, 'sended_num', Nvl(r_����.�ѷ�����, 0), 1);
          Zljsonputvalue(v_Jtmp, 'barcode_goods', r_����.��Ʒ����, 0);
          Zljsonputvalue(v_Jtmp, 'barcode_inside', r_����.�ڲ�����, 0, 2);
        
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
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Getexecutednum;
/

Create Or Replace Procedure Zl_Stuffsvr_Getstockcheck
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡ���Ŀ���鷽ʽ
  --��Σ�Json_In:��ʽ
  --  input
  --����: Json_Out,��ʽ����
  --  output
  --    code                N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    item_list             [����]
  --        warehouse_id    N   1   �ⷿID
  --        check_type      N   1   ��鷽ʽ��0-����飬1-�����ʾ����2-����ֹ
  ---------------------------------------------------------------------------

  v_Output Varchar2(32767);
Begin
  --�������

  For r_Data In (Select Distinct b.����id, Nvl(c.��鷽ʽ, 0) As ��鷽ʽ
                 From ��������˵�� B, ���ϳ����� C
                 Where b.����id = c.�ⷿid(+) And b.������� In (1, 2, 3) And b.�������� = '���ϲ���') Loop
  
    zlJsonPutValue(v_Output, 'warehouse_id', r_Data.����id, 1, 1);
    zlJsonPutValue(v_Output, 'check_type', r_Data.��鷽ʽ, 1, 2);
  
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || v_Output || ']}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Getstockcheck;
/

Create Or Replace Procedure Zl_Stuffsvr_Getstock
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡָ��������ָ���ⷿ�Ŀ��ÿ����
  --��Σ�Json_In:��ʽ
  --  input
  --    stuff_id        N   1   ����ID
  --    warehouse_ids   C   1   �ⷿids,����ö��ŷ���
  --    batch           N       ���Σ�<=0-���������Σ�>0ֻ��ĳ����
  --����: Json_Out,��ʽ����
  --  output
  --    code            N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message         C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    stock           N   1  ���ÿ��
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  n_����     ҩƷ�շ���¼.����%Type;
  n_����id   ҩƷ�շ���¼.ҩƷid%Type;
  n_������� ҩƷ���.�������� %Type;
  v_�ⷿids  Varchar2(4000);
Begin
  --�������
  j_Jsonin  := PLJson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  n_����id  := j_Json.Get_Number('stuff_id');
  v_�ⷿids := j_Json.Get_String('warehouse_ids');
  n_����    := j_Json.Get_Number('batch');

  If Nvl(n_����id, 0) = 0 Then
    Json_Out := zlJsonOut('δ�����������������Ϣ');
    Return;
  End If;

  --��ȡ���(�����������),ҩ��������(����=0,����Ϊҩ��)����Ч��
  If Nvl(n_����, 0) <= 0 Then
    Select Nvl(Sum(a.��������), 0)
    Into n_�������
    From ҩƷ��� A
    Where (Nvl(a.����, 0) = 0 Or a.Ч�� Is Null Or a.Ч�� > Trunc(Sysdate)) And a.���� = 1 And a.ҩƷid = n_����id And
          Instr(',' || v_�ⷿids || ',', ',' || a.�ⷿid || ',') > 0;
  Else
    Select Nvl(Sum(a.��������), 0)
    Into n_�������
    From ҩƷ��� A
    Where Nvl(a.����, 0) = n_���� And (a.Ч�� Is Null Or a.Ч�� > Trunc(Sysdate)) And a.���� = 1 And a.ҩƷid = n_����id And
          (Instr(',' || v_�ⷿids || ',', ',' || a.�ⷿid || ',') > 0 Or
           a.�ⷿid In (Select ����ⷿid
                      From ����ⷿ����
                      Where Instr(',' || v_�ⷿids || ',', ',' || ����id || ',') > 0 And Rownum < 2));
  End If;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","stock":' || n_������� || '}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Getstock;
/

Create Or Replace Procedure Zl_Stuffsvr_Getstockbatch
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ�������ȡ������Ŀ�漰�۸���Ϣ:����Ŀѡ������չʾ��漰�۸���Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --   stuff_ids            C   1   ����ID�������Ӣ�ĵĶ��ŷָ�
  --   warehouse_ids        C   0   �ⷿIDs���ⷿIDsΪNULLʱΪ���пⷿ
  --   return_price         N   1   �Ƿ񷵻��ۼۣ�1-���ؼ۸���Ϣ(�ۼ�);0-������
  --   type                 N   1   0-�����ؿⷿID;1-���ؿⷿID
  --����: Json_Out,��ʽ����
  --  output
  --    code                N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    item_list
  --      stuff_id            N   1   ����ID
  --      warehouse_id        N   1  �ⷿID
  --      stock               N   1   ��������
  --      price               N   1   ���ۼ�(���ؼ۸�ʱ���д���)
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  c_����ids  Clob;
  v_�ⷿids  Varchar2(32767);
  n_���ؼ۸� Number(2);
  n_Type     Number(2);
  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_����ids Collection_Type;
  I           Number;

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;
Begin
  --�������
  j_Jsonin   := PLJson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  c_����ids  := j_Json.Get_Clob('stuff_ids');
  v_�ⷿids  := j_Json.Get_String('warehouse_ids');
  n_���ؼ۸� := Nvl(j_Json.Get_Number('return_price'), 0);
  n_Type     := Nvl(j_Json.Get_Number('type'), 0);

  I := 0;
  While c_����ids Is Not Null Loop
    If Length(c_����ids) <= 4000 Then
      Col_����ids(I) := c_����ids;
      c_����ids := Null;
    Else
      Col_����ids(I) := Substr(c_����ids, 1, Instr(c_����ids, ',', 3980) - 1);
      c_����ids := Substr(c_����ids, Instr(c_����ids, ',', 3980) + 1);
    End If;
    I := I + 1;
  End Loop;

  v_Jtmp := Null;
  If n_���ؼ۸� = 0 Then
    If n_Type = 0 Then
      For I In 0 .. Col_����ids.Count - 1 Loop
        For c_��� In (Select /*+cardinality(b,10)*/
                      a.ҩƷid, Sum(Nvl(a.��������, 0)) As ���
                     From ҩƷ��� A, Table(f_Num2List(Col_����ids(I))) B
                     Where a.ҩƷid = b.Column_Value And a.���� = 1 And
                           (Instr(',' || v_�ⷿids || ',', ',' || a.�ⷿid || ',') > 0 Or v_�ⷿids Is Null) And
                           (Nvl(a.����, 0) = 0 Or a.Ч�� Is Null Or a.Ч�� > Trunc(Sysdate))
                     Group By a.ҩƷid
                     Having Sum(Nvl(a.��������, 0)) <> 0) Loop
        
          v_Jtmp := v_Jtmp || ',{"stuff_id":' || c_���.ҩƷid;
          v_Jtmp := v_Jtmp || ',"stock":' || zlJsonStr(c_���.���, 1);
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
    
    Elsif n_Type = 1 Then
      For I In 0 .. Col_����ids.Count - 1 Loop
        For c_��� In (Select a.ҩƷid, a.�ⷿid, a.���
                     From (Select /*+cardinality(b,10)*/
                             a.ҩƷid, a.�ⷿid, Sum(Nvl(a.��������, 0)) As ���
                            From ҩƷ��� A, Table(f_Num2List(Col_����ids(I))) B
                            Where a.ҩƷid = b.Column_Value And a.���� = 1 And
                                  (Instr(',' || v_�ⷿids || ',', ',' || a.�ⷿid || ',') > 0 Or v_�ⷿids Is Null) And
                                  (Nvl(a.����, 0) = 0 Or a.Ч�� Is Null Or a.Ч�� > Trunc(Sysdate))
                            Group By a.ҩƷid, a.�ⷿid
                            Having Sum(Nvl(a.��������, 0)) <> 0) A) Loop
        
          v_Jtmp := v_Jtmp || ',{"stuff_id":' || c_���.ҩƷid;
          v_Jtmp := v_Jtmp || ',"warehouse_id":' || c_���.�ⷿid;
          v_Jtmp := v_Jtmp || ',"stock":' || zlJsonStr(c_���.���, 1);
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
    End If;
  
    If c_Jtmp Is Null Then
      Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
    Else
      c_Jtmp   := c_Jtmp || v_Jtmp;
      Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || c_Jtmp || ']}}';
    End If;
    Return;
  End If;

  --�����۸�
  v_Jtmp := Null;
  For I In 0 .. Col_����ids.Count - 1 Loop
    For c_��� In (Select a.ҩƷid, Nvl(a.���, 0) As ���, Decode(Nvl(b.�Ƿ���, 0), 1, 0, Nvl(c.�ּ�, 0)) As �۸�
                 From (Select /*+cardinality(b,10)*/
                         a.ҩƷid, Sum(Nvl(a.��������, 0)) As ���
                        From ҩƷ��� A, Table(f_Num2List(Col_����ids(I))) B
                        Where a.ҩƷid = b.Column_Value And a.���� = 1 And
                              (Instr(',' || v_�ⷿids || ',', ',' || a.�ⷿid || ',') > 0 Or v_�ⷿids Is Null) And
                              (Nvl(a.����, 0) = 0 Or a.Ч�� Is Null Or a.Ч�� > Trunc(Sysdate))
                        Group By a.ҩƷid
                        Having Sum(Nvl(a.��������, 0)) <> 0) A, �շ���ĿĿ¼ B, �շѼ�Ŀ C
                 Where a.ҩƷid = c.�շ�ϸĿid And a.ҩƷid = b.Id And c.�۸�ȼ� Is Null And Sysdate Between c.ִ������ And
                       Nvl(c.��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))) Loop
    
      v_Jtmp := v_Jtmp || ',{"stuff_id":' || c_���.ҩƷid;
      v_Jtmp := v_Jtmp || ',"stock":' || zlJsonStr(c_���.���, 1);
      v_Jtmp := v_Jtmp || ',"price":' || zlJsonStr(c_���.�۸�, 1);
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
End Zl_Stuffsvr_Getstockbatch;
/

Create Or Replace Procedure Zl_Stuffsvr_Batchgetprice
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ�������ȡ�����ۼ�(������ʹ��)
  --��Σ�Json_In:��ʽ
  --  input
  --   stuff_ids    C   1   ����IDs�������Ӣ�ĵĶ��ŷָ�
  --����: Json_Out,��ʽ����
  --  output
  --    code          N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message         C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    item_list
  --      stuff_id N   1   ����ID
  --      price   N   1   ���ۼ�(���ؼ۸�ʱ���д���)
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_����ids Collection_Type;
  I           Integer;
  c_����ids   Clob;

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;
Begin
  --�������
  j_Jsonin  := PLJson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  c_����ids := j_Json.Get_Clob('stuff_ids');

  If c_����ids Is Null Then
    Json_Out := zlJsonOut('δ������Ч������id,����!');
  End If;

  I := 0;
  While c_����ids Is Not Null Loop
    If Length(c_����ids) <= 4000 Then
      Col_����ids(I) := c_����ids;
      c_����ids := Null;
    Else
      Col_����ids(I) := Substr(c_����ids, 1, Instr(c_����ids, ',', 3980) - 1);
      c_����ids := Substr(c_����ids, Instr(c_����ids, ',', 3980) + 1);
    End If;
    I := I + 1;
  End Loop;

  v_Jtmp := Null;
  For I In 0 .. Col_����ids.Count - 1 Loop
    --�����۸�
    For c_��� In (With c_ҩƷ��Ϣ As
                    (Select /*+cardinality(D,10)*/
                     d.Column_Value As ҩƷid
                    From Table(f_Num2List(Col_����ids(I))) D)
                   Select a.ҩƷid, Sum(a.ʵ�ʽ��) / Sum(a.ʵ������) As �۸�
                   From ҩƷ��� A, c_ҩƷ��Ϣ B
                   Where a.ҩƷid = b.ҩƷid And a.���� = 1 And (Nvl(a.����, 0) = 0 Or a.Ч�� Is Null Or a.Ч�� > Trunc(Sysdate))
                   Group By a.ҩƷid
                   Having Sum (Nvl(a.ʵ������, 0)) <> 0) Loop
    
      v_Jtmp := v_Jtmp || ',{"stuff_id":' || c_���.ҩƷid;
      v_Jtmp := v_Jtmp || ',"price":' || zlJsonStr(c_���.�۸�, 1);
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
End Zl_Stuffsvr_Batchgetprice;
/

Create Or Replace Procedure Zl_Stuffsvr_Checkexpirydate
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ�����������ĵ����Ч���Ƿ����
  --��Σ�Json_In:��ʽ
  --    input
  --        stuff_id        N   1   ����ID
  --        warehouse_id    N   1   �ⷿID
  --        quantity        N   1   ����
  --        batch           N   1   ���� 
  --����: Json_Out,��ʽ����
  --  output
  --    code                N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message             C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    exist_expiried      N   1   �Ƿ���ڹ���:0-�������ѹ�����Ŀ��1-�����ѹ�����Ŀ
  --    min_expirydate      C       ��С���Ч�ڣ�exist_expiried=1ʱ���أ���ʽ��yyyy-mm-dd
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  n_�ⷿid  ҩƷ�շ���¼.�ⷿid%Type;
  n_����id  ҩƷ�շ���¼.ҩƷid%Type;
  n_����    ҩƷ���.�������� %Type;
  n_����    ҩƷ���.���� %Type;
  d_Mindate Date;
  d_Sysdate Date;
  n_Find    Number(2);
  v_Tmp     Varchar2(20);
Begin
  --�������
  j_Jsonin := PLJson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_����id := j_Json.Get_Number('stuff_id');
  n_�ⷿid := j_Json.Get_Number('warehouse_id');
  n_����   := j_Json.Get_Number('quantity');
  n_����   := j_Json.Get_Number('batch');

  If Nvl(n_����id, 0) = 0 Then
    Json_Out := zlJsonOut('δ�����������������Ϣ');
    Return;
  End If;

  d_Sysdate := Sysdate;
  d_Mindate := To_Date('3000-01-01', 'yyyy-mm-dd');
  n_Find    := 0;
  --��һ���Բ��ϲ��ж�
  -- ��Ϊ���ܸ��������Ч�ڲ�ͬ, ���Ҫ�õ�����������С��Ч��
  For c_���� In (Select c.����, Nvl(b.����, 0) As ����, b.�������� As ���, b.���Ч��
               From �������� A, ҩƷ��� B, �շ���ĿĿ¼ C
               Where a.����id = b.ҩƷid And a.����id = c.Id And a.һ���Բ��� = 1 And b.���� = 1 And Nvl(b.��������, 0) > 0 And
                     a.���Ч�� Is Not Null And a.����id = n_����id And b.�ⷿid = n_�ⷿid And
                     Decode(n_����, Null, -1, b.����) = Nvl(n_����, -1)
               Order By Nvl(b.����, 0)) Loop
    If Nvl(n_Find, 0) = 0 Then
      n_Find := 1;
    End If;
    If c_����.���Ч�� < d_Mindate Then
      d_Mindate := c_����.���Ч��;
    End If;
  
    If Nvl(c_����.���, 0) < n_���� Then
      n_���� := n_���� - Nvl(c_����.���, 0);
    Else
      n_���� := 0;
    End If;
    If n_���� = 0 Then
      Exit;
    End If;
  
  End Loop;

  If d_Sysdate > d_Mindate And Nvl(n_Find, 0) = 1 Then
    v_Tmp  := To_Char(d_Mindate, 'yyyy -mm-dd');
    n_Find := 1;
  Else
    n_Find := 0;
  End If;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","exist_expiried":' || n_Find || ',"min_expirydate":"' || v_Tmp ||
              '"}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Checkexpirydate;
/

CREATE OR REPLACE Procedure Zl_Stuffsvr_Getprice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡָ�����ĵ��ۼۡ��ɱ���
  --��Σ�Json_In:��ʽ
  --  input
  --    stuff_id          N 1 ����ID
  --    warehouse_id      N 1 ҩ��ID
  --    quantity          N 1 ����
  --    batch             N   ���Σ�0-���������Σ�>0ֻ��ĳ����
  --    price_grade       C   �۸�ȼ�����
  --    item_list[]�б�
  --            stuff_id          N 1 ����ID
  --             warehouse_id      N 1 ҩ��ID
  --             quantity          N 1 ����
  --             batch             N   ���Σ�0-���������Σ�>0ֻ��ĳ����
  --             price_grade       C   �۸�ȼ�����
  --����: Json_Out,��ʽ����
  --  output
  --    code              N 1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message           C 1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    price             N 1 �ۼ�
  --    price_cost        N 1 �ɱ���
  --    quantity_remain   N 1 ʣ������������0��ʾ�����㹻������0���ʾ��������
  --    item_list[]�б�
  --            price             N 1 �ۼ�
  --            price_cost        N 1 �ɱ���
  --            quantity_remain   N 1 ʣ������������0��ʾ�����㹻������0���ʾ��������
  ---------------------------------------------------------------------------
  j_Jsonin Pljson;
  j_Json   Pljson;
  j_List   Pljson_List := Pljson_List();
  j_Tmpout Varchar2(32767);

  n_����id   ҩƷ���.ҩƷid%Type;
  n_�ⷿid   ҩƷ���.�ⷿid%Type;
  n_����     ҩƷ���.ʵ������%Type;
  n_����     ҩƷ���.����%Type;
  v_Temp     Varchar2(4000);
  n_����     ҩƷ���.�ɱ���%Type;
  n_�ɱ���   ҩƷ���.�ɱ���%Type;
  n_ʣ������ ҩƷ���.ʵ������%Type;
Begin
  --�������
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  
  If j_Json.Exist('item_list') Then
    j_List := j_Json.Get_Pljson_List('item_list');
    If j_List Is Not Null Then
      For I In 1 .. j_List.Count Loop
        j_Json := Pljson();
        j_Json := Pljson(j_List.Get(I));
      
        n_����id := j_Json.Get_Number('stuff_id');
        n_�ⷿid := j_Json.Get_Number('warehouse_id');
        n_����   := j_Json.Get_Number('quantity');
        n_����   := j_Json.Get_Number('batch');
      
        v_Temp := Zl_Fun_Getprice(n_����id, n_�ⷿid, n_����, 0, n_����);
        --�ֽ�
        n_����     := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1));
        v_Temp     := Substr(v_Temp, Instr(v_Temp, '|') + 1);
        n_�ɱ���   := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1));
        v_Temp     := Substr(v_Temp, Instr(v_Temp, '|') + 1);
        n_ʣ������ := To_Number(v_Temp);
      
        j_Tmpout := j_Tmpout || ',{"price":' || Zljsonstr(n_����, 1);
        j_Tmpout := j_Tmpout || ',"price_cost": ' || Zljsonstr(n_�ɱ���, 1);
        j_Tmpout := j_Tmpout || ',"quantity_remain":' || Zljsonstr(n_ʣ������, 1);
        j_Tmpout := j_Tmpout || '}';
      
      End Loop;
    End If;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || Substr(j_Tmpout, 2) || ']}}';
  Else
  
    n_����id := j_Json.Get_Number('stuff_id');
    n_�ⷿid := j_Json.Get_Number('warehouse_id');
    n_����   := j_Json.Get_Number('quantity');
    n_����   := j_Json.Get_Number('batch');
  
    v_Temp := Zl_Fun_Getprice(n_����id, n_�ⷿid, n_����, 0, n_����);
    --�ֽ�
    n_����     := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1));
    v_Temp     := Substr(v_Temp, Instr(v_Temp, '|') + 1);
    n_�ɱ���   := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1));
    v_Temp     := Substr(v_Temp, Instr(v_Temp, '|') + 1);
    n_ʣ������ := To_Number(v_Temp);
  
    Json_Out := '{"output":{"code":1,"message":"�ɹ�"';
    Json_Out := Json_Out || ',"price":' || Zljsonstr(n_����, 1);
    Json_Out := Json_Out || ',"price_cost": ' || Zljsonstr(n_�ɱ���, 1);
    Json_Out := Json_Out || ',"quantity_remain":' || Zljsonstr(n_ʣ������, 1) || '}}';
  End If;
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Getprice;
/

Create Or Replace Procedure Zl_Stuffsvr_Getidbybarcode
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ�������Ʒ������ڲ������ȡ����ID������
  --��Σ�Json_In:��ʽ
  --  input
  --    barcode             C 1 ��������봮
  --    type                N 0 1-������漰��Ч,0-����漰��Ч
  --    only_barcode_inside N 0 1-�����ڲ�������в���,0-����Ʒ���뼰�ڲ�������в���
  --����: Json_Out,��ʽ����
  --  output
  --    code              N 1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message           C 1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    item_list
  --      stuff_id        N 1   ����ID
  --      batch           N     ���Σ�0-���������Σ�>0ֻ��ĳ����
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  v_����   ҩƷ���.��Ʒ����%Type;
  n_Type   Number(1);
  n_Inside Number(1);

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;
Begin
  --�������
  j_Jsonin := PLJson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  v_����   := j_Json.Get_String('barcode');
  n_Type   := Nvl(j_Json.Get_Number('type'), 0);
  n_Inside := Nvl(j_Json.Get_Number('only_barcode_inside'), 0);

  v_Jtmp := Null;
  If n_Type = 0 Then
    If n_Inside = 1 Then
      For c_���� In (Select Distinct ҩƷid, ����
                   From ҩƷ���
                   Where (�ڲ����� = v_����) And �������� > 0 And (Ч�� Is Null Or Ч�� > Trunc(Sysdate))) Loop
      
        v_Jtmp := v_Jtmp || ',{"stuff_id":' || c_����.ҩƷid;
        v_Jtmp := v_Jtmp || ',"batch":' || Nvl(c_����.����, 0);
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
      For c_���� In (Select Distinct ҩƷid, ����
                   From ҩƷ���
                   Where (��Ʒ���� = v_���� Or �ڲ����� = v_����) And �������� > 0 And (Ч�� Is Null Or Ч�� > Trunc(Sysdate))) Loop
      
        v_Jtmp := v_Jtmp || ',{"stuff_id":' || c_����.ҩƷid;
        v_Jtmp := v_Jtmp || ',"batch":' || Nvl(c_����.����, 0);
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
  Else
    If n_Inside = 1 Then
      For c_���� In (Select Distinct ҩƷid, ����
                   From ҩƷ���
                   Where ��Ʒ���� Like v_���� || '%' Or �ڲ����� = v_���� || '%') Loop
      
        v_Jtmp := v_Jtmp || ',{"stuff_id":' || c_����.ҩƷid;
        v_Jtmp := v_Jtmp || ',"batch":' || Nvl(c_����.����, 0);
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
      For c_���� In (Select Distinct ҩƷid, ���� From ҩƷ��� Where �ڲ����� = v_���� || '%') Loop
      
        v_Jtmp := v_Jtmp || ',{"stuff_id":' || c_����.ҩƷid;
        v_Jtmp := v_Jtmp || ',"batch":' || Nvl(c_����.����, 0);
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
  End If;

  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Getidbybarcode;
/

Create Or Replace Procedure Zl_Stuffsvr_Checkstorelimit
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ָ��������ָ���ⷿ�Ŀ���Ƿ���ڴ�������
  --��Σ�Json_In:��ʽ
  --  input
  --    stuff_id            N   1   ����ID
  --    warehouse_id        N   1   �ⷿID
  --    stock               N   1   �������
  --����: Json_Out,��ʽ����
  --  output
  --    code                N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    below_limit_lower   N 1 1-���ڴ������ޣ�0-���ڵ��ڴ�������

  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  n_����id ҩƷ���.ҩƷid%Type;
  n_�ⷿid ҩƷ���.�ⷿid%Type;
  n_���   ҩƷ���.ʵ������%Type;
  n_Count  Number(1);
Begin
  --�������
  j_Jsonin := PLJson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_����id := j_Json.Get_Number('stuff_id');
  n_�ⷿid := j_Json.Get_Number('warehouse_id');
  n_���   := j_Json.Get_Number('stock');

  --��ȡҩƷ�����޶�
  Select Count(1)
  Into n_Count
  From ���ϴ����޶�
  Where ����id = n_����id And �ⷿid = n_�ⷿid And Nvl(����, 0) <> 0 And ���� > n_��� And Rownum < 2;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","below_limit_lower":' || n_Count || '}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Checkstorelimit;
/


Create Or Replace Procedure Zl_Stuffsvr_Adjustdata
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ��������תסԺʱ�������Ĺ�������
  --��Σ�Json_In:��ʽ
  --  input
  --    pati_id          N  1 ����ID
  --    pati_pageid      N  1 ��ҳID
  --    billtype         N  1 �������ͣ�1-�շѵ�;2-���ʵ�
  --    item_list
  --      stuff_no_old       C  1 ԭ���ݺ�
  --      stuffdtl_id_old    N  1 ԭ������ϸID
  --      stuff_no_new       C  1 �µ��ݺ�
  --      stuffdtl_id_new    N  1 �´�����ϸID
  --����: Json_Out,��ʽ����
  --  output
  --    code              N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message           C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Jsonin   PLJson;
  j_Json     PLJson;
  j_Jsonlist Pljson_List;

  n_����id ҩƷ�շ���¼.����id%Type;
  n_��ҳid ҩƷ�շ���¼.��ҳid%Type;
  n_����   ҩƷ�շ���¼.����%Type;

  v_ԭ���ݺ� δ��ҩƷ��¼.No%Type;
  n_ԭ��ϸid ҩƷ�շ���¼.����id%Type;
  v_�µ��ݺ� ҩƷ�շ���¼.No%Type;
  n_����ϸid ҩƷ�շ���¼.����id%Type;

  n_�������� ��������.��������%Type;

  n_Count Number;
  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin
  --�������
  j_Jsonin   := PLJson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_����id   := j_Json.Get_Number('pati_id');
  n_��ҳid   := j_Json.Get_Number('pati_pageid');
  n_����     := j_Json.Get_Number('billtype');
  j_Jsonlist := j_Json.Get_Pljson_List('item_list');

  If j_Jsonlist Is Not Null Then
    n_Count := j_Jsonlist.Count;
  End If;
  If n_Count = 0 Then
    v_Err_Msg := 'δ����ҩƷ������Ϣ!';
    Raise Err_Item;
  End If;

  For I In 1 .. j_Jsonlist.Count Loop
    j_Json     := PLJson();
    j_Json     := PLJson(j_Jsonlist.Get(I));
    v_ԭ���ݺ� := j_Json.Get_String('stuff_no_old');
    n_ԭ��ϸid := j_Json.Get_Number('stuffdtl_id_old');
    v_�µ��ݺ� := j_Json.Get_String('stuff_no_new');
    n_����ϸid := j_Json.Get_Number('stuffdtl_id_new');
  
    Update δ��ҩƷ��¼
    Set ���� = Decode(����, 24, 25, ����), ��ҳid = n_��ҳid, NO = v_�µ��ݺ�
    Where NO = v_ԭ���ݺ� And ���� = Decode(n_����, 1, 24, 2, 25, 26) And ����id = n_����id;
  
    Update ҩƷ�շ���¼
    Set ���� = Decode(����, 24, 25, ����), ����id = n_����ϸid, NO = v_�µ��ݺ�, ��ҳid = n_��ҳid, ������Դ = 2, ������Դ = 2
    Where NO = v_ԭ���ݺ� And ���� = Decode(n_����, 1, 24, 2, 25, 26) And ����id = n_ԭ��ϸid;
  
    Select Max(��������)
    Into n_��������
    From ��������
    Where ����id In (Select ҩƷid From ҩƷ�շ���¼ Where ���� = 25 And NO = v_�µ��ݺ� And ����id = n_����ϸid);
    If Nvl(n_��������, 0) = 1 Then
      --���±�������
      Update ҩƷ�շ���¼
      Set ����id = n_����ϸid, ��ҳid = n_��ҳid, ������Դ = 2, ������Դ = 2
      Where ���� = 21 And ����id = n_ԭ��ϸid;
    End If;
  End Loop;

  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Stuffsvr_Adjustdata;
/

Create Or Replace Procedure Zl_Stuffsvr_Getbatchnumber
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡ����������Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --    stuff_no            C  1 ���ݺ�
  --    stuffdtl_id         N  1 ������ϸID
  --    billtype          N  1 �������ͣ�1-�շѴ���,2-���ʵ�����,3-���ʱ���
  --����: Json_Out,��ʽ����
  --  output
  --    code              N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message           C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    batch_number      C   1   ����
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  v_���ݺ� ҩƷ�շ���¼.No%Type;
  n_��ϸid ҩƷ�շ���¼.����id%Type;
  n_����   ҩƷ�շ���¼.����%Type;
  v_����   ҩƷ�շ���¼.����%Type;
Begin

  --�������
  j_Jsonin := PLJson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  v_���ݺ� := j_Json.Get_String('stuff_no');
  n_��ϸid := j_Json.Get_Number('stuffdtl_id');
  n_����   := j_Json.Get_Number('billtype');

  Select Max(����)
  Into v_����
  From ҩƷ�շ���¼
  Where ���� = Decode(n_����, 1, 8, 2, 9, 10) And ����id = n_��ϸid And NO = v_���ݺ�;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","batch_number":"' || v_���� || '"}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Getbatchnumber;
/

Create Or Replace Procedure Zl_Stuffsvr_Checkcontainstuff
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���鵥���Ƿ񺬱�������
  --��Σ�Json_In:��ʽ
  --  input
  --    stuffdtl_ids      C   0  ������ϸid��,Ŀǰ����ķ���ID�����ö��ŷָ� ��1,2,3,4
  --����: Json_Out,��ʽ����
  --  output
  --    code              N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message           C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    contain_stuff     N   1   �Ƿ񺬱�������:1-���б�������,0-������������
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  v_Ids   Varchar2(32767);
  n_Count Number(1);
Begin
  --�������
  j_Jsonin := PLJson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  v_Ids    := j_Json.Get_String('stuffdtl_ids');

  Select /*+cardinality(j,10)*/
   Count(1)
  Into n_Count
  From ҩƷ�շ���¼ A, Table(f_Num2List(v_Ids)) J
  Where a.����id = j.Column_Value And a.���� = 21 And Rownum < 2;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","contain_stuff":' || Nvl(n_Count, 0) || '}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Checkcontainstuff;
/

Create Or Replace Procedure Zl_Stuffsvr_Getstuffputinbills
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡ�������ĵĵ�����Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --   query_type         N   1 ��ѯ��ʽ��
  --                                     ��ѯ��ⵥ�ݣ�0-������ⵥ�Ų�ѯ,1-��ѯʱ�������ʱ��,2-���һ�α������п�����ⵥ
  --                                     ��ѯ�����շѡ����˵���3-��ѯ�շѡ����˵�����Ҫ���billTypeʹ��
  --   warehouse_id       N   1 �ⷿID������ⷿ
  --   stuff_no             C   1 ��ⵥ��
  --   begin_time         C   1 ��ⵥ��ѯ��ʼʱ��
  --   end_time           C   1 ��ⵥ��ѯ��ֹʱ��
  --   billType           N     �������� 1-�շѵ�;2-���˵�
  --����: Json_Out,��ʽ����
  --  output
  --    code              N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message           C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    item_list                ����¼
  --      bill_id         N   1 ������ⵥid
  --      bill_no         C   1 ������ⵥ��
  --      serial_number   N     ���
  --      origin          C     ����
  --      provider        C     ��Ӧ��
  --      batch_number    C     ����
  --      production_date C     ��������
  --      expiry_date     C     Ч��
  --      sterilization_expiry_date C     ���Ч��
  --      putin_quantity  N     �������
  --      putin_price     N     ������ۼ�
  --      putin_money     N     ������۽��
  --      audit_date      C     �������
  --      stuff_id        N   1 ����ID
  --      batch           C     ����
  --      barcode_goods   C     ��Ʒ����
  --      barcode_inside  C     �ڲ�����
  --      stock           N     ���ÿ��
  --      stuffdtl_id     N     ������ϸid
  ---------------------------------------------------------------------------
  Cursor c_�����Ϣ Is
    Select a.Id, a.No, a.���, a.����, c.���� As ��Ӧ��, a.����, a.��������, a.Ч��, a.���Ч��, a.ʵ������ As �������, a.���ۼ� As ������ۼ�,
           a.���۽�� As ������۽��, a.�������, a.ҩƷid, a.����, b.��Ʒ����, b.�ڲ�����, b.�������� As ���
    From ҩƷ�շ���¼ A, ҩƷ��� B, ��Ӧ�� C
    Where a.���� = 15 And a.No Is Null And a.�ⷿid Is Null And a.�ⷿid = b.�ⷿid And Nvl(a.����, 0) = Nvl(b.����, 0) And
          a.��ҩ��λid = c.Id And Rownum < 1;
  r_�����Ϣ c_�����Ϣ%RowType;

  Type Ty_�����Ϣ Is Ref Cursor;
  c_�����Ϣ Ty_�����Ϣ; --��̬�α����

  j_Jsonin PLJson;
  j_Json   PLJson;

  n_��ѯ��ʽ Number(1);
  d_��ʼʱ�� Date;
  d_��ֹʱ�� Date;
  v_��ⵥ�� ҩƷ�շ���¼.No%Type;
  n_�ⷿid   ҩƷ���.�ⷿid%Type;
  n_Billtype ҩƷ�շ���¼.����%Type;

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;
Begin
  --�������
  j_Jsonin   := PLJson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_��ѯ��ʽ := j_Json.Get_Number('query_type');
  n_�ⷿid   := j_Json.Get_Number('warehouse_id');
  v_��ⵥ�� := j_Json.Get_String('stuff_no');
  d_��ʼʱ�� := To_Date(j_Json.Get_String('begin_time'), 'yyyy-mm-dd hh24:mi:ss');
  d_��ֹʱ�� := To_Date(j_Json.Get_String('end_time'), 'yyyy-mm-dd hh24:mi:ss');
  n_Billtype := j_Json.Get_Number('billType');
  If n_Billtype = 1 Then
    n_Billtype := 25;
  Else
    n_Billtype := 26;
  End If;

  --��ȡ���������Ϣ
  v_Jtmp := Null;
  If Nvl(n_��ѯ��ʽ, 0) = 3 Then
    For c_������Ϣ In (Select c.����id, c.����, c.��Ʒ����, c.�ڲ�����, Sum(b.��������) As ��������
                   From (Select a.����id, Max(a.ҩƷid) As ҩƷid, Max(a.����) As ����, Max(a.��Ʒ����) As ��Ʒ����, Max(a.�ڲ�����) As �ڲ�����
                          From ҩƷ�շ���¼ A
                          Where a.No = v_��ⵥ�� And a.���� = n_Billtype And Mod(a.��¼״̬, 3) In (0, 1)
                          Group By a.����id) C, ҩƷ��� B
                   Where c.ҩƷid = b.ҩƷid(+) And b.�ⷿid(+) = n_�ⷿid
                   Group By c.����id, c.����, c.��Ʒ����, c.�ڲ�����) Loop
    
      v_Jtmp := v_Jtmp || ',';
      zlJsonPutValue(v_Jtmp, 'stuffdtl_id', c_������Ϣ.����id, 1, 1);
      zlJsonPutValue(v_Jtmp, 'batch', c_������Ϣ.����);
      zlJsonPutValue(v_Jtmp, 'barcode_goods', c_������Ϣ.��Ʒ����);
      zlJsonPutValue(v_Jtmp, 'barcode_inside', c_������Ϣ.�ڲ�����);
      zlJsonPutValue(v_Jtmp, 'stock', c_������Ϣ.��������, 1, 2);
    
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
    If Nvl(n_��ѯ��ʽ, 0) = 0 Then
      --������ⵥ�Ų�ѯ
      Open c_�����Ϣ For
        Select a.Id, a.No, a.���, a.����, c.���� As ��Ӧ��, a.����, To_Char(a.��������, 'yyyy-mm-dd') As ��������,
               To_Char(a.Ч��, 'yyyy-mm-dd') As Ч��, To_Char(a.���Ч��, 'yyyy-mm-dd') As ���Ч��, a.ʵ������ As �������,
               LTrim(To_Char(a.���ۼ�, '9999990.00000')) As ������ۼ�, a.���۽�� As ������۽��, a.�������, a.ҩƷid, a.����, b.��Ʒ����, b.�ڲ�����,
               To_Char(b.��������, '9999990.00000') As ���
        From ҩƷ�շ���¼ A, ҩƷ��� B, ��Ӧ�� C
        Where a.���� = 15 And a.No = v_��ⵥ�� And a.�ⷿid = n_�ⷿid And a.�ⷿid = b.�ⷿid And Nvl(a.����, 0) = Nvl(b.����, 0) And
              a.��ҩ��λid = c.Id;
    Elsif Nvl(n_��ѯ��ʽ, 0) = 1 Then
      --��ѯʱ�������ʱ��
      Open c_�����Ϣ For
        Select a.Id, a.No, a.���, a.����, c.���� As ��Ӧ��, a.����, To_Char(a.��������, 'yyyy-mm-dd') As ��������,
               To_Char(a.Ч��, 'yyyy-mm-dd') As Ч��, To_Char(a.���Ч��, 'yyyy-mm-dd') As ���Ч��, a.ʵ������ As �������,
               LTrim(To_Char(a.���ۼ�, '9999990.00000')) As ������ۼ�, a.���۽�� As ������۽��, a.�������, a.ҩƷid, a.����, b.��Ʒ����, b.�ڲ�����,
               To_Char(b.��������, '9999990.00000') As ���
        From ҩƷ�շ���¼ A, ҩƷ��� B, ��Ӧ�� C
        Where a.���� = 15 And Decode(v_��ⵥ��, Null, '-', a.No) = Decode(v_��ⵥ��, Null, '-', v_��ⵥ��) And a.�ⷿid = n_�ⷿid And
              (a.������� Between d_��ʼʱ�� And d_��ֹʱ��) And a.�ⷿid = b.�ⷿid And Nvl(a.����, 0) = Nvl(b.����, 0) And
              a.��ҩ��λid = c.Id;
    Else
      --���һ�α������п�����ⵥ
      Open c_�����Ϣ For
        Select a.Id, a.No, a.���, a.����, c.���� As ��Ӧ��, a.����, To_Char(a.��������, 'yyyy-mm-dd') As ��������,
               To_Char(a.Ч��, 'yyyy-mm-dd') As Ч��, To_Char(a.���Ч��, 'yyyy-mm-dd') As ���Ч��, a.ʵ������ As �������,
               LTrim(To_Char(a.���ۼ�, '9999990.00000')) As ������ۼ�, a.���۽�� As ������۽��, a.�������, a.ҩƷid, a.����, b.��Ʒ����, b.�ڲ�����,
               To_Char(b.��������, '9999990.00000') As ���
        From ҩƷ�շ���¼ A, ҩƷ��� B, ��Ӧ�� C
        Where a.���� = 15 And a.�ⷿid = n_�ⷿid And a.�ⷿid = b.�ⷿid And Nvl(a.����, 0) = Nvl(b.����, 0) And a.��ҩ��λid = c.Id And
              a.No = (Select Max(NO) As NO
                      From ҩƷ�շ���¼ A1, ҩƷ��� B1
                      Where A1.������� Between Sysdate - 7 And Sysdate And A1.ҩƷid = B1.ҩƷid And A1.�ⷿid = B1.�ⷿid And
                            Nvl(A1.����, 0) = Nvl(B1.����, 0) And Nvl(B1.��������, 0) > 0 And A1.�ⷿid = n_�ⷿid);
    End If;
    Loop
      Fetch c_�����Ϣ
        Into r_�����Ϣ;
      Exit When c_�����Ϣ%NotFound;
    
      v_Jtmp := v_Jtmp || ',';
      zlJsonPutValue(v_Jtmp, 'bill_id', r_�����Ϣ.Id, 1, 1);
      zlJsonPutValue(v_Jtmp, 'bill_no', r_�����Ϣ.No);
      zlJsonPutValue(v_Jtmp, 'serial_number', r_�����Ϣ.���, 1);
      zlJsonPutValue(v_Jtmp, 'origin', r_�����Ϣ.����);
      zlJsonPutValue(v_Jtmp, 'provider', r_�����Ϣ.��Ӧ��);
      zlJsonPutValue(v_Jtmp, 'batch_number', r_�����Ϣ.����);
      zlJsonPutValue(v_Jtmp, 'production_date', r_�����Ϣ.��������);
      zlJsonPutValue(v_Jtmp, 'expiry_date', r_�����Ϣ.Ч��);
      zlJsonPutValue(v_Jtmp, 'sterilization_expiry_date', r_�����Ϣ.���Ч��);
      zlJsonPutValue(v_Jtmp, 'putin_quantity', r_�����Ϣ.�������, 1);
      zlJsonPutValue(v_Jtmp, 'putin_price', r_�����Ϣ.������ۼ�, 1);
      zlJsonPutValue(v_Jtmp, 'putin_money', r_�����Ϣ.������۽��, 1);
      zlJsonPutValue(v_Jtmp, 'audit_date', r_�����Ϣ.�������);
      zlJsonPutValue(v_Jtmp, 'stuff_id', r_�����Ϣ.ҩƷid, 1);
      zlJsonPutValue(v_Jtmp, 'batch', r_�����Ϣ.����);
      zlJsonPutValue(v_Jtmp, 'barcode_goods', r_�����Ϣ.��Ʒ����);
      zlJsonPutValue(v_Jtmp, 'barcode_inside', r_�����Ϣ.�ڲ�����);
      zlJsonPutValue(v_Jtmp, 'stock', r_�����Ϣ.���, 1, 2);
    
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
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Getstuffputinbills;
/

Create Or Replace Procedure Zl_Stuffsvr_Billaffirm
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����Ĵ�������ȷ�ϻ����Ĵ����շ�ȷ��
  --��Σ�Json_In:��ʽ
  --  input
  --    pati_id                   N  0 ����id������շ�ʱ��Ч
  --    pati_name                 C  0 ����������շ�ʱ��Ч
  --    pati_sex                  C  0 �Ա�����շ�ʱ��Ч
  --    pati_age                  C  0 ���䣺����շ�ʱ��Ч
  --    pati_outpno               C  0 ����ţ�����շ�ʱ��Ч
  --    auditor                   C  1 �����
  --    auditor_code              C  1 ����˱��
  --    audit_time                C  1 ���ʱ�䣺yyyy-mm-dd hh24:mi:ss
  --    item_list[]                   ���������б�[����]
  --      billtype                N  1 ��������:1 -�շѴ�����ҩ  ;2- ���ʵ�������ҩ;3- ���ʱ�����ҩ
  --      stuff_no                C  1 ���ݺ�
  --      stuffdtl_ids            C  0 ����ID,���Դ�����,�ö��ŷ���
  --      stuff_auto_send         N  0  �����Զ�����;0-���Զ�����;1-�Զ�����
  --      auto_send_ids           C  0  �Զ����ϵ���ϸids,����ö��ŷָ�
  --����: Json_Out,��ʽ����
  --  output
  --    code                      N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                   C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Jsonin      PLJson;
  j_Json        PLJson;
  j_Jsonlist_In Pljson_List;

  n_����id ҩƷ�շ���¼.����id%Type;
  v_����   ҩƷ�շ���¼.����%Type;
  v_�Ա�   ҩƷ�շ���¼.�Ա�%Type;
  v_����   ҩƷ�շ���¼.����%Type;
  v_����� Number(18);
  --�����������ⵥժҪ����ʽ����������:XXX    �Ա�:XXX    ����XXX    �����:XXX
  v_ժҪ     ҩƷ�շ���¼.ժҪ%Type;
  v_���ݺ�   ҩƷ�շ���¼.No%Type;
  v_��ϸids  Varchar2(4000);
  n_����     ҩƷ�շ���¼.����%Type;
  n_����_In  ҩƷ�շ���¼.����%Type;
  d_���ʱ�� Date;
  n_�Զ����� Number(1);
  v_Err      Varchar2(255);

  v_Nos         Varchar2(32767);
  v_������ϸid  Varchar2(400);
  v_������ϸids Varchar2(4000);
  v_�����      ��Ա��.����%Type;
  v_����˱��  ��Ա��.���%Type;
  Err_Custom Exception;
Begin
  --�������
  j_Jsonin      := PLJson(Json_In);
  j_Json        := j_Jsonin.Get_Pljson('input');
  n_����id      := j_Json.Get_Number('pati_id');
  v_����        := j_Json.Get_String('pati_name');
  v_�Ա�        := j_Json.Get_String('pati_sex');
  v_����        := j_Json.Get_String('pati_age');
  v_�����      := j_Json.Get_String('pati_outpno');
  v_�����      := j_Json.Get_String('auditor');
  v_����˱��  := j_Json.Get_String('auditor_code');
  d_���ʱ��    := To_Date(j_Json.Get_String('audit_time'), 'yyyy-mm-dd hh24:mi:ss');
  j_Jsonlist_In := j_Json.Get_Pljson_List('item_list');

  If Nvl(n_����id, 0) <> 0 Then
    v_ժҪ := '��������:' || v_���� || '    ' || '�Ա�:' || v_�Ա� || '    ' || '����' || v_���� || '    ' || '�����:' || v_�����;
  End If;

  If d_���ʱ�� Is Null Then
    d_���ʱ�� := Sysdate;
  End If;

  If j_Jsonlist_In.Count = 0 Then
    v_Err := 'δ�������ĵ�����Ϣ��';
    Raise Err_Custom;
  End If;

  For I In 1 .. j_Jsonlist_In.Count Loop
    j_Json       := PLJson();
    j_Json       := PLJson(j_Jsonlist_In.Get(I));
    n_����       := j_Json.Get_Number('billtype');
    v_���ݺ�     := j_Json.Get_String('stuff_no');
    v_��ϸids    := j_Json.Get_String('stuffdtl_ids');
    n_�Զ�����   := j_Json.Get_Number('stuff_auto_send');
    v_������ϸid := j_Json.Get_String('auto_send_ids');
  
    n_����_In := n_����;
    If n_���� = 1 Then
      n_���� := 24;
    Elsif n_���� = 2 Then
      n_���� := 25;
    Elsif n_���� = 3 Then
      n_���� := 26;
    Else
      v_Err := '���뵥��������Ч�����飡';
      Raise Err_Custom;
    End If;
  
    If Nvl(v_���ݺ�, '-') = '-' Then
      v_Err := 'δ���봦�����ţ����飡';
      Raise Err_Custom;
    End If;
  
    If Nvl(n_�Զ�����, 0) = 1 Then
      If v_������ϸid Is Null Then
        v_Nos := v_Nos || ',' || v_���ݺ�;
      Else
        v_������ϸids := v_������ϸids || ',' || v_������ϸid;
      End If;
    End If;
  
    If Nvl(v_��ϸids, '-') = '-' Then
      v_Err := 'δ���봦����ϸID�����飡';
      Raise Err_Custom;
    End If;
  
    Zl_�����շ���¼_�������(n_����, v_���ݺ�, v_��ϸids, d_���ʱ��, n_����id, v_����, v_�Ա�, v_����, v_ժҪ);
  End Loop;

  --�������ŷ���
  If Nvl(v_Nos, '-') <> '-' Then
    v_Nos := Substr(v_Nos, 2);
    Zl_�����շ���¼_�Զ�����_s(n_����, v_�����, v_����˱��, v_Nos, 0);
  End If;

  --������ID����
  If Nvl(v_������ϸids, '-') <> '-' Then
    v_������ϸids := Substr(v_������ϸids, 2);
    Zl_�����շ���¼_�Զ�����_s(n_����, v_�����, v_����˱��, v_������ϸids, 1);
  End If;

  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Stuffsvr_Billaffirm;
/

Create Or Replace Procedure Zl_Stuffsvr_Autosendstuff
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ������Զ����ϣ���NO��NO��ϸ��
  --��Σ�Json_In:��ʽ
  --  input
  --    billtype             N 1 ��������: 1-�շѴ������ϣ�2-���ʵ��������ϣ�3-���ʱ�������
  --    operator_name        C 1 ����Ա����
  --    operator_code        C 1 ����Ա���
  --    stuff_nos            C 1 ���ݺŴ���NO1,NO2...
  --    stuffdtl_ids         C 1 ������ϸid��,Ŀǰ����ķ���ID�����ö��ŷָ� ��1,2,3,4
  --    send_type            N 1 ��ҩ����,0-�� �����no�������ͷ�ҩ,1-ֻ�� ������ϸid����ҩ
  --����: Json_Out,��ʽ����
  --  output
  --    code                 N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message              C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
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
      n_���� := 24;
    Elsif n_���� = 2 Then
      n_���� := 25;
    Elsif n_���� = 3 Then
      n_���� := 26;
    Else
      v_Err := '����ڵ㡾billtype���������飡';
      Raise Err_Custom;
    End If;
    v_Nos := j_Json.Get_String('stuff_nos');
    If j_Json.Exist('stuffdtl_ids') Then
      v_Ids := j_Json.Get_String('stuffdtl_ids');
    End If;
    If v_Ids Is Null And v_Nos Is Null Then
      v_Err := 'δ�������ĵ��ݡ�rcp_nos���ڵ����ϸ��Ϣ��stuffdtl_ids���ڵ�';
      Raise Err_Custom;
      Return;
    End If;
  End If;

  If Nvl(n_Send_Type, 0) = 1 Then
    If j_Json.Exist('stuffdtl_ids') Then
      v_Ids := j_Json.Get_String('stuffdtl_ids');
    End If;
    If v_Ids Is Null Then
      v_Err := 'δ����������ϸ��Ϣ��stuffdtl_ids���ڵ�';
      Raise Err_Custom;
      Return;
    End If;
  End If;

  v_����Ա���� := j_Json.Get_String('operator_name');
  v_����Ա��� := j_Json.Get_String('operator_code');

  --�����ݺŷ���
  If v_Nos Is Not Null Then
    Zl_�����շ���¼_�Զ�����_s(n_����, v_����Ա����, v_����Ա���, v_Nos, 0);
  End If;

  --������ID����
  If v_Ids Is Not Null Then
    Zl_�����շ���¼_�Զ�����_s(n_����, v_����Ա����, v_����Ա���, v_Ids, 1);
  End If;

  Json_Out := '{"output":{"code":1,"message": "�ɹ�"}}';
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Stuffsvr_Autosendstuff;
/


Create Or Replace Procedure Zl_Stuffsvr_Autoreturnstuff
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ��Զ����ϣ���������ϸ������ID���ϣ�Ĭ����ȫ�ˣ�
  --��Σ�Json_In:��ʽ
  --  input
  --    audit_operator        C 1 �����
  --    stuffdtl_ids           ������ϸid,Ŀǰ����ķ���ID,������������(ð�żӶ������)������Ϊ�ձ�ʾȫ������ ������id1:����1,����id2:����2...
  --����: Json_Out,��ʽ����
  --  output
  --    code                 N 1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message              C 1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  d_����ʱ��   ҩƷ�շ���¼.�������%Type;
  v_����Ա���� ��Ա��.����%Type := Null;
  v_Err        Varchar2(255);
  Err_Custom Exception;
  v_Ids      Clob;
  n_����id   Number;
  n_����     ҩƷ�շ���¼.ʵ������%Type;
  n_��ҩ���� ҩƷ�շ���¼.ʵ������%Type;
  v_Tmp      Clob;
  v_Field    Varchar2(32767);
Begin
  --�������  
  j_Jsonin     := PLJson(Json_In);
  j_Json       := j_Jsonin.Get_Pljson('input');
  v_����Ա���� := j_Json.Get_String('audit_operator');

  If j_Json.Exist('stuffdtl_ids') Then
    v_Ids := j_Json.Get_Clob('stuffdtl_ids');
  End If;

  If v_Ids Is Null Then
    v_Err := 'δ������ϸ��Ϣ��stuffdtl_ids���ڵ�';
    Raise Err_Custom;
  End If;

  d_����ʱ�� := Sysdate;

  v_Tmp := v_Ids || ',';
  While Length(v_Tmp) <> 0 Loop
    v_Field  := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    n_����id := To_Number(Substr(v_Field, 1, Instr(v_Field, ':') - 1));
    n_����   := Substr(v_Field, Instr(v_Field, ':') + 1);
    v_Tmp    := Replace(',' || v_Tmp, ',' || v_Field || ',');
  
    If n_���� Is Not Null Then
      n_��ҩ���� := n_����;
    End If;
  
    --�ֽ���ҩ����
    For r_������ϸ In (Select a.�ⷿid, a.Id, Nvl(a.����, 1) * a.ʵ������ As ����
                   From ҩƷ�շ���¼ A
                   Where a.���� In (24, 25, 26) And (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0) And ����� Is Not Null And
                         a.����id = n_����id
                   Order By a.�ⷿid, a.ҩƷid, a.����) Loop
    
      If n_���� Is Null Then
        --���������Ϊ�ձ�ʾȫ��
      
        --�������ϣ�������ϸ��
        Zl_�����շ���¼_��������_s(r_������ϸ.Id, v_����Ա����, d_����ʱ��, Null, Null, Null, Null, 0, v_����Ա����);
      Else
        If n_��ҩ���� > 0 Then
          If n_��ҩ���� > r_������ϸ.���� Then
            --�������ϣ�������ϸ��
            Zl_�����շ���¼_��������_s(r_������ϸ.Id, v_����Ա����, d_����ʱ��, Null, Null, Null, r_������ϸ.����, 0, v_����Ա����);
          
            n_��ҩ���� := n_��ҩ���� - r_������ϸ.����;
          Else
            --�������ϣ�������ϸ��
            Zl_�����շ���¼_��������_s(r_������ϸ.Id, v_����Ա����, d_����ʱ��, Null, Null, Null, n_��ҩ����, 0, v_����Ա����);
          
            n_��ҩ���� := 0;
          End If;
        End If;
      End If;
    End Loop;
  End Loop;

  Json_Out := '{"output":{"code":1,"message": "�ɹ�"}}';

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Stuffsvr_Autoreturnstuff;
/

Create Or Replace Procedure Zl_Stuffsvr_Newbill
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

  ---------------------------billtype = 3,���ʱ�����ҩƷ�շ���¼.����=26��ʱ�������½ڵ�--------------------------------------
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
  ---------------------------billtype = 3,���ʱ�����ҩƷ�շ���¼.����=26��ʱ�������Ͻڵ�-----------------------------------------

  --     bill_list[]                      ���������б�[����]
  --        stuff_no                  C  1 NO
  --        charge_tag                N  1 �շѱ�־:0-δ�շѻ���ʻ���;1-���շѻ����
  --        fee_acnter                C    ������
  --        plcdept_id                C    ��������id��������)
  --        plcdept                   C    �����������ƣ�������)
  --        placer_id                 C    ����ҽʦid��������)
  --        placer                    C    ����ҽʦ��������)  ����
  --        apply_fee_category_code   C    ���뵥�ѱ����(ҽ�Ƹ��ʽ����)(������) ���ӣ�
  --        apply_fee_category_name   C    ���뵥�ѱ����ƣ�ҽ�Ƹ��ʽ���ƣ�(������) ���ӣ�
  --        operator_name             C  1 ����Ա����
  --        operator_code             C  1 ����Ա���
  --        create_time               C  1 �Ǽ�ʱ��:yyyy-mm-dd hh:mi:ss
  --        item_list[]                    ���������б�[����]

  ---------------------------billtype = 3,���ʱ�����ҩƷ�շ���¼.����=26��ʱ�������½ڵ�----------------------------------------
  --           pati_id                 N  1 ����ID
  --           pati_pageid             N    ��ҳID
  --           pati_name               C  1 ��������
  --           pati_sex                C  1 �Ա�
  --           pati_age                C  1 ����
  --           pati_identity           C    ���
  --           pati_birthdate          C    ��������:yyyy-mm-dd hh:mi:ss
  --           pati_idcard             C    ���֤��
  --           pati_wardarea_id        N    ���˲���ID
  --           pati_deptid             N  1 ���˿���ID
  ---------------------------billtype = 3,���ʱ�����ҩƷ�շ���¼.����=26��ʱ�������Ͻڵ�-----------------------------------------
  --           stuffdtl_id             N  1 ������ϸID
  --           serial_num              N  1 ���
  --           warehouse_id            N  1 �ⷿID
  --           is_bakstuff             N  1 �Ƿ񱸻�����:�и�ֵ���Ĳ���Ҫ���룬��0��ʾ�Ǹ�ֵ����ģʽ(��ɨ��ʱʹ��)
  --           bakstuff_batch             1 ������������
  --           stuff_id                N  1 ����ID
  --           baby_num                N    Ӥ�����

  ---------------------------���½ڵ�Ϊ��ѡ������ҽ����¼����-----------------------------------------------
  --           advice_id               N  0 ҽ��ID
  --           emergency_tag           N    ҽ����¼�еĽ�����־(0-��ͨ;1-����;2-��¼(��������Ч))
  --           effectivetime           N  0 ҽ����Ч
  --           freq_name               C  0 Ƶ������
  --           single                  N  0 ����
  ---------------------------���Ͻڵ�Ϊ��ѡ������ҽ����¼����-----------------------------------------------

  --           packages_num            N  1 ����
  --           outbound_num            N  1 ��������
  --           price                   N    �ۼ�
  --           warehouse_window        C  0 ���ϴ���
  --           memo                    C  0 ժҪ
  --           fee_source              N  0 ������Դ
  --           stuff_auto_send         N  0 �����Զ�����;0-���Զ�����;1-�Զ�����

  --����: Json_Out,��ʽ����
  --  output
  --    code                 N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message              C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
Begin
  --ֱ�ӵ�������ҵ����̣���θ�ʽһ�£�
  Zl_ҩƷ�շ���¼_Newstuffbill(Json_In, Json_Out);

  Json_Out := Zljsonout('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Stuffsvr_Newbill;
/

Create Or Replace Procedure Zl_Stuffsvr_Checkexistsbarcode
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ָ��������ҩƷ������Ƿ������Ʒ������ڲ�����
  --��Σ�Json_In:��ʽ
  --  input
  --    stuff_ids           C 1 ���������ids,����ö��ŷָ�
  --    barcode             C 0 ��ǰ��ѯ�����봮
  --    only_barcode_inside N 0 1-�����ڲ�������в���,0-����Ʒ���뼰�ڲ�������в���
  --����: Json_Out,��ʽ����
  --  output
  --    code                N 1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C 1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    stuff_ids           C 1  ���ص�ҩƷ����д�����Ʒ������ڲ����������ids,����ö��ŷָ�
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  v_����ids Varchar2(4000);
  v_����    ҩƷ���.��Ʒ����%Type;
  n_Inside  Number(1);
  v_Temp    Varchar2(4000);
Begin
  --�������
  j_Jsonin  := PLJson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  v_����ids := j_Json.Get_String('stuff_ids');
  v_����    := j_Json.Get_String('barcode');
  n_Inside  := Nvl(j_Json.Get_Number('only_barcode_inside'), 0);

  v_Temp := Null;
  If v_���� Is Null Then
    For c_���� In (Select /*+cardinality(b,10) */
                 Distinct ҩƷid
                 From ҩƷ��� A, Table(f_Str2List(v_����ids)) B
                 Where (Nvl(n_Inside, 0) = 0 And a.��Ʒ���� Is Not Null Or a.�ڲ����� Is Not Null) And a.�������� > 0 And
                       (a.Ч�� Is Null Or a.Ч�� > Trunc(Sysdate)) And a.ҩƷid = b.Column_Value And Not Exists
                  (Select 1
                        From ҩƷ���
                        Where ҩƷid = a.ҩƷid And (Nvl(n_Inside, 0) = 1 Or ��Ʒ���� Is Null) And �ڲ����� Is Null)) Loop
      v_Temp := v_Temp || ',' || c_����.ҩƷid;
    End Loop;
  Else
    For c_���� In (Select /*+cardinality(b,10) */
                 Distinct ҩƷid
                 From ҩƷ��� A, Table(f_Str2List(v_����ids)) B
                 Where (Nvl(n_Inside, 0) = 0 And a.��Ʒ���� Is Not Null Or a.�ڲ����� Is Not Null) And a.�������� > 0 And
                       (a.Ч�� Is Null Or a.Ч�� > Trunc(Sysdate)) And a.ҩƷid = b.Column_Value And Not Exists
                  (Select 1
                        From ҩƷ���
                        Where ҩƷid = a.ҩƷid And (Nvl(n_Inside, 0) = 0 And Nvl(��Ʒ����, '-') = v_���� Or Nvl(�ڲ�����, '-') = v_����)) And
                       Not Exists
                  (Select 1
                        From ҩƷ���
                        Where ҩƷid = a.ҩƷid And (Nvl(n_Inside, 0) = 1 Or ��Ʒ���� Is Null) And �ڲ����� Is Null)) Loop
      v_Temp := v_Temp || ',' || c_����.ҩƷid;
    End Loop;
  End If;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","stuff_ids":"' || Substr(v_Temp, 2) || '"}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Checkexistsbarcode;
/

Create Or Replace Procedure Zl_Stuffsvr_Delbill
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����Ĵ�������(���˷�)
  --��Σ�Json_In:��ʽ
  -- input
  --     billtype                 N   1   ��������:1 -�շѴ�������  ;2- ���ʵ���������
  --     stuff_no                 C   1   ���ݺ�,�иýڵ��ǰ�����NO�������ʣ���ʱֻ������������ϵͳ�ӿڣ�
  --     item_list[]                ���������б�[����]
  --          stuffdtl_id         N 1 ������ϸid,Ŀǰ����ķ���ID
  --          return_num          N 1 ��������

  --     return_list[]�Զ���ҩ�б�
  --           audit_operator        C 1 �����
  --           operator_time         C 1 ����ʱ��
  --           stuffdtl_ids          C 1 ������ϸid,Ŀǰ����ķ���ID,������������(ð�żӶ������)������Ϊ�ձ�ʾȫ������ ������id1:����1,����id2:����2...

  --����: Json_Out,��ʽ����
  -- output
  --    code                 N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message              C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Json     Pljson;
  o_Json     Pljson;
  j_Jsonlist Pljson_List;
  j_Item     Pljson;

  n_��������   ҩƷ�շ���¼.ʵ������%Type;
  n_������ϸid ҩƷ�շ���¼.����id%Type;
  d_����ʱ��   ҩƷ�շ���¼.�������%Type;
  v_����Ա���� ��Ա��.����%Type;
  v_Ids        Clob;
  n_����id     Number;
  n_����       ҩƷ�շ���¼.ʵ������%Type;
  n_����       Number(1);
  v_No         ҩƷ�շ���¼.No%Type;
  n_��������   ҩƷ�շ���¼.ʵ������%Type;

  v_Tmp   Clob;
  v_Field Varchar2(32767);
  v_Err   Varchar2(255);
  Err_Custom Exception;
Begin
  --�������
  j_Json := Pljson(Json_In);
  o_Json := j_Json.Get_Pljson('input');

  --��NO���Ϻ����ˣ�Ŀǰ����������ӿ�
  v_No := o_Json.Get_String('stuff_no');
  If v_No Is Not Null Then
    n_���� := o_Json.Get_Number('billtype');
  
    For r_������ϸ In (Select ����id, Sum(Nvl(����, 1) * ʵ������) As ��������
                   From ҩƷ�շ���¼
                   Where ���� = Decode(n_����, 1, 24, 2, 25, 26) And NO = v_No And ������� Is Null
                   Group By ����id
                   Order By ����id) Loop
      --������id��������
      Zl_�����շ���¼_�����˷�_s(r_������ϸ.����id, r_������ϸ.��������, 1);
    End Loop;
  End If;

  --�Զ�����
  j_Jsonlist := Pljson_List();
  j_Jsonlist := o_Json.Get_Pljson_List('return_list');
  If j_Jsonlist Is Not Null Then
    For I In 1 .. j_Jsonlist.Count Loop
      j_Item       := Pljson();
      j_Item       := Pljson(j_Jsonlist.Get(I));
      v_����Ա���� := j_Item.Get_String('audit_operator');
      d_����ʱ��   := To_Date(j_Item.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');
      v_Ids        := j_Item.Get_Clob('stuffdtl_ids');
      v_Tmp        := v_Ids || ',';
      While Length(v_Tmp) <> 0 Loop
        v_Field  := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
        n_����id := To_Number(Substr(v_Field, 1, Instr(v_Field, ':') - 1));
        n_����   := Substr(v_Field, Instr(v_Field, ':') + 1);
        v_Tmp    := Replace(',' || v_Tmp, ',' || v_Field || ',');
      
        If n_���� Is Not Null Then
          n_�������� := n_����;
        End If;
      
        --�ֽ���������
        For r_������ϸ In (Select a.�ⷿid, a.Id, Nvl(a.����, 1) * a.ʵ������ As ����
                       From ҩƷ�շ���¼ A
                       Where a.���� In (24, 25, 26) And (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0) And ����� Is Not Null And
                             a.����id = n_����id
                       Order By a.�ⷿid, a.ҩƷid, a.����) Loop
        
          If n_���� Is Null Then
            --���������Ϊ�ձ�ʾȫ��
            --�������ϣ�������ϸ��
            Zl_�����շ���¼_��������_s(r_������ϸ.Id, v_����Ա����, d_����ʱ��, Null, Null, Null, Null, 0, v_����Ա����);
          Else
            If n_�������� > 0 Then
              If n_�������� > r_������ϸ.���� Then
                --�������ϣ�������ϸ��
                Zl_�����շ���¼_��������_s(r_������ϸ.Id, v_����Ա����, d_����ʱ��, Null, Null, Null, r_������ϸ.����, 0, v_����Ա����);
              
                n_�������� := n_�������� - r_������ϸ.����;
              Else
                --�������ϣ�������ϸ��
                Zl_�����շ���¼_��������_s(r_������ϸ.Id, v_����Ա����, d_����ʱ��, Null, Null, Null, n_��������, 0, v_����Ա����);
              
                n_�������� := 0;
              End If;
            End If;
          End If;
        End Loop;
      End Loop;
    End Loop;
  End If;

  --ɾ���ĵ���
  n_������ϸid := Null;
  j_Jsonlist   := Pljson_List();
  j_Jsonlist   := o_Json.Get_Pljson_List('item_list');
  If j_Jsonlist Is Not Null Then
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json       := Pljson();
      o_Json       := Pljson(j_Jsonlist.Get(I));
      n_��������   := o_Json.Get_Number('return_num');
      n_������ϸid := o_Json.Get_Number('stuffdtl_id');
    
      If n_������ϸid Is Null Then
        v_Err := '����ڵ㡾stuffdtl_id���������飡';
        Raise Err_Custom;
      End If;
    
      If n_�������� Is Null Then
        v_Err := '����ڵ㡾return_num���������飡';
        Raise Err_Custom;
      End If;
      Zl_�����շ���¼_�����˷�_s(n_������ϸid, n_��������, 1);
    End Loop;
  End If;
  Json_Out := Zljsonout('�ɹ�', 1);
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Stuffsvr_Delbill;
/


Create Or Replace Procedure Zl_Stuffsvr_Getbakstockbatch
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ�������ȡ������Ŀ�漰�۸���Ϣ:����Ŀѡ������չʾ��漰�۸���Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --   stuff_ids            C   1   ����ID�������Ӣ�ĵĶ��ŷָ�
  --   warehouse_id         C   0   �ⷿID
  --����: Json_Out,��ʽ����
  --  output
  --    code                N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    item_list
  --      stuff_id            N   1   ����ID
  --      stock               N   1   ��������
  --      batch               N   0   ����
  --      batch_number        C   0   ����
  --      barcode_goods       C   0   ��Ʒ����
  --      barcode_inside      C   0   �ڲ�����
  --      provider            C   0   ��Ӧ��
  ---------------------------------------------------------------------------
  j_Jsonin  PLJson;
  j_Json    PLJson;
  c_����ids Clob;
  n_�ⷿid  ҩƷ���.�ⷿid%Type;
  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_����ids Collection_Type;
  I           Number;

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;
Begin
  --�������
  j_Jsonin  := PLJson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  c_����ids := j_Json.Get_Clob('stuff_ids');
  n_�ⷿid  := j_Json.Get_String('warehouse_id');

  I := 0;
  While c_����ids Is Not Null Loop
    If Length(c_����ids) <= 4000 Then
      Col_����ids(I) := c_����ids;
      c_����ids := Null;
    Else
      Col_����ids(I) := Substr(c_����ids, 1, Instr(c_����ids, ',', 3980) - 1);
      c_����ids := Substr(c_����ids, Instr(c_����ids, ',', 3980) + 1);
    End If;
    I := I + 1;
  End Loop;

  For I In 0 .. Col_����ids.Count - 1 Loop
    For c_��� In (Select /*+cardinality(b,10)*/
                  a.ҩƷid, Sum(Nvl(a.��������, 0)) As ���, a.����, a.�ϴ����� As ����, a.��Ʒ����, a.�ڲ�����, Max(c.����) As ��Ӧ��
                 From ҩƷ��� A, Table(f_Num2List(Col_����ids(I))) B, ��Ӧ�� C
                 Where a.ҩƷid = b.Column_Value And a.���� = 1 And a.�ϴι�Ӧ��id = c.Id(+) And a.�ⷿid = n_�ⷿid And
                       (a.Ч�� Is Null Or a.Ч�� > Trunc(Sysdate))
                 Group By a.ҩƷid, a.�ⷿid, a.����, a.�ϴ�����, a.��Ʒ����, a.�ڲ�����
                 Having Sum(Nvl(a.��������, 0)) > 0
                 Order By a.ҩƷid, a.����) Loop
    
      v_Jtmp := v_Jtmp || ',';
      zlJsonPutValue(v_Jtmp, 'stuff_id', c_���.ҩƷid, 1, 1);
      zlJsonPutValue(v_Jtmp, 'stock', c_���.���, 1);
      zlJsonPutValue(v_Jtmp, 'batch', c_���.����, 1);
      zlJsonPutValue(v_Jtmp, 'batch_number', c_���.����);
      zlJsonPutValue(v_Jtmp, 'barcode_goods', c_���.��Ʒ����);
      zlJsonPutValue(v_Jtmp, 'barcode_inside', c_���.�ڲ�����);
      zlJsonPutValue(v_Jtmp, 'provider', c_���.��Ӧ��, 0, 2);
    
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
End Zl_Stuffsvr_Getbakstockbatch;
/

Create Or Replace Procedure Zl_Stuffsvr_Getbill
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ�ִ��ҽ�������ջ�ʱ���漰�ļ��Ͳ�ѯͨ������id�������Ϣ
  --���      json
  --input     
  --  fee_ids                                    C 1 ����idƴ�������ŷָ� 
  --����      json
  --output      
  --  code                                       C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message                                    C 1 Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  item_list                                     ʱ����ϸ��Ϣ��֧�ֶ����[����]
  --    fee_id                                   N 1 ����id
  --    order_total_qunt                         N 1 ����
  --    drug_reocrd_id                           N 1 ҩƷ�շ�id
  --    billtype                                 N 1 ����
  --    drug_id                                  N 1 ҩƷid
  --    obj_dept_id                              N 1 �Է�����id
  --    warehouse_id                             N 1 �ⷿid
  --    batch                                    N 1 ����
  --    batch_number                             C 1 ����
  --    expiry_date                              C 1 ��Ч
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  v_����ids Varchar(4000);

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;
Begin
  --�������
  j_Jsonin  := PLJson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  v_����ids := j_Json.Get_String('fee_ids');

  v_Jtmp := Null;
  For R In (Select Nvl(b.����, 1) * b.ʵ������ As ����, b.Id As �շ�id, b.����, b.ҩƷid, b.�Է�����id, b.�ⷿid, b.����id, b.����, b.����,
                   To_Char(b.Ч��, 'yyyy-mm-dd hh24:mi:ss') As Ч��
            From ҩƷ�շ���¼ B
            Where b.����id In (Select /*+cardinality(x,10)*/
                              x.Column_Value
                             From Table(Cast(f_Num2List(v_����ids) As Zltools.t_Numlist)) X) And
                  (b.��¼״̬ = 1 Or Mod(b.��¼״̬, 3) = 0)
            Order By b.����id) Loop
  
    v_Jtmp := v_Jtmp || ',';
    zlJsonPutValue(v_Jtmp, 'fee_id', r.����id, 1, 1);
    zlJsonPutValue(v_Jtmp, 'order_total_qunt', r.����, 1);
    zlJsonPutValue(v_Jtmp, 'drug_reocrd_id', r.�շ�id, 1);
    zlJsonPutValue(v_Jtmp, 'billtype', r.����, 1);
    zlJsonPutValue(v_Jtmp, 'drug_id', r.ҩƷid, 1);
  
    zlJsonPutValue(v_Jtmp, 'obj_dept_id', r.�Է�����id, 1);
    zlJsonPutValue(v_Jtmp, 'warehouse_id', r.�ⷿid, 1);
    zlJsonPutValue(v_Jtmp, 'batch', r.����, 1);
    zlJsonPutValue(v_Jtmp, 'batch_number', r.����, 0);
    zlJsonPutValue(v_Jtmp, 'expiry_date', r.Ч��, 1, 2);
  
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
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Getbill;
/

Create Or Replace Procedure Zl_Stuffsvr_Getvirwarehouse
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡ����ⷿ
  --��Σ�Json_In:��ʽ
  --  input
  --����: Json_Out,��ʽ����
  --  output
  --    code                        N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                     C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    item_list     [����]
  --        dept_id                 N   1   ����ID
  --        warehouse_id            N   1   �ⷿID
  --        vir_warehouse_id    N   1   ����ⷿID
  ---------------------------------------------------------------------------

  v_List Varchar2(32767);
Begin
  --�������
  For r_Data In (Select Distinct ����id, �ⷿid, ����ⷿid From ����ⷿ����) Loop
    zlJsonPutValue(v_List, 'dept_id', r_Data.����id, 1, 1);
    zlJsonPutValue(v_List, 'warehouse_id', r_Data.�ⷿid, 1);
    zlJsonPutValue(v_List, 'vir_warehouse_id', r_Data.����ⷿid, 1, 2);
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || v_List || ']}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Getvirwarehouse;
/

Create Or Replace Procedure Zl_Stuffsvr_Getnotsendrec
(
  Json_In  Varchar,
  Json_Out Out Clob
) Is
  ---------------------------------------------------------------------------
  --���ܣ���ȡδ���ϼ�¼
  --��Σ�JSON��ʽ
  --input
  --  billtypes             C  1 �������ͣ������Ӣ�Ķ��ŷָ�:  1-�շѷ��ϵ�;2-���ʷ��ϵ�;3-���ʱ��ϵ�
  --  charge_tag            N  1 �շѱ�־:0-δ�շѻ���ʻ���;1-���շѻ����
  --  fee_source            C  1 ������Դ�������Ӣ�Ķ��ŷָ�:1-����,2-סԺ,4-���
  --  start_time            C  0 ��ʼʱ��:yyyy-mm-dd hh:mi:ss
  --  end_time              C  0 ����ʱ��:yyyy-mm-dd hh:mi:ss
  --���Σ�JSON��ʽ
  --output
  --  code  N  1  Ӧ����0-ʧ�ܣ�1-�ɹ�
  --  message C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  item_list[]
  --    billtype              N  1 ��������: 1-�շѷ��ϵ�;2-���ʷ��ϵ�;3-���ʱ��ϵ�
  --    stuff_no              C  1 ���ϵ���
  --    warehouse_id          N  1 �ⷿID
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  v_�������� Varchar2(100);
  n_�շѱ�־ Number(1);
  v_������Դ Varchar2(100);
  d_��ʼʱ�� Date;
  d_����ʱ�� Date;

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;
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
  v_�������� := Replace(v_��������, ',1,', ',24,');
  v_�������� := Replace(v_��������, ',2,', ',25,');
  v_�������� := Replace(v_��������, ',3,', ',26,');

  v_Jtmp := Null;
  For r_���� In (Select Decode(b.����, 24, 1, 25, 2, 26, 3) As ��������, b.No, b.�ⷿid
               From δ��ҩƷ��¼ B
               Where Instr(v_��������, ',' || b.���� || ',') > 0 And Nvl(b.���շ�, 0) = n_�շѱ�־ And
                     b.�������� Between Nvl(d_��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And Nvl(d_����ʱ��, Sysdate) And Exists
                (Select 1
                      From ҩƷ�շ���¼
                      Where ���� = b.���� And NO = b.No And
                            (Instr(',' || v_������Դ || ',', ',' || ������Դ || ',') > 0 Or ������Դ Is Null))) Loop
  
    v_Jtmp := v_Jtmp || ',{"billtype":' || r_����.��������;
    v_Jtmp := v_Jtmp || ',"stuff_no":"' || r_����.No || '"';
    v_Jtmp := v_Jtmp || ',"warehouse_id":' || Nvl(r_����.�ⷿid, 0);
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
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Getnotsendrec;
/
CREATE OR REPLACE Procedure Zl_Stuffsvr_Odr_Check
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ڷ����ջ��������ȡ���ͼ��
  --��Σ�Json_In:��ʽ
  --  input
  --     item_list[]�б�
  --               order_id                        N 1 ҽ��ID
  --               stuff_nos                         C 1 ���ݺ�ƴ��
  --����: Json_Out,��ʽ����
  --  output
  --    code                          N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Input    Pljson;
  j_Item     Pljson;
  j_List     Pljson_List := Pljson_List();
  n_ҽ��id   Number(18);
  n_Count    Number;
  v_Nos      Varchar2(30000);
  v_List     Varchar2(32767);
  v_Json_Out Varchar2(32767);
  v_Error    Varchar2(255);
  Err_Custom Exception;
Begin
  j_Item  := Pljson(Json_In);
  j_Input := j_Item.Get_Pljson('input');
  j_List  := j_Input.Get_Pljson_List('item_list');
  If j_List Is Not Null Then
    For I In 1 .. j_List.Count Loop
      j_Item   := Pljson();
      j_Item   := Pljson(j_List.Get(I));
      n_ҽ��id := j_Item.Get_Number('order_id');
      v_Nos    := j_Item.Get_String('stuff_nos');
    
      For r_���� In (Select /*+cardinality(j,10)*/
                    a.No, a.����id, a.ҩƷid, Sum(Nvl(a.����, 1) * Decode(a.�����, Null, 0, a.ʵ������)) As �ѷ�����
                   From ҩƷ�շ���¼ A, Table(f_Str2list(v_Nos)) J
                   Where a.No = j.Column_Value And a.ҽ��id = n_ҽ��id
                   Group By a.No, a.����id, a.ҩƷid) Loop
      
        v_List := v_List || ',{"stuffdtl_id":' || r_����.����id;
        v_List := v_List || ',"sended_num":' || Zljsonstr(r_����.�ѷ�����, 1);
        v_List := v_List || ',"order_id":' || n_ҽ��id;
        v_List := v_List || ',"stuff_id":' || r_����.ҩƷid;
        v_List := v_List || '}';
      
      End Loop;
    
    End Loop;
  
    If v_List Is Not Null Then
      v_List := ',"item_list":[' || Substr(v_List, 2) || ']';
    End If;
  
    v_Json_Out := '{"code":1,"message":"�ɹ�"';
    v_Json_Out := v_Json_Out || v_List;
    v_Json_Out := v_Json_Out || '}';
    Json_Out   := '{"output":' || v_Json_Out || '}';
  
  End If;
Exception
  When Err_Custom Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(v_Error) || '"}}';
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Odr_Check;
/
Create Or Replace Procedure Zl_Stuffsvr_Overdue_Recovery
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ�ҽ�����ڷ����ջ�������ش���
  --��Σ�Json_In:��ʽ
  --  input
  --     operator_name                      C 1 ����Ա����
  --     operator_time                      C 1 ����ʱ��
  --     item_list[]����ɾ���б�
  --                  stuffdtl_id            N 1 ������ϸid,Ŀǰ����ķ���ID
  --                  return_num             N 1 ��������
  --     roll_list[]�����ջ��б�
  --                  clinic_type            C 1 ҽ���������
  --                  stuff_no               C 1 �������ʵĵ��ݺ�
  --                  stuffdtl_id            N 1 ������ϸid,���ü�¼id
  --                  stuffdtl_id_old        N 1 ԭʼ������ϸid,���ü�¼id
  --                  packages_num           N 1 ����
  --                  outbound_num           N 1 ����
  --                  is_stuff_order         N 1 �����Ƿ��ǰ󶨵����ķ���0-������ҽ��,1-����ҽ��

  --����: Json_Out,��ʽ����
  --  output
  --    code                          N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Input Pljson;
  j_Item  Pljson;
  j_List  Pljson_List := Pljson_List();

  n_������ϸid ҩƷ�շ���¼.Id%Type;
  n_��������   ҩƷ�շ���¼.��д����%Type;
  v_��Ա����   Varchar2(3000);
  �ջ�ʱ��_In  Date;
  v_Dec        Number;
  -- v_�������   Varchar2(3000);
  v_�շ����    Number;
  No_In         Varchar2(3000);
  v_����id      Number;
  Old_����id    Number;
  v_��ǰ����    Number;
  v_��ǰ����    Number;
  v_�������    Varchar2(3000);
  n_����ҽ��    Number;
  n_����        Number;
  n_Count       Number;
  n_�շѱ�־    Number;
  n_�Զ�����    Number;
  v_������ϸids Varchar2(32767);

  Cursor c_Stuff Is
    Select b.����, Nvl(x.���÷���, 0) As ����, b.����, b.Ч��, x.���Ч��, b.Id As �շ�id, b.����id, b.��ҳid, b.�ⷿid, b.����, b.����, b.�Է�����id,
           b.���
    From ҩƷ�շ���¼ B, �������� X
    Where b.����id = Old_����id And b.ҩƷid = x.����id;

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
    
    P���� ҩƷ�շ���¼.����%Type,
    P���� ҩƷ�շ���¼.��д����%Type,
    P��� ҩƷ�շ���¼.���%Type
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
    If P��� Is Not Null Then
      Select Max(b.���ȼ�) Into v_���ȼ� From ��� B Where b.���� = P���;
    End If;
    If Sql%RowCount = 0 Then
      Insert Into δ��ҩƷ��¼
        (����, NO, ����id, ��ҳid, ����, ���ȼ�, �Է�����id, �ⷿid, ��������, ���շ�, ��ӡ״̬)
      Values
        (����_In, No_In, ����id_In, ��ҳid_In, ����_In, v_���ȼ�, �Է�����id_In, �ⷿid_In, �ջ�ʱ��_In, n_�շѱ�־, 0);
    End If;
  End;

Begin
  j_Item  := Pljson(Json_In);
  j_Input := j_Item.Get_Pljson('input');

  --ҩƷɾ���б�
  j_List := j_Input.Get_Pljson_List('item_list');
  If j_List Is Not Null Then
    For I In 1 .. j_List.Count Loop
      j_Item       := Pljson();
      j_Item       := Pljson(j_List.Get(I));
      n_������ϸid := j_Item.Get_Number('stuffdtl_id');
      n_��������   := j_Item.Get_Number('return_num');
      Zl_�����շ���¼_�����˷�_s(n_������ϸid, n_��������, 1);
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
    
      Select Nvl(Max(���), 0) + 1 Into v_�շ���� From ҩƷ�շ���¼ Where ���� = 25 And ��¼״̬ = 1 And NO = No_In;
    
      v_������� := j_Item.Get_String('clinic_type');
      No_In      := j_Item.Get_String('stuff_no');
      v_����id   := j_Item.Get_Number('stuffdtl_id');
      Old_����id := j_Item.Get_Number('stuffdtl_id_old');
      v_��ǰ���� := j_Item.Get_Number('packages_num');
      v_��ǰ���� := j_Item.Get_Number('outbound_num');
      n_����ҽ�� := j_Item.Get_Number('is_stuff_order'); --�����Ƿ��ǰ󶨵�ҩƷ����0-������ҽ��,1-����ҽ��
      n_�շѱ�־ := j_Item.Get_Number('charge_tag');
      n_�Զ����� := j_Item.Get_Number('stuff_auto_send');
      If n_�Զ����� = 1 Then
        v_������ϸids := v_������ϸids || ',' || v_����id;
      End If;
      For r_Stuff In c_Stuff Loop
        If n_����ҽ�� = 1 Then
          �����շ���¼_Insert(v_����id, r_Stuff.����, r_Stuff.����, r_Stuff.����, r_Stuff.Ч��, r_Stuff.���Ч��, r_Stuff.�շ�id,
                        r_Stuff.����id, r_Stuff.��ҳid, r_Stuff.�ⷿid, r_Stuff.����, r_Stuff.����, r_Stuff.�Է�����id, v_��ǰ����, v_��ǰ����,
                        r_Stuff.���);
        Else
          n_���� := v_��ǰ����;
          For r_Otherstuff In (Select b.����, Nvl(x.���÷���, 0) As ����, Nvl(b.����, 1) * b.ʵ������ As ����, b.����, b.Ч��, x.���Ч��,
                                      b.Id As �շ�id, b.����id, b.��ҳid, b.�ⷿid, b.����, b.����, b.�Է�����id, b.���
                               From ҩƷ�շ���¼ B, �������� X
                               Where b.����id = Old_����id And (b.��¼״̬ = 1 Or Mod(b.��¼״̬, 3) = 0) And b.ҩƷid = x.����id
                               Order By b.Id Desc) Loop
            If n_���� > 0 Then
              n_Count := r_Otherstuff.����;
              If n_���� < n_Count Then
                n_Count := n_����;
              End If;
              �����շ���¼_Insert(v_����id, r_Otherstuff.����, r_Otherstuff.����, r_Otherstuff.����, r_Otherstuff.Ч��,
                            r_Otherstuff.���Ч��, r_Otherstuff.�շ�id, r_Otherstuff.����id, r_Otherstuff.��ҳid,
                            r_Otherstuff.�ⷿid, r_Otherstuff.����, r_Otherstuff.����, r_Otherstuff.�Է�����id, 1, n_Count,
                            r_Otherstuff.���);
              n_���� := n_���� - r_Otherstuff.����;
            End If;
          End Loop;
        End If;
      End Loop;
    End Loop;
  End If;

  If v_������ϸids Is Not Null Then
    v_������ϸids := Substr(v_������ϸids, 2);
    Zl_�����շ���¼_�Զ�����_s(Null, v_��Ա����, Null, v_������ϸids, 1);
  End If;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Stuffsvr_Overdue_Recovery;
/