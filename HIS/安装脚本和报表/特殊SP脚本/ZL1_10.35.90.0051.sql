----------------------------------------------------------------------------------------------------------------
--���ű�֧�ִ�ZLHIS+ v10.35.90������ v10.35.90
--�������ݿռ�������ߵ�¼PLSQL��ִ�����нű�
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------
--�ṹ��������
------------------------------------------------------------------------------



------------------------------------------------------------------------------
--������������
------------------------------------------------------------------------------
--138241:���˺�,2019-02-28,��������ʶ�����
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, NULL, 0, 0, 0, 0, 0, 0, 320, '��������ʶ�����', '0', '0',
          ' ������������������ı���ɨ��¼�룬���ڼ��ʡ����ѵ�ģ����¼������������������ʱ��ֻ��ͨ��������¼��;��������ͨ�����롢���롢����Ƚ���¼�롣',
         '1-����ͨ��ɨ��¼���¼���������ϱ���������;0-�����ƣ�����ͨ�����롢���롢�����¼�뷽ʽ��ʶ����������', 
		 '����������Ϊ1-����ͨ��ɨ��¼���¼���������ϱ���������ʱ����Ҫ����������������������Ч��',
         '��������Ҫ�����ͨ�������ϸ������û���', Null
  From Dual;

--123946:��¶¶,2019-02-25,�����������������¼��ȡ�������½����ϵ�����
Insert into ���䷽ʽ values (3, '�ʹ���', 'PGC', 0);

Insert into ���䷽ʽ values (4, '���', 'ZC', 0);

Insert into ���䷽ʽ values (5, '��ǯ', 'CQ', 0);

Insert into ���䷽ʽ values (6, '�γ�', 'TC', 0);

Insert into ���䷽ʽ values (7, '����', 'TZ', 0);


-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------
--137272:���ϴ�,2019-02-25,ԤԼ��ǰ�����������
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1111, '����', User, 'Zl_�Һ����״̬_Lock', 'EXECUTE' From Dual;

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 9000, '����', User, 'Zl_�Һ����״̬_Lock', 'EXECUTE' From Dual;




-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--138040:Ƚ����,2019-02-28,ҽ�����ö൥��һ�ν��㲿���˷ѣ�ȫ�˺����յġ�������ü�¼.�Ƿ��ϴ�����д����
Create Or Replace Procedure Zl_�����շѼ�¼_����
(
  ԭ����id_In     ������ü�¼.����id%Type,
  ����id_In       ������ü�¼.����id%Type,
  ���ս���id_In   ������ü�¼.����id%Type,
  �ſ�ҽ������_In Varchar2 := Null
) As
  --�ſ�ҽ������_IN:����ö��ŷ���(ֻĳЩҽ������,�������ֽ�)
  Cursor c_Fee_Data Is
    Select ID
    From ������ü�¼ A
    Where ����id = ԭ����id_In And Not Exists
     (Select 1
           From ������ü�¼ B
           Where Mod(b.��¼����, 10) = 1 And a.No = b.No And a.��� = b.��� And ����id = ����id_In)
    Order By ID;

  v_����Ա��� ������ü�¼.����Ա���%Type;
  v_����Ա���� ������ü�¼.����Ա����%Type;
  d_�Ǽ�ʱ��   ������ü�¼.�Ǽ�ʱ��%Type;
  n_�ɿ���id   ������ü�¼.�ɿ���id%Type;
  n_����id     ������ü�¼.����id%Type;
  Err_Item Exception;
  v_Err_Msg    Varchar2(255);
  n_Array_Size Number := 200;
  t_����id     t_Numlist;
  n_������   ������ü�¼.ʵ�ս��%Type;
  n_�������   ����Ԥ����¼.��Ԥ��%Type;
  n_Count      Number(18);
Begin
  Begin
    Select ����Ա���, ����Ա����, �Ǽ�ʱ��, �ɿ���id, ����id
    Into v_����Ա���, v_����Ա����, d_�Ǽ�ʱ��, n_�ɿ���id, n_����id
    From ������ü�¼
    Where ����id = ����id_In And Rownum < 2;
  Exception
    When Others Then
      v_Err_Msg := 'NO';
  End;

  If Nvl(v_Err_Msg, '-') = 'NO' Then
    v_Err_Msg := '���ڲ�������,�õ��ݿ����Ѿ��������˷ѻ�ɾ��,�����ٽ����˷Ѳ�����';
    Raise Err_Item;
  End If;

  --1.�������ѡ������ǲ����˻򲿷�ִ�е�
  Insert Into ������ü�¼
    (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ,
     ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����, ִ��״̬, ����״̬, ִ��ʱ��,
     ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, ����, �ɿ���id, �Һ�id, ��ҳid)
    Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id,
           �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ������,
           ִ����, ִ��״̬, ����״̬, ִ��ʱ��, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, ����,
           �ɿ���id, �Һ�id, ��ҳid
    From (Select NO, Max(ʵ��Ʊ��) As ʵ��Ʊ��, 11 As ��¼����, 1 As ��¼״̬, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ,
                  �ѱ�, ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ, 1 As ����, Max(��ҩ����) As ��ҩ����, Sum(Nvl(����, 1) * Nvl(����, 0)) As ����,
                  Max(�Ӱ��־) As �Ӱ��־, Max(���ӱ�־) As ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, Avg(��׼����) As ��׼����, Sum(Ӧ�ս��) As Ӧ�ս��,
                  Sum(ʵ�ս��) As ʵ�ս��, ��������id, ������, ִ�в���id, Max(������) As ������, Max(ִ����) ִ����, Max(ִ��״̬) As ִ��״̬, 1 As ����״̬,
                  Max(ִ��ʱ��) ִ��ʱ��, v_����Ա��� As ����Ա���, v_����Ա���� As ����Ա����, ����ʱ��, d_�Ǽ�ʱ�� As �Ǽ�ʱ��, ���ս���id_In As ����id,
                  Sum(���ʽ��) As ���ʽ��, Max(������Ŀ��) As ������Ŀ��, ���մ���id, Sum(ͳ����) As ͳ����,
                  Max(Decode(��¼����, 1, ժҪ, 11, ժҪ, Null)) As ժҪ, 0 As �Ƿ��ϴ�, Max(���ձ���) As ���ձ���, Max(��������) As ��������,
                  Max(Decode(��¼����, 1, ����, 11, ����, Null)) As ����, n_�ɿ���id As �ɿ���id, Max(�Һ�id) As �Һ�id, Max(��ҳid) As ��ҳid
           From ������ü�¼
           Where Mod(��¼����, 10) = 1 And (NO, ���) In (Select NO, ��� From ������ü�¼ Where ����id = ����id_In)
           Group By NO, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀid,
                    �վݷ�Ŀ, ���ʷ���, ��������id, ������, ִ�в���id, ����ʱ��, ���մ���id
           Having Sum(Nvl(����, 1) * Nvl(����, 0)) <> 0);

  For c_���� In (Select NO, ���, ��������, �۸񸸺�, ������Ŀid, -1 * Sum(Nvl(����, 1) * Nvl(����, 0)) As ����, Sum(��׼����) As ��׼����,
                      -1 * Sum(Ӧ�ս��) As Ӧ�ս��, -1 * Sum(ʵ�ս��) As ʵ�ս��, -1 * Sum(ͳ����) As ͳ����, -1 * Sum(���ʽ��) As ���ʽ��
               From ������ü�¼
               Where ��¼���� = 11 And ����id = ���ս���id_In
               Group By NO, ���, ��������, �۸񸸺�, ������Ŀid) Loop
    Update ������ü�¼
    Set ���� = Nvl(����, 0) + Nvl(c_����.����, 0), ʵ�ս�� = Nvl(ʵ�ս��, 0) + Nvl(c_����.ʵ�ս��, 0),
        Ӧ�ս�� = Nvl(Ӧ�ս��, 0) + Nvl(c_����.Ӧ�ս��, 0), ���ʽ�� = Nvl(���ʽ��, 0) + Nvl(c_����.���ʽ��, 0),
        ͳ���� = Nvl(ͳ����, 0) + Nvl(c_����.ͳ����, 0)
    Where NO = c_����.No And ��� = c_����.��� And Nvl(��������, -1) = Nvl(c_����.��������, '-1') And
          Nvl(�۸񸸺�, -1) = Nvl(c_����.�۸񸸺�, '-1') And ������Ŀid = c_����.������Ŀid And ����id = ����id_In;
  End Loop;

  --2.�������δѡ�˷Ѳ���,��Ҫȫ���Ҳ���11�����ռ�¼
  Open c_Fee_Data;
  Loop
    Fetch c_Fee_Data Bulk Collect
      Into t_����id Limit n_Array_Size;
    Exit When t_����id.Count = 0;
  
    --�˷Ѽ�¼
    Forall I In 1 .. t_����id.Count
      Insert Into ������ü�¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id, �շ����, �շ�ϸĿid,
         ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����, ִ��״̬, ����״̬,
         ִ��ʱ��, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, ����, �ɿ���id, �Һ�id, ��ҳid)
        Select ���˷��ü�¼_Id.Nextval, a.No, a.ʵ��Ʊ��, 1, 2, a.���, a.��������, a.�۸񸸺�, a.����id, a.ҽ�����, a.�����־, a.����, a.�Ա�, a.����,
               a.��ʶ��, a.���ʽ, a.�ѱ�, a.���˿���id, a.�շ����, a.�շ�ϸĿid, a.���㵥λ, a.����, a.��ҩ����, -1 * a.����, a.�Ӱ��־, a.���ӱ�־,
               a.������Ŀid, a.�վݷ�Ŀ, a.���ʷ���, a.��׼����, -1 * a.Ӧ�ս��, -1 * a.ʵ�ս��, a.��������id, a.������, a.ִ�в���id, a.������, ִ����,
               Nvl(q.ִ��״̬, -1) As ִ��״̬, 1, a.ִ��ʱ��, v_����Ա���, v_����Ա����, a.����ʱ��, d_�Ǽ�ʱ��, ����id_In, -1 * a.���ʽ��, a.������Ŀ��,
               a.���մ���id, -1 * a.ͳ����, a.ժҪ, 0 As �Ƿ��ϴ�, a.���ձ���, a.��������, a.����, n_�ɿ���id As �ɿ���id, �Һ�id, ��ҳid
        From ������ü�¼ A,
             (Select j.No, j.���, Nvl(Max(j.ִ��״̬), 0) - 1 As ִ��״̬
               From ������ü�¼ M, ������ü�¼ J
               Where m.Id = t_����id(I) And m.No = j.No And m.��� = j.��� And Mod(j.��¼����, 10) = 1 And j.��¼״̬ = 2
               Group By j.No, j.���) Q
        Where ID = t_����id(I) And a.No = q.No(+) And a.��� = q.���(+);
  
    --��ԭ��¼״̬��1��Ϊ3
    Forall I In 1 .. t_����id.Count
      Update ������ü�¼ Set ��¼״̬ = 3 Where ID = t_����id(I) And ��¼״̬ = 1;
  
    --�����շѼ�¼
    If Nvl(���ս���id_In, 0) <> 0 Then
      Forall I In 1 .. t_����id.Count
        Insert Into ������ü�¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id, �շ����, �շ�ϸĿid,
           ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����, ִ��״̬,
           ����״̬, ִ��ʱ��, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, ����, �ɿ���id, �Һ�id,
           ��ҳid)
          Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, 11, 1, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id,
                 �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id,
                 ������, ִ����, ִ��״̬, 1, ִ��ʱ��, v_����Ա���, v_����Ա����, ����ʱ��, d_�Ǽ�ʱ��, ���ս���id_In, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ,
                 0 As �Ƿ��ϴ�, ���ձ���, ��������, ����, n_�ɿ���id As �ɿ���id, �Һ�id, ��ҳid
          From ������ü�¼
          Where ID = t_����id(I);
    End If;
  End Loop;
  Close c_Fee_Data;

  Select Count(1) Into n_Count From ����Ԥ����¼ Where ����id = ����id_In And ���㷽ʽ Is Null;
  If n_Count = 0 Then
    --�˷ѽ��㷽ʽ
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, ��������)
      Select ����Ԥ����¼_Id.Nextval, 3, Null, 2, n_����id, ���㷽ʽ, d_�Ǽ�ʱ��, v_����Ա���, v_����Ա����, -1 * ��Ԥ��, ����id_In, n_�ɿ���id,
             -1 * ����id_In, 2, 3
      From ����Ԥ����¼
      Where ����id = ԭ����id_In And ���㷽ʽ In (Select ���� From ���㷽ʽ Where ���� In (3, 4)) And
            Instr(',' || �ſ�ҽ������_In || ',', ',' || ���㷽ʽ || ',') = 0 And Mod(��¼����, 10) <> 1;
    --��ԭ����ȫ������
    --Insert Into ����Ԥ����¼
    --  (ID, ��¼����, NO, ��¼״̬, ����id, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־)
    --  Select ����Ԥ����¼_Id.Nextval, 3, Null, 2, n_����id, ���㷽ʽ, d_�Ǽ�ʱ��, v_����Ա���, v_����Ա����, -1 * ��Ԥ��, ����id_In, n_�ɿ���id,
    --         -1 * ����id_In, 2
    --  From ����Ԥ����¼
    --  Where ����id = ԭ����id_In And ���㷽ʽ = v_���� And Mod(��¼����, 10) <> 1;
  
    Select Sum(��Ԥ��) Into n_������� From ����Ԥ����¼ Where ����id = ����id_In;
    Select Sum(���ʽ��) Into n_������ From ������ü�¼ Where ����id = ����id_In;
  
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, ��������)
    Values
      (����Ԥ����¼_Id.Nextval, 3, Null, 2, n_����id, Null, d_�Ǽ�ʱ��, v_����Ա���, v_����Ա����, -1 * (Nvl(n_�������, 0) - Nvl(n_������, 0)),
       ����id_In, n_�ɿ���id, -1 * ����id_In, 1, 3);
  
  End If;
  If Nvl(���ս���id_In, 0) <> 0 Then
    Select Count(1) Into n_Count From ����Ԥ����¼ Where ����id = ���ս���id_In And ���㷽ʽ Is Null;
    If n_Count = 0 Then
      Select Sum(���ʽ��) Into n_������ From ������ü�¼ Where ����id = ���ս���id_In;
    
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, 3, Null, 1, n_����id, Null, d_�Ǽ�ʱ��, v_����Ա���, v_����Ա����, n_������, ���ս���id_In, n_�ɿ���id,
         -1 * ����id_In, 1, 3);
    End If;
  
    --����_In    Integer:=0, --0:����;1-סԺ
    --����_In    Integer:=1, --1-�շѵ�;2-���ʵ�
    --����_In    Integer:=0, --0:ɾ�����۵�;1-�շѻ����;2-�˷ѻ�����
    --No_In      ������ü�¼.No%Type,
    --ҽ��ids_In Varchar2 := Null
    For c_No In (Select Distinct NO From ������ü�¼ Where ��¼���� = 11 And ����id = ���ս���id_In) Loop
      Zl_ҽ������_�Ʒ�״̬_Update(0, 1, 2, c_No.No);
    End Loop;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����շѼ�¼_����;
/

--127817:��¶¶,2019-02-25,������ҳ��ȡ����ʱ������ȡ���������
Create Or Replace Function Zl_Adderss_Structure(v_Addressinfo Varchar2,n_Type Number := Null) Return Varchar2 Is
  --���ؽṹ��ʡ,ʡ����,�Ƿ�����,�Ƿ���ʾ,�Ƿ�ֻ�����⼶|��,�б���,�Ƿ�����,�Ƿ���ʾ,�Ƿ�ֻ�����⼶ 
  --          |����,���ر���,�Ƿ�����,�Ƿ���ʾ,�Ƿ�ֻ�����⼶|����,�������,�Ƿ�����,�Ƿ���ʾ,�Ƿ�ֻ�����⼶ 
  --          |�ֵ�,�ֵ�����,�Ƿ�����,�Ƿ���ʾ,�Ƿ�ֻ�����⼶ 
  v_ʡ       Varchar2(100);
  v_Codeʡ   Varchar2(15);
  v_Infoʡ   Varchar2(150);
  v_��       Varchar2(100);
  v_Code��   Varchar2(15);
  v_Info��   Varchar2(150);
  v_����     Varchar2(100);
  v_Code���� Varchar2(15);
  v_Info���� Varchar2(150);
  v_����     Varchar2(100);
  v_Code���� Varchar2(15);
  v_Info���� Varchar2(150);
  v_�ֵ�     Varchar2(500);
  v_Code�ֵ� Varchar2(15);
  v_Info�ֵ� Varchar2(550);
  v_Tmp      Varchar2(100);
  v_Adrstmp  Varchar2(500);
  n_Pos      Number(5);
  n_����     Number(1);
  n_����ʾ   Number(1);
  n_Count    Number(3);
  v_Return   Varchar2(700);
Begin
  --����ṹ���ĵ�ַ�����ý��е�ַ��׼���ָ���� 
  v_Adrstmp := v_Addressinfo;
  If v_Addressinfo Like '%,%,%,%,%' Then
    n_Pos     := Instr(v_Adrstmp, ',');
    v_ʡ      := Substr(v_Adrstmp, 1, n_Pos - 1);
    v_Adrstmp := Substr(v_Adrstmp, n_Pos + 1);
    n_Pos     := Instr(v_Adrstmp, ',');
    v_��      := Substr(v_Adrstmp, 1, n_Pos - 1);
    v_Adrstmp := Substr(v_Adrstmp, n_Pos + 1);
    n_Pos     := Instr(v_Adrstmp, ',');
    v_����    := Substr(v_Adrstmp, 1, n_Pos - 1);
    v_Adrstmp := Substr(v_Adrstmp, n_Pos + 1);
    n_Pos     := Instr(v_Adrstmp, ',');
    v_����    := Substr(v_Adrstmp, 1, n_Pos - 1);
    v_�ֵ�    := Substr(v_Adrstmp, n_Pos + 1);
    Select Max(����) Into v_Codeʡ From ���� Where ���� = v_ʡ And Nvl(����, 0) = 0;
    --ʡ����ַ��û�У��Ͳ������� 
    If v_Codeʡ Is Not Null Then
      Select Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ)
      Into v_Code��, n_����, n_����ʾ
      From ����
      Where ���� = v_�� And Nvl(����, 0) = 1 And �ϼ����� = v_Codeʡ;
      If v_Code�� Is Not Null Then
        v_Info�� := v_�� || ',' || v_Code�� || ',' || n_���� || ',' || n_����ʾ;
        Select Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ)
        Into v_Code����, n_����, n_����ʾ
        From ����
        Where ���� = v_���� And Nvl(����, 0) = 2 And �ϼ����� = v_Code��;
        --�����������ַ 
      Else
        Select Max(����), Max(�ϼ�����)
        Into v_Code����, v_Code��
        From ����
        Where ���� = v_���� And Nvl(����, 0) = 2 And �ϼ����� In (Select ���� From ���� Where �ϼ����� = v_Codeʡ);
        If v_Code�� Is Not Null Then
          Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ)
          Into v_��, v_Code��, n_����, n_����ʾ
          From ����
          Where ���� = v_Code��;
        End If;
        v_Info�� := v_�� || ',' || v_Code�� || ',' || n_���� || ',' || n_����ʾ;
        Select Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ)
        Into v_Code����, n_����, n_����ʾ
        From ����
        Where ���� = v_Code����;
      End If;
      v_Info���� := v_���� || ',' || v_Code���� || ',' || n_���� || ',' || n_����ʾ;
      If v_Code���� Is Not Null Then
        --������������ϸ��ַ�У��������������ַ�ṹ��¼�� 

        Select Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ)
        Into v_Code����, n_����, n_����ʾ
        From ����
        Where ���� = v_���� And Nvl(����, 0) = 3 And �ϼ����� = v_Code����;
        --�����������ַ 
        If v_Code���� Is Null Then
          Select Max(����), Max(�ϼ�����)
          Into v_Code�ֵ�, v_Code����
          From ����
          Where ���� = v_�ֵ� And Nvl(����, 0) = 4 And �ϼ����� In (Select ���� From ���� Where �ϼ����� = v_Code����);
          If v_Code���� Is Not Null Then
            Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ)
            Into v_����, v_Code����, n_����, n_����ʾ
            From ����
            Where ���� = v_Code����;
          End If;
        End If;
        v_Info���� := v_���� || ',' || v_Code���� || ',' || n_���� || ',' || n_����ʾ;
        If v_Code���� Is Not Null Then
          Select Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ)
          Into v_Code�ֵ�, n_����, n_����ʾ
          From ����
          Where ���� = v_�ֵ� And Nvl(����, 0) = 4 And �ϼ����� = v_Code����;
          v_Info�ֵ� := v_�ֵ� || ',' || v_Code�ֵ� || ',' || n_���� || ',' || n_����ʾ;
        End If;
      End If;
    End If;
    --�Ǳ�׼��ַ����������ַ����Ҫ�ָ�ʡ���У���, 
  Else
    v_Adrstmp := v_Addressinfo;
    v_Tmp     := Substr(v_Adrstmp, 1, 2);
    Select Max(����), Max(����) Into v_ʡ, v_Codeʡ From ���� Where ���� Like v_Tmp || '%' And Nvl(����, 0) = 0;
    --��ʡ����ַ��˵�����Խṹ�� 
    If v_Codeʡ Is Not Null Then
      --ʡ����ַ�Ǳ�׼�� 
      If Substr(v_Adrstmp, 1, Length(v_ʡ)) = v_ʡ Then
        v_Adrstmp := Substr(v_Adrstmp, Length(v_ʡ) + 1);
        --ʡ����ַ����׼,�����½�ʡ����������,��ʱ���м���ַ�����Ǳ�׼���ġ� 
      Else
        --���ж϶�����ַ�Ƿ���������ַ�벻��ʾ�ĵ�ַ 
        If v_Tmp = '����' Then
          v_Tmp := '���ɹ�';
        Elsif v_Tmp = '����' Then
          v_Tmp := '������';
        End If;
        v_Adrstmp := Substr(v_Adrstmp, Length(v_Tmp) + 1);
      End If;
      --�Ƚ�ȡ�м������������ؼ��֣���ƥ�� 
      v_Tmp := Substr(v_Adrstmp, 1, 2);
      If Nvl(n_Type, 0) <> 2 Then
        Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ), Count(1)
        Into v_��, v_Code��, n_����, n_����ʾ, n_Count
        From ����
        Where ���� Like v_Tmp || '%' And Nvl(����, 0) = 1 And �ϼ����� = v_Codeʡ;
	  End If;
      --���ڶ���ƥ�䣬��������ӳ��ȣ���ʱ��2�������ӵ�3����ƥ�� 
      If n_Count > 1 Then
        v_Tmp := Substr(v_Adrstmp, 1, 3);
        Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ)
        Into v_��, v_Code��, n_����, n_����ʾ
        From ����
        Where ���� Like v_Tmp || '%' And Nvl(����, 0) = 1 And �ϼ����� = v_Codeʡ;
      End If;
      --�ж��Ƿ���������ַ����ʾ�ĵ�ַ���µ�,������ڣ�����ݵ�������ַ��ȷ�������ַ 
      --������û�еڶ����������Ҫ�������ж�
      If v_Code�� Is Null Then
        Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ), Count(1), Max(�ϼ�����)
        Into v_����, v_Code����, n_����, n_����ʾ, n_Count, v_Code��
        From ����
        Where ���� Like v_Tmp || '%' And Nvl(����, 0) = 2 And �ϼ����� In (Select ���� From ���� Where �ϼ����� = v_Codeʡ);
        --���ڶ���ƥ�䣬��������ӳ��ȣ���ʱ��2�������ӵ�3����ƥ�� 
        If n_Count > 1 Then
          v_Tmp := Substr(v_Adrstmp, 1, 3);
          Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ), Max(�ϼ�����)
          Into v_����, v_Code����, n_����, n_����ʾ, v_Code��
          From ����
          Where ���� Like v_Tmp || '%' And Nvl(����, 0) = 2 And �ϼ����� In (Select ���� From ���� Where �ϼ����� = v_Codeʡ);
        End If;
        If v_Code�� Is Not Null Then
          v_Info���� := v_���� || ',' || v_Code���� || ',' || n_���� || ',' || n_����ʾ;
          v_Adrstmp  := Substr(v_Adrstmp, Length(v_����) + 1);
          Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ)
          Into v_��, v_Code��, n_����, n_����ʾ
          From ����
          Where ���� = v_Code��;
          v_Info�� := v_�� || ',' || v_Code�� || ',' || n_���� || ',' || n_����ʾ;
        End If;
      Else
        v_Info��  := v_�� || ',' || v_Code�� || ',' || n_���� || ',' || n_����ʾ;
        v_Adrstmp := Substr(v_Adrstmp, Length(v_��) + 1);
      End If;
      --û�����أ���������� 
      If Not v_Code�� Is Null And v_Code���� Is Null Then
        --�Ƚ�ȡ�ؼ������������ؼ��֣���ƥ�� 
        v_Tmp := Substr(v_Adrstmp, 1, 2);
        Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ), Count(1)
        Into v_����, v_Code����, n_����, n_����ʾ, n_Count
        From ����
        Where ���� Like v_Tmp || '%' And Nvl(����, 0) = 2 And �ϼ����� = v_Code��;
        --���ڶ���ƥ�䣬��������ӳ��ȣ���ʱ��2�������ӵ�3����ƥ�� 
        If n_Count > 1 Then
          v_Tmp := Substr(v_Adrstmp, 1, 3);
          Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ)
          Into v_����, v_Code����, n_����, n_����ʾ
          From ����
          Where ���� Like v_Tmp || '%' And Nvl(����, 0) = 2 And �ϼ����� = v_Code��;
        End If;
        If v_Code���� Is Null Then
          Select Max(�Ƿ�����), Max(�Ƿ���ʾ) Into n_����, n_����ʾ From ���� Where �ϼ����� = v_Code��;
          If Nvl(n_����, 0) = 1 Or Nvl(n_����ʾ, 0) = 1 Then
            Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ), Count(1), Max(�ϼ�����)
            Into v_����, v_Code����, n_����, n_����ʾ, n_Count, v_Code����
            From ����
            Where ���� Like v_Tmp || '%' And Nvl(����, 0) = 3 And �ϼ����� In (Select ���� From ���� Where �ϼ����� = v_Code��);
            --���ڶ���ƥ�䣬��������ӳ��ȣ���ʱ��2�������ӵ�3����ƥ�� 
            If n_Count > 1 Then
              v_Tmp := Substr(v_Adrstmp, 1, 3);
              Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ), Max(�ϼ�����)
              Into v_����, v_Code����, n_����, n_����ʾ, v_Code����
              From ����
              Where ���� Like v_Tmp || '%' And Nvl(����, 0) = 3 And �ϼ����� In (Select ���� From ���� Where �ϼ����� = v_Code��);
            End If;
          
            If v_Code���� Is Not Null Then
              v_Info���� := v_���� || ',' || v_Code���� || ',' || n_���� || ',' || n_����ʾ;
              v_Adrstmp  := Substr(v_Adrstmp, Length(v_����) + 1);
              Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ)
              Into v_����, v_Code����, n_����, n_����ʾ
              From ����
              Where ���� = v_Code����;
              v_Info���� := v_���� || ',' || v_Code���� || ',' || n_���� || ',' || n_����ʾ;
            End If;
          End If;
        Else
          v_Info���� := v_���� || ',' || v_Code���� || ',' || n_���� || ',' || n_����ʾ;
          v_Adrstmp  := Substr(v_Adrstmp, Length(v_����) + 1);
        End If;
      End If;
      If v_Code���� Is Not Null And v_Code���� Is Null Then
        --�Ƚ�ȡ���򼶵����������ؼ��֣���ƥ�� 
        v_Tmp := Substr(v_Adrstmp, 1, 2);
        Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ), Count(1)
        Into v_����, v_Code����, n_����, n_����ʾ, n_Count
        From ����
        Where ���� Like v_Tmp || '%' And Nvl(����, 0) = 3 And �ϼ����� = v_Code����;
        --���ڶ���ƥ�䣬��������ӳ��ȣ���ʱ��2�������ӵ�3����ƥ�� 
        If n_Count > 1 Then
          v_Tmp := Substr(v_Adrstmp, 1, 3);
          Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ), Count(1)
          Into v_����, v_Code����, n_����, n_����ʾ, n_Count
          From ����
          Where ���� Like v_Tmp || '%' And Nvl(����, 0) = 3 And �ϼ����� = v_Code����;
        End If;
        If v_Code���� Is Null Then
          Select Max(�Ƿ�����), Max(�Ƿ���ʾ) Into n_����, n_����ʾ From ���� Where �ϼ����� = v_Code����;
          If Nvl(n_����, 0) = 1 Or Nvl(n_����ʾ, 0) = 1 Then
            Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ), Count(1), Max(�ϼ�����)
            Into v_�ֵ�, v_Code�ֵ�, n_����, n_����ʾ, n_Count, v_Code����
            From ����
            Where ���� = v_Adrstmp And �ϼ����� In (Select ���� From ���� Where �ϼ����� = v_Code����);
          End If;
          If v_Code���� Is Not Null Then
            v_Info�ֵ� := v_�ֵ� || ',' || v_Code�ֵ� || ',' || n_���� || ',' || n_����ʾ;
            Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ)
            Into v_����, v_Code����, n_����, n_����ʾ
            From ����
            Where ���� = v_Code����;
            v_Info���� := v_���� || ',' || v_Code���� || ',' || n_���� || ',' || n_����ʾ;
          End If;
        Else
          v_Info���� := v_���� || ',' || v_Code���� || ',' || n_���� || ',' || n_����ʾ;
          v_Adrstmp  := Substr(v_Adrstmp, Length(v_����) + 1);
        End If;
        If v_Code���� Is Not Null And v_Code�ֵ� Is Null Then
          Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ)
          Into v_�ֵ�, v_Code�ֵ�, n_����, n_����ʾ
          From ����
          Where ���� = v_Adrstmp And Nvl(����, 0) = 4 And �ϼ����� = v_Code����;
          If v_Code�ֵ� Is Not Null Then
            v_Info�ֵ� := v_�ֵ� || ',' || v_Code�ֵ� || ',' || n_���� || ',' || n_����ʾ;
          End If;
        End If;
      End If;
    End If;
    If v_�ֵ� Is Null Then
      v_�ֵ� := v_Adrstmp;
    End If;
  End If;
  v_Infoʡ := v_ʡ || ',' || v_Codeʡ || ',,,';
  If v_Info�� Is Null Then
    v_Info�� := v_�� || ',,,';
  End If;
  --ֻ��ʡû���У��ж����Ƿ�ֻ�����⼶ 
  If Not v_Codeʡ Is Null And v_�� Is Null Then
    Select Count(1)
    Into n_Count
    From ����
    Where �ϼ����� = v_Codeʡ And Nvl(�Ƿ�����, 0) = 0 And Nvl(�Ƿ���ʾ, 0) = 0 And Rownum < 2;
    If n_Count = 0 Then
      Select Count(1) Into n_Count From ���� Where �ϼ����� = v_Codeʡ And Rownum < 2;
      If n_Count = 0 Then
        v_Info�� := v_Info�� || ',';
      Else
        v_Info�� := v_Info�� || ',1';
      End If;
    Else
      v_Info�� := v_Info�� || ',';
    End If;
  Else
    v_Info�� := v_Info�� || ',';
  End If;
  If v_Info���� Is Null Then
    v_Info���� := v_���� || ',,,';
  End If;
  --ֻ����û�����أ��ж�����ֻ�����⼶ 
  If Not v_Code�� Is Null And v_���� Is Null Then
    Select Count(1)
    Into n_Count
    From ����
    Where �ϼ����� = v_Code�� And Nvl(�Ƿ�����, 0) = 0 And Nvl(�Ƿ���ʾ, 0) = 0 And Rownum < 2;
    If n_Count = 0 Then
      Select Count(1) Into n_Count From ���� Where �ϼ����� = v_Code�� And Rownum < 2;
      If n_Count = 0 Then
        v_Info���� := v_Info���� || ',';
      Else
        v_Info���� := v_Info���� || ',1';
      End If;
    Else
      v_Info���� := v_Info���� || ',';
    End If;
  Else
    v_Info���� := v_Info���� || ',';
  End If;
  If v_Info���� Is Null Then
    v_Info���� := v_���� || ',,,';
  End If;
  --ֻ������û�������ж������Ƿ�ֻ��������¼� 
  If Not v_Code���� Is Null And v_���� Is Null Then
    Select Count(1)
    Into n_Count
    From ����
    Where �ϼ����� = v_Code���� And Nvl(�Ƿ�����, 0) = 0 And Nvl(�Ƿ���ʾ, 0) = 0 And Rownum < 2;
    If n_Count = 0 Then
      Select Count(1) Into n_Count From ���� Where �ϼ����� = v_Code���� And Rownum < 2;
      If n_Count = 0 Then
        v_Info���� := v_Info���� || ',';
      Else
        v_Info���� := v_Info���� || ',1';
      End If;
    Else
      v_Info���� := v_Info���� || ',';
    End If;
  Else
    v_Info���� := v_Info���� || ',';
  End If;
  If v_Info�ֵ� Is Null Then
    v_Info�ֵ� := v_�ֵ� || ',,,,';
  Else
    v_Info�ֵ� := v_Info�ֵ� || ',';
  End If;
  v_Return := v_Infoʡ || '|' || v_Info�� || '|' || v_Info���� || '|' || v_Info���� || '|' || v_Info�ֵ�;
  Return(v_Return);
End;
/


------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0051' Where ���=&n_System;
Commit;
