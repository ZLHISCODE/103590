----------------------------------------------------------------------------------------------------------------
--���ű�֧�ִ�ZLHIS+ v10.35.90������ v10.35.90
--�������ݿռ�������ߵ�¼PLSQL��ִ�����нű�
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------
--�ṹ��������
------------------------------------------------------------------------------
--119722:����,2019-01-22,�����ֶ�
Alter Table ҩƷ�ɹ��ƻ� Add ��Դ�ⷿ varchar2(200);
Alter Table ҩƷ�ɹ��ƻ� Add ��Դҩ�� varchar2(200);


------------------------------------------------------------------------------
--������������
------------------------------------------------------------------------------
--136629:����,2019-01-22,ȡ��ϵͳ��������ӣ�����סԺ�����Զ����ϵĲ���ֵ����
Delete From zlParameters Where ϵͳ = &n_System And ������ = '�Զ����ű���������' And ģ�� Is Null And ������ = 320;

Update zlParameters
Set Ӱ�����˵�� = 'ѡ��0�����Զ����ϣ�' || Chr(13) || ' ѡ��1����סԺ��ʿվ����ҽ����סԺҽ��վ����סԺ����ҽ��ʱ��������͵Ĳ��ǻ��۵������ڸ������õ����ģ��Զ����з��ϣ�' || Chr(13) ||
              ' ��¼��סԺ���ʵ������ʱ��Զ�����ʵ�ʱ���Ը������õ����ģ��Զ����з��ϲ�����' || Chr(13) || ' ����ҽ������վ��Ҫ�������Һ�ִ�п���һ�²��Զ����ϡ�' || Chr(13) ||
              ' ѡ��2���ڷ���ҽ����סԺ����ʱ��ֻ�б����ҿ����ĸ����������Ĳ��Զ�����', ����ֵ���� = '0-���Զ����ϣ�1-�Զ����ϣ�ҽ������վ��Ҫ�������Һ�ִ�п���һ�£���2-�����ҿ����Զ�����',
    ����˵�� = '����Ҳ�����ƵĲ���"92-���������Զ�����"',
    ����˵�� = '�Զ��������ڼ��ٻ�ʿ�Ĺ�������һ����˵��ѡ��1����������ڷ��ϲ��Ų��ǿ������ҵ������������ҩ�����������Ĳֿ������ϣ����ҹ���Ҫ���������ֻ���ֹ�ִ�з��ϲ���ʱ����ôѡ��2����Ӧ��������'
Where ϵͳ = &n_System And ������ = 'סԺ�����Զ�����' And ģ�� Is Null And ������ = 63;

-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--136632:���˺�,2019-01-23,���ӱ��ƿ����Զ����Ϲ���.
Create Or Replace Procedure Zl_סԺ���ʼ�¼_Verify
(
  No_In           סԺ���ü�¼.No%Type,
  ����Ա���_In   סԺ���ü�¼.����Ա���%Type,
  ����Ա����_In   סԺ���ü�¼.����Ա����%Type,
  ���_In         Varchar2 := Null,
  ����id_In       סԺ���ü�¼.����id%Type := Null,
  ���ʱ��_In     סԺ���ü�¼.�Ǽ�ʱ��%Type := Null,
  ���ϲ������_In Number := 0
) As
  --���ܣ����һ��סԺ���ʻ��۵�
  --������
  --    ���_IN����ʽ��"1,3,5,7,8",Ϊ�ձ�ʾ�������δ��˵���
  --    ����ID_IN��ֻ���ָ������,���ڰ�������˼��ʱ�
  --    ���ʱ��_IN�����ڲ�����Ҫͳһ���ƻ򷵻�ʱ��ĵط�
  --    ���ϲ������_in:1-���ϲ���ֱ�ӵ������,���Զ�����ʱ������鿪������;0-�Ƿ��ϲ������,���ݲ�����������鿪������
  --ֻ��ȡָ����ŵ�,δ��˵Ĳ��ݽ��д���

  Cursor c_Bill Is
    Select ID, ����id, ��ҳid, �շ�ϸĿid, ʵ�ս��, �����־, ������Ŀid, ִ�в���id, ��������id, ���˲���id, ���˿���id, ҽ�����

    From סԺ���ü�¼
    Where ��¼���� = 2 And ��¼״̬ = 0 And NO = No_In And
          (Instr(',' || ���_In || ',', ',' || Nvl(�۸񸸺�, ���) || ',') > 0 Or ���_In Is Null) And
          (����id + 0 = ����id_In Or ����id_In Is Null)
    Order By ���;

  --����а����������õ�δ������ʱ�����ݲ��������Ƿ��Զ�����
  Cursor c_Stuff Is
    Select ID, �ⷿid
    From ҩƷ�շ���¼ M
    Where NO = No_In And ���� In (25, 26) And �ⷿid Is Not Null And ��¼״̬ = 1 And ����� Is Null And Exists
     (Select 1
           From סԺ���ü�¼ A, �������� B
           Where a.Id = m.����id + 0 And a.��¼���� = 2 And a.��¼״̬ = 1 And a.No = No_In And
                 (Instr(',' || ���_In || ',', ',' || Nvl(a.�۸񸸺�, a.���) || ',') > 0 Or ���_In Is Null) And
                 (a.����id + 0 = ����id_In Or ����id_In Is Null) And a.�շ�ϸĿid = b.����id And b.�������� = 1)
    Order By �ⷿid, ҩƷid;
  --
  v_���Ϻ�         ҩƷ�շ���¼.���ܷ�ҩ��%Type;
  v_�ⷿid         ҩƷ�շ���¼.�ⷿid%Type;
  v_�շ�ids        Varchar2(4000);
  v_ҽ��ids        Varchar2(4000);
  v_Date           Date;
  n_��˱�־       ������ҳ.��˱�־%Type;
  n_סԺ״̬       ������ҳ.״̬%Type;
  n_������˷�ʽ   Number(2);
  n_δ��ƽ�ֹ���� Number(2);

  n_����id  ������ҳ.����id%Type;
  n_��ҳid  ������ҳ.��ҳid%Type;
  v_Err_Msg Varchar2(500);
  Err_Item Exception;
  v_����Ա���   ��Ա��.���%Type;
  v_����Ա����   ��Ա��.����%Type;
  v_Temp         Varchar2(225);
  n_�����Զ����� Number(2);
  n_��������id   סԺ���ü�¼.��������id%Type;
Begin
  If ���ʱ��_In Is Null Then
    Select Sysdate Into v_Date From Dual;
  Else
    v_Date := ���ʱ��_In;
  End If;

  v_����Ա��� := ����Ա���_In;
  v_����Ա���� := ����Ա����_In;
  If v_����Ա��� Is Null Then
    v_Temp := Zl_Identity(1);
    If Not (Nvl(v_Temp, '0') = '0' Or Nvl(v_Temp, '_') = '_') Then
      v_Temp       := Substr(v_Temp, Instr(v_Temp, ';') + 1);
      v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
      v_����Ա��� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
      v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
      v_����Ա���� := v_Temp;
    End If;
  End If;

  For r_Bill In c_Bill Loop
    If Nvl(n_��������id, 0) = 0 Then
      n_��������id := Nvl(r_Bill.��������id, 0);
    End If;
  
    Update סԺ���ü�¼
    Set ��¼״̬ = 1, ����Ա��� = v_����Ա���, ����Ա���� = v_����Ա����, �Ǽ�ʱ�� = v_Date --�Ѳ�����ҩƷ��¼��ʱ�䲻��
    Where ID = r_Bill.Id;
    If Nvl(n_����id, 0) <> Nvl(r_Bill.����id, 0) Then
      If Nvl(zl_GetSysParameter(185), 0) = 1 Then
        n_����id := Nvl(r_Bill.����id, 0);
        n_��ҳid := Nvl(r_Bill.��ҳid, 0);
      
        n_������˷�ʽ   := Nvl(zl_GetSysParameter(185), 0);
        n_δ��ƽ�ֹ���� := Nvl(zl_GetSysParameter(215), 0);
        If n_������˷�ʽ = 1 Or n_δ��ƽ�ֹ���� = 1 Then
          Begin
            Select ��˱�־, ״̬
            Into n_��˱�־, n_סԺ״̬
            From ������ҳ
            Where ����id = n_����id And ��ҳid = n_��ҳid;
          Exception
            When Others Then
              n_��˱�־ := 0;
              n_סԺ״̬ := 0;
          End;
          If n_δ��ƽ�ֹ���� = 1 And n_סԺ״̬ = 1 Then
            v_Err_Msg := '����δ���,��ֹ�Բ�����ط��õĲ���!';
            Raise Err_Item;
          End If;
        
          If n_������˷�ʽ = 1 Then
            If Nvl(n_��˱�־, 0) = 1 Then
              v_Err_Msg := '�ò���Ŀǰ������˷���,���ܽ��з�����ص���!';
              Raise Err_Item;
            End If;
            If Nvl(n_��˱�־, 0) = 2 Then
              v_Err_Msg := '�ò���Ŀǰ�Ѿ�����˷������,���ܽ��з�����ص���!';
              Raise Err_Item;
            End If;
          End If;
        End If;
      End If;
    
    End If;
  
    --ҩƷ�շ���¼.��������
    Update ҩƷ�շ���¼
    Set �������� = Decode(Sign(Nvl(�������, v_Date) - v_Date), -1, ��������, v_Date)
    Where NO = No_In And ���� In (9, 10, 25, 26) And ����id = r_Bill.Id;
  
    --�������
    Update �������
    Set ������� = Nvl(�������, 0) + Nvl(r_Bill.ʵ�ս��, 0)
    Where ����id = r_Bill.����id And ���� = 1 And ���� = 2;
  
    If Sql%RowCount = 0 Then
      Insert Into ������� (����id, ����, ����, �������, Ԥ�����) Values (r_Bill.����id, 1, 2, r_Bill.ʵ�ս��, 0);
    End If;
  
    --����δ�����
    Update ����δ�����
    Set ��� = Nvl(���, 0) + Nvl(r_Bill.ʵ�ս��, 0)
    Where ����id = r_Bill.����id And Nvl(��ҳid, 0) = Nvl(r_Bill.��ҳid, 0) And Nvl(���˲���id, 0) = Nvl(r_Bill.���˲���id, 0) And
          Nvl(���˿���id, 0) = Nvl(r_Bill.���˿���id, 0) And Nvl(��������id, 0) = Nvl(r_Bill.��������id, 0) And
          Nvl(ִ�в���id, 0) = Nvl(r_Bill.ִ�в���id, 0) And ������Ŀid + 0 = r_Bill.������Ŀid And ��Դ;�� + 0 = r_Bill.�����־;
  
    If Sql%RowCount = 0 Then
      Insert Into ����δ�����
        (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
      Values
        (r_Bill.����id, r_Bill.��ҳid, r_Bill.���˲���id, r_Bill.���˿���id, r_Bill.��������id, r_Bill.ִ�в���id, r_Bill.������Ŀid,
         r_Bill.�����־, Nvl(r_Bill.ʵ�ս��, 0));
    End If;
  
    If r_Bill.ҽ����� Is Not Null Then
      v_ҽ��ids := v_ҽ��ids || ',' || r_Bill.ҽ�����;
    End If;
  End Loop;

  --����ҽ�����ͼƷ�״̬
  If v_ҽ��ids Is Not Null Then
    v_ҽ��ids := Substr(v_ҽ��ids, 2);
    Zl_ҽ������_�Ʒ�״̬_Update(1, 2, 1, No_In, v_ҽ��ids);
  End If;

  --�ⷿ�е�ҩƷ��ȫ��������Ϊ���շ�
  Update δ��ҩƷ��¼
  Set ���շ� = 1, �������� = v_Date
  Where NO = No_In And ���� In (9, 10) And Nvl(���շ�, 0) = 0 And
        Nvl(�ⷿid, 0) Not In
        (Select Distinct Nvl(ִ�в���id, 0)
         From סԺ���ü�¼
         Where ��¼���� = 2 And NO = No_In And �շ���� In ('5', '6', '7') And ��¼״̬ = 0);

  Update δ��ҩƷ��¼
  Set ���շ� = 1, �������� = v_Date
  Where NO = No_In And ���� In (25, 26) And Nvl(���շ�, 0) = 0 And
        Nvl(�ⷿid, 0) Not In (Select Distinct Nvl(ִ�в���id, 0)
                             From סԺ���ü�¼
                             Where ��¼���� = 2 And NO = No_In And �շ���� = '4' And ��¼״̬ = 0);

  n_�����Զ����� := To_Number(Nvl(zl_GetSysParameter(63), '0'));
  --0-���Զ����ϣ�1-�Զ����ϣ�2-�����ҿ���ʱ�Զ�����
  If Nvl(n_�����Զ�����, 0) <> 0 Then
  
    --����������������Զ�����
    For r_Stuff In c_Stuff Loop
      --1.���ϲ���ֱ����˵ĵ���;��ֱ���ϣ������ݲ�����鿪������
      --2.����Ƿ��ϲ�����˵��Ҳ���Ϊ�����ҿ���ʱ�Զ����ϵ�,�������ʱ��������������ⷿ��ͬʱ���ŷ���
      --3.���������Ϊ�Զ����ϣ��򲻼�鿪�����ţ�ֱ�ӷ���
      If Nvl(���ϲ������_In, 0) = 1 Or Nvl(n_�����Զ�����, 0) = 1 Or
         (Nvl(n_�����Զ�����, 0) = 2 And Nvl(n_��������id, 0) = Nvl(r_Stuff.�ⷿid, 0)) Then
      
        If v_���Ϻ� Is Null Then
          v_���Ϻ� := Nextno(20);
        End If;
      
        If r_Stuff.�ⷿid <> Nvl(v_�ⷿid, 0) Then
          If Nvl(v_�ⷿid, 0) <> 0 And v_�շ�ids Is Not Null Then
            v_�շ�ids := Substr(v_�շ�ids, 2);
            Zl_ҩƷ�շ���¼_��������(v_�շ�ids, v_�ⷿid, v_����Ա����, Sysdate, 1, v_����Ա����, v_���Ϻ�, v_����Ա����);
          End If;
        
          v_�ⷿid  := r_Stuff.�ⷿid;
          v_�շ�ids := Null;
        End If;
      
        v_�շ�ids := v_�շ�ids || '|' || r_Stuff.Id || ',0';
      End If;
    End Loop;
    If Nvl(v_�ⷿid, 0) <> 0 And v_�շ�ids Is Not Null Then
      v_�շ�ids := Substr(v_�շ�ids, 2);
      Zl_ҩƷ�շ���¼_��������(v_�շ�ids, v_�ⷿid, v_����Ա����, Sysdate, 1, v_����Ա����, v_���Ϻ�, v_����Ա����);
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_סԺ���ʼ�¼_Verify;
/

--133895:���ϴ�,2019-01-23,�˺Ž�����Ϣ���ⲿ���벻�ڹ����м���
Create Or Replace Procedure Zl_���˹Һż�¼_Delete
(
  ���ݺ�_In       ������ü�¼.No%Type,
  ����Ա���_In   ������ü�¼.����Ա���%Type,
  ����Ա����_In   ������ü�¼.����Ա����%Type,
  ժҪ_In         ������ü�¼.ժҪ%Type := Null, --ԤԼȡ��ʱ ��д ���ԤԼȡ��ԭ��
  ɾ�������_In   Number := 0,
  ��ԭ���˽���_In Varchar2 := Null,
  �˷�����_In     In Number := 0, --0-ȫ�� 1-�˹Һŷ� 2-�˲����� 3-�˸��ӷ� 4-�˹Һ��벡�� 5-�˹Һ��븽��
  ��ָ������_In   ����Ԥ����¼.���㷽ʽ%Type := Null,
  �˺�����_In     Number := 1,
  �ջ�Ʊ�ݺ�_In   Varchar2 := Null,
  ����˵��_In     ����Ԥ����¼.����˵��%Type := Null,
  ���㷽ʽ_In     Varchar2 := Null,
  ��Ԥ��_In       ����Ԥ����¼.��Ԥ��%Type := Null
) As
  --�˷�����_In,��һ�¼�������²�׼���в����˷�
  --    2.�����ӿ�,��ʱ��֧��
  -- �ҺŷѲ����ѷֿ���,����
  --    ��ͨ���㷽ʽ:ԭ���㷽ʽ�˲��ַ���
  --    Ԥ����:Ԥ����,�˲���
  --    Ԥ��������ͨ���㷽ʽ���:�˿����ͨ���㷽ʽ������
  --    ���ѿ�:ԭ�������ò����������ѿ�
  --��ԭ���˽���_In:ָ�����˻���ԭ�����㷽ʽ(��ҽ���ĸ����˻�,�����˻������ֵ�),����ö�����
  --��ָ������_IN:ָ��ԭ���˽��㲿��,Ӧ���˸����ֽ��㷽ʽ,Ϊ��ʱȱʡ�˸��ֽ�,�����˸�ָ���Ľ��㷽ʽ

  --���α������ж��Ƿ񵥶��ղ�����,���ҺŻ��ܱ���
  Cursor c_Registinfo(v_״̬ ������ü�¼.��¼״̬%Type) Is
    Select a.����ʱ��, a.�Ǽ�ʱ��, c.����ʱ��, a.�շ�ϸĿid As ��Ŀid, c.ִ�в���id As ����id, c.ִ���� As ҽ������, d.Id As ҽ��id, c.�ű� As ����
    From ������ü�¼ A, �ҺŰ��� B, ���˹Һż�¼ C, ��Ա�� D
    Where a.��¼���� = 4 And a.��� = 1 And a.��¼״̬ = v_״̬ And c.No = a.No And c.ִ���� = d.����(+) And a.No = ���ݺ�_In And
          Nvl(a.���㵥λ, '�ű�') = c.�ű� And Rownum < 2;
  r_Registrow c_Registinfo%RowType;

  --���α������жϼ�¼�Ƿ����,�����û��ܱ���
  Cursor c_Moneyinfo(v_״̬ ������ü�¼.��¼״̬%Type) Is
    Select ���˿���id, ��������id, ִ�в���id, ������Ŀid, Nvl(Sum(Ӧ�ս��), 0) As Ӧ��, Nvl(Sum(ʵ�ս��), 0) As ʵ��, Nvl(Sum(���ʽ��), 0) As ����
    From ������ü�¼
    Where ��¼���� = 4 And ��¼״̬ = v_״̬ And NO = ���ݺ�_In
    Group By ���˿���id, ��������id, ִ�в���id, ������Ŀid;
  r_Moneyrow c_Moneyinfo%RowType;

  --�ù�����ڴ�����Ա�ɿ�������˵Ĳ�ͬ���㷽ʽ�Ľ��
  Cursor c_Opermoney(n_Id ����Ԥ����¼.����id%Type) Is
    Select Distinct b.���㷽ʽ, -1 * Nvl(b.��Ԥ��, 0) As ��Ԥ��
    From ����Ԥ����¼ B
    Where b.����id = n_Id And b.��¼���� = 4 And b.��¼״̬ = 2 And Nvl(b.��Ԥ��, 0) <> 0;

  v_Err_Msg Varchar(255);
  Err_Item Exception;

  n_����id ����Ԥ����¼.����id%Type;
  n_����id ������ü�¼.����id%Type;

  v_��ָ�����㷽ʽ ����Ԥ����¼.���㷽ʽ%Type;
  n_�˿���       ����Ԥ����¼.��Ԥ��%Type;
  n_��ӡid         Ʊ�ݴ�ӡ����.Id%Type;
  n_����id         ������Ϣ.����id%Type;
  n_�˷ѽ��       ����Ԥ����¼.��Ԥ��%Type;
  n_Ԥ�����       ����Ԥ����¼.��Ԥ��%Type; --ԭ��¼ Ԥ���ɿ���
  n_����ֵ         �������.Ԥ�����%Type;
  n_�Һ�id         ���˹Һż�¼.Id%Type;
  n_��id           ����ɿ����.Id%Type;

  n_�����˷�       Number; --��¼�Ƿ��Ǵ˵��ݵĵڶ����˷�
  n_����̨ǩ���Ŷ� Number;
  n_ԤԼ���ɶ���   Number;
  n_ԤԼ�Һ�       Number;
  n_�Һ����ɶ���   Number;
  d_Date           Date;
  n_����           ������ü�¼.���ʷ���%Type;
  n_����id1        ������Ϣ.����id%Type;
  n_���ض�         ������ü�¼.ʵ�ս��%Type;
  n_�ѽ���         Number;
  n_����id         �ҺŰ���.Id%Type;
  n_�ƻ�id         �ҺŰ��żƻ�.Id%Type;
  d_����ʱ��       Date;
  d_����ʱ��       ���˹Һż�¼.����ʱ��%Type;
  n_�����¼id     �ٴ������¼.Id%Type;
  v_����           �ҺŰ���.����%Type;
  n_���           ���˹Һż�¼.����%Type;
  v_ʱ���         ʱ���.ʱ���%Type;
  d_��鿪ʼʱ��   Date;
  d_������ʱ��   Date;
  v_Temp           Varchar2(500);
  v_����ids        Varchar2(500);
  n_���ﲡ��id     ������Ϣ.����id%Type;
  d_����ʱ��       ����ǼǼ�¼.����ʱ��%Type;
  v_��������       Varchar2(5000);
  v_��ǰ����       Varchar2(1000);
  v_���㷽ʽ       ����Ԥ����¼.���㷽ʽ%Type;
  n_��������־     Number;
  n_������       ����Ԥ����¼.��Ԥ��%Type;
  n_Count          Number;
Begin
  n_��id           := Zl_Get��id(����Ա����_In);
  v_��ָ�����㷽ʽ := ��ָ������_In;

  --�����ж�Ҫ�˺�/ȡ��ԤԼ�ļ�¼�Ƿ����
  Open c_Moneyinfo(1);
  Fetch c_Moneyinfo
    Into r_Moneyrow;
  If c_Moneyinfo%RowCount = 0 Then
    Close c_Moneyinfo;
    Open c_Moneyinfo(0);
    Fetch c_Moneyinfo
      Into r_Moneyrow;
    If c_Moneyinfo%RowCount = 0 Then
      v_Err_Msg := 'Ҫ����ĵ��ݲ����ڡ�';
      Raise Err_Item;
    End If;
    n_ԤԼ�Һ� := 1;
  End If;
  Close c_Moneyinfo;

  v_Temp := zl_GetSysParameter(256);
  If v_Temp Is Null Or Substr(v_Temp, 1, 1) = '0' Then
    Null;
  Else
    Begin
      d_����ʱ�� := To_Date(Substr(v_Temp, 3), 'YYYY-MM-DD hh24:mi:ss');
    Exception
      When Others Then
        d_����ʱ�� := Null;
    End;
  End If;

  Select �ű�, ����, ����ʱ�� Into v_����, n_���, d_����ʱ�� From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum < 2;

  Begin
    Select a.Id Into n_����id From �ҺŰ��� A Where a.���� = v_����;
  Exception
    When Others Then
      n_����id := -1;
  End;

  Begin
    Select ID
    Into n_�ƻ�id
    From �ҺŰ��żƻ�
    Where ����id = n_����id And ���ʱ�� Is Not Null And
          Nvl(��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) =
          (Select Max(a.��Чʱ��) As ��Ч
           From �ҺŰ��żƻ� A
           Where a.���ʱ�� Is Not Null And d_����ʱ�� Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                 a.ʧЧʱ�� And a.����id = n_����id) And
          d_����ʱ�� Between Nvl(��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And ʧЧʱ��;
  Exception
    When Others Then
      n_�ƻ�id := 0;
  End;

  Begin
    If Nvl(n_�ƻ�id, 0) = 0 Then
      Select Decode(To_Char(d_����ʱ��, 'D'), '1', ����, '2', ��һ, '3', �ܶ�, '4', ����, '5', ����, '6', ����, '7', ����, Null)
      Into v_ʱ���
      From �ҺŰ���
      Where ID = n_����id;
    Else
      Select Decode(To_Char(d_����ʱ��, 'D'), '1', ����, '2', ��һ, '3', �ܶ�, '4', ����, '5', ����, '6', ����, '7', ����, Null)
      Into v_ʱ���
      From �ҺŰ��żƻ�
      Where ID = n_�ƻ�id;
    End If;
  Exception
    When Others Then
      v_ʱ��� := Null;
  End;

  If v_ʱ��� Is Not Null And d_����ʱ�� Is Not Null Then
    --����Ƿ��ģʽ�ҺŰ���
    Select To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
           To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || ' ' || To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
    Into d_��鿪ʼʱ��, d_������ʱ��
    From ʱ���
    Where ʱ��� = v_ʱ��� And վ�� Is Null And ���� Is Null;
    If d_��鿪ʼʱ�� > d_������ʱ�� Then
      d_������ʱ�� := d_������ʱ�� + 1;
    End If;
    If d_������ʱ�� > d_����ʱ�� Then
      --��ȡ�����¼id
      Begin
        Select a.Id
        Into n_�����¼id
        From �ٴ������¼ A, �ٴ������Դ B
        Where a.��Դid = b.Id And b.���� = v_���� And �ϰ�ʱ�� = v_ʱ��� And d_����ʱ�� Between ��ʼʱ�� And ��ֹʱ��;
      Exception
        When Others Then
          n_�����¼id := Null;
      End;
    End If;
  End If;

  Begin
    Select Zl_Fun_Regcustomname Into v_Temp From Dual;
    If v_Temp Is Not Null Then
      v_����ids := Substr(v_Temp, Instr(v_Temp, '|') + 1);
    End If;
  Exception
    When Others Then
      v_����ids := Null;
  End;

  --1.ԤԼ����
  If Nvl(n_ԤԼ�Һ�, 0) = 1 Then
    --������Լ��
    Open c_Registinfo(0);
    Fetch c_Registinfo
      Into r_Registrow;
  
    Update ���˹ҺŻ���
    Set ��Լ�� = Nvl(��Լ��, 0) - 1
    Where ���� = Trunc(r_Registrow.����ʱ��) And ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And
          (���� = r_Registrow.���� Or ���� Is Null);
  
    If Sql%RowCount = 0 Then
      Insert Into ���˹ҺŻ���
        (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, ��Լ��)
      Values
        (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
         Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -1);
    End If;
    Close c_Registinfo;
  
    --���¹Һ����״̬
    Delete �Һ����״̬
    Where ״̬ = 2 And
          (����, ���, ����) = (Select ���㵥λ, ��ҩ����, Trunc(����ʱ��)
                          From ������ü�¼
                          Where ��¼���� = 4 And ��¼״̬ = 0 And ��� = 1 And Rownum = 1 And NO = ���ݺ�_In) Or
          (����, ���, ����) = (Select ���㵥λ, ��ҩ����, ����ʱ��
                          From ������ü�¼
                          Where ��¼���� = 4 And ��¼״̬ = 0 And ��� = 1 And Rownum = 1 And NO = ���ݺ�_In);
  
    --��Ӳ��˹Һż�¼�� ������¼
    Select ���˹Һż�¼_Id.Nextval, Sysdate Into n_�Һ�id, d_Date From Dual;
  
    Update ���˹Һż�¼ Set ��¼״̬ = 3 Where NO = ���ݺ�_In And ��¼״̬ = 1 And ��¼���� = 2;
    If Sql%NotFound Then
      v_Err_Msg := 'ԤԼ����' || ���ݺ�_In || '�������ڻ����ڲ���ԭ���Ѿ���ȡ��ԤԼ';
      Raise Err_Item;
    End If;
  
    Insert Into ���˹Һż�¼
      (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
       ����, ����, ����, ԤԼ, ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ҽ�Ƹ��ʽ)
      Select n_�Һ�id, NO, ��¼����, 2, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, d_Date, ����ʱ��,
             ����Ա���_In, ����Ա����_In, ����, ����, ����, ԤԼ, Nvl(ժҪ_In, ժҪ) As ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ҽ�Ƹ��ʽ
      From ���˹Һż�¼
      Where NO = ���ݺ�_In;
  
    If n_�����¼id Is Not Null Then
      Update �ٴ������¼ Set ��Լ�� = ��Լ�� - 1 Where ID = n_�����¼id And Nvl(��Լ��, 0) > 0;
      Update �ٴ�������ſ��� Set �Һ�״̬ = Null, ����Ա���� = Null Where ��¼id = n_�����¼id And ��� = n_���;
    End If;
  
    --Update ���˹Һż�¼ set ժҪ=nvl(ժҪ_IN,ժҪ) where NO=���ݺ�_IN;
    --ɾ��������ü�¼
    Delete From ������ü�¼ Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 0;
    --���ԤԼ���ɶ���ʱ��Ҫ�������
  
    n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
    If Nvl(n_ԤԼ���ɶ���, 0) = 1 Then
      --Ҫɾ������
      For v_�Һ� In (Select ID, ����, ִ�в���id, ִ���� From ���˹Һż�¼ Where NO = ���ݺ�_In) Loop
        Zl_�ŶӽкŶ���_Delete(v_�Һ�.ִ�в���id, v_�Һ�.Id);
      End Loop;
    End If;
    Return;
  End If;

  Select Nvl(���ʷ���, 0), ����id, Decode(Sign(Nvl(����id, 0)), 0, 0, 1)
  Into n_����, n_����id, n_�ѽ���
  From ������ü�¼
  Where ��¼���� = 4 And NO = ���ݺ�_In And ��¼״̬ In (1, 3) And Rownum < 2;

  --2.�ҺŴ���
  n_�ѽ��� := Nvl(n_�ѽ���, 0);

  If n_�ѽ��� = 1 And n_���� = 1 Then
    Select Sysdate, Null Into d_Date, n_����id From Dual;
  Else
    Select Sysdate, ���˽��ʼ�¼_Id.Nextval Into d_Date, n_����id From Dual;
  End If;

  ----0-ȫ�� 1-�˹Һŷ� 2-�˲�����
  If Nvl(�˷�����_In, 0) Not In (2, 3) Then
    --���ǹ��˲�����ʱ����
    --���¹Һ����״̬
    If �˺�����_In = 1 Then
      Delete �Һ����״̬
      Where ״̬ = 1 And
            (����, ���, ����) = (Select �ű�, ����, Trunc(����ʱ��) From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum = 1) Or
            (����, ���, ����) = (Select �ű�, ����, ����ʱ�� From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum = 1);
    Else
      Update �Һ����״̬
      Set ״̬ = 4
      Where ״̬ = 1 And
            (����, ���, ����) = (Select �ű�, ����, Trunc(����ʱ��) From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum = 1) Or
            (����, ���, ����) = (Select �ű�, ����, ����ʱ�� From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum = 1);
    End If;
  
    --���˾���״̬
    If n_����id Is Not Null Then
      Update ������Ϣ Set ����״̬ = 0, �������� = Null Where ����id = n_����id;
    
      --ɾ���������ش���,ֻ�е�ֻ��һ���Һż�¼���Ҳ��˽���������Һ����ڽ���ʱ�Żᴦ��
      If ɾ�������_In = 1 Then
        Delete ���ﲡ����¼ Where ����id = n_����id;
        Update ������Ϣ Set ����� = Null Where ����id = n_����id;
        --���ü�¼�����Һż����������￨����,�Լ����˽��Ѻ��˷ѻ����ʵķ���,�Һż�¼�������
        Update ������ü�¼ Set ��ʶ�� = Null Where �����־ = 1 And ����id = n_����id;
      End If;
    End If;
  
    --�����ʱ���˾��￨��,�˷�ʱ������￨��,�ڷǹ��˲�����ʱ
    n_����id1 := Null;
    Begin
      Select ����id
      Into n_����id1
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And ���ӱ�־ = 2 And Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
  
    If n_����id1 Is Not Null And Nvl(�˷�����_In, 0) Not In (2, 3) Then
      Update ������Ϣ
      Set ���￨�� = Null, ����֤�� = Null, Ic���� = Decode(Ic����, ���￨��, Null, Ic����)
      Where ����id = n_����id1;
    End If;
  
  End If;

  --���ǰ���Ƿ��Ѿ������˹�����
  Begin
    Select 1 Into n_�����˷� From ������ü�¼ Where ��¼���� = 4 And NO = ���ݺ�_In And ��¼״̬ = 3 And Rownum < 2;
  Exception
    When Others Then
      n_�����˷� := 0;
  End;

  If Nvl(�˷�����_In, 0) = 0 Or Nvl(�˷�����_In, 0) = 2 Then
    --ȫ��,�˲�����
    --������ü�¼��������¼
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
       ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��,
       ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
             ����Ա���_In, ����Ա����_In, ����ʱ��, d_Date, n_����id,
             Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����,
             Nvl(ժҪ_In, ժҪ) As ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
  
    --ԭʼ��¼
    If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
      Update ������ü�¼
      Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    Else
      Update ������ü�¼
      Set ��¼״̬ = 3
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    End If;
  Elsif Nvl(�˷�����_In, 0) = 1 Or Nvl(�˷�����_In, 0) = 4 Then
    --�˹Һŷ�,�˹Һ��벡����
    --������ü�¼��������¼
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
       ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��,
       ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
             ����Ա���_In, ����Ա����_In, ����ʱ��, d_Date, n_����id,
             Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����,
             Nvl(ժҪ_In, ժҪ) As ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') = 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
  
    --ԭʼ��¼
    If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
      Update ������ü�¼
      Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') = 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    Else
      Update ������ü�¼
      Set ��¼״̬ = 3
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') = 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    End If;
  Elsif Nvl(�˷�����_In, 0) = 3 Then
    --�˸��ӷ�
    --������ü�¼��������¼
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
       ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��,
       ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
             ����Ա���_In, ����Ա����_In, ����ʱ��, d_Date, n_����id,
             Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����,
             Nvl(ժҪ_In, ժҪ) As ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') > 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
  
    --ԭʼ��¼
    If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
      Update ������ü�¼
      Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') > 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    Else
      Update ������ü�¼
      Set ��¼״̬ = 3
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') > 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    End If;
  Elsif Nvl(�˷�����_In, 0) = 5 Then
    --�˹Һ��븽�ӷ�
    --������ü�¼��������¼
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
       ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��,
       ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
             ����Ա���_In, ����Ա����_In, ����ʱ��, d_Date, n_����id,
             Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����,
             Nvl(ժҪ_In, ժҪ) As ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Nvl(���ӱ�־, 0) <> 1;
  
    --ԭʼ��¼
    If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
      Update ������ü�¼
      Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Nvl(���ӱ�־, 0) <> 1;
    Else
      Update ������ü�¼
      Set ��¼״̬ = 3
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Nvl(���ӱ�־, 0) <> 1;
    End If;
  End If;

  n_����id := 0;
  If n_���� = 0 Then
    --��ȡ����ID
    Select Nvl(����id, 0)
    Into n_����id
    From ������ü�¼
    Where ��¼���� = 4 And ��¼״̬ = 3 And NO = ���ݺ�_In And Rownum < 2;
  End If;

  If n_���� = 1 Then
    --����
    For c_���� In (Select ʵ�ս��, ���˿���id, ��������id, ִ�в���id, ������Ŀid
                 From ������ü�¼
                 Where ��¼���� = 4 And ��¼״̬ = 3 And NO = ���ݺ�_In And
                       Nvl(���ӱ�־, 0) =
                       Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0)) And
                       Nvl(���ʷ���, 0) = 1) Loop
      --�������
      Update �������
      Set ������� = Nvl(�������, 0) - Nvl(c_����.ʵ�ս��, 0)
      Where ����id = Nvl(n_����id, 0) And ���� = 1 And ���� = 1
      Returning ������� Into n_���ض�;
    
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, ����, �������, Ԥ�����)
        Values
          (n_����id, 1, 1, -1 * Nvl(c_����.ʵ�ս��, 0), 0);
        n_���ض� := Nvl(c_����.ʵ�ս��, 0);
      End If;
      If Nvl(n_���ض�, 0) = 0 Then
        Delete �������
        Where ����id = Nvl(n_����id, 0) And ���� = 1 And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
      --����δ�����
      Update ����δ�����
      Set ��� = Nvl(���, 0) - Nvl(c_����.ʵ�ս��, 0)
      Where ����id = n_����id And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(c_����.���˿���id, 0) And
            Nvl(��������id, 0) = Nvl(c_����.��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(c_����.ִ�в���id, 0) And ������Ŀid + 0 = c_����.������Ŀid And
            ��Դ;�� + 0 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into ����δ�����
          (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
        Values
          (n_����id, Null, Null, c_����.���˿���id, c_����.��������id, c_����.ִ�в���id, c_����.������Ŀid, 1, -1 * Nvl(c_����.ʵ�ս��, 0));
      End If;
    End Loop;
    Delete ����δ�����
    Where ����id = n_����id And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���, 0) = 0 And ��Դ;�� + 0 = 1;
  End If;

  If n_���� = 0 Then
    --1.�˷�
    --���˹ҺŽ���:�ֽ�͸����ʻ�����
    If ���㷽ʽ_In Is Null And Nvl(��Ԥ��_In, 0) = 0 Then
      If ��ԭ���˽���_In Is Not Null Then
        --�˿����ȡ
        If Nvl(�˷�����_In, 0) <> 0 Or Nvl(n_�����˷�, 0) <> 0 Then
          --����ǵ����˲�����,����ֻ�˹Һŷ�,�Ȼ�ȡ�˷ѽ��
          Begin
            --��ȡ�����˿���
            Select -1 * Sum(Nvl(ʵ�ս��, 0)) As �տ���
            Into n_�˿���
            From ������ü�¼
            Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 2 And ����id = n_����id;
          
          Exception
            When Others Then
              v_Err_Msg := '���ݡ�' || ���ݺ�_In || '����' || Case Nvl(�˷�����_In, 0)
                             When 1 Then
                              '�Һŷ���'
                             When 2 Then
                              '������'
                           End || '�������ڲ���ԭ���Ѿ��������˷ѻ��ߵ��ݲ�����!';
              Raise Err_Item;
          End;
          Begin
            Select ��Ԥ��
            Into n_�˷ѽ��
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
          Exception
            When Others Then
              n_�˷ѽ�� := 0;
          End;
        
          --a.����Ľ��㷽ʽ
        
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In, -n_�˿���,
                   n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
        
          If n_�˷ѽ�� = 0 Then
            --b.����������ֽ�
            If n_�˿��� <> 0 Then
              If v_��ָ�����㷽ʽ Is Null Then
                --�˸��ֽ�
                Begin
                  Select ���� Into v_��ָ�����㷽ʽ From ���㷽ʽ Where ���� = 1;
                Exception
                  When Others Then
                    v_��ָ�����㷽ʽ := '�ֽ�';
                End;
              End If;
              Update ����Ԥ����¼
              Set ��Ԥ�� = ��Ԥ�� + (-1 * n_�˿���), ����˵�� = Nvl(����˵��_In, ����˵��)
              Where ��¼���� = 4 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_��ָ�����㷽ʽ;
              If Sql%RowCount = 0 Then
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����,
                   �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
                  Select ����Ԥ����¼_Id.Nextval, a.No, a.ʵ��Ʊ��, a.��¼����, 2, a.����id, a.��ҳid, a.����id, a.ժҪ, v_��ָ�����㷽ʽ, d_Date,
                         ����Ա���_In, ����Ա����_In, -1 * n_�˿���, n_����id, n_��id, Ԥ�����, Decode(����˵��_In, Null, �����id, Null), ���㿨���,
                         Decode(����˵��_In, Null, ����, Null), Decode(����˵��_In, Null, ������ˮ��, Null), Nvl(����˵��_In, ����˵��), ������λ, 4
                  From ����Ԥ����¼ A
                  Where a.��¼���� = 4 And a.��¼״̬ = 1 And a.����id = n_����id And Rownum < 2;
              End If;
            End If;
          End If;
        Else
          --a.����Ľ��㷽ʽԭ����
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In, -��Ԥ��,
                   n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
        
          --b.����������ֽ�
          Begin
            Select Sum(��Ԥ��)
            Into n_�˷ѽ��
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') > 0;
          Exception
            When Others Then
              n_�˷ѽ�� := 0;
          End;
          If n_�˷ѽ�� <> 0 Then
            If v_��ָ�����㷽ʽ Is Null Then
              --�˸��ֽ�
              Begin
                Select ���㷽ʽ
                Into v_��ָ�����㷽ʽ
                From ����Ԥ����¼
                Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
              
              Exception
                When Others Then
                  Begin
                    Select ���� Into v_��ָ�����㷽ʽ From ���㷽ʽ Where ���� = 1;
                  Exception
                    When Others Then
                      v_��ָ�����㷽ʽ := '�ֽ�';
                  End;
              End;
            End If;
            Update ����Ԥ����¼
            Set ��Ԥ�� = ��Ԥ�� + (-1 * n_�˷ѽ��), ����˵�� = Nvl(����˵��_In, ����˵��)
            Where ��¼���� = 4 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_��ָ�����㷽ʽ;
            If Sql%RowCount = 0 Then
              Insert Into ����Ԥ����¼
                (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
                 ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
                Select ����Ԥ����¼_Id.Nextval, a.No, a.ʵ��Ʊ��, a.��¼����, 2, a.����id, a.��ҳid, a.����id, a.ժҪ, v_��ָ�����㷽ʽ, d_Date,
                       ����Ա���_In, ����Ա����_In, -1 * n_�˷ѽ��, n_����id, n_��id, Ԥ�����, Decode(����˵��_In, Null, �����id, Null), ���㿨���,
                       Decode(����˵��_In, Null, ����, Null), Decode(����˵��_In, Null, ������ˮ��, Null), Nvl(����˵��_In, ����˵��), ������λ, 4
                From ����Ԥ����¼ A
                Where a.��¼���� = 4 And a.��¼״̬ = 1 And a.����id = n_����id And Rownum < 2;
            End If;
          End If;
        End If;
      Else
        --�˿����ȡ
        If Nvl(�˷�����_In, 0) <> 0 Or Nvl(n_�����˷�, 0) <> 0 Then
          --����ǵ����˲�����,����ֻ�˹Һŷ�,�Ȼ�ȡ�˷ѽ��
          Begin
            --��ȡ�����˿���
            Select -1 * Sum(Nvl(ʵ�ս��, 0)) As �տ���
            Into n_�˿���
            From ������ü�¼
            Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 2 And ����id = n_����id;
          Exception
            When Others Then
              v_Err_Msg := '���ݡ�' || ���ݺ�_In || '����' || Case Nvl(�˷�����_In, 0)
                             When 1 Then
                              '�Һŷ���'
                             When 2 Then
                              '������'
                           End || '�������ڲ���ԭ���Ѿ��������˷ѻ��ߵ��ݲ�����!';
              Raise Err_Item;
          End;
        End If;
        If Nvl(n_�����˷�, 0) = 0 And Nvl(�˷�����_In, 0) = 0 Then
          --�״�ȫ��
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In, -1 * ��Ԥ��,
                   n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id;
        Else
          --�����˷�,���߱��ε���һ����
          --�����˷�ʱ,��¼״̬=3 ,�״β�����,��¼״̬Ϊ1
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                   -1 * n_�˿���, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And ժҪ = 'ҽ���Һ�' And ��Ԥ�� = n_�˿��� And
                  Rownum < 2;
          If Sql%RowCount = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
               ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
              Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                     -1 * n_�˿���, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
              From ����Ԥ����¼
              Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And ��Ԥ�� = n_�˿��� And Rownum < 2;
          End If;
          If Sql%RowCount = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
               ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
              Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                     -1 * n_�˿���, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
              From ����Ԥ����¼
              Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And Rownum < 2;
            If Sql%RowCount = 0 Then
              --�����˷�,����ȫ��ʹ��Ԥ����ɷ�ʱ�Ŵ��ڴ������
              n_Ԥ����� := n_�˿���;
            End If;
          End If;
        End If;
      End If;
    Else
      --�����㷽ʽ��
      If ���㷽ʽ_In is Not Null then
         v_�������� := ���㷽ʽ_In || '|'; --�Կո�ֿ���|��β,û�н�������
         While v_�������� Is Not Null Loop
          v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '|') - 1);
          v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
        
          v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
          n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
        
          v_��ǰ����   := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
          n_��������־ := To_Number(v_��ǰ����);
        
          If n_��������־ = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
               ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
              Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, v_���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                     -1 * n_������, n_����id, n_��id, Ԥ�����, Null, Null, Null, Null, ����˵��_In, ������λ, 4
              From ����Ԥ����¼
              Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And Rownum < 2;
          Else
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
               ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
              Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, v_���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                     -1 * n_������, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, Nvl(����˵��_In, ����˵��), ������λ, 4
              
              From ����Ԥ����¼
              Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And
                    (�����id Is Not Null Or ���㿨��� Is Not Null) And Rownum < 2;
          End If;
          v_�������� := Substr(v_��������, Instr(v_��������, '|') + 1);
        End Loop;
      End IF;
      n_Ԥ����� := Nvl(��Ԥ��_In, 0);
    End if;
    --�״��˷�ʱ,��¼״̬�����Ϊ��3
    Update ����Ԥ����¼
    Set ��¼״̬ = 3
    Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id;
  
    --��Ԥ�� 1-ȫ�� 2-������,������ʱ��ȫ��ʹ��Ԥ�����нɿ�
    If Nvl(�˷�����_In, 0) = 0 Or (Nvl(�˷�����_In, 0) <> 0 And n_Ԥ����� <> 0) Then
      --���˹ҺŽ���:��Ԥ�����
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
        Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, d_Date,
               ����Ա����_In, ����Ա���_In, -1 * Decode(Nvl(�˷�����_In, 0) + Nvl(n_�����˷�, 0), 0, ��Ԥ��, n_Ԥ�����), n_����id, n_��id, Ԥ�����,
               �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
        From ����Ԥ����¼
        Where ��¼���� In (1, 11) And ����id = n_����id And Nvl(��Ԥ��, 0) <> 0 And
              Rownum = Decode(Nvl(�˷�����_In, 0) + Nvl(n_�����˷�, 0), 0, Rownum, 1);
    End If;
  
    --������Ԥ�����
    For c_Ԥ�� In (Select ����id, Ԥ�����, -1 * Sum(Nvl(��Ԥ��, 0)) As ��Ԥ��
                 From ����Ԥ����¼
                 Where ��¼���� In (1, 11) And ����id = n_����id
                 Group By ����id, Ԥ�����) Loop
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(c_Ԥ��.��Ԥ��, 0)
      Where ����id = c_Ԥ��.����id And ���� = Nvl(c_Ԥ��.Ԥ�����, 2) And ���� = 1
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, Ԥ�����, ����, ����)
        Values
          (c_Ԥ��.����id, Nvl(c_Ԥ��.��Ԥ��, 0), 1, Nvl(c_Ԥ��.Ԥ�����, 2));
        n_����ֵ := Nvl(c_Ԥ��.��Ԥ��, 0);
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From �������
        Where ����id = c_Ԥ��.����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
      End If;
    End Loop;
  
    If �ջ�Ʊ�ݺ�_In Is Not Null Then
      --���˹Һŷ�,������Ʊ��
      --�˿��ջ�Ʊ��(�����ϴιҺ�ʹ��Ʊ��,�����ջ�)
      Begin
        --�����һ�δ�ӡ��������ȡ
        Select ID
        Into n_��ӡid
        From (Select b.Id
               From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
               Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And b.�������� = 4 And b.No = ���ݺ�_In
               Order By a.ʹ��ʱ�� Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          n_��ӡid := Null;
      End;
    
      --���ջ�ԭƱ��
      If n_��ӡid Is Not Null Then
        Begin
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
            Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, ����Ա����_In, Ʊ�ݽ��
            From Ʊ��ʹ����ϸ
            Where ��ӡid = n_��ӡid And ���� = 1 And Instr(',' || �ջ�Ʊ�ݺ�_In || ',', ',' || ���� || ',') > 0;
        Exception
          When Others Then
            Delete From Ʊ��ʹ����ϸ Where ��ӡid = n_��ӡid And ���� = 2 And ԭ�� = 2;
            Insert Into Ʊ��ʹ����ϸ
              (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
              Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, ����Ա����_In, Ʊ�ݽ��
              From Ʊ��ʹ����ϸ
              Where ��ӡid = n_��ӡid And ���� = 1 And Instr(',' || �ջ�Ʊ�ݺ�_In || ',', ',' || ���� || ',') > 0;
        End;
      End If;
    End If;
  End If;

  --�����˲�������,��������ܼ�¼
  --��ػ��ܱ�Ĵ���

  --���˹ҺŻ���
  If Nvl(�˷�����_In, 0) Not In (2, 3) Then
    Open c_Registinfo(3);
    Fetch c_Registinfo
      Into r_Registrow;
  
    If c_Registinfo%RowCount = 0 Then
      --ֻ�ղ�����ʱ�޺ű�,������
      Close c_Registinfo;
    Else
    
      --��Ҫȷ���Ƿ�ԤԼ�Һ�
      --1.�������ԤԼ�ҺŲ����ĹҺż�¼,����Ҫ���ѹ����������ѽ���
      --2.����������Һ�,��ֻ���ѹ���
    
      Begin
        Select Decode(ԤԼ, Null, 0, 0, 0, 1) Into n_ԤԼ�Һ� From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum = 1;
      Exception
        When Others Then
          n_ԤԼ�Һ� := 0;
      End;
    
      Update ���˹ҺŻ���
      Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - n_ԤԼ�Һ�, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ�Һ�
      Where ���� = Trunc(r_Registrow.����ʱ��) And ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And
            (���� = r_Registrow.���� Or ���� Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into ���˹ҺŻ���
          (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, �����ѽ���, ��Լ��)
        Values
          (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
           Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -1, -1 * n_ԤԼ�Һ�, -1 * n_ԤԼ�Һ�);
      End If;
    
      If n_�����¼id Is Not Null Then
        Update �ٴ������¼
        Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - n_ԤԼ�Һ�, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ�Һ�
        Where ID = n_�����¼id And Nvl(��Լ��, 0) > 0;
        Update �ٴ�������ſ��� Set �Һ�״̬ = Null, ����Ա���� = Null Where ��¼id = n_�����¼id And ��� = n_���;
      End If;
    
      Close c_Registinfo;
    End If;
  End If;

  If n_���� = 0 Then
    --��Ա�ɿ����(���������ʻ��ȵĽ�����,�����˳�Ԥ����)
    For r_Opermoney In c_Opermoney(n_����id) Loop
      Update ��Ա�ɿ����
      Set ��� = Nvl(���, 0) + (-1 * r_Opermoney.��Ԥ��)
      Where �տ�Ա = ����Ա����_In And ���㷽ʽ = r_Opermoney.���㷽ʽ And ���� = 1
      Returning ��� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ��Ա�ɿ����
          (�տ�Ա, ���㷽ʽ, ����, ���)
        Values
          (����Ա����_In, r_Opermoney.���㷽ʽ, 1, -1 * r_Opermoney.��Ԥ��);
        n_����ֵ := r_Opermoney.��Ԥ��;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From ��Ա�ɿ����
        Where �տ�Ա = ����Ա����_In And ���㷽ʽ = r_Opermoney.���㷽ʽ And ���� = 1 And Nvl(���, 0) = 0;
      End If;
    End Loop;
  End If;
  If Nvl(�˷�����_In, 0) Not In (2, 3) Then
    n_�Һ����ɶ��� := Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113));
    If n_�Һ����ɶ��� <> 0 Then
      --Ҫɾ������
      For v_�Һ� In (Select ID, ����, ִ�в���id, ִ���� From ���˹Һż�¼ Where NO = ���ݺ�_In) Loop
        n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113, 1, Nvl(v_�Һ�.ִ�в���id, 0)));
        If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Then
          Zl_�ŶӽкŶ���_Delete(v_�Һ�.ִ�в���id, v_�Һ�.Id);
        End If;
      End Loop;
    End If;
  
    --ҽ�������ľ���ǼǼ�¼
    Begin
      Select ����id, ����ʱ�� Into n_���ﲡ��id, d_����ʱ�� From ���˹Һż�¼ Where NO = ���ݺ�_In;
      Delete From ����ǼǼ�¼ Where ����id = n_���ﲡ��id And ����ʱ�� = d_����ʱ�� And ��ҳid Is Null;
    Exception
      When Others Then
        Null;
    End;
  
    --���˹Һż�¼
    Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
  
    Update ���˹Һż�¼ Set ��¼״̬ = 3 Where NO = ���ݺ�_In And ��¼���� = 1 And ��¼״̬ = 1;
    If Sql%NotFound Then
      v_Err_Msg := '�Һŵ���' || ���ݺ�_In || '�������ڻ����ڲ���ԭ���Ѿ����˺�';
      Raise Err_Item;
    End If;
    Insert Into ���˹Һż�¼
      (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
       ����, ����, ����, ԤԼ, ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ����, ҽ�Ƹ��ʽ)
      Select n_�Һ�id, NO, ��¼����, 2, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, d_Date, ����ʱ��,
             ����Ա���_In, ����Ա����_In, ����, ����, ����, ԤԼ, Nvl(ժҪ_In, ժҪ) As ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ����, ҽ�Ƹ��ʽ
      From ���˹Һż�¼
      Where NO = ���ݺ�_In;
  End If;
  --��Ϣ����
  Begin
    Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
      Using 2, ���ݺ�_In;
  Exception
    When Others Then
      Null;
  End;
  b_Message.Zlhis_Regist_003(n_�Һ�id, ���ݺ�_In);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˹Һż�¼_Delete;
/

--133895:���ϴ�,2019-01-23,�˺Ž�����Ϣ���ⲿ���벻�ڹ����м���
Create Or Replace Procedure Zl_���˹Һż�¼_����_Delete
(
  ���ݺ�_In       ������ü�¼.No%Type,
  ����Ա���_In   ������ü�¼.����Ա���%Type,
  ����Ա����_In   ������ü�¼.����Ա����%Type,
  ժҪ_In         ������ü�¼.ժҪ%Type := Null, --ԤԼȡ��ʱ ��д ���ԤԼȡ��ԭ��
  ɾ�������_In   Number := 0,
  ��ԭ���˽���_In Varchar2 := Null,
  �˷�����_In     In Number := 0, --0-ȫ�� 1-�˹Һŷ� 2-�˲�����
  ��ָ������_In   ����Ԥ����¼.���㷽ʽ%Type := Null,
  �˺�����_In     Number := 1,
  ���㷽ʽ_In     Varchar2 := Null,
  ��Ԥ��_In       ����Ԥ����¼.��Ԥ��%Type := Null,
  �ջ�Ʊ�ݺ�_In   Varchar2 := Null,
  ����˵��_In     ����Ԥ����¼.����˵��%Type := Null
) As
  --�˷�����_In,��һ�¼�������²�׼���в����˷�
  --    2.�����ӿ�,��ʱ��֧��
  -- �ҺŷѲ����ѷֿ���,����
  --    ��ͨ���㷽ʽ:ԭ���㷽ʽ�˲��ַ���
  --    Ԥ����:Ԥ����,�˲���
  --    Ԥ��������ͨ���㷽ʽ���:�˿����ͨ���㷽ʽ������
  --    ���ѿ�:ԭ�������ò����������ѿ�
  --��ԭ���˽���_In:ָ�����˻���ԭ�����㷽ʽ(��ҽ���ĸ����˻�,�����˻������ֵ�),����ö�����
  --��ָ������_IN:ָ��ԭ���˽��㲿��,Ӧ���˸����ֽ��㷽ʽ,Ϊ��ʱȱʡ�˸��ֽ�,�����˸�ָ���Ľ��㷽ʽ

  --���α������ж��Ƿ񵥶��ղ�����,���ҺŻ��ܱ���
  Cursor c_Registinfo(v_״̬ ������ü�¼.��¼״̬%Type) Is
    Select a.����ʱ��, a.�Ǽ�ʱ��, c.����ʱ��, a.�շ�ϸĿid As ��Ŀid, c.ִ�в���id As ����id, c.ִ���� As ҽ������, d.Id As ҽ��id, c.�ű� As ����
    From ������ü�¼ A, ���˹Һż�¼ C, ��Ա�� D
    Where a.��¼���� = 4 And a.No = ���ݺ�_In And a.No = c.No And a.��¼״̬ = v_״̬ And c.ִ���� = d.����(+) And Rownum < 2;
  r_Registrow c_Registinfo%RowType;

  --���α������жϼ�¼�Ƿ����,�����û��ܱ���
  Cursor c_Moneyinfo(v_״̬ ������ü�¼.��¼״̬%Type) Is
    Select ���˿���id, ��������id, ִ�в���id, ������Ŀid, Nvl(Sum(Ӧ�ս��), 0) As Ӧ��, Nvl(Sum(ʵ�ս��), 0) As ʵ��, Nvl(Sum(���ʽ��), 0) As ����
    From ������ü�¼
    Where ��¼���� = 4 And ��¼״̬ = v_״̬ And NO = ���ݺ�_In
    Group By ���˿���id, ��������id, ִ�в���id, ������Ŀid;
  r_Moneyrow c_Moneyinfo%RowType;

  --�ù�����ڴ�����Ա�ɿ�������˵Ĳ�ͬ���㷽ʽ�Ľ��
  Cursor c_Opermoney(n_Id ����Ԥ����¼.����id%Type) Is
    Select Distinct b.���㷽ʽ, -1 * Nvl(b.��Ԥ��, 0) As ��Ԥ��
    From ����Ԥ����¼ B
    Where b.����id = n_Id And b.��¼���� = 4 And b.��¼״̬ = 2 And Nvl(b.��Ԥ��, 0) <> 0;

  v_Err_Msg Varchar(255);
  Err_Item Exception;

  n_����id ����Ԥ����¼.����id%Type;
  n_����id ������ü�¼.����id%Type;

  v_��ָ�����㷽ʽ ����Ԥ����¼.���㷽ʽ%Type;
  n_�˿���       ����Ԥ����¼.��Ԥ��%Type;
  n_��ӡid         Ʊ�ݴ�ӡ����.Id%Type;
  n_����id         ������Ϣ.����id%Type;
  n_�˷ѽ��       ����Ԥ����¼.��Ԥ��%Type;
  n_Ԥ�����       ����Ԥ����¼.��Ԥ��%Type; --ԭ��¼ Ԥ���ɿ���
  n_����ֵ         �������.Ԥ�����%Type;
  n_�Һ�id         ���˹Һż�¼.Id%Type;
  n_��id           ����ɿ����.Id%Type;

  n_�����˷�       Number; --��¼�Ƿ��Ǵ˵��ݵĵڶ����˷�
  n_����̨ǩ���Ŷ� Number;
  n_ԤԼ���ɶ���   Number;
  n_ԤԼ�Һ�       Number;
  n_�Һ����ɶ���   Number;
  d_Date           Date;
  n_����           ������ü�¼.���ʷ���%Type;
  n_����id1        ������Ϣ.����id%Type;
  n_���ض�         ������ü�¼.ʵ�ս��%Type;
  n_�ѽ���         Number;
  n_���           ���˹Һż�¼.����%Type;
  n_���ﲡ��id     ������Ϣ.����id%Type;
  d_����ʱ��       ����ǼǼ�¼.����ʱ��%Type;
  n_�����¼id     �ٴ������¼.Id%Type;
  v_��������       Varchar2(5000);
  v_��ǰ����       Varchar2(1000);
  v_����ids        Varchar2(500);
  v_Temp           Varchar2(500);
  v_���㷽ʽ       ����Ԥ����¼.���㷽ʽ%Type;
  n_��������־     Number;
  n_������       ����Ԥ����¼.��Ԥ��%Type;
  n_�����         Number;
  n_Count          Number;
Begin
  n_��id           := Zl_Get��id(����Ա����_In);
  v_��ָ�����㷽ʽ := ��ָ������_In;

  Select �����¼id, ���� Into n_�����¼id, n_��� From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum < 2;

  --�����ж�Ҫ�˺�/ȡ��ԤԼ�ļ�¼�Ƿ����
  Open c_Moneyinfo(1);
  Fetch c_Moneyinfo
    Into r_Moneyrow;
  If c_Moneyinfo%RowCount = 0 Then
    Close c_Moneyinfo;
    Open c_Moneyinfo(0);
    Fetch c_Moneyinfo
      Into r_Moneyrow;
    If c_Moneyinfo%RowCount = 0 Then
      v_Err_Msg := 'Ҫ����ĵ��ݲ����ڡ�';
      Raise Err_Item;
    End If;
    n_ԤԼ�Һ� := 1;
  End If;
  Close c_Moneyinfo;

  Begin
    Select Zl_Fun_Regcustomname Into v_Temp From Dual;
    If v_Temp Is Not Null Then
      v_����ids := Substr(v_Temp, Instr(v_Temp, '|') + 1);
    End If;
  Exception
    When Others Then
      v_����ids := Null;
  End;

  --1.ԤԼ����
  If Nvl(n_ԤԼ�Һ�, 0) = 1 Then
    --������Լ��
    Open c_Registinfo(0);
    Fetch c_Registinfo
      Into r_Registrow;
  
    n_����� := Null;
    Update �ٴ������¼ Set ��Լ�� = Nvl(��Լ��, 0) - 1 Where ID = n_�����¼id Returning ��Լ�� Into n_�����;
    If Nvl(n_�����, 0) < 0 Then
      Update �ٴ������¼ Set ��Լ�� = 0 Where ID = n_�����¼id;
    End If;
  
    Update ���˹ҺŻ���
    Set ��Լ�� = Nvl(��Լ��, 0) - 1
    Where ���� = Trunc(r_Registrow.����ʱ��) And Nvl(ҽ��id, 0) = Nvl(r_Registrow.ҽ��id, 0) And
          Nvl(ҽ������, '-') = Nvl(r_Registrow.ҽ������, '-') And ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And
          (���� = r_Registrow.���� Or ���� Is Null);
  
    If Sql%RowCount = 0 Then
      Insert Into ���˹ҺŻ���
        (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, ��Լ��)
      Values
        (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
         Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -1);
    End If;
  
    Close c_Registinfo;
  
    --���¹Һ����״̬
    Update �ٴ�������ſ���
    Set �Һ�״̬ = 0, ����Ա���� = Null
    Where �Һ�״̬ = 2 And ��¼id = n_�����¼id And ��� = n_���;
  
    Update �ٴ�������ſ���
    Set �Һ�״̬ = 4, ����Ա���� = Null
    Where �Һ�״̬ = 2 And ��¼id = n_�����¼id And ��ע = To_Char(n_���);
  
    --��Ӳ��˹Һż�¼�� ������¼
    Select ���˹Һż�¼_Id.Nextval, Sysdate Into n_�Һ�id, d_Date From Dual;
  
    Update ���˹Һż�¼ Set ��¼״̬ = 3 Where NO = ���ݺ�_In And ��¼״̬ = 1 And ��¼���� = 2;
    If Sql%NotFound Then
      v_Err_Msg := 'ԤԼ����' || ���ݺ�_In || '�������ڻ����ڲ���ԭ���Ѿ���ȡ��ԤԼ';
      Raise Err_Item;
    End If;
  
    Insert Into ���˹Һż�¼
      (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
       ����, ����, ����, ԤԼ, ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ҽ�Ƹ��ʽ, �����¼id, ԤԼ����Ա, ԤԼ����Ա���)
      Select n_�Һ�id, NO, ��¼����, 2, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, d_Date, ����ʱ��,
             ����Ա���_In, ����Ա����_In, ����, ����, ����, ԤԼ, Nvl(ժҪ_In, ժҪ) As ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ҽ�Ƹ��ʽ,
             n_�����¼id, ԤԼ����Ա, ԤԼ����Ա���
      From ���˹Һż�¼
      Where NO = ���ݺ�_In;
  
    Update ������ü�¼ Set ��¼״̬ = 3 Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 0;
    Insert Into ������ü�¼
      (ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, ���ʵ�id, ����id, ҽ�����, �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��, ���ʽ, ���˿���id, �ѱ�, �շ����,
       �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id,
       ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���, ����Ա����, ����id, ���ʽ��, ���մ���id, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���, �ɿ���id, ����״̬, ��ת��,
       �Һ�id, ��ҳid)
      Select ���˷��ü�¼_Id.Nextval, ��¼����, NO, ʵ��Ʊ��, 2, ���, ��������, �۸񸸺�, ���ʵ�id, ����id, ҽ�����, �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��, ���ʽ,
             ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, -1 * ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, -1 * Ӧ�ս��,
             -1 * ʵ�ս��, ������, ��������id, ������, ����ʱ��, d_Date, ִ�в���id, ִ����, -1, ִ��ʱ��, ����, ����Ա���_In, ����Ա����_In, Null, Null,
             ���մ���id, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���, �ɿ���id, ����״̬, ��ת��, �Һ�id, ��ҳid
      From ������ü�¼
      Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 3;
  
    --���ԤԼ���ɶ���ʱ��Ҫ�������
  
    n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
    If Nvl(n_ԤԼ���ɶ���, 0) = 1 Then
      --Ҫɾ������
      For v_�Һ� In (Select ID, ����, ִ�в���id, ִ���� From ���˹Һż�¼ Where NO = ���ݺ�_In) Loop
        Zl_�ŶӽкŶ���_Delete(v_�Һ�.ִ�в���id, v_�Һ�.Id);
      End Loop;
    End If;
    Return;
  End If;

  Select Nvl(���ʷ���, 0), ����id, Decode(Sign(Nvl(����id, 0)), 0, 0, 1)
  Into n_����, n_����id, n_�ѽ���
  From ������ü�¼
  Where ��¼���� = 4 And NO = ���ݺ�_In And ��¼״̬ In (1, 3) And Rownum < 2;

  --2.�ҺŴ���
  n_�ѽ��� := Nvl(n_�ѽ���, 0);

  If n_�ѽ��� = 1 And n_���� = 1 Then
    Select Sysdate, Null Into d_Date, n_����id From Dual;
  Else
    Select Sysdate, ���˽��ʼ�¼_Id.Nextval Into d_Date, n_����id From Dual;
  End If;

  ----0-ȫ�� 1-�˹Һŷ� 2-�˲�����
  If Nvl(�˷�����_In, 0) <> 2 Then
    --���ǹ��˲�����ʱ����
    --���¹Һ����״̬
    If �˺�����_In = 1 Then
      Update �ٴ�������ſ���
      Set �Һ�״̬ = 0, ����Ա���� = Null
      Where �Һ�״̬ = 1 And ��¼id = n_�����¼id And ��� = n_���;
    
      Update �ٴ�������ſ���
      Set �Һ�״̬ = 4, ����Ա���� = Null
      Where �Һ�״̬ = 1 And ��¼id = n_�����¼id And ��ע = To_Char(n_���);
    Else
      Update �ٴ�������ſ���
      Set �Һ�״̬ = 4, ����Ա���� = ����Ա����_In
      Where �Һ�״̬ = 1 And ��¼id = n_�����¼id And (��� = n_��� Or ��ע = To_Char(n_���));
    End If;
  
    --���˾���״̬
    If n_����id Is Not Null Then
      Update ������Ϣ Set ����״̬ = 0, �������� = Null Where ����id = n_����id;
    
      --ɾ���������ش���,ֻ�е�ֻ��һ���Һż�¼���Ҳ��˽���������Һ����ڽ���ʱ�Żᴦ��
      If ɾ�������_In = 1 Then
        Delete ���ﲡ����¼ Where ����id = n_����id;
        Update ������Ϣ Set ����� = Null Where ����id = n_����id;
        --���ü�¼�����Һż����������￨����,�Լ����˽��Ѻ��˷ѻ����ʵķ���,�Һż�¼�������
        Update ������ü�¼ Set ��ʶ�� = Null Where �����־ = 1 And ����id = n_����id;
      End If;
    End If;
  
    --�����ʱ���˾��￨��,�˷�ʱ������￨��,�ڷǹ��˲�����ʱ
    n_����id1 := Null;
    Begin
      Select ����id
      Into n_����id1
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And ���ӱ�־ = 2 And Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
  
    If n_����id1 Is Not Null And Nvl(�˷�����_In, 0) <> 2 Then
      Update ������Ϣ
      Set ���￨�� = Null, ����֤�� = Null, Ic���� = Decode(Ic����, ���￨��, Null, Ic����)
      Where ����id = n_����id1;
    End If;
  
  End If;

  --���ǰ���Ƿ��Ѿ������˹�����
  Begin
    Select 1 Into n_�����˷� From ������ü�¼ Where ��¼���� = 4 And NO = ���ݺ�_In And ��¼״̬ = 3 And Rownum < 2;
  Exception
    When Others Then
      n_�����˷� := 0;
  End;

  --������ü�¼
  --������¼
  If Nvl(�˷�����_In, 0) = 0 Or Nvl(�˷�����_In, 0) = 2 Then
    --ȫ��,�˲�����
    --������ü�¼��������¼
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
       ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��,
       ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
             ����Ա���_In, ����Ա����_In, ����ʱ��, d_Date, n_����id,
             Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����,
             Nvl(ժҪ_In, ժҪ) As ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
  
    --ԭʼ��¼
    If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
      Update ������ü�¼
      Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    Else
      Update ������ü�¼
      Set ��¼״̬ = 3
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    End If;
  Elsif Nvl(�˷�����_In, 0) = 1 Or Nvl(�˷�����_In, 0) = 4 Then
    --�˹Һŷ�,�˹Һ��벡����
    --������ü�¼��������¼
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
       ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��,
       ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
             ����Ա���_In, ����Ա����_In, ����ʱ��, d_Date, n_����id,
             Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����,
             Nvl(ժҪ_In, ժҪ) As ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') = 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
  
    --ԭʼ��¼
    If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
      Update ������ü�¼
      Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') = 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    Else
      Update ������ü�¼
      Set ��¼״̬ = 3
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') = 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    End If;
  Elsif Nvl(�˷�����_In, 0) = 3 Then
    --�˸��ӷ�
    --������ü�¼��������¼
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
       ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��,
       ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
             ����Ա���_In, ����Ա����_In, ����ʱ��, d_Date, n_����id,
             Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����,
             Nvl(ժҪ_In, ժҪ) As ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') > 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
  
    --ԭʼ��¼
    If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
      Update ������ü�¼
      Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') > 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    Else
      Update ������ü�¼
      Set ��¼״̬ = 3
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') > 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    End If;
  Elsif Nvl(�˷�����_In, 0) = 5 Then
    --�˹Һ��븽�ӷ�
    --������ü�¼��������¼
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
       ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��,
       ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
             ����Ա���_In, ����Ա����_In, ����ʱ��, d_Date, n_����id,
             Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����,
             Nvl(ժҪ_In, ժҪ) As ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Nvl(���ӱ�־, 0) <> 1;
  
    --ԭʼ��¼
    If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
      Update ������ü�¼
      Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Nvl(���ӱ�־, 0) <> 1;
    Else
      Update ������ü�¼
      Set ��¼״̬ = 3
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Nvl(���ӱ�־, 0) <> 1;
    End If;
  End If;

  n_����id := 0;
  If n_���� = 0 Then
    --��ȡ����ID
    Select Nvl(����id, 0)
    Into n_����id
    From ������ü�¼
    Where ��¼���� = 4 And ��¼״̬ = 3 And NO = ���ݺ�_In And
          Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0)) And
          Rownum = 1;
  End If;

  If n_���� = 1 Then
    --����
    For c_���� In (Select ʵ�ս��, ���˿���id, ��������id, ִ�в���id, ������Ŀid
                 From ������ü�¼
                 Where ��¼���� = 4 And ��¼״̬ = 3 And NO = ���ݺ�_In And
                       Nvl(���ӱ�־, 0) =
                       Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0)) And
                       Nvl(���ʷ���, 0) = 1) Loop
      --�������
      Update �������
      Set ������� = Nvl(�������, 0) - Nvl(c_����.ʵ�ս��, 0)
      Where ����id = Nvl(n_����id, 0) And ���� = 1 And ���� = 1
      Returning ������� Into n_���ض�;
    
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, ����, �������, Ԥ�����)
        Values
          (n_����id, 1, 1, -1 * Nvl(c_����.ʵ�ս��, 0), 0);
        n_���ض� := Nvl(c_����.ʵ�ս��, 0);
      End If;
      If Nvl(n_���ض�, 0) = 0 Then
        Delete �������
        Where ����id = Nvl(n_����id, 0) And ���� = 1 And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
      --����δ�����
      Update ����δ�����
      Set ��� = Nvl(���, 0) - Nvl(c_����.ʵ�ս��, 0)
      Where ����id = n_����id And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(c_����.���˿���id, 0) And
            Nvl(��������id, 0) = Nvl(c_����.��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(c_����.ִ�в���id, 0) And ������Ŀid + 0 = c_����.������Ŀid And
            ��Դ;�� + 0 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into ����δ�����
          (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
        Values
          (n_����id, Null, Null, c_����.���˿���id, c_����.��������id, c_����.ִ�в���id, c_����.������Ŀid, 1, -1 * Nvl(c_����.ʵ�ս��, 0));
      End If;
    End Loop;
    Delete ����δ�����
    Where ����id = n_����id And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���, 0) = 0 And ��Դ;�� + 0 = 1;
  End If;

  If n_���� = 0 Then
    --1.�˷�
    --���˹ҺŽ���:�ֽ�͸����ʻ�����
    If ���㷽ʽ_In Is Null And Nvl(��Ԥ��_In, 0) = 0 Then
      If ��ԭ���˽���_In Is Not Null Then
        --�˿����ȡ
        If Nvl(�˷�����_In, 0) <> 0 Or Nvl(n_�����˷�, 0) <> 0 Then
          --����ǵ����˲�����,����ֻ�˹Һŷ�,�Ȼ�ȡ�˷ѽ��
          Begin
            --��ȡ�����˿���
            Select Sum(Nvl(ʵ�ս��, 0)) As �տ���
            Into n_�˿���
            From ������ü�¼
            Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 3 And
                  Nvl(���ӱ�־, 0) =
                  Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
          
          Exception
            When Others Then
              v_Err_Msg := '���ݡ�' || ���ݺ�_In || '����' || Case Nvl(�˷�����_In, 0)
                             When 1 Then
                              '�Һŷ���'
                             When 2 Then
                              '������'
                           End || '�������ڲ���ԭ���Ѿ��������˷ѻ��ߵ��ݲ�����!';
              Raise Err_Item;
          End;
          Begin
            Select ��Ԥ��
            Into n_�˷ѽ��
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
          Exception
            When Others Then
              n_�˷ѽ�� := 0;
          End;
        
          --a.����Ľ��㷽ʽ
        
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In, -n_�˿���,
                   n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
        
          If n_�˷ѽ�� = 0 Then
            --b.����������ֽ�
            If n_�˿��� <> 0 Then
              If v_��ָ�����㷽ʽ Is Null Then
                --�˸��ֽ�
                Begin
                  Select ���� Into v_��ָ�����㷽ʽ From ���㷽ʽ Where ���� = 1;
                Exception
                  When Others Then
                    v_��ָ�����㷽ʽ := '�ֽ�';
                End;
              End If;
              Update ����Ԥ����¼
              Set ��Ԥ�� = ��Ԥ�� + (-1 * n_�˿���), ����˵�� = Nvl(����˵��_In, ����˵��)
              Where ��¼���� = 4 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_��ָ�����㷽ʽ;
              If Sql%RowCount = 0 Then
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����,
                   �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
                  Select ����Ԥ����¼_Id.Nextval, a.No, a.ʵ��Ʊ��, a.��¼����, 2, a.����id, a.��ҳid, a.����id, a.ժҪ, v_��ָ�����㷽ʽ, d_Date,
                         ����Ա���_In, ����Ա����_In, -1 * n_�˿���, n_����id, n_��id, Ԥ�����, Decode(����˵��_In, Null, �����id, Null), ���㿨���,
                         Decode(����˵��_In, Null, ����, Null), Decode(����˵��_In, Null, ������ˮ��, Null), Nvl(����˵��_In, ����˵��), ������λ,
                         4
                  From ����Ԥ����¼ A
                  Where a.��¼���� = 4 And a.��¼״̬ = 1 And a.����id = n_����id And Rownum < 2;
              End If;
            End If;
          End If;
        Else
          --a.����Ľ��㷽ʽԭ����
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In, -��Ԥ��,
                   n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
        
          --b.����������ֽ�
          Begin
            Select Sum(��Ԥ��)
            Into n_�˷ѽ��
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') > 0;
          Exception
            When Others Then
              n_�˷ѽ�� := 0;
          End;
          If n_�˷ѽ�� <> 0 Then
            If v_��ָ�����㷽ʽ Is Null Then
              --�˸��ֽ�
              Begin
                Select ���㷽ʽ
                Into v_��ָ�����㷽ʽ
                From ����Ԥ����¼
                Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And
                      Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
              
              Exception
                When Others Then
                  Begin
                    Select ���� Into v_��ָ�����㷽ʽ From ���㷽ʽ Where ���� = 1;
                  Exception
                    When Others Then
                      v_��ָ�����㷽ʽ := '�ֽ�';
                  End;
              End;
            End If;
            Update ����Ԥ����¼
            Set ��Ԥ�� = ��Ԥ�� + (-1 * n_�˷ѽ��), ����˵�� = Nvl(����˵��_In, ����˵��)
            Where ��¼���� = 4 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_��ָ�����㷽ʽ;
            If Sql%RowCount = 0 Then
              Insert Into ����Ԥ����¼
                (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����,
                 �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
                Select ����Ԥ����¼_Id.Nextval, a.No, a.ʵ��Ʊ��, a.��¼����, 2, a.����id, a.��ҳid, a.����id, a.ժҪ, v_��ָ�����㷽ʽ, d_Date,
                       ����Ա���_In, ����Ա����_In, -1 * n_�˷ѽ��, n_����id, n_��id, Ԥ�����, Decode(����˵��_In, Null, �����id, Null), ���㿨���,
                       Decode(����˵��_In, Null, ����, Null), Decode(����˵��_In, Null, ������ˮ��, Null), Nvl(����˵��_In, ����˵��), ������λ, 4
                From ����Ԥ����¼ A
                Where a.��¼���� = 4 And a.��¼״̬ = 1 And a.����id = n_����id And Rownum < 2;
            End If;
          End If;
        End If;
      Else
        --�˿����ȡ
        If Nvl(�˷�����_In, 0) <> 0 Or Nvl(n_�����˷�, 0) <> 0 Then
          --����ǵ����˲�����,����ֻ�˹Һŷ�,�Ȼ�ȡ�˷ѽ��
          Begin
            --��ȡ�����˿���
            Select Sum(Nvl(ʵ�ս��, 0)) As �տ���
            Into n_�˿���
            From ������ü�¼
            Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 3 And
                  Nvl(���ӱ�־, 0) =
                  Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
          Exception
            When Others Then
              v_Err_Msg := '���ݡ�' || ���ݺ�_In || '����' || Case Nvl(�˷�����_In, 0)
                             When 1 Then
                              '�Һŷ���'
                             When 2 Then
                              '������'
                           End || '�������ڲ���ԭ���Ѿ��������˷ѻ��ߵ��ݲ�����!';
              Raise Err_Item;
          End;
        End If;
        If Nvl(n_�����˷�, 0) = 0 And Nvl(�˷�����_In, 0) = 0 Then
          --�״�ȫ��
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                   -1 * ��Ԥ��, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id;
        Else
          --�����˷�,���߱��ε���һ����
          --�����˷�ʱ,��¼״̬=3 ,�״β�����,��¼״̬Ϊ1
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                   -1 * n_�˿���, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And ժҪ = 'ҽ���Һ�' And
                  ��Ԥ�� = n_�˿��� And Rownum < 2;
          If Sql%RowCount = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
               ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
              Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                     -1 * n_�˿���, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
              From ����Ԥ����¼
              Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And ��Ԥ�� = n_�˿��� And
                    Rownum < 2;
          End If;
          If Sql%RowCount = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
               ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
              Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                     -1 * n_�˿���, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
              From ����Ԥ����¼
              Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And Rownum < 2;
            If Sql%RowCount = 0 Then
              --�����˷�,����ȫ��ʹ��Ԥ����ɷ�ʱ�Ŵ��ڴ������
              n_Ԥ����� := n_�˿���;
            End If;
          End If;
        
        End If;
      End If;
    Else
      --�����㷽ʽ��
      If ���㷽ʽ_In Is Not Null Then
        v_�������� := ���㷽ʽ_In || '|'; --�Կո�ֿ���|��β,û�н�������
        While v_�������� Is Not Null Loop
          v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '|') - 1);
          v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
        
          v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
          n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
        
          v_��ǰ����   := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
          n_��������־ := To_Number(v_��ǰ����);
        
          If n_��������־ = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
               ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������, �������)
              Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, v_���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                     -1 * n_������, n_����id, n_��id, Ԥ�����, Null, Null, Null, Null, ����˵��_In, ������λ, 4, �������
              From ����Ԥ����¼
              Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And Rownum < 2;
          Else
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
               ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������, �������)
              Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, v_���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                     -1 * n_������, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, Nvl(����˵��_In, ����˵��), ������λ, 4, �������
              
              From ����Ԥ����¼
              Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And
                    (�����id Is Not Null Or ���㿨��� Is Not Null) And Rownum < 2;
          End If;
          v_�������� := Substr(v_��������, Instr(v_��������, '|') + 1);
        End Loop;
      End if;
      n_Ԥ����� := Nvl(��Ԥ��_In, 0);
    End If;
    --�״��˷�ʱ,��¼״̬�����Ϊ��3
    Update ����Ԥ����¼
    Set ��¼״̬ = 3
    Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id;
  
    --��Ԥ�� 1-ȫ�� 2-������,������ʱ��ȫ��ʹ��Ԥ�����нɿ�
    If Nvl(�˷�����_In, 0) = 0 Or (Nvl(�˷�����_In, 0) <> 0 And n_Ԥ����� <> 0) Then
      --���˹ҺŽ���:��Ԥ�����
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
        Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, d_Date,
               ����Ա����_In, ����Ա���_In, -1 * Decode(Nvl(�˷�����_In, 0) + Nvl(n_�����˷�, 0), 0, ��Ԥ��, n_Ԥ�����), n_����id, n_��id, Ԥ�����,
               �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
        From ����Ԥ����¼
        Where ��¼���� In (1, 11) And ����id = n_����id And Nvl(��Ԥ��, 0) <> 0 And
              Rownum = Decode(Nvl(�˷�����_In, 0) + Nvl(n_�����˷�, 0), 0, Rownum, 1);
    End If;
  
    --����˿ʽ���˿���
    Select Count(1) Into n_Count From ����Ԥ����¼ Where ����id = n_����id And ���㷽ʽ Is Null And Rownum < 2;
    IF n_Count > 0 Then
      v_Err_Msg := '������δ�ɿ������,������ɽ���!';
      Raise Err_Item;
    End if;
    
    Select a.ʵ��, b.��Ԥ��
    Into n_�˷ѽ��, n_�˿���
    From (Select Sum(ʵ�ս��) As ʵ�� From ������ü�¼ Where ����id = n_����id) a,
         (Select Sum(��Ԥ��) As ��Ԥ�� From ����Ԥ����¼ Where ����id = n_����id) b;
    IF Nvl(n_�˷ѽ��, 0) <> Nvl(n_�˿���, 0) Then
      v_Err_Msg := '��������˿��һ��,������ɽ���!';
      Raise Err_Item;
    End if;
    
    --������Ԥ�����
    For c_Ԥ�� In (Select ����id, Ԥ�����, -1 * Sum(Nvl(��Ԥ��, 0)) As ��Ԥ��
                 From ����Ԥ����¼
                 Where ��¼���� In (1, 11) And ����id = n_����id
                 Group By ����id, Ԥ�����) Loop
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(c_Ԥ��.��Ԥ��, 0)
      Where ����id = c_Ԥ��.����id And ���� = Nvl(c_Ԥ��.Ԥ�����, 2) And ���� = 1
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, Ԥ�����, ����, ����)
        Values
          (c_Ԥ��.����id, Nvl(c_Ԥ��.��Ԥ��, 0), 1, Nvl(c_Ԥ��.Ԥ�����, 2));
        n_����ֵ := Nvl(c_Ԥ��.��Ԥ��, 0);
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From �������
        Where ����id = c_Ԥ��.����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
      End If;
    End Loop;
  
    If �ջ�Ʊ�ݺ�_In Is Not Null Then
      --���˹Һŷ�,������Ʊ��
      --�˿��ջ�Ʊ��(�����ϴιҺ�ʹ��Ʊ��,�����ջ�)
      Begin
        --�����һ�δ�ӡ��������ȡ
        Select ID
        Into n_��ӡid
        From (Select b.Id
               From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
               Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And b.�������� = 4 And b.No = ���ݺ�_In
               Order By a.ʹ��ʱ�� Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          n_��ӡid := Null;
      End;
    
      --���ջ�ԭƱ��
      If n_��ӡid Is Not Null Then
        Begin
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
            Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, ����Ա����_In, Ʊ�ݽ��
            From Ʊ��ʹ����ϸ
            Where ��ӡid = n_��ӡid And ���� = 1 And Instr(',' || �ջ�Ʊ�ݺ�_In || ',', ',' || ���� || ',') > 0;
        Exception
          When Others Then
            Delete From Ʊ��ʹ����ϸ Where ��ӡid = n_��ӡid And ���� = 2 And ԭ�� = 2;
            Insert Into Ʊ��ʹ����ϸ
              (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
              Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, ����Ա����_In, Ʊ�ݽ��
              From Ʊ��ʹ����ϸ
              Where ��ӡid = n_��ӡid And ���� = 1 And Instr(',' || �ջ�Ʊ�ݺ�_In || ',', ',' || ���� || ',') > 0;
        End;
      End If;
    End If;
  End If;

  --�����˲�������,��������ܼ�¼
  --��ػ��ܱ�Ĵ���

  --���˹ҺŻ���
  Open c_Registinfo(3);
  Fetch c_Registinfo
    Into r_Registrow;

  If c_Registinfo%RowCount = 0 Then
    --ֻ�ղ�����ʱ�޺ű�,������
    Close c_Registinfo;
  Else
  
    --��Ҫȷ���Ƿ�ԤԼ�Һ�
    --1.�������ԤԼ�ҺŲ����ĹҺż�¼,����Ҫ���ѹ����������ѽ���
    --2.����������Һ�,��ֻ���ѹ���
  
    Begin
      Select Decode(ԤԼ, Null, 0, 0, 0, 1) Into n_ԤԼ�Һ� From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum = 1;
    Exception
      When Others Then
        n_ԤԼ�Һ� := 0;
    End;
    n_����� := Null;
    Update �ٴ������¼
    Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - n_ԤԼ�Һ�, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ�Һ�
    Where ID = n_�����¼id
    Returning �ѹ��� Into n_�����;
  
    If Nvl(n_�����, 0) < 0 Then
      Update �ٴ������¼ Set �ѹ��� = 0 Where ID = n_�����¼id;
    End If;
  
    Update ���˹ҺŻ���
    Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - n_ԤԼ�Һ�, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ�Һ�
    Where ���� = Trunc(r_Registrow.����ʱ��) And Nvl(ҽ��id, 0) = Nvl(r_Registrow.ҽ��id, 0) And
          Nvl(ҽ������, '-') = Nvl(r_Registrow.ҽ������, '-') And ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And
          (���� = r_Registrow.���� Or ���� Is Null);
  
    If Sql%RowCount = 0 Then
      Insert Into ���˹ҺŻ���
        (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, �����ѽ���, ��Լ��)
      Values
        (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
         Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -1, -1 * n_ԤԼ�Һ�, -1 * n_ԤԼ�Һ�);
    End If;
  
    Close c_Registinfo;
  End If;

  If n_���� = 0 Then
    --��Ա�ɿ����(���������ʻ��ȵĽ�����,�����˳�Ԥ����)
    For r_Opermoney In c_Opermoney(n_����id) Loop
      Update ��Ա�ɿ����
      Set ��� = Nvl(���, 0) + (-1 * r_Opermoney.��Ԥ��)
      Where �տ�Ա = ����Ա����_In And ���㷽ʽ = r_Opermoney.���㷽ʽ And ���� = 1
      Returning ��� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ��Ա�ɿ����
          (�տ�Ա, ���㷽ʽ, ����, ���)
        Values
          (����Ա����_In, r_Opermoney.���㷽ʽ, 1, -1 * r_Opermoney.��Ԥ��);
        n_����ֵ := r_Opermoney.��Ԥ��;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From ��Ա�ɿ����
        Where �տ�Ա = ����Ա����_In And ���㷽ʽ = r_Opermoney.���㷽ʽ And ���� = 1 And Nvl(���, 0) = 0;
      End If;
    End Loop;
  End If;
  If Nvl(�˷�����_In, 0) Not In (2, 3) Then
    n_�Һ����ɶ��� := Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113));
    If n_�Һ����ɶ��� <> 0 Then
      --Ҫɾ������
      For v_�Һ� In (Select ID, ����, ִ�в���id, ִ���� From ���˹Һż�¼ Where NO = ���ݺ�_In) Loop
        n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113, 1, Nvl(v_�Һ�.ִ�в���id, 0)));
        If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Then
          Zl_�ŶӽкŶ���_Delete(v_�Һ�.ִ�в���id, v_�Һ�.Id);
        End If;
      End Loop;
    End If;
  
    --ҽ�������ľ���ǼǼ�¼
    Begin
      Select ����id, ����ʱ�� Into n_���ﲡ��id, d_����ʱ�� From ���˹Һż�¼ Where NO = ���ݺ�_In;
      Delete From ����ǼǼ�¼ Where ����id = n_���ﲡ��id And ����ʱ�� = d_����ʱ�� And ��ҳid Is Null;
    Exception
      When Others Then
        Null;
    End;
  
    --���˹Һż�¼
    Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
  
    Update ���˹Һż�¼ Set ��¼״̬ = 3 Where NO = ���ݺ�_In And ��¼���� = 1 And ��¼״̬ = 1;
    If Sql%NotFound Then
      v_Err_Msg := '�Һŵ���' || ���ݺ�_In || '�������ڻ����ڲ���ԭ���Ѿ����˺�';
      Raise Err_Item;
    End If;
    Insert Into ���˹Һż�¼
      (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
       ����, ����, ����, ԤԼ, ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ����, ҽ�Ƹ��ʽ, �����¼id)
      Select n_�Һ�id, NO, ��¼����, 2, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, d_Date, ����ʱ��,
             ����Ա���_In, ����Ա����_In, ����, ����, ����, ԤԼ, Nvl(ժҪ_In, ժҪ) As ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ����, ҽ�Ƹ��ʽ,
             n_�����¼id
      From ���˹Һż�¼
      Where NO = ���ݺ�_In;
  End If;
  --��Ϣ����
  Begin
    Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
      Using 2, ���ݺ�_In;
  Exception
    When Others Then
      Null;
  End;
  b_Message.Zlhis_Regist_003(n_�Һ�id, ���ݺ�_In);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˹Һż�¼_����_Delete;
/

--119722:����,2019-01-22,���Ӵ���
Create Or Replace Procedure Zl_ҩƷ�ƻ���������_Insert
(
  Id_In       In ҩƷ�ɹ��ƻ�.No%Type,
  No_In       In ҩƷ�ɹ��ƻ�.No%Type,
  �ƻ�����_In In ҩƷ�ɹ��ƻ�.�ƻ�����%Type,
  �ڼ�_In     In ҩƷ�ɹ��ƻ�.�ڼ�%Type,
  �ⷿid_In   In ҩƷ�ɹ��ƻ�.�ⷿid%Type := Null,
  ҩ��id_In   In ҩƷ�ɹ��ƻ�.ҩ��id%Type := Null,
  ���Ʒ���_In In ҩƷ�ɹ��ƻ�.���Ʒ���%Type,
  ������_In   In ҩƷ�ɹ��ƻ�.������%Type,
  ��������_In In ҩƷ�ɹ��ƻ�.��������%Type,
  ����˵��_In In ҩƷ�ɹ��ƻ�.����˵��%Type := Null,
  ��Դ�ⷿ_In In ҩƷ�ɹ��ƻ�.��Դ�ⷿ%Type := Null,
  ��Դҩ��_In In ҩƷ�ɹ��ƻ�.��Դҩ��%Type := Null
) Is
Begin
  Insert Into ҩƷ�ɹ��ƻ�
    (ID, NO, �ƻ�����, �ڼ�, �ⷿid, ҩ��id, ���Ʒ���, ����˵��, ������, ��������, ��Դ�ⷿ, ��Դҩ��)
  Values
    (Id_In, No_In, �ƻ�����_In, �ڼ�_In, �ⷿid_In, ҩ��id_In, ���Ʒ���_In, ����˵��_In, ������_In, ��������_In, ��Դ�ⷿ_In, ��Դҩ��_In);
End Zl_ҩƷ�ƻ���������_Insert;
/



------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0047' Where ���=&n_System;
Commit;
