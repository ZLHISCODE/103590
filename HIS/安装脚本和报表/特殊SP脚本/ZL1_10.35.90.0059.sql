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
--138877:��͢��,2019-04-23,�������۲���תסԺ����ê�㶨��
Insert Into Zlmsg_Lists (Bz_Type, Code, Name, Key_Define, Note, Using)
Select '����', 'ZLHIS_PATIENT_029', '���۲���תסԺ����', '<root><����ID></����ID><��ҳID></��ҳID><�䶯ID></�䶯ID></root>', '���۲���תסԺ����ʱ', 1 From Dual;

-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--139063:Ƚ����,2019-04-25,�������۲��˰��������̾���
Create Or Replace Procedure Zl_������ʼ�¼_Delete
(
  No_In           ������ü�¼.No%Type,
  ���_In         Varchar2,
  ����Ա���_In   ������ü�¼.����Ա���%Type,
  ����Ա����_In   ������ü�¼.����Ա����%Type,
  ��Һ��ҩ���_In Number := 1,
  �Ǽ�ʱ��_In     סԺ���ü�¼.�Ǽ�ʱ��%Type := Sysdate
) As
  --���ܣ�����һ��������ʵ�����ָ������� 
  --��ţ���ʽ��"1,3,5,7,8",��"1:2:33456,3:2,5:2,7:2,8:2",ð��ǰ������ֱ�ʾ�к�,�м�����ֱ�ʾ�˵�����,��������ֱ�ʾ��ҩ��¼��ID,Ŀǰ�����������ʱ�Ŵ��� 
  --      Ϊ�ձ�ʾ�������пɳ����� 

  --���α�ΪҪ�˷ѵ��ݵ�����ԭʼ��¼ 
  Cursor c_Bill(n_��־ Number) Is
    Select a.Id, a.�۸񸸺�, a.���, a.ִ��״̬, a.�շ����, a.ҽ�����, a.����id, a.��ҳid, a.������Ŀid, a.��������id, a.ִ�в���id, a.���˲���id, a.���˿���id,
           a.ʵ�ս��, Decode(a.��¼״̬, 0, 1, 0) As ����, j.�������, m.��������
    From ������ü�¼ A, ����ҽ����¼ J, �������� M
    Where a.ҽ����� = j.Id(+) And a.�շ�ϸĿid + 0 = m.����id(+) And a.No = No_In And a.��¼���� = 2 And a.��¼״̬ In (0, 1, 3) And
          a.�����־ = n_��־
    Order By a.�շ�ϸĿid, a.���;

  --���α����ڴ�����ü�¼��� 
  Cursor c_Serial Is
    Select ���, �۸񸸺� From ������ü�¼ Where NO = No_In And ��¼���� = 2 And ��¼״̬ In (0, 1, 3) Order By ���;
  l_���� t_Numlist := t_Numlist();

  v_ҽ��ids  Varchar2(4000);
  n_����     ������ü�¼.�۸񸸺�%Type;
  n_�����־ ������ü�¼.�����־%Type;

  --�����˷Ѽ������ 
  n_ʣ������ Number;
  n_ʣ��Ӧ�� Number;
  n_ʣ��ʵ�� Number;
  n_ʣ��ͳ�� Number;

  n_׼������ Number;
  n_�˷Ѵ��� Number;
  n_�˷����� Number;
  n_�������� Number;

  n_Ӧ�ս�� Number;
  n_ʵ�ս�� Number;
  n_ͳ���� Number;

  v_���   Varchar2(4000);
  v_��ҩid Varchar2(4000);
  v_Tmp    Varchar2(4000);

  n_δִ������ ҩƷ�շ���¼.ʵ������%Type;
  n_��ִ������ ҩƷ�շ���¼.ʵ������%Type;

  n_Dec Number;

  n_Count   Number;
  d_Curdate Date;
  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin
  --�������ʱ,��ҩƷ�ᴫ���кŵ��������� 
  If Not ���_In Is Null Then
    If Instr(���_In, ':') > 0 Then
      --��ʽ��1:2:33456,3:2,5:2,7:2,8:2 
      For c_��� In (Select C1, C2 From Table(f_Str2list2(���_In, ',', ':'))) Loop
        v_��� := v_��� || ',' || c_���.C1;
        If Instr(c_���.C2, ':') > 0 Then
          v_��ҩid := v_��ҩid || ',' || Substr(c_���.C2, Instr(c_���.C2, ':') + 1);
        End If;
      End Loop;
      v_���   := Substr(v_���, 2);
      v_��ҩid := Substr(v_��ҩid, 2);
    Else
      v_��� := ���_In;
    End If;
  End If;

  --�Ƿ��Ѿ�ȫ����ȫִ��(ֻ�����ŵ��ݵļ��) 
  Select Nvl(Count(1), 0), Max(Nvl(�����־, 1))
  Into n_Count, n_�����־
  From ������ü�¼
  Where NO = No_In And ��¼���� = 2 And ��¼״̬ In (0, 1, 3) And Nvl(ִ��״̬, 0) <> 1;
  If n_Count = 0 Then
    v_Err_Msg := '�õ����е���Ŀ�Ѿ�ȫ����ȫִ�У�';
    Raise Err_Item;
  End If;

  If Nvl(n_�����־, 0) = 0 Then
    n_�����־ := 1;
  End If;

  --δ��ȫִ�е���Ŀ�Ƿ���ʣ������(ֻ�����ŵ��ݵļ��) 
  Select Nvl(Count(1), 0)
  Into n_Count
  From (Select ���, Sum(����) As ʣ������
         From (Select ��¼״̬, Nvl(�۸񸸺�, ���) As ���, Avg(Nvl(����, 1) * ����) As ����
                From ������ü�¼
                Where NO = No_In And ��¼���� = 2 And �����־ = n_�����־ And
                      Nvl(�۸񸸺�, ���) In
                      (Select Nvl(�۸񸸺�, ���)
                       From ������ü�¼
                       Where NO = No_In And ��¼���� = 2 And �����־ = n_�����־ And ��¼״̬ In (0, 1, 3) And Nvl(ִ��״̬, 0) <> 1)
                Group By ��¼״̬, Nvl(�۸񸸺�, ���))
         Group By ���
         Having Sum(����) <> 0);
  If n_Count = 0 Then
    v_Err_Msg := '�õ�����δ��ȫִ�в�����Ŀʣ������Ϊ��,û�п������ʵķ��ã�';
    Raise Err_Item;
  End If;

  --------------------------------------------------------------------------------- 
  --���ñ��� 
  Select Nvl(�Ǽ�ʱ��_In, Sysdate) Into d_Curdate From Dual;

  --���С��λ�� 
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  --ѭ������ÿ�з���(������Ŀ��) 
  For r_Bill In c_Bill(n_�����־) Loop
    If Instr(',' || v_��� || ',', ',' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || ',') > 0 Or v_��� Is Null Then
      If Nvl(r_Bill.ִ��״̬, 0) <> 1 Then
        --��ʣ������,ʣ��Ӧ��,ʣ��ʵ�� 
        Select Sum(Nvl(����, 1) * ����), Sum(Ӧ�ս��), Sum(ʵ�ս��), Sum(ͳ����)
        Into n_ʣ������, n_ʣ��Ӧ��, n_ʣ��ʵ��, n_ʣ��ͳ��
        From ������ü�¼
        Where NO = No_In And ��¼���� = 2 And ��� = r_Bill.���;
      
        n_�������� := 0;
        n_�˷����� := 0;
        If n_ʣ������ = 0 Then
          If v_��� Is Not Null Then
            v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����Ѿ�ȫ�����ʣ�';
            Raise Err_Item;
          End If;
          --�����δ�޶��к�,ԭʼ�����еĸñ��Ѿ�ȫ������(ִ��״̬=0��һ�ֿ���) 
        Else
          If Instr(���_In, ':') > 0 Then
            Select Max(a.C2) Into v_Tmp From Table(f_Str2list2(���_In, ',', ':')) A Where a.C1 = r_Bill.���;
            If Instr(v_Tmp, ':') > 0 Then
              n_�˷����� := Substr(v_Tmp, 1, Instr(v_Tmp, ':') - 1);
            Else
              n_�˷����� := To_Number(v_Tmp);
            End If;
            n_�������� := 1;
          End If;
        
          --׼������(��ҩƷ��ĿΪʣ������,ԭʼ����) 
          If Instr(',4,5,6,7,', r_Bill.�շ����) = 0 Or (r_Bill.�շ���� = '4' And Nvl(r_Bill.��������, 0) = 0) Then
            --@@@ 
            --��ҩƷ����(�Ծ���ҽ��ִ��Ϊ׼���м��) 
            --: 1.����ҽ�����͵�,����ҽ��ִ��Ϊ׼(�����ܰ���:���;����;����;������Ѫ) 
            --: 2.���ڲ���ҽ�ԼƼ��е��շѷ�ʽΪ:0-������ȡ ��,��֧�ֲ�����;�����������,��ֻ��ȫ�� 
            --: 3.������ҽ����,����ʣ������Ϊ׼ 
            n_Count := 0;
            If Instr(',C,D,F,G,K,', ',' || r_Bill.������� || ',') = 0 And r_Bill.������� Is Not Null Then
              Select Nvl(Sum(����), 0), Count(*)
              Into n_׼������, n_Count
              From (Select j.ҽ����� As ҽ��id, j.�շ�ϸĿid, Nvl(j.����, 1) * Nvl(j.����, 1) As ����
                     From ������ü�¼ J, ����ҽ����¼ M
                     Where j.ҽ����� = m.Id And j.No = No_In And j.��¼���� = 2 And j.��� = r_Bill.��� And j.��¼״̬ In (1, 3) And
                           Exists
                      (Select 1
                            From ����ҽ������ A
                            Where a.ҽ��id = j.ҽ����� And Nvl(a.ִ��״̬, 0) <> 1 And a.No || '' = No_In) And Exists
                      (Select 1
                            From ����ҽ���Ƽ� A
                            Where a.ҽ��id = j.ҽ����� And a.�շ�ϸĿid = j.�շ�ϸĿid And Nvl(a.�շѷ�ʽ, 0) = 0) And j.�۸񸸺� Is Null And
                           Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And
                           (j.��¼״̬ In (1, 3) And Not Exists
                            (Select 1
                             From ҩƷ�շ���¼
                             Where ����id = j.Id And Instr(',8,9,10,21,24,25,26,', ',' || ���� || ',') > 0) Or
                            j.��¼״̬ = 2 And Not Exists
                            (Select 1 From ҩƷ�շ���¼ Where NO = No_In And ���� In (8, 24) And ҩƷid = j.�շ�ϸĿid))
                     Union All
                     Select a.ҽ��id, a.�շ�ϸĿid, -1 * Nvl(a.����, 1) * Nvl(c.��������, 1) As ����
                     From ����ҽ���Ƽ� A, ����ҽ������ B, ����ҽ��ִ�� C, ������ü�¼ J, ����ҽ����¼ M
                     Where a.ҽ��id = b.ҽ��id And b.ҽ��id = c.ҽ��id And Nvl(a.�շѷ�ʽ, 0) = 0 And b.���ͺ� = c.���ͺ� And a.ҽ��id = m.Id And
                           Nvl(c.ִ�н��, 1) = 1 And Nvl(b.ִ��״̬, 0) <> 1 And a.ҽ��id = j.ҽ����� And a.�շ�ϸĿid = j.�շ�ϸĿid And
                           j.No = No_In And j.��¼���� = 2 And j.��� = r_Bill.��� And j.��¼״̬ In (1, 3) And j.�۸񸸺� Is Null And
                           Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And Not Exists
                      (Select 1
                            From ҩƷ�շ���¼
                            Where ����id = j.Id And Instr(',8,9,10,21,24,25,26,', ',' || ���� || ',') > 0) And Not Exists
                      (Select 1 From �������� Where ����id = j.�շ�ϸĿid And Nvl(��������, 0) = 1)
                     Union All
                     Select a.ҽ��id, a.�շ�ϸĿid, 0 As ����
                     From ����ҽ���Ƽ� A, ������ü�¼ J, ����ҽ����¼ M
                     Where a.ҽ��id = m.Id And a.ҽ��id = j.ҽ����� And a.�շ�ϸĿid = j.�շ�ϸĿid And Nvl(a.�շѷ�ʽ, 0) <> 0 And
                           j.No = No_In And j.��¼���� = 2 And Nvl(j.ִ��״̬, 0) = 2 And Not Exists
                      (Select 1 From �������� Where ����id = j.�շ�ϸĿid And Nvl(��������, 0) = 1) And
                           Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0);
            End If;
          
            If Nvl(n_Count, 0) = 0 Then
              n_׼������ := n_ʣ������;
            End If;
          Else
            Select Sum(Nvl(����, 1) * ʵ������)
            Into n_׼������
            From ҩƷ�շ���¼
            Where NO = No_In And ���� In (9, 25) And Mod(��¼״̬, 3) = 1 And ����� Is Null And ����id = r_Bill.Id;
          
            --���������õ��������� 
            If r_Bill.�շ���� = '4' And Nvl(n_׼������, 0) = 0 Then
              n_׼������ := n_ʣ������;
            End If;
          End If;
        
          If Nvl(n_�˷�����, 0) = 0 Then
            n_�˷����� := n_׼������;
          Else
            If n_׼������ < n_�˷����� Then
              v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з���׼���������㱾������������';
              Raise Err_Item;
            End If;
          End If;
        
          --���=ʣ����*(׼����/ʣ����) 
          n_Ӧ�ս�� := Round(n_ʣ��Ӧ�� * (n_�˷����� / n_ʣ������), n_Dec);
          n_ʵ�ս�� := Round(n_ʣ��ʵ�� * (n_�˷����� / n_ʣ������), n_Dec);
          n_ͳ���� := Round(n_ʣ��ͳ�� * (n_�˷����� / n_ʣ������), n_Dec);
        
          If Nvl(r_Bill.����, 0) = 0 Then
            --�ñ���Ŀ�ڼ������� 
            Select Nvl(Max(Abs(ִ��״̬)), 0) + 1
            Into n_�˷Ѵ���
            From ������ü�¼
            Where NO = No_In And ��¼���� = 2 And ��¼״̬ = 2 And ��� = r_Bill.���;
          
            --�����˷Ѽ�¼ 
            Insert Into ������ü�¼
              (ID, NO, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, Ӥ����, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id, �շ����,
               �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ������,
               ִ����, ִ��״̬, ִ��ʱ��, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ������Ŀ��, ���մ���id, ͳ����, ���ʵ�id, ժҪ, ���ձ���, �Ƿ���, ����, �Һ�id, ��ҳid,
               ���˲���id)
              Select ���˷��ü�¼_Id.Nextval, NO, ��¼����, 2, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, Ӥ����, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�,
                     ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ, Decode(Sign(n_�˷����� - Nvl(����, 1) * ����), 0, ����, 1), ��ҩ����,
                     Decode(Sign(n_�˷����� - Nvl(����, 1) * ����), 0, -1 * ����, -1 * n_�˷�����), �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���,
                     ��׼����, -1 * n_Ӧ�ս��, -1 * n_ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����, -1 * n_�˷Ѵ���, ִ��ʱ��, ����Ա���_In,
                     ����Ա����_In, ����ʱ��, d_Curdate, ������Ŀ��, ���մ���id, -1 * n_ͳ����, ���ʵ�id, ժҪ, ���ձ���, �Ƿ���, ����, �Һ�id, ��ҳid,
                     ���˲���id
              From ������ü�¼
              Where ID = r_Bill.Id;
          
            --������� 
            If n_�����־ <> 4 Then
              Update �������
              Set ������� = Nvl(�������, 0) - n_ʵ�ս��
              Where ����id = r_Bill.����id And ���� = 1 And ���� = 1;
              If Sql%RowCount = 0 Then
                Insert Into �������
                  (����id, ����, ����, �������, Ԥ�����)
                Values
                  (r_Bill.����id, 1, 1, -1 * n_ʵ�ս��, 0);
              End If;
            End If;
          
            --����δ����� 
            Update ����δ�����
            Set ��� = Nvl(���, 0) - n_ʵ�ս��
            Where ����id = r_Bill.����id And Nvl(��ҳid, 0) = Nvl(r_Bill.��ҳid, 0) And Nvl(���˲���id, 0) = Nvl(r_Bill.���˲���id, 0) And
                  Nvl(���˿���id, 0) = Nvl(r_Bill.���˿���id, 0) And Nvl(��������id, 0) = Nvl(r_Bill.��������id, 0) And
                  Nvl(ִ�в���id, 0) = Nvl(r_Bill.ִ�в���id, 0) And ������Ŀid + 0 = r_Bill.������Ŀid And ��Դ;�� + 0 = n_�����־;
            If Sql%RowCount = 0 Then
              Insert Into ����δ�����
                (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
              Values
                (r_Bill.����id, r_Bill.��ҳid, r_Bill.���˲���id, r_Bill.���˿���id, r_Bill.��������id, r_Bill.ִ�в���id, r_Bill.������Ŀid,
                 n_�����־, -1 * n_ʵ�ս��);
            End If;
          
            --���ԭ���ü�¼ 
            --ִ��״̬:ȫ������(׼����=ʣ����)���Ϊ0,������Ϊ1 
            If Instr(',4,5,6,7,', r_Bill.�շ����) = 0 Then
              --һ�������ҩƷ�����ĵ���Ŀ,�����ڲ������ʵ����,ֻ������������������ʱ,�Ż���ֲ�������,���� 
              --ִ��״ֻ̬������:0.δִ��;1��ִ��; 
              --������������˹����н���ִ��ǿ�Ƹ�Ϊ��2����ִ��,�����Ҫ�ڴ˴���Ϊ1��ִ��.δִ�еĲ���. 
              Update ������ü�¼
              Set ��¼״̬ = 3, ִ��״̬ = Decode(Sign(n_�˷����� - n_ʣ������), 0, 0, Decode(ִ��״̬, 2, 1, ִ��״̬))
              Where ID = r_Bill.Id;
            Else
              Select Nvl(Sum(Decode(�����, Null, 1, 0) * Nvl(����, 1) * ʵ������), 0),
                     Nvl(Sum(Decode(�����, Null, 0, 1) * Nvl(����, 1) * ʵ������), 0)
              Into n_δִ������, n_��ִ������
              From ҩƷ�շ���¼
              Where NO = No_In And ���� In (9, 10, 25, 26) And ����id = r_Bill.Id;
            
              Update ������ü�¼
              Set ��¼״̬ = 3,
                  ִ��״̬ = Decode(Sign(n_�˷����� - n_ʣ������), 0, 0,
                                 Decode(Sign(n_δִ������ - n_�˷�����), 1, Decode(n_��ִ������, 0, 0, 2), 1))
              Where ID = r_Bill.Id;
            End If;
          Else
            --���ۼ��˵� 
            If Nvl(n_��������, 0) = 0 Then
              l_����.Extend;
              l_����(l_����.Count) := r_Bill.Id;
            Else
              --�������� 
              --���۵�,�Ƚ���ص����ݴ������ڲ����� 
              Update סԺ���ü�¼
              Set ���� = 1, ���� = Nvl(����, 1) * ���� - n_�˷�����, Ӧ�ս�� = Nvl(Ӧ�ս��, 0) - n_Ӧ�ս��, ʵ�ս�� = Nvl(ʵ�ս��, 0) - n_ʵ�ս��,
                  �Ǽ�ʱ�� = d_Curdate, ͳ���� = Nvl(ͳ����, 0) - n_ͳ����
              Where ID = r_Bill.Id
              Returning ���� Into n_ʣ������;
              If Nvl(n_ʣ������, 0) <= 0 Then
                l_����.Extend;
                l_����(l_����.Count) := r_Bill.Id;
              End If;
            End If;
          
            If r_Bill.ҽ����� Is Not Null Then
              If Instr(',' || Nvl(v_ҽ��ids, '') || ',', ',' || r_Bill.ҽ����� || ',') = 0 Then
                v_ҽ��ids := Nvl(v_ҽ��ids, '') || ',' || r_Bill.ҽ�����;
              End If;
            End If;
          End If;
        End If;
      Else
        If v_��� Is Not Null Then
          v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����Ѿ���ȫִ��,�������ʣ�';
          Raise Err_Item;
        End If;
        --���:û�޶��к�,ԭʼ�����а����Ѿ���ȫִ�е� 
      End If;
    End If;
  End Loop;

  --��������ҩID,����ҩƷ�Ƿ�����Һ��ҩ���� 
  If v_��ҩid Is Null And ��Һ��ҩ���_In = 1 Then
    For v_���� In (Select ID
                 From ������ü�¼
                 Where NO = No_In And ��¼���� = 2 And ��¼״̬ In (0, 1, 3) And �շ���� In ('4', '5', '6', '7') And �����־ = n_�����־ And
                       (Instr(',' || v_��� || ',', ',' || ��� || ',') > 0 Or v_��� Is Null)) Loop
      Begin
        Select Count(1)
        Into n_Count
        From ��Һ��ҩ���� A, ҩƷ�շ���¼ B
        Where a.�շ�id = b.Id And b.����id = v_����.Id And Instr(',8,9,10,21,24,25,26,', ',' || b.���� || ',') > 0;
      Exception
        When Others Then
          n_Count := 0;
      End;
      If n_Count <> 0 Then
        v_Err_Msg := '�����Ѿ�������Һ��ҩ���ĵĴ�����ҩƷ���޷�������ʣ�';
        Raise Err_Item;
      End If;
    End Loop;
  End If;

  --------------------------------------------------------------------------------- 
  --ҩƷ��ش���:��Ҫ�Ƕ����������Ч.(�����ǲ���) 
  --���밴�ա��շ�ϸĿid���������򣬷�ֹ��������ҩƷ��桱�� 
  For v_���� In (Select ID, ���
               From ������ü�¼
               Where NO = No_In And ��¼���� = 2 And ��¼״̬ In (0, 1, 3) And �շ���� In ('4', '5', '6', '7') And �����־ = n_�����־ And
                     (Instr(',' || v_��� || ',', ',' || ��� || ',') > 0 Or v_��� Is Null)
               Order By �շ�ϸĿid) Loop
    --���ݷ���ID��������صĴ��� 
    n_�˷����� := 0;
    If Instr(���_In, ':') > 0 Then
      Select Max(a.C2) Into v_Tmp From Table(f_Str2list2(���_In, ',', ':')) A Where a.C1 = v_����.���;
      If Instr(v_Tmp, ':') > 0 Then
        n_�˷����� := Substr(v_Tmp, 1, Instr(v_Tmp, ':') - 1);
      Else
        n_�˷����� := To_Number(v_Tmp);
      End If;
    End If;
    Zl_ҩƷ�շ���¼_�����˷�(v_����.Id, n_�˷�����, v_��ҩid);
  End Loop;

  --ɾ�����ۼ�¼ 
  n_Count := l_����.Count;
  Forall I In 1 .. l_����.Count
    Delete From ������ü�¼ Where ID = l_����(I);

  --ɾ��֮����ͳһ������� 
  If n_Count > 0 Then
    n_Count := 1;
    For r_Serial In c_Serial Loop
      If r_Serial.�۸񸸺� Is Null Then
        n_���� := n_Count;
      End If;
    
      Update ������ü�¼
      Set ��� = n_Count, �۸񸸺� = Decode(�۸񸸺�, Null, Null, n_����)
      Where NO = No_In And ��¼���� = 2 And ��� = r_Serial.���;
    
      Update ������ü�¼ Set �������� = n_Count Where NO = No_In And ��¼���� = 2 And �������� = r_Serial.���;
    
      n_Count := n_Count + 1;
    End Loop;
  End If;

  --���ŵ���ȫ������ʱ��ɾ������ҽ������ 
  For c_ҽ�� In (Select Distinct ҽ�����
               From ������ü�¼
               Where NO = No_In And ��¼���� = 2 And ��¼״̬ = 3 And ҽ����� Is Not Null) Loop
    Select Nvl(Count(*), 0)
    Into n_Count
    From (Select ���, Sum(����) As ʣ������
           From (Select ��¼״̬, Nvl(�۸񸸺�, ���) As ���, Avg(Nvl(����, 1) * ����) As ����
                  From ������ü�¼
                  Where ��¼���� = 2 And ҽ����� + 0 = c_ҽ��.ҽ����� And NO = No_In
                  Group By ��¼״̬, Nvl(�۸񸸺�, ���))
           Group By ���
           Having Sum(����) <> 0);
  
    If n_Count = 0 Then
      Delete From ����ҽ������ Where ҽ��id = c_ҽ��.ҽ����� And ��¼���� = 2 And NO = No_In;
    End If;
  End Loop;

  If v_ҽ��ids Is Not Null Then
    --ҽ������ 
    --����_In    Integer:=0, --0:����;1-סԺ 
    --����_In    Integer:=1, --1-�շѵ�;2-���ʵ� 
    --����_In    Integer:=0, --0:ɾ�����۵�;1-�շѻ����;2-�˷ѻ����� 
    --No_In      ������ü�¼.No%Type, 
    --ҽ��ids_In Varchar2 := Null 
    v_ҽ��ids := Substr(v_ҽ��ids, 2);
    Zl_ҽ������_�Ʒ�״̬_Update(0, 2, 0, No_In, v_ҽ��ids);
  Else
    Zl_ҽ������_�Ʒ�״̬_Update(0, 2, 2, No_In, v_ҽ��ids);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_������ʼ�¼_Delete;
/

--140358:����,2019-04-24,Zl_Getpatient��ȡ����Ӥ��������Ϣʱ,��������û����������ȫ��ɨ��
Create Or Replace Function Zl_Getpatient
(
  Begintime_In In Date,
  Endtime_In   In Date,
  ����id_In    In ���ű�.Id%Type,
  ��ʽid_In    In ���˻����ļ�.��ʽid%Type,
  Type_In      In Varchar2,
  Typeall_In   In Number := 1,
  Split_In     In Varchar2 := ';'
) Return t_Numlist2
  Pipelined As
  n_Index  Number(1);
  v_Str    Varchar2(50);
  n_����id Number;
  n_��ҳid Number;
  P        Number;
  Out_Rec  t_Numobj2 := t_Numobj2(Null, Null);

  --��ȡ��Ӧ��������������Ժ���� 
  Cursor c_List_All Is
    Select b.����id, b.��ҳid
    From ������Ϣ A, ������ҳ B, ��Ժ���� C
    Where a.����id = b.����id And a.��ҳid = b.��ҳid And Nvl(b.��ҳid, 0) <> 0 And Nvl(b.����״̬, 0) <> 5 And b.���ʱ�� Is Null And
          a.����id = c.����id And c.����id = ����id_In;

  --��Ժ�����ڵĲ��� 
  Cursor c_List_Ry Is
    Select b.����id, b.��ҳid
    From ������Ϣ A, ������ҳ B, ��Ժ���� C
    Where a.����id = b.����id And a.��ҳid = b.��ҳid And Nvl(b.��ҳid, 0) <> 0 And Nvl(b.����״̬, 0) <> 5 And b.���ʱ�� Is Null And
          b.��Ժ���� Between Begintime_In And Endtime_In And a.����id = c.����id And c.����id = ����id_In;

  --���������ڵĲ��� 
  Cursor c_List_Ss Is
    Select b.����id, b.��ҳid
    From ������Ϣ A, ������ҳ B, ��Ժ���� F, ���˻����ļ� C, ���˻������� D, ���˻�����ϸ E
    Where a.����id = b.����id And a.��ҳid = b.��ҳid And a.����id = f.����id And Nvl(b.��ҳid, 0) <> 0 And f.����id = ����id_In And
          Nvl(b.����״̬, 0) <> 5 And b.���ʱ�� Is Null And b.����id = c.����id And b.��ҳid = c.��ҳid And c.��ʽid = ��ʽid_In And
          c.Id = d.�ļ�id And d.Id = e.��¼id And e.��¼���� = 4 And e.��Ŀ���� <> '����' And Nvl(e.���Ժϸ�, 0) <> 1 And e.��ֹ�汾 Is Null And
          d.����ʱ�� Between Begintime_In And Endtime_In
    
    Union
    --��ҽ������ȡ����������Ϣ 
    Select b.����id, b.��ҳid
    From ������Ϣ A, ������ҳ B, ��Ժ���� F,
         (Select d.����id, d.��ҳid
           From (Select Distinct a.����id, a.��ҳid
                  From ����ҽ����¼ A, ������ĿĿ¼ B
                  Where a.������Ŀid = b.Id And a.������� = 'F' And a.���id Is Null And a.ҽ��״̬ In (3, 8) And
                        a.��ʼִ��ʱ�� Between Begintime_In And Endtime_In
                  Union
                  Select Distinct a.����id, a.��ҳid
                  From ������������¼ A, ��Ժ���� F
                  Where a.����id = f.����id And a.��ҳid = f.��ҳid And f.����id = ����id_In And a.����ʱ�� Between Begintime_In And
                        Endtime_In) D
           Group By d.����id, d.��ҳid) C
    Where a.����id = b.����id And a.��ҳid = b.��ҳid And a.����id = f.����id And Nvl(b.��ҳid, 0) <> 0 And f.����id = ����id_In And
          Nvl(b.����״̬, 0) <> 5 And b.���ʱ�� Is Null And b.����id = c.����id And b.��ҳid = c.��ҳid;

  --���������´��ڳ���37.5�ȵĲ��� 
  Cursor c_List_Tw Is
    Select b.����id, b.��ҳid
    From ������Ϣ A, ������ҳ B, ��Ժ���� F, ���˻����ļ� C, ���˻������� D, ���˻�����ϸ E
    Where a.����id = b.����id And a.��ҳid = b.��ҳid And a.����id = f.����id And Nvl(b.��ҳid, 0) <> 0 And f.����id = ����id_In And
          Nvl(b.����״̬, 0) <> 5 And b.���ʱ�� Is Null And b.����id = c.����id And b.��ҳid = c.��ҳid And c.��ʽid = ��ʽid_In And
          c.Id = d.�ļ�id And d.Id = e.��¼id And e.��¼���� = 1 And e.��Ŀ��� = 1 And
          Length(Translate(e.��¼����, '-.0123456789' || e.��¼����, '-.0123456789')) = Length(e.��¼����) And
          Zl_To_Number(e.��¼����) >= 37.5 And e.��ֹ�汾 Is Null And d.����ʱ�� Between Begintime_In And Endtime_In;

  --Σ/�ز��� 
  Cursor c_List_Wz Is
    Select b.����id, b.��ҳid
    From ������Ϣ A, ������ҳ B, ��Ժ���� F
    Where a.����id = b.����id And a.��ҳid = b.��ҳid And a.����id = f.����id And Nvl(b.��ҳid, 0) <> 0 And f.����id = ����id_In And
          Nvl(b.����״̬, 0) <> 5 And b.���ʱ�� Is Null And Instr(',' || 'Σ,��' || ',', ',' || b.��ǰ���� || ',') > 0;

  --ת�������ڵĲ��� 
  Cursor c_List_Zr Is
    Select b.����id, b.��ҳid
    From ������Ϣ A, ������ҳ B, ���˱䶯��¼ C, ��Ժ���� F
    Where a.����id = b.����id And a.��ҳid = b.��ҳid And Nvl(b.��ҳid, 0) <> 0 And b.����id = c.����id And b.��ҳid = c.��ҳid And
          a.����id = f.����id And f.����id = ����id_In And Nvl(b.����״̬, 0) <> 5 And b.���ʱ�� Is Null And Nvl(c.���Ӵ�λ, 0) = 0 And
          c.����id + 0 = f.����id And c.��ʼԭ�� In (3, 15) And c.��ʼʱ�� Is Not Null And b.״̬ = 0 And c.��ʼʱ�� Between Begintime_In And
          Endtime_In;

  -- һ�������ϻ���ȼ��Ĳ��� 
  Cursor c_List_Yj Is
    Select b.����id, b.��ҳid
    From ������Ϣ A, ������ҳ B, ��Ժ���� F
    Where a.����id = b.����id And a.��ҳid = b.��ҳid And a.����id = f.����id And Nvl(b.��ҳid, 0) <> 0 And f.����id = ����id_In And
          Nvl(b.����״̬, 0) <> 5 And b.���ʱ�� Is Null And Zl_Patittendgrade(b.����id, b.��ҳid) <= 1;

  --����������ڵĲ��� 
  Cursor c_List_Fm Is
    Select b.����id, b.��ҳid
    From ������Ϣ A, ������ҳ B, ��Ժ���� F, ���˻����ļ� C, ���˻������� D, ���˻�����ϸ E
    Where a.����id = b.����id And a.��ҳid = b.��ҳid And a.����id = f.����id And
          Nvl(b.��ҳid,
              
              0) <> 0 And f.����id = ����id_In And Nvl(b.����״̬, 0) <> 5 And b.���ʱ�� Is Null And b.����id = c.����id And
          b.��ҳid = c.��ҳid And c.��ʽid = ��ʽid_In And c.Id = d.�ļ�id And d.Id = e.��¼id And e.��¼���� = 4 And e.��Ŀ���� = '����' And
          e.��ֹ�汾 Is Null And d.����ʱ�� Between Begintime_In And Endtime_In;

  Type v_����id_Type Is Table Of ������ҳ.����id%Type;
  v_����id v_����id_Type;
  Type v_��ҳid_Type Is Table Of ������ҳ.��ҳid%Type;
  v_��ҳid v_��ҳid_Type;
Begin

  v_Str := Type_In || Split_In;

  If Typeall_In = 1 Then
    Open c_List_All;
    Fetch c_List_All Bulk Collect
      Into v_����id, v_��ҳid;
    Close c_List_All;
    For I In 1 .. v_����id.Count Loop
      Out_Rec.C1 := v_����id(I);
      Out_Rec.C2 := v_��ҳid(I);
      Pipe Row(Out_Rec);
    End Loop;
  Else
    Loop
      P := Instr(v_Str, Split_In);
      Exit When(Nvl(P, 0) = 0);
      n_Index := Trim(Substr(v_Str, 1, P - 1));
      If n_Index Is Not Null Then
        If n_Index = 0 Then
          Open c_List_Ry;
          Fetch c_List_Ry Bulk Collect
            Into v_����id, v_��ҳid;
          Close c_List_Ry;
        Elsif n_Index = 1 Then
          Open c_List_Ss;
          Fetch c_List_Ss Bulk Collect
            Into v_����id, v_��ҳid;
          Close c_List_Ss;
        Elsif n_Index = 2 Then
          Open c_List_Tw;
          Fetch c_List_Tw Bulk Collect
            Into v_����id, v_��ҳid;
          Close c_List_Tw;
        Elsif n_Index = 3 Then
          Open c_List_Wz;
          Fetch c_List_Wz Bulk Collect
            Into v_����id, v_��ҳid;
          Close c_List_Wz;
        Elsif n_Index = 4 Then
          Open c_List_Zr;
          Fetch c_List_Zr Bulk Collect
            Into v_����id, v_��ҳid;
          Close c_List_Zr;
        Elsif n_Index = 5 Then
          Open c_List_Yj;
          Fetch c_List_Yj Bulk Collect
            Into v_����id, v_��ҳid;
          Close c_List_Yj;
        Elsif n_Index = 6 Then
          Open c_List_Fm;
          Fetch c_List_Fm Bulk Collect
            Into v_����id, v_��ҳid;
          Close c_List_Fm;
        End If;
      
        For I In 1 .. v_����id.Count Loop
          Out_Rec.C1 := v_����id(I);
          Out_Rec.C2 := v_��ҳid(I);
          Pipe Row(Out_Rec);
        End Loop;
      End If;
      v_Str := Substr(v_Str, P + 1);
    End Loop;
  End If;
  Return;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Getpatient;
/
--139097:������,2019-04-23,�������۲��˴���
Create Or Replace Procedure Zl_����ҽ����¼_����
(
  ҽ��id_In     In ����ҽ����¼.Id%Type,
  Flag_In       In Number := 0,
  ҽ������_In   In ����ҽ����¼.ҽ������%Type := Null,
  ��������_In   In ����ҽ��״̬.��������%Type := Null,
  ����Ա���_In In ��Ա��.���%Type := Null,
  ����Ա����_In In ��Ա��.����%Type := Null
  --���ܣ�����סԺҽ����״̬�������Ͳ���(������������ͨ������Zl_����ҽ����¼_��������������) 
  --������ҽ��ID_IN=һ��ҽ��ID 
  --      FLAG_IN=�������ݡ�����ֹͣ��0=���ִ����ֹʱ��,1=�������е�ִ����ֹʱ�䡣 
  --      ҽ������_IN=�ù��̱��������˵���ʱ���ã����ڴ�����ʾ�� 
  --      ��������_IN=�ù��̱��������˵���ʱ���ã����ں˶Ի������ݡ�0-���˷���,n=���˾���ҽ������ 
) Is
  --����ָ��ҽ���Ĳ�����¼,��һ��ΪҪ���˵�����(״̬��������) 
  --���������˷��ͺ���Զ�ֹͣ,�ڻ��˷���ʱ�Զ�����ֹͣ���� 
  Cursor c_Rolladvice Is
    Select b.������Ա, b.����ʱ��, 0 As ���ͺ�, a.���, Null As NO, b.��������, 0 As ִ��״̬, Sysdate + Null As �״�ʱ��, Sysdate + Null As ĩ��ʱ��,
           a.�ϴ�ִ��ʱ��, a.ҽ����Ч, a.������� As ���, a.������Ŀid, Null As ����, a.����id, a.��ҳid, a.Ӥ��, 0 As ��¼����, 0 As �������, 0 As ��������id,
           a.��˱��, a.����ҽ��, a.ִ�п���id, Nvl(a.���id, a.Id) As ��id, a.���id, a.Id As ҽ��id, -null As ��������, Null As ��������
    From ����ҽ����¼ A, ����ҽ��״̬ B
    Where a.Id = b.ҽ��id And (a.Id = ҽ��id_In Or a.���id = ҽ��id_In) And
          (Nvl(a.ҽ����Ч, 0) = 0 And b.�������� Not In (1, 2, 3) Or Nvl(a.ҽ����Ч, 0) = 1 And b.�������� Not In (1, 2, 3, 8))
    Union
    Select b.������ As ������Ա, b.����ʱ�� As ����ʱ��, b.���ͺ�, a.���, b.No, -null As ��������, b.ִ��״̬, b.�״�ʱ��, b.ĩ��ʱ��, a.�ϴ�ִ��ʱ��, a.ҽ����Ч,
           c.���, a.������Ŀid, c.�������� As ����, a.����id, a.��ҳid, a.Ӥ��, b.��¼����, b.�������, a.��������id, a.��˱��, a.����ҽ��, a.ִ�п���id,
           Nvl(a.���id, a.Id) As ��id, a.���id, a.Id As ҽ��id, b.��������, b.��������
    From ����ҽ����¼ A, ����ҽ������ B, ������ĿĿ¼ C
    Where a.Id = b.ҽ��id And a.������Ŀid = c.Id And (a.Id = ҽ��id_In Or a.���id = ҽ��id_In)
    Order By ����ʱ�� Desc, ���ͺ�, ���;
  r_Rolladvice c_Rolladvice%RowType;

  --��ʽͬc_Rolladvice��ֻȡ���Ͳ��������Զ����˴��� 
  Cursor c_Rollsend(v_���ͺ� ����ҽ������.���ͺ�%Type) Is
    Select Distinct b.ҽ��id, b.����ʱ�� As ����ʱ��, b.���ͺ�, b.ִ��״̬, a.������� As ���, c.��ǰ����id As ���˲���id, a.���˿���id,
                    b.ִ�в���id As ִ�п���id
    From ����ҽ����¼ A, ����ҽ������ B, ������ҳ C
    Where a.Id = b.ҽ��id And b.���ͺ� = v_���ͺ� And a.����id = c.����id And a.��ҳid = c.��ҳid And
          (a.Id = ҽ��id_In Or a.���id = ҽ��id_In)
    Order By b.����ʱ�� Desc, b.���ͺ�;

  --����ҽ��������NO������λ���Ҫ���ʵķ��ü�¼ 
  --һ��ҽ�������Ƕ���д�˷��ͼ�¼,�ҿ���NO��ͬ(ҩƷ��,�÷��巨��һ����) 
  --���ܷ��ͼ�¼�ļƷ�״̬(��������Ʒ�),�з��ü�¼��Ȼ�������� 
  --����ֻ��۸񸸺�Ϊ�յ�,�Ա�ȡ������� 
  --ֻ�ܼ�¼״̬Ϊ1�ķ���,���������ʻ򲿷����ʵļ�¼,���ٴ�������"��¼״̬=3"�Ķ�ȡ�����������жϣ������� 
  Cursor c_Rollmoneyout
  (
    v_���ͺ�    ����ҽ������.���ͺ�%Type,
    v_ҽ��id    ����ҽ����¼.Id%Type,
    t_Adviceids t_Numlist
  ) Is
    Select /*+ Rule*/
     a.Id, a.��¼״̬, a.No, a.���, a.�շ����, a.ִ��״̬, d.��������, a.ִ�в���id, a.��¼����
    From ������ü�¼ A, Table(t_Adviceids) B, ����ҽ������ C, �������� D
    Where c.ҽ��id = b.Column_Value And c.���ͺ� = v_���ͺ� And a.ҽ����� = b.Column_Value And
          (a.ҽ����� = v_ҽ��id Or Nvl(v_ҽ��id, 0) = 0) And a.��¼״̬ In (0, 1, 3) And a.No = c.No And a.��¼���� = c.��¼���� And
          a.�۸񸸺� Is Null And a.�շ�ϸĿid = d.����id(+)
    Order By a.No, a.���;

  Cursor c_Rollmoneyin
  (
    v_���ͺ�    ����ҽ������.���ͺ�%Type,
    v_ҽ��id    ����ҽ����¼.Id%Type,
    t_Adviceids t_Numlist
  ) Is
    Select /*+ Rule*/
     a.Id, a.��¼״̬, a.No, a.���, a.�շ����, a.ִ��״̬, d.��������, a.ִ�в���id, a.��¼����
    From סԺ���ü�¼ A, Table(t_Adviceids) B, ����ҽ������ C, �������� D
    Where c.ҽ��id = b.Column_Value And c.���ͺ� = v_���ͺ� And a.ҽ����� = b.Column_Value And
          (a.ҽ����� = v_ҽ��id Or Nvl(v_ҽ��id, 0) = 0) And a.��¼״̬ In (0, 1, 3) And a.No = c.No And a.��¼���� = c.��¼���� And
          a.�۸񸸺� Is Null And a.�շ�ϸĿid = d.����id(+)
    Order By a.No, a.���;

  --ȡ����סԺ����ʱ�Զ����ŵ�����(��û�����ϵ�) 
  Cursor c_Stuff_Drug(v_����id ҩƷ�շ���¼.����id%Type) Is
    Select ID
    From ҩƷ�շ���¼
    Where ����id = v_����id And (��¼״̬ = 1 Or Mod(��¼״̬, 3) = 0) And ����� Is Not Null
    Order By ҩƷid;

  --���ڴ�������ҽ���Ļ��� 
  Cursor c_Patilog
  (
    v_����id ���˱䶯��¼.����id%Type,
    v_��ҳid ���˱䶯��¼.��ҳid%Type
  ) Is
    Select *
    From ���˱䶯��¼
    Where ����id = v_����id And ��ҳid = v_��ҳid And ��ֹʱ�� Is Null
    Order By ��ʼʱ�� Desc;
  r_Patilog c_Patilog%RowType;

  Cursor c_Adviceids Is
    Select ID From ����ҽ����¼ Where ID = ҽ��id_In Or ���id = ҽ��id_In;
  t_Adviceids t_Numlist;

  v_ҽ��״̬     ����ҽ����¼.ҽ��״̬%Type;
  v_ҽ����Ч     ����ҽ����¼.ҽ����Ч%Type;
  v_����no       ����ҽ������.No%Type;
  v_�������     Varchar2(255);
  v_ĩ��ʱ��     ����ҽ������.ĩ��ʱ��%Type;
  v_����ʱ��     ����ҽ��״̬.����ʱ��%Type;
  v_��������     ������ĿĿ¼.��������%Type;
  v_ִ��Ƶ��     ������ĿĿ¼.ִ��Ƶ��%Type;
  v_�ϴ�ʱ��     ����ҽ����¼.�ϴ�ִ��ʱ��%Type;
  v_ִ��ʱ��     ����ҽ����¼.ִ��ʱ�䷽��%Type;
  v_��ʼִ��ʱ�� ����ҽ����¼.��ʼִ��ʱ��%Type;
  v_�ϴδ�ӡʱ�� ����ҽ����¼.�ϴδ�ӡʱ��%Type;
  v_Ƶ�ʼ��     ����ҽ����¼.Ƶ�ʼ��%Type;
  v_�����λ     ����ҽ����¼.�����λ%Type;
  v_���ͺ�       ����ҽ������.���ͺ�%Type;
  n_����ȼ�id   ���˱䶯��¼.����ȼ�id%Type;
  d_��ʼʱ��     ���˱䶯��¼.��ʼʱ��%Type;
  d_����ʱ��     ����ҽ��״̬.����ʱ��%Type;
  v_Tmp���ͺ�    ����ҽ������.���ͺ�%Type;
  n_ִ��         Number;

  Intdigit   Number(3);
  v_Update   Number(1);
  v_Count    Number(5);
  v_Temp     Varchar2(2000);
  v_��Ա��� ��Ա��.���%Type;
  v_��Ա���� ��Ա��.����%Type;
  v_Time     Varchar2(4000);
  n_Blndo    Number;

  v_Error Varchar2(2000);
  Err_Custom Exception;

  Function Checkmoneyundo
  (
    v_No       סԺ���ü�¼.No%Type,
    v_��¼���� סԺ���ü�¼.��¼����%Type,
    v_���     סԺ���ü�¼.���%Type,
    n_����     Number := 0 --0סԺ��1���� 
  ) Return Number Is
    n_Num      Number;
    n_ִ��״̬ Number;
  Begin
    n_Num := 0;
    If n_���� = 0 Then
      Select Nvl(Sum(Nvl(����, 1) * ����), 0) As ����
      Into n_Num
      From סԺ���ü�¼
      Where NO = v_No And ��¼���� = v_��¼���� And ��� = v_��� And ��¼״̬ In (2, 3);
      Select Nvl(ִ��״̬, 0)
      Into n_ִ��״̬
      From סԺ���ü�¼
      Where NO = v_No And ��¼���� = v_��¼���� And ��� = v_��� And ��¼״̬ = 3;
    Else
      Select Nvl(Sum(Nvl(����, 1) * ����), 0) As ����
      Into n_Num
      From ������ü�¼
      Where NO = v_No And ��¼���� = v_��¼���� And ��� = v_��� And ��¼״̬ In (2, 3);
      Select Nvl(ִ��״̬, 0)
      Into n_ִ��״̬
      From ������ü�¼
      Where NO = v_No And ��¼���� = v_��¼���� And ��� = v_��� And ��¼״̬ = 3;
    End If;
    If n_Num <> 0 Then
      n_Num := 1;
    End If;
    --�������¼����ִ�У�����ִ�еģ����Զ��ˡ� 
    If n_ִ��״̬ <> 0 Then
      n_Num := 0;
    End If;
    Return(n_Num);
  End;
Begin
  v_Tmp���ͺ� := -1;
  Open c_Rolladvice;
  Loop
    Fetch c_Rolladvice
      Into r_Rolladvice;
    If c_Rolladvice%RowCount = 0 Then
      Close c_Rolladvice;
      v_Error := Nvl(ҽ������_In, '��ҽ��') || '��ǰû�п��Ի��˵����ݡ�';
      Raise Err_Custom;
    End If;
    Exit When c_Rolladvice%NotFound;
    Exit When d_����ʱ�� <> r_Rolladvice.����ʱ�� And d_����ʱ�� Is Not Null;
    d_����ʱ�� := r_Rolladvice.����ʱ��;
  
    --�������˵���ʱ�ж� 
    If ҽ������_In Is Not Null Then
      If Nvl(r_Rolladvice.��������, 0) <> Nvl(��������_In, 0) Then
        v_Error := Nvl(ҽ������_In, '��ҽ��') || '�����뵱ǰҽ��һ����ˣ����ܸ�ҽ���Ѿ�ִ��������������';
        Raise Err_Custom;
      End If;
    End If;
  
    --һ�鷢�ͺ�ִֻ��һ�� 
    If v_Tmp���ͺ� <> r_Rolladvice.���ͺ� Then
      v_Tmp���ͺ� := r_Rolladvice.���ͺ�;
      n_ִ��      := 1;
    Else
      n_ִ�� := 0;
    End If;
  
    If n_ִ�� = 1 Then
      Open c_Adviceids;
      Fetch c_Adviceids Bulk Collect
        Into t_Adviceids;
      Close c_Adviceids;
    
      If r_Rolladvice.���ͺ� = 0 Then
        --����ҽ��״̬����(��ʱ��ؼ���) 
        --4-���ϣ�5-������6-��ͣ��7-���ã�8-ֹͣ��9-ȷ��ֹͣ��10-Ƥ�Խ��;13-ͣ������ 
        ------------------------------------------------------------------ 
        --���ֻ���˻ص�У��״̬ 
        If r_Rolladvice.�������� = 3 Then
          v_Error := Nvl(ҽ������_In, '��ҽ��') || '��ǰ����ͨ��У��״̬�������ٻ��ˡ�';
          Raise Err_Custom;
        Elsif r_Rolladvice.�������� = 4 And Nvl(r_Rolladvice.Ӥ��, 0) = 0 Then
          If r_Rolladvice.��� = 'H' Then
            Select ��������, ִ��Ƶ�� Into v_��������, v_ִ��Ƶ�� From ������ĿĿ¼ Where ID = r_Rolladvice.������Ŀid;
            If v_�������� = '1' And v_ִ��Ƶ�� = '2' Then
              v_Error := '����ȼ����Ϻ����ٻ��ˡ�';
              Raise Err_Custom;
            End If;
          End If;
        End If;
      
        --����Ƿ�������������֮ǰ�Ĳ��� 
        If r_Rolladvice.�������� <> 5 Then
          --ȡ�������ʱ�� 
          Select Nvl(ҽ������ʱ��, To_Date('1900-01-01', 'YYYY-MM-DD'))
          Into v_����ʱ��
          From ������ҳ
          Where ����id = r_Rolladvice.����id And ��ҳid = r_Rolladvice.��ҳid;
        
          If r_Rolladvice.����ʱ�� < v_����ʱ�� Then
            v_Error := '�ò������������֮ǰ�Ĳ��������ٻ��ˡ�';
            Raise Err_Custom;
          End If;
        End If;
      
        --ɾ��(����ҽ��)�����״̬������¼ 
        Delete /*+ Rule*/
        From ����ҽ��״̬
        Where ҽ��id In (Select Column_Value From Table(t_Adviceids)) And ����ʱ�� = r_Rolladvice.����ʱ��;
      
        --ȡɾ����Ӧ�ָ���ҽ��״̬ 
        Select ��������
        Into v_ҽ��״̬
        From ����ҽ��״̬
        Where ����ʱ�� = (Select Max(����ʱ��) From ����ҽ��״̬ Where ҽ��id = ҽ��id_In) And ҽ��id = ҽ��id_In;
      
        --�ָ�(����ҽ��)���˺��״̬ 
        Update ����ҽ����¼ Set ҽ��״̬ = v_ҽ��״̬ Where ID = ҽ��id_In Or ���id = ҽ��id_In;
      
        --��������Ĵ��� 
        If r_Rolladvice.�������� = 8 Then
          --�����ڷ����ջع���ҽ�� ���������������ģʽ�����ж϶�Ӧ�ġ����˷������ʡ������Ƿ�ȡ��������������ˣ��������� 
          --                       ����ǲ�����������ģʽ���������ٻ��ˡ� 
          --���ܳ��ڷ����ջ�ʱ��ȫ���ջ�(���ϴ�ִ��ʱ��) 
          Select /*+ Rule*/
           Nvl(Count(*), 0)
          Into v_Count
          From ����ҽ����¼ A, ����ҽ������ B
          Where b.ҽ��id = a.Id And (a.Id = ҽ��id_In Or a.���id = ҽ��id_In) And
                b.���ͺ� =
                (Select Max(���ͺ�) From ����ҽ������ Where ҽ��id In (Select Column_Value From Table(t_Adviceids))) And
                a.ִ����ֹʱ�� Is Not Null And ((a.�ϴ�ִ��ʱ�� < b.ĩ��ʱ��) Or (a.�ϴ�ִ��ʱ�� Is Null And b.ĩ��ʱ�� Is Not Null));
          If v_Count > 0 Then
            If zl_GetSysParameter('�����ջز�����������', 1254) = '1' Then
              v_Error := Nvl(ҽ������_In, '��ҽ��') || '�ѱ����ڷ����ջأ������ٳ���ֹͣ������';
              Raise Err_Custom;
            Else
              --����Ѿ�ȡ���������룬���������. 
              Select Count(1)
              Into v_Count
              From ���˷������� A, סԺ���ü�¼ B, ����ҽ����¼ C
              Where a.����id = b.Id And c.Id = b.ҽ����� And (c.Id = ҽ��id_In Or c.���id = ҽ��id_In);
              If v_Count > 0 Then
                v_Error := Nvl(ҽ������_In, '��ҽ��') || '�ѱ����ڷ����ջأ������ٳ���ֹͣ������';
                Raise Err_Custom;
              Else
                --�õ��ϴ�ִ��ʱ�����Ϣ 
                Select �ϴ�ִ��ʱ��, ִ��ʱ�䷽��, ��ʼִ��ʱ��, �ϴδ�ӡʱ��, Ƶ�ʼ��, �����λ
                Into v_�ϴ�ʱ��, v_ִ��ʱ��, v_��ʼִ��ʱ��, v_�ϴδ�ӡʱ��, v_Ƶ�ʼ��, v_�����λ
                From ����ҽ����¼
                Where ID = ҽ��id_In;
                v_�ϴ�ʱ�� := To_Date(To_Char(v_�ϴ�ʱ�� + 1 / 24 / 60 / 60, 'yyyy-MM-dd hh24:mi:ss'), 'yyyy-MM-dd hh24:mi:ss');
              
                --�޸��ϴ�ִ��ʱ��Ϊ�ջغ��ĩ��ִ��ʱ�䡣 
                v_ĩ��ʱ�� := Null;
                Begin
                  --һ��ҽ���ķ�����ĩʱ����ͬ,һ����ҩ��ȡ��С�� 
                  --ȡ���IDΪNULL��ҽ���ķ��ͼ�¼��ʱ�� 
                  --����ҩ;������ҩ�÷�����δ��д���ͼ�¼ 
                  Select /*+ Rule*/
                   ĩ��ʱ��, ���ͺ�
                  Into v_ĩ��ʱ��, v_���ͺ�
                  From ����ҽ������
                  Where ҽ��id In (Select Column_Value From Table(t_Adviceids)) And
                        ���ͺ� = (Select Max(���ͺ�)
                               From ����ҽ������
                               Where ҽ��id In (Select Column_Value From Table(t_Adviceids))) And Rownum = 1;
                Exception
                  When Others Then
                    Null;
                End;
                Update ����ҽ����¼ Set �ϴ�ִ��ʱ�� = v_ĩ��ʱ�� Where ID = ҽ��id_In Or ���id = ҽ��id_In;
              
                Select Count(1) Into v_Count From ����ҽ������ Where ҽ��id = ҽ��id_In And ���ͺ� = v_���ͺ�;
                If v_Count > 0 Then
                  --��ԭҽ��ִ��ʱ�� 
                  Select Zl_Adviceexetimes(ҽ��id_In, v_�ϴ�ʱ��, v_ĩ��ʱ��, v_ִ��ʱ��, v_��ʼִ��ʱ��, v_�ϴδ�ӡʱ��, v_Ƶ�ʼ��, v_�����λ, 0)
                  Into v_Time
                  From Dual;
                  Insert Into ҽ��ִ��ʱ��
                    (Ҫ��ʱ��, ҽ��id, ���ͺ�)
                    Select To_Date(Column_Value, 'yyyy-mm-dd hh24:mi:ss'), ҽ��id_In, v_���ͺ�
                    From Table(f_Str2list(v_Time));
                End If;
              End If;
            End If;
          End If;
        
          --����ȼ��䶯�������������䶯ʱ����������� 
          If r_Rolladvice.��� = 'H' And Nvl(r_Rolladvice.Ӥ��, 0) = 0 Then
            Select ��������, ִ��Ƶ�� Into v_��������, v_ִ��Ƶ�� From ������ĿĿ¼ Where ID = r_Rolladvice.������Ŀid;
            If v_�������� = '1' And v_ִ��Ƶ�� = '2' Then
              Select Count(*), Max(a.����ȼ�id), Max(a.��ʼʱ��)
              Into v_Count, n_����ȼ�id, d_��ʼʱ��
              From ���˱䶯��¼ A
              Where a.����id = r_Rolladvice.����id And a.��ҳid = r_Rolladvice.��ҳid And a.��ʼԭ�� = 6 And a.��ֹʱ�� Is Null And
                    a.���Ӵ�λ = 0;
              --���û���ҵ����һ���ǻ���ȼ��䶯���ֹ 
              If v_Count = 0 Then
                --ҽ������ȼ�����סʱ��Ļ���ȼ�һ��ʱҪ�����ж� 
                Select Count(*)
                Into v_Count
                From ���˱䶯��¼ A
                Where a.����id = r_Rolladvice.����id And a.��ҳid = r_Rolladvice.��ҳid And a.��ʼԭ�� = 6;
                If v_Count > 0 Then
                  v_Error := '���ڻ���ȼ�ҽ��ֹͣ��ò����Ѿ������������䶯��¼,���ܻ��˸�ҽ����ֹͣ������';
                  Raise Err_Custom;
                End If;
              Else
                --���n_����ȼ�IDΪNull�������Ƿ��ǵ�ǰ���˵�ҽ����Ӧ�ı䶯��¼,Ŀ�����ж������ȼ�ҽ��ʱҪ��˳����ˡ� 
                --���n_����ȼ�ID��ΪNull�����п�����У����һ������ȼ�ʱ���Զ�ֹͣ�ģ�δ�����䶯��¼�� 
                --     ����Ҫ��鵱ǰ���һ���䶯�Ļ���ȼ�ID�Ƿ��ǵ�ǰҽ���Ļ���ȼ�ID,Ŀ�����ж������ȼ�ҽ��ʱҪ��˳����ˣ����������Ҫ�ٳ������һ�α䶯��ֱ�ӻ���ҽ�����ɡ� 
                If n_����ȼ�id Is Null Then
                  Select Count(*)
                  Into v_Count
                  From ���˱䶯��¼ B, ����ҽ���Ƽ� C
                  Where b.����id = r_Rolladvice.����id And b.��ҳid = r_Rolladvice.��ҳid And c.ҽ��id = ҽ��id_In And
                        c.�շ�ϸĿid = b.����ȼ�id And b.��ֹʱ�� = d_��ʼʱ�� And b.��ֹԭ�� = 6 And b.���Ӵ�λ = 0;
                Else
                  --��ʼʱ��ֻȡ���ӶԱȣ�У�Ե�ʱ����ȼ��Ŀ�ʼʱ����ҽ����ʼʱ��+��ǰʱ������� 
                  Select Count(*)
                  Into v_Count
                  From ����ҽ���Ƽ� C, ����ҽ����¼ A
                  Where a.Id = c.ҽ��id And a.Id = ҽ��id_In And c.�շ�ϸĿid = n_����ȼ�id And
                        a.��ʼִ��ʱ�� = To_Date(To_Char(d_��ʼʱ��, 'yyyy-mm-dd hh24:mi'), 'yyyy-mm-dd hh24:mi');
                End If;
                If v_Count = 0 Then
                  v_Error := '�����˵�ҽ���������һ������ȼ�ҽ�����뽫����Ļ���ȼ�ҽ�����Ϻ��ٻ��˱���ҽ����';
                  Raise Err_Custom;
                End If;
              
                If n_����ȼ�id Is Null Then
                  --��ǰ������Ա 
                  If ����Ա���_In Is Not Null And ����Ա����_In Is Not Null Then
                    v_��Ա��� := ����Ա���_In;
                    v_��Ա���� := ����Ա����_In;
                  Else
                    v_Temp     := Zl_Identity;
                    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
                    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
                    v_��Ա��� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
                    v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
                  End If;
                
                  Zl_���˱䶯��¼_Undo(r_Rolladvice.����id, r_Rolladvice.��ҳid, v_��Ա���, v_��Ա����, '1', Null, Null, '����ȼ��䶯');
                End If;
              End If;
            End If;
          End If;
        
          If r_Rolladvice.��� = 'Z' And Instr(',9,10,', ',' || r_Rolladvice.���� || ',') > 0 And
             Nvl(r_Rolladvice.Ӥ��, 0) = 0 Then
            --���˲���ҽ��ʱ�����ñ䶯��¼���� 
            Zl_���˱䶯��¼_Undo(r_Rolladvice.����id, r_Rolladvice.��ҳid, v_��Ա���, v_��Ա����, Null, Null, Null, '�����䶯');
          End If;
        
          --����ҽ��ֹͣʱ,���ͣ��ҽ����ʱ��,�����ʵϰҽʦ�������˵ģ���ָ������״̬ 
          Update ����ҽ����¼
          Set ִ����ֹʱ�� = Decode(Flag_In, 1, ִ����ֹʱ��, Null), ͣ��ҽ�� = Null, ͣ��ʱ�� = Null,
              ��˱�� = Decode(r_Rolladvice.��˱��, 3, 2, r_Rolladvice.��˱��)
          Where ID = ҽ��id_In Or ���id = ҽ��id_In;
        Elsif r_Rolladvice.�������� = 9 Then
          --����ҽ��ȷ��ֹͣʱ,����Ƿ��Ѵ�ӡͣ��ʱ�� 
          Select /*+ Rule*/
           Count(*)
          Into v_Count
          From ����ҽ����ӡ
          Where ��ӡ��� = 1 And ҽ��id In (Select Column_Value From Table(t_Adviceids));
          If v_Count > 0 Then
            v_Error := Nvl(ҽ������_In, '��ҽ��') || '��ͣ��ʱ���Ѿ���ӡ�������ٳ���ȷ��ֹͣ������';
            Raise Err_Custom;
          End If;
        
          --����ҽ��ȷ��ֹͣʱ,���ͣ��ҽ����ʱ�� 
          Update ����ҽ����¼ Set ȷ��ͣ��ʱ�� = Null, ȷ��ͣ����ʿ = Null Where ID = ҽ��id_In Or ���id = ҽ��id_In;
        Elsif r_Rolladvice.�������� = 10 Then
          --���˱�עƤ�Խ��,ͬʱɾ�������Ǽ�(+)��(-),���ݼ�¼ʱ�� 
          --�����ļ�¼��ҽ�������޹ۣ�����Ҫ���� 
          Delete From ���˹�����¼
          Where ����id = r_Rolladvice.����id And Nvl(��ҳid, 0) = Nvl(r_Rolladvice.��ҳid, 0) And ��¼ʱ�� = r_Rolladvice.����ʱ�� And
                Nvl(���, 0) = 0;
        
          Update ����ҽ����¼ Set Ƥ�Խ�� = Null Where ID = ҽ��id_In Or ���id = ҽ��id_In;
        Elsif r_Rolladvice.�������� = 13 Then
          If Instr(r_Rolladvice.����ҽ��, '/') > 0 Then
            Update ����ҽ����¼ Set ��˱�� = 1 Where ID = ҽ��id_In Or ���id = ҽ��id_In;
          Else
            Update ����ҽ����¼ Set ��˱�� = Null Where ID = ҽ��id_In Or ���id = ҽ��id_In;
          End If;
        End If;
        --��������ҽ�����ϲ��� 
        --����ҽ�� 
        If r_Rolladvice.�������� = 4 And r_Rolladvice.��� = 'Z' Then
          Select Count(1) Into v_Count From ������ĿĿ¼ Where ID = r_Rolladvice.������Ŀid And �������� = '4';
          If v_Count = 1 Then
            b_Message.Zlhis_Cis_004(r_Rolladvice.����id, r_Rolladvice.��ҳid, ҽ��id_In);
          End If;
        End If;
      Else
        --����ҽ������(�Է��ͺŹؼ���) 
        ------------------------------------------------------------------ 
        --��ǰ������Ա 
        v_Temp     := Zl_Identity;
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
        v_��Ա��� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
        v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
      
        --����Ƿ�����Һ��Һ��¼�����Ƿ��Ѿ������������ѯ������˵������Һ��¼ 
        Begin
          Select Decode(Max(�Ƿ�����), 1, 1, 0)
          Into v_Count
          From ��Һ��ҩ��¼
          Where ҽ��id = ҽ��id_In And ���ͺ� = r_Rolladvice.���ͺ�;
        Exception
          When Others Then
            v_Count := -1;
        End;
      
        If v_Count = 1 Then
          v_Error := 'ҽ��"' || ҽ������_In || '"����ҺҩƷ���Ѿ�����Һ�����������������ܻ��˷��͡�';
          Raise Err_Custom;
        Elsif v_Count = 0 Then
          Zl_��Һ��ҩ��¼_ҽ������(ҽ��id_In, r_Rolladvice.���ͺ�, v_��Ա����, Sysdate);
        End If;
      
        --����Ƿ����δ��˵��������� 
        Select Count(*)
        Into v_Count
        From ����ҽ����¼ A, ����ҽ������ B, סԺ���ü�¼ C, ���˷������� D
        Where (a.Id = ҽ��id_In Or a.���id = ҽ��id_In) And a.Id = b.ҽ��id And b.ҽ��id = c.ҽ����� And c.Id = d.����id And
              c.��¼״̬ In (0, 1, 3) And d.״̬ = 0;
      
        If v_Count > 0 Then
          v_Error := 'ҽ��"' || ҽ������_In || '"����δ��˵��������룬��ȡ�����������������ٻ��˷��͡�';
          Raise Err_Custom;
        End If;
      
        --���ҽ���Ƿ������Ч��ҽ������ 
        Select Count(*)
        Into v_Count
        From ����ҽ������ A, סԺ���ü�¼ B
        Where a.ҽ��id = b.ҽ����� And a.No = b.No And b.��¼״̬ = 1 And b.ʵ�ս�� <> 0 And a.���ͺ� = r_Rolladvice.���ͺ� And
              a.ҽ��id In (Select Column_Value From Table(t_Adviceids));
        If v_Count > 0 Then
          v_Error := '��ҽ���»����ڸ�����Ŀ�����ȳ�����';
          Raise Err_Custom;
        End If;
      
        --���Ʒ����Զ�ִ��ʱ������Ҳ�Զ�����ִ��(����ʿվ�д˹���) 
        --�Ǹ������õ�����ҽ����ͬ��ͨҽ��ִ�д��� 
        Select ҽ����Ч Into v_ҽ����Ч From ����ҽ����¼ Where ID = ҽ��id_In;
        If Substr(zl_GetSysParameter('����ִ���Զ����', 1254), v_ҽ����Ч + 1, 1) = '1' Then
          Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into Intdigit From Dual;
        
          For r_Rollsend In c_Rollsend(r_Rolladvice.���ͺ�) Loop
            If Nvl(r_Rollsend.ִ��״̬, 0) = 1 And
               (Nvl(r_Rollsend.ִ�п���id, 0) = Nvl(r_Rollsend.���˲���id, 0) Or
                Nvl(r_Rollsend.ִ�п���id, 0) = Nvl(r_Rollsend.���˿���id, 0)) Then
            
              --ҽ����ִ��״̬ 
              Update ����ҽ������ Set ִ��״̬ = 0 Where ���ͺ� = r_Rollsend.���ͺ� And ҽ��id = r_Rollsend.ҽ��id;
              v_Update := 1;
            
              If r_Rolladvice.��¼���� = 2 And Nvl(r_Rolladvice.�������, 0) = 0 Then
                --���õ�ִ��״̬ 
                For r_Rollmoney In c_Rollmoneyin(r_Rollsend.���ͺ�, r_Rollsend.ҽ��id, t_Adviceids) Loop
                  n_Blndo := 0;
                  If r_Rollmoney.��¼״̬ <> 3 Then
                    n_Blndo := 1;
                  Else
                    n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.��¼����, r_Rollmoney.���);
                  End If;
                  If n_Blndo > 0 Then
                    If Not (r_Rollmoney.�շ���� = '4' And Nvl(r_Rollmoney.��������, 0) = 1) And
                       Not r_Rollmoney.�շ���� In ('5', '6', '7') Then
                      --��ͨ����ֱ��ȡ��ִ��״̬������ҩƷ�͸������õ����� 
                      Update סԺ���ü�¼
                      Set ִ��״̬ = 0, ִ��ʱ�� = Null, ִ���� = Null
                      Where NO = r_Rollmoney.No And ��¼���� = r_Rolladvice.��¼���� And ��¼״̬ = r_Rollmoney.��¼״̬ And
                            Nvl(�۸񸸺�, ���) = r_Rollmoney.��� And ҽ����� = r_Rollsend.ҽ��id;
                    Elsif r_Rollmoney.�շ���� = '4' And Nvl(r_Rollmoney.��������, 0) = 1 Then
                      --�������õ����ģ����Զ����� 
                      For r_Stuff In c_Stuff_Drug(r_Rollmoney.Id) Loop
                        Zl_�����շ���¼_��������(r_Stuff.Id, v_��Ա����, Sysdate, Null, Null, Null, Null, 0, v_��Ա����);
                      End Loop;
                    Elsif r_Rollmoney.�շ���� In ('5', '6', '7') Then
                      --סԺ���ҷ�ҩ��ҩƷ�Զ���ҩ 
                      If r_Rollmoney.ִ�в���id = r_Rollsend.���˲���id Or r_Rollmoney.ִ�в���id = r_Rollsend.���˿���id Then
                        For r_Drug In c_Stuff_Drug(r_Rollmoney.Id) Loop
                          Zl_ҩƷ�շ���¼_������ҩ(r_Drug.Id, v_��Ա����, Sysdate, Null, Null, Null, Null, Null, Null, Intdigit, 2);
                        End Loop;
                      End If;
                    End If;
                  End If;
                End Loop;
              Else
                --סԺ���˷��÷��͵�����������������Դ����סԺ�� 
                For r_Rollmoney In c_Rollmoneyout(r_Rollsend.���ͺ�, r_Rollsend.ҽ��id, t_Adviceids) Loop
                  n_Blndo := 0;
                  If r_Rollmoney.��¼״̬ <> 3 Then
                    n_Blndo := 1;
                  Else
                    n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.��¼����, r_Rollmoney.���, 1);
                  End If;
                  If n_Blndo > 0 Then
                    If Not (r_Rollmoney.�շ���� = '4' And Nvl(r_Rollmoney.��������, 0) = 1) And
                       Not r_Rollmoney.�շ���� In ('5', '6', '7') Then
                      --��ͨ����ֱ��ȡ��ִ��״̬������ҩƷ�͸������õ����� 
                      Update ������ü�¼
                      Set ִ��״̬ = 0, ִ��ʱ�� = Null, ִ���� = Null
                      Where NO = r_Rollmoney.No And ��¼���� = r_Rolladvice.��¼���� And ��¼״̬ = r_Rollmoney.��¼״̬ And
                            Nvl(�۸񸸺�, ���) = r_Rollmoney.��� And ҽ����� = r_Rollsend.ҽ��id;
                    Elsif r_Rollmoney.�շ���� = '4' And Nvl(r_Rollmoney.��������, 0) = 1 Then
                      --�������õ����ģ����Զ����� 
                      For r_Stuff In c_Stuff_Drug(r_Rollmoney.Id) Loop
                        Zl_�����շ���¼_��������(r_Stuff.Id, v_��Ա����, Sysdate, Null, Null, Null, Null, 0, v_��Ա����);
                      End Loop;
                    Elsif r_Rollmoney.�շ���� In ('5', '6', '7') Then
                      --�����ҷ�ҩ��ҩƷ�Զ���ҩ 
                      If r_Rollmoney.ִ�в���id = r_Rollsend.���˲���id Or r_Rollmoney.ִ�в���id = r_Rollsend.���˿���id Then
                        For r_Drug In c_Stuff_Drug(r_Rollmoney.Id) Loop
                          Zl_ҩƷ�շ���¼_������ҩ(r_Drug.Id, v_��Ա����, Sysdate, Null, Null, Null, Null, Null, Null, Intdigit, 1);
                        End Loop;
                      End If;
                    End If;
                  End If;
                End Loop;
              End If;
            End If;
          End Loop;
        End If;
        ------------------------------------------------------------------ 
        --�������ջصĳ���ҩƷҽ�����������(���˷��þͶ�����) 
        If Nvl(r_Rolladvice.ҽ����Ч, 0) = 0 Then
          If r_Rolladvice.�ϴ�ִ��ʱ�� Is Not Null And r_Rolladvice.ĩ��ʱ�� Is Not Null Then
            If r_Rolladvice.�ϴ�ִ��ʱ�� < r_Rolladvice.ĩ��ʱ�� Then
              v_Error := Nvl(ҽ������_In, '��ҽ��') || '������ڷ��͵������ѱ��ջأ������ٻ��ˡ�';
              Raise Err_Custom;
            End If;
          Elsif r_Rolladvice.�ϴ�ִ��ʱ�� Is Null And r_Rolladvice.ĩ��ʱ�� Is Not Null Then
            --�������ܱ�ȫ�������ջ� 
            v_Error := Nvl(ҽ������_In, '��ҽ��') || 'δ�����ͣ����͵������ѱ�ȫ�������ջأ������ٻ��ˡ�';
            Raise Err_Custom;
          End If;
        End If;
      
        If Nvl(r_Rolladvice.ִ��״̬, 0) In (1, 3) And v_Update <> 1 Then
          --1-��ȫִ��;3-����ִ�� 
          v_Error := Nvl(ҽ������_In, '��ҽ��') || '������͵������Ѿ�ִ�л�����ִ�У����ܻ��ˡ�';
          Raise Err_Custom;
        Else
          --������ҽ����ִ�У���ҲҪ���ƻ��ˣ����磺����Ĳɼ���ʽ�� 
          Select /*+ Rule*/
           Count(1)
          Into v_Count
          From ����ҽ������
          Where ҽ��id In (Select Column_Value From Table(t_Adviceids)) And ִ��״̬ In (1, 3) And
                ���ͺ� =
                (Select Max(���ͺ�) From ����ҽ������ Where ҽ��id In (Select Column_Value From Table(t_Adviceids)));
          If v_Count > 0 Then
            v_Error := Nvl(ҽ������_In, '��ҽ��') || '������͵������Ѿ�ִ�л�����ִ�У����ܻ��ˡ�';
            Raise Err_Custom;
          End If;
        End If;
      
        ------------------------------------------------------------------ 
        --������ҽ���ķ�������(��һ��ҽ�������в�ͬNO����) 
        --���ԭʼ�����ѱ�����(�򲿷�����),���ù��������ж� 
        v_����no   := Null;
        v_������� := Null;
        If r_Rolladvice.��¼���� = 2 And Nvl(r_Rolladvice.�������, 0) = 0 Then
          For r_Rollmoney In c_Rollmoneyin(r_Rolladvice.���ͺ�, Null, t_Adviceids) Loop
            --��Ӧ�ķ�����ִ�� 
            If Nvl(r_Rollmoney.ִ��״̬, 0) <> 0 Then
              v_Error := Nvl(ҽ������_In, '��ҽ��') || '���͵ķ��õ���"' || r_Rollmoney.No || '"�е������ѱ����ֻ���ȫִ�У����ܻ��ˡ�';
              Raise Err_Custom;
            End If;
            n_Blndo := 0;
            If r_Rollmoney.��¼״̬ <> 3 Then
              n_Blndo := 1;
            Else
              n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.��¼����, r_Rollmoney.���);
            End If;
            If n_Blndo > 0 Then
              --���ֽ������жϲ�����ҩ 
              If v_����no <> r_Rollmoney.No And v_������� Is Not Null Then
                Zl_סԺ���ʼ�¼_Delete(v_����no, Substr(v_�������, 2), v_��Ա���, v_��Ա����, 2, 0, 0);
                v_������� := Null;
              End If;
              v_����no   := r_Rollmoney.No;
              v_������� := v_������� || ',' || r_Rollmoney.���;
            End If;
          End Loop;
        Else
          For r_Rollmoney In c_Rollmoneyout(r_Rolladvice.���ͺ�, Null, t_Adviceids) Loop
            --��Ӧ�ķ�����ִ�� 
            If Nvl(r_Rollmoney.ִ��״̬, 0) <> 0 And Not (Nvl(r_Rollmoney.ִ��״̬, 0) = -1 And Nvl(r_Rollmoney.��¼״̬, 0) = 0) Then
              v_Error := Nvl(ҽ������_In, '��ҽ��') || '���͵ķ��õ���"' || r_Rollmoney.No || '"�е������ѱ����ֻ���ȫִ�У����ܻ��ˡ�';
              Raise Err_Custom;
            End If;
            --�շѵ������շ� 
            If r_Rollmoney.��¼״̬ = 1 And Not (r_Rolladvice.��¼���� = 2 And Nvl(r_Rolladvice.�������, 0) = 1) Then
              v_Error := Nvl(ҽ������_In, '��ҽ��') || '���͵����ﵥ��"' || r_Rollmoney.No || '"���շѣ����ܻ��ˡ�';
              Raise Err_Custom;
            End If;
            n_Blndo := 0;
            If r_Rollmoney.��¼״̬ <> 3 Then
              n_Blndo := 1;
            Else
              n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.��¼����, r_Rollmoney.���, 1);
            End If;
            If n_Blndo > 0 Then
              --���ֽ������жϲ�����ҩ 
              If v_����no <> r_Rollmoney.No And v_������� Is Not Null Then
                If r_Rolladvice.��¼���� = 2 And Nvl(r_Rolladvice.�������, 0) = 1 Then
                  --סԺ����Ϊ�������(���������ҽ������Ϊ������ʣ�����ҽ��û�л��˹���) 
                  Zl_������ʼ�¼_Delete(v_����no, v_�������, v_��Ա���, v_��Ա����, 0);
                Else
                  Zl_���ﻮ�ۼ�¼_Delete(v_����no, Substr(v_�������, 2));
                End If;
                v_������� := Null;
              End If;
              v_����no   := r_Rollmoney.No;
              v_������� := v_������� || ',' || r_Rollmoney.���;
            End If;
          End Loop;
        End If;
        If v_������� Is Not Null And v_����no Is Not Null Then
          v_������� := Substr(v_�������, 2);
          If r_Rolladvice.��¼���� = 2 And Nvl(r_Rolladvice.�������, 0) = 0 Then
            Zl_סԺ���ʼ�¼_Delete(v_����no, v_�������, v_��Ա���, v_��Ա����, 2, 0, 0);
          Elsif r_Rolladvice.��¼���� = 2 And Nvl(r_Rolladvice.�������, 0) = 1 Then
            --סԺ����Ϊ�������(���������ҽ������Ϊ������ʣ�����ҽ��û�л��˹���) 
            Zl_������ʼ�¼_Delete(v_����no, v_�������, v_��Ա���, v_��Ա����, 0);
          Else
            Zl_���ﻮ�ۼ�¼_Delete(v_����no, v_�������);
          End If;
        End If;
      
        --���˷��Ͳ�����ͨ�����ͼ�¼������Ϣ�����������ҽ��Ҫ��ǰ      
        For R In (Select a.����id, a.��ҳid, b.No, b.���ͺ�, b.��������, b.�״�ʱ��, b.ĩ��ʱ��, b.��������, a.Id, a.���id,
                         Nvl(a.���id, a.Id) As ��id, c.���, c.��������, a.ִ�п���id
                  From ����ҽ����¼ A, ����ҽ������ B, ������ĿĿ¼ C
                  Where a.Id = b.ҽ��id And a.������Ŀid = c.Id And b.���ͺ� = r_Rolladvice.���ͺ� And
                        b.ҽ��id In (Select Column_Value From Table(t_Adviceids))
                  Order By a.���) Loop
        
          --�˴������Ϣ����
          If r.��� = 'D' And r.���id Is Null Then
            --��� 
            b_Message.Zlhis_Cis_037(r.����id, r.��ҳid, Null, r.���ͺ�, r.��id, r.No, 2);
          Elsif r.��� = 'F' And r.���id Is Null Then
            --���� 
            b_Message.Zlhis_Cis_038(r.����id, r.��ҳid, Null, r.���ͺ�, r.��id, r.No);
          Elsif r.��� = 'K' And r.���id Is Null Then
            --��Ѫ 
            b_Message.Zlhis_Cis_039(r.����id, r.��ҳid, Null, r.���ͺ�, r.��id, r.No);
          Elsif r.��� = 'E' And r.�������� = '6' Then
            --����
            b_Message.Zlhis_Cis_036(r.����id, r.��ҳid, Null, r.���ͺ�, r.��id, r.No, 2);
          End If;
        
          Select Count(1) Into v_Count From ��������˵�� A Where a.����id = r.ִ�п���id And a.�������� = '����';
          If v_Count > 0 Then
            --����ִ��ҽ�����˷���
            b_Message.Zlhis_Cis_044(r.����id, r.��ҳid, r.���ͺ�, r.Id, r.No, r.��������, r.�״�ʱ��, r.ĩ��ʱ��, r.��������);
          End If;
        End Loop;
      
        --��Ѫҽ����ɾ������ҽ������ 
        Delete From ����ҽ������ Where ���ͺ� = r_Rolladvice.���ͺ� And ҽ��id = ҽ��id_In;
      
        --ɾ��ҽ��ִ��ʱ�� (����ҽ��ID�Ų����˼�¼) 
        Delete From ҽ��ִ��ʱ�� Where ���ͺ� = r_Rolladvice.���ͺ� And ҽ��id = ҽ��id_In;
      
        --ɾ�����ͼ�¼(����ҽ����) 
        Delete /*+ Rule*/
        From ����ҽ������
        Where ���ͺ� = r_Rolladvice.���ͺ� And ҽ��id In (Select Column_Value From Table(t_Adviceids));
      
        --���(����ҽ��)�ϴ�ִ��ʱ��(���ϴη��͵�ĩ��ִ��ʱ��) 
        --���г���(���������Գ���)����ʱ����д��ĩ��ʱ�� 
        --��������û�У���ֻ���ܷ�����һ�Ρ� 
        v_ĩ��ʱ�� := Null;
        Begin
          --һ��ҽ���ķ�����ĩʱ����ͬ,һ����ҩ��ȡ��С�� 
          --ȡ���IDΪNULL��ҽ���ķ��ͼ�¼��ʱ�� 
          --����ҩ;������ҩ�÷�����δ��д���ͼ�¼ 
          Select /*+ Rule*/
           ĩ��ʱ��
          Into v_ĩ��ʱ��
          From ����ҽ������
          Where ҽ��id In (Select Column_Value From Table(t_Adviceids)) And
                ���ͺ� =
                (Select Max(���ͺ�) From ����ҽ������ Where ҽ��id In (Select Column_Value From Table(t_Adviceids))) And
                Rownum = 1;
        Exception
          When Others Then
            Null;
        End;
        Update ����ҽ����¼ Set �ϴ�ִ��ʱ�� = v_ĩ��ʱ�� Where ID = ҽ��id_In Or ���id = ҽ��id_In;
      
        --������������ʱ��ͬʱ�Զ�����ֹͣ 
        If Nvl(r_Rolladvice.ҽ����Ч, 0) = 1 Then
          --ɾ��(��������)�����ֹͣ״̬������¼ 
          Delete /*+ Rule*/
          From ����ҽ��״̬
          Where ҽ��id In (Select Column_Value From Table(t_Adviceids)) And
                ����ʱ�� = (Select Max(����ʱ��) From ����ҽ��״̬ Where ҽ��id = ҽ��id_In) And �������� = 8;
          --r_RollAdvice.����ʱ��:����ʱ����ܲ����Զ�ֹͣʱ����ͬ�� 
        
          --ȡɾ����Ӧ�ָ���ҽ��״̬ 
          Select ��������
          Into v_ҽ��״̬
          From ����ҽ��״̬
          Where ����ʱ�� = (Select Max(����ʱ��) From ����ҽ��״̬ Where ҽ��id = ҽ��id_In) And ҽ��id = ҽ��id_In;
        
          --�ָ�(����ҽ��)���˺��״̬ 
          Update ����ҽ����¼
          Set ҽ��״̬ = v_ҽ��״̬, ִ����ֹʱ�� = Null, ͣ��ҽ�� = Null, ͣ��ʱ�� = Null
          Where ID = ҽ��id_In Or ���id = ҽ��id_In;
        End If;
      
        --סԺ����ҽ�����ͺ�Ļ���(3-ת��;5-��Ժ;6-תԺ,11-����) 
        If r_Rolladvice.��� = 'Z' And Instr(',3,5,6,11,', ',' || r_Rolladvice.���� || ',') > 0 And
           Nvl(r_Rolladvice.Ӥ��, 0) = 0 Then
          Open c_Patilog(r_Rolladvice.����id, r_Rolladvice.��ҳid);
          Fetch c_Patilog
            Into r_Patilog;
          If c_Patilog%Found Then
            If r_Rolladvice.���� = '3' And r_Patilog.��ʼԭ�� = 3 Then
              --ȡ������ת��״̬ 
              If r_Patilog.��ʼʱ�� Is Null Then
                --ת��ҽ�������⴦����һ������������ת��ҽ��ʱ��ֻ�ܻ��������һ��,70443 
                Select Count(1)
                Into v_Count
                From ����ҽ����¼ A, ������ĿĿ¼ B
                Where a.������Ŀid = b.Id And a.����id = r_Rolladvice.����id And a.��ҳid = r_Rolladvice.��ҳid And a.������� = 'Z' And
                      b.�������� = '3' And a.ҽ��״̬ = 8 And
                      a.��ʼִ��ʱ�� > (Select ��ʼִ��ʱ�� From ����ҽ����¼ Where ID = ҽ��id_In);
                If v_Count = 0 Then
                  Zl_���˱䶯��¼_Undo(r_Rolladvice.����id, r_Rolladvice.��ҳid, v_��Ա���, v_��Ա����, Null, Null, Null, 'ת��');
                Else
                  v_Error := '����ת���Ѿ���ƣ������ٻ��ˡ�';
                  Raise Err_Custom;
                End If;
              Else
                v_Error := '����ת���Ѿ���ƣ������ٻ��ˡ�';
                Raise Err_Custom;
              End If;
            Elsif r_Rolladvice.���� In ('5', '6', '11') And r_Patilog.��ʼԭ�� = 10 Then
              --ȡ������Ԥ��Ժ״̬ 
              Zl_���˱䶯��¼_Undo(r_Rolladvice.����id, r_Rolladvice.��ҳid, v_��Ա���, v_��Ա����, Null, Null, Null, 'Ԥ��Ժ');
            End If;
          End If;
          Close c_Patilog;
        End If;
      
        --���˲���ʱ�� 
        --1.�����¼�(ֻ��һ��ҽ����¼)��������7-����,8-����,11-���� 
        If r_Rolladvice.��� = 'F' Or r_Rolladvice.��� = 'Z' And Instr(',7,8,11,', ',' || r_Rolladvice.���� || ',') > 0 Then
          Zl_���Ӳ���ʱ��_Delete(r_Rolladvice.����id, r_Rolladvice.��ҳid, 'ҽ��', r_Rolladvice.��������id, ҽ��id_In);
        End If;
      
        --2.���⴦��֪��ͬ����(������ص�֪��ͬ�����ٴε��ã���Ϊ����������������Ŀ�����й�����֪��ͬ����) 
        If Instr('C,D,E,F,G,K,L', r_Rolladvice.���) > 0 Then
          For R In (Select a.Id, a.������� From ����ҽ����¼ A Where a.Id = ҽ��id_In Or a.���id = ҽ��id_In) Loop
            --���id��һ��ҽ����һ����������ģ�����Ҫ���ж�һ����� 
            If Instr('C,D,E,F,G,K,L', r.�������) > 0 Then
              Zl_���Ӳ���ʱ��_Delete(r_Rolladvice.����id, r_Rolladvice.��ҳid, 'ҽ��', r_Rolladvice.��������id, r.Id);
            End If;
          End Loop;
        End If;
      
        If r_Rolladvice.��� = 'Z' And r_Rolladvice.�������� = '6' Then
          --���� 
          b_Message.Zlhis_Cis_040(r_Rolladvice.����id, r_Rolladvice.��ҳid, r_Rolladvice.���ͺ�, r_Rolladvice.��id,
                                  r_Rolladvice.No);
        Elsif r_Rolladvice.��� = 'Z' And r_Rolladvice.�������� = '8' Then
          --���� 
          b_Message.Zlhis_Cis_041(r_Rolladvice.����id, r_Rolladvice.��ҳid, r_Rolladvice.���ͺ�, r_Rolladvice.��id,
                                  r_Rolladvice.No);
        Elsif r_Rolladvice.��� = 'Z' And r_Rolladvice.�������� = '11' Then
          --���� 
          b_Message.Zlhis_Cis_042(r_Rolladvice.����id, r_Rolladvice.��ҳid, r_Rolladvice.���ͺ�, r_Rolladvice.��id,
                                  r_Rolladvice.No);
        Elsif r_Rolladvice.��� = 'E' And r_Rolladvice.�������� = '5' Then
          --�������� 
          b_Message.Zlhis_Cis_043(r_Rolladvice.����id, r_Rolladvice.��ҳid, r_Rolladvice.���ͺ�, r_Rolladvice.��id,
                                  r_Rolladvice.No);
        Elsif r_Rolladvice.��� = 'H' And Nvl(r_Rolladvice.��������, '0') = '0' Then
          --������ 
          b_Message.Zlhis_Cis_007(r_Rolladvice.����id, r_Rolladvice.��ҳid, r_Rolladvice.���ͺ�, r_Rolladvice.��id,
                                  r_Rolladvice.No);
        End If;
      End If;
    End If;
    Exit When r_Rolladvice.���ͺ� = 0;
  End Loop;
  Close c_Rolladvice;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ����¼_����;
/
--138877:��͢��,2019-04-23,�������۲���תסԺ������Ϣ
Create Or Replace Package b_Message Is
  Procedure p_Msg_Todo_Insert
  (
    Msg_Code_In  Zlmsg_Todo.Msg_Code%Type,
    Key_Value_In Zlmsg_Todo.Key_Value%Type
  );
  --����ƽ̨��������
  Procedure Set_Platform_Call(Platform_Call Number);
  --��������
  Procedure Zlhis_Dict_001(Id_In ���ű�.Id%Type);
  --�޸Ĳ���
  Procedure Zlhis_Dict_002(����id_In ���ű�.Id%Type);
  --ͣ�ò���
  Procedure Zlhis_Dict_003(����id_In ���ű�.Id%Type);
  --���ò���
  Procedure Zlhis_Dict_004(����id_In ���ű�.Id%Type);
  --������Ա
  Procedure Zlhis_Dict_005(��Աid_In ��Ա��.Id%Type);
  --�޸���Ա
  Procedure Zlhis_Dict_006(��Աid_In ��Ա��.Id%Type);
  --ͣ����Ա
  Procedure Zlhis_Dict_007(��Աid_In ��Ա��.Id%Type);
  --������Ա
  Procedure Zlhis_Dict_008(��Աid_In ��Ա��.Id%Type);
  --�����շ���Ŀ
  Procedure Zlhis_Dict_009(ϸĿid_In �շ���ĿĿ¼.Id%Type);
  --�޸��շ���Ŀ
  Procedure Zlhis_Dict_010(ϸĿid_In �շ���ĿĿ¼.Id%Type);
  --ͣ���շ���Ŀ
  Procedure Zlhis_Dict_011(ϸĿid_In �շ���ĿĿ¼.Id%Type);
  --�����շ���Ŀ
  Procedure Zlhis_Dict_012(ϸĿid_In �շ���ĿĿ¼.Id%Type);
  --����������Ŀ
  Procedure Zlhis_Dict_013(����id_In ������ĿĿ¼.Id%Type);
  --�޸�������Ŀ
  Procedure Zlhis_Dict_014(����id_In ������ĿĿ¼.Id%Type);
  --ͣ��������Ŀ
  Procedure Zlhis_Dict_015(����id_In ������ĿĿ¼.Id%Type);
  --����������Ŀ
  Procedure Zlhis_Dict_016(����id_In ������ĿĿ¼.Id%Type);
  --����������Ŀ
  Procedure Zlhis_Dict_017(����id_In ������ĿĿ¼.Id%Type);
  --�޸ļ�����Ŀ
  Procedure Zlhis_Dict_018(����id_In ������ĿĿ¼.Id%Type);
  --ɾ��������Ŀ
  Procedure Zlhis_Dict_019
  (
    ����id_In ������ĿĿ¼.Id%Type,
    ����_In   ����������Ŀ.����%Type,
    ������_In ����������Ŀ.������%Type,
    Ӣ����_In ����������Ŀ.Ӣ����%Type
  );

  --������������Ŀ¼
  Procedure Zlhis_Dict_021(����id_In ��������Ŀ¼.Id%Type);
  --�޸ļ�������Ŀ¼
  Procedure Zlhis_Dict_022(����id_In ��������Ŀ¼.Id%Type);
  --ͣ�ü�������Ŀ¼
  Procedure Zlhis_Dict_023(����id_In ��������Ŀ¼.Id%Type);
  --���ü�������Ŀ¼
  Procedure Zlhis_Dict_024(����id_In ��������Ŀ¼.Id%Type);
  --����ҩƷ����
  Procedure Zlhis_Dict_025
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  );
  --�޸�ҩƷ����
  Procedure Zlhis_Dict_026
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  );
  --ɾ��ҩƷ����
  Procedure Zlhis_Dict_027
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type,
    ����_In ���Ʒ���Ŀ¼.����%Type,
    ����_In ���Ʒ���Ŀ¼.����%Type
  );
  --ͣ��ҩƷ����
  Procedure Zlhis_Dict_028
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  );
  --����ҩƷ����
  Procedure Zlhis_Dict_029
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  );
  --����ҩƷƷ��
  Procedure Zlhis_Dict_030
  (
    ���_In ������ĿĿ¼.���%Type,
    Id_In   ������ĿĿ¼.Id%Type
  );
  --�޸�ҩƷƷ��
  Procedure Zlhis_Dict_031
  (
    ���_In ������ĿĿ¼.���%Type,
    Id_In   ������ĿĿ¼.Id%Type
  );
  --ɾ��ҩƷƷ��
  Procedure Zlhis_Dict_032
  (
    ���_In ������ĿĿ¼.���%Type,
    Id_In   ������ĿĿ¼.Id%Type,
    ����_In ������ĿĿ¼.����%Type,
    ����_In ������ĿĿ¼.����%Type
  );
  --ͣ��ҩƷƷ��
  Procedure Zlhis_Dict_033
  (
    ���_In ������ĿĿ¼.���%Type,
    Id_In   ������ĿĿ¼.Id%Type
  );
  --����ҩƷƷ��
  Procedure Zlhis_Dict_034
  (
    ���_In ������ĿĿ¼.���%Type,
    Id_In   ������ĿĿ¼.Id%Type
  );
  --����ҩƷ���
  Procedure Zlhis_Dict_035
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  );
  --�޸�ҩƷ���
  Procedure Zlhis_Dict_036
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  );
  --ɾ��ҩƷ���
  Procedure Zlhis_Dict_037
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type,
    ����_In   �շ���ĿĿ¼.����%Type,
    ����_In   �շ���ĿĿ¼.����%Type,
    ���_In   �շ���ĿĿ¼.���%Type,
    ����_In   �շ���ĿĿ¼.����%Type
  );
  --ͣ��ҩƷ���
  Procedure Zlhis_Dict_038
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  );
  --����ҩƷ���
  Procedure Zlhis_Dict_039
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  );
  --����ҩƷ�洢�ⷿ
  Procedure Zlhis_Dict_040
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  );
  --����ҩƷ�����޶�
  Procedure Zlhis_Dict_041
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  );
  --��������Ʒ��
  Procedure Zlhis_Dict_042(Id_In ������ĿĿ¼.Id%Type);
  --�������Ĺ��
  Procedure Zlhis_Dict_043(Id_In �շ���ĿĿ¼.Id%Type);
  --�޸����Ĺ��
  Procedure Zlhis_Dict_044(Id_In �շ���ĿĿ¼.Id%Type);
  --ɾ�����Ĺ��
  Procedure Zlhis_Dict_045
  (
    Id_In   �շ���ĿĿ¼.Id%Type,
    ����_In �շ���ĿĿ¼.����%Type,
    ����_In �շ���ĿĿ¼.����%Type,
    ���_In �շ���ĿĿ¼.���%Type,
    ����_In �շ���ĿĿ¼.����%Type
  );
  --ͣ�����Ĺ��
  Procedure Zlhis_Dict_046(Id_In �շ���ĿĿ¼.Id%Type);
  --�������Ĺ��
  Procedure Zlhis_Dict_047(Id_In �շ���ĿĿ¼.Id%Type);
  --ҽ������
  Procedure Zlhis_Dict_048
  (
    ����_In       In ����֧����Ŀ.����%Type,
    �շ�ϸĿid_In In ����֧����Ŀ.�շ�ϸĿid%Type
  );
  --ɾ��ҽ������
  Procedure Zlhis_Dict_049
  (
    ����_In       In ����֧����Ŀ.����%Type,
    �շ�ϸĿid_In In ����֧����Ŀ.�շ�ϸĿid%Type,
    ��Ŀ����_In   In �շ���ĿĿ¼.����%Type,
    ��Ŀ����_In   In �շ���ĿĿ¼.����%Type,
    ҽ������_In   In ����֧����Ŀ.��Ŀ����%Type,
    ҽ������_In   In ����֧����Ŀ.��Ŀ����%Type
  );
  --�������ķ���
  Procedure Zlhis_Dict_050
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  );
  --�޸����ķ���
  Procedure Zlhis_Dict_051
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  );
  --ɾ�����ķ���
  Procedure Zlhis_Dict_052
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type,
    ����_In ���Ʒ���Ŀ¼.����%Type,
    ����_In ���Ʒ���Ŀ¼.����%Type
  );
  --�շѼ�Ŀ�䶯
  Procedure Zlhis_Dict_053
  (
    �շ���ĿId_In       �շ���ĿĿ¼.Id%Type
  );
  --�����շѶ��ձ䶯
  Procedure Zlhis_Dict_054
  (
    ������ĿId_In     ���Ʒ���Ŀ¼.Id%Type
  );
  --�������Ƽ������
  Procedure Zlhis_Dictpacs_001
  (
    ����_In   ���Ƽ������.����%Type,
    ����_In   ���Ƽ������.����%Type,
    ����_In   ���Ƽ������.����%Type,
    ������_In ���Ƽ������.������%Type
  );
  --�޸����Ƽ������
  Procedure Zlhis_Dictpacs_002
  (
    ����_In   ���Ƽ������.����%Type,
    ����_In   ���Ƽ������.����%Type,
    ����_In   ���Ƽ������.����%Type,
    ������_In ���Ƽ������.������%Type
  );
  --ɾ�����Ƽ������
  Procedure Zlhis_Dictpacs_003
  (
    ����_In   ���Ƽ������.����%Type,
    ����_In   ���Ƽ������.����%Type,
    ����_In   ���Ƽ������.����%Type,
    ������_In ���Ƽ������.������%Type
  );
  --�������Ƽ�鲿λ
  Procedure Zlhis_Dictpacs_004
  (
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ��ע_In     ���Ƽ�鲿λ.��ע%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    �����Ա�_In ���Ƽ�鲿λ.�����Ա�%Type
  );
  --�޸����Ƽ�鲿λ
  Procedure Zlhis_Dictpacs_005
  (
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ��ע_In     ���Ƽ�鲿λ.��ע%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    �����Ա�_In ���Ƽ�鲿λ.�����Ա�%Type
  );
  --ɾ�����Ƽ�鲿λ
  Procedure Zlhis_Dictpacs_006
  (
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ��ע_In     ���Ƽ�鲿λ.��ע%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    �����Ա�_In ���Ƽ�鲿λ.�����Ա�%Type
  );
  --����������Ŀ��λ
  Procedure Zlhis_Dictpacs_007
  (
    Id_In     ������Ŀ��λ.Id%Type,
    ��Ŀid_In ������Ŀ��λ.��Ŀid%Type,
    ����_In   ������Ŀ��λ.����%Type,
    ��λ_In   ������Ŀ��λ.��λ%Type,
    ����_In   ������Ŀ��λ.����%Type,
    Ĭ��_In   ������Ŀ��λ.Ĭ��%Type
  );
  --�޸�������Ŀ��λ
  Procedure Zlhis_Dictpacs_008
  (
    Id_In     ������Ŀ��λ.Id%Type,
    ��Ŀid_In ������Ŀ��λ.��Ŀid%Type,
    ����_In   ������Ŀ��λ.����%Type,
    ��λ_In   ������Ŀ��λ.��λ%Type,
    ����_In   ������Ŀ��λ.����%Type,
    Ĭ��_In   ������Ŀ��λ.Ĭ��%Type
  );
  --ɾ��������Ŀ��λ
  Procedure Zlhis_Dictpacs_009
  (
    Id_In     ������Ŀ��λ.Id%Type,
    ��Ŀid_In ������Ŀ��λ.��Ŀid%Type,
    ����_In   ������Ŀ��λ.����%Type,
    ��λ_In   ������Ŀ��λ.��λ%Type,
    ����_In   ������Ŀ��λ.����%Type,
    Ĭ��_In   ������Ŀ��λ.Ĭ��%Type
  );
  --�������Ƽ���걾
  Procedure Zlhis_Dictlis_004
  (
    ����_In     ���Ƽ���걾.����%Type,
    ����_In     ���Ƽ���걾.����%Type,
    ����_In     ���Ƽ���걾.����%Type,
    �����Ա�_In ���Ƽ���걾.�����Ա�%Type
  );
  --�޸����Ƽ���걾
  Procedure Zlhis_Dictlis_005
  (
    ����_In     ���Ƽ���걾.����%Type,
    ����_In     ���Ƽ���걾.����%Type,
    ����_In     ���Ƽ���걾.����%Type,
    �����Ա�_In ���Ƽ���걾.�����Ա�%Type
  );
  --ɾ��������Ŀ��λ
  Procedure Zlhis_Dictlis_006
  (
    ����_In     ���Ƽ���걾.����%Type,
    ����_In     ���Ƽ���걾.����%Type,
    ����_In     ���Ƽ���걾.����%Type,
    �����Ա�_In ���Ƽ���걾.�����Ա�%Type
  );
  --������Ѫ������
  Procedure Zlhis_Dictlis_007
  (
    ����_In   ��Ѫ������.����%Type,
    ����_In   ��Ѫ������.����%Type,
    ����_In   ��Ѫ������.����%Type,
    ��Ӽ�_In ��Ѫ������.��Ӽ�%Type,
    ��Ѫ��_In ��Ѫ������.��Ѫ��%Type,
    ���_In   ��Ѫ������.���%Type,
    ��ɫ_In   ��Ѫ������.��ɫ%Type,
    ����id_In ��Ѫ������.����id%Type
  );
  --�޸Ĳ�Ѫ������
  Procedure Zlhis_Dictlis_008
  (
    ����_In   ��Ѫ������.����%Type,
    ����_In   ��Ѫ������.����%Type,
    ����_In   ��Ѫ������.����%Type,
    ��Ӽ�_In ��Ѫ������.��Ӽ�%Type,
    ��Ѫ��_In ��Ѫ������.��Ѫ��%Type,
    ���_In   ��Ѫ������.���%Type,
    ��ɫ_In   ��Ѫ������.��ɫ%Type,
    ����id_In ��Ѫ������.����id%Type
  );
  --ɾ����Ѫ������
  Procedure Zlhis_Dictlis_009
  (
    ����_In   ��Ѫ������.����%Type,
    ����_In   ��Ѫ������.����%Type,
    ����_In   ��Ѫ������.����%Type,
    ��Ӽ�_In ��Ѫ������.��Ӽ�%Type,
    ��Ѫ��_In ��Ѫ������.��Ѫ��%Type,
    ���_In   ��Ѫ������.���%Type,
    ��ɫ_In   ��Ѫ������.��ɫ%Type,
    ����id_In ��Ѫ������.����id%Type
  );

  --ҩƷ��ҩ����
  Procedure Zlhis_Drug_001(No_In ҩƷ�շ���¼.No%Type);
  --ȡ��ҩƷ��ҩ����
  Procedure Zlhis_Drug_002(No_In ҩƷ�շ���¼.No%Type);
  --ҩƷ�ƿⵥ����
  Procedure Zlhis_Drug_003(No_In ҩƷ�շ���¼.No%Type);
  --ҩƷ�ƿⵥ����
  Procedure Zlhis_Drug_004(No_In ҩƷ�շ���¼.No%Type);

  --���ŷ�ҩ
  Procedure Zlhis_Drug_005
  (
    �ⷿid_In ҩƷ�շ���¼.�ⷿid%Type,
    �շ�id_In ҩƷ�շ���¼.Id%Type
  );
  --������ҩ
  Procedure Zlhis_Drug_006
  (
    �����շ�id_In ҩƷ�շ���¼.Id%Type,
    �����շ�id_In ҩƷ�շ���¼.Id%Type,
    ����_In       ҩƷ�շ���¼.ʵ������%Type,
    ����id_In     ������ü�¼.Id%Type
  );
  --ҩƷ����
  Procedure Zlhis_Drug_007(�۸�id_In ҩƷ�۸��¼.Id%Type);
  --���䷢��
  Procedure Zlhis_Drug_008(��¼ids_In Varchar2);
  --ҩƷ���ۼ�
  Procedure Zlhis_Drug_009
  (
    �۸�id_In ҩƷ�۸��¼.Id%Type,
    ʱ��_In   Number
  );
  --���ĵ��ɱ���
  Procedure Zlhis_Drug_010(�۸�id_In �ɱ��۵�����Ϣ.Id%Type);
  --���ĵ��ۼ�
  Procedure Zlhis_Drug_011
  (
    �۸�id_In �շѼ�Ŀ.Id%Type,
    ʱ��_In   Number
  );
  --2.ֹͣ����ҽ����סԺ
  Procedure Zlhis_Cis_002
  (
    ����id_In  In ����ҽ����¼.����id%Type,
    ��ҳid_In  In ����ҽ����¼.��ҳid%Type,
    ҽ��id_In  In ����ҽ����¼.Id%Type,
    ҽ��ids_In In Varchar2
  );
  --3.���ϻ���ҽ��������/סԺ
  Procedure Zlhis_Cis_003
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );

  --4.��������ҽ����סԺ
  Procedure Zlhis_Cis_004
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );

  --5.������������ҽ����סԺ
  Procedure Zlhis_Cis_005
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );

  --6.���߻�����ҽ����סԺ
  Procedure Zlhis_Cis_006
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );

  --7.�������߻�����ҽ����סԺ
  Procedure Zlhis_Cis_007
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  );

  --���ﻼ�߽���
  Procedure Zlhis_Cis_008
  (
    ����id_In In ����ҽ����¼.����id%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type
  );

  --���ﻼ��ȡ������
  Procedure Zlhis_Cis_009
  (
    ����id_In In ����ҽ����¼.����id%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type
  );

  --10.�´ﻼ����ϣ�����/סԺ
  Procedure Zlhis_Cis_010
  (
    ����id_In In ���˹Һż�¼.����id%Type,
    ����id_In In ���˹Һż�¼.Id%Type, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    ���id_In In ������ϼ�¼.Id%Type
  );
  --11.�����������
  Procedure Zlhis_Cis_011
  (
    ����id_In   In ���˹Һż�¼.����id%Type,
    ����id_In   In ���˹Һż�¼.Id%Type, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    Id_In       In ������ϼ�¼.Id%Type,
    ����id_In   In ������ϼ�¼.����id%Type,
    ���id_In   In ������ϼ�¼.���id%Type,
    �������_In In ������ϼ�¼.�������%Type
  );

  --����ִ��ҽ��У��
  Procedure Zlhis_Cis_012
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );

  --13.����Σ��ֵ�Ķ�֪ͨ
  Procedure Zlhis_Cis_014
  (
    ����id_In In ���˹Һż�¼.����id%Type,
    ����id_In In ���˹Һż�¼.Id%Type, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    ҽ��id_In In ����ҽ����¼.Id%Type,
    ��Ϣid_In In ҵ����Ϣ�嵥.Id%Type
  );

  --15.���߼������룬����/סԺ
  Procedure Zlhis_Cis_016
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    ������Դ_In In ����ҽ����¼.������Դ%Type
  );
  --16.���߼�����룬����/סԺ
  Procedure Zlhis_Cis_017
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    ������Դ_In In ����ҽ����¼.������Դ%Type
  );
  --17.�����������룬����/סԺ
  Procedure Zlhis_Cis_018
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );
  --18.������Ѫ���룬סԺ
  Procedure Zlhis_Cis_019
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );
  --19.���߻������룬סԺ
  Procedure Zlhis_Cis_020
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );
  --20.��������ҽ����סԺ
  Procedure Zlhis_Cis_021
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );
  --21.��������ҽ����סԺ
  Procedure Zlhis_Cis_022
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );
  --22.������������ҽ����סԺ
  Procedure Zlhis_Cis_023
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );

  --24.���Σ��ֵ�Ķ�֪ͨ
  Procedure Zlhis_Cis_025
  (
    ����id_In In ���˹Һż�¼.����id%Type,
    ����id_In In ���˹Һż�¼.Id%Type, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    ҽ��id_In In ����ҽ����¼.Id%Type,
    ��Ϣid_In In ҵ����Ϣ�嵥.Id%Type
  );

  --����ִ��ҽ������
  Procedure Zlhis_Cis_026
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );

  --�������߼�������
  Procedure Zlhis_Cis_036
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    No_In       In ����ҽ������.No%Type,
    ������Դ_In In ����ҽ����¼.������Դ%Type
  );

  --�������߼������
  Procedure Zlhis_Cis_037
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    No_In       In ����ҽ������.No%Type,
    ������Դ_In In ����ҽ����¼.������Դ%Type
  );

  --����������������
  Procedure Zlhis_Cis_038
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  );

  --����������Ѫ����
  Procedure Zlhis_Cis_039
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  );

  --�������߻�������
  Procedure Zlhis_Cis_040
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  );

  --������������ҽ��
  Procedure Zlhis_Cis_041
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  );

  --������������ҽ��
  Procedure Zlhis_Cis_042
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  );

  --������������ҽ��
  Procedure Zlhis_Cis_043
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  );

  --��������ִ��ҽ��
  Procedure Zlhis_Cis_044
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    No_In       In ����ҽ������.No%Type,
    ��������_In In ����ҽ������.��������%Type,
    �״�ʱ��_In In ����ҽ������.�״�ʱ��%Type,
    ĩ��ʱ��_In In ����ҽ������.ĩ��ʱ��%Type,
    ��������_In In ����ҽ������.��������%Type
  );
  --����ҽ��ִ�еǼ�
  Procedure Zlhis_Cis_050
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    Ҫ��ʱ��_In In ����ҽ��ִ��.Ҫ��ʱ��%Type,
    ִ��ʱ��_In In ����ҽ��ִ��.ִ��ʱ��%Type
  );

  --����ҽ��ȡ��ִ�еǼ�
  Procedure Zlhis_Cis_051
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    Ҫ��ʱ��_In In ����ҽ��ִ��.Ҫ��ʱ��%Type,
    ִ��ʱ��_In In ����ҽ��ִ��.ִ��ʱ��%Type,
    ��������_In In ����ҽ��ִ��.��������%Type,
    ִ�н��_In In ����ҽ��ִ��.ִ�н��%Type,
    ִ��ժҪ_In In ����ҽ��ִ��.ִ��ժҪ%Type,
    ִ�п���_In In ����ҽ��ִ��.ִ�п���id%Type,
    ִ����_In   In ����ҽ��ִ��.ִ����%Type,
    �˶���_In   In ����ҽ��ִ��.�˶���%Type,
    ��¼��Դ_In In ����ҽ��ִ��.��¼��Դ%Type
  );
  --����ҽ��ִ�����
  Procedure Zlhis_Cis_052
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );
  --����ҽ������ִ�����
  Procedure Zlhis_Cis_053
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );
  --�������뷢�ͺ��޸�
  Procedure Zlhis_Cis_056
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    ������Դ_In In ����ҽ����¼.������Դ%Type
  );

  --���ﻼ����ɾ���
  Procedure Zlhis_Cis_057
  (
    ����id_In In ����ҽ����¼.����id%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type
  );

  --���ﻼ��ȡ����ɾ���
  Procedure Zlhis_Cis_058
  (
    ����id_In In ����ҽ����¼.����id%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type
  );

  --ȷ��ֹͣ����ҽ�� 
  Procedure Zlhis_Cis_059
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );

  --26.��鱨����ɣ�������ʱ
  Procedure Zlhis_Pacs_001
  (
    ҽ��id_In   In Ӱ�����¼.ҽ��id%Type,
    ����id_Ins  In Varchar2,
    ��������_In In Number --1-�ϰ�PACS���棬2-�ϰ没���༭�����棬3-�°�༭������
  );
  --27.���״̬ͬ�������״̬�ı��
  Procedure Zlhis_Pacs_002
  (
    ҽ��id_In In Ӱ�����¼.ҽ��id%Type,
    ԭ״̬_In In ����ҽ������.ִ�й���%Type,
    ��״̬_In In ����ҽ������.ִ�й���%Type
  );
  --28.���״̬���ˣ����״̬���˺�
  Procedure Zlhis_Pacs_003
  (
    ҽ��id_In In Ӱ�����¼.ҽ��id%Type,
    ԭ״̬_In In ����ҽ������.ִ�й���%Type,
    ��״̬_In In ����ҽ������.ִ�й���%Type
  );
  --29.��鱨�泷��������������ʱ
  Procedure Zlhis_Pacs_004
  (
    ҽ��id_In   In Ӱ�����¼.ҽ��id%Type,
    ����id_Ins  In Varchar2,
    ��������_In In Number --1-�ϰ�PACS���棬2-�ϰ没���༭�����棬3-�°�༭������
  );
  --30.���Σ��ֵ֪ͨ����鷢��Σ��ֵʱ
  Procedure Zlhis_Pacs_005(ҽ��id_In In Ӱ�����¼.ҽ��id%Type);
  -- ���ԤԼ֪ͨ�����ԤԼʱ
  Procedure Zlhis_Pacs_006
  (
    ҽ��id_In In Ӱ�����¼.ҽ��id%Type,
    ԤԼid_In In Ris���ԤԼ.ԤԼid%Type
  );
  -- ȡ�����ԤԼ��ȡ��ԤԼʱ
  Procedure Zlhis_Pacs_007
  (
    ҽ��id_In       In Ӱ�����¼.ҽ��id%Type,
    ԤԼid_In       In Ris���ԤԼ.ԤԼid%Type,
    ԤԼ����_In     In Ris���ԤԼ.ԤԼ����%Type,
    ԤԼ���_In     In Ris���ԤԼ.���%Type,
    ����豸����_In In Ris���ԤԼ.����豸����%Type
  );

  --36.���߷���
  Procedure Zlhis_Patient_018
  (
    �䶯id_In   In ���˱䶯��¼.Id%Type,
    ����id_In   In ������Ϣ.����id%Type,
    �����id_In In ҽ�ƿ����.Id%Type,
    ����_In     In ����ҽ�ƿ���Ϣ.����%Type
  );

  --37.�����˿�
  Procedure Zlhis_Patient_019
  (
    �䶯id_In   In ���˱䶯��¼.Id%Type,
    ����id_In   In ������Ϣ.����id%Type,
    �����id_In In ҽ�ƿ����.Id%Type,
    ����_In     In ����ҽ�ƿ���Ϣ.����%Type
  );

  --38.�����˿�
  Procedure Zlhis_Patient_020
  (
    �䶯id_In   In ���˱䶯��¼.Id%Type,
    ����id_In   In ������Ϣ.����id%Type,
    �����id_In In ҽ�ƿ����.Id%Type,
    ԭ����_In   In ����ҽ�ƿ���Ϣ.����%Type,
    �¿���_In   In ����ҽ�ƿ���Ϣ.����%Type
  );

  --39.���˹ҺŵǼǣ�����ԤԼ�Ǽ�)
  Procedure Zlhis_Regist_001
  (
    �Һ�id_In In ���˹Һż�¼.Id%Type,
    No_In     In ���˹Һż�¼.No%Type
  );

  --40.���˷���
  Procedure Zlhis_Regist_002
  (
    �Һ�id_In In ���˹Һż�¼.Id%Type,
    No_In     In ���˹Һż�¼.No%Type,
    ����_In   In ���˹Һż�¼.����%Type
  );

  --41.�����˺�
  Procedure Zlhis_Regist_003
  (
    �Һ�id_In In ���˹Һż�¼.Id%Type,
    No_In     In ���˹Һż�¼.No%Type
  );

  --42.�ٴ����ﰲ�ŵ���
  Procedure Zlhis_Regist_004
  (
    �䶯ԭ��_In In Integer, --1-ͣ��;2-����;3-���ұ䶯
    ��¼id_In   In �ٴ������¼.Id%Type,
    �䶯id_In   In �ٴ�����䶯��¼.Id%Type
  );

  --43.���ﻼ�߹ҺŻ��Ų���
  Procedure Zlhis_Regist_005
  (
    No_In         In ���˹Һż�¼.No%Type,
    �䶯ԭ��_In   Integer, --1-����;2-����;3-ԤԼ���ڱ䶯,
    ����䶯id_In ����䶯��¼.Id%Type
  );

  --���������շѼ��������
  --��������_In:1-�շѽ��㣬2-�������
  Procedure Zlhis_Charge_002
  (
    ��������_In In Number,
    ����id_In   In ������ü�¼.����id%Type
  );

  --46.�����˷ѵ���
  Procedure Zlhis_Charge_004
  (
    �˷�����_In In Number,
    ����id_In   In ������ü�¼.����id%Type
  );

  --47.��Ԥ����
  Procedure Zlhis_Charge_005
  (
    Ԥ��id_In In ����Ԥ����¼.Id%Type,
    ���ݺ�_In In ����Ԥ����¼.No%Type
  );

  --48.��Ԥ����(����������Ԥ�����)
  Procedure Zlhis_Charge_006
  (
    ��Ԥ��id_In In ����Ԥ����¼.Id%Type,
    ���ݺ�_In   In ����Ԥ����¼.No%Type
  );

  --סԺ���ʵ���
  Procedure Zlhis_Charge_007
  (
    �շ����_In In סԺ���ü�¼.�շ����%Type,
    ����id_In   In סԺ���ü�¼.Id%Type
  );

  --סԺ���ʵ�������
  Procedure Zlhis_Charge_008
  (
    �շ����_In In סԺ���ü�¼.�շ����%Type,
    ����id_In   In סԺ���ü�¼.Id%Type,
    �շ�ids_In  In Varchar2 := Null --���ܷ���ID��Ӧ����շ�id����Ӧ��ʽ���շ�id,����|�շ�id,��������ҩƷ����
  );

  --53.סԺ������Ժ�Ǽ�
  Procedure Zlhis_Patient_001
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );
  --54.סԺ������Ժ���
  Procedure Zlhis_Patient_002
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );
  --56.סԺ���ߴ�λ���
  Procedure Zlhis_Patient_004
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );
  --57.סԺ���߲�����
  Procedure Zlhis_Patient_005
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );
  --58.סԺ���߱������
  Procedure Zlhis_Patient_006
  (
    ����id_In   In ������ҳ.����id%Type,
    ��ҳid_In   In ������ҳ.��ҳid%Type,
    ������ʽ_In In Varchar2
  );
  --59.סԺ����ҽ�����
  Procedure Zlhis_Patient_007
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );
  --סԺ���߻���ȼ����
  Procedure Zlhis_Patient_008
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );
  --60.סԺ����Ԥ��Ժ
  Procedure Zlhis_Patient_009
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );
  --61.סԺ���߳�Ժ
  Procedure Zlhis_Patient_010
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );
  --62.סԺ�����������Ǽ�
  Procedure Zlhis_Patient_011
  (
    ����id_In   In ������ҳ.����id%Type,
    ��ҳid_In   In ������ҳ.��ҳid%Type,
    Ӥ�����_In ����ҽ����¼.Ӥ��%Type
  );
  --63.סԺ����ת�����
  Procedure Zlhis_Patient_012
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );
  --64.�������Ǽ�����
  Procedure Zlhis_Patient_013
  (
    ����id_In   In ������ҳ.����id%Type,
    ��ҳid_In   In ������ҳ.��ҳid%Type,
    Ӥ�����_In ����ҽ����¼.Ӥ��%Type
  );
  --65.���ﻼ�ߵǼ�
  Procedure Zlhis_Patient_015(����id_In In ������ҳ.����id%Type);
  --66.������Ϣ�޸�
  Procedure Zlhis_Patient_016(����id_In In ������ҳ.����id%Type);

  --67.���ߺϲ�
  Procedure Zlhis_Patient_017 
  ( 
    ����id_In   In ������ҳ.����id%Type, 
    ԭ����id_In In ������ҳ.����id%Type,
    �仯ids_In  In Varchar2 
  ); 

  --69.����ת����ת��
  Procedure Zlhis_Patient_026
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );

  Procedure Zlhis_Patient_028(����id_In In ������ҳ.����id%Type);


  --79.���۲���תסԺ����
  Procedure Zlhis_Patient_029
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );


  --Ѫ��:������Ѫ���
  Procedure Zlhis_Blood_001(ҽ��id_In In ����ҽ����¼.Id%Type);
  --Ѫ��:������Ѫ�ܾ�
  Procedure Zlhis_Blood_002(ҽ��id_In In ����ҽ����¼.Id%Type);

  --70.����걾���
  Procedure Zlhis_Lis_001(�걾id_In In ����걾��¼.Id%Type);
  --71.����걾��˳���
  Procedure Zlhis_Lis_002(�걾id_In In ����걾��¼.Id%Type);
  --73.����걾�����ӡ
  Procedure Zlhis_Lis_004
  (
    ��������_In In ����ҽ������.��������%Type,
    ҽ��id_In   In ����ҽ������.ҽ��id%Type,
    ҽ��ids_In  In Varchar2
  );
  --74.����걾�����ӡ����
  Procedure Zlhis_Lis_005
  (
    ��������_In In ����ҽ������.��������%Type,
    ҽ��id_In   In ����ҽ������.ҽ��id%Type,
    ҽ��ids_In  In Varchar2
  );
  --75.����걾����
  Procedure Zlhis_Lis_006(�걾id_In In ����걾��¼.Id%Type);
  --76.����걾���ճ���
  Procedure Zlhis_Lis_007(�걾id_In In ����걾��¼.Id%Type);
  --77.����걾����
  Procedure Zlhis_Lis_008(�걾id_In In ����걾��¼.Id%Type);
End b_Message;
/
Create Or Replace Package Body b_Message Is
  --�Ƿ���ƽ̨����
  Is_Platform_Call Number(1) := 0;
  --��Ϣ��������
  Message_Creator Zlmsg_Todo.Creator%Type := Null;
  --������Ϣ��ѯ���
  Type Tmap_Msg_Using Is Table Of Number(1) Index By Varchar2(30);
  Zlmsg_Map Tmap_Msg_Using;
  --��Ϣ�Ƿ�����
  Function p_Msg_Using(Msg_Code_In Zlmsg_Lists.Code%Type) Return Number As
    n_Using Zlmsg_Lists.Using%Type;
    v_Code  Zlmsg_Lists.Code%Type;
  Begin
    If Is_Platform_Call = 1 Then
      Return 0;
    End If;
    v_Code := Upper(Msg_Code_In);
    Begin
      n_Using := Zlmsg_Map(v_Code);
      Return n_Using;
    Exception
      When No_Data_Found Then
        --����ȡMax�ݴ��������൱�����,�û�����û�в�ȡͬ���޸Ļ��Լ���������Ϣ���͵���δע�ᵽZlmsg_Lists���������������ִ���



      
        Select Nvl(Using, 0) Into n_Using From Zlmsg_Lists A Where Code = v_Code;
        Zlmsg_Map(v_Code) := n_Using;
        --��ѯ������Ϣ����Ա�������������ִ�д���
        If Message_Creator Is Null Then
          Message_Creator := Zl_Username;
        End If;
        Return n_Using;
    End;
  Exception
    When Others Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || 'δ��Zlmsg_Lists���ҵ���Ϣ"' || v_Code || '"������ϵ����Ա���д���' || '[ZLSOFT]');
      Return 0;
  End;
  Procedure p_Msg_Todo_Insert
  (
    Msg_Code_In  Zlmsg_Todo.Msg_Code%Type,
    Key_Value_In Zlmsg_Todo.Key_Value%Type
  ) Is
  Begin
    If p_Msg_Using(Msg_Code_In) = 0 Then
      Return;
    End If;
    Insert Into Zlmsg_Todo
      (ID, Msg_Code, Key_Value, State, Create_Time, Creator)
    Values
      (Zlmsg_Todo_Id.Nextval, Upper(Msg_Code_In), Key_Value_In, 0, Sysdate, Message_Creator);
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Msg_Todo_Insert;
  --���õ�ǰ�ỰΪƽ̨����
  Procedure Set_Platform_Call(Platform_Call Number) Is
  Begin
    Is_Platform_Call := Platform_Call;
  End Set_Platform_Call;
  --��ϢZlhis_Dict_001
  Procedure Zlhis_Dict_001(Id_In ���ű�.Id%Type) Is
    v_Define Xmltype;
    v_Value  Zlmsg_Todo.Key_Value%Type;
  Begin
    If p_Msg_Using('ZLHIS_DICT_001') = 0 Then
      Return;
    End If;
    Begin
      Select Xmltype(Key_Define) Into v_Define From Zlmsg_Lists Where Code = 'ZLHIS_DICT_001';
    Exception
      When Others Then
        v_Define := Xmltype('<root><ID>NULL</ID></root>');
    End;
    Select Updatexml(v_Define, '/root/ID/text()', Id_In).Getstringval() Into v_Value From Dual;
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_001', v_Value);
  End Zlhis_Dict_001;
  --�޸Ĳ���
  Procedure Zlhis_Dict_002(����id_In ���ű�.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_002', v_Value);
  End Zlhis_Dict_002;
  --ͣ�ò���
  Procedure Zlhis_Dict_003(����id_In ���ű�.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_003', v_Value);
  End Zlhis_Dict_003;
  --���ò���
  Procedure Zlhis_Dict_004(����id_In ���ű�.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_004', v_Value);
  End Zlhis_Dict_004;
  --������Ա
  Procedure Zlhis_Dict_005(��Աid_In ��Ա��.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><��ԱID>' || ��Աid_In || '</��ԱID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_005', v_Value);
  End Zlhis_Dict_005;
  --�޸���Ա
  Procedure Zlhis_Dict_006(��Աid_In ��Ա��.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><��ԱID>' || ��Աid_In || '</��ԱID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_006', v_Value);
  End Zlhis_Dict_006;
  --ͣ����Ա
  Procedure Zlhis_Dict_007(��Աid_In ��Ա��.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><��ԱID>' || ��Աid_In || '</��ԱID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_007', v_Value);
  End Zlhis_Dict_007;
  --������Ա
  Procedure Zlhis_Dict_008(��Աid_In ��Ա��.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><��ԱID>' || ��Աid_In || '</��ԱID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_008', v_Value);
  End Zlhis_Dict_008;
  --�����շ���Ŀ
  Procedure Zlhis_Dict_009(ϸĿid_In �շ���ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ϸĿID>' || ϸĿid_In || '</ϸĿID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_009', v_Value);
  End Zlhis_Dict_009;
  --�޸��շ���Ŀ
  Procedure Zlhis_Dict_010(ϸĿid_In �շ���ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ϸĿID>' || ϸĿid_In || '</ϸĿID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_010', v_Value);
  End Zlhis_Dict_010;
  --ͣ���շ���Ŀ
  Procedure Zlhis_Dict_011(ϸĿid_In �շ���ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ϸĿID>' || ϸĿid_In || '</ϸĿID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_011', v_Value);
  End Zlhis_Dict_011;
  --�����շ���Ŀ
  Procedure Zlhis_Dict_012(ϸĿid_In �շ���ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ϸĿID>' || ϸĿid_In || '</ϸĿID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_012', v_Value);
  End Zlhis_Dict_012;
  --����������Ŀ
  Procedure Zlhis_Dict_013(����id_In ������ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_013', v_Value);
  End Zlhis_Dict_013;
  --�޸�������Ŀ
  Procedure Zlhis_Dict_014(����id_In ������ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_014', v_Value);
  End Zlhis_Dict_014;
  --ͣ��������Ŀ
  Procedure Zlhis_Dict_015(����id_In ������ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_015', v_Value);
  End Zlhis_Dict_015;
  --����������Ŀ
  Procedure Zlhis_Dict_016(����id_In ������ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_016', v_Value);
  End Zlhis_Dict_016;
  --����������Ŀ
  Procedure Zlhis_Dict_017(����id_In ������ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><ϵͳ>1</ϵͳ></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_017', v_Value);
  End Zlhis_Dict_017;
  --�޸ļ�����Ŀ
  Procedure Zlhis_Dict_018(����id_In ������ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><ϵͳ>1</ϵͳ></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_018', v_Value);
  End Zlhis_Dict_018;
  --ɾ��������Ŀ
  Procedure Zlhis_Dict_019
  (
    ����id_In ������ĿĿ¼.Id%Type,
    ����_In   ����������Ŀ.����%Type,
    ������_In ����������Ŀ.������%Type,
    Ӣ����_In ����������Ŀ.Ӣ����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID>' || '<����>' || ����_In || '</����>' || '<������>' || ������_In || '</������>' ||
               '<Ӣ����>' || Ӣ����_In || '</Ӣ����>' || '<ϵͳ>1</ϵͳ></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_019', v_Value);
  End Zlhis_Dict_019;
  --������������Ŀ¼
  Procedure Zlhis_Dict_021(����id_In ��������Ŀ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_021', v_Value);
  End Zlhis_Dict_021;
  --�޸ļ�������Ŀ¼
  Procedure Zlhis_Dict_022(����id_In ��������Ŀ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_022', v_Value);
  End Zlhis_Dict_022;
  --ͣ�ü�������Ŀ¼
  Procedure Zlhis_Dict_023(����id_In ��������Ŀ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_023', v_Value);
  End Zlhis_Dict_023;
  --���ü�������Ŀ¼
  Procedure Zlhis_Dict_024(����id_In ��������Ŀ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_024', v_Value);
  End Zlhis_Dict_024;
  --����ҩƷ����
  Procedure Zlhis_Dict_025
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_025', v_Value);
  End Zlhis_Dict_025;
  --�޸�ҩƷ����
  Procedure Zlhis_Dict_026
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_026', v_Value);
  End Zlhis_Dict_026;
  --ɾ��ҩƷ����
  Procedure Zlhis_Dict_027
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type,
    ����_In ���Ʒ���Ŀ¼.����%Type,
    ����_In ���Ʒ���Ŀ¼.����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID><����>' || ����_In || '</����><����>' || ����_In ||
               '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_027', v_Value);
  End Zlhis_Dict_027;
  --ͣ��ҩƷ����
  Procedure Zlhis_Dict_028
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_028', v_Value);
  End Zlhis_Dict_028;
  --����ҩƷ����
  Procedure Zlhis_Dict_029
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_029', v_Value);
  End Zlhis_Dict_029;
  --����ҩƷƷ��
  Procedure Zlhis_Dict_030
  (
    ���_In ������ĿĿ¼.���%Type,
    Id_In   ������ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_030', v_Value);
  End Zlhis_Dict_030;
  --�޸�ҩƷƷ��
  Procedure Zlhis_Dict_031
  (
    ���_In ������ĿĿ¼.���%Type,
    Id_In   ������ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_031', v_Value);
  End Zlhis_Dict_031;
  --ɾ��ҩƷƷ��
  Procedure Zlhis_Dict_032
  (
    ���_In ������ĿĿ¼.���%Type,
    Id_In   ������ĿĿ¼.Id%Type,
    ����_In ������ĿĿ¼.����%Type,
    ����_In ������ĿĿ¼.����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ID>' || Id_In || '</ID><����>' || ����_In || '</����><����>' || ����_In ||
               '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_032', v_Value);
  End Zlhis_Dict_032;
  --ͣ��ҩƷƷ��
  Procedure Zlhis_Dict_033
  (
    ���_In ������ĿĿ¼.���%Type,
    Id_In   ������ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_033', v_Value);
  End Zlhis_Dict_033;
  --����ҩƷƷ��
  Procedure Zlhis_Dict_034
  (
    ���_In ������ĿĿ¼.���%Type,
    Id_In   ������ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_034', v_Value);
  End Zlhis_Dict_034;
  --����ҩƷ���
  Procedure Zlhis_Dict_035
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ҩƷID>' || ҩƷid_In || '</ҩƷID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_035', v_Value);
  End Zlhis_Dict_035;
  --�޸�ҩƷ���
  Procedure Zlhis_Dict_036
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ҩƷID>' || ҩƷid_In || '</ҩƷID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_036', v_Value);
  End Zlhis_Dict_036;
  --ɾ��ҩƷ���
  Procedure Zlhis_Dict_037
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type,
    ����_In   �շ���ĿĿ¼.����%Type,
    ����_In   �շ���ĿĿ¼.����%Type,
    ���_In   �շ���ĿĿ¼.���%Type,
    ����_In   �շ���ĿĿ¼.����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ҩƷID>' || ҩƷid_In || '</ҩƷID><����>' || ����_In || '</����><����>' || ����_In ||
               '</����><���>' || ���_In || '</���><����>' || ����_In || '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_037', v_Value);
  End Zlhis_Dict_037;
  --ͣ��ҩƷ���
  Procedure Zlhis_Dict_038
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ҩƷID>' || ҩƷid_In || '</ҩƷID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_038', v_Value);
  End Zlhis_Dict_038;
  --����ҩƷ���
  Procedure Zlhis_Dict_039
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ҩƷID>' || ҩƷid_In || '</ҩƷID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_039', v_Value);
  End Zlhis_Dict_039;
  --����ҩƷ�洢�ⷿ
  Procedure Zlhis_Dict_040
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ҩƷID>' || ҩƷid_In || '</ҩƷID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_040', v_Value);
  End Zlhis_Dict_040;
  --����ҩƷ�����޶�
  Procedure Zlhis_Dict_041
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ҩƷID>' || ҩƷid_In || '</ҩƷID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_041', v_Value);
  End Zlhis_Dict_041;
  --��������Ʒ��
  Procedure Zlhis_Dict_042(Id_In ������ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || Id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_042', v_Value);
  End Zlhis_Dict_042;
  --�������Ĺ��
  Procedure Zlhis_Dict_043(Id_In �շ���ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || Id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_043', v_Value);
  End Zlhis_Dict_043;
  --�޸����Ĺ��
  Procedure Zlhis_Dict_044(Id_In �շ���ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || Id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_044', v_Value);
  End Zlhis_Dict_044;
  --ɾ�����Ĺ��
  Procedure Zlhis_Dict_045
  (
    Id_In   �շ���ĿĿ¼.Id%Type,
    ����_In �շ���ĿĿ¼.����%Type,
    ����_In �շ���ĿĿ¼.����%Type,
    ���_In �շ���ĿĿ¼.���%Type,
    ����_In �շ���ĿĿ¼.����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || Id_In || '</����ID><����>' || ����_In || '</����><����>' || ����_In || '</����><���>' || ���_In ||
               '</���><����>' || ����_In || '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_045', v_Value);
  End Zlhis_Dict_045;
  --ͣ�����Ĺ��
  Procedure Zlhis_Dict_046(Id_In �շ���ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || Id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_046', v_Value);
  End Zlhis_Dict_046;
  --�������Ĺ��
  Procedure Zlhis_Dict_047(Id_In �շ���ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || Id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_047', v_Value);
  End Zlhis_Dict_047;
  --ҽ������
  Procedure Zlhis_Dict_048
  (
    ����_In       In ����֧����Ŀ.����%Type,
    �շ�ϸĿid_In In ����֧����Ŀ.�շ�ϸĿid%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><�շ�ϸĿID>' || �շ�ϸĿid_In || '</�շ�ϸĿID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_048', v_Value);
  End Zlhis_Dict_048;
  --ɾ��ҽ������
  Procedure Zlhis_Dict_049
  (
    ����_In       In ����֧����Ŀ.����%Type,
    �շ�ϸĿid_In In ����֧����Ŀ.�շ�ϸĿid%Type,
    ��Ŀ����_In   In �շ���ĿĿ¼.����%Type,
    ��Ŀ����_In   In �շ���ĿĿ¼.����%Type,
    ҽ������_In   In ����֧����Ŀ.��Ŀ����%Type,
    ҽ������_In   In ����֧����Ŀ.��Ŀ����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><�շ�ϸĿID>' || �շ�ϸĿid_In || '</�շ�ϸĿID><��Ŀ����>' || ��Ŀ����_In || '</��Ŀ����><��Ŀ����>' ||
               ��Ŀ����_In || '</��Ŀ����><ҽ������>' || ҽ������_In || '</ҽ������><ҽ������>' || ҽ������_In || '</ҽ������></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_049', v_Value);
  End Zlhis_Dict_049;
  --�������ķ���
  Procedure Zlhis_Dict_050
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_050', v_Value);
  End Zlhis_Dict_050;
  --�޸����ķ���
  Procedure Zlhis_Dict_051
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_051', v_Value);
  End Zlhis_Dict_051;
  --ɾ�����ķ���
  Procedure Zlhis_Dict_052
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type,
    ����_In ���Ʒ���Ŀ¼.����%Type,
    ����_In ���Ʒ���Ŀ¼.����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID><����>' || ����_In || '</����><����>' || ����_In ||
               '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_052', v_Value);
  End Zlhis_Dict_052;
  --�շѼ�Ŀ�䶯
  Procedure Zlhis_Dict_053
  (
    �շ���ĿId_In       �շ���ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�շ���ĿID>' || �շ���ĿId_In || '</�շ���ĿID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_053', v_Value);
  End Zlhis_Dict_053;

  --�����շѶ��ձ䶯
  Procedure Zlhis_Dict_054
  (
    ������ĿId_In     ���Ʒ���Ŀ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><������ĿID>' || ������ĿId_In || '</������ĿID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_054', v_Value);
  End Zlhis_Dict_054;
  --�������Ƽ������
  Procedure Zlhis_Dictpacs_001
  (
    ����_In   ���Ƽ������.����%Type,
    ����_In   ���Ƽ������.����%Type,
    ����_In   ���Ƽ������.����%Type,
    ������_In ���Ƽ������.������%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><������>' || ������_In ||
               '</������></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_001', v_Value);
  End Zlhis_Dictpacs_001;

  --�޸����Ƽ������
  Procedure Zlhis_Dictpacs_002
  (
    ����_In   ���Ƽ������.����%Type,
    ����_In   ���Ƽ������.����%Type,
    ����_In   ���Ƽ������.����%Type,
    ������_In ���Ƽ������.������%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><������>' || ������_In ||
               '</������></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_002', v_Value);
  End Zlhis_Dictpacs_002;
  --ɾ�����Ƽ������
  Procedure Zlhis_Dictpacs_003
  (
    ����_In   ���Ƽ������.����%Type,
    ����_In   ���Ƽ������.����%Type,
    ����_In   ���Ƽ������.����%Type,
    ������_In ���Ƽ������.������%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><������>' || ������_In ||
               '</������></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_003', v_Value);
  End Zlhis_Dictpacs_003;
  --�������Ƽ�鲿λ
  Procedure Zlhis_Dictpacs_004
  (
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ��ע_In     ���Ƽ�鲿λ.��ע%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    �����Ա�_In ���Ƽ�鲿λ.�����Ա�%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In ||
               '</����><��ע>' || ��ע_In || '</��ע><����>' || ����_In || '</����><�����Ա�>' || �����Ա�_In || '</�����Ա�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_004', v_Value);
  End Zlhis_Dictpacs_004;
  --�޸����Ƽ�鲿λ
  Procedure Zlhis_Dictpacs_005
  (
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ��ע_In     ���Ƽ�鲿λ.��ע%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    �����Ա�_In ���Ƽ�鲿λ.�����Ա�%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In ||
               '</����><��ע>' || ��ע_In || '</��ע><����>' || ����_In || '</����><�����Ա�>' || �����Ա�_In || '</�����Ա�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_005', v_Value);
  End Zlhis_Dictpacs_005;
  --ɾ�����Ƽ�鲿λ
  Procedure Zlhis_Dictpacs_006
  (
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ��ע_In     ���Ƽ�鲿λ.��ע%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    �����Ա�_In ���Ƽ�鲿λ.�����Ա�%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In ||
               '</����><��ע>' || ��ע_In || '</��ע><����>' || ����_In || '</����><�����Ա�>' || �����Ա�_In || '</�����Ա�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_006', v_Value);
  End Zlhis_Dictpacs_006;
  --����������Ŀ��λ
  Procedure Zlhis_Dictpacs_007
  (
    Id_In     ������Ŀ��λ.Id%Type,
    ��Ŀid_In ������Ŀ��λ.��Ŀid%Type,
    ����_In   ������Ŀ��λ.����%Type,
    ��λ_In   ������Ŀ��λ.��λ%Type,
    ����_In   ������Ŀ��λ.����%Type,
    Ĭ��_In   ������Ŀ��λ.Ĭ��%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ID>' || Id_In || '</ID><��ĿID>' || ��Ŀid_In || '</��ĿID><����>' || ����_In || '</����><��λ>' || ��λ_In ||
               '</��λ><����>' || ����_In || '</����><Ĭ��>' || Ĭ��_In || '</Ĭ��></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_007', v_Value);
  End Zlhis_Dictpacs_007;
  --�޸�������Ŀ��λ
  Procedure Zlhis_Dictpacs_008
  (
    Id_In     ������Ŀ��λ.Id%Type,
    ��Ŀid_In ������Ŀ��λ.��Ŀid%Type,
    ����_In   ������Ŀ��λ.����%Type,
    ��λ_In   ������Ŀ��λ.��λ%Type,
    ����_In   ������Ŀ��λ.����%Type,
    Ĭ��_In   ������Ŀ��λ.Ĭ��%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ID>' || Id_In || '</ID><��ĿID>' || ��Ŀid_In || '</��ĿID><����>' || ����_In || '</����><��λ>' || ��λ_In ||
               '</��λ><����>' || ����_In || '</����><Ĭ��>' || Ĭ��_In || '</Ĭ��></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_008', v_Value);
  End Zlhis_Dictpacs_008;
  --ɾ��������Ŀ��λ
  Procedure Zlhis_Dictpacs_009
  (
    Id_In     ������Ŀ��λ.Id%Type,
    ��Ŀid_In ������Ŀ��λ.��Ŀid%Type,
    ����_In   ������Ŀ��λ.����%Type,
    ��λ_In   ������Ŀ��λ.��λ%Type,
    ����_In   ������Ŀ��λ.����%Type,
    Ĭ��_In   ������Ŀ��λ.Ĭ��%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ID>' || Id_In || '</ID><��ĿID>' || ��Ŀid_In || '</��ĿID><����>' || ����_In || '</����><��λ>' || ��λ_In ||
               '</��λ><����>' || ����_In || '</����><Ĭ��>' || Ĭ��_In || '</Ĭ��></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_009', v_Value);
  End Zlhis_Dictpacs_009;
  --����������Ŀ��λ
  Procedure Zlhis_Dictlis_004
  (
    ����_In     ���Ƽ���걾.����%Type,
    ����_In     ���Ƽ���걾.����%Type,
    ����_In     ���Ƽ���걾.����%Type,
    �����Ա�_In ���Ƽ���걾.�����Ա�%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><�����Ա�>' || �����Ա�_In ||
               '</�����Ա�></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_004', v_Value);
  End Zlhis_Dictlis_004;
  --�޸�������Ŀ��λ
  Procedure Zlhis_Dictlis_005
  (
    ����_In     ���Ƽ���걾.����%Type,
    ����_In     ���Ƽ���걾.����%Type,
    ����_In     ���Ƽ���걾.����%Type,
    �����Ա�_In ���Ƽ���걾.�����Ա�%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><�����Ա�>' || �����Ա�_In ||
               '</�����Ա�></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_005', v_Value);
  End Zlhis_Dictlis_005;
  --ɾ��������Ŀ��λ
  Procedure Zlhis_Dictlis_006
  (
    ����_In     ���Ƽ���걾.����%Type,
    ����_In     ���Ƽ���걾.����%Type,
    ����_In     ���Ƽ���걾.����%Type,
    �����Ա�_In ���Ƽ���걾.�����Ա�%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><�����Ա�>' || �����Ա�_In ||
               '</�����Ա�></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_006', v_Value);
  End Zlhis_Dictlis_006;
  --������Ѫ������
  Procedure Zlhis_Dictlis_007
  (
    ����_In   ��Ѫ������.����%Type,
    ����_In   ��Ѫ������.����%Type,
    ����_In   ��Ѫ������.����%Type,
    ��Ӽ�_In ��Ѫ������.��Ӽ�%Type,
    ��Ѫ��_In ��Ѫ������.��Ѫ��%Type,
    ���_In   ��Ѫ������.���%Type,
    ��ɫ_In   ��Ѫ������.��ɫ%Type,
    ����id_In ��Ѫ������.����id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><��Ӽ�>' || ��Ӽ�_In ||
               '</��Ӽ�><��Ѫ��>' || ��Ѫ��_In || '</��Ѫ��><��ɫ>' || ��ɫ_In || '</��ɫ><���>' || ���_In || '</���><����ID_In>' || ����id_In ||
               '</����ID_In></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_007', v_Value);
  End Zlhis_Dictlis_007;
  --������Ѫ������
  Procedure Zlhis_Dictlis_008
  (
    ����_In   ��Ѫ������.����%Type,
    ����_In   ��Ѫ������.����%Type,
    ����_In   ��Ѫ������.����%Type,
    ��Ӽ�_In ��Ѫ������.��Ӽ�%Type,
    ��Ѫ��_In ��Ѫ������.��Ѫ��%Type,
    ���_In   ��Ѫ������.���%Type,
    ��ɫ_In   ��Ѫ������.��ɫ%Type,
    ����id_In ��Ѫ������.����id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><��Ӽ�>' || ��Ӽ�_In ||
               '</��Ӽ�><��Ѫ��>' || ��Ѫ��_In || '</��Ѫ��><��ɫ>' || ��ɫ_In || '</��ɫ><���>' || ���_In || '</���><����ID_In>' || ����id_In ||
               '</����ID_In></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_008', v_Value);
  End Zlhis_Dictlis_008;
  --������Ѫ������
  Procedure Zlhis_Dictlis_009
  (
    ����_In   ��Ѫ������.����%Type,
    ����_In   ��Ѫ������.����%Type,
    ����_In   ��Ѫ������.����%Type,
    ��Ӽ�_In ��Ѫ������.��Ӽ�%Type,
    ��Ѫ��_In ��Ѫ������.��Ѫ��%Type,
    ���_In   ��Ѫ������.���%Type,
    ��ɫ_In   ��Ѫ������.��ɫ%Type,
    ����id_In ��Ѫ������.����id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><��Ӽ�>' || ��Ӽ�_In ||
               '</��Ӽ�><��Ѫ��>' || ��Ѫ��_In || '</��Ѫ��><��ɫ>' || ��ɫ_In || '</��ɫ><���>' || ���_In || '</���><����ID_In>' || ����id_In ||
               '</����ID_In></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_009', v_Value);
  End Zlhis_Dictlis_009;
  --ҩƷ��ҩ����
  Procedure Zlhis_Drug_001(No_In ҩƷ�շ���¼.No%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���ݺ�>' || No_In || '</���ݺ�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_001', v_Value);
  End Zlhis_Drug_001;
  --ȡ��ҩƷ��ҩ����
  Procedure Zlhis_Drug_002(No_In ҩƷ�շ���¼.No%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���ݺ�>' || No_In || '</���ݺ�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_002', v_Value);
  End Zlhis_Drug_002;
  --ҩƷ�ƿⵥ����
  Procedure Zlhis_Drug_003(No_In ҩƷ�շ���¼.No%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���ݺ�>' || No_In || '</���ݺ�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_003', v_Value);
  End Zlhis_Drug_003;
  --ҩƷ�ƿⵥ����
  Procedure Zlhis_Drug_004(No_In ҩƷ�շ���¼.No%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���ݺ�>' || No_In || '</���ݺ�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_004', v_Value);
  End Zlhis_Drug_004;
  --���ŷ�ҩ
  Procedure Zlhis_Drug_005
  (
    �ⷿid_In ҩƷ�շ���¼.�ⷿid%Type,
    �շ�id_In ҩƷ�շ���¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�ⷿID>' || �ⷿid_In || '</�ⷿID><�շ�ID>' || �շ�id_In || '</�շ�ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_005', v_Value);
  End Zlhis_Drug_005;
  --������ҩ
  Procedure Zlhis_Drug_006
  (
    �����շ�id_In ҩƷ�շ���¼.Id%Type,
    �����շ�id_In ҩƷ�շ���¼.Id%Type,
    ����_In       ҩƷ�շ���¼.ʵ������%Type,
    ����id_In     ������ü�¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><������¼ID>' || �����շ�id_In || '</������¼ID><������¼ID>' || �����շ�id_In || '</������¼ID><����>' || ����_In ||
               '</����><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_006', v_Value);
  End Zlhis_Drug_006;
  --ҩƷ����
  Procedure Zlhis_Drug_007(�۸�id_In ҩƷ�۸��¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�۸�ID>' || �۸�id_In || '</�۸�ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_007', v_Value);
  End Zlhis_Drug_007;
  --���䷢��
  Procedure Zlhis_Drug_008(��¼ids_In Varchar2) Is
    v_Value  Zlmsg_Todo.Key_Value%Type;
    n_��¼id ��Һ��ҩ��¼.Id%Type;
    v_Tmp    Varchar2(4000);
	n_Length Number(18);
  Begin
    If ��¼ids_In Is Null Then
      v_Tmp := Null;
    Else
      v_Tmp := ��¼ids_In || ',';
    End If;
  
    v_Value := '<root><��¼IDS>';
  
    While v_Tmp Is Not Null Loop
      --�ֽⵥ��ID��
      n_��¼id := To_Number(Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1));
      v_Tmp    := Replace(',' || v_Tmp, ',' || n_��¼id || ',');
      
      --�жϵ�ǰ�����Ƿ񼴽���������                                                                        
      Select Lengthb(v_Value || '<��¼ID>' || n_��¼id || '</��¼ID>') Into n_Length From Dual;            
      If n_Length > 950 Then								                   
        v_Value := v_Value || '</��¼IDs></root>';                                                         
        b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_008', v_Value);                                            
        v_Value := '<root><��¼IDs>';                                                                      
      End If;

      v_Value := v_Value || '<��¼ID>' || n_��¼id || '</��¼ID>';
    End Loop;
  
    v_Value := v_Value || '</��¼IDS></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_008', v_Value);
  End Zlhis_Drug_008;
  --ҩƷ���ۼ�
  Procedure Zlhis_Drug_009
  (
    �۸�id_In ҩƷ�۸��¼.Id%Type,
    ʱ��_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�۸�ID>' || �۸�id_In || '</�۸�ID><ʱ��>' || ʱ��_In || '</ʱ��></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_009', v_Value);
  End Zlhis_Drug_009;
  --���ĵ��ɱ���
  Procedure Zlhis_Drug_010(�۸�id_In �ɱ��۵�����Ϣ.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�۸�ID>' || �۸�id_In || '</�۸�ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_010', v_Value);
  End Zlhis_Drug_010;
  --���ĵ��ۼ�
  Procedure Zlhis_Drug_011
  (
    �۸�id_In �շѼ�Ŀ.Id%Type,
    ʱ��_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�۸�ID>' || �۸�id_In || '</�۸�ID><ʱ��>' || ʱ��_In || '</ʱ��></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_011', v_Value);
  End Zlhis_Drug_011;

  --2.ֹͣ����ҽ����סԺ
  Procedure Zlhis_Cis_002
  (
    ����id_In  In ����ҽ����¼.����id%Type,
    ��ҳid_In  In ����ҽ����¼.��ҳid%Type,
    ҽ��id_In  In ����ҽ����¼.Id%Type,
    ҽ��ids_In In Varchar2
  ) Is
  Begin
    If p_Msg_Using('ZLHIS_CIS_002') = 0 Then
      Return;
    End If;
    If ҽ��id_In Is Not Null Then
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_002',
                                  '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><ID>' || ҽ��id_In ||
                                   '</ID></root>');
    Else
      For R In (Select '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><ID>' || ID || '</ID></root>' As Xml_Value
                From ����ҽ����¼
                Where ID In (Select Column_Value From Table(f_Num2list(ҽ��ids_In))) And ���id Is Null) Loop
        b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_002', r.Xml_Value);
      End Loop;
    End If;
  End Zlhis_Cis_002;
  --3.���ϻ���ҽ��������/סԺ
  Procedure Zlhis_Cis_003
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><ID>' ||
               ҽ��id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_003', v_Value);
  End Zlhis_Cis_003;

  --4.��������ҽ����סԺ
  Procedure Zlhis_Cis_004
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_004',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><ID>' || ҽ��id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_004;

  --5.������������ҽ����סԺ
  Procedure Zlhis_Cis_005
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_005',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><ID>' || ҽ��id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_005;

  --6.���߻�����ҽ����סԺ
  Procedure Zlhis_Cis_006
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_006',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In ||
                                 '</���ͺ�><ID>' || ҽ��id_In || '</ID></root>');
  End Zlhis_Cis_006;

  --7.�������߻�����ҽ����סԺ
  Procedure Zlhis_Cis_007
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_007', v_Value);
  End Zlhis_Cis_007;

  --���ﻼ�߽���
  Procedure Zlhis_Cis_008
  (
    ����id_In In ����ҽ����¼.����id%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_008', '<root><����ID>' || ����id_In || '</����ID><NO>' || �Һŵ�_In || '</NO></root>');
  End Zlhis_Cis_008;

  --���ﻼ��ȡ������
  Procedure Zlhis_Cis_009
  (
    ����id_In In ����ҽ����¼.����id%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_009', '<root><����ID>' || ����id_In || '</����ID><NO>' || �Һŵ�_In || '</NO></root>');
  End Zlhis_Cis_009;

  --10.�´ﻼ����ϣ�����/סԺ
  Procedure Zlhis_Cis_010
  (
    ����id_In In ���˹Һż�¼.����id%Type,
    ����id_In In ���˹Һż�¼.Id%Type, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    ���id_In In ������ϼ�¼.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_010',
                                '<root><����ID>' || ����id_In || '</����ID><����ID>' || ����id_In || '</����ID><ID>' || ���id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_010;
  --11.�����������
  Procedure Zlhis_Cis_011
  (
    ����id_In   In ���˹Һż�¼.����id%Type,
    ����id_In   In ���˹Һż�¼.Id%Type, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    Id_In       In ������ϼ�¼.Id%Type,
    ����id_In   In ������ϼ�¼.����id%Type,
    ���id_In   In ������ϼ�¼.���id%Type,
    �������_In In ������ϼ�¼.�������%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><����ID>' || ����id_In || '</����ID><ID>' || Id_In || '</ID><����ID>' ||
               ����id_In || '</����ID><���ID>' || ���id_In || '</���ID><�������>' || �������_In || '</�������></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_011', v_Value);
  End Zlhis_Cis_011;

  --����ִ��ҽ��У��
  Procedure Zlhis_Cis_012
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_012',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><ID>' || ҽ��id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_012;

  --13.����Σ��ֵ�Ķ�֪ͨ
  Procedure Zlhis_Cis_014
  (
    ����id_In In ���˹Һż�¼.����id%Type,
    ����id_In In ���˹Һż�¼.Id%Type, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    ҽ��id_In In ����ҽ����¼.Id%Type,
    ��Ϣid_In In ҵ����Ϣ�嵥.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><����ID>' || ����id_In || '</����ID><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ID>' ||
               ��Ϣid_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_014', v_Value);
  End Zlhis_Cis_014;
  --15.���߼������룬����/סԺ
  Procedure Zlhis_Cis_016
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    ������Դ_In In ����ҽ����¼.������Դ%Type --1-����;2-סԺ;3-����(������ڸ��ﲿ�Ž�����������);4-��첡��
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><������Դ>' || ������Դ_In || '</������Դ></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_016', v_Value);
  End Zlhis_Cis_016;
  --16.���߼�����룬����/סԺ
  Procedure Zlhis_Cis_017
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    ������Դ_In In ����ҽ����¼.������Դ%Type --1-����;2-סԺ;3-����(������ڸ��ﲿ�Ž�����������);4-��첡��
  ) Is
    v_Value    Zlmsg_Todo.Key_Value%Type;
    v_�������� ������ĿĿ¼.��������%Type;
  Begin
    Select Max(a.��������)
    Into v_��������
    From ������ĿĿ¼ A, ����ҽ����¼ B
    Where b.������Ŀid = a.Id And b.Id = ҽ��id_In;
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><������Դ>' || ������Դ_In || '</������Դ></root>';
    If v_�������� = '����' Then
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_054', v_Value);
    Else
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_017', v_Value);
    End If;
  End Zlhis_Cis_017;
  --17.�����������룬����/סԺ
  Procedure Zlhis_Cis_018
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_018', v_Value);
  End Zlhis_Cis_018;
  --18.������Ѫ���룬סԺ
  Procedure Zlhis_Cis_019
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_019', v_Value);
  End Zlhis_Cis_019;
  --19.���߻������룬סԺ
  Procedure Zlhis_Cis_020
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_020', v_Value);
  End Zlhis_Cis_020;
  --20.��������ҽ����סԺ
  Procedure Zlhis_Cis_021
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_021', v_Value);
  End Zlhis_Cis_021;
  --21.��������ҽ����סԺ
  Procedure Zlhis_Cis_022
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_022', v_Value);
  End Zlhis_Cis_022;
  --22.������������ҽ����סԺ
  Procedure Zlhis_Cis_023
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_023', v_Value);
  End Zlhis_Cis_023;

  --24.���Σ��ֵ�Ķ�֪ͨ
  Procedure Zlhis_Cis_025
  (
    ����id_In In ���˹Һż�¼.����id%Type,
    ����id_In In ���˹Һż�¼.Id%Type, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    ҽ��id_In In ����ҽ����¼.Id%Type,
    ��Ϣid_In In ҵ����Ϣ�嵥.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><����ID>' || ����id_In || '</����ID><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ID>' ||
               ��Ϣid_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_025', v_Value);
  End Zlhis_Cis_025;

  --����ִ��ҽ������
  Procedure Zlhis_Cis_026
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_026',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In ||
                                 '</���ͺ�><ID>' || ҽ��id_In || '</ID></root>');
  End Zlhis_Cis_026;

  --�������߼�������
  Procedure Zlhis_Cis_036
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    No_In       In ����ҽ������.No%Type,
    ������Դ_In In ����ҽ����¼.������Դ%Type --1-����;2-סԺ;3-����(������ڸ��ﲿ�Ž�����������);4-��첡��
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><NO>' || No_In || '</NO><������Դ>' || ������Դ_In ||
               '</������Դ></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_036', v_Value);
  End Zlhis_Cis_036;

  --�������߼������
  Procedure Zlhis_Cis_037
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    No_In       In ����ҽ������.No%Type,
    ������Դ_In In ����ҽ����¼.������Դ%Type --1-����;2-סԺ;3-����(������ڸ��ﲿ�Ž�����������);4-��첡��
  ) Is
    v_Value    Zlmsg_Todo.Key_Value%Type;
    v_�������� ������ĿĿ¼.��������%Type;
  Begin
    Select Max(a.��������)
    Into v_��������
    From ������ĿĿ¼ A, ����ҽ����¼ B
    Where b.������Ŀid = a.Id And b.Id = ҽ��id_In;
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><NO>' || No_In || '</NO><������Դ>' || ������Դ_In ||
               '</������Դ></root>';
    If v_�������� = '����' Then
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_055', v_Value);
    Else
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_037', v_Value);
    End If;
  End Zlhis_Cis_037;

  --����������������
  Procedure Zlhis_Cis_038
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_038', v_Value);
  End Zlhis_Cis_038;

  --����������Ѫ����
  Procedure Zlhis_Cis_039
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_039', v_Value);
  End Zlhis_Cis_039;

  --�������߻�������
  Procedure Zlhis_Cis_040
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_040', v_Value);
  End Zlhis_Cis_040;

  --������������ҽ��
  Procedure Zlhis_Cis_041
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_041', v_Value);
  End Zlhis_Cis_041;

  --������������ҽ��
  Procedure Zlhis_Cis_042
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_042', v_Value);
  End Zlhis_Cis_042;

  --������������ҽ��
  Procedure Zlhis_Cis_043
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_043', v_Value);
  End Zlhis_Cis_043;

  --��������ִ��ҽ��
  Procedure Zlhis_Cis_044
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    No_In       In ����ҽ������.No%Type,
    ��������_In In ����ҽ������.��������%Type,
    �״�ʱ��_In In ����ҽ������.�״�ʱ��%Type,
    ĩ��ʱ��_In In ����ҽ������.ĩ��ʱ��%Type,
    ��������_In In ����ҽ������.��������%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID><NO>' || No_In || '</NO><��������>' || ��������_In || '</��������><�״�ʱ��>' ||
               To_Char(�״�ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '</�״�ʱ��><ĩ��ʱ��>' ||
               To_Char(ĩ��ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '</ĩ��ʱ��><��������>' || ��������_In || '</��������></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_044', v_Value);
  End Zlhis_Cis_044;

  --����ҽ��ִ�еǼ�
  Procedure Zlhis_Cis_050
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    Ҫ��ʱ��_In In ����ҽ��ִ��.Ҫ��ʱ��%Type,
    ִ��ʱ��_In In ����ҽ��ִ��.ִ��ʱ��%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><Ҫ��ʱ��>' || To_Char(Ҫ��ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') ||
               '</Ҫ��ʱ��><ִ��ʱ��>' || To_Char(ִ��ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '</ִ��ʱ��></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_050', v_Value);
  End Zlhis_Cis_050;

  --����ҽ��ȡ��ִ�еǼ�
  Procedure Zlhis_Cis_051
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    Ҫ��ʱ��_In In ����ҽ��ִ��.Ҫ��ʱ��%Type,
    ִ��ʱ��_In In ����ҽ��ִ��.ִ��ʱ��%Type,
    ��������_In In ����ҽ��ִ��.��������%Type,
    ִ�н��_In In ����ҽ��ִ��.ִ�н��%Type,
    ִ��ժҪ_In In ����ҽ��ִ��.ִ��ժҪ%Type,
    ִ�п���_In In ����ҽ��ִ��.ִ�п���id%Type,
    ִ����_In   In ����ҽ��ִ��.ִ����%Type,
    �˶���_In   In ����ҽ��ִ��.�˶���%Type,
    ��¼��Դ_In In ����ҽ��ִ��.��¼��Դ%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><Ҫ��ʱ��>' || To_Char(Ҫ��ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') ||
               '</Ҫ��ʱ��><ִ��ʱ��>' || To_Char(ִ��ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '</ִ��ʱ��><��������>' || ��������_In ||
               '</��������><ִ�н��>' || ִ�н��_In || '</ִ�н��><ִ��ժҪ>' || ִ��ժҪ_In || '</ִ��ժҪ><ִ�п���ID>' || ִ�п���_In ||
               '</ִ�п���ID><ִ����>' || ִ����_In || '</ִ����><�˶���>' || �˶���_In || '</�˶���><��¼��Դ>' || ��¼��Դ_In || '</��¼��Դ></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_051', v_Value);
  End Zlhis_Cis_051;
  --����ҽ��ִ�����
  Procedure Zlhis_Cis_052
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_052',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In ||
                                 '</�Һŵ�><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID></root>');
  End Zlhis_Cis_052;
  --����ҽ������ִ�����
  Procedure Zlhis_Cis_053
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_053',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In ||
                                 '</�Һŵ�><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID></root>');
  End Zlhis_Cis_053;

  --�������뷢�ͺ��޸�
  Procedure Zlhis_Cis_056
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    ������Դ_In In ����ҽ����¼.������Դ%Type --1-����;2-סԺ;3-����(������ڸ��ﲿ�Ž�����������);4-��첡��
  ) Is
    v_Value    Zlmsg_Todo.Key_Value%Type;
    v_�������� ������ĿĿ¼.��������%Type;
  Begin
    Select Max(a.��������)
    Into v_��������
    From ������ĿĿ¼ A, ����ҽ����¼ B
    Where b.������Ŀid = a.Id And b.Id = ҽ��id_In;
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><������Դ>' || ������Դ_In || '</������Դ></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_056', v_Value);
  End Zlhis_Cis_056;

  --���ﻼ����ɾ���
  Procedure Zlhis_Cis_057
  (
    ����id_In In ����ҽ����¼.����id%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_057', '<root><����ID>' || ����id_In || '</����ID><NO>' || �Һŵ�_In || '</NO></root>');
  End Zlhis_Cis_057;

  --���ﻼ��ȡ����ɾ���
  Procedure Zlhis_Cis_058
  (
    ����id_In In ����ҽ����¼.����id%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_058', '<root><����ID>' || ����id_In || '</����ID><NO>' || �Һŵ�_In || '</NO></root>');
  End Zlhis_Cis_058;

  --ȷ��ֹͣ����ҽ�� 
  Procedure Zlhis_Cis_059
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_059','<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><ID>' || ҽ��id_In || '</ID></root>');
  End Zlhis_Cis_059;

  --26.��鱨����ɣ�������ʱ
  Procedure Zlhis_Pacs_001
  (
    ҽ��id_In   In Ӱ�����¼.ҽ��id%Type,
    ����id_Ins  In Varchar2,
    ��������_In In Number --1-�ϰ�PACS���棬2-�ϰ没���༭�����棬3-�°�༭������
  ) Is
  Begin
    If p_Msg_Using('ZLHIS_PACS_001') = 0 Then
      Return;
    End If;
    For R In (Select '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><����ID>' || Column_Value || '</����ID><��������>' || ��������_In ||
                      '<��������></root>' As Xml_Value
              From Table(f_Str2list(����id_Ins))) Loop
      b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_001', r.Xml_Value);
    End Loop;
  End Zlhis_Pacs_001;
  --27.���״̬ͬ�������״̬�ı��
  Procedure Zlhis_Pacs_002
  (
    ҽ��id_In In Ӱ�����¼.ҽ��id%Type,
    ԭ״̬_In In ����ҽ������.ִ�й���%Type,
    ��״̬_In In ����ҽ������.ִ�й���%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ԭ״̬>' || ԭ״̬_In || '</ԭ״̬><��״̬>' || ��״̬_In || '</��״̬></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_002', v_Value);
  End Zlhis_Pacs_002;
  --28.���״̬���ˣ����״̬���˺�
  Procedure Zlhis_Pacs_003
  (
    ҽ��id_In In Ӱ�����¼.ҽ��id%Type,
    ԭ״̬_In In ����ҽ������.ִ�й���%Type,
    ��״̬_In In ����ҽ������.ִ�й���%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ԭ״̬>' || ԭ״̬_In || '</ԭ״̬><��״̬>' || ��״̬_In || '</��״̬></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_003', v_Value);
  End Zlhis_Pacs_003;
  --29.��鱨�泷��������������ʱ
  Procedure Zlhis_Pacs_004
  (
    ҽ��id_In   In Ӱ�����¼.ҽ��id%Type,
    ����id_Ins  In Varchar2,
    ��������_In In Number --1-�ϰ�PACS���棬2-�ϰ没���༭�����棬3-�°�༭������
  ) Is
  Begin
    If p_Msg_Using('ZLHIS_PACS_004') = 0 Then
      Return;
    End If;
    For R In (Select '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><����ID>' || Column_Value || '</����ID><��������>' || ��������_In ||
                      '<��������></root>' As Xml_Value
              From Table(f_Str2list(����id_Ins))) Loop
      b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_004', r.Xml_Value);
    End Loop;
  End Zlhis_Pacs_004;
  --30.���Σ��ֵ֪ͨ����鷢��Σ��ֵʱ
  Procedure Zlhis_Pacs_005(ҽ��id_In In Ӱ�����¼.ҽ��id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_005', v_Value);
  End Zlhis_Pacs_005;
  -- ���ԤԼ֪ͨ�����ԤԼʱ
  Procedure Zlhis_Pacs_006
  (
    ҽ��id_In In Ӱ�����¼.ҽ��id%Type,
    ԤԼid_In In Ris���ԤԼ.ԤԼid%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ԤԼID>' || ԤԼid_In || '</ԤԼID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_006', v_Value);
  End Zlhis_Pacs_006;
  -- ȡ�����ԤԼ��ȡ��ԤԼʱ
  Procedure Zlhis_Pacs_007
  (
    ҽ��id_In       In Ӱ�����¼.ҽ��id%Type,
    ԤԼid_In       In Ris���ԤԼ.ԤԼid%Type,
    ԤԼ����_In     In Ris���ԤԼ.ԤԼ����%Type,
    ԤԼ���_In     In Ris���ԤԼ.���%Type,
    ����豸����_In In Ris���ԤԼ.����豸����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ԤԼID>' || ԤԼid_In || '</ԤԼID><ԤԼ����>' || ԤԼ����_In || '</ԤԼ����><ԤԼ���>' ||
               ԤԼ���_In || '</ԤԼ���><����豸����>' || ����豸����_In || '</����豸����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_007', v_Value);
  End Zlhis_Pacs_007;

  --36.���߷�����󶨿�
  Procedure Zlhis_Patient_018
  (
    �䶯id_In   In ���˱䶯��¼.Id%Type,
    ����id_In   In ������Ϣ.����id%Type,
    �����id_In In ҽ�ƿ����.Id%Type,
    ����_In     In ����ҽ�ƿ���Ϣ.����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�䶯ID>' || �䶯id_In || '</�䶯ID><����ID>' || ����id_In || '</����ID><�����ID>' || �����id_In ||
               '</�����ID><����>' || ����_In || '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_018', v_Value);
  End;

  --37.�����˿�
  Procedure Zlhis_Patient_019
  (
    �䶯id_In   In ���˱䶯��¼.Id%Type,
    ����id_In   In ������Ϣ.����id%Type,
    �����id_In In ҽ�ƿ����.Id%Type,
    ����_In     In ����ҽ�ƿ���Ϣ.����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�䶯ID>' || �䶯id_In || '</�䶯ID><����ID>' || ����id_In || '</����ID><�����ID>' || �����id_In ||
               '</�����ID><����>' || ����_In || '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_019', v_Value);
  End;

  --38.���߲���/����
  Procedure Zlhis_Patient_020
  (
    �䶯id_In   In ���˱䶯��¼.Id%Type,
    ����id_In   In ������Ϣ.����id%Type,
    �����id_In In ҽ�ƿ����.Id%Type,
    ԭ����_In   In ����ҽ�ƿ���Ϣ.����%Type,
    �¿���_In   In ����ҽ�ƿ���Ϣ.����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�䶯ID>' || �䶯id_In || '</�䶯ID><����ID>' || ����id_In || '</����ID><�����ID>' || �����id_In ||
               '</�����ID><ԭ����>' || ԭ����_In || '</ԭ����><�¿���>' || �¿���_In || '</�¿���></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_020', v_Value);
  End;

  --39.���˹ҺŵǼǣ�����ԤԼ�Ǽ�)
  Procedure Zlhis_Regist_001
  (
    �Һ�id_In In ���˹Һż�¼.Id%Type,
    No_In     In ���˹Һż�¼.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�Һ�ID>' || �Һ�id_In || '</�Һ�ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_001', v_Value);
  End;

  --40.���˷���
  Procedure Zlhis_Regist_002
  (
    �Һ�id_In In ���˹Һż�¼.Id%Type,
    No_In     In ���˹Һż�¼.No%Type,
    ����_In   In ���˹Һż�¼.����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�Һ�ID>' || �Һ�id_In || '</�Һ�ID><NO>' || No_In || '</NO><����>' || Nvl(����_In, '') || '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_002', v_Value);
  End;

  --41.�����˺ţ���ȡ��ԤԼ)
  Procedure Zlhis_Regist_003
  (
    �Һ�id_In In ���˹Һż�¼.Id%Type,
    No_In     In ���˹Һż�¼.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�Һ�ID>' || �Һ�id_In || '</�Һ�ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_003', v_Value);
  End;

  --42.�ٴ����ﰲ�ŵ���
  Procedure Zlhis_Regist_004
  (
    �䶯ԭ��_In In Integer, --1-ͣ��;2-����;3-���ұ䶯
    ��¼id_In   In �ٴ������¼.Id%Type,
    �䶯id_In   In �ٴ�����䶯��¼.Id%Type
    
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�䶯ԭ��>' || �䶯ԭ��_In || '</�䶯ԭ��><��¼ID>' || ��¼id_In || '</��¼ID><�䶯ID>' || �䶯id_In ||
               '</�䶯ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_004', v_Value);
  End;

  --43.���ﻼ�߹ҺŻ��Ų���
  Procedure Zlhis_Regist_005
  (
    No_In         In ���˹Һż�¼.No%Type,
    �䶯ԭ��_In   Integer, --1-����;2-����;3-ԤԼ���ڱ䶯,
    ����䶯id_In ����䶯��¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><NO>' || No_In || '</NO><�䶯ԭ��>' || �䶯ԭ��_In || '</�䶯ԭ��><����䶯ID>' || ����䶯id_In ||
               '</����䶯ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_005', v_Value);
  End;

  --���������շѼ��������
  Procedure Zlhis_Charge_002
  (
    ��������_In In Number,
    ����id_In   In ������ü�¼.����id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    --��������_In:1-�շѽ��㣬2-�������
    v_Value := '<root><��������>' || ��������_In || '</��������><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_002', v_Value);
  End;

  --46.�����˷ѵ���
  Procedure Zlhis_Charge_004
  (
    �˷�����_In In Number,
    ����id_In   In ������ü�¼.����id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    --�˷�����_In:1-�շѽ��㣬2-�������
    v_Value := '<root><�˷�����>' || �˷�����_In || '</�˷�����><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_004', v_Value);
  End;

  --47.��Ԥ����
  Procedure Zlhis_Charge_005
  (
    Ԥ��id_In In ����Ԥ����¼.Id%Type,
    ���ݺ�_In In ����Ԥ����¼.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><Ԥ��ID>' || Ԥ��id_In || '</Ԥ��ID><���ݺ�>' || ���ݺ�_In || '</���ݺ�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_005', v_Value);
  End;

  --48.��Ԥ����(����������Ԥ�����)
  Procedure Zlhis_Charge_006
  (
    ��Ԥ��id_In In ����Ԥ����¼.Id%Type,
    ���ݺ�_In   In ����Ԥ����¼.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><��Ԥ��ID>' || ��Ԥ��id_In || '</��Ԥ��ID><���ݺ�>' || ���ݺ�_In || '</���ݺ�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_006', v_Value);
  End;

  --סԺ���ʵ���
  Procedure Zlhis_Charge_007
  (
    �շ����_In In סԺ���ü�¼.�շ����%Type,
    ����id_In   In סԺ���ü�¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�շ����>' || �շ����_In || '</�շ����><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_007', v_Value);
  End;

  --סԺ���ʵ�������
  Procedure Zlhis_Charge_008
  (
    �շ����_In In סԺ���ü�¼.�շ����%Type,
    ����id_In   In סԺ���ü�¼.Id%Type,
    �շ�ids_In  In Varchar2 := Null --���ܷ���ID��Ӧ����շ�id����Ӧ��ʽ���շ�id,����|�շ�id,��������ҩƷ����
  ) Is
    v_Value   Zlmsg_Todo.Key_Value%Type;
    v_Tmp     Varchar2(4000);
    v_Infotmp Varchar2(4000);
    v_Fields  Varchar2(4000);
    v_�շ�id  Varchar2(50);
    v_����    Varchar2(20);
  Begin
    If p_Msg_Using('ZLHIS_CHARGE_008') = 0 Then
      Return;
    End If;
    v_Value := '<root><�շ����>' || �շ����_In || '</�շ����><����ID>' || ����id_In || '</����ID>';
  
    If �շ�ids_In Is Null Then
      v_Infotmp := Null;
      v_Tmp     := '<�շ�IDS>' || '<�շ�ID>' || '</�շ�ID>' || '<����>' || '</����>' || '</�շ�IDS>';
    Else
      v_Infotmp := �շ�ids_In || '|';
      While v_Infotmp Is Not Null Loop
        --�ֽ��շ�ID��
        v_Fields  := Substr(v_Infotmp, 1, Instr(v_Infotmp, '|') - 1);
        v_�շ�id  := Substr(v_Fields, 1, Instr(v_Fields, ',') - 1);
        v_����    := Substr(v_Fields, Instr(v_Fields, ',') + 1);
        v_Infotmp := Replace('|' || v_Infotmp, '|' || v_Fields || '|');
      
        v_Tmp := v_Tmp || '<�շ�IDS>' || '<�շ�ID>' || v_�շ�id || '</�շ�ID>' || '<����>' || v_���� || '</����>' || '</�շ�IDS>';
      End Loop;
    End If;
  
    v_Value := v_Value || v_Tmp || '</root>';
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_008', v_Value);
  End;

  --53.סԺ������Ժ�Ǽ�
  Procedure Zlhis_Patient_001
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
    n_�䶯id ���˱䶯��¼.Id%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_001') = 0 Then
      Return;
    End If;
    Select Max(ID)
    Into n_�䶯id
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ʼԭ�� = 1 And Nvl(���Ӵ�λ, 0) = 0;
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_001',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�䶯ID>' || n_�䶯id ||
                                 '</�䶯ID></root>');
  End Zlhis_Patient_001;
  --54.סԺ������Ժ���
  Procedure Zlhis_Patient_002
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
    n_�䶯id ���˱䶯��¼.Id%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_002') = 0 Then
      Return;
    End If;
    Select Max(ID)
    Into n_�䶯id
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0;
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_002',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�䶯ID>' || n_�䶯id ||
                                 '</�䶯ID></root>');
  End Zlhis_Patient_002;
  --56.סԺ���ߴ�λ���
  Procedure Zlhis_Patient_004
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
    v_ԭ����   Varchar2(255);
    v_�´���   Varchar2(255);
    n_�䶯id   Number(18);
    n_��ʼԭ�� Number(3);
    d_��ʼʱ�� Date;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_004') = 0 Then
      Return;
    End If;
    Select ID, ����, ��ʼʱ��, ��ʼԭ��
    Into n_�䶯id, v_�´���, d_��ʼʱ��, n_��ʼԭ��
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0;
  
    Select Max(����)
    Into v_ԭ����
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� = d_��ʼʱ�� And ��ֹԭ�� = n_��ʼԭ�� And Nvl(���Ӵ�λ, 0) = 0;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_004',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID>' || '<ԭ����>' ||
                                 v_ԭ���� || '</ԭ����>' || '<�´���>' || v_�´��� || '</�´���>' || '<�䶯ID>' || n_�䶯id || '</�䶯ID>' ||
                                 '</root>');
  End Zlhis_Patient_004;
  --57.סԺ���߲�����
  Procedure Zlhis_Patient_005
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
    n_�䶯id ���˱䶯��¼.Id%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_005') = 0 Then
      Return;
    End If;
    Select Max(ID)
    Into n_�䶯id
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0;
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_005',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�䶯ID>' || n_�䶯id ||
                                 '</�䶯ID></root>');
  End Zlhis_Patient_005;
  --58.סԺ���߱������
  Procedure Zlhis_Patient_006
  (
    ����id_In   In ������ҳ.����id%Type,
    ��ҳid_In   In ������ҳ.��ҳid%Type,
    ������ʽ_In In Varchar2
  ) Is
    n_����id     ���˱䶯��¼.����id%Type;
    n_����id     ���˱䶯��¼.����id%Type;
    n_����ȼ�id ���˱䶯��¼.����ȼ�id%Type;
    n_ҽ��С��id ���˱䶯��¼.ҽ��С��id%Type;
    v_����       ���˱䶯��¼.����%Type;
    v_���λ�ʿ   ���˱䶯��¼.���λ�ʿ%Type;
    v_����ҽʦ   ���˱䶯��¼.����ҽʦ%Type;
    v_����ҽʦ   ���˱䶯��¼.����ҽʦ%Type;
    v_����ҽʦ   ���˱䶯��¼.����ҽʦ%Type;
    v_����       ���˱䶯��¼.����%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_006') = 0 Then
      Return;
    End If;
    Select Max(����id), Max(����id), Max(����ȼ�id), Max(ҽ��С��id), Max(����), Max(���λ�ʿ), Max(����ҽʦ), Max(����ҽʦ), Max(����ҽʦ), Max(����)
    Into n_����id, n_����id, n_����ȼ�id, n_ҽ��С��id, v_����, v_���λ�ʿ, v_����ҽʦ, v_����ҽʦ, v_����ҽʦ, v_����
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And (��ֹʱ�� Is Null Or ��ֹԭ�� = 1) And Nvl(���Ӵ�λ, 0) = 0;
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_006',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><������ʽ>' || ������ʽ_In ||
                                 '</������ʽ><����ID>' || n_����id || '</����ID>' || '<����ID>' || n_����id || '</����ID>' || '<����ȼ�ID>' ||
                                 n_����ȼ�id || '</����ȼ�ID>' || '<ҽ��С��ID>' || n_ҽ��С��id || '</ҽ��С��ID>' || '<����>' || v_���� ||
                                 '</����>' || '<���λ�ʿ>' || v_���λ�ʿ || '</���λ�ʿ>' || '<����ҽʦ>' || v_����ҽʦ || '</����ҽʦ>' ||
                                 '<����ҽʦ>' || v_����ҽʦ || '</����ҽʦ>' || '<����ҽʦ>' || v_����ҽʦ || '</����ҽʦ>' || '<����>' || v_���� ||
                                 '</����>' || '</root>');
  End Zlhis_Patient_006;
  --59.סԺ����ҽ�����
  Procedure Zlhis_Patient_007
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
    v_ԭסԺҽ�� Varchar2(100);
    v_��סԺҽ�� Varchar2(100);
    v_ԭ����ҽ�� Varchar2(100);
    v_������ҽ�� Varchar2(100);
    v_ԭ����ҽ�� Varchar2(100);
    v_������ҽ�� Varchar2(100);
    v_ԭ���λ�ʿ Varchar2(100);
    v_�����λ�ʿ Varchar2(100);
    n_�䶯id     Number(18);
    n_��ʼԭ��   Number(3);
    d_��ʼʱ��   Date;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_007') = 0 Then
      Return;
    End If;
    Select ID, ����ҽʦ, ����ҽʦ, ����ҽʦ, ���λ�ʿ, ��ʼʱ��, ��ʼԭ��
    Into n_�䶯id, v_��סԺҽ��, v_������ҽ��, v_������ҽ��, v_�����λ�ʿ, d_��ʼʱ��, n_��ʼԭ��
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0;
  
    Select Max(����ҽʦ), Max(����ҽʦ), Max(����ҽʦ), Max(���λ�ʿ)
    Into v_ԭסԺҽ��, v_ԭ����ҽ��, v_ԭ����ҽ��, v_ԭ���λ�ʿ
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� = d_��ʼʱ�� And ��ֹԭ�� = n_��ʼԭ�� And Nvl(���Ӵ�λ, 0) = 0;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_007',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID>' || '<ԭסԺҽ��>' ||
                                 v_ԭסԺҽ�� || '</ԭסԺҽ��>' || '<��סԺҽ��>' || v_��סԺҽ�� || '</��סԺҽ��>' || '<ԭ����ҽ��>' || v_ԭ����ҽ�� ||
                                 '</ԭ����ҽ��>' || '<������ҽ��>' || v_������ҽ�� || '</������ҽ��>' || '<ԭ����ҽ��>' || v_ԭ����ҽ�� || '</ԭ����ҽ��>' ||
                                 '<������ҽ��>' || v_������ҽ�� || '</������ҽ��>' || '<ԭ���λ�ʿ>' || v_ԭ���λ�ʿ || '</ԭ���λ�ʿ>' || '<�����λ�ʿ>' ||
                                 v_�����λ�ʿ || '</�����λ�ʿ>' || '<�䶯ID>' || n_�䶯id || '</�䶯ID>' || '</root>');
  End Zlhis_Patient_007;
  --סԺ���߻���ȼ����
  Procedure Zlhis_Patient_008
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
    v_ԭ����ȼ�id Number(18);
    v_�»���ȼ�id Number(18);
    n_�䶯id       Number(18);
    n_��ʼԭ��     Number(3);
    d_��ʼʱ��     Date;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_008') = 0 Then
      Return;
    End If;
    Select ID, ����ȼ�id, ��ʼʱ��, ��ʼԭ��
    Into n_�䶯id, v_�»���ȼ�id, d_��ʼʱ��, n_��ʼԭ��
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0;
  
    Select Max(����ȼ�id)
    Into v_ԭ����ȼ�id
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� = d_��ʼʱ�� And ��ֹԭ�� = n_��ʼԭ�� And Nvl(���Ӵ�λ, 0) = 0;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_008',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID>' || '<ԭ����ȼ�ID>' ||
                                 v_ԭ����ȼ�id || '</ԭ����ȼ�ID>' || '<�»���ȼ�ID>' || v_�»���ȼ�id || '</�»���ȼ�ID>' || '<�䶯ID>' ||
                                 n_�䶯id || '</�䶯ID>' || '</root>');
  End Zlhis_Patient_008;
  --60.סԺ����Ԥ��Ժ
  Procedure Zlhis_Patient_009
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
    n_�䶯id ���˱䶯��¼.Id%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_009') = 0 Then
      Return;
    End If;
    Select Max(ID)
    Into n_�䶯id
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0;
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_009',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�䶯ID>' || n_�䶯id ||
                                 '</�䶯ID></root>');
  End Zlhis_Patient_009;
  --61.סԺ���߳�Ժ
  Procedure Zlhis_Patient_010
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_010',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID></root>');
  End Zlhis_Patient_010;
  --62.סԺ�����������Ǽ�
  Procedure Zlhis_Patient_011
  (
    ����id_In   In ������ҳ.����id%Type,
    ��ҳid_In   In ������ҳ.��ҳid%Type,
    Ӥ�����_In ����ҽ����¼.Ӥ��%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_011',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><Ӥ�����>' || Ӥ�����_In ||
                                 '</Ӥ�����></root>');
  End Zlhis_Patient_011;
  --63.סԺ����ת�����
  Procedure Zlhis_Patient_012
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
    v_ת������id Number(18);
    v_ת�����id Number(18);
    n_�䶯id     Number(18);
    n_��ʼԭ��   Number(3);
    d_��ʼʱ��   Date;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_012') = 0 Then
      Return;
    End If;
    Select ID, ����id, ��ʼʱ��, ��ʼԭ��
    Into n_�䶯id, v_ת�����id, d_��ʼʱ��, n_��ʼԭ��
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0;
  
    Select Max(����id)
    Into v_ת������id
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� = d_��ʼʱ�� And ��ֹԭ�� = n_��ʼԭ�� And Nvl(���Ӵ�λ, 0) = 0;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_012',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID>' || '<ת������ID>' ||
                                 v_ת������id || '</ת������ID>' || '<ת�����ID>' || v_ת�����id || '</ת�����ID>' || '<�䶯ID>' || n_�䶯id ||
                                 '</�䶯ID>' || '</root>');
  End Zlhis_Patient_012;
  --64.�������Ǽ�����
  Procedure Zlhis_Patient_013
  (
    ����id_In   In ������ҳ.����id%Type,
    ��ҳid_In   In ������ҳ.��ҳid%Type,
    Ӥ�����_In ����ҽ����¼.Ӥ��%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_013',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><Ӥ�����>' || Ӥ�����_In ||
                                 '</Ӥ�����></root>');
  End Zlhis_Patient_013;
  --65.���ﻼ�ߵǼ�
  Procedure Zlhis_Patient_015(����id_In In ������ҳ.����id%Type) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_015', '<root><����ID>' || ����id_In || '</����ID></root>');
  End Zlhis_Patient_015;
  --66.������Ϣ�޸�
  Procedure Zlhis_Patient_016(����id_In In ������ҳ.����id%Type) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_016', '<root><����ID>' || ����id_In || '</����ID></root>');
  End Zlhis_Patient_016;

  --67.���ߺϲ�
  Procedure Zlhis_Patient_017 
  ( 
    ����id_In   In ������ҳ.����id%Type, 
    ԭ����id_In In ������ҳ.����id%Type,
    �仯ids_In  In Varchar2
  ) Is 
  --������ 1����id,1��ҳid:1ԭ����id,1ԭ��ҳid; 2����id,2��ҳid:2ԭ����id,2ԭ��ҳid;��.
  Begin 
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_017', 
                                '<root><����ID>' || ����id_In || '</����ID><ԭ����ID>' || ԭ����id_In || '</ԭ����ID><CINFO>'||�仯ids_In||'</CINFO></root>'); 
  End Zlhis_Patient_017;

  --69.סԺ����ת�벡��
  Procedure Zlhis_Patient_026
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
    v_ת������id Number(18);
    v_ת�벡��id Number(18);
    n_�䶯id     Number(18);
    n_��ʼԭ��   Number(3);
    d_��ʼʱ��   Date;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_026') = 0 Then
      Return;
    End If;
    Select ID, ����id, ��ʼʱ��, ��ʼԭ��
    Into n_�䶯id, v_ת�벡��id, d_��ʼʱ��, n_��ʼԭ��
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0;
  
    Select Max(����id)
    Into v_ת������id
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� = d_��ʼʱ�� And ��ֹԭ�� = n_��ʼԭ�� And Nvl(���Ӵ�λ, 0) = 0;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_026',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID>' || '<ת������ID>' ||
                                 v_ת������id || '</ת������ID>' || '<ת�벡��ID>' || v_ת�벡��id || '</ת�벡��ID>' || '<�䶯ID>' || n_�䶯id ||
                                 '</�䶯ID>' || '</root>');
  End Zlhis_Patient_026;

  Procedure Zlhis_Patient_028(����id_In In ������ҳ.����id%Type) Is 
    v_����     ������Ϣ.����%Type; 
    v_�Ա�     ������Ϣ.�Ա�%Type; 
    v_����     ������Ϣ.����%Type; 
    v_�����   ������Ϣ.�����%Type; 
    v_���֤�� ������Ϣ.���֤��%Type; 
    v_�������� varchar2(50); 
  Begin 
    If p_Msg_Using('ZLHIS_PATIENT_028') = 0 Then 
      Return; 
    End If; 
    Select ����, �Ա�, ����, To_Char(��������, 'yyyymmdd'), �����, ���֤�� 
    Into v_����, v_�Ա�, v_����, v_��������, v_�����, v_���֤�� 
    From ������Ϣ 
    Where ����id = ����id_In; 
 
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_028', 
                                '<root><����ID>' || ����id_In || '</����ID><����>' || v_���� || '</����>' || '<�Ա�>' || v_�Ա� || 
                                 '</�Ա�>' || '<����>' || v_���� || '</����>' || '<��������>' || v_�������� || '</��������>' || '<�����>' || 
                                 v_����� || '</�����>' || '<���֤��>' || v_���֤�� || '</���֤��>' || '</root>'); 
  End Zlhis_Patient_028; 

  --79.���۲���תסԺ����
  Procedure Zlhis_Patient_029
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
    n_�䶯id Number(18);
  Begin
    Select max(ID)
    Into n_�䶯id
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0 And ��ʼԭ�� = 9;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_029',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�䶯ID>' || n_�䶯id ||
                                 '</�䶯ID></root>');
  
  End Zlhis_Patient_029;

  --Ѫ��:������Ѫ���
  Procedure Zlhis_Blood_001(ҽ��id_In In ����ҽ����¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If ҽ��id_In Is Not Null Then
      v_Value := '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_BLOOD_001', v_Value);
    End If;
  End Zlhis_Blood_001;

  --Ѫ��:���Ҿܾ���Ѫ
  Procedure Zlhis_Blood_002(ҽ��id_In In ����ҽ����¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If ҽ��id_In Is Not Null Then
      v_Value := '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_BLOOD_002', v_Value);
    End If;
  End Zlhis_Blood_002;

  --70.���鱨�����
  Procedure Zlhis_Lis_001(�걾id_In In ����걾��¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If �걾id_In Is Not Null Then
      v_Value := '<root><�걾ID>' || �걾id_In || '</�걾ID><ϵͳ>1</ϵͳ></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_001', v_Value);
    End If;
  End Zlhis_Lis_001;
  --71.���鱨����˳���
  Procedure Zlhis_Lis_002(�걾id_In In ����걾��¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If �걾id_In Is Not Null Then
      v_Value := '<root><�걾ID>' || �걾id_In || '</�걾ID><ϵͳ>1</ϵͳ></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_002', v_Value);
    End If;
  End Zlhis_Lis_002;
  --73.����걾�����ӡ
  Procedure Zlhis_Lis_004
  (
    ��������_In In ����ҽ������.��������%Type,
    ҽ��id_In   In ����ҽ������.ҽ��id%Type,
    ҽ��ids_In  In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If p_Msg_Using('ZLHIS_LIS_004') = 0 Then
      Return;
    End If;
    If ҽ��id_In Is Not Null Then
      v_Value := '<root><��������>' || ��������_In || '</��������><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ϵͳ>1</ϵͳ></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_004', v_Value);
    Else
      For R In (Select '<root><��������>' || ��������_In || '</��������><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ϵͳ>1</ϵͳ></root>' As Xml_Value
                From ����ҽ������
                Where ҽ��id In (Select Column_Value From Table(f_Num2list(ҽ��ids_In)))) Loop
        b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_004', r.Xml_Value);
      End Loop;
    End If;
  End Zlhis_Lis_004;
  --74.����걾�����ӡ����
  Procedure Zlhis_Lis_005
  (
    ��������_In In ����ҽ������.��������%Type,
    ҽ��id_In   In ����ҽ������.ҽ��id%Type,
    ҽ��ids_In  In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If p_Msg_Using('ZLHIS_LIS_005') = 0 Then
      Return;
    End If;
    If ҽ��id_In Is Not Null Then
      v_Value := '<root><��������>' || ��������_In || '</��������><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ϵͳ>1</ϵͳ></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_005', v_Value);
    Else
      For R In (Select '<root><��������>' || ��������_In || '</��������><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ϵͳ>1</ϵͳ></root>' As Xml_Value
                From ����ҽ������
                Where ҽ��id In (Select Column_Value From Table(f_Num2list(ҽ��ids_In)))) Loop
        b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_005', r.Xml_Value);
      End Loop;
    End If;
  End Zlhis_Lis_005;
  --75.����걾����
  Procedure Zlhis_Lis_006(�걾id_In In ����걾��¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If �걾id_In Is Not Null Then
      v_Value := '<root><�걾ID>' || �걾id_In || '</�걾ID><ϵͳ>1</ϵͳ></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_006', v_Value);
    End If;
  End Zlhis_Lis_006;
  --76.����걾���ճ���
  Procedure Zlhis_Lis_007(�걾id_In In ����걾��¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If �걾id_In Is Not Null Then
      v_Value := '<root><�걾ID>' || �걾id_In || '</�걾ID><ϵͳ>1</ϵͳ></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_007', v_Value);
    End If;
  End Zlhis_Lis_007;
  --77.����걾����
  Procedure Zlhis_Lis_008(�걾id_In In ����걾��¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If �걾id_In Is Not Null Then
      v_Value := '<root><�걾ID>' || �걾id_In || '</�걾ID><ϵͳ>1</ϵͳ></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_008', v_Value);
    End If;
  End Zlhis_Lis_008;

End b_Message;
/



--138877:��͢��,2019-04-23,�������۲���תסԺ������Ϣ
Create Or Replace Procedure Zl_���˱䶯��¼_תסԺ
(
  ����id_In ������ҳ.����id%Type,
  ��ҳid_In ������ҳ.��ҳid%Type,
  סԺ��_In ������Ϣ.סԺ��%Type
) Is
  --���ܣ���סԺ���۲���תΪסԺ����
  v_Count      Number;
  v_��Ժ����id Number;
  v_Date       Date;
  v_Temp       Varchar2(255);
  v_��Ա���   סԺ���ü�¼.����Ա���%Type;
  v_��Ա����   סԺ���ü�¼.����Ա����%Type;
  Cursor c_Oldinfo Is
    Select b.*
    From (Select c.*
           From ���˱䶯��¼ C
           Where c.����id = ����id_In And c.��ҳid = ��ҳid_In And
                 c.��ʼʱ�� = (Select Min(��ʼʱ��)
                           From ���˱䶯��¼
                           Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ʼʱ�� > v_Date)) A, ���˱䶯��¼ B
    
    Where b.����id = ����id_In And b.��ҳid = ��ҳid_In And a.��ʼʱ�� = b.��ֹʱ�� And a.��ʼԭ�� = b.��ֹԭ�� And a.���Ӵ�λ = b.���Ӵ�λ
    Union
    Select *
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null And ��ʼʱ�� <= v_Date;

  Cursor c_Endinfo Is
    Select * From ���˱䶯��¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null;
  r_Oldinfo c_Oldinfo%RowType;
  r_Endinfo c_Endinfo%RowType;

  v_��ֹԭ�� ���˱䶯��¼.��ֹԭ��%Type;
  v_��ֹʱ�� ���˱䶯��¼.��ֹʱ��%Type;
  v_��ֹ��Ա ���˱䶯��¼.��ֹ��Ա%Type;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --�����������
  Select Nvl(״̬, 0), ��Ժ����id
  Into v_Count, v_��Ժ����id
  From ������ҳ
  Where ����id = ����id_In And ��ҳid = ��ҳid_In And �������� = 2;
  If v_Count = 1 Then
    v_Error := '���˵�ǰ��δ���,����תΪסԺ���ˡ����Ƚ�������ƺ����ԡ�';
    Raise Err_Custom;
  Elsif v_Count = 2 Then
    v_Error := '���˵�ǰ����ת��,����תΪסԺ���ˡ����Ƚ�����ת�ƻ�ȡ��ת�ƺ����ԡ�';
    Raise Err_Custom;
  End If;

  Select Zl_סԺ�ձ�_Count(v_��Ժ����id, Trunc(Sysdate)) Into v_Count From Dual;
  If v_Count > 0 Then
    v_Error := '�Ѳ���ҵ��ʱ���ڵ�סԺ�ձ�,���ܰ����ҵ��!';
    Raise Err_Custom;
  End If;

  Select Sysdate Into v_Date From Dual;
  v_Temp     := Zl_Identity;
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_��Ա��� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);

  Open c_Oldinfo; --�����ȴ�
  Fetch c_Oldinfo
    Into r_Oldinfo;
  Open c_Endinfo;
  Fetch c_Endinfo
    Into r_Endinfo;
  If c_Endinfo%RowCount = 0 Then
    Close c_Endinfo;
    v_Error := 'δ���ָò��˵�ǰ��Ч�ı䶯��¼��';
    Raise Err_Custom;
  End If;
  Select Count(*)
  Into v_Count
  From ���˱䶯��¼
  Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ʼʱ�� Is Null And ��ֹʱ�� Is Null;
  If v_Count > 0 Then
    v_Error := '�ò�������ת�ƻ�ת���������ܽ��������䶯��';
    Raise Err_Custom;
  End If;

  --ȡ���ϴα䶯
  If r_Oldinfo.��ֹʱ�� Is Not Null Then
    v_��ֹʱ�� := r_Oldinfo.��ֹʱ��;
    v_��ֹԭ�� := r_Oldinfo.��ֹԭ��;
    v_��ֹ��Ա := r_Oldinfo.��ֹ��Ա;
    --ȡ���ϴα䶯
    Update ���˱䶯��¼
    Set ��ֹʱ�� = v_Date, ��ֹԭ�� = 9, ��ֹ��Ա = v_��Ա����, �ϴμ���ʱ�� = Null
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� = v_��ֹʱ�� And ��ֹԭ�� = v_��ֹԭ��;
    --���½����ļ�¼�����ֹͣ����������ɾ���ϴμ���ʱ��
    Update ���˱䶯��¼ Set �ϴμ���ʱ�� = Null Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ʼʱ�� > v_Date;
  Else
    Update ���˱䶯��¼
    Set ��ֹʱ�� = v_Date, ��ֹԭ�� = 9, ��ֹ��Ա = v_��Ա����
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null;
  End If;

  --�����䶯��¼
  While c_Oldinfo%Found Loop
    Insert Into ���˱䶯��¼
      (ID, ����id, ��ҳid, ��ʼʱ��, ��ʼԭ��, ���Ӵ�λ, ����id, ����id, ҽ��С��id, ����ȼ�id, ��λ�ȼ�id, ����, ���λ�ʿ, ����ҽʦ, ����ҽʦ, ����ҽʦ, ����, ����Ա���,
       ����Ա����, ��ֹʱ��, ��ֹԭ��, ��ֹ��Ա)
    Values
      (���˱䶯��¼_Id.Nextval, ����id_In, ��ҳid_In, v_Date, 9, r_Oldinfo.���Ӵ�λ, r_Oldinfo.����id, r_Oldinfo.����id, r_Oldinfo.ҽ��С��id,
       r_Oldinfo.����ȼ�id, r_Oldinfo.��λ�ȼ�id, r_Oldinfo.����, r_Oldinfo.���λ�ʿ, r_Oldinfo.����ҽʦ, r_Oldinfo.����ҽʦ, r_Oldinfo.����ҽʦ,
       r_Oldinfo.����, v_��Ա���, v_��Ա����, v_��ֹʱ��, v_��ֹԭ��, v_��ֹ��Ա);
    If Nvl(r_Oldinfo.���Ӵ�λ, 0) = 0 Then
      --����������дʱ��
      Zl_���Ӳ���ʱ��_Insert(����id_In, ��ҳid_In, 2, '��Ժ', r_Oldinfo.����id, r_Oldinfo.����ҽʦ, v_Date, v_Date);
    End If;
    Fetch c_Oldinfo
      Into r_Oldinfo;
  End Loop;

  Close c_Oldinfo;
  Close c_Endinfo;

  Update ������ҳ Set �������� = 0, סԺ�� = סԺ��_In Where ����id = ����id_In And ��ҳid = ��ҳid_In;
  Update ������Ϣ Set סԺ�� = סԺ��_In, סԺ���� = Nvl(סԺ����, 0) + 1 Where ����id = ����id_In;

  --�����������
  Select Count(*)
  Into v_Count
  From ���˱䶯��¼
  Where ����id = ����id_In And ��ҳid = ��ҳid_In And Nvl(���Ӵ�λ, 0) = 0 And ��ʼʱ�� Is Not Null And ��ֹʱ�� Is Null;

  If v_Count > 1 Then
    v_Error := '���ֲ��˴��ڷǷ��ı䶯��¼,��ǰ�������ܼ�����' || Chr(13) || Chr(10) || '��������������粢�����������,��ˢ�²���״̬�����ԣ�';
    Raise Err_Custom;
  End If;

 --����תסԺ��Ϣ����
  b_Message.Zlhis_Patient_029(����id_In, ��ҳid_In);
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˱䶯��¼_תסԺ;
/

--139595:��ҵ��,2019-04-15,����֧�������������
Create Or Replace Procedure Zl_��Һ��ҩ��¼_����
(
  ��ҩid_In   In Varchar2, --ID��:ID1,ID2....
  ������Ա_In In ��Һ��ҩ��¼.������Ա%Type := Null,
  ����ʱ��_In In ��Һ��ҩ��¼.����ʱ��%Type := Null,
  ����˵��_In In ��Һ��ҩ״̬.����˵��%Type := Null
) Is
  v_Tansid     Varchar2(20);
  v_Tmp        Varchar2(4000);
  v_Error      Varchar2(255);
  n_����״̬   ��Һ��ҩ��¼.����״̬%Type;
  n_People     Number(2);
  v_No         Varchar2(20);
  n_��Ŀid     Number(18);
  v_�շ���Ŀid Varchar2(200);
  n_Row        Number(2);
  n_Out        Number(10);
  n_Outnum     Number(10);
  n_Count      Number(18);
  n_Packet     Number(2);
  v_Usercode   Varchar2(100);
  v_������     Varchar2(20);
  Err_Custom Exception;

  Cursor c_Bill Is
    Select a.����id, a.��ҳid, a.��ʶ��, a.����, a.�Ա�, a.����, a.����, a.�ѱ�, a.���˲���id, a.���˿���id, a.Ӥ����, e.ҩƷid, b.�ⷿid, f.��ҩ����,
           c.ִ��ʱ��, g.���, 1 As �����־
    From סԺ���ü�¼ A, ҩƷ�շ���¼ B, ��Һ��ҩ��¼ C, ��Һ��ҩ���� D, ҩƷ��� E, ��ҺҩƷ���� F, �����շѷ��� G
    Where a.Id = b.����id And b.Id = d.�շ�id And d.��¼id = c.Id And b.ҩƷid = e.ҩƷid And b.ҩƷid = f.ҩƷid And
          Substr(f.��ҩ����, Instr(f.��ҩ����, '-') + 1) = g.��ҩ����(+) And Nvl(c.�Ƿ���, 0) <> 0 And c.Id = v_Tansid
    Union All
    Select a.����id, a.��ҳid, a.��ʶ��, a.����, a.�Ա�, a.����, '' As ����, a.�ѱ�, a.���˲���id, a.���˿���id, a.Ӥ����, e.ҩƷid, b.�ⷿid, f.��ҩ����,
           c.ִ��ʱ��, g.���, 0 As �����־
    From ������ü�¼ A, ҩƷ�շ���¼ B, ��Һ��ҩ��¼ C, ��Һ��ҩ���� D, ҩƷ��� E, ��ҺҩƷ���� F, �����շѷ��� G
    Where a.Id = b.����id And b.Id = d.�շ�id And d.��¼id = c.Id And b.ҩƷid = e.ҩƷid And b.ҩƷid = f.ҩƷid And
          Substr(f.��ҩ����, Instr(f.��ҩ����, '-') + 1) = g.��ҩ����(+) And Nvl(c.�Ƿ���, 0) <> 0 And c.Id = v_Tansid
    Order By ���;
Begin
  v_Usercode := Zl_Identity;
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ';') + 1);
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ',') + 1);
  v_Usercode := Substr(v_Usercode, 1, Instr(v_Usercode, ',') - 1);
  n_People   := Nvl(zl_GetSysParameter('���÷Ѱ�������ȡ', 1345), 0);
  n_Out      := Nvl(zl_GetSysParameter('��Ժ���˲������÷�', 1345), 0);
  n_Packet   := Nvl(zl_GetSysParameter('���ҩƷ�ڷ��ͻ�����ȡ���÷�', 1345), 0);

  If ��ҩid_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := ��ҩid_In || ',';
  End If;
  While v_Tmp Is Not Null Loop
    --�ֽⵥ��ID��
    v_Tansid := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp    := Replace(',' || v_Tmp, ',' || v_Tansid || ',');
  
    Begin
      Select ����״̬ Into n_����״̬ From ��Һ��ҩ��¼ Where ID = v_Tansid;
    
      If n_����״̬ > 4 Then
        v_Error := '�������ѱ����������ܽ��з��Ͳ�����';
        Raise Err_Custom;
      End If;
    Exception
      When Others Then
        v_Error := '�������ѱ�������';
        Raise Err_Custom;
    End;
  
    v_������ := '';
    Begin
      Select ������
      Into v_������
      From (Select e.������
             From ��Һ��ҩ��¼ A, ��Һ��ҩ���� B, ҩƷ�շ���¼ C, ��Һ̨ҩƷ���� D, ��Һ�������� E
             Where a.Id = b.��¼id And b.�շ�id = c.Id And c.ҩƷid = d.ҩƷid And c.�ⷿid = d.����id And d.��ҩ̨id = e.��ҩ̨id And
                   a.��ҩ���� = e.���� And e.���� = To_Date(To_Char(Sysdate, 'yyyy-mm-dd'), 'yyyy-mm-dd') And a.Id = v_Tansid And
                   Rownum = 1
             Order By d.��ҩ̨id)
      Where Rownum = 1;
    Exception
      When Others Then
        Null;
    End;
  
    Insert Into ��Һ��ҩ״̬
      (��ҩid, ��������, ������Ա, ����ʱ��, ����˵��, ʵ�ʹ�����Ա)
    Values
      (v_Tansid, 5, ������Ա_In, ����ʱ��_In, ����˵��_In, v_������);
    Update ��Һ��ҩ��¼ Set ����״̬ = 5, ������Ա = ������Ա_In, ����ʱ�� = ����ʱ��_In Where ID = v_Tansid;
  
    --���ҩƷ�շ�
    If n_Packet = 1 Then
      n_Count := 0;
      Select Nextno(14) Into v_No From Dual;
    
      For r_Bill In c_Bill Loop
        Select Count(����id)
        Into n_Outnum
        From ������ҳ
        Where ��ҳid = r_Bill.��ҳid And ����id = r_Bill.����id And (Nvl(״̬, 0) = 3 Or ��Ժ���� Is Not Null);
      
        --�Ȳ�ѯ�Ƿ��а���ҩ;����ȡ�����÷ѷ���
        Select Nvl(Max(��Ŀid), 0)
        Into n_��Ŀid
        From ��Һ��ҩ��¼ A, ����ҽ����¼ B, �����շѷ��� C
        Where a.Id = v_Tansid And a.ҽ��id = b.Id And b.������Ŀid = c.����id;
        If n_��Ŀid = 0 Then
          --���޶�Ӧ��ҩ;�������÷���ȡ���������ٲ�ѯ�Ƿ��а���ҩ������ȡ�����÷ѷ���
          Select Nvl(Max(��Ŀid), 0)
          Into n_��Ŀid
          From �����շѷ���
          Where ��ҩ���� = Substr(r_Bill.��ҩ����, Instr(r_Bill.��ҩ����, '-', 1, 1) + 1);
        End If;
      
        If n_��Ŀid <> 0 Then
          n_Row := 0;
        
          If n_People = 1 Then
            Select Count(��ҩid)
            Into n_Row
            From (Select ��ҩid
                   From ��Һ��ҩ���� A, סԺ���ü�¼ B, ��Һ��ҩ��¼ C
                   Where a.No = b.No And a.��ҩid = c.Id And b.����id = r_Bill.����id And b.��¼״̬ = 1 And b.�շ�ϸĿid = n_��Ŀid And
                         r_Bill.ִ��ʱ�� Between Trunc(c.ִ��ʱ��) And Trunc(c.ִ��ʱ�� + 1) - 1 / 24 / 60 / 60
                   Union All
                   Select ��ҩid
                   From ��Һ��ҩ���� A, ������ü�¼ B, ��Һ��ҩ��¼ C
                   Where a.No = b.No And a.��ҩid = c.Id And b.����id = r_Bill.����id And b.��¼״̬ = 1 And b.�շ�ϸĿid = n_��Ŀid And
                         r_Bill.ִ��ʱ�� Between Trunc(c.ִ��ʱ��) And Trunc(c.ִ��ʱ�� + 1) - 1 / 24 / 60 / 60);
          End If;
        Else
          n_Row := 1;
        End If;
      
        If n_Row = 0 And (n_Outnum = 0 Or n_Out = 0) Then
          For r_Item In (Select a.Id �շ�ϸĿid, a.��� �շ����, a.���㵥λ, a.�Ӱ�Ӽ� �Ӱ��־, d.Id ������Ŀid, d.�վݷ�Ŀ, b.�ּ�
                         From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ D
                         Where a.Id = b.�շ�ϸĿid And b.������Ŀid = d.Id And a.Id = n_��Ŀid And b.ִ������ <= Sysdate And
                               (b.��ֹ���� >= Sysdate Or b.��ֹ���� Is Null)) Loop
            If n_Count = 0 Then
              Insert Into ��Һ��ҩ���� (��ҩid, NO, ����id) Values (v_Tansid, v_No, r_Bill.����id);
            End If;
          
            n_Count := n_Count + 1;
          
            If r_Bill.�����־ = 1 Then
              Zl_סԺ���ʼ�¼_Insert(v_No, n_Count, r_Bill.����id, r_Bill.��ҳid, r_Bill.��ʶ��, r_Bill.����, r_Bill.�Ա�, r_Bill.����,
                               r_Bill.����, r_Bill.�ѱ�, r_Bill.���˲���id, r_Bill.���˿���id, r_Item.�Ӱ��־, r_Bill.Ӥ����, r_Bill.�ⷿid,
                               ������Ա_In, Null, r_Item.�շ�ϸĿid, r_Item.�շ����, r_Item.���㵥λ, Null, Null, Null, 1, 1, Null,
                               r_Bill.�ⷿid, Null, r_Item.������Ŀid, r_Item.�վݷ�Ŀ, r_Item.�ּ�, r_Item.�ּ�, r_Item.�ּ�, Null,
                               Sysdate, Sysdate, Null, Null, v_Usercode, ������Ա_In);
            Else
              Zl_������ʼ�¼_Insert(v_No, n_Count, r_Bill.����id, r_Bill.��ʶ��, r_Bill.����, r_Bill.�Ա�, r_Bill.����, r_Bill.�ѱ�,
                               r_Item.�Ӱ��־, r_Bill.Ӥ����, r_Bill.���˿���id, r_Bill.�ⷿid, ������Ա_In, Null, r_Item.�շ�ϸĿid,
                               r_Item.�շ����, r_Item.���㵥λ, 1, 1, Null, r_Bill.�ⷿid, Null, r_Item.������Ŀid, r_Item.�վݷ�Ŀ,
                               r_Item.�ּ�, r_Item.�ּ�, r_Item.�ּ�, Sysdate, Sysdate, Null, Null, v_Usercode, ������Ա_In);
            End If;
          End Loop;
        End If;
      
        If n_Row = 0 Then
          Exit;
        End If;
      End Loop;
    End If;
  End Loop;

  --��Ϣ����
  b_Message.Zlhis_Drug_008(��ҩid_In);
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Һ��ҩ��¼_����;
/

--139595:��ҵ��,2019-04-15,����֧�������������
Create Or Replace Procedure Zl_��Һ��ҩ��¼_�˲�
(
  ����id_In   In ��Һ��ҩ��¼.����id%Type,
  ҽ��id_In   In Varchar2, --��Һҽ����ҩ;����Ӧ��ҽ��ID:ҽ��ID1,ҽ��ID2...
  ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
  �˲���_In   In ��Һ��ҩ״̬.������Ա%Type,
  �˲�ʱ��_In In ��Һ��ҩ״̬.����ʱ��%Type
) Is
  v_Count    Number;
  v_���     Number;
  v_ִ��ʱ�� Date;

  v_���id      Number;
  v_New���id   Number;
  v_Old���id   Number;
  v_���ͺ�      Number;
  v_Tmp         Varchar2(200);
  I             Number;
  v_��ҩid      Number;
  v_����        Number;
  v_Maxno       Varchar2(4000);
  v_Lableno     Varchar2(200);
  v_Maxbatch    Number;
  v_Curdose     Number;
  v_Sumdose     Number;
  v_Drugcount   Number;
  v_Currdate    Date;
  n_Needcheck   Number;
  n_Lngid       ҩƷ�շ���¼.Id%Type;
  n_Count       Number(3);
  n_����        ҩƷ�շ���¼.����%Type;
  v_No          ҩƷ�շ���¼.No%Type;
  n_���ʹ���    Number(5);
  n_����id      ������Ϣ.����id%Type := 0;
  b_Change      Boolean := True;
  n_Sum         Number;
  n_��������    Number(1);
  n_Cur         Number(5);
  v_�ϴη��ͺ�  ����ҽ������.���ͺ�%Type;
  v_ҽ��ids     Varchar2(4000);
  v_Tansid      Varchar2(12);
  v_��ǰ����    Varchar2(20);
  n_Num         Number(8);
  d_Oldִ��ʱ�� Date;
  n_�Ƿ���    Number(1);
  n_���        Number(1);
  n_��ҩ��      Number(2);
  --���Ʋ���
  v_ҽ������       Number;
  v_��Һ����       Number;
  v_����Һ����     Varchar2(2000);
  v_����Һ��ҩ;�� Varchar2(2000);
  v_��Դ����       Varchar2(4000);
  v_Continue       Number := 1;
  v_Nodosage       Number := 0;
  v_�����ϴ�����   Number := 0;
  d_�ֹ����ʱ��   Date;
  n_Tpn���÷�ʽ    Number := 0;
  v_ҩƷ����       Varchar2(20);
  n_���ҩƷ����   Number(1);
  n_����ҩƷ����   Number(1);
  n_���ȼ�         Number := 999;
  n_�Զ�����       Number := 0;
  n_����id         Number := 0;
  n_Row            Number(2);
  n_��������       Number := 0;
  n_ʣ������       Number := 0;
  n_��������       Number := 0;
  n_�ۼ�����       Number := 0;
  n_ҽ��id         Number := 0;
  n_��д����       Number := 0;
  v_��ҩ����       Varchar2(20);
  v_ʱ�䴮         Varchar2(100);
  v_ʱ��ֵ         Date;
  v_Fields         Varchar2(100);
  v_�Ƿ�ı�       Varchar2(20);
  v_ʱ�䴮1        Varchar2(100);
  Err_Item Exception;
  n_��ͨ���С�� Number;

  Cursor c_ҽ����¼ Is
    Select /*+rule */
    Distinct e.ҽ��id As ���id, e.���ͺ�, b.Ƶ�ʼ��, b.�����λ, b.ִ��ʱ�䷽��, a.����, a.�Ա�, a.����, a.סԺ��, a.��ǰ���� As ����, a.��ǰ����id As ���˲���id,
             a.��ǰ����id As ���˿���id, e.�״�ʱ��, e.ĩ��ʱ��, b.��ʼִ��ʱ��, Nvl(e.��������, 0) As ����, e.����ʱ��, Decode(b.ҽ����Ч, 0, 1, 2) As ҽ������,
             b.������Ŀid As ��ҩ;��, b.����id, Nvl(c.ִ�б��, 0) As �Ƿ�tpn
    From ����ҽ������ E, ����ҽ����¼ B, ������Ϣ A, ������ĿĿ¼ C, Table(f_Num2list(ҽ��id_In)) D
    Where e.ҽ��id = b.Id And b.����id = a.����id And c.��� = 'E' And c.�������� = '2' And c.ִ�з��� = 1 And b.������Ŀid = c.Id And
          e.ҽ��id = d.Column_Value And e.���ͺ� = ���ͺ�_In
    Order By b.����id, e.ҽ��id, e.���ͺ�;

  Cursor c_����ҽ����¼ Is
    Select /*+rule */
    Distinct e.ҽ��id, e.���ͺ�, b.Ƶ�ʼ��, b.�����λ, b.ִ��ʱ�䷽��, a.����, a.�Ա�, a.����, a.סԺ��, a.��ǰ���� As ����, a.��ǰ����id As ���˲���id,
             a.��ǰ����id As ���˿���id, e.�״�ʱ��, e.ĩ��ʱ��, b.��ʼִ��ʱ��, Nvl(e.��������, 0) As ����, e.����ʱ��, Decode(b.ҽ����Ч, 0, 1, 2) As ҽ������,
             b.������Ŀid As ��ҩ;��, b.����id
    From ����ҽ������ E, ����ҽ����¼ B, ������Ϣ A, ������ĿĿ¼ C
    Where e.ҽ��id = b.Id And b.����id = a.����id And b.������Ŀid = c.Id And b.���id = v_���id And e.���ͺ� = ���ͺ�_In
    Order By e.ҽ��id, e.���ͺ�;

  Cursor c_�շ���¼ Is
    Select Distinct c.Id As �շ�id, c.���, c.ʵ������ As ����, Nvl(e.�Ƿ�������, 0) As �Ƿ�������, c.����, c.No
    From ����ҽ����¼ A, ����ҽ������ B, ҩƷ�շ���¼ C, סԺ���ü�¼ D, ��ҺҩƷ���� E
    Where a.Id = b.ҽ��id And c.����id = d.Id And a.Id = d.ҽ����� And b.No = c.No And c.ҩƷid = e.ҩƷid(+) And
          b.ִ�в���id + 0 = ����id_In And c.���� = 9 And c.������� Is Null And a.Id = n_ҽ��id And b.���ͺ� = ���ͺ�_In And c.��� < 1000
    Union All
    Select Distinct c.Id As �շ�id, c.���, c.ʵ������ As ����, Nvl(e.�Ƿ�������, 0) As �Ƿ�������, c.����, c.No
    From ����ҽ����¼ A, ����ҽ������ B, ҩƷ�շ���¼ C, ������ü�¼ D, ��ҺҩƷ���� E
    Where a.Id = b.ҽ��id And c.����id = d.Id And a.Id = d.ҽ����� And b.No = c.No And c.ҩƷid = e.ҩƷid(+) And
          b.ִ�в���id + 0 = ����id_In And c.���� = 9 And c.������� Is Null And a.Id = n_ҽ��id And b.���ͺ� = ���ͺ�_In And c.��� < 1000
    Order By NO, ���;

  Cursor c_ԭʼ�շ���¼ Is
    Select Distinct c.Id As �շ�id, c.���, c.ʵ������ As ����, Nvl(e.�Ƿ�������, 0) As �Ƿ�������, c.����, c.No
    From ����ҽ����¼ A, ����ҽ������ B, ҩƷ�շ���¼ C, סԺ���ü�¼ D, ��ҺҩƷ���� E
    Where a.Id = b.ҽ��id And c.����id = d.Id And a.Id = d.ҽ����� And b.No = c.No And c.ҩƷid = e.ҩƷid(+) And
          b.ִ�в���id + 0 = ����id_In And c.���� = 9 And c.������� Is Null And a.���id = v_���id And b.���ͺ� = ���ͺ�_In And c.��� < 1000
    Union All
    Select Distinct c.Id As �շ�id, c.���, c.ʵ������ As ����, Nvl(e.�Ƿ�������, 0) As �Ƿ�������, c.����, c.No
    From ����ҽ����¼ A, ����ҽ������ B, ҩƷ�շ���¼ C, ������ü�¼ D, ��ҺҩƷ���� E
    Where a.Id = b.ҽ��id And c.����id = d.Id And a.Id = d.ҽ����� And b.No = c.No And c.ҩƷid = e.ҩƷid(+) And
          b.ִ�в���id + 0 = ����id_In And c.���� = 9 And c.������� Is Null And a.���id = v_���id And b.���ͺ� = ���ͺ�_In And c.��� < 1000
    Order By NO, ���;

  Cursor c_��Һ����¼ Is
    Select a.Id, a.ִ��ʱ��, a.��ҩ����, a.ҽ��id, d.����ʱ��
    From ��Һ��ҩ��¼ A, ����ҽ����¼ B, ��ҩ�������� C, ����ҽ������ D
    Where a.ҽ��id = b.Id And a.��ҩ���� = c.���� And d.ҽ��id = a.ҽ��id And a.���ͺ� = d.���ͺ� And c.���� <> 0 And c.ҩƷ���� Is Null And
          b.����id = n_����id And a.����״̬ < 2 And a.ִ��ʱ�� Between Trunc(v_ʱ��ֵ) And Trunc(v_ʱ��ֵ + 1) - 1 / 24 / 60 / 60;

  v_��Һ����¼   c_��Һ����¼%RowType;
  v_ҽ����¼     c_ҽ����¼%RowType;
  v_�շ���¼     c_�շ���¼%RowType;
  v_����ҽ����¼ c_����ҽ����¼%RowType;

  Function Zl_Getpivaworkbatch
  (
    ִ��ʱ��_In In Date,
    ����ʱ��_In In Date,
    ҩƷ����_In In Varchar2 := Null
  ) Return Number As
  
    v_Exetime   Date;
    v_Starttime Date;
    v_Endtime   Date;
    v_Maxbatch  Number(2);
    v_Batch     Number;
    Cursor c_��ҩ���� Is
      Select ����, ��ҩʱ��, ��ҩʱ��, ���, ҩƷ����
      From ��ҩ��������
      Where ���� = 1 And ��������id = ����id_In
      Order By ҩƷ����, ����;
  
    v_��ҩ���� c_��ҩ����%RowType;
  Begin
    v_Exetime := To_Date(Substr(To_Char(ִ��ʱ��_In, 'yyyy-mm-dd hh24:mi'), 12), 'hh24:mi');
  
    Select Max(Nvl(����, 0)) + 1 Into v_Maxbatch From ��ҩ�������� Where ���� = 1 And ��������id = ����id_In;
  
    For v_��ҩ���� In c_��ҩ���� Loop
      v_Batch := 0;
    
      --���췢�͵�ҽ�����͵���������
      If (Trunc(ִ��ʱ��_In) >= Trunc(v_Currdate) And Trunc(����ʱ��_In) < Trunc(ִ��ʱ��_In)) Or n_�������� = 0 Then
        If v_��ҩ����.���� <> '0' And
           ((Nvl(v_��ҩ����.ҩƷ����, '0') <> '0' And v_��ҩ����.ҩƷ���� = ҩƷ����_In) Or Nvl(v_��ҩ����.ҩƷ����, '0') = '0') Then
          v_Starttime := To_Date(Substr(v_��ҩ����.��ҩʱ��, 1, Instr(v_��ҩ����.��ҩʱ��, '-') - 1), 'hh24:mi');
          v_Endtime   := To_Date(Substr(v_��ҩ����.��ҩʱ��, Instr(v_��ҩ����.��ҩʱ��, '-') + 1), 'hh24:mi');
        
          If v_Exetime >= v_Starttime And v_Exetime <= v_Endtime Then
            v_Batch := v_��ҩ����.����;
            n_���  := v_��ҩ����.���;
            Exit When v_Batch > 0;
          End If;
        End If;
      End If;
    End Loop;
  
    If v_Batch = 0 And (n_���ҩƷ���� <> 1 Or n_�������� = 1) Then
      v_Batch := v_Maxbatch;
    End If;
    Return(v_Batch);
  End;

  Function Zl_Getfirst
  (
    ��ҩid_In In Number,
    ����id_In In Number
  ) Return Number As
    n_First  Number;
    n_����id Number;
    Cursor c_���ȼ� Is
      Select ����id, ��ҩ����, ���ȼ�, Ƶ��
      From ��ҺҩƷ���ȼ�
      Where (����id = ����id_In Or ����id = 0)
      Order By ����id, ���ȼ� Desc;
  
    r_���ȼ� c_���ȼ�%RowType;
  Begin
    n_First := 0;
    For r_���ȼ� In c_���ȼ� Loop
      If n_����id <> 0 And r_���ȼ�.����id = 0 Then
        Exit;
      End If;
      n_����id := r_���ȼ�.����id;
    
      For r_��ҩ��¼ In (Select Distinct d.��ҩ����, e.ִ��Ƶ��
                     From ��Һ��ҩ��¼ A, ��Һ��ҩ���� B, ҩƷ�շ���¼ C, ��ҺҩƷ���� D, ����ҽ����¼ E
                     Where a.ҽ��id = e.Id And a.Id = b.��¼id And b.�շ�id = c.Id And c.ҩƷid = d.ҩƷid And a.Id = ��ҩid_In) Loop
        If Instr(r_��ҩ��¼.��ҩ����, r_���ȼ�.��ҩ����, 1) > 0 And (Instr(r_���ȼ�.Ƶ��, r_��ҩ��¼.ִ��Ƶ��, 1) > 0 Or r_���ȼ�.Ƶ�� = '����Ƶ��') Then
          n_First := r_���ȼ�.���ȼ�;
          Exit;
        End If;
      End Loop;
    End Loop;
  
    If n_First = 0 Then
      n_First := 999;
    End If;
    Return(n_First);
  End;
Begin
  n_Count          := 0;
  v_ҽ������       := Zl_To_Number(Nvl(zl_GetSysParameter('ҽ������', 1345), 1));
  v_��Һ����       := Zl_To_Number(Nvl(zl_GetSysParameter('ͬ������Һ����', 1345), 0));
  v_����Һ����     := Nvl(zl_GetSysParameter('����ҺҩƷ����', 1345), '');
  v_����Һ��ҩ;�� := Nvl(zl_GetSysParameter('��Һ��ҩ;��', 1345), '');
  v_��Դ����       := Nvl(zl_GetSysParameter('��Դ����', 1345), '');
  v_�����ϴ�����   := Zl_To_Number(Nvl(zl_GetSysParameter('�����ϴ�����', 1345), 0));
  n_Tpn���÷�ʽ    := Zl_To_Number(Nvl(zl_GetSysParameter('����Ӫ��ҩ�ﴦ�÷�ʽ', 1345), 0));
  n_���ҩƷ����   := Zl_To_Number(Nvl(zl_GetSysParameter('����ҩƷ����������ҩƷ�����ݸ�ҩʱ��û����ҩ���ε���Һ��Ĭ��Ϊ0���β����', 1345), 0));
  n_����ҩƷ����   := Zl_To_Number(Nvl(zl_GetSysParameter('����ҩƷ��ҩƷ����ָ������', 1345), 0));
  n_�Զ�����       := Zl_To_Number(Nvl(zl_GetSysParameter('�����Զ�����', 1345), 0));
  n_��������       := Zl_To_Number(Nvl(zl_GetSysParameter('���췢�͵�ҽ����������Һ��ȫ������������', 1345), 0));
  v_ҽ��ids        := ҽ��id_In;
  v_��ǰ����       := '';
  n_���ʹ���       := 0;

  --ȡ��ͨҵ�񾫶�λ��
  --���:1-ҩƷ 2-����
  --���ݣ�2-���ۼ� 4-���
  --��λ��ҩƷ:1-�ۼ� 5-��λ
  Select ���� Into n_��ͨ���С�� From ҩƷ���ľ��� Where ��� = 1 And ���� = 4 And ��λ = 5;

  Select Trunc(Sysdate) Into v_Currdate From Dual;

  Select Max(Nvl(����, 0)) + 1 Into v_Maxbatch From ��ҩ��������;

  --��鵱ǰ���˵�ҽ���Ƿ��н�����Ҫִ�е���Һ��������״̬��
  If Instr(v_ҽ��ids, ',') = 0 Then
    v_Tansid := v_ҽ��ids;
  Else
    v_Tansid := Substr(v_Tmp, 1, Instr(v_ҽ��ids, ',') - 1);
  End If;

  Select Count(ID)
  Into n_Num
  From ��Һ��ҩ��¼
  Where �Ƿ����� = 1 And ִ��ʱ�� Between Trunc(v_Currdate) And Trunc(v_Currdate + 1) - 1 / 24 / 60 / 60 And
        ҽ��id In
        (Select ���id
         From ����ҽ����¼
         Where ����id = (Select ����id From ����ҽ����¼ Where ���id = v_Tansid And Rownum < 2) And (������� = '5' Or ������� = '6')) And
        Rownum < 2;

  If n_Num > 0 Then
    Select ����
    Into v_��ǰ����
    From ��Һ��ҩ��¼
    Where �Ƿ����� = 1 And ִ��ʱ�� Between Trunc(v_Currdate) And Trunc(v_Currdate + 1) - 1 / 24 / 60 / 60 And
          ҽ��id In
          (Select ���id
           From ����ҽ����¼
           Where ����id = (Select ����id From ����ҽ����¼ Where ���id = v_Tansid And Rownum < 2) And (������� = '5' Or ������� = '6')) And
          Rownum < 2;
    Raise Err_Item;
  End If;

  --�Ƚ�ԭ�շ���¼����������µ��շ���¼��������ɾ��
  --Update ҩƷ�շ���¼
  --Set ��� = ��� + 10000
  --Where ID In (Select \*+rule *\
  --             Distinct c.Id
  --             From ����ҽ����¼ A, ����ҽ������ B, ҩƷ�շ���¼ C, סԺ���ü�¼ D, Table(f_Num2list(ҽ��id_In)) F
  --             Where a.Id = b.ҽ��id And c.����id = d.Id And a.Id = d.ҽ����� And b.No = c.No And b.ִ�в���id + 0 = ����id_In And
  --                   c.���� = 9 And c.������� Is Null And a.���id = f.Column_Value And b.���ͺ� = ���ͺ�_In And c.��� < 10000);

  For v_ҽ����¼ In c_ҽ����¼ Loop
    v_Continue := 1;
    n_����id   := v_ҽ����¼.����id;
    n_����id   := v_ҽ����¼.���˿���id;
  
    Select Count(1)
    Into v_Continue
    From (Select 1
           From ����ҽ����¼ A, ��Һ������ҩƷ B, סԺ���ü�¼ C
           Where c.�շ�ϸĿid = b.ҩƷid And c.ҽ����� = a.Id And a.���id = v_ҽ����¼.���id
           Union All
           Select 1
           From ����ҽ����¼ A, ��Һ������ҩƷ B, ������ü�¼ C
           Where c.�շ�ϸĿid = b.ҩƷid And c.ҽ����� = a.Id And a.���id = v_ҽ����¼.���id);
    If v_Continue = 0 Then
      v_Continue := 1;
    Else
      v_Continue := 0;
    End If;
  
    --�������Ʋ�����Һ��
    If (v_ҽ������ = 1 And v_ҽ����¼.ҽ������ <> 1) Or (v_ҽ������ = 2 And v_ҽ����¼.ҽ������ <> 2) Then
      v_Continue := 0;
    End If;
  
    If Not v_����Һ��ҩ;�� Is Null Then
      If Instr(',' || v_����Һ��ҩ;�� || ',', ',' || v_ҽ����¼.��ҩ;�� || ',') = 0 Then
        v_Continue := 0;
      End If;
    End If;
  
    If Not v_��Դ���� Is Null Then
      If Instr(',' || v_��Դ���� || ',', ',' || v_ҽ����¼.���˿���id || ',') = 0 Then
        v_Continue := 0;
      End If;
    End If;
  
    v_ҩƷ���� := Null;
    For r_ҩƷ���� In (Select Decode(Nvl(d.������, 0), 0, Decode(Nvl(d.�Ƿ�����ҩ, 0), 0, '', '����ҩ'), '������') ҩƷ����
                   From ����ҽ����¼ A, ҩƷ��� B, סԺ���ü�¼ C, ҩƷ���� D
                   Where c.�շ�ϸĿid = b.ҩƷid And b.ҩ��id = d.ҩ��id And c.ҽ����� = a.Id And a.���id = v_ҽ����¼.���id
                   Union All
                   Select Decode(Nvl(d.������, 0), 0, Decode(Nvl(d.�Ƿ�����ҩ, 0), 0, '', '����ҩ'), '������') ҩƷ����
                   From ����ҽ����¼ A, ҩƷ��� B, ������ü�¼ C, ҩƷ���� D
                   Where c.�շ�ϸĿid = b.ҩƷid And b.ҩ��id = d.ҩ��id And c.ҽ����� = a.Id And a.���id = v_ҽ����¼.���id) Loop
      If r_ҩƷ����.ҩƷ���� Is Not Null Then
        v_ҩƷ���� := r_ҩƷ����.ҩƷ����;
      End If;
    End Loop;
  
    If v_ҩƷ���� Is Null Then
      If v_ҽ����¼.�Ƿ�tpn = 2 Then
        v_ҩƷ���� := 'Ӫ��ҩ';
      End If;
    End If;
  
    If v_Continue = 1 Then
      v_Old���id := v_New���id;
      v_���id    := v_ҽ����¼.���id;
      v_New���id := v_���id;
      v_���ͺ�    := v_ҽ����¼.���ͺ�;
      v_���      := 0;
    
      If v_Continue = 1 Then
        --v_Count := Zl_Gettransexenumber(v_ҽ����¼.��ʼִ��ʱ��, v_ҽ����¼.�״�ʱ��, v_ҽ����¼.ĩ��ʱ��, v_ҽ����¼.Ƶ�ʼ��, v_ҽ����¼.�����λ, v_ҽ����¼.ִ��ʱ�䷽��);
        Select Count(ҽ��id)
        Into v_Count
        From ҽ��ִ��ʱ��
        Where ҽ��id = v_ҽ����¼.���id And ���ͺ� = v_ҽ����¼.���ͺ�;
      
        v_Nodosage := 0;
      
        For I In 1 .. v_Count Loop
          Select ��Һ��ҩ��¼_Id.Nextval Into v_��ҩid From Dual;
          v_��� := v_��� + 1;
        
          If I > 1 Then
            --��ҽ��ִ��ʱ�����ȡҽ����ִ��ʱ��
            Select Ҫ��ʱ��
            Into v_ִ��ʱ��
            From ҽ��ִ��ʱ��
            Where ҽ��id = v_ҽ����¼.���id And ���ͺ� = v_ҽ����¼.���ͺ� And Ҫ��ʱ�� > v_ִ��ʱ�� And Rownum = 1
            Order By Ҫ��ʱ��;
          Else
            Select Min(Ҫ��ʱ��)
            Into v_ִ��ʱ��
            From ҽ��ִ��ʱ��
            Where ҽ��id = v_ҽ����¼.���id And ���ͺ� = v_ҽ����¼.���ͺ� And Rownum = 1
            Order By Ҫ��ʱ��;
          End If;
        
          v_���� := 0;
        
          If d_Oldִ��ʱ�� <> Trunc(v_ִ��ʱ��) Or d_Oldִ��ʱ�� Is Null Then
            b_Change := True;
          End If;
        
          If b_Change = True Then
            If d_Oldִ��ʱ�� <> Trunc(v_ִ��ʱ��) Or d_Oldִ��ʱ�� Is Null Then
              d_Oldִ��ʱ�� := v_ִ��ʱ��;
            
              Select Count(Distinct a.��ҩ����)
              Into n_��ҩ��
              From ��Һ��ҩ��¼ A
              Where a.ҽ��id In (Select ID From ����ҽ����¼ Where ����id = v_ҽ����¼.����id And ���id Is Null) And
                    a.ִ��ʱ�� Between Trunc(v_ִ��ʱ�� - 1) And Trunc(v_ִ��ʱ��) - 1 / 24 / 60 / 60 And ����״̬ >= 2 And ����״̬ < 9;
            
              If n_��ҩ�� > 1 Then
                Update ��Һ��ҩ��¼
                Set �Ƿ�������� = 1
                Where ҽ��id In (Select ID From ����ҽ����¼ Where ����id = n_����id) And
                     
                      ִ��ʱ�� Between Trunc(v_ִ��ʱ��) And Trunc(v_ִ��ʱ�� + 1) - 1 / 24 / 60 / 60 And ����״̬ < 2;
                b_Change := False;
              
              End If;
            End If;
          End If;
        
          If b_Change = True Then
            n_����id := v_ҽ����¼.����id;
            Select Count(ID)
            
            Into n_Sum
            From ��Һ��ҩ��¼
            Where ҽ��id = v_ҽ����¼.���id And ִ��ʱ�� Between Trunc(v_ִ��ʱ�� - 1) And Trunc(v_ִ��ʱ��) - 1 / 24 / 60 / 60;
            If n_Sum = 0 Then
              Update ��Һ��ҩ��¼
              Set �Ƿ�������� = 1
              Where ҽ��id In (Select ID From ����ҽ����¼ Where ����id = n_����id) And
                   
                    ִ��ʱ�� Between Trunc(v_ִ��ʱ��) And Trunc(v_ִ��ʱ�� + 1) - 1 / 24 / 60 / 60 And ����״̬ < 2;
              b_Change := False;
            
            End If;
          
            If b_Change = True Then
              --�����Һ���Ƿ���������״̬
              Select Count(a.Id)
              Into n_Sum
              From ��Һ��ҩ��¼ A, ��Һ��ҩ���� B, ҩƷ�շ���¼ C
              Where a.Id = b.��¼id And b.�շ�id = c.Id And
                    a.ҽ��id In (Select ID
                               From ����ҽ����¼
                               Where ����id = (Select ����id From ����ҽ����¼ Where ID = v_ҽ����¼.���id And Rownum < 2)) And
                    a.ִ��ʱ�� Between Trunc(v_ִ��ʱ�� - 1) And Trunc(v_ִ��ʱ��) - 1 / 24 / 60 / 60 And a.���ʱ�� Is Not Null;
              If n_Sum <> 0 Then
                Update ��Һ��ҩ��¼
                Set �Ƿ�������� = 1
                Where ҽ��id In (Select ID From ����ҽ����¼ Where ����id = n_����id) And ִ��ʱ�� Between Trunc(v_ִ��ʱ��) And
                      Trunc(v_ִ��ʱ�� + 1) - 1 / 24 / 60 / 60 And ����״̬ < 2;
                b_Change := False;
              End If;
            
              Select Count(ҽ��id)
              Into n_Cur
              From ҽ��ִ��ʱ��
              Where ҽ��id = v_ҽ����¼.���id And Ҫ��ʱ�� Between Trunc(v_ִ��ʱ��) And Trunc(v_ִ��ʱ�� + 1) - 1 / 24 / 60 / 60;
            
              Select Count(ҽ��id)
              Into n_Sum
              From ҽ��ִ��ʱ��
              Where ҽ��id = v_ҽ����¼.���id And Ҫ��ʱ�� Between Trunc(v_ִ��ʱ�� - 1) And Trunc(v_ִ��ʱ��) - 1 / 24 / 60 / 60;
            
              If n_Sum <> n_Cur Then
                Update ��Һ��ҩ��¼
                Set �Ƿ�������� = 1
                Where ҽ��id In (Select ID From ����ҽ����¼ Where ����id = n_����id) And ִ��ʱ�� Between Trunc(v_ִ��ʱ��) And
                      Trunc(v_ִ��ʱ�� + 1) - 1 / 24 / 60 / 60 And ����״̬ < 2;
                b_Change := False;
              End If;
            End If;
          End If;
        
          If v_ʱ�䴮 <> Trunc(Sysdate) || ';false\' Or v_ʱ�䴮 Is Null Then
            If Trunc(v_ִ��ʱ��) = Trunc(Sysdate) Then
              If b_Change = False Then
                v_ʱ�䴮 := Trunc(v_ִ��ʱ��) || ';false\';
              Else
                v_ʱ�䴮 := Trunc(v_ִ��ʱ��) || ';true\';
              End If;
            End If;
          End If;
        
          If v_ʱ�䴮1 <> Trunc(Sysdate + 1) || ';false\' Or v_ʱ�䴮1 Is Null Then
            If Trunc(v_ִ��ʱ��) = Trunc(Sysdate + 1) Then
              If b_Change = False Then
                v_ʱ�䴮1 := Trunc(v_ִ��ʱ��) || ';false\';
              Else
                v_ʱ�䴮1 := Trunc(v_ִ��ʱ��) || ';true\';
              End If;
            End If;
          End If;
        
          If v_ҩƷ���� Is Null Or n_����ҩƷ���� = 0 Then
            v_���� := Zl_Getpivaworkbatch(v_ִ��ʱ��, Sysdate);
          Else
            --ҩƷ���Ͳ�Ϊ�գ�ֱ�Ӹ���ҩƷ����ƥ������
            v_���� := Zl_Getpivaworkbatch(v_ִ��ʱ��, Sysdate, v_ҩƷ����);
          End If;
        
          Select Count(ҽ��id)
          Into n_���ʹ���
          From ҽ��ִ��ʱ��
          Where ҽ��id = v_ҽ����¼.���id And Ҫ��ʱ�� <= v_ִ��ʱ��
          Order By Ҫ��ʱ��;
        
          If n_���ʹ��� > 99 Then
            n_���ʹ��� := Mod(n_���ʹ���, 99);
          End If;
        
          If Length(v_ҽ����¼.���id) > 9 Then
            If n_���ʹ��� < 10 Then
              Select '91' || Substr(To_Char(v_ҽ����¼.���id), Length(v_ҽ����¼.���id) - 8) || To_Char(v_ҽ����¼.���id) || '0' ||
                      To_Char(n_���ʹ���)
              Into v_Maxno
              From Dual;
            Else
              Select '91' || Substr(To_Char(v_ҽ����¼.���id), Length(v_ҽ����¼.���id) - 8) || To_Char(v_ҽ����¼.���id) ||
                      To_Char(n_���ʹ���)
              Into v_Maxno
              From Dual;
            End If;
          Else
            If n_���ʹ��� < 10 Then
              Select '91' || Substr('000000000', Length(v_ҽ����¼.���id) + 1) || To_Char(v_ҽ����¼.���id) || '0' ||
                      To_Char(n_���ʹ���)
              Into v_Maxno
              From Dual;
            Else
              Select '91' || Substr('000000000', Length(v_ҽ����¼.���id) + 1) || To_Char(v_ҽ����¼.���id) || To_Char(n_���ʹ���)
              Into v_Maxno
              From Dual;
            End If;
          End If;
          n_�������� := 0;
          If b_Change = False Then
            n_�������� := 1;
          End If;
        
          If v_���� <> 0 Then
            Select Nvl(Max(���), 0), Max(ҩƷ����)
            Into n_���, v_��ҩ����
            From ��ҩ��������
            Where ���� = v_���� And ��������id = ����id_In;
          End If;
        
          If (Trunc(v_ִ��ʱ��) <= v_Currdate Or n_��� <> 0) And v_��ҩ���� Is Null Then
            n_�Ƿ���     := 1;
            d_�ֹ����ʱ�� := Null;
          Else
            n_�Ƿ���     := 0;
            d_�ֹ����ʱ�� := Null;
          End If;
        
          --�����TPN��������������ζ�����Ϊ����
          If v_ҽ����¼.�Ƿ�tpn = 2 Then
            n_�Ƿ���     := 0;
            d_�ֹ����ʱ�� := Null;
          End If;
        
          If v_���� = 0 Then
            n_�Ƿ��� := 1;
          End If;
          --������ҩ��¼
          Insert Into ��Һ��ҩ��¼
            (ID, ����id, ���, ����, �Ա�, ����, סԺ��, ����, ���˲���id, ���˿���id, ִ��ʱ��, ҽ��id, ���ͺ�, ��ҩ����, ƿǩ��, �Ƿ��������, �Ƿ���, ���ʱ��, ����״̬,
             ������Ա, ����ʱ��)
          Values
            (v_��ҩid, ����id_In, v_���, v_ҽ����¼.����, v_ҽ����¼.�Ա�, v_ҽ����¼.����, v_ҽ����¼.סԺ��, v_ҽ����¼.����, v_ҽ����¼.���˲���id,
             v_ҽ����¼.���˿���id, v_ִ��ʱ��, v_ҽ����¼.���id, v_ҽ����¼.���ͺ�, v_����, v_Maxno, n_��������, n_�Ƿ���, d_�ֹ����ʱ��, 1, �˲���_In, �˲�ʱ��_In);
        
          Insert Into ��Һ��ҩ״̬ (��ҩid, ��������, ������Ա, ����ʱ��) Values (v_��ҩid, 1, �˲���_In, �˲�ʱ��_In);
        
          For v_����ҽ����¼ In c_����ҽ����¼ Loop
            n_ҽ��id   := v_����ҽ����¼.ҽ��id;
            n_�ۼ����� := 0;
            n_ʣ������ := 0;
          
            Select Nvl(Sum(ʵ������), 0)
            Into n_Sum
            From (Select c.ʵ������
                   From ����ҽ����¼ A, ����ҽ������ B, ҩƷ�շ���¼ C, סԺ���ü�¼ D
                   Where a.Id = b.ҽ��id And c.����id = d.Id And a.Id = d.ҽ����� And b.No = c.No And b.ִ�в���id + 0 = ����id_In And
                         c.���� = 9 And c.������� Is Null And a.Id = n_ҽ��id And b.���ͺ� = v_ҽ����¼.���ͺ� And c.��� < 1000
                   Union All
                   Select c.ʵ������
                   From ����ҽ����¼ A, ����ҽ������ B, ҩƷ�շ���¼ C, ������ü�¼ D
                   Where a.Id = b.ҽ��id And c.����id = d.Id And a.Id = d.ҽ����� And b.No = c.No And b.ִ�в���id + 0 = ����id_In And
                         c.���� = 9 And c.������� Is Null And a.Id = n_ҽ��id And b.���ͺ� = v_ҽ����¼.���ͺ� And c.��� < 1000);
          
            --������ҩ��¼��Ӧ��ҩƷ��¼
            For v_�շ���¼ In c_�շ���¼ Loop
              If v_�շ���¼.�Ƿ������� = 1 Then
                v_Nodosage := 1;
              End If;
            
              Select ҩƷ�շ���¼_Id.Nextval Into n_Lngid From Dual;
              n_�ۼ����� := n_�ۼ����� + v_�շ���¼.����;
            
              If n_ʣ������ = 0 Then
                n_ʣ������ := n_Sum / v_Count;
              End If;
              n_�������� := n_Sum / v_Count;
            
              If n_�ۼ����� >= n_Sum / v_Count * I Then
                n_Count := n_Count + 1;
                Insert Into ҩƷ�շ���¼
                  (ID, ��¼״̬, ����, NO, ���, �ⷿid, ��ҩ��λid, ������id, �Է�����id, ���ϵ��, ҩƷid, ����, ����, ����, ��������, Ч��, ����, ��д����, ʵ������,
                   �ɱ���, �ɱ����, ����, ���ۼ�, ���۽��, ���, ժҪ, ������, ��������, ��ҩ��, ��ҩ����, �����, �������, �۸�id, ����id, ����, Ƶ��, �÷�, ���, �������,
                   ���Ч��, ��Ʒ�ϸ�֤, ��ҩ��ʽ, ��ҩ����, ������, ��׼�ĺ�, ���ܷ�ҩ��, ע��֤��, �ⷿ��λ, ��Ʒ����, �ڲ�����, �˲���, �˲�����, ǩ��ȷ����, ǩ��ʱ��)
                  Select n_Lngid, ��¼״̬, ����, NO, n_Count + 1000, �ⷿid, ��ҩ��λid, ������id, �Է�����id, ���ϵ��, ҩƷid, ����, ����, ����,
                         ��������, Ч��, ����, n_ʣ������, n_ʣ������, �ɱ���, Round(�ɱ��� * n_ʣ������, n_��ͨ���С��), ����, ���ۼ�,
                         Round(���ۼ� * n_ʣ������, n_��ͨ���С��), Round(��� * (ʵ������ / n_ʣ������), n_��ͨ���С��), '����', ������, ��������, ��ҩ��,
                         ��ҩ����, �����, �������, �۸�id, ����id, ����, Ƶ��, �÷�, ���, �������, ���Ч��, ��Ʒ�ϸ�֤, ��ҩ��ʽ, ��ҩ����, ������, ��׼�ĺ�, ���ܷ�ҩ��,
                         ע��֤��, �ⷿ��λ, ��Ʒ����, �ڲ�����, �˲���, �˲�����, ǩ��ȷ����, ǩ��ʱ��
                  From ҩƷ�շ���¼
                  Where ID = v_�շ���¼.�շ�id;
              
                Zl_δ��ҩƷ��¼_Insert(n_Lngid);
              
                Insert Into ��Һ��ҩ���� (��¼id, �շ�id, ����) Values (v_��ҩid, n_Lngid, n_ʣ������);
              
                n_ʣ������ := 0;
                Exit;
              Elsif n_�ۼ����� > (n_Sum / v_Count * (I - 1)) Then
                n_Count    := n_Count + 1;
                n_��д���� := n_�ۼ����� - (n_Sum / v_Count * (I - 1)) - (n_�������� - n_ʣ������);
                Insert Into ҩƷ�շ���¼
                  (ID, ��¼״̬, ����, NO, ���, �ⷿid, ��ҩ��λid, ������id, �Է�����id, ���ϵ��, ҩƷid, ����, ����, ����, ��������, Ч��, ����, ��д����, ʵ������,
                   �ɱ���, �ɱ����, ����, ���ۼ�, ���۽��, ���, ժҪ, ������, ��������, ��ҩ��, ��ҩ����, �����, �������, �۸�id, ����id, ����, Ƶ��, �÷�, ���, �������,
                   ���Ч��, ��Ʒ�ϸ�֤, ��ҩ��ʽ, ��ҩ����, ������, ��׼�ĺ�, ���ܷ�ҩ��, ע��֤��, �ⷿ��λ, ��Ʒ����, �ڲ�����, �˲���, �˲�����, ǩ��ȷ����, ǩ��ʱ��)
                  Select n_Lngid, ��¼״̬, ����, NO, n_Count + 1000, �ⷿid, ��ҩ��λid, ������id, �Է�����id, ���ϵ��, ҩƷid, ����, ����, ����,
                         ��������, Ч��, ����, n_��д����, n_��д����, �ɱ���, Round(�ɱ��� * n_��д����, n_��ͨ���С��), ����, ���ۼ�,
                         Round(���ۼ� * n_��д����, n_��ͨ���С��), Round(��� * (ʵ������ / n_��д����), n_��ͨ���С��), '����', ������, ��������, ��ҩ��,
                         ��ҩ����, �����, �������, �۸�id, ����id, ����, Ƶ��, �÷�, ���, �������, ���Ч��, ��Ʒ�ϸ�֤, ��ҩ��ʽ, ��ҩ����, ������, ��׼�ĺ�, ���ܷ�ҩ��,
                         ע��֤��, �ⷿ��λ, ��Ʒ����, �ڲ�����, �˲���, �˲�����, ǩ��ȷ����, ǩ��ʱ��
                  From ҩƷ�շ���¼
                  Where ID = v_�շ���¼.�շ�id;
              
                Zl_δ��ҩƷ��¼_Insert(n_Lngid);
              
                Insert Into ��Һ��ҩ���� (��¼id, �շ�id, ����) Values (v_��ҩid, n_Lngid, n_��д����);
              
                n_ʣ������ := n_ʣ������ - n_��д����;
              End If;
            End Loop;
          End Loop;
          n_���ȼ� := Zl_Getfirst(v_��ҩid, v_ҽ����¼.���˿���id);
          Update ��Һ��ҩ��¼ Set ���ȼ� = n_���ȼ� Where ID = v_��ҩid;
        
        End Loop;
      
        For v_�շ���¼ In c_ԭʼ�շ���¼ Loop
          n_���� := v_�շ���¼.����;
        
          v_No := v_�շ���¼.No;
          Delete From ҩƷ�շ���¼ Where ID = v_�շ���¼.�շ�id;
        End Loop;
      
        --����ҩƷ���߲������õ�ҩƷĬ��Ϊ0����
        Select Count(�շ�id) Into n_Row From ��Һ��ҩ���� Where ��¼id = v_��ҩid;
        If (v_Nodosage = 1 Or n_Row = 1) And n_���ҩƷ���� = 1 Then
          Update ��Һ��ҩ��¼
          Set ��ҩ���� = 0, �Ƿ��� = 1
          Where ҽ��id = v_ҽ����¼.���id And ���ͺ� = v_ҽ����¼.���ͺ� And ����״̬ < 2;
        End If;
        --������ڡ��������á����Ե�ҩƷ��Ҳ����Ϊ���
        If v_Nodosage = 1 Then
          Update ��Һ��ҩ��¼
          Set �Ƿ��� = 1
          Where ҽ��id = v_ҽ����¼.���id And ���ͺ� = v_ҽ����¼.���ͺ� And ����״̬ < 2;
        End If;
      End If;
    End If;
  End Loop;

  For v_�շ���¼ In (Select ID From ҩƷ�շ���¼ Where ��� < 1000 And ���� = n_���� And NO = v_No) Loop
    n_Count := n_Count + 1;
    Update ҩƷ�շ���¼ Set ��� = n_Count + 1000, ժҪ = '����' Where ID = v_�շ���¼.Id;
  End Loop;

  Update ҩƷ�շ���¼
  Set ��� = ��� - 1000, ժҪ = 'ҽ������'
  Where ժҪ = '����' And ��� > 1000 And ���� = n_���� And NO = v_No;

  If n_�������� = 1 Then
  
    Select Count(a.Id)
    Into n_Sum
    From ��Һ��ҩ��¼ A, ����ҽ������ B
    Where a.ҽ��id = b.ҽ��id And a.���ͺ� = b.���ͺ� And
          a.ҽ��id In (Select ID From ����ҽ����¼ Where ����id = n_����id And ���id Is Null) And b.����ʱ�� Between Trunc(Sysdate) And
          Trunc(Sysdate + 1) - 1 / 24 / 60 / 60 And a.ִ��ʱ�� Between Trunc(Sysdate) And
          Trunc(Sysdate + 1) - 1 / 24 / 60 / 60 And ����״̬ < 9;
    If n_Sum <> 0 Then
      b_Change  := False;
      v_ʱ�䴮1 := Trunc(Sysdate + 1) || ';false\';
    
      Update ��Һ��ҩ��¼
      Set �Ƿ�������� = 1
      Where ҽ��id In (Select ID From ����ҽ����¼ Where ����id = n_����id) And ִ��ʱ�� Between Trunc(Sysdate + 1) And
            Trunc(Sysdate + 2) - 1 / 24 / 60 / 60 And ����״̬ < 2;
    End If;
  End If;
  If v_ʱ�䴮 Is Null Then
    v_ʱ�䴮 := v_ʱ�䴮1;
  Else
    v_ʱ�䴮 := v_ʱ�䴮 || v_ʱ�䴮1;
  End If;

  While v_ʱ�䴮 Is Not Null Loop
    --�ֽⵥ��ID��
    v_Fields   := Substr(v_ʱ�䴮, 1, Instr(v_ʱ�䴮, '\') - 1);
    v_ʱ��ֵ   := Substr(v_Fields, 1, Instr(v_Fields, ';') - 1);
    v_�Ƿ�ı� := Substr(v_Fields, Instr(v_Fields, ';') + 1);
  
    v_ʱ�䴮 := Replace('\' || v_ʱ�䴮, '\' || v_Fields || '\');
  
    If v_�Ƿ�ı� = 'true' Then
      b_Change := True;
    Else
      b_Change := False;
    End If;

    If b_Change = True Then
      Select Count(ҽ��id)
      Into n_Cur
      From (Select Distinct a.Ҫ��ʱ��, a.ҽ��id
             From ҽ��ִ��ʱ�� A, ��Һ��ҩ��¼ B
             Where a.Ҫ��ʱ�� + 0 = b.ִ��ʱ�� And a.ҽ��id = b.ҽ��id And b.ִ��ʱ�� + 0 Between Trunc(v_ʱ��ֵ) And
                   Trunc(v_ʱ��ֵ + 1) - 1 / 24 / 60 / 60 And
                   b.ҽ��id In (Select ID From ����ҽ����¼ Where ����id = n_����id And ���id Is Null));
    
      Select Count(ҽ��id)
      Into n_Sum
      From (Select Distinct a.Ҫ��ʱ��, a.ҽ��id
             From ҽ��ִ��ʱ�� A, ��Һ��ҩ��¼ B
             Where a.Ҫ��ʱ�� + 0 = b.ִ��ʱ�� And a.ҽ��id = b.ҽ��id And b.ִ��ʱ�� + 0 Between Trunc(v_ʱ��ֵ - 1) And
                   Trunc(v_ʱ��ֵ) - 1 / 24 / 60 / 60 And
                   b.ҽ��id In (Select ID From ����ҽ����¼ Where ����id = n_����id And ���id Is Null));
    
      If n_Cur <> n_Sum Then
        Update ��Һ��ҩ��¼
        Set �Ƿ�������� = 1
        Where ҽ��id In (Select ID From ����ҽ����¼ Where ����id = n_����id) And ִ��ʱ�� Between Trunc(v_ʱ��ֵ) And
              Trunc(v_ʱ��ֵ + 1) - 1 / 24 / 60 / 60 And ����״̬ < 2;
        b_Change := False;
      End If;
    End If;
  
    If v_�����ϴ����� = 1 And b_Change = True Then
      For v_��Һ����¼ In c_��Һ����¼ Loop
        Begin
          Select Distinct ��ҩ����
          Into v_����
          From ��Һ��ҩ��¼ A, ��Һ��ҩ���� B, ҩƷ�շ���¼ C
          Where a.Id = b.��¼id And b.�շ�id = c.Id And a.ҽ��id = v_��Һ����¼.ҽ��id And
                To_Char(a.ִ��ʱ��, 'hh24:mi:ss') = To_Char(v_��Һ����¼.ִ��ʱ��, 'hh24:mi:ss') And
                a.ִ��ʱ�� Between Trunc(v_��Һ����¼.ִ��ʱ�� - 1) And Trunc(v_��Һ����¼.ִ��ʱ��) - 1 / 24 / 60 / 60 And Rownum = 1;
        Exception
          When Others Then
            Begin
              Select Distinct ��ҩ����
              Into v_����
              From ��Һ��ҩ��¼ A
              Where a.ҽ��id = v_��Һ����¼.ҽ��id And To_Char(a.ִ��ʱ��, 'hh24:mi:ss') = To_Char(v_��Һ����¼.ִ��ʱ��, 'hh24:mi:ss') And
                    a.����״̬ <> 12 And a.ִ��ʱ�� Between Trunc(v_��Һ����¼.ִ��ʱ�� - 1) And Trunc(v_��Һ����¼.ִ��ʱ��) - 1 / 24 / 60 / 60 And
                    Rownum = 1;
            Exception
              When Others Then
                v_���� := v_��Һ����¼.��ҩ����;
            End;
        End;
      
        Update ��Һ��ҩ��¼
        Set �Ƿ�ȷ�ϵ��� = 0, �Ƿ�������� = 0
        Where ҽ��id In (Select ID From ����ҽ����¼ Where ����id = n_����id) And ִ��ʱ�� Between Trunc(v_��Һ����¼.ִ��ʱ��) And
              Trunc(v_��Һ����¼.ִ��ʱ�� + 1) - 1 / 24 / 60 / 60 And ����״̬ < 2;
      
        If v_��Һ����¼.��ҩ���� <> v_���� Then
          Update ��Һ��ҩ��¼ Set ��ҩ���� = v_���� Where ID = v_��Һ����¼.Id;
          Select Nvl(Max(���), 0) Into n_��� From ��ҩ�������� Where ���� = v_���� And ��������id = ����id_In;
          If n_��� <> 0 Then
            Update ��Һ��ҩ��¼ Set �Ƿ��� = n_��� Where ID = v_��Һ����¼.Id;
          Else
            Select Nvl(Max(���), 0)
            Into n_���
            From ��ҩ��������
            Where ���� = v_��Һ����¼.��ҩ���� And ��������id = ����id_In;
          
            If n_��� <> 0 Then
              Update ��Һ��ҩ��¼ Set �Ƿ��� = 0 Where ID = v_��Һ����¼.Id;
            End If;
          End If;
        End If;
      End Loop;
    End If;
  
    If n_�Զ����� = 1 And (b_Change = False Or v_�����ϴ����� = 0) Then
      For v_��Һ����¼ In c_��Һ����¼ Loop
        v_���� := Zl_Getpivaworkbatch(v_��Һ����¼.ִ��ʱ��, v_��Һ����¼.����ʱ��);
        Update ��Һ��ҩ��¼ Set ��ҩ���� = v_���� Where ID = v_��Һ����¼.Id;
      End Loop;
      Zl_��Һ��ҩ��¼_�Զ�����(n_����id, n_����id, ����id_In, v_ʱ��ֵ);
    End If;
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]����' || v_��ǰ���� || '����Һ���������б���������Һ��������ʧ�ܣ�[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Һ��ҩ��¼_�˲�;
/

--139595:��ҵ��,2019-04-15,����֧�������������
Create Or Replace Procedure Zl_��Һ��ҩ��¼_��ҩ
(
  ��ҩid_In   In Varchar2, --ID��:ID1,ID2.... 
  ������Ա_In In ��Һ��ҩ��¼.������Ա%Type := Null,
  ����ʱ��_In In ��Һ��ҩ��¼.����ʱ��%Type := Null,
  ����˵��_In In ��Һ��ҩ״̬.����˵��%Type := Null,
  �ƶ�����_In In Number := 0
) Is
  v_Tansid     Varchar2(20);
  v_Tmp        Varchar2(4000);
  v_No         Varchar2(20);
  v_Usercode   Varchar2(100);
  n_����״̬   ��Һ��ҩ��¼.����״̬%Type;
  v_Error      Varchar2(255);
  n_People     Number(1);
  n_Row        Number(2);
  d_ִ��ʱ��   Date;
  v_��ҩ����   Varchar2(50);
  n_��Ŀid     Number(18);
  v_�շ���Ŀid Varchar2(200);
  v_Info       Varchar2(200);
  v_Id         Varchar2(20);
  n_����       Number(2);
  n_Count      Number(18);
  n_Out        Number(10);
  n_Outnum     Number(10);
  n_���״̬   Number(1);
  v_�˶���     Varchar2(20);
  v_��Һ��     Varchar2(20);
  Err_Custom Exception;

  Cursor c_Bill Is
    Select a.����id, a.��ҳid, a.��ʶ��, a.����, a.�Ա�, a.����, a.����, a.�ѱ�, a.���˲���id, a.���˿���id, a.Ӥ����, e.ҩƷid, b.�ⷿid, f.��ҩ����, g.���,
           1 As �����־
    From סԺ���ü�¼ A, ҩƷ�շ���¼ B, ��Һ��ҩ��¼ C, ��Һ��ҩ���� D, ҩƷ��� E, ��ҺҩƷ���� F, �����շѷ��� G
    Where a.Id = b.����id And b.Id = d.�շ�id And d.��¼id = c.Id And b.ҩƷid = e.ҩƷid And b.ҩƷid = f.ҩƷid And
          Substr(f.��ҩ����, Instr(f.��ҩ����, '-') + 1) = g.��ҩ����(+) And Nvl(c.�Ƿ���, 0) <> 1 And c.Id = v_Tansid
    Union All
    Select a.����id, a.��ҳid, a.��ʶ��, a.����, a.�Ա�, a.����, '' As ����, a.�ѱ�, a.���˲���id, a.���˿���id, a.Ӥ����, e.ҩƷid, b.�ⷿid, f.��ҩ����,
           g.���, 0 As �����־
    From ������ü�¼ A, ҩƷ�շ���¼ B, ��Һ��ҩ��¼ C, ��Һ��ҩ���� D, ҩƷ��� E, ��ҺҩƷ���� F, �����շѷ��� G
    Where a.Id = b.����id And b.Id = d.�շ�id And d.��¼id = c.Id And b.ҩƷid = e.ҩƷid And b.ҩƷid = f.ҩƷid And
          Substr(f.��ҩ����, Instr(f.��ҩ����, '-') + 1) = g.��ҩ����(+) And Nvl(c.�Ƿ���, 0) <> 1 And c.Id = v_Tansid
    Order By ���;

Begin
  v_Usercode := Zl_Identity;
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ';') + 1);
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ',') + 1);
  v_Usercode := Substr(v_Usercode, 1, Instr(v_Usercode, ',') - 1);
  n_People   := Nvl(zl_GetSysParameter('���÷Ѱ�������ȡ', 1345), 0);
  n_Out      := Nvl(zl_GetSysParameter('��Ժ���˲������÷�', 1345), 0);

  If ��ҩid_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := ��ҩid_In || ',';
  End If;

  While v_Tmp Is Not Null Loop
    --�ֽⵥ��ID�� 
    v_Tansid := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp    := Replace(',' || v_Tmp, ',' || v_Tansid || ',');
  
    --��鵱ǰ��Һ����״̬�Ƿ�Ϊ����ҩ״̬ 
    Begin
      Select ����״̬, ִ��ʱ��, Nvl(�Ƿ���, 0)
      Into n_����״̬, d_ִ��ʱ��, n_���״̬
      From ��Һ��ҩ��¼
      Where ID = v_Tansid;
    
      If n_����״̬ > 3 Then
        v_Error := '�������ѱ����������ܽ��з�ҩ��';
        Raise Err_Custom;
      End If;
    Exception
      When Others Then
        v_Error := '�������ѱ�������';
        Raise Err_Custom;
    End;
  
    v_�˶��� := '';
    v_��Һ�� := '';
    Begin
      Select �˶���, ��Һ��
      Into v_�˶���, v_��Һ��
      From (Select e.�˶���, e.��Һ��
             From ��Һ��ҩ��¼ A, ��Һ��ҩ���� B, ҩƷ�շ���¼ C, ��Һ̨ҩƷ���� D, ��Һ�������� E
             Where a.Id = b.��¼id And b.�շ�id = c.Id And c.ҩƷid = d.ҩƷid And c.�ⷿid = d.����id And d.��ҩ̨id = e.��ҩ̨id And
                   a.��ҩ���� = e.���� And e.���� = To_Date(To_Char(Sysdate, 'yyyy-mm-dd'), 'yyyy-mm-dd') And a.Id = v_Tansid And
                   Rownum = 1
             Order By d.��ҩ̨id)
      Where Rownum = 1;
    Exception
      When Others Then
        Null;
    End;
  
    Update ��Һ��ҩ��¼ Set ����״̬ = 4, ������Ա = ������Ա_In, ����ʱ�� = ����ʱ��_In Where ID = v_Tansid;
    If �ƶ�����_In = 0 Then
      Insert Into ��Һ��ҩ״̬
        (��ҩid, ��������, ������Ա, ����ʱ��, ����˵��, ʵ�ʹ�����Ա)
      Values
        (v_Tansid, 3, ������Ա_In, ����ʱ��_In, ����˵��_In, v_�˶���);
    End If;
    Insert Into ��Һ��ҩ״̬
      (��ҩid, ��������, ������Ա, ����ʱ��, ����˵��, ʵ�ʹ�����Ա)
    Values
      (v_Tansid, 4, ������Ա_In, ����ʱ��_In, ����˵��_In, v_��Һ��);
  
    If n_���״̬ = 0 Then
      n_Count := 0;
      Select Nextno(14) Into v_No From Dual;
      For r_Bill In c_Bill Loop
        Select Count(����id)
        Into n_Outnum
        From ������ҳ
        Where ��ҳid = r_Bill.��ҳid And ����id = r_Bill.����id And (Nvl(״̬, 0) = 3 Or ��Ժ���� Is Not Null);
        If n_Count = 0 And (n_Outnum = 0 Or n_Out = 0) Then
          --��ȡ���Ϸ� 
          --v_�շ���Ŀid:='6970,2;6971,1;'; 
          Select Zl_Fun_Pivacustom(v_Tansid) Into v_�շ���Ŀid From Dual;
          While v_�շ���Ŀid Is Not Null Loop
            v_Info       := Substr(v_�շ���Ŀid, 1, Instr(v_�շ���Ŀid, ';') - 1);
            v_�շ���Ŀid := Replace(';' || v_�շ���Ŀid, ';' || v_Info || ';');
          
            v_Id   := Substr(v_Info, 1, Instr(v_Info, ',') - 1);
            v_Info := Replace(',' || v_Info, ',' || v_Id || ',');
          
            For r_Item In (Select a.Id �շ�ϸĿid, a.��� �շ����, a.���㵥λ, a.�Ӱ�Ӽ� �Ӱ��־, d.Id ������Ŀid, d.�վݷ�Ŀ, b.�ּ�
                           From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ D
                           Where a.Id = b.�շ�ϸĿid And b.������Ŀid = d.Id And a.Id = v_Id And b.ִ������ <= Sysdate And
                                 (b.��ֹ���� >= Sysdate Or b.��ֹ���� Is Null)) Loop
              If n_Count = 0 Then
                Insert Into ��Һ��ҩ���� (��ҩid, NO, ����id) Values (v_Tansid, v_No, r_Bill.����id);
              End If;
            
              n_Count := n_Count + 1;
            
              If r_Bill.�����־ = 1 Then
                Zl_סԺ���ʼ�¼_Insert(v_No, n_Count, r_Bill.����id, r_Bill.��ҳid, r_Bill.��ʶ��, r_Bill.����, r_Bill.�Ա�, r_Bill.����,
                                 r_Bill.����, r_Bill.�ѱ�, r_Bill.���˲���id, r_Bill.���˿���id, r_Item.�Ӱ��־, r_Bill.Ӥ����,
                                 r_Bill.�ⷿid, ������Ա_In, Null, r_Item.�շ�ϸĿid, r_Item.�շ����, r_Item.���㵥λ, Null, Null, Null,
                                 1, v_Info, Null, r_Bill.�ⷿid, Null, r_Item.������Ŀid, r_Item.�վݷ�Ŀ, r_Item.�ּ�,
                                 r_Item.�ּ� * v_Info, r_Item.�ּ� * v_Info, Null, Sysdate, Sysdate, Null, Null, v_Usercode,
                                 ������Ա_In);
              Else
                Zl_������ʼ�¼_Insert(v_No, n_Count, r_Bill.����id, r_Bill.��ʶ��, r_Bill.����, r_Bill.�Ա�, r_Bill.����, r_Bill.�ѱ�,
                                 r_Item.�Ӱ��־, r_Bill.Ӥ����, r_Bill.���˿���id, r_Bill.�ⷿid, ������Ա_In, Null, r_Item.�շ�ϸĿid,
                                 r_Item.�շ����, r_Item.���㵥λ, 1, 1, Null, r_Bill.�ⷿid, Null, r_Item.������Ŀid, r_Item.�վݷ�Ŀ,
                                 r_Item.�ּ�, r_Item.�ּ�, r_Item.�ּ�, Sysdate, Sysdate, Null, Null, v_Usercode, ������Ա_In);
              End If;
            End Loop;
          End Loop;
        End If;
      
        --�Ȳ�ѯ�Ƿ��а���ҩ;����ȡ�����÷ѷ���
        Select Nvl(Max(��Ŀid), 0)
        Into n_��Ŀid
        From ��Һ��ҩ��¼ A, ����ҽ����¼ B, �����շѷ��� C
        Where a.Id = v_Tansid And a.ҽ��id = b.Id And b.������Ŀid = c.����id;
        If n_��Ŀid = 0 Then
          --���޶�Ӧ��ҩ;�������÷���ȡ���������ٲ�ѯ�Ƿ��а���ҩ������ȡ�����÷ѷ���
          Select Nvl(Max(��Ŀid), 0)
          Into n_��Ŀid
          From �����շѷ���
          Where ��ҩ���� = Substr(r_Bill.��ҩ����, Instr(r_Bill.��ҩ����, '-', 1, 1) + 1);
        End If;
      
        If n_��Ŀid <> 0 Then
          n_Row := 0;
        
          If n_People = 1 Then
            Select Count(��ҩid)
            Into n_Row
            From (Select ��ҩid
                   From ��Һ��ҩ���� A, סԺ���ü�¼ B, ��Һ��ҩ��¼ C
                   Where a.No = b.No And a.��ҩid = c.Id And b.����id = r_Bill.����id And b.��¼״̬ = 1 And b.�շ�ϸĿid = n_��Ŀid And
                         d_ִ��ʱ�� Between Trunc(c.ִ��ʱ��) And Trunc(c.ִ��ʱ�� + 1) - 1 / 24 / 60 / 60
                   Union All
                   Select ��ҩid
                   From ��Һ��ҩ���� A, ������ü�¼ B, ��Һ��ҩ��¼ C
                   Where a.No = b.No And a.��ҩid = c.Id And b.����id = r_Bill.����id And b.��¼״̬ = 1 And b.�շ�ϸĿid = n_��Ŀid And
                         d_ִ��ʱ�� Between Trunc(c.ִ��ʱ��) And Trunc(c.ִ��ʱ�� + 1) - 1 / 24 / 60 / 60);
          End If;
        Else
          n_Row := 1;
        End If;
      
        If n_Row = 0 And (n_Outnum = 0 Or n_Out = 0) Then
          For r_Item In (Select a.Id �շ�ϸĿid, a.��� �շ����, a.���㵥λ, a.�Ӱ�Ӽ� �Ӱ��־, d.Id ������Ŀid, d.�վݷ�Ŀ, b.�ּ�
                         From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ D
                         Where a.Id = b.�շ�ϸĿid And b.������Ŀid = d.Id And a.Id = n_��Ŀid And b.ִ������ <= Sysdate And
                               (b.��ֹ���� >= Sysdate Or b.��ֹ���� Is Null)) Loop
            If n_Count = 0 Then
              Insert Into ��Һ��ҩ���� (��ҩid, NO, ����id) Values (v_Tansid, v_No, r_Bill.����id);
            End If;
          
            n_Count := n_Count + 1;
          
            If r_Bill.�����־ = 1 Then
              Zl_סԺ���ʼ�¼_Insert(v_No, n_Count, r_Bill.����id, r_Bill.��ҳid, r_Bill.��ʶ��, r_Bill.����, r_Bill.�Ա�, r_Bill.����,
                               r_Bill.����, r_Bill.�ѱ�, r_Bill.���˲���id, r_Bill.���˿���id, r_Item.�Ӱ��־, r_Bill.Ӥ����, r_Bill.�ⷿid,
                               ������Ա_In, Null, r_Item.�շ�ϸĿid, r_Item.�շ����, r_Item.���㵥λ, Null, Null, Null, 1, 1, Null,
                               r_Bill.�ⷿid, Null, r_Item.������Ŀid, r_Item.�վݷ�Ŀ, r_Item.�ּ�, r_Item.�ּ�, r_Item.�ּ�, Null,
                               Sysdate, Sysdate, Null, Null, v_Usercode, ������Ա_In);
            Else
              Zl_������ʼ�¼_Insert(v_No, n_Count, r_Bill.����id, r_Bill.��ʶ��, r_Bill.����, r_Bill.�Ա�, r_Bill.����, r_Bill.�ѱ�,
                               r_Item.�Ӱ��־, r_Bill.Ӥ����, r_Bill.���˿���id, r_Bill.�ⷿid, ������Ա_In, Null, r_Item.�շ�ϸĿid,
                               r_Item.�շ����, r_Item.���㵥λ, 1, 1, Null, r_Bill.�ⷿid, Null, r_Item.������Ŀid, r_Item.�վݷ�Ŀ,
                               r_Item.�ּ�, r_Item.�ּ�, r_Item.�ּ�, Sysdate, Sysdate, Null, Null, v_Usercode, ������Ա_In);
            End If;
          End Loop;
        End If;
      
        If n_Row = 0 Then
          Exit;
        End If;
      End Loop;
    End If;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Һ��ҩ��¼_��ҩ;
/

--139595:��ҵ��,2019-04-15,����֧�������������
Create Or Replace Procedure Zl_סԺ���ʼ�¼_��ҩ���
(
  Billid_In     In Varchar2, --ҩƷ�շ���¼ID��,��ʽ:"id1,id2,id3....."
  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
  ���ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type := Null
) As
  --���ܣ����һ��סԺ���ʻ��۵�
  --139595�޸�֧������������
  --������
  --    ���ʱ��_IN�����ڲ�����Ҫͳһ���ƻ򷵻�ʱ��ĵط�

  n_Billid Number;
  n_���   Number;
  Cursor c_Bill Is
    Select ID, NO, ��¼����, ��¼״̬, ����id, ��ҳid, �����־, ʵ�ս��, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ���˲���id, 1 As סԺ����
    From סԺ���ü�¼
    Where ID = (Select ����id From ҩƷ�շ���¼ Where ID = n_Billid) And ��¼״̬ = 0
    Union All
    Select ID, NO, ��¼����, ��¼״̬, ����id, ��ҳid, �����־, ʵ�ս��, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ���˲���id, 0 As סԺ����
    From ������ü�¼
    Where ID = (Select ����id From ҩƷ�շ���¼ Where ID = n_Billid) And ��¼״̬ = 0;

  v_Infotmp Varchar2(4000);
  v_Date    Date;
Begin
  If ���ʱ��_In Is Null Then
    Select Sysdate Into v_Date From Dual;
  Else
    v_Date := ���ʱ��_In;
  End If;

  v_Infotmp := Billid_In || ',';
  While v_Infotmp Is Not Null Loop
    --�ֽⵥ��ID��
    n_Billid  := Substr(v_Infotmp, 1, Instr(v_Infotmp, ',') - 1);
    v_Infotmp := Replace(',' || v_Infotmp, ',' || n_Billid || ',');
  
    For r_Bill In c_Bill Loop
      If r_Bill.סԺ���� = 1 Then
        Update סԺ���ü�¼
        Set ��¼״̬ = 1, ����Ա��� = ����Ա���_In, ����Ա���� = ����Ա����_In, �Ǽ�ʱ�� = v_Date --�Ѳ�����ҩƷ��¼��ʱ�䲻��
        Where ID = r_Bill.Id;
      Else
        Update ������ü�¼
        Set ��¼״̬ = 1, ����Ա��� = ����Ա���_In, ����Ա���� = ����Ա����_In, �Ǽ�ʱ�� = v_Date --�Ѳ�����ҩƷ��¼��ʱ�䲻��
        Where ID = r_Bill.Id;
      End If;
    
      --ҩƷ�շ���¼.��������
      Update ҩƷ�շ���¼
      Set �������� = Decode(Sign(Nvl(�������, v_Date) - v_Date), -1, ��������, v_Date)
      Where ID = n_Billid;
    
      If Nvl(r_Bill.�����־, 0) = 1 Or Nvl(r_Bill.�����־, 0) = 2 Then
        n_��� := r_Bill.�����־;
      Elsif Nvl(r_Bill.��ҳid, 0) = 0 Or Nvl(r_Bill.�����־, 0) = 4 Then
        n_��� := 1;
      Else
        n_��� := 2;
      End If;
    
      --�������
      Update �������
      Set ������� = Nvl(�������, 0) + Nvl(r_Bill.ʵ�ս��, 0)
      Where ����id = r_Bill.����id And ���� = 1 And ���� = n_���;
    
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, ����, �������, Ԥ�����)
        Values
          (r_Bill.����id, 1, n_���, r_Bill.ʵ�ս��, 0);
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
    
      --�ⷿ�е�ҩƷ��ȫ��������Ϊ���շ�
      If r_Bill.סԺ���� = 1 Then
        Update δ��ҩƷ��¼
        Set ���շ� = 1, �������� = v_Date
        Where NO = r_Bill.No And ���� In (9, 10) And Nvl(���շ�, 0) = 0 And
              Nvl(�ⷿid, 0) Not In
              (Select Distinct Nvl(ִ�в���id, 0)
               From סԺ���ü�¼
               Where ��¼���� = 2 And NO = r_Bill.No And �շ���� In ('5', '6', '7') And ��¼״̬ = 0);
      Else
        Update δ��ҩƷ��¼
        Set ���շ� = 1, �������� = v_Date
        Where NO = r_Bill.No And ���� In (9, 10) And Nvl(���շ�, 0) = 0 And
              Nvl(�ⷿid, 0) Not In
              (Select Distinct Nvl(ִ�в���id, 0)
               From ������ü�¼
               Where ��¼���� = 2 And NO = r_Bill.No And �շ���� In ('5', '6', '7') And ��¼״̬ = 0);
      End If;
    End Loop;
  End Loop;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_סԺ���ʼ�¼_��ҩ���;
/

--139595:��ҵ��,2019-04-15,����֧�������������
Create Or Replace Procedure Zl_��Һ��ҩ��¼_ȡ������(��ҩid_In In Varchar2 --ID��:ID1,ID2....
                                           ) Is
  v_Tansid   Varchar2(20);
  v_Tmp      Varchar2(4000);
  v_������Ա ��Һ��ҩ��¼.������Ա%Type;
  d_����ʱ�� ��Һ��ҩ��¼.����ʱ��%Type;
  n_���     Number(2);

  v_Error    Varchar2(255);
  n_����״̬ ��Һ��ҩ��¼.����״̬%Type;
  v_Usercode Varchar2(100);
  Err_Custom Exception;
Begin
  v_Usercode := Zl_Identity;
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ';') + 1);
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ',') + 1);
  v_Usercode := Substr(v_Usercode, 1, Instr(v_Usercode, ',') - 1);

  If ��ҩid_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := ��ҩid_In || ',';
  End If;
  While v_Tmp Is Not Null Loop
    --�ֽⵥ��ID��
    v_Tansid := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp    := Replace(',' || v_Tmp, ',' || v_Tansid || ',');
  
    --��鵱ǰ��Һ����״̬�Ƿ�Ϊ�ѷ���״̬
    Begin
      Select ����״̬ Into n_����״̬ From ��Һ��ҩ��¼ Where ID = v_Tansid;
    
      If n_����״̬ != 5 Then
        v_Error := '�������ѱ����������ܽ���ȡ�����Ͳ�����';
        Raise Err_Custom;
      End If;
    Exception
      When Others Then
        v_Error := '�������ѱ�������';
        Raise Err_Custom;
    End;
  
    Select ������Ա, ����ʱ��
    Into v_������Ա, d_����ʱ��
    From (Select ������Ա, ����ʱ�� From ��Һ��ҩ״̬ Where ��ҩid = v_Tansid And �������� = 4 Order By ����ʱ�� Desc)
    Where Rownum = 1;
  
    Update ��Һ��ҩ��¼ Set ����״̬ = 4, ������Ա = v_������Ա, ����ʱ�� = d_����ʱ�� Where ID = v_Tansid;
  
    --��[��Һ��ҩ״̬]���м�¼��ȡ�����͡��Ĳ���
    Insert Into ��Һ��ҩ״̬
      (��ҩid, ��������, ������Ա, ����ʱ��, ����˵��)
    Values
      (v_Tansid, 4, v_������Ա, Sysdate, 'ȡ������');
  
    Select �Ƿ��� Into n_��� From ��Һ��ҩ��¼ Where ID = v_Tansid;
    If n_��� <> 0 Then
      For r_Item In (Select a.No, b.���, 1 As סԺ����
                     From ��Һ��ҩ���� A, סԺ���ü�¼ B
                     Where a.����id = b.����id And a.No = b.No And b.��¼״̬ = 1 And a.��ҩid = v_Tansid
                     Union All
                     Select a.No, b.���, 0 As סԺ����
                     From ��Һ��ҩ���� A, ������ü�¼ B
                     Where a.����id = b.����id And a.No = b.No And b.��¼״̬ = 1 And a.��ҩid = v_Tansid) Loop
        If r_Item.No Is Not Null Then
          If r_Item.סԺ���� = 1 Then
            Zl_סԺ���ʼ�¼_Delete(r_Item.No, r_Item.���, v_Usercode, Zl_Username);
          Else
            Zl_������ʼ�¼_Delete(r_Item.No, r_Item.���, v_Usercode, Zl_Username);
          End If;
        End If;
      End Loop;
    End If;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Һ��ҩ��¼_ȡ������;
/

--139595:��ҵ��,2019-04-15,����֧�������������
Create Or Replace Procedure Zl_��Һ��ҩ��¼_ȡ����ҩ(��ҩid_In In varchar2 --ID��:ID1,ID2....
                                           ) Is
  v_Tansid   varchar2(20);
  v_Tmp      varchar2(4000);
  v_No       varchar2(20);
  v_Usercode varchar2(100);
  n_���     ��Һ��ҩ��¼.�Ƿ���%Type := 0;
  v_������Ա ��Һ��ҩ��¼.������Ա%Type;
  d_����ʱ�� ��Һ��ҩ��¼.����ʱ��%Type;

  v_Error    varchar2(255);
  n_����״̬ ��Һ��ҩ��¼.����״̬%Type;
  Err_Custom Exception;
  n_Row Number(10);
  n_Out Number(1);

Begin
  v_Usercode := Zl_Identity;
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ';') + 1);
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ',') + 1);
  v_Usercode := Substr(v_Usercode, 1, Instr(v_Usercode, ',') - 1);
  n_Out      := Nvl(zl_GetSysParameter('��Ժ���˲������÷�', 1345), 0);

  If ��ҩid_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := ��ҩid_In || ',';
  End If;

  While v_Tmp Is Not Null Loop
    --�ֽⵥ��ID�� 
    v_Tansid := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp    := Replace(',' || v_Tmp, ',' || v_Tansid || ',');
  
    --��鵱ǰ��Һ����״̬�Ƿ�Ϊ����ҩ״̬ 
    Begin
      Select ����״̬ Into n_����״̬ From ��Һ��ҩ��¼ Where ID = v_Tansid;
    
      If n_����״̬ != 4 Then
        v_Error := '�����ݵ�ǰ������ҩ״̬�����ܽ���ȡ����ҩ��';
        Raise Err_Custom;
      End If;
    Exception
      When Others Then
        v_Error := '�������ѱ�������';
        Raise Err_Custom;
    End;
  
    Select ������Ա, ����ʱ��
    Into v_������Ա, d_����ʱ��
    From (Select ������Ա, ����ʱ�� From ��Һ��ҩ״̬ Where ��ҩid = v_Tansid And �������� = 2 Order By ����ʱ�� Desc)
    Where Rownum = 1;
  
    Update ��Һ��ҩ��¼ Set ����״̬ = 2, ������Ա = v_������Ա, ����ʱ�� = d_����ʱ�� Where ID = v_Tansid;
  
    --��[��Һ��ҩ״̬]���м�¼��ȡ����ҩ���Ĳ���
    Insert Into ��Һ��ҩ״̬
      (��ҩid, ��������, ������Ա, ����ʱ��, ����˵��)
    Values
      (v_Tansid, 2, v_������Ա, Sysdate, 'ȡ����ҩ');
  
    Select �Ƿ��� Into n_��� From ��Һ��ҩ��¼ Where ID = v_Tansid;
    If n_��� <> 1 Then
      For r_Item In (Select a.No, b.���, 1 As סԺ����
                     From ��Һ��ҩ���� A, סԺ���ü�¼ B
                     Where a.����id = b.����id And a.No = b.No And b.��¼״̬ = 1 And a.��ҩid = v_Tansid
                     Union All
                     Select a.No, b.���, 0 As סԺ����
                     From ��Һ��ҩ���� A, ������ü�¼ B
                     Where a.����id = b.����id And a.No = b.No And b.��¼״̬ = 1 And a.��ҩid = v_Tansid) Loop
        If r_Item.No Is Not Null Then
          If r_Item.סԺ���� = 1 Then
             Zl_סԺ���ʼ�¼_Delete(r_Item.No, r_Item.���, v_Usercode, Zl_Username);
          Else
            Zl_������ʼ�¼_Delete(r_Item.No, r_Item.���, v_Usercode, Zl_Username);
          End If;
        End If;
      End Loop;
    Else
      Zl_��Һ��ҩ��¼_ȡ����ҩ(v_Tansid);
    End If;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Һ��ҩ��¼_ȡ����ҩ;
/

--139595:��ҵ��,2019-04-15,����֧�������������
Create Or Replace Procedure Zl_��Һ��ҩ��¼_ȡ����ҩ(��ҩid_In In Varchar2 --ID��:��ҩID1,��ҩID2....
                                           ) Is
  v_Id       Varchar2(20);
  v_��ҩid   Varchar2(20);
  v_Tmp      Varchar2(4000);
  v_Date     Date;
  v_������Ա ��Һ��ҩ��¼.������Ա%Type;
  d_����ʱ�� ��Һ��ҩ��¼.����ʱ��%Type;
  n_����״̬ ��Һ��ҩ��¼.����״̬%Type;

  v_ֹͣ��ҩids    Varchar2(4000); --���Ա�ҩ��ҩʱ�����˶Բ��ϣ���ȡ�������Һ���ġ�ȡ����ҩ������
  n_�Ա�ҩ����     ҩƷ�շ���¼.ʵ������%Type;
  n_�Ա�ҩ�������� ҩƷ�շ���¼.ʵ������%Type; --���Ա�ҩ��ҩƷ�շ���¼�п��Ա�����������
  n_����           Number; --1�����ﵥ�ݣ�2��סԺ����

  v_Error Varchar2(255);
  Err_Custom Exception;

  Cursor c_��ҩ���� Is
    Select /*+ rule*/
    Distinct c.��¼id, a.Id As ��ҩid, c.�շ�id, a.����, a.Ч��, a.����, c.���� As ��ҩ��, a.ҩƷid, a.����, b.����id
    From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B, ��Һ��ҩ���� C, Table(Cast(f_Num2list(��ҩid_In) As Zltools.t_Numlist)) D
    Where c.��¼id = d.Column_Value And b.Id = c.�շ�id And a.���� = b.���� And a.No = b.No And a.�ⷿid + 0 = b.�ⷿid And
          a.ҩƷid + 0 = b.ҩƷid And a.��� = b.��� And (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0)
    Order By a.ҩƷid, a.����;

  v_��ҩ���� c_��ҩ����%RowType;

  Cursor c_�Ա�ҩ��¼ Is
    Select Distinct a.Id, b.��������, c.����ϵ��, c.ҩƷid
    From ��Һ��ҩ��¼ A, ����ҽ����¼ B, ҩƷ��� C, Table(Cast(f_Num2list(��ҩid_In) As Zltools.t_Numlist)) D
    Where a.ҽ��id = b.���id And a.Id = d.Column_Value And b.�շ�ϸĿid = c.ҩƷid And b.ִ������ = 5 And b.ִ�б�� = 0 And
          b.�շ�ϸĿid In (Select d.ҩƷid From ��Һ�Ա�ҩ�嵥 D Where d.�Ƿ����� = 1);

  v_�Ա�ҩ��¼ c_�Ա�ҩ��¼%RowType;
Begin
  If ��ҩid_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := ��ҩid_In || ',';
  End If;
  While v_Tmp Is Not Null Loop
    --�ֽⵥ��ID��
    v_Id  := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp := Replace(',' || v_Tmp, ',' || v_Id || ',');
  
    --��鵱ǰ��Һ����״̬�Ƿ�Ϊ����ҩ״̬
    Begin
      Select ����״̬ Into n_����״̬ From ��Һ��ҩ��¼ Where ID = v_Id;
    
      If n_����״̬ != 2 Then
        v_Error := '�������ѱ����������ܽ���ȡ����ҩ������';
        Raise Err_Custom;
      End If;
    Exception
      When Others Then
        v_Error := '�������ѱ�������';
        Raise Err_Custom;
    End;
  
    Select ������Ա, ����ʱ��
    Into v_������Ա, d_����ʱ��
    From ��Һ��ҩ״̬
    Where ��ҩid = v_Id And �������� = 1 And Rownum = 1;
  
    Update ��Һ��ҩ��¼ Set ����״̬ = 1, ������Ա = v_������Ա, ����ʱ�� = d_����ʱ�� Where ID = v_Id;
  
    --��[��Һ��ҩ״̬]���м�¼��ȡ����ҩ���Ĳ���
    Insert Into ��Һ��ҩ״̬
      (��ҩid, ��������, ������Ա, ����ʱ��, ����˵��)
    Values
      (v_Id, 1, v_������Ա, Sysdate, 'ȡ����ҩ');
  
  End Loop;

  Select Sysdate Into v_Date From Dual;

  --������Һ�Ա�ҩ�嵥�����������ҩƷ���ѷ�ҩ����
  For v_�Ա�ҩ��¼ In c_�Ա�ҩ��¼ Loop
  
    n_�Ա�ҩ���� := v_�Ա�ҩ��¼.�������� / v_�Ա�ҩ��¼.����ϵ��;
  
    Select Sum(a.ʵ������)
    Into n_�Ա�ҩ��������
    From ҩƷ�շ���¼ A
    Where Mod(a.��¼״̬, 3) = 1 And a.�ƻ�id = v_�Ա�ҩ��¼.Id And a.ҩƷid = v_�Ա�ҩ��¼.ҩƷid And a.����� Is Not Null And
          a.������� Is Not Null;
  
    If n_�Ա�ҩ�������� < n_�Ա�ҩ���� Then
      --��������˶Բ��ϣ����ռ���ǰ��ҩid,��������ͬ���������Һ���Ķ�ӦҩƷ
      If v_ֹͣ��ҩids Is Null Then
        v_ֹͣ��ҩids := v_�Ա�ҩ��¼.Id;
      Else
        v_ֹͣ��ҩids := v_ֹͣ��ҩids || ',' || v_�Ա�ҩ��¼.Id;
      End If;
    
      Exit;
    
    End If;
  
    --����Һ����������Ա�ҩ,���ռ���ҩƷ�շ���¼���е�id
    For v_�Ա�ҩ�շ���¼ In (Select a.Id, a.����, a.Ч��, a.����, a.ʵ������ As ��ҩ��, a.����, a.����id
                      From ҩƷ�շ���¼ A
                      Where a.�ƻ�id = v_�Ա�ҩ��¼.Id And a.ҩƷid = v_�Ա�ҩ��¼.ҩƷid And a.����� Is Not Null And a.������� Is Not Null And
                            (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0)
                      Order By a.����) Loop
    
      --�ж�������������ﻹ��סԺ 
      Begin
        Select 1 Into n_���� From ������ü�¼ Where ID = v_�Ա�ҩ�շ���¼.����id;
      Exception
        When Others Then
          n_���� := 2;
      End;
    
      Zl_ҩƷ�շ���¼_������ҩ(v_�Ա�ҩ�շ���¼.Id, Zl_Username, v_Date, v_�Ա�ҩ�շ���¼.����, v_�Ա�ҩ�շ���¼.Ч��, v_�Ա�ҩ�շ���¼.����, v_�Ա�ҩ�շ���¼.��ҩ��, Null,
                     Zl_Username, 2, n_����);
    End Loop;
  End Loop;

  For v_��ҩ���� In c_��ҩ���� Loop
    --�ų����жϵ���Һ��
    If Instr(',' || v_ֹͣ��ҩids || ',', ',' || v_��ҩ����.��¼id || ',') = 0 Then
      Begin
        Select 1 Into n_���� From ������ü�¼ Where ID = v_��ҩ����.����id;
      Exception
        When Others Then
          n_���� := 2;
      End;
    
      --������ҩ
      Zl_ҩƷ�շ���¼_������ҩ(v_��ҩ����.��ҩid, Zl_Username, v_Date, v_��ҩ����.����, v_��ҩ����.Ч��, v_��ҩ����.����, v_��ҩ����.��ҩ��, Null, Zl_Username,
                     2, n_����);
    
      Select Max(a.Id)
      Into v_��ҩid
      From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B
      Where b.Id = v_��ҩ����.��ҩid And a.���� = b.���� And a.No = b.No And a.�ⷿid + 0 = b.�ⷿid And a.ҩƷid + 0 = b.ҩƷid And
            a.��� = b.��� And Mod(a.��¼״̬, 3) = 1 And a.������� Is Null;
    
      --�滻��Һ��ҩ�����е��շ�ID
      Update ��Һ��ҩ���� Set �շ�id = v_��ҩid Where ��¼id = v_��ҩ����.��¼id And �շ�id = v_��ҩ����.�շ�id;
    End If;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Һ��ҩ��¼_ȡ����ҩ;
/

--139595:��ҵ��,2019-04-15,����֧�������������
Create Or Replace Procedure Zl_��Һ��ҩ��¼_ȷ�Ͼܾ�
(
  ��ҩid_In   In ��Һ��ҩ��¼.Id%Type,
  ������Ա_In In ��Һ��ҩ��¼.������Ա%Type
) Is
  n_��ҩid   ��Һ��ҩ��¼.Id%Type;
  n_��ҩid   ҩƷ�շ���¼.Id%Type;
  n_��ҩid   ҩƷ�շ���¼.Id%Type;
  n_�շ�id   ҩƷ�շ���¼.Id%Type;
  n_����״̬ ��Һ��ҩ��¼.����״̬%Type;
  v_������Ա ��Һ��ҩ��¼.������Ա%Type;
  d_����ʱ�� ��Һ��ҩ��¼.����ʱ��%Type;
  v_Error    Varchar2(255);
  Err_Custom Exception;
  n_���� Number; --1�����ﵥ�ݣ�2��סԺ����

  Cursor c_��ҩ���� Is
    Select /*+ rule*/
    Distinct c.��¼id, a.Id As ��ҩid, c.�շ�id, a.����id
    From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B, ��Һ��ҩ���� C
    Where c.��¼id = ��ҩid_In And b.Id = c.�շ�id And a.���� = b.���� And a.No = b.No And a.�ⷿid + 0 = b.�ⷿid And
          a.ҩƷid + 0 = b.ҩƷid And a.��� = b.��� And (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0);

  r_��ҩ���� c_��ҩ����%RowType;

  Cursor c_��ҩ��¼ Is
    Select a.Id, a.����, a.Ч��, a.����, b.���� As ��ҩ��
    From ҩƷ�շ���¼ A, ��Һ��ҩ���� B
    Where a.Id = n_��ҩid And a.����� Is Not Null And b.�շ�id = n_�շ�id And b.��¼id = ��ҩid_In;

  r_��ҩ��¼ c_��ҩ��¼%RowType;

Begin
  Select Nvl(����״̬, 0) Into n_����״̬ From ��Һ��ҩ��¼ Where ID = ��ҩid_In;

  If n_����״̬ = 7 Then
    Insert Into ��Һ��ҩ״̬ (��ҩid, ��������, ������Ա, ����ʱ��) Values (��ҩid_In, 8, ������Ա_In, Sysdate);
  
    Select ������Ա, ����ʱ��
    Into v_������Ա, d_����ʱ��
    From ��Һ��ҩ״̬
    Where ��ҩid = ��ҩid_In And �������� = 1
    Order By ����ʱ�� Desc;
  
    Update ��Һ��ҩ��¼ Set ����״̬ = 1, ������Ա = v_������Ա, ����ʱ�� = d_����ʱ�� Where ID = ��ҩid_In;
  
    For r_��ҩ���� In c_��ҩ���� Loop
      n_��ҩid := r_��ҩ����.��ҩid;
      n_�շ�id := r_��ҩ����.�շ�id;
      For r_��ҩ��¼ In c_��ҩ��¼ Loop
        Begin
          Select 1 Into n_���� From ������ü�¼ Where ID = r_��ҩ����.����id;
        Exception
          When Others Then
            n_���� := 2;
        End;
      
        --������ҩ
        Zl_ҩƷ�շ���¼_������ҩ(r_��ҩ��¼.Id, Zl_Username, Sysdate, r_��ҩ��¼.����, r_��ҩ��¼.Ч��, r_��ҩ��¼.����, r_��ҩ��¼.��ҩ��, Null, Zl_Username);
      
        Select Max(a.Id)
        Into n_��ҩid
        From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B
        Where b.Id = n_��ҩid And a.���� = b.���� And a.No = b.No And a.�ⷿid + 0 = b.�ⷿid And a.ҩƷid + 0 = b.ҩƷid And
              a.��� = b.��� And Mod(a.��¼״̬, 3) = 1 And a.������� Is Null;
      
        --�滻��Һ��ҩ�����е��շ�ID
        Update ��Һ��ҩ���� Set �շ�id = n_��ҩid Where ��¼id = ��ҩid_In And �շ�id = r_��ҩ����.�շ�id;
      End Loop;
    End Loop;
  Else
    v_Error := '���������û�������Һ�������������ظ�������';
    Raise Err_Custom;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Һ��ҩ��¼_ȷ�Ͼܾ�;
/

--139595:��ҵ��,2019-04-15,����֧�������������
Create Or Replace Procedure Zl_��Һ��ҩ��¼_�������
(
  ��ҩid_In   In Varchar2, --ID��:ID1,��˱�־1,ID2,��˱�־2....
  ������Ա_In In ��Һ��ҩ��¼.������Ա%Type,
  ����ʱ��_In In ��Һ��ҩ��¼.����ʱ��%Type
) Is
  v_Tansid     Varchar2(20);
  v_Tansids    Varchar2(4000);
  v_Tmp        Varchar2(4000);
  v_Usercode   Varchar2(100);
  v_��ҩid     ҩƷ�շ���¼.Id%Type;
  n_Count      Number(1);
  d_���ʱ��   ҩƷ�շ���¼.�������%Type;
  v_No         ҩƷ�շ���¼.No%Type;
  v_�ϴ�no     ҩƷ�շ���¼.No%Type;
  n_��˱�־   Number(1);
  n_����״̬   Number(2);
  v_�շ�ids    Varchar2(4000);
  v_��ҩ����id ҩƷ�շ���¼.Id%Type;

  v_ԭʼid ҩƷ�շ���¼.Id%Type;
  v_Error  Varchar2(255);
  n_����   Number; --1�����ﵥ�ݣ�2��סԺ����
  Err_Custom Exception;

  Cursor c_���ʼ�¼ Is
    Select Distinct a.����id, b.����ʱ��
    From ҩƷ�շ���¼ A, ��Һ��ҩ��¼ B, ��Һ��ҩ���� C
    Where a.Id = c.�շ�id And b.Id = c.��¼id And b.Id = v_Tansid And b.����״̬ = 9;

  v_���ʼ�¼ c_���ʼ�¼%RowType;

  Cursor c_��ҩ��¼ Is
    Select /*+ rule*/
    Distinct a.Id As ��ҩid, c.�շ�id, c.����, a.ҩƷid, a.����, c.��¼id As ��ҩid, a.����id
    From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B, ��Һ��ҩ���� C, Table(Cast(f_Num2list(v_Tansids) As Zltools.t_Numlist)) D
    Where c.��¼id = d.Column_Value And b.Id = c.�շ�id And a.���� = b.���� And a.No = b.No And a.�ⷿid + 0 = b.�ⷿid And
          a.ҩƷid + 0 = b.ҩƷid And a.��� = b.��� And (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0)
    Order By a.ҩƷid, a.����;

  v_��ҩ��¼ c_��ҩ��¼%RowType;

  Cursor c_�������� Is
    Select /*+ rule*/
     a.No, a.��� || ':' || c.���� || ':' || c.��¼id As �������, 1 As סԺ����
    From סԺ���ü�¼ A, ҩƷ�շ���¼ B, ��Һ��ҩ���� C, Table(Cast(f_Num2list(v_Tansids) As Zltools.t_Numlist)) D
    Where a.Id = b.����id And b.Id = c.�շ�id And Mod(b.��¼״̬, 3) = 1 And c.��¼id = d.Column_Value
    Union All
    Select /*+ rule*/
     a.No, a.��� || ':' || c.���� || ':' || c.��¼id As �������, 0 As סԺ����
    From ������ü�¼ A, ҩƷ�շ���¼ B, ��Һ��ҩ���� C, Table(Cast(f_Num2list(v_Tansids) As Zltools.t_Numlist)) D
    Where a.Id = b.����id And b.Id = c.�շ�id And Mod(b.��¼״̬, 3) = 1 And c.��¼id = d.Column_Value;

  v_�������� c_��������%RowType;

  Cursor c_�Ա�ҩ��¼ Is
    Select Distinct a.Id, b.��������, c.����ϵ��, c.ҩƷid
    From ��Һ��ҩ��¼ A, ����ҽ����¼ B, ҩƷ��� C, Table(Cast(f_Num2list(v_Tansids) As Zltools.t_Numlist)) D
    Where a.ҽ��id = b.���id And a.Id = d.Column_Value And b.�շ�ϸĿid = c.ҩƷid And b.ִ������ = 5 And b.ִ�б�� = 0 And
          b.�շ�ϸĿid In (Select d.ҩƷid From ��Һ�Ա�ҩ�嵥 D Where d.�Ƿ����� = 1);

  v_�Ա�ҩ��¼ c_�Ա�ҩ��¼%RowType;
Begin
  If ��ҩid_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := ��ҩid_In || ',';
  End If;

  v_Usercode := Zl_Identity;
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ';') + 1);
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ',') + 1);
  v_Usercode := Substr(v_Usercode, 1, Instr(v_Usercode, ',') - 1);

  While v_Tmp Is Not Null Loop
    --�ֽⵥ��ID��
    v_Tansid   := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp      := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
    n_��˱�־ := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp      := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
  
    v_�շ�ids := Null;
  
    --ͳ�����ȷ�ϵ���Һ��(n_��˱�־ = 1)
    If n_��˱�־ = 1 Then
      If v_Tansids Is Null Then
        v_Tansids := v_Tansid;
      Else
        v_Tansids := v_Tansids || ',' || v_Tansid;
      End If;
    End If;
  
    --��鵱ǰ��Һ����״̬�Ƿ�Ϊ����ҩ״̬
    Begin
      Select ����״̬ Into n_����״̬ From ��Һ��ҩ��¼ Where ID = v_Tansid;
    
      If n_����״̬ <> 9 Then
        v_Error := '�������ѱ����������ܽ���������ˣ�';
        Raise Err_Custom;
      End If;
    Exception
      When Others Then
        v_Error := '�������ѱ�������';
        Raise Err_Custom;
    End;
  
    If n_��˱�־ = 1 Then
      n_����״̬ := 10;
    Elsif n_��˱�־ = 2 Then
      n_����״̬ := 11;
    End If;
  
    --������Һ����Ӧ���շ�NO
    Begin
      Select NO
      Into v_No
      From ҩƷ�շ���¼
      Where ID In (Select �շ�id From ��Һ��ҩ���� Where ��¼id In (Select ID From ��Һ��ҩ��¼ Where ID = v_Tansid)) And Rownum < 2;
    Exception
      When Others Then
        v_No := '';
    End;
  
    --�շ�NO��ͬ����ҩID�����ʱ���Դ�����Ϊ�ӳ�1��
    If v_No = v_�ϴ�no Then
      d_���ʱ�� := d_���ʱ�� + 1 / 24 / 60 / 60;
    Else
      d_���ʱ�� := ����ʱ��_In;
      v_�ϴ�no   := v_No;
    End If;
  
    --���ʼ�¼����
    For v_���ʼ�¼ In c_���ʼ�¼ Loop
      Zl_���˷�������_Audit(v_���ʼ�¼.����id, v_���ʼ�¼.����ʱ��, ������Ա_In, d_���ʱ��, n_��˱�־);
    End Loop;
  
    Select Count(*) Into n_Count From ��Һ��ҩ״̬ Where ��ҩid = v_Tansid And ����ʱ�� = ����ʱ��_In;
  
    If n_Count <> 1 Then
      Insert Into ��Һ��ҩ״̬
        (��ҩid, ��������, ������Ա, ����ʱ��)
      Values
        (v_Tansid, n_����״̬, ������Ա_In, ����ʱ��_In);
    End If;
    Update ��Һ��ҩ��¼ Set ������Ա = ������Ա_In, ����ʱ�� = ����ʱ��_In, ����״̬ = n_����״̬ Where ID = v_Tansid;
  End Loop;

  --����ҩ
  For v_��ҩ��¼ In c_��ҩ��¼ Loop
    --�ж�������������ﻹ��סԺ 
    Begin
      Select 1 Into n_���� From ������ü�¼ Where ID = v_��ҩ��¼.����id;
    Exception
      When Others Then
        n_���� := 2;
    End;
  
    Zl_ҩƷ�շ���¼_������ҩ(v_��ҩ��¼.��ҩid, ������Ա_In, ����ʱ��_In, Null, Null, Null, v_��ҩ��¼.����, Null, ������Ա_In, 2, n_����);
  
    --ȡ��ҩ����id
    Select a.Id
    Into v_��ҩid
    From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B
    Where b.Id = v_��ҩ��¼.��ҩid And a.���� = b.���� And a.No = b.No And a.�ⷿid + 0 = b.�ⷿid And a.ҩƷid + 0 = b.ҩƷid And
          a.��� = b.��� And Mod(a.��¼״̬, 3) = 1 And a.������� Is Null;
  
    --��Һ��ҩ�����е��շ�ID����Ϊ��ҩ�������շ�ID
    Update ��Һ��ҩ���� Set �շ�id = v_��ҩid Where ��¼id = v_��ҩ��¼.��ҩid And �շ�id = v_��ҩ��¼.�շ�id;
  
    If v_�շ�ids Is Null Then
      v_�շ�ids := v_��ҩid;
    Else
      v_�շ�ids := v_�շ�ids || ',' || v_��ҩid;
    End If;
  
    --ȡԭʼid
    Select a.Id
    Into v_ԭʼid
    From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B
    Where b.Id = v_��ҩ��¼.��ҩid And a.���� = b.���� And a.No = b.No And a.�ⷿid + 0 = b.�ⷿid And a.ҩƷid + 0 = b.ҩƷid And
          a.��� = b.��� And Mod(a.��¼״̬, 3) = 0 And a.������� Is Not Null;
  
    Insert Into ��Һ��ҩ����
      (��¼id, �շ�id, ����)
      Select ��¼id, v_ԭʼid, ���� From ��Һ��ҩ���� Where ��¼id = v_��ҩ��¼.��ҩid And �շ�id = v_��ҩid;
  
    v_�շ�ids := v_�շ�ids || ',' || v_ԭʼid;
  End Loop;

  --��������
  For v_�������� In c_�������� Loop
    If v_��������.סԺ���� = 1 Then
      Zl_סԺ���ʼ�¼_Delete(v_��������.No, v_��������.�������, v_Usercode, Zl_Username, 2, 1, 1, d_���ʱ��);
    Else
      Zl_������ʼ�¼_Delete(v_��������.No, v_��������.�������, v_Usercode, Zl_Username);
    End If;
  End Loop;

  --������Һ�Ա�ҩ�嵥�����������ҩƷ���ѷ�ҩ����
  For v_�Ա�ҩ��¼ In c_�Ա�ҩ��¼ Loop
    --����Һ����������Ա�ҩ,���ռ���ҩƷ�շ���¼���е�id
    For v_�Ա�ҩ�շ���¼ In (Select a.Id, a.����, a.Ч��, a.����, a.ʵ������ As ��ҩ��, a.����, a.����id
                      From ҩƷ�շ���¼ A
                      Where a.�ƻ�id = v_�Ա�ҩ��¼.Id And a.ҩƷid = v_�Ա�ҩ��¼.ҩƷid And a.����� Is Not Null And a.������� Is Not Null And
                            (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0)
                      Order By a.����) Loop
    
      --�ж�������������ﻹ��סԺ 
      Begin
        Select 1 Into n_���� From ������ü�¼ Where ID = v_�Ա�ҩ�շ���¼.����id;
      Exception
        When Others Then
          n_���� := 2;
      End;
    
      Zl_ҩƷ�շ���¼_������ҩ(v_�Ա�ҩ�շ���¼.Id, ������Ա_In, ����ʱ��_In, Null, Null, Null, v_�Ա�ҩ�շ���¼.��ҩ��, Null, ������Ա_In, 2, n_����);
    End Loop;
  End Loop;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Һ��ҩ��¼_�������;
/

--139595:��ҵ��,2019-04-15,����֧�������������
Create Or Replace Procedure Zl_ҩƷ�շ���¼_������ҩ
(
  Billinfo_In   In Varchar2, --��ʽ:"id1,����1|id2,����2|....."
  Partid_In     In ҩƷ�շ���¼.�ⷿid%Type,
  People_In     In ҩƷ�շ���¼.�����%Type,
  Date_In       In ҩƷ�շ���¼.�������%Type,
  ��ҩ��ʽ_In   In ҩƷ�շ���¼.��ҩ��ʽ%Type := 3,
  ��ҩ��_In     In ҩƷ�շ���¼.������%Type := Null,
  ���ܷ�ҩ��_In In ҩƷ�շ���¼.���ܷ�ҩ��%Type := Null,
  Intdigit_In   In Number := 2,
  ��ҩ��_In     In ҩƷ�շ���¼.��ҩ��%Type := Null,
  �˲���_In     In ҩƷ�շ���¼.�˲���%Type := Null
) Is
  --ֻ������
  v_Infotmp     Varchar2(4000);
  v_Fields      Varchar2(4000);
  n_Billid      ҩƷ�շ���¼.Id%Type;
  n_����        ҩƷ�շ���¼.����%Type;
  Lng������id Number(18);
  Int���ϵ��   Number;
  Intִ��״̬   Number;
  Int����       ҩƷ�շ���¼.����%Type;
  Strno         ҩƷ�շ���¼.No%Type;
  Lng�ⷿid     ҩƷ�շ���¼.�ⷿid%Type;
  LngҩƷid     ҩƷ�շ���¼.ҩƷid%Type;
  Lng����id     ҩƷ�շ���¼.����id%Type;
  Dbl�����     Number;
  v_���ۼ�      ҩƷ�շ���¼.���ۼ�%Type;
  Intδ����     δ��ҩƷ��¼.δ����%Type;
  v_�˲�����    ҩƷ�շ���¼.�˲�����%Type;
  --��д����
  Dblʵ������ ҩƷ�շ���¼.ʵ������%Type;
  Dblʵ�ʽ�� ҩƷ�շ���¼.���۽��%Type;
  Dbl�ɱ���� ҩƷ�շ���¼.�ɱ����%Type;
  Dblʵ�ʲ�� ҩƷ�շ���¼.���%Type;
  --2002-07-31����
  --LNGLAST���� ��ҩǰȷ��������(�Ѽ���������)
  Strҩ��           Varchar2(200);
  Dbl��������       ҩƷ�շ���¼.��д����%Type;
  Lnglast����       ҩƷ�շ���¼.����%Type;
  Lngcur����        ҩƷ�շ���¼.����%Type;
  Str����           ҩƷ�շ���¼.����%Type;
  StrЧ��           ҩƷ�շ���¼.Ч��%Type;
  n_�ϴι�Ӧ��id    ҩƷ���.�ϴι�Ӧ��id%Type;
  n_�ϴβɹ���      ҩƷ���.�ϴβɹ���%Type;
  v_�ϴβ���        ҩƷ���.�ϴβ���%Type;
  d_�ϴ���������    ҩƷ���.�ϴ���������%Type;
  v_��׼�ĺ�        ҩƷ���.��׼�ĺ�%Type;
  n_��¼״̬        ҩƷ�շ���¼.��¼״̬%Type;
  n_ƽ���ɱ���      ҩƷ���.ƽ���ɱ���%Type;
  n_��ҩ��ʽ        ҩƷ�շ���¼.��ҩ��ʽ%Type;
  v_ժҪ            ҩƷ�շ���¼.ժҪ%Type;
  Bln�շ��뷢ҩ���� Number(1);
  v_Error           Varchar2(255);
  Err_Custom Exception;
  n_��ͨ�ۼ�С��   Number;
  n_��ͨ���С��   Number;
  n_���۹���ģʽ Number(1);
  n_ҩƷ���۹��� Number(1);
  n_��������       δ��ҩƷ��¼.��������%Type;
  n_סԺ����       Number(1);
Begin
  Select Sysdate Into v_�˲����� From Dual;
  If Billinfo_In Is Null Then
    v_Infotmp := Null;
  Else
    v_Infotmp := Billinfo_In || '|';
  End If;
  While v_Infotmp Is Not Null Loop
    --�ֽⵥ��ID��
    v_Fields  := Substr(v_Infotmp, 1, Instr(v_Infotmp, '|') - 1);
    n_Billid  := Substr(v_Fields, 1, Instr(v_Fields, ',') - 1);
    n_����    := Substr(v_Fields, Instr(v_Fields, ',') + 1);
    v_Infotmp := Replace('|' || v_Infotmp, '|' || v_Fields || '|');
  
    --��ȡ���շ���¼�ĵ��ݡ�ҩƷID���ⷿID,���۽�ʵ��������������ID
    Begin
      Select a.����, a.No, a.ҩƷid, a.�ⷿid, a.����id, Nvl(a.���ۼ�, 0), Nvl(a.���۽��, 0), Nvl(a.ʵ������, 0) * Nvl(a.����, 1), a.������id,
             a.���ϵ��, Nvl(a.����, 0), a.����, a.Ч��, a.��ҩ��λid, a.����, a.��������, a.��׼�ĺ�, Nvl(a.��ҩ��ʽ, 0), a.ժҪ, ��¼״̬
      Into Int����, Strno, LngҩƷid, Lng�ⷿid, Lng����id, v_���ۼ�, Dblʵ�ʽ��, Dblʵ������, Lng������id, Int���ϵ��, Lnglast����, Str����, StrЧ��,
           n_�ϴι�Ӧ��id, v_�ϴβ���, d_�ϴ���������, v_��׼�ĺ�, n_��ҩ��ʽ, v_ժҪ, n_��¼״̬
      From ҩƷ�շ���¼ A
      Where a.Id = n_Billid And a.������� Is Null
      For Update Nowait;
    
      Select '[' || c.���� || ']' || c.���� Into Strҩ�� From �շ���ĿĿ¼ C Where c.Id = LngҩƷid;
    Exception
      When Others Then
        Int���� := 0;
        v_Error := '���������û���ִ�з�ҩ�������ظ�������';
        Raise Err_Custom;
    End;
  
    --ȡ��ͨҵ�񾫶�λ��
    --���:1-ҩƷ 2-����
    --���ݣ�2-���ۼ� 4-���
    --��λ��ҩƷ:1-�ۼ� 5-��λ
    Begin
      Select ���� Into n_��ͨ���С�� From ҩƷ���ľ��� Where ��� = 1 And ���� = 4 And ��λ = 5;
    Exception
      When Others Then
        n_��ͨ���С�� := 2;
    End;
  
    Begin
      Select ���� Into n_��ͨ�ۼ�С�� From ҩƷ���ľ��� Where ��� = 1 And ���� = 2 And ��λ = 1;
    Exception
      When Others Then
        n_��ͨ�ۼ�С�� := 2;
    End;
  
    Select Zl_To_Number(Nvl(zl_GetSysParameter(275), '0')) Into n_���۹���ģʽ From Dual;
  
    If n_��ҩ��ʽ = -1 Or v_ժҪ = '�ܷ�' Then
      Int���� := 0;
    End If;
  
    If Int���� > 0 Then
      If Nvl(n_����, 0) = 0 Then
        Lngcur���� := Lnglast����;
      Else
        Lngcur���� := Nvl(n_����, 0);
      End If;
    
      --����Ƿ��Ѿ���д�ⷿ
      Bln�շ��뷢ҩ���� := 0;
      If Lng�ⷿid Is Null Then
        Bln�շ��뷢ҩ���� := 1;
      End If;
      Lng�ⷿid := Partid_In;
    
      --ȡ����ҩƷ������
      Begin
        Select �ϴ�����, Ч��, Nvl(��������, 0), �ϴι�Ӧ��id, �ϴβ���, �ϴ���������, ��׼�ĺ�, �ϴβɹ���
        Into Str����, StrЧ��, Dbl��������, n_�ϴι�Ӧ��id, v_�ϴβ���, d_�ϴ���������, v_��׼�ĺ�, n_�ϴβɹ���
        From ҩƷ���
        Where �ⷿid + 0 = Lng�ⷿid And ҩƷid = LngҩƷid And ���� = 1 And Nvl(����, 0) = Lngcur����;
      Exception
        When Others Then
          n_�ϴβɹ��� := 0;
          Dbl��������  := 0;
      End;
    
      --���������������˳�
      If Lngcur���� <> Nvl(Lnglast����, 0) Then
        If Dbl�������� < Dblʵ������ And Lngcur���� <> 0 Then
          v_Error := Strҩ�� || '�Ŀ����������㣬������ֹ��';
          Raise Err_Custom;
        End If;
      End If;
    
      If n_���۹���ģʽ <> 0 Then
        Select Nvl(�Ƿ����۹���, 0) Into n_ҩƷ���۹��� From ҩƷ��� Where ҩƷid = LngҩƷid;
      End If;
    
      If n_��¼״̬ = 1 Then
        --ԭʼ��ҩ��¼��ȡ���¼۸�
        n_ƽ���ɱ��� := Zl_Fun_Getoutcost(LngҩƷid, Lngcur����, Lng�ⷿid);
      
        If n_���۹���ģʽ <> 0 And n_ҩƷ���۹��� = 1 And (v_���ۼ� = n_ƽ���ɱ��� Or Round(v_���ۼ�, n_��ͨ�ۼ�С��) = Round(n_ƽ���ɱ���, n_��ͨ�ۼ�С��)) Then
          Dbl�ɱ���� := Dblʵ�ʽ��;
        Else
          Dbl�ɱ���� := Round(n_ƽ���ɱ��� * Nvl(Dblʵ������, 0), n_��ͨ���С��);
        End If;
      Else
        --��ҩ�ٷ���¼��ȡԭʼ���ݼ۸�
        Select a.�ɱ���
        Into n_ƽ���ɱ���
        From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B
        Where b.Id = n_Billid And a.���� = b.���� And a.No = b.No And a.�ⷿid = b.�ⷿid And a.ҩƷid + 0 = b.ҩƷid And
              a.��� = b.��� And Nvl(a.����, 0) = Nvl(b.����, 0) And (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0);
      
        Dbl�ɱ���� := Round(n_ƽ���ɱ��� * Nvl(Dblʵ������, 0), n_��ͨ���С��);
      End If;
    
      Dblʵ�ʲ�� := Round(Dblʵ�ʽ�� - Dbl�ɱ����, n_��ͨ���С��);
    
      --����ҩƷ�շ���¼�����۽��ɱ������
      Update ҩƷ�շ���¼
      Set �ⷿid = Lng�ⷿid, �ɱ��� = n_ƽ���ɱ���, �ɱ���� = Dbl�ɱ����, ��� = Dblʵ�ʲ��, ���� = Lngcur����, ���� = Str����, Ч�� = StrЧ��,
          ��ҩ�� = ��ҩ��_In, �˲��� = �˲���_In, �˲����� = v_�˲�����, ����� = People_In, ������� = Date_In, ��ҩ��ʽ = ��ҩ��ʽ_In, ������ = ��ҩ��_In,
          ���ܷ�ҩ�� = ���ܷ�ҩ��_In, ��ҩ��λid = n_�ϴι�Ӧ��id, ���� = v_�ϴβ���, �������� = d_�ϴ���������, ��׼�ĺ� = v_��׼�ĺ�
      Where ID = n_Billid;
      --�����������
      If Sql%RowCount = 0 Then
        v_Error := 'Ҫ��ҩ��ҩƷ��¼"' || Strҩ�� || '"�����ڣ�������ֹ��';
        Raise Err_Custom;
      End If;
    
      --����סԺ���ü�¼��ִ��״̬(��ִ��)
      Select Decode(Sum(Nvl(����, 1) * ʵ������), Null, 1, 0, 1, 2)
      Into Intִ��״̬
      From ҩƷ�շ���¼
      Where ���� = Int���� And NO = Strno And ����id = Lng����id And ����� Is Null;
    
      --�ж����������������û���סԺ����
      Begin
        Select 1 Into n_סԺ���� From ������ü�¼ Where ID = Lng����id;
      Exception
        When Others Then
          n_סԺ���� := 2;
      End;
    
      If n_סԺ���� = 2 Then
        Update סԺ���ü�¼
        Set ִ��״̬ = Intִ��״̬, ִ���� = Decode(People_In, Null, Zl_Username, People_In), ִ��ʱ�� = Date_In, ִ�в���id = Partid_In
        Where ID = Lng����id;
      Else
        Update ������ü�¼
        Set ִ��״̬ = Intִ��״̬, ִ���� = Decode(People_In, Null, Zl_Username, People_In), ִ��ʱ�� = Date_In, ִ�в���id = Partid_In
        Where ID = Lng����id;
      End If;
    
      --����δ��ҩƷ��¼(���δ����Ϊ����ɾ��)
      Select Count(*)
      Into Intδ����
      From ҩƷ�շ���¼
      Where ���� = Int���� And (�ⷿid + 0 = Lng�ⷿid Or �ⷿid Is Null) And NO = Strno And ����� Is Null And
            Nvl(LTrim(RTrim(ժҪ)), 'С��') <> '�ܷ�';
    
      If Intδ���� = 0 Then
        Delete δ��ҩƷ��¼
        Where NO = Strno And ���� = Int���� And (�ⷿid + 0 = Lng�ⷿid Or �ⷿid Is Null)
        Returning �������� Into n_��������;
      
        --���´������ͣ������Ŵ�������
        Update ҩƷ�շ���¼ Set ע��֤�� = n_�������� Where ���� = Int���� And NO = Strno And �ⷿid = Lng�ⷿid;
      End If;
    
      --���¿��
      Zl_ҩƷ���_Update(n_Billid, 2, 1);
    
      Zl_δ��ҩƷ��¼_Delete(n_Billid);
    
      --�����������
      Zl_ҩƷ�շ���¼_��������(n_Billid);
    
      b_Message.Zlhis_Drug_005(Lng�ⷿid, n_Billid);
    End If;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�շ���¼_������ҩ;
/



------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0059' Where ���=&n_System;
Commit;
