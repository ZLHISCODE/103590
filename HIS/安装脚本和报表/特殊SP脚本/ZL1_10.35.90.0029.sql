----------------------------------------------------------------------------------------------------------------
--���ű�֧�ִ�ZLHIS+ v10.35.90������ v10.35.90
--�������ݿռ�������ߵ�¼PLSQL��ִ�����нű�
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------



------------------------------------------------------------------------------
--�ṹ��������
------------------------------------------------------------------------------
--129973:����,2018-09-10,�������ԡ��Ƿ�����������
Alter Table ҩƷ��� Add �Ƿ��������� number(1);

--130300:����,2018-09-11,�������ԡ��ϸ�����÷�������
Alter Table ҩƷ��� Add �ϸ�����÷����� number(1);
Alter Table ҩƷ���� Add �ϸ�����÷����� number(1);

------------------------------------------------------------------------------
--������������
------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--128798:��͢��,2018-09-14,Zl_����ҽ��ִ��_Insert�����ύ����ȷ
Create Or Replace Procedure Zl_����ҽ��ִ��_Insert
(
  ҽ��id_In       In ����ҽ��ִ��.ҽ��id%Type,
  ���ͺ�_In       In ����ҽ��ִ��.���ͺ�%Type,
  Ҫ��ʱ��_In     In ����ҽ��ִ��.Ҫ��ʱ��%Type,
  ��������_In     In ����ҽ��ִ��.��������%Type,
  ִ��ժҪ_In     In ����ҽ��ִ��.ִ��ժҪ%Type,
  ִ����_In       In ����ҽ��ִ��.ִ����%Type,
  ִ��ʱ��_In     In ����ҽ��ִ��.ִ��ʱ��%Type,
  ����ִ��_In     In Number := 0,
  �Զ����_In     In Number := 0,
  ִ�н��_In     In ����ҽ��ִ��.ִ�н��%Type := 1,
  δִ��ԭ��_In   In ����ҽ��ִ��.˵��%Type := Null,
  ����Ա���_In   In ��Ա��.���%Type := Null,
  ����Ա����_In   In ��Ա��.����%Type := Null,
  ִ�в���id_In   In ������ü�¼.ִ�в���id%Type := 0,
  ��Һ���_In     In Number := 0,
  ������Ŀ����_In In Number := 0,
  ��Һͨ��_In     In ����ҽ��ִ��.��Һͨ��%Type := Null,
  ��¼��Դ_In     In ����ҽ��ִ��.��¼��Դ%Type := Null,
  ִ�з�ʽ_In     In ����ҽ��ִ��.ִ�з�ʽ%Type := Null
  --������ҽ��ID_IN=����ִ�е�ҽ��ID���������Ϊ��ʾ�ļ�����Ŀ��ID��
  --      ִ�н��_In=1- ���   =0  -δִ��
  --      �����̨ʽ������ ����Ա���_In ����Ա����_In �������������봫��
  --ִ�в���id_In=������ָ��ִ�в��ŵķ��ã���������0ʱ������ִ�в���
  --��Һ���_In=�ƶ�����վ����ʱ���Ƿ�����Һ��Ϣ��
  --������Ŀ����_In=����Ǽ�����Ŀʱ����Ҫ���ʵ������ҽ������״̬
) Is
  --����Ҫִ�е�����¼,�������˸�������,��鲿λ�ļ�¼
  --��������,��ҩ�巨,�ɼ�������������,�������ֻ��д�ڵ�һ����Ŀ�ϣ���ִ��״̬��ͬ
  v_��id       ����ҽ����¼.Id%Type;
  v_�������   ����ҽ����¼.�������%Type;
  v_�Զ����   Number;
  v_������Դ   ����ҽ����¼.������Դ%Type;
  v_��������   ����ҽ������.��¼����%Type;
  v_��������   ������ĿĿ¼.��������%Type;
  n_ִ�з���   ������ĿĿ¼.ִ�з���%Type;
  v_����id     ������ҳ.��ǰ����id%Type;
  v_��Һ����   Varchar2(200);
  v_Count      Number;
  v_Temp       Varchar2(255);
  v_��Ա���   ��Ա��.���%Type;
  v_��Ա����   ��Ա��.����%Type;
  n_��Ч       ����ҽ����¼.ҽ����Ч%Type;
  n_������Ŀid ����ҽ����¼.������Ŀid%Type;
  v_����ִ��   Varchar2(5);
  n_��������   ����ҽ��ִ��.��������%Type;
  n_����id     ����ҽ����¼.����id%Type;
  n_��ҳid     ����ҽ����¼.��ҳid%Type;
  v_�Һŵ�     ����ҽ����¼.�Һŵ�%Type;

  n_ִ�д���   Number;
  n_ʣ�����   Number;
  n_ִ��״̬   Number;
  d_��ֹʱ��   Date;
  d_��ʼʱ��   Date;
  n_��������   Number;
  n_�Ǽ�����   Number;
  n_��������   Number;
  d_Ҫ��ʱ��   Date;
  n_ִ�п���id Number;
  n_��Ѫҽ��   Number(1);
  d_ִ��ʱ��   Date;

  v_Date  Date;
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --������죬��ֹ��������ִ�м�¼
  Begin
    Select (a.�������� - c.�ǼǴ���) As ʣ������, a.��������, a.ִ�в���id, Nvl(d.������Ŀid, 0), c.�ǼǴ���
    Into v_Count, n_��������, n_ִ�п���id, n_������Ŀid, n_�Ǽ�����
    From ����ҽ������ a,
         (Select ҽ��id_In As ҽ��id, ���ͺ�_In As ���ͺ�, Nvl(Sum(b.��������), 0) As �ǼǴ���
           From ����ҽ��ִ�� b
           Where b.ҽ��id = ҽ��id_In And b.���ͺ� = ���ͺ�_In) c, ����ҽ����¼ d
    Where a.ҽ��id = c.ҽ��id And a.���ͺ� = c.���ͺ� And a.ҽ��id = ҽ��id_In And a.ҽ��id = d.Id And a.���ͺ� = ���ͺ�_In;
  Exception
    When Others Then
      v_Count := ��������_In;
  End;
  v_����ִ�� := Zl_Getsysparameter(288);
  n_�������� := ��������_In;
  If ��������_In > v_Count And (Not (n_������Ŀid = 0 And v_����ִ�� = 1)) Then
    If Round(n_�Ǽ����� + ��������_In) = 1 Then
      --��������Ѫִ��
      n_�������� := 1 - n_�Ǽ�����;
    Else
      v_Error := '���ڲ������������Ѿ������˵Ǽǣ���ˢ�º����ԡ�';
      Raise Err_Custom;
    End If;
  End If;
  --��ǰ������Ա
  If ����Ա���_In Is Not Null And ����Ա����_In Is Not Null Then
    v_��Ա��� := ����Ա���_In;
    v_��Ա���� := ����Ա����_In;
  Else
    Begin
      Select ����, ��� Into v_��Ա����, v_��Ա��� From ��Ա�� Where ���� = ִ����_In;
    Exception
      When Others Then
        v_Temp     := Zl_Identity;
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
        v_��Ա��� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
        v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    End;
  End If;
  --��ҽ����ֹʱ����м��
  Select a.ִ����ֹʱ��, a.��ʼִ��ʱ��, a.ҽ����Ч, a.����id, a.��ҳid, a.�Һŵ�
  Into d_��ֹʱ��, d_��ʼʱ��, n_��Ч, n_����id, n_��ҳid, v_�Һŵ�
  From ����ҽ����¼ a
  Where a.Id = ҽ��id_In;
  If Not d_��ֹʱ�� Is Null And n_��Ч = 0 Then
    If Ҫ��ʱ��_In > d_��ֹʱ�� Then
      v_Error := 'Ҫ��ʱ�䳬����ҽ����ֹʱ�䣬��ȷ��ҽ���Ƿ���ǰֹͣ��';
      Raise Err_Custom;
    End If;
  End If;
  If Not d_��ʼʱ�� Is Null Then
    If ִ��ʱ��_In < d_��ʼʱ�� Then
      v_Error := 'ִ��ʱ��������ҽ���Ŀ�ʼִ��ʱ��''' || To_Char(d_��ʼʱ��, 'yyyy-mm-dd HH24:mi:ss') || '''��';
      Raise Err_Custom;
    End If;
  End If;
  Select Sysdate Into v_Date From Dual;
  Select a.������Դ, ִ�п���id, Nvl(a.���id, a.Id), Nvl(a.�������, '*'), Nvl(b.��������, '0') ��������, Nvl(b.ִ�з���, 0) ִ�з���
  Into v_������Դ, v_����id, v_��id, v_�������, v_��������, n_ִ�з���
  From ����ҽ����¼ a, ������ĿĿ¼ b
  Where a.Id = ҽ��id_In And a.������Ŀid = b.Id(+);

  If v_������Դ = 2 Then
    Select Decode(��¼����, 1, 1, Decode(�������, 1, 1, 2))
    Into v_��������
    From ����ҽ������
    Where ���ͺ� = ���ͺ�_In And ҽ��id = ҽ��id_In;
  Else
    v_�������� := 1;
  End If;

  --�ƶ�ϵͳ��Һ���
  If ��Һ���_In = 1 Then
    --��鵱ǰ�������������Ƿ������Һ�Ǽǹ���
    Select Nvl(Zl_Getsysparameter(184), '') Into v_��Һ���� From Dual;
  
    If v_��Һ���� Is Not Null And ִ�н��_In <> 0 Then
      If Instr(',' || v_��Һ���� || ',', ',' || v_����id || ',') > 0 Then
        v_����id   := 0;
        v_��Һ���� := 'Select 1 From ������Һ��¼ where ҽ��ID=:YZID AND ���ͺ�=:FSH AND Ҫ��ʱ��=:YQSJ';
        Begin
          Execute Immediate v_��Һ����
            Into v_����id
            Using ҽ��id_In, ���ͺ�_In, Ҫ��ʱ��_In;
        Exception
          When Others Then
            Null;
        End;
        If v_����id = 0 Then
          v_Error := '��ǰҽ����δ������Һ�����������ִ�еǼǣ�';
          Raise Err_Custom;
        End If;
      End If;
    End If;
    --��鵱ǰҽ���Ƿ�����Һ
  End If;

  --����ҽ��ִ��
  Select Count(1)
  Into v_Count
  From ����ҽ��ִ��
  Where ҽ��id = ҽ��id_In And ���ͺ� = ���ͺ�_In And ִ��ʱ�� = ִ��ʱ��_In;
  If v_Count > 0 Then
    v_Error := '��ָ����ִ��ʱ�䣬�Ѿ�ִ�й�����ҽ���������һ��ִ��ʱ�䡣';
    Raise Err_Custom;
  End If;
  Insert Into ����ҽ��ִ��
    (ҽ��id, ���ͺ�, Ҫ��ʱ��, ��������, ִ��ժҪ, ִ����, ִ��ʱ��, �Ǽ�ʱ��, �Ǽ���, ִ�н��, ˵��, ��Һͨ��, ִ�п���id, ��¼��Դ,ִ�з�ʽ)
  Values
    (ҽ��id_In, ���ͺ�_In, Ҫ��ʱ��_In, n_��������, ִ��ժҪ_In, ִ����_In, ִ��ʱ��_In, v_Date, v_��Ա����, ִ�н��_In, δִ��ԭ��_In, ��Һͨ��_In, n_ִ�п���id,
     ��¼��Դ_In,ִ�з�ʽ_In);

  b_Message.Zlhis_Cis_050(n_����id, n_��ҳid, v_�Һŵ�, ���ͺ�_In, ҽ��id_In, Ҫ��ʱ��_In, ִ��ʱ��_In);

  --���ü�¼��ִ��״̬���и���
  Select Decode(a.ִ��״̬, 1, a.��������, c.�ǼǴ���), Decode(a.ִ��״̬, 1, 0, a.�������� - c.�ǼǴ���), c.�ǼǴ���
  Into n_ִ�д���, n_ʣ�����, n_�Ǽ�����
  From ����ҽ������ a,
       (Select ҽ��id_In ҽ��id, ���ͺ�_In ���ͺ�, Nvl(Sum(b.��������), 0) As �ǼǴ���
         From ����ҽ��ִ�� b
         Where b.ҽ��id = ҽ��id_In And b.���ͺ� = ���ͺ�_In And Nvl(b.ִ�н��, 1) = 1) c
  Where a.ҽ��id = c.ҽ��id And a.���ͺ� = c.���ͺ� And a.ҽ��id = ҽ��id_In And a.���ͺ� = ���ͺ�_In;
  --���ȫ��ִ����״̬Ϊ1��δִ��״̬Ϊ0������ִ��״̬Ϊ2
  Select Decode(n_ʣ�����, 0, 1, Decode(n_ִ�д���, 0, 0, 2)) Into n_ִ��״̬ From Dual;

  --��д��ִ��״̬��ͱ��Ϊ����ִ��
  If Nvl(����ִ��_In, 0) = 1 Then
    Update ����ҽ������
    Set ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3)
    Where ִ��״̬ In (0, 3) And ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In;
  Else
    Update ����ҽ������
    Set ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3)
    Where ִ��״̬ In (0, 3) And ���ͺ� + 0 = ���ͺ�_In And
          ҽ��id In (Select Id
                   From ����ҽ����¼
                   Where Id = v_��id And Nvl(�������, '*') = v_�������
                   Union All
                   Select Id
                   From ����ҽ����¼
                   Where ���id = v_��id And Nvl(�������, '*') = v_�������);
  End If;

  --���¶�Ӧ�ķ���ִ��״̬Ϊ��ִ��(������ִ��)
  --��Ӧ�ô���ҩƷ�͸������õ�����
  If ִ�н��_In = 1 Then
    If n_ִ��״̬ != 0 Then
      d_ִ��ʱ�� := ִ��ʱ��_In;
    End If;
    If v_�������� = 2 Then
      If Nvl(����ִ��_In, 0) = 1 Then
        Update סԺ���ü�¼ A
        Set ִ��״̬ = n_ִ��״̬, ִ���� = Decode(n_ִ��״̬, 0, Null, ִ����_In), ִ��ʱ�� = d_ִ��ʱ��
        Where �շ���� Not In ('5', '6', '7') And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In) And Not Exists
         (Select 1 From �������� Where ����id = a.�շ�ϸĿid And �������� = 1) And a.��¼״̬ In (0, 1, 3) And
              (ҽ�����, NO, ��¼����) In
              (Select ҽ��id, NO, ��¼����
               From ����ҽ������
               Where ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3) And ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In);
      Else
        Update סԺ���ü�¼ A
        Set ִ��״̬ = n_ִ��״̬, ִ���� = Decode(n_ִ��״̬, 0, Null, ִ����_In), ִ��ʱ�� = d_ִ��ʱ��
        Where �շ���� Not In ('5', '6', '7') And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In) And Not Exists
         (Select 1 From �������� Where ����id = a.�շ�ϸĿid And �������� = 1) And a.��¼״̬ In (0, 1, 3) And
              (ҽ�����, NO, ��¼����) In (Select ҽ��id, NO, ��¼����
                                   From ����ҽ������
                                   Where ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3) And ���ͺ� + 0 = ���ͺ�_In And
                                         ҽ��id In (Select ID
                                                  From ����ҽ����¼
                                                  Where ID = v_��id And ������� = v_�������
                                                  Union All
                                                  Select ID
                                                  From ����ҽ����¼
                                                  Where ���id = v_��id And ������� = v_�������));
      End If;
    Else
      If Nvl(����ִ��_In, 0) = 1 Then
        --�������ﵥ��n_ִ��״̬����Ϊ0���Ǽ�ִ�������ѡ��ִ�н��Ϊδִ�У���������ж�
        Update ������ü�¼ A
        Set ִ��״̬ = n_ִ��״̬, ִ���� = Decode(n_ִ��״̬, 0, Null, ִ����_In), ִ��ʱ�� = d_ִ��ʱ��
        Where �շ���� Not In ('5', '6', '7') And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In) And Not Exists
         (Select 1 From �������� Where ����id = a.�շ�ϸĿid And �������� = 1) And a.��¼״̬ In (0, 1, 3) And
              (ҽ�����, NO, ��¼����) In
              (Select ҽ��id, NO, ��¼����
               From ����ҽ������
               Where ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3) And ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In);
      Else
        Update ������ü�¼ A
        Set ִ��״̬ = n_ִ��״̬, ִ���� = Decode(n_ִ��״̬, 0, Null, ִ����_In), ִ��ʱ�� = d_ִ��ʱ��
        Where �շ���� Not In ('5', '6', '7') And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In) And Not Exists
         (Select 1 From �������� Where ����id = a.�շ�ϸĿid And �������� = 1) And a.��¼״̬ In (0, 1, 3) And
              (ҽ�����, NO, ��¼����) In (Select ҽ��id, NO, ��¼����
                                   From ����ҽ������
                                   Where ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3) And ���ͺ� + 0 = ���ͺ�_In And
                                         ҽ��id In (Select Id
                                                  From ����ҽ����¼
                                                  Where Id = v_��id And ������� = v_�������
                                                  Union All
                                                  Select Id
                                                  From ����ҽ����¼
                                                  Where ���id = v_��id And ������� = v_�������));
      End If;
    End If;
    --�����Զ���ɲɼ�
    If v_������� = 'E' And v_�������� = '6' Then
      Update ����ҽ������ a
      Set a.������ = ִ����_In, a.����ʱ�� = ִ��ʱ��_In
      Where ҽ��id In
            (Select Id From ����ҽ����¼ Where Id = v_��id Union All Select Id From ����ҽ����¼ Where ���id = v_��id) And
            ���ͺ� = ���ͺ�_In;
    End If;
  
    --ִ�����δﵽ֮���Զ����ִ��(��Ҫ����PDA�Զ�ִ��)������������ƶ��ٴ�����ʿվ��PDAһ�¡�
    v_�Զ���� := �Զ����_In;
    If �Զ����_In = 1 Then
      --ҽ���Ѿ������״̬�����ٵ���ִ����ɹ��̴˴�����Ϊ���Զ����
      Select Max(a.ִ��״̬) Into v_Count From ����ҽ������ a Where a.ҽ��id = ҽ��id_In And a.���ͺ� = ���ͺ�_In;
      If v_Count = 1 Then
        v_�Զ���� := 0;
      End If;
      v_Count := Null;
    End If;
    --�ƶ���ִ���������Ѫҽ�������Զ����=0�����Զ����ҽ��ִ�У���Ѫ��ӿ��ڲ�����
    n_��Ѫҽ�� := 0;
    If v_������� = 'E' And v_�������� = '8' And n_ִ�з��� = 1 Then
      n_��Ѫҽ�� := Zl_To_Number(Nvl(Zl_Getsysparameter(236), '0'));
    End If;
    If Nvl(v_�Զ����, 0) = 0 And (v_������Դ = 2 Or v_������Դ = 1) And Instr('C,D', v_�������) = 0 And n_��Ѫҽ�� = 0 Then
      Begin
        Execute Immediate 'Select Count(1) From ZLMBSYSTEMS'
          Into v_Count;
      Exception
        When Others Then
          Null;
      End;
      If v_Count > 0 Then
        v_�Զ���� := 1;
      End If;
    End If;
  
    If Nvl(v_�Զ����, 0) = 1 Or ������Ŀ����_In = 1 Then
      Begin
        Select Decode(Sign(Nvl(Sum(b.��������), 0) - a.��������), 1, 1, 0, 1, 0)
        Into v_�Զ����
        From ����ҽ������ a, ����ҽ��ִ�� b
        Where a.ҽ��id = b.ҽ��id And a.���ͺ� = b.���ͺ� And a.ִ��״̬ In (0, 3) And a.ҽ��id = ҽ��id_In And a.���ͺ� = ���ͺ�_In
        Group By a.��������;
      Exception
        When Others Then
          Null;
      End;
    
      If Nvl(v_�Զ����, 0) = 1 Or ������Ŀ����_In = 1 Then
        Zl_����ҽ��ִ��_Finish(ҽ��id_In,
                         ���ͺ�_In,
                         Null,
                         ����ִ��_In,
                         v_��Ա���,
                         v_��Ա����,
                         ִ�в���id_In,
                         ������Ŀ����_In);
      End If;
    End If;
    --����ҽ��ִ�мƼ�.ִ��״̬
    If n_�������� > 0 Then
      Select Count(Distinct Ҫ��ʱ��) Into v_Count From ҽ��ִ�мƼ� Where ҽ��id = ҽ��id_In And ���ͺ� = ���ͺ�_In;
      If v_Count > 0 Then
        n_�������� := n_�������� / v_Count;
        --��ִ������+�������� �ܹ��ܹ�ִ�ж��ٸ�ʱ���,ȡ�������
        v_Count := Ceil((n_�Ǽ�����) / n_��������);
        --��ȡִ�н���Ҫ��ʱ��
        Select Ҫ��ʱ��
        Into d_Ҫ��ʱ��
        From (Select Ҫ��ʱ��, Rownum As ����
               From (Select Distinct Ҫ��ʱ��
                      From ҽ��ִ�мƼ�
                      Where ҽ��id = ҽ��id_In And ���ͺ� = ���ͺ�_In
                      Order By Ҫ��ʱ��))
        Where ���� = v_Count;
      
        If Not d_Ҫ��ʱ�� Is Null Then
          --�ȼ���Ƿ��Ѿ��˷�
          Select Max(Nvl(ִ��״̬, 0))
          Into v_Count
          From ҽ��ִ�мƼ�
          Where ҽ��id = ҽ��id_In And ���ͺ� = ���ͺ�_In And Ҫ��ʱ�� <= d_Ҫ��ʱ��;
          If v_Count = 2 Then
            v_Error := '��ָ����ִ��ʱ��ε�ҽ�������Ѿ����˷ѣ���������ִ�С�';
            Raise Err_Custom;
          End If;
          --���½���Ҫ��ʱ��֮ǰ(��)�ļ�¼ִ��״̬��
          Update ҽ��ִ�мƼ�
          Set ִ��״̬ = 1
          Where ҽ��id = ҽ��id_In And ���ͺ� = ���ͺ�_In And Ҫ��ʱ�� <= d_Ҫ��ʱ�� And Nvl(ִ��״̬, 0) <> 2;
        End If;
      End If;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����ҽ��ִ��_Insert;
/


--130541:������,2018-09-14,����ƽ̨��Ϣ����
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
    ԭ����id_In In ������ҳ.����id%Type
  );

  --69.����ת����ת��
  Procedure Zlhis_Patient_026
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );

  Procedure Zlhis_Patient_028(����id_In In ������ҳ.����id%Type);

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
    ԭ����id_In In ������ҳ.����id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_017',
                                '<root><����ID>' || ����id_In || '</����ID><ԭ����ID>' || ԭ����id_In || '</ԭ����ID></root>');
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

--122812:����,2018-09-13,Ʊ������¼�����ֶ�����,Ʊ�����ü�¼�����ֶ����ID
Create Or Replace Procedure Zl_���ѿ����ü�¼_Insert
(
  Id_In       ���ѿ����ü�¼.Id%Type,
  �ӿڱ��_In ���ѿ����ü�¼. �ӿڱ��%Type,
  ������_In   ���ѿ����ü�¼.������%Type,
  ǰ׺�ı�_In ���ѿ����ü�¼.ǰ׺�ı�%Type,
  ��ʼ����_In ���ѿ����ü�¼.��ʼ����%Type,
  ��ֹ����_In ���ѿ����ü�¼.��ֹ����%Type,
  ʹ�÷�ʽ_In ���ѿ����ü�¼.ʹ�÷�ʽ%Type,
  �Ǽ�ʱ��_In ���ѿ����ü�¼.�Ǽ�ʱ��%Type := Null,
  �Ǽ���_In   ���ѿ����ü�¼.�Ǽ���%Type := Null,
  ʣ������_In ���ѿ����ü�¼.ʣ������%Type := Null,
  ����_In     ���ѿ����ü�¼.����%Type := Null,
  ǩ����_In   ���ѿ����ü�¼.ǩ����%Type := Null,
  ���id_In   ���ѿ����ü�¼.���id%Type := Null
) Is
  v_Err_Msg Varchar2(500);
  Err_Item Exception;

  n_Count  Number(18);
  n_ʣ���� ���ѿ�����¼.ʣ������%Type;
Begin

  For r_��� In (Select ID, ǰ׺�ı�, �ӿڱ��, ��ʼ����, Nvl(��ֹ����, ��ʼ����) As ��ֹ����
               From ���ѿ�����¼
               Where ID = Nvl(���id_In, 0)) Loop
    --1. �����
    If ��ʼ����_In < r_���.��ʼ���� Or ��ʼ����_In > r_���.��ֹ���� Then
      --������ⷶΧ,��������
      v_Err_Msg := '��ǰ���õĿ�ʼ���š�' || ��ʼ����_In || '��������ⷶΧ' || Chr(10) || Chr(13) || '��' || r_���.��ʼ���� || '-' || r_���.��ֹ���� ||
                   '���������øÿ�Ƭ��';
      Raise Err_Item;
    End If;
  
    If ��ֹ����_In < r_���.��ʼ���� Or ��ֹ����_In > r_���.��ֹ���� Then
      --������ⷶΧ,��������
      v_Err_Msg := '��ǰ���õ���ֹ���š�' || ��ֹ����_In || '��������ⷶΧ' || Chr(10) || Chr(13) || '��' || r_���.��ʼ���� || '-' || r_���.��ֹ���� ||
                   '���������øÿ�Ƭ��';
      Raise Err_Item;
    End If;
    If r_���.�ӿڱ�� <> �ӿڱ��_In Then
      v_Err_Msg := '���Ŀ����' || Nvl(�ӿڱ��_In, '') || '�����������һ�¡�' || Nvl(r_���.�ӿڱ��, '') || '��!';
      Raise Err_Item;
    End If;
  
    --2.��鿨���Ƿ��Ѿ�������,�����ظ�����
    Select Count(1)
    Into n_Count
    From ���ѿ������¼
    Where ���id = Nvl(���id_In, 0) And ((��ʼ����_In Between ��ʼ���� And ��ֹ����) Or (��ֹ����_In Between ��ʼ���� And ��ֹ����) Or
          (��ʼ���� Between ��ʼ����_In And ��ֹ����_In) Or (��ֹ���� Between ��ʼ����_In And ��ֹ����_In));
  
    If n_Count <> 0 Then
      If ��ʼ����_In = ��ֹ����_In Then
        v_Err_Msg := '����:' || ��ʼ����_In || '�ڱ����¼���Ѿ����ڣ������ٽ������ã�';
      Else
        v_Err_Msg := '����:' || ��ʼ����_In || '-' || ��ֹ����_In || '�ڱ����¼���Ѿ����ڣ������ٽ������ã�';
      End If;
      Raise Err_Item;
    End If;
  
    --3.��鿨���Ƿ��Ѿ�������,���õĲ����ٽ��б���
    Select Count(1)
    Into n_Count
    From ���ѿ����ü�¼
    Where ���� = Nvl(����_In, 0) And ((��ʼ����_In Between ��ʼ���� And ��ֹ����) Or (��ֹ����_In Between ��ʼ���� And ��ֹ����) Or
          (��ʼ���� Between ��ʼ����_In And ��ֹ����_In) Or (��ֹ���� Between ��ʼ����_In And ��ֹ����_In));
    If n_Count <> 0 Then
      If ��ʼ����_In = ��ֹ����_In Then
        v_Err_Msg := '����:' || ��ʼ����_In || '�����ѿ����ü�¼���Ѿ����ڣ������ٽ������ã�';
      Else
        v_Err_Msg := '����:' || ��ʼ����_In || '-' || ��ֹ����_In || '�����ѿ����ü�¼���Ѿ����ڣ������ٽ������ã�';
      End If;
      Raise Err_Item;
    End If;
    --���ٿ��
    Update ���ѿ�����¼
    Set ʣ������ = Nvl(ʣ������, 0) - Nvl(ʣ������_In, 0)
    Where ID = Nvl(���id_In, 0) And �ӿڱ�� = �ӿڱ��_In
    Returning Nvl(ʣ������, 0) Into n_ʣ����;
  
    If n_ʣ���� < 0 Then
      v_Err_Msg := '��⿨Ƭ��ʣ��Ʊ�������㣬���飡';
      Raise Err_Item;
    End If;
    If n_ʣ���� = 0 Then
      Update ���ѿ�����¼ Set �Ƿ���ڿ� = 0 Where ID = Nvl(���id_In, 0) And �ӿڱ�� = �ӿڱ��_In;
    End If;
  End Loop;

  Insert Into ���ѿ����ü�¼
    (ID, �ӿڱ��, ������, ǰ׺�ı�, ��ʼ����, ��ֹ����, ʹ�÷�ʽ, �Ǽ�ʱ��, �Ǽ���, ʣ������, ����, ǩ����, ǩ��ʱ��, ���id)
  Values
    (Id_In, �ӿڱ��_In, ������_In, ǰ׺�ı�_In, ��ʼ����_In, ��ֹ����_In, ʹ�÷�ʽ_In, �Ǽ�ʱ��_In, �Ǽ���_In, ʣ������_In, ����_In, ǩ����_In,
     Decode(ǩ����_In, Null, Null + Sysdate, Sysdate), ���id_In);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���ѿ����ü�¼_Insert;
/

Create Or Replace Procedure Zl_���ѿ�����¼_Delete(Id_In In ���ѿ�����¼.Id%Type) Is
  v_Err_Msg varchar2(100);
  Err_Item Exception;

  n_Count Number(2);
Begin
  --����Ƿ���ڱ����¼ 
  Select Count(1) Into n_Count From ���ѿ������¼ Where ���id = Id_In And Rownum < 2;
  If n_Count = 1 Then
    v_Err_Msg := '����������Ѿ����ڱ����¼�������ٽ���ɾ����';
    Raise Err_Item;
  End If;

  --����Ƿ����ʹ�ü�¼,�������,�����������ɾ�� 
  Select Count(1) Into n_Count From ���ѿ����ü�¼ Where ���id = Id_In And Rownum < 2;
  If n_Count = 1 Then
    v_Err_Msg := '����������Ѿ��������ü�¼�������ٽ���ɾ����';
    Raise Err_Item;
  End If;

  --�������õ��Ѿ�ת�뵽��ʷ���ݿռ�,��˼�������Ƿ����,����,�϶��Ѿ�ʹ�� 
  Select Count(1) Into n_Count From ���ѿ�����¼ Where ID = Id_In And Nvl(�������, 0) > Nvl(ʣ������, 0);
  If n_Count = 1 Then
    v_Err_Msg := '����������Ѿ���ʹ�ã������ٽ���ɾ����';
    Raise Err_Item;
  End If;

  Delete From ���ѿ�����¼ Where ID = Id_In;
  If Sql%NotFound Then
    v_Err_Msg := '���ڲ���ԭ�򣬸�������ο����Ѿ�������ɾ���������ٽ���ɾ����';
    Raise Err_Item;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���ѿ�����¼_Delete;
/

Create Or Replace Procedure Zl_Ʊ������¼_Delete(Id_In In Ʊ������¼.ID%Type) Is
  v_Err_Msg varchar2(100);
  Err_Item Exception;
  n_Exists Number(2);
Begin

  --����Ƿ���ڱ����¼
  Begin
    Select 1 Into n_Exists From Ʊ�ݱ����¼ Where ���id = Id_In And Rownum = 1;
  Exception
    When Others Then
      n_Exists := 0;
  End;
  If n_Exists = 1 Then
    v_Err_Msg := '[ZLSOFT]Ʊ���Ѿ����ڱ����¼,�����ٽ���ɾ��![ZLSOFT]';
    Raise Err_Item;
  End If;

  --����Ƿ����ʹ�ü�¼,�������,�����������ɾ��
  Begin
    Select 1 Into n_Exists From Ʊ�����ü�¼ Where ���id = Id_In And Rownum = 1;
  Exception
    When Others Then
      n_Exists := 0;
  End;
  If n_Exists = 1 Then
    v_Err_Msg := '[ZLSOFT]Ʊ���Ѿ��������ü�¼,�����ٽ���ɾ��![ZLSOFT]';
    Raise Err_Item;
  End If;

  --�������õ��Ѿ�ת�뵽��ʷ���ݿռ�,��˼�������Ƿ����,����,�϶��Ѿ�ʹ��
  Begin
    Select 1 Into n_Exists From Ʊ������¼ Where ID = Id_In And Nvl(�������, 0) > Nvl(ʣ������, 0);
  Exception
    When Others Then
      n_Exists := 0;
  End;

  If n_Exists = 1 Then
    v_Err_Msg := '[ZLSOFT]Ʊ���Ѿ���ʹ��,�����ٽ���ɾ��![ZLSOFT]';
    Raise Err_Item;
  End If;

  Delete From Ʊ������¼ Where ID = Id_In;
  If Sql%NotFound Then
    v_Err_Msg := '[ZLSOFT]���ڲ���ԭ��,��������ο����Ѿ�������ɾ��,�����ٽ���ɾ��![ZLSOFT]';
    Raise Err_Item;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Ʊ������¼_Delete;
/

--130300:����,2018-09-11,���ֹ���Ʒ�����õġ��ϸ�����÷�������
Create Or Replace Procedure Zl_�÷�����_Update
(
  ҩ��id_In           In �����÷�����.��Ŀid%Type,
  ��������id_In       In Varchar2, --��"|"�ָ��Ĺ���ʵ������� 
  ��������_In         In ҩƷ����.��������%Type,
  �Ƴ�_In             In �����÷�����.�Ƴ�%Type,
  �÷�����_In         In Varchar2, --��"|"�ָ����÷��������ݣ�ÿ����¼��"�÷�ID^Ƶ��^���˼���^С������^ҽ������"��֯ 
  ��ʽ_In             In Number := 0, --0-������Ŀ����,1-��ǰ���;2-�ض�������Ŀ 
  ���_In             In Varchar2 := '0',
  ����id_In           In ������ĿĿ¼.����id%Type := 0,
  ҩƷid_In           In ҩƷ���.ҩƷid%Type := 0,
  �ϸ�����÷�����_In In ҩƷ���.�ϸ�����÷�����%Type := 0
) Is
  --ҩƷid_in :�����ڿ�˵���ǹ�񣬷���ΪƷ�ֻ��߷��� 
  v_Records      Varchar2(4000);
  v_Currrec      Varchar2(1000);
  v_Fields       Varchar2(1000);
  v_�÷�id       �����÷�����.�÷�id%Type;
  v_Ƶ��         �����÷�����.Ƶ��%Type;
  v_���˼���     �����÷�����.���˼���%Type;
  v_С������     �����÷�����.С������%Type;
  v_ҽ������     �����÷�����.ҽ������%Type;
  v_Dddֵ        �����÷�����.Dddֵ%Type;
  v_����         �����÷�����.����%Type;
  v_�Ƿ�Ƥ��     ҩƷ����.�Ƿ�Ƥ��%Type;
  n_ҩ��id       ҩƷ���.ҩ��id%Type;
  n_ҩƷ�÷����� Number; --0-ҩƷ�÷������в��������ݣ�1-ҩƷ�÷������д������� 

  Cursor c_Item Is
    Select i.Id
    From ������ĿĿ¼ I, ҩƷ���� T, (Select ҩƷ���� From ҩƷ���� Where ҩ��id = ҩ��id_In) C
    Where i.Id = t.ҩ��id And t.ҩƷ���� = c.ҩƷ���� And i.����id = ����id_In And i.Id <> ҩ��id_In;
Begin
  --Ʒ�ֺͷ��� 
  If ҩƷid_In = 0 Then
    For r_Item In (Select ID
                   From ������ĿĿ¼
                   Where (��ʽ_In = 0 And ID = ҩ��id_In) Or (��ʽ_In = 1 And ��� = ���_In) Or
                         (����id In (Select ID From ���Ʒ���Ŀ¼ Start With ID = ����id_In Connect By Prior ID = �ϼ�id))) Loop
      --��������
      If ��������id_In Is Not Null Then
        v_�Ƿ�Ƥ�� := 1;
      Else
        v_�Ƿ�Ƥ�� := 0;
      End If;
    
      Update ҩƷ����
      Set �������� = ��������_In, �Ƿ�Ƥ�� = v_�Ƿ�Ƥ��, �ϸ�����÷����� = �ϸ�����÷�����_In
      Where ҩ��id = r_Item.Id;
    
      For r_Spec In (Select b.ҩƷid From ҩƷ��� B Where b.ҩ��id = r_Item.Id) Loop
        Delete From ҩƷ�÷����� Where ҩƷid = r_Spec.ҩƷid And ���� = 0;
      End Loop;
    
      Delete From �����÷����� Where ��Ŀid = r_Item.Id And ���� = 0;
    
      v_Records := ��������id_In;
    
      While v_Records Is Not Null Loop
        v_Currrec := Substr(v_Records, 1, Instr(v_Records, '|') - 1);
        v_Fields  := v_Currrec;
        v_�÷�id  := To_Number(v_Fields);
        Insert Into �����÷����� (��Ŀid, ����, �÷�id) Values (r_Item.Id, 0, v_�÷�id);
        v_Records := Replace('|' || v_Records, '|' || v_Currrec || '|');
      
        --�ж���ҩƷ�÷��������Ƿ���ڶ�Ӧ��Ʒ�ֵĹ������� 
        Begin
          n_ҩƷ�÷����� := 0;
          Select 1
          Into n_ҩƷ�÷�����
          From ҩƷ��� A, ҩƷ�÷����� B
          Where a.ҩƷid = b.ҩƷid And a.ҩ��id = r_Item.Id And b.�÷�id = v_�÷�id And Rownum <= 1;
        Exception
          When Others Then
            n_ҩƷ�÷����� := 0;
        End;
      
        If n_ҩƷ�÷����� = 0 Then
          For r_ҩƷid In (Select ҩƷid From ҩƷ��� Where ҩ��id = r_Item.Id) Loop
            Insert Into ҩƷ�÷����� (ҩƷid, �÷�id, ����) Values (r_ҩƷid.ҩƷid, v_�÷�id, 0);
          End Loop;
        End If;
      End Loop;
    
      --�÷�����   Select ҩƷid From ҩƷ��� Where ҩ��id = r_Item.Id
      For r_Spec In (Select b.ҩƷid From ҩƷ��� B Where b.ҩ��id = r_Item.Id) Loop
        Delete From ҩƷ�÷����� Where ҩƷid = r_Spec.ҩƷid And ���� > 0;
      End Loop;
      Delete From �����÷����� Where ��Ŀid = r_Item.Id And ���� > 0;
    
      If �÷�����_In Is Null Then
        v_Records := Null;
      Else
        v_Records := �÷�����_In || '|';
      End If;
      v_���� := 0;
      While v_Records Is Not Null Loop
        v_Currrec  := Substr(v_Records, 1, Instr(v_Records, '|') - 1);
        v_Fields   := v_Currrec;
        v_�÷�id   := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
        v_Fields   := Substr(v_Fields, Instr(v_Fields, '^') + 1);
        v_Ƶ��     := Substr(v_Fields, 1, Instr(v_Fields, '^') - 1);
        v_Fields   := Substr(v_Fields, Instr(v_Fields, '^') + 1);
        v_���˼��� := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
        v_Fields   := Substr(v_Fields, Instr(v_Fields, '^') + 1);
        v_С������ := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
        v_Fields   := Substr(v_Fields, Instr(v_Fields, '^') + 1);
        v_ҽ������ := Substr(v_Fields, 1, Instr(v_Fields, '^') - 1);
        v_Fields   := Substr(v_Fields, Instr(v_Fields, '^') + 1);
        v_Dddֵ    := To_Number(v_Fields);
        v_����     := v_���� + 1;
        Insert Into �����÷�����
          (��Ŀid, ����, �÷�id, Ƶ��, ���˼���, С������, ҽ������, �Ƴ�, Dddֵ)
        Values
          (r_Item.Id, v_����, v_�÷�id, v_Ƶ��, v_���˼���, v_С������, v_ҽ������, �Ƴ�_In, v_Dddֵ);
        If ����id_In <> 0 Then
          For t_Item In c_Item Loop
            Delete From �����÷����� Where ��Ŀid = t_Item.Id And �÷�id = v_�÷�id And ���� > 0;
            Insert Into �����÷����� (��Ŀid, ����, �÷�id, Ƶ��) Values (t_Item.Id, v_����, v_�÷�id, v_Ƶ��);
          End Loop;
        End If;
      
        --�ж���ҩƷ�÷��������Ƿ���ڶ�Ӧ��Ʒ�ֵĹ������� 
        Begin
          n_ҩƷ�÷����� := 0;
          Select 1
          Into n_ҩƷ�÷�����
          From ҩƷ��� A, ҩƷ�÷����� B
          Where a.ҩƷid = b.ҩƷid And a.ҩ��id = r_Item.Id And b.�÷�id = v_�÷�id And Rownum <= 1;
        Exception
          When Others Then
            n_ҩƷ�÷����� := 0;
        End;
      
        If n_ҩƷ�÷����� = 0 Then
          For r_ҩƷid In (Select ҩƷid From ҩƷ��� Where ҩ��id = r_Item.Id) Loop
            Insert Into ҩƷ�÷�����
              (ҩƷid, �÷�id, Ƶ��, ���˼���, С������, ҽ������, �Ƴ�, Dddֵ, ����)
            Values
              (r_ҩƷid.ҩƷid, v_�÷�id, v_Ƶ��, v_���˼���, v_С������, v_ҽ������, �Ƴ�_In, v_Dddֵ, 1);
          End Loop;
        End If;
      
        v_Records := Replace('|' || v_Records, '|' || v_Currrec || '|');
      End Loop;
    End Loop;
  Else
    --��� 
    --��������
    If ��������id_In Is Not Null Then
      v_�Ƿ�Ƥ�� := 1;
    Else
      v_�Ƿ�Ƥ�� := 0;
    End If;
    Select ҩ��id Into n_ҩ��id From ҩƷ��� Where ҩƷid = ҩƷid_In;
    Update ҩƷ���� Set �������� = ��������_In, �Ƿ�Ƥ�� = v_�Ƿ�Ƥ�� Where ҩ��id = n_ҩ��id;
    Update ҩƷ��� Set �ϸ�����÷����� = �ϸ�����÷�����_In Where ҩƷid = ҩƷid_In;
  
    Delete From ҩƷ�÷����� Where ҩƷid = ҩƷid_In And ���� = 0;
    v_Records := ��������id_In;
  
    While v_Records Is Not Null Loop
      v_Currrec := Substr(v_Records, 1, Instr(v_Records, '|') - 1);
      v_Fields  := v_Currrec;
      v_�÷�id  := To_Number(v_Fields);
      Insert Into ҩƷ�÷����� (ҩƷid, �÷�id, ����) Values (ҩƷid_In, v_�÷�id, 0);
      v_Records := Replace('|' || v_Records, '|' || v_Currrec || '|');
    End Loop;
  
    --�÷�����
    Delete From ҩƷ�÷����� Where ҩƷid = ҩƷid_In And ���� > 0;
    If �÷�����_In Is Null Then
      v_Records := Null;
    Else
      v_Records := �÷�����_In || '|';
    End If;
    v_���� := 0;
    While v_Records Is Not Null Loop
      v_Currrec  := Substr(v_Records, 1, Instr(v_Records, '|') - 1);
      v_Fields   := v_Currrec;
      v_�÷�id   := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
      v_Fields   := Substr(v_Fields, Instr(v_Fields, '^') + 1);
      v_Ƶ��     := Substr(v_Fields, 1, Instr(v_Fields, '^') - 1);
      v_Fields   := Substr(v_Fields, Instr(v_Fields, '^') + 1);
      v_���˼��� := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
      v_Fields   := Substr(v_Fields, Instr(v_Fields, '^') + 1);
      v_С������ := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
      v_Fields   := Substr(v_Fields, Instr(v_Fields, '^') + 1);
      v_ҽ������ := Substr(v_Fields, 1, Instr(v_Fields, '^') - 1);
      v_Fields   := Substr(v_Fields, Instr(v_Fields, '^') + 1);
      v_Dddֵ    := To_Number(v_Fields);
      v_����     := v_���� + 1;
    
      Insert Into ҩƷ�÷�����
        (ҩƷid, �÷�id, Ƶ��, ���˼���, С������, ҽ������, �Ƴ�, Dddֵ, ����)
      Values
        (ҩƷid_In, v_�÷�id, v_Ƶ��, v_���˼���, v_С������, v_ҽ������, �Ƴ�_In, v_Dddֵ, 1);
    
      v_Records := Replace('|' || v_Records, '|' || v_Currrec || '|');
    End Loop;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�÷�����_Update;
/

--129973:����,2018-09-10,���Ӵ��Ρ��Ƿ�����������
Create Or Replace Procedure Zl_��ҩ���_Update
(
  ҩƷid_In         In ҩƷ���.ҩƷid%Type,
  ����_In           In �շ���ĿĿ¼.����%Type,
  ���_In           In �շ���ĿĿ¼.���%Type,
  ����_In           In �շ���ĿĿ¼.����%Type := Null,
  Ʒ��_In           In �շ���Ŀ����.����%Type := Null,
  ƴ��_In           In �շ���Ŀ����.����%Type := Null,
  ���_In           In �շ���Ŀ����.����%Type := Null,
  ������_In         In �շ���Ŀ����.����%Type := Null,
  ��ʶ��_In         In ҩƷ���.��ʶ��%Type := Null,
  ҩƷ��Դ_In       In ҩƷ���.ҩƷ��Դ%Type := Null,
  ��׼�ĺ�_In       In ҩƷ���.��׼�ĺ�%Type := Null,
  ע���̱�_In       In ҩƷ���.ע���̱�%Type := Null,
  �ۼ۵�λ_In       In �շ���ĿĿ¼.���㵥λ%Type := Null,
  ����ϵ��_In       In ҩƷ���.����ϵ��%Type := Null,
  ���ﵥλ_In       In ҩƷ���.���ﵥλ%Type := Null,
  �����װ_In       In ҩƷ���.�����װ%Type := Null,
  סԺ��λ_In       In ҩƷ���.סԺ��λ%Type := Null,
  סԺ��װ_In       In ҩƷ���.סԺ��װ%Type := Null,
  ҩ�ⵥλ_In       In ҩƷ���.ҩ�ⵥλ%Type := Null,
  ҩ���װ_In       In ҩƷ���.ҩ���װ%Type := Null,
  ���쵥λ_In       In ҩƷ���.���쵥λ%Type := 1,
  ���췧ֵ_In       In ҩƷ���.���췧ֵ%Type := Null,
  �Ƿ���_In       In �շ���ĿĿ¼.�Ƿ���%Type := Null,
  ָ��������_In     In ҩƷ���.ָ��������%Type := Null,
  ����_In           In ҩƷ���.����%Type := 95,
  ָ�����ۼ�_In     In ҩƷ���.ָ�����ۼ�%Type := Null,
  �ӳ���_In         In ҩƷ���.�ӳ���%Type := Null,
  ����ѱ���_In     In ҩƷ���.����ѱ���%Type := Null,
  ҩ�ۼ���_In       In ҩƷ���.ҩ�ۼ���%Type := Null,
  ��������_In       In �շ���ĿĿ¼.��������%Type := Null,
  �������_In       In �շ���ĿĿ¼.�������%Type := Null,
  Gmp��֤_In        In ҩƷ���.Gmp��֤%Type := 0,
  �б�ҩƷ_In       In ҩƷ���.�б�ҩƷ%Type := 0,
  ���ηѱ�_In       In �շ���ĿĿ¼.���ηѱ�%Type := 0,
  סԺ�ɷ����_In   In ҩƷ���.סԺ�ɷ����%Type := 0,
  ҩ�����_In       In ҩƷ���.ҩ�����%Type := Null,
  ҩ������_In       In ҩƷ���.ҩ������%Type := Null,
  ���Ч��_In       In ҩƷ���.���Ч��%Type := Null,
  ���������_In     In ҩƷ���.���������%Type := 0,
  �ɱ���_In         In ҩƷ���.�ɱ���%Type := 0,
  ��ǰ�ۼ�_In       In �շѼ�Ŀ.�ּ�%Type := 0,
  ����id_In         In �շѼ�Ŀ.������Ŀid%Type := Null,
  ��ͬ��λid_In     In ҩƷ���.��ͬ��λid%Type := Null,
  ˵��_In           In �շ���ĿĿ¼.˵��%Type := Null,
  ��̬����_In       In ҩƷ���.��̬����%Type := 0,
  ��ҩ����_In       In ҩƷ���.��ҩ����%Type := Null,
  ��ѡ��_In         In �շ���ĿĿ¼.��ѡ��%Type := Null,
  ��ֵ˰��_In       In ҩƷ���.��ֵ˰��%Type := Null,
  ����ҩ��_In       In ҩƷ���.����ҩ��%Type := Null,
  վ��_In           In �շ���ĿĿ¼.վ��%Type := Null,
  �Ƿ񳣱�_In       In ҩƷ���.�Ƿ񳣱�%Type := Null,
  �洢�¶�_In       In ��ҺҩƷ����.�洢�¶�%Type := Null,
  �洢����_In       In ��ҺҩƷ����.�洢����%Type := Null,
  ��ҩ����_In       In ��ҺҩƷ����.��ҩ����%Type := Null,
  �Ƿ�������_In   In ��ҺҩƷ����.�Ƿ�������%Type := Null,
  ����_In           In ҩƷ���.����%Type := Null,
  ������Ŀ_In       In �շ���ĿĿ¼.������Ŀ%Type := Null,
  ����ɷ����_In   In ҩƷ���.����ɷ����%Type := 0,
  Dddֵ_In          In ҩƷ���.Dddֵ%Type := 0,
  ��ΣҩƷ_In       ҩƷ���.��ΣҩƷ%Type := Null,
  �ͻ���λ_In       In ҩƷ���.�ͻ���λ%Type := Null,
  �ͻ���װ_In       In ҩƷ���.�ͻ���װ%Type := Null,
  ��Һע������_In   In ��ҺҩƷ����.��Һע������%Type := Null,
  �Ƿ��ҩ_In       In ҩƷ���.�Ƿ��ҩ%Type := Null,
  �Ƿ����۹���_In In ҩƷ���.�Ƿ����۹���%Type := Null,
  ��λ��_In         In ҩƷ���.��λ��%Type := Null,
  �Ƿ���������_In   In ҩƷ���.�Ƿ���������%Type := Null
) Is
  v_ҩ��id   ������ĿĿ¼.Id%Type;
  v_����     ������ĿĿ¼.����%Type;
  v_�Ƿ��� �շ���ĿĿ¼.�Ƿ���%Type; --������ҩƷ��ʱ��Ϊʱ�ۣ�ʱ��ҩƷֻ����δ������������޸�Ϊ���ۣ���������������޸Ķ������� 
  v_����     Number(2);
  Err_Notfind Exception;
  v_No           �շѼ�Ŀ.No%Type;
  v_Temp         �շ���ĿĿ¼.������Ŀ%Type;
  v_������Ŀ     �շ���ĿĿ¼.������Ŀ%Type;
  n_ָ�������   ҩƷ���.ָ�������%Type;
  n_ҩƷ�ϴ��ۼ� ҩƷ���.�ϴ��ۼ�%Type;
  n_���۽��     ҩƷ���.ʵ�ʽ��%Type;
  n_�շ�id       ҩƷ�շ���¼.Id%Type;
  n_��ͨ���С�� Number;
  n_���         Number(8);
  Classid        Number(18); --������
  v_Billno       ҩƷ�շ���¼.No%Type; --���۵���
  n_�۸�id       �շѼ�Ŀ.Id%Type;
  n_�շѼ�Ŀ�ּ� �շѼ�Ŀ.�ּ�%Type;
  n_ԭ��         ҩƷ�۸��¼.ԭ��%Type;
  n_ҩƷ�۸��¼ Number(1);
  v_���         �շ���ĿĿ¼.���%Type;
  --����->ʱ�ۺ����ҩƷ�۸��¼��ֵ

  Cursor c_Priceadjust Is
    Select s.ҩƷid, s.�ⷿid, Nvl(s.����, 0) As ����, s.�ϴι�Ӧ��id As ��Ӧ��id, s.�ϴ����� As ����, s.Ч��, s.�ϴβ��� As ����,
           Nvl(s.ʵ������, 0) As ʵ������, s.�ϴο��� As ����, Nvl(s.ʵ�ʽ��, 0) As ʵ�ʽ��, Nvl(s.ʵ�ʲ��, 0) As ʵ�ʲ��, Nvl(s.���ۼ�, 0) As ���ۼ�,
           s.ƽ���ɱ���, s.���Ч��, s.��׼�ĺ�, s.�ϴ��������� As ��������
    From ҩƷ��� S
    Where s.ҩƷid = ҩƷid_In And s.���� = 1 
    Order By s.ҩƷid, s.����, s.�ⷿid;

  r_Priceadjust c_Priceadjust%RowType;

Begin
  v_������Ŀ := ������Ŀ_In;
  --�жϲ�����Ŀ 
  If v_������Ŀ Is Null Then
    If ����id_In Is Not Null Then
      Begin
        Select ������Ŀ Into v_Temp From ������Ŀ Where ID = ����id_In;
      Exception
        When Others Then
          v_Temp := Null;
      End;
      If v_Temp Is Not Null Then
        v_������Ŀ := v_Temp;
      End If;
    End If;
  End If;
  --ͨ������ 
  Select ID, ����
  Into v_ҩ��id, v_����
  From ������ĿĿ¼
  Where ID = (Select ҩ��id From ҩƷ��� Where ҩƷid = ҩƷid_In);
  --ȡԭʼ�Ķ������� 
  Select �Ƿ��� Into v_�Ƿ��� From �շ���ĿĿ¼ Where ID = ҩƷid_In;
  --�����Ϣ 
  Update �շ���ĿĿ¼
  Set ���� = ����_In, ���� = v_����, ��� = ���_In, ���� = ����_In, ���㵥λ = �ۼ۵�λ_In, �������� = ��������_In, ������� = �������_In, ���ηѱ� = ���ηѱ�_In,
      ������Ŀ = v_������Ŀ, ˵�� = ˵��_In, ��ѡ�� = ��ѡ��_In, վ�� = վ��_In
  Where ID = ҩƷid_In
  Returning ��� Into v_���;
  If Sql%RowCount = 0 Then
    Raise Err_Notfind;
  End If;
  n_ָ������� := (1 - 1 / (1 + �ӳ���_In / 100)) * 100;
  Update ҩƷ���
  Set ��ʶ�� = ��ʶ��_In, ҩƷ��Դ = ҩƷ��Դ_In, ��׼�ĺ� = ��׼�ĺ�_In, ע���̱� = ע���̱�_In, ����ϵ�� = ����ϵ��_In, ���ﵥλ = ���ﵥλ_In, �����װ = �����װ_In,
      סԺ��λ = סԺ��λ_In, סԺ��װ = סԺ��װ_In, ҩ�ⵥλ = ҩ�ⵥλ_In, ҩ���װ = ҩ���װ_In, ���쵥λ = ���쵥λ_In, ���췧ֵ = ���췧ֵ_In, ָ�������� = ָ��������_In,
      ���� = ����_In, ָ�����ۼ� = ָ�����ۼ�_In, ָ������� = n_ָ�������, ����ѱ��� = ����ѱ���_In, ҩ�ۼ��� = ҩ�ۼ���_In, סԺ�ɷ���� = סԺ�ɷ����_In,
      ҩ����� = ҩ�����_In, ҩ������ = ҩ������_In, ���Ч�� = ���Ч��_In, �б�ҩƷ = �б�ҩƷ_In, Gmp��֤ = Gmp��֤_In, ��������� = ���������_In,
      ��ͬ��λid = ��ͬ��λid_In, ��̬���� = ��̬����_In, ��ҩ���� = ��ҩ����_In, ��ֵ˰�� = ��ֵ˰��_In, ����ҩ�� = ����ҩ��_In, �Ƿ񳣱� = �Ƿ񳣱�_In, ���� = ����_In,
      ����ɷ���� = ����ɷ����_In, Dddֵ = Dddֵ_In, ��ΣҩƷ = ��ΣҩƷ_In, �ͻ���λ = �ͻ���λ_In, �ͻ���װ = �ͻ���װ_In, �ӳ��� = �ӳ���_In, �Ƿ��ҩ = �Ƿ��ҩ_In,
      �Ƿ����۹��� = �Ƿ����۹���_In, ��λ�� = ��λ��_In, �Ƿ��������� = �Ƿ���������_In
  Where ҩƷid = ҩƷid_In;

  --�����޸ģ�����ҩƷ������ҩ���г�ҩ��ʱ��ȱʡ�������Ϊ�����סԺ������޸Ĺ��ҩƷʱ�����ٸ��ݹ��ҩƷ�ķ���������ҩƷ�ķ������ 
  --������Ŀ�������ĸ��� 
  --select nvl(sum(distinct I.�������),0) into v_���� 
  --from �շ���ĿĿ¼ I,ҩƷ��� S 
  --where I.ID=S.ҩƷID and S.ҩ��ID=v_ҩ��ID; 
  --update ������ĿĿ¼ 
  --set �������=decode(v_����,0,0,1,1,2,2,3) 
  --where ID=v_ҩ��ID; 

  --�����Ĵ��� 
  If ������_In Is Null Then
    Delete �շ���Ŀ���� Where �շ�ϸĿid = ҩƷid_In And ���� = 1 And ���� = 3;
  Else
    Update �շ���Ŀ���� Set ���� = v_����, ���� = ������_In Where �շ�ϸĿid = ҩƷid_In And ���� = 1 And ���� = 3;
    If Sql%RowCount = 0 Then
      Insert Into �շ���Ŀ���� (�շ�ϸĿid, ����, ����, ����, ����) Values (ҩƷid_In, v_����, 1, ������_In, 3);
    End If;
  End If;
  If Ʒ��_In Is Null Then
    Delete �շ���Ŀ���� Where �շ�ϸĿid = ҩƷid_In And ���� = 3;
  Else
    If ƴ��_In Is Null Then
      Delete �շ���Ŀ���� Where �շ�ϸĿid = ҩƷid_In And ���� = 3 And ���� = 1;
    Else
      Update �շ���Ŀ���� Set ���� = Ʒ��_In, ���� = ƴ��_In Where �շ�ϸĿid = ҩƷid_In And ���� = 3 And ���� = 1;
      If Sql%RowCount = 0 Then
        Insert Into �շ���Ŀ���� (�շ�ϸĿid, ����, ����, ����, ����) Values (ҩƷid_In, Ʒ��_In, 3, ƴ��_In, 1);
      End If;
    End If;
    If ���_In Is Null Then
      Delete �շ���Ŀ���� Where �շ�ϸĿid = ҩƷid_In And ���� = 3 And ���� = 2;
    Else
      Update �շ���Ŀ���� Set ���� = Ʒ��_In, ���� = ���_In Where �շ�ϸĿid = ҩƷid_In And ���� = 3 And ���� = 2;
      If Sql%RowCount = 0 Then
        Insert Into �շ���Ŀ���� (�շ�ϸĿid, ����, ����, ����, ����) Values (ҩƷid_In, Ʒ��_In, 3, ���_In, 2);
      End If;
    End If;
  End If;

  --������Ϣ������Ѿ��з�����������ֱ�Ӹ�����Щ��Ϣ 
  Select Nvl(Count(*), 0) Into v_���� From ҩƷ�շ���¼ Where ҩƷid = ҩƷid_In And Rownum < 2;
  If v_���� = 0 Then
    Update ҩƷ��� Set �ɱ��� = �ɱ���_In Where ҩƷid = ҩƷid_In;
    If ����id_In Is Not Null Then
      Update �շѼ�Ŀ
      Set �ּ� = ��ǰ�ۼ�_In, ������Ŀid = ����id_In, �䶯ԭ�� = 1, ����˵�� = '�޸Ķ���', ������ = User
      Where �շ�ϸĿid = ҩƷid_In And (Sysdate Between ִ������ And ��ֹ���� Or Sysdate >= ִ������ And ��ֹ���� Is Null) And �䶯ԭ�� = 1;
      If Sql%RowCount = 0 Then
        v_No := Nextno(9);
        Insert Into �շѼ�Ŀ
          (ID, ԭ��id, �շ�ϸĿid, ԭ��, �ּ�, ������Ŀid, �䶯ԭ��, ����˵��, ������, ִ������, ��ֹ����, NO, ���)
        Values
          (�շѼ�Ŀ_Id.Nextval, Null, ҩƷid_In, 0, ��ǰ�ۼ�_In, ����id_In, 1, '��������', User, Sysdate,
           To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
      End If;
    End If;
  Else
    --����ҵ������ʱ�������޸ļ۸��ǿ����޸�������Ŀ 
    Update �շѼ�Ŀ
    Set ������Ŀid = ����id_In
    Where �շ�ϸĿid = ҩƷid_In And (Sysdate Between ִ������ And ��ֹ���� Or Sysdate >= ִ������ And ��ֹ���� Is Null) And �䶯ԭ�� = 1;
  End If;

  --ʱ��->����
  If v_�Ƿ��� = 1 And �Ƿ���_In = 0 Then
    Update �շ���ĿĿ¼ Set �Ƿ��� = �Ƿ���_In Where ID = ҩƷid_In;
  
    Begin
      Select �ּ�, ID As �۸�id
      Into n_�շѼ�Ŀ�ּ�, n_�۸�id
      From �շѼ�Ŀ
      Where �շ�ϸĿid = ҩƷid_In And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And �䶯ԭ�� = 1;
    Exception
      When Others Then
        n_�շѼ�Ŀ�ּ� := Null;
        n_�۸�id       := Null;
    End;
  
    Begin
      Select �ϴ��ۼ� Into n_ҩƷ�ϴ��ۼ� From ҩƷ��� Where ҩƷid = ҩƷid_In;
    Exception
      When Others Then
        n_ҩƷ�ϴ��ۼ� := Null;
    End;
  
    If n_ҩƷ�ϴ��ۼ� Is Null Then
      n_ҩƷ�ϴ��ۼ� := n_�շѼ�Ŀ�ּ�;
    End If;
  
    Zl_�շѼ�Ŀ_Stop(ҩƷid_In, Sysdate - 1 / 24 / 60 / 60);
    v_No := Nextno(9);
    Insert Into �շѼ�Ŀ
      (ID, ԭ��id, �շ�ϸĿid, ԭ��, �ּ�, ������Ŀid, �䶯ԭ��, ����˵��, ������, ִ������, ��ֹ����, NO, ���)
    Values
      (�շѼ�Ŀ_Id.Nextval, n_�۸�id, ҩƷid_In, n_�շѼ�Ŀ�ּ�, n_ҩƷ�ϴ��ۼ�, ����id_In, 1, 'ʱ��ת����', Zl_Username, Sysdate,
       To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
  
    Select ���� Into n_��ͨ���С�� From ҩƷ���ľ��� Where ��� = 1 And ���� = 4 And ��λ = 5;
  
    --ȡ������ID
    Select ���id Into Classid From ҩƷ�������� Where ���� = 13;
  
    n_���   := 0;
    v_Billno := Null;
  
    For r_Priceadjust In c_Priceadjust Loop
      If n_ҩƷ�ϴ��ۼ� <> r_Priceadjust.���ۼ� Then
        If v_Billno Is Null Then
          Select Nextno(147) Into v_Billno From Dual;
        End If;
        n_��� := n_��� + 1;
        Select ҩƷ�շ���¼_Id.Nextval Into n_�շ�id From Dual;
        n_���۽�� := Round(n_ҩƷ�ϴ��ۼ� * r_Priceadjust.ʵ������, n_��ͨ���С��) -
                  Round(r_Priceadjust.���ۼ� * r_Priceadjust.ʵ������, n_��ͨ���С��);
        --��������Ӱ���¼
        Insert Into ҩƷ�շ���¼
          (ID, ��¼״̬, ����, NO, ���, ������id, ҩƷid, ����, ����, Ч��, ����, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ����, ���۽��, ���, ժҪ, ������,
           ��������, �ⷿid, ���ϵ��, �۸�id, �����, �������, ����, Ƶ��, ��ҩ��λid)
        Values
          (n_�շ�id, 1, 13, v_Billno, n_���, Classid, r_Priceadjust.ҩƷid, r_Priceadjust.����, r_Priceadjust.����,
           r_Priceadjust.Ч��, r_Priceadjust.����, 1, r_Priceadjust.ʵ������, 0, r_Priceadjust.���ۼ�, 0, n_ҩƷ�ϴ��ۼ�,
           r_Priceadjust.����, n_���۽��, n_���۽��, 'ʱ��ת����', Zl_Username, Sysdate, r_Priceadjust.�ⷿid, 1, n_�۸�id, Zl_Username,
           Sysdate, r_Priceadjust.ʵ�ʽ��, r_Priceadjust.ʵ�ʲ��, r_Priceadjust.��Ӧ��id);
      
        Zl_ҩƷ���_Update(n_�շ�id, 2, 0);
      End If;
    End Loop;
  
    --����->ʱ��
  Elsif v_�Ƿ��� = 0 And �Ƿ���_In = 1 Then
    Update �շ���ĿĿ¼ Set �Ƿ��� = �Ƿ���_In Where ID = ҩƷid_In;
    Begin
      Select �ּ�, ID As �۸�id
      Into n_�շѼ�Ŀ�ּ�, n_�۸�id
      From �շѼ�Ŀ
      Where �շ�ϸĿid = ҩƷid_In And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And �䶯ԭ�� = 1;
    Exception
      When Others Then
        n_�շѼ�Ŀ�ּ� := Null;
        n_�۸�id       := Null;
    End;
  
    Zl_�շѼ�Ŀ_Stop(ҩƷid_In, Sysdate - 1 / 24 / 60 / 60);
    v_No := Nextno(9);
    Insert Into �շѼ�Ŀ
      (ID, ԭ��id, �շ�ϸĿid, ԭ��, �ּ�, ������Ŀid, �䶯ԭ��, ����˵��, ������, ִ������, ��ֹ����, NO, ���)
    Values
      (�շѼ�Ŀ_Id.Nextval, n_�۸�id, ҩƷid_In, n_�շѼ�Ŀ�ּ�, n_�շѼ�Ŀ�ּ�, ����id_In, 1, '����תʱ��', Zl_Username, Sysdate,
       To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
  
    For r_Priceadjust In c_Priceadjust Loop
      n_ҩƷ�۸��¼ := 0;
      Begin
        Select 1, �ּ�
        Into n_ҩƷ�۸��¼, n_ԭ��
        From ҩƷ�۸��¼
        Where ҩƷid = r_Priceadjust.ҩƷid And �ⷿid = r_Priceadjust.�ⷿid And Nvl(����, 0) = r_Priceadjust.���� And ��¼״̬ = 1 And
              �۸����� = 1;
      Exception
        When Others Then
          n_ҩƷ�۸��¼ := 0;
          n_ԭ��         := n_�շѼ�Ŀ�ּ�;
      End;
    
      If n_ҩƷ�۸��¼ = 1 Then
        Zl_ҩƷ�۸��¼_Stop(1, r_Priceadjust.�ⷿid, r_Priceadjust.ҩƷid, r_Priceadjust.����, Sysdate - 1 / 24 / 60 / 60, 2);
      End If;
      Zl_ҩƷ�۸��¼_Insert(0, 1, r_Priceadjust.�ⷿid, r_Priceadjust.ҩƷid, r_Priceadjust.����, n_ԭ��, n_�շѼ�Ŀ�ּ�, Sysdate, '����תʱ��',
                       Zl_Username, Null, r_Priceadjust.��Ӧ��id, r_Priceadjust.����, r_Priceadjust.Ч��, r_Priceadjust.����,
                       r_Priceadjust.���Ч��, Null, Null, Null, Null, 1);
    
      Update ҩƷ���
      Set ���ۼ� = n_�շѼ�Ŀ�ּ�
      Where ���� = 1 And �ⷿid = r_Priceadjust.�ⷿid And ҩƷid = r_Priceadjust.ҩƷid And Nvl(����, 0) = r_Priceadjust.����;
    
    End Loop;
  End If;

  --ҩƷ�����̱Ƚ����� 
  If ����_In Is Not Null Then
    Update ҩƷ������ Set ���� = ����_In Where ���� = ����_In;
    If Sql%RowCount = 0 Then
      Insert Into ҩƷ������
        (����, ����, ����)
        Select Nvl(Max(To_Number(����)), 0) + 1, ����_In, zlSpellCode(����_In, 10) From ҩƷ������;
    End If;
  End If;

  --�޸���ҺҩƷ���� 
  Update ��ҺҩƷ����
  Set �洢�¶� = �洢�¶�_In, �洢���� = �洢����_In, ��ҩ���� = ��ҩ����_In, �Ƿ������� = �Ƿ�������_In, ��Һע������ = ��Һע������_In
  Where ҩƷid = ҩƷid_In;

  If Sql%NotFound Then
    Insert Into ��ҺҩƷ����
      (ҩƷid, �洢�¶�, �洢����, ��ҩ����, �Ƿ�������, ��Һע������)
    Values
      (ҩƷid_In, �洢�¶�_In, �洢����_In, ��ҩ����_In, �Ƿ�������_In, ��Һע������_In);
  End If;

  --ҩƷ���ȵ���(����ģʽʱ)
  Zl_ҩƷ���ľ���_���۵���;

  b_Message.Zlhis_Dict_036(v_���, ҩƷid_In);
Exception
  When Err_Notfind Then
    Raise_Application_Error(-20101, '[ZLSOFT]�ù�񲻴��ڣ������ѱ������û�ɾ����[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��ҩ���_Update;
/

Create Or Replace Procedure Zl_��ҩ���_Insert
(
  ҩ��id_In         In ҩƷ���.ҩ��id%Type,
  ҩƷid_In         In ҩƷ���.ҩƷid%Type,
  ����_In           In �շ���ĿĿ¼.����%Type,
  ���_In           In �շ���ĿĿ¼.���%Type,
  ����_In           In �շ���ĿĿ¼.����%Type := Null,
  Ʒ��_In           In �շ���Ŀ����.����%Type := Null,
  ƴ��_In           In �շ���Ŀ����.����%Type := Null,
  ���_In           In �շ���Ŀ����.����%Type := Null,
  ������_In         In �շ���Ŀ����.����%Type := Null,
  ��ʶ��_In         In ҩƷ���.��ʶ��%Type := Null,
  ҩƷ��Դ_In       In ҩƷ���.ҩƷ��Դ%Type := Null,
  ��׼�ĺ�_In       In ҩƷ���.��׼�ĺ�%Type := Null,
  ע���̱�_In       In ҩƷ���.ע���̱�%Type := Null,
  �ۼ۵�λ_In       In �շ���ĿĿ¼.���㵥λ%Type := Null,
  ����ϵ��_In       In ҩƷ���.����ϵ��%Type := Null,
  ���ﵥλ_In       In ҩƷ���.���ﵥλ%Type := Null,
  �����װ_In       In ҩƷ���.�����װ%Type := Null,
  סԺ��λ_In       In ҩƷ���.סԺ��λ%Type := Null,
  סԺ��װ_In       In ҩƷ���.סԺ��װ%Type := Null,
  ҩ�ⵥλ_In       In ҩƷ���.ҩ�ⵥλ%Type := Null,
  ҩ���װ_In       In ҩƷ���.ҩ���װ%Type := Null,
  ���쵥λ_In       In ҩƷ���.���쵥λ%Type := 1,
  ���췧ֵ_In       In ҩƷ���.���췧ֵ%Type := Null,
  �Ƿ���_In       In �շ���ĿĿ¼.�Ƿ���%Type := Null,
  ָ��������_In     In ҩƷ���.ָ��������%Type := Null,
  ����_In           In ҩƷ���.����%Type := 95,
  ָ�����ۼ�_In     In ҩƷ���.ָ�����ۼ�%Type := Null,
  �ӳ���_In         In ҩƷ���.�ӳ���%Type := Null,
  ����ѱ���_In     In ҩƷ���.����ѱ���%Type := Null,
  ҩ�ۼ���_In       In ҩƷ���.ҩ�ۼ���%Type := Null,
  ��������_In       In �շ���ĿĿ¼.��������%Type := Null,
  �������_In       In �շ���ĿĿ¼.�������%Type := Null,
  Gmp��֤_In        In ҩƷ���.Gmp��֤%Type := 0,
  �б�ҩƷ_In       In ҩƷ���.�б�ҩƷ%Type := 0,
  ���ηѱ�_In       In �շ���ĿĿ¼.���ηѱ�%Type := 0,
  סԺ�ɷ����_In   In ҩƷ���.סԺ�ɷ����%Type := 0,
  ҩ�����_In       In ҩƷ���.ҩ�����%Type := Null,
  ҩ������_In       In ҩƷ���.ҩ������%Type := Null,
  ���Ч��_In       In ҩƷ���.���Ч��%Type := Null,
  ���������_In     In ҩƷ���.���������%Type := 0,
  �ɱ���_In         In ҩƷ���.�ɱ���%Type := 0,
  ��ǰ�ۼ�_In       In �շѼ�Ŀ.�ּ�%Type := 0,
  ����id_In         In �շѼ�Ŀ.������Ŀid%Type := Null,
  ��ͬ��λid_In     In ҩƷ���.��ͬ��λid%Type := Null,
  ˵��_In           In �շ���ĿĿ¼.˵��%Type := Null,
  ��̬����_In       In ҩƷ���.��̬����%Type := 0,
  ��ҩ����_In       In ҩƷ���.��ҩ����%Type := Null,
  ��ѡ��_In         In �շ���ĿĿ¼.��ѡ��%Type := Null,
  ��ֵ˰��_In       In ҩƷ���.��ֵ˰��%Type := Null,
  ����ҩ��_In       In ҩƷ���.����ҩ��%Type := Null,
  վ��_In           In �շ���ĿĿ¼.վ��%Type := Null,
  �Ƿ񳣱�_In       In ҩƷ���.�Ƿ񳣱�%Type := Null,
  �洢�¶�_In       In ��ҺҩƷ����.�洢�¶�%Type := Null,
  �洢����_In       In ��ҺҩƷ����.�洢����%Type := Null,
  ��ҩ����_In       In ��ҺҩƷ����.��ҩ����%Type := Null,
  �Ƿ�������_In   In ��ҺҩƷ����.�Ƿ�������%Type := Null,
  ����_In           In ҩƷ���.����%Type := Null,
  ������Ŀ_In       In �շ���ĿĿ¼.������Ŀ%Type := Null,
  ����ɷ����_In   In ҩƷ���.����ɷ����%Type := 0,
  Dddֵ_In          In ҩƷ���.Dddֵ%Type := 0,
  ��ΣҩƷ_In       In ҩƷ���.��ΣҩƷ%Type := Null,
  �ͻ���λ_In       In ҩƷ���.�ͻ���λ%Type := Null,
  �ͻ���װ_In       In ҩƷ���.�ͻ���װ%Type := Null,
  ��Һע������_In   In ��ҺҩƷ����.��Һע������%Type := Null,
  �Ƿ��ҩ_In       In ҩƷ���.�Ƿ��ҩ%Type := Null,
  �Ƿ����۹���_In In ҩƷ���.�Ƿ����۹���%Type := Null,
  ��λ��_In         In ҩƷ���.��λ��%Type := Null,
  �Ƿ���������_In   In ҩƷ���.�Ƿ���������%Type := Null
) Is

  v_���       ������ĿĿ¼.���%Type;
  v_����       ������ĿĿ¼.����%Type;
  v_Kind       Varchar2(20);
  v_No         �շѼ�Ŀ.No%Type;
  v_Temp       �շ���ĿĿ¼.������Ŀ%Type;
  v_������Ŀ   �շ���ĿĿ¼.������Ŀ%Type;
  n_ָ������� ҩƷ���.ָ�������%Type;

  --�̵�ⷿ�Ĺ������� 
  Cursor c_Storageid Is
    Select Distinct ����id From ��������˵�� Where �������� Like v_Kind Or �������� = '�Ƽ���';
  r_Storageid c_Storageid%RowType;
Begin
  v_������Ŀ := ������Ŀ_In;
  --�жϲ�����Ŀ 
  If v_������Ŀ Is Null Then
    If ����id_In Is Not Null Then
      Begin
        Select ������Ŀ Into v_Temp From ������Ŀ Where ID = ����id_In;
      Exception
        When Others Then
          v_Temp := Null;
      End;
      If v_Temp Is Not Null Then
        v_������Ŀ := v_Temp;
      End If;
    End If;
  End If;
  --�������� 
  Select ���, ���� Into v_���, v_���� From ������ĿĿ¼ Where ID = ҩ��id_In;
  n_ָ������� := (1 - 1 / (1 + �ӳ���_In / 100)) * 100;
  --�����Ϣ 
  Insert Into �շ���ĿĿ¼
    (���, ID, ����, ����, ���, ����, ���㵥λ, ��������, �������, ���ηѱ�, �Ƿ���, ����ʱ��, ����ʱ��, ˵��, ��ѡ��, վ��, ������Ŀ)
  Values
    (v_���, ҩƷid_In, ����_In, v_����, ���_In, ����_In, �ۼ۵�λ_In, ��������_In, �������_In, ���ηѱ�_In, �Ƿ���_In, Sysdate,
     To_Date('3000-01-01', 'YYYY-MM-DD'), ˵��_In, ��ѡ��_In, վ��_In, v_������Ŀ);
  Insert Into ҩƷ���
    (ҩ��id, ҩƷid, ��ʶ��, ҩƷ��Դ, ��׼�ĺ�, ע���̱�, ����ϵ��, ���ﵥλ, �����װ, סԺ��λ, סԺ��װ, ҩ�ⵥλ, ҩ���װ, ���쵥λ, ���췧ֵ, ָ��������, ����, ָ�����ۼ�, ָ�������,
     ����ѱ���, ҩ�ۼ���, �ɱ���, Gmp��֤, �б�ҩƷ, ���������, סԺ�ɷ����, ҩ�����, ҩ������, ���Ч��, ��ͬ��λid, ��̬����, ��ҩ����, ��ֵ˰��, ����ҩ��, �Ƿ񳣱�, ����, ����ɷ����,
     Dddֵ, ��ΣҩƷ, �ͻ���λ, �ͻ���װ, �ӳ���, �Ƿ��ҩ, �Ƿ����۹���, ��λ��, �Ƿ���������)
  Values
    (ҩ��id_In, ҩƷid_In, ��ʶ��_In, ҩƷ��Դ_In, ��׼�ĺ�_In, ע���̱�_In, ����ϵ��_In, ���ﵥλ_In, �����װ_In, סԺ��λ_In, סԺ��װ_In, ҩ�ⵥλ_In, ҩ���װ_In,
     ���쵥λ_In, ���췧ֵ_In, ָ��������_In, ����_In, ָ�����ۼ�_In, n_ָ�������, ����ѱ���_In, ҩ�ۼ���_In, �ɱ���_In, Gmp��֤_In, �б�ҩƷ_In, ���������_In,
     סԺ�ɷ����_In, ҩ�����_In, ҩ������_In, ���Ч��_In, ��ͬ��λid_In, ��̬����_In, ��ҩ����_In, ��ֵ˰��_In, ����ҩ��_In, �Ƿ񳣱�_In, ����_In, ����ɷ����_In,
     Dddֵ_In, ��ΣҩƷ_In, �ͻ���λ_In, �ͻ���װ_In, �ӳ���_In, �Ƿ��ҩ_In, �Ƿ����۹���_In, ��λ��_In, �Ƿ���������_In);

  --�����޸ģ�����ҩƷ������ҩ���г�ҩ��ʱ��ȱʡ�������Ϊ�����סԺ����˽������ҩƷʱ�����ٸ��ݹ��ҩƷ�ķ���������ҩƷ�ķ������ 
  --������Ŀ�������ĸ��� 
  --select nvl(sum(distinct I.�������),0) into v_���� 
  --from �շ���ĿĿ¼ I,ҩƷ��� S 
  --where I.ID=S.ҩƷID and S.ҩ��ID=ҩ��ID_IN; 
  --update ������ĿĿ¼ 
  --set �������=decode(v_����,0,0,1,1,2,2,3) 
  --where ID=ҩ��ID_IN; 

  --�����Ĵ��� 
  Insert Into �շ���Ŀ����
    (�շ�ϸĿid, ����, ����, ����, ����)
    Select ҩƷid_In, ����, ����, ����, ���� From ������Ŀ���� Where ������Ŀid = ҩ��id_In;
  If ������_In Is Not Null Then
    Insert Into �շ���Ŀ���� (�շ�ϸĿid, ����, ����, ����, ����) Values (ҩƷid_In, v_����, 1, ������_In, 3);
  End If;
  If (Ʒ��_In Is Not Null) And (ƴ��_In Is Not Null) Then
    Insert Into �շ���Ŀ���� (�շ�ϸĿid, ����, ����, ����, ����) Values (ҩƷid_In, Ʒ��_In, 3, ƴ��_In, 1);
  End If;
  If (Ʒ��_In Is Not Null) And (���_In Is Not Null) Then
    Insert Into �շ���Ŀ���� (�շ�ϸĿid, ����, ����, ����, ����) Values (ҩƷid_In, Ʒ��_In, 3, ���_In, 2);
  End If;

  --������Ϣ 
  If ����id_In Is Not Null Then
    v_No := Nextno(9);
    Insert Into �շѼ�Ŀ
      (ID, ԭ��id, �շ�ϸĿid, ԭ��, �ּ�, ������Ŀid, �䶯ԭ��, ����˵��, ������, ִ������, ��ֹ����, NO, ���)
    Values
      (�շѼ�Ŀ_Id.Nextval, Null, ҩƷid_In, 0, ��ǰ�ۼ�_In, ����id_In, 1, '��������', User, Sysdate,
       To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
  End If;

  --ҩƷ�����̱Ƚ����� 
  If ����_In Is Not Null Then
    Update ҩƷ������ Set ���� = ����_In Where ���� = ����_In;
    If Sql%RowCount = 0 Then
      Insert Into ҩƷ������
        (����, ����, ����)
        Select Nvl(Max(To_Number(����)), 0) + 1, ����_In, zlSpellCode(����_In, 10) From ҩƷ������;
    End If;
  End If;

  --����ù��ķ������ 
  Insert Into �շ�ִ�п���
    (�շ�ϸĿid, ������Դ, ��������id, ִ�п���id)
    Select ҩƷid_In, ������Դ, ��������id, ִ�п���id From ����ִ�п��� Where ������Ŀid = ҩ��id_In;

  --�����̵����� 

  If v_��� = 5 Then
    v_Kind := '��ҩ%';
  Else
    v_Kind := '��ҩ%';
  End If;

  For r_Storageid In c_Storageid Loop
    Insert Into ҩƷ�����޶�
      (�ⷿid, ҩƷid, ����, ����, �̵�����, �ⷿ��λ)
    Values
      (r_Storageid.����id, ҩƷid_In, 0, 0, '1111', Null);
  End Loop;

  --������ҺҩƷ���� 
  Insert Into ��ҺҩƷ����
    (ҩƷid, �洢�¶�, �洢����, ��ҩ����, �Ƿ�������, ��Һע������)
  Values
    (ҩƷid_In, �洢�¶�_In, �洢����_In, ��ҩ����_In, �Ƿ�������_In, ��Һע������_In);

  --ҩƷ���ȵ���(����ģʽʱ)
  Zl_ҩƷ���ľ���_���۵���;

  b_Message.Zlhis_Dict_035(v_���, ҩƷid_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��ҩ���_Insert;
/



------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0029' Where ���=&n_System;
Commit;