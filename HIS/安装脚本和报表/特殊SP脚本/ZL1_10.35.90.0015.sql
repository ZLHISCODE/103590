----------------------------------------------------------------------------------------------------------------
--���ű�֧�ִ�ZLHIS+ v10.35.90������ v10.35.90
--�������ݿռ�������ߵ�¼PLSQL��ִ�����нű�
Define n_System=100;

----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------



------------------------------------------------------------------------------
--�ṹ��������
------------------------------------------------------------------------------
--122812:����,2018-06-07,Ʊ������¼�����ֶ�����,Ʊ�����ü�¼�����ֶ����ID
Alter Table Ʊ������¼ Add ���� varchar2(20);

Alter Table ���ѿ�����¼ Add ���� varchar2(20);

Alter Table Ʊ�����ü�¼ Add ���ID number(18);

Alter Table ���ѿ����ü�¼ Add ���ID number(18);

Create Index Ʊ������¼_IX_���� On Ʊ������¼(����) Tablespace zl9Indexhis;

Create Index ���ѿ�����¼_IX_���� On ���ѿ�����¼(����) Tablespace zl9Indexhis;

Create Index Ʊ�����ü�¼_IX_���ID On Ʊ�����ü�¼(���ID) Tablespace zl9Indexhis;

Create Index ���ѿ����ü�¼_IX_���ID On ���ѿ����ü�¼(���ID) Tablespace zl9Indexhis;

Alter Table Ʊ�����ü�¼ Add Constraint Ʊ�����ü�¼_FK_���ID foreign Key(���ID) References Ʊ������¼(ID) On Delete Cascade;
Alter Table ���ѿ����ü�¼ Add Constraint ���ѿ����ü�¼_FK_���ID foreign Key(���ID) References ���ѿ�����¼(ID) On Delete Cascade;

------------------------------------------------------------------------------
--������������
------------------------------------------------------------------------------
--122812:����,2018-06-07,Ʊ������¼�����ֶ�����,Ʊ�����ü�¼�����ֶ����ID
Update Ʊ������¼ Set ���� = ID;

Update Ʊ�����ü�¼ A
Set a.���id = a.����
Where Exists (Select 1 From Ʊ������¼ Where ID = a.����);

Update ���ѿ�����¼ Set ���� = ID;

Update ���ѿ����ü�¼ A
Set a.���id = a.����
Where Exists (Select 1 From ���ѿ�����¼ Where ID = a.����);





-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------





-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------




-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--126826:������,2018-06-08,��Ѫҽ��ִ�еǼǲ��Զ����
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
  ��¼��Դ_In     In ����ҽ��ִ��.��¼��Դ%Type := Null
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
    (ҽ��id, ���ͺ�, Ҫ��ʱ��, ��������, ִ��ժҪ, ִ����, ִ��ʱ��, �Ǽ�ʱ��, �Ǽ���, ִ�н��, ˵��, ��Һͨ��, ִ�п���id, ��¼��Դ)
  Values
    (ҽ��id_In, ���ͺ�_In, Ҫ��ʱ��_In, n_��������, ִ��ժҪ_In, ִ����_In, ִ��ʱ��_In, v_Date, v_��Ա����, ִ�н��_In, δִ��ԭ��_In, ��Һͨ��_In, n_ִ�п���id,
     ��¼��Դ_In);

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
    If v_�������� = 2 Then
      If Nvl(����ִ��_In, 0) = 1 Then
        Update סԺ���ü�¼ a
        Set ִ��״̬ = n_ִ��״̬, ִ���� = Decode(n_ִ��״̬, 0, Null, ִ����_In), ִ��ʱ�� = Decode(n_ִ��״̬, 0, Null, ִ��ʱ��_In)
        Where �շ���� Not In ('5', '6', '7') And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In) And Not Exists
         (Select 1 From �������� Where ����id = a.�շ�ϸĿid And �������� = 1) And a.��¼״̬ In (0, 1, 3) And
              (ҽ�����, No, ��¼����) In
              (Select ҽ��id, No, ��¼����
               From ����ҽ������
               Where ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3) And ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In);
      Else
        Update סԺ���ü�¼ a
        Set ִ��״̬ = n_ִ��״̬, ִ���� = Decode(n_ִ��״̬, 0, Null, ִ����_In), ִ��ʱ�� = Decode(n_ִ��״̬, 0, Null, ִ��ʱ��_In)
        Where �շ���� Not In ('5', '6', '7') And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In) And Not Exists
         (Select 1 From �������� Where ����id = a.�շ�ϸĿid And �������� = 1) And a.��¼״̬ In (0, 1, 3) And
              (ҽ�����, No, ��¼����) In (Select ҽ��id, No, ��¼����
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
    Else
      If Nvl(����ִ��_In, 0) = 1 Then
        --�������ﵥ��n_ִ��״̬����Ϊ0���Ǽ�ִ�������ѡ��ִ�н��Ϊδִ�У���������ж�
        Update ������ü�¼ a
        Set ִ��״̬ = n_ִ��״̬, ִ���� = Decode(n_ִ��״̬, 0, Null, ִ����_In), ִ��ʱ�� = Decode(n_ִ��״̬, 0, Null, ִ��ʱ��_In)
        Where �շ���� Not In ('5', '6', '7') And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In) And Not Exists
         (Select 1 From �������� Where ����id = a.�շ�ϸĿid And �������� = 1) And a.��¼״̬ In (0, 1, 3) And
              (ҽ�����, No, ��¼����) In
              (Select ҽ��id, No, ��¼����
               From ����ҽ������
               Where ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3) And ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In);
      Else
        Update ������ü�¼ a
        Set ִ��״̬ = n_ִ��״̬, ִ���� = Decode(n_ִ��״̬, 0, Null, ִ����_In), ִ��ʱ�� = Decode(n_ִ��״̬, 0, Null, ִ��ʱ��_In)
        Where �շ���� Not In ('5', '6', '7') And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In) And Not Exists
         (Select 1 From �������� Where ����id = a.�շ�ϸĿid And �������� = 1) And a.��¼״̬ In (0, 1, 3) And
              (ҽ�����, No, ��¼����) In (Select ҽ��id, No, ��¼����
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

--122812:����,2018-06-07,Ʊ������¼�����ֶ�����,Ʊ�����ü�¼�����ֶ����ID
CREATE OR REPLACE Procedure ZL_Ʊ������¼_INSERT
(
  Id_In       In Ʊ������¼.Id%Type,
  Ʊ��_In     In Ʊ������¼.Ʊ��%Type,
  ʹ�����_In In Ʊ������¼.ʹ�����%Type,
  ǰ׺�ı�_In In Ʊ������¼.ǰ׺�ı�%Type,
  ��ʼ����_In In Ʊ������¼.��ʼ����%Type,
  ��ֹ����_In In Ʊ������¼.��ֹ����%Type,
  �������_In In Ʊ������¼.�������%Type,
  ʣ������_In In Ʊ������¼.ʣ������%Type,
  ��ע_In     In Ʊ������¼.��ע%Type,
  �Ǽ���_In   In Ʊ������¼.�Ǽ���%Type,
  �޸ı�־_In Integer := 0,
  ����_In     In Ʊ������¼.����%Type := Null
) Is
  --�޸ı�־_In:0-����;1-�޸� 
  v_Err_Msg Varchar2(100);
  Err_Item Exception;
  n_Max_Len Number(18);
Begin
  If Nvl(�޸ı�־_In, 0) = 0 Then
    Insert Into Ʊ������¼
      (ID, Ʊ��, ʹ�����, ǰ׺�ı�, ��ʼ����, ��ֹ����, �������, ʣ������, ����Ʊ��, ��ע, �Ǽ���, �Ǽ�ʱ��, ����)
    Values
      (Id_In, Ʊ��_In, ʹ�����_In, ǰ׺�ı�_In, ��ʼ����_In, ��ֹ����_In, �������_In, ʣ������_In, Decode(Sign(Nvl(ʣ������_In, 0)), 1, 1, Null),
       ��ע_In, �Ǽ���_In, Sysdate, Nvl(����_In, Id_In));
    Return;
  End If;

  Begin
    Select Length(Min(��ʼ����))
    Into n_Max_Len
    From (Select Min(��ʼ����) As ��ʼ����
           From Ʊ�ݱ����¼
           Where ���id = Id_In
           Union All
           Select Min(��ʼ����) As ��ʼ���� From Ʊ�����ü�¼ Where ���� = Id_In);
  Exception
    When Others Then
      n_Max_Len := Null;
  End;

  If Not n_Max_Len Is Null Then
    If Length(��ʼ����_In) <> n_Max_Len Then
      v_Err_Msg := '[ZLSOFT]��������Ʊ���Ѿ���ʹ�ù�, ���볤�Ȳ��ܸı�,���볤��Ӧ����' || n_Max_Len || '![ZLSOFT]';
      Raise Err_Item;
    End If;
  End If;

  --�޸� 
  Update Ʊ�����ü�¼
  Set ʹ����� = ʹ�����_In, ���� = ����_In
  Where (Ʊ��, ���id) In (Select Ʊ��, ID From Ʊ������¼ Where ID = Id_In) And Nvl(ʣ������, 0) > 0;

  Update Ʊ������¼
  Set ǰ׺�ı� = ǰ׺�ı�_In, ��ʼ���� = ��ʼ����_In, ��ֹ���� = ��ֹ����_In, ������� = �������_In, ʣ������ = ʣ������_In,
      ����Ʊ�� = Decode(Sign(Nvl(ʣ������_In, 0)), -1, Null, 0, Null, 1), ��ע = ��ע_In, �Ǽ��� = �Ǽ���_In, �Ǽ�ʱ�� = Sysdate,
      ʹ����� = ʹ�����_In, ���� = ����_In
  Where ID = Id_In;
  If Sql%NotFound Then
    v_Err_Msg := '[ZLSOFT]�����Ʊ��δ�ҵ�,�����Ѿ�������ɾ��,�����޸�![ZLSOFT]';
    Raise Err_Item;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Ʊ������¼_Insert;
/

--122812:����,2018-06-07,Ʊ������¼�����ֶ�����,Ʊ�����ü�¼�����ֶ����ID
CREATE OR REPLACE Procedure Zl_���ѿ�����¼_Insert
(
  Id_In       In ���ѿ�����¼.Id%Type,
  �ӿڱ��_In In ���ѿ�����¼.�ӿڱ��%Type,
  ǰ׺�ı�_In In ���ѿ�����¼.ǰ׺�ı�%Type,
  ��ʼ����_In In ���ѿ�����¼.��ʼ����%Type,
  ��ֹ����_In In ���ѿ�����¼.��ֹ����%Type,
  �������_In In ���ѿ�����¼.�������%Type,
  ʣ������_In In ���ѿ�����¼.ʣ������%Type,
  ��ע_In     In ���ѿ�����¼.��ע%Type,
  �Ǽ���_In   In ���ѿ�����¼.�Ǽ���%Type,
  �޸ı�־_In Integer := 0,
  ����_In     In ���ѿ�����¼.����%Type := Null
) Is
  --�޸ı�־_In:0-����;1-�޸�
  v_Err_Msg Varchar2(100);
  Err_Item Exception;

  n_Max_Len Number(18);
Begin
  If Nvl(�޸ı�־_In, 0) = 0 Then
    Insert Into ���ѿ�����¼
      (ID, �ӿڱ��, ǰ׺�ı�, ��ʼ����, ��ֹ����, �������, ʣ������, �Ƿ���ڿ�, ��ע, �Ǽ���, �Ǽ�ʱ��, ����)
    Values
      (Id_In, �ӿڱ��_In, ǰ׺�ı�_In, ��ʼ����_In, ��ֹ����_In, �������_In, ʣ������_In, Decode(Sign(Nvl(ʣ������_In, 0)), 1, 1, 0), ��ע_In,
       �Ǽ���_In, Sysdate, Nvl(����_In, Id_In));
    Return;
  End If;

  Begin
    Select Length(Min(��ʼ����))
    Into n_Max_Len
    From (Select Min(��ʼ����) As ��ʼ����
           From ���ѿ������¼
           Where ���id = Id_In
           Union All
           Select Min(��ʼ����) As ��ʼ���� From ���ѿ����ü�¼ Where ���� = Id_In);
  Exception
    When Others Then
      n_Max_Len := Null;
  End;

  If Not n_Max_Len Is Null Then
    If Length(��ʼ����_In) <> n_Max_Len Then
      v_Err_Msg := '������ⵥ�Ѿ���ʹ�ù������ų��Ȳ��ܸı䣬���ų���Ӧ����' || n_Max_Len || '��';
      Raise Err_Item;
    End If;
  End If;

  --�޸�
  Update ���ѿ����ü�¼ Set �ӿڱ�� = �ӿڱ��_In, ���� = ����_In Where ���id = Id_In And Nvl(ʣ������, 0) > 0;

  Update ���ѿ�����¼
  Set ǰ׺�ı� = ǰ׺�ı�_In, ��ʼ���� = ��ʼ����_In, ��ֹ���� = ��ֹ����_In, ������� = �������_In, ʣ������ = ʣ������_In,
      �Ƿ���ڿ� = Decode(Sign(Nvl(ʣ������_In, 0)), 1, 1, 0), ��ע = ��ע_In, �Ǽ��� = �Ǽ���_In, �Ǽ�ʱ�� = Sysdate, �ӿڱ�� = �ӿڱ��_In,
      ���� = ����_In
  Where ID = Id_In;
  If Sql%NotFound Then
    v_Err_Msg := '����ⵥδ�ҵ��������Ѿ�������ɾ���������޸ģ�';
    Raise Err_Item;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���ѿ�����¼_Insert;
/

--122812:����,2018-06-07,Ʊ������¼�����ֶ�����,Ʊ�����ü�¼�����ֶ����ID
CREATE OR REPLACE Procedure Zl_Ʊ�����ü�¼_Insert
(
  Id_In       In Ʊ�����ü�¼.Id%Type,
  Ʊ��_In     In Ʊ�����ü�¼.Ʊ��%Type,
  ʹ�����_In In Ʊ�����ü�¼.ʹ�����%Type,
  ������_In   In Ʊ�����ü�¼.������%Type,
  ǰ׺�ı�_In In Ʊ�����ü�¼.ǰ׺�ı�%Type,
  ��ʼ����_In In Ʊ�����ü�¼.��ʼ����%Type,
  ��ֹ����_In In Ʊ�����ü�¼.��ֹ����%Type,
  ʹ�÷�ʽ_In In Ʊ�����ü�¼.ʹ�÷�ʽ%Type,
  �Ǽ�ʱ��_In In Ʊ�����ü�¼.�Ǽ�ʱ��%Type := Null,
  �Ǽ���_In   In Ʊ�����ü�¼.�Ǽ���%Type := Null,
  ʣ������_In In Ʊ�����ü�¼.ʣ������%Type := Null,
  ����_In     In Ʊ�����ü�¼.����%Type := Null,
  ǩ����_In   In Ʊ�����ü�¼.ǩ����%Type := Null,
  ���id_In   In Ʊ�����ü�¼.���id%Type := Null
) Is
  v_Err_Msg Varchar2(500);
  Err_Item Exception;
  n_Count  Number(18);
  n_ʣ���� Ʊ������¼.ʣ������%Type;
Begin

  For v_��� In (Select ID, ǰ׺�ı�, ʹ�����, ��ʼ����, Nvl(��ֹ����, ��ʼ����) As ��ֹ����
               From Ʊ������¼
               Where ID = Nvl(���id_In, 0) And Ʊ�� = Ʊ��_In) Loop
    --1. �����
    If ��ʼ����_In < v_���.��ʼ���� Or ��ʼ����_In > v_���.��ֹ���� Then
      --������ⷶΧ,��������
      v_Err_Msg := '[ZLSOFT]��ǰ���õĿ�ʼ���롺' || ��ʼ����_In || '��������ⷶΧ' || Chr(10) || Chr(13) || '��' || v_���.��ʼ���� || '-' ||
                   v_���.��ֹ���� || '���������ø�Ʊ��![ZLSOFT]';
      Raise Err_Item;
    End If;
  
    If ��ֹ����_In < v_���.��ʼ���� Or ��ֹ����_In > v_���.��ֹ���� Then
      --������ⷶΧ,��������
      v_Err_Msg := '[ZLSOFT]��ǰ���õ���ֹ���롺' || ��ֹ����_In || '��������ⷶΧ' || Chr(10) || Chr(13) || '��' || v_���.��ʼ���� || '-' ||
                   v_���.��ֹ���� || '���������ø�Ʊ��![ZLSOFT]';
      Raise Err_Item;
    End If;
    If Nvl(v_���.ʹ�����, 'LXH') <> Nvl(ʹ�����_In, 'LXH') Then
      v_Err_Msg := '[ZLSOFT]����ʹ�����' || Nvl(v_���.ʹ�����, '') || '�������õ����' || Nvl(ʹ�����_In, '') || '����һ��![ZLSOFT]';
      Raise Err_Item;
    End If;
  
    --2.���Ʊ���Ƿ��Ѿ�������,�����ظ�����
    Select Count(*)
    Into n_Count
    From Ʊ�ݱ����¼
    Where ���id = Nvl(���id_In, 0) And ((��ʼ����_In Between ��ʼ���� And ��ֹ����) Or (��ֹ����_In Between ��ʼ���� And ��ֹ����) Or
          (��ʼ���� Between ��ʼ����_In And ��ֹ����_In) Or (��ֹ���� Between ��ʼ����_In And ��ֹ����_In));
  
    If n_Count <> 0 Then
      If ��ʼ����_In = ��ֹ����_In Then
        v_Err_Msg := '[ZLSOFT]����:' || ��ʼ����_In || '�ڱ����¼���Ѿ�����,�����ٽ�������![ZLSOFT]';
      Else
        v_Err_Msg := '[ZLSOFT]����:' || ��ʼ����_In || '-' || ��ֹ����_In || '�ڱ����¼���Ѿ�����,�����ٽ�������![ZLSOFT]';
      End If;
      Raise Err_Item;
    End If;
  
    --3.���Ʊ���Ƿ��Ѿ�������,���õĲ����ٽ��б���
    Select Count(*)
    Into n_Count
    From Ʊ�����ü�¼
    Where ���� = Nvl(����_In, 0) And ((��ʼ����_In Between ��ʼ���� And ��ֹ����) Or (��ֹ����_In Between ��ʼ���� And ��ֹ����) Or
          (��ʼ���� Between ��ʼ����_In And ��ֹ����_In) Or (��ֹ���� Between ��ʼ����_In And ��ֹ����_In));
    If n_Count <> 0 Then
      If ��ʼ����_In = ��ֹ����_In Then
        v_Err_Msg := '[ZLSOFT]����:' || ��ʼ����_In || '��Ʊ�����ü�¼���Ѿ�����,�����ٽ�������![ZLSOFT]';
      Else
        v_Err_Msg := '[ZLSOFT]����:' || ��ʼ����_In || '-' || ��ֹ����_In || '��Ʊ�����ü�¼���Ѿ�����,�����ٽ�������![ZLSOFT]';
      End If;
      Raise Err_Item;
    End If;
    --���ٿ��
    Update Ʊ������¼
    Set ʣ������ = Nvl(ʣ������, 0) - Nvl(ʣ������_In, 0)
    Where ID = Nvl(���id_In, 0) And Ʊ�� = Ʊ��_In And Nvl(ʹ�����, 'LXH') = Nvl(ʹ�����_In, 'LXH')
    Returning Nvl(ʣ������, 0) Into n_ʣ����;
  
    If n_ʣ���� < 0 Then
      v_Err_Msg := '[ZLSOFT]���Ʊ�ݵ�ʣ��Ʊ��������,����![ZLSOFT]';
      Raise Err_Item;
    End If;
    If n_ʣ���� = 0 Then
      Update Ʊ������¼
      Set ����Ʊ�� = Null
      Where ID = Nvl(���id_In, 0) And Ʊ�� = Ʊ��_In And Nvl(ʹ�����, 'LXH') = Nvl(ʹ�����_In, 'LXH');
    End If;
  End Loop;

  Insert Into Ʊ�����ü�¼
    (ID, Ʊ��, ʹ�����, ������, ǰ׺�ı�, ��ʼ����, ��ֹ����, ʹ�÷�ʽ, �Ǽ�ʱ��, �Ǽ���, ʣ������, ����, ǩ����, ǩ��ʱ��, ���id)
  Values
    (Id_In, Ʊ��_In, ʹ�����_In, ������_In, ǰ׺�ı�_In, ��ʼ����_In, ��ֹ����_In, ʹ�÷�ʽ_In, �Ǽ�ʱ��_In, �Ǽ���_In, ʣ������_In, ����_In, ǩ����_In,
     Decode(ǩ����_In, Null, Null + Sysdate, Sysdate), ���id_In);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Ʊ�����ü�¼_Insert;
/

--122812:����,2018-06-07,Ʊ������¼�����ֶ�����,Ʊ�����ü�¼�����ֶ����ID
CREATE OR REPLACE Procedure Zl_���ѿ����ü�¼_Insert
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
  ���ID_In   ���ѿ����ü�¼.���ID%Type := Null
) Is
  v_Err_Msg Varchar2(500);
  Err_Item Exception;

  n_Count  Number(18);
  n_ʣ���� ���ѿ�����¼.ʣ������%Type;
Begin

  For r_��� In (Select ID, ǰ׺�ı�, �ӿڱ��, ��ʼ����, Nvl(��ֹ����, ��ʼ����) As ��ֹ����
               From ���ѿ�����¼
               Where ID = Nvl(���ID_In, 0)) Loop
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
    Where ���id = Nvl(���ID_In, 0) And ((��ʼ����_In Between ��ʼ���� And ��ֹ����) Or (��ֹ����_In Between ��ʼ���� And ��ֹ����) Or
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
    Where ID = Nvl(���ID_In, 0) And �ӿڱ�� = �ӿڱ��_In
    Returning Nvl(ʣ������, 0) Into n_ʣ����;

    If n_ʣ���� < 0 Then
      v_Err_Msg := '��⿨Ƭ��ʣ��Ʊ�������㣬���飡';
      Raise Err_Item;
    End If;
    If n_ʣ���� = 0 Then
      Update ���ѿ�����¼ Set �Ƿ���ڿ� = 0 Where ID = Nvl(���ID_In, 0) And �ӿڱ�� = �ӿڱ��_In;
    End If;
  End Loop;

  Insert Into ���ѿ����ü�¼
    (ID, �ӿڱ��, ������, ǰ׺�ı�, ��ʼ����, ��ֹ����, ʹ�÷�ʽ, �Ǽ�ʱ��, �Ǽ���, ʣ������, ����, ǩ����, ǩ��ʱ��)
  Values
    (Id_In, �ӿڱ��_In, ������_In, ǰ׺�ı�_In, ��ʼ����_In, ��ֹ����_In, ʹ�÷�ʽ_In, �Ǽ�ʱ��_In, �Ǽ���_In, ʣ������_In, ����_In, ǩ����_In,
     Decode(ǩ����_In, Null, Null + Sysdate, Sysdate));
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���ѿ����ü�¼_Insert;
/

--122812:����,2018-06-07,Ʊ������¼�����ֶ�����,Ʊ�����ü�¼�����ֶ����ID
CREATE OR REPLACE Procedure ZL_Ʊ�����ü�¼_UPDATE
(
  Id_In       In Ʊ�����ü�¼.Id%Type,
  ʹ�����_In In Ʊ�����ü�¼.ʹ�����%Type,
  ������_In   In Ʊ�����ü�¼.������%Type,
  ��ʼ����_In In Ʊ�����ü�¼.��ʼ����%Type,
  ��ֹ����_In In Ʊ�����ü�¼.��ֹ����%Type,
  ǰ׺�ı�_In In Ʊ�����ü�¼.ǰ׺�ı�%Type := Null,
  ʹ�÷�ʽ_In In Ʊ�����ü�¼.ʹ�÷�ʽ%Type := 1,
  �Ǽ�ʱ��_In In Ʊ�����ü�¼.�Ǽ�ʱ��%Type := Null,
  �Ǽ���_In   In Ʊ�����ü�¼.�Ǽ���%Type := Null,
  ����_In     In Ʊ�����ü�¼.����%Type := Null,
  ǩ����_In   In Ʊ�����ü�¼.ǩ����%Type := Null,
  ���id_In   In Ʊ�����ü�¼.���id%Type := Null
  
) Is
  Cursor c_���ü�¼ Is
    Select * From Ʊ�����ü�¼ Where ID = Id_In For Update;

  c_��¼     Ʊ�����ü�¼%RowType;
  n_ʹ������ Ʊ�����ü�¼.ʣ������%Type;
  n_ʣ������ Ʊ�����ü�¼.ʣ������%Type;
  n_ԭ������ Ʊ�����ü�¼.ʣ������%Type;
  n_�������� Ʊ�����ü�¼.ʣ������%Type;

  v_��ʼ���� Ʊ�����ü�¼.��ʼ����%Type;
  v_��ֹ���� Ʊ�����ü�¼.��ֹ����%Type;
  v_Err_Msg  Varchar2(500);
  Err_Item Exception;
  n_Count  Number(18);
  n_ʣ���� Ʊ������¼.ʣ������%Type;
Begin
  Open c_���ü�¼;
  Fetch c_���ü�¼
    Into c_��¼;

  If c_���ü�¼%NotFound Then
    --��¼δ�ҵ� 
    v_Err_Msg := '[ZLSOFT]������¼�Ѿ���ɾ���������޸ġ�[ZLSOFT]';
    Raise Err_Item;
  End If;

  Select Min(����), Max(����) Into v_��ʼ����, v_��ֹ���� From Ʊ��ʹ����ϸ Where ����id = Id_In;

  If ǰ׺�ı�_In Is Null Then
    n_ʣ������ := To_Number(��ֹ����_In) - To_Number(��ʼ����_In) + 1;
  Else
    n_ʣ������ := To_Number(Substr(��ֹ����_In, Length(ǰ׺�ı�_In) + 1)) - To_Number(Substr(��ʼ����_In, Length(ǰ׺�ı�_In) + 1)) + 1;
  End If;

  n_�������� := n_ʣ������;
  If c_��¼.ǰ׺�ı� Is Null Then
    n_ԭ������ := To_Number(c_��¼.��ֹ����) - To_Number(c_��¼.��ʼ����) + 1;
  Else
    n_ԭ������ := To_Number(Substr(c_��¼.��ֹ����, Length(c_��¼.ǰ׺�ı�) + 1)) - To_Number(Substr(c_��¼.��ʼ����, Length(c_��¼.ǰ׺�ı�) + 1)) + 1;
  End If;

  If v_��ʼ���� Is Not Null Then
    --�Ѿ�ʹ�ã���һЩ��Ŀ������֤ 
    If Nvl(ǰ׺�ı�_In, ' ') <> Nvl(c_��¼.ǰ׺�ı�, ' ') Then
      v_Err_Msg := '[ZLSOFT]������¼���õ�Ʊ���Ѿ�ʹ�ã������޸ĺ����ǰ׺��[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    If Length(��ʼ����_In) <> Length(c_��¼.��ʼ����) Then
      v_Err_Msg := '[ZLSOFT]������¼���õ�Ʊ���Ѿ�ʹ�ã������޸ĺ���ĳ��ȡ�[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    If ��ʼ����_In > v_��ʼ���� Then
      v_Err_Msg := '[ZLSOFT]������¼���õ�Ʊ���Ѿ�ʹ�ã���ʼ�������ֻ����' || v_��ʼ���� || '��[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    If ��ֹ����_In < v_��ֹ���� Then
      v_Err_Msg := '[ZLSOFT]������¼���õ�Ʊ���Ѿ�ʹ�ã���ֹ������Сֻ����' || v_��ֹ���� || '��[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    --����������� 
    If ǰ׺�ı�_In Is Null Then
      n_ʹ������ := To_Number(c_��¼.��ֹ����) - To_Number(c_��¼.��ʼ����) + 1 - c_��¼.ʣ������;
    Else
      n_ʹ������ := To_Number(Substr(c_��¼.��ֹ����, Length(ǰ׺�ı�_In) + 1)) - To_Number(Substr(c_��¼.��ʼ����, Length(ǰ׺�ı�_In) + 1)) + 1 -
                c_��¼.ʣ������;
    End If;
  
    n_ʣ������ := n_ʣ������ - n_ʹ������;
  End If;

  For v_��� In (Select ID, ǰ׺�ı�, ʹ�����, ��ʼ����, Nvl(��ֹ����, ��ʼ����) As ��ֹ����
               From Ʊ������¼
               Where ID = Nvl(���id_In, 0) And Ʊ�� = c_��¼.Ʊ��) Loop
  
    If Nvl(ʹ�����_In, 'LXH') <> Nvl(v_���.ʹ�����, 'LXH') Then
      v_Err_Msg := '[ZLSOFT]��ǰ���õ�ʹ�����' || Nvl(ʹ�����_In, '') || '��������ʹ�����һ�¡�' || Nvl(v_���.ʹ�����, '') || '��![ZLSOFT]';
      Raise Err_Item;
    End If;
  
    --1. ����� 
    If ��ʼ����_In < v_���.��ʼ���� Or ��ʼ����_In > v_���.��ֹ���� Then
      --������ⷶΧ,�������� 
      v_Err_Msg := '[ZLSOFT]��ǰ���õĿ�ʼ���롺' || ��ʼ����_In || '��������ⷶΧ' || Chr(10) || Chr(13) || '��' || v_���.��ʼ���� || '-' ||
                   v_���.��ֹ���� || '���������ø�Ʊ��![ZLSOFT]';
      Raise Err_Item;
    End If;
    If ��ֹ����_In < v_���.��ʼ���� Or ��ֹ����_In > v_���.��ֹ���� Then
      --������ⷶΧ,�������� 
      v_Err_Msg := '[ZLSOFT]��ǰ���õ���ֹ���롺' || ��ֹ����_In || '��������ⷶΧ' || Chr(10) || Chr(13) || '��' || v_���.��ʼ���� || '-' ||
                   v_���.��ֹ���� || '���������ø�Ʊ��![ZLSOFT]';
      Raise Err_Item;
    End If;
    --2.���Ʊ���Ƿ��Ѿ�������,�����ظ����� 
    Select Count(*)
    Into n_Count
    From Ʊ�ݱ����¼
    Where ���id = Nvl(���id_In, 0) And ((��ʼ����_In Between ��ʼ���� And ��ֹ����) Or (��ֹ����_In Between ��ʼ���� And ��ֹ����) Or
          (��ʼ���� Between ��ʼ����_In And ��ֹ����_In) Or (��ֹ���� Between ��ʼ����_In And ��ֹ����_In));
  
    If n_Count <> 0 Then
      If ��ʼ����_In = ��ֹ����_In Then
        v_Err_Msg := '[ZLSOFT]����:' || ��ʼ����_In || '�ڱ����¼���Ѿ�����,�����ٽ�������![ZLSOFT]';
      Else
        v_Err_Msg := '[ZLSOFT]����:' || ��ʼ����_In || '-' || ��ֹ����_In || '�ڱ����¼���Ѿ�����,�����ٽ�������![ZLSOFT]';
      End If;
      Raise Err_Item;
    End If;
  
    --3.���Ʊ���Ƿ��Ѿ�������,���õĲ����ٽ��б��� 
    Select Count(*)
    Into n_Count
    From Ʊ�����ü�¼
    Where ���� = Nvl(����_In, 0) And ID <> Id_In And
          ((��ʼ����_In Between ��ʼ���� And ��ֹ����) Or (��ֹ����_In Between ��ʼ���� And ��ֹ����) Or (��ʼ���� Between ��ʼ����_In And ��ֹ����_In) Or
          (��ֹ���� Between ��ʼ����_In And ��ֹ����_In));
    If n_Count <> 0 Then
      If ��ʼ����_In = ��ֹ����_In Then
        v_Err_Msg := '[ZLSOFT]����:' || ��ʼ����_In || '��Ʊ�����ü�¼���Ѿ�����,�����ٽ�������![ZLSOFT]';
      Else
        v_Err_Msg := '[ZLSOFT]����:' || ��ʼ����_In || '-' || ��ֹ����_In || '��Ʊ�����ü�¼���Ѿ�����,�����ٽ�������![ZLSOFT]';
      End If;
      Raise Err_Item;
    End If;
  
    --���ٿ�� 
    Update Ʊ������¼
    Set ʣ������ = Nvl(ʣ������, 0) + (Nvl(n_ԭ������, 0) - Nvl(n_��������, 0))
    Where ID = Nvl(���id_In, 0) And Ʊ�� = c_��¼.Ʊ��
    Returning Nvl(ʣ������, 0) Into n_ʣ����;
    If n_ʣ���� < 0 Then
      v_Err_Msg := '[ZLSOFT]���Ʊ�ݵ�ʣ��Ʊ��������,����![ZLSOFT]';
      Raise Err_Item;
    End If;
    Update Ʊ������¼
    Set ����Ʊ�� = Decode(Sign(Nvl(n_ʣ����, 0)), 1, 1, Null)
    Where ID = (Select ���id From Ʊ�����ü�¼ Where ID = Id_In) And Ʊ�� = c_��¼.Ʊ��;
  End Loop;

  Update Ʊ�����ü�¼
  Set ������ = ������_In, ǰ׺�ı� = ǰ׺�ı�_In, ��ʼ���� = ��ʼ����_In, ��ֹ���� = ��ֹ����_In, ʹ�÷�ʽ = ʹ�÷�ʽ_In, �Ǽ�ʱ�� = �Ǽ�ʱ��_In, �Ǽ��� = �Ǽ���_In,
      ʣ������ = n_ʣ������, ���� = ����_In, ʹ����� = ʹ�����_In, ǩ���� = ǩ����_In, ǩ��ʱ�� = Decode(ǩ����_In, Null, Null + Sysdate, Sysdate)
  Where ID = Id_In;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Ʊ�����ü�¼_Update;
/

--122812:����,2018-06-07,Ʊ������¼�����ֶ�����,Ʊ�����ü�¼�����ֶ����ID
CREATE OR REPLACE Procedure Zl_���ѿ����ü�¼_Update
(
  Id_In       ���ѿ����ü�¼.Id%Type,
  �ӿڱ��_In ���ѿ����ü�¼.�ӿڱ��%Type,
  ������_In   ���ѿ����ü�¼.������%Type,
  ��ʼ����_In ���ѿ����ü�¼.��ʼ����%Type,
  ��ֹ����_In ���ѿ����ü�¼.��ֹ����%Type,
  ǰ׺�ı�_In ���ѿ����ü�¼.ǰ׺�ı�%Type := Null,
  ʹ�÷�ʽ_In ���ѿ����ü�¼.ʹ�÷�ʽ%Type := 1,
  �Ǽ�ʱ��_In ���ѿ����ü�¼.�Ǽ�ʱ��%Type := Null,
  �Ǽ���_In   ���ѿ����ü�¼.�Ǽ���%Type := Null,
  ����_In     ���ѿ����ü�¼.����%Type := Null,
  ǩ����_In   ���ѿ����ü�¼.ǩ����%Type := Null,
  ���id_In   ���ѿ����ü�¼.���id%Type := Null
) Is
  Cursor c_���ü�¼ Is
    Select �ӿڱ��, ǰ׺�ı�, ��ʼ����, ��ֹ����, ʣ������ From ���ѿ����ü�¼ Where ID = Id_In For Update;

  c_��¼     c_���ü�¼%RowType;
  n_ʣ������ ���ѿ����ü�¼.ʣ������%Type;
  n_ԭ������ ���ѿ����ü�¼.ʣ������%Type;
  n_�������� ���ѿ����ü�¼.ʣ������%Type;

  v_��ʼ���� ���ѿ����ü�¼.��ʼ����%Type;
  v_��ֹ���� ���ѿ����ü�¼.��ֹ����%Type;
  n_ʣ����   ���ѿ�����¼.ʣ������%Type;
  n_Count    Number(18);

  v_Err_Msg Varchar2(500);
  Err_Item Exception;
Begin
  Open c_���ü�¼;
  Fetch c_���ü�¼
    Into c_��¼;

  If c_���ü�¼%NotFound Then
    --��¼δ�ҵ�
    v_Err_Msg := '������¼�Ѿ���ɾ���������޸ġ�';
    Raise Err_Item;
  End If;

  Select Min(����), Max(����) Into v_��ʼ����, v_��ֹ���� From ���ѿ�ʹ�ü�¼ Where ����id = Id_In;

  If ǰ׺�ı�_In Is Null Then
    n_ʣ������ := To_Number(��ֹ����_In) - To_Number(��ʼ����_In) + 1;
  Else
    n_ʣ������ := To_Number(Substr(��ֹ����_In, Length(ǰ׺�ı�_In) + 1)) - To_Number(Substr(��ʼ����_In, Length(ǰ׺�ı�_In) + 1)) + 1;
  End If;

  n_�������� := n_ʣ������;
  If c_��¼.ǰ׺�ı� Is Null Then
    n_ԭ������ := To_Number(c_��¼.��ֹ����) - To_Number(c_��¼.��ʼ����) + 1;
  Else
    n_ԭ������ := To_Number(Substr(c_��¼.��ֹ����, Length(c_��¼.ǰ׺�ı�) + 1)) - To_Number(Substr(c_��¼.��ʼ����, Length(c_��¼.ǰ׺�ı�) + 1)) + 1;
  End If;

  If v_��ʼ���� Is Not Null Then
    --�Ѿ�ʹ�ã���һЩ��Ŀ������֤
    If Nvl(ǰ׺�ı�_In, '-') <> Nvl(c_��¼.ǰ׺�ı�, '-') Then
      v_Err_Msg := '������¼���õĿ����Ѿ�ʹ�ã������޸Ŀ��ŵ�ǰ׺��';
      Raise Err_Item;
    End If;
  
    If Length(��ʼ����_In) <> Length(c_��¼.��ʼ����) Then
      v_Err_Msg := '������¼���õĿ����Ѿ�ʹ�ã������޸Ŀ��ŵĳ��ȡ�';
      Raise Err_Item;
    End If;
  
    If ��ʼ����_In > v_��ʼ���� Then
      v_Err_Msg := '������¼���õĿ����Ѿ�ʹ�ã���ʼ�������ֻ����' || v_��ʼ���� || '��';
      Raise Err_Item;
    End If;
  
    If ��ֹ����_In < v_��ֹ���� Then
      v_Err_Msg := '������¼���õĿ����Ѿ�ʹ�ã���ֹ������Сֻ����' || v_��ֹ���� || '��';
      Raise Err_Item;
    End If;
  
    --�����������
    n_ʣ������ := n_ʣ������ - (n_ԭ������ - c_��¼.ʣ������);
  End If;

  For v_��� In (Select ID, ǰ׺�ı�, �ӿڱ��, ��ʼ����, Nvl(��ֹ����, ��ʼ����) As ��ֹ����
               From ���ѿ�����¼
               Where ID = Nvl(���id_In, 0)) Loop
  
    If �ӿڱ��_In <> v_���.�ӿڱ�� Then
      v_Err_Msg := '��ǰ���õĿ����' || Nvl(�ӿڱ��_In, '') || '�����������һ�¡�' || Nvl(v_���.�ӿڱ��, '') || '��!';
      Raise Err_Item;
    End If;
  
    --1. �����
    If ��ʼ����_In < v_���.��ʼ���� Or ��ʼ����_In > v_���.��ֹ���� Then
      --������ⷶΧ,��������
      v_Err_Msg := '��ǰ���õĿ�ʼ���š�' || ��ʼ����_In || '��������ⷶΧ' || Chr(10) || Chr(13) || '��' || v_���.��ʼ���� || '-' || v_���.��ֹ���� ||
                   '���������øÿ�Ƭ��';
      Raise Err_Item;
    End If;
    If ��ֹ����_In < v_���.��ʼ���� Or ��ֹ����_In > v_���.��ֹ���� Then
      --������ⷶΧ,��������
      v_Err_Msg := '��ǰ���õ���ֹ���š�' || ��ֹ����_In || '��������ⷶΧ' || Chr(10) || Chr(13) || '��' || v_���.��ʼ���� || '-' || v_���.��ֹ���� ||
                   '���������øÿ�Ƭ��';
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
    Select Count(*)
    Into n_Count
    From ���ѿ����ü�¼
    Where ���� = Nvl(����_In, 0) And ID <> Id_In And
          ((��ʼ����_In Between ��ʼ���� And ��ֹ����) Or (��ֹ����_In Between ��ʼ���� And ��ֹ����) Or (��ʼ���� Between ��ʼ����_In And ��ֹ����_In) Or
          (��ֹ���� Between ��ʼ����_In And ��ֹ����_In));
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
    Set ʣ������ = Nvl(ʣ������, 0) + (Nvl(n_ԭ������, 0) - Nvl(n_��������, 0))
    Where ID = (Select ���id From ���ѿ����ü�¼ Where ID = Id_In)
    Returning Nvl(ʣ������, 0) Into n_ʣ����;
    If n_ʣ���� < 0 Then
      v_Err_Msg := '��⿨Ƭ��ʣ��Ʊ�������㣬���飡';
      Raise Err_Item;
    End If;
  
    Update ���ѿ�����¼ Set �Ƿ���ڿ� = Decode(Sign(Nvl(n_ʣ����, 0)), 1, 1, Null) Where ID = Nvl(���id_In, 0);
  End Loop;

  Update ���ѿ����ü�¼
  Set ������ = ������_In, ǰ׺�ı� = ǰ׺�ı�_In, ��ʼ���� = ��ʼ����_In, ��ֹ���� = ��ֹ����_In, ʹ�÷�ʽ = ʹ�÷�ʽ_In, �Ǽ�ʱ�� = �Ǽ�ʱ��_In, �Ǽ��� = �Ǽ���_In,
      ʣ������ = n_ʣ������, ���� = Nvl(����_In, 0), �ӿڱ�� = �ӿڱ��_In, ǩ���� = ǩ����_In,
      ǩ��ʱ�� = Decode(ǩ����_In, Null, Null + Sysdate, Sysdate)
  Where ID = Id_In;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���ѿ����ü�¼_Update;
/

CREATE OR REPLACE Procedure Zl_Ris���ԤԼ_Delete(ҽ��id_In In Ris���ԤԼ.ҽ��id%Type) Is
  v_ԤԼid       Ris���ԤԼ.ԤԼid%Type;
  v_ԤԼ����     Ris���ԤԼ.ԤԼ����%Type;
  v_ԤԼ���     Ris���ԤԼ.���%Type;
  v_����豸���� Ris���ԤԼ.����豸����%Type;
Begin
  v_ԤԼid := 0;
  Begin
    Select ԤԼid, ԤԼ����, ���, ����豸����
    Into v_ԤԼid, v_ԤԼ����, v_ԤԼ���, v_����豸����
    From Ris���ԤԼ
    Where ҽ��id = ҽ��id_In;
  Exception
    When Others Then
      Null;
  End;

  Delete Ris���ԤԼ Where ҽ��id = ҽ��id_In;

  --������Ϣ
  If v_ԤԼid <> 0 Then
    b_Message.Zlhis_Pacs_007(ҽ��id_In, v_ԤԼid, v_ԤԼ����, v_ԤԼ���, v_����豸����);
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Ris���ԤԼ_Delete;
/





------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0015' Where ���=&n_System;
Commit;