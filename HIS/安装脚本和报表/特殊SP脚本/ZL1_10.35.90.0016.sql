----------------------------------------------------------------------------------------------------------------
--���ű�֧�ִ�ZLHIS+ v10.35.90������ v10.35.90
--�������ݿռ�������ߵ�¼PLSQL��ִ�����нű�
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------



------------------------------------------------------------------------------
--�ṹ��������
------------------------------------------------------------------------------
--127220:������,2018-06-12,����ҽ��ִ��ʱ�Ĳ�����Դ
Alter Table ����ҽ��ִ�� Add ִ�з�ʽ Number(1);




------------------------------------------------------------------------------
--������������
------------------------------------------------------------------------------




-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------
--126262:����,2018-06-13,������ҩ�Ͳ��ŷ�ҩ�������Ӳ�������
Insert Into zlModuleRelas(ϵͳ,ģ��,����,���ϵͳ,���ģ��,�������,��ع���,ȱʡֵ)
Select 100,1341,A.* From (
Select ����,���ϵͳ,���ģ��,�������,��ع���,ȱʡֵ From zlModuleRelas Where 1 = 0
Union All Select NULL,100,1259,1,NULL,1 From Dual) A;

Insert Into zlModuleRelas(ϵͳ,ģ��,����,���ϵͳ,���ģ��,�������,��ع���,ȱʡֵ)
Select 100,1342,A.* From (
Select ����,���ϵͳ,���ģ��,�������,��ع���,ȱʡֵ From zlModuleRelas Where 1 = 0
Union All Select NULL,100,1259,1,NULL,1 From Dual) A;

Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��,ȱʡֵ)
Select  &n_System,1341,A.* From (
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0 
Union All Select '���Ӳ�������',30,'�и�Ȩ��ʱ�����Խ��е��Ӳ�������',1 From Dual) A;

Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��,ȱʡֵ)
Select  &n_System,1342,A.* From (
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0 
Union All Select '���Ӳ�������',23,'�и�Ȩ��ʱ�����Խ��е��Ӳ�������',1 From Dual) A;

-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------




-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--127220:������,2018-06-12,����ҽ��ִ��ʱ�Ĳ�����Դ
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
--127220:������,2018-06-12,����ҽ��ִ��ʱ�Ĳ�����Դ
Create Or Replace Procedure Zl_Third_Adviceoperation
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --���ܣ�ҽ��ִ�еǼ�/ȡ��ִ�еǼ� /д��
  --1������ÿ��ҽ��ִ��ʱ��ִ�еǼǣ�����Դ�����������Ƿ��ƶ���ִ��
  --2������ִ�е�ҽ��ȡ��ִ�в���
  --��Σ�xml_in
  --<IN>
  -- <TYPE>1</TYPE>   --�������ͣ�1��ִ�еǼǣ�0��ȡ��ִ�еǼ�
  -- <YZID>1162695</YZID>   --ҽ��id
  -- <FSH>202704</FSH>   --���ͺ�
  -- <YQSJ>2017-12-05 10:00:00</YQSJ>   --Ҫ��ʱ��
  -- <CZY></CZY>   --����Ա
  -- <CZSJ>2017-12-05 16:26:54</CZSJ>   --����ʱ��

  --���½ڵ�ȡ��ִ��ʱ����
  -- <ZXSM>PDAִ��</ZXSM>   --ִ��ժҪ
  -- <ZXCS></ZXCS>   --ִ�д���
  -- <SYTD>���ֱ�</SYTD>
  -- <JLLY>1</JLLY>    --��¼��Դ��0-PC�˵Ǽǣ�1-�ƶ��ٴ��Ǽǣ��ƶ��˹̶���1
  -- <ZXFS>2</ZXFS>  --ִ�з�ʽ��0-����(������)��1-�ֹ�ִ�У�2-ɨ��ִ��

  --</IN>

  --���Σ�Xml_Out
  --<OUTPUT>
  --   <RESULT>true</RESULT>
  --</OUTPUT>

  --ʧ�ܣ�
  --<OUTPUT>
  --   <RESULT>false</RESULT>
  --   <ERROR>
  --     <MSG>��ϸ������ʾ</MSG>
  --   </ERROR>
  --</OUTPUT>

  n_Type     Number;
  n_ҽ��id   ����ҽ����¼.Id%Type;
  n_���ͺ�   ����ҽ������.���ͺ�%Type;
  d_Ҫ��ʱ�� ����ҽ��ִ��.Ҫ��ʱ��%Type;
  d_ִ��ʱ�� ����ҽ��ִ��.ִ��ʱ��%Type;
  v_ִ��ժҪ ����ҽ��ִ��.ִ��ժҪ%Type;
  n_�������� ����ҽ��ִ��.��������%Type;
  v_��Һͨ�� ����ҽ��ִ��.��Һͨ��%Type;
  n_��¼��Դ ����ҽ��ִ��.��¼��Դ%Type;
  v_ִ����   ����ҽ��ִ��.ִ����%Type;
  n_ִ�з�ʽ ����ҽ��ִ��.ִ�з�ʽ%Type;
Begin
  Select Extractvalue(Value(A), 'IN/TYPE'), Extractvalue(Value(A), 'IN/YZID') As ҽ��id,
         Extractvalue(Value(A), 'IN/FSH') As ���ͺ�,
         To_Date(Extractvalue(Value(A), 'IN/YQSJ'), 'yyyy-mm-dd hh24:mi:ss') As Ҫ��ʱ��,
         Extractvalue(Value(A), 'IN/CZY') As ִ����,
         To_Date(Extractvalue(Value(A), 'IN/CZSJ'), 'yyyy-mm-dd hh24:mi:ss') As ִ��ʱ��,
         Extractvalue(Value(A), 'IN/ZXSM') As ִ��ժҪ, Extractvalue(Value(A), 'IN/ZXCS') As ��������,
         Extractvalue(Value(A), 'IN/SYTD') As ��Һͨ��, Extractvalue(Value(A), 'IN/JLLY') As ��¼��Դ,
         Extractvalue(Value(A), 'IN/ZXFS') As ִ�з�ʽ
  Into n_Type, n_ҽ��id, n_���ͺ�, d_Ҫ��ʱ��, v_ִ����, d_ִ��ʱ��, v_ִ��ժҪ, n_��������, v_��Һͨ��, n_��¼��Դ, n_ִ�з�ʽ
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If n_Type = 1 Then
    Zl_����ҽ��ִ��_Insert(n_ҽ��id, n_���ͺ�, d_Ҫ��ʱ��, n_��������, v_ִ��ժҪ, v_ִ����, d_ִ��ʱ��, 0, 0, 1, Null, Null, Null, 0, 0, 0, v_��Һͨ��,
                     n_��¼��Դ, n_ִ�з�ʽ);
  Else
    Zl_����ҽ��ִ��_Delete(n_ҽ��id, n_���ͺ�, d_ִ��ʱ��, 0, 0, 0);
  End If;
  Xml_Out := Xmltype('<OUTPUT><RESULT>true</RESULT></OUTPUT>');
Exception
  When Others Then
    Xml_Out := Xmltype('<OUTPUT><RESULT>false</RESULT><ERROR><MSG>' || SQLCode || '***' || SQLErrM ||
                       '</MSG></ERROR></OUTPUT>');
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Adviceoperation;
/





------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0016' Where ���=&n_System;
Commit;