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


-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--128798:��͢��,2018-08-03,����ִ��ʱ����ʾСʱ����
Create Or Replace Procedure Zl_����ҽ��ִ��_Cancel
(
  ҽ��id_In     ����ҽ��ִ��.ҽ��id%Type,
  ���ͺ�_In     ����ҽ��ִ��.���ͺ�%Type,
  ����ִ��_In   Number,
  ����Ա���_In ��Ա��.���%Type,
  ����Ա����_In ��Ա��.����%Type,
  ��id_In       ����ҽ��ִ��.ҽ��id%Type,
  �������_In   ����ҽ����¼.�������%Type,
  ִ�в���id_In ������ü�¼.ִ�в���id%Type
  --������ҽ��ID_IN=����ִ�е�ҽ��ID���������Ϊ��ʾ�ļ�����Ŀ��ID�� 
  --      ����ִ��_In=����ҽ������Ƿ���ö�ÿ����Ŀ��ɢ����ִ�еķ�ʽ 
) Is
  --ҽ����صķ��õ��� 
  Cursor c_No Is
    Select a.No || ':' || a.��¼����
    From ����ҽ������ A, ����ҽ����¼ B
    Where a.���ͺ� + 0 = ���ͺ�_In And a.ҽ��id = b.Id And b.������� = �������_In And (b.Id = ��id_In Or b.���id = ��id_In)
    Union
    Select NO || ':' || ��¼���� From ����ҽ������ Where ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In;

  Cursor c_Noone Is
    Select NO || ':' || ��¼����
    From ����ҽ������
    Where ���ͺ� + 0 = ���ͺ�_In And ҽ��id = ҽ��id_In
    Union
    Select NO || ':' || ��¼���� From ����ҽ������ Where ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In;

  r_No       t_Strlist;
  r_No_Stuff t_Strlist;
  r_Finish   t_Numlist;

  n_ִ�д��� Number;
  n_ʣ����� Number;
  d_ִ��ʱ��   Date;
  n_ִ��״̬ Number;

  --Ҫȡ��ִ�еķ�����(������ҩƷ�͸������õ�����) 
  Cursor c_Finish(r_No t_Strlist) Is
    Select /*+ RULE */
     a.Id
    From (Select Distinct a.Id, a.�շ����, a.�շ�ϸĿid
           From ������ü�¼ A, ����ҽ����¼ B,
                (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                         To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As ��¼����
                  From Table(r_No)) N
           Where a.ҽ����� = b.Id And (b.Id = ��id_In Or b.���id = ��id_In) And a.No = n.No And
                 (a.��¼���� = n.��¼���� Or a.��¼���� = 11 And n.��¼���� = 1) And a.��¼״̬ In (0, 1, 3) And
                 (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In)) A, �������� B
    Where a.�շ�ϸĿid = b.����id(+) And Not a.�շ���� In ('5', '6', '7') And Not (a.�շ���� = '4' And Nvl(b.��������, 0) = 1);

  Cursor c_Finishone(r_No t_Strlist) Is
    Select /*+ RULE */
     a.Id
    From (Select a.Id, a.�շ����, a.�շ�ϸĿid
           From ������ü�¼ A,
                (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                         To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As ��¼����
                  From Table(r_No)) N
           Where a.ҽ����� = ҽ��id_In And a.No = n.No And (a.��¼���� = n.��¼���� Or a.��¼���� = 11 And n.��¼���� = 1) And
                 a.��¼״̬ In (0, 1, 3) And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In)) A, �������� B
    Where a.�շ�ϸĿid = b.����id(+) And Not a.�շ���� In ('5', '6', '7') And Not (a.�շ���� = '4' And Nvl(b.��������, 0) = 1);

  --ȡ��ִ���а����������õķ�������ʱ�����ݲ��������Ƿ��Զ����� 
  --��������ҽ��Ŀǰ�����ڵ��������ִ�е���� 
  Cursor c_Stuff(r_No t_Strlist) Is
    Select /*+ RULE */
     b.Id
    From ������ü�¼ A, ҩƷ�շ���¼ B, ����ҽ����¼ C,
         (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                  To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As ��¼����
           From Table(r_No)) N
    Where a.Id = b.����id And (b.��¼״̬ = 1 Or Mod(b.��¼״̬, 3) = 0) And b.����� Is Not Null And a.�շ���� = '4' And a.��¼״̬ = 1 And
          a.ҽ����� = c.Id And c.������� = �������_In And (c.Id = ��id_In Or c.���id = ��id_In) And a.No = n.No And
          (a.��¼���� = n.��¼���� Or a.��¼���� = 11 And n.��¼���� = 1) And b.���� IN(24,25,26) And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In)
    Order By b.ҩƷid;
Begin
  Open c_Noone;
  Fetch c_Noone Bulk Collect
    Into r_No_Stuff;
  Close c_Noone;

  If Nvl(����ִ��_In, 0) = 0 Then
    Open c_No;
    Fetch c_No Bulk Collect
      Into r_No;
    Close c_No;
  Else
    r_No := r_No_Stuff;
  End If;

  --�����ÿ�����Ҫ����ҽ����� 
  --������ҩƷ�͸������õ����ģ���Ϊ��Щ��Ҫ���Ųű�ʾִ�� 
  If Nvl(����ִ��_In, 0) = 0 Then
    Open c_Finish(r_No);
    Fetch c_Finish Bulk Collect
      Into r_Finish;
    Close c_Finish;
  Else
    Open c_Finishone(r_No);
    Fetch c_Finishone Bulk Collect
      Into r_Finish;
    Close c_Finishone;
  End If;

  Select Decode(a.ִ��״̬, 1, a.��������, c.�ǼǴ���), Decode(a.ִ��״̬, 1, 0, a.�������� - c.�ǼǴ���)
  Into n_ִ�д���, n_ʣ�����
  From ����ҽ������ A,
       (Select ҽ��id_In ҽ��id, ���ͺ�_In ���ͺ�, Nvl(Sum(b.��������), 0) As �ǼǴ���
         From ����ҽ��ִ�� B
         Where b.ҽ��id = ҽ��id_In And b.���ͺ� = ���ͺ�_In And Nvl(b.ִ�н��, 1) <> 0) C
  Where a.ҽ��id = c.ҽ��id And a.���ͺ� = c.���ͺ� And a.ҽ��id = ҽ��id_In And a.���ͺ� = ���ͺ�_In;

  --���ȫ��ִ����״̬Ϊ1��δִ��״̬Ϊ0������ִ��״̬Ϊ2 
  Select Decode(n_ʣ�����, 0, 1, Decode(n_ִ�д���, 0, 0, 2)) Into n_ִ��״̬ From Dual;

  --�������ﵥ�ݣ������������շѣ�����ִ�У�2������ȫִ�У�1��,ִ��ʱ��Ϊִ����ɵ�ִ��ʱ�䣬ִ����Ϊִ����ɵ�ִ���� 
  Forall I In 1 .. r_Finish.Count
    Update ������ü�¼
    Set ִ��״̬ = n_ִ��״̬, ִ��ʱ�� = Decode(n_ִ��״̬, 0, d_ִ��ʱ��, ִ��ʱ��), ִ���� = Decode(n_ִ��״̬, 0, Null, ִ����)
    Where ID = r_Finish(I);

  --����������������Զ����� 
  For r_Stuff In c_Stuff(r_No_Stuff) Loop
    Zl_�����շ���¼_��������(r_Stuff.Id, ����Ա����_In, Sysdate, Null, Null, Null, Null, 0, ����Ա����_In);
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ��ִ��_Cancel;
/

--128798:��͢��,2018-08-03,����ִ��ʱ����ʾСʱ����
Create Or Replace Procedure Zl_סԺҽ��ִ��_Cancel
(
  ҽ��id_In     ����ҽ��ִ��.ҽ��id%Type,
  ���ͺ�_In     ����ҽ��ִ��.���ͺ�%Type,
  ����ִ��_In   Number,
  ����Ա���_In ��Ա��.���%Type,
  ����Ա����_In ��Ա��.����%Type,
  ��id_In       ����ҽ��ִ��.ҽ��id%Type,
  �������_In   ����ҽ����¼.�������%Type,
  ִ�в���id_In ������ü�¼.ִ�в���id%Type
  --������ҽ��ID_IN=����ִ�е�ҽ��ID���������Ϊ��ʾ�ļ�����Ŀ��ID��
  --      ����ִ��_In=����ҽ������Ƿ���ö�ÿ����Ŀ��ɢ����ִ�еķ�ʽ
) Is
  --ҽ����صķ��õ���
  Cursor c_No Is
    Select a.No || ':' || a.��¼����
    From ����ҽ������ A, ����ҽ����¼ B
    Where a.���ͺ� + 0 = ���ͺ�_In And a.ҽ��id = b.Id And b.������� = �������_In And (b.Id = ��id_In Or b.���id = ��id_In)
    Union
    Select NO || ':' || ��¼���� From ����ҽ������ Where ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In;

  Cursor c_Noone Is
    Select NO || ':' || ��¼����
    From ����ҽ������
    Where ���ͺ� + 0 = ���ͺ�_In And ҽ��id = ҽ��id_In
    Union
    Select NO || ':' || ��¼���� From ����ҽ������ Where ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In;
  r_No       t_Strlist;
  r_No_Stuff t_Strlist;
  r_Finish   t_Numlist;

  n_ִ�д��� Number;
  n_ʣ����� Number;
  d_ִ��ʱ��   Date;
  n_ִ��״̬ Number;
  n_Count    Number;
  --Ҫȡ��ִ�еķ�����(������ҩƷ�͸������õ�����)
  Cursor c_Finish(r_No t_Strlist) Is
    Select /*+ RULE */
     a.Id
    From (Select Distinct a.Id, a.�շ����, a.�շ�ϸĿid
           From סԺ���ü�¼ A, ����ҽ����¼ B,
                (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                         To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As ��¼����
                  From Table(r_No)) N
           Where a.ҽ����� = b.Id And (b.Id = ��id_In Or b.���id = ��id_In) And a.No = n.No And a.��¼���� = n.��¼���� And
                 a.��¼״̬ In (0, 1, 3) And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In)) A, �������� B
    Where a.�շ�ϸĿid = b.����id(+) And Not a.�շ���� In ('5', '6', '7') And Not (a.�շ���� = '4' And Nvl(b.��������, 0) = 1);

  Cursor c_Finishone(r_No t_Strlist) Is
    Select /*+ RULE */
     a.Id
    From (Select a.Id, a.�շ����, a.�շ�ϸĿid
           From סԺ���ü�¼ A,
                (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                         To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As ��¼����
                  From Table(r_No)) N
           Where a.ҽ����� = ҽ��id_In And a.No = n.No And a.��¼���� = n.��¼���� And a.��¼״̬ In (0, 1, 3) And
                 (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In)) A, �������� B
    Where a.�շ�ϸĿid = b.����id(+) And Not a.�շ���� In ('5', '6', '7') And Not (a.�շ���� = '4' And Nvl(b.��������, 0) = 1);

  --ȡ��ִ���а����������õķ�������ʱ�����ݲ��������Ƿ��Զ�����
  --��������ҽ��Ŀǰ�����ڵ��������ִ�е����
  Cursor c_Stuff(r_No t_Strlist) Is
    Select /*+ RULE */
     b.Id
    From סԺ���ü�¼ A, ҩƷ�շ���¼ B, ����ҽ����¼ C,
         (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                  To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As ��¼����
           From Table(r_No)) N
    Where a.Id = b.����id And (b.��¼״̬ = 1 Or Mod(b.��¼״̬, 3) = 0) And b.����� Is Not Null And a.�շ���� = '4' And a.��¼״̬ = 1 And
          a.ҽ����� = c.Id And c.������� = �������_In And (c.Id = ��id_In Or c.���id = ��id_In) And a.No = n.No And a.��¼���� = n.��¼���� And
          b.���� In (25, 26) And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In)
    Order By b.ҩƷid;

  --ȡ��ִ���а���ҩƷʱ������ִ�е��Զ���ҩ
  Cursor c_Drug(r_No t_Strlist) Is
    Select /*+ RULE */
     b.Id
    From סԺ���ü�¼ A, ҩƷ�շ���¼ B, ����ҽ����¼ C, ������ҳ D,
         (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                  To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As ��¼����
           From Table(r_No)) N
    Where a.Id = b.����id And (b.��¼״̬ = 1 Or Mod(b.��¼״̬, 3) = 0) And b.����� Is Not Null And b.�ⷿid = ִ�в���id_In And
          a.�շ���� In ('5', '6', '7') And a.��¼״̬ = 1 And a.ҽ����� = c.Id And c.������� = �������_In And c.����id = d.����id And
          c.��ҳid = d.��ҳid And (c.Id = ��id_In Or c.���id = ��id_In) And a.No = n.No And a.��¼���� = n.��¼����
    Order By b.ҩƷid;

  v_ҽ����Ч ����ҽ����¼.ҽ����Ч%Type;
Begin
  Open c_Noone;
  Fetch c_Noone Bulk Collect
    Into r_No_Stuff;
  Close c_Noone;

  If Nvl(����ִ��_In, 0) = 0 Then
    Open c_No;
    Fetch c_No Bulk Collect
      Into r_No;
    Close c_No;
  Else
    r_No := r_No_Stuff;
  End If;

  --�����ÿ�����Ҫ����ҽ�����
  --������ҩƷ�͸������õ����ģ���Ϊ��Щ��Ҫ���Ųű�ʾִ��
  If Nvl(����ִ��_In, 0) = 0 Then
    Open c_Finish(r_No);
    Fetch c_Finish Bulk Collect
      Into r_Finish;
    Close c_Finish;
  Else
    Open c_Finishone(r_No);
    Fetch c_Finishone Bulk Collect
      Into r_Finish;
    Close c_Finishone;
  End If;

  Select Decode(a.ִ��״̬, 1, a.��������, c.�ǼǴ���), Decode(a.ִ��״̬, 1, 0, a.�������� - c.�ǼǴ���)
  Into n_ִ�д���, n_ʣ�����
  From ����ҽ������ A,
       (Select ҽ��id_In ҽ��id, ���ͺ�_In ���ͺ�, Nvl(Sum(b.��������), 0) As �ǼǴ���
         From ����ҽ��ִ�� B
         Where b.ҽ��id = ҽ��id_In And b.���ͺ� = ���ͺ�_In And Nvl(b.ִ�н��, 1) <> 0) C
  Where a.ҽ��id = c.ҽ��id And a.���ͺ� = c.���ͺ� And a.ҽ��id = ҽ��id_In And a.���ͺ� = ���ͺ�_In;

  --���ȫ��ִ����״̬Ϊ1��δִ��״̬Ϊ0������ִ��״̬Ϊ2
  Select Decode(n_ʣ�����, 0, 1, Decode(n_ִ�д���, 0, 0, 2)) Into n_ִ��״̬ From Dual;

  Forall I In 1 .. r_Finish.Count
    Update סԺ���ü�¼
    Set ִ��״̬ = n_ִ��״̬, ִ��ʱ�� = Decode(n_ִ��״̬, 0, d_ִ��ʱ��, ִ��ʱ��), ִ���� = Decode(n_ִ��״̬, 0, Null, ִ����)
    Where ID = r_Finish(I);

  --����������������Զ�����
  For r_Stuff In c_Stuff(r_No_Stuff) Loop
    Zl_�����շ���¼_��������(r_Stuff.Id, ����Ա����_In, Sysdate, Null, Null, Null, Null, 0, ����Ա����_In);
  End Loop;

  --����ҩƷ�Զ���ҩ(ֻ�ڻ�ʿվ������ҩƷ�Ŵ���,�����ɲ������α��ж�)
  Select Max(a.ҽ����Ч), Max(Decode(b.����id, ִ�в���id_In, 1, 0))
  Into v_ҽ����Ч, n_Count
  From ����ҽ����¼ A, ���˱䶯��¼ B
  Where a.����id = b.����id And a.��ҳid = b.��ҳid And a.Id = ҽ��id_In;

  If Substr(Zl_Getsysparameter('����ִ���Զ����', 1254), v_ҽ����Ч + 1, 1) = '1' And n_Count = 1 Then
    For r_Drug In c_Drug(r_No_Stuff) Loop
      Zl_ҩƷ�շ���¼_������ҩ(r_Drug.Id, ����Ա����_In, Sysdate);
    End Loop;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_סԺҽ��ִ��_Cancel;
/

--128798:��͢��,2018-08-03,����ִ��ʱ����ʾСʱ����
Create Or Replace Procedure Zl_����ҽ��ִ��_Update
(
  ԭִ��ʱ��_In ����ҽ��ִ��.ִ��ʱ��%Type,
  ҽ��id_In     ����ҽ��ִ��.ҽ��id%Type,
  ���ͺ�_In     ����ҽ��ִ��.���ͺ�%Type,
  Ҫ��ʱ��_In   ����ҽ��ִ��.Ҫ��ʱ��%Type,
  ��������_In   ����ҽ��ִ��.��������%Type,
  ִ��ժҪ_In   ����ҽ��ִ��.ִ��ժҪ%Type,
  ִ����_In     ����ҽ��ִ��.ִ����%Type,
  ִ��ʱ��_In   ����ҽ��ִ��.ִ��ʱ��%Type,
  ִ�н��_In   ����ҽ��ִ��.ִ�н��%Type := 1,
  δִ��ԭ��_In ����ҽ��ִ��.˵��%Type := Null,
  ����ִ��_In   Number := 0,
  ����Ա���_In ��Ա��.���%Type := Null,
  ����Ա����_In ��Ա��.����%Type := Null,
  ִ�в���id_In ������ü�¼.ִ�в���id%Type := 0
  --������ҽ��ID_IN=����ִ�е�ҽ��ID���������Ϊ��ʾ�ļ�����Ŀ��ID��
  --ִ�в���id_In=������ָ��ִ�в��ŵķ��ã���������0ʱ������ִ�в���
) Is
  --����Ҫִ�е�����¼,�������˸�������,��鲿λ�ļ�¼
  --��������,��ҩ�巨,�ɼ�������������,�������ֻ��д�ڵ�һ����Ŀ�ϣ���ִ��״̬��ͬ
  v_Temp     Varchar2(255); 
  v_��Ա���� ��Ա��.����%Type;

  v_��id        ����ҽ����¼.Id%Type;
  v_�������    ����ҽ����¼.�������%Type;
  v_ִ�н��old ����ҽ��ִ��.ִ�н��%Type;
  n_��������old ����ҽ��ִ��.��������%Type;

  v_������Դ ����ҽ����¼.������Դ%Type;
  v_�������� ����ҽ������.��¼����%Type;

  n_ִ�д��� Number;
  n_ʣ����� Number;
  n_ִ��״̬ Number;
  n_�������� Number;
  n_�������� Number;
  v_Count    Number;
  n_�Ǽ����� Number;
  d_Ҫ��ʱ�� Date;
  d_ִ��ʱ�� Date;

  d_�Ǽ�ʱ��   ����ҽ��ִ��.�Ǽ�ʱ��%Type;
  n_ȡ��ִ��   Number;
  n_Diffday    Number(18, 3);
  n_ִ�п���id Number;

  v_Date  Date;
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --��ǰ������Ա
  If ����Ա���_In Is Not Null And ����Ա����_In Is Not Null Then
    v_��Ա���� := ����Ա����_In;
  Else
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  End If;

  Select Sysdate Into v_Date From Dual;
  Select Nvl(ִ�н��, 1), Nvl(��������, 0), �Ǽ�ʱ��
  Into v_ִ�н��old, n_��������old, d_�Ǽ�ʱ��
  From ����ҽ��ִ��
  Where ҽ��id = ҽ��id_In And ���ͺ� = ���ͺ�_In And ִ��ʱ�� = ԭִ��ʱ��_In;
  -----ȡ��ִ����Ч��������
  Select Zl_To_Number(Nvl(zl_GetSysParameter(220), '999')) Into n_ȡ��ִ�� From Dual;
  Select v_Date - d_�Ǽ�ʱ�� Into n_Diffday From Dual;
  --�Ǽ�ʱ�䳬��ȡ��ִ�������ļ�¼���������޸�ҽ��ִ�����
  If n_Diffday > n_ȡ��ִ�� Then
    v_Error := 'ҽ��ִ�еǼ�ʱ�䳬����ȡ��ִ����Ч�����������޸�ҽ��ִ�������';
    Raise Err_Custom;
  End If;
  Select ִ�в���id Into n_ִ�п���id From ����ҽ������ Where ҽ��id = ҽ��id_In And ���ͺ� = ���ͺ�_In;
  --����ҽ��ִ��
  Update ����ҽ��ִ��
  Set Ҫ��ʱ�� = Ҫ��ʱ��_In, �������� = ��������_In, ִ��ժҪ = ִ��ժҪ_In, ִ���� = ִ����_In, ִ��ʱ�� = ִ��ʱ��_In, �Ǽ�ʱ�� = v_Date, �Ǽ��� = v_��Ա����,
      ִ�н�� = ִ�н��_In, ˵�� = δִ��ԭ��_In, ִ�п���id = n_ִ�п���id
  Where ҽ��id = ҽ��id_In And ���ͺ� = ���ͺ�_In And ִ��ʱ�� = ԭִ��ʱ��_In;
  --����ִ�д�������ִ�н���޸ĺ���Ҫ���µ��ݵ�ִ��״̬
  If v_ִ�н��old <> ִ�н��_In Or n_��������old <> ��������_In Then
    Select ������Դ, Nvl(���id, ID), �������
    Into v_������Դ, v_��id, v_�������
    From ����ҽ����¼
    Where ID = ҽ��id_In;
  
    If v_������Դ = 2 Then
      Select Decode(��¼����, 1, 1, Decode(�������, 1, 1, 2))
      Into v_��������
      From ����ҽ������
      Where ���ͺ� = ���ͺ�_In And ҽ��id = ҽ��id_In;
    Else
      v_�������� := 1;
    End If;
  
    Select Decode(a.ִ��״̬, 1, a.��������, c.�ǼǴ���), Decode(a.ִ��״̬, 1, 0, a.�������� - c.�ǼǴ���), a.��������, c.�ǼǴ���


    Into n_ִ�д���, n_ʣ�����, n_��������, n_�Ǽ�����
    From ����ҽ������ A,
         (Select ҽ��id_In ҽ��id, ���ͺ�_In ���ͺ�, Nvl(Sum(b.��������), 0) As �ǼǴ���
           From ����ҽ��ִ�� B
           Where b.ҽ��id = ҽ��id_In And b.���ͺ� = ���ͺ�_In And Nvl(b.ִ�н��, 1) = 1) C
    Where a.ҽ��id = c.ҽ��id And a.���ͺ� = c.���ͺ� And a.ҽ��id = ҽ��id_In And a.���ͺ� = ���ͺ�_In;
  
    --���ȫ��ִ����״̬Ϊ1��δִ��״̬Ϊ0������ִ��״̬Ϊ2
    Select Decode(n_ʣ�����, 0, 1, Decode(n_ִ�д���, 0, 0, 2)) Into n_ִ��״̬ From Dual;
  
    --����ҽ��ִ�мƼ�.ִ��״̬
    If n_�������� > 0 Then
      Select Count(Distinct Ҫ��ʱ��) Into v_Count From ҽ��ִ�мƼ� Where ҽ��id = ҽ��id_In And ���ͺ� = ���ͺ�_In;
      If v_Count > 0 Then
        n_�������� := n_�������� / v_Count;
        --��ִ������+�������� �ܹ��ܹ�ִ�ж��ٸ�ʱ���,ȡ�������
        v_Count := Ceil((n_�Ǽ�����) / n_��������);
        If n_�Ǽ����� = 0 Then
          Update ҽ��ִ�мƼ�
          Set ִ��״̬ = 0
          Where ҽ��id = ҽ��id_In And ���ͺ� = ���ͺ�_In And Nvl(ִ��״̬, 0) <> 2;
        Else
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
            Update ҽ��ִ�мƼ�
            Set ִ��״̬ = 0
            Where ҽ��id = ҽ��id_In And ���ͺ� = ���ͺ�_In And Ҫ��ʱ�� > d_Ҫ��ʱ�� And Nvl(ִ��״̬, 0) <> 2;
          End If;
        End If;
      End If;
    End If;
  
    --ִ�д�����Ϊ0�ͱ��Ϊ����ִ��
    If Nvl(����ִ��_In, 0) = 1 Then
      Update ����ҽ������
      Set ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3), ����� = Null, ���ʱ�� = Null
      Where ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In;
    Else
      Update ����ҽ������
      Set ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3), ����� = Null, ���ʱ�� = Null
      Where ִ��״̬ In (0, 3) And ���ͺ� + 0 = ���ͺ�_In And
            ҽ��id In (Select ID From ����ҽ����¼ Where (ID = v_��id Or ���id = v_��id) And ������� = v_�������);
    End If;
  
    If v_�������� = 2 Then
      If Nvl(����ִ��_In, 0) = 1 Then
        Update סԺ���ü�¼ A
        Set ִ��״̬ = n_ִ��״̬, ִ���� = Decode(n_ִ��״̬, 0, Null, ִ����_In), ִ��ʱ�� = Decode(n_ִ��״̬, 0, d_ִ��ʱ��, ִ��ʱ��_In)
        Where �շ���� Not In ('5', '6', '7') And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In) And Not Exists
         (Select 1 From �������� Where ����id = a.�շ�ϸĿid And �������� = 1) And a.��¼״̬ In (0, 1, 3) And
              (ҽ�����, NO, ��¼����) In
              (Select ҽ��id, NO, ��¼����
               From ����ҽ������
               Where ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3) And ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In);
      Else
        Update סԺ���ü�¼ A
        Set ִ��״̬ = n_ִ��״̬, ִ���� = Decode(n_ִ��״̬, 0, Null, ִ����_In), ִ��ʱ�� = Decode(n_ִ��״̬, 0, d_ִ��ʱ��, ִ��ʱ��_In)
        Where �շ���� Not In ('5', '6', '7') And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In) And Not Exists
         (Select 1 From �������� Where ����id = a.�շ�ϸĿid And �������� = 1) And a.��¼״̬ In (0, 1, 3) And
              (ҽ�����, NO, ��¼����) In
              (Select ҽ��id, NO, ��¼����
               From ����ҽ������
               Where ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3) And ���ͺ� + 0 = ���ͺ�_In And
                     ҽ��id In
                     (Select ID From ����ҽ����¼ Where (ID = v_��id Or ���id = v_��id) And ������� = v_�������));
      End If;
    Else
      If Nvl(����ִ��_In, 0) = 1 Then
        Update ������ü�¼ A
        Set ִ��״̬ = n_ִ��״̬, ִ���� = Decode(n_ִ��״̬, 0, Null, ִ����_In), ִ��ʱ�� = Decode(n_ִ��״̬, 0, d_ִ��ʱ��, ִ��ʱ��_In)
        Where �շ���� Not In ('5', '6', '7') And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In) And Not Exists
         (Select 1 From �������� Where ����id = a.�շ�ϸĿid And �������� = 1) And a.��¼״̬ In (0, 1, 3) And
              (ҽ�����, NO, ��¼����) In
              (Select ҽ��id, NO, ��¼����
               From ����ҽ������
               Where ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3) And ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In);
      Else
        Update ������ü�¼ A
        Set ִ��״̬ = n_ִ��״̬, ִ���� = Decode(n_ִ��״̬, 0, Null, ִ����_In), ִ��ʱ�� = Decode(n_ִ��״̬, 0, d_ִ��ʱ��, ִ��ʱ��_In)
        Where �շ���� Not In ('5', '6', '7') And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In) And Not Exists
         (Select 1 From �������� Where ����id = a.�շ�ϸĿid And �������� = 1) And a.��¼״̬ In (0, 1, 3) And
              (ҽ�����, NO, ��¼����) In
              (Select ҽ��id, NO, ��¼����
               From ����ҽ������
               Where ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3) And ���ͺ� + 0 = ���ͺ�_In And
                     ҽ��id In
                     (Select ID From ����ҽ����¼ Where (ID = v_��id Or ���id = v_��id) And ������� = v_�������));
      End If;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ��ִ��_Update;
/

--128798:��͢��,2018-08-03,����ִ��ʱ����ʾСʱ����
CREATE OR REPLACE Procedure Zl_����ҽ��ִ��_Delete
(
  ҽ��id_In     In ����ҽ��ִ��.ҽ��id%Type,
  ���ͺ�_In     In ����ҽ��ִ��.���ͺ�%Type,
  ִ��ʱ��_In   In ����ҽ��ִ��.ִ��ʱ��%Type,
  ����ִ��_In   In Number := 0,
  �Զ�ȡ��_In   In Number := 0,
  ִ�в���id_In In ������ü�¼.ִ�в���id%Type := 0
  --������ҽ��ID_IN=����ִ�е�ҽ��ID���������Ϊ��ʾ�ļ�����Ŀ��ID��
  --ִ�в���id_In=������ָ��ִ�в��ŵķ��ã���������0ʱ������ִ�в���
) Is
  --����Ҫִ�е�����¼,�������˸�������,��鲿λ�ļ�¼
  --��������,��ҩ�巨,�ɼ�������������,�������ֻ��д�ڵ�һ����Ŀ�ϣ���ִ��״̬��ͬ
  v_��id     ����ҽ����¼.Id%Type;
  v_������� ����ҽ����¼.�������%Type;
  v_������Դ ����ҽ����¼.������Դ%Type;
  v_�������� ����ҽ������.��¼����%Type;
  v_�������� ������ĿĿ¼.��������%Type;

  n_����id   ����ҽ����¼.����id%Type;
  n_��ҳid   ����ҽ����¼.��ҳid%Type;
  v_�Һŵ�   ����ҽ����¼.�Һŵ�%Type;
  n_�������� ����ҽ��ִ��.��������%Type;
  v_ִ��ժҪ ����ҽ��ִ��.ִ��ժҪ%Type;
  n_ִ�п��� ����ҽ��ִ��.ִ�п���id%Type;
  v_ִ����   ����ҽ��ִ��.ִ����%Type;
  v_�˶���   ����ҽ��ִ��.�˶���%Type;
  n_��¼��Դ ����ҽ��ִ��.��¼��Դ%Type;

  v_�Զ�ȡ�� Number;
  v_ִ��״̬ Number;

  n_ִ�д��� Number;
  n_ʣ����� Number;
  n_ִ��״̬ Number;
  n_ִ�н�� Number;

  n_�������� Number;
  n_�������� Number;
  v_Count    Number;
  n_�Ǽ����� Number;
  d_Ҫ��ʱ�� Date;
  d_ִ��ʱ�� Date;

  d_�Ǽ�ʱ�� ����ҽ��ִ��.�Ǽ�ʱ��%Type;
  n_ȡ��ִ�� Number;
  n_Diffday  Number(18, 3);
  Err_Custom Exception;
  v_Error Varchar2(2000);
Begin
  Select a.������Դ, Nvl(a.���id, a.Id), Nvl(a.�������, '*'), Nvl(b.��������, '0') ��������, a.����id, a.��ҳid, a.�Һŵ�
  Into v_������Դ, v_��id, v_�������, v_��������, n_����id, n_��ҳid, v_�Һŵ�
  From ����ҽ����¼ A, ������ĿĿ¼ B
  Where a.Id = ҽ��id_In And a.������Ŀid = b.Id(+);

  Select Nvl(a.ִ�н��, 1), a.�Ǽ�ʱ��, a.Ҫ��ʱ��, a.��������, a.ִ��ժҪ, a.ִ�п���id, a.ִ����, a.�˶���, a.��¼��Դ
  Into n_ִ�н��, d_�Ǽ�ʱ��, d_Ҫ��ʱ��, n_��������,
       
       v_ִ��ժҪ, n_ִ�п���, v_ִ����, v_�˶���, n_��¼��Դ
  
  From ����ҽ��ִ�� A
  Where a.ҽ��id = ҽ��id_In And a.���ͺ� + 0 = ���ͺ�_In And a.ִ��ʱ�� = ִ��ʱ��_In;

  -----ȡ��ִ����Ч��������
  Select Zl_To_Number(Nvl(zl_GetSysParameter(220), '999')) Into n_ȡ��ִ�� From Dual;
  Select Sysdate - d_�Ǽ�ʱ�� Into n_Diffday From Dual;
  --�Ǽ�ʱ�䳬��ȡ��ִ�������ļ�¼��������ɾ��ҽ��ִ�м�¼
  If n_Diffday > n_ȡ��ִ�� Then
    v_Error := 'ҽ��ִ�еǼ�ʱ�䳬����ȡ��ִ����Ч����������ɾ��ҽ��ִ�м�¼��';
    Raise Err_Custom;
  End If;

  If v_������Դ = 2 Then
    Select Decode(��¼����, 1, 1, Decode(�������, 1, 1, 2))
    Into v_��������
    From ����ҽ������
    Where ���ͺ� = ���ͺ�_In And ҽ��id = ҽ��id_In;
  Else
    v_�������� := 1;
  End If;

  --����ҽ��ִ��
  Delete From ����ҽ��ִ�� Where ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In And ִ��ʱ�� = ִ��ʱ��_In;
  b_Message.Zlhis_Cis_051(n_����id, n_��ҳid, v_�Һŵ�, ���ͺ�_In, ҽ��id_In, d_Ҫ��ʱ��, ִ��ʱ��_In, n_��������, n_ִ�н��, v_ִ��ժҪ, n_ִ�п���,
                          v_ִ����, v_�˶���, n_��¼��Դ);
  d_Ҫ��ʱ�� := Null;

  --����δִ�е�ҽ��ִ�м�¼��ɾ����������ҽ�������Լ�������Ϣ��ִ��״̬
  If n_ִ�н�� <> 0 Then
    Select Decode(a.ִ��״̬, 1, a.��������, c.�ǼǴ���), Decode(a.ִ��״̬, 1, 0, a.�������� - c.�ǼǴ���), a.��������, c.�ǼǴ���


    
    Into n_ִ�д���, n_ʣ�����, n_��������, n_�Ǽ�����
    From ����ҽ������ A,
         (Select ҽ��id_In ҽ��id, ���ͺ�_In ���ͺ�, Nvl(Sum(b.��������), 0) As �ǼǴ���
           From ����ҽ��ִ�� B
           Where b.ҽ��id = ҽ��id_In And b.���ͺ� = ���ͺ�_In And Nvl(b.ִ�н��, 1) <> 0) C
    Where a.ҽ��id = c.ҽ��id And a.���ͺ� = c.���ͺ� And a.ҽ��id = ҽ��id_In And a.���ͺ� = ���ͺ�_In;
  
    --���ȫ��ִ����״̬Ϊ1��δִ��״̬Ϊ0������ִ��״̬Ϊ2
    Select Decode(n_ʣ�����, 0, 1, Decode(n_ִ�д���, 0, 0, 2)) Into n_ִ��״̬ From Dual;
  
    --����ҽ��ִ�мƼ�.ִ��״̬
    If n_�������� > 0 Then
      If n_�Ǽ����� > 0 Then
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
            Update ҽ��ִ�мƼ�
            Set ִ��״̬ = 0
            Where ҽ��id = ҽ��id_In And ���ͺ� = ���ͺ�_In And Ҫ��ʱ�� > d_Ҫ��ʱ�� And Nvl(ִ��״̬, 0) <> 2;
          End If;
        End If;
      Else
        Update ҽ��ִ�мƼ� Set ִ��״̬ = 0 Where ҽ��id = ҽ��id_In And ���ͺ� = ���ͺ�_In And Nvl(ִ��״̬, 0) <> 2;
      End If;
    End If;
  
    --���ִ�����ɾ���˾͸���ִ��״̬
    If Nvl(����ִ��_In, 0) = 1 Then
      Update ����ҽ������
      Set ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3), ����� = Null, ���ʱ�� = Null
      Where ִ��״̬ = 3 And ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In;
    Else
      Update ����ҽ������ A
      Set ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3), ����� = Null, ���ʱ�� = Null
      Where ִ��״̬ = 3 And ���ͺ� + 0 = ���ͺ�_In And
            ҽ��id In
            (Select ID From ����ҽ����¼ Where (ID = v_��id Or ���id = v_��id) And Nvl(�������, '*') = v_�������) And Not Exists
       (Select 1 From ����ҽ��ִ�� Where ���ͺ� + 0 = ���ͺ�_In And ҽ��id = a.ҽ��id);
    End If;
    --���¶�Ӧ�ķ���ִ��״̬Ϊδִ��
    --��Ӧ�ô���ҩƷ�͸������õ�����
    If v_�������� = 2 Then
      If Nvl(����ִ��_In, 0) = 1 Then
        Update סԺ���ü�¼ A
        Set ִ��״̬ = n_ִ��״̬, ִ���� = Decode(n_ִ��״̬, 0, Null, ִ����), ִ��ʱ�� = Decode(n_ִ��״̬, 0, d_ִ��ʱ��, ִ��ʱ��)
        Where �շ���� Not In ('5', '6', '7') And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In) And Not Exists
         (Select 1 From �������� Where ����id = a.�շ�ϸĿid And �������� = 1) And a.��¼״̬ In (0, 1, 3) And
              (ҽ�����, NO, ��¼����) In
              (Select ҽ��id, NO, ��¼����
               From ����ҽ������
               Where ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3) And ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In);
      Else
        Update סԺ���ü�¼ A
        Set ִ��״̬ = n_ִ��״̬, ִ���� = Decode(n_ִ��״̬, 0, Null, ִ����), ִ��ʱ�� = Decode(n_ִ��״̬, 0, d_ִ��ʱ��, ִ��ʱ��)
        Where �շ���� Not In ('5', '6', '7') And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In) And Not Exists
         (Select 1 From �������� Where ����id = a.�շ�ϸĿid And �������� = 1) And a.��¼״̬ In (0, 1, 3) And
              (ҽ�����, NO, ��¼����) In
              (Select ҽ��id, NO, ��¼����
               From ����ҽ������
               Where ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3) And ���ͺ� + 0 = ���ͺ�_In And
                     ҽ��id In
                     (Select ID From ����ҽ����¼ Where (ID = v_��id Or ���id = v_��id) And ������� = v_�������));
      End If;
    Else
      If Nvl(����ִ��_In, 0) = 1 Then
        Update ������ü�¼ A
        Set ִ��״̬ = n_ִ��״̬, ִ���� = Decode(n_ִ��״̬, 0, Null, ִ����), ִ��ʱ�� = Decode(n_ִ��״̬, 0, d_ִ��ʱ��, ִ��ʱ��)
        Where �շ���� Not In ('5', '6', '7') And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In) And Not Exists
         (Select 1 From �������� Where ����id = a.�շ�ϸĿid And �������� = 1) And a.��¼״̬ In (0, 1, 3) And
              (ҽ�����, NO, ��¼����) In
              (Select ҽ��id, NO, ��¼����
               From ����ҽ������
               Where ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3) And ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In);
      Else
        Update ������ü�¼ A
        Set ִ��״̬ = n_ִ��״̬, ִ���� = Decode(n_ִ��״̬, 0, Null, ִ����), ִ��ʱ�� = Decode(n_ִ��״̬, 0, d_ִ��ʱ��, ִ��ʱ��)
        Where �շ���� Not In ('5', '6', '7') And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In) And Not Exists
         (Select 1 From �������� Where ����id = a.�շ�ϸĿid And �������� = 1) And a.��¼״̬ In (0, 1, 3) And
              (ҽ�����, NO, ��¼����) In
              (Select ҽ��id, NO, ��¼����
               From ����ҽ������
               Where ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3) And ���ͺ� + 0 = ���ͺ�_In And
                     ҽ��id In
                     (Select ID From ����ҽ����¼ Where (ID = v_��id Or ���id = v_��id) And ������� = v_�������));
      End If;
    End If;
    --����ɼ��Զ�ȡ���ɼ��˲ɼ�ʱ��
    If v_������� = 'E' And v_�������� = '6' Then
      Update ����ҽ������ A
      Set a.������ = Null, a.����ʱ�� = Null
      Where ҽ��id In (Select ID From ����ҽ����¼ Where (ID = v_��id Or ���id = v_��id)) And ���ͺ� = ���ͺ�_In;
    End If;
  
    --�����ִ�еģ�ִ�����μ���֮���Զ�ȡ��ִ�����Ϊ����ִ�л�δִ��(��Ҫ����PDA�Զ�ִ��)
    If Nvl(�Զ�ȡ��_In, 0) = 1 Then
      Begin
        Select Decode(Sign(Nvl(Sum(b.��������), 0) - a.��������), -1, 1, 0), Decode(Sign(Nvl(Sum(b.��������), 0)), 0, 0, 3)
        Into v_�Զ�ȡ��, v_ִ��״̬
        From ����ҽ������ A, ����ҽ��ִ�� B
        Where a.ҽ��id = b.ҽ��id(+) And a.���ͺ� = b.���ͺ�(+) And a.ִ��״̬ = 1 And a.ҽ��id = ҽ��id_In And a.���ͺ� = ���ͺ�_In
        Group By a.��������;
      Exception
        When Others Then
          Null;
      End;
    
      If Nvl(v_�Զ�ȡ��, 0) = 1 Then
        Zl_����ҽ��ִ��_Cancel(ҽ��id_In, ���ͺ�_In, Null, ����ִ��_In, ִ�в���id_In);
      
        If v_ִ��״̬ = 3 Then
          Select Nvl(���id, ID), ������� Into v_��id, v_������� From ����ҽ����¼ Where ID = ҽ��id_In;
          Update ����ҽ������
          Set ִ��״̬ = 3, ����� = Null, ���ʱ�� = Null
          Where ���ͺ� + 0 = ���ͺ�_In And
                ҽ��id In (Select ID From ����ҽ����¼ Where (ID = v_��id Or ���id = v_��id) And ������� = v_�������);
        End If;
      
      End If;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ��ִ��_Delete;
/

--128798:��͢��,2018-08-03,����ִ��ʱ����ʾСʱ����
CREATE OR REPLACE Procedure Zl_����ҽ��ִ��_Insert
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
  d_ִ��ʱ��   Date;

  v_Date  Date;
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --������죬��ֹ��������ִ�м�¼
  Begin
    Select (a.�������� - c.�ǼǴ���) As ʣ������, a.��������, a.ִ�в���id, Nvl(d.������Ŀid, 0), c.�ǼǴ���
    Into v_Count, n_��������, n_ִ�п���id, n_������Ŀid, n_�Ǽ�����
    From ����ҽ������ A,
         (Select ҽ��id_In As ҽ��id, ���ͺ�_In As ���ͺ�, Nvl(Sum(b.��������), 0) As �ǼǴ���
           From ����ҽ��ִ�� B
           Where b.ҽ��id = ҽ��id_In And b.���ͺ� = ���ͺ�_In) C, ����ҽ����¼ D
    Where a.ҽ��id = c.ҽ��id And a.���ͺ� = c.���ͺ� And a.ҽ��id = ҽ��id_In And a.ҽ��id = d.Id And a.���ͺ� = ���ͺ�_In;
  Exception
    When Others Then
      v_Count := ��������_In;
  End;
  v_����ִ�� := zl_GetSysParameter(288);
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
  From ����ҽ����¼ A
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
  Select a.������Դ, ִ�п���id, Nvl(a.���id, a.Id), Nvl(a.�������, '*'), Nvl(b.��������, '0') ��������
  Into v_������Դ, v_����id, v_��id, v_�������, v_��������
  From ����ҽ����¼ A, ������ĿĿ¼ B
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
    Select Nvl(zl_GetSysParameter(184), '') Into v_��Һ���� From Dual;
  
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
  From ����ҽ������ A,
       (Select ҽ��id_In ҽ��id, ���ͺ�_In ���ͺ�, Nvl(Sum(b.��������), 0) As �ǼǴ���
         From ����ҽ��ִ�� B
         Where b.ҽ��id = ҽ��id_In And b.���ͺ� = ���ͺ�_In And Nvl(b.ִ�н��, 1) = 1) C
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
          ҽ��id In (Select ID
                   From ����ҽ����¼
                   Where ID = v_��id And Nvl(�������, '*') = v_�������
                   Union All
                   Select ID
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
                                         ҽ��id In (Select ID
                                                  From ����ҽ����¼
                                                  Where ID = v_��id And ������� = v_�������
                                                  Union All
                                                  Select ID
                                                  From ����ҽ����¼
                                                  Where ���id = v_��id And ������� = v_�������));
      End If;
    End If;
    --�����Զ���ɲɼ�
    If v_������� = 'E' And v_�������� = '6' Then
      Update ����ҽ������ A
      Set a.������ = ִ����_In, a.����ʱ�� = ִ��ʱ��_In
      Where ҽ��id In
            (Select ID From ����ҽ����¼ Where ID = v_��id Union All Select ID From ����ҽ����¼ Where ���id = v_��id) And
            ���ͺ� = ���ͺ�_In;
    End If;
  
    --ִ�����δﵽ֮���Զ����ִ��(��Ҫ����PDA�Զ�ִ��)������������ƶ��ٴ�����ʿվ��PDAһ�¡�
    v_�Զ���� := �Զ����_In;
    If �Զ����_In = 1 Then
      --ҽ���Ѿ������״̬�����ٵ���ִ����ɹ��̴˴�����Ϊ���Զ����
      Select Max(a.ִ��״̬) Into v_Count From ����ҽ������ A Where a.ҽ��id = ҽ��id_In And a.���ͺ� = ���ͺ�_In;
      If v_Count = 1 Then
        v_�Զ���� := 0;
      End If;
      v_Count := Null;
    End If;
  
    If Nvl(v_�Զ����, 0) = 0 And (v_������Դ = 2 Or v_������Դ = 1) And Instr('C,D', v_�������) = 0 Then
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
        From ����ҽ������ A, ����ҽ��ִ�� B
        Where a.ҽ��id = b.ҽ��id And a.���ͺ� = b.���ͺ� And a.ִ��״̬ In (0, 3) And a.ҽ��id = ҽ��id_In And a.���ͺ� = ���ͺ�_In
        Group By a.��������;
      Exception
        When Others Then
          Null;
      End;
    
      If Nvl(v_�Զ����, 0) = 1 Or ������Ŀ����_In = 1 Then
        Zl_����ҽ��ִ��_Finish(ҽ��id_In, ���ͺ�_In, Null, ����ִ��_In, v_��Ա���, v_��Ա����, ִ�в���id_In, ������Ŀ����_In);
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
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ��ִ��_Insert;
/

--129387:����,2018-08-03,���صĲ��˻�����Ϣ������������
Create Or Replace Procedure Zl_Third_Getpatiinfo
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  -------------------------------------------------------------------------------------------------- 
  --����:��ȡ���˻�����Ϣ
  --���:Xml_In: 
  --  <IN> 
  --      <BRID></BRID>     --����ID
  --      <SFZH></SFZH>     --���֤��
  --      <CXKLB></CXKLB>   --��ѯ�����
  --      <MZH></MZH>       --�����
  --      <GHDH></GHDH>     --�Һŵ���
  --      <YLKLB></YLKLB>   --ҽ�ƿ����ID��������
  --      <YLKH></YLKH>     --ҽ�ƿ���
  --      <BRXM></BRXM>     --��������
  --  </IN> 
  --  ����ʶ��˳��:
  --  1.���벡��ID ,�Բ���IDΪ׼
  --  2.����Һŵ������ԹҺŵ�Ϊ׼
  --  3.���뿨���ID�����Կ����ID�Ϳ���Ϊ׼
  --  4.��������ţ����������Ϊ׼
  --  5.�������֤���������֤�ź�����Ϊ׼  
  --����:Xml_Out 
  -- <OUTPUT>
  --   <BR>
  --     <BRID></BRID>       --����ID
  --     <XM></XM>           --����
  --     <XB></XB>           --�Ա�
  --     <Nl></NL>           --����
  --     <CSRQ></CSRQ>       --��������
  --     <MZH></MZH>         --�����
  --     <HY></HY>           --����
  --     <GJ></GJ>           --����
  --     <MZ></MZ>           --����
  --     <XL></XL>           --ѧ��
  --     <SF></SF>           --���
  --     <ZY></ZY>           --ְҵ
  --     <SFZH></SFZH>       --���֤��
  --     <FKFS></FKFS>       --���ʽ
  --     <LXFS></LXFS>       --��ϵ��ʽ
  --     <LXRXM></LXRXM>     --��ϵ������
  --     <LXRDH></LXRDH>     --��ϵ�˵绰
  --     <LXRDZ></LXRDZ>     --��ϵ�˵�ַ
  --     <LXDH></LXDH>       --��ϵ�绰
  --     <XJZDZ></XJZDZ>     --�־�ס��ַ 
  --     <HJDZ></HJDZ>       --������ַ
  --     <CSDD></CSDD>       --�����ص�
  --     <KSID></KSID>       --����ID
  --     <CXKH></CXKH>       --��ѯ����
  --     <GMS></GMS>         --����ʷ         
  --     <GHD></GHD>         --�Һŵ���
  --     <GHSJ></GHSJ>       --�Һ�ʱ��
  --     <JZSJ></JZSJ>       --����ʱ��
  --     <JZKS></JZKS>       --�������
  --     <JZYS></JZYS>       --����ҽ��
  --   </BR>
  -- </OUTPUT>
  -------------------------------------------------------------------------------------------------- 

  v_����id       Varchar2(30000);
  v_ҽ�ƿ�       Varchar2(500);
  v_�����       Varchar2(500);
  v_�Һŵ�       ���˹Һż�¼.No%Type;
  v_����         ����ҽ�ƿ���Ϣ.����%Type;
  v_����         ������Ϣ.����%Type;
  v_���֤��     ������Ϣ.���֤��%Type;
  v_��ѯ�����   Varchar2(20);
  n_��ѯ�����id ����ҽ�ƿ���Ϣ.�����id%Type;
  n_�����id     ҽ�ƿ����.Id%Type;
  v_No           ���˹Һż�¼.No%Type;
  v_���ʽ     ���˹Һż�¼.ҽ�Ƹ��ʽ%Type;
  d_�Һ�ʱ��     ���˹Һż�¼.�Ǽ�ʱ��%Type;
  d_����ʱ��     ���˹Һż�¼.ִ��ʱ��%Type;
  v_�������     ���ű�.����%Type;
  v_����ҽ��     ���˹Һż�¼.ִ����%Type;
  v_����ʷ       ���˹�����¼.ҩ����%Type;
  v_Temp         Varchar2(32767); --��ʱXML 
  x_Templet      Xmltype; --ģ��XML 
  v_Err_Msg      Varchar2(200);
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(a), 'IN/SFZH'), Extractvalue(Value(a), 'IN/YLKLB'), Extractvalue(Value(a), 'IN/YLKH'),
         Extractvalue(Value(a), 'IN/BRXM'), Extractvalue(Value(a), 'IN/MZH'), Extractvalue(Value(a), 'IN/GHDH'),
         Extractvalue(Value(a), 'IN/BRID'), Extractvalue(Value(a), 'IN/CXKLB')
  Into v_���֤��, v_ҽ�ƿ�, v_����, v_����, v_�����, v_�Һŵ�, v_����id, v_��ѯ�����
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) a;

  If v_��ѯ����� Is Not Null Then
    Select Max(Id) Into n_��ѯ�����id From ҽ�ƿ���� Where ���� = v_��ѯ�����;
    If n_��ѯ�����id Is Null Then
      n_��ѯ�����id := To_Number(v_��ѯ�����);
    End If;
  End If;

  If v_����id Is Null Then
  
    If v_���֤�� Is Null And v_ҽ�ƿ� Is Null And v_���� Is Null And v_���� Is Null And v_����� Is Null And v_�Һŵ� Is Null Then
      v_Err_Msg := 'δ�����κ�����,�޷���ɲ�ѯ!';
      Raise Err_Item;
    End If;
  
    If v_ҽ�ƿ� Is Not Null Then
      Select Max(Id) Into n_�����id From ҽ�ƿ���� Where ���� = v_ҽ�ƿ�;
      If n_�����id Is Null Then
        n_�����id := To_Number(v_ҽ�ƿ�);
      End If;
    End If;
  
    If v_�Һŵ� Is Null Then
      If Nvl(n_�����id, 0) = 0 Then
        If Nvl(v_�����, 0) <> 0 Then
          Select ����id Into v_����id From ������Ϣ Where ����� = v_�����;
        Else
          Select Max(����id)
          Into v_����id
          From ������Ϣ
          Where Nvl(���֤��, '-') = Nvl(v_���֤��, Nvl(���֤��, '-')) And ���� = Nvl(v_����, ����);
        End If;
      Else
        Select Max(����id) Into v_����id From ����ҽ�ƿ���Ϣ Where �����id = n_�����id And ���� = v_����;
      End If;
    Else
      Select Max(����id) Into v_����id From ���˹Һż�¼ Where No = v_�Һŵ� And ��¼���� = 1 And ��¼״̬ In (1, 3);
    End If;
  End If;
  If Nvl(v_����id, 0) = 0 Then
    v_Err_Msg := '���ݴ�������,�޷���ɲ�ѯ!';
    Raise Err_Item;
  End If;
  For r_�Һ� In (Select c.����id, c.��ǰ����id, c.�����, c.����, c.�Ա�, c.����, c.����״��, c.����, c.��������, c.���֤��, c.ְҵ, c.ѧ��, c.����, c.��ͥ�绰,
                      c.��ͥ��ַ, c.���ڵ�ַ, c.���, c.�ֻ���, c.��ϵ������, c.��ϵ�˵绰, c.��ϵ�˵�ַ, c.�����ص�, Max(f.����) As ����
               From ������Ϣ c, ����ҽ�ƿ���Ϣ f
               Where c.����id = v_����id And c.����id = f.����id(+) And f.�����id(+) = n_��ѯ�����id And Nvl(f.״̬, 0) = 0
               Group By c.����id, c.��ǰ����id, c.�����, c.����, c.�Ա�, c.��������, c.���֤��, c.ְҵ, c.ѧ��, c.����, c.��ͥ�绰, c.��ͥ��ַ, c.���ڵ�ַ,
                        c.���, c.�ֻ���, c.��ϵ������, c.��ϵ�˵绰, c.��ϵ�˵�ַ, c.�����ص�, c.����, c.����״��, c.����) Loop
    v_Temp := '<BR>';
  
    If v_�Һŵ� Is Null Then
      Select Max(No), Max(ҽ�Ƹ��ʽ), Max(�Ǽ�ʱ��), Max(ִ��ʱ��), Max(ִ����), Max(�������)
      Into v_No, v_���ʽ, d_�Һ�ʱ��, d_����ʱ��, v_����ҽ��, v_�������
      From (Select a.No, a.ҽ�Ƹ��ʽ, a.�Ǽ�ʱ��, a.ִ��ʱ��, a.ִ����, b.���� As �������
             From ���˹Һż�¼ a, ���ű� b
             Where a.ִ�в���id = b.Id(+) And a.����id = r_�Һ�.����id And a.��¼���� = 1 And a.��¼״̬ = 1
             Order By a.�Ǽ�ʱ�� Desc)
      Where Rownum < 2;
    Else
      Select Max(a.No), Max(a.ҽ�Ƹ��ʽ), Max(a.�Ǽ�ʱ��), Max(a.ִ��ʱ��), Max(a.ִ����), Max(b.����) As �������
      Into v_No, v_���ʽ, d_�Һ�ʱ��, d_����ʱ��, v_����ҽ��, v_�������
      From ���˹Һż�¼ a, ���ű� b
      Where a.ִ�в���id = b.Id(+) And a.No = v_�Һŵ� And a.��¼���� = 1 And a.��¼״̬ In (1, 3);
    End If;
  
    For r In (Select ҩ���� From ���˹�����¼ Where ����id = r_�Һ�.����id) Loop
      v_����ʷ := v_����ʷ || ',' || r.ҩ����;
    End Loop;
    v_����ʷ := Substr(v_����ʷ, 2);
  
    v_Temp := v_Temp || '<BRID>' || r_�Һ�.����id || '</BRID>';
    v_Temp := v_Temp || '<XM>' || r_�Һ�.���� || '</XM>';
    v_Temp := v_Temp || '<XB>' || r_�Һ�.�Ա� || '</XB>';
    v_Temp := v_Temp || '<NL>' || r_�Һ�.���� || '</NL>';
    v_Temp := v_Temp || '<CSRQ>' || To_Char(r_�Һ�.��������, 'yyyy-mm-dd hh24:mi:ss') || '</CSRQ>';
    v_Temp := v_Temp || '<MZH>' || r_�Һ�.����� || '</MZH>';
    v_Temp := v_Temp || '<HY>' || r_�Һ�.����״�� || '</HY>';
    v_Temp := v_Temp || '<GJ>' || r_�Һ�.���� || '</GJ>';
    v_Temp := v_Temp || '<MZ>' || r_�Һ�.���� || '</MZ>';
    v_Temp := v_Temp || '<XL>' || r_�Һ�.ѧ�� || '</XL>';
    v_Temp := v_Temp || '<SF>' || r_�Һ�.��� || '</SF>';
    v_Temp := v_Temp || '<ZY>' || r_�Һ�.ְҵ || '</ZY>';
    v_Temp := v_Temp || '<SFZH>' || r_�Һ�.���֤�� || '</SFZH>';
    v_Temp := v_Temp || '<FKFS>' || v_���ʽ || '</FKFS>';
    v_Temp := v_Temp || '<LXFS>' || r_�Һ�.�ֻ��� || '</LXFS>';
    v_Temp := v_Temp || '<LXRXM>' || r_�Һ�.��ϵ������ || '</LXRXM>';
    v_Temp := v_Temp || '<LXRDH>' || r_�Һ�.��ϵ�˵绰 || '</LXRDH>';
    v_Temp := v_Temp || '<LXRDZ>' || r_�Һ�.��ϵ�˵�ַ || '</LXRDZ>';
    v_Temp := v_Temp || '<LXDH>' || r_�Һ�.��ͥ�绰 || '</LXDH>';
    v_Temp := v_Temp || '<XJZDZ>' || r_�Һ�.��ͥ��ַ || '</XJZDZ>';
    v_Temp := v_Temp || '<HJDZ>' || r_�Һ�.���ڵ�ַ || '</HJDZ>';
    v_Temp := v_Temp || '<CSDD>' || r_�Һ�.�����ص� || '</CSDD>';
    v_Temp := v_Temp || '<KSID>' || r_�Һ�.��ǰ����id || '</KSID>';
    v_Temp := v_Temp || '<CXKH>' || r_�Һ�.���� || '</CXKH>';
    v_Temp := v_Temp || '<GMS>' || v_����ʷ || '</GMS>';
    v_Temp := v_Temp || '<GHD>' || v_No || '</GHD>';
    v_Temp := v_Temp || '<GHSJ>' || To_Char(d_�Һ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '</GHSJ>';
    v_Temp := v_Temp || '<JZSJ>' || To_Char(d_����ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '</JZSJ>';
    v_Temp := v_Temp || '<JZKS>' || v_������� || '</JZKS>';
    v_Temp := v_Temp || '<JZYS>' || v_����ҽ�� || '</JZYS>';
    v_Temp := v_Temp || '</BR>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  End Loop;

  Xml_Out := x_Templet;

Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Third_Getpatiinfo;
/





------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0023' Where ���=&n_System;
Commit;
