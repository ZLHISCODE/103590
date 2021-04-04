Create Or Replace Procedure Zl_��λ״����¼_Update
(
  ����id_In     In ��λ״����¼.����id%Type,
  ���_In       In ��λ״����¼.���%Type,
  �շ�ϸĿid_In In ��λ״����¼.�շ�ϸĿid%Type,
  ״̬_In       In ��λ״����¼.״̬%Type,
  ��ע_In       In ��λ״����¼.��ע%Type
) Is
Begin
  Update ��λ״����¼
  Set �շ�ϸĿid = �շ�ϸĿid_In, ״̬ = ״̬_In, ��ע = ��ע_In
  Where ����id = ����id_In And ��� = ���_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_��λ״����¼_Update;
/

CREATE OR REPLACE Procedure Zl_��λ״����¼_Insert
(
  ����id_In     In ��λ״����¼.����id%Type,
  ���_In       In ��λ״����¼.���%Type,
  ���_In       In ��λ״����¼.���%Type,
  ״̬_In       In ��λ״����¼.״̬%Type,
  �շ�ϸĿid_In In ��λ״����¼.�շ�ϸĿid%Type,
  ��ע_In       In ��λ״����¼.��ע%Type
) Is
Begin
  Insert Into ��λ״����¼
    (����id, ���, ���, ״̬, �շ�ϸĿid, ��ע)
  Values
    (����id_In, ���_In, ���_In, ״̬_In, �շ�ϸĿid_In, ��ע_In);
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_��λ״����¼_Insert;
/

Create Or Replace Procedure Zl_��λ״����¼_Delete
(
  ����id_In In ��λ״����¼.����id%Type,
  ���_In   In ��λ״����¼.���%Type
) Is
Begin
  Delete ��λ״����¼ Where nvl(����id,0) = 0 And ״̬ <> 1 And ����id = ����id_In And ��� = ���_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_��λ״����¼_Delete;
/

Create Or Replace Procedure Zl_��λ״����¼_Setseating
(
  ����id_In In ��λ״����¼.����id%Type,
  ���_In   In ��λ״����¼.���%Type,
  ���_In   In ��λ״����¼.���%Type,
  ����id_In In ��λ״����¼.����id%Type,
  NO_In   In ��λ״����¼.NO%Type
) Is
Begin
  If ����id_In <> 0 Then
    -- ռ��
    Update ��λ״����¼
    Set ����id = ����id_In, ״̬ = 1, NO = NO_In
    Where ����id = ����id_In And ��� = ���_In And ��� = ���_In And Nvl(״̬, 0) = 0 And Nvl(����id, 0) = 0;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_��λ״����¼_Setseating;
/

Create Or Replace Procedure Zl_��λ״����¼_Clear
(
  ����id_In In ��λ״����¼.����id%Type,
  ���_In   In ��λ״����¼.���%Type
) Is
Begin
  Update ��λ״����¼
  Set ����id = Null, ״̬ = 0, NO = Null
  Where Nvl(����id, 0) <> 0 And ״̬ = 1 And ����id = ����id_In And ��� = ���_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_��λ״����¼_Clear;
/

Create Or Replace Procedure Zl_����ҽ��ִ��_Transfusion
(
  ҽ��id_In   In ����ҽ��ִ��.ҽ��id%Type,
  ���ͺ�_In   In ����ҽ��ִ��.���ͺ�%Type,
  ִ��ʱ��_In In ����ҽ��ִ��.ִ��ʱ��%Type,
  ��ˮ��_In   In ����ҽ��ִ��.��ˮ��%Type,
  ��ҩ��_In   In ����ҽ��ִ��.��ҩ��%Type,
  ����_In     In ����ҽ��ִ��.����%Type,
  ���_In     In ����ҽ��ִ��.���%Type,
  ����_In     In ����ҽ��ִ��.����%Type,
  ��ϵ��_In   In ����ҽ��ִ��.��ϵ��%Type,
  Һ����_In   In ����ҽ��ִ��.Һ����%Type,
  ˵��_In     In ����ҽ��ִ��.˵��%Type,
  �ӵ���_In   In ����ҽ��ִ��.�ӵ���%Type,
  ��ʱ_In     In ����ҽ��ִ��.��ʱ%Type,
  ����_In     In ����ҽ��ִ��.����%Type
) Is

Begin

  Update ����ҽ��ִ��
  Set ��ˮ�� = ��ˮ��_In, ��ҩ�� = ��ҩ��_In, ���� = ����_In, ��� = ���_In, ���� = ����_In, ��ϵ�� = ��ϵ��_In,
      Һ���� = Һ����_In, ˵�� = ˵��_In, �ӵ��� = �ӵ���_In, ��ʱ = ��ʱ_In, ���� = ����_In
  Where ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In And ִ��ʱ�� = ִ��ʱ��_In;

Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����ҽ��ִ��_Transfusion;
/

Create Or Replace Procedure Zl_����ҽ��ִ��_Modify
(
  ��ˮ��_In In ����ҽ��ִ��.��ˮ��%Type,
  ҽ��id_In In ����ҽ��ִ��.ҽ��id%Type,
  ���ͺ�_In In ����ҽ��ִ��.���ͺ�%Type,
  ����_In   In ����ҽ��ִ��.����%Type,
  Һ����_In In ����ҽ��ִ��.Һ����%Type,
  ��ϵ��_In In ����ҽ��ִ��.��ϵ��%Type,
  ��ʱ_In   In ����ҽ��ִ��.��ʱ%Type,
  ˵��_In   In ����ҽ��ִ��.˵��%Type
) Is
  --����Ҫִ�е�����¼,�������˸�������,��鲿λ�ļ�¼ 
  --��������,��ҩ�巨,�ɼ�������������,�������ֻ��д�ڵ�һ����Ŀ�ϣ���ִ��״̬��ͬ 
  Cursor c_Advice Is
    Select A.ҽ��id, B.���id, B.�������
    From ����ҽ������ A, ����ҽ����¼ B
    Where (B.ID = ҽ��id_In Or (B.���id = ҽ��id_In And B.������� In ('F', 'D'))) And A.ҽ��id = B.ID And
          A.���ͺ� + 0 = ���ͺ�_In;

  v_Temp     Varchar2(255);
  v_��Ա���� ���˷��ü�¼.����Ա����%Type;
  v_Date  Date;
Begin
  v_Temp     := Zl_Identity;
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);

  Select Sysdate Into v_Date From Dual;

  For r_Advice In c_Advice Loop
    Update ����ҽ��ִ��
    Set ���� = ����_In, Һ���� = Һ����_In, ��ϵ�� = ��ϵ��_In, ��ʱ = ��ʱ_In, ˵�� = ˵��_In, �Ǽ�ʱ�� = v_Date,
        �Ǽ��� = v_��Ա����
    Where ҽ��id = r_Advice.ҽ��id And ���ͺ� + 0 = ���ͺ�_In And ��ˮ�� = ��ˮ��_In;
  End Loop;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����ҽ��ִ��_Modify;
/

Create Or Replace Procedure Zl_�ŶӼ�¼_Addqueue
(
	����id_In In �ŶӼ�¼.����id%Type,
	����id_In In �ŶӼ�¼.����id%Type,
	˳���_In In �ŶӼ�¼.˳���%Type
) Is

Begin
	-- һ��������һ������ֻ����һ���ŶӼ�¼ ,����,��ɾ���ÿ���ԭ�����ŶӼ�¼,��д���¼�¼.
	Delete �ŶӼ�¼ Where ����id = ����id_In And ����id = ����id_In;
	Insert Into �ŶӼ�¼
		(����id, ����id, ˳���, ��Ȩ��, ״̬, ��ע, ����)
	Values
		(����id_In, ����id_In, ˳���_In, 0, 1, '', Sysdate);
Exception
	When Others Then
		Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�ŶӼ�¼_Addqueue;
/

CREATE OR REPLACE Procedure Zl_�ŶӼ�¼_Update
(
	����id_In In �ŶӼ�¼.����id%Type,
	����id_In In �ŶӼ�¼.����id%Type,
	˳���_In In �ŶӼ�¼.˳���%Type,
	��Ȩ��_In In �ŶӼ�¼.��Ȩ��%Type,
	״̬_In   In �ŶӼ�¼.״̬%Type
) Is

Begin
	Update �ŶӼ�¼
	Set ��Ȩ�� = ��Ȩ��_In, ״̬ = ״̬_In
	Where ����id = ����id_In And ����id = ����id_In And ˳��� = ˳���_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);

End Zl_�ŶӼ�¼_Update;
/

Create Or Replace Procedure Zl_�ݴ�ҩƷ��¼_Insert
(
	No_In       In �ݴ�ҩƷ��¼.No%Type,
	���_In     In �ݴ�ҩƷ��¼.���%Type,
	����id_In   In �ݴ�ҩƷ��¼.����id %Type,
	ҽ��id_In   In �ݴ�ҩƷ��¼.ҽ��id%Type,
	���ͺ�_In   In �ݴ�ҩƷ��¼.���ͺ�%Type,
	ҩƷid_In   In �ݴ�ҩƷ��¼.ҩƷid %Type,
	ҩƷ����_In In �ݴ�ҩƷ��¼.ҩƷ����%Type,
	���_In     In �ݴ�ҩƷ��¼.���%Type,
	ִ�з���_In In �ݴ�ҩƷ��¼.ִ�з���%Type,
	ʹ��״̬_In In �ݴ�ҩƷ��¼.ʹ��״̬%Type,
	ժҪ_In     In �ݴ�ҩƷ��¼.ժҪ%Type,
	���ϵ��_In In �ݴ�ҩƷ��¼.���ϵ��%Type,
	��λ_In     In �ݴ�ҩƷ��¼.��λ%Type,
	����_In     In �ݴ�ҩƷ��¼.����%Type,
	����_In     In �ݴ�ҩƷ��¼.����%Type,
	����_In     In �ݴ�ҩƷ��¼.����%Type,
	���_In     In �ݴ�ҩƷ��¼.���%Type,
	����Ա_In   In �ݴ�ҩƷ��¼.����Ա%Type,
	����id_In   In �ݴ�ҩƷ��¼.����id%Type,
	�Ǽ�ʱ��_In In �ݴ�ҩƷ��¼.�Ǽ�ʱ��%Type
) Is
Begin
	Insert Into �ݴ�ҩƷ��¼
		(No, ���, ����id, ҽ��id, ���ͺ�, ҩƷid, ҩƷ����, ���, ִ�з���, ʹ��״̬, ժҪ, ���ϵ��, ��λ, ����, ����,
		 ����, ���, ����Ա, �Ǽ�ʱ��, ����id)
	Values
		(No_In, ���_In, ����id_In, ҽ��id_In, ���ͺ�_In, ҩƷid_In, ҩƷ����_In, ���_In, ִ�з���_In, ʹ��״̬_In,
		 ժҪ_In, ���ϵ��_In, ��λ_In, ����_In, ����_In, ����_In, ���_In, ����Ա_In, �Ǽ�ʱ��_In, ����id_In);
	-- �޸� ʹ��״̬
	If ���ϵ��_In = -1 Then
		Update �ݴ�ҩƷ��¼
		Set ʹ��״̬ = 1
		Where No = No_In And ��� = ���_In And ����id = ����id_In And ҽ��id = ҽ��id_In And ���ͺ� = ���ͺ�_In And
					ҩƷid = ҩƷid_In;
	End If;
Exception
	When Others Then
		Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�ݴ�ҩƷ��¼_Insert;
/

Create Or Replace Procedure Zl_�ݴ�ҩƷ��¼_Delete(No_In In �ݴ�ҩƷ��¼.NO%Type) Is
Begin
  Delete �ݴ�ҩƷ��¼ Where NO = No_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�ݴ�ҩƷ��¼_Delete;
/

Create Or Replace Procedure Zl_�ݴ�ҩƷ��¼_Undouse
(
  No_In       In �ݴ�ҩƷ��¼.NO%Type,
  ���_In     In �ݴ�ҩƷ��¼.���%Type,
  ���ϵ��_In In �ݴ�ҩƷ��¼.���ϵ��%Type,
  �Ǽ�ʱ��_In In �ݴ�ҩƷ��¼.�Ǽ�ʱ��%Type
) Is
  n_Use �ݴ�ҩƷ��¼.����%Type;
Begin
  Delete �ݴ�ҩƷ��¼ Where NO = No_In And ��� = ���_In And ���ϵ�� = ���ϵ��_In And �Ǽ�ʱ�� = �Ǽ�ʱ��_In;
  Select Sum(Nvl(����, 0)) Into n_Use From �ݴ�ҩƷ��¼ Where NO = No_In And ��� = ���_In And ���ϵ�� = ���ϵ��_In;
  If Nvl(n_Use, 0) = 0 Then
    Update �ݴ�ҩƷ��¼ Set ʹ��״̬ = 0 Where NO = No_In And ��� = ���_In And ���ϵ�� = 1;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�ݴ�ҩƷ��¼_Undouse;
/

Create Or Replace Procedure Zl_�ݴ�ҩƷ��¼_Adviceused
(
	No_In       In �ݴ�ҩƷ��¼.No%Type,
	���_In     In �ݴ�ҩƷ��¼.���%Type,
	ҽ��id_In   In �ݴ�ҩƷ��¼.ҽ��id%Type,
	���ͺ�_In   In �ݴ�ҩƷ��¼.���ͺ�%Type,
	ҩƷid_In   In �ݴ�ҩƷ��¼.ҩƷid %Type,
	����_In     In �ݴ�ҩƷ��¼.����%Type,
	����Ա_In   In �ݴ�ҩƷ��¼.����Ա%Type,
	�Ǽ�ʱ��_In In �ݴ�ҩƷ��¼.�Ǽ�ʱ��%Type
) Is
Begin
	Insert Into �ݴ�ҩƷ��¼
		(No, ���, ����id, ҽ��id, ���ͺ�, ҩƷid, ҩƷ����, ���, ִ�з���, ʹ��״̬, ժҪ, ���ϵ��, ��λ, ����, ����,
		 ����, ���, ����Ա, �Ǽ�ʱ��, ����id)
		Select b.No, b.���, b.����id, b.ҽ��id, b.���ͺ�, b.ҩƷid, b.ҩƷ����, b.���, b.ִ�з���, 1, b.ժҪ, -1, b.��λ,
					 b.����, ����_In, b.����, ����_In * b.����, ����Ա_In, �Ǽ�ʱ��_In, b.����id
		From �ݴ�ҩƷ��¼ b
		Where b.���ϵ�� = 1 And Nvl(b.ʹ��״̬, 0) = 0 And b.ҩƷid = ҩƷid_In And b.ҽ��id = ҽ��id_In And
					b.���ͺ� = ���ͺ�_In;

	Update �ݴ�ҩƷ��¼
	Set ʹ��״̬ = 1
	Where No = No_In And ��� = ���_In And ҽ��id = ҽ��id_In And ���ͺ� = ���ͺ�_In And ҩƷid = ҩƷid_In;
Exception
	When Others Then
		Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�ݴ�ҩƷ��¼_Adviceused;
/
-------------------------------------------------------------------------------------------------------------------
-- �������޸�ԭ�еĹ���
-------------------------------------------------------------------------------------------------------------------
Create Or Replace Procedure Zl_����ҽ��ִ��_Insert
(
  ҽ��id_In   ����ҽ��ִ��.ҽ��id%Type,
  ���ͺ�_In   ����ҽ��ִ��.���ͺ�%Type,
  Ҫ��ʱ��_In ����ҽ��ִ��.Ҫ��ʱ��%Type,
  ��������_In ����ҽ��ִ��.��������%Type,
  ִ��ժҪ_In ����ҽ��ִ��.ִ��ժҪ%Type,
  ִ����_In   ����ҽ��ִ��.ִ����%Type,
  ִ��ʱ��_In ����ҽ��ִ��.ִ��ʱ��%Type
) Is
  --����Ҫִ�е�����¼,�������˸�������,��鲿λ�ļ�¼ 
  --��������,��ҩ�巨,�ɼ�������������,�������ֻ��д�ڵ�һ����Ŀ�ϣ���ִ��״̬��ͬ 
  Cursor c_Advice Is
    Select A.ҽ��id, B.���id, B.�������
    From ����ҽ������ A, ����ҽ����¼ B
    Where (B.ID = ҽ��id_In Or (B.���id = ҽ��id_In And B.������� In ('F', 'D'))) And A.ҽ��id = B.ID And
          A.���ͺ� + 0 = ���ͺ�_In;

  v_Temp Varchar2(255);
  --v_��Ա��� ���˷��ü�¼.����Ա���%Type;
  v_��Ա���� ���˷��ü�¼.����Ա����%Type;

  v_Date  Date;
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  v_Temp := Zl_Identity;
  v_Temp := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  --v_��Ա��� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);

  Select Sysdate Into v_Date From Dual;

  For r_Advice In c_Advice Loop
    Insert Into ����ҽ��ִ��
      (ҽ��id, ���ͺ�, Ҫ��ʱ��, ��������, ִ��ժҪ, ִ����, ִ��ʱ��, �Ǽ�ʱ��, �Ǽ���)
    Values
      (r_Advice.ҽ��id, ���ͺ�_In, Ҫ��ʱ��_In, ��������_In, ִ��ժҪ_In, ִ����_In, ִ��ʱ��_In, v_Date, v_��Ա����);
  
    --��д��ִ��״̬��ͱ��Ϊ����ִ�� 
    If r_Advice.������� = 'C' And r_Advice.���id Is Not Null Then
      Update ����ҽ������
      Set ִ��״̬ = 3
      Where ���ͺ� + 0 = ���ͺ�_In And ҽ��id In (Select ID From ����ҽ����¼ Where ���id = r_Advice.���id);
    Else
      Update ����ҽ������ Set ִ��״̬ = 3 Where ҽ��id = r_Advice.ҽ��id And ���ͺ� + 0 = ���ͺ�_In;
    End If;
    --Beging 2007-01-04 ɾ��ʱ����Ƿ��ü�¼�е�ִ��״̬���������˷�
    Update ���˷��ü�¼
    Set ִ��״̬ = 2, ִ��ʱ�� = ִ��ʱ��_In, ִ���� = v_��Ա����
    Where Nvl(ִ��״̬, 0) = 0 And �շ���� Not In ('5', '6', '7') And
          ҽ����� + 0 In
          (Select ID
           From ����ҽ����¼
           Where ID = Nvl(r_Advice.ҽ��id, r_Advice.���id) Or ���id = Nvl(r_Advice.ҽ��id, r_Advice.���id)) And
          (��¼����, NO) In
          (Select ��¼����, NO
           From ����ҽ������
           Where ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In
           Union All
           Select ��¼����, NO From ����ҽ������ Where ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In);
    --End 2007-01-04 ɾ��ʱ����Ƿ��ü�¼�е�ִ��״̬���������˷�  
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����ҽ��ִ��_Insert;
/

Create Or Replace Procedure Zl_����ҽ��ִ��_Delete
(
  ҽ��id_In   ����ҽ��ִ��.ҽ��id%Type,
  ���ͺ�_In   ����ҽ��ִ��.���ͺ�%Type,
  ִ��ʱ��_In ����ҽ��ִ��.ִ��ʱ��%Type
) Is
  --����Ҫִ�е�����¼,�������˸�������,��鲿λ�ļ�¼ 
  --��������,��ҩ�巨,�ɼ�������������,�������ֻ��д�ڵ�һ����Ŀ�ϣ���ִ��״̬��ͬ 
  Cursor c_Advice Is
    Select A.ҽ��id, B.���id, B.�������
    From ����ҽ������ A, ����ҽ����¼ B
    Where (B.ID = ҽ��id_In Or (B.���id = ҽ��id_In And B.������� In ('F', 'D'))) And A.ҽ��id = B.ID And
          A.���ͺ� + 0 = ���ͺ�_In;

  v_Count Number;
Begin
  For r_Advice In c_Advice Loop
    Delete From ����ҽ��ִ�� Where ҽ��id = r_Advice.ҽ��id And ���ͺ� + 0 = ���ͺ�_In And ִ��ʱ�� = ִ��ʱ��_In;
    --Beging 2007-01-04 ɾ��ʱ��������ü�¼�е�ִ��״̬�������˷�
    Update ���˷��ü�¼
    Set ִ��״̬ = 0, ִ��ʱ�� = Null, ִ���� = Null
    Where Nvl(ִ��״̬, 0) = 2 And �շ���� Not In ('5', '6', '7') And
          ҽ����� + 0 In
          (Select ID
           From ����ҽ����¼
           Where ID = Nvl(r_Advice.ҽ��id, r_Advice.���id) Or ���id = Nvl(r_Advice.ҽ��id, r_Advice.���id)) And
          (��¼����, NO) In
          (Select ��¼����, NO
           From ����ҽ������
           Where ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In
           Union All
           Select ��¼����, NO From ����ҽ������ Where ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In);
    --End 2007-01-04 ɾ��ʱ��������ü�¼�е�ִ��״̬�������˷�  
  End Loop;

  --���ִ�����ɾ���˾ͱ��ִ��״̬Ϊδִ�� 
  Select Count(*) Into v_Count From ����ҽ��ִ�� Where ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In;
  If Nvl(v_Count, 0) = 0 Then
    For r_Advice In c_Advice Loop
      If r_Advice.������� = 'C' And r_Advice.���id Is Not Null Then
        Update ����ҽ������
        Set ִ��״̬ = 0
        Where ���ͺ� + 0 = ���ͺ�_In And ҽ��id In (Select ID From ����ҽ����¼ Where ���id = r_Advice.���id);
      Else
        Update ����ҽ������ Set ִ��״̬ = 0 Where ҽ��id = r_Advice.ҽ��id And ���ͺ� + 0 = ���ͺ�_In;
      End If;
    End Loop;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����ҽ��ִ��_Delete;
/

