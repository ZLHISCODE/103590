Create Or Replace Procedure Zl_���鱨�浥_Insert
(
	Id_In   In ����ҽ����¼.Id%Type,
	Type_In In Number -- 0=���� 1=ɾ��
) Is
	--HIS������LIS�ӿ�ʹ��
	v_��ҳid     ����ҽ����¼.��ҳid%Type;
	v_ҽ��id     ����ҽ����¼.Id%Type;
	v_��������id ����ҽ����¼.��������id%Type;
	v_������Դ   ����걾��¼.������Դ%Type;
	v_����id     ����걾��¼.����id%Type;
	v_Ӥ��       ����걾��¼.Ӥ��%Type;
	v_�����ļ�id ��������Ӧ��.�����ļ�id%Type;
	v_�����ļ��� �����ļ��б�.����%Type;
	v_�ļ�id     ���Ӳ�������.�ļ�id%Type;
	v_Temp       Varchar2(255);
	v_��Ա����id ������Ա.����id%Type;
	v_��Ա���   ��Ա��.���%Type;
	v_��Ա����   ��Ա��.����%Type;
	v_ִ��       Number;
	v_No         ����ҽ������.No%Type;
	v_����       ����ҽ������.��¼����%Type;
	v_���       Varchar2(1000);
	v_����       Number;
	v_Error      Varchar2(255);
	Err_Custom Exception;
	--���ҵ�ǰ�걾���������
	Cursor c_Samplequest Is
		Select Distinct Id As ҽ��id From ����ҽ����¼ Where Id_In In (Id, ���id);

	--δ��˵ķ�����(������ҩƷ)
	Cursor c_Verify(v_ҽ��id In Number) Is
		Select Distinct ��¼����, No, ���
		From סԺ���ü�¼
		Where �շ���� Not In ('5', '6', '7') And
					ҽ����� + 0 In (Select Id From ����ҽ����¼ Where v_ҽ��id In (Id, ���id)) And ���ʷ��� = 1 And
					��¼״̬ = 0 And �۸񸸺� Is Null And
					(��¼����, No) In
					(Select ��¼����, No
					 From ����ҽ������
					 Where ҽ��id = v_ҽ��id
					 Union All
					 Select ��¼����, No
					 From ����ҽ������
					 Where ҽ��id In (Select Id From ����ҽ����¼ Where v_ҽ��id In (Id, ���id)))
		Order By ��¼����, No, ���;

Begin
	--����Ա��Ϣ:����ID,��������;��ԱID,��Ա���,��Ա����
	v_Temp       := Zl_Identity;
	v_��Ա����id := To_Number(Substr(v_Temp, 1, Instr(v_Temp, ',') - 1));
	v_Temp       := Substr(v_Temp, Instr(v_Temp, ';') + 1);
	v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
	v_��Ա���   := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
	v_��Ա����   := Substr(v_Temp, Instr(v_Temp, ',') + 1);

	Select Distinct Nvl(b.��ҳid, 0), Nvl(b.���id, 0), Decode(b.������Դ, 2, 2, 4, 4, 1), Nvl(b.����id, 0),
									Nvl(b.��������id, 0), Nvl(b.Ӥ��, 0)
	Into v_��ҳid, v_ҽ��id, v_������Դ, v_����id, v_��������id, v_Ӥ��
	From ����ҽ����¼ b
	Where b.���id = Id_In;

	Begin
		Select �����ļ�id, c.����
		Into v_�����ļ�id, v_�����ļ���
		From ����ҽ����¼ a, ��������Ӧ�� b, �����ļ��б� c
		Where a.������Ŀid = b.������Ŀid And b.�����ļ�id = c.Id And a.���id = v_ҽ��id And b.Ӧ�ó��� = v_������Դ And
					Rownum <= 1;
	Exception
		When Others Then
			Return;
	End;

	If Type_In = 0 Then
		--����
		--ɾ����ǰ�ı����¼
		Begin
			Select ����id Into v_�ļ�id From ����ҽ������ Where ҽ��id = v_ҽ��id And Rownum <= 1;
			Delete ���Ӳ�����¼ Where Id = v_�ļ�id;
			Delete ���Ӳ������� Where �ļ�id = v_�ļ�id;
		Exception
			When Others Then
				Select ���Ӳ�����¼_Id.Nextval Into v_�ļ�id From Dual;
				--Insert Into ����ҽ������ (ҽ��id, ����id) Values (v_ҽ��id, v_�ļ�id);
		End;
	
		Insert Into ���Ӳ�����¼
			(Id, ������Դ, ����id, ��ҳid, Ӥ��, ����id, ��������, �ļ�id, ��������, ������, ����ʱ��, ������, ����ʱ��,
			 ���汾, ǩ������)
		Values
			(v_�ļ�id, v_������Դ, v_����id, v_��ҳid, v_Ӥ��, v_��������id, 7, v_�����ļ�id, v_�����ļ���, Null, Sysdate,
			 Null, Sysdate, 1, 0);
	
		Insert Into ����ҽ������ (ҽ��id, ����id) Values (v_ҽ��id, v_�ļ�id);
	
		Insert Into ���Ӳ�������
			(Id, �ļ�id, ��ʼ��, ��ֹ��, ��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���)
		Values
			(���Ӳ�������_Id.Nextval, v_�ļ�id, 1, 1, Null, 1, 2, Null, Null, 0, 0, 0, 0);
	
		Update ����ҽ������ Set ִ��״̬ = 1 Where ҽ��id In (Select Id From ����ҽ����¼ Where v_ҽ��id In (Id, ���id));
	
		--ִ�к��Զ���˶�Ӧ�ļ��ʻ��۵�(������ҩƷ)
		Select Zl_To_Number(Nvl(Zl_Getsysparameter(81), '0')) Into v_ִ�� From Dual;
		--2.��鵱ǰ�걾��ص��������ر걾�Ƿ�������
		For r_Samplequest In c_Samplequest Loop
		
			--r_SampleQuest.ҽ��id�����Ѿ����,�����������
		
			--2.����ִ�д���
			IF If v_���� = 1 Then
			Update ������ü�¼
			Set ִ��״̬ = 1, ִ��ʱ�� = Sysdate, ִ���� = v_��Ա����
			Where �շ���� Not In ('5', '6', '7') And
						(ҽ�����, ��¼����, No) In
						(Select ҽ��id, ��¼����, No
						 From ����ҽ������
						 Where ҽ��id = r_Samplequest.ҽ��id
						 Union All
						 Select ҽ��id, ��¼����, No
						 From ����ҽ������
						 Where ҽ��id In (Select Id From ����ҽ����¼ Where r_Samplequest.ҽ��id In (Id, ���id)));
			 ELSE 
			Update סԺ���ü�¼
			Set ִ��״̬ = 1, ִ��ʱ�� = Sysdate, ִ���� = v_��Ա����
			Where �շ���� Not In ('5', '6', '7') And
						(ҽ�����, ��¼����, No) In
						(Select ҽ��id, ��¼����, No
						 From ����ҽ������
						 Where ҽ��id = r_Samplequest.ҽ��id
						 Union All
						 Select ҽ��id, ��¼����, No
						 From ����ҽ������
						 Where ҽ��id In (Select Id From ����ҽ����¼ Where r_Samplequest.ҽ��id In (Id, ���id)));
		         END if;
			--3.�Զ���˼���
			If Nvl(v_ִ��, 0) = 1 Then
				For r_Verify In c_Verify(r_Samplequest.ҽ��id) Loop
					If r_Verify.No || ',' || r_Verify.��¼���� <> v_No || ',' || v_���� Then
						If v_��� Is Not Null Then
							If v_���� = 1 Then
								Zl_������ʼ�¼_Verify(v_No, v_��Ա���, v_��Ա����, Substr(v_���, 2));
							Elsif v_���� = 2 Then
								Zl_סԺ���ʼ�¼_Verify(v_No, v_��Ա���, v_��Ա����, Substr(v_���, 2));
							End If;
						End If;
						v_��� := Null;
					End If;
					v_No   := r_Verify.No;
					v_���� := r_Verify.��¼����;
					v_��� := v_��� || ',' || r_Verify.���;
				End Loop;
				If v_��� Is Not Null Then
					If v_���� = 1 Then
						Zl_������ʼ�¼_Verify(v_No, v_��Ա���, v_��Ա����, Substr(v_���, 2));
					Elsif v_���� = 2 Then
						Zl_סԺ���ʼ�¼_Verify(v_No, v_��Ա���, v_��Ա����, Substr(v_���, 2));
					End If;
				End If;
			End If;
		
		End Loop;
	Else
		--ɾ��
	
		v_���� := 0;
		Select Nvl(����״̬, 0) Into v_���� From ����ҽ������ Where ҽ��id = v_ҽ��id;
		If v_���� = 0 Then
			Select ����id Into v_�ļ�id From ����ҽ������ Where ҽ��id = v_ҽ��id And Rownum <= 1;
			Delete ����ҽ������ Where ҽ��id = v_ҽ��id;
			Delete ���Ӳ�����¼ Where Id = v_�ļ�id;
			Delete ���Ӳ������� Where �ļ�id = v_�ļ�id;
			Update ����ҽ������
			Set ִ��״̬ = 0
			Where ҽ��id In (Select Id From ����ҽ����¼ Where v_ҽ��id In (Id, ���id));
			For r_Samplequest In c_Samplequest Loop
				--2.����ִ�д���
				If v_���� = 1 Then
				Update ������ü�¼
				Set ִ��״̬ = 0, ִ��ʱ�� = Null, ִ���� = Null
				Where �շ���� Not In ('5', '6', '7') And
							(ҽ�����, ��¼����, No) In
							(Select ҽ��id, ��¼����, No
							 From ����ҽ������
							 Where ҽ��id = r_Samplequest.ҽ��id
							 Union All
							 Select ҽ��id, ��¼����, No
							 From ����ҽ������
							 Where ҽ��id In (Select Id From ����ҽ����¼ Where r_Samplequest.ҽ��id In (Id, ���id)));
				ELSE 
				Update סԺ���ü�¼
				Set ִ��״̬ = 0, ִ��ʱ�� = Null, ִ���� = Null
				Where �շ���� Not In ('5', '6', '7') And
							(ҽ�����, ��¼����, No) In
							(Select ҽ��id, ��¼����, No
							 From ����ҽ������
							 Where ҽ��id = r_Samplequest.ҽ��id
							 Union All
							 Select ҽ��id, ��¼����, No
							 From ����ҽ������
							 Where ҽ��id In (Select Id From ����ҽ����¼ Where r_Samplequest.ҽ��id In (Id, ���id)));
				END if;
			End Loop;
		Else
			v_Error := '�ñ����Ѿ���ҽ�����ģ�����ȡ��������ϵҽ����';
			Raise Err_Custom;
		End If;
	End If;
Exception
	When Err_Custom Then
		Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
	When Others Then
		Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_���鱨�浥_Insert;
/
Create Or Replace Procedure Zl_���Ӳ�����ʽ_Insert
(
  Id_In   In ���Ӳ�����ʽ.�ļ�id%Type,
  Txt_In  In Varchar2,
  ��ʼ_In In Number -- 1=��ʼ
) Is
  l_Blob Blob;
Begin

  If ��ʼ_In = 1 Then
    Delete ���Ӳ�����ʽ Where �ļ�id = Id_In;
  End If;
  If ��ʼ_In = 1 Then
    Update ���Ӳ�����ʽ Set ���� = Empty_Blob() Where �ļ�id = Id_In;
    If Sql%Rowcount = 0 Then
      Insert Into ���Ӳ�����ʽ (�ļ�id, ����) Values (Id_In, Empty_Blob());
    End If;
  End If;
  Select ���� Into l_Blob From ���Ӳ�����ʽ Where �ļ�id = Id_In For Update;
  Dbms_Lob.Writeappend(l_Blob, Length(Hextoraw(Txt_In)) / 2, Hextoraw(Txt_In));
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_���Ӳ�����ʽ_Insert;
/
Create Or Replace Procedure Zl_����ҽ�����_Edit
(
  Id_In   In ����ҽ����¼.Id%Type,
  Type_In In Number -- 1=���� 0=ȡ������
) Is
Begin
  Update ����ҽ������ Set ִ��״̬ = Type_In Where ҽ��id In (Select ID From ����ҽ����¼ Where Id_In In (ID, ���id));
  Update ������ü�¼
  Set ִ��״̬ = Type_In, ִ��ʱ�� = Null, ִ���� = Null
  Where ҽ����� In (Select ID From ����ҽ����¼ Where ������Դ<>2 AND Id_In In (ID, ���id));
Update סԺ���ü�¼
  Set ִ��״̬ = Type_In, ִ��ʱ�� = Null, ִ���� = Null
  Where ҽ����� In (Select ID From ����ҽ����¼ Where  ������Դ=2 AND Id_In In (ID, ���id));
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����ҽ�����_Edit;
/

--  2009-09-21 �������ָ�걣�����
Create Or Replace Procedure Zl_���ָ��_Externaledit
(
	����id_In     In ���������.����id%Type,
	����id_In     In ���������.����id%Type,
	�嵥id_In     In ���������.�嵥id%Type,
	���ָ��id_In In ���������.���ָ��id%Type,
	������_In     In ���������.�����%Type,
	����ʱ��_In   In ���������.���ʱ��%Type,
	���_In       In ���������.���%Type,
	��λ_In       In ���������.��λ%Type,
	�ο�_In       In ���������.�ο�%Type,
	����_In       In ���������.����%Type
) Is
Begin

	Update ���������
	Set ��� = ���_In, ���� = ����_In, ��λ = ��λ_In, �ο� = �ο�_In, ����� = ������_In, ���ʱ�� = ����ʱ��_In
	Where ����id = ����id_In And ����id = ����id_In And �嵥id = �嵥id_In And ���ָ��id = ���ָ��id_In;

	Update ���������
	Set ������ = ������_In, ����ʱ�� = ����ʱ��_In, ִ��״̬ = 1
	Where ����id = ����id_In And ����id = ����id_In And �嵥id = �嵥id_In;
Exception
	When Others Then
		Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_���ָ��_Externaledit;
/