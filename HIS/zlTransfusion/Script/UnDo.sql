Alter Table ������ĿĿ¼ drop Column ִ�з���;
Alter Table ҩƷ��� drop column ����;
Alter Table ����ҽ��ִ�� drop column ��ˮ��;
Alter Table ����ҽ��ִ�� drop column �ӵ���;
Alter Table ����ҽ��ִ�� drop column ��ҩ��;
Alter Table ����ҽ��ִ�� drop column ����;
Alter Table ����ҽ��ִ�� drop column ���;
Alter Table ����ҽ��ִ�� drop column ����;
Alter Table ����ҽ��ִ�� drop column ��ϵ��;
Alter Table ����ҽ��ִ�� drop column Һ����;
Alter Table ����ҽ��ִ�� drop column ��ʱ;
Alter Table ����ҽ��ִ�� drop column ����;
Alter Table ����ҽ��ִ�� drop column ˵��;

drop table ִ�д�ӡ��¼;
drop table �ݴ�ҩƷ��¼;
drop table ��λ״����¼;
drop table �ŶӼ�¼;

drop sequence ����ҽ��ִ��_��ˮ��;

drop procedure Zl_��λ״����¼_Update;
drop procedure Zl_��λ״����¼_Insert;
drop procedure Zl_��λ״����¼_Delete;
drop procedure Zl_��λ״����¼_Setseating;
drop procedure Zl_��λ״����¼_Clear;
drop procedure Zl_����ҽ��ִ��_Transfusion;
drop procedure Zl_����ҽ��ִ��_Modify;
drop procedure Zl_�ŶӼ�¼_Addqueue;
drop procedure Zl_�ŶӼ�¼_Update;
drop procedure Zl_�ݴ�ҩƷ��¼_Insert;
drop procedure Zl_�ݴ�ҩƷ��¼_Delete;
drop procedure Zl_�ݴ�ҩƷ��¼_Undouse;
drop procedure Zl_�ݴ�ҩƷ��¼_Adviceused;

--
delete zlComponent where ����='zl9Transfusion';
delete zlPrograms where ���=1264;
delete zlProgFuncs where ���=1264;
delete zlProgPrivs where ���=1264;
delete zlMenus where ����='������Һע�����';
delete ������Ʊ� where ��Ŀ���=19;

Delete zlNotices where ��������='[����][����]ʱ���ѵ�����鿴�����' And ϵͳ=100;

-- Ƥ������

--------------------------
-- ��ԭ����(10.16.0)
CREATE OR REPLACE Procedure ZL_����ҽ��ִ��_Insert(
	ҽ��ID_IN		����ҽ��ִ��.ҽ��ID%Type,
	���ͺ�_IN		����ҽ��ִ��.���ͺ�%Type,
	Ҫ��ʱ��_IN		����ҽ��ִ��.Ҫ��ʱ��%Type,
	��������_IN		����ҽ��ִ��.��������%Type,
	ִ��ժҪ_IN		����ҽ��ִ��.ִ��ժҪ%Type,
	ִ����_IN		����ҽ��ִ��.ִ����%Type,
	ִ��ʱ��_IN		����ҽ��ִ��.ִ��ʱ��%Type
--������ҽ��ID_IN=����ִ�е�ҽ��ID���������Ϊ��ʾ�ļ�����Ŀ��ID��
) IS
	--����Ҫִ�е�����¼,�������˸�������,��鲿λ�ļ�¼
	--��������,��ҩ�巨,�ɼ�������������,�������ֻ��д�ڵ�һ����Ŀ�ϣ���ִ��״̬��ͬ
	Cursor c_Advice IS
		Select A.ҽ��ID,B.���ID,B.�������
		From ����ҽ������ A,����ҽ����¼ B
		Where (B.ID=ҽ��ID_IN Or (B.���ID=ҽ��ID_IN And B.������� IN('F','D')))
			And A.ҽ��ID=B.ID And A.���ͺ�+0=���ͺ�_IN;

    v_Temp			Varchar2(255);
    v_��Ա���		���˷��ü�¼.����Ա���%Type;
    v_��Ա����		���˷��ü�¼.����Ա����%Type;

	v_Date			Date;
    v_Error			Varchar2(255);
    Err_Custom		Exception;
Begin
    --��ǰ������Ա
    v_Temp:=zl_Identity;
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_��Ա���:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_��Ա����:=Substr(v_Temp,Instr(v_Temp,',')+1);

	Select Sysdate Into v_Date From Dual;

    --����ҽ��ִ��
	For r_Advice In c_Advice Loop
		Insert Into ����ҽ��ִ��(
			ҽ��ID,���ͺ�,Ҫ��ʱ��,��������,ִ��ժҪ,ִ����,ִ��ʱ��,�Ǽ�ʱ��,�Ǽ���)
		Values(
			r_Advice.ҽ��ID,���ͺ�_IN,Ҫ��ʱ��_IN,��������_IN,ִ��ժҪ_IN,ִ����_IN,ִ��ʱ��_IN,v_Date,v_��Ա����);

		--��д��ִ��״̬��ͱ��Ϊ����ִ��
		If r_Advice.�������='C' And r_Advice.���ID IS Not NULL Then
			Update ����ҽ������ 
				Set ִ��״̬=3 
			Where ���ͺ�+0=���ͺ�_IN And ҽ��ID IN(
				Select ID From ����ҽ����¼ Where ���ID=r_Advice.���ID);
		Else
			Update ����ҽ������ Set ִ��״̬=3 Where ҽ��ID=r_Advice.ҽ��ID And ���ͺ�+0=���ͺ�_IN;
		End IF;
	End Loop;
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then Zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_����ҽ��ִ��_Insert;
/

CREATE OR REPLACE Procedure ZL_����ҽ��ִ��_Delete(
	ҽ��ID_IN		����ҽ��ִ��.ҽ��ID%Type,
	���ͺ�_IN		����ҽ��ִ��.���ͺ�%Type,
	ִ��ʱ��_IN		����ҽ��ִ��.ִ��ʱ��%Type
--������ҽ��ID_IN=����ִ�е�ҽ��ID���������Ϊ��ʾ�ļ�����Ŀ��ID��
) IS
	--����Ҫִ�е�����¼,�������˸�������,��鲿λ�ļ�¼
	--��������,��ҩ�巨,�ɼ�������������,�������ֻ��д�ڵ�һ����Ŀ�ϣ���ִ��״̬��ͬ
	Cursor c_Advice IS
		Select A.ҽ��ID,B.���ID,B.�������
		From ����ҽ������ A,����ҽ����¼ B
		Where (B.ID=ҽ��ID_IN Or (B.���ID=ҽ��ID_IN And B.������� IN('F','D')))
			And A.ҽ��ID=B.ID And A.���ͺ�+0=���ͺ�_IN;

	v_Count			Number;
Begin
    --����ҽ��ִ��
	For r_Advice In c_Advice Loop
		Delete From ����ҽ��ִ�� Where ҽ��ID=r_Advice.ҽ��ID And ���ͺ�+0=���ͺ�_IN And ִ��ʱ��=ִ��ʱ��_IN;
	End Loop;

	--���ִ�����ɾ���˾ͱ��ִ��״̬Ϊδִ��
	Select Count(*) Into v_Count From ����ҽ��ִ�� Where ҽ��ID=ҽ��ID_IN And ���ͺ�+0=���ͺ�_IN;
	If Nvl(v_Count,0)=0 Then
		For r_Advice In c_Advice Loop
			If r_Advice.�������='C' And r_Advice.���ID IS Not NULL Then
				Update ����ҽ������ 
					Set ִ��״̬=0
				Where ���ͺ�+0=���ͺ�_IN And ҽ��ID IN(
					Select ID From ����ҽ����¼ Where ���ID=r_Advice.���ID);
			Else
				Update ����ҽ������ Set ִ��״̬=0 Where ҽ��ID=r_Advice.ҽ��ID And ���ͺ�+0=���ͺ�_IN;
			End IF;
		End Loop;
	End IF;
Exception
    When OTHERS Then Zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_����ҽ��ִ��_Delete;
/

commit;