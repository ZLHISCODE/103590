--ҩ���Զ����ӿ�����ģ��
Insert Into zlPrograms
  (���, ����, ˵��, ϵͳ, ����)
  Select 1348, 'ҩ���Զ����ӿ�', 'HIS��ҩ���Զ��䡢��ҩϵͳ�ӿ�', &n_System, 'zlDrugPacker'
  From Dual
  Where Not Exists (Select 1 From zlPrograms Where ��� = 1348 And ���� = 'ҩ���Զ����ӿ�');

--ҩ���Զ����ӿ�����ģ��
Insert Into zlProgFuncs
  (ϵͳ, ���, ����, ����, ˵��)
  Select &n_System, 1348, '����', Null, Null
  From Dual
  Where Not Exists (Select 1 From zlProgFuncs Where ϵͳ = &n_System And ��� = 1348 And ���� = '����');

--ҩ���Զ����ӿ�����ģ��
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) 
  select &n_System,1348,'����',User,'���ű�','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='���ű�') union all
  select &n_System,1348,'����',User,'������Ա','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='������Ա') union all
  select &n_System,1348,'����',User,'��������˵��','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='��������˵��') union all
  select &n_System,1348,'����',User,'��Ա��','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='��Ա��') union all
  select &n_System,1348,'����',User,'�ϻ���Ա��','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='�ϻ���Ա��') union all
  select &n_System,1348,'����',User,'ҩƷ����','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='ҩƷ����') union all
  select &n_System,1348,'����',User,'ҩ����ҩ�豸','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='ҩ����ҩ�豸') union all
  select &n_System,1348,'����',User,'�Զ���ҩ����','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='�Զ���ҩ����') union all
  select &n_System,1348,'����',User,'ҩ���豸����','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='ҩ���豸����') union all
  select &n_System,1348,'����',User,'���˹Һż�¼','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='���˹Һż�¼') union all
  select &n_System,1348,'����',User,'������Ϣ','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='������Ϣ') union all
  select &n_System,1348,'����',User,'����ҽ������','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='����ҽ������') union all
  select &n_System,1348,'����',User,'����ҽ����¼','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='����ҽ����¼') union all
  select &n_System,1348,'����',User,'������ϼ�¼','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='������ϼ�¼') union all
  select &n_System,1348,'����',User,'��ҩ����','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='��ҩ����') union all
  select &n_System,1348,'����',User,'��Ӧ��','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='��Ӧ��') union all
  select &n_System,1348,'����',User,'������ü�¼','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='������ü�¼') union all
  select &n_System,1348,'����',User,'�շѼ�Ŀ','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='�շѼ�Ŀ') union all
  select &n_System,1348,'����',User,'�շ���Ŀ����','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='�շ���Ŀ����') union all
  select &n_System,1348,'����',User,'�շ���ĿĿ¼','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='�շ���ĿĿ¼') union all
  select &n_System,1348,'����',User,'δ��ҩƷ��¼','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='δ��ҩƷ��¼') union all
  select &n_System,1348,'����',User,'ҩƷ�����޶�','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='ҩƷ�����޶�') union all
  select &n_System,1348,'����',User,'ҩƷ���','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='ҩƷ���') union all
  select &n_System,1348,'����',User,'ҩƷ���','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='ҩƷ���') union all
  select &n_System,1348,'����',User,'ҩƷ������','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='ҩƷ������') union all
  select &n_System,1348,'����',User,'ҩƷ�շ���¼','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='ҩƷ�շ���¼') union all
  select &n_System,1348,'����',User,'ҩƷ����','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='ҩƷ����') union all
  select &n_System,1348,'����',User,'���Ʒ���Ŀ¼','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='���Ʒ���Ŀ¼') union all
  select &n_System,1348,'����',User,'������ĿĿ¼','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='������ĿĿ¼') union all
  select &n_System,1348,'����',User,'סԺ���ü�¼','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='סԺ���ü�¼') union all
  select &n_System,1348,'����',User,'Zl_ҩ����ҩ�豸_Insert','EXECUTE' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='Zl_ҩ����ҩ�豸_Insert') union all
  select &n_System,1348,'����',User,'Zl_ҩ����ҩ�豸_Update','EXECUTE' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='Zl_ҩ����ҩ�豸_Update') union all
  select &n_System,1348,'����',User,'Zl_ҩ����ҩ�豸_Delete','EXECUTE' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='Zl_ҩ����ҩ�豸_Delete') union all
  select &n_System,1348,'����',User,'Zl_ҩ����ҩ�豸_Switch','EXECUTE' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='Zl_ҩ����ҩ�豸_Switch') union all
  select &n_System,1348,'����',User,'Zl_ҩ���豸����_Update','EXECUTE' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='Zl_ҩ���豸����_Update') union all
  select &n_System,1348,'����',User,'Zl_δ��ҩƷ��¼_���䷢ҩ����','EXECUTE' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='Zl_δ��ҩƷ��¼_���䷢ҩ����');
  
--����ϵͳ����
Insert Into zlParameters(ID,ϵͳ,ģ��,������,������,����ֵ,ȱʡֵ,����˵��)
Select zlParameters_ID.Nextval, &n_System,-Null,222, 'ҩ���Զ�����ҩ�ӿ�','0','0','�Ƿ�����ҩ���Զ�����ҩ�ӿڣ�0-��������1-����' From Dual;


--���ݽṹ
--Drop Table ҩ����ҩ�豸;
Create Table ҩ����ҩ�豸(
   Id NUMBER(4),
   ���� VARCHAR2(20),
   ���� VARCHAR2(20),
   �ͺ� VARCHAR2(20),
   ������ VARCHAR2(100),
   ʹ�ò���ID NUMBER(18),
   �������� NUMBER(1),
   �������� VARCHAR2(200),
   ������� NUMBER(1),
   �Ƿ����� NUMBER(1))
   TABLESPACE ZL9MEDLST;
Alter Table ҩ����ҩ�豸 Add Constraint ҩ����ҩ�豸_PK Primary Key (ID) Using Index Tablespace ZL9INDEXHIS;
Alter Table ҩ����ҩ�豸 Add Constraint ҩ����ҩ�豸_UQ_���� Unique (����) Using Index Tablespace ZL9INDEXHIS;
Alter Table ҩ����ҩ�豸 Add constraint ҩ����ҩ�豸_UQ_ʹ�ò���ID unique (ʹ�ò���ID, ����, ����, �ͺ�) using index  tablespace ZL9INDEXHIS;
Alter table ҩ����ҩ�豸 add constraint ҩ����ҩ�豸_FK_ʹ�ò���ID foreign key (ʹ�ò���ID) references ���ű� (ID);

Create Sequence ҩ����ҩ�豸_ID Start With 1;

Create Table �Զ���ҩ����(
    Id NUMBER(4),
    ������ NUMBER(4),
    ������ VARCHAR2(100),
    ����ֵ VARCHAR2(4000),
    ȱʡֵ VARCHAR2(4000),
    ����˵�� VARCHAR2(255))
    TABLESPACE ZL9MEDLST;
Alter Table �Զ���ҩ���� Add Constraint �Զ���ҩ����_PK Primary Key(ID) Using Index PCTFREE 5;
Alter Table �Զ���ҩ���� Add Constraint �Զ���ҩ����_UQ_������ Unique(������) Using Index PCTFREE 5;
Alter Table �Զ���ҩ���� Add Constraint �Զ���ҩ����_UQ_������ Unique(������) Using Index PCTFREE 5;

Insert Into �Զ���ҩ����(ID,������,������,����ֵ,ȱʡֵ,����˵��)
Select 1,1,'ҩƷ����',NULL,NULL,'Null��ʾ����ҩƷ���ͣ������Ҫָ��ĳЩ���ͣ���ʽ��������,Ƭ��,��' From Dual;
  
  
Create Table ҩ���豸����(
   ����ID NUMBER(4),
   �豸ID NUMBER(4),
   ����ֵ VARCHAR2(4000))
   TABLESPACE ZL9MEDLST;
Alter Table ҩ���豸���� Add Constraint ҩ���豸����_PK Primary key (����ID, �豸ID) Using Index Tablespace ZL9INDEXHIS;
Alter Table ҩ���豸���� Add Constraint ҩ���豸����_FK_����ID Foreign key (����ID) references �Զ���ҩ���� (ID);
Alter Table ҩ���豸���� Add Constraint ҩ���豸����_FK_�豸ID Foreign key (�豸ID) references ҩ����ҩ�豸 (ID) On Delete Cascade;



--�豸����
CREATE OR REPLACE Procedure Zl_ҩ����ҩ�豸_Insert
(
  ����_In         In ҩ����ҩ�豸.����%Type,
  ����_In         In ҩ����ҩ�豸.����%Type,
  �ͺ�_In         In ҩ����ҩ�豸.�ͺ�%Type,
  ������_In       In ҩ����ҩ�豸.������%Type,
  ʹ�ò���id_In   In ҩ����ҩ�豸.ʹ�ò���id%Type,
  ��������_In     In ҩ����ҩ�豸.��������%Type,
  ��������_In     In ҩ����ҩ�豸.��������%Type,
  �Ƿ�����_In     In ҩ����ҩ�豸.�Ƿ�����%Type,
  �������_In     In ҩ����ҩ�豸.�������%Type
) Is
  n_�豸id Number;
Begin
  Select ҩ����ҩ�豸_Id.Nextval Into n_�豸id From Dual;

  Insert Into ҩ����ҩ�豸
    (ID, ����, ����, �ͺ�, ������, ʹ�ò���id, ��������, ��������, �������, �Ƿ�����)
  Values
    (n_�豸id, ����_In, ����_In, �ͺ�_In, ������_In, ʹ�ò���id_In, ��������_In, ��������_In, �������_In, �Ƿ�����_In);

  Insert Into ҩ���豸����
    (����id, �豸id, ����ֵ)
    Select 1, n_�豸id, Null From Dual;
    
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩ����ҩ�豸_Insert;
/

--�豸����
CREATE OR REPLACE Procedure Zl_ҩ����ҩ�豸_Update
(
  �豸id_In     In ҩ����ҩ�豸.Id%Type,
  ����_In       In ҩ����ҩ�豸.����%Type,
  ����_In       In ҩ����ҩ�豸.����%Type,
  �ͺ�_In       In ҩ����ҩ�豸.�ͺ�%Type,
  ������_In     In ҩ����ҩ�豸.������%Type,
  ʹ�ò���id_In In ҩ����ҩ�豸.ʹ�ò���id%Type,
  ��������_In   In ҩ����ҩ�豸.��������%Type,
  ��������_In   In ҩ����ҩ�豸.��������%Type,
  �Ƿ�����_In   In ҩ����ҩ�豸.�Ƿ�����%Type,
  �������_In   In ҩ����ҩ�豸.�������%Type
) Is
Begin
  Update ҩ����ҩ�豸
  Set ���� = ����_In, ���� = ����_In, �ͺ� = �ͺ�_In, ������ = ������_In, ʹ�ò���id = ʹ�ò���id_In, �������� = ��������_In, �������� = ��������_In,
      ������� = �������_In, �Ƿ����� = �Ƿ�����_In
  Where ID = �豸id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩ����ҩ�豸_Update;
/

--�豸ɾ��
Create Or Replace Procedure Zl_ҩ����ҩ�豸_Delete
(
  �豸id_In In ҩ����ҩ�豸.Id%Type
) Is
Begin
  Delete ҩ����ҩ�豸 Where ID = �豸id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩ����ҩ�豸_Delete;
/

--�豸����/ͣ������
Create Or Replace Procedure Zl_ҩ����ҩ�豸_Switch
(
  �豸id_In   In ҩ����ҩ�豸.Id%Type,
  �Ƿ�����_In In ҩ����ҩ�豸.�Ƿ�����%Type
) Is
Begin
  Update ҩ����ҩ�豸 Set �Ƿ����� = �Ƿ�����_In Where ID = �豸id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩ����ҩ�豸_Switch;
/

--�豸�����޸�
Create Or Replace Procedure Zl_ҩ���豸����_Update
(
  ����id_In In �Զ���ҩ����.Id%Type,
  �豸id_In In ҩ����ҩ�豸.Id%Type,
  ����ֵ_In In ҩ���豸����.����ֵ%Type
) Is
Begin
  Update ҩ���豸���� Set ����ֵ = ����ֵ_In Where ����id = ����id_In And �豸id = �豸id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩ���豸����_Update;
/

