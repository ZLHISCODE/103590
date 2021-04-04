--ҩ���Զ����ӿ�����ģ��
Insert Into zlPrograms
  (���, ����, ˵��, ϵͳ, ����)
  Select 1348, 'ҩ���Զ����ӿ�', 'HIS��ҩ���Զ��䡢��ҩϵͳ�ӿ�', &n_Syttem, 'zlDrugPacker'
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
  select &n_System,1348,'����',User,'ҩ���豸����','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='ҩ���豸����') union all
  select &n_System,1348,'����',User,'ҩ���豸����','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='ҩ���豸����') union all
  select &n_System,1348,'����',User,'ҩ��ע���豸','SELECT' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='ҩ��ע���豸') union all
  select &n_System,1348,'����',User,'ZL_ҩ���豸����_INSERT','EXECUTE' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='ZL_ҩ���豸����_INSERT') union all
  select &n_System,1348,'����',User,'ZL_ҩ���豸����_UPDATE','EXECUTE' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='ZL_ҩ���豸����_UPDATE') union all
  select &n_System,1348,'����',User,'ZL_ҩ���豸����_DELETE','EXECUTE' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='ZL_ҩ���豸����_DELETE') union all
  select &n_System,1348,'����',User,'ZL_ҩ��ע���豸_INSERT','EXECUTE' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='ZL_ҩ��ע���豸_INSERT') union all
  select &n_System,1348,'����',User,'ZL_ҩ��ע���豸_UPDATE','EXECUTE' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='ZL_ҩ��ע���豸_UPDATE') union all
  select &n_System,1348,'����',User,'ZL_ҩ��ע���豸_DELETE','EXECUTE' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='ZL_ҩ��ע���豸_DELETE') union all
  select &n_System,1348,'����',User,'ZL_ҩ��ע���豸_SWITCH','EXECUTE' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='ZL_ҩ��ע���豸_SWITCH') union all
  select &n_System,1348,'����',User,'ZL_ҩ��ע���豸_SETTING','EXECUTE' from dual where not exists(select 1 from zlprogprivs where ϵͳ=&n_System and ���=1348 and ����='����' and ����='ZL_ҩ��ע���豸_SETTING') ;

--ҩ���Զ����ӿ�����ģ�飨1348��
Insert Into zlParameters(ID,ϵͳ,ģ��,˽��,����,��Ȩ,�̶�,������,������,����ֵ,ȱʡֵ,����˵��)
Select Rownum+B.ID,A.* From (
  Select ϵͳ,ģ��,˽��,����,��Ȩ,�̶�,������,������,����ֵ,ȱʡֵ,����˵�� From zlParameters Where ID=0 Union All
  Select &n_Syttem,1348,0,0,0,0,1,'�������',NULL,'1','ҩ���豸��ִ�е�ҽ���򴦷���1-���2-סԺ' From Dual Union All
  Select &n_Syttem,1348,0,0,0,0,2,'��ҩ��Ӧҵ��',NULL,NULL,'1-�����շѣ�2-������ҩ��ҩ���ܣ�3-������ҩ��ҩ����' From Dual Union All
  Select &n_Syttem,1348,0,0,0,0,3,'���Ͷ�Ӧҵ��',NULL,NULL,'1-����ҩƷ������ҩ' From Dual Union All
  Select &n_Syttem,1348,0,0,0,0,4,'��������',NULL,NULL,'Null��ʾδѡ��ȫ����ʾ�������ֵ��ݣ�1-������2-����;3-���ʵ���' From Dual Union All
  Select &n_Syttem,1348,0,0,0,0,5,'ҩƷ����',NULL,NULL,'Null��ʾδѡ��ȫ����ʾ����ҩƷ���ͣ������Ҫָ��ĳЩ���ͣ���ʽ��������|Ƭ��|��' From Dual
  ) A,(Select Nvl(Max(ID),0) AS ID From zlParameters) B;



create table ҩ���豸����  (
   ID          NUMBER(10)      not null,
   ����        VARCHAR2(20)    not null,
   ��������    NUMBER(2)       not null,
   ��������    VARCHAR2(200)   not null
)
  tablespace ZL9MEDLST;

Alter Table ҩ���豸���� Add Constraint ҩ���豸����_PK Primary Key (ID) Using Index Tablespace ZL9INDEXHIS;
Alter Table ҩ���豸���� Add Constraint ҩ���豸����_UQ_���� Unique (����) Using Index Tablespace ZL9INDEXHIS;

create sequence ҩ���豸����_ID start with 1;




create table ҩ��ע���豸  (
   ID               NUMBER(18)            not null,
   ����             VARCHAR2(20)          not null,
   ����             VARCHAR2(20)          not null,
   �ͺ�             VARCHAR2(20),
   ������           VARCHAR2(100),
   ����ID           NUMBER(18)            not null,
   ����ID           NUMBER(10)            not null,
   ����             NUMBER(1)
)
  tablespace ZL9MEDLST;
 
Alter Table ҩ��ע���豸 Add Constraint ҩ��ע���豸_PK Primary Key (ID) Using Index Tablespace ZL9INDEXHIS;
Alter Table ҩ��ע���豸 Add Constraint ҩ��ע���豸_UQ_���� Unique (����) Using Index Tablespace ZL9INDEXHIS;
Alter Table ҩ��ע���豸 Add constraint ҩ��ע���豸_UQ_����ID unique (����ID, ����, �ͺ�) using index  tablespace ZL9INDEXHIS;

alter table ҩ��ע���豸 add constraint ҩ��ע���豸_FK_����ID foreign key (����ID) references ҩ���豸���� (ID);
alter table ҩ��ע���豸 add constraint ҩ��ע���豸_FK_����ID foreign key (����ID) references ���ű� (ID);

create sequence ҩ��ע���豸_ID start with 1;



create table ҩ���豸����  (
   ����ID             NUMBER(18)                      not null,
   �豸ID             NUMBER(18)                      not null,
   ����ֵ             VARCHAR2(4000)
);

alter table ҩ���豸���� add constraint ҩ���豸����_PK primary key (����ID, �豸ID) Using Index Tablespace ZL9INDEXHIS;

alter table ҩ���豸���� add constraint ҩ���豸����_FK_�豸ID foreign key (�豸ID) references ҩ��ע���豸 (ID);






CREATE OR REPLACE Procedure Zl_ҩ���豸����_Insert
(
  ����_In In ҩ���豸����.����%Type,
  ����_In In ҩ���豸����.��������%Type,
  ����_In In ҩ���豸����.��������%Type
) Is

Begin

  Insert Into ҩ���豸���� 
    (ID, ����, ��������, ��������) 
    Values 
    (ҩ���豸����_Id.Nextval, ����_In, ����_In, ����_In);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩ���豸����_Insert;
/

Create Or Replace Procedure Zl_ҩ���豸����_Delete(Id_In In ҩ���豸����.Id%Type) Is
Begin

  Delete ҩ���豸���� Where ID = Id_In;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩ���豸����_Delete;
/

CREATE OR REPLACE Procedure Zl_ҩ���豸����_Update
(
  Id_In   In ҩ���豸����.Id%Type,
  ����_In In ҩ���豸����.����%Type,
  ����_In In ҩ���豸����.��������%Type,
  ����_In In ҩ���豸����.��������%Type
) Is
Begin

  Update ҩ���豸���� Set ���� = ����_In, �������� = ����_In, �������� = ����_In Where ID = Id_In;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩ���豸����_Update;
/

Create Or Replace Procedure Zl_ҩ��ע���豸_Insert
(
  ����_In         In ҩ��ע���豸.����%Type,
  ����_In         In ҩ��ע���豸.����%Type,
  �ͺ�_In         In ҩ��ע���豸.�ͺ�%Type := Null,
  ������_In       In ҩ��ע���豸.������%Type := Null,
  ����id_In       In ҩ��ע���豸.����id%Type,
  ����id_In       In ҩ��ע���豸.����id%Type,
  ����_In         In ҩ��ע���豸.����%Type := Null,
  �������_In     In ҩ���豸����.����ֵ%Type,
  ��ҩ��Ӧҵ��_In In ҩ���豸����.����ֵ%Type := Null,
  ���Ͷ�Ӧҵ��_In In ҩ���豸����.����ֵ%Type := Null
) Is

  n_�豸id Number;

Begin

  Select ҩ��ע���豸_Id.Nextval Into n_�豸id From Dual;

  Insert Into ҩ��ע���豸
    (ID, ����, ����, �ͺ�, ������, ����id, ����id, ����)
  Values
    (n_�豸id, ����_In, ����_In, �ͺ�_In, ������_In, ����id_In, ����id_In, ����_In);

  Insert Into ҩ���豸����
    (����id, �豸id, ����ֵ)
    Select ID, n_�豸id, �������_In
    From Zlparameters
    Where ϵͳ = &n_Syttem And ģ�� = 1348 And ������ = 1
    Union All
    Select ID, n_�豸id, ��ҩ��Ӧҵ��_In
    From Zlparameters
    Where ϵͳ = &n_Syttem And ģ�� = 1348 And ������ = 2 And ��ҩ��Ӧҵ��_In Is Not Null
    Union All
    Select ID, n_�豸id, ���Ͷ�Ӧҵ��_In
    From Zlparameters
    Where ϵͳ = &n_Syttem And ģ�� = 1348 And ������ = 3 And ���Ͷ�Ӧҵ��_In Is Not Null;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩ��ע���豸_Insert;
/

Create Or Replace Procedure Zl_ҩ��ע���豸_Delete(�豸id_In In ҩ��ע���豸.Id%Type) Is
Begin

  Delete ҩ���豸���� Where �豸id = �豸id_In;

  Delete ҩ��ע���豸 Where ID = �豸id_In;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩ��ע���豸_Delete;
/

Create Or Replace Procedure Zl_ҩ��ע���豸_Switch
(
  �豸id_In In ҩ��ע���豸.Id%Type,
  ����_In   In ҩ��ע���豸.����%Type := Null
) Is
Begin

  Update ҩ��ע���豸 Set ���� = ����_In Where ID = �豸id_In;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩ��ע���豸_Switch;
/

Create Or Replace Procedure Zl_ҩ��ע���豸_Setting
(
  �豸id_In   In ҩ��ע���豸.Id%Type,
  ��������_In In ҩ���豸����.����ֵ%Type := Null,
  ҩƷ����_In In ҩ���豸����.����ֵ%Type := Null
) Is

  n_����id Number;

Begin

  --��������
  Select ID Into n_����id From Zlparameters Where ϵͳ = 100 And ģ�� = 1348 And ������ = 4;
  If Nvl(n_����id, 0) > 0 Then
--    If ��������_In Is Null Then
--      Delete ҩ���豸���� Where ����id = n_����id And �豸id = �豸id_In;
--    Else
      Update ҩ���豸���� Set ����ֵ = ��������_In Where ����id = n_����id And �豸id = �豸id_In;
      If Sql%NotFound Then
        Insert Into ҩ���豸���� (����id, �豸id, ����ֵ) Values (n_����id, �豸id_In, ��������_In);
      End If;
--    End If;
  End If;

  --ҩƷ����
  Select ID Into n_����id From Zlparameters Where ϵͳ = 100 And ģ�� = 1348 And ������ = 5;
  If Nvl(n_����id, 0) > 0 Then
--    If ҩƷ����_In Is Null Then
--      Delete ҩ���豸���� Where ����id = n_����id And �豸id = �豸id_In;
--    Else
      Update ҩ���豸���� Set ����ֵ = ҩƷ����_In Where ����id = n_����id And �豸id = �豸id_In;
      If Sql%NotFound Then
        Insert Into ҩ���豸���� (����id, �豸id, ����ֵ) Values (n_����id, �豸id_In, ҩƷ����_In);
      End If;
--    End If;
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩ��ע���豸_Setting;
/

Create Or Replace Procedure Zl_ҩ��ע���豸_Update
(
  �豸id_In       In ҩ��ע���豸.Id%Type,
  ����_In         In ҩ��ע���豸.����%Type,
  ����_In         In ҩ��ע���豸.����%Type,
  �ͺ�_In         In ҩ��ע���豸.�ͺ�%Type := Null,
  ������_In       In ҩ��ע���豸.������%Type := Null,
  ����id_In       In ҩ��ע���豸.����id%Type,
  ����id_In       In ҩ��ע���豸.����id%Type,
  ����_In         In ҩ��ע���豸.����%Type := Null,
  �������_In     In ҩ���豸����.����ֵ%Type,
  ��ҩ��Ӧҵ��_In In ҩ���豸����.����ֵ%Type := Null,
  ���Ͷ�Ӧҵ��_In In ҩ���豸����.����ֵ%Type := Null
) Is

  n_����id Number;

Begin

  Update ҩ��ע���豸
  Set ���� = ����_In, ���� = ����_In, �ͺ� = �ͺ�_In, ������ = ������_In, ����id = ����id_In, ����id = ����id_In, ���� = ����_In
  Where ID = �豸id_In;

  --�������
  Begin
    Select ID Into n_����id From Zlparameters Where ϵͳ = 100 And ģ�� = 1348 And ������ = 1;
  Exception
    When Others Then
      n_����id := Null;
  End;
  If n_����id Is Not Null Then
    Update ҩ���豸���� Set ����ֵ = �������_In Where ����id = n_����id And �豸id = �豸id_In;
    If Sql%NotFound Then
      Insert Into ҩ���豸����
        (����id, �豸id, ����ֵ)
        Select ID, �豸id_In, �������_In From Zlparameters Where ϵͳ = 100 And ģ�� = 1348 And ������ = 1;
    End If;
  End If;

  --��ҩҵ��
  Begin
    Select ID Into n_����id From Zlparameters Where ϵͳ = 100 And ģ�� = 1348 And ������ = 2;
  Exception
    When Others Then
      n_����id := Null;
  End;
  If n_����id Is Not Null Then
    Update ҩ���豸���� Set ����ֵ = ��ҩ��Ӧҵ��_In Where ����id = n_����id And �豸id = �豸id_In;
    If Sql%NotFound Then
      Insert Into ҩ���豸����
        (����id, �豸id, ����ֵ)
        Select ID, �豸id_In, ��ҩ��Ӧҵ��_In From Zlparameters Where ϵͳ = 100 And ģ�� = 1348 And ������ = 2;
    End If;
  End If;

  --����ҵ��
  Begin
    Select ID Into n_����id From Zlparameters Where ϵͳ = 100 And ģ�� = 1348 And ������ = 3;
  Exception
    When Others Then
      n_����id := Null;
  End;
  If n_����id Is Not Null Then
    Update ҩ���豸���� Set ����ֵ = ���Ͷ�Ӧҵ��_In Where ����id = n_����id And �豸id = �豸id_In;
    If Sql%NotFound Then
      Insert Into ҩ���豸����
        (����id, �豸id, ����ֵ)
        Select ID, �豸id_In, ���Ͷ�Ӧҵ��_In From Zlparameters Where ϵͳ = 100 And ģ�� = 1348 And ������ = 3;
    End If;
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩ��ע���豸_Update;
/
