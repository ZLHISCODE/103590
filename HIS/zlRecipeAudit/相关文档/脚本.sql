----------------
--ϵͳ�ű�
----------------
Insert Into zlBakTables(ϵͳ,���,����,���,ֱ��ת��,ͣ�ô�����)
Select &n_System,2,A.* From (
Select ����,���,ֱ��ת��,ͣ�ô����� From zlBakTables Where 1 = 0 Union All 
  Select '���������ϸ',14,1,-Null From Dual Union All 
  Select '���������',15,1,-Null From Dual Union All 
  Select '��������¼',18,1,-Null From Dual Union All 
Select ����,���,ֱ��ת��,ͣ�ô����� From ZLBAKTABLES Where 1 = 0) A;

Insert Into zlComponent(����,����,���汾,�ΰ汾,���汾,ϵͳ,ע���Ʒ����,ע���Ʒ����,ע���Ʒ�汾) Values('zl9RecipeAudit','������鲿��',10,35,0,&n_System,'����ҽԺ��Ϣϵͳ','ZLHIS+','10');

Insert Into zlPrograms(���,����,˵��,ϵͳ,����) Values(1351,'���ﴦ�����','ҩ��ʦ������ҽ���¿��Ĵ���������飬ֻ��ͨ���������շѺ��䷢ҩ��',&n_System,'zl9RecipeAudit'); 
Insert Into zlPrograms(���,����,˵��,ϵͳ,����) Values(1352,'סԺҩ�����','ҩ��ʦ��סԺҽ���¿���ҩ��������飬ֻ��ͨ���������䷢ҩ��',&n_System,'zl9RecipeAudit'); 
Insert Into zlPrograms(���,����,˵��,ϵͳ,����) Values(1353,'���������Ŀ','ȷ�����ﴦ����סԺҩ������Ҫ�����Щ��Ŀ��',&n_System,'zl9RecipeAudit'); 
Insert Into zlPrograms(���,����,˵��,ϵͳ,����) Values(1354,'�����������','���ﴦ�����Ĵ�����������������������ȡ�Ϳ�չ��顣',&n_System,'zl9RecipeAudit'); 
Insert Into zlPrograms(���,����,˵��,ϵͳ,����) Values(1355,'�������ͳ��','�ɷֱ������ﴦ����סԺҩ����ͳ�����ݣ����������',&n_System,'zl9RecipeAudit'); 

------
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��,ȱʡֵ)
Select &n_System,1351,A.* From (
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0 Union All
  Select '����',-NULL,NULL,1 From Dual Union All
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0) A;

Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��,ȱʡֵ)
Select &n_System,1352,A.* From (
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0 Union All
  Select '����',-NULL,NULL,1 From Dual Union All
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0) A;

Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��,ȱʡֵ)
Select &n_System,1353,A.* From (
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0 Union All
  Select '����',-NULL,NULL,1 From Dual Union All
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0) A;

Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��,ȱʡֵ)
Select &n_System,1354,A.* From (
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0 Union All
  Select '����',-NULL,NULL,1 From Dual Union All
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0) A;

Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��,ȱʡֵ)
Select &n_System,1355,A.* From (
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0 Union All
  Select '����',-NULL,NULL,1 From Dual Union All
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0) A;

------
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1351,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All
  Select '������Ϣ','SELECT' From Dual Union All
  Select '����ҽ����¼','SELECT' From Dual Union All
  Select '����ҽ������','SELECT' From Dual Union All
  Select '������Ա','SELECT' From Dual Union All
  Select '��������˵��','SELECT' From Dual Union All
  Select '������鳣������','SELECT' From Dual Union All
  Select '����������','SELECT' From Dual Union All  
  Select '���������Ŀ','SELECT' From Dual Union All
  Select '�����������','SELECT' From Dual Union All
  Select '��������¼','SELECT' From Dual Union All
  Select '���������ϸ','SELECT' From Dual Union All
  Select '���������','SELECT' From Dual Union All
  Select '������ü�¼','SELECT' From Dual Union All
  Select '�շ���ĿĿ¼','SELECT' From Dual Union All
  Select '�շ���Ŀ����','SELECT' From Dual Union All
  Select 'ZL_������鳣������_UPDATE','EXECUTE' From Dual Union All
  Select 'ZL_�������_AUDIT','EXECUTE' From Dual Union All
  Select 'ZL_�������_AUDIT_DETAIL','EXECUTE' From Dual Union All
  Select 'ZL_����������_SAVE','EXECUTE' From Dual Union All
  Select 'ZL_��������¼_LOCK','EXECUTE' From Dual Union All
  Select 'ZL_ҵ����Ϣ�嵥_INSERT','EXECUTE' From Dual Union All
  Select 'ZL_FUN_PATI_CALORIE','EXECUTE' From Dual Union All
  Select '���˹Һż�¼','SELECT' From Dual Union All
  Select '���˹�����¼','SELECT' From Dual Union All
  Select '������ϼ�¼','SELECT' From Dual Union All
  Select '��������Ŀ¼','SELECT' From Dual Union All
  Select '�������Ŀ¼','SELECT' From Dual Union All
  Select '������ĿĿ¼','SELECT' From Dual Union All
  Select 'ҩƷ���','SELECT' From Dual Union All
  Select '�����������','SELECT' From Dual Union All
  Select '���˻����¼','SELECT' From Dual Union All
  Select '���˻�������','SELECT' From Dual Union All
  Select '����Ƶ����Ŀ','SELECT' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1352,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All
  Select '������Ϣ','SELECT' From Dual Union All
  Select '����ҽ����¼','SELECT' From Dual Union All
  Select '����ҽ������','SELECT' From Dual Union All
  Select '������Ա','SELECT' From Dual Union All
  Select '��������˵��','SELECT' From Dual Union All
  Select '������鳣������','SELECT' From Dual Union All
  Select '����������','SELECT' From Dual Union All  
  Select '���������Ŀ','SELECT' From Dual Union All
  Select '��������¼','SELECT' From Dual Union All
  Select '���������ϸ','SELECT' From Dual Union All
  Select '���������','SELECT' From Dual Union All
  Select '�շ���ĿĿ¼','SELECT' From Dual Union All
  Select '�շ���Ŀ����','SELECT' From Dual Union All
  Select 'ZL_������鳣������_UPDATE','EXECUTE' From Dual Union All
  Select 'ZL_�������_AUDIT','EXECUTE' From Dual Union All
  Select 'ZL_�������_AUDIT_DETAIL','EXECUTE' From Dual Union All
  Select 'ZL_����������_SAVE','EXECUTE' From Dual Union All
  Select 'ZL_��������¼_LOCK','EXECUTE' From Dual Union All
  Select 'ZL_ҵ����Ϣ�嵥_INSERT','EXECUTE' From Dual Union All
  Select 'ZL_FUN_PATI_CALORIE','EXECUTE' From Dual Union All
  Select '������ҳ','SELECT' From Dual Union All
  Select '������ҳ�ӱ�','SELECT' From Dual Union All
  Select '���˹�����¼','SELECT' From Dual Union All
  Select '������ϼ�¼','SELECT' From Dual Union All
  Select '��������Ŀ¼','SELECT' From Dual Union All
  Select '�������Ŀ¼','SELECT' From Dual Union All
  Select '������ĿĿ¼','SELECT' From Dual Union All
  Select 'ҩƷ���','SELECT' From Dual Union All
  Select '�����������','SELECT' From Dual Union All
  Select '���˻����¼','SELECT' From Dual Union All
  Select '���˻�������','SELECT' From Dual Union All
  Select '����Ƶ����Ŀ','SELECT' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1353,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All
  Select '���������Ŀ','SELECT' From Dual Union All
  Select '���������Ŀ_ID','SELECT' From Dual Union All
  Select 'ZL_���������Ŀ_UPDATE','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1354,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All
  Select '��������˵��','SELECT' From Dual Union All
  Select '��Ա����˵��','SELECT' From Dual Union All
  Select '���Ʒ���Ŀ¼','SELECT' From Dual Union All
  Select '������ĿĿ¼','SELECT' From Dual Union All
  Select 'ҩƷ����','SELECT' From Dual Union All
  Select '��������Ŀ¼','SELECT' From Dual Union All
  Select '�������Ŀ¼','SELECT' From Dual Union All
  Select '�����������','SELECT' From Dual Union All
  Select 'ZL_�����������_UPDATE','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1355,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All
  Select '��������˵��','SELECT' From Dual Union All
  Select '��Ա����˵��','SELECT' From Dual Union All
  Select '����ҽ����¼','SELECT' From Dual Union All
  Select '�շ���ĿĿ¼','SELECT' From Dual Union All
  Select '���������Ŀ','SELECT' From Dual Union All
  Select '��������¼','SELECT' From Dual Union All
  Select '���������ϸ','SELECT' From Dual Union All
  Select '���������','SELECT' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

--[[zlModuleRelas]]
Insert Into zlModuleRelas(ϵͳ,ģ��,����,���ϵͳ,���ģ��,�������,��ع���,ȱʡֵ)
Select &n_System,1351,A.* From (
Select ����,���ϵͳ,���ģ��,�������,��ع���,ȱʡֵ From zlModuleRelas Where 1 = 0 Union All
  Select NULL,&n_System,9001,1,'����',1 From Dual Union All
Select ����,���ϵͳ,���ģ��,�������,��ع���,ȱʡֵ From zlModuleRelas Where 1 = 0) A;

Insert Into zlModuleRelas(ϵͳ,ģ��,����,���ϵͳ,���ģ��,�������,��ع���,ȱʡֵ)
Select &n_System,1352,A.* From (
Select ����,���ϵͳ,���ģ��,�������,��ع���,ȱʡֵ From zlModuleRelas Where 1 = 0 Union All
  Select NULL,&n_System,9001,1,'����',1 From Dual Union All
Select ����,���ϵͳ,���ģ��,�������,��ع���,ȱʡֵ From zlModuleRelas Where 1 = 0) A;

--[[zlProgRelas]]
--��


------
Insert Into zlMenus
  (���, ID, �ϼ�id, ����, ���, ˵��, ϵͳ, ģ��, �̱���, ͼ��)
  Select ���, Zlmenus_Id.Nextval, ID, '�������ϵͳ', Null, 'ҩ��ʦ��ҽ���¿������ﴦ����סԺҩ������ϵͳ��', &n_System, -Null, '�������ϵͳ', ͼ��
  From zlMenus
  Where ���� = 'ҽ�ƹ���ϵͳ' And ��� = 'ȱʡ' And ϵͳ = &n_System And ģ�� Is Null;

Insert Into zlMenus(���,ID,�ϼ�ID,����,���,˵��,ϵͳ,ģ��,�̱���,ͼ��) Select A.���,ZlMenus_ID.Nextval,A.ID,B.* From (
Select ���,ID From zlMenus Where ���� = '�������ϵͳ' And ��� = 'ȱʡ' And ϵͳ = &n_System And ģ�� Is Null) A,(Select ����,���,˵��,ϵͳ,ģ��,�̱���,ͼ�� From zlMenus Where 1 = 0 Union All
  Select '���ﴦ�����', Null,'ҩ��ʦ������ҽ���¿��Ĵ���������飬ֻ��ͨ���������շѺͷ�ҩ��' ,&n_System,1351,'���ﴦ�����' ,232 From Dual Union All
  Select 'סԺҩ�����', Null,'ҩ��ʦ��סԺҽ���¿���ҩ��������飬ֻ��ͨ�������ܼƷѺͷ�ҩ��' ,&n_System,1352,'סԺҩ�����' ,234 From Dual Union All
  Select '���������Ŀ', Null,'ȷ�����ﴦ����סԺҩ������Ҫ�����Щ��Ŀ��' ,&n_System,1353,'���������Ŀ' ,193 From Dual Union All
  Select '�����������', Null,'���ﴦ�����Ĵ�����������������������ȡ�Ϳ�չ��顣' ,&n_System,1354,'�����������' ,210 From Dual Union All
  Select '�������ͳ��', Null,'�ɷֱ������ﴦ����סԺҩ����ͳ�����ݣ����������' ,&n_System,1355,'�������ͳ��' ,179 From Dual Union All
Select ����,���,˵��,ϵͳ,ģ��,�̱���,ͼ�� From zlMenus Where 1 = 0) B;
  
Insert Into zlParameters(ID,ϵͳ,ģ��,˽��,����,��Ȩ,�̶�,����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵��)
Select zlParameters_ID.Nextval,&n_System,-Null,A.* From (
Select ˽��,����,��Ȩ,�̶�,����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵�� From zlParameters Where 1 = 0 Union All 
  Select 0,0,0,0,0,0,241,'�������',Null,'0','�Ƿ����ô������ϵͳ���������̿��ƣ�','0-�����סԺ�������ã�1-�������ã�סԺ�����ã�2-���ﲻ���ã�סԺ���ã�3-�����סԺ������',Null,Null,Null From Dual Union All
  Select 0,0,0,0,0,0,242,'������ʱ��',Null,'1','ȷ������ҩʦ�󷽵Ľ���ʱ����','1-��������ǰ��2-ҩ���䷢ҩǰ',Null,Null,Null From Dual Union All 
  Select 0,0,0,0,0,0,243,'����ҩʦ���ʱ��',Null,'10','����ҩʦ�����ʱ������λ���ӣ������������趨ʱ��ֵδ���Ĵ�����ҽʦ�ɷ���ͨ�������ⲡ�˳�ʱ�������ٴ����һ�ҩ����',Null,Null,Null,Null From Dual Union All
  Select 0,0,0,0,0,0,244,'�����������',Null,'1','ȷ�����ﴦ��/סԺҩ�����������ʲô��չ��','1-���ݡ�������������淶��28�2-���ݡ���������취��7��',Null,Null,Null From Dual Union All
  Select 0,0,0,0,0,0,245,'��������ҽ�����ϸ�ҽ��',Null,'0','����ҽ������������Ƿ�������ҽ���������ҩ����','0-�����ѣ�1-����',Null,Null,Null From Dual Union All
  Select 0,0,0,0,0,0,246,'����סԺҽ�����ϸ�ҽ��',Null,'0','סԺҽ������������Ƿ�������ҽ���������ҩ����','0-�����ѣ�1-����',Null,Null,Null From Dual Union All
Select ˽��,����,��Ȩ,�̶�,����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵�� From zlParameters Where 1 = 0) A;

----------------
--���ݽṹ
----------------
Create Sequence �����������_Id Start With 1;
Create Sequence ��������¼_Id Start With 1;
Create Sequence ���������Ŀ_Id Start With 1;

Create Table ����������(
       ������ Varchar2(15), 
       ������� Number(1),
       �Ƿ����� Number(1), 
       ������ʱ�� Date,
       ��Դ���� Varchar2(4000)) 
Pctfree 10 Initrans 1 
Tablespace Zl9baseitem;

Create Table �����������(
       ID Number(18), 
       ��� Number(2), 
       ҩ��id Number(18),
       ����id Number(18), 
       ҽ��id Number(18), 
       ���id Number(18), 
       ����id Number(18)) 
Pctfree 10 Initrans 1 
Tablespace Zl9baseitem;

Create Table ���������Ŀ(
       ID Number(18), 
       ��� Number(1),
       ���� Varchar2(10), 
       ��� Varchar2(50),
       ���� Varchar2(500), 
       �Ƿ��������� Number(1), 
       �Ƿ�סԺ���� Number(1), 
       ������� Number(1),
       PASS��� Varchar2(50),
       ������ Varchar2(100), 
       ����ʱ�� Date, 
       ����ʱ�� Date)
Pctfree 10 Initrans 1 
Tablespace Zl9baseitem;

Create Table ������鳣������(
       �û��� Varchar2(20),
       ���� Varchar2(500))
Pctfree 10 Initrans 1 
Tablespace Zl9medlst;

Create Table ��������¼(
       ID Number(18),
       ����id Number(18), 
       �Һ�id Number(18), 
       ��ҳid Number(18),
       �ύ�� Varchar2(100), 
       �ύʱ�� Date, 
       �ύ����id Number(18),
       ����� Number(1),
       ����� Varchar2(100), 
       ���ʱ�� Date, 
       ��ҩҩ��id Number(18), 
       �ۺ����� Varchar2(500),
       ״̬ Number(1),
       �����û� Varchar2(20),
       ����ʱ�� Date,
       ��ת�� Number(3)) 
Pctfree 10 Initrans 20 
Tablespace Zl9medlst;

Create Table ���������ϸ(
       ��id Number(18), 
       ҽ��id Number(18), 
       ����ύ Number(1), 
       ��ת�� Number(3)) 
Pctfree 10 Initrans 20 
Tablespace Zl9medlst;

Create Table ���������(
       ��id Number(18), 
       ҽ��id Number(18), 
       �����Ŀid Number(18),
       ����ύ Number(1),
       ҩʦ��� Number(1),
       �Զ���� Number(1),
       ���� Varchar2(500),
       ��ת�� Number(3)) 
Pctfree 10 Initrans 20 
Tablespace Zl9medlst;

Alter Table ���������� Add Constraint ����������_Pk Primary Key(������, �������) Using Index Tablespace Zl9indexhis;
Alter Table ����������� Add Constraint �����������_Pk Primary Key(ID) Using Index Tablespace Zl9indexhis;
--Alter Table ����������� Add Constraint �����������_Uq_����id Unique(����id, ���) Using Index Tablespace Zl9indexhis;
--Alter Table ����������� Add Constraint �����������_Uq_ҽ��id Unique(ҽ��id, ���) Using Index Tablespace Zl9indexhis;
--Alter Table ����������� Add Constraint �����������_Uq_���id Unique(���id, ���) Using Index Tablespace Zl9indexhis;
--Alter Table ����������� Add Constraint �����������_Uq_����id Unique(����id, ���) Using Index Tablespace Zl9indexhis;
--Alter Table ����������� Add Constraint �����������_Uq_ҩ��id Unique(ҩ��id, ���) Using Index Tablespace Zl9indexhis;
Alter Table ���������Ŀ Add Constraint ���������Ŀ_Pk Primary Key(ID) Using Index Tablespace Zl9indexhis;
Alter Table ��������¼ Add Constraint ��������¼_Pk Primary Key(ID) Using Index Tablespace Zl9indexhis;
Alter Table ���������ϸ Add Constraint ���������ϸ_Pk Primary Key(��id, ҽ��id) Using Index Tablespace Zl9indexhis;
Alter Table ������鳣������ Add Constraint ������鳣������_Pk Primary Key(�û���, ����) Using Index Tablespace Zl9indexhis;

Alter Table ���������Ŀ Add Constraint ���������Ŀ_Uq_���� Unique(����) Using Index Tablespace Zl9indexhis;
Alter Table ���������Ŀ Add Constraint ���������Ŀ_Uq_��� Unique(���) Using Index Tablespace Zl9indexhis;
Alter Table ��������� Add Constraint ���������_Uq_��Id Unique(��id, ҽ��id, �����Ŀid) Using Index Tablespace Zl9indexhis;
Alter Table ��������¼ Add Constraint ��������¼_Uq_�ύʱ�� Unique(�ύʱ��, �ύ����id, ����id, ��ҩҩ��id) Using Index Tablespace Zl9indexhis;

Alter Table ����������� Modify ��� Constraint �����������_NN_��� not null;
Alter Table ���������Ŀ Modify ���� Constraint ���������Ŀ_NN_���� Not Null;
Alter Table ���������Ŀ Modify ��� Constraint ���������Ŀ_NN_��� Not Null;
Alter Table ��������� Modify ��Id Constraint ���������_NN_��Id not null;

Alter Table ����������� Add Constraint �����������_Fk_����ID Foreign Key(����id) References ���ű�(ID) On Delete Cascade enable novalidate;
Alter Table ����������� Add Constraint �����������_Fk_ҽ��ID Foreign Key(ҽ��id) References ��Ա��(ID) On Delete Cascade enable novalidate;
Alter Table ����������� Add Constraint �����������_Fk_���ID Foreign Key(���id) References �������Ŀ¼(ID) On Delete Cascade enable novalidate;
Alter Table ����������� Add Constraint �����������_Fk_����ID Foreign Key(����id) References ��������Ŀ¼(ID) On Delete Cascade enable novalidate;
Alter Table ����������� Add Constraint �����������_Fk_ҩ��ID Foreign Key(ҩ��id) References ������ĿĿ¼(ID) On Delete Cascade enable novalidate;
Alter Table ��������¼ Add Constraint ��������¼_FK_����Id Foreign Key(����Id) References ������Ϣ(����ID) enable novalidate;
Alter Table ��������¼ Add Constraint ��������¼_FK_�Һ�Id Foreign Key(�Һ�Id) References ���˹Һż�¼(ID) enable novalidate;
Alter Table ��������¼ Add Constraint ��������¼_FK_�ύ����Id Foreign Key(�ύ����Id) References ���ű�(ID) enable novalidate;
Alter Table ��������¼ Add Constraint ��������¼_FK_��ҩҩ��Id Foreign Key(��ҩҩ��Id) References ���ű�(ID) enable novalidate;
Alter Table ���������ϸ Add Constraint ���������ϸ_Fk_��id Foreign Key(��id) References ��������¼(ID) On Delete Cascade enable novalidate;
Alter Table ���������ϸ Add Constraint ���������ϸ_Fk_ҽ��id Foreign Key(ҽ��id) References ����ҽ����¼(ID) enable novalidate;
Alter Table ��������� Add Constraint ���������_Fk_��id Foreign Key(��id) References ��������¼(ID) On Delete Cascade enable novalidate;
Alter Table ��������� Add Constraint ���������_Fk_�����Ŀid Foreign Key(�����Ŀid) References ���������Ŀ(ID) enable novalidate;
Alter Table ��������� Add Constraint ���������_Fk_ҽ��id Foreign Key(ҽ��id) References ����ҽ����¼(ID) Enable Novalidate;

Create Index ��������¼_Ix_�Һ�id On ��������¼(�Һ�id) Tablespace Zl9indexhis;
Create Index ��������¼_Ix_����id On ��������¼(����id, ��ҳid) Tablespace Zl9indexhis;
Create Index ��������¼_Ix_���ʱ�� On ��������¼(���ʱ��) Tablespace Zl9indexhis;
Create Index ��������¼_Ix_״̬ On ��������¼(״̬) Tablespace Zl9indexhis;
Create Index ��������¼_Ix_�����û� On ��������¼(�����û�) Tablespace Zl9indexhis;
Create Index ��������¼_Ix_��ת�� On ��������¼(��ת��) Tablespace Zl9indexhis;
Create Index ���������ϸ_Ix_ҽ��id On ���������ϸ(ҽ��id) Tablespace Zl9indexhis;
Create Index ���������ϸ_IX_��ת�� ON ���������ϸ(��ת��) Tablespace Zl9indexhis;
Create Index ���������_Ix_�����Ŀid On ���������(�����Ŀid) Tablespace Zl9indexhis;
Create Index ���������_IX_��ת�� ON ���������(��ת��) Tablespace Zl9indexhis;
  




CREATE OR REPLACE Procedure Zl_�����������_Update
(
  ���_In In �����������.���%Type,
  ���_In In Number,
  Ids_In  In Varchar2
) Is

  --v_Err_Msg Varchar2(2000);
  --Err_Item Exception;

Begin

  If ���_In = 1 Then
    Delete �����������;
  End If;

  If ���_In = 1 Then
    --����
    Insert Into �����������
      (ID, ���, ����id)
      Select �����������_Id.Nextval, ���_In, Column_Value From Table(f_Num2list(Ids_In));
  Elsif ���_In = 2 Then
    --ҽ��
    Insert Into �����������
      (ID, ���, ҽ��id)
      Select �����������_Id.Nextval, ���_In, Column_Value From Table(f_Num2list(Ids_In));
  Elsif ���_In = 3 Then
    --���
    Insert Into �����������
      (ID, ���, ���id)
      Select �����������_Id.Nextval, ���_In, Column_Value From Table(f_Num2list(Ids_In));
  Elsif ���_In = 4 Then
    --����
    Insert Into �����������
      (ID, ���, ����id)
      Select �����������_Id.Nextval, ���_In, Column_Value From Table(f_Num2list(Ids_In));
  Elsif ���_In = 5 Then
    --ҩ��
    Insert Into �����������
      (ID, ���, ҩ��id)
      Select �����������_Id.Nextval, ���_In, Column_Value From Table(f_Num2list(Ids_In));
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����������_Update;
/

Create Or Replace Procedure Zl_���������Ŀ_Update
(
  ��Ŀid_In       In ���������Ŀ.Id%Type,
  ���_In         In ���������Ŀ.���%Type,
  ����_In         In ���������Ŀ.����%Type,
  ���_In         In ���������Ŀ.���%Type,
  ����_In         In ���������Ŀ.����%Type,
  �Ƿ���������_In In ���������Ŀ.�Ƿ���������%Type,
  �Ƿ�סԺ����_In In ���������Ŀ.�Ƿ�סԺ����%Type,
  �������_In     In ���������Ŀ.�������%Type,
  Pass���_In     In ���������Ŀ.Pass���%Type,
  ������_In       In ���������Ŀ.������%Type,
  �Ƿ�����_In     In Number := Null
) Is

  n_Count Number(18);
  n_Id    Number(18);

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

Begin

  --�����ĿID�Ƿ����
  Select Count(1) Into n_Count From ���������Ŀ Where ID = ��Ŀid_In;

  If n_Count = 0 Then
    If ���_In = 4 Then
      --���4=�Զ��塱
      If Nvl(��Ŀid_In, 0) <= 0 Then
        --����ID
        Select ���������Ŀ_Id.Nextval Into n_Id From Dual;
      Else
        n_Id := ��Ŀid_In;
      End If;
      Insert Into ���������Ŀ
        (ID, ���, ����, ���, ����, �Ƿ���������, �Ƿ�סԺ����, �������, Pass���, ������, ����ʱ��, ����ʱ��)
      Values
        (n_Id, ���_In, ����_In, ���_In, ����_In, �Ƿ���������_In, �Ƿ�סԺ����_In, �������_In, Null, ������_In, Sysdate, Null);
    Else
      v_Err_Msg := 'δ�ҵ���Ŀ���ݣ�';
      Raise Err_Item;
    End If;
  Else
    If ���_In = 1 Then
      --1=��������취7��
      Update ���������Ŀ
      Set �Ƿ��������� = �Ƿ���������_In, �Ƿ�סԺ���� = �Ƿ�סԺ����_In, ������ = ������_In, ����ʱ�� = Sysdate
      Where ID = ��Ŀid_In;
    Elsif ���_In = 2 Then
      --2=������������淶28��
      Update ���������Ŀ
      Set �Ƿ��������� = �Ƿ���������_In, �Ƿ�סԺ���� = �Ƿ�סԺ����_In, ������ = ������_In, ����ʱ�� = Sysdate
      Where ID = ��Ŀid_In;
    Elsif ���_In = 3 Then
      --3=�̶�
      Update ���������Ŀ
      Set �Ƿ��������� = �Ƿ���������_In, �Ƿ�סԺ���� = �Ƿ�סԺ����_In, Pass��� = Pass���_In, ������ = ������_In, ����ʱ�� = Sysdate
      Where ID = ��Ŀid_In;
    Elsif ���_In = 4 Then
      --4=�Զ���
      If �Ƿ�����_In = 1 Then
        --��顰������������Ƿ�ʹ��
        Select Count(1) Into n_Count From ��������� Where �����Ŀid = ��Ŀid_In;
        If n_Count <= 0 Then
          Delete ���������Ŀ Where ID = ��Ŀid_In;
        Else
          Update ���������Ŀ Set ����ʱ�� = Sysdate Where ID = ��Ŀid_In;
        End If;
      Else
        Update ���������Ŀ
        Set ���� = ����_In, ��� = ���_In, ���� = ����_In, �Ƿ��������� = �Ƿ���������_In, �Ƿ�סԺ���� = �Ƿ�סԺ����_In, ������ = ������_In, ����ʱ�� = Sysdate
        Where ID = ��Ŀid_In;
      End If;
    Else
      v_Err_Msg := '��Ŀ�������ȷ��';
      Raise Err_Item;
    End If;
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���������Ŀ_Update;
/

CREATE OR REPLACE Procedure Zl_�������_Insert
(
  ��id_In     In ��������¼.Id%Type,
  ����id_In     In ��������¼.����id%Type,
  �Һ�id_In     In ��������¼.�Һ�id%Type,
  ��ҳid_In     In ��������¼.��ҳid%Type,
  �ύ����id_In In ��������¼.�ύ����id%Type,
  �ύ��_In     In ��������¼.�ύ��%Type,
  ��ҩҩ��id_In In ��������¼.��ҩҩ��id%Type,
  ҽ��id_In     In Varchar2
) Is

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

Begin

  --���������¼
  If Nvl(�Һ�id_In, 0) > 0 Then
    Insert Into ��������¼
      (ID, ����id, �Һ�id, �ύ��, �ύʱ��, �ύ����id, ��ҩҩ��id, ״̬)
    Values
      (��id_In, ����id_In, �Һ�id_In, �ύ��_In, Sysdate, �ύ����id_In, ��ҩҩ��id_In, 0);
  Else
    Insert Into ��������¼
      (ID, ����id, ��ҳid, �ύ��, �ύʱ��, �ύ����id, ��ҩҩ��id, ״̬)
    Values
      (��id_In, ����id_In, ��ҳid_In, �ύ��_In, Sysdate, �ύ����id_In, ��ҩҩ��id_In, 0);
  End If;

  --���������¼��Ӧ��ҽ��
  For r_Medical In (Select /*+RULE*/
                     ID
                    From ����ҽ����¼ A, Table(f_Num2list(ҽ��id_In, ',')) B
                    Where a.Id = b.Column_Value) Loop
    --���޸ľ�ҽ��id������ύ
    If r_Medical.Id Is Not Null Then
      Update ���������ϸ Set ����ύ = Null Where ҽ��id = r_Medical.Id;
    
      Insert Into ���������ϸ (��id, ҽ��id, ����ύ) Values (��id_In, r_Medical.Id, 1);
    End If;
  
  End Loop;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�������_Insert;
/

CREATE OR REPLACE Procedure Zl_�������_Auto
(
  ��id_In         In ���������.��id%Type,
  �Զ����_In       In ���������.�Զ����%Type,
  �����Ŀ��ҽ��_In In Varchar2
) Is

  v_Err_Msg Varchar2(255);
  Err_Item Exception;
Begin

  For r_Info In (Select /*+RULE*/
                  C1 �����Ŀid, C2 ҽ��id
                 From Table(f_Num2list2(�����Ŀ��ҽ��_In, '|', ','))) Loop
    If r_Info.ҽ��id Is Not Null Then
      --���޸ľ�ҽ��id������ύ
      Update ��������� Set ����ύ = Null Where ҽ��id = r_Info.ҽ��id;
    
      Insert Into ���������
        (��id, ҽ��id, �����Ŀid, ����ύ, �Զ����)
      Values
        (��id_In, Decode(r_Info.ҽ��id, 0, Null, r_Info.ҽ��id), r_Info.�����Ŀid, 1, �Զ����_In);
    End If;
  
  End Loop;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�������_Auto;
/

Create Or Replace Procedure Zl_�������_Cancel
(
  ҽ��id_In  In Varchar2,
  ����id_Out Out Varchar2
) Is

  n_Count   Number;
  v_Lockid  Varchar2(4000);
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

Begin

  --ȡҽ����Ӧ����ID
  For r_Info In (Select /* +RULE*/
                 Distinct b.Id, b.״̬
                 From ���������ϸ A, ��������¼ B, ����ҽ����¼ C, Table(f_Num2list(ҽ��id_In, ',')) D
                 Where a.��id = b.Id And a.ҽ��id = c.Id And c.���id = d.Column_Value And a.����ύ = 1 And
                       (b.״̬ Between 0 And 1 Or b.״̬ Is Null) And c.������� In ('5', '6', '7')) Loop
  
    Select Count(1) Into n_Count From ��������¼ Where ID = r_Info.Id And �����û� Is Not Null;
    If n_Count = 0 Then
      --δ����
      If Nvl(r_Info.״̬, 0) = 0 Then
        --δ��飬ֱ��ɾ����¼
        Delete ��������¼ Where ID = r_Info.Id And (״̬ = 0 Or ״̬ Is Null);
      Elsif r_Info.״̬ = 1 Then
        --����飬����״̬
        Update ��������¼ Set ״̬ = ״̬ + 10 Where ״̬ = 1;
      End If;
    Else
      --������
      Begin
        Select f_List2str(Cast(Collect(Cast(ҽ��id As Varchar2(20))) As t_Strlist), ',')
        Into v_Lockid
        From ���������ϸ
        Where ��id = r_Info.Id
        Order By ҽ��id;
      Exception
        When Others Then
          v_Lockid := Null;
      End;
    
      If v_Lockid Is Not Null Then
        ����id_Out := ����id_Out || ',' || v_Lockid;
      End If;
    End If;
  
  End Loop;

  If Substr(����id_Out, 1, 1) = ',' Then
    ����id_Out := Substr(����id_Out, 2, Length(����id_Out));
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�������_Cancel;
/

CREATE OR REPLACE Procedure Zl_�������_Update
(
  ҵ�����_In In Number,
  ����id_In   In ���ű�.Id%Type,
  ��id_In   In ��������¼.Id%Type
) Is
  --���ܣ���ҵ����𣬸��´�������¼��״̬
  --������
  --  ҵ�����_In��1-����ҵ��2-סԺҵ��
  --  ����id_In���ٴ�����ID
  --  ����ID_In����

  n_Param1  Number(10);
  n_Param2  Number(10);
  n_Count   Number(18);
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

Begin

  Select Nvl(zl_GetSysParameter('������ʱ��'), '0') Into n_Param1 From Dual;
  Select Nvl(zl_GetSysParameter('����ҩʦ���ʱ��'), '10') Into n_Param2 From Dual;

  If ҵ�����_In = 1 Then
    --����ҵ��
    Select Count(1)
    Into n_Count
    From (Select Max(������ʱ��) ������ʱ��
           From ����������
           Where Nvl(�������, 0) = 0 And ',' || ��Դ���� || ',' Like '%,' || ����id_In || ',%')
    Where ������ʱ�� <= (Sysdate - n_Param2 / 24 / 60);
  
    If n_Count > 0 Then
      --��ʱ��δ��飬���2-��ʱ����
      Update ��������¼
      Set ״̬ = 2, �����û� = Null, ����ʱ�� = Null
      Where (״̬ = 0 Or ״̬ Is Null) And ID = ��id_In;
    End If;
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�������_Update;
/

CREATE OR REPLACE Procedure Zl_����������_Save
(
  ���_In         In Number,
  ������_In       In ����������.������%Type,
  �������_In     In ����������.�������%Type,
  �Ƿ�����_In In ����������.�Ƿ�����%Type := Null,
  ��Դ����_In     In ����������.��Դ����%Type := Null
) Is

  --���ܣ����洦��������
  --������
  --  ���_In��1-������Դ���ң�2-����������ʱ��
  --  �������_In��0-���1-סԺ
  --  �Ƿ�����_In�����_In = 2���ò�������
  --  ��Դ����_In�����_In = 1���ò�������

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

Begin

  If ���_In = 1 Then
    --������Դ����
    Update ���������� Set ��Դ���� = ��Դ����_In Where ������ = ������_In And ������� = �������_In;
  
    If Sql%RowCount = 0 Then
      Insert Into ����������
        (������, �������, �Ƿ�����, ������ʱ��, ��Դ����)
      Values
        (������_In, �������_In, 0, Sysdate, ��Դ����_In);
    End If;
  Elsif ���_In = 2 Then
    --�����Ƿ����󷽡�������ʱ��
    Update ����������
    Set �Ƿ����� = �Ƿ�����_In, ������ʱ�� = Sysdate
    Where ������ = ������_In And ������� = �������_In;
  
    If Sql%RowCount = 0 Then
      Insert Into ����������
        (������, �������, �Ƿ�����, ������ʱ��, ��Դ����)
      Values
        (������_In, �������_In, �Ƿ�����_In, Sysdate, Null);
    End If;
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����������_Save;
/

CREATE OR REPLACE Procedure Zl_������鳣������_Update
(
  ���ܺ�_In In Number,
  �û���_In In ������鳣������.�û���%Type,
  ����_In   In ������鳣������.����%Type
) Is

  --���ܣ�������ɾ��������鳣������
  --������
  --  ���ܺ�_In��1-������0-ɾ��

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

Begin

  If ���ܺ�_In = 0 Then
    Delete ������鳣������ Where �û��� = �û���_In And ���� = ����_In;
  Elsif ���ܺ�_In = 1 Then
    Insert Into ������鳣������ (�û���, ����) Values (�û���_In, ����_In);
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_������鳣������_Update;
/

Create Or Replace Procedure Zl_��������¼_Lock
(
  Lock_In     In Number,
  ������_In   In ����������.������%Type,
  �������_In In ����������.�������%Type,
  ��id_In   In ��������¼.Id%Type := Null
) Is
  --���ܣ���������¼�ļ����������л�
  --������
  --  Lock_In��0-������1-����
  --  �������_In��0-���1-סԺ

  n_Count   Number(2);
  v_Err_Msg Varchar2(255);
  Err_Item Exception;
Begin

  If Lock_In = 1 Then
  
    --�ȼ��
    Select Count(1) Into n_Count From ��������¼ Where ID = ��id_In;
    If n_Count <= 0 Then
      v_Err_Msg := '�ô�������¼�ѱ�ɾ����';
      Raise Err_Item;
    End If;
    
    Select Count(1) Into n_Count From ��������¼ Where ID = ��id_In And ״̬ = 0;
    If n_Count <= 0 Then
      v_Err_Msg := '�ô�������¼�ѱ���飡';
      Raise Err_Item;
    End If;

    --����
    If Nvl(��id_In, 0) > 0 Then
      Update ���������� Set ������ʱ�� = Sysdate Where ������ = ������_In And ������� = �������_In;
    
      Update ��������¼
      Set �����û� = Upper(User), ����ʱ�� = Sysdate
      Where ID = ��id_In And (�����û� Is Null Or �����û� = Upper(User));
      If Sql%NotFound Then
        v_Err_Msg := '�ô�������¼�ѱ�������������';
        Raise Err_Item;
      End If;
    End If;
  
  Else
  
    --����
    Update ���������� Set ������ʱ�� = Sysdate Where ������ = ������_In And ������� = �������_In;
  
    If Nvl(��id_In, 0) = 0 Then
      --���е�ǰ�û���������¼
      Update ��������¼ Set �����û� = Null, ����ʱ�� = Null Where �����û� = Upper(User);
    Else
      Update ��������¼ Set �����û� = Null, ����ʱ�� = Null Where ID = ��id_In And �����û� = Upper(User);
      If Sql%NotFound Then
        v_Err_Msg := '�ô�������¼�ѱ������˽�����';
        Raise Err_Item;
      Else
        --������һСʱ�����ļ�¼
        Update ��������¼
        Set �����û� = Null, ����ʱ�� = Null
        Where ����ʱ�� < Sysdate - 1 / 24 And �����û� = Upper(User);
      End If;
    End If;
  
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��������¼_Lock;
/

Create Or Replace Procedure Zl_�������_Audit
(
  ��id_In   In ��������¼.Id%Type,
  �����_In In ��������¼.�����%Type,
  �����_In   In ��������¼.�����%Type,
  �ۺ�����_In In ��������¼.�ۺ�����%Type
) Is
  --���ܣ��ύ��������¼��Ϣ

  v_Err_Msg Varchar2(255);
  Err_Item Exception;
Begin

  Update ��������¼
  Set ����� = �����_In, ����� = �����_In, ���ʱ�� = Sysdate, �ۺ����� = �ۺ�����_In, ״̬ = 1, �����û� = Null, ����ʱ�� = Null
  Where ID = ��id_In And ���ʱ�� Is Null;
  If Sql%NotFound Then
    v_Err_Msg := '�ò��˵ļ�¼�ѱ���������飡';
    Raise Err_Item;
  End If; 

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�������_Audit;
/

Create Or Replace Procedure Zl_�������_Audit_Detail
(
  ��id_In     In ���������.��id%Type,
  ҽ��id_In     In ���������.ҽ��id%Type,
  �����Ŀid_In In ���������.�����Ŀid%Type,
  ҩʦ���_In   In ���������.ҩʦ���%Type
) Is
  --���ܣ��ύ�����������Ϣ(ҩʦ�����Ϣ)

  v_Err_Msg Varchar2(255);
  Err_Item Exception;
Begin

  --���¾�ҽ��ID�ġ�����ύ��
  Update ��������� Set ����ύ = Null Where ����ύ Is Not Null And ҽ��id = ҽ��id_In;

  --����ҩʦ�����Ϣ
  Update ���������
  Set ����ύ = Decode(Nvl(ҽ��id_In, 0), 0, Null, 1), ҩʦ��� = ҩʦ���_In
  Where ҽ��id = ҽ��id_In And �����Ŀid = �����Ŀid_In And ��id = ��id_In;
  If Sql%NotFound Then
    Insert Into ���������
      (��id, ҽ��id, �����Ŀid, ����ύ, ҩʦ���, �Զ����)
    Values
      (��id_In, ҽ��id_In, �����Ŀid_In, Decode(Nvl(ҽ��id_In, 0), 0, Null, 1), ҩʦ���_In, Null);
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�������_Audit_Detail;
/

CREATE OR REPLACE Function Zl_Fun_Pati_Calorie
(
  ����id_In In ������Ϣ.����id%Type,
  ��ҳid_In In ������Ϣ.��ҳid%Type,
  �Һ�id_In In ���˹Һż�¼.Id%Type
) Return Varchar2 Is

  --���ܣ�ͨ��������Ϣ����������˵�������Ҫ��
  v_Return  Varchar2(500);
  n_Sex     Number(1);
  n_Age     Number(5);
  n_Age_Var Number(10, 2);
  n_High    Number(5);
  n_Weight  Number(5);
  n_Calorie Number(10);
  n_Err     Number(1) := 1;
  v_Tmp     Varchar2(500);

  --��ȡ�����ַ�������ֵ
  Function Get_Age(����_In In Varchar2) Return Number Is
    v_Tmp Varchar2(100) := '';
    N     Number(3) := 1;
  Begin
    Loop
      If N > Length(����_In) Then
        Exit;
      End If;
      If Regexp_Like(Substr(����_In, N, 1), '[0-9]') Then
        v_Tmp := v_Tmp || Substr(����_In, N, 1);
      Else
        Exit;
      End If;
      N := N + 1;
    End Loop;
  
    Return v_Tmp;
  End;

Begin

  If ��ҳid_In Is Null And �Һ�id_In Is Null Then
    Return Null;
  End If;

  --�Ա�
  Begin
    Select Decode(�Ա�, '��', 1, 'Ů', 2, Null), ���� Into n_Sex, v_Tmp From ������Ϣ Where ����id = ����id_In;
  Exception
    When Others Then
      Select Null, Null Into n_Sex, v_Tmp From Dual;
  End;

  --����
  If v_Tmp Is Null Then
    Select 0, 0, 0 Into n_Age, n_Age_Var, n_Err From Dual;
  Else
    If v_Tmp Like '%��%' Then
      n_Age     := Get_Age(v_Tmp);
      n_Age_Var := 1;
    Elsif v_Tmp Like '%��%' Then
      n_Age     := Get_Age(v_Tmp);
      n_Age_Var := Round(1 / 12, 2);
    Elsif v_Tmp Like '%��%' Or v_Tmp Like '%��%' Then
      n_Age     := Get_Age(v_Tmp);
      n_Age_Var := Round(1 / 365, 2);
    Elsif v_Tmp Like '%Сʱ%' Or v_Tmp Like '%��%' Then
      n_Age     := 1;
      n_Age_Var := Round(1 / 365, 2);
    Else
      Select 0, 0, 0 Into n_Age, n_Age_Var, n_Err From Dual;
    End If;
  End If;

  If ��ҳid_In Is Not Null Then
    --סԺ
  
    --���
    Begin
      Select ���, ���� Into n_High, n_Weight From ������ҳ Where ����id = ����id_In And ��ҳid = ��ҳid_In;
      If Nvl(n_High, 0) = 0 Or Nvl(n_Weight, 0) = 0 Then
        n_Err := 0;
      End If;
    Exception
      When Others Then
        Select 0, 0, 0 Into n_High, n_Weight, n_Err From Dual;
    End;
  
  Else
    --����
  
    --���
    Begin
      Select b.��¼����
      Into n_High
      From ���˻����¼ A, ���˻������� B
      Where a.Id = b.��¼id And a.����id = ����id_In And a.��ҳid = �Һ�id_In And a.������Դ = 1 And b.��Ŀ���� = '���';
    Exception
      When Others Then
        Select 0, 0 Into n_High, n_Err From Dual;
    End;
  
    --����
    Begin
      Select b.��¼����
      Into n_Weight
      From ���˻����¼ A, ���˻������� B
      Where a.Id = b.��¼id And a.����id = ����id_In And a.��ҳid = �Һ�id_In And a.������Դ = 1 And b.��Ŀ���� = '����';
    Exception
      When Others Then
        Select 0, 0 Into n_Weight, n_Err From Dual;
    End;
  
  End If;

  --������Ҫ��
  Select Nvl(n_High, 0), Nvl(n_Weight, 0) Into n_High, n_Weight From Dual;
  If n_Sex = 1 Then
    n_Calorie := 66.5 + 13.8 * n_Weight + 5.0 * n_High - 6.8 * n_Age * n_Age_Var;
    v_Return  := '66.5 + 13.8 * ' || n_Weight || 'KG + 5.0 * ' || n_High || 'CM - 6.8 * ' ||
                 Round(n_Age * n_Age_Var, 2) || '�� = ' || n_Calorie * n_Err;
  Else
    n_Calorie := 655.1 + 9.6 * n_Weight + 1.8 * n_High - 4.7 * n_Age * n_Age_Var;
    v_Return  := '655.1 + 9.6 * ' || n_Weight || 'KG + 1.8 * ' || n_High || 'CM - 4.7 * ' ||
                 Round(n_Age * n_Age_Var, 2) || '�� = ' || n_Calorie * n_Err;
  End If;

  Return v_Return;

End Zl_Fun_Pati_Calorie;
/

Create Or Replace Procedure Zl1_Datamove_Reb
(
  System_In    In Number,
  Speedmode_In In Number,
  Func_In      In Number,
  Enable_In    In Number := 0,
  Parallel_In  In Number := 0,
  Rebscope_In  In Number := 0
) As
  --���ܣ�����ʷ����ת��֮ǰ�����ô��������Զ���ҵ��Լ����������ת��֮��������Щ�����Լ��ؽ���ת���������ջر��ת�����������Ŀռ�
  --������
  --System_In:    Ӧ��ϵͳ���,100=��׼��
  --speedmode_in������ת��ģʽ��0-����ģʽ��1-����ģʽ���ڿͻ���ͣ��ʱ��ת���ڼ����ת�����������Ψһ�������Լ�����������Լӿ���ת���ݵ�ɾ��������
  --func_in:      1=��������2=�Զ���ҵ��3=Լ����4=������5=�ؽ���ת��������6-�ջر��ת�����������Ŀռ䣬7-�����Ĵ洢�ռ䣨move�������ָ������õ�Լ��������
  --Enable_in:    0-���ã�1=���ã���func_inֵΪ1-4��Ч
  --rebScope_in:   Func_In=6ʱ��ָ�ؽ������ķ�Χ(0-���ú�����,1-���ú����༰ҽ����,2-ȫ��)��Func_In=7ʱָMove��ķ�Χ(0-���ú����࣬1-ȫ��)

  v_Sql Varchar2(1000);
  n_Do  Number(1);
  v_Tbs Varchar2(100);

  --ת������е�SQL��ѯ���������
  v_Indexeswithtag Varchar2(4000) := '������ü�¼_IX_����ID,סԺ���ü�¼_IX_����ID,���ò����¼_IX_����ID,���ò����¼_IX_�Ǽ�ʱ��,����Ԥ����¼_IX_��ҳID,����Ԥ����¼_IX_����ID,����Ԥ����¼_IX_�տ�ʱ��,������ü�¼_IX_�Ǽ�ʱ��,������ü�¼_IX_ҽ�����,סԺ���ü�¼_IX_�Ǽ�ʱ��,���˽��ʼ�¼_IX_�շ�ʱ��,���˽��ʼ�¼_IX_����id' ||
                                     ',ҩƷ�շ���¼_IX_����ID,�շ���¼������Ϣ_IX_�շ�ID,��Һ��ҩ����_IX_�շ�ID,ҩƷ����ƻ�_IX_����ID,ҩƷǩ����ϸ_IX_�շ�ID' ||
                                     ',��Ա����¼_IX_���ʱ��,��Ա�սɼ�¼_IX_�Ǽ�ʱ��,��Ա�ݴ��¼_IX_�ս�ID,��Ա�ݴ��¼_IX_�Ǽ�ʱ��,Ʊ�����ü�¼_IX_�Ǽ�ʱ��,Ʊ��ʹ����ϸ_IX_����ID,Ʊ�ݴ�ӡ��ϸ_IX_ʹ��ID' ||
                                     ',���˹Һż�¼_IX_�Ǽ�ʱ��,����ҽ������_IX_����ʱ��,����ҽ����¼_IX_�Һŵ�,����ҽ����¼_IX_��ҳID,����ҽ����¼_IX_���ID' ||
                                     ',������ҳ_IX_��Ժ����,סԺ���ü�¼_IX_����ID,���˹�����¼_IX_����ID,������ϼ�¼_IX_����ID,���������¼_IX_��ҳID' ||
                                     ',���˻����¼_IX_��ҳID,���˻�������_IX_��¼id,���˻����ļ�_IX_��ҳID,���˻�������_IX_�ļ�ID,���˻�����ϸ_IX_��¼ID,���˻����ӡ_IX_�ļ�ID' ||
                                     ',���Ӳ�����¼_IX_����ID,����ҽ������_IX_����ID,Ӱ�񱨸沵��_IX_ҽ��ID,������ļ�¼_IX_����ID,������ϼ�¼_IX_����ID' ||
                                     ',�����ٴ�·��_IX_����ID,���˺ϲ�·��_IX_��Ҫ·����¼ID,����·��ִ��_IX_·����¼ID,���˳�����¼_IX_·����¼ID,�������ҽ��_IX_ҽ��ID' ||
                                     ',Ӱ�����뵥ͼ��_IX_ҽ��ID,Ӱ���ղ�����_IX_ҽ��ID,����걾��¼_IX_ҽ��ID,������Ŀ�ֲ�_IX_�걾ID,���������¼_IX_�걾ID' ||
                                     ',���������¼_IX_�걾ID,����ͼ����_IX_�걾ID,������ռ�¼_IX_ҽ��ID,������ͨ���_IX_����걾ID,���������ϸ_IX_ҽ��ID';

  --ת������е�SQL��ѯ���������(������Ψһ����Ӧ������)
  v_Constraintswithtag Varchar2(4000) := '����Ԥ����¼_UQ_NO,���˽��ʼ�¼_UQ_NO,���˽��ʼ�¼_PK,������ü�¼_UQ_NO,סԺ���ü�¼_UQ_NO' ||
                                         ',���˿��������_PK,���ò����¼_PK,���˿������¼_PK,�������㽻��_PK,��Һ��ҩ��¼_PK,ҩƷǩ����¼_PK,Ʊ�ݴ�ӡ����_PK,���˹Һż�¼_PK,���˹ҺŻ���_UQ_����,����ת���¼_UQ_NO' ||
                                         ',���˻�����Ŀ_UQ_ҳ��,���˻���Ҫ������_UQ_ҳ��,����Ҫ������_PK,���Ӳ�����¼_PK,���Ӳ�������_PK,���Ӳ�����ʽ_PK,���Ӳ�������_UQ_�������,���Ӳ���ͼ��_PK,�����걨��¼_PK' ||
                                         ',���˺ϲ�·������_PK,����·������_PK,����·������_PK,����·��ָ��_UQ_����ָ��,����·��ҽ��_PK' ||
                                         ',����ҽ����¼_PK,����ҽ������_PK,����ҽ���Ƽ�_UQ_�շ�ϸĿID,����ҽ������_PK,����ҽ������_PK,����ҽ��ִ��_PK,ҽ��ִ��ʱ��_PK,ҽ��ִ�д�ӡ_PK,����ҽ����ӡ_UQ_ҽ��ID,��Ѫ�����¼_PK,��Ѫ������_PK' ||
                                         ',������ϼ�¼_PK,����ҽ��״̬_PK,ҽ��ǩ����¼_PK,����ҽ������_PK,���Ƶ��ݴ�ӡ_PK,ҽ��ִ�мƼ�_PK,ִ�д�ӡ��¼_PK' ||
                                         ',Ӱ�����¼_PK,Ӱ��������_UQ_���к�,Ӱ����ͼ��_UQ_ͼ���,Ӱ��Σ��ֵ��¼_UQ_ҽ��ID' ||
                                         ',����������Ŀ_PK,�����ʿؼ�¼_PK,����ǩ����¼_PK,�����Լ���¼_PK,�����ʿر���_PK,����ҩ�����_PK,��Ա�սɼ�¼_PK,��Ա�ս���ϸ_PK,��Ա�ս�Ʊ��_PK,��Ա�սɶ���_PK' ||
                                         ',��������¼_PK,���������_UQ_��ID';

  --���ܣ�1.���û���������ת�����������������,����ɾ�������¼ʱ���ӱ�ÿ�м�¼ִ��һ��SQL��ѯ��ɾ��
  --      2.���û�����������Ψһ��Լ��������ʱ���Զ�ɾ����Ӧ������������ʱ�Զ������������������ɾ������
  --���磺����ҽ������_FK_ҽ��ID�������Щ������ڵı�����δת����δ��zlbaktables���ж��壩��ִ��ǰ���鲢����ת����
  Procedure Setconstraintstatus As
  Begin
    --����ʱ���Ƚ�������ת��������������������ٽ���ת���������
    If Enable_In = 0 Then
      For R In (Select c.Table_Name, c.Constraint_Name, a.Table_Name As Ptable_Name
                From User_Constraints A, User_Constraints C, zlBakTables B
                Where a.Table_Name = b.���� And b.ֱ��ת�� = 1 And b.ϵͳ = System_In And a.Constraint_Type In ('P', 'U') And
                      c.r_Constraint_Name = a.Constraint_Name And c.Constraint_Type = 'R' And c.Status = 'ENABLED'
                Order By a.Table_Name) Loop
        v_Sql := 'Alter Table ' || r.Table_Name || ' Disable Constraint ' || r.Constraint_Name;
        Execute Immediate v_Sql;
      End Loop;
    
      If Speedmode_In = 1 Then
        --����������Ψһ������(����ɾ������������ʹskip_unusable_indexesΪtrue��Ҳ�޷�ɾ������Unusable״̬��Ψһ�������ı��еļ�¼)
        --����ת������е�SQL��ѯ���������(������Ψһ����Ӧ������)
        For R In (Select a.Table_Name, a.Constraint_Name
                  From User_Constraints A, zlBakTables T, User_Tables B
                  Where a.Table_Name = t.���� And t.ֱ��ת�� = 1 And t.ϵͳ = System_In And a.Status = 'ENABLED' And
                        a.Constraint_Type In ('P', 'U') And a.Table_Name = b.Table_Name And b.Iot_Type Is Null And
                        a.Constraint_Name Not In
                        (Select Upper(Column_Value) As Constraint_Name From Table(f_Str2list(v_Constraintswithtag)))
                  Order By Constraint_Name) Loop
          v_Sql := 'Alter Table ' || r.Table_Name || ' Disable Constraint ' || r.Constraint_Name ||
                   ' Cascade Drop Index';
          Execute Immediate v_Sql;
        End Loop;
      End If;
    Else
      --����ʱ��������������Ψһ��������������ת�����������������
      If Speedmode_In = 1 Then
        --���ؽ�������������Լ�����Ա��ؽ�����ʱ���ò���ִ������ʱ�䣬��������Լ��ʱҲ���Բ���novalidate��ʽ
        For R In (Select d.Table_Name, d.Constraint_Name, LTrim(Max(Sys_Connect_By_Path(d.Column_Name, ',')), ',') Colstr
                  From User_Cons_Columns D,
                       (Select a.Table_Name, a.Constraint_Name
                         From User_Constraints A, zlBakTables T
                         Where a.Table_Name = t.���� And t.ֱ��ת�� = 1 And t.ϵͳ = System_In And a.Status = 'DISABLED' And
                               a.Constraint_Type In ('P', 'U')) A
                  Where a.Constraint_Name = d.Constraint_Name And a.Table_Name = d.Table_Name
                  Start With d.Position = 1
                  Connect By Prior d.Position + 1 = d.Position And Prior d.Constraint_Name = d.Constraint_Name
                  Group By d.Table_Name, d.Constraint_Name
                  Order By Constraint_Name) Loop
          Update Zldatamovelog
          Set ��ǰ���� = '���ڻָ�Լ��:' || r.Constraint_Name
          Where ϵͳ = System_In And ���� = (Select Max(����) From Zldatamovelog Where ϵͳ = System_In);
        
          Select Tablespace_Name Into v_Tbs From User_Indexes Where Table_Name = r.Table_Name And Rownum < 2;
        
          --����������Ψһ��ʱ�������Ǳ�ɾ���˵ģ���������Ҫ��Create
          v_Sql := 'Create Unique Index ' || r.Constraint_Name || ' On ' || r.Table_Name || '(' || r.Colstr ||
                   ') Tablespace ' || v_Tbs || ' Nologging';
          Begin
            Execute Immediate v_Sql;
          Exception
            When Others Then
              Null; --������Щ������Ψһ�����Ǳ���ת���ڼ䱻���õģ�֮ǰ�ʹ��ڲ�Ψһ���ݣ�����Ψһ���������
          End;
        
          --���Զ�����Լ���������Ĺ���
          v_Sql := 'Alter Table ' || r.Table_Name || ' Enable Novalidate Constraint ' || r.Constraint_Name;
          Execute Immediate v_Sql;
        End Loop;
      End If;
    
      --��������ת�����������������
      For R In (Select c.Table_Name, c.Constraint_Name, a.Table_Name As Ptable_Name
                From User_Constraints A, User_Constraints C, zlBakTables B
                Where a.Table_Name = b.���� And b.ֱ��ת�� = 1 And b.ϵͳ = System_In And a.Constraint_Type In ('P', 'U') And
                      c.r_Constraint_Name = a.Constraint_Name And c.Constraint_Type = 'R' And c.Status = 'DISABLED'
                Order By a.Table_Name) Loop
        --Ϊ�˼ӿ��ٶȣ�����novalidate������֤��������
        --��������ת����������������zlbaktables�ж����ˣ���û�б�д��Ӧ������ת���ű���δ��֤�����ݿ�����Υ��Լ���������
        v_Sql := 'Alter Table ' || r.Table_Name || ' Enable Novalidate Constraint ' || r.Constraint_Name;
        Execute Immediate v_Sql;
      End Loop;
    End If;
  End Setconstraintstatus;

  --���ܣ�����ģʽʱ����LOB�������������������ģʽʱ������ת�������÷�ת������������(���磺����ҽ���Ƽ�_IX_�շ�ϸĿID)
  --˵��������������Ϊ�����ɾ�����ݵ�����
  Procedure Setindexstatus As
  Begin
    If Speedmode_In = 1 Then
      --����ת������е�SQL��ѯ���������
      For R In (Select /*+ rule*/
                 a.Index_Name
                From User_Indexes A, zlBakTables T
                Where a.Table_Name = t.���� And t.ֱ��ת�� = 1 And t.ϵͳ = System_In And t.ֱ��ת�� = 1 And
                      a.Index_Name <> a.Table_Name || '_IX_��ת��' And
                      a.Index_Name Not In
                      (Select Upper(Column_Value) As Index_Name From Table(f_Str2list(v_Indexeswithtag))) And
                      a.Status = Decode(Enable_In, 0, 'VALID', 'UNUSABLE') And a.Index_Type = 'NORMAL' And Not Exists
                 (Select 1
                       From User_Constraints C
                       Where c.Index_Name = a.Index_Name And c.Constraint_Type In ('P', 'U'))
                Order By Index_Name) Loop
      
        If Enable_In = 0 Then
          v_Sql := 'Alter Index ' || r.Index_Name || ' Unusable';
          Execute Immediate v_Sql;
        Else
          Update Zldatamovelog
          Set ��ǰ���� = '�����ؽ�����:' || r.Index_Name
          Where ϵͳ = System_In And ���� = (Select Max(����) From Zldatamovelog Where ϵͳ = System_In);
        
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Nologging';
          Begin
            Execute Immediate v_Sql;
            --�����ؽ��Ƚ������������ؽ�����Ҫ���������������������������ORA-00054: ��Դ��æ, ��ָ���� NOWAIT ��ʽ��ȡ��Դ


          Exception
            When Others Then
              If SQLErrM Like 'ORA-00054%' Then
                v_Sql := Replace(v_Sql, 'Rebuild', 'Rebuild Online');
                Execute Immediate v_Sql;
              End If;
          End;
        End If;
      End Loop;
    Else
      For R In (Select a.Index_Name
                From (Select d.Table_Name, d.Index_Name, LTrim(Max(Sys_Connect_By_Path(d.Column_Name, ',')), ',') Colstr
                       From User_Ind_Columns D, zlBakTables T, User_Indexes C
                       Where c.Table_Name = t.���� And t.ֱ��ת�� = 1 And t.ϵͳ = System_In And c.Uniqueness = 'NONUNIQUE' And
                             c.Index_Type = 'NORMAL' And c.Status = Decode(Enable_In, 0, 'VALID', 'UNUSABLE') And
                             c.Index_Name = d.Index_Name And c.Table_Name = d.Table_Name
                       Start With d.Column_Position = 1
                       Connect By Prior d.Column_Position + 1 = d.Column_Position And Prior d.Index_Name = d.Index_Name
                       Group By d.Table_Name, d.Index_Name) A,
                     (Select e.Table_Name, LTrim(Max(Sys_Connect_By_Path(e.Column_Name, ',')), ',') Colstr
                       From User_Cons_Columns E, User_Constraints F, zlBakTables T, User_Constraints C
                       Where e.Table_Name = t.���� And t.ֱ��ת�� = 1 And t.ϵͳ = System_In And
                             e.Constraint_Name = f.Constraint_Name And f.Constraint_Type = 'R' And
                             c.Constraint_Name = f.r_Constraint_Name And c.Table_Name Not In ('������ҳ', '������Ϣ') And
                             Not Exists
                        (Select 1 From zlBakTables G Where g.���� = c.Table_Name And g.ϵͳ = System_In)
                       Start With Nvl(e.Position, 1) = 1
                       Connect By Prior Nvl(e.Position, 1) + 1 = Nvl(e.Position, 1) And
                                  Prior e.Constraint_Name = e.Constraint_Name
                       Group By e.Table_Name, e.Constraint_Name) B
                Where a.Table_Name = b.Table_Name And a.Colstr = b.Colstr
                Order By Index_Name) Loop
      
        If Enable_In = 0 Then
          --���⴦�������������������ã�������ҩƷĿ¼�޸Ĺ�񣬲���ɿ���Ҫʹ��
          If r.Index_Name Not In ('����ҽ����¼_IX_�շ�ϸĿID', '��Ա�սɼ�¼_IX_�ɿ���ID') Then
            v_Sql := 'Alter Index ' || r.Index_Name || ' Unusable';
            Execute Immediate v_Sql;
          End If;
        Else
          Update Zldatamovelog
          Set ��ǰ���� = '�����ؽ�����:' || r.Index_Name
          Where ϵͳ = System_In And ���� = (Select Max(����) From Zldatamovelog Where ϵͳ = System_In);
        
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Online Nologging';
          Execute Immediate v_Sql;
          --�����ؽ��Ƚ������������ؽ�����Ҫ���������������������������ORA-00054: ��Դ��æ, ��ָ���� NOWAIT ��ʽ��ȡ��Դ  
        End If;
      End Loop;
    End If;
  End Setindexstatus;

  --���ܣ�ת�������ڼ䣬ͣ��ת�����ϵ����д�������ת�����ٻָ�
  Procedure Settriggerstatus As
  Begin
    For R In (Select Distinct a.Table_Name, t.ͣ�ô�����
              From User_Triggers A, zlBakTables T
              Where a.Status = Decode(Enable_In, 0, 'ENABLED', 'DISABLED') And a.Table_Name = t.���� And t.ֱ��ת�� = 1 And
                    t.ϵͳ = System_In) Loop
      If Enable_In = 0 Then
        v_Sql := 'Alter Table ' || r.Table_Name || ' DISABLE ALL TRIGGERS';
        Update zlBakTables Set ͣ�ô����� = 1 Where ϵͳ = System_In And ���� = r.Table_Name;
      Elsif Nvl(r.ͣ�ô�����, 0) = 1 Then
        v_Sql := 'Alter Table ' || r.Table_Name || ' ENABLE ALL TRIGGERS';
        Update zlBakTables Set ͣ�ô����� = Null Where ϵͳ = System_In And ���� = r.Table_Name;
      End If;
      Execute Immediate v_Sql;
    End Loop;
    Commit;
  End Settriggerstatus;

  --���ܣ�ת�������ڼ䣬ͣ�õ�ǰ�����ߵ������Զ���ҵ��ת����������
  Procedure Setjobstatus As
    v_Jobs Varchar2(4000);
  Begin
    --ͣ��
    If Enable_In = 0 Then
      For R In (Select Job From User_Jobs Where Broken = 'N') Loop
        Dbms_Job.Broken(r.Job, True);
        v_Jobs := v_Jobs || ',' || r.Job;
      End Loop;
    
      If v_Jobs Is Not Null Then
        v_Jobs := Substr(v_Jobs, 2);
        Update zlDataMove Set ͣ����ҵ�� = v_Jobs Where ϵͳ = System_In And ��� = 1;
      End If;
    Else
      --����
      Select ͣ����ҵ�� Into v_Jobs From zlDataMove Where ϵͳ = System_In And ��� = 1;
      If v_Jobs Is Not Null Then
        For R In (Select Job
                  From User_Jobs
                  Where Broken = 'Y' And Job In (Select Column_Value From Table(f_Num2list(v_Jobs)))) Loop
          Dbms_Job.Broken(r.Job, False);
        End Loop;
        Update zlDataMove Set ͣ����ҵ�� = Null Where ϵͳ = System_In And ��� = 1;
      End If;
    End If;
    --��ҵ���ú�����ύ�������Ч
    Commit;
  End Setjobstatus;
Begin
  If Parallel_In < 2 Then
    Execute Immediate 'Alter Session DISABLE PARALLEL DDL';
  Else
    If Speedmode_In = 1 And (Func_In In (6, 7) Or Func_In In (3, 4) And Enable_In = 1) Then
      --Ϊ�ؽ��������ò���ִ�У�����ͨ��������IO�豸�����ܣ�����̫�ߵĲ��жȷ����ή�����ܣ����и����ܴ洢�豸���ɼӴ��жȣ�
      --ִ���ؽ���������Զ�Ϊ�������ϲ��ж����ԣ������ȡ������Ӱ�����SQL��ִ�мƻ�(ȫ��ɨ��+���в�ѯ������),�ں���ȡ�������Ĳ��ж�


      Execute Immediate 'Alter Session FORCE PARALLEL DDL PARALLEL ' || Parallel_In;
    End If;
  End If;

  If Func_In = 1 Then
    --1.���ô�����
    Settriggerstatus;
  Elsif Func_In = 2 Then
    --2.�����Զ���ҵ
    Setjobstatus;
  Elsif Func_In = 3 Then
    --3.����Լ��״̬
    Setconstraintstatus;
  Elsif Func_In = 4 Then
    --4.��������״̬
    Setindexstatus;
  Elsif Func_In = 5 Then
    --5.�ؽ�"��ת��"����
    For R In (Select b.Index_Name
              From zlBakTables A, User_Indexes B
              Where a.���� = b.Table_Name And a.ֱ��ת�� = 1 And a.ϵͳ = System_In And b.Index_Name = b.Table_Name || '_IX_��ת��'
              Union All
              Select '������ҳ_IX_��ת��' From Dual Where System_In = 100) Loop
      Update Zldatamovelog
      Set ��ǰ���� = '�����ؽ�����:' || r.Index_Name
      Where ϵͳ = System_In And ���� = (Select Max(����) From Zldatamovelog Where ϵͳ = System_In);
    
      --��ʱ̫�̣����벢��DDL
      --����ת��ʱ����ؽ����������������������������������ORA-00054: ��Դ��æ, ��ָ���� NOWAIT ��ʽ��ȡ��Դ
      If Speedmode_In = 1 Then
        v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Nologging';
      Else
        v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Online Nologging';
      End If;
      Begin
        Execute Immediate v_Sql;
      Exception
        When Others Then
          If SQLErrM Like 'ORA-00054%' Then
            v_Sql := Replace(v_Sql, 'Rebuild', 'Rebuild Online');
            Execute Immediate v_Sql;
          End If;
      End;
    End Loop;
  
  Elsif Func_In = 6 Then
    --6.�ؽ����ת����ѯ���õ������������Ա����ؽ�����������һ��Ĳ�ѯʱ�䣩
    --����ҵ������ý׶��������ؽ���Щ�������Ա���һЩ����Ҫ���ؽ���ʱ
    For R In (Select b.Index_Name, a.���
              From User_Indexes B, zlBakTables A
              Where a.ϵͳ = System_In And a.���� = b.Table_Name And
                    b.Index_Name In
                    (Select Upper(Column_Value)
                     From Table(f_Str2list(v_Indexeswithtag))
                     Union
                     Select Upper(Column_Value) From Table(f_Str2list(v_Constraintswithtag)))
              Order By Index_Name) Loop
      n_Do := 0;
      If Rebscope_In = 0 Then
        If r.��� < 5 Then
          n_Do := 1; --�����ú�����
        End If;
      Elsif Rebscope_In = 1 Then
        If r.��� < 5 Or r.��� = 8 Then
          n_Do := 1; --�����ú����ࡢҽ����
        End If;
      Else
        n_Do := 1;
      End If;
    
      If n_Do = 1 Then
        Update Zldatamovelog
        Set ��ǰ���� = '�����ؽ�����:' || r.Index_Name
        Where ϵͳ = System_In And ���� = (Select Max(����) From Zldatamovelog Where ϵͳ = System_In);
      
        --v_Sql := 'Alter Index ' || r.Index_Name || ' shrink Space';
        --ʹ��shrink��ʽ���ܲ���ִ��,��������ٶȱ�rebuild PARALLEL 8 ��6��
        If Speedmode_In = 1 Then
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Nologging';
        Else
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Online Nologging';
        End If;
        Begin
          Execute Immediate v_Sql;
          --�����ؽ��Ƚ������������ؽ�����Ҫ���������������������������ORA-00054: ��Դ��æ, ��ָ���� NOWAIT ��ʽ��ȡ��Դ


        Exception
          When Others Then
            If SQLErrM Like 'ORA-00054%' Then
              v_Sql := Replace(v_Sql, 'Rebuild', 'Rebuild Online');
              Execute Immediate v_Sql;
            End If;
        End;
      End If;
    End Loop;
  
    --����������(����ת��ʱ��Ӱ��ҵ���ʹ�ã����Բ�֧��)
  Elsif Func_In = 7 And Speedmode_In = 1 Then
    --rebScope_in=0,ֻ�������С��5�ľ��ú���������á�ҩƷ��Ʊ�ݣ�������ȫ������
    For R In (Select a.���� As Table_Name
              From zlBakTables A
              Where a.ֱ��ת�� = 1 And (��� < Decode(Rebscope_In, 0, 5, 100))
              Order By ���, ���) Loop
    
      Update Zldatamovelog
      Set ��ǰ���� = '���������:' || r.Table_Name
      Where ϵͳ = System_In And ���� = (Select Max(����) From Zldatamovelog Where ϵͳ = System_In);
    
      --����п��еĿռ䣬����Ƶ�������ռ䣬ֻ���������ܾ����ƶ��ļ�β�������ݿ飬�Ա���б�ռ��ļ�������
      --��ǰ�������˻Ự����ǿ�Ʋ���
      v_Sql := 'Alter Table ' || r.Table_Name || ' Move Nologging';
      Execute Immediate v_Sql;
    
      --�����ƶ�Lob����
      For L In (Select Column_Name, Tablespace_Name From User_Lobs Where Table_Name = r.Table_Name) Loop
        v_Sql := 'Alter Table ' || r.Table_Name || ' Move Lob(' || l.Column_Name || ') Store as (Tablespace ' ||
                 l.Tablespace_Name || ') Nologging';
        Execute Immediate v_Sql;
      End Loop;
    
      v_Sql := 'Alter Table ' || r.Table_Name || ' Noparallel';
      Execute Immediate v_Sql;
    
      --move�󣬱���ص�������ȫ��ʧЧ����Ҫȫ���ؽ�
      For S In (Select Index_Name
                From User_Indexes
                Where Table_Name = r.Table_Name And Status = 'UNUSABLE' And
                      (Index_Name = r.Table_Name || '_IX_��ת��' Or
                      Index_Name In
                      (Select Upper(Column_Value)
                        From Table(f_Str2list(v_Indexeswithtag))
                        Union
                        Select Upper(Column_Value) From Table(f_Str2list(v_Constraintswithtag))))
                Order By Index_Name) Loop
        Update Zldatamovelog
        Set ��ǰ���� = '���ڻָ�ʧЧ����:' || s.Index_Name
        Where ϵͳ = System_In And ���� = (Select Max(����) From Zldatamovelog Where ϵͳ = System_In);
      
        --��ǰ�������˻Ự����ǿ�Ʋ���
        v_Sql := 'Alter Index ' || s.Index_Name || ' Rebuild Nologging';
        Execute Immediate v_Sql;
      End Loop;
    End Loop;
  End If;

  --ִ���ؽ���������Զ�Ϊ�������ϲ��ж����ԣ������ȡ������Ӱ�����SQL��ִ�мƻ�(ȫ��ɨ��+���в�ѯ������)
  ---------------------------------------------------------------------------------------------------
  If Speedmode_In = 1 And Parallel_In > 1 And (Func_In In (6, 7) Or Func_In In (3, 4) And Enable_In = 1) Then
    Execute Immediate 'ALTER Session DISABLE PARALLEL DDL';
  
    For R In (Select Index_Name From User_Indexes Where Degree Not In ('1', '0')) Loop
      v_Sql := 'Alter Index ' || r.Index_Name || ' Noparallel';
      Execute Immediate v_Sql;
    End Loop;
  End If;

  Update Zldatamovelog
  Set ��ǰ���� = '�ؽ����'
  Where ϵͳ = System_In And ���� = (Select Max(����) From Zldatamovelog Where ϵͳ = System_In);
  Commit;
  --�����̲����д����������ɵ��ù��̴���
End Zl1_Datamove_Reb;
/

Create Or Replace Procedure Zl1_Datamove_Tag
(
  d_End    In Date,
  n_����   In Number,
  n_System In Number
) As
  --���ܣ���Ǵ�ת��������
  --˵����Ϊ����Undo��ռ����͹��󣬷ֶ��ύ
Begin
  --1.���ú��㣨����,ҩƷ,�տ��Ʊ�ݵȣ�

  --*****���⴦������ҽԺ��������:
  --����IDΪ1��"ҽ������",Ϊ2��"�ɲ���":������Ƿ���壬������Ԥ����δ����ģ�ǿ��ת��
  Update /*+ rule*/ ����Ԥ����¼ L
  Set ��ת�� = n_����
  Where ����id In
        (Select Distinct a.����id --1.�����շѺ͹Һŵ��շѽ����¼(�ų�֮���˺ź��˷ѵ�,һ�ŵ�����ֻҪ����һ������)
         From ������ü�¼ A
         Where a.��ת�� Is Null And a.�Ǽ�ʱ�� < d_End And a.��¼���� In (1, 4) And
               (a.��¼״̬ In (1, 2) Or a.��¼״̬ = 3 And Not Exists
                (Select 1
                 From ������ü�¼ B
                 Where a.No = b.No And a.��¼���� = b.��¼���� And b.��¼״̬ = 2 And b.�Ǽ�ʱ�� >= d_End))
         Union All
         Select Distinct a.����id --2.ҽ��������
         From ���ò����¼ A
         Where a.��ת�� Is Null And a.�Ǽ�ʱ�� < d_End And a.��¼���� = 1 And
               (a.��¼״̬ In (1, 2) Or a.��¼״̬ = 3 And Not Exists
                (Select 1
                 From ���ò����¼ B
                 Where a.No = b.No And a.��¼���� = b.��¼���� And b.��¼״̬ In (1, 2) And b.�Ǽ�ʱ�� >= d_End))
         Union All
         Select Distinct a.����id --3.���￨���շѽ����¼(�ų�֮���˿��ѵ�,һ�ŵ�����ֻҪ����һ������)
         From סԺ���ü�¼ A
         Where a.��ת�� Is Null And a.�Ǽ�ʱ�� < d_End And a.��¼���� = 5 And a.���ʷ��� = 0 And
               (a.��¼״̬ In (1, 2) Or a.��¼״̬ = 3 And Not Exists
                (Select 1
                 From סԺ���ü�¼ B
                 Where a.No = b.No And a.��¼���� = b.��¼���� And b.��¼״̬ = 2 And b.�Ǽ�ʱ�� >= d_End))
         Union All --4.����(���ʵ�)��סԺ�Ľ��ʽ����¼
         Select ����id
         From (With Settle As (Select Distinct a.Id As ����id, a.����id --3.����(���ʵ�)��סԺ�Ľ��ʽ����¼(�ų�֮��������ϵ�)
                               From ���˽��ʼ�¼ A
                               Where a.��ת�� Is Null And a.�շ�ʱ�� < d_End And
                                     (a.��¼״̬ In (1, 2) Or a.��¼״̬ = 3 And Not Exists
                                      (Select 1
                                       From ���˽��ʼ�¼ B
                                       Where a.No = b.No And b.��¼״̬ = 2 And b.�շ�ʱ�� >= d_End)))
                Select ����id
                From Settle
                Minus
                --1.һ��Ԥ�����ʽ��ʳ��꣨����ID��ͬ������Щ����IDҪ�����ų�,���ⲿ�ֱ�ת����Ӱ������ļ����Ƿ���� 
                --2.��Щ���õ��ݵĽ���ID��Ӧ�Ŀ��ܻ�������NO����������ID(�������Ϻ�ֶ�ν��ʽ��壬���ܲ�����ת��ʱ��֮��)����Щ����IDҪ�����ų�,���ⲿ�ֱ�ת����Ӱ������ļ����Ƿ����
                --���ǵ�������ĸ����ԣ�Ϊ���߼���������ѯ���ܣ�������ID���ų�
                Select Distinct d.Id
                From ���˽��ʼ�¼ D,
                     (Select Distinct c.����id --���סԺ����һ��ᣬ�Լ�������ʺ�סԺ���ʿ���һ����ҳ�ͬһ��Ԥ�����������ﲻ����ҳID
                       From סԺ���ü�¼ C,
                            (Select Distinct d.No, d.���, Mod(d.��¼����, 10) As ��¼����
                              From סԺ���ü�¼ D,
                                   (Select s.����id
                                     From Settle S, ���˽��ʼ�¼ E
                                     Where s.����id = e.����id And
                                           (e.�շ�ʱ�� > d_End Or Exists (Select 1 From ��Ժ���� F Where s.����id = f.����id))) S --û�н�����֮��û���ٽ���ͳ��˴��ʣ����־Ͳ��ų�
                              Where d.����id = s.����id) D
                       Where c.No = d.No And Mod(c.��¼����, 10) = d.��¼���� And c.��� = d.��� --���ʺ����Ϻ��ٶ԰������ʵ����ʵĽ���IDΪ�յļ�¼,һ����ܼ����Ƿ����,���ֽ���IDΪ�յ�����ת���ں��浥��ת��                                        
                       Group By c.No, Mod(c.��¼����, 10), c.����id --һ�ŵ����е�һ�пɲ��ֽ��ʣ��Ե���Ϊ�������жϣ�����һ�ŵ��ݵ�����һ���ֱ�ת��
                       Having Nvl(Sum(c.ʵ�ս��), 0) <> Nvl(Sum(c.���ʽ��), 0) Or Exists (Select 1 --�ų�ת��ʱ��֮���ٴν��ʵ�(���Ϻ��ٴν���)������ԭʼ����ת�ߺ󣬺�������ʱ�޷���ȷ�ж�
                                                                                   From סԺ���ü�¼ E, ���˽��ʼ�¼ S
                                                                                   Where e.No = c.No And
                                                                                         Mod(e.��¼����, 10) = Mod(c.��¼����, 10) And
                                                                                         e.��¼���� In (12, 13, 15) And
                                                                                         e.����id = s.Id And s.�շ�ʱ�� >= d_End)
                       Union All
                       Select Distinct c.����id
                       From ������ü�¼ C,
                            (Select Distinct d.No, d.���, Mod(d.��¼����, 10) As ��¼����
                              From ������ü�¼ D, Settle S
                              Where d.����id = s.����id) D --��Ϊ�����ﲡ�ˣ����ԣ�ֻҪû�н���,�ò��˵Ķ���ת��
                       Where c.No = d.No And Mod(c.��¼����, 10) = d.��¼���� And c.��� = d.���
                       Group By c.No, Mod(c.��¼����, 10), c.����id
                       Having Nvl(Sum(c.ʵ�ս��), 0) <> Nvl(Sum(c.���ʽ��), 0) Or Exists (Select 1
                                                                                    From ������ü�¼ E, ���˽��ʼ�¼ S
                                                                                    Where e.No = c.No And
                                                                                          Mod(e.��¼����, 10) = Mod(c.��¼����, 10) And
                                                                                          e.��¼���� In (12, 13, 15) And
                                                                                          e.����id = s.Id And s.�շ�ʱ�� >= d_End)) N
                Where d.����id = n.����id)
                
         
         );

  --�ų�Ԥ����δ����ĺ�ת��ʱ��֮��ҩ�ļ�¼
  --��Ϊǰ���SQL����Ľ���ID���ܲ�ȫ�ǳ�Ԥ����(�����շѺ�סԺ���ʲ��ѵ�)�����ԣ���Ҫ����һ��SQL���ų�
  --���ڿ��ܴ��������쳣(סԺ���ý��ʳ�Ԥ�����Ϊ1������Ԥ��)������û�м�Ԥ����������޶�
  Update /*+ rule*/ ����Ԥ����¼
  Set ��ת�� = Null
  Where ��ת�� = n_���� And
        ����id In
        (Select Distinct d.����id
         From ����Ԥ����¼ D,
              --����D����Ϊ�˲��ͬһԤ�����ݵ���������ID����Ԥ�����Ԥ�����ϵģ��ٴγ�ͬһԤ�����ݣ�
              --�ò��˵����н���ID�Ķ���ת�������ⲿ�ֳ�Ԥ���Ľ���ID���ų���ԭʼԤ������ת�ߣ�������������ID�����õ��ݵ�һ����(ԭʼ���ʡ��������ϡ��ٴν�һ���֡��ٴν�ȫ��)ת��
              (Select Distinct l.����id
                From ����Ԥ����¼ L, ����Ԥ����¼ P --���ܱ��ν��ʳ��ֻ��ʣ��������Ҫ����L����ԭʼ��Ԥ���ĵ��ݣ��Լ���¼����Ϊ11�Ŀ��ܻ���ת��ʱ��֮��������ʣ���Ľ���ID
                Where l.��¼���� In (1, 11) And l.No = p.No And p.��¼���� In (1, 11) And p.��ת�� = n_����
                Group By l.No, l.����id
                Having Nvl(Sum(l.���), 0) <> Nvl(Sum(l.��Ԥ��), 0) And (Exists (Select 1
                                                                           From ����Ԥ����¼ E
                                                                           Where l.����id = e.����id And e.�տ�ʱ�� > d_End) Or Exists (Select 1
                                                                                                                               From ��Ժ���� E
                                                                                                                               Where l.����id =
                                                                                                                                     e.����id)) --û�г�����֮��û���ٳ���������ͳ��˴��ʣ������ø��Ľ��ʲ�������ʾ��Ԥ�����ɳ��������������־Ͳ��ų� 
                Or Nvl(Sum(l.���), 0) = Nvl(Sum(l.��Ԥ��), 0) And (Exists (Select 1
                                                                      From ����Ԥ����¼ E, ���˽��ʼ�¼ F --�ų�ת��ʱ��֮�����������ID���
                                                                      Where e.No = l.No And e.��¼���� = 11 And e.����id = f.Id And --��Ԥ��ʱ���տ�ʱ������ǽ�Ԥ�����ʱ�䣬����������Ҫ���������ʱ��
                                                                            f.�շ�ʱ�� >= d_End) Or Exists (Select 1
                                                                                                       From ����Ԥ����¼ E,
                                                                                                            ������ü�¼ F
                                                                                                       Where e.No = l.No And
                                                                                                             e.��¼���� = 11 And
                                                                                                             e.����id =
                                                                                                             f.����id And
                                                                                                             f.�Ǽ�ʱ�� >=
                                                                                                             d_End And
                                                                                                             f.��¼���� In
                                                                                                             (1, 4) And
                                                                                                             Nvl(f.���ʷ���, 0) <> 1) Or Exists (Select 1
                                                                                                                                            From ����Ԥ����¼ E,
                                                                                                                                                 סԺ���ü�¼ F
                                                                                                                                            Where e.No = l.No And
                                                                                                                                                  e.��¼���� = 11 And
                                                                                                                                                  e.����id =
                                                                                                                                                  f.����id And
                                                                                                                                                  f.�Ǽ�ʱ�� >=
                                                                                                                                                  d_End And
                                                                                                                                                  f.��¼���� In (5,
                                                                                                                                                             15) And
                                                                                                                                                  Nvl(f.���ʷ���,
                                                                                                                                                      0) <> 1))) N
         
         Where d.����id = n.����id);

  --Ϊ�˽����߼��ĸ����ԣ����ų���ת��ʱ��֮��ҩ��δ��ҩ�ķ��ü�¼��Ӧ�Ľ���ID������������Ľ������ݺͷ�������ǿ��ת�� 

  --Ԥ����û��ʹ�þ�ֱ�����˵ļ�¼(����IDΪ��)
  Update /*+ rule*/ ����Ԥ����¼
  Set ��ת�� = n_����
  Where ��¼���� = 1 And
        NO In (Select a.No
               From ����Ԥ����¼ A
               Where a.����id Is Null And a.��¼���� = 1 And a.��¼״̬ In (2, 3) And a.��ת�� Is Null And a.�տ�ʱ�� < d_End
               Group By a.No
               Having Sum(a.���) = 0);

  --��Ԥ�������ϵļ�¼����¼����Ϊ2����û�н���ID
  Update /*+ rule*/ ����Ԥ����¼
  Set ��ת�� = n_����
  Where ����id Is Null And ��¼���� = 2 And NO In (Select a.No From ����Ԥ����¼ A Where a.��ת�� = n_���� And a.��¼���� = 3);

  Update Zldatamovelog
  Set ��ǰ���� = '(1/10)�������ݱ����ɣ����ڱ�Ƿ�������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  Update /*+ rule*/ ���˽��ʼ�¼
  Set ��ת�� = n_����
  Where ID In (Select ����id From ����Ԥ����¼ Where ��ת�� = n_����);

  --�����޽���ļ�¼(Ϊ���������ܣ����жϷ��ã�ֻҪ����������Ԥ����¼�͵���������ý���)
  Update /*+ rule*/ ���˽��ʼ�¼ L
  Set ��ת�� = n_����
  Where �շ�ʱ�� < d_End And ��ת�� Is Null And Not Exists (Select 1 From ����Ԥ����¼ P Where l.Id = p.����id);

  Update /*+ rule*/ ���˿��������
  Set ��ת�� = n_����
  Where Ԥ��id In (Select a.Id From ����Ԥ����¼ A Where ��ת�� = n_����);

  Update /*+ rule*/ ���˿������¼
  Set ��ת�� = n_����
  Where ID In (Select ������id From ���˿�������� Where ��ת�� = n_����);

  Update /*+ rule*/ �������㽻��
  Set ��ת�� = n_����
  Where ����id In (Select a.Id From ����Ԥ����¼ A Where ��ת�� = n_����);

  --�ҺŴ��ۺ�ʵ�ս��Ϊ0��(û�ж�Ӧ��Ԥ����¼),��ʹ֮�����˺ŷ���Ҳ���ܣ���Ϊ���Ϊ�㲻Ӱ�����),�����Ѽ�ʹΪ��Ҳ��Ԥ����¼                 
  --���ݹҺż�¼����������ã���ֱ�Ӱ�ʱ����������Ҫ��
  Update /*+ rule*/ ������ü�¼
  Set ��ת�� = n_����
  Where NO In (Select NO From ���˹Һż�¼ Where ��ת�� Is Null And �Ǽ�ʱ�� < d_End) And ��¼���� = 4 And ʵ�ս�� = 0;

  --û�н��ʵ��ѳ����ļ��ʵ�����ۺ�ʵ�ս��Ϊ��ģ���û���������ʵ���ǿ��ת��
  Update /*+ rule*/ ������ü�¼
  Set ��ת�� = n_����
  Where ��¼���� = 2 And
        NO In
        (Select NO
         From (Select b.�Һŵ�, a.No, a.���, Sum(a.ʵ�ս��)
                From ������ü�¼ A, ����ҽ����¼ B
                Where a.ҽ����� = b.Id And a.����id Is Null And a.��¼���� = 2 And b.������Դ <> 4 And a.��ת�� Is Null And a.�Ǽ�ʱ�� < d_End
                Group By a.No, a.���, b.�Һŵ�
                Having Sum(a.ʵ�ս��) = 0 And Not Exists (Select 1
                                                      From ������ü�¼ C, ����ҽ����¼ D
                                                      Where b.�Һŵ� = d.�Һŵ� And d.Id = c.ҽ����� And d.������Դ <> 4 And c.��¼���� = 2 And
                                                            c.��ת�� Is Null
                                                      Group By c.No, c.���
                                                      Having Sum(a.ʵ�ս��) > 0)));

  --ֱ���շѵĺͽ����޽��㣨Ԥ������¼�ģ�Union����allȥ���ظ��Լ���in������
  Update /*+ rule*/ ������ü�¼
  Set ��ת�� = n_����
  Where ����id In (Select ����id
                 From ����Ԥ����¼
                 Where ��ת�� = n_����
                 Union
                 Select ID From ���˽��ʼ�¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ���ò����¼
  Set ��ת�� = n_����
  Where ����id In (Select ����id From ����Ԥ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ƾ����ӡ��¼
  Set ��ת�� = n_����
  Where (NO, ��¼����) In (Select NO, ��¼���� From ������ü�¼ Where ��ת�� = n_����);

  --��Ԥ����¼����Ϊ��ȡ���￨ֱ���շѵģ��޽���ID��,�ټӽ��ʼ�¼��Ϊ��ȡ�����޽��㣨Ԥ������¼��
  Update /*+ rule*/ סԺ���ü�¼
  Set ��ת�� = n_����
  Where ����id In (Select ����id
                 From ����Ԥ����¼
                 Where ��ת�� = n_����
                 Union
                 Select ID From ���˽��ʼ�¼ Where ��ת�� = n_����);

  --1.ת���������Ϻ󣬼��ʵ����ʵļ�¼������״̬Ϊ2��û�н���ID��(��¼״̬Ϊ3���н���ID��)����ǰ����ת����
  --2.δ���ʵ������(�ѳ����ļ��ʵ�)
  --3.û�н���ID�Ļ��ۼ�¼����Ϊת��
  --4.���շ�Ҳû�г�Ԥ��������ô���Ϊת��
  --������"��ת�� Is Null"��Ϊ�˴���������α��ת�������
  Update /*+ rule*/ ������ü�¼ A
  Set ��ת�� = n_����
  Where ((Exists (Select 1
                  From ������ü�¼ B
                  Where a.No = b.No And a.��¼���� = b.��¼���� And a.��� = b.��� And b.��¼״̬ = 3 And b.����id Is Not Null And
                        b.��ת�� + 0 = n_����) And ��¼״̬ = 2 Or Exists
         (Select 1
           From ������ü�¼ B
           Where a.No = b.No And a.��¼���� = b.��¼���� And a.��� = b.���
           Group By b.No, b.��¼����, b.���
           Having Nvl(Sum(b.ʵ�ս��), 0) = 0)) And ��¼���� = 2 Or ��¼״̬ = 0 Or ��¼���� = 1 And ʵ�ս�� = 0 And ���ʽ�� = 0) And
        ����id Is Null And ��ת�� Is Null And �Ǽ�ʱ�� < d_End;

  --1.ת���������Ϻ󣬼��ʵ����ʵļ�¼������״̬Ϊ2��û�н���ID��(��¼״̬Ϊ3���н���ID��)����ǰ����ת����
  --2.δ���ʵ������(�ѳ����ļ��ʵ�)
  --3.û�н���ID�Ļ��ۼ�¼����Ϊת��
  Update /*+ rule*/ סԺ���ü�¼ A
  Set ��ת�� = n_����
  Where ((Exists (Select 1
                  From סԺ���ü�¼ B
                  Where a.No = b.No And a.��¼���� = b.��¼���� And a.��� = b.��� And b.��¼״̬ = 3 And b.����id Is Not Null And
                        b.��ת�� + 0 = n_����) And ��¼״̬ = 2 Or Exists
         (Select 1
           From סԺ���ü�¼ B
           Where a.No = b.No And a.��¼���� = b.��¼���� And a.��� = b.���
           Group By b.No, b.��¼����, b.���
           Having Nvl(Sum(b.ʵ�ս��), 0) = 0)) And ��¼���� = 2 Or ��¼״̬ = 0) And ����id Is Null And ��ת�� Is Null And �Ǽ�ʱ�� < d_End;

  --���ڴ������ʲ�����Ժδ�����������ںܾ���ǰ����Щ���ݣ����Ԥ���ѳ��꣬����ΪҪת��
  Update /*+ rule*/ סԺ���ü�¼ A
  Set ��ת�� = n_����
  Where ��ת�� Is Null And ����id Is Null And
        (����id, ��ҳid) In (Select ����id, ��ҳid
                         From ������ҳ C
                         Where ��Ժ���� < d_End And ��ת�� Is Null And ����ת�� Is Null And Not Exists
                          (Select 1
                                From ����Ԥ����¼ B
                                Where b.����id = c.����id And b.Ԥ����� = 2 And b.��¼���� In (1, 11) Having
                                 Nvl(Sum(b.���), 0) - Nvl(Sum(b.��Ԥ��), 0) <> 0));

  Update Zldatamovelog
  Set ��ǰ���� = '(2/10)�������ݱ����ɣ����ڱ��ҩƷ����'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  Update /*+ Rule*/ ҩƷ�շ���¼ A
  Set ��ת�� = n_����
  Where Rowid In (Select m.Rowid
                  From ҩƷ�շ���¼ M, ������ü�¼ E
                  Where m.����id = e.Id And (e.��¼���� = 1 And m.���� In (8, 24) Or e.��¼���� = 2 And m.���� In (9, 25)) And
                        e.�շ���� In ('4', '5', '6', '7') And e.��ת�� = n_����
                  Union All
                  Select m.Rowid
                  From ҩƷ�շ���¼ M, סԺ���ü�¼ E
                  Where m.����id = e.Id And m.���� In (9, 10, 25, 26) And e.��¼���� = 2 And e.�շ���� In ('4', '5', '6', '7') And
                        e.��ת�� = n_����);

  Update /*+ rule*/ �շ���¼������Ϣ
  Set ��ת�� = n_����
  Where �շ�id In (Select ID From ҩƷ�շ���¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ��Һ��ҩ����
  Set ��ת�� = n_����
  Where �շ�id In (Select ID From ҩƷ�շ���¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ��Һ��ҩ��¼
  Set ��ת�� = n_����
  Where ID In (Select ��¼id From ��Һ��ҩ���� Where ��ת�� = n_����);

  Update /*+ rule*/ ��Һ��ҩ����
  Set ��ת�� = n_����
  Where ��ҩid In (Select ID From ��Һ��ҩ��¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ��Һ��ҩ״̬
  Set ��ת�� = n_����
  Where ��ҩid In (Select ID From ��Һ��ҩ��¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ҩƷ����ƻ�
  Set ��ת�� = n_����
  Where ����id In (Select ID From ҩƷ�շ���¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ҩƷǩ����ϸ
  Set ��ת�� = n_����
  Where �շ�id In (Select ID From ҩƷ�շ���¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ҩƷǩ����¼
  Set ��ת�� = n_����
  Where ID In (Select ǩ��id From ҩƷǩ����ϸ Where ��ת�� = n_����);

  Update Zldatamovelog
  Set ��ǰ���� = '(3/10)ҩƷ���ݱ����ɣ����ڱ�ǽɿ���Ʊ������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  Update /*+ rule*/ ��Ա����¼ Set ��ת�� = n_���� Where ��ת�� Is Null And ���ʱ�� < d_End;

  Update /*+ rule*/ ��Ա�սɼ�¼ Set ��ת�� = n_���� Where ��ת�� Is Null And �Ǽ�ʱ�� < d_End;

  Update /*+ rule*/ ��Ա�սɶ���
  Set ��ת�� = n_����
  Where �ս�id In (Select ID From ��Ա�սɼ�¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ��Ա�ս���ϸ
  Set ��ת�� = n_����
  Where �ս�id In (Select ID From ��Ա�սɼ�¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ��Ա�ս�Ʊ��
  Set ��ת�� = n_����
  Where �ս�id In (Select ID From ��Ա�սɼ�¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ��Ա�ݴ��¼
  Set ��ת�� = n_����
  Where �ս�id In (Select ID From ��Ա�սɼ�¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ��Ա�ݴ��¼ Set ��ת�� = n_���� Where ��ת�� Is Null And ��¼���� = 1 And �Ǽ�ʱ�� < d_End;

  Update /*+ rule*/ Ʊ�����ü�¼ A
  Set ��ת�� = n_����
  Where Not Exists
   (Select 1 From Ʊ��ʹ����ϸ B Where b.����id = a.Id And b.ʹ��ʱ�� >= d_End) And ��ת�� Is Null And ʣ������ = 0 And �Ǽ�ʱ�� < d_End;

  Update /*+ rule*/ Ʊ��ʹ����ϸ
  Set ��ת�� = n_����
  Where ����id In (Select ID From Ʊ�����ü�¼ Where ��ת�� = n_����);
  Update /*+ rule*/ Ʊ�ݴ�ӡ����
  Set ��ת�� = n_����
  Where ID In (Select ��ӡid From Ʊ��ʹ����ϸ Where ��ת�� = n_����);

  Update /*+ rule*/ Ʊ�ݴ�ӡ��ϸ
  Set ��ת�� = n_����
  Where ʹ��id In (Select ID From Ʊ��ʹ����ϸ Where ��ת�� = n_����);

  Update Zldatamovelog
  Set ��ǰ���� = '(4/10)�ɿ���Ʊ�����ݱ����ɣ����ڱ�Ǿ��Ｐ��������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  --2.���Ｐ��������
  --��ת�����������Һŷ���δת���ģ�ת��ʱ��֮�����ҽ����ҽ����Ӧ�ķ���δת����
  --��ʹ���ھ���(r.ִ��״̬ <> 2 )��Ҳǿ��ת��
  Update /*+ rule*/ ���˹Һż�¼ T
  Set ��ת�� = n_����
  Where Rowid In
        (Select Rowid
         From ���˹Һż�¼ R
         Where Not Exists (Select 1
                From ������ü�¼ A
                Where r.No = a.No And a.�Ǽ�ʱ�� < d_End And a.��¼���� = 4 And a.��ת�� Is Null) And Not Exists
          (Select 1
                From ����ҽ����¼ A
                Where a.�Һŵ� = r.No And a.������Դ <> 4 And Nvl(a.ͣ��ʱ��, a.����ʱ��) >= d_End) And Not Exists
          (Select 1
                From ������ü�¼ E, ����ҽ����¼ A
                Where r.No = a.�Һŵ� And a.Id = e.ҽ����� And a.������Դ <> 4 And e.��ת�� Is Null) And r.��ת�� Is Null And
               r.�Ǽ�ʱ�� < d_End);

  --������һ���ֹҺ�����δת�������ԣ����ܱ�����ݿ�����Һ����ݲ�ƥ��
  Update ���˹ҺŻ��� Set ��ת�� = n_���� Where ��ת�� Is Null And ���� < d_End;
  Update /*+ rule*/ ����ת���¼ Set ��ת�� = n_���� Where NO In (Select NO From ���˹Һż�¼ Where ��ת�� = n_����);

  --ͨ��"סԺ���ü�¼"����ѯ��������"���˽��ʼ�¼",��Ϊ��Ժδ������ʲ���Ҳת���˷���
  --��Ժ����������Ȼ��Ҫ����Ϊ����ĳ�ν���ת���ˣ������˵�ʱ��δ��Ժ(һ��סԺ��ν���)��
  --ͨ��ָ��������ʽ���������Ż���ȱʡ����"������ҳIX_��Ժ����"������Ч��̫�ͣ�
  Update /*+ rule*/ ������ҳ P
  Set ��ת�� = n_����
  Where Not Exists (Select 1 From סԺ���ü�¼ A Where a.����id = p.����id And a.��ҳid = p.��ҳid And a.��ת�� Is Null) And ��ת�� Is Null And
        ����ת�� Is Null And ��Ժ���� < d_End And
        (����id, ��ҳid) In (Select Distinct ����id, ��ҳid From סԺ���ü�¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ���˹�����¼
  Set ��ת�� = n_����
  Where (����id, ��ҳid) In (Select ����id, ID
                         From ���˹Һż�¼
                         Where ��ת�� = n_����
                         Union All
                         Select ����id, ��ҳid From ������ҳ Where ��ת�� = n_����);

  Update /*+ rule*/ ������ϼ�¼
  Set ��ת�� = n_����
  Where (����id, ��ҳid) In (Select ����id, ID
                         From ���˹Һż�¼
                         Where ��ת�� = n_����
                         Union All
                         Select ����id, ��ҳid From ������ҳ Where ��ת�� = n_����);

  Update /*+ rule*/ ���������¼
  Set ��ת�� = n_����
  Where (����id, ��ҳid) In (Select ����id, ID
                         From ���˹Һż�¼
                         Where ��ת�� = n_����
                         Union All
                         Select ����id, ��ҳid From ������ҳ Where ��ת�� = n_����);

  Update Zldatamovelog
  Set ��ǰ���� = '(5/10)���Ｐ�������ݱ����ɣ����ڱ�ǻ�������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  --3.��������
  Update /*+ rule*/ ���˻����ļ�
  Set ��ת�� = n_����
  Where (����id, ��ҳid) In (Select ����id, ��ҳid From ������ҳ Where ��ת�� = n_����);
  Update /*+ rule*/ ���˻�������
  Set ��ת�� = n_����
  Where �ļ�id In (Select ID From ���˻����ļ� Where ��ת�� = n_����);
  Update /*+ rule*/ ���˻�����ϸ
  Set ��ת�� = n_����
  Where ��¼id In (Select ID From ���˻������� Where ��ת�� = n_����);

  Update /*+ rule*/ ���˻����ӡ
  Set ��ת�� = n_����
  Where �ļ�id In (Select ID From ���˻����ļ� Where ��ת�� = n_����);
  Update /*+ rule*/ ���˻�����Ŀ
  Set ��ת�� = n_����
  Where �ļ�id In (Select ID From ���˻����ļ� Where ��ת�� = n_����);
  Update /*+ rule*/ ���˻���Ҫ������
  Set ��ת�� = n_����
  Where �ļ�id In (Select ID From ���˻����ļ� Where ��ת�� = n_����);
  Update /*+ rule*/ ����Ҫ������
  Set ��ת�� = n_����
  Where �ļ�id In (Select ID From ���˻����ļ� Where ��ת�� = n_����);

  --�ϰ滤��ϵͳ����
  Update /*+ rule*/ ���˻����¼
  Set ��ת�� = n_����
  Where (����id, ��ҳid) In (Select ����id, ��ҳid From ������ҳ Where ��ת�� = n_����);
  Update /*+ rule*/ ���˻�������
  Set ��ת�� = n_����
  Where ��¼id In (Select ID From ���˻����¼ Where ��ת�� = n_����);

  Update Zldatamovelog
  Set ��ǰ���� = '(6/10)�������ݱ����ɣ����ڱ�ǲ�������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  --4.��������
  Update /*+ rule*/ ���Ӳ�����¼
  Set ��ת�� = n_����
  Where ������Դ <> 4 And (����id, ��ҳid) In (Select ����id, ID
                                       From ���˹Һż�¼
                                       Where ��ת�� = n_����
                                       Union All
                                       Select ����id, ��ҳid From ������ҳ Where ��ת�� = n_����);

  --�ԵǼ��ಡ��(�޹Һŵ���)
  --����ID�����ظ�����Ϊ���鱨��֮��ģ���ι�����������һ�ű��棬���ڲ���ҽ��������У����ҽ��id��Ӧͬһ����ID
  --����ҽ�����ͼ�¼��ִ��״̬����Ϊ����û������ִ�еǼ�
  Update /*+ rule*/ ���Ӳ�����¼
  Set ��ת�� = n_����
  Where ID In (Select c.����id
               From ����ҽ������ A, ����ҽ����¼ B, ����ҽ������ C
               Where c.ҽ��id = b.Id And b.Id = a.ҽ��id And b.���id Is Null And Nvl(b.��ҳid, 0) = 0 And b.�Һŵ� Is Null And
                     a.�������� = 1 And a.��ת�� Is Null And a.����ʱ�� < d_End);

  Update /*+ rule*/ ���Ӳ�������
  Set ��ת�� = n_����
  Where ����id In (Select ID From ���Ӳ�����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ���Ӳ�����ʽ
  Set ��ת�� = n_����
  Where �ļ�id In (Select ID From ���Ӳ�����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ���Ӳ�������
  Set ��ת�� = n_����
  Where �ļ�id In (Select ID From ���Ӳ�����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ���Ӳ���ͼ��
  Set ��ת�� = n_����
  Where ����id In (Select ID From ���Ӳ������� Where ��ת�� = n_���� And �������� = 5);

  Update /*+ rule*/ ����ҽ������
  Set ��ת�� = n_����
  Where ����id In (Select ID From ���Ӳ�����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ Ӱ�񱨸沵��
  Set ��ת�� = n_����
  Where (ҽ��id, ����id) In (Select ҽ��id, ����id From ����ҽ������ Where ��ת�� = n_����);

  Update /*+ rule*/ ������ļ�¼
  Set ��ת�� = n_����
  Where ����id In (Select ID From ���Ӳ�����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ �����걨��¼
  Set ��ת�� = n_����
  Where �ļ�id In (Select ID From ���Ӳ�����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ������ϼ�¼
  Set ��ת�� = n_����
  Where ����id In (Select ID From ���Ӳ�����¼ Where ��ת�� = n_����);

  Update Zldatamovelog
  Set ��ǰ���� = '(7/10)�������ݱ����ɣ����ڱ���ٴ�·������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  --5.�ٴ�·��    
  Update /*+ rule*/ �����ٴ�·��
  Set ��ת�� = n_����
  Where (����id, ��ҳid) In (Select ����id, ��ҳid From ������ҳ Where ��ת�� = n_����);
  Update /*+ rule*/ ���˺ϲ�·��
  Set ��ת�� = n_����
  Where ��Ҫ·����¼id In (Select ID From �����ٴ�·�� Where ��ת�� = n_����);
  Update /*+ rule*/ ���˺ϲ�·������
  Set ��ת�� = n_����
  Where ·����¼id In (Select ID From �����ٴ�·�� Where ��ת�� = n_����);

  Update /*+ rule*/ ���˳�����¼
  Set ��ת�� = n_����
  Where ·����¼id In (Select ID From �����ٴ�·�� Where ��ת�� = n_����);

  Update /*+ rule*/ ����·��ִ��
  Set ��ת�� = n_����
  Where ·����¼id In (Select ID From �����ٴ�·�� Where ��ת�� = n_����);
  Update /*+ rule*/ ����·������
  Set ��ת�� = n_����
  Where ·����¼id In (Select ID From �����ٴ�·�� Where ��ת�� = n_����);
  Update /*+ rule*/ ����·������
  Set ��ת�� = n_����
  Where ·����¼id In (Select ID From �����ٴ�·�� Where ��ת�� = n_����);
  Update /*+ rule*/ ����·��ָ��
  Set ��ת�� = n_����
  Where ·����¼id In (Select ID From �����ٴ�·�� Where ��ת�� = n_����);
  Update /*+ rule*/ ����·��ҽ��
  Set ��ת�� = n_����
  Where ·��ִ��id In (Select ID From ����·��ִ�� Where ��ת�� = n_����);

  Update Zldatamovelog
  Set ��ǰ���� = '(8/10)�ٴ�·�����ݱ����ɣ����ڱ��ҽ������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  --6.ҽ�������飬���
  Update /*+ rule*/ ����ҽ����¼
  Set ��ת�� = n_����
  Where �Һŵ� In (Select NO From ���˹Һż�¼ Where ��ת�� = n_����) And ������Դ <> 4;
  Update /*+ rule*/ ����ҽ����¼
  Set ��ת�� = n_����
  Where (����id, ��ҳid) In (Select ����id, ��ҳid From ������ҳ Where ��ת�� = n_����);

  --�ԵǼ��ಡ��(�޹Һŵ�)������ҽ��������ǰ��ת����ʱ��ת��
  Update /*+ rule*/ ����ҽ����¼
  Set ��ת�� = n_����
  Where Rowid In (Select b.Rowid
                  From ����ҽ����¼ B, ����ҽ������ C
                  Where (b.���id = c.ҽ��id Or b.Id = c.ҽ��id) And c.��ת�� = n_����);

  Update /*+ rule*/ ����ҽ���Ƽ�
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ����ҽ������
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ����ҽ������
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ��Ѫ�����¼
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ��Ѫ������
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ����ҽ��ִ��
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ����ҽ����ӡ
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ҽ��ִ�д�ӡ
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ �������ҽ��
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ������ϼ�¼
  Set ��ת�� = n_����
  Where ID In (Select ���id From �������ҽ�� Where ��ת�� = n_����);

  Update /*+ rule*/ ����ҽ��״̬
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ҽ��ǩ����¼
  Set ��ת�� = n_����
  Where ID In (Select ǩ��id From ����ҽ��״̬ Where ��ת�� = n_���� And ǩ��id Is Not Null);

  Update /*+ rule*/ ����ҽ������
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ���Ƶ��ݴ�ӡ
  Set ��ת�� = n_����
  Where (NO, ��¼����) In (Select NO, ��¼���� From ����ҽ������ Where ��ת�� = n_����);

  Update /*+ rule*/ ҽ��ִ��ʱ��
  Set ��ת�� = n_����
  Where (ҽ��id, ���ͺ�) In (Select ҽ��id, ���ͺ� From ����ҽ������ Where ��ת�� = n_����);

  Update /*+ rule*/ ҽ��ִ�мƼ�
  Set ��ת�� = n_����
  Where (ҽ��id, ���ͺ�) In (Select ҽ��id, ���ͺ� From ����ҽ������ Where ��ת�� = n_����);

  Update /*+ rule*/ ִ�д�ӡ��¼
  Set ��ת�� = n_����
  Where (ҽ��id, ���ͺ�) In (Select ҽ��id, ���ͺ� From ����ҽ������ Where ��ת�� = n_����);

  Update /*+ rule*/ ���������ϸ
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ��������¼
  Set ��ת�� = n_����
  Where ID In (Select ��id From ���������ϸ Where ��ת�� = n_����);

  Update /*+ rule*/ ���������
  Set ��ת�� = n_����
  Where ��id In (Select ID From ��������¼ Where ��ת�� = n_����);

  Update Zldatamovelog
  Set ��ǰ���� = '(9/10)ҽ�����ݱ����ɣ����ڱ�Ǽ���������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  Update /*+ rule*/ Ӱ�����¼
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ Ӱ��������
  Set ��ת�� = n_����
  Where ���uid In (Select ���uid From Ӱ�����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ Ӱ����ͼ��
  Set ��ת�� = n_����
  Where ����uid In (Select ����uid From Ӱ�������� Where ��ת�� = n_����);

  Update /*+ rule*/ Ӱ�����뵥ͼ��
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ Ӱ���ղ�����
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ Ӱ��Σ��ֵ��¼
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update Zldatamovelog
  Set ��ǰ���� = '(10/10)Ӱ�����ݱ����ɣ����ڱ�Ǽ�������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  Update /*+ rule*/ ����걾��¼
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ����������Ŀ
  Set ��ת�� = n_����
  Where �걾id In (Select ID From ����걾��¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ������Ŀ�ֲ�
  Set ��ת�� = n_����
  Where �걾id In (Select ID From ����걾��¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ���������¼
  Set ��ת�� = n_����
  Where �걾id In (Select ID From ����걾��¼ Where ��ת�� = n_����);
  Update /*+ rule*/ �����ʿؼ�¼
  Set ��ת�� = n_����
  Where �걾id In (Select ID From ����걾��¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ���������¼
  Set ��ת�� = n_����
  Where �걾id In (Select ID From ����걾��¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ����ǩ����¼
  Set ��ת�� = n_����
  Where ����걾id In (Select ID From ����걾��¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ����ͼ����
  Set ��ת�� = n_����
  Where �걾id In (Select ID From ����걾��¼ Where ��ת�� = n_����);

  Update /*+ rule*/ �����Լ���¼
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ������ռ�¼
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ������ͨ���
  Set ��ת�� = n_����
  Where ����걾id In (Select ID From ����걾��¼ Where ��ת�� = n_����);
  Update /*+ rule*/ �����ʿر���
  Set ��ת�� = n_����
  Where ���id In (Select ID From ������ͨ��� Where ��ת�� = n_����);
  Update /*+ rule*/ ����ҩ�����
  Set ��ת�� = n_����
  Where ϸ�����id In (Select ID From ������ͨ��� Where ��ת�� = n_����);
  Update /*+ rule*/ ������ˮ�߱걾
  Set ��ת�� = n_����
  Where �걾id In (Select ID From ����걾��¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ������ˮ��ָ��
  Set ��ת�� = n_����
  Where �걾id In (Select ID From ����걾��¼ Where ��ת�� = n_����);

  Commit;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl1_Datamove_Tag;
/



---------------------
--zlAppData
---------------------
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 1, 'A01', 'ҩƷƤ��', '�涨������Ƥ�Ե�ҩƷ������ҽʦ�Ƿ�ע���������鼰������ж�', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 1, 'A02', '��ҩ���ٴ���ϵ������', '������ҩ���ٴ���ϵ������', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 1, 'A03', '�������÷�����ȷ��', '�������÷�����ȷ��', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 1, 'A04', '�������ҩ;���ĺ�����', 'ѡ�ü������ҩ;���ĺ�����', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 1, 'A05', '�Ƿ����ظ���ҩ����', '�Ƿ����ظ���ҩ����', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 1, 'A06', 'ҩ���໥���ú��������', '�Ƿ���Ǳ���ٴ������ҩ���໥���ú��������', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 1, 'A07', '������ҩ���������', '������ҩ���������', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 2, '1-1', '��������ȱ', '������ǰ�ǡ����ġ��������ȱ��', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 2, '1-2', 'ҩʦǩ��ǩ�²�һ��', 'ҽʦǩ����ǩ�²��淶������ǩ����ǩ�µ�������һ�µ�', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 2, '1-3', 'ҩʦδ�Դ���������������˵�', 'ҩʦδ�Դ���������������˵ģ�������ǵ���ˡ����䡢�˶ԡ���ҩ��Ŀ����˵���ҩʦ���˶Է�ҩҩʦǩ�������ߵ���ֵ�����δִ��˫ǩ���涨��', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 2, '1-4', '��������Ӥ�׶�δд���ա�����', '��������Ӥ�׶�����δд���ա������', 0, 0, 1, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 2, '1-5', '��ҩ���г�ҩ����ҩ��Ƭδ�ֱ𿪾ߴ���', '��ҩ���г�ҩ����ҩ��Ƭδ�ֱ𿪾ߴ�����', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 2, '1-6', 'δʹ��ҩƷ�淶���ƿ��ߴ���', 'δʹ��ҩƷ�淶���ƿ��ߴ�����', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 2, '1-7', 'ҩƷ��д���淶�����', 'ҩƷ�ļ����������������λ����д���淶�������', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 2, '1-8', '�÷�����ʹ�ú��������־�', '�÷�������ʹ�á���ҽ�����������á��Ⱥ��������־��', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 2, '1-9', '�����޸�δǩ����ҩƷ����δע��ԭ��', '�����޸�δǩ����ע���޸����ڣ���ҩƷ������ʹ��δע��ԭ����ٴ�ǩ����', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 2, '1-10', 'δд�ٴ���ϻ���д��ȫ', '���ߴ���δд�ٴ���ϻ��ٴ������д��ȫ��', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 2, '1-11', '�����ż��ﴦ��������ҩƷ', '�����ż��ﴦ����������ҩƷ��', 0, 0, 0, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 2, '1-12', '�ӳ���������δע������', '����������£����ﴦ������7�����������ﴦ������3�����������Բ������겡�������������Ҫ�ʵ��ӳ���������δע�����ɵ�', 0, 0, 0, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 2, '1-13', '�����������ҩƷδִ�й��ҹ涨', '��������ҩƷ������ҩƷ��ҽ���ö���ҩƷ��������ҩƷ���������ҩƷ����δִ�й����йع涨��', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 2, '1-14', 'δ������ҩ�������', 'ҽʦδ���տ���ҩ���ٴ�Ӧ�ù���涨���߿���ҩ�ﴦ����', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 2, '1-15', '��ҩ��Ƭδ����������ʹ������', '��ҩ��Ƭ����ҩ��δ���ա�������������ʹ����˳�����У���δ��Ҫ���עҩ����������������Ҫ���', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 2, '2-1', '��Ӧ֤������', '��Ӧ֤�����˵�', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 2, '2-2', '��ѡ��ҩƷ������', '��ѡ��ҩƷ�����˵�', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 2, '2-3', 'ҩƷ���ͻ��ҩ;��������', 'ҩƷ���ͻ��ҩ;�������˵�', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 2, '2-4', '���������ɲ���ѡ���һ���ҩ��', '���������ɲ���ѡ���һ���ҩ���', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 2, '2-5', '�÷�������������', '�÷������������˵�', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 2, '2-6', '������ҩ������', '������ҩ�����˵�', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 2, '2-7', '�ظ���ҩ', '�ظ���ҩ��', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 2, '2-8', '��������ɻ��߲����໥����', '��������ɻ��߲����໥���õ�', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 2, '2-9', '������ҩ������', '������ҩ�����������', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 2, '3-1', '����Ӧ֤��ҩ', '����Ӧ֤��ҩ', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 2, '3-2', '���������ɿ��߸߼�ҩ', '���������ɿ��߸߼�ҩ��', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 2, '3-3', '���������ɳ�˵������ҩ', '���������ɳ�˵������ҩ��', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 2, '3-4', '����������Ϊͬһ���߿�2������������ͬҩ��', '����������Ϊͬһ����ͬʱ����2������ҩ��������ͬҩ���', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 3, 'C01', 'PASS���', '������ҩ�����', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into ���������Ŀ (ID,���,����,���,����,�Ƿ���������,�Ƿ�סԺ����,�������,PASS���,������,����ʱ��) Select ���������Ŀ_ID.Nextval, 4, 'D01', '��ҩע�����������', '��ҩע����������ϣ������֣�', 0, 0, 2, Null, user, sysdate From Dual;

