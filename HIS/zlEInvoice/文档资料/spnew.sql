CREATE TABLE ����Ʊ�����(
	���   number(3),
	����   varchar2(50),
	����   varchar2(20),
	�Ƿ�����   number(2),
	���� varchar2(100),
	������ varchar2(100))
 TABLESPACE zl9Expense;

Alter Table ����Ʊ����� Add Constraint ����Ʊ�����_PK Primary Key(���) Using Index Tablespace zl9Indexhis;
Alter Table ����Ʊ����� Add Constraint ����Ʊ�����_UQ_����  Unique(����)  Using Index Tablespace zl9Indexhis;
 
Create Table ����Ʊ��վ�����(
 ���� Number(2),
 վ�� varchar2(50))
 TABLESPACE zl9Expense;
Alter Table ����Ʊ��վ����� Add Constraint ����Ʊ��վ�����_PK Primary Key(վ��,����) Using Index Tablespace zl9Indexhis;


CREATE TABLE ����Ʊ�ݿ�Ʊ��(
	ID   Number(18),
	�ϼ�ID   Number(18),
	����   varchar2(20),
	����   varchar2(50),
	����   varchar2(20),
	Ժ��   varchar2(50),
	�ͻ���   varchar2(50),
	����ID   number(18),
	λ��   varchar2(100),
	ĩ��   number(2),
	����ʱ��   date,
	����ʱ��   date)
 TABLESPACE zl9Expense;

Alter Table ����Ʊ�ݿ�Ʊ�� Add Constraint ����Ʊ�ݿ�Ʊ��_PK Primary Key(ID) Using Index Tablespace zl9Indexhis;
Alter Table ����Ʊ�ݿ�Ʊ�� Add Constraint ����Ʊ�ݿ�Ʊ��_UQ_����  Unique(����, ����ʱ��)  Using Index Tablespace zl9Indexhis;
Alter Table ����Ʊ�ݿ�Ʊ�� Add Constraint ����Ʊ�ݿ�Ʊ��_FK_�ϼ�ID Foreign Key (�ϼ�ID) References ����Ʊ�ݿ�Ʊ��(ID) on delete cascade;
Alter Table ����Ʊ�ݿ�Ʊ�� Add Constraint ����Ʊ�ݿ�Ʊ��_FK_����ID Foreign Key (����ID) References ���ű�(ID) on delete cascade;

CREATE INDEX ����Ʊ�ݿ�Ʊ��_IX_���� ON ����Ʊ�ݿ�Ʊ��(����) TABLESPACE zl9Indexhis;
CREATE INDEX ����Ʊ�ݿ�Ʊ��_IX_����ID ON ����Ʊ�ݿ�Ʊ��(����ID) TABLESPACE zl9Indexhis;


CREATE SEQUENCE ����Ʊ�ݿ�Ʊ��_ID START WITH 1;  


CREATE TABLE Ʊ�ݿ�Ʊ�����(
    Id Number(18),
	��Ʊ��ID Number(18),
	��ԱID Number(18),
	�ͻ��� varchar2(50))
TABLESPACE zl9Expense;

Alter Table Ʊ�ݿ�Ʊ����� Add Constraint Ʊ�ݿ�Ʊ�����_PK Primary Key(ID) Using Index Tablespace zl9Indexhis;
Alter Table Ʊ�ݿ�Ʊ����� Add Constraint Ʊ�ݿ�Ʊ�����_UQ_��Ʊ��ID  Unique(��Ʊ��ID, ��ԱID,�ͻ���)  Using Index Tablespace zl9Indexhis;
Alter Table Ʊ�ݿ�Ʊ����� Add Constraint Ʊ�ݿ�Ʊ�����_FK_��ԱID Foreign Key (��ԱID) References ��Ա��(ID) on delete cascade;
CREATE INDEX Ʊ�ݿ�Ʊ�����_IX_��ԱID ON Ʊ�ݿ�Ʊ�����(��ԱID) TABLESPACE zl9Indexhis;
CREATE INDEX Ʊ�ݿ�Ʊ�����_IX_�ͻ��� ON Ʊ�ݿ�Ʊ�����(�ͻ���) TABLESPACE zl9Indexhis;
CREATE SEQUENCE Ʊ�ݿ�Ʊ�����_ID START WITH 1;  


ALTER TABLE ����Ԥ����¼ Add  (�Ƿ����Ʊ�� number(2),Ԥ������Ʊ�� number(2));
ALTER TABLE ���˽��ʼ�¼ Add  �Ƿ����Ʊ�� number(2);

ALTER TABLE ��Լ��λ ADD (������ô��� varchar2(50));
ALTER TABLE ������� add(���ջ������� varchar2(50));
CREATE INDEX ��Լ��λ_IX_���� ON ��Լ��λ(����) TABLESPACE zl9Indexhis;

Create Table ����Ʊ��ʹ�ü�¼(
 ID Number(18),
 Ʊ�� number(2),
 ��¼״̬ number(2),
 ����ID number(18),
 ����ID number(18),
 ���� varchar2(100),
 �Ա� varchar2(4),
 ���� varchar2(20),
 ����� number(18),
 סԺ�� number(18),
 ���� Varchar2(50),
 ���� Varchar2(50),
 ������ Varchar2(20),
 ƾ֤���� Varchar2(50),
 ƾ֤���� Varchar2(50),
 ƾ֤������ Varchar2(20),
 Ʊ�ݽ�� number(16,5),
 ����ʱ�� varchar2(30),
 URL����  varchar2(2000),
 URL����  varchar2(2000),
 ԭƱ��ID number(18),
 �Ƿ񻻿� number(2),
 ֽ�ʷ�Ʊ�� Varchar2(50),
 ��ӡID Number(18),
 �˿�id number(18),
 ��ע varchar2(4000),
 ��Ʊ�� varchar2(100),
 ϵͳ��Դ varchar2(100),
 ����Ա��� varchar2(6),
 ����Ա���� varchar2(50),
 �Ǽ�ʱ�� Date,
 ��ת�� number(3))
 TABLESPACE zl9Expense PCTFREE 5 initrans 20;

Alter Table ����Ʊ��ʹ�ü�¼ Add Constraint ����Ʊ��ʹ�ü�¼_PK Primary Key(ID) Using Index Tablespace zl9Indexhis;
Alter table ����Ʊ��ʹ�ü�¼ Add Constraint ����Ʊ��ʹ�ü�¼_UQ_���� Unique(����,Ʊ��,��¼״̬,����)  Using Index Tablespace zl9Indexhis;
Alter Table ����Ʊ��ʹ�ü�¼ Add Constraint ����Ʊ��ʹ�ü�¼_FK_ԭƱ��ID Foreign Key (ԭƱ��ID) References ����Ʊ��ʹ�ü�¼(ID);
Alter Table ����Ʊ��ʹ�ü�¼ Add Constraint ����Ʊ��ʹ�ü�¼_FK_����ID Foreign Key (����ID) References ������Ϣ(����ID);

CREATE INDEX ����Ʊ��ʹ�ü�¼_IX_�Ǽ�ʱ�� ON ����Ʊ��ʹ�ü�¼(�Ǽ�ʱ��) TABLESPACE zl9Indexhis;
CREATE INDEX ����Ʊ��ʹ�ü�¼_IX_����ʱ�� ON ����Ʊ��ʹ�ü�¼(����ʱ��) TABLESPACE zl9Indexhis;
CREATE INDEX ����Ʊ��ʹ�ü�¼_IX_����ID ON ����Ʊ��ʹ�ü�¼(����ID) TABLESPACE zl9Indexhis;
CREATE INDEX ����Ʊ��ʹ�ü�¼_IX_ԭƱ��ID ON ����Ʊ��ʹ�ü�¼(ԭƱ��ID) TABLESPACE zl9Indexhis;

CREATE SEQUENCE ����Ʊ��ʹ�ü�¼_ID START WITH 1;  

CREATE TABLE ����Ʊ�ݶ�ά�� (
 ʹ�ü�¼ID number(18),
 ��ά�� clob,
 ��ת�� number(3)) 
TABLESPACE zl9Expense PCTFREE 20;

ALTER TABLE ����Ʊ�ݶ�ά�� ADD CONSTRAINT ����Ʊ�ݶ�ά��_PK PRIMARY KEY (ʹ�ü�¼ID) USING INDEX TABLESPACE zl9Indexhis;
ALTER TABLE ����Ʊ�ݶ�ά�� ADD CONSTRAINT ����Ʊ�ݶ�ά��_FK_ʹ�ü�¼ID  FOREIGN KEY (ʹ�ü�¼ID ) REFERENCES ����Ʊ��ʹ�ü�¼(ID)  On Delete Cascade;


ALTER TABLE Ʊ������¼ ADD (�Ƿ����� number(2));
ALTER TABLE Ʊ�����ü�¼ ADD (�Ƿ����� number(2));

ALTER TABLE Ʊ��ʹ����ϸ ADD (����Ʊ��ID number(18));
ALTER TABLE Ʊ��ʹ����ϸ ADD CONSTRAINT Ʊ��ʹ����ϸ_FK_����Ʊ��ID FOREIGN KEY(����Ʊ��ID) REFERENCES ����Ʊ��ʹ�ü�¼(ID);
CREATE INDEX Ʊ��ʹ����ϸ_IX_����Ʊ��ID ON Ʊ��ʹ����ϸ(����Ʊ��ID) TABLESPACE zl9Indexhis;

Insert into zlTables(ϵͳ,����,��ռ�,����) Values(&n_System,'����Ʊ�����','ZL9EXPENSE','A2');
Insert into zlTables(ϵͳ,����,��ռ�,����) Values(&n_System,'����Ʊ��վ�����','ZL9EXPENSE','A2');
Insert into zlTables(ϵͳ,����,��ռ�,����) Values(&n_System,'����Ʊ�ݿ�Ʊ��','ZL9EXPENSE','A2');
Insert into zlTables(ϵͳ,����,��ռ�,����) Values(&n_System,'Ʊ�ݿ�Ʊ�����','ZL9EXPENSE','A2');
Insert into zlTables(ϵͳ,����,��ռ�,����) Values(&n_System,'����Ʊ��ʹ�ü�¼','ZL9EXPENSE','B1');
Insert into zlTables(ϵͳ,����,��ռ�,����) Values(&n_System,'����Ʊ�ݶ�ά��','ZL9EXPENSE','B1');
Insert Into zlBakTables(ϵͳ,���,����,���,ֱ��ת��,ͣ�ô�����)
select &n_system,1,A.* From (
Select ����,���,ֱ��ת��,ͣ�ô����� From zlBakTables Where 1 = 0
Union All Select '����Ʊ��ʹ�ü�¼',18,1,-NULL From Dual
Union All Select '����Ʊ�ݶ�ά��',19,1,-NULL From Dual) A;

Insert Into zlParameters
(ID,ϵͳ,ģ��,˽��,����,��Ȩ,�̶�,����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵��)
Select Zlparameters_Id.Nextval, &n_System, -1*null, 0, 0, 0, 0, 0, 0,327 , '�Һŵ���Ʊ�ݿ���', '', '0|1|0:','��Ҫ���ƹҺ�ҵ���Ƿ����õ���Ʊ��'||CHR(13)||'1.�����˵���Ʊ�ݵ�ҵ��:ϵͳ�������ٰ������Ʊ�ݹ�����ϵ����Ʊ�ݹ���Ϳ��ƣ���ͨ��������������Ʊ�ݽӿڲ������е���Ʊ�ݵĿ��ߡ����ϡ�������,��ˣ��ͻ��˿ؼ�����Ҫ�в�����zlElectronicInvoice.dll�������������ص���Ʊ�ݽӿڵġ�'||CHR(13)||'2.δ���õ���Ʊ�ݵ�ҵ��:Ʊ�ݡ���ӡ�ȶ���HISϵͳ���й���Ϳ��ơ�',
'��ʽ:Ʊ�����ÿ���|Ʊ�ݹ������|ҽ�����ÿ���'||CHR(13)||'1. Ʊ�����ÿ���:��Ҫ���Ƶ���Ʊ���Ƿ����÷�ʽ��0-��ʾδ���õ���Ʊ��;1-�������õ���Ʊ��;2-�����վ�����õ���Ʊ��  '||CHR(13)||'2.Ʊ�ݹ������:��Ҫ�ǿ����Ƿ�HISϵͳ����Ʊ��:0-����HIS����Ʊ��;1-����Ʊ��ƽ̨'||chr(13)||'3.ҽ�����ÿ���:��ʽΪ���ñ�־:��������'||CHR(13)||'   a.���ñ�־:0-����δ����;1-��������'||CHR(13)||'   b.��������:�մ�������ҽ������;�ǿ�ʱ������ҽ����ţ����ҽ��ʱ�ö��ŷ���', '', '������ĳЩҽԺ��Ҫ���õ���Ʊ�ݹ���ҵ��', '��ҽԺ���õ���Ʊ�ݺ�һ�㲻�����˲�������������˲���������Ӱ�쵽�Һ�ҵ���Ʊ��ʹ�ü���ӡ��'
From Dual;

Insert Into zlParameters
(ID,ϵͳ,ģ��,˽��,����,��Ȩ,�̶�,����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵��)
Select Zlparameters_Id.Nextval, &n_System, -1*null, 0, 0, 0, 0, 0, 0,328 , '�շѵ���Ʊ�ݿ���', '', '0|1|0:','��Ҫ�����շ�ҵ���Ƿ����õ���Ʊ��'||CHR(13)||'1.�����˵���Ʊ�ݵ�ҵ��:ϵͳ�������ٰ������Ʊ�ݹ�����ϵ����Ʊ�ݹ���Ϳ��ƣ���ͨ��������������Ʊ�ݽӿڲ������е���Ʊ�ݵĿ��ߡ����ϡ�������,��ˣ��ͻ��˿ؼ�����Ҫ�в�����zlElectronicInvoice.dll�������������ص���Ʊ�ݽӿڵġ�'||CHR(13)||'2.δ���õ���Ʊ�ݵ�ҵ��:Ʊ�ݡ���ӡ�ȶ���HISϵͳ���й���Ϳ��ơ�',
'��ʽ:Ʊ�����ÿ���|Ʊ�ݹ������|ҽ�����ÿ���'||CHR(13)||'1. Ʊ�����ÿ���:��Ҫ���Ƶ���Ʊ���Ƿ����÷�ʽ��0-��ʾδ���õ���Ʊ��;1-�������õ���Ʊ��;2-�����վ�����õ���Ʊ��  '||CHR(13)||'2.Ʊ�ݹ������:��Ҫ�ǿ����Ƿ�HISϵͳ����Ʊ��:0-����HIS����Ʊ��;1-����Ʊ��ƽ̨'||chr(13)||'3.ҽ�����ÿ���:��ʽΪ���ñ�־:��������'||CHR(13)||'   a.���ñ�־:0-����δ����;1-��������'||CHR(13)||'   b.��������:�մ�������ҽ������;�ǿ�ʱ������ҽ����ţ����ҽ��ʱ�ö��ŷ���', '', '������ĳЩҽԺ��Ҫ���õ���Ʊ�ݹ���ҵ��', '��ҽԺ���õ���Ʊ�ݺ�һ�㲻�����˲�������������˲���������Ӱ�쵽�շ�ҵ���Ʊ��ʹ�ü���ӡ��'
From Dual;

Insert Into zlParameters
(ID,ϵͳ,ģ��,˽��,����,��Ȩ,�̶�,����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵��)
Select Zlparameters_Id.Nextval, &n_System, -1*null, 0, 0, 0, 0, 0, 0,329 , 'Ԥ������Ʊ�ݿ���', '', '0|0|1|0:','��Ҫ����Ԥ��ҵ���Ƿ����õ���Ʊ��'||CHR(13)||'1.�����˵���Ʊ�ݵ�ҵ��:ϵͳ�������ٰ������Ʊ�ݹ�����ϵ����Ʊ�ݹ���Ϳ��ƣ���ͨ��������������Ʊ�ݽӿڲ������е���Ʊ�ݵĿ��ߡ����ϡ�������,��ˣ��ͻ��˿ؼ�����Ҫ�в�����zlElectronicInvoice.dll�������������ص���Ʊ�ݽӿڵġ�'||CHR(13)||'2.δ���õ���Ʊ�ݵ�ҵ��:Ʊ�ݡ���ӡ�ȶ���HISϵͳ���й���Ϳ��ơ�',
'��ʽ:Ԥ�����|Ʊ�����ÿ���|Ʊ�ݹ������|ҽ�����ÿ���'||CHR(13)||'1. Ԥ�����:��Ҫ�������õ���Ʊ�ݵ�Ԥ�����ͣ�0-��ʾ����Ԥ��;1-��������Ԥ��;2-����סԺԤ��  '||CHR(13)||'2. Ʊ�����ÿ���:��Ҫ���Ƶ���Ʊ���Ƿ����÷�ʽ��0-��ʾδ���õ���Ʊ��;1-�������õ���Ʊ��;2-�����վ�����õ���Ʊ��  '||CHR(13)||'3.Ʊ�ݹ������:��Ҫ�ǿ����Ƿ�HISϵͳ����Ʊ��:0-����HIS����Ʊ��;1-����Ʊ��ƽ̨'||chr(13)||'4.ҽ�����ÿ���:��ʽΪ���ñ�־:��������'||CHR(13)||'   a.���ñ�־:0-����δ����;1-��������'||CHR(13)||'   b.��������:�մ�������ҽ������;�ǿ�ʱ������ҽ����ţ����ҽ��ʱ�ö��ŷ���', '', '������ĳЩҽԺ��Ҫ���õ���Ʊ�ݹ���ҵ��', '��ҽԺ���õ���Ʊ�ݺ�һ�㲻�����˲�������������˲���������Ӱ�쵽Ԥ��ҵ���Ʊ��ʹ�ü���ӡ��'
From Dual;

Insert Into zlParameters
(ID,ϵͳ,ģ��,˽��,����,��Ȩ,�̶�,����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵��)
Select Zlparameters_Id.Nextval, &n_System, -1*null, 0, 0, 0, 0, 0, 0,330 , '���ʵ���Ʊ�ݿ���', '', '0|1|0:','��Ҫ���ƽ���ҵ���Ƿ����õ���Ʊ��'||CHR(13)||'1.�����˵���Ʊ�ݵ�ҵ��:ϵͳ�������ٰ������Ʊ�ݹ�����ϵ����Ʊ�ݹ���Ϳ��ƣ���ͨ��������������Ʊ�ݽӿڲ������е���Ʊ�ݵĿ��ߡ����ϡ�������,��ˣ��ͻ��˿ؼ�����Ҫ�в�����zlElectronicInvoice.dll�������������ص���Ʊ�ݽӿڵġ�'||CHR(13)||'2.δ���õ���Ʊ�ݵ�ҵ��:Ʊ�ݡ���ӡ�ȶ���HISϵͳ���й���Ϳ��ơ�',
'��ʽ:Ʊ�����ÿ���|Ʊ�ݹ������|ҽ�����ÿ���'||CHR(13)||'1. Ʊ�����ÿ���:��Ҫ���Ƶ���Ʊ���Ƿ����÷�ʽ��0-��ʾδ���õ���Ʊ��;1-�������õ���Ʊ��;2-�����վ�����õ���Ʊ��  '||CHR(13)||'2.Ʊ�ݹ������:��Ҫ�ǿ����Ƿ�HISϵͳ����Ʊ��:0-����HIS����Ʊ��;1-����Ʊ��ƽ̨'||chr(13)||'3.ҽ�����ÿ���:��ʽΪ���ñ�־:��������'||CHR(13)||'   a.���ñ�־:0-����δ����;1-��������'||CHR(13)||'   b.��������:�մ�������ҽ������;�ǿ�ʱ������ҽ����ţ����ҽ��ʱ�ö��ŷ���', '', '������ĳЩҽԺ��Ҫ���õ���Ʊ�ݹ���ҵ��', '��ҽԺ���õ���Ʊ�ݺ�һ�㲻�����˲�������������˲���������Ӱ�쵽����ҵ���Ʊ��ʹ�ü���ӡ��'
From Dual;

Insert Into zlParameters
(ID,ϵͳ,ģ��,˽��,����,��Ȩ,�̶�,����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵��)
Select Zlparameters_Id.Nextval, &n_System, -1*null, 0, 0, 0, 0, 0, 0,331 , '���￨����Ʊ�ݿ���', '', '0|1|0:','��Ҫ���Ʒ���ҵ���Ƿ����õ���Ʊ��'||CHR(13)||'1.�����˵���Ʊ�ݵ�ҵ��:ϵͳ�������ٰ������Ʊ�ݹ�����ϵ����Ʊ�ݹ���Ϳ��ƣ���ͨ��������������Ʊ�ݽӿڲ������е���Ʊ�ݵĿ��ߡ����ϡ�������,��ˣ��ͻ��˿ؼ�����Ҫ�в�����zlElectronicInvoice.dll�������������ص���Ʊ�ݽӿڵġ�'||CHR(13)||'2.δ���õ���Ʊ�ݵ�ҵ��:Ʊ�ݡ���ӡ�ȶ���HISϵͳ���й���Ϳ��ơ�',
'��ʽ:Ʊ�����ÿ���|Ʊ�ݹ������|ҽ�����ÿ���'||CHR(13)||'1. Ʊ�����ÿ���:��Ҫ���Ƶ���Ʊ���Ƿ����÷�ʽ��0-��ʾδ���õ���Ʊ��;1-�������õ���Ʊ��;2-�����վ�����õ���Ʊ��  '||CHR(13)||'2.Ʊ�ݹ������:��Ҫ�ǿ����Ƿ�HISϵͳ����Ʊ��:0-����HIS����Ʊ��;1-����Ʊ��ƽ̨'||chr(13)||'3.ҽ�����ÿ���:��ʽΪ���ñ�־:��������'||CHR(13)||'   a.���ñ�־:0-����δ����;1-��������'||CHR(13)||'   b.��������:�մ�������ҽ������;�ǿ�ʱ������ҽ����ţ����ҽ��ʱ�ö��ŷ���', '', '������ĳЩҽԺ��Ҫ���õ���Ʊ�ݹ���ҵ��', '��ҽԺ���õ���Ʊ�ݺ�һ�㲻�����˲�������������˲���������Ӱ�쵽����ҵ���Ʊ��ʹ�ü���ӡ��'
From Dual;


Insert Into zlPrograms(���,����,˵��,ϵͳ,����) Values(1145,'����Ʊ�ݲ���','��Ҫ����Ҫ����ԹҺš��շѡ�Ԥ�������ʵ�ҵ��ĵ���Ʊ�ݵĿ�Ʊ����ӡ����Ʊ����Ʊ�Ȳ������д�Ȩ��ʱ��������Ե���Ʊ�ݵĿ��ߡ���ӡ����Ʊ����Ʊ�Ĳ�����',&n_System,'zL9CashBill');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��,ȱʡֵ)
Select &n_System,1145,A.* From (
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0
Union All Select '����',1,'',1 From Dual
Union All Select '��������',2,'��Բ������в�����Ȩ�ޡ��и�Ȩ��ʱ��������б��ز�������',0 From Dual
Union All Select '���ߵ���Ʊ��',3,'��Ҫ�ǿ����Ƿ������ߵ���Ʊ��Ȩ��',0 From Dual
Union All Select '����ֽ��Ʊ��',4,'��Ҫ�ǿ����Ƿ񻻿�ֽ��Ʊ��Ȩ��.',0 From Dual
Union All Select '���»���Ʊ��',5,'��Ҫ�ǿ����Ƿ����»���ֽ��Ʊ��Ȩ��.',0 From Dual
Union All Select '����ֽ��Ʊ��',6,'��Ҫ�ǿ����Ƿ����������ѻ�����ֽ��Ʊ��.',0 From Dual 
) A;
  
Insert Into zlModuleRelas(���ϵͳ,ģ��,����,ϵͳ,���ģ��,�������,��ع���,ȱʡֵ)
Select  &n_System,1145,A.* From (
Select ����,���ϵͳ,���ģ��,�������,��ع���,ȱʡֵ From zlModuleRelas Where 1 = 0
Union All Select '����',&n_System,1101,1,'����',1 From Dual
Union All Select '����',&n_System,1103,1,'����',1 From Dual
Union All Select '����',&n_System,1107,1,'����',1 From Dual
Union All Select '����',&n_System,1151,1,'����',1 From Dual
Union All Select '����',&n_System,1111,1,'����',1 From Dual
Union All Select '����',&n_System,1113,1,'����',1 From Dual
Union All Select '����',&n_System,1121,1,'����',1 From Dual
Union All Select '����',&n_System,1124,1,'����',1 From Dual
Union All Select '����',&n_System,1131,1,'����',1 From Dual
Union All Select '����',&n_System,1137,1,'����',1 From Dual
Union All Select '����',&n_System,1801,1,'����',1 From Dual
Union All Select '����',&n_System,1802,1,'����',1 From Dual
Union All Select '����',&n_System,1803,1,'����',1 From Dual
Union All Select '����',&n_System,1804,1,'����',1 From Dual
Union All Select '����',&n_System,1805,1,'����',1 From Dual
Union All Select '����',&n_System,1806,1,'����',1 From Dual
Union All Select '����',&n_System,1807,1,'����',1 From Dual
Union All Select '����',&n_System,1809,1,'����',1 From Dual
Union All Select '����',&n_System,1811,1,'����',1 From Dual
) A;

Insert Into zlProgPrivs(ϵͳ, ���, ����, ������, ����, Ȩ��)
Select &n_System, 1145, '����', User, A.*
From (Select ����, Ȩ�� From zlProgPrivs Where 1 = 0
Union All Select '����Ʊ��վ�����','SELECT' From Dual
Union All Select '����Ʊ��ʹ�ü�¼','SELECT' From Dual
Union All Select 'Ʊ������¼','SELECT' From Dual
Union All Select 'Ʊ�����ü�¼','SELECT' From Dual
Union All Select 'Ʊ��ʹ����ϸ','SELECT' From Dual
Union All Select 'Ʊ��ʹ�����','SELECT' From Dual
Union All Select 'Ʊ�ݴ�ӡ����','SELECT' From Dual
Union All Select '����Ʊ�ݶ�ά��','SELECT' From Dual
Union All Select '����Ʊ��ʹ�ü�¼_ID','SELECT' From Dual
Union All Select '����Ԥ����¼','SELECT' From Dual
Union All Select '������ü�¼','SELECT' From Dual
Union All Select '���ò����¼','SELECT' From Dual
Union All Select 'סԺ���ü�¼','SELECT' From Dual
Union All Select '���˹Һż�¼','SELECT' From Dual
Union All Select '���ս����¼','SELECT' From Dual
Union All Select '���˽��ʼ�¼','SELECT' From Dual
Union All Select '���˿������¼','SELECT' From Dual
Union All Select '�������㽻��','SELECT' From Dual
Union All Select '�����˿���Ϣ','SELECT' From Dual 
Union All Select '���ս�����ϸ','SELECT' From Dual
Union All Select '�������','SELECT' From Dual
Union All Select '������׼��Ŀ','SELECT' From Dual
Union All Select '����֧������','SELECT' From Dual
Union All Select '����֧����Ŀ','SELECT' From Dual
Union All Select '���൵�α���','SELECT' From Dual
Union All Select '�ʻ������Ϣ','SELECT' From Dual
Union All Select '�������Ҷ�Ӧ','SELECT' From Dual
Union All Select '�������','SELECT' From Dual
Union All Select '���㷽ʽ','SELECT' From Dual
Union All Select '���㷽ʽӦ��','SELECT' From Dual
Union All Select '��Ա��','SELECT' From Dual
Union All Select '���ű�','SELECT' From Dual
Union All Select '��������˵��','SELECT' From Dual
Union All Select '�շѷ���Ŀ¼','SELECT' From Dual
Union All Select '�շ��ض���Ŀ','SELECT' From Dual
Union All Select '�շ�ϸĿ','SELECT' From Dual
Union All Select '�շ���Ŀ����','SELECT' From Dual
Union All Select '�շ���Ŀ���','SELECT' From Dual
Union All Select '�շ���ĿĿ¼','SELECT' From Dual
Union All Select '�շ�ִ�п���','SELECT' From Dual
Union All Select '�վݷ�Ŀ','SELECT' From Dual
Union All Select '������Ŀ','SELECT' From Dual
Union All Select '�Ա�','SELECT' From Dual
Union All Select '�ѱ�','SELECT' From Dual
Union All Select '�ѱ����ÿ���','SELECT' From Dual
Union All Select '��������','SELECT' From Dual
Union All Select 'ҩƷ���','SELECT' From Dual
Union All Select 'ҩƷĿ¼','SELECT' From Dual
Union All Select 'ҩƷ����','SELECT' From Dual
Union All Select 'ҩƷ��Ϣ','SELECT' From Dual
Union All Select 'ҽ���������','SELECT' From Dual
Union All Select 'ҽ��������ϸ','SELECT' From Dual
Union All Select 'ҽ���˶Ա�','SELECT' From Dual
Union All Select 'ҽ�Ƹ��ʽ','SELECT' From Dual
Union All Select 'ҽ�ƿ���ʧ��ʽ','SELECT' From Dual
Union All Select '���Ʒ���Ŀ¼','SELECT' From Dual
Union All Select '���ƻ�����Ŀ','SELECT' From Dual
Union All Select '�����շѹ�ϵ','SELECT' From Dual
Union All Select '������ĿĿ¼','SELECT' From Dual
Union All Select '����ִ�п���','SELECT' From Dual
Union All Select '֤������','SELECT' From Dual
Union All Select '��������','SELECT' From Dual
Union All Select '���ѿ�����','SELECT' From Dual
Union All Select '���ѿ����Ŀ¼','SELECT' From Dual
Union All Select '���ѿ���Ϣ','SELECT' From Dual
Union All Select '�����ӿ�����','SELECT' From Dual
Union All Select '����Ʊ�ݿ�Ʊ��','SELECT' From Dual
Union All Select 'Ʊ�ݿ�Ʊ�����','SELECT' From Dual
Union All Select '����Ʊ�����','SELECT' From Dual
UNION ALL SELECT 'Zl_����Ʊ��ʹ�ü�¼_Insert','EXECUTE' From dual
UNION ALL SELECT 'Zl_����Ʊ��ʹ�ü�¼_Delete','EXECUTE' From dual
UNION ALL SELECT 'Zl_����Ʊ��ʹ�ü�¼_Update','EXECUTE' From dual
UNION ALL SELECT 'Zl_����Ʊ�ݶ�ά��_Update','EXECUTE' From dual
UNION ALL SELECT '����Ʊ�ݶ�ά��','UPDATE' From dual
UNION ALL SELECT 'Zl_ֽ��Ʊ��ʹ��_Update','EXECUTE' From dual
UNION ALL SELECT 'Zl_����Ʊ��վ�����_Update','EXECUTE' From dual
UNION ALL SELECT 'zl_����Ʊ�ݿ�Ʊ��_insert','EXECUTE' From dual
UNION ALL SELECT 'zl_����Ʊ�ݿ�Ʊ��_update','EXECUTE' From dual
UNION ALL SELECT 'zl_����Ʊ�ݿ�Ʊ��_start','EXECUTE' From dual
UNION ALL SELECT 'zl_����Ʊ�ݿ�Ʊ��_stop','EXECUTE' From dual
UNION ALL SELECT 'zl_����Ʊ�ݿ�Ʊ��_delete','EXECUTE' From dual
UNION ALL SELECT 'Zl_Ʊ�ݿ�Ʊ�����_Update','EXECUTE' From dual
UNION ALL SELECT 'Zl_�����ӿ�����_Set','EXECUTE' From dual
UNION ALL SELECT 'Zl_�����ӿ�����_Get','EXECUTE' From dual
) A;

Insert Into zlParameters
(ID,ϵͳ,ģ��,˽��,����,��Ȩ,�̶�,����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵��)
Select Zlparameters_Id.Nextval, &n_System, 1145, 0, 1,1, 0, 0, 0,1 , 'Ʊ�ݻ�����ʽ', '', '0','��Ҫ�����ڿ��ߵ���Ʊ�ݺ����ֽ��Ʊ�ݵĻ�������,�����֣����������Զ���������ʾ����.',
'0-��������1-�Զ�������2-��ʾ����.', '1.��ԹҺ�ҵ����Ҫ���á��Һŵ���Ʊ�ݿ��ơ�Ϊ���ò���Ч'||CHR(13)||'2.����շ�ҵ����Ҫ���á��շѵ���Ʊ�ݿ��ơ�Ϊ���ò���Ч'||CHR(13)||'3.���Ԥ��ҵ����Ҫ���á�Ԥ������Ʊ�ݿ��ơ�Ϊ���ò���Ч'||CHR(13)||'4.��Խ���ҵ����Ҫ���á����ʵ���Ʊ�ݿ��ơ�Ϊ���ò���Ч', '������ĳЩҽԺ��Ҫ���õ���Ʊ�ݹ���ҵ��ʱ��ͬ����Ҫ��������Ʊ��ҵ��(��Ҫ�ǹ�����ʹ��).', ''
From Dual;

Insert Into zlParameters
(ID,ϵͳ,ģ��,˽��,����,��Ȩ,�̶�,����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵��)
Select Zlparameters_Id.Nextval, &n_System, 1145, 0, 1,1, 0, 0, 0,2 , '��֪����ӡ��ʽ', '', '0','��Ҫ�����ڿ��ߵ���Ʊ�ݺ��Ƿ��ӡ��֪��������,�����֣�����ӡ����ӡ����ʾ��ӡ.',
'0-����ӡ��1-�Զ���ӡ��2-��ʾ��ӡ..', '�����˵���Ʊ��ҵ��(�Һţ��շѣ�Ԥ��������)�����󣬱�������Ч�����ñ���:zl1_INSIDE_1145', '������ĳЩҽԺ��Ҫ���õ���Ʊ�ݹ���ҵ��ʱ��ͬ����Ҫ��ӡ��֪��������.', ''
From Dual;

Insert Into zlParameters
(ID,ϵͳ,ģ��,˽��,����,��Ȩ,�̶�,����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵��)
Select Zlparameters_Id.Nextval, &n_System, 1145, 0, 1,1, 0, 0, 0,3 , '����Ʊ�ݴ�ӡ��ʽ', '', '0','��Ҫ�����ڿ��ߵ���Ʊ�ݺ��Ƿ��ӡ����Ʊ�ݸ�����,�����֣�����ӡ����ӡ����ʾ��ӡ.',
'0-����ӡ��1-�Զ���ӡ��2-��ʾ��ӡ..', '�����˵���Ʊ��ҵ��(�Һţ��շѣ�Ԥ��������)�����󣬱�������Ч�����ô�ӡ�ӿڽ��д�ӡ', '������ĳЩҽԺ��Ҫ���õ���Ʊ�ݹ���ҵ��ʱ��ͬ����Ҫ��ӡ����Ʊ�ݸ�����.', ''
From Dual;

Insert Into zlParameters
(ID,ϵͳ,ģ��,˽��,����,��Ȩ,�̶�,����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵��)
Select Zlparameters_Id.Nextval, &n_System, 1145, 0, 0,1, 0, 0, 0,4 , '��Ʊ����뷽ʽ', '', '1','��Ҫ���Ƶ���Ʊ�ݿ�Ʊ����뷽ʽ��0-���ͻ��˶��룬1-���շ�Ա���룬2-���ͻ��˺��շ�Ա����.',
'0-���ͻ��˶��룬1-���շ�Ա���룬2-���ͻ��˺��շ�Ա����', '�����˵���Ʊ��ҵ��(�Һţ��շѣ�Ԥ��������)�����󣬱�������Ч', '������ĳЩҽԺ��Ҫ���õ���Ʊ�ݹ���ҵ��ʱ��ͬ����Ҫ��ӡ����Ʊ�ݸ�����.', ''
From Dual;

Insert Into zlPrograms(���,����,˵��,ϵͳ,����) Values(1144,'����Ʊ�ݹ���','��Ҫ����Ե���Ʊ�ݵ���ػ�����Ŀ�Ķ��롢ֽ��Ʊ���·������ʼ�����Ʊ�ݿ��ߵȹ��ܵĹ���',&n_System,'zL9CashBill');
Insert Into Zlmenus
  (���, Id, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��)
  Select 'ȱʡ', Zlmenus_Id.Nextval, Id, '����Ʊ�ݹ���', '����Ʊ��', 'E', 246, '��Ҫ����Ե���Ʊ�ݵ���ػ�����Ŀ�Ķ��롢ֽ��Ʊ���·������ʼ�����Ʊ�ݿ��ߵȹ��ܵĹ���', &n_System, 1144
  From Zlmenus
  Where ���� = '��Ӫ����ϵͳ' And ��� = 'ȱʡ' And ϵͳ = &n_System And ģ�� Is Null AND ROWNUM <2;

Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��,ȱʡֵ)
Select &n_System,1144,A.* From (
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0
Union All Select '����',1,'',1 From Dual
Union All Select '�������ݹ���',2,'��Ҫ���ƻ������ݵ�һЩά�������磺վ����롢�վݷ�Ŀ����Ȼ���������.',0 From Dual
Union All Select '����Ʊ�ݺ˶�',3,'��Ҫ����Ե���Ʊ�ݵĺ˶Բ���',0 From Dual
Union All Select '�Һ�Ʊ�ݺ˶�',4,'��Ҫ����ԹҺ�ҵ�����ĵ���Ʊ�ݺ˶�.',0 From Dual
Union All Select '�շ�Ʊ�ݺ˶�',5,'��Ҫ������շ�ҵ�����ĵ���Ʊ�ݺ˶�.',0 From Dual
Union All Select '����Ԥ��Ʊ�ݺ˶�',6,'��Ҫ���������Ԥ��ҵ�����ĵ���Ʊ�ݺ˶�.',0 From Dual
Union All Select 'סԺԤ��Ʊ�ݺ˶�',7,'��Ҫ�����סԺԤ��ҵ�����ĵ���Ʊ�ݺ˶�.',0 From Dual
Union All Select '�������Ʊ�ݺ˶�',8,'��Ҫ������������ҵ�����ĵ���Ʊ�ݺ˶�.',0 From Dual
Union All Select 'סԺ����Ʊ�ݺ˶�',9,'��Ҫ�����סԺ����ҵ�����ĵ���Ʊ�ݺ˶�.',0 From Dual
Union All Select '����Ʊ�ݹ���',10,'��Ҫ�����δ���ߵĵ���Ʊ�ݽ����������߲���.',0 From Dual
Union All Select '���߹Һŵ���Ʊ��',11,'��Ҫ����ԹҺ�δ���ߵ���Ʊ�ݵļ�¼���е���Ʊ�ݵĿ���.',0 From Dual
Union All Select '�����շѵ���Ʊ��',12,'��Ҫ����������շ�δ���ߵ���Ʊ�ݵļ�¼���е���Ʊ�ݵĿ���.',0 From Dual
Union All Select '��������Ԥ������Ʊ��',13,'��Ҫ���������Ԥ��δ���ߵ���Ʊ�ݵļ�¼���е���Ʊ�ݵĿ���.',0 From Dual
Union All Select '����סԺԤ������Ʊ��',14,'��Ҫ�����סԺԤ��δ���ߵ���Ʊ�ݵļ�¼���е���Ʊ�ݵĿ���.',0 From Dual
Union All Select '����������ʵ���Ʊ��',15,'��Ҫ������������δ���ߵ���Ʊ�ݵļ�¼���е���Ʊ�ݵĿ���.',0 From Dual
Union All Select '����סԺ���ʵ���Ʊ��',16,'��Ҫ�����סԺ����δ���ߵ���Ʊ�ݵļ�¼���е���Ʊ�ݵĿ���.',0 From Dual
Union All Select 'ֽ��Ʊ�ݹ���',17,'��Ҫ�����δ����ֽ��Ʊ�ݵĵ���Ʊ�ݽ��л�������.',0 From Dual
Union All Select '�����Һ�Ʊ��',18,'��Ҫ����ԹҺ�δ����ֽ��Ʊ�ݵĵ���Ʊ�ݼ�¼���л���.',0 From Dual
Union All Select '�����շ�Ʊ��',19,'��Ҫ������շ�δ����ֽ��Ʊ�ݵĵ���Ʊ�ݼ�¼���л���.',0 From Dual
Union All Select '��������Ԥ��Ʊ��',20,'��Ҫ���������Ԥ��δ����ֽ��Ʊ�ݵĵ���Ʊ�ݼ�¼���л���.',0 From Dual
Union All Select '����סԺԤ��Ʊ��',21,'��Ҫ�����סԺԤ��δ����ֽ��Ʊ�ݵĵ���Ʊ�ݼ�¼���л���.',0 From Dual
Union All Select '�����������Ʊ��',22,'��Ҫ������������δ����ֽ��Ʊ�ݵĵ���Ʊ�ݼ�¼���л���.',0 From Dual
Union All Select '����סԺ����Ʊ��',23,'��Ҫ�����סԺ����δ����ֽ��Ʊ�ݵĵ���Ʊ�ݼ�¼���л���.',0 From Dual
) A;

--����Ʊ�ݺ˶�
Insert Into zlProgRelas(ϵͳ,���,���,����,��ϵ,����,�����ϵ)
Select &n_System,1144,1,A.* From (
Select ����,��ϵ,����,�����ϵ From zlProgRelas Where 1 = 0
Union All Select '����Ʊ�ݺ˶�',2,1,1 From Dual
Union All Select '�Һ�Ʊ�ݺ˶�',2,0,0 From Dual
Union All Select '�շ�Ʊ�ݺ˶�',2,0,0 From Dual
Union All Select '����Ԥ��Ʊ�ݺ˶�',2,0,0 From Dual
Union All Select 'סԺԤ��Ʊ�ݺ˶�',2,0,0 From Dual
Union All Select '�������Ʊ�ݺ˶�',2,0,0 From Dual
Union All Select 'סԺ����Ʊ�ݺ˶�',2,0,0 From Dual) A;

--����Ʊ�ݹ���
Insert Into zlProgRelas(ϵͳ,���,���,����,��ϵ,����,�����ϵ)
Select &n_System,1144,2,A.* From (
Select ����,��ϵ,����,�����ϵ From zlProgRelas Where 1 = 0
Union All Select '����Ʊ�ݹ���',2,1,1 From Dual
Union All Select '���߹Һŵ���Ʊ��',2,0,0 From Dual
Union All Select '�����շѵ���Ʊ��',2,0,0 From Dual
Union All Select '��������Ԥ������Ʊ��',2,0,0 From Dual
Union All Select '����סԺԤ������Ʊ��',2,0,0 From Dual
Union All Select '����������ʵ���Ʊ��',2,0,0 From Dual
Union All Select '����סԺ���ʵ���Ʊ��',2,0,0 From Dual) A;

--ֽ��Ʊ�ݹ���
Insert Into zlProgRelas(ϵͳ,���,���,����,��ϵ,����,�����ϵ)
Select &n_System,1144,3,A.* From (
Select ����,��ϵ,����,�����ϵ From zlProgRelas Where 1 = 0
Union All Select 'ֽ��Ʊ�ݹ���',2,1,1 From Dual
Union All Select '�����Һ�Ʊ��',2,0,0 From Dual
Union All Select '�����շ�Ʊ��',2,0,0 From Dual
Union All Select '��������Ԥ��Ʊ��',2,0,0 From Dual
Union All Select '����סԺԤ��Ʊ��',2,0,0 From Dual
Union All Select '�����������Ʊ��',2,0,0 From Dual
Union All Select '����סԺ����Ʊ��',2,0,0 From Dual) A;



Insert Into zlProgPrivs(ϵͳ, ���, ����, ������, ����, Ȩ��)
Select &n_System, 1144, '����', User, A.*
From (Select ����, Ȩ�� From zlProgPrivs Where 1 = 0
Union All Select '����Ʊ��ʹ�ü�¼','SELECT' From Dual
Union All Select '����Ʊ��վ�����','SELECT' From Dual
Union All Select 'Ʊ������¼','SELECT' From Dual
Union All Select 'Ʊ�����ü�¼','SELECT' From Dual
Union All Select 'Ʊ��ʹ����ϸ','SELECT' From Dual
Union All Select 'Ʊ��ʹ�����','SELECT' From Dual
Union All Select 'Ʊ�ݴ�ӡ����','SELECT' From Dual
Union All Select '����Ʊ�ݶ�ά��','SELECT' From Dual
Union All Select '����Ʊ��ʹ�ü�¼_ID','SELECT' From Dual
Union All Select '����Ʊ�ݿ�Ʊ��_ID','SELECT' From Dual
Union All Select '����Ԥ����¼','SELECT' From Dual
Union All Select '������ü�¼','SELECT' From Dual
Union All Select '���ò����¼','SELECT' From Dual
Union All Select 'סԺ���ü�¼','SELECT' From Dual
Union All Select '���˹Һż�¼','SELECT' From Dual
Union All Select '���ս����¼','SELECT' From Dual
Union All Select '���˽��ʼ�¼','SELECT' From Dual
Union All Select '���˿������¼','SELECT' From Dual
Union All Select '�������㽻��','SELECT' From Dual
Union All Select '�����˿���Ϣ','SELECT' From Dual 
Union All Select '���ս�����ϸ','SELECT' From Dual
Union All Select '�������','SELECT' From Dual
Union All Select '������׼��Ŀ','SELECT' From Dual
Union All Select '����֧������','SELECT' From Dual
Union All Select '����֧����Ŀ','SELECT' From Dual
Union All Select '���൵�α���','SELECT' From Dual
Union All Select '�ʻ������Ϣ','SELECT' From Dual
Union All Select '�������Ҷ�Ӧ','SELECT' From Dual
Union All Select '�������','SELECT' From Dual
Union All Select '���㷽ʽ','SELECT' From Dual
Union All Select '���㷽ʽӦ��','SELECT' From Dual
Union All Select '��Ա��','SELECT' From Dual
Union All Select '���ű�','SELECT' From Dual
Union All Select '��������˵��','SELECT' From Dual
Union All Select '�շѷ���Ŀ¼','SELECT' From Dual
Union All Select '�շ��ض���Ŀ','SELECT' From Dual
Union All Select '�շ�ϸĿ','SELECT' From Dual
Union All Select '�շ���Ŀ����','SELECT' From Dual
Union All Select '�շ���Ŀ���','SELECT' From Dual
Union All Select '�շ���ĿĿ¼','SELECT' From Dual
Union All Select '�շ�ִ�п���','SELECT' From Dual
Union All Select '�վݷ�Ŀ','SELECT' From Dual
Union All Select '������Ŀ','SELECT' From Dual
Union All Select '�Ա�','SELECT' From Dual
Union All Select '�ѱ�','SELECT' From Dual
Union All Select '�ѱ����ÿ���','SELECT' From Dual
Union All Select '��������','SELECT' From Dual
Union All Select 'ҩƷ���','SELECT' From Dual
Union All Select 'ҩƷĿ¼','SELECT' From Dual
Union All Select 'ҩƷ����','SELECT' From Dual
Union All Select 'ҩƷ��Ϣ','SELECT' From Dual
Union All Select 'ҽ���������','SELECT' From Dual
Union All Select 'ҽ��������ϸ','SELECT' From Dual
Union All Select 'ҽ���˶Ա�','SELECT' From Dual
Union All Select 'ҽ�Ƹ��ʽ','SELECT' From Dual
Union All Select 'ҽ�ƿ���ʧ��ʽ','SELECT' From Dual
Union All Select '���Ʒ���Ŀ¼','SELECT' From Dual
Union All Select '���ƻ�����Ŀ','SELECT' From Dual
Union All Select '�����շѹ�ϵ','SELECT' From Dual
Union All Select '������ĿĿ¼','SELECT' From Dual
Union All Select '����ִ�п���','SELECT' From Dual
Union All Select '֤������','SELECT' From Dual
Union All Select '��������','SELECT' From Dual
Union All Select '���ѿ�����','SELECT' From Dual
Union All Select '���ѿ����Ŀ¼','SELECT' From Dual
Union All Select '���ѿ���Ϣ','SELECT' From Dual
Union All Select '�����ӿ�����','SELECT' From Dual
Union All Select '����Ʊ�ݿ�Ʊ��','SELECT' From Dual
Union All Select 'Ʊ�ݿ�Ʊ�����','SELECT' From Dual
Union All Select '����Ʊ�����','SELECT' From Dual
UNION ALL SELECT 'Zl_����Ʊ��ʹ�ü�¼_Insert','EXECUTE' From dual
UNION ALL SELECT 'Zl_����Ʊ��ʹ�ü�¼_Delete','EXECUTE' From dual
UNION ALL SELECT 'Zl_����Ʊ��ʹ�ü�¼_Update','EXECUTE' From dual
UNION ALL SELECT 'Zl_����Ʊ�ݶ�ά��_Update','EXECUTE' From dual
UNION ALL SELECT '����Ʊ�ݶ�ά��','UPDATE' From dual
UNION ALL SELECT 'Zl_ֽ��Ʊ��ʹ��_Update','EXECUTE' From dual
UNION ALL SELECT 'Zl_����Ʊ��վ�����_Update','EXECUTE' From dual
UNION ALL SELECT 'zl_����Ʊ�ݿ�Ʊ��_insert','EXECUTE' From dual
UNION ALL SELECT 'zl_����Ʊ�ݿ�Ʊ��_update','EXECUTE' From dual
UNION ALL SELECT 'zl_����Ʊ�ݿ�Ʊ��_start','EXECUTE' From dual
UNION ALL SELECT 'zl_����Ʊ�ݿ�Ʊ��_stop','EXECUTE' From dual
UNION ALL SELECT 'zl_����Ʊ�ݿ�Ʊ��_delete','EXECUTE' From dual
UNION ALL SELECT 'Zl_Ʊ�ݿ�Ʊ�����_Update','EXECUTE' From dual
UNION ALL SELECT 'Zl_�����ӿ�����_Set','EXECUTE' From dual
UNION ALL SELECT 'Zl_�����ӿ�����_Get','EXECUTE' From dual
) A;

Insert Into zlProgPrivs(ϵͳ, ���, ����, ������, ����, Ȩ��)
Select &n_System, 1006, '����', User, A.*
From (Select ����, Ȩ�� From zlProgPrivs Where 1 = 0
Union All Select 'zlClients','SELECT' From Dual
Union All Select '�������','SELECT' From Dual
Union All Select '����Ʊ��վ�����','SELECT' From Dual
UNION ALL SELECT 'Zl_����Ʊ��վ�����_Update','EXECUTE' From dual
Union All Select '����Ʊ�����','SELECT' From Dual
Union All Select 'Zl_����Ʊ�����_Update','EXECUTE' From Dual
) A;

Create Or Replace Procedure Zl_����Ʊ��ʹ�ü�¼_Insert
(
  Id_In         In ����Ʊ��ʹ�ü�¼.Id%Type,
  Ʊ��_In       In ����Ʊ��ʹ�ü�¼.Ʊ��%Type,
  ����id_In     In ����Ʊ��ʹ�ü�¼.����id%Type,
  ����id_In     In ����Ʊ��ʹ�ü�¼.����id%Type,
  ����_In       In ����Ʊ��ʹ�ü�¼.����%Type,
  �Ա�_In       In ����Ʊ��ʹ�ü�¼.�Ա�%Type,
  ����_In       In ����Ʊ��ʹ�ü�¼.����%Type,
  �����_In     In ����Ʊ��ʹ�ü�¼.�����%Type,
  סԺ��_In     In ����Ʊ��ʹ�ü�¼.סԺ��%Type,
  Ʊ�ݽ��_In   In ����Ʊ��ʹ�ü�¼.Ʊ�ݽ��%Type,
  ��Ʊ��_In     In ����Ʊ��ʹ�ü�¼.��Ʊ��%Type,
  ϵͳ��Դ_In   In ����Ʊ��ʹ�ü�¼.ϵͳ��Դ%Type,
  ����ʱ��_In   In ����Ʊ��ʹ�ü�¼.����ʱ��%Type,
  ��ע_In       In ����Ʊ��ʹ�ü�¼.��ע%Type,
  ����Ա���_In In ����Ʊ��ʹ�ü�¼.����Ա���%Type,
  ����Ա����_In In ����Ʊ��ʹ�ü�¼.����Ա����%Type,
  �Ǽ�ʱ��_In   In ����Ʊ��ʹ�ü�¼.�Ǽ�ʱ��%Type,
  ԭƱ��id_In   In ����Ʊ��ʹ�ü�¼.ԭƱ��id%Type := Null,
  �˿�id_In     In ����Ʊ��ʹ�ü�¼.ԭƱ��id%Type := Null
) As
  n_��¼״̬ ����Ʊ��ʹ�ü�¼.��¼״̬%Type;
Begin
  n_��¼״̬ := 1;

  Insert Into ����Ʊ��ʹ�ü�¼
    (ID, Ʊ��, ��¼״̬, ����id, ����id, ����, �Ա�, ����, �����, סԺ��, Ʊ�ݽ��, ����ʱ��, ԭƱ��id, �˿�id, ��Ʊ��, ϵͳ��Դ, ��ע, ����Ա���, ����Ա����, �Ǽ�ʱ��)
  Values
    (Id_In, Ʊ��_In, n_��¼״̬, ����id_In, Decode(Nvl(����id_In, 0), 0, Null, ����id_In), ����_In, �Ա�_In, ����_In,
     Decode(Nvl(�����_In, 0), 0, Null, �����_In), Decode(Nvl(סԺ��_In, 0), 0, Null, סԺ��_In), Ʊ�ݽ��_In, ����ʱ��_In, ԭƱ��id_In,
     �˿�id_In, ��Ʊ��_In, ϵͳ��Դ_In, ��ע_In, ����Ա���_In, ����Ա����_In, Nvl(�Ǽ�ʱ��_In, Sysdate));
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ʊ��ʹ�ü�¼_Insert;
/

Create Or Replace Procedure Zl_����Ʊ�ݶ�ά��_Update
(
  ʹ�ü�¼id_In In ����Ʊ�ݶ�ά��.ʹ�ü�¼id%Type,
  �Ƿ�ɾ��_In   Number := 0,
  ��ά��_In     Varchar2 := Null
) As
  -- �Ƿ�ɾ��_IN:1-��ʾɾ��;0-��ʾ��ɾ��
  n_Count Number(18);
Begin
  If Nvl(�Ƿ�ɾ��_In, 0) = 1 Then
    Delete ����Ʊ�ݶ�ά�� Where ʹ�ü�¼id = ʹ�ü�¼id_In;
    Return;
  End If;
  Select Count(1) Into n_Count From ����Ʊ�ݶ�ά�� Where ʹ�ü�¼id = ʹ�ü�¼id_In;
  If n_Count = 0 Then
    Insert Into ����Ʊ�ݶ�ά�� (ʹ�ü�¼id, ��ά��) Values (ʹ�ü�¼id_In, ��ά��_In);
  Else
    Update ����Ʊ�ݶ�ά�� Set ��ά�� = ��ά��_In Where ʹ�ü�¼id = ʹ�ü�¼id_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ʊ�ݶ�ά��_Update;
/


Create Or Replace Procedure Zl_����Ʊ��ʹ�ü�¼_Update
(
  Id_In         In ����Ʊ��ʹ�ü�¼.Id%Type,
  ����_In       In ����Ʊ��ʹ�ü�¼.����%Type,
  ����_In       In ����Ʊ��ʹ�ü�¼.����%Type,
  ������_In     In ����Ʊ��ʹ�ü�¼.������%Type,
  ����ʱ��_In   In ����Ʊ��ʹ�ü�¼.����ʱ��%Type,
  Url����_In    In ����Ʊ��ʹ�ü�¼.Url����%Type,
  Url����_In    In ����Ʊ��ʹ�ü�¼.Url����%Type,
  ��ע_In       In ����Ʊ��ʹ�ü�¼.��ע%Type,
  ��Ʊ��_In     In ����Ʊ��ʹ�ü�¼.��Ʊ��%Type,
  ϵͳ��Դ_In   In ����Ʊ��ʹ�ü�¼.ϵͳ��Դ%Type,
  Ʊ�ݽ��_In   In ����Ʊ��ʹ�ü�¼.Ʊ�ݽ��%Type := Null,
  ƾ֤����_In   In ����Ʊ��ʹ�ü�¼.ƾ֤����%Type := Null,
  ƾ֤����_In   In ����Ʊ��ʹ�ü�¼.ƾ֤����%Type := Null,
  ƾ֤������_In In ����Ʊ��ʹ�ü�¼.ƾ֤������%Type := Null
) As
Begin

  Update ����Ʊ��ʹ�ü�¼
  Set ���� = Nvl(����_In, ����), ���� = Nvl(����_In, ����), ������ = Nvl(������_In, ������), ����ʱ�� = ����ʱ��_In, Url���� = Nvl(Url����_In, Url����),
      Url���� = Nvl(Url����_In, Url����), ��ע = Nvl(��ע_In, ��ע), ��Ʊ�� = Nvl(��Ʊ��_In, ��Ʊ��), ϵͳ��Դ = ϵͳ��Դ_In,
      Ʊ�ݽ�� = Nvl(Ʊ�ݽ��_In, Ʊ�ݽ��), ƾ֤���� = Nvl(ƾ֤����_In, ƾ֤����), ƾ֤���� = Nvl(ƾ֤����_In, ƾ֤����), ƾ֤������ = Nvl(ƾ֤������_In, ƾ֤������)
  Where ID = Id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ʊ��ʹ�ü�¼_Update;
/

Create Or Replace Procedure Zl_����Ʊ��ʹ�ü�¼_Delete
(
  Id_In           In ����Ʊ��ʹ�ü�¼.Id%Type,
  ��Ʊ��_In       In ����Ʊ��ʹ�ü�¼.��Ʊ��%Type,
  ϵͳ��Դ_In     In ����Ʊ��ʹ�ü�¼.ϵͳ��Դ%Type,
  ����ʱ��_In     In ����Ʊ��ʹ�ü�¼.����ʱ��%Type,
  ��ע_In         In ����Ʊ��ʹ�ü�¼.��ע%Type,
  ����Ա���_In   In ����Ʊ��ʹ�ü�¼.����Ա���%Type,
  ����Ա����_In   In ����Ʊ��ʹ�ü�¼.����Ա����%Type,
  �Ǽ�ʱ��_In     In ����Ʊ��ʹ�ü�¼.�Ǽ�ʱ��%Type,
  ԭ����Ʊ��id_In In ����Ʊ��ʹ�ü�¼.Id%Type
) As
  v_Err_Msg Varchar2(200);
  Err_Item Exception;
  n_�Ƿ񻻿� ����Ʊ��ʹ�ü�¼.�Ƿ񻻿�%Type;
Begin

  Update ����Ʊ��ʹ�ü�¼ Set ��¼״̬ = 3 Where ID = ԭ����Ʊ��id_In Returning Nvl(�Ƿ񻻿�, 0) Into n_�Ƿ񻻿�;
  If Sql%NotFound Then
    v_Err_Msg := 'δ�ҵ�ԭʼ�ĵ���Ʊ����Ϣ���������ϲ���!';
    Raise Err_Item;
  End If;
  If Nvl(n_�Ƿ񻻿�, 0) = 1 Then
    --��ǰ����Ʊ���Ѿ�����ֽ��Ʊ��
    v_Err_Msg := '��ǰ����Ʊ���Ѿ�����ֽ��Ʊ��,��Ҫ�ȳ��ֽ��Ʊ�ݺ�������ϵ��ӷ�Ʊ!';
    Raise Err_Item;
  End If;

  Insert Into ����Ʊ��ʹ�ü�¼
    (ID, Ʊ��, ��¼״̬, ����id, ����id, ����, �Ա�, ����, �����, סԺ��, ����, ����, ������, Ʊ�ݽ��, Url����, Url����, ����ʱ��, ԭƱ��id, ��ӡid, �Ƿ񻻿�, ֽ�ʷ�Ʊ��,
     ��Ʊ��, ϵͳ��Դ, ��ע, ����Ա���, ����Ա����, �Ǽ�ʱ��, �˿�id)
    Select Id_In, Ʊ��, 2, ����id, ����id, ����, �Ա�, ����, �����, סԺ��, ����, ����, ������, Ʊ�ݽ��, Url����, Url����, ����ʱ��_In, ԭ����Ʊ��id_In, ��ӡid,
           �Ƿ񻻿�, ֽ�ʷ�Ʊ��, Nvl(��Ʊ��_In, ��Ʊ��) As ��Ʊ��, Nvl(ϵͳ��Դ_In, ϵͳ��Դ) As ϵͳ��Դ, Nvl(��ע_In, ��ע) As ��ע, ����Ա���_In, ����Ա����_In,
           �Ǽ�ʱ��_In, �˿�id
    From ����Ʊ��ʹ�ü�¼
    Where ID = ԭ����Ʊ��id_In;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ʊ��ʹ�ü�¼_Delete;
/

Create Or Replace Procedure Zl_ֽ��Ʊ��ʹ��_Update
(
  ������Դ_In   Ʊ�ݴ�ӡ����.��������%Type,
  Ʊ��_In       Ʊ��ʹ����ϸ.Ʊ��%Type,
  ����id_In     ����Ԥ����¼.����id%Type,
  ����Ʊ��id_In ����Ʊ��ʹ�ü�¼.Id%Type,
  Ʊ�ݺ�_In     Varchar2,
  Ʊ�ݽ��_In   Ʊ��ʹ����ϸ.Ʊ�ݽ��%Type,
  ����id_In     Ʊ��ʹ����ϸ.����id%Type,
  ʹ����_In     Ʊ��ʹ����ϸ.ʹ����%Type,
  ʹ��ʱ��_In   Ʊ��ʹ����ϸ.ʹ��ʱ��%Type,
  ������ʽ_In   Integer := 0,
  �Ƿ񲹽���_In Number := 0,
  ��Ʊ����_In   Number := 0
) As
  --���ܣ��û������ؿ�������ֽ��Ʊ��
  --������
  --     ������ʽ_In:0-����;1-���»���;2-����Ʊ��;3-����Ʊ��
  --     ������Դ_IN =1-�շ�,2-Ԥ��(����Ѻ��),3-����,4-�Һ�,5-���￨
  --     Ʊ��_in: 1-�շ�,2-Ԥ��,3-����,4-�Һ�
  --      ����ID_In:Ʊ��_in=2:ԭԤ��ID;����Ϊ����ID(ԭ����ID)
  --      Ʊ�ݺ�_IN������ö��ŷ���;
  --      ����ID:���Ϊ0��NULL,��ʾ���ϸ����Ʊ�ݡ�
  --      �Ƿ񲹽���_In-0-���ǲ�����;1-�ǲ�����
  --      ��Ʊ����_In-0-���Ǻ�Ʊ,1-����˿�(��ת��Ԥ��)�����ĺ�Ʊ.Ŀǰ�����Ԥ����Ч(Ʊ��_In=2)

  c_No t_StrList := t_StrList();

  n_�ջ�id   Ʊ�ݴ�ӡ����.Id%Type;
  n_��ӡid   Ʊ�ݴ�ӡ����.Id%Type;
  v_��Ʊ��   Ʊ��ʹ����ϸ.����%Type;
  v_���Ʊ Ʊ��ʹ����ϸ.����%Type;
  n_������   Number(2);
  n_Count    Number(18);
  v_Error    Varchar2(255);
  Err_Custom Exception;
  n_ԭ�� Number(2);
Begin

  If Nvl(������Դ_In, 0) = 1 Then
    --�շ�
    If Nvl(�Ƿ񲹽���_In, 0) = 0 Then
      Select NO Bulk Collect Into c_No From (Select Distinct NO From ������ü�¼ Where ����id = ����id_In);
    
    Else
      Select NO Bulk Collect Into c_No From (Select Distinct NO From ���ò����¼ Where ����id = ����id_In);
    
    End If;
  
  Elsif Nvl(������Դ_In, 0) = 2 Then
    --Ԥ��
    Select NO Bulk Collect Into c_No From (Select Distinct NO From ����Ԥ����¼ Where ID = ����id_In);
  
  Elsif Nvl(������Դ_In, 0) = 3 Then
    --����
    Select NO Bulk Collect Into c_No From (Select Distinct NO From ���˽��ʼ�¼ Where ID = ����id_In);
  
  Elsif Nvl(������Դ_In, 0) = 4 Then
    --�Һ�
    If Nvl(�Ƿ񲹽���_In, 0) = 0 Then
      Select NO Bulk Collect Into c_No From (Select Distinct NO From ������ü�¼ Where ����id = ����id_In);
    
    Else
      Select NO Bulk Collect Into c_No From (Select Distinct NO From ���ò����¼ Where ����id = ����id_In);
    
    End If;
  Elsif Nvl(������Դ_In, 0) = 5 Then
    --���￨
    Select NO Bulk Collect Into c_No From (Select Distinct NO From סԺ���ü�¼ Where ����id = ����id_In);
  
  Else
    v_Error := '��Ч������Դ(' || Nvl(������Դ_In, 0) || '),�޷����л������ݣ�';
    Raise Err_Custom;
  End If;
  If c_No.Count = 0 Then
    v_Error := 'δ�ҵ���Ӧ�Ľ�������(' || Nvl(����id_In, 0) || '),�޷����л������ݣ�';
    Raise Err_Custom;
  End If;

  --1.���ջ�Ʊ��
  Begin
    If Nvl(��Ʊ����_In, 0) = 1 Then
      If Nvl(������ʽ_In, 0) > 0 Then
        Select ID
        Into n_�ջ�id
        From (Select b.Id
               From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
               Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� = 6 And b.�������� = ������Դ_In And b.No = c_No(1) And a.Ʊ�� = Ʊ��_In And
                     Not Exists
                (Select 1 From Ʊ��ʹ����ϸ B Where a.���� = b.���� And a.��ӡid = b.��ӡid And ���� = 2)
               Order By a.ʹ��ʱ�� Desc)
        Where Rownum < 2;
        n_ԭ�� := 7;
      Else
        n_������ := 1;
      End If;
    Else
      Select ID
      Into n_�ջ�id
      From (Select b.Id
             From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
             Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And b.�������� = ������Դ_In And b.No = c_No(1) And a.Ʊ�� = Ʊ��_In And
                   Not Exists
              (Select 1 From Ʊ��ʹ����ϸ B Where a.���� = b.���� And a.��ӡid = b.��ӡid And ���� = 2)
             Order By a.ʹ��ʱ�� Desc)
      Where Rownum < 2;
      n_ԭ�� := 2;
    End If;
  Exception
    When Others Then
      n_������ := 1;
  End;

  --�ջ�Ʊ��(������ǰδ����Ʊ��,�޷��ջ�)
  If n_�ջ�id Is Not Null Then
    --Decode(������ʽ_In, 2, 5, 2)
    Insert Into Ʊ��ʹ����ϸ
      (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ����, ʹ��ʱ��, Ʊ�ݽ��, ����Ʊ��id)
      Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, n_ԭ��, ����id, ��ӡid, ʹ����_In, ʹ��ʱ��_In, Ʊ�ݽ��, ����Ʊ��id
      From Ʊ��ʹ����ϸ A
      Where ��ӡid = n_�ջ�id And ���� = 1 And a.Ʊ�� = Ʊ��_In And Not Exists
       (Select 1 From Ʊ��ʹ����ϸ B Where a.���� = b.���� And ��ӡid = n_�ջ�id And ���� = 2);
  Else
    n_������ := 1;
  End If;

  --��Ʊ�ݺ�ʱ,���ô���Ʊ��
  If Ʊ�ݺ�_In Is Null Or Nvl(������ʽ_In, 0) >= 2 Then
  
    If Nvl(������Դ_In, 0) = 1 Then
      --�շ�
    
      If Nvl(�Ƿ񲹽���_In, 0) = 0 Then
        Update ������ü�¼
        Set ʵ��Ʊ�� = Null
        Where Mod(��¼����, 10) = 1 And NO In (Select Column_Value From Table(c_No));
      Else
        Update ���ò����¼
        Set ʵ��Ʊ�� = Null
        Where Mod(��¼����, 10) = 1 And NO In (Select Column_Value From Table(c_No));
      End If;
    
    Elsif Nvl(������Դ_In, 0) = 2 Then
      --Ԥ��
      If Not Nvl(��Ʊ����_In, 0) = 1 Then
        Update ����Ԥ����¼
        Set ʵ��Ʊ�� = Null
        Where Mod(��¼����, 10) = 1 And NO In (Select Column_Value From Table(c_No));
      End If;
    Elsif Nvl(������Դ_In, 0) = 3 Then
      --����
    
      Update ���˽��ʼ�¼ Set ʵ��Ʊ�� = v_��Ʊ�� Where NO In (Select Column_Value From Table(c_No));
    
    Elsif Nvl(������Դ_In, 0) = 4 Then
      --�Һ�
      If Nvl(�Ƿ񲹽���_In, 0) = 0 Then
        Update ������ü�¼
        Set ʵ��Ʊ�� = Null
        Where Mod(��¼����, 10) = 4 And NO In (Select Column_Value From Table(c_No));
      Else
        Update ���ò����¼
        Set ʵ��Ʊ�� = Null
        Where Mod(��¼����, 10) = 1 And NO In (Select Column_Value From Table(c_No));
      End If;
    
    End If;
    Update ����Ʊ��ʹ�ü�¼ Set ��ӡid = Null, �Ƿ񻻿� = 0, ֽ�ʷ�Ʊ�� = Null Where ID = Nvl(����Ʊ��id_In, 0);
    Return;
  End If;

  v_��Ʊ�� := Substr(Ʊ�ݺ�_In || ',', 1, Instr(Ʊ�ݺ�_In || ',', ',') - 1);

  --���·���Ʊ�ݲ���дƱ�ݴ�ӡ����
  Select Ʊ�ݴ�ӡ����_Id.Nextval Into n_��ӡid From Dual;

  Insert Into Ʊ�ݴ�ӡ����
    (ID, ��������, NO, ��ӡ����)
    Select Distinct n_��ӡid, ������Դ_In, NO, 0 From (Select Distinct Column_Value As NO From Table(c_No));

  If Nvl(��Ʊ����_In, 0) = 1 And Nvl(Ʊ��_In, 0) = 2 Then
    n_ԭ�� := 6;
  Else
    If n_������ = 1 Then
      n_ԭ�� := 1;
    Else
      n_ԭ�� := 3;
    End If;
  End If;

  Insert Into Ʊ��ʹ����ϸ
    (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��, ����Ʊ��id)
    Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��_In, Column_Value As ��Ʊ��, 1, n_ԭ��, Decode(Nvl(����id_In, 0), 0, Null, ����id_In), n_��ӡid,
           ʹ��ʱ��_In, ʹ����_In, Ʊ�ݽ��_In, ����Ʊ��id_In
    From Table(f_Str2List(Ʊ�ݺ�_In));

  If Nvl(����id_In, 0) <> 0 Then
    Select Count(*) As n_Count, Max(Column_Value) Into n_Count, v_���Ʊ From Table(f_Str2List(Ʊ�ݺ�_In));
  
    Update Ʊ�����ü�¼ Set ʣ������ = Nvl(ʣ������, 0) - n_Count, ��ǰ���� = v_���Ʊ Where ID = ����id_In;
  
  End If;
  If Nvl(������Դ_In, 0) = 1 Then
    --�շ�
  
    If Nvl(�Ƿ񲹽���_In, 0) = 0 Then
      Update ������ü�¼
      Set ʵ��Ʊ�� = v_��Ʊ��
      Where Mod(��¼����, 10) = 1 And NO In (Select Column_Value From Table(c_No));
    Else
      Update ���ò����¼
      Set ʵ��Ʊ�� = v_��Ʊ��
      Where Mod(��¼����, 10) = 1 And NO In (Select Column_Value From Table(c_No));
    End If;
  
  Elsif Nvl(������Դ_In, 0) = 2 Then
    --Ԥ��
    If Not Nvl(��Ʊ����_In, 0) = 1 Then
      Update ����Ԥ����¼
      Set ʵ��Ʊ�� = v_��Ʊ��
      Where Mod(��¼����, 10) = 1 And NO In (Select Column_Value From Table(c_No));
    End If;
  Elsif Nvl(������Դ_In, 0) = 3 Then
    --����
  
    Update ���˽��ʼ�¼ Set ʵ��Ʊ�� = v_��Ʊ�� Where NO In (Select Column_Value From Table(c_No));
  
  Elsif Nvl(������Դ_In, 0) = 4 Then
    --�Һ�
    If Nvl(�Ƿ񲹽���_In, 0) = 0 Then
      Update ������ü�¼
      Set ʵ��Ʊ�� = v_��Ʊ��
      Where Mod(��¼����, 10) = 4 And NO In (Select Column_Value From Table(c_No));
    Else
      Update ���ò����¼
      Set ʵ��Ʊ�� = v_��Ʊ��
      Where Mod(��¼����, 10) = 1 And NO In (Select Column_Value From Table(c_No));
    End If;
  
    --ELSIF nvl(������Դ_in,0)=5 THEN  --���￨
    --���￨��ʵ��Ʊ�ݣ����Ժ���չ
    --UPDATE סԺ���ü�¼ SET ʵƱƱ��=v_��Ʊ�� WHERE  mod(��¼����,10)=1 AND NO IN (Select Column_value From table(c_NO));
  End If;

  Update ����Ʊ��ʹ�ü�¼ Set ��ӡid = n_��ӡid, �Ƿ񻻿� = 1, ֽ�ʷ�Ʊ�� = v_��Ʊ�� Where ID = Nvl(����Ʊ��id_In, 0);

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ֽ��Ʊ��ʹ��_Update;
/


Create Or Replace Procedure Zl_����Ʊ��վ�����_Update
(
  ����_In In ����Ʊ��վ�����.����%Type,
  վ��_In In Clob := Null
) Is

  --˵����
  --     ����_In��1-�շ�,2-Ԥ��,3-����,4-�Һ�
  --     վ��_In������Ʊ��ʹ�ü�¼.վ��,����ö��ŷָ�,������վ���ʾ��ɾ�� ����Ʊ��վ�����
Begin
  Delete From ����Ʊ��վ����� A Where a.���� = ����_In;

  For r_վ�� In (Select Column_Value As վ�� From Table(f_Str2List(վ��_In))) Loop
    Insert Into ����Ʊ��վ����� (����, վ��) Values (����_In, r_վ��.վ��);
  End Loop;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ʊ��վ�����_Update;
/

Create Or Replace Function Zl_Fun_Isstarteinvoice
(
  ����_In     Integer,
  ����_In     ���ս����¼.����%Type := 0,
  ���վ��_In Integer := 1,
  ����_In     Integer := Null
) Return Number Is
  ---------------------------------------------------------------------------
  --���ܣ��ж�ָ�������Ƿ������˵���Ʊ�� 
  --����������_In��1-�շ�,2-Ԥ��,3-����,4-�Һ�;5-���￨
  --      ���վ��_IN-1-��ʾ��Ҫ���վ���Ƿ�����;0-����飬ֱ���ж�ҵ���Ƿ�����
  --      ����_In:null-����������;����=Ԥ���ͽ��ʣ��ֱ����1-�����2-סԺ
  --����:1-���õ���Ʊ��;0-δ���õ���Ʊ��
  ---------------------------------------------------------------------------
  v_������       Varchar2(100);
  v_Para         Varchar2(4000);
  v_ҽ��         Varchar2(4000);
  n_����Ʊ������ Number(2);
  n_ҽ������     Number(2);
  n_����         Number(2);

  n_Return Number(2);
Begin

  If Nvl(����_In, 0) = 1 Then
    v_Para := zl_GetSysParameter('�շѵ���Ʊ�ݿ���');
  Elsif Nvl(����_In, 0) = 2 Then
    v_Para := zl_GetSysParameter('Ԥ������Ʊ�ݿ���');
    --��ʽ��Ԥ�����|Ʊ�����ÿ���|Ʊ�ݹ������|ҽ�����ÿ���
    v_Para := Nvl(v_Para, '') || '|||||';
    n_���� := To_Number(Substr(v_Para, 1, Instr(v_Para, '|') - 1));
    If n_���� <> 0 And Nvl(����_In, 0) <> 0 And Nvl(����_In, 0) <> n_���� Then
      Return 0;
    End If;
    v_Para := Substr(v_Para, Instr(v_Para, '|') + 1);
  Elsif Nvl(����_In, 0) = 3 Then
    v_Para := zl_GetSysParameter('���ʵ���Ʊ�ݿ���');
  Elsif Nvl(����_In, 0) = 4 Then
    v_Para := zl_GetSysParameter('�Һŵ���Ʊ�ݿ���');
  Elsif Nvl(����_In, 0) = 5 Then
    v_Para := zl_GetSysParameter('���￨����Ʊ�ݿ���');
  Else
    Return 0;
  End If;

  --��ʽ��Ʊ�����ÿ���|Ʊ�ݹ������|ҽ�����ÿ���
  v_Para         := Nvl(v_Para, '') || '|||||';
  n_����Ʊ������ := To_Number(Substr(v_Para, 1, Instr(v_Para, '|') - 1));
  If Nvl(n_����Ʊ������, 0) = 0 Then
    Return 0;
  End If;

  n_Return := 1;
  If Nvl(n_����Ʊ������, 0) = 2 And Nvl(���վ��_In, 1) = 1 Then
    --0-��ʾδ���õ���Ʊ��;1-�������õ���Ʊ��;2-�����վ�����õ���Ʊ��
    Select Terminal Into v_������ From V$session Where Audsid = Userenv('sessionid');
    Select Nvl(Max(1), 0) Into n_Return From ����Ʊ��վ����� Where ���� = Nvl(����_In, 0) And վ�� = v_������;
  End If;
  If Nvl(����_In, 0) = 0 Then
    --��ҽ����ֱ�ӷ���
    Return n_Return;
  End If;

  --ҽ����֤
  v_Para := Substr(v_Para, Instr(v_Para, '|') + 1);
  v_Para := Substr(v_Para, Instr(v_Para, '|') + 1);
  v_Para := Substr(v_Para, 1, Instr(v_Para, '|') - 1);
  If Instr(v_Para, ':') > 0 Then
    --ҽ����أ����ñ�־:��������
    v_ҽ��     := Substr(v_Para, Instr(v_Para, ':') + 1);
    n_ҽ������ := To_Number(Substr(v_Para, 1, Instr(v_Para, ':') - 1));
  End If;
  If Nvl(n_ҽ������, 0) = 0 Or v_ҽ�� Is Null Then
    Return 0;
  End If;
  If Instr(',' || v_ҽ�� || ',', ',' || ����_In || ',') > 0 Then
    Return 1;
  End If;
  Return 0;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Fun_Isstarteinvoice;
/


Create Or Replace Procedure Zl_�������ʽ���_Update
(
  ����id_In       ������ü�¼.����id%Type,
  ����id_In       ����Ԥ����¼.����id%Type,
  ���ս���_In     Varchar2,
  �������_In     �������.����%Type,
  ֧����ʽ_In     ���㷽ʽ.����%Type,
  ����Ա���_In   ����Ԥ����¼.����Ա���%Type,
  ����Ա����_In   ����Ԥ����¼.����Ա����%Type,
  �տ�ʱ��_In     ����Ԥ����¼.�տ�ʱ��%Type,
  ��ɽ���_In     Number := 0,
  �Ƿ����Ʊ��_In ����Ԥ����¼.�Ƿ����Ʊ��%Type := Null,
  ����_In         ���ս����¼.����%Type := Null
) As
  ------------------------------------------------------------------------------------------------------------------------------
  --����:�շѽ���ʱ,�޸Ľ���������Ϣ
  -- ���ս���_In:(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������||.."
  -- �Ƿ����Ʊ��_In-Ϊ��ʱ���ڲ��������Խ����ж��Ƿ�����
  -- ��������_IN:1-�������;2-סԺ����
  -- ��ɽ���_In:1-����շ�;0-δ����շ�
  ------------------------------------------------------------------------------------------------------------------------------
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_��������     Varchar2(500);
  v_��ǰ����     Varchar2(50);
  n_Ԥ��id       ����Ԥ����¼.Id%Type;
  n_��ҳid       ����Ԥ����¼.��ҳid%Type;
  n_����id       ����Ԥ����¼.����id%Type;
  v_���㷽ʽ     ����Ԥ����¼.���㷽ʽ%Type;
  n_������     ����Ԥ����¼.��Ԥ��%Type;
  n_ʣ���       ����Ԥ����¼.��Ԥ��%Type;
  n_���ʽ��     ����Ԥ����¼.��Ԥ��%Type;
  n_����ֵ       ��Ա�ɿ����.���%Type;
  v_����       ���㷽ʽ.����%Type;
  n_�����     ����Ԥ����¼.��Ԥ��%Type;
  n_�Ƿ����Ʊ�� ����Ԥ����¼.�Ƿ����Ʊ��%Type;
  n_Count        Number;
  n_Havenull     Number;
  l_Ԥ��id       t_NumList := t_NumList();
  n_�ɿ���id     ����Ԥ����¼.�ɿ���id%Type;
  n_��ֵid       ����Ԥ����¼.Id%Type;
  d_����ʱ��     ����Ԥ����¼.����ʱ��%Type;
  v_������Ա     ����Ԥ����¼.������Ա%Type;
  n_����         ���ս����¼.����%Type;
  Cursor c_Balance_Record Is
    Select Max(m.����id) As ����id, Max(NO) As NO, Max(Nvl(�տ�ʱ��_In, m.�շ�ʱ��)) As �շ�ʱ��, Max(Nvl(����Ա���_In, m.����Ա���)) As ����Ա���,
           Max(Nvl(m.����Ա����, ����Ա����_In)) As ����Ա����, Max(Nvl(n_�ɿ���id, m.�ɿ���id)) As �ɿ���id, Max(��������) As ��������
    
    From ���˽��ʼ�¼ M
    Where m.Id = ����id_In;
  r_Balance_Record c_Balance_Record%RowType;

  Cursor c_Balancedata Is
    Select ��¼����, NO, ��¼״̬, ����id, ��ҳid, ����id, ���㷽ʽ, Nvl(�տ�ʱ��_In, �տ�ʱ��) As �տ�ʱ��, Nvl(����Ա���_In, ����Ա���) As ����Ա���,
           Nvl(����Ա����_In, ����Ա����) As ����Ա����, ��Ԥ��, ����id, Nvl(n_�ɿ���id, �ɿ���id) As �ɿ���id, �������
    From ����Ԥ����¼
    Where ����id = ����id_In And ���㷽ʽ Is Null;

  r_Balancedata c_Balancedata%RowType;

  Procedure ����Ԥ����¼_��Ԥ��
  (
    ��Ԥ��_In        In Out ����Ԥ����¼.��Ԥ��%Type,
    Ԥ�����_In      ����Ԥ����¼.Ԥ�����%Type,
    ��Ԥ������ids_In Varchar2 := Null,
    ��������_In      Number := 2
  ) As
    --���������_In  0-Ԥ�������ʱ����ȥ������1-��ȥ������2-���ݽ��ж��ٳ���١�
    --��Ԥ��_In:������������_In=2���򷵻�δ��̯��ɵĽ�����;����ΪNULL��0
    v_��Ԥ������ids Varchar2(4000);
    n_����ֵ        ��Ա�ɿ����.���%Type;
    n_Ԥ�����      ����Ԥ����¼.��Ԥ��%Type;
    n_��Ԥ��        ����Ԥ����¼.��Ԥ��%Type;
    n_�Ự��        ����Ԥ����¼.�Ự��%Type; --��ʽ��SID+'_'+SERIAL#
    n_��id          ����ɿ����.Id%Type;
  Begin
    v_��Ԥ������ids := Nvl(��Ԥ������ids_In, ����id_In);
    n_��id          := Zl_Get��id(����Ա����_In);
  
    --Ԥ�����
    If Nvl(��Ԥ��_In, 0) = 0 Then
      Return;
    End If;
    Select Max(Sid || '_' || Serial#) Into n_�Ự�� From V$session Where Audsid = Userenv('sessionid');
  
    n_Ԥ����� := ��Ԥ��_In; --�Ƚ����ã��������Լ���
    --���������㷽ʽΪ���տ����Ԥ���
    For c_��Ԥ�� In (Select a.No, b.Ԥ����� As ���, Nvl(a.����id, 0) As ����id, a.����id, a.��¼״̬, a.Id, a.�տ�ʱ��, a.��������id
                  From ����Ԥ����¼ A, Ԥ��������� B
                  Where a.Id = b.Ԥ��id And b.����id In (Select Column_Value From Table(f_Num2List(v_��Ԥ������ids))) And
                        Nvl(b.Ԥ�����, 2) = Ԥ�����_In And a.���㷽ʽ Not In (Select ���� From ���㷽ʽ Where ���� = 5) And
                        Nvl(a.У�Ա�־, 0) = 0
                  Order By Decode(����id, Nvl(����id_In, 0), 0, 1), a.�տ�ʱ��) Loop
    
      If c_��Ԥ��.��� - n_Ԥ����� < 0 Then
        n_��Ԥ�� := c_��Ԥ��.���;
      Else
        n_��Ԥ�� := n_Ԥ�����;
      End If;
    
      If c_��Ԥ��.����id = 0 Then
        --��һ�γ�Ԥ��(���Ͻ���ID,���Ϊ0)
        Update ����Ԥ����¼
        Set ��Ԥ�� = 0, ����id = ����id_In, ������� = -1 * ����id_In, �������� = ��������_In, �Ự�� = n_�Ự��
        Where ID = c_��Ԥ��.Id;
      End If;
      --���ϴ�ʣ���
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, ��������, �Ự��, ��������id, ����ʱ��, ������Ա, У�Ա�־)
        Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��_In,
               ����Ա����_In, ����Ա���_In, n_��Ԥ��, ����id_In, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, -1 * ����id_In,
               ��������_In, n_�Ự��, c_��Ԥ��.��������id, �տ�ʱ��_In, ����Ա����_In, 0
        From ����Ԥ����¼
        Where NO = c_��Ԥ��.No And ��¼״̬ = c_��Ԥ��.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
    
      --��������(���㷽ʽΪNULL��)
      Update ����Ԥ����¼
      Set ��Ԥ�� = ��Ԥ�� - n_��Ԥ��
      Where ����id = ����id_In And ���㷽ʽ Is Null
      Returning Nvl(��Ԥ��, 0) Into n_����ֵ;
    
      --���²���Ԥ�����
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) - n_��Ԥ��
      Where ����id = c_��Ԥ��.����id And ���� = 1 And ���� = Ԥ�����_In
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ������� (����id, ����, Ԥ�����, ����) Values (c_��Ԥ��.����id, Ԥ�����_In, -1 * n_��Ԥ��, 1);
        n_����ֵ := -1 * n_��Ԥ��;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From �������
        Where ����id = c_��Ԥ��.����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
    
      --����Ԥ���������
      Update Ԥ���������
      Set Ԥ����� = Nvl(Ԥ�����, 0) - n_��Ԥ��
      Where ����id = c_��Ԥ��.����id And Ԥ��id = c_��Ԥ��.Id
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into Ԥ���������
          (Ԥ��id, ����id, Ԥ�����, Ԥ�����)
        Values
          (c_��Ԥ��.Id, c_��Ԥ��.����id, Ԥ�����_In, -1 * n_��Ԥ��);
        n_����ֵ := -1 * n_��Ԥ��;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From Ԥ��������� Where Ԥ��id = c_��Ԥ��.Id And Nvl(Ԥ�����, 0) = 0;
      End If;
    
      --����Ƿ��Ѿ�������
      If c_��Ԥ��.��� < n_Ԥ����� Then
        n_Ԥ����� := n_Ԥ����� - c_��Ԥ��.���;
      Else
        n_Ԥ����� := 0;
      End If;
      If n_Ԥ����� = 0 Then
        Exit;
      End If;
    
    End Loop;
    --������Ƿ��㹻
    ��Ԥ��_In := n_Ԥ�����;
  End ����Ԥ����¼_��Ԥ��;

Begin

  If ����Ա����_In Is Null Then
    n_�ɿ���id := Null;
  Else
    n_�ɿ���id := Zl_Get��id(����Ա����_In);
  End If;

  --0.��ʽ����
  Select Max(Decode(���㷽ʽ, Null, 1, 0)) Into n_Havenull From ����Ԥ����¼ Where ����id = ����id_In;

  If Nvl(n_Count, 0) = 0 Then
    --���ӽ��㷽ʽΪNULL�ļ�¼
    Begin
      Select ���� Into v_���� From ���㷽ʽ Where Nvl(����, 0) = 9;
    Exception
      When Others Then
        v_���� := '����';
    End;
  End If;

  --1.���ӽ��㷽ʽΪ�յĽ�������
  If Nvl(n_Havenull, 0) = 0 Then
  
    Open c_Balance_Record;
    Fetch c_Balance_Record
      Into r_Balance_Record;
  
    If c_Balance_Record%RowCount = 0 Then
      Close c_Balance_Record;
      v_Err_Msg := 'δ�ҵ����ʼ�¼,������Ϊ����ԭ��ɾ���˽�������,�����²�������!';
      Raise Err_Item;
    End If;
  
    Select Sum(Nvl(���ʽ��, 0))
    Into n_������
    From (Select Sum(���ʽ��) As ���ʽ��
           From סԺ���ü�¼
           Where ����id = ����id_In
           Union All
           Select Sum(���ʽ��) As ���ʽ��
           From ������ü�¼
           Where ����id = ����id_In);
  
    n_����� := Round(n_������ - Round(Nvl(n_������, 0), 2), 6);
    n_������ := Round(Nvl(n_������, 0), 2);
  
    Select a.��ҳid, a.��Ժ����id
    Into n_��ҳid, n_����id
    From ������ҳ A, ������Ϣ B
    Where a.����id = ����id_In And a.����id = b.����id And a.��ҳid = b.��ҳid;
  
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ����id, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, У�Ա�־, ��������)
    Values
      (����Ԥ����¼_Id.Nextval, 2, r_Balance_Record.No, 1, Decode(����id_In, 0, Null, ����id_In), Decode(n_��ҳid, 0, Null, n_��ҳid),
       Decode(n_����id, 0, Null, n_����id), Null, r_Balance_Record.�շ�ʱ��, r_Balance_Record.����Ա���, r_Balance_Record.����Ա����,
       n_������, ����id_In, r_Balance_Record.�ɿ���id, 1, 2);
  
    --����(�Ȼ��ܺ���������
    If n_����� <> 0 Then
      Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, ��������, ��������id)
      Values
        (n_Ԥ��id, 2, r_Balance_Record.No, 1, Decode(����id_In, 0, Null, ����id_In), v_����, r_Balance_Record.�շ�ʱ��,
         r_Balance_Record.����Ա���, r_Balance_Record.����Ա����, n_�����, ����id_In, r_Balance_Record.�ɿ���id, -1 * ����id_In, 1, 2,
         n_Ԥ��id);
    End If;
    Close c_Balance_Record;
  End If;

  Open c_Balancedata;
  Fetch c_Balancedata
    Into r_Balancedata;

  --1.�Ȼ�������
  n_������ := 0;
  n_ʣ���   := 0;
  For c_���� In (Select a.Id, a.No, a.��¼���� As ��¼����, a.���㷽ʽ, a.����id, a.��Ԥ��, a.Ԥ�����, b.����,
                      Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, 0, a.Id)) As Ԥ��id
               From ����Ԥ����¼ A, ���㷽ʽ B
               Where a.����id = ����id_In And a.���㷽ʽ = b.����(+)) Loop
  
    If c_����.���� <> 9 And c_����.���㷽ʽ Is Not Null Then
      If c_����.��¼���� = 1 Or c_����.��¼���� = 11 Then
      
        Update �������
        Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(c_����.��Ԥ��, 0)
        Where ����id = Nvl(c_����.����id, 0) And ���� = 1 And ���� = Nvl(c_����.Ԥ�����, 2)
        Returning Ԥ����� Into n_����ֵ;
        If Sql%NotFound Then
          Insert Into �������
            (����id, ����, Ԥ�����, ����)
          Values
            (c_����.����id, Nvl(c_����.Ԥ�����, 2), Nvl(c_����.��Ԥ��, 0), 1);
        
          n_����ֵ := Nvl(c_����.��Ԥ��, 0);
        End If;
      
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete ������� Where ���� = 1 And ����id = c_����.����id And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
        End If;
      
        --����Ԥ���������
        n_��ֵid := c_����.Ԥ��id;
        If Nvl(n_��ֵid, 0) = 0 Then
          Select Max(ID) Into n_��ֵid From ����Ԥ����¼ Where NO = c_����.No And ��¼���� = 1 And ��¼״̬ <> 2;
        End If;
      
        If n_��ֵid <> 0 Then
        
          Update Ԥ���������
          Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(c_����.��Ԥ��, 0)
          Where ����id = c_����.����id And Ԥ��id = n_��ֵid
          Returning Ԥ����� Into n_����ֵ;
        
          If Sql%RowCount = 0 Then
            Insert Into Ԥ���������
              (Ԥ��id, ����id, Ԥ�����, Ԥ�����)
            Values
              (n_��ֵid, c_����.����id, Nvl(c_����.Ԥ�����, 2), Nvl(c_����.��Ԥ��, 0));
            n_����ֵ := Nvl(c_����.��Ԥ��, 0);
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From Ԥ��������� Where Ԥ��id = n_��ֵid And Nvl(Ԥ�����, 0) = 0;
          End If;
        
        End If;
      End If;
      n_������ := Nvl(n_������, 0) + Nvl(c_����.��Ԥ��, 0);
      If c_����.��¼���� = 11 Then
        l_Ԥ��id.Extend;
        l_Ԥ��id(l_Ԥ��id.Count) := c_����.Id;
      Else
        Update ����Ԥ����¼ Set ����id = Null, ��Ԥ�� = Null Where ID = c_����.Id;
      End If;
    
    End If;
  
    n_ʣ��� := Nvl(n_ʣ���, 0) + Nvl(c_����.��Ԥ��, 0);
  End Loop;
  n_ʣ��� := Nvl(n_ʣ���, 0) - Nvl(n_�����, 0);

  If Nvl(n_������, 0) <> 0 Then
    Update ����Ԥ����¼ Set ��Ԥ�� = Nvl(��Ԥ��, 0) + Nvl(n_������, 0) Where ����id = ����id_In And ���㷽ʽ Is Null;
    If Sql%NotFound Then
      v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[���˽��ʴ���]�����½��ʣ�';
      Raise Err_Item;
    End If;
  End If;
  If l_Ԥ��id.Count <> 0 Then
    Forall I In 1 .. l_Ԥ��id.Count
      Delete ����Ԥ����¼ Where ID = l_Ԥ��id(I);
  End If;

  --2.�ٴ���ҽ������

  If Not ���ս���_In Is Null Then
    n_������ := 0;
    v_�������� := ���ս���_In || '||';
    n_Ԥ��id   := Null;
    d_����ʱ�� := Sysdate;
    v_������Ա := zl_UserName;
  
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      n_������ := To_Number(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
    
      If n_������ <> 0 Then
        n_ʣ��� := Nvl(n_ʣ���, 0) - Nvl(n_������, 0);
        If Nvl(n_Ԥ��id, 0) = 0 Then
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
        End If;
      
        If Nvl(n_��ֵid, 0) = 0 Then
          n_��ֵid := n_Ԥ��id;
        End If;
      
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, У�Ա�־, ��������, �ɿλ,
           ��������id, ����ʱ��, ������Ա)
        Values
          (n_Ԥ��id, 2, r_Balancedata.No, 1, r_Balancedata.����id, r_Balancedata.��ҳid, r_Balancedata.����id, '���ս���', v_���㷽ʽ,
           r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���, r_Balancedata.����Ա����, n_������, r_Balancedata.����id, r_Balancedata.�ɿ���id,
           2, 2, �������_In, n_��ֵid, d_����ʱ��, v_������Ա);
      
        --��������(���㷽ʽΪNULL��)
        Update ����Ԥ����¼
        Set ��Ԥ�� = ��Ԥ�� - n_������
        Where ����id = ����id_In And ���㷽ʽ Is Null
        Returning Nvl(��Ԥ��, 0) Into n_����ֵ;
      End If;
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
      n_Ԥ��id   := Null;
    End Loop;
  
    If Nvl(��ɽ���_In, 0) = 1 Then
      --ҽ����ر�Ĵ���
      Update ���ս�����ϸ Set ��־ = 2 Where ����id = ����id_In;
    End If;
  End If;

  --Ԥ�����
  If Nvl(n_ʣ���, 0) <> 0 Then
  
    If Nvl(����id_In, 0) = 0 Then
      v_Err_Msg := '����ȷ�����˵Ĳ���ID,�շѲ���ʹ��Ԥ�������,�������ʧ�ܣ�';
      Raise Err_Item;
    End If;
  
    ����Ԥ����¼_��Ԥ��(n_ʣ���, 2, Null, 2);
  
    If Nvl(n_ʣ���, 0) > 0 Then
    
      If ֧����ʽ_In Is Null Then
        v_Err_Msg := '����ȷ���ɿʽ������ɿʽ�Ƿ���ȷ,�������ʧ�ܣ�';
        Raise Err_Item;
      End If;
      n_���ʽ�� := n_ʣ���;
    
      Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��,
         ����id, �ɿ�, �Ҳ�, �ɿ���id, У�Ա�־, ��������, ��������id)
      Values
        (n_Ԥ��id, r_Balancedata.No, Null, 2, 1, ����id_In, r_Balancedata.��ҳid, r_Balancedata.����id, Null, ֧����ʽ_In, Null,
         '���ʽɿ�', Null, Null, Null, r_Balancedata.�տ�ʱ��, ����Ա���_In, ����Ա����_In, n_���ʽ��, ����id_In, Null, Null,
         r_Balancedata.�ɿ���id, 2, 2, n_Ԥ��id);
      --��������(���㷽ʽΪNULL��)
      Update ����Ԥ����¼
      Set ��Ԥ�� = ��Ԥ�� - n_���ʽ��
      Where ����id = ����id_In And ���㷽ʽ Is Null
      Returning Nvl(��Ԥ��, 0) Into n_����ֵ;
    
    End If;
  
  End If;
  If Nvl(��ɽ���_In, 0) = 0 Then
    Close c_Balancedata;
    Return;
  End If;

  -----------------------------------------------------------------------------------------
  --����շ�,��Ҫ������Ա�ɿ����,Ԥ����¼(���㷽ʽ=NULL)
  --1.ɾ�����㷽ʽΪNULL��Ԥ����¼
  Delete ����Ԥ����¼ Where ����id = ����id_In And ���㷽ʽ Is Null And Nvl(��Ԥ��, 0) = 0;
  If Sql%NotFound Then
    Select Count(*) Into n_Count From ����Ԥ����¼ A Where ����id = ����id_In And ���㷽ʽ Is Null;
    If n_Count <> 0 Then
      v_Err_Msg := '������δ�ɿ������,������ɽ���!';
    Else
      v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�!';
    End If;
    Raise Err_Item;
  End If;

  --������Ϊ��ʱ������һ�����Ϊ0�Ĳ���Ԥ����¼
  Select Count(*) Into n_Count From ����Ԥ����¼ A Where ����id = ����id_In;

  If n_Count = 0 Then
    v_���㷽ʽ := ֧����ʽ_In;
    If v_���㷽ʽ Is Null Then
      Select Max(���㷽ʽ) Into v_���㷽ʽ From ���㷽ʽӦ�� Where Ӧ�ó��� = '����' And Nvl(ȱʡ��־, 0) = 1;
      If v_���㷽ʽ Is Null Then
        Select Nvl(Max(����), '�ֽ�') Into v_���㷽ʽ From ���㷽ʽ Where Nvl(����, 0) = 1;
      End If;
    End If;
  
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, У�Ա�־, �����id, ���㿨���, ����,
       ������ˮ��, ����˵��, �������, ��������)
    Values
      (����Ԥ����¼_Id.Nextval, 2, r_Balancedata.No, 1, r_Balancedata.����id, r_Balancedata.��ҳid, r_Balancedata.����id, '���ʽɿ�',
       v_���㷽ʽ, r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���, r_Balancedata.����Ա����, 0, r_Balancedata.����id, r_Balancedata.�ɿ���id,
       2, Null, Null, Null, Null, Null, Null, 2);
  End If;

  n_�Ƿ����Ʊ�� := �Ƿ����Ʊ��_In;
  If �Ƿ����Ʊ��_In Is Null Then
    n_���� := Nvl(����_In, 0);
    If ����_In Is Null Then
      Select Max(����) Into n_���� From ���ս����¼ Where ��¼id = ����id_In And ���� = 2;
    End If;
    n_�Ƿ����Ʊ�� := Zl_Fun_Isstarteinvoice(3, n_����);
  End If;
  --2.����ɿ����ݺ��Ҳ����ݼ�У�Ա�־����Ϊ0
  Update ����Ԥ����¼ Set У�Ա�־ = 0, �Ƿ����Ʊ�� = n_�Ƿ����Ʊ�� Where ����id = ����id_In;

  --3.���·���״̬
  Update ���˽��ʼ�¼ Set ����״̬ = Null,�Ƿ����Ʊ�� = n_�Ƿ����Ʊ�� Where ID = ����id_In;

  --4.������Ա�ɿ�����
  For c_�ɿ� In (Select ���㷽ʽ, ����Ա����, Sum(Nvl(a.��Ԥ��, 0)) As ��Ԥ��
               From ����Ԥ����¼ A
               Where a.����id = ����id_In And Mod(a.��¼����, 10) <> 1
               Group By ���㷽ʽ, ����Ա����) Loop
  
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + Nvl(c_�ɿ�.��Ԥ��, 0)
    Where �տ�Ա = c_�ɿ�.����Ա���� And ���� = 1 And ���㷽ʽ = c_�ɿ�.���㷽ʽ;
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ����
        (�տ�Ա, ���㷽ʽ, ����, ���)
      Values
        (c_�ɿ�.����Ա����, c_�ɿ�.���㷽ʽ, 1, Nvl(c_�ɿ�.��Ԥ��, 0));
    End If;
  End Loop;
  Close c_Balancedata;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�������ʽ���_Update;
/


Create Or Replace Procedure Zl_���˽��ʽ���_Modify
(
  ��������_In      Number,
  ����id_In        ������ü�¼.����id%Type,
  ����id_In        ����Ԥ����¼.����id%Type,
  ���㷽ʽ_In      Varchar2,
  ��Ԥ��_In        ����Ԥ����¼.��Ԥ��%Type := Null,
  ��֧Ʊ��_In      ����Ԥ����¼.��Ԥ��%Type := Null,
  �����id_In      ����Ԥ����¼.�����id%Type := Null,
  ����_In          ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In      ����Ԥ����¼.����˵��%Type := Null,
  �ɿ�_In          ����Ԥ����¼.�ɿ�%Type := Null,
  �Ҳ�_In          ����Ԥ����¼.�Ҳ�%Type := Null,
  �����_In      ������ü�¼.ʵ�ս��%Type := Null,
  ��������_In      Number := 2,
  ȱʡ���㷽ʽ_In  ���㷽ʽ.����%Type := Null,
  ����Ա���_In    ����Ԥ����¼.����Ա���%Type := Null,
  ����Ա����_In    ����Ԥ����¼.����Ա����%Type := Null,
  �տ�ʱ��_In      ����Ԥ����¼.�տ�ʱ��%Type := Null,
  ��Ԥ������ids_In Varchar2 := Null,
  ��ɽ���_In      Number := 0,
  У�Ա�־_In      Number := 2,
  Ԥ��id_In        ����Ԥ����¼.Id%Type := Null,
  ��������id_In    ����Ԥ����¼.Id%Type := Null,
  ���ԭ����_In    Number := 0,
  ���ӱ�־_In      ����Ԥ����¼.Id%Type := Null,
  ����Ự_In      Number := 1,
  �Ƿ����Ʊ��_In  ����Ԥ����¼.�Ƿ����Ʊ��%Type := Null
) As
  ------------------------------------------------------------------------------------------------------------------------------
  --����:�շѽ���ʱ,�޸Ľ���������Ϣ
  --��������_In:
  --   0-��ͨ�շѷ�ʽ:
  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ|����||.." ;Ҳ�������.
  --     ����֧Ʊ��_In:����漰��֧Ʊ,���뱾�ε���֧Ʊ��,�������շ�ʱ,������
  --   1.����������:
  --     �ٽ��㷽ʽ_IN:���Դ��������㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
  --     ����֧Ʊ��_In:������
  --     �ۿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
  --     ��У�Ա�־_IN:���Դ��룬����Ϊ2:1-��У�ԵĽ���;2-�ӿڵ��óɹ���֧���ɹ�
  --     ���Ƿ�ת��_IN:�������Ŵ���
  --     @��������id_In:��������Ŵ���
  --     @Ԥ��id_In:��������Ŵ���
  --   2-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
  --     ����֧Ʊ��_In:������
  --   3-���ѿ�����:
  --     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."  ���ѿ�ID:Ϊ��ʱ,���ݿ����Զ���λ
  --     �ڳ�Ԥ��_In: ������
  --     ����֧Ʊ��_In:������
  -- ��Ԥ��_In: ���ڳ�Ԥ��ʱ,����
  -- �����_In:��������ʱ,����
  --  ��������_IN:1-�������;2-סԺ����
  --��Ԥ������ids_In:����ö��ַ���,��Ԥ��ʱ��Ч(��Ԥ�������ҵ���������һ��),��Ҫ��ʹ�ü�����Ԥ����
  -- ��ɽ���_In:1-����շ�;0-δ����շ�
  -- ���ӱ�־_IN:���ڽ���ҵ��(�������˿��)��NULl or 0-��ͨҵ��;1-�ֽ����˿�,2-����һ�ν��׽ӿ��˿�;3-ת�ʷ�ʽ�˿�
  -- ����Ự_In:1-��ʾ�ӷŻỰ��0-��ʾ������Ự
  --�Ƿ����Ʊ��_In:null-��ʾ�����ڲ�ֱ���жϣ��ǿձ�ʾֱ���Դ����Ϊ׼
  ------------------------------------------------------------------------------------------------------------------------------
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_��������   Varchar2(500);
  v_��ǰ����   Varchar2(300);
  v_����       ����ҽ�ƿ���Ϣ.����%Type;
  n_���ѿ�id   ���ѿ���Ϣ.Id%Type;
  n_�����id   ����Ԥ����¼.���㿨���%Type;
  v_����       Varchar2(100);
  n_Ԥ��id     ����Ԥ����¼.Id%Type;
  n_��������id ����Ԥ����¼.Id%Type;
  v_���㷽ʽ   ����Ԥ����¼.���㷽ʽ%Type;
  n_������   ����Ԥ����¼.��Ԥ��%Type;
  n_����ֵ     ��Ա�ɿ����.���%Type;
  n_��Ԥ��     ����Ԥ����¼.��Ԥ��%Type;
  v_��֧Ʊ     ����Ԥ����¼.���㷽ʽ%Type;
  v_�������   ����Ԥ����¼.�������%Type;
  v_����ժҪ   ����Ԥ����¼.ժҪ%Type;
  v_����     ���㷽ʽ.����%Type;
  n_�����   ����Ԥ����¼.��Ԥ��%Type;
  n_����id     ����Ԥ����¼.����id%Type;
  n_Count      Number;
  n_Havenull   Number;
  l_Ԥ��id     t_NumList := t_NumList();
  n_�ɿ���id   ����Ԥ����¼.�ɿ���id%Type;
  n_���ý��   ������ü�¼.���ʽ��%Type;
  n_���ʽ��   ����Ԥ����¼.��Ԥ��%Type;
  v_������Ա   ����Ԥ����¼.������Ա%Type;
  d_����ʱ��   ����Ԥ����¼.����ʱ��%Type;

  n_����         ���ս����¼.����%Type;
  n_У�Ա�־     ����Ԥ����¼.У�Ա�־%Type;
  n_�Ƿ�δ��     �����˿���Ϣ.�Ƿ�δ��%Type;
  n_�˿���     �����˿���Ϣ.���%Type;
  v_�Ự��       ����Ԥ����¼.�Ự��%Type; --��ʽ��SID+'_'+SERIAL#
  n_�Ƿ����Ʊ�� Number(2);
  Cursor c_Balance_Record Is
    Select ����id, NO, Nvl(�տ�ʱ��_In, �շ�ʱ��) As �շ�ʱ��, Nvl(����Ա���_In, ����Ա���) As ����Ա���, Nvl(����Ա����_In, ����Ա����) As ����Ա����,
           Nvl(n_�ɿ���id, �ɿ���id) As �ɿ���id, �������� As ��������, ��ҳid
    From ���˽��ʼ�¼
    Where ID = ����id_In;

  r_Balance_Record c_Balance_Record%RowType;

  Cursor c_Balancedata Is
    Select ��¼����, NO, ��¼״̬, ����id, ��ҳid, ����id, ���㷽ʽ, Nvl(�տ�ʱ��_In, �տ�ʱ��) As �տ�ʱ��, Nvl(����Ա���_In, ����Ա���) As ����Ա���,
           Nvl(����Ա����_In, ����Ա����) As ����Ա����, ��Ԥ��, ����id, Nvl(n_�ɿ���id, �ɿ���id) As �ɿ���id, �������
    From ����Ԥ����¼
    Where ����id = ����id_In And ���㷽ʽ Is Null;

  r_Balancedata c_Balancedata%RowType;

Begin

  n_����� := �����_In;

  If Nvl(����Ự_In, 0) = 1 Then
    Select Max(Sid || '_' || Serial#) Into v_�Ự�� From V$session Where Audsid = Userenv('sessionid');
  End If;

  --0.��ʽ����
  Select Count(1), Max(Decode(���㷽ʽ, Null, 1, 0)), Sum(��Ԥ��), Max(Decode(���㷽ʽ, Null, �ɿ���id, 0))
  Into n_Count, n_Havenull, n_��Ԥ��, n_�ɿ���id
  From ����Ԥ����¼
  Where ����id = ����id_In;

  If Nvl(n_�ɿ���id, 0) = 0 Then
    n_�ɿ���id := Null;
  End If;

  --1.���ӽ��㷽ʽΪ�յĽ�������
  If Nvl(n_Havenull, 0) = 0 Then
    n_Count := 0;
    Open c_Balance_Record;
    Fetch c_Balance_Record
      Into r_Balance_Record;
  
    If c_Balance_Record%NotFound Then
      Close c_Balance_Record;
      v_Err_Msg := 'δ�ҵ�ָ���Ľ�������,��ǰ�������ʧ�ܣ�';
      Raise Err_Item;
    End If;
  
    n_�ɿ���id := r_Balance_Record.�ɿ���id;
    If Nvl(n_�ɿ���id, 0) = 0 Then
      n_�ɿ���id := Null;
    End If;
    Select Sum(Nvl(���ʽ��, 0))
    Into n_������
    From (Select Sum(���ʽ��) As ���ʽ��
           From סԺ���ü�¼
           Where ����id = ����id_In
           Union All
           Select Sum(���ʽ��) As ���ʽ��
           From ������ü�¼
           Where ����id = ����id_In);
  
    n_����� := n_������ - Round(Nvl(n_������, 0), 6);
    n_������ := Round(Nvl(n_������, 0) - Nvl(n_��Ԥ��, 0), 6);
  
    n_����id := Null;
    If Nvl(r_Balance_Record.��������, 0) = 2 Then
      --סԺ�Ĳ��п���ID
      Select Max(��ǰ����id) Into n_����id From ������Ϣ Where ����id = ����id_In;
    End If;
  
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ����id, ��ҳid, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, У�Ա�־, ��������, �Ự��)
    Values
      (����Ԥ����¼_Id.Nextval, 2, r_Balance_Record.No, 1, r_Balance_Record.����id, n_����id, r_Balance_Record.��ҳid, Null,
       r_Balance_Record.�շ�ʱ��, r_Balance_Record.����Ա���, r_Balance_Record.����Ա����, n_������, ����id_In, r_Balance_Record.�ɿ���id, 1,
       2, v_�Ự��);
  
    n_����� := Nvl(n_�����, 0) + Nvl(�����_In, 0);
    Close c_Balance_Record;
  End If;

  Open c_Balancedata;
  Fetch c_Balancedata
    Into r_Balancedata;

  If c_Balancedata%NotFound Then
    Close c_Balancedata;
    v_Err_Msg := 'δ�ҵ�ָ���Ľ�������,�������ʧ�ܣ�';
    Raise Err_Item;
  End If;

  If Nvl(���ԭ����_In, 0) = 1 Then
    --����У�Ա�־Ϊ1�ļ�¼
    n_������ := 0;
  
    For c_У�� In (Select ID, ��Ԥ��
                 From ����Ԥ����¼
                 Where ��¼���� = 2 And ����id = r_Balancedata.����id And Nvl(�����id, 0) = �����id_In And
                       ��������id = Nvl(��������id_In, 0)) Loop
      n_������ := Round(n_������ + Nvl(c_У��.��Ԥ��, 0), 5);
    
      l_Ԥ��id.Extend;
      l_Ԥ��id(l_Ԥ��id.Count) := c_У��.Id;
    
    End Loop;
  
    If n_������ <> 0 Then
      Update ����Ԥ����¼
      Set ��Ԥ�� = Nvl(��Ԥ��, 0) + Nvl(n_������, 0)
      Where ����id = ����id_In And ���㷽ʽ Is Null;
      If Sql%NotFound Then
        v_Err_Msg := ' ������Ϣ����, ������Ϊ����ԭ����ɽ�����Ϣ����, ���� [ �շѽ��㴰�� ] �������շѣ� ';
        Raise Err_Item;
      End If;
    
    End If;
    If l_Ԥ��id.Count <> 0 Then
      --Ԥ��ɾ�����󣬴��Ͻ���ID
      Forall I In 1 .. l_Ԥ��id.Count
        Delete ����Ԥ����¼ Where ID = l_Ԥ��id(I) And ����id + 0 = ����id_In;
    End If;
  
  End If;

  --2.��������
  If Nvl(n_�����, 0) <> 0 Then
  
    Select Nvl(Max(����), '����') Into v_���� From ���㷽ʽ Where Nvl(����, 0) = 9;
  
    Update ����Ԥ����¼
    Set ��Ԥ�� = Nvl(��Ԥ��, 0) + Nvl(n_�����, 0)
    Where ����id = ����id_In And ���㷽ʽ = v_����;
  
    If Sql%NotFound Then
      Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
    
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, У�Ա�־, �����id, ���㿨���, ����,
         ������ˮ��, ����˵��, �������, ��������, ��������id, �Ự��)
      Values
        (n_Ԥ��id, 2, r_Balancedata.No, 1, r_Balancedata.����id, r_Balancedata.����id, r_Balancedata.��ҳid, Null, v_����,
         r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���, r_Balancedata.����Ա����, n_�����, r_Balancedata.����id, r_Balancedata.�ɿ���id,
         2, Null, Null, ����_In, ������ˮ��_In, ����˵��_In, Null, 2, n_Ԥ��id, v_�Ự��);
    End If;
  
    Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - Nvl(n_�����, 0) Where ����id = ����id_In And ���㷽ʽ Is Null;
    If Sql%NotFound Then
      v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�';
      Raise Err_Item;
    End If;
  End If;

  --3.�����Ԥ��
  If Nvl(��Ԥ��_In, 0) <> 0 Then
    If Nvl(����id_In, 0) = 0 Then
      v_Err_Msg := '����ȷ�����˵Ĳ���ID,����ʹ��Ԥ�������,�������ʧ�ܣ�';
      Raise Err_Item;
    End If;
    Zl_����Ԥ����¼_��Ԥ��(����id_In, ����id_In, ��Ԥ��_In, ��������_In, r_Balancedata.����Ա���, r_Balancedata.����Ա����, r_Balancedata.�տ�ʱ��,
                  ��Ԥ������ids_In, 2, 1);
  End If;

  n_Ԥ��id     := Ԥ��id_In;
  n_��������id := ��������id_In;

  --4.������ͨ������Ϣ
  If ��������_In = 0 Then
  
    If Nvl(��֧Ʊ��_In, 0) <> 0 Then
      --0-��ͨ�շѷ�ʽ����֧Ʊ
      Select Max(b.����)
      Into v_��֧Ʊ
      From ���㷽ʽӦ�� A, ���㷽ʽ B
      Where a.Ӧ�ó��� = '����' And b.���� = a.���㷽ʽ And Nvl(b.Ӧ����, 0) = 1;
    
      If v_��֧Ʊ Is Null Then
        v_Err_Msg := '�ڽ��㳡����,�����ڽ�������ΪӦ����Ľ��㷽ʽ,����[���㷽ʽ]�����ã�';
        Raise Err_Item;
      End If;
      Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
    
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, У�Ա�־, �����id, ���㿨���, ����,
         ������ˮ��, ����˵��, �������, ��������, ��������id, �Ự��)
      Values
        (n_Ԥ��id, 2, r_Balancedata.No, 1, r_Balancedata.����id, r_Balancedata.����id, r_Balancedata.��ҳid, Null, v_��֧Ʊ,
         r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���, r_Balancedata.����Ա����, ��֧Ʊ��_In, r_Balancedata.����id, r_Balancedata.�ɿ���id,
         У�Ա�־_In, Null, Null, ����_In, ������ˮ��_In, ����˵��_In, Null, 2, n_Ԥ��id, v_�Ự��);
    
      Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - ��֧Ʊ��_In Where ����id = ����id_In And ���㷽ʽ Is Null;
      If Sql%NotFound Then
        v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�';
        Raise Err_Item;
      End If;
    
    End If;
  
    n_Ԥ��id := Ԥ��id_In;
    --�����շѽ��� :��ʽΪ:"���㷽ʽ|������|�������|����ժҪ|����||.."
    v_�������� := ���㷽ʽ_In || '||';
  
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
    
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
    
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
    
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      v_������� := LTrim(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_����     := Null;
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      If Instr(v_��ǰ����, '|') > 0 Then
      
        v_����ժҪ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
        v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
        v_����     := LTrim(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
      Else
        v_����ժҪ := v_��ǰ����;
      End If;
    
      If Nvl(n_������, 0) <> 0 Then
        If Nvl(n_Ԥ��id, 0) = 0 Then
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          n_��������id := n_Ԥ��id;
        End If;
      
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, У�Ա�־, �����id, ���㿨���, ����,
           ������ˮ��, ����˵��, �������, ��������, ��������id, �Ự��)
        Values
          (����Ԥ����¼_Id.Nextval, 2, r_Balancedata.No, 1, r_Balancedata.����id, r_Balancedata.����id, r_Balancedata.��ҳid,
           v_����ժҪ, v_���㷽ʽ, r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���, r_Balancedata.����Ա����, n_������, r_Balancedata.����id,
           r_Balancedata.�ɿ���id, У�Ա�־_In, Null, Null, v_����, ������ˮ��_In, ����˵��_In, v_�������, 2, n_��������id, v_�Ự��);
      
        Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - n_������ Where ����id = ����id_In And ���㷽ʽ Is Null;
        If Sql%NotFound Then
          v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�';
          Raise Err_Item;
        End If;
      End If;
      n_Ԥ��id   := Null;
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  --5.���������㽻��
  If ��������_In = 1 Then
    v_��ǰ���� := ���㷽ʽ_In;
    v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
    v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
    n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
    v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
    v_������� := LTrim(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
    v_����ժҪ := LTrim(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
  
    If Nvl(n_������, 0) <> 0 Then
    
      If Nvl(n_Ԥ��id, 0) = 0 Then
        Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      End If;
    
      n_У�Ա�־ := У�Ա�־_In;
      If Nvl(n_��������id, 0) = 0 Then
        n_��������id := n_Ԥ��id;
      Else
        Select Count(1), Max(a.�Ƿ�δ��), -1 * Nvl(Sum(a.���), 0)
        Into n_Count, n_�Ƿ�δ��, n_�˿���
        From �����˿���Ϣ A, ����Ԥ����¼ B
        Where a.��¼id = b.Id And a.����id = r_Balancedata.����id And b.��������id = n_��������id;
        If n_Count > 1 Then
          --Ԥ�������˿�ʱ����������ID��ͬ����Ҫ�ϲ���ע��ֻҪ��һ�ʻ�û����У�Ա�־����1
          If Nvl(n_�Ƿ�δ��, 0) = 1 Then
            n_У�Ա�־ := 1;
          End If;
          n_������ := n_�˿���;
        End If;
      End If;
    
      v_������Ա := zl_UserName;
    
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, У�Ա�־, �����id, ���㿨���, ����,
         ������ˮ��, ����˵��, �������, ��������, ��������id, ���ӱ�־, ����ʱ��, ������Ա, �Ự��)
      Values
        (n_Ԥ��id, 2, r_Balancedata.No, 1, r_Balancedata.����id, r_Balancedata.����id, r_Balancedata.��ҳid, v_����ժҪ, v_���㷽ʽ,
         r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���, r_Balancedata.����Ա����, n_������, r_Balancedata.����id, r_Balancedata.�ɿ���id,
         n_У�Ա�־, �����id_In, Null, ����_In, ������ˮ��_In, ����˵��_In, v_�������, 2, n_��������id, ���ӱ�־_In, Sysdate, v_������Ա, v_�Ự��);
    
      Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - n_������ Where ����id = ����id_In And ���㷽ʽ Is Null;
    
      If Sql%NotFound Then
        v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�';
        Raise Err_Item;
      End If;
    
      --��������������Ϣ����
      Zl_Custom_Balance_Update(n_Ԥ��id);
    
    End If;
  
  End If;

  --6.ҽ������(���ô˹���,��ȡƽ����̯�ķ�ʽ��̯�������):�������ҽ���ᴦ��,����ȫ��
  If ��������_In = 2 Then
  
    --2.1����Ƿ��Ѿ�����ҽ����������,������ɾ��
    n_������ := 0;
    For v_ҽ�� In (Select ID, ���㷽ʽ, ��Ԥ��
                 From ����Ԥ����¼ A
                 Where ����id = ����id_In And �����id Is Null And Exists
                  (Select 1 From ���㷽ʽ Where a.���㷽ʽ = ���� And ���� In (3, 4)) And Mod(��¼����, 10) <> 1) Loop
      n_������ := n_������ + Nvl(v_ҽ��.��Ԥ��, 0);
      l_Ԥ��id.Extend;
      l_Ԥ��id(l_Ԥ��id.Count) := v_ҽ��.Id;
    End Loop;
  
    If Nvl(n_������, 0) <> 0 Then
      Update ����Ԥ����¼
      Set ��Ԥ�� = Nvl(��Ԥ��, 0) + Nvl(n_������, 0)
      Where ����id = ����id_In And ���㷽ʽ Is Null;
      If Sql%NotFound Then
        v_Err_Msg := ' ������Ϣ����, ������Ϊ����ԭ����ɽ�����Ϣ����, ���� [ �շѽ��㴰�� ] �������շѣ� ';
        Raise Err_Item;
      End If;
    End If;
  
    If l_Ԥ��id.Count <> 0 Then
      Forall I In 1 .. l_Ԥ��id.Count
        Delete ����Ԥ����¼ Where ID = l_Ԥ��id(I);
    End If;
  
    If ���㷽ʽ_In Is Not Null Then
      v_�������� := ���㷽ʽ_In || '||';
    End If;
    n_Ԥ��id     := Ԥ��id_In;
    n_��������id := ��������id_In;
  
    d_����ʱ�� := Sysdate;
    v_������Ա := zl_UserName;
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      n_������ := To_Number(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
    
      If n_������ <> 0 Then
      
        If Nvl(n_Ԥ��id, 0) = 0 Then
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
        End If;
      
        If Nvl(n_��������id, 0) = 0 Then
          n_��������id := n_Ԥ��id;
        End If;
      
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, У�Ա�־, ��������, ��������id,
           ����ʱ��, ������Ա, �Ự��)
        Values
          (n_Ԥ��id, 2, r_Balancedata.No, 1, r_Balancedata.����id, r_Balancedata.����id, r_Balancedata.��ҳid, ' ���ս��� ', v_���㷽ʽ,
           r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���, r_Balancedata.����Ա����, n_������, r_Balancedata.����id, r_Balancedata.�ɿ���id,
           У�Ա�־_In, 2, n_��������id, d_����ʱ��, v_������Ա, v_�Ự��);
      
        --��������(���㷽ʽΪNULL��)
        Update ����Ԥ����¼
        Set ��Ԥ�� = ��Ԥ�� - n_������
        Where ����id = ����id_In And ���㷽ʽ Is Null
        Returning Nvl(��Ԥ��, 0) Into n_����ֵ;
        n_Ԥ��id := Null;
      End If;
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  
    --ҽ����ر�Ĵ���
    Update ���ս�����ϸ Set ��־ = 2 Where ����id = ����id_In;
  
  End If;

  --7-���ѿ���������
  If ��������_In = 3 Then
    v_��������   := ���㷽ʽ_In || '||';
    n_Ԥ��id     := Ԥ��id_In;
    n_��������id := ��������id_In;
  
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      --�����ID|����|���ѿ�ID|���ѽ��
      n_�����id := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      v_����     := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_���ѿ�id := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_������ := To_Number(v_��ǰ����);
    
      Select Max(����), Max(���㷽ʽ) Into v_����, v_���㷽ʽ From ���ѿ����Ŀ¼ Where ��� = n_�����id;
      If v_���� Is Null Then
        v_Err_Msg := ' δ�ҵ���Ӧ�Ľ��㿨�ӿ�, ����ˢ������ʧ�� ! ';
        Raise Err_Item;
      End If;
      If v_���㷽ʽ Is Null Then
        v_Err_Msg := v_���� || ' δ���ö�Ӧ�Ľ��㷽ʽ, ����ˢ������ʧ�� ! ';
        Raise Err_Item;
      End If;
    
      If Nvl(n_������, 0) <> 0 Then
        Update ����Ԥ����¼
        Set ��Ԥ�� = Nvl(��Ԥ��, 0) + n_������
        Where ��¼���� = 2 And ����id = r_Balancedata. ����id And ���㷽ʽ = v_���㷽ʽ And ���㿨��� = n_�����id
        Returning ID Into n_Ԥ��id;
        If Sql%NotFound Then
        
          If Nvl(n_Ԥ��id, 0) = 0 Then
            Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          End If;
        
          If Nvl(n_��������id, 0) = 0 Then
            n_��������id := n_Ԥ��id;
          End If;
        
          Insert Into ����Ԥ����¼
            (ID, ��¼����, NO, ��¼״̬, ����id, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ���㿨���, У�Ա�־, ��������,
             ��������id, �Ự��)
          Values
            (n_Ԥ��id, 2, r_Balancedata.No, 1, r_Balancedata. ����id, r_Balancedata.����id, r_Balancedata.��ҳid, Null, v_���㷽ʽ,
             r_Balancedata. �տ�ʱ��, r_Balancedata. ����Ա���, r_Balancedata. ����Ա����, n_������, r_Balancedata. ����id,
             r_Balancedata. �ɿ���id, n_�����id, У�Ա�־_In, 2, n_��������id, v_�Ự��);
        End If;
      
        Zl_���˿������¼_֧��(n_�����id, v_����, n_���ѿ�id, n_������, n_Ԥ��id, r_Balancedata. ����Ա���, r_Balancedata. ����Ա����,
                      r_Balancedata. �տ�ʱ��);
      
        --��������(���㷽ʽΪNULL��)
        Update ����Ԥ����¼
        Set ��Ԥ�� = ��Ԥ�� - n_������
        Where ����id = r_Balancedata. ����id And ���㷽ʽ Is Null And Nvl(У�Ա�־, 0) = 1
        Returning Nvl(��Ԥ��, 0) Into n_����ֵ;
        n_Ԥ��id := Null;
      End If;
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  If Nvl(��ɽ���_In, 0) = 0 Then
    Close c_Balancedata;
    Return;
  End If;

  -----------------------------------------------------------------------------------------
  --����շ�,��Ҫ������Ա�ɿ����,Ԥ����¼(���㷽ʽ=NULL)

  --�������δ�˿�ļ�¼�����²���Ԥ����¼��У�Ա�־
  --���ڶ��Ԥ�����������ID��ͬʱ���ϲ�Ϊ��һ������Ԥ����¼��ֻ��ȫ���ɹ�ʱУ�Ա�־�Ÿ���Ϊ2��������1
  Delete �����˿���Ϣ Where Nvl(�Ƿ�δ��, 0) = 1 And ����id = ����id_In;
  For c_��¼ In (Select ID, ��������id, ��Ԥ��, �����id
               From ����Ԥ����¼
               Where ��¼���� = 2 And ��Ԥ�� < 0 And ����id = ����id_In And �����id Is Not Null And У�Ա�־ = 1) Loop
  
    Select -1 * Nvl(Sum(a.���), 0)
    Into n_������
    From �����˿���Ϣ A, ����Ԥ����¼ B
    Where a.��¼id = b.Id And a.����id = ����id_In And b.��������id = c_��¼.��������id And a.�����id = c_��¼.�����id;
  
    Update ����Ԥ����¼ Set ��Ԥ�� = n_������, У�Ա�־ = 2 Where ID = c_��¼.Id;
    Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� + c_��¼.��Ԥ�� - n_������ Where ����id = ����id_In And ���㷽ʽ Is Null;
  End Loop;

  --1.ɾ�����㷽ʽΪNULL��Ԥ����¼
  Delete ����Ԥ����¼ Where ����id = ����id_In And ���㷽ʽ Is Null And Nvl(��Ԥ��, 0) = 0;
  If Sql%NotFound Then
    Select Count(*) Into n_Count From ����Ԥ����¼ A Where ����id = ����id_In And ���㷽ʽ Is Null;
    If n_Count <> 0 Then
      v_Err_Msg := ' ������δ�ɿ������, ������ɽ��� ! ';
    Else
      v_Err_Msg := ' ������Ϣ����, ������Ϊ����ԭ����ɽ�����Ϣ����, ���� [ �շѽ��㴰�� ] �������շ�! ';
    End If;
    Raise Err_Item;
  End If;

  Select Count(*), Max(b.����)
  Into n_Count, v_Err_Msg
  From ����Ԥ����¼ A, ҽ�ƿ���� B
  Where a.�����id = b.Id And a.����id = ����id_In And a.�����id Is Not Null And a.������ˮ�� Is Null;
  --����������Ҫ���������Ϸ���
  If n_Count <> 0 Then
    v_Err_Msg := v_Err_Msg || '�޽�����ˮ�ţ������׳ɹ�,����ϵͳ����Ա��ϵ!';
    Raise Err_Item;
  End If;

  --������Ϊ��ʱ������һ�����Ϊ0�Ĳ���Ԥ����¼
  Select Count(*) Into n_Count From ����Ԥ����¼ A Where ����id = ����id_In;

  If n_Count = 0 Then
    v_���㷽ʽ := ȱʡ���㷽ʽ_In;
    If v_���㷽ʽ Is Null Then
      Select Max(���㷽ʽ) Into v_���㷽ʽ From ���㷽ʽӦ�� Where Ӧ�ó��� = '����' And Nvl(ȱʡ��־, 0) = 1;
      If v_���㷽ʽ Is Null Then
        Select Nvl(Max(����), ' �ֽ� ') Into v_���㷽ʽ From ���㷽ʽ Where Nvl(����, 0) = 1;
      End If;
    End If;
    If Nvl(n_Ԥ��id, 0) = 0 Then
      Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
    End If;
  
    If Nvl(n_��������id, 0) = 0 Then
      n_��������id := n_Ԥ��id;
    End If;
  
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, У�Ա�־, �����id, ���㿨���, ����,
       ������ˮ��, ����˵��, �������, ��������, ��������id, �Ự��)
    Values
      (n_Ԥ��id, 2, Null, 1, r_Balancedata.����id, r_Balancedata.����id, Null, Null, v_���㷽ʽ, r_Balancedata.�տ�ʱ��,
       r_Balancedata.����Ա���, r_Balancedata.����Ա����, 0, r_Balancedata.����id, r_Balancedata.�ɿ���id, 2, Null, Null, Null, Null,
       ����˵��_In, Null, 2, n_��������id, v_�Ự��);
  End If;
  n_�Ƿ����Ʊ�� := �Ƿ����Ʊ��_In;
  If �Ƿ����Ʊ��_In Is Null Then
    Select Nvl(Max(����), 0) Into n_���� From ���ս����¼ Where ��¼id = ����id_In And ���� = 2;
    n_�Ƿ����Ʊ�� := Zl_Fun_Isstarteinvoice(3, n_����);
  End If;

  --2.����ɿ����ݺ��Ҳ����ݼ�У�Ա�־����Ϊ0
  Update ����Ԥ����¼
  Set �ɿ� = �ɿ�_In, �Ҳ� = �Ҳ�_In, У�Ա�־ = 0, �Ự�� = Null, �Ƿ����Ʊ�� = n_�Ƿ����Ʊ��
  Where ����id = ����id_In;

  --3.���·���״̬
  Update ���˽��ʼ�¼ Set ����״̬ = Null,�Ƿ����Ʊ�� = n_�Ƿ����Ʊ�� Where ID = ����id_In;

  --4.������Ա�ɿ�����
  For c_�ɿ� In (Select ���㷽ʽ, ����Ա����, Sum(Nvl(a.��Ԥ��, 0)) As ��Ԥ��
               From ����Ԥ����¼ A
               Where a.����id = ����id_In And Mod(a.��¼����, 10) <> 1
               Group By ���㷽ʽ, ����Ա����) Loop
  
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + Nvl(c_�ɿ�.��Ԥ��, 0)
    Where �տ�Ա = c_�ɿ�.����Ա���� And ���� = 1 And ���㷽ʽ = c_�ɿ�.���㷽ʽ;
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ����
        (�տ�Ա, ���㷽ʽ, ����, ���)
      Values
        (c_�ɿ�.����Ա����, c_�ɿ�.���㷽ʽ, 1, Nvl(c_�ɿ�.��Ԥ��, 0));
    End If;
  End Loop;
  Close c_Balancedata;

  --5.���ü�¼��Ԥ����¼ƥ����
  Select Nvl(Sum(���ý��), 0), Nvl(Sum(���ʽ��), 0)
  Into n_���ý��, n_���ʽ��
  From (Select Nvl(Sum(���ʽ��), 0) As ���ý��, 0 As ���ʽ��
         From ������ü�¼
         Where ����id = ����id_In
         Union All
         Select Nvl(Sum(���ʽ��), 0) As ���ý��, 0 As ���ʽ��
         From סԺ���ü�¼
         Where ����id = ����id_In
         Union All
         Select 0 As ���ý��, Nvl(Sum(��Ԥ��), 0) As ���ʽ��
         From ����Ԥ����¼
         Where ����id = ����id_In);

  If Nvl(n_���ý��, 0) <> Nvl(n_���ʽ��, 0) Then
    v_Err_Msg := ' ������Ϣ�������Ϣ��ƥ��, �޷���ɽ��� ! ';
    Raise Err_Item;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˽��ʽ���_Modify;
/

Create Or Replace Procedure Zl_���˽�������_Modify
(
  ��������_In      Number,
  ����id_In        ���˽��ʼ�¼.����id%Type,
  ����id_In        ����Ԥ����¼.����id%Type,
  ���㷽ʽ_In      Varchar2,
  �����id_In      ����Ԥ����¼.�����id%Type := Null,
  ����_In          ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In      ����Ԥ����¼.����˵��%Type := Null,
  �ɿ�_In          ����Ԥ����¼.�ɿ�%Type := Null,
  �Ҳ�_In          ����Ԥ����¼.�Ҳ�%Type := Null,
  �����_In      ����Ԥ����¼.��Ԥ��%Type := Null,
  Ԥ�����_In      ����Ԥ����¼.��Ԥ��%Type := Null,
  ����Ա���_In    ����Ԥ����¼.����Ա���%Type := Null,
  ����Ա����_In    ����Ԥ����¼.����Ա����%Type := Null,
  �տ�ʱ��_In      ����Ԥ����¼.�տ�ʱ��%Type := Null,
  ��Ԥ������ids_In Varchar2 := Null,
  �������_In      Number := 0,
  У�Ա�־_In      Number := 0,
  ��������id_In    ����Ԥ����¼.Id%Type := Null,
  ���ԭ����_In    Number := 0,
  Ԥ��id_In        ����Ԥ����¼.Id%Type := Null,
  ����Ự_In      Number := 1
) As
  ------------------------------------------------------------------------------------------------------------------------------
  --����:�շѽ���ʱ,�޸Ľ���������Ϣ
  --��������_In:
  --   0-ԭ����:
  --       ��������������
  --   1-��ͨ�˷ѷ�ʽ:
  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
  --   2.�������˷ѽ���:
  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
  --     �ڿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
  --     ��������ID_IN:
  --     ���ԭ����_In:1-��ʾ�ڸ�������ǰ�����ԭ���Ľ�����Ϣ(������ID+��������ID�����);0-��ʾ�����
  --   3-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
  --   4-���ѿ�����:
  --     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."
  -- Ԥ�����_In:����漰Ԥ����,���뱾�ε���Ԥ�����Ԥ����� ������<0ʱ ��ʾ��Ԥ����;>0 ʱ:��ʾ��Ԥ����
  -- �����_In:��������ʱ,����
  -- ����˷�_In:0-δ����˷�;1-�쳣����˷�;2-����˷�
  -- ��Ԥ������ids_In:����ö��ַ���,��Ԥ��ʱ��Ч(��Ԥ�������ҵ���������һ��),��Ҫ��ʹ�ü�����Ԥ����
  -- У�Ա�־_In:0-��ɻ���ҪУ��;1-��ҪУ��;2-�ӿ��Ѿ����óɹ�
  --����Ự_In��1-��ʾ�ӷŻỰ��0-��ʾ������Ự
  ------------------------------------------------------------------------------------------------------------------------------
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_�������� Varchar2(500);
  v_��ǰ���� Varchar2(500);
  v_����     ����ҽ�ƿ���Ϣ.����%Type;
  n_���ѿ�id ���ѿ���Ϣ.Id%Type;
  n_�����id ����Ԥ����¼.���㿨���%Type;
  v_����     Varchar2(100);
  n_Ԥ��id   ����Ԥ����¼.Id%Type;
  v_���㷽ʽ ����Ԥ����¼.���㷽ʽ%Type;
  n_������ ����Ԥ����¼.��Ԥ��%Type;
  n_����ֵ   ��Ա�ɿ����.���%Type;
  v_������� ����Ԥ����¼.�������%Type;
  v_����ժҪ ����Ԥ����¼.ժҪ%Type;
  v_����   ���㷽ʽ.����%Type;
  n_�ɿ���id ����Ԥ����¼.�ɿ���id%Type;
  n_�쳣���� Number(3);
  n_У�Ա�־ ����Ԥ����¼.У�Ա�־%Type;

  d_����ʱ�� ����Ԥ����¼.����ʱ��%Type;
  v_������Ա ����Ԥ����¼.������Ա%Type;
  n_Dec      Number; --���С��λ��

  n_Count        Number;
  n_Havenull     Number;
  l_Ԥ��id       t_NumList := t_NumList();
  n_ԭ����id     ����Ԥ����¼.����id%Type;
  n_����id       ����Ԥ����¼.����id%Type;
  n_ԭԤ��id     ����Ԥ����¼.Id%Type;
  n_��ֵid       ����Ԥ����¼.Id%Type;
  v_�Ự��       ����Ԥ����¼.�Ự��%Type; --��ʽ��SID+'_'+SERIAL#
  n_�Ƿ����Ʊ�� ����Ԥ����¼.�Ƿ����Ʊ��%Type;

  Cursor c_Balance_Record Is
    Select Max(NO) As NO, Max(m.����id) As ����id, Max(Nvl(�տ�ʱ��_In, m.�շ�ʱ��)) As �Ǽ�ʱ��, Max(Nvl(����Ա���_In, m.����Ա���)) As ����Ա���,
           Max(Nvl(����Ա����_In, m.����Ա����)) As ����Ա����, Sum(���ʽ��) As ������, Max(Nvl(n_�ɿ���id, m.�ɿ���id)) As �ɿ���id,
           Max(��������) As ��������
    From ���˽��ʼ�¼ M
    Where ID = ����id_In;
  r_Balance_Record c_Balance_Record%RowType;

  Cursor c_Balance_Data Is
    Select ��¼����, NO, ��¼״̬, ����id, ����id, ��ҳid, ���㷽ʽ, Nvl(�տ�ʱ��_In, �տ�ʱ��) As �տ�ʱ��, Nvl(����Ա���_In, ����Ա���) As ����Ա���,
           Nvl(����Ա����_In, ����Ա����) As ����Ա����, ��Ԥ��, ����id, Nvl(n_�ɿ���id, �ɿ���id) As �ɿ���id, �������
    From ����Ԥ����¼
    Where ����id = ����id_In And ���㷽ʽ Is Null;
  r_Balance_Data c_Balance_Data%RowType;
  n_����       ����Ԥ����¼.��Ԥ��%Type;

Begin

  If ����Ա����_In Is Not Null Then
    n_�ɿ���id := Zl_Get��id(����Ա����_In);
  End If;

  If Nvl(����Ự_In, 0) = 1 Then
    Select Max(Sid || '_' || Serial#) Into v_�Ự�� From V$session Where Audsid = Userenv('sessionid');
  End If;

  Open c_Balance_Record;
  Fetch c_Balance_Record
    Into r_Balance_Record;

  Open c_Balance_Data;
  Fetch c_Balance_Data
    Into r_Balance_Data;

  If r_Balance_Record.No Is Null Then
    v_Err_Msg := 'δ�ҵ�ָ���Ľ������ϼ�¼��';
    Raise Err_Item;
  End If;

  If ��������_In = 0 Then
    --ԭ������
    Select Max(ID) Into n_ԭ����id From ���˽��ʼ�¼ Where ��¼״̬ In (1, 3) And NO = r_Balance_Data.No;
    If n_ԭ����id Is Null Then
      v_Err_Msg := 'û�з���Ҫ���ϵĽ��ʵ���,�����Ѿ����ϣ�';
      Raise Err_Item;
    End If;
  
    Delete ����Ԥ����¼ Where ����id = ����id_In And ���㷽ʽ Is Null;
    -- ��Ԥ��
    If Nvl(Ԥ�����_In, 0) <> 0 Then
    
      Zl_���˽���Ԥ��_Cancel(����id_In, n_ԭ����id, ����id_In, -1 * Ԥ�����_In, r_Balance_Data.����Ա���, r_Balance_Data.����Ա����,
                       r_Balance_Data.�տ�ʱ��, r_Balance_Record.�ɿ���id);
    Else
    
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ����id, ��ҳid, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, ��������, ��������id, �Ự��, �Ƿ����Ʊ��)
        Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ����id, ��ҳid, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�,
               r_Balance_Record.�Ǽ�ʱ��, r_Balance_Record.����Ա����, r_Balance_Record.����Ա���, -1 * ��Ԥ��, ����id_In,
               r_Balance_Record.�ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 2, 2, ��������id, v_�Ự��, �Ƿ����Ʊ��
        From ����Ԥ����¼
        Where ����id = n_ԭ����id And ��¼���� In (1, 11) And Nvl(��Ԥ��, 0) <> 0;
    
      For r_Ԥ�� In (Select ����id, NO, Ԥ�����, Sum(��Ԥ��) As ��Ԥ��, Max(Decode(��¼����, 1, Decode(��¼״̬, 2, 0, ID), 0)) As Ԥ��id
                   From ����Ԥ����¼
                   Where ����id = ����id_In And Mod(��¼����, 10) = 1
                   Group By ����id, NO, Ԥ�����) Loop
        --�������(Ԥ��)
        Update �������
        Set Ԥ����� = Nvl(Ԥ�����, 0) - Nvl(r_Ԥ��.��Ԥ��, 0) --ע:�µĽ���ID�������Ǹ������
        Where ����id = r_Ԥ��.����id And ���� = Nvl(r_Ԥ��.Ԥ�����, 2) And ���� = 1
        Returning Ԥ����� Into n_����ֵ;
      
        If Sql%RowCount = 0 Then
          Insert Into �������
            (����id, ����, ����, Ԥ�����, �������)
          Values
            (r_Ԥ��.����id, 1, Nvl(r_Ԥ��.Ԥ�����, 2), -1 * r_Ԥ��.��Ԥ��, 0);
          n_����ֵ := -1 * r_Ԥ��.��Ԥ��;
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete ������� Where ���� = 1 And ����id = r_Ԥ��.����id And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
        End If;
      
        --����Ԥ���������
        n_��ֵid := Nvl(r_Ԥ��.Ԥ��id, 0);
        If n_��ֵid = 0 Then
          Select Max(ID) Into n_��ֵid From ����Ԥ����¼ Where NO = r_Ԥ��.No And ��¼���� = 1 And ��¼״̬ <> 2;
        End If;
        If n_��ֵid <> 0 Then
          Update Ԥ���������
          Set Ԥ����� = Nvl(Ԥ�����, 0) - Nvl(r_Ԥ��.��Ԥ��, 0)
          Where ����id = r_Ԥ��.����id And Ԥ��id = n_��ֵid
          Returning Ԥ����� Into n_����ֵ;
        
          If Sql%RowCount = 0 Then
            Insert Into Ԥ���������
              (Ԥ��id, ����id, Ԥ�����, Ԥ�����)
            Values
              (n_��ֵid, r_Ԥ��.����id, Nvl(r_Ԥ��.Ԥ�����, 2), Nvl(-1 * r_Ԥ��.��Ԥ��, 0));
            n_����ֵ := -1 * Nvl(r_Ԥ��.��Ԥ��, 0);
          End If;
        
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From Ԥ��������� Where Ԥ��id = n_��ֵid And Nvl(Ԥ�����, 0) = 0;
          End If;
        
        End If;
      End Loop;
    
    End If;
  
    --������
    Insert Into ����Ԥ����¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��, ����id,
       �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, ��������, ��������id, ����ʱ��, ������Ա, �Ự��)
      Select ����Ԥ����¼_Id.Nextval, a.No, ʵ��Ʊ��, 12, a.��¼״̬, a.����id, a.��ҳid, a.����id, Null, a.���㷽ʽ, a.�������, a.ժҪ, a.�ɿλ,
             a.��λ������, a.��λ�ʺ�, r_Balance_Record.�Ǽ�ʱ��, r_Balance_Record.����Ա����, r_Balance_Record.����Ա���, -1 * ��Ԥ��, ����id_In,
             r_Balance_Record.�ɿ���id, a.Ԥ�����, a.�����id, a.���㿨���, a.����, a.������ˮ��, a.����˵��, a.������λ, Nvl(У�Ա�־_In, 0), 2,
             a.��������id,
             Decode(a.�����id, Null, Decode(Nvl(b.����, 0), 3, r_Balance_Record.�Ǽ�ʱ��, 4, r_Balance_Record.�Ǽ�ʱ��, Null),
                     r_Balance_Record.�Ǽ�ʱ��),
             Decode(a.�����id, Null, Decode(Nvl(b.����, 0), 3, r_Balance_Record.����Ա����, 4, r_Balance_Record.����Ա����, Null),
                     r_Balance_Record.����Ա����), v_�Ự��
      From ����Ԥ����¼ A, ���㷽ʽ B
      Where ����id = n_ԭ����id And a.���㷽ʽ = b.����(+) And Mod(��¼����, 10) <> 1 And Nvl(��Ԥ��, 0) >= 0;
  
    Update ����Ԥ����¼ Set �ɿ� = �ɿ�_In, �Ҳ� = �Ҳ�_In, У�Ա�־ = 0 Where ����id = ����id_In;
  
    Select Count(1)
    Into n_�쳣����
    From ���˽��ʼ�¼
    Where NO = r_Balance_Data.No And ��¼״̬ = 3 And ����״̬ = 1 And Rownum < 2;
  
    If Nvl(У�Ա�־_In, 0) = 0 And Nvl(n_�쳣����, 0) <> 0 Then
      Update ���˽��ʼ�¼
      Set ����״̬ = Decode(n_�쳣����, 1, 2, Null)
      Where NO = r_Balance_Data.No And ����״̬ Is Not Null;
    
      Update ����Ԥ����¼ Set �Ự�� = Null Where ����id = ����id_In And �Ự�� Is Not Null;
    End If;
    Close c_Balance_Record;
  
    --��Ҫ����������������Ϣ���¹���
    For c_�������� In (Select ID From ����Ԥ����¼ Where ����id = ����id_In And �����id Is Not Null Order By �����id) Loop
      --��������������Ϣ����
      Zl_Custom_Balance_Update(c_��������.Id);
    End Loop;
  
    Return;
  End If;

  --0.��ʽ����
  Select Count(1), Max(Decode(���㷽ʽ, Null, 1, 0))
  Into n_Count, n_Havenull
  From ����Ԥ����¼
  Where ����id = ����id_In;

  --���С��λ��
  Select zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  --1.���ӽ��㷽ʽΪ�յĽ�������
  n_���� := �����_In;

  If Nvl(n_Havenull, 0) = 0 Then
    n_Count := 0;
    Select Sum(���ʽ��)
    Into n_������
    From (Select Sum(���ʽ��) As ���ʽ��
           From ������ü�¼
           Where ����id = ����id_In
           Union All
           Select Sum(���ʽ��)
           From סԺ���ü�¼
           Where ����id = ����id_In);
  
    Begin
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, У�Ա�־, ��������, �Ự��)
      Values
        (����Ԥ����¼_Id.Nextval, 12, r_Balance_Record.No, 1, r_Balance_Record.����id, Null, r_Balance_Record.�Ǽ�ʱ��,
         r_Balance_Record.����Ա���, r_Balance_Record.����Ա����, n_������, ����id_In, r_Balance_Record.�ɿ���id, 2, 2, v_�Ự��);
    Exception
      When Others Then
        n_Count := -1;
    End;
  
    If n_Count = -1 Then
      v_Err_Msg := 'δ�ҵ�ָ���Ľ�����������,�˷Ѳ���ʧ�ܣ�';
      Raise Err_Item;
    End If;
  End If;

  --��������
  If Nvl(n_����, 0) <> 0 Then
    Select Nvl(Max(����), '����') Into v_���� From ���㷽ʽ Where Nvl(����, 0) = 9;
  
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, У�Ա�־, ��������, �Ự��)
    Values
      (����Ԥ����¼_Id.Nextval, 12, r_Balance_Record.No, 1, r_Balance_Record.����id, v_����, r_Balance_Record.�Ǽ�ʱ��,
       r_Balance_Record.����Ա���, r_Balance_Record.����Ա����, n_����, ����id_In, r_Balance_Record.�ɿ���id, 2, 2, v_�Ự��);
  
    --��������(���㷽ʽΪNULL��)
    Update ����Ԥ����¼
    Set ��Ԥ�� = ��Ԥ�� - Nvl(n_����, 0)
    Where ����id = r_Balance_Data.����id And ���㷽ʽ Is Null
    Returning Nvl(��Ԥ��, 0) Into n_����ֵ;
  
  End If;

  If Nvl(Ԥ�����_In, 0) <> 0 Then
    Select Max(ID) Into n_ԭ����id From ���˽��ʼ�¼ Where ��¼״̬ In (1, 3) And NO = r_Balance_Data.No;
    If n_ԭ����id Is Null Then
      v_Err_Msg := 'û�з���Ҫ���ϵĽ��ʵ���,�����Ѿ����ϣ�';
      Raise Err_Item;
    End If;
  
    Zl_���˽���Ԥ��_Cancel(����id_In, n_ԭ����id, ����id_In, -1 * Ԥ�����_In, r_Balance_Data.����Ա���, r_Balance_Data.����Ա����,
                     r_Balance_Data.�տ�ʱ��, r_Balance_Record.�ɿ���id);
  End If;

  If ��������_In = 1 Then
    --   1-��ͨ�˷ѷ�ʽ:
    --�����շѽ��� :��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.."
    v_�������� := ���㷽ʽ_In || '||';
    n_Ԥ��id   := Ԥ��id_In;
  
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
    
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
    
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
    
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      v_������� := LTrim(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_����ժҪ := LTrim(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
    
      --���жϡ�������Ƿ�Ϊ�㣬�п����Ѿ����꣬����ʱ���㷽ʽΪ�յ��ؽ�ͳ�����¼�ĳ�Ԥ��֮��Ϊ��
      If v_���㷽ʽ Is Not Null Then
        --If Nvl(n_������, 0) <> 0 Then
        n_������ := Nvl(n_������, 0);
        If Nvl(n_������, 0) <> 0 Then
          If Nvl(n_Ԥ��id, 0) = 0 Then
            Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          End If;
        
          Insert Into ����Ԥ����¼
            (ID, ��¼����, NO, ��¼״̬, ����id, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, У�Ա�־, �����id, ���㿨���,
             ����, ������ˮ��, ����˵��, �������, ��������, �Ự��)
          Values
            (n_Ԥ��id, 12, r_Balance_Data.No, 1, r_Balance_Data.����id, r_Balance_Data.����id, r_Balance_Data.��ҳid, v_����ժҪ,
             v_���㷽ʽ, r_Balance_Data.�տ�ʱ��, r_Balance_Data.����Ա���, r_Balance_Data.����Ա����, n_������, r_Balance_Data.����id,
             r_Balance_Data.�ɿ���id, 2, Null, Null, ����_In, ������ˮ��_In, ����˵��_In, Null, 2, v_�Ự��);
        
          Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - n_������ Where ����id = ����id_In And ���㷽ʽ Is Null;
          If Sql%NotFound Then
            v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�';
            Raise Err_Item;
          End If;
          n_Ԥ��id := Null;
        End If;
      End If;
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  If ��������_In = 2 Then
  
    If Nvl(���ԭ����_In, 0) = 1 And Nvl(��������id_In, 0) <> 0 Then
      --��ԭ���㷽ʽΪ�յĽ�����
      --���������Ⲣ������
      Update ����Ԥ����¼
      Set ��Ԥ�� = ��Ԥ��
      Where ����id = ����id_In And ��������id = ��������id_In And Mod(��¼����, 10) <> 1;
    
      Select Sum(��Ԥ��)
      Into n_������
      From ����Ԥ����¼
      Where ����id = ����id_In And ��������id = ��������id_In And Mod(��¼����, 10) <> 1;
    
      Update ����Ԥ����¼
      Set ��Ԥ�� = Nvl(��Ԥ��, 0) + Nvl(n_������, 0)
      Where ����id = ����id_In And ���㷽ʽ Is Null;
    
      Delete ����Ԥ����¼ Where ����id = ����id_In And ��������id = ��������id_In And Mod(��¼����, 10) <> 1;
    End If;
  
    d_����ʱ�� := Sysdate;
    v_������Ա := zl_UserName;
    --   2.�������˷ѽ���:
    v_��ǰ���� := ���㷽ʽ_In;
    v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
    v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
    n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
    v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
    v_������� := LTrim(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
    v_����ժҪ := LTrim(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
  
    If Nvl(n_������, 0) <> 0 Then
      n_Ԥ��id := Ԥ��id_In;
      If Nvl(n_Ԥ��id, 0) = 0 Then
        Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      End If;
    
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, У�Ա�־, �����id, ���㿨���, ����,
         ������ˮ��, ����˵��, �������, ��������, ��������id, ����ʱ��, ������Ա, �Ự��)
      Values
        (n_Ԥ��id, 12, r_Balance_Data.No, 1, r_Balance_Data.����id, r_Balance_Data.����id, r_Balance_Data.��ҳid, v_����ժҪ,
         v_���㷽ʽ, r_Balance_Data.�տ�ʱ��, r_Balance_Data.����Ա���, r_Balance_Data.����Ա����, n_������, r_Balance_Data.����id,
         r_Balance_Data.�ɿ���id, У�Ա�־_In, �����id_In, Null, ����_In, ������ˮ��_In, ����˵��_In, v_�������, 2,
         Decode(Nvl(��������id_In, 0), 0, n_Ԥ��id, ��������id_In), d_����ʱ��, v_������Ա, v_�Ự��);
    
      Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - n_������ Where ����id = ����id_In And ���㷽ʽ Is Null;
      If Sql%NotFound Then
        v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�';
        Raise Err_Item;
      End If;
    
      --��������������Ϣ����
      Zl_Custom_Balance_Update(n_Ԥ��id);
    End If;
  End If;

  If ��������_In = 3 Then
    --   3-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
    --3.1����Ƿ��Ѿ�����ҽ����������,������ɾ��
    n_������ := 0;
  
    If У�Ա�־_In = 0 Then
      n_У�Ա�־ := 2;
    Else
      n_У�Ա�־ := 1;
    End If;
  
    For v_ҽ�� In (Select ID, ���㷽ʽ, ��Ԥ��
                 From ����Ԥ����¼ A
                 Where ����id = ����id_In And �����id Is Null And Exists
                  (Select 1 From ���㷽ʽ Where a.���㷽ʽ = ���� And ���� In (3, 4)) And Mod(��¼����, 10) <> 1) Loop
      n_������ := n_������ + Nvl(v_ҽ��.��Ԥ��, 0);
      l_Ԥ��id.Extend;
      l_Ԥ��id(l_Ԥ��id.Count) := v_ҽ��.Id;
    End Loop;
  
    If Nvl(n_������, 0) <> 0 Then
      Update ����Ԥ����¼
      Set ��Ԥ�� = Nvl(��Ԥ��, 0) + Nvl(n_������, 0)
      Where ����id = ����id_In And ���㷽ʽ Is Null;
      If Sql%NotFound Then
        v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�';
        Raise Err_Item;
      End If;
    End If;
    If l_Ԥ��id.Count <> 0 Then
      Forall I In 1 .. l_Ԥ��id.Count
        Delete ����Ԥ����¼ Where ID = l_Ԥ��id(I);
    End If;
  
    If ���㷽ʽ_In Is Not Null Then
      v_�������� := ���㷽ʽ_In || '||';
    End If;
    d_����ʱ�� := Sysdate;
    v_������Ա := zl_UserName;
    n_Ԥ��id   := Ԥ��id_In;
  
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      n_������ := To_Number(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
    
      For c_������Ϣ In (Select ��¼����, NO, ��¼״̬, ����id, ����id, ��ҳid, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, ��������id
                     From ����Ԥ����¼
                     Where ����id = ����id_In And ���㷽ʽ Is Null) Loop
        If Nvl(n_Ԥ��id, 0) = 0 Then
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
        End If;
      
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, У�Ա�־, ��������, ����ʱ��, ������Ա,
           ��������id, �Ự��)
        Values
          (n_Ԥ��id, 12, c_������Ϣ.No, 1, c_������Ϣ.����id, c_������Ϣ.����id, c_������Ϣ.��ҳid, '���ս���', v_���㷽ʽ, c_������Ϣ.�տ�ʱ��, c_������Ϣ.����Ա���,
           c_������Ϣ.����Ա����, n_������, c_������Ϣ.����id, c_������Ϣ.�ɿ���id, n_У�Ա�־, 2, d_����ʱ��, v_������Ա,
           Decode(Nvl(��������id_In, 0), 0, n_Ԥ��id, ��������id_In), v_�Ự��);
        n_Ԥ��id := Null;
      End Loop;
    
      --��������(���㷽ʽΪNULL��)
      Update ����Ԥ����¼
      Set ��Ԥ�� = ��Ԥ�� - n_������
      Where ����id = ����id_In And ���㷽ʽ Is Null
      Returning Nvl(��Ԥ��, 0) Into n_����ֵ;
    
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  --4-���ѿ���������
  If ��������_In = 4 Then
    v_�������� := ���㷽ʽ_In || '||';
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      --�����ID|����|���ѿ�ID|���ѽ��
      n_�����id := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      v_����     := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_���ѿ�id := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_������ := To_Number(v_��ǰ����);
    
      Begin
        Select ����, ���㷽ʽ Into v_����, v_���㷽ʽ From ���ѿ����Ŀ¼ Where ��� = n_�����id;
      Exception
        When Others Then
          v_���� := Null;
      End;
      If v_���� Is Null Then
        v_Err_Msg := 'δ�ҵ���Ӧ�Ľ��㿨�ӿ�,����ˢ������ʧ��!';
        Raise Err_Item;
      End If;
      If v_���㷽ʽ Is Null Then
        v_Err_Msg := v_���� || 'δ���ö�Ӧ�Ľ��㷽ʽ,����ˢ������ʧ��!';
        Raise Err_Item;
      End If;
      If Nvl(n_������, 0) <> 0 Then
        n_����id := ����id_In;
      
        Update ����Ԥ����¼
        Set ��Ԥ�� = Nvl(��Ԥ��, 0) + n_������
        Where ����id = n_����id And ���㷽ʽ = v_���㷽ʽ And ���㿨��� = n_�����id
        Returning ID Into n_Ԥ��id;
        If Sql%NotFound Then
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (ID, ��¼����, NO, ��¼״̬, ����id, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ���㿨���, У�Ա�־, ��������,
             �Ự��)
          Values
            (n_Ԥ��id, 12, r_Balance_Data.No, 1, r_Balance_Data. ����id, r_Balance_Data.����id, r_Balance_Data.��ҳid, Null,
             v_���㷽ʽ, r_Balance_Data. �տ�ʱ��, r_Balance_Data. ����Ա���, r_Balance_Data. ����Ա����, n_������, n_����id,
             r_Balance_Data. �ɿ���id, n_�����id, 2, 2, v_�Ự��);
        End If;
      
        Begin
          Select b.Id
          Into n_ԭԤ��id
          From ���˽��ʼ�¼ A, ����Ԥ����¼ B
          Where a.Id = b.����id And a.��¼״̬ In (1, 3) And a.No = r_Balance_Data.No And b.���㿨��� = n_�����id;
        Exception
          When Others Then
            Begin
              v_Err_Msg := 'û�з���' || v_���� || '��ԭ�������ݣ�';
              Raise Err_Item;
            End;
        End;
      
        --���뿨�����¼
        Zl_���˿������¼_�˿�(n_�����id, v_����, n_���ѿ�id, -1 * n_������, n_ԭԤ��id, n_Ԥ��id, r_Balance_Data. ����Ա���,
                      r_Balance_Data. ����Ա����, r_Balance_Data. �տ�ʱ��);
      
        --��������(���㷽ʽΪNULL��)
        Update ����Ԥ����¼
        Set ��Ԥ�� = ��Ԥ�� - n_������
        Where ����id = n_����id And ���㷽ʽ Is Null
        Returning Nvl(��Ԥ��, 0) Into n_����ֵ;
      End If;
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  If Nvl(�������_In, 0) = 0 Then
    Return;
  End If;

  -----------------------------------------------------------------------------------------
  --����շ�,��Ҫ������Ա�ɿ����,Ԥ����¼(���㷽ʽ=NULL)
  If Nvl(�������_In, 0) = 1 Then
  
    --�쳣����˷�
    Update ����Ԥ����¼ Set �ɿ� = �ɿ�_In, �Ҳ� = �Ҳ�_In, У�Ա�־ = 0, �Ự�� = Null Where ����id = ����id_In;
  
    Update ���˽��ʼ�¼
    Set ����״̬ = 2
    Where ID In (Select ID From ���˽��ʼ�¼ A Where NO = r_Balance_Data.No) And ����״̬ Is Not Null;
  
    Return;
  End If;

  --1.ɾ�����㷽ʽΪNULL��Ԥ����¼
  Delete ����Ԥ����¼ Where ����id = ����id_In And ���㷽ʽ Is Null And Nvl(��Ԥ��, 0) = 0;
  If Sql%NotFound Then
    Select Count(*) Into n_Count From ����Ԥ����¼ A Where ����id = ����id_In And ���㷽ʽ Is Null;
    If n_Count <> 0 Then
      v_Err_Msg := '������δ�˿������,������ɽ������ϲ���!';
    Else
      v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[���ʴ���]���������ϣ�!';
    End If;
    Raise Err_Item;
  End If;

  --������Ϊ��ʱ������һ�����Ϊ0�Ĳ���Ԥ����¼
  Select Count(*) Into n_Count From ����Ԥ����¼ A Where ����id = ����id_In;

  If n_Count = 0 Then
    If v_���㷽ʽ Is Null Then
      Begin
        Select ���㷽ʽ Into v_���㷽ʽ From ���㷽ʽӦ�� Where Ӧ�ó��� = '����' And Nvl(ȱʡ��־, 0) = 1;
      Exception
        When Others Then
          v_���㷽ʽ := Null;
      End;
      If v_���㷽ʽ Is Null Then
        Begin
          Select ���� Into v_���㷽ʽ From ���㷽ʽ Where Nvl(����, 0) = 1;
        Exception
          When Others Then
            v_���㷽ʽ := '�ֽ�';
        End;
      End If;
    End If;
  
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, У�Ա�־, �����id, ���㿨���, ����,
       ������ˮ��, ����˵��, �������, ��������)
    Values
      (����Ԥ����¼_Id.Nextval, 12, r_Balance_Data.No, 1, r_Balance_Data.����id, r_Balance_Data.����id, r_Balance_Data.��ҳid, Null,
       v_���㷽ʽ, r_Balance_Data.�տ�ʱ��, r_Balance_Data.����Ա���, r_Balance_Data.����Ա����, 0, r_Balance_Data.����id,
       r_Balance_Data.�ɿ���id, 2, Null, Null, Null, Null, ����˵��_In, Null, 2);
  End If;

  --���µ���Ʊ��
  Select Max(�Ƿ����Ʊ��)
  Into n_�Ƿ����Ʊ��
  From ����Ԥ����¼
  Where ����id In (Select ID From ���˽��ʼ�¼ Where ��¼״̬ In (1, 3) And NO = r_Balance_Data.No);

  --2.����ɿ����ݺ��Ҳ����ݼ�У�Ա�־����Ϊ0
  Update ����Ԥ����¼
  Set �ɿ� = �ɿ�_In, �Ҳ� = �Ҳ�_In, У�Ա�־ = 0, �Ự�� = Null, �Ƿ����Ʊ�� = n_�Ƿ����Ʊ��
  Where ����id = ����id_In;

  --3.���·���״̬
  Update ���˽��ʼ�¼ Set ����״̬ = Null, �Ƿ����Ʊ�� = n_�Ƿ����Ʊ�� Where ID = ����id_In;

  --4.Ʊ�ݻ���
  -- ���ܴ��ں�Լ��λ�����˴�ӡ, ���Դ��ڶ���Ʊ��
  For c_Ʊ�� In (Select ID As ��ӡid
               From (Select b.Id
                      From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
                      Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And b.�������� = 3 And b.No = r_Balance_Data.No
                      Order By a.ʹ��ʱ�� Desc)
               Where Rownum < 2) Loop
  
    --�����ջ�Ʊ��(������ǰû��ʹ��Ʊ��,�޷��ջ�)
    If c_Ʊ��.��ӡid Is Not Null Then
      Insert Into Ʊ��ʹ����ϸ
        (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
        Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, r_Balance_Data.�տ�ʱ��, r_Balance_Data.����Ա����, Ʊ�ݽ��
        From Ʊ��ʹ����ϸ A
        Where ��ӡid = c_Ʊ��.��ӡid And Ʊ�� In (1, 3) And ���� = 1 And Not Exists
         (Select 1 From Ʊ��ʹ����ϸ B Where a.���� = b.���� And ��ӡid = c_Ʊ��.��ӡid And ���� = 2);
    End If;
  
  End Loop;

  --5.�����ܶ�Ҫ�������Ϣ����һ��
  Select Sum(��Ԥ��), Sum(���ʽ��)
  Into n_����ֵ, n_������
  From (Select Sum(��Ԥ��) As ��Ԥ��, 0 As ���ʽ��
         From ����Ԥ����¼
         Where ����id = ����id_In
         Union All
         Select 0, Sum(���ʽ��)
         From ������ü�¼
         Where ����id = ����id_In
         Union All
         Select 0, Sum(���ʽ��) As ���ʽ��
         From סԺ���ü�¼
         Where ����id = ����id_In);

  If Nvl(n_����ֵ, 0) <> Nvl(n_������, 0) Then
    v_Err_Msg := '�����ܶ�������ܶһ��,���ܽ������ϲ���������ϵͳ����Ա��ϵ!';
    Raise Err_Item;
  End If;

  --5.������Ա�ɿ�����
  For c_�ɿ� In (Select ���㷽ʽ, ����Ա����, Sum(Nvl(a.��Ԥ��, 0)) As ��Ԥ��
               From ����Ԥ����¼ A
               Where a.����id = ����id_In And Mod(a.��¼����, 10) <> 1
               Group By ���㷽ʽ, ����Ա����) Loop
  
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + Nvl(c_�ɿ�.��Ԥ��, 0)
    Where �տ�Ա = c_�ɿ�.����Ա���� And ���� = 1 And ���㷽ʽ = c_�ɿ�.���㷽ʽ;
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ����
        (�տ�Ա, ���㷽ʽ, ����, ���)
      Values
        (c_�ɿ�.����Ա����, c_�ɿ�.���㷽ʽ, 1, Nvl(c_�ɿ�.��Ԥ��, 0));
    End If;
  End Loop;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˽�������_Modify;
/
 

Create Or Replace Procedure Zl_ҽ�ƿ�����_Modify
(
  ���ݺ�_In       סԺ���ü�¼.No%Type,
  ����id_In       סԺ���ü�¼.����id%Type,
  ���㷽ʽ_In     ����Ԥ����¼.���㷽ʽ%Type := Null,
  ������_In     ����Ԥ����¼.��Ԥ��%Type := 0,
  ��ɱ�־_In     Number := 0,
  �����id_In     ����Ԥ����¼.�����id%Type := Null,
  ���ѿ�_In       Number := 0,
  ����_In         ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In   ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In     ����Ԥ����¼.����˵��%Type := Null,
  ��ͨ����_In     Number := 0,
  �������_In     ����Ԥ����¼.�������%Type := Null,
  ժҪ_In         ����Ԥ����¼.ժҪ%Type := Null,
  У�Ա�־_In     ����Ԥ����¼.У�Ա�־%Type := 2,
  ��������id_In   ����Ԥ����¼.��������id%Type := Null,
  �Ƿ����Ʊ��_In ����Ԥ����¼.�Ƿ����Ʊ��%Type := Null
) As
  ----------------------------------------------------------------------------
  --����:
  --�Ƿ����Ʊ��_In:null-��ʾ�����ڲ�ֱ���жϣ��ǿձ�ʾֱ���Դ����Ϊ׼
  --                ��ע������˷ѣ��ò���ʧЧ)
  ----------------------------------------------------------------------------

  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  n_������     ����Ԥ����¼.��Ԥ��%Type;
  n_��Ԥ��       ����Ԥ����¼.��Ԥ��%Type;
  n_Ԥ��id       ����Ԥ����¼.Id%Type;
  n_�����id     ����Ԥ����¼.�����id%Type;
  n_���ѿ�id     ����Ԥ����¼.�����id%Type;
  v_���㷽ʽ     ����Ԥ����¼.���㷽ʽ%Type;
  n_ԭԤ��id     ����Ԥ����¼.Id%Type;
  n_�������     ����Ԥ����¼.��Ԥ��%Type;
  n_�շ�����     Number;
  n_�˷�         Number;
  n_����         Number;
  n_Count        Number;
  d_Date         ����Ԥ����¼.�տ�ʱ��%Type;
  v_����Ա���   ����Ԥ����¼.����Ա���%Type;
  v_����Ա����   ����Ԥ����¼.����Ա����%Type;
  n_ԭ����id     ����Ԥ����¼.����id%Type;
  n_�Ƿ����Ʊ�� Number(2);
  n_����         ���ս����¼.����%Type;
Begin

  Select Nvl(Max(���ʷ���), 0), Max(����id), Max(Decode(��¼״̬, 3, ����״̬, 0)), Max(Decode(��¼״̬, 3, 1, 0))
  Into n_����, n_ԭ����id, n_�շ�����, n_�˷�
  From סԺ���ü�¼
  Where NO = ���ݺ�_In And ��¼���� = 5 And ��¼״̬ In (1, 3);

  If Nvl(����id_In, 0) = 0 Or n_���� = 1 Then
    Return;
  End If;

  If n_�˷� = 1 Then
    n_������ := -1 * Nvl(������_In, 0);
  Else
    n_������ := Nvl(������_In, 0);
  End If;
  If ��ͨ����_In = 0 And Nvl(�����id_In, 0) <> 0 Then
    If Nvl(���ѿ�_In, 0) = 0 Then
      n_�����id := �����id_In;
    Else
      n_���ѿ�id := �����id_In;
    End If;
  End If;
  If ���㷽ʽ_In Is Null Then
    Update ����Ԥ����¼
    Set У�Ա�־ = У�Ա�־_In
    Where ��¼���� = 5 And ����id = ����id_In And Rownum < 2 Return ID, �����id, �տ�ʱ��, ����Ա���, ����Ա���� Into n_Ԥ��id, n_�����id, d_Date,
     v_����Ա���, v_����Ա����;
  Else
    Update ����Ԥ����¼
    Set ���㷽ʽ = ���㷽ʽ_In, ��Ԥ�� = n_������, У�Ա�־ = У�Ա�־_In, �����id = n_�����id, ���㿨��� = n_���ѿ�id, ���� = ����_In, ������ˮ�� = ������ˮ��_In,
        ����˵�� = ����˵��_In, ������� = �������_In, ժҪ = Nvl(ժҪ_In, ժҪ), ��������id = Nvl(��������id_In, ID)
    Where ��¼���� = 5 And ����id = ����id_In And Rownum < 2 Return ID, �����id, �տ�ʱ��, ����Ա���, ����Ա���� Into n_Ԥ��id, n_�����id, d_Date,
     v_����Ա���, v_����Ա����;
  End If;

  --���������������½ӿ���Ϣ
  If У�Ա�־_In = 2 Then
    If Nvl(n_�����id, 0) <> 0 Then
      Zl_Custom_Balance_Update(n_Ԥ��id);
      Update ����Ԥ����¼
      Set ����ʱ�� = �տ�ʱ��, ������Ա = ����Ա����
      Where ��¼���� = 5 And Nvl(�����id, 0) > 0 And ����id = ����id_In;
    End If;
  
    If Nvl(n_���ѿ�id, 0) <> 0 Then
      If n_�˷� = 0 Then
        Zl_���˿������¼_֧��(n_���ѿ�id, ����_In, 0, n_������, n_Ԥ��id, v_����Ա���, v_����Ա����, d_Date);
      Else
        Select Nvl(ID, 0), -1 * Nvl(��Ԥ��, 0)
        Into n_ԭԤ��id, n_�������
        From ����Ԥ����¼
        Where NO = ���ݺ�_In And ��¼���� = 5 And ��¼״̬ = 3 And ���㷽ʽ = ���㷽ʽ_In And ���㿨��� = �����id_In;
        If n_ԭԤ��id = 0 Then
          v_Err_Msg := 'δ�ҵ�ԭ�����¼��';
          Raise Err_Item;
        End If;
        If n_������� <> n_������ Then
          v_Err_Msg := '���ѿ��˿��һ�£�';
          Raise Err_Item;
        End If;
        Zl_���˿������¼_�˿�(�����id_In, ����_In, 0, -1 * n_������, n_ԭԤ��id, n_Ԥ��id, v_����Ա���, v_����Ա����, d_Date);
      End If;
    End If;
  End If;

  If Nvl(��ɱ�־_In, 0) = 0 Then
    Return;
  End If;

  --1.�ȼ�����Ƿ�һ��
  Select Nvl(Sum(ʵ�ս��), 0) Into n_������ From סԺ���ü�¼ Where ����id = ����id_In;
  Select Nvl(Sum(��Ԥ��), 0), Max(���㷽ʽ) Into n_��Ԥ��, v_���㷽ʽ From ����Ԥ����¼ Where ����id = ����id_In;
  If n_������ <> n_��Ԥ�� Then
    v_Err_Msg := '���ѽ�����Ϣ����ʵ�ս��(' || n_������ || ')�������(' || n_��Ԥ�� || ')��һ�£�������ɽ��㣡';
    Raise Err_Item;
  End If;

  Select Count(1) Into n_Count From ����Ԥ����¼ Where ����id = ����id_In And �����id Is Not Null And У�Ա�־ = 1;
  If n_Count > 0 Then
    v_Err_Msg := '������δ���ýӿ�֧����������ɽ��㣡';
    Raise Err_Item;
  End If;
  If v_���㷽ʽ Is Null And Nvl(n_�շ�����, 0) = 0 Then
    v_Err_Msg := '����δָ���Ľ��㷽ʽ��������ɽ��㣡';
    Raise Err_Item;
  End If;
  If Nvl(n_�˷�, 0) = 1 Then
    Select Max(�Ƿ����Ʊ��) Into n_�Ƿ����Ʊ�� From ����Ԥ����¼ Where ����id = n_ԭ����id;
  Else
    n_�Ƿ����Ʊ�� := �Ƿ����Ʊ��_In;
    If �Ƿ����Ʊ��_In Is Null Then
      Select Nvl(Max(����), 0) Into n_���� From ���ս����¼ Where ��¼id = ����id_In And ���� = 2;
      n_�Ƿ����Ʊ�� := Zl_Fun_Isstarteinvoice(3, n_����);
    End If;
  End If;

  --2.����ɿ����ݺ��Ҳ����ݼ�У�Ա�־����Ϊ0
  Update ����Ԥ����¼ Set У�Ա�־ = 0, �Ƿ����Ʊ�� = n_�Ƿ����Ʊ�� Where ����id = ����id_In;

  If Nvl(n_�շ�����, 0) = 1 Then
    Update ����Ԥ����¼ Set ���㷽ʽ = Null Where NO = ���ݺ�_In And ��¼���� = 5 And ��¼״̬ = 3 And У�Ա�־ = 1;
  End If;

  --3.���·���״̬
  If Nvl(n_�շ�����, 0) = 0 Then
    Update סԺ���ü�¼ Set ����״̬ = 0 Where ����id = ����id_In;
  End If;

  --4.������Ա�ɿ�����,Not Exists����Ҫ������������������ϵ�ԭʼ���ݽ���ɹ��˵ģ�����Ҳ���˷ѽӿڣ������ܸ��½ɿ����
  For c_�ɿ� In (Select ���㷽ʽ, ����Ա����, Sum(Nvl(a.��Ԥ��, 0)) As ��Ԥ��
               From ����Ԥ����¼ A
               Where a.����id = ����id_In And Mod(a.��¼����, 10) <> 1 And ���㷽ʽ Is Not Null And Not Exists
                (Select 1
                      From ����Ԥ����¼ B
                      Where b.No = a.No And b.��¼���� = a.��¼���� And b.��¼״̬ = 3 And b.��������id = a.��������id And
                            Nvl(b.У�Ա�־, 0) <> 0)
               Group By ���㷽ʽ, ����Ա����) Loop
  
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + Nvl(c_�ɿ�.��Ԥ��, 0)
    Where �տ�Ա = c_�ɿ�.����Ա���� And ���� = 1 And ���㷽ʽ = c_�ɿ�.���㷽ʽ;
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ����
        (�տ�Ա, ���㷽ʽ, ����, ���)
      Values
        (c_�ɿ�.����Ա����, c_�ɿ�.���㷽ʽ, 1, Nvl(c_�ɿ�.��Ԥ��, 0));
    End If;
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҽ�ƿ�����_Modify;
/



Create Or Replace Procedure Zl_����Ԥ����¼_��Ԥ��
(
  ����id_In        ����Ԥ����¼.����id%Type,
  ����id_In        ����Ԥ����¼.����id%Type,
  ��Ԥ��_In        ����Ԥ����¼.��Ԥ��%Type := Null,
  Ԥ�����_In      ����Ԥ����¼.Ԥ�����%Type,
  ����Ա���_In    ����Ԥ����¼.����Ա���%Type,
  ����Ա����_In    ����Ԥ����¼.����Ա����%Type,
  �տ�ʱ��_In      ����Ԥ����¼.�տ�ʱ��%Type,
  ��Ԥ������ids_In Varchar2 := Null,
  ��������_In      Number := 3,
  ���������_In  Number := 0,
  У�Ա�־_In      ����Ԥ����¼.У�Ա�־%Type := Null,
  �����ƻỰ_In    ����Ԥ����¼.�Ự��%Type := 0,
  �Ƿ����Ʊ��_In  ����Ԥ����¼.�Ƿ����Ʊ��%Type := Null
) As
  --���������_In  0-Ԥ�������ʱ����ȥ������1-��ȥ�������
  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  v_��Ԥ������ids Varchar2(4000);
  n_����ֵ        ��Ա�ɿ����.���%Type;
  n_Ԥ�����      ����Ԥ����¼.��Ԥ��%Type;
  n_��Ԥ��        ����Ԥ����¼.��Ԥ��%Type;
  n_�Ự��        ����Ԥ����¼.�Ự��%Type; --��ʽ��SID+'_'+SERIAL#
  n_��id          ����ɿ����.Id%Type;
Begin
  v_��Ԥ������ids := Nvl(��Ԥ������ids_In, ����id_In);
  n_��id          := Zl_Get��id(����Ա����_In);

  If Nvl(�����ƻỰ_In, 0) = 0 Then
    Select Max(Sid || '_' || Serial#) Into n_�Ự�� From V$session Where Audsid = Userenv('sessionid');
  End If;

  --Ԥ�����
  If Nvl(��Ԥ��_In, 0) <> 0 Then
    If Nvl(����id_In, 0) = 0 Then
      v_Err_Msg := '����ȷ�����˵Ĳ���ID,�շѲ���ʹ��Ԥ�������,�������ʧ�ܣ�';
      Raise Err_Item;
    End If;
  
    --���������
    Select Nvl(Sum(Nvl(Ԥ�����, 0) - Decode(���������_In, 1, Nvl(�������, 0), 0)), 0)
    Into n_Ԥ�����
    From �������
    Where ����id In (Select Column_Value From Table(f_Num2List(v_��Ԥ������ids))) And Nvl(����, 0) = 1 And ���� = Ԥ�����_In;
  
    If Nvl(n_Ԥ�����, 0) < Nvl(��Ԥ��_In, 0) Then
      v_Err_Msg := '���˵ĵ�ǰԤ�����Ϊ ' || LTrim(To_Char(n_Ԥ�����, '9999999990.00')) || '��С�ڱ���֧����� ' ||
                   LTrim(To_Char(��Ԥ��_In, '9999999990.00')) || '��֧��ʧ�ܣ�';
      Raise Err_Item;
    End If;
  
    n_Ԥ����� := ��Ԥ��_In;
  
    --�Ƚ����ã��������Լ���
    --���������㷽ʽΪ���տ����Ԥ���
    For c_��Ԥ�� In (Select a.No, b.Ԥ����� As ���, Nvl(a.����id, 0) As ����id, a.����id, a.��¼״̬, a.Id, a.�տ�ʱ��, a.��������id
                  From ����Ԥ����¼ A, Ԥ��������� B
                  Where a.Id = b.Ԥ��id And b.����id In (Select Column_Value From Table(f_Num2List(v_��Ԥ������ids))) And
                        Nvl(b.Ԥ�����, 2) = Ԥ�����_In And a.���㷽ʽ Not In (Select ���� From ���㷽ʽ Where ���� = 5) And
                        Nvl(a.У�Ա�־, 0) = 0
                  Order By Decode(����id, Nvl(����id_In, 0), 0, 1), a.�տ�ʱ��) Loop
    
      If c_��Ԥ��.��� - n_Ԥ����� < 0 Then
        n_��Ԥ�� := c_��Ԥ��.���;
      Else
        n_��Ԥ�� := n_Ԥ�����;
      End If;
    
      If c_��Ԥ��.����id = 0 Then
        --��һ�γ�Ԥ��(���Ͻ���ID,���Ϊ0)
        Update ����Ԥ����¼
        Set ��Ԥ�� = 0, ����id = ����id_In, ������� = -1 * ����id_In, �������� = ��������_In, �Ự�� = n_�Ự��,
            �Ƿ����Ʊ�� = Decode(Nvl(�Ƿ����Ʊ��, 0), 1, 1, �Ƿ����Ʊ��_In)
        Where ID = c_��Ԥ��.Id;
      End If;
      --���ϴ�ʣ���
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, ��������, �Ự��, ��������id, ����ʱ��, ������Ա, У�Ա�־, �Ƿ����Ʊ��)
        Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��_In,
               ����Ա����_In, ����Ա���_In, n_��Ԥ��, ����id_In, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, -1 * ����id_In,
               ��������_In, n_�Ự��, c_��Ԥ��.��������id, �տ�ʱ��_In, ����Ա����_In, У�Ա�־_In, Decode(Nvl(�Ƿ����Ʊ��, 0), 1, 1, �Ƿ����Ʊ��_In)
        From ����Ԥ����¼
        Where NO = c_��Ԥ��.No And ��¼״̬ = c_��Ԥ��.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
    
      --��������(���㷽ʽΪNULL��)
      Update ����Ԥ����¼
      Set ��Ԥ�� = ��Ԥ�� - n_��Ԥ��
      Where ����id = ����id_In And ���㷽ʽ Is Null
      Returning Nvl(��Ԥ��, 0) Into n_����ֵ;
    
      --���²���Ԥ�����
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) - n_��Ԥ��
      Where ����id = c_��Ԥ��.����id And ���� = 1 And ���� = Ԥ�����_In
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ������� (����id, ����, Ԥ�����, ����) Values (c_��Ԥ��.����id, Ԥ�����_In, -1 * n_��Ԥ��, 1);
        n_����ֵ := -1 * n_��Ԥ��;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From �������
        Where ����id = c_��Ԥ��.����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
    
      --����Ԥ���������
      Update Ԥ���������
      Set Ԥ����� = Nvl(Ԥ�����, 0) - n_��Ԥ��
      Where ����id = c_��Ԥ��.����id And Ԥ��id = c_��Ԥ��.Id
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into Ԥ���������
          (Ԥ��id, ����id, Ԥ�����, Ԥ�����)
        Values
          (c_��Ԥ��.Id, c_��Ԥ��.����id, Ԥ�����_In, -1 * n_��Ԥ��);
        n_����ֵ := -1 * n_��Ԥ��;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From Ԥ��������� Where Ԥ��id = c_��Ԥ��.Id And Nvl(Ԥ�����, 0) = 0;
      End If;
    
      --����Ƿ��Ѿ�������
      If c_��Ԥ��.��� < n_Ԥ����� Then
        n_Ԥ����� := n_Ԥ����� - c_��Ԥ��.���;
      Else
        n_Ԥ����� := 0;
      End If;
      If n_Ԥ����� = 0 Then
        Exit;
      End If;
    
    End Loop;
    --������Ƿ��㹻
    If Abs(n_Ԥ�����) > 0 Then
      v_Err_Msg := '���˵ĵ�ǰԤ�����С�ڱ���֧����� ' || LTrim(To_Char(��Ԥ��_In, '9999999990.00')) || '�����ܼ���������';
      Raise Err_Item;
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ԥ����¼_��Ԥ��;
/

Create Or Replace Procedure Zl_����ԤԼ�Һ�_����
(
  No_In            In ���˹Һż�¼.No%Type,
  ����_In          In ���˹Һż�¼.����%Type,
  ����id_In        In ������ü�¼.����id%Type := Null,
  �����id_In      In ����Ԥ����¼.�����id%Type := Null,
  ����_In          In ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In    In ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In      In ����Ԥ����¼.����˵��%Type := Null,
  ����ʱ��_In      In ���˹Һż�¼.����ʱ��%Type := Null,
  ��Ԥ������ids_In In Varchar2 := Null,
  �Ƿ����Ʊ��_In  ����Ԥ����¼.�Ƿ����Ʊ��%Type := Null
  --�ù�������ֱ�����ԤԼ�ҺŽ��ա������Ҫ��ҽ��վʹ�á�
  -- ��Ԥ������ids_In:����ö��ַ���,��Ԥ��ʱ��Ч(��Ԥ�������ҵ���������һ��),��Ҫ��ʹ�ü�����Ԥ����
) As
  --�ű���Ϣ
  Cursor c_Regist Is
    Select b.����id, b.��Ŀid, b.ҽ��id, b.ҽ������, b.����
    From ������ü�¼ A, �ҺŰ��� B
    Where a.��¼���� = 4 And a.��¼״̬ = 1 And a.No = No_In And a.��� = 1 And a.���㵥λ = b.����;
  r_Regist c_Regist%RowType;

  Cursor c_Registnew Is
    Select b.����id, b.��Ŀid, b.ҽ��id, b.ҽ������, c.����
    From ���˹Һż�¼ A, �ٴ������¼ B, �ٴ������Դ C
    Where a.��¼���� = 1 And a.��¼״̬ = 1 And a.No = No_In And a.�����¼id = b.Id And b.��Դid = c.Id;
  r_Registnew c_Registnew%RowType;

  v_����no       ������ü�¼.No%Type;
  v_Temp         Varchar2(255);
  v_��Ա���     ������ü�¼.����Ա���%Type;
  v_��Ա����     ������ü�¼.����Ա����%Type;
  v_�Һ����ɶ��� Varchar2(2);
  v_�ŶӺ���     �ŶӽкŶ���.�ŶӺ���%Type;
  v_ԤԼ��ʽ     ���˹Һż�¼.ԤԼ��ʽ%Type;

  n_����id   ���˹Һż�¼.����id%Type;
  n_�����   ���˹Һż�¼.�����%Type;
  n_�ҺŽ�� ������ü�¼.ʵ�ս��%Type;
  n_ʣ���� �������.Ԥ�����%Type;
  n_����id   ������ü�¼.����id%Type;

  d_Date     Date;
  n_�����Ŷ� Number(18);
  n_�Ŷ�     Number(18);
  v_���㷽ʽ ����Ԥ����¼.���㷽ʽ%Type;
  v_�������� ҽ�ƿ����.����%Type;

  v_Error Varchar2(255);
  Err_Custom Exception;
  n_��id         ����ɿ����.Id%Type;
  v_�Ŷ����     �ŶӽкŶ���.�Ŷ����%Type;
  n_����ģʽ     ������Ϣ.����ģʽ%Type;
  v_���ʽ     ���˹Һż�¼.ҽ�Ƹ��ʽ%Type;
  n_�����¼id   �ٴ������¼.Id%Type;
  n_�Ƿ����Ʊ�� Number(2);
  n_����         ���ս����¼.����%Type;
Begin
  Begin
    Select a.����id, a.��ʶ��, Nvl(b.Ԥ�����, 0) - Nvl(b.�������, 0) As ���, Sum(a.ʵ�ս��) As n_�ҺŽ��, Substr(a.����, 1, 10)
    Into n_����id, n_�����, n_ʣ����, n_�ҺŽ��, v_ԤԼ��ʽ
    From ������ü�¼ A, ������� B
    Where a.����id = b.����id(+) And b.����(+) = 1 And b.����(+) = 1 And a.No = No_In And a.��¼���� = 4 And a.��¼״̬ = 0
    Group By a.����id, a.��ʶ��, Nvl(b.Ԥ�����, 0) - Nvl(b.�������, 0), a.����;
  Exception
    When Others Then
      v_Error := 'ԤԼ�Һ���Ϣ�����ڣ����ܸ�ԤԼ�Һ��ѱ����ա�';
      Raise Err_Custom;
  End;
  If ����ʱ��_In Is Null Then
    Select Sysdate Into d_Date From Dual;
  Else
    d_Date := ����ʱ��_In;
  End If;

  Begin
    Select �����¼id Into n_�����¼id From ���˹Һż�¼ Where NO = No_In;
  Exception
    When Others Then
      n_�����¼id := Null;
  End;

  n_�Ƿ����Ʊ�� := �Ƿ����Ʊ��_In;
  If �Ƿ����Ʊ��_In Is Null Then
    Select Nvl(Max(����), 0) Into n_���� From ���ս����¼ Where ��¼id = ����id_In And ���� = 1;
    n_�Ƿ����Ʊ�� := Zl_Fun_Isstarteinvoice(4, n_����);
  End If;

  --��ǰ������Ա
  v_Temp     := Zl_Identity;
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_��Ա��� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  n_��id     := Zl_Get��id(v_��Ա����);
  n_����ģʽ := 0;
  If Nvl(n_����id, 0) <> 0 Then
    Select Nvl(����ģʽ, 0) Into n_����ģʽ From ������Ϣ Where ����id = n_����id;
  End If;
  If n_����ģʽ = 0 Then
    If Nvl(����id_In, 0) = 0 Then
      Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
    Else
      n_����id := ����id_In;
    End If;
  End If;

  --������Ϣ�Ĳ���
  If n_����� Is Null Then
    Select To_Number(Nextno(3)) Into n_����� From Dual;
  End If;

  If n_����id Is Null Then
    Select To_Number(Nextno(1)) Into n_����id From Dual;
    Insert Into ������Ϣ
      (����id, �����, ����, �Ա�, ����, �ѱ�, ҽ�Ƹ��ʽ, �Ǽ�ʱ��)
      Select n_����id, n_�����, a.����, a.�Ա�, a.����, a.�ѱ�, b.����, d_Date
      From ������ü�¼ A, ҽ�Ƹ��ʽ B
      Where a.���ʽ = b.����(+) And a.No = No_In And a.��¼���� = 4 And a.��¼״̬ = 0 And a.��� = 1;
  End If;

  --���²�����Ϣ����������Ϣ
  Update ������Ϣ Set ����ʱ�� = d_Date, ����״̬ = 2, �������� = ����_In Where ����id = n_����id;

  --����������ü�¼����������Ϣ
  Update ������ü�¼
  Set ��¼״̬ = 1, ����id = Decode(n_����ģʽ, 1, Null, n_����id), ���ʽ�� = Decode(n_����ģʽ, 1, Null, ʵ�ս��), ��ҩ���� = ����_In, ִ���� = v_��Ա����,
      ִ��״̬ = 2, ִ��ʱ�� = d_Date, ����id = Decode(����id, Null, n_����id, ����id), ��ʶ�� = Decode(��ʶ��, Null, n_�����, ��ʶ��),
      �Ǽ�ʱ�� = d_Date, ����Ա��� = v_��Ա���, ����Ա���� = v_��Ա����, �ɿ���id = n_��id, ���ʷ��� = Decode(n_����ģʽ, 1, 1, 0)
  Where NO = No_In And ��¼���� = 4 And ��¼״̬ = 0;

  Update ���˹Һż�¼
  Set ��¼���� = 1, ������ = v_��Ա����, ����ʱ�� = d_Date, ���� = ����_In, ִ���� = v_��Ա����, ִ��ʱ�� = d_Date, ִ��״̬ = 2,
      ����id = Decode(����id, Null, n_����id, ����id), ����� = Decode(�����, Null, n_�����, �����)
  Where NO = No_In And ��¼״̬ = 1 And ��¼���� = 2;

  b_Message.Zlhis_Cis_008(n_����id, No_In);

  If Sql%NotFound Then
    --�������˹Һż�¼����������Ϣ
    Begin
      Select a.����
      Into v_���ʽ
      From ҽ�Ƹ��ʽ A, ������ü�¼ B
      Where b.No = No_In And b.��¼���� = 4 And b.��¼״̬ = 1 And b.��� = 1 And a.���� = b.���ʽ And Rownum < 2;
      Insert Into ���˹Һż�¼
        (ID, NO, ����id, ��¼����, ��¼״̬, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ����ʱ��, �Ǽ�ʱ��, ����Ա���, ����Ա����,
         ԤԼ, ԤԼ��ʽ, ����ʱ��, ������, ԤԼʱ��, ҽ�Ƹ��ʽ, �����¼id, �Һ���Ŀid, �ѱ�)
        Select ���˹Һż�¼_Id.Nextval, No_In, ����id, 1, 1, ��ʶ��, ����, �Ա�, ����, ���㵥λ, �Ӱ��־, ����_In, Null, ִ�в���id, v_��Ա����, 2, d_Date,
               ����ʱ��, �Ǽ�ʱ��, ����Ա���, ����Ա����, 1, Substr(����, 1, 10) As ԤԼ��ʽ, d_Date, v_��Ա����, ����ʱ��, v_���ʽ, n_�����¼id, �շ�ϸĿid,
               �ѱ�
        
        From ������ü�¼
        Where NO = No_In And ��¼���� = 4 And ��¼״̬ = 1 And ��� = 1;
    Exception
      When Others Then
        v_Error := '��ԤԼ�Һ��ѱ����ա�';
        Raise Err_Custom;
    End;
  End If;

  v_�Һ����ɶ��� := zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113));
  If v_�Һ����ɶ��� <> 0 Then
    For c_�Һ� In (Select ID, ִ�в���id, ����, ����_In As ����, �Ǽ�ʱ��, ִ���� As ִ����, ����id, �ű�, ����
                 From ���˹Һż�¼
                 Where NO = No_In And Rownum = 1) Loop
      Begin
        Select 1,
               Case
                 When �Ŷ�ʱ�� < Trunc(Sysdate) Then
                  1
                 Else
                  0
               End
        Into n_�Ŷ�, n_�����Ŷ�
        From �ŶӽкŶ���
        Where ҵ������ = 0 And ҵ��id = c_�Һ�.Id And Rownum <= 1;
      Exception
        When Others Then
          n_�Ŷ� := 0;
      End;
    
      If n_�Ŷ� = 0 Then
        --�����Ŷ�
        --��������
        --.����ִ�в��š� �ķ�ʽ���ɶ���
        v_�ŶӺ��� := Zlgetnextqueue(c_�Һ�.ִ�в���id, c_�Һ�.Id, c_�Һ�.�ű� || '|' || Nvl(c_�Һ�.����, 0));
        v_�Ŷ���� := Zlgetsequencenum(0, c_�Һ�.Id, 0);
        --  ��������_In , ҵ������_In, ҵ��id_In,����id_In,�ŶӺ���_In,�Ŷӱ��_In,��������_In,����ID_IN, ����_In, ҽ������_In, �Ŷ�ʱ��_In
        Zl_�ŶӽкŶ���_Insert(c_�Һ�.ִ�в���id, 0, c_�Һ�.Id, c_�Һ�.ִ�в���id, v_�ŶӺ���, Null, c_�Һ�.����, c_�Һ�.����id, c_�Һ�.����, c_�Һ�.ִ����,
                         Sysdate, v_ԤԼ��ʽ, Null, v_�Ŷ����);
      Elsif Nvl(n_�����Ŷ�, 0) = 1 Then
        --���¶��к�
        v_�ŶӺ��� := Zlgetnextqueue(c_�Һ�.ִ�в���id, c_�Һ�.Id, c_�Һ�.�ű� || '|' || Nvl(c_�Һ�.����, 0));
        v_�Ŷ���� := Zlgetsequencenum(0, c_�Һ�.Id, 1);
        --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In, ҽ������_In ,�ŶӺ���_In
        Zl_�ŶӽкŶ���_Update(c_�Һ�.ִ�в���id, 0, c_�Һ�.Id, c_�Һ�.ִ�в���id, c_�Һ�.����, c_�Һ�.����, c_�Һ�.ִ����, v_�ŶӺ���, v_�Ŷ����);
      
      Else
        --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In, ҽ������_In ,�ŶӺ���_In
        Zl_�ŶӽкŶ���_Update(c_�Һ�.ִ�в���id, 0, c_�Һ�.Id, c_�Һ�.ִ�в���id, c_�Һ�.����, c_�Һ�.����, c_�Һ�.ִ����);
      End If;
      --���պ�,�������
      Update �ŶӽкŶ��� Set �Ŷ�״̬ = 2 Where ҵ������ = 0 And ҵ��id = c_�Һ�.Id;
    End Loop;
  End If;

  --�Һŷ��ý���
  If Nvl(n_�ҺŽ��, 0) <> 0 Then
  
    If Nvl(n_ʣ����, 0) >= Nvl(n_�ҺŽ��, 0) And Nvl(�����id_In, 0) = 0 And n_����ģʽ = 0 Then
      --��Ԥ����ʽ����
      Zl_����Ԥ����¼_��Ԥ��(n_����id, n_����id, n_�ҺŽ��, 1, v_��Ա���, v_��Ա����, d_Date, ��Ԥ������ids_In, 4, 1, Null, 0, n_�Ƿ����Ʊ��);
    Elsif Nvl(�����id_In, 0) > 0 And n_����ģʽ = 0 Then
    
      Begin
        Select ���㷽ʽ, ���� Into v_���㷽ʽ, v_�������� From ҽ�ƿ���� Where ID = �����id_In;
      Exception
        When Others Then
          v_�������� := Null;
      End;
      If v_�������� Is Null Then
        v_Error := 'δ�ҵ������ӿ�,����ҽ�ƿ����������.';
        Raise Err_Custom;
      End If;
      If v_���㷽ʽ Is Null Then
        v_Error := v_�������� || 'δ���ö�Ӧ�Ľ��㷽ʽ,����ҽ�ƿ����������.';
        Raise Err_Custom;
      End If;
    
      --�������ӿ�֧��
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��, ����id,
         �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, Ԥ�����, �������, ��������, �Ƿ����Ʊ��)
        Select ����Ԥ����¼_Id.Nextval, NO, Null, 4, 1, ����id, ���˿���id, Null, v_���㷽ʽ, Null, 'ҽ��վ�ҺŽ���', Null, Null, Null, �Ǽ�ʱ��,
               ����Ա����, ����Ա���, n_�ҺŽ��, ����id, �ɿ���id, �����id_In, Null, ����_In, ������ˮ��_In, ����˵��_In, v_��������, Null, ����id, 4,
               n_�Ƿ����Ʊ��
        From ������ü�¼
        Where NO = No_In And ��¼���� = 4 And ��¼״̬ = 1 And ��� = 1;
    
      Update ��Ա�ɿ����
      Set ��� = Nvl(���, 0) + n_�ҺŽ��
      Where �տ�Ա = v_��Ա���� And ���� = 1 And ���㷽ʽ = v_���㷽ʽ;
      If Sql%RowCount = 0 Then
        Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (v_��Ա����, v_���㷽ʽ, 1, n_�ҺŽ��);
      End If;
    Else
      If n_����ģʽ = 1 Then
        --����
        For c_���� In (Select ʵ�ս��, ���˿���id, ��������id, ִ�в���id, ������Ŀid
                     From ������ü�¼
                     Where ��¼���� = 4 And ��¼״̬ = 1 And NO = No_In And Nvl(���ʷ���, 0) = 1) Loop
          --�������
          Update �������
          Set ������� = Nvl(�������, 0) + Nvl(c_����.ʵ�ս��, 0)
          Where ����id = Nvl(n_����id, 0) And ���� = 1 And ���� = 1;
        
          If Sql%RowCount = 0 Then
            Insert Into �������
              (����id, ����, ����, �������, Ԥ�����)
            Values
              (n_����id, 1, 1, Nvl(c_����.ʵ�ս��, 0), 0);
          End If;
        
          --����δ�����
          Update ����δ�����
          Set ��� = Nvl(���, 0) + Nvl(c_����.ʵ�ս��, 0)
          Where ����id = n_����id And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(c_����.���˿���id, 0) And
                Nvl(��������id, 0) = Nvl(c_����.��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(c_����.ִ�в���id, 0) And
                ������Ŀid + 0 = c_����.������Ŀid And ��Դ;�� + 0 = 1;
        
          If Sql%RowCount = 0 Then
            Insert Into ����δ�����
              (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
            Values
              (n_����id, Null, Null, c_����.���˿���id, c_����.��������id, c_����.ִ�в���id, c_����.������Ŀid, 1, Nvl(c_����.ʵ�ս��, 0));
          End If;
        End Loop;
      
      Else
      
        --���ɻ��۵��շ�(����������)
        v_Temp := zl_GetSysParameter('�Һ�ģʽ', 9000);
        If Nvl(v_Temp, '0') = '0' Then
          v_Error := '����ʣ����' || To_Char(Nvl(n_ʣ����, 0), '0.00') || ' ����ҺŽ��' || To_Char(Nvl(n_�ҺŽ��, 0), '0.00') ||
                     '���������ԤԼ���ա�';
          Raise Err_Custom;
        End If;
      
        Select Nextno(13) Into v_����no From Dual;
      
        Insert Into ������ü�¼
          (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �����־, ����id, ��ʶ��, ���ʽ, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
           ����, ��ҩ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ���ʷ���, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ��״̬, ժҪ,
           ����, �ɿ���id)
          Select ���˷��ü�¼_Id.Nextval, 1, v_����no, 0, a.���, a.��������, a.�۸񸸺�, a.�����־, a.����id, a.��ʶ��, a.���ʽ, a.����, a.�Ա�, a.����,
                 a.���˿���id, a.�ѱ�, a.�շ����, a.�շ�ϸĿid, b.���㵥λ, a.����, a.����, Null, Null, Null, a.������Ŀid, a.�վݷ�Ŀ, a.��׼����,
                 a.Ӧ�ս��, a.ʵ�ս��, 0, v_��Ա����, a.ִ�в���id, v_��Ա����, d_Date, d_Date, a.ִ�в���id, 0, '�Һ�:' || No_In, a.����, n_��id
          From ������ü�¼ A, �շ���ĿĿ¼ B
          Where a.�շ�ϸĿid = b.Id And a.No = No_In And a.��¼���� = 4 And a.��¼״̬ = 1;
      
        --�Һű����շ�
        Update ������ü�¼
        Set Ӧ�ս�� = 0, ʵ�ս�� = 0, ���ʽ�� = 0
        Where NO = No_In And ��¼���� = 4 And ��¼״̬ = 1;
      End If;
    End If;
  End If;

  --���˹ҺŻ���(ֻ����һ��,�ҵ�����ȡ�����Ѳ�����)
  If n_�����¼id Is Null Then
    Open c_Regist;
    Fetch c_Regist
      Into r_Regist;
    Update ���˹ҺŻ���
    Set �ѹ��� = Nvl(�ѹ���, 0) + 1, �����ѽ��� = Nvl(�����ѽ���, 0) + 1
    Where ���� = Trunc(d_Date) And Nvl(����id, 0) = Nvl(r_Regist.����id, 0) And Nvl(��Ŀid, 0) = Nvl(r_Regist.��Ŀid, 0) And
          Nvl(ҽ������, 'ҽ��') = Nvl(r_Regist.ҽ������, 'ҽ��') And Nvl(ҽ��id, 0) = Nvl(r_Regist.ҽ��id, 0) And
          (���� = r_Regist.���� Or ���� Is Null);
    If Sql%RowCount = 0 Then
      Insert Into ���˹ҺŻ���
        (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���)
      Values
        (Trunc(d_Date), r_Regist.����id, r_Regist.��Ŀid, r_Regist.ҽ������, r_Regist.ҽ��id, r_Regist.����, 1);
    End If;
    Close c_Regist;
  Else
    Open c_Registnew;
    Fetch c_Registnew
      Into r_Registnew;
    Update �ٴ������¼ Set �ѹ��� = Nvl(�ѹ���, 0) + 1, �����ѽ��� = Nvl(�����ѽ���, 0) + 1 Where ID = n_�����¼id;
    Update ���˹ҺŻ���
    Set �ѹ��� = Nvl(�ѹ���, 0) + 1, �����ѽ��� = Nvl(�����ѽ���, 0) + 1
    Where ���� = Trunc(d_Date) And Nvl(����id, 0) = Nvl(r_Registnew.����id, 0) And Nvl(��Ŀid, 0) = Nvl(r_Registnew.��Ŀid, 0) And
          Nvl(ҽ������, 'ҽ��') = Nvl(r_Registnew.ҽ������, 'ҽ��') And Nvl(ҽ��id, 0) = Nvl(r_Registnew.ҽ��id, 0) And
          (���� = r_Registnew.���� Or ���� Is Null);
    If Sql%RowCount = 0 Then
      Insert Into ���˹ҺŻ���
        (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���)
      Values
        (Trunc(d_Date), r_Registnew.����id, r_Registnew.��Ŀid, r_Registnew.ҽ������, r_Registnew.ҽ��id, r_Registnew.����, 1);
    End If;
    Close c_Registnew;
  End If;

  --���˵�����Ϣ
  If n_����id Is Not Null Then
    Update ������Ϣ
    Set ������ = Null, ������ = Null, �������� = Null
    Where ����id = n_����id And Exists (Select 1
           From ���˵�����¼
           Where ����id = n_����id And ��ҳid Is Not Null And
                 �Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ���˵�����¼ Where ����id = n_����id));
    If Sql%RowCount > 0 Then
      Update ���˵�����¼
      Set ����ʱ�� = d_Date
      Where ����id = n_����id And ��ҳid Is Not Null And Nvl(����ʱ��, d_Date) > d_Date;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ԤԼ�Һ�_����;
/

Create Or Replace Procedure Zl_ҽ�ƿ���¼_Delete
(
  ���ݺ�_In     סԺ���ü�¼.No%Type,
  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
  �˷�����_In   Integer := 0,
  �˷ѷ�ʽ_In   ����Ԥ����¼.���㷽ʽ%Type := Null,
  ����id_In     סԺ���ü�¼.����id%Type := 0
) As
  --�˷�����_In:0-�˿������������ѣ�1-�˿����������ѣ�2-�����˲�����
  Cursor c_Cardinfo Is
    Select a.Id As ����id, Nvl(a.���ʷ���, 0) As ����, a.����id, a.ʵ��Ʊ��, a.����id, Nvl(a.��ҳid, 0) As ��ҳid,
           Nvl(a.���˲���id, 0) As ���˲���id, Nvl(a.���˿���id, 0) As ���˿���id, Nvl(a.��������id, 0) As ��������id,
           Nvl(a.ִ�в���id, 0) As ִ�в���id, a.������Ŀid, a.ʵ�ս��, b.���㷽ʽ, b.��Ԥ��, b.�����id, b.����, b.���㿨���, b.�������, a.����,
           b.Id As Ԥ��id, a.ժҪ, a.����״̬
    From סԺ���ü�¼ A, ����Ԥ����¼ B
    Where a.��¼���� = 5 And a.��¼״̬ = 1 And a.No = ���ݺ�_In And a.����id = b.����id(+) And a.���ӱ�־ <> 8;
  r_Cardrow c_Cardinfo%RowType;

  Cursor c_Booksinfo Is
    Select a.Id As ����id, Nvl(a.���ʷ���, 0) As ����, a.����id, a.ʵ��Ʊ��, a.����id, Nvl(a.��ҳid, 0) As ��ҳid,
           Nvl(a.���˲���id, 0) As ���˲���id, Nvl(a.���˿���id, 0) As ���˿���id, Nvl(a.��������id, 0) As ��������id,
           Nvl(a.ִ�в���id, 0) As ִ�в���id, a.������Ŀid, a.ʵ�ս��, b.���㷽ʽ, b.��Ԥ��, b.�����id, b.����, b.���㿨���, b.�������, a.����,
           b.Id As Ԥ��id, a.ժҪ, a.����״̬
    From סԺ���ü�¼ A, ����Ԥ����¼ B
    Where a.��¼���� = 5 And a.��¼״̬ = 1 And a.No = ���ݺ�_In And a.����id = b.����id(+) And a.���ӱ�־ = 8;
  r_Booksrow c_Booksinfo%RowType;

  v_����id     סԺ���ü�¼.Id%Type;
  v_����id     סԺ���ü�¼.����id%Type;
  n_����ֵ     �������.�������%Type;
  n_�����id   Number(18);
  v_����״̬   ������ü�¼.��¼״̬%Type;
  n_����id     סԺ���ü�¼.Id%Type;
  n_�˷ѽ��   ����Ԥ����¼.��Ԥ��%Type;
  v_Date       Date;
  n_����       Number(1);
  n_����id     סԺ���ü�¼.����id%Type;
  n_��ҳid     סԺ���ü�¼.��ҳid%Type;
  n_���˿���id סԺ���ü�¼.���˿���id%Type;
  n_���˲���id סԺ���ü�¼.���˲���id%Type;
  n_��������id סԺ���ü�¼.��������id%Type;
  n_ִ�в���id סԺ���ü�¼.ִ�в���id%Type;
  n_������Ŀid סԺ���ü�¼.������Ŀid%Type;
  n_�����˷�   Number; --��¼�Ƿ��Ǵ˵��ݵĵڶ����˷�
  n_��Ԥ��id   ����Ԥ����¼.Id%Type;
  n_У�Ա�־   ����Ԥ����¼.У�Ա�־%Type;
  v_Err_Msg    Varchar2(200);
  Err_Item Exception;
  n_��id ����ɿ����.Id%Type;

Begin
  Begin
    Select 1 Into n_�����˷� From סԺ���ü�¼ Where ��¼���� = 5 And NO = ���ݺ�_In And ��¼״̬ = 3 And Rownum < 2;
  Exception
    When Others Then
      n_�����˷� := 0;
  End;

  If �˷�����_In <> 2 Then
    Open c_Cardinfo;
    Fetch c_Cardinfo
      Into r_Cardrow;
    n_��id := Zl_Get��id(����Ա����_In);
  
    --�����ж�Ҫ�˿��ļ�¼�Ƿ����
    If c_Cardinfo%RowCount = 0 Then
      Close c_Cardinfo;
      v_Err_Msg := '[ZLSOFT]û�з���Ҫ�˿��ļ�¼,�ü�¼�����Ѿ��˳���[ZLSOFT]';
      Raise Err_Item;
    Else
      Select Sysdate Into v_Date From Dual;
      Select ���˷��ü�¼_Id.Nextval Into v_����id From Dual;
    
      If r_Cardrow.���� = 0 Then
        If Nvl(����id_In, 0) = 0 Then
          Select ���˽��ʼ�¼_Id.Nextval Into v_����id From Dual;
        Else
          v_����id := ����id_In;
        End If;
        n_У�Ա�־ := 1;
      Else
        n_У�Ա�־ := 0;
      End If;
    
      --�˳����￨���ü�¼
      Insert Into סԺ���ü�¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ����id, ��ҳid, ���˲���id, ���˿���id, �����־, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ����,
         �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ����id,
         ���ʽ��, �ɿ���id, ����, ժҪ, ����״̬)
        Select v_����id, NO, ʵ��Ʊ��, ��¼����, 2, ���, ����id, ��ҳid, ���˲���id, ���˿���id, �����־, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
               -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ����Ա���_In, ����Ա����_In,
               ����ʱ��, v_Date, v_����id, Decode(v_����id, Null, Null, -���ʽ��), n_��id, ����, ժҪ, n_У�Ա�־
        From סԺ���ü�¼
        Where ID = r_Cardrow.����id;
    
      Update סԺ���ü�¼ Set ��¼״̬ = 3 Where ID = r_Cardrow.����id;
    
      --����˲����ѣ���Ҫͬʱ��������
      If Nvl(�˷�����_In, 0) = 1 Then
        Begin
          Select ID
          Into n_����id
          From סԺ���ü�¼
          Where ��¼���� = 5 And ��¼״̬ = 1 And NO = ���ݺ�_In And ���ӱ�־ = 8;
        Exception
          When Others Then
            n_����id := 0;
        End;
        If n_����id <> 0 Then
          Select ���˷��ü�¼_Id.Nextval Into v_����id From Dual;
        
          Insert Into סԺ���ü�¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ����id, ��ҳid, ���˲���id, ���˿���id, �����־, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ����,
             �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ����id,
             ���ʽ��, �ɿ���id, ����, ժҪ, ����״̬)
            Select v_����id, NO, ʵ��Ʊ��, ��¼����, 2, ���, ����id, ��ҳid, ���˲���id, ���˿���id, �����־, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ,
                   ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ����Ա���_In,
                   ����Ա����_In, ����ʱ��, v_Date, v_����id, Decode(v_����id, Null, Null, -���ʽ��), n_��id, ����, ժҪ, n_У�Ա�־
            
            From סԺ���ü�¼
            Where ID = n_����id;
        
          Update סԺ���ü�¼ Set ��¼״̬ = 3 Where ID = n_����id;
        End If;
      End If;
      --���������۵���������ۻ�δ�շѣ�ֱ��ɾ��
      Begin
        Select Nvl(��¼״̬, -1)
        Into v_����״̬
        From ������ü�¼
        Where ����id = r_Cardrow.����id And ��¼���� = 1 And NO = r_Cardrow.ժҪ;
      Exception
        When Others Then
          v_����״̬ := -1;
      End;
      If v_����״̬ = 0 Then
        Zl_���ﻮ�ۼ�¼_Delete(r_Cardrow.ժҪ);
      End If;
    
      If Nvl(�˷�����_In, 0) = 1 Then
        n_�˷ѽ�� := -1 * r_Cardrow.��Ԥ��;
      Else
        n_�˷ѽ�� := -1 * r_Cardrow.ʵ�ս��;
      End If;
    
      --Ԥ���������յĽ�����
      If r_Cardrow.���� = 0 Then
        Select ����Ԥ����¼_Id.Nextval Into n_��Ԥ��id From Dual;
        If �˷ѷ�ʽ_In Is Null Then
          Insert Into ����Ԥ����¼
            (ID, NO, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���,
             ����, ������ˮ��, ����˵��, ������λ, ��������, ��������id, У�Ա�־, �Ƿ����Ʊ��)
            Select n_��Ԥ��id, NO, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, �˷ѷ�ʽ_In, v_Date, ����Ա���_In, ����Ա����_In, n_�˷ѽ��, v_����id,
                   n_��id, Ԥ�����, Null, Null, Null, Null, Null, ������λ, 5, n_��Ԥ��id, n_У�Ա�־, �Ƿ����Ʊ��
            From ����Ԥ����¼
            Where ��¼���� = 5 And ��¼״̬ = Decode(n_�����˷�, 0, 1, 3) And ����id = r_Cardrow.����id;
        Else
          Insert Into ����Ԥ����¼
            (ID, NO, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���,
             ����, ������ˮ��, ����˵��, ������λ, ��������, ��������id, У�Ա�־, �Ƿ����Ʊ��)
            Select n_��Ԥ��id, NO, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, �˷ѷ�ʽ_In, v_Date, ����Ա���_In, ����Ա����_In, n_�˷ѽ��, v_����id,
                   n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 5, ��������id, n_У�Ա�־, �Ƿ����Ʊ��
            From ����Ԥ����¼
            Where ��¼���� = 5 And ��¼״̬ = Decode(n_�����˷�, 0, 1, 3) And ����id = r_Cardrow.����id;
        End If;
      
        Update ����Ԥ����¼ Set ��¼״̬ = 3 Where ��¼���� = 5 And ��¼״̬ = 1 And ����id = r_Cardrow.����id;
      End If;
    
      --����ҽ�ƿ�״̬Ϊ�˿��쳣
      If Nvl(r_Cardrow.����״̬, 0) = 0 Then
        n_�����id := To_Number(Nvl(r_Cardrow.����, '0'));
        Update ����ҽ�ƿ���Ϣ Set ״̬ = 3 Where �����id = n_�����id And ���� = r_Cardrow.ʵ��Ʊ��;
      End If;
    
      --����ͨ����Ϣ
      n_����       := r_Cardrow.����;
      n_����id     := r_Cardrow.����id;
      n_��ҳid     := r_Cardrow.��ҳid;
      n_���˿���id := r_Cardrow.���˿���id;
      n_���˲���id := r_Cardrow.���˲���id;
      n_��������id := r_Cardrow.��������id;
      n_ִ�в���id := r_Cardrow.ִ�в���id;
      n_������Ŀid := r_Cardrow.������Ŀid;
    
      Close c_Cardinfo;
    End If;
  Else
    Open c_Booksinfo;
    Fetch c_Booksinfo
      Into r_Booksrow;
    n_��id := Zl_Get��id(����Ա����_In);
  
    --�����ж�Ҫ�˿��ļ�¼�Ƿ����
    If c_Booksinfo%RowCount = 0 Then
      Close c_Booksinfo;
      v_Err_Msg := '[ZLSOFT]û�з���Ҫ�˷ѵļ�¼,�ü�¼�����Ѿ��˳���[ZLSOFT]';
      Raise Err_Item;
    Else
      Select Sysdate Into v_Date From Dual;
      Select ���˷��ü�¼_Id.Nextval Into v_����id From Dual;
    
      If r_Booksrow.���� = 0 Then
        If Nvl(����id_In, 0) = 0 Then
          Select ���˽��ʼ�¼_Id.Nextval Into v_����id From Dual;
        Else
          v_����id := ����id_In;
        End If;
        n_У�Ա�־ := 1;
      Else
        n_У�Ա�־ := 0;
      End If;
    
      n_�˷ѽ�� := -1 * r_Booksrow.ʵ�ս��;
    
      --�˳������ѷ��ü�¼
      Insert Into סԺ���ü�¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ����id, ��ҳid, ���˲���id, ���˿���id, �����־, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ����,
         �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ����id,
         ���ʽ��, �ɿ���id, ����, ժҪ, ����״̬)
        Select v_����id, NO, ʵ��Ʊ��, ��¼����, 2, ���, ����id, ��ҳid, ���˲���id, ���˿���id, �����־, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
               -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ����Ա���_In, ����Ա����_In,
               ����ʱ��, v_Date, v_����id, Decode(v_����id, Null, Null, -���ʽ��), n_��id, ����, ժҪ, n_У�Ա�־
        From סԺ���ü�¼
        Where ID = r_Booksrow.����id;
    
      Update סԺ���ü�¼ Set ��¼״̬ = 3 Where ID = r_Booksrow.����id;
    
      --Ԥ���������յĽ�����
      If r_Booksrow.���� = 0 Then
        Select ����Ԥ����¼_Id.Nextval Into n_��Ԥ��id From Dual;
        If �˷ѷ�ʽ_In Is Null Then
          Insert Into ����Ԥ����¼
            (ID, NO, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���,
             ����, ������ˮ��, ����˵��, ������λ, ��������, ��������id, У�Ա�־, �Ƿ����Ʊ��)
            Select n_��Ԥ��id, NO, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, �˷ѷ�ʽ_In, v_Date, ����Ա���_In, ����Ա����_In, n_�˷ѽ��, v_����id,
                   n_��id, Ԥ�����, Null, Null, Null, Null, Null, ������λ, 5, ID, n_У�Ա�־, �Ƿ����Ʊ��
            From ����Ԥ����¼
            Where ��¼���� = 5 And ��¼״̬ = Decode(n_�����˷�, 0, 1, 3) And ����id = r_Booksrow.����id;
        Else
          Insert Into ����Ԥ����¼
            (ID, NO, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���,
             ����, ������ˮ��, ����˵��, ������λ, ��������, ��������id, У�Ա�־, �Ƿ����Ʊ��)
            Select n_��Ԥ��id, NO, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, �˷ѷ�ʽ_In, v_Date, ����Ա���_In, ����Ա����_In, n_�˷ѽ��, v_����id,
                   n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 5, ��������id, n_У�Ա�־, �Ƿ����Ʊ��
            From ����Ԥ����¼
            Where ��¼���� = 5 And ��¼״̬ = Decode(n_�����˷�, 0, 1, 3) And ����id = r_Booksrow.����id;
        End If;
      
        Update ����Ԥ����¼ Set ��¼״̬ = 3 Where ��¼���� = 5 And ��¼״̬ = 1 And ����id = r_Booksrow.����id;
      End If;
    
      --����ͨ����Ϣ
      n_����       := r_Booksrow.����;
      n_����id     := r_Booksrow.����id;
      n_��ҳid     := r_Booksrow.��ҳid;
      n_���˿���id := r_Booksrow.���˿���id;
      n_���˲���id := r_Booksrow.���˲���id;
      n_��������id := r_Booksrow.��������id;
      n_ִ�в���id := r_Booksrow.ִ�в���id;
      n_������Ŀid := r_Booksrow.������Ŀid;
      Close c_Booksinfo;
    End If;
  End If;
  ----------------------------------------------------------------------------------------------------------------------------------------

  --��ػ��ܱ�Ĵ���
  If n_���� = 1 Then
    --����'�������'
    Update �������
    Set ������� = Nvl(�������, 0) + n_�˷ѽ��
    Where ���� = 1 And ����id = n_����id And Nvl(����, 2) = Decode(Nvl(n_��ҳid, 0), 0, 1, 2)
    Returning ������� Into n_����ֵ;
  
    If Sql%RowCount = 0 Then
      Insert Into �������
        (����id, ����, ����, Ԥ�����, �������)
      Values
        (n_����id, 1, Decode(Nvl(n_��ҳid, 0), 0, 1, 2), 0, n_�˷ѽ��);
      n_����ֵ := n_�˷ѽ��;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete ������� Where ���� = 1 And ����id = n_����id And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
    End If;
  
    --����'����δ�����'
    Update ����δ�����
    Set ��� = Nvl(���, 0) + n_�˷ѽ��
    Where ����id = n_����id And Nvl(��ҳid, 0) = n_��ҳid And Nvl(���˲���id, 0) = n_���˲���id And Nvl(���˿���id, 0) = n_���˿���id And
          Nvl(��������id, 0) = n_��������id And Nvl(ִ�в���id, 0) = n_ִ�в���id And ������Ŀid + 0 = n_������Ŀid And ��Դ;�� = 3;
  
    If Sql%RowCount = 0 Then
      Insert Into ����δ�����
        (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
      Values
        (n_����id, Decode(n_��ҳid, 0, Null, n_��ҳid), Decode(n_���˲���id, 0, Null, n_���˲���id),
         Decode(n_���˿���id, 0, Null, n_���˿���id), Decode(n_��������id, 0, Null, n_��������id), Decode(n_ִ�в���id, 0, Null, n_ִ�в���id),
         n_������Ŀid, 3, n_�˷ѽ��);
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20999, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҽ�ƿ���¼_Delete;
/

Create Or Replace Procedure Zl_�����շѽ���_Modify
(
  ��������_In      Number,
  ����id_In        ������ü�¼.����id%Type,
  ����id_In        ����Ԥ����¼.����id%Type,
  ���㷽ʽ_In      Varchar2,
  ��Ԥ��_In        ����Ԥ����¼.��Ԥ��%Type := Null,
  ��֧Ʊ��_In      ����Ԥ����¼.��Ԥ��%Type := Null,
  �����id_In      ����Ԥ����¼.�����id%Type := Null,
  ����_In          ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In      ����Ԥ����¼.����˵��%Type := Null,
  �ɿ�_In          ����Ԥ����¼.�ɿ�%Type := Null,
  �Ҳ�_In          ����Ԥ����¼.�Ҳ�%Type := Null,
  �����_In      ������ü�¼.ʵ�ս��%Type := Null,
  ��ɽ���_In      Number := 0,
  ȱʡ���㷽ʽ_In  ���㷽ʽ.����%Type := Null,
  ��Ԥ������ids_In Varchar2 := Null,
  ���½������_In  Number := 1,
  ��������id_In    ����Ԥ����¼.��������id%Type := Null,
  ɾ��ԭ����_In    Number := 0,
  У�Ա�־_In      ����Ԥ����¼.У�Ա�־%Type := 0,
  �����ƻỰ_In    ����Ԥ����¼.�Ự��%Type := 0,
  �Ƿ����Ʊ��_In  ����Ԥ����¼.�Ƿ����Ʊ��%Type := Null
) As
  ------------------------------------------------------------------------------------------------------------------------------ 
  --����:�շѽ���ʱ,�޸Ľ���������Ϣ 
  --��������_In: 
  --   0-��ͨ�շѷ�ʽ: 
  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������. 
  --     ����֧Ʊ��_In:����漰��֧Ʊ,���뱾�ε���֧Ʊ��,�������շ�ʱ,������ 
  --   1.����������: 
  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ" 
  --     ����֧Ʊ��_In:������ 
  --     �ۿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ���� 
  --   2-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���) 
  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
  --     ����֧Ʊ��_In:������
  --   3-���ѿ�����:
  --     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."  ���ѿ�ID:Ϊ��ʱ,���ݿ����Զ���λ 
  --     ����֧Ʊ��_In:������ 
  --   4-���������㣬���ֽ��㷽ʽ: 
  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ|���ݺ�|�Ƿ���ͨ����|����" 
  --     ����֧Ʊ��_In:������ 
  --     �ۿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ���� 

  -- ��Ԥ��_In: ���ڳ�Ԥ��ʱ,���� 
  -- �����_In:��������ʱ,���� 
  -- ��ɽ���_In:1-����շ�;0-δ����շ� 
  -- ��Ԥ������ids_In:����ö��ַ���,��Ԥ��ʱ��Ч(��Ԥ�������ҵ���������һ��),��Ҫ��ʹ�ü�����Ԥ���� 
  --���½������_In  �Ƿ������Ա��������Ҫ�Ǵ���ͳһ����Ա��¼��̨����������� 
  --��������id_In ��������_In Ϊ1,4ʱ���봫�� 
  --ɾ��ԭ����_in ��������_InΪ4ʱ��Ч��������㷽ʽʱ���ö�θù��� 
  --У�Ա�־_In  ��������_InΪ4ʱ��Ч 
    --�Ƿ����Ʊ��_In:null-��ʾ�����ڲ�ֱ���жϣ��ǿձ�ʾֱ���Դ����Ϊ׼
  ------------------------------------------------------------------------------------------------------------------------------ 
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_�������� Varchar2(500);
  v_��ǰ���� Varchar2(50);
  v_����     ����ҽ�ƿ���Ϣ.����%Type;
  n_���ѿ�id ���ѿ���Ϣ.Id%Type;
  v_����     ���ѿ����Ŀ¼.����%Type;
  n_�����id ����Ԥ����¼.���㿨���%Type;
  n_Ԥ��id   ����Ԥ����¼.Id%Type;

  v_���㷽ʽ ����Ԥ����¼.���㷽ʽ%Type;
  n_������ ����Ԥ����¼.��Ԥ��%Type;
  v_������� ����Ԥ����¼.�������%Type;
  v_����ժҪ ����Ԥ����¼.ժҪ%Type;
  v_������Ա ����Ԥ����¼.������Ա%Type;

  n_����ֵ   ��Ա�ɿ����.���%Type;
  n_��Ԥ��   ����Ԥ����¼.��Ԥ��%Type;
  v_��֧Ʊ   ����Ԥ����¼.���㷽ʽ%Type;
  v_����   ���㷽ʽ.����%Type;
  n_Count    Number;
  n_Havenull Number;
  l_Ԥ��id   t_Numlist := t_Numlist();
  n_�Ự��   ����Ԥ����¼.�Ự��%Type; --��ʽ��SID+'_'+SERIAL# 
n_�Ƿ����Ʊ�� Number(2);
n_����         ���ս����¼.����%Type;
  Cursor c_Feedata Is
    Select Max(m.����id) As ����id, Max(m.�Ǽ�ʱ��) As �Ǽ�ʱ��, Max(m.����Ա���) As ����Ա���, Max(m.����Ա����) As ����Ա����, Sum(���ʽ��) As ������,
           Max(m.�ɿ���id) As �ɿ���id
    From ������ü�¼ M
    Where m.����id = ����id_In;
  r_Feedata c_Feedata%RowType;

  Cursor c_Balancedata Is
    Select ��¼����, NO, ��¼״̬, ����id, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������
    From ����Ԥ����¼
    Where ����id = ����id_In And ���㷽ʽ Is Null;
  r_Balancedata c_Balancedata%RowType;

Begin
  IF Nvl(�����ƻỰ_In, 0) = 0 then
    Begin
      Select Sid || '_' || Serial# Into n_�Ự�� From V$session Where Audsid = Userenv('sessionid');
    Exception
      When Others Then
        n_�Ự�� := Null;
    End;
  End IF;
  v_������Ա := zl_UserName;

  Begin
    Select ���� Into v_���� From ���㷽ʽ Where Nvl(����, 0) = 9;
  Exception
    When Others Then
      v_���� := '����';
  End;

  --0.��ʽ���� 
  Select Count(1), Max(Decode(���㷽ʽ, Null, 1, 0))
  Into n_Count, n_Havenull
  From ����Ԥ����¼
  Where ����id = ����id_In;

  --1.���ӽ��㷽ʽΪ�յĽ������� 
  n_������ := 0;
  n_Count    := 0;
  Open c_Feedata;
  Begin
    Fetch c_Feedata
      Into r_Feedata;
    --�������������㷽ʽΪnull�ļ�¼ 
    Select Nvl(Sum(��Ԥ��), 0) Into n_������ From ����Ԥ����¼ Where ����id = ����id_In;
    If Nvl(n_Havenull, 0) = 0 Or Round(Nvl(r_Feedata.������, 0), 6) <> Round(Nvl(n_������, 0), 6) Then
      --��ɾ�����ڵĽ��㷽ʽΪnull�ļ�¼ 
      Delete From ����Ԥ����¼ Where ����id = ����id_In And ���㷽ʽ Is Null;
      Select Nvl(Sum(��Ԥ��), 0) Into n_������ From ����Ԥ����¼ Where ����id = ����id_In;
    
      n_������ := Round(Nvl(r_Feedata.������, 0) - n_������, 6);
    
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־, ��������, �Ự��)
      Values
        (����Ԥ����¼_Id.Nextval, 3, Null, 1, Decode(����id_In, 0, Null, ����id_In), Null, r_Feedata.�Ǽ�ʱ��, r_Feedata.����Ա���,
         r_Feedata.����Ա����, n_������, ����id_In, r_Feedata.�ɿ���id, Sysdate, v_������Ա, -1 * ����id_In, 1, 3, n_�Ự��);
    End If;
  Exception
    When Others Then
      n_Count := 1;
  End;
  Close c_Feedata;
  If n_Count = 1 Then
    v_Err_Msg := 'δ�ҵ�ָ�����շ���ϸ����,�������ʧ�ܣ�';
    Raise Err_Item;
  End If;

  If ��������_In = 0 And Nvl(��֧Ʊ��_In, 0) <> 0 Then
    Begin
      Select b.����
      Into v_��֧Ʊ
      From ���㷽ʽӦ�� A, ���㷽ʽ B
      Where a.Ӧ�ó��� = '�շ�' And b.���� = a.���㷽ʽ And Nvl(b.Ӧ����, 0) = 1 And Rownum <= 1;
    Exception
      When Others Then
        v_��֧Ʊ := '��';
    End;
    If v_��֧Ʊ = '��' Then
      v_Err_Msg := '�ڽ��㳡����,�����ڽ�������ΪӦ����Ľ��㷽ʽ,����[���㷽ʽ]�����ã�';
      Raise Err_Item;
    End If;
  End If;

  Open c_Balancedata;
  Fetch c_Balancedata
    Into r_Balancedata;

  If Nvl(�����_In, 0) <> 0 Then
    Update ����Ԥ����¼
    Set ��Ԥ�� = Nvl(��Ԥ��, 0) + Nvl(�����_In, 0)
    Where ����id = ����id_In And ���㷽ʽ = v_����;
    If Sql%NotFound Then
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־, �����id,
         ���㿨���, ����, ������ˮ��, ����˵��, �������, ��������, �Ự��)
      Values
        (����Ԥ����¼_Id.Nextval, 3, Null, 1, r_Balancedata.����id, Null, Null, v_����, r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���,
         r_Balancedata.����Ա����, �����_In, ����id_In, r_Balancedata.�ɿ���id, Sysdate, v_������Ա, r_Balancedata.�������, 2, Null, Null,
         ����_In, ������ˮ��_In, ����˵��_In, Null, 3, n_�Ự��);
    End If;
  
    Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - Nvl(�����_In, 0) Where ����id = ����id_In And ���㷽ʽ Is Null;
  End If;

  --Ԥ����� 
  If Nvl(��Ԥ��_In, 0) <> 0 Then
    If Nvl(����id_In, 0) = 0 Then
      v_Err_Msg := '����ȷ�����˵Ĳ���ID,�շѲ���ʹ��Ԥ�������,�������ʧ�ܣ�';
      Raise Err_Item;
    End If;
  
    Zl_����Ԥ����¼_��Ԥ��(����id_In, ����id_In, ��Ԥ��_In, 1, r_Balancedata.����Ա���, r_Balancedata.����Ա����, r_Balancedata.�տ�ʱ��,
                  ��Ԥ������ids_In, 3, 1);
  End If;

  If ��������_In = 0 Then
    If Nvl(��֧Ʊ��_In, 0) <> 0 Then
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־, �����id,
         ���㿨���, ����, ������ˮ��, ����˵��, �������, ��������, �Ự��)
      Values
        (����Ԥ����¼_Id.Nextval, 3, Null, 1, r_Balancedata.����id, Null, Null, v_��֧Ʊ, r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���,
         r_Balancedata.����Ա����, ��֧Ʊ��_In, ����id_In, r_Balancedata.�ɿ���id, Sysdate, v_������Ա, r_Balancedata.�������, 2, Null, Null,
         ����_In, ������ˮ��_In, ����˵��_In, Null, 3, n_�Ự��);
    
      Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - ��֧Ʊ��_In Where ����id = ����id_In And ���㷽ʽ Is Null;
    End If;
  
    --�����շѽ��� :��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." 
    v_�������� := ���㷽ʽ_In || '||';
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      v_������� := LTrim(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_����ժҪ := LTrim(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
    
      If Nvl(n_������, 0) <> 0 Then
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־,
           �����id, ���㿨���, ����, ������ˮ��, ����˵��, �������, ��������, �Ự��)
        Values
          (����Ԥ����¼_Id.Nextval, 3, Null, 1, r_Balancedata.����id, Null, v_����ժҪ, v_���㷽ʽ, r_Balancedata.�տ�ʱ��,
           r_Balancedata.����Ա���, r_Balancedata.����Ա����, n_������, ����id_In, r_Balancedata.�ɿ���id, Sysdate, v_������Ա,
           r_Balancedata.�������, 2, Null, Null, ����_In, ������ˮ��_In, ����˵��_In, v_�������, 3, n_�Ự��);
      
        Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - n_������ Where ����id = ����id_In And ���㷽ʽ Is Null;
      End If;
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  --1.���������㽻�� 
  If ��������_In = 1 Then
    v_��ǰ���� := ���㷽ʽ_In;
    v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
    v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
    n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
    v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
    v_������� := LTrim(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
    v_����ժҪ := LTrim(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
  
    If Nvl(n_������, 0) <> 0 Then
      Select Count(1) Into n_Count From ����Ԥ����¼ Where ID = ��������id_In And Rownum < 2;
      If n_Count = 0 And Nvl(��������id_In, 0) <> 0 Then
        n_Ԥ��id := ��������id_In;
      Else
        Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      End If;
    
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־, �����id,
         ���㿨���, ����, ��������id, ������ˮ��, ����˵��, �������, ��������, �Ự��)
      Values
        (n_Ԥ��id, 3, Null, 1, r_Balancedata.����id, Null, v_����ժҪ, v_���㷽ʽ, r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���,
         r_Balancedata.����Ա����, n_������, ����id_In, r_Balancedata.�ɿ���id, Sysdate, v_������Ա, r_Balancedata.�������, 2, �����id_In,
         Null, ����_In, ��������id_In, ������ˮ��_In, ����˵��_In, v_�������, 3, n_�Ự��);
    
      Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - n_������ Where ����id = ����id_In And ���㷽ʽ Is Null;
    
      --��������������Ϣ����
      Zl_Custom_Balance_Update(n_Ԥ��id);
    End If;
  End If;

  --2.ҽ������(���ô˹���,��ȡƽ����̯�ķ�ʽ��̯�������):�������ҽ���ᴦ��,����ȫ�� 
  If ��������_In = 2 Then
    --2.1����Ƿ��Ѿ�����ҽ����������,������ɾ�� 
    n_������ := 0;
    For v_ҽ�� In (Select ID, ���㷽ʽ, ��Ԥ��
                 From ����Ԥ����¼ A
                 Where a.����id = ����id_In And Exists
                  (Select 1 From ���㷽ʽ Where a.���㷽ʽ = ���� And ���� In (3, 4)) And Mod(��¼����, 10) <> 1) Loop
      n_������ := n_������ + Nvl(v_ҽ��.��Ԥ��, 0);
      l_Ԥ��id.Extend;
      l_Ԥ��id(l_Ԥ��id.Count) := v_ҽ��.Id;
    End Loop;
  
    Update ����Ԥ����¼ Set ��Ԥ�� = Nvl(��Ԥ��, 0) + Nvl(n_������, 0) Where ����id = ����id_In And ���㷽ʽ Is Null;
  
    Forall I In 1 .. l_Ԥ��id.Count
      Delete From ����Ԥ����¼ Where ID = l_Ԥ��id(I);
  
    If ���㷽ʽ_In Is Not Null Then
      v_�������� := ���㷽ʽ_In || '||';
    End If;
  
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      n_������ := To_Number(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
    
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־, ��������,
         �Ự��)
      Values
        (����Ԥ����¼_Id.Nextval, 3, Null, 1, r_Balancedata.����id, Null, '���ս���', v_���㷽ʽ, r_Balancedata.�տ�ʱ��,
         r_Balancedata.����Ա���, r_Balancedata.����Ա����, n_������, ����id_In, r_Balancedata.�ɿ���id, Sysdate, v_������Ա,
         r_Balancedata.�������, 1, 3, n_�Ự��);
    
      --��������(���㷽ʽΪNULL��) 
      Update ����Ԥ����¼
      Set ��Ԥ�� = ��Ԥ�� - n_������
      Where ����id = ����id_In And ���㷽ʽ Is Null
      Returning Nvl(��Ԥ��, 0) Into n_����ֵ;
    
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  --3-���ѿ��������� 
  If ��������_In = 3 Then
    v_�������� := ���㷽ʽ_In || '||';
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      --�����ID|����|���ѿ�ID|���ѽ�� 
      n_�����id := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      v_����     := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_���ѿ�id := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_������ := To_Number(v_��ǰ����);
      Begin
        Select ����, ���㷽ʽ Into v_����, v_���㷽ʽ From ���ѿ����Ŀ¼ Where ��� = �����id_In;
      Exception
        When Others Then
          v_���� := Null;
      End;
      If v_���� Is Null Then
        v_Err_Msg := 'δ�ҵ���Ӧ�Ľ��㿨�ӿ�,����ˢ������ʧ��!';
        Raise Err_Item;
      End If;
      If v_���㷽ʽ Is Null Then
        v_Err_Msg := v_���� || 'δ���ö�Ӧ�Ľ��㷽ʽ,����ˢ������ʧ��!';
        Raise Err_Item;
      End If;
    
      If Nvl(n_������, 0) <> 0 Then
        Update ����Ԥ����¼
        Set ��Ԥ�� = Nvl(��Ԥ��, 0) + n_������
        Where ����id = r_Balancedata. ����id And ���㷽ʽ = v_���㷽ʽ And ���㿨��� = n_�����id
        Returning ID Into n_Ԥ��id;
        If Sql%NotFound Then
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, ���㿨���,
             У�Ա�־, ��������, �Ự��)
          Values
            (n_Ԥ��id, 3, Null, 1, r_Balancedata. ����id, Null, Null, v_���㷽ʽ, r_Balancedata. �տ�ʱ��, r_Balancedata. ����Ա���,
             r_Balancedata. ����Ա����, n_������, r_Balancedata. ����id, r_Balancedata. �ɿ���id, Sysdate, v_������Ա,
             r_Balancedata. �������, n_�����id, 2, 3, n_�Ự��);
        End If;
      
        Zl_���˿������¼_֧��(n_�����id, v_����, n_���ѿ�id, n_������, n_Ԥ��id, r_Balancedata. ����Ա���, r_Balancedata. ����Ա����,
                      r_Balancedata. �տ�ʱ��);
      
        --��������(���㷽ʽΪNULL��) 
        Update ����Ԥ����¼
        Set ��Ԥ�� = ��Ԥ�� - n_������
        Where ����id = r_Balancedata. ����id And ���㷽ʽ Is Null And Nvl(У�Ա�־, 0) = 1
        Returning Nvl(��Ԥ��, 0) Into n_����ֵ;
      End If;
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  --4.���������㣬���ֽ��㷽ʽ 
  If ��������_In = 4 Then
    If Nvl(ɾ��ԭ����_In, 0) = 1 Then
      --1.1����Ƿ��Ѿ�������������������,������ɾ�� 
      n_������ := 0;
      For c_���� In (Select ID, ���㷽ʽ, ��Ԥ��
                   From ����Ԥ����¼ A
                   Where ����id = ����id_In And �����id = �����id_In And ��������id = ��������id_In) Loop
        n_������ := n_������ + Nvl(c_����.��Ԥ��, 0);
        l_Ԥ��id.Extend;
        l_Ԥ��id(l_Ԥ��id.Count) := c_����.Id;
      End Loop;
    
      Update ����Ԥ����¼
      Set ��Ԥ�� = Nvl(��Ԥ��, 0) + Nvl(n_������, 0)
      Where ����id = ����id_In And ���㷽ʽ Is Null;
    
      Forall I In 1 .. l_Ԥ��id.Count
        Delete From ����Ԥ����¼ Where ID = l_Ԥ��id(I);
    End If;
  
    n_Ԥ��id := 0;
    --��ʽ�����㷽ʽ|������|�������|����ժҪ|���ݺ�|�Ƿ���ͨ����|���� 
    For c_���� In (Select Max(Decode(���, 1, ֵ, Null)) As ���㷽ʽ, Zl_To_Number(Max(Decode(���, 2, ֵ, ''))) As ������,
                        Trim(Max(Decode(���, 3, ֵ, ''))) As �������, Trim(Max(Decode(���, 4, ֵ, ''))) As ����ժҪ,
                        Trim(Max(Decode(���, 5, ֵ, ''))) As ���ݺ�, Zl_To_Number(Max(Decode(���, 6, ֵ, ''))) As �Ƿ���ͨ����,
                        Trim(Max(Decode(���, 7, ֵ, ''))) As ����
                 From (Select Rownum As ���, Column_Value As ֵ From Table(f_Str2list(���㷽ʽ_In, '|')))
                 Having Nvl(Zl_To_Number(Max(Decode(���, 2, ֵ, ''))), 0) <> 0) Loop
    
      Update ����Ԥ����¼
      Set ��Ԥ�� = ��Ԥ�� + c_����.������
      Where ����id = ����id_In And ���㷽ʽ = c_����.���㷽ʽ And ��������id = ��������id_In
      Returning ID Into n_Ԥ��id;
      If Sql%NotFound Then
        Select Count(1) Into n_Count From ����Ԥ����¼ Where ID = ��������id_In And Rownum < 2;
        If n_Count = 0 Then
          n_Ԥ��id := ��������id_In;
        Else
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
        End If;
      
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־,
           �����id, ����, ��������id, ������ˮ��, ����˵��, �������, ��������, �Ự��)
        Values
          (n_Ԥ��id, 3, Null, 1, r_Balancedata.����id, Null, c_����.����ժҪ, c_����.���㷽ʽ, r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���,
           r_Balancedata.����Ա����, c_����.������, ����id_In, r_Balancedata.�ɿ���id, Sysdate, v_������Ա, r_Balancedata.�������, У�Ա�־_In,
           Decode(c_����.�Ƿ���ͨ����, 1, Null, �����id_In), Decode(c_����.�Ƿ���ͨ����, 1, Null, Nvl(c_����.����, ����_In)), ��������id_In,
           ������ˮ��_In, ����˵��_In, c_����.�������, 3, n_�Ự��);
      End If;
    
      Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - c_����.������ Where ����id = ����id_In And ���㷽ʽ Is Null;
    
      If c_����.���ݺ� Is Not Null Then
        Insert Into ҽ��������ϸ
          (����id, NO, ���㷽ʽ, ���, �����id, ��������id, ������ˮ��, ����˵��)
        Values
          (����id_In, c_����.���ݺ�, c_����.���㷽ʽ, c_����.������, Decode(c_����.�Ƿ���ͨ����, 1, Null, �����id_In), ��������id_In, ������ˮ��_In,
           ����˵��_In);
      End If;
    
      --��������������Ϣ����
      Zl_Custom_Balance_Update(n_Ԥ��id);
    End Loop;
  End If;

  If Nvl(��ɽ���_In, 0) = 0 Then
    Return;
  End If;

  ----------------------------------------------------------------------------------------- 
  --����շ�,��Ҫ������Ա�ɿ����,Ԥ����¼(���㷽ʽ=NULL) 

  --1.ɾ�����㷽ʽΪNULL��Ԥ����¼ 
  Delete ����Ԥ����¼ Where ����id = ����id_In And ���㷽ʽ Is Null And Nvl(��Ԥ��, 0) = 0;
  If Sql%NotFound Then
    Select Count(*) Into n_Count From ����Ԥ����¼ A Where ����id = ����id_In And ���㷽ʽ Is Null;
    If n_Count <> 0 Then
      v_Err_Msg := '������δ�ɿ������,������ɽ���!';
    Else
      v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�!';
    End If;
    Raise Err_Item;
  End If;

  --���������ü�¼�벡��Ԥ����¼�Ľ���Ƿ���� 
  n_������ := 0;
  n_��Ԥ��   := 0;
  Select Nvl(Sum(ʵ�ս��), 0) Into n_������ From ������ü�¼ Where ����id = ����id_In;
  Select Nvl(Sum(��Ԥ��), 0) Into n_��Ԥ�� From ����Ԥ����¼ Where ����id = ����id_In;
  If n_������ <> n_��Ԥ�� Then
    v_Err_Msg := '������Ϣ����ʵ�ս��(' || n_������ || ')�������(' || n_��Ԥ�� || ')��һ�£�������ɽ��㣡';
    Raise Err_Item;
  End If;

  --������Ϊ��ʱ������һ�����Ϊ0�Ĳ���Ԥ����¼ 
  Select Count(1) Into n_Count From ����Ԥ����¼ A Where ����id = ����id_In;
  If n_Count = 0 Then
    v_���㷽ʽ := ȱʡ���㷽ʽ_In;
    If v_���㷽ʽ Is Null Then
      Begin
        Select ���㷽ʽ Into v_���㷽ʽ From ���㷽ʽӦ�� Where Ӧ�ó��� = '�շ�' And Nvl(ȱʡ��־, 0) = 1;
      Exception
        When Others Then
          v_���㷽ʽ := Null;
      End;
      If v_���㷽ʽ Is Null Then
        Begin
          Select ���� Into v_���㷽ʽ From ���㷽ʽ Where Nvl(����, 0) = 1;
        Exception
          When Others Then
            v_���㷽ʽ := '�ֽ�';
        End;
      End If;
    End If;
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־, �����id,
       ���㿨���, ����, ������ˮ��, ����˵��, �������, ��������, �Ự��)
    Values
      (����Ԥ����¼_Id.Nextval, 3, Null, 1, r_Balancedata.����id, Null, Null, v_���㷽ʽ, r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���,
       r_Balancedata.����Ա����, 0, ����id_In, r_Balancedata.�ɿ���id, Sysdate, v_������Ա, r_Balancedata.�������, 2, Null, Null, Null,
       Null, ����˵��_In, Null, 3, n_�Ự��);
  End If;

  n_�Ƿ����Ʊ�� := �Ƿ����Ʊ��_In;
  If �Ƿ����Ʊ��_In Is Null Then
    Select Nvl(Max(����), 0) Into n_���� From ���ս����¼ Where ��¼id = ����id_In And ���� = 1;
    n_�Ƿ����Ʊ�� := Zl_Fun_Isstarteinvoice(1, n_����);
  End If;

  --2.����ɿ����ݺ��Ҳ����ݼ�У�Ա�־����Ϊ0���Ự�Ÿ���ΪNULL 
  Update ����Ԥ����¼ Set �ɿ� = �ɿ�_In, �Ҳ� = �Ҳ�_In, У�Ա�־ = 0, �Ự�� = Null,�Ƿ����Ʊ��=n_�Ƿ����Ʊ�� Where ����id = ����id_In;

  --3.���·���״̬ 
  Update ������ü�¼ Set ����״̬ = 0 Where ����id = ����id_In;

  --4.������Ա�ɿ����� 
  If Nvl(���½������_In, 1) = 1 Then
    For c_�ɿ� In (Select ���㷽ʽ, ����Ա����, Sum(Nvl(a.��Ԥ��, 0)) As ��Ԥ��
                 From ����Ԥ����¼ A
                 Where a.����id = ����id_In And Mod(a.��¼����, 10) <> 1
                 Group By ���㷽ʽ, ����Ա����) Loop
    
      Update ��Ա�ɿ����
      Set ��� = Nvl(���, 0) + Nvl(c_�ɿ�.��Ԥ��, 0)
      Where �տ�Ա = c_�ɿ�.����Ա���� And ���� = 1 And ���㷽ʽ = c_�ɿ�.���㷽ʽ;
      If Sql%RowCount = 0 Then
        Insert Into ��Ա�ɿ����
          (�տ�Ա, ���㷽ʽ, ����, ���)
        Values
          (c_�ɿ�.����Ա����, c_�ɿ�.���㷽ʽ, 1, Nvl(c_�ɿ�.��Ԥ��, 0));
      End If;
    End Loop;
  End If;

  --5.���ҵ�����ݴ��� 
  Zl_�����շѼ�¼_����շ�(����id_In);

  --��Ϣ���ɴ��� 
  --��������:1-�շѽ��㣬2-������� 
  --����ID:����id 
  b_Message.Zlhis_Charge_002(1, ����id_In);

  --�շѺ�������� 
  Begin
    Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
      Using 4, ����id_In;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����շѽ���_Modify;
/

CREATE OR REPLACE Procedure Zl_����Ʊ�ݿ�Ʊ��_Insert
(
  Id_In       In ����Ʊ�ݿ�Ʊ��.Id%Type,
  �ϼ�id_In   In ����Ʊ�ݿ�Ʊ��.�ϼ�id%Type,
  ����_In     In ����Ʊ�ݿ�Ʊ��.����%Type,
  ����_In     In ����Ʊ�ݿ�Ʊ��.����%Type,
  ����_In     In ����Ʊ�ݿ�Ʊ��.����%Type,
  Ժ��_In     In ����Ʊ�ݿ�Ʊ��.Ժ��%Type,
  �ͻ���_In   In ����Ʊ�ݿ�Ʊ��.�ͻ���%Type,
  ����id_In   In ����Ʊ�ݿ�Ʊ��.����id%Type,
  λ��_In     In ����Ʊ�ݿ�Ʊ��.λ��%Type,
  ĩ��_In     In ����Ʊ�ݿ�Ʊ��.ĩ��%Type := Null,
  ����ʱ��_In In ����Ʊ�ݿ�Ʊ��.����ʱ��%Type := Null
  
) Is
  d_����ʱ�� Date;
Begin

  d_����ʱ�� := ����ʱ��_In;
  If d_����ʱ�� Is Null Then
    d_����ʱ�� := Sysdate;
  End If;

  Insert Into ����Ʊ�ݿ�Ʊ��
    (ID, �ϼ�id, ����, ����, ����, Ժ��, �ͻ���, ����id, λ��, ĩ��, ����ʱ��, ����ʱ��)
  Values
    (Id_In, �ϼ�id_In, ����_In, ����_In, ����_In, Ժ��_In, �ͻ���_In, ����id_In, λ��_In, ĩ��_In, d_����ʱ��,
     To_Date('3000-01-01', 'yyyy-mm-dd hh24:mi:ss'));

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ʊ�ݿ�Ʊ��_Insert;
/
CREATE OR REPLACE Procedure Zl_����Ʊ�ݿ�Ʊ��_Update
(
  Id_In     In ����Ʊ�ݿ�Ʊ��.Id%Type,
  �ϼ�id_In In ����Ʊ�ݿ�Ʊ��.�ϼ�id%Type,
  ����_In   In ����Ʊ�ݿ�Ʊ��.����%Type,
  ����_In   In ����Ʊ�ݿ�Ʊ��.����%Type,
  ����_In   In ����Ʊ�ݿ�Ʊ��.����%Type,
  Ժ��_In   In ����Ʊ�ݿ�Ʊ��.Ժ��%Type,
  �ͻ���_In In ����Ʊ�ݿ�Ʊ��.�ͻ���%Type,
  ����id_In In ����Ʊ�ݿ�Ʊ��.����id%Type,
  λ��_In   In ����Ʊ�ݿ�Ʊ��.λ��%Type
) Is
  n_�ϼ�id ����Ʊ�ݿ�Ʊ��.�ϼ�id%Type;
Begin
  Select �ϼ�id Into n_�ϼ�id From ����Ʊ�ݿ�Ʊ�� Where ID = Id_In;

  Update ����Ʊ�ݿ�Ʊ��
  Set �ϼ�id = �ϼ�id_In, ���� = ����_In, ���� = ����_In, ���� = ����_In, Ժ�� = Ժ��_In, �ͻ��� = �ͻ���_In, ����id = ����id_In, λ�� = λ��_In
  Where ID = Id_In;

  Update ����Ʊ�ݿ�Ʊ��
  Set ���� = ����_In || Substr(����, Length(����_In) + 1)
  Where ID In (Select ID From ����Ʊ�ݿ�Ʊ�� Start With �ϼ�id = Id_In Connect By Prior ID = �ϼ�id);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ʊ�ݿ�Ʊ��_Update;
/
CREATE OR REPLACE Procedure Zl_����Ʊ�ݿ�Ʊ��_Start(Id_In In ����Ʊ�ݿ�Ʊ��.Id%Type) Is

Begin
  Update ����Ʊ�ݿ�Ʊ�� Set ����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd hh24:mi:ss') Where ID = Id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ʊ�ݿ�Ʊ��_Start;
/
CREATE OR REPLACE Procedure Zl_����Ʊ�ݿ�Ʊ��_Stop(Id_In In ����Ʊ�ݿ�Ʊ��.Id%Type) Is

Begin
  Update ����Ʊ�ݿ�Ʊ�� Set ����ʱ�� = Sysdate Where ID = Id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ʊ�ݿ�Ʊ��_Stop;
/
Create Or Replace Procedure Zl_����Ʊ�ݿ�Ʊ��_Delete(Id_In In ����Ʊ�ݿ�Ʊ��.Id%Type) Is
Begin
  Delete From ����Ʊ�ݿ�Ʊ�� Where ID = Id_In;
  Delete From Ʊ�ݿ�Ʊ����� Where ��Ʊ��id = Id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ʊ�ݿ�Ʊ��_Delete;
/
Create Or Replace Procedure Zl_Ʊ�ݿ�Ʊ�����_Update
(
  ����_In     In Number,
  Id_In       In Ʊ�ݿ�Ʊ�����.Id%Type := Null,
  ��Ʊ��id_In In ����Ʊ�ݿ�Ʊ��.Id%Type := Null,
  ��Աid_In   In Ʊ�ݿ�Ʊ�����.��Աid%Type := Null,
  �ͻ���_In   In Ʊ�ݿ�Ʊ�����.�ͻ���%Type := Null
) Is
  --˵��
  --����_In:0-����;1-�޸�;2-ɾ��;3-ɾ������
Begin
  --ɾ������
  If Nvl(����_In, 0) = 3 Then
    Delete From Ʊ�ݿ�Ʊ�����;
    Return;
  End If;
  --ɾ��
  If Nvl(����_In, 0) = 2 Then
    Delete From Ʊ�ݿ�Ʊ����� Where ID = Id_In;
    Return;
  End If;
  --�޸�
  If Nvl(����_In, 0) = 1 Then
    Update Ʊ�ݿ�Ʊ����� Set ��Աid = ��Աid_In, �ͻ��� = �ͻ���_In Where ��Ʊ��id = Id_In;
    Return;
  End If;

  --����
  Insert Into Ʊ�ݿ�Ʊ����� (ID, ��Ʊ��id, ��Աid, �ͻ���) Values (Id_In, ��Ʊ��id_In, ��Աid_In, �ͻ���_In);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Ʊ�ݿ�Ʊ�����_Update;
/


Create Or Replace Procedure Zl_����Ԥ����¼_Insert
(
  Id_In           ����Ԥ����¼.Id%Type,
  ���ݺ�_In       ����Ԥ����¼.No%Type,
  Ʊ�ݺ�_In       Ʊ��ʹ����ϸ.����%Type,
  ����id_In       ����Ԥ����¼.����id%Type,
  ��ҳid_In       ����Ԥ����¼.��ҳid%Type,
  ����id_In       ����Ԥ����¼.����id%Type,
  ���_In         ����Ԥ����¼.���%Type,
  ���㷽ʽ_In     ����Ԥ����¼.���㷽ʽ%Type,
  �������_In     ����Ԥ����¼.�������%Type,
  �ɿλ_In     ����Ԥ����¼.�ɿλ%Type,
  ��λ������_In   ����Ԥ����¼.��λ������%Type,
  ��λ�ʺ�_In     ����Ԥ����¼.��λ�ʺ�%Type,
  ժҪ_In         ����Ԥ����¼.ժҪ%Type,
  ����Ա���_In   ����Ԥ����¼.����Ա���%Type,
  ����Ա����_In   ����Ԥ����¼.����Ա����%Type,
  ����id_In       Ʊ��ʹ����ϸ.����id%Type,
  Ԥ�����_In     ����Ԥ����¼.Ԥ�����%Type := Null,
  �����id_In     ����Ԥ����¼.�����id%Type := Null,
  ���㿨���_In   ����Ԥ����¼.���㿨���%Type := Null,
  ����_In         ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In   ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In     ����Ԥ����¼.����˵��%Type := Null,
  ������λ_In     ����Ԥ����¼.������λ%Type := Null,
  �տ�ʱ��_In     ����Ԥ����¼.�տ�ʱ��%Type := Null,
  ��������_In     Integer := 0,
  ����id_In       ����Ԥ����¼.����id%Type := Null,
  ��������_In     ����Ԥ����¼.��������%Type := Null,
  �˿���_In     Number := 0,
  ǿ������_In     Number := 0,
  ���½������_In Number := 1,
  �Ƿ�ת��_In     Number := 0,
  У�Ա�־_In     ����Ԥ����¼.У�Ա�־%Type := Null,
  ����״̬_In     Number := 0,
  Ԥ������Ʊ��_In ����Ԥ����¼.Ԥ������Ʊ��%Type := Null,
  ����_In         ���ս����¼.����%Type := Null
) As
  ----------------------------------------------
  --��������_In:0-������Ԥ��;1-����Ϊδ��Ч��Ԥ����;3-����˿�
  --����ID_IN:>0ʱ,��ʾĳ�ν���ʱ,ͬ��������Ԥ����¼
  --�˿���_In;0-�����˿����Ƿ�����˲�����1-����˿���
  --���½������_In:0-�� zl_��Ա�ɿ����_Update �и��£�1-�ڱ������и���
  --ǿ������_In:0-��ǿ�ƣ�1-�����������ѿ����������ֵ�ǿ�����ֽ������
  --�Ƿ�ת��_In:0-ԭ���˻����֣�1-ת�˵�֧�ֵ���������
  --����״̬_In:0-�������㣬1-����Ϊ�쳣���ݣ�2-����쳣����

  v_Err_Msg Varchar2(200);
  Err_Item Exception;

  v_����     ���㷽ʽ.����%Type;
  v_��ӡid   Ʊ�ݴ�ӡ����.Id%Type;
  v_����     ������Ϣ.��������%Type;
  v_Date     Date;
  n_����ֵ   �������.Ԥ�����%Type;
  n_��id     ����ɿ����.Id%Type;
  n_������� �������.Ԥ�����%Type;
  n_����Ԥ�� �������.Ԥ�����%Type;
  n_�˿��� ����Ԥ����¼.���%Type;
  n_ʣ���   ����Ԥ����¼.���%Type;
  n_����id   ���˽��ʼ�¼.Id%Type;
  n_����     ���ս����¼.����%Type;

  n_Ԥ������Ʊ�� Number(2);
  Cursor c_��Ԥ�� Is
    Select a.Id, a.No, a.����id, a.Ԥ�����, a.�����id, a.����, a.������ˮ��, a.����˵��, 0 As ���, a.�տ�ʱ��, a.��� As Ԥ����, a.��������id
    From ����Ԥ����¼ A
    Where Rownum < 2;
  r_��Ԥ�� c_��Ԥ��%RowType;

  Type Ty_ʣ��� Is Ref Cursor;
  c_ʣ��� Ty_ʣ���; --��̬�α����
Begin

  n_Ԥ������Ʊ�� := Ԥ������Ʊ��_In;
  If n_Ԥ������Ʊ�� Is Null Then
    n_���� := ����_In;
    If ����_In Is Null Then
      Select Nvl(Max(����), 0) Into n_���� From ���ս����¼ Where ��¼id = Id_In And ���� = 3;
    End If;
    n_Ԥ������Ʊ�� := Zl_Fun_Isstarteinvoice(2, n_����, Ԥ�����_In);
  End If;

  v_Date := �տ�ʱ��_In;
  If v_Date Is Null Then
    Select Sysdate Into v_Date From Dual;
  End If;
  n_��id := Zl_Get��id(����Ա����_In);

  If Not (��������_In = 3 And ����״̬_In = 2) Then
  
    If ����״̬_In = 0 Or ����״̬_In = 1 Then
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ, �ɿ���id,
         Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, ��Ԥ��, ��������, У�Ա�־, ��������id, ����ʱ��, ������Ա, Ԥ������Ʊ��)
      Values
        (Id_In, ���ݺ�_In, Decode(����״̬_In, 0, Ʊ�ݺ�_In, Null), 1, Decode(����״̬_In, 1, 0, 1), ����id_In,
         Decode(��ҳid_In, 0, Null, ��ҳid_In), Decode(����id_In, 0, Null, ����id_In), ���_In, ���㷽ʽ_In, �������_In, v_Date, �ɿλ_In,
         ��λ������_In, ��λ�ʺ�_In, ����Ա���_In, ����Ա����_In, ժҪ_In, n_��id, Ԥ�����_In, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In,
         ������λ_In, ����id_In, Decode(����id_In, Null, Null, 0), ��������_In, У�Ա�־_In, Id_In, �տ�ʱ��_In, ����Ա����_In, n_Ԥ������Ʊ��);
    
      If Nvl(�����id_In, 0) <> 0 Then
        --�Զ�����̵���
        Zl_Custom_Balance_Update(Id_In);
      End If;
    End If;
  
    If ��������_In = 0 Then
      --����Ԥ���������
      Insert Into Ԥ��������� (Ԥ��id, ����id, Ԥ�����, Ԥ�����) Values (Id_In, ����id_In, Ԥ�����_In, ���_In);
    
    Elsif ��������_In = 1 Then
      --�ݲ�������ܱ�
      Return;
    Elsif ��������_In = 3 Then
      --����Ԥ���������
      Insert Into Ԥ��������� (Ԥ��id, ����id, Ԥ�����, Ԥ�����) Values (Id_In, ����id_In, Ԥ�����_In, ���_In);
    
      --����һ��ԭԤ��ID�ĳ�����¼��ͬʱҲ����һ������˿�ĳ�����¼
      --���տ���ܽ��г���
      Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
      If Nvl(�����id_In, 0) = 0 And Nvl(���㿨���_In, 0) = 0 Then
        --���֣�������ͨ���㷽ʽ���֡�ǿ�����֡���������������
        Open c_ʣ��� For
          Select Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, 0, a.Id), 0)) As ID, a.No, a.����id, a.Ԥ�����, a.�����id, a.����,
                 a.������ˮ��, a.����˵��, Min(Decode(Sign(a.���), -1, 0, 1)) As ���, Min(Decode(a.��¼����, 1, a.�տ�ʱ��, Null)) As �տ�ʱ��,
                 Nvl(Sum(a.���), 0) - Nvl(Sum(a.��Ԥ��), 0) As Ԥ����, Max(a.��������id) As ��������id
          From ����Ԥ����¼ A, ҽ�ƿ���� B, ���ѿ����Ŀ¼ C
          Where a.����id = ����id_In And a.��¼���� In (1, 11) And a.Ԥ����� = Nvl(Ԥ�����_In, 2) And a.�����id = b.Id(+) And
                Decode(ǿ������_In, 1, 1, Nvl(b.�Ƿ�����, 1)) = 1 And a.�����id = c.���(+) And
                Decode(ǿ������_In, 1, 1, Nvl(c.�Ƿ�����, 1)) = 1 And a.���㷽ʽ Not In (Select ���� From ���㷽ʽ Where ���� = 5)
          Group By a.No, a.����id, a.Ԥ�����, a.�����id, a.����, a.������ˮ��, a.����˵��
          Having Nvl(Sum(a.���), 0) - Nvl(Sum(a.��Ԥ��), 0) <> 0
          Order By ���, �տ�ʱ��;
      Elsif Nvl(�Ƿ�ת��_In, 0) = 1 Then
        --ת�ˣ��������������ֻ���ǿ�����֣�����Ŀ��ſ��ܲ���ԭ����,�����ͬ�ֿ�����Ԥ���ɿ��̯
        --Ŀǰֻ֧��ͬһ�ֿ�ת��
        Open c_ʣ��� For
          Select Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, 0, a.Id), 0)) As ID, a.No, a.����id, a.Ԥ�����, a.�����id, a.����,
                 a.������ˮ��, a.����˵��, Min(Decode(Sign(a.���), -1, 0, 1)) As ���, Min(Decode(a.��¼����, 1, a.�տ�ʱ��, Null)) As �տ�ʱ��,
                 Nvl(Sum(a.���), 0) - Nvl(Sum(a.��Ԥ��), 0) As Ԥ����, Max(a.��������id) As ��������id
          From ����Ԥ����¼ A, ҽ�ƿ���� B
          Where a.����id = ����id_In And a.��¼���� In (1, 11) And a.Ԥ����� = Nvl(Ԥ�����_In, 2) And a.�����id = b.Id(+) And
                Nvl(�����id, 0) = Nvl(�����id_In, 0) And Nvl(������ˮ��, '-') = Nvl(������ˮ��_In, '-') And
                a.���㷽ʽ Not In (Select ���� From ���㷽ʽ Where ���� = 5)
          Group By a.No, a.����id, a.Ԥ�����, a.�����id, a.����, a.������ˮ��, a.����˵��
          Having Nvl(Sum(a.���), 0) - Nvl(Sum(a.��Ԥ��), 0) <> 0
          Order By ���, �տ�ʱ��;
      Else
        --�����������������ѿ������ݿ����ID�����㿨��š����š�������ˮ��ȱʡԭԤ����¼���������ȷ��Ψһ����з�̯
        Open c_ʣ��� For
          Select Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, 0, a.Id), 0)) As ID, a.No, a.����id, a.Ԥ�����, a.�����id, a.����,
                 a.������ˮ��, a.����˵��, Min(Decode(Sign(a.���), -1, 0, 1)) As ���, Min(Decode(a.��¼����, 1, a.�տ�ʱ��, Null)) As �տ�ʱ��,
                 Nvl(Sum(a.���), 0) - Nvl(Sum(a.��Ԥ��), 0) As Ԥ����, Max(a.��������id) As ��������id
          From ����Ԥ����¼ A
          Where a.����id = ����id_In And a.��¼���� In (1, 11) And a.Ԥ����� = Nvl(Ԥ�����_In, 2) And
                Nvl(a.�����id, 0) = Nvl(�����id_In, 0) And Nvl(a.���㿨���, 0) = Nvl(���㿨���_In, 0) And
                Nvl(a.����, '-') = Nvl(����_In, '-') And Nvl(������ˮ��, '-') = Nvl(������ˮ��_In, '-') And
                a.���㷽ʽ Not In (Select ���� From ���㷽ʽ Where ���� = 5)
          Group By a.No, a.����id, a.Ԥ�����, a.�����id, a.����, a.������ˮ��, a.����˵��
          Having Nvl(Sum(a.���), 0) - Nvl(Sum(a.��Ԥ��), 0) <> 0
          Order By ���, �տ�ʱ��;
      End If;
    
      n_ʣ���   := -1 * ���_In;
      n_�˿��� := 0;
      Loop
        Fetch c_ʣ���
          Into r_��Ԥ��;
        Exit When c_ʣ���%NotFound;
        If r_��Ԥ��.No <> ���ݺ�_In Then
          If n_ʣ��� > r_��Ԥ��.Ԥ���� Then
            n_�˿��� := r_��Ԥ��.Ԥ����;
            n_ʣ���   := n_ʣ��� - n_�˿���;
          Else
            n_�˿��� := n_ʣ���;
            n_ʣ���   := 0;
          End If;
        
          If Nvl(n_�˿���, 0) <> 0 Then
            Update ����Ԥ����¼ Set ����id = n_����id Where NO = r_��Ԥ��.No And ��¼���� = 1 And ����id Is Null;
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, ���, ���㷽ʽ, �������, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, �տ�ʱ��, ����Ա����,
               ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, ��Ԥ��, ��������, ��������id, ����ʱ��, ������Ա, У�Ա�־, Ԥ������Ʊ��)
              Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, 1, ����id, ��ҳid, Ԥ�����, ����id, Null, ���㷽ʽ, �������, �ɿλ, ��λ������, ��λ�ʺ�,
                     ����Ա���_In, v_Date, ����Ա����_In, ժҪ, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, n_����id, n_�˿���, Null,
                     Nvl(r_��Ԥ��.��������id, r_��Ԥ��.Id), �տ�ʱ��_In, ����Ա����_In, У�Ա�־_In, n_Ԥ������Ʊ��
              From ����Ԥ����¼
              Where NO = r_��Ԥ��.No And ��¼���� In (1, 11) And Rownum < 2;
          
            --����Ԥ���������
            Update Ԥ���������
            Set Ԥ����� = Nvl(Ԥ�����, 0) - n_�˿���
            Where ����id = r_��Ԥ��.����id And Ԥ��id = r_��Ԥ��.Id
            Returning Ԥ����� Into n_����ֵ;
            If Sql%RowCount = 0 Then
              Insert Into Ԥ���������
                (Ԥ��id, ����id, Ԥ�����, Ԥ�����)
              Values
                (r_��Ԥ��.Id, r_��Ԥ��.����id, 1, -1 * n_�˿���);
              n_����ֵ := -1 * n_�˿���;
            End If;
            If Nvl(n_����ֵ, 0) = 0 Then
              Delete From Ԥ��������� Where Ԥ��id = r_��Ԥ��.Id And Nvl(Ԥ�����, 0) = 0;
            End If;
          
          End If;
        
          If n_ʣ��� = 0 Then
            Exit;
          End If;
        End If;
      End Loop;
    
      If n_ʣ��� <> 0 And Nvl(�˿���_In, 0) = 1 Then
        v_Err_Msg := '�˿�����ڲ���ʣ��Ԥ����';
        Raise Err_Item;
      End If;
    
      n_�˿��� := -1 * (-1 * ���_In - n_ʣ���);
      If n_�˿��� <> 0 Then
        Update ����Ԥ����¼ Set ����id = n_����id Where ID = Id_In;
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ, �ɿ���id,
           Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, ��Ԥ��, ��������, ��������id, ����ʱ��, ������Ա, У�Ա�־, Ԥ������Ʊ��)
        Values
          (����Ԥ����¼_Id.Nextval, ���ݺ�_In, Decode(����״̬_In, 0, Ʊ�ݺ�_In, Null), 11, 1, ����id_In,
           Decode(��ҳid_In, 0, Null, ��ҳid_In), Decode(����id_In, 0, Null, ����id_In), Null, ���㷽ʽ_In, �������_In, v_Date, �ɿλ_In,
           ��λ������_In, ��λ�ʺ�_In, ����Ա���_In, ����Ա����_In, ժҪ_In, n_��id, Ԥ�����_In, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In,
           ������λ_In, n_����id, n_�˿���, Null, Id_In, �տ�ʱ��_In, ����Ա����_In, У�Ա�־_In, n_Ԥ������Ʊ��);
        --����Ԥ���������
        Update Ԥ���������
        Set Ԥ����� = Nvl(Ԥ�����, 0) - n_�˿���
        Where ����id = ����id_In And Ԥ��id = Id_In
        Returning Ԥ����� Into n_����ֵ;
        If Sql%RowCount = 0 Then
          Insert Into Ԥ��������� (Ԥ��id, ����id, Ԥ�����, Ԥ�����) Values (Id_In, ����id_In, 1, -1 * n_�˿���);
          n_����ֵ := -1 * n_�˿���;
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From Ԥ��������� Where Ԥ��id = Id_In And Nvl(Ԥ�����, 0) = 0;
        End If;
      End If;
    
      If ���_In < 0 And Nvl(ǿ������_In, 0) = 0 Then
        Begin
          Select Nvl(Ԥ�����, 0) - Nvl(�������, 0)
          Into n_�������
          From �������
          Where ���� = 1 And ����id = ����id_In And Nvl(����, 0) = Nvl(Ԥ�����_In, 0);
        Exception
          When Others Then
            Null;
        End;
        --����˿�Ҫ��������Ԥ���Ƿ�֧������
      
        For c_����Ԥ�� In (Select a.Ԥ��id, a.Ԥ�����, a.�����id, a.���㿨��� As ���ѽӿ�id, Nvl(b.����, c.���) As ����, Nvl(b.����, c.����) As ����,
                              Decode(b.����, Null, c.�Ƿ�ȫ��, b.�Ƿ�ȫ��) As �Ƿ�ȫ��, Decode(b.����, Null, c.�Ƿ�����, b.�Ƿ�����) As �Ƿ�����,
                              a.����, a.������ˮ��, a.����˵��, a.Ԥ�����
                       From (Select a.Ԥ�����, Nvl(a.�����id, 0) As �����id, Nvl(a.���㿨���, 0) As ���㿨���, a.����, a.������ˮ��, a.����˵��,
                                     Max(Decode(Sign(���), -1, Decode(a.��¼״̬, 1, 0, 2, 0, ID), ID)) As Ԥ��id,
                                     Nvl(Sum(���), 0) - Nvl(Sum(Nvl(��Ԥ��, 0)), 0) As Ԥ�����
                              From ����Ԥ����¼ A
                              Where a.����id = ����id_In And (Nvl(a.���㿨���, 0) <> 0 Or Nvl(�����id, 0) <> 0)
                              Group By a.Ԥ�����, Nvl(a.�����id, 0), Nvl(a.���㿨���, 0), a.����, a.������ˮ��, a.����˵��
                              Having Nvl(Sum(���), 0) - Nvl(Sum(Nvl(��Ԥ��, 0)), 0) <> 0) A, ҽ�ƿ���� B, ���ѿ����Ŀ¼ C
                       Where a.Ԥ����� = Nvl(Ԥ�����_In, 0) And a.�����id = b.Id(+) And a.���㿨��� = c.���(+) And
                             Nvl(a.Ԥ�����, 0) <> 0
                       Order By ����, a.����, a.������ˮ��, a.����˵��) Loop
        
          If Instr(',7,8,', ',' || v_���� || ',') = 0 And Nvl(c_����Ԥ��.�Ƿ�����, 0) = 0 And Nvl(c_����Ԥ��.Ԥ�����, 0) > 0 Then
            n_����Ԥ�� := Nvl(n_����Ԥ��, 0) + Nvl(c_����Ԥ��.Ԥ�����, 0);
          Elsif Instr(',7,8,', ',' || v_���� || ',') > 0 Then
            If Nvl(c_����Ԥ��.����, '0') = Nvl(����_In, '0') And Nvl(c_����Ԥ��.������ˮ��, '0') = Nvl(������ˮ��_In, '0') And
               Nvl(c_����Ԥ��.����˵��, '0') = Nvl(����˵��_In, '0') Then
              n_����Ԥ�� := Nvl(n_����Ԥ��, 0) + Nvl(c_����Ԥ��.Ԥ�����, 0);
            End If;
          End If;
        End Loop;
      
        If Instr(',7,8,', ',' || v_���� || ',') > 0 And Nvl(n_����Ԥ��, 0) < 0 And ��������_In = 3 Then
          v_Err_Msg := '�˿�����ڲ�������Ԥ����';
          Raise Err_Item;
        Elsif Nvl(n_�������, 0) < 0 And �˿���_In = 1 Then
          v_Err_Msg := '�˿�����ڲ���ʣ��Ԥ����';
          Raise Err_Item;
        Elsif Instr(',7,8,', ',' || v_���� || ',') = 0 And Nvl(n_�������, 0) - Nvl(n_����Ԥ��, 0) < 0 And ��������_In = 3 And
              �˿���_In = 1 Then
          v_Err_Msg := '�˿�����ڲ���ʣ��Ԥ����';
          Raise Err_Item;
        End If;
      End If;
    End If;
    --�������(Ԥ���������)
  
    Select Max(����) Into v_���� From ���㷽ʽ Where ���� = ���㷽ʽ_In;
  
    If Nvl(v_����, 1) <> 5 Then
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) + ���_In
      Where ���� = 1 And ����id = ����id_In And Nvl(����, 0) = Nvl(Ԥ�����_In, 0)
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, ����, Ԥ�����, �������)
        Values
          (����id_In, 1, Nvl(Ԥ�����_In, 0), ���_In, 0);
        n_����ֵ := ���_In;
      End If;
      If Nvl(���_In, 0) = 0 Then
        Delete From ������� Where ����id = ����id_In And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
      End If;
    End If;
  End If;
  --������ɲŴ���Ʊ�ݣ���Ϣ��
  If ����״̬_In = 1 Then
    Return;
    --�����쳣����
  Elsif ����״̬_In = 2 Then
    If ��������_In = 3 Then
      Update ����Ԥ����¼ Set ��¼״̬ = 1, ʵ��Ʊ�� = Ʊ�ݺ�_In Where ID = Id_In Return ����id Into n_����id;
      Update ����Ԥ����¼
      Set ʵ��Ʊ�� = Ʊ�ݺ�_In
      Where NO = (Select NO From ����Ԥ����¼ Where ID = Id_In) And ��¼���� = 11;
      Update ����Ԥ����¼
      Set �տ�ʱ�� = v_Date, ����Ա��� = ����Ա���_In, ����Ա���� = ����Ա����_In, �ɿ���id = n_��id, ����ʱ�� = v_Date, ������Ա = ����Ա����_In,
          Ԥ������Ʊ�� = n_Ԥ������Ʊ��
      Where ����id = n_����id And Nvl(У�Ա�־, 0) = 1;
      Update ����Ԥ����¼ Set У�Ա�־ = Null Where ����id = n_����id;
      --�Զ�����̵���
      Zl_Custom_Balance_Update(Id_In);
    Else
      --���²��������
      Update ����Ԥ����¼
      Set ��¼״̬ = 1, У�Ա�־ = Null, ʵ��Ʊ�� = Ʊ�ݺ�_In, �տ�ʱ�� = v_Date, ����Ա��� = ����Ա���_In, ����Ա���� = ����Ա����_In, �ɿ���id = n_��id,
          ����ʱ�� = v_Date, ������Ա = ����Ա����_In, ����id = Decode(����id_In, 0, Null, ����id_In), ��� = ���_In, ���㷽ʽ = ���㷽ʽ_In,
          ������� = �������_In, �ɿλ = �ɿλ_In, ��λ������ = ��λ������_In, ��λ�ʺ� = ��λ�ʺ�_In, ժҪ = ժҪ_In, �����id = �����id_In,
          ���㿨��� = ���㿨���_In, ���� = ����_In, Ԥ������Ʊ�� = n_Ԥ������Ʊ��
      Where ID = Id_In;
      --�Զ�����̵���
      Zl_Custom_Balance_Update(Id_In);
    
    End If;
  End If;
  --����Ʊ��
  If Ʊ�ݺ�_In Is Not Null Then
    Select Ʊ�ݴ�ӡ����_Id.Nextval Into v_��ӡid From Dual;
  
    --����Ʊ��
    Insert Into Ʊ�ݴ�ӡ���� (ID, ��������, NO) Values (v_��ӡid, 2, ���ݺ�_In);
  
    Insert Into Ʊ��ʹ����ϸ
      (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
    Values
      (Ʊ��ʹ����ϸ_Id.Nextval, 2, Ʊ�ݺ�_In, 1, 1, ����id_In, v_��ӡid, v_Date, ����Ա����_In, ���_In);
    --״̬�Ķ�
    Update Ʊ�����ü�¼
    Set ��ǰ���� = Ʊ�ݺ�_In, ʣ������ = Decode(Sign(ʣ������ - 1), -1, 0, ʣ������ - 1), ʹ��ʱ�� = Sysdate
    Where ID = Nvl(����id_In, 0);
  End If;

  --��ػ��ܱ���

  --��Ա�ɿ����(����)
  If Nvl(���½������_In, 1) = 1 Then
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + ���_In
    Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = ���㷽ʽ_In
    Returning ��� Into n_����ֵ;
  
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (����Ա����_In, ���㷽ʽ_In, 1, ���_In);
      n_����ֵ := ���_In;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ��Ա�ɿ����
      Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = ���㷽ʽ_In And Nvl(���, 0) = 0;
    End If;
  End If;
  --����ʱ�����Ĵ���
  Select Nvl(��������, 0) Into v_���� From ������Ϣ Where ����id = ����id_In;
  If v_���� = 1 And Nvl(���_In, 0) > 0 Then
    Update ������Ϣ
    Set ������ = Decode(Sign(Nvl(������, 0) - Nvl(���_In, 0)), 1, Nvl(������, 0) - Nvl(���_In, 0), Null),
        ������ = Decode(Sign(Nvl(������, 0) - Nvl(���_In, 0)), 1, ������, Null),
        �������� = Decode(Sign(Nvl(������, 0) - Nvl(���_In, 0)), 1, ��������, Null)
    Where ����id = ����id_In;
  End If;
  If ��������_In <> 1 And ����id_In Is Null Then
    If ���_In < 0 Then
      b_Message.Zlhis_Charge_006(Id_In, ���ݺ�_In);
    Else
      b_Message.Zlhis_Charge_005(Id_In, ���ݺ�_In);
    End If;
    --��Ϣ����;
    Begin
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
        Using 11, Id_In;
    Exception
      When Others Then
        Null;
    End;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ԥ����¼_Insert;
/


Create Or Replace Procedure Zl_����Ԥ����¼_Modify
(
  Id_In           ����Ԥ����¼.Id%Type,
  ���㷽ʽ_In     ����Ԥ����¼.���㷽ʽ%Type,
  ������_In     ����Ԥ����¼.���%Type,
  �������_In     ����Ԥ����¼.�������%Type,
  ����_In         ����Ԥ����¼.����%Type,
  ������ˮ��_In   ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In     ����Ԥ����¼.����˵��%Type := Null,
  ����ժҪ_In     ����Ԥ����¼.ժҪ%Type,
  ����Ա����_In   ����Ԥ����¼.����Ա����%Type,
  ��ͨ����_In     Number := 0,
  Ԥ������Ʊ��_In ����Ԥ����¼.Ԥ������Ʊ��%Type := Null,
  ����_In         ���ս����¼.����%Type := Null
) As
  --����:���������ӿڷ�����Ϣ����Ԥ����¼
  --��ͨ����_In: 0-���濨���ID��1-�����濨���ID

  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  v_���㷽ʽ     ����Ԥ����¼.���㷽ʽ%Type;
  n_�����id     ����Ԥ����¼.�����id%Type;
  n_���         ����Ԥ����¼.���%Type;
  n_����ֵ       �������.Ԥ�����%Type;
  n_���         �������.Ԥ�����%Type;
  n_����id       ����Ԥ����¼.Id%Type;
  n_Ԥ�����     ����Ԥ����¼.Ԥ�����%Type;
  n_����id       ����Ԥ����¼.����id%Type;
  n_��¼״̬     ����Ԥ����¼.��¼״̬%Type;
  n_Ԥ������Ʊ�� Number(2);
  n_����         ���ս����¼.����%Type;
Begin

  n_Ԥ������Ʊ�� := Ԥ������Ʊ��_In;
  Begin
    Select ����id, ���㷽ʽ, �����id, ���, Ԥ�����, ����id, ��¼״̬
    Into n_����id, v_���㷽ʽ, n_�����id, n_���, n_Ԥ�����, n_����id, n_��¼״̬
    From ����Ԥ����¼
    Where ID = Id_In;
  Exception
    When Others Then
      v_Err_Msg := 'δ�ҵ��������ݣ����飡';
      Raise Err_Item;
  End;

  If n_Ԥ������Ʊ�� Is Null Then
    n_���� := ����_In;
    If ����_In Is Null Then
      Select Nvl(Max(����), 0) Into n_���� From ���ս����¼ Where ��¼id = Id_In And ���� = 3;
    End If;
    n_Ԥ������Ʊ�� := Zl_Fun_Isstarteinvoice(2, n_����, n_Ԥ�����);
  End If;

  If Nvl(��ͨ����_In, 0) = 1 Then
    n_�����id := Null;
  End If;
  Update ����Ԥ����¼
  Set ���㷽ʽ = Nvl(���㷽ʽ_In, ���㷽ʽ), ��� = Nvl(������_In, ���), ������� = Nvl(�������_In, �������), ժҪ = Nvl(����ժҪ_In, ժҪ),
      �����id = n_�����id, ���� = Nvl(����_In, ����), ������ˮ�� = Nvl(������ˮ��_In, ������ˮ��), ����˵�� = Nvl(����˵��_In, ����˵��), Ԥ������Ʊ�� = n_Ԥ������Ʊ��
  Where ID = Id_In;
  --�Զ�����̵���
  Zl_Custom_Balance_Update(Id_In);

  --������Ԥ���������
  If Nvl(n_��¼״̬, 0) = 1 And Nvl(n_����id, 0) = 0 Then
    If Nvl(������_In, 0) <> Nvl(n_���, 0) Then
      n_��� := Nvl(n_���, 0) - Nvl(������_In, 0);
      --����Ԥ���������
      Update Ԥ���������
      Set Ԥ����� = Nvl(Ԥ�����, 0) - n_���
      Where ����id = n_����id And Ԥ��id = Id_In
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into Ԥ��������� (Ԥ��id, ����id, Ԥ�����, Ԥ�����) Values (Id_In, n_����id, 1, Nvl(������_In, 0));
        n_����ֵ := Nvl(������_In, 0);
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From Ԥ��������� Where Ԥ��id = Id_In And Nvl(Ԥ�����, 0) = 0;
      End If;
      --�������
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) - n_���
      Where ���� = 1 And ����id = n_����id And Nvl(����, 0) = Nvl(n_Ԥ�����, 0)
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, ����, Ԥ�����, �������)
        Values
          (n_����id, 1, Nvl(n_Ԥ�����, 0), Nvl(������_In, 0), 0);
        n_����ֵ := Nvl(������_In, 0);
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From ������� Where ����id = n_����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
      End If;
    End If;
  End If;

  --��Ա�ɿ����
  If Nvl(���㷽ʽ_In, Nvl(v_���㷽ʽ, '-')) <> Nvl(v_���㷽ʽ, '-') Then
    --ԭ���㷽ʽ
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) - n_���
    Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = v_���㷽ʽ
    Returning ��� Into n_����ֵ;
  
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (����Ա����_In, v_���㷽ʽ, 1, -1 * n_���);
      n_����ֵ := -1 * n_���;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ��Ա�ɿ����
      Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ And Nvl(���, 0) = 0;
    End If;
    --�½��㷽ʽ
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + ������_In
    Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = ���㷽ʽ_In
    Returning ��� Into n_����ֵ;
  
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (����Ա����_In, ���㷽ʽ_In, 1, ������_In);
      n_����ֵ := ������_In;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ��Ա�ɿ����
      Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = ���㷽ʽ_In And Nvl(���, 0) = 0;
    End If;
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ԥ����¼_Modify;
/


Create Or Replace Procedure Zl_���˷���Ʊ��_Print
(
  No_In       Varchar2,
  Ʊ�ݺ�_In   Ʊ��ʹ����ϸ.����%Type,
  ����id_In   Ʊ��ʹ����ϸ.����id%Type,
  ʹ����_In   Ʊ��ʹ����ϸ.ʹ����%Type,
  ʹ��ʱ��_In Ʊ��ʹ����ϸ.ʹ��ʱ��%Type,
  ��������_In Number,
  Ʊ������_In Number := 1
) As
  --���ܣ�����ҽ�ƿ�ʹ������Ʊ��
  --������
  --      NO_IN       =     �����շѵĵ��ݺš���ʽΪ��A0000001
  --      Ʊ�ݺ�_IN   =     Ҫʹ�õĿ�ʼƱ�ݺš���Ʊ�ݺ�Ӧ�ò�Ϊ�գ�Ϊ��ʱ����������,�˿�ʱ����ԭʼƱ�ݺ�
  --      ����ID_IN   =     �ϸ����Ʊ��ʱ��Ϊʹ��Ʊ�ݵ��������Ρ����ϸ����ʱ��ΪNULL��
  --      Ʊ������_In =     ʵ�������Ʊ�ݴ�ӡ����
  --      ��������_In =     1-������2-�˿���3-�ش�4-����5-����
  --���α�����Ʊ�ݷ�Χ�ж�
  Cursor c_Fact Is
    Select * From Ʊ�����ü�¼ Where ID = Nvl(����id_In, 0);
  r_Factrow c_Fact%RowType;

  v_�ջ�id     Ʊ�ݴ�ӡ����.Id%Type;
  v_Ʊ�ݺ�     Ʊ��ʹ����ϸ.����%Type;
  v_��ǰƱ�ݺ� Ʊ��ʹ����ϸ.����%Type;
  n_��ӡid     Ʊ�ݴ�ӡ����.Id%Type;

  n_Ʊ�ݽ�� Ʊ��ʹ����ϸ.Ʊ�ݽ��%Type;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin

  --��Ʊ�ݺ�ʱ,���ô���Ʊ��
  If Ʊ�ݺ�_In Is Null Then
    Return;
  End If;

  --�˿�
  If ��������_In = 2 Then
    Begin
      --�����һ�δ�ӡ��������ȡ
      Select ID
      Into v_�ջ�id
      From (Select b.Id
             From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
             Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And a.Ʊ�� = 1 And b.�������� = 5 And b.No = No_In And
                   Not Exists
              (Select 1 From Ʊ��ʹ����ϸ B Where a.���� = b.���� And a.��ӡid = b.��ӡid And ���� = 2)
             Order By a.ʹ��ʱ�� Desc)
      Where Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
  
    If v_�ջ�id Is Not Null Then
      Insert Into Ʊ��ʹ����ϸ
        (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
        Select Ʊ��ʹ����ϸ_Id.Nextval, 1, Ʊ�ݺ�_In, 2, 2, ����id, ��ӡid, ʹ��ʱ��_In, ʹ����_In
        From Ʊ��ʹ����ϸ
        Where ��ӡid = v_�ջ�id And Ʊ�� = 1 And ���� = 1;
    End If;
    Return;
  End If;

  --�ش��ջ�ԭʼƱ��
  If ��������_In = 3 Or ��������_In = 5 Then
    Begin
      --�����һ�δ�ӡ��������ȡ
      Select ID
      Into v_�ջ�id
      From (Select b.Id
             From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
             Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And a.Ʊ�� = 1 And b.�������� = 5 And b.No = No_In And
                   Not Exists
              (Select 1 From Ʊ��ʹ����ϸ B Where a.���� = b.���� And a.��ӡid = b.��ӡid And ���� = 2)
             Order By a.ʹ��ʱ�� Desc)
      Where Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
  
    If v_�ջ�id Is Not Null Then
      Begin
        Insert Into Ʊ��ʹ����ϸ
          (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
          Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, Decode(��������_In, 5, 2, 4), ����id, ��ӡid, ʹ��ʱ��_In, ʹ����_In, Ʊ�ݽ��
          From Ʊ��ʹ����ϸ
          Where ��ӡid = v_�ջ�id And Ʊ�� = 1 And ���� = 1;
      Exception
        When Others Then
          Null;
      End;
    End If;
  End If;

  --Ʊ�ݴ�ӡ���
  Select Nvl(Sum(ʵ�ս��), 0) Into n_Ʊ�ݽ�� From סԺ���ü�¼ Where ��¼���� = 5 And NO = No_In;

  Select Ʊ�ݴ�ӡ����_Id.Nextval Into n_��ӡid From Dual;
  --���ɵ��ݵ�Ʊ�ݴ�ӡ����
  Insert Into Ʊ�ݴ�ӡ���� (ID, ��������, NO) Values (n_��ӡid, 5, No_In);

  --������Ʊ��
  v_Ʊ�ݺ� := Ʊ�ݺ�_In;
  If Nvl(����id_In, 0) <> 0 Then
    Open c_Fact;
    Fetch c_Fact
      Into r_Factrow;
    If c_Fact%RowCount = 0 Then
      v_Error := '��Ч��Ʊ���������Σ��޷���ɹҺ�Ʊ�ݷ��������';
      Close c_Fact;
      Raise Err_Custom;
    Elsif Nvl(r_Factrow.ʣ������, 0) < Ʊ������_In Then
      v_Error := '��ǰ���ε�ʣ����������' || Ʊ������_In || '�ţ��޷���ɹҺ�Ʊ�ݷ��������';
      Close c_Fact;
      Raise Err_Custom;
    End If;
  End If;
  For I In 1 .. Ʊ������_In Loop
    --���Ʊ�ݷ�Χ�Ƿ���ȷ
    If Nvl(����id_In, 0) <> 0 Then
      If Not (Upper(v_Ʊ�ݺ�) >= Upper(r_Factrow.��ʼ����) And Upper(v_Ʊ�ݺ�) <= Upper(r_Factrow.��ֹ����) And
          Length(v_Ʊ�ݺ�) = Length(r_Factrow.��ֹ����)) Then
        v_Error := '�õ�����Ҫ��ӡ����Ʊ��,��Ʊ�ݺ�"' || v_Ʊ�ݺ� || '"����Ʊ�����õĺ��뷶Χ��';
        Close c_Fact;
        Raise Err_Custom;
      End If;
    End If;
  
    --����Ʊ��
    Insert Into Ʊ��ʹ����ϸ
      (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
    Values
      (Ʊ��ʹ����ϸ_Id.Nextval, 1, v_Ʊ�ݺ�, 1, Decode(��������_In, 3, 3, 1), ����id_In, n_��ӡid, ʹ��ʱ��_In, ʹ����_In, n_Ʊ�ݽ��);
  
    v_��ǰƱ�ݺ� := v_Ʊ�ݺ�;
    --��һ��Ʊ�ݺ�
    v_Ʊ�ݺ� := Zl_Incstr(v_Ʊ�ݺ�);
  End Loop;

  If Nvl(����id_In, 0) <> 0 Then
    Update Ʊ�����ü�¼
    Set ʹ��ʱ�� = ʹ��ʱ��_In, ��ǰ���� = v_��ǰƱ�ݺ�, ʣ������ = Nvl(ʣ������, 0) - Ʊ������_In
    Where ID = ����id_In;
    Close c_Fact;
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˷���Ʊ��_Print;
/


Create Or Replace Procedure Zl_�����ӿ�����_Set
(
  �ӿ���_In �����ӿ�����.�ӿ���%Type,
  ����_In   �����ӿ�����.������%Type,
  ����ֵ_In �����ӿ�����.����ֵ%Type
) As
  v_Error Varchar2(255);
Begin
  If zl_To_Number(����_In) <> 0 Then
    Update �����ӿ����� Set ����ֵ = ����ֵ_In Where �ӿ��� = �ӿ���_In And ������ = zl_To_Number(����_In);
  Else
    Update �����ӿ����� Set ����ֵ = ����ֵ_In Where �ӿ��� = �ӿ���_In And ������ = ����_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����ӿ�����_Set;
/

Create Or Replace Function Zl_�����ӿ�����_Get
(
  �ӿ���_In �����ӿ�����.�ӿ���%Type,
  ����_In   �����ӿ�����.������%Type ) Return Varchar2  
 As 
  v_����ֵ  �����ӿ�����.����ֵ%Type;
Begin
  If zl_To_Number(����_In) <> 0 Then
    SELECT  ����ֵ INTO v_����ֵ From  �����ӿ����� Where �ӿ��� = �ӿ���_In And ������ = zl_To_Number(����_In);
  Else
    SELECT  ����ֵ INTO v_����ֵ From  �����ӿ����� Where �ӿ��� = �ӿ���_In And ������ = ����_In;
  End If;
  return v_����ֵ;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����ӿ�����_Get;
/
Create Or Replace Procedure Zl_����Ԥ����¼_תԤ��
(
  Ʊ�ݺ�_In     Ʊ��ʹ����ϸ.����%Type,
  ����id_In     ����Ԥ����¼.����id%Type,
  ��ҳid_In     ����Ԥ����¼.��ҳid%Type,
  ����id_In     ����Ԥ����¼.����id%Type,
  ���_In       ����Ԥ����¼.���%Type,
  ����Ա���_In ����Ԥ����¼.����Ա���%Type,
  ����Ա����_In ����Ԥ����¼.����Ա����%Type,
  �տ�ʱ��_In   ����Ԥ����¼.�տ�ʱ��%Type,
  ����id_In     Ʊ��ʹ����ϸ.����id%Type,
  Ԥ�����_In   ����Ԥ����¼.Ԥ�����%Type,
  ժҪ_In       ����Ԥ����¼.ժҪ%Type
) As
  ------------------------------------------------------------
  --Ԥ�����_In:1-����תסԺ;2-סԺת����
  ------------------------------------------------------------
  v_Err_Msg Varchar2(100);
  Err_Item Exception;

  v_��ӡid Ʊ�ݴ�ӡ����.Id%Type;
  n_����ֵ �������.Ԥ�����%Type;

  n_Id       ����Ԥ����¼.Id%Type;
  v_No       ����Ԥ����¼.No%Type;
  n_���     ����Ԥ����¼.���%Type;
  n_Ԥ��     ����Ԥ����¼.���%Type;
  l_Ʊ��no   t_StrList := t_StrList();
  n_Ԥ����� ����Ԥ����¼.Ԥ�����%Type;

  n_��id         ����Ԥ����¼.�ɿ���id%Type;
  d_�տ�ʱ��     ����Ԥ����¼.�տ�ʱ��%Type;
  n_Ԥ������Ʊ�� Number(2);
  n_����         ������Ϣ.����%Type;
  n_��������     Number(2);

  Procedure ����Ԥ����¼_Insert
  (
    Id_In           ����Ԥ����¼.Id%Type,
    ���ݺ�_In       ����Ԥ����¼.No%Type,
    ����id_In       ����Ԥ����¼.����id%Type,
    ��ҳid_In       ����Ԥ����¼.��ҳid%Type,
    ����id_In       ����Ԥ����¼.����id%Type,
    ��ֵ���_In     ����Ԥ����¼.���%Type,
    ���㷽ʽ_In     ����Ԥ����¼.���㷽ʽ%Type,
    �������_In     ����Ԥ����¼.�������%Type,
    �ɿλ_In     ����Ԥ����¼.�ɿλ%Type,
    ��λ������_In   ����Ԥ����¼.��λ������%Type,
    ��λ�ʺ�_In     ����Ԥ����¼.��λ�ʺ�%Type,
    ժҪ_In         ����Ԥ����¼.ժҪ%Type,
    ����Ա���_In   ����Ԥ����¼.����Ա���%Type,
    ����Ա����_In   ����Ԥ����¼.����Ա����%Type,
    Ԥ�����_In     ����Ԥ����¼.Ԥ�����%Type := Null,
    �����id_In     ����Ԥ����¼.�����id%Type := Null,
    ���㿨���_In   ����Ԥ����¼.���㿨���%Type := Null,
    ����_In         ����Ԥ����¼.����%Type := Null,
    ������ˮ��_In   ����Ԥ����¼.������ˮ��%Type := Null,
    ����˵��_In     ����Ԥ����¼.����˵��%Type := Null,
    ������λ_In     ����Ԥ����¼.������λ%Type := Null,
    �տ�ʱ��_In     ����Ԥ����¼.�տ�ʱ��%Type := Null,
    ��id_In         ����Ԥ����¼.�ɿ���id%Type := Null,
    ��������id_In   ����Ԥ����¼.��������id%Type := Null,
    Ԥ������Ʊ��_In ����Ԥ����¼.Ԥ������Ʊ��%Type
  ) As
    n_����ֵ �������.Ԥ�����%Type;
  Begin
  
    Insert Into ����Ԥ����¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ, �ɿ���id,
       Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, ��Ԥ��, ��������, У�Ա�־, ��������id, ����ʱ��, ������Ա, Ԥ������Ʊ��)
    Values
      (Id_In, ���ݺ�_In, Null, 1, 1, ����id_In, Decode(��ҳid_In, 0, Null, ��ҳid_In), Decode(����id_In, 0, Null, ����id_In),
       ��ֵ���_In, ���㷽ʽ_In, �������_In, �տ�ʱ��_In, �ɿλ_In, ��λ������_In, ��λ�ʺ�_In, ����Ա���_In, ����Ա����_In, ժҪ_In, ��id_In, Ԥ�����_In,
       �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, ������λ_In, Null, Null, Null, Null, ��������id_In, �տ�ʱ��_In, ����Ա����_In,
       Ԥ������Ʊ��_In);
  
    If Nvl(�����id_In, 0) <> 0 Then
      --�Զ�����̵���
      Zl_Custom_Balance_Update(Id_In);
    End If;
    --����Ԥ���������
    Insert Into Ԥ��������� (Ԥ��id, ����id, Ԥ�����, Ԥ�����) Values (Id_In, ����id_In, Ԥ�����_In, ��ֵ���_In);
  
    --�������(Ԥ���������)
    Update �������
    Set Ԥ����� = Nvl(Ԥ�����, 0) + ��ֵ���_In
    Where ���� = 1 And ����id = ����id_In And Nvl(����, 0) = Nvl(Ԥ�����_In, 0)
    Returning Ԥ����� Into n_����ֵ;
  
    If Sql%RowCount = 0 Then
      Insert Into �������
        (����id, ����, ����, Ԥ�����, �������)
      Values
        (����id_In, 1, Nvl(Ԥ�����_In, 0), ��ֵ���_In, 0);
      n_����ֵ := ��ֵ���_In;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ������� Where ����id = ����id_In And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
    End If;
    If ��ֵ���_In < 0 Then
      b_Message.Zlhis_Charge_006(Id_In, ���ݺ�_In);
    Else
      b_Message.Zlhis_Charge_005(Id_In, ���ݺ�_In);
    End If;
    --��Ϣ����;
    Begin
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
        Using 11, Id_In;
    Exception
      When Others Then
        Null;
    End;
  End ����Ԥ����¼_Insert;

  Procedure ����Ԥ����¼_Strict
  (
    Id_In           ����Ԥ����¼.Id%Type,
    ԭԤ��id_In     ����Ԥ����¼.Id%Type,
    ���ݺ�_In       ����Ԥ����¼.No%Type,
    ����id_In       ����Ԥ����¼.����id%Type,
    ��ҳid_In       ����Ԥ����¼.��ҳid%Type,
    ����id_In       ����Ԥ����¼.����id%Type,
    �������_In     ����Ԥ����¼.���%Type,
    ���㷽ʽ_In     ����Ԥ����¼.���㷽ʽ%Type,
    �������_In     ����Ԥ����¼.�������%Type,
    �ɿλ_In     ����Ԥ����¼.�ɿλ%Type,
    ��λ������_In   ����Ԥ����¼.��λ������%Type,
    ��λ�ʺ�_In     ����Ԥ����¼.��λ�ʺ�%Type,
    ժҪ_In         ����Ԥ����¼.ժҪ%Type,
    ����Ա���_In   ����Ԥ����¼.����Ա���%Type,
    ����Ա����_In   ����Ԥ����¼.����Ա����%Type,
    Ԥ�����_In     ����Ԥ����¼.Ԥ�����%Type := Null,
    �����id_In     ����Ԥ����¼.�����id%Type := Null,
    ���㿨���_In   ����Ԥ����¼.���㿨���%Type := Null,
    ����_In         ����Ԥ����¼.����%Type := Null,
    ������ˮ��_In   ����Ԥ����¼.������ˮ��%Type := Null,
    ����˵��_In     ����Ԥ����¼.����˵��%Type := Null,
    ������λ_In     ����Ԥ����¼.������λ%Type := Null,
    �տ�ʱ��_In     ����Ԥ����¼.�տ�ʱ��%Type := Null,
    ��������id_In   ����Ԥ����¼.��������id%Type := Null,
    �ɿ���id_In     ����Ԥ����¼.�ɿ���id%Type := Null,
    Ԥ������Ʊ��_In ����Ԥ����¼.Ԥ������Ʊ��%Type
  ) As
  
    n_����ֵ �������.Ԥ�����%Type;
    n_����id ���˽��ʼ�¼.Id%Type;
  Begin
  
    Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
  
    Insert Into ����Ԥ����¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ, �ɿ���id,
       Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, ��Ԥ��, ��������, У�Ա�־, ��������id, ����ʱ��, ������Ա, ���ӱ�־, Ԥ������Ʊ��)
    Values
      (Id_In, ���ݺ�_In, Null, 1, 1, ����id_In, Decode(��ҳid_In, 0, Null, ��ҳid_In), Decode(����id_In, 0, Null, ����id_In),
       �������_In, ���㷽ʽ_In, �������_In, �տ�ʱ��_In, �ɿλ_In, ��λ������_In, ��λ�ʺ�_In, ����Ա���_In, ����Ա����_In, ժҪ_In, �ɿ���id_In, Ԥ�����_In,
       �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, ������λ_In, n_����id, �������_In, Null, Null, ��������id_In, �տ�ʱ��_In, ����Ա����_In,
       Decode(Ԥ�����_In, 1, 2, 3), Ԥ������Ʊ��_In);
  
    If Nvl(�����id_In, 0) <> 0 Then
      --�Զ�����̵���
      Zl_Custom_Balance_Update(Id_In);
    End If;
  
    Update ����Ԥ����¼ Set ����id = n_����id, ��Ԥ�� = 0 Where ID = ԭԤ��id_In And ����id Is Null;
  
    Insert Into ����Ԥ����¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ, �ɿ���id,
       Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, ��Ԥ��, ��������, У�Ա�־, ��������id, ����ʱ��, ������Ա)
      Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���,
             ����Ա����, ժҪ, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, n_����id, Round(-1 * �������_In, 2), ��������, У�Ա�־,
             ��������id, ����ʱ��, ������Ա
      From ����Ԥ����¼
      Where ID = ԭԤ��id_In;
  
    --����Ԥ���������
    Update Ԥ���������
    Set Ԥ����� = Nvl(Ԥ�����, 0) + �������_In
    Where ����id = ����id_In And Ԥ��id = ԭԤ��id_In
    Returning Ԥ����� Into n_����ֵ;
    If Sql%RowCount = 0 Then
      Insert Into Ԥ��������� (Ԥ��id, ����id, Ԥ�����, Ԥ�����) Values (ԭԤ��id_In, ����id_In, 1, �������_In);
      n_����ֵ := �������_In;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From Ԥ��������� Where Ԥ��id = ԭԤ��id_In And Nvl(Ԥ�����, 0) = 0;
    End If;
  
    --�������(Ԥ���������)
    Update �������
    Set Ԥ����� = Nvl(Ԥ�����, 0) + �������_In
    Where ���� = 1 And ����id = ����id_In And Nvl(����, 0) = Nvl(Ԥ�����_In, 0)
    Returning Ԥ����� Into n_����ֵ;
  
    If Sql%RowCount = 0 Then
      Insert Into �������
        (����id, ����, ����, Ԥ�����, �������)
      Values
        (����id_In, 1, Nvl(Ԥ�����_In, 0), �������_In, 0);
      n_����ֵ := �������_In;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ������� Where ����id = ����id_In And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
    End If;
    If �������_In < 0 Then
      b_Message.Zlhis_Charge_006(Id_In, ���ݺ�_In);
    Else
      b_Message.Zlhis_Charge_005(Id_In, ���ݺ�_In);
    End If;
  
    --��Ϣ����;
    Begin
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
        Using 11, Id_In;
    Exception
      When Others Then
        Null;
    End;
  
  End ����Ԥ����¼_Strict;

Begin

  Select Nvl(Sum(Nvl(Ԥ�����, 0)), 0) Into n_����ֵ From ������� Where ����id = ����id_In And ���� = Ԥ�����_In;
  If Nvl(n_����ֵ, 0) - Nvl(���_In, 0) < 0 Then
    v_Err_Msg := '[ZLSOFT]' || Case
                   When Nvl(Ԥ�����_In, 0) = 1 Then
                    '����Ԥ��'
                   Else
                    'סԺԤ��'
                 End || '����![ZLSOFT]';
    Raise Err_Item;
  End If;

  n_��id := Zl_Get��id(����Ա����_In);

  d_�տ�ʱ�� := �տ�ʱ��_In;
  If d_�տ�ʱ�� Is Null Then
    Select Sysdate Into d_�տ�ʱ�� From Dual;
  End If;

  n_��� := ���_In;

  For v_Ԥ�� In (Select a.Id, a.���㷽ʽ, a.�������, a.�����id, a.���㿨���, a.����, a.������ˮ��, a.����˵��, a.������λ, b.Ԥ����� As ���, a.��������id,
                      a.�ɿλ, a.��λ������, a.��λ�ʺ�, a.Ԥ������Ʊ��
               From ����Ԥ����¼ A, Ԥ��������� B
               Where a.Id = b.Ԥ��id And b.����id = ����id_In And b.Ԥ����� = Ԥ�����_In And Not Exists
                (Select 1 From ���㷽ʽ Where a.���㷽ʽ = ���� And ���� = 5)
               Order By a.�տ�ʱ��) Loop
  
    Select ����Ԥ����¼_Id.Nextval Into n_Id From Dual;
    Select Decode(Ԥ�����_In, 1, 2, 1) Into n_Ԥ����� From Dual;
    v_No := Nextno(11);
    l_Ʊ��no.Extend;
    l_Ʊ��no(l_Ʊ��no.Count) := v_No;
  
    n_Ԥ�� := Nvl(v_Ԥ��.���, 0);
    If n_��� < Nvl(v_Ԥ��.���, 0) Then
      n_Ԥ�� := n_���;
    End If;
    n_��� := n_��� - n_Ԥ��;
    Select ���� Into n_�������� From ���㷽ʽ Where ���� = v_Ԥ��.���㷽ʽ;
    If n_�������� = 3 Then
      Select To_Number(����) Into n_���� From ������Ϣ Where ����id = ����id_In;
    Else
      n_���� := 0;
    End If;
    n_Ԥ������Ʊ�� := Zl_Fun_Isstarteinvoice(2, n_����, n_Ԥ�����);
  
    --1.����һ����ֵ��¼
    ����Ԥ����¼_Insert(n_Id, v_No, ����id_In, ��ҳid_In, ����id_In, n_Ԥ��, v_Ԥ��.���㷽ʽ, v_Ԥ��.�������, v_Ԥ��.�ɿλ, v_Ԥ��.��λ������, v_Ԥ��.��λ�ʺ�,
                  ժҪ_In, ����Ա���_In, ����Ա����_In, n_Ԥ�����, v_Ԥ��.�����id, v_Ԥ��.���㿨���, v_Ԥ��.����, v_Ԥ��.������ˮ��, v_Ԥ��.����˵��, v_Ԥ��.������λ,
                  d_�տ�ʱ��, n_��id, v_Ԥ��.��������id, n_Ԥ������Ʊ��);
  
    Update ����Ԥ����¼ Set ʵ��Ʊ�� = Ʊ�ݺ�_In Where ID = n_Id;
  
    v_No := Nextno(11);
    l_Ʊ��no.Extend;
    l_Ʊ��no(l_Ʊ��no.Count) := v_No;
  
    Select ����Ԥ����¼_Id.Nextval Into n_Id From Dual;
    Select Decode(Ԥ�����_In, 1, 1, 2) Into n_Ԥ����� From Dual;
    --2.��ԭ��Ԥ��
    ����Ԥ����¼_Strict(n_Id, v_Ԥ��.Id, v_No, ����id_In, ��ҳid_In, ����id_In, -1 * n_Ԥ��, v_Ԥ��.���㷽ʽ, v_Ԥ��.�������, v_Ԥ��.�ɿλ,
                  v_Ԥ��.��λ������, v_Ԥ��.��λ�ʺ�, ժҪ_In, ����Ա���_In, ����Ա����_In, n_Ԥ�����, v_Ԥ��.�����id, v_Ԥ��.���㿨���, v_Ԥ��.����, v_Ԥ��.������ˮ��,
                  v_Ԥ��.����˵��, v_Ԥ��.������λ, d_�տ�ʱ��, v_Ԥ��.��������id, n_��id, v_Ԥ��.Ԥ������Ʊ��);
  
    Update ����Ԥ����¼
    Set ʵ��Ʊ�� = Ʊ�ݺ�_In
    Where NO = (Select NO From ����Ԥ����¼ Where ID = n_Id) And ��¼���� In (1, 11);
  
    If n_��� <= 0 Then
      Exit;
    End If;
  End Loop;

  If Nvl(n_���, 0) <> 0 Then
    v_Err_Msg := '[ZLSOFT]' || Case
                   When Nvl(Ԥ�����_In, 0) = 1 Then
                    '����Ԥ��'
                   Else
                    'סԺԤ��'
                 End || '����![ZLSOFT]';
    Raise Err_Item;
  End If;

  --����Ʊ��
  If Ʊ�ݺ�_In Is Not Null Then
    Select Ʊ�ݴ�ӡ����_Id.Nextval Into v_��ӡid From Dual;
  
    --����Ʊ��
    Insert Into Ʊ�ݴ�ӡ����
      (ID, ��������, NO)
      Select v_��ӡid, 2, Column_Value From Table(l_Ʊ��no);
  
    Insert Into Ʊ��ʹ����ϸ
      (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
    Values
      (Ʊ��ʹ����ϸ_Id.Nextval, 2, Ʊ�ݺ�_In, 1, 1, ����id_In, v_��ӡid, �տ�ʱ��_In, ����Ա����_In, ���_In);
    --״̬�Ķ�
    Update Ʊ�����ü�¼
    Set ��ǰ���� = Ʊ�ݺ�_In, ʣ������ = Decode(Sign(ʣ������ - 1), -1, 0, ʣ������ - 1), ʹ��ʱ�� = Sysdate
    Where ID = Nvl(����id_In, 0);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ԥ����¼_תԤ��;
/

Create Or Replace Procedure Zl_����Ԥ����¼_����˿�
(
  Id_In           ����Ԥ����¼.Id%Type,
  ���ݺ�_In       ����Ԥ����¼.No%Type,
  Ʊ�ݺ�_In       Ʊ��ʹ����ϸ.����%Type,
  ����id_In       ����Ԥ����¼.����id%Type,
  ��ҳid_In       ����Ԥ����¼.��ҳid%Type,
  ����id_In       ����Ԥ����¼.����id%Type,
  ���_In         ����Ԥ����¼.���%Type,
  ���㷽ʽ_In     ����Ԥ����¼.���㷽ʽ%Type,
  �������_In     ����Ԥ����¼.�������%Type,
  �ɿλ_In     ����Ԥ����¼.�ɿλ%Type,
  ��λ������_In   ����Ԥ����¼.��λ������%Type,
  ��λ�ʺ�_In     ����Ԥ����¼.��λ�ʺ�%Type,
  ժҪ_In         ����Ԥ����¼.ժҪ%Type,
  ����Ա���_In   ����Ԥ����¼.����Ա���%Type,
  ����Ա����_In   ����Ԥ����¼.����Ա����%Type,
  ����id_In       Ʊ��ʹ����ϸ.����id%Type,
  Ԥ�����_In     ����Ԥ����¼.Ԥ�����%Type := Null,
  �����id_In     ����Ԥ����¼.�����id%Type := Null,
  ���㿨���_In   ����Ԥ����¼.���㿨���%Type := Null,
  ����_In         ����Ԥ����¼.����%Type := Null,
  ��������id_In   ����Ԥ����¼.��������id%Type := Null,
  ������ˮ��_In   ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In     ����Ԥ����¼.����˵��%Type := Null,
  ������λ_In     ����Ԥ����¼.������λ%Type := Null,
  �տ�ʱ��_In     ����Ԥ����¼.�տ�ʱ��%Type := Null,
  У�Ա�־_In     ����Ԥ����¼.У�Ա�־%Type := Null,
  ������Ϣ_In     Varchar2 := Null,
  ����������_In   Number := 0,
  ����״̬_In     Number := 0,
  ����id_In       ����Ԥ����¼.����id%Type := Null,
  �������_In     ����Ԥ����¼.�������%Type := Null,
  Ԥ������Ʊ��_In ����Ԥ����¼.Ԥ������Ʊ��%Type := Null,
  ����_In         ���ս����¼.����%Type := Null
) As
  ----------------------------------------------
  --����˿����
  --������Ϣ_In:ԭԤ��ID|���||....
  --����������_IN:0-��ʾ��Ҫ����Ԥ����¼�����²������;1-��ʾֻ���½�����Ϣ�е���������
  --����״̬_IN:0-��ʾ��ɽ���;1-��ʾδ��ɽ���;
  v_Err_Msg Varchar2(200);
  Err_Item Exception;

  n_��ӡid       Ʊ�ݴ�ӡ����.Id%Type;
  v_����         ������Ϣ.��������%Type;
  d_�տ�ʱ��     Date;
  n_����ֵ       �������.Ԥ�����%Type;
  n_��id         ����ɿ����.Id%Type;
  n_����id       ����Ԥ����¼.����id%Type;
  n_�������     ����Ԥ����¼.�������%Type;
  n_Count        Number(18);
  n_Ԥ�����     �������.Ԥ�����%Type;
  n_Ԥ������Ʊ�� Number(2);
  n_����         ���ս����¼.����%Type;

Begin

  n_Ԥ������Ʊ�� := Ԥ������Ʊ��_In;
  If n_Ԥ������Ʊ�� Is Null Then
    n_���� := ����_In;
    If ����_In Is Null Then
      Select Nvl(Max(����), 0) Into n_���� From ���ս����¼ Where ��¼id = Id_In And ���� = 3;
    End If;
    n_Ԥ������Ʊ�� := Zl_Fun_Isstarteinvoice(2, n_����, Ԥ�����_In);
  End If;
  n_��id := Zl_Get��id(����Ա����_In);
  If ����������_In = 0 Then
    d_�տ�ʱ�� := �տ�ʱ��_In;
    If d_�տ�ʱ�� Is Null Then
      Select Sysdate Into d_�տ�ʱ�� From Dual;
    End If;
    n_������� := �������_In;
    n_����id   := ����id_In;
    If Nvl(n_����id, 0) = 0 Then
      Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
    End If;
    If Nvl(n_�������, 0) = 0 Then
      n_������� := -1 * n_����id;
    End If;
    --Ϊ�˲������������������(���_InΪ����)
    Update �������
    Set Ԥ����� = Nvl(Ԥ�����, 0) + ���_In
    Where ����id = ����id_In And ���� = Ԥ�����_In And ���� = 1
    Returning Ԥ����� Into n_Ԥ�����;
    If Sql%RowCount = 0 Then
      Insert Into �������
        (����id, ����, ����, Ԥ�����, �������)
      Values
        (����id_In, 1, Nvl(Ԥ�����_In, 0), ���_In, 0);
      n_Ԥ����� := ���_In;
    End If;
  
    Insert Into ����Ԥ����¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ, �ɿ���id,
       Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��Ԥ��, ��������, У�Ա�־, ��������id, ����ʱ��, ������Ա, ���ӱ�־, Ԥ������Ʊ��)
    Values
      (Id_In, ���ݺ�_In, Decode(����״̬_In, 0, Ʊ�ݺ�_In, Null), 1, 0, ����id_In, Decode(��ҳid_In, 0, Null, ��ҳid_In),
       Decode(����id_In, 0, Null, ����id_In), ���_In, ���㷽ʽ_In, �������_In, d_�տ�ʱ��, �ɿλ_In, ��λ������_In, ��λ�ʺ�_In, ����Ա���_In,
       ����Ա����_In, ժҪ_In, n_��id, Ԥ�����_In, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, ������λ_In, n_����id, n_�������, Null,
       Null, У�Ա�־_In, Decode(Nvl(��������id_In, 0), 0, Id_In, ��������id_In), �տ�ʱ��_In, ����Ա����_In, 1, n_Ԥ������Ʊ��);
  
    --����Ԥ���������
    Insert Into Ԥ��������� (Ԥ��id, ����id, Ԥ�����, Ԥ�����) Values (Id_In, ����id_In, Ԥ�����_In, ���_In);
  End If;

  If ����������_In = 1 Then
    Select Max(����id), Max(�տ�ʱ��), Max(1) Into n_����id, d_�տ�ʱ��, n_Count From ����Ԥ����¼ Where ID = Id_In;
    If n_Count = 0 Then
      v_Err_Msg := 'δ�ҵ��˿��¼�����飡';
      Raise Err_Item;
    End If;
  End If;

  If ������Ϣ_In Is Not Null Then
    Zl_����Ԥ����¼_Relevance(����id_In, Id_In, ������Ϣ_In, n_����id, ����Ա���_In, ����Ա����_In, �տ�ʱ��_In, У�Ա�־_In, n_��id);
  End If;

  If ����״̬_In = 1 Then
    Return;
  End If;

  --���¼�¼״̬1
  Update ����Ԥ����¼
  Set ��¼״̬ = 1, У�Ա�־ = 0, ʵ��Ʊ�� = Ʊ�ݺ�_In
  Where NO = ���ݺ�_In And ��¼���� = 1 And ��¼״̬ = 0
  Returning ����id Into n_����id;
  If Sql%NotFound Then
    v_Err_Msg := 'δ�ҵ�ָ���ĵ���(' || ���ݺ�_In || ',������Ϊ����ԭ�������˿���飡';
    Raise Err_Item;
  End If;
  Update ����Ԥ����¼ Set У�Ա�־ = 0 Where ����id = n_����id And Nvl(У�Ա�־, 0) <> 0;

  --����Ʊ��
  If Ʊ�ݺ�_In Is Not Null Then
    Select Ʊ�ݴ�ӡ����_Id.Nextval Into n_��ӡid From Dual;
  
    --����Ʊ��
    Insert Into Ʊ�ݴ�ӡ���� (ID, ��������, NO) Values (n_��ӡid, 2, ���ݺ�_In);
  
    Insert Into Ʊ��ʹ����ϸ
      (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
    Values
      (Ʊ��ʹ����ϸ_Id.Nextval, 2, Ʊ�ݺ�_In, 1, 1, ����id_In, n_��ӡid, d_�տ�ʱ��, ����Ա����_In, ���_In);
    --״̬�Ķ�
    Update Ʊ�����ü�¼
    Set ��ǰ���� = Ʊ�ݺ�_In, ʣ������ = Decode(Sign(ʣ������ - 1), -1, 0, ʣ������ - 1), ʹ��ʱ�� = Sysdate
    Where ID = Nvl(����id_In, 0);
  
    Update ����Ԥ����¼ Set ʵ��Ʊ�� = Ʊ�ݺ�_In Where ����id = ����id_In And ��¼���� = 11 And NO = ���ݺ�_In;
  
  End If;

  --��Ա�ɿ����(����)
  Update ��Ա�ɿ����
  Set ��� = Nvl(���, 0) + ���_In
  Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = ���㷽ʽ_In
  Returning ��� Into n_����ֵ;

  If Sql%RowCount = 0 Then
    Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (����Ա����_In, ���㷽ʽ_In, 1, ���_In);
    n_����ֵ := ���_In;
  End If;
  If Nvl(n_����ֵ, 0) = 0 Then
    Delete From ��Ա�ɿ���� Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = ���㷽ʽ_In And Nvl(���, 0) = 0;
  End If;

  --����ʱ�����Ĵ���
  Select Nvl(��������, 0) Into v_���� From ������Ϣ Where ����id = ����id_In;

  If v_���� = 1 And Nvl(���_In, 0) > 0 Then
    Update ������Ϣ
    Set ������ = Decode(Sign(Nvl(������, 0) - Nvl(���_In, 0)), 1, Nvl(������, 0) - Nvl(���_In, 0), Null),
        ������ = Decode(Sign(Nvl(������, 0) - Nvl(���_In, 0)), 1, ������, Null),
        �������� = Decode(Sign(Nvl(������, 0) - Nvl(���_In, 0)), 1, ��������, Null)
    Where ����id = ����id_In;
  End If;

  If Nvl(n_Ԥ�����, 0) = 0 Then
    Delete From �������
    Where ����id = ����id_In And ���� = Ԥ�����_In And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0 And ���� = 1;
  End If;

  If ���_In < 0 Then
    b_Message.Zlhis_Charge_006(Id_In, ���ݺ�_In);
  Else
    b_Message.Zlhis_Charge_005(Id_In, ���ݺ�_In);
  End If;

  --��Ϣ����;
  Begin
    Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
      Using 11, Id_In;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ԥ����¼_����˿�;
/

Create Or Replace Package b_Einvoice_Request Is
  ------------------------------------------------------------------
  --����Ʊ��ҵ����
  --  1.Einvoice_Start-�жϵ���Ʊ���Ƿ�����(����:1-����;0-δ����)
  --  2.EInvoice_Create-����Ʊ�ݿ���(����1-�ɹ�;0-ʧ��)
  --  3.Einvoice_Cancel_Check-����Ʊ������ǰ���(����:1-�Ϸ�;0-���Ϸ�)
  --  4.Einvoice_Cancel-����Ʊ������(����1-�ɹ�;0-ʧ��)
  ------------------------------------------------------------------
  --1.�жϵ���Ʊ���Ƿ�����
  Function Einvoice_Start
  (
    ҵ�񳡾�_In Integer,
    ����_In     ���ս����¼.����%Type,
    ����_In     Integer := Null
  ) Return Number;

  --2.����Ʊ�ݿ���
  Function Einvoice_Create
  (
    ҵ�񳡾�_In  Integer,
    ����id_In    ����Ԥ����¼.����id%Type,
    ����id_In    ����Ԥ����¼.����id%Type,
    ������Ϣ_Out Out Varchar2
  ) Return Number;

  --3.����Ʊ�����ϼ��
  Function Einvoice_Cancel_Check
  (
    ҵ�񳡾�_In  Integer,
    ����id_In    ����Ԥ����¼.����id%Type,
    ������Ϣ_Out Out Varchar2
  ) Return Number;

  --4.����Ʊ������
  Function Einvoice_Cancel
  (
    ҵ�񳡾�_In  Integer,
    ����id_In    ����Ԥ����¼.����id%Type,
    ������Ϣ_Out Out Varchar2
  ) Return Number;
End b_Einvoice_Request;
/

Create Or Replace Package Body b_Einvoice_Request Is
  ------------------------------------------------------------------
  --����Ʊ��ҵ����
  --  1.Einvoice_Start-�жϵ���Ʊ���Ƿ�����(����:1-����;0-δ����)
  --  2.EInvoice_Create-����Ʊ�ݿ���(����1-�ɹ�;0-ʧ��)
  --  3.Einvoice_Cancel_Check-����Ʊ������ǰ���(����:1-�Ϸ�;0-���Ϸ�)
  --  4.Einvoice_Cancel-����Ʊ������(����1-�ɹ�;0-ʧ��)
  ------------------------------------------------------------------

  Function Einvoice_Start
  (
    ҵ�񳡾�_In Integer,
    ����_In     ���ս����¼.����%Type,
    ����_In     Integer := Null
  ) Return Number Is
    ------------------------------------------------------------------
    --����:�жϵ���Ʊ���Ƿ�����
    --���:ҵ�񳡾�_In-1-�շ�,2-Ԥ��,3-����,4-�Һ�;5-���￨
    --     ����_in:NULL-����������;��Գ���Ϊ���˼�Ԥ��:1-����;2-סԺ;
    --����:������Ϣ_Out-���صĴ�����Ϣ
    --����:1-����;0-δ����
    -------------------------------------------------------------------
    v_������   ����Ʊ�����.������%Type;
    v_Sql      Varchar2(1000);
    n_Return   Number(2);
    n_����     Number(2);
    n_Err_Code Number(18);
    Err_Item   Exception;
    Err_Custom Exception;
  Begin
  
    If Nvl(ҵ�񳡾�_In, 0) = 2 And Nvl(����_In, 0) = 1 Then
      --����Ԥ�����ݲ�֧��
      Return 0;
    End If;
  
    Begin
      n_���� := 1;
      Select ������ Into v_������ From ����Ʊ����� Where �Ƿ����� = 1;
    Exception
      When Others Then
        n_���� := 0;
    End;
    If n_���� = 0 Or v_������ Is Null Then
      --δ���û��ް����ƣ�ֱ�ӷ���0����ʾ�ɹ�;
      Return 0;
    End If;
  
    n_Err_Code := Null;
    Begin
      v_Sql := 'begin :n_return:=' || v_������ || '.Einvoice_Start(:1,:2,:3); end;';
      Execute Immediate v_Sql
        Using Out n_Return, In ҵ�񳡾�_In, ����_In, ����_In;
      Return n_Return;
    Exception
      When Others Then
        n_Err_Code := SQLCode;
    End;
    If n_Err_Code = -6550 Or n_Err_Code Is Null Then
      Return 0;
    End If;
    Return 0;
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Einvoice_Start;

  Function Einvoice_Create
  (
    ҵ�񳡾�_In  Integer,
    ����id_In    ����Ԥ����¼.����id%Type,
    ����id_In    ����Ԥ����¼.����id%Type,
    ������Ϣ_Out Out Varchar2
  ) Return Number Is
    ------------------------------------------------------------------
    --����:����Ʊ�ݿ���
    --���:ҵ�񳡾�_In-1-�շ�,2-Ԥ��,3-����,4-�Һ�;5-���￨
    --     ����ID_In-ҵ�񳡾�_In=2(Ԥ��)ʱ��ԭԤ��ID,ҵ�񳡾�_In<>2(Ԥ��)ʱ��ԭ����ID
    --     ����ID_In-ҵ�񳡾�_In=2(Ԥ��)ʱ���˿��Ԥ��ID,ҵ�񳡾�_In<>2(Ԥ��)ʱ����ǰ�˷ѵĽ���ID,�����˷�ʱ��Ч;
    --����:������Ϣ_Out-���صĴ�����Ϣ
    --����:1-�ɹ�;0-ʧ��
    -------------------------------------------------------------------
    v_������      ����Ʊ�����.������%Type;
    v_Sql         Varchar2(1000);
    n_Return      Number(2);
    v_Err_Msg_Out Varchar2(4000);
    n_����        Number(2);
    n_Err_Code    Number(18);
    Err_Item   Exception;
    Err_Custom Exception;
  Begin
  
    Begin
      n_���� := 1;
      Select ������ Into v_������ From ����Ʊ����� Where �Ƿ����� = 1;
    Exception
      When Others Then
        n_���� := 0;
    End;
    If n_���� = 0 Or v_������ Is Null Then
      --δ���û��ް����ƣ�ֱ�ӷ���1����ʾ�ɹ�;
      Return 1;
    End If;
  
    n_Err_Code := Null;
    Begin
      v_Sql := 'begin :n_return:=' || v_������ || '.EInvoice_Create(:1,:2,:3,:v_Err_Msg_out); end;';
      Execute Immediate v_Sql
        Using Out n_Return, In ҵ�񳡾�_In, ����id_In, ����id_In, Out v_Err_Msg_Out;
      ������Ϣ_Out := v_Err_Msg_Out;
      Return n_Return;
    Exception
      When Others Then
        n_Err_Code    := SQLCode;
        v_Err_Msg_Out := SQLErrM;
    End;
    If n_Err_Code = -6550 Or n_Err_Code Is Null Then
      --û�д˹��̣�����true
      Return 1;
    End If;
    Raise Err_Item;
  
  Exception
    When Err_Item Then
      zl_ErrorCenter(n_Err_Code, v_Err_Msg_Out);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Einvoice_Create;

  Function Einvoice_Cancel_Check
  (
    ҵ�񳡾�_In  Integer,
    ����id_In    ����Ԥ����¼.����id%Type,
    ������Ϣ_Out Out Varchar2
  ) Return Number Is
    ------------------------------------------------------------------
    --����:����Ʊ������
    --���:ҵ�񳡾�_In-1-�շ�,2-Ԥ��,3-����,4-�Һ�;5-���￨
    --     ����ID_In-ҵ�񳡾�_In=2(Ԥ��)ʱ��ԭԤ��ID,ҵ�񳡾�_In<>2(Ԥ��)ʱ��ԭ����ID 
    --����:������Ϣ_Out-���صĴ�����Ϣ
    --����:1-�ɹ�;0-ʧ��
    -------------------------------------------------------------------
    v_������      ����Ʊ�����.������%Type;
    v_Sql         Varchar2(1000);
    n_Return      Number(2);
    v_Err_Msg_Out Varchar2(4000);
    n_����        Number(2);
    n_Err_Code    Number(18);
    Err_Item   Exception;
    Err_Custom Exception;
  Begin
  
    If ҵ�񳡾�_In = 2 Then
      --Ԥ����
      Select Max(Nvl(Ԥ������Ʊ��, 0)) Into n_Return From ����Ԥ����¼ Where ����id = ����id_In;
    Else
      --��Ԥ�����շѡ����ʡ��Һż����￨
      Select Max(Nvl(�Ƿ����Ʊ��, 0)) Into n_Return From ����Ԥ����¼ Where ����id = ����id_In;
    End If;
    If Nvl(n_Return, 0) = 0 Then
      --�ü�¼��δ���õ���Ʊ�ݵģ�ֱ�ӷ���1;
      Return 1;
    End If;
  
    Begin
      n_���� := 1;
      Select ������ Into v_������ From ����Ʊ����� Where �Ƿ����� = 1;
    Exception
      When Others Then
        n_���� := 0;
    End;
  
    If n_���� = 0 Or v_������ Is Null Then
      ������Ϣ_Out := '����Ʊ��δ���ã����ڴ����н����˷ѻ��˿�ڴ˴���ֹ�������Ʊ�ݳ�졣';
      Return 0;
    End If;
  
    --����Ƿ񻻿�������Ʊ��
    For c_����Ʊ�� In (Select ID, �Ƿ񻻿�, ֽ�ʷ�Ʊ��
                   From ����Ʊ��ʹ�ü�¼
                   Where Ʊ�� = ҵ�񳡾�_In And ��¼״̬ = 1 And ����id = ����id_In) Loop
      --��Ե���Ʊ�ݽ��д���
      If Nvl(c_����Ʊ��.�Ƿ񻻿�, 0) = 1 Then
        --����ֽ�ʷ�Ʊ�ţ���ֹ���ϲ���
        ������Ϣ_Out := '�Ѿ�����ֽ�ʷ�Ʊ(' || c_����Ʊ��.ֽ�ʷ�Ʊ�� || ')���ڴ˴���ֹ�������Ʊ�ݳ�졣';
        Return 0;
      End If;
      n_Err_Code := Null;
      Begin
        v_Sql := 'begin :n_return:=' || v_������ || '.Einvoice_Cancel_Check(:1,:2,:v_Err_Msg_out); end;';
        Execute Immediate v_Sql
          Using Out n_Return, In ҵ�񳡾�_In, c_����Ʊ��.Id, Out v_Err_Msg_Out;
        ������Ϣ_Out := v_Err_Msg_Out;
        Return n_Return;
      Exception
        When Others Then
          n_Err_Code    := SQLCode;
          v_Err_Msg_Out := SQLErrM;
      End;
    
      If n_Err_Code = -6550 Then
        ������Ϣ_Out := '����Ʊ��δ���ã����ڴ����н����˷ѻ��˿�ڴ˴���ֹ�������Ʊ�ݳ�졣';
        Return 0;
      End If;
      If Not n_Err_Code Is Null Then
        Raise Err_Item;
      End If;
    
    End Loop;
    Return 1;
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg_Out || '[ZLSOFT]');
    When Err_Item Then
      zl_ErrorCenter(n_Err_Code, v_Err_Msg_Out);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Einvoice_Cancel_Check;

  Function Einvoice_Cancel
  (
    ҵ�񳡾�_In  Integer,
    ����id_In    ����Ԥ����¼.����id%Type,
    ������Ϣ_Out Out Varchar2
  ) Return Number Is
    ------------------------------------------------------------------
    --����:����Ʊ������
    --���:ҵ�񳡾�_In-1-�շ�,2-Ԥ��,3-����,4-�Һ�;5-���￨
    --     ����ID_In-ҵ�񳡾�_In=2(Ԥ��)ʱ��ԭԤ��ID,ҵ�񳡾�_In<>2(Ԥ��)ʱ��ԭ����ID 
    --����:������Ϣ_Out-���صĴ�����Ϣ
    --����:1-�ɹ�;0-ʧ��
    -------------------------------------------------------------------
    v_������      ����Ʊ�����.������%Type;
    v_Sql         Varchar2(1000);
    n_Return      Number(2);
    v_Err_Msg_Out Varchar2(4000);
    n_����        Number(2);
    n_Err_Code    Number(18);
    Err_Item   Exception;
    Err_Custom Exception;
  Begin
  
    If ҵ�񳡾�_In = 2 Then
      --Ԥ����
      Select Max(Nvl(Ԥ������Ʊ��, 0)) Into n_Return From ����Ԥ����¼ Where ����id = ����id_In;
    Else
      --��Ԥ�����շѡ����ʡ��Һż����￨
      Select Max(Nvl(�Ƿ����Ʊ��, 0)) Into n_Return From ����Ԥ����¼ Where ����id = ����id_In;
    End If;
    If Nvl(n_Return, 0) = 0 Then
      --�ü�¼��δ���õ���Ʊ�ݵģ�ֱ�ӷ���1;
      Return 1;
    End If;
  
    Begin
      n_���� := 1;
      Select ������ Into v_������ From ����Ʊ����� Where �Ƿ����� = 1;
    Exception
      When Others Then
        n_���� := 0;
    End;
  
    If n_���� = 0 Or v_������ Is Null Then
      ������Ϣ_Out := '����Ʊ��δ���ã����ڴ����н����˷ѻ��˿�ڴ˴���ֹ�������Ʊ�ݳ�졣';
      Return 0;
    End If;
  
    --����Ƿ񻻿�������Ʊ��
    For c_����Ʊ�� In (Select ID, �Ƿ񻻿�, ֽ�ʷ�Ʊ��
                   From ����Ʊ��ʹ�ü�¼
                   Where Ʊ�� = ҵ�񳡾�_In And ��¼״̬ = 1 And ����id = ����id_In) Loop
      --��Ե���Ʊ�ݽ��д���
      If Nvl(c_����Ʊ��.�Ƿ񻻿�, 0) = 1 Then
        --����ֽ�ʷ�Ʊ�ţ���ֹ���ϲ���
        ������Ϣ_Out := '�Ѿ�����ֽ�ʷ�Ʊ(' || c_����Ʊ��.ֽ�ʷ�Ʊ�� || ')���ڴ˴���ֹ�������Ʊ�ݳ�졣';
        Return 0;
      End If;
      n_Err_Code := Null;
    
      --���Ⲣ��ԭ�򣬻�����Ҫ�Ƚ��м�����Ʊ���Ƿ������졣
      Begin
        v_Sql := 'begin :n_return:=' || v_������ || '.Einvoice_Cancel_Check(:1,:2,:v_Err_Msg_out); end;';
        Execute Immediate v_Sql
          Using Out n_Return, In ҵ�񳡾�_In, c_����Ʊ��.Id, Out v_Err_Msg_Out;
        ������Ϣ_Out := v_Err_Msg_Out;
        Return n_Return;
      Exception
        When Others Then
          n_Err_Code    := SQLCode;
          v_Err_Msg_Out := SQLErrM;
      End;
    
      If n_Err_Code = -6550 Then
        ������Ϣ_Out := '����Ʊ��δ���ã����ڴ����н����˷ѻ��˿�ڴ˴���ֹ�������Ʊ�ݳ�졣';
        Return 0;
      End If;
      If Not n_Err_Code Is Null Then
        Raise Err_Item;
      End If;
    
      --���е���Ʊ�ݳ�촦��
      Begin
        v_Sql := 'begin :n_return:=' || v_������ || '.Einvoice_Cancel(:1,:2,:v_Err_Msg_out); end;';
        Execute Immediate v_Sql
          Using Out n_Return, In ҵ�񳡾�_In, c_����Ʊ��.Id, Out v_Err_Msg_Out;
        ������Ϣ_Out := v_Err_Msg_Out;
        Return n_Return;
      Exception
        When Others Then
          n_Err_Code    := SQLCode;
          v_Err_Msg_Out := SQLErrM;
      End;
    
      If n_Err_Code = -6550 Then
        ������Ϣ_Out := '����Ʊ��δ���ã����ڴ����н����˷ѻ��˿�ڴ˴���ֹ�������Ʊ�ݳ�졣';
        Return 0;
      End If;
      If Not n_Err_Code Is Null Then
        Raise Err_Item;
      End If;
    End Loop;
    Return 1;
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg_Out || '[ZLSOFT]');
    When Err_Item Then
      zl_ErrorCenter(n_Err_Code, v_Err_Msg_Out);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Einvoice_Cancel;
End b_Einvoice_Request;
/


Create Or Replace Procedure Zl_Third_Regist
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --����:HIS�Һ�
  --���:Xml_In:
  --<IN>
  --   <CZFS>3</CZFS>    //������ʽ
  --   <CZJLID>1</CZJLID>    //�����¼ID
  --   <HM>����</HM>    //����
  --   <HX>����</HX>     //����
  --   <JKFS>0</JKFS>  //�ɿʽ,0-�ҺŻ�ԤԼ�ɿ�;1-ԤԼ���ɿ�
  --   <YYSJ>2014-10-21 </YYSJ>    //ԤԼ���� YYYY-MM-DD,��ʱ�η���ſ�����Ҫ����ʱ��
  --   <JE>���</JE>     //���
  --   <HZDW>������λ</HZDW>        //������λ����
  --   <YYFS>֧����<YYFS>    //ԤԼ��ʽ,����������֧����
  --   <BRID>����ID</BRID>     //����ID
  --   <SFZH>���֤��</SFZH>     //���֤��
  --   <XM>����</XM>            //����
  --   <BRLX></BRLX>             //ҽ����������
  --   <FB>��ͨ</FB>               //���˷ѱ𣬿��Բ���
  --   <JQM>������</JQM>            //������
  --   <JSMS>1</JSMS>          //����ģʽ��0-��ͨģʽ��1-�첽����ģʽ
  --   <CZLX>0</CZLX>          //�������ͣ�����ģʽΪ1ʱ���룬0-��ʼ���㣬1-��ɽ��㣬2-���˽���
  --   <JZID>1</JZID>          //����ID����������Ϊ1��2ʱ����
  --   <ZFBZH>֧�������ں�UserID</ZFBZH>
  --   <ZFBXCY>֧����С����UserID</ZFBXCY>
  --   <WXGZHID>΢�Ź��ں�OpenID</WXGZH>
  --   <WXXCXID>΢��С����OpenID</WXXCXID>
  --   <JSLIST>          //�����б���������Ϊ2ʱ�ɲ�����
  --     <JS>            //������Ϣ���Һŷ�ҽ������Ŀǰ��֧��һ�����ṹ���շ�һ��
  --       <JSKLB>���㿨���</JSKLB>    //���㿨���
  --       <JSKH>֧�����ʺ�</JSKH>           //���㿨��(֧�����ʺ�)
  --       <JYSM>����˵��</JYSM>            //˵�����̶���֧����
  --       <JYLSH>��ˮ��</JYLSH>           //��ˮ�ţ���������
  --       <JSFS>���㷽ʽ</JSFS>            //���㷽ʽ:�ֽ�֧Ʊ�������������,���Դ���
  --       <JSJE>������</JSJE>            //������
  --       <ZY>ժҪ</ZY>                  //ժҪ
  --       <SFCYJ></SFCYJ>              //�Ƿ��Ԥ�����Һ�Ŀǰ����
  --       <SFXFK></SFXFK>              //�Ƿ����ѿ�,�Һ�Ŀǰ����
  --       <EXPENDLIST>                 //��չ��Ϣ
  --         <EXPEND>
  --           <JYMC>��������</JYMC>        //��������
  --           <JYLR>��������<JYLR>         //��������
  --         </EXPEND>
  --         <EXPEND>
  --           ...
  --         </EXPEND>
  --       </EXPENDLIST>
  --     </JS>
  --   </JSLIST>
  --</IN>

  --����:Xml_Out
  --<OUTPUT>
  --  <GHDH>�Һŵ���</GHDH>          //�Һŵ���
  --  <CZSJ>����ʱ��</CZSJ>          //HIS�ĵǼ�ʱ��
  --  <JZID>����ID</JZID>          //���ν���ID
  --  <KPBZ>��Ʊ��־</KPBZ> //1-�ɹ����ߵ���Ʊ��;0-δ��Ʊ�ɹ���־
  --  <URL>H5ҳ��URL</URL>
  --  <NETURL>����H5ҳ��URL</NETURL>
  --  <FPTT>��Ʊ̧ͷ</FPTT>        //��������
  --  <FPH>��Ʊ��</FPH>             //��Ʊ���
  --  <FPJE>��Ʊ���</FPJE>        //100.00
  --  <KPRQ>��Ʊ����</KPRQ>   //yyyy-mm-dd
  --  <ERROR><MSG>������Ϣ</MSG></ERROR>  //����ʱ����
  --</ OUTPUT>
  --------------------------------------------------------------------------------------------------
  v_����     �ҺŰ���.����%Type;
  n_����     �Һ����״̬.���%Type;
  d_ԭʼʱ�� Date;
  n_Ӧ�ս�� ������ü�¼.Ӧ�ս��%Type;
  v_ԤԼ��ʽ ԤԼ��ʽ.����%Type;
  v_������λ ���˹Һż�¼.������λ%Type;
  n_����id   ������Ϣ.����id%Type;
  v_�������� ������Ϣ.��������%Type;
  v_�ѱ�     ������ü�¼.�ѱ�%Type;
  v_������   �Һ����״̬.������%Type;
  n_�ɿʽ Number(3);
  n_��¼id   �ٴ������¼.Id%Type;
  v_���֤�� ������Ϣ.���֤��%Type;
  v_����     ������ü�¼.����%Type;
  n_����ģʽ Number(1); --0-��ͨģʽ��1-�첽����ģʽ
  n_�������� Number(1); --����ģʽΪ1ʱ���룬0 - ��ʼ���㣬1 - ��ɽ��㣬2 - ���˽���

  v_Para     Varchar2(2000);
  n_�Һ�ģʽ Number(3);
  d_����ʱ�� Date;
  d_����ʱ�� Date;
  d_�Ǽ�ʱ�� Date;

  n_������ʽ   Number;
  v_��ˮ��     ����Ԥ����¼.������ˮ��%Type;
  v_˵��       ������ü�¼.ժҪ%Type;
  v_��������� ҽ�ƿ����.����%Type;
  v_���㿨��   ����Ԥ����¼.����%Type;
  v_No         ���˹Һż�¼.No%Type;
  v_���㷽ʽ   ҽ�ƿ����.���㷽ʽ%Type;

  n_����id   ������ü�¼.����id%Type;
  n_�����id ҽ�ƿ����.Id%Type;
  v_�Ű�     �ҺŰ���.����%Type;
  n_����id   �ҺŰ���.Id%Type;
  n_�ƻ�id   �ҺŰ��żƻ�.Id%Type;
  n_Ԥ��id   ����Ԥ����¼.Id%Type;
  n_��ſ��� �ҺŰ���.��ſ���%Type;
  v_����     �ҺŰ�������.������Ŀ%Type;
  v_�ֽ�     ���㷽ʽ.����%Type;
  n_��ʱ��   Number(3);
  v_�������� Varchar2(3000);
  v_���ս��� Varchar2(1000);
  n_Step     Number(2);

  v_�����     �������׼�¼.���%Type;
  n_��Ԥ��     ����Ԥ����¼.��Ԥ��%Type;
  n_��������id ����Ԥ����¼.��������id%Type;

  v_Temp    Varchar2(32767); --��ʱXML
  x_Templet Xmltype; --ģ��XML

  n_Count     Number(3);
  n_Checkmzlg Number(2);
  v_Err_Msg   Varchar2(200);
  Err_Item    Exception;
  Err_Special Exception;
  n_�Ƿ����Ʊ��       ����Ԥ����¼.�Ƿ����Ʊ��%Type;
  v_֧�������ں�userid Varchar2(100);
  v_֧����С����userid Varchar2(100);
  v_΢�Ź��ں�openid   Varchar2(100);
  v_΢��С����openid   Varchar2(100);
  n_��Ʊ��־           Number(2);
  v_��������           ����Ʊ��ʹ�ü�¼.����%Type;
  v_��Ʊ���           ����Ʊ��ʹ�ü�¼.����%Type;
  v_��Ʊ����           Varchar2(20);
  n_��Ʊ���           ����Ʊ��ʹ�ü�¼.Ʊ�ݽ��%Type;
  v_Url                ����Ʊ��ʹ�ü�¼.Url����%Type;
  v_Url����            ����Ʊ��ʹ�ü�¼.Url����%Type;

Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/HM'), Extractvalue(Value(A), 'IN/HX'),
         To_Date(Extractvalue(Value(A), 'IN/YYSJ'), 'YYYY-MM-DD hh24:mi:ss'), Extractvalue(Value(A), 'IN/JE'),
         Extractvalue(Value(A), 'IN/YYFS'), Extractvalue(Value(A), 'IN/HZDW'), Extractvalue(Value(A), 'IN/BRID'),
         Extractvalue(Value(A), 'IN/BRLX'), Extractvalue(Value(A), 'IN/FB'), Extractvalue(Value(A), 'IN/JQM'),
         Extractvalue(Value(A), 'IN/JKFS'), Extractvalue(Value(A), 'IN/CZJLID'), Extractvalue(Value(A), 'IN/SFZH'),
         Extractvalue(Value(A), 'IN/XM'), Extractvalue(Value(A), 'IN/JSMS'), Extractvalue(Value(A), 'IN/CZLX'),
         Extractvalue(Value(A), 'IN/JZID'), Extractvalue(Value(A), 'IN/ZFBZH'), Extractvalue(Value(A), 'IN/ZFBXCY'),
         Extractvalue(Value(A), 'IN/WXGZHID'), Extractvalue(Value(A), 'IN/WXXCXID')
  Into v_����, n_����, d_ԭʼʱ��, n_Ӧ�ս��, v_ԤԼ��ʽ, v_������λ, n_����id, v_��������, v_�ѱ�, v_������, n_�ɿʽ, n_��¼id, v_���֤��, v_����, n_����ģʽ,
       n_��������, n_����id, v_֧�������ں�userid, v_֧����С����userid, v_΢�Ź��ں�openid, v_΢��С����openid
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If Nvl(n_����id, 0) = 0 And Not v_���֤�� Is Null And Not v_���� Is Null Then
    n_����id := Zl_Third_Getpatiid(v_���֤��, v_����);
  End If;
  If Nvl(n_����id, 0) = 0 Then
    v_Err_Msg := '�޷�ȷ��������Ϣ,����!';
    Raise Err_Item;
  End If;

  If Not v_֧�������ں�userid Is Null Then
    Zl_������Ϣ�ӱ�_Update(n_����id, Upper('֧�������ں�UserID'), v_֧�������ں�userid);
  End If;

  If Not v_֧����С����userid Is Null Then
    Zl_������Ϣ�ӱ�_Update(n_����id, Upper('֧����С����UserID'), v_֧�������ں�userid);
  End If;

  If Not v_΢�Ź��ں�openid Is Null Then
    Zl_������Ϣ�ӱ�_Update(n_����id, Upper('΢�Ź��ں�OpenID'), v_֧�������ں�userid);
  End If;

  If Not v_΢��С����openid Is Null Then
    Zl_������Ϣ�ӱ�_Update(n_����id, Upper('΢��С����OpenID'), v_֧�������ں�userid);
  End If;

  If Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) <> 0 Then
    If Nvl(n_����id, 0) = 0 Then
      v_Err_Msg := 'û��ָ����صĽ������ݣ�';
      Raise Err_Item;
    End If;
  
    Begin
      Select NO, ����ʱ��, �Ǽ�ʱ��
      Into v_No, d_����ʱ��, d_�Ǽ�ʱ��
      From ������ü�¼
      Where ��¼���� = 4 And Nvl(����״̬, 0) = 1 And ����id = n_����id And Rownum < 2;
    Exception
      When Others Then
        v_Err_Msg := 'û���ҵ�ָ����������ݣ������ѱ�����';
        Raise Err_Item;
    End;
  End If;

  If Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) = 2 Then
    Zl_���˹Һż�¼_Cancel(n_����id);
  
    v_Temp := '<GHDH>' || v_No || '</GHDH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<CZSJ>' || To_Char(d_�Ǽ�ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<JZID>' || n_����id || '</JZID>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    Xml_Out := x_Templet;
    Return;
  End If;

  v_Para     := zl_GetSysParameter(256);
  n_�Һ�ģʽ := Substr(v_Para, 1, 1);
  Begin
    d_����ʱ�� := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
  Exception
    When Others Then
      d_����ʱ�� := Null;
  End;

  If Nvl(n_����ģʽ, 0) = 0 Or Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) = 0 Then
    If Sysdate - 10 > Nvl(d_����ʱ��, Sysdate - 30) Then
      If n_�Һ�ģʽ = 1 And Nvl(d_ԭʼʱ��, Sysdate) > Nvl(d_����ʱ��, Sysdate - 30) And n_��¼id Is Null Then
        v_Err_Msg := 'ϵͳ��ǰ���ڳ�����Ű�ģʽ������Ĳ����޷�ȷ���ҺŰ��ţ������ԣ�';
        Raise Err_Item;
      End If;
    Else
      If n_�Һ�ģʽ = 1 And Nvl(d_ԭʼʱ��, Sysdate) > Nvl(d_����ʱ��, Sysdate - 30) And n_��¼id Is Null Then
        Begin
          Select a.Id
          Into n_��¼id
          From �ٴ������¼ A, �ٴ������Դ B
          Where a.��Դid = b.Id And b.���� = v_���� And Nvl(d_ԭʼʱ��, Sysdate) Between a.��ʼʱ�� And a.��ֹʱ��;
        Exception
          When Others Then
            v_Err_Msg := 'ϵͳ��ǰ���ڳ�����Ű�ģʽ������Ĳ����޷�ȷ���ҺŰ��ţ������ԣ�';
            Raise Err_Item;
        End;
      End If;
    End If;
  
    n_Checkmzlg := To_Number(Nvl(zl_GetSysParameter(323), '0'));
    For c_���׼�¼ In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As ���㿨���,
                          Extractvalue(b.Column_Value, '/JS/JSFS') As ���㷽ʽ,
                          Extractvalue(b.Column_Value, '/JS/SFCYJ') As �Ƿ��Ԥ��,
                          Extractvalue(b.Column_Value, '/JS/JYLSH') As ������ˮ��,
                          Extractvalue(b.Column_Value, '/JS/JSKH') As ���㿨��, Extractvalue(b.Column_Value, '/JS/ZY') As ժҪ
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
    
      --��Ԥ������Ҫ����������
      If Nvl(c_���׼�¼.�Ƿ��Ԥ��, 0) = 0 Then
        If c_���׼�¼.���㿨��� Is Null Then
          v_����� := c_���׼�¼.���㷽ʽ;
        Else
          Select Decode(Translate(Nvl(c_���׼�¼.���㿨���, 'abcd'), '#1234567890', '#'), Null, 1, 0)
          Into n_Count
          From Dual;
          If Nvl(n_Count, 0) = 1 Then
            Select Max(����) Into v_����� From ҽ�ƿ���� Where ID = To_Number(c_���׼�¼.���㿨���);
          Else
            Select Max(����) Into v_����� From ҽ�ƿ���� Where ���� = c_���׼�¼.���㿨���;
          End If;
        End If;
      
        If v_����� Is Null Then
          v_Err_Msg := '��֧�ֵĽ��㷽ʽ,���飡';
          Raise Err_Item;
        End If;
      
        --����һ�����㷽ʽ�ż�齻����
        n_Step := Nvl(n_Step, 0) + 1;
        If Zl_Fun_�������׼�¼_Locked(v_�����, c_���׼�¼.������ˮ��, c_���׼�¼.���㿨��, c_���׼�¼.ժҪ, 4) = 0 And n_Step = 1 Then
          v_Err_Msg := '������ˮ��Ϊ:' || c_���׼�¼.������ˮ�� || '�Ľ������ڽ����У��������ٴ��ύ�˽���!';
          Raise Err_Special;
        End If;
      Else
        If Nvl(n_Checkmzlg, 0) <> 0 Then
          Select Count(1)
          Into n_Count
          From ������ҳ A, ������Ϣ B
          Where a.����id = n_����id And a.�������� = 1 And a.����id = b.����id And a.��ҳid = b.��ҳid And Nvl(b.��Ժ, 0) = 1;
          If n_Count <> 0 Then
            v_Err_Msg := '�������۲��˲���ʹ������Ԥ����';
            Raise Err_Item;
          End If;
        End If;
      End If;
    End Loop;
  
    If v_�������� Is Not Null Then
      Select Count(1) Into n_Count From �������� Where ���� = v_��������;
      If n_Count = 0 Then
        v_Err_Msg := 'û�з���Ϊ(' || v_�������� || ')�Ĳ������ͣ�';
        Raise Err_Item;
      End If;
      Update ������Ϣ Set �������� = Nvl(��������, v_��������) Where ����id = n_����id;
    End If;
  
    d_�Ǽ�ʱ�� := Sysdate;
    v_No       := Nextno(12);
    Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
  
    If n_��¼id Is Null Then
      Select C2
      Into v_����
      From Table(f_Str2List2('1:����,2:��һ,3:�ܶ�,4:����,5:����,6:����,7:����'))
      Where C1 = To_Char(d_ԭʼʱ��, 'D');
    
      Begin
        Select ID
        Into n_�ƻ�id
        From (Select ID
               From �ҺŰ��żƻ�
               Where ���� = v_���� And d_ԭʼʱ�� Between Nvl(��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And ʧЧʱ�� And
                     ���ʱ�� Is Not Null
               Order By ��Чʱ�� Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          Select ID Into n_����id From �ҺŰ��� Where ���� = v_����;
      End;
    
      d_����ʱ�� := d_ԭʼʱ��;
      If Nvl(n_�ƻ�id, 0) <> 0 Then
        --�Ӽƻ���ȡ��Ϣ
        Select Decode(To_Char(d_ԭʼʱ��, 'D'), '1', a.����, '2', a.��һ, '3', a.�ܶ�, '4', a.����, '5', a.����, '6', a.����, '7', a.����,
                       Null), Nvl(a.��ſ���, 0)
        Into v_�Ű�, n_��ſ���
        From �ҺŰ��żƻ� A
        Where a.Id = n_�ƻ�id;
        Select Count(1) Into n_��ʱ�� From �Һżƻ�ʱ�� Where ���� = v_���� And �ƻ�id = n_�ƻ�id And Rownum < 2;
      
        --������λ���
        If v_������λ Is Not Null Then
          Select Count(1)
          Into n_Count
          From ������λ�ƻ�����
          Where �ƻ�id = n_�ƻ�id And ���� = 0 And ������λ = v_������λ;
          If n_Count = 1 Then
            v_Err_Msg := '����ĺ�����λ�ڴ˺����ϱ����ã�';
            Raise Err_Item;
          End If;
        End If;
      
        If n_��ʱ�� = 1 And n_��ſ��� = 0 Then
          Select ���
          Into n_����
          From �Һżƻ�ʱ��
          Where �ƻ�id = n_�ƻ�id And ���� = v_���� And To_Char(��ʼʱ��, 'hh24:mi:ss') = To_Char(d_����ʱ��, 'hh24:mi:ss');
        Else
          Begin
            Select To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || '' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
            Into d_����ʱ��
            From �Һżƻ�ʱ��
            Where �ƻ�id = n_�ƻ�id And ���� = v_���� And ��� = Nvl(n_����, 0);
          Exception
            When Others Then
              If n_��ʱ�� = 1 And n_��ſ��� = 1 Then
                Select To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || '' || To_Char(Max(����ʱ��), 'hh24:mi:ss'),
                                'YYYY-MM-DD hh24:mi:ss')
                Into d_����ʱ��
                From �Һżƻ�ʱ��
                Where �ƻ�id = n_�ƻ�id And ���� = v_����;
              Else
                Select To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || '' || To_Char(��ʼʱ��, 'hh24:mi:ss'),
                                'YYYY-MM-DD hh24:mi:ss')
                Into d_����ʱ��
                From ʱ���
                Where ʱ��� = v_�Ű�;
              End If;
              If d_����ʱ�� < d_�Ǽ�ʱ�� Then
                d_����ʱ�� := d_�Ǽ�ʱ��;
              End If;
          End;
        End If;
      Else
        --�Ӱ��Ŷ�ȡ��Ϣ
        Select Decode(To_Char(d_ԭʼʱ��, 'D'), '1', b.����, '2', b.��һ, '3', b.�ܶ�, '4', b.����, '5', b.����, '6', b.����, '7', b.����,
                       Null), Nvl(b.��ſ���, 0)
        Into v_�Ű�, n_��ſ���
        From �ҺŰ��� B
        Where b.Id = n_����id;
        Select Count(1) Into n_��ʱ�� From �ҺŰ���ʱ�� Where ���� = v_���� And ����id = n_����id And Rownum < 2;
      
        --������λ���
        If v_������λ Is Not Null Then
          Select Count(1)
          Into n_Count
          From ������λ���ſ���
          Where ����id = n_����id And ���� = 0 And ������λ = v_������λ;
          If n_Count = 1 Then
            v_Err_Msg := '����ĺ�����λ�ڴ˺����ϱ����ã�';
            Raise Err_Item;
          End If;
        End If;
      
        If n_��ʱ�� = 1 And n_��ſ��� = 0 Then
          d_����ʱ�� := d_ԭʼʱ��;
          Select ���
          Into n_����
          From �ҺŰ���ʱ��
          Where ����id = n_����id And ���� = v_���� And To_Char(��ʼʱ��, 'hh24:mi:ss') = To_Char(d_����ʱ��, 'hh24:mi:ss');
        Else
          Begin
            Select To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || '' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
            Into d_����ʱ��
            From �ҺŰ���ʱ��
            Where ����id = n_����id And ���� = v_���� And ��� = Nvl(n_����, 0);
          Exception
            When Others Then
              If n_��ʱ�� = 1 And n_��ſ��� = 1 Then
                Select To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || '' || To_Char(Max(����ʱ��), 'hh24:mi:ss'),
                                'YYYY-MM-DD hh24:mi:ss')
                Into d_����ʱ��
                From �ҺŰ���ʱ��
                Where ����id = n_����id And ���� = v_����;
              Else
                Select To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || '' || To_Char(��ʼʱ��, 'hh24:mi:ss'),
                                'YYYY-MM-DD hh24:mi:ss')
                Into d_����ʱ��
                From ʱ���
                Where ʱ��� = v_�Ű�;
              End If;
              If d_����ʱ�� < d_�Ǽ�ʱ�� Then
                d_����ʱ�� := d_�Ǽ�ʱ��;
              End If;
          End;
        End If;
      End If;
    Else
      --������Ű�ģʽ
      Begin
        Select ��ʼʱ�� Into d_����ʱ�� From �ٴ�������ſ��� Where ��¼id = n_��¼id And ��� = n_����;
      Exception
        When Others Then
          d_����ʱ�� := d_ԭʼʱ��;
      End;
    End If;
  
    --�Ȳ������˹Һż�¼�Ͳ��˷��ü�¼
    If Nvl(n_�ɿʽ, 0) = 0 Then
      If Trunc(d_����ʱ��) <> Trunc(Sysdate) Then
        n_������ʽ := 3;
      Else
        n_������ʽ := 1;
      End If;
    Else
      n_������ʽ := 2;
    End If;
    Zl_���������Һ�_Insert(n_������ʽ, n_����id, v_����, n_����, v_No, Null, Null, Null, d_����ʱ��, d_�Ǽ�ʱ��, v_������λ, n_Ӧ�ս��, Null, Null,
                     Null, Null, v_ԤԼ��ʽ, Null, Null, Null, 1, n_����id, 0, Null, Null, Null, 1, v_�ѱ�, Null, v_������, 1, 0,
                     n_��¼id, 0, Null, 1, 0);
  End If;

  Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
  For r_���� In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As ���㿨���,
                      Extractvalue(b.Column_Value, '/JS/JSKH') As ���㿨��,
                      Extractvalue(b.Column_Value, '/JS/SFCYJ') As �Ƿ��Ԥ��,
                      Extractvalue(b.Column_Value, '/JS/JSFS') As ���㷽ʽ,
                      Extractvalue(b.Column_Value, '/JS/JYLSH') As ������ˮ��,
                      Extractvalue(b.Column_Value, '/JS/JYSM') As ����˵��, Extractvalue(b.Column_Value, '/JS/JSJE') As ������
               From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
  
    If Nvl(r_����.�Ƿ��Ԥ��, 0) = 0 Then
      If r_����.���㷽ʽ Is Null Then
        Begin
          Select b.���㷽ʽ, b.Id
          Into v_���㷽ʽ, n_�����id
          From ҽ�ƿ���� B
          Where b.���� = r_����.���㿨��� And Rownum < 2;
        Exception
          When Others Then
            v_Err_Msg := 'û�з��ָý��㿨�������Ϣ';
            Raise Err_Item;
        End;
        v_�������� := v_�������� || '|' || v_���㷽ʽ || ',' || r_����.������ || ',,';
      Else
        v_���㷽ʽ := r_����.���㷽ʽ;
        Select Count(1) Into n_Count From ���㷽ʽ Where ���� = v_���㷽ʽ And ���� In (3, 4);
        If n_Count = 1 And r_����.���㿨��� Is Null Then
          v_���ս��� := v_���ս��� || '||' || v_���㷽ʽ || '|' || r_����.������;
        Else
          v_�������� := v_�������� || '|' || v_���㷽ʽ || ',' || r_����.������ || ',,';
        End If;
      End If;
    
      If r_����.���㿨��� Is Not Null Then
        v_��������   := v_�������� || '1,';
        v_��������� := r_����.���㿨���;
        v_���㿨��   := r_����.���㿨��;
        v_��ˮ��     := r_����.������ˮ��;
        v_˵��       := r_����.����˵��;
        If n_�����id Is Null Then
          Begin
            Select b.Id Into n_�����id From ҽ�ƿ���� B Where b.���� = r_����.���㿨��� And Rownum < 2;
          Exception
            When Others Then
              v_Err_Msg := 'û�з��ָý��㿨�������Ϣ';
              Raise Err_Item;
          End;
        End If;
      
        Select Decode(Translate(Nvl(r_����.���㿨���, 'abcd'), '#1234567890', '#'), Null, 1, 0)
        Into n_Count
        From Dual;
        If Nvl(n_Count, 0) = 1 Then
          Select Max(����) Into v_����� From ҽ�ƿ���� Where ID = To_Number(r_����.���㿨���);
        Else
          Select Max(����) Into v_����� From ҽ�ƿ���� Where ���� = r_����.���㿨���;
        End If;
      
        If Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) = 0 Then
          Select Max(��������id)
          Into n_��������id
          From ����Ԥ����¼
          Where ��¼���� Not In (1, 11) And ����id = n_����id And �����id = n_�����id And Rownum < 2;
          If Nvl(n_��������id, 0) = 0 Then
            n_��������id := n_Ԥ��id;
          Else
            Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          End If;
        
          Insert Into ����Ԥ����¼
            (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��,
             ������λ, ��������, �������, ����ʱ��, ������Ա, ��������id, У�Ա�־)
            Select n_Ԥ��id, 4, 1, v_No, ����id, v_���㷽ʽ, r_����.������, �Ǽ�ʱ��, ����Ա���, ����Ա����, ����id, Null, �ɿ���id, n_�����id, Null,
                   v_���㿨��, v_��ˮ��, v_˵��, v_������λ, 4, '�����ӿڹҺ�', �Ǽ�ʱ��, ����Ա����, n_��������id, 1
            From ������ü�¼
            Where ��¼���� = 4 And ����id = n_����id And Rownum < 2;
        Else
          If Nvl(n_����ģʽ, 0) = 1 Then
            Delete From ����Ԥ����¼ Where ��¼���� Not In (1, 11) And ����id = n_����id And n_�����id = n_�����id;
          End If;
          v_�������� := v_�������� || n_Ԥ��id;
        End If;
      Else
        v_�������� := v_�������� || '0,';
        v_�����   := r_����.���㷽ʽ;
      End If;
    
      If Nvl(n_����ģʽ, 0) = 0 Or Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) = 1 Then
        Update �������׼�¼
        Set ҵ�����id = n_����id
        Where ��ˮ�� = r_����.������ˮ�� And ��� = v_����� And ҵ������ = 4;
      End If;
    Else
      n_��Ԥ�� := r_����.������;
    End If;
  End Loop;

  If Nvl(n_����ģʽ, 0) = 0 Or Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) = 1 Then
    If v_�������� Is Not Null Then
      v_�������� := Substr(v_��������, 2);
    Else
      Begin
        Select ���� Into v_�ֽ� From ���㷽ʽ Where ���� = 1;
      Exception
        When Others Then
          v_�ֽ� := '�ֽ�';
      End;
      v_�������� := v_�ֽ� || ',' || 0 || ',,0';
    End If;
  
    If Nvl(n_�ɿʽ, 0) = 0 Then
      If Trunc(d_����ʱ��) <> Trunc(Sysdate) Then
        n_������ʽ := 3;
      Else
        n_������ʽ := 1;
      End If;
    Else
      n_������ʽ := 2;
    End If;
    Zl_���������Һ�_Insert(n_������ʽ, n_����id, v_����, n_����, v_No, Null, v_��������, Null, d_����ʱ��, d_�Ǽ�ʱ��, v_������λ, n_Ӧ�ս��, Null, Null,
                     v_��ˮ��, v_˵��, v_ԤԼ��ʽ, n_Ԥ��id, n_�����id, Null, 1, n_����id, 0, v_���ս���, n_��Ԥ��, v_���㿨��, 1, v_�ѱ�, Null,
                     v_������, 1, 0, n_��¼id, 0, Null, 1, 1);
  
    If Nvl(n_�����id, 0) <> 0 Then
      For c_��չ��Ϣ In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As ��������,
                            Extractvalue(b.Column_Value, '/EXPEND/JYLR') As ��������
                     From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS/EXPENDLIST/EXPEND'))) B) Loop
        Zl_�������㽻��_Insert(n_�����id, 0, v_���㿨��, n_����id, c_��չ��Ϣ.�������� || '|' || c_��չ��Ϣ.��������, 0);
      End Loop;
    
    End If;
  
    --����Ʊ�ݴ���  
    n_�Ƿ����Ʊ�� := b_Einvoice_Request.Einvoice_Start(4, Null);
    Update ����Ԥ����¼ Set �Ƿ����Ʊ�� = n_�Ƿ����Ʊ�� Where ����id = n_����id;
  
    If Nvl(n_�Ƿ����Ʊ��, 0) = 1 Then
      --��Ҫ���ߵ���Ʊ��
      If b_Einvoice_Request.Einvoice_Create(4, n_����id, Null, v_Err_Msg) = 0 Then
        --����Ʊ�ݿ��߳ɹ�
        Raise Err_Item;
      End If;
    
      Select Max(1), Max(����), Max(����), Max(To_Char(����ʱ��, 'yyyy-mm-dd')), Max(Url����), Max(Url����), Max(Ʊ�ݽ��)
      Into n_��Ʊ��־, v_��������, v_��Ʊ���, v_��Ʊ����, v_Url, v_Url����, n_��Ʊ���
      From ����Ʊ��ʹ�ü�¼
      Where ����id = n_����id And Ʊ�� = 4 And ��¼״̬ = 1;
    
      If v_�������� Is Not Null Then
        v_���� := v_��������;
      End If;
    End If;
  End If;
  v_Temp := '<GHDH>' || v_No || '</GHDH>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<CZSJ>' || To_Char(d_�Ǽ�ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<JZID>' || n_����id || '</JZID>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<KPBZ>' || Nvl(n_��Ʊ��־, 0) || '</KPBZ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;

  v_Temp := '<URL>' || Nvl(v_Url, '') || '</URL>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;

  v_Temp := '<NETURL>' || Nvl(v_Url����, '') || '</NETURL>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;

  v_Temp := '<FPTT>' || v_���� || '</FPTT>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;

  v_Temp := '<FPH>' || v_��Ʊ��� || '</FPH>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<FPJE>' || Nvl(n_��Ʊ���, 0) || '</FPJE>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<KPRQ>' || v_��Ʊ���� || '</KPRQ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Xml_Out := x_Templet;

Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Err_Special Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20105, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Regist;
/


Create Or Replace Procedure Zl_Third_Registdelcheck
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --����:HIS�˺ż��
  --���:Xml_In:
  --<IN>
  --  <GHDH>A000001</GHDH>    //�Һŵ���
  --  <JSKLB>֧����</JSKLB>      //���㿨���
  --  <JCFP>1</JCFP>            //��鷢Ʊ
  --  <GHJE>20</GHJE>            //�ҺŽ��
  --  <LSH>34563</LSH>           //������ˮ��
  --  <JKFS>0</JKFS>             //�ɿʽ,0-�ҺŻ�ԤԼ�ɿ�;1-ԤԼ���ɿ�
  --  <YYFS></YYFS>              //�ɿʽ=1ʱ���룬ԤԼ��ԤԼ��ʽ
  --  <XL></XL>                  //����
  --</IN>

  --����:Xml_Out
  --<OUTPUT>
  -- <ERROR><MSG></MSG></ERROR> //Ϊ�ձ�ʾ���ɹ�
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  v_�����     Varchar2(100);
  v_No         ���˹Һż�¼.No%Type;
  n_�ҺŽ��   ������ü�¼.ʵ�ս��%Type;
  v_���㷽ʽ   ҽ�ƿ����.���㷽ʽ%Type;
  n_ʵ�ս��   ������ü�¼.ʵ�ս��%Type;
  v_������ˮ�� ����Ԥ����¼.������ˮ��%Type;
  n_����       Number(3);
  v_Type       Varchar2(50);
  v_Temp       Varchar2(32767); --��ʱXML
  x_Templet    Xmltype; --ģ��XML

  n_�ѿ�ҽ�� Number(2);
  n_��鷢Ʊ Number(3);
  n_�Ƿ��ӡ Number(3);
  n_�ɿʽ Number(3);
  n_����     ������Ϣ.����%Type;
  v_ԤԼ��ʽ ���˹Һż�¼.ԤԼ��ʽ%Type;
  v_�շѵ�   ������ü�¼.No%Type;
  n_����id   ������ü�¼.����id%Type;
  v_Err_Msg  Varchar2(200);
  Err_Item Exception;
  n_Count Number;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/GHDH'), Extractvalue(Value(A), 'IN/JSKLB'), Extractvalue(Value(A), 'IN/GHJE'),
         Extractvalue(Value(A), 'IN/LSH'), To_Number(Extractvalue(Value(A), 'IN/JCFP')),
         To_Number(Extractvalue(Value(A), 'IN/JKFS')), Extractvalue(Value(A), 'IN/YYFS'),
         To_Number(Extractvalue(Value(A), 'IN/XL'))
  Into v_No, v_�����, n_�ҺŽ��, v_������ˮ��, n_��鷢Ʊ, n_�ɿʽ, v_ԤԼ��ʽ, n_����
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  Select Max(�շѵ�) Into v_�շѵ� From ���˹Һż�¼ Where NO = v_No;

  n_�ɿʽ := Nvl(n_�ɿʽ, 0);
  If v_����� Is Not Null And n_�ɿʽ = 0 Then
    Select Nvl2(Translate(v_�����, '\1234567890', '\'), 'Char', 'Num') Into v_Type From Dual;
    If v_Type = 'Num' Then
      --������ǿ����ID
      Select ���㷽ʽ Into v_���㷽ʽ From ҽ�ƿ���� Where ID = To_Number(v_�����);
    Else
      --������ǿ��������
      Select ���㷽ʽ Into v_���㷽ʽ From ҽ�ƿ���� Where ���� = v_�����;
    End If;
    If Nvl(n_�ɿʽ, 0) = 0 Then
      If Nvl(n_����, 0) = 0 Then
        Select Nvl(Max(1), 0)
        Into n_����
        From ����Ԥ����¼ A,
             (Select Distinct ����id
               From ������ü�¼
               Where NO = v_No And ��¼���� = 4
               Union
               Select Distinct ����id
               From סԺ���ü�¼
               Where NO = v_No And ��¼���� = 5
               Union
               Select Distinct ����id
               From ������ü�¼
               Where NO = v_�շѵ� And ��¼���� = 1) B
        Where a.����id = b.����id And ���㷽ʽ <> v_���㷽ʽ And Mod(��¼����, 10) <> 1 And Rownum < 2;
      Else
        Select Nvl(Max(1), 0)
        Into n_����
        From ����Ԥ����¼ A,
             (Select Distinct ����id
               From ������ü�¼
               Where NO = v_No And ��¼���� = 4
               Union
               Select Distinct ����id
               From סԺ���ü�¼
               Where NO = v_No And ��¼���� = 5
               Union
               Select Distinct ����id
               From ������ü�¼
               Where NO = v_�շѵ� And ��¼���� = 1) B, ���㷽ʽ C
        Where a.����id = b.����id And ���㷽ʽ <> v_���㷽ʽ And Mod(��¼����, 10) <> 1 And a.���㷽ʽ = c.���� And c.���� Not In (3, 4) And
              Rownum < 2;
        If n_���� = 0 Then
          Select Nvl(Max(1), 0)
          Into n_����
          From ���ս����¼ A,
               (Select Distinct ����id
                 From ������ü�¼
                 Where NO = v_No And ��¼���� = 4
                 Union
                 Select Distinct ����id
                 From סԺ���ü�¼
                 Where NO = v_No And ��¼���� = 5
                 Union
                 Select Distinct ����id
                 From ������ü�¼
                 Where NO = v_�շѵ� And ��¼���� = 1) B
          Where a.��¼id = b.����id And ���� <> n_���� And Rownum < 2;
        End If;
      End If;
      If n_���� = 1 Then
        v_Err_Msg := '����ĹҺŵ��ݰ���' || v_���㷽ʽ || '����Ľ��㷽ʽ,�޷��˺�!';
        Raise Err_Item;
      End If;
    Else
      Begin
        Select 1 Into n_���� From ���˹Һż�¼ A Where a.No = v_No And a.ԤԼ��ʽ = v_ԤԼ��ʽ And Rownum < 2;
      Exception
        When Others Then
          n_���� := 0;
      End;
      If n_���� = 0 Then
        v_Err_Msg := '����ĹҺŵ��ݲ���' || v_ԤԼ��ʽ || 'ԤԼ��,�޷��˺�!';
        Raise Err_Item;
      End If;
    End If;
  End If;

  If n_�ɿʽ = 0 Then
    Select Sum(ʵ�ս��), Max(Decode(��¼״̬, 2, 0, ����id))
    Into n_ʵ�ս��, n_����id
    From ������ü�¼
    Where NO = v_No And ��¼���� = 4;
    If Not v_�շѵ� Is Null Then
      Select Sum(ʵ�ս��) Into n_ʵ�ս�� From ������ü�¼ Where NO = v_�շѵ� And ��¼���� = 1;
    End If;
    If n_ʵ�ս�� <> n_�ҺŽ�� Then
      v_Err_Msg := '������˿�����ʵ�ʹҺŽ���������!';
      Raise Err_Item;
    End If;
    --����Ʊ�ݼ��
    If b_Einvoice_Request.Einvoice_Cancel_Check(4, n_����id, v_Err_Msg) = 0 Then
      --ʧ�ܺ�ֱ���״�
      Raise Err_Item;
    End If;
  End If;

  --��������飬�Ѵ��ڲ��������ݵģ������˺�
  Begin
    Select 1
    Into n_����
    From ���ò����¼ A,
         (Select Distinct ����id
           From ������ü�¼
           Where NO = v_No And ��¼���� = 4
           Union
           Select Distinct ����id
           From סԺ���ü�¼
           Where NO = v_No And ��¼���� = 5
           Union
           Select Distinct ����id
           From ������ü�¼
           Where NO = v_�շѵ� And ��¼���� = 1) B
    Where a.�շѽ���id = b.����id And a.��¼���� = 1 And a.���ӱ�־ = 1 And Nvl(a.����״̬, 0) <> 2 And Rownum < 2;
  Exception
    When Others Then
      n_���� := 0;
  End;
  If n_���� = 1 Then
    v_Err_Msg := '����ĹҺŵ����Ѿ������˶��ν���,�޷��˺�!';
    Raise Err_Item;
  End If;
  --ҽ����飬�Ѿ�����ҽ���ģ������˺�
  Begin
    Select Distinct 1 Into n_�ѿ�ҽ�� From ����ҽ����¼ Where �Һŵ� = v_No;
  Exception
    When Others Then
      n_�ѿ�ҽ�� := 0;
  End;
  If n_�ѿ�ҽ�� = 1 Then
    v_Err_Msg := '����ĹҺŵ����Ѿ�����ҽ��,�޷��˺�!';
    Raise Err_Item;
  End If;
  If Nvl(n_��鷢Ʊ, 0) = 1 Then
    Select Max(Decode(a.ʵ��Ʊ��, Null, 0, 1)) Into n_�Ƿ��ӡ From ������ü�¼ A Where NO = v_No And ��¼���� = 4;
    If Nvl(n_�Ƿ��ӡ, 0) = 1 Then
      v_Err_Msg := '�����˺ŵĵ����ѿ���Ʊ,�����˷�!';
      Raise Err_Item;
    End If;
    Select Max(Decode(a.ʵ��Ʊ��, Null, 0, 1))
    Into n_�Ƿ��ӡ
    From ������ü�¼ A
    Where NO = v_�շѵ� And ��¼���� = 1;
    If Nvl(n_�Ƿ��ӡ, 0) = 1 Then
      v_Err_Msg := '�����˺ŵĵ����ѿ���Ʊ,�����˷�!';
      Raise Err_Item;
    End If;
  End If;

  Select Count(1) Into n_Count From ���˹Һż�¼ Where NO = v_No And ��¼��־ = -1;
  If n_Count <> 0 Then
    v_Err_Msg := '�����˺ŵĵ��ݴ��ڽ����쳣״̬,�������˷�!';
    Raise Err_Item;
  End If;

  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Registdelcheck;
/

Create Or Replace Procedure Zl_Third_Registdel
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  -------------------------------------------------------------------------------------------------- 
  --����:HIS�˺� 
  --���:Xml_In: 
  --<IN> 
  --  <GHDH>A000001</GHDH>    //�Һŵ��� 
  --  <JSKLB>֧����</JSKLB>      //���㿨��� 
  --  <JCFP>1</JCFP>            //��鷢Ʊ 
  --  <GHJE>20</GHJE>            //�ҺŽ�� 
  --  <JSMS>1</JSMS>          //����ģʽ��0-��ͨģʽ��1-�첽����ģʽ
  --  <CZLX>0</CZLX>          //�������ͣ�����ģʽΪ1ʱ���룬0-��ʼ���㣬1-��ɽ��㣬2-���˽���
  --  <CXID>1</CXID>          //��������ID����������Ϊ1��2ʱ���� 
  --  <JKFS>0</JKFS>             //�ɿʽ,0-�ҺŻ�ԤԼ�ɿ�;1-ԤԼ���ɿ� 
  --  <YYFS></YYFS>              //�ɿʽ=1ʱ���룬ԤԼ��ԤԼ��ʽ 
  --  <LSH>34563</LSH>           //������ˮ��
  --</IN> 

  --����:Xml_Out 
  --<OUTPUT> 
  --  <CZSJ>����ʱ��</CZSJ>          //HIS�ĵǼ�ʱ�� 
  --  <YJZID>ԭ����ID</YJZID> 
  --  <CXID>����ID</CXID> 
  --  <KPBZ>��Ʊ��־</KPBZ> //�����˲���Ч:1-�ɹ����ߵ���Ʊ��;0-δ��Ʊ�ɹ���־
  --  <URL>H5ҳ��URL</URL>
  --  <NETURL>����H5ҳ��URL</NETURL>
  --  <FPTT>��Ʊ̧ͷ</FPTT>        //��������
  --  <FPH>��Ʊ��</FPH>             //��Ʊ���
  --  <FPJE>��Ʊ���</FPJE>        //100.00
  --  <KPRQ>��Ʊ����</KPRQ>   //yyyy-mm-dd
  --  <ERROR><MSG></MSG></ERROR> //Ϊ�ձ�ʾȡ���Һųɹ� 
  --</OUTPUT> 
  -------------------------------------------------------------------------------------------------- 
  v_No         ���˹Һż�¼.No%Type;
  v_�����     Varchar2(100);
  n_�ҺŽ��   ������ü�¼.ʵ�ս��%Type;
  v_������ˮ�� ����Ԥ����¼.������ˮ��%Type;
  n_��鷢Ʊ   Number(3);
  n_�ɿʽ   Number(3);
  v_ԤԼ��ʽ   ���˹Һż�¼.ԤԼ��ʽ%Type;
  n_����ģʽ   Number(1); --0-��ͨģʽ��1-�첽����ģʽ
  n_��������   Number(1); --����ģʽΪ1ʱ���룬0 - ��ʼ���㣬1 - ��ɽ��㣬2 - ���˽���

  v_����Ա���   ������ü�¼.����Ա���%Type;
  v_����Ա����   ������ü�¼.����Ա����%Type;
  v_���㷽ʽ     ҽ�ƿ����.���㷽ʽ%Type;
  v_Type         Varchar2(50);
  n_�ѿ�ҽ��     Number(2);
  n_�Ƿ��ӡ     Number(3);
  n_����id       ������ü�¼.����id%Type;
  n_����id       ������ü�¼.����id%Type;
  n_�Һ�ԭ����id ������ü�¼.����id%Type;
  n_�Һų���id   ������ü�¼.����id%Type;
  n_ʣ����     ������ü�¼.����id%Type;
  d_�Ǽ�ʱ��     Date;
  v_�շѵ�       ������ü�¼.No%Type;
  n_����id       ������ü�¼.����id%Type;
  n_�����id     ҽ�ƿ����.Id%Type;
  v_�˷ѽ���     Varchar2(1000);
  n_Temp         Number(18);
  n_Ԥ��id       ����Ԥ����¼.Id%Type;
  n_�Ƿ����Ʊ�� ����Ԥ����¼.�Ƿ����Ʊ��%Type;

  n_��Ʊ��־ Number(2);
  v_�������� ����Ʊ��ʹ�ü�¼.����%Type;
  v_��Ʊ��� ����Ʊ��ʹ�ü�¼.����%Type;
  v_��Ʊ���� Varchar2(20);
  n_��Ʊ��� ����Ʊ��ʹ�ü�¼.Ʊ�ݽ��%Type;
  v_Url      ����Ʊ��ʹ�ü�¼.Url����%Type;
  v_Url����  ����Ʊ��ʹ�ü�¼.Url����%Type;

  v_Temp    Varchar2(32767); --��ʱXML 
  x_Templet Xmltype; --ģ��XML 

  n_Count   Number;
  v_Err_Msg Varchar2(200);
  Err_Item Exception;

Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/GHDH'), Extractvalue(Value(A), 'IN/JSKLB'), Extractvalue(Value(A), 'IN/GHJE'),
         Extractvalue(Value(A), 'IN/LSH'), To_Number(Extractvalue(Value(A), 'IN/JCFP')),
         Nvl(To_Number(Extractvalue(Value(A), 'IN/JKFS')), 0), Extractvalue(Value(A), 'IN/YYFS'),
         Extractvalue(Value(A), 'IN/JSMS'), Extractvalue(Value(A), 'IN/CZLX'), Extractvalue(Value(A), 'IN/CXID')
  Into v_No, v_�����, n_�ҺŽ��, v_������ˮ��, n_��鷢Ʊ, n_�ɿʽ, v_ԤԼ��ʽ, n_����ģʽ, n_��������, n_����id
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  Select Max(����id) Into n_����id From ������ü�¼ Where NO = v_No And ��¼���� = 4 And ��¼״̬ In (1, 3);

  --��Ҫ�Ե���Ʊ�ݳ�촦��
  If b_Einvoice_Request.Einvoice_Cancel(4, n_����id, v_Err_Msg) = 0 Then
    --����Ʊ������ʧ�� 
    Raise Err_Item;
  End If;

  If Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) <> 0 Then
    If Nvl(n_����id, 0) = 0 And Nvl(n_����id, 0) <> 0 Then
      v_Err_Msg := 'û��ָ����صĽ������ݣ�';
      Raise Err_Item;
    End If;
  
    Begin
      Select �Ǽ�ʱ�� Into d_�Ǽ�ʱ�� From ���˹Һż�¼ Where NO = v_No And ��¼״̬ In (1, 3) And Rownum < 2;
    Exception
      When Others Then
        v_Err_Msg := 'û���ҵ�ָ������ؽ������ݣ������ѱ�����';
        Raise Err_Item;
    End;
  End If;

  If Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) = 2 Then
    --ɾ����������
    Zl_���˹Һż�¼_Cancel(n_����id);
  
    v_Temp := '<CZSJ>' || To_Char(d_�Ǽ�ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<YJZID>' || n_����id || '</YJZID>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<CXID>' || n_����id || '</CXID>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<KPBZ>' || 0 || '</KPBZ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<URL>' || '' || '</URL>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<NETURL>' || '' || '</NETURL>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<FPTT>' || '' || '</FPTT>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<FPH>' || '' || '</FPH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<FPJE>' || '' || '</FPJE>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<KPRQ>' || '' || '</KPRQ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    Xml_Out := x_Templet;
    Return;
  End If;

  --��ȡ����Ա��Ϣ
  v_����Ա��� := Zl_����Ա��Ϣ(1);
  v_����Ա���� := Zl_����Ա��Ϣ(2);

  Select Max(�շѵ�) Into v_�շѵ� From ���˹Һż�¼ Where NO = v_No;
  If Nvl(n_����ģʽ, 0) = 0 Or Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) = 0 Then
    If n_�ɿʽ = 1 Then
      Select Count(1)
      Into n_Count
      From ������ü�¼
      Where NO = v_No And ��¼���� = 4 And ����id Is Not Null And Rownum < 2;
      If n_Count = 0 Then
        Select Count(1)
        Into n_Count
        From ������ü�¼
        Where NO In (Select /*+cardinality(B,10)*/
                      Column_Value
                     From Table(f_Str2List(v_�շѵ�)) B) And ��¼���� = 1 And ����id Is Not Null And Rownum < 2;
      End If;
      If n_Count <> 0 Then
        v_Err_Msg := '����ĹҺŵ��ݲ���ԤԼ�Һŵ�,�޷�ȡ��ԤԼ!';
        Raise Err_Item;
      End If;
    
      Select Count(1) Into n_Count From ���˹Һż�¼ A Where a.No = v_No And a.ԤԼ��ʽ = v_ԤԼ��ʽ And Rownum < 2;
      If n_Count = 0 Then
        v_Err_Msg := '����ĹҺŵ��ݲ���' || v_ԤԼ��ʽ || 'ԤԼ��,�޷�ȡ��ԤԼ!';
        Raise Err_Item;
      End If;
    End If;
  
    If v_����� Is Not Null And Nvl(n_�ɿʽ, 0) = 0 Then
      Select Nvl2(Translate(v_�����, '\1234567890', '\'), 'Char', 'Num') Into v_Type From Dual;
      If v_Type = 'Num' Then
        --������ǿ����ID 
        Select ���㷽ʽ, ID Into v_���㷽ʽ, n_�����id From ҽ�ƿ���� Where ID = To_Number(v_�����);
      Else
        --������ǿ�������� 
        Select ���㷽ʽ, ID Into v_���㷽ʽ, n_�����id From ҽ�ƿ���� Where ���� = v_�����;
      End If;
    
      --Ҫ�˵ĵ��ݲ����Ըý��㿨����ģ����ֹ�˺� 
      Select Count(1)
      Into n_Count
      From ����Ԥ����¼ A,
           (Select Distinct ����id
             From ������ü�¼
             Where NO = v_No And ��¼���� = 4
             Union
             Select Distinct ����id
             From סԺ���ü�¼
             Where NO = v_No And ��¼���� = 5
             Union
             Select Distinct ����id
             From ������ü�¼
             Where NO In (Select /*+cardinality(B,10)*/
                           Column_Value
                          From Table(f_Str2List(v_�շѵ�)) B) And ��¼���� = 1) B
      Where a.����id = b.����id And a.�����id = n_�����id And Rownum < 2;
      If n_Count = 0 Then
        v_Err_Msg := '����ĹҺŵ��ݲ���' || v_���㷽ʽ || '�����,�޷��˺�!';
        Raise Err_Item;
      End If;
    End If;
  
    --��������飬�Ѵ��ڲ��������ݵģ������˺� 
    Select Count(1)
    Into n_Count
    From ���ò����¼ A,
         (Select Distinct ����id
           From ������ü�¼
           Where NO = v_No And ��¼���� = 4
           Union
           Select Distinct ����id
           From סԺ���ü�¼
           Where NO = v_No And ��¼���� = 5
           Union
           Select Distinct ����id
           From ������ü�¼
           Where NO In (Select /*+cardinality(B,10)*/
                         Column_Value
                        From Table(f_Str2List(v_�շѵ�)) B) And ��¼���� = 1) B
    Where a.�շѽ���id = b.����id And a.��¼���� = 1 And a.���ӱ�־ = 1 And Nvl(a.����״̬, 0) <> 2 And Rownum < 2;
    If n_Count = 1 Then
      v_Err_Msg := '����ĹҺŵ����Ѿ������˶��ν���,�޷��˺�!';
      Raise Err_Item;
    End If;
  
    --ҽ����飬�Ѿ�����ҽ���ģ������˺� 
    Select Count(1) Into n_�ѿ�ҽ�� From ����ҽ����¼ Where �Һŵ� = v_No And Rownum < 2;
    If n_�ѿ�ҽ�� = 1 Then
      v_Err_Msg := '����ĹҺŵ����Ѿ�����ҽ��,�޷��˺�!';
      Raise Err_Item;
    End If;
  
    If Nvl(n_��鷢Ʊ, 0) = 1 Then
      Select Max(Decode(a.ʵ��Ʊ��, Null, 0, 1)) Into n_�Ƿ��ӡ From ������ü�¼ A Where NO = v_No And ��¼���� = 4;
      If Nvl(n_�Ƿ��ӡ, 0) = 1 Then
        v_Err_Msg := '�����˺ŵĵ����ѿ���Ʊ,�����˷�!';
        Raise Err_Item;
      End If;
    
      Select Max(Decode(a.ʵ��Ʊ��, Null, 0, 1))
      Into n_�Ƿ��ӡ
      From ������ü�¼ A
      Where NO In (Select /*+cardinality(B,10)*/
                    Column_Value
                   From Table(f_Str2List(v_�շѵ�)) B) And ��¼���� = 1;
      If Nvl(n_�Ƿ��ӡ, 0) = 1 Then
        v_Err_Msg := '�����˺ŵĵ����ѿ���Ʊ,�����˷�!';
        Raise Err_Item;
      End If;
    End If;
  
    d_�Ǽ�ʱ�� := Sysdate;
    Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
    Zl_���������Һ�_Delete(v_No, v_������ˮ��, '�ƶ�ƽ̨�˺�', d_�Ǽ�ʱ��, Null, 1, 0, n_����id);
  
    If Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) = 0 And Nvl(n_�����id, 0) > 0 Then
      For c_��¼ In (Select NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, -1 * ��Ԥ�� As ��Ԥ��, ������λ, �����id, ����, ������ˮ��, ��������,
                          ��������id
                   From ����Ԥ����¼
                   Where ��¼���� = 4 And ��¼״̬ In (1, 3) And ����id = n_����id And �����id = n_�����id) Loop
      
        Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ������ˮ��, ����˵��,
           ������λ, �������, �����id, ����, ��������, ����ʱ��, ������Ա, ��������id, У�Ա�־)
          Select n_Ԥ��id, c_��¼.No, c_��¼.ʵ��Ʊ��, c_��¼.��¼����, 2, c_��¼.����id, c_��¼.��ҳid, c_��¼.����id, c_��¼.ժҪ, c_��¼.���㷽ʽ, d_�Ǽ�ʱ��,
                 ����Ա���, ����Ա����, c_��¼.��Ԥ��, n_����id, �ɿ���id, c_��¼.������ˮ��, '�ƶ�ƽ̨�˺�', c_��¼.������λ, n_����id, c_��¼.�����id, c_��¼.����,
                 c_��¼.��������, d_�Ǽ�ʱ��, ����Ա����, c_��¼.��������id, 1
          From ������ü�¼
          Where ��¼���� = 4 And ����id = n_����id And Rownum < 2;
      End Loop;
    End If;
  Else
    If v_����� Is Not Null And Nvl(n_�ɿʽ, 0) = 0 Then
      Select Nvl2(Translate(v_�����, '\1234567890', '\'), 'Char', 'Num') Into v_Type From Dual;
      If v_Type = 'Num' Then
        Select ID Into n_�����id From ҽ�ƿ���� Where ID = To_Number(v_�����);
      Else
        Select ID Into n_�����id From ҽ�ƿ���� Where ���� = v_�����;
      End If;
      Delete From ����Ԥ����¼ Where ����id = n_����id And n_�����id = n_�����id;
    End If;
  End If;

  If Nvl(n_����ģʽ, 0) = 0 Or Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) = 1 Then
    Zl_���������Һ�_Delete(v_No, v_������ˮ��, '�ƶ�ƽ̨�˺�', d_�Ǽ�ʱ��, Null, 1, 1, n_����id);
    n_�Һ�ԭ����id := n_����id;
    n_�Һų���id   := n_����id;
  
    --ͬ�������۵� 
    If v_�շѵ� Is Not Null Then
      n_Temp := 0;
      For c_�Һ� In (Select NO, Max(��¼״̬) As ��¼״̬, Max(����id) As ����id, Max(Decode(��¼״̬, 2, 0, ����id)) As ԭ����id,
                          Max(Decode(��¼״̬, 2, ����id, 0)) As ����id
                   From ������ü�¼
                   Where NO In (Select /*+cardinality(B,10)*/
                                 Column_Value
                                From Table(f_Str2List(v_�շѵ�)) B) And ��¼���� = 1) Loop
      
        If Nvl(c_�Һ�.��¼״̬, 0) = 0 Then
          Zl_���ﻮ�ۼ�¼_Delete(c_�Һ�.No);
          n_����id := c_�Һ�.ԭ����id;
          n_����id := c_�Һ�.����id;
        Elsif Nvl(c_�Һ�.��¼״̬, 0) = 1 Then
          If v_���㷽ʽ Is Null Then
            v_Err_Msg := '���ιҺŵ����˿�ʧ��,����!';
            Raise Err_Item;
          End If;
          --��Ҫ�Ե���Ʊ�ݳ�촦��
          If b_Einvoice_Request.Einvoice_Cancel(1, c_�Һ�.ԭ����id, v_Err_Msg) = 0 Then
            --����Ʊ������ʧ�� 
            Raise Err_Item;
          End If;
        
          Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
          Zl_�����շѼ�¼_����(c_�Һ�.No, v_����Ա���, v_����Ա����, Null, d_�Ǽ�ʱ��, Null, n_����id);
          v_�˷ѽ��� := v_���㷽ʽ || '|' || -1 * n_�ҺŽ�� || '|' || ' |' || ' ';
          Zl_�����˷ѽ���_Modify(2, n_����id, n_����id, v_�˷ѽ���, 0, n_�����id, Null, v_������ˮ��, Null, 0, 0, 0, 2);
        
          n_����id := c_�Һ�.ԭ����id;
          n_����id := c_�Һ�.����id;
          n_Temp   := n_Temp + 1;
        Else
          n_����id := c_�Һ�.ԭ����id;
          n_����id := c_�Һ�.����id;
        End If;
      
      End Loop;
    
      If n_Temp > 1 Then
        v_Err_Msg := '���ιҺŴ��ڶ���շѣ������˷Ѻ����˺�!';
        Raise Err_Item;
      End If;
    End If;
  
    --�������Ʊ��
    n_Count := 0;
    Select Sum(���ʽ��) Into n_ʣ���� From ������ü�¼ Where NO = v_No And ��¼���� = 4;
    Select Max(�Ƿ����Ʊ��) Into n_�Ƿ����Ʊ�� From ����Ԥ����¼ Where ����id = n_�Һ�ԭ����id;
  
    Update ����Ԥ����¼ Set �Ƿ����Ʊ�� = n_�Ƿ����Ʊ�� Where ����id = n_�Һų���id;
  
    If Nvl(n_ʣ����, 0) <> 0 And Nvl(n_�Ƿ����Ʊ��, 0) = 1 Then
      --���ֽ��ʣ���Ҫ���¿��ߵ���Ʊ��
      If b_Einvoice_Request.Einvoice_Create(4, n_�Һ�ԭ����id, n_�Һų���id, v_Err_Msg) = 0 Then
        Raise Err_Item;
      End If;
    
      Select Max(1), Max(����), Max(����), Max(To_Char(����ʱ��, 'yyyy-mm-dd')), Max(Url����), Max(Url����), Max(Ʊ�ݽ��)
      Into n_��Ʊ��־, v_��������, v_��Ʊ���, v_��Ʊ����, v_Url, v_Url����, n_��Ʊ���
      From ����Ʊ��ʹ�ü�¼
      Where ����id = n_����id And Ʊ�� = 4 And ��¼״̬ = 1;
    
      n_Count := 1;
    End If;
  
    If Nvl(n_�Һ�ԭ����id, 0) <> Nvl(n_����id, 0) Then
      --�շѲ��ֵĴ���
      Select Max(�Ƿ����Ʊ��) Into n_�Ƿ����Ʊ�� From ����Ԥ����¼ Where ����id = n_����id;
      If Nvl(n_�Ƿ����Ʊ��, 0) = 1 Then
      
        Update ����Ԥ����¼ Set �Ƿ����Ʊ�� = n_�Ƿ����Ʊ�� Where ����id = n_����id;
      
        Select Sum(���ʽ��)
        Into n_ʣ����
        From ������ü�¼
        Where NO In (Select Distinct NO From ������ü�¼ Where ����id = n_����id) And Mod(��¼����, 10) = 1;
      
        If Nvl(n_ʣ����, 10) <> 0 Then
          --���ֿ���
          If b_Einvoice_Request.Einvoice_Create(1, n_����id, n_����id, v_Err_Msg) = 0 Then
            If Nvl(n_Count, 0) = 0 Then
              --�ҺŴ����˵���Ʊ�ݣ������״�
              Raise Err_Item;
            End If;
          Else
            If Nvl(n_��Ʊ��־, 0) = 1 Then
              Select Max(1), Max(����), Max(����), Max(To_Char(����ʱ��, 'yyyy-mm-dd')), Max(Url����), Max(Url����), Max(Ʊ�ݽ��)
              Into n_��Ʊ��־, v_��������, v_��Ʊ���, v_��Ʊ����, v_Url, v_Url����, n_��Ʊ���
              From ����Ʊ��ʹ�ü�¼
              Where ����id = n_����id And Ʊ�� = 1 And ��¼״̬ = 1;
            
            End If;
          End If;
        End If;
      End If;
    End If;
  End If;

  v_Temp := '<CZSJ>' || To_Char(d_�Ǽ�ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<YJZID>' || n_����id || '</YJZID>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<CXID>' || n_����id || '</CXID>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<KPBZ>' || Nvl(n_��Ʊ��־, 0) || '</KPBZ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<URL>' || Nvl(v_Url, '') || '</URL>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<NETURL>' || Nvl(v_Url����, '') || '</NETURL>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<FPTT>' || v_�������� || '</FPTT>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<FPH>' || v_��Ʊ��� || '</FPH>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<FPJE>' || Nvl(n_��Ʊ���, 0) || '</FPJE>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<KPRQ>' || v_��Ʊ���� || '</KPRQ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Registdel;
/


Create Or Replace Procedure Zl_�����շѼ�¼_Insert
(
  No_In            ������ü�¼.No%Type,
  ���_In          ������ü�¼.���%Type,
  ����id_In        ������ü�¼.����id%Type,
  ������Դ_In      Number,
  ��ʶ��_In        ������ü�¼.��ʶ��%Type,
  ���ʽ_In      ������ü�¼.���ʽ%Type,
  ����_In          ������ü�¼.����%Type,
  �Ա�_In          ������ü�¼.�Ա�%Type,
  ����_In          ������ü�¼.����%Type,
  �ѱ�_In          ������ü�¼.�ѱ�%Type,
  �Ӱ��־_In      ������ü�¼.�Ӱ��־%Type,
  ���˿���id_In    ������ü�¼.���˿���id%Type,
  ��������id_In    ������ü�¼.��������id%Type,
  ������_In        ������ü�¼.������%Type,
  ��������_In      ������ü�¼.��������%Type,
  �շ�ϸĿid_In    ������ü�¼.�շ�ϸĿid%Type,
  �շ����_In      ������ü�¼.�շ����%Type,
  ���㵥λ_In      ������ü�¼.���㵥λ%Type,
  ������Ŀ��_In    ������ü�¼.������Ŀ��%Type,
  ���մ���id_In    ������ü�¼.���մ���id%Type,
  ��ҩ����_In      ������ü�¼.��ҩ����%Type,
  ����_In          ������ü�¼.����%Type,
  ����_In          ������ü�¼.����%Type,
  ���ӱ�־_In      ������ü�¼.���ӱ�־%Type,
  ִ�в���id_In    ������ü�¼.ִ�в���id%Type,
  �۸񸸺�_In      ������ü�¼.�۸񸸺�%Type,
  ������Ŀid_In    ������ü�¼.������Ŀid%Type,
  �վݷ�Ŀ_In      ������ü�¼.�վݷ�Ŀ%Type,
  ��׼����_In      ������ü�¼.��׼����%Type,
  Ӧ�ս��_In      ������ü�¼.Ӧ�ս��%Type,
  ʵ�ս��_In      ������ü�¼.ʵ�ս��%Type,
  ͳ����_In      ������ü�¼.ͳ����%Type,
  ����ʱ��_In      ������ü�¼.����ʱ��%Type,
  �Ǽ�ʱ��_In      ������ü�¼.�Ǽ�ʱ��%Type,
  ԭno_In          ������ü�¼.No%Type,
  ����id_In        ������ü�¼.����id%Type,
  �շѽ���_In      Varchar2,
  ��Ԥ����_In      ����Ԥ����¼.��Ԥ��%Type,
  ���ս���_In      Varchar2,
  ����Ա���_In    ������ü�¼.����Ա���%Type,
  ����Ա����_In    ������ü�¼.����Ա����%Type,
  ժҪ_In          ������ü�¼.ժҪ%Type := Null,
  �Ƿ���_In      ������ü�¼.�Ƿ���%Type := 0,
  �÷�_In          ҩƷ�շ���¼.�÷�%Type := Null, --�÷�[|�巨]
  �ɿ�_In          ����Ԥ����¼.�ɿ�%Type := Null,
  �Ҳ�_In          ����Ԥ����¼.�Ҳ�%Type := Null,
  ��ҩ��̬_In      ������ü�¼.����%Type := Null,
  ��Ԥ������ids_In Varchar2 := Null,
  �Ƿ����Ʊ��_In  ����Ԥ����¼.�Ƿ����Ʊ��%Type := 0
) As
  --���ܣ�����һ�������շѵ���
  --������
  --  ������Դ_IN:1-����;2-סԺ  סԺ�����շ�ʱ�á�
  --  ԭNO_IN:�޸ı����µ���ʱ�á�Ŀǰ���ڴ����ҩƷ�շ���¼��ժҪ�С�
  --         ԭ����(��¼״̬=2)��¼�޸Ĳ������µ��ݺš�
  --         �µ���(��¼״̬=1)��¼���޸ĵ�ԭ���ݺš�
  -- �շѽ���_IN:��ʽ="���㷽ʽ|������|�������|����ժҪ||.....",ע���޽�������ժҪʱҪ�ÿո����
  -- ���ս���_IN:��ʽ="���㷽ʽ|������||....."
  -- ��Ԥ������ids_In:����ö��ַ���,��Ԥ��ʱ��Ч(��Ԥ�������ҵ���������һ��),��Ҫ��ʹ�ü�����Ԥ����
  -- �Ƿ����Ʊ��_In �Ƿ�ʹ�õ���Ʊ��
  v_����id ������ü�¼.Id%Type;

  v_�÷� ҩƷ�շ���¼.�÷�%Type;
  v_�巨 ҩƷ�շ���¼.���%Type;
  ------------------------------------------------------------
  --���㷽ʽ��
  v_�������� Varchar2(3000);
  v_��ǰ���� Varchar2(150);
  v_���㷽ʽ ����Ԥ����¼.���㷽ʽ%Type;
  v_������ ����Ԥ����¼.��Ԥ��%Type;
  v_������� ����Ԥ����¼.�������%Type;
  v_����ժҪ ����Ԥ����¼.ժҪ%Type;
  n_����ֵ   �������.�������%Type;

  v_Dec        Number;
  v_���ʽ   ҽ�Ƹ��ʽ.����%Type;
  v_�ѱ�����   �ѱ�.����%Type;
  n_�²���ģʽ Number;

  --��ʱ����
  Err_Custom Exception;
  v_Error       Varchar2(255);
  n_��id        ����ɿ����.Id%Type;
  n_����С��    Number;
  v_Strtmpbefor Varchar2(4000);
  v_Msg         Varchar2(4000);

Begin
  n_��id := Zl_Get��id(����Ա����_In);

  --���С��λ��
  Select zl_To_Number(Nvl(zl_GetSysParameter(9), '2')), zl_To_Number(Nvl(zl_GetSysParameter(157), '5'))
  Into v_Dec, n_����С��
  From Dual;

  --������ü�¼
  Select ���˷��ü�¼_Id.Nextval Into v_����id From Dual;
  Insert Into ������ü�¼
    (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �����־, ����id, ��ʶ��, ���ʽ, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��,
     ���մ���id, ����, ����, ��ҩ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����, ���ʷ���, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id,
     ִ��״̬, ����id, ���ʽ��, ����Ա���, ����Ա����, ժҪ, �Ƿ���, ����, �ɿ���id)
  Values
    (v_����id, 1, No_In, 1, ���_In, Decode(��������_In, 0, Null, ��������_In), Decode(�۸񸸺�_In, 0, Null, �۸񸸺�_In),
     Decode(������Դ_In, 1, 1, 2), Decode(����id_In, 0, Null, ����id_In), Decode(��ʶ��_In, 0, Null, ��ʶ��_In), ���ʽ_In, ����_In, �Ա�_In,
     ����_In, ���˿���id_In, �ѱ�_In, �շ����_In, �շ�ϸĿid_In, ���㵥λ_In, ������Ŀ��_In, ���մ���id_In, ����_In, ����_In, ��ҩ����_In, �Ӱ��־_In, ���ӱ�־_In,
     ������Ŀid_In, �վݷ�Ŀ_In, ��׼����_In, Ӧ�ս��_In, ʵ�ս��_In, ͳ����_In, 0, ����Ա����_In, ��������id_In, ������_In, ����ʱ��_In, �Ǽ�ʱ��_In, ִ�в���id_In,
     0, ����id_In, ʵ�ս��_In, ����Ա���_In, ����Ա����_In, ժҪ_In, �Ƿ���_In, ��ҩ��̬_In, n_��id);

  If ���_In = 1 Then
    --����Ԥ����¼(��һ��ʱ����)
    --��������
    If �շѽ���_In Is Not Null Then
      --�����շѽ���
      v_�������� := �շѽ���_In || '||';
      While v_�������� Is Not Null Loop
      
        v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
        v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      
        v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
        v_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      
        v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
        v_������� := LTrim(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
        v_����ժҪ := LTrim(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
      
        If Nvl(v_������, 0) <> 0 Then
          Insert Into ����Ԥ����¼
            (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ�, �Ҳ�, �ɿ���id, ��������)
          Values
            (����Ԥ����¼_Id.Nextval, 3, No_In, 1, Decode(����id_In, 0, Null, ����id_In), Null, v_����ժҪ, v_���㷽ʽ, v_�������, �Ǽ�ʱ��_In,
             ����Ա���_In, ����Ա����_In, v_������, ����id_In, Decode(v_��������, �շѽ���_In || '||', �ɿ�_In, Null),
             Decode(v_��������, �շѽ���_In || '||', �Ҳ�_In, Null), n_��id, 3);
        End If;
      
        v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
      End Loop;
    End If;
  
    --�������ս���
    If ���ս���_In Is Not Null Then
      v_�������� := ���ս���_In || '||';
      While v_�������� Is Not Null Loop
        v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      
        v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
        v_������ := To_Number(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
      
        If Nvl(v_������, 0) <> 0 Then
          Insert Into ����Ԥ����¼
            (ID, ��¼����, NO, ��¼״̬, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ��������)
          Values
            (����Ԥ����¼_Id.Nextval, 3, No_In, 1, Decode(����id_In, 0, Null, ����id_In), '���ս���', v_���㷽ʽ, �Ǽ�ʱ��_In, ����Ա���_In,
             ����Ա����_In, v_������, ����id_In, n_��id, 3);
        End If;
        v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
      End Loop;
    End If;
  
    --Ԥ������
    If Nvl(��Ԥ����_In, 0) <> 0 Then
      Zl_����Ԥ����¼_��Ԥ��(����id_In, ����id_In, ��Ԥ����_In, 1, ����Ա���_In, ����Ա����_In, �Ǽ�ʱ��_In, ��Ԥ������ids_In, 3);
    End If;
  End If;

  Update ����Ԥ����¼ Set �Ƿ����Ʊ�� = �Ƿ����Ʊ��_In Where ����id = ����id_In And ��¼���� <> 1;

  --��ػ��ܱ�Ĵ���
  --����"��Ա�ɿ����"(ע��Ҫ��������ʻ��Ľ���)
  n_����ֵ := 0;
  If ���_In = 1 Then
    --�����շѽ���
    If �շѽ���_In Is Not Null Then
      v_�������� := �շѽ���_In || '||';
      While v_�������� Is Not Null Loop
        v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
        v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
        v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
        v_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      
        If Nvl(v_������, 0) <> 0 Then
          Update ��Ա�ɿ����
          Set ��� = Nvl(���, 0) + Nvl(v_������, 0)
          Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ
          Returning ��� + n_����ֵ Into n_����ֵ;
        
          If Sql%RowCount = 0 Then
            Insert Into ��Ա�ɿ����
              (�տ�Ա, ���㷽ʽ, ����, ���)
            Values
              (����Ա����_In, v_���㷽ʽ, 1, Nvl(v_������, 0));
            n_����ֵ := n_����ֵ + Nvl(v_������, 0);
          End If;
        End If;
      
        v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
      End Loop;
    End If;
  
    --�������ս���
    If ���ս���_In Is Not Null Then
      v_�������� := ���ս���_In || '||';
      While v_�������� Is Not Null Loop
        v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      
        v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
        v_������ := To_Number(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
      
        If Nvl(v_������, 0) <> 0 Then
          Update ��Ա�ɿ����
          Set ��� = Nvl(���, 0) + Nvl(v_������, 0)
          Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ
          Returning ��� + n_����ֵ Into n_����ֵ;
        
          If Sql%RowCount = 0 Then
            Insert Into ��Ա�ɿ����
              (�տ�Ա, ���㷽ʽ, ����, ���)
            Values
              (����Ա����_In, v_���㷽ʽ, 1, Nvl(v_������, 0));
            n_����ֵ := n_����ֵ + Nvl(v_������, 0);
          End If;
        End If;
      
        v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
      End Loop;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ��Ա�ɿ���� Where ���� = 1 And �տ�Ա = ����Ա����_In And Nvl(���, 0) = 0;
    End If;
  End If;

  --ҩƷ���������ϲ���
  If �շ����_In In ('4', '5', '6', '7') Then
    --ҩƷ�÷��巨�ֽ�
    If �÷�_In Is Not Null Then
      If Instr(�÷�_In, '|') > 0 Then
        v_�÷� := Substr(�÷�_In, 1, Instr(�÷�_In, '|') - 1);
        v_�巨 := Substr(�÷�_In, Instr(�÷�_In, '|') + 1);
      Else
        v_�÷� := �÷�_In;
      End If;
    End If;
    Zl_ҩƷ�շ���¼_���۳���(v_����id, ԭno_In, Null, Null, v_�÷�, v_�巨);
  End If;

  --���²��ݲ�����Ϣ
  If ���_In = 1 And ����id_In Is Not Null Then
    If ���ʽ_In Is Not Null And Nvl(������Դ_In, 1) = 1 Then
      Select ���� Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = ���ʽ_In;
    End If;
    Select Max(����) Into v_�ѱ����� From �ѱ� Where ���� = �ѱ�_In; --2-��̬�ѱ𲻸���
  
    Select Zl_Fun_Checkidentify(0, ����id_In, v_Strtmpbefor) Into v_Strtmpbefor From Dual;
    Update ������Ϣ
    Set �Ա� = Decode(����, '�²���', Nvl(�Ա�_In, �Ա�), �Ա�), ���� = Decode(����, '�²���', Nvl(����_In, ����), ����),
        ���� = Decode(����, '�²���', ����_In, ����), ҽ�Ƹ��ʽ = Nvl(v_���ʽ, ҽ�Ƹ��ʽ), �ѱ� = Decode(v_�ѱ�����, 1, �ѱ�_In, �ѱ�)
    Where ����id = ����id_In;
    Select Zl_Fun_Checkidentify(1, ����id_In, v_Strtmpbefor) Into v_Msg From Dual;
    Select zl_To_Number(Nvl(zl_GetSysParameter('�Զ���������', '1111'), '0')) Into n_�²���ģʽ From Dual;
    If n_�²���ģʽ = 1 Then
      Update ���˹Һż�¼
      Set ���� = ����_In, �Ա� = �Ա�_In, ���� = ����_In
      Where ����id = ����id_In And ���� = '�²���';
      Update ������ü�¼
      Set ���� = ����_In, �Ա� = �Ա�_In, ���� = ����_In, ���ʽ = ���ʽ_In
      Where ����id = ����id_In And ���� = '�²���';
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����շѼ�¼_Insert;
/
Create Or Replace Procedure Zl_�����շѼ�¼_Insert
(
  No_In            ������ü�¼.No%Type,
  ����id_In        ������ü�¼.����id%Type,
  ������Դ_In      Number,
  ���ʽ_In      ������ü�¼.���ʽ%Type,
  ����_In          ������ü�¼.����%Type,
  �Ա�_In          ������ü�¼.�Ա�%Type,
  ����_In          ������ü�¼.����%Type,
  ���˿���id_In    ������ü�¼.���˿���id%Type,
  ��������id_In    ������ü�¼.��������id%Type,
  ������_In        ������ü�¼.������%Type,
  �շѽ���_In      Varchar2,
  ��Ԥ����_In      ����Ԥ����¼.��Ԥ��%Type,
  ���ս���_In      Varchar2,
  ����id_In        ������ü�¼.����id%Type,
  ����ʱ��_In      ������ü�¼.����ʱ��%Type,
  ����Ա���_In    ������ü�¼.����Ա���%Type,
  ����Ա����_In    ������ü�¼.����Ա����%Type,
  ��ҩ����_In      Varchar2,
  �Ƿ���_In      ������ü�¼.�Ƿ���%Type := 0,
  �ɿ�_In          ����Ԥ����¼.�ɿ�%Type := Null,
  �Ҳ�_In          ����Ԥ����¼.�Ҳ�%Type := Null,
  ����������_In    Varchar2 := Null,
  �Ǽ�ʱ��_In      ������ü�¼.�Ǽ�ʱ��%Type := Null,
  �������_In      ����Ԥ����¼.�������%Type := Null,
  ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In      ����Ԥ����¼.����˵��%Type := Null,
  ���շ�_In      Number := 0,
  ��Ԥ������ids_In Varchar2 := Null,
  �Ƿ����Ʊ��_In  ����Ԥ����¼.�Ƿ����Ʊ��%Type := 0
) As
  --���ܣ������շ�ʱ��ȡ���۵�����
  --������
  -- ��ҩ����_In:ִ�в���ID1|��ҩ����1;...;ִ�в���IDn|��ҩ����n
  -- ������Դ_IN:1-����;2-סԺ
  -- �շѽ���_IN:��ʽ="���㷽ʽ|������|�������|����ժҪ||.....",ע���޽�������ժҪʱҪ�ÿո����
  -- ���ս���_IN:��ʽ="���㷽ʽ|������||....."
  -- ����������_In:��ʽ=�����Id|�Ƿ����ѿ�|������|����|��ע||...
  -- ������ˮ��_In�ͽ���˵��_In:�շѽ���_INʱ��Ч.
  -- ��Ԥ������ids_In:����ö��ַ���,��Ԥ��ʱ��Ч(��Ԥ�������ҵ���������һ��),��Ҫ��ʹ�ü�����Ԥ����
  -- �Ƿ����Ʊ��_In �Ƿ�ʹ�õ���Ʊ�� 
  --˵����
  --        1.��ȡ���۷���ʱ,�ż��������ػ���,�ڻ���ʱ������;��ҩƷ��ػ���(��������)����ʱ�Ѿ����㡣
  --        2.��ȡ���۷���ʱ,Ŀǰ���漰������δ������չ�����,�ɻ���ʱֱ�Ӵ���  

  --=================================
  --��ע���ù���Ŀǰֻ�м��շ�ʹ�ã�
  --=================================

  Cursor c_Price Is
    Select ID
    From ������ü�¼
    Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 0 And ����Ա���� Is Null
    Order By ���;

  n_Array_Size Number := 200;
  t_����id     t_NumList;
  v_��������   ���ű�.����%Type;

  --Ԥ���������ر���
  v_�������� Varchar2(3000);
  v_��ǰ���� Varchar2(150);
  v_���㷽ʽ ����Ԥ����¼.���㷽ʽ%Type;
  n_������ ����Ԥ����¼.��Ԥ��%Type;
  v_������� ����Ԥ����¼.�������%Type;
  v_����ժҪ ����Ԥ����¼.ժҪ%Type;

  n_����id   ������ü�¼.����id%Type;
  v_��ʶ��   ������ü�¼.��ʶ��%Type;
  v_���ʽ ҽ�Ƹ��ʽ.����%Type;
  n_����ֵ   �������.Ԥ�����%Type;

  --��ʱ����
  n_Count      Number;
  n_�²���ģʽ Number;
  v_����no     ҩƷ�շ���¼.No%Type;
  v_Date       Date;
  v_Err_Msg    Varchar2(255);
  Err_Item Exception;
  n_��id        ����ɿ����.Id%Type;
  n_�����id    ҽ�ƿ����.Id%Type;
  n_���ѿ�      Number;
  v_����        ����Ԥ����¼.����%Type;
  v_������      Varchar2(100);
  n_Ԥ��id      ����Ԥ����¼.Id%Type;
  n_���ѿ�id    ���ѿ���Ϣ.Id%Type;
  v_Strtmpbefor Varchar2(4000);
  v_Msg         Varchar2(4000);
Begin
  n_��id := Zl_Get��id(����Ա����_In);

  Select Count(ID), Max(����id)
  Into n_Count, n_����id
  From ������ü�¼
  Where ��¼���� = 1 And ��¼״̬ = 0 And NO = No_In And ����Ա���� Is Null;
  If n_Count = 0 Then
    v_Err_Msg := '���ܶ�ȡ���۵�����,�õ��ݿ����Ѿ�ɾ�����Ѿ��շѣ�';
    Raise Err_Item;
  End If;
  If Nvl(n_����id, 0) <> 0 And Nvl(n_����id, 0) <> Nvl(����id_In, 0) Then
    v_Err_Msg := '���ݡ�' || No_In || '�����ǵ�ǰ���˵ķ��ã����ܶ�������շѣ�';
    Raise Err_Item;
  End If;

  v_Date := �Ǽ�ʱ��_In;
  If v_Date Is Null Then
    Select Sysdate Into v_Date From Dual;
  End If;
  If Nvl(����id_In, 0) <> 0 Then
    Select Decode(��ǰ����id, Null, �����, סԺ��) Into v_��ʶ�� From ������Ϣ Where ����id = ����id_In;
  End If;

  ------------------------------------------------------------------------------------------------------------------------
  --��������
  Open c_Price;
  Loop
    Fetch c_Price Bulk Collect
      Into t_����id Limit n_Array_Size;
    Exit When t_����id.Count = 0;
  
    --ѭ������������ü�¼
    Forall I In 1 .. t_����id.Count
    --ִ��״̬����ֶβ�����,�ڻ���ʱ����;��Ϊ����δ�շѷ�ҩ,������ִ�еĻ��۵��������շѲ����ġ�
    --Ϊ��֤��Ԥ�������¼��ʱ����ͬ,������д�Ǽ�ʱ��,��ҩƷ���ֲ��䶯��
      Update ������ü�¼
      Set ��¼״̬ = 1, ����id = Decode(����id_In, 0, Null, ����id_In), ��ʶ�� = v_��ʶ��, ���ʽ = ���ʽ_In, ���� = ����_In, ���� = ����_In,
          �Ա� = �Ա�_In,
          --���ܱ���ҽ�����͵�����
          ���˿���id = Nvl(���˿���id_In, ���˿���id), ��������id = Nvl(��������id_In, ��������id), ������ = Nvl(������_In, ������), ���ʽ�� = ʵ�ս��,
          ����id = ����id_In, ����ʱ�� = ����ʱ��_In, �Ǽ�ʱ�� = v_Date, ����Ա��� = ����Ա���_In, ����Ա���� = ����Ա����_In, �Ƿ��� = �Ƿ���_In,
          �ɿ���id = n_��id
      Where ID = t_����id(I) And ��¼״̬ = 0;
  
    If Sql%RowCount <> t_����id.Count Then
      v_Err_Msg := '���ڲ�������,�õ��ݿ����Ѿ�ɾ�����Ѿ��շѣ�';
      Raise Err_Item;
    End If;
  
  End Loop;
  Close c_Price;
  ------------------------------------------------------------------------------------------------------------------------

  --Ԥ������ؽ���
  --�շѽ���
  If �շѽ���_In Is Not Null Then
    v_�������� := �շѽ���_In || '||';
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      v_������� := LTrim(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_����ժҪ := LTrim(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
    
      If Nvl(n_������, 0) <> 0 Then
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ�, �Ҳ�, �ɿ���id, �������, ������ˮ��,
           ����˵��, ��������)
        Values
          (����Ԥ����¼_Id.Nextval, 3, No_In, 1, Decode(����id_In, 0, Null, ����id_In), Null, v_����ժҪ, v_���㷽ʽ, v_�������, v_Date,
           ����Ա���_In, ����Ա����_In, n_������, ����id_In, Decode(v_��������, �շѽ���_In || '||', �ɿ�_In, Null),
           Decode(v_��������, �շѽ���_In || '||', �Ҳ�_In, Null), n_��id, �������_In, ������ˮ��_In, ����˵��_In, 3);
      End If;
    
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  If ���ս���_In Is Not Null Then
    --�������ս���
    v_�������� := ���ս���_In || '||';
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      n_������ := To_Number(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
    
      If Nvl(n_������, 0) <> 0 Then
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, ��������)
        Values
          (����Ԥ����¼_Id.Nextval, 3, No_In, 1, Decode(����id_In, 0, Null, ����id_In), '���ս���', v_���㷽ʽ, v_Date, ����Ա���_In,
           ����Ա����_In, n_������, ����id_In, n_��id, �������_In, 3);
      End If;
    
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  If ����������_In Is Not Null Then
    v_�������� := ����������_In || '||';
    While v_�������� Is Not Null Loop
      --�����Id|�Ƿ����ѿ�|������|����|��ע||...
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      n_�����id := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_���ѿ�   := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      v_����     := LTrim(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      v_����ժҪ := v_��ǰ����;
    
      If n_���ѿ� = 1 Then
        Select ���㷽ʽ, ���� Into v_���㷽ʽ, v_������ From ���ѿ����Ŀ¼ Where ��� = n_�����id;
        If v_���㷽ʽ Is Null Then
          v_Err_Msg := v_������ || 'δ���ý��㷽ʽ����,�������ѿ��н�������,����ʧ�ܣ�';
          Raise Err_Item;
        End If;
      Else
        Select ���㷽ʽ, ���� Into v_���㷽ʽ, v_������ From ҽ�ƿ���� Where ID = n_�����id;
        If v_���㷽ʽ Is Null Then
          v_Err_Msg := v_������ || 'δ���ý��㷽ʽ����,����ҽ�ƿ������н�������,����ʧ�ܣ�';
          Raise Err_Item;
        End If;
      End If;
      If Nvl(n_������, 0) <> 0 Then
        Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, �����id, ���㿨���, ���㷽ʽ, �������, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ�, �Ҳ�, �ɿ���id,
           �������, ����, ��������)
        Values
          (n_Ԥ��id, 3, No_In, 1, Decode(����id_In, 0, Null, ����id_In), Null, v_����ժҪ, Decode(n_���ѿ�, 1, Null, n_�����id),
           Decode(n_���ѿ�, 0, Null, n_�����id), v_���㷽ʽ, v_�������, v_Date, ����Ա���_In, ����Ա����_In, n_������, ����id_In, Null, Null,
           n_��id, �������_In, v_����, 3);
      
        --���������
        If n_���ѿ� = 1 Then
          Zl_���˿������¼_֧��(n_�����id, v_����, n_���ѿ�id, n_������, n_Ԥ��id, ����Ա���_In, ����Ա����_In, v_Date);
        End If;
      End If;
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  --Ԥ������
  If Nvl(��Ԥ����_In, 0) <> 0 Then
    Zl_����Ԥ����¼_��Ԥ��(����id_In, ����id_In, ��Ԥ����_In, 1, ����Ա���_In, ����Ա����_In, �Ǽ�ʱ��_In, ��Ԥ������ids_In, 3, 1);
  End If;

  Update ����Ԥ����¼ Set �Ƿ����Ʊ�� = �Ƿ����Ʊ��_In Where ����id = ����id_In And ��¼���� <> 1;

  --��ػ��ܱ�Ĵ���

  --����"��Ա�ɿ����"
  --�շѽ���
  n_����ֵ := 0;
  If �շѽ���_In Is Not Null Then
    v_�������� := �շѽ���_In || '||';
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
    
      If Nvl(n_������, 0) <> 0 Then
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + Nvl(n_������, 0)
        Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ
        Returning Nvl(���, 0) + n_����ֵ Into n_����ֵ;
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ����
            (�տ�Ա, ���㷽ʽ, ����, ���)
          Values
            (����Ա����_In, v_���㷽ʽ, 1, Nvl(n_������, 0));
          n_����ֵ := Nvl(n_����ֵ, 0) + Nvl(n_������, 0);
        End If;
      End If;
    
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  --�������ս���
  If ���ս���_In Is Not Null Then
    v_�������� := ���ս���_In || '||';
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      n_������ := To_Number(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
    
      If Nvl(n_������, 0) <> 0 Then
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + Nvl(n_������, 0)
        Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ
        Returning Nvl(���, 0) + n_����ֵ Into n_����ֵ;
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ����
            (�տ�Ա, ���㷽ʽ, ����, ���)
          Values
            (����Ա����_In, v_���㷽ʽ, 1, Nvl(n_������, 0));
          n_����ֵ := Nvl(n_����ֵ, 0) + Nvl(n_������, 0);
        End If;
      End If;
    
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  If ����������_In Is Not Null Then
    v_�������� := ����������_In || '||';
    While v_�������� Is Not Null Loop
      --�����Id|�Ƿ����ѿ�|������|����|��ע||...
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      n_�����id := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_���ѿ�   := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      v_����     := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') + 1);
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      v_����ժҪ := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
    
      If n_���ѿ� = 1 Then
        Select ���㷽ʽ, ���� Into v_���㷽ʽ, v_������ From ���ѿ����Ŀ¼ Where ��� = n_�����id;
        If v_���㷽ʽ Is Null Then
          v_Err_Msg := v_������ || 'δ���ý��㷽ʽ����,�������ѿ��н�������,����ʧ�ܣ�';
          Raise Err_Item;
        End If;
      Else
        Select ���㷽ʽ, ���� Into v_���㷽ʽ, v_������ From ҽ�ƿ���� Where ID = n_�����id;
        If v_���㷽ʽ Is Null Then
          v_Err_Msg := v_������ || 'δ���ý��㷽ʽ����,����ҽ�ƿ������н�������,����ʧ�ܣ�';
          Raise Err_Item;
        End If;
      End If;
    
      If Nvl(n_������, 0) <> 0 Then
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + Nvl(n_������, 0)
        Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ
        Returning Nvl(���, 0) + n_����ֵ Into n_����ֵ;
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ����
            (�տ�Ա, ���㷽ʽ, ����, ���)
          Values
            (����Ա����_In, v_���㷽ʽ, 1, Nvl(n_������, 0));
          n_����ֵ := Nvl(n_����ֵ, 0) + Nvl(n_������, 0);
        End If;
      End If;
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  If Nvl(n_����ֵ, 0) = 0 Then
    Delete From ��Ա�ɿ���� Where ���� = 1 And �տ�Ա = ����Ա����_In And Nvl(���, 0) = 0;
  End If;

  --ҩƷ���ַǷ�����Ϣ���޸�
  --ҩƷδ����¼(����ѷ�ҩ���޸Ĳ���),���뷢ҩʱ�޿ⷿID
  --���ܴ��ڲ��Ϻ�ҩƷ�ⷿ��ͬ���������޷�ҩ����
  Update δ��ҩƷ��¼
  Set ����id = Decode(����id_In, 0, Null, ����id_In), ���� = ����_In, �Է�����id = ��������id_In, ���շ� = 1, �������� = v_Date
  Where ���� = 24 And NO = No_In And
        Nvl(�ⷿid, 0) In (Select Distinct Nvl(ִ�в���id, 0)
                         From ������ü�¼
                         Where ��¼���� = 1 And ��¼״̬ = 1 And NO = No_In And �շ���� = '4');

  Update δ��ҩƷ��¼
  Set ����id = Decode(����id_In, 0, Null, ����id_In), ���� = ����_In, �Է�����id = ��������id_In, ���շ� = 1, �������� = v_Date
  Where ���� = 8 And NO = No_In And
        Nvl(�ⷿid, 0) In (Select Distinct Nvl(ִ�в���id, 0)
                         From ������ü�¼
                         Where ��¼���� = 1 And ��¼״̬ = 1 And NO = No_In And �շ���� In ('5', '6', '7'));

  --ҩƷ�շ���¼(�����Ѿ���ҩ��ȡ����ҩ,���м�¼����)
  Update ҩƷ�շ���¼
  Set �Է�����id = ��������id_In, �������� = Decode(Sign(Nvl(�������, v_Date) - v_Date), -1, ��������, v_Date)
  Where ���� = 24 And NO = No_In And
        ����id + 0 In (Select ID From ������ü�¼ Where ��¼���� = 1 And ��¼״̬ = 1 And NO = No_In And �շ���� = '4');

  -------------------------------------------------------------------------------------------
  --����������
  n_Count := Null;
  Begin
    Select Count(*), Max(a.No)
    Into n_Count, v_����no
    From ҩƷ�շ���¼ A, ������ü�¼ B
    Where a.����id = b.Id And b.�շ���� = '4' And b.��¼���� = 1 And b.��¼״̬ = 1 And
          Instr(',8,9,10,21,24,25,26,', ',' || a.���� || ',') > 0 And b.No = No_In And Rownum <= 1;
  Exception
    When Others Then
      Null;
  End;
  If Nvl(n_Count, 0) > 0 Then
    If Nvl(���˿���id_In, 0) <> 0 Then
      Select ���� Into v_�������� From ���ű� Where ID = ���˿���id_In;
    End If;
    v_Err_Msg := LPad(' ', 4);
    v_Err_Msg := Substr('��������:' || ����_In || v_Err_Msg || '�Ա�:' || �Ա�_In || v_Err_Msg || '����' || ����_In || v_Err_Msg ||
                        '�����:' || Nvl(v_��ʶ��, '') || v_Err_Msg || '���˿���:' || v_��������, 1, 100);
  
    Update ҩƷ�շ���¼
    Set �Է�����id = ��������id_In, �������� = Decode(Sign(Nvl(�������, v_Date) - v_Date), -1, ��������, v_Date), ժҪ = v_Err_Msg
    Where ���� = 21 And NO = v_����no And
          ����id + 0 In (Select ID From ������ü�¼ Where ��¼���� = 1 And ��¼״̬ = 1 And NO = No_In And �շ���� = '4');
  End If;

  Update ҩƷ�շ���¼
  Set �Է�����id = ��������id_In, �������� = Decode(Sign(Nvl(�������, v_Date) - v_Date), -1, ��������, v_Date)
  Where ���� = 8 And NO = No_In And
        ����id + 0 In
        (Select ID From ������ü�¼ Where ��¼���� = 1 And ��¼״̬ = 1 And NO = No_In And �շ���� In ('5', '6', '7'));
  If Not ��ҩ����_In Is Null Then
    --���·�ҩ����
    If Nvl(���շ�_In, 0) <> 0 Then
      Update ������ü�¼
      Set ��ҩ���� = ��ҩ����_In
      Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 1 And �շ���� = 'Z';
    Else
      For v_���� In (Select To_Number(C1) As C1, C2 From Table(f_Str2List2(��ҩ����_In, ';', '|'))) Loop
        Update ������ü�¼
        Set ��ҩ���� = Nvl(v_����.C2, ��ҩ����)
        Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 1 And ִ�в���id = Nvl(v_����.C1, ִ�в���id) And �շ���� In ('5', '6', '7');
      
        Update ҩƷ�շ���¼
        Set ��ҩ���� = Nvl(v_����.C2, ��ҩ����)
        Where ���� = 8 And NO = No_In And �ⷿid = Nvl(v_����.C1, �ⷿid) And
              ����id + 0 In (Select ID
                           From ������ü�¼
                           Where ��¼���� = 1 And ��¼״̬ = 1 And NO = No_In And �շ���� In ('5', '6', '7'));
      
        Update δ��ҩƷ��¼
        Set ��ҩ���� = Nvl(v_����.C2, ��ҩ����)
        Where ���� = 8 And NO = No_In And �ⷿid = Nvl(v_����.C1, �ⷿid) And
              Nvl(�ⷿid, 0) In (Select Distinct Nvl(ִ�в���id, 0)
                               From ������ü�¼
                               Where ��¼���� = 1 And ��¼״̬ = 1 And NO = No_In And �շ���� In ('5', '6', '7'));
      End Loop;
    End If;
  End If;

  --���²��ݲ�����Ϣ
  If ����id_In Is Not Null Then
    If ���ʽ_In Is Not Null And ������Դ_In = 1 Then
      Select ���� Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = ���ʽ_In;
    End If;
  
    --ͨ�����۵��շ�ʱ������ķѱ�,��Ϊ���ò������
  
    Select Zl_Fun_Checkidentify(0, ����id_In, v_Strtmpbefor) Into v_Strtmpbefor From Dual;
    Update ������Ϣ
    Set �Ա� = Nvl(�Ա�_In, �Ա�), ���� = Nvl(����_In, ����), ���� = Decode(����, '�²���', ����_In, ����), ҽ�Ƹ��ʽ = Nvl(v_���ʽ, ҽ�Ƹ��ʽ)
    Where ����id = ����id_In;
    Select Zl_Fun_Checkidentify(1, ����id_In, v_Strtmpbefor) Into v_Msg From Dual;
    Select zl_To_Number(Nvl(zl_GetSysParameter('�Զ���������', '1111'), '0')) Into n_�²���ģʽ From Dual;
    If n_�²���ģʽ = 1 Then
    
      Update ���˹Һż�¼
      Set ���� = ����_In, �Ա� = �Ա�_In, ���� = ����_In
      Where ����id = ����id_In And ���� = '�²���';
    
      Update ������ü�¼
      Set ���� = ����_In, �Ա� = �Ա�_In, ���� = ����_In, ���ʽ = ���ʽ_In
      Where ����id = ����id_In And ���� = '�²���';
    End If;
  End If;
  --ҽ������
  --����_In    Integer:=0, --0:����;1-סԺ
  --����_In    Integer:=1, --1-�շѵ�;2-���ʵ�
  --����_In    Integer:=0, --0:ɾ�����۵�;1-�շѻ����;2-�˷ѻ�����
  --No_In      ������ü�¼.No%Type,
  --ҽ��ids_In varchar2 := Null
  Zl_ҽ������_�Ʒ�״̬_Update(0, 1, 1, No_In);

  --������Ϣ
  Begin
    Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
      Using 4, ����id_In;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����շѼ�¼_Insert;
/
Create Or Replace Procedure Zl_�����շѽ���_Modify
(
  ��������_In      Number,
  ����id_In        ������ü�¼.����id%Type,
  ����id_In        ����Ԥ����¼.����id%Type,
  ���㷽ʽ_In      Varchar2,
  ��Ԥ��_In        ����Ԥ����¼.��Ԥ��%Type := Null,
  ��֧Ʊ��_In      ����Ԥ����¼.��Ԥ��%Type := Null,
  �����id_In      ����Ԥ����¼.�����id%Type := Null,
  ����_In          ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In      ����Ԥ����¼.����˵��%Type := Null,
  �ɿ�_In          ����Ԥ����¼.�ɿ�%Type := Null,
  �Ҳ�_In          ����Ԥ����¼.�Ҳ�%Type := Null,
  �����_In      ������ü�¼.ʵ�ս��%Type := Null,
  ��ɽ���_In      Number := 0,
  ȱʡ���㷽ʽ_In  ���㷽ʽ.����%Type := Null,
  ��Ԥ������ids_In Varchar2 := Null,
  ���½������_In  Number := 1,
  ��������id_In    ����Ԥ����¼.��������id%Type := Null,
  ɾ��ԭ����_In    Number := 0,
  У�Ա�־_In      ����Ԥ����¼.У�Ա�־%Type := 0,
  �����ƻỰ_In    ����Ԥ����¼.�Ự��%Type := 0,
  �Ƿ����Ʊ��_In  ����Ԥ����¼.�Ƿ����Ʊ��%Type := 0
) As
  ------------------------------------------------------------------------------------------------------------------------------ 
  --����:�շѽ���ʱ,�޸Ľ���������Ϣ 
  --��������_In: 
  --   0-��ͨ�շѷ�ʽ: 
  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������. 
  --     ����֧Ʊ��_In:����漰��֧Ʊ,���뱾�ε���֧Ʊ��,�������շ�ʱ,������ 
  --   1.����������: 
  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ" 
  --     ����֧Ʊ��_In:������ 
  --     �ۿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ���� 
  --   2-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���) 
  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
  --     ����֧Ʊ��_In:������
  --   3-���ѿ�����:
  --     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."  ���ѿ�ID:Ϊ��ʱ,���ݿ����Զ���λ 
  --     ����֧Ʊ��_In:������ 
  --   4-���������㣬���ֽ��㷽ʽ: 
  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ|���ݺ�|�Ƿ���ͨ����|����" 
  --     ����֧Ʊ��_In:������ 
  --     �ۿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ���� 

  -- ��Ԥ��_In: ���ڳ�Ԥ��ʱ,���� 
  -- �����_In:��������ʱ,���� 
  -- ��ɽ���_In:1-����շ�;0-δ����շ� 
  -- ��Ԥ������ids_In:����ö��ַ���,��Ԥ��ʱ��Ч(��Ԥ�������ҵ���������һ��),��Ҫ��ʹ�ü�����Ԥ���� 
  -- ���½������_In  �Ƿ������Ա��������Ҫ�Ǵ���ͳһ����Ա��¼��̨����������� 
  -- ��������id_In ��������_In Ϊ1,4ʱ���봫�� 
  -- ɾ��ԭ����_in ��������_InΪ4ʱ��Ч��������㷽ʽʱ���ö�θù��� 
  -- У�Ա�־_In  ��������_InΪ4ʱ��Ч 
  -- �Ƿ����Ʊ��_In �Ƿ�ʹ�õ���Ʊ�ݣ���ɽ���_In=1 ʱ����
  ------------------------------------------------------------------------------------------------------------------------------ 
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_�������� Varchar2(500);
  v_��ǰ���� Varchar2(50);
  v_����     ����ҽ�ƿ���Ϣ.����%Type;
  n_���ѿ�id ���ѿ���Ϣ.Id%Type;
  v_����     ���ѿ����Ŀ¼.����%Type;
  n_�����id ����Ԥ����¼.���㿨���%Type;
  n_Ԥ��id   ����Ԥ����¼.Id%Type;

  v_���㷽ʽ ����Ԥ����¼.���㷽ʽ%Type;
  n_������ ����Ԥ����¼.��Ԥ��%Type;
  v_������� ����Ԥ����¼.�������%Type;
  v_����ժҪ ����Ԥ����¼.ժҪ%Type;
  v_������Ա ����Ԥ����¼.������Ա%Type;

  n_����ֵ   ��Ա�ɿ����.���%Type;
  n_��Ԥ��   ����Ԥ����¼.��Ԥ��%Type;
  v_��֧Ʊ   ����Ԥ����¼.���㷽ʽ%Type;
  v_����   ���㷽ʽ.����%Type;
  n_Count    Number;
  n_Havenull Number;
  l_Ԥ��id   t_NumList := t_NumList();
  n_�Ự��   ����Ԥ����¼.�Ự��%Type; --��ʽ��SID+'_'+SERIAL# 

  Cursor c_Feedata Is
    Select Max(m.����id) As ����id, Max(m.�Ǽ�ʱ��) As �Ǽ�ʱ��, Max(m.����Ա���) As ����Ա���, Max(m.����Ա����) As ����Ա����, Sum(���ʽ��) As ������,
           Max(m.�ɿ���id) As �ɿ���id
    From ������ü�¼ M
    Where m.����id = ����id_In;
  r_Feedata c_Feedata%RowType;

  Cursor c_Balancedata Is
    Select ��¼����, NO, ��¼״̬, ����id, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������
    From ����Ԥ����¼
    Where ����id = ����id_In And ���㷽ʽ Is Null;
  r_Balancedata c_Balancedata%RowType;

Begin
  If Nvl(�����ƻỰ_In, 0) = 0 Then
    Begin
      Select Sid || '_' || Serial# Into n_�Ự�� From V$session Where Audsid = Userenv('sessionid');
    Exception
      When Others Then
        n_�Ự�� := Null;
    End;
  End If;
  v_������Ա := zl_UserName;

  Begin
    Select ���� Into v_���� From ���㷽ʽ Where Nvl(����, 0) = 9;
  Exception
    When Others Then
      v_���� := '����';
  End;

  --0.��ʽ���� 
  Select Count(1), Max(Decode(���㷽ʽ, Null, 1, 0))
  Into n_Count, n_Havenull
  From ����Ԥ����¼
  Where ����id = ����id_In;

  --1.���ӽ��㷽ʽΪ�յĽ������� 
  n_������ := 0;
  n_Count    := 0;
  Open c_Feedata;
  Begin
    Fetch c_Feedata
      Into r_Feedata;
    --�������������㷽ʽΪnull�ļ�¼ 
    Select Nvl(Sum(��Ԥ��), 0) Into n_������ From ����Ԥ����¼ Where ����id = ����id_In;
    If Nvl(n_Havenull, 0) = 0 Or Round(Nvl(r_Feedata.������, 0), 6) <> Round(Nvl(n_������, 0), 6) Then
      --��ɾ�����ڵĽ��㷽ʽΪnull�ļ�¼ 
      Delete From ����Ԥ����¼ Where ����id = ����id_In And ���㷽ʽ Is Null;
      Select Nvl(Sum(��Ԥ��), 0) Into n_������ From ����Ԥ����¼ Where ����id = ����id_In;
    
      n_������ := Round(Nvl(r_Feedata.������, 0) - n_������, 6);
    
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־, ��������, �Ự��)
      Values
        (����Ԥ����¼_Id.Nextval, 3, Null, 1, Decode(����id_In, 0, Null, ����id_In), Null, r_Feedata.�Ǽ�ʱ��, r_Feedata.����Ա���,
         r_Feedata.����Ա����, n_������, ����id_In, r_Feedata.�ɿ���id, Sysdate, v_������Ա, -1 * ����id_In, 1, 3, n_�Ự��);
    End If;
  Exception
    When Others Then
      n_Count := 1;
  End;
  Close c_Feedata;
  If n_Count = 1 Then
    v_Err_Msg := 'δ�ҵ�ָ�����շ���ϸ����,�������ʧ�ܣ�';
    Raise Err_Item;
  End If;

  If ��������_In = 0 And Nvl(��֧Ʊ��_In, 0) <> 0 Then
    Begin
      Select b.����
      Into v_��֧Ʊ
      From ���㷽ʽӦ�� A, ���㷽ʽ B
      Where a.Ӧ�ó��� = '�շ�' And b.���� = a.���㷽ʽ And Nvl(b.Ӧ����, 0) = 1 And Rownum <= 1;
    Exception
      When Others Then
        v_��֧Ʊ := '��';
    End;
    If v_��֧Ʊ = '��' Then
      v_Err_Msg := '�ڽ��㳡����,�����ڽ�������ΪӦ����Ľ��㷽ʽ,����[���㷽ʽ]�����ã�';
      Raise Err_Item;
    End If;
  End If;

  Open c_Balancedata;
  Fetch c_Balancedata
    Into r_Balancedata;

  If Nvl(�����_In, 0) <> 0 Then
    Update ����Ԥ����¼
    Set ��Ԥ�� = Nvl(��Ԥ��, 0) + Nvl(�����_In, 0)
    Where ����id = ����id_In And ���㷽ʽ = v_����;
    If Sql%NotFound Then
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־, �����id,
         ���㿨���, ����, ������ˮ��, ����˵��, �������, ��������, �Ự��)
      Values
        (����Ԥ����¼_Id.Nextval, 3, Null, 1, r_Balancedata.����id, Null, Null, v_����, r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���,
         r_Balancedata.����Ա����, �����_In, ����id_In, r_Balancedata.�ɿ���id, Sysdate, v_������Ա, r_Balancedata.�������, 2, Null, Null,
         ����_In, ������ˮ��_In, ����˵��_In, Null, 3, n_�Ự��);
    End If;
  
    Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - Nvl(�����_In, 0) Where ����id = ����id_In And ���㷽ʽ Is Null;
  End If;

  --Ԥ����� 
  If Nvl(��Ԥ��_In, 0) <> 0 Then
    If Nvl(����id_In, 0) = 0 Then
      v_Err_Msg := '����ȷ�����˵Ĳ���ID,�շѲ���ʹ��Ԥ�������,�������ʧ�ܣ�';
      Raise Err_Item;
    End If;
  
    Zl_����Ԥ����¼_��Ԥ��(����id_In, ����id_In, ��Ԥ��_In, 1, r_Balancedata.����Ա���, r_Balancedata.����Ա����, r_Balancedata.�տ�ʱ��,
                  ��Ԥ������ids_In, 3, 1);
  End If;

  If ��������_In = 0 Then
    If Nvl(��֧Ʊ��_In, 0) <> 0 Then
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־, �����id,
         ���㿨���, ����, ������ˮ��, ����˵��, �������, ��������, �Ự��)
      Values
        (����Ԥ����¼_Id.Nextval, 3, Null, 1, r_Balancedata.����id, Null, Null, v_��֧Ʊ, r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���,
         r_Balancedata.����Ա����, ��֧Ʊ��_In, ����id_In, r_Balancedata.�ɿ���id, Sysdate, v_������Ա, r_Balancedata.�������, 2, Null, Null,
         ����_In, ������ˮ��_In, ����˵��_In, Null, 3, n_�Ự��);
    
      Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - ��֧Ʊ��_In Where ����id = ����id_In And ���㷽ʽ Is Null;
    End If;
  
    --�����շѽ��� :��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." 
    v_�������� := ���㷽ʽ_In || '||';
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      v_������� := LTrim(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_����ժҪ := LTrim(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
    
      If Nvl(n_������, 0) <> 0 Then
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־,
           �����id, ���㿨���, ����, ������ˮ��, ����˵��, �������, ��������, �Ự��)
        Values
          (����Ԥ����¼_Id.Nextval, 3, Null, 1, r_Balancedata.����id, Null, v_����ժҪ, v_���㷽ʽ, r_Balancedata.�տ�ʱ��,
           r_Balancedata.����Ա���, r_Balancedata.����Ա����, n_������, ����id_In, r_Balancedata.�ɿ���id, Sysdate, v_������Ա,
           r_Balancedata.�������, 2, Null, Null, ����_In, ������ˮ��_In, ����˵��_In, v_�������, 3, n_�Ự��);
      
        Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - n_������ Where ����id = ����id_In And ���㷽ʽ Is Null;
      End If;
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  --1.���������㽻�� 
  If ��������_In = 1 Then
    v_��ǰ���� := ���㷽ʽ_In;
    v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
    v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
    n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
    v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
    v_������� := LTrim(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
    v_����ժҪ := LTrim(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
  
    If Nvl(n_������, 0) <> 0 Then
      Select Count(1) Into n_Count From ����Ԥ����¼ Where ID = ��������id_In And Rownum < 2;
      If n_Count = 0 And Nvl(��������id_In, 0) <> 0 Then
        n_Ԥ��id := ��������id_In;
      Else
        Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      End If;
    
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־, �����id,
         ���㿨���, ����, ��������id, ������ˮ��, ����˵��, �������, ��������, �Ự��)
      Values
        (n_Ԥ��id, 3, Null, 1, r_Balancedata.����id, Null, v_����ժҪ, v_���㷽ʽ, r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���,
         r_Balancedata.����Ա����, n_������, ����id_In, r_Balancedata.�ɿ���id, Sysdate, v_������Ա, r_Balancedata.�������, 2, �����id_In,
         Null, ����_In, ��������id_In, ������ˮ��_In, ����˵��_In, v_�������, 3, n_�Ự��);
    
      Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - n_������ Where ����id = ����id_In And ���㷽ʽ Is Null;
    
      --��������������Ϣ����
      Zl_Custom_Balance_Update(n_Ԥ��id);
    End If;
  End If;

  --2.ҽ������(���ô˹���,��ȡƽ����̯�ķ�ʽ��̯�������):�������ҽ���ᴦ��,����ȫ�� 
  If ��������_In = 2 Then
    --2.1����Ƿ��Ѿ�����ҽ����������,������ɾ�� 
    n_������ := 0;
    For v_ҽ�� In (Select ID, ���㷽ʽ, ��Ԥ��
                 From ����Ԥ����¼ A
                 Where a.����id = ����id_In And Exists
                  (Select 1 From ���㷽ʽ Where a.���㷽ʽ = ���� And ���� In (3, 4)) And Mod(��¼����, 10) <> 1) Loop
      n_������ := n_������ + Nvl(v_ҽ��.��Ԥ��, 0);
      l_Ԥ��id.Extend;
      l_Ԥ��id(l_Ԥ��id.Count) := v_ҽ��.Id;
    End Loop;
  
    Update ����Ԥ����¼ Set ��Ԥ�� = Nvl(��Ԥ��, 0) + Nvl(n_������, 0) Where ����id = ����id_In And ���㷽ʽ Is Null;
  
    Forall I In 1 .. l_Ԥ��id.Count
      Delete From ����Ԥ����¼ Where ID = l_Ԥ��id(I);
  
    If ���㷽ʽ_In Is Not Null Then
      v_�������� := ���㷽ʽ_In || '||';
    End If;
  
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      n_������ := To_Number(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
    
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־, ��������,
         �Ự��)
      Values
        (����Ԥ����¼_Id.Nextval, 3, Null, 1, r_Balancedata.����id, Null, '���ս���', v_���㷽ʽ, r_Balancedata.�տ�ʱ��,
         r_Balancedata.����Ա���, r_Balancedata.����Ա����, n_������, ����id_In, r_Balancedata.�ɿ���id, Sysdate, v_������Ա,
         r_Balancedata.�������, 1, 3, n_�Ự��);
    
      --��������(���㷽ʽΪNULL��) 
      Update ����Ԥ����¼
      Set ��Ԥ�� = ��Ԥ�� - n_������
      Where ����id = ����id_In And ���㷽ʽ Is Null
      Returning Nvl(��Ԥ��, 0) Into n_����ֵ;
    
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  --3-���ѿ��������� 
  If ��������_In = 3 Then
    v_�������� := ���㷽ʽ_In || '||';
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      --�����ID|����|���ѿ�ID|���ѽ�� 
      n_�����id := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      v_����     := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_���ѿ�id := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_������ := To_Number(v_��ǰ����);
      Begin
        Select ����, ���㷽ʽ Into v_����, v_���㷽ʽ From ���ѿ����Ŀ¼ Where ��� = �����id_In;
      Exception
        When Others Then
          v_���� := Null;
      End;
      If v_���� Is Null Then
        v_Err_Msg := 'δ�ҵ���Ӧ�Ľ��㿨�ӿ�,����ˢ������ʧ��!';
        Raise Err_Item;
      End If;
      If v_���㷽ʽ Is Null Then
        v_Err_Msg := v_���� || 'δ���ö�Ӧ�Ľ��㷽ʽ,����ˢ������ʧ��!';
        Raise Err_Item;
      End If;
    
      If Nvl(n_������, 0) <> 0 Then
        Update ����Ԥ����¼
        Set ��Ԥ�� = Nvl(��Ԥ��, 0) + n_������
        Where ����id = r_Balancedata. ����id And ���㷽ʽ = v_���㷽ʽ And ���㿨��� = n_�����id
        Returning ID Into n_Ԥ��id;
        If Sql%NotFound Then
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, ���㿨���,
             У�Ա�־, ��������, �Ự��)
          Values
            (n_Ԥ��id, 3, Null, 1, r_Balancedata. ����id, Null, Null, v_���㷽ʽ, r_Balancedata. �տ�ʱ��, r_Balancedata. ����Ա���,
             r_Balancedata. ����Ա����, n_������, r_Balancedata. ����id, r_Balancedata. �ɿ���id, Sysdate, v_������Ա,
             r_Balancedata. �������, n_�����id, 2, 3, n_�Ự��);
        End If;
      
        Zl_���˿������¼_֧��(n_�����id, v_����, n_���ѿ�id, n_������, n_Ԥ��id, r_Balancedata. ����Ա���, r_Balancedata. ����Ա����,
                      r_Balancedata. �տ�ʱ��);
      
        --��������(���㷽ʽΪNULL��) 
        Update ����Ԥ����¼
        Set ��Ԥ�� = ��Ԥ�� - n_������
        Where ����id = r_Balancedata. ����id And ���㷽ʽ Is Null And Nvl(У�Ա�־, 0) = 1
        Returning Nvl(��Ԥ��, 0) Into n_����ֵ;
      End If;
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  --4.���������㣬���ֽ��㷽ʽ 
  If ��������_In = 4 Then
    If Nvl(ɾ��ԭ����_In, 0) = 1 Then
      --1.1����Ƿ��Ѿ�������������������,������ɾ�� 
      n_������ := 0;
      For c_���� In (Select ID, ���㷽ʽ, ��Ԥ��
                   From ����Ԥ����¼ A
                   Where ����id = ����id_In And �����id = �����id_In And ��������id = ��������id_In) Loop
        n_������ := n_������ + Nvl(c_����.��Ԥ��, 0);
        l_Ԥ��id.Extend;
        l_Ԥ��id(l_Ԥ��id.Count) := c_����.Id;
      End Loop;
    
      Update ����Ԥ����¼
      Set ��Ԥ�� = Nvl(��Ԥ��, 0) + Nvl(n_������, 0)
      Where ����id = ����id_In And ���㷽ʽ Is Null;
    
      Forall I In 1 .. l_Ԥ��id.Count
        Delete From ����Ԥ����¼ Where ID = l_Ԥ��id(I);
    End If;
  
    n_Ԥ��id := 0;
    --��ʽ�����㷽ʽ|������|�������|����ժҪ|���ݺ�|�Ƿ���ͨ����|���� 
    For c_���� In (Select Max(Decode(���, 1, ֵ, Null)) As ���㷽ʽ, zl_To_Number(Max(Decode(���, 2, ֵ, ''))) As ������,
                        Trim(Max(Decode(���, 3, ֵ, ''))) As �������, Trim(Max(Decode(���, 4, ֵ, ''))) As ����ժҪ,
                        Trim(Max(Decode(���, 5, ֵ, ''))) As ���ݺ�, zl_To_Number(Max(Decode(���, 6, ֵ, ''))) As �Ƿ���ͨ����,
                        Trim(Max(Decode(���, 7, ֵ, ''))) As ����
                 From (Select Rownum As ���, Column_Value As ֵ From Table(f_Str2List(���㷽ʽ_In, '|')))
                 Having Nvl(zl_To_Number(Max(Decode(���, 2, ֵ, ''))), 0) <> 0) Loop
    
      Update ����Ԥ����¼
      Set ��Ԥ�� = ��Ԥ�� + c_����.������
      Where ����id = ����id_In And ���㷽ʽ = c_����.���㷽ʽ And ��������id = ��������id_In
      Returning ID Into n_Ԥ��id;
      If Sql%NotFound Then
        Select Count(1) Into n_Count From ����Ԥ����¼ Where ID = ��������id_In And Rownum < 2;
        If n_Count = 0 Then
          n_Ԥ��id := ��������id_In;
        Else
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
        End If;
      
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־,
           �����id, ����, ��������id, ������ˮ��, ����˵��, �������, ��������, �Ự��)
        Values
          (n_Ԥ��id, 3, Null, 1, r_Balancedata.����id, Null, c_����.����ժҪ, c_����.���㷽ʽ, r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���,
           r_Balancedata.����Ա����, c_����.������, ����id_In, r_Balancedata.�ɿ���id, Sysdate, v_������Ա, r_Balancedata.�������, У�Ա�־_In,
           Decode(c_����.�Ƿ���ͨ����, 1, Null, �����id_In), Decode(c_����.�Ƿ���ͨ����, 1, Null, Nvl(c_����.����, ����_In)), ��������id_In,
           ������ˮ��_In, ����˵��_In, c_����.�������, 3, n_�Ự��);
      End If;
    
      Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - c_����.������ Where ����id = ����id_In And ���㷽ʽ Is Null;
    
      If c_����.���ݺ� Is Not Null Then
        Insert Into ҽ��������ϸ
          (����id, NO, ���㷽ʽ, ���, �����id, ��������id, ������ˮ��, ����˵��)
        Values
          (����id_In, c_����.���ݺ�, c_����.���㷽ʽ, c_����.������, Decode(c_����.�Ƿ���ͨ����, 1, Null, �����id_In), ��������id_In, ������ˮ��_In,
           ����˵��_In);
      End If;
    
      --��������������Ϣ����
      Zl_Custom_Balance_Update(n_Ԥ��id);
    End Loop;
  End If;

  If Nvl(��ɽ���_In, 0) = 0 Then
    Return;
  End If;

  ----------------------------------------------------------------------------------------- 
  --����շ�,��Ҫ������Ա�ɿ����,Ԥ����¼(���㷽ʽ=NULL) 

  --1.ɾ�����㷽ʽΪNULL��Ԥ����¼ 
  Delete ����Ԥ����¼ Where ����id = ����id_In And ���㷽ʽ Is Null And Nvl(��Ԥ��, 0) = 0;
  If Sql%NotFound Then
    Select Count(*) Into n_Count From ����Ԥ����¼ A Where ����id = ����id_In And ���㷽ʽ Is Null;
    If n_Count <> 0 Then
      v_Err_Msg := '������δ�ɿ������,������ɽ���!';
    Else
      v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�!';
    End If;
    Raise Err_Item;
  End If;

  --���������ü�¼�벡��Ԥ����¼�Ľ���Ƿ���� 
  n_������ := 0;
  n_��Ԥ��   := 0;
  Select Nvl(Sum(ʵ�ս��), 0) Into n_������ From ������ü�¼ Where ����id = ����id_In;
  Select Nvl(Sum(��Ԥ��), 0) Into n_��Ԥ�� From ����Ԥ����¼ Where ����id = ����id_In;
  If n_������ <> n_��Ԥ�� Then
    v_Err_Msg := '������Ϣ����ʵ�ս��(' || n_������ || ')�������(' || n_��Ԥ�� || ')��һ�£�������ɽ��㣡';
    Raise Err_Item;
  End If;

  --������Ϊ��ʱ������һ�����Ϊ0�Ĳ���Ԥ����¼ 
  Select Count(1) Into n_Count From ����Ԥ����¼ A Where ����id = ����id_In;
  If n_Count = 0 Then
    v_���㷽ʽ := ȱʡ���㷽ʽ_In;
    If v_���㷽ʽ Is Null Then
      Begin
        Select ���㷽ʽ Into v_���㷽ʽ From ���㷽ʽӦ�� Where Ӧ�ó��� = '�շ�' And Nvl(ȱʡ��־, 0) = 1;
      Exception
        When Others Then
          v_���㷽ʽ := Null;
      End;
      If v_���㷽ʽ Is Null Then
        Begin
          Select ���� Into v_���㷽ʽ From ���㷽ʽ Where Nvl(����, 0) = 1;
        Exception
          When Others Then
            v_���㷽ʽ := '�ֽ�';
        End;
      End If;
    End If;
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־, �����id,
       ���㿨���, ����, ������ˮ��, ����˵��, �������, ��������, �Ự��)
    Values
      (����Ԥ����¼_Id.Nextval, 3, Null, 1, r_Balancedata.����id, Null, Null, v_���㷽ʽ, r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���,
       r_Balancedata.����Ա����, 0, ����id_In, r_Balancedata.�ɿ���id, Sysdate, v_������Ա, r_Balancedata.�������, 2, Null, Null, Null,
       Null, ����˵��_In, Null, 3, n_�Ự��);
  End If;

  --2.����ɿ����ݺ��Ҳ����ݼ�У�Ա�־����Ϊ0���Ự�Ÿ���ΪNULL 
  Update ����Ԥ����¼
  Set �ɿ� = �ɿ�_In, �Ҳ� = �Ҳ�_In, У�Ա�־ = 0, �Ự�� = Null, �Ƿ����Ʊ�� = �Ƿ����Ʊ��_In
  Where ����id = ����id_In And ��¼���� <> 1;

  --3.���·���״̬ 
  Update ������ü�¼ Set ����״̬ = 0 Where ����id = ����id_In;

  --4.������Ա�ɿ����� 
  If Nvl(���½������_In, 1) = 1 Then
    For c_�ɿ� In (Select ���㷽ʽ, ����Ա����, Sum(Nvl(a.��Ԥ��, 0)) As ��Ԥ��
                 From ����Ԥ����¼ A
                 Where a.����id = ����id_In And Mod(a.��¼����, 10) <> 1
                 Group By ���㷽ʽ, ����Ա����) Loop
    
      Update ��Ա�ɿ����
      Set ��� = Nvl(���, 0) + Nvl(c_�ɿ�.��Ԥ��, 0)
      Where �տ�Ա = c_�ɿ�.����Ա���� And ���� = 1 And ���㷽ʽ = c_�ɿ�.���㷽ʽ;
      If Sql%RowCount = 0 Then
        Insert Into ��Ա�ɿ����
          (�տ�Ա, ���㷽ʽ, ����, ���)
        Values
          (c_�ɿ�.����Ա����, c_�ɿ�.���㷽ʽ, 1, Nvl(c_�ɿ�.��Ԥ��, 0));
      End If;
    End Loop;
  End If;

  --5.���ҵ�����ݴ��� 
  Zl_�����շѼ�¼_����շ�(����id_In);

  --��Ϣ���ɴ��� 
  --��������:1-�շѽ��㣬2-������� 
  --����ID:����id 
  b_Message.Zlhis_Charge_002(1, ����id_In);

  --�շѺ�������� 
  Begin
    Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
      Using 4, ����id_In;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����շѽ���_Modify;
/
Create Or Replace Procedure Zl_�����˷ѽ���_Modify
(
  ��������_In      Number,
  ����id_In        ������ü�¼.����id%Type,
  ����id_In        ����Ԥ����¼.����id%Type,
  ���㷽ʽ_In      Varchar2,
  ��Ԥ��_In        ����Ԥ����¼.��Ԥ��%Type := Null,
  �����id_In      ����Ԥ����¼.�����id%Type := Null,
  ����_In          ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In      ����Ԥ����¼.����˵��%Type := Null,
  �ɿ�_In          ����Ԥ����¼.�ɿ�%Type := Null,
  �Ҳ�_In          ����Ԥ����¼.�Ҳ�%Type := Null,
  �����_In      ������ü�¼.ʵ�ս��%Type := Null,
  ����˷�_In      Number := 0,
  ԭ����id_In      ����Ԥ����¼.����id%Type := Null,
  ʣ��תԤ��_In    Number := 0,
  ȱʡ���㷽ʽ_In  ���㷽ʽ.����%Type := Null,
  ��Ԥ������ids_In Varchar2 := Null,
  ��������id_In    ����Ԥ����¼.��������id%Type := Null,
  ɾ��ԭ����_In    Number := 0,
  У�Ա�־_In      ����Ԥ����¼.У�Ա�־%Type := 0
) As
  ------------------------------------------------------------------------------------------------------------------------------ 
  --����:�շѽ���ʱ,�޸Ľ���������Ϣ 
  --��������_In: 
  --   0-ԭ���� 
  --      ԭ������һ��ȫ��,����У�Ա�־��Ϊ1,ҽ�����óɹ���,����Ϊ2,��ɺ���0 
  --   1-��ͨ�˷ѷ�ʽ: 
  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������. 
  --   2.�������˷ѽ���: 
  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ" 
  --     �ڿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ���� 
  --   3-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���) 
  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
  --   4-���ѿ�����:
  --     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||." 
  --   5.�������˷ѽ��㣬���ֽ��㷽ʽ: 
  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ|���ݺ�|�Ƿ���ͨ����" 
  --     �ڿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ���� 

  -- ��Ԥ��_In:����漰Ԥ����,���뱾�ε���Ԥ�� ������<0ʱ ��ʾ��Ԥ������ֵ;>0 ʱ:��ʾ��Ԥ���� 
  -- ʣ��תԤ��_In: 1��ʾ��ʣ���˿��ת��Ϊ��ֵ���;0��ʾ��Ԥ�� 
  -- �����_In:��������ʱ,���� 
  -- ����˷�_In:0-δ����˷�;1-�쳣����˷�;2-����˷� 
  -- ԭ����ID_IN:ԭ����ʱ,����(���ԭ����δ����ʱ,�������һ�ν���Ϊ׼) 
  -- ��Ԥ������ids_In:����ö��ַ���,��Ԥ��ʱ��Ч(��Ԥ�������ҵ���������һ��),��Ҫ��ʹ�ü�����Ԥ���� 
  -- ��������id_In ��������_In Ϊ3,5ʱ���봫�� 
  -- ɾ��ԭ����_in ��������_InΪ5ʱ��Ч��������㷽ʽʱ���ö�θù��� 
  -- У�Ա�־_In  ��������_InΪ5ʱ��Ч
  ------------------------------------------------------------------------------------------------------------------------------ 
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_�������� Varchar2(500);
  v_��ǰ���� Varchar2(50);
  v_����     ����ҽ�ƿ���Ϣ.����%Type;
  n_���ѿ�id ���ѿ���Ϣ.Id%Type;
  v_����     ���ѿ����Ŀ¼.����%Type;
  n_�����id ����Ԥ����¼.���㿨���%Type;
  n_ԭԤ��id ����Ԥ����¼.Id%Type;
  n_Ԥ��id   ����Ԥ����¼.Id%Type;
  v_���㷽ʽ ����Ԥ����¼.���㷽ʽ%Type;
  n_������ ����Ԥ����¼.��Ԥ��%Type;
  n_����ֵ   ��Ա�ɿ����.���%Type;
  n_Ԥ����� ����Ԥ����¼.��Ԥ��%Type;
  n_��Ԥ��   ����Ԥ����¼.��Ԥ��%Type;
  v_������� ����Ԥ����¼.�������%Type;
  v_����ժҪ ����Ԥ����¼.ժҪ%Type;
  v_����   ���㷽ʽ.����%Type;
  n_��¼״̬ ����Ԥ����¼.��¼״̬%Type;
  n_��ֵid   ����Ԥ����¼.Id%Type;

  v_�˷ѽ��� ���㷽ʽ.����%Type;
  v_No       ����Ԥ����¼.No%Type;
  n_Dec      Number; --���С��λ�� 

  n_Count    Number;
  n_Havenull Number;
  l_Ԥ��id   t_NumList := t_NumList();
  n_ԭ����id ����Ԥ����¼.����id%Type;
  n_�ؽ�id   ����Ԥ����¼.����id%Type;
  n_����id   ����Ԥ����¼.����id%Type;
  n_������� ����Ԥ����¼.����id%Type;
  v_Msg      Varchar2(500);
  n_�Ự��   ����Ԥ����¼.�Ự��%Type; --��ʽ��SID+'_'+SERIAL# 
  v_������Ա ����Ԥ����¼.������Ա%Type;
  n_�첽���� Number;

  n_����תסԺ�˷� Number;
  n_�Ƿ����Ʊ��   ����Ԥ����¼.�Ƿ����Ʊ��%Type;

  Cursor c_Feedata Is
    Select Max(NO) As NO, Max(m.����id) As ����id, Max(m.�Ǽ�ʱ��) As �Ǽ�ʱ��, Max(m.����Ա���) As ����Ա���, Max(m.����Ա����) As ����Ա����,
           Sum(���ʽ��) As ������, Max(m.�ɿ���id) As �ɿ���id
    From ������ü�¼ M
    Where m.����id = ����id_In;
  r_Feedata c_Feedata%RowType;

  Cursor c_Balancedata Is
    Select ��¼����, NO, ��¼״̬, ����id, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������
    From ����Ԥ����¼
    Where ����id = ����id_In And ���㷽ʽ Is Null;
  r_Balancedata c_Balancedata%RowType;

Begin
  Begin
    Select Sid || '_' || Serial# Into n_�Ự�� From V$session Where Audsid = Userenv('sessionid');
  Exception
    When Others Then
      n_�Ự�� := Null;
  End;
  v_������Ա := zl_UserName;

  Begin
    Select ���� Into v_�˷ѽ��� From ���㷽ʽ Where ���� = 1;
  Exception
    When Others Then
      v_�˷ѽ��� := '�ֽ�';
  End;
  Select Count(1) Into n_�첽���� From ���ý������ A Where a.�����־ = 1 And a.����id = ����id_In And Rownum < 2;

  --���С��λ�� 
  Select zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  --0.��ʽ���� 
  Select Count(1), Max(Decode(���㷽ʽ, Null, 1, 0)), Max(�������)
  Into n_Count, n_Havenull, n_�������
  From ����Ԥ����¼
  Where ����id = ����id_In;

  If Nvl(n_Count, 0) = 0 Or Nvl(�����_In, 0) <> 0 Then
    --���ӽ��㷽ʽΪNULL�ļ�¼ 
    Begin
      Select ���� Into v_���� From ���㷽ʽ Where Nvl(����, 0) = 9;
    Exception
      When Others Then
        v_���� := '����';
    End;
  End If;

  --1.���ӽ��㷽ʽΪ�յĽ������� 
  n_Count := 0;
  Open c_Feedata;
  Begin
    Fetch c_Feedata
      Into r_Feedata;
    If Nvl(n_Havenull, 0) = 0 Then
      n_������ := Round(Nvl(r_Feedata.������, 0), n_Dec);
    
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־, ��������, �Ự��)
      Values
        (����Ԥ����¼_Id.Nextval, 3, Null, 2, Decode(����id_In, 0, Null, ����id_In), Null, r_Feedata.�Ǽ�ʱ��, r_Feedata.����Ա���,
         r_Feedata.����Ա����, n_������, ����id_In, r_Feedata.�ɿ���id, Sysdate, v_������Ա, -1 * ����id_In, 1, 3, n_�Ự��);
    
      --����(�Ȼ��ܺ��������� 
      If n_������ <> Nvl(r_Feedata.������, 0) Then
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־, ��������, �Ự��)
        Values
          (����Ԥ����¼_Id.Nextval, 3, Null, 1, Decode(����id_In, 0, Null, ����id_In), v_����, r_Feedata.�Ǽ�ʱ��, r_Feedata.����Ա���,
           r_Feedata.����Ա����, Nvl(r_Feedata.������, 0) - n_������, ����id_In, r_Feedata.�ɿ���id, Sysdate, v_������Ա, -1 * ����id_In, 1,
           3, n_�Ự��);
      End If;
      n_������� := -1 * ����id_In;
    End If;
  Exception
    When Others Then
      n_Count := 1;
  End;
  Close c_Feedata;
  If n_Count = 1 Then
    v_Err_Msg := 'δ�ҵ�ָ�����շ���ϸ����,�������ʧ�ܣ�';
    Raise Err_Item;
  End If;

  Open c_Balancedata;
  Fetch c_Balancedata
    Into r_Balancedata;

  n_ԭ����id := ԭ����id_In;
  If Nvl(n_ԭ����id, 0) = 0 Then
    Select Max(b.����id)
    Into n_ԭ����id
    From ������ü�¼ A, ������ü�¼ B
    Where a.����id = ����id_In And a.No = b.No And b.��¼���� = 1 And b.��¼״̬ In (1, 3);
  End If;

  If Nvl(n_ԭ����id, 0) = 0 Then
    v_Err_Msg := 'δ�ҵ�ԭ��������,����ԭ���ˣ�';
    Raise Err_Item;
  End If;

  If ��������_In = 0 Then
    --0.ԭ���� 
    --1.ֻ�������ѿ����� 
    Select Count(1)
    Into n_Count
    From ����Ԥ����¼ A, ���˿������¼ B
    Where a.Id = b.����id And a.��¼���� = 3 And a.����id = n_ԭ����id And Rownum < 2;
    If n_Count <> 0 Then
      Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
    
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��,
         ������Ա, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������, �Ự��)
        Select n_Ԥ��id, ��¼����, NO, 2, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, r_Balancedata. �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�,
               r_Balancedata.����Ա���, r_Balancedata.����Ա����, -1 * ��Ԥ��, ����id_In, r_Balancedata.�ɿ���id, Sysdate, v_������Ա, Ԥ�����,
               �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 2, r_Balancedata.�������, Mod(��¼����, 10), n_�Ự��
        
        From ����Ԥ����¼ A
        Where a.��¼���� = 3 And a.����id = n_ԭ����id And Exists (Select 1 From ���˿������¼ Where ����id = a.Id);
    
      --�շ�ʱ����ʹ���˶������ѿ� 
      For c_��¼ In (Select a.Id, c.�ӿڱ��, c.���ѿ�id, c.����, -1 * Sum(c.Ӧ�ս��) As ������
                   From ����Ԥ����¼ A, ���˿������¼ C
                   Where a.Id = c.����id And a.��¼���� = 3 And a.��¼״̬ In (1, 3) And a.����id = n_ԭ����id
                   Group By a.Id, c.�ӿڱ��, c.���ѿ�id, c.����) Loop
      
        Zl_���˿������¼_�˿�(c_��¼.�ӿڱ��, c_��¼.����, c_��¼.���ѿ�id, c_��¼.������, c_��¼.Id, n_Ԥ��id, r_Balancedata. ����Ա���,
                      r_Balancedata. ����Ա����, r_Balancedata. �տ�ʱ��);
      End Loop;
    End If;
  
    --2.ҽ�� 
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ��������id, ����ʱ��, ������Ա, Ԥ�����,
       �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������, �Ự��)
      Select ����Ԥ����¼_Id.Nextval, ��¼����, NO, 2, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���,
             r_Balancedata.����Ա����, -1 * ��Ԥ��, r_Balancedata.����id, r_Balancedata.�ɿ���id, ��������id, Sysdate, v_������Ա, Ԥ�����,
             �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 1 As У�Ա�־, r_Balancedata.�������, Mod(��¼����, 10), n_�Ự��
      From ����Ԥ����¼ A, (Select ���� From ���㷽ʽ Where ���� In (3, 4)) J
      Where a.��¼״̬ In (1, 3) And a.���㷽ʽ = j.���� And a.���㷽ʽ Is Not Null And a.����id = n_ԭ����id And a.�����id Is Null;
  
    --���½��㷽ʽΪNULL �ļ�¼ 
    Select Sum(��Ԥ��) Into n_����ֵ From ����Ԥ����¼ Where ����id = ����id_In And ���㷽ʽ Is Not Null;
    Select Sum(���ʽ��) Into n_������ From ������ü�¼ Where ����id = ����id_In;
    Update ����Ԥ����¼
    Set ��Ԥ�� = Nvl(n_������, 0) - Nvl(n_����ֵ, 0)
    Where ����id = ����id_In And ���㷽ʽ Is Null;
  End If;

  n_�ؽ�id := 0;
  If ��������_In <> 0 Then
    --����ȫ��ʱ,����Ƿ�����������շ����ݵ� 
    Begin
      Select ����id Into n_�ؽ�id From ����Ԥ����¼ Where ������� = n_������� And ����id <> ����id_In And Rownum < 2;
    Exception
      When Others Then
        n_�ؽ�id := 0;
    End;
  End If;

  --��Ҫ��������� 
  If Nvl(�����_In, 0) <> 0 Then
    --���ѷ������յĽ����¼�� 
    n_����id   := ����id_In;
    n_��¼״̬ := 2;
    If Nvl(n_�ؽ�id, 0) <> 0 Then
      n_����id   := n_�ؽ�id;
      n_��¼״̬ := 1;
    End If;
  
    Update ����Ԥ����¼
    Set ��Ԥ�� = Nvl(��Ԥ��, 0) + Nvl(�����_In, 0)
    Where ����id = n_����id And ���㷽ʽ = v_����;
    If Sql%NotFound Then
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־, �����id,
         ���㿨���, ����, ������ˮ��, ����˵��, �������, ��������, �Ự��)
      Values
        (����Ԥ����¼_Id.Nextval, 3, Null, n_��¼״̬, r_Balancedata.����id, Null, Null, v_����, r_Balancedata.�տ�ʱ��,
         r_Balancedata.����Ա���, r_Balancedata.����Ա����, �����_In, n_����id, r_Balancedata.�ɿ���id, Sysdate, v_������Ա,
         r_Balancedata.�������, 2, Null, Null, ����_In, ������ˮ��_In, ����˵��_In, Null, 3, n_�Ự��);
    End If;
  
    Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - Nvl(�����_In, 0) Where ����id = n_����id And ���㷽ʽ Is Null;
  End If;

  --Ԥ�����:����ǳ�Ԥ��,��Ҫ�ȴ����Ԥ���� 
  If Nvl(��Ԥ��_In, 0) <> 0 Then
    If Nvl(r_Balancedata.����id, 0) = 0 Then
      v_Err_Msg := '����ȷ��������Ϣ,����ʹ��Ԥ������㣡';
      Raise Err_Item;
    End If;
  
    n_Ԥ����� := ��Ԥ��_In;
    If n_Ԥ����� < 0 And Nvl(ʣ��תԤ��_In, 0) = 1 Then
      --     �ڳ�Ԥ��_In:����漰Ԥ����,���뱾�ε���Ԥ�� ������<0ʱ ��ʾ��Ԥ������ֵ;>0 ʱ:��ʾ��Ԥ���� 
      --     ��ʣ��תԤ��_In: 1��ʾ��ʣ���˿��ת��Ϊ��ֵ���;0��ʾ��Ԥ�� 
    
      --1.�����ɳ�ֵԤ�� 
      v_No := Nextno(11);
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ���, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־, �����id,
         ���㿨���, ����, ������ˮ��, ����˵��, �������, Ԥ�����, ��������, �Ự��)
      Values
        (����Ԥ����¼_Id.Nextval, 1, v_No, 1, r_Balancedata.����id, Null, '�˷�����Ԥ��', v_�˷ѽ���, r_Balancedata.�տ�ʱ��,
         r_Balancedata.����Ա���, r_Balancedata.����Ա����, -1 * n_Ԥ�����, ����id_In, r_Balancedata.�ɿ���id, Sysdate, v_������Ա,
         r_Balancedata.�������, 0, Null, Null, ����_In, ������ˮ��_In, ����˵��_In, Null, 1, Null, n_�Ự��);
    
      --Ԥ��������� 
      Insert Into Ԥ���������
        (Ԥ��id, ����id, Ԥ�����, Ԥ�����)
      Values
        (����Ԥ����¼_Id.Currval, r_Balancedata.����id, 1, -1 * n_Ԥ�����);
    
      --���²������ 
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) + (-1 * n_Ԥ�����)
      Where ����id = ����id_In And ���� = 1 And ���� = 1
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ������� (����id, ����, Ԥ�����, ����) Values (����id_In, 1, -1 * n_Ԥ�����, 1);
        n_����ֵ := -1 * n_Ԥ�����;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From �������
        Where ����id = r_Balancedata.����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
    
      --2.�����˷Ѽ�¼ 
      If Nvl(n_�ؽ�id, 0) <> 0 Then
        Select Sum(��Ԥ��)
        Into n_����ֵ
        From ����Ԥ����¼
        Where ����id = ����id_In And ���㷽ʽ Is Null And Nvl(��Ԥ��, 0) <> 0;
        If Nvl(n_����ֵ, 0) <> 0 Then
          Insert Into ����Ԥ����¼
            (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־,
             �����id, ���㿨���, ����, ������ˮ��, ����˵��, �������, ��������, �Ự��)
          Values
            (����Ԥ����¼_Id.Nextval, 3, Null, 2, r_Balancedata.����id, Null, v_����ժҪ, v_�˷ѽ���, r_Balancedata.�տ�ʱ��,
             r_Balancedata.����Ա���, r_Balancedata.����Ա����, n_����ֵ, ����id_In, r_Balancedata.�ɿ���id, Sysdate, v_������Ա,
             r_Balancedata.�������, 2, Null, Null, ����_In, ������ˮ��_In, ����˵��_In, Null, 3, n_�Ự��);
        
          Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - Nvl(n_����ֵ, 0) Where ����id = ����id_In And ���㷽ʽ Is Null;
        End If;
        n_������ := n_������ - Nvl(n_����ֵ, 0);
      
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־,
           �����id, ���㿨���, ����, ������ˮ��, ����˵��, �������, ��������, �Ự��)
        Values
          (����Ԥ����¼_Id.Nextval, 3, Null, 1, r_Balancedata.����id, Null, v_����ժҪ, v_�˷ѽ���, r_Balancedata.�տ�ʱ��,
           r_Balancedata.����Ա���, r_Balancedata.����Ա����, n_������, n_�ؽ�id, r_Balancedata.�ɿ���id, Sysdate, v_������Ա,
           r_Balancedata.�������, 2, Null, Null, ����_In, ������ˮ��_In, ����˵��_In, Null, 3, n_�Ự��);
      
        Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� + n_Ԥ����� Where ����id = ����id_In And ���㷽ʽ Is Null;
      Else
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־,
           �����id, ���㿨���, ����, ������ˮ��, ����˵��, �������, ��������, �Ự��)
        Values
          (����Ԥ����¼_Id.Nextval, 3, Null, 2, r_Balancedata.����id, Null, v_����ժҪ, v_���㷽ʽ, r_Balancedata.�տ�ʱ��,
           r_Balancedata.����Ա���, r_Balancedata.����Ա����, -1 * n_Ԥ�����, ����id_In, r_Balancedata.�ɿ���id, Sysdate, v_������Ա,
           r_Balancedata.�������, 2, Null, Null, ����_In, ������ˮ��_In, ����˵��_In, Null, 3, n_�Ự��);
      
        Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� + n_Ԥ����� Where ����id = ����id_In And ���㷽ʽ Is Null;
      End If;
    End If;
  
    If Nvl(n_Ԥ�����, 0) < 0 And Nvl(ʣ��תԤ��_In, 0) = 0 Then
      If Nvl(n_�ؽ�id, 0) <> 0 Then
        --1.��Ԥ���� 
        For v_��Ԥ�� In (Select Max(a.Id) As ID, Max(a.No) As NO, a.����id, Max(a.�տ�ʱ��) As �տ�ʱ��, Sum(Nvl(a.��Ԥ��, 0)) As ���
                      From ����Ԥ����¼ A,
                           (Select Distinct a.����id
                             From ������ü�¼ A, ������ü�¼ B
                             Where a.No = b.No And Mod(a.��¼����, 10) = 1 And b.����id = n_ԭ����id) B
                      Where a.����id = b.����id And Mod(a.��¼����, 10) = 1 And Nvl(a.Ԥ�����, 0) = 1
                      Group By NO, ����id
                      Having Sum(Nvl(a.��Ԥ��, 0)) > 0
                      Order By �տ�ʱ�� Desc) Loop
          Insert Into ����Ԥ����¼
            (ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������,
             У�Ա�־, �����id, ���㿨���, ����, ������ˮ��, ����˵��, �������, Ԥ�����, ��������, �Ự��, ��������id)
            Select ����Ԥ����¼_Id.Nextval, 11, NO, ʵ��Ʊ��, ��¼״̬, ����id, Null, ժҪ, ���㷽ʽ, r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���,
                   r_Balancedata.����Ա����, -1 * v_��Ԥ��.���, ����id_In, r_Balancedata.�ɿ���id, Sysdate, v_������Ա, r_Balancedata.�������,
                   2, �����id, ���㿨���, ����, ������ˮ��, ����˵��, �������, Ԥ�����, 3, n_�Ự��, ��������id
            From ����Ԥ����¼
            Where ID = v_��Ԥ��.Id;
        
          --����Ԥ��������� 
          Select Max(ID) Into n_��ֵid From ����Ԥ����¼ Where NO = v_��Ԥ��.No And ��¼���� = 1 And ��¼״̬ <> 2;
          If Nvl(n_��ֵid, 0) <> 0 Then
            Update Ԥ���������
            Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(v_��Ԥ��.���, 0)
            Where ����id = v_��Ԥ��.����id And Ԥ��id = n_��ֵid
            Returning Ԥ����� Into n_����ֵ;
            If Sql%RowCount = 0 Then
              Insert Into Ԥ���������
                (Ԥ��id, ����id, Ԥ�����, Ԥ�����)
              Values
                (n_��ֵid, v_��Ԥ��.����id, 1, Nvl(v_��Ԥ��.���, 0));
              n_����ֵ := Nvl(v_��Ԥ��.���, 0);
            End If;
            If Nvl(n_����ֵ, 0) = 0 Then
              Delete From Ԥ��������� Where Ԥ��id = n_��ֵid And Nvl(Ԥ�����, 0) = 0;
            End If;
          End If;
        
          --���²������ 
          Update �������
          Set Ԥ����� = Nvl(Ԥ�����, 0) + v_��Ԥ��.���
          Where ����id = v_��Ԥ��.����id And ���� = 1 And ���� = 1
          Returning Ԥ����� Into n_����ֵ;
          If Sql%RowCount = 0 Then
            Insert Into ������� (����id, ����, Ԥ�����, ����) Values (v_��Ԥ��.����id, 1, v_��Ԥ��.���, 1);
            n_����ֵ := v_��Ԥ��.���;
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From �������
            Where ����id = v_��Ԥ��.����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
          End If;
        
          Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - (-1 * v_��Ԥ��.���) Where ����id = ����id_In And ���㷽ʽ Is Null;
        
          n_Ԥ����� := n_Ԥ����� - (-1 * v_��Ԥ��.���);
        End Loop;
      
        --2.��Ԥ���� 
        If n_Ԥ����� <> 0 Then
          For v_��Ԥ�� In (Select Max(a.Id) As ID, a.No, a.����id, a.Ԥ�����, Max(a.�տ�ʱ��) As �տ�ʱ��, Sum(Nvl(a.��Ԥ��, 0)) As ���
                        From ����Ԥ����¼ A,
                             (Select Distinct a.����id
                               From ������ü�¼ A, ������ü�¼ B
                               Where a.No = b.No And Mod(a.��¼����, 10) = 1 And b.����id = n_ԭ����id) B
                        Where a.����id = b.����id And Mod(a.��¼����, 10) = 1 And Nvl(a.Ԥ�����, 0) = 1 And a.����id <> ����id_In
                        Group By a.No, a.����id, a.Ԥ�����
                        Having Sum(Nvl(a.��Ԥ��, 0)) > 0
                        Order By �տ�ʱ�� Desc) Loop
          
            If v_��Ԥ��.��� - n_Ԥ����� < 0 Then
              n_������ := v_��Ԥ��.���;
              n_Ԥ����� := n_Ԥ����� - v_��Ԥ��.���;
            Else
              n_������ := n_Ԥ�����;
              n_Ԥ����� := 0;
            End If;
          
            Insert Into ����Ԥ����¼
              (ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������,
               У�Ա�־, �����id, ���㿨���, ����, ������ˮ��, ����˵��, �������, Ԥ�����, ��������, �Ự��, ��������id)
              Select ����Ԥ����¼_Id.Nextval, 11, NO, ʵ��Ʊ��, ��¼״̬, ����id, Null, ժҪ, ���㷽ʽ, r_Balancedata.�տ�ʱ��,
                     r_Balancedata.����Ա���, r_Balancedata.����Ա����, n_������, n_�ؽ�id, r_Balancedata.�ɿ���id, Sysdate, v_������Ա,
                     r_Balancedata.�������, 2, �����id, ���㿨���, ����, ������ˮ��, ����˵��, �������, Ԥ�����, 3, n_�Ự��, ��������id
              From ����Ԥ����¼
              Where ID = v_��Ԥ��.Id;
          
            --����Ԥ��������� 
            Select Max(ID) Into n_��ֵid From ����Ԥ����¼ Where NO = v_��Ԥ��.No And ��¼���� = 1 And ��¼״̬ <> 2;
            If Nvl(n_��ֵid, 0) <> 0 Then
              Update Ԥ���������
              Set Ԥ����� = Nvl(Ԥ�����, 0) + (-1 * n_������)
              Where ����id = v_��Ԥ��.����id And Ԥ��id = n_��ֵid
              Returning Ԥ����� Into n_����ֵ;
              If Sql%RowCount = 0 Then
                Insert Into Ԥ���������
                  (Ԥ��id, ����id, Ԥ�����, Ԥ�����)
                Values
                  (n_��ֵid, v_��Ԥ��.����id, Nvl(v_��Ԥ��.Ԥ�����, 2), -1 * n_������);
                n_����ֵ := -1 * n_������;
              End If;
              If Nvl(n_����ֵ, 0) = 0 Then
                Delete From Ԥ��������� Where Ԥ��id = n_��ֵid And Nvl(Ԥ�����, 0) = 0;
              End If;
            End If;
          
            --���²������ 
            Update �������
            Set Ԥ����� = Nvl(Ԥ�����, 0) + (-1 * n_������)
            Where ����id = v_��Ԥ��.����id And ���� = 1 And ���� = 1
            Returning Ԥ����� Into n_����ֵ;
            If Sql%RowCount = 0 Then
              Insert Into ������� (����id, ����, Ԥ�����, ����) Values (v_��Ԥ��.����id, 1, -1 * n_������, 1);
              n_����ֵ := -1 * n_������;
            End If;
            If Nvl(n_����ֵ, 0) = 0 Then
              Delete From �������
              Where ����id = v_��Ԥ��.����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
            End If;
          
            Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - n_������ Where ����id = n_�ؽ�id And ���㷽ʽ Is Null;
          
            n_����ֵ := 1;
            If n_Ԥ����� = 0 Then
              Exit;
            End If;
          End Loop;
        End If;
      Else
        --��Ԥ���� 
        n_����ֵ   := 0;
        n_Ԥ����� := -1 * n_Ԥ�����;
      
        For v_��Ԥ�� In (Select Max(a.Id) As ID, a.No, a.����id, a.Ԥ�����, Max(a.�տ�ʱ��) As �տ�ʱ��, Sum(Nvl(a.��Ԥ��, 0)) As ���
                      From ����Ԥ����¼ A,
                           (Select Distinct a.����id
                             From ������ü�¼ A, ������ü�¼ B
                             Where a.No = b.No And Mod(a.��¼����, 10) = 1 And b.����id = n_ԭ����id) B
                      Where a.����id = b.����id And Mod(a.��¼����, 10) = 1 And Nvl(a.Ԥ�����, 0) = 1
                      Group By a.No, a.����id, a.Ԥ�����
                      Having Sum(Nvl(a.��Ԥ��, 0)) > 0
                      Order By �տ�ʱ�� Desc) Loop
        
          If v_��Ԥ��.��� - n_Ԥ����� < 0 Then
            n_������ := -1 * v_��Ԥ��.���;
            n_Ԥ����� := n_Ԥ����� - v_��Ԥ��.���;
          Else
            n_������ := -1 * n_Ԥ�����;
            n_Ԥ����� := 0;
          End If;
        
          Insert Into ����Ԥ����¼
            (ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������,
             У�Ա�־, �����id, ���㿨���, ����, ������ˮ��, ����˵��, �������, Ԥ�����, ��������, �Ự��, ��������id)
            Select ����Ԥ����¼_Id.Nextval, 11, NO, ʵ��Ʊ��, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���,
                   r_Balancedata.����Ա����, n_������, ����id_In, r_Balancedata.�ɿ���id, Sysdate, v_������Ա, r_Balancedata.�������, 2,
                   �����id, ���㿨���, ����, ������ˮ��, ����˵��, �������, Ԥ�����, 3, n_�Ự��, ��������id
            From ����Ԥ����¼
            Where ID = v_��Ԥ��.Id;
        
          --����Ԥ��������� 
          Select Max(ID) Into n_��ֵid From ����Ԥ����¼ Where NO = v_��Ԥ��.No And ��¼���� = 1 And ��¼״̬ <> 2;
          If Nvl(n_��ֵid, 0) <> 0 Then
            Update Ԥ���������
            Set Ԥ����� = Nvl(Ԥ�����, 0) + (-1 * n_������)
            Where ����id = v_��Ԥ��.����id And Ԥ��id = n_��ֵid
            Returning Ԥ����� Into n_����ֵ;
            If Sql%RowCount = 0 Then
              Insert Into Ԥ���������
                (Ԥ��id, ����id, Ԥ�����, Ԥ�����)
              Values
                (n_��ֵid, v_��Ԥ��.����id, Nvl(v_��Ԥ��.Ԥ�����, 2), -1 * n_������);
              n_����ֵ := -1 * n_������;
            End If;
            If Nvl(n_����ֵ, 0) = 0 Then
              Delete From Ԥ��������� Where Ԥ��id = n_��ֵid And Nvl(Ԥ�����, 0) = 0;
            End If;
          End If;
        
          --���²������ 
          Update �������
          Set Ԥ����� = Nvl(Ԥ�����, 0) + (-1 * n_������)
          Where ����id = v_��Ԥ��.����id And ���� = 1 And ���� = 1
          Returning Ԥ����� Into n_����ֵ;
          If Sql%RowCount = 0 Then
            Insert Into ������� (����id, ����, Ԥ�����, ����) Values (v_��Ԥ��.����id, 1, -1 * n_������, 1);
            n_����ֵ := -1 * n_������;
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From �������
            Where ����id = v_��Ԥ��.����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
          End If;
        
          Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - n_������ Where ����id = ����id_In And ���㷽ʽ Is Null;
        
          n_����ֵ := 1;
          If n_Ԥ����� = 0 Then
            Exit;
          End If;
        End Loop;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        v_Err_Msg := 'δ�ҵ�ԭʼ�ĳ�Ԥ����¼,���ܻ���Ԥ���';
        Raise Err_Item;
      End If;
    
      If Nvl(n_Ԥ�����, 0) <> 0 Then
        v_Err_Msg := '��ǰ��Ԥ���������շѽ����еĳ�Ԥ����,���ܻ���Ԥ���';
        Raise Err_Item;
      End If;
    End If;
  
    n_Ԥ����� := ��Ԥ��_In;
    If Nvl(n_Ԥ�����, 0) > 0 Then
      --��Ԥ���� 
      Zl_����Ԥ����¼_��Ԥ��(����id_In, ����id_In, n_Ԥ�����, 1, r_Balancedata.����Ա���, r_Balancedata.����Ա����, r_Balancedata.�տ�ʱ��,
                    ��Ԥ������ids_In, 3, 1);
    End If;
  End If;

  --1-��ͨ�˷ѷ�ʽ 
  If ��������_In = 1 Then
    --�����շѽ��� :��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." 
    v_�������� := ���㷽ʽ_In || '||';
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      v_���㷽ʽ := Nvl(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1), ȱʡ���㷽ʽ_In);
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      v_������� := LTrim(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_����ժҪ := LTrim(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
    
      If Nvl(n_������, 0) <> 0 Then
        n_������ := Nvl(n_������, 0);
        If Nvl(n_�ؽ�id, 0) <> 0 Then
          --1.�Ȱ����ַ�ʽȫ�� 
          --2.�ٰ����ַ�ʽ�տ� 
          --3.�����˿�=1+2 
          Select Sum(��Ԥ��)
          Into n_����ֵ
          From ����Ԥ����¼
          Where ����id = ����id_In And ���㷽ʽ Is Null And Nvl(��Ԥ��, 0) <> 0;
          If Nvl(n_����ֵ, 0) <> 0 Then
            Insert Into ����Ԥ����¼
              (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־,
               �����id, ���㿨���, ����, ������ˮ��, ����˵��, �������, ��������, �Ự��)
            Values
              (����Ԥ����¼_Id.Nextval, 3, Null, 2, r_Balancedata.����id, Null, v_����ժҪ, v_���㷽ʽ, r_Balancedata.�տ�ʱ��,
               r_Balancedata.����Ա���, r_Balancedata.����Ա����, n_����ֵ, ����id_In, r_Balancedata.�ɿ���id, Sysdate, v_������Ա,
               r_Balancedata.�������, 2, Null, Null, ����_In, ������ˮ��_In, ����˵��_In, Null, 3, n_�Ự��);
          
            Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - Nvl(n_����ֵ, 0) Where ����id = ����id_In And ���㷽ʽ Is Null;
          End If;
          n_������ := n_������ - Nvl(n_����ֵ, 0);
        
          If Nvl(n_������, 0) <> 0 Then
            Insert Into ����Ԥ����¼
              (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־,
               �����id, ���㿨���, ����, ������ˮ��, ����˵��, �������, ��������, �Ự��)
            Values
              (����Ԥ����¼_Id.Nextval, 3, Null, 1, r_Balancedata.����id, Null, v_����ժҪ, v_���㷽ʽ, r_Balancedata.�տ�ʱ��,
               r_Balancedata.����Ա���, r_Balancedata.����Ա����, n_������, n_�ؽ�id, r_Balancedata.�ɿ���id, Sysdate, v_������Ա,
               r_Balancedata.�������, 2, Null, Null, ����_In, ������ˮ��_In, ����˵��_In, Null, 3, n_�Ự��);
          
            Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - n_������ Where ����id = n_�ؽ�id And ���㷽ʽ Is Null;
          End If;
        Else
          If Nvl(n_������, 0) <> 0 Then
            Insert Into ����Ԥ����¼
              (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־,
               �����id, ���㿨���, ����, ������ˮ��, ����˵��, �������, ��������, �Ự��)
            Values
              (����Ԥ����¼_Id.Nextval, 3, Null, 2, r_Balancedata.����id, Null, v_����ժҪ, v_���㷽ʽ, r_Balancedata.�տ�ʱ��,
               r_Balancedata.����Ա���, r_Balancedata.����Ա����, n_������, ����id_In, r_Balancedata.�ɿ���id, Sysdate, v_������Ա,
               r_Balancedata.�������, 2, Null, Null, ����_In, ������ˮ��_In, ����˵��_In, Null, 3, n_�Ự��);
          
            Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - n_������ Where ����id = ����id_In And ���㷽ʽ Is Null;
          End If;
        End If;
      End If;
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  --2.�������˷ѽ��� 
  If ��������_In = 2 Then
    v_��ǰ���� := ���㷽ʽ_In;
    v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
    v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
    n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
    v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
    v_������� := LTrim(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
    v_����ժҪ := LTrim(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
  
    If Nvl(n_������, 0) <> 0 Then
      If Nvl(n_�ؽ�id, 0) <> 0 Then
        --1.�Ȱ����ַ�ʽȫ�� 
        --2.�ٰ����ַ�ʽ�տ� 
        --3.�����˿�=1+2 
        Select Sum(��Ԥ��)
        Into n_����ֵ
        From ����Ԥ����¼
        Where ����id = ����id_In And ���㷽ʽ Is Null And Nvl(��Ԥ��, 0) <> 0;
        If Nvl(n_����ֵ, 0) <> 0 Then
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־,
             �����id, ���㿨���, ����, ������ˮ��, ����˵��, �������, ��������, �Ự��, ��������id)
          Values
            (n_Ԥ��id, 3, Null, 2, r_Balancedata.����id, Null, v_����ժҪ, v_���㷽ʽ, r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���,
             r_Balancedata.����Ա����, n_����ֵ, ����id_In, r_Balancedata.�ɿ���id, Sysdate, v_������Ա, r_Balancedata.�������, 2, �����id_In,
             Null, ����_In, ������ˮ��_In, ����˵��_In, Null, 3, n_�Ự��, ��������id_In);
        
          Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - Nvl(n_����ֵ, 0) Where ����id = ����id_In And ���㷽ʽ Is Null;
        
          --��������������Ϣ����
          Zl_Custom_Balance_Update(n_Ԥ��id);
        End If;
        n_������ := n_������ - Nvl(n_����ֵ, 0);
      
        Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־,
           �����id, ���㿨���, ����, ������ˮ��, ����˵��, �������, ��������, �Ự��, ��������id)
        Values
          (n_Ԥ��id, 3, Null, 1, r_Balancedata.����id, Null, v_����ժҪ, v_���㷽ʽ, r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���,
           r_Balancedata.����Ա����, n_������, n_�ؽ�id, r_Balancedata.�ɿ���id, Sysdate, v_������Ա, r_Balancedata.�������, 2, �����id_In,
           Null, ����_In, ������ˮ��_In, ����˵��_In, Null, 3, n_�Ự��, ��������id_In);
      
        Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - n_������ Where ����id = n_�ؽ�id And ���㷽ʽ Is Null;
      
        --��������������Ϣ����
        Zl_Custom_Balance_Update(n_Ԥ��id);
      Else
        Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־,
           �����id, ���㿨���, ����, ������ˮ��, ����˵��, �������, ��������, �Ự��, ��������id)
        Values
          (n_Ԥ��id, 3, Null, 2, r_Balancedata.����id, Null, v_����ժҪ, v_���㷽ʽ, r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���,
           r_Balancedata.����Ա����, n_������, ����id_In, r_Balancedata.�ɿ���id, Sysdate, v_������Ա, r_Balancedata.�������, 2, �����id_In,
           Null, ����_In, ������ˮ��_In, ����˵��_In, v_�������, 3, n_�Ự��, ��������id_In);
      
        Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - n_������ Where ����id = ����id_In And ���㷽ʽ Is Null;
      
        --��������������Ϣ����
        Zl_Custom_Balance_Update(n_Ԥ��id);
      End If;
    End If;
  End If;

  --3-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���) 
  If ��������_In = 3 Then
    --3.1����Ƿ��Ѿ�����ҽ����������,������ɾ�� 
    n_������ := 0;
    For v_ҽ�� In (Select ID, ���㷽ʽ, ��Ԥ��
                 From ����Ԥ����¼ A
                 Where ����id = ����id_In And Exists
                  (Select 1 From ���㷽ʽ Where a.���㷽ʽ = ���� And ���� In (3, 4)) And Mod(��¼����, 10) <> 1) Loop
      n_������ := n_������ + Nvl(v_ҽ��.��Ԥ��, 0);
      l_Ԥ��id.Extend;
      l_Ԥ��id(l_Ԥ��id.Count) := v_ҽ��.Id;
    End Loop;
  
    Update ����Ԥ����¼ Set ��Ԥ�� = Nvl(��Ԥ��, 0) + Nvl(n_������, 0) Where ����id = ����id_In And ���㷽ʽ Is Null;
  
    Forall I In 1 .. l_Ԥ��id.Count
      Delete ����Ԥ����¼ Where ID = l_Ԥ��id(I);
  
    If ���㷽ʽ_In Is Not Null Then
      v_�������� := ���㷽ʽ_In || '||';
    End If;
  
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      n_������ := To_Number(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
    
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־, ��������,
         �Ự��)
      Values
        (����Ԥ����¼_Id.Nextval, 3, Null, 2, r_Balancedata.����id, Null, '���ս���', v_���㷽ʽ, r_Balancedata.�տ�ʱ��,
         r_Balancedata.����Ա���, r_Balancedata.����Ա����, n_������, ����id_In, r_Balancedata.�ɿ���id, Sysdate, v_������Ա,
         r_Balancedata.�������, 1, 3, n_�Ự��);
    
      --��������(���㷽ʽΪNULL��) 
      Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - n_������ Where ����id = ����id_In And ���㷽ʽ Is Null;
    
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  --4-���ѿ��������� 
  If ��������_In = 4 Then
    Begin
      --��ȡ��һ���տ����ID 
      Select Max(a.����id)
      Into n_ԭ����id
      From ������ü�¼ A, (Select Distinct NO From ������ü�¼ Where ����id = n_ԭ����id) M
      Where a.No = m.No And Mod(a.��¼����, 10) = 1 And a.��¼״̬ In (1, 3) And Nvl(a.����״̬, 0) <> 1 And
            a.�Ǽ�ʱ�� + 0 =
            (Select Max(m.�Ǽ�ʱ��)
             From ������ü�¼ M, (Select Distinct NO From ������ü�¼ Where ����id = n_ԭ����id) J
             Where m.No = j.No And Mod(m.��¼����, 10) = 1 And m.��¼״̬ In (1, 3) And Nvl(m.����״̬, 0) <> 1);
    
    Exception
      When Others Then
        v_Err_Msg := 'δ�ҵ�ԭ�������ݣ�';
        Raise Err_Item;
    End;
  
    v_�������� := ���㷽ʽ_In || '||';
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      --�����ID|����|���ѿ�ID|���ѽ�� 
      n_�����id := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      v_����     := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_���ѿ�id := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_������ := To_Number(v_��ǰ����);
    
      Begin
        Select ����, ���㷽ʽ Into v_����, v_���㷽ʽ From ���ѿ����Ŀ¼ Where ��� = n_�����id;
      Exception
        When Others Then
          v_���� := Null;
      End;
      If v_���� Is Null Then
        v_Err_Msg := 'δ�ҵ���Ӧ�Ľ��㿨�ӿ�,����ˢ������ʧ��!';
        Raise Err_Item;
      End If;
      If v_���㷽ʽ Is Null Then
        v_Err_Msg := v_���� || 'δ���ö�Ӧ�Ľ��㷽ʽ,����ˢ������ʧ��!';
        Raise Err_Item;
      End If;
      If Nvl(n_������, 0) <> 0 Then
        n_����id := ����id_In;
      
        If Nvl(n_�ؽ�id, 0) <> 0 Then
          For c_��¼ In (Select a.Id, c.�ӿڱ��, c.���ѿ�id, c.����, c.Ӧ�ս�� As ������
                       From ����Ԥ����¼ A, ���˿������¼ C
                       Where a.Id = c.����id And a.��¼���� = 3 And a.��¼״̬ In (1, 3) And a.����id = n_ԭ����id And c.�ӿڱ�� = n_�����id And
                             c.���ѿ�id = n_���ѿ�id) Loop
          
            If Nvl(c_��¼.������, 0) <> 0 Then
              Update ����Ԥ����¼
              Set ��Ԥ�� = Nvl(��Ԥ��, 0) + c_��¼.������
              Where ����id = ����id_In And ���㷽ʽ = v_���㷽ʽ And ���㿨��� = n_�����id
              Returning ID Into n_Ԥ��id;
              If Sql%NotFound Then
                Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
                Insert Into ����Ԥ����¼
                  (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������,
                   ���㿨���, У�Ա�־, ��������, �Ự��)
                Values
                  (n_Ԥ��id, 3, Null, 2, r_Balancedata. ����id, Null, Null, v_���㷽ʽ, r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���,
                   r_Balancedata. ����Ա����, c_��¼.������, ����id_In, r_Balancedata. �ɿ���id, Sysdate, v_������Ա, r_Balancedata.�������,
                   n_�����id, 2, 3, n_�Ự��);
              End If;
            
              Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - c_��¼.������ Where ����id = ����id_In And ���㷽ʽ Is Null;
            
              --���뿨�����¼ 
              Zl_���˿������¼_�˿�(c_��¼.�ӿڱ��, c_��¼.����, c_��¼.���ѿ�id, -1 * c_��¼.������, c_��¼.Id, n_Ԥ��id, r_Balancedata. ����Ա���,
                            r_Balancedata. ����Ա����, r_Balancedata. �տ�ʱ��);
            
              n_������ := n_������ - c_��¼.������;
            End If;
          End Loop;
          n_����id := n_�ؽ�id;
        End If;
      
        Update ����Ԥ����¼
        Set ��Ԥ�� = Nvl(��Ԥ��, 0) + n_������
        Where ����id = n_����id And ���㷽ʽ = v_���㷽ʽ And ���㿨��� = n_�����id
        Returning ID Into n_Ԥ��id;
        If Sql%NotFound Then
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, ���㿨���,
             У�Ա�־, ��������, �Ự��)
          Values
            (n_Ԥ��id, 3, Null, Decode(Nvl(n_�ؽ�id, 0), 0, 2, 1), r_Balancedata. ����id, Null, Null, v_���㷽ʽ,
             r_Balancedata. �տ�ʱ��, r_Balancedata. ����Ա���, r_Balancedata. ����Ա����, n_������, n_����id, r_Balancedata. �ɿ���id,
             Sysdate, v_������Ա, r_Balancedata. �������, n_�����id, 2, 3, n_�Ự��);
        End If;
      
        If Nvl(n_�ؽ�id, 0) = 0 Then
          Begin
            Select ID Into n_ԭԤ��id From ����Ԥ����¼ A Where ����id = n_ԭ����id And ���㿨��� = n_�����id;
          Exception
            When Others Then
              v_Err_Msg := 'δ�ҵ�ԭ�����¼��';
              Raise Err_Item;
          End;
        
          Zl_���˿������¼_�˿�(n_�����id, v_����, n_���ѿ�id, -1 * n_������, n_ԭԤ��id, n_Ԥ��id, r_Balancedata. ����Ա���,
                        r_Balancedata. ����Ա����, r_Balancedata. �տ�ʱ��);
        Else
          Zl_���˿������¼_֧��(n_�����id, v_����, n_���ѿ�id, n_������, n_Ԥ��id, r_Balancedata. ����Ա���, r_Balancedata. ����Ա����,
                        r_Balancedata. �տ�ʱ��);
        End If;
      
        --��������(���㷽ʽΪNULL��) 
        Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - n_������ Where ����id = n_����id And ���㷽ʽ Is Null;
      End If;
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  --5-������ҽ������ 
  If ��������_In = 5 Then
    If Nvl(ɾ��ԭ����_In, 0) = 1 Then
      --1.1����Ƿ��Ѿ�������������������,������ɾ�� 
      n_������ := 0;
      For c_���� In (Select ID, ���㷽ʽ, ��Ԥ��
                   From ����Ԥ����¼ A
                   Where ����id = ����id_In And �����id = �����id_In And Nvl(��������id, 0) = Nvl(��������id_In, 0)) Loop
        n_������ := n_������ + Nvl(c_����.��Ԥ��, 0);
        l_Ԥ��id.Extend;
        l_Ԥ��id(l_Ԥ��id.Count) := c_����.Id;
      End Loop;
    
      Update ����Ԥ����¼
      Set ��Ԥ�� = Nvl(��Ԥ��, 0) + Nvl(n_������, 0)
      Where ����id = ����id_In And ���㷽ʽ Is Null;
    
      Forall I In 1 .. l_Ԥ��id.Count
        Delete From ����Ԥ����¼ Where ID = l_Ԥ��id(I);
    
      Delete From ҽ��������ϸ
      Where ����id = ����id_In And �����id = �����id_In And Nvl(��������id, 0) = Nvl(��������id_In, 0);
    End If;
  
    --��ʽ�����㷽ʽ|������|�������|����ժҪ|���ݺ�|�Ƿ���ͨ���� 
    For c_���� In (Select Max(Decode(���, 1, ֵ, Null)) As ���㷽ʽ, zl_To_Number(Max(Decode(���, 2, ֵ, ''))) As ������,
                        Trim(Max(Decode(���, 3, ֵ, ''))) As �������, Trim(Max(Decode(���, 4, ֵ, ''))) As ����ժҪ,
                        Trim(Max(Decode(���, 5, ֵ, ''))) As ���ݺ�, zl_To_Number(Max(Decode(���, 6, ֵ, ''))) As �Ƿ���ͨ����
                 From (Select Rownum As ���, Column_Value As ֵ From Table(f_Str2List(���㷽ʽ_In, '|')))
                 Having Nvl(zl_To_Number(Max(Decode(���, 2, ֵ, ''))), 0) <> 0) Loop
    
      Update ����Ԥ����¼
      Set ��Ԥ�� = ��Ԥ�� + c_����.������
      Where ����id = ����id_In And �����id = �����id_In And ���㷽ʽ = c_����.���㷽ʽ And Nvl(��������id, 0) = Nvl(��������id_In, 0)
      Returning ID Into n_Ԥ��id;
      If Sql%NotFound Then
        Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־,
           �����id, ����, ��������id, ������ˮ��, ����˵��, �������, ��������, �Ự��)
        Values
          (n_Ԥ��id, 3, Null, 2, r_Balancedata.����id, Null, c_����.����ժҪ, c_����.���㷽ʽ, r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���,
           r_Balancedata.����Ա����, c_����.������, ����id_In, r_Balancedata.�ɿ���id, Sysdate, v_������Ա, r_Balancedata.�������, У�Ա�־_In,
           Decode(c_����.�Ƿ���ͨ����, 1, Null, �����id_In), Decode(c_����.�Ƿ���ͨ����, 1, Null, ����_In), ��������id_In, ������ˮ��_In, ����˵��_In,
           c_����.�������, 3, n_�Ự��);
      End If;
    
      Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - c_����.������ Where ����id = ����id_In And ���㷽ʽ Is Null;
    
      If c_����.���ݺ� Is Not Null Then
        Insert Into ҽ��������ϸ
          (����id, NO, ���㷽ʽ, ���, �����id, ��������id, ������ˮ��, ����˵��)
        Values
          (����id_In, c_����.���ݺ�, c_����.���㷽ʽ, c_����.������, Decode(c_����.�Ƿ���ͨ����, 1, Null, �����id_In), ��������id_In, ������ˮ��_In,
           ����˵��_In);
      End If;
    
      --��������������Ϣ����
      Zl_Custom_Balance_Update(n_Ԥ��id);
    End Loop;
  End If;

  If Nvl(����˷�_In, 0) = 0 Then
    Return;
  End If;

  ----------------------------------------------------------------------------------------- 
  --����շ�,��Ҫ������Ա�ɿ����,Ԥ����¼(���㷽ʽ=NULL) 
  -- ����˷�_In:0-δ����˷�;1-�쳣����˷�;2-����˷� 
  If Nvl(����˷�_In, 0) = 1 Then
    --������ȫ��ԭ���˺����������� 
    Select Count(1)
    Into n_Count
    From ����Ԥ����¼ A
    Where a.����id = n_ԭ����id And Nvl(a.У�Ա�־, 0) = 2 And Nvl(a.��Ԥ��, 0) <> 0 And Not Exists
     (Select 1
           From ����Ԥ����¼
           Where ����id = ����id_In And ���㷽ʽ = a.���㷽ʽ And Nvl(��������id, 0) = Nvl(a.��������id, 0) And Nvl(У�Ա�־, 0) = 2);
    If n_Count <> 0 Then
      v_Err_Msg := '������δ���ϵĽ��ף�����������ϣ�';
      Raise Err_Item;
    End If;
  
    Update ����Ԥ����¼ Set У�Ա�־ = 0, �Ự�� = Null Where ����id = ����id_In;
  
    If Nvl(n_�첽����, 0) <> 0 Then
      --�ָ����۵�״̬ 
    
      --��δ��ҩƷ��¼���Ϊδ�շ�״̬ 
      For c_No In (Select Distinct NO From ������ü�¼ Where ��¼���� = 1 And ����id = ����id_In) Loop
        Update δ��ҩƷ��¼ A
        Set ���շ� = 0
        Where ���� In (8, 24) And NO = c_No.No And Exists
         (Select 1
               From ҩƷ�շ���¼
               Where ���� = a.���� And Nvl(�ⷿid, 0) = Nvl(a.�ⷿid, 0) And NO = c_No.No And Mod(��¼״̬, 3) = 1 And ����� Is Null);
      End Loop;
    
      Delete From ������ü�¼ Where ��¼���� = 1 And ����id = ����id_In;
    
      Update ������ü�¼
      Set ��¼״̬ = 1, ����id = Null, ���ʽ�� = Null, ����Ա��� = Null, ����Ա���� = Null, �ɿ���id = Null
      Where ��¼���� = 1 And ����id = n_ԭ����id;
    
      --���ԭԤ����¼ 
      Update ����Ԥ����¼ Set ��¼״̬ = 3 Where ����id = n_ԭ����id And Mod(��¼����, 10) <> 1;
    End If;
    Return;
  End If;

  Select Max(a.�Ƿ����Ʊ��)
  Into n_�Ƿ����Ʊ��
  From ����Ԥ����¼ A
  Where a.����id = n_ԭ����id And a.��¼���� In (11, 3);

  --1.���ҵ�����ݴ��� 
  --  �������תסԺ�˷��쳣����˷�ʱ�����������ҵ������
  Select Count(1)
  Into n_����תסԺ�˷�
  From ����Ԥ����¼
  Where ����id = ����id_In And ���㷽ʽ Is Null And Nvl(���ӱ�־, 0) = 2 And Rownum < 2;
  If Nvl(n_�첽����, 0) = 1 And Nvl(n_����תסԺ�˷�, 0) = 0 Then
    Zl_�����շѼ�¼_����˷�(����id_In, r_Balancedata.����Ա����, n_�ؽ�id, n_�Ƿ����Ʊ��);
  End If;

  --2.ɾ�����㷽ʽΪNULL��Ԥ����¼ 
  --���㷽ʽΪNULL�ĳ�����¼���ؽ��¼�Ľ��֮��Ϊ�㣬˵�������ȫ������ 
  If Nvl(n_�ؽ�id, 0) <> 0 Then
    Select Sum(Nvl(��Ԥ��, 0))
    Into n_��Ԥ��
    From ����Ԥ����¼
    Where ����id In (����id_In, n_�ؽ�id) And ���㷽ʽ Is Null;
    If Nvl(n_��Ԥ��, 0) <> 0 Then
      v_Err_Msg := '������δ�ɿ������,������ɽ���!';
      Raise Err_Item;
    Else
      Delete ����Ԥ����¼ Where ����id = ����id_In And ���㷽ʽ Is Null And Nvl(��Ԥ��, 0) = 0;
      If Sql%NotFound Then
        Update ����Ԥ����¼ Set ���㷽ʽ = v_�˷ѽ��� Where ����id = ����id_In And ���㷽ʽ Is Null;
        If Sql%NotFound Then
          v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�˷ѽ��㴰��]�������շѣ�!';
          Raise Err_Item;
        End If;
      End If;
    
      Delete ����Ԥ����¼ Where ����id = n_�ؽ�id And ���㷽ʽ Is Null And Nvl(��Ԥ��, 0) = 0;
      If Sql%NotFound Then
        Update ����Ԥ����¼ Set ���㷽ʽ = v_�˷ѽ��� Where ����id = n_�ؽ�id And ���㷽ʽ Is Null;
        If Sql%NotFound Then
          v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�˷ѽ��㴰��]�������շѣ�!';
          Raise Err_Item;
        End If;
      End If;
    End If;
    Update ����Ԥ����¼ Set �ɿ� = �ɿ�_In, �Ҳ� = �Ҳ�_In, У�Ա�־ = 0 Where ����id = n_�ؽ�id;
  Else
    Delete ����Ԥ����¼ Where ����id = ����id_In And ���㷽ʽ Is Null And Nvl(��Ԥ��, 0) = 0;
    If Sql%NotFound Then
      Select Count(1) Into n_Count From ����Ԥ����¼ A Where ����id = ����id_In And ���㷽ʽ Is Null;
      If n_Count <> 0 Then
        v_Err_Msg := '������δ�ɿ������,������ɽ���!';
      Else
        v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�˷ѽ��㴰��]�������շѣ�!';
      End If;
      Raise Err_Item;
    End If;
  End If;

  --3.���������ü�¼�벡��Ԥ����¼�Ľ���Ƿ���� 
  n_������ := 0;
  n_��Ԥ��   := 0;
  Select Nvl(Sum(ʵ�ս��), 0)
  Into n_������
  From ������ü�¼
  Where ����id In (Select ����id From ����Ԥ����¼ Where ������� = n_�������);
  Select Nvl(Sum(��Ԥ��), 0) Into n_��Ԥ�� From ����Ԥ����¼ Where ������� = n_�������;
  If n_������ <> n_��Ԥ�� Then
    v_Err_Msg := '������Ϣ����ʵ�ս��(' || n_������ || ')�������(' || n_��Ԥ�� || ')��һ�£�������ɽ��㣡';
    Raise Err_Item;
  End If;

  --4.������Ϊ��ʱ������һ�����Ϊ0�Ĳ���Ԥ����¼ 
  Select Count(1) Into n_Count From ����Ԥ����¼ A Where ����id = ����id_In;
  If n_Count = 0 Then
    v_���㷽ʽ := ȱʡ���㷽ʽ_In;
    If v_���㷽ʽ Is Null Then
      Begin
        Select ���㷽ʽ Into v_���㷽ʽ From ���㷽ʽӦ�� Where Ӧ�ó��� = '�շ�' And Nvl(ȱʡ��־, 0) = 1;
      Exception
        When Others Then
          v_���㷽ʽ := Null;
      End;
      If v_���㷽ʽ Is Null Then
        Begin
          Select ���� Into v_���㷽ʽ From ���㷽ʽ Where Nvl(����, 0) = 1;
        Exception
          When Others Then
            v_���㷽ʽ := '�ֽ�';
        End;
      End If;
    End If;
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־, �����id,
       ���㿨���, ����, ������ˮ��, ����˵��, �������, ��������, �Ự��)
    Values
      (����Ԥ����¼_Id.Nextval, 3, Null, 1, r_Balancedata.����id, Null, Null, v_���㷽ʽ, r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���,
       r_Balancedata.����Ա����, 0, ����id_In, r_Balancedata.�ɿ���id, Sysdate, v_������Ա, r_Balancedata.�������, 2, Null, Null, Null,
       Null, ����˵��_In, Null, 3, n_�Ự��);
  End If;

  --5.����ɿ����ݺ��Ҳ����ݼ�У�Ա�־����Ϊ0���Ự�Ÿ���ΪNULL  
  Update ����Ԥ����¼
  Set �ɿ� = �ɿ�_In, �Ҳ� = �Ҳ�_In, У�Ա�־ = 0, �Ự�� = Null, �Ƿ����Ʊ�� = n_�Ƿ����Ʊ��
  Where ����id In (����id_In, n_�ؽ�id);

  --6.���·���״̬ 
  Update ������ü�¼ Set ����״̬ = 0 Where ����id In (����id_In, n_�ؽ�id);

  --7.������Ա�ɿ����� 
  For c_�ɿ� In (Select ���㷽ʽ, ����Ա����, Sum(Nvl(a.��Ԥ��, 0)) As ��Ԥ��
               From ����Ԥ����¼ A
               Where a.����id In (����id_In, n_�ؽ�id) And Mod(a.��¼����, 10) <> 1
               Group By ���㷽ʽ, ����Ա����) Loop
  
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + Nvl(c_�ɿ�.��Ԥ��, 0)
    Where �տ�Ա = c_�ɿ�.����Ա���� And ���� = 1 And ���㷽ʽ = c_�ɿ�.���㷽ʽ;
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ����
        (�տ�Ա, ���㷽ʽ, ����, ���)
      Values
        (c_�ɿ�.����Ա����, c_�ɿ�.���㷽ʽ, 1, Nvl(c_�ɿ�.��Ԥ��, 0));
    End If;
  End Loop;

  --8.��Ϣ���ɴ��� 
  b_Message.Zlhis_Charge_004(1, ����id_In);
  If Nvl(n_�ؽ�id, 0) <> 0 Then
    b_Message.Zlhis_Charge_002(1, n_�ؽ�id);
  End If;

  --9.��Ϣ���� 
  Select ����id_In || ',' || ����id_In || ',' || Decode(����˷�_In, 2, 0, 0, 0, 1) Into v_Msg From Dual;
  Begin
    Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
      Using 5, v_Msg;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����˷ѽ���_Modify;
/
Create Or Replace Procedure Zl_�����շѼ�¼_����˷�
(
  ����id_In     ����Ԥ����¼.����id%Type,
  ����Ա����_In ����Ԥ����¼.����Ա����%Type,
  ���ս���id_In ����Ԥ����¼.����id%Type := Null,
  ����Ʊ��_In   ����Ԥ����¼.�Ƿ����Ʊ��%Type := 0
) As
  --���ܣ������˷���ɺ��첽����ʱ�������ҵ������
  n_ԭ����id ����Ԥ����¼.����id%Type;
  n_ʣ������ ������ü�¼.����%Type;
  n_׼������ ������ü�¼.����%Type;
  n_ִ��״̬ ������ü�¼.ִ��״̬%Type;

  v_Para         Varchar2(1000);
  n_����ģʽ     Number(3);
  n_�ֱ��ӡ     Number;
  n_Onepatiprint Number;
  n_��ӡid       Ʊ�ݴ�ӡ����.Id%Type;
  l_ʹ��id       t_NumList := t_NumList();
  n_����Ʊ��     Number;
  n_������       Number;

  n_Count Number;
  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin
  --��ԭ����ID
  Select Max(����id)
  Into n_ԭ����id
  From (With c_���� As (Select NO From ������ü�¼ Where Mod(��¼����, 10) = 1 And ����id = ����id_In)
         Select Max(a.����id) As ����id
         From ������ü�¼ A, c_���� M
         Where a.No = m.No And Mod(a.��¼����, 10) = 1 And a.��¼״̬ In (1, 3) And Nvl(a.����״̬, 0) <> 1 And
               a.�Ǽ�ʱ�� + 0 =
               (Select Max(m.�Ǽ�ʱ��)
                From ������ü�¼ M, c_���� J
                Where m.No = j.No And Mod(m.��¼����, 10) = 1 And m.��¼״̬ In (1, 3) And Nvl(m.����״̬, 0) <> 1));


  If Nvl(���ս���id_In, 0) <> 0 Then
    --��������ʱ��ԭ��¼�Ǳ�ȫ�������˵�
    Update ������ü�¼ Set ��¼״̬ = 3 Where Mod(��¼����, 10) = 1 And ��¼״̬ = 1 And ����id = n_ԭ����id;
  End If;
  --���ԭԤ����¼
  Update ����Ԥ����¼ Set ��¼״̬ = 3 Where ����id = n_ԭ����id And Mod(��¼����, 10) <> 1;

  --����˿�ʱһ��ͨ���տ��¼
  Update ����Ԥ����¼
  Set ��¼״̬ = 3
  Where ��¼���� Not In (1, 11) And ��¼״̬ = 1 And ����id <> ����id_In And
        (�����id, ��������id) In (Select �����id, ��������id
                            From ����Ԥ����¼
                            Where ��¼���� Not In (1, 11) And ����id = ����id_In And �����id Is Not Null);

  --���밴�ա��շ�ϸĿid���������򣬷�ֹ����������ҩƷ��桱��
  For c_No In (Select NO, ���
               From ������ü�¼
               Where Mod(��¼����, 10) = 1 And ����id In (����id_In, ���ս���id_In)
               Group By NO, ���, �շ�ϸĿid
               Having Nvl(Sum(Nvl(����, 1) * ����), 0) <> 0
               Order By �շ�ϸĿid) Loop
  
    For r_Bill In (Select a.Id, a.No, a.�շ����, a.ҽ�����, a.���, a.�۸񸸺�, a.�շ�ϸĿid, a.ִ��״̬, j.�������, m.��������,
                          Nvl(j.ҽ��״̬, 0) As ҽ��״̬
                   From ������ü�¼ A, ����ҽ����¼ J, �������� M
                   Where a.ҽ����� = j.Id(+) And a.�շ�ϸĿid + 0 = m.����id(+) And a.��¼���� = 1 And a.��¼״̬ In (1, 3) And
                         Nvl(a.ִ��״̬, 0) <> 1 And a.No = c_No.No And a.��� = c_No.���) Loop
    
      --ҩƷ�����������
      If Instr(',4,5,6,7,', r_Bill.�շ����) > 0 Then
        Zl_ҩƷ�շ���¼_�����˷�(r_Bill.Id);
      End If;
    
      Select Nvl(Sum(Nvl(����, 1) * ����), 0)
      Into n_ʣ������
      From ������ü�¼
      Where Mod(��¼����, 10) = 1 And NO = r_Bill.No And ��� = r_Bill.���;
    
      n_׼������ := 0;
      --׼������(��ҩƷ��ĿΪʣ������,ԭʼ����)
      If Instr(',4,5,6,7,', r_Bill.�շ����) = 0 Or (r_Bill.�շ���� = '4' And Nvl(r_Bill.��������, 0) = 0) Then
        --��ҩƷ����(�Ծ���ҽ��ִ��Ϊ׼���м��)
        --: 1.����ҽ��ִ�мƼ۵�,����ҽ��ִ��Ϊ׼(�����ܰ���:���;����;����;������Ѫ),��ִ�еĲ������˷�
        --: 2.������ҽ��ִ�мƼ۵�,����ʣ������Ϊ׼
        --: 3.ҽ�������˵�,����ʣ������Ϊ׼(����ҽ����¼.ҽ��״̬=4��ʾ����ҽ������ɾ��"����ҽ������",����ҩ�������Ϻ���ҩ)
        --: 4.����ҽ������.ִ��״̬=1�����ִ�У�ʱ��׼����Ϊ0�����ٸ���ҽ��ִ�мƼ���ͳ��׼����
        If Instr(',C,D,F,G,K,', ',' || r_Bill.������� || ',') = 0 And r_Bill.������� Is Not Null And r_Bill.ҽ��״̬ <> 4 Then
          Select Nvl(Sum(Decode(b.ִ��״̬, 1, 0, 1) * Decode(c.ִ��״̬, 0, 1, 0) * c.����), 0)
          Into n_׼������
          From ����ҽ������ B, ҽ��ִ�мƼ� C
          Where b.ҽ��id = r_Bill.ҽ����� And b.No = r_Bill.No And b.ҽ��id = c.ҽ��id And b.���ͺ� = c.���ͺ� And
                c.�շ�ϸĿid + 0 = r_Bill.�շ�ϸĿid And b.��¼���� = 1;
        End If;
      Else
        Select Nvl(Sum(Nvl(����, 1) * ʵ������), 0)
        Into n_׼������
        From ҩƷ�շ���¼
        Where ���� In (8, 24) And Mod(��¼״̬, 3) = 1 And ����� Is Null And NO = r_Bill.No And ����id = r_Bill.Id;
      End If;
    
      --���ԭ���ü�¼
      n_ִ��״̬ := Case
                  When n_ʣ������ = n_׼������ Then
                   0
                  When n_׼������ = 0 Then
                   1
                  Else
                   2
                End;
      Update ������ü�¼
      Set ִ��״̬ = n_ִ��״̬
      Where Mod(��¼����, 10) = 1 And ��¼״̬ In (1, 3) And NO = c_No.No And ��� = c_No.���;
    
      Update ������ü�¼
      Set ��¼״̬ = 3
      Where Mod(��¼����, 10) = 1 And ��¼״̬ = 1 And ����id = n_ԭ����id And NO = c_No.No And ��� = c_No.���;
    End Loop;
  End Loop;

  If Nvl(����Ʊ��_In, 0) = 0 Then
    --���ñ�־||NO;ִ�п���(����);�վݷ�Ŀ(��ҳ����,����);�շ�ϸĿ(����)
    v_Para     := Nvl(zl_GetSysParameter('Ʊ�ݷ������', 1121), '0||0;0;0,0;0');
    n_����ģʽ := zl_To_Number(Substr(v_Para, 1, 1));
    n_�ֱ��ӡ := Nvl(zl_GetSysParameter('���ŵ����շѷֱ��ӡ', 1121), '0');
  End If;

  For c_No In (Select NO
               From ������ü�¼
               Where Mod(��¼����, 10) = 1 And ����id In (����id_In, ���ս���id_In)
               Group By NO, ���
               Having Nvl(Sum(Nvl(����, 1) * ����), 0) <> 0) Loop
  
    If Nvl(����Ʊ��_In, 0) = 0 Then
      --�Ƿ񰴲���һ�δ�ӡ��
      Select Count(1)
      Into n_Onepatiprint
      From Ʊ�ݴ�ӡ���� A1, Ʊ�ݴ�ӡ���� A2
      Where A1.Id = A2.Id And A1.�������� = A2.�������� And A1.No = c_No.No And A1.�������� = 1 And Nvl(A2.��ӡ����, 0) = 1;
    
      --�˷�Ʊ�ݻ���(��ȫ��ʱ�Ż��գ�������ʱ���ش�����л���)
      If n_����ģʽ = 0 And n_�ֱ��ӡ = 1 And n_Onepatiprint = 0 Then
        Select Count(1)
        Into n_������
        From (Select 1
               From ������ü�¼ A
               Where Mod(a.��¼����, 10) = 1 And a.No = c_No.No
               Group By a.No, a.���
               Having Nvl(Sum(Nvl(a.����, 1) * a.����), 0) <> 0);
      Else
        Select Count(1)
        Into n_������
        From (Select 1
               From ������ü�¼ A
               Where a.No In
                     (Select a.No
                      From ������ü�¼ A, ������ü�¼ B
                      Where a.����id = b.����id And Mod(a.��¼����, 10) = 1 And Mod(b.��¼����, 10) = 1 And b.No = c_No.No) And
                     Mod(a.��¼����, 10) = 1
               Group By a.No, a.���
               Having Nvl(Sum(Nvl(a.����, 1) * a.����), 0) <> 0);
      End If;
    
      If n_����ģʽ <> 0 And n_Onepatiprint <> 0 And n_������ = 0 Then
        n_����Ʊ�� := 0;
      Elsif n_������ = 0 Then
        n_����Ʊ�� := 1;
      Else
        n_����Ʊ�� := 0;
      End If;
    
      If n_����Ʊ�� = 1 Then
        If n_����ģʽ <> 0 Then
          --�ջ�Ʊ��
          Select ʹ��id
          Bulk Collect
          Into l_ʹ��id
          From (Select Distinct b.ʹ��id From Ʊ�ݴ�ӡ��ϸ B Where b.No = c_No.No And Nvl(b.Ʊ��, 0) = 1);
        
          n_����ģʽ := l_ʹ��id.Count;
          If l_ʹ��id.Count <> 0 Then
            --������ռ�¼
            Forall I In 1 .. l_ʹ��id.Count
              Insert Into Ʊ��ʹ����ϸ
                (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ����, ʹ��ʱ��, Ʊ�ݽ��)
                Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, ����Ա����_In, Sysdate, Ʊ�ݽ��
                From Ʊ��ʹ����ϸ A
                Where ID = l_ʹ��id(I) And ���� = 1 And Not Exists
                 (Select 1 From Ʊ��ʹ����ϸ Where ���� = a.���� And Ʊ�� = a.Ʊ�� And Nvl(����, 0) <> 1);
          
            Forall I In 1 .. l_ʹ��id.Count
              Update Ʊ�ݴ�ӡ��ϸ Set �Ƿ���� = 1 Where ʹ��id = l_ʹ��id(I) And Nvl(�Ƿ����, 0) = 0;
          
          End If;
        End If;
        If n_����ģʽ = 0 Then
          --��ȡ�������һ�εĴ�ӡID(�����Ƕ��ŵ����շѴ�ӡ)
          --����=1��ԭ��=6Ϊ�˷Ѵ�ӡƱ��(��Ʊ)��������
          Select Max(ID)
          Into n_��ӡid
          From (Select b.Id
                 From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
                 Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And b.�������� = 1 And b.No = c_No.No
                 Order By a.ʹ��ʱ�� Desc)
          Where Rownum < 2;
        
          --������ǰû�д�ӡ,���ջ�
          If n_��ӡid Is Not Null Then
            --a.���ŵ���ѭ������ʱֻ���ջ�һ��
            Select Count(1) Into n_Count From Ʊ��ʹ����ϸ Where Ʊ�� = 1 And ���� = 2 And ��ӡid = n_��ӡid;
            If n_Count = 0 Then
              Insert Into Ʊ��ʹ����ϸ
                (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
                Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, Sysdate, ����Ա����_In, Ʊ�ݽ��
                From Ʊ��ʹ����ϸ
                Where ��ӡid = n_��ӡid And Ʊ�� = 1 And ���� = 1;
            Else
              --b.�����˷Ѷ���ջ�ʱ,���һ��ȫ���ջ�Ҫ�ſ����ջص�
              Insert Into Ʊ��ʹ����ϸ
                (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
                Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, Sysdate, ����Ա����_In, Ʊ�ݽ��
                From Ʊ��ʹ����ϸ A
                Where ��ӡid = n_��ӡid And Ʊ�� = 1 And ���� = 1 And Not Exists
                 (Select 1
                       From Ʊ��ʹ����ϸ B
                       Where a.���� = b.���� And ��ӡid = n_��ӡid And Ʊ�� = 1 And ���� = 2);
            End If;
          End If;
        End If;
      End If;
    End If;
  
    --���ҽ��ִ�мƼ�.����ID
    For c_���� In (Select Distinct a.Id, a.ҽ����� As ҽ��id, a.�շ�ϸĿid, b.���ͺ�
                 From ������ü�¼ A, ����ҽ������ B
                 Where a.ҽ����� = b.ҽ��id And a.No = b.No And a.����id = ����id_In And a.�۸񸸺� Is Null And b.��¼���� = 1) Loop
      Update ҽ��ִ�мƼ�
      Set ����id = Null
      Where ҽ��id = c_����.ҽ��id And ���ͺ� = c_����.���ͺ� And �շ�ϸĿid = c_����.�շ�ϸĿid And ִ��״̬ = 2 And ����id = c_����.Id;
    End Loop;
  
    --ɾ������ҽ������(���һ��ɾ��ʱ)
    For c_ҽ�� In (Select Distinct ҽ�����
                 From ������ü�¼
                 Where Mod(��¼����, 10) = 1 And ��¼״̬ In (1, 3) And ҽ����� Is Not Null And NO = c_No.No) Loop
    
      Select Count(1)
      Into n_Count
      From (Select 1
             From ������ü�¼
             Where Mod(��¼����, 10) = 1 And ҽ����� + 0 = c_ҽ��.ҽ����� And NO = c_No.No
             Group By ���
             Having Nvl(Sum(Nvl(����, 1) * ����), 0) <> 0);
    
      If n_Count = 0 Then
        Delete From ����ҽ������ Where ҽ��id = c_ҽ��.ҽ����� And ��¼���� = 1 And NO = c_No.No;
      End If;
    End Loop;
  
    --����_In    Integer:=0, --0:����;1-סԺ
    --����_In    Integer:=1, --1-�շѵ�;2-���ʵ�
    --����_In    Integer:=0, --0:ɾ�����۵�;1-�շѻ����;2-�˷ѻ�����
    --No_In      ������ü�¼.No%Type,
    --ҽ��ids_In Varchar2 := Null
    Zl_ҽ������_�Ʒ�״̬_Update(0, 1, 2, c_No.No);
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����շѼ�¼_����˷�;
/
Create Or Replace Procedure Zl_������շ�_Delete
(
  No_In         ������ü�¼.No%Type,
  ����Ա���_In ������ü�¼.����Ա���%Type,
  ����Ա����_In ������ü�¼.����Ա����%Type
) As
  --���ܣ�ɾ��һ��������շѵ���

  --���α�ΪҪ�˷ѵ��ݵ�����ԭʼ��¼
  Cursor c_Bill Is
    Select a.Id, a.No, a.���ӱ�־, a.�շ�ϸĿid, a.���, a.�۸񸸺�, a.ִ��״̬, a.�շ����, a.����, a.����, a.ҽ�����, j.�������, m.��������
    From ������ü�¼ A, ����ҽ����¼ J, �������� M
    Where a.ҽ����� = j.Id(+) And a.No = No_In And a.��¼���� = 1 And a.��¼״̬ In (1, 3) And a.�շ�ϸĿid + 0 = m.����id(+)
    Order By a.�շ�ϸĿid, a.���;

  --�ù�����ڴ�����Ա�ɿ�������˵Ĳ�ͬ���㷽ʽ�Ľ��
  Cursor c_Money(����id_In ����Ԥ����¼.����id%Type) Is
    Select ���㷽ʽ, ��Ԥ��
    From ����Ԥ����¼
    Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = ����id_In And ���㷽ʽ Is Not Null And Nvl(��Ԥ��, 0) <> 0 And Nvl(У�Ա�־, 0) = 0;

  --���α����ڲ����շ�ʱʹ�ù��ĳ�Ԥ�����¼
  Cursor c_Deposit(V����id ����Ԥ����¼.����id%Type) Is
    Select NO, ID, ����id, ��Ԥ�� As ���, Ԥ�����
    From ����Ԥ����¼
    Where ��¼���� In (1, 11) And ��¼״̬ In (1, 3) And ����id = V����id And Nvl(��Ԥ��, 0) <> 0
    Order By ID Desc;

  n_����id   ������Ϣ.����id%Type;
  n_����id   ������ü�¼.����id%Type;
  n_������� ����Ԥ����¼.�������%Type;
  n_��ӡid   Ʊ�ݴ�ӡ����.Id%Type;

  n_Ԥ����� ����Ԥ����¼.��Ԥ��%Type;
  n_����ֵ   ����Ԥ����¼.��Ԥ��%Type;
  n_��ֵid   ����Ԥ����¼.Id%Type;
  --�����˷Ѽ������
  n_ʣ������ Number;
  n_ʣ��Ӧ�� Number;
  n_ʣ��ʵ�� Number;
  n_ʣ��ͳ�� Number;
  n_׼������ Number;
  n_�˷Ѵ��� Number;

  n_Ӧ�ս�� Number;
  n_ʵ�ս�� Number;
  n_ͳ���� Number;
  n_�ܽ��   Number;
  n_����״̬ ������ü�¼.����״̬%Type;
  n_�����˷� Number; --�Ƿ��һ���˷���ȫ���˷�,��ÿ���˷ѹ������жϵõ���
  n_��id     ����ɿ����.Id%Type;

  v_�˷ѽ��� ���㷽ʽ.����%Type;
  l_ʹ��id   t_NumList := t_NumList();
  n_Dec      Number;
  d_Date     Date;
  n_Count    Number;
  n_ԭ����id Number;

  Err_Item Exception;
  v_Err_Msg      Varchar2(255);
  n_����ģʽ     Number(3);
  v_Para         Varchar2(1000);
  n_ҽ��ִ�мƼ� Number;
  n_�Ƿ����Ʊ�� ����Ԥ����¼.�Ƿ����Ʊ��%Type;
Begin
  n_��id := Zl_Get��id(����Ա����_In);

  --�Ƿ��Ѿ�ȫ����ȫִ��(ֻ�Ǹõ������ŵ��ݵļ��)
  Select Nvl(Count(*), 0)
  Into n_Count
  From ������ü�¼
  Where NO = No_In And ��¼���� = 1 And ��¼״̬ In (1, 3) And Nvl(ִ��״̬, 0) <> 1;
  If n_Count = 0 Then
    v_Err_Msg := '�õ����е���Ŀ�Ѿ�ȫ����ȫִ�У�';
    Raise Err_Item;
  End If;

  --δ��ȫִ�е���Ŀ�Ƿ���ʣ������(ֻ�����ŵ��ݵļ��)
  --ִ��״̬��ԭʼ��¼���ж�
  Select Nvl(Count(*), 0)
  Into n_Count
  From (Select ���, Sum(����) As ʣ������
         From (Select ��¼״̬, Nvl(�۸񸸺�, ���) As ���, Avg(Nvl(����, 1) * ����) As ����
                From ������ü�¼
                Where NO = No_In And ��¼���� = 1 And Nvl(���ӱ�־, 0) <> 9 And
                      Nvl(�۸񸸺�, ���) In
                      (Select Nvl(�۸񸸺�, ���)
                       From ������ü�¼
                       Where NO = No_In And ��¼���� = 1 And ��¼״̬ In (1, 3) And Nvl(ִ��״̬, 0) <> 1)
                Group By ��¼״̬, Nvl(�۸񸸺�, ���))
         Group By ���
         Having Sum(����) <> 0);
  If n_Count = 0 Then
    v_Err_Msg := '�õ�����δ��ȫִ�в�����Ŀʣ������Ϊ��,û�п����˷ѵķ��ã�';
    Raise Err_Item;
  End If;
  --ȷ���Ƿ���ҽ��ִ�мƼ��д�������,�����������,�����ҽ��ִ�мƼ۽����˷�,���򰴾ɷ�ʽ���д���
  Select Count(1)
  Into n_ҽ��ִ�мƼ�
  From ������ü�¼ A, ҽ��ִ�мƼ� B
  Where a.ҽ����� = b.ҽ��id And a.��¼���� = 1 And a.No = No_In And a.��¼״̬ In (1, 3) And Rownum = 1;

  ---------------------------------------------------------------------------------
  --���ñ���
  Select Sysdate Into d_Date From Dual;
  Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
  n_������� := Null;

  --���С��λ��
  Select zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  --��ȡ���㷽ʽ����
  Begin
    Select ���� Into v_�˷ѽ��� From ���㷽ʽ Where ���� = 1;
  Exception
    When Others Then
      v_�˷ѽ��� := '�ֽ�';
  End;

  ---------------------------------------------------------------------------------
  --ѭ������ÿ�з���(������Ŀ��)
  n_�ܽ��   := 0;
  n_�����˷� := 1;
  For r_Bill In c_Bill Loop
    If Nvl(r_Bill.ִ��״̬, 0) <> 1 Then
      --��ʣ������,ʣ��Ӧ��,ʣ��ʵ��
      Select Sum(Nvl(����, 1) * ����), Sum(Ӧ�ս��), Sum(ʵ�ս��), Sum(ͳ����)
      Into n_ʣ������, n_ʣ��Ӧ��, n_ʣ��ʵ��, n_ʣ��ͳ��
      From ������ü�¼
      Where NO = No_In And ��¼���� = 1 And ��� = r_Bill.���;
    
      If n_ʣ������ = 0 Then
        --�����δ�޶��к�,ԭʼ�����еĸñ��Ѿ�ȫ���˷�(ִ��״̬=0��һ�ֿ���)
        n_�����˷� := 0;
      Else
        --׼������(��ҩƷ��ĿΪʣ������,ԭʼ����)
        If Instr(',4,5,6,7,', r_Bill.�շ����) = 0 Or (r_Bill.�շ���� = '4' And Nvl(r_Bill.��������, 0) = 0) Then
          --@@@
          --��ҩƷ����(�Ծ���ҽ��ִ��Ϊ׼���м��)
          --: 1.����ҽ�����͵�,����ҽ��ִ��Ϊ׼(�����ܰ���:���;����;����;������Ѫ)
          --: 2.������ҽ����,����ʣ������Ϊ׼
          n_Count := 0;
          If Instr(',C,D,F,G,K,', ',' || r_Bill.������� || ',') = 0 And r_Bill.������� Is Not Null Then
            If n_ҽ��ִ�мƼ� = 1 Then
              Select Decode(Sign(Sum(����)), -1, 0, Sum(����)), Count(*)
              Into n_׼������, n_Count
              From (Select Max(Decode(a.��¼״̬, 2, 0, a.Id)) As ID, Max(a.ҽ�����) As ҽ��id, Max(a.�շ�ϸĿid) As �շ�ϸĿid,
                            Sum(Nvl(a.����, 1) * Nvl(a.����, 1)) As ����,
                            Sum(Decode(a.��¼״̬, 2, 0, Nvl(a.����, 1) * Nvl(a.����, 1))) As ԭʼ����
                     From ������ü�¼ A, ����ҽ����¼ M
                     Where a.ҽ����� = m.Id And Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And
                           Instr('5,6,7', a.�շ����) = 0 And a.No = No_In And a.��� = r_Bill.��� And a.��¼���� = 1 And
                           a.��¼״̬ In (1, 2, 3) And a.�۸񸸺� Is Null
                     Group By a.���
                     Union All
                     Select a.Id, a.ҽ����� As ҽ��id, a.�շ�ϸĿid, -1 * b.���� As ��ִ��, 0 ԭʼ����
                     From ������ü�¼ A, ҽ��ִ�мƼ� B, ����ҽ����¼ M
                     Where a.ҽ����� = b.ҽ��id And a.�շ�ϸĿid = b.�շ�ϸĿid + 0 And a.ҽ����� = m.Id And
                           Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And Instr('5,6,7', a.�շ����) = 0 And
                           (Exists
                            (Select 1
                             From ����ҽ��ִ��
                             Where b.ҽ��id = ҽ��id And b.���ͺ� = ���ͺ� And b.Ҫ��ʱ�� = Ҫ��ʱ�� And Nvl(ִ�н��, 0) = 1) Or Exists
                            (Select 1
                             From ����ҽ������
                             Where b.ҽ��id = ҽ��id And b.���ͺ� = ���ͺ� And Nvl(ִ��״̬, 0) = 1)) And Not Exists
                      (Select 1
                            From ����ҽ������
                            Where a.ҽ����� = ҽ��id And a.No = NO And Mod(a.��¼����, 10) = ��¼����) And a.No = No_In And
                           a.��� = r_Bill.��� And a.��¼���� = 1 And a.��¼״̬ In (1, 3) ��and a.�۸񸸺� Is Null) Q1
              Where Not Exists (Select 1 From ҩƷ�շ���¼ Where ����id = Q1.Id) Having Max(ID) <> 0;
            Else
              Select Nvl(Sum(����), 0), Count(*)
              Into n_׼������, n_Count
              From (Select a.ҽ��id, a.�շ�ϸĿid, Nvl(a.����, 1) * Nvl(b.��������, 1) As ����
                     From ����ҽ���Ƽ� A, ����ҽ������ B, ������ü�¼ J, ����ҽ����¼ M
                     Where a.ҽ��id = b.ҽ��id And a.ҽ��id = m.Id And Nvl(b.ִ��״̬, 0) <> 1 And a.ҽ��id = j.ҽ����� And
                           a.�շ�ϸĿid = j.�շ�ϸĿid And j.No = No_In And j.��¼���� = 1 And j.��� = r_Bill.��� And j.��¼״̬ In (1, 3) And
                           j.�۸񸸺� Is Null And Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And Exists
                      (Select 1
                            From ����ҽ���Ƽ� A
                            Where a.ҽ��id = j.ҽ����� And a.�շ�ϸĿid = j.�շ�ϸĿid And Nvl(a.�շѷ�ʽ, 0) = 0) And Not Exists
                      (Select 1 From ҩƷ�շ���¼ Where ����id = j.Id)
                     Union All
                     Select a.ҽ��id, a.�շ�ϸĿid, -1 * Nvl(a.����, 1) * Nvl(c.��������, 1) As ����
                     From ����ҽ���Ƽ� A, ����ҽ������ B, ����ҽ��ִ�� C, ������ü�¼ J, ����ҽ����¼ M
                     Where a.ҽ��id = b.ҽ��id And b.ҽ��id = c.ҽ��id And b.���ͺ� = c.���ͺ� And a.ҽ��id = m.Id And Nvl(c.ִ�н��, 1) = 1 And
                           Nvl(b.ִ��״̬, 0) <> 1 And a.ҽ��id = j.ҽ����� And a.�շ�ϸĿid = j.�շ�ϸĿid And j.No = No_In And
                           j.��¼���� = 1 And Nvl(a.�շѷ�ʽ, 0) = 0 And j.��� = r_Bill.��� And j.��¼״̬ In (1, 3) And j.�۸񸸺� Is Null And
                           Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And Not Exists
                      (Select 1 From ҩƷ�շ���¼ Where ����id = j.Id) And Not Exists
                      (Select 1 From �������� Where ����id = j.�շ�ϸĿid And Nvl(��������, 0) = 1)
                     Union All
                     Select a.ҽ����� As ҽ��id, a.�շ�ϸĿid, Nvl(a.����, 1) * a.���� As ����
                     From ������ü�¼ A, ����ҽ����¼ M
                     Where a.ҽ����� = m.Id And Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And a.No = No_In And
                           a.��¼���� = 1 And a.��� = r_Bill.��� And a.��¼״̬ = 2 And a.�۸񸸺� Is Null And Not Exists
                      (Select 1 From ҩƷ�շ���¼ Where NO = No_In And ���� In (8, 24) And ҩƷid = a.�շ�ϸĿid));
            End If;
          End If;
          If Nvl(n_Count, 0) <> 0 And n_׼������ = 0 Then
            v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з�����ִ��,�������˷ѣ�';
            Raise Err_Item;
          End If;
        
          If Nvl(n_Count, 0) = 0 Then
            n_׼������ := n_ʣ������;
          End If;
        
        Else
          Select Nvl(Sum(Nvl(����, 1) * ʵ������), 0), Count(*)
          Into n_׼������, n_Count
          From ҩƷ�շ���¼
          Where NO = No_In And ���� In (8, 24) And Mod(��¼״̬, 3) = 1 --@@@
                And ����� Is Null And ����id = r_Bill.Id;
        
          --��ʣ��������׼�������������������
          --1.���������õ������޶�Ӧ���շ���¼,��ʱʹ��ʣ������
          --2.��������,��ʱ�ѷ�ҩ����
          If n_׼������ = 0 Then
            If r_Bill.�շ���� = '4' Then
              If n_Count > 0 Then
                v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����ѷ���,�����Ϻ����˷ѣ�';
                Raise Err_Item;
              Else
                n_׼������ := n_ʣ������;
              End If;
            Else
              v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����ѷ�ҩ,����ҩ�����˷ѣ�';
              Raise Err_Item;
            End If;
          End If;
        End If;
      
        --�Ƿ񲿷��˷�
        If r_Bill.ִ��״̬ = 2 Or n_׼������ <> Nvl(r_Bill.����, 1) * r_Bill.���� Then
          n_�����˷� := 0;
        End If;
      
        --����������ü�¼
        n_����״̬ := 0;
        --�ñ���Ŀ�ڼ����˷�
        Select Nvl(Max(Abs(ִ��״̬)), 0) + 1
        Into n_�˷Ѵ���
        From ������ü�¼
        Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 2 And Nvl(ִ��״̬, 0) < 0 And ��� = r_Bill.���;
      
        n_Ӧ�ս�� := Round(n_ʣ��Ӧ�� * (n_׼������ / n_ʣ������), n_Dec);
        n_ʵ�ս�� := Round(n_ʣ��ʵ�� * (n_׼������ / n_ʣ������), n_Dec);
        n_ͳ���� := Round(n_ʣ��ͳ�� * (n_׼������ / n_ʣ������), n_Dec);
        n_�ܽ��   := n_�ܽ�� + n_ʵ�ս��;
      
        --�����˷Ѽ�¼
        Insert Into ������ü�¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id, �շ����, �շ�ϸĿid,
           ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����, ִ��״̬,
           ����״̬, ִ��ʱ��, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, ����, �ɿ���id)
          Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�,
                 ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ, Decode(Sign(n_׼������ - Nvl(����, 1) * ����), 0, ����, 1), ��ҩ����,
                 Decode(Sign(n_׼������ - Nvl(����, 1) * ����), 0, -1 * ����, -1 * n_׼������), �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����,
                 -1 * n_Ӧ�ս��, -1 * n_ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����, -1 * n_�˷Ѵ���, n_����״̬, ִ��ʱ��, ����Ա���_In, ����Ա����_In,
                 ����ʱ��, d_Date, n_����id, -1 * n_ʵ�ս��, ������Ŀ��, ���մ���id, -1 * n_ͳ����, ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���,
                 ��������, ����, n_��id
          From ������ü�¼
          Where ID = r_Bill.Id;
      
        --���ԭ���ü�¼
        --ִ��״̬:ȫ������(׼����=ʣ����)���Ϊ0,������Ϊ1,�쳣�շѵ�,���Ǳ���9
        Update ������ü�¼
        Set ��¼״̬ = 3, ִ��״̬ = Decode(Nvl(ִ��״̬, 0), 9, 9, Decode(Sign(n_׼������ - n_ʣ������), 0, 0, 1))
        Where ID = r_Bill.Id;
      End If;
    Else
      --���:û�޶��к�,ԭʼ�����а����Ѿ���ȫִ�е�
      n_�����˷� := 0;
    End If;
  End Loop;

  ---------------------------------------------------------------------------------
  --������Ԥ����¼
  --�Զ���������,Ĭ�ϱ���һλ
  n_�ܽ�� := Round(n_�ܽ��, 1);
  --ԭ���ݵĽ���ID
  Select ����id, ����id
  Into n_ԭ����id, n_����id
  From ������ü�¼
  Where NO = No_In And ��¼���� = 1 And ��¼״̬ In (1, 3) And Rownum = 1;

  If n_�����˷� = 1 Then
    --���ݵ�һ���˷���ȫ������
    --��Ԥ�����ּ�¼
    Insert Into ����Ԥ����¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��, ����id,
       �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������)
      Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, d_Date,
             ����Ա����_In, ����Ա���_In, -1 * ��Ԥ��, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 0, n_�������, 3
      From ����Ԥ����¼
      Where ��¼���� In (1, 11) And ����id = n_ԭ����id And Nvl(��Ԥ��, 0) <> 0;
    --������Ԥ�����
    For v_Ԥ�� In (Select NO, Ԥ�����, Nvl(Sum(Nvl(��Ԥ��, 0)), 0) As Ԥ�����, ����id
                 From ����Ԥ����¼
                 Where ��¼���� In (1, 11) And ����id = n_ԭ����id
                 Group By NO, Ԥ�����, ����id
                 Having Sum(Nvl(��Ԥ��, 0)) <> 0) Loop
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(v_Ԥ��.Ԥ�����, 0)
      Where ����id = v_Ԥ��.����id And ���� = 1 And ���� = Nvl(v_Ԥ��.Ԥ�����, 2)
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, Ԥ�����, ����)
        Values
          (v_Ԥ��.����id, Nvl(v_Ԥ��.Ԥ�����, 2), Nvl(v_Ԥ��.Ԥ�����, 0), 1);
        n_����ֵ := n_Ԥ�����;
      End If;
      If n_����ֵ = 0 Then
        Delete From �������
        Where ����id = v_Ԥ��.����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
      End If;
    
      --����Ԥ���������
      Select Max(ID) Into n_��ֵid From ����Ԥ����¼ Where NO = v_Ԥ��.No And ��¼���� = 1 And ��¼״̬ <> 2;
      If Nvl(n_��ֵid, 0) <> 0 Then
        Update Ԥ���������
        Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(v_Ԥ��.Ԥ�����, 0)
        Where ����id = v_Ԥ��.����id And Ԥ��id = n_��ֵid
        Returning Ԥ����� Into n_����ֵ;
        If Sql%RowCount = 0 Then
          Insert Into Ԥ���������
            (Ԥ��id, ����id, Ԥ�����, Ԥ�����)
          Values
            (n_��ֵid, v_Ԥ��.����id, Nvl(v_Ԥ��.Ԥ�����, 2), Nvl(v_Ԥ��.Ԥ�����, 0));
          n_����ֵ := Nvl(v_Ԥ��.Ԥ�����, 0);
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From Ԥ��������� Where Ԥ��id = n_��ֵid And Nvl(Ԥ�����, 0) = 0;
        End If;
      End If;
    End Loop;
  
    --ԭ���˻�(��Ԥ����ǰ���Ѵ���)
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����,
       �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������)
      Select ����Ԥ����¼_Id.Nextval, ��¼����, NO, 2, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, d_Date, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���_In, ����Ա����_In,
             -1 * ��Ԥ��, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 0, n_�������, 3
      From ����Ԥ����¼ A, (Select ���� From ���㷽ʽ Where ���� In (3, 4)) J,
           (Select m.Id As Ԥ��id From ����Ԥ����¼ M Where m.����id = n_ԭ����id And m.��¼���� = 3 And m.��¼״̬ = 1) Q
      Where a.��¼���� = 3 And a.��¼״̬ = 1 And a.����id = n_ԭ����id And a.Id = q.Ԥ��id(+) And a.���㷽ʽ = j.����(+);
  Else
    --�����˷�ֱ����Ϊָ�����㷽ʽ
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
       ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������)
      Select ����Ԥ����¼_Id.Nextval, 3, No_In, 2, ����id, ��ҳid, '�����˷ѽ���', v_�˷ѽ���, d_Date, Null, Null, Null, ����Ա���_In, ����Ա����_In,
             -1 * n_�ܽ��, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 0, n_�������, 3
      
      From ����Ԥ����¼
      Where ��¼���� = 3 And ��¼״̬ In (1, 3) And ����id = n_ԭ����id And Rownum = 1;
  
    --����շ�ʱֻʹ����Ԥ����,��Ҫ��Ԥ��,���ҿ����ж�ʳ�Ԥ��
    If Sql%RowCount = 0 Then
      n_Ԥ����� := n_�ܽ��;
    
      For r_Deposit In c_Deposit(n_ԭ����id) Loop
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
           ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������)
          Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�,
                 d_Date, ����Ա����_In, ����Ա���_In, Decode(Sign(r_Deposit.��� - n_Ԥ�����), -1, -1 * r_Deposit.���, -1 * n_Ԥ�����),
                 n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 0, n_�������, 3
          From ����Ԥ����¼
          Where ID = r_Deposit.Id;
      
        --���²���Ԥ�����
        Update �������
        Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(r_Deposit.���, 0)
        Where ����id = r_Deposit.����id And ���� = 1 And ���� = 1
        Returning Ԥ����� Into n_����ֵ;
        If Sql%RowCount = 0 Then
          Insert Into ������� (����id, ����, Ԥ�����, ����) Values (r_Deposit.����id, 1, Nvl(r_Deposit.���, 0), 1);
          n_����ֵ := Nvl(r_Deposit.���, 0);
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From �������
          Where ����id = r_Deposit.����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
        End If;
        --����Ԥ���������
        Select Max(ID) Into n_��ֵid From ����Ԥ����¼ Where NO = r_Deposit.No And ��¼���� = 1 And ��¼״̬ <> 2;
        If Nvl(n_��ֵid, 0) <> 0 Then
          Update Ԥ���������
          Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(r_Deposit.���, 0)
          Where ����id = r_Deposit.����id And Ԥ��id = n_��ֵid
          Returning Ԥ����� Into n_����ֵ;
          If Sql%RowCount = 0 Then
            Insert Into Ԥ���������
              (Ԥ��id, ����id, Ԥ�����, Ԥ�����)
            Values
              (n_��ֵid, r_Deposit.����id, Nvl(r_Deposit.Ԥ�����, 2), Nvl(r_Deposit.���, 0));
            n_����ֵ := Nvl(r_Deposit.���, 0);
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From Ԥ��������� Where Ԥ��id = n_��ֵid And Nvl(Ԥ�����, 0) = 0;
          End If;
        End If;
        --����Ƿ��Ѿ�������
        If r_Deposit.��� < n_Ԥ����� Then
          n_Ԥ����� := n_Ԥ����� - r_Deposit.���;
        Else
          n_Ԥ����� := 0;
        End If;
        If n_Ԥ����� = 0 Then
          Exit;
        End If;
      End Loop;
    
    End If;
  End If;
  --����ԭ��¼
  Update ����Ԥ����¼ Set ��¼״̬ = 3 Where ��¼���� = 3 And ��¼״̬ In (1, 3) And ����id = n_ԭ����id;

  Select Nvl(Sum(Nvl(���ʽ��, 0)), 0) Into n_ʵ�ս�� From ������ü�¼ Where ����id = n_����id;
  Select Nvl(Sum(Nvl(��Ԥ��, 0)), 0) Into n_����ֵ From ����Ԥ����¼ Where ����id = n_����id;

  n_ʵ�ս�� := n_ʵ�ս�� - n_����ֵ;

  If n_ʵ�ս�� <> 0 Then
    --δ�ҵ����²��������
    Zl_���շ����_Insert(No_In, n_����id, n_����id, n_ʵ�ս��, d_Date, ����Ա���_In, ����Ա����_In, 1);
  End If;

  --���� �Ƿ����Ʊ�� ���
  Select Max(a.�Ƿ����Ʊ��)
  Into n_�Ƿ����Ʊ��
  From ����Ԥ����¼ A
  Where a.����id = n_ԭ����id And a.��¼���� In (11, 3);

  Update ����Ԥ����¼ Set �Ƿ����Ʊ�� = n_�Ƿ����Ʊ�� Where ����id = n_����id;

  --��Ա�ɿ����(ע����Ԥ����¼�����Ŵ������������ʻ��ȵĽ�����,�����˳�Ԥ����)
  For r_Moneyrow In c_Money(n_����id) Loop
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + r_Moneyrow.��Ԥ��
    Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_Moneyrow.���㷽ʽ
    Returning ��� Into n_����ֵ;
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ����
        (�տ�Ա, ���㷽ʽ, ����, ���)
      Values
        (����Ա����_In, r_Moneyrow.���㷽ʽ, 1, r_Moneyrow.��Ԥ��);
      n_����ֵ := r_Moneyrow.��Ԥ��;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ��Ա�ɿ����
      Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_Moneyrow.���㷽ʽ And Nvl(���, 0) = 0;
    End If;
  End Loop;

  ---------------------------------------------------------------------------------
  --�˷�Ʊ�ݻ���(��ȫ��ʱ�Ż���,�����������ش�����л���)
  If Nvl(n_�Ƿ����Ʊ��, 0) = 0 Then
    --���ñ�־||NO;ִ�п���(����);�վݷ�Ŀ(��ҳ����,����);�շ�ϸĿ(����)
    v_Para     := Nvl(zl_GetSysParameter('Ʊ�ݷ������', 1121), '0||0;0;0,0;0');
    n_����ģʽ := zl_To_Number(Substr(v_Para, 1, 1));
    If n_����ģʽ <> 0 Then
      --�ջ�Ʊ��
      Select ʹ��id
      Bulk Collect
      Into l_ʹ��id
      From (Select Distinct b.ʹ��id From Ʊ�ݴ�ӡ��ϸ B Where b.No = No_In And Nvl(b.Ʊ��, 0) = 1);
    
      n_����ģʽ := l_ʹ��id.Count;
      If l_ʹ��id.Count <> 0 Then
        --������ռ�¼
        Forall I In 1 .. l_ʹ��id.Count
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ����, ʹ��ʱ��, Ʊ�ݽ��)
            Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, ����Ա����_In, d_Date, Ʊ�ݽ��
            From Ʊ��ʹ����ϸ A
            Where ID = l_ʹ��id(I) And ���� = 1 And Not Exists
             (Select 1 From Ʊ��ʹ����ϸ Where ���� = a.���� And Ʊ�� = a.Ʊ�� And Nvl(����, 0) <> 1);
      
        Forall I In 1 .. l_ʹ��id.Count
          Update Ʊ�ݴ�ӡ��ϸ Set �Ƿ���� = 1 Where ʹ��id = l_ʹ��id(I) And Nvl(�Ƿ����, 0) = 0;
      
      End If;
    End If;
  
    If n_����ģʽ = 0 Then
      --��ȡ�������һ�εĴ�ӡID(�����Ƕ��ŵ����շѴ�ӡ)
      Begin
        --����=1��ԭ��=6Ϊ�˷Ѵ�ӡƱ��(��Ʊ)��������
        Select ID
        Into n_��ӡid
        From (Select b.Id
               From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
               Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And b.�������� = 1 And b.No = No_In
               Order By a.ʹ��ʱ�� Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          Null;
      End;
      --������ǰû�д�ӡ,���ջ�
      If n_��ӡid Is Not Null Then
        --a.���ŵ���ѭ������ʱֻ���ջ�һ��
        Select Count(*) Into n_Count From Ʊ��ʹ����ϸ Where Ʊ�� = 1 And ���� = 2 And ��ӡid = n_��ӡid;
        If n_Count = 0 Then
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
            Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, ����Ա����_In, Ʊ�ݽ��
            From Ʊ��ʹ����ϸ
            Where ��ӡid = n_��ӡid And Ʊ�� = 1 And ���� = 1;
        Else
          --b.�����˷Ѷ���ջ�ʱ,���һ��ȫ���ջ�Ҫ�ſ����ջص�
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
            Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, ����Ա����_In, Ʊ�ݽ��
            From Ʊ��ʹ����ϸ A
            Where ��ӡid = n_��ӡid And Ʊ�� = 1 And ���� = 1 And Not Exists
             (Select 1 From Ʊ��ʹ����ϸ B Where a.���� = b.���� And ��ӡid = n_��ӡid And Ʊ�� = 1 And ���� = 2);
        End If;
      End If;
    End If;
  End If;

  ---------------------------------------------------------------------------------
  --ҩƷ�����������
  --���밴�ա��շ�ϸĿid���������򣬷�ֹ��������ҩƷ��桱��
  For r_Expenses In (Select ID
                     From ������ü�¼
                     Where NO = No_In And ��¼���� = 1 And ��¼״̬ In (1, 3) And �շ���� In ('4', '5', '6', '7')
                     Order By �շ�ϸĿid) Loop
    Zl_ҩƷ�շ���¼_�����˷�(r_Expenses.Id);
  End Loop;

  --ҽ������
  --ɾ������ҽ������(���һ��ɾ��ʱ)
  For c_ҽ�� In (Select Distinct ҽ�����
               From ������ü�¼
               Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 3 And ҽ����� Is Not Null) Loop
    Select Nvl(Count(*), 0)
    Into n_Count
    From (Select ���, Sum(����) As ʣ������
           From (Select ��¼״̬, ִ��״̬, Nvl(�۸񸸺�, ���) As ���, Avg(Nvl(����, 1) * ����) As ����
                  From ������ü�¼
                  Where ��¼���� = 1 And Nvl(���ӱ�־, 0) <> 9 And ҽ����� + 0 = c_ҽ��.ҽ����� And NO = No_In
                  Group By ��¼״̬, ִ��״̬, Nvl(�۸񸸺�, ���))
           Group By ���
           Having Sum(����) <> 0);
  
    If n_Count = 0 Then
      Delete From ����ҽ������ Where ҽ��id = c_ҽ��.ҽ����� And ��¼���� = 1 And NO = No_In;
    End If;
  End Loop;

  --����_In    Integer:=0, --0:����;1-סԺ
  --����_In    Integer:=1, --1-�շѵ�;2-���ʵ�
  --����_In    Integer:=0, --0:ɾ�����۵�;1-�շѻ����;2-�˷ѻ�����
  --No_In      ������ü�¼.No%Type,
  --ҽ��ids_In varchar2 := Null
  Zl_ҽ������_�Ʒ�״̬_Update(0, 1, 2, No_In);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_������շ�_Delete;
/
Create Or Replace Procedure Zl_���ò������_Modify
(
  ��������_In     Number,
  ����id_In       In ���ò����¼.����id%Type,
  ���㷽ʽ_In     Varchar2,
  �����id_In     ����Ԥ����¼.�����id%Type := Null,
  ����_In         ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In   ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In     ����Ԥ����¼.����˵��%Type := Null,
  �����_In     ������ü�¼.ʵ�ս��%Type := Null,
  ��ɽ���_In     Number := 0,
  ��Ԥ��_In       ����Ԥ����¼.��Ԥ��%Type := Null,
  У�Ա�־_In     ����Ԥ����¼.У�Ա�־%Type := 0,
  �Ƿ����Ʊ��_In ����Ԥ����¼.�Ƿ����Ʊ��%Type := Null
) As
  ------------------------------------------------------------------------------------------------------------------------------ 
  --����:���ղ������ʱ,�޸Ľ���������Ϣ 
  --��������_In: 
  --   0-��ͨ���㷽ʽ: 
  --     ���㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������|�������|����ժҪ||.. ;Ҳ�������. 
  --   1.����������: 
  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:���㷽ʽ|������|�������|����ժҪ 
  --     �ڿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ���� 
  --   2-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���) 
  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.. 
  -- �����_In:��������ʱ,���� 
  -- ��ɽ���_In:1-��ɲ������;0-δ��ɲ������;2-������쳣���� 
  -- ��Ԥ��_In:��Ԥ�����˿�ʱΪ�����տ�ʱΪ�� 
  -- У�Ա�־_In  ��������_InΪ1ʱ��Ч 
  --�Ƿ����Ʊ��_In:null-��ʾ�����ڲ�ֱ���жϣ��ǿձ�ʾֱ���Դ����Ϊ׼
  ------------------------------------------------------------------------------------------------------------------------------ 
  v_����   ���㷽ʽ.����%Type;
  n_Count    Number(18);
  n_�Ự��   ����Ԥ����¼.�Ự��%Type; --��ʽ��SID+'_'+SERIAL# 
  v_������Ա ����Ԥ����¼.������Ա%Type;

  v_�������� Varchar2(4000);
  v_��ǰ���� Varchar2(4000);
  v_���㷽ʽ ����Ԥ����¼.���㷽ʽ%Type;
  n_������ ����Ԥ����¼.��Ԥ��%Type;
  v_������� ����Ԥ����¼.�������%Type;
  v_����ժҪ ����Ԥ����¼.ժҪ%Type;

  n_��Ԥ�� ����Ԥ����¼.��Ԥ��%Type;
  n_����ֵ ����Ԥ����¼.��Ԥ��%Type;
  n_Ԥ��id ����Ԥ����¼.Id%Type;
  l_Ԥ��id t_NumList := t_NumList();

  n_�Ƿ����Ʊ�� Number(2);
  n_����         ���ս����¼.����%Type;

  Cursor c_Balance Is
    Select ��¼����, NO, ��¼״̬, ʵ��Ʊ��, ����id, �շѽ���id, ����״̬, ����Ա���, ����Ա����, �Ǽ�ʱ��, �ɿ���id, ����id, �������, ���ӱ�־
    From ���ò����¼ A
    Where ����id = ����id_In And ��¼���� = 1 And Rownum < 2;
  r_Balance c_Balance%RowType;

  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin
  Begin
    Select Sid || '_' || Serial# Into n_�Ự�� From V$session Where Audsid = Userenv('sessionid');
  Exception
    When Others Then
      n_�Ự�� := Null;
  End;
  v_������Ա := zl_UserName;

  Select Count(1) Into n_Count From ���ò����¼ Where ����id = ����id_In And Rownum < 2 And ��¼���� = 1;
  If n_Count = 0 Then
    v_Err_Msg := 'δ�ҵ�ҽ�����������ݣ����ܼ�������!';
    Raise Err_Item;
  End If;

  Open c_Balance;
  Fetch c_Balance
    Into r_Balance;

  If Nvl(�����_In, 0) <> 0 Then
    Begin
      Select ���� Into v_���� From ���㷽ʽ Where Nvl(����, 0) = 9;
    Exception
      When Others Then
        v_���� := '����';
    End;
  
    Update ����Ԥ����¼
    Set ��Ԥ�� = Nvl(��Ԥ��, 0) + Nvl(�����_In, 0)
    Where ����id = ����id_In And ���㷽ʽ = v_����;
    If Sql%NotFound Then
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־, �����id,
         ���㿨���, ����, ������ˮ��, ����˵��, �������, ��������, �Ự��)
      Values
        (����Ԥ����¼_Id.Nextval, 6, r_Balance.No, r_Balance.��¼״̬, r_Balance.����id, Null, Null, v_����, r_Balance.�Ǽ�ʱ��,
         r_Balance.����Ա���, r_Balance.����Ա����, �����_In, r_Balance.����id, r_Balance.�ɿ���id, Sysdate, v_������Ա, r_Balance.�������, 2,
         Null, Null, ����_In, ������ˮ��_In, ����˵��_In, Null, 6, n_�Ự��);
      Update ����Ԥ����¼ Set ��Ԥ�� = Nvl(��Ԥ��, 0) - �����_In Where ����id = ����id_In And ���㷽ʽ Is Null;
    End If;
  End If;

  --��Ԥ���� 
  If Nvl(��Ԥ��_In, 0) < 0 Then
    n_��Ԥ�� := -1 * ��Ԥ��_In;
    For v_��Ԥ�� In (Select Max(ID) As Ԥ��id, NO, ����id, Max(�տ�ʱ��) As �տ�ʱ��, Sum(Nvl(��Ԥ��, 0)) As ���
                  From ����Ԥ����¼
                  Where Mod(��¼����, 10) = 1 And Nvl(Ԥ�����, 0) = 1 And
                        ����id In
                        (Select a.����id
                         From ������ü�¼ A, ������ü�¼ B
                         Where a.��¼���� = b.��¼���� And a.No = b.No And a.��� = b.��� And b.��¼״̬ <> 2 And
                               b.����id In
                               (Select �շѽ���id From ���ò����¼ Where ��¼���� = 1 And ������� = r_Balance.�������))
                  Group By NO, ����id
                  Having Sum(Nvl(��Ԥ��, 0)) > 0
                  Order By �տ�ʱ�� Desc) Loop
    
      If v_��Ԥ��.��� - n_��Ԥ�� < 0 Then
        n_������ := -1 * v_��Ԥ��.���;
        n_��Ԥ��   := n_��Ԥ�� - v_��Ԥ��.���;
      Else
        n_������ := -1 * n_��Ԥ��;
        n_��Ԥ��   := 0;
      End If;
    
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־,
         �����id, ���㿨���, ����, ������ˮ��, ����˵��, �������, Ԥ�����, ��������, �Ự��)
        Select ����Ԥ����¼_Id.Nextval, 11, NO, ʵ��Ʊ��, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, r_Balance.�Ǽ�ʱ��, r_Balance.����Ա���,
               r_Balance.����Ա����, n_������, r_Balance.����id, r_Balance.�ɿ���id, Sysdate, v_������Ա, r_Balance.�������, 2, �����id,
               ���㿨���, ����, ������ˮ��, ����˵��, �������, Ԥ�����, 6, n_�Ự��
        From ����Ԥ����¼
        Where ID = v_��Ԥ��.Ԥ��id;
    
      --����Ԥ��������� 
      Select Max(ID) Into n_Ԥ��id From ����Ԥ����¼ Where NO = v_��Ԥ��.No And ��¼���� = 1 And ��¼״̬ <> 2;
      If Nvl(n_Ԥ��id, 0) <> 0 Then
        Update Ԥ���������
        Set Ԥ����� = Nvl(Ԥ�����, 0) + (-1 * n_������)
        Where ����id = v_��Ԥ��.����id And Ԥ��id = n_Ԥ��id
        Returning Ԥ����� Into n_����ֵ;
        If Sql%RowCount = 0 Then
          Insert Into Ԥ���������
            (Ԥ��id, ����id, Ԥ�����, Ԥ�����)
          Values
            (n_Ԥ��id, v_��Ԥ��.����id, 1, -1 * n_������);
          n_����ֵ := -1 * n_������;
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From Ԥ��������� Where Ԥ��id = n_Ԥ��id And Nvl(Ԥ�����, 0) = 0;
        End If;
      End If;
    
      --���²������ 
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) + (-1 * n_������)
      Where ����id = v_��Ԥ��.����id And ���� = 1 And ���� = 1
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ������� (����id, ����, Ԥ�����, ����) Values (v_��Ԥ��.����id, 1, -1 * n_������, 1);
        n_����ֵ := -1 * n_������;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From �������
        Where ����id = v_��Ԥ��.����id And ���� = 1 And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
    
      Update ����Ԥ����¼ Set ��Ԥ�� = Nvl(��Ԥ��, 0) - n_������ Where ����id = ����id_In And ���㷽ʽ Is Null;
    
      If n_��Ԥ�� = 0 Then
        Exit;
      End If;
    End Loop;
    If n_��Ԥ�� <> 0 Then
      v_Err_Msg := '��ǰ�˿��������Ԥ������˽��˿�ʧ�ܣ�';
      Raise Err_Item;
    End If;
  End If;

  --0.��ͨ���㷽ʽ 
  If Nvl(��������_In, 0) = 0 Then
    --�����շѽ��� :��ʽΪ:���㷽ʽ|������|�������|����ժҪ||.. 
    v_�������� := ���㷽ʽ_In || '||';
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      v_������� := LTrim(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_����ժҪ := LTrim(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
    
      If Nvl(n_������, 0) <> 0 Then
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־,
           �����id, ���㿨���, ����, ������ˮ��, ����˵��, �������, ��������, �Ự��)
        Values
          (����Ԥ����¼_Id.Nextval, 6, r_Balance.No, r_Balance.��¼״̬, r_Balance.����id, Null, v_����ժҪ, v_���㷽ʽ, r_Balance.�Ǽ�ʱ��,
           r_Balance.����Ա���, r_Balance.����Ա����, n_������, r_Balance.����id, r_Balance.�ɿ���id, Sysdate, v_������Ա, r_Balance.�������, 2,
           Null, Null, ����_In, ������ˮ��_In, ����˵��_In, v_�������, 6, n_�Ự��);
      
        Update ����Ԥ����¼ Set ��Ԥ�� = Nvl(��Ԥ��, 0) - n_������ Where ����id = ����id_In And ���㷽ʽ Is Null;
      End If;
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  --1.���������� 
  If Nvl(��������_In, 0) = 1 Then
    v_��ǰ���� := ���㷽ʽ_In;
    v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
    v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
    n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
    v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
    v_������� := LTrim(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
    v_����ժҪ := LTrim(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
  
    If Nvl(n_������, 0) <> 0 Then
      Update ����Ԥ����¼
      Set ��Ԥ�� = Nvl(��Ԥ��, 0) + n_������
      Where ����id = ����id_In And ���㷽ʽ = v_���㷽ʽ And �����id = �����id_In
      Returning ID Into n_Ԥ��id;
      If Sql%NotFound Then
        Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־,
           �����id, ���㿨���, ����, ������ˮ��, ����˵��, �������, ��������, �Ự��)
        Values
          (n_Ԥ��id, 6, r_Balance.No, r_Balance.��¼״̬, r_Balance.����id, Null, v_����ժҪ, v_���㷽ʽ, r_Balance.�Ǽ�ʱ��,
           r_Balance.����Ա���, r_Balance.����Ա����, n_������, r_Balance.����id, r_Balance.�ɿ���id, Sysdate, v_������Ա, r_Balance.�������,
           У�Ա�־_In, �����id_In, Null, ����_In, ������ˮ��_In, ����˵��_In, v_�������, 6, n_�Ự��);
      End If;
      Update ����Ԥ����¼ Set ��Ԥ�� = Nvl(��Ԥ��, 0) - n_������ Where ����id = ����id_In And ���㷽ʽ Is Null;
    
      --��������������Ϣ����
      Zl_Custom_Balance_Update(n_Ԥ��id);
    End If;
  End If;

  --2.ҽ������ 
  If Nvl(��������_In, 0) = 2 Then
    --2.1����Ƿ��Ѿ�����ҽ����������,������ɾ�� 
    n_������ := 0;
    For v_ҽ�� In (Select ID, ���㷽ʽ, ��Ԥ��
                 From ����Ԥ����¼ A
                 Where ����id = ����id_In And Exists
                  (Select 1 From ���㷽ʽ Where a.���㷽ʽ = ���� And ���� In (3, 4)) And Mod(��¼����, 10) <> 1) Loop
      n_������ := n_������ + Nvl(v_ҽ��.��Ԥ��, 0);
      l_Ԥ��id.Extend;
      l_Ԥ��id(l_Ԥ��id.Count) := v_ҽ��.Id;
    End Loop;
  
    If l_Ԥ��id.Count <> 0 Then
      Forall I In 1 .. l_Ԥ��id.Count
        Delete ����Ԥ����¼ Where ID = l_Ԥ��id(I);
    End If;
  
    --��ɾ�����㷽ʽΪ�յļ�¼ 
    Delete ����Ԥ����¼ Where ����id = ����id_In And ���㷽ʽ Is Null;
    If ���㷽ʽ_In Is Not Null Then
      v_�������� := ���㷽ʽ_In || '||';
    End If;
    n_��Ԥ�� := 0;
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      n_������ := To_Number(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־, �����id,
         ���㿨���, ����, ������ˮ��, ����˵��, �������, ��������, �Ự��)
      Values
        (����Ԥ����¼_Id.Nextval, 6, r_Balance.No, r_Balance.��¼״̬, r_Balance.����id, Null, '���ս���', v_���㷽ʽ, r_Balance.�Ǽ�ʱ��,
         r_Balance.����Ա���, r_Balance.����Ա����, Nvl(n_������, 0), r_Balance.����id, r_Balance.�ɿ���id, Sysdate, v_������Ա,
         r_Balance.�������, 2, Null, Null, Null, Null, Null, Null, 6, n_�Ự��);
      n_��Ԥ�� := n_��Ԥ�� + Nvl(n_������, 0);
    
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  
    --������㷽ʽΪNULL 
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־, �����id,
       ���㿨���, ����, ������ˮ��, ����˵��, �������, ��������, �Ự��)
    Values
      (����Ԥ����¼_Id.Nextval, 6, r_Balance.No, r_Balance.��¼״̬, r_Balance.����id, Null, '', Null, r_Balance.�Ǽ�ʱ��,
       r_Balance.����Ա���, r_Balance.����Ա����, -1 * n_��Ԥ��, r_Balance.����id, r_Balance.�ɿ���id, Sysdate, v_������Ա, r_Balance.�������, 1,
       Null, Null, Null, Null, Null, Null, 6, n_�Ự��);
    --ҽ����ر�Ĵ��� 
    Update ���ս�����ϸ Set ��־ = 2 Where ����id = ����id_In;
  End If;

  If Nvl(��ɽ���_In, 0) = 0 Then
    Return;
  End If;

  If Nvl(��ɽ���_In, 0) = 2 Then
    --1.����У�Ա�־ 
    Update ����Ԥ����¼ Set У�Ա�־ = 0 Where NO = r_Balance.No;
    Update ���ò����¼ Set ����״̬ = 2 Where NO = r_Balance.No;
    If Sql%NotFound Then
      v_Err_Msg := 'δ�ҵ�ҽ�����������ݣ����ܱ����˽��������ϲ���!';
      Raise Err_Item;
    End If;
    Return;
  End If;

  Delete ����Ԥ����¼ Where ����id = ����id_In And Mod(��¼����, 10) <> 1 And ���㷽ʽ Is Null And Nvl(��Ԥ��, 0) = 0;
  If Sql%NotFound Then
    Select Count(1) Into n_Count From ����Ԥ����¼ A Where ����id = ����id_In And ���㷽ʽ Is Null;
    If n_Count <> 0 Then
      v_Err_Msg := '������δ�ɿ�����ݣ�������ɽ��㣡';
    Else
      v_Err_Msg := '������Ϣ���󣬿�����Ϊ����ԭ����ɽ�����Ϣ��������[���ղ������]�����½��㣡';
    End If;
    Raise Err_Item;
  End If;

  --1.�����쳣״̬ 
  Update ���ò����¼ Set ����״̬ = 0 Where ������� = r_Balance.�������;
  If Sql%NotFound Then
    v_Err_Msg := 'δ�ҵ�ҽ�����������ݣ����ܱ����˽������˷ѻ����ϲ���!';
    Raise Err_Item;
  End If;

  n_�Ƿ����Ʊ�� := �Ƿ����Ʊ��_In;
  If �Ƿ����Ʊ��_In Is Null Then
    Select Nvl(Max(����), 0) Into n_���� From ���ս����¼ Where ��¼id = ����id_In And ���� = 1;
    If Nvl(r_Balance.���ӱ�־, 0) = 1 Then
      n_�Ƿ����Ʊ�� := Zl_Fun_Isstarteinvoice(4, n_����);
    Else
      n_�Ƿ����Ʊ�� := Zl_Fun_Isstarteinvoice(1, n_����);
    End If;
  End If;

  --2.����У�Ա�־,�Ự�� 
  Update ����Ԥ����¼ Set У�Ա�־ = 0, �Ự�� = Null, �Ƿ����Ʊ�� = n_�Ƿ����Ʊ�� Where ����id = ����id_In;

  --3.������Ա�ɿ����� 
  For c_�ɿ� In (Select a.���㷽ʽ, a.����Ա����, Nvl(Sum(a.��Ԥ��), 0) As ��Ԥ��
               From ����Ԥ����¼ A
               Where a.������� = r_Balance.������� And Mod(a.��¼����, 10) <> 1
               Group By a.���㷽ʽ, a.����Ա����
               Having Nvl(Sum(a.��Ԥ��), 0) <> 0) Loop
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + c_�ɿ�.��Ԥ��
    Where �տ�Ա = c_�ɿ�.����Ա���� And ���� = 1 And ���㷽ʽ = c_�ɿ�.���㷽ʽ;
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ����
        (�տ�Ա, ���㷽ʽ, ����, ���)
      Values
        (c_�ɿ�.����Ա����, c_�ɿ�.���㷽ʽ, 1, c_�ɿ�.��Ԥ��);
    End If;
  End Loop;

  --��Ϣ���ɴ��� 
  --��������:1-�շѽ��㣬2-������� 
  --����ID:����id 
  b_Message.Zlhis_Charge_002(2, ����id_In);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���ò������_Modify;
/
Create Or Replace Procedure Zl_���ò������_����˷�
(
  ����id_In     In ���ò����¼.����id%Type,
  ���㷽ʽ_In   Varchar2,
  �����id_In   ����Ԥ����¼.�����id%Type := Null,
  ����_In       ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
  �����_In   ������ü�¼.ʵ�ս��%Type := Null,
  ��ɽ���_In   Number := 1,
  ��Ԥ��_In     ����Ԥ����¼.��Ԥ��%Type := Null,
  У�Ա�־_In   ����Ԥ����¼.У�Ա�־%Type := 0
) As
  ------------------------------------------------------------------------------------------------------------------------------
  --����:�������˷�
  --   ���㷽ʽ_IN:��ʽΪ:"���㷽ʽ|������|�������|����ժҪ" ;Ҳ������գ�������˿�ʱΪ�����տ�ʱΪ��
  --   �����������贫�뿨���ID_IN,����_IN,������ˮ��_IN,����˵��_In
  --   �����_In ��������ʱ,����
  --   ��ɽ���_In  1-��ɲ�������˷�;0-δ��ɲ�������˷�
  --   ��Ԥ��_In ��Ԥ�����˿�ʱΪ�����տ�ʱΪ��
  --   У�Ա�־_In  �������˷�ʱ��Ч
  ------------------------------------------------------------------------------------------------------------------------------
  v_����   ���㷽ʽ.����%Type;
  n_Count    Number(18);
  n_�Ự��   ����Ԥ����¼.�Ự��%Type; --��ʽ��SID+'_'+SERIAL#
  v_������Ա ����Ԥ����¼.������Ա%Type;

  v_��ǰ���� Varchar2(4000);
  v_�������� Varchar2(4000);
  v_���㷽ʽ ����Ԥ����¼.���㷽ʽ%Type;
  n_������ ����Ԥ����¼.��Ԥ��%Type;
  v_������� ����Ԥ����¼.�������%Type;
  v_����ժҪ ����Ԥ����¼.ժҪ%Type;

  n_Ԥ����� ����Ԥ����¼.��Ԥ��%Type;
  n_��Ԥ��   ����Ԥ����¼.��Ԥ��%Type;
  n_ʣ���� ����Ԥ����¼.��Ԥ��%Type;
  n_����ֵ   ����Ԥ����¼.��Ԥ��%Type;
  n_Ԥ��id   ����Ԥ����¼.Id%Type;
  n_Rowcount Number;
  n_Currrow  Number;

  n_�Ƿ����Ʊ�� ����Ԥ����¼.�Ƿ����Ʊ��%Type;

  Cursor c_Balance Is
    Select ��¼����, NO, ��¼״̬, ʵ��Ʊ��, ����id, �շѽ���id, ����״̬, ����Ա���, ����Ա����, �Ǽ�ʱ��, �ɿ���id, ����id, �������, ���ӱ�־
    From ���ò����¼ A
    Where ����id = ����id_In And ��¼���� = 1 And Rownum < 2;
  r_Balance c_Balance%RowType;

  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin
  Begin
    Select Sid || '_' || Serial# Into n_�Ự�� From V$session Where Audsid = Userenv('sessionid');
  Exception
    When Others Then
      n_�Ự�� := Null;
  End;
  v_������Ա := zl_UserName;

  Select Count(1) Into n_Count From ���ò����¼ Where ����id = ����id_In And ��¼���� = 1 And Rownum < 2;
  If n_Count = 0 Then
    v_Err_Msg := 'δ�ҵ�ҽ�����������ݣ����ܼ�������!';
    Raise Err_Item;
  End If;

  Open c_Balance;
  Fetch c_Balance
    Into r_Balance;

  If Nvl(�����_In, 0) <> 0 Then
    Begin
      Select ���� Into v_���� From ���㷽ʽ Where Nvl(����, 0) = 9;
    Exception
      When Others Then
        v_���� := '����';
    End;
  
    Update ����Ԥ����¼
    Set ��Ԥ�� = Nvl(��Ԥ��, 0) + Nvl(�����_In, 0)
    Where ����id = ����id_In And ���㷽ʽ = v_����;
    If Sql%NotFound Then
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־, �����id,
         ���㿨���, ����, ������ˮ��, ����˵��, �������, ��������, �Ự��)
      Values
        (����Ԥ����¼_Id.Nextval, 6, r_Balance.No, r_Balance.��¼״̬, r_Balance.����id, Null, Null, v_����, r_Balance.�Ǽ�ʱ��,
         r_Balance.����Ա���, r_Balance.����Ա����, �����_In, r_Balance.����id, r_Balance.�ɿ���id, Sysdate, v_������Ա, r_Balance.�������, 2,
         Null, Null, ����_In, ������ˮ��_In, ����˵��_In, Null, 6, n_�Ự��);
    End If;
  
    Update ����Ԥ����¼
    Set ��Ԥ�� = Nvl(��Ԥ��, 0) - Nvl(�����_In, 0)
    Where ���㷽ʽ Is Null And ����id = ����id_In;
    If Sql%NotFound Then
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־, �����id,
         ���㿨���, ����, ������ˮ��, ����˵��, �������, ��������, �Ự��)
      Values
        (����Ԥ����¼_Id.Nextval, 6, r_Balance.No, r_Balance.��¼״̬, r_Balance.����id, Null, Null, Null, r_Balance.�Ǽ�ʱ��,
         r_Balance.����Ա���, r_Balance.����Ա����, -1 * �����_In, r_Balance.����id, r_Balance.�ɿ���id, Sysdate, v_������Ա,
         r_Balance.�������, 2, Null, Null, Null, Null, Null, Null, 6, n_�Ự��);
    End If;
  End If;

  --�������תסԺʱ���˿�����ܴ���ʣ��δ�˽��
  --������㷽ʽΪNULL�ļ�¼�����˿������δ�˽��ʱ������ȫ���ӵ����һ����¼��
  Select Count(1) Into n_Rowcount From ����Ԥ����¼ Where ������� = r_Balance.������� And ���㷽ʽ Is Null;

  --��Ԥ����
  If Nvl(��Ԥ��_In, 0) < 0 Then
    n_Ԥ����� := -1 * ��Ԥ��_In;
    For v_��Ԥ�� In (Select Max(ID) As Ԥ��id, NO, ����id, Max(�տ�ʱ��) As �տ�ʱ��, Sum(Nvl(��Ԥ��, 0)) As ���
                  From ����Ԥ����¼
                  Where Mod(��¼����, 10) = 1 And Nvl(Ԥ�����, 0) = 1 And
                        ����id In
                        (
                         --���ý���ID
                         Select a.����id As ԭ����id
                         From ������ü�¼ A, ������ü�¼ B
                         Where a.��¼���� = b.��¼���� And a.No = b.No And a.��� = b.��� And b.��¼״̬ <> 2 And
                               b.����id In
                               (Select �շѽ���id From ���ò����¼ Where ��¼���� = 1 And ������� = r_Balance.�������)
                         Union All
                         --����������ID
                         Select ����id
                         From ����Ԥ����¼
                         Where ������� In (Select a.�������
                                        From ���ò����¼ A, ���ò����¼ B
                                        Where a.No = b.No And a.��¼���� = b.��¼���� And a.���ӱ�־ = b.���ӱ�־ And b.��¼���� = 1 And
                                              b.������� = r_Balance.�������))
                  Group By NO, ����id
                  Having Sum(Nvl(��Ԥ��, 0)) > 0
                  Order By �տ�ʱ�� Desc) Loop
    
      If v_��Ԥ��.��� - n_Ԥ����� < 0 Then
        n_������ := -1 * v_��Ԥ��.���;
        n_Ԥ����� := n_Ԥ����� - v_��Ԥ��.���;
      Else
        n_������ := -1 * n_Ԥ�����;
        n_Ԥ����� := 0;
      End If;
    
      n_ʣ���� := Nvl(n_������, 0);
      n_Currrow  := 0;
      --��Ҫ���ݡ�����������ȴ����տ�����ģ��ٴ����˿������
      For c_���� In (Select ����id, ��¼����, ����id, Nvl(Sum(��Ԥ��), 0) As ������
                   From ����Ԥ����¼
                   Where ������� = r_Balance.������� And ���㷽ʽ Is Null
                   Group By ����id, ����id, ��¼����
                   Order By ������ Desc) Loop
      
        n_Currrow := n_Currrow + 1;
        If c_����.������ < n_ʣ���� Then
          n_��Ԥ��   := n_ʣ����;
          n_ʣ���� := 0;
        Else
          n_��Ԥ��   := c_����.������;
          n_ʣ���� := n_ʣ���� - c_����.������;
        End If;
      
        --�˿������δ�˽��ʱ������ȫ���ӵ����һ����¼��
        If n_Currrow = n_Rowcount And n_ʣ���� <> 0 Then
          n_��Ԥ��   := n_��Ԥ�� + n_ʣ����;
          n_ʣ���� := 0;
        End If;
      
        If n_��Ԥ�� <> 0 Then
          Insert Into ����Ԥ����¼
            (ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������,
             У�Ա�־, �����id, ���㿨���, ����, ������ˮ��, ����˵��, �������, Ԥ�����, ��������, �Ự��)
            Select ����Ԥ����¼_Id.Nextval, 11, NO, ʵ��Ʊ��, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, r_Balance.�Ǽ�ʱ��, r_Balance.����Ա���,
                   r_Balance.����Ա����, n_��Ԥ��, c_����.����id, r_Balance.�ɿ���id, Sysdate, v_������Ա, r_Balance.�������, 2, �����id, ���㿨���,
                   ����, ������ˮ��, ����˵��, �������, Ԥ�����, Mod(c_����.��¼����, 10), n_�Ự��
            From ����Ԥ����¼
            Where ID = v_��Ԥ��.Ԥ��id;
        
          --����Ԥ���������
          Select Max(ID) Into n_Ԥ��id From ����Ԥ����¼ Where NO = v_��Ԥ��.No And ��¼���� = 1 And ��¼״̬ <> 2;
          If Nvl(n_Ԥ��id, 0) <> 0 Then
            Update Ԥ���������
            Set Ԥ����� = Nvl(Ԥ�����, 0) + (-1 * n_��Ԥ��)
            Where ����id = v_��Ԥ��.����id And Ԥ��id = n_Ԥ��id
            Returning Ԥ����� Into n_����ֵ;
            If Sql%RowCount = 0 Then
              Insert Into Ԥ���������
                (Ԥ��id, ����id, Ԥ�����, Ԥ�����)
              Values
                (n_Ԥ��id, v_��Ԥ��.����id, 1, -1 * n_��Ԥ��);
              n_����ֵ := -1 * n_��Ԥ��;
            End If;
            If Nvl(n_����ֵ, 0) = 0 Then
              Delete From Ԥ��������� Where Ԥ��id = n_Ԥ��id And Nvl(Ԥ�����, 0) = 0;
            End If;
          End If;
        
          --���²������
          Update �������
          Set Ԥ����� = Nvl(Ԥ�����, 0) + (-1 * n_��Ԥ��)
          Where ����id = v_��Ԥ��.����id And ���� = 1 And ���� = 1
          Returning Ԥ����� Into n_����ֵ;
          If Sql%RowCount = 0 Then
            Insert Into ������� (����id, ����, Ԥ�����, ����) Values (c_����.����id, 1, -1 * n_��Ԥ��, 1);
            n_����ֵ := -1 * n_��Ԥ��;
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From �������
            Where ����id = c_����.����id And ���� = 1 And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
          End If;
        
          Update ����Ԥ����¼ Set ��Ԥ�� = Nvl(��Ԥ��, 0) - n_��Ԥ�� Where ���㷽ʽ Is Null And ����id = c_����.����id;
        End If;
      
        If n_ʣ���� = 0 Then
          Exit;
        End If;
      End Loop;
      If n_ʣ���� <> 0 Then
        v_Err_Msg := '��ǰ�˿��������ʣ��δ�˽��˿�ʧ�ܣ�';
        Raise Err_Item;
      End If;
    
      If n_Ԥ����� = 0 Then
        Exit;
      End If;
    End Loop;
    If n_Ԥ����� <> 0 Then
      v_Err_Msg := '��ǰ�˿��������Ԥ������˽��˿�ʧ�ܣ�';
      Raise Err_Item;
    End If;
  End If;

  --�����շѽ��� :��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.."
  v_�������� := ���㷽ʽ_In || '||';
  While v_�������� Is Not Null Loop
    v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
    v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
    v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
    n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
    v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
    v_������� := LTrim(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
    v_����ժҪ := LTrim(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
  
    n_ʣ���� := Nvl(n_������, 0);
    n_Currrow  := 0;
    -- n_������ Ϊ������Ҫ���ݡ�����������ȴ����˿�����ģ��ٴ����տ������
    --����������ʹ���н��㷽ʽΪ�յļ�¼�ĳ�Ԥ�����Ϊ��
    For c_���� In (Select ����id, ��¼����, ��¼״̬, NO, Nvl(Sum(��Ԥ��), 0) As ������
                 From ����Ԥ����¼
                 Where ������� = r_Balance.������� And ���㷽ʽ Is Null
                 Group By ����id, ��¼����, NO, ��¼״̬
                 Order By -1 * ������) Loop
    
      n_Currrow := n_Currrow + 1;
      If c_����.������ < n_ʣ���� Then
        n_��Ԥ��   := n_ʣ����;
        n_ʣ���� := 0;
      Else
        n_��Ԥ��   := c_����.������;
        n_ʣ���� := n_ʣ���� - c_����.������;
      End If;
    
      --�˿������δ�˽��ʱ������ȫ���ӵ����һ����¼��
      If n_Currrow = n_Rowcount And n_ʣ���� <> 0 Then
        n_��Ԥ��   := n_��Ԥ�� + n_ʣ����;
        n_ʣ���� := 0;
      End If;
    
      If n_��Ԥ�� <> 0 Then
        Update ����Ԥ����¼
        Set ��Ԥ�� = Nvl(��Ԥ��, 0) + n_��Ԥ��
        Where ����id = c_����.����id And ���㷽ʽ = v_���㷽ʽ And �����id = �����id_In
        Returning ID Into n_Ԥ��id;
        If Sql%NotFound Then
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ����ʱ��, ������Ա, �������, У�Ա�־,
             �����id, ���㿨���, ����, ������ˮ��, ����˵��, �������, ��������, �Ự��)
          Values
            (n_Ԥ��id, c_����.��¼����, c_����.No, c_����.��¼״̬, r_Balance.����id, Null, v_����ժҪ, v_���㷽ʽ, r_Balance.�Ǽ�ʱ��,
             r_Balance.����Ա���, r_Balance.����Ա����, n_��Ԥ��, c_����.����id, r_Balance.�ɿ���id, Sysdate, v_������Ա, r_Balance.�������,
             Decode(Nvl(�����id_In, 0), 0, 2, У�Ա�־_In), �����id_In, Null, ����_In, ������ˮ��_In, ����˵��_In, v_�������,
             Mod(c_����.��¼����, 10), n_�Ự��);
        End If;
      
        Update ����Ԥ����¼ Set ��Ԥ�� = Nvl(��Ԥ��, 0) - n_��Ԥ�� Where ���㷽ʽ Is Null And ����id = c_����.����id;
      
        If Nvl(�����id_In, 0) <> 0 Then
          --��������������Ϣ����
          Zl_Custom_Balance_Update(n_Ԥ��id);
        End If;
      End If;
    
      If n_ʣ���� = 0 Then
        Exit;
      End If;
    End Loop;
    If n_ʣ���� <> 0 Then
      v_Err_Msg := '��ǰ�˿��������ʣ��δ�˽��˿�ʧ�ܣ�';
      Raise Err_Item;
    End If;
  
    v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
  End Loop;

  If Nvl(��ɽ���_In, 0) = 0 Then
    Return;
  End If;

  Delete From ����Ԥ����¼
  Where ������� = r_Balance.������� And ���㷽ʽ Is Null And Mod(��¼����, 10) <> 1 And Nvl(��Ԥ��, 0) = 0;
  If Sql%NotFound Then
    Select Count(1) Into n_Count From ����Ԥ����¼ A Where ����id = ����id_In And ���㷽ʽ Is Null;
    If n_Count <> 0 Then
      v_Err_Msg := '������δ�ɿ�����ݣ�������ɽ��㣡';
    Else
      v_Err_Msg := '������Ϣ���󣬿�����Ϊ����ԭ����ɽ�����Ϣ��������[���ղ������]�����½��㣡';
    End If;
    Raise Err_Item;
  End If;

  --һ�ν����ж������㷽ʽΪNULL�ļ�¼
  Select Count(1) Into n_Count From ����Ԥ����¼ A Where ������� = r_Balance.������� And ���㷽ʽ Is Null;
  If n_Count <> 0 Then
    v_Err_Msg := '������Ϣ��������[���ղ������]�����½��㣡';
    Raise Err_Item;
  End If;

  --1.�����쳣״̬
  Update ������ü�¼
  Set ����״̬ = 0
  Where Nvl(����״̬, 0) = 1 And ����id In (Select Distinct ����id From ����Ԥ����¼ Where ������� = r_Balance.�������);

  Update ���ò����¼ Set ����״̬ = 0 Where ������� = r_Balance.�������;
  If Sql%NotFound Then
    v_Err_Msg := 'δ�ҵ�ҽ�����������ݣ����ܲ����˽������˷ѻ����ϲ���!';
    Raise Err_Item;
  End If;

  --2.����У�Ա�־,�Ự��
  Update ����Ԥ����¼ Set У�Ա�־ = 0, �Ự�� = Null Where ������� = r_Balance.�������;

  --3.���� �Ƿ����Ʊ�� ��� 
  Select Max(a.�Ƿ����Ʊ��)
  Into n_�Ƿ����Ʊ��
  From ����Ԥ����¼ A,
       (Select ����id
         From (Select b.����id
                From ���ò����¼ B
                Where b.No = r_Balance.No And b.��¼���� = 1 And b.��¼״̬ In (1, 3)
                Order By b.�Ǽ�ʱ��)
         Where Rownum < 2) B
  Where a.����id = b.����id;

  Update ����Ԥ����¼ Set �Ƿ����Ʊ�� = n_�Ƿ����Ʊ�� Where ������� = r_Balance.������� And �������� = 6;

  --4.������Ա�ɿ�����
  For c_�ɿ� In (Select a.���㷽ʽ, a.����Ա����, Nvl(Sum(a.��Ԥ��), 0) As ��Ԥ��
               From ����Ԥ����¼ A
               Where a.������� = r_Balance.������� And Mod(a.��¼����, 10) <> 1
               Group By a.���㷽ʽ, a.����Ա����
               Having Nvl(Sum(a.��Ԥ��), 0) <> 0) Loop
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + c_�ɿ�.��Ԥ��
    Where �տ�Ա = c_�ɿ�.����Ա���� And ���� = 1 And ���㷽ʽ = c_�ɿ�.���㷽ʽ;
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ����
        (�տ�Ա, ���㷽ʽ, ����, ���)
      Values
        (c_�ɿ�.����Ա����, c_�ɿ�.���㷽ʽ, 1, c_�ɿ�.��Ԥ��);
    End If;
  End Loop;

  --��Ϣ���ɴ���
  For c_����id In (Select Distinct ����id From ����Ԥ����¼ Where ��¼���� = 3 And ������� = r_Balance.�������) Loop
    b_Message.Zlhis_Charge_004(2, c_����id.����id);
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���ò������_����˷�;
/

Create Or Replace Procedure Zl_����תסԺ_�շ�ת��
(
  No_In           סԺ���ü�¼.No%Type,
  ����Ա���_In   סԺ���ü�¼.����Ա���%Type,
  ����Ա����_In   סԺ���ü�¼.����Ա����%Type,
  �˷�ʱ��_In     סԺ���ü�¼.����ʱ��%Type,
  �����˷�_In     Number := 0,
  ��Ժ����id_In   סԺ���ü�¼.��������id%Type := Null,
  ��ҳid_In       סԺ���ü�¼.��ҳid%Type := Null,
  ���㷽ʽ_In     ����Ԥ����¼.���㷽ʽ%Type := Null,
  ����id_In       ����Ԥ����¼.����id%Type := Null,
  ԭ����id_In     ����Ԥ����¼.����id%Type := Null,
  ����_In       ����Ԥ����¼.��Ԥ��%Type := Null,
  �ɿ���id_In     ����Ԥ����¼.�ɿ���id%Type := Null,
  Ԥ������Ʊ��_In ����Ԥ����¼.Ԥ������Ʊ��%Type := 0
) As
  --���:
  --  �����˷�_In:0-����תסԺ��������;1-�����˷�ģʽ;=1ʱ:��Ժ����id_In����ҳID_IN���Բ�����
  --  �ɿ���ID_In:NULL��ʾ���ɿ���ID�������¶�ȡ);0-��ʾ�Ѿ���ȡ�������ٶ�ȡ;>0��ʾ�Ѿ���ȡ������Ľɿ��� 
  --  Ԥ������Ʊ��_In:Ԥ�����Ƿ����õ���Ʊ��
  n_Count      Number(5);
  n_ԭ����id   סԺ���ü�¼.����id%Type;
  n_ʵ�ս��   ������ü�¼.ʵ�ս��%Type;
  n_��id       ����ɿ����.Id%Type;
  n_����id     ������Ϣ.����id%Type;
  v_Ԥ��no     ����Ԥ����¼.No%Type;
  n_Ԥ�����   ����Ԥ����¼.��Ԥ��%Type;
  n_��ӡid     Ʊ��ʹ����ϸ.��ӡid%Type;
  n_��������id סԺ���ü�¼.��������id%Type;
  v_������     ������ü�¼.������%Type;
  n_����id     ������ü�¼.����id%Type;
  v_����     ���㷽ʽ.����%Type;
  n_����     ����Ԥ����¼.��Ԥ��%Type;
  n_����ֵ     �������.�������%Type;

  n_ʣ��Ԥ��     ����Ԥ����¼.��Ԥ��%Type;
  v_���㷽ʽ     ���㷽ʽ.����%Type;
  v_ȱʡ���㷽ʽ ���㷽ʽ.����%Type;
  v_Nos          Varchar2(3000);
  v_����ids      Varchar2(3000);
  v_ԭ����ids    Varchar2(3000);
  n_Tempid       ����Ԥ����¼.Id%Type;
  n_ҽ��         Number;
  n_����         Number;
  n_�����˷�     Number;
  n_�˷�����     Number;
  n_����״̬     ������ü�¼.����״̬%Type;
  n_��������id   ����Ԥ����¼.��������id%Type;
  n_��ֵid       ����Ԥ����¼.Id%Type;
  n_�Ƿ����ҽ�� Number(2);
  Err_Item Exception;
  v_Err_Msg Varchar2(200);

  Procedure ����Ԥ����_Del
  (
    ����id_In    ����Ԥ����¼.����id%Type,
    ����ids_In   Varchar2,
    ��Ԥ����_In  ����Ԥ����¼.��Ԥ��%Type,
    �˿�ϼ�_Out Out ����Ԥ����¼.��Ԥ��%Type
  ) As
    --��Ԥ����_In�������ʱ����ʾȫ��,������Ԥ���ʽ�����˿�
    n_ȫ��     Number(2);
    n_��Ԥ���� ����Ԥ����¼.��Ԥ��%Type;
    n_��Ԥ��   ����Ԥ����¼.��Ԥ��%Type;
  Begin
    n_ȫ��     := 1;
    n_��Ԥ���� := Nvl(��Ԥ����_In, 0);
    If Nvl(��Ԥ����_In, 0) <> 0 Then
      n_ȫ�� := 0;
    End If;
  
    �˿�ϼ�_Out := 0;
    For r_Prepay In (Select NO, Max(Decode(��¼����, 1, ʵ��Ʊ��, Null)) As ʵ��Ʊ��, ����id, ��ҳid, Max(����id) As ����id,
                            Max(���㷽ʽ) As ���㷽ʽ, Max(�������) As �������, Max(�ɿλ) As �ɿλ, Max(��λ������) As ��λ������,
                            Max(��λ�ʺ�) As ��λ�ʺ�, Min(�տ�ʱ��) As �տ�ʱ��, Sum(��Ԥ��) As ��Ԥ��, �����id, ���㿨���, Max(����) As ����, ������ˮ��,
                            Max(����˵��) As ����˵��, Max(������λ) As ������λ, ��������, Decode(Nvl(�����id, 0), 0, 0, ��������id) As ��������id,
                            Max(Ԥ�����) As Ԥ�����, Max(����ʱ��) As ����ʱ��, Max(������Ա) As ������Ա, Max(Decode(��¼����, 1, ID, 0)) As Ԥ��id
                     From ����Ԥ����¼ A
                     Where ��¼���� In (1, 11) And a.����id In (Select Column_Value From Table(f_Str2List(����ids_In))) And
                           Nvl(��Ԥ��, 0) <> 0
                     Group By NO, ����id, ��ҳid, �����id, ���㿨���, ������ˮ��, ��������, ��������id) Loop
    
      n_��Ԥ�� := Nvl(r_Prepay.��Ԥ��, 0);
      If n_ȫ�� = 0 Then
        If n_��Ԥ���� <> 0 Then
          If n_��Ԥ���� > n_��Ԥ�� Then
            n_��Ԥ���� := Round(n_��Ԥ���� - n_��Ԥ��, 6);
          Else
            n_��Ԥ��   := Nvl(n_��Ԥ����, 0);
            n_��Ԥ���� := 0;
          End If;
        Else
          Exit;
        End If;
      
      End If;
      Select ����Ԥ����¼_Id.Nextval Into n_Tempid From Dual;
    
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, Ԥ�����, ��������, ��������id, ����ʱ��, ������Ա)
        Select n_Tempid, r_Prepay.No, r_Prepay.ʵ��Ʊ��, 11, 1, r_Prepay.����id, r_Prepay.��ҳid, r_Prepay.����id, Null,
               r_Prepay.���㷽ʽ, r_Prepay.�������, Null, r_Prepay.�ɿλ, r_Prepay.��λ������, r_Prepay.��λ�ʺ�, �˷�ʱ��_In, ����Ա����_In,
               ����Ա���_In, -1 * n_��Ԥ��, ����id_In, n_��id, r_Prepay.�����id, r_Prepay.���㿨���, r_Prepay.����, r_Prepay.������ˮ��,
               r_Prepay.����˵��, r_Prepay.������λ, -1 * ����id_In, Nvl(r_Prepay.Ԥ�����, 1), r_Prepay.��������, r_Prepay.��������id,
               r_Prepay.����ʱ��, r_Prepay.������Ա
        From Dual;
    
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) + n_��Ԥ��
      Where ����id = r_Prepay.����id And ���� = Nvl(r_Prepay.Ԥ�����, 1) And ���� = 1
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, Ԥ�����, ����)
        Values
          (r_Prepay.����id, Nvl(r_Prepay.Ԥ�����, 1), n_��Ԥ��, 1);
        n_����ֵ := n_��Ԥ��;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From �������
        Where ����id = r_Prepay.����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
      End If;
    
      n_��ֵid := r_Prepay.Ԥ��id;
      If Nvl(n_��ֵid, 0) = 0 Then
        Select Max(ID) Into n_��ֵid From ����Ԥ����¼ Where NO = r_Prepay.No And ��¼���� = 1 And ��¼״̬ In (1, 3);
      End If;
      If Nvl(n_��ֵid, 0) <> 0 Then
        Update Ԥ���������
        Set Ԥ����� = Nvl(Ԥ�����, 0) + n_��Ԥ��
        Where ����id = r_Prepay.����id And Ԥ��id = n_��ֵid
        Returning Ԥ����� Into n_����ֵ;
      
        If Sql%RowCount = 0 Then
          Insert Into Ԥ���������
            (Ԥ��id, ����id, Ԥ�����, Ԥ�����)
          Values
            (n_��ֵid, r_Prepay.����id, Nvl(r_Prepay.Ԥ�����, 1), n_��Ԥ��);
          n_����ֵ := n_��Ԥ��;
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From Ԥ��������� Where Ԥ��id = n_��ֵid And Nvl(Ԥ�����, 0) = 0;
        End If;
      End If;
      �˿�ϼ�_Out := �˿�ϼ�_Out + n_��Ԥ��;
    End Loop;
  
  End ����Ԥ����_Del;

  ---------------------------------------------------------------------------------------------
  --���ɳ�Ԥ��������Ԥ����
  Procedure ���˽���_Strict
  (
    ����id_In       ����Ԥ����¼.����id%Type,
    ����ids_In      Varchar2,
    ʣ���_In       ����Ԥ����¼.��Ԥ��%Type := Null,
    �Ƿ�����Ԥ��_In Number := 1,
    ����ҽ��_In     Number := 1
  ) As
    --ʣ���_In�������ʱ����ʾȫ��,����ʣ�������˿�
    --�Ƿ�����Ԥ��_In:1-�����µ�סԺԤ��;0-�������µ�סԺԤ��
    --����ҽ��_In:1-��ʾ��ҽ�����г��������򲻳����ⲿ������
    n_�˿���   ����Ԥ����¼.��Ԥ��%Type;
    n_��Ԥ��     ����Ԥ����¼.��Ԥ��%Type;
    n_ȫ��       Number(2);
    n_��������id ����Ԥ����¼.��������id%Type;
  Begin
    n_ȫ��     := 1;
    n_�˿��� := Nvl(ʣ���_In, 0);
    If Nvl(ʣ���_In, 0) <> 0 Then
      n_ȫ�� := 0;
    End If;
  
    For r_Pay In (Select a.���㷽ʽ, Sum(a.��Ԥ��) As ��Ԥ��, 2 As Ԥ�����, a.�����id, a.���㿨���, a.����, a.������ˮ�� As ������ˮ��,
                         Max(Decode(a.��¼״̬, 2, Null, a.����˵��)) As ����˵��, a.������λ, b.����, a.��������id,
                         Max(Decode(a.��¼״̬, 2, 0, a.Id)) As Ԥ��id, Max(�������) As �������, Max(ժҪ) As ժҪ,
                         Sign(Sum(a.��Ԥ��)) As ��־
                  From ����Ԥ����¼ A, ���㷽ʽ B
                  Where a.��¼���� = 3 And a.��¼״̬ In (1, 2, 3) And
                        a.����id In (Select Column_Value From Table(f_Str2List(����ids_In))) And a.���㷽ʽ = b.����(+) And
                        a.���㷽ʽ Is Not Null
                  Group By a.���㷽ʽ, Ԥ�����, a.�����id, a.���㿨���, a.����, b.����, a.������ˮ��, a.������λ, a.��������id
                  Having Sum(a.��Ԥ��) <> 0
                  Order By �����id, ��־ Desc, ����) Loop
    
      If (Nvl(����ҽ��_In, 0) = 1) Or (Instr('34', Nvl(r_Pay.����, 0))) = 0 Or
         (����ҽ��_In = 0 And Instr('34', Nvl(r_Pay.����, 0)) > 0 And Nvl(r_Pay.�����id, 0) <> 0) Then
      
        n_��Ԥ�� := Nvl(r_Pay.��Ԥ��, 0);
      
        If n_ȫ�� = 0 Then
          If n_�˿��� <> 0 Then
            If n_�˿��� > n_��Ԥ�� Then
              n_�˿��� := Round(n_�˿��� - n_��Ԥ��, 6);
            Else
              n_��Ԥ��   := Nvl(n_�˿���, 0);
              n_�˿��� := 0;
            End If;
          Else
            Exit;
          End If;
        End If;
      
        --4.1����Ԥ����� (�����ڲ����˷ѵ����)
        --���е���,����������Ԥ�����
        --��Ϊ�տ�������ɿ�,������Ա�ɿ�����ޱ仯
        --һ��ͨ:ֻ�к���ҽ������ֽ��㷽ʽ�ģ��Ż���ýӿڴ���
        --1.һ��ͨ
        n_Count      := 0;
        n_��������id := r_Pay.��������id;
        If r_Pay.�����id Is Not Null Or Nvl(r_Pay.���㿨���, 0) <> 0 Then
          If Nvl(r_Pay.��������id, 0) = 0 And Nvl(r_Pay.�����id, 0) <> 0 Then
          
            n_��������id := r_Pay.Ԥ��id;
            Update ����Ԥ����¼
            Set ��������id = n_��������id
            Where ����id In (Select Column_Value From Table(f_Str2List(����ids_In))) And ��¼���� = 3 And ��¼״̬ In (2, 3) And
                  �����id = r_Pay.�����id;
          End If;
        
          --�����������ѿ�(���ѿ������ʱ����)
          If Instr('34', r_Pay.����) = 0 And Nvl(r_Pay.���㿨���, 0) = 0 Then
            --����������Ҫ�� ���Ƿ���ֽ��㷽ʽ������Ƕ��ֽ��㷽ʽ����Ҫ�����˿�ӿ�
            Select Count(Distinct ���㷽ʽ)
            Into n_Count
            From ����Ԥ����¼
            Where ����id In (Select Column_Value From Table(f_Str2List(����ids_In))) And �����id = r_Pay.�����id And
                  Nvl(��������id, 0) = Nvl(r_Pay.��������id, 0);
          Else
            --���ѿ��������˿����ҽ����:��Ҫ���ýӿ�
            n_Count := 2;
          End If;
        
          If n_Count > 1 Then
          
            --��Ҫ���ýӿ��˵ģ������ӱ�־��Ϊ1,�Ա��˿�
            Update ����Ԥ����¼
            Set ��Ԥ�� = ��Ԥ�� + (-1 * n_��Ԥ��)
            Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = ����id_In And ���㷽ʽ = r_Pay.���㷽ʽ And Nvl(��������id, 0) = Nvl(n_��������id, 0) And
                  Nvl(�����id, 0) = Nvl(r_Pay.�����id, 0) And Nvl(���㿨���, 0) = Nvl(r_Pay.���㿨���, 0);
          
            If Sql%RowCount = 0 Then
              Select ����Ԥ����¼_Id.Nextval Into n_Tempid From Dual;
              Insert Into ����Ԥ����¼
                (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
                 �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������, ��������id, ������Ա, ����ʱ��, ���ӱ�־)
              Values
                (n_Tempid, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * n_��Ԥ��, r_Pay.���㷽ʽ, Null, �˷�ʱ��_In, Null,
                 Null, Null, ����Ա���_In, ����Ա����_In, '', n_��id, r_Pay.�����id, r_Pay.���㿨���, r_Pay.����, r_Pay.������ˮ��, r_Pay.����˵��,
                 r_Pay.������λ, ����id_In, -1 * ����id_In, 1, 3, Decode(n_��������id, 0, Null, n_��������id), ����Ա����_In, �˷�ʱ��_In, -1);
            
              --��������������Ϣ����
              Zl_Custom_Balance_Update(n_Tempid);
            End If;
          
            n_Count    := 2;
            n_����״̬ := 1;
          End If;
        End If;
      
        If n_Count <= 1 Then
          --������Ԥ����
          v_���㷽ʽ := Nvl(r_Pay.���㷽ʽ, v_ȱʡ���㷽ʽ);
        
          --ҽ�������Ѳ�����Ԥ����
          If Instr('349', r_Pay.����) = 0 And Nvl(�Ƿ�����Ԥ��_In, 0) = 1 And n_��Ԥ�� <> 0 Then
            --һ��ͨ��ÿһ�ʶ�����һ��Ԥ�����¼
            --������ͬһ�ֽ��㷽ʽֻ����һ��Ԥ�����¼
            Update ����Ԥ����¼
            Set ��� = Nvl(���, 0) + n_��Ԥ��
            Where ��¼���� = 1 And ��¼״̬ = 1 And �տ�ʱ�� = �˷�ʱ��_In And ����id + 0 = n_����id And ���㷽ʽ = v_���㷽ʽ And
                  (Nvl(�����id, 0) = 0 And Nvl(r_Pay.�����id, 0) = 0)
            Returning ID Into n_��ֵid;
            If Sql%RowCount = 0 Then
              v_Ԥ��no := Nextno(11);
              Select ����Ԥ����¼_Id.Nextval Into n_Tempid From Dual;
              Insert Into ����Ԥ����¼
                (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ, �ɿ���id,
                 Ԥ�����, �����id, ��������id, ������Ա, ����ʱ��, ����, ����˵��, ������ˮ��, �������, Ԥ������Ʊ��)
              Values
                (n_Tempid, v_Ԥ��no, Null, 1, 1, n_����id, ��ҳid_In, ��Ժ����id_In, n_��Ԥ��, v_���㷽ʽ, �˷�ʱ��_In, Null, Null, Null,
                 ����Ա���_In, ����Ա����_In, '����תסԺԤ��', n_��id, r_Pay.Ԥ�����, r_Pay.�����id, Nvl(r_Pay.��������id, n_Tempid), ����Ա����_In,
                 �˷�ʱ��_In, r_Pay.����, r_Pay.����˵��, r_Pay.������ˮ��, r_Pay.�������, Ԥ������Ʊ��_In);
              n_��ֵid := n_Tempid;
            End If;
            n_����״̬ := 1;
          
            --�������
            Update �������
            Set Ԥ����� = Nvl(Ԥ�����, 0) + n_��Ԥ��
            Where ���� = 1 And ����id = n_����id And ���� = 2
            Returning Ԥ����� Into n_����ֵ;
            If Sql%RowCount = 0 Then
              Insert Into ������� (����id, ����, ����, Ԥ�����, �������) Values (n_����id, 1, 2, n_��Ԥ��, 0);
              n_����ֵ := n_��Ԥ��;
            End If;
            If Nvl(n_����ֵ, 0) = 0 Then
              Delete From �������
              Where ����id = n_����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
            End If;
          
            If Nvl(n_��ֵid, 0) <> 0 Then
              Update Ԥ���������
              Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(n_��Ԥ��, 0)
              Where ����id = n_����id And Ԥ��id = n_��ֵid
              Returning Ԥ����� Into n_����ֵ;
            
              If Sql%RowCount = 0 Then
                Insert Into Ԥ���������
                  (Ԥ��id, ����id, Ԥ�����, Ԥ�����)
                Values
                  (n_��ֵid, n_����id, 2, Nvl(n_��Ԥ��, 0));
                n_����ֵ := Nvl(n_��Ԥ��, 0);
              End If;
              If Nvl(n_����ֵ, 0) = 0 Then
                Delete From Ԥ��������� Where Ԥ��id = n_��ֵid And Nvl(Ԥ�����, 0) = 0;
              End If;
            End If;
          End If;
        
          Update ����Ԥ����¼
          Set ��Ԥ�� = ��Ԥ�� + (-1 * n_��Ԥ��), ��������id = Nvl(��������id, n_��������id)
          Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = ����id_In And ���㷽ʽ = v_���㷽ʽ And Nvl(��������id, 0) = Nvl(n_��������id, 0) And
                Nvl(�����id, 0) = Nvl(r_Pay.�����id, 0) And Nvl(���㿨���, 0) = Nvl(r_Pay.���㿨���, 0);
          If Sql%RowCount = 0 Then
            Select ����Ԥ����¼_Id.Nextval Into n_Tempid From Dual;
          
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
               �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������, ��������id, ������Ա, ����ʱ��)
            Values
              (n_Tempid, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * n_��Ԥ��, v_���㷽ʽ, Null, �˷�ʱ��_In, Null, Null,
               Null, ����Ա���_In, ����Ա����_In, '', n_��id, r_Pay.�����id, r_Pay.���㿨���, r_Pay.����, r_Pay.������ˮ��, r_Pay.����˵��,
               r_Pay.������λ, ����id_In, -1 * ����id_In, 1, 3, Decode(n_��������id, 0, Null, n_��������id), ����Ա����_In, �˷�ʱ��_In);
          End If;
        
          --4.2�ɿ����ݴ���
          --   ��Ϊû��ʵ���ղ��˵�Ǯ,���Բ�����
          --�����˷��������ԭԤ����¼
          If r_Pay.���� In (3, 4) Then
            Update ��Ա�ɿ����
            Set ��� = Nvl(���, 0) - n_��Ԥ��
            Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_Pay.���㷽ʽ
            Returning ��� Into n_����ֵ;
            If Sql%RowCount = 0 Then
              Insert Into ��Ա�ɿ����
                (�տ�Ա, ���㷽ʽ, ����, ���)
              Values
                (����Ա����_In, r_Pay.���㷽ʽ, 1, -1 * n_��Ԥ��);
              n_����ֵ := n_��Ԥ��;
            End If;
            If Nvl(n_����ֵ, 0) = 0 Then
              Delete From ��Ա�ɿ����
              Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_Pay.���㷽ʽ And Nvl(���, 0) = 0;
            End If;
          End If;
        End If;
      End If;
    End Loop;
  
  End ���˽���_Strict;

  Procedure ���ü�¼_Strict
  (
    ����id_In   ������ü�¼.����id%Type,
    Nos_In      Varchar2,
    �ɿ���id_In ������ü�¼.�ɿ���id%Type
  ) As
  
  Begin
  
    --���·�����˼�¼
    Update ������˼�¼
    Set ��¼״̬ = 2
    Where ����id In (Select a.Id
                   From ������ü�¼ A
                   Where a.No In (Select Column_Value From Table(f_Str2List(Nos_In))) And Mod(a.��¼����, 10) = 1 And
                         a.��¼״̬ In (1, 3)) And ���� = 1;
    --���������¼
    Update ������ü�¼
    Set ��¼״̬ = 3
    Where Mod(��¼����, 10) = 1 And ��¼״̬ = 1 And NO In (Select Column_Value From Table(f_Str2List(Nos_In)));
  
    For r_Clinic In (Select Min(Mod(a.��¼����, 10)) As ��¼����, a.No, a.���, a.��������, a.�۸񸸺�, a.ҽ�����, a.����id, a.����, a.�Ա�, a.����,
                            a.���˿���id, a.�ѱ�, a.�շ����, a.�շ�ϸĿid, a.���㵥λ, a.������Ŀ��, a.���մ���id, a.���ձ���, a.��������, a.��ҩ����, a.����,
                            Sum(a.����) As ����, a.�Ӱ��־, a.���ӱ�־, a.������Ŀid, a.�վݷ�Ŀ, a.��׼����, Sum(a.Ӧ�ս��) As Ӧ�ս��,
                            Sum(a.ʵ�ս��) As ʵ�ս��, Sum(a.ͳ����) As ͳ����, a.��������id, a.������, a.ִ�в���id, a.������,
                            Max(a.���ʵ�id) As ���ʵ�id, Max(a.�Ƿ���) As �Ƿ���, a.����ʱ��, Min(a.ʵ��Ʊ��) As ʵ��Ʊ��,
                            Nvl(Min(Decode(a.��¼״̬, 2, a.ִ��״̬, 0)), 0) - 1 As ִ��״̬, �Һ�id, ��ҳid, ���˲���id
                     From ������ü�¼ A
                     Where a.No In (Select Column_Value From Table(f_Str2List(Nos_In))) And Mod(a.��¼����, 10) = 1 And
                           Nvl(a.���ӱ�־, 0) Not In (8, 9)
                     Group By a.No, a.���, a.��������, a.�۸񸸺�, a.ҽ�����, a.����id, a.����, a.�Ա�, a.����, a.���˿���id, a.�ѱ�, a.�շ����,
                              a.�շ�ϸĿid, a.���㵥λ, a.������Ŀ��, a.���մ���id, a.���ձ���, a.��������, a.��ҩ����, a.����, a.�Ӱ��־, a.���ӱ�־, a.������Ŀid,
                              a.�վݷ�Ŀ, a.��׼����, a.��������id, a.������, a.ִ�в���id, a.������, a.����ʱ��, �Һ�id, ��ҳid, ���˲���id
                     Having Sum(a.����) <> 0) Loop
      Insert Into ������ü�¼
        (ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, �����־, ҽ�����, ����id, ��ʶ��, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ,
         ������Ŀ��, ���մ���id, ���ձ���, ��������, ��ҩ����, ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����, ���ʷ���, ��������id, ������,
         ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ������, ����Ա���, ����Ա����, ���ʵ�id, ժҪ, �Ƿ���, �ɿ���id, ����id, ���ʽ��, ִ��״̬, ����״̬, �Һ�id, ��ҳid, ���˲���id)
      Values
        (���˷��ü�¼_Id.Nextval, r_Clinic.��¼����, r_Clinic.No, r_Clinic.ʵ��Ʊ��, 2, r_Clinic.���, r_Clinic.��������, r_Clinic.�۸񸸺�, 1,
         r_Clinic.ҽ�����, r_Clinic.����id, '', r_Clinic.����, r_Clinic.�Ա�, r_Clinic.����, r_Clinic.���˿���id, r_Clinic.�ѱ�,
         r_Clinic.�շ����, r_Clinic.�շ�ϸĿid, r_Clinic.���㵥λ, r_Clinic.������Ŀ��, r_Clinic.���մ���id, r_Clinic.���ձ���, r_Clinic.��������,
         r_Clinic.��ҩ����, r_Clinic.����, -1 * r_Clinic.����, r_Clinic.�Ӱ��־, r_Clinic.���ӱ�־, r_Clinic.������Ŀid, r_Clinic.�վݷ�Ŀ,
         r_Clinic.��׼����, -1 * r_Clinic.Ӧ�ս��, -1 * r_Clinic.ʵ�ս��, -1 * r_Clinic.ͳ����, 0, r_Clinic.��������id, r_Clinic.������,
         r_Clinic.����ʱ��, �˷�ʱ��_In, r_Clinic.ִ�в���id, r_Clinic.������, ����Ա���_In, ����Ա����_In, r_Clinic.���ʵ�id, '', r_Clinic.�Ƿ���,
         �ɿ���id_In, ����id_In, -1 * r_Clinic.ʵ�ս��, r_Clinic.ִ��״̬, 0, r_Clinic.�Һ�id, r_Clinic.��ҳid, r_Clinic.���˲���id);
    End Loop;
  End ���ü�¼_Strict;

  Procedure ҽ��������ϸ_Strict
  (
    ����id_In    ����Ԥ����¼.����id%Type,
    ����ids_In   Varchar2,
    Nos_In       Varchar2,
    �ɿ���id_In  ������ü�¼.�ɿ���id%Type,
    �˿�ϼ�_Out Out ����Ԥ����¼.��Ԥ��%Type,
    ������_In    Number := 0
  ) As
    --������_In:1-ֻ����ҽ����ϸ����;0-���˳�������Ҫ������Ԥ����¼
    v_���� ����Ԥ����¼.����%Type;
  Begin
    --ҽ���˿�
    �˿�ϼ�_Out := 0;
    For r_ҽ�� In (Select ����id, NO, ���㷽ʽ, ���, ��ע, �����id, ��������id, ������ˮ��, ����˵��
                 From ҽ��������ϸ A
                 Where Instr(',' || Nos_In || ',', ',' || NO || ',') > 0 And
                       ����id In (Select Column_Value From Table(f_Str2List(����ids_In)))) Loop
    
      If Nvl(������_In, 0) = 0 Then
      
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) - r_ҽ��.���
        Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_ҽ��.���㷽ʽ
        Returning ��� Into n_����ֵ;
      
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ����
            (�տ�Ա, ���㷽ʽ, ����, ���)
          Values
            (����Ա����_In, r_ҽ��.���㷽ʽ, 1, -1 * r_ҽ��.���);
          n_����ֵ := r_ҽ��.���;
        End If;
      
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From ��Ա�ɿ����
          Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_ҽ��.���㷽ʽ And Nvl(���, 0) = 0;
        End If;
      
        Select Max(��������id), Max(����)
        Into n_��������id, v_����
        From ����Ԥ����¼
        Where ����id = r_ҽ��.����id And ���㷽ʽ = r_ҽ��.���㷽ʽ And Nvl(�����id, 0) = Nvl(r_ҽ��.�����id, 0);
      
        Update ����Ԥ����¼
        Set ��Ԥ�� = ��Ԥ�� + (-1 * r_ҽ��.���)
        Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = ����id_In And ���㷽ʽ = r_ҽ��.���㷽ʽ And Nvl(�����id, 0) = Nvl(r_ҽ��.�����id, 0) And
              Nvl(��������id, 0) = Nvl(n_��������id, 0);
        If Sql%RowCount = 0 Then
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
             �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������, ��������id, ����ʱ��, ������Ա, ���ӱ�־)
          Values
            (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * r_ҽ��.���, r_ҽ��.���㷽ʽ, Null, �˷�ʱ��_In,
             Null, Null, Null, ����Ա���_In, ����Ա����_In, r_ҽ��.��ע, �ɿ���id_In, r_ҽ��.�����id, Null, v_����, r_ҽ��.������ˮ��, r_ҽ��.����˵��,
             Null, ����id_In, -1 * ����id_In, Decode(Nvl(r_ҽ��.�����id, 0), 0, 0, 1), 3, n_��������id, �˷�ʱ��_In, ����Ա����_In,
             Decode(Nvl(r_ҽ��.�����id, 0), 0, Null, -1));
        End If;
      
        Update ����Ԥ����¼
        Set ��¼״̬ = 3
        Where ��¼���� = 3 And ��¼״̬ = 1 And ����id In (Select Column_Value From Table(f_Str2List(v_ԭ����ids))) And
              ���㷽ʽ = r_ҽ��.���㷽ʽ;
      End If;
    
      Update ҽ��������ϸ
      Set ��� = ��� + (-1 * r_ҽ��.���)
      Where NO = r_ҽ��.No And ����id = ����id_In And ���㷽ʽ = r_ҽ��.���㷽ʽ And Nvl(�����id, 0) = Nvl(r_ҽ��.�����id, 0) And
            Nvl(��������id, 0) = Nvl(r_ҽ��.��������id, 0);
    
      If Sql%RowCount = 0 Then
        Insert Into ҽ��������ϸ
          (����id, NO, ���㷽ʽ, ���, �����id, ��������id, ������ˮ��, ����˵��)
        Values
          (����id_In, r_ҽ��.No, r_ҽ��.���㷽ʽ, -1 * r_ҽ��.���, r_ҽ��.�����id, r_ҽ��.��������id, r_ҽ��.������ˮ��, r_ҽ��.����˵��);
      End If;
      �˿�ϼ�_Out := �˿�ϼ�_Out + Nvl(r_ҽ��.���, 0);
    End Loop;
  End ҽ��������ϸ_Strict;

  Procedure �������Ʊ��_����(Nos_In Varchar2) As
    n_��ӡid Ʊ�ݴ�ӡ����.Id%Type;
  Begin
    --2.Ʊ���ջ�
    --������ǰû�д�ӡ,���ջ�
    For r_Nos In (Select Column_Value As NO From Table(f_Str2List(Nos_In))) Loop
    
      Select Nvl(Max(ID), 0)
      Into n_��ӡid
      From (Select b.Id
             From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
             Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And b.�������� = 1 And b.No = r_Nos.No
             Order By a.ʹ��ʱ�� Desc)
      Where Rownum < 2;
      If n_��ӡid > 0 Then
        --���ŵ���ѭ������ʱֻ���ջ�һ��
        Select Count(��ӡid) Into n_Count From Ʊ��ʹ����ϸ Where Ʊ�� = 1 And ���� = 2 And ��ӡid = n_��ӡid;
        If n_Count = 0 Then
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
            Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, �˷�ʱ��_In, ����Ա����_In, Ʊ�ݽ��
            From Ʊ��ʹ����ϸ
            Where ��ӡid = n_��ӡid And Ʊ�� = 1 And ���� = 1;
        End If;
      End If;
    End Loop;
  End �������Ʊ��_����;

  Procedure ����תס����_Over
  (
    ����id_In   ����Ԥ����¼.����id%Type,
    ԭ����id_In ����Ԥ����¼.����id%Type := Null
  ) As
    n_����״̬ Number(2);
  Begin
  
    Delete From ����Ԥ����¼ Where ����id = Nvl(ԭ����id_In, 0) And ժҪ = 'Ԥ����ʱ��¼' And ��¼���� = 3;
  
    Delete From ����Ԥ����¼
    Where ����id = ����id_In And ��¼���� = 3 And ��¼״̬ = 2 And ��Ԥ�� = 0 And ���㷽ʽ Is Not Null;
  
    --�Ϸ��Լ�飬��Ԥ��Ҫ����ý���ϼ�һ�£���һ��ʱ��ֱ���˳�
    Select Nvl(Max(1), 0) Into n_����״̬ From ����Ԥ����¼ Where ����id = ����id_In And ���ӱ�־ = -1;
  
    Update ������ü�¼ Set ����״̬ = Nvl(n_����״̬, 0) Where ����id = ����id_In;
  
    If Nvl(n_����״̬, 0) = 0 Then
      --�������쳣��ֱ�Ӹ���
      Update ����Ԥ����¼ Set У�Ա�־ = 0, ���ӱ�־ = Null Where ����id = ����id_In;
    Else
      Update ����Ԥ����¼ A
      Set У�Ա�־ = 2
      Where ����id = ����id_In And
            ((Nvl(���ӱ�־, 0) <> -1 And �����id Is Not Null) Or
            (Exists (Select 1 From ���㷽ʽ Where a.���㷽ʽ = ���� And ���� In (3, 4) And �����id Is Null)));
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ����id, ��ҳid, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, У�Ա�־, ��������, �������)
        Select ����Ԥ����¼_Id.Nextval, 3, Null, 2, n_����id, Null, Null, Null, �˷�ʱ��_In, ����Ա���_In, ����Ա����_In, Null, ����id_In,
               n_��id, 0, 3, -1 * ����id_In
        From Dual;
    End If;
  End ����תס����_Over;

Begin
  n_��id := �ɿ���id_In;
  If n_��id Is Null Then
    n_��id := Zl_Get��id(����Ա����_In);
  End If;
  If Nvl(n_��id, 0) = 0 Then
    n_��id := Null;
  End If;

  --����
  Select Max(����) Into v_���� From ���㷽ʽ Where ���� = 9 And Rownum < 2;
  If v_���� Is Null Then
    v_Err_Msg := 'û�з��������㷽ʽ�������Ƿ���ȷ���ã�';
    Raise Err_Item;
  End If;

  n_����       := ����_In;
  v_ȱʡ���㷽ʽ := ���㷽ʽ_In;
  If v_ȱʡ���㷽ʽ Is Null Then
    Select Nvl(Max(����), '�ֽ�') Into v_ȱʡ���㷽ʽ From ���㷽ʽ Where ���� = 1 And Rownum < 2;
  End If;

  If ԭ����id_In Is Null Then
  
    Select Count(NO), Sum(ʵ�ս��)
    Into n_Count, n_ʵ�ս��
    From ������ü�¼
    Where NO = No_In And Mod(��¼����, 10) = 1;
  
    If n_Count = 0 Or n_ʵ�ս�� = 0 Then
      v_Err_Msg := '����' || No_In || '�����շѵ��ݻ��򲢷�ԭ�����˲����˸õ���,����תΪסԺ����.';
      Raise Err_Item;
    End If;
  
    Select ����id, ����id, ��������id, ������
    Into n_ԭ����id, n_����id, n_��������id, v_������
    From ������ü�¼
    Where NO = No_In And Mod(��¼����, 10) = 1 And ��¼״̬ In (1, 3) And Rownum < 2;
  
    --1.1���Ϸ��ü�¼
    n_����id := ����id_In;
    If n_����id Is Null Then
      Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
    End If;
  
    ���ü�¼_Strict(n_����id, No_In, n_��id);
  
    --1.2����Ԥ����¼
    --���ϳ�Ԥ������
    For r_����id In (Select Distinct ����id
                   From ������ü�¼
                   Where NO In (Select Distinct NO
                                From ������ü�¼
                                Where ����id In (Select ����id
                                               From ����Ԥ����¼
                                               Where ������� In (Select b.�������
                                                              From ������ü�¼ A, ����Ԥ����¼ B
                                                              Where a.No = No_In And b.������� < 0 And Mod(a.��¼����, 10) = 1 And
                                                                    a.��¼״̬ <> 0 And a.����id = b.����id))) And
                         Mod(��¼����, 10) = 1 And ��¼״̬ <> 0
                   Union
                   Select Distinct ����id
                   From ������ü�¼
                   Where NO In (Select Distinct NO
                                From ������ü�¼
                                Where ����id In (Select a.����id
                                               From ������ü�¼ A, ����Ԥ����¼ B
                                               Where a.No = No_In And b.������� > 0 And Mod(a.��¼����, 10) = 1 And a.��¼״̬ <> 0 And
                                                     a.����id = b.����id)) And Mod(��¼����, 10) = 1 And ��¼״̬ <> 0) Loop
      v_ԭ����ids := v_ԭ����ids || ',' || r_����id.����id;
    End Loop;
  
    v_ԭ����ids := Substr(v_ԭ����ids, 2);
  
    Select Nvl(Max(1), 0)
    Into n_ҽ��
    From ���ս����¼
    Where ��¼id In (Select Column_Value From Table(f_Str2List(v_ԭ����ids))) And Rownum < 2 And �����id Is Not Null;
  
    If n_ҽ�� = 1 Then
    
      Select Nvl(Max(1), 0)
      Into n_����
      From ҽ��������ϸ
      Where NO = No_In And ����id In (Select Column_Value From Table(f_Str2List(v_ԭ����ids))) And Rownum < 2 And
            �����id Is Not Null;
    
      If n_���� = 0 Then
        v_Err_Msg := '��ǰ����' || No_In || '������ҽ��������ϸ,�޷���������תסԺ!';
        Raise Err_Item;
      End If;
    
    End If;
  
    --�Ƚ������Ѵ���
    n_����ֵ := Round(n_ʵ�ս��, 2) - Nvl(n_ʵ�ս��, 0);
    If Nvl(n_����ֵ, 0) <> 0 Then
      n_���� := Nvl(n_����, 0) + n_����ֵ;
    End If;
    n_ʵ�ս�� := Round(n_ʵ�ս��, 2);
  
    --ҽ����ϸ����: ����id_In, ����ids_In,Nos_In,�ɿ���id_In �˿�ϼ�_In out
    ҽ��������ϸ_Strict(n_����id, v_ԭ����ids, No_In, n_��id, n_Ԥ�����);
    n_ʵ�ս�� := n_ʵ�ս�� - Nvl(n_Ԥ�����, 0);
  
    If n_ʵ�ս�� <> 0 Then
      --��Ԥ����:
      ����Ԥ����_Del(n_����id, v_ԭ����ids, n_ʵ�ս��, n_Ԥ�����);
      n_ʵ�ս�� := n_ʵ�ս�� - Nvl(n_Ԥ�����, 0);
    End If;
  
    --2.Ʊ���ջ�
    �������Ʊ��_����(No_In);
  
    --3.�ɿ����ݴ���(
    --   �����������:
    --    1. ת������ֱ�����ʵ�,��ɿ����ݲ�����;
    --    2. ��ת��,�ٵ������˿���Ʊ,����Ҫ���нɿ����ݴ���
  
    If Nvl(�����˷�_In, 0) = 1 Or n_ʵ�ս�� <> 0 Then
      If Nvl(�����˷�_In, 0) = 1 Then
        --������Ԥ��:�˿�תԤ��(������Ʊ��,�ɲ���Աͨ���ش����)
        ���˽���_Strict(n_����id, v_ԭ����ids, n_ʵ�ս��, 0, 0);
      Elsif n_ʵ�ս�� <> 0 Then
        --����Ԥ��:
        ���˽���_Strict(n_����id, v_ԭ����ids, n_ʵ�ս��, 1, 0);
      End If;
    End If;
  
    If n_���� Is Not Null Then
      Update ����Ԥ����¼ Set ��Ԥ�� = Nvl(��Ԥ��, 0) + n_���� Where ����id = n_����id And ���㷽ʽ = v_����;
      If Sql%NotFound Then
        Select ����Ԥ����¼_Id.Nextval Into n_Tempid From Dual;
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
           �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������, ��������id)
        Values
          (n_Tempid, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, n_����, v_����, Null, �˷�ʱ��_In, Null, Null, Null,
           ����Ա���_In, ����Ա����_In, '', n_��id, Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 0, 3, n_Tempid);
      End If;
    End If;
  
    ����תס����_Over(n_����id, n_ԭ����id);
    Return;
  End If;

  --ҽ��������ת��
  For r_Nos In (Select Distinct a.No
                From ������ü�¼ A
                Where Mod(a.��¼����, 10) = 1 And a.��¼״̬ In (1, 3) And a.����id = ԭ����id_In) Loop
    v_Nos := v_Nos || ',' || r_Nos.No;
  End Loop;
  v_Nos := Substr(v_Nos, 2);

  For r_����ids In (Select Distinct a.����id
                  From ������ü�¼ A
                  Where a.No In (Select Column_Value From Table(f_Str2List(v_Nos))) And Mod(a.��¼����, 10) = 1 And
                        a.��¼״̬ <> 0) Loop
    v_����ids := v_����ids || ',' || r_����ids.����id;
  End Loop;

  v_����ids := Substr(v_����ids, 2);
  Select Count(a.No), Sum(a.ʵ�ս��)
  Into n_Count, n_ʵ�ս��
  From ������ü�¼ A
  Where a.No In (Select Column_Value From Table(f_Str2List(v_Nos))) And Mod(a.��¼����, 10) = 1;

  If n_Count = 0 Or n_ʵ�ս�� = 0 Then
    v_Err_Msg := '���ν��㲻���շѻ��򲢷�ԭ�����˲����˸ý���,����תΪסԺ����.';
    Raise Err_Item;
  End If;

  Select ����id, ����id, ��������id, ������
  Into n_ԭ����id, n_����id, n_��������id, v_������
  From ������ü�¼
  Where ����id = ԭ����id_In And Mod(��¼����, 10) = 1 And ��¼״̬ In (1, 3) And Rownum < 2;

  Select Nvl(Max(1), 0)
  Into n_�����˷�
  From ������ü�¼ A
  Where Mod(a.��¼����, 10) = 1 And a.��¼״̬ = 2 And a.����id In (Select Column_Value From Table(f_Str2List(v_����ids))) And
        Rownum < 2;

  Begin
    Select 0
    Into n_�����˷�
    From ������ü�¼ A
    Where ��¼���� = 11 And a.����id In (Select Column_Value From Table(f_Str2List(v_����ids))) And Rownum < 2;
  Exception
    When Others Then
      Null;
  End;

  n_�˷����� := 0;
  --ֻ�д��ڲ�����ʱ������Ҫ�� ����㷽ʽ�к��˶�����������Ϣ
  Select Count(*) - Max(Decode(����, 1, 1, 0)) As ͳ������, Sum(Decode(����, 1, 1, 0) * ��Ԥ��) As ʣ��Ԥ��, Sum(��Ԥ��) As ʣ���˿�,
         Max(�Ƿ���ҽ��) As �Ƿ���ҽ��
  Into n_�˷�����, n_ʣ��Ԥ��, n_����ֵ, n_�Ƿ����ҽ��
  From (Select Mod(a.��¼����, 10) As ����, Decode(Mod(a.��¼����, 10), 1, '��Ԥ��', a.���㷽ʽ) As ���㷽ʽ, Max(1) As �˷�����, Sum(��Ԥ��) As ��Ԥ��,
                Max(Decode(m.����, 3, 1, 4, 1, 0)) As �Ƿ���ҽ��
         From ����Ԥ����¼ A, ���㷽ʽ M
         Where a.���㷽ʽ = m.����(+) And a.��¼״̬ <> 0 And ����id In (Select Column_Value From Table(f_Str2List(v_����ids)))
         Group By Mod(a.��¼����, 10), Decode(Mod(a.��¼����, 10), 1, '��Ԥ��', a.���㷽ʽ)
         Having Sum(��Ԥ��) <> 0);

  If Round(Nvl(n_ʵ�ս��, 0), 5) <> Round(Nvl(n_����ֵ, 0), 5) Then
    v_Err_Msg := '���ν����ʣ�����δ�˿��������Ϣ��δ�˿���Ϣ����,����תΪסԺ����.' || Chr(13) || '����ʣ��ϼ�:' ||
                 LTrim(To_Char(n_ʵ�ս��, '9999999990.99')) || Chr(13) || '����ʣ��ϼ�:' ||
                 LTrim(To_Char(n_����ֵ, '9999999990.99'));
    Raise Err_Item;
  End If;

  --�Ƚ������Ѵ���
  n_����ֵ := Round(n_ʵ�ս��, 2) - Nvl(n_ʵ�ս��, 0);
  If Nvl(n_����ֵ, 0) <> 0 Then
    n_���� := Nvl(n_����, 0) + n_����ֵ;
  End If;
  n_ʵ�ս�� := Round(n_ʵ�ս��, 2);
  --1.1���Ϸ��ü�¼
  n_����id := ����id_In;
  If n_����id Is Null Then
    Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
  End If;

  ���ü�¼_Strict(n_����id, v_Nos, n_��id);

  --����ҽ��(��ǰ��ԭ����ID������Ӧ���������� ���ֵ�������ID
  ҽ��������ϸ_Strict(n_����id, v_����ids, v_Nos, n_��id, n_Ԥ�����, 1);

  --1.2����Ԥ����¼
  --���ϳ�Ԥ������
  If (n_�����˷� = 0 Or n_�˷����� = 0) And Nvl(�����˷�_In, 0) = 0 Then
    --1.Ԥ��ԭ����
    ����Ԥ����_Del(n_����id, v_����ids, Null, n_����ֵ);
    n_ʵ�ս�� := n_ʵ�ս�� - Nvl(n_����ֵ, 0);
  Elsif n_�˷����� >= 1 And n_ʣ��Ԥ�� >= n_ʵ�ս�� And n_ʣ��Ԥ�� <> 0 And Nvl(n_�Ƿ����ҽ��, 0) = 0 Then
    --2.��������Ԥ���㹻:�ڲ�����ʱ��δ��Ԥ���𣬶�����Ԥ���������ʣ��Ľ�ֱ�ӷ���Ԥ����
    --Ԥ����ָ�������
    ����Ԥ����_Del(n_����id, v_����ids, n_ʵ�ս��, n_����ֵ);
    n_ʵ�ս�� := 0;
  Elsif n_�˷����� = 1 And n_ʣ��Ԥ�� < n_ʵ�ս�� And n_ʣ��Ԥ�� <> 0 And Nvl(n_�Ƿ����ҽ��, 0) = 0 Then
    --3.ֻ��һ�ֽ����Ҳ�����:ʣ��ֻ��һ�ֽ��㷽ʽ�Ҵ���Ԥ��С����ʣ���,��ȫ��Ԥ����
    ����Ԥ����_Del(n_����id, v_����ids, Null, n_Ԥ�����);
    n_ʵ�ս�� := n_ʵ�ս�� - Nvl(n_Ԥ�����, 0);
  Else
    --4.���ڶ�����¼ʱ����Ҫ����Ԥ�������ų���Ȼ����ͳ�ƣ����ⲿ��תΪȱʡ���㷽ʽ
    Select ����Ԥ����¼_Id.Nextval Into n_Tempid From Dual;
    Insert Into ����Ԥ����¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��, ����id,
       �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, ��������, ��������id)
      Select n_Tempid, Max(NO), Max(ʵ��Ʊ��), 3, 3, ����id, Max(��ҳid) As ��ҳid, Max(����id) As ����id, Null, v_ȱʡ���㷽ʽ, Max(�������),
             'Ԥ����ʱ��¼', Null, Null, Null, Max(�տ�ʱ��), ����Ա����_In, ����Ա���_In, Sum(��Ԥ��), n_ԭ����id, Null, Null, Null, Null, Null,
             Null, -1 * n_ԭ����id, 3, Null
      From ����Ԥ����¼ A
      Where ��¼���� In (1, 11) And a.����id In (Select Column_Value From Table(f_Str2List(v_����ids))) And Nvl(��Ԥ��, 0) <> 0
      Group By ����id;
  End If;

  --��������ɷѼ�ҽ������(����һ��ͨ):У�Ա�־��ͳ�Ƹ���Ϊ1
  If n_ʵ�ս�� <> 0 Then
  
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����,
       �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, ��������, У�Ա�־, ��������id)
      Select ����Ԥ����¼_Id.Nextval, ��¼����, NO, 2, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �˷�ʱ��_In, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���_In, ����Ա����_In,
             0, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, -1 * n_����id, ��������, 1, ��������id
      From ����Ԥ����¼ A
      Where a.��¼���� = 3 And a.��¼״̬ = 1 And a.����id In (Select Column_Value From Table(f_Str2List(v_����ids))) And
            (a.�����id Is Null And a.���㿨��� Is Null);
    Update ����Ԥ����¼
    Set ��¼״̬ = 3
    Where ��¼���� = 3 And ��¼״̬ = 1 And ����id In (Select Column_Value From Table(f_Str2List(v_����ids)));
  
  End If;

  --2.Ʊ���ջ�
  �������Ʊ��_����(v_Nos);

  --3.�ɿ����ݴ���(
  --   �����������:
  --    1. ת������ֱ�����ʵ�,��ɿ����ݲ�����;
  --    2. ��ת��,�ٵ������˿���Ʊ,����Ҫ���нɿ����ݴ���
  If Nvl(�����˷�_In, 0) = 1 Then
    --�������
    ���˽���_Strict(n_����id, v_����ids, Null, 0);
  Elsif n_ʵ�ս�� <> 0 Then
    If n_�����˷� = 0 Then
      n_ʵ�ս�� := Null; --���ǲ����ˣ���ȫ��
    End If;
    --4.�˿�תԤ��(������Ʊ��,�ɲ���Աͨ���ش����)
    ���˽���_Strict(n_����id, v_����ids, n_ʵ�ս��);
  End If;

  If n_���� Is Not Null Then
  
    Update ����Ԥ����¼
    Set ��Ԥ�� = ��Ԥ�� - n_����
    Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_ȱʡ���㷽ʽ;
  
    Update ����Ԥ����¼
    Set ��Ԥ�� = ��Ԥ�� + n_����
    Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_����;
  
    If Sql%RowCount = 0 Then
      Select ����Ԥ����¼_Id.Nextval Into n_Tempid From Dual;
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ, �ɿ���id,
         �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������, ��������id)
      Values
        (n_Tempid, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, n_����, v_����, Null, �˷�ʱ��_In, Null, Null, Null,
         ����Ա���_In, ����Ա����_In, '', n_��id, Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 0, 3, n_Tempid);
    End If;
  End If;

  ����תס����_Over(n_����id, n_ԭ����id);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����תסԺ_�շ�ת��;
/
Create Or Replace Procedure Zl_����תסԺ_������ת��
(
  No_In           ���ò����¼.No%Type,
  ���ó���id_In   ����Ԥ����¼.����id%Type,
  �������id_In   ����Ԥ����¼.����id%Type,
  �������_In     ����Ԥ����¼.�������%Type,
  �˷�ʱ��_In     סԺ���ü�¼.����ʱ��%Type,
  ����Ա���_In   סԺ���ü�¼.����Ա���%Type,
  ����Ա����_In   סԺ���ü�¼.����Ա����%Type,
  ��ҳid_In       ����Ԥ����¼.��ҳid%Type,
  ��Ժ����id_In   ����Ԥ����¼.����id%Type,
  ���㷽ʽ_In     ����Ԥ����¼.���㷽ʽ%Type := Null,
  ����_In       ����Ԥ����¼.��Ԥ��%Type := Null,
  Ԥ������Ʊ��_In ����Ԥ����¼.Ԥ������Ʊ��%Type := 0
) As
  --���ܣ��Է��ò�������������ý���תסԺ���ô��� 
  --��Σ� 
  --  ���㷽ʽ_In ��Ϊ�գ���ʾ���г�Ԥ����ķ�ҽ�����ȫ����Ϊָ���Ľ��㷽ʽ�� 
  --              Ϊ�գ���ʾ���г�Ԥ����ķ�ҽ�����ȫ��תΪסԺԤ����  
  --  Ԥ������Ʊ��_In:Ԥ�����Ƿ����õ���Ʊ��
  Err_Item Exception;
  v_Err_Msg Varchar2(200);
  n_����ֵ  ����Ԥ����¼.��Ԥ��%Type;

  n_��id     ����ɿ����.Id%Type;
  v_����   ���㷽ʽ.����%Type;
  n_����   ����Ԥ����¼.��Ԥ��%Type;
  n_Dec      Number; --���С��λ�� 
  n_�첽���� Number;

  v_Nos    Varchar2(4000);
  n_����id ����Ԥ����¼.����id%Type;

  n_���˽�� ����Ԥ����¼.��Ԥ��%Type;
  n_δ�˽�� ����Ԥ����¼.��Ԥ��%Type;
  n_��Ԥ��   ����Ԥ����¼.��Ԥ��%Type;
  v_���㷽ʽ Varchar2(4000);
  v_Ԥ��no   ����Ԥ����¼.No%Type;

  n_�Ƿ����Ʊ�� ����Ԥ����¼.�Ƿ����Ʊ��%Type;

  --����Ԥ����� 
  Procedure ����Ԥ����¼_Insert
  (
    ����id_In     ����Ԥ����¼.����id%Type,
    ���_In       ����Ԥ����¼.���%Type,
    ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type,
    �տ�ʱ��_In   ����Ԥ����¼.�տ�ʱ��%Type,
    �������_In   ����Ԥ����¼.�������%Type,
    �����id_In   ����Ԥ����¼.�����id%Type := Null,
    ����_In       ����Ԥ����¼.����%Type := Null,
    ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
    ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
    ��������id_In ����Ԥ����¼.��������id%Type := Null
  ) As
    n_��ֵid ����Ԥ����¼.Id%Type;
    v_Ԥ��no ����Ԥ����¼.No%Type;
    n_����ֵ ����Ԥ����¼.���%Type;
  Begin
    If Nvl(���_In, 0) = 0 Or ���㷽ʽ_In Is Null Then
      Return;
    End If;
  
    --һ��ͨ��ÿһ�ʶ�����һ��Ԥ�����¼ 
    --������ͬһ�ֽ��㷽ʽֻ����һ��Ԥ�����¼ 
    Update ����Ԥ����¼
    Set ��� = Nvl(���, 0) + ���_In
    Where ��¼���� = 1 And ��¼״̬ = 1 And �տ�ʱ�� = �տ�ʱ��_In And ����id + 0 = ����id_In And ���㷽ʽ = ���㷽ʽ_In And
          (Nvl(�����id, 0) = 0 And Nvl(�����id_In, 0) = 0)
    Returning ID Into n_��ֵid;
    If Sql%RowCount = 0 Then
      v_Ԥ��no := Nextno(11);
      Select ����Ԥ����¼_Id.Nextval Into n_��ֵid From Dual;
    
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ, �ɿ���id, Ԥ�����,
         �����id, ����, ����˵��, ������ˮ��, �������, ��������id, ������Ա, ����ʱ��, Ԥ������Ʊ��)
      Values
        (n_��ֵid, v_Ԥ��no, Null, 1, 1, ����id_In, ��ҳid_In, ��Ժ����id_In, ���_In, ���㷽ʽ_In, �տ�ʱ��_In, Null, Null, Null, ����Ա���_In,
         ����Ա����_In, '����תסԺԤ��', n_��id, 2, �����id_In, ����_In, ����˵��_In, ������ˮ��_In, �������_In, Nvl(��������id_In, n_��ֵid), ����Ա����_In,
         �տ�ʱ��_In, Ԥ������Ʊ��_In);
    End If;
  
    Update Ԥ���������
    Set Ԥ����� = Nvl(Ԥ�����, 0) + ���_In
    Where ����id = ����id_In And Ԥ��id = n_��ֵid
    Returning Ԥ����� Into n_����ֵ;
    If Sql%RowCount = 0 Then
      Insert Into Ԥ��������� (Ԥ��id, ����id, Ԥ�����, Ԥ�����) Values (n_��ֵid, ����id_In, 2, ���_In);
      n_����ֵ := ���_In;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From Ԥ��������� Where Ԥ��id = n_��ֵid And Nvl(Ԥ�����, 0) = 0;
    End If;
  
    Update �������
    Set Ԥ����� = Nvl(Ԥ�����, 0) + ���_In
    Where ���� = 1 And ����id = ����id_In And ���� = 2
    Returning Ԥ����� Into n_����ֵ;
    If Sql%RowCount = 0 Then
      Insert Into ������� (����id, ����, ����, Ԥ�����, �������) Values (����id_In, 1, 2, ���_In, 0);
      n_����ֵ := ���_In;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ������� Where ����id = ����id_In And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
    End If;
  End;
Begin
  n_��id := Zl_Get��id(����Ա����_In);
  --���� 
  Begin
    Select ���� Into v_���� From ���㷽ʽ Where ���� = 9 And Rownum < 2;
  Exception
    When Others Then
      v_Err_Msg := 'û�з��������㷽ʽ�������Ƿ���ȷ���ã�';
      Raise Err_Item;
  End;
  n_���� := Nvl(����_In, 0);

  --���С��λ�� 
  Select zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  Select f_List2Str(Cast(Collect(a.No) As t_StrList), ',', 1), Max(a.����id)
  Into v_Nos, n_����id
  From ������ü�¼ A, ���ò����¼ B
  Where a.����id = b.�շѽ���id And b.��¼���� = 1 And b.���ӱ�־ = 0 And b.No = No_In;
  If v_Nos Is Null Then
    v_Err_Msg := 'δ�ҵ�ԭҽ�����������ݣ�����ת��ʧ��!';
    Raise Err_Item;
  End If;

  --1.���·�����˼�¼ 
  Update ������˼�¼
  Set ��¼״̬ = 2
  Where ���� = 1 And ����id In (Select /*+cardinality(b,10)*/
                             a.Id
                            From ������ü�¼ A, (Select Column_Value As NO From Table(f_Str2List(v_Nos))) B
                            Where a.No = b.No And Mod(a.��¼����, 10) = 1 And a.��¼״̬ In (1, 3));

  --2.����������ü�¼ 
  Update ������ü�¼
  Set ��¼״̬ = 3
  Where Mod(��¼����, 10) = 1 And ��¼״̬ = 1 And NO In (Select Column_Value As NO From Table(f_Str2List(v_Nos)));

  --����ԭ�շѼ�¼�Ƿ����ɶ���������ȷ�����ϼ�¼�Ƿ�Ҳ���ɶ������� 
  Select /*+cardinality(c,10)*/
   Count(1)
  Into n_�첽����
  From ���ý������ A, ������ü�¼ B, (Select Column_Value As NO From Table(f_Str2List(v_Nos))) C
  Where a.����id = b.Id And a.�����־ = 1 And b.��¼���� = 1 And b.No = c.No And Rownum < 2;

  For c_���� In (Select /*+cardinality(b,10)*/
                a.No, a.ʵ��Ʊ��, a.���, a.��������, a.�۸񸸺�, a.����id, a.ҽ�����, a.�����־, a.����, a.�Ա�, a.����, a.��ʶ��, a.���ʽ, a.���˿���id,
                a.�ѱ�, a.�շ����, a.�շ�ϸĿid, a.���㵥λ, a.��ҩ����, Sum(Nvl(a.����, 1) * a.����) As ����, a.�Ӱ��־, a.���ӱ�־, a.Ӥ����, a.������Ŀid,
                a.�վݷ�Ŀ, a.��׼����, Sum(a.Ӧ�ս��) As Ӧ�ս��, Sum(a.ʵ�ս��) As ʵ�ս��, a.������, a.��������id, a.������, a.����ʱ��, a.ִ�в���id, a.ִ����,
                Nvl(Min(Decode(a.��¼״̬, 2, a.ִ��״̬, 0)), 0) - 1 As ִ��״̬, a.����, Sum(a.���ʽ��) As ���ʽ��, Max(���մ���id) As ���մ���id,
                Max(������Ŀ��) As ������Ŀ��, Max(���ձ���) As ���ձ���, Max(��������) As ��������, Sum(a.ͳ����) As ͳ����, �Ƿ���, a.�Һ�id, a.��ҳid,
                a.���˲���id
               From ������ü�¼ A, (Select Column_Value As NO From Table(f_Str2List(v_Nos))) B
               Where a.No = b.No And a.��¼���� In (1, 11)
               Group By a.No, a.ʵ��Ʊ��, a.���, a.��������, a.�۸񸸺�, a.����id, a.ҽ�����, a.�����־, a.����, a.�Ա�, a.����, a.��ʶ��, a.���ʽ,
                        a.���˿���id, a.�ѱ�, a.�շ����, a.�շ�ϸĿid, a.���㵥λ, a.��ҩ����, a.�Ӱ��־, a.���ӱ�־, a.Ӥ����, a.������Ŀid, a.�վݷ�Ŀ, a.��׼����,
                        a.������, a.��������id, a.������, a.����ʱ��, a.ִ�в���id, a.ִ����, a.����, �Ƿ���, a.�Һ�id, a.��ҳid, a.���˲���id
               Having Nvl(Sum(Nvl(a.����, 1) * a.����), 0) <> 0) Loop
  
    Insert Into ������ü�¼
      (ID, ��¼����, NO, ��¼״̬, ʵ��Ʊ��, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid,
       ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����,
       ִ��״̬, ִ��ʱ��, ����, ����Ա���, ����Ա����, ����id, ���ʽ��, ���մ���id, ������Ŀ��, ���ձ���, ��������, ͳ����, ժҪ, �Ƿ���, �ɿ���id, ����״̬, �Һ�id, ��ҳid,
       ���˲���id)
    Values
      (���˷��ü�¼_Id.Nextval, 1, c_����.No, 2, c_����.ʵ��Ʊ��, c_����.���, c_����.��������, c_����.�۸񸸺�, c_����.����id, c_����.ҽ�����, c_����.�����־,
       c_����.����, c_����.�Ա�, c_����.����, c_����.��ʶ��, c_����.���ʽ, c_����.���˿���id, c_����.�ѱ�, c_����.�շ����, c_����.�շ�ϸĿid, c_����.���㵥λ, 1,
       c_����.��ҩ����, -1 * c_����.����, c_����.�Ӱ��־, c_����.���ӱ�־, c_����.Ӥ����, c_����.������Ŀid, c_����.�վݷ�Ŀ, c_����.��׼����, -1 * c_����.Ӧ�ս��,
       -1 * c_����.ʵ�ս��, c_����.������, c_����.��������id, c_����.������, c_����.����ʱ��, �˷�ʱ��_In, c_����.ִ�в���id, c_����.ִ����, c_����.ִ��״̬, Null,
       c_����.����, ����Ա���_In, ����Ա����_In, ���ó���id_In, -1 * c_����.���ʽ��, c_����.���մ���id, c_����.������Ŀ��, c_����.���ձ���, c_����.��������,
       -1 * c_����.ͳ����, '', c_����.�Ƿ���, n_��id, 0, c_����.�Һ�id, c_����.��ҳid, c_����.���˲���id);
  
    If n_�첽���� = 1 Then
      Insert Into ���ý������
        (�����־, ����id, �Ƿ�����, ����id, ���ʽ��, ����Ա���, ����Ա����)
      Values
        (1, ���˷��ü�¼_Id.Currval, 0, ���ó���id_In, -1 * c_����.���ʽ��, ����Ա���_In, ����Ա����_In);
    End If;
  End Loop;
  Zl_�����˷ѽ���_Modify(1, n_����id, ���ó���id_In, Null);

  --3.���ϲ�������¼��ͬʱ�ѽ�����Ʊ�ݻ��պ�ҽ��ԭ���ˣ� 
  Zl_���ò����¼_Delete(No_In, �������id_In, Null, �������_In, ���ó���id_In, ����Ա���_In, ����Ա����_In, �˷�ʱ��_In);
  Update ���ò����¼ Set ����״̬ = 0 Where ������� = �������_In;
  --����Ϊҽ���ӿ��ѵ��óɹ� 
  Update ����Ԥ����¼
  Set У�Ա�־ = 2
  Where ��¼���� = 6 And ����id = �������id_In And ���㷽ʽ In (Select ���� From ���㷽ʽ Where ���� In (3, 4));

  --4.�������ݴ��� 
  Select -1 * Nvl(Sum(a.��Ԥ��), 0)
  Into n_δ�˽��
  From ����Ԥ����¼ A
  Where a.������� = �������_In And a.���㷽ʽ Is Null;
  If Nvl(n_����, 0) = 0 Then
    n_���� := Round(n_δ�˽��, n_Dec) - n_δ�˽��;
  End If;
  n_δ�˽�� := n_δ�˽�� - n_����;

  For r_Ԥ�� In (Select Case
                        When Mod(a.��¼����, 10) = 1 Then
                         1
                        When Nvl(a.�����id, 0) <> 0 Then
                         2
                        Else
                         0
                      End As ����, a.����id, Nvl(a.��Ԥ��, 0) As ��Ԥ��, a.No, a.����id, a.���㷽ʽ, a.�������, a.�����id, a.����, a.������ˮ��,
                      a.����˵��, a.��������id
               From ����Ԥ����¼ A, ���㷽ʽ B
               Where a.���㷽ʽ = b.���� And a.��¼״̬ In (1, 3) And b.���� Not In (3, 4, 9) And
                     a.����id In (Select �շѽ���id From ���ò����¼ Where ��¼���� = 1 And ���ӱ�־ = 0 And NO = No_In)) Loop
  
    --���ǵ��ֽ��㷽ʽ 
    If r_Ԥ��.���� = 1 Then
      --Ԥ���� 
      Zl_���ò������_����˷�(�������id_In, Null, Null, Null, Null, Null, n_����, 0, -1 * n_δ�˽��);
      Exit;
    Elsif r_Ԥ��.���� = 2 Then
      --һ��ͨ 
      Select Nvl(Sum(���), 0) Into n_���˽�� From �����˿���Ϣ Where ��¼id = r_Ԥ��.����id;
      If r_Ԥ��.��Ԥ�� - n_���˽�� > 0 Then
        If r_Ԥ��.��Ԥ�� - n_���˽�� > n_δ�˽�� Then
          n_��Ԥ�� := n_δ�˽��;
        Else
          n_��Ԥ�� := r_Ԥ��.��Ԥ�� - n_���˽��;
        End If;
      
        v_���㷽ʽ := r_Ԥ��.���㷽ʽ || '|' || -1 * n_��Ԥ�� || '| | ';
        Zl_���ò������_����˷�(�������id_In, v_���㷽ʽ, r_Ԥ��.�����id, r_Ԥ��.����, r_Ԥ��.������ˮ��, r_Ԥ��.����˵��, n_����, 0, 0, 2);
        Zl_�����˿���Ϣ_Insert(�������_In, r_Ԥ��.����id, n_��Ԥ��, r_Ԥ��.����, r_Ԥ��.������ˮ��, r_Ԥ��.����˵��, 0, 0, 0, r_Ԥ��.�����id, r_Ԥ��.������ˮ��,
                         r_Ԥ��.����˵��);
      
        --תΪסԺԤ���� 
        ����Ԥ����¼_Insert(r_Ԥ��.����id, n_��Ԥ��, r_Ԥ��.���㷽ʽ, �˷�ʱ��_In, r_Ԥ��.�������, r_Ԥ��.�����id, r_Ԥ��.����, r_Ԥ��.������ˮ��, r_Ԥ��.����˵��,
                      r_Ԥ��.��������id);
      
        n_δ�˽�� := n_δ�˽�� - n_��Ԥ��;
        n_����   := 0;
      End If;
      If n_δ�˽�� = 0 Then
        Exit;
      End If;
    Else
      --������ҽ�����㷽ʽ 
      --���㷽ʽ|������|�������|����ժҪ 
      v_���㷽ʽ := r_Ԥ��.���㷽ʽ || '|' || -1 * n_δ�˽�� || '| | ';
      Zl_���ò������_����˷�(�������id_In, v_���㷽ʽ, Null, Null, Null, Null, n_����, 0);
    
      --תΪסԺԤ���� 
      ����Ԥ����¼_Insert(r_Ԥ��.����id, n_δ�˽��, r_Ԥ��.���㷽ʽ, �˷�ʱ��_In, r_Ԥ��.�������);
      Exit;
    End If;
  End Loop;

  --5.ת����ɴ��� 
  Delete From ����Ԥ����¼ Where ����id = �������id_In And ���㷽ʽ Is Null And Nvl(��Ԥ��, 0) = 0;
  If Sql%NotFound Then
    v_Err_Msg := '������δ�ɿ�����ݣ�������ɽ��㣡';
    Raise Err_Item;
  End If;
  Delete From ����Ԥ����¼ Where ����id = ���ó���id_In And ���㷽ʽ Is Null And Nvl(��Ԥ��, 0) = 0;
  If Sql%NotFound Then
    v_Err_Msg := '������δ�ɿ�����ݣ�������ɽ��㣡';
    Raise Err_Item;
  End If;
  Update ����Ԥ����¼ Set У�Ա�־ = 0, �Ự�� = Null Where ������� = �������_In;

  --6.���� �Ƿ����Ʊ�� ��� 
  Select Max(a.�Ƿ����Ʊ��)
  Into n_�Ƿ����Ʊ��
  From ����Ԥ����¼ A,
       (Select ����id
         From (Select b.����id
                From ���ò����¼ B
                Where b.No = No_In And b.��¼���� = 1 And b.��¼״̬ In (1, 3)
                Order By b.�Ǽ�ʱ��)
         Where Rownum < 2) B
  Where a.����id = b.����id;

  Update ����Ԥ����¼ Set �Ƿ����Ʊ�� = n_�Ƿ����Ʊ�� Where ������� = �������_In And �������� = 6;

  --��Ա�ɿ�����Ҫ��ҽ���� 
  For c_Ԥ�� In (Select a.���㷽ʽ, a.����Ա����, Nvl(Sum(a.��Ԥ��), 0) As ��Ԥ��
               From ����Ԥ����¼ A, ���㷽ʽ B
               Where a.���㷽ʽ = b.���� And b.���� In (3, 4) And a.������� = �������_In
               Group By a.���㷽ʽ, a.����Ա����) Loop
  
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + c_Ԥ��.��Ԥ��
    Where �տ�Ա = c_Ԥ��.����Ա���� And ���� = 1 And ���㷽ʽ = c_Ԥ��.���㷽ʽ
    Returning ��� Into n_����ֵ;
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ����
        (�տ�Ա, ���㷽ʽ, ����, ���)
      Values
        (c_Ԥ��.����Ա����, c_Ԥ��.���㷽ʽ, 1, c_Ԥ��.��Ԥ��);
      n_����ֵ := c_Ԥ��.��Ԥ��;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ��Ա�ɿ����
      Where �տ�Ա = c_Ԥ��.����Ա���� And ���� = 1 And ���㷽ʽ = c_Ԥ��.���㷽ʽ And Nvl(���, 0) = 0;
    End If;
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����תסԺ_������ת��;
/
Create Or Replace Procedure Zl_����תסԺ_��������
(
  No_In           ���˽��ʼ�¼.No%Type,
  ����id_In       ���˽��ʼ�¼.Id%Type,
  ��ҳid_In       ����Ԥ����¼.��ҳid%Type,
  ��Ժ����id_In   ����Ԥ����¼.����id%Type,
  �������_In     Number := 0,
  Ԥ������Ʊ��_In ����Ԥ����¼.Ԥ������Ʊ��%Type := 0
) As
  --���ܣ��������תסԺ���Ͻ��ʽ������ݣ���������ģʽ���� 
  --��Σ� 
  --  �������_In:0-��ʼ��������;1-��ɽ������� 
  --  Ԥ������Ʊ��_In:Ԥ�����Ƿ����õ���Ʊ�ݣ��������ʱ����
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_����ֵ   ����Ԥ����¼.��Ԥ��%Type;
  v_���㷽ʽ ����Ԥ����¼.���㷽ʽ%Type;
  n_������ ����Ԥ����¼.��Ԥ��%Type;
  n_ԭ����id ���˽��ʼ�¼.Id%Type;
  n_���ʽ�� ���˽��ʼ�¼.���ʽ��%Type;

  n_��Ԥ��       ����Ԥ����¼.��Ԥ��%Type;
  n_�Ƿ�תΪԤ�� Number(2);
  n_Ԥ��id       ����Ԥ����¼.Id%Type;
  n_����       ����Ԥ����¼.��Ԥ��%Type;
  n_������֧Ʊ   Number(2);

  Cursor c_Balance_Data Is
    Select NO, ����id, ����id, ��ҳid, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, �ɿ���id
    From ����Ԥ����¼
    Where ����id = ����id_In And ���㷽ʽ Is Null;
  r_Balance_Data c_Balance_Data%RowType;

  Procedure ��Ա�ɿ����_Update
  (
    �տ�Ա_In   ��Ա�ɿ����.�տ�Ա%Type,
    ���㷽ʽ_In ��Ա�ɿ����.���㷽ʽ%Type,
    ���_In     ��Ա�ɿ����.���%Type
  ) As
    --���ܣ����� ��Ա�ɿ���� 
    n_����ֵ ��Ա�ɿ����.���%Type;
  Begin
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + Nvl(���_In, 0)
    Where �տ�Ա = �տ�Ա_In And ���� = 1 And ���㷽ʽ = ���㷽ʽ_In
    Returning ��� Into n_����ֵ;
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (�տ�Ա_In, ���㷽ʽ_In, 1, Nvl(���_In, 0));
      n_����ֵ := Nvl(���_In, 0);
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ��Ա�ɿ���� Where �տ�Ա = �տ�Ա_In And ���� = 1 And ���㷽ʽ = ���㷽ʽ_In And Nvl(���, 0) = 0;
    End If;
  End ��Ա�ɿ����_Update;

  Procedure סԺԤ����_Insert
  (
    ����id_In     ����Ԥ����¼.����id%Type,
    ���_In       ����Ԥ����¼.���%Type,
    ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type,
    �տ�ʱ��_In   ����Ԥ����¼.�տ�ʱ��%Type,
    �������_In   ����Ԥ����¼.�������%Type,
    ����Ա���_In ����Ԥ����¼.����Ա���%Type,
    ����Ա����_In ����Ԥ����¼.����Ա����%Type,
    �ɿ���id_In   ����Ԥ����¼.�ɿ���id%Type,
    �����id_In   ����Ԥ����¼.�����id%Type := Null,
    ����_In       ����Ԥ����¼.����%Type := Null,
    ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
    ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
    ��������id_In ����Ԥ����¼.��������id%Type := Null,
    ������Ա_In   ����Ԥ����¼.������Ա%Type := Null,
    ����ʱ��_In   ����Ԥ����¼.����ʱ��%Type := Null
  ) As
    --���ܣ����� סԺԤ���� 
    n_Ԥ��id ����Ԥ����¼.Id%Type;
    v_Ԥ��no ����Ԥ����¼.No%Type;
    n_����ֵ ����Ԥ����¼.���%Type;
  Begin
    If Nvl(���_In, 0) = 0 Or ���㷽ʽ_In Is Null Then
      Return;
    End If;
  
    --һ��ͨ��ÿһ�ʶ�����һ��Ԥ�����¼ 
    --������ͬһ�ֽ��㷽ʽֻ����һ��Ԥ�����¼ 
    Update ����Ԥ����¼
    Set ��� = Nvl(���, 0) + ���_In
    Where ��¼���� = 1 And ��¼״̬ = 1 And �տ�ʱ�� = �տ�ʱ��_In And ����id + 0 = ����id_In And ���㷽ʽ = ���㷽ʽ_In And
          (Nvl(�����id, 0) = 0 And Nvl(�����id_In, 0) = 0)
    Returning ID Into n_Ԥ��id;
    If Sql%RowCount = 0 Then
      v_Ԥ��no := Nextno(11);
      Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
    
      Insert Into ����Ԥ����¼
        (ID, NO, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ժҪ, �ɿ���id, Ԥ�����, �����id, ����, ����˵��, ������ˮ��,
         �������, ��������id, ������Ա, ����ʱ��, Ԥ������Ʊ��)
      Values
        (n_Ԥ��id, v_Ԥ��no, 1, 1, ����id_In, ��ҳid_In, ��Ժ����id_In, ���_In, ���㷽ʽ_In, �տ�ʱ��_In, ����Ա���_In, ����Ա����_In, '����תסԺԤ��',
         �ɿ���id_In, 2, �����id_In, ����_In, ����˵��_In, ������ˮ��_In, �������_In, Nvl(��������id_In, n_Ԥ��id), Nvl(������Ա_In, ����Ա����_In),
         Nvl(����ʱ��_In, �տ�ʱ��_In), Ԥ������Ʊ��_In);
    End If;
  
    Update Ԥ���������
    Set Ԥ����� = Nvl(Ԥ�����, 0) + ���_In
    Where ����id = ����id_In And Ԥ��id = n_Ԥ��id
    Returning Ԥ����� Into n_����ֵ;
    If Sql%RowCount = 0 Then
      Insert Into Ԥ��������� (Ԥ��id, ����id, Ԥ�����, Ԥ�����) Values (n_Ԥ��id, ����id_In, 2, ���_In);
      n_����ֵ := ���_In;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From Ԥ��������� Where Ԥ��id = n_Ԥ��id And Nvl(Ԥ�����, 0) = 0;
    End If;
  
    Update �������
    Set Ԥ����� = Nvl(Ԥ�����, 0) + ���_In
    Where ���� = 1 And ����id = ����id_In And ���� = 2
    Returning Ԥ����� Into n_����ֵ;
    If Sql%RowCount = 0 Then
      Insert Into ������� (����id, ����, ����, Ԥ�����, �������) Values (����id_In, 1, 2, ���_In, 0);
      n_����ֵ := ���_In;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ������� Where ����id = ����id_In And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
    End If;
  
    --������Ա�ɿ���� 
    ��Ա�ɿ����_Update(����Ա����_In, ���㷽ʽ_In, ���_In);
  End סԺԤ����_Insert;
Begin
  Open c_Balance_Data;
  Fetch c_Balance_Data
    Into r_Balance_Data;
  If r_Balance_Data.No Is Null Then
    v_Err_Msg := 'δ�ҵ�ָ���Ľ������Ͻ������ݣ�';
    Raise Err_Item;
  End If;

  Select Max(ID), Nvl(Sum(���ʽ��), 0)
  Into n_ԭ����id, n_���ʽ��
  From ���˽��ʼ�¼
  Where ��¼״̬ In (1, 3) And NO = No_In;
  If Nvl(n_ԭ����id, 0) = 0 Then
    v_Err_Msg := 'û�з���Ҫ���ϵĽ��ʵ��ݣ������Ѿ����ϣ�';
    Raise Err_Item;
  End If;

  If Nvl(�������_In, 0) = 0 Then
    --���ͣ�0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ,4-һ��ͨ(��),5-���ѿ� 
    --��ԭ�����ϴ���
    For r_Pay In (Select Case
                            When Mod(a.��¼����, 10) = 1 Then
                             1
                            When Instr(',3,4,', ',' || b.���� || ',') > 0 And a.�����id Is Null Then
                             2
                            When Nvl(a.�����id, 0) <> 0 Then
                             3
                            When j.���㷽ʽ Is Not Null Then
                             4
                            When a.���㿨��� Is Not Null Then
                             5
                            Else
                             0
                          End As ����, a.Id, a.No, a.����id, a.����id, a.��ҳid, a.���㷽ʽ,
                         Nvl(Decode(a.���㿨���, Null, a.��Ԥ��, -1 * p.Ӧ�ս��), 0) As ��Ԥ��, a.�������, a.�����id,
                         Decode(a.���㿨���, Null, a.����, p.����) As ����, a.������ˮ��, a.����˵��, a.������λ, a.���㿨���, p.���ѿ�id,
                         Nvl(b.����, 1) As ��������, a.��������id, a.������Ա, a.����ʱ��, b.Ӧ����
                  From ����Ԥ����¼ A, ���㷽ʽ B, һ��ͨĿ¼ J, ���˿������¼ P
                  Where a.���㷽ʽ = b.����(+) And a.���㷽ʽ = j.���㷽ʽ(+) And a.Id = p.����id(+) And a.���㿨��� = p.�ӿڱ��(+) And
                        a.����id = n_ԭ����id
                  Order By ��Ԥ��) Loop
      n_��Ԥ�� := Nvl(r_Pay.��Ԥ��, 0);
    
      --1-Ԥ����,�����ʱ�ٴ��� 
      --2-ҽ��,�����д��� 
      If r_Pay.���� = 1 Or r_Pay.���� = 2 Then
        n_�Ƿ�תΪԤ�� := 0;
      Else
        n_�Ƿ�תΪԤ�� := 1;
      End If;
    
      --3-һ��ͨ 
      If r_Pay.���� = 3 Then
        --��Ҫ����Ƿ���ֽ��㷽ʽ��ҽ��������ǣ�����Ҫ���ýӿ��˿� 
        Select Count(Distinct ���㷽ʽ) + Max(Decode(b.����, 3, 2, 4, 2, 0))
        Into n_�Ƿ�תΪԤ��
        From ����Ԥ����¼ A, ���㷽ʽ B
        Where a.���㷽ʽ = b.����(+) And ����id = n_ԭ����id And �����id = r_Pay.�����id And Nvl(��������id, 0) = Nvl(r_Pay.��������id, 0);
      
        If Nvl(n_�Ƿ�תΪԤ��, 0) = 1 And n_��Ԥ�� < 0 Then
          --���ֽ��㷽ʽ�Ҳ���ҽ����ͬʱ����ʱ���˿�/ת�ˣ��򲻴��� 
          n_�Ƿ�תΪԤ�� := 0;
        End If;
      End If;
    
      --4-һ��ͨ(��),ԭ���˻� 
    
      --5-���ѿ�,ԭ���˻� 
      If r_Pay.���� = 5 Then
        Update ����Ԥ����¼
        Set ��Ԥ�� = Nvl(��Ԥ��, 0) - n_��Ԥ��
        Where ����id = ����id_In And ���㷽ʽ = r_Pay.���㷽ʽ And ���㿨��� = r_Pay.���㿨���
        Returning ID Into n_Ԥ��id;
        If Sql%NotFound Then
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (ID, ��¼����, NO, ��¼״̬, ����id, ����id, ��ҳid, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ���㿨���, У�Ա�־, ��������)
          Values
            (n_Ԥ��id, 12, r_Pay.No, 1, r_Pay.����id, r_Pay.����id, r_Pay.��ҳid, r_Pay.���㷽ʽ, r_Balance_Data.�տ�ʱ��,
             r_Balance_Data.����Ա���, r_Balance_Data.����Ա����, -1 * n_��Ԥ��, ����id_In, r_Balance_Data.�ɿ���id, r_Pay.���㿨���, 2, 2);
        End If;
      
        --���뿨�����¼ 
        Zl_���˿������¼_�˿�(r_Pay.���㿨���, r_Pay.����, r_Pay.���ѿ�id, n_��Ԥ��, r_Pay.Id, n_Ԥ��id, r_Balance_Data.����Ա���,
                      r_Balance_Data.����Ա����, r_Balance_Data.�տ�ʱ��);
      
        Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� + n_��Ԥ�� Where ����id = ����id_In And ���㷽ʽ Is Null;
        n_�Ƿ�תΪԤ�� := 0;
      End If;
    
      --0-��ͨ����,תΪסԺԤ�� 
      If r_Pay.���� = 0 Then
        If Nvl(r_Pay.��������, 1) = 1 Or Nvl(r_Pay.��������, 1) = 9 Or Nvl(r_Pay.Ӧ����, 0) = 1 Then
          --�ֽ�����ѡ�Ӧ����(��֧Ʊ)�Ȳ��˿�,�����ʱ��תΪסԺԤ�� 
          n_�Ƿ�תΪԤ�� := 0;
        End If;
        If Nvl(r_Pay.��������, 1) = 2 And Instr(r_Pay.���㷽ʽ, '֧Ʊ') > 0 Then
          --������֧Ʊ��ֱ�Ӱ�����תΪסԺԤ�� 
          Select Count(1)
          Into n_������֧Ʊ
          From ����Ԥ����¼ A, ���㷽ʽ B
          Where a.���㷽ʽ = b.���� And ����id = n_ԭ����id And Nvl(b.Ӧ����, 0) = 1 And Rownum < 2;
          If Nvl(n_������֧Ʊ, 0) > 0 Then
            n_�Ƿ�תΪԤ�� := 0;
          End If;
        End If;
      End If;
    
      If n_�Ƿ�תΪԤ�� > 0 Then
        Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, ����Ա���, ����Ա����, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��,
           ����˵��, ������λ, ����id, У�Ա�־, ��������, ��������id, ������Ա, ����ʱ��)
        Values
          (n_Ԥ��id, 12, r_Pay.No, 1, r_Pay.����id, r_Pay.��ҳid, r_Pay.����id, -1 * n_��Ԥ��, r_Pay.���㷽ʽ, r_Pay.�������,
           r_Balance_Data.�տ�ʱ��, r_Balance_Data.����Ա���, r_Balance_Data.����Ա����, r_Balance_Data.�ɿ���id, r_Pay.�����id,
           r_Pay.���㿨���, r_Pay.����, r_Pay.������ˮ��, r_Pay.����˵��, r_Pay.������λ, ����id_In, Decode(n_�Ƿ�תΪԤ��, 1, 2, 1), 2,
           r_Pay.��������id, r_Pay.������Ա, r_Pay.����ʱ��);
      
        --תΪסԺԤ�� 
        If n_�Ƿ�תΪԤ�� = 1 Then
          סԺԤ����_Insert(r_Pay.����id, n_��Ԥ��, r_Pay.���㷽ʽ, r_Balance_Data.�տ�ʱ��, r_Pay.�������, r_Balance_Data.����Ա���,
                       r_Balance_Data.����Ա����, r_Balance_Data.�ɿ���id, r_Pay.�����id, r_Pay.����, r_Pay.������ˮ��, r_Pay.����˵��,
                       r_Pay.��������id, r_Pay.������Ա, r_Pay.����ʱ��);
        End If;
      
        Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� + n_��Ԥ�� Where ����id = ����id_In And ���㷽ʽ Is Null;
      End If;
    End Loop;
    Return;
  End If;

  -------------------------------------------------------------------------------------------------------------- 
  --����������תסԺ�������� 
  Select -1 * Nvl(��Ԥ��, 0) Into n_������ From ����Ԥ����¼ Where ����id = ����id_In And ���㷽ʽ Is Null;

  --1.��ʣ���������Ԥ���� 
  If Nvl(n_������, 0) > 0 Then
    For r_Ԥ�� In (Select NO, ʵ��Ʊ��, ��¼״̬, ����id, ��ҳid, ����id, ���㷽ʽ, Max(�������) As �������, Max(ժҪ) As ժҪ, Max(�ɿλ) As �ɿλ,
                        Max(��λ������) As ��λ������, Max(��λ�ʺ�) As ��λ�ʺ�, Sum(��Ԥ��) As ��Ԥ��, Max(Ԥ�����) As Ԥ�����, �����id, ���㿨���,
                        Max(����) As ����, Max(��������id) As ��������id, Max(������ˮ��) As ������ˮ��, Max(����˵��) As ����˵��, Max(������λ) As ������λ,
                        Nvl(Max(�Ƿ�ת�ʼ�����), 0) As �Ƿ�ת�ʼ�����, Max(Ԥ��id) As Ԥ��id, Max(����ʱ��) As ����ʱ��, Max(������Ա) As ������Ա
                 From (Select a.No, a.ʵ��Ʊ��, a.��¼״̬, ����id, ��ҳid, ����id, a.���㷽ʽ, a.�������, a.ժҪ, a.�ɿλ, a.��λ������, a.��λ�ʺ�, a.��Ԥ��,
                               a.Ԥ�����, a.�����id, a.���㿨���, a.��������id, a.����, a.������ˮ��, a.����˵��, a.������λ, b.�Ƿ�ת�ʼ�����,
                               Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, 0, a.Id), 0) As Ԥ��id, a.����ʱ�� As ����ʱ��, a.������Ա As ������Ա
                        From ����Ԥ����¼ A, ҽ�ƿ���� B
                        Where a.����id = n_ԭ����id And a.��¼���� In (1, 11) And Nvl(a.��Ԥ��, 0) <> 0 And a.�����id = b.Id(+)
                        Union All
                        Select a.No, a.ʵ��Ʊ��, a.��¼״̬, a.����id, ��ҳid, a.����id, a.���㷽ʽ, '' || ������� As �������, '' As ժҪ, '' As �ɿλ,
                               '' As ��λ������, '' As ��λ�ʺ�, -1 * b.��� As ��Ԥ��, a.Ԥ�����, a.�����id, a.���㿨���, a.��������id, '' As ����,
                               '' As ������ˮ��, '' As ����˵��, '' As ������λ, 0 As �Ƿ�ת�ʼ�����,
                               Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, 0, a.Id), 0) As Ԥ��id, a.����ʱ�� As ����ʱ��, a.������Ա As ������Ա
                        From ����Ԥ����¼ A, �����˿���Ϣ B
                        Where b.����id = n_ԭ����id And a.Id = b.��¼id And Nvl(b.�Ƿ�δ��, 0) <> 1)
                 Group By NO, ʵ��Ʊ��, ��¼״̬, ����id, ��ҳid, ����id, ���㷽ʽ, �����id, ���㿨���
                 Having Nvl(Sum(��Ԥ��), 0) <> 0
                 Order By Ԥ����� Desc, �Ƿ�ת�ʼ����� Desc, ����ʱ�� Desc) Loop
      n_��Ԥ�� := Nvl(r_Ԥ��.��Ԥ��, 0);
    
      If n_������ > n_��Ԥ�� Then
        n_������ := Round(n_������ - n_��Ԥ��, 6);
      Else
        n_��Ԥ��   := Nvl(n_������, 0);
        n_������ := 0;
      End If;
    
      Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, У�Ա�־, �����id, ���㿨���, ����,
         ������ˮ��, ����˵��, �������, ��������, ��������id, ����ʱ��, ������Ա)
      Values
        (n_Ԥ��id, 12, r_Balance_Data.No, 1, r_Balance_Data.����id, r_Balance_Data.����id, r_Balance_Data.��ҳid, Null,
         r_Ԥ��.���㷽ʽ, r_Balance_Data.�տ�ʱ��, r_Balance_Data.����Ա���, r_Balance_Data.����Ա����, -1 * n_��Ԥ��, ����id_In,
         r_Balance_Data.�ɿ���id, 2, r_Ԥ��.�����id, r_Ԥ��.���㿨���, r_Ԥ��.����, r_Ԥ��.������ˮ��, r_Ԥ��.����˵��, r_Ԥ��.�������, 2, r_Ԥ��.��������id,
         r_Ԥ��.����ʱ��, r_Ԥ��.������Ա);
    
      --תΪסԺԤ�� 
      סԺԤ����_Insert(r_Ԥ��.����id, n_��Ԥ��, r_Ԥ��.���㷽ʽ, r_Balance_Data.�տ�ʱ��, r_Ԥ��.�������, r_Balance_Data.����Ա���,
                   r_Balance_Data.����Ա����, r_Balance_Data.�ɿ���id, r_Ԥ��.�����id, r_Ԥ��.����, r_Ԥ��.������ˮ��, r_Ԥ��.����˵��, r_Ԥ��.��������id,
                   r_Ԥ��.������Ա, r_Ԥ��.����ʱ��);
    
      Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� + n_��Ԥ�� Where ����id = ����id_In And ���㷽ʽ Is Null;
    
      If n_������ = 0 Then
        Exit;
      End If;
    End Loop;
  End If;

  n_������ := -1 * n_������;
  --2.��δ�˽��ȫ����"�ֽ�"תΪסԺԤ�������˿� 
  n_��Ԥ�� := Zl_Cent_Money(n_������);
  n_���� := Round(n_������ - n_��Ԥ��, 6);
  If n_��Ԥ�� <> 0 Then
    Select Nvl(Max(����), '�ֽ�') Into v_���㷽ʽ From ���㷽ʽ Where Nvl(����, 0) = 1 And Rownum < 1;
  
    Update ����Ԥ����¼
    Set ��Ԥ�� = Nvl(��Ԥ��, 0) + n_��Ԥ��
    Where ����id = ����id_In And ���㷽ʽ = v_���㷽ʽ And Rownum < 2;
    If Sql%NotFound Then
      Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, У�Ա�־, ��������)
      Values
        (n_Ԥ��id, 12, r_Balance_Data.No, 1, r_Balance_Data.����id, r_Balance_Data.����id, r_Balance_Data.��ҳid, Null, v_���㷽ʽ,
         r_Balance_Data.�տ�ʱ��, r_Balance_Data.����Ա���, r_Balance_Data.����Ա����, n_��Ԥ��, ����id_In, r_Balance_Data.�ɿ���id, 2, 2);
    End If;
  
    --תΪסԺԤ�� 
    סԺԤ����_Insert(r_Balance_Data.����id, -1 * n_��Ԥ��, v_���㷽ʽ, r_Balance_Data.�տ�ʱ��, Null, r_Balance_Data.����Ա���,
                 r_Balance_Data.����Ա����, r_Balance_Data.�ɿ���id);
  
    Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - n_��Ԥ�� Where ����id = ����id_In And ���㷽ʽ Is Null;
  End If;

  --3.��ɽ������� 
  Zl_���˽�������_Modify(1, r_Balance_Data.����id, ����id_In, Null, Null, Null, Null, Null, Null, Null, n_����, Null,
                   r_Balance_Data.����Ա���, r_Balance_Data.����Ա����, r_Balance_Data.�տ�ʱ��, Null, 2);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����תסԺ_��������;
/
Create Or Replace Procedure Zl_�������תסԺ_Modify
(
  ��������_In   Number,
  ����id_In     ����Ԥ����¼.����id%Type,
  ����id_In     ���˽��ʼ�¼.����id%Type,
  ���㷽ʽ_In   Varchar2,
  ����Ա���_In ����Ԥ����¼.����Ա���%Type := Null,
  ����Ա����_In ����Ԥ����¼.����Ա����%Type := Null,
  ����˷�_In   Number := 0,
  ��������id_In ����Ԥ����¼.Id%Type := Null,
  �˿�ʱ��_In   ����Ԥ����¼.�տ�ʱ��%Type := Null,
  У�Ա�־_In   ����Ԥ����¼.У�Ա�־%Type := Null,
  �����_In   ����Ԥ����¼.��Ԥ��%Type := Null,
  �����id_In   ����Ԥ����¼.�����id%Type := Null,
  ����_In       ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
  ���ԭ����_In Number := 0
) As
  ------------------------------------------------------------------------------------------------------------------------------
  --����:�շѽ���ʱ,�޸Ľ���������Ϣ
  --��������_In:
  --   0-������У�Ա�־:ֻ���¹�������ID��У�Ա�־
  --   1-��ͨ�˷ѷ�ʽ:
  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
  --   2.�������˷ѽ���:
  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
  --     �ڿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
  --     ��������ID_IN:
  --     ���ԭ����_In:1-��ʾ�ڸ�������ǰ�����ԭ���Ľ�����Ϣ(������ID+��������ID�����);0-��ʾ�����
  --   3-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
  --   4-���ѿ�����:
  --     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."
  -- ����˷�_In:0-δ����˷�;1-�쳣����˷�;2-����˷�
  -- ��Ԥ������ids_In:����ö��ַ���,��Ԥ��ʱ��Ч(��Ԥ�������ҵ���������һ��),��Ҫ��ʹ�ü�����Ԥ����
  -- У�Ա�־_In:0-��ɻ���ҪУ��;1-��ҪУ��;2-�ӿ��Ѿ����óɹ�
  ------------------------------------------------------------------------------------------------------------------------------
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_��������  Varchar2(500);
  v_��ǰ����  Varchar2(500);
  v_ԭ����ids Varchar2(500);
  v_���㷽ʽ  ����Ԥ����¼.���㷽ʽ%Type;
  n_������  ����Ԥ����¼.��Ԥ��%Type;
  n_����ֵ    ��Ա�ɿ����.���%Type;
  v_�������  ����Ԥ����¼.�������%Type;
  v_����ժҪ  ����Ԥ����¼.ժҪ%Type;
  v_����    ���㷽ʽ.����%Type;
  n_Ԥ��id    ����Ԥ����¼.Id%Type;
  n_�ɿ���id  ����Ԥ����¼.�ɿ���id%Type;
  n_У�Ա�־  ����Ԥ����¼.У�Ա�־%Type;
  n_�������  ����Ԥ����¼.��Ԥ��%Type;
  d_����ʱ��  ����Ԥ����¼.����ʱ��%Type;
  v_������Ա  ����Ԥ����¼.������Ա%Type;
  n_Dec       Number; --���С��λ��

  n_Count  Number;
  l_Ԥ��id t_NumList := t_NumList();
  n_���� ����Ԥ����¼.��Ԥ��%Type;

  n_�Ƿ����Ʊ�� ����Ԥ����¼.�Ƿ����Ʊ��%Type;

  Procedure Zl_Square_Update
  (
    ����ids_In    Varchar2,
    Ԥ��id_In     ����Ԥ����¼.Id%Type,
    �ֽ���id_In   ����Ԥ����¼.����id%Type,
    �ɿ���id_In   ����Ԥ����¼.�ɿ���id%Type,
    �˿�ʱ��_In   ����Ԥ����¼.�տ�ʱ��%Type,
    �������_In   ����Ԥ����¼.�������%Type,
    �˷ѽ��_In   ����Ԥ����¼.��Ԥ��%Type := Null,
    ���㿨���_In ����Ԥ����¼.���㿨���%Type := Null
  ) As
    n_Ԥ��id   ����Ԥ����¼.Id%Type;
    n_������ ����Ԥ����¼.��Ԥ��%Type;
    n_���ν�� ����Ԥ����¼.��Ԥ��%Type;
  Begin
  
    n_������ := Nvl(�˷ѽ��_In, 0);
  
    n_Ԥ��id := Ԥ��id_In;
    --�������ѿ�,���㿨��������Ѿ�������
    For v_У�� In (Select Min(a.Id) As Ԥ��id, c.���ѿ�id, -1 * Nvl(Sum(c.Ӧ�ս��), 0) As ������, c.�ӿڱ��, c.����
                 From ����Ԥ����¼ A, ���˿������¼ C
                 Where a.Id = c.����id And a.���㿨��� = ���㿨���_In And a.��¼���� = 3 And
                       a.����id In (Select Column_Value From Table(f_Str2List(����ids_In)))
                 Group By c.���ѿ�id, c.�ӿڱ��, c.����) Loop
    
      If v_У��.������ < n_������ Then
        n_���ν�� := v_У��.������;
        n_������ := n_������ - v_У��.������;
      Else
        n_���ν�� := n_������;
        n_������ := 0;
      End If;
    
      --����ʱ,ֻ����һ��
      If n_Ԥ��id = 0 Then
        Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id,
           Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������)
          Select n_Ԥ��id, ��¼����, NO, 2, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �˿�ʱ��_In, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���_In, ����Ա����_In,
                 -1 * �˷ѽ��_In, �ֽ���id_In, �ɿ���id_In, Ԥ�����, �����id, Nvl(���㿨���, v_У��.�ӿڱ��), ����, ������ˮ��, ����˵��, ������λ, 2, �������_In,
                 ��������
          From ����Ԥ����¼ A
          Where ID = v_У��.Ԥ��id;
      End If;
    
      Zl_���˿������¼_�˿�(v_У��.�ӿڱ��, v_У��.����, v_У��.���ѿ�id, n_���ν��, v_У��.Ԥ��id, n_Ԥ��id, ����Ա���_In, ����Ա����_In, �˿�ʱ��_In);
    
      If n_������ = 0 Then
        Exit;
      End If;
    End Loop;
  End;
Begin

  If ����Ա����_In Is Null Then
    n_�ɿ���id := Null;
  Else
    n_�ɿ���id := Zl_Get��id(����Ա����_In);
  End If;

  Select Count(1) Into n_Count From ������ü�¼ Where ����id = ����id_In And Rownum < 2;
  If n_Count = 0 Then
    v_Err_Msg := 'δ�ҵ�ָ���������շѵ��˷Ѽ�¼,���飡';
    Raise Err_Item;
  End If;

  Select Count(1) Into n_Count From ����Ԥ����¼ Where ����id = ����id_In And ���㷽ʽ Is Null;

  If n_Count = 0 Then
    --������㷽ʽΪNULL�ļ�¼���Ա��˿�
    Select Sum(Nvl(���ʽ��, 0)) - Sum(��Ԥ��)
    Into n_������
    From (Select Sum(���ʽ��) As ���ʽ��, 0 As ��Ԥ��
           From סԺ���ü�¼
           Where ����id = ����id_In
           Union All
           Select Sum(���ʽ��) As ���ʽ��, 0 As ��Ԥ��
           From ������ü�¼
           Where ����id = ����id_In
           Union All
           Select 0 As ���ʽ��, Sum(��Ԥ��) As ��Ԥ��
           From ����Ԥ����¼
           Where ����id = ����id_In);
  
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ����id, ��ҳid, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, У�Ա�־, ��������)
      Select ����Ԥ����¼_Id.Nextval, 3, Null, 2, ����id_In, Null, Null, Null, �˿�ʱ��_In, ����Ա���_In, ����Ա����_In, n_������, ����id_In,
             n_�ɿ���id, 0, 3
      From Dual;
  End If;

  If ��������_In = 0 Then
    --������У�Ա�־:ֻ���¹�������ID��У�Ա�־
    Update ����Ԥ����¼
    Set У�Ա�־ = У�Ա�־_In
    Where ����id = ����id_In And Nvl(��������id, 0) = Nvl(��������id_In, 0);
    Return;
  End If;

  --���С��λ��
  Select zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  --1.���ӽ��㷽ʽΪ�յĽ�������
  n_���� := �����_In;
  --��������
  If Nvl(n_����, 0) <> 0 Then
    Select Nvl(Max(����), '����') Into v_���� From ���㷽ʽ Where Nvl(����, 0) = 9;
  
    Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� + Nvl(n_����, 0) Where ����id = ����id_In And ���㷽ʽ = v_����;
    If Sql%NotFound Then
    
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �������, �ɿ���id, У�Ա�־, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, 3, Null, 1, ����id_In, v_����, �˿�ʱ��_In, ����Ա���_In, ����Ա����_In, n_����, ����id_In, -1 * ����id_In,
         n_�ɿ���id, 2, 3);
    End If;
    --��������(���㷽ʽΪNULL��)
    Update ����Ԥ����¼
    Set ��Ԥ�� = ��Ԥ�� - Nvl(n_����, 0)
    Where ����id = ����id_In And ���㷽ʽ Is Null
    Returning Nvl(��Ԥ��, 0) Into n_����ֵ;
  End If;

  If ��������_In = 1 Then
    --   1-��ͨ�˷ѷ�ʽ:
    --�����շѽ��� :��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.."
    v_�������� := ���㷽ʽ_In || '||';
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
    
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
    
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
    
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      v_������� := LTrim(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_����ժҪ := LTrim(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
    
      --���жϡ�������Ƿ�Ϊ�㣬�п����Ѿ����꣬����ʱ���㷽ʽΪ�յ��ؽ�ͳ�����¼�ĳ�Ԥ��֮��Ϊ��
      If v_���㷽ʽ Is Not Null Then
        --If Nvl(n_������, 0) <> 0 Then
        n_������ := Nvl(n_������, 0);
        If Nvl(n_������, 0) <> 0 Then
          Insert Into ����Ԥ����¼
            (ID, ��¼����, NO, ��¼״̬, ����id, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �������, �ɿ���id, У�Ա�־, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, �������, ��������, ��������id, ���ӱ�־)
          Values
            (����Ԥ����¼_Id.Nextval, 3, Null, 1, ����id_In, Null, Null, v_����ժҪ, v_���㷽ʽ, �˿�ʱ��_In, ����Ա���_In, ����Ա����_In, n_������,
             ����id_In, -1 * ����id_In, n_�ɿ���id, У�Ա�־_In, Null, Null, ����_In, ������ˮ��_In, ����˵��_In, Null, 3, ��������id_In, -1);
        
          Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - n_������ Where ����id = ����id_In And ���㷽ʽ Is Null;
          If Sql%NotFound Then
            v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�';
            Raise Err_Item;
          End If;
        End If;
      End If;
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  If ��������_In = 2 And ���㷽ʽ_In Is Not Null Then
    --�������˷ѽ���
    If Nvl(���ԭ����_In, 0) = 1 And Nvl(��������id_In, 0) <> 0 Then
      --��ԭ���㷽ʽΪ�յĽ�����
      --���������Ⲣ������
      Update ����Ԥ����¼
      Set ��Ԥ�� = ��Ԥ��
      Where ����id = ����id_In And ��������id = ��������id_In And Mod(��¼����, 10) <> 1;
    
      Select Sum(��Ԥ��)
      Into n_������
      From ����Ԥ����¼
      Where ����id = ����id_In And ��������id = ��������id_In And Mod(��¼����, 10) <> 1;
    
      Update ����Ԥ����¼
      Set ��Ԥ�� = Nvl(��Ԥ��, 0) + Nvl(n_������, 0)
      Where ����id = ����id_In And ���㷽ʽ Is Null;
    
      Delete ����Ԥ����¼ Where ����id = ����id_In And ��������id = ��������id_In And Mod(��¼����, 10) <> 1;
    End If;
  
    d_����ʱ�� := Sysdate;
    v_������Ա := zl_UserName;
    --   2.�������˷ѽ���:
    v_��ǰ���� := ���㷽ʽ_In;
    v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
    v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
    n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
    v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
    v_������� := LTrim(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
    v_����ժҪ := LTrim(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
  
    If Nvl(n_������, 0) <> 0 Then
      --�ȸ��£�
      Update ����Ԥ����¼
      Set ��Ԥ�� = Nvl(��Ԥ��, 0) + n_������
      Where ����id = ����id_In And �����id = �����id_In And ��������id = ��������id_In And ���㷽ʽ = v_���㷽ʽ;
    
      If Sql%NotFound Then
        Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �������, �ɿ���id, У�Ա�־, �����id,
           ���㿨���, ����, ������ˮ��, ����˵��, �������, ��������, ��������id, ����ʱ��, ������Ա, ���ӱ�־)
        Values
          (n_Ԥ��id, 3, Null, 2, ����id_In, Null, Null, v_����ժҪ, v_���㷽ʽ, �˿�ʱ��_In, ����Ա���_In, ����Ա����_In, n_������, ����id_In,
           -1 * ����id_In, n_�ɿ���id, У�Ա�־_In, �����id_In, Null, ����_In, ������ˮ��_In, ����˵��_In, v_�������, 3, ��������id_In, d_����ʱ��,
           v_������Ա, -1);
      
        --��������������Ϣ����
        Zl_Custom_Balance_Update(n_Ԥ��id);
      
      End If;
    
      Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - n_������ Where ����id = ����id_In And ���㷽ʽ Is Null;
      If Sql%NotFound Then
        v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�';
        Raise Err_Item;
      End If;
    End If;
  End If;

  If ��������_In = 3 Then
    --   3-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
    --3.1����Ƿ��Ѿ�����ҽ����������,������ɾ��
    n_������ := 0;
  
    If У�Ա�־_In = 0 Then
      n_У�Ա�־ := 2;
    Else
      n_У�Ա�־ := 1;
    End If;
  
    For v_ҽ�� In (Select ID, ���㷽ʽ, ��Ԥ��
                 From ����Ԥ����¼ A
                 Where ����id = ����id_In And �����id Is Null And Exists
                  (Select 1 From ���㷽ʽ Where a.���㷽ʽ = ���� And ���� In (3, 4)) And Mod(��¼����, 10) <> 1 And �����id Is Null) Loop
      n_������ := n_������ + Nvl(v_ҽ��.��Ԥ��, 0);
      l_Ԥ��id.Extend;
      l_Ԥ��id(l_Ԥ��id.Count) := v_ҽ��.Id;
    End Loop;
  
    If Nvl(n_������, 0) <> 0 Then
      Update ����Ԥ����¼
      Set ��Ԥ�� = Nvl(��Ԥ��, 0) + Nvl(n_������, 0)
      Where ����id = ����id_In And ���㷽ʽ Is Null;
      If Sql%NotFound Then
        v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�';
        Raise Err_Item;
      End If;
    End If;
  
    If l_Ԥ��id.Count <> 0 Then
      Forall I In 1 .. l_Ԥ��id.Count
        Delete ����Ԥ����¼ Where ID = l_Ԥ��id(I);
    End If;
  
    If ���㷽ʽ_In Is Not Null Then
      v_�������� := ���㷽ʽ_In || '||';
    End If;
    d_����ʱ�� := Sysdate;
    v_������Ա := zl_UserName;
  
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      n_������ := To_Number(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
    
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �������, �ɿ���id, У�Ա�־, ��������, ��������id,
         ����ʱ��, ������Ա, ���ӱ�־)
      Values
        (����Ԥ����¼_Id.Nextval, 3, Null, 2, ����id_In, Null, Null, '���ս���', v_���㷽ʽ, �˿�ʱ��_In, ����Ա���_In, ����Ա����_In, n_������,
         ����id_In, -1 * ����id_In, n_�ɿ���id, n_У�Ա�־, 3, ��������id_In, d_����ʱ��, v_������Ա, -1);
    
      --��������(���㷽ʽΪNULL��)
      Update ����Ԥ����¼
      Set ��Ԥ�� = ��Ԥ�� - n_������
      Where ����id = ����id_In And ���㷽ʽ Is Null
      Returning Nvl(��Ԥ��, 0) Into n_����ֵ;
    
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  --4-���ѿ���������
  If ��������_In = 4 Then
    Null;
  End If;

  If Nvl(����˷�_In, 0) = 0 Then
    Return;
  End If;

  -----------------------------------------------------------------------------------------
  --����շ�,��Ҫ������Ա�ɿ����,Ԥ����¼(���㷽ʽ=NULL)
  If Nvl(����˷�_In, 0) = 1 Then
    Update ����Ԥ����¼ Set У�Ա�־ = 0 Where ����id = ����id_In;
    Return;
  End If;

  --�������ѿ�
  v_ԭ����ids := Null;
  For c_ԭ���� In (Select Distinct a.����id
                From ����Ԥ����¼ A,
                     (Select Distinct ����id
                       From ������ü�¼
                       Where NO In (Select Distinct NO From ������ü�¼ Where ����id = ����id_In And Mod(��¼����, 10) = 1)) B
                Where a.����id = b.����id And Mod(a.��¼����, 10) <> 1 And ��¼״̬ In (3, 1) And a.���㿨��� Is Not Null) Loop
    v_ԭ����ids := Nvl(v_ԭ����ids, '') || ',' || c_ԭ����.����id;
  End Loop;

  If v_ԭ����ids Is Not Null Then
    v_ԭ����ids := Substr(v_ԭ����ids, 2);
  End If;

  For c_���ѿ� In (Select ID, ���㿨���, ���㷽ʽ, ��Ԥ��, ��������id
                From ����Ԥ����¼
                Where ����id = ����id_In And ���ӱ�־ = -1 And Nvl(���㿨���, 0) <> 0) Loop
    n_������� := Nvl(c_���ѿ�.��Ԥ��, 0);
    If n_������� <> 0 Then
      Zl_Square_Update(v_ԭ����ids, c_���ѿ�.Id, ����id_In, n_�ɿ���id, �˿�ʱ��_In, -1 * ����id_In, -1 * n_�������, c_���ѿ�.���㿨���);
    End If;
  End Loop;

  --1.ɾ�����㷽ʽΪNULL��Ԥ����¼
  Delete ����Ԥ����¼ Where ����id = ����id_In And ���㷽ʽ Is Null And Nvl(��Ԥ��, 0) = 0;
  If Sql%NotFound Then
    Select Count(*) Into n_Count From ����Ԥ����¼ A Where ����id = ����id_In And ���㷽ʽ Is Null;
    If n_Count <> 0 Then
      v_Err_Msg := '������δ�˿������,������ɽ������ϲ���!';
    Else
      v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[���ʴ���]���������ϣ�!';
    End If;
    Raise Err_Item;
  End If;

  --������Ϊ��ʱ������һ�����Ϊ0�Ĳ���Ԥ����¼
  Select Count(*) Into n_Count From ����Ԥ����¼ A Where ����id = ����id_In;

  If n_Count = 0 Then
    If v_���㷽ʽ Is Null Then
      Select Max(���㷽ʽ) Into v_���㷽ʽ From ���㷽ʽӦ�� Where Ӧ�ó��� = '�շ�' And Nvl(ȱʡ��־, 0) = 1;
      If v_���㷽ʽ Is Null Then
        Select Nvl(Max(����), '�ֽ�') Into v_���㷽ʽ From ���㷽ʽ Where Nvl(����, 0) = 1;
      End If;
    End If;
  
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �������, �ɿ���id, У�Ա�־, �����id, ���㿨���,
       ����, ������ˮ��, ����˵��, �������, ��������)
    Values
      (����Ԥ����¼_Id.Nextval, 3, Null, 2, ����id_In, Null, Null, Null, v_���㷽ʽ, �˿�ʱ��_In, ����Ա���_In, ����Ա����_In, 0, ����id_In,
       -1 * ����id_In, n_�ɿ���id, 2, Null, Null, Null, Null, ����˵��_In, Null, 2);
  End If;

  --4.�����ܶ�Ҫ�������Ϣ����һ��
  Select Sum(��Ԥ��), Sum(���ʽ��)
  Into n_����ֵ, n_������
  From (Select Sum(��Ԥ��) As ��Ԥ��, 0 As ���ʽ��
         From ����Ԥ����¼
         Where ����id = ����id_In
         Union All
         Select 0, Sum(���ʽ��)
         From ������ü�¼
         Where ����id = ����id_In
         Union All
         Select 0, Sum(���ʽ��) As ���ʽ��
         From סԺ���ü�¼
         Where ����id = ����id_In);

  If Nvl(n_����ֵ, 0) <> Nvl(n_������, 0) Then
    v_Err_Msg := '�����ܶ�������ܶһ��,���ܽ������ϲ���������ϵͳ����Ա��ϵ!';
    Raise Err_Item;
  End If;

  --5.������Ա�ɿ�����
  For c_���� In (Select ����Ա����, ���㷽ʽ, -1 * Sum(��Ԥ��) As ��Ԥ��
               From ����Ԥ����¼
               Where ����id = ����id_In And ���ӱ�־ = -1
               Group By ����Ա����, ���㷽ʽ) Loop
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) - Nvl(c_����.��Ԥ��, 0)
    Where �տ�Ա = c_����.����Ա���� And ���� = 1 And ���㷽ʽ = c_����.���㷽ʽ
    Returning ��� Into n_����ֵ;
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ����
        (�տ�Ա, ���㷽ʽ, ����, ���)
      Values
        (c_����.����Ա����, c_����.���㷽ʽ, 1, -1 * Nvl(c_����.��Ԥ��, 0));
      n_����ֵ := Nvl(c_����.��Ԥ��, 0);
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ��Ա�ɿ����
      Where �տ�Ա = c_����.����Ա���� And ���� = 1 And ���㷽ʽ = c_����.���㷽ʽ And Nvl(���, 0) = 0;
    End If;
  End Loop;

  --2.����ɿ����ݺ��Ҳ����ݼ�У�Ա�־����Ϊ0  
  Select Max(a.�Ƿ����Ʊ��)
  Into n_�Ƿ����Ʊ��
  From ����Ԥ����¼ A,
       (Select Max(b.����id) As ����id
         From ������ü�¼ A, ������ü�¼ B
         Where a.����id = ����id_In And a.No = b.No And b.��¼���� = 1 And b.��¼״̬ In (1, 3)) B
  Where a.����id = b.����id And a.��¼���� In (11, 3);

  Update ����Ԥ����¼ Set У�Ա�־ = 0, ���ӱ�־ = Null, �Ƿ����Ʊ�� = n_�Ƿ����Ʊ�� Where ����id = ����id_In;

  --3.���·���״̬
  Update ������ü�¼ Set ����״̬ = Null Where ����id = ����id_In;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�������תסԺ_Modify;
/

Create Or Replace Procedure Zl_���˹Һ��շ�_Modify
(
  ���ݺ�_In     ������ü�¼.No%Type,
  ����id_In     ������ü�¼.����id%Type,
  ������Ϣ_In   Varchar2,
  ��������_In   Number := 0,
  ��ɱ�־_In   Number := 0,
  ���ɶ���_In   Number := 0,
  �˺�����_In   Number := 1,
  �ջ�Ʊ�ݺ�_In Varchar2 := Null,
  ��������_In   Number := 0,
  ��������id_In ����Ԥ����¼.��������id%Type := Null,
  �����id_In   ����Ԥ����¼.�����id%Type := Null,
  ����_In       ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
  ��ͨ����_In   Number := 0,
  У�Ա�־_In   Number := 2,
  ����Ʊ��_In   ����Ԥ����¼.Ԥ������Ʊ��%Type := 0
) As
  --����:���¸��¹ҺŽ�����Ϣ,��̯��ָ���ĵ�����
  --������Ϣ_In:Ϊ��ʱ,��ʾֻ����Ԥ���ı�־(��Ԥ����ͽ���һ��ʱ,�Ż�ʹ�ô˷�ʽ)
  --��������_In��
  -- 0-��ͨ��ʽ���㣺�����������㷽ʽ,��ʽΪ:"���㷽ʽ,������,�������,����ժҪ|.." ;
  --                 Ҳ�������.Ϊ��ʱֻ��������1��2�ҿ����ID=null��У�Ա�־
  -- 1-��������ֻ�ܴ���һ�����㷽ʽ,��ʽΪ:"���㷽ʽ,������,�������,����ժҪ"
  --           Ҳ�������.Ϊ��ʱֻ��������7,8�ҿ����ID=�����ID_In��У�Ա�־
  -- 2-���ѿ���ֻ�ܴ���һ�����㷽ʽ,��ʽΪ:"���㷽ʽ,������"
  -- 3-Ԥ��֧�����ش���ֻ�ܴ���һ��Ԥ������IDs ��ʽΪ:"������|��Ԥ������ids"
  -- 4-ҽ�����㣺��������,��ʽΪ:"���㷽ʽ,������|.."
  --             Ҳ�������.Ϊ��ʱֻ��������3��4�ҿ����ID=BULL��У�Ա�־
  --                                   �ڶ������㷽ʽ���Ժ�����������_In = 1
  --��������_In����������Ͻ�����룬��һ�����㷽ʽ��ԭ�����¼�ϸ��£���ɾ��������ͬ��������ID��¼,�������±�ʶ����������������Ҫ����
  --��ɱ�־_In������������ɵ��շ�
  --��ͨ����_In: �����ӿڷ��صĽ��㷽ʽ�Ƿ񱣴濨���ID
  --У�Ա�־_In:ҽ��������������ʱ��Ч
  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  n_��id     ����Ԥ����¼.�ɿ���id%Type;
  v_�������� Varchar2(3000);
  v_��ǰ���� Varchar2(200);

  v_�ֽ�          ����Ԥ����¼.���㷽ʽ%Type;
  v_���㷽ʽ      ����Ԥ����¼.���㷽ʽ%Type;
  n_������      ����Ԥ����¼.��Ԥ��%Type;
  v_�������      ����Ԥ����¼.�������%Type;
  v_����ժҪ      ����Ԥ����¼.ժҪ%Type;
  n_��Ԥ��        ����Ԥ����¼.��Ԥ��%Type;
  n_�������      ����Ԥ����¼.��Ԥ��%Type;
  v_��Ԥ������ids Varchar2(1000);
  n_Ԥ��id        ����Ԥ����¼.Id%Type;
  n_ԭԤ��id      ����Ԥ����¼.Id%Type;
  n_����id        ����Ԥ����¼.Id%Type;
  n_��ֵid        ����Ԥ����¼.Id%Type;
  v_����Ա���    ����Ԥ����¼.����Ա���%Type;
  v_����Ա����    ����Ԥ����¼.����Ա����%Type;
  d_Date          ������ü�¼.�Ǽ�ʱ��%Type;
  v_No            ������ü�¼.No%Type;
  n_����id        ������ü�¼.����id%Type;
  n_�����id      ����Ԥ����¼.�����id%Type;
  n_����ֵ        �������.Ԥ�����%Type;
  n_�շ�����      Number;
  n_Count         Number;
  l_Ԥ��id        t_Numlist := t_Numlist();

  n_ԤԼ���ɶ���   Number;
  n_����̨ǩ���Ŷ� Number;
  n_����           Number;
  n_�˺�           Number; --0-�Һţ�1-�˺�
  n_�˷�           Number;
  n_ԤԼ�Һ�       Number; --0-�Һţ�1-ԤԼ�Һ�
  n_ԤԼ��־       Number; --0-�Һţ�1-ԤԼ�Һ�
  n_�Һ�id         ���˹Һż�¼.Id%Type;
  n_��¼id         ���˹Һż�¼.�����¼id%Type;
  v_����           �ҺŰ�������.������Ŀ%Type;
  d_����ʱ��       ���˹Һż�¼.����ʱ��%Type;
  n_���           ������ü�¼.���%Type;
  n_����id         �ҺŰ���.Id%Type;
  n_�ƻ�id         �ҺŰ��żƻ�.Id%Type;
  v_����           ���˹Һż�¼.�ű�%Type;
  n_����           ���˹Һż�¼.����%Type;
  n_��ӡid         Ʊ�ݴ�ӡ����.Id%Type;

  n_�Ŷ�       Number;
  n_�����Ŷ�   Number;
  n_��ʱ����ʾ Number;
  n_��ʱ��     Number;
  v_�ŶӺ���   �ŶӽкŶ���.�ŶӺ���%Type;
  v_�Ŷ����   �ŶӽкŶ���.�Ŷ����%Type;
  v_��������   �ŶӽкŶ���.��������%Type;
  d_�Ŷ�ʱ��   �ŶӽкŶ���.�Ŷ�ʱ��%Type;
  n_�Ƿ����Ʊ�� Number(2);
  n_����         ���ս����¼.����%Type;

  Cursor c_Registinfo(v_״̬ ������ü�¼.��¼״̬%Type) Is
    Select a.����ʱ��, a.�Ǽ�ʱ��, c.����ʱ��, Nvl(c.�Һ���Ŀid, a.�շ�ϸĿid) As ��Ŀid, c.ִ�в���id As ����id, c.ִ���� As ҽ������, d.Id As ҽ��id,
           c.�ű� As ����, c.����, a.����Ա����
    From ������ü�¼ a, ���˹Һż�¼ c, ��Ա�� d
    Where a.��¼���� = 4 And a.No = ���ݺ�_In And a.No = c.No And a.��¼״̬ = v_״̬ And c.ִ���� = d.����(+) And Rownum < 2;
  r_Registrow c_Registinfo%Rowtype;
Begin
  --n_�˺ţ�0-�ҺŸ��£�1-�˺Ÿ���
  Select Max(Id), Decode(Nvl(Max(��¼״̬), 0), 3, 1, 0), Decode(Nvl(Max(��¼����), 0), 2, 1, 0),
         Decode(Nvl(Max(ԤԼ), 0), 1, 1, 0), Nvl(Max(�����¼id), 0), Nvl(Max(�ű�), '0'), Max(����ʱ��), Nvl(Max(����), 0)
  Into n_�Һ�id, n_�˺�, n_ԤԼ�Һ�, n_ԤԼ��־, n_��¼id, v_����, d_����ʱ��, n_����
  From ���˹Һż�¼
  Where No = ���ݺ�_In And ��¼״̬ In (1, 3) And Rownum < 2;

  Select Nvl(Max(���ʷ���), 0), Min(���)
  Into n_����, n_���
  From ������ü�¼
  Where No = ���ݺ�_In And ��¼���� = 4 And
        �Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ������ü�¼ Where No = ���ݺ�_In And ��¼���� = 4)
  Order By ���, �Ǽ�ʱ�� Desc;

  Select Nvl(Max(����), '�ֽ�') Into v_�ֽ� From ���㷽ʽ Where ���� = 1;
  If Nvl(n_���, 0) <> 1 Then
    n_�Һ�id := 0;
    n_�˺�   := 0;
  End If;
  If ��ͨ����_In = 0 And Nvl(�����id_In, 0) <> 0 Then
    n_�����id := �����id_In;
  End If;

  If n_ԤԼ�Һ� = 0 And n_���� = 0 Then
    Select Count(1)
    Into n_�շ�����
    From ������ü�¼
    Where No = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 3 And ����״̬ = 1 And Rownum < 2;
    Select Count(1) Into n_�˷� From ������ü�¼ Where ����id = ����id_In And ��¼״̬ = 2 And Rownum < 2;
  
    Select No, ����Ա���, ����Ա����, ����id, �Ǽ�ʱ��, �ɿ���id
    Into v_No, v_����Ա���, v_����Ա����, n_����id, d_Date, n_��id
    From ������ü�¼
    Where ����id = ����id_In And Rownum < 2;
  
    If Nvl(��������_In, 0) = 0 Then
      --0.��ͨ����
      If ������Ϣ_In Is Null Then
        Update ����Ԥ����¼ a
        Set a.У�Ա�־ = 2
        Where a.����id = ����id_In And Nvl(�����id, 0) = 0 And a.���㷽ʽ Is Not Null And Exists
         (Select 1 From ���㷽ʽ Where a.���㷽ʽ = ���� And ���� In (1, 2));
      Else
        n_Count    := 0;
        v_�������� := ������Ϣ_In || '|';
        While v_�������� Is Not Null Loop
          v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '|') - 1) || ',,,';
          v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
        
          v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
          n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
          If n_�˷� = 1 Then
            n_������ := -1 * n_������;
          Else
            v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
            v_������� := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
          
            v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
            v_����ժҪ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
          End If;
        
          If v_���㷽ʽ Is Null Or n_������ Is Null Then
            v_Err_Msg := '���㷽ʽ����ȷ��';
            Raise Err_Item;
          End If;
        
          If n_�˷� = 0 And Nvl(n_������, 0) = 0 Then
            Delete ����Ԥ����¼
            Where Nvl(�����id, 0) = 0 And ��¼���� = 4 And ����id = ����id_In And ���㷽ʽ = v_���㷽ʽ And
                  Nvl(��������id, 0) = Nvl(��������id_In, 0) Return Nvl(Sum(��Ԥ��), 0) Into n_������;
            n_��Ԥ�� := Nvl(n_��Ԥ��, 0) + n_������;
          Else
            Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
            Insert Into ����Ԥ����¼
              (Id, ��¼����, No, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ�, �Ҳ�, �ɿ���id, Ԥ�����,
               ������λ, �������, ����˵��, У�Ա�־, ��ת��, ��������, �Ự��, ��������id)
              Select n_Ԥ��id, ��¼����, No, ��¼״̬, ����id, ��ҳid, ����id, Nvl(v_����ժҪ, '�Һ��շ�'), v_���㷽ʽ, v_�������, �տ�ʱ��, ����Ա���, ����Ա����,
                     n_������, ����id, �ɿ�, �Ҳ�, �ɿ���id, Ԥ�����, ������λ, �������, ����˵��_In, 2, ��ת��, ��������, �Ự��, n_Ԥ��id
              From ����Ԥ����¼
              Where Nvl(�����id, 0) = Nvl(n_�����id, 0) And ����id = ����id_In And ��¼���� = 4 And Rownum < 2;
            n_��Ԥ�� := Nvl(n_��Ԥ��, 0) - n_������;
          End If;
        
          v_�������� := Substr(v_��������, Instr(v_��������, '|') + 1);
        End Loop;
      End If;
    End If;
  
    If Nvl(��������_In, 0) = 1 Then
      --1.������
      If ������Ϣ_In Is Null Then
        Update ����Ԥ����¼
        Set ������ˮ�� = Nvl(������ˮ��_In, ������ˮ��), ����˵�� = Nvl(����˵��_In, ����˵��), У�Ա�־ = У�Ա�־_In, ���� = Nvl(����_In, ����)
        Where �����id = �����id_In And ����id = ����id_In
        Returning Id Bulk Collect Into l_Ԥ��id;
      
        If У�Ա�־_In = 2 Then
          --���������������½ӿ���Ϣ
          For i In 1 .. l_Ԥ��id.Count Loop
            Zl_Custom_Balance_Update(l_Ԥ��id(i));
          End Loop;
        End If;
      Else
        n_Count    := 0;
        v_�������� := ������Ϣ_In || '|';
        While v_�������� Is Not Null Loop
          v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '|') - 1);
          v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
        
          v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
          n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
        
          If n_�˷� = 1 Then
            n_������ := -1 * n_������;
          End If;
        
          v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
          v_������� := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
        
          v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
          v_����ժҪ := v_��ǰ����;
        
          If v_���㷽ʽ Is Null Or n_������ Is Null Then
            v_Err_Msg := '���㷽ʽ����ȷ��';
            Raise Err_Item;
          End If;
        
          n_Count := n_Count + 1;
          If Nvl(��������_In, 0) = 1 Then
            Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
            Insert Into ����Ԥ����¼
              (Id, ��¼����, No, ��¼״̬, ����id, ��ҳid, ����id, �ɿλ, ��λ������, ��λ�ʺ�, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id,
               �ɿ�, �Ҳ�, �ɿ���id, Ԥ�����, �����id, ����, ������ˮ��, ����˵��, ������λ, �������, У�Ա�־, ��ת��, ��������, �Ự��, ��������id)
              Select n_Ԥ��id, ��¼����, No, ��¼״̬, ����id, ��ҳid, ����id, �ɿλ, ��λ������, ��λ�ʺ�,
                     Decode(n_�˷�, 1, Null, Nvl(v_����ժҪ, ժҪ)), v_���㷽ʽ, Decode(n_�˷�, 1, Null, Nvl(v_�������, �������)), �տ�ʱ��,
                     ����Ա���, ����Ա����, n_������, ����id, �ɿ�, �Ҳ�, �ɿ���id, Ԥ�����, n_�����id, ����_In, ������ˮ��_In, ����˵��_In, ������λ, �������,
                     У�Ա�־_In, ��ת��, ��������, �Ự��, ��������id_In
              From ����Ԥ����¼
              Where ����id = ����id_In And ��¼���� = 4 And Rownum < 2;
          Else
            If n_Count = 1 Then
              --��һ�α���ͬһ��������ID���ܽ���������һ����ɾ�������ļ�¼���Ա����ֱ�Ӹ���
              Select Nvl(Sum(��Ԥ��), 0)
              Into n_��Ԥ��
              From ����Ԥ����¼
              Where Nvl(��������id, 0) = Nvl(��������id_In, 0) And Nvl(�����id, 0) = Nvl(�����id_In, 0) And ����id = ����id_In And
                    У�Ա�־ = 1;
              Update ����Ԥ����¼
              Set ��Ԥ�� = n_������, �����id = Decode(��������_In, 2, Null, n_�����id),
                  ժҪ = Decode(n_�˷�, 1, Null, Nvl(v_����ժҪ, ժҪ)), ������� = Decode(n_�˷�, 1, Null, Nvl(v_�������, �������)),
                  ���㿨��� = Decode(��������_In, 2, n_�����id, Null), ���� = ����_In, ������ˮ�� = ������ˮ��_In, ����˵�� = ����˵��_In,
                  У�Ա�־ = У�Ա�־_In
              Where Nvl(��������id, 0) = Nvl(��������id_In, 0) And ���㷽ʽ = v_���㷽ʽ And ����id = ����id_In And У�Ա�־ = 1 And Rownum < 2
               Return Id Into n_Ԥ��id;
            
              If Sql%Notfound Then
                Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
                Insert Into ����Ԥ����¼
                  (Id, ��¼����, No, ��¼״̬, ����id, ��ҳid, ����id, �ɿλ, ��λ������, ��λ�ʺ�, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��,
                   ����id, �ɿ�, �Ҳ�, �ɿ���id, Ԥ�����, �����id, ����, ������ˮ��, ����˵��, ������λ, �������, У�Ա�־, ��ת��, ��������, �Ự��, ��������id)
                  Select n_Ԥ��id, ��¼����, No, ��¼״̬, ����id, ��ҳid, ����id, �ɿλ, ��λ������, ��λ�ʺ�,
                         Decode(n_�˷�, 1, Null, Nvl(v_����ժҪ, ժҪ)), v_���㷽ʽ,
                         Decode(n_�˷�, 1, Null, Nvl(v_�������, �������)), �տ�ʱ��, ����Ա���, ����Ա����, n_������, ����id, �ɿ�, �Ҳ�, �ɿ���id, Ԥ�����,
                         n_�����id, ����_In, ������ˮ��_In, ����˵��_In, ������λ, �������, У�Ա�־_In, ��ת��, ��������, �Ự��, ��������id_In
                  From ����Ԥ����¼
                  Where ����id = ����id_In And ��¼���� = 4 And Rownum < 2;
              End If;
              Delete From ����Ԥ����¼
              Where Id <> Nvl(n_Ԥ��id, 0) And Nvl(��������id, 0) = Nvl(��������id_In, 0) And Nvl(�����id, 0) = Nvl(�����id_In, 0) And
                    ����id = ����id_In And У�Ա�־ = 1;
            Else
              Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
              Insert Into ����Ԥ����¼
                (Id, ��¼����, No, ��¼״̬, ����id, ��ҳid, ����id, �ɿλ, ��λ������, ��λ�ʺ�, ժҪ, ���, ���㷽ʽ, �������, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��,
                 ����id, �ɿ�, �Ҳ�, �ɿ���id, Ԥ�����, �����id, ����, ������ˮ��, ����˵��, ������λ, �������, У�Ա�־, ��ת��, ��������, �Ự��, ��������id)
                Select n_Ԥ��id, ��¼����, No, ��¼״̬, ����id, ��ҳid, ����id, �ɿλ, ��λ������, ��λ�ʺ�,
                       Decode(n_�˷�, 1, Null, Nvl(v_����ժҪ, ժҪ)), ���, v_���㷽ʽ,
                       Decode(n_�˷�, 1, Null, Nvl(v_�������, �������)), �տ�ʱ��, ����Ա���, ����Ա����, n_������, ����id, �ɿ�, �Ҳ�, �ɿ���id, Ԥ�����,
                       �����id_In, ����_In, ������ˮ��_In, ����˵��_In, ������λ, �������, У�Ա�־_In, ��ת��, ��������, �Ự��, ��������id_In
                From ����Ԥ����¼
                Where ����id = ����id_In And ��¼���� = 4 And Rownum < 2;
            End If;
          End If;
          v_�������� := Substr(v_��������, Instr(v_��������, '|') + 1);
          n_��Ԥ��   := Nvl(n_��Ԥ��, 0) - n_������;
          If У�Ա�־_In = 2 Then
            --���������������½ӿ���Ϣ
            Zl_Custom_Balance_Update(n_Ԥ��id);
          End If;
        End Loop;
      End If;
      If У�Ա�־_In = 2 Then
        Update ����Ԥ����¼
        Set ����ʱ�� = �տ�ʱ��, ������Ա = ����Ա����
        Where ��¼���� = 4 And Nvl(�����id, 0) > 0 And ����id = ����id_In;
      End If;
    End If;
  
    If Nvl(��������_In, 0) = 2 Then
      --2.���ѿ�
      v_�������� := ������Ϣ_In || '|';
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '|') - 1);
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
    
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
      n_������ := To_Number(v_��ǰ����);
      If n_�˷� = 1 Then
        n_������ := -1 * n_������;
      End If;
    
      If v_���㷽ʽ Is Null Or n_������ Is Null Then
        v_Err_Msg := '���㷽ʽ����ȷ��';
        Raise Err_Item;
      End If;
    
      If n_������ <> 0 Then
        If n_�˷� = 0 Then
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (Id, ��¼����, No, ��¼״̬, ����id, ��ҳid, ����id, �ɿλ, ��λ������, ��λ�ʺ�, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ�,
             �Ҳ�, �ɿ���id, Ԥ�����, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, У�Ա�־, ��ת��, ��������, �Ự��, ��������id)
            Select n_Ԥ��id, ��¼����, No, ��¼״̬, ����id, ��ҳid, ����id, �ɿλ, ��λ������, ��λ�ʺ�, ժҪ, v_���㷽ʽ, �������, �տ�ʱ��, ����Ա���, ����Ա����,
                   n_������, ����id, �ɿ�, �Ҳ�, �ɿ���id, Ԥ�����, �����id_In, ����_In, ������ˮ��_In, ����˵��_In, ������λ, �������, 2, ��ת��, ��������, �Ự��,
                   n_Ԥ��id
            From ����Ԥ����¼
            Where ����id = ����id_In And ��¼���� = 4 And Rownum < 2;
        
          Zl_���˿������¼_֧��(�����id_In, ����_In, 0, n_������, n_Ԥ��id, v_����Ա���, v_����Ա����, d_Date);
          n_��Ԥ�� := Nvl(n_��Ԥ��, 0) - n_������;
        Else
          Select Nvl(Id, 0), -1 * Nvl(��Ԥ��, 0)
          Into n_ԭԤ��id, n_�������
          From ����Ԥ����¼
          Where No = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 3 And ���㷽ʽ = v_���㷽ʽ And ���㿨��� = �����id_In;
          If n_ԭԤ��id = 0 Then
            v_Err_Msg := 'δ�ҵ�ԭ�����¼��';
            Raise Err_Item;
          End If;
          If n_������� <> n_������ Then
            v_Err_Msg := '���ѿ��˿��һ�£�';
            Raise Err_Item;
          End If;
        
          Update ����Ԥ����¼
          Set У�Ա�־ = 2
          Where ����id = ����id_In And ���㷽ʽ = v_���㷽ʽ And ���㿨��� = �����id_In
          Returning Id Into n_Ԥ��id;
          If Sql%Notfound Then
            Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
            Insert Into ����Ԥ����¼
              (Id, ��¼����, No, ��¼״̬, ����id, ��ҳid, ����id, �ɿλ, ��λ������, ��λ�ʺ�, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id,
               �ɿ�, �Ҳ�, �ɿ���id, Ԥ�����, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, У�Ա�־, ��ת��, ��������, �Ự��, ��������id)
              Select n_Ԥ��id, ��¼����, No, ��¼״̬, ����id, ��ҳid, ����id, �ɿλ, ��λ������, ��λ�ʺ�, ժҪ, v_���㷽ʽ, �������, �տ�ʱ��, ����Ա���, ����Ա����,
                     n_������, ����id, �ɿ�, �Ҳ�, �ɿ���id, Ԥ�����, �����id_In, ����_In, ������ˮ��_In, ����˵��_In, ������λ, �������, 2, ��ת��, ��������,
                     �Ự��, n_Ԥ��id
              From ����Ԥ����¼
              Where ����id = ����id_In And ��¼���� = 4 And Rownum < 2;
            n_��Ԥ�� := Nvl(n_��Ԥ��, 0) - n_������;
          End If;
          Zl_���˿������¼_�˿�(�����id_In,
                        ����_In,
                        0,
                        -1 * n_������,
                        n_ԭԤ��id,
                        n_Ԥ��id,
                        v_����Ա���,
                        v_����Ա����,
                        d_Date);
        End If;
      End If;
    End If;
  
    If Nvl(��������_In, 0) = 3 Then
      --3.Ԥ��
      v_��ǰ����      := ������Ϣ_In || '|';
      n_��Ԥ��        := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ����      := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      v_��Ԥ������ids := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
    
      If Nvl(n_��Ԥ��, 0) <> 0 Then
        If n_�˷� = 1 Then
          n_������� := n_��Ԥ��;
          For c_Ԥ�� In (Select Max(Id) As Ԥ��id, No, ����id, Ԥ�����, Sum(��Ԥ��) As ��Ԥ��
                       From ����Ԥ����¼
                       Where ��¼���� In (1, 11) And
                             ����id In (Select Distinct ����id From ������ü�¼ Where No = ���ݺ�_In And ��¼���� = 4)
                       Group By No, ����id, Ԥ�����, ����id
                       Having Sum(��Ԥ��) > 0) Loop
            If n_������� > Nvl(c_Ԥ��.��Ԥ��, 0) Then
              n_������ := Nvl(c_Ԥ��.��Ԥ��, 0);
              n_������� := n_������� - Nvl(c_Ԥ��.��Ԥ��, 0);
            Else
              n_������ := n_�������;
              n_������� := 0;
            End If;
          
            Insert Into ����Ԥ����¼
              (Id, No, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���,
               ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������, ��������id, ������Ա, ����ʱ��, У�Ա�־)
              Select ����Ԥ����¼_Id.Nextval, No, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�,
                     d_Date, v_����Ա����, v_����Ա���, -1 * n_������, ����id_In, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4,
                     ��������id, v_����Ա����, d_Date, 2
              From ����Ԥ����¼
              Where Id = c_Ԥ��.Ԥ��id And Rownum < 2;
          
            --����Ԥ���������
            Select Max(Id) Into n_��ֵid From ����Ԥ����¼ Where No = c_Ԥ��.No And ��¼���� = 1 And ��¼״̬ <> 2;
            If Nvl(n_��ֵid, 0) <> 0 Then
              Update Ԥ���������
              Set Ԥ����� = Nvl(Ԥ�����, 0) + n_������
              Where ����id = c_Ԥ��.����id And Ԥ��id = n_��ֵid
              Returning Ԥ����� Into n_����ֵ;
              If Sql%Rowcount = 0 Then
                Insert Into Ԥ���������
                  (Ԥ��id, ����id, Ԥ�����, Ԥ�����)
                Values
                  (n_��ֵid, c_Ԥ��.����id, Nvl(c_Ԥ��.Ԥ�����, 2), n_������);
                n_����ֵ := n_������;
              End If;
              If Nvl(n_����ֵ, 0) = 0 Then
                Delete From Ԥ��������� Where Ԥ��id = n_��ֵid And Nvl(Ԥ�����, 0) = 0;
              End If;
            End If;
            If n_������� = 0 Then
              Exit;
            End If;
          End Loop;
          If Nvl(n_�������, 0) <> 0 Then
            v_Err_Msg := '��Ԥ��������֧����Ԥ�������飡';
            Raise Err_Item;
          End If;
        
          Update �������
          Set Ԥ����� = Nvl(Ԥ�����, 0) + n_��Ԥ��
          Where ����id = n_����id And ���� = 1 And ���� = 1
          Returning Ԥ����� Into n_����ֵ;
          If Sql%Rowcount = 0 Then
            Insert Into ������� (����id, Ԥ�����, ����, ����) Values (n_����id, n_��Ԥ��, 1, 1);
            n_����ֵ := n_��Ԥ��;
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From �������
            Where ����id = n_����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
          End If;
        Else
          Zl_����Ԥ����¼_��Ԥ��(n_����id,
                        ����id_In,
                        n_��Ԥ��,
                        1,
                        v_����Ա���,
                        v_����Ա����,
                        d_Date,
                        v_��Ԥ������ids,
                        4,
                        1,
                        2,
                        1);
          n_��Ԥ�� := 0; --Zl_����Ԥ����¼_��Ԥ�� �г����NULL�Ľ��
        End If;
      End If;
    End If;
  
    If Nvl(��������_In, 0) = 4 Then
      --4.ҽ��
      If ������Ϣ_In Is Null Then
        --Ԥ����ͽ���һ��ʱ,�Ż�ֻ���±�־
        Update ����Ԥ����¼ a
        Set a.У�Ա�־ = У�Ա�־_In
        Where a.����id = ����id_In And Nvl(�����id, 0) = 0 And Exists
         (Select 1 From ���㷽ʽ Where a.���㷽ʽ = ���� And ���� In (3, 4));
        --ҽ����ر�Ĵ���
        Update ���ս�����ϸ Set ��־ = У�Ա�־_In Where ����id = ����id_In;
      Else
        --ɾ������ҽ����������(�������㷽ʽ��ɾ��)
        Select Nvl(Sum(a.��Ԥ��), 0)
        Into n_��Ԥ��
        From ����Ԥ����¼ a
        Where a.����id = ����id_In And a.��¼���� = 4 And Nvl(a.�����id, 0) = 0 And a.���㷽ʽ Is Not Null And Exists
         (Select 1 From ���㷽ʽ Where ���� In (3, 4) And a.���㷽ʽ = ����);
      
        n_Count    := 0;
        v_�������� := ������Ϣ_In || '|';
        While v_�������� Is Not Null Loop
          v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '|') - 1) || ',,,';
          v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
        
          v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
          n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
          If n_�˷� = 1 Then
            n_������ := -1 * n_������;
          End If;
        
          If v_���㷽ʽ Is Null Or n_������ Is Null Then
            v_Err_Msg := '���㷽ʽ����ȷ��';
            Raise Err_Item;
          End If;
        
          n_Count := n_Count + 1;
          If n_Count = 1 Then
            --��һ�α���ͬһ��������ID���ܽ���������һ����ɾ�������ļ�¼���Ա����ֱ�Ӹ���
            Update ����Ԥ����¼
            Set ��Ԥ�� = n_������, ժҪ = 'ҽ���Һ�', У�Ա�־ = У�Ա�־_In, ��������id = Id
            Where Nvl(�����id, 0) = 0 And ���㷽ʽ = v_���㷽ʽ And ����id = ����id_In And Rownum < 2 Return Id Into n_����id;
          
            If Sql%Notfound Then
              Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
              n_����id := n_Ԥ��id;
              Insert Into ����Ԥ����¼
                (Id, ��¼����, No, ��¼״̬, ����id, ��ҳid, ����id, �ɿλ, ��λ������, ��λ�ʺ�, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id,
                 �ɿ�, �Ҳ�, �ɿ���id, Ԥ�����, ������λ, �������, У�Ա�־, ��ת��, ��������, �Ự��, ��������id)
                Select n_Ԥ��id, ��¼����, No, ��¼״̬, ����id, ��ҳid, ����id, �ɿλ, ��λ������, ��λ�ʺ�, 'ҽ���Һ�', v_���㷽ʽ, Null, �տ�ʱ��, ����Ա���,
                       ����Ա����, n_������, ����id, �ɿ�, �Ҳ�, �ɿ���id, Ԥ�����, ������λ, �������, У�Ա�־_In, ��ת��, ��������, �Ự��, n_����id
                From ����Ԥ����¼
                Where ����id = ����id_In And ��¼���� = 4 And Rownum < 2;
            End If;
          
            Delete ����Ԥ����¼ a
            Where a.����id = ����id_In And a.��¼���� = 4 And Id <> Nvl(n_����id, 0) And Nvl(a.�����id, 0) = 0 And
                  a.���㷽ʽ Is Not Null And Exists (Select 1 From ���㷽ʽ Where ���� In (3, 4) And a.���㷽ʽ = ����);
            n_��Ԥ�� := Nvl(n_��Ԥ��, 0) - n_������;
          Else
            Insert Into ����Ԥ����¼
              (Id, ��¼����, No, ��¼״̬, ����id, ��ҳid, ����id, �ɿλ, ��λ������, ��λ�ʺ�, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id,
               �ɿ�, �Ҳ�, �ɿ���id, Ԥ�����, �����id, ����, ������ˮ��, ����˵��, ������λ, �������, У�Ա�־, ��ת��, ��������, �Ự��, ��������id)
              Select ����Ԥ����¼_Id.Nextval, ��¼����, No, ��¼״̬, ����id, ��ҳid, ����id, �ɿλ, ��λ������, ��λ�ʺ�, 'ҽ���Һ�', v_���㷽ʽ, Null, �տ�ʱ��,
                     ����Ա���, ����Ա����, n_������, ����id, �ɿ�, �Ҳ�, �ɿ���id, Ԥ�����, n_�����id, ����_In, ������ˮ��_In, ����˵��_In, ������λ, �������,
                     У�Ա�־_In, ��ת��, ��������, �Ự��, n_����id
              From ����Ԥ����¼
              Where Nvl(��������id, 0) = Nvl(n_����id, 0) And Nvl(�����id, 0) = Nvl(n_�����id, 0) And ����id = ����id_In And ��¼���� = 4 And
                    Rownum < 2;
            n_��Ԥ�� := Nvl(n_��Ԥ��, 0) - n_������;
          End If;
          v_�������� := Substr(v_��������, Instr(v_��������, '|') + 1);
        End Loop;
        Update ���ս�����ϸ Set ��־ = У�Ա�־_In Where ����id = ����id_In;
      End If;
    End If;
  
    If Nvl(n_��Ԥ��, 0) <> 0 Then
      --��δ������Ľ��ۼƵ����㷽ʽΪNULL�ļ�¼�У����������˽����¼����NULL�п۳�
      Update ����Ԥ����¼
      Set ��Ԥ�� = Nvl(��Ԥ��, 0) + n_��Ԥ��
      Where ��¼���� = 4 And ���㷽ʽ Is Null And ����id = ����id_In And У�Ա�־ = 1;
      If Sql%Notfound Then
        Insert Into ����Ԥ����¼
          (Id, ��¼����, No, ��¼״̬, ����id, ��ҳid, ����id, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, У�Ա�־, ��������)
          Select ����Ԥ����¼_Id.Nextval, ��¼����, No, ��¼״̬, ����id, ��ҳid, ����id, Null, �տ�ʱ��, ����Ա���, ����Ա����, n_��Ԥ��, ����id, �ɿ���id, 1,
                 ��������
          From ����Ԥ����¼
          Where ��¼���� = 4 And ����id = ����id_In And Rownum < 2;
      End If;
      n_��Ԥ�� := 0;
    End If;
  
    If Nvl(��ɱ�־_In, 0) = 0 Then
      Return;
    End If;
  
    --1.�ȼ�����Ƿ�һ��
    Select Nvl(Sum(ʵ�ս��), 0) Into n_������ From ������ü�¼ Where ����id = ����id_In;
    If n_������ = 0 Then
      --0�������⴦��
      Update ����Ԥ����¼
      Set ���㷽ʽ = v_�ֽ�, У�Ա�־ = 0
      Where ���㷽ʽ Is Null And Nvl(��Ԥ��, 0) = 0 And ����id = ����id_In;
    End If;
    Select Nvl(Sum(��Ԥ��), 0) Into n_��Ԥ�� From ����Ԥ����¼ Where ����id = ����id_In;
    If n_������ <> n_��Ԥ�� Then
      v_Err_Msg := '������Ϣ����ʵ�ս��(' || n_������ || ')�������(' || n_��Ԥ�� || ')��һ�£�������ɽ��㣡';
      Raise Err_Item;
    End If;
    Delete From ����Ԥ����¼ Where ���㷽ʽ Is Null And Nvl(��Ԥ��, 0) = 0 And ����id = ����id_In;
    --2.����Ƿ����δУ�Եļ�¼
    If Nvl(n_�շ�����, 0) = 1 Then
      Update ����Ԥ����¼ Set У�Ա�־ = 0 Where ����id = ����id_In And ���㷽ʽ Is Null;
    Else
      Select Count(1) Into n_Count From ����Ԥ����¼ Where ����id = ����id_In And ���㷽ʽ Is Null;
      If n_Count > 0 Then
        v_Err_Msg := '���������л�����δ��������ݣ�������ɽ��㣡';
        Raise Err_Item;
      End If;
    End If;
    Select Count(1) Into n_Count From ����Ԥ����¼ Where ����id = ����id_In And У�Ա�־ = 1 And Nvl(��Ԥ��, 0) <> 0;
    If n_Count > 0 Then
      v_Err_Msg := '���������л�����δУ��֧����ʽ��������ɽ��㣡';
      Raise Err_Item;
    End If;
  
    --3.����Ԥ����¼��У�Ա�־
    If Nvl(n_�˷�, 0) = 1 Then
      Select Max(�Ƿ����Ʊ��) Into n_�Ƿ����Ʊ�� From ����Ԥ����¼ 
       Where ����id In (Select ����ID From ������ü�¼ Where no = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 3);
    Else
      n_�Ƿ����Ʊ�� := ����Ʊ��_In;
      If ����Ʊ��_In Is Null Then
        Select Nvl(Max(����), 0) Into n_���� From ���ս����¼ Where ��¼id = ����id_In And ���� = 2;
        n_�Ƿ����Ʊ�� := Zl_Fun_Isstarteinvoice(4, n_����);
      End If;
    End If;
    Update ����Ԥ����¼ Set У�Ա�־ = 0, �Ƿ����Ʊ�� = n_�Ƿ����Ʊ�� Where ����id = ����id_In;
    If Nvl(n_�շ�����, 0) = 1 Then
      Update ����Ԥ����¼ Set ���㷽ʽ = Null Where No = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 3 And У�Ա�־ = 1;
    End If;
  
    --4.���·���״̬,����Ƕ��쳣�������ϣ���ԭʼ��¼���˷Ѽ�¼�ķ���״̬��Ӧ��Ϊ�쳣
    If Nvl(n_�շ�����, 0) = 0 Then
      Update ������ü�¼ Set ����״̬ = 0 Where ����id = ����id_In;
    End If;
  
    --5.���¹Һż�¼��־
    If Nvl(n_�շ�����, 0) = 0 And Nvl(n_���, 0) = 1 Then
      Update ���˹Һż�¼ Set ��¼��־ = 0 Where ��¼״̬ = Decode(n_�˺�, 1, 2, 1) And No = ���ݺ�_In;
    End If;
  
    --6.������Ա�ɿ�����,Not Exists����Ҫ������������Һ����ϵ�ԭʼ���ݽ���ɹ��˵ģ�����Ҳ���˷ѽӿڣ������ܸ��½ɿ����
    If Nvl(n_�շ�����, 0) = 0 Then
      For c_�ɿ� In (Select ���㷽ʽ, ����Ա����, Sum(Nvl(a.��Ԥ��, 0)) As ��Ԥ��
                   From ����Ԥ����¼ a
                   Where a.����id = ����id_In And Mod(a.��¼����, 10) <> 1 And ���㷽ʽ Is Not Null And Not Exists
                    (Select 1
                          From ����Ԥ����¼ b
                          Where b.No = a.No And b.��¼���� = a.��¼���� And b.��¼״̬ = 3 And b.��������id = a.��������id And
                                Nvl(b.У�Ա�־, 0) <> 0)
                   Group By ���㷽ʽ, ����Ա����) Loop
      
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + Nvl(c_�ɿ�.��Ԥ��, 0)
        Where �տ�Ա = c_�ɿ�.����Ա���� And ���� = 1 And ���㷽ʽ = c_�ɿ�.���㷽ʽ;
        If Sql%Rowcount = 0 Then
          Insert Into ��Ա�ɿ����
            (�տ�Ա, ���㷽ʽ, ����, ���)
          Values
            (c_�ɿ�.����Ա����, c_�ɿ�.���㷽ʽ, 1, Nvl(c_�ɿ�.��Ԥ��, 0));
        End If;
      End Loop;
    End If;
  Else
    If Nvl(��ɱ�־_In, 0) = 0 Then
      Return;
    End If;
    Update ������ü�¼ Set ����״̬ = 0 Where ��¼���� = 4 And No = ���ݺ�_In;
    Update ���˹Һż�¼ Set ��¼��־ = 0 Where No = ���ݺ�_In;
  End If;

  If n_�˺� = 1 Then
    --�Һű�������ɹҺ�ʱ���ɶ��У��˺��ڿ�ʼ�˺�ʱ��ȡ������
    Open c_Registinfo(3);
    Fetch c_Registinfo
      Into r_Registrow;
  
    Update ���˹ҺŻ���
    Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - n_ԤԼ��־, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ��־
    Where ���� = Trunc(r_Registrow.����ʱ��) And ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And
          (���� = r_Registrow.���� Or ���� Is Null);
  
    If Sql%Rowcount = 0 Then
      Insert Into ���˹ҺŻ���
        (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, �����ѽ���, ��Լ��)
      Values
        (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
         Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -1, -1 * n_ԤԼ��־, -1 * n_ԤԼ��־);
    End If;
  
    If �˺�����_In = 1 Or (�˺�����_In = 2 And Trunc(r_Registrow.����ʱ��) <> Trunc(Sysdate)) Then
      Delete �Һ����״̬
      Where ״̬ = 1 And
            (����, ���, ����) = (Select �ű�, ����, Trunc(����ʱ��) From ���˹Һż�¼ Where No = ���ݺ�_In And Rownum < 2) Or
            (����, ���, ����) = (Select �ű�, ����, ����ʱ�� From ���˹Һż�¼ Where No = ���ݺ�_In And Rownum < 2);
    Else
      Update �Һ����״̬
      Set ״̬ = 4
      Where ״̬ = 1 And
            (����, ���, ����) = (Select �ű�, ����, Trunc(����ʱ��) From ���˹Һż�¼ Where No = ���ݺ�_In And Rownum < 2) Or
            (����, ���, ����) = (Select �ű�, ����, ����ʱ�� From ���˹Һż�¼ Where No = ���ݺ�_In And Rownum < 2);
    End If;
  
    If n_��¼id <> 0 Then
      If �˺�����_In = 1 Or (�˺�����_In = 2 And Trunc(r_Registrow.����ʱ��) <> Trunc(Sysdate)) Then
        Update �ٴ�������ſ���
        Set �Һ�״̬ = 0, ����Ա���� = Null
        Where �Һ�״̬ = 1 And ��¼id = n_��¼id And ��� = r_Registrow.����;
      
        Update �ٴ�������ſ���
        Set �Һ�״̬ = 4, ����Ա���� = Null
        Where �Һ�״̬ = 1 And ��¼id = n_��¼id And ��ע = To_Char(r_Registrow.����);
      Else
        Update �ٴ�������ſ���
        Set �Һ�״̬ = 4, ����Ա���� = r_Registrow.����Ա����
        Where �Һ�״̬ = 1 And ��¼id = n_��¼id And (��� = r_Registrow.���� Or ��ע = To_Char(r_Registrow.����));
      End If;
    
      Update �ٴ������¼
      Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - n_ԤԼ��־, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ��־
      Where Id = n_��¼id;
    End If;
  
    --ҽ�������ľ���ǼǼ�¼
    Begin
      Delete From ����ǼǼ�¼ Where ����id = n_����id And ����ʱ�� = d_����ʱ�� And ��ҳid Is Null;
    Exception
      When Others Then
        Null;
    End;
  Elsif Nvl(n_�˷�, 0) = 0 Then
    --0-����������;1-��ҽ�������̨�Ŷ�;2-�ȷ���,��ҽ��վ
    If Nvl(n_ԤԼ�Һ�, 0) = 1 Then
      n_ԤԼ���ɶ��� := Zl_To_Number(Zl_Getsysparameter('ԤԼ���ɶ���', 1113));
    End If;
  
    If Nvl(���ɶ���_In, 0) <> 0 And Nvl(n_ԤԼ�Һ�, 0) = 0 Or Nvl(n_ԤԼ���ɶ���, 0) = 1 Then
      For v_�Һ� In (Select Id, ����, ����, ִ����, ִ�в���id, ����ʱ��, �ű�, ����, ԤԼ��ʽ
                   From ���˹Һż�¼
                   Where No = ���ݺ�_In) Loop
        n_����̨ǩ���Ŷ� := Zl_To_Number(Zl_Getsysparameter('����̨ǩ���Ŷ�', 1113, 1, Nvl(v_�Һ�.ִ�в���id, 0)));
        If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Or Nvl(n_ԤԼ���ɶ���, 0) = 1 Then
          Begin
            Select 1,
                   Case
                     When �Ŷ�ʱ�� < Trunc(Sysdate) Then
                      1
                     Else
                      0
                   End
            Into n_�Ŷ�, n_�����Ŷ�
            From �ŶӽкŶ���
            Where ҵ������ = 0 And ҵ��id = v_�Һ�.Id And Rownum <= 1;
          Exception
            When Others Then
              n_�Ŷ� := 0;
          End;
          If n_�Ŷ� = 0 Then
            If n_ԤԼ���ɶ��� = 1 Then
              If Nvl(n_��¼id, 0) = 0 Then
                Select Decode(To_Char(d_����ʱ��, 'D'),
                               '1',
                               '����',
                               '2',
                               '��һ',
                               '3',
                               '�ܶ�',
                               '4',
                               '����',
                               '5',
                               '����',
                               '6',
                               '����',
                               '7',
                               '����',
                               Null)
                Into v_����
                From Dual;
                Select Max(Id) Into n_����id From �ҺŰ��� Where ���� = v_����;
                Select Max(Id)
                Into n_�ƻ�id
                From �ҺŰ��żƻ�
                Where ����id = n_����id And ���ʱ�� Is Not Null And
                      Nvl(��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) =
                      (Select Max(a.��Чʱ��) As ��Ч
                       From �ҺŰ��żƻ� a
                       Where a.���ʱ�� Is Not Null And d_����ʱ�� Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                             a.ʧЧʱ�� And a.����id = n_����id) And
                      d_����ʱ�� Between Nvl(��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And ʧЧʱ��;
              
                If Nvl(n_�ƻ�id, 0) = 0 Then
                  Select Count(Rownum)
                  Into n_��ʱ��
                  From �ҺŰ���ʱ��
                  Where ���� = v_���� And ����id = n_����id And Rownum <= 1;
                Else
                  Select Count(Rownum)
                  Into n_��ʱ��
                  From �Һżƻ�ʱ��
                  Where ���� = v_���� And �ƻ�id = n_�ƻ�id And Rownum <= 1;
                End If;
              Else
                Select Nvl(�Ƿ��ʱ��, 0) Into n_��ʱ�� From �ٴ������¼ Where Id = n_��¼id;
              End If;
              n_��ʱ����ʾ := Nvl(Zl_To_Number(Zl_Getsysparameter(270)), 0);
              If n_��ʱ����ʾ = 1 And n_��ʱ�� = 1 Then
                n_��ʱ����ʾ := 1;
              Else
                n_��ʱ����ʾ := Null;
              End If;
            End If;
            --��������
            --����ִ�в��š���������
            n_�Һ�id   := v_�Һ�.Id;
            v_�������� := v_�Һ�.ִ�в���id;
            v_�ŶӺ��� := Zlgetnextqueue(v_�Һ�.ִ�в���id, n_�Һ�id, v_�Һ�.�ű� || '|' || v_�Һ�.����);
            v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 0);
          
            --�Һ�id_In,����_In,����_In,ȱʡ����_In,��չ_In(������)
            d_�Ŷ�ʱ�� := Zl_Get_Queuedate(n_�Һ�id, v_�Һ�.�ű�, v_�Һ�.����, v_�Һ�.����ʱ��);
            --   ��������_In , ҵ������_In, ҵ��id_In,����id_In,�ŶӺ���_In,�Ŷӱ��_In,��������_In,����ID_IN, ����_In, ҽ������_In,
            Zl_�ŶӽкŶ���_Insert(v_��������,
                             0,
                             n_�Һ�id,
                             v_�Һ�.ִ�в���id,
                             v_�ŶӺ���,
                             Null,
                             v_�Һ�.����,
                             n_����id,
                             v_�Һ�.����,
                             v_�Һ�.ִ����,
                             d_�Ŷ�ʱ��,
                             v_�Һ�.ԤԼ��ʽ,
                             n_��ʱ����ʾ,
                             v_�Ŷ����);
          
            --�Һ������Ŷ�
            If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Then
              Update ���˹Һż�¼ Set ��¼��־ = 1 Where Id = n_�Һ�id;
            End If;
          
          Elsif Nvl(n_�����Ŷ�, 0) = 1 Then
            --���¶��к�
            v_�ŶӺ��� := Zlgetnextqueue(v_�Һ�.ִ�в���id, v_�Һ�.Id, v_�Һ�.�ű� || '|' || Nvl(v_�Һ�.����, 0));
            v_�Ŷ���� := Zlgetsequencenum(0, v_�Һ�.Id, 1);
            --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In, ҽ������_In ,�ŶӺ���_In
            Zl_�ŶӽкŶ���_Update(v_�Һ�.ִ�в���id,
                             0,
                             v_�Һ�.Id,
                             v_�Һ�.ִ�в���id,
                             v_�Һ�.����,
                             v_�Һ�.����,
                             v_�Һ�.ִ����,
                             v_�ŶӺ���,
                             v_�Ŷ����);
          
          Else
            --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In, ҽ������_In ,�ŶӺ���_In
            Zl_�ŶӽкŶ���_Update(v_�Һ�.ִ�в���id,
                             0,
                             v_�Һ�.Id,
                             v_�Һ�.ִ�в���id,
                             v_�Һ�.����,
                             v_�Һ�.����,
                             v_�Һ�.ִ����);
          End If;
        End If;
      End Loop;
    End If;
  End If;

  If �ջ�Ʊ�ݺ�_In Is Not Null And Nvl(n_�˷�, 0) = 1 Then
    --���˹Һŷ�,������Ʊ��
    --�˿��ջ�Ʊ��(�����ϴιҺ�ʹ��Ʊ��,�����ջ�)
    Begin
      --�����һ�δ�ӡ��������ȡ
      Select Id
      Into n_��ӡid
      From (Select b.Id
             From Ʊ��ʹ����ϸ a, Ʊ�ݴ�ӡ���� b
             Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And b.�������� = 4 And b.No = ���ݺ�_In
             Order By a.ʹ��ʱ�� Desc)
      Where Rownum < 2;
    Exception
      When Others Then
        n_��ӡid := Null;
    End;
  
    --���ջ�ԭƱ��
    If n_��ӡid Is Not Null Then
      Begin
        Insert Into Ʊ��ʹ����ϸ
          (Id, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
          Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, v_����Ա����, Ʊ�ݽ��
          From Ʊ��ʹ����ϸ
          Where ��ӡid = n_��ӡid And ���� = 1 And Instr(',' || �ջ�Ʊ�ݺ�_In || ',', ',' || ���� || ',') > 0;
      Exception
        When Others Then
          Delete From Ʊ��ʹ����ϸ Where ��ӡid = n_��ӡid And ���� = 2 And ԭ�� = 2;
          Insert Into Ʊ��ʹ����ϸ
            (Id, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
            Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, v_����Ա����, Ʊ�ݽ��
            From Ʊ��ʹ����ϸ
            Where ��ӡid = n_��ӡid And ���� = 1 And Instr(',' || �ջ�Ʊ�ݺ�_In || ',', ',' || ���� || ',') > 0;
      End;
    End If;
  End If;

  --��Ϣ����
  If Nvl(n_�˷�, 0) = 0 Then
    Begin
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
        Using 1, n_�Һ�id;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_001(n_�Һ�id, ���ݺ�_In);
  Else
    Begin
      Select Id Into n_�Һ�id From ���˹Һż�¼ Where No = ���ݺ�_In And ��¼״̬ = 2 And Rownum < 2;
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
        Using 2, ���ݺ�_In;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_003(n_�Һ�id, ���ݺ�_In);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_���˹Һ��շ�_Modify;
/

Create Or Replace Procedure Zl_Third_Payment
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is

  --------------------------------------------------------------------------------------------------
  --����:�����ӿ�֧��
  --���:Xml_In:
  --<IN>
  --        <NO></NO>                       //�շѵ��ݺŴ�,���ŷָ�������ݺ�
  --        <JE></JE>                       //�ܽ��
  --        <BRID>����ID</BRID>             //����ID
  --        <XM>����</XM>                   //����
  --        <SFZH>���֤��</SFZH>           //���֤��
  --        <SFGH></SFGH>                   //�Ƿ�Һŵ�
  --        <WCJE>����</WCJE>             //������ʱ,���ܽ��-���ν�������ܶ�Ϊ׼
  --        <JSMS>1</JSMS>          //����ģʽ��0-��ͨģʽ��1-�첽����ģʽ
  --        <CZLX>0</CZLX>          //�������ͣ�����ģʽΪ1ʱ���룬0-��ʼ���㣬1-��ɽ��㣬2-���˽���
  --        <JZID>1</JZID>          //����ID����������Ϊ1��2ʱ����
  --    <ZFBZH>֧�������ں�UserID</ZFBZH>
  --    <ZFBXCY>֧����С����UserID</ZFBXCY>
  --    <WXGZHID>΢�Ź��ں�OpenID</WXGZH>
  --    <WXXCXID>΢��С����OpenID</WXXCXID>
  --        <JSLIST>          //�����б���������Ϊ2ʱ�ɲ�����
  --         <JS>
  --              <JSKLB>֧�������</JSKLB >
  --              <JSKH>֧������</ JSKH >
  --              <JSFS>֧����ʽ</JSFS> //֧����ʽ:�ֽ�;֧Ʊ,�����������,���Դ���
  --              <JSJE>֧�����</JSJE>
  --              <JYLSH>������ˮ��</JYLSH>
  --              <JYSM>����˵��</JYSM>
  --              <ZY>ժҪ</ZY>
  --              <SFCYJ>�Ƿ��Ԥ��</SFCYJ>  //�ʳ�Ԥ��ʱ,ֻ��JSJE�ڵ�:1-��Ԥ��
  --              <SFXFK>�Ƿ����ѿ�</SFXFK>  //(1-�����ѿ�),���ѿ�ʱ,������㿨���,���㿨��,������Ƚӵ�
  --              <DJH>S0000001</DJH> //�ֵ��ݽ���ʱ����
  --              <EXPENDLIST>  //��չ������Ϣ
  --                  <EXPEND>
  --                        <JYMC >��������</��������>
  --                        <JYLR>��������</JYLR>
  --                  </EXPEND>
  --              </EXPENDLIST>
  --         </JS>
  --       </JSLIST >
  --</IN>

  --����:Xml_Out
  --  <OUTPUT>
  --    <JZID>����ID</JZID>       //����ID
  --    <CZSJ>����ʱ��</CZSJ>          //HIS�ĵǼ�ʱ��
  --  <KPBZ>��Ʊ��־</KPBZ> //1-�ɹ����ߵ���Ʊ��;0-δ��Ʊ�ɹ���־
  --  <URL>H5ҳ��URL</URL>
  --  <NETURL>����H5ҳ��URL</NETURL>
  --  <FPTT>��Ʊ̧ͷ</FPTT>        //��������
  --  <FPH>��Ʊ��</FPH>             //��Ʊ���
  --  <FPJE>��Ʊ���</FPJE>        //100.00
  --  <KPRQ>��Ʊ����</KPRQ>   //yyyy-mm-dd
  --    �D�D�������д�������˵����ȷִ��
  --    <ERROR>
  --      <MSG>������Ϣ</MSG>
  --    </ERROR>
  --  </OUTPUT>
  --------------------------------------------------------------------------------------------------
  v_Nos      Varchar2(4000);
  n_������ ������ü�¼.ʵ�ս��%Type;

  n_�����id   ҽ�ƿ����.Id%Type;
  n_���㿨��� ����Ԥ����¼.���㿨���%Type;
  v_���㷽ʽ   Varchar2(2000);
  n_����id     ������ü�¼.����id%Type;
  v_���֤��   ������Ϣ.���֤��%Type;
  v_����       ������ü�¼.����%Type;
  v_�Ա�       ������ü�¼.�Ա�%Type;
  v_����       ������ü�¼.����%Type;
  n_����ģʽ   Number(1); --0-��ͨģʽ��1-�첽����ģʽ
  n_��������   Number(1); --����ģʽΪ1ʱ���룬0 - ��ʼ���㣬1 - ��ɽ��㣬2 - ���˽���

  n_��������id ����Ԥ����¼.��������id%Type;
  n_Ԥ��id     ����Ԥ����¼.Id%Type;
  n_���ѿ�     Number;
  n_ɾ��ԭ���� Number;

  v_ҽ�Ƹ��ʽ���� ҽ�Ƹ��ʽ.����%Type;
  v_���ʽ         ҽ�Ƹ��ʽ.����%Type;
  v_����Ա����       ������ü�¼.����Ա���%Type;
  v_����Ա����       ������ü�¼.����Ա����%Type;
  n_����id           ������ü�¼.����id%Type;
  n_���ʽ��         ������ü�¼.���ʽ��%Type;
  d_�շ�ʱ��         ����Ԥ����¼.�տ�ʱ��%Type;
  n_���ѿ�id         ���ѿ���Ϣ.Id%Type;
  v_�շѽ���         Varchar2(2000);
  v_��ͨ����         Varchar2(4000);
  n_�Ƿ�Һ�         Number(3);
  n_Ԥ��֧��         ������ü�¼.ʵ�ս��%Type;
  n_��֧ͨ��         ������ü�¼.ʵ�ս��%Type;
  v_���㿨��         ����Ԥ����¼.����%Type;
  v_������ˮ��       ����Ԥ����¼.������ˮ��%Type;
  v_����˵��         ����Ԥ����¼.����˵��%Type;
  v_ժҪ             ����Ԥ����¼.ժҪ%Type;
  n_����id           �ҺŰ���.����id%Type;
  n_��Ŀid           �ҺŰ���.��Ŀid%Type;
  n_ҽ��id           �ҺŰ���.ҽ��id%Type;
  v_ҽ������         �ҺŰ���.ҽ������%Type;
  v_����             �ҺŰ���.����%Type;
  n_�����           ������Ϣ.�����%Type;
  d_����ʱ��         ���˹Һż�¼.����ʱ��%Type;
  v_�ѱ�             ������Ϣ.�ѱ�%Type;
  n_����             ���˹Һż�¼.����%Type;
  v_Para             Varchar2(500);
  n_�Һ�ģʽ         Number(3);
  d_����ʱ��         Date;
  v_��ʱ���㷽ʽ     ����Ԥ����¼.���㷽ʽ%Type;
  n_�����¼id       �ٴ������¼.Id%Type;
  n_���             ������ü�¼.���%Type;
  v_������Ŀid       Varchar2(500);
  v_��������         Varchar2(500);
  v_����ֵ           Varchar2(100);
  n_Cursor           Number(3);
  n_ʵ�ս��         ������ü�¼.ʵ�ս��%Type;
  v_ʵ��             Varchar2(500);
  n_��������         ������ü�¼.��������%Type;
  n_���˿���id       ������ü�¼.���˿���id%Type;
  n_ִ�в���id       ������ü�¼.ִ�в���id%Type;
  v_No               ������ü�¼.No%Type;
  v_��ͨ�ȼ�         Varchar2(100);
  v_Pricegrade       Varchar2(500);
  n_ҽ��֧��         ����Ԥ����¼.��Ԥ��%Type;
  n_Exists           Number;
  v_վ��             ���ű�.վ��%Type;
  n_�������         ����Ԥ����¼.�������%Type;
  n_ҵ������         �������׼�¼.ҵ������%Type;
  v_Temp             Varchar2(32767); --��ʱXML
  x_Templet          Xmltype; --ģ��XML
  v_�����           �������׼�¼.���%Type;
  v_����Ա           ������ü�¼.����Ա����%Type;
  v_��ҩ����         Varchar2(4000);
  n_����           ����Ԥ����¼.��Ԥ��%Type;
  n_��������         Number;
  n_ʵ����           Number(3);
  n_��֤             Number(3);
  n_Step             Number(2);
  n_Checkmzlg        Number(2);
  n_Count            Number(2);

  n_�Ƿ����Ʊ��       ����Ԥ����¼.�Ƿ����Ʊ��%Type;
  v_֧�������ں�userid Varchar2(100);
  v_֧����С����userid Varchar2(100);
  v_΢�Ź��ں�openid   Varchar2(100);
  v_΢��С����openid   Varchar2(100);
  n_��Ʊ��־           Number(2);

  v_��Ʊ���� Varchar2(20);
  v_�������� ����Ʊ��ʹ�ü�¼.����%Type;
  v_��Ʊ��� ����Ʊ��ʹ�ü�¼.����%Type;
  n_��Ʊ��� ����Ʊ��ʹ�ü�¼.Ʊ�ݽ��%Type;
  v_Url      ����Ʊ��ʹ�ü�¼.Url����%Type;
  v_Url����  ����Ʊ��ʹ�ü�¼.Url����%Type;

  Type Price_Type Is Record(
    ��Ŀid ������ü�¼.�շ�ϸĿid%Type,
    ����   ������ü�¼.����%Type,
    ����   ������ü�¼.��׼����%Type,
    Ӧ��   ������ü�¼.Ӧ�ս��%Type,
    ʵ��   ������ü�¼.ʵ�ս��%Type); --����Price��¼���� 
  Type Price_Type_Array Is Table Of Price_Type Index By Binary_Integer; --������Price��¼���������� 
  Price_Rec       Price_Type; --�������������ͣ�Price��¼����
  Price_Rec_Array Price_Type_Array; --�������������ͣ����Price��¼����������

  v_Err_Msg Varchar2(200);
  Err_Item    Exception;
  Err_Special Exception;

  --��ȡ���������
  Function Get_Cardname
  (
    �����_In Varchar2,
    ���ѿ�_In Number
  ) Return Varchar2 As
    v_����       ҽ�ƿ����.����%Type;
    n_By_Id_Find Number;
  
    v_Err_Msg Varchar2(200);
    Err_Item Exception;
  Begin
    If �����_In Is Null Then
      Return Null;
    End If;
  
    Select Decode(Translate(Nvl(�����_In, 'abcd'), '#1234567890', '#'), Null, 1, 0) Into n_By_Id_Find From Dual;
  
    --1.�������ID����ҽ�ƿ�
    If Nvl(���ѿ�_In, 0) = 0 And Nvl(n_By_Id_Find, 0) = 1 Then
      Begin
        Select ���� Into v_���� From ҽ�ƿ���� Where ID = To_Number(�����_In);
      Exception
        When Others Then
          v_Err_Msg := 'δ�ҵ�ָ���Ľ���֧����Ϣ��';
          Raise Err_Item;
      End;
    End If;
  
    --2.�������Ʋ���ҽ�ƿ�
    If Nvl(���ѿ�_In, 0) = 0 And Nvl(n_By_Id_Find, 0) = 0 Then
      Begin
        Select ���� Into v_���� From ҽ�ƿ���� Where ���� = �����_In;
      Exception
        When Others Then
          v_Err_Msg := �����_In || '������!';
          Raise Err_Item;
      End;
    End If;
  
    --3.�������ID�������ѿ�
    If Nvl(���ѿ�_In, 0) = 1 And Nvl(n_By_Id_Find, 0) = 1 Then
      Begin
        Select ���� Into v_���� From ���ѿ����Ŀ¼ Where ��� = To_Number(�����_In);
      Exception
        When Others Then
          v_Err_Msg := 'δ�ҵ�ָ���Ľ���֧����Ϣ��';
          Raise Err_Item;
      End;
    End If;
  
    --3.�������Ʋ������ѿ�
    If Nvl(���ѿ�_In, 0) = 1 And Nvl(n_By_Id_Find, 0) = 0 Then
      Begin
        Select ���� Into v_���� From ���ѿ����Ŀ¼ Where ���� = �����_In;
      Exception
        When Others Then
          v_Err_Msg := �����_In || '������!';
          Raise Err_Item;
      End;
    End If;
  
    Return v_����;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  End;

  --��ȡ�����ID
  Function Get_Cardtypeid
  (
    �����_In    Varchar2,
    ���ѿ�_In    Number,
    ���㷽ʽ_Out In Out ҽ�ƿ����.���㷽ʽ%Type
  ) Return Number As
    n_�����id ҽ�ƿ����.Id%Type;
    v_����     ҽ�ƿ����.����%Type;
    n_����     ҽ�ƿ����.�Ƿ�����%Type;
    v_���㷽ʽ ҽ�ƿ����.���㷽ʽ%Type;
  
    n_By_Id_Find Number;
    v_Err_Msg    Varchar2(200);
    Err_Item Exception;
  Begin
    If �����_In Is Null Then
      Return 0;
    End If;
  
    Select Decode(Translate(Nvl(�����_In, 'abcd'), '#1234567890', '#'), Null, 1, 0) Into n_By_Id_Find From Dual;
  
    --1.�������ID����ҽ�ƿ�
    If Nvl(���ѿ�_In, 0) = 0 And Nvl(n_By_Id_Find, 0) = 1 Then
      Begin
        Select ID, ���㷽ʽ, ����, �Ƿ�����
        Into n_�����id, v_����, v_���㷽ʽ, n_����
        From ҽ�ƿ����
        Where ID = To_Number(�����_In);
      Exception
        When Others Then
          v_Err_Msg := 'δ�ҵ�ָ���Ľ���֧����Ϣ��';
          Raise Err_Item;
      End;
    End If;
  
    --2.�������Ʋ���ҽ�ƿ�
    If Nvl(���ѿ�_In, 0) = 0 And Nvl(n_By_Id_Find, 0) = 0 Then
      Begin
        Select ID, ���㷽ʽ, ����, �Ƿ�����
        Into n_�����id, v_����, v_���㷽ʽ, n_����
        From ҽ�ƿ����
        Where ���� = �����_In;
      Exception
        When Others Then
          v_Err_Msg := �����_In || '������!';
          Raise Err_Item;
      End;
    End If;
  
    --3.�������ID�������ѿ�
    If Nvl(���ѿ�_In, 0) = 1 And Nvl(n_By_Id_Find, 0) = 1 Then
      Begin
        Select ���, ���㷽ʽ, ����, ����
        Into n_�����id, v_����, v_���㷽ʽ, n_����
        From ���ѿ����Ŀ¼
        Where ��� = To_Number(�����_In);
      Exception
        When Others Then
          v_Err_Msg := 'δ�ҵ�ָ���Ľ���֧����Ϣ��';
          Raise Err_Item;
      End;
    End If;
  
    --3.�������Ʋ������ѿ�
    If Nvl(���ѿ�_In, 0) = 1 And Nvl(n_By_Id_Find, 0) = 0 Then
      Begin
        Select ���, ���㷽ʽ, ����, ����
        Into n_�����id, v_����, v_���㷽ʽ, n_����
        From ���ѿ����Ŀ¼
        Where ���� = �����_In;
      Exception
        When Others Then
          v_Err_Msg := �����_In || '������!';
          Raise Err_Item;
      End;
    End If;
  
    If Nvl(n_����, 0) = 0 Then
      v_Err_Msg := v_���� || 'δ���ã���������нɷѣ�';
      Raise Err_Item;
    End If;
  
    If ���㷽ʽ_Out Is Null Then
      If v_���㷽ʽ Is Null Then
        v_Err_Msg := v_���� || 'δ���ý��㷽ʽ����������нɷѣ�';
        Raise Err_Item;
      End If;
    
      ���㷽ʽ_Out := v_���㷽ʽ;
    End If;
  
    Return n_�����id;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  End;

  Procedure Thirdcard_Balance
  (
    ����id_In     ����Ԥ����¼.����id%Type,
    ����id_In     ����Ԥ����¼.����id%Type,
    ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type,
    �����_In     ����Ԥ����¼.�����id%Type,
    ����_In       ����Ԥ����¼.����%Type,
    ֧�����_In   ����Ԥ����¼.��Ԥ��%Type,
    ������ˮ��_In ����Ԥ����¼.������ˮ��%Type,
    ����˵��_In   ����Ԥ����¼.����˵��%Type,
    No_In         ����Ԥ����¼.No%Type,
    ��������id_In ����Ԥ����¼.��������id%Type,
    Xmlexpned_In  Xmltype,
    ����ģʽ_In   Number := 0,
    ��������_In   Number := 0,
    ɾ��ԭ����_In Number := 0
  ) Is
    --��Σ�
    --         ����ģʽ_in   0-��ͨģʽ��1-�첽����ģʽ
    --        ��������_in   ����ģʽΪ1ʱ���룬0 - ��ʼ���㣬1 - ��ɽ��㣬2 - ���˽���
    --        ɾ��ԭ����_in ��������_InΪ1ʱ��Ч��������㷽ʽʱ���ö�θù���
    v_�շѽ��� Varchar2(2000);
    n_У�Ա�־ ����Ԥ����¼.У�Ա�־%Type;
  Begin
    If Nvl(����ģʽ_In, 0) = 1 And Nvl(��������_In, 0) = 0 Then
      n_У�Ա�־ := 1;
    Else
      n_У�Ա�־ := 2;
    End If;
  
    --���㷽ʽ|������|�������|����ժҪ|���ݺ�|�Ƿ���ͨ����
    v_�շѽ��� := ���㷽ʽ_In || '|' || ֧�����_In || '| | |' || No_In || '|0';
    Zl_�����շѽ���_Modify(4, ����id_In, ����id_In, v_�շѽ���, 0, 0, �����_In, ����_In, ������ˮ��_In, ����˵��_In, 0, 0, 0, 0, Null, Null, 0,
                     ��������id_In, ɾ��ԭ����_In, n_У�Ա�־);
  
    If Nvl(����ģʽ_In, 0) = 1 And Nvl(��������_In, 0) = 0 Then
      Return;
    End If;
  
    --������չ������Ϣ 
    For c_��չ In (Select Extractvalue(j.Column_Value, '/EXPEND/JYMC') As Jymc,
                        Extractvalue(j.Column_Value, '/EXPEND/JYLR') As Jylr
                 From Table(Xmlsequence(Extract(Xmlexpned_In, '/EXPENDLIST/EXPEND'))) J) Loop
      Zl_�������㽻��_Insert(�����_In, 0, ����_In, ����id_In, c_��չ.Jymc || '|' || c_��չ.Jylr);
    End Loop;
  End Thirdcard_Balance;

  Procedure Squarecard_Balance
  (
    ����id_In     ����Ԥ����¼.����id%Type,
    ����id_In     ����Ԥ����¼.����id%Type,
    �����_In     ����Ԥ����¼.�����id%Type,
    ����_In       ����Ԥ����¼.����%Type,
    ֧�����_In   ����Ԥ����¼.��Ԥ��%Type,
    ������ˮ��_In ����Ԥ����¼.������ˮ��%Type,
    ����˵��_In   ����Ԥ����¼.����˵��%Type,
    Xmlexpned_In  Xmltype
  ) Is
    v_�շѽ��� Varchar2(2000);
    n_���ѿ�id ���ѿ���Ϣ.Id%Type;
  Begin
    Select ID
    Into n_���ѿ�id
    From ���ѿ���Ϣ
    Where �ӿڱ�� = �����_In And ���� = ����_In And
          ��� = (Select Max(���) From ���ѿ���Ϣ Where �ӿڱ�� = �����_In And ���� = ����_In);
  
    --���㷽ʽ_IN��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||....
    v_�շѽ��� := �����_In || '|' || ����_In || '|' || n_���ѿ�id || '|' || ֧�����_In;
    Zl_�����շѽ���_Modify(3, ����id_In, ����id_In, v_�շѽ���, 0, 0, �����_In, ����_In, ������ˮ��_In, ����˵��_In, 0, 0, 0, 0);
  
    --������չ������Ϣ
    For c_��չ In (Select Extractvalue(j.Column_Value, '/EXPEND/JYMC') As Jymc,
                        Extractvalue(j.Column_Value, '/EXPEND/JYLR') As Jylr
                 From Table(Xmlsequence(Extract(Xmlexpned_In, '/EXPENDLIST/EXPEND'))) J) Loop
      Zl_�������㽻��_Insert(�����_In, 1, ����_In, ����id_In, c_��չ.Jymc || '|' || c_��չ.Jylr, 0);
    End Loop;
  End Squarecard_Balance;

Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/NO'), To_Number(Extractvalue(Value(A), 'IN/BRID')),
         To_Number(Extractvalue(Value(A), 'IN/JE')), To_Number(Extractvalue(Value(A), 'IN/WCJE')),
         To_Number(Extractvalue(Value(A), 'IN/SFGH')), Extractvalue(Value(A), 'IN/ZD'),
         Extractvalue(Value(A), 'IN/SFZH'), Extractvalue(Value(A), 'IN/XM'), Extractvalue(Value(A), 'IN/JSMS'),
         Extractvalue(Value(A), 'IN/CZLX'), Extractvalue(Value(A), 'IN/JZID'), Extractvalue(Value(A), 'IN/ZFBZH'),
         Extractvalue(Value(A), 'IN/ZFBXCY'), Extractvalue(Value(A), 'IN/WXGZHID'), Extractvalue(Value(A), 'IN/WXXCXID')
  Into v_Nos, n_����id, n_������, n_����, n_�Ƿ�Һ�, v_վ��, v_���֤��, v_����, n_����ģʽ, n_��������, n_����id, v_֧�������ں�userid, v_֧����С����userid,
       v_΢�Ź��ں�openid, v_΢��С����openid
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If Nvl(n_����id, 0) = 0 And Not v_���֤�� Is Null And Not v_���� Is Null Then
    n_����id := Zl_Third_Getpatiid(v_���֤��, v_����);
  End If;

  --0.��ؼ��
  If Nvl(n_����id, 0) = 0 Then
    v_Err_Msg := '������Чʶ�������,������ɷ�!';
    Raise Err_Item;
  End If;

  If Not v_֧�������ں�userid Is Null Then
    Zl_������Ϣ�ӱ�_Update(n_����id, Upper('֧�������ں�UserID'), v_֧�������ں�userid);
  End If;

  If Not v_֧����С����userid Is Null Then
    Zl_������Ϣ�ӱ�_Update(n_����id, Upper('֧����С����UserID'), v_֧�������ں�userid);
  End If;

  If Not v_΢�Ź��ں�openid Is Null Then
    Zl_������Ϣ�ӱ�_Update(n_����id, Upper('΢�Ź��ں�OpenID'), v_֧�������ں�userid);
  End If;

  If Not v_΢��С����openid Is Null Then
    Zl_������Ϣ�ӱ�_Update(n_����id, Upper('΢��С����OpenID'), v_֧�������ں�userid);
  End If;

  If v_Nos Is Null Then
    v_Err_Msg := 'û��ָ����ص��շѵ���,������ɷ�!';
    Raise Err_Item;
  End If;

  If Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) <> 0 Then
    If Nvl(n_����id, 0) = 0 Then
      v_Err_Msg := 'û��ָ����صĽ������ݣ�';
      Raise Err_Item;
    End If;
  
    Begin
      Select �տ�ʱ��
      Into d_�շ�ʱ��
      From ����Ԥ����¼
      Where ����id = n_����id And Nvl(У�Ա�־, 0) = 1 And Rownum < 2;
    Exception
      When Others Then
        v_Err_Msg := 'û���ҵ�ָ������ؽ������ݣ������ѱ�����';
        Raise Err_Item;
    End;
  End If;

  If Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) = 2 Then
    If Nvl(n_�Ƿ�Һ�, 0) = 0 Then
      --ɾ���������ݣ��ָ����۵�
      Zl_���˽����¼_Delete(n_����id);
      Zl_�����շѽ���_Cancel(n_����id);
    Else
      Zl_���˹Һż�¼_Cancel(n_����id);
    End If;
  
    v_Temp := '<CZSJ>' || To_Char(d_�շ�ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<JZID>' || n_����id || '</JZID>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<KPBZ>' || 0 || '</KPBZ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<URL>' || '' || '</URL>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<NETURL>' || '' || '</NETURL>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<FPTT>' || v_���� || '</FPTT>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<FPH>' || '' || '</FPH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<FPJE>' || '' || '</FPJE>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<KPRQ>' || '' || '</KPRQ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    Xml_Out := x_Templet;
    Return;
  End If;

  --��Աid,��Ա���,��Ա����
  v_Temp := Zl_Identity(1);
  If Nvl(v_Temp, '0') = '0' Or Nvl(v_Temp, '_') = '_' Then
    v_Err_Msg := 'ϵͳ�����ϱ���Ч�Ĳ���Ա,������ɷ�!';
    Raise Err_Item;
  End If;
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_����Ա���� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_����Ա���� := v_Temp;

  Begin
    Select b.����, b.����, a.����, a.�Ա�, a.����
    Into v_ҽ�Ƹ��ʽ����, v_���ʽ, v_����, v_�Ա�, v_����
    From ������Ϣ A, ҽ�Ƹ��ʽ B
    Where a.ҽ�Ƹ��ʽ = b.����(+) And a.����id = n_����id;
  Exception
    When Others Then
      v_Err_Msg := 'ָ���Ľɷѵ����в�����Чʶ����,������ɷ�!';
      Raise Err_Item;
  End;

  n_Checkmzlg := To_Number(Nvl(zl_GetSysParameter(323), '0'));
  Select Decode(Nvl(n_�Ƿ�Һ�, 0), 0, 3, 4) Into n_ҵ������ From Dual;
  If Nvl(n_����ģʽ, 0) = 0 Or Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) = 0 Then
    For c_���׼�¼ In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As ���㿨���,
                          Extractvalue(b.Column_Value, '/JS/JSFS') As ���㷽ʽ,
                          Extractvalue(b.Column_Value, '/JS/JYLSH') As ������ˮ��,
                          Extractvalue(b.Column_Value, '/JS/JSKH') As ���㿨��, Extractvalue(b.Column_Value, '/JS/ZY') As ժҪ,
                          Extractvalue(b.Column_Value, '/JS/SFXFK') As �Ƿ����ѿ�,
                          Extractvalue(b.Column_Value, '/JS/SFCYJ') As �Ƿ��Ԥ��
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
    
      If Nvl(c_���׼�¼.�Ƿ��Ԥ��, 0) = 0 Then
        If c_���׼�¼.���㿨��� Is Null Then
          v_����� := c_���׼�¼.���㷽ʽ;
        Else
          v_����� := Get_Cardname(c_���׼�¼.���㿨���, c_���׼�¼.�Ƿ����ѿ�);
        End If;
      
        If v_����� Is Null Then
          v_Err_Msg := '��֧�ֵĽ��㷽ʽ,���飡';
          Raise Err_Item;
        End If;
      
        --����һ�����㷽ʽ�ż�齻����
        n_Step := Nvl(n_Step, 0) + 1;
        If Zl_Fun_�������׼�¼_Locked(v_�����, c_���׼�¼.������ˮ��, c_���׼�¼.���㿨��, c_���׼�¼.ժҪ, n_ҵ������) = 0 And n_Step = 1 Then
          v_Err_Msg := '������ˮ��Ϊ:' || c_���׼�¼.������ˮ�� || '�Ľ������ڽ����У��������ٴ��ύ�˽���!';
          Raise Err_Special;
        End If;
      Else
        If Nvl(n_Checkmzlg, 0) <> 0 Then
          Select Count(1)
          Into n_Count
          From ������ҳ A, ������Ϣ B
          Where a.����id = n_����id And a.�������� = 1 And a.����id = b.����id And a.��ҳid = b.��ҳid And Nvl(b.��Ժ, 0) = 1;
          If n_Count <> 0 Then
            v_Err_Msg := '�������۲��˲���ʹ������Ԥ����';
            Raise Err_Item;
          End If;
        End If;
      End If;
    End Loop;
  End If;

  --���õ���
  If Nvl(n_�Ƿ�Һ�, 0) = 0 Then
    If Nvl(n_����ģʽ, 0) = 0 Or Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) = 0 Then
      --1.���з����շѴ���
      --��ȡ��ҩ����
      v_��ҩ���� := Zl_Getclinicchargepaywins(v_Nos);
    
      Select ���˽��ʼ�¼_Id.Nextval, Sysdate Into n_����id, d_�շ�ʱ�� From Dual;
    
      n_���ʽ�� := 0;
      For c_�ɷѵ� In (Select /*+ rule */
                     a.No, Max(a.��������id) As ��������id, Max(a.���˿���id) As ���˿���id, Max(a.����id) As ����id, Sum(ʵ�ս��) As ʵ�ս��,
                     Max(a.������) As ������
                    From ������ü�¼ A, Table(f_Str2List(v_Nos)) J
                    Where a.��¼���� = 1 And a.No = j.Column_Value
                    Group By a.No) Loop
        If Nvl(c_�ɷѵ�.����id, 0) <> n_����id Then
          v_Err_Msg := '�ɷѵ���:' || c_�ɷѵ�.No || '�뵱ǰ������ݲ���,������ɷ�!';
          Raise Err_Item;
        End If;
      
        n_���ʽ�� := n_���ʽ�� + Nvl(c_�ɷѵ�.ʵ�ս��, 0);
        Zl_���˻����շ�_Insert(c_�ɷѵ�.No, n_����id, 1, v_ҽ�Ƹ��ʽ����, v_����, v_�Ա�, v_����, c_�ɷѵ�.���˿���id, c_�ɷѵ�.��������id, c_�ɷѵ�.������,
                         n_����id, d_�շ�ʱ��, v_����Ա����, v_����Ա����, v_��ҩ����, 0, d_�շ�ʱ��);
      End Loop;
    
      --����ܽ���Ƿ���ȷ
      If Nvl(n_����, 0) = 0 Then
        n_���� := Nvl(n_���ʽ��, 0) - Nvl(n_������, 0);
        If Abs(n_����) > 1.00 Then
          v_Err_Msg := '���ݽɿ�����ʵ�ʽ�������̫��!';
          Raise Err_Item;
        End If;
      End If;
      If Nvl(n_���ʽ��, 0) <> Nvl(n_������, 0) + Nvl(n_����, 0) Then
        v_Err_Msg := 'ָ���Ľɷѵ��ݵ��ܽ���,������ѡ��ɷѵ���!';
        Raise Err_Item;
      End If;
    End If;
  
    --2.ȷ��֧����ʽ 
    n_�������   := -1 * n_����id;
    n_ɾ��ԭ���� := 0;
    If Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) = 1 Then
      n_ɾ��ԭ���� := 1;
    End If;
    For c_���㷽ʽ In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As ���㿨���,
                          Extractvalue(b.Column_Value, '/JS/JSKH') As ���㿨��,
                          Extractvalue(b.Column_Value, '/JS/JSFS') As ���㷽ʽ,
                          Extractvalue(b.Column_Value, '/JS/JSJE') As ������,
                          Extractvalue(b.Column_Value, '/JS/JYLSH') As ������ˮ��,
                          Extractvalue(b.Column_Value, '/JS/JYSM') As ����˵��, Extractvalue(b.Column_Value, '/JS/ZY') As ժҪ,
                          Extractvalue(b.Column_Value, '/JS/SFCYJ') As �Ƿ��Ԥ��,
                          Extractvalue(b.Column_Value, '/JS/SFXFK') As �Ƿ����ѿ�,
                          Extractvalue(b.Column_Value, '/JS/DJH') As ���ݺ�,
                          Extract(b.Column_Value, '/JS/EXPENDLIST') As Expend
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
    
      n_���ѿ�   := Nvl(c_���㷽ʽ.�Ƿ����ѿ�, 0);
      v_���㷽ʽ := c_���㷽ʽ.���㷽ʽ;
      --1.����������
      If c_���㷽ʽ.���㿨��� Is Not Null And n_���ѿ� = 0 Then
        n_�����id := Get_Cardtypeid(c_���㷽ʽ.���㿨���, 0, v_���㷽ʽ);
        Select Max(��������id)
        Into n_��������id
        From ����Ԥ����¼
        Where ����id = n_����id And �����id = n_�����id And Rownum < 2;
        If Nvl(n_��������id, 0) = 0 Then
          Select ����Ԥ����¼_Id.Nextval Into n_��������id From Dual;
        End If;
      
        Thirdcard_Balance(n_����id, n_����id, v_���㷽ʽ, n_�����id, c_���㷽ʽ.���㿨��, c_���㷽ʽ.������, c_���㷽ʽ.������ˮ��, c_���㷽ʽ.����˵��,
                          c_���㷽ʽ.���ݺ�, n_��������id, c_���㷽ʽ.Expend, n_����ģʽ, n_��������, n_ɾ��ԭ����);
      
        n_ɾ��ԭ���� := 0;
        If Nvl(n_����ģʽ, 0) = 0 Or Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) = 1 Then
          If c_���㷽ʽ.���㿨��� Is Null Then
            v_����� := v_���㷽ʽ;
          Else
            v_����� := Get_Cardname(c_���㷽ʽ.���㿨���, n_���ѿ�);
          End If;
          Update �������׼�¼
          Set ҵ�����id = n_�������
          Where ��ˮ�� = c_���㷽ʽ.������ˮ�� And ��� = v_����� And ҵ������ = n_ҵ������;
        End If;
      
        --��ɽ���ʱ�Ŵ���������������
      Elsif Nvl(n_����ģʽ, 0) = 0 Or Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) = 1 Then
        --2.���ѿ�����                           
        If c_���㷽ʽ.���㿨��� Is Not Null And n_���ѿ� = 1 Then
          n_�����id := Get_Cardtypeid(c_���㷽ʽ.���㿨���, 1, v_���㷽ʽ);
          Squarecard_Balance(n_����id, n_����id, n_�����id, c_���㷽ʽ.���㿨��, c_���㷽ʽ.������, c_���㷽ʽ.������ˮ��, c_���㷽ʽ.����˵��,
                             c_���㷽ʽ.Expend);
        
          --3.��Ԥ����
        Elsif Nvl(c_���㷽ʽ.�Ƿ��Ԥ��, 0) = 1 Then
          Zl_�����շѽ���_Modify(0, n_����id, n_����id, Null, c_���㷽ʽ.������, 0, Null, Null, Null, Null, 0, 0, 0, 0);
        
          --4.��ͨ����
        Else
          If v_���㷽ʽ Is Null Then
            v_Err_Msg := 'δָ��֧����ʽ��������ɿ�!';
            Raise Err_Item;
          End If;
        
          --���㷽ʽ|������|�������|����ժҪ||..
          v_�շѽ��� := v_���㷽ʽ || '|' || c_���㷽ʽ.������ || '| | ';
          v_��ͨ���� := v_��ͨ���� || '||' || v_�շѽ���;
        End If;
      
        If c_���㷽ʽ.���㿨��� Is Null Then
          v_����� := v_���㷽ʽ;
        Else
          v_����� := Get_Cardname(c_���㷽ʽ.���㿨���, n_���ѿ�);
        End If;
        Update �������׼�¼
        Set ҵ�����id = n_�������
        Where ��ˮ�� = c_���㷽ʽ.������ˮ�� And ��� = v_����� And ҵ������ = n_ҵ������;
      End If;
    End Loop;
  
    If Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) = 0 Then
      v_Temp := '<CZSJ>' || To_Char(d_�շ�ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
      v_Temp := '<JZID>' || n_����id || '</JZID>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
      v_Temp := '<KPBZ>' || 0 || '</KPBZ>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    
      v_Temp := '<URL>' || '' || '</URL>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    
      v_Temp := '<NETURL>' || '' || '</NETURL>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    
      v_Temp := '<FPTT>' || v_���� || '</FPTT>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    
      v_Temp := '<FPH>' || '' || '</FPH>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
      v_Temp := '<FPJE>' || '' || '</FPJE>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
      v_Temp := '<KPRQ>' || '' || '</KPRQ>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    
      Xml_Out := x_Templet;
      Return;
    End If;
  
    --5.��ͨ���㼰��ɽ���
    If v_��ͨ���� Is Not Null Then
      v_��ͨ���� := Substr(v_��ͨ����, 3);
    End If;
  
    --6.����Ʊ�ݴ���
    n_�Ƿ����Ʊ�� := b_Einvoice_Request.Einvoice_Start(1, Null);
    Zl_�����շѽ���_Modify(0, n_����id, n_����id, v_��ͨ����, Null, 0, Null, Null, Null, Null, 0, 0, n_����, 1, Null, Null, 1, Null, 0,
                     0, 0, n_�Ƿ����Ʊ��);
    If Nvl(n_�Ƿ����Ʊ��, 0) = 1 Then
      If b_Einvoice_Request.Einvoice_Create(1, n_����id, Null, v_Err_Msg) = 0 Then
        --����Ʊ�ݿ��߳ɹ�
        Raise Err_Item;
      End If;
    
      Select Max(1), Max(����), Max(����), Max(To_Char(����ʱ��, 'yyyy-mm-dd')), Max(Url����), Max(Url����), Max(Ʊ�ݽ��)
      Into n_��Ʊ��־, v_��������, v_��Ʊ���, v_��Ʊ����, v_Url, v_Url����, n_��Ʊ���
      From ����Ʊ��ʹ�ü�¼
      Where ����id = n_����id And Ʊ�� = 1 And ��¼״̬ = 1;
    
      If v_�������� Is Not Null Then
        v_���� := v_��������;
      End If;
    
    End If;
  
    v_Temp := '<CZSJ>' || To_Char(d_�շ�ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<JZID>' || n_����id || '</JZID>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<KPBZ>' || Nvl(n_��Ʊ��־, 0) || '</KPBZ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<URL>' || Nvl(v_Url, '') || '</URL>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<NETURL>' || Nvl(v_Url����, '') || '</NETURL>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<FPTT>' || v_���� || '</FPTT>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<FPH>' || v_��Ʊ��� || '</FPH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<FPJE>' || Nvl(n_��Ʊ���, 0) || '</FPJE>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<KPRQ>' || v_��Ʊ���� || '</KPRQ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    Xml_Out := x_Templet;
    Return;
  End If;

  --=================================================================================
  --�Һŵ��� 
  n_���ʽ�� := 0;
  v_Para     := zl_GetSysParameter(256);
  n_�Һ�ģʽ := Substr(v_Para, 1, 1);
  Begin
    d_����ʱ�� := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
  Exception
    When Others Then
      d_����ʱ�� := Null;
  End;

  --ʵ���Ƽ��
  n_ʵ���� := To_Number(Nvl(zl_GetSysParameter(319), '0'));
  If n_ʵ���� = 1 Then
    Select Count(1) Into n_��֤ From ����ʵ����Ϣ Where ����id = n_����id And Rownum < 2;
    If n_��֤ = 0 Then
      v_Err_Msg := '����δʵ����֤�����ܹҺš�';
      Raise Err_Item;
    End If;
  End If;

  Begin
    Select a.ִ�в���id, a.�շ�ϸĿid, c.Id, a.ִ����, b.�ű�, b.�����, b.����ʱ��, a.�ѱ�, b.����, b.�����¼id
    Into n_����id, n_��Ŀid, n_ҽ��id, v_ҽ������, v_����, n_�����, d_����ʱ��, v_�ѱ�, n_����, n_�����¼id
    From ������ü�¼ A, ���˹Һż�¼ B, ��Ա�� C
    Where a.No = v_Nos And a.��¼���� = 4 And a.��� = 1 And a.No = b.No And a.ִ���� = c.����(+);
  Exception
    When Others Then
      v_Err_Msg := 'û���ҵ�ָ���ĵ������ݣ�';
      Raise Err_Item;
  End;

  --ԤԼ����
  If n_�Һ�ģʽ = 1 Then
    If d_����ʱ�� > d_����ʱ�� And n_�����¼id Is Null Then
      n_�Һ�ģʽ := 0;
    End If;
  End If;

  If Nvl(n_����ģʽ, 0) = 0 Or Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) = 0 Then
    v_Pricegrade := Zl_Get_Pricegrade(v_վ��);
    v_��ͨ�ȼ�   := Substr(v_Pricegrade, 1, Instr(v_Pricegrade, '|') - 1);
  
    Select ���˽��ʼ�¼_Id.Nextval, Sysdate Into n_����id, d_�շ�ʱ�� From Dual;
  
    For c_���� In (Select 1 As ˳���, b.No, b.�վݷ�Ŀ, b.����id, b.ִ�в���id, b.���˿���id, b.������, b.�շ����, b.������Ŀid, b.���ӱ�־,
                        To_Char(b.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��, b.�۸񸸺�, b.��������, b.���, b.�շ�ϸĿid, b.���㵥λ,
                        Max(m.����) As ����, Max(m.���) As ���, Sum(b.��׼����) As ����, Avg(Nvl(b.����, 1) * b.����) As ����,
                        Sum(b.Ӧ�ս��) As Ӧ�ս��, Sum(b.ʵ�ս��) As ʵ�ս��, Max(j.����) As ��������, Max(q.����) As ִ�п���
                 From ������ü�¼ B, �շ���ĿĿ¼ M, ���ű� J, ���ű� Q
                 Where b.No = v_Nos And b.��¼���� = 4 And Nvl(b.����״̬, 0) = 0 And b.�շ�ϸĿid = m.Id And b.��������id = j.Id(+) And
                       b.ִ�в���id = q.Id(+)
                 Group By b.No, b.�վݷ�Ŀ, b.����id, b.ִ�в���id, b.���˿���id, b.������, b.������Ŀid, b.�շ����, b.�Ǽ�ʱ��, b.�۸񸸺�, b.��������, b.���,
                          b.�շ�ϸĿid, b.���㵥λ, b.���ӱ�־
                 Order By ���) Loop
    
      Zl_����ԤԼ�Һż�¼_Update(c_����.No, c_����.���, c_����.�۸񸸺�, c_����.��������, c_����.�շ����, c_����.�շ�ϸĿid, c_����.����, c_����.����, c_����.������Ŀid,
                         c_����.�վݷ�Ŀ, c_����.Ӧ�ս��, c_����.ʵ�ս��, c_����.���ӱ�־, Null, Null, Null, Null, c_����.���˿���id, c_����.ִ�в���id);
    
      n_���ʽ��   := n_���ʽ�� + c_����.ʵ�ս��;
      n_���       := c_����.���;
      n_���˿���id := c_����.���˿���id;
      n_ִ�в���id := c_����.ִ�в���id;
      v_No         := c_����.No;
    End Loop;
  
    Begin
      Select Zl_Fun_Customregexpenses(n_����id, 0, v_����, v_����, v_�Ա�, v_����, v_���֤��, v_�ѱ�, v_���ʽ)
      Into v_������Ŀid
      From Dual;
    Exception
      When Others Then
        v_������Ŀid := Null;
    End;
    If v_������Ŀid Is Not Null Then
      If Instr(v_������Ŀid, '|') > 0 Then
        v_��������   := v_������Ŀid || ','; --�Կո�ֿ���|��β,û�н�������
        v_������Ŀid := '';
        n_Cursor     := 0;
        While v_�������� Is Not Null Loop
          v_����ֵ         := Substr(v_��������, 1, Instr(v_��������, ',') - 1);
          Price_Rec.��Ŀid := To_Number(Substr(v_����ֵ, 1, Instr(v_����ֵ, '|') - 1));
          v_����ֵ         := Substr(v_����ֵ, Instr(v_����ֵ, '|') + 1);
          Price_Rec.����   := To_Number(Substr(v_����ֵ, 1, Instr(v_����ֵ, '|') - 1));
          v_����ֵ         := Substr(v_����ֵ, Instr(v_����ֵ, '|') + 1);
          Price_Rec.����   := To_Number(Substr(v_����ֵ, 1, Instr(v_����ֵ, '|') - 1));
          v_����ֵ         := Substr(v_����ֵ, Instr(v_����ֵ, '|') + 1);
          Price_Rec.Ӧ��   := To_Number(Substr(v_����ֵ, 1, Instr(v_����ֵ, '|') - 1));
          v_����ֵ         := Substr(v_����ֵ, Instr(v_����ֵ, '|') + 1);
          Price_Rec.ʵ��   := To_Number(v_����ֵ);
        
          n_Cursor := n_Cursor + 1;
          Price_Rec_Array(n_Cursor) := Price_Rec;
          v_�������� := Substr(v_��������, Instr(v_��������, ',') + 1);
          v_������Ŀid := v_������Ŀid || ',' || Price_Rec_Array(n_Cursor).��Ŀid;
        End Loop;
      
        If v_������Ŀid Is Not Null Then
          v_������Ŀid := Substr(v_������Ŀid, 2);
        End If;
      
        For c_������Ŀ In (Select /*+cardinality(D,10)*/
                        5 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, c.Id As ������Ŀid,
                        c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, -1 As ִ�п�������
                       From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, Table(f_Str2List(v_������Ŀid)) D
                       Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = d.Column_Value And Sysdate Between b.ִ������ And
                             Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And
                             (b.�۸�ȼ� = v_��ͨ�ȼ� Or
                             (b.�۸�ȼ� Is Null And Not Exists
                              (Select 1
                                From �շѼ�Ŀ
                                Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = v_��ͨ�ȼ� And Sysdate Between ִ������ And
                                      Nvl(��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')))))) Loop
        
          n_��� := n_��� + 1;
          For n_Cursor In 1 .. Price_Rec_Array.Count Loop
            If c_������Ŀ.��Ŀid = Price_Rec_Array(n_Cursor).��Ŀid Then
              Zl_����ԤԼ�Һż�¼_Update(v_No, n_���, Null, Null, c_������Ŀ.���, c_������Ŀ.��Ŀid, Price_Rec_Array(n_Cursor).����,
                                 Price_Rec_Array(n_Cursor).����, c_������Ŀ.������Ŀid, c_������Ŀ.�վݷ�Ŀ, Price_Rec_Array(n_Cursor).Ӧ��,
                                 Price_Rec_Array(n_Cursor).ʵ��, Null, Null, Null, Null, Null, n_���˿���id, n_ִ�в���id);
            
              n_ʵ�ս�� := Price_Rec_Array(n_Cursor).ʵ��;
              n_���ʽ�� := n_���ʽ�� + n_ʵ�ս��;
              Exit;
            End If;
          End Loop;
        End Loop;
      Else
        For c_������Ŀ In (Select /*+cardinality(D,10)*/
                        5 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, 1 As ����, c.Id As ������Ŀid,
                        c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, -1 As ִ�п�������
                       From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, Table(f_Str2List(v_������Ŀid)) D
                       Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = d.Column_Value And Sysdate Between b.ִ������ And
                             Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And
                             (b.�۸�ȼ� = v_��ͨ�ȼ� Or
                             (b.�۸�ȼ� Is Null And Not Exists
                              (Select 1
                                From �շѼ�Ŀ
                                Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = v_��ͨ�ȼ� And Sysdate Between ִ������ And
                                      Nvl(��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')))))
                       Union All
                       Select /*+cardinality(E,10)*/
                        6 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, d.�������� As ����,
                        c.Id As ������Ŀid, c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, -1 As ִ�п�������
                       From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, �շѴ�����Ŀ D, Table(f_Str2List(v_������Ŀid)) E
                       Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = d.����id And d.����id = e.Column_Value And
                             Sysdate Between b.ִ������ And Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And
                             (b.�۸�ȼ� = v_��ͨ�ȼ� Or
                             (b.�۸�ȼ� Is Null And Not Exists
                              (Select 1
                                From �շѼ�Ŀ
                                Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = v_��ͨ�ȼ� And Sysdate Between ִ������ And
                                      Nvl(��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')))))) Loop
          n_��� := n_��� + 1;
          If c_������Ŀ.���� = 5 Then
            n_�������� := n_���;
          End If;
        
          v_ʵ��     := Zl_Actualmoney(v_�ѱ�, c_������Ŀ.��Ŀid, c_������Ŀ.������Ŀid, c_������Ŀ.���� * c_������Ŀ.����);
          n_ʵ�ս�� := To_Number(Substr(v_ʵ��, Instr(v_ʵ��, ':') + 1));
        
          If c_������Ŀ.���� = 5 Then
            Zl_����ԤԼ�Һż�¼_Update(v_No, n_���, Null, Null, c_������Ŀ.���, c_������Ŀ.��Ŀid, c_������Ŀ.����, c_������Ŀ.����, c_������Ŀ.������Ŀid,
                               c_������Ŀ.�վݷ�Ŀ, c_������Ŀ.���� * c_������Ŀ.����, n_ʵ�ս��, Null, Null, Null, Null, Null, n_���˿���id,
                               n_ִ�в���id);
          Else
            Zl_����ԤԼ�Һż�¼_Update(v_No, n_���, Null, n_��������, c_������Ŀ.���, c_������Ŀ.��Ŀid, c_������Ŀ.����, c_������Ŀ.����, c_������Ŀ.������Ŀid,
                               c_������Ŀ.�վݷ�Ŀ, c_������Ŀ.���� * c_������Ŀ.����, n_ʵ�ս��, Null, Null, Null, Null, Null, n_���˿���id,
                               n_ִ�в���id);
          End If;
          n_���ʽ�� := n_���ʽ�� + n_ʵ�ս��;
        End Loop;
      End If;
    End If;
  
    --����ܽ���Ƿ���ȷ
    If Nvl(n_����, 0) = 0 Then
      n_���� := Nvl(n_���ʽ��, 0) - Nvl(n_������, 0);
      If Abs(n_����) > 1.00 Then
        v_Err_Msg := '���ݽɿ�����ʵ�ʽ�������̫��!';
        Raise Err_Item;
      End If;
    End If;
    If Nvl(n_���ʽ��, 0) <> Nvl(n_������, 0) + Nvl(n_����, 0) Then
      Select Max(����Ա����) Into v_����Ա From ������ü�¼ Where ��¼���� = 4 And NO = v_Nos;
      If v_����Ա = v_����Ա���� Then
        v_Err_Msg := 'ָ���Ľɷѵ��ݵ��ܽ���,������ѡ��ɷѵ���!';
        Raise Err_Special;
      Else
        v_Err_Msg := 'ָ���Ľɷѵ��ݵ��ܽ���,������ѡ��ɷѵ���!';
        Raise Err_Item;
      End If;
    End If;
  
    n_��������id := 0;
    For c_���㷽ʽ In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As ���㿨���,
                          Extractvalue(b.Column_Value, '/JS/JSKH') As ���㿨��,
                          Extractvalue(b.Column_Value, '/JS/JSFS') As ���㷽ʽ,
                          Extractvalue(b.Column_Value, '/JS/JSJE') As ������,
                          Extractvalue(b.Column_Value, '/JS/JYLSH') As ������ˮ��,
                          Extractvalue(b.Column_Value, '/JS/JYSM') As ����˵��, Extractvalue(b.Column_Value, '/JS/ZY') As ժҪ,
                          Extractvalue(b.Column_Value, '/JS/SFCYJ') As �Ƿ��Ԥ��,
                          Extractvalue(b.Column_Value, '/JS/SFXFK') As �Ƿ����ѿ�
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
    
      n_���ѿ�       := Nvl(c_���㷽ʽ.�Ƿ����ѿ�, 0);
      v_��ʱ���㷽ʽ := c_���㷽ʽ.���㷽ʽ;
    
      If Nvl(c_���㷽ʽ.�Ƿ��Ԥ��, 0) = 1 Then
        n_Ԥ��֧�� := c_���㷽ʽ.������;
      Else
        n_��֧ͨ�� := Nvl(n_��֧ͨ��, 0) + c_���㷽ʽ.������;
      
        If c_���㷽ʽ.���㿨��� Is Not Null Then
          --���������㷽ʽ
          If n_���ѿ� = 0 Then
            n_�����id := Get_Cardtypeid(c_���㷽ʽ.���㿨���, 0, v_��ʱ���㷽ʽ);
          Else
            n_���㿨��� := Get_Cardtypeid(c_���㷽ʽ.���㿨���, 1, v_��ʱ���㷽ʽ);
          End If;
          v_���㿨��   := c_���㷽ʽ.���㿨��;
          v_������ˮ�� := c_���㷽ʽ.������ˮ��;
          v_����˵��   := c_���㷽ʽ.����˵��;
          v_ժҪ       := c_���㷽ʽ.ժҪ;
        
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          If Nvl(n_��������id, 0) = 0 Then
            n_��������id := n_Ԥ��id;
          End If;
          v_���㷽ʽ := v_���㷽ʽ || '|' || v_��ʱ���㷽ʽ || ',' || c_���㷽ʽ.������ || ',,1' || ',' || n_Ԥ��id || ',' || n_��������id;
        Else
          Select Nvl(Max(1), 0) Into n_Exists From ���㷽ʽ Where ���� = Nvl(v_��ʱ���㷽ʽ, '-') And ���� In (3, 4);
        
          If n_Exists = 1 Then
            n_ҽ��֧�� := c_���㷽ʽ.������;
          Else
            --�������㷽ʽ
            v_���㷽ʽ := v_���㷽ʽ || '|' || v_��ʱ���㷽ʽ || ',' || c_���㷽ʽ.������ || ',,0';
          End If;
        End If;
      End If;
    End Loop;
  
    If v_���㷽ʽ Is Not Null Then
      v_���㷽ʽ := Substr(v_���㷽ʽ, 2);
    End If;
  
    If n_�Һ�ģʽ = 0 Then
      Zl_ԤԼ�ҺŽ���_Insert(v_Nos, Null, Null, n_����id, Zl_Get_��������(v_����), n_����id, n_�����, v_����, v_�Ա�, v_����, v_ҽ�Ƹ��ʽ����, v_�ѱ�,
                       v_���㷽ʽ, n_��֧ͨ��, n_Ԥ��֧��, n_ҽ��֧��, d_����ʱ��, v_����Ա����, v_����Ա����, d_�շ�ʱ��, n_�����id, n_���㿨���, v_���㿨��,
                       v_������ˮ��, v_����˵��, Null, 0, 0, Null, 1);
    Else
      Zl_ԤԼ�ҺŽ���_����_Insert(v_Nos, Null, Null, n_����id, Zl_Get_��������(v_����, n_�����¼id), n_����id, n_�����, v_����, v_�Ա�, v_����,
                          v_ҽ�Ƹ��ʽ����, v_�ѱ�, v_���㷽ʽ, n_��֧ͨ��, n_Ԥ��֧��, Null, d_����ʱ��, v_����Ա����, v_����Ա����, d_�շ�ʱ��, n_�����id,
                          n_���㿨���, v_���㿨��, v_������ˮ��, v_����˵��, Null, 0, 0, Null, 1);
    End If;
  
  End If;

  If Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) = 0 Then
    v_Temp := '<CZSJ>' || To_Char(d_�շ�ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<JZID>' || n_����id || '</JZID>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    Xml_Out := x_Templet;
    Return;
  End If;

  Zl_ԤԼ�ҺŽ���_��Ÿ���(v_Nos, Null, v_����Ա����, v_����Ա����, d_����ʱ��, d_�շ�ʱ��);
  n_�������� := 0;
  For c_���㷽ʽ In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As ���㿨���,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As ���㿨��,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As ���㷽ʽ,
                        Extractvalue(b.Column_Value, '/JS/JSJE') As ������,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As ������ˮ��,
                        Extractvalue(b.Column_Value, '/JS/JYSM') As ����˵��, Extractvalue(b.Column_Value, '/JS/ZY') As ժҪ,
                        Extractvalue(b.Column_Value, '/JS/SFCYJ') As �Ƿ��Ԥ��,
                        Extractvalue(b.Column_Value, '/JS/SFXFK') As �Ƿ����ѿ�
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
  
    v_���㷽ʽ := c_���㷽ʽ.���㷽ʽ;
    If c_���㷽ʽ.���㿨��� Is Null Then
      v_����� := v_���㷽ʽ;
    Else
      v_����� := Get_Cardname(c_���㷽ʽ.���㿨���, c_���㷽ʽ.�Ƿ����ѿ�);
      If Nvl(c_���㷽ʽ.�Ƿ����ѿ�, 0) = 0 Then
        n_�����id := Get_Cardtypeid(c_���㷽ʽ.���㿨���, 0, v_���㷽ʽ);
      
        If Nvl(n_�����id, 0) <> 0 Then
          v_���㿨�� := c_���㷽ʽ.���㿨��;
        
          Select Max(��������id)
          Into n_��������id
          From ����Ԥ����¼
          Where ����id = n_����id And �����id = n_�����id And Rownum < 2;
          If Nvl(n_��������id, 0) = 0 Then
            Select ����Ԥ����¼_Id.Nextval Into n_��������id From Dual;
          End If;
        
          Zl_���˹Һ��շ�_Modify(v_Nos, n_����id, v_���㷽ʽ || ',' || c_���㷽ʽ.������ || ',,' || c_���㷽ʽ.ժҪ, 1, 0, 0, 1, Null, n_��������,
                           n_��������id, n_�����id, v_���㿨��, c_���㷽ʽ.������ˮ��, c_���㷽ʽ.����˵��);
        
          For c_��չ��Ϣ In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As ��������,
                                Extractvalue(b.Column_Value, '/EXPEND/JYLR') As ��������
                         From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS/EXPENDLIST/EXPEND'))) B) Loop
            Zl_�������㽻��_Insert(n_�����id, 0, v_���㿨��, n_����id, c_��չ��Ϣ.�������� || '|' || c_��չ��Ϣ.��������, 0);
          End Loop;
        End If;
      End If;
    End If;
    n_�������� := 1;
  
    Update �������׼�¼
    Set ҵ�����id = n_����id
    Where ��ˮ�� = c_���㷽ʽ.������ˮ�� And ��� = v_����� And ҵ������ = n_ҵ������;
  End Loop;

  --6.����Ʊ�ݴ���
  n_�Ƿ����Ʊ�� := b_Einvoice_Request.Einvoice_Start(4, Null);
  Zl_���˹Һ��շ�_Modify(v_Nos, n_����id, Null, 0, 1, 0, 1, Null, 0, Null, Null, Null, Null, Null, 0, 2, n_�Ƿ����Ʊ��);

  --�������
  Zl_���˹ҺŻ���_Update(v_ҽ������, n_ҽ��id, n_��Ŀid, n_����id, d_����ʱ��, 2, v_����, 1, n_�����¼id);
  If Nvl(n_�Ƿ����Ʊ��, 0) = 1 Then
    If b_Einvoice_Request.Einvoice_Create(4, n_����id, Null, v_Err_Msg) = 0 Then
      --����Ʊ�ݿ��߳ɹ�
      Raise Err_Item;
    End If;
  
    Select Max(1), Max(����), Max(����), Max(To_Char(����ʱ��, 'yyyy-mm-dd')), Max(Url����), Max(Url����), Max(Ʊ�ݽ��)
    Into n_��Ʊ��־, v_��������, v_��Ʊ���, v_��Ʊ����, v_Url, v_Url����, n_��Ʊ���
    From ����Ʊ��ʹ�ü�¼
    Where ����id = n_����id And Ʊ�� = 4 And ��¼״̬ = 1;
  
    If v_�������� Is Not Null Then
      v_���� := v_��������;
    End If;
  
  End If;
  v_Temp := '<CZSJ>' || To_Char(d_�շ�ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<JZID>' || n_����id || '</JZID>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;

  v_Temp := '<KPBZ>' || Nvl(n_��Ʊ��־, 0) || '</KPBZ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;

  v_Temp := '<URL>' || Nvl(v_Url, '') || '</URL>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;

  v_Temp := '<NETURL>' || Nvl(v_Url����, '') || '</NETURL>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;

  v_Temp := '<FPTT>' || v_���� || '</FPTT>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;

  v_Temp := '<FPH>' || v_��Ʊ��� || '</FPH>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<FPJE>' || Nvl(n_��Ʊ���, 0) || '</FPJE>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<KPRQ>' || v_��Ʊ���� || '</KPRQ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;

  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Err_Special Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20105, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Payment;
/


Create Or Replace Procedure Zl_Third_Charge_Delcheck
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  -------------------------------------------------------------------------------------------------- 
  --����:�����˷Ѽ�� 
  --���:Xml_In: 
  --<IN> 
  --    <BRID>����ID</BRID> 
  --    <XM>����</XM> 
  --    <SFZH>���֤��</SFZH> 
  --    <JE></JE> //�˿��ܽ�� 
  --    <JSKLB></JSKLB>     //���㿨��� 
  --    <TFZY>�˷�ժҪ</TFZY> 
  --    <JCFP>1</JCFP>      //��鷢Ʊ,0-�����;1-���;Ϊ1ʱ����ӡ�˷�Ʊ�ĵ��ݲ����˷� 
  --    <XL>����</XL>         //ҽ����������,�մ�����ͨ���� 
  --    <FYLIST> 
  --        <FY> 
  --           <DJH>�˿�ݺ�</DJH> 
  --           <XH>�˿����(��ʽ:1,2,3..Ϊ�մ�����ʣ������)</DJH> 
  --        <FY> 
  --    </FYLIST> 
  --    <TKLIST> 
  --        <TK> 
  --            <TKKLB>�˿���</TKKLB> 
  --            <TKKH>�˿��</TKKH> 
  --            <TKFS>�˿ʽ</TKFS> //�˿ʽ:�ֽ�;֧Ʊ,�����������,���Դ��� 
  --            <TKJE>֧�����</TKJE> 
  --            <JYLSH>������ˮ��</JYLSH> 
  --            <JYSM>����˵��</JYSM> 
  --            <TKZY>ժҪ</TKZY> 
  --            <TYJK>�˻�Ԥ����</TYJK> //�ʳ�Ԥ��ʱ,ֻ��JSJE�ڵ�:1-��Ԥ�� 
  --            <SFXFK>�Ƿ����ѿ�</SFXFK>   //(1-�����ѿ�),���ѿ�ʱ,������㿨���,���㿨��,������Ƚӵ� 
  --            <DJH>S0000001</DJH> //�ֵ��ݽ���ʱ���� 
  --            <EXPENDLIST>  //��չ������Ϣ 
  --                <EXPEND> 
  --                    <JYMC>��������</JYMC> 
  --                    <JYLR>��������</JYLR> 
  --                </EXPEND> 
  --            </EXPENDLIST> 
  --        </TK> 
  --    </TKLIST> 
  --</IN> 

  --����:Xml_Out 
  --  <OUT> 
  --    �D�D�������д�������˵��ͨ����� 
  --    <ERROR> 
  --      <MSG>������Ϣ</MSG> 
  --    </ERROR> 
  --  </OUT> 
  -------------------------------------------------------------------------------------------------- 
  n_�˿��ܶ� ������ü�¼.ʵ�ս��%Type;

  n_����id     ������ü�¼.����id%Type;
  v_����       ������Ϣ.����%Type;
  v_���֤��   ������Ϣ.���֤��%Type;
  n_���ݲ���id ������ü�¼.����id%Type;
  n_����id     ������ü�¼.����id%Type;
  n_ԭ������� ����Ԥ����¼.�������%Type;
  v_���㿨��� Varchar2(100);
  v_���㷽ʽ   ҽ�ƿ����.���㷽ʽ%Type;
  n_�����id   ҽ�ƿ����.Id%Type;
  v_������     ҽ�ƿ����.����%Type;

  v_ժҪ     ������ü�¼.ժҪ%Type;
  n_Count    Number(18);
  n_Temp     Number(18);
  n_��鷢Ʊ Number(3);
  n_�Ƿ��ӡ Number(3);
  n_�˷�ģʽ Number(3);
  n_״̬     Number(3);
  n_����     ������Ϣ.����%Type;

  v_Temp    Varchar2(32767); --��ʱXML 
  x_Templet Xmltype; --ģ��XML 

  v_Err_Msg Varchar2(200);
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  --0.��ȡ����еĲ���ID����Ϣ 
  Select To_Number(Extractvalue(Value(A), 'IN/BRID')), To_Number(Extractvalue(Value(A), 'IN/JE')),
         Extractvalue(Value(A), 'IN/TFZY'), To_Number(Extractvalue(Value(A), 'IN/JCFP')),
         Extractvalue(Value(A), 'IN/JSKLB'), To_Number(Extractvalue(Value(A), 'IN/XL')),
         Extractvalue(Value(A), 'IN/SFZH'), Extractvalue(Value(A), 'IN/XM')
  Into n_����id, n_�˿��ܶ�, v_ժҪ, n_��鷢Ʊ, v_���㿨���, n_����, v_���֤��, v_����
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  If Nvl(n_����id, 0) = 0 And Not v_���֤�� Is Null And Not v_���� Is Null Then
    n_����id := Zl_Third_Getpatiid(v_���֤��, v_����);
  End If;
  --0.��ؼ�� 
  If Nvl(n_����id, 0) = 0 Then
    v_Err_Msg := '������Чʶ�������,�������˷Ѳ���!';
    Raise Err_Item;
  End If;

  n_�˷�ģʽ := zl_GetSysParameter('�����˷���������');

  If v_���㿨��� Is Not Null Then
    Begin
      n_�����id := To_Number(v_���㿨���);
    Exception
      When Others Then
        n_�����id := 0;
    End;
    If n_�����id = 0 Then
      Begin
        Select ID, ���� Into n_�����id, v_������ From ҽ�ƿ���� Where ���� = v_���㿨���;
      Exception
        When Others Then
          v_Err_Msg := '�޷�ȷ�ϴ���Ľ��㿨��';
          Raise Err_Item;
      End;
    Else
      Begin
        Select ���� Into v_������ From ҽ�ƿ���� Where ID = n_�����id;
      Exception
        When Others Then
          v_Err_Msg := '�޷�ȷ�ϴ���Ľ��㿨��';
          Raise Err_Item;
      End;
    End If;
  Else
    n_�����id := 0;
  End If;

  If Nvl(n_�����id, 0) <> 0 Then
    Select ���㷽ʽ Into v_���㷽ʽ From ҽ�ƿ���� Where ID = n_�����id;
  End If;

  --��Աid,��Ա���,��Ա���� 
  v_Temp := Zl_Identity(1);
  If Nvl(v_Temp, '0') = '0' Or Nvl(v_Temp, '_') = '_' Then
    v_Err_Msg := 'ϵͳ�����ϱ���Ч�Ĳ���Ա,�������˷�!';
    Raise Err_Item;
  End If;

  --1.�˷Ѽ�� 
  n_Count      := 0;
  n_ԭ������� := 0;
  For c_���� In (Select Extractvalue(b.Column_Value, '/FY/DJH') As ���ݺ�, Extractvalue(b.Column_Value, '/FY/XH') As �˿����
               From Table(Xmlsequence(Extract(Xml_In, '/IN/FYLIST/FY'))) B) Loop
  
    If c_����.���ݺ� Is Null Then
      v_Err_Msg := 'δȷ��ָ���˷ѵĵ��ݺ�,�����˷�!';
      Raise Err_Item;
    End If;
  
    If Nvl(n_�˷�ģʽ, 0) = 1 Then
      Begin
        Select Nvl(״̬, 0) Into n_״̬ From �����˷����� Where NO = c_����.���ݺ� And Mod(��¼����, 10) = 1;
      Exception
        When Others Then
          n_״̬ := 0;
      End;
      If n_״̬ <> 1 Then
        v_Err_Msg := '��ǰΪ�˷�����ģʽ,�˷�֮ǰ�����벢���ͨ���õ���!';
        Raise Err_Item;
      End If;
    End If;
  
    Begin
      Select a.�������, a.����id, a.����id
      Into n_Temp, n_����id, n_���ݲ���id
      From ����Ԥ����¼ A, ������ü�¼ B
      Where a.����id = b.����id And b.No = c_����.���ݺ� And b.��¼���� = 1 And Nvl(b.����״̬, 0) = 0 And b.��¼״̬ In (1, 3) And
            Rownum < 2;
    Exception
      When Others Then
        n_Temp := Null;
    End;
  
    If n_Temp Is Null Then
      v_Err_Msg := 'ָ���ĵ��ݺ�:' || c_����.���ݺ� || 'δ�ҵ�,�����˷�!';
      Raise Err_Item;
    End If;
  
    If Nvl(n_���ݲ���id, 0) = 0 Then
      Begin
        Select ����id
        Into n_���ݲ���id
        From ������ü�¼
        Where NO = c_����.���ݺ� And ��¼���� = 1 And Nvl(����״̬, 0) = 0 And ��¼״̬ In (1, 3) And Rownum < 2;
      Exception
        When Others Then
          n_���ݲ���id := 0;
      End;
    End If;
  
    If Nvl(n_����id, 0) <> Nvl(n_���ݲ���id, 0) Then
      v_Err_Msg := '�����˷ѵ��շѵ�:' || c_����.���ݺ� || '���ǵ�ǰ���˵��շѵ�,�����˷�!';
      Raise Err_Item;
    End If;
  
    If n_ԭ������� <> 0 And n_ԭ������� <> n_Temp Then
      v_Err_Msg := '�����˷ѵĵ��ݺŲ���һ���շѽ���,�����˷�!';
      Raise Err_Item;
    End If;
    n_ԭ������� := n_Temp;
  
    Select Count(1) Into n_Temp From ���ò����¼ Where �շѽ���id = n_����id;
    If Nvl(n_Temp, 0) <> 0 Then
      v_Err_Msg := '�����˷ѵĵ��ݺ��Ѿ������˱��ղ������,�����˷�!';
      Raise Err_Item;
    End If;
  
    If Nvl(n_�����id, 0) <> 0 Then
      If Nvl(n_����, 0) = 0 Then
        Select Count(1) Into n_Temp From ����Ԥ����¼ Where ����id = n_����id And �����id <> n_�����id;
      Else
        Select Count(1)
        Into n_Temp
        From ����Ԥ����¼ A, ���㷽ʽ B
        Where a.����id = n_����id And �����id <> n_�����id And a.���㷽ʽ = b.���� And b.���� Not In (3, 4);
        If n_Temp = 0 Then
          Select Nvl(Max(1), 0)
          Into n_Temp
          From ���ս����¼ A
          Where a.��¼id = n_����id And ���� <> n_���� And Rownum < 2;
        End If;
      End If;
      If Nvl(n_Temp, 0) > 0 Then
        v_Err_Msg := '�����˷ѵĵ��ݰ���' || v_���㷽ʽ || '����Ľ��㷽ʽ,�����˷�!';
        Raise Err_Item;
      End If;
    End If;
  
    If Nvl(n_��鷢Ʊ, 0) = 1 Then
      Select Max(Decode(a.ʵ��Ʊ��, Null, 0, 1))
      Into n_�Ƿ��ӡ
      From ������ü�¼ A
      Where NO = c_����.���ݺ� And ��¼���� = 1;
      If Nvl(n_�Ƿ��ӡ, 0) = 1 Then
        v_Err_Msg := '�����˷ѵĵ��ݺ��ѿ���Ʊ,�����˷�!';
        Raise Err_Item;
      End If;
    End If;
  
    --����Ʊ�ݼ�� 
    If b_Einvoice_Request.Einvoice_Cancel_Check(1, n_����id, v_Err_Msg) = 0 Then
      --ʧ�ܺ�ֱ���״�
      Raise Err_Item;
    End If;
  
    n_Count := n_Count + 1;
  End Loop;

  If n_Count = 0 Then
    v_Err_Msg := 'δȷ��������Ҫ�˷ѵĵ���,�����˷�!';
    Raise Err_Item;
  End If;

  --2.֧����ʽ��� 
  n_Count := 0;
  For c_���㷽ʽ In (Select Extractvalue(b.Column_Value, '/TK/TKKLB') As �����, Extractvalue(b.Column_Value, '/TK/TKKH') As ����,
                        Extractvalue(b.Column_Value, '/TK/TKFS') As ���㷽ʽ,
                        -1 * To_Number(Extractvalue(b.Column_Value, '/TK/TKJE')) As �˿���,
                        Extractvalue(b.Column_Value, '/TK/JYLSH') As ������ˮ��,
                        Extractvalue(b.Column_Value, '/TK/JYSM') As ����˵��, Extractvalue(b.Column_Value, '/TK/TKZY') As ժҪ,
                        Extractvalue(b.Column_Value, '/TK/TYJK') As �Ƿ���Ԥ��,
                        Extractvalue(b.Column_Value, '/TK/SFXFK') As �Ƿ����ѿ�,
                        Extract(b.Column_Value, '/TK/EXPENDLIST') As Expend
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/TKLIST/TK'))) B) Loop
  
    --1.�˻������� 
    If c_���㷽ʽ.����� Is Not Null And Nvl(c_���㷽ʽ.�Ƿ����ѿ�, 0) = 0 Then
      --1.���������� 
      Null;
    Elsif c_���㷽ʽ.����� Is Not Null And Nvl(c_���㷽ʽ.�Ƿ����ѿ�, 0) = 1 Then
      --2.���ѿ����� 
      Null;
    Elsif Nvl(c_���㷽ʽ.�Ƿ���Ԥ��, 0) = 1 Then
      --3.��Ԥ���� 
      Null;
    Else
      --4.��ͨ���� 
      If c_���㷽ʽ.���㷽ʽ Is Null Then
        v_Err_Msg := 'δָ��֧����ʽ,�����˷�!';
        Raise Err_Item;
      End If;
    End If;
    n_Count := n_Count + 1;
  End Loop;

  If n_Count = 0 Then
    v_Err_Msg := '������Чȷ�ϵ�ǰ��֧����ʽ,�����˷�!';
    Raise Err_Item;
  End If;

  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Charge_Delcheck;
/

Create Or Replace Procedure Zl_Third_Charge_Del
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is

  -------------------------------------------------------------------------------------------------- 
  --����:�����˷ѽ��� 
  --���:Xml_In: 
  --<IN> 
  --    <BRID>����ID</BRID> 
  --    <XM>����</XM> 
  --    <SFZH>���֤��</SFZH> 
  --    <JE></JE> //�˿��ܽ�� 
  --    <JSKLB></JSKLB>     //���㿨��� 
  --    <TFZY>�˷�ժҪ</TFZY> 
  --    <JCFP>1</JCFP>      //��鷢Ʊ 
  --    <JSMS>1</JSMS>          //����ģʽ��0-��ͨģʽ��1-�첽����ģʽ 
  --    <CZLX>0</CZLX>          //�������ͣ�����ģʽΪ1ʱ���룬0-��ʼ���㣬1-��ɽ��㣬2-���˽��� 
  --    <CXID>1</CXID>          //��������ID����������Ϊ1��2ʱ���� 
  --    <FYLIST> 
  --        <FY> 
  --           <DJH>�˿�ݺ�</DJH> 
  --           <XH>�˿����(��ʽ:1,2,3..Ϊ�մ�����ʣ������)</DJH> 
  --        <FY> 
  --    </FYLIST> 
  --    <TKLIST>          //�����б���������Ϊ2ʱ�ɲ����� 
  --        <TK> 
  --            <TKKLB>�˿���</TKKLB> 
  --            <TKKH>�˿��</TKKH> 
  --            <TKFS>�˿ʽ</TKFS> //�˿ʽ:�ֽ�;֧Ʊ,�����������,���Դ��� 
  --            <TKJE>֧�����</TKJE> 
  --            <JYLSH>������ˮ��</JYLSH> 
  --            <TKZY>ժҪ</TKZY> 
  --            <TYJK>�˻�Ԥ����</TYJK> //�ʳ�Ԥ��ʱ,ֻ��JSJE�ڵ�:1-��Ԥ�� 
  --            <SFXFK>�Ƿ����ѿ�</SFXFK>   //(1-�����ѿ�),���ѿ�ʱ,������㿨���,���㿨��,������Ƚӵ� 
  --            <DJH>S0000001</DJH> //�ֵ��ݽ���ʱ���� 
  --            <EXPENDLIST>  //��չ������Ϣ 
  --                <EXPEND> 
  --                    <JYMC>��������</JYMC> 
  --                    <JYLR>��������</JYLR> 
  --                </EXPEND> 
  --            </EXPENDLIST> 
  --        </TK> 
  --    </TKLIST> 
  --</IN> 

  --����:Xml_Out 
  --  <OUTPUT> 
  --    <CZSJ>����ʱ��</CZSJ>          //HIS�ĵǼ�ʱ�� 
  --    <YJZID>ԭ����ID</YJZID>       //ԭ����ID 
  --    <CXID>����ID</CXID>          //����ID 
  --  <KPBZ>��Ʊ��־</KPBZ> //1-�ɹ����ߵ���Ʊ��;0-δ��Ʊ�ɹ���־
  --  <URL>H5ҳ��URL</URL>
  --  <NETURL>����H5ҳ��URL</NETURL>
  --  <FPTT>��Ʊ̧ͷ</FPTT>        //��������
  --  <FPH>��Ʊ��</FPH>             //��Ʊ���
  --  <FPJE>��Ʊ���</FPJE>        //100.00
  --  <KPRQ>��Ʊ����</KPRQ>   //yyyy-mm-dd
  --    �D�D�������д�������˵����ȷִ�� 
  --    <ERROR> 
  --      <MSG>������Ϣ</MSG> 
  --    </ERROR> 
  --  </OUTPUT> 
  -------------------------------------------------------------------------------------------------- 
  n_�˿��ܶ� ������ü�¼.ʵ�ս��%Type;
  n_�����id ҽ�ƿ����.Id%Type;
  v_���㷽ʽ Varchar2(2000);
  n_����ģʽ Number(1); --0-��ͨģʽ��1-�첽����ģʽ 
  n_�������� Number(1); --����ģʽΪ1ʱ���룬0 - ��ʼ���㣬1 - ��ɽ��㣬2 - ���˽��� 

  n_����id     ������ü�¼.����id%Type;
  v_����       ������Ϣ.����%Type;
  v_���֤��   ������Ϣ.���֤��%Type;
  n_���ݲ���id ������ü�¼.����id%Type;
  v_����Ա���� ������ü�¼.����Ա���%Type;
  v_����Ա���� ������ü�¼.����Ա����%Type;
  n_����id     ������ü�¼.����id%Type;
  n_����id     ������ü�¼.����id%Type;
  n_���ʽ��   ������ü�¼.���ʽ��%Type;
  n_����     ����Ԥ����¼.��Ԥ��%Type;
  n_ԭ������� ����Ԥ����¼.�������%Type;
  l_�Һŵ�     t_StrList := t_StrList();
  n_�������   ����Ԥ����¼.�������%Type;
  n_��鷢Ʊ   Number(3);
  n_�Ƿ��ӡ   Number(3);
  v_���㿨��� Varchar2(100);
  v_����ids    Varchar2(1000);

  n_���ѿ�     Number;
  n_���ѿ�id   ���ѿ���Ϣ.Id%Type;
  v_ժҪ       ������ü�¼.ժҪ%Type;
  n_Count      Number(18);
  n_Billcount  Number(18);
  n_��������id ����Ԥ����¼.��������id%Type;
  n_ɾ��ԭ���� Number;

  d_�˷�ʱ�� ����Ԥ����¼.�տ�ʱ��%Type;
  v_�Һŵ�   ���˹Һż�¼.No%Type;
  v_�շѵ�   ������ü�¼.No%Type;

  v_�˷ѽ��� Varchar2(2000);
  v_��ͨ���� Varchar2(4000);
  n_ʣ���� ������ü�¼.���ʽ��%Type;

  n_�Ƿ����Ʊ�� ����Ԥ����¼.�Ƿ����Ʊ��%Type;
  n_��Ʊ��־     Number(2);
  v_��������     ����Ʊ��ʹ�ü�¼.����%Type;
  v_��Ʊ���     ����Ʊ��ʹ�ü�¼.����%Type;
  v_��Ʊ����     Varchar2(20);
  n_��Ʊ���     ����Ʊ��ʹ�ü�¼.Ʊ�ݽ��%Type;
  v_Url          ����Ʊ��ʹ�ü�¼.Url����%Type;
  v_Url����      ����Ʊ��ʹ�ü�¼.Url����%Type;

  v_Temp    Varchar2(32767); --��ʱXML 
  v_Err_Msg Varchar2(200);
  Err_Item Exception;

  --��ȡ��������� 
  Function Get_Cardname
  (
    �����_In Varchar2,
    ���ѿ�_In Number
  ) Return Varchar2 As
    v_����       ҽ�ƿ����.����%Type;
    n_By_Id_Find Number;
  
    v_Err_Msg Varchar2(200);
    Err_Item Exception;
  Begin
    If �����_In Is Null Then
      Return Null;
    End If;
  
    Select Decode(Translate(Nvl(�����_In, 'abcd'), '#1234567890', '#'), Null, 1, 0) Into n_By_Id_Find From Dual;
  
    --1.�������ID����ҽ�ƿ� 
    If Nvl(���ѿ�_In, 0) = 0 And Nvl(n_By_Id_Find, 0) = 1 Then
      Begin
        Select ���� Into v_���� From ҽ�ƿ���� Where ID = To_Number(�����_In);
      Exception
        When Others Then
          v_Err_Msg := 'δ�ҵ�ָ���Ľ���֧����Ϣ��';
          Raise Err_Item;
      End;
    End If;
  
    --2.�������Ʋ���ҽ�ƿ� 
    If Nvl(���ѿ�_In, 0) = 0 And Nvl(n_By_Id_Find, 0) = 0 Then
      Begin
        Select ���� Into v_���� From ҽ�ƿ���� Where ���� = �����_In;
      Exception
        When Others Then
          v_Err_Msg := �����_In || '������!';
          Raise Err_Item;
      End;
    End If;
  
    --3.�������ID�������ѿ� 
    If Nvl(���ѿ�_In, 0) = 1 And Nvl(n_By_Id_Find, 0) = 1 Then
      Begin
        Select ���� Into v_���� From ���ѿ����Ŀ¼ Where ��� = To_Number(�����_In);
      Exception
        When Others Then
          v_Err_Msg := 'δ�ҵ�ָ���Ľ���֧����Ϣ��';
          Raise Err_Item;
      End;
    End If;
  
    --3.�������Ʋ������ѿ� 
    If Nvl(���ѿ�_In, 0) = 1 And Nvl(n_By_Id_Find, 0) = 0 Then
      Begin
        Select ���� Into v_���� From ���ѿ����Ŀ¼ Where ���� = �����_In;
      Exception
        When Others Then
          v_Err_Msg := �����_In || '������!';
          Raise Err_Item;
      End;
    End If;
  
    Return v_����;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  End;

  --��ȡ�����ID 
  Function Get_Cardtypeid
  (
    �����_In    Varchar2,
    ���ѿ�_In    Number,
    ���㷽ʽ_Out In Out ҽ�ƿ����.���㷽ʽ%Type
  ) Return Number As
    n_�����id ҽ�ƿ����.Id%Type;
    v_����     ҽ�ƿ����.����%Type;
    n_����     ҽ�ƿ����.�Ƿ�����%Type;
    v_���㷽ʽ ҽ�ƿ����.���㷽ʽ%Type;
  
    n_By_Id_Find Number;
    v_Err_Msg    Varchar2(200);
    Err_Item Exception;
  Begin
    If �����_In Is Null Then
      Return 0;
    End If;
  
    Select Decode(Translate(Nvl(�����_In, 'abcd'), '#1234567890', '#'), Null, 1, 0) Into n_By_Id_Find From Dual;
  
    --1.�������ID����ҽ�ƿ� 
    If Nvl(���ѿ�_In, 0) = 0 And Nvl(n_By_Id_Find, 0) = 1 Then
      Begin
        Select ID, ���㷽ʽ, ����, �Ƿ�����
        Into n_�����id, v_����, v_���㷽ʽ, n_����
        From ҽ�ƿ����
        Where ID = To_Number(�����_In);
      Exception
        When Others Then
          v_Err_Msg := 'δ�ҵ�ָ���Ľ���֧����Ϣ��';
          Raise Err_Item;
      End;
    End If;
  
    --2.�������Ʋ���ҽ�ƿ� 
    If Nvl(���ѿ�_In, 0) = 0 And Nvl(n_By_Id_Find, 0) = 0 Then
      Begin
        Select ID, ���㷽ʽ, ����, �Ƿ�����
        Into n_�����id, v_����, v_���㷽ʽ, n_����
        From ҽ�ƿ����
        Where ���� = �����_In;
      Exception
        When Others Then
          v_Err_Msg := �����_In || '������!';
          Raise Err_Item;
      End;
    End If;
  
    --3.�������ID�������ѿ� 
    If Nvl(���ѿ�_In, 0) = 1 And Nvl(n_By_Id_Find, 0) = 1 Then
      Begin
        Select ���, ���㷽ʽ, ����, ����
        Into n_�����id, v_����, v_���㷽ʽ, n_����
        From ���ѿ����Ŀ¼
        Where ��� = To_Number(�����_In);
      Exception
        When Others Then
          v_Err_Msg := 'δ�ҵ�ָ���Ľ���֧����Ϣ��';
          Raise Err_Item;
      End;
    End If;
  
    --3.�������Ʋ������ѿ� 
    If Nvl(���ѿ�_In, 0) = 1 And Nvl(n_By_Id_Find, 0) = 0 Then
      Begin
        Select ���, ���㷽ʽ, ����, ����
        Into n_�����id, v_����, v_���㷽ʽ, n_����
        From ���ѿ����Ŀ¼
        Where ���� = �����_In;
      Exception
        When Others Then
          v_Err_Msg := �����_In || '������!';
          Raise Err_Item;
      End;
    End If;
  
    If Nvl(n_����, 0) = 0 Then
      v_Err_Msg := v_���� || 'δ���ã���������нɷѣ�';
      Raise Err_Item;
    End If;
  
    If ���㷽ʽ_Out Is Null Then
      If v_���㷽ʽ Is Null Then
        v_Err_Msg := v_���� || 'δ���ý��㷽ʽ����������нɷѣ�';
        Raise Err_Item;
      End If;
    
      ���㷽ʽ_Out := v_���㷽ʽ;
    End If;
  
    Return n_�����id;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  End;

  Procedure Thirdcard_Balance
  (
    ����id_In     ����Ԥ����¼.����id%Type,
    ����id_In     ����Ԥ����¼.����id%Type,
    ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type,
    �����_In     ����Ԥ����¼.�����id%Type,
    ����_In       ����Ԥ����¼.����%Type,
    �˿���_In   ����Ԥ����¼.��Ԥ��%Type,
    ������ˮ��_In ����Ԥ����¼.������ˮ��%Type,
    ����˵��_In   ����Ԥ����¼.����˵��%Type,
    No_In         ����Ԥ����¼.No%Type,
    ��������id_In ����Ԥ����¼.��������id%Type,
    ժҪ_In       ����Ԥ����¼.ժҪ%Type,
    Xmlexpned_In  Xmltype,
    ����ģʽ_In   Number := 0,
    ��������_In   Number := 0,
    ɾ��ԭ����_In Number := 0
  ) Is
    --��Σ� 
    --         ����ģʽ_in   0-��ͨģʽ��1-�첽����ģʽ 
    --        ��������_in   ����ģʽΪ1ʱ���룬0 - ��ʼ���㣬1 - ��ɽ��㣬2 - ���˽��� 
    --        ɾ��ԭ����_in ��������_InΪ1ʱ��Ч��������㷽ʽʱ���ö�θù��� 
    v_�˷ѽ��� Varchar2(2000);
    n_У�Ա�־ ����Ԥ����¼.У�Ա�־%Type;
  Begin
    If Nvl(����ģʽ_In, 0) = 1 And Nvl(��������_In, 0) = 0 Then
      n_У�Ա�־ := 1;
    Else
      n_У�Ա�־ := 2;
    End If;
  
    --���㷽ʽ|������|�������|����ժҪ|���ݺ�|�Ƿ���ͨ���� 
    v_�˷ѽ��� := ���㷽ʽ_In || '|' || �˿���_In || '| |' || ժҪ_In || '|' || No_In || '|0';
    Zl_�����˷ѽ���_Modify(5, ����id_In, ����id_In, v_�˷ѽ���, 0, �����_In, ����_In, ������ˮ��_In, ����˵��_In, 0, 0, 0, 0, Null, 0, Null, Null,
                     ��������id_In, ɾ��ԭ����_In, n_У�Ա�־);
  
    If Nvl(����ģʽ_In, 0) = 1 And Nvl(��������_In, 0) = 0 Then
      Return;
    End If;
  
    --������չ������Ϣ 
    For c_��չ In (Select Extractvalue(j.Column_Value, '/EXPEND/JYMC') As Jymc,
                        Extractvalue(j.Column_Value, '/EXPEND/JYLR') As Jylr
                 From Table(Xmlsequence(Extract(Xmlexpned_In, '/EXPENDLIST/EXPEND'))) J) Loop
      Zl_�������㽻��_Insert(�����_In, 0, ����_In, ����id_In, c_��չ.Jymc || '|' || c_��չ.Jylr);
    End Loop;
  End Thirdcard_Balance;

  Procedure Squarecard_Balance
  (
    ����id_In     ����Ԥ����¼.����id%Type,
    ����id_In     ����Ԥ����¼.����id%Type,
    �����_In     ����Ԥ����¼.�����id%Type,
    ����_In       ����Ԥ����¼.����%Type,
    �˿���_In   ����Ԥ����¼.��Ԥ��%Type,
    ������ˮ��_In ����Ԥ����¼.������ˮ��%Type,
    ����˵��_In   ����Ԥ����¼.����˵��%Type,
    Xmlexpned_In  Xmltype
  ) Is
    v_�˷ѽ��� Varchar2(2000);
    n_���ѿ�id ���ѿ���Ϣ.Id%Type;
  Begin
    Select ID
    Into n_���ѿ�id
    From ���ѿ���Ϣ
    Where �ӿڱ�� = �����_In And ���� = ����_In And
          ��� = (Select Max(���) From ���ѿ���Ϣ Where �ӿڱ�� = �����_In And ���� = ����_In);
  
    --�����ID|����|���ѿ�ID|���ѽ��||. 
    v_�˷ѽ��� := �����_In || '|' || ����_In || '|' || n_���ѿ�id || '|' || �˿���_In;
    Zl_�����˷ѽ���_Modify(4, ����id_In, ����id_In, v_�˷ѽ���, 0, Null, ����_In, ������ˮ��_In, ����˵��_In, 0, 0, 0, 0);
  
    --������չ������Ϣ 
    For c_��չ In (Select Extractvalue(j.Column_Value, '/EXPEND/JYMC') As Jymc,
                        Extractvalue(j.Column_Value, '/EXPEND/JYLR') As Jylr
                 From Table(Xmlsequence(Extract(Xmlexpned_In, '/EXPENDLIST/EXPEND'))) J) Loop
      Zl_�������㽻��_Insert(�����_In, 1, ����_In, ����id_In, c_��չ.Jymc || '|' || c_��չ.Jylr, 0);
    End Loop;
  End Squarecard_Balance;

Begin
  --0.��ȡ����еĲ���ID����Ϣ 
  Select To_Number(Extractvalue(Value(A), 'IN/BRID')), To_Number(Extractvalue(Value(A), 'IN/JE')),
         Extractvalue(Value(A), 'IN/TFZY'), To_Number(Extractvalue(Value(A), 'IN/JCFP')),
         Extractvalue(Value(A), 'IN/JSKLB'), Extractvalue(Value(A), 'IN/SFZH'), Extractvalue(Value(A), 'IN/XM'),
         Extractvalue(Value(A), 'IN/JSMS'), Extractvalue(Value(A), 'IN/CZLX'), Extractvalue(Value(A), 'IN/CXID')
  Into n_����id, n_�˿��ܶ�, v_ժҪ, n_��鷢Ʊ, v_���㿨���, v_���֤��, v_����, n_����ģʽ, n_��������, n_����id
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If Nvl(n_����id, 0) = 0 And Not v_���֤�� Is Null And Not v_���� Is Null Then
    n_����id := Zl_Third_Getpatiid(v_���֤��, v_����);
  End If;

  If Nvl(n_����id, 0) = 0 Then
    v_Err_Msg := '�޷�ȷ��������Ϣ,����!';
    Raise Err_Item;
  End If;

  If Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) <> 0 Then
    --n_����ģʽ --0-��ͨģʽ��1-�첽����ģʽ 
    --n_�������� :����ģʽΪ1ʱ���룬0 - ��ʼ���㣬1 - ��ɽ��㣬2 - ���˽��� 
  
    If Nvl(n_����id, 0) = 0 Then
      v_Err_Msg := 'û��ָ����صĽ������ݣ�';
      Raise Err_Item;
    End If;
  
    Begin
      Select �տ�ʱ��
      Into d_�˷�ʱ��
      From ����Ԥ����¼
      Where ����id = n_����id And Nvl(У�Ա�־, 0) = 1 And Rownum < 2;
    Exception
      When Others Then
        v_Err_Msg := 'û���ҵ�ָ������ؽ������ݣ������ѱ�����';
        Raise Err_Item;
    End;
  
    Select f_List2Str(Cast(Collect(To_Char(a.����id)) As t_StrList), ',', 1)
    Into v_����ids
    From ������ü�¼ A, ������ü�¼ B
    Where a.No = b.No And a.��¼���� = b.��¼���� And a.��¼���� = 1 And a.��¼״̬ In (1, 3) And b.����id = n_����id;
  End If;

  If Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) = 2 Then
    --ɾ���������ݣ��ָ����۵� 
    --n_����ģʽ --0-��ͨģʽ��1-�첽����ģʽ 
    --n_�������� :����ģʽΪ1ʱ���룬0 - ��ʼ���㣬1 - ��ɽ��㣬2 - ���˽��� 
  
    Zl_���˽����¼_Delete(n_����id);
    Zl_�����˷ѽ���_Cancel(n_����id);
  
    v_Temp  := '<CZSJ>' || To_Char(d_�˷�ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    v_Temp  := v_Temp || '<YJZID>' || v_����ids || '</YJZID>';
    v_Temp  := v_Temp || '<CXID>' || n_����id || '</CXID>';
    v_Temp  := v_Temp || '<KPBZ>' || 0 || '</KPBZ>';
    v_Temp  := v_Temp || '<URL>' || '' || '</URL>';
    v_Temp  := v_Temp || '<NETURL>' || '' || '</NETURL>';
    v_Temp  := v_Temp || '<FPTT>' || v_���� || '</FPTT>';
    v_Temp  := v_Temp || '<FPH>' || '' || '</FPH>';
    v_Temp  := v_Temp || '<FPJE>' || '' || '</FPJE>';
    v_Temp  := v_Temp || '<KPRQ>' || '' || '</KPRQ>';
    Xml_Out := Xmltype('<OUTPUT>' || v_Temp || '</OUTPUT>');
    Return;
  End If;

  --��Աid,��Ա���,��Ա���� 
  v_Temp       := Zl_Identity(1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_����Ա���� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_����Ա���� := v_Temp;

  If v_���㿨��� Is Not Null Then
    n_�����id := Get_Cardtypeid(v_���㿨���, 0, v_���㷽ʽ);
    If Nvl(n_�����id, 0) = 0 Then
      v_Err_Msg := '�޷�ȷ�ϴ���Ľ��㿨��';
      Raise Err_Item;
    End If;
  End If;

  If Nvl(n_����ģʽ, 0) = 0 Or Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) = 0 Then
    --n_����ģʽ --0-��ͨģʽ��1-�첽����ģʽ 
    --n_�������� :����ģʽΪ1ʱ���룬0 - ��ʼ���㣬1 - ��ɽ��㣬2 - ���˽��� 
    --1.�Ƚ����˷� 
    Select ���˽��ʼ�¼_Id.Nextval, Sysdate Into n_����id, d_�˷�ʱ�� From Dual;
    n_Billcount  := 0;
    n_ԭ������� := 0;
    For c_���� In (Select Extractvalue(b.Column_Value, '/FY/DJH') As ���ݺ�, Extractvalue(b.Column_Value, '/FY/XH') As �˿����
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/FYLIST/FY'))) B) Loop
      Begin
        Select �������, ����id, ����id
        Into n_�������, n_����id, n_���ݲ���id
        From ����Ԥ����¼
        Where ����id In (Select ����id
                       From ������ü�¼
                       Where NO = c_����.���ݺ� And ��¼���� = 1 And Nvl(����״̬, 0) = 0 And ��¼״̬ In (1, 3) And Rownum < 2) And
              Rownum < 2;
      Exception
        When Others Then
          n_������� := Null;
      End;
    
      If Instr(',' || v_����ids || ',', ',' || n_����id || ',') = 0 Then
        v_����ids := v_����ids || ',' || n_����id;
        --��Ҫ�Ե���Ʊ�ݳ�촦��
        If b_Einvoice_Request.Einvoice_Cancel(1, n_����id, v_Err_Msg) = 0 Then
          --����Ʊ������ʧ�� 
          Raise Err_Item;
        End If;
      End If;
      If n_������� Is Null Then
        v_Err_Msg := 'ָ���ĵ��ݺ�:' || c_����.���ݺ� || 'δ�ҵ�,�����˷�!';
        Raise Err_Item;
      End If;
    
      --�ҺŻ���ģʽ
      Select Max(Substr(a.ժҪ, 4))
      Into v_�Һŵ�
      From ������ü�¼ A
      Where a.No = c_����.���ݺ� And a.��¼���� = 1 And a.ժҪ Like '�Һ�:%' And Rownum < 2;
      If v_�Һŵ� Is Not Null Then
        Select Max(�շѵ�) Into v_�շѵ� From ���˹Һż�¼ Where NO = v_�Һŵ�;
        If v_�շѵ� Is Not Null Then
          Select Count(1)
          Into n_Count
          From ������ü�¼ A
          Where a.��¼���� = 1 And a.��¼״̬ = 1 And a.No <> c_����.���ݺ� And a.��� = 1 And
                a.No In (Select /*+ cardinality(b, 10) */
                          Column_Value
                         From Table(f_Str2List(v_�շѵ�)) B);
          If n_Count = 0 Then
            l_�Һŵ�.Extend;
            l_�Һŵ�(l_�Һŵ�.Count) := v_�Һŵ�;
          End If;
        End If;
      End If;
    
      If Nvl(n_���ݲ���id, 0) = 0 Then
        Select Nvl(Max(����id), 0)
        Into n_���ݲ���id
        From ������ü�¼
        Where NO = c_����.���ݺ� And ��¼���� = 1 And Nvl(����״̬, 0) = 0 And ��¼״̬ In (1, 3) And Rownum < 2;
      End If;
    
      If Nvl(n_����id, 0) <> Nvl(n_���ݲ���id, 0) Then
        v_Err_Msg := '�����˷ѵ��շѵ�:' || c_����.���ݺ� || '���ǵ�ǰ���˵��շѵ�,�����˷�!';
        Raise Err_Item;
      End If;
    
      If n_ԭ������� <> 0 And n_ԭ������� <> n_������� Then
        v_Err_Msg := '�����˷ѵĵ��ݲ���һ���շѽ���,�����˷�!';
        Raise Err_Item;
      End If;
    
      n_ԭ������� := n_�������;
      Select Count(1) Into n_Count From ���ò����¼ Where �շѽ���id = n_����id And Rownum < 2;
      If n_Count <> 0 Then
        v_Err_Msg := '�����˷ѵĵ����Ѿ������˱��ղ������,�����˷�!';
        Raise Err_Item;
      End If;
    
      If v_���㿨��� Is Not Null Then
        Select Count(1) Into n_Count From ����Ԥ����¼ Where ����id = n_����id And �����id = n_�����id;
        If n_Count = 0 Then
          v_Err_Msg := '�����˷ѵĵ��ݲ���' || v_���㷽ʽ || '�����,�����˷�!';
          Raise Err_Item;
        End If;
      End If;
    
      If Nvl(n_��鷢Ʊ, 0) = 1 Then
        Select Max(Decode(a.ʵ��Ʊ��, Null, 0, 1))
        Into n_�Ƿ��ӡ
        From ������ü�¼ A
        Where NO = c_����.���ݺ� And ��¼���� = 1;
        If Nvl(n_�Ƿ��ӡ, 0) = 1 Then
          v_Err_Msg := '�����˷ѵĵ��ݺ��ѿ���Ʊ,�����˷�!';
          Raise Err_Item;
        End If;
      End If;
    
      Zl_�����շѼ�¼_����(c_����.���ݺ�, v_����Ա����, v_����Ա����, c_����.�˿����, d_�˷�ʱ��, v_ժҪ, n_����id);
      n_Billcount := n_Billcount + 1;
    End Loop;
    If n_Billcount = 0 Then
      v_Err_Msg := 'δȷ��������Ҫ�˷ѵĵ���,�����˷�!';
      Raise Err_Item;
    End If;
  
    --����ܽ���Ƿ���ȷ 
    Select Sum(���ʽ��) Into n_���ʽ�� From ������ü�¼ Where ����id = n_����id;
    n_���� := -1 * Nvl(n_���ʽ��, 0) - Nvl(n_�˿��ܶ�, 0);
    If Abs(n_����) > 1.00 Then
      v_Err_Msg := '���ݽɿ�����ʵ�ʽ�������̫��!';
      Raise Err_Item;
    End If;
  End If;

  --2.�����˷ѵĽ�����Ϣ 
  --2.ȷ��֧����ʽ 
  n_ɾ��ԭ���� := 0;
  If Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) = 1 Then
    n_ɾ��ԭ���� := 1;
  End If;
  For c_���㷽ʽ In (Select Extractvalue(b.Column_Value, '/TK/TKKLB') As �����, Extractvalue(b.Column_Value, '/TK/TKKH') As ����,
                        Extractvalue(b.Column_Value, '/TK/TKFS') As ���㷽ʽ,
                        -1 * To_Number(Extractvalue(b.Column_Value, '/TK/TKJE')) As �˿���,
                        Extractvalue(b.Column_Value, '/TK/JYLSH') As ������ˮ��,
                        Extractvalue(b.Column_Value, '/TK/JYSM') As ����˵��, Extractvalue(b.Column_Value, '/TK/TKZY') As ժҪ,
                        Extractvalue(b.Column_Value, '/TK/TYJK') As �Ƿ���Ԥ��,
                        Extractvalue(b.Column_Value, '/TK/SFXFK') As �Ƿ����ѿ�,
                        Extractvalue(b.Column_Value, '/TK/DJH') As ���ݺ�,
                        Extract(b.Column_Value, '/TK/EXPENDLIST') As Expend
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/TKLIST/TK'))) B) Loop
  
    n_���ѿ� := Nvl(c_���㷽ʽ.�Ƿ����ѿ�, 0);
    --1.���������� 
    If c_���㷽ʽ.����� Is Not Null And n_���ѿ� = 0 Then
      n_�����id := Get_Cardtypeid(c_���㷽ʽ.�����, 0, v_���㷽ʽ);
      Select Max(��������id)
      Into n_��������id
      From ����Ԥ����¼ A,
           (Select a.����id
             From ������ü�¼ A, ������ü�¼ B
             Where a.No = b.No And Mod(a.��¼����, 10) = 1 And a.��¼״̬ In (1, 3) And b.����id = n_����id) B
      Where a.����id = b.����id And Mod(a.��¼����, 10) <> 1 And a.�����id = n_�����id;
    
      Thirdcard_Balance(n_����id, n_����id, Nvl(c_���㷽ʽ.���㷽ʽ, v_���㷽ʽ), n_�����id, c_���㷽ʽ.����, c_���㷽ʽ.�˿���, c_���㷽ʽ.������ˮ��,
                        c_���㷽ʽ.����˵��, c_���㷽ʽ.���ݺ�, n_��������id, c_���㷽ʽ.ժҪ, c_���㷽ʽ.Expend, n_����ģʽ, n_��������, n_ɾ��ԭ����);
      n_ɾ��ԭ���� := 0;
    
      --��ɽ���ʱ�Ŵ��������������� 
    Elsif Nvl(n_����ģʽ, 0) = 0 Or Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) = 1 Then
      --2.���ѿ����� 
      If c_���㷽ʽ.����� Is Not Null And n_���ѿ� = 1 Then
        n_�����id := Get_Cardtypeid(c_���㷽ʽ.�����, 1, v_���㷽ʽ);
        Squarecard_Balance(n_����id, n_����id, n_�����id, c_���㷽ʽ.����, c_���㷽ʽ.�˿���, c_���㷽ʽ.������ˮ��, c_���㷽ʽ.����˵��, c_���㷽ʽ.Expend);
      
        --3.��Ԥ���� 
      Elsif Nvl(c_���㷽ʽ.�Ƿ���Ԥ��, 0) = 1 Then
        Zl_�����˷ѽ���_Modify(1, n_����id, n_����id, Null, c_���㷽ʽ.�˿���, Null, Null, Null, Null, 0, 0, 0, 0);
      
        --4.��ͨ���� 
      Else
        If c_���㷽ʽ.���㷽ʽ Is Null Then
          v_Err_Msg := 'δָ��ָ����ʽ�����ʽɿ�!';
          Raise Err_Item;
        End If;
      
        --���㷽ʽ|������|�������|����ժҪ||.. 
        v_�˷ѽ��� := c_���㷽ʽ.���㷽ʽ || '|' || c_���㷽ʽ.�˿��� || '| |' || Nvl(c_���㷽ʽ.ժҪ, '  ');
        v_��ͨ���� := Nvl(v_��ͨ����, '') || '||' || v_�˷ѽ���;
      End If;
    End If;
  End Loop;

  If Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) = 0 Then
    --n_����ģʽ --0-��ͨģʽ��1-�첽����ģʽ 
    --n_�������� :����ģʽΪ1ʱ���룬0 - ��ʼ���㣬1 - ��ɽ��㣬2 - ���˽��� 
  
    v_Temp  := '<CZSJ>' || To_Char(d_�˷�ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    v_Temp  := v_Temp || '<YJZID>' || v_����ids || '</YJZID>';
    v_Temp  := v_Temp || '<CXID>' || n_����id || '</CXID>';
    v_Temp  := v_Temp || '<KPBZ>' || 0 || '</KPBZ>';
    v_Temp  := v_Temp || '<URL>' || '' || '</URL>';
    v_Temp  := v_Temp || '<NETURL>' || '' || '</NETURL>';
    v_Temp  := v_Temp || '<FPTT>' || v_���� || '</FPTT>';
    v_Temp  := v_Temp || '<FPH>' || '' || '</FPH>';
    v_Temp  := v_Temp || '<FPJE>' || '' || '</FPJE>';
    v_Temp  := v_Temp || '<KPRQ>' || '' || '</KPRQ>';
    Xml_Out := Xmltype('<OUTPUT>' || v_Temp || '</OUTPUT>');
  
    Return;
  End If;

  --5.��ͨ���㼰��ɽ��� 
  If v_��ͨ���� Is Not Null Then
    v_��ͨ���� := Substr(v_��ͨ����, 3);
  End If;

  Zl_�����˷ѽ���_Modify(1, n_����id, n_����id, v_��ͨ����, 0, Null, Null, Null, Null, 0, 0, n_����, 2);

  If v_����ids Is Not Null Then
    v_����ids := Substr(v_����ids, 2);
  End If;

  If l_�Һŵ�.Count <> 0 Then
    For I In 0 .. l_�Һŵ�.Count Loop
      v_Temp := '<GHDH>' || l_�Һŵ�(I) || '</GHDH>';
      v_Temp := v_Temp || '<JSKLB>' || v_���㿨��� || '</JSKLB>';
      v_Temp := v_Temp || '<GHJE>' || 0 || '</GHJE>';
      Zl_Third_Registdel(Xmltype('<IN>' || v_Temp || '</IN>'), Xml_Out);
    End Loop;
  End If;

  If Nvl(n_����ģʽ, 0) = 0 Or Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) = 1 Then
    --n_����ģʽ --0-��ͨģʽ��1-�첽����ģʽ 
    --n_�������� :����ģʽΪ1ʱ���룬0 - ��ʼ���㣬1 - ��ɽ��㣬2 - ���˽��� 
    --����Ʊ�ݴ���
  
    For c_ԭ���� In (Select Distinct a.����id
                  From ������ü�¼ A, ������ü�¼ B
                  Where a.No = b.No And a.��¼���� = b.��¼���� And a.��¼���� = 1 And a.��¼״̬ In (1, 3) And b.����id = n_����id) Loop
      Select Sum(���ʽ��)
      Into n_ʣ����
      From ������ü�¼
      Where NO In (Select Distinct NO From ������ü�¼ Where ����id = c_ԭ����.����id) And Mod(��¼����, 10) = 1;
    
      Select Max(�Ƿ����Ʊ��) Into n_�Ƿ����Ʊ�� From ����Ԥ����¼ Where ����id = c_ԭ����.����id;
    
      If Nvl(n_ʣ����, 0) <> 0 And Nvl(n_�Ƿ����Ʊ��, 0) = 1 Then
        --�����ˣ���Ҫ���¿��ߵ���Ʊ��
        If b_Einvoice_Request.Einvoice_Create(1, c_ԭ����.����id, n_����id, v_Err_Msg) = 0 Then
          Raise Err_Item;
        End If;
        Select Max(1), Max(����), Max(����), Max(To_Char(����ʱ��, 'yyyy-mm-dd')), Max(Url����), Max(Url����), Max(Ʊ�ݽ��)
        Into n_��Ʊ��־, v_��������, v_��Ʊ���, v_��Ʊ����, v_Url, v_Url����, n_��Ʊ���
        From ����Ʊ��ʹ�ü�¼
        Where ����id = n_����id And Ʊ�� = 1 And ��¼״̬ = 1;
      
        If v_�������� Is Not Null Then
          v_���� := v_��������;
        End If;
      End If;
    End Loop;
  End If;
  v_Temp  := '<CZSJ>' || To_Char(d_�˷�ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
  v_Temp  := v_Temp || '<YJZID>' || v_����ids || '</YJZID>';
  v_Temp  := v_Temp || '<CXID>' || n_����id || '</CXID>';
  v_Temp  := v_Temp || '<KPBZ>' || Nvl(n_��Ʊ��־, 0) || '</KPBZ>';
  v_Temp  := v_Temp || '<URL>' || Nvl(v_Url, '') || '</URL>';
  v_Temp  := v_Temp || '<NETURL>' || Nvl(v_Url����, '') || '</NETURL>';
  v_Temp  := v_Temp || '<FPTT>' || v_���� || '</FPTT>';
  v_Temp  := v_Temp || '<FPH>' || v_��Ʊ��� || '</FPH>';
  v_Temp  := v_Temp || '<FPJE>' || Nvl(n_��Ʊ���, 0) || '</FPJE>';
  v_Temp  := v_Temp || '<KPRQ>' || v_��Ʊ���� || '</KPRQ>';
  Xml_Out := Xmltype('<OUTPUT>' || v_Temp || '</OUTPUT>');

Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Charge_Del;
/

Create Or Replace Procedure Zl_Third_Deposit_Recharge
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --����:Ԥ�����ֵ
  --���:Xml_In:
  --    <IN>
  --        <BRID>����ID</BRID>
  --        <ZYID>��ҳID</ZYID>
  --        <XM>����</XM>
  --        <SFZH>���֤��</SFZH>
  --        <SFMZ>�Ƿ�����</SFMZ>   //1-������,0-סԺ
  --        <SFYJ>�Ƿ�Ѻ��</SFYJ>   //�Ƿ�ΪѺ��0-Ԥ���ɿ1-Ѻ��ɿ�
  --        <JSMS>0</JSMS>          //����ģʽ��0-��ͨģʽ��1-�첽����ģʽ
  --        <CZLX>0</CZLX>          //�������ͣ�����ģʽΪ1ʱ���룬0-��ʼ���㣬1-��ɽ��㣬2-���˽���
  --        <YJDH></YJDH>           //Ԥ�����ţ���������Ϊ1��2ʱ����
  --    <ZFBZH>֧�������ں�UserID</ZFBZH>
  --    <ZFBXCY>֧����С����UserID</ZFBXCY>
  --    <WXGZHID>΢�Ź��ں�OpenID</WXGZH>
  --    <WXXCXID>΢��С����OpenID</WXXCXID>
  --        <JSLIST>                //�����б���������Ϊ2ʱ�ɲ�����
  --            <JS>
  --              <JSKLB>֧�������</JSKLB >
  --              <JSKH>֧������</JSKH>
  --              <JYLSH>������ˮ��</JYLSH>
  --              <JYSM>����˵��</JYSM>
  --              <JSFS>֧����ʽ</JSFS> //��ֵ��ʽ:�ֽ�;֧Ʊ,�����������,���Դ���
  --              <JSJE>���׽��</JSJE> //��ֵ���
  --              <ZY>ժҪ</ZY>
  --              <SFXFK>�Ƿ����ѿ�</SFXFK>
  --              <JSHM>�������(���Բ���)</JSHM>
  --              <JKDW>�ɿλ(���Բ���)</JKDW>
  --              <DWKFH>��λ������(���Բ���)</DWKFH>
  --              <DWZH>��λ�ʺ�(���Բ���)</DWZH>
  --              <HZDW>������λ(���Բ���)</HZDW>
  --              <EXPENDLIST>         //��չ������Ϣ
  --                   <EXPEND>
  --                        <JYMC>��������</JYMC>
  --                        <JYLR>��������</JYLR>
  --                   </EXPEND>
  --              </EXPENDLIST >
  --            </JS>
  --         </JSLIST>
  --    </IN>
  --����:Xml_Out
  --  <OUTPUT>
  --    <CZSJ>����ʱ��</CZSJ>          //HIS�ĵǼ�ʱ��
  --    <YJDH>Ԥ������(������ŷָ�)</YJDH>
  --  <KPBZ>��Ʊ��־</KPBZ> //1-�ɹ����ߵ���Ʊ��;0-δ��Ʊ�ɹ���־
  --  <URL>H5ҳ��URL</URL>
  --  <NETURL>����H5ҳ��URL</NETURL>
  --  <FPTT>��Ʊ̧ͷ</FPTT>        //��������
  --  <FPH>��Ʊ��</FPH>             //��Ʊ���
  --  <FPJE>��Ʊ���</FPJE>        //100.00
  --  <KPRQ>��Ʊ����</KPRQ>   //yyyy-mm-dd
  --    �D�D�������д�������˵����ȷִ��
  --    <ERROR>
  --      <MSG>������Ϣ</MSG>
  --    </ERROR>
  --  </OUTPUT>
  --------------------------------------------------------------------------------------------------
  v_���㷽ʽ   Varchar2(2000);
  v_No         ����Ԥ����¼.No%Type;
  v_����Ա���� ����Ԥ����¼.����Ա���%Type;
  v_����Ա���� ����Ԥ����¼.����Ա����%Type;
  n_����ģʽ   Number(2); --0-��ͨģʽ��1-�첽����ģʽ
  n_��������   Number(2); --����ģʽΪ1ʱ���룬0-��ʼ���㣬1-��ɽ��㣬2-���˽��� 

  n_�����id           ҽ�ƿ����.Id%Type;
  n_����id             ������ü�¼.����id%Type;
  v_����               ������Ϣ.����%Type;
  v_���֤��           ������Ϣ.���֤��%Type;
  n_��ҳid             ����Ԥ����¼.��ҳid%Type;
  n_����id             ����Ԥ����¼.����id%Type;
  n_Ԥ��id             ����Ԥ����¼.Id%Type;
  n_���㿨���         ����Ԥ����¼.���㿨���%Type;
  n_Ԥ�����           ����Ԥ����¼.Ԥ�����%Type;
  n_���ѿ�             Number(2);
  n_����Ԥ��           Number(2);
  v_�����             �������׼�¼.���%Type;
  n_Step               Number(2);
  d_�Ǽ�ʱ��           Date;
  n_����               Number(1);
  n_״̬               Number(1);
  n_У�Ա�־           ����Ԥ����¼.У�Ա�־%Type;
  n_Billcount          Number;
  n_�Ƿ�Ѻ��           Number(1);
  n_Ԥ������Ʊ��       ����Ԥ����¼.Ԥ������Ʊ��%Type;
  v_֧�������ں�userid Varchar2(100);
  v_֧����С����userid Varchar2(100);
  v_΢�Ź��ں�openid   Varchar2(100);
  v_΢��С����openid   Varchar2(100);
  n_��Ʊ��־           Number(2);
  v_��������           ����Ʊ��ʹ�ü�¼.����%Type;
  v_��Ʊ���           ����Ʊ��ʹ�ü�¼.����%Type;
  v_��Ʊ����           Varchar2(20);
  n_��Ʊ���           ����Ʊ��ʹ�ü�¼.Ʊ�ݽ��%Type;
  v_Url                ����Ʊ��ʹ�ü�¼.Url����%Type;
  v_Url����            ����Ʊ��ʹ�ü�¼.Url����%Type;

  v_Temp    Varchar2(32767); --��ʱXML
  v_Err_Msg Varchar2(200);
  Err_Special Exception;
  Err_Item    Exception;

  Function Get������
  (
    �����_In Varchar2,
    ���ѿ�_In Number
  ) Return Varchar2 As
    v_����� Varchar2(200);
    n_Num    Number(1);
  Begin
    Select Decode(Translate(�����_In, '#1234567890', '#'), Null, 1, 0) Into n_Num From Dual;
    If Nvl(���ѿ�_In, 0) = 1 Then
      If Nvl(n_Num, 0) = 1 Then
        Select Max(����) Into v_����� From ���ѿ����Ŀ¼ Where ��� = To_Number(�����_In);
      Else
        Select Max(����) Into v_����� From ���ѿ����Ŀ¼ Where ���� = �����_In;
      End If;
    Else
      If Nvl(n_Num, 0) = 1 Then
        Select Max(����) Into v_����� From ҽ�ƿ���� Where ID = To_Number(�����_In);
      Else
        Select Max(����) Into v_����� From ҽ�ƿ���� Where ���� = �����_In;
      End If;
    End If;
    Return v_�����;
  End Get������;

  Function Get���㷽ʽ
  (
    �����_In    Varchar2,
    ���ѿ�_In    Number,
    �����id_Out Out ����Ԥ����¼.�����id%Type
  ) Return Varchar2 As
    --�����_In ���������
  Begin
    If Nvl(���ѿ�_In, 0) = 1 Then
      Begin
        Select ���, ���㷽ʽ, Decode(Nvl(����, 0), 1, Null, ���� || 'δ���ã���������нɷѣ�')
        Into �����id_Out, v_���㷽ʽ, v_Err_Msg
        From ���ѿ����Ŀ¼
        Where ���� = �����_In;
      Exception
        When Others Then
          v_Err_Msg := �����_In || '�����ڣ�';
      End;
    
      If v_Err_Msg Is Not Null Then
        Raise Err_Item;
      End If;
      If v_���㷽ʽ Is Null Then
        v_Err_Msg := �����_In || 'δ���ý��㷽ʽ���������ѿ����������ý��㷽ʽ��';
        Raise Err_Item;
      End If;
    Else
      Begin
        Select ID, ���㷽ʽ, Decode(Nvl(�Ƿ�����, 0), 1, Null, ���� || 'δ���ã���������нɷѣ�')
        Into �����id_Out, v_���㷽ʽ, v_Err_Msg
        From ҽ�ƿ����
        Where ���� = �����_In;
      Exception
        When Others Then
          v_Err_Msg := �����_In || '�����ڣ�';
      End;
    
      If v_Err_Msg Is Not Null Then
        Raise Err_Item;
      End If;
      If v_���㷽ʽ Is Null Then
        v_Err_Msg := �����_In || 'δ���ý��㷽ʽ������ҽ�ƿ���������ý��㷽ʽ��';
        Raise Err_Item;
      End If;
    End If;
  
    Return v_���㷽ʽ;
  End Get���㷽ʽ;
Begin
  Select Extractvalue(Value(A), 'IN/BRID'), To_Number(Extractvalue(Value(A), 'IN/ZYID')),
         To_Number(Extractvalue(Value(A), 'IN/SFMZ')), Extractvalue(Value(A), 'IN/SFZH'),
         Extractvalue(Value(A), 'IN/XM'), Extractvalue(Value(A), 'IN/JSMS'), Extractvalue(Value(A), 'IN/CZLX'),
         Extractvalue(Value(A), 'IN/YJDH'), Extractvalue(Value(A), 'IN/SFYJ'), Extractvalue(Value(A), 'IN/ZFBZH'),
         Extractvalue(Value(A), 'IN/ZFBXCY'), Extractvalue(Value(A), 'IN/WXGZHID'), Extractvalue(Value(A), 'IN/WXXCXID')
  Into n_����id, n_��ҳid, n_����Ԥ��, v_���֤��, v_����, n_����ģʽ, n_��������, v_No, n_�Ƿ�Ѻ��, v_֧�������ں�userid, v_֧����С����userid, v_΢�Ź��ں�openid,
       v_΢��С����openid
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If Nvl(n_����Ԥ��, 0) = 1 And Nvl(n_����id, 0) = 0 Then
    If Not v_���֤�� Is Null And Not v_���� Is Null Then
      n_����id := Zl_Third_Getpatiid(v_���֤��, v_����);
    End If;
  End If;

  --0.��ؼ��
  If Nvl(n_����id, 0) = 0 Then
    v_Err_Msg := '������Чʶ������ݣ��������ֵ��';
    Raise Err_Item;
  End If;

  If Not v_֧�������ں�userid Is Null Then
    Zl_������Ϣ�ӱ�_Update(n_����id, Upper('֧�������ں�UserID'), v_֧�������ں�userid);
  End If;

  If Not v_֧����С����userid Is Null Then
    Zl_������Ϣ�ӱ�_Update(n_����id, Upper('֧����С����UserID'), v_֧�������ں�userid);
  End If;

  If Not v_΢�Ź��ں�openid Is Null Then
    Zl_������Ϣ�ӱ�_Update(n_����id, Upper('΢�Ź��ں�OpenID'), v_֧�������ں�userid);
  End If;

  If Not v_΢��С����openid Is Null Then
    Zl_������Ϣ�ӱ�_Update(n_����id, Upper('΢��С����OpenID'), v_֧�������ں�userid);
  End If;

  Begin
    Select Nullif(Nvl(a.��ǰ����id, b.��Ժ����id), 0)
    Into n_����id
    From ������Ϣ A, ������ҳ B
    Where a.����id = b.����id(+) And a.��ҳid = b.��ҳid(+) And a.����id = n_����id;
  Exception
    When Others Then
      v_Err_Msg := '������Чʶ������ݣ��������ֵ��';
      Raise Err_Item;
  End;

  If Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) <> 0 Then
    If v_No Is Null Then
      v_Err_Msg := 'û��ָ����صĽ������ݣ�';
      Raise Err_Item;
    End If;
  
    Begin
      If n_�Ƿ�Ѻ�� = 1 Then
        Select �տ�ʱ��, ID
        Into d_�Ǽ�ʱ��, n_Ԥ��id
        From ����Ѻ���¼
        Where ��¼״̬ = 0 And NO = v_No And Nvl(У�Ա�־, 0) = 1 And Rownum < 2;
      Else
        Select �տ�ʱ��, ID
        Into d_�Ǽ�ʱ��, n_Ԥ��id
        From ����Ԥ����¼
        Where ��¼���� = 1 And ��¼״̬ = 0 And NO = v_No And Nvl(У�Ա�־, 0) = 1 And Rownum < 2;
      End If;
    Exception
      When Others Then
        v_Err_Msg := 'û���ҵ�ָ������ؽ������ݣ������ѱ�����';
        Raise Err_Item;
    End;
  End If;

  If Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) = 2 Then
    --ɾ���������ݣ��ָ����۵�
    If n_�Ƿ�Ѻ�� = 1 Then
      Zl_����Ѻ���쳣��¼_Delete(v_No);
    Else
      Zl_����Ԥ���쳣��¼_Delete(v_No);
    End If;
    v_Temp  := '<YJDH>' || v_No || '</YJDH>';
    v_Temp  := v_Temp || '<CZSJ>' || To_Char(d_�Ǽ�ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    v_Temp  := v_Temp || '<KPBZ>' || 0 || '</KPBZ>';
    v_Temp  := v_Temp || '<URL>' || '' || '</URL>';
    v_Temp  := v_Temp || '<NETURL>' || '' || '</NETURL>';
    v_Temp  := v_Temp || '<FPTT>' || v_���� || '</FPTT>';
    v_Temp  := v_Temp || '<FPH>' || '' || '</FPH>';
    v_Temp  := v_Temp || '<FPJE>' || '' || '</FPJE>';
    v_Temp  := v_Temp || '<KPRQ>' || '' || '</KPRQ>';
    Xml_Out := Xmltype('<OUTPUT>' || v_Temp || '</OUTPUT>');
    Return;
  End If;

  --����Ա��Ϣ:����ID,��������;��ԱID,��Ա���,��Ա����
  v_Temp := Zl_Identity;
  If Nvl(v_Temp, '0') = '0' Or Nvl(v_Temp, '_') = '_' Then
    v_Err_Msg := '����ʶ����Ч�Ĳ���Ա��������ɷѣ�';
    Raise Err_Item;
  End If;
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_����Ա���� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_����Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);

  If Nvl(n_����ģʽ, 0) = 0 Or Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) = 0 Then
    For c_���׼�¼ In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As ���㿨���,
                          Extractvalue(b.Column_Value, '/JS/JSFS') As ���㷽ʽ,
                          Extractvalue(b.Column_Value, '/JS/JYLSH') As ������ˮ��,
                          Extractvalue(b.Column_Value, '/JS/SFXFK') As �Ƿ����ѿ�,
                          Extractvalue(b.Column_Value, '/JS/JSKH') As ���㿨��, Extractvalue(b.Column_Value, '/JS/ZY') As ժҪ
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
    
      If c_���׼�¼.���㿨��� Is Null Then
        v_����� := c_���׼�¼.���㷽ʽ;
      Else
        v_����� := Get������(c_���׼�¼.���㿨���, c_���׼�¼.�Ƿ����ѿ�);
      End If;
      If v_����� Is Null Then
        v_Err_Msg := '��֧�ֵĽ��㷽ʽ�����飡';
        Raise Err_Item;
      End If;
    
      --����һ�����㷽ʽ�ż�齻����
      n_Step := Nvl(n_Step, 0) + 1;
      If Zl_Fun_�������׼�¼_Locked(v_�����, c_���׼�¼.������ˮ��, c_���׼�¼.���㿨��, c_���׼�¼.ժҪ, 1) = 0 And n_Step = 1 Then
        v_Err_Msg := '������ˮ��Ϊ:' || c_���׼�¼.������ˮ�� || '�Ľ������ڽ����У��������ٴ��ύ�˽��ף�';
        Raise Err_Special;
      End If;
    End Loop;
  End If;

  --2.ȷ��֧����ʽ 
  If Nvl(n_����Ԥ��, 0) = 0 Then
    n_Ԥ����� := 2;
  Else
    n_Ԥ����� := 1;
  End If;
  d_�Ǽ�ʱ��  := Sysdate;
  n_Billcount := 0;
  For c_���㷽ʽ In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As ���㿨���,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As ���㿨��,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As ���㷽ʽ,
                        Extractvalue(b.Column_Value, '/JS/JSJE') As ������,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As ������ˮ��,
                        Extractvalue(b.Column_Value, '/JS/JYSM') As ����˵��, Extractvalue(b.Column_Value, '/JS/ZY') As ժҪ,
                        Extractvalue(b.Column_Value, '/JS/SFXFK') As �Ƿ����ѿ�,
                        Extractvalue(b.Column_Value, '/JS/JSHM') As �������,
                        Extractvalue(b.Column_Value, '/JS/JKDW') As �ɿλ,
                        Extractvalue(b.Column_Value, '/JS/DWKFH') As ��λ������,
                        Extractvalue(b.Column_Value, '/JS/DWZH') As ��λ�ʺ�,
                        Extractvalue(b.Column_Value, '/JS/HZDW') As ������λ,
                        Extract(b.Column_Value, '/JS/EXPENDLIST') As Expend
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
  
    If Nvl(c_���㷽ʽ.������, 0) = 0 Then
      v_Err_Msg := '����ĳ�ֵ���Ϊ�㣬û��Ҫ���г�ֵ���������ֵ����Ƿ������';
      Raise Err_Item;
    End If;
  
    n_���ѿ�     := Nvl(c_���㷽ʽ.�Ƿ����ѿ�, 0);
    n_���㿨��� := Null;
    n_�����id   := Null;
    If c_���㷽ʽ.���㿨��� Is Null Then
      v_�����   := c_���㷽ʽ.���㷽ʽ;
      v_���㷽ʽ := c_���㷽ʽ.���㷽ʽ;
    Else
      v_�����   := Get������(c_���㷽ʽ.���㿨���, n_���ѿ�);
      v_���㷽ʽ := Get���㷽ʽ(v_�����, n_���ѿ�, n_�����id);
      If Nvl(n_���ѿ�, 0) = 1 Then
        n_���㿨��� := n_�����id;
        n_�����id   := Null;
      End If;
    End If;
    If v_���㷽ʽ Is Null Then
      v_Err_Msg := 'δȷ�����γ�ֵ��֧����ʽ������֧����ʽ�Ƿ������';
      Raise Err_Item;
    End If;
  
    If Nvl(n_����ģʽ, 0) = 0 Or Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) = 0 Then
      If Nvl(n_Billcount, 0) > 0 Then
        v_Err_Msg := 'Ŀǰֻ֧��һ��֧����ʽ�������ֵ��Ϣ�Ƿ������';
        Raise Err_Item;
      End If;
    
      v_No := Nextno(11);
      Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
    End If;
  
    --��������_In:0-������Ԥ��;1-����Ϊδ��Ч��Ԥ����;3-����˿�
    --����״̬_In:0-�������㣬1-����Ϊ�쳣���ݣ�2-����쳣���� 
    If Nvl(n_����ģʽ, 0) = 0 Then
      n_����     := 0;
      n_״̬     := 0;
      n_У�Ա�־ := 0;
    Else
      If Nvl(n_��������, 0) = 0 Then
        n_����     := 1;
        n_״̬     := 1;
        n_У�Ա�־ := 1;
      Else
        n_����     := 0;
        n_״̬     := 2;
        n_У�Ա�־ := 0;
      End If;
    End If;
    If n_�Ƿ�Ѻ�� = 1 Then
      Zl_����Ѻ���¼_Insert(n_Ԥ��id, v_No, Null, n_����id, n_��ҳid, n_����id, c_���㷽ʽ.�ɿλ, c_���㷽ʽ.��λ������, c_���㷽ʽ.��λ�ʺ�, c_���㷽ʽ.ժҪ,
                       c_���㷽ʽ.������, v_���㷽ʽ, c_���㷽ʽ.�������, n_Ԥ�����, Null, v_����Ա����, v_����Ա����, c_���㷽ʽ.���㿨��, n_�����id,
                       c_���㷽ʽ.������ˮ��, c_���㷽ʽ.����˵��, n_У�Ա�־, n_״̬);
    Else
      Zl_����Ԥ����¼_Insert(n_Ԥ��id, v_No, Null, n_����id, n_��ҳid, n_����id, c_���㷽ʽ.������, v_���㷽ʽ, c_���㷽ʽ.�������, c_���㷽ʽ.�ɿλ,
                       c_���㷽ʽ.��λ������, c_���㷽ʽ.��λ�ʺ�, c_���㷽ʽ.ժҪ, v_����Ա����, v_����Ա����, Null, n_Ԥ�����, n_�����id, n_���㿨���,
                       c_���㷽ʽ.���㿨��, c_���㷽ʽ.������ˮ��, c_���㷽ʽ.����˵��, c_���㷽ʽ.������λ, d_�Ǽ�ʱ��, n_����, Null, Null, 0, 0, 1, 0, n_У�Ա�־,
                       n_״̬, 0);
    
    End If;
    If Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) = 0 Then
      v_Temp := '<YJDH>' || v_No || '</YJDH>';
      v_Temp := v_Temp || '<CZSJ>' || To_Char(d_�Ǽ�ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
      v_Temp := v_Temp || '<KPBZ>' || 0 || '</KPBZ>';
      v_Temp := v_Temp || '<URL>' || '' || '</URL>';
      v_Temp := v_Temp || '<NETURL>' || '' || '</NETURL>';
      v_Temp := v_Temp || '<FPTT>' || v_���� || '</FPTT>';
      v_Temp := v_Temp || '<FPH>' || '' || '</FPH>';
      v_Temp := v_Temp || '<FPJE>' || '' || '</FPJE>';
      v_Temp := v_Temp || '<KPRQ>' || '' || '</KPRQ>';
    
      Xml_Out := Xmltype('<OUTPUT>' || v_Temp || '</OUTPUT>');
      Return;
    End If;
  
    If Nvl(n_���ѿ�, 0) = 1 Then
      Zl_�����ӿڸ���_Update(n_���㿨���, 1, c_���㷽ʽ.���㿨��, n_Ԥ��id, c_���㷽ʽ.������ˮ��, c_���㷽ʽ.����˵��, 1, 0);
    End If;
  
    --������չ������Ϣ
    For c_��չ In (Select Extractvalue(j.Column_Value, '/EXPEND/JYMC') As Jymc,
                        Extractvalue(j.Column_Value, '/EXPEND/JYLR') As Jylr
                 From Table(Xmlsequence(Extract(c_���㷽ʽ.Expend, '/EXPENDLIST/EXPEND'))) J) Loop
    
      If Nvl(n_���ѿ�, 0) = 1 Then
        Zl_�������㽻��_Insert(n_���㿨���, 1, c_���㷽ʽ.���㿨��, n_Ԥ��id, c_��չ.Jymc || '|' || c_��չ.Jylr, 1);
      Else
        Zl_�������㽻��_Insert(n_�����id, 0, c_���㷽ʽ.���㿨��, n_Ԥ��id, c_��չ.Jymc || '|' || c_��չ.Jylr, 1 + Nvl(n_�Ƿ�Ѻ��, 0));
      End If;
    End Loop;
    Update �������׼�¼
    Set ҵ�����id = n_Ԥ��id
    Where ��ˮ�� = c_���㷽ʽ.������ˮ�� And ��� = v_����� And ҵ������ = 1;
  
    n_Billcount := n_Billcount + 1;
  End Loop;

  If Nvl(n_Billcount, 0) = 0 Then
    v_Err_Msg := '������Чȷ�ϵ�ǰ��ֵ��֧����ʽ��';
    Raise Err_Item;
  End If;
  If Nvl(n_�Ƿ�Ѻ��, 0) = 0 Then
    --����Ʊ�ݴ���  
    n_Ԥ������Ʊ�� := b_Einvoice_Request.Einvoice_Start(2, Null,n_����Ԥ��);
    Update ����Ԥ����¼ Set Ԥ������Ʊ�� = n_Ԥ������Ʊ�� Where ID = n_Ԥ��id;
    --��Ҫ���ߵ���Ʊ��
    If Nvl(n_Ԥ������Ʊ��, 0) = 1 Then
      If b_Einvoice_Request.Einvoice_Create(2, n_Ԥ��id, Null, v_Err_Msg) = 0 Then
        --����Ʊ�ݿ��߳ɹ�
        Raise Err_Item;
      End If;
    
      Select Max(1), Max(����), Max(����), Max(To_Char(����ʱ��, 'yyyy-mm-dd')), Max(Url����), Max(Url����), Max(Ʊ�ݽ��)
      Into n_��Ʊ��־, v_��������, v_��Ʊ���, v_��Ʊ����, v_Url, v_Url����, n_��Ʊ���
      From ����Ʊ��ʹ�ü�¼
      Where ����id = n_Ԥ��id And Ʊ�� = 2 And ��¼״̬ = 1;
    
      If v_�������� Is Not Null Then
        v_���� := v_��������;
      End If;
    
    End If;
  End If;
  v_Temp  := '<YJDH>' || v_No || '</YJDH>';
  v_Temp  := v_Temp || '<CZSJ>' || To_Char(d_�Ǽ�ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
  v_Temp  := v_Temp || '<KPBZ>' || Nvl(n_��Ʊ��־, 0) || '</KPBZ>';
  v_Temp  := v_Temp || '<URL>' || Nvl(v_Url, '') || '</URL>';
  v_Temp  := v_Temp || '<NETURL>' || Nvl(v_Url����, '') || '</NETURL>';
  v_Temp  := v_Temp || '<FPTT>' || v_���� || '</FPTT>';
  v_Temp  := v_Temp || '<FPH>' || v_��Ʊ��� || '</FPH>';
  v_Temp  := v_Temp || '<FPJE>' || Nvl(n_��Ʊ���, 0) || '</FPJE>';
  v_Temp  := v_Temp || '<KPRQ>' || v_��Ʊ���� || '</KPRQ>';
  Xml_Out := Xmltype('<OUTPUT>' || v_Temp || '</OUTPUT>');
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Err_Special Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20105, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Deposit_Recharge;
/

Create Or Replace Procedure Zl_Third_Settlement
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --����:�����ӿ�֧��
  --���:Xml_In:
  --<IN>
  --  <BRID>����ID</BRID>         //����ID
  --  <XM>����</XM>               //����
  --  <SFZH>���֤��</SFZH>       //���֤��
  --  <ZYID>��ҳID</ZYID>         //��ҳID
  --  <JSLX>2</JSLX>         //��������,1-����,2-סԺ��Ĭ��Ϊ 2
  --  <JE></JE>         //���ν����ܽ��
  --  <NO></NO>         //���ʵķ��õ��ݺ�(������ʵ�),Ŀǰ����������=1ʱ��ʹ��
  --  <JZKNO></JZKNO>   //���ʵľ��￨���ݺ�,Ŀǰ����������=1ʱ��ʹ��
  --  <JZSJ></JZSJ>     //����ʱ��
  --  <JSMS>1</JSMS>    //����ģʽ��0-��ͨģʽ��1-�첽����ģʽ
  --  <CZLX>0</CZLX>    //�������ͣ�����ģʽΪ1ʱ���룬0-��ʼ���㣬1-��ɽ��㣬2-���˽���
  --  <JZID>1</JZID>    //����ID����������Ϊ1��2ʱ����
  --  <ZFBZH>֧�������ں�UserID</ZFBZH>
  --  <ZFBXCY>֧����С����UserID</ZFBXCY>
  --  <WXGZHID>΢�Ź��ں�OpenID</WXGZH>
  --  <WXXCXID>΢��С����OpenID</WXXCXID>
  --  <JSLIST>          //�����б���������Ϊ2ʱ�ɲ����� 
  --    <JS>
  --      <JSKLB>֧�������</JSKLB >
  --      <JSKH>֧������</ JSKH >
  --      <JSFS>֧����ʽ</JSFS> //֧����ʽ:�ֽ�;֧Ʊ,�����������,���Դ���
  --      <JSJE>������</JSJE> //�������Ϊ����SFCYJΪ1ʱΪ�ܵĳ�Ԥ�����
  --      <JYLSH>������ˮ��</JYLSH>
  --      <JYSM>����˵��</JYSM>
  --      <ZY>ժҪ</ZY>
  --      <SFXFK>�Ƿ����ѿ�</SFXFK>  //(1-�����ѿ�),���ѿ�ʱ,������㿨���,���㿨��,������Ƚӵ�
  --      <SFCYJ>�Ƿ��Ԥ��</SFCYJ>  //�Ƿ��Ԥ����0-���㣬1-��Ԥ��.�ʳ�Ԥ��ʱ,ֻ��JSJE�ڵ�
  --      <CYJLIST>  //��Ԥ����ϣ��Ƿ��Ԥ��=1ʱ���룬��ʱ�Դ���ĵ��ݽ��г�Ԥ������������ýڵ㣬��HIS���Ƚ����ù�����г�Ԥ�����ʱ���ܽ�����Ԥ������������Ԥ�����ֻ��ʹ��Ԥ������н��ʣ�
  --        <TKFS>�˿ʽ<TKFS>  //�˿ʽ�������˿�ʱ���룺0-�ֽ����˿�;1-����һ�ν��׽ӿ��˿�;2-ת�ʷ�ʽ�˿�(�ݲ�֧��)(Ӧ������һ�£���Ҫ�����첽ģʽ�£����´��ڴ����쳣����)
  --        <ITEM>  //˵�������˱��г壻�絥��A001��Ԥ������1000�����ν���800������������¼�ٳ�1000������200��Sum(����-�˽��)=���ʽ��
  --          <DJH>Ԥ����ݺ�</DJH>
  --          <JE>���׽��</JE>
  --          <SFTK>�Ƿ���Ԥ����</SFTK>  //�Ƿ���Ԥ���0-��Ԥ����;1-��Ԥ����
  --          <JYLSH>������ˮ��</JYLSH>  //�˿����ˮ��
  --          <JYSM>����˵��</JYSM>  //�˿��˵��
  --          <EXPENDLIST>  //�˿�׵���չ��Ϣ���˿ʽ=0ʱ����
  --            <EXPEND>
  --              <JYMC>��������</JYMC> //��������
  --              <JYLR>��������</JYLR> //��������
  --            </EXPEND>
  --          </EXPENDLIST>
  --        </ITEM>
  --        <EXPENDLIST>  //�˿�׵���չ��Ϣ���˿ʽ=1��2ʱ����
  --          <EXPEND>      
  --            <JYMC>��������</JYMC> //��������
  --            <JYLR>��������</JYLR> //��������
  --          </EXPEND>
  --        </EXPENDLIST>
  --      </CYJLIST>
  --      <EXPENDLIST>  //��չ������Ϣ
  --        <EXPEND>
  --          <JYMC>��������</JYMC> //��������
  --          <JYLR>��������</JYLR> //��������
  --        </EXPEND>
  --      </EXPENDLIST>
  --    </JS>
  --  </JSLIST >
  --</IN>

  --����:Xml_Out
  --<OUT>
  --  <CZSJ>����ʱ��</CZSJ>          //HIS�ĵǼ�ʱ��
  --  <JZID>����ID</JZID>
  --  <KPBZ>��Ʊ��־</KPBZ> //1-�ɹ����ߵ���Ʊ��;0-δ��Ʊ�ɹ���־
  --  <URL>H5ҳ��URL</URL>
  --  <NETURL>����H5ҳ��URL</NETURL>
  --  <FPTT>��Ʊ̧ͷ</FPTT>        //��������
  --  <FPH>��Ʊ��</FPH>             //��Ʊ���
  --  <FPJE>��Ʊ���</FPJE>        //100.00
  --  <KPRQ>��Ʊ����</KPRQ>   //yyyy-mm-dd
  --  <ERROR>  //���޸ô�������˵����ȷִ��
  --    <MSG>������Ϣ</MSG>
  --  </ERROR>
  --</OUT>
  --------------------------------------------------------------------------------------------------
  n_��ҳid       ������ҳ.��ҳid%Type;
  n_����id       ������ҳ.����id%Type;
  v_����         ������Ϣ.����%Type;
  v_���֤��     ������Ϣ.���֤��%Type;
  n_�����ܶ�     ����Ԥ����¼.��Ԥ��%Type;
  n_��������     Number(3);
  v_���￨���ݺ� Varchar2(20000);
  d_����ʱ��     Date;
  v_���ݺ�       Varchar2(20000);
  n_����ģʽ     Number(1); --0-��ͨģʽ��1-�첽����ģʽ
  n_��������     Number(1); --����ģʽΪ1ʱ���룬0 - ��ʼ���㣬1 - ��ɽ��㣬2 - ���˽���

  v_����Ա���� ���˽��ʼ�¼.����Ա���%Type;
  v_����Ա���� ���˽��ʼ�¼.����Ա����%Type;
  n_����id     ���˽��ʼ�¼.Id%Type;
  n_�����ʽ�� ����Ԥ����¼.��Ԥ��%Type;
  d_��ʼ����   Date;
  d_��������   Date;
  d_��С����   Date;
  d_�������   Date;
  n_��������id ����Ԥ����¼.��������id%Type;
  n_Ԥ��id     ����Ԥ����¼.Id%Type;
  n_ɾ��ԭ���� Number;

  v_���ʵ��� ���˽��ʼ�¼.No%Type;
  d_����ʱ�� ���˽��ʼ�¼.�շ�ʱ��%Type;
  n_����id   ���˽��ʼ�¼.Id%Type;

  n_���㿨��� ���ѿ����Ŀ¼.���%Type;
  n_ʱ������   Number(3);
  v_No         ���˽��ʼ�¼.No%Type;
  n_�����id   ҽ�ƿ����.Id%Type;
  v_���㷽ʽ   ����Ԥ����¼.���㷽ʽ%Type;
  v_Temp       Varchar2(500);
  v_Ids        Varchar2(20000);
  x_Templet    Xmltype; --ģ��XML
  v_���ѿ����� Varchar2(20000);
  n_������   ����Ԥ����¼.��Ԥ��%Type;
  n_Step       Number(2);

  v_Err_Msg Varchar2(200);
  Err_Item    Exception;
  Err_Special Exception;
  v_����� �������׼�¼.���%Type;
  n_Number Number(2);

  n_����id   ������ü�¼.Id%Type;
  n_��¼���� ������ü�¼.��¼����%Type;
  v_����no   ������ü�¼.No%Type;
  n_���     ������ü�¼.���%Type;
  n_��¼״̬ ������ü�¼.��¼״̬%Type;
  n_ִ��״̬ ������ü�¼.ִ��״̬%Type;
  n_δ���� ������ü�¼.ʵ�ս��%Type;
  n_���ʽ�� ������ü�¼.ʵ�ս��%Type;
  n_����   ������ü�¼.ʵ�ս��%Type;
  Type t_���ý�����ϸ Is Ref Cursor;
  c_���ý�����ϸ t_���ý�����ϸ;

  Type Ty_Ԥ���� Is Record(
    �˿ʽ   Number(1), --0-�ֽ����˿�;1-����һ�ν��׽ӿ��˿�;2-ת�ʷ�ʽ�˿�(�ݲ�֧��)
    ���ݺ�     ����Ԥ����¼.No%Type,
    ��Ԥ��     ����Ԥ����¼.��Ԥ��%Type,
    ������ˮ�� ����Ԥ����¼.������ˮ��%Type,
    ����˵��   ����Ԥ����¼.����˵��%Type,
    
    Ԥ��id       ����Ԥ����¼.Id%Type,
    ���㷽ʽ     ����Ԥ����¼.���㷽ʽ%Type,
    �����id     ����Ԥ����¼.�����id%Type,
    ����         ����Ԥ����¼.����%Type,
    ԭ������ˮ�� ����Ԥ����¼.������ˮ��%Type,
    ԭ����˵��   ����Ԥ����¼.����˵��%Type,
    ��������id   ����Ԥ����¼.��������id%Type,
    �Ƿ�ת��     Number(1),
    ������չ��Ϣ Xmltype, --�ֽ����˿����չ��Ϣ
    
    �Ƿ��˿� Number(1),
    ��չ��Ϣ Xmltype);
  Type t_Ԥ���� Is Table Of Ty_Ԥ����;
  l_Ԥ���� t_Ԥ����;

  n_��Ԥ�������       Number(1);
  n_��Ԥ�����         ����Ԥ����¼.��Ԥ��%Type;
  n_��Ԥ�����         ����Ԥ����¼.��Ԥ��%Type;
  n_������Ԥ��         Number(1);
  n_�Ƿ����Ʊ��       ����Ԥ����¼.�Ƿ����Ʊ��%Type;
  v_֧�������ں�userid Varchar2(100);
  v_֧����С����userid Varchar2(100);
  v_΢�Ź��ں�openid   Varchar2(100);
  v_΢��С����openid   Varchar2(100);
  n_��Ʊ��־           Number(2);
  v_��������           ����Ʊ��ʹ�ü�¼.����%Type;
  v_��Ʊ���           ����Ʊ��ʹ�ü�¼.����%Type;
  v_��Ʊ����           Varchar2(20);
  n_��Ʊ���           ����Ʊ��ʹ�ü�¼.Ʊ�ݽ��%Type;
  v_Url                ����Ʊ��ʹ�ü�¼.Url����%Type;
  v_Url����            ����Ʊ��ʹ�ü�¼.Url����%Type;

  Procedure ����Ԥ����_List
  (
    Ԥ������Ϣ_In  Xmltype,
    Ԥ����֧��_In  Number,
    Ԥ����_Out     In Out t_Ԥ����,
    ������Ԥ��_Out In Out Number
  ) As
    n_ԭԤ��id ����Ԥ����¼.Id%Type;
    n_ʣ���� ����Ԥ����¼.��Ԥ��%Type;
  
    v_���㷽ʽ     ����Ԥ����¼.���㷽ʽ%Type;
    n_�����id     ����Ԥ����¼.�����id%Type;
    v_����         ����Ԥ����¼.����%Type;
    v_ԭ������ˮ�� ����Ԥ����¼.������ˮ��%Type;
    v_ԭ����˵��   ����Ԥ����¼.����˵��%Type;
    n_��������id   ����Ԥ����¼.��������id%Type;
  
    n_�˿ʽ Number(1);
    x_��չ��Ϣ Xmltype;
  
    I         Number(18);
    v_Err_Msg Varchar2(200);
    Err_Item Exception;
  Begin
    If Ԥ����_Out Is Null Then
      Ԥ����_Out := t_Ԥ����();
    End If;
  
    --�˿ʽ:0-�ֽ����˿�;1-����һ�ν��׽ӿ��˿�;2-ת�ʷ�ʽ�˿�(�ݲ�֧��)
    Select Extractvalue(Value(A), 'CYJLIST/TKFS'), Extract(a.Column_Value, 'CYJLIST/EXPENDLIST')
    Into n_�˿ʽ, x_��չ��Ϣ
    From Table(Xmlsequence(Extract(Ԥ������Ϣ_In, 'CYJLIST'))) A;
  
    For r_Ԥ���� In (Select Extractvalue(b.Column_Value, 'ITEM/DJH') As ���ݺ�, Extractvalue(b.Column_Value, 'ITEM/JE') As ��Ԥ��,
                         Extractvalue(b.Column_Value, 'ITEM/SFTK') As �Ƿ��˿�,
                         Extractvalue(b.Column_Value, 'ITEM/JYLSH') As ������ˮ��,
                         Extractvalue(b.Column_Value, 'ITEM/JYSM') As ����˵��,
                         Extract(b.Column_Value, 'ITEM/EXPENDLIST') As ��չ��Ϣ
                  From Table(Xmlsequence(Extract(Ԥ������Ϣ_In, 'CYJLIST/ITEM'))) B
                  Order By Nvl(�Ƿ��˿�, 0)) Loop
    
      Select Max(Decode(��¼����, 1, ID, 0)), Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0))
      Into n_ԭԤ��id, n_ʣ����
      From ����Ԥ����¼
      Where ��¼���� In (1, 11) And NO = r_Ԥ����.���ݺ�;
      If Nvl(n_ԭԤ��id, 0) = 0 Then
        v_Err_Msg := 'Ԥ�����[' || r_Ԥ����.���ݺ� || ']�����ڣ�����ʧ�ܣ�';
        Raise Err_Item;
      End If;
    
      If Nvl(r_Ԥ����.�Ƿ��˿�, 0) = 0 Then
        If Nvl(n_ʣ����, 0) < Nvl(r_Ԥ����.��Ԥ��, 0) And Nvl(Ԥ����֧��_In, 0) = 1 Then
          v_Err_Msg := 'Ԥ�����[' || r_Ԥ����.���ݺ� || ']���㣬����ʧ�ܣ�';
          Raise Err_Item;
        End If;
      End If;
    
      Begin
        Select ID, ���㷽ʽ, �����id, ����, ������ˮ��, ����˵��, ��������id
        Into n_ԭԤ��id, v_���㷽ʽ, n_�����id, v_����, v_ԭ������ˮ��, v_ԭ����˵��, n_��������id
        From ����Ԥ����¼
        Where ��¼���� = 1 And NO = r_Ԥ����.���ݺ�;
      Exception
        When Others Then
          v_Err_Msg := 'Ԥ�����[' || r_Ԥ����.���ݺ� || ']�����ڣ�����ʧ�ܣ�';
          Raise Err_Item;
      End;
    
      If Nvl(r_Ԥ����.�Ƿ��˿�, 0) = 1 Then
        ������Ԥ��_Out := 1;
      End If;
    
      Ԥ����_Out.Extend();
      I := Ԥ����_Out.Count;
      Ԥ����_Out(I).���ݺ� := r_Ԥ����.���ݺ�;
      Ԥ����_Out(I).��Ԥ�� := r_Ԥ����.��Ԥ��;
      Ԥ����_Out(I).�Ƿ��˿� := r_Ԥ����.�Ƿ��˿�;
      Ԥ����_Out(I).������ˮ�� := r_Ԥ����.������ˮ��;
      Ԥ����_Out(I).����˵�� := r_Ԥ����.����˵��;
    
      Ԥ����_Out(I).Ԥ��id := n_ԭԤ��id;
      Ԥ����_Out(I).���㷽ʽ := v_���㷽ʽ;
      Ԥ����_Out(I).�����id := n_�����id;
      Ԥ����_Out(I).���� := v_����;
      Ԥ����_Out(I).ԭ������ˮ�� := v_ԭ������ˮ��;
      Ԥ����_Out(I).ԭ����˵�� := v_ԭ����˵��;
      Ԥ����_Out(I).��������id := n_��������id;
      If Nvl(n_�˿ʽ, 0) = 2 Then
        Ԥ����_Out(I).�Ƿ�ת�� := 1;
      Else
        Ԥ����_Out(I).�Ƿ�ת�� := 0;
      End If;
      Ԥ����_Out(I).������չ��Ϣ := r_Ԥ����.��չ��Ϣ;
    
      Ԥ����_Out(I).�˿ʽ := n_�˿ʽ;
      Ԥ����_Out(I).��չ��Ϣ := x_��չ��Ϣ;
    End Loop;
  
    If Nvl(������Ԥ��_Out, 0) = 1 And Nvl(n_�˿ʽ, 0) = 2 Then
      v_Err_Msg := 'Ԥ�����ݲ�֧��ת���˿����ʧ�ܣ�';
      Raise Err_Item;
    End If;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  End;

  Procedure �������㽻��_Save
  (
    ����id_In   ����Ԥ����¼.����id%Type,
    �����id_In ����Ԥ����¼.�����id%Type,
    ����_In     ����Ԥ����¼.����%Type,
    ��չ��Ϣ_In Xmltype
  ) As
  Begin
    If ��չ��Ϣ_In Is Null Then
      Return;
    End If;
  
    For c_��չ In (Select Extractvalue(j.Column_Value, 'EXPEND/JYMC') As ����,
                        Extractvalue(j.Column_Value, 'EXPEND/JYLR') As ����
                 From Table(Xmlsequence(Extract(��չ��Ϣ_In, 'EXPENDLIST/EXPEND'))) J) Loop
      Zl_�������㽻��_Insert(�����id_In, 0, ����_In, ����id_In, c_��չ.���� || '|' || c_��չ.����);
    End Loop;
  End;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/ZYID'), To_Number(Extractvalue(Value(A), 'IN/BRID')),
         To_Number(Extractvalue(Value(A), 'IN/JE')), Nvl(To_Number(Extractvalue(Value(A), 'IN/JSLX')), 2),
         Extractvalue(Value(A), 'IN/NO'), To_Number(Extractvalue(Value(A), 'IN/JZKNO')),
         To_Date(Extractvalue(Value(A), 'IN/JZSJ'), 'yyyy-mm-dd hh24:mi:ss'), Extractvalue(Value(A), 'IN/SFZH'),
         Extractvalue(Value(A), 'IN/XM'), Extractvalue(Value(A), 'IN/JSMS'), Extractvalue(Value(A), 'IN/CZLX'),
         Extractvalue(Value(A), 'IN/JZID'), Extractvalue(Value(A), 'IN/ZFBZH'), Extractvalue(Value(A), 'IN/ZFBXCY'),
         Extractvalue(Value(A), 'IN/WXGZHID'), Extractvalue(Value(A), 'IN/WXXCXID')
  Into n_��ҳid, n_����id, n_�����ܶ�, n_��������, v_���ݺ�, v_���￨���ݺ�, d_����ʱ��, v_���֤��, v_����, n_����ģʽ, n_��������, n_����id, v_֧�������ں�userid,
       v_֧����С����userid, v_΢�Ź��ں�openid, v_΢��С����openid
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If n_�������� = 1 And Nvl(n_����id, 0) = 0 And Not v_���֤�� Is Null And Not v_���� Is Null Then
    n_����id := Zl_Third_Getpatiid(v_���֤��, v_����);
  End If;

  --0.��ؼ��
  If Nvl(n_����id, 0) = 0 Then
    v_Err_Msg := '������Чʶ�������,���������!';
    Raise Err_Item;
  End If;

  If Not v_֧�������ں�userid Is Null Then
    Zl_������Ϣ�ӱ�_Update(n_����id, Upper('֧�������ں�UserID'), v_֧�������ں�userid);
  End If;

  If Not v_֧����С����userid Is Null Then
    Zl_������Ϣ�ӱ�_Update(n_����id, Upper('֧����С����UserID'), v_֧�������ں�userid);
  End If;

  If Not v_΢�Ź��ں�openid Is Null Then
    Zl_������Ϣ�ӱ�_Update(n_����id, Upper('΢�Ź��ں�OpenID'), v_֧�������ں�userid);
  End If;

  If Not v_΢��С����openid Is Null Then
    Zl_������Ϣ�ӱ�_Update(n_����id, Upper('΢��С����OpenID'), v_֧�������ں�userid);
  End If;

  If Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) <> 0 Then
    If Nvl(n_����id, 0) = 0 Then
      v_Err_Msg := 'û��ָ����صĽ������ݣ�';
      Raise Err_Item;
    End If;
  
    Begin
      Select �տ�ʱ��, ��������id, �����id
      Into d_����ʱ��, n_��������id, n_�����id
      From ����Ԥ����¼
      Where ����id = n_����id And Nvl(У�Ա�־, 0) = 1 And �����id Is Not Null And Rownum < 2;
    Exception
      When Others Then
        v_Err_Msg := 'û���ҵ�ָ������ؽ������ݣ������ѱ�����';
        Raise Err_Item;
    End;
  End If;

  v_����Ա���� := Zl_����Ա��Ϣ(1);
  v_����Ա���� := Zl_����Ա��Ϣ(2);

  --��1���첽ģʽ���˽���
  If Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) = 2 Then
    --ɾ����������
    Zl_���˽��ʽ���_Delete(n_����id, n_�����id, n_��������id);
    --����ԭ����
    Begin
      Select NO Into v_���ʵ��� From ���˽��ʼ�¼ Where ID = n_����id And Nvl(����״̬, 0) = 1 And Rownum < 2;
    Exception
      When Others Then
        v_Err_Msg := 'û���ҵ�ԭ�������ݣ������ѱ�����';
        Raise Err_Item;
    End;
  
    d_����ʱ�� := Sysdate;
    Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
    Zl_���˽��ʼ�¼_Cancel(v_���ʵ���, n_����id, v_����Ա����, v_����Ա����, d_����ʱ��);
    Zl_���˽�������_Modify(0, n_����id, n_����id, Null, Null, Null, Null, Null, Null, Null, Null, Null, v_����Ա����, v_����Ա����, d_����ʱ��,
                     Null, 1);
  
    v_Temp := '<CZSJ>' || To_Char(d_����ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<JZID>' || n_����id || '</JZID>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<KPBZ>' || 0 || '</KPBZ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<URL>' || '' || '</URL>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<NETURL>' || '' || '</NETURL>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<FPTT>' || v_���� || '</FPTT>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<FPH>' || '' || '</FPH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<FPJE>' || '' || '</FPJE>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<KPRQ>' || '' || '</KPRQ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    Xml_Out := x_Templet;
    Return;
  End If;

  --��2�����첽ģʽ���첽ģʽ��ʼ����
  If Nvl(n_����ģʽ, 0) = 0 Or Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) = 0 Then
    --��2.1������������֧������
    For c_���׼�¼ In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As ���㿨���,
                          Extractvalue(b.Column_Value, '/JS/JSFS') As ���㷽ʽ,
                          Extractvalue(b.Column_Value, '/JS/JYLSH') As ������ˮ��,
                          Extractvalue(b.Column_Value, '/JS/JSKH') As ���㿨��, Extractvalue(b.Column_Value, '/JS/ZY') As ժҪ,
                          Extractvalue(b.Column_Value, '/JS/SFCYJ') As �Ƿ��Ԥ��,
                          Extractvalue(b.Column_Value, '/JS/SFXFK') As �Ƿ����ѿ�
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
    
      If Not (c_���׼�¼.���㿨��� Is Null Or Nvl(c_���׼�¼.�Ƿ����ѿ�, '0') = '1' Or Nvl(c_���׼�¼.�Ƿ��Ԥ��, 0) = 1) Then
        Select Decode(Translate(Nvl(c_���׼�¼.���㿨���, 'abcd'), '#1234567890', '#'), Null, 1, 0)
        Into n_Number
        From Dual;
      
        If Nvl(n_Number, 0) = 1 Then
          Select Max(����) Into v_����� From ҽ�ƿ���� Where ID = To_Number(c_���׼�¼.���㿨���);
        Else
          Select Max(����) Into v_����� From ҽ�ƿ���� Where ���� = c_���׼�¼.���㿨���;
        End If;
        If v_����� Is Null Then
          v_Err_Msg := '��֧�ֵĽ��㷽ʽ,���飡';
          Raise Err_Item;
        End If;
      
        --����һ�����㷽ʽ�ż�齻����
        n_Step := Nvl(n_Step, 0) + 1;
        If Zl_Fun_�������׼�¼_Locked(v_�����, c_���׼�¼.������ˮ��, c_���׼�¼.���㿨��, c_���׼�¼.ժҪ, 2) = 0 And n_Step = 1 Then
          v_Err_Msg := '������ˮ��Ϊ:' || c_���׼�¼.������ˮ�� || '�Ľ������ڽ����У��������ٴ��ύ�˽���!';
          Raise Err_Special;
        End If;
      End If;
    End Loop;
  
    --��2.2�����ý��ʼ�¼
    Select Nvl(zl_GetSysParameter('���ʷ���ʱ��', 1137), 0) Into n_ʱ������ From Dual;
    If n_�������� = 2 Then
      Open c_���ý�����ϸ For
        Select Max(Decode(����id, Null, ID, Null)) As ID, Mod(��¼����, 10) As ��¼����, NO, ���, ��¼״̬, ִ��״̬,
               Trunc(Min(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ��Сʱ��, Trunc(Max(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ���ʱ��,
               Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) As ���, Sum(Nvl(���ʽ��, 0)) As ���ʽ��
        From סԺ���ü�¼
        Where ����id = n_����id And ��¼״̬ <> 0 And ��ҳid = n_��ҳid And ���ʷ��� = 1
        Group By Mod(��¼����, 10), NO, ���, ��¼״̬, ִ��״̬
        Having(Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) <> 0) Or (Sum(Nvl(ʵ�ս��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Sum(Nvl(���ʽ��, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(���ʽ��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Mod(Count(*), 2) = 0
        Order By NO, ���;
    Else
      If v_���ݺ� Is Null And v_���￨���ݺ� Is Null Then
        Open c_���ý�����ϸ For
          Select Min(Decode(����id, Null, ID, Null)) As ID, Mod(��¼����, 10) As ��¼����, NO, ���, ��¼״̬, ִ��״̬,
                 Trunc(Min(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ��Сʱ��, Trunc(Max(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ���ʱ��,
                 Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) As ���, Sum(Nvl(���ʽ��, 0)) As ���ʽ��
          From ������ü�¼
          Where ����id = n_����id And ��¼״̬ <> 0 And ���ʷ��� = 1
          Group By Mod(��¼����, 10), NO, ���, ��¼״̬, ִ��״̬
          Having(Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) <> 0) Or (Sum(Nvl(ʵ�ս��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Sum(Nvl(���ʽ��, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(���ʽ��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Mod(Count(*), 2) = 0
          Union All
          Select Min(Decode(����id, Null, ID, Null)) As ID, Mod(��¼����, 10) As ��¼����, NO, ���, ��¼״̬, ִ��״̬,
                 Trunc(Min(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ��Сʱ��, Trunc(Max(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ���ʱ��,
                 Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) As ���, Sum(Nvl(���ʽ��, 0)) As ���ʽ��
          From סԺ���ü�¼
          Where ����id = n_����id And ��¼״̬ <> 0 And ���ʷ��� = 1 And Mod(��¼����, 10) = 5
          Group By Mod(��¼����, 10), NO, ���, ��¼״̬, ִ��״̬
          Having(Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) <> 0) Or (Sum(Nvl(ʵ�ս��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Sum(Nvl(���ʽ��, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(���ʽ��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Mod(Count(*), 2) = 0
          Order By NO, ���;
      Elsif v_���ݺ� Is Not Null And v_���￨���ݺ� Is Not Null Then
        Open c_���ý�����ϸ For
          Select Min(Decode(����id, Null, ID, Null)) As ID, Mod(��¼����, 10) As ��¼����, NO, ���, ��¼״̬, ִ��״̬,
                 Trunc(Min(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ��Сʱ��, Trunc(Max(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ���ʱ��,
                 Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) As ���, Sum(Nvl(���ʽ��, 0)) As ���ʽ��
          From ������ü�¼
          Where ����id + 0 = n_����id And ��¼״̬ <> 0 And Mod(��¼����, 10) = 2 And ���ʷ��� = 1 And
                NO In (Select /*+cardinality(b,10) */
                        Column_Value
                       From Table(f_Str2List(v_���ݺ�)) B)
          Group By Mod(��¼����, 10), NO, ���, ��¼״̬, ִ��״̬
          Having(Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) <> 0) Or (Sum(Nvl(ʵ�ս��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Sum(Nvl(���ʽ��, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(���ʽ��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Mod(Count(*), 2) = 0
          Union All
          Select Min(Decode(����id, Null, ID, Null)) As ID, Mod(��¼����, 10) As ��¼����, NO, ���, ��¼״̬, ִ��״̬,
                 Trunc(Min(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ��Сʱ��, Trunc(Max(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ���ʱ��,
                 Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) As ���, Sum(Nvl(���ʽ��, 0)) As ���ʽ��
          From סԺ���ü�¼
          Where ����id + 0 = n_����id And ��¼״̬ <> 0 And ���ʷ��� = 1 And Mod(��¼����, 10) = 5 And
                NO In (Select /*+cardinality(b,10) */
                        Column_Value
                       From Table(f_Str2List(v_���￨���ݺ�)) B)
          Group By Mod(��¼����, 10), NO, ���, ��¼״̬, ִ��״̬
          Having(Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) <> 0) Or (Sum(Nvl(ʵ�ս��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Sum(Nvl(���ʽ��, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(���ʽ��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Mod(Count(*), 2) = 0
          Order By ��¼����, NO, ���;
      Elsif v_���ݺ� Is Not Null Then
        Open c_���ý�����ϸ For
          Select Min(Decode(����id, Null, ID, Null)) As ID, Mod(��¼����, 10) As ��¼����, NO, ���, ��¼״̬, ִ��״̬,
                 Trunc(Min(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ��Сʱ��, Trunc(Max(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ���ʱ��,
                 Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) As ���, Sum(Nvl(���ʽ��, 0)) As ���ʽ��
          From ������ü�¼
          Where ����id + 0 = n_����id And ��¼״̬ <> 0 And Mod(��¼����, 10) = 2 And ���ʷ��� = 1 And
                NO In (Select /*+cardinality(b,10) */
                        Column_Value
                       From Table(f_Str2List(v_���ݺ�)) B)
          Group By Mod(��¼����, 10), NO, ���, ��¼״̬, ִ��״̬
          Having(Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) <> 0) Or (Sum(Nvl(ʵ�ս��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Sum(Nvl(���ʽ��, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(���ʽ��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Mod(Count(*), 2) = 0
          Order By NO, ���;
      Else
        Open c_���ý�����ϸ For
          Select Min(Decode(����id, Null, ID, Null)) As ID, Mod(��¼����, 10) As ��¼����, NO, ���, ��¼״̬, ִ��״̬,
                 Trunc(Min(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ��Сʱ��, Trunc(Max(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ���ʱ��,
                 Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) As ���, Sum(Nvl(���ʽ��, 0)) As ���ʽ��
          From סԺ���ü�¼
          Where ����id + 0 = n_����id And ��¼״̬ <> 0 And ���ʷ��� = 1 And Mod(��¼����, 10) = 5 And
                NO In (Select /*+cardinality(b,10) */
                        Column_Value
                       From Table(f_Str2List(v_���￨���ݺ�)) B)
          Group By Mod(��¼����, 10), NO, ���, ��¼״̬, ִ��״̬
          Having(Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) <> 0) Or (Sum(Nvl(ʵ�ս��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Sum(Nvl(���ʽ��, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(���ʽ��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Mod(Count(*), 2) = 0
          Order By NO, ���;
      End If;
    End If;
  
    Select ���˽��ʼ�¼_Id.Nextval, Sysdate, Nextno(15) Into n_����id, d_����ʱ��, v_No From Dual;
    n_�����ʽ�� := 0;
    Loop
      Fetch c_���ý�����ϸ
        Into n_����id, n_��¼����, v_����no, n_���, n_��¼״̬, n_ִ��״̬, d_��С����, d_�������, n_δ����, n_���ʽ��;
      Exit When c_���ý�����ϸ%NotFound;
    
      n_�����ʽ�� := n_�����ʽ�� + Nvl(n_δ����, 0);
      If d_��ʼ���� Is Null Then
        d_��ʼ���� := d_��С����;
      Elsif d_��ʼ���� > d_��С���� Then
        d_��ʼ���� := d_��С����;
      End If;
      If d_�������� Is Null Then
        d_�������� := d_�������;
      Elsif d_�������� < d_������� Then
        d_�������� := d_�������;
      End If;
    
      If Nvl(n_���ʽ��, 0) = 0 Then
        If n_����id Is Not Null Then
          If Length(v_Ids || ',' || n_����id) > 4000 Then
            v_Ids := Substr(v_Ids, 2);
            Zl_���ʷ��ü�¼_Batch(v_Ids, n_����id, n_����id);
            v_Ids := '';
          End If;
          v_Ids := v_Ids || ',' || n_����id;
        Else
          Zl_���ʷ��ü�¼_Insert(0, v_����no, n_��¼����, n_��¼״̬, n_ִ��״̬, n_���, n_δ����, n_����id);
        End If;
      Else
        Zl_���ʷ��ü�¼_Insert(0, v_����no, n_��¼����, n_��¼״̬, n_ִ��״̬, n_���, n_δ����, n_����id);
      End If;
    End Loop;
  
    If v_Ids Is Not Null Then
      v_Ids := Substr(v_Ids, 2);
      Zl_���ʷ��ü�¼_Batch(v_Ids, n_����id, n_����id);
    End If;
  
    If Round(n_�����ʽ��, 6) <> Nvl(n_�����ܶ�, 0) Then
      v_Err_Msg := '����Ľ��ʽ����ʵ�ʽ��ʽ���,���������!';
      Raise Err_Item;
    End If;
  
    Zl_���˽��ʼ�¼_Insert(n_����id, v_No, n_����id, d_����ʱ��, d_��ʼ����, d_��������, 0, 0, n_��ҳid, Null, n_��������, Null, n_��������, 1, n_��ҳid,
                     n_�����ܶ�);
  
    --��2.3����������Ԥ�ȱ���,������֧���ʹ����˿��Ԥ����֧��
    n_��Ԥ������� := 0;
    n_������Ԥ��   := 0;
    l_Ԥ����       := t_Ԥ����();
    For r_���㷽ʽ In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As ���㿨���,
                          Extractvalue(b.Column_Value, '/JS/JSKH') As ���㿨��,
                          Extractvalue(b.Column_Value, '/JS/JSFS') As ���㷽ʽ,
                          Extractvalue(b.Column_Value, '/JS/JSJE') As ������,
                          Extractvalue(b.Column_Value, '/JS/JYLSH') As ������ˮ��,
                          Extractvalue(b.Column_Value, '/JS/JYSM') As ����˵��, Extractvalue(b.Column_Value, '/JS/ZY') As ժҪ,
                          Extractvalue(b.Column_Value, '/JS/SFXFK') As �Ƿ����ѿ�,
                          Extractvalue(b.Column_Value, '/JS/SFCYJ') As �Ƿ��Ԥ��,
                          Extract(b.Column_Value, '/JS/CYJLIST') As ��Ԥ���б�
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
    
      If Nvl(r_���㷽ʽ.�Ƿ��Ԥ��, 0) = 0 Then
        n_��Ԥ������� := 1;
        n_�����id     := Null;
        If r_���㷽ʽ.���㿨��� Is Not Null And Nvl(r_���㷽ʽ.�Ƿ����ѿ�, 0) = 0 Then
          Select Decode(Translate(Nvl(r_���㷽ʽ.���㿨���, 'abcd'), '#1234567890', '#'), Null, 1, 0)
          Into n_Number
          From Dual;
        
          If Nvl(n_Number, 0) = 1 Then
            Select Max(ID) Into n_�����id From ҽ�ƿ���� Where ID = r_���㷽ʽ.���㿨��� And Nvl(�Ƿ�����, 0) = 1;
          Else
            Select Max(ID) Into n_�����id From ҽ�ƿ���� Where ���� = r_���㷽ʽ.���㿨��� And Nvl(�Ƿ�����, 0) = 1;
          End If;
        
          If n_�����id Is Null Then
            v_Err_Msg := 'δ�ҵ���Ӧ��ҽ�ƿ���Ϣ!';
            Raise Err_Item;
          End If;
        
          Select Max(��������id)
          Into n_��������id
          From ����Ԥ����¼
          Where ����id = n_����id And �����id = n_�����id And Rownum < 2;
          If Nvl(n_��������id, 0) = 0 Then
            Select ����Ԥ����¼_Id.Nextval Into n_��������id From Dual;
            n_Ԥ��id := n_��������id;
          Else
            n_Ԥ��id := Null;
          End If;
        
          v_���㷽ʽ := r_���㷽ʽ.���㷽ʽ || '|' || r_���㷽ʽ.������ || '|';
          Zl_���˽��ʽ���_Modify(1, n_����id, n_����id, v_���㷽ʽ, Null, 0, n_�����id, r_���㷽ʽ.���㿨��, r_���㷽ʽ.������ˮ��, r_���㷽ʽ.����˵��, 0, 0, 0,
                           n_��������, Null, v_����Ա����, v_����Ա����, d_����ʱ��, Null, 0, 1, n_Ԥ��id, n_��������id);
        End If;
      
      Elsif r_���㷽ʽ.��Ԥ���б� Is Not Null Then
        --��*����ָ�����ݽ���Ԥ����֧��
        ����Ԥ����_List(r_���㷽ʽ.��Ԥ���б�, 1, l_Ԥ����, n_������Ԥ��);
      End If;
    End Loop;
  
    --��*����ָ�����ݽ���Ԥ����֧��
    If l_Ԥ����.Count > 0 And Nvl(n_������Ԥ��, 0) = 1 Then
      --  1.���������Ԥ�����ֻ��ʹ��Ԥ������н���
      If Nvl(n_��Ԥ�������, 0) = 1 Then
        v_Err_Msg := '���ڶ�Ԥ��������˿�ʱ��ȫ�����ʽ�����ʹ��Ԥ�������֧����';
        Raise Err_Item;
      End If;
    
      --�ȱ���Ԥ��������
      n_��Ԥ����� := 0;
      n_��Ԥ����� := 0;
      For I In 1 .. l_Ԥ����.Count Loop
        If Nvl(l_Ԥ����(I).�Ƿ��˿�, 0) = 0 Then
          --��Ԥ����  
          Zl_����Ԥ����¼_Insert(l_Ԥ����(I).Ԥ��id, l_Ԥ����(I).���ݺ�, 1, l_Ԥ����(I).��Ԥ��, n_����id, n_����id, v_����Ա����, v_����Ա����, d_����ʱ��);
        
          n_��Ԥ����� := n_��Ԥ����� + Nvl(l_Ԥ����(I).��Ԥ��, 0);
        Else
          --��Ԥ����  
          Zl_�����˿���Ϣ_Insert(n_����id, l_Ԥ����(I).Ԥ��id, l_Ԥ����(I).��Ԥ��, l_Ԥ����(I).����, Null, Null, 0, 1, l_Ԥ����(I).�Ƿ�ת��,
                           l_Ԥ����(I).�����id, l_Ԥ����(I).ԭ������ˮ��, l_Ԥ����(I).ԭ����˵��);
          Zl_���˽��ʽ���_Modify(1, n_����id, n_����id, l_Ԥ����(I).���㷽ʽ || '|' || -1 * l_Ԥ����(I).��Ԥ�� || '| | ', Null, Null,
                           l_Ԥ����(I).�����id, l_Ԥ����(I).����, l_Ԥ����(I).ԭ������ˮ��, l_Ԥ����(I).ԭ����˵��, Null, Null, Null, n_��������, Null,
                           v_����Ա����, v_����Ա����, d_����ʱ��, Null, 0, 1, Null, l_Ԥ����(I).��������id, 0, Nvl(l_Ԥ����(I).�˿ʽ, 0) + 1);
        
          n_��Ԥ����� := n_��Ԥ����� + Nvl(l_Ԥ����(I).��Ԥ��, 0);
        End If;
      End Loop;
    
      --  2.���˱��г壻�絥��A001��Ԥ������1000�����ν���800������������¼�ٳ�1000������200��Sum(����-�˽��)=���ʽ��
      --  ˵���������������
      If Abs(Nvl(n_�����ܶ�, 0) - (Nvl(n_��Ԥ�����, 0) - Nvl(n_��Ԥ�����, 0))) >= 1.00 Then
        v_Err_Msg := 'Ԥ����֧���������ʽ���ȣ���������㣡';
        Raise Err_Item;
      End If;
    End If;
  End If;

  --��3���첽ģʽ��ʼ�����Ѵ����꣬���ؽ��
  If Nvl(n_����ģʽ, 0) = 1 And Nvl(n_��������, 0) = 0 Then
    v_Temp := '<CZSJ>' || To_Char(d_����ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<JZID>' || n_����id || '</JZID>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<KPBZ>' || 0 || '</KPBZ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<URL>' || '' || '</URL>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<NETURL>' || '' || '</NETURL>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<FPTT>' || v_���� || '</FPTT>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<FPH>' || '' || '</FPH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<FPJE>' || '' || '</FPJE>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<KPRQ>' || '' || '</KPRQ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    Xml_Out := x_Templet;
    Return;
  End If;

  --��4�����첽ģʽ���첽ģʽ��ɽ��ף�������������
  n_������     := 0;
  n_ɾ��ԭ����   := 1;
  n_���ʽ��     := 0;
  n_��Ԥ������� := 0;
  n_������Ԥ��   := 0;
  l_Ԥ����       := t_Ԥ����();
  For r_���㷽ʽ In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As ���㿨���,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As ���㿨��,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As ���㷽ʽ,
                        Extractvalue(b.Column_Value, '/JS/JSJE') As ������,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As ������ˮ��,
                        Extractvalue(b.Column_Value, '/JS/JYSM') As ����˵��, Extractvalue(b.Column_Value, '/JS/ZY') As ժҪ,
                        Extractvalue(b.Column_Value, '/JS/SFXFK') As �Ƿ����ѿ�,
                        Extractvalue(b.Column_Value, '/JS/SFCYJ') As �Ƿ��Ԥ��,
                        Extract(b.Column_Value, '/JS/CYJLIST') As ��Ԥ���б�,
                        Extract(b.Column_Value, '/JS/EXPENDLIST') As Expend
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
  
    v_�����   := r_���㷽ʽ.���㷽ʽ;
    n_���ʽ�� := n_���ʽ�� + Nvl(r_���㷽ʽ.������, 0);
    If Nvl(r_���㷽ʽ.�Ƿ��Ԥ��, 0) = 0 Then
      n_��Ԥ������� := 1;
      n_�����id     := Null;
      If r_���㷽ʽ.���㿨��� Is Not Null Then
        Select Decode(Translate(Nvl(r_���㷽ʽ.���㿨���, 'abcd'), '#1234567890', '#'), Null, 1, 0)
        Into n_Number
        From Dual;
      
        If Nvl(r_���㷽ʽ.�Ƿ����ѿ�, 0) = 1 Then
          If Nvl(n_Number, 0) = 1 Then
            Select Max(���), Max(����)
            Into n_���㿨���, v_�����
            From ���ѿ����Ŀ¼
            Where ��� = r_���㷽ʽ.���㿨��� And Nvl(����, 0) = 1;
          Else
            Select Max(���), Max(����)
            Into n_���㿨���, v_�����
            From ���ѿ����Ŀ¼
            Where ���� = r_���㷽ʽ.���㿨��� And Nvl(����, 0) = 1;
          End If;
        
          If n_���㿨��� Is Null Then
            v_Err_Msg := 'δ�ҵ���Ӧ�����ѿ���Ϣ';
            Raise Err_Item;
          End If;
        Else
          If Nvl(n_Number, 0) = 1 Then
            Select Max(ID), Max(����)
            Into n_�����id, v_�����
            From ҽ�ƿ����
            Where ID = r_���㷽ʽ.���㿨��� And Nvl(�Ƿ�����, 0) = 1;
          Else
            Select Max(ID), Max(����)
            Into n_�����id, v_�����
            From ҽ�ƿ����
            Where ���� = r_���㷽ʽ.���㿨��� And Nvl(�Ƿ�����, 0) = 1;
          End If;
        
          If n_�����id Is Null Then
            v_Err_Msg := 'δ�ҵ���Ӧ��ҽ�ƿ���Ϣ!';
            Raise Err_Item;
          End If;
        End If;
      End If;
    
      If n_�����id Is Not Null Then
        --������
        Select Max(��������id)
        Into n_��������id
        From ����Ԥ����¼
        Where ����id = n_����id And �����id = n_�����id And Rownum < 2;
        If n_ɾ��ԭ���� = 1 Then
          n_Ԥ��id := n_��������id;
        Else
          n_Ԥ��id := Null;
        End If;
      
        v_���㷽ʽ := r_���㷽ʽ.���㷽ʽ || '|' || r_���㷽ʽ.������ || '|';
        Zl_���˽��ʽ���_Modify(1, n_����id, n_����id, v_���㷽ʽ, Null, 0, n_�����id, r_���㷽ʽ.���㿨��, r_���㷽ʽ.������ˮ��, r_���㷽ʽ.����˵��, 0, 0, 0,
                         n_��������, Null, v_����Ա����, v_����Ա����, d_����ʱ��, Null, 0, 1, n_Ԥ��id, n_��������id, n_ɾ��ԭ����);
        n_ɾ��ԭ���� := 0;
      
        �������㽻��_Save(n_����id, n_�����id, r_���㷽ʽ.���㿨��, r_���㷽ʽ.Expend);
      Else
        If n_���㿨��� Is Not Null Then
          --���ѿ�
          v_���ѿ����� := Nvl(v_���ѿ�����, '') || '||' || n_���㿨��� || '|' || r_���㷽ʽ.���㿨�� || '|0|' || r_���㷽ʽ.������;
        Else
          --��������
          v_���㷽ʽ := r_���㷽ʽ.���㷽ʽ || '|' || r_���㷽ʽ.������ || '||';
          Zl_���˽��ʽ���_Modify(0, n_����id, n_����id, v_���㷽ʽ, Null, 0, n_�����id, r_���㷽ʽ.���㿨��, r_���㷽ʽ.������ˮ��, r_���㷽ʽ.����˵��, 0, 0, 0,
                           n_��������, Null, v_����Ա����, v_����Ա����, d_����ʱ��, Null, 0);
        End If;
      End If;
    
      n_������ := n_������ + Nvl(r_���㷽ʽ.������, 0);
    Elsif r_���㷽ʽ.��Ԥ���б� Is Null Then
      --��**����Ԥ��,ĿǰĬ��ȫ��
      n_��Ԥ����� := r_���㷽ʽ.������;
      Zl_���˽��ʽ���_Modify(0, n_����id, n_����id, Null, n_��Ԥ�����, 0, Null, Null, Null, Null, 0, 0, 0, n_��������, Null, v_����Ա����,
                       v_����Ա����, d_����ʱ��, Null, 0);
    
      n_������ := n_������ + Nvl(r_���㷽ʽ.������, 0);
    Else
      --��*����ָ�����ݽ���Ԥ����֧��
      ����Ԥ����_List(r_���㷽ʽ.��Ԥ���б�, 0, l_Ԥ����, n_������Ԥ��);
    End If;
  
    Update �������׼�¼
    Set ҵ�����id = n_����id
    Where ��ˮ�� = Nvl(r_���㷽ʽ.������ˮ��, '-') And ��� = v_����� And ҵ������ = 2;
  End Loop;

  --���ѿ�����
  If v_���ѿ����� Is Not Null Then
    v_���ѿ����� := Substr(v_���ѿ�����, 3);
    Zl_���˽��ʽ���_Modify(3, n_����id, n_����id, v_���ѿ�����, Null, 0, Null, Null, Null, Null, 0, 0, 0, n_��������, Null, v_����Ա����,
                     v_����Ա����, d_����ʱ��, Null, 0);
  End If;

  --��*����ָ�����ݽ���Ԥ����֧��
  If l_Ԥ����.Count > 0 Then
    n_��Ԥ����� := 0;
    n_��Ԥ����� := 0;
    If Nvl(n_������Ԥ��, 0) = 0 Then
      --����Ԥ�����������
      For I In 1 .. l_Ԥ����.Count Loop
        Zl_����Ԥ����¼_Insert(l_Ԥ����(I).Ԥ��id, l_Ԥ����(I).���ݺ�, 1, l_Ԥ����(I).��Ԥ��, n_����id, n_����id, v_����Ա����, v_����Ա����, d_����ʱ��);
      
        n_��Ԥ����� := n_��Ԥ����� + Nvl(l_Ԥ����(I).��Ԥ��, 0);
      End Loop;
    Else
      --����Ԥ�����������
      --  1.���������Ԥ�����ֻ��ʹ��Ԥ������н���
      If Nvl(n_��Ԥ�������, 0) = 1 Then
        v_Err_Msg := '���ڶ�Ԥ��������˿�ʱ������ȫ�����ʽ�����ʹ��Ԥ�������֧����';
        Raise Err_Item;
      End If;
    
      For I In 1 .. l_Ԥ����.Count Loop
        If Nvl(l_Ԥ����(I).�Ƿ��˿�, 0) = 0 Then
          n_��Ԥ����� := n_��Ԥ����� + Nvl(l_Ԥ����(I).��Ԥ��, 0);
        Else
          Zl_�����˿���Ϣ_Insert(n_����id, l_Ԥ����(I).Ԥ��id, l_Ԥ����(I).��Ԥ��, l_Ԥ����(I).����, l_Ԥ����(I).������ˮ��, l_Ԥ����(I).����˵��, 1, 0,
                           l_Ԥ����(I).�Ƿ�ת��, l_Ԥ����(I).�����id, l_Ԥ����(I).ԭ������ˮ��, l_Ԥ����(I).ԭ����˵��);
          Zl_���˽��ʽ���_Modify(1, n_����id, n_����id, l_Ԥ����(I).���㷽ʽ || '|' || -1 * l_Ԥ����(I).��Ԥ�� || '| | ', Null, Null,
                           l_Ԥ����(I).�����id, l_Ԥ����(I).����, l_Ԥ����(I).ԭ������ˮ��, l_Ԥ����(I).ԭ����˵��, Null, Null, Null, n_��������, Null,
                           v_����Ա����, v_����Ա����, d_����ʱ��, Null, 0, 2, Null, l_Ԥ����(I).��������id, 1, Nvl(l_Ԥ����(I).�˿ʽ, 0) + 1);
        
          --������չ������Ϣ
          If Nvl(l_Ԥ����(I).�˿ʽ, 0) = 0 Then
            �������㽻��_Save(n_����id, l_Ԥ����(I).�����id, l_Ԥ����(I).����, l_Ԥ����(I).������չ��Ϣ);
          End If;
        
          n_��Ԥ����� := n_��Ԥ����� + Nvl(l_Ԥ����(I).��Ԥ��, 0);
        End If;
      End Loop;
    
      --  2.���˱��г壻�絥��A001��Ԥ������1000�����ν���800������������¼�ٳ�1000������200��Sum(����-�˽��)=���ʽ��
      If Abs(Nvl(n_�����ܶ�, 0) - (Nvl(n_��Ԥ�����, 0) - Nvl(n_��Ԥ�����, 0))) > 1.00 Then
        v_Err_Msg := 'Ԥ����֧���������ʽ���ȣ���������㣡';
        Raise Err_Item;
      End If;
    End If;
  
    If Nvl(l_Ԥ����(1).�˿ʽ, 0) <> 0 Then
      For I In 1 .. l_Ԥ����.Count Loop
        If Nvl(l_Ԥ����(I).�Ƿ��˿�, 0) = 1 Then
          �������㽻��_Save(n_����id, l_Ԥ����(I).�����id, l_Ԥ����(I).����, l_Ԥ����(I).��չ��Ϣ);
          Exit;
        End If;
      End Loop;
    End If;
  
    n_������ := n_������ + (Nvl(n_��Ԥ�����, 0) - Nvl(n_��Ԥ�����, 0));
  End If;

  n_���� := Round(Nvl(n_�����ܶ�, 0) - Nvl(n_������, 0), 6);
  If Abs(Nvl(n_����, 0)) > 1 Then
    v_Err_Msg := '���������������1.00��С��-1.00Ԫ,��������ʲ���,����!';
    Raise Err_Item;
  End If;

  --��5����ɽ��㣬���ؽ��

  --����Ʊ�ݴ���  
  n_�Ƿ����Ʊ�� := b_Einvoice_Request.Einvoice_Start(3, Null, n_��������);
  Zl_���˽��ʽ���_Modify(0, n_����id, n_����id, '', Null, 0, Null, Null, Null, Null, 0, 0, n_����, n_��������, Null, v_����Ա����, v_����Ա����,
                   d_����ʱ��, Null, 1, 2, Null, Null, 0, Null, 0, n_�Ƿ����Ʊ��);
  --��Ҫ���ߵ���Ʊ��
  If Nvl(n_�Ƿ����Ʊ��, 0) = 1 Then
    If b_Einvoice_Request.Einvoice_Create(3, n_����id, Null, v_Err_Msg) = 0 Then
      --����Ʊ�ݿ��߳ɹ�
      Raise Err_Item;
    End If;
  
    Select Max(1), Max(����), Max(����), Max(To_Char(����ʱ��, 'yyyy-mm-dd')), Max(Url����), Max(Url����), Max(Ʊ�ݽ��)
    Into n_��Ʊ��־, v_��������, v_��Ʊ���, v_��Ʊ����, v_Url, v_Url����, n_��Ʊ���
    From ����Ʊ��ʹ�ü�¼
    Where ����id = n_����id And Ʊ�� = 3 And ��¼״̬ = 1;
  
    If v_�������� Is Not Null Then
      v_���� := v_��������;
    End If;
  End If;
  v_Temp := '<CZSJ>' || To_Char(d_����ʱ��, 'YYYY-MM-DD hh23:mi:ss') || '</CZSJ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<JZID>' || n_����id || '</JZID>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<KPBZ>' || Nvl(n_��Ʊ��־, 0) || '</KPBZ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<URL>' || Nvl(v_Url, '') || '</URL>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<NETURL>' || Nvl(v_Url����, '') || '</NETURL>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<FPTT>' || v_���� || '</FPTT>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<FPH>' || v_��Ʊ��� || '</FPH>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<FPJE>' || Nvl(n_��Ʊ���, 0) || '</FPJE>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<KPRQ>' || v_��Ʊ���� || '</KPRQ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Err_Special Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20105, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Settlement;
/

Create Or Replace Procedure Zl_����Ʊ�����_Update(���_In ����Ʊ�����.���%Type) As
  --���ܣ��޸ĵ���Ʊ���������á�ͣ��
  v_Error Varchar2(255);
  Err_Item Exception;
Begin

  If ���_In Is Null Then
    v_Error := '�������Ʊ�����ı��,���飡';
    Raise Err_Item;
  End If;

  --��ͣ��ԭ����Ʊ�ݽӿ�
  Update ����Ʊ����� Set �Ƿ����� = Null Where Nvl(�Ƿ�����, 0) = 1;

  --�������ֵ���Ʊ�ݽӿ�
  Update ����Ʊ����� Set �Ƿ����� = 1 Where ��� = ���_In;
  If Sql%NotFound Then
    v_Error := '������δ�ҵ���Ӧ�ĵ���Ʊ���������,���飡';
    Raise Err_Item;
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ʊ�����_Update;
/
Create Or Replace Procedure Zl_��Լ��λ_Insert
(
  Id_In           In ��Լ��λ.Id%Type,
  �ϼ�id_In       In ��Լ��λ.�ϼ�id%Type,
  ����_In         In ��Լ��λ.����%Type,
  ����_In         In ��Լ��λ.����%Type,
  ����_In         In ��Լ��λ.����%Type := Null,
  ��ַ_In         In ��Լ��λ.��ַ%Type := Null,
  �绰_In         In ��Լ��λ.�绰%Type := Null,
  ��������_In     In ��Լ��λ.��������%Type := Null,
  �ʺ�_In         In ��Լ��λ.�ʺ�%Type := Null,
  ��ϵ��_In       In ��Լ��λ.��ϵ��%Type := Null,
  ĩ��_In         In ��Լ��λ.ĩ��%Type := 1,
  �����ʼ�_In     In ��Լ��λ.�����ʼ�%Type := Null,
  ˵��_In         In ��Լ��λ.˵��%Type := Null,
  վ��_In         In ��Լ��λ.վ��%Type := Null,
  ������ô���_In In ��Լ��λ.������ô���%Type := Null
) Is
Begin
  --���Ȳ����¼ 
  Insert Into ��Լ��λ
    (ID, ����, ����, ����, ��ַ, �绰, ��������, �ʺ�, ��ϵ��, �ϼ�id, ����ʱ��, ����ʱ��, ĩ��, �����ʼ�, ˵��, վ��, ������ô���)
  Values
    (Id_In, ����_In, ����_In, ����_In, ��ַ_In, �绰_In, ��������_In, �ʺ�_In, ��ϵ��_In, �ϼ�id_In, Sysdate,
     To_Date('3000-01-01', 'yyyy-mm-dd'), ĩ��_In, �����ʼ�_In, ˵��_In, վ��_In, ������ô���_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Լ��λ_Insert;
/

Create Or Replace Procedure Zl_��Լ��λ_Update
(
  Id_In           In ��Լ��λ.Id%Type,
  �ϼ�id_In       In ��Լ��λ.�ϼ�id%Type,
  ����_In         In ��Լ��λ.����%Type,
  ����_In         In ��Լ��λ.����%Type,
  ����_In         In ��Լ��λ.����%Type,
  ��ַ_In         In ��Լ��λ.��ַ%Type := Null,
  �绰_In         In ��Լ��λ.�绰%Type := Null,
  ��������_In     In ��Լ��λ.��������%Type := Null,
  �ʺ�_In         In ��Լ��λ.�ʺ�%Type := Null,
  ��ϵ��_In       In ��Լ��λ.��ϵ��%Type := Null,
  ԭ����_In       In Number,
  �����ʼ�_In     In ��Լ��λ.�����ʼ�%Type := Null,
  ˵��_In         In ��Լ��λ.˵��%Type := Null,
  վ��_In         In ��Լ��λ.վ��%Type := Null,
  ������ô���_In In ��Լ��λ.������ô���%Type := Null
) Is
Begin
  --���Ȳ����޸ļ�¼ 
  Update ��Լ��λ
  Set ���� = ����_In, ���� = ����_In, ���� = ����_In, ��ַ = ��ַ_In, �绰 = �绰_In, �������� = ��������_In, �ʺ� = �ʺ�_In, ��ϵ�� = ��ϵ��_In,
      �ϼ�id = �ϼ�id_In, �����ʼ� = �����ʼ�_In, ˵�� = ˵��_In, վ�� = վ��_In, ������ô��� = ������ô���_In
  Where ID = Id_In;

  --�������¼�ҲҪ�޸ı��� 
  Update ��Լ��λ
  Set ���� = ����_In || Substr(����, ԭ����_In)
  Where ID In (Select ID From ��Լ��λ Start With �ϼ�id = Id_In Connect By Prior ID = �ϼ�id);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Լ��λ_Update;
/
Create Or Replace Procedure Zl_����Ԥ����¼_Delete
(
  Id_In         ����Ԥ����¼.Id%Type,
  ժҪ_In       ����Ԥ����¼.ժҪ%Type,
  ����Ա���_In ����Ԥ����¼.����Ա���%Type,
  ����Ա����_In ����Ԥ����¼.����Ա����%Type,
  �ʻ��˷�_In   Number := 1,
  ��Ԥ��id_In   ����Ԥ����¼.Id%Type := Null,
  Ʊ�ݺ�_In     ����Ԥ����¼.ʵ��Ʊ��%Type := Null,
  ����id_In     Ʊ�����ü�¼.Id%Type := Null,
  У�Ա�־_In   ����Ԥ����¼.У�Ա�־%Type := Null,
  ����ģʽ_In   Number := 0,
  ����״̬_In   Number := 0,
  ��������_In   Number := 0,
  ���ַ�ʽ_In   ����Ѻ���¼.���㷽ʽ%Type := Null
) As
  --У�Ա�־_In  �������˿�ʱ���ȴ���1������У�Ա�־Ϊ1�ļ�¼���ٴ���ջ�0����Ϊ0�������˿��¼��
  --����ģʽ_In  0-ͬ����ɣ�1-�첽���
  --����״̬_In  ����ģʽ_In=1ʱ��0-�쳣״̬��1-��ɽ���
  --��������_In:����֧����Ԥ�����Ƿ�����  0�������֣�1������
  Cursor c_Moneyinfo Is
    Select ID, NO, ���, ���㷽ʽ, ����id, Ԥ�����
    From ����Ԥ����¼
    Where ID = Id_In And ��¼���� = 1 And (��¼״̬ = 1 Or ��¼״̬ = 3);
  r_Moneyrow c_Moneyinfo%RowType;

  v_��ӡid   Ʊ�ݴ�ӡ����.Id%Type;
  v_����     ���㷽ʽ.����%Type;
  v_�ֽ�     ���㷽ʽ.����%Type;
  v_�����ʻ� ���㷽ʽ.����%Type;
  n_����ֵ   �������.Ԥ�����%Type;
  n_Ԥ��id   ����Ԥ����¼.Id%Type;
  v_No       ����Ԥ����¼.No%Type;
  n_�����id ����Ԥ����¼.�����id%Type;
  v_Date     Date;
  Err_Custom Exception;
  n_��id         ����ɿ����.Id%Type;
  v_Msg          Varchar2(500);
  n_Ԥ������Ʊ�� Number(1);
Begin
  n_Ԥ��id := ��Ԥ��id_In;

  --��ȡ���㷽ʽ����
  Select Max(Nvl(����, '�ֽ�')) Into v_�ֽ� From ���㷽ʽ Where ���� = 1;
  Select Max(Nvl(����, '�����ʻ�')) Into v_�����ʻ� From ���㷽ʽ Where ���� = 3;

  Open c_Moneyinfo;
  Fetch c_Moneyinfo
    Into r_Moneyrow;

  --�����ж�Ҫ�˿�ļ�¼�Ƿ����
  If c_Moneyinfo%RowCount = 0 Then
    Close c_Moneyinfo;
    Raise Err_Custom;
  End If;
  Select Sysdate Into v_Date From Dual;
  If n_Ԥ��id Is Null Then
    Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
  End If;
  n_��id := Zl_Get��id(����Ա����_In);

  If Not (����ģʽ_In = 1 And ����״̬_In = 1) Then
  
    If Nvl(�ʻ��˷�_In, 0) = 1 Then
      --֧�ָ����ʻ��˷�,��������
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���, ���㷽ʽ, �������, �տ�ʱ��, ����Ա���, ����Ա����, �ɿλ, ��λ������, ��λ�ʺ�, �ɿ���id,
         Ԥ�����, �����id, ����, ������ˮ��, ����˵��, ������λ, ���㿨���, У�Ա�־, ��������id, ����ʱ��, ������Ա, Ԥ������Ʊ��)
        Select n_Ԥ��id, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ_In, -1 * ���, Decode(��������_In, 1, ���ַ�ʽ_In, ���㷽ʽ),
               Decode(��������_In, 1, Null, �������), v_Date, ����Ա���_In, ����Ա����_In, �ɿλ, ��λ������, ��λ�ʺ�, n_��id, Ԥ�����,
               Decode(��������_In, 1, Null, �����id), Decode(��������_In, 1, Null, ����), Decode(��������_In, 1, Null, ������ˮ��),
               Decode(��������_In, 1, Null, ����˵��), ������λ, ���㿨���, У�Ա�־_In, ��������id, Decode(��������_In, 1, Null, v_Date),
               Decode(��������_In, 1, Null, ����Ա����_In), Ԥ������Ʊ��
        From ����Ԥ����¼
        Where ID = Id_In;
    Else
      --��֧��ʱ,������ֽ�,��¼����Ϊ2��ժҪ���־,Ϊ3�ĸ����������ժҪ
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���, ���㷽ʽ, �������, �տ�ʱ��, ����Ա���, ����Ա����, �ɿλ, ��λ������, ��λ�ʺ�, �ɿ���id,
         Ԥ�����, �����id, ����, ������ˮ��, ����˵��, ������λ, ���㿨���, У�Ա�־, ��������id, ����ʱ��, ������Ա, Ԥ������Ʊ��)
        Select n_Ԥ��id, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, Nvl(ժҪ_In, '�����ʻ��˿�'), -1 * ���,
               Decode(���㷽ʽ, v_�����ʻ�, v_�ֽ�, Decode(��������_In, 1, ���ַ�ʽ_In, ���㷽ʽ)), Decode(��������_In, 1, Null, �������), v_Date,
               ����Ա���_In, ����Ա����_In, Decode(���㷽ʽ, v_�����ʻ�, Null, �ɿλ), Decode(���㷽ʽ, v_�����ʻ�, Null, ��λ������),
               Decode(���㷽ʽ, v_�����ʻ�, Null, ��λ�ʺ�), n_��id, Ԥ�����, Decode(��������_In, 1, Null, �����id),
               Decode(��������_In, 1, Null, ����), Decode(��������_In, 1, Null, ������ˮ��), Decode(��������_In, 1, Null, ����˵��), ������λ,
               ���㿨���, У�Ա�־_In, ��������id, Decode(��������_In, 1, Null, v_Date), Decode(��������_In, 1, Null, ����Ա����_In), Ԥ������Ʊ��
        From ����Ԥ����¼
        Where ID = Id_In;
    End If;
    Select �����id Into n_�����id From ����Ԥ����¼ Where ID = Id_In;
    If Nvl(n_�����id, 0) <> 0 Then
      --�Զ�����̵���
      Zl_Custom_Balance_Update(n_Ԥ��id);
    End If;
    Update ����Ԥ����¼ Set ��¼״̬ = 3 Where ID = Id_In;
    --����(Ԥ��)���(���������ֽ��Ǹ����ʻ���Ӧ�ü���)
    --�ж�Ҫ�˿������
    Select b.���� Into v_���� From ����Ԥ����¼ A, ���㷽ʽ B Where a.���㷽ʽ = b.����(+) And a.Id = Id_In;
    If Nvl(v_����, 1) <> 5 Then
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) - r_Moneyrow.���
      Where ���� = 1 And ����id = r_Moneyrow.����id And Nvl(����, 2) = Nvl(r_Moneyrow.Ԥ�����, 2)
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, ����, Ԥ�����, �������)
        Values
          (r_Moneyrow.����id, 1, Nvl(r_Moneyrow.Ԥ�����, 2), -r_Moneyrow.���, 0);
        n_����ֵ := -r_Moneyrow.���;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From �������
        Where ����id = r_Moneyrow.����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
      End If;
    End If;
  
    --Ԥ���������
    Update Ԥ���������
    Set Ԥ����� = Nvl(Ԥ�����, 0) - r_Moneyrow.���
    Where Ԥ��id = Id_In
    Returning Ԥ����� Into n_����ֵ;
    If Sql%RowCount = 0 Then
      Insert Into Ԥ���������
        (Ԥ��id, ����id, Ԥ�����, Ԥ�����)
      Values
        (r_Moneyrow.Id, r_Moneyrow.����id, Nvl(r_Moneyrow.Ԥ�����, 2), -r_Moneyrow.���);
      n_����ֵ := -r_Moneyrow.���;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From Ԥ���������
      Where Ԥ��id = r_Moneyrow.Id And Nvl(Ԥ�����, 2) = Nvl(r_Moneyrow.Ԥ�����, 2) And Nvl(Ԥ�����, 0) = 0;
    End If;
  End If;

  --�첽�������ʱ��ִ�д�����
  If ����ģʽ_In = 1 Then
    If ����״̬_In = 1 Then
      --���������˿������У�Ա�־Ϊ0�ļ�¼����
      Update ����Ԥ����¼
      Set У�Ա�־ = У�Ա�־_In, �տ�ʱ�� = v_Date, ����Ա��� = ����Ա���_In, ����Ա���� = ����Ա����_In, �ɿ���id = n_��id, ����ʱ�� = v_Date,
          ������Ա = ����Ա����_In
      Where ID = n_Ԥ��id;
      --�Զ�����̵���
      Zl_Custom_Balance_Update(n_Ԥ��id);
    Else
      Return;
    End If;
  End If;

  --������ػ��ܱ�
  --��Ա�ɿ����(ע�������������ʻ��Ľ��㷽ʽ)
  If Nvl(�ʻ��˷�_In, 0) = 1 Then
    --֧���˸����ʻ�ʱ�Ĵ���
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) - r_Moneyrow.���
    Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = Decode(��������_In, 1, ���ַ�ʽ_In, r_Moneyrow.���㷽ʽ)
    Returning ��� Into n_����ֵ;
  
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ����
        (�տ�Ա, ���㷽ʽ, ����, ���)
      Values
        (����Ա����_In, Decode(��������_In, 1, ���ַ�ʽ_In, r_Moneyrow.���㷽ʽ), 1, -r_Moneyrow.���);
      n_����ֵ := -r_Moneyrow.���;
    
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ��Ա�ɿ����
      Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = Decode(��������_In, 1, ���ַ�ʽ_In, r_Moneyrow.���㷽ʽ) And Nvl(���, 0) = 0;
    End If;
  Else
    --��֧��ʱ�Ĵ���
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) - r_Moneyrow.���
    Where ���� = 1 And �տ�Ա = ����Ա����_In And
          ���㷽ʽ = Decode(r_Moneyrow.���㷽ʽ, v_�����ʻ�, v_�ֽ�, Decode(��������_In, 1, ���ַ�ʽ_In, r_Moneyrow.���㷽ʽ))
    Returning ��� Into n_����ֵ;
  
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ����
        (�տ�Ա, ���㷽ʽ, ����, ���)
      Values
        (����Ա����_In, Decode(r_Moneyrow.���㷽ʽ, v_�����ʻ�, v_�ֽ�, Decode(��������_In, 1, ���ַ�ʽ_In, r_Moneyrow.���㷽ʽ)), 1,
         -r_Moneyrow.���);
      n_����ֵ := -r_Moneyrow.���;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ��Ա�ɿ����
      Where �տ�Ա = ����Ա����_In And ���� = 1 And
            ���㷽ʽ = Decode(r_Moneyrow.���㷽ʽ, v_�����ʻ�, v_�ֽ�, Decode(��������_In, 1, ���ַ�ʽ_In, r_Moneyrow.���㷽ʽ)) And
            Nvl(���, 0) = 0;
    End If;
  End If;
  --�����ջ�Ʊ��(������ǰû��ʹ��Ʊ��,�޷��ջ�)
  Begin
    Select Nvl(Ԥ������Ʊ��, 0) Into n_Ԥ������Ʊ�� From ����Ԥ����¼ Where ID = Id_In;
    If n_Ԥ������Ʊ�� = 0 Then
      Select ID
      Into v_��ӡid
      From (Select b.Id
             From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
             Where a.��ӡid = b.Id And a.���� = 1 And b.�������� = 2 And b.No = r_Moneyrow.No
             Order By a.ʹ��ʱ�� Desc)
      Where Rownum < 2;
    End If;
  Exception
    When Others Then
      Null;
  End;

  If v_��ӡid Is Not Null Then
    Insert Into Ʊ��ʹ����ϸ
      (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
      Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, v_Date, ����Ա����_In, Ʊ�ݽ��
      From Ʊ��ʹ����ϸ
      Where ��ӡid = v_��ӡid And Ʊ�� = 2 And ���� = 1;
  End If;

  --����Ʊ��
  If Ʊ�ݺ�_In Is Not Null Then
    Select Ʊ�ݴ�ӡ����_Id.Nextval Into v_��ӡid From Dual;
  
    --����Ʊ��
    Insert Into Ʊ�ݴ�ӡ���� (ID, ��������, NO) Values (v_��ӡid, 2, r_Moneyrow.No);
  
    Insert Into Ʊ��ʹ����ϸ
      (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
    Values
      (Ʊ��ʹ����ϸ_Id.Nextval, 2, Ʊ�ݺ�_In, 1, 6, ����id_In, v_��ӡid, v_Date, ����Ա����_In, -1 * r_Moneyrow.���);
  
    --״̬�Ķ�
    Update Ʊ�����ü�¼
    Set ��ǰ���� = Ʊ�ݺ�_In, ʣ������ = Decode(Sign(ʣ������ - 1), -1, 0, ʣ������ - 1), ʹ��ʱ�� = Sysdate
    Where ID = Nvl(����id_In, 0);
  
  End If;

  Close c_Moneyinfo;

  --��Ϣ����;
  Select NO Into v_No From ����Ԥ����¼ Where ID = n_Ԥ��id;
  b_Message.Zlhis_Charge_006(n_Ԥ��id, v_No);
  Select Id_In || ',' || �ʻ��˷�_In Into v_Msg From Dual;
  Begin
    Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
      Using 12, v_Msg;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20999, '[ZLSOFT]û�з���Ҫ�˿��Ԥ����¼,�ü�¼�����Ѿ��˳���[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ԥ����¼_Delete;
/

Create Or Replace Procedure Zl1_Datamove_Tag
(
  d_End            In Date,
  n_����           In Number,
  n_System         In Number,
  n_Ԥ��ʣ������� In ����Ԥ����¼.���%Type := 10 --�����˲�����δ����ã�Ҳ������Ժ����ʱ������δ�����Ԥ������ָ��ֵ���µ�����ǿ��ת���������������δת���Ӷ�Ӱ��ת���ٶ�
  
) As
  --���ܣ���Ǵ�ת��������
  --˵����Ϊ����Undo��ռ����͹��󣬷ֶ��ύ
  d_Lastend Date; --����ת����ֹʱ�䣨d_EndΪ����ת����ֹʱ�䣩

  --�ݹ�ȡ����һ��Ԥ������е�һ���ֱ����Ϊ��ת����������
  Procedure Datamove_Tag_Update
  (
    ����id_In t_NumList,
    d_End     In Date,
    n_����    In Number
  ) As
  
    c_����id t_NumList := t_NumList();
    c_No     t_StrList := t_StrList();
  Begin
    --1.1һ��Ԥ�����ݱ��������ID���ˣ��ҳ����е�һ���ֱ����Ϊ��ת�������ݣ��磺
    --   NO=A001 ��¼����=11 ����ID=10 ��ת��=1
    --   NO=A001 ��¼����=11 ����ID=11 ��ת��=NULL
    If ����id_In Is Null Then
      Select Distinct a.No Bulk Collect
      Into c_No
      From ����Ԥ����¼ A
      Where a.��¼���� In (1, 11) And a.��ת�� = n_���� And Exists
       (Select 1 From ����Ԥ����¼ Where NO = a.No And ��¼���� In (1, 11) And ��ת�� Is Null);
    Else
      Select Distinct a.No Bulk Collect
      Into c_No
      From ����Ԥ����¼ A
      Where a.����id In (Select /*+cardinality(b,10) */
                        Column_Value
                       From Table(����id_In) B) And a.��¼���� In (1, 11) And a.��ת�� Is Null And Exists
       (Select 1 From ����Ԥ����¼ Where NO = a.No And ��¼���� In (1, 11) And ��ת�� + 0 = n_����);
    End If;
  
    If c_No.Count = 0 Then
      Return;
    End If;
  
    --1.2ȡ�����
    Forall I In 1 .. c_No.Count
      Update ����Ԥ����¼ Set ��ת�� = Null Where NO = c_No(I) And ��¼���� In (1, 11);
  
    --------------------------------------------------------------------------------------------------------
    --2.1һ������ID���˶���Ԥ�����ݣ��ҳ����е�һ���ֱ����Ϊ��ת�������ݣ��磺
    --   NO=A001 ��¼����=11 ����ID=20 ��ת��=1
    --   NO=A002 ��¼����=11 ����ID=20 ��ת��=NULL
    Select Distinct a.����id Bulk Collect
    Into c_����id
    From ����Ԥ����¼ A
    Where a.No In (Select /*+cardinality(b,10) */
                    Column_Value
                   From Table(c_No) B) And a.��¼���� In (1, 11) And a.��ת�� Is Null And a.�տ�ʱ�� + 0 < d_End And Exists
     (Select 1 From ����Ԥ����¼ Where ����id = a.����id And ��ת�� + 0 = n_����);
  
    If c_����id.Count = 0 Then
      Return;
    End If;
  
    --2.2ȡ�����(����һ�ν��ʵ��������㷽ʽ�ļ�¼)
    Forall I In 1 .. c_����id.Count
      Update ����Ԥ����¼ Set ��ת�� = Null Where ����id = c_����id(I);
  
    --�ݹ����
    Datamove_Tag_Update(c_����id, d_End, n_����);
  End Datamove_Tag_Update;
Begin
  Select ������������ Into d_Lastend From zlDataMove Where ϵͳ = n_System And ��� = 1;
  If d_Lastend Is Null Then
    Return;
  End If;
  --�¼��Ӳ�ѯע�������Ż������ܹ������ݹ��˵���С�������ŵ����Exists��������ǰ��

  --1.���ú��㣨����,ҩƷ,�տ��Ʊ�ݵȣ�
  --����ҵ����ԭʼҵ��ķ���ʱ����ͬ���Ǽ�ʱ�䲻ͬ������Ҫ������ʱ������ѯ.
  --��������������ж������ID�����漰������õ��ݣ���Щ����Ҫһ��ת�����ų�ת��������Ӱ������ж��Ƿ����
  --1.һ�ŷ��õ��ݵ�һ�з��û���з��ÿ��ֶܷ�ν��ʣ��ж����ͬ�Ľ���ID��
  --2.�������Ϻ�Ҳ���ֶܷ�ν���(һ�ŵ��ݶ����ͬ�Ľ���ID)
  --3.�������Ϻ�������������õ���һ���(һ�ŵ��ݵĶ������ID���漰�������NO����ЩNO����֮ǰ�������Ϲ�������������ID)
  --���ǵ�������ĸ����ԣ�Ϊ���߼���������ѯ���ܣ�������ID���ų�(�ò��˵Ľ������ݶ���ת��)

  Update /*+ rule*/ ����Ԥ����¼ L
  Set ��ת�� = n_����
  Where ����id In
        (Select Distinct a.����id --1.�����շѺ͹Һŵ��շѽ����¼
         From ������ü�¼ A
         Where (a.��¼״̬ In (1, 2) Or a.��¼״̬ = 3 And Not Exists
                (Select 1
                 From ������ü�¼ B
                 Where a.No = b.No And a.��¼���� = b.��¼���� And b.��¼״̬ = 2 And b.�Ǽ�ʱ�� >= d_Lastend)) And a.��ת�� Is Null And
               a.��¼���� In (1, 4) And a.����ʱ�� < d_End And a.�Ǽ�ʱ�� < d_Lastend
         Union All
         Select Distinct b.����id --2.ҽ��������(û�з���ʱ���ֶ�,���ϼ�¼�ĵǼ�ʱ�䲻ͬ��Ϊ�˰��շѺ����ϵ�һ����ת��������Ҫ����B��)
         From ���ò����¼ A, ���ò����¼ B
         Where a.��ת�� Is Null And a.No = b.No And a.��¼���� = b.��¼���� And a.�Ǽ�ʱ�� < d_End
         Union All
         Select Distinct a.����id --3.���￨���շѽ����¼(�ų�֮���˿��ѵ�,һ�ŵ�����ֻҪ����һ������)
         From סԺ���ü�¼ A
         Where (a.��¼״̬ In (1, 2) Or a.��¼״̬ = 3 And Not Exists
                (Select 1
                 From סԺ���ü�¼ B
                 Where a.No = b.No And a.��¼���� = b.��¼���� And b.��¼״̬ = 2 And b.�Ǽ�ʱ�� >= d_Lastend)) And a.��ת�� Is Null And
               a.���ʷ��� = 0 And a.��¼���� = 5 And a.����ʱ�� < d_End
         Union All --4.סԺ���ʷ��õĽ��ʽ����¼
         Select ����id
         From (With Settle As (Select Distinct c.����id
                               From (Select Distinct b.No, b.���, Mod(b.��¼����, 10) As ��¼����
                                      From (Select Distinct b.Id
                                             From ���˽��ʼ�¼ A, ���˽��ʼ�¼ B --���ϵĽ��ʵ����շ�ʱ�������ָ��ʱ��֮������Ҫ����B��
                                             Where (a.��¼״̬ In (1, 2) Or a.��¼״̬ = 3 And Not Exists
                                                    (Select 1
                                                     From ���˽��ʼ�¼ C
                                                     Where a.No = c.No And c.��¼״̬ = 2 And c.�շ�ʱ�� >= d_Lastend)) And
                                                   a.��ת�� Is Null And a.No = b.No And (a.�������� = 2 Or Nvl(a.��������, 0) = 0) And
                                                   a.�շ�ʱ�� < d_End) A, סԺ���ü�¼ B
                                      Where a.Id = b.����id) B, סԺ���ü�¼ C --ͨ��C���ҵ���Щ���õ��ݵ����н���IDһ��ת(������ת��ʱ��֮��)
                               Where c.No = b.No And Mod(c.��¼����, 10) = b.��¼���� And c.��� = b.���)
                Select ����id
                From Settle
                Minus
                Select Distinct a.Id
                From ���˽��ʼ�¼ A,
                     (Select Distinct ����id
                       From (Select c.����id, c.No, Mod(c.��¼����, 10) As ��¼����, Nvl(Sum(c.ʵ�ս��), 0) As ʵ�ս��,
                                     Nvl(Sum(c.���ʽ��), 0) As ���ʽ��
                              From סԺ���ü�¼ C, Settle S
                              Where c.����id = s.����id
                              Group By c.No, Mod(c.��¼����, 10), c.����id) C
                       Where c.ʵ�ս�� <> c.���ʽ�� And Exists (Select 1 From ��Ժ���� F Where c.����id = f.����id) --��Ժ����û�н����Ҳת�ߣ�����Ҫʱ�ٳ�أ��������ų���������̫��
                             Or Exists (Select 1
                              From סԺ���ü�¼ E, ���˽��ʼ�¼ S
                              Where e.No = c.No And Mod(e.��¼����, 10) = c.��¼���� And e.����id = s.Id And
                                    s.��ת�� Is Null And s.�շ�ʱ�� >= d_Lastend)) N --��ʹ���ڱ���ת��ʱ��֮����壬ֻҪ����������ת��ʱ��֮�󣬾Ͳ��ų�
                
                Where a.����id = n.����id And (a.�������� = 2 Or Nvl(a.��������, 0) = 0))
                Union All --5.������ʷ��õĽ��ʽ����¼
                Select ����id
                From (With Settle As (Select Distinct c.����id
                                      From (Select Distinct b.No, b.���, Mod(b.��¼����, 10) As ��¼����
                                             From (Select Distinct b.Id
                                                    From ���˽��ʼ�¼ A, ���˽��ʼ�¼ B
                                                    Where a.��ת�� Is Null And a.No = b.No And (a.�������� = 1 Or Nvl(a.��������, 0) = 0) And
                                                          a.�շ�ʱ�� < d_End) A, ������ü�¼ B
                                             Where a.Id = b.����id) B, ������ü�¼ C
                                      Where c.No = b.No And Mod(c.��¼����, 10) = b.��¼���� And c.��� = b.���)
                       Select ����id
                       From Settle
                       Minus
                       Select Distinct a.Id
                       From ���˽��ʼ�¼ A,
                            (Select Distinct c.����id
                              From (Select c.����id, c.No, Mod(c.��¼����, 10) As ��¼����, Nvl(Sum(c.ʵ�ս��), 0) As ʵ�ս��,
                                            Nvl(Sum(c.���ʽ��), 0) As ���ʽ��
                                     From ������ü�¼ C, Settle S
                                     Where c.����id = s.����id
                                     Group By c.No, Mod(c.��¼����, 10), c.����id) C
                              Where c.ʵ�ս�� <> c.���ʽ�� --���ﲡ��û�н���Ĳ�ת��
                                    Or Exists (Select 1
                                     From ������ü�¼ E, ���˽��ʼ�¼ S
                                     Where e.No = c.No And Mod(e.��¼����, 10) = c.��¼���� And e.����id = s.Id And
                                           s.��ת�� Is Null And s.�շ�ʱ�� >= d_Lastend)) N
                       Where a.����id = n.����id And (a.�������� = 1 Or Nvl(a.��������, 0) = 0))
                       
         
         
         );

  --�ų�Ԥ����δ�����
  --Ϊ�˽����߼��ĸ����ԣ����ų���ת��ʱ��֮��ҩ��δ��ҩ�ķ��ü�¼��Ӧ�Ľ���ID������������Ľ������ݺͷ�������ǿ��ת��
  --��Ϊǰ���SQL����Ľ���ID���ܲ�ȫ�ǳ�Ԥ����(�����շѺ�סԺ���ʲ��ѵ�)�����ԣ���Ҫ����һ��SQL���ų�
  Update /*+ rule*/ ����Ԥ����¼
  Set ��ת�� = Null
  Where ��ת�� = n_���� And
        ����id In
        (Select Distinct d.����id --�õ�����ص����г�Ԥ���Ľ���ID����ת��
         From ����Ԥ����¼ D,
              (Select Distinct l.No
                From (Select l.No, l.����id, l.Ԥ�����, Nvl(Sum(l.���), 0) As ���, Nvl(Sum(l.��Ԥ��), 0) As ��Ԥ��,
                              Sum(Decode(l.��ת��, Null, Decode(����id, Null, Decode(��¼״̬, 2, 0, 1), 1), 0)) As δת��
                       From ����Ԥ����¼ L --���ܰ�����IDȷ�ϱ��δ�ת���ĳ��ֻ��ʣ��������Ҫ����L����ԭʼ��Ԥ���ĵ��ݣ��Լ���¼����Ϊ11�Ŀ��ܻ���ת��ʱ��֮��������ʣ���Ľ���ID
                       Where l.��¼���� In (1, 11) And
                             l.No In
                             (Select Distinct p.No From ����Ԥ����¼ P Where p.��¼���� In (1, 11) And p.��ת�� = n_����)
                       Group By l.No, l.����id, l.Ԥ�����) L --���סԺ����һ�ν��壬���ԣ����ܼ���ҳID
                Where δת�� > 0 --ֻҪ��Ԥ�����ݻ���δת����Ԥ�����Ԥ����¼����ת��������ת��һ���ֵ��º����жϴ���
                      Or
                      l.��� <> l.��Ԥ�� And
                      (Exists (Select 1
                               From ����Ԥ����¼ E --ʣ��Ԥ���һ���ø�����Ԥ�����˿NO�Ų�ͬ���������൱���ǳ����ˣ����ų�
                               Where e.����id = l.����id And e.Ԥ����� = l.Ԥ����� And e.��¼���� In (1, 11) And
                                     (e.��ת�� = n_���� Or e.��ת�� Is Null And e.����id Is Null And e.��¼���� = 1 And �տ�ʱ�� < d_End)
                                Having Abs(Nvl(Sum(e.���), 0) - Nvl(Sum(e.��Ԥ��), 0)) > n_Ԥ��ʣ�������) --���С�ڵ���n���ų����������3�ֽ���IDΪ�յ�Ҫ����һ��
                       Or l.Ԥ����� = 2 And Exists (Select 1 From ��Ժ���� E Where l.����id = e.����id) Or Exists
                       (Select 1
                        From ����δ����� E
                        Where l.����id = e.����id And (l.Ԥ����� = 1 And e.��ҳid Is Null Or l.Ԥ����� = 2 And e.��ҳid Is Not Null)))) N
         Where d.No = n.No And d.��¼���� In (1, 11));

  --��������3�ֽ���IDΪ�յ�Ԥ����¼
  --1.Ԥ����û��ʹ�þ�ֱ�����˵ļ�¼(����IDΪ��)
  Update /*+ rule*/ ����Ԥ����¼
  Set ��ת�� = n_����
  Where ��¼���� = 1 And
        NO In (Select a.No
               From ����Ԥ����¼ A
               Where a.����id Is Null And a.��¼���� = 1 And a.��¼״̬ In (2, 3) And a.��ת�� Is Null And a.�տ�ʱ�� < d_End
               Group By a.No
               Having Sum(a.���) = 0);

  --2.��Ԥ������˿�ļ�¼������IDΪ�գ���¼״̬Ϊ2��
  Update /*+ rule*/ ����Ԥ����¼
  Set ��ת�� = n_����
  Where ����id Is Null And ��¼���� = 1 And ��¼״̬ = 2 And
        NO In (Select a.No From ����Ԥ����¼ A Where a.��¼���� = 1 And a.��¼״̬ = 3 And a.��ת�� = n_����);

  --�ų�ͬһ��Ԥ����ݲ��ּ�¼�����Ϊת����,ֻҪ�в�ת���ģ������ŵ��ݶ���ת��
  --����2���й���Ӱ�죬����Ҫ������֮��ִ��
  --ҪӰ���3��������жϣ�����Ҫ������֮ǰִ��
  Datamove_Tag_Update(Null, d_End, n_����);

  --3.Ԥ����δ����ʱ�ý�����Ԥ�����˿�(����IDΪ�գ����Ҹ�ԭʼ�ĳ�Ԥ����NOû�й�����ϵ)
  --��������"��� < 0"����Ϊ����Ԥ����û��ʹ�ù�����ֱ���ý�����Ԥ�����˿�����
  Update /*+ rule*/ ����Ԥ����¼ L
  Set ��ת�� = n_����
  Where Exists (Select 1
         From ����Ԥ����¼ E
         Where e.����id = l.����id And e.Ԥ����� = l.Ԥ����� And e.��¼���� In (1, 11) And
               (e.��ת�� = n_���� Or e.��ת�� Is Null And e.����id Is Null And e.��¼���� = 1 And ��¼״̬ = 1 And �տ�ʱ�� < d_End)
         Group By e.����id
         Having Abs(Nvl(Sum(e.���), 0) - Nvl(Sum(e.��Ԥ��), 0)) <= n_Ԥ��ʣ�������) --���С�ڵ���nҪת������ǰ�桰�ų�Ԥ����δ����ġ�Ҫ����һ��
       
        And Exists (Select 1
         From ����Ԥ����¼ E
         Where e.����id = l.����id And e.Ԥ����� = l.Ԥ����� And e.��¼���� In (1, 11) And e.��ת�� = n_����) And
        ��ת�� Is Null And ����id Is Null And ��¼���� = 1 And ��¼״̬ = 1 And �տ�ʱ�� < d_End;

  Update /*+ rule*/ ����Ѻ���¼ Set ��ת�� = n_���� Where ��¼״̬ In (2, 3) And ��ת�� Is Null And �տ�ʱ�� < d_End;

  Update /*+ rule*/ �������㽻��
  Set ��ת�� = n_����
  Where ����id In (Select a.Id From ����Ѻ���¼ A Where ��ת�� = n_���� And Nvl(����, 0) = 2);

  Update /*+ rule*/ ����Ʊ��ʹ�ü�¼
  Set ��ת�� = n_����
  Where Ʊ�� = 2 And
        ����id In
        (Select ID From ����Ԥ����¼ Where ��ת�� = n_���� And Nvl(Ԥ������Ʊ��, 0) = 1 And Mod(��¼����, 10) = 1);

  Update /*+ rule*/ ����Ʊ��ʹ�ü�¼
  Set ��ת�� = n_����
  Where Ʊ�� <> 2 And
        ����id In
        (Select ID From ����Ԥ����¼ Where ��ת�� = n_���� And Nvl(�Ƿ����Ʊ��, 0) = 1 And Mod(��¼����, 10) <> 1);

  Update /*+ rule*/ ����Ʊ�ݶ�ά��
  Set ��ת�� = n_����
  Where ʹ�ü�¼id In (Select ID From ����Ʊ��ʹ�ü�¼ Where ��ת�� = n_����);

  --Ԥ��Ʊ�ݣ����ϸ���ƣ�
  Update /*+ rule*/ Ʊ��ʹ����ϸ
  Set ��ת�� = n_����
  Where Ʊ�� = 2 And ���� In (Select Distinct ʵ��Ʊ��
                          From ����Ԥ����¼
                          Where Mod(��¼����, 0) = 1 And Nvl(У�Ա�־, 0) = 0 And ��ת�� = n_����) And Nvl(����id, 0) = 0;

  Update zlDataMovelog
  Set ��ǰ���� = '(1/11)�������ݱ����ɣ����ڱ�Ƿ�������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  Update /*+ rule*/ ���˽��ʼ�¼
  Set ��ת�� = n_����
  Where ID In (Select ����id From ����Ԥ����¼ Where ��ת�� = n_����);

  --�����޽���ļ�¼(Ϊ���������ܣ����жϷ��ã�ֻҪ����������Ԥ����¼�͵���������ý���)
  Update /*+ rule*/ ���˽��ʼ�¼ L
  Set ��ת�� = n_����
  Where �շ�ʱ�� < d_End And ��ת�� Is Null And Not Exists (Select 1 From ����Ԥ����¼ P Where l.Id = p.����id);

  --����Ʊ�ݣ����ϸ���ƣ�
  Update /*+ rule*/ Ʊ��ʹ����ϸ
  Set ��ת�� = n_����
  Where Ʊ�� = 3 And ���� In (Select Distinct ʵ��Ʊ�� From ���˽��ʼ�¼ Where Nvl(����״̬, 0) = 0 And ��ת�� = n_����) And Nvl(����id, 0) = 0;

  Update /*+ rule*/ ���˿������¼
  Set ��ת�� = n_����
  Where ��¼���� = 4 And ����id In (Select a.Id From ����Ԥ����¼ A Where ��ת�� = n_����);

  Update /*+ rule*/ �������㽻��
  Set ��ת�� = n_����
  Where ����id In (Select a.Id From ����Ԥ����¼ A Where ��ת�� = n_����) And Nvl(����, 0) = 0;

  Update /*+ rule*/ �����˿���Ϣ
  Set ��ת�� = n_����
  Where (��¼id, ����id) In (Select a.Id, a.����id From ����Ԥ����¼ A Where ��ת�� = n_����);

  --1.�Һŷ����쳣����
  --a.����IDΪ�գ�ʵ�ս����ܲ�Ϊ�㣩
  --b.����ID��Ϊ�գ����ۺ�ʵ�ս��Ϊ0��Ӧ�ս�������������ĹҺŷ��ã�û�йҺż�¼��Ҳû��Ԥ����¼
  --������ʱ��ת������Ϊ�պ��˵ķ���ʱ����ͬ���Ǽ�ʱ�䲻ͬ��
  Update /*+ rule*/ ������ü�¼
  Set ��ת�� = n_����
  Where ��ת�� Is Null And ����ʱ�� < d_End And ��¼���� = 4 And (ʵ�ս�� = 0 Or ����id Is Null);

  --2.ֱ���շѵĺͽ����޽��㣨Ԥ������¼�ģ�Union����allȥ���ظ��Լ���in������
  Update /*+ rule*/ ������ü�¼
  Set ��ת�� = n_����
  Where ����id In (Select ����id
                 From ����Ԥ����¼
                 Where ��ת�� = n_����
                 Union
                 Select ID From ���˽��ʼ�¼ Where ��ת�� = n_����);

  --3.û�н���id������(������ʱ��)
  --a.δ���ʵĻ��ۼ�¼
  --b.δ�շѵ������
  --������"��ת�� Is Null"��Ϊ�˴���������α��ת�������
  Update /*+ rule*/ ������ü�¼ A
  Set ��ת�� = n_����
  Where (��¼״̬ = 0 Or ��¼���� = 1 And ʵ�ս�� = 0 And ���ʽ�� = 0) And ����id Is Null And ��ת�� Is Null And ����ʱ�� < d_End;

  --4.û�н���id������(������ʱ��)
  --δ���ʵ�������ʷ���(����)���ò���û��Ԥ�������Ҳ���������ת��ʱ��֮����δ��������ʷ���
  Update /*+ rule*/ ������ü�¼ A
  Set ��ת�� = n_����
  Where Not Exists (Select 1
         From ����Ԥ����¼ B
         Where b.����id = a.����id And b.��ת�� Is Null And b.Ԥ����� = 1 And b.��¼���� In (1, 11) Having
          Nvl(Sum(b.���), 0) <> Nvl(Sum(b.��Ԥ��), 0)) And Not Exists
   (Select 1
         From ������ü�¼ B
         Where a.����id = b.����id And b.��¼���� = 2 And b.����id Is Null And b.��ת�� Is Null And b.�Ǽ�ʱ�� > = d_Lastend) And
        ��¼���� = 2 And ����id Is Null And ��ת�� Is Null And ����ʱ�� < d_End;

  --5.û�н���id������(������ʱ��)
  --���������ļ��ʼ�¼����¼״̬Ϊ2�����Ǽ�ʱ������ڵ�ǰָ��ת��ʱ��֮�󣬶�ԭʼ���ʼ�¼����¼״̬Ϊ3�����Ǽ�ʱ����ָ��ת��ʱ��֮ǰ��ǰ�����ߵķ���ʱ������ͬ�ġ�
  --a.δ���ʵ�����ʷ��û���ۺ�ʵ�ս��Ϊ��ģ�����ģ�����û�й�ѡ������ý��ʣ�
  --b.�������Ϻ󣬼��ʵ����ʵļ�¼������IDΪ���Ҽ�¼״̬Ϊ2�ģ�����¼״̬Ϊ3�����н���ID������ǰ����ת��.
  Update /*+ rule*/ ������ü�¼ A
  Set ��ת�� = n_����
  Where (Exists (Select 1
                 From ������ü�¼ B
                 Where a.No = b.No And a.��¼���� = b.��¼���� And a.��� = b.��� And b.��¼״̬ = 3 And b.����id Is Not Null And
                       b.��ת�� + 0 = n_����) And ��¼״̬ = 2 Or Exists
         (Select 1
          From ������ü�¼ B
          Where a.No = b.No And a.��¼���� = b.��¼���� And a.��� = b.��� And b.����id Is Null
          Group By b.No, b.��¼����, b.���
          Having Nvl(Sum(b.ʵ�ս��), 0) = 0)) And ��¼���� = 2 And ����id Is Null And ��ת�� Is Null And ����ʱ�� < d_End;

  --6.�н���id�������(������ʱ��)
  --a.���ѱ���ۺ���ʽ��Ϊ����շѼ�¼,
  --b.һ�ŵ�����ͬ����ID�Ľ��ʽ��֮��Ϊ0(������Ϊ��)
  --��ʹ��ת��ʱ��֮��ҩ�ģ�Ҳǿ��ת����Ϊ�˼����߼������ԣ���߲�ѯ���ܣ�
  Update /*+ rule*/ ������ü�¼ A
  Set ��ת�� = n_����
  Where (���ʽ�� = 0 Or Exists
         (Select 1 From ������ü�¼ C Where a.����id = c.����id Group By c.����id, c.No Having Sum(c.���ʽ��) = 0)) And Not Exists
   (Select 1 From ����Ԥ����¼ B Where a.����id = b.����id And b.��ת�� Is Null) And ��¼���� = 1 And ����id Is Not Null And
        ��ת�� Is Null And ����ʱ�� < d_End;

  --�շ�Ʊ�ݣ����ϸ���ƣ�
  Update /*+ rule*/ Ʊ��ʹ����ϸ
  Set ��ת�� = n_����
  Where Ʊ�� = 1 And
        ���� In
        (Select ʵ��Ʊ�� From ������ü�¼ Where ��ת�� = n_���� And Mod(��¼����, 10) = 1 And Nvl(����״̬, 0) = 0) And Nvl(����id, 0) = 0;

  --�Һ�Ʊ�ݣ����ϸ���ƣ�
  Update /*+ rule*/ Ʊ��ʹ����ϸ
  Set ��ת�� = n_����
  Where Ʊ�� = 4 And
        ���� In (Select ʵ��Ʊ�� From ������ü�¼ Where ��ת�� = n_���� And ��¼���� = 4 And Nvl(����״̬, 0) = 0) And Nvl(����id, 0) = 0;

  Update /*+ rule*/ ���ý������
  Set ��ת�� = n_����
  Where ����id In (Select ����id From ����Ԥ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ҽ��������ϸ
  Set ��ת�� = n_����
  Where ����id In (Select ����id From ����Ԥ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ���ò����¼
  Set ��ת�� = n_����
  Where ����id In (Select ����id From ����Ԥ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ƾ����ӡ��¼
  Set ��ת�� = n_����
  Where (NO, ��¼����) In (Select NO, ��¼���� From ������ü�¼ Where ��ת�� = n_����);

  --1.��Ԥ����¼����Ϊ��ȡ���￨ֱ���շѵģ��޽���ID��,�ټӽ��ʼ�¼��Ϊ��ȡ�����޽��㣨Ԥ������¼��
  Update /*+ rule*/ סԺ���ü�¼
  Set ��ת�� = n_����
  Where ����id In (Select ����id
                 From ����Ԥ����¼
                 Where ��ת�� = n_����
                 Union
                 Select ID From ���˽��ʼ�¼ Where ��ת�� = n_����);

  --2.û�н���id������(������ʱ��)
  --���������ļ��ʼ�¼����¼״̬Ϊ2����ԭʼ��¼�ͳ�����¼�ķ���ʱ������ͬ�ġ�
  --1)ת���������Ϻ󣬼��ʵ����ʵļ�¼����¼״̬Ϊ2����û�н���ID����(��¼״̬Ϊ3���н���ID��)����ǰ����ת����
  --2)δ���ʵ������(�ѳ����ļ��ʵ�����ۺ�ʵ�ս��Ϊ��)
  --3)û�н���ID�Ļ��ۼ�¼����Ϊת��
  Update /*+ rule*/ סԺ���ü�¼ A
  Set ��ת�� = n_����
  Where ((Exists (Select 1
                  From סԺ���ü�¼ B
                  Where a.No = b.No And a.��¼���� = b.��¼���� And a.��� = b.��� And b.��¼״̬ = 3 And b.����id Is Not Null And
                        b.��ת�� + 0 = n_����) And ��¼״̬ = 2 Or Exists
         (Select 1
           From סԺ���ü�¼ B
           Where a.No = b.No And a.��¼���� = b.��¼���� And a.��� = b.��� And b.����id Is Null
           Group By b.No, b.��¼����, b.���
           Having Nvl(Sum(b.ʵ�ս��), 0) = 0)) And a.��¼���� In (2, 3, 5) Or a.��¼״̬ = 0) And a.����id Is Null And a.��ת�� Is Null And
        a.����ʱ�� < d_End;

  --3.��Ժδ���ʵģ����ʲ��ˣ�����Ϊ�Ǻܾ���ǰ����Щ���ݣ����Ԥ���ѳ��꣬����ΪҪת��
  --ȥ��������ҳ�е�"����ת�� is null"������������ΪһЩ���˿�����֮ǰ����������ת����
  Update /*+ rule*/ סԺ���ü�¼ A
  Set ��ת�� = n_����
  Where ��ת�� Is Null And ����id Is Null And
        (����id, ��ҳid) In (Select ����id, ��ҳid
                         From ������ҳ C
                         Where ��Ժ���� < d_End And ��ת�� Is Null And Not Exists
                          (Select 1
                                From ����Ԥ����¼ B
                                Where b.����id = c.����id And b.��ת�� Is Null And b.Ԥ����� = 2 And b.��¼���� In (1, 11) Having
                                 Nvl(Sum(b.���), 0) <> Nvl(Sum(b.��Ԥ��), 0)));
  --���￨Ʊ�ݣ����ϸ���ƣ�
  Update /*+ rule*/ Ʊ��ʹ����ϸ
  Set ��ת�� = n_����
  Where Ʊ�� = 5 And
        ���� In
        (Select ʵ��Ʊ�� From סԺ���ü�¼ Where ��ת�� = n_���� And Mod(��¼����, 10) = 5 And Nvl(����״̬, 0) = 0) And Nvl(����id, 0) = 0;

  Update /*+ rule*/ �����嵥��ӡ
  Set ��ת�� = n_����
  Where (NO, Mod(��¼����, 10), Decode(��¼״̬, 3, 1, ��¼״̬), ���) In
        (Select NO, Mod(��¼����, 10) As ��¼����, Decode(��¼״̬, 3, 1, ��¼״̬) As ��¼״̬, ���
         From ������ü�¼
         Where ��ת�� = n_����
         Union
         Select NO, Mod(��¼����, 10) As ��¼����, Decode(��¼״̬, 3, 1, ��¼״̬) As ��¼״̬, ���
         From סԺ���ü�¼
         Where ��ת�� = n_����);

  Update /*+ rule*/ ���ñ䶯��¼
  Set ��ת�� = n_����
  Where ����id In (Select ID From סԺ���ü�¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ���˷�������
  Set ��ת�� = n_����
  Where ����id In (Select ID From סԺ���ü�¼ Where ��ת�� = n_����);

  Update zlDataMovelog
  Set ��ǰ���� = '(2/11)�������ݱ����ɣ����ڱ��ҩƷ����'
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

  Update /*+ rule*/ ҩƷ�շ������־ A
  Set ��ת�� = n_����
  Where (a.������, a.����) In (Select b.No, b.���� From ҩƷ�շ���¼ B Where b.��ת�� = n_����);

  Update /*+ rule*/ ҩƷ�շ�סԺ��־
  Set ��ת�� = n_����
  Where �շ�id In (Select ID From ҩƷ�շ���¼ Where ��ת�� = n_����);

  Update /*+ rule*/ δ��ҩƷ��¼
  Set ��ת�� = n_����
  Where �շ�id In (Select ID From ҩƷ�շ���¼ Where ��ת�� = n_����);

  Update zlDataMovelog
  Set ��ǰ���� = '(3/11)ҩƷ���ݱ����ɣ����ڱ�ǽɿ���Ʊ������'
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
  Where Not Exists (Select 1 From Ʊ��ʹ����ϸ B Where b.����id = a.Id And b.ʹ��ʱ�� >= d_Lastend) And ��ת�� Is Null And ʣ������ = 0 And
        �Ǽ�ʱ�� < d_End;

  Update /*+ rule*/ Ʊ��ʹ����ϸ
  Set ��ת�� = n_����
  Where ����id In (Select ID From Ʊ�����ü�¼ Where ��ת�� = n_����);

  Update /*+ rule*/ Ʊ�ݴ�ӡ����
  Set ��ת�� = n_����
  Where ID In (Select ��ӡid From Ʊ��ʹ����ϸ Where ��ת�� = n_����);

  Update /*+ rule*/ Ʊ�ݴ�ӡ��ϸ
  Set ��ת�� = n_����
  Where ʹ��id In (Select ID From Ʊ��ʹ����ϸ Where ��ת�� = n_����);

  Update /*+ rule*/ ��������¼
  Set ��ת�� = n_����
  Where �Һ�id In (Select ID From ���˹Һż�¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ��������¼
  Set ��ת�� = n_����
  Where ����id In (Select ID From ��������¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ���ﲡ������
  Set ��ת�� = n_����
  Where ����id In (Select ID From ��������¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ���ﲡ������ָ��
  Set ��ת�� = n_����
  Where ����id In (Select ID From ���ﲡ������ Where ��ת�� = n_����);

  Update zlDataMovelog
  Set ��ǰ���� = '(4/11)�ɿ���Ʊ�����ݱ����ɣ����ڱ�Ǿ��Ｐ��������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  --2.���Ｐ��������
  --��ת�����������Һŷ���δת���ģ�����ת��ʱ��֮�����ҽ������Щҽ����Ϊʱ��û�е�����Ӧת������ҽ����Ӧ�ķ���δת����
  --��ʹ���ھ���(r.ִ��״̬ <> 2 )��Ҳǿ��ת��(ҽ������û��ʹ����ɾ��﹦��)
  Update /*+ rule*/ ���˹Һż�¼ T
  Set ��ת�� = n_����
  Where Rowid In
        (Select Rowid
         From ���˹Һż�¼ R
         Where Not Exists (Select 1 From ������ü�¼ A Where r.No = a.No And a.��¼���� = 4 And a.��ת�� Is Null) And Not Exists
          (Select 1
                From ����ҽ����¼ A
                Where a.�Һŵ� = r.No And a.��ת�� Is Null And a.������Դ <> 4 And Nvl(a.ͣ��ʱ��, a.����ʱ��) >= d_Lastend) And
               Not Exists (Select 1
                From ������ü�¼ E, ����ҽ����¼ A
                Where r.No = a.�Һŵ� And a.Id = e.ҽ����� And a.������Դ <> 4 And e.��ת�� Is Null) And
               r.��ת�� Is Null And r.����ʱ�� < d_End);

  --������һ���ֹҺ�����δת�������ԣ����ܱ�����ݿ�����Һ����ݲ�ƥ��
  Update ���˹ҺŻ��� Set ��ת�� = n_���� Where ��ת�� Is Null And ���� < d_End;
  Update /*+ rule*/ ����ת���¼ Set ��ת�� = n_���� Where NO In (Select NO From ���˹Һż�¼ Where ��ת�� = n_����);

  --ͨ��"סԺ���ü�¼"����ѯ��������"���˽��ʼ�¼",��Ϊ��Ժδ������ʲ���Ҳת���˷���
  --��Ժ����������Ȼ��Ҫ����Ϊ����ĳ�ν���ת���ˣ�������������ת����ֹʱ��֮ǰ��δ��Ժ(һ��סԺ��ν���)��
  --ͨ��ָ��������ʽ���������Ż���ȱʡ����"������ҳIX_��Ժ����"������Ч��̫�ͣ�
  --����"����ת�� is null"����������Ϊһ��סԺ��ν���ʱ������粻ͬ��ת������(ת����ֹʱ��)�����ֶν��ᱻ���¶�Ρ�
  Update /*+ rule*/ ������ҳ P
  Set ��ת�� = n_����
  Where Not Exists
   (Select 1 From סԺ���ü�¼ A Where a.����id = p.����id And a.��ҳid = p.��ҳid And a.��ת�� Is Null) And ��ת�� Is Null And
        ��Ժ���� < d_Lastend And (����id, ��ҳid) In (Select Distinct ����id, ��ҳid From סԺ���ü�¼ Where ��ת�� = n_����);

  --�ѳ�Ժ����û�з��õģ�Ҳ���Ϊת�����Ա�ת����������
  Update /*+ rule*/ ������ҳ P
  Set ��ת�� = n_����
  Where Not Exists (Select 1 From סԺ���ü�¼ A Where a.����id = p.����id And a.��ҳid = p.��ҳid) And ��ת�� Is Null And ����ת�� Is Null And
        ��Ժ���� < d_End;

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

  Update zlDataMovelog
  Set ��ǰ���� = '(5/11)���Ｐ�������ݱ����ɣ����ڱ�ǻ�������'
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

  Update /*+ rule*/ ���˻������
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

  Update zlDataMovelog
  Set ��ǰ���� = '(6/11)�������ݱ����ɣ����ڱ�ǲ�������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  --4.��������
  Update /*+ rule*/ ���Ӳ�����¼
  Set ��ת�� = n_����
  Where ������Դ = 1 And (����id, ��ҳid) In (Select ����id, ID From ���˹Һż�¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ���Ӳ�����¼
  Set ��ת�� = n_����
  Where ������Դ = 2 And (����id, ��ҳid) In (Select ����id, ��ҳid From ������ҳ Where ��ת�� = n_����);

  --�ԵǼ��ಡ��(�޹Һŵ���)
  --����ID�����ظ�����Ϊ���鱨��֮��ģ���ι�����������һ�ű��棬���ڲ���ҽ��������У����ҽ��id��Ӧͬһ����ID
  --Ϊ�������ܣ�����ҽ�����ͼ�¼�ķ���ʱ���ѯ�������þ�ȷ��ʱ�䣬��Ϊֱ�ӵǼǵļ���ҽ����һ�㿪��ʱ���뷢��ʱ������
  --��Щ���⣨�������ݣ��Һŵ�Ϊ�յ�ҽ����������ԴΪ3�ģ�ֱ�ӵǼǵļ�����ҽ����������������ԴΪ1��4�ģ���������ҽ��������ҳID���ܲ���0
  Update /*+ rule*/ ���Ӳ�����¼
  Set ��ת�� = n_����
  Where ��ת�� Is Null And �������� = 7 And ID In (Select c.����id
                                            From ����ҽ����¼ B, ����ҽ������ C
                                            Where c.ҽ��id = b.Id And b.������Դ <> 2 And b.�Һŵ� Is Null And b.���id Is Null And
                                                  b.��ת�� Is Null And b.����ʱ�� < d_End);

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
  Where ����id In (Select ID From ���Ӳ�����¼ Where �������� = 7 And ��ת�� = n_����);

  Update /*+ rule*/ Ӱ�񱨸沵��
  Set ��ת�� = n_����
  Where (ҽ��id, ����id) In (Select ҽ��id, ����id From ����ҽ������ Where ��ת�� = n_����);

  Update /*+ rule*/ ҽ����������
  Set ��ת�� = n_����
  Where ID In (Select ����id From ����ҽ������ Where ��ת�� = n_����);

  Update /*+ rule*/ ������ļ�¼
  Set ��ת�� = n_����
  Where ����id In (Select ID From ���Ӳ�����¼ Where �������� = 7 And ��ת�� = n_����);

  Update /*+ rule*/ �����걨��¼
  Set ��ת�� = n_����
  Where �ļ�id In (Select ID From ���Ӳ�����¼ Where �������� = 5 And ��ת�� = n_����);

  Update /*+ rule*/ �������淴��
  Set ��ת�� = n_����
  Where �ļ�id In (Select ID From ���Ӳ�����¼ Where �������� = 5 And ��ת�� = n_����);

  Update /*+ rule*/ �����걨����
  Set ��ת�� = n_����
  Where �걨id In (Select ID From ���Ӳ�����¼ Where �������� = 5 And ��ת�� = n_����);

  Update /*+ rule*/ ������ϼ�¼
  Set ��ת�� = n_����
  Where ����id In (Select ID From ���Ӳ�����¼ Where ��ת�� = n_����);

  Update zlDataMovelog
  Set ��ǰ���� = '(7/11)�������ݱ����ɣ����ڱ���ٴ�·������'
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
  Update /*+ rule*/ ����·��ҽ������
  Set ��ת�� = n_����
  Where ·��ִ��id In (Select ID From ����·��ִ�� Where ��ת�� = n_����);

  Update zlDataMovelog
  Set ��ǰ���� = '(8/11)�ٴ�·�����ݱ����ɣ����ڱ��ҽ������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  --6.ҽ�������飬���
  --���ϲ�����Դ��������ԴΪ3���ԵǼ��ಡ�������˹Һŵ���ҽ����ת���˶�ҽ������û��ת��
  Update /*+ rule*/ ����ҽ����¼
  Set ��ת�� = n_����
  Where �Һŵ� In (Select NO From ���˹Һż�¼ Where ��ת�� = n_����) And ������Դ = 1;

  --���ϲ�����Դ������ ��ҽ����ת���˶�ҽ������û��ת��
  Update /*+ rule*/ ����ҽ����¼
  Set ��ת�� = n_����
  Where (����id, ��ҳid) In (Select ����id, ��ҳid From ������ҳ Where ��ת�� = n_����) And ������Դ = 2;

  --�ԵǼ��ಡ��(�޹Һŵ�)������ҽ��������ǰ��ת����ʱ��ת��
  Update /*+ rule*/ ����ҽ����¼
  Set ��ת�� = n_����
  Where ��ת�� Is Null And Rowid In (Select b.Rowid
                                  From ����ҽ����¼ B, ����ҽ������ C
                                  Where (b.���id = c.ҽ��id Or b.Id = c.ҽ��id) And c.��ת�� = n_����);

  --�ԵǼ��ಡ��(�޹Һŵ�)��û��ҽ������
  Update /*+ rule*/ ����ҽ����¼ A
  Set ��ת�� = n_����
  Where Not Exists (Select 1 From ����ҽ������ B Where a.Id = b.ҽ��id) And Not Exists
   (Select 1 From ����ҽ������ B Where a.���id = b.ҽ��id) And �Һŵ� Is Null And ������Դ = 3 And ��ת�� Is Null And ����ʱ�� < d_End;

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
  Update /*+ rule*/ ��Ѫ������Ŀ
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

  Update /*+ rule*/ ҽ��ִ�����
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

  Update /*+ rule*/ Ris���ԤԼ
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ �������Լ�¼
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ҽ�����뵥�ļ�
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ����Σ��ֵ��¼
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ����Σ��ֵ����
  Set ��ת�� = n_����
  Where Σ��ֵid In (Select ID From ����Σ��ֵ��¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ����Σ��ֵҽ��
  Set ��ת�� = n_����
  Where Σ��ֵid In (Select ID From ����Σ��ֵ��¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ҩ������˵��
  Set ��ת�� = n_����
  Where ҽ��a In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ҩ������˵��
  Set ��ת�� = n_����
  Where ҽ��b In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update zlDataMovelog
  Set ��ǰ���� = '(9/11)ҽ�����ݱ����ɣ����ڱ�Ǽ���������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  Update /*+ rule*/ Ӱ�����¼
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ Ӱ�񱨸��¼
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ Ӱ�񱨸������¼
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

  Update /*+ rule*/ Ӱ��ԤԼ��¼
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update zlDataMovelog
  Set ��ǰ���� = '(10/11)Ӱ�����ݱ����ɣ����ڱ�Ǽ�������'
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
  Update zlDataMovelog
  Set ��ǰ���� = '(11/11)�������ݱ����ɣ����ڱ�������ٴ�·������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  --11.�����ٴ�·������
  Update /*+ rule*/ ��������·����¼
  Set ��ת�� = n_����
  Where �Һ�id In (Select ID From ���˹Һż�¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ��������·����¼ A
  Set ��ת�� = Null
  Where Exists (Select 1 From ��������·����¼ C Where a.·����¼id = c.·����¼id And c.��ת�� Is Null) And ��ת�� = n_����;

  Update /*+ rule*/ ��������·��
  Set ��ת�� = n_����
  Where ID In (Select ·����¼id From ��������·����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ �������������¼
  Set ��ת�� = n_����
  Where ·����¼id In (Select ID From ��������·�� Where ��ת�� = n_����);

  Update /*+ rule*/ ��������·��ִ��
  Set ��ת�� = n_����
  Where ·����¼id In (Select ID From ��������·�� Where ��ת�� = n_����);

  Update /*+ rule*/ ��������·��ָ��
  Set ��ת�� = n_����
  Where ·����¼id In (Select ID From ��������·�� Where ��ת�� = n_����);

  Update /*+ rule*/ ��������·������
  Set ��ת�� = n_����
  Where ·����¼id In (Select ID From ��������·�� Where ��ת�� = n_����);

  Update /*+ rule*/ ��������·������
  Set ��ת�� = n_����
  Where ·����¼id In (Select ID From ��������·�� Where ��ת�� = n_����);

  Update /*+ rule*/ ��������·��ҽ��
  Set ��ת�� = n_����
  Where ·��ִ��id In (Select ID From ��������·��ִ�� Where ��ת�� = n_����);

  --12.������ҩ�嵥
  Update /*+ rule*/ ������ҩ�嵥
  Set ��ת�� = n_����
  Where (����id, ��ҳid) In (Select ����id, ��ҳid From ������ҳ Where ��ת�� = n_����);

  Update /*+ rule*/ ������ҩ�䷽
  Set ��ת�� = n_����
  Where �䷽id In (Select ID From ������ҩ�嵥 Where ��ת�� = n_����);

  Commit;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl1_Datamove_Tag;
/

--����ZL1_INSIDE_1145/����Ʊ�ݸ�֪��
Insert Into zlReports(ID,����ID,���,����,˵��,����,��ӡ��,��ֽ,Ʊ��,��ӡ��ʽ,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��,��ֹ��ʼʱ��,��ֹ����ʱ��) Values(zlReports_ID.NextVal,Null,'ZL1_INSIDE_1145','����Ʊ�ݸ�֪��','����Ʊ��ר��','J~(f^sl{}+=EpjkvM"QT','Microsoft XPS Document Writer',15,0,0,&n_system,Null,Null,Sysdate,Sysdate,To_Date('2020-05-01 00:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('2020-05-01 00:00:00','YYYY-MM-DD HH24:MI:SS'));
Insert Into zlRPTPuts(����ID,ϵͳ,����ID,����) Values(zlReports_ID.CurrVal,&n_system,1103,'����Ʊ�ݸ�֪��');
Insert Into zlRPTPuts(����ID,ϵͳ,����ID,����) Values(zlReports_ID.CurrVal,&n_system,1107,'����Ʊ�ݸ�֪��');
Insert Into zlRPTPuts(����ID,ϵͳ,����ID,����) Values(zlReports_ID.CurrVal,&n_system,1111,'����Ʊ�ݸ�֪��');
Insert Into zlRPTPuts(����ID,ϵͳ,����ID,����) Values(zlReports_ID.CurrVal,&n_system,1121,'����Ʊ�ݸ�֪��');
Insert Into zlRPTPuts(����ID,ϵͳ,����ID,����) Values(zlReports_ID.CurrVal,&n_system,1124,'����Ʊ�ݸ�֪��');
Insert Into zlRPTPuts(����ID,ϵͳ,����ID,����) Values(zlReports_ID.CurrVal,&n_system,1137,'����Ʊ�ݸ�֪��');
Insert Into zlRPTPuts(����ID,ϵͳ,����ID,����) Values(zlReports_ID.CurrVal,&n_system,1144,'����Ʊ�ݸ�֪��');
Insert Into zlRPTPuts(����ID,ϵͳ,����ID,����) Values(zlReports_ID.CurrVal,&n_system,1145,'����Ʊ�ݸ�֪��');
Insert Into zlRPTFmts(����ID,���,˵��,W,H,ֽ��,ֽ��,��ֽ̬��,ͼ��,�Ƿ�ͣ��,ͣ��ԭ��) Values(zlReports_ID.CurrVal,1,'�ҺŸ�֪��',5874,3855,256,1,0,0,Null,Null);
Insert Into zlRPTFmts(����ID,���,˵��,W,H,ֽ��,ֽ��,��ֽ̬��,ͼ��,�Ƿ�ͣ��,ͣ��ԭ��) Values(zlReports_ID.CurrVal,2,'�շѸ�֪��',7749,4305,256,1,0,0,Null,Null);
Insert Into zlRPTFmts(����ID,���,˵��,W,H,ֽ��,ֽ��,��ֽ̬��,ͼ��,�Ƿ�ͣ��,ͣ��ԭ��) Values(zlReports_ID.CurrVal,3,'���ʸ�֪��',4494,4742,256,1,0,0,Null,Null);
Insert Into zlRPTFmts(����ID,���,˵��,W,H,ֽ��,ֽ��,��ֽ̬��,ͼ��,�Ƿ�ͣ��,ͣ��ԭ��) Values(zlReports_ID.CurrVal,4,'Ԥ����֪��',4464,4455,256,1,0,0,Null,Null);
Insert Into zlRPTFmts(����ID,���,˵��,W,H,ֽ��,ֽ��,��ֽ̬��,ͼ��,�Ƿ�ͣ��,ͣ��ԭ��) Values(zlReports_ID.CurrVal,5,'ҽ�ƿ���֪��',6039,4487,256,1,0,0,Null,Null);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����,˵��,�������ӱ��) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'��ά��','��ά��,205',User||'.����Ʊ�ݶ�ά��',0,Null,Null);
Insert Into zlRPTSQLs(ԴID,�к�,����)  Select zlRPTDatas_ID.CurrVal,a.* From (
Select 1,'select  zltools.Zlbase64.Decode(��ά��) as ��ά�� from ����Ʊ�ݶ�ά�� where ʹ�ü�¼id=[0]' From Dual) a;
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����,����) Values(zlRPTDatas_ID.CurrVal,Null,0,'����Ʊ��ID',1,Null,0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����,˵��,�������ӱ��) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'�Һ�','NO,202|���,139|����,139|����,139|Ӧ�ս��,139|ʵ�ս��,139|����,202|����,202|����,202|�Ա�,202|����,202|����,202|���,202|��λ,202|��ʶ��,139',User||'.���ò����¼,'||User||'.������ü�¼,'||User||'.����Ʊ��ʹ�ü�¼,'||User||'.�շ���ĿĿ¼',0,Null,Null);
Insert Into zlRPTSQLs(ԴID,�к�,����)  Select zlRPTDatas_ID.CurrVal,a.* From (
Select 1,'Select a.No, a.���, Sum(a.����) As ����, Avg(a.��׼����) As ����, Sum(a.Ӧ�ս��) As Ӧ�ս��, Sum(a.ʵ�ս��) As ʵ�ս��, Max(a.����) As ����,' From Dual
Union All Select 2,'       Max(a.����) As ����, Max(a.����) As ����, Max(a.�Ա�) As �Ա�, Max(a.����) As ����, Max(a.����) As ����, Max(a.���) As ���,' From Dual
Union All Select 3,'       Max(a.��λ) As ��λ,Max(��ʶ��) As ��ʶ��' From Dual
Union All Select 4,'From (' From Dual
Union All Select 5,'       --2.�󵥼�' From Dual
Union All Select 6,'       Select a.No, Nvl(a.�۸񸸺�, a.���) As ���, Avg(Nvl(a.����, 1) * a.����) As ����, Sum(a.��׼����) As ��׼����, Sum(a.Ӧ�ս��) As Ӧ�ս��,' From Dual
Union All Select 7,'               Sum(a.ʵ�ս��) As ʵ�ս��, Max(b.����) As ����, Max(b.����) As ����, Max(a.��ʶ��) As ��ʶ��, Max(a.����) As ����, Max(a.�Ա�) As �Ա�,' From Dual
Union All Select 8,'               Max(a.����) As ����, Max(c.����) As ����, Max(c.���) As ���, Max(c.���㵥λ) As ��λ' From Dual
Union All Select 9,'       From ������ü�¼ A,' From Dual
Union All Select 10,'             (' From Dual
Union All Select 11,'               --1.�ҵ�ԭʼ��������е���' From Dual
Union All Select 12,'               Select Distinct a.��¼����, a.No, a.���, d.����, d.Url���� As ����' From Dual
Union All Select 13,'               From ������ü�¼ A, ����Ʊ��ʹ�ü�¼ D' From Dual
Union All Select 14,'               Where a.����id = d.����id And d.Id = [0] And Not Exists (Select 1 From ���ò����¼ Where �շѽ���id = a.����id) And d.Ʊ�� = 4' From Dual
Union All Select 15,'               Union All' From Dual
Union All Select 16,'               Select Distinct a.��¼����, a.No, a.���, d.����, d.Url���� As ����' From Dual
Union All Select 17,'               From ������ü�¼ A, ���ò����¼ B, ����Ʊ��ʹ�ü�¼ D' From Dual
Union All Select 18,'               Where a.����id = b.�շѽ���id And b.����id = d.����id And d.Id = [0] And d.Ʊ�� = 4) B, �շ���ĿĿ¼ C' From Dual
Union All Select 19,'       Where a.No = b.No And Mod(a.��¼����, 10) = b.��¼���� And a.��� = b.��� And a.�շ�ϸĿid = c.Id' From Dual
Union All Select 20,'       Group By a.��¼����, a.��¼״̬, a.No, Nvl(a.�۸񸸺�, a.���)) A' From Dual
Union All Select 21,'Group By a.No, a.���' From Dual
Union All Select 22,'Having Nvl(Sum(a.����), 0) <> 0' From Dual) a;
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����,����) Values(zlRPTDatas_ID.CurrVal,Null,0,'����Ʊ��ID',1,Null,0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����,˵��,�������ӱ��) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'����','�վݷ�Ŀ,202|����,202|�Ա�,202|����,202|�����,139|סԺ��,139|���,139|����,202|����,202',User||'.סԺ���ü�¼,'||User||'.���˽��ʼ�¼,'||User||'.����Ʊ��ʹ�ü�¼,'||User||'.������ü�¼,'||User||'.������Ϣ',0,Null,Null);
Insert Into zlRPTSQLs(ԴID,�к�,����)  Select zlRPTDatas_ID.CurrVal,a.* From (
Select 1,'Select a.�վݷ�Ŀ,b.����, b.�Ա�, b.����,max(b.�����) as �����,max(b.סԺ��) as סԺ��, Sum(a.���) As ���, ' From Dual
Union All Select 2,'       Max(����) As ����,  Max(a.����) As ����' From Dual
Union All Select 3,'From (Select b.����id, a.�վݷ�Ŀ, Sum(a.���ʽ��) As ���, Max(c.Url����) As ����, Max(c.����) As ����' From Dual
Union All Select 4,'       From סԺ���ü�¼ A, ���˽��ʼ�¼ B, ����Ʊ��ʹ�ü�¼ C' From Dual
Union All Select 5,'       Where a.����id = b.Id And b.Id = c.����id And c.Ʊ�� = 3 And c.��¼״̬ = 1 And b.��¼״̬ In (1, 3) And c.Id = [0]' From Dual
Union All Select 6,'       Group By b.����id, a.�վݷ�Ŀ' From Dual
Union All Select 7,'       Union All' From Dual
Union All Select 8,'       Select b.����id, a.�վݷ�Ŀ, Sum(a.���ʽ��) As ���, Max(c.Url����) As ����, Max(c.����) As ����' From Dual
Union All Select 9,'       From ������ü�¼ A, ���˽��ʼ�¼ B, ����Ʊ��ʹ�ü�¼ C' From Dual
Union All Select 10,'       Where a.����id = b.Id And b.Id = c.����id And c.Ʊ�� = 3 And c.��¼״̬ = 1 And b.��¼״̬ In (1, 3) And c.Id = [0]' From Dual
Union All Select 11,'       Group By b.����id, a.�վݷ�Ŀ) A, ������Ϣ B' From Dual
Union All Select 12,'Where a.����id = b.����id' From Dual
Union All Select 13,'Group By  a.�վݷ�Ŀ,b.����, b.�Ա�, b.����' From Dual) a;
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����,����) Values(zlRPTDatas_ID.CurrVal,Null,0,'����Ʊ��ID',1,Null,0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����,˵��,�������ӱ��) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'�շ�','NO,202|���,139|����,139|����,139|Ӧ�ս��,139|ʵ�ս��,139|����,202|����,202|����,202|�Ա�,202|����,202|����,202|���,202|��λ,202|��ʶ��,139',User||'.���ò����¼,'||User||'.������ü�¼,'||User||'.����Ʊ��ʹ�ü�¼,'||User||'.�շ���ĿĿ¼',0,Null,Null);
Insert Into zlRPTSQLs(ԴID,�к�,����)  Select zlRPTDatas_ID.CurrVal,a.* From (
Select 1,'Select a.No, a.���, Sum(a.����) As ����, Avg(a.��׼����) As ����, Sum(a.Ӧ�ս��) As Ӧ�ս��, Sum(a.ʵ�ս��) As ʵ�ս��, Max(a.����) As ����,' From Dual
Union All Select 2,'       Max(a.����) As ����, Max(a.����) As ����, Max(a.�Ա�) As �Ա�, Max(a.����) As ����, Max(a.����) As ����, Max(a.���) As ���,' From Dual
Union All Select 3,'       Max(a.��λ) As ��λ,Max(��ʶ��) As ��ʶ��' From Dual
Union All Select 4,'From (' From Dual
Union All Select 5,'       --2.�󵥼�' From Dual
Union All Select 6,'       Select a.No, Nvl(a.�۸񸸺�, a.���) As ���, Avg(Nvl(a.����, 1) * a.����) As ����, Sum(a.��׼����) As ��׼����, Sum(a.Ӧ�ս��) As Ӧ�ս��,' From Dual
Union All Select 7,'               Sum(a.ʵ�ս��) As ʵ�ս��, Max(b.����) As ����, Max(b.����) As ����, Max(a.��ʶ��) As ��ʶ��, Max(a.����) As ����, Max(a.�Ա�) As �Ա�,' From Dual
Union All Select 8,'               Max(a.����) As ����, Max(c.����) As ����, Max(c.���) As ���, Max(c.���㵥λ) As ��λ' From Dual
Union All Select 9,'       From ������ü�¼ A,' From Dual
Union All Select 10,'             (' From Dual
Union All Select 11,'               --1.�ҵ�ԭʼ��������е���' From Dual
Union All Select 12,'               Select Distinct a.��¼����, a.No, a.���, d.����, d.Url���� As ����' From Dual
Union All Select 13,'               From ������ü�¼ A, ����Ʊ��ʹ�ü�¼ D' From Dual
Union All Select 14,'               Where a.����id = d.����id And d.Id = [0] And Not Exists (Select 1 From ���ò����¼ Where �շѽ���id = a.����id) And d.Ʊ�� = 1' From Dual
Union All Select 15,'               Union All' From Dual
Union All Select 16,'               Select Distinct a.��¼����, a.No, a.���, d.����, d.Url���� As ����' From Dual
Union All Select 17,'               From ������ü�¼ A, ���ò����¼ B, ����Ʊ��ʹ�ü�¼ D' From Dual
Union All Select 18,'               Where a.����id = b.�շѽ���id And b.����id = d.����id And d.Id = [0] And d.Ʊ�� = 1) B, �շ���ĿĿ¼ C' From Dual
Union All Select 19,'       Where a.No = b.No And Mod(a.��¼����, 10) = b.��¼���� And a.��� = b.��� And a.�շ�ϸĿid = c.Id' From Dual
Union All Select 20,'       Group By a.��¼����, a.��¼״̬, a.No, Nvl(a.�۸񸸺�, a.���)) A' From Dual
Union All Select 21,'Group By a.No, a.���' From Dual
Union All Select 22,'Having Nvl(Sum(a.����), 0) <> 0' From Dual) a;
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����,����) Values(zlRPTDatas_ID.CurrVal,Null,0,'����Ʊ��ID',1,Null,0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����,˵��,�������ӱ��) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'ҽ�ƿ�','��ʶ��,131|NO,202|����,202|�Ա�,202|����,202|����,202|����,131|����,139|ʵ�ս��,131|����,202|����,202',User||'.סԺ���ü�¼,'||User||'.�շ���ĿĿ¼,'||User||'.����Ʊ��ʹ�ü�¼',0,Null,Null);
Insert Into zlRPTSQLs(ԴID,�к�,����)  Select zlRPTDatas_ID.CurrVal,a.* From (
Select 1,'Select a.��ʶ��, a.No, a.����, a.�Ա�, a.����, b.����, a.��׼���� As ����, Nvl(a.����, 1) * a.���� As ����, a.ʵ�ս��, ' From Dual
Union All Select 2,'       c.����,c.Url���� As ����' From Dual
Union All Select 3,'From סԺ���ü�¼ A, �շ���ĿĿ¼ B, ����Ʊ��ʹ�ü�¼ C' From Dual
Union All Select 4,'Where a.����Id = c.����id And c.Ʊ�� = 5 And c.��¼״̬ = 1 And a.�շ�ϸĿid = b.Id And a.��¼���� = 5 And a.��¼״̬ = 1 And c.Id = [0]' From Dual) a;
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����,����) Values(zlRPTDatas_ID.CurrVal,Null,0,'����Ʊ��ID',1,Null,0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����,˵��,�������ӱ��) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'Ԥ��','NO,202|���,131|���㷽ʽ,202|����,202|����,202|����,202|�Ա�,202|����,202|�����,131|סԺ��,131',User||'.����Ԥ����¼,'||User||'.����Ʊ��ʹ�ü�¼',0,Null,Null);
Insert Into zlRPTSQLs(ԴID,�к�,����)  Select zlRPTDatas_ID.CurrVal,a.* From (
Select 1,'Select a.No, a.���, a.���㷽ʽ, b.����, b.Url���� As ����, b.����, b.�Ա�, b.����,b.�����,b.סԺ��' From Dual
Union All Select 2,'From ����Ԥ����¼ A, ����Ʊ��ʹ�ü�¼ B' From Dual
Union All Select 3,'Where a.��¼״̬ = 1 And a.��¼���� = 1 And a.Id = b.����id And b.Ʊ�� = 2 And b.Id = [0]' From Dual) a;
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����,����) Values(zlRPTDatas_ID.CurrVal,Null,0,'����Ʊ��ID',1,Null,0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ6',2,Null,0,Null,0,'���ӷ�Ʊ��:[�Һ�.����]',Null,450,720,1980,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ1',2,Null,0,Null,0,'��ʶ��:[�Һ�.��ʶ��]',Null,450,1035,2160,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ3',2,Null,0,Null,0,'�Ա�:[�Һ�.�Ա�]',Null,450,1380,1800,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ5',2,Null,0,Null,0,'�ҺŸ�֪��',Null,1995,255,1650,330,0,0,1,'����',16,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ2',2,Null,0,Null,0,'����:[�Һ�.����]',Null,3450,1035,1800,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ4',2,Null,0,Null,0,'����:[�Һ�.����]',Null,3465,1380,1800,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ7',2,Null,0,Null,0,'[��ά��.��ά��]',Null,4215,30,1390,1200,0,0,1,'����',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'�����1',4,Null,0,Null,0,'�Һ�',Null,465,1755,5030,1395,255,0,0,'����',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-1,0,Null,Null,'[�Һ�.NO]','4^225^NO',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-2,1,Null,Null,'[�Һ�.����]','4^225^����',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-3,2,Null,Null,'[�Һ�.����]','4^225^����',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-4,3,Null,Null,'[�Һ�.����]','4^225^����',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-5,4,Null,Null,'[�Һ�.ʵ�ս��]','4^225^ʵ�ս��',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,'��ǩ5',2,Null,0,Null,0,'���ӷ�Ʊ��:[�շ�.����]',Null,435,825,1980,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,'��ǩ2',2,Null,0,Null,0,'����:[�շ�.����]',Null,450,1200,1440,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,'�շѸ�֪��',2,Null,0,Null,0,'�շѸ�֪��',Null,2520,210,1650,330,0,0,1,'����',16,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,'��ǩ1',2,Null,0,Null,0,'��־��:[�շ�.��ʶ��]',Null,3000,810,1800,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,'��ǩ3',2,Null,0,Null,0,'�Ա�:[�շ�.�Ա�]',Null,3030,1140,1440,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,'��ǩ4',2,Null,0,Null,0,'����:[�շ�.����]',Null,5790,1155,1440,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,'��ǩ6',2,Null,0,Null,0,'[��ά��.��ά��]',Null,5835,30,1375,1125,0,0,0,'����',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,'�����1',4,Null,0,Null,0,'�շ�',Null,420,1590,7030,2475,255,0,0,'����',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,Null,6,zlRPTItems_ID.CurrVal-1,0,Null,Null,'[�շ�.NO]','4^300^NO',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,Null,6,zlRPTItems_ID.CurrVal-2,1,Null,Null,'[�շ�.����]','4^300^����',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,Null,6,zlRPTItems_ID.CurrVal-3,2,Null,Null,'[�շ�.���]','4^300^���',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,Null,6,zlRPTItems_ID.CurrVal-4,3,Null,Null,'[�շ�.��λ]','4^300^��λ',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,Null,6,zlRPTItems_ID.CurrVal-5,4,Null,Null,'[�շ�.����]','4^300^����',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,Null,6,zlRPTItems_ID.CurrVal-6,5,Null,Null,'[�շ�.����]','4^300^����',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,Null,6,zlRPTItems_ID.CurrVal-7,6,Null,Null,'[�շ�.ʵ�ս��]','4^300^ʵ�ս��',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,3,'��ǩ7',2,Null,0,Null,0,'���ӷ�Ʊ��:[����.����]',Null,375,855,1980,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,3,'��ǩ3',2,Null,0,Null,0,'����:[����.����]',Null,375,1245,2160,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,3,'��ǩ5',2,Null,0,Null,0,'����:[����.����]',Null,375,1605,1440,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,3,'��ǩ2',2,Null,0,Null,0,'�����:[����.�����]',Null,375,1995,1800,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,3,'��ǩ1',2,Null,0,Null,0,'���ʸ�֪��',Null,1215,255,1650,330,0,0,1,'����',16,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,3,'��ǩ4',2,Null,0,Null,0,'�Ա�:[����.�Ա�]',Null,2265,1230,1440,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,3,'��ǩ6',2,Null,0,Null,0,'סԺ��:[����.סԺ��]',Null,2315,1995,1800,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,3,'��ǩ8',2,Null,0,Null,0,'[��ά��.��ά��]',Null,2940,150,1345,1080,0,0,0,'����',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,3,'�����1',4,Null,0,Null,0,'����',Null,405,2385,3601,1935,255,0,0,'����',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,3,Null,6,zlRPTItems_ID.CurrVal-1,0,Null,Null,'[����.�վݷ�Ŀ]','4^330^�վݷ�Ŀ',0,0,1860,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,3,Null,6,zlRPTItems_ID.CurrVal-2,1,Null,Null,'[����.���]','4^330^���',0,0,1485,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,4,'��ǩ8',2,Null,0,Null,0,'���ӷ�Ʊ��:[Ԥ��.����]',Null,435,855,1980,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,4,'��ǩ3',2,Null,0,Null,0,'����:[Ԥ��.����]',Null,435,1200,1440,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,4,'��ǩ5',2,Null,0,Null,0,'����:[Ԥ��.����]',Null,435,1530,1440,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,4,'��ǩ2',2,Null,0,Null,0,'�����:[Ԥ��.�����]',Null,435,1845,1800,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,4,'��ǩ1',2,Null,0,Null,0,'Ԥ����֪��',Null,1305,315,1650,330,0,0,1,'����',16,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,4,'��ǩ6',2,Null,0,Null,0,'סԺ��:[Ԥ��.סԺ��]',Null,2375,1845,1800,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,4,'��ǩ4',2,Null,0,Null,0,'�Ա�:[Ԥ��.�Ա�]',Null,2400,1200,1440,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,4,'��ǩ9',2,Null,0,Null,0,'[��ά��.��ά��]',Null,2910,75,1300,1245,0,0,0,'����',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,4,'�����1',4,Null,0,Null,0,'Ԥ��',Null,450,2235,3480,1770,255,0,0,'����',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,4,Null,6,zlRPTItems_ID.CurrVal-1,0,Null,Null,'[Ԥ��.NO]','4^270^NO',0,0,1110,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,4,Null,6,zlRPTItems_ID.CurrVal-2,1,Null,Null,'[Ԥ��.���]','4^270^���',0,0,1230,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,4,Null,6,zlRPTItems_ID.CurrVal-3,2,Null,Null,'[Ԥ��.���㷽ʽ]','4^270^���㷽ʽ',0,0,1050,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,5,'��ǩ6',2,Null,0,Null,0,'���ӷ�Ʊ��:[ҽ�ƿ�.����]',Null,495,890,2160,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,5,'��ǩ2',2,Null,0,Null,0,'��ʶ��:[ҽ�ƿ�.��ʶ��]',Null,495,1320,1980,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,5,'��ǩ4',2,Null,0,Null,0,'�Ա�:[ҽ�ƿ�.�Ա�]',Null,495,1785,1620,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,5,'��ǩ1',2,Null,0,Null,0,'ҽ�ƿ���֪��',Null,1935,285,1980,330,0,0,1,'����',16,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,5,'��ǩ3',2,Null,0,Null,0,'����:[ҽ�ƿ�.����]',Null,3705,1320,1620,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,5,'��ǩ5',2,Null,0,Null,0,'����:[ҽ�ƿ�.����]',Null,3705,1785,1620,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,5,'��ǩ7',2,Null,0,Null,0,'[��ά��.��ά��]',Null,4125,150,1300,1110,0,0,0,'����',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,5,'�����1',4,Null,0,Null,0,'ҽ�ƿ�',Null,495,2280,5030,1545,255,0,0,'����',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,5,Null,6,zlRPTItems_ID.CurrVal-1,0,Null,Null,'[ҽ�ƿ�.NO]','4^225^NO',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,5,Null,6,zlRPTItems_ID.CurrVal-2,1,Null,Null,'[ҽ�ƿ�.����]','4^225^����',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,5,Null,6,zlRPTItems_ID.CurrVal-3,2,Null,Null,'[ҽ�ƿ�.����]','4^225^����',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,5,Null,6,zlRPTItems_ID.CurrVal-4,3,Null,Null,'[ҽ�ƿ�.����]','4^225^����',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת,��ֵ�Ԫ��,�Զ����,�������) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,5,Null,6,zlRPTItems_ID.CurrVal-5,4,Null,Null,'[ҽ�ƿ�.ʵ�ս��]','4^225^ʵ�ս��',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);

--����ZL1_INSIDE_1145/����Ʊ�ݸ�֪��
Insert into zlProgFuncs(ϵͳ,���,����,˵��) Values(n_System,1103,'����Ʊ�ݸ�֪��','����Ʊ��ר��');
Insert into zlProgFuncs(ϵͳ,���,����,˵��) Values(n_System,1107,'����Ʊ�ݸ�֪��','����Ʊ��ר��');
Insert into zlProgFuncs(ϵͳ,���,����,˵��) Values(n_System,1111,'����Ʊ�ݸ�֪��','����Ʊ��ר��');
Insert into zlProgFuncs(ϵͳ,���,����,˵��) Values(n_System,1121,'����Ʊ�ݸ�֪��','����Ʊ��ר��');
Insert into zlProgFuncs(ϵͳ,���,����,˵��) Values(n_System,1124,'����Ʊ�ݸ�֪��','����Ʊ��ר��');
Insert into zlProgFuncs(ϵͳ,���,����,˵��) Values(n_System,1137,'����Ʊ�ݸ�֪��','����Ʊ��ר��');
Insert into zlProgFuncs(ϵͳ,���,����,˵��) Values(n_System,1144,'����Ʊ�ݸ�֪��','����Ʊ��ר��');
Insert into zlProgFuncs(ϵͳ,���,����,˵��) Values(n_System,1145,'����Ʊ�ݸ�֪��','����Ʊ��ר��');
Insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
select &n_system,1103,'����Ʊ�ݸ�֪��',User,'���˽��ʼ�¼','SELECT' From Dual
Union All select &n_system,1103,'����Ʊ�ݸ�֪��',User,'����Ԥ����¼','SELECT' From Dual
Union All select &n_system,1103,'����Ʊ�ݸ�֪��',User,'����Ʊ�ݶ�ά��','SELECT' From Dual
Union All select &n_system,1103,'����Ʊ�ݸ�֪��',User,'����Ʊ��ʹ�ü�¼','SELECT' From Dual
Union All select &n_system,1103,'����Ʊ�ݸ�֪��',User,'���ò����¼','SELECT' From Dual
Union All select &n_system,1103,'����Ʊ�ݸ�֪��',User,'������ü�¼','SELECT' From Dual
Union All select &n_system,1103,'����Ʊ�ݸ�֪��',User,'�շ���ĿĿ¼','SELECT' From Dual
Union All select &n_system,1103,'����Ʊ�ݸ�֪��',User,'סԺ���ü�¼','SELECT' From Dual
Union All select &n_system,1107,'����Ʊ�ݸ�֪��',User,'���˽��ʼ�¼','SELECT' From Dual
Union All select &n_system,1107,'����Ʊ�ݸ�֪��',User,'����Ԥ����¼','SELECT' From Dual
Union All select &n_system,1107,'����Ʊ�ݸ�֪��',User,'����Ʊ�ݶ�ά��','SELECT' From Dual
Union All select &n_system,1107,'����Ʊ�ݸ�֪��',User,'����Ʊ��ʹ�ü�¼','SELECT' From Dual
Union All select &n_system,1107,'����Ʊ�ݸ�֪��',User,'���ò����¼','SELECT' From Dual
Union All select &n_system,1107,'����Ʊ�ݸ�֪��',User,'������ü�¼','SELECT' From Dual
Union All select &n_system,1107,'����Ʊ�ݸ�֪��',User,'סԺ���ü�¼','SELECT' From Dual
Union All select &n_system,1111,'����Ʊ�ݸ�֪��',User,'���˽��ʼ�¼','SELECT' From Dual
Union All select &n_system,1111,'����Ʊ�ݸ�֪��',User,'����Ԥ����¼','SELECT' From Dual
Union All select &n_system,1111,'����Ʊ�ݸ�֪��',User,'����Ʊ�ݶ�ά��','SELECT' From Dual
Union All select &n_system,1111,'����Ʊ�ݸ�֪��',User,'����Ʊ��ʹ�ü�¼','SELECT' From Dual
Union All select &n_system,1111,'����Ʊ�ݸ�֪��',User,'���ò����¼','SELECT' From Dual
Union All select &n_system,1111,'����Ʊ�ݸ�֪��',User,'������ü�¼','SELECT' From Dual
Union All select &n_system,1111,'����Ʊ�ݸ�֪��',User,'סԺ���ü�¼','SELECT' From Dual
Union All select &n_system,1121,'����Ʊ�ݸ�֪��',User,'���˽��ʼ�¼','SELECT' From Dual
Union All select &n_system,1121,'����Ʊ�ݸ�֪��',User,'����Ԥ����¼','SELECT' From Dual
Union All select &n_system,1121,'����Ʊ�ݸ�֪��',User,'����Ʊ�ݶ�ά��','SELECT' From Dual
Union All select &n_system,1121,'����Ʊ�ݸ�֪��',User,'����Ʊ��ʹ�ü�¼','SELECT' From Dual
Union All select &n_system,1121,'����Ʊ�ݸ�֪��',User,'���ò����¼','SELECT' From Dual
Union All select &n_system,1121,'����Ʊ�ݸ�֪��',User,'������ü�¼','SELECT' From Dual
Union All select &n_system,1121,'����Ʊ�ݸ�֪��',User,'סԺ���ü�¼','SELECT' From Dual
Union All select &n_system,1124,'����Ʊ�ݸ�֪��',User,'���˽��ʼ�¼','SELECT' From Dual
Union All select &n_system,1124,'����Ʊ�ݸ�֪��',User,'����Ԥ����¼','SELECT' From Dual
Union All select &n_system,1124,'����Ʊ�ݸ�֪��',User,'����Ʊ�ݶ�ά��','SELECT' From Dual
Union All select &n_system,1124,'����Ʊ�ݸ�֪��',User,'����Ʊ��ʹ�ü�¼','SELECT' From Dual
Union All select &n_system,1124,'����Ʊ�ݸ�֪��',User,'���ò����¼','SELECT' From Dual
Union All select &n_system,1124,'����Ʊ�ݸ�֪��',User,'������ü�¼','SELECT' From Dual
Union All select &n_system,1124,'����Ʊ�ݸ�֪��',User,'סԺ���ü�¼','SELECT' From Dual
Union All select &n_system,1137,'����Ʊ�ݸ�֪��',User,'���˽��ʼ�¼','SELECT' From Dual
Union All select &n_system,1137,'����Ʊ�ݸ�֪��',User,'����Ԥ����¼','SELECT' From Dual
Union All select &n_system,1137,'����Ʊ�ݸ�֪��',User,'����Ʊ�ݶ�ά��','SELECT' From Dual
Union All select &n_system,1137,'����Ʊ�ݸ�֪��',User,'����Ʊ��ʹ�ü�¼','SELECT' From Dual
Union All select &n_system,1137,'����Ʊ�ݸ�֪��',User,'���ò����¼','SELECT' From Dual
Union All select &n_system,1137,'����Ʊ�ݸ�֪��',User,'������ü�¼','SELECT' From Dual
Union All select &n_system,1137,'����Ʊ�ݸ�֪��',User,'סԺ���ü�¼','SELECT' From Dual
Union All select &n_system,1144,'����Ʊ�ݸ�֪��',User,'���˽��ʼ�¼','SELECT' From Dual
Union All select &n_system,1144,'����Ʊ�ݸ�֪��',User,'������Ϣ','SELECT' From Dual
Union All select &n_system,1144,'����Ʊ�ݸ�֪��',User,'����Ԥ����¼','SELECT' From Dual
Union All select &n_system,1144,'����Ʊ�ݸ�֪��',User,'����Ʊ�ݶ�ά��','SELECT' From Dual
Union All select &n_system,1144,'����Ʊ�ݸ�֪��',User,'����Ʊ��ʹ�ü�¼','SELECT' From Dual
Union All select &n_system,1144,'����Ʊ�ݸ�֪��',User,'���ò����¼','SELECT' From Dual
Union All select &n_system,1144,'����Ʊ�ݸ�֪��',User,'������ü�¼','SELECT' From Dual
Union All select &n_system,1144,'����Ʊ�ݸ�֪��',User,'�շ���ĿĿ¼','SELECT' From Dual
Union All select &n_system,1144,'����Ʊ�ݸ�֪��',User,'סԺ���ü�¼','SELECT' From Dual
Union All select &n_system,1145,'����Ʊ�ݸ�֪��',User,'���˽��ʼ�¼','SELECT' From Dual
Union All select &n_system,1145,'����Ʊ�ݸ�֪��',User,'������Ϣ','SELECT' From Dual
Union All select &n_system,1145,'����Ʊ�ݸ�֪��',User,'����Ԥ����¼','SELECT' From Dual
Union All select &n_system,1145,'����Ʊ�ݸ�֪��',User,'����Ʊ�ݶ�ά��','SELECT' From Dual
Union All select &n_system,1145,'����Ʊ�ݸ�֪��',User,'����Ʊ��ʹ�ü�¼','SELECT' From Dual
Union All select &n_system,1145,'����Ʊ�ݸ�֪��',User,'���ò����¼','SELECT' From Dual
Union All select &n_system,1145,'����Ʊ�ݸ�֪��',User,'������ü�¼','SELECT' From Dual
Union All select &n_system,1145,'����Ʊ�ݸ�֪��',User,'�շ���ĿĿ¼','SELECT' From Dual
Union All select &n_system,1145,'����Ʊ�ݸ�֪��',User,'סԺ���ü�¼','SELECT' From Dual;


