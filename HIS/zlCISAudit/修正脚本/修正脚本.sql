
-----------------------------------------------------------------------------------------------------------------------
--    ���Ӳ������鵵   ��ռ�:zl9CISAduit
-----------------------------------------------------------------------------------------------------------------------
--���Ӳ������鵵
Create Sequence �����ύ��¼_ID start with 1;
Create Sequence ����������¼_ID start with 1;
Create Sequence �������ļ�¼_ID start with 1;
Create Sequence ��������¼_ID start with 1;
Create Sequence �������ַ���_ID start with 1;
Create Sequence �������ֱ�׼_ID start with 1;
Create Sequence �������ֽ��_ID start with 1;
Create Sequence ����������ϸ_ID start with 1;

Create Table �����ύ��¼(
    ID			Number(18),
    ����id		Number(18),
    ��ҳid		Number(5),
    ��¼״̬	Number(3),
    �ύ��		Varchar2(20),
    �ύʱ��	Date,
    ������		Varchar2(20),
    ����ʱ��	Date,
    �鵵��		Varchar2(20),
    �鵵ʱ��	Date,
    ������		Varchar2(20),
    ����ʱ��	Date,
    ��������	Varchar2(255))
    TableSpace zl9CISAduit
    PCTFREE 5 PCTUSED 85;

Create Table ����������ǩ(
    �ύid		Number(18),
    ���Ķ���	Number(3),
    �ļ�id		Number(18),
    ����ʱ��	Date)
    TableSpace zl9CISAduit
    PCTFREE 5 PCTUSED 85;

Create Table ����������¼(
    ID			Number(18),
    ���id		Number(18),
    �ύid		Number(18),
    ����id		Number(18),
    ��ҳid		Number(5),
    ��������	Number(3),
    �ļ�id		Number(18),
    ��¼����	Number(3),
    ��¼״̬	Number(3),
    �������	Varchar2(255),
    ������Ŀid	Number(18),
    ������		Varchar2(20),
    ����ʱ��	Date,
    ��������	Date,
    ����˵��	Varchar2(255),
    ������		Varchar2(20),
    ����ʱ��	Date)
    TableSpace zl9CISAduit
    PCTFREE 5 PCTUSED 85;

Create Table ����������ʷ(
    ID			Number(18),
    ���id		Number(18),
    �ύid		Number(18),
    ����id		Number(18),
    ��ҳid		Number(5),
    ��������	Number(3),
    �ļ�id		Number(18),
    ��¼����	Number(3),
    ��¼״̬	Number(3),
    �������	Varchar2(255),
    ������Ŀid	Number(18),
    ������		Varchar2(20),
    ����ʱ��	Date,
    ��������	Date,
    ����˵��	Varchar2(255),
    ������		Varchar2(20),
    ����ʱ��	Date)
    TableSpace zl9CISAduit
    PCTFREE 5 PCTUSED 85;

Create Table �������ļ�¼(
	ID		Number(18),
	No		Varchar2(10),
	��¼״̬	Number(3),   
	������	Varchar2(20),	
	��������	Varchar2(255),
	����ʱ��	Date,
	��������	Date,
	����ʱ��	Date,
	��������	Date,
	��׼��	Varchar2(20),
	��׼ʱ��	Date,
	�ܽ�����	Varchar2(255),
	�ܽ���	Varchar2(20),
	�ܽ�ʱ��	Date,
	�Ǽ�ʱ��	Date)
	TABLESPACE zl9CISAduit
	PCTFREE 5 PCTUSED 85;

Create Table ������������(
    ����id		Number(18),
    ����id		Number(18),
    ��ҳid		Number(5))
    TableSpace zl9CISAduit
    PCTFREE 5 PCTUSED 85;

Create Table ����������Ա(
    ����id		Number(18),
    ��Աid		Number(18))
    TableSpace zl9CISAduit
    PCTFREE 5 PCTUSED 85;

Create Table ��������¼(
    ID			Number(18),
    ����id		Number(18),
    ��ҳid		Number(5),
    ��¼״̬	Number(3),
    �����		Varchar2(20),
    ���ʱ��	Date,
    �������	Varchar2(255),
    �����		Varchar2(20),
    ���ʱ��	Date)
    TableSpace zl9CISAduit
    PCTFREE 5 PCTUSED 85;

--�������ֲ���
Create Table �������ַ���(
	ID number(18) not null,
	���� varchar2(50),
	�ܷ� number(8,2) default 100,
	��ֵ number(8,2),
	��ֵ number(8,2),
	���� varchar2(10),
	���� varchar2(10),
	ѡ�� number(1) default 0,
	����ʱ�� Date,
	ͣ��ʱ�� Date)
    	TABLESPACE zl9CISAduit
    	PCTFREE 5  PCTUSED 85;

Create Table �������ֱ�׼(
	ID number(18) not null,
	�ϼ�ID number(18),
	����ID number(18),
	���� varchar2(50),
	���� varchar2(4000),
	��׼��ֵ number(8,2),
	ȱ�ݵȼ� varchar2(2),
	���ֵ�λ varchar2(8),
	�ϼ���� NUMBER(18),
	��� NUMBER(18))
    	TABLESPACE zl9CISAduit
    	PCTFREE 5  PCTUSED 85;

Create Table �������ֽ��(
	ID number(18) not null,
	����ID number(18),
	��ҳID number(5),
	����ID number(18),
	�ܷ� number(8,2),
	�ȼ� varchar2(2),
	�����޸� number(1),
	��ע	varchar(50),
	������ varchar2(20),
	����ʱ�� Date,
	����� varchar2(20),
	���ʱ�� Date)
	TABLESPACE zl9CISAduit
	PCTFREE 10 PCTUSED 80;

Create Table ����������ϸ(
	ID number(18) not null,
	����ID number(18),
	���ֱ�׼ID number(18),
	������� number(8,2),
	ȱ�ݵȼ� varchar2(2),
	�ɷ��޸� Number(1) Default 0,
	��ע	varchar(50))
	TABLESPACE zl9CISAduit
	PCTFREE 10 PCTUSED 80;


--�޸����еı�
Alter Table ������ҳ Add ����״̬ Number(3);
Alter Table ������ҳ Add ���ʱ�� Date;

--���Ӳ������鵵
Alter Table �����ύ��¼ Add Constraint �����ύ��¼_PK Primary Key (ID) Using Index Pctfree 0 Tablespace zl9indexcis;
Alter Table �����ύ��¼ Add Constraint �����ύ��¼_CK_��¼״̬ Check (��¼״̬ IN(1,2,3,4,5));
Alter Table �����ύ��¼ Add Constraint �����ύ��¼_FK_����ID Foreign Key (����ID) References ������Ϣ(����ID);
Alter Table �����ύ��¼ Add Constraint �����ύ��¼_FK_��ҳID Foreign Key (����ID,��ҳID) References ������ҳ(����ID,��ҳID);
Alter Table ����������ǩ Add Constraint ����������ǩ_PK Primary Key (�ύID,���Ķ���,�ļ�ID) Using Index Pctfree 0 Tablespace zl9indexcis;
Alter Table ����������ǩ Add Constraint ����������ǩ_FK_�ύID Foreign Key (�ύID) References �����ύ��¼(ID) On Delete Cascade;
Alter Table ����������ǩ Add Constraint ����������ǩ_CK_���Ķ��� Check (���Ķ��� IN(1,2,3,4,5,6,7,8));
Alter Table ����������¼ Add Constraint ����������¼_PK Primary Key (ID) Using Index Pctfree 0 Tablespace zl9indexcis;
Alter Table ����������¼ Add Constraint ����������¼_FK_���ID Foreign Key (���ID) References ����������¼(ID) On Delete Cascade;
Alter Table ����������¼ Add Constraint ����������¼_FK_�ύID Foreign Key (�ύID) References �����ύ��¼(ID);
Alter Table ����������¼ Add Constraint ����������¼_CK_��¼���� Check (��¼���� IN(1,2));
Alter Table ����������¼ Add Constraint ����������¼_CK_��¼״̬ Check (��¼״̬ IN(1,2,3));
Alter Table ����������¼ Add Constraint ����������¼_CK_�������� Check (�������� IN(1,2,3,4,5,6,7,8));
Alter Table ����������¼ Add Constraint ����������¼_FK_����ID Foreign Key (����ID) References ������Ϣ(����ID);
Alter Table ����������¼ Add Constraint ����������¼_FK_��ҳID Foreign Key (����ID,��ҳID) References ������ҳ(����ID,��ҳID);
Alter Table ����������ʷ Add Constraint ����������ʷ_PK Primary Key (ID) Using Index Pctfree 0 Tablespace zl9indexcis;
Alter Table ����������ʷ Add Constraint ����������ʷ_FK_���ID Foreign Key (���ID) References ����������ʷ(ID) On Delete Cascade;
Alter Table ����������ʷ Add Constraint ����������ʷ_FK_�ύID Foreign Key (�ύID) References �����ύ��¼(ID);
Alter Table ����������ʷ Add Constraint ����������ʷ_CK_��¼���� Check (��¼���� IN(1,2));
Alter Table ����������ʷ Add Constraint ����������ʷ_CK_��¼״̬ Check (��¼״̬ IN(1,2,3));
Alter Table ����������ʷ Add Constraint ����������ʷ_CK_�������� Check (�������� IN(1,2,3,4,5,6,7,8));
Alter Table ����������ʷ Add Constraint ����������ʷ_FK_����ID Foreign Key (����ID) References ������Ϣ(����ID);
Alter Table ����������ʷ Add Constraint ����������ʷ_FK_��ҳID Foreign Key (����ID,��ҳID) References ������ҳ(����ID,��ҳID);

Alter Table �������ļ�¼ Add Constraint �������ļ�¼_PK Primary Key (ID) Using Index Pctfree 0 Tablespace zl9indexcis;
Alter Table �������ļ�¼ Add Constraint �������ļ�¼_CK_��¼״̬ Check (��¼״̬ IN(1,2,3));
Alter Table ������������ Add Constraint ������������_PK Primary Key (����ID,����ID,��ҳID) Using Index Pctfree 0 Tablespace zl9indexcis;
Alter Table ������������ Add Constraint ������������_FK_����ID Foreign Key (����ID) References �������ļ�¼(ID) On Delete Cascade;
Alter Table ������������ Add Constraint ������������_FK_����ID Foreign Key (����ID) References ������Ϣ(����ID);
Alter Table ������������ Add Constraint ������������_FK_��ҳID Foreign Key (����ID,��ҳID) References ������ҳ(����ID,��ҳID);
Alter Table ����������Ա Add Constraint ����������Ա_PK Primary Key (����ID,��ԱID) Using Index Pctfree 0 Tablespace zl9indexcis;
Alter Table ����������Ա Add Constraint ����������Ա_FK_����ID Foreign Key (����ID) References �������ļ�¼(ID) On Delete Cascade;
Alter Table ��������¼ Add Constraint ��������¼_PK Primary Key (ID) Using Index Pctfree 0 Tablespace zl9indexcis;
Alter Table ��������¼ Add Constraint ��������¼_CK_��¼״̬ Check (��¼״̬ IN(1,2));
Alter Table ��������¼ Add Constraint ��������¼_FK_����ID Foreign Key (����ID) References ������Ϣ(����ID);
Alter Table ��������¼ Add Constraint ��������¼_FK_��ҳID Foreign Key (����ID,��ҳID) References ������ҳ(����ID,��ҳID);

--�������ֲ���
ALTER TABLE �������ַ��� ADD CONSTRAINT �������ַ���_PK PRIMARY KEY (ID) USING INDEX PCTFREE 5 TABLESPACE zl9indexcis;
ALTER TABLE �������ַ��� Add CONSTRAINT �������ַ���_CK_ѡ�� CHECK (ѡ�� IN(0,1));
ALTER TABLE �������ֱ�׼ ADD CONSTRAINT �������ֱ�׼_PK PRIMARY KEY (ID) USING INDEX PCTFREE 5 TABLESPACE zl9indexcis;
ALTER TABLE �������ֱ�׼ ADD CONSTRAINT �������ֱ�׼_FK_�ϼ�ID FOREIGN KEY (�ϼ�ID) REFERENCES �������ֱ�׼(ID) ON DELETE CASCADE;
ALTER TABLE �������ֱ�׼ ADD CONSTRAINT �������ֱ�׼_FK_����ID FOREIGN KEY (����ID) REFERENCES �������ַ���(ID) ON DELETE CASCADE;
ALTER TABLE �������ֽ�� ADD CONSTRAINT �������ֽ��_PK PRIMARY KEY (ID) USING INDEX PCTFREE 5 TABLESPACE zl9indexcis;
ALTER TABLE �������ֽ�� Add CONSTRAINT �������ֽ��_UQ_����ID_��ҳID UNIQUE (����ID,��ҳID) USING INDEX PCTFREE 5 TABLESPACE zl9indexcis;
ALTER TABLE �������ֽ�� ADD CONSTRAINT �������ֽ��_FK_����ID_��ҳID FOREIGN KEY (����ID,��ҳID) REFERENCES ������ҳ(����ID,��ҳID) ON DELETE CASCADE;
ALTER TABLE �������ֽ�� ADD CONSTRAINT �������ֽ��_FK_����ID FOREIGN KEY (����ID) REFERENCES �������ַ���(ID) ON DELETE CASCADE;
ALTER TABLE ����������ϸ ADD CONSTRAINT ����������ϸ_PK PRIMARY KEY (ID) USING INDEX PCTFREE 5 TABLESPACE zl9indexcis;
ALTER TABLE ����������ϸ ADD CONSTRAINT ����������ϸ_FK_���ֱ�׼ID FOREIGN KEY (���ֱ�׼ID) REFERENCES �������ֱ�׼(ID) ON DELETE CASCADE;
ALTER TABLE ����������ϸ ADD CONSTRAINT ����������ϸ_FK_����ID FOREIGN KEY (����ID) REFERENCES �������ֽ��(ID) ON DELETE CASCADE;
ALTER TABLE ����������ϸ Add CONSTRAINT ����������ϸ_CK_�ɷ��޸� CHECK (�ɷ��޸� IN(0,1));

-----------------------------------------------------------------------------------------------------------------------
---���Ӳ������鵵
-----------------------------------------------------------------------------------------------------------------------
Create Index �����ύ��¼_IX_����id On �����ύ��¼(����id) Pctfree 5 Tablespace zl9indexcis
/
Create Index ����������ǩ_IX_�ύid On ����������ǩ(�ύid) Pctfree 5 Tablespace zl9indexcis
/
Create Index ����������¼_IX_�ύid On ����������¼(�ύid) Pctfree 5 Tablespace zl9indexcis
/
Create Index ����������¼_IX_���id On ����������¼(���id) Pctfree 5 Tablespace zl9indexcis
/
Create Index ����������¼_IX_����ʱ�� On ����������¼(����ʱ��) Pctfree 5 Tablespace zl9indexcis
/
Create Index ����������¼_IX_����ʱ�� On ����������¼(����ʱ��) Pctfree 5 Tablespace zl9indexcis
/
Create Index ����������ʷ_IX_�ύid On ����������ʷ(�ύid) Pctfree 5 Tablespace zl9indexcis
/
Create Index ����������ʷ_IX_���id On ����������ʷ(���id) Pctfree 5 Tablespace zl9indexcis
/
Create Index ����������ʷ_IX_����ʱ�� On ����������ʷ(����ʱ��) Pctfree 5 Tablespace zl9indexcis
/
Create Index ����������ʷ_IX_����ʱ�� On ����������ʷ(����ʱ��) Pctfree 5 Tablespace zl9indexcis
/
Create Index ������������_IX_����id On ������������(����id) Pctfree 5 Tablespace zl9indexcis
/
Create Index ������������_IX_����id On ������������(����id) Pctfree 5 Tablespace zl9indexcis
/
Create Index ����������Ա_IX_����id On ����������Ա(����id) Pctfree 5 Tablespace zl9indexcis
/

--�������ֲ���
Create Index �������ֱ�׼_IX_����ID on �������ֱ�׼(����ID) PCTFREE 5 TABLESPACE zl9indexcis
/
Create Index �������ֱ�׼_IX_�ϼ�ID on �������ֱ�׼(�ϼ�ID) PCTFREE 5 TABLESPACE zl9indexcis
/
Create Index �������ֽ��_IX_����ID on �������ֽ��(����ID) PCTFREE 5 TABLESPACE zl9indexcis
/
Create Index ����������ϸ_IX_���ID on ����������ϸ(����ID) PCTFREE 5 TABLESPACE zl9indexcis
/
Create Index ����������ϸ_IX_���ֱ�׼ID on ����������ϸ(���ֱ�׼ID) PCTFREE 5 TABLESPACE zl9indexcis
/


--��ͼ����
create or replace view �������ֱ�׼��ͼ as
select decode(T.�ϼ����,null,���,T.�ϼ����) as �ϼ����, decode(T.���,null,T.ID,T.���) as ���,T.ID,T.�ϼ�ID,T.����ID,T.��Ŀ,T.��׼��ֵ,T.����Ҫ��,T.ȱ������,T.�۷ֱ�׼,decode(T.�������,0,'��','��') as ����
from
(
  select B.�ϼ����,A.���,A.����ID,
  A.ID,
  A.�ϼ�ID,
  decode(A.�������,0,decode(A.�ϼ�ID,Null,A.����,B.����),A.����) as ��Ŀ,
  decode(A.�������,0,decode(A.�ϼ�ID,Null,A.��׼��ֵ,B.��׼��ֵ),B.��׼��ֵ) as ��׼��ֵ,
  decode(A.�������,0,decode(A.�ϼ�ID,Null,A.����,B.����),A.����) as ����Ҫ��,
  A.���� as ȱ������,
  DECODE(A.ȱ�ݵȼ�,NULL,decode(sign(A.��׼��ֵ-1),-1,To_CHAR(A.��׼��ֵ,'0.9'),To_Char(A.��׼��ֵ))||decode(A.���ֵ�λ,NULL,'','/'||A.���ֵ�λ),A.ȱ�ݵȼ�) as �۷ֱ�׼,
  A.�������
  from
      (
          select AA.���,AA.ID,AA.����ID,AA.�ϼ�ID,AA.����,AA.����,AA.��׼��ֵ,AA.ȱ�ݵȼ�,AA.���ֵ�λ,count(BB.ID) as �������
          from �������ֱ�׼ AA,�������ֱ�׼ BB
          where AA.ID=BB.�ϼ�ID(+)
          group by AA.���,AA.ID,AA.����ID,AA.�ϼ�ID,AA.����,AA.����,AA.��׼��ֵ,AA.ȱ�ݵȼ�,AA.���ֵ�λ
      ) A,
      (
          select ��� as �ϼ����,ID,����,��׼��ֵ,���� from �������ֱ�׼
      ) B
  where A.�ϼ�ID=B.ID(+)
) T
order by decode(T.�ϼ����,null,���,T.�ϼ����),decode(T.���,null,T.ID,T.���);


create or replace view ��������������ͼ as
Select Tb.סԺ��, Tb.����, Tb.�Ա�, Ta.*
From (Select T1.����id, T1.��ҳid, T1.��Ժ����, T1.��Ժ����, T2.���� As ��Ժ����, T3.���� As ��Ժ����, T1.����ҽʦ,
              T1.���λ�ʿ, T1.סԺҽʦ, T1.��Ŀ����, T1.���id, T1.����id, T1.�ܷ�, T1.�ȼ�, T1.������,
              To_Char(T1.����ʱ��, 'YYYY-MM-DD') As ����ʱ��, T1.�����, To_Char(T1.���ʱ��, 'YYYY-MM-DD') As ���ʱ��,
              T1.�����޸�, T1.��ע
       From (Select A.����id, A.��ҳid, A.��Ժ����id, A.��Ժ����id, A.��Ժ����, A.��Ժ����, A.����ҽʦ, A.���λ�ʿ,
                     A.סԺҽʦ, A.��Ŀ����, B.ID As ���id, B.����id, B.�ܷ�, B.�ȼ�, B.������, B.����ʱ��, B.�����,
                     B.���ʱ��, B.�����޸�, B.��ע
              From ������ҳ A, �������ֽ�� B
              Where A.����id = B.����id(+) And A.��ҳid = B.��ҳid(+)) T1, ���ű� T2, ���ű� T3
       Where T1.��Ժ����id = T2.ID And T1.��Ժ����id = T3.ID) Ta, ������Ϣ Tb
Where Ta.����id = Tb.����id;



--zlStreamTabs����
-----------------------------------------------------------------------------------------------------------------------



--zlBakTables����
-----------------------------------------------------------------------------------------------------------------------


--zlComponent����
-----------------------------------------------------------------------------------------------------------------------
Insert Into zlComponent(����,����,���汾,�ΰ汾,���汾,ϵͳ) Values('zl9CISAduit','���Ӳ������鵵',10,20,0,100);

--zlPrograms����
-----------------------------------------------------------------------------------------------------------------------
Insert into zlPrograms(���,����,˵��,ϵͳ,����) Values(1550,'���ֱ�׼ά��','��ɲ������ֱ�׼��ɾ�ĺ�ѡ�á�',100,'zl9CISAduit');
Insert Into zlPrograms(���,����,˵��,ϵͳ,����) Values(1560,'���Ӳ������','��Ժ���˵ĵ��Ӳ��������͹鵵�Լ���Ժ���˵ĵ��Ӳ����ĳ�鼰��档',100,'zl9CISAduit');
Insert Into zlPrograms(���,����,˵��,ϵͳ,����) Values(1561,'���Ӳ�������','�鵵���˵ĵ��Ӳ������ĵ�����Ͳ��ġ�',100,'zl9CISAduit');
Insert into zlPrograms(���,����,˵��,ϵͳ,����) Values(1562,'���Ӳ�������','����¼������ֱ�׼�����в����������ֺ���ˡ�',100,'zl9CISAduit');


--zlProgFuncs����
-----------------------------------------------------------------------------------------------------------------------
--���ֱ�׼ά��
Insert Into zlProgFuncs(ϵͳ,���,����,˵��)  values (100,1550,'����','');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��)  values (100,1550,'��ɾ��','�д�Ȩ�޵��û����Խ��в������ֱ�׼����ɾ�ļ�ѡ�ò�����');

--���Ӳ������
Insert Into zlProgFuncs(ϵͳ,���,����,˵��)  values (100,1560,'����','');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��)  values (100,1560,'��������','���ñ�ģ����ص�ȫ�ֲ���');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��)  values (100,1560,'������','��ʼ��鲡�˵ĵ��Ӳ���');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��)  values (100,1560,'�ܾ����','�ܾ���鲡�˵ĵ��Ӳ���');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��)  values (100,1560,'��鲡��','�Բ��˵ĵ��Ӳ���������鴦��');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��)  values (100,1560,'�鵵����','�Բ��˵ĵ��Ӳ������й鵵����');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��)  values (100,1560,'��没��','����Ժ���˵ĵ��Ӳ������з�洦��');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��)  values (100,1560,'��ⲡ��','����Ժ���˵ĵ��Ӳ������н�⴦��');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��)  values (100,1560,'���˽���','�����Ѿ���ʼ���Ĳ��˵��Ӳ���');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��)  values (100,1560,'���˾ܾ�','���˱��ܾ��Ĳ��˵��Ӳ���');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��)  values (100,1560,'���˹鵵','�����Ѿ��鵵�Ĳ��˵��Ӳ���');

--���Ӳ�������
Insert Into zlProgFuncs(ϵͳ,���,����,˵��)  values (100,1561,'����','');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��)  values (100,1561,'��������','���ñ�ģ����ص�ȫ�ֲ���');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��)  values (100,1561,'�Ǽ�����','�Ǽǡ��޸ĺ�ɾ�����Ӳ����Ľ������뵥');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��)  values (100,1561,'��������','��׼��ܾ��µǼǵĽ������뵥');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��)  values (100,1561,'���Ĳ���','�����Ѿ���׼�Ĳ��˵ĵ��Ӳ�������');
--�����������
Insert Into zlProgFuncs(ϵͳ,���,����,˵��)  values (100,1562,'����','');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��)  values (100,1562,'����','�д�Ȩ�޵��û����Խ��в������ֲ�����');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��)  values (100,1562,'���','�д�Ȩ�޵��û����Խ��в������ֽ������˲�����');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��)  values (100,1562,'ȡ�����','�д�Ȩ�޵��û����Զ��Ѿ���˲�������ȡ����˵Ĳ�����');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��)  values (100,1562,'���п���','�д�Ȩ�޵��û����Զ����п��ҽ������֣�����ֻ�ܶԱ����Ҳ����������֡�');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��)  values (100,1562,'�޸���������','�д�Ȩ�޵��û������޸��������ֽ��������ֻ���޸ı������ֽ����');


--zlProgPrivs����
-----------------------------------------------------------------------------------------------------------------------
--���ֱ�׼ά��
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1550,'����',user,'�������ַ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1550,'����',user,'�������ֽ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1550,'����',user,'�������ַ���_ID','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1550,'����',user,'�������ֱ�׼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1550,'����',user,'�������ֱ�׼_ID','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1550,'����',user,'�������ֱ�׼��ͼ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1550,'��ɾ��',user,'ZL_�������ַ���_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1550,'��ɾ��',user,'ZL_�������ַ���_Update','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1550,'��ɾ��',user,'ZL_�������ַ���_Delete','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1550,'��ɾ��',user,'ZL_�������ֱ�׼_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1550,'��ɾ��',user,'ZL_�������ֱ�׼_Update','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1550,'��ɾ��',user,'ZL_�������ֱ�׼_Delete','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1550,'��ɾ��',user,'ZL_�������ַ���_ѡ��','EXECUTE');

--���Ӳ������
--����
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'�����ύ��¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'����������¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'����������ʷ','SELECT');

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'������Ϣ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'������ҳ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'������ҳ�ӱ�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'��Ϸ������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'���˱䶯��¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'��λ״����¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'���˹�����¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'������ϼ�¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'���������¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'������������¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'�������Ҷ�Ӧ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'����ҽ��״̬','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'����ҽ������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'���˷��ü�¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'ҩƷ����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'ҩƷ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'������ĿĿ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'��������Ӧ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'���Ӳ�������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'Zl_Lob_Read','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'���Ӳ�������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'��˽������Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'����������Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'�����ļ��ṹ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'���˻�������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'�����¼��Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'���¼�¼��Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'���Ӳ�����¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'����ҽ������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'����ҽ����¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'�����ļ��б�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'����ҳ���ʽ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'����Ӧ�ÿ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'����',user,'���˻����¼','SELECT');
--������
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'������',user,'zl_�����ύ��¼_Receive','EXECUTE');
--�ܾ����
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'�ܾ����',user,'zl_�����ύ��¼_Refuse','EXECUTE');
--�鵵����
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'�鵵����',user,'zl_�����ύ��¼_Archive','EXECUTE');
--���˽���
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'���˽���',user,'zl_�����ύ��¼_UnReceive','EXECUTE');
--���˾ܾ�
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'���˾ܾ�',user,'zl_�����ύ��¼_UnRefuse','EXECUTE');
--���˹鵵
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'���˹鵵',user,'zl_�����ύ��¼_UnArchive','EXECUTE');
--��没��
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'��没��',user,'zl_��������¼_Lock','EXECUTE');
--��ⲡ��
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'��ⲡ��',user,'zl_��������¼_UnLock','EXECUTE');
--��鲡��
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'��鲡��',user,'�������ַ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'��鲡��',user,'�������ֱ�׼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'��鲡��',user,'����������¼_ID','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'��鲡��',user,'����ʱ��Ҫ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'��鲡��',user,'���˹Һż�¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'��鲡��',user,'������д�¼�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'��鲡��',user,'�������ݼ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'��鲡��',user,'����ʱ�޼��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'��鲡��',user,'zl_����������¼_Update','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'��鲡��',user,'zl_����������¼_Delete','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'��鲡��',user,'zl_����������¼_Finish','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'��鲡��',user,'zl_����������¼_RollBackFinish','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'��鲡��',user,'Zl_�������ݼ��_Neaten','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1560,'��鲡��',user,'Zl_����ʱ�޼��_Neaten','EXECUTE');

--���Ӳ�������
--����
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'������Ϣ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'������ҳ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'������ҳ�ӱ�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'��Ϸ������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'���˱䶯��¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'��λ״����¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'���˹�����¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'������ϼ�¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'���������¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'������������¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'�������Ҷ�Ӧ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'����ҽ��״̬','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'����ҽ������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'���˷��ü�¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'ҩƷ����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'ҩƷ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'������ĿĿ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'��������Ӧ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'���Ӳ�������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'Zl_Lob_Read','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'���Ӳ�������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'��˽������Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'����������Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'�����ļ��ṹ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'���˻�������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'�����¼��Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'���¼�¼��Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'���Ӳ�����¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'����ҽ������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'����ҽ����¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'�����ļ��б�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'����ҳ���ʽ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'����Ӧ�ÿ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'���˻����¼','SELECT');

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'�������ļ�¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'����������Ա','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'����',user,'������������','SELECT');
--�Ǽ�����
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'�Ǽ�����',user,'�������ļ�¼_ID','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'�Ǽ�����',user,'�Ա�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'�Ǽ�����',user,'����״��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'�Ǽ�����',user,'�����������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'�Ǽ�����',user,'��������Ŀ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'�Ǽ�����',user,'zl_����������Ա_Update','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'�Ǽ�����',user,'zl_������������_Update','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'�Ǽ�����',user,'zl_�������ļ�¼_Update','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'�Ǽ�����',user,'zl_�������ļ�¼_Delete','EXECUTE');
--��������
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'��������',user,'zl_�������ļ�¼_Authorize','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'��������',user,'zl_�������ļ�¼_Refuse','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1561,'��������',user,'zl_�������ļ�¼_Rollback','EXECUTE');

--�����������
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'�������ֽ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'����������ϸ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'�������ֱ�׼��ͼ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'��������������ͼ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'�������ַ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'�������ֱ�׼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'������ҳ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'������ҳ�ӱ�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'���˹���ҩ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'������Ϣ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'���ű�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'������Ա','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'��������˵��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'��������Ŀ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'��Ա��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'�ϻ���Ա��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'���������¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'��Ϸ������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'������ϼ�¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'סԺ������¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'�ٴ�����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'����״��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'ְҵ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'Ѫ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'ҽ�Ƹ��ʽ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'�����������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'������Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'�������ֽ��_ID','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'����������ϸ_ID','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'ZL_�������ֽ��_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'ZL_�������ֽ��_Update','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'ZL_�������ֽ��_Delete','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'����',user,'ZL_����������ϸ_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'���',user,'ZL_�������ֽ��_���','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'���',user,'������ҳ�ӱ�','INSERT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1562,'ȡ�����',user,'ZL_�������ֽ��_ȡ�����','EXECUTE');

--zlMenus����
-----------------------------------------------------------------------------------------------------------------------
Insert Into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Values('ȱʡ',zlMenus_id.nextval,null,'���Ӳ�������','��������','A',99,'���Ӳ��������鵵����',100,NULL);
Insert Into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Values('ȱʡ',zlMenus_id.nextval,   zlMenus_id.nextval-1,'���ֱ�׼ά��','���ֱ�׼','A',231,'��ɲ������ֱ�׼��ɾ�ĺ�ѡ�á�',100,1550);
Insert Into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Values('ȱʡ',zlMenus_id.nextval,   zlMenus_id.nextval-2,'���Ӳ������','�������','B',232,'��Ժ���˵ĵ��Ӳ��������͹鵵�Լ���Ժ���˵ĵ��Ӳ����ĳ�鼰��档',100,1560);
Insert Into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Values('ȱʡ',zlMenus_id.nextval,   zlMenus_id.nextval-3,'���Ӳ�������','��������','C',141,'�鵵���˵ĵ��Ӳ������ĵ�����Ͳ��ġ�',100,1561);
Insert Into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Values('ȱʡ',zlMenus_id.nextval,   zlMenus_id.nextval-4,'���Ӳ�������','��������','D',136,'����¼������ֱ�׼�����в����������ֺ���ˡ�',100,1562);
Insert Into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Values('ȱʡ',zlMenus_id.nextval,   zlMenus_id.nextval-5,'�����������','�������','E',99,'���Ӳ������鵵����ر������',100,NULL);

--zlBaseCode����
-----------------------------------------------------------------------------------------------------------------------

--zlParameters����
--1560:���Ӳ������
Insert Into zlParameters(ID,ϵͳ,ģ��,˽��,������,������,����ֵ,ȱʡֵ,����˵��)
Select Rownum+B.ID,A.* From (
	Select ϵͳ,ģ��,˽��,������,������,����ֵ,ȱʡֵ,����˵�� From zlParameters Where ID=0 Union ALL
	Select 100,1560,1,1,'��λ����','����','����','���Ҷ�λ���˵ķ�ʽ' From Dual Union ALL
	Select 100,1560,1,2,'�ϴ�״̬','0','0','����Ա�ϴ�ѡ��Ĳ����б����' From Dual Union ALL
	Select 100,1560,1,3,'���ȱʡ��Χ','��  ��','��  ��','��ʾδ�鵵���˵�ȱʡʱ�䷶Χ' From Dual Union ALL
	Select 100,1560,1,4,'�鵵ȱʡ��Χ','��  ��','��  ��','��ʾ�ѹ鵵���˵�ȱʡʱ�䷶Χ' From Dual Union ALL
	Select 100,1560,0,5,'������������','7','7','����������Ҫ���ٴ����Ҵ������������' From Dual Union ALL
	Select 100,1560,0,6,'δ����ˢ��Ƶ��','5','5','�Զ�ˢ��δ���鷴�������ʱ��������λ������' From Dual Union ALL
	Select 100,1560,1,7,'������ⷶΧ','��  ��','��  ��','��ʾ����ɵ������ʱ�䷶Χ' From Dual
  ) A,(Select Nvl(Max(ID),0) AS ID From zlParameters) B;

--1561:���Ӳ�������
Insert Into zlParameters(ID,ϵͳ,ģ��,˽��,������,������,����ֵ,ȱʡֵ,����˵��)
Select Rownum+B.ID,A.* From (
	Select ϵͳ,ģ��,˽��,������,������,����ֵ,ȱʡֵ,����˵�� From zlParameters Where ID=0 Union ALL
	Select 100,1561,1,1,'��λ����','No','No','���Ҷ�λ��������ķ�ʽ' From Dual Union ALL
	Select 100,1561,1,2,'�ϴ�״̬','0','0','����Ա�ϴ�ѡ����б����' From Dual Union ALL
	Select 100,1561,1,3,'�Ǽ�ȱʡ��Χ','��  ��','��  ��','��ʾ�Ǽ������ȱʡʱ�䷶Χ' From Dual Union ALL
	Select 100,1561,0,4,'������������','7','7','�������ĵ���������' From Dual
  ) A,(Select Nvl(Max(ID),0) AS ID From zlParameters) B;

--������zlParameters������
Select zlParameters_ID.Nextval From zlParameters;

--������Ʊ�����
Insert Into ������Ʊ�(��Ŀ���,��Ŀ����,������,�Զ���ȱ,��Ź���)
Select 86,'�����������뵥','',1,0 From Dual;

------------------------------------------------------------------------------------------------------------------------------------------
--		�����嵥
------------------------------------------------------------------------------------------------------------------------------------------
--�������ַ���
CREATE OR REPLACE PROCEDURE ZL_�������ַ���_Insert
(	ID_in IN �������ַ���.ID%TYPE,
	����_in IN �������ַ���.����%TYPE,
	�ܷ�_in IN �������ַ���.�ܷ�%TYPE,
	��ֵ_in IN �������ַ���.��ֵ%TYPE,
	��ֵ_in IN �������ַ���.��ֵ%TYPE,
	����_in IN �������ַ���.����%TYPE,
	����_in IN �������ַ���.����%TYPE,
	ѡ��_in IN �������ַ���.ѡ��%TYPE,
	����ʱ��_in IN �������ַ���.����ʱ��%TYPE,
	ͣ��ʱ��_in IN �������ַ���.ͣ��ʱ��%TYPE
)
IS
BEGIN
  if ѡ��_in=1 then
     update �������ַ��� 
     set ѡ��=0
     where ����=����_in;  
  end if;
  
	INSERT INTO �������ַ���
		(ID,����,�ܷ�,��ֵ,��ֵ,����,����,ѡ��,����ʱ��,ͣ��ʱ��)
	VALUES
		(ID_in,����_in,�ܷ�_in,��ֵ_in,��ֵ_in,����_in,����_in,ѡ��_in,����ʱ��_in,ͣ��ʱ��_in);
    
EXCEPTION
	WHEN OTHERS THEN
		ZL_ErrorCenter (SQLCODE, SQLERRM);
END  ZL_�������ַ���_Insert;
/

CREATE OR REPLACE PROCEDURE ZL_�������ַ���_Update
(	ID_in IN �������ַ���.ID%TYPE,
	����_in IN �������ַ���.����%TYPE,
	�ܷ�_in IN �������ַ���.�ܷ�%TYPE,
	��ֵ_in IN �������ַ���.��ֵ%TYPE,
	��ֵ_in IN �������ַ���.��ֵ%TYPE,
	����_in IN �������ַ���.����%TYPE,
	����_in IN �������ַ���.����%TYPE,
	ѡ��_in IN �������ַ���.ѡ��%TYPE,
	����ʱ��_in IN �������ַ���.����ʱ��%TYPE,
	ͣ��ʱ��_in IN �������ַ���.ͣ��ʱ��%TYPE
)
IS
BEGIN
  if ѡ��_in=1 then
     update �������ַ��� 
     set ѡ��=0
     where ����=����_in;  
  end if;

	Update �������ַ���
	set	����=����_in,�ܷ�=�ܷ�_in,��ֵ=��ֵ_in,��ֵ=��ֵ_in,����=����_in,����=����_in,ѡ��=ѡ��_in,����ʱ��=����ʱ��_in,ͣ��ʱ��=ͣ��ʱ��_in
	where ID=ID_in;
	
	IF SQL%NOTFOUND THEN
		---���û�и��µ�����ô����һ��
		INSERT INTO �������ַ���
			(ID,����,�ܷ�,��ֵ,��ֵ,����,����,ѡ��,����ʱ��,ͣ��ʱ��)
		VALUES
			(ID_in,����_in,�ܷ�_in,��ֵ_in,��ֵ_in,����_in,����_in,ѡ��_in,����ʱ��_in,ͣ��ʱ��_in);
	END IF;

EXCEPTION
	WHEN OTHERS THEN
		ZL_ErrorCenter (SQLCODE, SQLERRM);
END  ZL_�������ַ���_Update;
/

CREATE OR REPLACE PROCEDURE ZL_�������ַ���_Delete
(
	ID_in IN �������ַ���.ID%TYPE
)
IS
BEGIN
	DELETE FROM �������ַ���
		WHERE ID = ID_in;
EXCEPTION
	WHEN OTHERS THEN
		ZL_ErrorCenter (SQLCODE, SQLERRM);
END  ZL_�������ַ���_Delete;
/

CREATE OR REPLACE PROCEDURE ZL_�������ַ���_ѡ��
(	ID_in IN �������ַ���.ID%TYPE,
  ѡ��_in IN �������ַ���.ѡ��%TYPE
)
IS
BEGIN
  if ѡ��_in=1 then
     update �������ַ���
     set ѡ��=0
     where ����=(select ���� from �������ַ��� where ID=ID_in);
  end if;

	update �������ַ���
	set ѡ��=ѡ��_in
  where ID=ID_in; 

EXCEPTION
	WHEN OTHERS THEN
		ZL_ErrorCenter (SQLCODE, SQLERRM);
END  ZL_�������ַ���_ѡ��;
/

--�������ֱ�׼

CREATE OR REPLACE PROCEDURE ZL_�������ֱ�׼_Insert
(	ID_in IN �������ֱ�׼.ID%TYPE,
	�ϼ�ID_in IN �������ֱ�׼.�ϼ�ID%TYPE,
	����ID_in IN �������ֱ�׼.����ID%TYPE,
	����_in IN �������ֱ�׼.����%TYPE,
	����_in IN �������ֱ�׼.����%TYPE,
	��׼��ֵ_in IN �������ֱ�׼.��׼��ֵ%TYPE,
	ȱ�ݵȼ�_in IN �������ֱ�׼.ȱ�ݵȼ�%TYPE,
	���ֵ�λ_in IN �������ֱ�׼.���ֵ�λ%TYPE,
  ��׼ID_IN IN �������ֱ�׼.���%TYPE
)
IS
  ��׼��� NUMBER;
BEGIN
  if ��׼ID_IN=0 or ��׼ID_IN is null then
     if �ϼ�ID_IN=0 or �ϼ�ID_IN is null then
        select decode(max(���),null,0,max(���)+1) into ��׼��� from �������ֱ�׼ where ����ID=����ID_IN and �ϼ�ID is null;
     else
        select decode(max(���),null,0,max(���)+1) into ��׼��� from �������ֱ�׼ where ����ID=����ID_IN and �ϼ�ID=�ϼ�ID_IN;
     end if;
     --������ֱ�׼
     INSERT INTO �������ֱ�׼
       (ID,�ϼ�ID,����ID,����,����,��׼��ֵ,ȱ�ݵȼ�,���ֵ�λ,���)
     VALUES
       (ID_IN,�ϼ�ID_IN,����ID_IN,����_IN,����_IN,��׼��ֵ_IN,ȱ�ݵȼ�_IN,���ֵ�λ_IN,��׼���);       
  else
     --�������ֱ�׼
    if �ϼ�ID_IN=0 or �ϼ�ID_IN is null then   --Ϊ������Ŀ
       select ��� into ��׼��� from �������ֱ�׼ where ID=��׼ID_IN;
       update �������ֱ�׼ set ���=���+1 where �ϼ�ID is null and ���>=��׼��� and ����ID=����ID_IN;
       INSERT INTO �������ֱ�׼
      	 (ID,�ϼ�ID,����ID,����,����,��׼��ֵ,ȱ�ݵȼ�,���ֵ�λ,���)
       VALUES
      	 (ID_IN,�ϼ�ID_IN,����ID_IN,����_IN,����_IN,��׼��ֵ_IN,ȱ�ݵȼ�_IN,���ֵ�λ_IN,��׼���);     
    else                        --Ϊ���ֱ�׼
       select ��� into ��׼��� from �������ֱ�׼ where ID=��׼ID_IN;
       update �������ֱ�׼ set ���=���+1 where �ϼ�ID=�ϼ�ID_in  and ���>=��׼��� and ����ID=����ID_IN;
       INSERT INTO �������ֱ�׼
      	 (ID,�ϼ�ID,����ID,����,����,��׼��ֵ,ȱ�ݵȼ�,���ֵ�λ,���)
       VALUES
      	 (ID_IN,�ϼ�ID_IN,����ID_IN,����_IN,����_IN,��׼��ֵ_IN,ȱ�ݵȼ�_IN,���ֵ�λ_IN,��׼���);     
    end if;
  end if;
EXCEPTION
	WHEN OTHERS THEN
		ZL_ErrorCenter (SQLCODE, SQLERRM);
END  ZL_�������ֱ�׼_Insert;
/

CREATE OR REPLACE PROCEDURE ZL_�������ֱ�׼_Update
(	ID_in IN �������ֱ�׼.ID%TYPE,
	�ϼ�ID_in IN �������ֱ�׼.�ϼ�ID%TYPE,
	����ID_in IN �������ֱ�׼.����ID%TYPE,
	����_in IN �������ֱ�׼.����%TYPE,
	����_in IN �������ֱ�׼.����%TYPE,
	��׼��ֵ_in IN �������ֱ�׼.��׼��ֵ%TYPE,
	ȱ�ݵȼ�_in IN �������ֱ�׼.ȱ�ݵȼ�%TYPE,
	���ֵ�λ_in IN �������ֱ�׼.���ֵ�λ%TYPE
)
IS
BEGIN
	Update �������ֱ�׼
	set �ϼ�ID=�ϼ�ID_IN,����ID=����ID_IN,����=����_IN,����=����_IN,��׼��ֵ=��׼��ֵ_IN,ȱ�ݵȼ�=ȱ�ݵȼ�_IN,���ֵ�λ=���ֵ�λ_IN
	where ID=ID_IN;
	
	IF SQL%NOTFOUND THEN
		---���û�и��µ�����ô����һ��
		INSERT INTO �������ֱ�׼
			(ID,�ϼ�ID,����ID,����,����,��׼��ֵ,ȱ�ݵȼ�,���ֵ�λ)
		VALUES
			(ID_IN,�ϼ�ID_IN,����ID_IN,����_IN,����_IN,��׼��ֵ_IN,ȱ�ݵȼ�_IN,���ֵ�λ_IN);
	END IF;

EXCEPTION
	WHEN OTHERS THEN
		ZL_ErrorCenter (SQLCODE, SQLERRM);
END  ZL_�������ֱ�׼_Update;
/

CREATE OR REPLACE PROCEDURE ZL_�������ֱ�׼_Delete
(
	ID_in IN �������ֱ�׼.ID%TYPE,
  ɾ��������Ŀ_in IN NUMBER
)
IS
  lng�ϼ�ID NUMBER;
  lng���������� NUMBER;
BEGIN
  if ɾ��������Ŀ_in=1 then
    select decode(�ϼ�ID,null,0,�ϼ�ID) into lng�ϼ�ID
      from �������ֱ�׼ where ID=ID_IN;
  end if;
  
	DELETE FROM �������ֱ�׼
		WHERE ID = ID_in;
    
  if ɾ��������Ŀ_in=1 then
 
    select decode(����,'��',1,0) into lng����������
      from �������ֱ�׼��ͼ where ID=lng�ϼ�ID;
    if lng����������=1 then
       DELETE FROM �������ֱ�׼	WHERE ID = lng�ϼ�ID;
    end if;
  end if;

EXCEPTION
	WHEN OTHERS THEN
		ZL_ErrorCenter (SQLCODE, SQLERRM);
END  ZL_�������ֱ�׼_Delete;
/


--�������ֽ��
Create Or Replace Procedure Zl_�������ֽ��_Insert
(
  Id_In       In �������ֽ��.ID%Type,
  ����id_In   In �������ֽ��.����id%Type,
  ��ҳid_In   In �������ֽ��.��ҳid%Type,
  ����id_In   In �������ֽ��.����id%Type,
  �ܷ�_In     In �������ֽ��.�ܷ�%Type,
  �ȼ�_In     In �������ֽ��.�ȼ�%Type,
  ��ע_In     In �������ֽ��.��ע%Type,
  ������_In   In �������ֽ��.������%Type,
  ����ʱ��_In In �������ֽ��.����ʱ��%Type,
  �����_In   In �������ֽ��.�����%Type,
  ���ʱ��_In In �������ֽ��.���ʱ��%Type,
  �����޸�_In In �������ֽ��.�����޸�%Type
) Is
Begin
  Insert Into �������ֽ��
    (ID, ����id, ��ҳid, ����id, �ܷ�, �ȼ�, ��ע, ������, ����ʱ��, �����, ���ʱ��, �����޸�)
  Values
    (Id_In, ����id_In, ��ҳid_In, ����id_In, �ܷ�_In, �ȼ�_In, ��ע_In, ������_In, ����ʱ��_In, �����_In, ���ʱ��_In,
     �����޸�_In);
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�������ֽ��_Insert;
/

Create Or Replace Procedure Zl_�������ֽ��_Update
(
  Id_In       In �������ֽ��.ID%Type,
  ����id_In   In �������ֽ��.����id%Type,
  ��ҳid_In   In �������ֽ��.��ҳid%Type,
  ����id_In   In �������ֽ��.����id%Type,
  �ܷ�_In     In �������ֽ��.�ܷ�%Type,
  �ȼ�_In     In �������ֽ��.�ȼ�%Type,
  ��ע_In     In �������ֽ��.��ע%Type,
  ������_In   In �������ֽ��.������%Type,
  ����ʱ��_In In �������ֽ��.����ʱ��%Type,
  �����_In   In �������ֽ��.�����%Type,
  ���ʱ��_In In �������ֽ��.���ʱ��%Type,
  �����޸�_In In �������ֽ��.�����޸�%Type
) Is
Begin
  Update �������ֽ��
  Set ����id = ����id_In, ��ҳid = ��ҳid_In, ����id = ����id_In, �ܷ� = �ܷ�_In, �ȼ� = �ȼ�_In, ��ע = ��ע_In,
      ������ = ������_In, ����ʱ�� = ����ʱ��_In, ����� = �����_In, ���ʱ�� = ���ʱ��_In, �����޸� = �����޸�_In
  Where ID = Id_In;

  If Sql%NotFound Then
    ---���û�и��µ�����ô����һ��
    Insert Into �������ֽ��
      (ID, ����id, ��ҳid, ����id, �ܷ�, �ȼ�, ��ע, ������, ����ʱ��, �����, ���ʱ��, �����޸�)
    Values
      (Id_In, ����id_In, ��ҳid_In, ����id_In, �ܷ�_In, �ȼ�_In, ��ע_In, ������_In, ����ʱ��_In, �����_In,
       ���ʱ��_In, �����޸�_In);
  End If;

Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�������ֽ��_Update;
/


CREATE OR REPLACE PROCEDURE ZL_�������ֽ��_Delete
(
	ID_in IN �������ֽ��.ID%TYPE
)
IS
BEGIN
	DELETE FROM �������ֽ��
		WHERE ID = ID_in;
EXCEPTION
	WHEN OTHERS THEN
		ZL_ErrorCenter (SQLCODE, SQLERRM);
END  ZL_�������ֽ��_Delete;
/

  --����������ϸ

Create Or Replace Procedure Zl_����������ϸ_Insert
(
  Id_In         In ����������ϸ.ID%Type,
  ����id_In     In ����������ϸ.����id%Type,
  ���ֱ�׼id_In In ����������ϸ.���ֱ�׼id%Type,
  �������_In   In ����������ϸ.�������%Type,
  ȱ�ݵȼ�_In   In ����������ϸ.ȱ�ݵȼ�%Type,
  �ɷ��޸�_In   In ����������ϸ.�ɷ��޸�%Type,
  ��ע_In       In �������ֽ��.��ע%Type
  
) Is
Begin
  Insert Into ����������ϸ
    (ID, ����id, ���ֱ�׼id, �������, ȱ�ݵȼ�, �ɷ��޸�, ��ע)
  Values
    (Id_In, ����id_In, ���ֱ�׼id_In, �������_In, ȱ�ݵȼ�_In, �ɷ��޸�_In, ��ע_In);
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����������ϸ_Insert;
/

Create Or Replace Procedure Zl_����������ϸ_Update
(
  Id_In         In ����������ϸ.ID%Type,
  ����id_In     In ����������ϸ.����id%Type,
  ���ֱ�׼id_In In ����������ϸ.���ֱ�׼id%Type,
  �������_In   In ����������ϸ.�������%Type,
  ȱ�ݵȼ�_In   In ����������ϸ.ȱ�ݵȼ�%Type,
  �ɷ��޸�_In   In ����������ϸ.�ɷ��޸�%Type,
  ��ע_In       In �������ֽ��.��ע%Type
) Is
Begin
  Update ����������ϸ
  Set ����id = ����id_In, ���ֱ�׼id = ���ֱ�׼id_In, ������� = �������_In, ȱ�ݵȼ� = ȱ�ݵȼ�_In,
      �ɷ��޸� = �ɷ��޸�_In, ��ע = ��ע_In
  Where ID = Id_In;

  If Sql%NotFound Then
    ---���û�и��µ�����ô����һ�� 
    Insert Into ����������ϸ
      (ID, ����id, ���ֱ�׼id, �������, ȱ�ݵȼ�, �ɷ��޸�, ��ע)
    Values
      (Id_In, ����id_In, ���ֱ�׼id_In, �������_In, ȱ�ݵȼ�_In, �ɷ��޸�_In, ��ע_In);
  End If;

Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����������ϸ_Update;
/


CREATE OR REPLACE PROCEDURE ZL_����������ϸ_Delete
(
	ID_in IN ����������ϸ.ID%TYPE
)
IS
BEGIN
	DELETE FROM ����������ϸ
		WHERE ID = ID_in;
EXCEPTION
	WHEN OTHERS THEN
		ZL_ErrorCenter (SQLCODE, SQLERRM);
END  ZL_����������ϸ_Delete;
/

  --�������ֽ��-�����ȡ�����

Create Or Replace Procedure Zl_�������ֽ��_���
(
  Id_In     In �������ֽ��.ID%Type,
  �����_In In �������ֽ��.�����%Type
) Is
  n_��Ŀ         Number;
  n_�Զ�д��ӱ� Number;
  n_����id       Number;
  n_��ҳid       Number;
  v_�ȼ�         Varchar2(2);
  v_�ɵȼ�ֵ     Varchar2(2);
Begin
  Update �������ֽ�� Set ����� = �����_In, ���ʱ�� = Sysdate Where ID = Id_In;

  Select ����id Into n_����id From �������ֽ�� Where ID = Id_In;
  Select ��ҳid Into n_��ҳid From �������ֽ�� Where ID = Id_In;
  Select �ȼ� Into v_�ȼ� From �������ֽ�� Where ID = Id_In;

  Select Count(*) Into n_��Ŀ From ������ҳ�ӱ� Where ����id = n_����id And ��ҳid = n_��ҳid And ��Ϣ�� = '��������';

  n_�Զ�д��ӱ� := To_Number(Zl_Getsysparameter(90, 0, 0), '9999999');

  If n_��Ŀ = 0 And n_�Զ�д��ӱ� = 1 And v_�ȼ� <> '��' Then
    Insert Into ������ҳ�ӱ� (����id, ��ҳid, ��Ϣ��, ��Ϣֵ) Values (n_����id, n_��ҳid, '��������', v_�ȼ�);
  Else
    If n_��Ŀ = 1 And n_�Զ�д��ӱ� = 1 And v_�ȼ� <> '��' Then
      Select ��Ϣֵ
      Into v_�ɵȼ�ֵ
      From ������ҳ�ӱ�
      Where ����id = n_����id And ��ҳid = n_��ҳid And ��Ϣ�� = '��������';
      If v_�ɵȼ�ֵ Is Null Then
        Update ������ҳ�ӱ� Set ��Ϣֵ = v_�ȼ� Where ����id = n_����id And ��ҳid = n_��ҳid And ��Ϣ�� = '��������';
      End If;
    End If;
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�������ֽ��_���;
/

CREATE OR REPLACE PROCEDURE ZL_�������ֽ��_ȡ�����
(	ID_in IN �������ֽ��.ID%TYPE
)
IS
BEGIN
	Update �������ֽ��
	set �����=NULL,���ʱ��=NULL
	where ID=ID_IN;


EXCEPTION
	WHEN OTHERS THEN
		ZL_ErrorCenter (SQLCODE, SQLERRM);
END  ZL_�������ֽ��_ȡ�����;
/
----------------------------------------------------------------------------
---  UPDATE   for   �������ļ�¼
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_�������ļ�¼_Update(
	ID_IN		IN	�������ļ�¼.ID%TYPE,
	No_IN		IN	�������ļ�¼.No%TYPE,
	������_IN		IN	�������ļ�¼.������%TYPE,
	��������_IN	IN	�������ļ�¼.��������%TYPE,
	����ʱ��_IN	IN	�������ļ�¼.����ʱ��%TYPE,
	��������_IN	IN	�������ļ�¼.��������%TYPE,
	�Ǽ�ʱ��_IN	IN	�������ļ�¼.�Ǽ�ʱ��%TYPE:=Sysdate
)
IS
BEGIN
	Update �������ļ�¼ Set No=No_IN,
				������=������_IN,
				��������=��������_IN,
				����ʱ��=����ʱ��_IN,
				��������=��������_IN
	Where ID=ID_IN And ��¼״̬=1;
	
	If SQL%RowCount=0 Then
		Insert Into �������ļ�¼(ID,No,��¼״̬,������,��������,����ʱ��,��������,�Ǽ�ʱ��) 
		VALUES (ID_IN,No_IN,1,������_IN,��������_IN,����ʱ��_IN,��������_IN,�Ǽ�ʱ��_IN);
	End If;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_�������ļ�¼_Update;
/

----------------------------------------------------------------------------
---  Update   for   ����������Ա
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_����������Ա_Update(
	����id_IN		IN	����������Ա.����id%TYPE,
	��Աid_In		IN	Varchar2:=Null
)
IS
	strTmp			Varchar2(4000);
	intPos			Number(18);
	n_��Աid			����������Ա.��Աid%TYPE;
BEGIN
	If ��Աid_IN Is Null Then
		Delete From ����������Ա Where ����id=����id_IN;
	Else
		strTmp := ��Աid_In||';';
		WHILE strTmp IS NOT NULL LOOP
			intPos := INSTR (strTmp, ';');
			IF intPos >0 Then
				n_��Աid := To_Number(SUBSTR (strTmp, 1, intPos - 1));
				strTmp := SUBSTR (strTmp, intPos + 1);
				If n_��Աid>0 Then
					Insert Into ����������Ա(����id,��Աid) values (����id_IN,n_��Աid);
				End If;
			End If;       
		END LOOP;
	End If;

EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_����������Ա_Update;
/

----------------------------------------------------------------------------
---  Update   for   ������������
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_������������_Update(
	����id_IN		IN	������������.����id%TYPE,
	����id_In		IN	Varchar2:=Null
)
IS
	strTmp			Varchar2(4000);
	str����			Varchar2(4000);
	intPos			Number(18);
	n_����id			������������.����id%TYPE;
	n_��ҳid			������������.��ҳid%TYPE;
BEGIN
	If ����id_IN Is Null Then
		Delete From ������������ Where ����id=����id_IN;
	Else
		strTmp := ����id_In||';';
		WHILE strTmp IS NOT NULL LOOP
			intPos := INSTR (strTmp, ';');

			IF intPos >0 Then
				
				str���� := SUBSTR (strTmp, 1, intPos - 1)||':';
				strTmp := SUBSTR (strTmp, intPos + 1);
				
				If str���� Is Not Null Then
					intPos := INSTR (str����, ':');
					n_����id := To_Number(SUBSTR (str����, 1, intPos - 1));				
					str���� := SUBSTR (str����, intPos + 1);

					intPos := INSTR (str����, ':');
					n_��ҳid := To_Number(SUBSTR (str����, 1, intPos - 1));
					str���� := SUBSTR (str����,intPos + 1);

					If n_����id>0 And n_��ҳid>0 Then
						Insert Into ������������(����id,����id,��ҳid) values (����id_IN,n_����id,n_��ҳid);
					End If;
				End If;
			End If;      
		END LOOP;
	End If;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_������������_Update;
/

----------------------------------------------------------------------------
---  Delete   for   �������ļ�¼
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_�������ļ�¼_Delete(
	ID_IN		IN	�������ļ�¼.ID%TYPE
)
IS
BEGIN
	Delete From �������ļ�¼ Where ID=ID_IN And ��¼״̬=1;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_�������ļ�¼_Delete;
/

----------------------------------------------------------------------------
---  Authorize   for   �������ļ�¼
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_�������ļ�¼_Authorize(
	ID_IN		IN	�������ļ�¼.ID%TYPE,
	����ʱ��_In	In	�������ļ�¼.����ʱ��%TYPE,
	��������_In	In	�������ļ�¼.��������%TYPE,
	��׼��_In		In	�������ļ�¼.��׼��%TYPE,	
	��׼ʱ��_In	In	�������ļ�¼.��׼ʱ��%TYPE:=Sysdate
)
IS
BEGIN
	Update �������ļ�¼ Set ����ʱ��=����ʱ��_In,��������=��������_In,��׼��=��׼��_In,��׼ʱ��=��׼ʱ��_In,��¼״̬=2 Where ID=ID_IN And ��¼״̬=1;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_�������ļ�¼_Authorize;
/

----------------------------------------------------------------------------
---  Refuse   for   �������ļ�¼
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_�������ļ�¼_Refuse(
	ID_IN		IN	�������ļ�¼.ID%TYPE,
	�ܽ���_In		In	�������ļ�¼.�ܽ���%TYPE,
	�ܽ�����_In	In	�������ļ�¼.�ܽ�����%TYPE,
	�ܽ�ʱ��_In	In	�������ļ�¼.�ܽ�ʱ��%TYPE:=Sysdate
)
IS
BEGIN
	Update �������ļ�¼ Set �ܽ���=�ܽ���_In,�ܽ�����=�ܽ�����_In,�ܽ�ʱ��=�ܽ�ʱ��_In,��¼״̬=3 Where ID=ID_IN And ��¼״̬=1;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_�������ļ�¼_Refuse;
/

----------------------------------------------------------------------------
---  Rollback   for   �������ļ�¼
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_�������ļ�¼_Rollback(
	ID_IN		IN	�������ļ�¼.ID%TYPE,
	��������_In	In	Number
)
IS
BEGIN
	If ��������_In=1 Then
		Update �������ļ�¼ Set ����ʱ��=Null,��������=Null,��׼��=Null,��׼ʱ��=Null,��¼״̬=1 Where ID=ID_IN And ��¼״̬=2;
	ElsIf ��������_In=2 Then
		Update �������ļ�¼ Set �ܽ���=Null,�ܽ�����=Null,�ܽ�ʱ��=Null,��¼״̬=1 Where ID=ID_IN And ��¼״̬=3;
	End If;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_�������ļ�¼_Rollback;
/

----------------------------------------------------------------------------
---  Update   for   ����������¼
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_����������¼_Update(
	ID_In			In	����������¼.ID%TYPE,
	���ID_In		In	����������¼.���ID%TYPE,
	�ύID_In		In	����������¼.�ύID%TYPE,
	����ID_In		In	����������¼.����ID%TYPE,
	��ҳID_In		In	����������¼.��ҳID%TYPE,
	��������_In	In	����������¼.��������%TYPE,
	�ļ�ID_In		In	����������¼.�ļ�ID%TYPE,
	�������_In	In	����������¼.�������%TYPE,
	������ĿID_In	In	����������¼.������ĿID%TYPE,
	������_In		In	����������¼.������%TYPE,
	����ʱ��_In	In	����������¼.����ʱ��%TYPE,
	��������_In	In	����������¼.��������%TYPE
)
IS
BEGIN
	
	Update ����������¼ Set 	�ύID=Decode(�ύID_In,0,Null,�ύID_In),
						����ID=����ID_In,
						��ҳID=��ҳID_In,
						��������=��������_In,
						��¼����=Decode(�ύID_In,Null,1,0,1,2),
						��¼״̬=1,
						�������=�������_In,
						������ĿID=Decode(������ĿID_In,0,Null,������ĿID_In),
						�ļ�ID=Decode(�ļ�ID_In,0,Null,�ļ�ID_In),
						������=������_In,
						����ʱ��=����ʱ��_In,
						��������=��������_In
	Where ID=ID_In;

	If SQL%RowCount=0 Then
		Insert Into ����������¼(ID,���ID,�ύID,����ID,��ҳID,��������,�ļ�ID,��¼����,��¼״̬,�������,������ĿID,������,����ʱ��,��������)
		Values (ID_In,Decode(���ID_In,0,Null,���ID_In),Decode(�ύID_In,0,Null,�ύID_In),����ID_In,��ҳID_In,��������_In,Decode(�ļ�ID_In,0,Null,�ļ�ID_In),Decode(�ύID_In,Null,1,0,1,2),1,�������_In,Decode(������ĿID_In,0,Null,������ĿID_In),������_In,����ʱ��_In,��������_In);
		Update ������ҳ Set ����״̬=4 Where ����id=����ID_In And ��ҳID=��ҳID_In;
	End If;

EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_����������¼_Update;
/
----------------------------------------------------------------------------
---  Finish   for   ����������¼
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_����������¼_Finish(
	ID_In			In	����������¼.ID%TYPE
)
IS
BEGIN	
	Update ����������¼ Set ��¼״̬=3 Where ID=ID_In And ��¼״̬<>3;
	
	If SQL%RowCount>0 Then
		Update ������ҳ a Set ����״̬=3 Where (a.����id,a.��ҳID) In (Select ����id,��ҳid From ����������¼ Where ID=ID_In) And Not Exists (Select 1 From ����������¼ b Where a.����id=b.����id And a.��ҳid=b.��ҳid And ��¼״̬=1);
	End If;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_����������¼_Finish;
/

----------------------------------------------------------------------------
---  RollBackFinish   for   ����������¼
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_����������¼_RollBackFinish(
	ID_In			In	����������¼.ID%TYPE
)
IS
BEGIN	
	Update ����������¼ Set ��¼״̬=Decode(������,Null,1,2) Where ID=ID_In;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_����������¼_RollBackFinish;
/

----------------------------------------------------------------------------
---  Delete   for   ����������¼
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_����������¼_Delete(
	ID_In			In	����������¼.ID%TYPE
)
IS
BEGIN	
	Delete ����������¼ Where ID=ID_In And ��¼״̬=1;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_����������¼_Delete;
/
----------------------------------------------------------------------------
---  Commit   for   �����ύ��¼
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_�����ύ��¼_Commit(
	�ύid_In		In	Varchar2,
	��¼״̬_In	In	�����ύ��¼.��¼״̬%Type,
	������_In		In	�����ύ��¼.������%Type,
	����ʱ��_In	In	�����ύ��¼.����ʱ��%Type:=Sysdate,
	��������_In	In	�����ύ��¼.��������%Type:=Null
)
IS
	v_Tmp			Varchar2(4000);
	n_�ύid			Number(18);
	n_Pos			Number(18);
BEGIN		
	If �ύid_In Is Not Null Then
		v_Tmp := �ύid_In||',';
		WHILE v_Tmp IS NOT NULL LOOP
			n_Pos := INSTR (v_Tmp, ',');
			IF n_Pos >0 Then
				n_�ύid := To_Number(SUBSTR (v_Tmp, 1, n_Pos - 1));
				v_Tmp := SUBSTR (v_Tmp, n_Pos + 1);

				If n_�ύid>0 Then

					For r_List In (Select ID,������ From �����ύ��¼ Where ID=n_�ύid And ��¼״̬<>��¼״̬_In) Loop
						If ��¼״̬_In=3 Then
							--���մ���
							Update ������ҳ Set ����״̬=3 Where (����id,��ҳid) In (Select ����id,��ҳid From �����ύ��¼ Where ID=r_List.ID) And Nvl(����״̬,0)<>3;
							Update �����ύ��¼ Set ��¼״̬=3,������=������_In,����ʱ��=����ʱ��_In Where ID=r_List.ID And ��¼״̬<>3;
						ElsIf ��¼״̬_In=2 Then
							--�ܾ����
							Update ������ҳ Set ����״̬=2 Where (����id,��ҳid) In (Select ����id,��ҳid From �����ύ��¼ Where ID=r_List.ID) And Nvl(����״̬,0)<>2;
							Update �����ύ��¼ Set ��¼״̬=2,������=������_In,����ʱ��=����ʱ��_In,��������=��������_In Where ID=r_List.ID And ��¼״̬<>2;
						ElsIf ��¼״̬_In=5 Then
							--���鵵
							Update ������ҳ Set ����״̬=5 Where (����id,��ҳid) In (Select ����id,��ҳid From �����ύ��¼ Where ID=r_List.ID) And Nvl(����״̬,0)<>5;
							Update �����ύ��¼ Set ��¼״̬=5,�鵵��=������_In,�鵵ʱ��=����ʱ��_In Where ID=r_List.ID And ��¼״̬<>5;
						End If;
					End Loop;
				End If;
			End If;
		End Loop;
	End If;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_�����ύ��¼_Commit;
/
----------------------------------------------------------------------------
---  Receive   for   �����ύ��¼
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_�����ύ��¼_Receive(
	�ύid_In		In	Varchar2,
	������_In		In	�����ύ��¼.������%Type,
	����ʱ��_In	In	�����ύ��¼.����ʱ��%Type:=Sysdate
)
IS
BEGIN	

	If �ύid_In Is Not Null Then
		zl_�����ύ��¼_Commit(�ύid_In,3,������_In,����ʱ��_In);
	End If;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_�����ύ��¼_Receive;
/
----------------------------------------------------------------------------
---  Refuse   for   �����ύ��¼
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_�����ύ��¼_Refuse(
	�ύid_In		In	Varchar2,
	������_In		In	�����ύ��¼.������%Type,
	����ʱ��_In	In	�����ύ��¼.����ʱ��%Type:=Sysdate,
	��������_In	In	�����ύ��¼.��������%Type:=Null
)
IS
BEGIN	

	If �ύid_In Is Not Null Then
		zl_�����ύ��¼_Commit(�ύid_In,2,������_In,����ʱ��_In,��������_In);
	End If;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_�����ύ��¼_Refuse;
/
----------------------------------------------------------------------------
---  Archive   for   �����ύ��¼
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_�����ύ��¼_Archive(
	�ύid_In		In	Varchar2,
	�鵵��_In		In	�����ύ��¼.�鵵��%Type,
	�鵵ʱ��_In	In	�����ύ��¼.�鵵ʱ��%Type:=Sysdate
)
IS
	n_Count			Number(18);
	v_Error			Varchar2(255);
	Err_Custom		Exception;
BEGIN	
	--����Ƿ������еķ������ⶼ�Ѿ������
	Select Count(1) Into n_Count From ����������¼ Where �ύid=�ύid_In And ��¼״̬ In (1,2);
	If n_Count>0 Then
		v_Error:='��ǰ���˻���δ���ķ������⡣';
		Raise Err_Custom;
	End If;

	If �ύid_In Is Not Null Then		
		zl_�����ύ��¼_Commit(�ύid_In,5,�鵵��_In,�鵵ʱ��_In);
	End If;
EXCEPTION
	When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_�����ύ��¼_Archive;
/
----------------------------------------------------------------------------
---  RollBack   for   �����ύ��¼
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_�����ύ��¼_RollBack(
	�ύid_In			In	Varchar2,
	����״̬_In	In	Number:=1
)
IS	
	v_Tmp			Varchar2(4000);
	n_�ύid			Number(18);
	n_Pos			Number(18);
BEGIN		
	If �ύid_In Is Not Null Then
		v_Tmp := �ύid_In||',';
		WHILE v_Tmp IS NOT NULL LOOP
			n_Pos := INSTR (v_Tmp, ',');
			IF n_Pos >0 Then
				n_�ύid := To_Number(SUBSTR (v_Tmp, 1, n_Pos - 1));
				v_Tmp := SUBSTR (v_Tmp, n_Pos + 1);

				If n_�ύid>0 Then

					For r_List In (Select ID,������ From �����ύ��¼ Where ID=n_�ύid And ��¼״̬=����״̬_In) Loop
						If ����״̬_In=3 Then
							--���˽���
							Update ������ҳ Set ����״̬=1 Where (����id,��ҳid) In (Select ����id,��ҳid From �����ύ��¼ Where ID=r_List.ID) And Nvl(����״̬,0)<>1;
							Update �����ύ��¼ Set ��¼״̬=1,������=Null,����ʱ��=Null Where ID=r_List.ID And ��¼״̬<>1;
						ElsIf ����״̬_In=2 Then
							--���˾ܾ�
							Update ������ҳ Set ����״̬=1 Where (����id,��ҳid) In (Select ����id,��ҳid From �����ύ��¼ Where ID=r_List.ID) And Nvl(����״̬,0)<>1;
							Update �����ύ��¼ Set ��¼״̬=1,������=Null,����ʱ��=Null,��������=Null Where ID=r_List.ID And ��¼״̬<>1;
						ElsIf ����״̬_In=5 Then
							--���˹鵵
							If r_List.������ Is Null Then
								Update ������ҳ Set ����״̬=1 Where (����id,��ҳid) In (Select ����id,��ҳid From �����ύ��¼ Where ID=r_List.ID) And Nvl(����״̬,0)<>1;
								Update �����ύ��¼ Set ��¼״̬=1,�鵵��=Null,�鵵ʱ��=Null Where ID=r_List.ID And ��¼״̬<>1;
							Else
								Update ������ҳ Set ����״̬=3 Where (����id,��ҳid) In (Select ����id,��ҳid From �����ύ��¼ Where ID=r_List.ID) And Nvl(����״̬,0)<>3;
								Update �����ύ��¼ Set ��¼״̬=3,�鵵��=Null,�鵵ʱ��=Null Where ID=r_List.ID And ��¼״̬<>3;
							End If;
						End If;
					End Loop;
				End If;
			End If;
		End Loop;
	End If;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_�����ύ��¼_RollBack;
/
----------------------------------------------------------------------------
---  UnReceive   for   �����ύ��¼
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_�����ύ��¼_UnReceive(
	�ύid_In		In	Varchar2
)
IS
BEGIN	
	If �ύid_In Is Not Null Then
		zl_�����ύ��¼_RollBack(�ύid_In,3);
	End If;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_�����ύ��¼_UnReceive;
/
----------------------------------------------------------------------------
---  UnRefuse   for   �����ύ��¼
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_�����ύ��¼_UnRefuse(
	�ύid_In		In	Varchar2
)
IS
BEGIN	

	If �ύid_In Is Not Null Then
		zl_�����ύ��¼_RollBack(�ύid_In,2);
	End If;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_�����ύ��¼_UnRefuse;
/
----------------------------------------------------------------------------
---  UnArchive   for   �����ύ��¼
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_�����ύ��¼_UnArchive(
	�ύid_In		In	Varchar2
)
IS
BEGIN	
	If �ύid_In Is Not Null Then		
		zl_�����ύ��¼_RollBack(�ύid_In,5);
	End If;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_�����ύ��¼_UnArchive;
/

----------------------------------------------------------------------------
---  Lock   for   ��������¼
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_��������¼_Lock(
	����ID_In		In	��������¼.����ID%Type,
	��ҳID_In		In	��������¼.��ҳID%Type,
	�����_In		In	��������¼.�����%Type,
	���ʱ��_In	In	��������¼.���ʱ��%Type:=Sysdate,
	�������_In	In	��������¼.�������%Type:=Null
)
IS
	v_Error			Varchar2(255);
	Err_Custom		Exception;
BEGIN	
	Update ������ҳ Set ���ʱ��=���ʱ��_In Where ����id=����ID_In And ��ҳID=��ҳID_In And ���ʱ�� Is Null;
	If SQL%RowCount=0 Then
		v_Error:='��ǰ�����Ѿ������򲻴��ڵĲ��ˡ�';
		Raise Err_Custom;
	End If;

	Insert Into ��������¼(ID,����ID,��ҳID,��¼״̬,�����,���ʱ��,�������)
	Select  ��������¼_ID.NextVal ,����ID_In,��ҳID_In,1,�����_In,���ʱ��_In,�������_In From Dual;

EXCEPTION
	When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_��������¼_Lock;
/

----------------------------------------------------------------------------
---  UnLock   for   ��������¼
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_��������¼_UnLock(
	����ID_In		In	��������¼.����ID%Type,
	��ҳID_In		In	��������¼.��ҳID%Type
)
IS
	v_Error			Varchar2(255);
	Err_Custom		Exception;
BEGIN		
	Update ������ҳ Set ���ʱ��=Null Where ����id=����ID_In And ��ҳID=��ҳID_In And ���ʱ�� Is Not Null;
	If SQL%RowCount=0 Then
		v_Error:='��ǰ�����Ѿ�������򲻴��ڵĲ��ˡ�';
		Raise Err_Custom;
	End If;

	Update ��������¼ Set ��¼״̬=2 Where ����id=����ID_In And ��ҳID=��ҳID_In And ��¼״̬=1;
EXCEPTION
	When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_��������¼_UnLock;
/

------------------------------------------------------------------------------------------------------------------------------------------
--		������
------------------------------------------------------------------------------------------------------------------------------------------
--����ZL1_INSIDE_1562_1/���������ֽ����
Insert Into zlReports(ID,���,����,˵��,����,��ӡ��,��ֽ,Ʊ��,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��) Values(zlReports_ID.NextVal,'ZL1_INSIDE_1562_1','���������ֽ����','��ӡ���ֽ����','Jf9m dmyc96Rhfo1H*W^','Microsoft Office Document Image Writer',1,0,100,1562,'��ӡ���ֽ����',Sysdate,Sysdate);
Insert Into zlRPTFmts(����ID,���,˵��,W,H,ֽ��,ֽ��,��ֽ̬��,ͼ��) Values(zlReports_ID.CurrVal,1,'���������ֽ��ͳ�Ʊ�1',11904,16838,9,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ2',2,Null,0,'�����1',11,'���λ:[��λ����]',Null,795,1305,1710,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,1);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ7',2,Null,0,'�����1',11,'��Ժ����:[�������ֽ��_����.��Ժ����]',Null,795,2105,3330,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ14',2,Null,0,'�����1',21,'������:[�������ֽ��_����.������]',Null,795,15008,2970,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ16',2,Null,0,'�����1',21,'�����:[�������ֽ��_����.�����]',Null,795,15300,2970,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ3',2,Null,0,'�����1',21,'�Ʊ���:[����Ա����]',Null,795,15592,1710,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ15',2,Null,0,'�����1',22,'����ʱ��:[�������ֽ��_����.����ʱ��]',Null,4482,15008,3330,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ17',2,Null,0,'�����1',22,'���ʱ��:[�������ֽ��_����.���ʱ��]',Null,4482,15300,3330,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ9',2,Null,0,'�����1',11,'ס Ժ ��:[�������ֽ��_����.סԺ��]',Null,795,1590,3150,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ6',2,Null,0,'�����1',12,'סԺ����:[�������ֽ��_����.סԺ����]',Null,4482,1845,3330,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ10',2,Null,0,'�����1',12,'סԺҽʦ:[�������ֽ��_����.סԺҽʦ]',Null,4482,2105,3330,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ1',2,Null,0,'�����1',12,'�������ֽ����',Null,4835,675,2625,360,0,0,1,'����_GB2312',18,1,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ4',2,Null,0,'�����1',22,'��[ҳ��]ҳ',Null,5698,15592,900,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ13',2,Null,0,'�����1',13,'��    ��:[�������ֽ��_����.�ȼ�]',Null,8530,1590,2970,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ12',2,Null,0,'�����1',13,'��    ��:[�������ֽ��_����.�ܷ�]',Null,8530,1845,2970,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ11',2,Null,0,'�����1',13,'��Ժ����:[�������ֽ��_����.��Ժ����]',Null,8170,2105,3330,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ5',2,Null,0,'�����1',23,'�������ڣ�[YYYY-MM-DD]',Null,9520,15592,1980,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'�����1',4,Null,0,Null,0,'����������ϸ_����',Null,795,2385,10705,12478,255,0,1,'����',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-1,0,Null,Null,'[����������ϸ_����.��Ŀ]','4^435^��Ŀ',0,0,930,0,255,1,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-2,1,Null,Null,'[����������ϸ_����.��׼��ֵ]','4^435^��׼��ֵ',0,0,555,0,255,1,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-3,2,Null,Null,'[����������ϸ_����.����Ҫ��]','4^435^����Ҫ��',0,0,2145,0,255,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-4,3,Null,Null,'[����������ϸ_����.ȱ������]','4^435^ȱ������',0,0,3540,0,255,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-5,4,Null,Null,'[����������ϸ_����.�۷ֱ�׼]','4^435^�۷ֱ�׼',0,0,1140,0,255,1,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-6,5,Null,Null,'[����������ϸ_����.����]','4^435^����',0,0,1140,0,255,1,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ18',2,Null,0,'�����1',12,'�����޸�:[�������ֽ��_����.�����޸�]',Null,4482,1590,3330,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ8',2,Null,0,'�����1',11,'��    ��:[�������ֽ��_����.����]',Null,795,1845,2970,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����,˵��) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'�������ֽ��_����','ID,131|סԺ����,131|��Ժ����,200|����,200|סԺ��,131|סԺҽʦ,200|��Ժ����,200|�ܷ�,200|�ȼ�,200|�����޸�,200|������,200|����ʱ��,200|�����,200|���ʱ��,200',User||'.���ű�,'||User||'.�������ֽ��,'||User||'.������ҳ,'||User||'.������Ϣ',0,Null);
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,1,'select A.ID,A.��ҳID as סԺ����,(select ���ű�.���� from ���ű� where ���ű�.id=B.��Ժ����ID) as ��Ժ����,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,2,'C.����,C.סԺ��,B.סԺҽʦ,TO_CHAR(B.��Ժ����,''YYYY-MM-DD'') as ��Ժ����,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,3,'decode(A.�ȼ�,''��'',''-     '',A.�ܷ�) as �ܷ�,decode(A.�ȼ�,''��'',''���ϸ�'',A.�ȼ�) as �ȼ�,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,4,'decode(A.�����޸�,null,''��'',0,''��'',1,''��'') as �����޸�,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,5,'A.������,TO_CHAR(A.����ʱ��,''YYYY-MM-DD'') as ����ʱ��,A.�����,TO_CHAR(A.���ʱ��,''YYYY-MM-DD'') as ���ʱ�� ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,6,'from �������ֽ�� A, ������ҳ B,������Ϣ C');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,7,'where A.����ID=C.����ID and A.����ID= B.����ID and A.��ҳID=B.��ҳID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,8,'and ID=[0]');
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,Null,0,'���ID',1,Null,0,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����,˵��) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'����������ϸ_����','ID,131|��Ŀ,200|��׼��ֵ,200|����Ҫ��,200|ȱ������,200|�۷ֱ�׼,200|����,200',User||'.����������ϸ,'||User||'.�������ֽ��,'||User||'.�������ֱ�׼��ͼ',0,Null);
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,1,'select A.ID,A.��Ŀ,TO_CHAR(A.��׼��ֵ)||''��'' as ��׼��ֵ,A.����Ҫ��,A.ȱ������,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,2,'decode(A.�۷ֱ�׼,''��'',''�׼�'',''��'',''�Ҽ�'',''��'',''����'',''��'',''������'',A.�۷ֱ�׼) as �۷ֱ�׼,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,3,'(');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,4,'select decode(ȱ�ݵȼ�,null,to_CHAR(�������),''��'',''������'',ȱ�ݵȼ�||''��'') ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,5,'from ����������ϸ where ���ֱ�׼ID=A.ID and ����ID=[0]');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,6,') as ����  ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,7,'from �������ֱ�׼��ͼ A  ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,8,'where A.����=''��'' and A.����ID=(select B.����ID from �������ֽ�� B where B.ID=[0])  ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,9,'order by A.�ϼ�ID,A.ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,10,Null);
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,Null,0,'���ID',1,Null,0,Null,Null,Null,Null,Null,Null);

--����ZL1_REPORT_1570/ҽʦ��������ͳ�Ʊ�
Insert Into zlReports(ID,���,����,˵��,����,��ӡ��,��ֽ,Ʊ��,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��) Values(zlReports_ID.NextVal,'ZL1_REPORT_1570','ҽʦ��������ͳ�Ʊ�','ҽʦ��������ͳ�Ʊ�','Ww:mXk|ws35VitmcW*O]','Microsoft Office Document Image Writer',1,0,100,1570,'����',Sysdate,Sysdate);
Insert Into zlRPTFmts(����ID,���,˵��,W,H,ֽ��,ֽ��,��ֽ̬��,ͼ��) Values(zlReports_ID.CurrVal,1,'ҽʦ��������ͳ�Ʊ�1',11904,16838,9,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ2',2,Null,0,'���ܱ�1',11,'���λ:[��λ����]',Null,270,935,1710,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,1);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ3',2,Null,0,'���ܱ�1',21,'�Ʊ���:[����Ա����]',Null,270,15325,1710,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ1',2,Null,0,Null,0,'סԺҽʦ��������ͳ�Ʊ�',Null,3822,395,4125,360,0,0,1,'����_GB2312',18,1,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ5',2,Null,0,'���ܱ�1',22,'��[ҳ��]ҳ',Null,5295,15325,900,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ6',2,Null,0,'���ܱ�1',13,'ͳ��ʱ��:[=��ʼ����] �� [=��������]',Null,8070,935,3150,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ4',2,Null,0,'���ܱ�1',23,'�������ڣ�[YYYY-MM-DD]',Null,9240,15325,1980,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'���ܱ�1',5,Null,0,Null,0,'������������_����',Null,270,1245,10950,13980,255,0,0,'����',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,7,zlRPTItems_ID.CurrVal-1,0,Null,Null,'סԺҽʦ',Null,0,0,795,0,255,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-2,0,Null,Null,'��������',Null,0,0,945,0,255,2,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-3,1,Null,Null,'��������',Null,0,0,930,0,255,2,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-4,2,Null,Null,'���󲡰�',Null,0,0,885,0,255,2,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-5,3,Null,Null,'�׵�',Null,0,0,855,0,255,2,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-6,4,Null,Null,'�ҵ�',Null,0,0,810,0,255,2,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-7,5,Null,Null,'����',Null,0,0,810,0,255,2,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-8,6,Null,Null,'�����޸���',Null,0,0,930,0,255,2,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-9,7,Null,Null,'�׵���',Null,0,0,915,0,255,2,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-10,8,Null,Null,'�ҵ���',Null,0,0,930,0,255,2,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-11,9,Null,Null,'������',Null,0,0,900,0,255,2,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-12,10,Null,Null,'�����޸���',Null,0,0,1005,0,255,2,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����,˵��) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'������������_����','סԺҽʦ,200|��������,139|��������,139|���󲡰�,139|�׵�,139|�ҵ�,139|����,139|�����޸���,139|�׵���,139|�ҵ���,139|������,139|�����޸���,139',User||'.��������������ͼ',1,Null);
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,1,'Select סԺҽʦ,count(*) as ��������,count(����ʱ��) as ��������,count(���ʱ��) as ���󲡰�,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,2,'       count(decode(�ȼ�,''��'',���ʱ��,null)) as �׵�,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,3,'       count(decode(�ȼ�,''��'',���ʱ��,null)) as �ҵ�,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,4,'       count(decode(�ȼ�,''��'',���ʱ��,null)) as ����,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,5,'       count(decode(�����޸�,1,���ʱ��,null)) as �����޸���,       ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,6,'       round(decode(count(���ʱ��),0,0,count(decode(�ȼ�,''��'',���ʱ��,null))/count(���ʱ��))*100,1) as �׵���,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,7,'       round(decode(count(���ʱ��),0,0,count(decode(�ȼ�,''��'',���ʱ��,null))/count(���ʱ��))*100,1) as �ҵ���,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,8,'       round(decode(count(���ʱ��),0,0,count(decode(�ȼ�,''��'',���ʱ��,null))/count(���ʱ��))*100,1) as ������,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,9,'       round(decode(count(���ʱ��),0,0,count(decode(�����޸�,1,���ʱ��,null))/count(���ʱ��))*100,1) as �����޸���');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,10,'From ��������������ͼ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,11,'where ��Ժ���� between [0] and [1]+1 ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,12,'group by סԺҽʦ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,13,Null);
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,14,'union all');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,15,'Select ''�ϼ�'',count(*) as ��������,count(����ʱ��) as ��������,count(���ʱ��) as ���󲡰�,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,16,'       count(decode(�ȼ�,''��'',���ʱ��,null)) as �׵�,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,17,'       count(decode(�ȼ�,''��'',���ʱ��,null)) as �ҵ�,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,18,'       count(decode(�ȼ�,''��'',���ʱ��,null)) as ����,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,19,'       count(decode(�����޸�,1,���ʱ��,null)) as �����޸���,       ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,20,'       round(decode(count(���ʱ��),0,0,count(decode(�ȼ�,''��'',���ʱ��,null))/count(���ʱ��))*100,1) as �׵���,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,21,'       round(decode(count(���ʱ��),0,0,count(decode(�ȼ�,''��'',���ʱ��,null))/count(���ʱ��))*100,1) as �ҵ���,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,22,'       round(decode(count(���ʱ��),0,0,count(decode(�ȼ�,''��'',���ʱ��,null))/count(���ʱ��))*100,1) as ������,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,23,'       round(decode(count(���ʱ��),0,0,count(decode(�����޸�,1,���ʱ��,null))/count(���ʱ��))*100,1) as �����޸���');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,24,'From ��������������ͼ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,25,'where ��Ժ���� between [0] and [1]+1 ');
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,Null,0,'��ʼ����',2,CHR(38)||'ǰһ������',0,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,Null,1,'��������',2,CHR(38)||'��ǰ����',0,Null,Null,Null,Null,Null,Null);

--����ZL1_REPORT_1571/���Ҳ�������ͳ�Ʊ�
Insert Into zlReports(ID,���,����,˵��,����,��ӡ��,��ֽ,Ʊ��,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��) Values(zlReports_ID.NextVal,'ZL1_REPORT_1571','���Ҳ�������ͳ�Ʊ�','���Ҳ�������ͳ�Ʊ�','Ew:mXj|ws!5VitlcW*]]','Microsoft Office Document Image Writer',1,0,100,1571,'����',Sysdate,Sysdate);
Insert Into zlRPTFmts(����ID,���,˵��,W,H,ֽ��,ֽ��,��ֽ̬��,ͼ��) Values(zlReports_ID.CurrVal,1,'���Ҳ�������ͳ�Ʊ�1',11904,16838,9,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ2',2,Null,0,'���ܱ�1',11,'���λ:[��λ����]',Null,255,1190,1710,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,1);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ3',2,Null,0,'���ܱ�1',21,'�Ʊ���:[����Ա����]',Null,255,15295,1710,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ1',2,Null,0,Null,0,'���Ҳ�������ͳ�Ʊ�',Null,4205,645,3375,360,0,1,1,'����_GB2312',18,1,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ5',2,Null,0,'���ܱ�1',22,'��[ҳ��]ҳ',Null,5385,15295,900,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ6',2,Null,0,'���ܱ�1',13,'ͳ��ʱ��:[=��ʼ����] �� [=��������]',Null,8265,1190,3150,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ4',2,Null,0,'���ܱ�1',23,'�������ڣ�[YYYY-MM-DD]',Null,9435,15295,1980,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'���ܱ�1',5,Null,0,Null,0,'������������_����',Null,255,1500,11160,13695,255,0,0,'����',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,7,zlRPTItems_ID.CurrVal-1,0,Null,Null,'����',Null,0,0,840,0,255,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-2,0,Null,Null,'��������',Null,0,0,855,0,255,2,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-3,1,Null,Null,'��������',Null,0,0,930,0,255,2,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-4,2,Null,Null,'���󲡰�',Null,0,0,900,0,255,2,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-5,3,Null,Null,'�׵�',Null,0,0,840,0,255,2,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-6,4,Null,Null,'�ҵ�',Null,0,0,870,0,255,2,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-7,5,Null,Null,'����',Null,0,0,825,0,255,2,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-8,6,Null,Null,'�����޸���',Null,0,0,1020,0,255,2,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-9,7,Null,Null,'�׵���',Null,0,0,1005,0,255,2,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-10,8,Null,Null,'�ҵ���',Null,0,0,1005,0,255,2,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-11,9,Null,Null,'������',Null,0,0,1005,0,255,2,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-12,10,Null,Null,'�����޸���',Null,0,0,1005,0,255,2,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����,˵��) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'������������_����','����,200|��������,139|��������,139|���󲡰�,139|�׵�,139|�ҵ�,139|����,139|�����޸���,139|�׵���,139|�ҵ���,139|������,139|�����޸���,139',User||'.��������������ͼ',1,Null);
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,1,'Select ��Ժ���� as ����,count(*) as ��������,count(����ʱ��) as ��������,count(���ʱ��) as ���󲡰�,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,2,'       count(decode(�ȼ�,''��'',���ʱ��,null)) as �׵�,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,3,'       count(decode(�ȼ�,''��'',���ʱ��,null)) as �ҵ�,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,4,'       count(decode(�ȼ�,''��'',���ʱ��,null)) as ����,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,5,'       count(decode(�����޸�,1,���ʱ��,null)) as �����޸���,       ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,6,'       round(decode(count(���ʱ��),0,0,count(decode(�ȼ�,''��'',���ʱ��,null))/count(���ʱ��))*100,1) as �׵���,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,7,'       round(decode(count(���ʱ��),0,0,count(decode(�ȼ�,''��'',���ʱ��,null))/count(���ʱ��))*100,1) as �ҵ���,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,8,'       round(decode(count(���ʱ��),0,0,count(decode(�ȼ�,''��'',���ʱ��,null))/count(���ʱ��))*100,1) as ������,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,9,'       round(decode(count(���ʱ��),0,0,count(decode(�����޸�,1,���ʱ��,null))/count(���ʱ��))*100,1) as �����޸���');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,10,'From ��������������ͼ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,11,'where ��Ժ���� between [0] and [1]+1 ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,12,'group by ��Ժ����');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,13,Null);
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,14,'union all');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,15,'Select ''�ϼ�'',count(*) as ��������,count(����ʱ��) as ��������,count(���ʱ��) as ���󲡰�,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,16,'       count(decode(�ȼ�,''��'',���ʱ��,null)) as �׵�,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,17,'       count(decode(�ȼ�,''��'',���ʱ��,null)) as �ҵ�,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,18,'       count(decode(�ȼ�,''��'',���ʱ��,null)) as ����,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,19,'       count(decode(�����޸�,1,���ʱ��,null)) as �����޸���,       ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,20,'       round(decode(count(���ʱ��),0,0,count(decode(�ȼ�,''��'',���ʱ��,null))/count(���ʱ��))*100,1) as �׵���,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,21,'       round(decode(count(���ʱ��),0,0,count(decode(�ȼ�,''��'',���ʱ��,null))/count(���ʱ��))*100,1) as �ҵ���,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,22,'       round(decode(count(���ʱ��),0,0,count(decode(�ȼ�,''��'',���ʱ��,null))/count(���ʱ��))*100,1) as ������,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,23,'       round(decode(count(���ʱ��),0,0,count(decode(�����޸�,1,���ʱ��,null))/count(���ʱ��))*100,1) as �����޸���');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,24,'From ��������������ͼ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,25,'where ��Ժ���� between [0] and [1]+1 ');
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,Null,0,'��ʼ����',2,CHR(38)||'ǰһ������',0,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,Null,1,'��������',2,CHR(38)||'��ǰ����',0,Null,Null,Null,Null,Null,Null);

--����ZL1_REPORT_1572/�������ֽ���嵥
Insert Into zlReports(ID,���,����,˵��,����,��ӡ��,��ֽ,Ʊ��,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��) Values(zlReports_ID.NextVal,'ZL1_REPORT_1572','�������ֽ���嵥','�������ֽ���嵥','Le(jHbyys+6Rbirs_)FH','Microsoft Office Document Image Writer',1,0,100,1572,'����',Sysdate,Sysdate);
Insert Into zlRPTFmts(����ID,���,˵��,W,H,ֽ��,ֽ��,��ֽ̬��,ͼ��) Values(zlReports_ID.CurrVal,1,'11',11904,16838,9,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ2',2,Null,0,'�����1',11,'���λ:[��λ����]',Null,345,995,1710,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,1);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ3',2,Null,0,'�����1',21,'�Ʊ���:[����Ա����]',Null,345,15370,1710,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ1',2,Null,0,Null,0,'�������ֽ���嵥',Null,4015,525,3000,360,0,0,1,'����_GB2312',18,1,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ5',2,Null,0,'�����1',22,'��[ҳ��]ҳ',Null,5452,15370,900,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ6',2,Null,0,'�����1',13,'ͳ��ʱ��:[=��ʼ����] �� [=��������]',Null,8309,995,3150,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ4',2,Null,0,'�����1',23,'�������ڣ�[YYYY-MM-DD]',Null,9479,15370,1980,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'�����1',4,Null,0,Null,0,'������������_����',Null,345,1305,11114,13965,255,0,0,'����',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-1,0,Null,Null,'[������������_����.סԺ��]','4^255^סԺ��',0,0,960,0,255,2,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-2,1,Null,Null,'[������������_����.����]','4^255^����',0,0,825,0,255,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-3,2,Null,Null,'[������������_����.��Ժ����]','4^255^��Ժ����',0,0,1005,0,255,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-4,3,Null,Null,'[������������_����.���λ�ʿ]','4^255^���λ�ʿ',0,0,1005,0,255,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-5,4,Null,Null,'[������������_����.סԺҽʦ]','4^255^סԺҽʦ',0,0,900,0,255,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-6,5,Null,Null,'[������������_����.�ܷ�]','4^255^�ܷ�',0,0,675,0,255,2,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-7,6,Null,Null,'[������������_����.�ȼ�]','4^255^�ȼ�',0,0,585,0,255,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-8,7,Null,Null,'[������������_����.�����޸�]','4^255^�����޸�',0,0,885,0,255,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-9,8,Null,Null,'[������������_����.������]','4^255^������',0,0,1005,0,255,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-10,9,Null,Null,'[������������_����.����ʱ��]','4^255^����ʱ��',0,0,1005,0,255,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-11,10,Null,Null,'[������������_����.�����]','4^255^�����',0,0,1005,0,255,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-12,11,Null,Null,'[������������_����.���ʱ��]','4^255^���ʱ��',0,0,1005,0,255,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����,˵��) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'������������_����','סԺ��,131|����,200|��Ժ����,200|���λ�ʿ,200|סԺҽʦ,200|�ܷ�,131|�ȼ�,200|�����޸�,200|������,200|����ʱ��,200|�����,200|���ʱ��,200',User||'.��������������ͼ',0,Null);
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,1,'Select סԺ��,����,��Ժ����,���λ�ʿ,סԺҽʦ,�ܷ�,�ȼ�,decode(�����޸�,1,''��'','''') as �����޸�,������,����ʱ��,�����,���ʱ��');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,2,'From ��������������ͼ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,3,'where ����ʱ�� is not null and ��Ժ���� between [0] and [1]+1 ');
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,Null,0,'��ʼ����',2,CHR(38)||'ǰһ������',0,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,Null,1,'��������',2,CHR(38)||'��ǰ����',0,Null,Null,Null,Null,Null,Null);

--����ZL1_INSIDE_1562_1/���������ֽ����
Insert into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1562,'��ӡ���ֽ����',Null);
Insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1562,'��ӡ���ֽ����',User,'�������ֱ�׼��ͼ','SELECT');
Insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1562,'��ӡ���ֽ����',User,'�������ֽ��','SELECT');
Insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1562,'��ӡ���ֽ����',User,'����������ϸ','SELECT');
Insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1562,'��ӡ���ֽ����',User,'������ҳ','SELECT');
Insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1562,'��ӡ���ֽ����',User,'������Ϣ','SELECT');
Insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1562,'��ӡ���ֽ����',User,'���ű�','SELECT');
--����ZL1_REPORT_1570/ҽʦ��������ͳ�Ʊ�
Insert into zlPrograms(���,����,˵��,ϵͳ,����) Values(1570,'ҽʦ��������ͳ�Ʊ�','ҽʦ��������ͳ�Ʊ�',100,'zl9Report');
Insert into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1570,'����',Null);
Insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1570,'����',User,'��������������ͼ','SELECT');
Insert into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Select 'ȱʡ',zlMenus_ID.Nextval,ID,'ҽʦ��������ͳ�Ʊ�','ҽʦ��������ͳ�Ʊ�',Null,105,'ҽʦ��������ͳ�Ʊ�',100,1570 From zlMenus Where ϵͳ=100 And ���='ȱʡ' And ����='�����������' And ģ�� is NULL;

--����ZL1_REPORT_1571/���Ҳ�������ͳ�Ʊ�
Insert into zlPrograms(���,����,˵��,ϵͳ,����) Values(1571,'���Ҳ�������ͳ�Ʊ�','���Ҳ�������ͳ�Ʊ�',100,'zl9Report');
Insert into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1571,'����',Null);
Insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1571,'����',User,'��������������ͼ','SELECT');
Insert into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Select 'ȱʡ',zlMenus_ID.Nextval,ID,'���Ҳ�������ͳ�Ʊ�','���Ҳ�������ͳ�Ʊ�',Null,105,'���Ҳ�������ͳ�Ʊ�',100,1571 From zlMenus Where ϵͳ=100 And ���='ȱʡ' And ����='�����������' And ģ�� is NULL;

--����ZL1_REPORT_1572/�������ֽ���嵥
Insert into zlPrograms(���,����,˵��,ϵͳ,����) Values(1572,'�������ֽ���嵥','�������ֽ���嵥',100,'zl9Report');
Insert into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1572,'����',Null);
Insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1572,'����',User,'��������������ͼ','SELECT');
Insert into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Select 'ȱʡ',zlMenus_ID.Nextval,ID,'�������ֽ���嵥','�������ֽ���嵥',Null,105,'�������ֽ���嵥',100,1572 From zlMenus Where ϵͳ=100 And ���='ȱʡ' And ����='�����������' And ģ�� is NULL;
