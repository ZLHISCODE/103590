--------------------------------------------------------------------------------------------------------------------------
--�ٴ�·�����ݽṹ����
--------------------------------------------------------------------------------------------------------------------------
Create Table �ٴ���������(
    ���� VARCHAR2(1),
    ���� VARCHAR2(20),
    ���� VARCHAR2(10))
    TABLESPACE zl9BaseItem
    PCTFREE 5  
		PCTUSED 85;
Alter Table �ٴ��������� Add Constraint �ٴ���������_PK Primary Key (����) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table �ٴ��������� Add Constraint �ٴ���������_UQ_���� Unique (����) Using Index Pctfree 5 Tablespace zl9indexcis;


Create Table ·���������(
    ���� VARCHAR2(5),
    ���� VARCHAR2(100),
    ���� VARCHAR2(10),
		�ϼ� VARCHAR2(5),
		ĩ�� NUMBER(1))
    TABLESPACE zl9BaseItem
    PCTFREE 5  
		PCTUSED 85;
Alter Table ·��������� Add Constraint ·���������_PK Primary Key (����) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table ·��������� Add Constraint ·���������_UQ_���� Unique (�ϼ�,����) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table ·��������� Add Constraint ·���������_CK_ĩ�� Check (ĩ�� in(0,1));


Create Table ���쳣��ԭ��(
    ���� VARCHAR2(5),
    ���� VARCHAR2(100),
    ���� VARCHAR2(10),
		�ϼ� VARCHAR2(5),
		ĩ�� NUMBER(1))
    TABLESPACE zl9BaseItem
    PCTFREE 5  
		PCTUSED 85;
Alter Table ���쳣��ԭ�� Add Constraint ���쳣��ԭ��_PK Primary Key (����) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table ���쳣��ԭ�� Add Constraint ���쳣��ԭ��_UQ_���� Unique (�ϼ�,����) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table ���쳣��ԭ�� Add Constraint ���쳣��ԭ��_CK_ĩ�� Check (ĩ�� in(0,1));


Create Sequence �ٴ�·��ͼ��_ID Start With 1;
CREATE TABLE �ٴ�·��ͼ��(
		ID NUMBER(18),
		ͼ�� BLOB,
		���� NUMBER(1))
		LOB(ͼ��) Store as (Cache)
    TABLESPACE zl9BaseItem
    PCTFREE 20
    PCTUSED 70;
Alter Table �ٴ�·��ͼ�� Add Constraint �ٴ�·��ͼ��_PK Primary Key (ID) Using Index Pctfree 5 Tablespace zl9indexcis;

Create Sequence �ٴ�·��Ŀ¼_ID Start With 1;
CREATE TABLE �ٴ�·��Ŀ¼(
    ID NUMBER(18),
		���� VARCHAR2(50),
    ���� VARCHAR2(5),
    ���� VARCHAR2(100),
    ͨ�� NUMBER(1),
    ���°汾 NUMBER(3),
    �������� VARCHAR2(20),
    ���ò��� VARCHAR2(20),
		�����Ա� NUMBER(1),
		�������� VARCHAR2(10),
    ˵�� VARCHAR2(200))
    TABLESPACE zl9BaseItem
    PCTFREE 5
    PCTUSED 85;
Alter Table �ٴ�·��Ŀ¼ Add Constraint �ٴ�·��Ŀ¼_PK Primary Key (ID) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table �ٴ�·��Ŀ¼ Add Constraint �ٴ�·��Ŀ¼_UQ_���� Unique (����,����) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table �ٴ�·��Ŀ¼ Add Constraint �ٴ�·��Ŀ¼_UQ_���� Unique (����,����) Using Index Pctfree 5 Tablespace zl9indexcis;


CREATE TABLE �ٴ�·������(
    ·��ID NUMBER(18),
    ����ID NUMBER(18),
		���ID NUMBER(18))
    TABLESPACE zl9BaseItem
    PCTFREE 5
    PCTUSED 85;
Alter Table �ٴ�·������ Add Constraint �ٴ�·������_UQ_����ID Unique (·��ID,����ID,���ID) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table �ٴ�·������ Add Constraint �ٴ�·������_FK_·��ID Foreign Key (·��ID) References �ٴ�·��Ŀ¼(ID) On Delete Cascade;
Alter Table �ٴ�·������ Add Constraint �ٴ�·������_FK_����ID Foreign Key (����ID) References ��������Ŀ¼(ID) On Delete Cascade;
Alter Table �ٴ�·������ Add Constraint �ٴ�·������_FK_���ID Foreign Key (���ID) References �������Ŀ¼(ID) On Delete Cascade;


CREATE TABLE �ٴ�·������(
    ·��ID NUMBER(18),
    ����ID NUMBER(18))
    TABLESPACE zl9BaseItem
    PCTFREE 5
    PCTUSED 85;
Alter Table �ٴ�·������ Add Constraint �ٴ�·������_PK Primary Key (·��ID,����ID) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table �ٴ�·������ Add Constraint �ٴ�·������_FK_·��ID Foreign Key (·��ID) References �ٴ�·��Ŀ¼(ID) On Delete Cascade;
Alter Table �ٴ�·������ Add Constraint �ٴ�·������_FK_����ID Foreign Key (����ID) References ���ű�(ID) On Delete Cascade;


CREATE TABLE �ٴ�·���ļ�(
    ·��ID NUMBER(18),
		�ļ��� VARCHAR2(200),
    ���� BLOB,
		������ VARCHAR2(20),
		����ʱ�� DATE)
    TABLESPACE zl9BaseItem
    PCTFREE 20
    PCTUSED 70;
Alter Table �ٴ�·���ļ� Add Constraint �ٴ�·���ļ�_PK Primary Key (·��ID,�ļ���) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table �ٴ�·���ļ� Add Constraint �ٴ�·���ļ�_FK_·��ID Foreign Key (·��ID) References �ٴ�·��Ŀ¼(ID) On Delete Cascade;

CREATE TABLE �ٴ�·���汾(
    ·��ID NUMBER(18),
    �汾�� NUMBER(3),
    ��׼סԺ�� VARCHAR2(10),
    ��׼���� VARCHAR2(20),
    �汾˵�� VARCHAR2(200),
    ������ VARCHAR2(20),
    ����ʱ�� DATE,
    ����� VARCHAR2(20),
    ���ʱ�� DATE,
		ͣ���� VARCHAR2(20),
    ͣ��ʱ�� DATE)
    TABLESPACE zl9BaseItem
    PCTFREE 5
    PCTUSED 85;
Alter Table �ٴ�·���汾 Add Constraint �ٴ�·���汾_PK Primary Key (·��ID,�汾��) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table �ٴ�·���汾 Add Constraint �ٴ�·���汾_FK_·��ID Foreign Key (·��ID) References �ٴ�·��Ŀ¼(ID) On Delete Cascade;


Create Sequence �ٴ�·���׶�_ID Start With 1;
CREATE TABLE �ٴ�·���׶�(
		ID NUMBER(18),
    ·��ID NUMBER(18),
    �汾�� NUMBER(3),
		��ID NUMBER(18),
    ��� NUMBER(5),
    ���� VARCHAR2(50),
    ��ʼ���� NUMBER(3),
    �������� NUMBER(3),
    ��־ VARCHAR2(10),
    ˵�� VARCHAR2(200))
    TABLESPACE zl9BaseItem
    PCTFREE 5
    PCTUSED 85;
Alter Table �ٴ�·���׶� Add Constraint �ٴ�·���׶�_PK Primary Key (ID) Using Index Pctfree 5 Tablespace zl9indexcis;
--���ʹ���ӳ�Լ��Ϊ��·����������������֮ǰ��Ų��ظ����
Alter Table �ٴ�·���׶� Add Constraint �ٴ�·���׶�_UQ_��� Unique (·��ID,�汾��,��ID,���) Deferrable Initially Deferred Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table �ٴ�·���׶� Add Constraint �ٴ�·���׶�_FK_�汾�� Foreign Key (·��ID,�汾��) References �ٴ�·���汾(·��ID,�汾��) On Delete Cascade;
Alter Table �ٴ�·���׶� Add Constraint �ٴ�·���׶�_FK_��ID Foreign Key (��ID) References �ٴ�·���׶�(ID) On Delete Cascade;


CREATE TABLE �ٴ�·������(
    ·��ID NUMBER(18),
    �汾�� NUMBER(3),
    ��� NUMBER(5),
		���� VARCHAR2(50))
    TABLESPACE zl9BaseItem
    PCTFREE 5
    PCTUSED 85;
Alter Table �ٴ�·������ Add Constraint �ٴ�·������_PK Primary Key (·��ID,�汾��,���) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table �ٴ�·������ Add Constraint �ٴ�·������_UQ_���� Unique (·��ID,�汾��,����) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table �ٴ�·������ Add Constraint �ٴ�·������_FK_�汾�� Foreign Key (·��ID,�汾��) References �ٴ�·���汾(·��ID,�汾��) On Delete Cascade;

Create Sequence �ٴ�·����Ŀ_ID Start With 1;
CREATE TABLE �ٴ�·����Ŀ(
		ID NUMBER(18),
    ·��ID NUMBER(18),
    �汾�� NUMBER(3),
    �׶�ID NUMBER(18),
		���� VARCHAR2(50),
		��Ŀ��� NUMBER(5),
		��Ŀ���� VARCHAR2(1000),
		ִ�з�ʽ NUMBER(1),
		ִ���� NUMBER(1),
		��Ŀ��� VARCHAR2(500),
		ͼ��ID NUMBER(18))
    TABLESPACE zl9BaseItem
    PCTFREE 5
    PCTUSED 85;
Alter Table �ٴ�·����Ŀ Add Constraint �ٴ�·����Ŀ_PK Primary Key (ID) Using Index Pctfree 5 Tablespace zl9indexcis;
--���ʹ���ӳ�Լ��Ϊ��·����������������֮ǰ��Ų��ظ����
Alter Table �ٴ�·����Ŀ Add Constraint �ٴ�·����Ŀ_UQ_��Ŀ��� Unique (�׶�ID,����,��Ŀ���) Deferrable Initially Deferred Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table �ٴ�·����Ŀ Add Constraint �ٴ�·����Ŀ_UQ_��Ŀ���� Unique (�׶�ID,��Ŀ����) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table �ٴ�·����Ŀ Add Constraint �ٴ�·����Ŀ_FK_�汾�� Foreign Key (·��ID,�汾��) References �ٴ�·���汾(·��ID,�汾��) On Delete Cascade;
Alter Table �ٴ�·����Ŀ Add Constraint �ٴ�·����Ŀ_FK_�׶�ID Foreign Key (�׶�ID) References �ٴ�·���׶�(ID) On Delete Cascade;
Alter Table �ٴ�·����Ŀ Add Constraint �ٴ�·����Ŀ_FK_ͼ��ID Foreign Key (ͼ��ID) References �ٴ�·��ͼ��(ID);
Create Index �ٴ�·����Ŀ_IX_�汾�� On �ٴ�·����Ŀ(·��ID,�汾��) Pctfree 5 Tablespace zl9indexcis
/
Create Index �ٴ�·����Ŀ_IX_�׶�ID On �ٴ�·����Ŀ(�׶�ID) Pctfree 5 Tablespace zl9indexcis
/
Create Index �ٴ�·����Ŀ_IX_ͼ��ID On �ٴ�·����Ŀ(ͼ��ID) Pctfree 5 Tablespace zl9indexcis
/


Create Sequence ·��ҽ������_ID Start With 1;
CREATE TABLE ·��ҽ������(
		ID NUMBER(18),
    ���ID NUMBER(18),
    ��� NUMBER(5),
    ��Ч NUMBER(1),
    ������ĿID NUMBER(18),
		�շ�ϸĿID NUMBER(18),
		ҽ������ VARCHAR2(1000),
		�������� NUMBER(16,5),
		�ܸ����� NUMBER(16,5),
		�걾��λ VARCHAR2(60),
		��鷽�� VARCHAR2(30),
		ҽ������ VARCHAR2(1000),
		ִ��Ƶ�� VARCHAR2(20),
		Ƶ�ʴ��� NUMBER(3),
		Ƶ�ʼ�� NUMBER(3),
		�����λ VARCHAR2(4),
		ִ������ NUMBER(1),
		ִ�п���ID NUMBER(18),
		ʱ�䷽�� VARCHAR2(50))
    TABLESPACE zl9BaseItem
    PCTFREE 5
    PCTUSED 85;
Alter Table ·��ҽ������ Add Constraint ·��ҽ������_PK Primary Key (ID) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table ·��ҽ������ Add Constraint ·��ҽ������_FK_���ID Foreign Key (���ID) References ·��ҽ������(ID) Deferrable Initially Deferred;
Alter Table ·��ҽ������ Add Constraint ·��ҽ������_FK_������ĿID Foreign Key (������ĿID) References ������ĿĿ¼(ID);
Alter Table ·��ҽ������ Add Constraint ·��ҽ������_FK_�շ�ϸĿID Foreign Key (�շ�ϸĿID) References �շ���ĿĿ¼(ID);
Alter Table ·��ҽ������ Add Constraint ·��ҽ������_FK_ִ�п���ID Foreign Key (ִ�п���ID) References ���ű�(ID);
Create Index ·��ҽ������_IX_���ID On ·��ҽ������(���ID) Pctfree 5 Tablespace zl9indexcis
/
Create Index ·��ҽ������_IX_������ĿID On ·��ҽ������(������ĿID) Pctfree 5 Tablespace zl9indexcis
/
Create Index ·��ҽ������_IX_�շ�ϸĿID On ·��ҽ������(�շ�ϸĿID) Pctfree 5 Tablespace zl9indexcis
/
Create Index ·��ҽ������_IX_ִ�п���ID On ·��ҽ������(ִ�п���ID) Pctfree 5 Tablespace zl9indexcis
/


CREATE TABLE �ٴ�·��ҽ��(
		·����ĿID NUMBER(18),
    ҽ������ID NUMBER(18))
    TABLESPACE zl9BaseItem
    PCTFREE 5
    PCTUSED 85;
Alter Table �ٴ�·��ҽ�� Add Constraint �ٴ�·��ҽ��_PK Primary Key (·����ĿID,ҽ������ID) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table �ٴ�·��ҽ�� Add Constraint �ٴ�·��ҽ��_FK_·����ĿID Foreign Key (·����ĿID) References �ٴ�·����Ŀ(ID) On Delete Cascade;
Alter Table �ٴ�·��ҽ�� Add Constraint �ٴ�·��ҽ��_FK_ҽ������ID Foreign Key (ҽ������ID) References ·��ҽ������(ID) On Delete Cascade;


CREATE TABLE �ٴ�·������(
		��ĿID NUMBER(18),
    �ļ�ID NUMBER(18))
    TABLESPACE zl9BaseItem
    PCTFREE 5
    PCTUSED 85;
Alter Table �ٴ�·������ Add Constraint �ٴ�·������_PK Primary Key (��ĿID,�ļ�ID) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table �ٴ�·������ Add Constraint �ٴ�·������_FK_��ĿID Foreign Key (��ĿID) References �ٴ�·����Ŀ(ID) On Delete Cascade;
Alter Table �ٴ�·������ Add Constraint �ٴ�·������_FK_�ļ�ID Foreign Key (�ļ�ID) References �����ļ��б�(ID) On Delete Cascade;

Create Sequence �ٴ�·������_ID Start With 1;
CREATE TABLE �ٴ�·������(
		ID NUMBER(18),
    ·��ID NUMBER(18),
    �汾�� NUMBER(3),
		�׶�ID NUMBER(18),
		�������� NUMBER(1))
    TABLESPACE zl9BaseItem
    PCTFREE 5
    PCTUSED 85;
Alter Table �ٴ�·������ Add Constraint �ٴ�·������_PK Primary Key (ID) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table �ٴ�·������ Add Constraint �ٴ�·������_UQ_�������� Unique (·��ID,�汾��,�׶�ID,��������) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table �ٴ�·������ Add Constraint �ٴ�·������_FK_�汾�� Foreign Key (·��ID,�汾��) References �ٴ�·���汾(·��ID,�汾��) On Delete Cascade;
Alter Table �ٴ�·������ Add Constraint �ٴ�·������_FK_�׶�ID Foreign Key (�׶�ID) References �ٴ�·���׶�(ID) On Delete Cascade;


Create Sequence ·������ָ��_ID Start With 1;
CREATE TABLE ·������ָ��(
		ID NUMBER(18),
    ����ID NUMBER(18),
    ��� NUMBER(5),
		����ָ�� VARCHAR2(200),
		ָ������ NUMBER(1),
		ָ���� VARCHAR2(500))
    TABLESPACE zl9BaseItem
    PCTFREE 5
    PCTUSED 85;
Alter Table ·������ָ�� Add Constraint ·������ָ��_PK Primary Key (ID) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table ·������ָ�� Add Constraint ·������ָ��_UQ_��� Unique (����ID,���) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table ·������ָ�� Add Constraint ·������ָ��_UQ_����ָ�� Unique (����ID,����ָ��) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table ·������ָ�� Add Constraint ·������ָ��_FK_����ID Foreign Key (����ID) References �ٴ�·������(ID) On Delete Cascade;


CREATE TABLE ·����������(
		����ID NUMBER(18),
    ָ��ID NUMBER(18),
    ��ĿID NUMBER(18),
		��ϵʽ VARCHAR2(5),
		����ֵ VARCHAR2(50),
		������� NUMBER(1))
    TABLESPACE zl9BaseItem
    PCTFREE 5
    PCTUSED 85;
Alter Table ·���������� Add Constraint ·����������_UQ_���� Unique (ָ��ID,��ĿID,��ϵʽ,����ֵ) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table ·���������� Add Constraint ·����������_FK_����ID Foreign Key (����ID) References �ٴ�·������(ID) On Delete Cascade;
Alter Table ·���������� Add Constraint ·����������_FK_ָ��ID Foreign Key (ָ��ID) References ·������ָ��(ID) On Delete Cascade;
Alter Table ·���������� Add Constraint ·����������_FK_��ĿID Foreign Key (��ĿID) References �ٴ�·����Ŀ(ID) On Delete Cascade;
Create Index ·����������_IX_����ID On ·����������(����ID) Pctfree 5 Tablespace zl9indexcis
/


Create Sequence �����ٴ�·��_ID Start With 1;
CREATE TABLE �����ٴ�·��(
		ID NUMBER(18),
    ����ID NUMBER(18),
    ��ҳID NUMBER(5),
		����ID NUMBER(18),
		·��ID NUMBER(18),
		�汾�� NUMBER(3),
		������ VARCHAR2(20),
		����ʱ�� DATE,
		����˵�� VARCHAR2(1000),
		����ʱ�� DATE,
		״̬ NUMBER(1),
		��ǰ����   NUMBER(18),
		��ǰ�׶�ID NUMBER(18),
		ǰһ�׶�ID NUMBER(18))
    TABLESPACE zl9CISRec
    PCTFREE 5
    PCTUSED 85;
Alter Table �����ٴ�·�� Add Constraint �����ٴ�·��_PK Primary Key (ID) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table �����ٴ�·�� Add Constraint �����ٴ�·��_FK_��ҳID Foreign Key (����ID,��ҳID) References ������ҳ(����ID,��ҳID);
Alter Table �����ٴ�·�� Add Constraint �����ٴ�·��_FK_����ID Foreign Key (����ID) References ���ű�(ID);
Alter Table �����ٴ�·�� Add Constraint �����ٴ�·��_FK_�汾�� Foreign Key (·��ID,�汾��) References �ٴ�·���汾(·��ID,�汾��);
Create Index �����ٴ�·��_IX_����ID On �����ٴ�·��(����ID,��ҳID) Pctfree 5 Tablespace zl9indexcis
/
Create Index �����ٴ�·��_IX_����ID On �����ٴ�·��(����ID) Pctfree 5 Tablespace zl9indexcis
/
Create Index �����ٴ�·��_IX_·��ID On �����ٴ�·��(·��ID,�汾��) Pctfree 5 Tablespace zl9indexcis
/
Create Index �����ٴ�·��_IX_����ʱ�� On �����ٴ�·��(����ʱ��) Pctfree 5 Tablespace zl9indexcis
/


Create Sequence ����·��ִ��_ID Start With 1;
CREATE TABLE ����·��ִ��(
		ID NUMBER(18),
		·����¼ID NUMBER(18),
		�׶�ID NUMBER(18),		
		���� DATE,
		���� NUMBER(5),
		���� VARCHAR2(50),
		��ĿID NUMBER(18),
		��Ŀ��� NUMBER(5),
		��Ŀ���� VARCHAR2(1000),
		ִ���� NUMBER(1),
		��Ŀ��� VARCHAR2(500),
		���ԭ�� VARCHAR2(1000),
		ͼ��ID NUMBER(18),
		ִ���� VARCHAR2(20),
		ִ��ʱ�� DATE,
		ִ�н�� VARCHAR2(50),
		ִ��˵�� VARCHAR2(200),
		�Ǽ��� VARCHAR2(20),
		�Ǽ�ʱ�� DATE)
    TABLESPACE zl9CISRec
    PCTFREE 5
    PCTUSED 85;
Alter Table ����·��ִ�� Add Constraint ����·��ִ��_PK Primary Key (ID) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table ����·��ִ�� Add Constraint ����·��ִ��_UQ_��Ŀ���� Unique (·����¼ID,�׶�ID,����,��ĿID,��Ŀ����) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table ����·��ִ�� Add Constraint ����·��ִ��_FK_·����¼ID Foreign Key (·����¼ID) References �����ٴ�·��(ID);
Alter Table ����·��ִ�� Add Constraint ����·��ִ��_FK_�׶�ID Foreign Key (�׶�ID) References �ٴ�·���׶�(ID);
Alter Table ����·��ִ�� Add Constraint ����·��ִ��_FK_��ĿID Foreign Key (��ĿID) References �ٴ�·����Ŀ(ID);
Alter Table ����·��ִ�� Add Constraint ����·��ִ��_FK_ͼ��ID Foreign Key (ͼ��ID) References �ٴ�·��ͼ��(ID);
Create Index ����·��ִ��_IX_���� On ����·��ִ��(����) Pctfree 5 Tablespace zl9indexcis
/
Create Index ����·��ִ��_IX_·����¼ID On ����·��ִ��(·����¼ID) Pctfree 5 Tablespace zl9indexcis
/
Create Index ����·��ִ��_IX_�׶�ID On ����·��ִ��(�׶�ID) Pctfree 5 Tablespace zl9indexcis
/
Create Index ����·��ִ��_IX_��ĿID On ����·��ִ��(��ĿID) Pctfree 5 Tablespace zl9indexcis
/
Create Index ����·��ִ��_IX_ͼ��ID On ����·��ִ��(ͼ��ID) Pctfree 5 Tablespace zl9indexcis
/
Create Index ����·��ִ��_IX_�Ǽ�ʱ�� On ����·��ִ��(�Ǽ�ʱ��) Pctfree 5 Tablespace zl9indexcis
/


CREATE TABLE ����·������(
		·����¼ID NUMBER(18),
		�׶�ID NUMBER(18),
		���� DATE,
    ���� NUMBER(5),
		������ VARCHAR2(50),
		����ʱ�� DATE,
		������� NUMBER(2),
		����˵�� VARCHAR2(1000),
		�Ǽ��� VARCHAR2(20),
		�Ǽ�ʱ�� DATE)
    TABLESPACE zl9CISRec
    PCTFREE 5
    PCTUSED 85;
Alter Table ����·������ Add Constraint ����·������_PK Primary Key (·����¼ID,�׶�ID,����) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table ����·������ Add Constraint ����·������_FK_�׶�ID Foreign Key (�׶�ID) References �ٴ�·���׶�(ID);
Alter Table ����·������ Add Constraint ����·������_FK_·����¼ID Foreign Key (·����¼ID) References �����ٴ�·��(ID);
Create Index ����·������_IX_���� On ����·������(����) Pctfree 5 Tablespace zl9indexcis
/
Create Index ����·������_IX_�Ǽ�ʱ�� On ����·������(�Ǽ�ʱ��) Pctfree 5 Tablespace zl9indexcis
/


CREATE TABLE ����·��ָ��(
		·����¼ID NUMBER(18),
		�׶�ID NUMBER(18),
		���� DATE,
    ���� NUMBER(5),
		�������� NUMBER(1),
		����ָ�� VARCHAR2(50),
		ָ������ NUMBER(1),
		ָ���� VARCHAR2(50))
    TABLESPACE zl9CISRec
    PCTFREE 5
    PCTUSED 85;
Alter Table ����·��ָ�� Add Constraint ����·��ָ��_UQ_����ָ�� Unique (·����¼ID,�׶�ID,����,����ָ��) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table ����·��ָ�� Add Constraint ����·��ָ��_FK_�׶�ID Foreign Key (�׶�ID) References �ٴ�·���׶�(ID);
Alter Table ����·��ָ�� Add Constraint ����·��ָ��_FK_·����¼ID Foreign Key (·����¼ID) References �����ٴ�·��(ID);
Create Index ����·��ָ��_IX_���� On ����·��ָ��(����) Pctfree 5 Tablespace zl9indexcis
/

CREATE TABLE ����·��ҽ��(
		·��ִ��ID NUMBER(18),
    ����ҽ��ID NUMBER(18))
    TABLESPACE zl9CISRec
    PCTFREE 5
    PCTUSED 85;
Alter Table ����·��ҽ�� Add Constraint ����·��ҽ��_PK Primary Key (·��ִ��ID,����ҽ��ID) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table ����·��ҽ�� Add Constraint ����·��ҽ��_FK_·��ִ��ID Foreign Key (·��ִ��ID) References ����·��ִ��(ID);
Alter Table ����·��ҽ�� Add Constraint ����·��ҽ��_FK_����ҽ��ID Foreign Key (����ҽ��ID) References ����ҽ����¼(ID);


--��ԭ���Ӳ�����¼�ĸ���
Alter Table ���Ӳ�����¼ Add ·��ִ��ID Number(18);
Alter Table ���Ӳ�����¼ Add Constraint ���Ӳ�����¼_FK_·��ִ��ID Foreign Key (·��ִ��ID) References ����·��ִ��(ID);
Create Index ���Ӳ�����¼_IX_·��ִ��ID On ���Ӳ�����¼(·��ִ��ID) Pctfree 5 Tablespace zl9indexcis
/

--------------------------------------------------------------------------------------------------------------------------
--�ٴ�·���������ݲ���
--------------------------------------------------------------------------------------------------------------------------
Insert Into zlBaseCode(ϵͳ,����,�̶�,˵��,����) Values(100,'�ٴ���������',1,'�ٴ�·���������Ӧ����ĸ��ӡ������̶ȣ��Լ�ʵʩ·�������׳̶Ƚ��е�һ�����ֱ�׼��','ҽ�ƹ���');
Insert Into zlBaseCode(ϵͳ,����,�̶�,˵��,����) Values(100,'·���������',0,'�ٴ�·����Ŀִ��ʱ�ĳ������','ҽ�ƹ���');
Insert Into zlBaseCode(ϵͳ,����,�̶�,˵��,����) Values(100,'���쳣��ԭ��',0,'�ٴ�·�������ı���ԭ��','ҽ�ƹ���');

Insert Into �ٴ���������(����,����,����)
	Select 'A','������ͨ��','DCPTX' From Dual Union ALL
	Select 'B','������֢��','DCJZX' From Dual Union ALL
	Select 'C','����������','FZYNX' From Dual Union ALL
	Select 'D','����Σ����','FZWZX' From Dual;

Insert Into zlStreamTabs(System_NO,Table_Name,Dml_Handle,Repeat_Way,Fixation)
Select 100,'�ٴ���������',0,2,1 From Dual Union All
Select 100,'·���������',0,2,1 From Dual Union All
Select 100,'���쳣��ԭ��',0,2,1 From Dual Union All
Select 100,'�ٴ�·��ͼ��',0,2,1 From Dual Union All
Select 100,'�ٴ�·��Ŀ¼',0,2,1 From Dual Union All
Select 100,'�ٴ�·���ļ�',0,2,1 From Dual Union All
Select 100,'�ٴ�·������',0,2,1 From Dual Union All
Select 100,'�ٴ�·������',0,2,1 From Dual Union All
Select 100,'�ٴ�·���汾',0,2,1 From Dual Union All
Select 100,'�ٴ�·���׶�',0,2,1 From Dual Union All
Select 100,'�ٴ�·������',0,2,1 From Dual Union All
Select 100,'�ٴ�·����Ŀ',0,2,1 From Dual Union All
Select 100,'·��ҽ������',0,2,1 From Dual Union All
Select 100,'�ٴ�·��ҽ��',0,2,1 From Dual Union All
Select 100,'�ٴ�·������',0,2,1 From Dual Union All
Select 100,'�ٴ�·������',0,2,1 From Dual Union All
Select 100,'·������ָ��',0,2,1 From Dual Union All
Select 100,'·����������',0,2,1 From Dual;

Insert Into zlStreamTabs(System_NO,Table_Name,Dml_Handle,Repeat_Way,Fixation)
Select 100,'�����ٴ�·��',0,3,1 From Dual Union All
Select 100,'����·��ִ��',0,3,1 From Dual Union All
Select 100,'����·������',0,3,1 From Dual Union ALL
Select 100,'����·��ָ��',0,3,1 From Dual Union ALL
Select 100,'����·��ҽ��',0,3,1 From Dual;

Insert Into zlBakTables(ϵͳ,����)
Select 100,'�����ٴ�·��' From Dual Union ALL
Select 100,'����·��ִ��' From Dual Union ALL
Select 100,'����·������' From Dual Union ALL
Select 100,'����·��ָ��' From Dual Union ALL
Select 100,'����·��ҽ��' From Dual;

Insert Into zlPrograms(���,����,˵��,ϵͳ,����) Values(1078,'�ٴ�·������','���ٴ�·���Ļ�����Ϣ��·������Ϣ�����汾�仯���ж��塢ά��',100,'zl9CISJob');
Insert Into zlPrograms(���,����,˵��,ϵͳ,����) Values(1256,'�ٴ�·��Ӧ��','�����ٴ�·���ĵ��룬���ɣ�ִ�У������ȹ���Ӧ��',100,'zl9CISJob');
Insert Into zlPrograms(���,����,˵��,ϵͳ,����) Values(1275,'�ٴ�·������','�Ը����ٴ�·����Ӧ���������ϸ���в��ģ�����',100,'zl9CISJob');

Insert Into zlProgFuncs(ϵͳ,���,����) Values(100,1078,'����');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1078,'��ɾ��',1,'���ٴ�·���������ӡ��޸ġ�ɾ����Ӧ�÷�Χ�Ȼ�����Ϣά����Ȩ��');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1078,'����XML',2,'��XML�ļ������ٴ�·����Ȩ��');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1078,'����XML',3,'���ٴ�·��������XML�ļ���Ȩ��');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1078,'���',4,'���ƶ��õ��ٴ�·�������������Ч��Ȩ�ޣ����и�Ȩ��ͬʱ����ȡ�����');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1078,'ͣ��',5,'���Ѿ����Ӧ�õ��ٴ�·������ͣ�õ�Ȩ�ޣ����и�Ȩ��ͬʱ���������Ѿ�ͣ�õ�·��');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1078,'·�������',6,'���ٴ�·����ķ��࣬ʱ��׶Σ���Ŀ���汾����Ϣ������ƶ����Ȩ��');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1078,'���������',7,'���ٴ�·���ĵ������������׶�������������������������ƶ����Ȩ��');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1078,'ȫԺ·��',8,'��ȫԺ���ٴ�·�����й����Ȩ�ޣ������и�Ȩ��ֻ�ܹ����Ƶ��ٴ�·��');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1078,'ͼ������',9,'���ٴ�·��������Ŀ���Զ�Ӧ��ͼ�����������ɾ��Ȩ��');

Insert Into zlProgFuncs(ϵͳ,���,����) Values(100,1256,'����');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1256,'����·��',1,'������Ժ������Ʋ��˽������������������ٴ�·����Ȩ��');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1256,'����·��',2,'�Բ��˽���·����Ŀ���ɵ�Ȩ�ޣ�����߱�ҽ���´�Ͳ������ɵ�Ȩ�޲������ɶ�Ӧ��Ŀ');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1256,'ִ��·��',3,'�Բ����ٴ�·�������ݽ���ִ�е�Ȩ�ޣ�����߱�ҽ��ֹͣȨ�޲��ܽ�������ִ��');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1256,'�׶�����',4,'�Բ����ٴ�·����ÿ���׶ν���������Ȩ��');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1256,'����·��',5,'�Զ�����Ϊ����ٴ�·��ִ�е�Ȩ��');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1256,'·������Ŀ',6,'���ٴ�·����ƶ���õķ�Χ֮�����·����Ŀ��Ȩ��');

Insert Into zlProgFuncs(ϵͳ,���,����) Values(100,1275,'����');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1275,'ȫԺ·��',1,'��ȫԺ���ٴ�·�����и��ٵ�Ȩ�ޣ������и�Ȩ��ֻ�ܸ��ٱ��Ƶ��ٴ�·��');

--- �ٴ�·�������1��
Insert Into zlProgRelas(ϵͳ,���,����,���,��ϵ,����,�����ϵ) Values(100,1078,'·�������',1,2,1,0);
Insert Into zlProgRelas(ϵͳ,���,����,���,��ϵ,����,�����ϵ) Values(100,1078,'���������',1,2,0,0);

--1078:�ٴ�·������(����)
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'���ű�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'��������˵��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'��Ա��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'��Ա����˵��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'������Ա','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'�ϻ���Ա��','SELECT');

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'����','SELECT');

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'�ٴ���������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'·���������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'���쳣��ԭ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'�ٴ�·��ͼ��','SELECT');

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'�ٴ�·��Ŀ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'�ٴ�·������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'�ٴ�·������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'�ٴ�·���ļ�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'�ٴ�·���汾','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'�ٴ�·������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'�ٴ�·���׶�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'�ٴ�·����Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'�ٴ�·������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'·������ָ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'·����������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'�ٴ�·��ҽ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'·��ҽ������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'�ٴ�·������','SELECT');

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'�����ļ��б�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'��������Ŀ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'�������Ŀ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'������ϱ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'������ĿĿ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'�շ���ĿĿ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'ҩƷ���','SELECT');

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'�����÷�����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'����Ƶ����Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'������Ŀ����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'���Ʒ���Ŀ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'����ִ�п���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'���Ƹ�����Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'������Ŀ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'���Ƽ�鲿λ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'���Ƽ���걾','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'�շ���Ŀ����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'�շ�ִ�п���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'��������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'ҩƷ����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'��������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'ҽ�����ݶ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'��ҩ�����ע','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'������Ŀ�ο�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'���鱨����Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'������Ϣ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'������ҳ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'���˹Һż�¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'����ҽ����¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'������������¼','SELECT');

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����',user,'Zl_Lob_Read','EXECUTE');

--1078:�ٴ�·������(��ɾ��)
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'��ɾ��',user,'Zl_Lob_Append','EXECUTE');

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'��ɾ��',user,'Zl_�ٴ�·��Ŀ¼_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'��ɾ��',user,'Zl_�ٴ�·��Ŀ¼_Update','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'��ɾ��',user,'Zl_�ٴ�·��Ŀ¼_Delete','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'��ɾ��',user,'Zl_�ٴ�·���ļ�_Delete','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'��ɾ��',user,'Zl_�ٴ�·���ļ�_Insert','EXECUTE');

--1078:�ٴ�·������(����XML)
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����XML',user,'zl_�ٴ�·��Ŀ¼_Update','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����XML',user,'zl_�ٴ�·��Ŀ¼_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����XML',user,'Zl_�ٴ�·���汾_Delete','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����XML',user,'Zl_�ٴ�·���汾_Update','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����XML',user,'Zl_·������ָ��_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����XML',user,'Zl_·����������_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����XML',user,'Zl_·��ҽ������_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����XML',user,'Zl_�ٴ�·������_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����XML',user,'Zl_�ٴ�·���׶�_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'����XML',user,'Zl_�ٴ�·����Ŀ_Insert','EXECUTE');

--1078:�ٴ�·������(���)
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'���',user,'Zl_�ٴ�·���汾_Audit','EXECUTE');

--1078:�ٴ�·������(ͣ��)
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'ͣ��',user,'Zl_�ٴ�·���汾_Stop','EXECUTE');

--1078:�ٴ�·������(·�������)
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'·�������',user,'Zl_�ٴ�·���汾_Delete','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'·�������',user,'Zl_�ٴ�·���汾_Copy','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'·�������',user,'Zl_�ٴ�·���汾_Update','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'·�������',user,'Zl_·������ָ��_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'·�������',user,'Zl_·����������_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'·�������',user,'Zl_�ٴ�·���׶�_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'·�������',user,'Zl_�ٴ�·���׶�_Delete','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'·�������',user,'Zl_�ٴ�·���׶�_Update','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'·�������',user,'Zl_�ٴ�·������_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'·�������',user,'Zl_·��ҽ������_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'·�������',user,'Zl_�ٴ�·����Ŀ_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'·�������',user,'Zl_�ٴ�·����Ŀ_Delete','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'·�������',user,'Zl_�ٴ�·����Ŀ_Update','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'·�������',user,'Zl_GetPathCharge','EXECUTE');

--1078:�ٴ�·������(ͼ������)
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'ͼ������',user,'Zl_Lob_Append','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'ͼ������',user,'Zl_�ٴ�·��ͼ��_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1078,'ͼ������',user,'Zl_�ٴ�·��ͼ��_Delete','EXECUTE');

--1275:�ٴ�·������(����)
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'���ű�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'��������˵��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'��Ա��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'��Ա����˵��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'������Ա','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'�ϻ���Ա��','SELECT');

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'�ٴ�·��ͼ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'�ٴ�·��Ŀ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'�ٴ�·������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'�ٴ�·������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'�ٴ�·���ļ�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'�ٴ�·���汾','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'�ٴ�·������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'�ٴ�·���׶�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'�ٴ�·����Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'�ٴ�·������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'·������ָ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'·����������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'�ٴ�·��ҽ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'·��ҽ������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'�ٴ�·������','SELECT');

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'�����ļ��б�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'������ĿĿ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'�շ���ĿĿ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'ҩƷ���','SELECT');

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'������Ϣ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'������ҳ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'�����ٴ�·��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'����·��ִ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'����·������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'����·��ָ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'����·��ҽ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'����ҽ����¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'����ҽ��״̬','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'����ҽ������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'����ҽ������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'���Ӳ�����¼','SELECT');

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1275,'����',user,'Zl_Lob_Read','EXECUTE');


--1256:�ٴ�·��Ӧ��(����)
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����',user,'�ٴ�·��Ŀ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����',user,'�ٴ�·���汾','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����',user,'�ٴ�·���׶�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����',user,'�ٴ�·������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����',user,'�ٴ�·����Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����',user,'����·��ִ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����',user,'����·������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����',user,'�ٴ�·��ҽ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����',user,'�ٴ�·������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����',user,'·��ҽ������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����',user,'�����ļ��б�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����',user,'�ٴ�·��ͼ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����',user,'�ٴ�·������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����',user,'·������ָ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����',user,'·����������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����',user,'�����ٴ�·��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����',user,'����·��ָ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����',user,'������ҳ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����',user,'������Ϣ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����',user,'���˱䶯��¼','SELECT');


--1256:�ٴ�·��Ӧ��(����·��)
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����·��',user,'�ٴ�·������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����·��',user,'�ٴ�·������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����·��',user,'������ϼ�¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����·��',user,'Zl_����·������_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����·��',user,'Zl_����·������_Delete','EXECUTE');

--1256:�ٴ�·��Ӧ��(����·��)
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����·��',user,'����·��ҽ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����·��',user,'���Ӳ�����¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����·��',user,'������������¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����·��',user,'Zl_����·������_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����·��',user,'Zl_����·������_Delete','EXECUTE');

--1256:�ٴ�·��Ӧ��(·������Ŀ)
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'·������Ŀ',user,'·���������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'·������Ŀ',user,'Zl_����·������_Insert','EXECUTE');

--1256:�ٴ�·��Ӧ��(�׶�����)
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'�׶�����',user,'���쳣��ԭ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'�׶�����',user,'����ģ�����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'�׶�����',user,'�������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'�׶�����',user,'Zl_����·������_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'�׶�����',user,'Zl_����·������_Delete','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'�׶�����',user,'Zl_Getpathcharge','EXECUTE');

--1256:�ٴ�·��Ӧ��(ִ��·��)
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'ִ��·��',user,'����·��ҽ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'ִ��·��',user,'Zl_����·��ִ��_Update','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'ִ��·��',user,'Zl_����·��ִ��_Delete','EXECUTE');

--1256:�ٴ�·��Ӧ��(����·��)
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����·��',user,'Zl_����·������_Update','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1256,'����·��',user,'Zl_����·������_Delete','EXECUTE');

--Insert Into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Values('ȱʡ',zlMenus_id.nextval,   zlMenus_id.nextval-20,'ҽ��������Ŀ','ҽ������','D',99,'����������������ƴ�ʩӦ�õ���ػ�����',100,NULL);
--Insert Into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Values('ȱʡ',zlMenus_id.nextval,      zlMenus_id.nextval-10,'�ٴ�·������','·������','J',99,'���ٴ�·���Ļ�����Ϣ��·������Ϣ�����汾�仯���ж��塢ά��',100,1078);

--Insert Into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Values('ȱʡ',zlMenus_id.nextval,null,'�������ϼ���','��������','E',99,'��������벡��������ѯ',100,NULL);
--Insert Into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Values('ȱʡ',zlMenus_id.nextval,   zlMenus_id.nextval-5,'�ٴ�·������','·������','G',129,'�Ը����ٴ�·����Ӧ���������ϸ���в��ģ�����',100,1275);

Insert Into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) 
	Select 'ȱʡ',zlMenus_id.nextval,ID,'�ٴ�·������','·������','J',99,'���ٴ�·���Ļ�����Ϣ��·������Ϣ�����汾�仯���ж��塢ά��',100,1078 
	From zlMenus Where ���='ȱʡ' And ����='ҽ��������Ŀ' And ϵͳ=100 And ģ�� Is Null;

Insert Into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) 
	Select 'ȱʡ',zlMenus_id.nextval,ID,'�ٴ�·������','·������','G',129,'�Ը����ٴ�·����Ӧ���������ϸ���в��ģ�����',100,1275
	From zlMenus Where ���='ȱʡ' And ����='�������ϼ���' And ϵͳ=100 And ģ�� Is Null;

--------------------------------------------------------------------------------------------------------------------------
--�ٴ�·���洢���̲���
--------------------------------------------------------------------------------------------------------------------------
Create Or Replace Procedure Zl_�ٴ�·��Ŀ¼_Insert
(
  ����_In     �ٴ�·��Ŀ¼.����%Type,
  ����_In     �ٴ�·��Ŀ¼.����%Type,
  ����_In     �ٴ�·��Ŀ¼.����%Type,
  ˵��_In     �ٴ�·��Ŀ¼.˵��%Type,
  ��������_In �ٴ�·��Ŀ¼.��������%Type,
  ���ò���_In �ٴ�·��Ŀ¼.���ò���%Type,
  �����Ա�_In �ٴ�·��Ŀ¼.�����Ա�%Type,
  ��������_In �ٴ�·��Ŀ¼.��������%Type,
  ͨ��_In     �ٴ�·��Ŀ¼.ͨ��%Type,
  ����ids_In  Varchar2 := Null,
  ����ids_In  Varchar2 := Null,
  ·��id_In   �ٴ�·��Ŀ¼.Id%Type := Null
  --������
  --����IDs_IN����Ϊָ������Ӧ��ʱ���룬��ʽΪ"����ID1,����ID2,..."
  --����IDs_IN�������ʽΪ"����ID1,����ID2,...;���ID1,���ID2,..."
  --·��id_In���Ƿ����ⲿȷ���µ�ID
) Is
  v_·��id �ٴ�·��Ŀ¼.Id%Type;

  v_������ Varchar2(4000);
  v_��ǰid Number(18);

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  If ·��id_In Is Not Null Then
    v_·��id := ·��id_In;
  Else
    Select �ٴ�·��Ŀ¼_Id.Nextval Into v_·��id From Dual;
  End If;

  Insert Into �ٴ�·��Ŀ¼
    (ID, ����, ����, ����, ˵��, ��������, ���ò���, �����Ա�, ��������, ͨ��)
  Values
    (v_·��id, ����_In, ����_In, ����_In, ˵��_In, ��������_In, ���ò���_In, �����Ա�_In, ��������_In, ͨ��_In);

  If ͨ��_In = 2 And ����ids_In Is Not Null Then
    v_������ := ����ids_In || ',';
    While v_������ Is Not Null Loop
      v_��ǰid := To_Number(Substr(v_������, 1, Instr(v_������, ',') - 1));
      v_������ := Substr(v_������, Instr(v_������, ',') + 1);
    
      Insert Into �ٴ�·������ (·��id, ����id) Values (v_·��id, v_��ǰid);
    End Loop;
  End If;

  If ����ids_In Is Not Null Then
    v_������ := Substr(����ids_In, 1, Instr(����ids_In, ';') - 1);
    If v_������ Is Not Null Then
      v_������ := v_������ || ',';
      While v_������ Is Not Null Loop
        v_��ǰid := To_Number(Substr(v_������, 1, Instr(v_������, ',') - 1));
        v_������ := Substr(v_������, Instr(v_������, ',') + 1);
      
        Insert Into �ٴ�·������ (·��id, ����id) Values (v_·��id, v_��ǰid);
      End Loop;
    End If;
  
    v_������ := Substr(����ids_In, Instr(����ids_In, ';') + 1);
    If v_������ Is Not Null Then
      v_������ := v_������ || ',';
      While v_������ Is Not Null Loop
        v_��ǰid := To_Number(Substr(v_������, 1, Instr(v_������, ',') - 1));
        v_������ := Substr(v_������, Instr(v_������, ',') + 1);
      
        Insert Into �ٴ�·������ (·��id, ���id) Values (v_·��id, v_��ǰid);
      End Loop;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ�·��Ŀ¼_Insert;
/

Create Or Replace Procedure Zl_�ٴ�·��Ŀ¼_Update
(
  ·��id_In   �ٴ�·��Ŀ¼.Id%Type,
  ����_In     �ٴ�·��Ŀ¼.����%Type,
  ����_In     �ٴ�·��Ŀ¼.����%Type,
  ����_In     �ٴ�·��Ŀ¼.����%Type,
  ˵��_In     �ٴ�·��Ŀ¼.˵��%Type,
  ��������_In �ٴ�·��Ŀ¼.��������%Type,
  ���ò���_In �ٴ�·��Ŀ¼.���ò���%Type,
  �����Ա�_In �ٴ�·��Ŀ¼.�����Ա�%Type,
  ��������_In �ٴ�·��Ŀ¼.��������%Type,
  ͨ��_In     �ٴ�·��Ŀ¼.ͨ��%Type,
  ����ids_In  Varchar2 := Null,
  ����ids_In  Varchar2 := Null
  --������
  --����IDs_IN����Ϊָ������Ӧ��ʱ���룬��ʽΪ"����ID1,����ID2,..."
  --����IDs_IN�������ʽΪ"����ID1,����ID2,...;���ID1,���ID2,..."
) Is
  v_������ Varchar2(4000);
  v_��ǰid Number(18);

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  Update �ٴ�·��Ŀ¼
  Set ���� = ����_In, ���� = ����_In, ���� = ����_In, ˵�� = ˵��_In, �������� = ��������_In, ���ò��� = ���ò���_In, �����Ա� = �����Ա�_In, �������� = ��������_In,
      ͨ�� = ͨ��_In
  Where ID = ·��id_In;

  Delete From �ٴ�·������ Where ·��id = ·��id_In;
  If ͨ��_In = 2 And ����ids_In Is Not Null Then
    v_������ := ����ids_In || ',';
    While v_������ Is Not Null Loop
      v_��ǰid := To_Number(Substr(v_������, 1, Instr(v_������, ',') - 1));
      v_������ := Substr(v_������, Instr(v_������, ',') + 1);
    
      Insert Into �ٴ�·������ (·��id, ����id) Values (·��id_In, v_��ǰid);
    End Loop;
  End If;

  Delete From �ٴ�·������ Where ·��id = ·��id_In;
  If ����ids_In Is Not Null Then
    v_������ := Substr(����ids_In, 1, Instr(����ids_In, ';') - 1);
    If v_������ Is Not Null Then
      v_������ := v_������ || ',';
      While v_������ Is Not Null Loop
        v_��ǰid := To_Number(Substr(v_������, 1, Instr(v_������, ',') - 1));
        v_������ := Substr(v_������, Instr(v_������, ',') + 1);
      
        Insert Into �ٴ�·������ (·��id, ����id) Values (·��id_In, v_��ǰid);
      End Loop;
    End If;
  
    v_������ := Substr(����ids_In, Instr(����ids_In, ';') + 1);
    If v_������ Is Not Null Then
      v_������ := v_������ || ',';
      While v_������ Is Not Null Loop
        v_��ǰid := To_Number(Substr(v_������, 1, Instr(v_������, ',') - 1));
        v_������ := Substr(v_������, Instr(v_������, ',') + 1);
      
        Insert Into �ٴ�·������ (·��id, ���id) Values (·��id_In, v_��ǰid);
      End Loop;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ�·��Ŀ¼_Update;
/

Create Or Replace Procedure Zl_�ٴ�·��Ŀ¼_Delete(·��id_In �ٴ�·��Ŀ¼.Id%Type) Is
  v_Count Number;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --�����������˵İ汾��������ɾ��
  Select Count(*) Into v_Count From �ٴ�·���汾 Where ·��id = ·��id_In And ���ʱ�� Is Not Null;
  If Nvl(v_Count, 0) > 0 Then
    v_Error := '���ٴ�·�������Ѿ���˵�·����汾��������ɾ����';
    Raise Err_Custom;
  End If;

  Delete From �ٴ�·��Ŀ¼ Where ID = ·��id_In;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ�·��Ŀ¼_Delete;
/

Create Or Replace Procedure Zl_�ٴ�·���ļ�_Delete
(
  ·��id_In �ٴ�·���ļ�.·��id%Type,
  �ļ���_In �ٴ�·���ļ�.�ļ���%Type
) Is
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  Delete From �ٴ�·���ļ� Where ·��id = ·��id_In And �ļ��� = �ļ���_In;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ�·���ļ�_Delete;
/

Create Or Replace Procedure Zl_�ٴ�·���ļ�_Insert
(
  ·��id_In �ٴ�·���ļ�.·��id%Type,
  �ļ���_In �ٴ�·���ļ�.�ļ���%Type
) Is
  v_Temp     Varchar2(255);
  v_��Ա���� ����ҽ��״̬.������Ա%Type;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  v_Temp     := Zl_Identity;
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);

  Insert Into �ٴ�·���ļ� (·��id, �ļ���, ������, ����ʱ��) Values (·��id_In, �ļ���_In, v_��Ա����, Sysdate);
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ�·���ļ�_Insert;
/

Create Or Replace Procedure Zl_�ٴ�·��ͼ��_Insert(ͼ��id_In �ٴ�·��ͼ��.Id%Type) Is
Begin
  Insert Into �ٴ�·��ͼ�� (ID, ����) Values (ͼ��id_In, 0);
End Zl_�ٴ�·��ͼ��_Insert;
/

Create Or Replace Procedure Zl_�ٴ�·��ͼ��_Delete(ͼ��id_In �ٴ�·��ͼ��.Id%Type) Is
Begin
  Delete From �ٴ�·��ͼ�� Where ID = ͼ��id_In;
End Zl_�ٴ�·��ͼ��_Delete;
/

Create Or Replace Function Zl_Lob_Read
(
  Tab_In   In Number,
  Key_In   In Varchar2,
  Pos_In   In Number,
  Moved_In In Number := 0
  --����˵���� 
  --Tab_In������LOB�����ݱ�
  --        0-�������ͼ��;1-�����ļ���ʽ;2-�����ļ�ͼ��;3-�������ĸ�ʽ;4-��������ͼ��; 
  --        5-���Ӳ�����ʽ;6-���Ӳ���ͼ��;7-����ҳ���ʽ��8-���Ӳ�������;9-�����ص���� 
  --        10-�ٴ�·���ļ�,11-�ٴ�·��ͼ��
  --Key_In�����ݼ�¼�Ĺؼ��� 
  --Pos_In����0��ʼ���϶�ȡ��ֱ������Ϊ��
  --Moved_In: 0������¼,1��ȡת���󱸱��¼ 
) Return Varchar2 Is
  l_Blob   Blob;
  v_Buffer Varchar2(32767);
  n_Amount Number := 2000;
  n_Offset Number := 1;
Begin
  If Tab_In = 0 Then
    Select ͼ�� Into l_Blob From �������ͼ�� Where ���� = Key_In;
  Elsif Tab_In = 1 Then
    Select ���� Into l_Blob From �����ļ���ʽ Where �ļ�id = To_Number(Key_In);
  Elsif Tab_In = 2 Then
    Select ͼ�� Into l_Blob From �����ļ�ͼ�� Where ����id = To_Number(Key_In);
  Elsif Tab_In = 3 Then
    Select ���� Into l_Blob From �������ĸ�ʽ Where �ļ�id = To_Number(Key_In);
  Elsif Tab_In = 4 Then
    Select ͼ�� Into l_Blob From ��������ͼ�� Where ����id = To_Number(Key_In);
  Elsif Tab_In = 5 Then
    If Moved_In = 0 Then
      Select ���� Into l_Blob From ���Ӳ�����ʽ Where �ļ�id = To_Number(Key_In);
    Else
      Select ���� Into l_Blob From H���Ӳ�����ʽ Where �ļ�id = To_Number(Key_In);
    End If;
  Elsif Tab_In = 6 Then
    If Moved_In = 0 Then
      Select ͼ�� Into l_Blob From ���Ӳ���ͼ�� Where ����id = To_Number(Key_In);
    Else
      Select ͼ�� Into l_Blob From H���Ӳ���ͼ�� Where ����id = To_Number(Key_In);
    End If;
  Elsif Tab_In = 7 Then
    Select ͼ��
    Into l_Blob
    From ����ҳ���ʽ
    Where ���� = To_Number(Substr(Key_In, 1, 1)) And ��� = Substr(Key_In, 3);
  Elsif Tab_In = 8 Then
    If Moved_In = 0 Then
      Select ����
      Into l_Blob
      From ���Ӳ�������
      Where ����id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And ��� = Substr(Key_In, Instr(Key_In, ',') + 1);
    Else
      Select ����
      Into l_Blob
      From H���Ӳ�������
      Where ����id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And ��� = Substr(Key_In, Instr(Key_In, ',') + 1);
    End If;
  Elsif Tab_In = 9 Then
    Select ���ͼ�� Into l_Blob From �����ص���� Where ��� = To_Number(Key_In);
  Elsif Tab_In = 10 Then
    Select ����
    Into l_Blob
    From �ٴ�·���ļ�
    Where ·��id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And �ļ��� = Substr(Key_In, Instr(Key_In, ',') + 1);
  Elsif Tab_In = 11 Then
    Select ͼ�� Into l_Blob From �ٴ�·��ͼ�� Where ID = To_Number(Key_In);
  End If;

  n_Offset := n_Offset + Pos_In * n_Amount;
  Dbms_Lob.Read(l_Blob, n_Amount, n_Offset, v_Buffer);
  Return v_Buffer;
Exception
  When No_Data_Found Then
    Return Null;
End Zl_Lob_Read;
/

Create Or Replace Procedure Zl_Lob_Append
(
  Tab_In In Number,
  Key_In In Varchar2,
  Txt_In In Varchar2, --16���Ƶ��ļ�Ƭ�λ�����Ƭ��
  Cls_In In Number := 0 --�Ƿ����ԭ�������ݣ���һƬ�δ���ʱΪ1���Ժ�Ϊ0
  --����˵����
  --Tab_In������LOB�����ݱ�
  --        0-�������ͼ��;1-�����ļ���ʽ;2-�����ļ�ͼ��;3-�������ĸ�ʽ;4-��������ͼ��;
  --        5-���Ӳ�����ʽ;6-���Ӳ���ͼ��;7-����ҳ���ʽ��8-���Ӳ�������;9-�����ص����
  --        10-�ٴ�·���ļ�,11-�ٴ�·��ͼ��
  --Key_In�����ݼ�¼�Ĺؼ���
  --Txt_In��16���Ƶ��ļ�Ƭ�λ�����Ƭ��
  --Cls_In���Ƿ����ԭ�������ݣ���һƬ�δ���ʱΪ1���Ժ�Ϊ0
) Is
  l_Blob Blob;
Begin
  If Tab_In = 0 Then
    If Cls_In = 1 Then
      Update �������ͼ�� Set ͼ�� = Empty_Blob() Where ���� = Key_In;
    End If;
    Select ͼ�� Into l_Blob From �������ͼ�� Where ���� = Key_In For Update;
  Elsif Tab_In = 1 Then
    If Cls_In = 1 Then
      Update �����ļ���ʽ Set ���� = Empty_Blob() Where �ļ�id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into �����ļ���ʽ (�ļ�id, ����) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select ���� Into l_Blob From �����ļ���ʽ Where �ļ�id = To_Number(Key_In) For Update;
  Elsif Tab_In = 2 Then
    If Cls_In = 1 Then
      Update �����ļ�ͼ�� Set ͼ�� = Empty_Blob() Where ����id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into �����ļ�ͼ�� (����id, ͼ��) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select ͼ�� Into l_Blob From �����ļ�ͼ�� Where ����id = To_Number(Key_In) For Update;
  Elsif Tab_In = 3 Then
    If Cls_In = 1 Then
      Update �������ĸ�ʽ Set ���� = Empty_Blob() Where �ļ�id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into �������ĸ�ʽ (�ļ�id, ����) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select ���� Into l_Blob From �������ĸ�ʽ Where �ļ�id = To_Number(Key_In) For Update;
  Elsif Tab_In = 4 Then
    If Cls_In = 1 Then
      Update ��������ͼ�� Set ͼ�� = Empty_Blob() Where ����id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into ��������ͼ�� (����id, ͼ��) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select ͼ�� Into l_Blob From ��������ͼ�� Where ����id = To_Number(Key_In) For Update;
  Elsif Tab_In = 5 Then
    If Cls_In = 1 Then
      Update ���Ӳ�����ʽ Set ���� = Empty_Blob() Where �ļ�id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into ���Ӳ�����ʽ (�ļ�id, ����) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select ���� Into l_Blob From ���Ӳ�����ʽ Where �ļ�id = To_Number(Key_In) For Update;
  Elsif Tab_In = 6 Then
    If Cls_In = 1 Then
      Update ���Ӳ���ͼ�� Set ͼ�� = Empty_Blob() Where ����id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into ���Ӳ���ͼ�� (����id, ͼ��) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select ͼ�� Into l_Blob From ���Ӳ���ͼ�� Where ����id = To_Number(Key_In) For Update;
  Elsif Tab_In = 7 Then
    If Cls_In = 1 Then
      Update ����ҳ���ʽ
      Set ͼ�� = Empty_Blob()
      Where ���� = To_Number(Substr(Key_In, 1, 1)) And ��� = Substr(Key_In, 3);
    End If;
    Select ͼ��
    Into l_Blob
    From ����ҳ���ʽ
    Where ���� = To_Number(Substr(Key_In, 1, 1)) And ��� = Substr(Key_In, 3)
    For Update;
  Elsif Tab_In = 8 Then
    If Cls_In = 1 Then
      Update ���Ӳ�������
      Set ���� = Empty_Blob()
      Where ����id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And ��� = Substr(Key_In, Instr(Key_In, ',') + 1);
    End If;
    Select ����
    Into l_Blob
    From ���Ӳ�������
    Where ����id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And ��� = Substr(Key_In, Instr(Key_In, ',') + 1)
    For Update;
  Elsif Tab_In = 9 Then
    If Cls_In = 1 Then
      Update �����ص���� Set ���ͼ�� = Empty_Blob() Where ��� = To_Number(Key_In);
    End If;
    Select ���ͼ�� Into l_Blob From �����ص���� Where ��� = To_Number(Key_In) For Update;
  Elsif Tab_In = 10 Then
    If Cls_In = 1 Then
      Update �ٴ�·���ļ�
      Set ���� = Empty_Blob()
      Where ·��id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And
            �ļ��� = Substr(Key_In, Instr(Key_In, ',') + 1);
    End If;
    Select ����
    Into l_Blob
    From �ٴ�·���ļ�
    Where ·��id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And �ļ��� = Substr(Key_In, Instr(Key_In, ',') + 1)
    For Update;
  Elsif Tab_In = 11 Then
    If Cls_In = 1 Then
      Update �ٴ�·��ͼ�� Set ͼ�� = Empty_Blob() Where ID = To_Number(Key_In);
    End If;
    Select ͼ�� Into l_Blob From �ٴ�·��ͼ�� Where ID = To_Number(Key_In) For Update;
  End If;
  Dbms_Lob.Writeappend(l_Blob, Length(Hextoraw(Txt_In)) / 2, Hextoraw(Txt_In));
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Lob_Append;
/

Create Or Replace Procedure Zl_�ٴ�·���汾_Update
(
  ·��id_In     �ٴ�·���汾.·��id%Type,
  �汾��_In     �ٴ�·���汾.�汾��%Type,
  ��׼סԺ��_In �ٴ�·���汾.��׼סԺ��%Type,
  ��׼����_In   �ٴ�·���汾.��׼����%Type,
  �汾˵��_In   �ٴ�·���汾.�汾˵��%Type
) Is
Begin
  Update �ٴ�·���汾
  Set ��׼סԺ�� = ��׼סԺ��_In, ��׼���� = ��׼����_In, �汾˵�� = �汾˵��_In
  Where ·��id = ·��id_In And �汾�� = �汾��_In;
  If Sql%RowCount = 0 Then
    Insert Into �ٴ�·���汾
      (·��id, �汾��, ��׼סԺ��, ��׼����, �汾˵��, ������, ����ʱ��)
    Values
      (·��id_In, �汾��_In, ��׼סԺ��_In, ��׼����_In, �汾˵��_In, zl_UserName, Sysdate);
  Else
    --ɾ����Ӧ�ĵ���������Ϣ,�������±���
    Delete From �ٴ�·������ Where ·��id = ·��id_In And �汾�� = �汾��_In And �������� = 1;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ�·���汾_Update;
/

Create Or Replace Procedure Zl_�ٴ�·������_Insert
(
  ·��id_In �ٴ�·������.·��id%Type,
  �汾��_In �ٴ�·������.�汾��%Type,
  ���_In   �ٴ�·������.���%Type,
  ����_In   �ٴ�·������.����%Type,
  Clear_In  Number := 0
  --������
  --  Clear_IN������ǰ�Ƿ������ǰ�汾·�������з���
) Is
Begin
  If Nvl(Clear_In, 0) = 1 And ���_In = 1 Then
    Delete From �ٴ�·������ Where ·��id = ·��id_In And �汾�� = �汾��_In;
  End If;
  Insert Into �ٴ�·������ (·��id, �汾��, ���, ����) Values (·��id_In, �汾��_In, ���_In, ����_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ�·������_Insert;
/

Create Or Replace Procedure Zl_�ٴ�·���׶�_Delete(�׶�id_In Varchar2) Is
  --������
  --  �׶�ID_IN��ID1,ID2,...
Begin
  Delete /*+ Rule*/
  From �ٴ�·���׶�
  Where ID In (Select * From Table(Cast(f_Num2list(�׶�id_In) As Zltools.t_Numlist)));
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ�·���׶�_Delete;
/

Create Or Replace Procedure Zl_�ٴ�·���׶�_Insert
(
  Id_In       �ٴ�·���׶�.Id%Type,
  ·��id_In   �ٴ�·���׶�.·��id%Type,
  �汾��_In   �ٴ�·���׶�.�汾��%Type,
  ��id_In     �ٴ�·���׶�.��id%Type,
  ���_In     �ٴ�·���׶�.���%Type,
  ����_In     �ٴ�·���׶�.����%Type,
  ��ʼ����_In �ٴ�·���׶�.��ʼ����%Type,
  ��������_In �ٴ�·���׶�.��������%Type,
  ��־_In     �ٴ�·���׶�.��־%Type,
  ˵��_In     �ٴ�·���׶�.˵��%Type
) Is
Begin
  Insert Into �ٴ�·���׶�
    (ID, ·��id, �汾��, ��id, ���, ����, ��ʼ����, ��������, ��־, ˵��)
  Values
    (Id_In, ·��id_In, �汾��_In, ��id_In, ���_In, ����_In, ��ʼ����_In, ��������_In, ��־_In, ˵��_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ�·���׶�_Insert;
/

Create Or Replace Procedure Zl_�ٴ�·���׶�_Update
(
  Id_In       �ٴ�·���׶�.Id%Type,
  ·��id_In   �ٴ�·���׶�.·��id%Type,
  �汾��_In   �ٴ�·���׶�.�汾��%Type,
  ���_In     �ٴ�·���׶�.���%Type,
  ����_In     �ٴ�·���׶�.����%Type,
  ��ʼ����_In �ٴ�·���׶�.��ʼ����%Type,
  ��������_In �ٴ�·���׶�.��������%Type,
  ��־_In     �ٴ�·���׶�.��־%Type,
  ˵��_In     �ٴ�·���׶�.˵��%Type
) Is
Begin
  Update �ٴ�·���׶�
  Set ��� = ���_In, ���� = ����_In, ��ʼ���� = ��ʼ����_In, �������� = ��������_In, ��־ = ��־_In, ˵�� = ˵��_In
  Where ID = Id_In And ·��id = ·��id_In And �汾�� = �汾��_In;

  --ɾ����Ӧ�Ľ׶�������Ϣ,�������±���
  Delete From �ٴ�·������ Where ·��id = ·��id_In And �汾�� = �汾��_In And �׶�id = Id_In And �������� = 2;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ�·���׶�_Update;
/

Create Or Replace Procedure Zl_·��ҽ������_Insert
(
  Id_In         ·��ҽ������.Id%Type,
  ���id_In     ·��ҽ������.���id%Type,
  ���_In       ·��ҽ������.���%Type,
  ��Ч_In       ·��ҽ������.��Ч%Type,
  ������Ŀid_In ·��ҽ������.������Ŀid%Type,
  ҽ������_In   ·��ҽ������.ҽ������%Type,
  ��������_In   ·��ҽ������.��������%Type,
  �ܸ�����_In   ·��ҽ������.�ܸ�����%Type,
  �շ�ϸĿid_In ·��ҽ������.�շ�ϸĿid%Type,
  �걾��λ_In   ·��ҽ������.�걾��λ%Type,
	��鷽��_In   ·��ҽ������.��鷽��%Type,
  ִ��Ƶ��_In   ·��ҽ������.ִ��Ƶ��%Type,
  Ƶ�ʴ���_In   ·��ҽ������.Ƶ�ʴ���%Type,
  Ƶ�ʼ��_In   ·��ҽ������.Ƶ�ʼ��%Type,
  �����λ_In   ·��ҽ������.�����λ%Type,
  ҽ������_In   ·��ҽ������.ҽ������%Type,
  ִ������_In   ·��ҽ������.ִ������%Type,
  ִ�п���id_In ·��ҽ������.ִ�п���id%Type,
  ʱ�䷽��_In   ·��ҽ������.ʱ�䷽��%Type,
  ·��id_In     �ٴ�·����Ŀ.·��id%Type := Null,
  �汾��_In     �ٴ�·����Ŀ.�汾��%Type := Null
) Is
  --������
  --  ·��ID_IN,�汾��_IN��������ʱ����ʾҪ���ָ���汾��·�����е�����ҽ�����ݺ͹�������
Begin
  If ·��id_In Is Not Null And �汾��_In Is Not Null Then
    --�ἶ��ɾ��
    --Delete From �ٴ�·��ҽ��
    --Where ·����Ŀid In (Select ID From �ٴ�·����Ŀ Where ·��id = ·��id_In And �汾�� = �汾��_In);
  
    Delete From ·��ҽ������
    Where ID In (Select ҽ������id
                 From �ٴ�·����Ŀ A, �ٴ�·��ҽ�� B
                 Where a.Id = b.·����Ŀid And a.·��id = ·��id_In And a.�汾�� = �汾��_In);
  End If;

  Insert Into ·��ҽ������
    (ID, ���id, ���, ��Ч, ������Ŀid, ҽ������, ��������, �ܸ�����, �շ�ϸĿid, �걾��λ, ��鷽��, ִ��Ƶ��, Ƶ�ʴ���, Ƶ�ʼ��, �����λ, ҽ������, ִ������, ִ�п���id, ʱ�䷽��)
  Values
    (Id_In, ���id_In, ���_In, ��Ч_In, ������Ŀid_In, ҽ������_In, ��������_In, �ܸ�����_In, �շ�ϸĿid_In, �걾��λ_In, ��鷽��_In, ִ��Ƶ��_In, Ƶ�ʴ���_In,
     Ƶ�ʼ��_In, �����λ_In, ҽ������_In, ִ������_In, ִ�п���id_In, ʱ�䷽��_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_·��ҽ������_Insert;
/

Create Or Replace Procedure Zl_�ٴ�·����Ŀ_Insert
(
  Id_In       �ٴ�·����Ŀ.Id%Type,
  ·��id_In   �ٴ�·����Ŀ.·��id%Type,
  �汾��_In   �ٴ�·����Ŀ.�汾��%Type,
  �׶�id_In   �ٴ�·����Ŀ.�׶�id%Type,
  ����_In     �ٴ�·����Ŀ.����%Type,
  ��Ŀ���_In �ٴ�·����Ŀ.��Ŀ���%Type,
  ��Ŀ����_In �ٴ�·����Ŀ.��Ŀ����%Type,
  ִ�з�ʽ_In �ٴ�·����Ŀ.ִ�з�ʽ%Type,
  ִ����_In   �ٴ�·����Ŀ.ִ����%Type,
  ��Ŀ���_In �ٴ�·����Ŀ.��Ŀ���%Type,
  ͼ��id_In   �ٴ�·����Ŀ.ͼ��id%Type,
  ҽ��id_In   Varchar2,
  ����id_In   Varchar2
  --������
  --   ҽ��ID_IN����Ӧ·��ҽ�����ݵ�ID����ʽΪID1,ID2,....
  --   ����ID_IN����Ӧ�����ļ��б��ID����ʽΪID1,ID2,...
) Is
Begin
  Insert Into �ٴ�·����Ŀ
    (ID, ·��id, �汾��, �׶�id, ����, ��Ŀ���, ��Ŀ����, ִ�з�ʽ, ִ����, ��Ŀ���, ͼ��id)
  Values
    (Id_In, ·��id_In, �汾��_In, �׶�id_In, ����_In, ��Ŀ���_In, ��Ŀ����_In, ִ�з�ʽ_In, ִ����_In, ��Ŀ���_In, ͼ��id_In);

  --����ҽ������
  If ҽ��id_In Is Not Null Then
    For r_Row In (Select /*+ Rule*/
                   Column_Value As ID
                  From Table(Cast(f_Num2list(ҽ��id_In) As Zltools.t_Numlist))) Loop
      Insert Into �ٴ�·��ҽ�� (·����Ŀid, ҽ������id) Values (Id_In, r_Row.Id);
    End Loop;
  End If;

  --����������
  If ����id_In Is Not Null Then
    For r_Row In (Select /*+ Rule*/
                   Column_Value As ID
                  From Table(Cast(f_Num2list(����id_In) As Zltools.t_Numlist))) Loop
      Insert Into �ٴ�·������ (��Ŀid, �ļ�id) Values (Id_In, r_Row.Id);
    End Loop;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ�·����Ŀ_Insert;
/

Create Or Replace Procedure Zl_�ٴ�·����Ŀ_Update
(
  Id_In       �ٴ�·����Ŀ.Id%Type,
  ·��id_In   �ٴ�·����Ŀ.·��id%Type,
  �汾��_In   �ٴ�·����Ŀ.�汾��%Type,
  ��Ŀ���_In �ٴ�·����Ŀ.��Ŀ���%Type,
  ��Ŀ����_In �ٴ�·����Ŀ.��Ŀ����%Type,
  ִ�з�ʽ_In �ٴ�·����Ŀ.ִ�з�ʽ%Type,
  ִ����_In   �ٴ�·����Ŀ.ִ����%Type,
  ��Ŀ���_In �ٴ�·����Ŀ.��Ŀ���%Type,
  ͼ��id_In   �ٴ�·����Ŀ.ͼ��id%Type,
  ҽ��id_In   Varchar2,
  ����id_In   Varchar2
) Is
Begin
  Update �ٴ�·����Ŀ
  Set ��Ŀ��� = ��Ŀ���_In, ��Ŀ���� = ��Ŀ����_In, ִ�з�ʽ = ִ�з�ʽ_In, ִ���� = ִ����_In, ��Ŀ��� = ��Ŀ���_In, ͼ��id = ͼ��id_In
  Where ID = Id_In And ·��id = ·��id_In And �汾�� = �汾��_In;

  --����ҽ������
  Delete From �ٴ�·��ҽ�� Where ·����Ŀid = Id_In;
  If ҽ��id_In Is Not Null Then
    For r_Row In (Select /*+ Rule*/
                   Column_Value As ID
                  From Table(Cast(f_Num2list(ҽ��id_In) As Zltools.t_Numlist))) Loop
      Insert Into �ٴ�·��ҽ�� (·����Ŀid, ҽ������id) Values (Id_In, r_Row.Id);
    End Loop;
  End If;

  --����������
  Delete From �ٴ�·������ Where ��Ŀid = Id_In;
  If ����id_In Is Not Null Then
    For r_Row In (Select /*+ Rule*/
                   Column_Value As ID
                  From Table(Cast(f_Num2list(����id_In) As Zltools.t_Numlist))) Loop
      Insert Into �ٴ�·������ (��Ŀid, �ļ�id) Values (Id_In, r_Row.Id);
    End Loop;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ�·����Ŀ_Update;
/

Create Or Replace Procedure Zl_�ٴ�·����Ŀ_Delete(��Ŀid_In Varchar2) Is
  --������
  --  ��Ŀid_In��ID1,ID2,...
Begin
  Delete /*+ Rule*/
  From �ٴ�·����Ŀ
  Where ID In (Select * From Table(Cast(f_Num2list(��Ŀid_In) As Zltools.t_Numlist)));
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ�·����Ŀ_Delete;
/

Create Or Replace Procedure Zl_�ٴ�·���汾_Audit
(
  ·��id_In �ٴ�·����Ŀ.·��id%Type,
  �汾��_In �ٴ�·����Ŀ.�汾��%Type,
  ���_In   Number
  --������
  --   ���_IN��1=ͨ����ˣ�-1=ȡ�����
) Is
  v_Date  Date;
  v_Count Number;
  v_User  ��Ա��.����%Type;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  If ���_In = 1 Then
    --���
    Select Sysdate Into v_Date From Dual;
    Select zl_UserName Into v_User From Dual;
  
    Update �ٴ�·���汾
    Set ����� = v_User, ���ʱ�� = v_Date
    Where ·��id = ·��id_In And �汾�� = �汾��_In And ���ʱ�� Is Null;
    If Sql%RowCount > 0 Then
      --�Զ�ͣ��֮ǰ�İ汾
      Update �ٴ�·���汾
      Set ͣ���� = v_User, ͣ��ʱ�� = v_Date
      Where ·��id = ·��id_In And �汾�� < �汾��_In And ͣ��ʱ�� Is Null;
    
      Update �ٴ�·��Ŀ¼ Set ���°汾 = �汾��_In Where ID = ·��id_In;
    End If;
  Elsif ���_In = -1 Then
    --ȡ�����
    Select Count(*) Into v_Count From �����ٴ�·�� Where ·��id = ·��id_In And �汾�� = �汾��_In And Rownum = 1;
    If Nvl(v_Count, 0) > 0 Then
      v_Error := '�ð汾���ٴ�·���Ѿ���ʹ�ã�����ȡ����ˡ�';
      Raise Err_Custom;
    End If;
  
    Select Count(*) Into v_Count From �ٴ�·���汾 Where ·��id = ·��id_In And �汾�� > �汾��_In;
    If Nvl(v_Count, 0) > 0 Then
      v_Error := '�ð汾������������µİ汾������ȡ����ˡ�';
      Raise Err_Custom;
    End If;
  
    Select �����, ���ʱ�� Into v_User, v_Date From �ٴ�·���汾 Where ·��id = ·��id_In And �汾�� = �汾��_In;
  
    Update �ٴ�·���汾
    Set ����� = Null, ���ʱ�� = Null
    Where ·��id = ·��id_In And �汾�� = �汾��_In And ���ʱ�� Is Not Null;
    If Sql%RowCount > 0 Then
      --�ָ�֮ǰ���ʱ�Զ�ͣ�õİ汾(�ֹ�ͣ�õĲ�����)
      Select Max(�汾��)
      Into v_Count
      From �ٴ�·���汾
      Where ·��id = ·��id_In And �汾�� < �汾��_In And ͣ���� = v_User And ͣ��ʱ�� = v_Date;
      If Nvl(v_Count, 0) > 0 Then
        Update �ٴ�·���汾 Set ͣ���� = Null, ͣ��ʱ�� = Null Where ·��id = ·��id_In And �汾�� = v_Count;
      End If;
    
      --�������°汾
      Select Max(�汾��)
      Into v_Count
      From �ٴ�·���汾
      Where ·��id = ·��id_In And ���ʱ�� Is Not Null And ͣ��ʱ�� Is Null;
      If Nvl(v_Count, 0) > 0 Then
        Update �ٴ�·��Ŀ¼ Set ���°汾 = v_Count Where ID = ·��id_In;
      Else
        Update �ٴ�·��Ŀ¼ Set ���°汾 = Null Where ID = ·��id_In;
      End If;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ�·���汾_Audit;
/

Create Or Replace Procedure Zl_�ٴ�·���汾_Stop
(
  ·��id_In �ٴ�·����Ŀ.·��id%Type,
  �汾��_In �ٴ�·����Ŀ.�汾��%Type,
  ͣ��_In   Number
  --������
  --   ͣ��_In��1=ͣ�ã�-1=ȡ��ͣ��
) Is
  v_Date  Date;
  v_Count Number;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  If ͣ��_In = 1 Then
    Select ���ʱ�� Into v_Date From �ٴ�·���汾 Where ·��id = ·��id_In And �汾�� = �汾��_In;
    If v_Date Is Null Then
      v_Error := '�ð汾���ٴ�·����δ��ˣ�����Ҫͣ�á�';
      Raise Err_Custom;
    End If;
  
    Update �ٴ�·���汾
    Set ͣ���� = zl_UserName, ͣ��ʱ�� = Sysdate
    Where ·��id = ·��id_In And �汾�� = �汾��_In And ͣ��ʱ�� Is Null;
  Elsif ͣ��_In = -1 Then
    Select Count(*)
    Into v_Count
    From �ٴ�·���汾
    Where ·��id = ·��id_In And �汾�� > �汾��_In And (ͣ��ʱ�� Is Not Null Or ���ʱ�� Is Not Null);
    If Nvl(v_Count, 0) > 0 Then
      v_Error := '�ð汾������������Ѿ���˻���ͣ�õİ汾������ȡ��ͣ�á�';
      Raise Err_Custom;
    End If;
  
    Update �ٴ�·���汾
    Set ͣ���� = Null, ͣ��ʱ�� = Null
    Where ·��id = ·��id_In And �汾�� = �汾��_In And ͣ��ʱ�� Is Not Null;
  End If;

  --�������°汾
  Select Max(�汾��)
  Into v_Count
  From �ٴ�·���汾
  Where ·��id = ·��id_In And ���ʱ�� Is Not Null And ͣ��ʱ�� Is Null;
  If Nvl(v_Count, 0) > 0 Then
    Update �ٴ�·��Ŀ¼ Set ���°汾 = v_Count Where ID = ·��id_In;
  Else
    Update �ٴ�·��Ŀ¼ Set ���°汾 = Null Where ID = ·��id_In;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ�·���汾_Stop;
/

Create Or Replace Procedure Zl_�ٴ�·���汾_Delete
(
  ·��id_In �ٴ�·���汾.·��id%Type,
  �汾��_In �ٴ�·���汾.�汾��%Type
) Is
  v_Count Number;
Begin
  Delete From ·��ҽ������
  Where ID In (Select ҽ������id
               From �ٴ�·����Ŀ A, �ٴ�·��ҽ�� B
               Where a.Id = b.·����Ŀid And a.·��id = ·��id_In And a.�汾�� = �汾��_In);

  --�������Զ�����ɾ��
  Delete From �ٴ�·���汾 Where ·��id = ·��id_In And �汾�� = �汾��_In;

  --�������°汾
  Select Max(�汾��)
  Into v_Count
  From �ٴ�·���汾
  Where ·��id = ·��id_In And ���ʱ�� Is Not Null And ͣ��ʱ�� Is Null;
  If Nvl(v_Count, 0) > 0 Then
    Update �ٴ�·��Ŀ¼ Set ���°汾 = v_Count Where ID = ·��id_In;
  Else
    Update �ٴ�·��Ŀ¼ Set ���°汾 = Null Where ID = ·��id_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ�·���汾_Delete;
/

Create Or Replace Procedure Zl_·������ָ��_Insert
(
  ·��id_In   �ٴ�·������.·��id%Type,
  �汾��_In   �ٴ�·������.�汾��%Type,
  �׶�id_In   �ٴ�·������.�׶�id%Type,
  ��������_In �ٴ�·������.��������%Type,
  ָ��id_In   ·������ָ��.Id%Type,
  ���_In     ·������ָ��.���%Type,
  ����ָ��_In ·������ָ��.����ָ��%Type,
  ָ������_In ·������ָ��.ָ������%Type,
  ָ����_In ·������ָ��.ָ����%Type
) Is
  v_����id �ٴ�·������.Id%Type;
Begin
  Begin
    Select ID
    Into v_����id
    From �ٴ�·������
    Where ·��id = ·��id_In And �汾�� = �汾��_In And Nvl(�׶�id, 0) = Nvl(�׶�id_In, 0) And �������� = ��������_In;
  Exception
    When Others Then
      Null;
  End;

  If v_����id Is Null Then
    Select �ٴ�·������_Id.Nextval Into v_����id From Dual;
    Insert Into �ٴ�·������
      (ID, ·��id, �汾��, �׶�id, ��������)
    Values
      (v_����id, ·��id_In, �汾��_In, �׶�id_In, ��������_In);
  End If;

  Insert Into ·������ָ��
    (ID, ����id, ���, ����ָ��, ָ������, ָ����)
  Values
    (ָ��id_In, v_����id, ���_In, ����ָ��_In, ָ������_In, ָ����_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_·������ָ��_Insert;
/

Create Or Replace Procedure Zl_·����������_Insert
(
  ·��id_In   �ٴ�·������.·��id%Type,
  �汾��_In   �ٴ�·������.�汾��%Type,
  �׶�id_In   �ٴ�·������.�׶�id%Type,
  ��������_In �ٴ�·������.��������%Type,
  ָ��id_In   ·����������.ָ��id%Type,
  ��Ŀid_In   ·����������.��Ŀid%Type,
  ��ϵʽ_In   ·����������.��ϵʽ%Type,
  ����ֵ_In   ·����������.����ֵ%Type,
  �������_In ·����������.�������%Type
) Is
  v_����id �ٴ�·������.Id%Type;
Begin
  Begin
    Select ID
    Into v_����id
    From �ٴ�·������
    Where ·��id = ·��id_In And �汾�� = �汾��_In And Nvl(�׶�id, 0) = Nvl(�׶�id_In, 0) And �������� = ��������_In;
  Exception
    When Others Then
      Null;
  End;

  If v_����id Is Null Then
    Select �ٴ�·������_Id.Nextval Into v_����id From Dual;
    Insert Into �ٴ�·������
      (ID, ·��id, �汾��, �׶�id, ��������)
    Values
      (v_����id, ·��id_In, �汾��_In, �׶�id_In, ��������_In);
  End If;

  Insert Into ·����������
    (����id, ָ��id, ��Ŀid, ��ϵʽ, ����ֵ, �������)
  Values
    (v_����id, ָ��id_In, ��Ŀid_In, ��ϵʽ_In, ����ֵ_In, �������_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_·����������_Insert;
/

Create Or Replace Procedure Zl_�ٴ�·���汾_Copy
(
  Դ·��id_In   �ٴ�·���汾.·��id%Type,
  Դ�汾��_In   �ٴ�·���汾.�汾��%Type,
  Ŀ��·��id_In �ٴ�·���汾.·��id%Type,
  Ŀ��汾��_In �ٴ�·���汾.�汾��%Type
  --���ܣ����Ʋ����µ��ٴ�·���汾
  --������
  --  Դ�汾��_In�����δָ��(0��NULL)����ȡ������Ч�İ汾��
  --  Ŀ�걾��_In�����δָ��(0��NULL)��������µİ汾��
) Is
  v_Դ�汾��   �ٴ�·���汾.�汾��%Type;
  v_Ŀ��汾�� �ٴ�·���汾.�汾��%Type;

  v_Advice_Id Number;
  v_Step_Id   Number;
  v_Item_Id   Number;
  v_Eval_Id   Number;
  v_Mark_Id   Number;

  v_Error Varchar2(255);
  Err_Custom Exception;

  --����������غ���
  Type t_Id_Table Is Table Of Number;
  Arr_Id t_Id_Table;

  Procedure Adjuest_Sequence_Advice(n_Count Number) Is
  Begin
    Select ·��ҽ������_Id.Nextval Bulk Collect Into Arr_Id From Dual Connect By Rownum <= n_Count;
  End;
  Procedure Adjuest_Sequence_Step(n_Count Number) Is
  Begin
    Select �ٴ�·���׶�_Id.Nextval Bulk Collect Into Arr_Id From Dual Connect By Rownum <= n_Count;
  End;
  Procedure Adjuest_Sequence_Item(n_Count Number) Is
  Begin
    Select �ٴ�·����Ŀ_Id.Nextval Bulk Collect Into Arr_Id From Dual Connect By Rownum <= n_Count;
  End;
  Procedure Adjuest_Sequence_Eval(n_Count Number) Is
  Begin
    Select �ٴ�·������_Id.Nextval Bulk Collect Into Arr_Id From Dual Connect By Rownum <= n_Count;
  End;
  Procedure Adjuest_Sequence_Mark(n_Count Number) Is
  Begin
    Select ·������ָ��_Id.Nextval Bulk Collect Into Arr_Id From Dual Connect By Rownum <= n_Count;
  End;
Begin
  --ȷ��Դ·���汾��
  v_Դ�汾�� := Nvl(Դ�汾��_In, 0);
  If v_Դ�汾�� = 0 Then
    Select ���°汾 Into v_Դ�汾�� From �ٴ�·��Ŀ¼ Where ID = Դ·��id_In;
    If Nvl(v_Դ�汾��, 0) = 0 Then
      v_Error := 'Ҫ���Ƶ���Դ�ٴ�·����û�п��õ���Ч�汾��';
      Raise Err_Custom;
    End If;
  End If;

  --ȷ��Ŀ��·���汾��
  v_Ŀ��汾�� := Nvl(Ŀ��汾��_In, 0);
  If v_Ŀ��汾�� = 0 Then
    Select Nvl(Max(�汾��), 0) + 1 Into v_Ŀ��汾�� From �ٴ�·���汾 Where ·��id = Ŀ��·��id_In;
  Else
    Zl_�ٴ�·���汾_Delete(Ŀ��·��id_In, Ŀ��汾��_In);
  End If;

  --·��ҽ������
  Select ·��ҽ������_Id.Currval, ·��ҽ������_Id.Nextval Into v_Advice_Id, v_Advice_Id From Dual;

  Select v_Advice_Id - Nvl(Min(ID), 0) + 1
  Into v_Advice_Id
  From ·��ҽ������
  Where ID In (Select b.ҽ������id
               From �ٴ�·����Ŀ A, �ٴ�·��ҽ�� B
               Where a.Id = b.·����Ŀid And a.·��id = Դ·��id_In And a.�汾�� = v_Դ�汾��);

  Insert Into ·��ҽ������
    (ID, ���id, ���, ��Ч, ������Ŀid, ҽ������, ��������, �ܸ�����, �շ�ϸĿid, �걾��λ, ��鷽��, ִ��Ƶ��, Ƶ�ʴ���, Ƶ�ʼ��, �����λ, ҽ������, ִ������, ִ�п���id, ʱ�䷽��)
    Select ID + v_Advice_Id, ���id + v_Advice_Id, ���, ��Ч, ������Ŀid, ҽ������, ��������, �ܸ�����, �շ�ϸĿid, �걾��λ, ��鷽��, ִ��Ƶ��, Ƶ�ʴ���, Ƶ�ʼ��,
           �����λ, ҽ������, ִ������, ִ�п���id, ʱ�䷽��
    From ·��ҽ������
    Where ID In (Select b.ҽ������id
                 From �ٴ�·����Ŀ A, �ٴ�·��ҽ�� B
                 Where a.Id = b.·����Ŀid And a.·��id = Դ·��id_In And a.�汾�� = v_Դ�汾��);
  Adjuest_Sequence_Advice(v_Advice_Id); --��������

  --�ٴ�·���汾
  Insert Into �ٴ�·���汾
    (·��id, �汾��, ��׼סԺ��, ��׼����, �汾˵��, ������, ����ʱ��)
    Select Ŀ��·��id_In, v_Ŀ��汾��, ��׼סԺ��, ��׼����, �汾˵��, ������, ����ʱ��
    From �ٴ�·���汾
    Where ·��id = Դ·��id_In And �汾�� = v_Դ�汾��;

  --�ٴ�·������
  Insert Into �ٴ�·������
    (·��id, �汾��, ���, ����)
    Select Ŀ��·��id_In, v_Ŀ��汾��, ���, ����
    From �ٴ�·������
    Where ·��id = Դ·��id_In And �汾�� = v_Դ�汾��;

  --�ٴ�·���׶�
  Select �ٴ�·���׶�_Id.Currval, �ٴ�·���׶�_Id.Nextval Into v_Step_Id, v_Step_Id From Dual;
  Select v_Step_Id - Nvl(Min(ID), 0) + 1
  Into v_Step_Id
  From �ٴ�·���׶�
  Where ·��id = Դ·��id_In And �汾�� = v_Դ�汾��;

  Insert Into �ٴ�·���׶�
    (ID, ·��id, �汾��, ��id, ���, ����, ��ʼ����, ��������, ��־, ˵��)
    Select ID + v_Step_Id, Ŀ��·��id_In, v_Ŀ��汾��, ��id + v_Step_Id, ���, ����, ��ʼ����, ��������, ��־, ˵��
    From �ٴ�·���׶�
    Where ·��id = Դ·��id_In And �汾�� = v_Դ�汾��;
  Adjuest_Sequence_Step(v_Step_Id); --��������

  --�ٴ�·����Ŀ
  Select �ٴ�·����Ŀ_Id.Currval, �ٴ�·����Ŀ_Id.Nextval Into v_Item_Id, v_Item_Id From Dual;
  Select v_Item_Id - Nvl(Min(ID), 0) + 1
  Into v_Item_Id
  From �ٴ�·����Ŀ
  Where ·��id = Դ·��id_In And �汾�� = v_Դ�汾��;

  Insert Into �ٴ�·����Ŀ
    (ID, ·��id, �汾��, �׶�id, ����, ��Ŀ���, ��Ŀ����, ִ�з�ʽ, ִ����, ��Ŀ���, ͼ��id)
    Select ID + v_Item_Id, Ŀ��·��id_In, v_Ŀ��汾��, �׶�id + v_Step_Id, ����, ��Ŀ���, ��Ŀ����, ִ�з�ʽ, ִ����, ��Ŀ���, ͼ��id
    From �ٴ�·����Ŀ
    Where ·��id = Դ·��id_In And �汾�� = v_Դ�汾��;
  Adjuest_Sequence_Item(v_Item_Id); --��������

  --�ٴ�·��ҽ��
  Insert Into �ٴ�·��ҽ��
    (·����Ŀid, ҽ������id)
    Select b.·����Ŀid + v_Item_Id, b.ҽ������id + v_Advice_Id
    From �ٴ�·����Ŀ A, �ٴ�·��ҽ�� B
    Where a.Id = b.·����Ŀid And a.·��id = Դ·��id_In And a.�汾�� = v_Դ�汾��;

  --�ٴ�·������
  Insert Into �ٴ�·������
    (��Ŀid, �ļ�id)
    Select b.��Ŀid + v_Item_Id, b.�ļ�id
    From �ٴ�·����Ŀ A, �ٴ�·������ B
    Where a.Id = b.��Ŀid And a.·��id = Դ·��id_In And a.�汾�� = v_Դ�汾��;

  --�ٴ�·������
  Select �ٴ�·������_Id.Currval, �ٴ�·������_Id.Nextval Into v_Eval_Id, v_Eval_Id From Dual;
  Select v_Eval_Id - Nvl(Min(ID), 0) + 1
  Into v_Eval_Id
  From �ٴ�·������
  Where ·��id = Դ·��id_In And �汾�� = v_Դ�汾��;

  Insert Into �ٴ�·������
    (ID, ·��id, �汾��, �׶�id, ��������)
    Select ID + v_Eval_Id, Ŀ��·��id_In, v_Ŀ��汾��, �׶�id + v_Step_Id, ��������
    From �ٴ�·������
    Where ·��id = Դ·��id_In And �汾�� = v_Դ�汾��;
  Adjuest_Sequence_Eval(v_Eval_Id); --��������

  --·������ָ��
  Select ·������ָ��_Id.Currval, ·������ָ��_Id.Nextval Into v_Mark_Id, v_Mark_Id From Dual;
  Select v_Mark_Id - Nvl(Min(ID), 0) + 1
  Into v_Mark_Id
  From ·������ָ��
  Where ����id In (Select ID From �ٴ�·������ Where ·��id = Դ·��id_In And �汾�� = v_Դ�汾��);

  Insert Into ·������ָ��
    (ID, ����id, ���, ����ָ��, ָ������, ָ����)
    Select ID + v_Mark_Id, ����id + v_Eval_Id, ���, ����ָ��, ָ������, ָ����
    From ·������ָ��
    Where ����id In (Select ID From �ٴ�·������ Where ·��id = Դ·��id_In And �汾�� = v_Դ�汾��);
  Adjuest_Sequence_Mark(v_Mark_Id); --��������

  --·����������
  Insert Into ·����������
    (����id, ָ��id, ��Ŀid, ��ϵʽ, ����ֵ, �������)
    Select ����id + v_Eval_Id, ָ��id + v_Mark_Id, ��Ŀid + v_Item_Id, ��ϵʽ, ����ֵ, �������
    From ·����������
    Where ����id In (Select ID From �ٴ�·������ Where ·��id = Դ·��id_In And �汾�� = v_Դ�汾��);
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ�·���汾_Copy;
/












--�ٴ�·��Ӧ����ع���
---------------------------------------------------------------------------------------------------------------------------

Create Or Replace Procedure Zl_����·������_Insert
(
  ����id_In   �����ٴ�·��.����id%Type,
  ��ҳid_In   �����ٴ�·��.��ҳid%Type,
  ����id_In   �����ٴ�·��.����id%Type,
  ·��id_In   �����ٴ�·��.·��id%Type,
  �汾��_In   �����ٴ�·��.�汾��%Type,
  ·����¼_In �����ٴ�·��.Id%Type,
  ������_In   �����ٴ�·��.������%Type,
  ����˵��_In �����ٴ�·��.����˵��%Type,
  ���ϵ���_In �����ٴ�·��.״̬%Type, --0=������,1=����
  ָ������_In Varchar2, --ָ������|ָ����|ָ������||...,ĩβ��||,����Ϊ��
  ���_In     Number
) Is
  v_Str   Varchar2(4000);
  v_Tmp   Varchar2(1000);
  v_Index Number;
  I       Number(5) := 1;

  l_ָ������ t_Strlist := t_Strlist();
  l_ָ���� t_Strlist := t_Strlist();
  l_ָ������ t_Numlist := t_Numlist();

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  If ���_In = 1 Then
    Insert Into �����ٴ�·��
      (ID, ����id, ��ҳid, ����id, ·��id, �汾��, ������, ����ʱ��, ����˵��, ״̬)
    Values
      (·����¼_In, ����id_In, ��ҳid_In, ����id_In, ·��id_In, �汾��_In, ������_In, Sysdate, ����˵��_In, ���ϵ���_In);
  End If;

  If Not ָ������_In Is Null Then
    v_Str := ָ������_In;
    Loop
      v_Index := Instr(v_Str, '||');
      Exit When(Nvl(v_Index, 0) = 0);
      l_ָ������.Extend;
      l_ָ����.Extend;
      l_ָ������.Extend;
    
      v_Tmp := Substr(v_Str, 1, v_Index - 1);
      l_ָ������(I) := Substr(v_Tmp, 1, Instr(v_Tmp, '|') - 1);
      v_Tmp := Substr(v_Tmp, Instr(v_Tmp, '|') + 1);
      l_ָ����(I) := Trim(Substr(v_Tmp, 1, Instr(v_Tmp, '|') - 1));
      v_Tmp := Substr(v_Tmp, Instr(v_Tmp, '|') + 1);
      l_ָ������(I) := To_Number(v_Tmp);
    
      v_Str := Substr(v_Str, v_Index + 2);
      I     := I + 1;
    End Loop;
  
    Forall I In 1 .. l_ָ������.Count
      Insert Into ����·��ָ��
        (·����¼id, ��������, ����ָ��, ָ����, ָ������)
      Values
        (·����¼_In, 1, l_ָ������(I), l_ָ����(I), l_ָ������(I));
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����·������_Insert;
/


Create Or Replace Procedure Zl_����·������_Delete(����·��id_In �����ٴ�·��.Id%Type) Is
  v_Count �����ٴ�·��.Id%Type;
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  Select Nvl(Max(·����¼id), 0) Into v_Count From ����·��ִ�� Where ·����¼id = ����·��id_In;

  If v_Count = 0 Then
    Delete ����·��ָ�� Where ·����¼id = ����·��id_In;
    Delete �����ٴ�·�� Where ID = ����·��id_In;
  Else
    v_Error := '�ò��˵�·����������·����Ŀ,����ȡ�����롣';
    Raise Err_Custom;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����·������_Delete;
/

Create Or Replace Procedure Zl_����·������_Insert
(
  ����_In       Number, --1=����,2=�޸�
  ·����¼id_In �����ٴ�·��.Id%Type,
  �׶�id_In     �ٴ�·���׶�.Id%Type,
  ����_In       ����·������.����%Type,
  ����_In       ����·������.����%Type,
  ������_In     ����·������.������%Type,
  �������_In   ����·������.�������%Type,
  ����˵��_In   ����·������.����˵��%Type,
  �Ǽ���_In     ����·������.�Ǽ���%Type,
  ָ������_In   Varchar2, --ָ������|ָ����|ָ������||...,ĩβ��||,����Ϊ��
  ���_In       Number
) Is
  v_Str   Varchar2(4000);
  v_Tmp   Varchar2(1000);
  v_Index Number;
  I       Number(5) := 1;

  l_ָ������ t_Strlist := t_Strlist();
  l_ָ���� t_Strlist := t_Strlist();
  l_ָ������ t_Numlist := t_Numlist();

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin

  If ���_In = 1 Then
    If ����_In = 1 Then
      Insert Into ����·������
        (·����¼id, �׶�id, ����, ����, ������, ����ʱ��, �������, ����˵��, �Ǽ���, �Ǽ�ʱ��)
      Values
        (·����¼id_In, �׶�id_In, ����_In, ����_In, ������_In, Sysdate, �������_In, ����˵��_In, �Ǽ���_In, Sysdate);
    Else
      Update ����·������
      Set ������ = ������_In, ����ʱ�� = Sysdate, ������� = �������_In, ����˵�� = ����˵��_In, �Ǽ��� = �Ǽ���_In, �Ǽ�ʱ�� = Sysdate
      Where ·����¼id = ·����¼id_In And �׶�id = �׶�id_In And ���� = ����_In;
    End If;
  End If;

  If Not ָ������_In Is Null Then
    v_Str := ָ������_In;
    Loop
      v_Index := Instr(v_Str, '||');
      Exit When(Nvl(v_Index, 0) = 0);
      l_ָ������.Extend;
      l_ָ����.Extend;
      l_ָ������.Extend;
    
      v_Tmp := Substr(v_Str, 1, v_Index - 1);
      l_ָ������(I) := Substr(v_Tmp, 1, Instr(v_Tmp, '|') - 1);
      v_Tmp := Substr(v_Tmp, Instr(v_Tmp, '|') + 1);
      l_ָ����(I) := Trim(Substr(v_Tmp, 1, Instr(v_Tmp, '|') - 1));
      v_Tmp := Substr(v_Tmp, Instr(v_Tmp, '|') + 1);
      l_ָ������(I) := To_Number(v_Tmp);
    
      v_Str := Substr(v_Str, v_Index + 2);
      I     := I + 1;
    End Loop;
  
    If ����_In = 1 Then
      Forall I In 1 .. l_ָ������.Count
      
        Insert Into ����·��ָ��
          (·����¼id, �׶�id, ����, ����, ��������, ����ָ��, ָ����, ָ������)
        Values
          (·����¼id_In, �׶�id_In, ����_In, ����_In, 2, l_ָ������(I), l_ָ����(I), l_ָ������(I));
    Else
      Forall I In 1 .. l_ָ������.Count
        Update ����·��ָ��
        Set ָ���� = l_ָ����(I)
        Where ·����¼id = ·����¼id_In And �׶�id = �׶�id_In And ���� = ����_In And ����ָ�� = l_ָ������(I);
    End If;
  End If;

  If �������_In = -1 Then
    If ����_In = 2 Then
      v_Index := 0;
      Select ��ǰ�׶�id Into v_Index From �����ٴ�·�� Where ID = ·����¼id_In;
      If v_Index <> �׶�id_In Then
        v_Error := '�ò����������˴��յ�·����Ŀ,�����޸��������������·����';
        Raise Err_Custom;
      End If;
    End If;
    Update �����ٴ�·��
    Set ����ʱ�� = Sysdate, ״̬ = 3, ǰһ�׶�id = �׶�id_In, ��ǰ�׶�id = Null, ��ǰ���� = Null
    Where ID = ·����¼id_In;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����·������_Insert;
/

Create Or Replace Procedure Zl_����·������_Delete
(
  ·����¼id_In ����·��ִ��.Id%Type,
  �׶�id_In     ����·��ִ��.�׶�id%Type,
  ����_In       ����·��ִ��.����%Type
) Is
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin

  --�������Ϊ����ʱ�Զ�������,ȡ�������Զ�ȡ������
  Delete ����·������ Where ·����¼id = ·����¼id_In And �׶�id = �׶�id_In And ���� = ����_In;
  Delete ����·��ָ�� Where ·����¼id = ·����¼id_In And �׶�id = �׶�id_In And ���� = ����_In;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����·������_Delete;
/


Create Or Replace Procedure Zl_����·������_Insert
(
  ���_In        Number,
  ����id_In      �����ٴ�·��.����id%Type,
  ��ҳid_In      �����ٴ�·��.��ҳid%Type,
  Ӥ��_In        ���Ӳ�����¼.Ӥ��%Type,
  ����id_In      �����ٴ�·��.����id%Type,
  ·����¼id_In  ����·��ִ��.·����¼id%Type,
  �׶�id_In      ����·��ִ��.�׶�id%Type,
  ����_In        ����·��ִ��.����%Type,
  ����_In        ����·��ִ��.����%Type,
  ����_In        ����·��ִ��.����%Type,
  ��Ŀid_In      ����·��ִ��.��Ŀid%Type,
  ҽ��ids_In     Varchar2,
  �����ļ�ids_In Varchar2,
  ���˲���ids_In Varchar2,
  �Ǽ���_In      ����·��ִ��.�Ǽ���%Type,
  �Ǽ�ʱ��_In    ����·��ִ��.�Ǽ�ʱ��%Type,
  ��Ŀ����_In    ����·��ִ��.��Ŀ����%Type := Null,
  ִ����_In      ����·��ִ��.ִ����%Type := Null,
  ��Ŀ���_In    ����·��ִ��.��Ŀ���%Type := Null,
  ͼ��id_In      ����·��ִ��.ͼ��id%Type := Null,
  ���ԭ��_In    ����·��ִ��.���ԭ��%Type := Null
) Is
  v_��ǰ�׶�id �����ٴ�·��.��ǰ�׶�id%Type;
  v_·��ִ��id ����·��ִ��.Id%Type;
  v_����id     ���Ӳ�����¼.Id%Type;
  t_Advice     t_Numlist;
  t_File       t_Numlist;
  t_Doc        t_Numlist;

  v_Id         ���Ӳ�������.Id%Type;
  v_��id       ���Ӳ�������.��id%Type;
  v_��ǰ��id   ���Ӳ�������.��id%Type;
  v_�������   ���Ӳ�������.�������%Type;
  v_ԭ������� ���Ӳ�������.��id%Type;
  v_�����ı�   ���Ӳ�������.�����ı�%Type;
  n_Ԥ�����id ���Ӳ�������.Ԥ�����id%Type;

  v_��Ŀ��� ����·��ִ��.��Ŀ���%Type;
  v_Error    Varchar2(255);
  Err_Custom Exception;
Begin
  If ���_In = 1 And ��Ŀ����_In Is Null Then
    Select Nvl(��ǰ�׶�id, 0) Into v_��ǰ�׶�id From �����ٴ�·�� Where ID = ·����¼id_In;
    If v_��ǰ�׶�id <> �׶�id_In Then
      Update �����ٴ�·�� Set ǰһ�׶�id = ��ǰ�׶�id, ��ǰ�׶�id = �׶�id_In Where ID = ·����¼id_In;
    End If;
    Update �����ٴ�·�� Set ��ǰ���� = ����_In Where ID = ·����¼id_In;
  End If;

  --��ӵ�·������Ŀ
  If ��Ŀ����_In Is Not Null Then
    Select Max(��Ŀ���) + 1
    Into v_��Ŀ���
    From ����·��ִ��
    Where ·����¼id = ·����¼id_In And �׶�id = �׶�id_In And ���� = ����_In And ���� = ����_In;
    If v_��Ŀ��� Is Null Then
      --��������·����Ŀ֮��,��ʹ�п�ѡ����Ŀ���ܻ�δ����(���Բ�������,����Ԥ�����)
      Select Nvl(Max(��Ŀ���), 0) + 1
      Into v_��Ŀ���
      From �ٴ�·����Ŀ A, �����ٴ�·�� B
      Where a.·��id = b.·��id And a.�汾�� = b.�汾�� And b.Id = ·����¼id_In And a.�׶�id = �׶�id_In And ���� = ����_In;
    End If;
  End If;

  Select ����·��ִ��_Id.Nextval Into v_·��ִ��id From Dual;
  Insert Into ����·��ִ��
    (ID, ·����¼id, �׶�id, ����, ����, ����, ��Ŀid, �Ǽ���, �Ǽ�ʱ��, ��Ŀ���, ��Ŀ����, ִ����, ��Ŀ���, ͼ��id, ���ԭ��)
  Values
    (v_·��ִ��id, ·����¼id_In, �׶�id_In, ����_In, ����_In, ����_In, ��Ŀid_In, �Ǽ���_In, �Ǽ�ʱ��_In, v_��Ŀ���, ��Ŀ����_In, ִ����_In, ��Ŀ���_In,
     ͼ��id_In, ���ԭ��_In);

  If ҽ��ids_In Is Not Null Then
    Select Column_Value Bulk Collect Into t_Advice From Table(f_Num2list(ҽ��ids_In));
    Forall I In 1 .. t_Advice.Count
      Insert Into ����·��ҽ�� (·��ִ��id, ����ҽ��id) Values (v_·��ִ��id, t_Advice(I));
  End If;

  If ���˲���ids_In Is Not Null Then
    Select Column_Value Bulk Collect Into t_Doc From Table(f_Num2list(���˲���ids_In));
    Select Column_Value Bulk Collect Into t_File From Table(f_Num2list(�����ļ�ids_In));
    For I In 1 .. t_Doc.Count Loop
      v_����id := t_Doc(I);
    
      Insert Into ���Ӳ�����¼
        (ID, ������Դ, ����id, ��ҳid, Ӥ��, ����id, ��������, �ļ�id, ��������, ������, ����ʱ��, ������, ����ʱ��, ���汾, ǩ������, �༭��ʽ, ·��ִ��id)
        Select v_����id, 2, ����id_In, ��ҳid_In, Ӥ��_In, ����id_In, ����, ID, ����, �Ǽ���_In, �Ǽ�ʱ��_In, �Ǽ���_In, �Ǽ�ʱ��_In, 1, 0, 0,
               v_·��ִ��id
        From �����ļ��б�
        Where ID = t_File(I);
    
      v_������� := 0;
      For Rs In (Select ID, �ļ�id, Nvl(��id, 0) As ��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���, Ԥ�����id, �������, ʹ��ʱ��,
                        ����Ҫ��id, �滻��, Ҫ������, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ��
                 From �����ļ��ṹ
                 Where �ļ�id = t_File(I)
                 Order By �������) Loop
      
        v_������� := v_������� + 1;
        Select ���Ӳ�������_Id.Nextval Into v_Id From Dual;
      
        If Rs.��id = 0 Then
          v_��ǰ��id := v_Id;
          v_��id     := Null;
          If Rs.�������� = 1 And Not Rs.Ԥ�����id Is Null Then
            n_Ԥ�����id := Rs.Ԥ�����id;
          Else
            n_Ԥ�����id := Null;
          End If;
        Else
          --�������Ϊ�յ�ʱ�򣬸�ID�Ͳ��ǰ���˳����ˣ���Ҫ���²���
          If Rs.������� Is Null Then
            n_Ԥ�����id := Null;
            Select ������� Into v_ԭ������� From �����ļ��ṹ Where ID = Rs.��id;
            If v_ԭ������� Is Null Then
              v_��id := Null;
            Else
              Select ID Into v_��id From ���Ӳ������� Where �ļ�id = v_����id And ������� = v_ԭ�������;
            End If;
          Else
            v_��id := v_��ǰ��id;
          End If;
        End If;
      
        If Rs.�������� = 4 And Rs.�滻�� = 1 Then
          v_�����ı� := Zl_Replace_Element_Value(Rs.Ҫ������, ����id_In, ��ҳid_In, 2, Null, Ӥ��_In);
        Else
          v_�����ı� := Rs.�����ı�;
        End If;
      
        Insert Into ���Ӳ�������
          (ID, �ļ�id, ��ʼ��, ��ֹ��, ��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���, Ԥ�����id, �������, ʹ��ʱ��, ����Ҫ��id, �滻��,
           Ҫ������, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ��)
        Values
          (v_Id, v_����id, 1, 0, v_��id, v_�������, Rs.��������, Rs.������, Rs.��������, Null, Rs.�����д�, v_�����ı�, Rs.�Ƿ���, Rs.Ԥ�����id,
           Rs.�������, Rs.ʹ��ʱ��, Rs.����Ҫ��id, Rs.�滻��, Rs.Ҫ������, Rs.Ҫ������, Rs.Ҫ�س���, Rs.Ҫ��С��, Rs.Ҫ�ص�λ, Rs.Ҫ�ر�ʾ, Rs.������̬, Rs.Ҫ��ֵ��);
      
        If Rs.�������� = 5 Then
          Insert Into ���Ӳ���ͼ�� (����id, ͼ��) Values (v_Id, (Select ͼ�� From �����ļ�ͼ�� Where ����id = Rs.Id));
        End If;
      
      End Loop;
    
      Insert Into ���Ӳ�����ʽ
        (�ļ�id, ����)
      Values
        (v_����id, (Select ���� From �����ļ���ʽ Where �ļ�id = t_File(I)));
    End Loop;
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����·������_Insert;
/



Create Or Replace Procedure Zl_����·������_Delete(ִ�м�¼id_In ����·��ִ��.Id%Type) Is
  t_Id t_Numlist;

  --����ҽ��,�����׶δ���ʱ��ɾ��
  Cursor c_Advice Is
    Select ����ҽ��id
    From ����·��ҽ�� A
    Where ·��ִ��id = ִ�м�¼id_In And Not Exists
     (Select 1 From ����·��ҽ�� B Where a.����ҽ��id = b.����ҽ��id And a.·��ִ��id <> b.·��ִ��id);

  Cursor c_Doc Is
    Select ID From ���Ӳ�����¼ Where ·��ִ��id = ִ�м�¼id_In;

  v_�׶�id     ����·��ִ��.�׶�id%Type;
  v_·����¼id ����·��ִ��.·����¼id%Type;
  v_�Ǽ�ʱ��   ����·��ִ��.�Ǽ�ʱ��%Type;
  v_����       ����·��ִ��.����%Type;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --�Ƿ�����ȡ�����߼������ڽ�������м��
  Delete ����·��ҽ�� Where ·��ִ��id = ִ�м�¼id_In;
  Open c_Advice;
  Fetch c_Advice Bulk Collect
    Into t_Id;
  Close c_Advice;
  If t_Id.Count > 0 Then
    For I In 1 .. t_Id.Count Loop
      Zl_����ҽ����¼_Delete(t_Id(I), 0);
    End Loop;
  End If;

  Open c_Doc;
  Fetch c_Doc Bulk Collect
    Into t_Id;
  Close c_Doc;
  If t_Id.Count > 0 Then
    For I In 1 .. t_Id.Count Loop
      Zl_���Ӳ�����¼_Delete(t_Id(I));
    End Loop;
  End If;
  Delete ����·��ִ�� Where ID = ִ�м�¼id_In Returning ·����¼id, �׶�id Into v_·����¼id, v_�׶�id;

  Select Max(����) Into v_���� From ����·��ִ�� Where ·����¼id = v_·����¼id And �׶�id = v_�׶�id;
  --�����ǰ�׶ε����һ��ִ�м�¼��ɾ��(ȫ�����ǷǱ���ִ�е������)
  If v_���� Is Null Then
    --a.�����ǰû���κ�ִ�м�¼
    Select Max(����) Into v_���� From ����·��ִ�� Where ·����¼id = v_·����¼id;
    If v_���� Is Null Then
      Update �����ٴ�·�� Set ǰһ�׶�id = Null, ��ǰ�׶�id = Null, ��ǰ���� = Null, ״̬ = 1 Where ID = v_·����¼id;
    Else
      --b.���˵�ǰһ���׶�
      Select Max(�׶�id)
      Into v_�׶�id
      From ����·��ִ��
      Where ·����¼id = v_·����¼id And
            �Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��)
                    From ����·��ִ��
                    Where ·����¼id = v_·����¼id And �׶�id <> (Select ǰһ�׶�id From �����ٴ�·�� Where ID = v_·����¼id));
    
      Update �����ٴ�·��
      Set ��ǰ�׶�id = ǰһ�׶�id, ǰһ�׶�id = v_�׶�id, ��ǰ���� = v_����, ״̬ = 1
      Where ID = v_·����¼id;
    End If;
  Else
    Update �����ٴ�·�� Set ��ǰ���� = v_���� Where ID = v_·����¼id And ��ǰ���� <> v_����;
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����·������_Delete;
/


Create Or Replace Procedure Zl_����·��ִ��_Update
(
  ִ����_In   ����·��ִ��.ִ����%Type,
  ִ��ʱ��_In ����·��ִ��.ִ��ʱ��%Type,
  ִ������_In Varchar2 --ID|ִ�н��|ִ��˵��||...ĩβ��||,��ִ��˵��Ϊ��ʱ,Ҫ�ӿո�,�����||ճ��
) Is
  v_Str   Varchar2(4000);
  v_Tmp   Varchar2(1000);
  v_Index Number;
  I       Number(5) := 1;

  v_Id       t_Numlist := t_Numlist();
  v_ִ�н�� t_Strlist := t_Strlist();
  v_ִ��˵�� t_Strlist := t_Strlist();

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  v_Str := ִ������_In;
  Loop
    v_Index := Instr(v_Str, '||');
    Exit When(Nvl(v_Index, 0) = 0);
    v_Id.Extend;
    v_ִ�н��.Extend;
    v_ִ��˵��.Extend;
  
    v_Tmp := Substr(v_Str, 1, v_Index - 1);
    v_Id(I) := Substr(v_Tmp, 1, Instr(v_Tmp, '|') - 1);
    v_Tmp := Substr(v_Tmp, Instr(v_Tmp, '|') + 1);
    v_ִ�н��(I) := Trim(Substr(v_Tmp, 1, Instr(v_Tmp, '|') - 1));
    v_Tmp := Substr(v_Tmp, Instr(v_Tmp, '|') + 1);
    v_ִ��˵��(I) := Trim(v_Tmp);
  
    v_Str := Substr(v_Str, v_Index + 2);
    I     := I + 1;
  End Loop;

  Forall I In 1 .. v_Id.Count
    Update ����·��ִ��
    Set ִ���� = ִ����_In, ִ��ʱ�� = ִ��ʱ��_In, ִ�н�� = v_ִ�н��(I), ִ��˵�� = v_ִ��˵��(I)
    Where ID = v_Id(I);

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����·��ִ��_Update;
/

Create Or Replace Procedure Zl_����·��ִ��_Delete(·��ִ��id_In ����·��ִ��.Id%Type) Is
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin

  Update ����·��ִ�� Set ִ���� = Null, ִ��ʱ�� = Null, ִ�н�� = Null, ִ��˵�� = Null Where ID = ·��ִ��id_In;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����·��ִ��_Delete;
/


Create Or Replace Procedure Zl_����·������_Update(·����¼id_In �����ٴ�·��.Id%Type) Is
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin

  Update �����ٴ�·��
  Set ����ʱ�� = Sysdate, ״̬ = 2, ǰһ�׶�id = ��ǰ�׶�id, ��ǰ�׶�id = Null, ��ǰ���� = Null
  Where ID = ·����¼id_In;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����·������_Update;
/


Create Or Replace Procedure Zl_����·������_Delete
(
  ·����¼id_In �����ٴ�·��.Id%Type,
  ��������_In   �����ٴ�·��.״̬%Type
) Is
  v_�׶�id     ����·������.�׶�id%Type;
  v_ǰһ�׶�id ����·������.�׶�id%Type;
  v_����       ����·������.����%Type;
  v_����       ����·������.����%Type;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin

  Select ǰһ�׶�id Into v_�׶�id From �����ٴ�·�� Where ID = ·����¼id_In;
  Select Max(����), Max(����)
  Into v_����, v_����
  From ����·��ִ��
  Where ·����¼id = ·����¼id_In And �׶�id = v_�׶�id;

  If ��������_In = 3 Then
    --�������Ϊ����ʱ�Զ�������,ȡ�������Զ�ȡ������
    Delete ����·������ Where ·����¼id = ·����¼id_In And �׶�id = v_�׶�id And ���� = v_����;
    Delete ����·��ָ�� Where ·����¼id = ·����¼id_In And �׶�id = v_�׶�id And ���� = v_����;
  End If;

  --b.���˵�ǰһ���׶�
  Select Max(�׶�id)
  Into v_ǰһ�׶�id
  From ����·��ִ��
  Where ·����¼id = ·����¼id_In And
        �Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ����·��ִ�� Where ·����¼id = ·����¼id_In And �׶�id <> v_�׶�id);

  Update �����ٴ�·��
  Set ����ʱ�� = Null, ״̬ = 1, ǰһ�׶�id = v_ǰһ�׶�id, ��ǰ�׶�id = v_�׶�id, ��ǰ���� = v_����
  Where ID = ·����¼id_In;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����·������_Delete;
/

Create Or Replace Function Zl_Getpathcharge
(
  ����id_In   ������ҳ.����id%Type, --��ȷ������ʱ����0
  ��ҳid_In   ������ҳ.��ҳid%Type, --��ȷ������ʱ����0
  ·��id_In   �ٴ�·����Ŀ.·��id%Type,
  �汾��_In   �ٴ�·����Ŀ.�汾��%Type,
  �׶�id_In   �ٴ�·����Ŀ.�׶�id%Type, --û��ָ���׶�ʱ,���ݵ�ǰ������ȷ��ȱʡ�Ľ׶�
  ����_In     ����·��ִ��.����%Type, --��ǰ�׶�����ִ�е�����
  ��Ժʱ��_In Date --������Ժ�����ʱ��,���ڼ����ϴ�ִ��ʱ��(Ƶ��Ϊÿn��m��ʱ),��ȷ������ʱ���뵱ǰϵͳʱ��
) Return Number As
  v_Error Varchar2(255);
  Err_Custom Exception;

  v_Tmp Varchar2(1000);
  n_Tmp Number(8);

  v_�ѱ�         ������ҳ.�ѱ�%Type;
  n_ʵ�ս��     Number(16, 5);
  n_Ӧ�ս��     Number(16, 5);
  n_������     Number(16, 5);
  n_ʵ�պϼ�     Number(16, 5);
  n_�Ʒ�����     Number(16, 5);
  n_����         Number(16, 5);
  n_���ܼ����ۿ� Number(1);
  n_������id     Number(8);
  n_����         Number(8);
  n_Day          Number(8); --��ǰ������������
  n_Lastday      Number(8);

  l_�ɼ����� Boolean;
  l_��ҩ�巨 Boolean;
  l_��ҩ�÷� Boolean;
  l_��ҩ;�� Boolean;
  l_��Ѫ;�� Boolean;

  v_Lasttype   ������ĿĿ¼.���%Type;
  n_Lastsum    ·��ҽ������.�ܸ�����%Type;
  n_Last���id ·��ҽ������.���id%Type;
  n_Lastid     ·��ҽ������.Id%Type;
  n_Lastamount Number(8);
  n_Last����   Number(8);
  l_Last�巨   Boolean;
  l_Do         Boolean;
  l_Firstday   Boolean;

  n_�׶�id     �ٴ�·���׶�.Id%Type;
  n_ǰһ�׶�id �ٴ�·���׶�.Id%Type;
  l_Rate       t_Strlist;

  --ȡҩƷ�����Ϣ(δ��ȷ���ʱ,ȡ����һ�����)
  Cursor Mediinfo
  (
    ������Ŀid_In Number,
    �շ�ϸĿid_In Number
  ) Is
    Select g.Id As �շ�ϸĿid, Nvl(f.����ϵ��, 1) As ����ϵ��, f.�ɷ����, Nvl(g.�Ƿ���, 0) �Ƿ���, h.ȱʡ�۸�, h.�ּ�, g.���ηѱ�, h.������Ŀid,
           Nvl(h.�����շ���, 1) �����շ���
    From ҩƷ��� F, �շ���ĿĿ¼ G, �շѼ�Ŀ H
    Where f.ҩ��id = ������Ŀid_In And f.ҩƷid = Nvl(�շ�ϸĿid_In, f.ҩƷid) And f.ҩƷid = g.Id And g.Id = h.�շ�ϸĿid And
          Sysdate Between h.ִ������ And Nvl(h.��ֹ����, Sysdate + 1)
    Order By g.����;
  r_Medi Mediinfo%Rowtype;

  --����:��ȡָ����������ȱʡʱ��׶�ID  
  Function Getphaseid(n_Day Number) Return Number As
    n_Id �ٴ�·���׶�.Id%Type;
  Begin
    For R In (Select ID
              From �ٴ�·���׶�
              Where n_Day Between Nvl(��ʼ����, n_Day) And Nvl(��������, Nvl(��ʼ����, n_Day)) And ·��id = ·��id_In And �汾�� = �汾��_In
              Order By Decode(��id, Null, 0, 1), ���) Loop
      n_Id := r.Id;
      Exit;
    End Loop;
    Return n_Id;
  End Getphaseid;

  --����:��ȡָʱ��׶�(�����Ƿ�֧)��ǰһʱ��׶�id 
  Function Getprephaseid(n_�׶�id �ٴ�·���׶�.Id%Type) Return Number As
    n_Id �ٴ�·���׶�.Id%Type;
  Begin
    Select Nvl(Max(b.Id), 0)
    Into n_Id
    From �ٴ�·���׶� A, �ٴ�·���׶� B
    Where a.·��id = b.·��id And a.�汾�� = b.�汾�� And a.Id = n_�׶�id And b.��� = a.��� - 1;
    Return n_Id;
  End Getprephaseid;

  --����:��ȡָ��·����Ŀ�Ŀ�ʼִ������(��Ժʱ��Ϊ��һ��)
  Function Getitembeginday(n_·����Ŀid �ٴ�·����Ŀ.Id%Type) Return Number As
    n_Preday Number(8);
    n_Id     �ٴ�·���׶�.Id%Type;
    n_Preid  �ٴ�·���׶�.Id%Type;
    n_Tmp    Number(8);
    n_Return Number(8);
  Begin
    n_Preday := ����_In - 1;
    If n_Preday = 0 Or n_ǰһ�׶�id = 0 Then
      --��ǰ�ǵ�һ����һ���׶�
      n_Return := 1;
    Else
      n_Id    := n_�׶�id;
      n_Preid := n_ǰһ�׶�id;
      Loop
        --���ǰһ�׶��Ƿ�����ͬ��·����Ŀ
        Select Nvl(Count(p.Id), 0)
        Into n_Tmp
        From �ٴ�·����Ŀ T, �ٴ�·����Ŀ P
        Where t.·��id = p.·��id And t.�汾�� = p.�汾�� And t.��Ŀ���� = p.��Ŀ���� And t.Id = n_·����Ŀid And p.�׶�id = n_Preid;
        If n_Tmp = 0 Then
          Exit;
        End If;
      
        n_Id := n_Preid; --�����,������ǰ��
        Select Nvl(Max(b.Id), 0)
        Into n_Preid
        From �ٴ�·���׶� A, �ٴ�·���׶� B
        Where a.·��id = b.·��id And a.�汾�� = b.�汾�� And a.Id = n_Id And b.��� = a.��� - 1;
        If n_Preid = 0 Then
          Exit;
        End If;
      End Loop;
    
      Select Nvl(��ʼ����, 0) Into n_Tmp From �ٴ�·���׶� Where ID = n_Id;
      If n_Tmp = 0 Then
        --�����ڼ�������׶β���������,����ȡǰһ���׶εĽ�������+1        
        Select Nvl(Nvl(��������, ��ʼ����), 0) + 1 Into n_Tmp From �ٴ�·���׶� Where ID = n_Preid;
      End If;
      If n_Tmp <= 1 Then
        n_Return := 1;
      Else
        n_Return := n_Tmp;
      End If;
    End If;
    Return n_Return;
  End Getitembeginday;

  --����:��ȡʱ��ҩƷ��Ӧ�ս��(��Ϊ�ǹ���,���ܳ���ģʽ  )
  Function Getʱ��ҩƷ���
  (
    n_������     In Number,
    n_ִ�п���id In Number,
    n_�շ�ϸĿid In Number
  ) Return Number As
    n_�ܽ��   Number(16, 5);
    n_�������� Number(16, 5);
    n_�������� Number(16, 5);
  Begin
    n_�ܽ��   := 0;
    n_�������� := n_������;
    For D In (Select Nvl(��������, 0) As ���, Nvl(���ۼ�, Nvl(Decode(Nvl(ʵ������, 0), 0, 0, ʵ�ʽ�� / ʵ������), 0)) As ʱ��
              From ҩƷ���
              Where �ⷿid = n_ִ�п���id And ҩƷid = n_�շ�ϸĿid And Nvl(��������, 0) > 0 And ���� = 1 And
                    (Nvl(����, 0) = 0 Or Ч�� Is Null Or Ч�� > Trunc(Sysdate))
              Order By Nvl(����, 0)) Loop
      If n_�������� <= d.��� Then
        n_�������� := n_��������;
      Else
        n_�������� := d.���;
      End If;
      n_�ܽ��   := n_�ܽ�� + n_�������� * d.ʱ��;
      n_�������� := n_�������� - n_��������;
      If n_�������� = 0 Then
        Exit;
      End If;
    End Loop;
    Return n_�ܽ��;
  End Getʱ��ҩƷ���;
Begin
  Select To_Char(Sysdate, 'D') - 1 Into n_Day From Dual;
  If n_Day = 0 Then
    n_Day := 7;
  End If;

  If Nvl(�׶�id_In, 0) = 0 Then
    n_�׶�id := Getphaseid(����_In);
    n_Tmp    := n_�׶�id;
  Else
    n_�׶�id := �׶�id_In; --�����ǰ�Ƿ�֧,����ȱʡ��֧    
    Select Nvl(��id, ID) Into n_Tmp From �ٴ�·���׶� Where ID = n_�׶�id;
  End If;
  n_ǰһ�׶�id := Getprephaseid(n_Tmp);

  l_Firstday := False;
  Select Nvl(��ʼ����, 0) Into n_Tmp From �ٴ�·���׶� Where ID = n_�׶�id;
  If n_Tmp = 0 Then
    --�����ڼ�������׶β���������,����ȡǰһ���׶εĽ�������+1 
    Select Nvl(Nvl(��������, ��ʼ����), 0) + 1 Into n_Tmp From �ٴ�·���׶� Where ID = n_ǰһ�׶�id;
  End If;
  If n_Tmp = ����_In Then
    l_Firstday := True;
  End If;

  If Nvl(����id_In, 0) <> 0 Then
    Select �ѱ� Into v_�ѱ� From ������ҳ Where ����id = ����id_In And ��ҳid = ��ҳid_In;
  End If;
  n_���ܼ����ۿ� := Nvl(Zl_Getsysparameter(93), 0);
  n_ʵ�պϼ�     := 0;

  --Ժ��ִ�к���ִ�еĶ�������,���Ƽ۵ĳ���
  For R In (Select Nvl(c.���id, c.Id) As ��id, Nvl(e.���, c.���) As ���, c.���, a.Id, c.���id, c.��Ч, d.���, c.������Ŀid, c.�շ�ϸĿid,
                   c.ִ�п���id, c.�걾��λ, c.��鷽��, Nvl(c.��������, 1) As ��������, c.�ܸ�����, c.ִ��Ƶ��, c.Ƶ�ʴ���, c.Ƶ�ʼ��, c.�����λ, c.ִ������,
                   c.ʱ�䷽��, Nvl(d.�������, 0) As �������, a.ִ�з�ʽ
            From �ٴ�·����Ŀ A, �ٴ�·��ҽ�� B, ·��ҽ������ C, ������ĿĿ¼ D, ·��ҽ������ E
            Where a.·��id = ·��id_In And a.�汾�� = �汾��_In And a.�׶�id = n_�׶�id And a.ִ�з�ʽ Not In (0, 3) And a.Id = b.·����Ŀid And
                  b.ҽ������id = c.Id And c.������Ŀid = d.Id And c.ִ������ Not In (0, 5) And d.�Ƽ����� <> 1 And c.���id = e.Id(+)
            Order By a.��Ŀ���, ���, ��id, c.���) Loop
    l_Do := True;
    If r.ִ�з�ʽ = 2 Then
      l_Do := l_Firstday; --����ִ��һ�ε���Ŀ,���ڱ��׶εĵ�һ��ʱ����
    End If;
    If l_Do Then
      --1.��������
      n_����     := 0;
      l_��ҩ�巨 := False;
      l_��Ѫ;�� := False;
      l_��ҩ�÷� := False;
      l_�ɼ����� := False;
      l_��ҩ;�� := False;
      If r.��� = 'E' And r.���id Is Not Null And r.���id = n_Last���id Then
        If v_Lasttype = '7' Then
          l_��ҩ�巨 := True;
        Elsif v_Lasttype = 'K' Then
          l_��Ѫ;�� := True;
        End If;
      Elsif r.��� = 'E' And r.���id Is Null And r.Id = n_Last���id Then
        If l_Last�巨 Or v_Lasttype = '7' Then
          l_��ҩ�÷� := True;
        Elsif v_Lasttype = 'C' Then
          l_�ɼ����� := True;
        Elsif v_Lasttype In ('5', '6') Then
          l_��ҩ;�� := True;
        End If;
      End If;
      If r.��� In ('5', '6', '7') Then
        Open Mediinfo(r.������Ŀid, r.�շ�ϸĿid);
        Fetch Mediinfo
          Into r_Medi;
        Close Mediinfo;
      End If;
    
      --���� 
      If r.��Ч = 0 Then
        --a.��Ҫҽ����һ���ɼ��ļ�����Ŀ,��ҩƷ(��Ϊ�ǹ���,�����Ƕ�̬�����������Ӱ��)
        If (r.���id Is Null And Not l_�ɼ����� And Not l_��ҩ�÷� And Not l_��ҩ;��) Or (r.���id Is Not Null And r.��� = 'C') Or
           r.��� In ('5', '6', '7') Then
        
          If r.ʱ�䷽�� Is Null And (Nvl(r.Ƶ�ʴ���, 0) = 0 Or Nvl(r.Ƶ�ʼ��, 0) = 0 Or r.�����λ Is Null) Then
            n_���� := 1; --��������Ŀ --��Ϊ�ǹ���,��Ϊ��������ֹʱ��,��ÿ��һ����
          Else
            Select Column_Value Bulk Collect Into l_Rate From Table(f_Str2list(r.ʱ�䷽��, '-')); --ִ��Ƶ��Ϊ"��ѡƵ��"����Ŀ             
            Case r.�����λ
              When '��' Then
                --��:ÿ�����Σ� 1/8:00-3/8:00-5/8:00��1/8-3/8-5/8                  
                For I In 1 .. l_Rate.Count Loop
                  If n_Day = Substr(l_Rate(I), 1, Instr(l_Rate(I), '/') - 1) Then
                    n_���� := n_���� + 1;
                  End If;
                End Loop;
              When '��' Then
                --��:ÿ�����Σ�8:00-12:00-16:00 �� 8:12:16                     
                If r.Ƶ�ʼ�� = 1 Then
                  If ����_In = 1 Then
                    For I In 1 .. l_Rate.Count Loop
                      If ��Ժʱ��_In <= To_Date(To_Char(��Ժʱ��_In, 'yyyy-mm-dd') || ' ' || l_Rate(I), 'yyyy-mm-dd hh24:mi') Then
                        n_���� := n_���� + 1; --��Ժ���� 
                      End If;
                    End Loop;
                  Else
                    n_���� := r.Ƶ�ʴ���;
                  End If;
                Else
                  n_Lastday := ����_In - Getitembeginday(r.Id); --��:����һ�Σ� 1/8 �� 1/8:00                
                  For I In 1 .. l_Rate.Count Loop
                    If n_Lastday = Substr(l_Rate(I), 1, Instr(l_Rate(I), '/') - 1) Then
                      n_���� := n_���� + 1;
                    End If;
                  End Loop;
                End If;
              When 'Сʱ' Then
                If ����_In = 1 Then
                  n_���� := Trunc(Trunc((Trunc(��Ժʱ��_In + 1) - ��Ժʱ��_In) * 24) / r.Ƶ�ʼ��); --��Ժ����          
                Else
                  n_���� := Trunc(24 / r.Ƶ�ʼ��);
                End If;
              When '����' Then
                If ����_In = 1 Then
                  n_���� := Trunc(Trunc((Trunc(��Ժʱ��_In + 1) - ��Ժʱ��_In) * 24 * 60) / r.Ƶ�ʼ��); --��Ժ����     
                Else
                  n_���� := Trunc((24 * 60) / r.Ƶ�ʼ��);
                End If;
            End Case;
          End If;
          If r.��� = '7' Then
            n_���� := r.�ܸ����� * n_����;
            If r_Medi.�ɷ���� = 0 Then
              n_���� := n_���� * r.�������� / r_Medi.����ϵ��;
            Else
              n_���� := n_���� * Ceil(r.�������� / r_Medi.����ϵ��); --�ܸ�����=����    
            End If;
            n_Last���� := n_����; --��ҩ�巨���÷����ܸ�����Ϊ���� 
          Elsif r.��� = '5' Or r.��� = '6' Then
            If r_Medi.�ɷ���� = 0 Then
              n_���� := n_���� * r.�������� / r_Medi.����ϵ��; --�ɷ���
            Elsif r_Medi.�ɷ���� = 1 Then
              n_���� := Ceil(n_���� * r.�������� / r_Medi.����ϵ��); --������
            Elsif r_Medi.�ɷ���� = 2 Then
              n_���� := n_���� * Ceil(r.�������� / r_Medi.����ϵ��); --һ����
            Else
              n_���� := Ceil(n_���� * r.�������� / r_Medi.����ϵ��); --�ɷ����<0  :n���ڷ���ʹ����Ч,����̫����,�����ɷ��㴦��
            End If;
          Else
            If r.������� = 1 Then
              n_���� := Ceil(r.�������� * n_����); --ȡ�������--ȡ�����㣬�����ڿ�ѡƵ�ʵļ�������ʱ���ƴγ���ҽ����
            Else
              n_���� := r.�������� * n_����;
            End If;
          End If;
        Elsif l_��ҩ�巨 Or l_��ҩ�÷� Then
          n_���� := n_Last����; --b.��ҩ�巨���÷�Ϊ���� 
          n_���� := n_Lastamount;
        Elsif l_��ҩ;�� Then
          n_���� := n_Lastamount; --c.��ҩ;��     
          n_���� := n_Lastamount;
        Elsif l_��Ѫ;�� Then
          n_���� := n_Lastamount; --d.��Ѫ;����ִ�д���      
          n_���� := n_Lastamount;
        Elsif r.���id Is Not Null Or l_�ɼ����� Then
          n_���� := n_Lastsum; --e.����ҽ����걾�ɼ�����(�����Ϻ�������ϲ�����Ϊ����,���Դ˶β���ִ��)   
          n_���� := n_Lastamount;
        End If;
      Else
        --����
        If r.��� = '7' Then
          n_���� := r.�ܸ�����;
          If r_Medi.�ɷ���� = 0 Then
            n_���� := r.�ܸ����� * r.�������� / r_Medi.����ϵ��;
          Else
            n_���� := r.�ܸ����� * Ceil(r.�������� / r_Medi.����ϵ��); --�ܸ�����=����    
          End If;
        Elsif r.��� In ('5', '6') Then
          If Nvl(r.Ƶ�ʴ���, 0) = 0 Or Nvl(r.Ƶ�ʼ��, 0) = 0 Then
            n_���� := 1; --һ���Ե�����ҩƷ
            --��Ϊû������,���Բ���Ƶ�ʼ���
          Elsif r_Medi.�ɷ���� = 0 And Nvl(r.��������, 0) <> 0 Then
            --�ɷ���ҩƷʱ,�������Ե����ı��������ҩ;���Ĵ���,����һ��Ƶ�����ڵĴ�������
            n_���� := Trunc(r.�ܸ����� * r_Medi.����ϵ�� / r.��������);
          Else
            n_���� := r.Ƶ�ʴ���;
          End If;
          n_���� := r.�ܸ�����;
        Elsif l_��ҩ�巨 Or l_��ҩ�÷� Or l_��ҩ;�� Then
          n_���� := n_Lastamount; --��ҩ;��,��ҩ�÷�,�巨�Ĵ���
          n_���� := n_Lastamount;
        Elsif (r.���id Is Null And Not l_�ɼ�����) Or (r.���id Is Not Null And r.��� = 'C') Then
          --��Ҫҽ����һ���ɼ��ļ�����Ŀ
          n_���� := Nvl(r.�ܸ�����, 1);
          n_���� := Ceil(n_���� / r.��������);
        Elsif l_��Ѫ;�� Then
          n_���� := n_Lastamount; --d.��Ѫ;����ִ�д���  
          n_���� := n_Lastamount;
        Elsif r.���id Is Not Null Or l_�ɼ����� Then
          n_���� := n_Lastsum; --e.����ҽ����걾�ɼ�����    
          n_���� := n_Lastamount;
        End If;
      End If;
      n_Lastamount := n_����;
      n_Lastsum    := n_����;
      v_Lasttype   := r.���;
      n_Last���id := r.���id;
      n_Lastid     := r.Id;
      l_Last�巨   := l_��ҩ�巨;
    
      --2.����ʵ�ս��(�����ǼӰ�Ӽ�)    
      n_ʵ�ս�� := 0;
      n_Ӧ�ս�� := 0;
      n_������id := 0;
      n_������ := 0;
      If r.��� In ('4', '5', '6', '7') Then
      
        If r_Medi.�Ƿ��� = 0 Then
          n_Ӧ�ս�� := r_Medi.�ּ� * n_����;
        Else
          n_Ӧ�ս�� := Getʱ��ҩƷ���(n_����, r.ִ�п���id, r_Medi.�շ�ϸĿid);
        End If;
        If Not (v_�ѱ� Is Null Or r_Medi.���ηѱ� = 1) Then
          v_Tmp      := Zl_Actualmoney(v_�ѱ�, r_Medi.�շ�ϸĿid, r_Medi.������Ŀid, n_Ӧ�ս��, n_����, r.ִ�п���id);
          n_ʵ�ս�� := Substr(v_Tmp, Instr(v_Tmp, ':') + 1);
        Else
          n_ʵ�ս�� := n_Ӧ�ս��;
        End If;
        n_ʵ�պϼ� := n_ʵ�պϼ� + n_ʵ�ս��;
      
      Else
        For D In (Select c.���, c.Id As �շ�ϸĿid, a.�շ�����, b.������Ŀid, Decode(c.�Ƿ���, 1, b.ȱʡ�۸�, b.�ּ�) As ����, c.�Ƿ���,
                         Nvl(a.������Ŀ, 0) As ����, d.��������, c.���ηѱ�, Nvl(a.��������, 0) As ��������, Nvl(a.�շѷ�ʽ, 0) As �շѷ�ʽ, b.�����շ���
                  From �����շѹ�ϵ A, �շѼ�Ŀ B, �շ���ĿĿ¼ C, �������� D
                  Where a.������Ŀid = r.������Ŀid And (r.��� <> 'D' Or r.��� = 'D' And a.��鲿λ = r.�걾��λ And a.��鷽�� = r.��鷽��) And
                        a.�շ���Ŀid = b.�շ�ϸĿid And a.�շ���Ŀid = c.Id And a.�շ���Ŀid = d.����id(+) And c.������� In (2, 3) And
                        (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null) And Sysdate Between b.ִ������ And
                        Nvl(b.��ֹ����, Sysdate + 1) And
                        (a.�շѷ�ʽ = 1 And c.��� = '4' And a.�շ���Ŀid = r_Medi.�շ�ϸĿid Or Not (a.�շѷ�ʽ = 1 And c.��� = '4'))
                  Order By ��������, ����, a.�շ���Ŀid) Loop
          n_�Ʒ����� := n_���� * d.�շ�����;
        
          If d.�Ƿ��� = 1 And (d.��� In ('5', '6', '7') Or (d.��� = '4' And d.�������� = 1)) Then
            n_Ӧ�ս�� := Getʱ��ҩƷ���(n_�Ʒ�����, r.ִ�п���id, d.�շ�ϸĿid); --ʱ�۷�ҩ��ҩƷ��������õ�����    
          Elsif r.��� = 'F' And r.���id Is Not Null Then
            n_Ӧ�ս�� := d.���� * Nvl(d.�����շ���, 100) / 100 * n_�Ʒ�����;
          Else
            n_Ӧ�ս�� := d.���� * n_�Ʒ�����;
          End If;
          If n_Ӧ�ս�� <> 0 Then
            If d.���� = 0 And n_���ܼ����ۿ� = 1 And n_������id = 0 Then
              n_������id := d.������Ŀid; --SQL����������ǰ��,ֻȡ����Ŀ�ĵ�һ������
            End If;
          
            If n_������id <> 0 Then
              n_������ := n_������ + n_Ӧ�ս��;
              n_ʵ�ս�� := n_Ӧ�ս��;
            Elsif v_�ѱ� Is Null Or d.���ηѱ� = 1 Then
              n_ʵ�ս�� := n_Ӧ�ս��;
            Else
              v_Tmp      := Zl_Actualmoney(v_�ѱ�, d.�շ�ϸĿid, d.������Ŀid, n_Ӧ�ս��, n_�Ʒ�����, r.ִ�п���id);
              n_ʵ�ս�� := Substr(v_Tmp, Instr(v_Tmp, ':') + 1);
            End If;
            n_ʵ�պϼ� := n_ʵ�պϼ� + n_ʵ�ս��;
          End If;
        End Loop;
        If n_������id <> 0 And n_������ <> 0 Then
          v_Tmp      := Zl_Actualmoney(v_�ѱ�, 0, n_������id, n_������);
          n_ʵ�ս�� := Substr(v_Tmp, Instr(v_Tmp, ':') + 1);
          n_ʵ�պϼ� := n_ʵ�պϼ� + n_ʵ�ս��;
        End If;
      End If;
    End If;
  End Loop;
  Return n_ʵ�պϼ�;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Getpathcharge;
/


Create Or Replace Procedure Zl_���˱䶯��¼_Change
(
  ����id_In     ������ҳ.����id%Type,
  ��ҳid_In     ������ҳ.��ҳid%Type,
  ת�����id_In ���˱䶯��¼.����id%Type,
  ����Ա���_In ���˱䶯��¼.����Ա���%Type,
  ����Ա����_In ���˱䶯��¼.����Ա����%Type
) As
  -----------------------------------------------------------
  --˵��������ת�ƵǼ�
  -----------------------------------------------------------    
  v_Count Number;
  v_����  ������Ϣ.����%Type;
  Err_Custom Exception;
  v_Error Varchar2(255);
Begin
  --�����жϸò����Ƿ��ڵȴ�ת�ƻ����״̬
  Select Count(*)
  Into v_Count
  From ������ҳ
  Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��Ժ���� Is Null And Nvl(״̬, 0) In (0, 3);
  If v_Count = 0 Then
    v_Error := '����ʧ��,�ò���������ת��״̬����δ���,����ת�ơ�';
    Raise Err_Custom;
  End If;
  
   --�ٴ�·������ִ��ʱ������ת��
  Select Max(b.״̬)
  Into v_Count
  From ������ҳ A, �����ٴ�·�� B
  Where a.����id = ����id_In And a.��ҳid = ��ҳid_In And a.����id = b.����id And a.��ҳid = b.��ҳid And a.��Ժ����id = b.����id;
  If v_Count = 1 Then
    v_Error := '�ò��˵��ٴ�·������ִ����,����ת�ơ�';
    Raise Err_Custom;
  End If;

  --����д��ʼʱ�����ֹʱ��
  Insert Into ���˱䶯��¼
    (ID, ����id, ��ҳid, ��ʼʱ��, ��ʼԭ��, ����id, ����id, ����Ա���, ����Ա����)
  Values
    (���˱䶯��¼_Id.Nextval, ����id_In, ��ҳid_In, Null, 3, Null, ת�����id_In, ����Ա���_In, ����Ա����_In);

  Update ������ҳ
  Set ״̬ = 2, ���� = Zl_Age_Calc(����id)
  Where ����id = ����id_In And ��ҳid = ��ҳid_In Returning ���� Into v_����;

  Update ������Ϣ Set ���� = v_���� Where ����id = ����id_In;

  --�����������
  Select Count(*)
  Into v_Count
  From ���˱䶯��¼
  Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ʼʱ�� Is Null And ��ֹʱ�� Is Null;
  If v_Count > 1 Then
    v_Error := '���ֲ��˴��ڷǷ��ı䶯��¼,��ǰ�������ܼ�����' || Chr(13) || Chr(10) ||
               '��������������粢�����������,��ˢ�²���״̬�����ԣ�';
    Raise Err_Custom;
  End If;

  Select Count(*) Into v_Count From ������ҳ Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��Ժ���� Is Null;
  If v_Count = 0 Then
    v_Error := '����ʧ��,�ò����ѳ�Ժ,���ܽ��е�ǰ����.' || Chr(13) || Chr(10) ||
               '��������������粢�����������,��ˢ�²���״̬��';
    Raise Err_Custom;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_���˱䶯��¼_Change;
/

Create Or Replace Procedure Zl_���˱䶯��¼_Out
(
  ����id_In       ������ҳ.����id%Type,
  ��ҳid_In       ������ҳ.��ҳid%Type,
  ����id_In       ������ϼ�¼.����id%Type,
  ���id_In       ������ϼ�¼.���id%Type,
  ��Ժ���_In     ������ϼ�¼.�������%Type,
  ��Ժ���_In     ������ϼ�¼.��Ժ���%Type,
  ��ҽ����id_In   ������ϼ�¼.����id%Type,
  ��ҽ���id_In   ������ϼ�¼.���id%Type,
  ��ҽ���_In     ������ϼ�¼.�������%Type,
  ��ҽ��Ժ���_In ������ϼ�¼.��Ժ���%Type,
  �Ƿ�����_In     ������ҳ.�Ƿ�ȷ��%Type, --ͬʱ��Ϊ��ҽ���Ƿ�����
  ��Ժ��ʽ_In     ������ҳ.��Ժ��ʽ%Type,
  ��Ժʱ��_In     ������ҳ.��Ժ����%Type,
  �����־_In     ������ҳ.�����־%Type, --0/NULL-�����1-�£�2-�꣬3-�ܣ�4-�죬9-����
  ��������_In     ������ҳ.��������%Type,
  ʬ���־_In     ������ҳ.ʬ���־%Type,
  ����Ա���_In   ���˱䶯��¼.����Ա���%Type,
  ����Ա����_In   ���˱䶯��¼.����Ա����%Type
) As
  -----------------------------------------------------------
  --˵�������˳�Ժ
  -----------------------------------------------------------
  Cursor c_Bedinfo Is
    Select ����id, ���� From ��λ״����¼ Where ����id = ����id_In;

  v_Count Number;
  Err_Custom Exception;
  v_Error    Varchar2(255);
  v_�������� Date;
  v_Ӧ��ʱ�� Date;
  v_�����   zlSystems.�����%Type;
  v_����     ������Ϣ.����%Type;
  v_Sql      Varchar2(1000);
  v_��Ժ���� Number;
Begin
  --�����жϸò����Ƿ��ѳ�Ժ
  Select Count(*) Into v_Count From ������ҳ Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��Ժ���� Is Null;

  If v_Count = 0 Then
    v_Error := '����ʧ��,�ò��˿����Ѿ���Ժ��';
    Raise Err_Custom;
  End If;

  --�����жϸò����Ƿ��ڵȴ�ת�ƻ����״̬
  Select Count(*)
  Into v_Count
  From ������ҳ
  Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��Ժ���� Is Null And Nvl(״̬, 0) In (0, 3);
  If v_Count = 0 Then
    v_Error := '����ʧ��,�ò���������ת��״̬����δ���,���ܳ�Ժ��';
    Raise Err_Custom;
  End If;
  
   --�ٴ�·������ִ��ʱ�������Ժ
  Select Max(b.״̬)
  Into v_Count
  From ������ҳ A, �����ٴ�·�� B
  Where a.����id = ����id_In And a.��ҳid = ��ҳid_In And a.����id = b.����id And a.��ҳid = b.��ҳid And a.��Ժ����id = b.����id;
  If v_Count = 1 Then
    v_Error := '�ò��˵��ٴ�·������ִ����,���ܳ�Ժ��';
    Raise Err_Custom;
  End If;

  --�ж��Ƿ����סԺ�ձ�
  Select Nvl(��Ժ����id, ��Ժ����id) Into v_��Ժ���� From ������ҳ Where ����id = ����id_In And ��ҳid = ��ҳid_In;
  Select Zl_סԺ�ձ�_Count(v_��Ժ����, ��Ժʱ��_In) Into v_Count From Dual;
  If v_Count > 0 Then
    v_Error := '�Ѳ���ҵ��ʱ���ڵ�סԺ�ձ�,���ܰ����ҵ��!';
    Raise Err_Custom;
  End If;

  --�ж��Ƿ��벡��ϵͳ����
  Begin
    Select ����� Into v_����� From zlSystems Where Floor(��� / 100) = 3;
  Exception
    When Others Then
      Null;
  End;
  --��Ժ�䶯
  Update ���˱䶯��¼
  Set ��ֹʱ�� = ��Ժʱ��_In, ��ֹԭ�� = 1, ��ֹ��Ա = ����Ա����_In
  Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null;

  --��λ��¼
  For r_Bedrow In c_Bedinfo Loop
    Update ��λ״����¼
    Set ״̬ = '�մ�', ����id = Null, ����id = Decode(����, 1, Null, ����id)
    Where ����id = r_Bedrow.����id And ���� = r_Bedrow.����;
  End Loop;

  --������ҳ
  Update ������ҳ
  Set ״̬ = 0, ��Ժ���� = ��Ժʱ��_In, ��Ժ��ʽ = ��Ժ��ʽ_In,
      סԺ���� = Decode(Trunc(��Ժʱ��_In) - Trunc(��Ժ����), 0, 1, Trunc(��Ժʱ��_In) - Trunc(��Ժ����)), �����־ = �����־_In,
      �������� = Decode(��������_In, 0, Null, ��������_In), ʬ���־ = ʬ���־_In, �Ƿ�ȷ�� = Decode(Nvl(�Ƿ�����_In, 0), 0, 1, 0),
      ���� = Zl_Age_Calc(����id), ����״̬ = Null
  Where ����id = ����id_In And ��ҳid = ��ҳid_In
  Returning ���� Into v_����;

  --���������¼
  If v_����� = 100 Then
    If Nvl(��������_In, 0) <> 0 Then
      If �����־_In = 1 Then
        v_�������� := Add_Months(��Ժʱ��_In, ��������_In);
      Elsif �����־_In = 2 Then
        v_�������� := Add_Months(��Ժʱ��_In, 12 * ��������_In);
      Elsif �����־_In = 3 Then
        v_�������� := ��Ժʱ��_In + 7 * ��������_In;
      Elsif �����־_In = 4 Then
        v_�������� := ��Ժʱ��_In + ��������_In;
      End If;
    Else
      v_�������� := To_Date('3000-1-1', 'YYYY-MM-DD');
    End If;
  
    If �����־_In = 1 Or �����־_In = 2 Or �����־_In = 3 Or �����־_In = 4 Then
      If v_�������� > Add_Months(��Ժʱ��_In, 3) Then
        v_Ӧ��ʱ�� := Trunc(Add_Months(��Ժʱ��_In, 3));
      Else
        v_Ӧ��ʱ�� := v_��������;
      End If;
      v_Sql := 'Insert Into �����¼ (ID, ����id, ��ҳid, ��������, Ӧ��ʱ��) Values (�����¼_Id.Nextval,' || ����id_In || ',' || ��ҳid_In ||
               ',TO_DATE(''' || To_Char(v_��������, 'YYYY-MM-DD') || ''',' || '''YYYY-MM-DD ''' || '),' || 'TO_DATE(''' ||
               To_Char(v_Ӧ��ʱ��, 'YYYY-MM-DD') || ''',' || '''YYYY-MM-DD ''' || '))';
      Execute Immediate v_Sql;
    End If;
  End If;
  --������Ϣ
  Update ������Ϣ
  Set ��ǰ����id = Null, ��ǰ����id = Null, ��ǰ���� = Null, ��Ժʱ�� = ��Ժʱ��_In, ���� = v_����, ��Ժ = Null
  Where ����id = ����id_In;

  --��Ժ���
  If ��Ժ���_In Is Not Null Or ����id_In Is Not Null Or ���id_In Is Not Null Then
    Delete From ������ϼ�¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ������� = 3 And ��ϴ��� = 1 And ��¼��Դ = 2;
    Insert Into ������ϼ�¼
      (ID, ����id, ��ҳid, ��¼��Դ, �������, ��ϴ���, ����id, ���id, �������, �Ƿ�����, ��Ժ���, ��¼����, ��¼��)
    Values
      (������ϼ�¼_Id.Nextval, ����id_In, ��ҳid_In, 2, 3, 1, ����id_In, ���id_In, ��Ժ���_In, �Ƿ�����_In, ��Ժ���_In, Sysdate, ����Ա����_In);
  End If;

  --��ҽ��Ժ���
  If ��ҽ���_In Is Not Null Or ��ҽ����id_In Is Not Null Or ��ҽ���id_In Is Not Null Then
    Delete From ������ϼ�¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ������� = 13 And ��ϴ��� = 1 And ��¼��Դ = 2;
    Insert Into ������ϼ�¼
      (ID, ����id, ��ҳid, ��¼��Դ, �������, ��ϴ���, ����id, ���id, �������, �Ƿ�����, ��Ժ���, ��¼����, ��¼��)
    Values
      (������ϼ�¼_Id.Nextval, ����id_In, ��ҳid_In, 2, 13, 1, ��ҽ����id_In, ��ҽ���id_In, ��ҽ���_In, Null, ��ҽ��Ժ���_In, Sysdate, ����Ա����_In);
  End If;

  --PDAͬ����־д��
  If Zl_Pda_Enabled > 0 Then
    Zl_Pdasynch_Log(1, ����id_In, 2);
  End IF;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˱䶯��¼_Out;
/

