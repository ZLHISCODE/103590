ALTER TABLE ԤԼ��ʽ add ԤԼ���� number(5);

ALTER  Table ʱ��� modify ʱ��� varchar2(10);

ALTER  Table ʱ��� ADD (
   վ�� varchar2(1),
   ���� varchar2(10),
   ����Ԥ��ʱ�� number(18),
   ��Ϣʱ�� varchar2(200));
Alter Table ʱ���  drop Constraint ʱ���_PK   Cascade Drop Index;
Alter Table ʱ��� Add Constraint ʱ���_UQ_ʱ��� Unique (ʱ���,����,վ��) Using Index Tablespace zl9Indexhis;
Alter Table ʱ��� Modify ʱ��� Constraint ʱ���_NN_ʱ��� Not Null;

Create Table ����ͣ��ԭ��(
   ���� varchar2(5),
   ���� varchar2(50),
   ���� varchar2(20),
   ȱʡ��־ number(1) default 0)
TABLESPACE zl9BaseItem ;

Alter Table ����ͣ��ԭ��  Add Constraint ����ͣ��ԭ��_PK  Primary Key (����) Using Index Tablespace zl9Indexhis;
Alter Table ����ͣ��ԭ�� Add Constraint ����ͣ��ԭ��_UQ_���� Unique (����) Using Index Tablespace zl9Indexhis;

Create Sequence ��������_ID start with 1;
ALTER TABLE �������� add(ID number(18));

 
Declare
  Cursor c_�������� Is
    Select ��������_Id.Nextval ID, ���� From �������� Where ID Is Null;
  n_Array_Size Number := 200;

  t_Id       t_Numlist;
  t_�������� t_Strlist;
Begin

  Open c_��������;

  Loop
    Fetch c_�������� Bulk Collect
      Into t_Id, t_�������� Limit n_Array_Size;
    Exit When t_��������.Count = 0;
  
    --ѭ������������ü�¼
    Forall I In 1 .. t_��������.Count
      Update �������� Set ID = t_Id(I) Where ���� = t_��������(I);
  End Loop;
  COMMIT ;
  Close c_��������;
End;
/

Alter Table ��������  drop Constraint ��������_PK   Cascade Drop Index;
Alter Table ��������  Add Constraint ��������_PK  Primary Key (ID) Using Index Tablespace zl9Indexhis;
Alter Table �������� Add Constraint ��������_UQ_���� Unique (����) Using Index Tablespace zl9Indexhis;


CREATE TABLE �����������ÿ��� (
	����ID number(18),
	����ID number(18),
	ȱʡ��־ number(2)) 
TABLESPACE zl9BaseItem ;

Alter Table �����������ÿ���  Add Constraint �����������ÿ���_PK  Primary Key (����ID,����ID) Using Index Tablespace zl9Indexhis;
Alter Table �����������ÿ��� Add Constraint �����������ÿ���_FK_����ID Foreign Key (����ID) References ��������( ID) ;
Alter Table �����������ÿ��� Add Constraint �����������ÿ���_FK_����ID Foreign Key (����ID) References ���ű�( ID) ;

Create Index �����������ÿ���_IX_����id on �����������ÿ���(����id) Tablespace zl9Indexhis;


Create Table �������ձ�(
   ��� number(18),
   �������� varchar2(50),
   ���� number(18),
   ��ʼ���� Date,
   ��ֹ���� Date,
   ��ע varchar2(1000),
   ����Һ� varchar2(500),
   ����ԤԼ varchar2(500))
TABLESPACE zl9BaseItem ;

Alter Table �������ձ�  Add Constraint �������ձ�_PK  Primary Key (��ʼ����,���,��������,����) Using Index Tablespace zl9Indexhis;

ALTER TABLE �Һź�����λ ADD ����ʱ�� number(18) ;

   

Create Sequence �ٴ������Դ_ID start with 1;
Create Table �ٴ������Դ(
   ID number(18) not null,
   ���� varchar2(10),
   ���� varchar2(5),
   ����id number(18),
   ��ĿID number(18),
   ҽ��id number(18),
   ҽ������ varchar2(50),
   �Ƿ񽨲��� number(2) default 0,
   ԤԼ���� number(3),
   ����Ƶ�� number(3),
   ���տ���״̬ number(2) ,
   �Ƿ���ջ��� number(2) default 0,
   �Ƿ��ٴ��Ű� number(2) default 0,
   �Ű෽ʽ number(2),
   �Ƿ�ɾ�� number(2) default 0,
   ����ʱ�� Date,
   ����ʱ�� Date)
TABLESPACE zl9BaseItem ;

Alter Table �ٴ������Դ Add Constraint �ٴ������Դ_PK  Primary Key (ID) Using Index Tablespace zl9Indexhis;
Alter Table �ٴ������Դ Add Constraint �ٴ������Դ_UQ_���� Unique (����,����ʱ��) Using Index Tablespace zl9Indexhis;
Alter Table �ٴ������Դ Add Constraint �ٴ������Դ_UQ_������Ŀ Unique (����ID,��ĿID,ҽ������,ҽ��ID,����ʱ��) Using Index Tablespace zl9Indexhis;
Alter Table �ٴ������Դ Add Constraint �ٴ������Դ_FK_����ID Foreign Key (����ID) References ���ű�( ID) ;
Alter Table �ٴ������Դ Add Constraint �ٴ������Դ_FK_��ĿID Foreign Key (��ĿID) References �շ���ĿĿ¼(ID) ;
Alter Table �ٴ������Դ Add Constraint �ٴ������Դ_FK_ҽ��id Foreign Key (ҽ��id) References ��Ա��(ID) ;
 
Create Index �ٴ������Դ_IX_��ĿID on �ٴ������Դ(��ĿID) Tablespace zl9Indexhis;
Create Index �ٴ������Դ_IX_ҽ��id on �ٴ������Դ(ҽ��id) Tablespace zl9Indexhis;
Create Index �ٴ������Դ_IX_ҽ������ on �ٴ������Դ(ҽ������) Tablespace zl9Indexhis;


Create Sequence �ٴ������Դ����_ID start with 1;
Create Table �ٴ������Դ����(
   ID number(18) not null,
   ��ԴID number(18),
   �ϰ�ʱ�� varchar2(10),
   �޺��� number(10),
   ��Լ�� number(10),
   �Ƿ���ſ��� number(2) default 0,
   �Ƿ��ʱ��  NUMBER(2),
   ԤԼ���� number(2),
   �Ƿ��ռ number(2) default 0,   
   ���﷽ʽ number(3),
   ����ID number(18))
TABLESPACE zl9BaseItem ;

Alter Table �ٴ������Դ����  Add Constraint �ٴ������Դ����_PK  Primary Key (ID) Using Index Tablespace zl9Indexhis;
Alter Table �ٴ������Դ���� Add Constraint �ٴ������Դ����_FK_��ԴID Foreign Key (��ԴID) References �ٴ������Դ( ID) ;
Alter Table �ٴ������Դ����  Add Constraint �ٴ������Դ����_UQ_��ԴID  Unique (��ԴID,�ϰ�ʱ��) Using Index Tablespace zl9Indexhis; 
Alter Table �ٴ������Դ���� Add Constraint �ٴ������Դ����_FK_����ID Foreign Key (����ID) References ��������( ID) ;
create Index �ٴ������Դ����_IX_����ID on �ٴ������Դ����(����ID) Tablespace zl9Indexhis;




Create Table �ٴ������Դ����(
   ����ID number(18),
   ����ID number(18))
TABLESPACE zl9BaseItem ;

Alter Table �ٴ������Դ����  Add Constraint �ٴ������Դ����_PK  Primary Key (����ID,����ID) Using Index Tablespace zl9Indexhis;
Alter Table �ٴ������Դ���� Add Constraint �ٴ������Դ����_FK_����ID Foreign Key (����ID) References �ٴ������Դ����( ID) ;
Alter Table �ٴ������Դ���� Add Constraint �ٴ������Դ����_FK_����ID Foreign Key (����ID) References ��������( ID) ;
create Index �ٴ������Դ����_IX_����ID on �ٴ������Դ����(����ID) Tablespace zl9Indexhis;



Create Table �ٴ������Դʱ��(
   ����ID number(18),
   ��� number(18),
   ��ʼʱ�� Date,
   ��ֹʱ�� Date,
   �������� number(10),
   �Ƿ�ԤԼ number(2))
TABLESPACE zl9BaseItem;

Alter Table �ٴ������Դʱ��  Add Constraint �ٴ������Դʱ��_PK  Primary Key (����ID,���) Using Index Tablespace zl9Indexhis;
Alter Table �ٴ������Դʱ�� Add Constraint �ٴ������Դʱ��_FK_����ID Foreign Key (����ID) References �ٴ������Դ����( ID) ;

Create Table �ٴ������Դ����(
   ����ID number(18),
   ���� number(2),
   ���� number(2),
   ���� varchar2(50),
   ��� number(18),
   ���Ʒ�ʽ number(2),
   ���� number(16,5))
TABLESPACE zl9BaseItem ;

Alter Table �ٴ������Դ����  Add Constraint �ٴ������Դ����_PK  Primary Key (����ID,����,����,����,���) Using Index Tablespace zl9Indexhis;
Alter Table �ٴ������Դ���� Add Constraint �ٴ������Դ����_FK_����ID Foreign Key (����ID) References �ٴ������Դ����(ID);

Create Sequence �ٴ������_ID start with 1;

Create Table �ٴ������(
   ID number(18) not null,
   �Ű෽ʽ number(18),
   ������� varchar2(50),
   ��� number(4),
   �·� number(2),
   ���� number(2),
   Ӧ�÷�Χ number(2),
   ����ID number(18),
   ��ע varchar2(100),
   ������ varchar2(50),
   ����ʱ�� Date)
TABLESPACE zl9BaseItem ;

Alter Table �ٴ������  Add Constraint �ٴ������_PK  Primary Key (ID) Using Index Tablespace zl9Indexhis;
Alter Table �ٴ������  Add Constraint �ٴ������_UQ_�������  Unique (���,�·�,����,�������,�Ű෽ʽ) Using Index Tablespace zl9Indexhis;
Alter Table �ٴ������ Add Constraint �ٴ������_FK_����ID Foreign Key (����ID) References ���ű�(ID) ;

Create Index �ٴ������_IX_����ID on �ٴ������(����ID) Tablespace zl9Indexhis;



Create Sequence �ٴ����ﰲ��_ID start with 1;
Create Table �ٴ����ﰲ��(
   ID number(18) not null,
   ����ID number(18),
   ��ԴID number(18),
   ��ĿID number(18),
   ҽ��id number(18),
   ҽ������ varchar2(50),
   �Ű���� number(2),
   �Ƿ��������� number(2),
   �Ƿ����ճ��� number(2),
   ��ʼʱ�� Date,
   ��ֹʱ�� Date,
   ����Ա���� varchar2(50),
   �Ǽ�ʱ�� Date,
   ԭ��ֹʱ�� Date)
TABLESPACE zl9BaseItem ;

Alter Table �ٴ����ﰲ�� Add Constraint �ٴ����ﰲ��_PK  Primary Key (ID) Using Index Tablespace zl9Indexhis;
Alter Table �ٴ����ﰲ��  Add Constraint �ٴ����ﰲ��_UQ_����ID  Unique (����ID,��ԴID,��ʼʱ��) Using Index Tablespace zl9Indexhis;

Alter Table �ٴ����ﰲ�� Add Constraint �ٴ����ﰲ��_FK_��ԴID Foreign Key (��ԴID) References �ٴ������Դ( ID) ;
Alter Table �ٴ����ﰲ�� Add Constraint �ٴ����ﰲ��_FK_����ID Foreign Key (����ID) References �ٴ������( ID) ;
Alter Table �ٴ����ﰲ�� Add Constraint �ٴ����ﰲ��_FK_��ĿID Foreign Key (��ĿID) References �շ���ĿĿ¼(ID) ;
Alter Table �ٴ����ﰲ�� Add Constraint �ٴ����ﰲ��_FK_ҽ��id Foreign Key (ҽ��id) References ��Ա��(ID);



Create Index �ٴ����ﰲ��_IX_��ĿID on �ٴ����ﰲ��(��ĿID) Tablespace zl9Indexhis;
Create Index �ٴ����ﰲ��_IX_ҽ��id on �ٴ����ﰲ��(ҽ��id) Tablespace zl9Indexhis;
Create Index �ٴ����ﰲ��_IX_��ԴID on �ٴ����ﰲ��(��ԴID) Tablespace zl9Indexhis;


Create Sequence �ٴ���������_ID start with 1;
Create Table �ٴ���������(
   ID     number(18),
   ����ID number(18),
   ������Ŀ varchar2(20),
   �ϰ�ʱ�� varchar2(10),
   �޺��� number(10),
   ��Լ�� number(10),
   �Ƿ���ſ��� number(2),
   �Ƿ��ʱ�� NUMBER(2),
   ԤԼ���� number(2),
   ���﷽ʽ number(2),
   ����ID number(18),
   �Ƿ��ռ number(2))
TABLESPACE zl9BaseItem ;

Alter Table �ٴ���������  Add Constraint �ٴ���������_PK  Primary Key (ID) Using Index Tablespace zl9Indexhis;
Alter Table �ٴ���������  Add Constraint �ٴ���������_UQ_����ID  Unique (����ID,������Ŀ,�ϰ�ʱ��) Using Index Tablespace zl9Indexhis;
Alter Table �ٴ��������� Add Constraint �ٴ���������_FK_����ID Foreign Key (����ID) References �ٴ����ﰲ��( ID) ;
Alter Table �ٴ��������� Add Constraint �ٴ���������_FK_����id Foreign Key (����id) References ��������(ID) ;
Create Index �ٴ���������_IX_����ID on �ٴ���������(����ID) Tablespace zl9Indexhis;



Create Table �ٴ�����ʱ��(
   ����ID number(18),
   ��� number(18),
   ��ʼʱ�� Date,
   ��ֹʱ�� Date,
   �������� number(10),
   �Ƿ�ԤԼ number(2))
TABLESPACE zl9BaseItem ;

Alter Table �ٴ�����ʱ��  Add Constraint �ٴ�����ʱ��_PK  Primary Key (����ID,���) Using Index Tablespace zl9Indexhis;
Alter Table �ٴ�����ʱ�� Add Constraint �ٴ�����ʱ��_FK_����ID Foreign Key (����ID) References �ٴ���������( ID) ;


Create Table �ٴ���������(
   ����ID number(18),
   ����ID number(18))
TABLESPACE zl9BaseItem ;
Alter Table �ٴ���������  Add Constraint �ٴ���������_PK  Primary Key (����ID,����id) Using Index Tablespace zl9Indexhis;
Alter Table �ٴ��������� Add Constraint �ٴ���������_FK_����ID Foreign Key (����ID) References �ٴ���������( ID) ;

Alter Table �ٴ��������� Add Constraint �ٴ���������_FK_����id Foreign Key (����id) References ��������(ID) ;
Create Index �ٴ���������_IX_����ID on �ٴ���������(����ID) Tablespace zl9Indexhis;


Create Table �ٴ�����Һſ���(
   ����ID number(18),
   ���� number(2),
   ���� number(2),
   ���� varchar2(50),
   ��� number(18),
   ���Ʒ�ʽ number(2),
   ���� number(16,5))
TABLESPACE zl9BaseItem ;

Alter Table �ٴ�����Һſ���  Add Constraint �ٴ�����Һſ���_PK  Primary Key (����ID,���,����,����,����) Using Index Tablespace zl9Indexhis;
Alter Table �ٴ�����Һſ��� Add Constraint �ٴ�����Һſ���_FK_����ID Foreign Key (����ID) References �ٴ���������( ID) ;
 

Create Sequence �ٴ������¼_ID start with 1;
Create Table �ٴ������¼(
   ID number(18) not null,
   ����ID number(18),
   ��ԴID number(18),
   �������� Date,
   �ϰ�ʱ�� varchar2(10),
   ��ʼʱ�� Date,
   ��ֹʱ�� Date,
   ͣ�￪ʼʱ�� Date,
   ͣ����ֹʱ�� Date,
   ͣ��ԭ�� varchar2(50),
   ȱʡԤԼʱ�� Date,
   ��ǰ�Һ�ʱ�� Date,
   �޺��� number(10),
   �ѹ��� number(10),
   ��Լ�� number(10),
   ��Լ�� number(10),
   �����ѽ��� number(10),
   �Ƿ���ſ��� number(2) default 0,
   �Ƿ��ʱ�� number(2) default 0,
   ԤԼ���� number(2),
   �Ƿ��ռ number(2),
   ��ĿID number(18),
   ����ID number(18),
   ҽ��id number(18),
   ҽ������ varchar2(50),
   ����ҽ��id number(18),
   ����ҽ������ varchar2(50),
   ���﷽ʽ number(2),
   ����ID Number(18),
   �Ƿ����� number(2) default 0,
   �Ƿ���ʱ���� number(2) default 0,
   �Ǽ��� varchar2(50),
   �Ǽ�ʱ�� Date,
   �Ƿ񷢲� number(2) default 0)
TABLESPACE zl9BaseItem;

Alter Table �ٴ������¼  Add Constraint �ٴ������¼_PK  Primary Key (ID) Using Index Tablespace zl9Indexhis;

Alter Table �ٴ������¼  Add Constraint �ٴ������¼_UQ_��������  Unique (��������,��ԴID,�ϰ�ʱ��) Using Index Tablespace zl9Indexhis;

Alter Table �ٴ������¼ Add Constraint �ٴ������¼_FK_����ID Foreign Key (����ID) References �ٴ����ﰲ��( ID) ;
Alter Table �ٴ������¼ Add Constraint �ٴ������¼_FK_��ԴID Foreign Key (��ԴID) References �ٴ������Դ( ID) ;

Alter Table �ٴ������¼ Add Constraint �ٴ������¼_FK_��ĿID Foreign Key (��ĿID) References �շ���ĿĿ¼(ID) ;
Alter Table �ٴ������¼ Add Constraint �ٴ������¼_FK_����ID Foreign Key (����ID) References ���ű�(ID) ;
Alter Table �ٴ������¼ Add Constraint �ٴ������¼_FK_ҽ��id Foreign Key (ҽ��id) References ��Ա��(ID) ;
Alter Table �ٴ������¼ Add Constraint �ٴ������¼_FK_����ҽ��id Foreign Key (����ҽ��id) References ��Ա��(ID) ;
Alter Table �ٴ������¼ Add Constraint �ٴ������¼_FK_����id Foreign Key (����id) References ��������(ID) ;
Create Index �ٴ������¼_IX_����ID on �ٴ������¼(����ID) Tablespace zl9Indexhis;

Create Index �ٴ������¼_IX_����ID on �ٴ������¼(����ID) Tablespace zl9Indexhis;
Create Index �ٴ������¼_IX_��ԴID on �ٴ������¼(��ԴID) Tablespace zl9Indexhis;
Create Index �ٴ������¼_IX_����ҽ��id on �ٴ������¼(����ҽ��id) Tablespace zl9Indexhis;

Create Index �ٴ������¼_IX_ҽ��id on �ٴ������¼(ҽ��id) Tablespace zl9Indexhis;
Create Index �ٴ������¼_IX_��ĿID on �ٴ������¼(��ĿID) Tablespace zl9Indexhis;
Create Index �ٴ������¼_IX_����ID on �ٴ������¼(����ID) Tablespace zl9Indexhis;
Create Index �ٴ������¼_IX_��ʼʱ�� on �ٴ������¼(��ʼʱ��,��ԴID) Tablespace zl9Indexhis;


Create Sequence �ٴ�����ͣ���¼_ID start with 1;
Create Table �ٴ�����ͣ���¼(
   ID number(18) not null,
   ��¼ID number(18),
   ��ʼʱ�� Date,
   ��ֹʱ�� Date,
   ͣ��ԭ�� varchar2(50),
   ����ҽ��ID number(18),
   ����ҽ������ varchar2(50),
   ������ varchar2(50),
   ����ʱ�� Date,
   ������ varchar2(50),
   ����ʱ�� Date,
   ȡ���� varchar2(50),
   ȡ��ʱ�� Date)
TABLESPACE zl9BaseItem ;

Alter Table �ٴ�����ͣ���¼  Add Constraint �ٴ�����ͣ���¼_PK  Primary Key (ID) Using Index Tablespace zl9Indexhis;
Alter Table �ٴ�����ͣ���¼ Add Constraint �ٴ�����ͣ���¼_FK_��¼ID Foreign Key (��¼ID) References �ٴ������¼( ID) ;
Alter Table �ٴ�����ͣ���¼ Add Constraint �ٴ�����ͣ���¼_FK_����ҽ��ID Foreign Key (����ҽ��ID) References ��Ա��( ID) ;

Create Index �ٴ�����ͣ���¼_IX_��¼ID on �ٴ�����ͣ���¼(��¼ID) Tablespace zl9Indexhis;
Create Index �ٴ�����ͣ���¼_IX_����ҽ��ID on �ٴ�����ͣ���¼(����ҽ��ID) Tablespace zl9Indexhis;
Create Index �ٴ�����ͣ���¼_IX_����ʱ�� on �ٴ�����ͣ���¼(����ʱ��) Tablespace zl9Indexhis;
Create Index �ٴ�����ͣ���¼_IX_����ʱ�� on �ٴ�����ͣ���¼(����ʱ��) Tablespace zl9Indexhis;

Create Table �ٴ�������ſ���(
   ��¼ID number(18),
   ��� number(18),
   ԤԼ˳��� number(18),
   ��ʼʱ�� Date,
   ��ֹʱ�� Date,
   ���� number(10),
   �Ƿ�ԤԼ number(2),
   �Һ�״̬ number(2),
   ����ʱ�� Date,
   ����   number(2),
   ���� varchar2(50),
   ����Ա���� varchar2(50),
   ����վIP varchar2(20),
   ����վ���� varchar2(200),
   ��ע varchar2(100))
TABLESPACE zl9BaseItem ;

Alter Table �ٴ�������ſ���  Add Constraint �ٴ�������ſ���_UQ_��¼ID  Unique (��¼ID,���,ԤԼ˳���) Using Index Tablespace zl9Indexhis;
Alter Table �ٴ�������ſ��� Add Constraint �ٴ�������ſ���_FK_��¼ID Foreign Key (��¼ID) References �ٴ������¼( ID) ;
Alter Table �ٴ�������ſ��� Modify ��¼ID Constraint �ٴ�������ſ���_NN_��¼ID Not Null;


Create Table �ٴ��������Ҽ�¼(
   ��¼ID number(18),
   ����ID Number(18),
   ��ǰ���� number(1) default 0)
TABLESPACE zl9BaseItem ;

Alter Table �ٴ��������Ҽ�¼ Add Constraint �ٴ��������Ҽ�¼_PK  Primary Key (��¼ID,����ID) Using Index Tablespace zl9Indexhis;
Alter Table �ٴ��������Ҽ�¼ Add Constraint �ٴ��������Ҽ�¼_FK_��¼ID Foreign Key (��¼ID) References �ٴ������¼( ID) ;
Alter Table �ٴ��������Ҽ�¼ Add Constraint �ٴ��������Ҽ�¼_FK_����id Foreign Key (����id) References ��������(ID) ;
Create Index �ٴ��������Ҽ�¼_IX_����ID on �ٴ��������Ҽ�¼(����ID) Tablespace zl9Indexhis;




Create Table �ٴ�����Һſ��Ƽ�¼(
   ��¼ID number(18),
   ���� number(2),
   ���� number(2),
   ���� varchar2(50),
   ��� number(18),
   ���Ʒ�ʽ number(2),
   ���� number(16,5))
TABLESPACE zl9BaseItem ;

Alter Table �ٴ�����Һſ��Ƽ�¼  Add Constraint �ٴ�����Һſ��Ƽ�¼_PK  Primary Key (��¼ID,����,���,����,����) Using Index Tablespace zl9Indexhis;
Alter Table �ٴ�����Һſ��Ƽ�¼ Add Constraint �ٴ�����Һſ��Ƽ�¼_FK_��¼ID Foreign Key (��¼ID) References �ٴ������¼(ID) ;


Create Sequence ���˷�����Ϣ��¼_ID start with 1;
Create Table ���˷�����Ϣ��¼(
   ID number(18) not null,
   ֪ͨ���� number(18),
   ��¼ID number(18),
   �Һ�ID number(18),
   ��ԴID number(18),
   ���� varchar2(10),
   ����ID number(18),
   ��ĿID number(18),
   ҽ��ID number(18),
   ҽ������ varchar2(50),
   ����ID number(18),
   ���﷽ʽ number(2),
   ���� number(10),
   ��ʼʱ�� Date,
   ��ֹʱ�� Date,
   ֪ͨԭ�� varchar2(100),
   �Ǽ��� varchar2(50),
   �Ǽ�ʱ�� Date,
   ����˵�� varchar2(100),
   ������ varchar2(50),
   ����ʱ�� Date)
TABLESPACE zl9BaseItem ;

Alter Table ���˷�����Ϣ��¼  Add Constraint ���˷�����Ϣ��¼_PK  Primary Key (ID) Using Index Tablespace zl9Indexhis;
Alter Table ���˷�����Ϣ��¼ Add Constraint ���˷�����Ϣ��¼_FK_��ԴID Foreign Key (��ԴID) References �ٴ������Դ( ID) ;
Alter Table ���˷�����Ϣ��¼ Add Constraint ���˷�����Ϣ��¼_FK_��¼ID Foreign Key (��¼ID) References �ٴ������¼( ID) ;

Alter Table ���˷�����Ϣ��¼ Add Constraint ���˷�����Ϣ��¼_FK_��ĿID Foreign Key (��ĿID) References �շ���ĿĿ¼(ID) ;
Alter Table ���˷�����Ϣ��¼ Add Constraint ���˷�����Ϣ��¼_FK_ҽ��id Foreign Key (ҽ��id) References ��Ա��(ID) ;
Alter Table ���˷�����Ϣ��¼ Add Constraint ���˷�����Ϣ��¼_FK_����ID Foreign Key (����ID) References ���ű�(ID) ;
Alter Table ���˷�����Ϣ��¼ Add Constraint ���˷�����Ϣ��¼_FK_����ID Foreign Key (����ID) References ������Ϣ(����ID) ;

Create Index ���˷�����Ϣ��¼_IX_�Ǽ�ʱ�� on ���˷�����Ϣ��¼(�Ǽ�ʱ��) Tablespace zl9Indexhis;
Create Index ���˷�����Ϣ��¼_IX_����ʱ�� on ���˷�����Ϣ��¼(����ʱ��) Tablespace zl9Indexhis;

Create Index ���˷�����Ϣ��¼_IX_����ID on ���˷�����Ϣ��¼(����ID) Tablespace zl9Indexhis;
Create Index ���˷�����Ϣ��¼_IX_�Һ�ID on ���˷�����Ϣ��¼(�Һ�ID) Tablespace zl9Indexhis;
Create Index ���˷�����Ϣ��¼_IX_����ID on ���˷�����Ϣ��¼(����) Tablespace zl9Indexhis;
Create Index ���˷�����Ϣ��¼_IX_��ԴID on ���˷�����Ϣ��¼(��ԴID) Tablespace zl9Indexhis;
Create Index ���˷�����Ϣ��¼_IX_��¼ID on ���˷�����Ϣ��¼(��¼ID) Tablespace zl9Indexhis;
Create Index ���˷�����Ϣ��¼_IX_����ID on ���˷�����Ϣ��¼(����ID) Tablespace zl9Indexhis;
Create Index ���˷�����Ϣ��¼_IX_��ĿID on ���˷�����Ϣ��¼(��ĿID) Tablespace zl9Indexhis;
Create Index ���˷�����Ϣ��¼_IX_ҽ��ID on ���˷�����Ϣ��¼(ҽ��ID) Tablespace zl9Indexhis;



Create Sequence �ٴ�����䶯��¼_ID start with 1;
Create Table �ٴ�����䶯��¼(
   ID number(18) not null,
   ��¼ID number(18),
   �䶯���� number(2),
   ԭԤԼ���� number(2),
   ��ԤԼ���� number(2),
   ԭ���� number(10),
   ������ number(10),
   ԭ���﷽ʽ number(2),
   ԭ�������� varchar2(20),
   ԭ����ID number(18),
   �ַ��﷽ʽ number(2),
   ���������� varchar2(20),
   ������ID number(18),
   ����Ա���� varchar2(50),
   �Ǽ�ʱ�� Date)
TABLESPACE zl9BaseItem ;

Alter Table �ٴ�����䶯��¼  Add Constraint �ٴ�����䶯��¼_PK  Primary Key (ID) Using Index Tablespace zl9Indexhis;
Alter Table �ٴ�����䶯��¼ Add Constraint �ٴ�����䶯��¼_FK_��¼ID Foreign Key (��¼ID) References �ٴ������¼( ID) ;
Create Index �ٴ�����䶯��¼_IX_��¼ID on �ٴ�����䶯��¼(��¼ID) Tablespace zl9Indexhis;
Create Index �ٴ�����䶯��¼_IX_�Ǽ�ʱ�� on �ٴ�����䶯��¼(�Ǽ�ʱ��) Tablespace zl9Indexhis;

Alter Table �ٴ�����䶯��¼ Add Constraint �ٴ�����䶯��¼_FK_ԭ����id Foreign Key (ԭ����id) References ��������(ID) ;
Create Index �ٴ�����䶯��¼_IX_ԭ����ID on �ٴ�����䶯��¼(ԭ����ID) Tablespace zl9Indexhis;

Alter Table �ٴ�����䶯��¼ Add Constraint �ٴ�����䶯��¼_FK_������id Foreign Key (������id) References ��������(ID) ;
Create Index �ٴ�����䶯��¼_IX_������ID on �ٴ�����䶯��¼(������ID) Tablespace zl9Indexhis;


Create Table �ٴ�����䶯��ϸ(
   �䶯ID number(18),
   �䶯���� number(2),
   ���� number(2),
   ���� varchar2(50),
   ��� number(18),
   ���Ʒ�ʽ number(2),
   ���� number(10),
   ����ID number(18),
   �������� varchar2(20))
TABLESPACE zl9BaseItem ;

Alter Table �ٴ�����䶯��ϸ  Add Constraint �ٴ�����䶯��ϸ_PK  Primary Key (�䶯ID,����,�䶯����,���) Using Index Tablespace zl9Indexhis;
Alter Table �ٴ�����䶯��ϸ Add Constraint �ٴ�����䶯��ϸ_FK_�䶯ID Foreign Key (�䶯ID) References �ٴ�����䶯��¼( ID) ;

Alter Table �ٴ�����䶯��ϸ Add Constraint �ٴ�����䶯��ϸ_FK_����id Foreign Key (����id) References ��������(ID) ;
Create Index �ٴ�����䶯��ϸ_IX_����ID on �ٴ�����䶯��ϸ(����ID) Tablespace zl9Indexhis;


ALTER TABLE ���˹Һż�¼ ADD (�����¼ID number(18));
Alter Table ���˹Һż�¼ Add Constraint ���˹Һż�¼_FK_�����¼ID Foreign Key (�����¼ID) References �ٴ������¼( ID) ;
Create Index ���˹Һż�¼_�����¼ID on ���˹Һż�¼(�����¼ID) Tablespace zl9Indexhis;


--����ͣ��ԭ������
insert into ����ͣ��ԭ��(����,����,����,ȱʡ��־) values ('01','����','SS',0);
insert into ����ͣ��ԭ��(����,����,����,ȱʡ��־) values ('02','����','HZ',0);
insert into ����ͣ��ԭ��(����,����,����,ȱʡ��־) values ('03','����','GX',0);
insert into ����ͣ��ԭ��(����,����,����,ȱʡ��־) values ('04','����','BJ',0);
insert into ����ͣ��ԭ��(����,����,����,ȱʡ��־) values ('05','�¼�','SJ',0);
insert into ����ͣ��ԭ��(����,����,����,ȱʡ��־) values ('06','����','QT',0);


--�����������ÿ�������
Insert Into �����������ÿ���
  (����id, ����id, ȱʡ��־)
  Select *
  From (Select Distinct q.Id, m.����id, 0 As ȱʡ��־
         From (Select Distinct b.����id, ��������
                From �ҺŰ������� A, �ҺŰ��� B
                Where a.�ű�id = b.Id
                Union All
                Select Distinct c.����id, ��������
                From �Һżƻ����� A, �ҺŰ��żƻ� B, �ҺŰ��� C
                Where a.�ƻ�id = b.Id And b.����id = c.Id) M, �������� Q
         Where m.�������� = q.����)��


--ģ����ش���
--1114:�ٴ����ﰲ��
Insert Into zlPrograms(���,����,˵��,ϵͳ,����) Values( 1114,'�ٴ����ﰲ��','�Ա���λ�ٴ����ҵĳ��ﰲ�Ž��й���',&n_System,'zl9RegEvent'); 
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��,ȱʡֵ)
Select &n_System,1114,A.* From (
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0 Union All 
    Select '����',-Null,NULL,1 From Dual Union All 
    Select 'ʱ�������',1,'���ӡ�ɾ�����޸ĵĹҺ�ʱ��Ĳ���Ȩ�ޡ��и�Ȩ��ʱ������ԹҺ���Ŀ�ĸ�ʱ��ν��ж���',1 From Dual Union All 
    Select '�ڼ�������',2,'���ӡ�ɾ�����޸ķ����ڼ��յĲ���Ȩ�ޡ��и�Ȩ��ʱ������Ը������ڼ��ս��ж���',1 From Dual Union All 
    Select '������������',3,'���ӡ�ɾ�����޸��������ҵĲ���Ȩ�ޡ��и�Ȩ��ʱ������Ը��������ҽ�������',1 From Dual Union All 
    Select '�����Դ����',4,'���ӡ�ɾ�����޸ġ�ͣ�ü����ó����Դ�Ĳ���Ȩ�ޡ��и�Ȩ��ʱ��������Ը���Դ�������ӣ��޸ģ�ɾ��,ͣ�ü�����',1 From Dual Union All 
    Select 'ģ�����',5,'���ӡ�ɾ�����޸ĳ���ģ��Ĳ���Ȩ�ޡ��и�Ȩ��ʱ���������ģ��������ӣ��޸ļ�ɾ����',1 From Dual Union All 
    Select '���ﰲ��',6,'��Ը���Դ�ĳ�����а��ŵĲ���Ȩ�ޣ��и�Ȩ��ʱ��������Ը���Դ�ĳ�����а��š�',1 From Dual Union All 
    Select '��������',7,'���ָ���İ��ű���з����������и�Ȩ��ʱ��������Գ������з���������',1 From Dual Union All 
    Select 'ȡ������',8,'����Ѿ������İ��ű����ȡ���������и�Ȩ��ʱ��������Գ�������ȡ������������',1 From Dual Union All 
    Select '��ʱ���ﰲ��',9,'���Ѿ������İ��ŵ���δ���г��ﰲ�ŵĺ�Դ������ʱ���ﰲ�Ų������и�Ȩ��ʱ���������δ����ĺ�Դ������ʱ���ﰲ�š�',1 From Dual Union All 
    Select 'ͣ��',10,'���Ѿ������İ��Ž���ͣ��������и�Ȩ��ʱ��������Ժ�Դ����ͣ�������',1 From Dual Union All 
    Select '����',11,'���Ѿ������İ��Ž�������������и�Ȩ��ʱ��������Ժ�Դ�������������',1 From Dual Union All 
    Select '�Ӻ�',12,'���Ѿ������İ��Ž��мӺŲ������и�Ȩ��ʱ��������Ժ�Դ���мӺŲ�����',1 From Dual Union All 
    Select '����',13,'���Ѿ������İ��Ž��м��Ų������и�Ȩ��ʱ��������Ժ�Դ���м��Ų�����',1 From Dual Union All 
    Select '������������',14,'���Ѿ������İ��Ž������ҵ����������и�Ȩ��ʱ��������Ժ�Դ�������ҵ���������',1 From Dual Union All 
    Select '����ԤԼ�Һ�',15,'���Ѿ������İ��Ž��к�����λ��ԤԼ��ʽ��ԤԼ���Ƶ����������и�Ȩ��ʱ��������Ժ�Դ���к�����λ��ԤԼ��ʽ��ԤԼ���Ƶ�����',1 From Dual Union All 
    Select 'ͣ������',16,'��ҽ���ϳ�ʱ��ͣ�������������д�Ȩ��ʱ������������������',1 From Dual Union All 
    Select 'ͣ������',17,'���ͣ��������������������д�Ȩ��ʱ������������������������',1 From Dual Union All 
    Select '���������ͣ������',18,'���������ͣ���������,�д�Ȩ��ʱ���������������д���뵥�Ĳ�����',1 From Dual Union All 
    Select '���п���',19,'�������и�Ȩ��ʱ��ֻ�ܲ鿴�ʹ������ﲿ�·Ÿ�������صĺ�Դ��',1 From Dual Union All 
    Select '��������',20,'���ٴ����ﰲ�ŵĲ������ý��в�����Ȩ�ޡ��и�Ȩ��ʱ��������б��ز�������',1 From Dual Union All 
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0) A;


--1115:���߷�������
Insert Into zlPrograms(���,����,˵��,ϵͳ,����) Values( 1115,'���߷�������','��ͣ��������ԭ�����ͷ���Աͨ���绰֪ͨ�����Լ��Բ��˵�ԤԼ��Ϣ�����˺š����Ｐ����Ȳ���',&n_System,'zl9RegEvent'); 
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��,ȱʡֵ)
Select &n_System,1115,A.* From (
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0 Union All 
    Select '����',-Null,NULL,1 From Dual Union All 
    Select 'ͣ����Ϣ����',1,'ʵ������ͣ����������ʱ����Ҫ�Զ�Ӧ��ԤԼ����ȡ�������������Ĳ���Ȩ�ޡ��и�Ȩ��ʱ���������Ա��ͣ�����������Ӧ��ԤԼ����ȡ������������������',1 From Dual Union All 
    Select 'ԤԼ�Ǽ���Ϣ����',2,'ʵ�ֶԸ��ﲡ�˵��ڵ�ԤԼ�Ǽ���Ϣ�������Ѹ���������ԤԼ�ҺŵĲ���Ȩ�ޡ��и�Ȩ��ʱ���������Ա�Ե��ڵ�ԤԼ�Ǽǵ���Ϣ֪ͨ���߻������ԤԼ�ҺŵĲ���Ȩ�ޡ�',1 From Dual Union All 
    Select 'ԤԼ�ҺŵǼ�',3,'ʵ��ԤԼ�ҺŵĲ���Ȩ�ޡ��и�Ȩ��ʱ���������Ա����ԤԼ�ҺŲ�����',1 From Dual Union All 
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0) A;


Insert Into zlMenus(���,ID,�ϼ�ID,����,���,˵��,ϵͳ,ģ��,�̱���,ͼ��) Select A.���,ZlMenus_ID.Nextval,A.ID,B.* 
From (
	Select ���,ID From zlMenus Where ���� = '�ż���Һ�ϵͳ' And ��� = 'ȱʡ' And ϵͳ = &n_System And ģ�� Is Null) A,
	(	Select ����,���,˵��,ϵͳ,ģ��,�̱���,ͼ�� From zlMenus Where 1 = 0 Union All
		Select '�ٴ����ﰲ��' ,'A' ,'�Ա���λ�ٴ����ҵĳ��ﰲ�Ž��й���' ,&n_System,1114,'���ﰲ��' ,236 From Dual Union All
		Select '���߷�������' ,'B' ,'ʵ��ԤԼ�ҺŵǼ�,ȡ��ԤԼ���������Ĳ���Ȩ�ޡ��и�Ȩ��ʱ���������Ա����ԤԼ�Ǽ�,ȡ��ԤԼ����������������' ,&n_System,1115,'��������' ,220 From Dual Union All
		Select ����,���,˵��,ϵͳ,ģ��,�̱���,ͼ�� From zlMenus Where 1 = 0
          ) B;

--������ش���
Insert Into zlParameters(ID,ϵͳ,ģ��,˽��,����,��Ȩ,�̶�,����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵��)
Select zlParameters_ID.Nextval,&n_System,-Null,-Null,-Null,-Null,-Null,A.* From (
Select ����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵�� From zlParameters Where 1 = 0 Union All 
     Select 0,0,256,'�Һ��Ű�ģʽ','0','0','1. Ӱ��Һţ����ڣ�����ƽ̨�������ȣ���ȡ������:' || chr(10) || 'a.�������Ϊ�����ƻ��Ű�ģʽ��,�����ݡ�����+�ƻ����ķ�ʽ�����Ű࣬���ڹҺ�ҵ���ȡ��Ч��Դʱ�Ǵӡ��ҺŰ��š��ȱ���ȡ��' || chr(10) || 'b.�������Ϊ����������Ű�ģʽ��, �����ݡ���������̶�������³�����ܳ�����ķ�ʽ�����Ű࣬���ڹҺ�ҵ���ȡ��Ч��Դʱ�Ǵӡ���������¼���ȱ���ȡ����' || chr(10) || '2.Ӱ��ҺŴ��ڵ�չ����ʽ:' || chr(10) || '    a.�������Ϊ�����ƻ��Ű�ģʽ��,�ҺŴ�����ߵĹҺŰ������ݽ�����һ�����յķ�ʽչ��.' || chr(10) || '    b.������á�������Ű�ģʽ�����ҺŴ������ֻչ�ֵ����ָ�����ڵĹҺŰ������ݡ�','0-�ƻ��Ű�ģʽ,1-������Ű�ģʽ',Null,'1.���ҽԺҵ��ϼ򵥣�һ������������ҽԺ��,�ٴ����ҳ�����Թ̶���ֻ�н��ٵĳ���仯ʱ��ʹ�á��ƻ��Ű�ģʽ��' || chr(10) || '2.���ҽԺҵ��ϸ��ӣ�һ��������ҽԺ�����ٴ����ҳ�����ʱ�仯�ϴ�ʱ���»����Ű�ʱ��ʹ�á�������Ű�ģʽ��', '��ҽԺ����HISϵͳ��һ�㲻�����˲�������������˲���������ֱ��Ӱ�쵽�Һ�ҵ��(���ڣ�����ƽ̨������ϵͳ�ȣ���' 
     From Dual Union All 
Select ����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵�� From zlParameters Where 1 = 0) A;

--�����ű�
Insert Into zlParameters(ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
Select Zlparameters_Id.Nextval,&n_System,1114,A.* From (
  Select ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵�� From zlParameters Where 1 = 0 Union All 
  Select 1, -null, -null, 1, -null, -null, 1, '��ʾͣ�ú�Դ', '0', '0', '�鿴�ٴ������Դʱ�Ƿ���ʾ��ͣ�ú�Դ��', '0-����ʾ��1-��ʾ��', Null, '�������û��ڲ鿴��Դʱͣ�ú�Դ�ĸ��Ի���ʾ��ʽ', Null From Dual Union All
  Select -null, -null, 1, 1, -null, -null, 2, 'ֻ����ѡԺ��ҽ��', '0', '0', '���к�Դ����ʱ��ѡ���ҽ���Ƿ�ֻ��ѡ��Ժ�ڵ�ҽ����', '1-��Ժ��ҽ��;0-����ѡ����ԮҽԺ��Ժ��ҽ����', Null, '������û����Ԯҽ�����û���', Null From Dual Union All
  Select -null, -null, 0, 0, -null, -null, 3, 'ԤԼ�嵥���Ʒ�ʽ', '0', '0', 'ͣ��ʱ�Ƿ�ԤԼ�嵥�����Excel�С�', '0-�������Excel��1-�Զ������Excel,2-ѡ�������Excel��', Null, '��������ͣ��ҺŰ���ʱ��Ҫѡ���Ե����ԤԼ�嵥��ҵ��', Null From Dual Union All
  Select  -null, -null, 0, 0, -null, -null, 4, 'ԤԼ�嵥��ӡ��ʽ', '0', '0', '0-����ӡ,1-�Զ���ӡ,2-ѡ���Ƿ��ӡ��', '0-����ӡ,1-�Զ���ӡ,2-ѡ���Ƿ��ӡ��', Null, '��������ͣ��ҺŰ���ʱ��Ҫѡ���ԵĴ�ӡԤԼ�嵥��ҵ��', Null From Dual Union All
  Select  -null, -null, 0, 0, -null, -null, 5, '������ӡ��ʽ', '0', '0', '0-����ӡ,1-�Զ���ӡ,2-ѡ���Ƿ��ӡ��', '0-����ӡ,1-�Զ���ӡ,2-ѡ���Ƿ��ӡ��', Null, '�������ڷ��������ʱ��Ҫѡ���ԵĴ�ӡ������ҵ��', Null From Dual Union All
  Select  -null, -null, 0, 1, -null, -null, 6, '������ҽ��ͬ������ԤԼ�Һŵ�', '0', '0', '������ʱ������������ҽ��ͬ�����������漰�����ԤԼ�Һŵ���', '1-����ͬ������,0-�����¡�', Null, '������������ĳ�������Դ��ҽ��ҵ��ʱ��ѡ���Ƿ�ͬ������ԤԼ�Һŵ���ҽ����Ϣ', Null From Dual Union All
  Select 1, -null, -null, 1, -null, -null, 7, '��ʾȱʡ������Ϣ', '', '1', '�ں�Դ�����п�����ѡ���Դʱ���Ƿ����·���ʾ���Ƶ������Ϣ�����磺ȱʡ�������Ϣ��������Ϣ������ԤԼ������Ϣ�ȣ���', '0-����ʾ��1-��ʾ��', Null, '�������û��ڲ鿴��Դʱ��Ҫ����鿴�����õĳ����ϰ�ʱ�ΰ�����Ϣ��ҵ��', Null From Dual Union All
  Select ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵�� From zlParameters Where 1 = 0) A;
  
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1111, 0, 0, 0, 0, 68, '����ͬ���޹�N����', '0',
         '0', '����һ������ͬһ����һ������,ֻ�ܹ�N���š�', '0-������;N-��������', Null, Null, Null
  From Dual
  Where Not Exists
   (Select 1
         From zlParameters
         Where ������ = '����ͬ���޹�N����' And Nvl(ģ��, 0) = 1111 And Nvl(ϵͳ, 0) = &n_System);

Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1111, 0, 0, 0, 0, 69, '���˹Һſ�������', '0',
         '0', 'ͬһ������ͬһʱ���ܷ�Ҷ�����ҡ�', '0-������,>=1��ʾ�������Ƶ�����', Null, Null, Null
  From Dual
  Where Not Exists
   (Select 1
         From zlParameters
         Where ������ = '���˹Һſ�������' And Nvl(ģ��, 0) = 1111 And Nvl(ϵͳ, 0) = &n_System);	

Update zlParameters
Set ������ = '����ͬ����ԼN����', Ӱ�����˵�� = '��ԤԼʱ,ͬһ������ͬһʱ�估ͬһ���ҵ���������', ����ֵ����= '0-������;N-����ԤԼ����'
Where ������ = '����ͬ����Լһ����' And ģ�� = 1111 And ϵͳ = &n_System And Not Exists
 (Select 1
       From zlParameters
       Where ������ = '����ͬ����ԼN����' And Nvl(ģ��, 0) = 1111 And Nvl(ϵͳ, 0) = &n_System);

Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1802, 0, 0, 0, 0, 0, 0, 39, '�Һ�ʱѡ��ʱ��', '1',
         '1', '�����ҺŹ�ר�Һŷ�ʱ�εĺű�ʱ���Ƿ��ṩʱ��ѡ��������û�ѡ��Һ�ʱ��', '0-�����ã�1-���á�', Null, Null, Null
  From Dual
  Where Not Exists
   (Select 1
         From zlParameters
         Where ������ = '�Һ�ʱѡ��ʱ��' And Nvl(ģ��, 0) = 1802 And Nvl(ϵͳ, 0) = &n_System);
         
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1803, 0, 0, 0, 0, 0, 0, 39, 'ԤԼʱѡ��ʱ��', '1',
         '1', '����ԤԼ�ҷ�ʱ�εĺű�ʱ���Ƿ��ṩʱ��ѡ��������û�ѡ��ԤԼʱ��', '0-�����ã�1-���á�', Null, Null, Null
  From Dual
  Where Not Exists
   (Select 1
         From zlParameters
         Where ������ = 'ԤԼʱѡ��ʱ��' And Nvl(ģ��, 0) = 1803 And Nvl(ϵͳ, 0) = &n_System);


Insert Into zlProgFuncs
  (ϵͳ, ���, ����, ����, ˵��, ȱʡֵ)
  Select &n_System, 9000, 'ԤԼ�Ǽ�', 7, '������ҽ��������סԺҽ�������ȵ�ԤԼ�ǼǵĲ���Ȩ�ޡ��и�Ȩ��ʱ���������ԤԼ�Ǽǲ�����', 1
  From Dual
  Where Not Exists (Select 1 From zlProgFuncs Where ϵͳ = &n_System And ��� = 9000 And ���� = 'ԤԼ�Ǽ�');


--Ȩ�޽ű�
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1114,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All
Select '�ٴ������_ID','SELECT' From Dual Union All
Select '�ٴ����ﰲ��_ID','SELECT' From Dual Union All
Select '�ٴ������Դ_ID','SELECT' From Dual Union All
Select '�ٴ������Դ����_ID','SELECT' From Dual Union All
Select '�ٴ������¼_ID','SELECT' From Dual Union All
Select '�ٴ�����䶯��¼_ID','SELECT' From Dual Union All
Select '�ٴ���������_ID','SELECT' From Dual Union All
Select '�ٴ������Դ����','SELECT' From Dual Union All
Select '�ٴ������Դʱ��','SELECT' From Dual Union All
Select '�ٴ������Դ����','SELECT' From Dual Union All
Select '���˹Һż�¼','SELECT' From Dual Union All
Select '���ű�','SELECT' From Dual Union All
Select '��������˵��','SELECT' From Dual Union All
Select '����ͣ��ԭ��','SELECT' From Dual Union All
Select '�������ձ�','SELECT' From Dual Union All
Select '�ҺŰ���','SELECT' From Dual Union All
Select '�Һź�����λ','SELECT' From Dual Union All
Select '����','SELECT' From Dual Union All
Select '�ٴ����ﰲ��','SELECT' From Dual Union All
Select '�ٴ�����䶯��¼','SELECT' From Dual Union All
Select '�ٴ�����䶯��ϸ','SELECT' From Dual Union All
Select '�ٴ������','SELECT' From Dual Union All
Select '�ٴ�����Һſ���','SELECT' From Dual Union All
Select '�ٴ�����Һſ��Ƽ�¼','SELECT' From Dual Union All
Select '�ٴ������Դ','SELECT' From Dual Union All
Select '�ٴ������Դ����','SELECT' From Dual Union All
Select '�ٴ������¼','SELECT' From Dual Union All
Select '�ٴ�����ʱ��','SELECT' From Dual Union All
Select '�ٴ�����ͣ���¼','SELECT' From Dual Union All
Select '�ٴ���������','SELECT' From Dual Union All
Select '�ٴ�������ſ���','SELECT' From Dual Union All
Select '�ٴ���������','SELECT' From Dual Union All
Select '�ٴ��������Ҽ�¼','SELECT' From Dual Union All
Select '��������','SELECT' From Dual Union All
Select '�����������ÿ���','SELECT' From Dual Union All
Select '�ϻ���Ա��','SELECT' From Dual Union All
Select 'ʱ���','SELECT' From Dual Union All
Select '�շ���ĿĿ¼','SELECT' From Dual Union All
Select 'ԤԼ��ʽ','SELECT' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,������,����,����,Ȩ��)
Select &n_System,1114,User,A.* From (
Select ����,����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All
Select 'ʱ�������','Zl_�ϰ�ʱ��_Modify','EXECUTE' From Dual Union All
Select 'ʱ�������','Zl_�ϰ�ʱ��_Delete','EXECUTE' From Dual Union All
Select '�ڼ�������','Zl_�������ձ�_Modify','EXECUTE' From Dual Union All
Select '�ڼ�������','Zl_�������ձ�_Delete','EXECUTE' From Dual Union All
Select '������������','Zl_��������_Modify','EXECUTE' From Dual Union All
Select '������������','Zl_��������_Delete','EXECUTE' From Dual Union All
Select '�����Դ����','Zl_�ٴ������Դ_Stopandstart','EXECUTE' From Dual Union All
Select '�����Դ����','Zl_�ٴ������Դ_Modify','EXECUTE' From Dual Union All
Select '�����Դ����','Zl_�ٴ������Դ_Delete','EXECUTE' From Dual Union All
Select '�����Դ����','Zl_�ٴ������Դ����_Modify','EXECUTE' From Dual Union All
Select 'ģ�����','Zl_�ٴ������_Add','EXECUTE' From Dual Union All
Select 'ģ�����','Zl_�ٴ������_Update','EXECUTE' From Dual Union All
Select 'ģ�����','Zl_�ٴ������_Delete','EXECUTE' From Dual Union All
Select 'ģ�����','Zl_�ٴ����ﰲ��_Delete','EXECUTE' From Dual Union All
Select 'ģ�����','Zl_�ٴ������ϰ�ʱ��_Delete','EXECUTE' From Dual Union All
Select 'ģ�����','Zl_�ٴ����ﰲ��_Insert','EXECUTE' From Dual Union All
Select 'ģ�����','Zl_�ٴ���������_Insert','EXECUTE' From Dual Union All
Select 'ģ�����','Zl_�ٴ�����Һſ���_Insert','EXECUTE' From Dual Union All
Select '���ﰲ��','Zl_�ٴ������_Delete','EXECUTE' From Dual Union All
Select '���ﰲ��','Zl_�ٴ������_����','EXECUTE' From Dual Union All
Select '���ﰲ��','Zl1_Auto_Buildingregisterplan','EXECUTE' From Dual Union All
Select '���ﰲ��','Zl_�ٴ������_Add','EXECUTE' From Dual Union All
Select '���ﰲ��','Zl_�ٴ������_Update','EXECUTE' From Dual Union All
Select '���ﰲ��','Zl_�ٴ����ﰲ��_Delete','EXECUTE' From Dual Union All
Select '���ﰲ��','Zl_�ٴ������ϰ�ʱ��_Delete','EXECUTE' From Dual Union All
Select '���ﰲ��','Zl_�ٴ����ﰲ��_Insert','EXECUTE' From Dual Union All
Select '���ﰲ��','Zl_�ٴ���������_Insert','EXECUTE' From Dual Union All
Select '���ﰲ��','Zl_�ٴ�����Һſ���_Insert','EXECUTE' From Dual Union All
Select '���ﰲ��','Zl_�ٴ������¼_Insert','EXECUTE' From Dual Union All
Select '���ﰲ��','Zl_�ٴ�����Һſ��Ƽ�¼_Insert','EXECUTE' From Dual Union All
Select '���ﰲ��','Zl_�ٴ����ﰲ��_Applyto','EXECUTE' From Dual Union All
Select '���ﰲ��','Zl_Buildregisterplanbyrecord','EXECUTE' From Dual Union All
Select '���ﰲ��','Zl_�ٴ����ﰲ��_BatchDelete','EXECUTE' From Dual Union All
Select '���ﰲ��','Zl_Buildregisterfixedrule','EXECUTE' From Dual Union All
Select '���ﰲ��','Zl_Buildregisterplanbytemplet','EXECUTE' From Dual Union All
Select '���ﰲ��','Zl_�ٴ����ﰲ��_��ſ���','EXECUTE' From Dual Union All
Select '���ﰲ��','Zl_�ٴ������¼_Batchlock','EXECUTE' From Dual Union All
Select '��������','Zl_�ٴ����ﰲ��_Publish','EXECUTE' From Dual Union All
Select 'ȡ������','Zl_�ٴ����ﰲ��_Publish','EXECUTE' From Dual Union All
Select '��ʱ���ﰲ��','Zl_�ٴ����ﰲ��_Delete','EXECUTE' From Dual Union All
Select '��ʱ���ﰲ��','Zl_�ٴ����ﰲ��_Insert','EXECUTE' From Dual Union All
Select '��ʱ���ﰲ��','Zl_�ٴ������ϰ�ʱ��_Delete','EXECUTE' From Dual Union All
Select '��ʱ���ﰲ��','Zl_�ٴ������¼_Insert','EXECUTE' From Dual Union All
Select '��ʱ���ﰲ��','Zl_�ٴ�����Һſ��Ƽ�¼_Insert','EXECUTE' From Dual Union All
Select 'ͣ��','Zl_�ٴ������¼_Stopvisit','EXECUTE' From Dual Union All
Select '����','Zl_�ٴ������¼_Replacedoctor','EXECUTE' From Dual Union All
Select '����','Zl1_Ex_Isdoctorsamelevel','EXECUTE' From Dual Union All
Select '�Ӻ�','Zl_�ٴ�������ſ��Ʊ䶯','EXECUTE' From Dual Union All
Select '����','Zl_�ٴ�������ſ��Ʊ䶯','EXECUTE' From Dual Union All
Select '������������','Zl_�ٴ���������_Update','EXECUTE' From Dual Union All
Select '����ԤԼ�Һ�','Zl_�ٴ�����ԤԼ���Ʊ䶯','EXECUTE' From Dual Union All
Select '����ԤԼ�Һ�','Zl_�ٴ�������ſ���_Update','EXECUTE' From Dual Union All
Select 'ͣ������','Zl_�ٴ�����ͣ��_Apply','EXECUTE' From Dual Union All
Select 'ͣ������','Zl_�ٴ�����ͣ��_Audit','EXECUTE' From Dual Union All
Select ����,����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,������,����,����,Ȩ��)
Select &n_System,1115,User,A.* From (
Select ����,����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All
Select 'ͣ����Ϣ����','Zl_���߷�������_����','EXECUTE' From Dual Union All
Select 'ͣ����Ϣ����','Zl_���߷�������_����','EXECUTE' From Dual Union All
Select 'ͣ����Ϣ����','Zl_���߷�������_����','EXECUTE' From Dual Union All
Select 'ͣ����Ϣ����','zl_���˹Һż�¼_����_DELETE','EXECUTE' From Dual Union All
Select 'ԤԼ�Ǽ���Ϣ����','Zl_���߷�������_����','EXECUTE' From Dual Union All
Select 'ԤԼ�Ǽ���Ϣ����','Zl_���߷�������_����','EXECUTE' From Dual Union All
Select 'ԤԼ�Ǽ���Ϣ����','Zl_���߷�������_����','EXECUTE' From Dual Union All
Select 'ԤԼ�Ǽ���Ϣ����','zl_���˹Һż�¼_����_DELETE','EXECUTE' From Dual Union All
Select '����','Zl1_Fun_Getreturnvisit','EXECUTE' From Dual Union All
Select '����','Zl_���߷�������_����','EXECUTE' From Dual Union All
Select '����','������ü�¼','SELECT' From Dual Union All
Select '����','����Ԥ����¼','SELECT' From Dual Union All
Select '����','���㷽ʽ','SELECT' From Dual Union All
Select '����','���˹Һż�¼','SELECT' From Dual Union All
Select '����','ҽ�Ƹ��ʽ','SELECT' From Dual Union All
Select '����','���˷�����Ϣ��¼','SELECT' From Dual Union All
Select '����','������Ϣ','SELECT' From Dual Union All
Select '����','���ű�','SELECT' From Dual Union All
Select '����','�շ���ĿĿ¼','SELECT' From Dual Union All
Select '����','�ٴ������¼','SELECT' From Dual Union All
Select '����','�ٴ������Դ','SELECT' From Dual Union All
Select '����','�շѼ�Ŀ','SELECT' From Dual Union All
Select '����','������Ŀ','SELECT' From Dual Union All
Select '����','�շѴ�����Ŀ','SELECT' From Dual Union All
Select '����','�ٴ�������ſ���','SELECT' From Dual Union All
Select '����','�շ��ض���Ŀ','SELECT' From Dual Union All
Select '����','����ǼǼ�¼','SELECT' From Dual Union All
Select '����','��Ա��','SELECT' From Dual Union All
Select '����','����䶯��¼','SELECT' From Dual Union All
Select ����,����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,������,����,����,Ȩ��)
Select &n_System,1111,User,A.* From (
Select ����,����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All
Select 'ȡ��ԤԼ','Zl_���˹Һż�¼_����_Delete','EXECUTE' From Dual Union All
Select '�˺�','Zl_���˹Һż�¼_����_Delete','EXECUTE' From Dual Union All
Select '�Һ�','Zl_���˹Һż�¼_����_Insert','EXECUTE' From Dual Union All
Select 'ԤԼ�Һ�','Zl_���˹Һż�¼_����_Insert','EXECUTE' From Dual Union All
Select '����','Zl_�Һ����״̬_����_Delete','EXECUTE' From Dual Union All
Select ����,����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 9000, 'ԤԼ�Ǽ�', User, 'Zl_����ԤԼ�Ǽ�_Insert', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 9000 And ���� = 'ԤԼ�Ǽ�' And Upper(����) = Upper('Zl_����ԤԼ�Ǽ�_Insert'));

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 9000, 'ԤԼ�Ǽ�', User, 'ZL_���߷�������_����', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 9000 And ���� = 'ԤԼ�Ǽ�' And Upper(����) = Upper('ZL_���߷�������_����'));

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 9000, '����', User, 'Zl_�Һ����״̬_����_Delete', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 9000 And ���� = '����' And Upper(����) = Upper('Zl_�Һ����״̬_����_Delete'));

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1111,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All 
Select '�ٴ������¼','SELECT' From Dual Union All
Select '�ٴ������Դ','SELECT' From Dual Union All
Select '�ٴ�����Һſ��Ƽ�¼','SELECT' From Dual Union All
Select '�ٴ��������Ҽ�¼','SELECT' From Dual Union All
Select '�ٴ�������ſ���','SELECT' From Dual Union All
Select '�ٴ����ﰲ��','SELECT' From Dual Union All
Select '�ٴ������','SELECT' From Dual Union All
Select 'Zl1_Auto_Buildingregisterplan','EXECUTE' From Dual Union All
Select 'Zl_ԤԼ��ʽ_Check','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1113,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All 
Select '�ٴ������¼','SELECT' From Dual Union All
Select '�ٴ������Դ','SELECT' From Dual Union All
Select '�ٴ�����Һſ��Ƽ�¼','SELECT' From Dual Union All
Select '�ٴ��������Ҽ�¼','SELECT' From Dual Union All
Select '�ٴ�������ſ���','SELECT' From Dual Union All
Select '�ٴ����ﰲ��','SELECT' From Dual Union All
Select '�ٴ������','SELECT' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,9000,'�Һ�',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All 
Select '�ٴ������¼','SELECT' From Dual Union All
Select '�ٴ������Դ','SELECT' From Dual Union All
Select '�ٴ�����Һſ��Ƽ�¼','SELECT' From Dual Union All
Select '�ٴ��������Ҽ�¼','SELECT' From Dual Union All
Select '�ٴ�������ſ���','SELECT' From Dual Union All
Select '�ٴ����ﰲ��','SELECT' From Dual Union All
Select '�ٴ������','SELECT' From Dual Union All
Select 'Zl1_Auto_Buildingregisterplan','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,9000,'ԤԼ',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All 
Select '�ٴ������¼','SELECT' From Dual Union All
Select '�ٴ������Դ','SELECT' From Dual Union All
Select '�ٴ�����Һſ��Ƽ�¼','SELECT' From Dual Union All
Select '�ٴ��������Ҽ�¼','SELECT' From Dual Union All
Select '�ٴ�������ſ���','SELECT' From Dual Union All
Select '�ٴ����ﰲ��','SELECT' From Dual Union All
Select '�ٴ������','SELECT' From Dual Union All
Select 'Zl1_Auto_Buildingregisterplan','EXECUTE' From Dual Union All
Select 'Zl_ԤԼ��ʽ_Check','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1539,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All 
Select '�ٴ������¼','SELECT' From Dual Union All
Select '�ٴ������Դ','SELECT' From Dual Union All
Select '����','SELECT' From Dual Union All
Select '�ٴ��������Ҽ�¼','SELECT' From Dual Union All
Select '�ٴ�������ſ���','SELECT' From Dual Union All
Select '�շ���Ŀ���','SELECT' From Dual Union All
Select '�շ���ĿĿ¼','SELECT' From Dual Union All
Select 'Zl1_Auto_Buildingregisterplan','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1802,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All 
Select '�ٴ������¼','SELECT' From Dual Union All
Select '�ٴ������Դ','SELECT' From Dual Union All
Select '�ٴ��������Ҽ�¼','SELECT' From Dual Union All
Select '�ٴ�������ſ���','SELECT' From Dual Union All
Select 'Zl1_Auto_Buildingregisterplan','EXECUTE' From Dual Union All
Select 'Zl_���˹Һż�¼_����_Insert','EXECUTE' From Dual Union All
Select 'Zl_ԤԼ�ҺŽ���_����_Insert','EXECUTE' From Dual Union All
Select 'NextReservationNum','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1803,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All 
Select '�ٴ������¼','SELECT' From Dual Union All
Select '�ٴ������Դ','SELECT' From Dual Union All
Select '�ٴ��������Ҽ�¼','SELECT' From Dual Union All
Select '�ٴ�������ſ���','SELECT' From Dual Union All
Select 'Zl1_Auto_Buildingregisterplan','EXECUTE' From Dual Union All
Select 'Zl_���˹Һż�¼_����_Insert','EXECUTE' From Dual Union All
Select 'Zl_ԤԼ�ҺŽ���_����_Insert','EXECUTE' From Dual Union All
Select 'NextReservationNum','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into Zlprocedure
  (ID, ����, ����, ״̬, ������, ˵��)
  Select Zlprocedure_Id.Nextval, 2, 'Zl1_Ex_Isdoctorsamelevel', 3, User, '�Ƚ�����ҽ����ְ���С' From Dual;
  
Insert Into zlAutoJobs
  (ϵͳ, ����, ���, ����, ˵��, ����, ����, ִ��ʱ��, ���ʱ��)
  Select &n_System, 1, 11, '�����¼�Զ�����', '���̶�ʱ����ɶԹҺŹ̶����ų����¼�Զ����ɡ�', 'Zl1_Auto_BuildingRegisterPlan', Null,
         Trunc(Sysdate) + 1 / 24, 1
  From Dual
  Where Not Exists (Select 1 From zlAutoJobs Where ϵͳ = &n_System And ���� = 1 And ��� = 11);

Insert Into zlBaseCode
  (ϵͳ, ����, �̶�, ˵��, ����)
Values
  (&n_System, '����ͣ��ԭ��', 0, '�ٴ����ﰲ�ŵĳ���ͣ��ԭ��', 'ҽ�ƹ���');
  
--���ݴ���
Insert Into zlProgFuncs
  (ϵͳ, ���, ����, ����, ˵��, ȱʡֵ)
  Select ϵͳ, ���, '���շѺ�', ����, '�����˹��շѺŵĲ���Ȩ�ޡ��и�Ȩ��ʱ������Բ��˽��йҷ��ò�Ϊ0�ĺ�', ȱʡֵ
  From zlProgFuncs
  Where ��� = 1111 And ϵͳ = &n_System And ���� = '�Һ�';

Update Zlprogrelas Set ���� = '���շѺ�' Where ��� = 1111 And ϵͳ = &n_System And ���� = '�Һ�';

Update zlProgPrivs Set ���� = '���շѺ�' Where ��� = 1111 And ϵͳ = &n_System And ���� = '�Һ�';

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ����, ������, Ȩ��)
  Select ϵͳ, ���, '����Ѻ�', ����, ������, Ȩ��
  From zlProgPrivs
  Where ��� = 1111 And ϵͳ = &n_System And ���� = '���շѺ�';

Update zlRoleGrant Set ���� = '���շѺ�' Where ��� = 1111 And ϵͳ = &n_System And ���� = '�Һ�';

Delete From zlProgFuncs Where ��� = 1111 And ϵͳ = &n_System And ���� = '�Һ�';


  --���̽ű�
Create Or Replace Procedure Zl_�ϰ�ʱ��_Modify
(
  ��������_In     Number,
  վ��_In         ʱ���.վ��%Type,
  ����_In         ʱ���.����%Type,
  ʱ���_In       ʱ���.ʱ���%Type,
  ��ʼʱ��_In     ʱ���.��ʼʱ��%Type,
  ��ֹʱ��_In     ʱ���.��ֹʱ��%Type,
  ��Ϣʱ��_In     ʱ���.��Ϣʱ��%Type,
  ȱʡʱ��_In     ʱ���.ȱʡʱ��%Type,
  ��ǰʱ��_In     ʱ���.��ǰʱ��%Type,
  ����Ԥ��ʱ��_In ʱ���.����Ԥ��ʱ��%Type,
  ԭվ��_In       ʱ���.վ��%Type := Null,
  ԭ����_In       ʱ���.����%Type := Null,
  ԭʱ���_In     ʱ���.ʱ���%Type := Null
) As
  --�������޸��ϰ�ʱ��
  --��������_In 0-������1-�޸�
  --ԭվ��_In��ԭ����_In��ԭʱ���_In �޸�ʱ����
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_Count Number;
Begin
  If Nvl(��������_In, 0) = 0 Then
    --�����ϰ�ʱ��
    Begin
      Select 1
      Into n_Count
      From ʱ���
      Where Nvl(վ��, '-') = Nvl(վ��_In, '-') And Nvl(����, '-') = Nvl(����_In, '-') And ʱ��� = ʱ���_In;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count > 0 Then
      v_Err_Msg := '��ǰվ���Ѵ�����ͬ������ϰ�ʱ��Ρ�' || ʱ���_In || '����';
      Raise Err_Item;
    End If;
  
    Insert Into ʱ���
      (վ��, ����, ʱ���, ��ʼʱ��, ��ֹʱ��, ��Ϣʱ��, ȱʡʱ��, ��ǰʱ��, ����Ԥ��ʱ��)
    Values
      (վ��_In, ����_In, ʱ���_In, ��ʼʱ��_In, ��ֹʱ��_In, ��Ϣʱ��_In, Nvl(ȱʡʱ��_In, ��ʼʱ��_In), Nvl(��ǰʱ��_In, ��ʼʱ��_In), ����Ԥ��ʱ��_In);
    Return;
  End If;

  --�޸�ʱ�����ԭ�ϰ�ʱ���Ƿ�ʹ�ã���ʹ�õĲ����޸�վ�㡢���ࡢʱ���
  --����ɾ����ʹ�õķ�Χ������һ��,��ʹ�õ�ʱ��ֻҪ��һ�����ɣ���ͬվ�㣬��ͬ������ܻ��ж��ͬ����ʱ��Σ�

  If Nvl(ԭվ��_In, '-') <> Nvl(վ��_In, '-') Or Nvl(ԭ����_In, '-') <> Nvl(����_In, '-') Or ԭʱ���_In <> ʱ���_In Then
    --�ٴ������Դ����
    Begin
      Select 1
      Into n_Count
      From (Select b.�ϰ�ʱ��, c.վ��, a.����, Row_Number() Over(Partition By b.�ϰ�ʱ�� Order By b.�ϰ�ʱ��, c.վ�� Desc, a.���� Desc) As ���
             From �ٴ������Դ A, �ٴ������Դ���� B, ���ű� C
             Where a.Id = b.��Դid And a.����id = c.Id)
      Where ��� = 1 And Nvl(վ��, '-') = Nvl(ԭվ��_In, '-') And Nvl(����, '-') = Nvl(ԭ����_In, '-') And �ϰ�ʱ�� = ԭʱ���_In And
            Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    --�ٴ���������(�̶�����ģ��)
    If Nvl(n_Count, 0) = 0 Then
      Begin
        Select 1
        Into n_Count
        From (Select a.�ϰ�ʱ��, c.վ��, b.����,
                      Row_Number() Over(Partition By a.�ϰ�ʱ�� Order By a.�ϰ�ʱ��, c.վ�� Desc, b.���� Desc) As ���
               From �ٴ��������� A, �ٴ����ﰲ�� D, �ٴ������Դ B, ���ű� C
               Where a.����id = d.Id And d.��Դid = b.Id And b.����id = c.Id)
        Where ��� = 1 And Nvl(վ��, '-') = Nvl(ԭվ��_In, '-') And Nvl(����, '-') = Nvl(ԭ����_In, '-') And �ϰ�ʱ�� = ԭʱ���_In And
              Rownum < 2;
      Exception
        When Others Then
          n_Count := 0;
      End;
    End If;
    --�ٴ������¼
    --����飬��Ϊ�ñ�̫������ϰ�ʱ�ε���Ϣ����������������У�û���ҵ��ϰ�ʱ��ʱ������������������ȡ
    If n_Count > 0 Then
      v_Err_Msg := '�ϰ�ʱ��Ρ�' || ԭʱ���_In || '���ѱ�ʹ�ã������޸���վ�㡢���༰ʱ������ƣ�';
      Raise Err_Item;
    End If;
  
    Begin
      Select 1
      Into n_Count
      From ʱ���
      Where Nvl(վ��, '-') = Nvl(վ��_In, '-') And Nvl(����, '-') = Nvl(����_In, '-') And ʱ��� = ʱ���_In;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count > 0 Then
      v_Err_Msg := '��ǰվ���Ѵ�����ͬ������ϰ�ʱ��Ρ�' || ʱ���_In || '����';
      Raise Err_Item;
    End If;
  End If;

  Update ʱ���
  Set վ�� = վ��_In, ���� = ����_In, ʱ��� = ʱ���_In, ��ʼʱ�� = ��ʼʱ��_In, ��ֹʱ�� = ��ֹʱ��_In, ��Ϣʱ�� = ��Ϣʱ��_In, ȱʡʱ�� = Nvl(ȱʡʱ��_In, ��ʼʱ��_In),
      ��ǰʱ�� = Nvl(��ǰʱ��_In, ��ʼʱ��_In), ����Ԥ��ʱ�� = ����Ԥ��ʱ��_In
  Where Nvl(վ��, '-') = Nvl(ԭվ��_In, '-') And Nvl(����, '-') = Nvl(ԭ����_In, '-') And ʱ��� = ԭʱ���_In;
  If Sql%NotFound Then
    Insert Into ʱ���
      (վ��, ����, ʱ���, ��ʼʱ��, ��ֹʱ��, ��Ϣʱ��, ȱʡʱ��, ��ǰʱ��, ����Ԥ��ʱ��)
    Values
      (վ��_In, ����_In, ʱ���_In, ��ʼʱ��_In, ��ֹʱ��_In, ��Ϣʱ��_In, Nvl(ȱʡʱ��_In, ��ʼʱ��_In), Nvl(��ǰʱ��_In, ��ʼʱ��_In), ����Ԥ��ʱ��_In);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ϰ�ʱ��_Modify;
/
Create Or Replace Procedure Zl_�ϰ�ʱ��_Delete
(
  վ��_In   ʱ���.վ��%Type,
  ����_In   ʱ���.����%Type,
  ʱ���_In ʱ���.ʱ���%Type
) As
  -- ɾ���ϰ�ʱ��
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_Count Number;
Begin
  --���ݼ�飬�ϰ�ʱ����ѱ�ʹ������ɾ��
  --����ɾ����ʹ�õķ�Χ������һ��,��ʹ�õ�ʱ��ֻҪ��һ�����ɣ���ͬվ�㣬��ͬ������ܻ��ж��ͬ����ʱ��Σ�

  --�ٴ������Դ����
  Begin
    Select 1
    Into n_Count
    From (Select b.�ϰ�ʱ��, c.վ��, a.����, Row_Number() Over(Partition By b.�ϰ�ʱ�� Order By b.�ϰ�ʱ��, c.վ�� Desc, a.���� Desc) As ���
           From �ٴ������Դ A, �ٴ������Դ���� B, ���ű� C
           Where a.Id = b.��Դid And a.����id = c.Id)
    Where ��� = 1 And Nvl(վ��, '-') = Nvl(վ��_In, '-') And Nvl(����, '-') = Nvl(����_In, '-') And �ϰ�ʱ�� = ʱ���_In And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;
  --�ٴ���������(�̶�����ģ��)
  If Nvl(n_Count, 0) = 0 Then
    Begin
      Select 1
      Into n_Count
      From (Select a.�ϰ�ʱ��, c.վ��, b.����, Row_Number() Over(Partition By a.�ϰ�ʱ�� Order By a.�ϰ�ʱ��, c.վ�� Desc, b.���� Desc) As ���
             From �ٴ��������� A, �ٴ����ﰲ�� D, �ٴ������Դ B, ���ű� C
             Where a.����id = d.Id And d.��Դid = b.Id And b.����id = c.Id)
      Where ��� = 1 And Nvl(վ��, '-') = Nvl(վ��_In, '-') And Nvl(����, '-') = Nvl(����_In, '-') And �ϰ�ʱ�� = ʱ���_In And
            Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
  End If;
  --�ٴ������¼
  --����飬��Ϊ�ñ�̫������ϰ�ʱ�ε���Ϣ����������������У�û���ҵ��ϰ�ʱ��ʱ������������������ȡ

  If n_Count > 0 Then
    v_Err_Msg := '��ǰ�ϰ�ʱ����ѱ�ʹ�ã�����ɾ����';
    Raise Err_Item;
  End If;

  Delete From ʱ���
  Where Nvl(վ��, '-') = Nvl(վ��_In, '-') And Nvl(����, '-') = Nvl(����_In, '-') And ʱ��� = ʱ���_In;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ϰ�ʱ��_Delete;
/

Create Or Replace Procedure Zl_�������ձ�_Modify
(
  ��������_In Number,
  ���_In     �������ձ�.���%Type,
  ��������_In �������ձ�.��������%Type,
  ��ʼ����_In �������ձ�.��ʼ����%Type,
  ��ֹ����_In �������ձ�.��ֹ����%Type,
  ��ע_In     �������ձ�.��ע%Type,
  �������_In Varchar2 := Null,
  ����ԤԼ_In Varchar2 := Null,
  ����Һ�_In Varchar2 := Null
) As
  --�������޸ķ����ڼ���
  --      ��������_In 0-������1-�޸�
  --      �������_In ��ʽ������ʱ��1~ ԭ�ϰ�ʱ��1;����ʱ��2~ ԭ�ϰ�ʱ��2;
  --      ����ԤԼ_in ����ԤԼ������,��ʽ��yyyy-mm-dd;yyyy-mm-dd;...
  --      ����Һ�_in ����Һŵ�����,��ʽ��yyyy-mm-dd;yyyy-mm-dd;...
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_Count Number;

  v_������� Varchar2(4000);
  v_��ǰ��Ŀ Varchar2(4000);
  d_��ʼ���� Date;
  d_��ֹ���� Date;
Begin
  If ��������_In = 0 Then
    --����
    Begin
      Select 1
      Into n_Count
      From �������ձ�
      Where ���� = 0 And ��� = ���_In And �������� = ��������_In And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count > 0 Then
      v_Err_Msg := ���_In || '���Ѵ��ڡ�' || ��������_In || '����';
      Raise Err_Item;
    End If;
  
    Begin
      Select 1
      Into n_Count
      From �ٴ������¼
      Where �������� Between ��ʼ����_In And ��ֹ����_In And Nvl(�Ƿ񷢲�, 0) = 1 And (Nvl(��Լ��, 0) <> 0 Or Nvl(�ѹ���, 0) <> 0) And
            Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count > 0 Then
      v_Err_Msg := '��ǰ�ڼ��յ�ʱ�䷶Χ������ԤԼ�ҺŲ��ˣ��������ã�';
      Raise Err_Item;
    End If;
  
    Insert Into �������ձ�
      (���, ��������, ����, ��ʼ����, ��ֹ����, ��ע, ����ԤԼ, ����Һ�)
    Values
      (���_In, ��������_In, 0, ��ʼ����_In, ��ֹ����_In, ��ע_In, ����ԤԼ_In, ����Һ�_In);
  
    If �������_In Is Not Null Then
      v_������� := �������_In || ';';
    End If;
    While v_������� Is Not Null Loop
      v_��ǰ��Ŀ := Substr(v_�������, 0, Instr(v_�������, ';') - 1);
      d_��ʼ���� := To_Date(Substr(v_��ǰ��Ŀ, 0, Instr(v_��ǰ��Ŀ, '~') - 1), 'yyyy-mm-dd');
      d_��ֹ���� := To_Date(Substr(v_��ǰ��Ŀ, Instr(v_��ǰ��Ŀ, '~') + 1), 'yyyy-mm-dd');
    
      Insert Into �������ձ�
        (���, ��������, ����, ��ʼ����, ��ֹ����, ��ע)
      Values
        (���_In, ��������_In, 1, d_��ʼ����, d_��ֹ����, Null);
    
      v_������� := Substr(v_�������, Instr(v_�������, ';') + 1);
    End Loop;
  
  Elsif ��������_In = 1 Then
    --�޸�
    Begin
      Select ��ʼ����
      Into d_��ʼ����
      From �������ձ�
      Where ���� = 0 And ��� = ���_In And �������� = ��������_In And Rownum < 2;
    Exception
      When Others Then
        d_��ʼ���� := Null;
    End;
    If d_��ʼ���� Is Null Then
      v_Err_Msg := ���_In || '�겻���ڡ�' || ��������_In || '����';
      Raise Err_Item;
    End If;
  
    If Sysdate > d_��ʼ���� Then
      v_Err_Msg := '��ǰʱ���Ѿ������˽ڼ��տ�ʼʱ�䣬�����޸ģ�';
      Raise Err_Item;
    End If;
  
    Begin
      Select 1
      Into n_Count
      From �ٴ������¼
      Where �������� Between ��ʼ����_In And ��ֹ����_In And Nvl(�Ƿ񷢲�, 0) = 1 And (Nvl(��Լ��, 0) <> 0 Or Nvl(�ѹ���, 0) <> 0) And
            Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count > 0 Then
      v_Err_Msg := '��ǰ�ڼ��յ�ʱ�䷶Χ������ԤԼ�ҺŲ��ˣ������޸ģ�';
      Raise Err_Item;
    End If;
  
    Update �������ձ�
    Set ��ʼ���� = ��ʼ����_In, ��ֹ���� = ��ֹ����_In, ��ע = ��ע_In, ����ԤԼ = ����ԤԼ_In, ����Һ� = ����Һ�_In
    Where ��� = ���_In And Nvl(����, 0) = 0 And �������� = ��������_In;
  
    --��ɾ����������
    Delete From �������ձ� Where ��� = ���_In And Nvl(����, 0) = 1 And �������� = ��������_In;
    If �������_In Is Not Null Then
      v_������� := �������_In || ';';
    End If;
    While v_������� Is Not Null Loop
      v_��ǰ��Ŀ := Substr(v_�������, 0, Instr(v_�������, ';') - 1);
      d_��ʼ���� := To_Date(Substr(v_��ǰ��Ŀ, 0, Instr(v_��ǰ��Ŀ, '~') - 1), 'yyyy-mm-dd');
      d_��ֹ���� := To_Date(Substr(v_��ǰ��Ŀ, Instr(v_��ǰ��Ŀ, '~') + 1), 'yyyy-mm-dd');
    
      Insert Into �������ձ�
        (���, ��������, ����, ��ʼ����, ��ֹ����, ��ע)
      Values
        (���_In, ��������_In, 1, d_��ʼ����, d_��ֹ����, Null);
    
      v_������� := Substr(v_�������, Instr(v_�������, ';') + 1);
    End Loop;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�������ձ�_Modify;
/
Create Or Replace Procedure Zl_�������ձ�_Delete
(
  ���_In     �������ձ�.���%Type,
  ��������_In �������ձ�.��������%Type
) As
  --ɾ�������ڼ���
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  d_��ʼ���� Date;
Begin
  Begin
    Select ��ʼ����
    Into d_��ʼ����
    From �������ձ�
    Where ���� = 0 And ��� = ���_In And �������� = ��������_In And Rownum < 2;
  Exception
    When Others Then
      d_��ʼ���� := Null;
  End;
  If d_��ʼ���� Is Null Then
    v_Err_Msg := ���_In || '�겻���ڡ�' || ��������_In || '����';
    Raise Err_Item;
  End If;

  If Sysdate > d_��ʼ���� Then
    v_Err_Msg := '��ǰʱ���Ѿ������˽ڼ��տ�ʼʱ�䣬�����޸ģ�';
    Raise Err_Item;
  End If;

  Delete From �������ձ� Where ��� = ���_In And �������� = ��������_In;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�������ձ�_Delete;
/
Create Or Replace Procedure Zl_��������_Modify
(
  ��������_In Number,
  Id_In       ��������.Id%Type,
  ����_In     ��������.����%Type := Null,
  ����_In     ��������.����%Type := Null,
  ����_In     ��������.����%Type := Null,
  λ��_In     ��������.λ��%Type := Null,
  վ��_In     ��������.վ��%Type := Null,
  ���ÿ���_In Varchar2 := Null
) As
  --�������޸���������
  --��������_In 0-������1-�޸�
  --���ÿ���_In ��ʽ������ID����ʽ������1;����2;����3;...
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_Id    ��������.Id%Type;
  n_Count Number;
Begin
  If ��������_In = 0 Then
    --����
    Begin
      Select 1 Into n_Count From �������� Where ���� = ����_In And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count > 0 Then
      v_Err_Msg := ����_In || ' �Ѵ��ڣ�';
      Raise Err_Item;
    End If;
  
    Select ��������_Id.Nextval Into n_Id From Dual;
    Insert Into �������� (ID, ����, ����, ����, λ��, վ��) Values (n_Id, ����_In, ����_In, ����_In, λ��_In, վ��_In);
  
    --���������������ÿ���
    If Not ���ÿ���_In Is Null Then
      Insert Into �����������ÿ���
        (����id, ����id)
        Select n_Id, Column_Value As ����id From Table(f_Num2list(���ÿ���_In, ';'));
    End If;
  
    Return;
  End If;

  --�޸�
  Begin
    Select 1 Into n_Count From �������� Where ���� = ����_In And ID <> Id_In And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count > 0 Then
    v_Err_Msg := ����_In || ' �Ѵ��ڣ�';
    Raise Err_Item;
  End If;

  Update �������� Set ���� = ����_In, ���� = ����_In, ���� = ����_In, λ�� = λ��_In, վ�� = վ��_In Where ID = Id_In;

  --��ɾ��
  Delete From �����������ÿ��� Where ����id = Id_In;
  --���������������ÿ���
  If Not ���ÿ���_In Is Null Then
    Insert Into �����������ÿ���
      (����id, ����id)
      Select Id_In, Column_Value As ����id From Table(f_Num2list(���ÿ���_In, ';'));
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��������_Modify;
/
Create Or Replace Procedure Zl_��������_Delete(Id_In ��������.Id%Type) As
  --ɾ����������
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_Count Number(2);
Begin
  Begin
    Select 1 Into n_Count From �ٴ������Դ���� Where ����id = Id_In And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;

  If Nvl(n_Count, 0) = 0 Then
    Begin
      Select 1 Into n_Count From �ٴ��������� Where ����id = Id_In And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
  End If;

  If Nvl(n_Count, 0) = 0 Then
    Begin
      Select 1 Into n_Count From �ٴ��������Ҽ�¼ Where ����id = Id_In And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
  End If;
  If Nvl(n_Count, 0) > 0 Then
    v_Err_Msg := '��ǰ�����ѱ�ʹ�ã�����ɾ����';
    Raise Err_Item;
  End If;

  Delete From �����������ÿ��� Where ����id = Id_In;
  Delete From �������� Where ID = Id_In;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��������_Delete;
/
Create Or Replace Procedure Zl_�ٴ������Դ_Stopandstart
(
  Id_In   �ٴ������Դ.Id%Type,
  ͣ��_In Number := 0
) As
Begin
  If Nvl(ͣ��_In, 0) = 1 Then
    Update �ٴ������Դ Set ����ʱ�� = Sysdate Where ID = Id_In;
  Else
    Update �ٴ������Դ Set ����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd'), �Ƿ�ɾ�� = 0 Where ID = Id_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ������Դ_Stopandstart;
/
Create Or Replace Procedure Zl_�ٴ������Դ_Modify
(
  ��������_In     Number,
  Id_In           �ٴ������Դ.Id%Type,
  ����_In         �ٴ������Դ.����%Type := Null,
  ����_In         �ٴ������Դ.����%Type := Null,
  ����id_In       �ٴ������Դ.����id%Type := 0,
  ��Ŀid_In       �ٴ������Դ.��Ŀid%Type := 0,
  ҽ��id_In       �ٴ������Դ.ҽ��id%Type := Null,
  ҽ������_In     �ٴ������Դ.ҽ������%Type := Null,
  �Ƿ񽨲���_In   �ٴ������Դ.�Ƿ񽨲���%Type := 0,
  ԤԼ����_In     �ٴ������Դ.ԤԼ����%Type := 0,
  ����Ƶ��_In     �ٴ������Դ.����Ƶ��%Type := 0,
  ���տ���״̬_In �ٴ������Դ.���տ���״̬%Type := 0,
  �Ƿ���ջ���_In �ٴ������Դ.�Ƿ���ջ���%Type := 0,
  �Ƿ��ٴ��Ű�_In �ٴ������Դ.�Ƿ��ٴ��Ű�%Type := 0,
  �Ű෽ʽ_In     �ٴ������Դ.�Ű෽ʽ%Type := 0
) As
  --��������_In 0-������1-�޸ģ�2-ɾ��
  --��������_In ����ID����ʽ������ID1;����ID2;����ID13;...
  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  n_��Դid �ٴ������Դ.Id%Type;
  n_Count  Number;
Begin

  If ��������_In = 0 Then
    --���Ӻ�Դ
    n_��Դid := Id_In;
  
    If Nvl(n_��Դid, 0) = 0 Then
      Select �ٴ������Դ_Id.Nextval Into n_��Դid From Dual;
    End If;
    Insert Into �ٴ������Դ
      (ID, ����, ����, ����id, ��Ŀid, ҽ��id, ҽ������, �Ƿ񽨲���, ԤԼ����, ����Ƶ��, ���տ���״̬, �Ƿ���ջ���, �Ƿ��ٴ��Ű�, �Ű෽ʽ, �Ƿ�ɾ��, ����ʱ��, ����ʱ��)
    Values
      (n_��Դid, ����_In, ����_In, ����id_In, ��Ŀid_In, ҽ��id_In, ҽ������_In, �Ƿ񽨲���_In, ԤԼ����_In, ����Ƶ��_In, ���տ���״̬_In, �Ƿ���ջ���_In,
       �Ƿ��ٴ��Ű�_In, �Ű෽ʽ_In, 0, Sysdate, To_Date('3000-01-01', 'yyyy-mm-dd'));
  
    Return;
  End If;

  --�޸ĺ�Դ
  Update �ٴ������Դ
  Set ���� = ����_In, ���� = ����_In, ����id = ����id_In, ��Ŀid = ��Ŀid_In, ҽ��id = ҽ��id_In, ҽ������ = ҽ������_In, �Ƿ񽨲��� = �Ƿ񽨲���_In,
      ԤԼ���� = ԤԼ����_In, ����Ƶ�� = ����Ƶ��_In, ���տ���״̬ = ���տ���״̬_In, �Ƿ���ջ��� = �Ƿ���ջ���_In, �Ƿ��ٴ��Ű� = �Ƿ��ٴ��Ű�_In, �Ű෽ʽ = �Ű෽ʽ_In
  Where ID = Id_In And Nvl(�Ƿ�ɾ��, 0) = 0 And Nvl(����ʱ��, Sysdate) >= Sysdate;
  If Sql%NotFound Then
    v_Err_Msg := '��ǰ��Դ�����ѱ�����ɾ����ͣ�ã����ܶԸú�Դ��Ϣ�����޸�!';
    Raise Err_Item;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ������Դ_Modify;
/

Create Or Replace Procedure Zl_�ٴ������Դ_Delete(Id_In �ٴ������Դ.Id%Type) As
  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  n_Count  Number;
  l_����id t_Numlist := t_Numlist();
Begin
  Select Count(1) Into n_Count From �ٴ����ﰲ�� Where ��Դid = Id_In;

  If n_Count = 0 Then
  
    Select ID Bulk Collect Into l_����id From �ٴ������Դ���� Where ��Դid = Id_In;
  
    Forall I In 1 .. l_����id.Count
      Delete �ٴ������Դʱ�� Where ����id = l_����id(I);
  
    Forall I In 1 .. l_����id.Count
      Delete �ٴ������Դ���� Where ����id = l_����id(I);
  
    Forall I In 1 .. l_����id.Count
      Delete �ٴ������Դ���� Where ����id = l_����id(I);
  
    Delete �ٴ������Դ���� Where ��Դid = Id_In;
    --��ɾ��
  
    Delete From �ٴ������Դ Where ID = Id_In;
    If Sql%NotFound Then
      v_Err_Msg := '��ǰ��Դ�����ѱ�����ɾ����������ɾ��!';
      Raise Err_Item;
    End If;
    Return;
  End If;
  Update �ٴ������Դ Set �Ƿ�ɾ�� = 1, ����ʱ�� = Sysdate Where ID = Id_In And Nvl(�Ƿ�ɾ��, 0) = 0;
  If Sql%NotFound Then
    v_Err_Msg := '��ǰ��Դ�����ѱ�����ɾ����������ɾ��!';
    Raise Err_Item;
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ������Դ_Delete;
/

Create Or Replace Procedure Zl_�ٴ������Դ����_Modify
(
  Id_In           �ٴ������Դ����.Id%Type,
  ��Դid_In       �ٴ������Դ����.��Դid%Type,
  �ϰ�ʱ��_In     �ٴ������Դ����.�ϰ�ʱ��%Type,
  �޺���_In       �ٴ������Դ����.�޺���%Type,
  ��Լ��_In       �ٴ������Դ����.��Լ��%Type,
  �Ƿ���ſ���_In �ٴ������Դ����.�Ƿ���ſ���%Type,
  �Ƿ��ʱ��_In   �ٴ������Դ����.�Ƿ��ʱ��%Type,
  ԤԼ����_In     �ٴ������Դ����.ԤԼ����%Type,
  �Ƿ��ռ_In     �ٴ������Դ����.�Ƿ��ռ%Type,
  ���﷽ʽ_In     �ٴ������Դ����.���﷽ʽ%Type,
  ����id_In       �ٴ������Դ����.����id%Type,
  ��Դ����_In     Varchar2 := Null,
  ��Դʱ��_In     Varchar2 := Null,
  ��Դ����_In     Varchar2 := Null,
  ɾ����Դ����_In Integer := 0
  
) As
  --��Դʱ��_IN:���,��ʼʱ��(HH:MM:SS),��ֹʱ(HH:MM:SS)��,����,�Ƿ�ԤԼ|...
  --��Դ����_IN:����id1,����id2,....
  --��Դ����_IN:����,����,����,���Ʒ�ʽ,���,����|
  --ɾ����Դ����_in:1-��������ǰ����ɾ����Դ����,0-��ɾ�����ݣ�ֱ�Ӳ���

  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  l_����id   t_Numlist := t_Numlist();
  n_Count    Number;
  v_��ʼʱ�� Varchar2(20);
  v_��ֹʱ�� Varchar2(20);

  n_���     �ٴ������Դʱ��.���%Type;
  d_��ʼʱ�� �ٴ������Դʱ��.��ʼʱ��%Type;
  d_��ֹʱ�� �ٴ������Դʱ��.��ֹʱ��%Type;
  n_����     �ٴ������Դʱ��.��������%Type;
  n_�Ƿ�ԤԼ �ٴ������Դʱ��.�Ƿ�ԤԼ%Type;
  n_����     �ٴ������Դ����.����%Type;
  n_����     �ٴ������Դ����.����%Type;
  v_����     �ٴ������Դ����.����%Type;
  n_���Ʒ�ʽ �ٴ������Դ����.���Ʒ�ʽ%Type;
  n_�������� �ٴ������Դ����.����%Type;
Begin
  If Nvl(ɾ����Դ����_In, 0) = 1 Then
    Select ID Bulk Collect Into l_����id From �ٴ������Դ���� Where ��Դid = ��Դid_In;
    Forall I In 1 .. l_����id.Count
      Delete �ٴ������Դʱ�� Where ����id = l_����id(I);
  
    Forall I In 1 .. l_����id.Count
      Delete �ٴ������Դ���� Where ����id = l_����id(I);
  
    Forall I In 1 .. l_����id.Count
      Delete �ٴ������Դ���� Where ����id = l_����id(I);
  
    Delete �ٴ������Դ���� Where ��Դid = ��Դid_In;
    Delete From �ٴ������Դ���� Where ��Դid = ��Դid_In;
  
  End If;

  Select Count(1) Into n_Count From �ٴ������Դ���� Where ID = Id_In;
  If n_Count = 0 Then
    Insert Into �ٴ������Դ����
      (ID, ��Դid, �ϰ�ʱ��, �޺���, ��Լ��, �Ƿ���ſ���, �Ƿ��ʱ��, ԤԼ����, �Ƿ��ռ, ���﷽ʽ, ����id)
    Values
      (Id_In, ��Դid_In, �ϰ�ʱ��_In, �޺���_In, ��Լ��_In, �Ƿ���ſ���_In, �Ƿ��ʱ��_In, ԤԼ����_In, �Ƿ��ռ_In, ���﷽ʽ_In, ����id_In);
  
  End If;

  If ��Դʱ��_In Is Not Null Then
    --�����Դȱʡʱ���
    For c_ʱ��μ� In (Select Rownum As ���, Column_Value As ֵ From Table(f_Str2list(��Դʱ��_In, '|'))) Loop
      n_���     := Null;
      v_��ʼʱ�� := Null;
      v_��ֹʱ�� := Null;
      n_����     := Null;
      n_�Ƿ�ԤԼ := Null;
      For c_ʱ��� In (Select Rownum As ���, Column_Value As ֵ From Table(f_Str2list(c_ʱ��μ�.ֵ)) Order By ���) Loop
        If c_ʱ���.��� = 1 Then
          n_��� := To_Number(c_ʱ���.ֵ);
        End If;
      
        If c_ʱ���.��� = 2 Then
          v_��ʼʱ�� := c_ʱ���.ֵ;
        End If;
      
        If c_ʱ���.��� = 3 Then
          v_��ֹʱ�� := c_ʱ���.ֵ;
        End If;
      
        If c_ʱ���.��� = 4 Then
          n_���� := To_Number(c_ʱ���.ֵ);
        End If;
      
        If c_ʱ���.��� = 5 Then
          n_�Ƿ�ԤԼ := To_Number(c_ʱ���.ֵ);
        End If;
      
      End Loop;
      d_��ʼʱ�� := To_Date('3000-01-01 ' || Nvl(v_��ʼʱ��, ''), 'yyyy-mm-dd hh24:mi:ss');
      d_��ֹʱ�� := To_Date('3000-01-01 ' || Nvl(v_��ֹʱ��, ''), 'yyyy-mm-dd hh24:mi:ss');
    
      If d_��ʼʱ�� >= d_��ֹʱ�� Then
        d_��ֹʱ�� := d_��ֹʱ�� + 1;
      End If;
    
      If Nvl(n_���, 0) <> 0 Then
        Insert Into �ٴ������Դʱ��
          (����id, ���, ��ʼʱ��, ��ֹʱ��, ��������, �Ƿ�ԤԼ)
        Values
          (Id_In, n_���, d_��ʼʱ��, d_��ֹʱ��, n_����, n_�Ƿ�ԤԼ);
      End If;
    End Loop;
  
  End If;

  --�����Դ��ȱʡ����
  --��Դ����_IN:����,����,����,���Ʒ�ʽ,���,����|
  If ��Դ����_In Is Not Null Then
    For c_ʱ��μ� In (Select Rownum As ���, Column_Value As ֵ From Table(f_Str2list(��Դ����_In, '|'))) Loop
      n_����     := Null;
      n_����     := Null;
      v_����     := Null;
      n_���     := Null;
      n_���Ʒ�ʽ := Null;
      n_�������� := Null;
    
      --����,����,����,���Ʒ�ʽ,���,����|
      For c_ʱ��� In (Select Rownum As ���, Column_Value As ֵ From Table(f_Str2list(c_ʱ��μ�.ֵ)) Order By ���) Loop
        If c_ʱ���.��� = 1 Then
          n_���� := To_Number(c_ʱ���.ֵ);
        End If;
      
        If c_ʱ���.��� = 2 Then
          n_���� := To_Number(c_ʱ���.ֵ);
        End If;
      
        If c_ʱ���.��� = 3 Then
          v_���� := c_ʱ���.ֵ;
        End If;
      
        If c_ʱ���.��� = 4 Then
          n_���Ʒ�ʽ := To_Number(c_ʱ���.ֵ);
        End If;
      
        If c_ʱ���.��� = 5 Then
          n_��� := To_Number(c_ʱ���.ֵ);
        End If;
      
        If c_ʱ���.��� = 6 Then
          n_�������� := To_Number(c_ʱ���.ֵ);
        End If;
      
      End Loop;
    
      If v_���� Is Not Null Then
        Insert Into �ٴ������Դ����
          (����id, ����, ����, ����, ���, ���Ʒ�ʽ, ����)
        Values
          (Id_In, n_����, n_����, v_����, n_���, n_���Ʒ�ʽ, n_��������);
      
      End If;
    End Loop;
  End If;
  --�����Դ����
  If ��Դ����_In Is Not Null Then
    Insert Into �ٴ������Դ����
      (����id, ����id)
      Select Id_In As ����id, Column_Value As ����id From Table(f_Num2list(��Դ����_In));
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ������Դ����_Modify;
/

Create Or Replace Procedure Zl_�ٴ������_����(����_In �ҺŰ���.����%Type) As
  -------------------------------------------------------------------------
  --����˵���������ٴ������,��Ҫ�Ǹ��ݹҺŰ��ţ��Һżƻ����ŵȱ�������ݵ���,��������:
  -------------------------------------------------------------------------

  Err_Item Exception;
  v_Err_Msg Varchar2(200);

  l_����id t_Numlist := t_Numlist();
  n_Count  Number(18);

  v_ʱ��� Varchar2(4000);

  Procedure Zl_Register_Import(����_In �ҺŰ���.����%Type) As
    n_��Դid   �ٴ������Դ.Id%Type;
    d_��ʼʱ�� �ٴ������Դ.����ʱ��%Type;
    n_����id   �ٴ������.Id%Type;
    n_����id   �ٴ����ﰲ��.Id%Type;
    d_��ֹʱ�� �ٴ������Դ.����ʱ��%Type;
    n_����id   ��������.Id%Type;
    n_ԤԼ���� �ٴ���������.ԤԼ����%Type;
  
    l_����id t_Numlist := t_Numlist();
    n_����id �ٴ������Դ����.Id%Type;
  Begin
    --���ҡ���Ŀ��ҽ����ͬ��ֻ����һ��
    For c_��Դ In (Select ID, ����, ����, ����id, ��Ŀid, ҽ������, Decode(ҽ��id, 0, Null, ҽ��id) As ҽ��id, ���, ����, ��һ, �ܶ�, ����, ����, ����, ����,
                        ��������, ���﷽ʽ, ��ſ���, ��ʼʱ��, ��ֹʱ��, ͣ������, ִ��ʱ��, ִ�мƻ�id, �Ƿ�ɾ��, Ĭ��ʱ�μ��, ԤԼ����
                 From �ҺŰ��� A
                 Where ���� = ����_In And Not Exists (Select 1
                        From �ٴ������Դ
                        Where ����id = a.����id And ��Ŀid = a.��Ŀid And ҽ������ = a.ҽ������ And
                              ҽ��id = Decode(a.ҽ��id, 0, Null, a.ҽ��id))) Loop
    
      Select �ٴ������Դ_Id.Nextval Into n_��Դid From Dual;
    
      Select Nvl(Min(��ʼʱ��), Sysdate)
      Into d_��ʼʱ��
      From (Select Min(��ʼʱ��) As ��ʼʱ��
             From �ҺŰ���ʱ��
             Where ����id = c_��Դ.Id
             Union All
             Select Min(a.��Чʱ��) As ��ʼʱ�� From �ҺŰ��żƻ� A Where a.����id = c_��Դ.Id);
    
      --1.�����ٴ������Դ
      Insert Into �ٴ������Դ
        (ID, ����, ����, ����id, ��Ŀid, ҽ��id, ҽ������, �Ƿ񽨲���, ԤԼ����, ����Ƶ��, ���տ���״̬, �Ƿ��ٴ��Ű�, �Ű෽ʽ, �Ƿ�ɾ��, ����ʱ��, ����ʱ��)
      Values
        (n_��Դid, c_��Դ.����, c_��Դ.����, c_��Դ.����id, c_��Դ.��Ŀid, c_��Դ.ҽ��id, c_��Դ.ҽ������, c_��Դ.��������, c_��Դ.ԤԼ����, c_��Դ.Ĭ��ʱ�μ��, 2, 0,
         0, c_��Դ.�Ƿ�ɾ��, d_��ʼʱ��, Nvl(c_��Դ.ͣ������, To_Date('3000-01-01', 'yyyy-mm-dd')));
    
      --2.�����ٴ�����ͣ���¼
      Insert Into �ٴ�����ͣ���¼
        (ID, ��¼id, ��ʼʱ��, ��ֹʱ��, ͣ��ԭ��, ����ҽ��id, ����ҽ������, ������, ����ʱ��, ������, ����ʱ��, ȡ����, ȡ��ʱ��)
        Select �ٴ�����ͣ���¼_Id.Nextval As ID, Null As ��¼id, a.��ʼֹͣʱ��, a.����ֹͣʱ��, a.��ע, Null As ����ҽ��id, Null As ����ҽ������, b.ҽ������,
               a.�ƶ�����, a.�ƶ���, a.�ƶ�����, Null As ȡ����, Null As ȡ��ʱ��
        From �ҺŰ���ͣ��״̬ A, �ҺŰ��� B
        Where a.����id = b.Id And b.Id = c_��Դ.Id And b.ҽ��id Is Not Null And Not Exists
         (Select 1
               From �ٴ�����ͣ���¼
               Where ��¼id Is Null And ������ = b.ҽ������ And ��ʼʱ�� = a.��ʼֹͣʱ�� And ��ֹʱ�� = a.����ֹͣʱ��);
    
      --3.������صĳ��������
      --3.1 �̶������
      --    ����ʱ����+�̶������,���磺2015��̶������
      Begin
        Select ID Into n_����id From �ٴ������ Where ��ע = 'ϵͳ����' And ������ Is Null;
      Exception
        When Others Then
          n_����id := Null;
      End;
      If Nvl(n_����id, 0) = 0 Then
        Select �ٴ������_Id.Nextval Into n_����id From Dual;
        Insert Into �ٴ������
          (ID, �Ű෽ʽ, �������, ���, �·�, ����, Ӧ�÷�Χ, ����id, ��ע, ������, ����ʱ��)
        Values
          (n_����id, 0, To_Char(Sysdate, 'yyyy') || '��̶������', To_Number(To_Char(Sysdate, 'yyyy')), Null, Null, Null, Null,
           'ϵͳ����', Null, Null);
      End If;
    
      --3.2�����ٴ����ﰲ��
      d_��ʼʱ�� := Sysdate;
      d_��ֹʱ�� := Sysdate;
      For c_���� In (Select *
                   From (Select ID As ����id, -1 * Null As �ƻ�id, ����id, ��Ŀid, ҽ������, Decode(ҽ��id, 0, Null, ҽ��id) As ҽ��id, ����,
                                 ��һ, �ܶ�, ����, ����, ����, ����, ���﷽ʽ, ��ſ���, Nvl(��ʼʱ��, Sysdate - 3) As ��ʼʱ��,
                                 Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��
                          From �ҺŰ���
                          Where ID = c_��Դ.Id And Not Exists
                           (Select 1 From �ҺŰ��żƻ� Where ����id = c_��Դ.Id And ͣ������ Is Null) And
                                Not (���� Is Null And ��һ Is Null And �ܶ� Is Null And ���� Is Null And ���� Is Null And ���� Is Null And
                                 ���� Is Null)
                          Union All
                          Select a.����id As ����id, a.Id As �ƻ�id, b.����id, a.��Ŀid, a.ҽ������,
                                 Decode(a.ҽ��id, 0, Null, a.ҽ��id) As ҽ��id, a.����, a.��һ, a.�ܶ�, a.����, a.����, a.����, a.����, a.���﷽ʽ,
                                 a.��ſ���, Nvl(a.��Чʱ��, Sysdate - 3) As ��ʼʱ��,
                                 Nvl(a.ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��
                          From �ҺŰ��żƻ� A, �ҺŰ��� B
                          Where a.����id = b.Id And b.Id = c_��Դ.Id And b.ͣ������ Is Null And
                                Not (a.���� Is Null And a.��һ Is Null And a.�ܶ� Is Null And a.���� Is Null And a.���� Is Null And
                                 a.���� Is Null And a.���� Is Null))
                   Order By ��ʼʱ��) Loop
        If Nvl(n_����id, 0) <> 0 Then
          If c_����.��ʼʱ�� < Sysdate Then
            --������ʧЧ����
            Select ID Bulk Collect Into l_����id From �ٴ��������� Where ����id = n_����id;
          
            Forall I In 1 .. l_����id.Count
              Delete From �ٴ�����ʱ�� Where ����id = l_����id(I);
          
            Forall I In 1 .. l_����id.Count
              Delete From �ٴ�����Һſ��� Where ����id = l_����id(I);
          
            Forall I In 1 .. l_����id.Count
              Delete From �ٴ��������� Where ����id = l_����id(I);
          
            Forall I In 1 .. l_����id.Count
              Delete From �ٴ��������� Where ID = l_����id(I);
          
            Delete From �ٴ����ﰲ�� Where ID = n_����id;
          Else
            --���ϴεĿ�ʼʱ����Ϊ���ε���ֹʱ��
            Update �ٴ����ﰲ��
            Set ��ֹʱ�� = c_����.��ʼʱ�� - 1 / 24 / 60 / 60, ԭ��ֹʱ�� = c_����.��ʼʱ�� - 1 / 24 / 60 / 60
            Where ID = n_����id;
          End If;
        End If;
      
        If c_����.��ֹʱ�� > d_��ֹʱ�� Then
          d_��ֹʱ�� := c_����.��ֹʱ��;
        End If;
      
        Select �ٴ����ﰲ��_Id.Nextval Into n_����id From Dual;
        n_����id := Null;
        If Nvl(c_����.���﷽ʽ, 0) = 1 Then
          Begin
            If Nvl(c_����.�ƻ�id, 0) <> 0 Then
              Select a.Id
              Into n_����id
              From �������� A, �Һżƻ����� B
              Where a.���� = b.�������� And b.�ƻ�id = c_����.�ƻ�id And Rownum < 2;
            Else
              Select a.Id
              Into n_����id
              From �������� A, �ҺŰ������� B
              Where a.���� = b.�������� And b.�ű�id = c_����.����id And Rownum < 2;
            End If;
          Exception
            When Others Then
              n_����id := Null;
          End;
        End If;
      
        --a.�ٴ����ﰲ��
        Insert Into �ٴ����ﰲ��
          (ID, ����id, ��Դid, ��Ŀid, ҽ��id, ҽ������, �Ű����, �Ƿ���������, �Ƿ����ճ���, ��ʼʱ��, ��ֹʱ��, ����Ա����, �Ǽ�ʱ��, ԭ��ֹʱ��)
        Values
          (n_����id, n_����id, n_��Դid, c_����.��Ŀid, c_����.ҽ��id, c_����.ҽ������, Null, Null, Null, c_����.��ʼʱ��, d_��ֹʱ��, Zl_Username,
           Sysdate, d_��ֹʱ��);
      
        --b.�ٴ���������
        If Nvl(c_����.�ƻ�id, 0) <> 0 Then
          n_ԤԼ���� := 0;
          Begin
            Select 1 Into n_ԤԼ���� From �Һżƻ����� Where �ƻ�id = c_����.�ƻ�id And ��Լ�� = 0 And Rownum < 2;
          Exception
            When Others Then
              Null;
          End;
        
          Select Count(1) Into n_Count From �Һżƻ����� Where �ƻ�id = c_����.�ƻ�id And Rownum < 2;
          If n_Count = 0 Then
            Insert Into �ٴ���������
              (ID, ����id, �޺���, ��Լ��, �Ƿ���ſ���, �Ƿ��ʱ��, ԤԼ����, ������Ŀ, �ϰ�ʱ��, ���﷽ʽ, ����id)
              Select �ٴ���������_Id.Nextval, n_����id, Null, Null, c_����.��ſ���, Decode(b.����, Null, 0, 1) As �Ƿ��ʱ��, n_ԤԼ����, a.������Ŀ,
                     a.�ϰ�ʱ��, c_����.���﷽ʽ, n_����id
              From (Select '����' As ������Ŀ, c_����.���� As �ϰ�ʱ��
                     From Dual
                     Where c_����.���� Is Not Null
                     Union All
                     Select '��һ', c_����.��һ
                     From Dual
                     Where c_����.��һ Is Not Null
                     Union All
                     Select '�ܶ�', c_����.�ܶ�
                     From Dual
                     Where c_����.�ܶ� Is Not Null
                     Union All
                     Select '����', c_����.����
                     From Dual
                     Where c_����.���� Is Not Null
                     Union All
                     Select '����', c_����.����
                     From Dual
                     Where c_����.���� Is Not Null
                     Union All
                     Select '����', c_����.����
                     From Dual
                     Where c_����.���� Is Not Null
                     Union All
                     Select '����', c_����.���� From Dual Where c_����.���� Is Not Null) A,
                   (Select Distinct ���� From �Һżƻ�ʱ�� Where �ƻ�id = c_����.�ƻ�id) B
              Where a.������Ŀ = b.����(+);
          Else
            Insert Into �ٴ���������
              (ID, ����id, ������Ŀ, �ϰ�ʱ��, �޺���, ��Լ��, �Ƿ���ſ���, �Ƿ��ʱ��, ԤԼ����, ���﷽ʽ, ����id)
              Select �ٴ���������_Id.Nextval, n_����id, ������Ŀ,
                     Decode(������Ŀ, '����', c_����.����, '��һ', c_����.��һ, '�ܶ�', c_����.�ܶ�, '����', c_����.����, '����', c_����.����, '����',
                             c_����.����, '����', c_����.����, Null), �޺���, ��Լ��, c_����.��ſ���, Decode(b.����, Null, 0, 1) As �Ƿ��ʱ��,
                     n_ԤԼ����, c_����.���﷽ʽ, n_����id
              From �Һżƻ����� A, (Select Distinct ���� From �Һżƻ�ʱ�� Where �ƻ�id = c_����.�ƻ�id) B
              Where a.������Ŀ = b.����(+) And �ƻ�id = c_����.�ƻ�id;
          End If;
        Else
          n_ԤԼ���� := 0;
          Begin
            Select 1 Into n_ԤԼ���� From �ҺŰ������� Where ����id = c_����.����id And ��Լ�� = 0 And Rownum < 2;
          Exception
            When Others Then
              Null;
          End;
        
          Select Count(1) Into n_Count From �ҺŰ������� Where ����id = c_����.����id And Rownum < 2;
          If n_Count = 0 Then
            Insert Into �ٴ���������
              (ID, ����id, �޺���, ��Լ��, �Ƿ���ſ���, �Ƿ��ʱ��, ԤԼ����, ������Ŀ, �ϰ�ʱ��, ���﷽ʽ, ����id)
              Select �ٴ���������_Id.Nextval, n_����id, Null, Null, c_����.��ſ���, Decode(b.����, Null, 0, 1) As �Ƿ��ʱ��, n_ԤԼ����, a.������Ŀ,
                     a.�ϰ�ʱ��, c_����.���﷽ʽ, n_����id
              From (Select '����' As ������Ŀ, c_����.���� As �ϰ�ʱ��
                     From Dual
                     Where c_����.���� Is Not Null
                     Union All
                     Select '��һ', c_����.��һ
                     From Dual
                     Where c_����.��һ Is Not Null
                     Union All
                     Select '�ܶ�', c_����.�ܶ�
                     From Dual
                     Where c_����.�ܶ� Is Not Null
                     Union All
                     Select '����', c_����.����
                     From Dual
                     Where c_����.���� Is Not Null
                     Union All
                     Select '����', c_����.����
                     From Dual
                     Where c_����.���� Is Not Null
                     Union All
                     Select '����', c_����.����
                     From Dual
                     Where c_����.���� Is Not Null
                     Union All
                     Select '����', c_����.���� From Dual Where c_����.���� Is Not Null) A,
                   (Select Distinct ���� From �ҺŰ���ʱ�� Where ����id = c_����.����id) B
              Where a.������Ŀ = b.����(+);
          Else
            Insert Into �ٴ���������
              (ID, ����id, ������Ŀ, �ϰ�ʱ��, �޺���, ��Լ��, �Ƿ���ſ���, �Ƿ��ʱ��, ԤԼ����, ���﷽ʽ, ����id)
              Select �ٴ���������_Id.Nextval, n_����id, ������Ŀ,
                     Decode(������Ŀ, '����', c_����.����, '��һ', c_����.��һ, '�ܶ�', c_����.�ܶ�, '����', c_����.����, '����', c_����.����, '����',
                             c_����.����, '����', c_����.����, Null), �޺���, ��Լ��, c_����.��ſ���, Decode(b.����, Null, 0, 1) As �Ƿ��ʱ��,
                     n_ԤԼ����, c_����.���﷽ʽ, n_����id
              From �ҺŰ������� A, (Select Distinct ���� From �ҺŰ���ʱ�� Where ����id = c_����.����id) B
              Where a.������Ŀ = b.����(+) And ����id = c_����.����id;
          End If;
        End If;
      
        --c.�ٴ���������
        If Nvl(c_����.���﷽ʽ, 0) > 0 Then
          If Nvl(c_����.�ƻ�id, 0) <> 0 Then
            Insert Into �ٴ���������
              (����id, ����id)
              Select a.Id, b.����id
              From �ٴ��������� A,
                   (Select Distinct a.Id As ����id
                     From �������� A, �Һżƻ����� B
                     Where a.���� = b.�������� And b.�ƻ�id = c_����.�ƻ�id) B
              Where a.����id = n_����id;
          Else
            Insert Into �ٴ���������
              (����id, ����id)
              Select a.Id, b.����id
              From �ٴ��������� A,
                   (Select Distinct a.Id As ����id
                     From �������� A, �ҺŰ������� B
                     Where a.���� = b.�������� And b.�ű�id = c_����.����id) B
              Where a.����id = n_����id;
          End If;
        End If;
      
        --D.�ٴ�����ʱ��
        If Nvl(c_����.�ƻ�id, 0) <> 0 Then
          Insert Into �ٴ�����ʱ��
            (����id, ���, ��ʼʱ��, ��ֹʱ��, ��������, �Ƿ�ԤԼ)
            Select a.Id, b.���, b.��ʼʱ��, b.����ʱ��, b.��������, b.�Ƿ�ԤԼ
            From �ٴ��������� A,
                 (Select n_����id As ����id, ����,
                          Decode(����, '����', c_����.����, '��һ', c_����.��һ, '�ܶ�', c_����.�ܶ�, '����', c_����.����, '����', c_����.����, '����',
                                  c_����.����, '����', c_����.����, Null) As �ϰ�ʱ��, ���, ��ʼʱ��, ����ʱ��, ��������, �Ƿ�ԤԼ
                   From �Һżƻ�ʱ��
                   Where �ƻ�id = c_����.�ƻ�id) B
            Where a.����id = b.����id And a.������Ŀ = b.���� And a.�ϰ�ʱ�� = b.�ϰ�ʱ��;
        
        Else
          Insert Into �ٴ�����ʱ��
            (����id, ���, ��ʼʱ��, ��ֹʱ��, ��������, �Ƿ�ԤԼ)
            Select a.Id, b.���, b.��ʼʱ��, b.����ʱ��, b.��������, b.�Ƿ�ԤԼ
            From �ٴ��������� A,
                 (Select n_����id As ����id, ����,
                          Decode(����, '����', c_����.����, '��һ', c_����.��һ, '�ܶ�', c_����.�ܶ�, '����', c_����.����, '����', c_����.����, '����',
                                  c_����.����, '����', c_����.����, Null) As �ϰ�ʱ��, ���, ��ʼʱ��, ����ʱ��, ��������, �Ƿ�ԤԼ
                   From �ҺŰ���ʱ��
                   Where ����id = c_����.����id) B
            Where a.����id = b.����id And a.������Ŀ = b.���� And a.�ϰ�ʱ�� = b.�ϰ�ʱ��;
        End If;
      
        --����ʱ�ε���ſ��ƺ����������
        For c_������Ŀ In (Select ID, �޺���
                       From �ٴ���������
                       Where ����id = n_����id And Nvl(�޺���, 0) <> 0 And Nvl(�Ƿ���ſ���, 0) = 1 And Nvl(�Ƿ��ʱ��, 0) = 0) Loop
          For I In 1 .. c_������Ŀ.�޺��� Loop
            Insert Into �ٴ�����ʱ�� (����id, ���, ��������, �Ƿ�ԤԼ) Values (c_������Ŀ.Id, I, 1, 1);
          End Loop;
        End Loop;
      
        --�κ�һ����������ԤԼʱ��ʾȫ������ԤԼ
        Update �ٴ�����ʱ�� A
        Set a.�Ƿ�ԤԼ = 1
        Where ����id In (Select ID From �ٴ��������� Where ����id = n_����id) And Not Exists
         (Select 1 From �ٴ�����ʱ�� B Where a.����id = b.����id And Nvl(b.�Ƿ�ԤԼ, 0) = 1);
      
        --E.������λ�Һſ���
        If Nvl(c_����.�ƻ�id, 0) <> 0 Then
        
          Insert Into �ٴ�����Һſ���
            (����id, ����, ����, ����, ���, ���Ʒ�ʽ, ����)
            Select a.Id, b.����, b.����, b.������λ, b.���, b.���Ʒ�ʽ, b.����
            From �ٴ��������� A,
                 (Select 1 As ����, 1 As ����, ������λ, n_����id As ����id, ������Ŀ,
                          Decode(������Ŀ, '����', c_����.����, '��һ', c_����.��һ, '�ܶ�', c_����.�ܶ�, '����', c_����.����, '����', c_����.����, '����',
                                  c_����.����, '����', c_����.����, Null) As �ϰ�ʱ��, ���,
                          Case
                             When Nvl(���, 0) = 0 And Nvl(����, 0) = 0 Then
                              0
                             When ��� = 0 And Nvl(����, 0) <> 0 Then
                              2
                             When Nvl(���, 0) <> 0 And Nvl(����, 0) <> 0 Then
                              3
                             Else
                              4
                           End As ���Ʒ�ʽ, ����
                   From ������λ�ƻ�����
                   Where �ƻ�id = c_����.�ƻ�id And
                         Decode(������Ŀ, '����', c_����.����, '��һ', c_����.��һ, '�ܶ�', c_����.�ܶ�, '����', c_����.����, '����', c_����.����, '����',
                                c_����.����, '����', c_����.����, Null) Is Not Null) B
            Where a.����id = b.����id And a.������Ŀ = b.������Ŀ And a.�ϰ�ʱ�� = b.�ϰ�ʱ��;
        
        Else
          Insert Into �ٴ�����Һſ���
            (����id, ����, ����, ����, ���, ���Ʒ�ʽ, ����)
            Select a.Id, b.����, b.����, b.������λ, b.���, b.���Ʒ�ʽ, b.����
            From �ٴ��������� A,
                 (Select 1 As ����, 1 As ����, ������λ, n_����id As ����id, ������Ŀ,
                          Decode(������Ŀ, '����', c_����.����, '��һ', c_����.��һ, '�ܶ�', c_����.�ܶ�, '����', c_����.����, '����', c_����.����, '����',
                                  c_����.����, '����', c_����.����, Null) As �ϰ�ʱ��, ���,
                          Case
                            When Nvl(���, 0) = 0 And Nvl(����, 0) = 0 Then
                             0
                            When ��� = 0 And Nvl(����, 0) <> 0 Then
                             2
                            When Nvl(���, 0) <> 0 And Nvl(����, 0) <> 0 Then
                             3
                            Else
                             4
                          End As ���Ʒ�ʽ, ����
                   From ������λ���ſ���
                   Where ����id = c_����.����id And
                         Decode(������Ŀ, '����', c_����.����, '��һ', c_����.��һ, '�ܶ�', c_����.�ܶ�, '����', c_����.����, '����', c_����.����, '����',
                                c_����.����, '����', c_����.����, Null) Is Not Null) B
            Where a.����id = b.����id And a.������Ŀ = b.������Ŀ And a.�ϰ�ʱ�� = b.�ϰ�ʱ��;
        End If;
      End Loop;
    
      --4.����һ�ݳ�����Ϣ��Ϊ��Դ������Ϣ
      --˵����1.ͬһ��Դ�������/�ƻ�ʱֻ�������һ�����Ŷ�Ӧ�ĳ�����Ϣ
      --      2.�ϰ�ʱ�ΰ���������(��һ������)ȡ��һ��
      For c_���� In (Select ID, �ϰ�ʱ��, �޺���, ��Լ��, �Ƿ���ſ���, �Ƿ��ʱ��, ԤԼ����, �Ƿ��ռ, ���﷽ʽ, ����id
                   From (Select ID, �ϰ�ʱ��, �޺���, ��Լ��, �Ƿ���ſ���, �Ƿ��ʱ��, ԤԼ����, �Ƿ��ռ, ���﷽ʽ, ����id,
                                 Row_Number() Over(Partition By �ϰ�ʱ�� Order By Decode(������Ŀ, '��һ', 1, '�ܶ�', 2, '����', 3, '����', 4, '����', 5, '����', 6, '����', 7)) As ���
                          From �ٴ���������
                          Where ����id = n_����id)
                   Where ��� = 1) Loop
        --a.�ٴ������Դ����
        Select �ٴ������Դ����_Id.Nextval Into n_����id From Dual;
        Insert Into �ٴ������Դ����
          (ID, ��Դid, �ϰ�ʱ��, �޺���, ��Լ��, �Ƿ���ſ���, �Ƿ��ʱ��, ԤԼ����, �Ƿ��ռ, ���﷽ʽ, ����id)
        Values
          (n_����id, n_��Դid, c_����.�ϰ�ʱ��, c_����.�޺���, c_����.��Լ��, c_����.�Ƿ���ſ���, c_����.�Ƿ��ʱ��, c_����.ԤԼ����, c_����.�Ƿ��ռ, c_����.���﷽ʽ,
           c_����.����id);
        --b.�ٴ������Դ����
        Insert Into �ٴ������Դ����
          (����id, ����id)
          Select n_����id, ����id From �ٴ��������� Where ����id = c_����.Id;
        --c.�ٴ������Դʱ��
        Insert Into �ٴ������Դʱ��
          (����id, ���, ��ʼʱ��, ��ֹʱ��, ��������, �Ƿ�ԤԼ)
          Select n_����id, ���, ��ʼʱ��, ��ֹʱ��, ��������, �Ƿ�ԤԼ From �ٴ�����ʱ�� Where ����id = c_����.Id;
        --d.�ٴ������Դ����
        Insert Into �ٴ������Դ����
          (����id, ����, ����, ����, ���, ���Ʒ�ʽ, ����)
          Select n_����id, ����, ����, ����, ���, ���Ʒ�ʽ, ���� From �ٴ�����Һſ��� Where ����id = c_����.Id;
      End Loop;
    End Loop;
  End;
Begin
  Select Count(1) Into n_Count From �ٴ������ Where Rownum < 2;
  If n_Count <> 0 Then
    v_Err_Msg := '��ǰ�Ѿ����ڳ�����ˣ��������ٵ��룡';
    Raise Err_Item;
  End If;

  Begin
    Select f_List2str(Cast(Collect(s.ʱ���) As t_Strlist))
    Into v_ʱ���
    From (Select ʱ���, Row_Number() Over(Partition By ʱ��� Order By ʱ���) As ���
           From (Select Decode(b.�к�, 1, a.��һ, 2, a.�ܶ�, 3, a.����, 4, a.����, 5, a.����, 6, a.����, a.����) As ʱ���
                  From �ҺŰ��� A, (Select Level As �к� From Dual Connect By Level <= 7) B
                  Where a.ͣ������ Is Null And Nvl(a.��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate And
                        (����_In = a.���� Or ����_In Is Null)
                  Union All
                  Select Decode(n.�к�, 1, m.��һ, 2, m.�ܶ�, 3, m.����, 4, m.����, 5, m.����, 6, m.����, m.����) As ʱ���
                  From �ҺŰ��żƻ� M, (Select Level As �к� From Dual Connect By Level <= 7) N
                  Where Nvl(m.ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate And (����_In = m.���� Or ����_In Is Null))) S,
         ʱ��� T
    Where s.ʱ��� = t.ʱ���(+) And t.ʱ��� Is Null And s.��� = 1;
  Exception
    When Others Then
      v_ʱ��� := Null;
  End;

  If v_ʱ��� Is Not Null Then
    v_Err_Msg := 'ԭ�ҺŰ����е��ϰ�ʱ��Ρ�' || v_ʱ��� || '�������ڣ������ڡ���������>�ϰ�ʱ���������ӣ�';
    Raise Err_Item;
  End If;

  If Not ����_In Is Null Then
    Zl_Register_Import(����_In);
    Return;
  End If;

  --ɾ���������к�Դ���ڵ���֮ǰ�ѽ�������ʾ
  Select ID Bulk Collect Into l_����id From �ٴ������Դ����;

  Forall I In 1 .. l_����id.Count
    Delete From �ٴ������Դ���� Where ����id = l_����id(I);

  Forall I In 1 .. l_����id.Count
    Delete From �ٴ������Դʱ�� Where ����id = l_����id(I);

  Forall I In 1 .. l_����id.Count
    Delete From �ٴ������Դ���� Where ����id = l_����id(I);

  Forall I In 1 .. l_����id.Count
    Delete From �ٴ������Դ���� Where ID = l_����id(I);

  Delete From �ٴ������Դ;

  For c_��Դ In (Select ����
               From �ҺŰ���
               Where Nvl(�Ƿ�ɾ��, 0) = 0 And
                     Nvl(ͣ������, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd')) Loop
  
    Zl_Register_Import(c_��Դ.����);
  End Loop;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ������_����;
/
Create Or Replace Procedure Zl1_Auto_Buildingregisterplan(�Һ�ʱ��_In In Date := Null) As

  -------------------------------------------------------------------------
  --����˵�����Զ������ٴ������¼
  --          1�����ݺ�Դ�Զ�����ԤԼ���ڵ��ٴ������¼;
  --          2��ԤԼ������ȷ��:��ԴԤԼ����-->ԤԼ��ʽ��������ȡ���)-->ϵͳԤԼ����
  --���:�Һ�ʱ��_IN:NULLʱ���Զ�����;����ֻ���ָ�������Ƿ������˳����¼û��

  -------------------------------------------------------------------------
  n_ȱʡԤԼ���� Number(10);
  v_����Ա����   �ٴ����ﰲ��.����Ա����%Type;
  n_��¼id       �ٴ������¼.Id%Type;
  n_����id       �ٴ����ﰲ��.Id%Type;

  d_�������� Date;
  d_�Ǽ����� Date;
  d_�������� Date;
  d_��ǰ���� Date;
  d_��ʼ���� Date;
  d_��ֹ���� Date;
  v_ͣ��ԭ�� �ٴ������¼.ͣ��ԭ��%Type;
  v_������Ŀ �ٴ���������.������Ŀ%Type;

  n_�ڼ���   Number(2);
  v_�������� �������ձ�.��������%Type;
  n_�Ƿ���� Number(2);
  n_Count    Number(18);
Begin

  Select Max(ԤԼ����) Into n_ȱʡԤԼ���� From ԤԼ��ʽ;
  If Nvl(n_ȱʡԤԼ����, 0) = 0 Then
    n_ȱʡԤԼ���� := To_Number(Nvl(zl_GetSysParameter('�Һ�����ԤԼ����'), '0'));
  End If;
  If Nvl(n_ȱʡԤԼ����, 0) = 0 Then
    n_ȱʡԤԼ���� := 7;
  End If;

  d_��ǰ����   := Trunc(Nvl(�Һ�ʱ��_In, Sysdate));
  d_�Ǽ�����   := Sysdate;
  v_����Ա���� := Zl_Username;
  For c_��Դ In (Select c.Id, c.����, c.����, c.����id, c.ҽ������, Decode(Nvl(c.ԤԼ����, 0), 0, n_ȱʡԤԼ����, c.ԤԼ����) As ԤԼ����,
                      Nvl(b.վ��, '-') As վ��, Nvl(c.�Ƿ���ջ���, 0) As �Ƿ���ջ���, Nvl(c.���տ���״̬, 0) As ���տ���״̬
               From �ٴ������Դ C, ���ű� B
               Where c.����id = b.Id And Nvl(c.�Ƿ�ɾ��, 0) = 0 And
                     (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd')) And Exists
                (Select 1
                      From �ٴ����ﰲ�� M, �ٴ������ N
                      Where m.����id = n.Id And Nvl(n.�Ű෽ʽ, 0) = 0 And n.����ʱ�� Is Not Null And m.��Դid = c.Id And
                            m.��ֹʱ�� >= d_��ǰ����) And Not Exists
                (Select 1
                      From �ٴ������¼
                      Where ��Դid = c.Id And �������� = d_��ǰ���� + Decode(Nvl(c.ԤԼ����, 0), 0, n_ȱʡԤԼ����, c.ԤԼ����))) Loop
  
    For c_������Ϣ In (Select a.����, b.����id, b.�Ƿ���������, b.�Ƿ����ճ���,
                          Decode(To_Char(a.����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                                  '7', '����', Null) As ������Ŀ
                   From (Select Trunc(Sysdate) + ���� As ����
                          From (Select Level - 1 As ���� From Dual Connect By Level <= c_��Դ.ԤԼ���� + 1)
                          Minus
                          Select Trunc(��������) As ����
                          From �ٴ������¼
                          Where �������� Between Trunc(Sysdate) And Trunc(Sysdate) + c_��Դ.ԤԼ���� And ��Դid = c_��Դ.Id) A,
                        (Select m.Id As ����id, m.��ʼʱ��, m.��ֹʱ��, m.�Ƿ���������, m.�Ƿ����ճ���
                          From �ٴ����ﰲ�� M, �ٴ������ N
                          Where m.��Դid = c_��Դ.Id And m.����id = n.Id And Nvl(n.�Ű෽ʽ, 0) = 0 And n.����ʱ�� Is Not Null) B
                   Where a.���� Between b.��ʼʱ�� And b.��ֹʱ��) Loop
      d_�������� := c_������Ϣ.����;
      v_������Ŀ := c_������Ϣ.������Ŀ;
      n_����id   := c_������Ϣ.����id;
    
      n_�ڼ���   := 0;
      n_�Ƿ���� := 1;
      d_��ʼ���� := Null;
      d_��ֹ���� := Null;
      v_ͣ��ԭ�� := Null;
      Begin
        --��Ҫȷ��
        Select 1, ��������
        Into n_�ڼ���, v_��������
        From �������ձ�
        Where d_�������� Between ��ʼ���� And ��ֹ���� And ���� = 0;
      Exception
        When Others Then
          Null;
      End;
      --���տ���״̬��0-���ϰ�;1-�ϰ��ҿ���ԤԼ;2-�ϰ൫������ԤԼ
      If Nvl(c_��Դ.���տ���״̬, 0) = 0 And n_�ڼ��� = 1 Then
        n_�Ƿ���� := 0;
        v_ͣ��ԭ�� := v_��������;
      End If;
    
      d_�������� := Null;
      If Nvl(c_��Դ.�Ƿ���ջ���, 0) = 1 And n_�ڼ��� = 0 Then
        Begin
          --��Ҫȷ����ǰ�����Ƿ���ĳһ�컻�ݹ�����
          --��ʼ���ڣ�ԭ����Ϣ��(��������) �� ��ֹ���ڣ�ԭ���ϰ���(����������)
          Select ��ֹ���� Into d_�������� From �������ձ� Where ��ʼ���� = d_�������� And ���� = 1;
        Exception
          When Others Then
            Null;
        End;
        --��ǰ�ǻ����գ����������������ն�Ӧ���ϰ�
        If Not d_�������� Is Null Then
          n_�Ƿ���� := 1;
          Select Decode(To_Char(d_��������, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7',
                         '����', Null)
          Into v_������Ŀ
          From Dual;
        End If;
      End If;
    
      If Nvl(n_����id, 0) = 0 Then
        n_�Ƿ���� := 0;
      End If;
      --��������Ƿ����
      If n_�Ƿ���� = 1 Then
        Select Count(*) Into n_Count From �ٴ��������� Where ����id = n_����id And ������Ŀ = v_������Ŀ;
        If Nvl(n_Count, 0) = 0 Then
          n_�Ƿ���� := 0;
        End If;
      End If;
    
      If Nvl(n_�Ƿ����, 0) = 0 Then
        --�����ٴ������¼(ʱ���ΪNULL �Ŀռ�¼)
        Insert Into �ٴ������¼
          (ID, ����id, ��Դid, ��������, �Ǽ���, �Ǽ�ʱ��)
          Select �ٴ������¼_Id.Nextval, c_������Ϣ.����id, a.Id As ID, c_������Ϣ.����, v_����Ա����, d_�Ǽ����� As �Ǽ�ʱ��
          From �ٴ������Դ A, �ٴ����ﰲ�� B
          Where a.Id = b.��Դid And b.Id = c_������Ϣ.����id;
      Else
        --�����������
        Begin
          Select Min(��ʼʱ��), Max(��ֹʱ��), Max(ͣ��ԭ��)
          Into d_��ʼ����, d_��ʼ����, v_ͣ��ԭ��
          From �ٴ�����ͣ���¼
          Where ��¼id Is Null And c_������Ϣ.���� Between ��ʼʱ�� And ��ֹʱ�� And ������ = c_��Դ.ҽ������ And ������ Is Not Null And
                ȡ���� Is Null And Rownum < 2;
        Exception
          When Others Then
            d_��ʼ���� := Null;
            d_��ֹ���� := Null;
            v_ͣ��ԭ�� := Null;
        End;
      
        For c_��¼ In (With c_ʱ��� As
                        (Select ʱ���, ��ʼʱ��, ��ֹʱ��, ����, վ��, ȱʡʱ��, ��ǰʱ��
                        From (Select ʱ���, ��ʼʱ��, ��ֹʱ��, ����, վ��, ȱʡʱ��, ��ǰʱ��,
                                      Row_Number() Over(Partition By ʱ��� Order By ʱ���, վ�� Asc, ���� Asc) As ���
                               From ʱ���
                               Where Nvl(վ��, c_��Դ.վ��) = c_��Դ.վ�� And Nvl(����, c_��Դ.����) = c_��Դ.����)
                        Where ��� = 1)
                       Select c_������Ϣ.����id As ����id, B1.��Դid, c_������Ϣ.���� As ��������, m.�ϰ�ʱ��, m.Id As ����id,
                              To_Date(To_Char(c_������Ϣ.����, 'yyyy-mm-dd ') || To_Char(j.��ʼʱ��, 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                              To_Date(To_Char(c_������Ϣ.����, 'yyyy-mm-dd ') || To_Char(j.��ֹʱ��, 'hh24:mi:ss'),
                                      'yyyy-mm-dd hh24:mi:ss') + Case
                                When j.��ֹʱ�� <= j.��ʼʱ�� Then
                                 1
                                Else
                                 0
                              End As ��ֹʱ��, d_��ʼ���� As ͣ�￪ʼʱ��, d_��ֹ���� As ͣ����ֹʱ��, v_ͣ��ԭ�� As ͣ��ԭ��,
                              To_Date(To_Char(c_������Ϣ.����, 'yyyy-mm-dd ') || To_Char(Nvl(j.ȱʡʱ��, j.��ʼʱ��), 'hh24:mi:ss'),
                                      'yyyy-mm-dd hh24:mi:ss') + Case
                                When j.ȱʡʱ�� < j.��ʼʱ�� Then
                                 1
                                Else
                                 0
                              End As ȱʡԤԼʱ��,
                              To_Date(To_Char(c_������Ϣ.����, 'yyyy-mm-dd ') || To_Char(Nvl(j.��ǰʱ��, j.��ʼʱ��), 'hh24:mi:ss'),
                                      'yyyy-mm-dd hh24:mi:ss') + Case
                                When j.��ʼʱ�� < j.��ǰʱ�� Then
                                 -1
                                Else
                                 0
                              End As ��ǰ�Һ�ʱ��, m.�޺���, 0 As �ѹ���, m.��Լ��, 0 As ��Լ��, 0 As �����ѽ���, m.�Ƿ���ſ���, m.�Ƿ��ʱ��, m.ԤԼ����,
                              B1.��Ŀid, B1.ҽ��id, B1.ҽ������, Null As ����ҽ��id, Null As ����ҽ������, m.���﷽ʽ, m.����id, 0 As �Ƿ�����,
                              0 As �Ƿ���ʱ����, v_����Ա���� As ����Ա����, d_�Ǽ����� As �Ǽ�ʱ��, v_������Ŀ As ������Ŀ
                       From �ٴ����ﰲ�� B1, �ٴ��������� M, c_ʱ��� J
                       Where B1.Id = n_����id And B1.Id = m.����id And m.������Ŀ = v_������Ŀ And m.�ϰ�ʱ�� = j.ʱ��� And
                             To_Date(To_Char(c_������Ϣ.����, 'yyyy-mm-dd ') || To_Char(j.��ʼʱ��, 'hh24:mi:ss'),
                                     'yyyy-mm-dd hh24:mi:ss') >= B1.��ʼʱ��) Loop
        
          Select �ٴ������¼_Id.Nextval Into n_��¼id From Dual;
          Insert Into �ٴ������¼
            (ID, ����id, ��Դid, ��������, �ϰ�ʱ��, ��ʼʱ��, ��ֹʱ��, ͣ�￪ʼʱ��, ͣ����ֹʱ��, ͣ��ԭ��, ȱʡԤԼʱ��, ��ǰ�Һ�ʱ��, �޺���, �ѹ���, ��Լ��, ��Լ��, �����ѽ���,
             �Ƿ���ſ���, �Ƿ��ʱ��, ԤԼ����, ��Ŀid, ����id, ҽ��id, ҽ������, ����ҽ��id, ����ҽ������, ���﷽ʽ, ����id, �Ƿ�����, �Ƿ���ʱ����, �Ǽ���, �Ǽ�ʱ��, �Ƿ񷢲�)
          Values
            (n_��¼id, c_��¼.����id, c_��¼.��Դid, c_��¼.��������, c_��¼.�ϰ�ʱ��, c_��¼.��ʼʱ��, c_��¼.��ֹʱ��,
             Case When c_��¼.ͣ�￪ʼʱ�� Is Null Then Null When c_��¼.ͣ�￪ʼʱ�� < c_��¼.��ʼʱ�� Then c_��¼.��ʼʱ�� Else c_��¼.ͣ�￪ʼʱ�� End,
             Case When c_��¼.ͣ����ֹʱ�� Is Null Then Null When c_��¼.ͣ����ֹʱ�� > c_��¼.��ֹʱ�� Then c_��¼.��ֹʱ�� Else c_��¼.ͣ����ֹʱ�� End,
             c_��¼.ͣ��ԭ��, c_��¼.ȱʡԤԼʱ��, c_��¼.��ǰ�Һ�ʱ��, c_��¼.�޺���, c_��¼.�ѹ���, c_��¼.��Լ��, c_��¼.��Լ��, c_��¼.�����ѽ���, c_��¼.�Ƿ���ſ���,
             c_��¼.�Ƿ��ʱ��, c_��¼.ԤԼ����, c_��¼.��Ŀid, c_��Դ.����id, c_��¼.ҽ��id, c_��¼.ҽ������, c_��¼.����ҽ��id, c_��¼.����ҽ������, c_��¼.���﷽ʽ,
             c_��¼.����id, c_��¼.�Ƿ�����, c_��¼.�Ƿ���ʱ����, c_��¼.����Ա����, d_�Ǽ�����, 1);
        
          --�����ٴ�������ſ���
          Insert Into �ٴ�������ſ���
            (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ)
            Select n_��¼id, ���,
                   To_Date(To_Char(c_��¼.��������, 'yyyy-mm-dd ') || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
                   To_Date(To_Char(c_��¼.��������, 'yyyy-mm-dd ') || To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') + Case
                     When ��ֹʱ�� <= ��ʼʱ�� Then
                      1
                     Else
                      0
                   End, ��������, �Ƿ�ԤԼ
            From �ٴ�����ʱ��
            Where ����id = c_��¼.����id;
        
          --���������λ�Һſ��Ƽ�¼
          Insert Into �ٴ�����Һſ��Ƽ�¼
            (����, ����, ����, ��¼id, ���, ���Ʒ�ʽ, ����)
            Select ����, ����, ����, n_��¼id, ���, ���Ʒ�ʽ, ���� From �ٴ�����Һſ��� Where ����id = c_��¼.����id;
        
          --�����ٴ��������Ҽ�¼
          Insert Into �ٴ��������Ҽ�¼
            (��¼id, ����id)
            Select n_��¼id, ����id From �ٴ��������� Where ����id = c_��¼.����id;
        
        End Loop;
      End If;
      --һ��һ�ύ
      Commit;
    End Loop;
  End Loop;
End Zl1_Auto_Buildingregisterplan;
/
Create Or Replace Procedure Zl_�ٴ������_Add
(
  ��������_In Number,
  ����id_In   �ٴ������.Id%Type,
  �������_In �ٴ������.�������%Type,
  վ��_In     ���ű�.վ��%Type,
  ����Ա_In   �ٴ����ﰲ��.����Ա����%Type,
  ����ʱ��_In �ٴ����ﰲ��.�Ǽ�ʱ��%Type,
  ��ʼʱ��_In �ٴ����ﰲ��.��ʼʱ��%Type := Null,
  ��ֹʱ��_In �ٴ����ﰲ��.��ֹʱ��%Type := Null,
  ���_In     �ٴ������.���%Type := Null,
  �·�_In     �ٴ������.�·�%Type := Null,
  ����_In     �ٴ������.����%Type := Null,
  Ӧ�÷�Χ_In �ٴ������.Ӧ�÷�Χ%Type := Null,
  ����id_In   �ٴ������.����id%Type := Null,
  ��ע_In     �ٴ������.��ע%Type := Null,
  ��Աid_In   ��Ա��.Id%Type := Null,
  ɾ������_In Number := 0
) As
  --���ܣ����ӳ�����ģ��
  --������
  --        ��������_In 1-ģ�壬2-�̶�����, 3-�°��ţ�4-�ܰ���
  --        ��Աid_In ���̶���������Ч����Ϊ0��null��ʾ�ٴ�������Ա�����
  --        ɾ������_In �̶��Ű�תΪ���Ű�/���Ű�ʱ�����ƶ����Ű�/���Ű�ʱ�Ƿ�ɾ���³����ʱ����δʹ�õĳ����¼
  --˵����
  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  n_Count Number(8);

  n_����id �ٴ������.Id%Type;
  l_��¼id t_Numlist := t_Numlist();
  l_����id t_Numlist := t_Numlist();
Begin
  n_����id := ����id_In;
  If Nvl(n_����id, 0) = 0 Then
    Select �ٴ������_Id.Nextval Into n_����id From Dual;
  End If;

  --�Ű෽ʽ��0-�̶��Ű�;1-�����Ű�;2-�����Ű�;3-ģ��
  --============================================================================================================================================
  --1.ģ��
  If Nvl(��������_In, 0) = 1 Then
    Begin
      Select 1 Into n_Count From �ٴ������ Where ������� = �������_In And �Ű෽ʽ = 3 And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If Nvl(n_Count, 0) > 0 Then
      v_Err_Msg := '��ǰ�Ѵ�����Ϊ��' || �������_In || '����ģ�壡';
      Raise Err_Item;
    End If;
  
    --����Ƿ��пɲ�������Ч��Դ
    Begin
      Select 1
      Into n_Count
      From �ٴ������Դ A, ���ű� D
      Where a.����id = d.Id And a.�Ű෽ʽ In (1, 2) And Nvl(a.�Ƿ�ɾ��, 0) = 0 And
            (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd'))
           --��ǰ��Ա�ɲ����ĺ�Դ
            And (Nvl(��Աid_In, 0) = 0 Or
            (Nvl(a.�Ƿ��ٴ��Ű�, 0) = 1 And Exists (Select 1 From ������Ա Where ����id = a.����id And ��Աid = ��Աid_In)))
           --վ��
            And (d.վ�� Is Null Or d.վ�� = վ��_In) And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If Nvl(n_Count, 0) = 0 Then
      v_Err_Msg := '��ǰ�޿ɰ��»����Ű�ĺ�Դ����������ģ�壬���ȵ�����������>�ٴ���Դ��������ӳ����Դ��';
      Raise Err_Item;
    End If;
  
    --ģ�壬�϶����³����
    Insert Into �ٴ������
      (ID, �Ű෽ʽ, �������, Ӧ�÷�Χ, ����id, ��ע, ������, ����ʱ��)
    Values
      (n_����id, 3, �������_In, Ӧ�÷�Χ_In, ����id_In, ��ע_In, ����Ա_In, ����ʱ��_In);
  
    --�ٴ����ﰲ��
    For c_��Դ In (Select �ٴ����ﰲ��_Id.Nextval As ����id, n_����id As ����id, a.Id As ��Դid, a.��Ŀid, a.ҽ��id, a.ҽ������
                 From �ٴ������Դ A, ���ű� D
                 Where a.����id = d.Id And a.�Ű෽ʽ In (1, 2) And Nvl(a.�Ƿ�ɾ��, 0) = 0 And
                       (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd'))
                      --��ǰ��Ա�ɲ����ĺ�Դ
                       And (Nvl(��Աid_In, 0) = 0 Or (Nvl(a.�Ƿ��ٴ��Ű�, 0) = 1 And Exists
                        (Select 1 From ������Ա Where ����id = a.����id And ��Աid = ��Աid_In)))
                      --վ��
                       And (d.վ�� Is Null Or d.վ�� = վ��_In)) Loop
    
      Insert Into �ٴ����ﰲ��
        (ID, ����id, ��Դid, ��Ŀid, ҽ��id, ҽ������, ����Ա����, �Ǽ�ʱ��)
      Values
        (c_��Դ.����id, c_��Դ.����id, c_��Դ.��Դid, c_��Դ.��Ŀid, c_��Դ.ҽ��id, c_��Դ.ҽ������, ����Ա_In, ����ʱ��_In);
    End Loop;
    Return;
  End If;

  --============================================================================================================================================
  --2.�̶��Ű�
  If Nvl(��������_In, 0) = 2 Then
    Begin
      Select 1 Into n_Count From �ٴ������ Where ������� = �������_In And �Ű෽ʽ = 0 And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If Nvl(n_Count, 0) > 0 Then
      v_Err_Msg := '��ǰ�Ѵ�����Ϊ��' || �������_In || '���Ĺ̶������';
      Raise Err_Item;
    End If;
  
    Begin
      Select 1
      Into n_Count
      From �ٴ����ﰲ�� A, �ٴ������ B
      Where a.����id = b.Id And b.�Ű෽ʽ = 0 And a.��ʼʱ�� = ��ʼʱ��_In And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If Nvl(n_Count, 0) > 0 Then
      v_Err_Msg := '�Ѵ���Ϊ��ǰ��ʼʱ��Ĺ̶����ţ�';
      Raise Err_Item;
    End If;
  
    --����Ƿ�����Ч��Դ
    Begin
      Select 1
      Into n_Count
      From �ٴ������Դ A, ���ű� D
      Where a.����id = d.Id And a.�Ű෽ʽ = 0 And Nvl(a.�Ƿ�ɾ��, 0) = 0 And
            (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd'))
           --վ��
            And (d.վ�� Is Null Or d.վ�� = վ��_In) And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If Nvl(n_Count, 0) = 0 Then
      v_Err_Msg := '��ǰ�޿ɰ��̶��Ű�ĺ�Դ�����������̶����ţ����ȵ�����������>�ٴ���Դ��������ӳ����Դ��';
      Raise Err_Item;
    End If;
  
    --�̶����ţ��϶����³����,ֻ����"���п���"Ȩ�޵��˲�������
    Insert Into �ٴ������
      (ID, �Ű෽ʽ, �������, ���)
    Values
      (n_����id, 0, �������_In, To_Number(To_Char(��ʼʱ��_In, 'yyyy')));
  
    --ȱʡ������һ����Ч�ĳ��ﰲ��
    For c_��Դ In (Select �ٴ����ﰲ��_Id.Nextval As ����id, n_����id As ����id, ԭ����id, ��Դid, ��Ŀid, ҽ��id, ҽ������
                 From (Select a.Id As ԭ����id, b.Id As ��Դid, b.��Ŀid, b.ҽ��id, b.ҽ������,
                               Row_Number() Over(Partition By b.Id Order By a.��ʼʱ�� Desc) As ���
                        From �ٴ����ﰲ�� A, �ٴ������Դ B, �ٴ������ C, ���ű� D
                        Where a.��Դid = b.Id And a.����id = c.Id And b.����id = d.Id
                             --��Դ����
                              And b.�Ű෽ʽ = 0 And Nvl(b.�Ƿ�ɾ��, 0) = 0 And
                              (b.����ʱ�� Is Null Or b.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd'))
                             --��һ�γ��ﰲ������
                              And c.������ Is Not Null And c.�Ű෽ʽ = 0 And (d.վ�� Is Null Or d.վ�� = վ��_In)) M
                 Where ��� = 1) Loop
    
      Insert Into �ٴ����ﰲ��
        (ID, ����id, ��Դid, ��Ŀid, ҽ��id, ҽ������, ��ʼʱ��, ��ֹʱ��, ����Ա����, �Ǽ�ʱ��, ԭ��ֹʱ��)
      Values
        (c_��Դ.����id, c_��Դ.����id, c_��Դ.��Դid, c_��Դ.��Ŀid, c_��Դ.ҽ��id, c_��Դ.ҽ������, ��ʼʱ��_In, ��ֹʱ��_In, ����Ա_In, ����ʱ��_In, ��ֹʱ��_In);
    
      --���Ƴ��ﰲ��
      For c_���� In (Select ID From �ٴ��������� Where ����id = c_��Դ.ԭ����id) Loop
        Zl_�ٴ���������_Copy(c_����.Id, c_��Դ.����id);
      End Loop;
    End Loop;
  
    --��������һ����Ч���ﰲ�ŵĺ�Դ
    For c_��Դ In (Select �ٴ����ﰲ��_Id.Nextval As ����id, n_����id As ����id, a.Id As ��Դid, a.��Ŀid, a.ҽ��id, a.ҽ������
                 From �ٴ������Դ A, ���ű� D
                 Where a.����id = d.Id And a.�Ű෽ʽ = 0 And Nvl(a.�Ƿ�ɾ��, 0) = 0 And
                       (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd'))
                      --վ��
                       And (d.վ�� Is Null Or d.վ�� = վ��_In)
                      
                       And Not Exists (Select 1 From �ٴ����ﰲ�� Where ����id = n_����id And ��Դid = a.Id)) Loop
    
      Insert Into �ٴ����ﰲ��
        (ID, ����id, ��Դid, ��Ŀid, ҽ��id, ҽ������, ��ʼʱ��, ��ֹʱ��, ����Ա����, �Ǽ�ʱ��, ԭ��ֹʱ��)
      Values
        (c_��Դ.����id, c_��Դ.����id, c_��Դ.��Դid, c_��Դ.��Ŀid, c_��Դ.ҽ��id, c_��Դ.ҽ������, ��ʼʱ��_In, ��ֹʱ��_In, ����Ա_In, ����ʱ��_In, ��ֹʱ��_In);
    End Loop;
    Return;
  End If;

  --============================================================================================================================================
  --���Űࡢ���Ű�
  --����Ƿ�����Ч��Դ
  Begin
    Select 1
    Into n_Count
    From �ٴ������Դ A, ���ű� B
    Where a.����id = b.Id
         --��Ч��Դ
          And Nvl(a.�Ƿ�ɾ��, 0) = 0 And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd')) And
          (
          --���Ű�
           Nvl(��������_In, 0) = 3 And a.�Ű෽ʽ = 1
          --���Ű�
           Or Nvl(��������_In, 0) = 4 And
           (
           --��ǰ���������ʱ�䷶Χ�ڲ��������Ű�
            a.�Ű෽ʽ = 2 And Not Exists
            (Select 1
                From �ٴ����ﰲ�� P, �ٴ������ Q
                Where p.����id = q.Id And p.��Դid = a.Id And
                      Not (p.��ֹʱ�� < Trunc(��ʼʱ��_In, 'MONTH') Or p.��ʼʱ�� > Last_Day(��ʼʱ��_In)) And q.�Ű෽ʽ = 1)
           --��ǰ�ѵ���Ϊ�����Ű�,���Ǳ������������Ű࣬����ʣ�µĲ��ֽ��������ܽ����Ű�
            Or a.�Ű෽ʽ = 1 And Exists
            (Select 1
                From �ٴ����ﰲ�� P, �ٴ������ Q
                Where p.����id = q.Id And p.��Դid = a.Id And
                      Not (p.��ֹʱ�� < Trunc(��ʼʱ��_In, 'MONTH') Or p.��ʼʱ�� > Last_Day(��ʼʱ��_In)) And q.�Ű෽ʽ = 2)))
         --��Դ�ڸó����ʱ�䷶Χ���޳����¼
          And Not Exists
     (Select 1
           From �ٴ������¼ O, �ٴ����ﰲ�� P, �ٴ������ Q
           Where o.����id = p.Id And p.����id = q.Id And p.��Դid = a.Id And o.�������� Between ��ʼʱ��_In And ��ֹʱ��_In And
                 (q.�Ű෽ʽ In (1, 2)
                 --ԭ��Ϊ�̶����ﰲ��
                 Or q.�Ű෽ʽ = 0 And (Nvl(ɾ������_In, 0) = 0 Or Nvl(ɾ������_In, 0) = 1 And Exists
                  (Select 1 From ���˹Һż�¼ Where �����¼id = a.Id))))
         --��ǰ��Ա�ɲ����ĺ�Դ
          And (Nvl(��Աid_In, 0) = 0 Or
          (Nvl(a.�Ƿ��ٴ��Ű�, 0) = 1 And Exists (Select 1 From ������Ա Where ����id = a.����id And ��Աid = ��Աid_In)))
         --վ��
          And (b.վ�� Is Null Or b.վ�� = վ��_In) And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If Nvl(n_Count, 0) = 0 Then
    If Nvl(��������_In, 0) = 3 Then
      v_Err_Msg := '��ǰ�޿ɰ����Ű�ĺ�Դ�����������³�������ȵ�����������>�ٴ���Դ��������ӳ����Դ��';
    Else
      v_Err_Msg := '��ǰ�޿ɰ����Ű�ĺ�Դ�����������ܳ�������ȵ�����������>�ٴ���Դ��������ӳ����Դ��';
    End If;
    Raise Err_Item;
  End If;

  --�������ڣ��������������ֱ����ó��������ϴ���Ч��Դ���ż���
  --�漰���ٴ��Ű࣬��ǰ����Ա����ֻ�ܲ���ĳһ���ֺ�Դ
  Begin
    Select 1 Into n_Count From �ٴ������ Where ID = n_����id;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If Nvl(n_Count, 0) = 0 Then
    Insert Into �ٴ������
      (ID, �Ű෽ʽ, �������, ���, �·�, ����)
    Values
      (n_����id,
       Case
          When Nvl(��������_In, 0) = 3 Then
           1
          Else
           2
        End, �������_In, ���_In, �·�_In, ����_In);
  End If;

  --�����ǰ�����ʱ�䷶Χ���޹Һ�����ԤԼ�ĳ����¼(�̶�����)����ɾ���ⲿ�ֳ����¼(��ɾ�������ʱ�ɻָ�)��
  --���޸Ĺ̶����ŵ���ֹʱ�䣬��������ѯ��
  If Nvl(ɾ������_In, 0) = 1 Then
    For c_���� In (Select b.Id As ����id
                 From �ٴ����ﰲ�� B, �ٴ������ C, �ٴ������Դ D
                 Where b.����id = c.Id And b.��Դid = d.Id
                      --��Դ
                       And Nvl(d.�Ƿ�ɾ��, 0) = 0 And (d.����ʱ�� Is Null Or d.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd')) And
                       Nvl(d.�Ű෽ʽ, 0) = Decode(Nvl(��������_In, 0), 3, 1, 2)
                      --�����б�ʹ���˵ĳ����¼
                       And c.�Ű෽ʽ = 0 And b.��ֹʱ�� >= ��ʼʱ��_In And Not Exists
                  (Select 1
                        From �ٴ������¼ M, ���˹Һż�¼ N
                        Where m.����id = b.Id And m.Id = n.�����¼id And m.�������� >= ��ʼʱ��_In)
                      --��ǰ��Ա�ɲ����ĺ�Դ
                       And (Nvl(��Աid_In, 0) = 0 Or (Nvl(d.�Ƿ��ٴ��Ű�, 0) = 1 And Exists
                        (Select 1 From ������Ա Where ����id = d.����id And ��Աid = ��Աid_In)))) Loop
      l_����id.Extend();
      l_����id(l_����id.Count) := c_����.����id;
    
      For c_��¼ In (Select ID As ��¼id From �ٴ������¼ Where ����id = c_����.����id And �������� >= ��ʼʱ��_In) Loop
        l_��¼id.Extend();
        l_��¼id(l_��¼id.Count) := c_��¼.��¼id;
      End Loop;
    End Loop;
    Zl_�ٴ������¼_Batchdelete(l_��¼id);
    Forall I In 1 .. l_����id.Count
      Update �ٴ����ﰲ�� A
      Set a.��ֹʱ�� = ��ʼʱ��_In - 1 / 24 / 60 / 60
      Where a.Id = l_����id(I) And Not Exists (Select 1 From �ٴ������¼ Where ����id = a.Id And �������� >= ��ʼʱ��_In);
  End If;

  --ȱʡ������һ����Ч�ĳ��ﰲ��
  For c_��Դ In (Select �ٴ����ﰲ��_Id.Nextval As ����id, n_����id As ����id, ԭ����id, ��Դid, ��Ŀid, ҽ��id, ҽ������
               From (Select a.Id As ԭ����id, b.Id As ��Դid, b.��Ŀid, b.ҽ��id, b.ҽ������,
                             Row_Number() Over(Partition By b.Id Order By a.��ʼʱ�� Desc) As ���
                      From �ٴ����ﰲ�� A, �ٴ������Դ B, �ٴ������ C, ���ű� D
                      Where a.��Դid = b.Id And a.����id = c.Id And b.����id = d.Id
                           --��Ч��Դ
                            And Nvl(b.�Ƿ�ɾ��, 0) = 0 And
                            Nvl(b.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                            (
                            --���Ű�
                             Nvl(��������_In, 0) = 3 And b.�Ű෽ʽ = 1
                            --���Ű�
                             Or
                             Nvl(��������_In, 0) = 4 And
                             (
                             --��ǰ���������ʱ�䷶Χ�ڲ��������Ű�
                              b.�Ű෽ʽ = 2 And Not Exists
                              (Select 1
                               From �ٴ����ﰲ�� P, �ٴ������ Q
                               Where p.����id = q.Id And p.��Դid = b.Id And
                                     Not (p.��ֹʱ�� < Trunc(��ʼʱ��_In, 'MONTH') Or p.��ʼʱ�� > Last_Day(��ʼʱ��_In)) And q.�Ű෽ʽ = 1)
                             --��ǰ�ѵ���Ϊ�����Ű�,���Ǳ������������Ű࣬����ʣ�µĲ��ֽ��������ܽ����Ű�
                              Or b.�Ű෽ʽ = 1 And Exists
                              (Select 1
                               From �ٴ����ﰲ�� P, �ٴ������ Q
                               Where p.����id = q.Id And p.��Դid = b.Id And
                                     Not (p.��ֹʱ�� < Trunc(��ʼʱ��_In, 'MONTH') Or p.��ʼʱ�� > Last_Day(��ʼʱ��_In)) And q.�Ű෽ʽ = 2)))
                           --��һ����Ч���ﰲ��
                            And c.������ Is Not Null And c.�Ű෽ʽ = Decode(Nvl(��������_In, 0), 3, 1, 2)
                           --��Դ�ڸó����ʱ�䷶Χ���޳����¼
                            And Not Exists (Select 1
                             From �ٴ������¼ P
                             Where p.��Դid = b.Id And p.�������� Between ��ʼʱ��_In And ��ֹʱ��_In)
                           --��ǰ��Ա�ɲ����ĺ�Դ
                            And (Nvl(��Աid_In, 0) = 0 Or
                            (Nvl(b.�Ƿ��ٴ��Ű�, 0) = 1 And Exists
                             (Select 1 From ������Ա Where ����id = b.����id And ��Աid = ��Աid_In)))
                           --վ��
                            And (d.վ�� Is Null Or d.վ�� = վ��_In))
               Where ��� = 1) Loop
  
    Insert Into �ٴ����ﰲ��
      (ID, ����id, ��Դid, ��Ŀid, ҽ��id, ҽ������, ��ʼʱ��, ��ֹʱ��, ����Ա����, �Ǽ�ʱ��, ԭ��ֹʱ��)
    Values
      (c_��Դ.����id, c_��Դ.����id, c_��Դ.��Դid, c_��Դ.��Ŀid, c_��Դ.ҽ��id, c_��Դ.ҽ������, ��ʼʱ��_In, ��ֹʱ��_In, ����Ա_In, ����ʱ��_In, ��ֹʱ��_In);
  
    --���Ƴ��ﰲ��
    For c_��¼ In (Select a.Id, b.����
                 From �ٴ������¼ A,
                      (Select Trunc(��ʼʱ��_In) + Level - 1 As ����
                        From Dual
                        Connect By Level <= Trunc(��ֹʱ��_In) - Trunc(��ʼʱ��_In) + 1) B
                 Where a.����id = c_��Դ.ԭ����id
                      --���Ű�
                       And (Nvl(��������_In, 0) = 3 And To_Char(a.��������, 'dd') = To_Char(b.����, 'dd')
                       --���Ű�
                       Or Nvl(��������_In, 0) = 4 And To_Char(a.��������, 'D') = To_Char(b.����, 'D'))) Loop
      Zl_�ٴ������¼_Copy(c_��¼.Id, c_��Դ.����id, c_��¼.����, ����Ա_In, ����ʱ��_In);
    End Loop;
  End Loop;

  --��������һ����Ч���ﰲ�ŵĺ�Դ
  For c_��Դ In (Select �ٴ����ﰲ��_Id.Nextval As ����id, n_����id As ����id, a.Id As ��Դid, a.��Ŀid, a.ҽ��id, a.ҽ������
               From �ٴ������Դ A, ���ű� D
               Where a.����id = d.Id
                    --��Ч��Դ
                     And Nvl(a.�Ƿ�ɾ��, 0) = 0 And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd')) And
                     (
                     --���Ű�
                      Nvl(��������_In, 0) = 3 And a.�Ű෽ʽ = 1
                     --���Ű�
                      Or Nvl(��������_In, 0) = 4 And
                      (
                      --��ǰ���������ʱ�䷶Χ�ڲ��������Ű�
                       a.�Ű෽ʽ = 2 And Not Exists
                       (Select 1
                           From �ٴ����ﰲ�� P, �ٴ������ Q
                           Where p.����id = q.Id And p.��Դid = a.Id And
                                 Not (p.��ֹʱ�� < Trunc(��ʼʱ��_In, 'MONTH') Or p.��ʼʱ�� > Last_Day(��ʼʱ��_In)) And q.�Ű෽ʽ = 1)
                      --��ǰ�ѵ���Ϊ�����Ű�,���Ǳ������������Ű࣬����ʣ�µĲ��ֽ��������ܽ����Ű�
                       Or a.�Ű෽ʽ = 1 And Exists
                       (Select 1
                           From �ٴ����ﰲ�� P, �ٴ������ Q
                           Where p.����id = q.Id And p.��Դid = a.Id And
                                 Not (p.��ֹʱ�� < Trunc(��ʼʱ��_In, 'MONTH') Or p.��ʼʱ�� > Last_Day(��ʼʱ��_In)) And q.�Ű෽ʽ = 2)))
                    --��Դ�ڸó����ʱ�䷶Χ���޳����¼
                     And Not Exists
                (Select 1
                      From �ٴ������¼ P
                      Where p.��Դid = a.Id And p.�������� Between ��ʼʱ��_In And ��ֹʱ��_In)
                    --��ǰ��Ա�ɲ����ĺ�Դ
                     And (Nvl(��Աid_In, 0) = 0 Or (Nvl(a.�Ƿ��ٴ��Ű�, 0) = 1 And Exists
                      (Select 1 From ������Ա Where ����id = a.����id And ��Աid = ��Աid_In)))
                    --վ��
                     And (d.վ�� Is Null Or d.վ�� = վ��_In)
                    
                     And Not Exists (Select 1 From �ٴ����ﰲ�� Where ����id = n_����id And ��Դid = a.Id)) Loop
  
    Insert Into �ٴ����ﰲ��
      (ID, ����id, ��Դid, ��Ŀid, ҽ��id, ҽ������, ��ʼʱ��, ��ֹʱ��, ����Ա����, �Ǽ�ʱ��, ԭ��ֹʱ��)
    Values
      (c_��Դ.����id, c_��Դ.����id, c_��Դ.��Դid, c_��Դ.��Ŀid, c_��Դ.ҽ��id, c_��Դ.ҽ������, ��ʼʱ��_In, ��ֹʱ��_In, ����Ա_In, ����ʱ��_In, ��ֹʱ��_In);
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ������_Add;
/
Create Or Replace Procedure Zl_�ٴ���������_Copy
(
  ԭ����id_In �ٴ���������.Id%Type,
  ����id_In   �ٴ���������.����id%Type
) As
  --�����ٴ���������
  n_����id �ٴ���������.Id%Type;
Begin
  Select �ٴ���������_Id.Nextval Into n_����id From Dual;

  Insert Into �ٴ���������
    (ID, ����id, ������Ŀ, �ϰ�ʱ��, �޺���, ��Լ��, �Ƿ���ſ���, �Ƿ��ʱ��, ԤԼ����, ���﷽ʽ, ����id, �Ƿ��ռ)
    Select n_����id, ����id_In, a.������Ŀ, a.�ϰ�ʱ��, a.�޺���, a.��Լ��, a.�Ƿ���ſ���, a.�Ƿ��ʱ��, a.ԤԼ����, a.���﷽ʽ, a.����id, a.�Ƿ��ռ
    From �ٴ��������� A
    Where a.Id = ԭ����id_In;

  Insert Into �ٴ���������
    (����id, ����id)
    Select n_����id, ����id From �ٴ��������� Where ����id = ԭ����id_In;

  Insert Into �ٴ�����ʱ��
    (����id, ���, ��ʼʱ��, ��ֹʱ��, ��������, �Ƿ�ԤԼ)
    Select n_����id, ���, ��ʼʱ��, ��ֹʱ��, ��������, �Ƿ�ԤԼ From �ٴ�����ʱ�� Where ����id = ԭ����id_In;

  Insert Into �ٴ�����Һſ���
    (����id, ����, ����, ����, ���, ���Ʒ�ʽ, ����)
    Select n_����id, ����, ����, ����, ���, ���Ʒ�ʽ, ���� From �ٴ�����Һſ��� Where ����id = ԭ����id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ���������_Copy;
/
Create Or Replace Procedure Zl_�ٴ������¼_Copy
(
  ԭ��¼id_In   �ٴ������¼.Id%Type,
  ����id_In     �ٴ���������.����id%Type,
  ��������_In   �ٴ������¼.��������%Type,
  ����Ա����_In �ٴ������¼.�Ǽ���%Type,
  �Ǽ�ʱ��_In   �ٴ������¼.�Ǽ�ʱ��%Type
) As
  --�����ٴ������¼
  n_��¼id �ٴ������¼.Id%Type;

  d_��ʼʱ�� �ٴ������¼.��ʼʱ��%Type;
Begin
  Select �ٴ������¼_Id.Nextval Into n_��¼id From Dual;
  Begin
    Select a.��ʼʱ�� Into d_��ʼʱ�� From �ٴ������¼ A Where a.Id = ԭ��¼id_In;
  Exception
    When Others Then
      Return;
  End;

  Insert Into �ٴ������¼
    (ID, ����id, ��Դid, ��������, �ϰ�ʱ��, ��ʼʱ��, ��ֹʱ��, ͣ�￪ʼʱ��, ͣ����ֹʱ��, ͣ��ԭ��, ȱʡԤԼʱ��, ��ǰ�Һ�ʱ��, �޺���, �ѹ���, ��Լ��, ��Լ��, �����ѽ���, �Ƿ���ſ���,
     �Ƿ��ʱ��, ԤԼ����, �Ƿ��ռ, ��Ŀid, ����id, ҽ��id, ҽ������, ����ҽ��id, ����ҽ������, ���﷽ʽ, ����id, �Ƿ�����, �Ƿ���ʱ����, �Ǽ���, �Ǽ�ʱ��)
    Select n_��¼id, ����id_In, a.��Դid, ��������_In, a.�ϰ�ʱ��,
           To_Date(To_Char(��������_In, 'yyyy-mm-dd ') || To_Char(a.��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
           To_Date(To_Char(��������_In, 'yyyy-mm-dd ') || To_Char(a.��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') + Case
             When Trunc(��ֹʱ��) > Trunc(d_��ʼʱ��) Then
              1
             Else
              0
           End, Null As ͣ�￪ʼʱ��, Null As ͣ����ֹʱ��, Null As ͣ��ԭ��,
           To_Date(To_Char(��������_In, 'yyyy-mm-dd ') || To_Char(a.ȱʡԤԼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') + Case
             When Trunc(ȱʡԤԼʱ��) > Trunc(d_��ʼʱ��) Then
              1
             Else
              0
           End,
           To_Date(To_Char(��������_In, 'yyyy-mm-dd ') || To_Char(a.��ǰ�Һ�ʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') + Case
             When Trunc(��ǰ�Һ�ʱ��) < Trunc(d_��ʼʱ��) Then
              -1
             Else
              0
           End, a.�޺���, 0 As �ѹ���, a.��Լ��, 0 As ��Լ��, 0 As �����ѽ���, a.�Ƿ���ſ���, a.�Ƿ��ʱ��, a.ԤԼ����, a.�Ƿ��ռ, a.��Ŀid, a.����id, a.ҽ��id,
           a.ҽ������, Null As ����ҽ��id, Null As ����ҽ������, a.���﷽ʽ, a.����id, 0 As �Ƿ�����, 0 As �Ƿ���ʱ����, ����Ա����_In, �Ǽ�ʱ��_In
    From �ٴ������¼ A
    Where a.Id = ԭ��¼id_In;

  Insert Into �ٴ��������Ҽ�¼
    (��¼id, ����id)
    Select n_��¼id, ����id From �ٴ��������Ҽ�¼ Where ��¼id = ԭ��¼id_In;

  --��ʱ�β�����ŵģ���ԤԼ�Һ�ʱ��������¼����дԤԼ˳���
  Insert Into �ٴ�������ſ���
    (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ)
    Select n_��¼id, ���,
           To_Date(To_Char(��������_In, 'yyyy-mm-dd ') || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') + Case
             When Trunc(��ʼʱ��) > Trunc(d_��ʼʱ��) Then
              1
             Else
              0
           End,
           To_Date(To_Char(��������_In, 'yyyy-mm-dd ') || To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') + Case
             When Trunc(��ֹʱ��) > Trunc(d_��ʼʱ��) Then
              1
             Else
              0
           End, ����, �Ƿ�ԤԼ
    From �ٴ�������ſ���
    Where ԤԼ˳��� Is Null And ��¼id = ԭ��¼id_In;

  Insert Into �ٴ�����Һſ��Ƽ�¼
    (��¼id, ����, ����, ����, ���, ���Ʒ�ʽ, ����)
    Select n_��¼id, ����, ����, ����, ���, ���Ʒ�ʽ, ���� From �ٴ�����Һſ��Ƽ�¼ Where ��¼id = ԭ��¼id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ������¼_Copy;
/
Create Or Replace Procedure Zl_�ٴ������¼_Batchdelete(��¼id_In t_Numlist) As
  --ɾ���ٴ������¼
Begin
  Forall I In 1 .. ��¼id_In.Count
    Delete From �ٴ�����䶯��ϸ Where �䶯id In (Select ID From �ٴ�����䶯��¼ Where ��¼id = ��¼id_In(I));

  Forall I In 1 .. ��¼id_In.Count
    Delete From �ٴ�����䶯��¼ Where ��¼id = ��¼id_In(I);

  Forall I In 1 .. ��¼id_In.Count
    Delete From �ٴ�����ͣ���¼ Where ��¼id = ��¼id_In(I);

  Forall I In 1 .. ��¼id_In.Count
    Delete From �ٴ�������ſ��� Where ��¼id = ��¼id_In(I);

  Forall I In 1 .. ��¼id_In.Count
    Delete From �ٴ��������Ҽ�¼ Where ��¼id = ��¼id_In(I);

  Forall I In 1 .. ��¼id_In.Count
    Delete From �ٴ�����Һſ��Ƽ�¼ Where ��¼id = ��¼id_In(I);

  Forall I In 1 .. ��¼id_In.Count
    Delete From �ٴ������¼ Where ID = ��¼id_In(I);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ������¼_Batchdelete;
/
Create Or Replace Procedure Zl_Buildregisterfixedrule
(
  Id_In         �ٴ������.Id%Type,
  Newid_In      �ٴ������.Id%Type,
  �������_In   �ٴ������.�������%Type,
  ��ʼʱ��_In   �ٴ����ﰲ��.��ʼʱ��%Type,
  ��ֹʱ��_In   �ٴ����ﰲ��.��ֹʱ��%Type,
  ����Ա����_In �ٴ����ﰲ��.����Ա����%Type := Null,
  �Ǽ�ʱ��_In   �ٴ����ﰲ��.�Ǽ�ʱ��%Type := Null,
  վ��_In       ���ű�.վ��%Type
) As
  -------------------------------------------------------------------------
  --���ܣ��������й̶������������ɳ��µĹ̶������
  -------------------------------------------------------------------------
  n_Count Number;

  n_����id �ٴ������.Id%Type;

  v_����Ա   �ٴ����ﰲ��.����Ա����%Type;
  d_�Ǽ�ʱ�� Date;
  v_Err_Msg  Varchar2(255);
  Err_Item Exception;
Begin
  Begin
    Select 1 Into n_Count From �ٴ������ Where ID = Id_In;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count = 0 Then
    v_Err_Msg := 'δ����ԭ�������Ϣ��';
    Raise Err_Item;
  End If;

  --����Ƿ�����Ч��Դ
  Begin
    Select 1
    Into n_Count
    From �ٴ������Դ A, ���ű� B
    Where a.����id = b.Id And a.�Ű෽ʽ = 0 And Nvl(a.�Ƿ�ɾ��, 0) = 0 And
          (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd'))
         --վ��
          And (b.վ�� Is Null Or b.վ�� = վ��_In) And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If Nvl(n_Count, 0) = 0 Then
    v_Err_Msg := '��ǰ����������޿ɰ��̶��Ű�ĺ�Դ�����������µĹ̶����ţ�';
    Raise Err_Item;
  End If;

  Begin
    Select 1
    Into n_Count
    From �ٴ����ﰲ�� A, �ٴ������ B
    Where a.����id = b.Id And b.�Ű෽ʽ = 0 And a.��ʼʱ�� = ��ʼʱ��_In And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If Nvl(n_Count, 0) <> 0 Then
    v_Err_Msg := '�Ѵ���Ϊ��ǰ��ʼʱ��Ĺ̶����ţ�';
    Raise Err_Item;
  End If;

  n_����id := Newid_In;
  If Nvl(n_����id, 0) = 0 Then
    Select �ٴ������_Id.Nextval Into n_����id From Dual;
  End If;

  Insert Into �ٴ������
    (ID, �Ű෽ʽ, �������, ���)
  Values
    (n_����id, 0, �������_In, To_Number(To_Char(��ʼʱ��_In, 'yyyy')));

  d_�Ǽ�ʱ�� := Nvl(�Ǽ�ʱ��_In, Sysdate);
  v_����Ա   := Nvl(����Ա����_In, Zl_Username);

  For c_��Դ In (Select �ٴ����ﰲ��_Id.Nextval As ����id, n_����id As ����id, ԭ����id, ��Դid, ��Ŀid, ҽ��id, ҽ������
               From (Select b.Id As ԭ����id, b.��Դid, c.��Ŀid, c.ҽ��id, c.ҽ������,
                             Row_Number() Over(Partition By c.Id Order By b.��ʼʱ�� Desc) As ���
                      From �ٴ����ﰲ�� B, �ٴ������Դ C, ���ű� D
                      Where b.��Դid = c.Id And c.����id = d.Id And b.����id = Id_In
                           --��Դ����
                            And c.�Ű෽ʽ = 0 And Nvl(c.�Ƿ�ɾ��, 0) = 0 And
                            (c.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') Or c.����ʱ�� Is Null)
                           --վ��
                            And (d.վ�� Is Null Or d.վ�� = վ��_In)) M
               Where ��� = 1) Loop
  
    Insert Into �ٴ����ﰲ��
      (ID, ����id, ��Դid, ��Ŀid, ҽ��id, ҽ������, ��ʼʱ��, ��ֹʱ��, ����Ա����, �Ǽ�ʱ��, ԭ��ֹʱ��)
    Values
      (c_��Դ.����id, c_��Դ.����id, c_��Դ.��Դid, c_��Դ.��Ŀid, c_��Դ.ҽ��id, c_��Դ.ҽ������, ��ʼʱ��_In, ��ֹʱ��_In, v_����Ա, d_�Ǽ�ʱ��, ��ֹʱ��_In);
  
    --��������
    For c_���� In (Select ID, ����id, ������Ŀ, �ϰ�ʱ��, �޺���, ��Լ��, �Ƿ���ſ���, �Ƿ��ʱ��, ԤԼ����, ���﷽ʽ, ����id, �Ƿ��ռ
                 From �ٴ���������
                 Where ����id = c_��Դ.ԭ����id) Loop
    
      Zl_�ٴ���������_Copy(c_����.Id, c_��Դ.����id);
    End Loop;
  End Loop;

  --����û�еĳ��ﰲ�ŵĺ�Դ
  For c_��Դ In (Select �ٴ����ﰲ��_Id.Nextval As ����id, n_����id As ����id, a.Id As ��Դid, a.��Ŀid, a.ҽ��id, a.ҽ������
               From �ٴ������Դ A, ���ű� D
               Where a.����id = d.Id And a.�Ű෽ʽ = 0 And Nvl(a.�Ƿ�ɾ��, 0) = 0 And
                     (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd'))
                    --վ��
                     And (d.վ�� Is Null Or d.վ�� = վ��_In)
                    
                     And Not Exists (Select 1 From �ٴ����ﰲ�� Where ����id = n_����id And ��Դid = a.Id)) Loop
  
    Insert Into �ٴ����ﰲ��
      (ID, ����id, ��Դid, ��Ŀid, ҽ��id, ҽ������, ��ʼʱ��, ��ֹʱ��, ����Ա����, �Ǽ�ʱ��, ԭ��ֹʱ��)
    Values
      (c_��Դ.����id, c_��Դ.����id, c_��Դ.��Դid, c_��Դ.��Ŀid, c_��Դ.ҽ��id, c_��Դ.ҽ������, ��ʼʱ��_In, ��ֹʱ��_In, v_����Ա, d_�Ǽ�ʱ��, ��ֹʱ��_In);
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Buildregisterfixedrule;
/
Create Or Replace Procedure Zl_Buildregisterplanbyrecord
(
  ԭ����id_In   �ٴ������.Id%Type,
  �³���id_In   �ٴ������.Id%Type,
  �Ű෽ʽ_In   �ٴ������.�Ű෽ʽ%Type,
  �������_In   �ٴ������.�������%Type,
  ���_In       �ٴ������.���%Type,
  �·�_In       �ٴ������.�·�%Type,
  ����_In       �ٴ������.����%Type,
  ��ʼʱ��_In   �ٴ����ﰲ��.��ʼʱ��%Type,
  ��ֹʱ��_In   �ٴ����ﰲ��.��ֹʱ��%Type,
  ����Ա����_In �ٴ����ﰲ��.����Ա����%Type,
  �Ǽ�ʱ��_In   �ٴ����ﰲ��.�Ǽ�ʱ��%Type,
  վ��_In       ���ű�.վ��%Type,
  ��Աid_In     ��Ա��.Id%Type := Null,
  ɾ������_In   Number := 0
) As
  -------------------------------------------------------------------------
  --���ܣ����ݳ����¼�����µĳ����¼���°���/�ܰ��ţ�
  --������
  --        ��Աid_In ���̶���������Ч����Ϊ0��null��ʾ�ٴ�������Ա�����
  --        ɾ������_In �̶��Ű�תΪ���Ű�/���Ű�ʱ�����ƶ����Ű�/���Ű�ʱ�Ƿ�ɾ���³����ʱ����δʹ�õĳ����¼
  --˵����
  -------------------------------------------------------------------------
  n_Count Number;

  l_��¼id t_Numlist := t_Numlist();
  l_����id t_Numlist := t_Numlist();

  v_Err_Msg Varchar2(255);
  Err_Item Exception;
Begin
  Begin
    Select 1
    Into n_Count
    From �ٴ������Դ A, ���ű� B
    Where a.����id = b.Id
         --��Ч��Դ
          And Nvl(a.�Ƿ�ɾ��, 0) = 0 And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd')) And
          (
          --���Ű�
           Nvl(�Ű෽ʽ_In, 0) = 1 And a.�Ű෽ʽ = 1
          --���Ű�
           Or Nvl(�Ű෽ʽ_In, 0) = 2 And
           (
           --��ǰ���������ʱ�䷶Χ�ڲ��������Ű�
            a.�Ű෽ʽ = 2 And Not Exists
            (Select 1
                From �ٴ����ﰲ�� P, �ٴ������ Q
                Where p.����id = q.Id And p.��Դid = a.Id And
                      Not (p.��ֹʱ�� < Trunc(��ʼʱ��_In, 'MONTH') Or p.��ʼʱ�� > Last_Day(��ʼʱ��_In)) And q.�Ű෽ʽ = 1)
           --��ǰ�ѵ���Ϊ�����Ű�,���Ǳ������������Ű࣬����ʣ�µĲ��ֽ��������ܽ����Ű�
            Or a.�Ű෽ʽ = 1 And Exists
            (Select 1
                From �ٴ����ﰲ�� P, �ٴ������ Q
                Where p.����id = q.Id And p.��Դid = a.Id And
                      Not (p.��ֹʱ�� < Trunc(��ʼʱ��_In, 'MONTH') Or p.��ʼʱ�� > Last_Day(��ʼʱ��_In)) And q.�Ű෽ʽ = 2)))
         --��Դ�ڸó����ʱ�䷶Χ���޳����¼
          And Not Exists
     (Select 1
           From �ٴ������¼ O, �ٴ����ﰲ�� P, �ٴ������ Q
           Where o.����id = p.Id And p.����id = q.Id And p.��Դid = a.Id And o.�������� Between ��ʼʱ��_In And ��ֹʱ��_In And
                 (q.�Ű෽ʽ In (1, 2)
                 --ԭ��Ϊ�̶����ﰲ��
                 Or q.�Ű෽ʽ = 0 And (Nvl(ɾ������_In, 0) = 0 Or Nvl(ɾ������_In, 0) = 1 And Exists
                  (Select 1 From ���˹Һż�¼ Where �����¼id = a.Id))))
         --��ǰ��Ա�ɲ����ĺ�Դ
          And (Nvl(��Աid_In, 0) = 0 Or
          (Nvl(a.�Ƿ��ٴ��Ű�, 0) = 1 And Exists (Select 1 From ������Ա Where ����id = a.����id And ��Աid = ��Աid_In)))
         --վ��
          And (b.վ�� Is Null Or b.վ�� = վ��_In) And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count = 0 Then
    If Nvl(�Ű෽ʽ_In, 0) = 1 Then
      v_Err_Msg := '��ǰ����������޿ɰ����Ű�ĺ�Դ�����������µĳ����';
    Else
      v_Err_Msg := '��ǰ����������޿ɰ����Ű�ĺ�Դ�����������µĳ����';
    End If;
    Raise Err_Item;
  End If;

  --��������Ƿ����
  Begin
    Select 1 Into n_Count From �ٴ������ Where ID = �³���id_In;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If Nvl(n_Count, 0) = 0 Then
    Insert Into �ٴ������
      (ID, �Ű෽ʽ, �������, ���, �·�, ����)
    Values
      (�³���id_In, �Ű෽ʽ_In, �������_In, ���_In, �·�_In, ����_In);
  End If;

  --�����ǰ�����ʱ�䷶Χ���޹Һ�����ԤԼ�ĳ����¼(�̶�����)����ɾ���ⲿ�ֳ����¼(��ɾ�������ʱ�ɻָ�)��
  --���޸Ĺ̶����ŵ���ֹʱ�䣬��������ѯ��
  If Nvl(ɾ������_In, 0) = 1 Then
    For c_���� In (Select b.Id As ����id
                 From �ٴ����ﰲ�� B, �ٴ������ C, �ٴ������Դ D
                 Where b.����id = c.Id And b.��Դid = d.Id
                      --��Դ
                       And Nvl(d.�Ƿ�ɾ��, 0) = 0 And (d.����ʱ�� Is Null Or d.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd')) And
                       Nvl(d.�Ű෽ʽ, 0) = �Ű෽ʽ_In
                      --�����б�ʹ���˵ĳ����¼
                       And c.�Ű෽ʽ = 0 And b.��ֹʱ�� >= ��ʼʱ��_In And Not Exists
                  (Select 1
                        From �ٴ������¼ M, ���˹Һż�¼ N
                        Where m.����id = b.Id And m.Id = n.�����¼id And m.�������� >= ��ʼʱ��_In)
                      --��ǰ��Ա�ɲ����ĺ�Դ
                       And (Nvl(��Աid_In, 0) = 0 Or (Nvl(d.�Ƿ��ٴ��Ű�, 0) = 1 And Exists
                        (Select 1 From ������Ա Where ����id = d.����id And ��Աid = ��Աid_In)))) Loop
      l_����id.Extend();
      l_����id(l_����id.Count) := c_����.����id;
    
      For c_��¼ In (Select ID As ��¼id From �ٴ������¼ Where ����id = c_����.����id And �������� >= ��ʼʱ��_In) Loop
        l_��¼id.Extend();
        l_��¼id(l_��¼id.Count) := c_��¼.��¼id;
      End Loop;
    End Loop;
  
    Zl_�ٴ������¼_Batchdelete(l_��¼id);
    Forall I In 1 .. l_����id.Count
      Update �ٴ����ﰲ�� A
      Set a.��ֹʱ�� = ��ʼʱ��_In - 1 / 24 / 60 / 60
      Where a.Id = l_����id(I) And Not Exists (Select 1 From �ٴ������¼ Where ����id = a.Id And �������� >= ��ʼʱ��_In);
  End If;

  For c_��Դ In (Select �ٴ����ﰲ��_Id.Nextval As ����id, �³���id_In As ����id, b.Id As ԭ����id, b.��Դid, c.��Ŀid, c.ҽ��id, c.ҽ������
               From �ٴ����ﰲ�� B, �ٴ������Դ C, ���ű� D
               Where b.��Դid = c.Id And c.����id = d.Id And b.����id = ԭ����id_In
                    --��Ч��Դ
                     And Nvl(c.�Ƿ�ɾ��, 0) = 0 And (c.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') Or c.����ʱ�� Is Null) And
                     (
                     --���Ű�
                      Nvl(�Ű෽ʽ_In, 0) = 1 And c.�Ű෽ʽ = 1
                     -- ���Ű�
                      Or Nvl(�Ű෽ʽ_In, 0) = 2 And
                      (
                      --��ǰ���������ʱ�䷶Χ�ڲ��������Ű�
                       c.�Ű෽ʽ = 2 And Not Exists
                       (Select 1
                           From �ٴ����ﰲ�� P, �ٴ������ Q
                           Where p.����id = q.Id And p.��Դid = c.Id And
                                 Not (p.��ֹʱ�� < Trunc(��ʼʱ��_In, 'MONTH') Or p.��ʼʱ�� > Last_Day(��ʼʱ��_In)) And q.�Ű෽ʽ = 1)
                      --��ǰ�ѵ���Ϊ�����Ű�,���Ǳ������������Ű࣬����ʣ�µĲ��ֽ��������ܽ����Ű�
                       Or c.�Ű෽ʽ = 1 And Exists
                       (Select 1
                           From �ٴ����ﰲ�� P, �ٴ������ Q
                           Where p.����id = q.Id And p.��Դid = c.Id And
                                 Not (p.��ֹʱ�� < Trunc(��ʼʱ��_In, 'MONTH') Or p.��ʼʱ�� > Last_Day(��ʼʱ��_In)) And q.�Ű෽ʽ = 2)))
                    --��Դ�ڸó����ʱ�䷶Χ���޳����¼
                     And Not Exists
                (Select 1
                      From �ٴ������¼ P
                      Where p.��Դid = c.Id And p.�������� Between ��ʼʱ��_In And ��ֹʱ��_In)
                    --��ǰ��Ա�ɲ����ĺ�Դ
                     And (Nvl(��Աid_In, 0) = 0 Or (Nvl(c.�Ƿ��ٴ��Ű�, 0) = 1 And Exists
                      (Select 1 From ������Ա Where ����id = c.����id And ��Աid = ��Աid_In)))
                    --վ��
                     And (d.վ�� Is Null Or d.վ�� = վ��_In)) Loop
  
    Insert Into �ٴ����ﰲ��
      (ID, ����id, ��Դid, ��Ŀid, ҽ��id, ҽ������, ��ʼʱ��, ��ֹʱ��, ����Ա����, �Ǽ�ʱ��, ԭ��ֹʱ��)
    Values
      (c_��Դ.����id, c_��Դ.����id, c_��Դ.��Դid, c_��Դ.��Ŀid, c_��Դ.ҽ��id, c_��Դ.ҽ������, ��ʼʱ��_In, ��ֹʱ��_In, ����Ա����_In, �Ǽ�ʱ��_In, ��ֹʱ��_In);
  
    --�����¼
    For c_��¼ In (Select a.Id, b.����
                 From �ٴ������¼ A,
                      (Select Trunc(��ʼʱ��_In) + Level - 1 As ����
                        From Dual
                        Connect By Level <= Trunc(��ֹʱ��_In) - Trunc(��ʼʱ��_In) + 1) B
                 Where a.����id = c_��Դ.ԭ����id
                      --���Ű�
                       And (Nvl(�Ű෽ʽ_In, 0) = 1 And To_Char(a.��������, 'dd') = To_Char(b.����, 'dd')
                       --���Ű�
                       Or Nvl(�Ű෽ʽ_In, 0) = 2 And To_Char(a.��������, 'D') = To_Char(b.����, 'D'))) Loop
    
      Zl_�ٴ������¼_Copy(c_��¼.Id, c_��Դ.����id, c_��¼.����, ����Ա����_In, �Ǽ�ʱ��_In);
    End Loop;
  End Loop;

  --����û�еĳ��ﰲ�ŵĺ�Դ
  For c_��Դ In (Select �ٴ����ﰲ��_Id.Nextval As ����id, �³���id_In As ����id, a.Id As ��Դid, a.��Ŀid, a.ҽ��id, a.ҽ������
               From �ٴ������Դ A, ���ű� D
               Where a.����id = d.Id
                    --��Ч��Դ
                     And Nvl(a.�Ƿ�ɾ��, 0) = 0 And (a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') Or a.����ʱ�� Is Null) And
                     (
                     --���Ű�
                      Nvl(�Ű෽ʽ_In, 0) = 1 And a.�Ű෽ʽ = 1
                     -- ���Ű�
                      Or Nvl(�Ű෽ʽ_In, 0) = 2 And
                      (
                      --��ǰ���������ʱ�䷶Χ�ڲ��������Ű�
                       a.�Ű෽ʽ = 2 And Not Exists
                       (Select 1
                           From �ٴ����ﰲ�� P, �ٴ������ Q
                           Where p.����id = q.Id And p.��Դid = a.Id And
                                 Not (p.��ֹʱ�� < Trunc(��ʼʱ��_In, 'MONTH') Or p.��ʼʱ�� > Last_Day(��ʼʱ��_In)) And q.�Ű෽ʽ = 1)
                      --��ǰ�ѵ���Ϊ�����Ű�,���Ǳ������������Ű࣬����ʣ�µĲ��ֽ��������ܽ����Ű�
                       Or a.�Ű෽ʽ = 1 And Exists
                       (Select 1
                           From �ٴ����ﰲ�� P, �ٴ������ Q
                           Where p.����id = q.Id And p.��Դid = a.Id And
                                 Not (p.��ֹʱ�� < Trunc(��ʼʱ��_In, 'MONTH') Or p.��ʼʱ�� > Last_Day(��ʼʱ��_In)) And q.�Ű෽ʽ = 2)))
                    --��Դ�ڸó����ʱ�䷶Χ���޳����¼
                     And Not Exists
                (Select 1
                      From �ٴ������¼ P
                      Where p.��Դid = a.Id And p.�������� Between ��ʼʱ��_In And ��ֹʱ��_In)
                    --��ǰ��Ա�ɲ����ĺ�Դ
                     And (Nvl(��Աid_In, 0) = 0 Or (Nvl(a.�Ƿ��ٴ��Ű�, 0) = 1 And Exists
                      (Select 1 From ������Ա Where ����id = a.����id And ��Աid = ��Աid_In)))
                    --վ��
                     And (d.վ�� Is Null Or d.վ�� = վ��_In)
                    
                     And Not Exists (Select 1 From �ٴ����ﰲ�� Where ����id = �³���id_In And ��Դid = a.Id)) Loop
  
    Insert Into �ٴ����ﰲ��
      (ID, ����id, ��Դid, ��Ŀid, ҽ��id, ҽ������, ��ʼʱ��, ��ֹʱ��, ����Ա����, �Ǽ�ʱ��, ԭ��ֹʱ��)
    Values
      (c_��Դ.����id, c_��Դ.����id, c_��Դ.��Դid, c_��Դ.��Ŀid, c_��Դ.ҽ��id, c_��Դ.ҽ������, ��ʼʱ��_In, ��ֹʱ��_In, ����Ա����_In, �Ǽ�ʱ��_In, ��ֹʱ��_In);
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Buildregisterplanbyrecord;
/
Create Or Replace Procedure Zl_Buildregisterplanbytemplet
(
  ģ��id_In   �ٴ������.Id%Type,
  ��Աid_In   ��Ա��.Id%Type,
  ����id_In   �ٴ������.Id%Type,
  �Ű෽ʽ_In �ٴ������.�Ű෽ʽ%Type,
  �������_In �ٴ������.�������%Type,
  ���_In     �ٴ������.���%Type,
  �·�_In     �ٴ������.�·�%Type,
  ����_In     �ٴ������.����%Type,
  ��ʼʱ��_In �ٴ����ﰲ��.��ʼʱ��%Type,
  ��ֹʱ��_In �ٴ����ﰲ��.��ֹʱ��%Type,
  ����Ա_In   �ٴ����ﰲ��.����Ա����%Type,
  �Ǽ�ʱ��_In �ٴ����ﰲ��.�Ǽ�ʱ��%Type,
  վ��_In     ���ű�.վ��%Type,
  ɾ������_In Number := 0
) As
  -------------------------------------------------------------------------
  --����˵��������ģ���Զ������ٴ������¼
  --������
  --        ��Աid_In ���̶���������Ч����Ϊ0��null��ʾ�ٴ�������Ա�����
  --        ɾ������_In �̶��Ű�תΪ���Ű�/���Ű�ʱ�����ƶ����Ű�/���Ű�ʱ�Ƿ�ɾ���³����ʱ����δʹ�õĳ����¼
  --˵����
  -------------------------------------------------------------------------
  Err_Item Exception;
  v_Err_Msg Varchar2(200);
  n_Count   Number(18);

  d_��ѯ���� Date;
  n_��ѯ���� Number;
  v_������Ŀ �ٴ���������.������Ŀ%Type;

  n_�Ƿ���� Number(2);

  l_��¼id t_Numlist := t_Numlist();
  l_����id t_Numlist := t_Numlist();

  Procedure Isvisit
  (
    ����id_In       �ٴ����ﰲ��.Id%Type,
    �Ű����_In     �ٴ����ﰲ��.�Ű����%Type,
    ��������_In     �ٴ������¼.��������%Type,
    ��ѯ��ʼʱ��_In �ٴ����ﰲ��.��ʼʱ��%Type,
    ������Ŀ_In     Out �ٴ���������.������Ŀ%Type,
    �Ƿ����_In     Out Number
  ) As
    --�ж��Ƿ�������ȡ������Ŀ
    d_��ѯ���� Date;
    n_��ѯ���� Number;
  Begin
    �Ƿ����_In := 1;
    --��������Ƿ����
    If �Ű����_In = 1 Then
      --�����Ű�
      Select Decode(To_Char(��������_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����',
                     Null)
      Into ������Ŀ_In
      From Dual;
      Select Count(1) Into n_Count From �ٴ��������� Where ����id = ����id_In And ������Ŀ = ������Ŀ_In;
      If Nvl(n_Count, 0) = 0 Then
        �Ƿ����_In := 0;
      End If;
    Elsif �Ű����_In = 2 Then
      --�����Ű�
      ������Ŀ_In := '����';
      If Mod(To_Number(To_Char(��������_In, 'dd')), 2) <> 1 Then
        �Ƿ����_In := 0;
      End If;
    Elsif �Ű����_In = 3 Then
      --˫���Ű�
      ������Ŀ_In := '˫��';
      If Mod(To_Number(To_Char(��������_In, 'dd')), 2) <> 0 Then
        �Ƿ����_In := 0;
      End If;
    Elsif �Ű����_In = 4 Or �Ű����_In = 5 Then
      --4-������ѭ,5-��ѭ������
      If �Ű����_In = 4 Then
        d_��ѯ���� := To_Date(To_Char(��������_In, 'yyyy-mm') || To_Char(��ѯ��ʼʱ��_In, '-dd'), 'yyyy-mm-dd');
      Else
        d_��ѯ���� := ��ѯ��ʼʱ��_In;
      End If;
      Begin
        Select To_Number(Substr(������Ŀ, 1, Instr(������Ŀ, '��') - 1))
        Into n_��ѯ����
        From �ٴ���������
        Where ����id = ����id_In And Rownum < 2;
      Exception
        When Others Then
          n_��ѯ���� := 0;
      End;
      If Nvl(n_��ѯ����, 0) > 0 Then
        ������Ŀ_In := n_��ѯ���� || '��';
        If Mod(Trunc(��������_In) - Trunc(d_��ѯ����), n_��ѯ���� + 1) <> 0 Then
          �Ƿ����_In := 0;
        End If;
      End If;
    Elsif �Ű����_In = 6 Then
      --�ض�����
      ������Ŀ_In := To_Number(To_Char(��������_In, 'dd')) || '��';
      Select Count(1) Into n_Count From �ٴ��������� Where ����id = ����id_In And ������Ŀ = ������Ŀ_In;
      If Nvl(n_Count, 0) = 0 Then
        �Ƿ����_In := 0;
      End If;
    End If;
  End;
Begin
  Begin
    Select 1
    Into n_Count
    From �ٴ������Դ A, ���ű� B
    Where a.����id = b.Id
         --��Ч��Դ
          And Nvl(a.�Ƿ�ɾ��, 0) = 0 And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd')) And
          (
          --���Ű�
           Nvl(�Ű෽ʽ_In, 0) = 1 And a.�Ű෽ʽ = 1
          --���Ű�
           Or Nvl(�Ű෽ʽ_In, 0) = 2 And
           (
           --��ǰ���������ʱ�䷶Χ�ڲ��������Ű�
            a.�Ű෽ʽ = 2 And Not Exists
            (Select 1
                From �ٴ����ﰲ�� P, �ٴ������ Q
                Where p.����id = q.Id And p.��Դid = a.Id And
                      Not (p.��ֹʱ�� < Trunc(��ʼʱ��_In, 'MONTH') Or p.��ʼʱ�� > Last_Day(��ʼʱ��_In)) And q.�Ű෽ʽ = 1)
           --��ǰ�ѵ���Ϊ�����Ű�,���Ǳ������������Ű࣬����ʣ�µĲ��ֽ��������ܽ����Ű�
            Or a.�Ű෽ʽ = 1 And Exists
            (Select 1
                From �ٴ����ﰲ�� P, �ٴ������ Q
                Where p.����id = q.Id And p.��Դid = a.Id And
                      Not (p.��ֹʱ�� < Trunc(��ʼʱ��_In, 'MONTH') Or p.��ʼʱ�� > Last_Day(��ʼʱ��_In)) And q.�Ű෽ʽ = 2)))
         --��Դ�ڸó����ʱ�䷶Χ���޳����¼
          And Not Exists
     (Select 1
           From �ٴ������¼ O, �ٴ����ﰲ�� P, �ٴ������ Q
           Where o.����id = p.Id And p.����id = q.Id And p.��Դid = a.Id And o.�������� Between ��ʼʱ��_In And ��ֹʱ��_In And
                 (q.�Ű෽ʽ In (1, 2)
                 --ԭ��Ϊ�̶����ﰲ��
                 Or q.�Ű෽ʽ = 0 And (Nvl(ɾ������_In, 0) = 0 Or Nvl(ɾ������_In, 0) = 1 And Exists
                  (Select 1 From ���˹Һż�¼ Where �����¼id = a.Id))))
         --��ǰ��Ա�ɲ����ĺ�Դ
          And (Nvl(��Աid_In, 0) = 0 Or
          (Nvl(a.�Ƿ��ٴ��Ű�, 0) = 1 And Exists (Select 1 From ������Ա Where ����id = a.����id And ��Աid = ��Աid_In)))
         --վ��
          And (b.վ�� Is Null Or b.վ�� = վ��_In) And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count = 0 Then
    If Nvl(�Ű෽ʽ_In, 0) = 1 Then
      v_Err_Msg := '��ǰ����������޿ɰ����Ű�ĺ�Դ�����������µĳ����';
    Else
      v_Err_Msg := '��ǰ����������޿ɰ����Ű�ĺ�Դ�����������µĳ����';
    End If;
    Raise Err_Item;
  End If;

  --��������Ƿ����
  Begin
    Select 1 Into n_Count From �ٴ������ Where ID = ����id_In;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If Nvl(n_Count, 0) = 0 Then
    Insert Into �ٴ������
      (ID, �Ű෽ʽ, �������, ���, �·�, ����)
    Values
      (����id_In, �Ű෽ʽ_In, �������_In, ���_In, �·�_In, ����_In);
  End If;

  --�����ǰ�����ʱ�䷶Χ���޹Һ�����ԤԼ�ĳ����¼(�̶�����)����ɾ���ⲿ�ֳ����¼(��ɾ�������ʱ�ɻָ�)��
  --���޸Ĺ̶����ŵ���ֹʱ�䣬��������ѯ��
  If Nvl(ɾ������_In, 0) = 1 Then
    For c_���� In (Select b.Id As ����id
                 From �ٴ����ﰲ�� B, �ٴ������ C, �ٴ������Դ D
                 Where b.����id = c.Id And b.��Դid = d.Id
                      --��Դ
                       And Nvl(d.�Ƿ�ɾ��, 0) = 0 And (d.����ʱ�� Is Null Or d.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd')) And
                       Nvl(d.�Ű෽ʽ, 0) = �Ű෽ʽ_In
                      --�����б�ʹ���˵ĳ����¼
                       And c.�Ű෽ʽ = 0 And b.��ֹʱ�� >= ��ʼʱ��_In And Not Exists
                  (Select 1
                        From �ٴ������¼ M, ���˹Һż�¼ N
                        Where m.����id = b.Id And m.Id = n.�����¼id And m.�������� >= ��ʼʱ��_In)
                      --��ǰ��Ա�ɲ����ĺ�Դ
                       And (Nvl(��Աid_In, 0) = 0 Or (Nvl(d.�Ƿ��ٴ��Ű�, 0) = 1 And Exists
                        (Select 1 From ������Ա Where ����id = d.����id And ��Աid = ��Աid_In)))) Loop
      l_����id.Extend();
      l_����id(l_����id.Count) := c_����.����id;
    
      For c_��¼ In (Select ID As ��¼id From �ٴ������¼ Where ����id = c_����.����id And �������� >= ��ʼʱ��_In) Loop
        l_��¼id.Extend();
        l_��¼id(l_��¼id.Count) := c_��¼.��¼id;
      End Loop;
    End Loop;
  
    Zl_�ٴ������¼_Batchdelete(l_��¼id);
    Forall I In 1 .. l_����id.Count
      Update �ٴ����ﰲ�� A
      Set a.��ֹʱ�� = ��ʼʱ��_In - 1 / 24 / 60 / 60
      Where a.Id = l_����id(I) And Not Exists (Select 1 From �ٴ������¼ Where ����id = a.Id And �������� >= ��ʼʱ��_In);
  End If;

  For c_��Դ In (Select �ٴ����ﰲ��_Id.Nextval As ����id, ����id_In As ����id, b.Id As ԭ����id, b.��Դid, c.����id, c.��Ŀid, c.ҽ��id, c.ҽ������,
                      b.�Ű����, b.�Ƿ���������, b.�Ƿ����ճ���, b.��ʼʱ��, c.����, Nvl(d.վ��, '-') As վ��
               From �ٴ����ﰲ�� B, �ٴ������Դ C, ���ű� D
               Where b.��Դid = c.Id And c.����id = d.Id And b.����id = ģ��id_In
                    --��Ч��Դ
                     And Nvl(c.�Ƿ�ɾ��, 0) = 0 And (c.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') Or c.����ʱ�� Is Null) And
                     (
                     --���Ű�
                      Nvl(�Ű෽ʽ_In, 0) = 1 And c.�Ű෽ʽ = 1
                     -- ���Ű�
                      Or Nvl(�Ű෽ʽ_In, 0) = 2 And
                      (
                      --��ǰ���������ʱ�䷶Χ�ڲ��������Ű�
                       c.�Ű෽ʽ = 2 And Not Exists
                       (Select 1
                           From �ٴ����ﰲ�� P, �ٴ������ Q
                           Where p.����id = q.Id And p.��Դid = c.Id And
                                 Not (p.��ֹʱ�� < Trunc(��ʼʱ��_In, 'MONTH') Or p.��ʼʱ�� > Last_Day(��ʼʱ��_In)) And q.�Ű෽ʽ = 1)
                      --��ǰ�ѵ���Ϊ�����Ű�,���Ǳ������������Ű࣬����ʣ�µĲ��ֽ��������ܽ����Ű�
                       Or c.�Ű෽ʽ = 1 And Exists
                       (Select 1
                           From �ٴ����ﰲ�� P, �ٴ������ Q
                           Where p.����id = q.Id And p.��Դid = c.Id And
                                 Not (p.��ֹʱ�� < Trunc(��ʼʱ��_In, 'MONTH') Or p.��ʼʱ�� > Last_Day(��ʼʱ��_In)) And q.�Ű෽ʽ = 2)))
                    --��Դ�ڸó����ʱ�䷶Χ���޳����¼
                     And Not Exists
                (Select 1
                      From �ٴ������¼ P
                      Where p.��Դid = c.Id And p.�������� Between ��ʼʱ��_In And ��ֹʱ��_In)
                    --��ǰ��Ա�ɲ����ĺ�Դ
                     And (Nvl(��Աid_In, 0) = 0 Or (Nvl(c.�Ƿ��ٴ��Ű�, 0) = 1 And Exists
                      (Select 1 From ������Ա Where ����id = c.����id And ��Աid = ��Աid_In)))
                    --վ��
                     And (d.վ�� Is Null Or d.վ�� = վ��_In)) Loop
  
    Insert Into �ٴ����ﰲ��
      (ID, ����id, ��Դid, ��Ŀid, ҽ��id, ҽ������, ��ʼʱ��, ��ֹʱ��, ����Ա����, �Ǽ�ʱ��, ԭ��ֹʱ��)
    Values
      (c_��Դ.����id, c_��Դ.����id, c_��Դ.��Դid, c_��Դ.��Ŀid, c_��Դ.ҽ��id, c_��Դ.ҽ������, ��ʼʱ��_In, ��ֹʱ��_In, ����Ա_In, �Ǽ�ʱ��_In, ��ֹʱ��_In);
  
    --�ٴ������¼
    For c_���� In (Select Trunc(��ʼʱ��_In) + Level - 1 As ����,
                        Decode(To_Char(Trunc(��ʼʱ��_In) + Level - 1, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                '����', '6', '����', '7', '����', Null) As ����
                 From Dual
                 Connect By Level <= Trunc(��ֹʱ��_In) - Trunc(��ʼʱ��_In) + 1) Loop
    
      Isvisit(c_��Դ.ԭ����id, c_��Դ.�Ű����, c_����.����, c_��Դ.��ʼʱ��, v_������Ŀ, n_�Ƿ����);
    
      --�Ƿ����������ղ�����
      --�Ű����:1-�����Ű�;2-�����Ű�;3-˫���Ű�;4-������ѭ;5-��ѭ������;6-�ض�����
      If Instr(',2,3,4,5,', c_��Դ.�Ű����) > 0 And
         (Nvl(c_��Դ.�Ƿ���������, 0) = 0 And c_����.���� = '����' Or Nvl(c_��Դ.�Ƿ����ճ���, 0) = 0 And c_����.���� = '����') Then
        n_�Ƿ���� := 0;
      End If;
    
      If Nvl(n_�Ƿ����, 0) = 1 Then
        For c_��¼ In (With c_ʱ��� As
                        (Select ʱ���, ��ʼʱ��, ��ֹʱ��, ����, վ��, ȱʡʱ��, ��ǰʱ��
                        From (Select ʱ���, ��ʼʱ��, ��ֹʱ��, ����, վ��, ȱʡʱ��, ��ǰʱ��,
                                      Row_Number() Over(Partition By ʱ��� Order By ʱ���, վ�� Asc, ���� Asc) As ���
                               From ʱ���
                               Where Nvl(վ��, c_��Դ.վ��) = c_��Դ.վ�� And Nvl(����, c_��Դ.����) = c_��Դ.����)
                        Where ��� = 1)
                       Select �ٴ������¼_Id.Nextval As ��¼id, m.Id As ����id, m.�ϰ�ʱ��,
                              To_Date(To_Char(c_����.����, 'yyyy-mm-dd ') || To_Char(j.��ʼʱ��, 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                              To_Date(To_Char(c_����.����, 'yyyy-mm-dd ') || To_Char(j.��ֹʱ��, 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') + Case
                                 When j.��ֹʱ�� <= j.��ʼʱ�� Then
                                  1
                                 Else
                                  0
                               End As ��ֹʱ��,
                              To_Date(To_Char(c_����.����, 'yyyy-mm-dd ') || To_Char(Nvl(j.ȱʡʱ��, j.��ʼʱ��), 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') + Case
                                 When j.ȱʡʱ�� < j.��ʼʱ�� Then
                                  1
                                 Else
                                  0
                               End As ȱʡԤԼʱ��,
                              To_Date(To_Char(c_����.����, 'yyyy-mm-dd ') || To_Char(Nvl(j.��ǰʱ��, j.��ʼʱ��), 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') + Case
                                 When j.��ʼʱ�� < j.��ǰʱ�� Then
                                  -1
                                 Else
                                  0
                               End As ��ǰ�Һ�ʱ��, m.�޺���, m.��Լ��, m.�Ƿ���ſ���, m.�Ƿ��ʱ��, m.ԤԼ����, a.��Ŀid, a.ҽ��id, a.ҽ������, m.���﷽ʽ,
                              m.����id, m.�Ƿ��ռ
                       From �ٴ����ﰲ�� A, �ٴ��������� M, c_ʱ��� J
                       Where a.Id = m.����id And m.�ϰ�ʱ�� = j.ʱ��� And a.Id = c_��Դ.ԭ����id And m.������Ŀ = v_������Ŀ) Loop
        
          Insert Into �ٴ������¼
            (ID, ����id, ��Դid, ��������, �ϰ�ʱ��, ��ʼʱ��, ��ֹʱ��, ȱʡԤԼʱ��, ��ǰ�Һ�ʱ��, �޺���, ��Լ��, �Ƿ���ſ���, �Ƿ��ʱ��, ԤԼ����, ��Ŀid, ����id, ҽ��id,
             ҽ������, ���﷽ʽ, ����id, �Ǽ���, �Ǽ�ʱ��, �Ƿ��ռ)
          Values
            (c_��¼.��¼id, c_��Դ.����id, c_��Դ.��Դid, c_����.����, c_��¼.�ϰ�ʱ��, c_��¼.��ʼʱ��, c_��¼.��ֹʱ��, c_��¼.ȱʡԤԼʱ��, c_��¼.��ǰ�Һ�ʱ��,
             c_��¼.�޺���, c_��¼.��Լ��, c_��¼.�Ƿ���ſ���, c_��¼.�Ƿ��ʱ��, c_��¼.ԤԼ����, c_��¼.��Ŀid, c_��Դ.����id, c_��¼.ҽ��id, c_��¼.ҽ������,
             c_��¼.���﷽ʽ, c_��¼.����id, ����Ա_In, �Ǽ�ʱ��_In, c_��¼.�Ƿ��ռ);
        
          --�����ٴ�������ſ���
          Insert Into �ٴ�������ſ���
            (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ)
            Select c_��¼.��¼id, ���,
                   To_Date(To_Char(c_����.����, 'yyyy-mm-dd ') || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
                   To_Date(To_Char(c_����.����, 'yyyy-mm-dd ') || To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') + Case
                     When ��ֹʱ�� <= ��ʼʱ�� Then
                      1
                     Else
                      0
                   End, ��������, �Ƿ�ԤԼ
            From �ٴ�����ʱ��
            Where ����id = c_��¼.����id;
        
          --���������λ�Һſ��Ƽ�¼
          Insert Into �ٴ�����Һſ��Ƽ�¼
            (����, ����, ����, ��¼id, ���, ���Ʒ�ʽ, ����)
            Select ����, ����, ����, c_��¼.��¼id, ���, ���Ʒ�ʽ, ����
            From �ٴ�����Һſ���
            Where ����id = c_��¼.����id;
        
          --�����ٴ��������Ҽ�¼
          Insert Into �ٴ��������Ҽ�¼
            (��¼id, ����id)
            Select c_��¼.��¼id, ����id From �ٴ��������� Where ����id = c_��¼.����id;
        End Loop;
      End If;
    End Loop;
  End Loop;

  --����û�еĳ��ﰲ�ŵĺ�Դ
  For c_��Դ In (Select �ٴ����ﰲ��_Id.Nextval As ����id, ����id_In As ����id, a.Id As ��Դid, a.��Ŀid, a.ҽ��id, a.ҽ������
               From �ٴ������Դ A, ���ű� D
               Where a.����id = d.Id
                    --��Ч��Դ
                     And Nvl(a.�Ƿ�ɾ��, 0) = 0 And (a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') Or a.����ʱ�� Is Null) And
                     (
                     --���Ű�
                      Nvl(�Ű෽ʽ_In, 0) = 1 And a.�Ű෽ʽ = 1
                     -- ���Ű�
                      Or Nvl(�Ű෽ʽ_In, 0) = 2 And
                      (
                      --��ǰ���������ʱ�䷶Χ�ڲ��������Ű�
                       a.�Ű෽ʽ = 2 And Not Exists
                       (Select 1
                           From �ٴ����ﰲ�� P, �ٴ������ Q
                           Where p.����id = q.Id And p.��Դid = a.Id And
                                 Not (p.��ֹʱ�� < Trunc(��ʼʱ��_In, 'MONTH') Or p.��ʼʱ�� > Last_Day(��ʼʱ��_In)) And q.�Ű෽ʽ = 1)
                      --��ǰ�ѵ���Ϊ�����Ű�,���Ǳ������������Ű࣬����ʣ�µĲ��ֽ��������ܽ����Ű�
                       Or a.�Ű෽ʽ = 1 And Exists
                       (Select 1
                           From �ٴ����ﰲ�� P, �ٴ������ Q
                           Where p.����id = q.Id And p.��Դid = a.Id And
                                 Not (p.��ֹʱ�� < Trunc(��ʼʱ��_In, 'MONTH') Or p.��ʼʱ�� > Last_Day(��ʼʱ��_In)) And q.�Ű෽ʽ = 2)))
                    --��Դ�ڸó����ʱ�䷶Χ���޳����¼
                     And Not Exists
                (Select 1
                      From �ٴ������¼ P
                      Where p.��Դid = a.Id And p.�������� Between ��ʼʱ��_In And ��ֹʱ��_In)
                    --��ǰ��Ա�ɲ����ĺ�Դ
                     And (Nvl(��Աid_In, 0) = 0 Or (Nvl(a.�Ƿ��ٴ��Ű�, 0) = 1 And Exists
                      (Select 1 From ������Ա Where ����id = a.����id And ��Աid = ��Աid_In)))
                    --վ��
                     And (d.վ�� Is Null Or d.վ�� = վ��_In)
                    
                     And Not Exists (Select 1 From �ٴ����ﰲ�� Where ����id = ����id_In And ��Դid = a.Id)) Loop
  
    Insert Into �ٴ����ﰲ��
      (ID, ����id, ��Դid, ��Ŀid, ҽ��id, ҽ������, ��ʼʱ��, ��ֹʱ��, ����Ա����, �Ǽ�ʱ��, ԭ��ֹʱ��)
    Values
      (c_��Դ.����id, c_��Դ.����id, c_��Դ.��Դid, c_��Դ.��Ŀid, c_��Դ.ҽ��id, c_��Դ.ҽ������, ��ʼʱ��_In, ��ֹʱ��_In, ����Ա_In, �Ǽ�ʱ��_In, ��ֹʱ��_In);
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Buildregisterplanbytemplet;
/
Create Or Replace Procedure Zl_�ٴ������_Delete
(
  Id_In     �ٴ������.Id%Type,
  ��Աid_In ��Ա��.Id%Type := Null,
  վ��_In   ���ű�.վ��%Type
) As
  --���ܣ�ɾ���ٴ������
  --������
  --        ��Աid_In ���̶���������Ч����Ϊ0��null��ʾ�ٴ�������Ա��ɾ��
  n_Count    Number;
  n_�Ű෽ʽ �ٴ������.�Ű෽ʽ%Type;
  n_����id   �ٴ������.Id%Type;

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  l_��¼id t_Numlist := t_Numlist();
  l_����id t_Numlist := t_Numlist();
Begin
  Begin
    Select 1 Into n_Count From �ٴ������ Where �Ű෽ʽ <> 3 And ������ Is Not Null And ID = Id_In;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count <> 0 Then
    v_Err_Msg := '��ǰ������ѷ���������ɾ����';
    Raise Err_Item;
  End If;

  Begin
    Select �Ű෽ʽ Into n_�Ű෽ʽ From �ٴ������ Where ID = Id_In;
  Exception
    When Others Then
      v_Err_Msg := '�������Ϣδ�ҵ���';
      Raise Err_Item;
  End;

  If Nvl(n_�Ű෽ʽ, 0) = 0 Or Nvl(n_�Ű෽ʽ, 0) = 3 Then
    --�̶�����/ģ��
    --ɾ���ٴ���������
    Select b.Id Bulk Collect
    Into l_����id
    From �ٴ����ﰲ�� A, �ٴ��������� B
    Where a.Id = b.����id And a.����id = Id_In;
  
    Forall I In 1 .. l_����id.Count
      Delete From �ٴ�����ʱ�� Where ����id = l_����id(I);
  
    Forall I In 1 .. l_����id.Count
      Delete From �ٴ��������� Where ����id = l_����id(I);
  
    Forall I In 1 .. l_����id.Count
      Delete From �ٴ�����Һſ��� Where ����id = l_����id(I);
  
    Forall I In 1 .. l_����id.Count
      Delete From �ٴ��������� Where ID = l_����id(I);
  
    --ɾ���ٴ����ﰲ��
    Delete From �ٴ����ﰲ�� Where ����id = Id_In;
  
    --ɾ���ٴ������
    Delete �ٴ������ Where ID = Id_In;
  
    Return;
  End If;

  --========================================================================================================
  --�³����/�ܳ����
  --ֻ�ܴ����һ����ʼɾ��
  Begin
    Select ID
    Into n_����id
    From (Select a.Id
           From �ٴ������ A, �ٴ����ﰲ�� B, �ٴ������Դ C, ���ű� D
           Where a.�Ű෽ʽ = n_�Ű෽ʽ And a.Id = b.����id And b.��Դid = c.Id And c.����id = d.Id
                --��ǰ��Ա�ɲ����ĺ�Դ
                 And (Nvl(��Աid_In, 0) = 0 Or (Nvl(c.�Ƿ��ٴ��Ű�, 0) = 1 And Exists
                  (Select 1 From ������Ա Where ����id = c.����id And ��Աid = ��Աid_In)))
                --վ��
                 And (d.վ�� Is Null Or d.վ�� = վ��_In)
           Order By a.��� Desc, a.�·� Desc, a.���� Desc)
    Where Rownum < 2;
  Exception
    When Others Then
      n_����id := 0;
  End;
  If Nvl(n_����id, 0) <> 0 And Nvl(n_����id, 0) <> Id_In Then
    v_Err_Msg := '��������һ�������ʼɾ����';
    Raise Err_Item;
  End If;

  --�ָ��̶����ŵ���ֹʱ��
  For c_���� In (Select a.Id, a.ԭ��ֹʱ��
               From �ٴ����ﰲ�� A, �ٴ����ﰲ�� B, �ٴ������ C
               Where a.��Դid = b.��Դid And a.��ֹʱ�� = b.��ʼʱ�� - 1 / 24 / 60 / 60 And a.����id = c.Id And c.�Ű෽ʽ = 0 And
                     b.����id = Id_In) Loop
    Update �ٴ����ﰲ�� Set ��ֹʱ�� = c_����.ԭ��ֹʱ�� Where ID = c_����.Id;
  End Loop;

  --ɾ���ٴ������¼
  Select a.Id Bulk Collect
  Into l_��¼id
  From �ٴ������¼ A, �ٴ����ﰲ�� B, �ٴ������Դ C, ���ű� D
  Where a.����id = b.Id And a.��Դid = c.Id And c.����id = d.Id And b.����id = Id_In
       --��ǰ��Ա�ɲ����ĺ�Դ
        And (Nvl(��Աid_In, 0) = 0 Or
        (Nvl(c.�Ƿ��ٴ��Ű�, 0) = 1 And Exists (Select 1 From ������Ա Where c.����id = ����id And ��Աid = ��Աid_In)))
       --վ��
        And (d.վ�� Is Null Or d.վ�� = վ��_In);

  Zl_�ٴ������¼_Batchdelete(l_��¼id);

  --ɾ���ٴ����ﰲ��
  Delete From �ٴ����ﰲ�� A
  Where a.����id = Id_In And Exists
   (Select 1
         From �ٴ������Դ B, ���ű� D
         Where a.��Դid = b.Id And b.����id = d.Id
              --��ǰ��Ա�ɲ����ĺ�Դ
               And (Nvl(��Աid_In, 0) = 0 Or (Nvl(b.�Ƿ��ٴ��Ű�, 0) = 1 And Exists
                (Select 1 From ������Ա Where b.����id = ����id And ��Աid = ��Աid_In)))
              --վ��
               And (d.վ�� Is Null Or d.վ�� = վ��_In));

  --ɾ���ٴ������
  Delete �ٴ������ A
  Where a.Id = Id_In And Not Exists (Select 1 From �ٴ����ﰲ�� Where ����id = a.Id And ��Դid Is Not Null);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ������_Delete;
/

Create Or Replace Procedure Zl_�ٴ����ﰲ��_Applyto
(
  Ӧ������_In     Number,
  ԭid_In         �ٴ����ﰲ��.Id%Type,
  ԭ��Ŀ_In       Varchar2,
  ��id_In         �ٴ����ﰲ��.Id%Type,
  ����Ŀ_In       Varchar2,
  �Ƿ���ʱ����_In Number := 0
) As
  -------------------------------------------------------------------------
  --���ܣ���ĳ�����ڵİ���Ӧ������������
  --������
  --     ԭId_In ��Ӧ�õİ���ID
  --     ԭ��Ŀ_in ��Ӧ�õ���Ŀ
  --           1.ģ���̶������������Ŀ����"����"
  --           2.�����¼���������ڣ���"2016-01-02"
  --     ��id_In Ӧ���ڵİ���ID
  --     ����Ŀ_In Ӧ���ڵ���Ŀ�������"|"�ָ���
  --           1.ģ���̶������������Ŀ����Ŀ1|��Ŀ2|...����"����|����"
  --           2.�����¼���������ڣ�����1|����2|...����"2016-01-02|2016-01-05"
  --     Ӧ������_In 0-ģ���̶������,1-�����¼
  --˵����
  -------------------------------------------------------------------------
  n_Count    Number;
  n_����id   �ٴ���������.Id%Type;
  n_��¼id   �ٴ������¼.Id%Type;
  d_�������� �ٴ������¼.��������%Type;

  v_Err_Msg Varchar2(255);
  Err_Item Exception;
Begin
  --��鱻Ӧ�õİ����Ƿ���Ч
  If Nvl(Ӧ������_In, 0) = 0 Then
    Select Count(1)
    Into n_Count
    From �ٴ����ﰲ�� A, �ٴ��������� B
    Where a.Id = b.����id And a.Id = ԭid_In And b.������Ŀ = ԭ��Ŀ_In;
  Else
    Select Count(1)
    Into n_Count
    From �ٴ����ﰲ�� A, �ٴ������¼ B
    Where a.Id = b.����id And a.Id = ԭid_In And b.�������� = To_Date(ԭ��Ŀ_In, 'yyyy-mm-dd');
  End If;
  If n_Count = 0 Then
    v_Err_Msg := '��Ӧ�õİ���δ������Ч���ϰ�ʱ�Σ�����Ӧ�����������ţ�';
    Raise Err_Item;
  End If;

  Select Count(1) Into n_Count From �ٴ����ﰲ�� Where ID = ��id_In;
  If n_Count = 0 Then
    v_Err_Msg := 'δ�����㽫ҪӦ���ڵ��ٴ����ﰲ�ż�¼��';
    Raise Err_Item;
  End If;
  If ����Ŀ_In Is Null Then
    v_Err_Msg := '��Ӧ���ڵ���Ŀ��';
    Raise Err_Item;
  End If;

  If Nvl(Ӧ������_In, 0) = 0 Then
    --ģ���̶������
    For c_������Ŀ In (Select Column_Value As ��Ŀ From Table(f_Str2list(����Ŀ_In, '|'))) Loop
      --��ɾ������ʱ��
      Zl_�ٴ������ϰ�ʱ��_Delete(��id_In, c_������Ŀ.��Ŀ, 0);
    
      For c_ʱ�� In (Select ID, ����id, ������Ŀ, �ϰ�ʱ��, �޺���, ��Լ��, �Ƿ���ſ���, �Ƿ��ʱ��, ԤԼ����, ���﷽ʽ, ����id, �Ƿ��ռ
                   From �ٴ���������
                   Where ����id = ԭid_In And ������Ŀ = ԭ��Ŀ_In) Loop
      
        Select �ٴ���������_Id.Nextval Into n_����id From Dual;
        Insert Into �ٴ���������
          (ID, ����id, ������Ŀ, �ϰ�ʱ��, �޺���, ��Լ��, �Ƿ���ſ���, �Ƿ��ʱ��, ԤԼ����, ���﷽ʽ, ����id, �Ƿ��ռ)
        Values
          (n_����id, ��id_In, c_������Ŀ.��Ŀ, c_ʱ��.�ϰ�ʱ��, c_ʱ��.�޺���, c_ʱ��.��Լ��, c_ʱ��.�Ƿ���ſ���, c_ʱ��.�Ƿ��ʱ��, c_ʱ��.ԤԼ����, c_ʱ��.���﷽ʽ,
           c_ʱ��.����id, c_ʱ��.�Ƿ��ռ);
      
        Insert Into �ٴ���������
          (����id, ����id)
          Select n_����id, ����id From �ٴ��������� Where ����id = c_ʱ��.Id;
      
        Insert Into �ٴ�����ʱ��
          (����id, ���, ��ʼʱ��, ��ֹʱ��, ��������, �Ƿ�ԤԼ)
          Select n_����id, ���, ��ʼʱ��, ��ֹʱ��, ��������, �Ƿ�ԤԼ From �ٴ�����ʱ�� Where ����id = c_ʱ��.Id;
      
        Insert Into �ٴ�����Һſ���
          (����id, ����, ����, ����, ���, ���Ʒ�ʽ, ����)
          Select n_����id, ����, ����, ����, ���, ���Ʒ�ʽ, ���� From �ٴ�����Һſ��� Where ����id = c_ʱ��.Id;
      End Loop;
    End Loop;
  Else
    --�����¼
    For c_�������� In (Select Column_Value As ���� From Table(f_Str2list(����Ŀ_In, '|'))) Loop
      d_�������� := To_Date(c_��������.����, 'yyyy-mm-dd');
      --���ܶ���ʷ�İ��Ž��г��ﰲ�Ų���
      If Trunc(Sysdate + 1) > d_�������� Then
        v_Err_Msg := '���ܶԵ�ǰ���ڼ���ǰ�����ڽ��г��ﰲ�ţ�';
        Raise Err_Item;
      End If;
    
      --��鵱ǰ�����Ƿ������������������
      --һ����Դĳһ��İ���ֻ����һ�����������
      Begin
        Select 1
        Into n_Count
        From �ٴ������¼ A, �ٴ����ﰲ�� B, �ٴ����ﰲ�� C
        Where a.����id = b.Id And a.��Դid = c.��Դid And a.�������� = d_�������� And c.Id = ��id_In And b.Id <> ��id_In And Rownum < 2;
      Exception
        When Others Then
          n_Count := 0;
      End;
      If Nvl(n_Count, 0) = 1 Then
        v_Err_Msg := '����(' || To_Char(d_��������, 'yyyy-mm-dd') || ')��������������н����˰��ţ������ظ����ţ�';
        Raise Err_Item;
      End If;
    
      --��ɾ������ʱ��
      Zl_�ٴ������ϰ�ʱ��_Delete(��id_In, To_Char(d_��������, 'yyyy-mm-dd'), 1);
    
      For c_ʱ�� In (Select ID, ����id, ��Դid, ��������, �ϰ�ʱ��, ��ʼʱ��, ��ֹʱ��, ȱʡԤԼʱ��, ��ǰ�Һ�ʱ��, �޺���, ��Լ��, �Ƿ���ſ���, �Ƿ��ʱ��, ԤԼ����, �Ƿ��ռ,
                          ��Ŀid, ����id, ҽ��id, ҽ������, ���﷽ʽ, ����id
                   From �ٴ������¼
                   Where ����id = ԭid_In And �������� = To_Date(ԭ��Ŀ_In, 'yyyy-mm-dd')) Loop
      
        Select �ٴ������¼_Id.Nextval Into n_��¼id From Dual;
        Insert Into �ٴ������¼
          (ID, ����id, ��Դid, ��������, �ϰ�ʱ��, ��ʼʱ��, ��ֹʱ��, ȱʡԤԼʱ��, ��ǰ�Һ�ʱ��, �޺���, ��Լ��, �Ƿ���ſ���, �Ƿ��ʱ��, ԤԼ����, �Ƿ��ռ, ��Ŀid, ����id,
           ҽ��id, ҽ������, ���﷽ʽ, ����id, �Ƿ���ʱ����, �Ǽ���, �Ǽ�ʱ��)
          Select n_��¼id, a.Id, a.��Դid, d_��������, c_ʱ��.�ϰ�ʱ��,
                 To_Date(To_Char(d_��������, 'yyyy-mm-dd') || ' ' || To_Char(c_ʱ��.��ʼʱ��, 'hh24:mi:ss'),
                          'yyyy-mm-dd hh24:mi:ss'),
                 To_Date(To_Char(d_��������, 'yyyy-mm-dd') || ' ' || To_Char(c_ʱ��.��ֹʱ��, 'hh24:mi:ss'),
                          'yyyy-mm-dd hh24:mi:ss'),
                 To_Date(To_Char(d_��������, 'yyyy-mm-dd') || ' ' || To_Char(c_ʱ��.ȱʡԤԼʱ��, 'hh24:mi:ss'),
                          'yyyy-mm-dd hh24:mi:ss'),
                 To_Date(To_Char(d_��������, 'yyyy-mm-dd') || ' ' || To_Char(c_ʱ��.��ǰ�Һ�ʱ��, 'hh24:mi:ss'),
                          'yyyy-mm-dd hh24:mi:ss'), c_ʱ��.�޺���, c_ʱ��.��Լ��, c_ʱ��.�Ƿ���ſ���, c_ʱ��.�Ƿ��ʱ��, c_ʱ��.ԤԼ����, c_ʱ��.�Ƿ��ռ,
                 a.��Ŀid, b.����id, a.ҽ��id, a.ҽ������, c_ʱ��.���﷽ʽ, c_ʱ��.����id, Nvl(�Ƿ���ʱ����_In, 0), Zl_Username, Sysdate
          From �ٴ����ﰲ�� A, �ٴ������Դ B
          Where a.Id = ��id_In And a.��Դid = b.Id;
      
        Insert Into �ٴ��������Ҽ�¼
          (��¼id, ����id)
          Select n_��¼id, ����id From �ٴ��������Ҽ�¼ Where ��¼id = c_ʱ��.Id;
      
        Insert Into �ٴ�������ſ���
          (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ)
          Select n_��¼id, ���,
                 To_Date(To_Char(d_��������, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
                 To_Date(To_Char(d_��������, 'yyyy-mm-dd') || ' ' || To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
                 ����, �Ƿ�ԤԼ
          From �ٴ�������ſ���
          Where ԤԼ˳��� Is Null And ��¼id = c_ʱ��.Id;
      
        Insert Into �ٴ�����Һſ��Ƽ�¼
          (��¼id, ����, ����, ����, ���, ���Ʒ�ʽ, ����)
          Select n_��¼id, ����, ����, ����, ���, ���Ʒ�ʽ, ���� From �ٴ�����Һſ��Ƽ�¼ Where ��¼id = c_ʱ��.Id;
      
      End Loop;
    End Loop;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ����ﰲ��_Applyto;
/

Create Or Replace Procedure Zl_�ٴ����ﰲ��_Batchdelete
(
  ����id_In �ٴ������.Id%Type,
  ��Աid_In ��Ա��.Id%Type := 0,
  վ��_In   ���ű�.վ��%Type := Null,
  ��Դid_In �ٴ����ﰲ��.��Դid%Type := 0
) As
  --���ܣ�����ɾ���ٴ����ﰲ��
  --������
  --      ��Աid_In ������0��ɾ����Ա���ڿ��ҵ����к�Դ����
  --      ��Դid_In ������0��ɾ���ú�Դ�����а���
  --˵���������Աid_In=0�Һ�Դid_In=0 ��ɾ���ó��������к�Դ�����а���
  n_Count    Number(8);
  n_�����¼ Number(1);

  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  l_����id t_Numlist := t_Numlist();
  l_��¼id t_Numlist := t_Numlist();
Begin
  Begin
    Select 1
    Into n_Count
    From �ٴ������ A
    Where a.Id = ����id_In And a.������ Is Not Null And a.�Ű෽ʽ <> 3 And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count <> 0 Then
    v_Err_Msg := '��ǰ������ѷ������������޸İ��ţ�';
    Raise Err_Item;
  End If;

  Begin
    Select 1 Into n_�����¼ From �ٴ������ A Where a.Id = ����id_In And a.�Ű෽ʽ In (1, 2) And Rownum < 2;
  Exception
    When Others Then
      n_�����¼ := 0;
  End;

  If Nvl(n_�����¼, 0) = 0 Then
    --ɾ���ٴ��������/ģ��
    Select a.Id Bulk Collect
    Into l_����id
    From �ٴ��������� A, �ٴ����ﰲ�� B, �ٴ������Դ C, ���ű� D
    Where a.����id = b.Id And b.��Դid = c.Id And c.����id = d.Id And b.����id = ����id_In And
          (
          --ɾ���ó��������к�Դ�����а���
           (Nvl(��Դid_In, 0) = 0 And Nvl(��Աid_In, 0) = 0)
          --ɾ���ú�Դ�����а���
           Or (Nvl(��Դid_In, 0) <> 0 And b.��Դid = ��Դid_In)
          --ɾ����Ա���ڿ��ҵ����к�Դ����
           Or (Nvl(��Աid_In, 0) <> 0 And Exists (Select 1 From ������Ա Where ����id = c.����id And ��Աid = ��Աid_In)))
         --վ��
          And (d.վ�� Is Null Or d.վ�� = վ��_In);
  
    Forall I In 1 .. l_����id.Count
      Delete From �ٴ�����Һſ��� Where ����id = l_����id(I);
  
    Forall I In 1 .. l_����id.Count
      Delete From �ٴ�����ʱ�� Where ����id = l_����id(I);
  
    Forall I In 1 .. l_����id.Count
      Delete From �ٴ��������� Where ����id = l_����id(I);
  
    Forall I In 1 .. l_����id.Count
      Delete From �ٴ��������� Where ID = l_����id(I);
  Else
    --ɾ���ٴ������¼
    Select a.Id Bulk Collect
    Into l_��¼id
    From �ٴ������¼ A, �ٴ����ﰲ�� B, �ٴ������Դ C, ���ű� D
    Where a.����id = b.Id And b.��Դid = c.Id And c.����id = d.Id And b.����id = ����id_In And
          (
          --ɾ���ó��������к�Դ�����а���
           (Nvl(��Դid_In, 0) = 0 And Nvl(��Աid_In, 0) = 0)
          --ɾ���ú�Դ�����а���
           Or (Nvl(��Դid_In, 0) <> 0 And b.��Դid = ��Դid_In)
          --ɾ����Ա���ڿ��ҵ����к�Դ����
           Or (Nvl(��Աid_In, 0) <> 0 And Exists (Select 1 From ������Ա Where ����id = c.����id And ��Աid = ��Աid_In)))
         --վ��
          And (d.վ�� Is Null Or d.վ�� = վ��_In);
  
    Zl_�ٴ������¼_Batchdelete(l_��¼id);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ����ﰲ��_Batchdelete;
/

Create Or Replace Procedure Zl_�ٴ����ﰲ��_Publish
(
  Id_In       �ٴ������.Id%Type,
  ������_In   �ٴ������.������%Type := Null,
  ����ʱ��_In �ٴ������.����ʱ��%Type := Null,
  ȡ������_In Number := 0
) As
  --������ȡ����������
  --������
  --        ȡ������_In �Ƿ�ȡ������
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_Count    Number(2);
  n_�Ű෽ʽ �ٴ������.�Ű෽ʽ%Type;
  l_��¼id   t_Numlist := t_Numlist();

  d_ͣ�￪ʼʱ�� �ٴ������¼.ͣ�￪ʼʱ��%Type;
  d_ͣ����ֹʱ�� �ٴ������¼.ͣ����ֹʱ��%Type;
  v_ͣ��ԭ��     �ٴ������¼.ͣ��ԭ��%Type;
  d_ԭ�ϰ�����   �ٴ������¼.��������%Type;
  d_��������     �ٴ������¼.��������%Type;
  n_����         Number(2);

  n_�Һ��Ű�ģʽ Number(2);
  v_�Һ��Ű�ģʽ Varchar2(255);

  Function Fun_Isclinicvisit
  (
    ��Դid_In       In �ٴ������Դ.Id%Type,
    ��ʼʱ��_In     In �ٴ������¼.��ʼʱ��%Type,
    ��ֹʱ��_In     In �ٴ������¼.��ֹʱ��%Type,
    ͣ�￪ʼʱ��_In Out �ٴ������¼.ͣ�￪ʼʱ��%Type,
    ͣ����ֹʱ��_In Out �ٴ������¼.ͣ����ֹʱ��%Type,
    ͣ��ԭ��_In     Out �ٴ������¼.ͣ��ԭ��%Type,
    ԭ�ϰ�����_In   Out �ٴ������¼.��������%Type,
    ��������_In     Out �ٴ������¼.��������%Type
  ) Return Number As
    --���ܣ��ж�ҽ����ĳ��ʱ�䷶Χ�Ƿ����
    --     ��ÿһ��ʱ��ν��м��(��������+�ϰ�ʱ���)
    --��Σ�
    --     ��Դid_In���ٴ������ԴID
    --     ��ʼʱ��_In���ϰ�ʱ�εĿ�ʼʱ��
    --     ��ֹʱ��_In���ϰ�ʱ�ε���ֹʱ��
    --���Σ�
    --     ͣ��ԭ��_In��������ʱ����ͣ��ԭ�򣨶���Ե�һ��Ϊ׼�������򷵻ؿ�
    --���أ�
    --     0-������
    --     1-�ڷ����ڼ����ڣ�ͬʱ�ٴ������Դ.���տ���״̬=0(0-���ϰ�;1-�ϰ��ҿ���ԤԼ;2-�ϰ൫������ԤԼ)
    --     2-��ͣ�ﰲ��ʱ�䷶Χ��
    --     else-��������
    --˵����
    --     1)ȫ�������ڲ�����ʱ�䷶Χ��-->����
    --     2)ȫ�����ڲ�����ʱ�䷶Χ��-->������
    --     3)�����ڲ�����ʱ�䷶Χ��-->������
  
    n_���տ���״̬ �ٴ������Դ.���տ���״̬%Type;
    n_�Ƿ���ջ��� �ٴ������Դ.�Ƿ���ջ���%Type;
  Begin
    --�����ڼ���
    Begin
      Select Nvl(b.���տ���״̬, 0), Nvl(�Ƿ���ջ���, 0)
      Into n_���տ���״̬, n_�Ƿ���ջ���
      From �ٴ������Դ B
      Where b.Id = ��Դid_In;
    Exception
      When Others Then
        n_���տ���״̬ := 0;
        n_�Ƿ���ջ��� := 0;
    End;
  
    If Nvl(n_���տ���״̬, 0) = 0 Then
      --���տ���״̬:0-���ϰ�;1-�ϰ��ҿ���ԤԼ;2-�ϰ൫������ԤԼ
      Begin
        Select a.��ʼ����, a.��ֹ����, a.��������
        Into ͣ�￪ʼʱ��_In, ͣ����ֹʱ��_In, ͣ��ԭ��_In
        From �������ձ� A
        Where a.���� = 0 And ��ʼʱ��_In < a.��ֹ���� And ��ֹʱ��_In > a.��ʼ���� And Rownum < 2;
      
        If ͣ�￪ʼʱ��_In < ��ʼʱ��_In Then
          ͣ�￪ʼʱ��_In := ��ʼʱ��_In;
        End If;
        If ͣ����ֹʱ��_In > ��ֹʱ��_In Then
          ͣ����ֹʱ��_In := ��ֹʱ��_In;
        End If;
      
        --ȷ���Ƿ���Ҫ����
        If Nvl(n_�Ƿ���ջ���, 0) = 1 Then
          --1.ǰ��Ļ�������
          Begin
            --��ʼ���ڣ�ԭ����Ϣ��(��������) �� ��ֹ���ڣ�ԭ���ϰ���(����������)
            Select a.��ֹ����
            Into ԭ�ϰ�����_In
            From �������ձ� A
            Where a.���� = 1 And ��ʼʱ��_In < a.��ʼ���� + 1 - 1 / 24 / 60 / 60 And ��ֹʱ��_In > a.��ʼ���� And Rownum < 2;
          Exception
            When Others Then
              ԭ�ϰ�����_In := Null;
          End;
        
          --2.����Ļ���ǰ�棬���ܺ���Ļ�û�з������ڷ���ǰ��ĳ����ʱ��û�л���
          Begin
            --��ʼ���ڣ�ԭ����Ϣ��(��������) �� ��ֹ���ڣ�ԭ���ϰ���(����������)
            Select a.��ʼ����
            Into ��������_In
            From �������ձ� A
            Where a.���� = 1 And ��ʼʱ��_In < a.��ֹ���� + 1 - 1 / 24 / 60 / 60 And ��ֹʱ��_In > a.��ֹ���� And Rownum < 2;
          Exception
            When Others Then
              ��������_In := Null;
          End;
        End If;
      
        Return 1;
      Exception
        When Others Then
          ͣ�￪ʼʱ��_In := Null;
          ͣ����ֹʱ��_In := Null;
          ͣ��ԭ��_In     := Null;
      End;
    End If;
  
    --ͣ�ﰲ��
    Begin
      Select a.��ʼʱ��, a.��ֹʱ��, a.ͣ��ԭ��
      Into ͣ�￪ʼʱ��_In, ͣ����ֹʱ��_In, ͣ��ԭ��_In
      From �ٴ�����ͣ���¼ A, �ٴ������Դ B
      Where a.������ = b.ҽ������ And a.��¼id Is Null And a.����ʱ�� Is Not Null And a.ȡ���� Is Null And b.ҽ��id Is Not Null And
            b.Id = ��Դid_In And Not (��ʼʱ��_In >= a.��ֹʱ�� Or ��ֹʱ��_In <= a.��ʼʱ��) And Rownum < 2;
    
      If ͣ�￪ʼʱ��_In < ��ʼʱ��_In Then
        ͣ�￪ʼʱ��_In := ��ʼʱ��_In;
      End If;
      If ͣ����ֹʱ��_In > ��ֹʱ��_In Then
        ͣ����ֹʱ��_In := ��ֹʱ��_In;
      End If;
    
      Return 2;
    Exception
      When Others Then
        ͣ�￪ʼʱ��_In := Null;
        ͣ����ֹʱ��_In := Null;
        ͣ��ԭ��_In     := Null;
    End;
  
    Return - 1;
  Exception
    When Others Then
      Return 0;
  End Fun_Isclinicvisit;
Begin
  Begin
    Select Nvl(�Ű෽ʽ, 0) Into n_�Ű෽ʽ From �ٴ������ Where ID = Id_In;
  Exception
    When Others Then
      v_Err_Msg := '�������Ϣδ�ҵ���';
      Raise Err_Item;
  End;

  If Nvl(ȡ������_In, 0) = 0 Then
    --��������
    If Nvl(n_�Ű෽ʽ, 0) = 0 Then
      Begin
        Select 1
        Into n_Count
        From �ٴ����ﰲ�� A, �ٴ��������� B, �ٴ������ C
        Where a.Id = b.����id And a.����id = c.Id And c.�Ű෽ʽ = 0 And c.Id = Id_In And Rownum < 2;
      Exception
        When Others Then
          n_Count := 0;
      End;
      If n_Count = 0 Then
        v_Err_Msg := '��ǰ���������Ч�İ��ţ����ܷ�����';
        Raise Err_Item;
      End If;
    Else
      Begin
        Select 1
        Into n_Count
        From �ٴ����ﰲ�� A, �ٴ������¼ B, �ٴ������ C
        Where a.Id = b.����id And a.����id = c.Id And c.�Ű෽ʽ In (1, 2) And c.Id = Id_In And Rownum < 2;
      Exception
        When Others Then
          n_Count := 0;
      End;
      If n_Count = 0 Then
        v_Err_Msg := '��ǰ���������Ч�İ��ţ����ܷ�����';
        Raise Err_Item;
      End If;
    
      Begin
        Select 1
        Into n_Count
        From �ٴ������¼ A, �ٴ����ﰲ�� B
        Where a.��Դid = b.��Դid And a.�������� Between b.��ʼʱ�� And b.��ֹʱ�� And a.����id <> b.Id And b.����id = Id_In And Rownum < 2;
      Exception
        When Others Then
          n_Count := 0;
      End;
      If n_Count <> 0 Then
        v_Err_Msg := '��ǰ������еĲ��ֺ�Դ�ڵ�ǰ��������Чʱ�䷶Χ���Ѿ�������Ч�İ��ţ����ܷ�����';
        Raise Err_Item;
      End If;
    
      Begin
        Select 1 Into n_Count From �ٴ����ﰲ�� A Where a.����id = Id_In And a.��ʼʱ�� < Sysdate And Rownum < 2;
      Exception
        When Others Then
          n_Count := 0;
      End;
      If n_Count <> 0 Then
        v_Err_Msg := '��ǰʱ������˳����Ŀ�ʼʱ�䣬���ܷ�����';
        Raise Err_Item;
      End If;
    End If;
  
    --������ڶ��δ�����İ��ű����������������ڵİ��ţ����밴��С��Чʱ����з���
    Begin
      Select 1
      Into n_Count
      From �ٴ����ﰲ�� A, �ٴ������ B, �ٴ����ﰲ�� C
      Where a.����id = b.Id And b.�Ű෽ʽ = Nvl(n_�Ű෽ʽ, 0) And a.��Դid = c.��Դid And a.��ʼʱ�� < c.��ʼʱ�� And b.Id <> c.����id And
            b.������ Is Null And c.����id = Id_In And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count <> 0 Then
      v_Err_Msg := '��ǰ���ڶ��δ�����İ��ű����밴��С��ʼʱ����з�����';
      Raise Err_Item;
    End If;
  
    Update �ٴ������ Set ������ = ������_In, ����ʱ�� = ����ʱ��_In Where ID = Id_In;
  
    --ɾ������ʱ�а��ţ����Ǻ�Դ�ѱ�ͣ�õļ�¼
    For c_���� In (Select a.Id
                 From �ٴ����ﰲ�� A, �ٴ������Դ B
                 Where a.��Դid = b.Id And a.����id = Id_In And
                       Not (Nvl(b.�Ƿ�ɾ��, 0) = 0 And (b.����ʱ�� Is Null Or b.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd')))) Loop
      Zl_�ٴ����ﰲ��_Delete(c_����.Id, Nvl(n_�Ű෽ʽ, 0));
    End Loop;
  
    --�̶������޸ĵ�ǰ��Ч���ŵ���ֹʱ�䣬ͬһʱ��ͬһ��Դ��Ч�̶�����ֻ����һ��
    If Nvl(n_�Ű෽ʽ, 0) = 0 Then
      For c_���� In (Select a.Id, b.��ʼʱ��
                   From �ٴ����ﰲ�� A, �ٴ����ﰲ�� B, �ٴ������ C
                   Where a.��Դid = b.��Դid And a.��ֹʱ�� > b.��ʼʱ�� And a.����id = c.Id And Nvl(c.�Ű෽ʽ, 0) = 0 And
                         a.����id <> b.����id And c.������ Is Not Null And b.����id = Id_In) Loop
        Update �ٴ����ﰲ�� Set ��ֹʱ�� = c_����.��ʼʱ�� - 1 Where ID = c_����.Id;
      End Loop;
    
      --"���Ű�"/"���Ű�"���������ĺ�Դ�����ڵ�ǰ�̶����ŵ���Чʱ�������г����¼,��Ҫ�����̶����ŵ���Чʱ��
      For c_���� In (Select a.Id, b.��ֹʱ��
                   From �ٴ����ﰲ�� A, �ٴ����ﰲ�� B, �ٴ������ C
                   Where a.��Դid = b.��Դid And b.����id = c.Id And c.�Ű෽ʽ In (1, 2) And a.����id = Id_In And a.��ʼʱ�� < b.��ֹʱ��) Loop
        Update �ٴ����ﰲ�� Set ��ʼʱ�� = c_����.��ֹʱ�� + 1 Where ID = c_����.Id;
      End Loop;
    Else
      --�°���/�ܰ��Ų���ͣ����Ϣ
      For c_��¼ In (Select a.����id, a.��������, a.Id, a.��Դid, a.��ʼʱ��, a.��ֹʱ��
                   From �ٴ������¼ A, �ٴ����ﰲ�� B
                   Where a.����id = b.Id And b.����id = Id_In
                   Order By a.��������, a.�ϰ�ʱ��) Loop
      
        n_���� := Fun_Isclinicvisit(c_��¼.��Դid, c_��¼.��ʼʱ��, c_��¼.��ֹʱ��, d_ͣ�￪ʼʱ��, d_ͣ����ֹʱ��, v_ͣ��ԭ��, d_ԭ�ϰ�����, d_��������);
        If Nvl(n_����, 0) = 1 Or Nvl(n_����, 0) = 2 Then
          --�ڼ��ջ���ͣ�ﰲ��
          Update �ٴ������¼
          Set ͣ�￪ʼʱ�� = d_ͣ�￪ʼʱ��, ͣ����ֹʱ�� = d_ͣ����ֹʱ��, ͣ��ԭ�� = v_ͣ��ԭ��
          Where ID = c_��¼.Id;
        
          --����ͣ���¼
          Insert Into �ٴ�����ͣ���¼
            (ID, ��¼id, ��ʼʱ��, ��ֹʱ��, ͣ��ԭ��, ������, ����ʱ��, ������, ����ʱ��)
          Values
            (�ٴ�����ͣ���¼_Id.Nextval, c_��¼.Id, d_ͣ�￪ʼʱ��, d_ͣ����ֹʱ��, v_ͣ��ԭ��, ������_In, ����ʱ��_In, ������_In, ����ʱ��_In);
        End If;
      
        If Nvl(n_����, 0) = 1 Then
          --���л��ݴ���
          If d_ԭ�ϰ����� Is Not Null Then
            --����Ϣ�ջ����ϰ�����
            Begin
              Select 1
              Into n_Count
              From �ٴ������¼
              Where ��Դid = c_��¼.��Դid And �������� = d_ԭ�ϰ����� And Nvl(�Ƿ񷢲�, 0) = 1 And Rownum < 2;
            Exception
              When Others Then
                n_Count := 0;
            End;
            If n_Count > 0 Then
              --��ɾ�����е�
              Select ID Bulk Collect Into l_��¼id From �ٴ������¼ Where ID = c_��¼.Id;
              Zl_�ٴ������¼_Batchdelete(l_��¼id);
            
              For c_���ݼ�¼ In (Select ID
                             From �ٴ������¼
                             Where ��Դid = c_��¼.��Դid And �������� = d_ԭ�ϰ����� And Nvl(�Ƿ񷢲�, 0) = 1) Loop
                Zl_�ٴ������¼_Copy(c_���ݼ�¼.Id, c_��¼.����id, c_��¼.��������, ������_In, ����ʱ��_In);
              End Loop;
            End If;
          End If;
        
          If d_�������� Is Not Null Then
            --���ϰ����ڻ�����Ϣ�գ���������������еİ����ѱ�ʹ���򲻻���
            Begin
              Select 1
              Into n_Count
              From �ٴ������¼
              Where ��Դid = c_��¼.��Դid And �������� = c_��¼.�������� And Rownum < 2;
            Exception
              When Others Then
                n_Count := 0;
            End;
            If n_Count > 0 Then
              Begin
                Select 1
                Into n_Count
                From �ٴ������¼ A, ���˹Һż�¼ B
                Where a.Id = b.�����¼id And a.��Դid = c_��¼.��Դid And a.�������� = d_�������� And Rownum < 2;
              Exception
                When Others Then
                  n_Count := 0;
              End;
              If n_Count = 0 Then
                --��ɾ�����е�
                Select ID Bulk Collect
                Into l_��¼id
                From �ٴ������¼
                Where ��Դid = c_��¼.��Դid And �������� = d_��������;
                Zl_�ٴ������¼_Batchdelete(l_��¼id);
              
                For c_���ݼ�¼ In (Select ID From �ٴ������¼ Where ��Դid = c_��¼.��Դid And �������� = c_��¼.��������) Loop
                  Zl_�ٴ������¼_Copy(c_���ݼ�¼.Id, c_��¼.����id, d_��������, ������_In, ����ʱ��_In);
                End Loop;
              End If;
            End If;
          End If;
        End If;
      End Loop;
    
      --�޸��ٴ������¼�е�"�Ƿ񷢲�"
      Select a.Id Bulk Collect
      Into l_��¼id
      From �ٴ������¼ A, �ٴ����ﰲ�� B
      Where a.����id = b.Id And b.����id = Id_In;
    
      Forall I In 1 .. l_��¼id.Count
        Update �ٴ������¼ Set �Ƿ񷢲� = 1 Where ID = l_��¼id(I);
    End If;
    Return;
  End If;

  --==================================================================================================================
  --ȡ������
  Begin
    Select 1
    Into n_Count
    From �ٴ����ﰲ�� A, �ٴ����ﰲ�� B, �ٴ������ C, �ٴ������ D
    Where a.��ʼʱ�� > b.��ʼʱ�� And a.��Դid = b.��Դid And a.����id = c.Id And b.����id = d.Id And c.�Ű෽ʽ = d.�Ű෽ʽ And
          c.����ʱ�� Is Not Null And b.����id = Id_In And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count <> 0 Then
    v_Err_Msg := '��ǰ�����ʼʱ��֮�󻹴����ѷ����˵ĳ����������Ե�ǰ��������ȡ��������';
    Raise Err_Item;
  End If;

  Begin
    Select 1
    Into n_Count
    From ���˹Һż�¼ C, �ٴ������¼ A, �ٴ����ﰲ�� B
    Where c.�����¼id = a.Id And a.����id = b.Id And b.����id = Id_In And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count <> 0 Then
    v_Err_Msg := '��ǰ�����İ����ѱ�ʹ�ã�������ȡ��������';
    Raise Err_Item;
  End If;

  Begin
    Select 1 Into n_Count From �ٴ����ﰲ�� A Where a.����id = Id_In And Sysdate > a.��ʼʱ�� And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count <> 0 Then
    Select Nvl(zl_GetSysParameter(256), '|') Into v_�Һ��Ű�ģʽ From Dual;
    n_�Һ��Ű�ģʽ := To_Number(Substr(v_�Һ��Ű�ģʽ || '|', 1, Instr(v_�Һ��Ű�ģʽ || '|', '|') - 1));
    If Nvl(n_�Һ��Ű�ģʽ, 0) = 1 Then
      --û�л��Һ��Ű�ģʽʱ����ȡ������
      v_Err_Msg := '��ǰ�����Ѿ��ڵ�ǰ���ŵ���Чʱ�䷶Χ�ڻ��ߴ����˵�ǰ���ŵ���ֹʱ�䣬������ȡ��������';
      Raise Err_Item;
    End If;
  End If;

  Update �ٴ������ Set ������ = Null, ����ʱ�� = Null Where ID = Id_In;
  If Sql%NotFound Then
    v_Err_Msg := '�������Ϣδ�ҵ���';
    Raise Err_Item;
  End If;

  --�̶�����ȡ������ʱɾ�������¼���ָ�ԭ����
  If Nvl(n_�Ű෽ʽ, 0) = 0 Then
  
    --��ԭ��һ����Ч���ŵ���ֹʱ��
    For c_���� In (Select Distinct a.Id, a.ԭ��ֹʱ��
                 From �ٴ����ﰲ�� A, �ٴ����ﰲ�� B, �ٴ������ C
                 Where a.��Դid = b.��Դid And a.��ֹʱ�� = b.��ʼʱ�� - 1 And a.����id = c.Id And c.�Ű෽ʽ = 0 And c.������ Is Not Null And
                       b.����id = Id_In And a.����id <> Id_In) Loop
      Update �ٴ����ﰲ��
      Set ��ֹʱ�� = Nvl(c_����.ԭ��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd'))
      Where ID = c_����.Id;
    End Loop;
  
    --ɾ�������¼
    Select a.Id Bulk Collect
    Into l_��¼id
    From �ٴ������¼ A, �ٴ����ﰲ�� B
    Where a.����id = b.Id And b.����id = Id_In;
  
    Zl_�ٴ������¼_Batchdelete(l_��¼id);
  Else
    --�°���/�ܰ������ͣ����Ϣ�����޸��Ƿ񷢲�
    Select a.Id Bulk Collect
    Into l_��¼id
    From �ٴ������¼ A, �ٴ����ﰲ�� B
    Where a.����id = b.Id And b.����id = Id_In And a.ͣ�￪ʼʱ�� Is Not Null;
  
    Forall I In 1 .. l_��¼id.Count
      Delete From �ٴ�����ͣ���¼ Where ��¼id = l_��¼id(I);
  
    --�޸��ٴ������¼�е�"�Ƿ񷢲�"
    Select a.Id Bulk Collect
    Into l_��¼id
    From �ٴ������¼ A, �ٴ����ﰲ�� B
    Where a.����id = b.Id And b.����id = Id_In;
  
    Forall I In 1 .. l_��¼id.Count
      Update �ٴ������¼
      Set ͣ�￪ʼʱ�� = Null, ͣ����ֹʱ�� = Null, ͣ��ԭ�� = Null, �Ƿ񷢲� = 0
      Where ID = l_��¼id(I);
  
    --���ݵĲ��ٻָ�
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ����ﰲ��_Publish;
/

Create Or Replace Procedure Zl_�ٴ������_Update
(
  ��������_In Number,
  Id_In       �ٴ������.Id%Type,
  �������_In �ٴ������.�������%Type := Null,
  ��ʼʱ��_In �ٴ����ﰲ��.��ʼʱ��%Type := Null,
  ��ֹʱ��_In �ٴ����ﰲ��.��ֹʱ��%Type := Null,
  Ӧ�÷�Χ_In �ٴ������.Ӧ�÷�Χ%Type := Null,
  ����id_In   �ٴ������.����id%Type := Null,
  ��ע_In     �ٴ������.��ע%Type := Null
) As
  --�����������Ϣ�����ģ��͹̶�����
  --��������_In 1-ģ�壬2-�̶�����
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_Count Number;
Begin
  --ģ��
  If Nvl(��������_In, 0) = 1 Then
    Update �ٴ������
    Set ������� = �������_In, Ӧ�÷�Χ = Ӧ�÷�Χ_In, ����id = ����id_In, ��ע = ��ע_In
    Where ID = Id_In;
    If Sql%NotFound Then
      v_Err_Msg := '�������Ϣδ�ҵ���';
      Raise Err_Item;
    End If;
    Return;
  End If;

  --�̶�����
  Begin
    Select 1 Into n_Count From �ٴ������ Where ������ Is Not Null And ID = Id_In And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count <> 0 Then
    v_Err_Msg := '��ǰ������ѷ�������������е�����';
    Raise Err_Item;
  End If;

  Update �ٴ������ Set ������� = �������_In Where ID = Id_In;
  If Sql%NotFound Then
    v_Err_Msg := '�������Ϣδ�ҵ���';
    Raise Err_Item;
  End If;

  Update �ٴ����ﰲ��
  Set ��ʼʱ�� = Nvl(��ʼʱ��_In, ��ʼʱ��), ��ֹʱ�� = Nvl(��ֹʱ��_In, ��ֹʱ��), ����Ա���� = Nvl(����Ա����, Zl_Username), �Ǽ�ʱ�� = Nvl(�Ǽ�ʱ��, Sysdate),
      ԭ��ֹʱ�� = Nvl(��ֹʱ��_In, ԭ��ֹʱ��)
  Where ����id = Id_In;
  If Sql%NotFound Then
    --����һ���޺�Դ�ĳ��ﰲ�ţ����ڼ�¼��������Ϣ
    Insert Into �ٴ����ﰲ��
      (ID, ����id, ��Դid, ��ʼʱ��, ��ֹʱ��, ����Ա����, �Ǽ�ʱ��, ԭ��ֹʱ��)
    Values
      (�ٴ����ﰲ��_Id.Nextval, Id_In, Null, ��ʼʱ��_In, ��ֹʱ��_In, Zl_Username, Sysdate, ��ֹʱ��_In);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ������_Update;
/
Create Or Replace Procedure Zl_�ٴ����ﰲ��_Delete
(
  Id_In       �ٴ����ﰲ��.Id%Type,
  �����¼_In Number := 0
) As
  --���ܣ�ɾ���ٴ����ﰲ��
  --������

  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  l_����id t_Numlist := t_Numlist();
  l_��¼id t_Numlist := t_Numlist();
Begin

  If Nvl(�����¼_In, 0) = 0 Then
    --ɾ���ٴ��������/ģ��
    Select ID Bulk Collect Into l_����id From �ٴ��������� Where ����id = Id_In;
  
    Forall I In 1 .. l_����id.Count
      Delete From �ٴ�����Һſ��� Where ����id = l_����id(I);
  
    Forall I In 1 .. l_����id.Count
      Delete From �ٴ�����ʱ�� Where ����id = l_����id(I);
  
    Forall I In 1 .. l_����id.Count
      Delete From �ٴ��������� Where ����id = l_����id(I);
  
    Delete From �ٴ��������� Where ����id = Id_In;
  Else
    --ɾ���ٴ������¼
    Select ID Bulk Collect Into l_��¼id From �ٴ������¼ Where ����id = Id_In;
  
    Zl_�ٴ������¼_Batchdelete(l_��¼id);
  End If;
  Delete From �ٴ����ﰲ�� Where ID = Id_In;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ����ﰲ��_Delete;
/
Create Or Replace Procedure Zl_�ٴ����ﰲ��_��ſ���
(
  ����id_In   �ٴ������.Id%Type,
  ��ſ���_In �ٴ���������.�Ƿ���ſ���%Type,
  վ��_In     ���ű�.վ��%Type := Null,
  ��Աid_In   ��Ա��.Id%Type := 0
) As
  --ȫ��������ſ��ƻ���ȫ��ȡ����ſ���
  --������
  --      ��Աid_In ������0���޸���Ա���ڿ��ҵ����к�Դ���ţ������޸����к�Դ�İ���
  n_Count    Number(2);
  n_�����¼ Number(2);

  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  l_����id t_Numlist := t_Numlist();

  --���α����ڶ�ȡ�����ٴ����ﰲ�ŵ�ID
  Cursor c_����
  (
    ����id_In �ٴ������.Id%Type,
    ��Աid_In ��Ա��.Id%Type := 0
  ) Is
    Select b.Id
    From �ٴ����ﰲ�� B, �ٴ������Դ C
    Where b.��Դid = c.Id And b.����id = ����id_In And
          (Nvl(��Աid_In, 0) = 0 Or
          (Nvl(��Աid_In, 0) <> 0 And Exists (Select 1 From ������Ա Where ����id = c.����id And ��Աid = ��Աid_In))) And Exists
     (Select 1 From ���ű� Where ID = c.����id And (վ��_In Is Null Or (վ�� Is Null Or վ�� = վ��_In)));
Begin
  Select Count(1)
  Into n_Count
  From �ٴ������ A
  Where a.Id = ����id_In And a.������ Is Not Null And a.�Ű෽ʽ <> 3 And Rownum < 2;
  If n_Count <> 0 Then
    v_Err_Msg := '��ǰ������ѷ������������޸ģ�';
    Raise Err_Item;
  End If;

  Select Count(1) Into n_Count From �ٴ������ A Where a.Id = ����id_In And a.�Ű෽ʽ In (1, 2) And Rownum < 2;
  If n_Count <> 0 Then
    n_�����¼ := 1;
  End If;

  Open c_����(����id_In, ��Աid_In);
  Fetch c_���� Bulk Collect
    Into l_����id;
  Close c_����;

  If Nvl(n_�����¼, 0) = 0 Then
    --�ٴ��������ƻ�ģ��
    Forall I In 1 .. l_����id.Count
      Update �ٴ���������
      Set �Ƿ���ſ��� = ��ſ���_In
      Where (�޺��� Is Not Null Or ��Լ�� Is Not Null) And ����id = l_����id(I);
  
  Else
    --�ٴ������¼
    Forall I In 1 .. l_����id.Count
      Update �ٴ������¼
      Set �Ƿ���ſ��� = ��ſ���_In
      Where (�޺��� Is Not Null Or ��Լ�� Is Not Null) And ����id = l_����id(I);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ����ﰲ��_��ſ���;
/
Create Or Replace Procedure Zl_�ٴ������ϰ�ʱ��_Delete
(
  ����id_In   �ٴ���������.����id%Type,
  ��Ŀ_In     �ٴ���������.������Ŀ%Type,
  �����¼_In Number := 0
) As
  --���ܣ�ɾ���ٴ��������/��¼
  --������
  --      �����¼_In:�Ƿ��ǶԳ����¼����ɾ��
  --      ɾ�����ﰲ��_In:ɾ������ʱ��ʱ�Ƿ�ɾ�����ż�¼
  l_����id t_Numlist := t_Numlist();
  l_��¼id t_Numlist := t_Numlist();

  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin
  If Nvl(�����¼_In, 0) = 0 Then
    --ɾ���ٴ��������/ģ��
    Select ID Bulk Collect Into l_����id From �ٴ��������� Where ����id = ����id_In And ������Ŀ = ��Ŀ_In;
  
    Forall I In 1 .. l_����id.Count
      Delete From �ٴ�����Һſ��� Where ����id = l_����id(I);
  
    Forall I In 1 .. l_����id.Count
      Delete From �ٴ�����ʱ�� Where ����id = l_����id(I);
  
    Forall I In 1 .. l_����id.Count
      Delete From �ٴ��������� Where ����id = l_����id(I);
  
    Forall I In 1 .. l_����id.Count
      Delete From �ٴ��������� Where ID = l_����id(I);
  Else
    --ɾ���ٴ������¼
    Select ID Bulk Collect
    Into l_��¼id
    From �ٴ������¼
    Where ����id = ����id_In And �������� = To_Date(��Ŀ_In, 'yyyy-mm-dd');
  
    Zl_�ٴ������¼_Batchdelete(l_��¼id);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ������ϰ�ʱ��_Delete;
/
Create Or Replace Procedure Zl_�ٴ����ﰲ��_Insert
(
  Id_In           �ٴ����ﰲ��.Id%Type,
  ����id_In       �ٴ����ﰲ��.����id%Type,
  ��Դid_In       �ٴ����ﰲ��.��Դid%Type,
  ��Ŀid_In       �ٴ����ﰲ��.��Ŀid%Type,
  ҽ��id_In       �ٴ����ﰲ��.ҽ��id%Type,
  ҽ������_In     �ٴ����ﰲ��.ҽ������%Type,
  �Ű����_In     �ٴ����ﰲ��.�Ű����%Type,
  �Ƿ���������_In �ٴ����ﰲ��.�Ƿ���������%Type,
  �Ƿ����ճ���_In �ٴ����ﰲ��.�Ƿ����ճ���%Type,
  ��ʼʱ��_In     �ٴ����ﰲ��.��ʼʱ��%Type,
  ��ֹʱ��_In     �ٴ����ﰲ��.��ֹʱ��%Type,
  ����Ա����_In   �ٴ����ﰲ��.����Ա����%Type,
  �Ǽ�ʱ��_In     �ٴ����ﰲ��.�Ǽ�ʱ��%Type
) As
  --���ܣ����������ٴ����ﰲ��
  --������
  n_Count Number;

  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin

  Update �ٴ����ﰲ��
  Set ����id = ����id_In, ��Դid = ��Դid_In, ��Ŀid = ��Ŀid_In, ҽ��id = ҽ��id_In, ҽ������ = ҽ������_In, �Ű���� = �Ű����_In, �Ƿ��������� = �Ƿ���������_In,
      �Ƿ����ճ��� = �Ƿ����ճ���_In, ��ʼʱ�� = ��ʼʱ��_In, ��ֹʱ�� = ��ֹʱ��_In, ����Ա���� = ����Ա����_In, �Ǽ�ʱ�� = �Ǽ�ʱ��_In, ԭ��ֹʱ�� = ��ֹʱ��_In
  Where ID = Id_In;
  If Sql% NotFound Then
    Insert Into �ٴ����ﰲ��
      (ID, ����id, ��Դid, ��Ŀid, ҽ��id, ҽ������, �Ű����, �Ƿ���������, �Ƿ����ճ���, ��ʼʱ��, ��ֹʱ��, ����Ա����, �Ǽ�ʱ��, ԭ��ֹʱ��)
    Values
      (Id_In, ����id_In, ��Դid_In, ��Ŀid_In, ҽ��id_In, ҽ������_In, �Ű����_In, �Ƿ���������_In, �Ƿ����ճ���_In, ��ʼʱ��_In, ��ֹʱ��_In, ����Ա����_In,
       �Ǽ�ʱ��_In, ��ֹʱ��_In);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ����ﰲ��_Insert;
/
Create Or Replace Procedure Zl_�ٴ���������_Insert
(
  Id_In           �ٴ���������.Id%Type,
  ����id_In       �ٴ���������.����id%Type,
  ������Ŀ_In     �ٴ���������.������Ŀ%Type,
  �ϰ�ʱ��_In     �ٴ���������.�ϰ�ʱ��%Type,
  �޺���_In       �ٴ���������.�޺���%Type,
  ��Լ��_In       �ٴ���������.��Լ��%Type,
  �Ƿ��ʱ��_In   �ٴ���������.�Ƿ��ʱ��%Type,
  �Ƿ���ſ���_In �ٴ���������.�Ƿ���ſ���%Type,
  ԤԼ����_In     �ٴ���������.ԤԼ����%Type,
  �Ƿ��ռ_In     �ٴ���������.�Ƿ��ռ%Type,
  ���﷽ʽ_In     �ٴ���������.���﷽ʽ%Type := Null,
  ����_In         Varchar2 := Null,
  ʱ��_In         Varchar2 := Null,
  ɾ�����_In     Number := 0
) As
  --���ܣ����������ٴ���������
  --������
  --     ����_In:����1,����2,...
  --     ʱ��_In:���,��ʼʱ��,��ֹʱ��,��������,ԤԼ��־|...
  --     ɾ�����_In:�Ƿ�ɾ���������ʱ��
  v_���� Varchar2(100);
  n_���� �ٴ���������.����id%Type;

  v_ʱ��     Varchar2(5000);
  n_���     �ٴ�����ʱ��.���%Type;
  d_��ʼʱ�� �ٴ�����ʱ��.��ʼʱ��%Type;
  d_��ֹʱ�� �ٴ�����ʱ��.��ֹʱ��%Type;
  n_�������� �ٴ�����ʱ��.��������%Type;
  n_�Ƿ�ԤԼ �ٴ�����ʱ��.�Ƿ�ԤԼ%Type;
  v_��ǰ��� Varchar2(100);

  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin

  Update �ٴ���������
  Set �޺��� = �޺���_In, ��Լ�� = ��Լ��_In, �Ƿ��ʱ�� = �Ƿ��ʱ��_In, �Ƿ���ſ��� = �Ƿ���ſ���_In, ԤԼ���� = ԤԼ����_In, �Ƿ��ռ = �Ƿ��ռ_In, ���﷽ʽ = ���﷽ʽ_In,
      ����id = Null
  Where ID = Id_In;
  If Sql% NotFound Then
    Insert Into �ٴ���������
      (ID, ����id, ������Ŀ, �ϰ�ʱ��, �޺���, ��Լ��, �Ƿ��ʱ��, �Ƿ���ſ���, ԤԼ����, �Ƿ��ռ, ���﷽ʽ)
    Values
      (Id_In, ����id_In, ������Ŀ_In, �ϰ�ʱ��_In, �޺���_In, ��Լ��_In, �Ƿ��ʱ��_In, �Ƿ���ſ���_In, ԤԼ����_In, �Ƿ��ռ_In, ���﷽ʽ_In);
  End If;

  Delete From �ٴ��������� Where ����id = Id_In;
  --��������
  If ����_In Is Not Null Then
    v_���� := ����_In || ',';
  End If;
  While v_���� Is Not Null Loop
    n_���� := To_Number(Substr(v_����, 1, Instr(v_����, ',') - 1));
    If Nvl(���﷽ʽ_In, 0) = 1 Then
      Update �ٴ��������� Set ����id = n_���� Where ID = Id_In;
    End If;
    Insert Into �ٴ��������� (����id, ����id) Values (Id_In, n_����);
    v_���� := Substr(v_����, Instr(v_����, ',') + 1);
  End Loop;

  --����ʱ��
  If Nvl(ɾ�����_In, 0) = 1 Then
    --ɾ���������ʱ��
    Delete �ٴ�����ʱ�� Where ����id = Id_In;
  End If;
  If ʱ��_In Is Not Null Then
    v_ʱ�� := ʱ��_In || '|';
  End If;
  While v_ʱ�� Is Not Null Loop
    v_��ǰ��� := Substr(v_ʱ��, 1, Instr(v_ʱ��, '|') - 1);
    n_���     := To_Number(Substr(v_��ǰ���, 1, Instr(v_��ǰ���, ',') - 1));
    v_��ǰ��� := Substr(v_��ǰ���, Instr(v_��ǰ���, ',') + 1);
    d_��ʼʱ�� := To_Date(Substr(v_��ǰ���, 1, Instr(v_��ǰ���, ',') - 1), 'yyyy-mm-dd hh24:mi:ss');
    v_��ǰ��� := Substr(v_��ǰ���, Instr(v_��ǰ���, ',') + 1);
    d_��ֹʱ�� := To_Date(Substr(v_��ǰ���, 1, Instr(v_��ǰ���, ',') - 1), 'yyyy-mm-dd hh24:mi:ss');
    v_��ǰ��� := Substr(v_��ǰ���, Instr(v_��ǰ���, ',') + 1);
    n_�������� := To_Number(Substr(v_��ǰ���, 1, Instr(v_��ǰ���, ',') - 1));
    v_��ǰ��� := Substr(v_��ǰ���, Instr(v_��ǰ���, ',') + 1);
    n_�Ƿ�ԤԼ := To_Number(v_��ǰ���);
    If Nvl(n_���, 0) <> 0 Then
      Insert Into �ٴ�����ʱ��
        (����id, ���, ��ʼʱ��, ��ֹʱ��, �Ƿ�ԤԼ, ��������)
      Values
        (Id_In, n_���, d_��ʼʱ��, d_��ֹʱ��, n_�Ƿ�ԤԼ, n_��������);
    End If;
    v_ʱ�� := Substr(v_ʱ��, Instr(v_ʱ��, '|') + 1);
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ���������_Insert;
/
Create Or Replace Procedure Zl_�ٴ�����Һſ���_Insert
(
  ����id_In   �ٴ�����Һſ���.����id%Type,
  ����_In     �ٴ�����Һſ���.����%Type,
  ����_In     �ٴ�����Һſ���.����%Type,
  ����_In     �ٴ�����Һſ���.����%Type,
  ���Ʒ�ʽ_In �ٴ�����Һſ���.���Ʒ�ʽ%Type,
  �Ƿ��ռ_In �ٴ���������.�Ƿ��ռ%Type,
  ���ſ���_In Varchar2,
  ɾ��_In     Number := 0
) As
  --����:���������ٴ�����Һſ���
  --����:
  --    ����_In:1-��������;2-ԤԼ��ʽ
  --    ���ſ���_in:���1,����|���2,����|...
  --    ɾ��_in:�Ƿ�ɾ�����е�
  v_���     Varchar2(5000);
  v_��ǰ��Ŀ Varchar2(5000);
  n_���     �ٴ�����Һſ���.���%Type;
  n_����     �ٴ�����Һſ���.����%Type;
  n_Count    Number;

  Err_Item Exception;
  v_Err_Msg Varchar2(100);
Begin
  If Nvl(����_In, 0) = 1 Then
    --������λ
    Select Count(1) Into n_Count From �Һź�����λ Where ���� = ����_In;
    If n_Count = 0 Then
      v_Err_Msg := '[ZLSOFT]�Һź�����λδ�ҵ������飡[ZLSOFT]';
      Raise Err_Item;
    End If;
  Elsif Nvl(����_In, 0) = 2 Then
    --ԤԼ��ʽ
    Select Count(1) Into n_Count From ԤԼ��ʽ Where ���� = ����_In;
    If n_Count = 0 Then
      v_Err_Msg := '[ZLSOFT]�Һ�ԤԼ��ʽδ�ҵ������飡[ZLSOFT]';
      Raise Err_Item;
    End If;
  End If;

  Update �ٴ��������� Set �Ƿ��ռ = �Ƿ��ռ_In Where ID = ����id_In;

  If Nvl(ɾ��_In, 0) = 1 Then
    --ɾ�����е�
    Delete From �ٴ�����Һſ��� Where ����id = ����id_In And ���� = ����_In And ���� = ����_In And ���� = ����_In;
  End If;

  v_��� := ���ſ���_In || '|';
  While v_��� Is Not Null Loop
    v_��ǰ��Ŀ := Substr(v_���, 1, Instr(v_���, '|') - 1);
    n_���     := To_Number(Substr(v_��ǰ��Ŀ, 1, Instr(v_��ǰ��Ŀ, ',') - 1));
    v_��ǰ��Ŀ := Substr(v_��ǰ��Ŀ, Instr(v_��ǰ��Ŀ, ',') + 1);
    n_����     := To_Number(v_��ǰ��Ŀ);
    If Nvl(n_����, 0) <> 0 Then
      Insert Into �ٴ�����Һſ���
        (����id, ����, ����, ����, ���, ����, ���Ʒ�ʽ)
      Values
        (����id_In, ����_In, ����_In, ����_In, n_���, n_����, ���Ʒ�ʽ_In);
    End If;
    v_��� := Substr(v_���, Instr(v_���, '|') + 1);
  End Loop;

  --ÿһ��������λ����ԤԼ��ʽ���ٵ���һ����¼
  Select Count(1)
  Into n_Count
  From �ٴ�����Һſ���
  Where ����id = ����id_In And ���� = ����_In And ���� = ����_In And ���� = ����_In;
  If n_Count = 0 Then
    Insert Into �ٴ�����Һſ���
      (����id, ����, ����, ����, ���, ����, ���Ʒ�ʽ)
    Values
      (����id_In, ����_In, ����_In, ����_In, 0, 0, ���Ʒ�ʽ_In);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ�����Һſ���_Insert;
/
Create Or Replace Procedure Zl_�ٴ������¼_Insert
(
  Id_In           �ٴ������¼.Id%Type,
  ����id_In       �ٴ������¼.����id%Type,
  ��Դid_In       �ٴ������¼.��Դid%Type,
  ��������_In     �ٴ������¼.��������%Type,
  �ϰ�ʱ��_In     �ٴ������¼.�ϰ�ʱ��%Type,
  ��ʼʱ��_In     �ٴ������¼.��ʼʱ��%Type,
  ��ֹʱ��_In     �ٴ������¼.��ֹʱ��%Type,
  ȱʡԤԼʱ��_In �ٴ������¼.ȱʡԤԼʱ��%Type,
  ��ǰ�Һ�ʱ��_In �ٴ������¼.��ǰ�Һ�ʱ��%Type,
  �޺���_In       �ٴ������¼.�޺���%Type,
  ��Լ��_In       �ٴ������¼.��Լ��%Type,
  �Ƿ���ſ���_In �ٴ������¼.�Ƿ���ſ���%Type,
  �Ƿ��ʱ��_In   �ٴ������¼.�Ƿ��ʱ��%Type,
  ԤԼ����_In     �ٴ������¼.ԤԼ����%Type,
  �Ƿ��ռ_In     �ٴ������¼.�Ƿ��ռ%Type,
  ��Ŀid_In       �ٴ������¼.��Ŀid%Type,
  ����id_In       �ٴ������¼.����id%Type,
  ҽ��id_In       �ٴ������¼.ҽ��id%Type,
  ҽ������_In     �ٴ������¼.ҽ������%Type,
  ���﷽ʽ_In     �ٴ������¼.���﷽ʽ%Type,
  �Ƿ���ʱ����_In �ٴ������¼.�Ƿ���ʱ����%Type,
  �Ǽ���_In       �ٴ������¼.�Ǽ���%Type,
  �Ǽ�ʱ��_In     �ٴ������¼.�Ǽ�ʱ��%Type,
  �Ƿ񷢲�_In     �ٴ������¼.�Ƿ񷢲�%Type,
  ����_In         Varchar2 := Null,
  ʱ��_In         Varchar2 := Null,
  ɾ�����_In     Number := 0
) As
  --���ܣ����������ٴ������¼
  --������
  --     ����_In:����1,����2,...
  --     ʱ��_In:���,��ʼʱ��,��ֹʱ��,��������,ԤԼ��־|...
  --     ɾ�����_In:�Ƿ�ɾ���������ʱ��
  v_ʱ��     Varchar2(5000);
  n_���     �ٴ�������ſ���.���%Type;
  d_��ʼʱ�� �ٴ�������ſ���.��ʼʱ��%Type;
  d_��ֹʱ�� �ٴ�������ſ���.��ֹʱ��%Type;
  n_����     �ٴ�������ſ���.����%Type;
  n_�Ƿ�ԤԼ �ٴ�������ſ���.�Ƿ�ԤԼ%Type;
  v_��ǰ��� Varchar2(100);

  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin
  Update �ٴ������¼
  Set ����id = ����id_In, ��Դid = ��Դid_In, �������� = ��������_In, �ϰ�ʱ�� = �ϰ�ʱ��_In, ��ʼʱ�� = ��ʼʱ��_In, ��ֹʱ�� = ��ֹʱ��_In, ȱʡԤԼʱ�� = ȱʡԤԼʱ��_In,
      ��ǰ�Һ�ʱ�� = ��ǰ�Һ�ʱ��_In, �޺��� = �޺���_In, ��Լ�� = ��Լ��_In, �Ƿ���ſ��� = �Ƿ���ſ���_In, �Ƿ��ʱ�� = �Ƿ��ʱ��_In, ԤԼ���� = ԤԼ����_In,
      �Ƿ��ռ = �Ƿ��ռ_In, ��Ŀid = ��Ŀid_In, ����id = ����id_In, ҽ��id = ҽ��id_In, ҽ������ = ҽ������_In, ���﷽ʽ = ���﷽ʽ_In, �Ƿ���ʱ���� = �Ƿ���ʱ����_In,
      �Ǽ��� = �Ǽ���_In, �Ǽ�ʱ�� = �Ǽ�ʱ��_In, ����id = Null, �Ƿ񷢲� = �Ƿ񷢲�_In
  Where ID = Id_In;
  If Sql% NotFound Then
    Insert Into �ٴ������¼
      (ID, ����id, ��Դid, ��������, �ϰ�ʱ��, ��ʼʱ��, ��ֹʱ��, ȱʡԤԼʱ��, ��ǰ�Һ�ʱ��, �޺���, ��Լ��, �Ƿ���ſ���, �Ƿ��ʱ��, ԤԼ����, �Ƿ��ռ, ��Ŀid, ����id, ҽ��id,
       ҽ������, ���﷽ʽ, �Ƿ���ʱ����, �Ǽ���, �Ǽ�ʱ��, �Ƿ񷢲�)
    Values
      (Id_In, ����id_In, ��Դid_In, ��������_In, �ϰ�ʱ��_In, ��ʼʱ��_In, ��ֹʱ��_In, ȱʡԤԼʱ��_In, ��ǰ�Һ�ʱ��_In, �޺���_In, ��Լ��_In, �Ƿ���ſ���_In,
       �Ƿ��ʱ��_In, ԤԼ����_In, �Ƿ��ռ_In, ��Ŀid_In, ����id_In, ҽ��id_In, ҽ������_In, ���﷽ʽ_In, �Ƿ���ʱ����_In, �Ǽ���_In, �Ǽ�ʱ��_In, �Ƿ񷢲�_In);
  End If;

  Delete From �ٴ��������Ҽ�¼ Where ��¼id = Id_In;
  --��������
  If ����_In Is Not Null Then
    Insert Into �ٴ��������Ҽ�¼
      (��¼id, ����id)
      Select Id_In, Column_Value From Table(f_Str2list(����_In));
  
    If Nvl(���﷽ʽ_In, 0) = 1 Then
      Update �ٴ������¼ Set ����id = To_Number(����_In) Where ID = Id_In;
    End If;
  End If;

  --����ʱ��
  If Nvl(ɾ�����_In, 0) = 1 Then
    --ɾ���������ʱ��
    Delete �ٴ�������ſ��� Where ��¼id = Id_In;
  End If;
  If ʱ��_In Is Not Null Then
    v_ʱ�� := ʱ��_In || '|';
  End If;
  While v_ʱ�� Is Not Null Loop
    v_��ǰ��� := Substr(v_ʱ��, 1, Instr(v_ʱ��, '|') - 1);
    n_���     := To_Number(Substr(v_��ǰ���, 1, Instr(v_��ǰ���, ',') - 1));
    v_��ǰ��� := Substr(v_��ǰ���, Instr(v_��ǰ���, ',') + 1);
    d_��ʼʱ�� := To_Date(Substr(v_��ǰ���, 1, Instr(v_��ǰ���, ',') - 1), 'yyyy-mm-dd hh24:mi:ss');
    v_��ǰ��� := Substr(v_��ǰ���, Instr(v_��ǰ���, ',') + 1);
    d_��ֹʱ�� := To_Date(Substr(v_��ǰ���, 1, Instr(v_��ǰ���, ',') - 1), 'yyyy-mm-dd hh24:mi:ss');
    v_��ǰ��� := Substr(v_��ǰ���, Instr(v_��ǰ���, ',') + 1);
    n_����     := To_Number(Substr(v_��ǰ���, 1, Instr(v_��ǰ���, ',') - 1));
    v_��ǰ��� := Substr(v_��ǰ���, Instr(v_��ǰ���, ',') + 1);
    n_�Ƿ�ԤԼ := To_Number(v_��ǰ���);
    If Nvl(n_���, 0) > 0 Then
      Insert Into �ٴ�������ſ���
        (��¼id, ���, ��ʼʱ��, ��ֹʱ��, �Ƿ�ԤԼ, ����)
      Values
        (Id_In, n_���, d_��ʼʱ��, d_��ֹʱ��, n_�Ƿ�ԤԼ, n_����);
    End If;
    v_ʱ�� := Substr(v_ʱ��, Instr(v_ʱ��, '|') + 1);
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ������¼_Insert;
/
Create Or Replace Procedure Zl_�ٴ�����Һſ��Ƽ�¼_Insert
(
  ��¼id_In   �ٴ�����Һſ��Ƽ�¼.��¼id%Type,
  ����_In     �ٴ�����Һſ��Ƽ�¼.����%Type,
  ����_In     �ٴ�����Һſ��Ƽ�¼.����%Type,
  ����_In     �ٴ�����Һſ��Ƽ�¼.����%Type,
  ���Ʒ�ʽ_In �ٴ�����Һſ��Ƽ�¼.���Ʒ�ʽ%Type,
  �Ƿ��ռ_In �ٴ������¼.�Ƿ��ռ%Type,
  ���ſ���_In Varchar2,
  ɾ��_In     Number := 0
) As
  --����:���������ٴ�����Һſ��Ƽ�¼
  --����:
  --    ����_In:1-��������;2-ԤԼ��ʽ
  --    ���ſ���_in:���1,����|���2,����|...
  --    ɾ��_in:�Ƿ�ɾ�����е�
  v_���     Varchar2(5000);
  v_��ǰ��Ŀ Varchar2(5000);
  n_���     �ٴ�����Һſ��Ƽ�¼.���%Type;
  n_����     �ٴ�����Һſ��Ƽ�¼.����%Type;
  n_Count    Number;

  Err_Item Exception;
  v_Err_Msg Varchar2(100);
Begin
  If Nvl(����_In, 0) = 1 Then
    --������λ
    Select Count(1) Into n_Count From �Һź�����λ Where ���� = ����_In;
    If n_Count = 0 Then
      v_Err_Msg := '[ZLSOFT]�Һź�����λδ�ҵ������飡[ZLSOFT]';
      Raise Err_Item;
    End If;
  Elsif Nvl(����_In, 0) = 2 Then
    --ԤԼ��ʽ
    Select Count(1) Into n_Count From ԤԼ��ʽ Where ���� = ����_In;
    If n_Count = 0 Then
      v_Err_Msg := '[ZLSOFT]�Һ�ԤԼ��ʽδ�ҵ������飡[ZLSOFT]';
      Raise Err_Item;
    End If;
  End If;

  Update �ٴ������¼ Set �Ƿ��ռ = �Ƿ��ռ_In Where ID = ��¼id_In;

  If Nvl(ɾ��_In, 0) = 1 Then
    --ɾ�����е�
    Delete From �ٴ�����Һſ��Ƽ�¼ Where ��¼id = ��¼id_In And ���� = ����_In And ���� = ����_In And ���� = ����_In;
  End If;

  v_��� := ���ſ���_In || '|';
  While v_��� Is Not Null Loop
    v_��ǰ��Ŀ := Substr(v_���, 1, Instr(v_���, '|') - 1);
    n_���     := To_Number(Substr(v_��ǰ��Ŀ, 1, Instr(v_��ǰ��Ŀ, ',') - 1));
    v_��ǰ��Ŀ := Substr(v_��ǰ��Ŀ, Instr(v_��ǰ��Ŀ, ',') + 1);
    n_����     := To_Number(v_��ǰ��Ŀ);
    If Nvl(n_����, 0) <> 0 Then
      Insert Into �ٴ�����Һſ��Ƽ�¼
        (��¼id, ����, ����, ����, ���, ����, ���Ʒ�ʽ)
      Values
        (��¼id_In, ����_In, ����_In, ����_In, n_���, n_����, ���Ʒ�ʽ_In);
    End If;
    v_��� := Substr(v_���, Instr(v_���, '|') + 1);
  End Loop;

  --ÿһ��������λ����ԤԼ��ʽ���ٵ���һ����¼
  Select Count(1)
  Into n_Count
  From �ٴ�����Һſ��Ƽ�¼
  Where ��¼id = ��¼id_In And ���� = ����_In And ���� = ����_In And ���� = ����_In;
  If n_Count = 0 Then
    Insert Into �ٴ�����Һſ��Ƽ�¼
      (��¼id, ����, ����, ����, ���, ����, ���Ʒ�ʽ)
    Values
      (��¼id_In, ����_In, ����_In, ����_In, 0, 0, ���Ʒ�ʽ_In);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ�����Һſ��Ƽ�¼_Insert;
/
Create Or Replace Procedure Zl_�ٴ���������_Update
(
  Id_In       �ٴ���������.Id%Type,
  ���﷽ʽ_In �ٴ���������.���﷽ʽ%Type := Null,
  ����_In     Varchar2 := Null,
  �����¼_In Number := 0
) As
  --���ܣ������ٴ���������
  --������
  --     ����_In:����1,����2,...
  --     �����¼_In:�Ƿ��ǶԳ����¼����ɾ��
  n_Count  Number;
  n_�䶯id �ٴ�����䶯��¼.Id%Type;
  v_����   �ٴ�����䶯��¼.����������%Type;

  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin
  If Nvl(�����¼_In, 0) = 0 Then
    Update �ٴ��������� Set ���﷽ʽ = ���﷽ʽ_In Where ID = Id_In;
  
    Delete From �ٴ��������� Where ����id = Id_In;
    --��������
    If ����_In Is Not Null Then
    
      Insert Into �ٴ���������
        (����id, ����id)
        Select Id_In, Column_Value From Table(f_Str2list(����_In, ','));
    
      If Nvl(���﷽ʽ_In, 0) = 1 Then
        Update �ٴ��������� Set ����id = To_Number(����_In) Where ID = Id_In;
      End If;
    End If;
    Return;
  End If;

  --�ٴ�����䶯��Ϣ
  Select Count(1)
  Into n_Count
  From �ٴ������ A, �ٴ����ﰲ�� B, �ٴ������¼ C
  Where a.Id = b.����id And b.Id = c.����id And a.������ Is Not Null And c.Id = Id_In;
  If Nvl(n_Count, 0) = 0 Then
    v_Err_Msg := '�����¼�����ڣ�';
    Raise Err_Item;
  End If;

  Select �ٴ�����䶯��¼_Id.Nextval Into n_�䶯id From Dual;
  Insert Into �ٴ�����䶯��¼
    (ID, ��¼id, �䶯����, ԭ���﷽ʽ, ԭ����id, ԭ��������, �ַ��﷽ʽ, ����Ա����, �Ǽ�ʱ��)
    Select n_�䶯id, a.Id, 3, a.���﷽ʽ, a.����id, b.����, ���﷽ʽ_In, Zl_Username, Sysdate
    From �ٴ������¼ A, �������� B
    Where a.����id = b.Id(+) And a.Id = Id_In;

  Insert Into �ٴ�����䶯��ϸ
    (�䶯id, �䶯����, ���, ����id, ��������, ����)
    Select n_�䶯id, 1, ���, ����id, ����, '-'
    From (Select Rownum As ���, a.����id, b.����
           From �ٴ��������Ҽ�¼ A, �������� B
           Where a.����id = b.Id(+) And a.��¼id = Id_In);

  Update �ٴ������¼ Set ���﷽ʽ = ���﷽ʽ_In Where ID = Id_In;
  Delete From �ٴ��������Ҽ�¼ Where ��¼id = Id_In;

  --�ٴ�����䶯����Ϣ
  If ����_In Is Not Null Then
    Insert Into �ٴ��������Ҽ�¼
      (��¼id, ����id)
      Select Id_In, Column_Value From Table(f_Str2list(����_In, ','));
  
    Insert Into �ٴ�����䶯��ϸ
      (�䶯id, �䶯����, ���, ����id, ��������, ����)
      Select n_�䶯id, 2, Rownum, a.Id, a.����, '-'
      From �������� A, (Select Column_Value As ����id From Table(f_Str2list(����_In, ','))) B
      Where a.Id = b.����id;
  
    If Nvl(���﷽ʽ_In, 0) = 1 Then
      Update �ٴ������¼ Set ����id = To_Number(����_In) Where ID = Id_In;
    
      Update �ٴ�����䶯��¼
      Set ������id = To_Number(����_In),
          ���������� =
           (Select ���� From �������� Where ID = To_Number(����_In))
      Where ID = n_�䶯id
      Returning ���������� Into v_����;
      --���˹Һż�¼
      Update ���˹Һż�¼ Set ���� = v_���� Where �����¼id = Id_In;
      --������ü�¼
      Update ������ü�¼
      Set ��ҩ���� = v_����
      Where ��¼���� = 4 And NO In (Select NO From ���˹Һż�¼ Where ��¼���� = 1 And �����¼id = Id_In);
    End If;
  
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ���������_Update;
/
Create Or Replace Procedure Zl_�ٴ������¼_Batchlock
(
  Ids_In      Varchar2,
  ȡ������_In Number := 0
) As
  -- Ids_In �������������������ö��ŷָ�
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_Count Number;
Begin
  If Nvl(ȡ������_In, 0) = 1 Then
    Update �ٴ������¼ Set �Ƿ����� = 0 Where ID In (Select Column_Value From Table(f_Str2list(Ids_In, ',')));
  Else
    Update �ٴ������¼ Set �Ƿ����� = 1 Where ID In (Select Column_Value From Table(f_Str2list(Ids_In, ',')));
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ������¼_Batchlock;
/

Create Or Replace Procedure Zl_�ٴ������¼_Stopvisit
(
  ��¼id_In   �ٴ�����ͣ���¼.��¼id%Type,
  ��ʼʱ��_In �ٴ�����ͣ���¼.��ʼʱ��%Type := Null,
  ��ֹʱ��_In �ٴ�����ͣ���¼.��ֹʱ��%Type := Null,
  ͣ��ԭ��_In �ٴ�����ͣ���¼.ͣ��ԭ��%Type := Null,
  ����Ա_In   �ٴ�����ͣ���¼.������%Type := Null,
  ����ʱ��_In �ٴ�����ͣ���¼.����ʱ��%Type := Null,
  ȡ��ͣ��_In Number := 0
) As
  --���ܣ�ͣ�����ȡ��ͣ��
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_Count Number;
  d_Cur   Date;
  v_����  �ٴ������Դ.����%Type;
Begin
  If Nvl(ȡ��ͣ��_In, 0) = 0 Then
    --ͣ��
    If ��ʼʱ��_In <= Sysdate Then
      v_Err_Msg := 'ͣ��ʱ��Ŀ�ʼʱ��С���˵�ǰʱ�䣬���ܽ���ͣ�������';
      Raise Err_Item;
    End If;
  
    Insert Into �ٴ�����ͣ���¼
      (ID, ��¼id, ��ʼʱ��, ��ֹʱ��, ͣ��ԭ��, ������, ����ʱ��, ������, ����ʱ��)
      Select �ٴ�����ͣ���¼_Id.Nextval, ��¼id_In, ��ʼʱ��_In, ��ֹʱ��_In, ͣ��ԭ��_In, Nvl(a.ҽ������, ����Ա_In), ����ʱ��_In, ����Ա_In, ����ʱ��_In
      From �ٴ������¼ A
      Where ID = ��¼id_In;
  
    Update �ٴ������¼
    Set ͣ�￪ʼʱ�� = ��ʼʱ��_In, ͣ����ֹʱ�� = ��ֹʱ��_In, ͣ��ԭ�� = ͣ��ԭ��_In
    Where ID = ��¼id_In;
  
    Insert Into ���˷�����Ϣ��¼
      (ID, ֪ͨ����, ��¼id, �Һ�id, ��Դid, ����, ����id, ��Ŀid, ҽ��id, ҽ������, ����id, �Ǽ���, �Ǽ�ʱ��, ֪ͨԭ��)
      Select ���˷�����Ϣ��¼_Id.Nextval, 1, ��¼id_In, �Һ�id, ��Դid, ����, ����id, ��Ŀid, ҽ��id, ҽ������, ����id, ����Ա_In, ����ʱ��_In,
             'ҽ��' || ͣ��ԭ��_In || '����ͣ��'
      From (Select b.Id As �Һ�id, c.Id As ��Դid, c.����, c.����id, a.��Ŀid, a.ҽ��id, a.ҽ������, b.����id
             From �ٴ������¼ A, ���˹Һż�¼ B, �ٴ������Դ C
             Where a.Id = b.�����¼id And a.��Դid = c.Id And b.��¼״̬ = 1 And a.Id = ��¼id_In And
                   (b.��¼���� = 1 And b.����ʱ�� Between a.ͣ�￪ʼʱ�� And a.ͣ����ֹʱ�� Or
                   b.��¼���� = 2 And b.ԤԼʱ�� Between a.ͣ�￪ʼʱ�� And a.ͣ����ֹʱ��));
  
    --��Ϣ����
    -- ͣ������(1-ͣ��,2-ȡ��ͣ��),�����¼ID,ͣ�����
    Begin
      Select b.���� Into v_���� From �ٴ������¼ A, �ٴ������Դ B Where a.��Դid = b.Id And a.Id = ��¼id_In;
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
        Using 17, 1 || ',' || ��¼id_In || ',' || v_����;
    Exception
      When Others Then
        Null;
    End;
  Else
    --ȡ��ͣ��
    --���ݼ��
    Select ͣ�￪ʼʱ�� Into d_Cur From �ٴ������¼ Where ID = ��¼id_In And ͣ�￪ʼʱ�� Is Not Null;
    If d_Cur <= Sysdate Then
      v_Err_Msg := 'ͣ��ʱ��Ŀ�ʼʱ����С���˵�ǰʱ�䣬���ܽ���ȡ��ͣ�������';
      Raise Err_Item;
    End If;
    Select Count(1)
    Into n_Count
    From ���˷�����Ϣ��¼
    Where ��¼id = ��¼id_In And ֪ͨ���� = 1 And ������ Is Not Null;
    If n_Count <> 0 Then
      v_Err_Msg := '�ó����¼���ڲ��˷�����Ϣ��Ϣ��¼�����ѱ�����������ȡ��ͣ�������';
      Raise Err_Item;
    End If;
  
    Update �ٴ�����ͣ���¼
    Set ȡ���� = ����Ա_In, ȡ��ʱ�� = ����ʱ��_In
    Where ��¼id = ��¼id_In And ����ҽ������ Is Null And ȡ���� Is Null;
  
    Update �ٴ������¼
    Set ͣ�￪ʼʱ�� = Null, ͣ����ֹʱ�� = Null, ͣ��ԭ�� = Null
    Where ID = ��¼id_In And ͣ�￪ʼʱ�� Is Not Null;
  
    Delete ���˷�����Ϣ��¼ Where ��¼id = ��¼id_In And ֪ͨ���� = 1 And ������ Is Null;
  
    --��Ϣ����
    -- ͣ������(1-ͣ��,2-ȡ��ͣ��),�����¼ID,ͣ�����
    Begin
      Select b.���� Into v_���� From �ٴ������¼ A, �ٴ������Դ B Where a.��Դid = b.Id And a.Id = ��¼id_In;
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
        Using 17, 2 || ',' || ��¼id_In || ',' || v_����;
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
End Zl_�ٴ������¼_Stopvisit;
/


Create Or Replace Function Zl1_Ex_Isdoctorsamelevel
(
  ��ҽ��id_In   In ��Ա��.Id%Type,
  ��ҽ������_In In ��Ա��.����%Type,
  ��ҽ��id_In   In ��Ա��.Id%Type,
  ��ҽ������_In In ��Ա��.����%Type
) Return Number
--����˵�����Ƚ�����ҽ����ְ���С��
  --����˵�����ҺŰ�������ʱ���ã��������ҽ����ְ���Ƿ���ڵ���ԭҽ��ְ��
  --���˵����
  --     ��ҽ��ID����ԱID,Ժ��ҽ������NULL
  --     ��ҽ����������Ա����
  --     ��ҽ��ID����ԱID,Ժ��ҽ������NULL
  --     ��ҽ����������Ա����
  --�������أ�
  --     -1 - ��ҽ����ְ�������ҽ����ְ��
  --     0 - ��ҽ����ְ�������ҽ����ְ��
  --     1 - ��ҽ����ְ��С����ҽ����ְ��
  --˵�������ݡ�רҵ����ְ�����ж�,����С�ı�ʾְ��Խ��,û������רҵ����ְ���ҽ����ʾְ�����
 Is
  n_a Number;
  n_b Number;
Begin
  If Nvl(��ҽ��id_In, 0) = 0 Then
    --Ժ��ҽ��
    n_a := -1;
  Else
    Begin
      Select To_Number(Nvl(b.����, -1))
      Into n_a
      From ��Ա�� A, רҵ����ְ�� B
      Where a.רҵ����ְ�� = b.����(+) And a.Id = ��ҽ��id_In;
    Exception
      When Others Then
        n_a := -1;
    End;
  End If;
  If Nvl(��ҽ��id_In, 0) = 0 Then
    --Ժ��ҽ��
    n_b := -1;
  Else
    Begin
      Select To_Number(Nvl(b.����, -1))
      Into n_b
      From ��Ա�� A, רҵ����ְ�� B
      Where a.רҵ����ְ�� = b.����(+) And a.Id = ��ҽ��id_In;
    Exception
      When Others Then
        n_b := -1;
    End;
  End If;

  If n_a = -1 And n_b = -1 Then
    Return 0;
  Elsif n_a = -1 Then
    Return 1;
  Elsif n_b = -1 Then
    Return - 1;
  Else
    If n_a = n_b Then
      Return 0;
    Elsif n_a > n_b Then
      Return 1;
    Else
      Return - 1;
    End If;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl1_Ex_Isdoctorsamelevel;
/

Create Or Replace Procedure Zl_�ٴ������¼_Replacedoctor
(
  ��¼id_In       �ٴ�����ͣ���¼.��¼id%Type,
  ��ʼʱ��_In     �ٴ�����ͣ���¼.��ʼʱ��%Type := Null,
  ��ֹʱ��_In     �ٴ�����ͣ���¼.��ֹʱ��%Type := Null,
  ͣ��ԭ��_In     �ٴ�����ͣ���¼.ͣ��ԭ��%Type := Null,
  ����ҽ��id_In   �ٴ�����ͣ���¼.����ҽ��id%Type := Null,
  ����ҽ������_In �ٴ�����ͣ���¼.����ҽ������%Type := Null,
  ����Ա����_In   �ٴ�����ͣ���¼.������%Type := Null,
  ����Ա���_In   ��Ա��.���%Type := Null,
  ����ʱ��_In     �ٴ�����ͣ���¼.����ʱ��%Type := Null,
  ȡ������_In     Number := 0
) As
  --���ܣ��������ȡ������
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_Count        Number;
  d_Cur          Date;
  n_Updatedoctor Number(2);
  v_����         �ٴ������Դ.����%Type;
Begin
  If Nvl(ȡ������_In, 0) = 0 Then
    --����
    Begin
      Select 1 Into n_Count From �ٴ������¼ A Where ID = ��¼id_In And ͣ�￪ʼʱ�� Is Not Null;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If Nvl(n_Count, 0) <> 0 Then
      v_Err_Msg := '��ǰ�����¼�ѱ�ͣ����������';
      Raise Err_Item;
    End If;
  
    If ��ʼʱ��_In <= Sysdate Then
      v_Err_Msg := 'ͣ��ʱ��Ŀ�ʼʱ��С���˵�ǰʱ�䣬���ܽ������������';
      Raise Err_Item;
    End If;
  
    If Nvl(����ҽ��id_In, 0) <> 0 Then
      Begin
        Select 1
        Into n_Count
        From �ٴ������¼ A
        Where ID = ��¼id_In And Nvl(ҽ��id, ����ҽ��id) = ����ҽ��id_In And Rownum < 2;
      Exception
        When Others Then
          n_Count := 0;
      End;
      If Nvl(n_Count, 0) <> 0 Then
        v_Err_Msg := '����ҽ������Ϊԭ����ҽ������ѡ������ҽ����';
        Raise Err_Item;
      End If;
    End If;
  
    --�ڸ�ʱ���ڣ�����ҽ�����ܴ��������ĳ��ﰲ��
    --��A[A1,A2],B[B1,B2],��BΪ�ջ���ȫ������A��(A1<=B1,A2>=B2).��ôX[X1,X2]��A-B�н�������
    --(X1>=A1 And X1<=NVL(B1,A2)) Or (X2>=A1 And X2<=NVL(B1,A2)) Or (X1>=NVL(B2,A1) And X1<=A2) Or (X2>=NVL(B2,A1) And X2<=A2)
    If Nvl(����ҽ��id_In, 0) = 0 Then
      Select Count(1)
      Into n_Count
      From �ٴ������¼ A, �ٴ������¼ B
      Where a.�������� = b.�������� And Nvl(a.����ҽ������, a.ҽ������) = ����ҽ������_In And Nvl(a.����ҽ��id, a.ҽ��id) Is Null And b.Id = ��¼id_In And
            ((��ʼʱ��_In Between a.��ʼʱ�� And Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��)) Or (��ֹʱ��_In Between a.��ʼʱ�� And Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��)) Or
            (��ʼʱ��_In Between Nvl(a.ͣ����ֹʱ��, a.��ʼʱ��) And a.��ֹʱ��) Or (��ֹʱ��_In Between Nvl(a.ͣ����ֹʱ��, a.��ʼʱ��) And a.��ֹʱ��));
    Else
      Select Count(1)
      Into n_Count
      From �ٴ������¼ A, �ٴ������¼ B
      Where a.�������� = b.�������� And Nvl(a.����ҽ��id, a.ҽ��id) = ����ҽ��id_In And b.Id = ��¼id_In And
            ((��ʼʱ��_In Between a.��ʼʱ�� And Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��)) Or (��ֹʱ��_In Between a.��ʼʱ�� And Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��)) Or
            (��ʼʱ��_In Between Nvl(a.ͣ����ֹʱ��, a.��ʼʱ��) And a.��ֹʱ��) Or (��ֹʱ��_In Between Nvl(a.ͣ����ֹʱ��, a.��ʼʱ��) And a.��ֹʱ��));
    End If;
    If n_Count <> 0 Then
      v_Err_Msg := '����ҽ��������ʱ�䷶Χ���Ѵ����������ﰲ�ţ���ѡ������ҽ����';
      Raise Err_Item;
    End If;
    --����Ϊͬ�������ϵ�ҽ��
    Select Zl1_Ex_Isdoctorsamelevel(a.ҽ��id, a.ҽ������, ����ҽ��id_In, ����ҽ������_In)
    Into n_Count
    From �ٴ������¼ A
    Where ID = ��¼id_In;
    If n_Count = -1 Then
      v_Err_Msg := '����ҽ���ļ���С����ԭ����ҽ���ļ��𣬲����������ѡ������ҽ����';
      Raise Err_Item;
    End If;
  
    Insert Into �ٴ�����ͣ���¼
      (ID, ��¼id, ��ʼʱ��, ��ֹʱ��, ͣ��ԭ��, ����ҽ��id, ����ҽ������, ������, ����ʱ��, ������, ����ʱ��)
      Select �ٴ�����ͣ���¼_Id.Nextval, ��¼id_In, ��ʼʱ��_In, ��ֹʱ��_In, ͣ��ԭ��_In, ����ҽ��id_In, ����ҽ������_In, Nvl(a.ҽ������, ����Ա����_In),
             ����ʱ��_In, ����Ա����_In, ����ʱ��_In
      From �ٴ������¼ A
      Where ID = ��¼id_In;
  
    Update �ٴ������¼ Set ����ҽ��id = ����ҽ��id_In, ����ҽ������ = ����ҽ������_In Where ID = ��¼id_In;
  
    Insert Into ���˷�����Ϣ��¼
      (ID, ֪ͨ����, ��¼id, �Һ�id, ��Դid, ����, ����id, ��Ŀid, ҽ��id, ҽ������, ����id, �Ǽ���, �Ǽ�ʱ��, ֪ͨԭ��)
      Select ���˷�����Ϣ��¼_Id.Nextval, 2, ��¼id_In, �Һ�id, ��Դid, ����, ����id, ��Ŀid, ҽ��id, ҽ������, ����id, ����Ա����_In, ����ʱ��_In,
             'ҽ��' || ͣ��ԭ��_In || '��������'
      From (Select b.Id As �Һ�id, c.Id As ��Դid, c.����, c.����id, a.��Ŀid, a.ҽ��id, a.ҽ������, b.����id
             From �ٴ������¼ A, ���˹Һż�¼ B, �ٴ������Դ C
             Where a.Id = b.�����¼id And a.��Դid = c.Id And b.��¼״̬ = 1 And a.Id = ��¼id_In And
                   (b.��¼���� = 1 And b.����ʱ�� Between a.��ʼʱ�� And a.��ֹʱ�� Or b.��¼���� = 2 And b.ԤԼʱ�� Between a.��ʼʱ�� And a.��ֹʱ��));
  
    --��Ϣ����
    -- ��������(1-����,2-ȡ������),�����¼ID,�������
    Begin
      Select b.���� Into v_���� From �ٴ������¼ A, �ٴ������Դ B Where a.��Դid = b.Id And a.Id = ��¼id_In;
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
        Using 18, 1 || ',' || ��¼id_In || ',' || v_����;
    Exception
      When Others Then
        Null;
    End;
  
    --������ҽ��ͬ������ԤԼ�Һŵ�
    n_Updatedoctor := zl_GetSysParameter('������ҽ��ͬ������ԤԼ�Һŵ�', 1114);
    If Nvl(n_Updatedoctor, 0) = 1 Then
      For c_��¼ In (Select a.Id, b.No
                   From ���˷�����Ϣ��¼ A, ���˹Һż�¼ B
                   Where a.�Һ�id = b.Id And a.��¼id = ��¼id_In And a.֪ͨ���� = 2 And b.��¼���� In (1, 2) And b.��¼״̬ = 1) Loop
        Zl_���߷�������_����(c_��¼.Id, c_��¼.No, '������ҽ��ͬ������ԤԼ�Һŵ�', ����Ա����_In, ����Ա���_In);
      End Loop;
    End If;
  Else
    --���ݼ��
    Select ��ֹʱ��
    Into d_Cur
    From �ٴ������¼
    Where ID = ��¼id_In And ����ҽ������ Is Not Null And ͣ�￪ʼʱ�� Is Null;
    If d_Cur <= Sysdate Then
      v_Err_Msg := '��ֹʱ����С���˵�ǰʱ�䣬���ܽ���ȡ�����������';
      Raise Err_Item;
    End If;
    Select Count(1)
    Into n_Count
    From ���˷�����Ϣ��¼
    Where ��¼id = ��¼id_In And ֪ͨ���� = 2 And ������ Is Not Null;
    If n_Count <> 0 Then
      v_Err_Msg := '�ó����¼���ڲ��˷�����Ϣ��Ϣ��¼�����ѱ�����������ȡ�����������';
      Raise Err_Item;
    End If;
  
    Update �ٴ�����ͣ���¼
    Set ȡ���� = ����Ա����_In, ȡ��ʱ�� = ����ʱ��_In
    Where ��¼id = ��¼id_In And ����ҽ������ Is Not Null And ȡ���� Is Null;
  
    Update �ٴ������¼
    Set ����ҽ��id = Null, ����ҽ������ = Null
    Where ID = ��¼id_In And ����ҽ������ Is Not Null And ͣ�￪ʼʱ�� Is Null;
  
    Delete ���˷�����Ϣ��¼ Where ��¼id = ��¼id_In And ֪ͨ���� = 2 And ������ Is Null;
  
    --��Ϣ����
    -- ��������(1-����,2-ȡ������),�����¼ID,�������
    Begin
      Select b.���� Into v_���� From �ٴ������¼ A, �ٴ������Դ B Where a.��Դid = b.Id And a.Id = ��¼id_In;
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
        Using 18, 2 || ',' || ��¼id_In || ',' || v_����;
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
End Zl_�ٴ������¼_Replacedoctor;
/
Create Or Replace Procedure Zl_�ٴ�����ͣ��_Apply
(
  ��������_In Number,
  Id_In       �ٴ�����ͣ���¼.Id%Type,
  ��ʼʱ��_In �ٴ�����ͣ���¼.��ʼʱ��%Type,
  ��ֹʱ��_In �ٴ�����ͣ���¼.��ֹʱ��%Type,
  ͣ��ԭ��_In �ٴ�����ͣ���¼.ͣ��ԭ��%Type,
  ������_In   �ٴ�����ͣ���¼.������%Type,
  ����ʱ��_In �ٴ�����ͣ���¼.����ʱ��%Type
) As
  --���ܣ��˷������Լ�ȡ������
  --������
  --        ��������_In��0-���룬else-ȡ������
  --˵����
  n_Id    �ٴ�����ͣ���¼.Id%Type;
  n_Count Number;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  If ��������_In = 0 Then
    --����
    Begin
      Select 1
      Into n_Count
      From �ٴ�����ͣ���¼
      Where ��¼id Is Null And Not (��ʼʱ�� > ��ֹʱ��_In Or ��ֹʱ�� < ��ʼʱ��_In) And ������ = ������_In And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If Nvl(n_Count, 0) > 0 Then
      v_Error := 'ҽ�� ' || ������_In || ' �ڵ�ǰͣ��ʱ�䷶Χ���Ѵ���ͣ�ﰲ�ţ������ظ����룡';
      Raise Err_Custom;
    End If;
  
    If Nvl(Id_In, 0) = 0 Then
      Select �ٴ�����ͣ���¼_Id.Nextval Into n_Id From Dual;
    End If;
  
    Insert Into �ٴ�����ͣ���¼
      (ID, ��ʼʱ��, ��ֹʱ��, ͣ��ԭ��, ������, ����ʱ��)
    Values
      (n_Id, ��ʼʱ��_In, ��ֹʱ��_In, ͣ��ԭ��_In, ������_In, ����ʱ��_In);
  Else
    --ȡ������
    Select Count(1) Into n_Count From �ٴ�����ͣ���¼ Where ID = Id_In;
    If Nvl(n_Count, 0) = 0 Then
      v_Error := '�������ѱ�ȡ�����룬��ˢ�º�鿴...';
      Raise Err_Custom;
    End If;
  
    --���ͨ����������ȡ������
    Select Count(1) Into n_Count From �ٴ�����ͣ���¼ Where ID = Id_In And (������ Is Null Or ȡ���� Is Not Null);
    If Nvl(n_Count, 0) = 0 Then
      v_Error := '�������ѱ���ˣ�����ȡ�����롣';
      Raise Err_Custom;
    End If;
  
    Delete �ٴ�����ͣ���¼ Where ID = Id_In;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ�����ͣ��_Apply;
/
Create Or Replace Procedure Zl_�ٴ�����ͣ��_Audit
(
  ��������_In Number,
  Id_In       �ٴ�����ͣ���¼.Id%Type,
  ������_In   �ٴ�����ͣ���¼.������%Type,
  ����ʱ��_In �ٴ�����ͣ���¼.����ʱ��%Type
) As
  --���ܣ����ͣ�ﰲ��
  --������
  --       ״̬_In��1-��ˣ�2-ȡ�����
  n_Count Number;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  If Nvl(��������_In, 0) = 1 Then
    --���
    Select Count(1) Into n_Count From �ٴ�����ͣ���¼ Where ID = Id_In;
    If Nvl(n_Count, 0) = 0 Then
      v_Error := '�������ѱ�ȡ�����룬��ˢ�º�鿴...';
      Raise Err_Custom;
    End If;
  
    Select Count(1) Into n_Count From �ٴ�����ͣ���¼ Where ID = Id_In And (������ Is Null Or ȡ���� Is Not Null);
    If Nvl(n_Count, 0) = 0 Then
      v_Error := '�������ѱ���ˣ������ٴ���ˣ�';
      Raise Err_Custom;
    End If;
  
    Update �ٴ�����ͣ���¼
    Set ������ = ������_In, ����ʱ�� = ����ʱ��_In, ȡ���� = Null, ȡ��ʱ�� = Null
    Where ID = Id_In;
  
    --�Գ����¼����ͣ����
    For c_��¼ In (Select a.Id,
                        Case
                          When a.��ʼʱ�� < b.��ʼʱ�� Then
                           b.��ʼʱ��
                          Else
                           a.��ʼʱ��
                        End As ͣ�￪ʼʱ��,
                        Case
                          When a.��ֹʱ�� > b.��ֹʱ�� Then
                           b.��ֹʱ��
                          Else
                           a.��ֹʱ��
                        End As ͣ����ֹʱ��, b.ͣ��ԭ��, c.����
                 From �ٴ������¼ A, �ٴ�����ͣ���¼ B, �ٴ������Դ C
                 Where ((a.����ҽ������ Is Null And a.ҽ��id Is Not Null And a.ҽ������ = b.������) Or
                       (a.����ҽ������ Is Not Null And a.����ҽ��id Is Not Null And a.����ҽ������ = b.������)) And a.��Դid = c.Id And
                       b.Id = Id_In And Not (a.��ʼʱ�� > b.��ֹʱ�� Or a.��ֹʱ�� < b.��ʼʱ��)
                      --ֻ�����ѷ����˵�
                       And Exists (Select 1
                        From �ٴ����ﰲ�� C, �ٴ������ D
                        Where c.����id = d.Id And c.Id = a.����id And d.����ʱ�� Is Not Null)) Loop
    
      Update �ٴ������¼
      Set ͣ�￪ʼʱ�� = c_��¼.ͣ�￪ʼʱ��, ͣ����ֹʱ�� = c_��¼.ͣ����ֹʱ��, ͣ��ԭ�� = c_��¼.ͣ��ԭ��
      Where ID = c_��¼.Id;
    
      Insert Into ���˷�����Ϣ��¼
        (ID, ֪ͨ����, ��¼id, �Һ�id, ��Դid, ����, ����id, ��Ŀid, ҽ��id, ҽ������, ����id, �Ǽ���, �Ǽ�ʱ��)
        Select ���˷�����Ϣ��¼_Id.Nextval, 1, a.Id, b.Id, c.Id, c.����, c.����id, a.��Ŀid, a.ҽ��id, a.ҽ������, b.����id, ������_In, ����ʱ��_In
        From �ٴ������¼ A, ���˹Һż�¼ B, �ٴ������Դ C
        Where a.Id = b.�����¼id And a.��Դid = c.Id And b.��¼״̬ = 1 And a.Id = c_��¼.Id And
              (b.��¼���� = 1 And b.����ʱ�� Between a.ͣ�￪ʼʱ�� And a.ͣ����ֹʱ�� Or
              b.��¼���� = 2 And b.ԤԼʱ�� Between a.ͣ�￪ʼʱ�� And a.ͣ����ֹʱ��);
    
      --��Ϣ����
      -- ͣ������(1-ͣ��,2-ȡ��ͣ��),�����¼ID,ͣ�����
      Begin
        Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
          Using 17, 1 || ',' || c_��¼.Id || ',' || c_��¼.����;
      Exception
        When Others Then
          Null;
      End;
    End Loop;
  Else
    --ȡ�����
    Select Count(1) Into n_Count From �ٴ�����ͣ���¼ Where ID = Id_In And ������ Is Not Null And ȡ���� Is Null;
    If Nvl(n_Count, 0) = 0 Then
      v_Error := 'ԭ��˼�¼δ�ҵ�����ˢ�º�鿴...';
      Raise Err_Custom;
    End If;
  
    Select Count(1)
    Into n_Count
    From ���˷�����Ϣ��¼
    Where ��¼id In (Select a.Id
                   From �ٴ������¼ A, �ٴ�����ͣ���¼ B
                   Where Nvl(a.����ҽ������, a.ҽ������) = b.������ And b.Id = Id_In And
                         (a.��ʼʱ�� Between b.��ʼʱ�� And b.��ֹʱ�� Or a.��ֹʱ�� Between b.��ʼʱ�� And b.��ֹʱ��)) And ������ Is Not Null;
    If Nvl(n_Count, 0) <> 0 Then
      v_Error := '��ͣ�ﰲ�ŵĲ���ͣ����Ϣ�ѱ���������ȡ��������';
      Raise Err_Custom;
    End If;
  
    Update �ٴ�����ͣ���¼ Set ȡ���� = ������_In, ȡ��ʱ�� = ����ʱ��_In Where ID = Id_In;
  
    For c_��¼ In (Select a.Id, c.����
                 From �ٴ������¼ A, �ٴ�����ͣ���¼ B, �ٴ������Դ C
                 Where ((a.����ҽ������ Is Null And a.ҽ��id Is Not Null And a.ҽ������ = b.������) Or
                       (a.����ҽ������ Is Not Null And a.����ҽ��id Is Not Null And a.����ҽ������ = b.������)) And a.��Դid = c.Id And
                       b.Id = Id_In And (a.��ʼʱ�� Between b.��ʼʱ�� And b.��ֹʱ�� Or a.��ֹʱ�� Between b.��ʼʱ�� And b.��ֹʱ��) And
                       Exists (Select 1
                        From �ٴ����ﰲ�� C, �ٴ������ D
                        Where c.����id = d.Id And c.Id = a.����id And d.����ʱ�� Is Not Null)) Loop
    
      Update �ٴ������¼ Set ͣ�￪ʼʱ�� = Null, ͣ����ֹʱ�� = Null, ͣ��ԭ�� = Null Where ID = c_��¼.Id;
    
      Delete ���˷�����Ϣ��¼ Where ��¼id = c_��¼.Id And ֪ͨ���� = 1 And ������ Is Null;
    
      --��Ϣ����
      -- ͣ������(1-ͣ��,2-ȡ��ͣ��),�����¼ID,ͣ�����
      Begin
        Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
          Using 17, 2 || ',' || c_��¼.Id || ',' || c_��¼.����;
      Exception
        When Others Then
          Null;
      End;
    End Loop;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ�����ͣ��_Audit;
/

Create Or Replace Procedure Zl_�ٴ�����ԤԼ���Ʊ䶯
(
  �䶯����_In   �ٴ�����䶯��ϸ.�䶯����%Type,
  Id_In         �ٴ�����䶯��¼.Id%Type,
  ��¼id_In     �ٴ�����䶯��¼.��¼id%Type,
  ��ԤԼ����_In �ٴ�����䶯��¼.��ԤԼ����%Type := Null
) As
  --����:�޸�ԤԼ����ʱ�������ٴ�����䶯��¼/��ϸ
  --����:
  --     �䶯����_In  1-�䶯ǰ;2-�䶯��
  Err_Item Exception;
  v_Err_Msg Varchar2(100);
Begin
  If Nvl(�䶯����_In, 0) = 1 Then
    --�䶯ǰ
    Insert Into �ٴ�����䶯��¼
      (ID, ��¼id, �䶯����, ԭԤԼ����, ��ԤԼ����, ����Ա����, �Ǽ�ʱ��)
      Select Id_In, ��¼id_In, 4, ԤԼ����, Nvl(��ԤԼ����_In, ԤԼ����), Zl_Username, Sysdate
      From �ٴ������¼
      Where ID = ��¼id_In;
  
    Insert Into �ٴ�����䶯��ϸ
      (�䶯id, �䶯����, ����, ����, ���, ���Ʒ�ʽ, ����)
      Select Id_In, 1, ����, ����, ���, ���Ʒ�ʽ, ���� From �ٴ�����Һſ��Ƽ�¼ Where ��¼id = ��¼id_In;
  Else
    --�䶯��
    Insert Into �ٴ�����䶯��ϸ
      (�䶯id, �䶯����, ����, ����, ���, ���Ʒ�ʽ, ����)
      Select Id_In, 2, ����, ����, ���, ���Ʒ�ʽ, ���� From �ٴ�����Һſ��Ƽ�¼ Where ��¼id = ��¼id_In;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ�����ԤԼ���Ʊ䶯;
/

Create Or Replace Procedure Zl_�ٴ�������ſ��Ʊ䶯
(
  ��¼id_In     �ٴ�����䶯��¼.��¼id%Type,
  �޺���_In     �ٴ������¼.�޺���%Type,
  ��Լ��_In     �ٴ������¼.��Լ��%Type,
  ����Ա����_In �ٴ�����䶯��¼.����Ա����%Type := Null,
  �Ǽ�ʱ��_In   �ٴ�����䶯��¼.�Ǽ�ʱ��%Type := Null
) As
  --����:�޸��ٴ�������ſ���ʱ�������ٴ�����䶯��¼/��ϸ
  --����:
  n_ԭ�޺��� �ٴ������¼.�޺���%Type;
  n_ԭ��Լ�� �ٴ������¼.��Լ��%Type;

  v_����Ա���� �ٴ�����䶯��¼.����Ա����%Type := Null;
  d_�Ǽ�ʱ��   �ٴ�����䶯��¼.�Ǽ�ʱ��%Type := Null;

  Err_Item Exception;
  v_Err_Msg Varchar2(100);
Begin
  Begin
    Select �޺���, ��Լ�� Into n_ԭ�޺���, n_ԭ��Լ�� From �ٴ������¼ Where ID = ��¼id_In;
  Exception
    When Others Then
      v_Err_Msg := 'δ���ֳ����¼��';
      Raise Err_Item;
  End;

  --������Լ���޺���������Լ��Ϊ���ʾ��ֹԤԼ
  Update �ٴ������¼
  Set ��Լ�� = ��Լ��_In, �޺��� = �޺���_In, ԤԼ���� = Decode(Nvl(��Լ��_In, 0), 0, 1, ԤԼ����)
  Where ID = ��¼id_In;

  v_����Ա���� := Nvl(����Ա����_In, Zl_Username);
  d_�Ǽ�ʱ��   := Nvl(�Ǽ�ʱ��_In, Sysdate);
  If Nvl(n_ԭ�޺���, 0) <> Nvl(�޺���_In, 0) Then
    Insert Into �ٴ�����䶯��¼
      (ID, ��¼id, �䶯����, ԭ����, ������, ����Ա����, �Ǽ�ʱ��)
    Values
      (�ٴ�����䶯��¼_Id.Nextval, ��¼id_In, 1, n_ԭ�޺���, �޺���_In, v_����Ա����, d_�Ǽ�ʱ��);
  End If;
  If Nvl(n_ԭ��Լ��, 0) <> Nvl(��Լ��_In, 0) Then
    Insert Into �ٴ�����䶯��¼
      (ID, ��¼id, �䶯����, ԭ����, ������, ����Ա����, �Ǽ�ʱ��)
    Values
      (�ٴ�����䶯��¼_Id.Nextval, ��¼id_In, 2, n_ԭ��Լ��, ��Լ��_In, v_����Ա����, d_�Ǽ�ʱ��);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ�������ſ��Ʊ䶯;
/
Create Or Replace Procedure Zl_�ٴ�������ſ���_Update
(
  ��¼id_In   �ٴ������¼.Id%Type,
  ʱ��_In     Varchar2 := Null,
  ɾ�����_In Number := 0
) As
  --���ܣ������ٴ��������
  --������
  --     ʱ��_In:���,��ʼʱ��,��ֹʱ��,��������,ԤԼ��־|...
  --     ɾ�����_In:�Ƿ�ɾ���������ʱ��
  n_���     �ٴ�������ſ���.���%Type;
  d_��ʼʱ�� �ٴ�������ſ���.��ʼʱ��%Type;
  d_��ֹʱ�� �ٴ�������ſ���.��ֹʱ��%Type;
  n_����     �ٴ�������ſ���.����%Type;
  n_�Ƿ�ԤԼ �ٴ�������ſ���.�Ƿ�ԤԼ%Type;

  n_ʱ����� t_Numlist := t_Numlist();

  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin
  If Nvl(ɾ�����_In, 0) = 1 Then
    --ɾ������û�е��������ʱ��
    If ʱ��_In Is Not Null Then
      For c_ʱ�μ� In (Select Rownum As ���, Column_Value As ֵ From Table(f_Str2list(ʱ��_In, '|'))) Loop
        --���,��ʼʱ��,��ֹʱ��,��������,ԤԼ��־
        For c_ʱ�� In (Select Column_Value As ֵ From Table(f_Str2list(c_ʱ�μ�.ֵ)) Where Rownum = 1) Loop
        
          n_ʱ�����.Extend();
          n_ʱ�����(n_ʱ�����.Count) := To_Number(c_ʱ��.ֵ);
        End Loop;
      End Loop;
    End If;
  
    Delete �ٴ�������ſ��� Where ��¼id = ��¼id_In And ��� Not In (Select Column_Value From Table(n_ʱ�����));
    v_Err_Msg := n_ʱ�����.Count;
  End If;

  If ʱ��_In Is Not Null Then
    For c_ʱ�μ� In (Select Rownum As ���, Column_Value As ֵ From Table(f_Str2list(ʱ��_In, '|'))) Loop
      --���,��ʼʱ��,��ֹʱ��,��������,ԤԼ��־
      For c_ʱ�� In (Select Rownum As ���, Column_Value As ֵ From Table(f_Str2list(c_ʱ�μ�.ֵ))) Loop
        If c_ʱ��.��� = 1 Then
          n_��� := To_Number(c_ʱ��.ֵ);
        End If;
      
        If c_ʱ��.��� = 2 Then
          d_��ʼʱ�� := To_Date(c_ʱ��.ֵ, 'yyyy-mm-dd hh24:mi:ss');
        End If;
      
        If c_ʱ��.��� = 3 Then
          d_��ֹʱ�� := To_Date(c_ʱ��.ֵ, 'yyyy-mm-dd hh24:mi:ss');
        End If;
      
        If c_ʱ��.��� = 4 Then
          n_���� := To_Number(c_ʱ��.ֵ);
        End If;
      
        If c_ʱ��.��� = 5 Then
          n_�Ƿ�ԤԼ := To_Number(c_ʱ��.ֵ);
        End If;
      End Loop;
    
      If Nvl(n_���, 0) <> 0 Then
        Update �ٴ�������ſ���
        Set ��ʼʱ�� = d_��ʼʱ��, ��ֹʱ�� = d_��ֹʱ��, �Ƿ�ԤԼ = n_�Ƿ�ԤԼ, ���� = n_����
        Where ��¼id = ��¼id_In And ��� = n_���;
        If Sql%NotFound Then
          Insert Into �ٴ�������ſ���
            (��¼id, ���, ��ʼʱ��, ��ֹʱ��, �Ƿ�ԤԼ, ����)
          Values
            (��¼id_In, n_���, d_��ʼʱ��, d_��ֹʱ��, n_�Ƿ�ԤԼ, n_����);
        End If;
      End If;
    End Loop;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ�������ſ���_Update;
/


Create Or Replace Procedure Zl_���߷�������_����
(
  ��Ϣid_In     ���˷�����Ϣ��¼.Id%Type,
  No_In         ���˹Һż�¼.No%Type,
  ����˵��_In   ���˷�����Ϣ��¼.����˵��%Type,
  ����Ա����_In ���˷�����Ϣ��¼.������%Type,
  ����Ա���_In ���˹Һż�¼.����Ա���%Type
) As
  v_ԭִ����   ���˹Һż�¼.ִ����%Type;
  v_ִ����     ���˹Һż�¼.ִ����%Type;
  n_ԭִ����id �ٴ������¼.ҽ��id%Type;
  n_ִ����id   �ٴ������¼.����ҽ��id%Type;
  d_��������   �ٴ������¼.��������%Type;
  n_�������   ���˹Һż�¼.����%Type;
  n_��Ŀid     �ٴ������¼.��Ŀid%Type;
  n_����id     �ٴ������Դ.����id%Type;
  n_�Һ�״̬   Number(3); --1=�Һ�,2=ԤԼδ�տ�,3=ԤԼ���տ�
  v_����       �ٴ������Դ.����%Type;
  n_�䶯id     ����䶯��¼.Id%Type;
  v_Err_Msg    Varchar2(500);
  Err_Item Exception;
Begin
  --��ȡ����ҽ��
  Select b.����ҽ������, b.ҽ������, b.ҽ��id, b.����ҽ��id, ��������, b.��Ŀid, c.����id, c.����
  Into v_ִ����, v_ԭִ����, n_ԭִ����id, n_ִ����id, d_��������, n_��Ŀid, n_����id, v_����
  From ���˷�����Ϣ��¼ A, �ٴ������¼ B, �ٴ������Դ C
  Where a.Id = ��Ϣid_In And a.��¼id = b.Id And b.��Դid = c.Id;

  Select Decode(Nvl(ԤԼ, 0), 0, 1, Decode(����ʱ��, Null, 2, 3)), ����
  Into n_�Һ�״̬, n_�������
  From ���˹Һż�¼
  Where NO = No_In And ��¼״̬ = 1;

  --����䶯��¼
  Select ����䶯��¼_Id.Nextval Into n_�䶯id From Dual;
  b_Message.Zlhis_Regist_005(No_In, n_�䶯id, 1);
  Zl_����䶯��¼_Insert(No_In, 5, '���߷�����������', ����Ա����_In, ����Ա���_In, v_����, n_����id, n_��Ŀid, n_ִ����id, v_ִ����, Null, n_�������,
                   Sysdate, n_�䶯id);

  --���²��˹Һż�¼
  Update ���˹Һż�¼ Set ִ���� = v_ִ���� Where NO = No_In And ��¼״̬ = 1;
  --����������ü�¼
  Update ������ü�¼ Set ִ���� = v_ִ���� Where NO = No_In And ��¼���� = 4;
  --���»��߷����¼
  Update ���˷�����Ϣ��¼ Set ������ = ����Ա����_In, ����ʱ�� = Sysdate, ����˵�� = ����˵��_In Where ID = ��Ϣid_In;
  --���²��˹ҺŻ���
  If n_�Һ�״̬ = 1 Then
    Update ���˹ҺŻ���
    Set �ѹ��� = Nvl(�ѹ���, 0) - 1
    Where ���� = Trunc(d_��������) And Nvl(ҽ������, '-') = Nvl(v_ԭִ����, '-') And Nvl(ҽ��id, 0) = Nvl(n_ԭִ����id, 0) And
          Nvl(����id, 0) = Nvl(n_����id, 0) And Nvl(��Ŀid, 0) = Nvl(n_��Ŀid, 0) And (���� = v_���� Or ���� Is Null);
    Update ���˹ҺŻ���
    Set �ѹ��� = Nvl(�ѹ���, 0) + 1
    Where ���� = Trunc(d_��������) And Nvl(ҽ������, '-') = Nvl(v_ִ����, '-') And Nvl(ҽ��id, 0) = Nvl(n_ִ����id, 0) And
          Nvl(����id, 0) = Nvl(n_����id, 0) And Nvl(��Ŀid, 0) = Nvl(n_��Ŀid, 0) And (���� = v_���� Or ���� Is Null);
    If Sql%RowCount = 0 Then
      Insert Into ���˹ҺŻ���
        (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, ��Լ��, �����ѽ���)
      Values
        (Trunc(d_��������), n_����id, n_��Ŀid, v_ִ����, Decode(n_ִ����id, 0, Null, n_ִ����id), v_����, 1, 0, 0);
    End If;
  End If;
  If n_�Һ�״̬ = 2 Then
    Update ���˹ҺŻ���
    Set ��Լ�� = Nvl(��Լ��, 0) - 1
    Where ���� = Trunc(d_��������) And Nvl(ҽ������, '-') = Nvl(v_ԭִ����, '-') And Nvl(ҽ��id, 0) = Nvl(n_ԭִ����id, 0) And
          Nvl(����id, 0) = Nvl(n_����id, 0) And Nvl(��Ŀid, 0) = Nvl(n_��Ŀid, 0) And (���� = v_���� Or ���� Is Null);
    Update ���˹ҺŻ���
    Set ��Լ�� = Nvl(��Լ��, 0) + 1
    Where ���� = Trunc(d_��������) And Nvl(ҽ������, '-') = Nvl(v_ִ����, '-') And Nvl(ҽ��id, 0) = Nvl(n_ִ����id, 0) And
          Nvl(����id, 0) = Nvl(n_����id, 0) And Nvl(��Ŀid, 0) = Nvl(n_��Ŀid, 0) And (���� = v_���� Or ���� Is Null);
    If Sql%RowCount = 0 Then
      Insert Into ���˹ҺŻ���
        (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, ��Լ��, �����ѽ���)
      Values
        (Trunc(d_��������), n_����id, n_��Ŀid, v_ִ����, Decode(n_ִ����id, 0, Null, n_ִ����id), v_����, 0, 1, 0);
    End If;
  End If;
  If n_�Һ�״̬ = 3 Then
    Update ���˹ҺŻ���
    Set ��Լ�� = Nvl(��Լ��, 0) - 1, �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - 1
    Where ���� = Trunc(d_��������) And Nvl(ҽ������, '-') = Nvl(v_ԭִ����, '-') And Nvl(ҽ��id, 0) = Nvl(n_ԭִ����id, 0) And
          Nvl(����id, 0) = Nvl(n_����id, 0) And Nvl(��Ŀid, 0) = Nvl(n_��Ŀid, 0) And (���� = v_���� Or ���� Is Null);
    Update ���˹ҺŻ���
    Set ��Լ�� = Nvl(��Լ��, 0) + 1, �ѹ��� = Nvl(�ѹ���, 0) + 1, �����ѽ��� = Nvl(�����ѽ���, 0) + 1
    Where ���� = Trunc(d_��������) And Nvl(ҽ������, '-') = Nvl(v_ִ����, '-') And Nvl(ҽ��id, 0) = Nvl(n_ִ����id, 0) And
          Nvl(����id, 0) = Nvl(n_����id, 0) And Nvl(��Ŀid, 0) = Nvl(n_��Ŀid, 0) And (���� = v_���� Or ���� Is Null);
    If Sql%RowCount = 0 Then
      Insert Into ���˹ҺŻ���
        (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, ��Լ��, �����ѽ���)
      Values
        (Trunc(d_��������), n_����id, n_��Ŀid, v_ִ����, Decode(n_ִ����id, 0, Null, n_ִ����id), v_����, 1, 1, 1);
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���߷�������_����;
/


Create Or Replace Procedure Zl_���߷�������_����
(
  ��Ϣid_In     ���˷�����Ϣ��¼.Id%Type,
  No_In         ���˹Һż�¼.No%Type,
  �������_In   ���˹Һż�¼.����%Type,
  ����ʱ��_In   ���˹Һż�¼.ԤԼʱ��%Type,
  ����id_In     �ٴ������¼.Id%Type,
  ����˵��_In   ���˷�����Ϣ��¼.����˵��%Type,
  ����Ա����_In ���˷�����Ϣ��¼.������%Type,
  ����Ա���_In ���˹Һż�¼.����Ա���%Type
  
) As
  Cursor c_Registinfo Is
    Select a.����ʱ��, a.�Ǽ�ʱ��, c.����ʱ��, a.�շ�ϸĿid As ��Ŀid, c.ִ�в���id As ����id, c.ִ���� As ҽ������, d.Id As ҽ��id, c.�ű� As ����, c.����
    From ������ü�¼ A, �ҺŰ��� B, ���˹Һż�¼ C, ��Ա�� D
    Where a.��¼���� = 4 And c.No = a.No And c.ִ���� = d.����(+) And a.No = No_In And Nvl(a.���㵥λ, '�ű�') = c.�ű� And Rownum < 2;
  r_Registrow   c_Registinfo%RowType;
  v_�ű�        ���˹Һż�¼.�ű�%Type;
  n_ִ�в���id  ���˹Һż�¼.ִ�в���id%Type;
  n_��Ŀid      �ٴ������¼.��Ŀid%Type;
  v_ִ����      ���˹Һż�¼.ִ����%Type;
  n_ִ����id    ��Ա��.Id%Type;
  n_������      Number(2);
  n_�շ�        Number(2);
  n_Exists      Number(3);
  v_Temp        Varchar2(500);
  v_�շ���Ŀids Varchar2(500);
  v_Err_Msg     Varchar2(500);
  n_������id    �շ���ĿĿ¼.Id%Type;
  n_���        ������ü�¼.���%Type;
  n_ԤԼ        ���˹Һż�¼.ԤԼ%Type;
  n_ʵ�ս��    ������ü�¼.ʵ�ս��%Type;
  n_Ӧ�ս��    ������ü�¼.Ӧ�ս��%Type;
  v_�ѱ�        ������ü�¼.�ѱ�%Type;
  n_����id      ���˹Һż�¼.����id%Type;
  n_����        �ٴ������¼.�ѹ���%Type;
  n_����        �ٴ������¼.�޺���%Type;
  n_ԭ����      ���˹Һż�¼.����%Type;
  n_ԭ��¼id    �ٴ������¼.Id%Type;
  n_ԭ�Һ�״̬  �ٴ�������ſ���.�Һ�״̬%Type;
  v_ԭ����Ա    �ٴ�������ſ���.����Ա����%Type;
  v_ԭ��ע      �ٴ�������ſ���.��ע%Type;
  n_��ſ���    �ٴ������¼.�Ƿ���ſ���%Type;
  n_ԤԼ˳���  �ٴ�������ſ���.ԤԼ˳���%Type;
  n_ʵ�����    �ٴ�������ſ���.���%Type;
  n_�䶯id      ����䶯��¼.Id%Type;
  Err_Item Exception;
Begin
  Begin
    Select 1, ����id Into n_Exists, n_����id From ���˹Һż�¼ Where NO = No_In And ��¼״̬ = 1;
  Exception
    When Others Then
      v_Err_Msg := '���ݺ�Ϊ' || No_In || '��ԤԼ��¼������,�޷�����!';
      Raise Err_Item;
  End;
  Begin
    Select �ѱ� Into v_�ѱ� From ������ü�¼ Where NO = No_In And ��¼���� = 4;
  Exception
    When Others Then
      Begin
        Select �ѱ� Into v_�ѱ� From ������Ϣ Where ����id = n_����id;
      Exception
        When Others Then
          v_�ѱ� := Null;
      End;
  End;

  Select b.����, b.����id, Nvl(c.����, a.ҽ������), a.��Ŀid, c.Id, Nvl(a.�Ƿ���ſ���, 0)
  Into v_�ű�, n_ִ�в���id, v_ִ����, n_��Ŀid, n_ִ����id, n_��ſ���
  From �ٴ������¼ A, �ٴ������Դ B, ��Ա�� C
  Where a.Id = ����id_In And a.��Դid = b.Id And a.ҽ��id = c.Id(+);

  Select Max(1) Into n_�շ� From ������ü�¼ Where NO = No_In And ��¼���� = 4 And ���ʽ�� Is Not Null;
  Select Max(1)
  Into n_������
  From ������ü�¼ A, �շ��ض���Ŀ B
  Where a.No = No_In And a.��¼���� = 4 And a.�շ�ϸĿid = b.�շ�ϸĿid And b.�ض���Ŀ = '������';

  --����䶯��¼
  Select ����䶯��¼_Id.Nextval Into n_�䶯id From Dual;
  b_Message.Zlhis_Regist_005(No_In, n_�䶯id, 2);
  Zl_����䶯��¼_Insert(No_In, 4, '���߷������Ļ���', ����Ա����_In, ����Ա���_In, v_�ű�, n_ִ�в���id, n_��Ŀid, n_ִ����id, v_ִ����, Null, �������_In,
                   ����ʱ��_In, n_�䶯id);

  --���»��߷����¼
  Update ���˷�����Ϣ��¼ Set ������ = ����Ա����_In, ����ʱ�� = Sysdate, ����˵�� = ����˵��_In Where ID = ��Ϣid_In;

  --���²��˹ҺŻ���(����)
  Select ԤԼ, �����¼id Into n_ԤԼ, n_ԭ��¼id From ���˹Һż�¼ Where NO = No_In And ��¼״̬ = 1;

  --��黻���¼�Ƿ������㹻
  If n_ԤԼ = 0 Then
    Select �ѹ���, �޺��� Into n_����, n_���� From �ٴ������¼ Where ID = ����id_In;
    If Not n_���� Is Null Then
      If Nvl(n_����, 0) >= n_���� Then
        v_Err_Msg := 'Ҫ����ļ�¼�Ѿ����������������' || n_���� || ',�޷�����!';
        Raise Err_Item;
      End If;
    End If;
  Else
    If n_�շ� = 1 Then
      Select �ѹ���, �޺��� Into n_����, n_���� From �ٴ������¼ Where ID = ����id_In;
      If Not n_���� Is Null Then
        If Nvl(n_����, 0) >= n_���� Then
          v_Err_Msg := 'Ҫ����ļ�¼�Ѿ����������������' || n_���� || ',�޷�����!';
          Raise Err_Item;
        End If;
      End If;
    Else
      Select ��Լ��, ��Լ�� Into n_����, n_���� From �ٴ������¼ Where ID = ����id_In;
      If Not n_���� Is Null Then
        If Nvl(n_����, 0) >= n_���� Then
          v_Err_Msg := 'Ҫ����ļ�¼�Ѿ����������������' || n_���� || ',�޷�����!';
          Raise Err_Item;
        End If;
      End If;
    End If;
  End If;

  If n_ԤԼ = 0 Then
    Open c_Registinfo;
    Fetch c_Registinfo
      Into r_Registrow;
  
    n_ԭ���� := r_Registrow.����;
    Update �ٴ������¼ Set �ѹ��� = Nvl(�ѹ���, 0) - 1 Where ID = n_ԭ��¼id;
    Update ���˹ҺŻ���
    Set �ѹ��� = Nvl(�ѹ���, 0) - 1
    Where ���� = Trunc(r_Registrow.����ʱ��) And ҽ������ = r_Registrow.ҽ������ And ҽ��id = r_Registrow.ҽ��id And
          ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And (���� = r_Registrow.���� Or ���� Is Null);
  
    If Sql%RowCount = 0 Then
      Insert Into ���˹ҺŻ���
        (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���)
      Values
        (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
         Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -1);
    End If;
    Close c_Registinfo;
  Else
    If n_�շ� = 1 Then
      Open c_Registinfo;
      Fetch c_Registinfo
        Into r_Registrow;
    
      n_ԭ���� := r_Registrow.����;
    
      Update �ٴ������¼
      Set ��Լ�� = Nvl(��Լ��, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - 1, �ѹ��� = Nvl(�ѹ���, 0) - 1
      Where ID = n_ԭ��¼id;
    
      Update ���˹ҺŻ���
      Set ��Լ�� = Nvl(��Լ��, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - 1, �ѹ��� = Nvl(�ѹ���, 0) - 1
      Where ���� = Trunc(r_Registrow.����ʱ��) And ҽ������ = r_Registrow.ҽ������ And ҽ��id = r_Registrow.ҽ��id And
            ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And (���� = r_Registrow.���� Or ���� Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into ���˹ҺŻ���
          (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, ��Լ��, �����ѽ���, �ѹ���)
        Values
          (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
           Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -1, -1, -1);
      End If;
      Close c_Registinfo;
    Else
      Open c_Registinfo;
      Fetch c_Registinfo
        Into r_Registrow;
    
      n_ԭ���� := r_Registrow.����;
    
      Update �ٴ������¼ Set ��Լ�� = Nvl(��Լ��, 0) - 1 Where ID = n_ԭ��¼id;
      Update ���˹ҺŻ���
      Set ��Լ�� = Nvl(��Լ��, 0) - 1
      Where ���� = Trunc(r_Registrow.����ʱ��) And ҽ������ = r_Registrow.ҽ������ And ҽ��id = r_Registrow.ҽ��id And
            ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And (���� = r_Registrow.���� Or ���� Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into ���˹ҺŻ���
          (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, ��Լ��)
        Values
          (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
           Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -1);
      End If;
      Close c_Registinfo;
    End If;
  End If;

  If n_��ſ��� = 0 And Nvl(�������_In, 0) <> 0 Then
    Select Max(ԤԼ˳���)
    Into n_ԤԼ˳���
    From �ٴ�������ſ���
    Where ��¼id = ����id_In And ��� = �������_In And ԤԼ˳��� Is Not Null;
    If n_ԤԼ˳��� Is Null Then
      n_ԤԼ˳��� := 1;
    Else
      n_ԤԼ˳��� := n_ԤԼ˳��� + 1;
    End If;
    n_ʵ����� := To_Number(�������_In || n_ԤԼ˳���);
  Else
    n_ʵ����� := �������_In;
  End If;
  --���²��˹Һż�¼
  Update ���˹Һż�¼
  Set �ű� = v_�ű�, ִ�в���id = n_ִ�в���id, ִ���� = v_ִ����, ���� = n_ʵ�����, ����ʱ�� = ����ʱ��_In, ԤԼʱ�� = ����ʱ��_In, �����¼id = ����id_In
  Where NO = No_In And ��¼״̬ = 1;

  --����������ü�¼
  If n_������ = 1 Then
    Select �շ�ϸĿid Into n_������id From �շ��ض���Ŀ Where �ض���Ŀ = '������';
    v_�շ���Ŀids := n_��Ŀid || ',' || n_������id;
  Else
    v_�շ���Ŀids := n_��Ŀid;
  End If;
  Update ������ü�¼
  Set ���˿���id = n_ִ�в���id, ���㵥λ = v_�ű�, ��ҩ���� = n_ʵ�����, ִ�в���id = n_ִ�в���id, ִ���� = v_ִ����, ����ʱ�� = ����ʱ��_In
  Where NO = No_In And ��¼���� = 4;
  n_��� := 1;
  For c_Item In (Select 1 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, 1 As ����, c.Id As ������Ŀid,
                        c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, Null As ��������
                 From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C
                 Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = n_��Ŀid And Sysdate Between b.ִ������ And
                       Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))
                 Union All
                 Select 2 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, 1 As ����, c.Id As ������Ŀid,
                        c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, Null As ��������
                 From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C
                 Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = n_������id And Sysdate Between b.ִ������ And
                       Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))
                 Union All
                 Select 3 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, d.�������� As ����,
                        c.Id As ������Ŀid, c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, 1 As ��������
                 From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, �շѴ�����Ŀ D
                 Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = d.����id And
                       d.����id In (Select Column_Value From Table(f_Str2list(v_�շ���Ŀids))) And Sysdate Between b.ִ������ And
                       Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))
                 Order By ����, ��Ŀ����, �������) Loop
    n_Ӧ�ս�� := c_Item.���� * c_Item.����;
  
    If Nvl(c_Item.���ηѱ�, 0) <> 1 Then
      --����:
      v_Temp     := Zl_Actualmoney(v_�ѱ�, c_Item.��Ŀid, c_Item.������Ŀid, n_Ӧ�ս��);
      v_Temp     := Substr(v_Temp, Instr(v_Temp, ':') + 1);
      n_ʵ�ս�� := Zl_To_Number(v_Temp);
    Else
      n_ʵ�ս�� := n_Ӧ�ս��;
    End If;
  
    If n_�շ� = 1 Then
      Update ������ü�¼
      Set �շ���� = c_Item.���, �շ�ϸĿid = c_Item.��Ŀid, ������Ŀid = c_Item.������Ŀid, �վݷ�Ŀ = c_Item.�վݷ�Ŀ, ���� = c_Item.����,
          ��׼���� = c_Item.����, Ӧ�ս�� = n_Ӧ�ս��, ʵ�ս�� = n_ʵ�ս��, ���ʽ�� = n_ʵ�ս��
      Where ��� = n_��� And ��¼���� = 4 And NO = No_In;
      If Sql%RowCount = 0 Then
        Insert Into ������ü�¼
          (ID, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, NO, ʵ��Ʊ��, �����־, �Ӱ��־, ���ӱ�־, ��ҩ����, ����id, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, ���˿���id,
           �շ����, ���㵥λ, �շ�ϸĿid, ������Ŀid, �վݷ�Ŀ, ����, ����, ��׼����, Ӧ�ս��, ʵ�ս��, ���ʽ��, ����id, ���ʷ���, ��������id, ������, ������, ִ�в���id, ִ����,
           ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ���մ���id, ������Ŀ��, ���ձ���, ͳ����, ժҪ, ����, �ɿ���id)
          Select ���˷��ü�¼_Id.Nextval, 4, ��¼״̬, n_���, Null, Null, NO, ʵ��Ʊ��, �����־, Null, Null, ��ҩ����, ����id, ��ʶ��, ���ʽ, ����, �Ա�,
                 ����, �ѱ�, ���˿���id, c_Item.���, ���㵥λ, c_Item.��Ŀid, c_Item.������Ŀid, c_Item.�վݷ�Ŀ, 1, c_Item.����, c_Item.����,
                 n_Ӧ�ս��, n_ʵ�ս��, n_ʵ�ս��, ����id, 0, ��������id, ������, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ���մ���id, ������Ŀ��,
                 ���ձ���, ͳ����, ժҪ, ����, �ɿ���id
          From ������ü�¼
          Where ��¼���� = 4 And NO = No_In And ��� = 1;
      End If;
    Else
      Update ������ü�¼
      Set �շ�ϸĿid = c_Item.��Ŀid, ������Ŀid = c_Item.������Ŀid, �վݷ�Ŀ = c_Item.�վݷ�Ŀ, ��׼���� = c_Item.����, Ӧ�ս�� = c_Item.����,
          ʵ�ս�� = c_Item.����
      Where ��� = n_��� And ��¼���� = 4 And NO = No_In;
      If Sql%RowCount = 0 Then
        Insert Into ������ü�¼
          (ID, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, NO, ʵ��Ʊ��, �����־, �Ӱ��־, ���ӱ�־, ��ҩ����, ����id, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, ���˿���id,
           �շ����, ���㵥λ, �շ�ϸĿid, ������Ŀid, �վݷ�Ŀ, ����, ����, ��׼����, Ӧ�ս��, ʵ�ս��, ���ʽ��, ����id, ���ʷ���, ��������id, ������, ������, ִ�в���id, ִ����,
           ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ���մ���id, ������Ŀ��, ���ձ���, ͳ����, ժҪ, ����, �ɿ���id)
          Select ���˷��ü�¼_Id.Nextval, 4, ��¼״̬, n_���, Null, Null, NO, ʵ��Ʊ��, �����־, Null, Null, ��ҩ����, ����id, ��ʶ��, ���ʽ, ����, �Ա�,
                 ����, �ѱ�, ���˿���id, c_Item.���, ���㵥λ, c_Item.��Ŀid, c_Item.������Ŀid, c_Item.�վݷ�Ŀ, 1, c_Item.����, c_Item.����,
                 n_Ӧ�ս��, n_ʵ�ս��, Null, ����id, 0, ��������id, ������, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ���մ���id, ������Ŀ��,
                 ���ձ���, ͳ����, ժҪ, ����, �ɿ���id
          From ������ü�¼
          Where ��¼���� = 4 And NO = No_In And ��� = 1;
      End If;
    End If;
    n_��� := n_��� + 1;
  End Loop;

  --���²��˹ҺŻ���(����)
  If n_ԤԼ = 0 Then
    Open c_Registinfo;
    Fetch c_Registinfo
      Into r_Registrow;
  
    Update �ٴ������¼ Set �ѹ��� = Nvl(�ѹ���, 0) + 1 Where ID = ����id_In;
  
    Update ���˹ҺŻ���
    Set �ѹ��� = Nvl(�ѹ���, 0) + 1
    Where ���� = Trunc(r_Registrow.����ʱ��) And ҽ������ = r_Registrow.ҽ������ And ҽ��id = r_Registrow.ҽ��id And
          ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And (���� = r_Registrow.���� Or ���� Is Null);
  
    If Sql%RowCount = 0 Then
      Insert Into ���˹ҺŻ���
        (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���)
      Values
        (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
         Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, 1);
    End If;
    Close c_Registinfo;
  Else
    If n_�շ� = 1 Then
      Open c_Registinfo;
      Fetch c_Registinfo
        Into r_Registrow;
    
      Update �ٴ������¼
      Set ��Լ�� = Nvl(��Լ��, 0) + 1, �����ѽ��� = Nvl(�����ѽ���, 0) + 1, �ѹ��� = Nvl(�ѹ���, 0) + 1
      Where ID = ����id_In;
    
      Update ���˹ҺŻ���
      Set ��Լ�� = Nvl(��Լ��, 0) + 1, �����ѽ��� = Nvl(�����ѽ���, 0) + 1, �ѹ��� = Nvl(�ѹ���, 0) + 1
      Where ���� = Trunc(r_Registrow.����ʱ��) And ҽ������ = r_Registrow.ҽ������ And ҽ��id = r_Registrow.ҽ��id And
            ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And (���� = r_Registrow.���� Or ���� Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into ���˹ҺŻ���
          (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, ��Լ��, �����ѽ���, �ѹ���)
        Values
          (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
           Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, 1, 1, 1);
      End If;
      Close c_Registinfo;
    Else
      Open c_Registinfo;
      Fetch c_Registinfo
        Into r_Registrow;
    
      Update �ٴ������¼ Set ��Լ�� = Nvl(��Լ��, 0) + 1 Where ID = ����id_In;
    
      Update ���˹ҺŻ���
      Set ��Լ�� = Nvl(��Լ��, 0) + 1
      Where ���� = Trunc(r_Registrow.����ʱ��) And ҽ������ = r_Registrow.ҽ������ And ҽ��id = r_Registrow.ҽ��id And
            ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And (���� = r_Registrow.���� Or ���� Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into ���˹ҺŻ���
          (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, ��Լ��)
        Values
          (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
           Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, 1);
      End If;
      Close c_Registinfo;
    End If;
  End If;
  --�������
  Begin
    Select �Һ�״̬, ����Ա����, ��ע
    Into n_ԭ�Һ�״̬, v_ԭ����Ա, v_ԭ��ע
    From �ٴ�������ſ���
    Where ��¼id = n_ԭ��¼id And (��� = n_ԭ���� Or ��ע = n_ԭ����);
  
    If n_��ſ��� = 0 And Nvl(�������_In, 0) <> 0 Then
      Insert Into �ٴ�������ſ���
        (��¼id, ���, ԤԼ˳���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����Ա����, ��ע)
        Select ��¼id, ���, n_ԤԼ˳���, ��ʼʱ��, ��ֹʱ��, 1, �Ƿ�ԤԼ, n_ԭ�Һ�״̬, v_ԭ����Ա, n_ʵ�����
        From �ٴ�������ſ���
        Where ��¼id = ����id_In And ��� = �������_In And ԤԼ˳��� Is Null;
    Else
      Update �ٴ�������ſ���
      Set �Һ�״̬ = n_ԭ�Һ�״̬, ����Ա���� = v_ԭ����Ա, ��ע = v_ԭ��ע
      Where ��¼id = ����id_In And ��� = �������_In;
    End If;
  
    Update �ٴ�������ſ���
    Set �Һ�״̬ = Null, ����Ա���� = Null, ��ע = Null
    Where ��¼id = n_ԭ��¼id And (��� = n_ԭ���� Or ��ע = n_ԭ����);
  Exception
    When Others Then
      Null;
  End;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���߷�������_����;
/


Create Or Replace Procedure Zl_���߷�������_����
(
  ��Ϣid_In     ���˷�����Ϣ��¼.Id%Type,
  ����˵��_In   ���˷�����Ϣ��¼.����˵��%Type,
  ����Ա����_In ���˷�����Ϣ��¼.������%Type,
  ����Ա���_In ���˹Һż�¼.����Ա���%Type,
  �Һ�id_In     ���˹Һż�¼.Id%Type := Null,
  ������ʽ_In   Number := 0
) As
  --������ʽ_IN:0-��������,1-ȡ��ԤԼ�Ǽ�
  v_Err_Msg Varchar2(200);
  Err_Item Exception;
Begin
  --���»��߷����¼
  If ������ʽ_In = 0 Then
    If �Һ�id_In Is Null Then
      Update ���˷�����Ϣ��¼
      Set ������ = ����Ա����_In, ����ʱ�� = Sysdate, ����˵�� = ����˵��_In
      Where ID = ��Ϣid_In;
    Else
      Update ���˷�����Ϣ��¼
      Set ������ = ����Ա����_In, ����ʱ�� = Sysdate, ����˵�� = ����˵��_In, �Һ�id = �Һ�id_In
      Where ID = ��Ϣid_In;
    End If;
  Else
    Delete From ���˷�����Ϣ��¼ Where ID = ��Ϣid_In;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���߷�������_����;
/


Create Or Replace Procedure Zl_���˹ҺŻ���_Update
(
  ҽ������_In   �ҺŰ���.ҽ������%Type,
  ҽ��id_In     �ҺŰ���.ҽ��id%Type,
  �շ�ϸĿid_In ������ü�¼.�շ�ϸĿid%Type,
  ִ�в���id_In ������ü�¼.ִ�в���id%Type,
  ����ʱ��_In   ������ü�¼.����ʱ��%Type,
  ԤԼ��־_In   Number := 0, --�Ƿ�ΪԤԼ����:0-��ԤԼ�Һ�; 1-ԤԼ�Һ�,2-ԤԼ����;3-�շ�ԤԼ
  ����_In       �ҺŰ���.����%Type := Null,
  ��������_In   Number := 0, --�Ƿ�ӿڵ���
  �����¼id_In �ٴ������¼.Id%Type := Null
) As
  --����ʱ��_In:ԤԼʱ,ΪԤԼʱ��;����Ϊ�Ǽ�ʱ��
  v_Date    Date;
  n_ԤԼ��  ���˹ҺŻ���.��Լ��%Type;
  n_ʱ��    Number := 0;
  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  n_����ģʽ Number := 0;
  n_��Դid   �ٴ������¼.��Դid%Type;
Begin
  If �����¼id_In Is Null Then
    Begin
      Select 1
      Into n_ʱ��
      From Dual
      Where Exists (Select 1
             From �ҺŰ���ʱ�� A, �ҺŰ��� B
             Where a.����id = b.Id And b.���� = ����_In And Rownum <= 1
             Union All
             Select 1
             From �Һżƻ�ʱ�� C, �ҺŰ��żƻ� D ��
             Where c.�ƻ�id = d.Id And d.���� = ����_In And d.��Чʱ�� > Sysdate And Rownum <= 1);
    Exception
      When Others Then
        n_ʱ�� := 0;
    End;
    n_����ģʽ := Nvl(zl_GetSysParameter('ԤԼ����ģʽ', 1111), 0);
    --��ʱ�εĺű�ֻ�ܵ������
    If n_ʱ�� = 1 And Nvl(ԤԼ��־_In, 0) = 2 And ��������_In = 0 And n_����ģʽ = 0 Then
      If Trunc(����ʱ��_In) <> Trunc(Sysdate) Then
        v_Err_Msg := '��ʱ�ε�ԤԼ�Һŵ�ֻ�ܵ�����գ�';
        Raise Err_Item;
      End If;
    End If;
    If Nvl(ԤԼ��־_In, 0) <> 2 Or ��������_In = 1 Then
      v_Date := Trunc(����ʱ��_In);
    Else
      If n_����ģʽ = 0 Then
        v_Date := Trunc(Sysdate);
      Else
        v_Date := Trunc(����ʱ��_In);
      End If;
    End If;
  
    n_ԤԼ�� := 0;
    If Nvl(ԤԼ��־_In, 0) <> 1 Then
      --��ԤԼ�Һ�;��ԤԼ����
      If Nvl(ԤԼ��־_In, 0) = 2 And v_Date <> Trunc(����ʱ��_In) Then
        --1.��ȥԤԼ���ڵ�ԤԼ��;
        --2-���ϵ�ǰԤԼ���ڵĹҺ���;
        Update ���˹ҺŻ���
        Set ��Լ�� = Nvl(��Լ��, 0) - 1
        Where ���� = Trunc(����ʱ��_In) And Nvl(����id, 0) = ִ�в���id_In And Nvl(��Ŀid, 0) = �շ�ϸĿid_In And
              (���� = ����_In Or ���� Is Null)
        Returning ��Լ�� Into n_ԤԼ��;
      
        If n_ԤԼ�� < 0 Then
          Update ���˹ҺŻ���
          Set ��Լ�� = 0
          Where ���� = Trunc(����ʱ��_In) And Nvl(����id, 0) = ִ�в���id_In And Nvl(��Ŀid, 0) = �շ�ϸĿid_In And
                (���� = ����_In Or ���� Is Null)
          Returning ��Լ�� Into n_ԤԼ��;
        End If;
        n_ԤԼ�� := 1;
      Elsif Nvl(ԤԼ��־_In, 0) = 3 Then
        n_ԤԼ�� := 1;
      End If;
      Update ���˹ҺŻ���
      Set �ѹ��� = Nvl(�ѹ���, 0) + 1, �����ѽ��� = Nvl(�����ѽ���, 0) + Decode(ԤԼ��־_In, 0, 0, 1), ��Լ�� = Nvl(��Լ��, 0) + Nvl(n_ԤԼ��, 0)
      Where ���� = Decode(ԤԼ��־_In, 2, Trunc(v_Date), Trunc(����ʱ��_In)) And Nvl(����id, 0) = ִ�в���id_In And
            Nvl(��Ŀid, 0) = �շ�ϸĿid_In And (���� = ����_In Or ���� Is Null);
      If Sql%RowCount = 0 Then
        Insert Into ���˹ҺŻ���
          (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, ��Լ��, �����ѽ���)
        Values
          (Decode(ԤԼ��־_In, 2, Trunc(v_Date), Trunc(����ʱ��_In)), ִ�в���id_In, �շ�ϸĿid_In, ҽ������_In,
           Decode(ҽ��id_In, 0, Null, ҽ��id_In), ����_In, 1, Nvl(n_ԤԼ��, 0), Decode(ԤԼ��־_In, 0, 0, 1));
      End If;
    Else
      --ԤԼ�Һ�
      Update ���˹ҺŻ���
      Set ��Լ�� = Nvl(��Լ��, 0) + 1
      Where ���� = Trunc(v_Date) And Nvl(����id, 0) = ִ�в���id_In And Nvl(��Ŀid, 0) = �շ�ϸĿid_In And (���� = ����_In Or ���� Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into ���˹ҺŻ���
          (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, ��Լ��)
        Values
          (Trunc(v_Date), ִ�в���id_In, �շ�ϸĿid_In, ҽ������_In, Decode(ҽ��id_In, 0, Null, ҽ��id_In), ����_In, 1);
      End If;
    End If;
  Else
    --������Ű�ģʽ
    Begin
      Select Nvl(�Ƿ��ʱ��, 0) Into n_ʱ�� From �ٴ������¼ Where ID = �����¼id_In;
    Exception
      When Others Then
        n_ʱ�� := 0;
    End;
    Select ��Դid Into n_��Դid From �ٴ������¼ Where ID = �����¼id_In;
    n_����ģʽ := Nvl(zl_GetSysParameter('ԤԼ����ģʽ', 1111), 0);
    --��ʱ�εĺű�ֻ�ܵ������
    If n_ʱ�� = 1 And Nvl(ԤԼ��־_In, 0) = 2 And ��������_In = 0 And n_����ģʽ = 0 Then
      If Trunc(����ʱ��_In) <> Trunc(Sysdate) Then
        v_Err_Msg := '��ʱ�ε�ԤԼ�Һŵ�ֻ�ܵ�����գ�';
        Raise Err_Item;
      End If;
    End If;
    If Nvl(ԤԼ��־_In, 0) <> 2 Or ��������_In = 1 Then
      v_Date := Trunc(����ʱ��_In);
    Else
      If n_����ģʽ = 0 Then
        v_Date := Trunc(Sysdate);
      Else
        v_Date := Trunc(����ʱ��_In);
      End If;
    End If;
  
    n_ԤԼ�� := 0;
    If Nvl(ԤԼ��־_In, 0) <> 1 Then
      --��ԤԼ�Һ�;��ԤԼ����
      If Nvl(ԤԼ��־_In, 0) = 2 And v_Date <> Trunc(����ʱ��_In) Then
        --1.��ȥԤԼ���ڵ�ԤԼ��;
        --2-���ϵ�ǰԤԼ���ڵĹҺ���;
        Update �ٴ������¼ Set ��Լ�� = Nvl(��Լ��, 0) - 1 Where ID = �����¼id_In Returning ��Լ�� Into n_ԤԼ��;
        If n_ԤԼ�� < 0 Then
          Update �ٴ������¼ Set ��Լ�� = 0 Where ID = �����¼id_In Returning ��Լ�� Into n_ԤԼ��;
        End If;
        Update ���˹ҺŻ���
        Set ��Լ�� = Nvl(��Լ��, 0) - 1
        Where ���� = Trunc(����ʱ��_In) And Nvl(ҽ��id, 0) = Nvl(ҽ��id_In, 0) And Nvl(ҽ������, '-') = Nvl(ҽ������_In, '-') And
              Nvl(����id, 0) = ִ�в���id_In And Nvl(��Ŀid, 0) = �շ�ϸĿid_In And (���� = ����_In Or ���� Is Null)
        Returning ��Լ�� Into n_ԤԼ��;
        If n_ԤԼ�� < 0 Then
          Update ���˹ҺŻ���
          Set ��Լ�� = 0
          Where ���� = Trunc(����ʱ��_In) And Nvl(ҽ��id, 0) = Nvl(ҽ��id_In, 0) And Nvl(ҽ������, '-') = Nvl(ҽ������_In, '-') And
                Nvl(����id, 0) = ִ�в���id_In And Nvl(��Ŀid, 0) = �շ�ϸĿid_In And (���� = ����_In Or ���� Is Null)
          Returning ��Լ�� Into n_ԤԼ��;
        End If;
        n_ԤԼ�� := 1;
      Elsif Nvl(ԤԼ��־_In, 0) = 3 Then
        n_ԤԼ�� := 1;
      End If;
    
      Update �ٴ������¼
      Set �ѹ��� = Nvl(�ѹ���, 0) + 1, �����ѽ��� = Nvl(�����ѽ���, 0) + Decode(ԤԼ��־_In, 0, 0, 1), ��Լ�� = Nvl(��Լ��, 0) + Nvl(n_ԤԼ��, 0)
      Where ID = �����¼id_In;
    
      Update ���˹ҺŻ���
      Set �ѹ��� = Nvl(�ѹ���, 0) + 1, �����ѽ��� = Nvl(�����ѽ���, 0) + Decode(ԤԼ��־_In, 0, 0, 1), ��Լ�� = Nvl(��Լ��, 0) + Nvl(n_ԤԼ��, 0)
      Where ���� = Decode(ԤԼ��־_In, 2, Trunc(v_Date), Trunc(����ʱ��_In)) And Nvl(����id, 0) = ִ�в���id_In And
            Nvl(��Ŀid, 0) = �շ�ϸĿid_In And Nvl(ҽ��id, 0) = Nvl(ҽ��id_In, 0) And Nvl(ҽ������, '-') = Nvl(ҽ������_In, '-') And
            (���� = ����_In Or ���� Is Null);
      If Sql%RowCount = 0 Then
        Insert Into ���˹ҺŻ���
          (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, ��Լ��, �����ѽ���)
        Values
          (Decode(ԤԼ��־_In, 2, Trunc(v_Date), Trunc(����ʱ��_In)), ִ�в���id_In, �շ�ϸĿid_In, ҽ������_In,
           Decode(ҽ��id_In, 0, Null, ҽ��id_In), ����_In, 1, Nvl(n_ԤԼ��, 0), Decode(ԤԼ��־_In, 0, 0, 1));
      End If;
    Else
      --ԤԼ�Һ�
      Update �ٴ������¼ Set ��Լ�� = Nvl(��Լ��, 0) + 1 Where ID = �����¼id_In;
      Update ���˹ҺŻ���
      Set ��Լ�� = Nvl(��Լ��, 0) + 1
      Where ���� = Trunc(v_Date) And Nvl(ҽ��id, 0) = Nvl(ҽ��id_In, 0) And Nvl(ҽ������, '-') = Nvl(ҽ������_In, '-') And
            Nvl(����id, 0) = ִ�в���id_In And Nvl(��Ŀid, 0) = �շ�ϸĿid_In And (���� = ����_In Or ���� Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into ���˹ҺŻ���
          (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, ��Լ��)
        Values
          (Trunc(v_Date), ִ�в���id_In, �շ�ϸĿid_In, ҽ������_In, Decode(ҽ��id_In, 0, Null, ҽ��id_In), ����_In, 1);
      End If;
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˹ҺŻ���_Update;
/


Create Or Replace Procedure Zl_���˹Һż�¼_Insert
(
  ����id_In        ������ü�¼.����id%Type,
  �����_In        ������ü�¼.��ʶ��%Type,
  ����_In          ������ü�¼.����%Type,
  �Ա�_In          ������ü�¼.�Ա�%Type,
  ����_In          ������ü�¼.����%Type,
  ���ʽ_In      ������ü�¼.���ʽ%Type, --���ڴ�Ų��˵�ҽ�Ƹ��ʽ���
  �ѱ�_In          ������ü�¼.�ѱ�%Type,
  ���ݺ�_In        ������ü�¼.No%Type,
  Ʊ�ݺ�_In        ������ü�¼.ʵ��Ʊ��%Type,
  ���_In          ������ü�¼.���%Type,
  �۸񸸺�_In      ������ü�¼.�۸񸸺�%Type,
  ��������_In      ������ü�¼.��������%Type,
  �շ����_In      ������ü�¼.�շ����%Type,
  �շ�ϸĿid_In    ������ü�¼.�շ�ϸĿid%Type,
  ����_In          ������ü�¼.����%Type,
  ��׼����_In      ������ü�¼.��׼����%Type,
  ������Ŀid_In    ������ü�¼.������Ŀid%Type,
  �վݷ�Ŀ_In      ������ü�¼.�վݷ�Ŀ%Type,
  ���㷽ʽ_In      ����Ԥ����¼.���㷽ʽ%Type, --�ֽ�Ľ�������
  Ӧ�ս��_In      ������ü�¼.Ӧ�ս��%Type,
  ʵ�ս��_In      ������ü�¼.ʵ�ս��%Type,
  ���˿���id_In    ������ü�¼.���˿���id%Type,
  ��������id_In    ������ü�¼.��������id%Type,
  ִ�в���id_In    ������ü�¼.ִ�в���id%Type,
  ����Ա���_In    ������ü�¼.����Ա���%Type,
  ����Ա����_In    ������ü�¼.����Ա����%Type,
  ����ʱ��_In      ������ü�¼.����ʱ��%Type,
  �Ǽ�ʱ��_In      ������ü�¼.�Ǽ�ʱ��%Type,
  ҽ������_In      �ҺŰ���.ҽ������%Type,
  ҽ��id_In        �ҺŰ���.ҽ��id%Type,
  ������_In        Number, --������¼�Ƿ���������
  ����_In          Number,
  �ű�_In          �ҺŰ���.����%Type,
  ����_In          ������ü�¼.��ҩ����%Type,
  ����id_In        ������ü�¼.����id%Type,
  ����id_In        Ʊ��ʹ����ϸ.����id%Type,
  Ԥ��֧��_In      ����Ԥ����¼.��Ԥ��%Type, --ˢ���Һ�ʱʹ�õ�Ԥ�����,���Ϊ1����.
  �ֽ�֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�ֽ�֧�����ݽ��,���Ϊ1����.
  ����֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�����ʻ�֧�����,,���Ϊ1����.
  ���մ���id_In    ������ü�¼.���մ���id%Type,
  ������Ŀ��_In    ������ü�¼.������Ŀ��%Type,
  ͳ����_In      ������ü�¼.ͳ����%Type,
  ժҪ_In          ������ü�¼.ժҪ%Type, --ԤԼ�Һ�ժҪ��Ϣ
  ԤԼ�Һ�_In      Number := 0, --ԤԼ�Һ�ʱ��(��¼״̬=0,����ʱ��ΪԤԼʱ��),��ʱ����Ҫ���������ز���
  �շ�Ʊ��_In      Number := 0, --�Һ��Ƿ�ʹ���շ�Ʊ��
  ���ձ���_In      ������ü�¼.���ձ���%Type,
  ����_In          ���˹Һż�¼.����%Type := 0,
  ����_In          �Һ����״̬.���%Type := Null, --ԤԼʱ������ü�¼�ķ�ҩ�����ֶ�,�Һ�ʱ����Һż�¼
  ����_In          ���˹Һż�¼.����%Type := Null,
  ԤԼ����_In      Number := 0,
  ԤԼ��ʽ_In      ԤԼ��ʽ.����%Type := Null,
  ���ɶ���_In      Number := 0,
  �����id_In      ����Ԥ����¼.�����id%Type := Null,
  ���㿨���_In    ����Ԥ����¼.���㿨���%Type := Null,
  ����_In          ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In      ����Ԥ����¼.����˵��%Type := Null,
  ������λ_In      ����Ԥ����¼.������λ%Type := Null,
  ��������_In      Number := 0,
  ����_In          ���˹Һż�¼.����%Type := Null,
  ����ģʽ_In      Number := 0,
  ���ʷ���_In      Number := 0,
  �˺�����_In      Number := 1,
  ��Ԥ������ids_In Varchar2 := Null,
  �������˷ѱ�_In  Number := 0
) As
  ---------------------------------------------------------------------------
  --
  --����:
  --     ��������_in:0-�����ҺŻ���ԤԼ 1-����Աӵ�мӺ�Ȩ�޼Ӻ�
  --     �������˷ѱ�_In:0-���޸Ĳ��˷ѱ� 1-�޸Ĳ��˷ѱ�
  ----------------------------------------------------------------------------
  --���α������շѳ�Ԥ���Ŀ���Ԥ���б�
  --��ID�������ȳ��ϴ�δ����ġ�
  Cursor c_Deposit
  (
    v_����id        ������Ϣ.����id%Type,
    v_��Ԥ������ids Varchar2
  ) Is
    Select *
    From (Select a.Id, a.����id, a.��¼״̬, a.Ԥ�����, a.No, Nvl(a.���, 0) As ���
           From ����Ԥ����¼ A,
                (Select NO, Sum(Nvl(a.���, 0)) As ���
                  From ����Ԥ����¼ A
                  Where a.����id Is Null And Nvl(a.���, 0) <> 0 And
                        a.����id In (Select Column_Value From Table(f_Num2list(v_��Ԥ������ids))) And Nvl(a.Ԥ�����, 2) = 1
                  Group By NO
                  Having Sum(Nvl(a.���, 0)) <> 0) B
           Where a.����id Is Null And Nvl(a.���, 0) <> 0 And a.���㷽ʽ Not In (Select ���� From ���㷽ʽ Where ���� = 5) And
                 a.No = b.No And a.����id In (Select Column_Value From Table(f_Num2list(v_��Ԥ������ids))) And
                 Nvl(a.Ԥ�����, 2) = 1
           Union All
           Select 0 As ID, Max(����id) As ����id, ��¼״̬, Ԥ�����, NO, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���
           From ����Ԥ����¼
           Where ��¼���� In (1, 11) And ����id Is Not Null And Nvl(���, 0) <> Nvl(��Ԥ��, 0) And
                 ����id In (Select Column_Value From Table(f_Num2list(v_��Ԥ������ids))) And Nvl(Ԥ�����, 2) = 1 Having
            Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
           Group By ��¼״̬, NO, Ԥ�����)
    Order By Decode(����id, Nvl(v_����id, 0), 0, 1), ID, NO;
  --���ܣ�����һ�в��˹Һŷ��ã�����������ܵ�����Ԥ����¼
  --       ͬʱ������صĻ��ܱ�(���˹ҺŻ��ܡ����û���)
  --       ��һ�з��ô���Ʊ��ʹ�����(����ID_IN>0)

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_�ŶӺ��� �ŶӽкŶ���.�ŶӺ���%Type;
  v_�ֽ�     ���㷽ʽ.����%Type;
  v_�����ʻ� ���㷽ʽ.����%Type;
  v_�������� �ŶӽкŶ���.��������%Type;

  n_��ʱ��       Number;
  n_ʱ���޺�     Number;
  n_ʱ����Լ     Number;
  d_ʱ��ʱ��     Date;
  d_������ʱ�� Date;
  n_׷�Ӻ�       Number := 0; --����ʱ�ι��� ׷�ӹҺŵ���� 
  n_��Լ��       ���˹ҺŻ���.��Լ��%Type;
  n_�ѽ���       ���˹ҺŻ���.�����ѽ���%Type;
  n_ԤԼ��Чʱ�� Number;
  n_ʧЧ��       Number;
  n_ʧԼ�Һ�     Number := 0;
  n_��������     Number;
  n_����         Number := 0;

  n_��ӡid        Ʊ�ݴ�ӡ����.Id%Type;
  n_����id        ������ü�¼.Id%Type;
  n_Ԥ�����      ����Ԥ����¼.���%Type;
  n_��ǰ���      ����Ԥ����¼.���%Type;
  n_����ֵ        ����Ԥ����¼.���%Type;
  n_Ԥ��id        ����Ԥ����¼.Id%Type;
  n_���ѿ�id      ���ѿ�Ŀ¼.Id%Type;
  n_�Һ�id        ���˹Һż�¼.Id%Type;
  v_��Ԥ������ids Varchar2(4000);

  n_��id           ����ɿ����.Id%Type;
  n_�����         ������Ϣ.�����%Type;
  n_���           �Һ����״̬.���%Type;
  n_�������       �Һ����״̬.���%Type;
  n_��ſ���       �ҺŰ���.��ſ���%Type;
  n_����̨ǩ���Ŷ� Number;
  n_Count          Number;
  n_�޺���         Number(18);
  n_���ƿ�         Number;
  d_�Ŷ�ʱ��       Date;
  d_���ʱ��       Date;
  v_�ѱ�           �ѱ�.����%Type;
  n_ʱ�����       Number := -1;
  n_ԤԼ���ɶ���   Number;
  n_����id         �ҺŰ���.Id%Type;
  n_�ƻ�id         �ҺŰ��żƻ�.Id%Type := 0;
  v_����           �ҺŰ�������.������Ŀ%Type;
  n_��Լ��         Number(18);
  n_�ѹ���         Number(4) := 0;

  n_�ҳ��������� Number(4) := 0;
  n_����ģʽ       ������Ϣ.����ģʽ%Type;
  v_�Ŷ����       �ŶӽкŶ���.�Ŷ����%Type;
  v_������         �Һ����״̬.������%Type;
  v_��Ų���Ա     �Һ����״̬.����Ա����%Type;
  v_��Ż�����     �Һ����״̬.������%Type;
  v_���ʽ       ���˹Һż�¼.ҽ�Ƹ��ʽ%Type;
  v_Temp           Varchar2(3000);
  d_����ʱ��       Date;
Begin
  --��ȡ��ǰ��������
  Select Terminal Into v_������ From V$session Where Audsid = Userenv('sessionid');
  n_��id          := Zl_Get��id(����Ա����_In);
  v_��Ԥ������ids := Nvl(��Ԥ������ids_In, ����id_In);

  If �ѱ�_In Is Null Then
    Begin
      Select ���� Into v_�ѱ� From �ѱ� Where ȱʡ��־ = 1 And Rownum < 2;
      Update ������Ϣ Set �ѱ� = v_�ѱ� Where ����id = ����id_In;
    Exception
      When Others Then
        v_Err_Msg := '�޷�ȷ�����˷ѱ�����ȱʡ�ѱ��Ƿ���ȷ���ã�';
        Raise Err_Item;
    End;
  Else
    v_�ѱ� := �ѱ�_In;
    If Nvl(�������˷ѱ�_In, 0) = 1 Then
      Begin
        Update ������Ϣ Set �ѱ� = v_�ѱ� Where ����id = ����id_In;
      Exception
        When Others Then
          v_Err_Msg := 'û���ҵ���Ӧ�Ĳ��ˣ�';
          Raise Err_Item;
      End;
    End If;
  End If;

  If �����_In Is Not Null Then
    Begin
      Select Nvl(�����, 0) Into n_����� From ������Ϣ Where ����id = ����id_In;
    Exception
      When Others Then
        n_����� := 0;
    End;
    If n_����� = 0 Then
      Update ������Ϣ Set ����� = �����_In Where ����id = ����id_In;
    End If;
  End If;

  Begin
    Delete From �Һ����״̬
    Where ���� = �ű�_In And ���� = ����ʱ��_In And ��� = ����_In And ״̬ = 3 And ����Ա���� = ����Ա����_In;
  Exception
    When Others Then
      Null;
  End;

  v_Temp := zl_GetSysParameter(256);
  If v_Temp Is Null Or Substr(v_Temp, 1, 1) = '0' Then
    Null;
  Else
    Begin
      d_����ʱ�� := To_Date(Substr(v_Temp, 3), 'YYYY-MM-DD hh24:mi:ss');
      If ����ʱ��_In > d_����ʱ�� Then
        v_Err_Msg := '��ǰ�Һŵķ���ʱ��' || To_Char(����ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '�Ѿ������˳�����Ű�ģʽ,������ʹ�üƻ��Ű�ģʽ�Һ�!';
        Raise Err_Item;
      End If;
    Exception
      When Others Then
        Null;
    End;
  End If;

  If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
    --�ҺŻ���ԤԼ����
    --��Ϊ�����а��ձ൥�ݺŹ���,�չҺ������ܳ���10000��,����Ҫ���ΨһԼ����
    Select Count(*)
    Into n_Count
    From ������ü�¼
    Where ��¼���� = 4 And ��¼״̬ In (1, 3) And ��� = ���_In And NO = ���ݺ�_In;
    If n_Count <> 0 Then
      v_Err_Msg := '�Һŵ��ݺ��ظ�,���ܱ��棡' || Chr(13) || '���ʹ���˰���˳����,���չҺ������ܳ���10000�˴Ρ�';
      Raise Err_Item;
    End If;
  
    --��ȡ���㷽ʽ����
    Begin
      Select ���� Into v_�ֽ� From ���㷽ʽ Where ���� = 1;
    Exception
      When Others Then
        v_�ֽ� := '�ֽ�';
    End;
    Begin
      Select ���� Into v_�����ʻ� From ���㷽ʽ Where ���� = 3;
    Exception
      When Others Then
        v_�����ʻ� := '�����ʻ�';
    End;
  End If;

  n_��� := ����_In;
  Select Decode(To_Char(����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', Null)
  Into v_����
  From Dual;

  --�ҺŻ�ȡ����
  Begin
    Select a.Id, a.��ſ���, Nvl(b.�޺���, 0), Nvl(b.��Լ��, 0)
    Into n_����id, n_��ſ���, n_�޺���, n_��Լ��
    From �ҺŰ��� A, �ҺŰ������� B
    Where a.Id = b.����id(+) And b.������Ŀ(+) = v_���� And a.���� = �ű�_In;
  
  Exception
    When Others Then
      n_����id := -1;
  End;

  --����ǲ����ѻ��ߺű�Ϊ��ʱ�����
  If Nvl(������_In, 0) = 0 Or �ű�_In Is Not Null Then
    If n_����id = -1 Then
      v_Err_Msg := '������Ӧ�ĹҺŰ�������,����';
      Raise Err_Item;
    End If;
  End If;

  If Nvl(ԤԼ�Һ�_In, 0) = 1 Then
    --���Ȼ�ȡ�ƻ�
    Begin
      Select ID
      Into n_�ƻ�id
      From �ҺŰ��żƻ�
      Where ����id = n_����id And ���ʱ�� Is Not Null And
            Nvl(��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) =
            (Select Max(a.��Чʱ��) As ��Ч
             From �ҺŰ��żƻ� A
             Where a.���ʱ�� Is Not Null And ����ʱ��_In Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                   Nvl(a.ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) And a.����id = n_����id) And
            ����ʱ��_In Between Nvl(��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
            Nvl(ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd'));
    
    Exception
      When Others Then
        n_�ƻ�id := 0;
    End;
    If Nvl(n_�ƻ�id, 0) <> 0 Then
      Begin
        --��ȡ�ƻ�������
        Select a.Id, a.��ſ���, Nvl(b.�޺���, 0) As �޺���, Nvl(b.��Լ��, 0) As ��Լ��
        Into n_�ƻ�id, n_��ſ���, n_�޺���, n_��Լ��
        From �ҺŰ��żƻ� A, �Һżƻ����� B
        Where a.���� = �ű�_In And a.Id = n_�ƻ�id And a.���ʱ�� Is Not Null And a.Id = b.�ƻ�id(+) And b.������Ŀ(+) = v_����;
      Exception
        When Others Then
          v_Err_Msg := '������Ӧ�ĹҺŰ��Ż�ƻ�����,����';
          Raise Err_Item;
      End;
    End If;
  End If;

  --��ȡ�Ƿ��ʱ��
  If Nvl(n_�ƻ�id, 0) = 0 Then
    Select Count(Rownum) Into n_��ʱ�� From �ҺŰ���ʱ�� Where ���� = v_���� And ����id = n_����id And Rownum <= 1;
  Else
    Select Count(Rownum) Into n_��ʱ�� From �Һżƻ�ʱ�� Where ���� = v_���� And �ƻ�id = n_�ƻ�id And Rownum <= 1;
  End If;

  --��ʱ�ιҺ�ʱ�ж��Ƿ��ǹ��ڹҺ� Ҳ����׷�Ӻŵ���� Ŀǰֻ���ר�Һŷ�ʱ�ν��д���
  If Nvl(ԤԼ�Һ�_In, 0) = 0 And n_��ʱ�� > 0 And Nvl(n_��ſ���, 0) = 1 And ����_In Is Null And Nvl(��������_In, 0) = 0 Then
    --����ʱ��_in>Sysdate ����ʱ��>����ʱ��ʱ��--����_in is null
    Begin
      Select Max(To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'))
      Into d_������ʱ��
      From �ҺŰ���ʱ��
      Where ����id = n_����id And ���� = v_���� And Nvl(��������, 0) <> 0;
      n_׷�Ӻ� := Case Sign(����ʱ��_In - d_������ʱ��)
                 When -1 Then
                  0
                 Else
                  1
               End;
    Exception
      When Others Then
        n_׷�Ӻ� := 0;
    End;
  End If;
  d_ʱ��ʱ�� := ����ʱ��_In;

  If ���_In = 1 And Nvl(ԤԼ�Һ�_In, 0) = 0 And n_��ʱ�� > 0 Then
    --�Һ�ʱ��� �Ƿ����ʱ��,����ʱ��,��ʱ�ε�����������ȡ����
    Begin
      Select Nvl(���, 0),
             To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
             ��������, Decode(Nvl(�Ƿ�ԤԼ, 0), 0, 0, ��������)
      Into n_ʱ�����, d_ʱ��ʱ��, n_ʱ���޺�, n_ʱ����Լ
      From �ҺŰ���ʱ��
      Where ����id = n_����id And ���� = v_���� And
            (���, ����id, ����) In (Select Nvl(Max(���), -1), ����id, ����
                               From �ҺŰ���ʱ��
                               Where ����id = n_����id And ���� = v_���� And
                                     Decode(��������_In + n_׷�Ӻ�, 0, To_Char(����ʱ��_In, 'hh24:mi'),
                                            To_Char(Nvl(��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')) =
                                     To_Char(Nvl(��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
                               Group By ����id, ����);
    Exception
      When Others Then
        n_ʱ����� := -1;
        n_��ʱ��   := 0;
        d_ʱ��ʱ�� := ����ʱ��_In;
        n_ʱ���޺� := 0;
        n_ʱ����Լ := 0;
    End;
  End If;

  If ���_In = 1 And Nvl(ԤԼ�Һ�_In, 0) = 1 And n_��ʱ�� > 0 Then
    --ԤԼ��,ȡ�ƻ�
    Begin
      If Nvl(n_�ƻ�id, 0) = 0 Then
        --û�ƻ���Ч,ȡ���ŵ�����
        Select Nvl(���, 0),
               To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(c.��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ʱ��ʱ��,
               ��������, Decode(Nvl(�Ƿ�ԤԼ, 0), 0, 0, ��������)
        Into n_ʱ�����, d_ʱ��ʱ��, n_ʱ���޺�, n_ʱ����Լ
        From �ҺŰ���ʱ�� C
        Where ����id = n_����id And ���� = v_���� And
              (���, ����id, ����) In
              (Select Nvl(Max(c.���), -1), ����id, ����
               From �ҺŰ���ʱ�� C
               Where ����id = n_����id And c.���� = v_���� And
                     Decode(��������_In, 1, To_Char(Nvl(c.��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi'),
                            To_Char(����ʱ��_In, 'hh24:mi')) =
                     To_Char(Nvl(c.��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
               Group By ����id, ����);
      Else
        --�мƻ���Чȡ�ƻ�
        --û��Ч�������ǴӹҺżƻ�ʱ�β�ѯ      
        Select Nvl(���, -1),
               To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(c.��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ʱ��ʱ��,
               ��������, Decode(Nvl(�Ƿ�ԤԼ, 0), 0, 0, ��������)
        Into n_ʱ�����, d_ʱ��ʱ��, n_ʱ���޺�, n_ʱ����Լ
        From �Һżƻ�ʱ�� C
        Where �ƻ�id = n_�ƻ�id And ���� = v_���� And
              (���, �ƻ�id, ����) In
              (Select Nvl(Max(c.���), -1), �ƻ�id, ����
               From �Һżƻ�ʱ�� C
               Where �ƻ�id = n_�ƻ�id And c.���� = v_���� And
                     Decode(��������_In, 1, To_Char(Nvl(c.��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi'),
                            To_Char(����ʱ��_In, 'hh24:mi')) =
                     To_Char(Nvl(c.��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
               Group By �ƻ�id, ����);
      End If;
    Exception
      When Others Then
        n_ʱ����� := -1;
        n_��ʱ��   := 0;
        d_ʱ��ʱ�� := ����ʱ��_In;
        n_ʱ���޺� := 0;
        n_ʱ����Լ := 0;
    End;
  End If;

  If ���_In = 1 Then
  
    --��ȡ��ǰδʹ�õ����
    If Nvl(n_��ſ���, 0) = 1 And n_��ʱ�� = 0 Then
      --<��ſ��� δ����ʱ�� ��ȡ���õ�������,�Լ��Ѿ�ʹ�õ�����>     
      Begin
        --������
        If �˺�����_In = 1 Then
          Select Max(Nvl(���, 0)), Sum(Decode(���, Nvl(����_In, 0), 0, 1)), Sum(Nvl(ԤԼ, 0))
          Into n_�������, n_��������, n_��Լ��
          From �Һ����״̬
          Where ���� = �ű�_In And ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ״̬ Not In (4, 5);
        Else
          Select Max(Nvl(���, 0)), Sum(Decode(���, Nvl(����_In, 0), 0, 1)), Sum(Nvl(ԤԼ, 0))
          Into n_�������, n_��������, n_��Լ��
          From �Һ����״̬
          Where ���� = �ű�_In And ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ״̬ <> 5;
        End If;
      Exception
        When Others Then
          n_������� := 0;
          n_�������� := 0;
      End;
      If n_��� Is Null Then
        n_��� := Nvl(n_�������, 0) + 1;
      End If;
      --<��ſ��� δ����ʱ�� ��ȡ���õ������� �Լ��Ѿ�ʹ�õ����� --end>
    
      --�ǼӺŵ������Ҫ����Ƿ񳬹�������
      If ��������_In = 0 Then
        If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
          --�Һ�ʱ���       
          --������ſ���δ��ʱ�� �ﵽ������
          If n_�޺��� <= n_�������� And n_�޺��� > 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(Trunc(����ʱ��_In), 'yyyy-mm-dd ') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        
        Else
          --ԤԼʱ���
          If n_��Լ�� = 0 Then
            n_��Լ�� := n_�޺���;
          End If;
        
          If n_��Լ�� <= n_��Լ�� And n_��Լ�� > 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(Trunc(����ʱ��_In), 'yyyy-mm-dd') || '�Ѵﵽ�����Լ����';
            Raise Err_Item;
          End If;
        End If;
      
      Else
        Null;
        --������ſ���,δ��ʱ�� �Ӻ����   ������,����Ժ������������Ժ󲹳�
      End If;
    
    Elsif Nvl(n_��ʱ��, 0) > 0 And Nvl(n_��ſ���, 0) = 0 Then
      --<--��ͨ�ŷ�ʱ�� ����ֻ��ԤԼһ�����-->
      If ��������_In = 0 Then
        --<����ԤԼ�Һ�-->
        Begin
          Select Count(0) As �ѹ���, Nvl(Sum(Decode(Nvl(Sign(a.���� - d_ʱ��ʱ��), 0), 0, 1, 0)), 0) As ��Լ��
          Into n_�ѹ���, n_��Լ��
          From �Һ����״̬ A
          Where a.���� = �ű�_In And ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And
                ״̬ Not In (4, 5);
        Exception
          When Others Then
            n_�ѹ��� := 0;
            n_��Լ�� := 0;
        End;
      
        n_ʱ����Լ := n_ʱ���޺�; --��ͨ�ŷ�ʱ�ε����,29��n_ʱ����Լʼ����0 �������⴦��
        --�����������  
        If n_��Լ�� = 0 Then
          n_��Լ�� := n_�޺���;
        End If;
        If n_ʱ����Լ <= n_��Լ�� Or n_��Լ�� <= n_��Լ�� Then
          v_Err_Msg := '�ű�' || �ű�_In || '��ʱ��' || To_Char(d_ʱ��ʱ��, 'hh24:mi') || '�Ѵﵽ�����Լ����';
          Raise Err_Item;
        End If;
        If n_�޺��� <= n_�ѹ��� Then
          v_Err_Msg := '�ű�' || �ű�_In || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
          Raise Err_Item;
        End If;
      End If;
    
      --û�дﵽʱ�ε��޺��� �����ڵ�ǰʱ������׷��
    
      --��ȡ����ҳ���������
      If �˺�����_In = 1 Then
        Select Nvl(Max(���), 0)
        Into n_�ҳ���������
        From �Һ����״̬ A
        Where a.���� = d_ʱ��ʱ�� And ���� = �ű�_In And ״̬ Not In (4, 5);
      Else
        Select Nvl(Max(���), 0)
        Into n_�ҳ���������
        From �Һ����״̬ A
        Where a.���� = d_ʱ��ʱ�� And ���� = �ű�_In And ״̬ <> 5;
      End If;
    
      --�������
      n_��� := RPad(Nvl(n_ʱ�����, 0), Length(n_�޺���) + Length(Nvl(n_ʱ�����, 0)), 0) + n_��Լ�� + 1;
      If n_��� <= Nvl(n_�ҳ���������, 0) Then
        n_��� := Nvl(n_�ҳ���������, 0) + 1;
      End If;
    
      --<--��ͨ�ŷ�ʱ��--End>
    Elsif Nvl(n_��ʱ��, 0) > 0 And Nvl(n_��ſ���, 0) = 1 Then
      --<������ſ��� ����ʱ��
      --ר�Һŷ�ʱ��
      Begin
        If �˺�����_In = 1 Then
          Select Max(���), Sum(Decode(���, Nvl(����_In, 0), 0, 1)), Sum(Decode(Sign(���� - d_ʱ��ʱ��), 0, 1, 0))
          Into n_�������, n_�ѹ���, n_��������
          From �Һ����״̬
          Where ���� = �ű�_In And ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ״̬ Not In (4, 5);
        Else
          Select Max(���), Sum(Decode(���, Nvl(����_In, 0), 0, 1)), Sum(Decode(Sign(���� - d_ʱ��ʱ��), 0, 1, 0))
          Into n_�������, n_�ѹ���, n_��������
          From �Һ����״̬
          Where ���� = �ű�_In And ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ״̬ <> 5;
        End If;
      Exception
        When Others Then
          n_������� := 0;
          n_�������� := 0;
          n_�ѹ���   := 0;
      End;
    
      n_ʧЧ�� := 0;
      If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
        n_ԤԼ��Чʱ�� := Zl_To_Number(zl_GetSysParameter('ԤԼ��Чʱ��', 1111));
        n_ʧԼ�Һ�     := Zl_To_Number(zl_GetSysParameter('ʧԼ���ڹҺ�', 1111));
        If Nvl(n_ԤԼ��Чʱ��, 0) <> 0 And Nvl(n_ʧԼ�Һ�, 0) > 0 Then
          Begin
            Select Sum(Decode(Sign((Sysdate - n_ԤԼ��Чʱ�� / 24 / 60) - ����), 1, 1, 0))
            Into n_ʧЧ��
            From �Һ����״̬
            Where ���� = �ű�_In And ���� Between Trunc(Sysdate) And Sysdate And Nvl(ԤԼ, 0) = 1 And ״̬ = 2;
          Exception
            When Others Then
              n_ʧЧ�� := 0;
          End;
        End If;
      End If;
    
      If ��������_In = 0 Then
        If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
        
          --�Һ� ׷�Ӻ���ʱ�����ʱ���޺���
          If n_ʱ���޺� <= n_�������� And Nvl(n_׷�Ӻ�, 0) = 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��ʱ��' || To_Char(d_ʱ��ʱ��, 'hh24:mi') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
          If n_�޺��� <= n_�ѹ��� - n_ʧЧ�� Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        
        Else
          --�Һ�
          If n_��Լ�� = 0 Then
            n_��Լ�� := n_�޺���;
          End If;
          If n_��Լ�� <= n_�������� Then
            v_Err_Msg := '�ű�' || �ű�_In || '��ʱ��' || To_Char(d_ʱ��ʱ��, 'hh24:mi') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
          If n_�޺��� <= n_�ѹ��� Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        End If;
      End If;
      If n_��� Is Null Then
        --�������
        If Nvl(n_�������, 0) < Nvl(n_ʱ�����, 0) Then
          n_������� := Nvl(n_ʱ�����, 0);
        End If;
        n_��� := Nvl(n_�������, 0) + 1;
      End If;
    Elsif Nvl(n_��ʱ��, 0) = 0 And Nvl(n_��ſ���, 0) = 0 And Nvl(������_In, 0) = 0 And Nvl(�ű�_In, 0) > 0 Then
      ---<--��ͨ��  -->
      Begin
        Select �ѹ���, ��Լ��
        Into n_��������, n_��Լ��
        From ���˹ҺŻ���
        Where ���� = Trunc(����ʱ��_In) And ���� = �ű�_In;
      Exception
        When Others Then
          n_�������� := 0;
          n_��Լ��   := 0;
      End;
      If Nvl(��������_In, 0) = 0 Then
        If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
          --�Һ�
          If Nvl(n_��������, 0) >= Nvl(n_�޺���, 0) And Nvl(n_�޺���, 0) > 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        Else
          --ԤԼ
          If (Nvl(n_��Լ��, 0) > 0 And Nvl(n_��Լ��, 0) >= Nvl(n_��Լ��, 0)) Or
             (Nvl(n_��������, 0) >= Nvl(n_�޺���, 0) And Nvl(n_�޺���, 0) > 0) Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        End If;
      End If;
      ---<--��ͨ��  -->
    End If;
  End If;

  --���¹Һ����״̬
  If ���_In = 1 And Not n_��� Is Null Then
    If n_��ʱ�� = 1 Then
      d_���ʱ�� := ����ʱ��_In;
    Else
      d_���ʱ�� := Trunc(����ʱ��_In);
    End If;
    --������ŵĴ���
    Begin
      Select ����Ա����, ������
      Into v_��Ų���Ա, v_��Ż�����
      From �Һ����״̬
      Where ״̬ = 5 And ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_���;
      n_���� := 1;
    Exception
      When Others Then
        v_��Ų���Ա := Null;
        v_��Ż����� := Null;
        n_����       := 0;
    End;
    If n_���� = 0 Then
      Update �Һ����״̬
      Set ״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1), ԤԼ = Decode(ԤԼ����_In, 1, 1, 0), �Ǽ�ʱ�� = Sysdate
      Where ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_��� And ״̬ = 3 And ����Ա���� = ����Ա����_In;
      If Sql%RowCount = 0 Then
        Begin
          If Nvl(n_��ʱ��, 0) = 0 Or Nvl(ԤԼ�Һ�_In, 0) = 1 Or (Nvl(n_��ſ���, 0) = 0 And Nvl(����_In, 0) = 0) Then
            Insert Into �Һ����״̬
              (����, ����, ���, ״̬, ����Ա����, ԤԼ, �Ǽ�ʱ��)
            Values
              (�ű�_In, d_���ʱ��, n_���, Decode(ԤԼ�Һ�_In, 1, 2, 1), ����Ա����_In, Decode(ԤԼ����_In, 1, 1, 0), Sysdate);
            If Nvl(ԤԼ�Һ�_In, 0) = 0 And ԤԼ��ʽ_In Is Not Null Then
              Update �Һ����״̬ Set ԤԼ = 1 Where ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_���;
            End If;
          Elsif Nvl(n_��ʱ��, 0) > 0 Then
            --��ʱ�κ�ר�Һ� ʧԼ��ԤԼ������Һ�
            Update �Һ����״̬
            Set ״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1), ����Ա���� = ����Ա����_In, ԤԼ = Decode(ԤԼ����_In, 1, 1, 0), �Ǽ�ʱ�� = Sysdate
            Where ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_��� And ״̬ = 2;
            If Sql%NotFound Then
              Insert Into �Һ����״̬
                (����, ����, ���, ״̬, ����Ա����, ԤԼ, �Ǽ�ʱ��)
              Values
                (�ű�_In, d_���ʱ��, n_���, Decode(ԤԼ�Һ�_In, 1, 2, 1), ����Ա����_In, Decode(ԤԼ����_In, 1, 1, 0), Sysdate);
            End If;
            If Nvl(ԤԼ�Һ�_In, 0) = 0 And ԤԼ��ʽ_In Is Not Null Then
              Update �Һ����״̬ Set ԤԼ = 1 Where ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_���;
            End If;
          End If;
        Exception
          When Others Then
            v_Err_Msg := '���' || n_��� || '�ѱ�ʹ��,������ѡ��һ�����.';
            Raise Err_Item;
        End;
      End If;
    Else
      If ����Ա����_In <> v_��Ų���Ա Or v_������ <> v_��Ż����� Then
        v_Err_Msg := '���' || n_��� || '�ѱ�������' || v_������ || '����,������ѡ��һ�����.';
        Raise Err_Item;
      Else
        Update �Һ����״̬
        Set ״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1), ԤԼ = Decode(ԤԼ����_In, 1, 1, 0), �Ǽ�ʱ�� = Sysdate
        Where ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_��� And ״̬ = 5 And ����Ա���� = ����Ա����_In And ������ = v_������;
        If Nvl(ԤԼ�Һ�_In, 0) = 0 And ԤԼ��ʽ_In Is Not Null Then
          Update �Һ����״̬ Set ԤԼ = 1 Where ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_���;
        End If;
      End If;
    End If;
  End If;

  --�������˹Һŷ���(���ܵ����ǻ������������)
  Select ���˷��ü�¼_Id.Nextval Into n_����id From Dual; --Ӧ��ͨ������õ�

  Insert Into ������ü�¼
    (ID, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, NO, ʵ��Ʊ��, �����־, �Ӱ��־, ���ӱ�־, ��ҩ����, ����id, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, ���˿���id, �շ����,
     ���㵥λ, �շ�ϸĿid, ������Ŀid, �վݷ�Ŀ, ����, ����, ��׼����, Ӧ�ս��, ʵ�ս��, ���ʽ��, ����id, ���ʷ���, ��������id, ������, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����,
     ����ʱ��, �Ǽ�ʱ��, ���մ���id, ������Ŀ��, ���ձ���, ͳ����, ժҪ, ����, �ɿ���id)
  Values
    (n_����id, 4, Decode(ԤԼ�Һ�_In, 1, 0, 1), ���_In, Decode(�۸񸸺�_In, 0, Null, �۸񸸺�_In), ��������_In, ���ݺ�_In, Ʊ�ݺ�_In, 1, ����_In,
     ������_In, Decode(ԤԼ�Һ�_In, 1, To_Char(n_���), ����_In), Decode(����id_In, 0, Null, ����id_In),
     Decode(�����_In, 0, Null, �����_In), ���ʽ_In, ����_In, Decode(����_In, Null, Null, �Ա�_In), Decode(����_In, Null, Null, ����_In),
     v_�ѱ�, ���˿���id_In, �շ����_In, �ű�_In, �շ�ϸĿid_In, ������Ŀid_In, �վݷ�Ŀ_In, 1, ����_In, ��׼����_In, Ӧ�ս��_In, ʵ�ս��_In,
     Decode(ԤԼ�Һ�_In, 1, Null, Decode(Nvl(���ʷ���_In, 0), 1, Null, ʵ�ս��_In)),
     Decode(ԤԼ�Һ�_In, 1, Null, Decode(Nvl(���ʷ���_In, 0), 1, Null, ����id_In)), Decode(Nvl(���ʷ���_In, 0), 1, 1, 0), ��������id_In,
     ����Ա����_In, Decode(ԤԼ�Һ�_In, 1, ����Ա����_In, Null), ִ�в���id_In, ҽ������_In, ����Ա���_In, ����Ա����_In, ����ʱ��_In, �Ǽ�ʱ��_In, ���մ���id_In,
     ������Ŀ��_In, ���ձ���_In, ͳ����_In, ժҪ_In, ԤԼ��ʽ_In, Decode(ԤԼ�Һ�_In, 1, Null, n_��id));

  --���ܽ��㵽����Ԥ����¼
  If Nvl(ԤԼ�Һ�_In, 0) = 0 And Nvl(���ʷ���_In, 0) = 0 Then
  
    If (Nvl(�ֽ�֧��_In, 0) <> 0 Or (Nvl(�ֽ�֧��_In, 0) = 0 And Nvl(����֧��_In, 0) = 0 And Nvl(Ԥ��֧��_In, 0) = 0)) And ���_In = 1 Then
      Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      Insert Into ����Ԥ����¼
        (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
         ��������)
      Values
        (n_Ԥ��id, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), Nvl(���㷽ʽ_In, v_�ֽ�), Nvl(�ֽ�֧��_In, 0), �Ǽ�ʱ��_In,
         ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�', n_��id, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, ������λ_In, 4);
    
      If Nvl(���㿨���_In, 0) <> 0 And Nvl(�ֽ�֧��_In, 0) <> 0 Then
      
        n_���ѿ�id := Null;
        Begin
          Select Nvl(���ƿ�, 0), 1 Into n_���ƿ�, n_Count From �����ѽӿ�Ŀ¼ Where ��� = ���㿨���_In;
        Exception
          When Others Then
            n_Count := 0;
        End;
        If n_Count = 0 Then
          v_Err_Msg := 'û�з���ԭ���㿨����Ӧ���,���ܼ���������';
          Raise Err_Item;
        End If;
        If n_���ƿ� = 1 Then
          Select ID
          Into n_���ѿ�id
          From ���ѿ�Ŀ¼
          Where �ӿڱ�� = ���㿨���_In And ���� = ����_In And
                ��� = (Select Max(���) From ���ѿ�Ŀ¼ Where �ӿڱ�� = ���㿨���_In And ���� = ����_In);
        End If;
        Zl_���˿������¼_Insert(���㿨���_In, n_���ѿ�id, ���㷽ʽ_In, �ֽ�֧��_In, ����_In, Null, �Ǽ�ʱ��_In, Null, ����id_In, n_Ԥ��id);
      End If;
    
    End If;
  
    --����ҽ���Һ�
    If Nvl(����֧��_In, 0) <> 0 And ���_In = 1 Then
      Insert Into ����Ԥ����¼
        (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), v_�����ʻ�, ����֧��_In, �Ǽ�ʱ��_In, ����Ա���_In,
         ����Ա����_In, ����id_In, 'ҽ���Һ�', n_��id, 4);
    End If;
  
    --���ھ��￨ͨ��Ԥ����Һ�
    If Nvl(Ԥ��֧��_In, 0) <> 0 And ���_In = 1 Then
      n_Ԥ����� := Ԥ��֧��_In;
      For r_Deposit In c_Deposit(����id_In, v_��Ԥ������ids) Loop
        n_��ǰ��� := Case
                    When r_Deposit.��� - n_Ԥ����� < 0 Then
                     r_Deposit.���
                    Else
                     n_Ԥ�����
                  End;
      
        If r_Deposit.Id <> 0 Then
          --��һ�γ�Ԥ��(���Ͻ���ID,���Ϊ0)
          Update ����Ԥ����¼ Set ��Ԥ�� = 0, ����id = ����id_In, �������� = 4 Where ID = r_Deposit.Id;
        
        End If;
        --���ϴ�ʣ���
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���,
           ��Ԥ��, ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
          Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�,
                 �Ǽ�ʱ��_In, ����Ա����_In, ����Ա���_In, n_��ǰ���, ����id_In, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
          From ����Ԥ����¼
          Where NO = r_Deposit.No And ��¼״̬ = r_Deposit.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
      
        --���²���Ԥ�����
        Update �������
        Set Ԥ����� = Nvl(Ԥ�����, 0) - n_��ǰ���
        Where ����id = r_Deposit.����id And ���� = 1 And ���� = Nvl(r_Deposit.Ԥ�����, 2);
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
      If n_Ԥ����� > 0 Then
        v_Err_Msg := 'Ԥ���಻��֧������֧�����,���ܼ���������';
        Raise Err_Item;
      
      End If;
      Delete From ������� Where ����id = ����id_In And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
    End If;
  
    --��ػ��ܱ�Ĵ���
    --��Ա�ɿ����
    If ���_In = 1 Then
      If Nvl(�ֽ�֧��_In, 0) <> 0 Then
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + �ֽ�֧��_In
        Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = Nvl(���㷽ʽ_In, v_�ֽ�)
        Returning ��� Into n_����ֵ;
      
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ����
            (�տ�Ա, ���㷽ʽ, ����, ���)
          Values
            (����Ա����_In, Nvl(���㷽ʽ_In, v_�ֽ�), 1, �ֽ�֧��_In);
          n_����ֵ := �ֽ�֧��_In;
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From ��Ա�ɿ����
          Where �տ�Ա = ����Ա����_In And ���㷽ʽ = Nvl(���㷽ʽ_In, v_�ֽ�) And ���� = 1 And Nvl(���, 0) = 0;
        End If;
      End If;
    
      If Nvl(����֧��_In, 0) <> 0 Then
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + ����֧��_In
        Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = v_�����ʻ�
        Returning ��� Into n_����ֵ;
      
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (����Ա����_In, v_�����ʻ�, 1, ����֧��_In);
          n_����ֵ := ����֧��_In;
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From ��Ա�ɿ����
          Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_�����ʻ� And Nvl(���, 0) = 0;
        End If;
      End If;
    End If;
  End If;

  --���˹ҺŻ���(ֻ����һ��,�ҵ�����ȡ�����Ѳ�����)
  If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
    If Nvl(���ʷ���_In, 0) = 0 Then
      --����Ʊ��ʹ�����
      If ���_In = 1 And Ʊ�ݺ�_In Is Not Null Then
        Select Ʊ�ݴ�ӡ����_Id.Nextval Into n_��ӡid From Dual;
      
        --����Ʊ��
        Insert Into Ʊ�ݴ�ӡ���� (ID, ��������, NO) Values (n_��ӡid, 4, ���ݺ�_In);
      
        Insert Into Ʊ��ʹ����ϸ
          (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
        Values
          (Ʊ��ʹ����ϸ_Id.Nextval, Decode(�շ�Ʊ��_In, 1, 1, 4), Ʊ�ݺ�_In, 1, 1, ����id_In, n_��ӡid, �Ǽ�ʱ��_In, ����Ա����_In);
      
        --״̬�Ķ�
        Update Ʊ�����ü�¼
        Set ��ǰ���� = Ʊ�ݺ�_In, ʣ������ = Decode(Sign(ʣ������ - 1), -1, 0, ʣ������ - 1), ʹ��ʱ�� = Sysdate
        Where ID = Nvl(����id_In, 0);
      End If;
    End If;
    --���˱��ξ���(�Է���ʱ��Ϊ׼)
    If Nvl(����id_In, 0) <> 0 And ���_In = 1 Then
      Update ������Ϣ Set ����ʱ�� = ����ʱ��_In, ����״̬ = 1, �������� = ����_In Where ����id = ����id_In;
    End If;
  End If;

  If Nvl(ԤԼ�Һ�_In, 0) = 0 And Nvl(���ʷ���_In, 0) = 1 Then
    --����
    If Nvl(����id_In, 0) = 0 Then
      v_Err_Msg := 'Ҫ��Բ��˵ĹҺŷѽ��м��ʣ������ǽ������˲��ܼ��ʹҺš�';
      Raise Err_Item;
    End If;
  
    --�������
    Update �������
    Set ������� = Nvl(�������, 0) + Nvl(ʵ�ս��_In, 0)
    Where ����id = Nvl(����id_In, 0) And ���� = 1 And ���� = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into ������� (����id, ����, ����, �������, Ԥ�����) Values (����id_In, 1, 1, Nvl(ʵ�ս��_In, 0), 0);
    End If;
  
    --����δ�����
    Update ����δ�����
    Set ��� = Nvl(���, 0) + Nvl(ʵ�ս��_In, 0)
    Where ����id = ����id_In And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(���˿���id_In, 0) And
          Nvl(��������id, 0) = Nvl(��������id_In, 0) And Nvl(ִ�в���id, 0) = Nvl(ִ�в���id_In, 0) And ������Ŀid + 0 = ������Ŀid_In And
          ��Դ;�� + 0 = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into ����δ�����
        (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
      Values
        (����id_In, Null, Null, ���˿���id_In, ��������id_In, ִ�в���id_In, ������Ŀid_In, 1, Nvl(ʵ�ս��_In, 0));
    End If;
  End If;

  --���˹Һż�¼
  If �ű�_In Is Not Null And ���_In = 1 Then
    --And Nvl(ԤԼ�Һ�_In, 0) = 0
    Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
    Begin
      Select ���� Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = ���ʽ_In And Rownum < 2;
    Exception
      When Others Then
        v_���ʽ := Null;
    End;
    Insert Into ���˹Һż�¼
      (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ԤԼʱ��, ����Ա���,
       ����Ա����, ����, ����, ����, ԤԼ, ԤԼ��ʽ, ժҪ, ������ˮ��, ����˵��, ������λ, ����ʱ��, ������, ԤԼ����Ա, ԤԼ����Ա���, ����, ҽ�Ƹ��ʽ)
    Values
      (n_�Һ�id, ���ݺ�_In, Decode(Nvl(ԤԼ�Һ�_In, 0), 1, 2, 1), 1, ����id_In, �����_In, ����_In, �Ա�_In, ����_In, �ű�_In, ����_In, ����_In,
       Null, ִ�в���id_In, ҽ������_In, 0, Null, �Ǽ�ʱ��_In, ����ʱ��_In, Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����ʱ��_In, Null), ����Ա���_In,
       ����Ա����_In, ����_In, n_���, ����_In, Decode(ԤԼ����_In, 1, 1, 0), ԤԼ��ʽ_In, ժҪ_In, ������ˮ��_In, ����˵��_In, ������λ_In,
       Decode(Nvl(ԤԼ�Һ�_In, 0), 0, �Ǽ�ʱ��_In, Null), Decode(Nvl(ԤԼ�Һ�_In, 0), 0, ����Ա����_In, Null),
       Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����Ա����_In, Null), Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����Ա���_In, Null),
       Decode(Nvl(����_In, 0), 0, Null, ����_In), v_���ʽ);
    If Nvl(ԤԼ�Һ�_In, 0) = 0 And ԤԼ��ʽ_In Is Not Null Then
      Update ���˹Һż�¼
      Set ԤԼ = 1, ԤԼʱ�� = ����ʱ��_In, ԤԼ����Ա = ����Ա����_In, ԤԼ����Ա��� = ����Ա���_In
      Where ID = n_�Һ�id;
    End If;
    n_ԤԼ���ɶ��� := 0;
    If Nvl(ԤԼ�Һ�_In, 0) = 1 Then
      n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
    End If;
  
    --0-����������;1-��ҽ�������̨�Ŷ�;2-�ȷ���,��ҽ��վ
    If Nvl(���ɶ���_In, 0) <> 0 And Nvl(ԤԼ�Һ�_In, 0) = 0 Or n_ԤԼ���ɶ��� = 1 Then
      n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113));
      If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Or n_ԤԼ���ɶ��� = 1 Then
      
        --��������
        --.����ִ�в��š� �ķ�ʽ���ɶ���
        v_�������� := ִ�в���id_In;
        v_�ŶӺ��� := Zlgetnextqueue(ִ�в���id_In, n_�Һ�id, �ű�_In || '|' || n_���);
        v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 0);
        d_�Ŷ�ʱ�� := Zl_Get_Queuedate(n_�Һ�id, �ű�_In, n_���, ����ʱ��_In);
        --  ��������_In , ҵ������_In, ҵ��id_In,����id_In,�ŶӺ���_In,�Ŷӱ��_In,��������_In,����ID_IN, ����_In, ҽ������_In, �Ŷ�ʱ��_In
        Zl_�ŶӽкŶ���_Insert(v_��������, 0, n_�Һ�id, ִ�в���id_In, v_�ŶӺ���, Null, ����_In, ����id_In, ����_In, ҽ������_In, d_�Ŷ�ʱ��, ԤԼ��ʽ_In,
                         Null, v_�Ŷ����);
      End If;
    End If;
  End If;
  --���˵�����Ϣ
  If ����id_In Is Not Null And ���_In = 1 Then
    --ȡ����:
    If Nvl(n_����ģʽ, 0) <> Nvl(����ģʽ_In, 0) Then
      --����ģʽ��ȷ��
    
      v_Err_Msg := Null;
      Begin
        Select Nvl(����ģʽ, 0) Into n_����ģʽ From ������Ϣ Where ����id = ����id_In;
      Exception
        When Others Then
          v_Err_Msg := 'δ�ҵ�ָ���Ĳ�����Ϣ,������Һ�';
      End;
    
      If v_Err_Msg Is Not Null Then
        Raise Err_Item;
      End If;
      If n_����ģʽ = 1 And Nvl(����ģʽ_In, 0) = 0 Then
        --�����Ѿ���"�����ƺ�����",������"�Ƚ�������Ƶ�",�����Ƿ����δ������
        Select Count(1)
        Into n_Count
        From ����δ�����
        Where ����id = ����id_In And (��Դ;�� In (1, 4) Or ��Դ;�� = 3 And Nvl(��ҳid, 0) = 0) And Nvl(���, 0) <> 0 And Rownum < 2;
        If Nvl(n_Count, 0) <> 0 Then
          --����δ�������ݣ������Ƚ���������ִ��
          v_Err_Msg := '��ǰ���˵ľ���ģʽΪ�����ƺ�����Ҵ���δ����ã�����������ò��˵ľ���ģʽ,������ȶ�δ����ý��ʺ��ٹҺŻ򲻵������˵ľ���ģʽ!';
          Raise Err_Item;
        End If;
        --���
        --δ����ҽ��ҵ��ģ�����ʱ�͹Һŵ�,��Ҫ��֤ͬһ�εľ���ģʽ��һ����(�����Ѿ���飬�����ٴ���)
      End If;
      Update ������Ϣ Set ����ģʽ = ����ģʽ_In Where ����id = ����id_In;
    End If;
  
    Update ������Ϣ
    Set ������ = Null, ������ = Null, �������� = Null
    Where ����id = ����id_In And Exists (Select 1
           From ���˵�����¼
           Where ����id = ����id_In And ��ҳid Is Not Null And
                 �Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ���˵�����¼ Where ����id = ����id_In));
  
    If Sql%RowCount > 0 Then
      Update ���˵�����¼
      Set ����ʱ�� = Sysdate
      Where ����id = ����id_In And ��ҳid Is Not Null And Nvl(����ʱ��, Sysdate) > Sysdate;
    End If;
  End If;
  If ���_In = 1 Then
    --��Ϣ����
    Begin
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
        Using 1, n_�Һ�id;
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
End Zl_���˹Һż�¼_Insert;
/

Create Or Replace Procedure Zl_���˹Һż�¼_����_Delete
(
  ���ݺ�_In       ������ü�¼.No%Type,
  ����Ա���_In   ������ü�¼.����Ա���%Type,
  ����Ա����_In   ������ü�¼.����Ա����%Type,
  ժҪ_In         ������ü�¼.ժҪ%Type := Null, --ԤԼȡ��ʱ ��д ���ԤԼȡ��ԭ��
  ɾ�������_In   Number := 0,
  ��ԭ���˽���_In Varchar2 := Null,
  �˷�����_In     In Number := 0, --0-ȫ�� 1-�˹Һŷ� 2-�˲�����
  ��ָ������_In   ����Ԥ����¼.���㷽ʽ%Type := Null,
  �˺�����_In     Number := 1,
  ���㷽ʽ_In     Varchar2 := Null
) As
  --�˷�����_In,��һ�¼�������²�׼���в����˷�
  --    2.�����ӿ�,��ʱ��֧��
  -- �ҺŷѲ����ѷֿ���,����
  --    ��ͨ���㷽ʽ:ԭ���㷽ʽ�˲��ַ���
  --    Ԥ����:Ԥ����,�˲���
  --    Ԥ��������ͨ���㷽ʽ���:�˿����ͨ���㷽ʽ������
  --    ���ѿ�:ԭ�������ò����������ѿ�
  --��ԭ���˽���_In:ָ�����˻���ԭ�����㷽ʽ(��ҽ���ĸ����˻�,�����˻������ֵ�),����ö�����
  --��ָ������_IN:ָ��ԭ���˽��㲿��,Ӧ���˸����ֽ��㷽ʽ,Ϊ��ʱȱʡ�˸��ֽ�,�����˸�ָ���Ľ��㷽ʽ

  --���α������ж��Ƿ񵥶��ղ�����,���ҺŻ��ܱ���
  Cursor c_Registinfo(v_״̬ ������ü�¼.��¼״̬%Type) Is
    Select a.����ʱ��, a.�Ǽ�ʱ��, c.����ʱ��, a.�շ�ϸĿid As ��Ŀid, c.ִ�в���id As ����id, c.ִ���� As ҽ������, d.Id As ҽ��id, c.�ű� As ����
    From ������ü�¼ A, ���˹Һż�¼ C, ��Ա�� D
    Where a.��¼���� = 4 And a.��¼״̬ = v_״̬ And c.No = a.No And c.ִ���� = d.����(+) And a.No = ���ݺ�_In And
          Nvl(a.���㵥λ, '�ű�') = c.�ű� And Rownum < 2;
  r_Registrow c_Registinfo%RowType;

  --���α������жϼ�¼�Ƿ����,�����û��ܱ���
  Cursor c_Moneyinfo(v_״̬ ������ü�¼.��¼״̬%Type) Is
    Select ���˿���id, ��������id, ִ�в���id, ������Ŀid, Nvl(Sum(Ӧ�ս��), 0) As Ӧ��, Nvl(Sum(ʵ�ս��), 0) As ʵ��, Nvl(Sum(���ʽ��), 0) As ����
    From ������ü�¼
    Where ��¼���� = 4 And ��¼״̬ = v_״̬ And NO = ���ݺ�_In
    Group By ���˿���id, ��������id, ִ�в���id, ������Ŀid;
  r_Moneyrow c_Moneyinfo%RowType;

  --�ù�����ڴ�����Ա�ɿ�������˵Ĳ�ͬ���㷽ʽ�Ľ��
  Cursor c_Opermoney Is
    Select Distinct b.���㷽ʽ, -1 * Nvl(b.��Ԥ��, 0) As ��Ԥ��
    From ������ü�¼ A, ����Ԥ����¼ B
    Where a.����id = b.����id And a.No = ���ݺ�_In And a.��¼���� = 4 And a.��¼״̬ = 2 And b.��¼���� = 4 And b.��¼״̬ = 2 And
          Nvl(b.��Ԥ��, 0) <> 0 And
          Nvl(a.���ӱ�־, 0) =
          Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(a.���ӱ�־, 0), 1, -1, Nvl(a.���ӱ�־, 0)), Nvl(a.���ӱ�־, 0));

  v_Err_Msg Varchar(255);
  Err_Item Exception;

  n_����id ����Ԥ����¼.����id%Type;
  n_����id ������ü�¼.����id%Type;

  v_��ָ�����㷽ʽ ����Ԥ����¼.���㷽ʽ%Type;
  n_�˿���       ����Ԥ����¼.��Ԥ��%Type;
  n_��ӡid         Ʊ�ݴ�ӡ����.Id%Type;
  n_����id         ������Ϣ.����id%Type;
  n_�˷ѽ��       ����Ԥ����¼.��Ԥ��%Type;
  n_Ԥ�����       ����Ԥ����¼.��Ԥ��%Type; --ԭ��¼ Ԥ���ɿ���
  n_����ֵ         �������.Ԥ�����%Type;
  n_�Һ�id         ���˹Һż�¼.Id%Type;
  n_��id           ����ɿ����.Id%Type;

  n_�����˷�       Number; --��¼�Ƿ��Ǵ˵��ݵĵڶ����˷�
  n_����̨ǩ���Ŷ� Number;
  n_ԤԼ���ɶ���   Number;
  n_ԤԼ�Һ�       Number;
  n_�Һ����ɶ���   Number;
  d_Date           Date;
  n_����           ������ü�¼.���ʷ���%Type;
  n_����id1        ������Ϣ.����id%Type;
  n_���ض�         ������ü�¼.ʵ�ս��%Type;
  n_�ѽ���         Number;
  n_���           ���˹Һż�¼.����%Type;
  n_���ﲡ��id     ������Ϣ.����id%Type;
  d_����ʱ��       ����ǼǼ�¼.����ʱ��%Type;
  n_�����¼id     �ٴ������¼.Id%Type;
  v_��������       Varchar2(5000);
  v_��ǰ����       Varchar2(1000);
  v_���㷽ʽ       ����Ԥ����¼.���㷽ʽ%Type;
  n_��������־     Number;
  n_������       ����Ԥ����¼.��Ԥ��%Type;
Begin
  n_��id           := Zl_Get��id(����Ա����_In);
  v_��ָ�����㷽ʽ := ��ָ������_In;

  Select �����¼id, ���� Into n_�����¼id, n_��� From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum < 2;

  --�����ж�Ҫ�˺�/ȡ��ԤԼ�ļ�¼�Ƿ����
  Open c_Moneyinfo(1);
  Fetch c_Moneyinfo
    Into r_Moneyrow;
  If c_Moneyinfo%RowCount = 0 Then
    Close c_Moneyinfo;
    Open c_Moneyinfo(0);
    Fetch c_Moneyinfo
      Into r_Moneyrow;
    If c_Moneyinfo%RowCount = 0 Then
      v_Err_Msg := 'Ҫ����ĵ��ݲ����ڡ�';
      Raise Err_Item;
    End If;
    n_ԤԼ�Һ� := 1;
  End If;
  Close c_Moneyinfo;

  --1.ԤԼ����
  If Nvl(n_ԤԼ�Һ�, 0) = 1 Then
    --������Լ��
    Open c_Registinfo(0);
    Fetch c_Registinfo
      Into r_Registrow;
  
    Update �ٴ������¼ Set ��Լ�� = Nvl(��Լ��, 0) - 1 Where ID = n_�����¼id;
    Update ���˹ҺŻ���
    Set ��Լ�� = Nvl(��Լ��, 0) - 1
    Where ���� = Trunc(r_Registrow.����ʱ��) And Nvl(ҽ��id, 0) = Nvl(r_Registrow.ҽ��id, 0) And
          Nvl(ҽ������, '-') = Nvl(r_Registrow.ҽ������, '-') And ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And
          (���� = r_Registrow.���� Or ���� Is Null);
  
    If Sql%RowCount = 0 Then
      Insert Into ���˹ҺŻ���
        (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, ��Լ��)
      Values
        (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
         Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -1);
    End If;
  
    Close c_Registinfo;
  
    --���¹Һ����״̬
    Update �ٴ�������ſ���
    Set �Һ�״̬ = 0, ����Ա���� = Null
    Where �Һ�״̬ = 2 And ��¼id = n_�����¼id And ��� = n_���;
  
    Update �ٴ�������ſ���
    Set �Һ�״̬ = 4, ����Ա���� = Null
    Where �Һ�״̬ = 2 And ��¼id = n_�����¼id And ��ע = n_���;
  
    --��Ӳ��˹Һż�¼�� ������¼
    Select ���˹Һż�¼_Id.Nextval, Sysdate Into n_�Һ�id, d_Date From Dual;
  
    Update ���˹Һż�¼ Set ��¼״̬ = 3 Where NO = ���ݺ�_In And ��¼״̬ = 1 And ��¼���� = 2;
    If Sql%NotFound Then
      v_Err_Msg := 'ԤԼ����' || ���ݺ�_In || '�������ڻ����ڲ���ԭ���Ѿ���ȡ��ԤԼ';
      Raise Err_Item;
    End If;
  
    Insert Into ���˹Һż�¼
      (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
       ����, ����, ����, ԤԼ, ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ҽ�Ƹ��ʽ, �����¼id, ԤԼ����Ա, ԤԼ����Ա���)
      Select n_�Һ�id, NO, ��¼����, 2, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, d_Date, ����ʱ��,
             ����Ա���_In, ����Ա����_In, ����, ����, ����, ԤԼ, Nvl(ժҪ_In, ժҪ) As ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ҽ�Ƹ��ʽ,
             n_�����¼id, ԤԼ����Ա, ԤԼ����Ա���
      From ���˹Һż�¼
      Where NO = ���ݺ�_In;
  
    Update ������ü�¼ Set ��¼״̬ = 3 Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 0;
    Insert Into ������ü�¼
      (ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, ���ʵ�id, ����id, ҽ�����, �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��, ���ʽ, ���˿���id, �ѱ�, �շ����,
       �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id,
       ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���, ����Ա����, ����id, ���ʽ��, ���մ���id, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���, �ɿ���id, ����״̬, ��ת��,
       �Һ�id, ��ҳid)
      Select ���˷��ü�¼_Id.Nextval, ��¼����, NO, ʵ��Ʊ��, 2, ���, ��������, �۸񸸺�, ���ʵ�id, ����id, ҽ�����, �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��, ���ʽ,
             ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, -1 * ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, -1 * Ӧ�ս��,
             -1 * ʵ�ս��, ������, ��������id, ������, ����ʱ��, d_Date, ִ�в���id, ִ����, -1, ִ��ʱ��, ����, ����Ա���_In, ����Ա����_In, Null, Null,
             ���մ���id, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���, �ɿ���id, ����״̬, ��ת��, �Һ�id, ��ҳid
      From ������ü�¼
      Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 3;
  
    --���ԤԼ���ɶ���ʱ��Ҫ�������
  
    n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
    If Nvl(n_ԤԼ���ɶ���, 0) = 1 Then
      --Ҫɾ������
      For v_�Һ� In (Select ID, ����, ִ�в���id, ִ���� From ���˹Һż�¼ Where NO = ���ݺ�_In) Loop
        Zl_�ŶӽкŶ���_Delete(v_�Һ�.ִ�в���id, v_�Һ�.Id);
      End Loop;
    End If;
    Return;
  End If;

  Select Nvl(���ʷ���, 0), ����id, Decode(Sign(Nvl(����id, 0)), 0, 0, 1)
  Into n_����, n_����id, n_�ѽ���
  From ������ü�¼
  Where ��¼���� = 4 And NO = ���ݺ�_In And ��¼״̬ In (1, 3) And Rownum < 2;

  --2.�ҺŴ���
  n_�ѽ��� := Nvl(n_�ѽ���, 0);

  If n_�ѽ��� = 1 And n_���� = 1 Then
    Select Sysdate, Null Into d_Date, n_����id From Dual;
  Else
    Select Sysdate, ���˽��ʼ�¼_Id.Nextval Into d_Date, n_����id From Dual;
  End If;

  ----0-ȫ�� 1-�˹Һŷ� 2-�˲�����
  If Nvl(�˷�����_In, 0) <> 2 Then
    --���ǹ��˲�����ʱ����
    --���¹Һ����״̬
    If �˺�����_In = 1 Then
      Update �ٴ�������ſ���
      Set �Һ�״̬ = 0, ����Ա���� = Null
      Where �Һ�״̬ = 1 And ��¼id = n_�����¼id And ��� = n_���;
    
      Update �ٴ�������ſ���
      Set �Һ�״̬ = 4, ����Ա���� = Null
      Where �Һ�״̬ = 1 And ��¼id = n_�����¼id And ��ע = n_���;
    Else
      Update �ٴ�������ſ���
      Set �Һ�״̬ = 4, ����Ա���� = ����Ա����_In
      Where �Һ�״̬ = 1 And ��¼id = n_�����¼id And (��� = n_��� Or ��ע = n_���);
    End If;
  
    --���˾���״̬
    If n_����id Is Not Null Then
      Update ������Ϣ Set ����״̬ = 0, �������� = Null Where ����id = n_����id;
    
      --ɾ���������ش���,ֻ�е�ֻ��һ���Һż�¼���Ҳ��˽���������Һ����ڽ���ʱ�Żᴦ��
      If ɾ�������_In = 1 Then
        Delete ���ﲡ����¼ Where ����id = n_����id;
        Update ������Ϣ Set ����� = Null Where ����id = n_����id;
        --���ü�¼�����Һż����������￨����,�Լ����˽��Ѻ��˷ѻ����ʵķ���,�Һż�¼�������
        Update ������ü�¼ Set ��ʶ�� = Null Where �����־ = 1 And ����id = n_����id;
      End If;
    End If;
  
    --�����ʱ���˾��￨��,�˷�ʱ������￨��,�ڷǹ��˲�����ʱ
    n_����id1 := Null;
    Begin
      Select ����id
      Into n_����id1
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And ���ӱ�־ = 2 And Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
  
    If n_����id1 Is Not Null And Nvl(�˷�����_In, 0) <> 2 Then
      Update ������Ϣ
      Set ���￨�� = Null, ����֤�� = Null, Ic���� = Decode(Ic����, ���￨��, Null, Ic����)
      Where ����id = n_����id1;
    End If;
  
  End If;

  --���ǰ���Ƿ��Ѿ������˹�����
  Begin
    Select 1 Into n_�����˷� From ������ü�¼ Where ��¼���� = 4 And NO = ���ݺ�_In And ��¼״̬ = 3 And Rownum < 2;
  Exception
    When Others Then
      n_�����˷� := 0;
  End;

  --������ü�¼
  --������¼
  Insert Into ������ü�¼
    (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
     ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��,
     ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
    Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
           �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
           ����Ա���_In, ����Ա����_In, ����ʱ��, d_Date, n_����id,
           Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����,
           Nvl(ժҪ_In, ժҪ) As ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
    From ������ü�¼
    Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And
          Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));

  --ԭʼ��¼
  If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
    Update ������ü�¼
    Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
    Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And
          Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
  Else
    Update ������ü�¼
    Set ��¼״̬ = 3
    Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And
          Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
  End If;

  n_����id := 0;
  If n_���� = 0 Then
    --��ȡ����ID
    Select Nvl(����id, 0)
    Into n_����id
    From ������ü�¼
    Where ��¼���� = 4 And ��¼״̬ = 3 And NO = ���ݺ�_In And
          Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0)) And
          Rownum = 1;
  End If;

  If n_���� = 1 Then
    --����
    For c_���� In (Select ʵ�ս��, ���˿���id, ��������id, ִ�в���id, ������Ŀid
                 From ������ü�¼
                 Where ��¼���� = 4 And ��¼״̬ = 3 And NO = ���ݺ�_In And
                       Nvl(���ӱ�־, 0) =
                       Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0)) And
                       Nvl(���ʷ���, 0) = 1) Loop
      --�������
      Update �������
      Set ������� = Nvl(�������, 0) - Nvl(c_����.ʵ�ս��, 0)
      Where ����id = Nvl(n_����id, 0) And ���� = 1 And ���� = 1
      Returning ������� Into n_���ض�;
    
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, ����, �������, Ԥ�����)
        Values
          (n_����id, 1, 1, -1 * Nvl(c_����.ʵ�ս��, 0), 0);
        n_���ض� := Nvl(c_����.ʵ�ս��, 0);
      End If;
      If Nvl(n_���ض�, 0) = 0 Then
        Delete �������
        Where ����id = Nvl(n_����id, 0) And ���� = 1 And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
      --����δ�����
      Update ����δ�����
      Set ��� = Nvl(���, 0) - Nvl(c_����.ʵ�ս��, 0)
      Where ����id = n_����id And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(c_����.���˿���id, 0) And
            Nvl(��������id, 0) = Nvl(c_����.��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(c_����.ִ�в���id, 0) And ������Ŀid + 0 = c_����.������Ŀid And
            ��Դ;�� + 0 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into ����δ�����
          (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
        Values
          (n_����id, Null, Null, c_����.���˿���id, c_����.��������id, c_����.ִ�в���id, c_����.������Ŀid, 1, -1 * Nvl(c_����.ʵ�ս��, 0));
      End If;
    End Loop;
    Delete ����δ�����
    Where ����id = n_����id And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���, 0) = 0 And ��Դ;�� + 0 = 1;
  End If;

  If n_���� = 0 Then
    --1.�˷�
    --���˹ҺŽ���:�ֽ�͸����ʻ�����
    If ���㷽ʽ_In Is Null Then
      If ��ԭ���˽���_In Is Not Null Then
        --�˿����ȡ
        If Nvl(�˷�����_In, 0) <> 0 Or Nvl(n_�����˷�, 0) <> 0 Then
          --����ǵ����˲�����,����ֻ�˹Һŷ�,�Ȼ�ȡ�˷ѽ��
          Begin
            --��ȡ�����˿���
            Select Sum(Nvl(ʵ�ս��, 0)) As �տ���
            Into n_�˿���
            From ������ü�¼
            Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 3 And
                  Nvl(���ӱ�־, 0) =
                  Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
          
          Exception
            When Others Then
              v_Err_Msg := '���ݡ�' || ���ݺ�_In || '����' || Case Nvl(�˷�����_In, 0)
                             When 1 Then
                              '�Һŷ���'
                             When 2 Then
                              '������'
                           End || '�������ڲ���ԭ���Ѿ��������˷ѻ��ߵ��ݲ�����!';
              Raise Err_Item;
          End;
          Begin
            Select ��Ԥ��
            Into n_�˷ѽ��
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
          Exception
            When Others Then
              n_�˷ѽ�� := 0;
          End;
        
          --a.����Ľ��㷽ʽ
        
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In, -n_�˿���,
                   n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
        
          If n_�˷ѽ�� = 0 Then
            --b.����������ֽ�
            If n_�˿��� <> 0 Then
              If v_��ָ�����㷽ʽ Is Null Then
                --�˸��ֽ�
                Begin
                  Select ���� Into v_��ָ�����㷽ʽ From ���㷽ʽ Where ���� = 1;
                Exception
                  When Others Then
                    v_��ָ�����㷽ʽ := '�ֽ�';
                End;
              End If;
              Update ����Ԥ����¼
              Set ��Ԥ�� = ��Ԥ�� + (-1 * n_�˿���)
              Where ��¼���� = 4 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_��ָ�����㷽ʽ;
              If Sql%RowCount = 0 Then
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����,
                   �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
                  Select ����Ԥ����¼_Id.Nextval, a.No, a.ʵ��Ʊ��, a.��¼����, 2, a.����id, a.��ҳid, a.����id, a.ժҪ, v_��ָ�����㷽ʽ, d_Date,
                         ����Ա���_In, ����Ա����_In, -1 * n_�˿���, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
                  From ����Ԥ����¼ A
                  Where a.��¼���� = 4 And a.��¼״̬ = 1 And a.����id = n_����id And Rownum < 2;
              End If;
            End If;
          End If;
        Else
          --a.����Ľ��㷽ʽԭ����
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In, -��Ԥ��,
                   n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
        
          --b.����������ֽ�
          Begin
            Select Sum(��Ԥ��)
            Into n_�˷ѽ��
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') > 0;
          Exception
            When Others Then
              n_�˷ѽ�� := 0;
          End;
          If n_�˷ѽ�� <> 0 Then
            If v_��ָ�����㷽ʽ Is Null Then
              --�˸��ֽ�
              Begin
                Select ���㷽ʽ
                Into v_��ָ�����㷽ʽ
                From ����Ԥ����¼
                Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And
                      Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
              
              Exception
                When Others Then
                  Begin
                    Select ���� Into v_��ָ�����㷽ʽ From ���㷽ʽ Where ���� = 1;
                  Exception
                    When Others Then
                      v_��ָ�����㷽ʽ := '�ֽ�';
                  End;
              End;
            End If;
            Update ����Ԥ����¼
            Set ��Ԥ�� = ��Ԥ�� + (-1 * n_�˷ѽ��)
            Where ��¼���� = 4 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_��ָ�����㷽ʽ;
            If Sql%RowCount = 0 Then
              Insert Into ����Ԥ����¼
                (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����,
                 �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
                Select ����Ԥ����¼_Id.Nextval, a.No, a.ʵ��Ʊ��, a.��¼����, 2, a.����id, a.��ҳid, a.����id, a.ժҪ, v_��ָ�����㷽ʽ, d_Date,
                       ����Ա���_In, ����Ա����_In, -1 * n_�˷ѽ��, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
                From ����Ԥ����¼ A
                Where a.��¼���� = 4 And a.��¼״̬ = 1 And a.����id = n_����id And Rownum < 2;
            End If;
          End If;
        End If;
      Else
        --�˿����ȡ
        If Nvl(�˷�����_In, 0) <> 0 Or Nvl(n_�����˷�, 0) <> 0 Then
          --����ǵ����˲�����,����ֻ�˹Һŷ�,�Ȼ�ȡ�˷ѽ��
          Begin
            --��ȡ�����˿���
            Select Sum(Nvl(ʵ�ս��, 0)) As �տ���
            Into n_�˿���
            From ������ü�¼
            Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 3 And
                  Nvl(���ӱ�־, 0) =
                  Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
          Exception
            When Others Then
              v_Err_Msg := '���ݡ�' || ���ݺ�_In || '����' || Case Nvl(�˷�����_In, 0)
                             When 1 Then
                              '�Һŷ���'
                             When 2 Then
                              '������'
                           End || '�������ڲ���ԭ���Ѿ��������˷ѻ��ߵ��ݲ�����!';
              Raise Err_Item;
          End;
        End If;
        If Nvl(n_�����˷�, 0) = 0 And Nvl(�˷�����_In, 0) = 0 Then
          --�״�ȫ��
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                   -1 * ��Ԥ��, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id;
        Else
          --�����˷�,���߱��ε���һ����
          --�����˷�ʱ,��¼״̬=3 ,�״β�����,��¼״̬Ϊ1
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                   -1 * n_�˿���, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And ժҪ = 'ҽ���Һ�' And
                  ��Ԥ�� = n_�˿��� And Rownum < 2;
          If Sql%RowCount = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
               ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
              Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                     -1 * n_�˿���, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
              From ����Ԥ����¼
              Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And ��Ԥ�� = n_�˿��� And
                    Rownum < 2;
          End If;
          If Sql%RowCount = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
               ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
              Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                     -1 * n_�˿���, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
              From ����Ԥ����¼
              Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And Rownum < 2;
            If Sql%RowCount = 0 Then
              --�����˷�,����ȫ��ʹ��Ԥ����ɷ�ʱ�Ŵ��ڴ������
              n_Ԥ����� := n_�˿���;
            End If;
          End If;
        
        End If;
      End If;
    Else
      --�����㷽ʽ��
      v_�������� := ���㷽ʽ_In || '|'; --�Կո�ֿ���|��β,û�н�������
      While v_�������� Is Not Null Loop
        v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '|') - 1);
        v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
      
        v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
        n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
      
        v_��ǰ����   := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
        n_��������־ := To_Number(v_��ǰ����);
      
        If n_��������־ = 0 Then
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������, �������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, v_���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                   -1 * n_������, n_����id, n_��id, Ԥ�����, Null, Null, Null, Null, Null, ������λ, 4, �������
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And Rownum < 2;
        Else
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������, �������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, v_���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                   -1 * n_������, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4, �������
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And
                  (�����id Is Not Null Or ���㿨��� Is Not Null) And Rownum < 2;
        End If;
        v_�������� := Substr(v_��������, Instr(v_��������, '|') + 1);
      End Loop;
    End If;
    --�״��˷�ʱ,��¼״̬�����Ϊ��3
    Update ����Ԥ����¼
    Set ��¼״̬ = 3
    Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id;
  
    --��Ԥ�� 1-ȫ�� 2-������,������ʱ��ȫ��ʹ��Ԥ�����нɿ�
    If Nvl(�˷�����_In, 0) = 0 Or (Nvl(�˷�����_In, 0) <> 0 And n_Ԥ����� <> 0) Then
      --���˹ҺŽ���:��Ԥ�����
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
        Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, d_Date,
               ����Ա����_In, ����Ա���_In, -1 * Decode(Nvl(�˷�����_In, 0) + Nvl(n_�����˷�, 0), 0, ��Ԥ��, n_Ԥ�����), n_����id, n_��id, Ԥ�����,
               �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
        From ����Ԥ����¼
        Where ��¼���� In (1, 11) And ����id = n_����id And Nvl(��Ԥ��, 0) <> 0 And
              Rownum = Decode(Nvl(�˷�����_In, 0) + Nvl(n_�����˷�, 0), 0, Rownum, 1);
    End If;
  
    --������Ԥ�����
    For c_Ԥ�� In (Select ����id, Ԥ�����, -1 * Sum(Nvl(��Ԥ��, 0)) As ��Ԥ��
                 From ����Ԥ����¼
                 Where ��¼���� In (1, 11) And ����id = n_����id
                 Group By ����id, Ԥ�����) Loop
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(c_Ԥ��.��Ԥ��, 0)
      Where ����id = c_Ԥ��.����id And ���� = Nvl(c_Ԥ��.Ԥ�����, 2) And ���� = 1
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, Ԥ�����, ����, ����)
        Values
          (c_Ԥ��.����id, Nvl(c_Ԥ��.��Ԥ��, 0), 1, Nvl(c_Ԥ��.Ԥ�����, 2));
        n_����ֵ := Nvl(c_Ԥ��.��Ԥ��, 0);
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From �������
        Where ����id = c_Ԥ��.����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
      End If;
    End Loop;
  
    If Nvl(�˷�����_In, 0) <> 2 Then
      --���˹Һŷ�,������Ʊ��
      --�˿��ջ�Ʊ��(�����ϴιҺ�ʹ��Ʊ��,�����ջ�)
      Begin
        --�����һ�εĴ�ӡ������ȡ
        Select ID
        Into n_��ӡid
        From (Select b.Id
               From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
               Where a.��ӡid = b.Id And a.���� = 1 And b.�������� = 4 And b.No = ���ݺ�_In
               Order By a.ʹ��ʱ�� Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          Null;
      End;
    
      If n_��ӡid Is Not Null Then
        Insert Into Ʊ��ʹ����ϸ
          (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
          Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, ����Ա����_In
          From Ʊ��ʹ����ϸ
          Where ��ӡid = n_��ӡid And ���� = 1;
      End If;
    End If;
  End If;

  --�����˲�������,��������ܼ�¼
  --��ػ��ܱ�Ĵ���

  --���˹ҺŻ���
  Open c_Registinfo(3);
  Fetch c_Registinfo
    Into r_Registrow;

  If c_Registinfo%RowCount = 0 Then
    --ֻ�ղ�����ʱ�޺ű�,������
    Close c_Registinfo;
  Else
  
    --��Ҫȷ���Ƿ�ԤԼ�Һ�
    --1.�������ԤԼ�ҺŲ����ĹҺż�¼,����Ҫ���ѹ����������ѽ���
    --2.����������Һ�,��ֻ���ѹ���
  
    Begin
      Select Decode(ԤԼ, Null, 0, 0, 0, 1) Into n_ԤԼ�Һ� From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum = 1;
    Exception
      When Others Then
        n_ԤԼ�Һ� := 0;
    End;
  
    Update �ٴ������¼
    Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - n_ԤԼ�Һ�, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ�Һ�
    Where ID = n_�����¼id;
  
    Update ���˹ҺŻ���
    Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - n_ԤԼ�Һ�, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ�Һ�
    Where ���� = Trunc(r_Registrow.����ʱ��) And Nvl(ҽ��id, 0) = Nvl(r_Registrow.ҽ��id, 0) And
          Nvl(ҽ������, '-') = Nvl(r_Registrow.ҽ������, '-') And ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And
          (���� = r_Registrow.���� Or ���� Is Null);
  
    If Sql%RowCount = 0 Then
      Insert Into ���˹ҺŻ���
        (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, �����ѽ���, ��Լ��)
      Values
        (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
         Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -1, -1 * n_ԤԼ�Һ�, -1 * n_ԤԼ�Һ�);
    End If;
  
    Close c_Registinfo;
  End If;

  If n_���� = 0 Then
    --��Ա�ɿ����(���������ʻ��ȵĽ�����,�����˳�Ԥ����)
    For r_Opermoney In c_Opermoney Loop
      Update ��Ա�ɿ����
      Set ��� = Nvl(���, 0) + (-1 * r_Opermoney.��Ԥ��)
      Where �տ�Ա = ����Ա����_In And ���㷽ʽ = r_Opermoney.���㷽ʽ And ���� = 1
      Returning ��� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ��Ա�ɿ����
          (�տ�Ա, ���㷽ʽ, ����, ���)
        Values
          (����Ա����_In, r_Opermoney.���㷽ʽ, 1, -1 * r_Opermoney.��Ԥ��);
        n_����ֵ := r_Opermoney.��Ԥ��;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From ��Ա�ɿ����
        Where �տ�Ա = ����Ա����_In And ���㷽ʽ = r_Opermoney.���㷽ʽ And ���� = 1 And Nvl(���, 0) = 0;
      End If;
    End Loop;
  End If;
  If Nvl(�˷�����_In, 0) <> 2 Then
    n_�Һ����ɶ��� := Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113));
    If n_�Һ����ɶ��� <> 0 Then
      n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113));
      If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Then
      
        --Ҫɾ������
        For v_�Һ� In (Select ID, ����, ִ�в���id, ִ���� From ���˹Һż�¼ Where NO = ���ݺ�_In) Loop
          Zl_�ŶӽкŶ���_Delete(v_�Һ�.ִ�в���id, v_�Һ�.Id);
        End Loop;
      End If;
    End If;
  
    --ҽ�������ľ���ǼǼ�¼
    Begin
      Select ����id, ����ʱ�� Into n_���ﲡ��id, d_����ʱ�� From ���˹Һż�¼ Where NO = ���ݺ�_In;
      Delete From ����ǼǼ�¼ Where ����id = n_���ﲡ��id And ����ʱ�� = d_����ʱ�� And ��ҳid Is Null;
    Exception
      When Others Then
        Null;
    End;
  
    --���˹Һż�¼
    Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
  
    Update ���˹Һż�¼ Set ��¼״̬ = 3 Where NO = ���ݺ�_In And ��¼���� = 1 And ��¼״̬ = 1;
    If Sql%NotFound Then
      v_Err_Msg := '�Һŵ���' || ���ݺ�_In || '�������ڻ����ڲ���ԭ���Ѿ����˺�';
      Raise Err_Item;
    End If;
    Insert Into ���˹Һż�¼
      (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
       ����, ����, ����, ԤԼ, ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ����, ҽ�Ƹ��ʽ, �����¼id)
      Select n_�Һ�id, NO, ��¼����, 2, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, d_Date, ����ʱ��,
             ����Ա���_In, ����Ա����_In, ����, ����, ����, ԤԼ, Nvl(ժҪ_In, ժҪ) As ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ����, ҽ�Ƹ��ʽ,
             n_�����¼id
      From ���˹Һż�¼
      Where NO = ���ݺ�_In;
  End If;
  --��Ϣ����
  Begin
    Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
      Using 2, ���ݺ�_In;
  Exception
    When Others Then
      Null;
  End;
  b_Message.Zlhis_Regist_003(n_�Һ�id, ���ݺ�_In);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˹Һż�¼_����_Delete;
/

Create Or Replace Procedure Zl_���˹Һż�¼_����_Insert
(
  �����¼id_In    �ٴ������¼.Id%Type,
  ����id_In        ������ü�¼.����id%Type,
  �����_In        ������ü�¼.��ʶ��%Type,
  ����_In          ������ü�¼.����%Type,
  �Ա�_In          ������ü�¼.�Ա�%Type,
  ����_In          ������ü�¼.����%Type,
  ���ʽ_In      ������ü�¼.���ʽ%Type, --���ڴ�Ų��˵�ҽ�Ƹ��ʽ���
  �ѱ�_In          ������ü�¼.�ѱ�%Type,
  ���ݺ�_In        ������ü�¼.No%Type,
  Ʊ�ݺ�_In        ������ü�¼.ʵ��Ʊ��%Type,
  ���_In          ������ü�¼.���%Type,
  �۸񸸺�_In      ������ü�¼.�۸񸸺�%Type,
  ��������_In      ������ü�¼.��������%Type,
  �շ����_In      ������ü�¼.�շ����%Type,
  �շ�ϸĿid_In    ������ü�¼.�շ�ϸĿid%Type,
  ����_In          ������ü�¼.����%Type,
  ��׼����_In      ������ü�¼.��׼����%Type,
  ������Ŀid_In    ������ü�¼.������Ŀid%Type,
  �վݷ�Ŀ_In      ������ü�¼.�վݷ�Ŀ%Type,
  ���㷽ʽ_In      Varchar2,
  Ӧ�ս��_In      ������ü�¼.Ӧ�ս��%Type,
  ʵ�ս��_In      ������ü�¼.ʵ�ս��%Type,
  ���˿���id_In    ������ü�¼.���˿���id%Type,
  ��������id_In    ������ü�¼.��������id%Type,
  ִ�в���id_In    ������ü�¼.ִ�в���id%Type,
  ����Ա���_In    ������ü�¼.����Ա���%Type,
  ����Ա����_In    ������ü�¼.����Ա����%Type,
  ����ʱ��_In      ������ü�¼.����ʱ��%Type,
  �Ǽ�ʱ��_In      ������ü�¼.�Ǽ�ʱ��%Type,
  ҽ������_In      �ҺŰ���.ҽ������%Type,
  ҽ��id_In        �ҺŰ���.ҽ��id%Type,
  ������_In        Number, --������¼�Ƿ���������
  ����_In          Number,
  �ű�_In          �ҺŰ���.����%Type,
  ����_In          ������ü�¼.��ҩ����%Type,
  ����id_In        ������ü�¼.����id%Type,
  ����id_In        Ʊ��ʹ����ϸ.����id%Type,
  Ԥ��֧��_In      ����Ԥ����¼.��Ԥ��%Type, --ˢ���Һ�ʱʹ�õ�Ԥ�����,���Ϊ1����.
  �ֽ�֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�ֽ�֧�����ݽ��,���Ϊ1����.
  ����֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�����ʻ�֧�����,,���Ϊ1����.
  ���մ���id_In    ������ü�¼.���մ���id%Type,
  ������Ŀ��_In    ������ü�¼.������Ŀ��%Type,
  ͳ����_In      ������ü�¼.ͳ����%Type,
  ժҪ_In          ������ü�¼.ժҪ%Type, --ԤԼ�Һ�ժҪ��Ϣ
  ԤԼ�Һ�_In      Number := 0, --ԤԼ�Һ�ʱ��(��¼״̬=0,����ʱ��ΪԤԼʱ��),��ʱ����Ҫ���������ز���
  �շ�Ʊ��_In      Number := 0, --�Һ��Ƿ�ʹ���շ�Ʊ��
  ���ձ���_In      ������ü�¼.���ձ���%Type,
  ����_In          ���˹Һż�¼.����%Type := 0,
  ����_In          �Һ����״̬.���%Type := Null, --ԤԼʱ������ü�¼�ķ�ҩ�����ֶ�,�Һ�ʱ����Һż�¼
  ����_In          ���˹Һż�¼.����%Type := Null,
  ԤԼ����_In      Number := 0,
  ԤԼ��ʽ_In      ԤԼ��ʽ.����%Type := Null,
  ���ɶ���_In      Number := 0,
  �����id_In      ����Ԥ����¼.�����id%Type := Null,
  ���㿨���_In    ����Ԥ����¼.���㿨���%Type := Null,
  ����_In          ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In      ����Ԥ����¼.����˵��%Type := Null,
  ������λ_In      ����Ԥ����¼.������λ%Type := Null,
  ��������_In      Number := 0,
  ����_In          ���˹Һż�¼.����%Type := Null,
  ����ģʽ_In      Number := 0,
  ���ʷ���_In      Number := 0,
  �˺�����_In      Number := 1,
  ��Ԥ������ids_In Varchar2 := Null,
  �������˷ѱ�_In  Number := 0,
  ԤԼ˳���_In    �ٴ�������ſ���.ԤԼ˳���%Type := Null
) As
  ---------------------------------------------------------------------------
  --
  --����:
  --     ��������_in:0-�����ҺŻ���ԤԼ 1-����Աӵ�мӺ�Ȩ�޼Ӻ�
  --     �������˷ѱ�_In:0-���޸Ĳ��˷ѱ� 1-�޸Ĳ��˷ѱ�
  ----------------------------------------------------------------------------
  --���α������շѳ�Ԥ���Ŀ���Ԥ���б�
  --��ID�������ȳ��ϴ�δ����ġ�
  Cursor c_Deposit
  (
    v_����id        ������Ϣ.����id%Type,
    v_��Ԥ������ids Varchar2
  ) Is
    Select *
    From (Select a.Id, a.����id, a.��¼״̬, a.Ԥ�����, a.No, Nvl(a.���, 0) As ���
           From ����Ԥ����¼ A,
                (Select NO, Sum(Nvl(a.���, 0)) As ���
                  From ����Ԥ����¼ A
                  Where a.����id Is Null And Nvl(a.���, 0) <> 0 And
                        a.����id In (Select Column_Value From Table(f_Num2list(v_��Ԥ������ids))) And Nvl(a.Ԥ�����, 2) = 1
                  Group By NO
                  Having Sum(Nvl(a.���, 0)) <> 0) B
           Where a.����id Is Null And Nvl(a.���, 0) <> 0 And a.���㷽ʽ Not In (Select ���� From ���㷽ʽ Where ���� = 5) And
                 a.No = b.No And a.����id In (Select Column_Value From Table(f_Num2list(v_��Ԥ������ids))) And
                 Nvl(a.Ԥ�����, 2) = 1
           Union All
           Select 0 As ID, Max(����id) As ����id, ��¼״̬, Ԥ�����, NO, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���
           From ����Ԥ����¼
           Where ��¼���� In (1, 11) And ����id Is Not Null And Nvl(���, 0) <> Nvl(��Ԥ��, 0) And
                 ����id In (Select Column_Value From Table(f_Num2list(v_��Ԥ������ids))) And Nvl(Ԥ�����, 2) = 1 Having
            Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
           Group By ��¼״̬, NO, Ԥ�����)
    Order By Decode(����id, Nvl(v_����id, 0), 0, 1), ID, NO;
  --���ܣ�����һ�в��˹Һŷ��ã�����������ܵ�����Ԥ����¼
  --       ͬʱ������صĻ��ܱ�(���˹ҺŻ��ܡ����û���)
  --       ��һ�з��ô���Ʊ��ʹ�����(����ID_IN>0)

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_�ŶӺ��� �ŶӽкŶ���.�ŶӺ���%Type;
  v_�ֽ�     ���㷽ʽ.����%Type;
  v_�����ʻ� ���㷽ʽ.����%Type;
  v_�������� �ŶӽкŶ���.��������%Type;

  n_��ʱ��       Number;
  n_ʱ���޺�     Number;
  n_ʱ����Լ     Number;
  d_ʱ��ʱ��     Date;
  d_������ʱ�� Date;
  n_׷�Ӻ�       Number := 0; --����ʱ�ι��� ׷�ӹҺŵ����
  n_��Լ��       ���˹ҺŻ���.��Լ��%Type;
  n_�ѽ���       ���˹ҺŻ���.�����ѽ���%Type;
  n_ԤԼ��Чʱ�� Number;
  n_ʧЧ��       Number;
  n_ʧԼ�Һ�     Number := 0;
  n_��������     Number;
  n_����         Number := 0;

  n_��ӡid        Ʊ�ݴ�ӡ����.Id%Type;
  n_����id        ������ü�¼.Id%Type;
  n_Ԥ�����      ����Ԥ����¼.���%Type;
  n_��ǰ���      ����Ԥ����¼.���%Type;
  n_����ֵ        ����Ԥ����¼.���%Type;
  n_Ԥ��id        ����Ԥ����¼.Id%Type;
  n_���ѿ�id      ���ѿ�Ŀ¼.Id%Type;
  n_�Һ�id        ���˹Һż�¼.Id%Type;
  v_��Ԥ������ids Varchar2(4000);

  n_��id           ����ɿ����.Id%Type;
  n_�����         ������Ϣ.�����%Type;
  n_���           �Һ����״̬.���%Type;
  n_�������       �Һ����״̬.���%Type;
  n_��ſ���       �ҺŰ���.��ſ���%Type;
  n_����̨ǩ���Ŷ� Number;
  n_Count          Number;
  n_�޺���         Number(18);
  n_���ƿ�         Number;
  d_�Ŷ�ʱ��       Date;
  v_���㷽ʽ��¼   Varchar2(1000);
  d_���ʱ��       Date;
  v_�ѱ�           �ѱ�.����%Type;
  n_ʱ�����       Number := -1;
  n_ԤԼ���ɶ���   Number;
  v_���㷽ʽ       ���㷽ʽ.����%Type;
  v_��������       Varchar2(1000);
  v_��ǰ����       Varchar2(200);
  v_�������       ����Ԥ����¼.�������%Type;
  n_������       ����Ԥ����¼.��Ԥ��%Type;
  n_��������־     Number(2);
  n_����id         �ҺŰ���.Id%Type;
  n_ԤԼ˳���     �ٴ�������ſ���.ԤԼ˳���%Type;
  n_�ƻ�id         �ҺŰ��żƻ�.Id%Type := 0;
  v_����           �ҺŰ�������.������Ŀ%Type;
  n_��Լ��         Number(18);
  n_�ѹ���         Number(4) := 0;

  n_�ҳ��������� Number(4) := 0;
  n_����ģʽ       ������Ϣ.����ģʽ%Type;
  v_�Ŷ����       �ŶӽкŶ���.�Ŷ����%Type;
  v_������         �Һ����״̬.������%Type;
  v_��Ų���Ա     �Һ����״̬.����Ա����%Type;
  v_��Ż�����     �Һ����״̬.������%Type;
  v_���ʽ       ���˹Һż�¼.ҽ�Ƹ��ʽ%Type;
  n_״̬           �ٴ�������ſ���.�Һ�״̬%Type;
Begin
  --��ȡ��ǰ��������
  Select Terminal Into v_������ From V$session Where Audsid = Userenv('sessionid');
  n_��id          := Zl_Get��id(����Ա����_In);
  v_��Ԥ������ids := Nvl(��Ԥ������ids_In, ����id_In);

  If �ѱ�_In Is Null Then
    Begin
      Select ���� Into v_�ѱ� From �ѱ� Where ȱʡ��־ = 1 And Rownum < 2;
      Update ������Ϣ Set �ѱ� = v_�ѱ� Where ����id = ����id_In;
    Exception
      When Others Then
        v_Err_Msg := '�޷�ȷ�����˷ѱ�����ȱʡ�ѱ��Ƿ���ȷ���ã�';
        Raise Err_Item;
    End;
  Else
    v_�ѱ� := �ѱ�_In;
    If Nvl(�������˷ѱ�_In, 0) = 1 Then
      Begin
        Update ������Ϣ Set �ѱ� = v_�ѱ� Where ����id = ����id_In;
      Exception
        When Others Then
          v_Err_Msg := 'û���ҵ���Ӧ�Ĳ��ˣ�';
          Raise Err_Item;
      End;
    End If;
  End If;

  If �����_In Is Not Null Then
    Begin
      Select Nvl(�����, 0) Into n_����� From ������Ϣ Where ����id = ����id_In;
    Exception
      When Others Then
        n_����� := 0;
    End;
    If n_����� = 0 Then
      Update ������Ϣ Set ����� = �����_In Where ����id = ����id_In;
    End If;
  End If;

  Begin
    Update �ٴ�������ſ���
    Set �Һ�״̬ = 0
    Where ��¼id = �����¼id_In And ��� = ����_In And Nvl(�Һ�״̬, 0) = 3 And ����Ա���� = ����Ա����_In;
  Exception
    When Others Then
      Null;
  End;

  If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
    --�ҺŻ���ԤԼ����
    --��Ϊ�����а��ձ൥�ݺŹ���,�չҺ������ܳ���10000��,����Ҫ���ΨһԼ����
    Select Count(*)
    Into n_Count
    From ������ü�¼
    Where ��¼���� = 4 And ��¼״̬ In (1, 3) And ��� = ���_In And NO = ���ݺ�_In;
    If n_Count <> 0 Then
      v_Err_Msg := '�Һŵ��ݺ��ظ�,���ܱ��棡' || Chr(13) || '���ʹ���˰���˳����,���չҺ������ܳ���10000�˴Ρ�';
      Raise Err_Item;
    End If;
  
    --��ȡ���㷽ʽ����
    Begin
      Select ���� Into v_�ֽ� From ���㷽ʽ Where ���� = 1;
    Exception
      When Others Then
        v_�ֽ� := '�ֽ�';
    End;
    Begin
      Select ���� Into v_�����ʻ� From ���㷽ʽ Where ���� = 3;
    Exception
      When Others Then
        v_�����ʻ� := '�����ʻ�';
    End;
  End If;

  n_��� := ����_In;

  --��ȡ�Ƿ��ʱ��
  Begin
    Select Nvl(�Ƿ��ʱ��, 0), Nvl(�Ƿ���ſ���, 0), �޺���, ��Լ��
    Into n_��ʱ��, n_��ſ���, n_�޺���, n_��Լ��
    From �ٴ������¼
    Where ID = �����¼id_In;
  Exception
    When Others Then
      n_��ʱ��   := 0;
      n_��ſ��� := 0;
      n_�޺���   := Null;
      n_��Լ��   := Null;
  End;

  If n_��� Is Null And n_��ʱ�� = 1 And n_��ſ��� = 0 Then
    Begin
      Select ��� Into n_��� From �ٴ�������ſ��� Where ��¼id = �����¼id_In And ��ʼʱ�� = ����ʱ��_In;
    Exception
      When Others Then
        n_��� := Null;
    End;
  End If;

  --��ʱ�ιҺ�ʱ�ж��Ƿ��ǹ��ڹҺ� Ҳ����׷�Ӻŵ���� Ŀǰֻ���ר�Һŷ�ʱ�ν��д���
  If Nvl(ԤԼ�Һ�_In, 0) = 0 And n_��ʱ�� > 0 And Nvl(n_��ſ���, 0) = 1 And ����_In Is Null And Nvl(��������_In, 0) = 0 Then
    Begin
      Select Max(To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'))
      Into d_������ʱ��
      From �ٴ�������ſ���
      Where ��¼id = �����¼id_In And Nvl(����, 0) <> 0;
    
      n_׷�Ӻ� := Case Sign(����ʱ��_In - d_������ʱ��)
                 When -1 Then
                  0
                 Else
                  1
               End;
    Exception
      When Others Then
        n_׷�Ӻ� := 0;
    End;
  End If;
  d_ʱ��ʱ�� := ����ʱ��_In;

  If ���_In = 1 And n_��ʱ�� > 0 Then
    If Nvl(n_��ſ���, 0) = 1 Then
      Begin
        Select Nvl(���, 0),
               To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
               ����, Decode(Nvl(�Ƿ�ԤԼ, 0), 0, 0, ����)
        Into n_ʱ�����, d_ʱ��ʱ��, n_ʱ���޺�, n_ʱ����Լ
        From �ٴ�������ſ���
        Where ��¼id = �����¼id_In And ��� = n_���;
      Exception
        When Others Then
          n_ʱ����� := -1;
          n_��ʱ��   := 0;
          d_ʱ��ʱ�� := ����ʱ��_In;
          n_ʱ���޺� := 0;
          n_ʱ����Լ := 0;
      End;
    Else
      --�Һ�ʱ��� �Ƿ����ʱ��,����ʱ��,��ʱ�ε�����������ȡ����
      Begin
        Select Nvl(���, 0),
               To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
               ����, Decode(Nvl(�Ƿ�ԤԼ, 0), 0, 0, ����)
        Into n_ʱ�����, d_ʱ��ʱ��, n_ʱ���޺�, n_ʱ����Լ
        From �ٴ�������ſ���
        Where ��¼id = �����¼id_In And ��� = n_��� And ԤԼ˳��� Is Null;
      Exception
        When Others Then
          n_ʱ����� := -1;
          n_��ʱ��   := 0;
          d_ʱ��ʱ�� := ����ʱ��_In;
          n_ʱ���޺� := 0;
          n_ʱ����Լ := 0;
      End;
    End If;
  End If;

  If ���_In = 1 Then
    --��ȡ��ǰδʹ�õ����
    If Nvl(n_��ſ���, 0) = 1 And n_��ʱ�� = 0 Then
      --<��ſ��� δ����ʱ�� ��ȡ���õ�������,�Լ��Ѿ�ʹ�õ�����>
      Begin
        --������
        Select Count(1) Into n_�������� From ���˹Һż�¼ Where �����¼id = �����¼id_In And ��¼״̬ = 1;
        Select Max(���) Into n_������� From �ٴ�������ſ��� Where ��¼id = �����¼id_In;
      Exception
        When Others Then
          n_������� := 0;
          n_�������� := 0;
      End;
      Begin
        --������
        Select Sum(Nvl(����, 0))
        
        Into n_��Լ��
        From �ٴ�������ſ���
        Where ��¼id = �����¼id_In And Nvl(�Һ�״̬, 0) = 2;
      Exception
        When Others Then
          n_��Լ�� := 0;
      End;
    
      If n_��� Is Null Then
        n_��� := Nvl(n_�������, 0) + 1;
      End If;
      --<��ſ��� δ����ʱ�� ��ȡ���õ������� �Լ��Ѿ�ʹ�õ����� --end>
    
      --�ǼӺŵ������Ҫ����Ƿ񳬹�������
      If ��������_In = 0 Then
        If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
          --�Һ�ʱ���
          --������ſ���δ��ʱ�� �ﵽ������
          If n_�޺��� <= n_�������� And n_�޺��� > 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(Trunc(����ʱ��_In), 'yyyy-mm-dd ') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        
        Else
          --ԤԼʱ���
          If n_��Լ�� = 0 Then
            n_��Լ�� := n_�޺���;
          End If;
        
          If n_��Լ�� <= n_��Լ�� And n_��Լ�� > 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(Trunc(����ʱ��_In), 'yyyy-mm-dd') || '�Ѵﵽ�����Լ����';
            Raise Err_Item;
          End If;
        End If;
      
      Else
        Null;
        --������ſ���,δ��ʱ�� �Ӻ����   ������,����Ժ������������Ժ󲹳�
      End If;
    
    Elsif Nvl(n_��ʱ��, 0) > 0 And Nvl(n_��ſ���, 0) = 0 Then
      --<--��ͨ�ŷ�ʱ�� ����ֻ��ԤԼһ�����-->
      If ��������_In = 0 Then
        --<����ԤԼ�Һ�-->
        Begin
          Select Count(0) As �ѹ���, Nvl(Sum(Decode(Nvl(Sign(a.��ʼʱ�� - d_ʱ��ʱ��), 0), 0, 1, 0)), 0) As ��Լ��
          Into n_�ѹ���, n_��Լ��
          From �ٴ�������ſ��� A
          Where ��¼id = �����¼id_In And �Һ�״̬ Not In (0, 4, 5);
        Exception
          When Others Then
            n_�ѹ��� := 0;
            n_��Լ�� := 0;
        End;
      
        n_ʱ����Լ := n_ʱ���޺�; --��ͨ�ŷ�ʱ�ε����,29��n_ʱ����Լʼ����0 �������⴦��
        --�����������
        If n_��Լ�� = 0 Then
          n_��Լ�� := n_�޺���;
        End If;
        If n_ʱ����Լ <= n_��Լ�� Or n_��Լ�� <= n_��Լ�� Then
          v_Err_Msg := '�ű�' || �ű�_In || '��ʱ��' || To_Char(d_ʱ��ʱ��, 'hh24:mi') || '�Ѵﵽ�����Լ����';
          Raise Err_Item;
        End If;
        If n_�޺��� <= n_�ѹ��� Then
          v_Err_Msg := '�ű�' || �ű�_In || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
          Raise Err_Item;
        End If;
      End If;
    
      --û�дﵽʱ�ε��޺��� �����ڵ�ǰʱ������׷��
    
      --��ȡ����ҳ���������
      Select Nvl(Max(���), 0)
      Into n_�ҳ���������
      From �ٴ�������ſ��� A
      Where ��¼id = �����¼id_In And ԤԼ˳��� Is Null And �Һ�״̬ Not In (0, 5);
      If ԤԼ˳���_In Is Not Null Then
        n_ԤԼ˳��� := ԤԼ˳���_In;
      Else
        Begin
          Select Nvl(Max(ԤԼ˳���), 0) + 1
          Into n_ԤԼ˳���
          From �ٴ�������ſ���
          Where ��¼id = �����¼id_In And ��� = n_ʱ����� And ԤԼ˳��� Is Not Null;
        Exception
          When Others Then
            n_ԤԼ˳��� := Null;
        End;
      End If;
      --�������
      n_��� := RPad(Nvl(n_ʱ�����, 0), Length(n_�޺���) + Length(Nvl(n_ʱ�����, 0)), 0) + n_ԤԼ˳���;
      If n_ԤԼ˳��� Is Null Then
        n_��� := Nvl(n_�ҳ���������, 0) + 1;
      End If;
    
      --<--��ͨ�ŷ�ʱ��--End>
    Elsif Nvl(n_��ʱ��, 0) > 0 And Nvl(n_��ſ���, 0) = 1 Then
      --<������ſ��� ����ʱ��
      --ר�Һŷ�ʱ��
      Begin
        Select Max(���), Sum(Decode(���, Nvl(����_In, 0), 0, 1)), Sum(Decode(Sign(��ʼʱ�� - d_ʱ��ʱ��), 0, 1, 0))
        Into n_�������, n_�ѹ���, n_��������
        From �ٴ�������ſ���
        Where ��¼id = �����¼id_In And �Һ�״̬ Not In (0, 4, 5);
      Exception
        When Others Then
          n_������� := 0;
          n_�������� := 0;
          n_�ѹ���   := 0;
      End;
    
      n_ʧЧ�� := 0;
      If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
        n_ԤԼ��Чʱ�� := Zl_To_Number(zl_GetSysParameter('ԤԼ��Чʱ��', 1111));
        n_ʧԼ�Һ�     := Zl_To_Number(zl_GetSysParameter('ʧԼ���ڹҺ�', 1111));
        If Nvl(n_ԤԼ��Чʱ��, 0) <> 0 And Nvl(n_ʧԼ�Һ�, 0) > 0 Then
          Begin
            Select Sum(Decode(Sign((Sysdate - n_ԤԼ��Чʱ�� / 24 / 60) - ��ʼʱ��), 1, 1, 0))
            Into n_ʧЧ��
            From �ٴ�������ſ���
            Where ��¼id = �����¼id_In And ��ʼʱ�� Between Trunc(Sysdate) And Sysdate And Nvl(�Һ�״̬, 0) = 2;
          Exception
            When Others Then
              n_ʧЧ�� := 0;
          End;
        End If;
      End If;
    
      If ��������_In = 0 Then
        If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
        
          --�Һ� ׷�Ӻ���ʱ�����ʱ���޺���
          If n_ʱ���޺� <= n_�������� And Nvl(n_׷�Ӻ�, 0) = 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��ʱ��' || To_Char(d_ʱ��ʱ��, 'hh24:mi') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
          If n_�޺��� <= n_�ѹ��� - n_ʧЧ�� Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        
        Else
          --�Һ�
          If n_��Լ�� = 0 Then
            n_��Լ�� := n_�޺���;
          End If;
          If n_��Լ�� <= n_�������� Then
            v_Err_Msg := '�ű�' || �ű�_In || '��ʱ��' || To_Char(d_ʱ��ʱ��, 'hh24:mi') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
          If n_�޺��� <= n_�ѹ��� Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        End If;
      End If;
      If n_��� Is Null Then
        --�������
        If Nvl(n_�������, 0) < Nvl(n_ʱ�����, 0) Then
          n_������� := Nvl(n_ʱ�����, 0);
        End If;
        n_��� := Nvl(n_�������, 0) + 1;
      End If;
    Elsif Nvl(n_��ʱ��, 0) = 0 And Nvl(n_��ſ���, 0) = 0 And Nvl(������_In, 0) = 0 And Nvl(�ű�_In, 0) > 0 Then
      ---<--��ͨ��  -->
      Begin
        Select �ѹ���, ��Լ�� Into n_��������, n_��Լ�� From �ٴ������¼ Where ID = �����¼id_In;
      Exception
        When Others Then
          n_�������� := 0;
          n_��Լ��   := 0;
      End;
      If Nvl(��������_In, 0) = 0 Then
        If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
          --�Һ�
          If Nvl(n_��������, 0) >= Nvl(n_�޺���, 0) And Nvl(n_�޺���, 0) > 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        Else
          --ԤԼ
          If (Nvl(n_��Լ��, 0) > 0 And Nvl(n_��Լ��, 0) >= Nvl(n_��Լ��, 0)) Or
             (Nvl(n_��������, 0) >= Nvl(n_�޺���, 0) And Nvl(n_�޺���, 0) > 0) Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        End If;
      End If;
      ---<--��ͨ��  -->
    End If;
  End If;

  --���¹Һ����״̬
  If ���_In = 1 And Not n_��� Is Null Then
    If n_��ʱ�� = 1 Then
      d_���ʱ�� := ����ʱ��_In;
    Else
      d_���ʱ�� := Trunc(����ʱ��_In);
    End If;
    --������ŵĴ���
    Begin
      If n_ԤԼ˳��� Is Null Then
        Select ����Ա����, ����վ����
        Into v_��Ų���Ա, v_��Ż�����
        From �ٴ�������ſ���
        Where Nvl(�Һ�״̬, 0) = 5 And ��¼id = �����¼id_In And ��� = n_���;
      Else
        Select ����Ա����, ����վ����
        Into v_��Ų���Ա, v_��Ż�����
        From �ٴ�������ſ���
        Where Nvl(�Һ�״̬, 0) = 5 And ��¼id = �����¼id_In And ��� = n_ʱ����� And ԤԼ˳��� = n_ԤԼ˳���;
      End If;
      n_���� := 1;
    Exception
      When Others Then
        v_��Ų���Ա := Null;
        v_��Ż����� := Null;
        n_����       := 0;
    End;
    If n_���� = 0 Then
      If n_ԤԼ˳��� Is Null Then
        Update �ٴ�������ſ���
        Set �Һ�״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1)
        Where ��¼id = �����¼id_In And ��� = n_��� And Nvl(�Һ�״̬, 0) = 3 And ����Ա���� = ����Ա����_In;
      Else
        Update �ٴ�������ſ���
        Set �Һ�״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1)
        Where ��¼id = �����¼id_In And ��� = n_��� And ԤԼ˳��� = n_ԤԼ˳��� And Nvl(�Һ�״̬, 0) = 3 And ����Ա���� = ����Ա����_In;
      End If;
      If Sql%RowCount = 0 Then
        Begin
          If Nvl(n_��ʱ��, 0) > 0 Then
            If Nvl(n_��ſ���, 0) = 1 Then
              --��ʱ�κ�ר�Һ� ʧԼ��ԤԼ������Һ�
              Update �ٴ�������ſ���
              Set �Һ�״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1), ����Ա���� = ����Ա����_In
              Where ��¼id = �����¼id_In And ��� = n_��� And Nvl(�Һ�״̬, 0) In (0, 2);
              If Sql%NotFound Then
                Begin
                  Select �Һ�״̬ Into n_״̬ From �ٴ�������ſ��� Where ��¼id = �����¼id_In And ��� = n_���;
                  v_Err_Msg := '���' || n_��� || '�ѱ�ʹ��,������ѡ��һ�����.';
                  Raise Err_Item;
                Exception
                  When Others Then
                    Insert Into �ٴ�������ſ���
                      (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����ʱ��, ����, ����, ����Ա����, ��ע)
                      Select �����¼id_In, n_���, d_���ʱ��, d_���ʱ��, 1, Decode(ԤԼ�Һ�_In, 1, 1, 0), Decode(ԤԼ�Һ�_In, 1, 2, 1),
                             Null, Null, Null, ����Ա����_In, '׷�Ӻ�'
                      From Dual;
                End;
              End If;
            Else
              If Nvl(ԤԼ����_In, 0) = 1 Then
                Insert Into �ٴ�������ſ���
                  (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����ʱ��, ����, ����, ����Ա����, ��ע, ԤԼ˳���)
                  Select ��¼id, ���, ��ʼʱ��, ��ֹʱ��, 1, 1, Decode(ԤԼ�Һ�_In, 1, 2, 1), Null, Null, Null, ����Ա����_In, n_���, n_ԤԼ˳���
                  From �ٴ�������ſ���
                  Where ��¼id = �����¼id_In And ��� = n_ʱ����� And ԤԼ˳��� Is Null;
              End If;
            End If;
          Else
            If Nvl(n_��ſ���, 0) = 1 Then
              Update �ٴ�������ſ���
              Set �Һ�״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1)
              Where ��¼id = �����¼id_In And ��� = n_��� And Nvl(�Һ�״̬, 0) = 0;
            
              If Sql%RowCount = 0 Then
                Insert Into �ٴ�������ſ���
                  (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����ʱ��, ����, ����, ����Ա����, ��ע)
                  Select �����¼id_In, n_���, ����ʱ��_In, ����ʱ��_In, 1, Decode(ԤԼ�Һ�_In, 1, 1, 0), Decode(ԤԼ�Һ�_In, 1, 2, 1), Null,
                         Null, Null, ����Ա����_In, '׷�Ӻ�'
                  From Dual;
              End If;
            End If;
          End If;
        Exception
          When Others Then
            v_Err_Msg := '���' || n_��� || '�ѱ�ʹ��,������ѡ��һ�����.';
            Raise Err_Item;
        End;
      End If;
    Else
      If ����Ա����_In <> v_��Ų���Ա Or v_������ <> v_��Ż����� Then
        v_Err_Msg := '���' || n_��� || '�ѱ�������' || v_������ || '����,������ѡ��һ�����.';
        Raise Err_Item;
      Else
        If n_ԤԼ˳��� Is Null Then
          Update �ٴ�������ſ���
          Set �Һ�״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1)
          Where ��¼id = �����¼id_In And ��� = n_��� And Nvl(�Һ�״̬, 0) = 5 And ����Ա���� = ����Ա����_In And ����վ���� = v_������;
        Else
          Update �ٴ�������ſ���
          Set �Һ�״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1)
          Where ��¼id = �����¼id_In And ��� = n_ʱ����� And ԤԼ˳��� = n_ԤԼ˳��� And Nvl(�Һ�״̬, 0) = 5 And ����Ա���� = ����Ա����_In And
                ����վ���� = v_������;
        End If;
      End If;
    End If;
  End If;

  --�������˹Һŷ���(���ܵ����ǻ������������)
  Select ���˷��ü�¼_Id.Nextval Into n_����id From Dual; --Ӧ��ͨ������õ�

  Insert Into ������ü�¼
    (ID, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, NO, ʵ��Ʊ��, �����־, �Ӱ��־, ���ӱ�־, ��ҩ����, ����id, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, ���˿���id, �շ����,
     ���㵥λ, �շ�ϸĿid, ������Ŀid, �վݷ�Ŀ, ����, ����, ��׼����, Ӧ�ս��, ʵ�ս��, ���ʽ��, ����id, ���ʷ���, ��������id, ������, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����,
     ����ʱ��, �Ǽ�ʱ��, ���մ���id, ������Ŀ��, ���ձ���, ͳ����, ժҪ, ����, �ɿ���id)
  Values
    (n_����id, 4, Decode(ԤԼ�Һ�_In, 1, 0, 1), ���_In, Decode(�۸񸸺�_In, 0, Null, �۸񸸺�_In), ��������_In, ���ݺ�_In, Ʊ�ݺ�_In, 1, ����_In,
     ������_In, Decode(ԤԼ�Һ�_In, 1, To_Char(n_���), ����_In), Decode(����id_In, 0, Null, ����id_In),
     Decode(�����_In, 0, Null, �����_In), ���ʽ_In, ����_In, Decode(����_In, Null, Null, �Ա�_In), Decode(����_In, Null, Null, ����_In),
     v_�ѱ�, ���˿���id_In, �շ����_In, �ű�_In, �շ�ϸĿid_In, ������Ŀid_In, �վݷ�Ŀ_In, 1, ����_In, ��׼����_In, Ӧ�ս��_In, ʵ�ս��_In,
     Decode(ԤԼ�Һ�_In, 1, Null, Decode(Nvl(���ʷ���_In, 0), 1, Null, ʵ�ս��_In)),
     Decode(ԤԼ�Һ�_In, 1, Null, Decode(Nvl(���ʷ���_In, 0), 1, Null, ����id_In)), Decode(Nvl(���ʷ���_In, 0), 1, 1, 0), ��������id_In,
     ����Ա����_In, Decode(ԤԼ�Һ�_In, 1, ����Ա����_In, Null), ִ�в���id_In, ҽ������_In, ����Ա���_In, ����Ա����_In, ����ʱ��_In, �Ǽ�ʱ��_In, ���մ���id_In,
     ������Ŀ��_In, ���ձ���_In, ͳ����_In, ժҪ_In, ԤԼ��ʽ_In, Decode(ԤԼ�Һ�_In, 1, Null, n_��id));

  --���ܽ��㵽����Ԥ����¼
  If Nvl(ԤԼ�Һ�_In, 0) = 0 And Nvl(���ʷ���_In, 0) = 0 Then
    If Nvl(�ֽ�֧��_In, 0) = 0 And Nvl(����֧��_In, 0) = 0 And Nvl(Ԥ��֧��_In, 0) = 0 And ���_In = 1 Then
      Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      Insert Into ����Ԥ����¼
        (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
         ��������)
      Values
        (n_Ԥ��id, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), v_�ֽ�, 0, �Ǽ�ʱ��_In, ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�',
         n_��id, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, ������λ_In, 4);
    End If;
    If Nvl(�ֽ�֧��_In, 0) <> 0 And ���_In = 1 Then
      v_��������     := ���㷽ʽ_In || '|'; --�Կո�ֿ���|��β,û�н�������
      v_���㷽ʽ��¼ := '';
      While v_�������� Is Not Null Loop
        v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '|') - 1);
        v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
      
        v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
        n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
      
        v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
        v_������� := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
      
        v_��ǰ����   := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
        n_��������־ := To_Number(v_��ǰ����);
      
        If Instr('|' || v_���㷽ʽ��¼ || '|', '|' || Nvl(v_���㷽ʽ, v_�ֽ�) || '|') <> 0 Then
          v_Err_Msg := 'ʹ�����ظ��Ľ��㷽ʽ,����!';
          Raise Err_Item;
        Else
          v_���㷽ʽ��¼ := v_���㷽ʽ��¼ || '|' || Nvl(v_���㷽ʽ, v_�ֽ�);
        End If;
      
        If n_��������־ = 0 Then
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��,
             ������λ, ��������, �������)
          Values
            (n_Ԥ��id, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), Nvl(v_���㷽ʽ, v_�ֽ�), Nvl(n_������, 0), �Ǽ�ʱ��_In,
             ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�', n_��id, Null, Null, Null, Null, Null, ������λ_In, 4, v_�������);
        Else
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��,
             ������λ, ��������, �������)
          Values
            (n_Ԥ��id, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), Nvl(v_���㷽ʽ, v_�ֽ�), Nvl(n_������, 0), �Ǽ�ʱ��_In,
             ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�', n_��id, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, ������λ_In, 4,
             v_�������);
          If Nvl(���㿨���_In, 0) <> 0 Then
            n_���ѿ�id := Null;
            Begin
              Select Nvl(���ƿ�, 0), 1 Into n_���ƿ�, n_Count From �����ѽӿ�Ŀ¼ Where ��� = ���㿨���_In;
            Exception
              When Others Then
                n_Count := 0;
            End;
            If n_Count = 0 Then
              v_Err_Msg := 'û�з���ԭ���㿨����Ӧ���,���ܼ���������';
              Raise Err_Item;
            End If;
            If n_���ƿ� = 1 Then
              Select ID
              Into n_���ѿ�id
              From ���ѿ�Ŀ¼
              Where �ӿڱ�� = ���㿨���_In And ���� = ����_In And
                    ��� = (Select Max(���) From ���ѿ�Ŀ¼ Where �ӿڱ�� = ���㿨���_In And ���� = ����_In);
            End If;
            Zl_���˿������¼_Insert(���㿨���_In, n_���ѿ�id, Nvl(v_���㷽ʽ, v_�ֽ�), Nvl(n_������, 0), ����_In, Null, �Ǽ�ʱ��_In, Null, ����id_In,
                              n_Ԥ��id);
          End If;
        End If;
      
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + n_������
        Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = Nvl(v_���㷽ʽ, v_�ֽ�)
        Returning ��� Into n_����ֵ;
      
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ����
            (�տ�Ա, ���㷽ʽ, ����, ���)
          Values
            (����Ա����_In, Nvl(v_���㷽ʽ, v_�ֽ�), 1, n_������);
          n_����ֵ := n_������;
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From ��Ա�ɿ����
          Where �տ�Ա = ����Ա����_In And ���㷽ʽ = Nvl(v_���㷽ʽ, v_�ֽ�) And ���� = 1 And Nvl(���, 0) = 0;
        End If;
      
        v_�������� := Substr(v_��������, Instr(v_��������, '|') + 1);
      End Loop;
    End If;
  
    --����ҽ���Һ�
    If Nvl(����֧��_In, 0) <> 0 And ���_In = 1 Then
      Insert Into ����Ԥ����¼
        (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), v_�����ʻ�, ����֧��_In, �Ǽ�ʱ��_In, ����Ա���_In,
         ����Ա����_In, ����id_In, 'ҽ���Һ�', n_��id, 4);
    End If;
  
    --���ھ��￨ͨ��Ԥ����Һ�
    If Nvl(Ԥ��֧��_In, 0) <> 0 And ���_In = 1 Then
      n_Ԥ����� := Ԥ��֧��_In;
      For r_Deposit In c_Deposit(����id_In, v_��Ԥ������ids) Loop
        n_��ǰ��� := Case
                    When r_Deposit.��� - n_Ԥ����� < 0 Then
                     r_Deposit.���
                    Else
                     n_Ԥ�����
                  End;
      
        If r_Deposit.Id <> 0 Then
          --��һ�γ�Ԥ��(���Ͻ���ID,���Ϊ0)
          Update ����Ԥ����¼ Set ��Ԥ�� = 0, ����id = ����id_In, �������� = 4 Where ID = r_Deposit.Id;
        
        End If;
        --���ϴ�ʣ���
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���,
           ��Ԥ��, ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
          Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�,
                 �Ǽ�ʱ��_In, ����Ա����_In, ����Ա���_In, n_��ǰ���, ����id_In, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
          From ����Ԥ����¼
          Where NO = r_Deposit.No And ��¼״̬ = r_Deposit.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
      
        --���²���Ԥ�����
        Update �������
        Set Ԥ����� = Nvl(Ԥ�����, 0) - n_��ǰ���
        Where ����id = r_Deposit.����id And ���� = 1 And ���� = Nvl(r_Deposit.Ԥ�����, 2);
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
      If n_Ԥ����� > 0 Then
        v_Err_Msg := 'Ԥ���಻��֧������֧�����,���ܼ���������';
        Raise Err_Item;
      
      End If;
      Delete From ������� Where ����id = ����id_In And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
    End If;
  
    --��ػ��ܱ�Ĵ���
    --��Ա�ɿ����
    If ���_In = 1 Then
      If Nvl(����֧��_In, 0) <> 0 Then
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + ����֧��_In
        Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = v_�����ʻ�
        Returning ��� Into n_����ֵ;
      
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (����Ա����_In, v_�����ʻ�, 1, ����֧��_In);
          n_����ֵ := ����֧��_In;
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From ��Ա�ɿ����
          Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_�����ʻ� And Nvl(���, 0) = 0;
        End If;
      End If;
    End If;
  End If;

  --���˹ҺŻ���(ֻ����һ��,�ҵ�����ȡ�����Ѳ�����)
  If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
    If Nvl(���ʷ���_In, 0) = 0 Then
      --����Ʊ��ʹ�����
      If ���_In = 1 And Ʊ�ݺ�_In Is Not Null Then
        Select Ʊ�ݴ�ӡ����_Id.Nextval Into n_��ӡid From Dual;
      
        --����Ʊ��
        Insert Into Ʊ�ݴ�ӡ���� (ID, ��������, NO) Values (n_��ӡid, 4, ���ݺ�_In);
      
        Insert Into Ʊ��ʹ����ϸ
          (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
        Values
          (Ʊ��ʹ����ϸ_Id.Nextval, Decode(�շ�Ʊ��_In, 1, 1, 4), Ʊ�ݺ�_In, 1, 1, ����id_In, n_��ӡid, �Ǽ�ʱ��_In, ����Ա����_In);
      
        --״̬�Ķ�
        Update Ʊ�����ü�¼
        Set ��ǰ���� = Ʊ�ݺ�_In, ʣ������ = Decode(Sign(ʣ������ - 1), -1, 0, ʣ������ - 1), ʹ��ʱ�� = Sysdate
        Where ID = Nvl(����id_In, 0);
      End If;
    End If;
    --���˱��ξ���(�Է���ʱ��Ϊ׼)
    If Nvl(����id_In, 0) <> 0 And ���_In = 1 Then
      Update ������Ϣ Set ����ʱ�� = ����ʱ��_In, ����״̬ = 1, �������� = ����_In Where ����id = ����id_In;
    End If;
  End If;

  If Nvl(ԤԼ�Һ�_In, 0) = 0 And Nvl(���ʷ���_In, 0) = 1 Then
    --����
    If Nvl(����id_In, 0) = 0 Then
      v_Err_Msg := 'Ҫ��Բ��˵ĹҺŷѽ��м��ʣ������ǽ������˲��ܼ��ʹҺš�';
      Raise Err_Item;
    End If;
  
    --�������
    Update �������
    Set ������� = Nvl(�������, 0) + Nvl(ʵ�ս��_In, 0)
    Where ����id = Nvl(����id_In, 0) And ���� = 1 And ���� = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into ������� (����id, ����, ����, �������, Ԥ�����) Values (����id_In, 1, 1, Nvl(ʵ�ս��_In, 0), 0);
    End If;
  
    --����δ�����
    Update ����δ�����
    Set ��� = Nvl(���, 0) + Nvl(ʵ�ս��_In, 0)
    Where ����id = ����id_In And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(���˿���id_In, 0) And
          Nvl(��������id, 0) = Nvl(��������id_In, 0) And Nvl(ִ�в���id, 0) = Nvl(ִ�в���id_In, 0) And ������Ŀid + 0 = ������Ŀid_In And
          ��Դ;�� + 0 = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into ����δ�����
        (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
      Values
        (����id_In, Null, Null, ���˿���id_In, ��������id_In, ִ�в���id_In, ������Ŀid_In, 1, Nvl(ʵ�ս��_In, 0));
    End If;
  End If;

  --���˹Һż�¼
  If �ű�_In Is Not Null And ���_In = 1 Then
    --And Nvl(ԤԼ�Һ�_In, 0) = 0
    Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
    Begin
      Select ���� Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = ���ʽ_In And Rownum < 2;
    Exception
      When Others Then
        v_���ʽ := Null;
    End;
    Insert Into ���˹Һż�¼
      (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ԤԼʱ��, ����Ա���,
       ����Ա����, ����, ����, ����, ԤԼ, ԤԼ��ʽ, ժҪ, ������ˮ��, ����˵��, ������λ, ����ʱ��, ������, ԤԼ����Ա, ԤԼ����Ա���, ����, ҽ�Ƹ��ʽ, �����¼id)
    Values
      (n_�Һ�id, ���ݺ�_In, Decode(Nvl(ԤԼ�Һ�_In, 0), 1, 2, 1), 1, ����id_In, �����_In, ����_In, �Ա�_In, ����_In, �ű�_In, ����_In, ����_In,
       Null, ִ�в���id_In, ҽ������_In, 0, Null, �Ǽ�ʱ��_In, ����ʱ��_In, Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����ʱ��_In, Null), ����Ա���_In,
       ����Ա����_In, ����_In, n_���, ����_In, Decode(ԤԼ����_In, 1, 1, 0), ԤԼ��ʽ_In, ժҪ_In, ������ˮ��_In, ����˵��_In, ������λ_In,
       Decode(Nvl(ԤԼ�Һ�_In, 0), 0, �Ǽ�ʱ��_In, Null), Decode(Nvl(ԤԼ�Һ�_In, 0), 0, ����Ա����_In, Null),
       Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����Ա����_In, Null), Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����Ա���_In, Null),
       Decode(Nvl(����_In, 0), 0, Null, ����_In), v_���ʽ, �����¼id_In);
  
    n_ԤԼ���ɶ��� := 0;
    If Nvl(ԤԼ�Һ�_In, 0) = 1 Then
      n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
    End If;
  
    --0-����������;1-��ҽ�������̨�Ŷ�;2-�ȷ���,��ҽ��վ
    If Nvl(���ɶ���_In, 0) <> 0 And Nvl(ԤԼ�Һ�_In, 0) = 0 Or n_ԤԼ���ɶ��� = 1 Then
      n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113));
      If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Or n_ԤԼ���ɶ��� = 1 Then
      
        --��������
        --.����ִ�в��š� �ķ�ʽ���ɶ���
        v_�������� := ִ�в���id_In;
        v_�ŶӺ��� := Zlgetnextqueue(ִ�в���id_In, n_�Һ�id, �ű�_In || '|' || n_���);
        v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 0);
        d_�Ŷ�ʱ�� := Zl_Get_Queuedate(n_�Һ�id, �ű�_In, n_���, ����ʱ��_In);
        --  ��������_In , ҵ������_In, ҵ��id_In,����id_In,�ŶӺ���_In,�Ŷӱ��_In,��������_In,����ID_IN, ����_In, ҽ������_In, �Ŷ�ʱ��_In
        Zl_�ŶӽкŶ���_Insert(v_��������, 0, n_�Һ�id, ִ�в���id_In, v_�ŶӺ���, Null, ����_In, ����id_In, ����_In, ҽ������_In, d_�Ŷ�ʱ��, ԤԼ��ʽ_In,
                         Null, v_�Ŷ����);
      End If;
    End If;
  End If;
  --���˵�����Ϣ
  If ����id_In Is Not Null And ���_In = 1 Then
    --ȡ����:
    If Nvl(n_����ģʽ, 0) <> Nvl(����ģʽ_In, 0) Then
      --����ģʽ��ȷ��
    
      v_Err_Msg := Null;
      Begin
        Select Nvl(����ģʽ, 0) Into n_����ģʽ From ������Ϣ Where ����id = ����id_In;
      Exception
        When Others Then
          v_Err_Msg := 'δ�ҵ�ָ���Ĳ�����Ϣ,������Һ�';
      End;
    
      If v_Err_Msg Is Not Null Then
        Raise Err_Item;
      End If;
      If n_����ģʽ = 1 And Nvl(����ģʽ_In, 0) = 0 Then
        --�����Ѿ���"�����ƺ�����",������"�Ƚ�������Ƶ�",�����Ƿ����δ������
        Select Count(1)
        Into n_Count
        From ����δ�����
        Where ����id = ����id_In And (��Դ;�� In (1, 4) Or ��Դ;�� = 3 And Nvl(��ҳid, 0) = 0) And Nvl(���, 0) <> 0 And Rownum < 2;
        If Nvl(n_Count, 0) <> 0 Then
          --����δ�������ݣ������Ƚ���������ִ��
          v_Err_Msg := '��ǰ���˵ľ���ģʽΪ�����ƺ�����Ҵ���δ����ã�����������ò��˵ľ���ģʽ,������ȶ�δ����ý��ʺ��ٹҺŻ򲻵������˵ľ���ģʽ!';
          Raise Err_Item;
        End If;
        --���
        --δ����ҽ��ҵ��ģ�����ʱ�͹Һŵ�,��Ҫ��֤ͬһ�εľ���ģʽ��һ����(�����Ѿ���飬�����ٴ���)
      End If;
      Update ������Ϣ Set ����ģʽ = ����ģʽ_In Where ����id = ����id_In;
    End If;
  
    Update ������Ϣ
    Set ������ = Null, ������ = Null, �������� = Null
    Where ����id = ����id_In And Exists (Select 1
           From ���˵�����¼
           Where ����id = ����id_In And ��ҳid Is Not Null And
                 �Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ���˵�����¼ Where ����id = ����id_In));
  
    If Sql%RowCount > 0 Then
      Update ���˵�����¼
      Set ����ʱ�� = Sysdate
      Where ����id = ����id_In And ��ҳid Is Not Null And Nvl(����ʱ��, Sysdate) > Sysdate;
    End If;
  End If;
  If ���_In = 1 Then
    --��Ϣ����
    Begin
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
        Using 1, n_�Һ�id;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_001(n_�Һ�id, ���ݺ�_In);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˹Һż�¼_����_Insert;
/



Create Or Replace Procedure Zl_���˹Һż�¼_��������
(
  No_In       ���˹Һż�¼.No%Type := Null,
  ����id_In   ���˹Һż�¼.����id%Type := Null,
  ����_In     ���˹Һż�¼.����%Type := Null,
  ҽ��_In     ���˹Һż�¼.ִ����%Type := Null,
  ����ʱ��_In ���˹Һż�¼.����ʱ��%Type := Null,
  ��������_In Integer := 1,
  ԤԼ��ʽ_In ԤԼ��ʽ.����%Type := Null
) As
  v_Id           ������ü�¼.Id%Type := Null;
  v_�Һ����ɶ��� Varchar2(2);
  v_�ŶӺ���     �ŶӽкŶ���.�ŶӺ���%Type;
  v_�Ŷ����     �ŶӽкŶ���.�Ŷ����%Type;
  n_�����Ŷ�     Number(18);
  n_��������     ���˹Һż�¼.��¼����%Type;
  n_����id       ��������.Id%Type;
  n_�Ŷ�         Number(18);
  v_Error        Varchar2(255);
  Err_Custom Exception;
Begin
  If Nvl(��������_In, 0) = 2 Then
    Begin
      Select ID Into v_Id From ���˹Һż�¼ Where NO = No_In;
      Update �ŶӽкŶ��� Set ���� = ����_In, ҽ������ = ҽ��_In Where ҵ��id = v_Id;
    Exception
      When Others Then
        Null;
    End;
  Else
    Begin
      Select ID, ��¼���� Into v_Id, n_�������� From ���˹Һż�¼ Where NO = No_In And Nvl(ִ��״̬, 0) = 0;
      Select ID Into n_����id From �������� Where ���� = ����_In;
    Exception
      When Others Then
        Null;
    End;
    If v_Id Is Null Then
      v_Error := '�����Ѿ�������Ѿ��˺ţ������ٷ��';
      Raise Err_Custom;
    End If;
    If Nvl(��������_In, 0) = 1 Then
      If Nvl(����id_In, 0) <> 0 Then
        --���²�����Ϣ
        Update ������Ϣ Set �������� = ����_In Where ����id = ����id_In And ����״̬ = 1;
      End If;
    
      --���·��ü�¼
      Update ������ü�¼
      Set ��ҩ���� = ����_In, ִ���� = ҽ��_In
      Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_��������, 0), 2, 0, 1) And NO = No_In;
      --���²��˹Һż�¼
      Update ���˹Һż�¼
      Set ���� = ����_In, ִ���� = ҽ��_In, ����ʱ�� = Decode(����ʱ��_In, Null, ����ʱ��, ����ʱ��_In)
      Where NO = No_In;
    End If;
    v_�Һ����ɶ��� := Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113));
    If v_�Һ����ɶ��� <> 0 Then
      For c_�Һ� In (Select ID, ִ�в���id, ����, ����_In As ����, �Ǽ�ʱ��, ҽ��_In As ִ����, ����id, �ű�, ����
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
                           Nvl(����ʱ��_In, Sysdate), ԤԼ��ʽ_In, Null, v_�Ŷ����);
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
      End Loop;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˹Һż�¼_��������;
/


Create Or Replace Procedure Zl_���˹Һż�¼_����
(
  No_In         ���˹Һż�¼.No%Type,
  �ű�_In       ���˹Һż�¼.�ű�%Type,
  ����_In       ���˹Һż�¼.����%Type,
  ����id_In     ���˹Һż�¼.ִ�в���id%Type,
  ԭҽ��_In     ���˹Һż�¼.ִ����%Type,
  ԭҽ��id_In   ���˹ҺŻ���.ҽ��id%Type,
  ��ҽ��_In     ���˹Һż�¼.ִ����%Type,
  ��ҽ��id_In   ���˹ҺŻ���.ҽ��id%Type,
  �����¼id_In �ٴ������¼.Id%Type := Null
  --���ܣ���ɲ��˻��Ź��ܣ��ڹҺ���ĿID��ͬ������¡�
) As
  Cursor c_Bill Is
    Select ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, ���ʵ�id, ����id, ҽ�����, �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��, ���ʽ, ���˿���id, �ѱ�,
           �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id, ������, ����ʱ��,
           �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���, ����Ա����, ����id, ���ʽ��, ���մ���id, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���
    From ������ü�¼
    Where ��¼���� = 4 And ��¼״̬ = 1 And NO = No_In
    Order By ���;

  v_����id       ������ü�¼.Id%Type;
  v_�ֶ�������   �ŶӽкŶ���.��������%Type;
  v_�Һ����ɶ��� Varchar2(2);
  v_ԤԼ�Һ�     Number(2);
  n_ҵ��id       ���˹Һż�¼.Id%Type;
  v_�ŶӺ���     �ŶӽкŶ���.�ŶӺ���%Type;
  v_�ű�         ���˹Һż�¼.�ű�%Type;
  n_����         ���˹Һż�¼.����%Type;
  v_�Ŷ����     �ŶӽкŶ���.�Ŷ����%Type;
  v_Temp         Varchar2(500);
  v_����Ա���   ����䶯��¼.����Ա���%Type;
  v_����Ա����   ����䶯��¼.����Ա����%Type;
  n_ҽ��id       ��Ա��.Id%Type;
  n_����id       ��������.Id%Type;
  n_ԭ�����¼id �ٴ������¼.Id%Type;
  n_�䶯id       ����䶯��¼.Id%Type;
  v_Error        Varchar2(255);
  Err_Custom Exception;
Begin
  v_����id := 0;
  If �����¼id_In Is Null Then
    Begin
      Select ����id Into v_����id From ���˹Һż�¼ Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 1;
    Exception
      When Others Then
        Null;
    End;
    If v_����id = 0 Then
      v_Error := 'û���ҵ����˵ĹҺ���Ϣ��';
      Raise Err_Custom;
    Elsif v_����id Is Null Then
      v_Error := 'û���ҵ�������Ϣ��';
      Raise Err_Custom;
    End If;
  
    ---�ȸ��²�����Ϣ�ľ������Һ�״̬
    Update ������Ϣ Set �������� = ����_In, ����״̬ = 1 Where ����id = v_����id And ����״̬ In (1, 2);
  
    For r_Bill In c_Bill Loop
      If r_Bill.��� = 1 Then
      
        --��Ҫȷ���Ƿ�ԤԼ�Һ�
        --1.�������ԤԼ�ҺŲ����ĹҺż�¼,����Ҫ���ѹ����������ѽ���
        --2.����������Һ�,��ֻ���ѹ���
        Begin
          Select Decode(ԤԼ, Null, 0, 0, 0, 1) Into v_ԤԼ�Һ� From ���˹Һż�¼ Where NO = r_Bill.No And Rownum = 1;
        Exception
          When Others Then
            v_ԤԼ�Һ� := 0;
        End;
      
        --�ָ���ǰ�ĹҺŻ���
        Update ���˹ҺŻ���
        Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - v_ԤԼ�Һ�, ��Լ�� = Nvl(��Լ��, 0) - v_ԤԼ�Һ�
        Where ���� = Trunc(r_Bill.�Ǽ�ʱ��) And Nvl(����id, 0) = Nvl(r_Bill.ִ�в���id, 0) And Nvl(��Ŀid, 0) = Nvl(r_Bill.�շ�ϸĿid, 0) And
              (���� = r_Bill.���㵥λ Or ���� Is Null);
        If Sql%RowCount = 0 Then
          Insert Into ���˹ҺŻ���
            (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, ��Լ��, �����ѽ���)
          Values
            (Trunc(r_Bill.�Ǽ�ʱ��), r_Bill.ִ�в���id, r_Bill.�շ�ϸĿid, ԭҽ��_In, Decode(ԭҽ��id_In, 0, Null, ԭҽ��id_In), r_Bill.���㵥λ,
             -1, -1 * v_ԤԼ�Һ�, -1 * v_ԤԼ�Һ�);
        End If;
      
        ----Ȼ���ٸ��¹ҺŻ���
        Update ���˹ҺŻ���
        Set �ѹ��� = Nvl(�ѹ���, 0) + 1, �����ѽ��� = Nvl(�����ѽ���, 0) + v_ԤԼ�Һ�, ��Լ�� = Nvl(��Լ��, 0) + v_ԤԼ�Һ�
        Where ���� = Trunc(r_Bill.�Ǽ�ʱ��) And Nvl(����id, 0) = ����id_In And Nvl(��Ŀid, 0) = Nvl(r_Bill.�շ�ϸĿid, 0) And
              (���� = �ű�_In Or ���� Is Null);
        If Sql%RowCount = 0 Then
          Insert Into ���˹ҺŻ���
            (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, ��Լ��, �����ѽ���)
          Values
            (Trunc(r_Bill.�Ǽ�ʱ��), ����id_In, r_Bill.�շ�ϸĿid, ��ҽ��_In, Decode(��ҽ��id_In, 0, Null, ��ҽ��id_In), �ű�_In, 1, v_ԤԼ�Һ�,
             v_ԤԼ�Һ�);
        End If;
      End If;
    
      ---���¹Һż�¼
      Update ������ü�¼
      Set ִ�в���id = ����id_In, ���˿���id = ����id_In, ���㵥λ = �ű�_In, ��ҩ���� = ����_In,
          --���˲���id = ����id_In,
          ִ���� = ��ҽ��_In, ִ��״̬ = 0, ִ��ʱ�� = Null
      Where ID = r_Bill.Id;
    
      --���²��˹Һż�¼
      If r_Bill.��� = 1 Then
        v_Temp := Zl_Identity(1);
        Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_Temp From Dual;
        Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into v_����Ա��� From Dual;
        Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_����Ա���� From Dual;
        Begin
          Select ID Into n_ҽ��id From ��Ա�� Where ���� = ��ҽ��_In And Rownum < 2;
        Exception
          When Others Then
            n_ҽ��id := Null;
        End;
        Select ����䶯��¼_Id.Nextval Into n_�䶯id From Dual;
        b_Message.Zlhis_Regist_005(r_Bill.No, n_�䶯id, 2);
        Zl_����䶯��¼_Insert(r_Bill.No, 2, '���ﻻ��', v_����Ա����, v_����Ա���, �ű�_In, ����id_In, Null, n_ҽ��id, ��ҽ��_In, ����_In, n_����,
                         Null, n_�䶯id);
        v_�Һ����ɶ��� := Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113));
        If v_�Һ����ɶ��� <> 0 Then
          v_�ֶ������� := ����id_In;
          Select ID, �ű�, Nvl(����, 0)
          Into n_ҵ��id, v_�ű�, n_����
          From ���˹Һż�¼
          Where NO = r_Bill.No And Rownum = 1;
          --Zlgetnextqueue(ִ�в���id_In Number,ҵ��id_In     Number := Null)
          v_�ŶӺ��� := Zlgetnextqueue(����id_In, n_ҵ��id, v_�ű� || '|' || n_����);
          v_�Ŷ���� := Zlgetsequencenum(0, n_ҵ��id, 1);
          --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In , ҽ������_In
          Zl_�ŶӽкŶ���_Update(v_�ֶ�������, 0, n_ҵ��id, ����id_In, r_Bill.����, ����_In, ��ҽ��_In, v_�ŶӺ���, v_�Ŷ����);
        End If;
        Update ���˹Һż�¼
        Set ִ�в���id = ����id_In, �ű� = �ű�_In, ���� = ����_In, ִ���� = ��ҽ��_In, ִ��״̬ = 0, ִ��ʱ�� = Null
        Where NO = r_Bill.No;
      End If;
    End Loop;
  Else
    --������Ű�ģʽ
    Begin
      Select ����id, �����¼id
      Into v_����id, n_ԭ�����¼id
      From ���˹Һż�¼
      Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 1;
      Select ID Into n_����id From �������� Where ���� = ����_In;
    Exception
      When Others Then
        Null;
    End;
    If v_����id = 0 Then
      v_Error := 'û���ҵ����˵ĹҺ���Ϣ��';
      Raise Err_Custom;
    Elsif v_����id Is Null Then
      v_Error := 'û���ҵ�������Ϣ��';
      Raise Err_Custom;
    End If;
  
    ---�ȸ��²�����Ϣ�ľ������Һ�״̬
    Update ������Ϣ Set �������� = ����_In, ����״̬ = 1 Where ����id = v_����id And ����״̬ In (1, 2);
  
    For r_Bill In c_Bill Loop
      If r_Bill.��� = 1 Then
      
        --��Ҫȷ���Ƿ�ԤԼ�Һ�
        --1.�������ԤԼ�ҺŲ����ĹҺż�¼,����Ҫ���ѹ����������ѽ���
        --2.����������Һ�,��ֻ���ѹ���
        Begin
          Select Decode(ԤԼ, Null, 0, 0, 0, 1) Into v_ԤԼ�Һ� From ���˹Һż�¼ Where NO = r_Bill.No And Rownum = 1;
        Exception
          When Others Then
            v_ԤԼ�Һ� := 0;
        End;
      
        --�ָ���ǰ�ĹҺŻ���
        Update ���˹ҺŻ���
        Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - v_ԤԼ�Һ�, ��Լ�� = Nvl(��Լ��, 0) - v_ԤԼ�Һ�
        Where ���� = Trunc(r_Bill.�Ǽ�ʱ��) And Nvl(ҽ��id, 0) = Nvl(ԭҽ��id_In, 0) And Nvl(ҽ������, '-') = Nvl(ԭҽ��_In, '-') And
              Nvl(����id, 0) = Nvl(r_Bill.ִ�в���id, 0) And Nvl(��Ŀid, 0) = Nvl(r_Bill.�շ�ϸĿid, 0) And
              (���� = r_Bill.���㵥λ Or ���� Is Null);
        If Sql%RowCount = 0 Then
          Insert Into ���˹ҺŻ���
            (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, ��Լ��, �����ѽ���)
          Values
            (Trunc(r_Bill.�Ǽ�ʱ��), r_Bill.ִ�в���id, r_Bill.�շ�ϸĿid, ԭҽ��_In, Decode(ԭҽ��id_In, 0, Null, ԭҽ��id_In), r_Bill.���㵥λ,
             -1, -1 * v_ԤԼ�Һ�, -1 * v_ԤԼ�Һ�);
        End If;
        Update �ٴ������¼
        Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - v_ԤԼ�Һ�, ��Լ�� = Nvl(��Լ��, 0) - v_ԤԼ�Һ�
        Where ID = n_ԭ�����¼id;
      
        ----Ȼ���ٸ��¹ҺŻ���
        Update ���˹ҺŻ���
        Set �ѹ��� = Nvl(�ѹ���, 0) + 1, �����ѽ��� = Nvl(�����ѽ���, 0) + v_ԤԼ�Һ�, ��Լ�� = Nvl(��Լ��, 0) + v_ԤԼ�Һ�
        Where ���� = Trunc(r_Bill.�Ǽ�ʱ��) And Nvl(����id, 0) = ����id_In And Nvl(��Ŀid, 0) = Nvl(r_Bill.�շ�ϸĿid, 0) And
              (���� = �ű�_In Or ���� Is Null);
        If Sql%RowCount = 0 Then
          Insert Into ���˹ҺŻ���
            (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, ��Լ��, �����ѽ���)
          Values
            (Trunc(r_Bill.�Ǽ�ʱ��), ����id_In, r_Bill.�շ�ϸĿid, ��ҽ��_In, Decode(��ҽ��id_In, 0, Null, ��ҽ��id_In), �ű�_In, 1, v_ԤԼ�Һ�,
             v_ԤԼ�Һ�);
        End If;
        Update �ٴ������¼
        Set �ѹ��� = Nvl(�ѹ���, 0) + 1, �����ѽ��� = Nvl(�����ѽ���, 0) + v_ԤԼ�Һ�, ��Լ�� = Nvl(��Լ��, 0) + v_ԤԼ�Һ�
        Where ID = �����¼id_In;
      End If;
    
      ---���¹Һż�¼
      Update ������ü�¼
      Set ִ�в���id = ����id_In, ���˿���id = ����id_In, ���㵥λ = �ű�_In, ��ҩ���� = ����_In,
          --���˲���id = ����id_In,
          ִ���� = ��ҽ��_In, ִ��״̬ = 0, ִ��ʱ�� = Null
      Where ID = r_Bill.Id;
    
      --���²��˹Һż�¼
      If r_Bill.��� = 1 Then
        v_Temp := Zl_Identity(1);
        Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_Temp From Dual;
        Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into v_����Ա��� From Dual;
        Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_����Ա���� From Dual;
        Begin
          Select ID Into n_ҽ��id From ��Ա�� Where ���� = ��ҽ��_In And Rownum < 2;
        Exception
          When Others Then
            n_ҽ��id := Null;
        End;
        Select ����䶯��¼_Id.Nextval Into n_�䶯id From Dual;
        b_Message.Zlhis_Regist_005(r_Bill.No, n_�䶯id, 2);
        Zl_����䶯��¼_Insert(r_Bill.No, 2, '���ﻻ��', v_����Ա����, v_����Ա���, �ű�_In, ����id_In, Null, n_ҽ��id, ��ҽ��_In, ����_In, n_����,
                         Null, n_�䶯id);
        v_�Һ����ɶ��� := Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113));
        If v_�Һ����ɶ��� <> 0 Then
          v_�ֶ������� := ����id_In;
          Select ID, �ű�, Nvl(����, 0)
          Into n_ҵ��id, v_�ű�, n_����
          From ���˹Һż�¼
          Where NO = r_Bill.No And Rownum = 1;
          --Zlgetnextqueue(ִ�в���id_In Number,ҵ��id_In     Number := Null)
          v_�ŶӺ��� := Zlgetnextqueue(����id_In, n_ҵ��id, v_�ű� || '|' || n_����);
          v_�Ŷ���� := Zlgetsequencenum(0, n_ҵ��id, 1);
          --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In , ҽ������_In
          Zl_�ŶӽкŶ���_Update(v_�ֶ�������, 0, n_ҵ��id, ����id_In, r_Bill.����, ����_In, ��ҽ��_In, v_�ŶӺ���, v_�Ŷ����);
        End If;
        Update ���˹Һż�¼
        Set ִ�в���id = ����id_In, �ű� = �ű�_In, ���� = ����_In, ִ���� = ��ҽ��_In, ִ��״̬ = 0, ִ��ʱ�� = Null, �����¼id = �����¼id_In
        Where NO = r_Bill.No;
      End If;
    End Loop;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˹Һż�¼_����;
/


Create Or Replace Procedure Zl_���˹Һż�¼_��������
(
  Nos_In        In Varchar2 := Null,
  �ºű�_In     In ���˹Һż�¼.�ű�%Type := Null,
  ��ҽ������_In In �ҺŰ���.ҽ������%Type := Null,
  ��ҽ��id_In   In �ҺŰ���.ҽ��id%Type := Null,
  �¿���id_In   In �ҺŰ���.����id%Type := Null,
  ԭҽ������_In In �ҺŰ���.ҽ������%Type := Null,
  ԭҽ��id_In   In �ҺŰ���.ҽ��id%Type := Null,
  ԭ�ű�_In     In ���˹Һż�¼.�ű�%Type := Null,
  ����Ա����_In In �Һ����״̬.����Ա����%Type := Null,
  ԭ����id_In   In �ٴ������¼.Id%Type := Null,
  �³���id_In   In �ٴ������¼.Id%Type := Null
  --����: ��ɲ����������Ź���,�ڹҺ���Ŀ��ͬ,�޺�����ͬ,��Լ����ͬ,������ͬ������¡�
  --����˵��:  Nos_In :��Ҫ�����Ű�Ĳ��˹Һż�¼���ݼ�:��ʽ: M000001|M000002|..........
) As
  --��ȡ��Ӧ�Һż�¼��������ü�¼��Ϣ
  Cursor c_Bill(c_No ���˹Һż�¼.No%Type) Is
    Select ID, ���, NO, ����ʱ��, ִ�в���id, �շ�ϸĿid, ���㵥λ
    From ������ü�¼
    Where ��¼���� = 4 And ��¼״̬ In (0, 1) And NO = c_No
    Order By ���;
  --��ȡ��Ӧ�Ű�ķ�������
  Cursor c_ƽ������(c_ָ������ű�id �ҺŰ���.Id%Type) Is
    Select �ű�id, ��������, ��ǰ���� From �ҺŰ������� Where �ű�id = c_ָ������ű�id;
  Cursor c_����ƽ������ Is
    Select ��¼id, ����id, ��ǰ���� From �ٴ��������Ҽ�¼ Where ��¼id = �³���id_In;

  --��������
  r_ƽ������         �ҺŰ�������%RowType;
  r_����ƽ������     �ٴ��������Ҽ�¼%RowType;
  r_Bill             c_Bill%RowType;
  v_Nos              Varchar(2000);
  v_No               ���˹Һż�¼.No%Type;
  n_����id           ���˹Һż�¼.����id%Type;
  n_ԭ���           ���˹Һż�¼.����%Type;
  d_ԭ��������       ���˹Һż�¼.ԤԼʱ��%Type;
  n_�Ƿ��ѱ��ҳ�     Number(1);
  n_��¼����         Number(1);
  n_ԤԼ             Number(1);
  n_�Һ�״̬         Number(1); --0-�����Һ�:1-ԤԼ�Һ�;2-ԤԼ�ҺŽ���
  v_�¾�������       ������Ϣ.��������%Type;
  n_���﷽ʽ         Number(1); --0-������:1-ָ������:2-��̬����:3-ƽ������
  n_ָ������ű�id   Number(10);
  n_������������     Number(3);
  n_�Ƿ��ҵ��������� Number(1); --0:δ�ҵ�:1-�ҵ��������ʶδ����:2-�޸ĵ�һ�����ݱ�ʶ
  n_Index            Number(1); --��ǰ��¼��������ֵ
  v_�ֶ�������       �ŶӽкŶ���.��������%Type;
  n_�Һ����ɶ���     Number;
  n_ԤԼ���ɶ���     Number;
  v_�ŶӺ���         �ŶӽкŶ���.�ŶӺ���%Type;
  n_ҵ��id           ���˹Һż�¼.Id%Type;
  v_Temp             Varchar2(500);
  v_����Ա���       ����䶯��¼.����Ա���%Type;
  v_����Ա����       ����䶯��¼.����Ա����%Type;
  n_ҽ��id           ��Ա��.Id%Type;
  v_Error            Varchar2(255);
  n_�Һ��Ű�ģʽ     Number(3);
  n_�����¼id       �ٴ������¼.Id%Type;
  n_�䶯id           ����䶯��¼.Id%Type;
  Err_Custom Exception;
Begin
  n_�Һ��Ű�ģʽ := Zl_To_Number(Nvl(zl_GetSysParameter('�Һ��Ű�ģʽ'), 0));
  If n_�Һ��Ű�ģʽ = 0 Then
    --�ƻ��Ű�ģʽ
    --����Ƿ���ڸùҺż�¼
    If Nos_In Is Not Null Then
      v_Nos := Nos_In || '|';
      While v_Nos Is Not Null Loop
        --��ʼ������
        n_����id           := 0;
        n_ԭ���           := 0;
        d_ԭ��������       := Null;
        n_�Ƿ��ѱ��ҳ�     := 0;
        n_��¼����         := 0;
        n_ԤԼ             := 0;
        n_�Һ�״̬         := 0;
        v_�¾�������       := '';
        n_���﷽ʽ         := 0;
        n_ָ������ű�id   := 0;
        v_�ֶ�������       := '';
        v_�ŶӺ���         := '';
        n_ҵ��id           := 0;
        n_������������     := 0;
        n_�Һ����ɶ���     := 0;
        n_ԤԼ���ɶ���     := 0;
        n_�Ƿ��ҵ��������� := 0;
        n_Index            := 0;
      
        v_No  := Substr(v_Nos, 1, Instr(v_Nos, '|') - 1);
        v_Nos := Substr(v_Nos, Instr(v_Nos, '|') + 1);
        --����Ƿ���ڸùҺż�¼
        Begin
          Select a.Id, a.����id, a.����, Nvl(b.����, Nvl(a.ԤԼʱ��, a.����ʱ��)), a.��¼����, Nvl(a.ԤԼ, 0)
          Into n_ҵ��id, n_����id, n_ԭ���, d_ԭ��������, n_��¼����, n_ԤԼ
          From ���˹Һż�¼ A, �Һ����״̬ B
          Where NO = v_No And ��¼���� In (1, 2) And ��¼״̬ = 1 And a.�ű� = b.����(+) And
                Trunc(Nvl(ԤԼʱ��, ����ʱ��)) = Trunc(b.����(+)) And a.���� = b.���(+);
        Exception
          When Others Then
            Null;
        End;
        If n_����id = 0 Then
          v_Error := 'û���ҵ����˵ĹҺ���Ϣ';
          Raise Err_Custom;
        End If;
        --�жϵ�ǰ�Һ�״̬
        n_�Һ�״̬ := 0; --�����Һ�
        If n_��¼���� = 1 And n_ԤԼ = 1 Then
          n_�Һ�״̬ := 2; --ԤԼ����
        End If;
        If n_��¼���� = 2 And n_ԤԼ = 1 Then
          n_�Һ�״̬ := 1; --ԤԼ
        End If;
      
        --��黻�ŵ��ºű��Ƿ��ѱ��ҳ�
        Begin
          Select a.״̬
          Into n_�Ƿ��ѱ��ҳ�
          From �Һ����״̬ A
          Where a.���� = d_ԭ�������� And a.���� = �ºű�_In And a.��� = n_ԭ���;
        Exception
          When Others Then
            n_�Ƿ��ѱ��ҳ� := 0;
        End;
        If n_�Ƿ��ѱ��ҳ� > 0 Then
          v_Error := 'Ҫ���ĺű��ѱ��ҳ�';
          Raise Err_Custom;
        End If;
        --ԤԼ���յ�����½��з������ҵĻ�ȡ
        If n_�Һ�״̬ = 2 Then
          --��ȡ�ºű�����
          --˵��:ԤԼ������£�����Ҫ�����˲��û�ȡ��������
          --     ���յ������,��Ҫ���з���,�����Ҫ��ȡ��������
          --��ȡ���﷽ʽ
          Begin
            Select ID, Nvl(���﷽ʽ, 0) Into n_ָ������ű�id, n_���﷽ʽ From �ҺŰ��� Where ���� = �ºű�_In;
          Exception
            When Others Then
              n_���﷽ʽ       := 0;
              n_ָ������ű�id := 0;
          End;
        
          Begin
            If n_���﷽ʽ = 0 Then
              --������
              v_�¾������� := '';
            End If;
            If n_���﷽ʽ = 1 Then
              --ָ������
              Select �������� Into v_�¾������� From �ҺŰ������� Where �ű�id = n_ָ������ű�id;
            End If;
            If n_���﷽ʽ = 2 Then
              --��̬����
              Select ��������
              Into v_�¾�������
              From (Select ��������, Sum(Num) As Num
                     From (Select ��������, 0 As Num
                            From �ҺŰ�������
                            Where �ű�id = n_ָ������ű�id
                            Union All
                            Select ����, Count(����) As Num
                            From ���˹Һż�¼
                            Where Nvl(ִ��״̬, 0) = 0 And ��¼���� = 1 And ��¼״̬ = 1 And ����ʱ�� Between Trunc(Sysdate) And Sysdate And
                                  �ű� = �ºű�_In And ���� In (Select �������� From �ҺŰ������� Where �ű�id = n_ָ������ű�id)
                            Group By ����)
                     Group By ��������
                     Order By Num)
              Where Rownum = 1;
            End If;
            If n_���﷽ʽ = 3 Then
              --ƽ������
              --��ȡ��ǰ�����µ���������
              Select Count(1) Into n_������������ From �ҺŰ������� Where �ű�id = n_ָ������ű�id;
            
              Open c_ƽ������(n_ָ������ű�id);
              Loop
                Fetch c_ƽ������
                  Into r_ƽ������;
                Exit When c_ƽ������%NotFound;
                n_Index := n_Index + 1;
                --�ҵ��˶�Ӧ�ķ�������,��Ҫ�޸���һ�����ҵĵ�ǰ����Ϊ1(�������������һ�εķ�������)
                If n_�Ƿ��ҵ��������� = 1 Then
                  Update �ҺŰ�������
                  Set ��ǰ���� = 1
                  Where �ű�id = r_ƽ������.�ű�id And �������� = r_ƽ������.��������;
                  Exit;
                End If;
              
                If Nvl(r_ƽ������.��ǰ����, 0) = 1 Then
                  v_�¾������� := r_ƽ������.��������;
                  Update �ҺŰ�������
                  Set ��ǰ���� = 0
                  Where �ű�id = r_ƽ������.�ű�id And �������� = r_ƽ������.��������;
                  n_�Ƿ��ҵ��������� := 1;
                End If;
              
                If n_������������ = 1 And n_�Ƿ��ҵ��������� = 1 Then
                  n_�Ƿ��ҵ��������� := 2;
                  Exit;
                End If;
                If n_������������ > 1 And n_�Ƿ��ҵ��������� = 1 Then
                  --�α��Ѿ��������,������ӵ�һ�����ݿ�ʼ�޸ı�ʶ
                  If n_Index >= n_������������ Then
                    n_�Ƿ��ҵ��������� := 2;
                    Exit;
                  End If;
                End If;
              End Loop;
              Close c_ƽ������;
              --��������ֵ
              n_Index := 0;
              --��һ�η���
              If Nvl(v_�¾�������, ' ') = ' ' Or v_�¾������� Is Null Then
                Open c_ƽ������(n_ָ������ű�id);
                Loop
                  Fetch c_ƽ������
                    Into r_ƽ������;
                  Exit When c_ƽ������%NotFound;
                  n_Index := n_Index + 1;
                
                  If n_�Ƿ��ҵ��������� = 1 Then
                    Update �ҺŰ�������
                    Set ��ǰ���� = 1
                    Where �ű�id = r_ƽ������.�ű�id And �������� = r_ƽ������.��������;
                    Exit;
                  End If;
                
                  Update �ҺŰ�������
                  Set ��ǰ���� = 0
                  Where �ű�id = r_ƽ������.�ű�id And �������� = r_ƽ������.��������;
                  v_�¾������� := r_ƽ������.��������;
                
                  n_�Ƿ��ҵ��������� := 1;
                  If n_������������ = 1 And n_�Ƿ��ҵ��������� = 1 Then
                    n_�Ƿ��ҵ��������� := 2;
                    Exit;
                  End If;
                
                  If n_������������ > 1 And n_�Ƿ��ҵ��������� = 1 Then
                    --�α��Ѿ��������,������ӵ�һ�����ݿ�ʼ�޸ı�ʶ
                    If n_Index >= n_������������ Then
                      n_�Ƿ��ҵ��������� := 2;
                      Exit;
                    End If;
                  End If;
                
                End Loop;
                Close c_ƽ������;
              End If;
            
              If n_�Ƿ��ҵ��������� = 2 Then
                Open c_ƽ������(n_ָ������ű�id);
                Loop
                  Fetch c_ƽ������
                    Into r_ƽ������;
                  Exit When c_ƽ������%NotFound;
                  Update �ҺŰ�������
                  Set ��ǰ���� = 1
                  Where �ű�id = r_ƽ������.�ű�id And �������� = r_ƽ������.��������;
                  Exit;
                End Loop;
                Close c_ƽ������;
              End If;
            End If;
          Exception
            When Others Then
              v_�¾������� := '';
          End;
        End If;
      
        --���²�����Ϣ�ľ������Һ�״̬
        Update ������Ϣ Set �������� = v_�¾�������, ����״̬ = 1 Where ����id = n_����id And ����״̬ In (1, 2);
      
        --���α�
        Open c_Bill(v_No);
        Loop
          Fetch c_Bill
            Into r_Bill;
          Exit When c_Bill%NotFound;
          If r_Bill.��� = 1 Then
            --��Ҫȷ���Ƿ�ԤԼ�Һ�
            --1.�����ԤԼ�ҺŲ����ĹҺż�¼,����Ҫ����Լ��
            --2.�����ԤԼ�Һ�Ϊ���յĹҺż�¼,����Ҫ���ѹ����������ѽ�����
            --3.����������Һ�,��ֻ���ѹ���
            --�ָ���ǰ�ĹҺŻ���
            Update ���˹ҺŻ���
            Set �ѹ��� = Nvl(�ѹ���, 0) + Decode(n_�Һ�״̬, 0, -1, 2, -1, 0), �����ѽ��� = Nvl(�����ѽ���, 0) + Decode(n_�Һ�״̬, 2, -1, 0),
                ��Լ�� = Nvl(��Լ��, 0) + Decode(n_�Һ�״̬, 0, 0, -1)
            Where ���� = Trunc(r_Bill.����ʱ��) And Nvl(����id, 0) = Nvl(r_Bill.ִ�в���id, 0) And
                  Nvl(��Ŀid, 0) = Nvl(r_Bill.�շ�ϸĿid, 0) And (���� = r_Bill.���㵥λ Or ���� Is Null);
            If Sql%RowCount = 0 Then
              Insert Into ���˹ҺŻ���
                (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, ��Լ��, �����ѽ���)
              Values
                (Trunc(r_Bill.����ʱ��), r_Bill.ִ�в���id, r_Bill.�շ�ϸĿid, ԭҽ������_In, Decode(ԭҽ��id_In, 0, Null, ԭҽ��id_In),
                 r_Bill.���㵥λ, 0, 0, 0);
            End If;
          
            ----Ȼ���ٸ��¹ҺŻ���
            Update ���˹ҺŻ���
            Set �ѹ��� = Nvl(�ѹ���, 0) + Decode(n_�Һ�״̬, 0, 1, 2, 1, 0), �����ѽ��� = Nvl(�����ѽ���, 0) + Decode(n_�Һ�״̬, 2, 1, 0),
                ��Լ�� = Nvl(��Լ��, 0) + Decode(n_�Һ�״̬, 0, 0, 1)
            Where ���� = Trunc(r_Bill.����ʱ��) And Nvl(����id, 0) = �¿���id_In And Nvl(��Ŀid, 0) = Nvl(r_Bill.�շ�ϸĿid, 0) And
                 
                  (���� = �ºű�_In Or ���� Is Null);
            If Sql%RowCount = 0 Then
              Insert Into ���˹ҺŻ���
                (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, ��Լ��, �����ѽ���)
              Values
                (Trunc(r_Bill.����ʱ��), �¿���id_In, r_Bill.�շ�ϸĿid, ��ҽ������_In, Decode(��ҽ��id_In, 0, Null, ��ҽ��id_In), �ºű�_In,
                 Decode(n_�Һ�״̬, 0, 1, 2, 1, 0), Decode(n_�Һ�״̬, 0, 0, 1), Decode(n_�Һ�״̬, 2, 1, 0));
            End If;
          End If;
        
          ---���¹Һż�¼
          If n_�Һ�״̬ = 1 Then
            --ԤԼ
            Update ������ü�¼
            Set ִ�в���id = �¿���id_In, ���˿���id = �¿���id_In, ���㵥λ = �ºű�_In, ��ҩ���� = n_ԭ���,
                --���˲���id = ����id_In,
                ִ���� = ��ҽ������_In, ִ��״̬ = 0, ִ��ʱ�� = Null
            Where ID = r_Bill.Id;
          Else
            --�ҺŻ����
            Update ������ü�¼
            Set ִ�в���id = �¿���id_In, ���˿���id = �¿���id_In, ���㵥λ = �ºű�_In, ��ҩ���� = v_�¾�������,
                --���˲���id = ����id_In,
                ִ���� = ��ҽ������_In, ִ��״̬ = 0, ִ��ʱ�� = Null
            Where ID = r_Bill.Id;
          End If;
        
          --���²��˹Һż�¼
          If r_Bill.��� = 1 Then
            v_Temp := Zl_Identity(1);
            Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_Temp From Dual;
            Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into v_����Ա��� From Dual;
            Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_����Ա���� From Dual;
            Begin
              Select ID Into n_ҽ��id From ��Ա�� Where ���� = ��ҽ������_In And Rownum < 2;
            Exception
              When Others Then
                n_ҽ��id := Null;
            End;
            Select ����䶯��¼_Id.Nextval Into n_�䶯id From Dual;
            b_Message.Zlhis_Regist_005(r_Bill.No, n_�䶯id, 2);
            Zl_����䶯��¼_Insert(r_Bill.No, 1, '��������', v_����Ա����, v_����Ա���, �ºű�_In, �¿���id_In, Null, n_ҽ��id, ��ҽ������_In, v_�¾�������,
                             n_ԭ���, Null, n_�䶯id);
            --�޸Ķ�����Ϣ
            Update �ŶӽкŶ���
            Set ҽ������ = ��ҽ������_In, ���� = v_�¾�������
            Where ҵ��id = n_ҵ��id And ҵ������ = 0;
          
            Update ���˹Һż�¼
            Set ִ�в���id = �¿���id_In, �ű� = �ºű�_In, ���� = v_�¾�������, ִ���� = ��ҽ������_In, ִ��״̬ = 0, ִ��ʱ�� = Null, ���� = n_ԭ���
            Where NO = r_Bill.No;
            --�޸ĹҺ����״̬
            If n_ԭ��� Is Not Null Then
              --1.�ָ���ǰ�Һ����״̬
              Delete �Һ����״̬ Where ���� = d_ԭ�������� And ��� = n_ԭ��� And ���� = ԭ�ű�_In;
              --2.�������ź�Һ����״̬
              Insert Into �Һ����״̬
                (����, ����, ���, ����Ա����, ״̬, ԤԼ, �Ǽ�ʱ��)
              Values
                (�ºű�_In, d_ԭ��������, n_ԭ���, ����Ա����_In, Decode(n_�Һ�״̬, 1, 2, 1), Decode(n_�Һ�״̬, 0, 0, 1), Sysdate);
            End If;
          End If;
        End Loop;
        Close c_Bill;
      End Loop;
    End If;
  Else
    --������Ű�ģʽ
    --����Ƿ���ڸùҺż�¼
    If Nos_In Is Not Null Then
      v_Nos := Nos_In || '|';
      While v_Nos Is Not Null Loop
        --��ʼ������
        n_����id           := 0;
        n_ԭ���           := 0;
        d_ԭ��������       := Null;
        n_�Ƿ��ѱ��ҳ�     := 0;
        n_��¼����         := 0;
        n_ԤԼ             := 0;
        n_�Һ�״̬         := 0;
        v_�¾�������       := '';
        n_���﷽ʽ         := 0;
        n_ָ������ű�id   := 0;
        v_�ֶ�������       := '';
        v_�ŶӺ���         := '';
        n_ҵ��id           := 0;
        n_������������     := 0;
        n_�Һ����ɶ���     := 0;
        n_ԤԼ���ɶ���     := 0;
        n_�Ƿ��ҵ��������� := 0;
        n_Index            := 0;
      
        v_No  := Substr(v_Nos, 1, Instr(v_Nos, '|') - 1);
        v_Nos := Substr(v_Nos, Instr(v_Nos, '|') + 1);
        --����Ƿ���ڸùҺż�¼
        Begin
          Select a.Id, a.����id, a.����, Nvl(b.��ʼʱ��, Nvl(a.ԤԼʱ��, a.����ʱ��)), a.��¼����, Nvl(a.ԤԼ, 0), a.�����¼id
          Into n_ҵ��id, n_����id, n_ԭ���, d_ԭ��������, n_��¼����, n_ԤԼ, n_�����¼id
          From ���˹Һż�¼ A, �ٴ�������ſ��� B
          Where NO = v_No And ��¼���� In (1, 2) And ��¼״̬ = 1 And a.���� = b.���(+) And a.�����¼id = b.��¼id(+);
        Exception
          When Others Then
            Null;
        End;
        If n_����id = 0 Then
          v_Error := 'û���ҵ����˵ĹҺ���Ϣ';
          Raise Err_Custom;
        End If;
        --�жϵ�ǰ�Һ�״̬
        n_�Һ�״̬ := 0; --�����Һ�
        If n_��¼���� = 1 And n_ԤԼ = 1 Then
          n_�Һ�״̬ := 2; --ԤԼ����
        End If;
        If n_��¼���� = 2 And n_ԤԼ = 1 Then
          n_�Һ�״̬ := 1; --ԤԼ
        End If;
      
        --��黻�ŵ��ºű��Ƿ��ѱ��ҳ�
        Begin
          Select a.�Һ�״̬
          Into n_�Ƿ��ѱ��ҳ�
          From �ٴ�������ſ��� A
          Where a.��¼id = �³���id_In And a.��� = n_ԭ���;
        Exception
          When Others Then
            n_�Ƿ��ѱ��ҳ� := 0;
        End;
        If n_�Ƿ��ѱ��ҳ� > 0 Then
          v_Error := 'Ҫ���ĺű��ѱ��ҳ�';
          Raise Err_Custom;
        End If;
        --ԤԼ���յ�����½��з������ҵĻ�ȡ
        If n_�Һ�״̬ = 2 Then
          --��ȡ�ºű�����
          --˵��:ԤԼ������£�����Ҫ�����˲��û�ȡ��������
          --     ���յ������,��Ҫ���з���,�����Ҫ��ȡ��������
          --��ȡ���﷽ʽ
          Begin
            Select Nvl(���﷽ʽ, 0) Into n_���﷽ʽ From �ٴ������¼ Where ID = �³���id_In;
          Exception
            When Others Then
              n_���﷽ʽ := 0;
          End;
        
          Begin
            If n_���﷽ʽ = 0 Then
              --������
              v_�¾������� := '';
            End If;
            If n_���﷽ʽ = 1 Then
              --ָ������
              Select b.����
              Into v_�¾�������
              From �ٴ��������Ҽ�¼ A, �������� B
              Where a.����id = b.Id And a.��¼id = �³���id_In;
            End If;
            If n_���﷽ʽ = 2 Then
              --��̬����
              Select ��������
              Into v_�¾�������
              From (Select ��������, Sum(Num) As Num
                     From (Select b.���� As ��������, 0 As Num
                            From �ٴ��������Ҽ�¼ A, �������� B
                            Where a.��¼id = �³���id_In
                            Union All
                            Select ����, Count(����) As Num
                            From ���˹Һż�¼
                            Where Nvl(ִ��״̬, 0) = 0 And ��¼���� = 1 And ��¼״̬ = 1 And ����ʱ�� Between Trunc(Sysdate) And Sysdate And
                                  �ű� = �ºű�_In And ���� In (Select b.����
                                                         From �ٴ��������Ҽ�¼ A, �������� B
                                                         Where a.����id = b.Id And a.��¼id = �³���id_In)
                            Group By ����)
                     Group By ��������
                     Order By Num)
              Where Rownum = 1;
            End If;
            If n_���﷽ʽ = 3 Then
              --ƽ������
              --��ȡ��ǰ�����µ���������
              Select Count(1) Into n_������������ From �ٴ��������Ҽ�¼ Where ��¼id = �³���id_In;
            
              Open c_����ƽ������;
              Loop
                Fetch c_����ƽ������
                  Into r_����ƽ������;
                Exit When c_����ƽ������%NotFound;
                n_Index := n_Index + 1;
                --�ҵ��˶�Ӧ�ķ�������,��Ҫ�޸���һ�����ҵĵ�ǰ����Ϊ1(�������������һ�εķ�������)
                If n_�Ƿ��ҵ��������� = 1 Then
                  Update �ٴ��������Ҽ�¼
                  Set ��ǰ���� = 1
                  Where ��¼id = r_����ƽ������.��¼id And ����id = r_����ƽ������.����id;
                  Exit;
                End If;
              
                If Nvl(r_ƽ������.��ǰ����, 0) = 1 Then
                  v_�¾������� := r_ƽ������.��������;
                  Update �ٴ��������Ҽ�¼
                  Set ��ǰ���� = 0
                  Where ��¼id = r_����ƽ������.��¼id And ����id = r_����ƽ������.����id;
                  n_�Ƿ��ҵ��������� := 1;
                End If;
              
                If n_������������ = 1 And n_�Ƿ��ҵ��������� = 1 Then
                  n_�Ƿ��ҵ��������� := 2;
                  Exit;
                End If;
                If n_������������ > 1 And n_�Ƿ��ҵ��������� = 1 Then
                  --�α��Ѿ��������,������ӵ�һ�����ݿ�ʼ�޸ı�ʶ
                  If n_Index >= n_������������ Then
                    n_�Ƿ��ҵ��������� := 2;
                    Exit;
                  End If;
                End If;
              End Loop;
              Close c_����ƽ������;
              --��������ֵ
              n_Index := 0;
              --��һ�η���
              If Nvl(v_�¾�������, ' ') = ' ' Or v_�¾������� Is Null Then
                Open c_����ƽ������;
                Loop
                  Fetch c_����ƽ������
                    Into r_����ƽ������;
                  Exit When c_����ƽ������%NotFound;
                  n_Index := n_Index + 1;
                
                  If n_�Ƿ��ҵ��������� = 1 Then
                    Update �ٴ��������Ҽ�¼
                    Set ��ǰ���� = 1
                    Where ��¼id = r_����ƽ������.��¼id And ����id = r_����ƽ������.����id;
                    Exit;
                  End If;
                
                  Update �ٴ��������Ҽ�¼
                  Set ��ǰ���� = 0
                  Where ��¼id = r_����ƽ������.��¼id And ����id = r_����ƽ������.����id;
                  v_�¾������� := r_����ƽ������.����id;
                
                  n_�Ƿ��ҵ��������� := 1;
                  If n_������������ = 1 And n_�Ƿ��ҵ��������� = 1 Then
                    n_�Ƿ��ҵ��������� := 2;
                    Exit;
                  End If;
                
                  If n_������������ > 1 And n_�Ƿ��ҵ��������� = 1 Then
                    --�α��Ѿ��������,������ӵ�һ�����ݿ�ʼ�޸ı�ʶ
                    If n_Index >= n_������������ Then
                      n_�Ƿ��ҵ��������� := 2;
                      Exit;
                    End If;
                  End If;
                
                End Loop;
                Close c_����ƽ������;
              End If;
            
              If n_�Ƿ��ҵ��������� = 2 Then
                Open c_����ƽ������;
                Loop
                  Fetch c_����ƽ������
                    Into r_����ƽ������;
                  Exit When c_����ƽ������%NotFound;
                  Update �ٴ��������Ҽ�¼
                  Set ��ǰ���� = 1
                  Where ��¼id = r_����ƽ������.��¼id And ����id = r_����ƽ������.����id;
                  Exit;
                End Loop;
                Close c_����ƽ������;
              End If;
            End If;
          Exception
            When Others Then
              v_�¾������� := '';
          End;
        End If;
      
        --���²�����Ϣ�ľ������Һ�״̬
        Update ������Ϣ
        Set �������� =
             (Select ���� From �������� Where ID = v_�¾�������), ����״̬ = 1
        Where ����id = n_����id And ����״̬ In (1, 2);
      
        --���α�
        Open c_Bill(v_No);
        Loop
          Fetch c_Bill
            Into r_Bill;
          Exit When c_Bill%NotFound;
          If r_Bill.��� = 1 Then
            --��Ҫȷ���Ƿ�ԤԼ�Һ�
            --1.�����ԤԼ�ҺŲ����ĹҺż�¼,����Ҫ����Լ��
            --2.�����ԤԼ�Һ�Ϊ���յĹҺż�¼,����Ҫ���ѹ����������ѽ�����
            --3.����������Һ�,��ֻ���ѹ���
            --�ָ���ǰ�ĹҺŻ���
            Update ���˹ҺŻ���
            Set �ѹ��� = Nvl(�ѹ���, 0) + Decode(n_�Һ�״̬, 0, -1, 2, -1, 0), �����ѽ��� = Nvl(�����ѽ���, 0) + Decode(n_�Һ�״̬, 2, -1, 0),
                ��Լ�� = Nvl(��Լ��, 0) + Decode(n_�Һ�״̬, 0, 0, -1)
            Where ���� = Trunc(r_Bill.����ʱ��) And Nvl(ҽ��id, 0) = Nvl(ԭҽ��id_In, 0) And Nvl(ҽ������, '-') = Nvl(ԭҽ������_In, '-') And
                  Nvl(����id, 0) = Nvl(r_Bill.ִ�в���id, 0) And Nvl(��Ŀid, 0) = Nvl(r_Bill.�շ�ϸĿid, 0) And
                  (���� = r_Bill.���㵥λ Or ���� Is Null);
            If Sql%RowCount = 0 Then
              Insert Into ���˹ҺŻ���
                (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, ��Լ��, �����ѽ���)
              Values
                (Trunc(r_Bill.����ʱ��), r_Bill.ִ�в���id, r_Bill.�շ�ϸĿid, ԭҽ������_In, Decode(ԭҽ��id_In, 0, Null, ԭҽ��id_In),
                 r_Bill.���㵥λ, 0, 0, 0);
            End If;
          
            Update �ٴ������¼
            Set �ѹ��� = Nvl(�ѹ���, 0) + Decode(n_�Һ�״̬, 0, -1, 2, -1, 0), �����ѽ��� = Nvl(�����ѽ���, 0) + Decode(n_�Һ�״̬, 2, -1, 0),
                ��Լ�� = Nvl(��Լ��, 0) + Decode(n_�Һ�״̬, 0, 0, -1)
            Where ID = ԭ����id_In;
          
            ----Ȼ���ٸ��¹ҺŻ���
            Update ���˹ҺŻ���
            Set �ѹ��� = Nvl(�ѹ���, 0) + Decode(n_�Һ�״̬, 0, 1, 2, 1, 0), �����ѽ��� = Nvl(�����ѽ���, 0) + Decode(n_�Һ�״̬, 2, 1, 0),
                ��Լ�� = Nvl(��Լ��, 0) + Decode(n_�Һ�״̬, 0, 0, 1)
            Where ���� = Trunc(r_Bill.����ʱ��) And Nvl(ҽ��id, 0) = Nvl(��ҽ��id_In, 0) And Nvl(ҽ������, '-') = Nvl(��ҽ������_In, '-') And
                  Nvl(����id, 0) = �¿���id_In And Nvl(��Ŀid, 0) = Nvl(r_Bill.�շ�ϸĿid, 0) And (���� = �ºű�_In Or ���� Is Null);
            If Sql%RowCount = 0 Then
              Insert Into ���˹ҺŻ���
                (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, ��Լ��, �����ѽ���)
              Values
                (Trunc(r_Bill.����ʱ��), �¿���id_In, r_Bill.�շ�ϸĿid, ��ҽ������_In, Decode(��ҽ��id_In, 0, Null, ��ҽ��id_In), �ºű�_In,
                 Decode(n_�Һ�״̬, 0, 1, 2, 1, 0), Decode(n_�Һ�״̬, 0, 0, 1), Decode(n_�Һ�״̬, 2, 1, 0));
            End If;
            Update �ٴ������¼
            Set �ѹ��� = Nvl(�ѹ���, 0) + Decode(n_�Һ�״̬, 0, 1, 2, 1, 0), �����ѽ��� = Nvl(�����ѽ���, 0) + Decode(n_�Һ�״̬, 2, 1, 0),
                ��Լ�� = Nvl(��Լ��, 0) + Decode(n_�Һ�״̬, 0, 0, 1)
            Where ID = �³���id_In;
          End If;
        
          ---���¹Һż�¼
          If n_�Һ�״̬ = 1 Then
            --ԤԼ
            Update ������ü�¼
            Set ִ�в���id = �¿���id_In, ���˿���id = �¿���id_In, ���㵥λ = �ºű�_In, ��ҩ���� = n_ԭ���,
                --���˲���id = ����id_In,
                ִ���� = ��ҽ������_In, ִ��״̬ = 0, ִ��ʱ�� = Null
            Where ID = r_Bill.Id;
          Else
            --�ҺŻ����
            Update ������ü�¼
            Set ִ�в���id = �¿���id_In, ���˿���id = �¿���id_In, ���㵥λ = �ºű�_In, ��ҩ���� = v_�¾�������,
                --���˲���id = ����id_In,
                ִ���� = ��ҽ������_In, ִ��״̬ = 0, ִ��ʱ�� = Null
            Where ID = r_Bill.Id;
          End If;
        
          --���²��˹Һż�¼
          If r_Bill.��� = 1 Then
            v_Temp := Zl_Identity(1);
            Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_Temp From Dual;
            Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into v_����Ա��� From Dual;
            Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_����Ա���� From Dual;
            Begin
              Select ID Into n_ҽ��id From ��Ա�� Where ���� = ��ҽ������_In And Rownum < 2;
            Exception
              When Others Then
                n_ҽ��id := Null;
            End;
            Select ����䶯��¼_Id.Nextval Into n_�䶯id From Dual;
            b_Message.Zlhis_Regist_005(r_Bill.No, n_�䶯id, 2);
            Zl_����䶯��¼_Insert(r_Bill.No, 1, '��������', v_����Ա����, v_����Ա���, �ºű�_In, �¿���id_In, Null, n_ҽ��id, ��ҽ������_In, v_�¾�������,
                             n_ԭ���, Null, n_�䶯id);
            --�޸Ķ�����Ϣ
            Update �ŶӽкŶ���
            Set ҽ������ = ��ҽ������_In, ���� = v_�¾�������
            Where ҵ��id = n_ҵ��id And ҵ������ = 0;
          
            Update ���˹Һż�¼
            Set ִ�в���id = �¿���id_In, �ű� = �ºű�_In, ���� = v_�¾�������, ִ���� = ��ҽ������_In, ִ��״̬ = 0, ִ��ʱ�� = Null, ���� = n_ԭ���,
                �����¼id = �³���id_In
            Where NO = r_Bill.No;
            --�޸ĹҺ����״̬
            If n_ԭ��� Is Not Null Then
              --1.�ָ���ǰ�Һ����״̬
              Update �ٴ�������ſ���
              Set �Һ�״̬ = 0, ����Ա���� = Null, ����վip = Null, ����վ���� = Null
              Where ��¼id = ԭ����id_In And ��� = n_ԭ���;
              --2.�������ź�Һ����״̬
              Update �ٴ�������ſ���
              Set �Һ�״̬ = Decode(n_�Һ�״̬, 1, 2, 1), ����Ա���� = v_����Ա����, ����վip = Null, ����վ���� = Null
              Where ��¼id = �³���id_In And ��� = n_ԭ���;
            End If;
          End If;
        End Loop;
        Close c_Bill;
      End Loop;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˹Һż�¼_��������;
/



Create Or Replace Procedure Zl_���˽������
(
  ����id_In ������Ϣ.����id%Type,
  No_In     ���˹Һż�¼.No%Type,
  ����_In   ���˹Һż�¼.����%Type := Null,
  ִ����_In ���˹Һż�¼.ִ����%Type := Null,
  ժҪ_In   ���˹Һż�¼.ժҪ%Type := Null,
  ��ʿ_In   ���˹Һż�¼.���ӱ�־%Type := Null
) As
  v_�Һ�id     ���˹Һż�¼.Id%Type;
  v_ִ�в���id ���˹Һż�¼.ִ�в���id%Type;
  v_����ʱ��   ���˹Һż�¼.ִ��ʱ��%Type;
  v_���ʱ��   ���˹Һż�¼.ִ��ʱ��%Type;
  v_���       Varchar2(100);
  n_����id     ��������.Id%Type;
Begin
  v_���ʱ�� := Sysdate;

  Update ������Ϣ Set ����״̬ = 0 Where ����id = ����id_In And ����״̬ In (1, 2); --1-�ȴ�����,2-���ھ���;
  Begin
    Select ID Into n_����id From �������� Where ���� = ����_In;
  Exception
    When Others Then
      Null;
  End;
  --ִ��ʱ�䱣���˹Һż�¼һ��
  Update ������ü�¼
  Set ִ���� = Decode(ִ����_In, Null, ִ����, ִ����_In), ִ��״̬ = 1, ��ҩ���� = ����_In, ���� = Decode(ժҪ_In, Null, ����, ժҪ_In), Ӥ���� = ��ʿ_In
  Where NO = No_In And ��¼���� = 4 And ��¼״̬ In (1, 3) And Nvl(ִ��״̬, 0) In (0, 2);

  --���˹Һż�¼
  Update ���˹Һż�¼
  Set ִ���� = Decode(ִ����_In, Null, ִ����, ִ����_In), ִ��״̬ = 1, ���� = ����_In, ���ʱ�� = v_���ʱ��, ժҪ = Decode(ժҪ_In, Null, ժҪ, ժҪ_In),
      ���ӱ�־ = ��ʿ_In
  Where NO = No_In And Nvl(ִ��״̬, 0) In (0, 2) And ��¼״̬ = 1 And ��¼���� = 1
  Returning ID, ִ�в���id, ִ��ʱ��, Decode(����, 1, '����', Decode(����, 1, '����', '����')) Into v_�Һ�id, v_ִ�в���id, v_����ʱ��, v_���;

  If v_�Һ�id Is Not Null Then
    --��������ʱ�����ܱ��ε���û�н���Update��������û�з���ֵ
    --�����,�ŶӽкŸ���Ϊ���
    Update �ŶӽкŶ��� Set �Ŷ�״̬ = 4 Where ҵ������ = 0 And ҵ��id = v_�Һ�id;
  
    --����ʱ������
    Zl_���Ӳ���ʱ��_Insert(����id_In, v_�Һ�id, 1, v_���, v_ִ�в���id, ִ����_In, v_����ʱ��, v_���ʱ��);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˽������;
/



Create Or Replace Procedure Zl_����ԤԼ�Һ�_Clear As
  v_ԤԼ����     Number;
  v_����Ա����   ��Ա��.���� %Type;
  v_����Ա���   ��Ա��.���%Type;
  n_�Һ�id       ���˹Һż�¼.Id%Type;
  n_�Һ��Ű�ģʽ Number(3);
  n_�����¼id   �ٴ������¼.Id%Type;
  Cursor c_Clear Is
    Select a.No, a.����ʱ��, b.����id, b.��Ŀid, b.ҽ������, b.ҽ��id, b.����
    From ������ü�¼ A, �ҺŰ��� B
    Where a.���㵥λ = b.���� And a.��¼���� = 4 And a.��¼״̬ = 0 And a.��� = 1 And �Ǽ�ʱ�� >= Sysdate - v_ԤԼ���� And
          ����ʱ�� < Trunc(Sysdate);
  Cursor c_����clear Is
    Select b.Id, a.No, a.����ʱ��, d.����id, b.��Ŀid, b.ҽ������, b.ҽ��id, d.����, c.����
    From ������ü�¼ A, �ٴ������¼ B, ���˹Һż�¼ C, �ٴ������Դ D
    Where a.No = c.No And c.�����¼id = b.Id And b.��Դid = d.Id And a.��¼���� = 4 And a.��¼״̬ = 0 And a.��� = 1 And
          a.�Ǽ�ʱ�� >= Sysdate - v_ԤԼ���� And a.����ʱ�� < Trunc(Sysdate);
Begin
  Select Zl_To_Number(Nvl(zl_GetSysParameter(66), '15')) Into v_ԤԼ���� From Dual;
  Begin
    Select b.����, b.���
    Into v_����Ա����, v_����Ա���
    From �ϻ���Ա�� A, ��Ա�� B
    Where a.��Աid = b.Id And a.�û��� = Upper(User);
  Exception
    When Others Then
      Null;
  End;
  n_�Һ��Ű�ģʽ := Zl_To_Number(Nvl(zl_GetSysParameter(253), 0));
  If n_�Һ��Ű�ģʽ = 0 Then
    For r_Clear In c_Clear Loop
      Update ���˹ҺŻ���
      Set ��Լ�� = Nvl(��Լ��, 0) - 1
      Where ���� = Trunc(r_Clear.����ʱ��) And ����id = r_Clear.����id And ��Ŀid = r_Clear.��Ŀid And
            Nvl(ҽ������, 'ҽ��') = Nvl(r_Clear.ҽ������, 'ҽ��') And Nvl(ҽ��id, 0) = Nvl(r_Clear.ҽ��id, 0) And
            (���� = r_Clear.���� Or ���� Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into ���˹ҺŻ���
          (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, ��Լ��)
        Values
          (Trunc(r_Clear.����ʱ��), r_Clear.����id, r_Clear.��Ŀid, r_Clear.ҽ������, Decode(r_Clear.ҽ��id, 0, Null, r_Clear.ҽ��id),
           r_Clear.����, -1);
      End If;
      --ɾ��������ü�¼
      Delete From ������ü�¼ Where NO = r_Clear.No And ��¼���� = 4 And ��¼״̬ = 0;
      Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
      Update ���˹Һż�¼ Set ��¼״̬ = 3 Where NO = r_Clear.No;
      Insert Into ���˹Һż�¼
        (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ԤԼʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���,
         ����Ա����, ����, ����, ����, ԤԼ, ժҪ, ������ˮ��, ����˵��, ������λ, ҽ�Ƹ��ʽ)
        Select n_�Һ�id, NO, ��¼����, 2, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ԤԼʱ��, �Ǽ�ʱ��, ����ʱ��,
               v_����Ա���, v_����Ա����, ����, ����, ����, ԤԼ, ժҪ, ������ˮ��, ����˵��, ������λ, ҽ�Ƹ��ʽ
        From ���˹Һż�¼
        Where NO = r_Clear.No And ��¼״̬ = 3;
    End Loop;
  Else
    --������Ű�ģʽ
    For r_����clear In c_����clear Loop
      Update ���˹ҺŻ���
      Set ��Լ�� = Nvl(��Լ��, 0) - 1
      Where ���� = Trunc(r_����clear.����ʱ��) And ����id = r_����clear.����id And ��Ŀid = r_����clear.��Ŀid And
            Nvl(ҽ������, 'ҽ��') = Nvl(r_����clear.ҽ������, 'ҽ��') And Nvl(ҽ��id, 0) = Nvl(r_����clear.ҽ��id, 0) And
            (���� = r_����clear.���� Or ���� Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into ���˹ҺŻ���
          (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, ��Լ��)
        Values
          (Trunc(r_����clear.����ʱ��), r_����clear.����id, r_����clear.��Ŀid, r_����clear.ҽ������,
           Decode(r_����clear.ҽ��id, 0, Null, r_����clear.ҽ��id), r_����clear.����, -1);
      End If;
      Update �ٴ������¼ Set ��Լ�� = Nvl(��Լ��, 0) - 1 Where ID = r_����clear.Id;
      Update �ٴ�������ſ���
      Set �Һ�״̬ = 0, ����Ա���� = Null, ����վip = Null, ����վ���� = Null
      Where ��¼id = r_����clear.Id And (��� = r_����clear.���� Or ��ע = r_����clear.����);
      --������ü�¼
      Update ������ü�¼ Set ��¼״̬ = 3 Where NO = r_����clear.No And ��¼���� = 4 And ��¼״̬ = 0;
      Insert Into ������ü�¼
        (ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, ���ʵ�id, ����id, ҽ�����, �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��, ���ʽ, ���˿���id, �ѱ�,
         �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��,
         ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���, ����Ա����, ����id, ���ʽ��, ���մ���id, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���, �ɿ���id,
         ����״̬, ��ת��, �Һ�id, ��ҳid)
        Select ���˷��ü�¼_Id.Nextval, ��¼����, NO, ʵ��Ʊ��, 2, ���, ��������, �۸񸸺�, ���ʵ�id, ����id, ҽ�����, �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��,
               ���ʽ, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, -1 * ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, -1 * Ӧ�ս��,
               -1 * ʵ�ս��, ������, ��������id, ������, ����ʱ��, Sysdate, ִ�в���id, ִ����, -1, ִ��ʱ��, ����, v_����Ա���, v_����Ա����, Null, Null,
               ���մ���id, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���, �ɿ���id, ����״̬, ��ת��, �Һ�id, ��ҳid
        From ������ü�¼
        Where NO = r_����clear.No And ��¼���� = 4 And ��¼״̬ = 3;
    
      Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
      Update ���˹Һż�¼ Set ��¼״̬ = 3 Where NO = r_����clear.No;
      Insert Into ���˹Һż�¼
        (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ԤԼʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���,
         ����Ա����, ����, ����, ����, ԤԼ, ժҪ, ������ˮ��, ����˵��, ������λ, ҽ�Ƹ��ʽ)
        Select n_�Һ�id, NO, ��¼����, 2, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ԤԼʱ��, �Ǽ�ʱ��, ����ʱ��,
               v_����Ա���, v_����Ա����, ����, ����, ����, ԤԼ, ժҪ, ������ˮ��, ����˵��, ������λ, ҽ�Ƹ��ʽ
        From ���˹Һż�¼
        Where NO = r_����clear.No And ��¼״̬ = 3;
    End Loop;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ԤԼ�Һ�_Clear;
/



Create Or Replace Procedure Zl_����ԤԼ�Һ�_Defer
(
  �ű�_In       ������ü�¼.��ҩ����%Type,
  ԤԼ����_In   ������ü�¼.����ʱ��%Type,
  ��������_In   ������ü�¼.����ʱ��%Type,
  ����Ա����_In �Һ����״̬.����Ա����%Type,
  ��¼id_In     �ٴ������¼.Id%Type := Null
) As
  v_Do     Number(1);
  v_ҽ��   �ҺŰ���.ҽ������%Type;
  v_ҽ��id �ҺŰ���.ҽ��id%Type;
  v_����   Number;
  n_��¼id �ٴ������¼.Id%Type;
Begin
  If ��¼id_In Is Null Then
    v_���� := Trunc(��������_In) - Trunc(ԤԼ����_In);
    For c_Fee In (Select Distinct NO, ��ҩ���� ����, ִ�в���id, �շ�ϸĿid
                  From ������ü�¼
                  Where ��¼���� = 4 And ��¼״̬ = 0 And ��� = 1 And ���㵥λ = �ű�_In And ����ʱ�� Between Trunc(ԤԼ����_In) And
                        Trunc(ԤԼ����_In + 1) - 1 / 24 / 60 / 60) Loop
      v_Do := 1;
      --�Һ����״̬
      If Not c_Fee.���� Is Null Then
        Begin
          Update �Һ����״̬
          Set ���� = ���� + v_����, �Ǽ�ʱ�� = Sysdate
          Where ���� = �ű�_In And Trunc(����) = Trunc(ԤԼ����_In) And ��� = c_Fee.���� And ״̬ = 2 And ����Ա���� = ����Ա����_In;
        Exception
          --�����������������ʹ��,���ԤԼ�ҺŲ�����
          When Others Then
            --�����Ԥ����,������ֱ��ʹ��
            Update �Һ����״̬
            Set ״̬ = 2, �Ǽ�ʱ�� = Sysdate
            Where ���� = �ű�_In And Trunc(����) = Trunc(��������_In) And ��� = c_Fee.���� And ״̬ = 3 And ����Ա���� = ����Ա����_In;
            If Sql%RowCount = 0 Then
              v_Do := 0;
            Else
              Delete �Һ����״̬
              Where ���� = �ű�_In And Trunc(����) = Trunc(ԤԼ����_In) And ��� = c_Fee.���� And ״̬ = 2 And ����Ա���� = ����Ա����_In;
            End If;
        End;
      End If;
    
      If v_Do = 1 Then
        --ԤԼ��¼
        Update ������ü�¼
        Set ����ʱ�� = To_Date(To_Char(��������_In, 'yyyy-mm-dd') || To_Char(����ʱ��, ' hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
        Where ��¼���� = 4 And ��¼״̬ = 0 And NO = c_Fee.No;
        Update ���˹Һż�¼
        Set ����ʱ�� = To_Date(To_Char(��������_In, 'yyyy-mm-dd') || To_Char(����ʱ��, ' hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
        Where ��¼���� = 2 And ��¼״̬ = 1 And NO = c_Fee.No;
        --���˹ҺŻ���
        Begin
          Select ҽ������, ҽ��id Into v_ҽ��, v_ҽ��id From �ҺŰ��� Where ���� = �ű�_In;
        Exception
          When Others Then
            Null;
        End;
        Update ���˹ҺŻ���
        Set ��Լ�� = Nvl(��Լ��, 0) - 1
        Where ���� = Trunc(ԤԼ����_In) And Nvl(����id, 0) = c_Fee.ִ�в���id And Nvl(��Ŀid, 0) = c_Fee.�շ�ϸĿid And
              Nvl(ҽ������, 'ҽ��') = Nvl(v_ҽ��, 'ҽ��') And Nvl(ҽ��id, 0) = Nvl(v_ҽ��id, 0) And (���� = �ű�_In Or ���� Is Null);
      
        Update ���˹ҺŻ���
        Set ��Լ�� = Nvl(��Լ��, 0) + 1
        Where ���� = Trunc(��������_In) And Nvl(����id, 0) = c_Fee.ִ�в���id And Nvl(��Ŀid, 0) = c_Fee.�շ�ϸĿid And
              Nvl(ҽ������, 'ҽ��') = Nvl(v_ҽ��, 'ҽ��') And Nvl(ҽ��id, 0) = Nvl(v_ҽ��id, 0) And (���� = �ű�_In Or ���� Is Null);
        If Sql%RowCount = 0 Then
          Insert Into ���˹ҺŻ���
            (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, ��Լ��)
          Values
            (Trunc(��������_In), c_Fee.ִ�в���id, c_Fee.�շ�ϸĿid, v_ҽ��, Decode(v_ҽ��id, 0, Null, v_ҽ��id), �ű�_In, 1);
        End If;
      End If;
    End Loop;
  Else
    --������Ű�ģʽ
    v_���� := Trunc(��������_In) - Trunc(ԤԼ����_In);
    For c_Fee In (Select Distinct NO, ��ҩ���� ����, ִ�в���id, �շ�ϸĿid
                  From ������ü�¼
                  Where ��¼���� = 4 And ��¼״̬ = 0 And ��� = 1 And ���㵥λ = �ű�_In And ����ʱ�� Between Trunc(ԤԼ����_In) And
                        Trunc(ԤԼ����_In + 1) - 1 / 24 / 60 / 60) Loop
      v_Do := 1;
      --�Һ����״̬
      If Not c_Fee.���� Is Null Then
        Select c.Id
        Into n_��¼id
        From �ٴ������¼ A, �ٴ������Դ B, �ٴ������¼ C
        Where a.Id = ��¼id_In And a.��Դid = b.Id And b.Id = c.��Դid And c.�������� = a.�������� + v_����;
      
        Update �ٴ�������ſ���
        Set �Һ�״̬ = 2, ����Ա���� = ����Ա����_In
        Where ��¼id = n_��¼id And (��� = c_Fee.���� Or ��ע = c_Fee.����) And �Һ�״̬ = 0;
      
        If Sql%RowCount = 0 Then
          --�����Ԥ����,������ֱ��ʹ��
          Update �ٴ�������ſ���
          Set �Һ�״̬ = 2, ����Ա���� = ����Ա����_In
          Where ��¼id = n_��¼id And (��� = c_Fee.���� Or ��ע = c_Fee.����) And �Һ�״̬ = 3 And ����Ա���� = ����Ա����_In;
          If Sql%RowCount = 0 Then
            v_Do := 0;
          End If;
        End If;
      
        Update �ٴ�������ſ���
        Set �Һ�״̬ = 0, ����Ա���� = Null, ����վip = Null, ����վ���� = Null
        Where ��¼id = ��¼id_In And (��� = c_Fee.���� Or ��ע = c_Fee.����);
      End If;
    
      If v_Do = 1 Then
        --ԤԼ��¼
        Update ������ü�¼
        Set ����ʱ�� = To_Date(To_Char(��������_In, 'yyyy-mm-dd') || To_Char(����ʱ��, ' hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
        Where ��¼���� = 4 And ��¼״̬ = 0 And NO = c_Fee.No;
        Update ���˹Һż�¼
        Set ����ʱ�� = To_Date(To_Char(��������_In, 'yyyy-mm-dd') || To_Char(����ʱ��, ' hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
        Where ��¼���� = 2 And ��¼״̬ = 1 And NO = c_Fee.No;
        --���˹ҺŻ���
        Begin
          Select ҽ������, ҽ��id Into v_ҽ��, v_ҽ��id From �ٴ������¼ Where ID = ��¼id_In;
        Exception
          When Others Then
            Null;
        End;
        Update �ٴ������¼ Set ��Լ�� = Nvl(��Լ��, 0) - 1 Where ID = ��¼id_In;
        Update �ٴ������¼ Set ��Լ�� = Nvl(��Լ��, 0) + 1 Where ID = n_��¼id;
      
        Update ���˹ҺŻ���
        Set ��Լ�� = Nvl(��Լ��, 0) - 1
        Where ���� = Trunc(ԤԼ����_In) And Nvl(����id, 0) = c_Fee.ִ�в���id And Nvl(��Ŀid, 0) = c_Fee.�շ�ϸĿid And
              Nvl(ҽ������, 'ҽ��') = Nvl(v_ҽ��, 'ҽ��') And Nvl(ҽ��id, 0) = Nvl(v_ҽ��id, 0) And (���� = �ű�_In Or ���� Is Null);
      
        Update ���˹ҺŻ���
        Set ��Լ�� = Nvl(��Լ��, 0) + 1
        Where ���� = Trunc(��������_In) And Nvl(����id, 0) = c_Fee.ִ�в���id And Nvl(��Ŀid, 0) = c_Fee.�շ�ϸĿid And
              Nvl(ҽ������, 'ҽ��') = Nvl(v_ҽ��, 'ҽ��') And Nvl(ҽ��id, 0) = Nvl(v_ҽ��id, 0) And (���� = �ű�_In Or ���� Is Null);
        If Sql%RowCount = 0 Then
          Insert Into ���˹ҺŻ���
            (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, ��Լ��)
          Values
            (Trunc(��������_In), c_Fee.ִ�в���id, c_Fee.�շ�ϸĿid, v_ҽ��, Decode(v_ҽ��id, 0, Null, v_ҽ��id), �ű�_In, 1);
        End If;
      End If;
    End Loop;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ԤԼ�Һ�_Defer;
/



Create Or Replace Procedure Zl_�Һ����״̬_Update
(
  ����_In       �Һ����״̬.����%Type,
  ����_In       �Һ����״̬.����%Type,
  ���_In       �Һ����״̬.���%Type,
  ״̬_In       �Һ����״̬.״̬%Type,
  ����Ա����_In �Һ����״̬.����Ա����%Type,
  ����_In       Number, --1-����,0-ɾ��
  ��ע_In       �Һ����״̬.��ע%Type := Null,
  ����id_In     �ٴ������¼.Id%Type := Null,
  ԤԼ˳���_In �ٴ�������ſ���.ԤԼ˳���%Type := Null
) As

  v_����         �Һ����״̬.����Ա����%Type;
  v_״̬         �Һ����״̬.״̬%Type;
  n_�Һ��Ű�ģʽ Number(3);
  v_Error        Varchar2(255);
  Err_Custom Exception;
Begin
  n_�Һ��Ű�ģʽ := Zl_To_Number(Nvl(zl_GetSysParameter('�Һ��Ű�ģʽ'), 0));
  If n_�Һ��Ű�ģʽ = 0 Then
    If ����_In = 1 Then
      --�����Һ����״̬
      Begin
        Select ����Ա����, ״̬
        Into v_����, v_״̬
        From �Һ����״̬
        Where ���� = ����_In And ���� = ����_In And ��� = ���_In;
      Exception
        When Others Then
          Null;
      End;
    
      If v_���� Is Null Then
        Insert Into �Һ����״̬
          (����, ����, ���, ״̬, ����Ա����, ��ע, �Ǽ�ʱ��)
        Values
          (����_In, ����_In, ���_In, ״̬_In, ����Ա����_In, ��ע_In, Sysdate);
      Else
        v_Error := '���' || ���_In || '�ѱ�����Ա' || v_����;
        If v_״̬ = 1 Then
          v_Error := v_Error || 'ʹ��';
        Elsif v_״̬ = 2 Then
          v_Error := v_Error || 'ԤԼ';
        Elsif v_״̬ = 3 Then
          v_Error := v_Error || 'Ԥ��';
        End If;
        Raise Err_Custom;
      End If;
    Else
      Begin
        Select ����Ա����, ״̬
        Into v_����, v_״̬
        From �Һ����״̬
        Where ���� = ����_In And ���� = ����_In And ��� = ���_In;
      Exception
        When Others Then
          Null;
      End;
    
      If v_���� <> ����Ա����_In And v_״̬ = 3 Then
        --ȡ��Ԥ�����
        v_Error := '���' || ���_In || '���ɲ���Ա' || v_���� || 'Ԥ����,������ȡ��!';
        Raise Err_Custom;
      Else
        Delete �Һ����״̬ Where ���� = ����_In And ���� = ����_In And ��� = ���_In;
      End If;
    End If;
  Else
    --������Ű�ģʽ
    If ����_In = 1 Then
      --�����Һ����״̬
      Begin
        If ԤԼ˳���_In Is Null Then
          Select ����Ա����, �Һ�״̬
          Into v_����, v_״̬
          From �ٴ�������ſ���
          Where ��¼id = ����id_In And ��� = ���_In;
        Else
          Select ����Ա����, �Һ�״̬
          Into v_����, v_״̬
          From �ٴ�������ſ���
          Where ��¼id = ����id_In And ��� = ���_In And ԤԼ˳��� = ԤԼ˳���_In;
        End If;
      Exception
        When Others Then
          Null;
      End;
    
      If v_���� Is Null Then
        If ԤԼ˳���_In Is Null Then
          Update �ٴ�������ſ���
          Set �Һ�״̬ = ״̬_In, ����Ա���� = ����Ա����_In
          Where ��¼id = ����id_In And ��� = ���_In;
        Else
          Update �ٴ�������ſ���
          Set �Һ�״̬ = ״̬_In, ����Ա���� = ����Ա����_In
          Where ��¼id = ����id_In And ��� = ���_In And ԤԼ˳��� = ԤԼ˳���_In;
        End If;
      Else
        v_Error := '���' || ���_In || '�ѱ�����Ա' || v_����;
        If v_״̬ = 1 Then
          v_Error := v_Error || 'ʹ��';
        Elsif v_״̬ = 2 Then
          v_Error := v_Error || 'ԤԼ';
        Elsif v_״̬ = 3 Then
          v_Error := v_Error || 'Ԥ��';
        End If;
        Raise Err_Custom;
      End If;
    Else
      Begin
        If ԤԼ˳���_In Is Null Then
          Select ����Ա����, �Һ�״̬
          Into v_����, v_״̬
          From �ٴ�������ſ���
          Where ��¼id = ����id_In And ��� = ���_In;
        Else
          Select ����Ա����, �Һ�״̬
          Into v_����, v_״̬
          From �ٴ�������ſ���
          Where ��¼id = ����id_In And ��� = ���_In And ԤԼ˳��� = ԤԼ˳���_In;
        End If;
      Exception
        When Others Then
          Null;
      End;
    
      If v_���� <> ����Ա����_In And v_״̬ = 3 Then
        --ȡ��Ԥ�����
        v_Error := '���' || ���_In || '���ɲ���Ա' || v_���� || 'Ԥ����,������ȡ��!';
        Raise Err_Custom;
      Else
        If ԤԼ˳���_In Is Null Then
          Update �ٴ�������ſ���
          Set �Һ�״̬ = 0, ����Ա���� = Null, ����վip = Null, ����վ���� = Null
          Where ��¼id = ����id_In And ��� = ���_In;
        Else
          Update �ٴ�������ſ���
          Set �Һ�״̬ = 0, ����Ա���� = Null, ����վip = Null, ����վ���� = Null
          Where ��¼id = ����id_In And ��� = ���_In And ԤԼ˳��� = ԤԼ˳���_In;
        End If;
      End If;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�Һ����״̬_Update;
/



Create Or Replace Procedure Zl_ԤԼ�ҺŽ���_����_Insert
(
  No_In            ������ü�¼.No%Type,
  Ʊ�ݺ�_In        ������ü�¼.ʵ��Ʊ��%Type,
  ����id_In        Ʊ��ʹ����ϸ.����id%Type,
  ����id_In        ������ü�¼.����id%Type,
  ����_In          ������ü�¼.��ҩ����%Type,
  ����id_In        ������ü�¼.����id%Type,
  �����_In        ������ü�¼.��ʶ��%Type,
  ����_In          ������ü�¼.����%Type,
  �Ա�_In          ������ü�¼.�Ա�%Type,
  ����_In          ������ü�¼.����%Type,
  ���ʽ_In      ������ü�¼.���ʽ%Type, --���ڴ�Ų��˵�ҽ�Ƹ��ʽ���
  �ѱ�_In          ������ü�¼.�ѱ�%Type,
  ���㷽ʽ_In      Varchar2, --�ֽ�Ľ�������
  �ֽ�֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�ֽ�֧�����ݽ��
  Ԥ��֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱʹ�õ�Ԥ�����
  ����֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�����ʻ�֧�����
  ����ʱ��_In      ������ü�¼.����ʱ��%Type,
  ����_In          �Һ����״̬.���%Type,
  ����Ա���_In    ������ü�¼.����Ա���%Type,
  ����Ա����_In    ������ü�¼.����Ա����%Type,
  ���ɶ���_In      Number := 0,
  �Ǽ�ʱ��_In      ������ü�¼.�Ǽ�ʱ��%Type := Null,
  �����id_In      ����Ԥ����¼.�����id%Type := Null,
  ���㿨���_In    ����Ԥ����¼.���㿨���%Type := Null,
  ����_In          ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In      ����Ԥ����¼.����˵��%Type := Null,
  ����_In          ���˹Һż�¼.����%Type := Null,
  ����ģʽ_In      Number := 0,
  ���ʷ���_In      Number := 0,
  ��Ԥ������ids_In Varchar2 := Null,
  ��������_In      Number := 0
) As
  --���α������շѳ�Ԥ���Ŀ���Ԥ���б�
  --��ID�������ȳ��ϴ�δ����ġ�
  Cursor c_Deposit
  (
    v_����id        ������Ϣ.����id%Type,
    v_��Ԥ������ids Varchar2
  ) Is
    Select *
    From (Select a.Id, a.����id, a.��¼״̬, a.Ԥ�����, a.No, Nvl(a.���, 0) As ���
           From ����Ԥ����¼ A,
                (Select NO, Sum(Nvl(a.���, 0)) As ���
                  From ����Ԥ����¼ A
                  Where a.����id Is Null And Nvl(a.���, 0) <> 0 And
                        a.����id In (Select Column_Value From Table(f_Num2list(v_��Ԥ������ids))) And Nvl(a.Ԥ�����, 2) = 1
                  Group By NO
                  Having Sum(Nvl(a.���, 0)) <> 0) B
           Where a.����id Is Null And Nvl(a.���, 0) <> 0 And a.���㷽ʽ Not In (Select ���� From ���㷽ʽ Where ���� = 5) And
                 a.No = b.No And a.����id In (Select Column_Value From Table(f_Num2list(v_��Ԥ������ids))) And
                 Nvl(a.Ԥ�����, 2) = 1
           Union All
           Select 0 As ID, Max(����id) As ����id, ��¼״̬, Ԥ�����, NO, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���
           From ����Ԥ����¼
           Where ��¼���� In (1, 11) And ����id Is Not Null And Nvl(���, 0) <> Nvl(��Ԥ��, 0) And
                 ����id In (Select Column_Value From Table(f_Num2list(v_��Ԥ������ids))) And Nvl(Ԥ�����, 2) = 1 Having
            Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
           Group By ��¼״̬, Ԥ�����, NO)
    Order By Decode(����id, Nvl(v_����id, 0), 0, 1), ID, NO, Ԥ�����;

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_�ֽ�     ���㷽ʽ.����%Type;
  v_�����ʻ� ���㷽ʽ.����%Type;
  v_�������� �ŶӽкŶ���.��������%Type;
  v_�ű�     ������ü�¼.���㵥λ%Type;
  v_����     ������ü�¼.��ҩ����%Type;
  v_�ŶӺ��� �ŶӽкŶ���.�ŶӺ��� %Type;
  v_ԤԼ��ʽ ���˹Һż�¼.ԤԼ��ʽ %Type;

  n_��ӡid        Ʊ�ݴ�ӡ����.Id%Type;
  n_Ԥ�����      ����Ԥ����¼.���%Type;
  n_����ֵ        ����Ԥ����¼.���%Type;
  v_��Ԥ������ids Varchar2(4000);

  n_�Һ�id         ���˹Һż�¼.Id%Type;
  n_����̨ǩ���Ŷ� Number;
  n_��id           ����ɿ����.Id%Type;
  n_Count          Number(18);
  n_�Ŷ�           Number;
  n_�����Ŷ�       Number;
  n_��ǰ���       ����Ԥ����¼.���%Type;
  n_Ԥ��id         ����Ԥ����¼.Id%Type;
  n_���ѿ�id       ���ѿ�Ŀ¼.Id%Type;
  n_���ƿ�         Number;

  d_Date         Date;
  d_ԤԼʱ��     ������ü�¼.����ʱ��%Type;
  d_����ʱ��     Date;
  d_�Ŷ�ʱ��     Date;
  n_ʱ��         Number := 0;
  n_����         Number := 0;
  v_��������     Varchar2(2000);
  v_��ǰ����     Varchar2(500);
  n_������     ����Ԥ����¼.��Ԥ��%Type;
  v_�������     ����Ԥ����¼.�������%Type;
  v_���㷽ʽ     ����Ԥ����¼.���㷽ʽ%Type;
  n_��������־   Number(3);
  v_�Ŷ����     �ŶӽкŶ���.�Ŷ����%Type;
  n_����ģʽ     ������Ϣ.����ģʽ%Type;
  n_Ʊ��         Ʊ��ʹ����ϸ.Ʊ��%Type;
  v_���ʽ     ���˹Һż�¼.ҽ�Ƹ��ʽ%Type;
  n_����ģʽ     Number := 0;
  n_�����¼id   ���˹Һż�¼.�����¼id%Type;
  n_�³����¼id ���˹Һż�¼.�����¼id%Type;
  n_��Դid       �ٴ������¼.��Դid%Type;
  n_ԤԼ˳���   �ٴ�������ſ���.ԤԼ˳���%Type;
  v_Registtemp   Varchar2(500);
  n_���         Number(3);
Begin
  n_��id          := Zl_Get��id(����Ա����_In);
  v_��Ԥ������ids := Nvl(��Ԥ������ids_In, ����id_In);
  n_����ģʽ      := Nvl(zl_GetSysParameter('ԤԼ����ģʽ', 1111), 0);

  --��ȡ���㷽ʽ����
  Begin
    Select ���� Into v_�ֽ� From ���㷽ʽ Where ���� = 1;
  Exception
    When Others Then
      v_�ֽ� := '�ֽ�';
  End;
  Begin
    Select ���� Into v_�����ʻ� From ���㷽ʽ Where ���� = 3;
  Exception
    When Others Then
      v_�����ʻ� := '�����ʻ�';
  End;
  If �Ǽ�ʱ��_In Is Null Then
    Select Sysdate Into d_Date From Dual;
  Else
    d_Date := �Ǽ�ʱ��_In;
  End If;

  --���¹Һ����״̬
  Begin
    Select �ű�, ����, Trunc(����ʱ��), ����ʱ��, ԤԼ��ʽ, �����¼id
    Into v_�ű�, v_����, d_ԤԼʱ��, d_����ʱ��, v_ԤԼ��ʽ, n_�����¼id
    From ���˹Һż�¼
    Where ��¼���� = 2 And ��¼״̬ = 1 And Rownum = 1 And NO = No_In;
  Exception
    When Others Then
      v_Err_Msg := '��ǰԤԼ�Һŵ��ѱ������˽���';
      Raise Err_Item;
  End;

  --�ж��Ƿ��ʱ��
  Select Nvl(�Ƿ��ʱ��, 0), ��Դid Into n_ʱ��, n_��Դid From �ٴ������¼ Where ID = n_�����¼id;

  If n_ʱ�� = 0 And ��������_In = 0 Then
    If n_����ģʽ = 0 Then
      If Trunc(����ʱ��_In) = Trunc(Sysdate) Then
        d_����ʱ�� := ����ʱ��_In;
      Else
        d_����ʱ�� := Sysdate;
      End If;
    Else
      d_����ʱ�� := ����ʱ��_In;
    End If;
  Else
    If Not ����ʱ��_In Is Null Then
      d_����ʱ�� := ����ʱ��_In;
    End If;
  End If;
  If Not v_���� Is Null Then
    If ����_In Is Null Then
      Update �ٴ�������ſ��� Set �Һ�״̬ = 0 Where (��� = v_���� Or ��ע = v_����) And ��¼id = n_�����¼id;
    Else
      If Trunc(d_ԤԼʱ��) <> Trunc(Sysdate) And n_����ģʽ = 0 Then
        If n_ʱ�� = 0 And ��������_In = 0 Then
          --��ǰ���ջ��ӳٽ���
          Update �ٴ�������ſ��� Set �Һ�״̬ = 0 Where ��� = v_���� And ��¼id = n_�����¼id;
        
          Begin
            Select ID
            Into n_�³����¼id
            From �ٴ������¼
            Where ��Դid = n_��Դid And �������� = Trunc(Sysdate) And Rownum < 2;
          Exception
            When Others Then
              v_Err_Msg := '���յ���û�г��ﰲ��,�޷�����!';
              Raise Err_Item;
          End;
        
          Begin
            Select 1
            Into n_����
            From �ٴ�������ſ���
            Where ��¼id = n_�³����¼id And ��� = v_���� And Nvl(�Һ�״̬, 0) = 0;
          Exception
            When Others Then
              n_���� := 0;
          End;
        
          If n_���� = 0 Then
            Update �ٴ�������ſ���
            Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In
            Where ��¼id = n_�³����¼id And ��� = v_���� And Nvl(�Һ�״̬, 0) = 0;
          Else
            --�����ѱ�ʹ�õ����
            Select Min(���) Into v_���� From �ٴ�������ſ��� Where ��¼id = n_�³����¼id And Nvl(�Һ�״̬, 0) = 0;
            If v_���� Is Null Then
              v_Err_Msg := '���յ���û�п������,�޷�����!';
              Raise Err_Item;
            End If;
            Update �ٴ�������ſ���
            Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In
            Where ��¼id = n_�³����¼id And ��� = v_���� And Nvl(�Һ�״̬, 0) = 0;
          End If;
        Else
          Begin
            Select ID
            Into n_�³����¼id
            From �ٴ������¼
            Where ��Դid = n_��Դid And �������� = Trunc(Sysdate) And Rownum < 2;
          Exception
            When Others Then
              v_Err_Msg := '���յ���û�г��ﰲ��,�޷�����!';
              Raise Err_Item;
          End;
          Update �ٴ�������ſ���
          Set �Һ�״̬ = 0, ����Ա���� = ����Ա����_In
          Where (��� = v_���� Or ��ע = v_����) And ��¼id = n_�����¼id And Nvl(�Һ�״̬, 0) = 2
          Returning ԤԼ˳��� Into n_ԤԼ˳���;
        
          Update �ٴ�������ſ���
          Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In, ԤԼ˳��� = n_ԤԼ˳���
          Where ��� = v_���� And ��¼id = n_�³����¼id And Nvl(�Һ�״̬, 0) = 0;
          If Sql% RowCount = 0 Then
            v_Err_Msg := '���' || v_���� || '�ѱ�������ʹ��,������ѡ��һ�����.';
            Raise Err_Item;
          End If;
        End If;
      Else
        Update �ٴ�������ſ���
        Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In
        Where (��� = v_���� Or ��ע = v_����) And ��¼id = n_�����¼id;
        If Sql%RowCount = 0 Then
          v_Err_Msg := '���' || v_���� || '�ѱ�������ʹ��,������ѡ��һ�����.';
          Raise Err_Item;
        End If;
      End If;
    End If;
  Else
    If Not ����_In Is Null Then
      If Trunc(d_ԤԼʱ��) <> Trunc(Sysdate) And n_����ģʽ = 0 Then
        Begin
          Select ID
          Into n_�³����¼id
          From �ٴ������¼
          Where ��Դid = n_��Դid And �������� = Trunc(Sysdate) And Rownum < 2;
        Exception
          When Others Then
            v_Err_Msg := '���յ���û�г��ﰲ��,�޷�����!';
            Raise Err_Item;
        End;
        Update �ٴ�������ſ���
        Set �Һ�״̬ = 0, ����Ա���� = ����Ա����_In
        Where (��� = ����_In Or ��ע = ����_In) And ��¼id = n_�����¼id And Nvl(�Һ�״̬, 0) = 2
        Returning ԤԼ˳��� Into n_ԤԼ˳���;
        Update �ٴ�������ſ���
        Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In, ԤԼ˳��� = n_ԤԼ˳���
        Where ��� = ����_In And ��¼id = n_�³����¼id And Nvl(�Һ�״̬, 0) = 0;
        If Sql%RowCount = 0 Then
          v_Err_Msg := '���' || ����_In || '�ѱ�������ʹ��,������ѡ��һ�����.';
          Raise Err_Item;
        End If;
      Else
        Update �ٴ�������ſ���
        Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In
        Where (��� = ����_In Or ��ע = ����_In) And ��¼id = n_�����¼id;
      
      End If;
      v_���� := ����_In;
    Else
      v_���� := Null;
    End If;
  End If;

  --����������ü�¼
  Update ������ü�¼
  Set ��¼״̬ = 1, ʵ��Ʊ�� = Decode(Nvl(���ʷ���_In, 0), 1, Null, Ʊ�ݺ�_In), ����id = Decode(Nvl(���ʷ���_In, 0), 1, Null, ����id_In),
      ���ʽ�� = Decode(Nvl(���ʷ���_In, 0), 1, Null, ʵ�ս��), ��ҩ���� = ����_In, ����id = ����id_In, ��ʶ�� = �����_In, ���� = ����_In, ���� = ����_In,
      �Ա� = �Ա�_In, ���ʽ = ���ʽ_In, �ѱ� = �ѱ�_In, ����ʱ�� = d_����ʱ��, �Ǽ�ʱ�� = d_Date, ����Ա��� = ����Ա���_In, ����Ա���� = ����Ա����_In,
      �ɿ���id = n_��id, ���ʷ��� = Decode(Nvl(���ʷ���_In, 0), 1, 1, 0)
  Where ��¼���� = 4 And ��¼״̬ = 0 And NO = No_In;

  v_Registtemp := zl_GetSysParameter('�Һ��Ű�ģʽ');
  If Substr(v_Registtemp, 1, 1) = 1 Then
    Begin
      If To_Date(Substr(v_Registtemp, 3), 'yyyy-mm-dd hh24:mi:ss') > d_����ʱ�� Then
        v_Err_Msg := '����ʱ��' || To_Char(d_����ʱ��, 'yyyy-mm-dd hh24:mi:ss') || 'δ���ó�����Ű�ģʽ,Ŀǰ�޷�����!';
        Raise Err_Item;
      End If;
    Exception
      When Others Then
        Null;
    End;
    Begin
      Select 1
      Into n_���
      From �ٴ������¼
      Where ID = Nvl(n_�³����¼id, n_�����¼id) And d_����ʱ�� Between ͣ�￪ʼʱ�� And ͣ����ֹʱ��;
    Exception
      When Others Then
        n_��� := 0;
    End;
    If n_��� = 1 Then
      v_Err_Msg := '����ʱ��' || To_Char(d_����ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '�İ����Ѿ���ͣ��,�޷�����!';
      Raise Err_Item;
    End If;
  End If;

  --���˹Һż�¼
  Update ���˹Һż�¼
  Set ������ = ����Ա����_In, ����ʱ�� = d_Date, ��¼���� = 1, ����id = ����id_In, ����� = �����_In, ����ʱ�� = d_����ʱ��, ���� = ����_In, �Ա� = �Ա�_In,
      ���� = ����_In, ����Ա��� = ����Ա���_In, ����Ա���� = ����Ա����_In, ���� = Decode(Nvl(����_In, 0), 0, Null, ����_In), ���� = v_����, ���� = ����_In,
      �����¼id = Nvl(n_�³����¼id, n_�����¼id)
  Where ��¼״̬ = 1 And NO = No_In And ��¼���� = 2
  Returning ID Into n_�Һ�id;
  If Sql%NotFound Then
    Begin
      Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
      Begin
        Select ���� Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = ���ʽ_In And Rownum < 2;
      Exception
        When Others Then
          v_���ʽ := Null;
      End;
      Insert Into ���˹Һż�¼
        (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
         ժҪ, ����, ԤԼ, ԤԼ��ʽ, ������, ����ʱ��, ԤԼʱ��, ����, ҽ�Ƹ��ʽ, �����¼id)
        Select n_�Һ�id, No_In, 1, 1, ����id_In, �����_In, ����_In, �Ա�_In, ����_In, ���㵥λ, �Ӱ��־, ����_In, Null, ִ�в���id, ִ����, 0, Null,
               �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����, ժҪ, v_����, 1, Substr(����, 1, 10) As ԤԼ��ʽ, ����Ա����_In, Nvl(�Ǽ�ʱ��_In, Sysdate), ����ʱ��,
               Decode(Nvl(����_In, 0), 0, Null, ����_In), v_���ʽ, Nvl(n_�³����¼id, n_�����¼id)
        From ������ü�¼
        Where ��¼���� = 4 And ��¼״̬ = 1 And Rownum = 1 And NO = No_In;
    Exception
      When Others Then
        v_Err_Msg := '���ڲ���ԭ��,���ݺ�Ϊ��' || No_In || '���Ĳ���' || ����_In || '�Ѿ�������';
        Raise Err_Item;
    End;
  End If;

  --0-����������;1-��ҽ�������̨�Ŷ�;2-�ȷ���,��ҽ��վ
  If Nvl(���ɶ���_In, 0) <> 0 Then
    n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113));
    If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Then
      For v_�Һ� In (Select ID, ����, ����, ִ����, ִ�в���id, ����ʱ��, �ű�, ���� From ���˹Һż�¼ Where NO = No_In) Loop
      
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
          --��������
          --����ִ�в��š���������
          n_�Һ�id   := v_�Һ�.Id;
          v_�������� := v_�Һ�.ִ�в���id;
          v_�ŶӺ��� := Zlgetnextqueue(v_�Һ�.ִ�в���id, n_�Һ�id, v_�Һ�.�ű� || '|' || v_�Һ�.����);
          v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 0);
        
          --�Һ�id_In,����_In,����_In,ȱʡ����_In,��չ_In(������)
          d_�Ŷ�ʱ�� := Zl_Get_Queuedate(n_�Һ�id, v_�Һ�.�ű�, v_�Һ�.����, v_�Һ�.����ʱ��);
          --   ��������_In , ҵ������_In, ҵ��id_In,����id_In,�ŶӺ���_In,�Ŷӱ��_In,��������_In,����ID_IN, ����_In, ҽ������_In,
          Zl_�ŶӽкŶ���_Insert(v_��������, 0, n_�Һ�id, v_�Һ�.ִ�в���id, v_�ŶӺ���, Null, ����_In, ����id_In, v_�Һ�.����, v_�Һ�.ִ����, d_�Ŷ�ʱ��,
                           v_ԤԼ��ʽ, Null, v_�Ŷ����);
        Elsif Nvl(n_�����Ŷ�, 0) = 1 Then
          --���¶��к�
          v_�ŶӺ��� := Zlgetnextqueue(v_�Һ�.ִ�в���id, v_�Һ�.Id, v_�Һ�.�ű� || '|' || Nvl(v_�Һ�.����, 0));
          v_�Ŷ���� := Zlgetsequencenum(0, v_�Һ�.Id, 1);
          --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In, ҽ������_In ,�ŶӺ���_In
          Zl_�ŶӽкŶ���_Update(v_�Һ�.ִ�в���id, 0, v_�Һ�.Id, v_�Һ�.ִ�в���id, v_�Һ�.����, v_�Һ�.����, v_�Һ�.ִ����, v_�ŶӺ���, v_�Ŷ����);
        
        Else
          --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In, ҽ������_In ,�ŶӺ���_In
          Zl_�ŶӽкŶ���_Update(v_�Һ�.ִ�в���id, 0, v_�Һ�.Id, v_�Һ�.ִ�в���id, v_�Һ�.����, v_�Һ�.����, v_�Һ�.ִ����);
        End If;
      End Loop;
    End If;
  End If;

  --���ܽ��㵽����Ԥ����¼
  If Nvl(���ʷ���_In, 0) = 0 Then
    If Nvl(�ֽ�֧��_In, 0) = 0 And Nvl(����֧��_In, 0) = 0 And Nvl(Ԥ��֧��_In, 0) = 0 Then
      Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      Insert Into ����Ԥ����¼
        (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
         ��������)
      Values
        (n_Ԥ��id, 4, 1, No_In, Decode(����id_In, 0, Null, ����id_In), v_�ֽ�, 0, �Ǽ�ʱ��_In, ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�',
         n_��id, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, Null, 4);
    End If;
    If Nvl(�ֽ�֧��_In, 0) <> 0 Then
      v_�������� := ���㷽ʽ_In || '|'; --�Կո�ֿ���|��β,û�н�������
      While v_�������� Is Not Null Loop
        v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '|') - 1);
        v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
      
        v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
        n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
      
        v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
        v_������� := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
      
        v_��ǰ����   := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
        n_��������־ := To_Number(v_��ǰ����);
      
        If n_��������־ = 0 Then
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��,
             ������λ, ��������, �������)
          Values
            (n_Ԥ��id, 4, 1, No_In, Decode(����id_In, 0, Null, ����id_In), Nvl(v_���㷽ʽ, v_�ֽ�), Nvl(n_������, 0), �Ǽ�ʱ��_In,
             ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�', n_��id, Null, Null, Null, Null, Null, Null, 4, v_�������);
        Else
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��,
             ������λ, ��������, �������)
          Values
            (n_Ԥ��id, 4, 1, No_In, Decode(����id_In, 0, Null, ����id_In), Nvl(v_���㷽ʽ, v_�ֽ�), Nvl(n_������, 0), �Ǽ�ʱ��_In,
             ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�', n_��id, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, Null, 4, v_�������);
          If Nvl(���㿨���_In, 0) <> 0 Then
            n_���ѿ�id := Null;
            Begin
              Select Nvl(���ƿ�, 0), 1 Into n_���ƿ�, n_Count From �����ѽӿ�Ŀ¼ Where ��� = ���㿨���_In;
            Exception
              When Others Then
                n_Count := 0;
            End;
            If n_Count = 0 Then
              v_Err_Msg := 'û�з���ԭ���㿨����Ӧ���,���ܼ���������';
              Raise Err_Item;
            End If;
            If n_���ƿ� = 1 Then
              Select ID
              Into n_���ѿ�id
              From ���ѿ�Ŀ¼
              Where �ӿڱ�� = ���㿨���_In And ���� = ����_In And
                    ��� = (Select Max(���) From ���ѿ�Ŀ¼ Where �ӿڱ�� = ���㿨���_In And ���� = ����_In);
            End If;
            Zl_���˿������¼_Insert(���㿨���_In, n_���ѿ�id, ���㷽ʽ_In, �ֽ�֧��_In, ����_In, Null, �Ǽ�ʱ��_In, Null, ����id_In, n_Ԥ��id);
          End If;
        End If;
      
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + n_������
        Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = Nvl(v_���㷽ʽ, v_�ֽ�)
        Returning ��� Into n_����ֵ;
      
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ����
            (�տ�Ա, ���㷽ʽ, ����, ���)
          Values
            (����Ա����_In, Nvl(v_���㷽ʽ, v_�ֽ�), 1, n_������);
          n_����ֵ := n_������;
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From ��Ա�ɿ����
          Where �տ�Ա = ����Ա����_In And ���㷽ʽ = Nvl(v_���㷽ʽ, v_�ֽ�) And ���� = 1 And Nvl(���, 0) = 0;
        End If;
      
        v_�������� := Substr(v_��������, Instr(v_��������, '|') + 1);
      End Loop;
    End If;
  End If;

  --���ھ��￨ͨ��Ԥ����Һ�
  If Nvl(Ԥ��֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 Then
    n_Ԥ����� := Ԥ��֧��_In;
    For r_Deposit In c_Deposit(����id_In, v_��Ԥ������ids) Loop
      n_��ǰ��� := Case
                  When r_Deposit.��� - n_Ԥ����� < 0 Then
                   r_Deposit.���
                  Else
                   n_Ԥ�����
                End;
      If r_Deposit.Id <> 0 Then
        --��һ�γ�Ԥ��(���Ͻ���ID,���Ϊ0)
        Update ����Ԥ����¼ Set ��Ԥ�� = 0, ����id = ����id_In, �������� = 4 Where ID = r_Deposit.Id;
      End If;
      --���ϴ�ʣ���
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, Ԥ�����, �������, ��������)
        Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �Ǽ�ʱ��_In,
               ����Ա����_In, ����Ա���_In, n_��ǰ���, ����id_In, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, Ԥ�����, ����id_In, 4
        From ����Ԥ����¼
        Where NO = r_Deposit.No And ��¼״̬ = r_Deposit.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
    
      --���²���Ԥ�����
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) - n_��ǰ���
      Where ����id = r_Deposit.����id And ���� = 1 And ���� = Nvl(r_Deposit.Ԥ�����, 2)
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, Ԥ�����, ����)
        Values
          (r_Deposit.����id, Nvl(r_Deposit.Ԥ�����, 2), -1 * n_��ǰ���, 1);
        n_����ֵ := -1 * n_��ǰ���;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From �������
        Where ����id = r_Deposit.����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
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

  --����ҽ���Һ�
  If Nvl(����֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 Then
    Insert Into ����Ԥ����¼
      (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
       Ԥ�����, �������, ��������)
    Values
      (����Ԥ����¼_Id.Nextval, 4, 1, No_In, ����id_In, v_�����ʻ�, ����֧��_In, d_Date, ����Ա���_In, ����Ա����_In, ����id_In, 'ҽ���Һ�', n_��id,
       �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, Null, Null, ����id_In, 4);
  End If;

  --��ػ��ܱ�Ĵ���
  --��Ա�ɿ����
  If Nvl(����֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 Then
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + ����֧��_In
    Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = v_�����ʻ�
    Returning ��� Into n_����ֵ;
  
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (����Ա����_In, v_�����ʻ�, 1, ����֧��_In);
      n_����ֵ := ����֧��_In;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ��Ա�ɿ���� Where �տ�Ա = ����Ա����_In And ���� = 1 And Nvl(���, 0) = 0;
    End If;
  End If;

  --����Ʊ��ʹ�����
  If Ʊ�ݺ�_In Is Not Null And Nvl(���ʷ���_In, 0) = 0 Then
    Select Ʊ�ݴ�ӡ����_Id.Nextval Into n_��ӡid From Dual;
  
    --��ǰƱ�ݵ�Ʊ��
    Select Ʊ�� Into n_Ʊ�� From Ʊ�����ü�¼ Where ID = Nvl(����id_In, 0);
    --����Ʊ��
    Insert Into Ʊ�ݴ�ӡ���� (ID, ��������, NO) Values (n_��ӡid, 4, No_In);
  
    Insert Into Ʊ��ʹ����ϸ
      (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
    Values
      (Ʊ��ʹ����ϸ_Id.Nextval, n_Ʊ��, Ʊ�ݺ�_In, 1, 1, ����id_In, n_��ӡid, d_Date, ����Ա����_In);
  
    --״̬�Ķ�
    Update Ʊ�����ü�¼
    Set ��ǰ���� = Ʊ�ݺ�_In, ʣ������ = Decode(Sign(ʣ������ - 1), -1, 0, ʣ������ - 1), ʹ��ʱ�� = d_Date
    Where ID = Nvl(����id_In, 0);
  End If;

  If Nvl(���ʷ���_In, 0) = 1 Then
    --����
    If Nvl(����id_In, 0) = 0 Then
      v_Err_Msg := 'Ҫ��Բ��˵ĹҺŷѽ��м��ʣ������ǽ������˲��ܼ��ʹҺš�';
      Raise Err_Item;
    End If;
    For c_���� In (Select ʵ�ս��, ���˿���id, ��������id, ִ�в���id, ������Ŀid
                 From ������ü�¼
                 Where ��¼���� = 4 And ��¼״̬ = 1 And NO = No_In And Nvl(���ʷ���, 0) = 1) Loop
      --�������
      Update �������
      Set ������� = Nvl(�������, 0) + Nvl(c_����.ʵ�ս��, 0)
      Where ����id = Nvl(����id_In, 0) And ���� = 1 And ���� = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, ����, �������, Ԥ�����)
        Values
          (����id_In, 1, 1, Nvl(c_����.ʵ�ս��, 0), 0);
      End If;
    
      --����δ�����
      Update ����δ�����
      Set ��� = Nvl(���, 0) + Nvl(c_����.ʵ�ս��, 0)
      Where ����id = ����id_In And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(c_����.���˿���id, 0) And
            Nvl(��������id, 0) = Nvl(c_����.��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(c_����.ִ�в���id, 0) And ������Ŀid + 0 = c_����.������Ŀid And
            ��Դ;�� + 0 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into ����δ�����
          (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
        Values
          (����id_In, Null, Null, c_����.���˿���id, c_����.��������id, c_����.ִ�в���id, c_����.������Ŀid, 1, Nvl(c_����.ʵ�ս��, 0));
      End If;
    End Loop;
  End If;
  If Nvl(����id_In, 0) <> 0 Then
    n_����ģʽ := 0;
    Update ������Ϣ
    Set ����ʱ�� = d_����ʱ��, ����״̬ = 1, �������� = ����_In
    Where ����id = ����id_In
    Returning Nvl(����ģʽ, 0) Into n_����ģʽ;
    --ȡ����:
    If Nvl(n_����ģʽ, 0) <> Nvl(����ģʽ_In, 0) Then
      --����ģʽ��ȷ��
      If n_����ģʽ = 1 And Nvl(����ģʽ_In, 0) = 0 Then
        --�����Ѿ���"�����ƺ�����",������"�Ƚ�������Ƶ�",�����Ƿ����δ������
        Select Count(1)
        Into n_Count
        From ����δ�����
        Where ����id = ����id_In And (��Դ;�� In (1, 4) Or ��Դ;�� = 3 And Nvl(��ҳid, 0) = 0) And Nvl(���, 0) <> 0 And Rownum < 2;
        If Nvl(n_Count, 0) <> 0 Then
          --����δ�������ݣ������Ƚ���������ִ��
          v_Err_Msg := '��ǰ���˵ľ���ģʽΪ�����ƺ�����Ҵ���δ����ã�����������ò��˵ľ���ģʽ,������ȶ�δ����ý��ʺ��ٹҺŻ򲻵������˵ľ���ģʽ!';
          Raise Err_Item;
        End If;
        --���
        --δ����ҽ��ҵ��ģ�����ʱ�͹Һŵ�,��Ҫ��֤ͬһ�εľ���ģʽ��һ����(�����Ѿ���飬�����ٴ���)
      End If;
      Update ������Ϣ Set ����ģʽ = ����ģʽ_In Where ����id = ����id_In;
    End If;
  End If;

  --���˵�����Ϣ
  If ����id_In Is Not Null Then
    Update ������Ϣ
    Set ������ = Null, ������ = Null, �������� = Null
    Where ����id = ����id_In And Exists (Select 1
           From ���˵�����¼
           Where ����id = ����id_In And ��ҳid Is Not Null And
                 �Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ���˵�����¼ Where ����id = ����id_In));
    If Sql%RowCount > 0 Then
      Update ���˵�����¼
      Set ����ʱ�� = d_Date
      Where ����id = ����id_In And ��ҳid Is Not Null And Nvl(����ʱ��, d_Date) > d_Date;
    End If;
  End If;
  --��Ϣ����
  Begin
    Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
      Using 1, n_�Һ�id;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ԤԼ�ҺŽ���_����_Insert;
/




Create Or Replace Procedure Zl_����ԤԼ�Ǽ�_Insert
(
  ����id_In     ������Ϣ.����id%Type,
  ��Դid_In     �ٴ������Դ.Id%Type,
  ���﷽ʽ_In   ���˷�����Ϣ��¼.���﷽ʽ%Type,
  ����_In       ���˷�����Ϣ��¼.����%Type,
  ˵��_In       ���˷�����Ϣ��¼.֪ͨԭ��%Type,
  ����ʱ��_In   ���˷�����Ϣ��¼.��ʼʱ��%Type,
  ��������_In   Number,
  ����Ա����_In ���˹Һż�¼.����Ա����%Type,
  ����Ա���_In ���˹Һż�¼.����Ա���%Type
) As
  d_��ʼʱ�� ���˷�����Ϣ��¼.��ʼʱ��%Type;
  d_����ʱ�� ���˷�����Ϣ��¼.��ֹʱ��%Type;
  v_����     ���˷�����Ϣ��¼.����%Type;
  n_����id   ���˷�����Ϣ��¼.����id%Type;
  n_��Ŀid   ���˷�����Ϣ��¼.��Ŀid%Type;
  n_ҽ��id   ���˷�����Ϣ��¼.ҽ��id%Type;
  v_ҽ������ ���˷�����Ϣ��¼.ҽ������%Type;
  v_Err_Msg  Varchar2(200);
  Err_Item Exception;
Begin
  Begin
    Select ����, ����id, ��Ŀid, ҽ��id, ҽ������
    Into v_����, n_����id, n_��Ŀid, n_ҽ��id, v_ҽ������
    From �ٴ������Դ
    Where ID = ��Դid_In;
  Exception
    When Others Then
      v_Err_Msg := 'û���ҵ���Դ��Ϣ,ԤԼ�Ǽ�ʧ��';
      Raise Err_Item;
  End;
  d_��ʼʱ�� := Trunc(����ʱ��_In);
  d_����ʱ�� := d_��ʼʱ�� + Nvl(��������_In, 1);
  Insert Into ���˷�����Ϣ��¼
    (ID, ֪ͨ����, ��Դid, ����, ����id, ��Ŀid, ҽ��id, ҽ������, ����id, ���﷽ʽ, ����, ��ʼʱ��, ��ֹʱ��, ֪ͨԭ��, �Ǽ���, �Ǽ�ʱ��)
  Values
    (���˷�����Ϣ��¼_Id.Nextval, 3, ��Դid_In, v_����, n_����id, n_��Ŀid, n_ҽ��id, v_ҽ������, ����id_In, ���﷽ʽ_In, ����_In, d_��ʼʱ��, d_����ʱ��,
     ˵��_In, ����Ա����_In, Sysdate);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ԤԼ�Ǽ�_Insert;
/


CREATE OR REPLACE Function Zl_ԤԼ��ʽ_Check
(
  ��¼id_In   �ٴ������¼.Id%Type,
  ���_In     �ٴ�������ſ���.���%Type,
  ԤԼ��ʽ_In ԤԼ��ʽ.����%Type
) Return Number Is
  --����:ԤԼʱ�����Ӧ��ԤԼ��ʽ�Ƿ����
  --����:0-���δͨ��,1-���ͨ��
  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  n_���Ʒ�ʽ �ٴ�����Һſ��Ƽ�¼.���Ʒ�ʽ%Type;
  n_��Լ��   �ٴ������¼.��Լ��%Type;
  n_����     �ٴ�����Һſ��Ƽ�¼.����%Type;
  n_��Լ��   �ٴ������¼.��Լ��%Type;
  n_��ʱ��   �ٴ������¼.�Ƿ��ʱ��%Type;
  n_��ſ��� �ٴ������¼.�Ƿ���ſ���%Type;
Begin
  Begin
    Select ���Ʒ�ʽ
    Into n_���Ʒ�ʽ
    From �ٴ�����Һſ��Ƽ�¼
    Where ���� = 2 And ���� = 1 And ��¼id = ��¼id_In And Rownum < 2;
  Exception
    When Others Then
      Return 1;
  End;
  If n_���Ʒ�ʽ = 0 Then
    Return 0;
  End If;
  If n_���Ʒ�ʽ = 1 Or n_���Ʒ�ʽ = 2 Then
    Select Nvl(��Լ��, �޺���) Into n_��Լ�� From �ٴ������¼ Where ID = ��¼id_In;
    Select ����
    Into n_����
    From �ٴ�����Һſ��Ƽ�¼
    Where ���� = 2 And ���� = 1 And ���� = ԤԼ��ʽ_In And ��¼id = ��¼id_In;
    If n_���Ʒ�ʽ = 1 Then
      n_��Լ�� := Round(n_��Լ�� * n_���� / 100);
    Else
      n_��Լ�� := n_����;
    End If;
    Select Count(1)
    Into n_��Լ��
    From ���˹Һż�¼
    Where �����¼id = ��¼id_In And ��¼״̬ = 1 And ԤԼ��ʽ = ԤԼ��ʽ_In;
    If n_��Լ�� >= n_��Լ�� Then
      Return 0;
    End If;
  End If;
  If n_���Ʒ�ʽ = 3 Then
    Select ����
    Into n_��Լ��
    From �ٴ�����Һſ��Ƽ�¼
    Where ���� = 2 And ���� = 1 And ���� = ԤԼ��ʽ_In And ��¼id = ��¼id_In And ��� = ���_In;
    Select �Ƿ��ʱ��, �Ƿ���ſ��� Into n_��ʱ��, n_��ſ��� From �ٴ������¼ Where ID = ��¼id_In;
    If n_��ſ��� = 1 Then
      Select Nvl(Max(1), 0) Into n_��Լ�� From ���˹Һż�¼ Where �����¼id = ��¼id_In And ���� = ���_In;
    Else
      Select Count(1)
      Into n_��Լ��
      From �ٴ�������ſ��� A, ���˹Һż�¼ B
      Where a.��¼id = ��¼id_In And a.ԤԼ˳��� Is Not Null And Nvl(a.�Һ�״̬, 0) <> 0 And a.��ע = b.���� And b.ԤԼ��ʽ = ԤԼ��ʽ_In And
            b.��¼״̬ = 1;
    End If;
    If n_��Լ�� >= n_��Լ�� Then
      Return 0;
    End If;
  End If;
  If n_���Ʒ�ʽ = 4 Then
    Return 1;
  End If;
  Return 1;
Exception
  When Err_Item Then
    Return 0;
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ԤԼ��ʽ_Check;
/


Create Or Replace Procedure Zl_���������Һ�_Insert
(
  ������ʽ_In      Integer,
  ����id_In        ������ü�¼.����id%Type,
  ����_In          �ҺŰ���.����%Type,
  ����_In          �Һ����״̬.���%Type,
  ���ݺ�_In        ������ü�¼.No%Type,
  Ʊ�ݺ�_In        ������ü�¼.ʵ��Ʊ��%Type,
  ���㷽ʽ_In      Varchar2,
  ժҪ_In          ������ü�¼.ժҪ%Type, --ԤԼ�Һ�ժҪ��Ϣ
  ����ʱ��_In      ������ü�¼.����ʱ��%Type,
  �Ǽ�ʱ��_In      ������ü�¼.�Ǽ�ʱ��%Type,
  ������λ_In      �Һź�����λ.����%Type,
  �ҺŽ��ϼ�_In  ������ü�¼.ʵ�ս��%Type,
  ����id_In        Ʊ��ʹ����ϸ.����id%Type,
  �շ�Ʊ��_In      Number := 0, --�Һ��Ƿ�ʹ���շ�Ʊ��
  ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type,
  ����˵��_In      ����Ԥ����¼.����˵��%Type,
  ԤԼ��ʽ_In      ԤԼ��ʽ.����%Type := Null,
  Ԥ��id_In        ����Ԥ����¼.Id%Type := Null,
  �����id_In      ����Ԥ����¼.�����id%Type := Null,
  �������״̬_In  Number := 0,
  �Ƿ������豸_In  Number := 0,
  ����id_In        ������ü�¼.����id%Type := Null,
  ��������_In      Number := 0,
  ���ս���_In      Varchar2 := Null,
  ��Ԥ��_In        Number := Null,
  ֧������_In      ����Ԥ����¼.����%Type := Null,
  �˺�����_In      Number := 1,
  �ѱ�_In          ������ü�¼.�ѱ�%Type := Null,
  ��Ԥ������ids_In Varchar2 := Null,
  ������_In        �Һ����״̬.������%Type := Null,
  ��������_In      Number := 0,
  ������_In      Number := 0,
  �����¼id_In    �ٴ������¼.Id%Type := Null
) As
  --���ܣ������������йҺ�(����ԤԼ;ԤԼ�ҺŲ��ۿ�;ԤԼ�Һſۿ�)
  --���:������ʽ_IN:1-��ʾ�Һ�,2-��ʾԤԼ�ҺŲ��ۿ�,3-��ʾԤԼ�Һ�,�ۿ�
  --      ���㷽ʽ_IN:֧�ֶ��ֽ��㷽ʽ,���ֽ��㷽ʽʱ�������ʽ����:���㷽ʽ����1,���,�������,��������־|���㷽ʽ����2,���,�������,��������־|...
  --      �������״̬_In:1-��ʾǿ�Ƽ���Һ����״̬����;�������������Ż�����ʱ��ʱ�ż���.
  --      �Ƿ������豸_In:1-��ʾ��ҽԺ�������豸���д˺����ĵ���,�����豸���ô˺��� ����Ӻ�,��������
  --      ��������_In :0-������������ 1-��ʾ�Ե��ݽ�������,����δ��Ч�ĵ�����Ϣ;2-�������ļ�¼���н���-������������:δ��Ч�ĵ��������пۿ���ɺ���н���
  --      ���ս���_IN:��ʽ="���㷽ʽ|������||....."
  --      ��Ԥ������ids_In:����ö��ַ���,��Ԥ��ʱ��Ч(��Ԥ�������ҵ���������һ��),��Ҫ��ʹ�ü�����Ԥ����
  Err_Item Exception;
  v_Err_Msg  Varchar2(255);
  n_��ӡid   Ʊ�ݴ�ӡ����.Id%Type;
  n_����ֵ   ����Ԥ����¼.���%Type;
  v_�ŶӺ��� Varchar2(20);
  v_�������� �ŶӽкŶ���.��������%Type;
  n_Ԥ��id   ����Ԥ����¼.Id%Type;
  n_�Һ�id   ���˹Һż�¼.Id%Type;
  v_�������� Varchar2(3000);
  v_��ǰ���� Varchar2(150);

  v_���㷽ʽ       ����Ԥ����¼.���㷽ʽ%Type;
  n_������       ����Ԥ����¼.��Ԥ��%Type;
  n_����ϼ�       Number(16, 5);
  n_Ԥ�����       ����Ԥ����¼.��Ԥ��%Type;
  n_��id           ����ɿ����.Id%Type;
  d_�Ŷ�ʱ��       Date;
  n_����           Number;
  n_ͬ����Լһ���� Number(18);
  n_����ԤԼ������ Number(18);
  n_��Լ����       Number(18);

  n_������λ����       Number(18);
  n_�Ƿ񿪷�           Number(1);
  n_Count              Number(18);
  n_�к�               Number(18);
  n_���               ���˹Һż�¼.����%Type;
  n_����id             ������ü�¼.Id%Type;
  n_�۸񸸺�           Number(18);
  n_ԭ��Ŀid           �շ���ĿĿ¼.Id%Type;
  n_ԭ������Ŀid       �շ���ĿĿ¼.Id%Type;
  v_����               ���˹Һż�¼.����%Type;
  n_����id             �ҺŰ���.Id%Type;
  n_ʵ�ս��ϼ�       ������ü�¼.ʵ�ս��%Type;
  n_��������id         ������ü�¼.��������id%Type;
  n_ʵ�ս��           ������ü�¼.ʵ�ս��%Type;
  n_Ӧ�ս��           ������ü�¼.ʵ�ս��%Type;
  n_����id             ���˽��ʼ�¼.Id%Type;
  v_Temp               Varchar2(500);
  n_ԤԼʱ�����       Number;
  n_ԤԼ����           Number;
  d_ʱ�ο�ʼʱ��       Date;
  v_�շ���Ŀids        Varchar2(300);
  n_ԤԼ����           ������λ�ҺŻ���.��Լ��%Type;
  n_����               ���˹Һż�¼.����%Type;
  d_�Ǽ�ʱ��           Date;
  v_����Ա���         ��Ա��.���%Type;
  v_����Ա����         ��Ա��.����%Type;
  n_ԤԼ               Integer;
  v_����               �ҺŰ���ʱ��.����%Type;
  n_���÷�ʱ��         Integer;
  n_�ѹ���             ���˹ҺŻ���.�ѹ���%Type;
  n_��Լ��             ���˹ҺŻ���.��Լ��%Type;
  n_�����ѽ���         ���˹ҺŻ���.��Լ��%Type;
  n_ԤԼ���ɶ���       Number;
  d_Date               Date;
  n_�Һ����           Number;
  v_�Ŷӱ��           �ŶӽкŶ���.�Ŷӱ��%Type;
  v_�Ŷ����           �ŶӽкŶ���.�Ŷ����%Type;
  v_������             �Һ����״̬.������%Type;
  v_��Ų���Ա         �Һ����״̬.����Ա����%Type;
  v_��Ż�����         �Һ����״̬.������%Type;
  n_�������           Number := 0;
  n_������id           �շ��ض���Ŀ.�շ�ϸĿid%Type;
  v_���ʽ           ���˹Һż�¼.ҽ�Ƹ��ʽ%Type;
  v_�ѱ�               ������ü�¼.�ѱ�%Type;
  n_���ηѱ�           Number(3) := 0;
  v_��Ԥ������ids      Varchar2(4000);
  n_Tmp����id          �ҺŰ���.Id%Type;
  n_�ƻ�id             �ҺŰ��żƻ�.Id%Type;
  v_����               ������Ϣ.����%Type;
  n_������λ������ģʽ Number;
  n_�Һ��Ű�ģʽ       Number;
  d_����ʱ��           Date;

  Cursor c_Pati(n_����id ������Ϣ.����id%Type) Is
    Select a.����id, a.����, a.�Ա�, a.����, a.סԺ��, a.�����, a.�ѱ�, a.����, c.���� As ���ʽ
    From ������Ϣ A, ҽ�Ƹ��ʽ C
    Where a.����id = n_����id And a.ҽ�Ƹ��ʽ = c.����(+);

  r_Pati c_Pati%RowType;

  --���α������շѳ�Ԥ���Ŀ���Ԥ���б�
  --��ID�������ȳ��ϴ�δ����ġ�
  Cursor c_Deposit
  (
    n_����id        ������Ϣ.����id%Type,
    v_��Ԥ������ids Varchar2
  ) Is
    Select *
    From (Select a.Id, a.����id, a.��¼״̬, a.Ԥ�����, a.No, Nvl(a.���, 0) As ���
           From ����Ԥ����¼ A,
                (Select NO, Sum(Nvl(a.���, 0)) As ���
                  From ����Ԥ����¼ A
                  Where a.����id Is Null And Nvl(a.���, 0) <> 0 And
                        a.����id In (Select Column_Value From Table(f_Num2list(v_��Ԥ������ids))) And Nvl(a.Ԥ�����, 2) = 1
                  Group By NO
                  Having Sum(Nvl(a.���, 0)) <> 0) B
           Where a.����id Is Null And Nvl(a.���, 0) <> 0 And a.���㷽ʽ Not In (Select ���� From ���㷽ʽ Where ���� = 5) And
                 a.No = b.No And a.����id In (Select Column_Value From Table(f_Num2list(v_��Ԥ������ids))) And
                 Nvl(a.Ԥ�����, 2) = 1
           Union All
           Select 0 As ID, Max(����id) As ����id, ��¼״̬, Ԥ�����, NO, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���
           From ����Ԥ����¼
           Where ��¼���� In (1, 11) And ����id Is Not Null And Nvl(���, 0) <> Nvl(��Ԥ��, 0) And
                 ����id In (Select Column_Value From Table(f_Num2list(v_��Ԥ������ids))) And Nvl(Ԥ�����, 2) = 1 Having
            Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
           Group By ��¼״̬, NO, Ԥ�����)
    Order By Decode(����id, Nvl(n_����id, 0), 0, 1), ID, NO;

  Cursor c_����
  (
    v_����        �ҺŰ���.����%Type,
    d_����ʱ��_In Date
  ) Is
    Select *
    From (With ����ʱ��� As (Select ʱ���
                         From (Select ʱ���,
                                       To_Date(Decode(Sign(��ʼʱ�� - ��ֹʱ��), 1, '3000-01-09 ' || To_Char(��ʼʱ��, 'HH24:MI:SS'),
                                                       '3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS')), 'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                       To_Date(Decode(Sign(��ʼʱ�� - ��ֹʱ��), 1, '3000-01-11 ' || To_Char(��ֹʱ��, 'HH24:MI:SS'),
                                                       '3000-01-10 ' || To_Char(��ֹʱ��, 'HH24:MI:SS')), 'yyyy-mm-dd hh24:mi:ss') As ��ֹʱ��,
                                       To_Date('3000-01-10 ' || To_Char(d_����ʱ��_In, 'HH24:MI:SS'), 'yyyy-mm-dd hh24:mi:ss') As ��ǰʱ��,
                                       To_Date('3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS'), 'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��1,
                                       To_Date('3000-01-10 ' || To_Char(��ֹʱ��, 'HH24:MI:SS'), 'yyyy-mm-dd hh24:mi:ss') As ��ֹʱ��1
                                From ʱ���)
                         Where ��ǰʱ�� Between ��ʼʱ�� And ��ֹʱ��1 Or ��ǰʱ�� Between ��ʼʱ��1 And ��ֹʱ��)
           Select Distinct p.Id, p.����, p.����, p.����id, b.���� As ���ұ���, b.���� As ��������, p.��Ŀid, c.���� As ��Ŀ����, c.���� As ��Ŀ����,
                           p.ҽ��id, d.��� As ҽ�����, p.ҽ������, p.�޺���, p.��Լ��, p.���� As ��, p.��һ As һ, p.�ܶ� As ��, p.���� As ��,
                           p.���� As ��, p.���� As ��, p.���� As ��, p.��ſ���, p.�ƻ�id
           From (Select p.Id, p.����, p.����, p.����id, p.��Ŀid, p.ҽ��id, p.ҽ������, b.�޺���, b.��Լ��, Nvl(p.��������, 0) As ��������, p.����, p.��һ,
                         p.�ܶ�, p.����, p.����, p.����, p.����, p.���﷽ʽ, p.��ſ���,
                         Decode(To_Char(d_����ʱ��_In, 'D'), '1', p.����, '2', p.��һ, '3', p.�ܶ�, '4', p.����, '5', p.����, '6', p.����,
                                 '7', p.����, Null) As �Ű�, Null As �ƻ�id
                  From �ҺŰ��� P, �ҺŰ������� B
                  Where p.ͣ������ Is Null And p.Id = b.����id(+) And
                        b.������Ŀ(+) = Decode(To_Char(d_����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����',
                                           '6', '����', '7', '����', Null) And
                        d_����ʱ��_In Between Nvl(p.��ʼʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                        Nvl(p.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Not Exists
                   (Select 1
                         From �ҺŰ��żƻ�
                         Where ����id = p.Id And (d_����ʱ��_In Between ��Чʱ�� And ʧЧʱ��) And ���ʱ�� Is Not Null) And Not Exists
                   (Select 1
                         From �ҺŰ���ͣ��״̬
                         Where ����id = p.Id And d_����ʱ��_In Between ��ʼֹͣʱ�� And ����ֹͣʱ��) And p.���� = v_����
                  Union All
                  Select c.Id, c.����, c.����, c.����id, p.��Ŀid, p.ҽ��id, p.ҽ������, b.�޺���, b.��Լ��, Nvl(c.��������, 0) As ��������, p.����, p.��һ,
                         p.�ܶ�, p.����, p.����, p.����, p.����, p.���﷽ʽ, p.��ſ���,
                         Decode(To_Char(d_����ʱ��_In, 'D'), '1', p.����, '2', p.��һ, '3', p.�ܶ�, '4', p.����, '5', p.����, '6', p.����,
                                 '7', p.����, Null) As �Ű�, p.Id As �ƻ�id
                  From �ҺŰ��żƻ� P, �ҺŰ��� C, �Һżƻ����� B,
                       (Select Max(a.��Чʱ��) As ��Ч, ����id
                         From �ҺŰ��żƻ� A, �ҺŰ��� B
                         Where a.����id = b.Id And a.���ʱ�� Is Not Null And
                               ����ʱ��_In Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                               Nvl(a.ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) And b.���� = ����_In
                         Group By ����id) E
                  Where p.����id = c.Id And p.Id = b.�ƻ�id(+) And p.��Чʱ�� = e.��Ч And p.����id = e.����id And
                        Nvl(p.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) = Nvl(e.��Ч, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                        b.������Ŀ(+) = Decode(To_Char(d_����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����',
                                           '6', '����', '7', '����', Null) And (d_����ʱ��_In Between p.��Чʱ�� And p.ʧЧʱ��) And
                        p.���ʱ�� Is Not Null And Not Exists
                   (Select 1
                         From �ҺŰ���ͣ��״̬
                         Where ����id = c.Id And d_����ʱ��_In Between ��ʼֹͣʱ�� And ����ֹͣʱ��) And p.���� = v_����) P, ���ű� B, �շ���ĿĿ¼ C,
                ��Ա�� D
           Where p.����id = b.Id And p.ҽ��id = d.Id(+) And p.��Ŀid = c.Id And
                 (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And
                 (Nvl(p.ҽ��id, 0) = 0 Or Exists
                  (Select 1
                   From ��Ա�� Q
                   Where p.ҽ��id = q.Id And (q.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or q.����ʱ�� Is Null))) And Exists
            (Select 1 From ����ʱ��� Where ʱ��� = p.�Ű�))
           Order By ����;


  r_���� c_����%RowType;

  Function Zl_����(����_In �ҺŰ���.����%Type) Return Varchar2 As
    n_���﷽ʽ �ҺŰ���.���﷽ʽ%Type;
    n_����id   �ҺŰ���.Id%Type;
    v_����     ���˹Һż�¼.����%Type;
    v_Rowid    Varchar2(500);
    n_Next     Integer;
    n_First    Integer;
  Begin
  
    If ��������_In = 2 Then
      --�Ե��ݽ��н���,���ȼ���Ƿ��������
      Select Count(Rowid) Into n_���� From ���˹Һż�¼ Where����¼״̬ = 0 And NO = ���ݺ�_In And ����id = ����id_In;
      If n_���� = 0 Then
        v_Err_Msg := '���ݺ�Ϊ(' || ���ݺ�_In || ')�ĵ���,�����ڻ����Ѿ�������!';
        Raise Err_Item;
      End If;
      Select Max(����) Into n_���� From ���˹Һż�¼ Where����¼״̬ = 0 And NO = ���ݺ�_In And ����id = ����id_In;
    End If;
  
    Begin
      Select ID, Nvl(���﷽ʽ, 0) Into n_����id, n_���﷽ʽ From �ҺŰ��� Where ���� = ����_In;
    Exception
      When Others Then
        n_����id := -1;
    End;
  
    If n_����id = -1 Then
      v_Err_Msg := '����(' || ����_In || ')δ�ҵ�!';
      Raise Err_Item;
    End If;
    --0-�����1-ָ�����ҡ�2-��̬���3-ƽ������,��Ӧ������������
    v_���� := Null;
    If n_���﷽ʽ = 1 Then
      --1-ָ������
      Begin
        Select �������� Into v_���� From �ҺŰ������� Where �ű�id = n_����id;
      Exception
        When Others Then
          v_���� := Null;
      End;
    End If;
    If n_���﷽ʽ = 2 Then
      --2-��̬����:�ø��ű���Һ�δ�������ٵ�����
      For c_���� In (Select ��������, Sum(Num) As Num
                   From (Select ��������, 0 As Num
                          From �ҺŰ�������
                          Where �ű�id = n_����id
                          Union All
                          Select ����, Count(����) As Num
                          From ���˹Һż�¼
                          Where Nvl(ִ��״̬, 0) = 0 And ����ʱ�� Between Trunc(Sysdate) And Sysdate And �ű� = ����_In And
                                ���� In (Select �������� From �ҺŰ������� Where �ű�id = n_����id)
                          Group By ����)
                   Group By ��������
                   Order By Num) Loop
        v_���� := c_����.��������;
        Exit;
      End Loop;
    End If;
    If n_���﷽ʽ = 3 Then
    
      --ƽ�������ǰ����=1��ʾ�´�Ӧȡ�ĵ�ǰ����
      n_Next  := 0;
      n_First := 1;
      For c_���� In (Select Rowid As Rid, �ű�id, ��������, ��ǰ���� From �ҺŰ������� Where �ű�id = n_����id) Loop
        If n_First = 1 Then
          v_Rowid := c_����.Rid;
        End If;
        If n_Next = 1 Then
          v_���� := c_����.��������;
          Update �ҺŰ������� Set ��ǰ���� = 1 Where Rowid = c_����.Rid;
          Exit;
        End If;
        If Nvl(c_����.��ǰ����, 0) = 1 Then
          Update �ҺŰ������� Set ��ǰ���� = 0 Where Rowid = c_����.Rid;
          n_Next := 1;
        End If;
      End Loop;
      If v_���� Is Null Then
        Update �ҺŰ������� Set ��ǰ���� = 1 Where Rowid = v_Rowid Returning �������� Into v_����;
      End If;
    End If;
  
    Return v_����;
  End;

  Function Zl_����Ա
  (
    Type_In     Integer,
    Splitstr_In Varchar2
  ) Return Varchar2 As
    n_Step Number(18);
    v_Sub  Varchar2(1000);
    --Type_In:0-��ȡȱʡ����ID;1-��ȡ����Ա���;2-��ȡ����Ա����
    -- SplitStr:��ʽΪ:����ID,��������;��ԱID,��Ա���,��Ա����(��Zl_Identity��ȡ��)
  Begin
    If Type_In = 0 Then
      --ȱʡ����
      n_Step := Instr(Splitstr_In, ',');
      v_Sub  := Substr(Splitstr_In, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 1 Then
      --����Ա����
      n_Step := Instr(Splitstr_In, ';');
      v_Sub  := Substr(Splitstr_In, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 2 Then
      --����Ա����
      n_Step := Instr(Splitstr_In, ';');
      v_Sub  := Substr(Splitstr_In, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      Return v_Sub;
    End If;
  End;

  Procedure Zl_���������Һ�_����_Insert
  (
    ��¼id_In        �ٴ������¼.Id%Type,
    ������ʽ_In      Integer,
    ����id_In        ������ü�¼.����id%Type,
    ����_In          �ҺŰ���.����%Type,
    ����_In          �Һ����״̬.���%Type,
    ���ݺ�_In        ������ü�¼.No%Type,
    Ʊ�ݺ�_In        ������ü�¼.ʵ��Ʊ��%Type,
    ���㷽ʽ_In      Varchar2,
    ժҪ_In          ������ü�¼.ժҪ%Type, --ԤԼ�Һ�ժҪ��Ϣ
    ����ʱ��_In      ������ü�¼.����ʱ��%Type,
    �Ǽ�ʱ��_In      ������ü�¼.�Ǽ�ʱ��%Type,
    ������λ_In      �Һź�����λ.����%Type,
    �ҺŽ��ϼ�_In  ������ü�¼.ʵ�ս��%Type,
    ����id_In        Ʊ��ʹ����ϸ.����id%Type,
    �շ�Ʊ��_In      Number := 0, --�Һ��Ƿ�ʹ���շ�Ʊ��
    ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type,
    ����˵��_In      ����Ԥ����¼.����˵��%Type,
    ԤԼ��ʽ_In      ԤԼ��ʽ.����%Type := Null,
    Ԥ��id_In        ����Ԥ����¼.Id%Type := Null,
    �����id_In      ����Ԥ����¼.�����id%Type := Null,
    �������״̬_In  Number := 0,
    �Ƿ������豸_In  Number := 0,
    ����id_In        ������ü�¼.����id%Type := Null,
    ��������_In      Number := 0,
    ���ս���_In      Varchar2 := Null,
    ��Ԥ��_In        Number := Null,
    ֧������_In      ����Ԥ����¼.����%Type := Null,
    �˺�����_In      Number := 1,
    �ѱ�_In          ������ü�¼.�ѱ�%Type := Null,
    ��Ԥ������ids_In Varchar2 := Null,
    ������_In        �Һ����״̬.������%Type := Null,
    ��������_In      Number := 0,
    ������_In      Number := 0
  ) As
    --���ܣ������������йҺ�(����ԤԼ;ԤԼ�ҺŲ��ۿ�;ԤԼ�Һſۿ�),������Ű�ģʽ��ʹ��
    --���: ������ʽ_IN:1-��ʾ�Һ�,2-��ʾԤԼ�ҺŲ��ۿ�,3-��ʾԤԼ�Һ�,�ۿ�
    --      �������״̬_In:1-��ʾǿ�Ƽ���Һ����״̬����;�������������Ż�����ʱ��ʱ�ż���.
    --      �Ƿ������豸_In:1-��ʾ��ҽԺ�������豸���д˺����ĵ���,�����豸���ô˺��� ����Ӻ�,��������
    --      ��������_In :0-������������ 1-��ʾ�Ե��ݽ�������,����δ��Ч�ĵ�����Ϣ;2-�������ļ�¼���н���-������������:δ��Ч�ĵ��������пۿ���ɺ���н���
    --      ���ս���_IN:��ʽ="���㷽ʽ|������||....."
    --      ��Ԥ������ids_In:����ö��ַ���,��Ԥ��ʱ��Ч(��Ԥ�������ҵ���������һ��),��Ҫ��ʹ�ü�����Ԥ����
    Err_Item Exception;
    v_Err_Msg  Varchar2(255);
    n_��ӡid   Ʊ�ݴ�ӡ����.Id%Type;
    n_����ֵ   ����Ԥ����¼.���%Type;
    v_�ŶӺ��� Varchar2(20);
    v_�������� �ŶӽкŶ���.��������%Type;
    n_Ԥ��id   ����Ԥ����¼.Id%Type;
    n_�Һ�id   ���˹Һż�¼.Id%Type;
    v_�������� Varchar2(3000);
    v_��ǰ���� Varchar2(150);
  
    v_���㷽ʽ       ����Ԥ����¼.���㷽ʽ%Type;
    n_������       ����Ԥ����¼.��Ԥ��%Type;
    n_����ϼ�       Number(16, 5);
    n_Ԥ�����       ����Ԥ����¼.��Ԥ��%Type;
    n_��id           ����ɿ����.Id%Type;
    d_�Ŷ�ʱ��       Date;
    n_����           Number;
    n_ͬ����Լһ���� Number(18);
    n_����ԤԼ������ Number(18);
    n_��Լ����       Number(18);
  
    n_������λ����       Number(18);
    n_�Ƿ񿪷�           Number(1);
    n_Count              Number(18);
    n_�к�               Number(18);
    n_���               ���˹Һż�¼.����%Type;
    n_����id             ������ü�¼.Id%Type;
    n_�۸񸸺�           Number(18);
    n_ԭ��Ŀid           �շ���ĿĿ¼.Id%Type;
    n_ԭ������Ŀid       �շ���ĿĿ¼.Id%Type;
    v_����               ���˹Һż�¼.����%Type;
    n_����id             �ҺŰ���.Id%Type;
    n_ʵ�ս��ϼ�       ������ü�¼.ʵ�ս��%Type;
    n_��������id         ������ü�¼.��������id%Type;
    n_ʵ�ս��           ������ü�¼.ʵ�ս��%Type;
    n_Ӧ�ս��           ������ü�¼.ʵ�ս��%Type;
    n_����id             ���˽��ʼ�¼.Id%Type;
    v_Temp               Varchar2(500);
    v_���㷽ʽ��¼       Varchar2(1000);
    n_ԤԼʱ�����       Number;
    n_��ſ���           �ٴ������¼.�Ƿ���ſ���%Type;
    n_��Լ��             �ٴ������¼.��Լ��%Type;
    n_��Ŀid             �ٴ������¼.��Ŀid%Type;
    n_����id             �ٴ������¼.����id%Type;
    d_��ֹʱ��           �ٴ������¼.��ֹʱ��%Type;
    v_ҽ������           �ٴ������¼.ҽ������%Type;
    n_ҽ��id             �ٴ������¼.ҽ��id%Type;
    n_ԤԼ˳���         �ٴ�������ſ���.ԤԼ˳���%Type;
    n_ԤԼ����           Number;
    d_ʱ�ο�ʼʱ��       Date;
    d_ʱ����ֹʱ��       Date;
    v_�շ���Ŀids        Varchar2(300);
    n_��������־         Number;
    n_ԤԼ����           ������λ�ҺŻ���.��Լ��%Type;
    n_����               ���˹Һż�¼.����%Type;
    d_�Ǽ�ʱ��           Date;
    n_���ʽ��           ����Ԥ����¼.��Ԥ��%Type;
    v_�������           ����Ԥ����¼.�������%Type;
    v_����Ա���         ��Ա��.���%Type;
    v_����Ա����         ��Ա��.����%Type;
    n_ԤԼ               Integer;
    v_�ֽ�               ����Ԥ����¼.���㷽ʽ%Type;
    v_����               �ҺŰ���ʱ��.����%Type;
    n_���÷�ʱ��         Integer;
    n_�ѹ���             ���˹ҺŻ���.�ѹ���%Type;
    n_��Լ��             ���˹ҺŻ���.��Լ��%Type;
    n_�����ѽ���         ���˹ҺŻ���.��Լ��%Type;
    n_ԤԼ���ɶ���       Number;
    n_�޺���             �ٴ������¼.�޺���%Type;
    d_Date               Date;
    n_�Һ����           Number;
    v_�Ŷӱ��           �ŶӽкŶ���.�Ŷӱ��%Type;
    v_�Ŷ����           �ŶӽкŶ���.�Ŷ����%Type;
    v_������             �Һ����״̬.������%Type;
    v_��Ų���Ա         �Һ����״̬.����Ա����%Type;
    v_��Ż�����         �Һ����״̬.������%Type;
    n_�������           Number := 0;
    n_������id           �շ��ض���Ŀ.�շ�ϸĿid%Type;
    v_���ʽ           ���˹Һż�¼.ҽ�Ƹ��ʽ%Type;
    v_�ѱ�               ������ü�¼.�ѱ�%Type;
    n_���ηѱ�           Number(3) := 0;
    v_��Ԥ������ids      Varchar2(4000);
    v_����               ������Ϣ.����%Type;
    n_������λ������ģʽ Number;
    n_ͬ���޺���         Number;
    n_ͬ����Լ��         Number;
    n_���˹Һſ�����     Number;
    n_Exists             Number(5);
  
    Cursor c_Pati(n_����id ������Ϣ.����id%Type) Is
      Select a.����id, a.����, a.�Ա�, a.����, a.סԺ��, a.�����, a.�ѱ�, a.����, c.���� As ���ʽ
      From ������Ϣ A, ҽ�Ƹ��ʽ C
      Where a.����id = n_����id And a.ҽ�Ƹ��ʽ = c.����(+);
  
    r_Pati c_Pati%RowType;
  
    --���α������շѳ�Ԥ���Ŀ���Ԥ���б�
    --��ID�������ȳ��ϴ�δ����ġ�
    Cursor c_Deposit
    (
      n_����id        ������Ϣ.����id%Type,
      v_��Ԥ������ids Varchar2
    ) Is
      Select *
      From (Select a.Id, a.����id, a.��¼״̬, a.Ԥ�����, a.No, Nvl(a.���, 0) As ���
             From ����Ԥ����¼ A,
                  (Select NO, Sum(Nvl(a.���, 0)) As ���
                    From ����Ԥ����¼ A
                    Where a.����id Is Null And Nvl(a.���, 0) <> 0 And
                          a.����id In (Select Column_Value From Table(f_Num2list(v_��Ԥ������ids))) And Nvl(a.Ԥ�����, 2) = 1
                    Group By NO
                    Having Sum(Nvl(a.���, 0)) <> 0) B
             Where a.����id Is Null And Nvl(a.���, 0) <> 0 And a.���㷽ʽ Not In (Select ���� From ���㷽ʽ Where ���� = 5) And
                   a.No = b.No And a.����id In (Select Column_Value From Table(f_Num2list(v_��Ԥ������ids))) And
                   Nvl(a.Ԥ�����, 2) = 1
             Union All
             Select 0 As ID, Max(����id) As ����id, ��¼״̬, Ԥ�����, NO, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���
             From ����Ԥ����¼
             Where ��¼���� In (1, 11) And ����id Is Not Null And Nvl(���, 0) <> Nvl(��Ԥ��, 0) And
                   ����id In (Select Column_Value From Table(f_Num2list(v_��Ԥ������ids))) And Nvl(Ԥ�����, 2) = 1 Having
              Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
             Group By ��¼״̬, NO, Ԥ�����)
      Order By Decode(����id, Nvl(n_����id, 0), 0, 1), ID, NO;
  
    Function Zl_����(��¼id_In �ٴ������¼.Id%Type) Return Varchar2 As
      n_���﷽ʽ �ٴ������¼.���﷽ʽ%Type;
      v_����     ���˹Һż�¼.����%Type;
      v_Rowid    Varchar2(500);
      n_Next     Integer;
      n_First    Integer;
    Begin
    
      If ��������_In = 2 Then
        --�Ե��ݽ��н���,���ȼ���Ƿ��������
        Select Count(Rowid)
        Into n_����
        From ���˹Һż�¼ Where����¼״̬ = 0 And NO = ���ݺ�_In And ����id = ����id_In;
        If n_���� = 0 Then
          v_Err_Msg := '���ݺ�Ϊ(' || ���ݺ�_In || ')�ĵ���,�����ڻ����Ѿ�������!';
          Raise Err_Item;
        End If;
        Select Max(����) Into n_���� From ���˹Һż�¼ Where����¼״̬ = 0 And NO = ���ݺ�_In And ����id = ����id_In;
      End If;
    
      Begin
        Select Nvl(���﷽ʽ, 0) Into n_���﷽ʽ From �ٴ������¼ Where ID = ��¼id_In;
      Exception
        When Others Then
          v_Err_Msg := '�����¼(' || ��¼id_In || ')δ�ҵ�!';
          Raise Err_Item;
      End;
    
      --0-�����1-ָ�����ҡ�2-��̬���3-ƽ������,��Ӧ������������
      v_���� := Null;
      If n_���﷽ʽ = 1 Then
        --1-ָ������
        Begin
          Select b.���� Into v_���� From �ٴ��������Ҽ�¼ A, �������� B Where a.����id = b.Id And a.��¼id = ��¼id_In;
        Exception
          When Others Then
            v_���� := Null;
        End;
      End If;
      If n_���﷽ʽ = 2 Then
        --2-��̬����:�ø��ű���Һ�δ�������ٵ�����
        For c_���� In (Select ��������, Sum(Num) As Num
                     From (Select b.���� As ��������, 0 As Num
                            From �ٴ��������Ҽ�¼ A, �������� B
                            Where a.����id = b.Id And a.��¼id = ��¼id_In
                            Union All
                            Select ����, Count(����) As Num
                            From ���˹Һż�¼
                            Where Nvl(ִ��״̬, 0) = 0 And ����ʱ�� Between Trunc(Sysdate) And Sysdate And �ű� = ����_In And
                                  ���� In (Select d.����
                                         From �ٴ��������Ҽ�¼ C, �������� D
                                         Where c.����id = d.Id And c.��¼id = ��¼id_In)
                            Group By ����)
                     Group By ��������
                     Order By Num) Loop
          v_���� := c_����.��������;
          Exit;
        End Loop;
      End If;
      If n_���﷽ʽ = 3 Then
        --ƽ�������ǰ����=1��ʾ�´�Ӧȡ�ĵ�ǰ����
        n_Next  := 0;
        n_First := 1;
        For c_���� In (Select a.Rowid As Rid, b.���� As ��������, a.��ǰ����
                     From �ٴ��������Ҽ�¼ A, �������� B
                     Where a.����id = b.Id And a.��¼id = ��¼id_In) Loop
          If n_First = 1 Then
            v_Rowid := c_����.Rid;
          End If;
          If n_Next = 1 Then
            v_���� := c_����.��������;
            Update �ٴ��������Ҽ�¼ Set ��ǰ���� = 1 Where Rowid = c_����.Rid;
            Exit;
          End If;
          If Nvl(c_����.��ǰ����, 0) = 1 Then
            Update �ٴ��������Ҽ�¼ Set ��ǰ���� = 0 Where Rowid = c_����.Rid;
            n_Next := 1;
          End If;
        End Loop;
        If v_���� Is Null Then
          Update �ٴ��������Ҽ�¼ Set ��ǰ���� = 1 Where Rowid = v_Rowid Returning ����id Into v_����;
          Select ���� Into v_���� From �������� Where ID = v_����;
        End If;
      End If;
      Return v_����;
    End;
  
    Function Zl_����Ա
    (
      Type_In     Integer,
      Splitstr_In Varchar2
    ) Return Varchar2 As
      n_Step Number(18);
      v_Sub  Varchar2(1000);
      --Type_In:0-��ȡȱʡ����ID;1-��ȡ����Ա���;2-��ȡ����Ա����
      -- SplitStr:��ʽΪ:����ID,��������;��ԱID,��Ա���,��Ա����(��Zl_Identity��ȡ��)
    Begin
      If Type_In = 0 Then
        --ȱʡ����
        n_Step := Instr(Splitstr_In, ',');
        v_Sub  := Substr(Splitstr_In, 1, n_Step - 1);
        Return v_Sub;
      End If;
      If Type_In = 1 Then
        --����Ա����
        n_Step := Instr(Splitstr_In, ';');
        v_Sub  := Substr(Splitstr_In, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, 1, n_Step - 1);
        Return v_Sub;
      End If;
      If Type_In = 2 Then
        --����Ա����
        n_Step := Instr(Splitstr_In, ';');
        v_Sub  := Substr(Splitstr_In, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        Return v_Sub;
      End If;
    End;
  
  Begin
    v_��Ԥ������ids := Nvl(��Ԥ������ids_In, ����id_In);
  
    Begin
      Select ���� Into v_�ֽ� From ���㷽ʽ Where ���� = 1;
    Exception
      When Others Then
        v_�ֽ� := '�ֽ�';
    End;
  
    If �ѱ�_In Is Null Then
      Select Zl_Custom_Getpatifeetype(1, ����id_In) Into v_�ѱ� From Dual;
    Else
      v_�ѱ� := �ѱ�_In;
    End If;
    If v_�ѱ� Is Null Then
      n_���ηѱ� := 1;
      Select ���� Into v_�ѱ� From �ѱ� Where ȱʡ��־ = 1 And Rownum < 2;
    End If;
    Update ������Ϣ Set �ѱ� = v_�ѱ� Where ����id = ����id_In;
  
    If ��������_In = 1 Then
      Select Zl_Age_Calc(����id_In) Into v_���� From Dual;
      If v_���� Is Not Null Then
        Update ������Ϣ Set ���� = v_���� Where ����id = ����id_In;
      End If;
    End If;
    --��ȡ��ǰ��������
    If ������_In Is Not Null Then
      v_������ := ������_In;
    Else
      Select Terminal Into v_������ From V$session Where Audsid = Userenv('sessionid');
    End If;
    n_ʵ�ս��ϼ� := 0;
  
    Select Count(*) + 1
    Into n_�Һ����
    From ���˹Һż�¼
    Where �����¼id = ��¼id_In And �Ǽ�ʱ�� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In + 1) - 1 / 24 / 60 / 60;
  
    --����ID,��������;��ԱID,��Ա���,��Ա����
    v_Temp := Zl_Identity(0);
    If Nvl(v_Temp, ' ') = ' ' Then
      v_Err_Msg := '��ǰ������Աδ���ö�Ӧ����Ա��ϵ,���ܼ�����';
      Raise Err_Item;
    End If;
  
    If �Ǽ�ʱ��_In Is Null Then
      d_�Ǽ�ʱ�� := Sysdate;
    Else
      d_�Ǽ�ʱ�� := �Ǽ�ʱ��_In;
    End If;
    If Trunc(Sysdate) > Trunc(����ʱ��_In) Then
      v_Err_Msg := '���ܹ���ǰ�ĺ�(' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || ')��';
      Raise Err_Item;
    End If;
  
    v_Temp           := Nvl(zl_GetSysParameter('����ͬ������', 1111), '0,0|0,0');
    n_ͬ���޺���     := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
    v_Temp           := Substr(v_Temp, Instr(v_Temp, '|') + 1);
    n_ͬ����Լ��     := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
    v_Temp           := Nvl(zl_GetSysParameter('����ԤԼ������', 1111), '0|0');
    n_����ԤԼ������ := Substr(v_Temp, Instr(v_Temp, '|') + 1);
    n_���˹Һſ����� := Substr(v_Temp, 1, Instr(v_Temp, '|') - 1);
    n_��������id     := To_Number(Zl_����Ա(0, v_Temp));
    v_����Ա���     := Zl_����Ա(1, v_Temp);
    v_����Ա����     := Zl_����Ա(2, v_Temp);
    n_��id           := Zl_Get��id(v_����Ա����);
  
    If ������ʽ_In <> 1 Then
      --ԤԼ����Ƿ���Ӻ�����λ����
      --��������˺�����λ���� ��
      Begin
        Select 1
        Into n_������λ����
        From �ٴ�����Һſ��Ƽ�¼
        Where ��¼id = ��¼id_In And ���� = 1 And ���� = 1 And ���Ʒ�ʽ <> 4 And Rownum < 2;
      Exception
        When Others Then
          n_������λ���� := 0;
      End;
    End If;
  
    If ������ʽ_In <> 2 Then
      v_���� := Zl_����(��¼id_In);
    End If;
  
    --��Ϊ�����а��ձ൥�ݺŹ���,�չҺ������ܳ���10000��,����Ҫ���ΨһԼ����
    Select Count(*) Into n_Count From ������ü�¼ Where ��¼���� = 4 And ��¼״̬ In (1, 3) And NO = ���ݺ�_In;
    If n_Count <> 0 Then
      v_Err_Msg := '�Һŵ��ݺ��ظ�,���ܱ��棡' || Chr(13) || '���ʹ���˰���˳����,���չҺ������ܳ���10000�˴Ρ�';
      Raise Err_Item;
    End If;
  
    Open c_Pati(����id_In);
    n_Count := 0;
    Begin
      Fetch c_Pati
        Into r_Pati;
    Exception
      When Others Then
        n_Count := -1;
    End;
    If n_Count = -1 Then
      v_Err_Msg := '����δ�ҵ������ܼ�����';
      Raise Err_Item;
    End If;
  
    Begin
      Select Nvl(�Ƿ��ʱ��, 0), �޺���, �ѹ���, �����ѽ���, ��Լ��, �Ƿ���ſ���, ��Լ��, ��Ŀid, ����id, ҽ��id, ҽ������
      Into n_���÷�ʱ��, n_�޺���, n_�ѹ���, n_�����ѽ���, n_��Լ��, n_��ſ���, n_��Լ��, n_��Ŀid, n_����id, n_ҽ��id, v_ҽ������
      From �ٴ������¼
      Where ID = ��¼id_In;
    Exception
      When Others Then
        n_Count := -1;
    End;
    If n_Count = -1 Then
      v_Err_Msg := '�úű�û����' || To_Char(����ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '�н��а��š�';
      Raise Err_Item;
    End If;
  
    --�Բ������ƽ��м��
    --����ԤԼ���ۿ�ʱ���м��
    If ������ʽ_In = 2 Then
      If Nvl(n_ͬ����Լ��, 0) <> 0 Or Nvl(n_����ԤԼ������, 0) <> 0 Then
        n_��Լ���� := 0;
        For c_Chkitem In (Select Distinct ִ�в���id
                          From ���˹Һż�¼
                          Where ��¼״̬ = 1 And ��¼���� = 2 And ԤԼʱ�� Between Trunc(����ʱ��_In) And
                                Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60) Loop
          n_��Լ���� := n_��Լ���� + 1;
        End Loop;
        If n_��Լ���� >= Nvl(n_����ԤԼ������, 0) And Nvl(n_����ԤԼ������, 0) > 0 Then
          v_Err_Msg := 'ͬһ�������ͬʱ��ԤԼ[' || Nvl(n_����ԤԼ������, 0) || ']������,������ԤԼ��';
          Raise Err_Item;
        End If;
      
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ��¼״̬ = 1 And ��¼���� = 2 And ԤԼʱ�� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And
              ִ�в���id = n_����id;
        If n_Count >= Nvl(n_ͬ����Լ��, 0) And Nvl(n_ͬ����Լ��, 0) > 0 Then
          v_Err_Msg := '�ò����Ѿ��ڸÿ���ԤԼ��' || n_Count || '��,������ԤԼ��';
          Raise Err_Item;
        End If;
      End If;
    Else
      If Nvl(n_ͬ���޺���, 0) <> 0 Or Nvl(n_���˹Һſ�����, 0) <> 0 Then
        n_��Լ���� := 0;
        For c_Chkitem In (Select Distinct ִ�в���id
                          From ���˹Һż�¼
                          Where ��¼״̬ = 1 And ��¼���� = 1 And ����ʱ�� Between Trunc(����ʱ��_In) And
                                Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60) Loop
          n_��Լ���� := n_��Լ���� + 1;
        End Loop;
        If n_��Լ���� >= Nvl(n_���˹Һſ�����, 0) And Nvl(n_���˹Һſ�����, 0) > 0 Then
          v_Err_Msg := 'ͬһ�������ͬʱ�ܹҺ�[' || Nvl(n_���˹Һſ�����, 0) || ']������,�����ٹҺţ�';
          Raise Err_Item;
        End If;
      
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ��¼״̬ = 1 And ��¼���� = 1 And ����ʱ�� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And
              ִ�в���id = n_����id;
        If n_Count >= Nvl(n_ͬ���޺���, 0) And Nvl(n_ͬ���޺���, 0) > 0 Then
          v_Err_Msg := '�ò����Ѿ��ڸÿ��ҹҺ���' || n_Count || '��,�����ٹҺţ�';
          Raise Err_Item;
        End If;
      End If;
    End If;
  
    d_Date         := Null;
    d_ʱ�ο�ʼʱ�� := Null;
  
    If Nvl(n_�޺���, 0) >= 0 Or n_�޺��� Is Null Then
      If n_���÷�ʱ�� = 1 Then
        If Nvl(n_��ſ���, 0) = 1 Then
          If Nvl(�Ƿ������豸_In, 0) = 0 Then
            Select Count(*), Max(��ʼʱ��)
            Into n_Count, d_ʱ�ο�ʼʱ��
            From �ٴ�������ſ���
            Where ��¼id = ��¼id_In And ��� = Nvl(����_In, 0);
          
            v_Temp := '�Һ�';
            If ������ʽ_In > 1 Then
              v_Temp := 'ԤԼ�Һ�';
            End If;
          
            If n_Count = 0 Then
              v_Err_Msg := '�ű�Ϊ' || ����_In || '�ĹҺŰ����в��������Ϊ' || Nvl(����_In, 0) || '�İ���,������' || v_Temp || '��';
              Raise Err_Item;
            End If;
          End If;
        
          --�����,����ѡ��Һ�
          If Trunc(Sysdate) = Trunc(����ʱ��_In) Then
            --�ҵ���ĺ�
            v_Temp := To_Char(Sysdate, 'yyyy-mm-dd') || ' ';
            For v_ʱ�� In (Select To_Date(v_Temp || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                To_Date(To_Char(Sysdate + Decode(Sign(��ʼʱ�� - ��ֹʱ��), 1, 1, 0), 'yyyy-mm-dd') || ' ' ||
                                         To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ��ֹʱ��, ����, �Ƿ�ԤԼ
                         From �ٴ�������ſ���
                         Where ��¼id = ��¼id_In And ��� = Nvl(����_In, 0)) Loop
              If Sysdate > v_ʱ��.��ֹʱ�� Then
                v_Err_Msg := '�ű�Ϊ' || ����_In || '������Ϊ' || Nvl(����_In, 0) || '�İ���,�Ѿ�����ʱ��,������' || v_Temp || '��';
                Raise Err_Item;
              End If;
            End Loop;
          End If;
        Elsif ������ʽ_In > 1 Then
          --δ������ŵ�,��Ҫ���ԤԼ�����
          n_Count := 0;
          For v_ʱ�� In (Select ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ
                       From �ٴ�������ſ���
                       Where ��¼id = ��¼id_In And
                             (('3000-01-10 ' || To_Char(����ʱ��_In, 'HH24:MI:SS') Between
                             Decode(Sign(��ʼʱ�� - ��ֹʱ�� - 1 / 24 / 60 / 60), 1,
                                      '3000-01-09 ' || To_Char(��ʼʱ��, 'HH24:MI:SS'),
                                      '3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS')) And
                             '3000-01-10 ' || To_Char(��ֹʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS')) Or
                             ('3000-01-10 ' || To_Char(����ʱ��_In, 'HH24:MI:SS') Between
                             '3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS') And
                             Decode(Sign(��ʼʱ�� - ��ֹʱ�� - 1 / 24 / 60 / 60), 1,
                                      '3000-01-11 ' || To_Char(��ֹʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS'),
                                      '3000-01-10 ' || To_Char(��ֹʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS'))))) Loop
            n_ԤԼʱ����� := v_ʱ��.���;
            d_ʱ�ο�ʼʱ�� := v_ʱ��.��ʼʱ��;
            d_ʱ����ֹʱ�� := v_ʱ��.��ֹʱ��;
          
            Select Count(*), Max(���), Max(ԤԼ˳���) + 1
            Into n_Count, n_ԤԼ����, n_ԤԼ˳���
            From �ٴ�������ſ���
            Where ��¼id = ��¼id_In And Nvl(�Һ�״̬, 0) Not In (0, 4, 5);
          
            If Nvl(n_Count, 0) > Nvl(v_ʱ��.����, 0) And ��������_In <> 2 Then
              v_Err_Msg := '�ű�Ϊ' || ����_In || '�ĹҺŰ�������' || To_Char(v_ʱ��.��ʼʱ��, 'hh24:mi:ss') || '��' ||
                           To_Char(v_ʱ��.��ֹʱ��, 'hh24:mi:ss') || '���ֻ��ԤԼ' || Nvl(v_ʱ��.����, 0) || '��,�����ٽ���ԤԼ�Һţ�';
              Raise Err_Item;
            End If;
            n_Count := 1;
          End Loop;
        
          If n_Count = 0 Then
            v_Err_Msg := '�ű�Ϊ' || ����_In || '�ĹҺŰ�����û����صİ��żƻ�(' || To_Char(����ʱ��_In, 'YYYY-mm-dd HH24:MI:SS') ||
                         '),���ܽ���ԤԼ�Һţ�';
            Raise Err_Item;
          End If;
        End If;
      End If;
    End If;
  
    If ������ʽ_In = 1 And ��������_In <> 2 Then
      --�ҺŹ���:
      --  �ѹ������ܴ����޺���
      If n_�ѹ��� >= Nvl(n_�޺���, 0) And n_�޺��� Is Not Null Then
        v_Err_Msg := '�úű�����Ѵﵽ�޺��� ' || Nvl(n_�޺���, 0) || '�����ٹҺţ�';
        Raise Err_Item;
      End If;
    End If;
  
    If ������ʽ_In > 1 Then
      --ԤԼ����ؼ��
      --����:
      --   1.����Լ���ܳ�����Լ��
      --   2.����Ƿ�����ʱ�ε�
      If n_��Լ�� >= Nvl(n_��Լ��, 0) And Nvl(n_��Լ��, 0) <> 0 And n_��Լ�� Is Not Null And ��������_In <> 2 Then
        v_Err_Msg := '�úű��Ѵﵽ��Լ�� ' || Nvl(n_��Լ��, 0) || '������ԤԼ�Һţ�';
        Raise Err_Item;
      End If;
      If ԤԼ��ʽ_In Is Not Null Then
        Select Zl_ԤԼ��ʽ_Check(��¼id_In, ����_In, ԤԼ��ʽ_In) Into n_Exists From Dual;
        If n_Exists = 0 Then
          v_Err_Msg := '�����ԤԼ��ʽ' || ԤԼ��ʽ_In || 'ԤԼ�����ﵽ����,���ܼ�����';
          Raise Err_Item;
        End If;
      End If;
    End If;
    If n_������λ���� > 0 And ������ʽ_In <> 1 And ������λ_In Is Not Null Then
      If Nvl(n_��ſ���, 0) = 1 And Nvl(����_In, 0) = 0 Then
        v_Err_Msg := '��ǰ����ʹ������ſ���,��ȷ������ҪԤԼ�����,���ܼ�����';
        Raise Err_Item;
      End If; --Nvl(r_����.��ſ���, 0) =0
    
      n_��� := Case
                When Nvl(n_��ſ���, 0) = 1 Or n_���÷�ʱ�� = 1 And ������ʽ_In > 1 Then
                 Nvl(����_In, 0)
                Else
                 0
              End;
    
      --������λ����ģʽ
      Select Nvl(���Ʒ�ʽ, 0)
      Into n_������λ������ģʽ
      From �ٴ�����Һſ��Ƽ�¼
      Where ��¼id = ��¼id_In And ���� = ������λ_In And ���� = 1 And ���� = 1 And Rownum < 2;
    
      If n_������λ������ģʽ = 0 Then
        v_Err_Msg := '��ǰ����(' || Nvl(����_In, 0) || 'δ����' || ������λ_In || '��ԤԼ,���ܼ�����';
        Raise Err_Item;
      End If;
      If n_������λ������ģʽ = 1 Or n_������λ������ģʽ = 2 Then
        Select ����
        Into n_Count
        From �ٴ�����Һſ��Ƽ�¼
        Where ��¼id = ��¼id_In And ���� = ������λ_In And ���� = 1 And ���� = 1;
        If n_������λ������ģʽ = 1 Then
          n_Count := Round(Nvl(n_��Լ��, n_�޺���) * n_Count / 100);
        End If;
        Select Count(1)
        Into n_Exists
        From ���˹Һż�¼
        Where ��¼״̬ = 1 And �����¼id = ��¼id_In And ������λ = ������λ_In;
        If n_Exists >= n_Count Then
          v_Err_Msg := '��ǰ����(' || Nvl(����_In, 0) || '�Ѿ�����' || ������λ_In || '��ԤԼ����,���ܼ�����';
          Raise Err_Item;
        End If;
      End If;
      --������ż��
      If n_������λ������ģʽ = 3 Then
        For c_������λ In (Select ���, ����
                       From �ٴ�����Һſ��Ƽ�¼
                       Where ��¼id = ��¼id_In And ���� = ������λ_In And ���� = 1 And ���� = 1 And ��� = ����_In) Loop
          If n_��ſ��� = 1 Then
            Begin
              Select 1
              Into n_Count
              From �ٴ�������ſ���
              Where ��¼id = ��¼id_In And ��� = ����_In And Nvl(�Һ�״̬, 0) = 0;
            Exception
              When Others Then
                n_Count := 0;
            End;
            If n_Count = 1 Then
              n_�Ƿ񿪷� := 1;
            Else
              v_Err_Msg := '��ǰ���(' || Nvl(����_In, 0) || '�Ѿ�����' || ������λ_In || '��ԤԼ����,���ܼ�����';
              Raise Err_Item;
            End If;
          Else
            Select Count(1)
            Into n_Count
            From �ٴ�������ſ���
            Where ��¼id = ��¼id_In And ��� = ����_In And ԤԼ˳��� Is Not Null And Nvl(�Һ�״̬, 0) <> 0;
            If n_Count >= c_������λ.���� Then
              v_Err_Msg := '��ǰ���(' || Nvl(����_In, 0) || '�Ѿ�����' || ������λ_In || '��ԤԼ����,���ܼ�����';
              Raise Err_Item;
            Else
              n_�Ƿ񿪷� := 1;
            End If;
          End If;
        End Loop;
      
        If Nvl(n_�Ƿ񿪷�, 0) = 0 Then
          v_Err_Msg := '��ǰ���(' || Nvl(����_In, 0) || 'δ����,���ܼ�����';
          Raise Err_Item;
        End If;
      End If;
    End If;
  
    --����޺�������Լ��
    n_�к�         := 1;
    n_ԭ��Ŀid     := 0;
    n_ԭ������Ŀid := 0;
    n_ʵ�ս��ϼ� := 0;
    If ��������_In <> 1 Then
      If ������ʽ_In <> 2 Then
        If Nvl(����id_In, 0) = 0 Then
          --����Ӧ�ó�����
          Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
        Else
          n_����id := ����id_In;
        End If;
      Else
        n_����id := Null;
      End If;
    End If;
  
    If Nvl(������_In, 0) = 1 Then
      Begin
        Select �շ�ϸĿid Into n_������id From �շ��ض���Ŀ Where �ض���Ŀ = '������';
        v_�շ���Ŀids := n_��Ŀid || ',' || n_������id;
      Exception
        When Others Then
          v_Err_Msg := '����ȷ��������,�Һ�ʧ��!';
          Raise Err_Item;
      End;
    Else
      v_�շ���Ŀids := n_��Ŀid;
    End If;
  
    For c_Item In (Select 1 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, 1 As ����,
                          c.Id As ������Ŀid, c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, Null As ��������
                   From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C
                   Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = n_��Ŀid And Sysdate Between b.ִ������ And
                         Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))
                   Union All
                   Select 2 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, 1 As ����,
                          c.Id As ������Ŀid, c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, Null As ��������
                   From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C
                   Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = n_������id And Sysdate Between b.ִ������ And
                         Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))
                   Union All
                   Select 3 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, d.�������� As ����,
                          c.Id As ������Ŀid, c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, 1 As ��������
                   From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, �շѴ�����Ŀ D
                   Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = d.����id And
                         d.����id In (Select Column_Value From Table(f_Str2list(v_�շ���Ŀids))) And Sysdate Between b.ִ������ And
                         Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))
                   Order By ����, ��Ŀ����, �������) Loop
      n_�۸񸸺� := Null;
      If n_ԭ��Ŀid = c_Item.��Ŀid Then
        If n_ԭ������Ŀid <> c_Item.������Ŀid Then
          n_�۸񸸺� := n_�к�;
        End If;
        n_ԭ������Ŀid := c_Item.������Ŀid;
      End If;
      n_ԭ��Ŀid := c_Item.��Ŀid;
      n_Ӧ�ս�� := Round(c_Item.���� * c_Item.����, 5);
      n_ʵ�ս�� := n_Ӧ�ս��;
      If Nvl(c_Item.���ηѱ�, 0) <> 1 And n_���ηѱ� = 0 Then
        --����:
        v_Temp     := Zl_Actualmoney(r_Pati.�ѱ�, c_Item.��Ŀid, c_Item.������Ŀid, n_Ӧ�ս��);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ':') + 1);
        n_ʵ�ս�� := Zl_To_Number(v_Temp);
      End If;
      n_ʵ�ս��ϼ� := Nvl(n_ʵ�ս��ϼ�, 0) + n_ʵ�ս��;
    
      --�������ݲ���������
      If ��������_In <> 1 Then
        --�������˹Һŷ���(���ܵ����ǻ������������)
        Select ���˷��ü�¼_Id.Nextval Into n_����id From Dual;
        --:������ʽ_IN:1-��ʾ�Һ�,2-��ʾԤԼ�ҺŲ��ۿ�,3-��ʾԤԼ�Һ�,�ۿ�
        Insert Into ������ü�¼
          (ID, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, NO, ʵ��Ʊ��, �����־, �Ӱ��־, ���ӱ�־, ��ҩ����, ����id, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, ���˿���id,
           �շ����, ���㵥λ, �շ�ϸĿid, ������Ŀid, �վݷ�Ŀ, ����, ����, ��׼����, Ӧ�ս��, ʵ�ս��, ���ʽ��, ����id, ���ʷ���, ��������id, ������, ������, ִ�в���id, ִ����,
           ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ���մ���id, ������Ŀ��, ���ձ���, ͳ����, ժҪ, ����, �ɿ���id)
        Values
          (n_����id, 4, Decode(������ʽ_In, 2, 0, 1), n_�к�, n_�۸񸸺�, c_Item.��������, ���ݺ�_In, Ʊ�ݺ�_In, 1, Null, Null,
           Decode(������ʽ_In, 2, To_Char(����_In), v_����), r_Pati.����id, r_Pati.�����, r_Pati.���ʽ, r_Pati.����, r_Pati.�Ա�,
           r_Pati.����, r_Pati.�ѱ�, n_����id, c_Item.���, ����_In, c_Item.��Ŀid, c_Item.������Ŀid, c_Item.�վݷ�Ŀ, 1, c_Item.����,
           c_Item.����, n_Ӧ�ս��, n_ʵ�ս��, Decode(������ʽ_In, 2, Null, n_ʵ�ս��), n_����id, 0, n_��������id, v_����Ա����,
           Decode(������ʽ_In, 2, v_����Ա����, Null), n_����id, v_ҽ������, v_����Ա���, v_����Ա����, ����ʱ��_In, d_�Ǽ�ʱ��, Null, 0, Null, Null,
           ժҪ_In, ԤԼ��ʽ_In, Decode(������ʽ_In, 2, Null, n_��id));
      End If;
      n_�к� := n_�к� + 1;
    
    End Loop;
  
    If Round(Nvl(�ҺŽ��ϼ�_In, 0), 5) <> Round(Nvl(n_ʵ�ս��ϼ�, 0), 5) Then
      v_Err_Msg := '���ιҺŽ���ȷ,��������ΪҽԺ�����˼۸�,�����»�ȡ�Һ��շ���Ŀ�ļ۸�,���ܼ�����';
      Raise Err_Item;
    End If;
  
    If n_���÷�ʱ�� = 1 Then
      d_Date := To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(d_ʱ�ο�ʼʱ��, 'hh24:mi:ss'),
                        'yyyy-mm-dd hh24:mi:ss');
    Else
      d_Date := Trunc(����ʱ��_In);
    End If;
  
    --���¹Һ����״̬
    If ��������_In <> 2 Then
      n_���� := ����_In;
    End If;
    Begin
      Select 1
      Into n_Count
      From �ٴ�������ſ���
      Where ��¼id = ��¼id_In And ��� = n_���� And Nvl(�Һ�״̬, 0) Not In (0, 5);
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count = 1 Then
      If n_���÷�ʱ�� = 0 And Nvl(n_��ſ���, 0) = 1 Then
        n_���� := Null;
      End If;
      If n_���÷�ʱ�� = 1 And Nvl(n_��ſ���, 0) = 1 Then
        v_Err_Msg := '��ǰ����ѱ�ʹ�ã�������ѡ��һ����ţ�';
        Raise Err_Item;
      End If;
    End If;
    n_Count := 0;
    If n_���÷�ʱ�� = 0 And Nvl(n_��ſ���, 0) = 1 And n_���� Is Null And ��������_In <> 2 Then
      Select Nvl(Min(���), 0)
      Into n_����
      From �ٴ�������ſ���
      Where ��¼id = ��¼id_In And Nvl(�Һ�״̬, 0) = 5 And ����Ա���� = v_����Ա���� And ����վ���� = v_������;
      If n_���� = 0 Then
        Select Nvl(Max(���), 0) Into n_���� From �ٴ�������ſ��� Where ��¼id = ��¼id_In And Nvl(�Һ�״̬, 0) = 0;
        If n_���� = 0 Then
          Select Nvl(Max(���), 0) + 1 Into n_���� From �ٴ�������ſ��� Where ��¼id = ��¼id_In;
        End If;
      End If;
    End If;
  
    If n_���÷�ʱ�� = 1 And ��������_In <> 2 Then
      If ������ʽ_In > 1 And Nvl(n_��ſ���, 0) = 0 Then
        --����:ԤԼʱ�����||ԤԼ��
        If Nvl(n_ԤԼ����, 0) = 0 Then
          v_Temp := Nvl(n_��Լ��, 0);
          v_Temp := LTrim(RTrim(v_Temp));
          v_Temp := LPad(Nvl(n_ԤԼ����, 0) + 1, Length(v_Temp), '0');
          v_Temp := n_ԤԼʱ����� || v_Temp;
          n_���� := To_Number(v_Temp);
        Else
          n_���� := n_ԤԼ���� + 1;
        End If;
      End If;
    End If;
  
    If Nvl(n_��ſ���, 0) = 1 Or (������ʽ_In > 1 And n_���÷�ʱ�� = 1) Or �������״̬_In = 1 Then
      --������ŵĴ���
      Begin
        Select ����Ա����, ����վ����
        Into v_��Ų���Ա, v_��Ż�����
        From �ٴ�������ſ���
        Where �Һ�״̬ = 5 And ��¼id = ��¼id_In And ��� = n_����;
        n_������� := 1;
      Exception
        When Others Then
          v_��Ų���Ա := Null;
          v_��Ż����� := Null;
          n_�������   := 0;
      End;
      If n_������� = 0 Then
        If n_���÷�ʱ�� = 1 And n_��ſ��� = 0 Then
          Insert Into �ٴ�������ſ���
            (��¼id, ���, ԤԼ˳���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����, ����, ����Ա����, ��ע)
            Select ��¼id_In, n_ԤԼʱ�����, n_ԤԼ˳���, d_ʱ�ο�ʼʱ��, d_ʱ����ֹʱ��, 1, Decode(������ʽ_In, 1, 0, 1), Decode(������ʽ_In, 2, 2, 1),
                   1, ������λ_In, v_����Ա����, n_����
            From Dual;
        Else
          Update �ٴ�������ſ���
          Set �Һ�״̬ = Decode(������ʽ_In, 2, 2, 1), ����Ա���� = v_����Ա����
          Where ��¼id = ��¼id_In And ��� = n_����;
        End If;
        If Sql%RowCount = 0 Then
          Begin
            If n_���÷�ʱ�� = 1 Then
              --��ʱ��
              If n_��ſ��� = 1 Then
                --��ſ���
                Select Max(��ֹʱ��) Into d_��ֹʱ�� From �ٴ�������ſ��� Where ��¼id = ��¼id_In;
                If Sysdate > d_��ֹʱ�� Then
                  d_��ֹʱ�� := Sysdate;
                End If;
                Insert Into �ٴ�������ſ���
                  (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����, ����, ����Ա����)
                  Select ��¼id_In, n_����, d_��ֹʱ��, d_��ֹʱ��, 1, Decode(������ʽ_In, 1, 0, 1), Decode(������ʽ_In, 2, 2, 1), 1,
                         ������λ_In, v_����Ա����
                  From Dual;
              Else
                --��ʱ��,����ſ���
                Null;
              End If;
            Else
              --����ʱ��
              Insert Into �ٴ�������ſ���
                (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����, ����, ����Ա����)
                Select ��¼id_In, n_����, ��ʼʱ��, ��ֹʱ��, 1, Decode(������ʽ_In, 1, 0, 1), Decode(������ʽ_In, 2, 2, 1), 1, ������λ_In,
                       v_����Ա����
                From �ٴ�������ſ���
                Where ��¼id = ��¼id_In And ��� = 1;
            End If;
          Exception
            When Others Then
              v_Err_Msg := '���' || n_���� || '�ѱ�ʹ��,������ѡ��һ�����.';
              Raise Err_Item;
          End;
        End If;
      Else
        If v_����Ա���� <> v_��Ų���Ա Or v_������ <> v_��Ż����� Then
          v_Err_Msg := '���' || n_���� || '�ѱ�����' || v_������ || '����,������ѡ��һ�����.';
          Raise Err_Item;
        Else
          Update �ٴ�������ſ���
          Set �Һ�״̬ = Decode(������ʽ_In, 2, 2, 1), ����ʱ�� = Null
          Where ��¼id = ��¼id_In And ��� = n_���� And �Һ�״̬ = 5 And ����Ա���� = v_����Ա���� And ����վ���� = v_������;
        End If;
      End If;
    End If;
  
    --�������ݲ������κ� ����
    If ������ʽ_In <> 2 And ��������_In <> 1 Then
      --�Һ�,ԤԼ�Һ��Ѿ��ۿ��
      n_Ԥ��id := Ԥ��id_In;
      If Nvl(n_Ԥ��id, 0) = 0 Then
        Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      End If;
      n_����ϼ� := 0;
      If ���ս���_In Is Not Null Then
        --�������ս���
        v_�������� := ���ս���_In || '||';
        n_����ϼ� := 0;
        While v_�������� Is Not Null Loop
          v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
          v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
          n_������ := To_Number(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
          If Nvl(n_������, 0) <> 0 Then
            Insert Into ����Ԥ����¼
              (ID, ��¼����, NO, ��¼״̬, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, ��������)
            Values
              (n_Ԥ��id, 4, ���ݺ�_In, 1, Decode(����id_In, 0, Null, ����id_In), '���ս���', v_���㷽ʽ, d_�Ǽ�ʱ��, v_����Ա���, v_����Ա����,
               n_������, n_����id, n_��id, n_����id, 4);
            Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          End If;
          n_����ϼ� := Nvl(n_����ϼ�, 0) + Nvl(n_������, 0);
          v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
        End Loop;
      End If;
    
      If Nvl(��Ԥ��_In, 0) <> 0 Then
        --������Ԥ��
        n_����ϼ� := n_����ϼ� + Nvl(��Ԥ��_In, 0);
        n_Ԥ����� := ��Ԥ��_In;
        For r_Deposit In c_Deposit(����id_In, v_��Ԥ������ids) Loop
          n_������ := Case
                      When r_Deposit.��� - n_Ԥ����� < 0 Then
                       r_Deposit.���
                      Else
                       n_Ԥ�����
                    End;
          If r_Deposit.Id <> 0 Then
            --��һ�γ�Ԥ��(���Ͻ���ID,���Ϊ0)
            Update ����Ԥ����¼ Set ��Ԥ�� = 0, ����id = n_����id, �������� = 4 Where ID = r_Deposit.Id;
          End If;
          --���ϴ�ʣ���
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����,
             ����Ա���, ��Ԥ��, ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������,
                   ��λ�ʺ�, d_�Ǽ�ʱ��, v_����Ա����, v_����Ա���, n_������, n_����id, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, n_����id, 4
            From ����Ԥ����¼
            Where NO = r_Deposit.No And ��¼״̬ = r_Deposit.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
        
          --���²���Ԥ�����
          Update �������
          Set Ԥ����� = Nvl(Ԥ�����, 0) - n_������
          Where ����id = r_Deposit.����id And ���� = 1 And ���� = Nvl(r_Deposit.Ԥ�����, 2)
          Returning Ԥ����� Into n_����ֵ;
          If Sql%RowCount = 0 Then
            Insert Into ������� (����id, Ԥ�����, ����, ����) Values (r_Deposit.����id, -1 * n_������, 1, 1);
            n_����ֵ := -1 * n_������;
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From �������
            Where ����id = r_Deposit.����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
          End If;
        
          --����Ƿ��Ѿ�������
          If r_Deposit.��� <= n_������ Then
            n_Ԥ����� := n_Ԥ����� - r_Deposit.���;
          Else
            n_Ԥ����� := 0;
          End If;
          If n_Ԥ����� = 0 Then
            Exit;
          End If;
        End Loop;
        If n_Ԥ����� > 0 Then
          v_Err_Msg := 'Ԥ������֧������֧�����,���ܼ���������';
          Raise Err_Item;
        End If;
      End If;
      --ʣ�����,��ָ�����㷽֧��
      n_������ := Nvl(n_ʵ�ս��ϼ�, 0) - Nvl(n_����ϼ�, 0);
      If Nvl(n_������, 0) < 0 Then
        v_Err_Msg := '�Һŵ���ؽ�������˵�ǰʵ����,���ܼ���������';
        Raise Err_Item;
      End If;
    
      If Nvl(n_������, 0) <> 0 Or (Nvl(n_������, 0) = 0 And Nvl(��Ԥ��_In, 0) = 0) Then
        If ���㷽ʽ_In Is Null Then
          v_Err_Msg := 'δ����ָ���Ľ��㷽ʽ,���ܼ���������';
          Raise Err_Item;
        End If;
      
        If Nvl(Ԥ��id_In, 0) <> 0 Then
          --�����Ԥ��ID_In��Ҫ��Ϊ�˽����������,���ҽ������վ���˸�ID,��Ҫ���µ�ID���и���,����������ת���ID
          Update ����Ԥ����¼ Set ID = n_Ԥ��id Where ID = Nvl(Ԥ��id_In, 0);
          n_Ԥ��id := Nvl(Ԥ��id_In, 0);
        End If;
        If Instr(���㷽ʽ_In, ',') = 0 Then
          --ֻ����һ�ֽ��㷽ʽ��
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��,
             ������λ, ��������, �������)
          Values
            (n_Ԥ��id, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), Nvl(v_���㷽ʽ, v_�ֽ�), Nvl(n_������, 0), �Ǽ�ʱ��_In,
             v_����Ա���, v_����Ա����, n_����id, '�Һ��շ�', n_��id, �����id_In, Null, ֧������_In, ������ˮ��_In, ����˵��_In, ������λ_In, 4, v_�������);
        Else
          v_��������     := ���㷽ʽ_In || '|'; --�Կո�ֿ���|��β,û�н�������
          n_Exists       := 0;
          v_���㷽ʽ��¼ := '';
          While v_�������� Is Not Null Loop
            v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '|') - 1);
            v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
          
            v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
            n_���ʽ�� := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
          
            v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
            v_������� := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
          
            v_��ǰ����   := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
            n_��������־ := To_Number(v_��ǰ����);
          
            If Instr('|' || v_���㷽ʽ��¼ || '|', '|' || Nvl(v_���㷽ʽ, v_�ֽ�) || '|') <> 0 Then
              v_Err_Msg := 'ʹ�����ظ��Ľ��㷽ʽ,����!';
              Raise Err_Item;
            Else
              v_���㷽ʽ��¼ := v_���㷽ʽ��¼ || '|' || Nvl(v_���㷽ʽ, v_�ֽ�);
            End If;
          
            If n_��������־ = 0 Then
              Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
              Insert Into ����Ԥ����¼
                (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��,
                 ����˵��, ������λ, ��������, �������)
              Values
                (n_Ԥ��id, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), Nvl(v_���㷽ʽ, v_�ֽ�), Nvl(n_���ʽ��, 0), �Ǽ�ʱ��_In,
                 v_����Ա���, v_����Ա����, n_����id, '�Һ��շ�', n_��id, Null, Null, Null, Null, Null, ������λ_In, 4, v_�������);
            Else
              If n_Exists = 1 Then
                v_Err_Msg := 'Ŀǰ�ҺŽ�֧��һ���������㷽ʽ,���ܼ���������';
                Raise Err_Item;
              End If;
              Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
              Insert Into ����Ԥ����¼
                (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��,
                 ����˵��, ������λ, ��������, �������)
              Values
                (n_Ԥ��id, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), Nvl(v_���㷽ʽ, v_�ֽ�), Nvl(n_���ʽ��, 0), �Ǽ�ʱ��_In,
                 v_����Ա���, v_����Ա����, n_����id, '�Һ��շ�', n_��id, �����id_In, Null, ֧������_In, ������ˮ��_In, ����˵��_In, ������λ_In, 4, v_�������);
              n_Exists := 1;
            End If;
          
            v_�������� := Substr(v_��������, Instr(v_��������, '|') + 1);
          End Loop;
        End If;
      End If;
    
      --������Ա�ɿ�����
    
      For v_�ɿ� In (Select ���㷽ʽ, Sum(Nvl(a.��Ԥ��, 0)) As ��Ԥ��
                   From ����Ԥ����¼ A
                   Where a.����id = n_����id And Mod(a.��¼����, 10) <> 1 And Nvl(����id, 0) = Nvl(����id_In, 0)
                   Group By ���㷽ʽ) Loop
      
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + Nvl(v_�ɿ�.��Ԥ��, 0)
        Where �տ�Ա = v_����Ա���� And ���� = 1 And ���㷽ʽ = v_�ɿ�.���㷽ʽ
        Returning ��� Into n_����ֵ;
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ����
            (�տ�Ա, ���㷽ʽ, ����, ���)
          Values
            (v_����Ա����, v_�ɿ�.���㷽ʽ, 1, Nvl(v_�ɿ�.��Ԥ��, 0));
          n_����ֵ := Nvl(v_�ɿ�.��Ԥ��, 0);
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From ��Ա�ɿ����
          Where �տ�Ա = v_����Ա���� And ���㷽ʽ = v_�ɿ�.���㷽ʽ And ���� = 1 And Nvl(���, 0) = 0;
        End If;
      End Loop;
    End If;
  
    --����Һż�¼
    If ��������_In = 2 Then
      Begin
        Select ID Into n_�Һ�id From ���˹Һż�¼ Where����¼״̬ = 0 And NO = ���ݺ�_In And ����id = ����id_In;
      Exception
        When Others Then
          Null;
      End;
    Else
      Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
    End If;
  
    Update ���˹Һż�¼
    Set ��¼���� = Decode(������ʽ_In, 2, 2, 1), ��¼״̬ = Decode(��������_In, 1, 0, 1), ����� = r_Pati.�����, ����Ա���� = v_����Ա����,
        ����Ա��� = v_����Ա���, ԤԼ = Decode(������ʽ_In, 1, 0, 1),
        ������ = Decode(��������_In, 1, Null, Decode(������ʽ_In, 2, Null, v_����Ա����)),
        ����ʱ�� = Case ��������_In
                  When 1 Then
                   Null
                  Else
                   Case ������ʽ_In
                     When 2 Then
                      Null
                     Else
                      d_�Ǽ�ʱ��
                   End
                End, ������ˮ�� = Nvl(������ˮ��_In, ������ˮ��), ����˵�� = Nvl(����˵��_In, ����˵��), ������λ = Nvl(������λ_In, ������λ),
        ԤԼ����Ա = Decode(������ʽ_In, 1, Nvl(ԤԼ����Ա, Null), Nvl(ԤԼ����Ա, v_����Ա����)),
        ԤԼ����Ա��� = Decode(������ʽ_In, 1, Nvl(ԤԼ����Ա���, Null), Nvl(ԤԼ����Ա���, v_����Ա���)), �����¼id = ��¼id_In
    Where ID = n_�Һ�id;
    If Sql%NotFound Then
      Begin
        Select ���� Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = r_Pati.���ʽ And Rownum < 2;
      Exception
        When Others Then
          v_���ʽ := Null;
      End;
      Insert Into ���˹Һż�¼
        (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ԤԼʱ��, ����Ա���,
         ����Ա����, ����, ����, ԤԼ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ҽ�Ƹ��ʽ, ԤԼ����Ա, ԤԼ����Ա���, �����¼id)
      Values
        (n_�Һ�id, ���ݺ�_In, Decode(������ʽ_In, 2, 2, 1), Decode(��������_In, 1, 0, 1), r_Pati.����id, r_Pati.�����, r_Pati.����,
         r_Pati.�Ա�, r_Pati.����, ����_In, 0, v_����, Null, n_����id, v_ҽ������, 0, Null, d_�Ǽ�ʱ��, ����ʱ��_In,
         Case When(Nvl(������ʽ_In, 0)) = 1 Then Null Else ����ʱ��_In End, v_����Ա���, v_����Ա����, 0, n_����, Decode(������ʽ_In, 1, 0, 1),
         Decode(������ʽ_In, 2, Null, v_����Ա����), Decode(������ʽ_In, 2, To_Date(Null), d_�Ǽ�ʱ��), ������ˮ��_In, ����˵��_In, ������λ_In,
         v_���ʽ, Decode(������ʽ_In, 1, Null, v_����Ա����), Decode(������ʽ_In, 1, Null, v_����Ա���), ��¼id_In);
    End If;
    --�������ݲ��ܲ�������
    If ��������_In <> 1 Then
      n_ԤԼ���ɶ��� := 0;
      If ������ʽ_In > 1 Then
        n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
      End If;
      --�Һź��շѵ�ԤԼ��ֱ�ӽ������(�շ�ԤԼȱ�ٽ��չ���,����ֱ�Ӻ͹Һ�һ��ֱ�ӽ������)
      If ������ʽ_In <> 2 Or n_ԤԼ���ɶ��� = 1 Then
        If Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113)) <> 0 Then
          --�Ŷӽк�ģʽ:-0-����������;1-��ҽ�������̨�Ŷ�;2-�ȷ���,��ҽ��վ
          If Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113)) = 0 Or n_ԤԼ���ɶ��� = 1 Then
            --��������
            --.����ִ�в��š� �ķ�ʽ���ɶ���
            v_�������� := n_����id;
            v_�ŶӺ��� := Zlgetnextqueue(n_����id, n_�Һ�id, ����_In || '|' || ����_In);
            --�Һ�id_In,����_In,����_In,ȱʡ����_In,��չ_In(������)
            v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 0);
            --�Һ�id_In,����_In,����_In,ȱʡ����_In,��չ_In(������)
            d_�Ŷ�ʱ�� := Zl_Get_Queuedate(n_�Һ�id, ����_In, ����_In, d_Date);
            --  ��������_In , ҵ������_In, ҵ��id_In,����id_In,�ŶӺ���_In,v_�Ŷӱ��,��������_In,����ID_IN, ����_In, ҽ������_In, �Ŷ�ʱ��_In
            Zl_�ŶӽкŶ���_Insert(v_��������, 0, n_�Һ�id, n_����id, v_�ŶӺ���, Null, r_Pati.����, r_Pati.����id, v_����, v_ҽ������, d_�Ŷ�ʱ��,
                             ԤԼ��ʽ_In, n_���÷�ʱ��, v_�Ŷ����);
          End If;
        End If;
      End If;
    
      If Nvl(������ʽ_In, 0) = 1 Then
        --����Ʊ��ʹ�����
        If Ʊ�ݺ�_In Is Not Null Then
          Select Ʊ�ݴ�ӡ����_Id.Nextval Into n_��ӡid From Dual;
          --����Ʊ��
          Insert Into Ʊ�ݴ�ӡ���� (ID, ��������, NO) Values (n_��ӡid, 4, ���ݺ�_In);
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
          Values
            (Ʊ��ʹ����ϸ_Id.Nextval, Decode(�շ�Ʊ��_In, 1, 1, 4), Ʊ�ݺ�_In, 1, 1, ����id_In, n_��ӡid, d_�Ǽ�ʱ��, v_����Ա����);
          --״̬�Ķ�
          Update Ʊ�����ü�¼
          Set ��ǰ���� = Ʊ�ݺ�_In, ʣ������ = Decode(Sign(ʣ������ - 1), -1, 0, ʣ������ - 1), ʹ��ʱ�� = Sysdate
          Where ID = Nvl(����id_In, 0);
        End If;
        --���˱��ξ���(�Է���ʱ��Ϊ׼)
        If Nvl(r_Pati.����id, 0) <> 0 Then
          Update ������Ϣ Set ����ʱ�� = ����ʱ��_In, ����״̬ = 1, �������� = v_���� Where ����id = r_Pati.����id;
        End If;
      End If;
    End If;
    --���˹ҺŻ���
    --��������ʱ�����ٶԻ��ܵ��ݽ���ͳ���� ����������ʱ�Ѿ������˻���
    If ��������_In <> 2 Then
      --������ʽ_IN:1-��ʾ�Һ�,2-��ʾԤԼ�ҺŲ��ۿ�,3-��ʾԤԼ�Һ�,�ۿ�
      --�Ƿ�ΪԤԼ����:0-��ԤԼ�Һ�; 1-ԤԼ�Һ�,2-ԤԼ����;3-�շ�ԤԼ
      n_ԤԼ := Case
                When Nvl(������ʽ_In, 0) = 1 Then
                 0
                When Nvl(������ʽ_In, 0) = 2 Then
                 1
                When Nvl(������ʽ_In, 0) = 3 Then
                 3
                Else
                 0
              End;
      Zl_���˹ҺŻ���_Update(v_ҽ������, n_ҽ��id, n_��Ŀid, n_����id, ����ʱ��_In, n_ԤԼ, ����_In, 0, ��¼id_In);
    End If;
    --��Ϣ����
    Begin
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
        Using 1, n_�Һ�id;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_001(n_�Һ�id, ���ݺ�_In);
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Err_Msg || '+[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End;

Begin
  n_�Һ��Ű�ģʽ := Nvl(zl_GetSysParameter('�Һ��Ű�ģʽ'), 0);
  If n_�Һ��Ű�ģʽ = 1 Then
    --������Ű�ģʽ
    Zl_���������Һ�_����_Insert(�����¼id_In, ������ʽ_In, ����id_In, ����_In, ����_In, ���ݺ�_In, Ʊ�ݺ�_In, ���㷽ʽ_In, ժҪ_In, ����ʱ��_In, �Ǽ�ʱ��_In,
                        ������λ_In, �ҺŽ��ϼ�_In, ����id_In, �շ�Ʊ��_In, ������ˮ��_In, ����˵��_In, ԤԼ��ʽ_In, Ԥ��id_In, �����id_In, �������״̬_In,
                        �Ƿ������豸_In, ����id_In, ��������_In, ���ս���_In, ��Ԥ��_In, ֧������_In, �˺�����_In, �ѱ�_In, ��Ԥ������ids_In, ������_In,
                        ��������_In, ������_In);
  Else
  
    v_��Ԥ������ids := Nvl(��Ԥ������ids_In, ����id_In);
    v_Temp          := zl_GetSysParameter(256);
    If v_Temp Is Null Or Substr(v_Temp, 1, 1) = '0' Then
      Null;
    Else
      Begin
        d_����ʱ�� := To_Date(Substr(v_Temp, 3), 'YYYY-MM-DD hh24:mi:ss');
        If ����ʱ��_In > d_����ʱ�� Then
          v_Err_Msg := '��ǰ�Һŵķ���ʱ��' || To_Char(����ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '�Ѿ������˳�����Ű�ģʽ,������ʹ�üƻ��Ű�ģʽ�Һ�!';
          Raise Err_Item;
        End If;
      Exception
        When Others Then
          Null;
      End;
    End If;
    If �ѱ�_In Is Null Then
      Select Zl_Custom_Getpatifeetype(1, ����id_In) Into v_�ѱ� From Dual;
    Else
      v_�ѱ� := �ѱ�_In;
    End If;
    If v_�ѱ� Is Null Then
      n_���ηѱ� := 1;
      Select ���� Into v_�ѱ� From �ѱ� Where ȱʡ��־ = 1 And Rownum < 2;
    End If;
    Update ������Ϣ Set �ѱ� = v_�ѱ� Where ����id = ����id_In;
  
    If ��������_In = 1 Then
      Select Zl_Age_Calc(����id_In) Into v_���� From Dual;
      If v_���� Is Not Null Then
        Update ������Ϣ Set ���� = v_���� Where ����id = ����id_In;
      End If;
    End If;
    --��ȡ��ǰ��������
    If ������_In Is Not Null Then
      v_������ := ������_In;
    Else
      Select Terminal Into v_������ From V$session Where Audsid = Userenv('sessionid');
    End If;
    n_ʵ�ս��ϼ� := 0;
    Select Count(*) + 1
    Into n_�Һ����
    From ���˹Һż�¼
    Where �ű� = ����_In And �Ǽ�ʱ�� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In + 1) - 1 / 24 / 60 / 60;
    --Begin
    --����ID,��������;��ԱID,��Ա���,��Ա����
    v_Temp := Zl_Identity(0);
    If Nvl(v_Temp, ' ') = ' ' Then
      v_Err_Msg := '��ǰ������Աδ���ö�Ӧ����Ա��ϵ,���ܼ�����';
      Raise Err_Item;
    End If;
  
    If �Ǽ�ʱ��_In Is Null Then
      d_�Ǽ�ʱ�� := Sysdate;
    Else
      d_�Ǽ�ʱ�� := �Ǽ�ʱ��_In;
    End If;
    If Trunc(Sysdate) > Trunc(����ʱ��_In) Then
      v_Err_Msg := '���ܹ���ǰ�ĺ�(' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || ')��';
      Raise Err_Item;
    End If;
    n_ͬ����Լһ���� := Nvl(zl_GetSysParameter('����ͬ����Լһ����', 1111), 0);
    n_����ԤԼ������ := Nvl(zl_GetSysParameter('����ԤԼ������', 1111), 0);
    n_��������id     := To_Number(Zl_����Ա(0, v_Temp));
    v_����Ա���     := Zl_����Ա(1, v_Temp);
    v_����Ա����     := Zl_����Ա(2, v_Temp);
    n_��id           := Zl_Get��id(v_����Ա����);
  
    If ������ʽ_In <> 1 Then
      --ԤԼ����Ƿ���Ӻ�����λ����
      --��������˺�����λ���� ��
      Begin
        Select ID
        Into n_�ƻ�id
        From �ҺŰ��żƻ�
        Where ���� = ����_In And ����ʱ��_In Between Nvl(��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
              Nvl(ʧЧʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Rownum < 2
        Order By ��Чʱ�� Desc;
      Exception
        When Others Then
          Select ID Into n_Tmp����id From �ҺŰ��� Where ���� = ����_In;
      End;
      If Nvl(n_�ƻ�id, 0) <> 0 Then
        Select Count(0)
        Into n_������λ����
        From ������λ�ƻ�����
        Where ������λ = ������λ_In And �ƻ�id = n_�ƻ�id And Rownum < 2;
      Else
        Select Count(0)
        Into n_������λ����
        From ������λ���ſ���
        Where ������λ = ������λ_In And ����id = n_Tmp����id And Rownum < 2;
      End If;
    End If;
  
    If ������ʽ_In <> 2 Then
      v_���� := Zl_����(����_In);
    End If;
    If ������ʽ_In <> 2 And ���㷽ʽ_In Is Not Null Then
      --�����㷽ʽ�Ƿ��걸
      Select Count(*) Into n_Count From ���㷽ʽ Where ���� = Nvl(���㷽ʽ_In, 'Lxh') And ���� In (2, 7, 8);
      If Nvl(�����id_In, 0) <> 0 And n_Count = 0 Then
        Select Count(1)
        Into n_Count
        From ҽ�ƿ����
        Where ID = Nvl(�����id_In, 0) And ���㷽ʽ = Nvl(���㷽ʽ_In, 'lxh');
      End If;
      If n_Count = 0 Then
        v_Err_Msg := '���㷽ʽ(' || ���㷽ʽ_In || ')δ����,���ڽ��㷽ʽ���������á�';
        Raise Err_Item;
      End If;
    End If;
  
    --��Ϊ�����а��ձ൥�ݺŹ���,�չҺ������ܳ���10000��,����Ҫ���ΨһԼ����
    Select Count(*) Into n_Count From ������ü�¼ Where ��¼���� = 4 And ��¼״̬ In (1, 3) And NO = ���ݺ�_In;
    If n_Count <> 0 Then
      v_Err_Msg := '�Һŵ��ݺ��ظ�,���ܱ��棡' || Chr(13) || '���ʹ���˰���˳����,���չҺ������ܳ���10000�˴Ρ�';
      Raise Err_Item;
    End If;
  
    Open c_Pati(����id_In);
    n_Count := 0;
    Begin
      Fetch c_Pati
        Into r_Pati;
    Exception
      When Others Then
        n_Count := -1;
    End;
    If n_Count = -1 Then
      v_Err_Msg := '����δ�ҵ������ܼ�����';
      Raise Err_Item;
    End If;
  
    Open c_����(����_In, ����ʱ��_In);
    Begin
      Fetch c_����
        Into r_����;
    Exception
      When Others Then
        n_Count := -1;
    End;
    If n_Count = -1 Then
      v_Err_Msg := '�úű�û����' || To_Char(����ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '�н��а��š�';
      Raise Err_Item;
    End If;
  
    Select Decode(To_Char(����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����',
                   '����')
    Into v_����
    From Dual;
    Begin
      If r_����.�ƻ�id Is Null Then
        Select 1 Into n_���÷�ʱ�� From �ҺŰ���ʱ�� Where ����id = r_����.Id And ���� = v_���� And Rownum < 2;
      Else
        Select 1 Into n_���÷�ʱ�� From �Һżƻ�ʱ�� Where �ƻ�id = r_����.�ƻ�id And ���� = v_���� And Rownum < 2;
      End If;
    Exception
      When Others Then
        n_���÷�ʱ�� := 0;
    End;
  
    --�Բ������ƽ��м��
    --����ԤԼ���ۿ�ʱ���м��
    If ������ʽ_In = 2 Then
      If Nvl(n_ͬ����Լһ����, 0) <> 0 Or Nvl(n_����ԤԼ������, 0) <> 0 Then
        n_��Լ���� := 0;
        For c_Chkitem In (Select Count(1) As ��Լ, a.ִ�в���id As ����id, Nvl(k.����, '') As ����
                          From ���˹Һż�¼ A, ������Ϣ B, ���ű� K
                          Where a.����id = b.����id And a.����id = ����id_In And a.ִ�в���id = k.Id(+) And a.��¼���� = 2 And ��¼״̬ = 1 And
                                a.ԤԼʱ�� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60
                          Group By a.ִ�в���id, k.����) Loop
          If Nvl(n_ͬ����Լһ����, 0) <> 0 And c_Chkitem.����id = r_����.����id Then
          
            v_Err_Msg := '�ò����Ѿ��ڿ���[' || c_Chkitem.���� || ']������ԤԼ,������ԤԼ��';
            Raise Err_Item;
          
            If Nvl(n_����ԤԼ������, 0) > 0 And c_Chkitem.����id <> r_����.����id Then
              n_��Լ���� := n_��Լ���� + 1;
            End If;
          End If;
        End Loop;
        If n_��Լ���� >= Nvl(n_����ԤԼ������, 0) And Nvl(n_����ԤԼ������, 0) > 0 Then
          v_Err_Msg := 'ͬһ���������ͬʱԤԼ[' || Nvl(n_����ԤԼ������, 0) || ']������,������ԤԼ��';
          Raise Err_Item;
        End If;
      End If;
    End If;
  
    d_Date         := Null;
    d_ʱ�ο�ʼʱ�� := Null;
  
    If Nvl(r_����.�޺���, 0) >= 0 Or r_����.�޺��� Is Null Then
    
      Select Nvl(Sum(Nvl(b.�ѹ���, 0)), 0), Nvl(Sum(Nvl(b.�����ѽ���, 0)), 0), Nvl(Sum(Nvl(b.��Լ��, 0)), 0)
      Into n_�ѹ���, n_�����ѽ���, n_��Լ��
      From �ҺŰ��� A, ���˹ҺŻ��� B
      Where a.����id = b.����id And a.��Ŀid = b.��Ŀid And a.���� = ����_In And b.���� Between Trunc(����ʱ��_In) And
            Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And (a.���� = b.���� Or b.���� Is Null) And Nvl(a.ҽ��id, 0) = Nvl(b.ҽ��id, 0) And
            Nvl(a.ҽ������, 'ҽ��') = Nvl(b.ҽ������, 'ҽ��');
    
      If n_���÷�ʱ�� = 1 Then
        If Nvl(r_����.��ſ���, 0) = 1 Then
          If Nvl(�Ƿ������豸_In, 0) = 0 Then
            If r_����.�ƻ�id Is Null Then
              Select Count(*), Max(��ʼʱ��)
              Into n_Count, d_ʱ�ο�ʼʱ��
              From �ҺŰ���ʱ��
              Where ����id = r_����.Id And ���� = v_���� And ��� = Nvl(����_In, 0);
            Else
              Select Count(*), Max(��ʼʱ��)
              Into n_Count, d_ʱ�ο�ʼʱ��
              From �Һżƻ�ʱ��
              Where �ƻ�id = r_����.�ƻ�id And ���� = v_���� And ��� = Nvl(����_In, 0);
            End If;
            v_Temp := '�Һ�';
            If ������ʽ_In > 1 Then
              v_Temp := 'ԤԼ�Һ�';
            End If;
          
            If n_Count = 0 Then
              v_Err_Msg := '�ű�Ϊ' || ����_In || '�ĹҺŰ����в��������Ϊ' || Nvl(����_In, 0) || '�İ���,������' || v_Temp || '��';
              Raise Err_Item;
            End If;
          End If;
          --�����,����ѡ��Һ�
          If Trunc(Sysdate) = Trunc(����ʱ��_In) Then
            --�ҵ���ĺ�
            v_Temp := To_Char(Sysdate, 'yyyy-mm-dd') || ' ';
            If r_����.�ƻ�id Is Null Then
              For v_ʱ�� In (Select To_Date(v_Temp || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                  To_Date(To_Char(Sysdate + Decode(Sign(��ʼʱ�� - ����ʱ��), 1, 1, 0), 'yyyy-mm-dd') || ' ' ||
                                           To_Char(����ʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, ��������, �Ƿ�ԤԼ
                           From �ҺŰ���ʱ��
                           Where ����id = r_����.Id And ���� = v_���� And ��� = Nvl(����_In, 0)) Loop
                If Sysdate > v_ʱ��.����ʱ�� Then
                  v_Err_Msg := '�ű�Ϊ' || ����_In || '������Ϊ' || Nvl(����_In, 0) || '�İ���,�Ѿ�����ʱ��,������' || v_Temp || '��';
                  Raise Err_Item;
                End If;
              End Loop;
            Else
              For v_ʱ�� In (Select To_Date(v_Temp || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                  To_Date(To_Char(Sysdate + Decode(Sign(��ʼʱ�� - ����ʱ��), 1, 1, 0), 'yyyy-mm-dd') || ' ' ||
                                           To_Char(����ʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, ��������, �Ƿ�ԤԼ
                           From �Һżƻ�ʱ��
                           Where �ƻ�id = r_����.�ƻ�id And ���� = v_���� And ��� = Nvl(����_In, 0)) Loop
                If Sysdate > v_ʱ��.����ʱ�� Then
                  v_Err_Msg := '�ű�Ϊ' || ����_In || '������Ϊ' || Nvl(����_In, 0) || '�İ���,�Ѿ�����ʱ��,������' || v_Temp || '��';
                  Raise Err_Item;
                End If;
              End Loop;
            End If;
          End If;
        Elsif ������ʽ_In > 1 Then
          --δ������ŵ�,��Ҫ���ԤԼ�����
          n_Count := 0;
          If r_����.�ƻ�id Is Null Then
            For v_ʱ�� In (Select ���, ��ʼʱ��, ����ʱ��, ��������, �Ƿ�ԤԼ
                         From �ҺŰ���ʱ��
                         Where ����id = r_����.Id And ���� = v_���� And
                               (('3000-01-10 ' || To_Char(����ʱ��_In, 'HH24:MI:SS') Between
                               Decode(Sign(��ʼʱ�� - ����ʱ�� - 1 / 24 / 60 / 60), 1,
                                        '3000-01-09 ' || To_Char(��ʼʱ��, 'HH24:MI:SS'),
                                        '3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS')) And
                               '3000-01-10 ' || To_Char(����ʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS')) Or
                               ('3000-01-10 ' || To_Char(����ʱ��_In, 'HH24:MI:SS') Between
                               '3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS') And
                               Decode(Sign(��ʼʱ�� - ����ʱ�� - 1 / 24 / 60 / 60), 1,
                                        '3000-01-11 ' || To_Char(����ʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS'),
                                        '3000-01-10 ' || To_Char(����ʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS'))))) Loop
              n_ԤԼʱ����� := v_ʱ��.���;
              d_ʱ�ο�ʼʱ�� := v_ʱ��.��ʼʱ��;
            
              Select Count(*), Max(���)
              Into n_Count, n_ԤԼ����
              From �Һ����״̬
              Where ���� = ����_In And ���� = ����ʱ��_In And ״̬ Not In (4, 5);
            
              If Nvl(n_Count, 0) > Nvl(v_ʱ��.��������, 0) And ��������_In <> 2 Then
                v_Err_Msg := '�ű�Ϊ' || ����_In || '�ĹҺŰ�������' || To_Char(v_ʱ��.��ʼʱ��, 'hh24:mi:ss') || '��' ||
                             To_Char(v_ʱ��.����ʱ��, 'hh24:mi:ss') || '���ֻ��ԤԼ' || Nvl(v_ʱ��.��������, 0) || '��,�����ٽ���ԤԼ�Һţ�';
                Raise Err_Item;
              End If;
              n_Count := 1;
            End Loop;
          Else
            For v_ʱ�� In (Select ���, ��ʼʱ��, ����ʱ��, ��������, �Ƿ�ԤԼ
                         From �Һżƻ�ʱ��
                         Where �ƻ�id = r_����.�ƻ�id And ���� = v_���� And
                               (('3000-01-10 ' || To_Char(����ʱ��_In, 'HH24:MI:SS') Between
                               Decode(Sign(��ʼʱ�� - ����ʱ�� - 1 / 24 / 60 / 60), 1,
                                        '3000-01-09 ' || To_Char(��ʼʱ��, 'HH24:MI:SS'),
                                        '3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS')) And
                               '3000-01-10 ' || To_Char(����ʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS')) Or
                               ('3000-01-10 ' || To_Char(����ʱ��_In, 'HH24:MI:SS') Between
                               '3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS') And
                               Decode(Sign(��ʼʱ�� - ����ʱ�� - 1 / 24 / 60 / 60), 1,
                                        '3000-01-11 ' || To_Char(����ʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS'),
                                        '3000-01-10 ' || To_Char(����ʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS'))))) Loop
              n_ԤԼʱ����� := v_ʱ��.���;
              d_ʱ�ο�ʼʱ�� := v_ʱ��.��ʼʱ��;
            
              Select Count(*), Max(���)
              Into n_Count, n_ԤԼ����
              From �Һ����״̬
              Where ���� = ����_In And ���� = ����ʱ��_In And ״̬ Not In (4, 5);
            
              If Nvl(n_Count, 0) > Nvl(v_ʱ��.��������, 0) And ��������_In <> 2 Then
                v_Err_Msg := '�ű�Ϊ' || ����_In || '�ĹҺŰ�������' || To_Char(v_ʱ��.��ʼʱ��, 'hh24:mi:ss') || '��' ||
                             To_Char(v_ʱ��.����ʱ��, 'hh24:mi:ss') || '���ֻ��ԤԼ' || Nvl(v_ʱ��.��������, 0) || '��,�����ٽ���ԤԼ�Һţ�';
                Raise Err_Item;
              End If;
              n_Count := 1;
            End Loop;
          End If;
        
          If n_Count = 0 Then
            v_Err_Msg := '�ű�Ϊ' || ����_In || '�ĹҺŰ�����û����صİ��żƻ�(' || To_Char(����ʱ��_In, 'YYYY-mm-dd HH24:MI:SS') ||
                         '),���ܽ���ԤԼ�Һţ�';
            Raise Err_Item;
          End If;
        End If;
      End If;
    End If;
  
    If ������ʽ_In = 1 And ��������_In <> 2 Then
      --�ҺŹ���:
      --  �ѹ������ܴ����޺���
      If n_�ѹ��� >= Nvl(r_����.�޺���, 0) And r_����.�޺��� Is Not Null Then
        v_Err_Msg := '�úű�����Ѵﵽ�޺��� ' || Nvl(r_����.�޺���, 0) || '�����ٹҺţ�';
        Raise Err_Item;
      End If;
    End If;
  
    If ������ʽ_In > 1 Then
      --ԤԼ����ؼ��
      --����:
      --   1.����Լ���ܳ�����Լ��
      --   2.����Ƿ�����ʱ�ε�
      If n_��Լ�� >= Nvl(r_����.��Լ��, 0) And Nvl(r_����.��Լ��, 0) <> 0 And r_����.��Լ�� Is Not Null And ��������_In <> 2 Then
        v_Err_Msg := '�úű��Ѵﵽ��Լ�� ' || Nvl(r_����.��Լ��, 0) || '������ԤԼ�Һţ�';
        Raise Err_Item;
      End If;
    End If;
    If n_������λ���� > 0 And ������ʽ_In <> 1 And ������λ_In Is Not Null Then
    
      If Nvl(r_����.��ſ���, 0) = 1 And Nvl(����_In, 0) = 0 Then
        v_Err_Msg := '��ǰ����ʹ������ſ���,��ȷ������ҪԤԼ�����,���ܼ�����';
        Raise Err_Item;
      End If; --Nvl(r_����.��ſ���, 0) =0
    
      n_��� := Case
                When Nvl(r_����.��ſ���, 0) = 1 Or n_���÷�ʱ�� = 1 And ������ʽ_In > 1 Then
                 Nvl(����_In, 0)
                Else
                 0
              End;
    
      --������λ������ģʽ
      Begin
        If Nvl(n_�ƻ�id, 0) <> 0 Then
          Select 0
          Into n_���
          From ������λ�ƻ�����
          Where ������λ = ������λ_In And �ƻ�id = n_�ƻ�id And
                ������Ŀ = Decode(To_Char(����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                              '7', '����', Null) And ���� <> 0 And ��� = 0 And Rownum < 2;
        Else
          Select 0
          Into n_���
          From ������λ���ſ���
          Where ������λ = ������λ_In And ����id = n_Tmp����id And
                ������Ŀ = Decode(To_Char(����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                              '7', '����', Null) And ���� <> 0 And ��� = 0 And Rownum < 2;
        End If;
        n_������λ������ģʽ := 1;
      Exception
        When Others Then
          n_������λ������ģʽ := 0;
      End;
      --������ż��
      For c_������λ In (Select c.���, ����
                     From �ҺŰ��� A, ������λ���ſ��� C
                     Where a.���� = ����_In And Decode(To_Char(����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                   '����', '6', '����', '7', '����', Null) = c.������Ŀ(+) And a.Id = c.����id And
                           c.������λ = ������λ_In And c.��� = n_��� And Not Exists
                      (Select 1
                            From �ҺŰ��żƻ� D
                            Where d.����id = a.Id And d.���ʱ�� Is Not Null And
                                  ����ʱ��_In Between Nvl(d.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                                  Nvl(d.ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')))
                     Union All
                     Select c.���, ����
                     From �ҺŰ��żƻ� A, �ҺŰ��� D, ������λ�ƻ����� C,
                          (Select Max(a.��Чʱ��) As ��Ч, ����id
                            From �ҺŰ��żƻ� A, �ҺŰ��� B
                            Where a.����id = b.Id And a.���ʱ�� Is Not Null And
                                  ����ʱ��_In Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                                  Nvl(a.ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) And b.���� = ����_In
                            Group By ����id) E
                     Where a.����id = d.Id And a.���ʱ�� Is Not Null And d.���� = ����_In And a.����id = e.����id And
                           Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) =
                           Nvl(e.��Ч, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                           Decode(To_Char(����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                                  '7', '����', Null) = c.������Ŀ(+) And a.Id = c.�ƻ�id And c.������λ = ������λ_In And c.��� = n_��� And
                           ����ʱ��_In Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                           Nvl(a.ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd'))) Loop
      
        If Nvl(r_����.��ſ���, 0) = 1 And c_������λ.��� = n_��� And n_������λ������ģʽ = 0 Then
          n_�Ƿ񿪷� := 1;
          Exit;
        Elsif (Nvl(r_����.��ſ���, 0) = 0 And c_������λ.��� = n_���) Or n_������λ������ģʽ = 1 Then
          Begin
            Select Nvl(��Լ��, 0)
            Into n_ԤԼ����
            From ������λ�ҺŻ���
            Where ������λ = ������λ_In And ���� = Trunc(����ʱ��_In) And ���� = ����_In;
          Exception
            When Others Then
              n_ԤԼ���� := 0;
          End;
          If c_������λ.���� <= n_ԤԼ���� And Nvl(c_������λ.����, 0) > 0 And ��������_In <> 2 Then
            v_Err_Msg := '�úű��Ѵﵽ��Լ�� ' || Nvl(c_������λ.����, 0) || '������ԤԼ�Һţ�';
            Raise Err_Item;
          End If;
          n_�Ƿ񿪷� := 1;
          Exit;
        End If;
      
      End Loop;
    
      If Nvl(n_�Ƿ񿪷�, 0) = 0 Then
        v_Err_Msg := '��ǰ���(' || Nvl(����_In, 0) || 'δ����,���ܼ�����';
        Raise Err_Item;
      End If;
    End If;
  
    --����޺�������Լ��
    n_�к�         := 1;
    n_ԭ��Ŀid     := 0;
    n_ԭ������Ŀid := 0;
    n_ʵ�ս��ϼ� := 0;
    If ��������_In <> 1 Then
      If ������ʽ_In <> 2 Then
        If Nvl(����id_In, 0) = 0 Then
          --����Ӧ�ó�����
          Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
        Else
          n_����id := ����id_In;
        End If;
      Else
        n_����id := Null;
      End If;
    End If;
  
    If Nvl(������_In, 0) = 1 Then
      Begin
        Select �շ�ϸĿid Into n_������id From �շ��ض���Ŀ Where �ض���Ŀ = '������';
        v_�շ���Ŀids := r_����.��Ŀid || ',' || n_������id;
      Exception
        When Others Then
          v_Err_Msg := '����ȷ��������,�Һ�ʧ��!';
          Raise Err_Item;
      End;
    Else
      v_�շ���Ŀids := r_����.��Ŀid;
    End If;
  
    For c_Item In (Select 1 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, 1 As ����,
                          c.Id As ������Ŀid, c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, Null As ��������
                   From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C
                   Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = r_����.��Ŀid And Sysdate Between b.ִ������ And
                         Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))
                   Union All
                   Select 2 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, 1 As ����,
                          c.Id As ������Ŀid, c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, Null As ��������
                   From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C
                   Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = n_������id And Sysdate Between b.ִ������ And
                         Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))
                   Union All
                   Select 3 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, d.�������� As ����,
                          c.Id As ������Ŀid, c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, 1 As ��������
                   From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, �շѴ�����Ŀ D
                   Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = d.����id And
                         d.����id In (Select Column_Value From Table(f_Str2list(v_�շ���Ŀids))) And Sysdate Between b.ִ������ And
                         Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))
                   Order By ����, ��Ŀ����, �������) Loop
      n_�۸񸸺� := Null;
      If n_ԭ��Ŀid = c_Item.��Ŀid Then
        If n_ԭ������Ŀid <> c_Item.������Ŀid Then
          n_�۸񸸺� := n_�к�;
        End If;
        n_ԭ������Ŀid := c_Item.������Ŀid;
      End If;
      n_ԭ��Ŀid := c_Item.��Ŀid;
      n_Ӧ�ս�� := Round(c_Item.���� * c_Item.����, 5);
      n_ʵ�ս�� := n_Ӧ�ս��;
      If Nvl(c_Item.���ηѱ�, 0) <> 1 And n_���ηѱ� = 0 Then
        --����:
        v_Temp     := Zl_Actualmoney(r_Pati.�ѱ�, c_Item.��Ŀid, c_Item.������Ŀid, n_Ӧ�ս��);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ':') + 1);
        n_ʵ�ս�� := Zl_To_Number(v_Temp);
      End If;
      n_ʵ�ս��ϼ� := Nvl(n_ʵ�ս��ϼ�, 0) + n_ʵ�ս��;
    
      --�������ݲ���������
      If ��������_In <> 1 Then
        --�������˹Һŷ���(���ܵ����ǻ������������)
        Select ���˷��ü�¼_Id.Nextval Into n_����id From Dual;
        --:������ʽ_IN:1-��ʾ�Һ�,2-��ʾԤԼ�ҺŲ��ۿ�,3-��ʾԤԼ�Һ�,�ۿ�
        Insert Into ������ü�¼
          (ID, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, NO, ʵ��Ʊ��, �����־, �Ӱ��־, ���ӱ�־, ��ҩ����, ����id, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, ���˿���id,
           �շ����, ���㵥λ, �շ�ϸĿid, ������Ŀid, �վݷ�Ŀ, ����, ����, ��׼����, Ӧ�ս��, ʵ�ս��, ���ʽ��, ����id, ���ʷ���, ��������id, ������, ������, ִ�в���id, ִ����,
           ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ���մ���id, ������Ŀ��, ���ձ���, ͳ����, ժҪ, ����, �ɿ���id)
        Values
          (n_����id, 4, Decode(������ʽ_In, 2, 0, 1), n_�к�, n_�۸񸸺�, c_Item.��������, ���ݺ�_In, Ʊ�ݺ�_In, 1, Null, Null,
           Decode(������ʽ_In, 2, To_Char(����_In), v_����), r_Pati.����id, r_Pati.�����, r_Pati.���ʽ, r_Pati.����, r_Pati.�Ա�,
           r_Pati.����, r_Pati.�ѱ�, r_����.����id, c_Item.���, ����_In, c_Item.��Ŀid, c_Item.������Ŀid, c_Item.�վݷ�Ŀ, 1, c_Item.����,
           c_Item.����, n_Ӧ�ս��, n_ʵ�ս��, Decode(������ʽ_In, 2, Null, n_ʵ�ս��), n_����id, 0, n_��������id, v_����Ա����,
           Decode(������ʽ_In, 2, v_����Ա����, Null), r_����.����id, r_����.ҽ������, v_����Ա���, v_����Ա����, ����ʱ��_In, d_�Ǽ�ʱ��, Null, 0, Null,
           Null, ժҪ_In, ԤԼ��ʽ_In, Decode(������ʽ_In, 2, Null, n_��id));
      End If;
      n_�к� := n_�к� + 1;
    
    End Loop;
  
    If Round(Nvl(�ҺŽ��ϼ�_In, 0), 5) <> Round(Nvl(n_ʵ�ս��ϼ�, 0), 5) Then
      v_Err_Msg := '���ιҺŽ���ȷ,��������ΪҽԺ�����˼۸�,�����»�ȡ�Һ��շ���Ŀ�ļ۸�,���ܼ�����';
      Raise Err_Item;
    End If;
  
    If n_���÷�ʱ�� = 1 Then
      d_Date := To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(d_ʱ�ο�ʼʱ��, 'hh24:mi:ss'),
                        'yyyy-mm-dd hh24:mi:ss');
    Else
      d_Date := Trunc(����ʱ��_In);
    End If;
  
    --���¹Һ����״̬
    If ��������_In <> 2 Then
      n_���� := ����_In;
    End If;
    Begin
      Select 1
      Into n_Count
      From �Һ����״̬
      Where Trunc(����) = Trunc(����ʱ��_In) And ���� = ����_In And ��� = n_���� And ״̬ <> 5;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count = 1 Then
      If n_���÷�ʱ�� = 0 And Nvl(r_����.��ſ���, 0) = 1 Then
        n_���� := Null;
      End If;
      If n_���÷�ʱ�� = 1 And Nvl(r_����.��ſ���, 0) = 1 Then
        v_Err_Msg := '��ǰ����ѱ�ʹ�ã�������ѡ��һ����ţ�';
        Raise Err_Item;
      End If;
    End If;
    n_Count := 0;
    If n_���÷�ʱ�� = 0 And Nvl(r_����.��ſ���, 0) = 1 And n_���� Is Null And ��������_In <> 2 Then
      If �˺�����_In = 1 Then
        Select Nvl(Max(���), 0) + 1
        Into n_����
        From �Һ����״̬
        Where ���� = Trunc(����ʱ��_In) And ���� = r_����.���� And ״̬ Not In (4, 5);
      Else
        Select Nvl(Max(���), 0) + 1
        Into n_����
        From �Һ����״̬
        Where ���� = Trunc(����ʱ��_In) And ���� = r_����.���� And ״̬ <> 5;
      End If;
    End If;
    If n_���÷�ʱ�� = 1 And ��������_In <> 2 Then
    
      If ������ʽ_In > 1 And Nvl(r_����.��ſ���, 0) = 0 Then
        --����:ԤԼʱ�����||ԤԼ��
        If Nvl(n_ԤԼ����, 0) = 0 Then
          v_Temp := Nvl(r_����.��Լ��, 0);
          v_Temp := LTrim(RTrim(v_Temp));
          v_Temp := LPad(Nvl(n_ԤԼ����, 0) + 1, Length(v_Temp), '0');
          v_Temp := n_ԤԼʱ����� || v_Temp;
          n_���� := To_Number(v_Temp);
        Else
          n_���� := n_ԤԼ���� + 1;
        End If;
      End If;
    End If;
  
    If Nvl(r_����.��ſ���, 0) = 1 Or (������ʽ_In > 1 And n_���÷�ʱ�� = 1) Or �������״̬_In = 1 Then
      --������ŵĴ���
      Begin
        Select ����Ա����, ������
        Into v_��Ų���Ա, v_��Ż�����
        From �Һ����״̬
        Where ״̬ = 5 And ���� = ����_In And Trunc(����) = Trunc(d_Date) And ��� = n_����;
        n_������� := 1;
      Exception
        When Others Then
          v_��Ų���Ա := Null;
          v_��Ż����� := Null;
          n_�������   := 0;
      End;
      If n_������� = 0 Then
        Update �Һ����״̬
        Set ״̬ = Decode(������ʽ_In, 2, 2, 1), ԤԼ = Decode(������ʽ_In, 1, 0, 1), �Ǽ�ʱ�� = Sysdate
        Where ���� = ����_In And ���� = d_Date And ��� = n_���� And ����Ա���� = v_����Ա����;
        If Sql%RowCount = 0 Then
          Begin
            Insert Into �Һ����״̬
              (����, ����, ���, ״̬, ����Ա����, ԤԼ, �Ǽ�ʱ��)
            Values
              (����_In, d_Date, n_����, Decode(������ʽ_In, 2, 2, 1), v_����Ա����, Decode(������ʽ_In, 1, 0, 1), Sysdate);
          
            If n_������λ���� > 0 And ������ʽ_In > 1 And Nvl(n_�Ƿ񿪷�, 0) = 1 Then
              Update ������λ�ҺŻ���
              Set ��Լ�� = ��Լ�� + Decode(������ʽ_In, 2, 1, 0), �ѽ��� = �ѽ��� + Decode(������ʽ_In, 3, 1, 0)
              Where ���� = ����_In And ���� = d_Date And ��� = n_���� And ������λ = ������λ_In;
              If Sql%NotFound Then
                Insert Into ������λ�ҺŻ���
                  (����, ����, ���, ������λ, ��Լ��, �ѽ���)
                Values
                  (����_In, d_Date, n_����, ������λ_In, Decode(������ʽ_In, 1, 0, 1), Decode(������ʽ_In, 3, 1, 0));
              End If;
            End If;
          Exception
            When Others Then
              v_Err_Msg := '���' || n_���� || '�ѱ�ʹ��,������ѡ��һ�����.';
              Raise Err_Item;
          End;
        End If;
      Else
        If v_����Ա���� <> v_��Ų���Ա Or v_������ <> v_��Ż����� Then
          v_Err_Msg := '���' || n_���� || '�ѱ�������' || v_������ || '����,������ѡ��һ�����.';
          Raise Err_Item;
        Else
          Update �Һ����״̬
          Set ״̬ = Decode(������ʽ_In, 2, 2, 1), ԤԼ = Decode(������ʽ_In, 1, 0, 1), �Ǽ�ʱ�� = Sysdate
          Where ���� = ����_In And Trunc(����) = Trunc(d_Date) And ��� = n_���� And ״̬ = 5 And ����Ա���� = v_����Ա���� And ������ = v_������;
        End If;
      End If;
    End If;
  
    --�������ݲ������κ� ����
    If ������ʽ_In <> 2 And ��������_In <> 1 Then
      --�Һ�,ԤԼ�Һ��Ѿ��ۿ��
      n_Ԥ��id := Ԥ��id_In;
      If Nvl(n_Ԥ��id, 0) = 0 Then
        Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      End If;
      n_����ϼ� := 0;
      If ���ս���_In Is Not Null Then
        --�������ս���
        v_�������� := ���ս���_In || '||';
        n_����ϼ� := 0;
        While v_�������� Is Not Null Loop
          v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
          v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
          n_������ := To_Number(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
          If Nvl(n_������, 0) <> 0 Then
            Insert Into ����Ԥ����¼
              (ID, ��¼����, NO, ��¼״̬, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, ��������)
            Values
              (n_Ԥ��id, 4, ���ݺ�_In, 1, Decode(����id_In, 0, Null, ����id_In), '���ս���', v_���㷽ʽ, d_�Ǽ�ʱ��, v_����Ա���, v_����Ա����,
               n_������, n_����id, n_��id, n_����id, 4);
            Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          End If;
          n_����ϼ� := Nvl(n_����ϼ�, 0) + Nvl(n_������, 0);
          v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
        End Loop;
      End If;
    
      If Nvl(��Ԥ��_In, 0) <> 0 Then
        --������Ԥ��
        n_����ϼ� := n_����ϼ� + Nvl(��Ԥ��_In, 0);
        n_Ԥ����� := ��Ԥ��_In;
        For r_Deposit In c_Deposit(����id_In, v_��Ԥ������ids) Loop
          n_������ := Case
                      When r_Deposit.��� - n_Ԥ����� < 0 Then
                       r_Deposit.���
                      Else
                       n_Ԥ�����
                    End;
          If r_Deposit.Id <> 0 Then
            --��һ�γ�Ԥ��(���Ͻ���ID,���Ϊ0)
            Update ����Ԥ����¼ Set ��Ԥ�� = 0, ����id = n_����id, �������� = 4 Where ID = r_Deposit.Id;
          End If;
          --���ϴ�ʣ���
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����,
             ����Ա���, ��Ԥ��, ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������,
                   ��λ�ʺ�, d_�Ǽ�ʱ��, v_����Ա����, v_����Ա���, n_������, n_����id, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, n_����id, 4
            From ����Ԥ����¼
            Where NO = r_Deposit.No And ��¼״̬ = r_Deposit.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
        
          --���²���Ԥ�����
          Update �������
          Set Ԥ����� = Nvl(Ԥ�����, 0) - n_������
          Where ����id = r_Deposit.����id And ���� = 1 And ���� = Nvl(r_Deposit.Ԥ�����, 2)
          Returning Ԥ����� Into n_����ֵ;
          If Sql%RowCount = 0 Then
            Insert Into ������� (����id, Ԥ�����, ����, ����) Values (r_Deposit.����id, -1 * n_������, 1, 1);
            n_����ֵ := -1 * n_������;
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From �������
            Where ����id = r_Deposit.����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
          End If;
        
          --����Ƿ��Ѿ�������
          If r_Deposit.��� <= n_������ Then
            n_Ԥ����� := n_Ԥ����� - r_Deposit.���;
          Else
            n_Ԥ����� := 0;
          End If;
          If n_Ԥ����� = 0 Then
            Exit;
          End If;
        End Loop;
        If n_Ԥ����� > 0 Then
          v_Err_Msg := 'Ԥ������֧������֧�����,���ܼ���������';
          Raise Err_Item;
        End If;
      End If;
      --ʣ�����,��ָ�����㷽֧��
      n_������ := Nvl(n_ʵ�ս��ϼ�, 0) - Nvl(n_����ϼ�, 0);
      If Nvl(n_������, 0) < 0 Then
        v_Err_Msg := '�Һŵ���ؽ�������˵�ǰʵ����,���ܼ���������';
        Raise Err_Item;
      End If;
      If Nvl(n_������, 0) <> 0 Or (Nvl(n_������, 0) = 0 And Nvl(��Ԥ��_In, 0) = 0) Then
        If ���㷽ʽ_In Is Null Then
          v_Err_Msg := 'δ����ָ���Ľ��㷽ʽ,���ܼ���������';
          Raise Err_Item;
        End If;
      
        If Nvl(Ԥ��id_In, 0) <> 0 Then
          --�����Ԥ��ID_In��Ҫ��Ϊ�˽����������,���ҽ������վ���˸�ID,��Ҫ���µ�ID���и���,����������ת���ID
          Update ����Ԥ����¼ Set ID = n_Ԥ��id Where ID = Nvl(Ԥ��id_In, 0);
          n_Ԥ��id := Nvl(Ԥ��id_In, 0);
        End If;
      
        Insert Into ����Ԥ����¼
          (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, ������ˮ��, ����˵��, �������, ������λ, �����id, ����,
           ��������)
        Values
          (n_Ԥ��id, 4, 1, ���ݺ�_In, r_Pati.����id, ���㷽ʽ_In, Nvl(n_������, 0), d_�Ǽ�ʱ��, v_����Ա���, v_����Ա����, n_����id,
           ������λ_In || '�ɿ�', n_��id, ������ˮ��_In, ����˵��_In, n_����id, ������λ_In, �����id_In, ֧������_In, 4);
      End If;
    
      --������Ա�ɿ�����
    
      For v_�ɿ� In (Select ���㷽ʽ, Sum(Nvl(a.��Ԥ��, 0)) As ��Ԥ��
                   From ����Ԥ����¼ A
                   Where a.����id = n_����id And Mod(a.��¼����, 10) <> 1 And Nvl(����id, 0) = Nvl(����id_In, 0)
                   Group By ���㷽ʽ) Loop
      
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + Nvl(v_�ɿ�.��Ԥ��, 0)
        Where �տ�Ա = v_����Ա���� And ���� = 1 And ���㷽ʽ = v_�ɿ�.���㷽ʽ
        Returning ��� Into n_����ֵ;
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ����
            (�տ�Ա, ���㷽ʽ, ����, ���)
          Values
            (v_����Ա����, v_�ɿ�.���㷽ʽ, 1, Nvl(v_�ɿ�.��Ԥ��, 0));
          n_����ֵ := Nvl(v_�ɿ�.��Ԥ��, 0);
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From ��Ա�ɿ����
          Where �տ�Ա = v_����Ա���� And ���㷽ʽ = ���㷽ʽ_In And ���� = 1 And Nvl(���, 0) = 0;
        End If;
      
      End Loop;
    
    End If;
  
    --����Һż�¼
    If ��������_In = 2 Then
      Begin
        Select ID Into n_�Һ�id From ���˹Һż�¼ Where����¼״̬ = 0 And NO = ���ݺ�_In And ����id = ����id_In;
      Exception
        When Others Then
          Null;
      End;
    Else
      Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
    End If;
  
    Update ���˹Һż�¼
    Set ��¼���� = Decode(������ʽ_In, 2, 2, 1), ��¼״̬ = Decode(��������_In, 1, 0, 1), ����� = r_Pati.�����, ����Ա���� = v_����Ա����,
        ����Ա��� = v_����Ա���, ԤԼ = Decode(������ʽ_In, 1, 0, 1),
        ������ = Decode(��������_In, 1, Null, Decode(������ʽ_In, 2, Null, v_����Ա����)),
        ����ʱ�� = Case ��������_In
                  When 1 Then
                   Null
                  Else
                   Case ������ʽ_In
                     When 2 Then
                      Null
                     Else
                      d_�Ǽ�ʱ��
                   End
                End, ������ˮ�� = Nvl(������ˮ��_In, ������ˮ��), ����˵�� = Nvl(����˵��_In, ����˵��), ������λ = Nvl(������λ_In, ������λ),
        ԤԼ����Ա = Decode(������ʽ_In, 1, Nvl(ԤԼ����Ա, Null), Nvl(ԤԼ����Ա, v_����Ա����)),
        ԤԼ����Ա��� = Decode(������ʽ_In, 1, Nvl(ԤԼ����Ա���, Null), Nvl(ԤԼ����Ա���, v_����Ա���))
    Where ID = n_�Һ�id;
    If Sql%NotFound Then
      Begin
        Select ���� Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = r_Pati.���ʽ And Rownum < 2;
      Exception
        When Others Then
          v_���ʽ := Null;
      End;
      Insert Into ���˹Һż�¼
        (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ԤԼʱ��, ����Ա���,
         ����Ա����, ����, ����, ԤԼ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ҽ�Ƹ��ʽ, ԤԼ����Ա, ԤԼ����Ա���)
      Values
        (n_�Һ�id, ���ݺ�_In, Decode(������ʽ_In, 2, 2, 1), Decode(��������_In, 1, 0, 1), r_Pati.����id, r_Pati.�����, r_Pati.����,
         r_Pati.�Ա�, r_Pati.����, ����_In, 0, v_����, Null, r_����.����id, r_����.ҽ������, 0, Null, d_�Ǽ�ʱ��, ����ʱ��_In,
         Case When(Nvl(������ʽ_In, 0)) = 1 Then Null Else ����ʱ��_In End, v_����Ա���, v_����Ա����, 0, n_����, Decode(������ʽ_In, 1, 0, 1),
         Decode(������ʽ_In, 2, Null, v_����Ա����), Decode(������ʽ_In, 2, To_Date(Null), d_�Ǽ�ʱ��), ������ˮ��_In, ����˵��_In, ������λ_In,
         v_���ʽ, Decode(������ʽ_In, 1, Null, v_����Ա����), Decode(������ʽ_In, 1, Null, v_����Ա���));
    End If;
    --�������ݲ��ܲ�������
    If ��������_In <> 1 Then
      n_ԤԼ���ɶ��� := 0;
      If ������ʽ_In > 1 Then
        n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
      End If;
      --�Һź��շѵ�ԤԼ��ֱ�ӽ������(�շ�ԤԼȱ�ٽ��չ���,����ֱ�Ӻ͹Һ�һ��ֱ�ӽ������)
      If ������ʽ_In <> 2 Or n_ԤԼ���ɶ��� = 1 Then
        If Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113)) <> 0 Then
          --�Ŷӽк�ģʽ:-0-����������;1-��ҽ�������̨�Ŷ�;2-�ȷ���,��ҽ��վ
          If Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113)) = 0 Or n_ԤԼ���ɶ��� = 1 Then
            --��������
            --.����ִ�в��š� �ķ�ʽ���ɶ���
            v_�������� := r_����.����id;
            v_�ŶӺ��� := Zlgetnextqueue(r_����.����id, n_�Һ�id, ����_In || '|' || ����_In);
            --�Һ�id_In,����_In,����_In,ȱʡ����_In,��չ_In(������)
            v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 0);
            --�Һ�id_In,����_In,����_In,ȱʡ����_In,��չ_In(������)
            d_�Ŷ�ʱ�� := Zl_Get_Queuedate(n_�Һ�id, ����_In, ����_In, d_Date);
            --  ��������_In , ҵ������_In, ҵ��id_In,����id_In,�ŶӺ���_In,v_�Ŷӱ��,��������_In,����ID_IN, ����_In, ҽ������_In, �Ŷ�ʱ��_In
            Zl_�ŶӽкŶ���_Insert(v_��������, 0, n_�Һ�id, r_����.����id, v_�ŶӺ���, Null, r_Pati.����, r_Pati.����id, v_����, r_����.ҽ������,
                             d_�Ŷ�ʱ��, ԤԼ��ʽ_In, n_���÷�ʱ��, v_�Ŷ����);
          End If;
        End If;
      End If;
    
      If Nvl(������ʽ_In, 0) = 1 Then
        --����Ʊ��ʹ�����
        If Ʊ�ݺ�_In Is Not Null Then
          Select Ʊ�ݴ�ӡ����_Id.Nextval Into n_��ӡid From Dual;
          --����Ʊ��
          Insert Into Ʊ�ݴ�ӡ���� (ID, ��������, NO) Values (n_��ӡid, 4, ���ݺ�_In);
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
          Values
            (Ʊ��ʹ����ϸ_Id.Nextval, Decode(�շ�Ʊ��_In, 1, 1, 4), Ʊ�ݺ�_In, 1, 1, ����id_In, n_��ӡid, d_�Ǽ�ʱ��, v_����Ա����);
          --״̬�Ķ�
          Update Ʊ�����ü�¼
          Set ��ǰ���� = Ʊ�ݺ�_In, ʣ������ = Decode(Sign(ʣ������ - 1), -1, 0, ʣ������ - 1), ʹ��ʱ�� = Sysdate
          Where ID = Nvl(����id_In, 0);
        End If;
        --���˱��ξ���(�Է���ʱ��Ϊ׼)
        If Nvl(r_Pati.����id, 0) <> 0 Then
          Update ������Ϣ Set ����ʱ�� = ����ʱ��_In, ����״̬ = 1, �������� = v_���� Where ����id = r_Pati.����id;
        End If;
      End If;
    End If;
    --���˹ҺŻ���
    --��������ʱ�����ٶԻ��ܵ��ݽ���ͳ���� ����������ʱ�Ѿ������˻���
    If ��������_In <> 2 Then
      --������ʽ_IN:1-��ʾ�Һ�,2-��ʾԤԼ�ҺŲ��ۿ�,3-��ʾԤԼ�Һ�,�ۿ�
      --�Ƿ�ΪԤԼ����:0-��ԤԼ�Һ�; 1-ԤԼ�Һ�,2-ԤԼ����;3-�շ�ԤԼ
      n_ԤԼ := Case
                When Nvl(������ʽ_In, 0) = 1 Then
                 0
                When Nvl(������ʽ_In, 0) = 2 Then
                 1
                When Nvl(������ʽ_In, 0) = 3 Then
                 3
                Else
                 0
              End;
      Zl_���˹ҺŻ���_Update(r_����.ҽ������, r_����.ҽ��id, r_����.��Ŀid, r_����.����id, ����ʱ��_In, n_ԤԼ, ����_In);
    End If;
    --��Ϣ����
    Begin
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
        Using 1, n_�Һ�id;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_001(n_�Һ�id, ���ݺ�_In);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Err_Msg || '+[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���������Һ�_Insert;
/


Create Or Replace Procedure Zl_Third_Getdeptlist
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --����:��ȡ�ɹҺſ���

  --���:Xml_In:
  --<IN>
  --  <CXTS>14</CXTS>        //��ѯ����
  --  <HZDW>֧����</HZDW>    //������λ
  --  <ZD></ZD>              //վ��
  --</IN>

  --����:Xml_Out
  --<OUTPUT>
  -- <KSLIST>
  --  <KS>
  --    <ID>����ID</ID>       //����ID
  --    <MC>��������</MC>     //��������
  --  </KS>
  --  <KS>
  --    ...
  --  </KS>
  -- </KSLIST>
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  v_Temp      Varchar(5000); --��ʱXML
  x_Templet   Xmltype; --ģ��XML
  v_Para      Varchar2(4000);
  n_��ѯ����  Number(5);
  n_ԤԼ����  Number(5);
  n_Add_Lists Number(3);
  v_������λ  ������λ���ſ���.������λ%Type;
  n_վ��      ���ű�.վ��%Type;
  v_Err_Msg   Varchar2(200);
  d_����ʱ��  Date;
  Err_Item Exception;
  n_�Һ�ģʽ Number(3);
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');
  Select Extractvalue(Value(A), 'IN/CXTS'), Extractvalue(Value(A), 'IN/HZDW'), Extractvalue(Value(A), 'IN/ZD')
  Into n_��ѯ����, v_������λ, n_վ��
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  v_Para     := zl_GetSysParameter('�Һ��Ű�ģʽ');
  n_ԤԼ���� := zl_GetSysParameter(66);
  n_�Һ�ģʽ := To_Number(Substr(v_Para, 1, 1));
  If n_�Һ�ģʽ = 1 Then
    Begin
      d_����ʱ�� := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
    Exception
      When Others Then
        d_����ʱ�� := Null;
    End;
  End If;
  If n_�Һ�ģʽ = 0 Then
    If n_��ѯ���� Is Null Then
      If v_������λ Is Null Then
        For r_Dept In (Select Distinct a.����id, b.����
                       From �ҺŰ��� A, ���ű� B
                       Where a.ͣ������ Is Null And a.����id = b.Id And (b.վ�� = Nvl(n_վ��, 0) Or b.վ�� Is Null)) Loop
        
          If Nvl(n_Add_Lists, 0) = 0 Then
            --����DJList�ڵ�
            Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype('<KSLIST></KSLIST>')) Into x_Templet From Dual;
            n_Add_Lists := 1;
          End If;
          v_Temp := '<KS>' || '<ID>' || r_Dept.����id || '</ID>' || '<MC>' || r_Dept.���� || '</MC>' || '</KS>';
          Select Appendchildxml(x_Templet, '/OUTPUT/KSLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
        End Loop;
      Else
        For r_Dept In (Select Distinct ����id, ����
                       From (Select b.����id, d.����
                              From (Select a.Id, Nvl(a.ԤԼ����, n_ԤԼ����) As ԤԼ����
                                     From �ҺŰ��� A
                                     Where a.ͣ������ Is Null And Not Exists
                                      (Select 1 From ������λ���ſ��� Where ����id = a.Id And ������λ = v_������λ)
                                     Union All
                                     Select a.Id, Nvl(a.ԤԼ����, n_ԤԼ����) As ԤԼ����
                                     From �ҺŰ��� A, ������λ���ſ��� B
                                     Where a.ͣ������ Is Null And ������λ = v_������λ And a.Id = b.����id And b.���� <> 0) A, �ҺŰ��� B,
                                   �ҺŰ��żƻ� C, ���ű� D
                              Where a.Id = b.Id And c.����id = a.Id And c.���ʱ�� Is Not Null And
                                    ((c.��Чʱ�� < Sysdate And c.ʧЧʱ�� > Sysdate + a.ԤԼ����) Or
                                    (c.��Чʱ�� Between Sysdate And Sysdate + a.ԤԼ����) Or
                                    (c.ʧЧʱ�� Between Sysdate And Sysdate + a.ԤԼ����)) And Not Exists
                               (Select 1
                                     From ������λ�ƻ�����
                                     Where �ƻ�id = c.Id And ������λ = v_������λ And ���� = 0) And b.����id = d.Id And
                                    (d.վ�� = Nvl(n_վ��, 0) Or d.վ�� Is Null) And
                                    (Sysdate Between d.����ʱ�� And d.����ʱ�� Or Sysdate >= d.����ʱ�� And d.����ʱ�� Is Null)
                              Union All
                              Select b.����id, d.����
                              From (Select a.Id
                                     From �ҺŰ��� A
                                     Where a.ͣ������ Is Null And Not Exists
                                      (Select 1 From ������λ���ſ��� Where ����id = a.Id And ������λ = v_������λ)
                                     Union All
                                     Select a.Id
                                     From �ҺŰ��� A, ������λ���ſ��� B
                                     Where a.ͣ������ Is Null And ������λ = v_������λ And a.Id = b.����id And b.���� <> 0) A, �ҺŰ��� B,
                                   ���ű� D
                              Where a.Id = b.Id And Not Exists
                               (Select 1 From �ҺŰ��żƻ� Where ����id = a.Id And Rownum < 2) And b.����id = d.Id And
                                    (d.վ�� = Nvl(n_վ��, 0) Or d.վ�� Is Null) And
                                    (Sysdate Between d.����ʱ�� And d.����ʱ�� Or Sysdate >= d.����ʱ�� And d.����ʱ�� Is Null))) Loop
          If Nvl(n_Add_Lists, 0) = 0 Then
            --����DJList�ڵ�
            Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype('<KSLIST></KSLIST>')) Into x_Templet From Dual;
            n_Add_Lists := 1;
          End If;
          v_Temp := '<KS>' || '<ID>' || r_Dept.����id || '</ID>' || '<MC>' || r_Dept.���� || '</MC>' || '</KS>';
          Select Appendchildxml(x_Templet, '/OUTPUT/KSLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
        End Loop;
      End If;
    Else
      If v_������λ Is Null Then
        For r_Dept In (Select Distinct a.����id, b.����
                       From �ҺŰ��� A, ���ű� B
                       Where a.ͣ������ Is Null And a.����id = b.Id And (b.վ�� = Nvl(n_վ��, 0) Or b.վ�� Is Null)) Loop
        
          If Nvl(n_Add_Lists, 0) = 0 Then
            --����DJList�ڵ�
            Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype('<KSLIST></KSLIST>')) Into x_Templet From Dual;
            n_Add_Lists := 1;
          End If;
          v_Temp := '<KS>' || '<ID>' || r_Dept.����id || '</ID>' || '<MC>' || r_Dept.���� || '</MC>' || '</KS>';
          Select Appendchildxml(x_Templet, '/OUTPUT/KSLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
        End Loop;
      Else
        For r_Dept In (Select Distinct ����id, ����
                       From (Select b.����id, d.����
                              From (Select a.Id
                                     From �ҺŰ��� A
                                     Where a.ͣ������ Is Null And Not Exists
                                      (Select 1 From ������λ���ſ��� Where ����id = a.Id And ������λ = v_������λ)
                                     Union All
                                     Select a.Id
                                     From �ҺŰ��� A, ������λ���ſ��� B
                                     Where a.ͣ������ Is Null And ������λ = v_������λ And a.Id = b.����id And b.���� <> 0) A, �ҺŰ��� B,
                                   �ҺŰ��żƻ� C, ���ű� D
                              Where a.Id = b.Id And c.����id = a.Id And c.���ʱ�� Is Not Null And
                                    ((c.��Чʱ�� < Sysdate And c.ʧЧʱ�� > Sysdate + n_��ѯ����) Or
                                    (c.��Чʱ�� Between Sysdate And Sysdate + n_��ѯ����) Or
                                    (c.ʧЧʱ�� Between Sysdate And Sysdate + n_��ѯ����)) And Not Exists
                               (Select 1
                                     From ������λ�ƻ�����
                                     Where �ƻ�id = c.Id And ������λ = v_������λ And ���� = 0) And b.����id = d.Id And
                                    (d.վ�� = Nvl(n_վ��, 0) Or d.վ�� Is Null) And
                                    (Sysdate Between d.����ʱ�� And d.����ʱ�� Or Sysdate >= d.����ʱ�� And d.����ʱ�� Is Null)
                              Union All
                              Select b.����id, d.����
                              From (Select a.Id
                                     From �ҺŰ��� A
                                     Where a.ͣ������ Is Null And Not Exists
                                      (Select 1 From ������λ���ſ��� Where ����id = a.Id And ������λ = v_������λ)
                                     Union All
                                     Select a.Id
                                     From �ҺŰ��� A, ������λ���ſ��� B
                                     Where a.ͣ������ Is Null And ������λ = v_������λ And a.Id = b.����id And b.���� <> 0) A, �ҺŰ��� B,
                                   ���ű� D
                              Where a.Id = b.Id And Not Exists
                               (Select 1 From �ҺŰ��żƻ� Where ����id = a.Id And Rownum < 2) And b.����id = d.Id And
                                    (d.վ�� = Nvl(n_վ��, 0) Or d.վ�� Is Null) And
                                    (Sysdate Between d.����ʱ�� And d.����ʱ�� Or Sysdate >= d.����ʱ�� And d.����ʱ�� Is Null))) Loop
          If Nvl(n_Add_Lists, 0) = 0 Then
            --����DJList�ڵ�
            Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype('<KSLIST></KSLIST>')) Into x_Templet From Dual;
            n_Add_Lists := 1;
          End If;
          v_Temp := '<KS>' || '<ID>' || r_Dept.����id || '</ID>' || '<MC>' || r_Dept.���� || '</MC>' || '</KS>';
          Select Appendchildxml(x_Templet, '/OUTPUT/KSLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
        End Loop;
      End If;
    End If;
  Else
    --������Ű�ģʽ
    If n_��ѯ���� Is Null Then
      If v_������λ Is Null Then
        For r_Dept In (Select Distinct a.����id, b.����
                       From �ҺŰ��� A, ���ű� B
                       Where a.ͣ������ Is Null And a.����id = b.Id And (b.վ�� = Nvl(n_վ��, 0) Or b.վ�� Is Null) And
                             Sysdate < d_����ʱ��
                       Union
                       Select Distinct a.����id, b.����
                       From �ٴ������¼ A, ���ű� B, �ٴ������Դ C
                       Where a.��Դid = c.Id And a.��ʼʱ�� > d_����ʱ�� And a.�������� >= Trunc(Sysdate) And
                             a.�������� <= Trunc(Sysdate + Nvl(c.ԤԼ����, n_ԤԼ����)) And Nvl(a.�Ƿ񷢲�, 0) = 1 And a.����id = b.Id And
                             (b.վ�� = Nvl(n_վ��, 0) Or b.վ�� Is Null)) Loop
          If Nvl(n_Add_Lists, 0) = 0 Then
            Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype('<KSLIST></KSLIST>')) Into x_Templet From Dual;
            n_Add_Lists := 1;
          End If;
          v_Temp := '<KS>' || '<ID>' || r_Dept.����id || '</ID>' || '<MC>' || r_Dept.���� || '</MC>' || '</KS>';
          Select Appendchildxml(x_Templet, '/OUTPUT/KSLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
        End Loop;
      Else
        For r_Dept In (Select Distinct ����id, ����
                       From (Select b.����id, d.����
                              From (Select a.Id, Nvl(a.ԤԼ����, n_ԤԼ����) As ԤԼ����
                                     From �ҺŰ��� A
                                     Where a.ͣ������ Is Null And Not Exists
                                      (Select 1 From ������λ���ſ��� Where ����id = a.Id And ������λ = v_������λ)
                                     Union All
                                     Select a.Id, Nvl(a.ԤԼ����, n_ԤԼ����) As ԤԼ����
                                     From �ҺŰ��� A, ������λ���ſ��� B
                                     Where a.ͣ������ Is Null And ������λ = v_������λ And a.Id = b.����id And b.���� <> 0) A, �ҺŰ��� B,
                                   �ҺŰ��żƻ� C, ���ű� D
                              Where a.Id = b.Id And Sysdate < d_����ʱ�� And c.����id = a.Id And c.���ʱ�� Is Not Null And
                                    ((c.��Чʱ�� < Sysdate And c.ʧЧʱ�� > Sysdate + a.ԤԼ����) Or
                                    (c.��Чʱ�� Between Sysdate And Sysdate + a.ԤԼ����) Or
                                    (c.ʧЧʱ�� Between Sysdate And Sysdate + a.ԤԼ����)) And Not Exists
                               (Select 1
                                     From ������λ�ƻ�����
                                     Where �ƻ�id = c.Id And ������λ = v_������λ And ���� = 0) And b.����id = d.Id And
                                    (d.վ�� = Nvl(n_վ��, 0) Or d.վ�� Is Null) And
                                    (Sysdate Between d.����ʱ�� And d.����ʱ�� Or Sysdate >= d.����ʱ�� And d.����ʱ�� Is Null)
                              Union All
                              Select b.����id, d.����
                              From (Select a.Id
                                     From �ҺŰ��� A
                                     Where a.ͣ������ Is Null And Not Exists
                                      (Select 1 From ������λ���ſ��� Where ����id = a.Id And ������λ = v_������λ)
                                     Union All
                                     Select a.Id
                                     From �ҺŰ��� A, ������λ���ſ��� B
                                     Where a.ͣ������ Is Null And ������λ = v_������λ And a.Id = b.����id And b.���� <> 0) A, �ҺŰ��� B,
                                   ���ű� D
                              Where a.Id = b.Id And Sysdate < d_����ʱ�� And Not Exists
                               (Select 1 From �ҺŰ��żƻ� Where ����id = a.Id And Rownum < 2) And b.����id = d.Id And
                                    (d.վ�� = Nvl(n_վ��, 0) Or d.վ�� Is Null) And
                                    (Sysdate Between d.����ʱ�� And d.����ʱ�� Or Sysdate >= d.����ʱ�� And d.����ʱ�� Is Null))
                       Union
                       Select Distinct ����id, ����
                       From (Select a.����id, b.����
                              From �ٴ������¼ A, ���ű� B, �ٴ������Դ C
                              Where a.��Դid = c.Id And a.��ʼʱ�� > d_����ʱ�� And a.�������� >= Trunc(Sysdate) And
                                    a.�������� <= Trunc(Sysdate + Nvl(c.ԤԼ����, n_ԤԼ����)) And Nvl(a.�Ƿ񷢲�, 0) = 1 And
                                    a.����id = b.Id And (b.վ�� = Nvl(n_վ��, 0) Or b.վ�� Is Null) And Not Exists
                               (Select 1 From �ٴ�����Һſ��Ƽ�¼ Where ��¼id = a.Id And ���� = 1 And ���� = 1)
                              Union
                              Select a.����id, b.����
                              From �ٴ������¼ A, ���ű� B, �ٴ������Դ C
                              Where a.��Դid = c.Id And a.��ʼʱ�� > d_����ʱ�� And a.�������� >= Trunc(Sysdate) And
                                    a.�������� <= Trunc(Sysdate + Nvl(c.ԤԼ����, n_ԤԼ����)) And Nvl(a.�Ƿ񷢲�, 0) = 1 And
                                    a.����id = b.Id And (b.վ�� = Nvl(n_վ��, 0) Or b.վ�� Is Null) And Exists
                               (Select 1
                                     From �ٴ�����Һſ��Ƽ�¼
                                     Where ��¼id = a.Id And ���� = v_������λ And ���� = 1 And ���� = 1 And ���Ʒ�ʽ <> 0))) Loop
          If Nvl(n_Add_Lists, 0) = 0 Then
            Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype('<KSLIST></KSLIST>')) Into x_Templet From Dual;
            n_Add_Lists := 1;
          End If;
          v_Temp := '<KS>' || '<ID>' || r_Dept.����id || '</ID>' || '<MC>' || r_Dept.���� || '</MC>' || '</KS>';
          Select Appendchildxml(x_Templet, '/OUTPUT/KSLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
        End Loop;
      End If;
    Else
      If v_������λ Is Null Then
        For r_Dept In (Select Distinct a.����id, b.����
                       From �ҺŰ��� A, ���ű� B
                       Where a.ͣ������ Is Null And a.����id = b.Id And (b.վ�� = Nvl(n_վ��, 0) Or b.վ�� Is Null) And
                             Sysdate < d_����ʱ��
                       Union
                       Select Distinct a.����id, b.����
                       From �ٴ������¼ A, ���ű� B
                       Where a.�������� >= Trunc(Sysdate) And a.��ʼʱ�� > d_����ʱ�� And a.�������� <= Trunc(Sysdate + n_��ѯ����) And
                             Nvl(a.�Ƿ񷢲�, 0) = 1 And a.����id = b.Id And (b.վ�� = Nvl(n_վ��, 0) Or b.վ�� Is Null)) Loop
          If Nvl(n_Add_Lists, 0) = 0 Then
            Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype('<KSLIST></KSLIST>')) Into x_Templet From Dual;
            n_Add_Lists := 1;
          End If;
          v_Temp := '<KS>' || '<ID>' || r_Dept.����id || '</ID>' || '<MC>' || r_Dept.���� || '</MC>' || '</KS>';
          Select Appendchildxml(x_Templet, '/OUTPUT/KSLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
        End Loop;
      Else
        For r_Dept In (Select Distinct ����id, ����
                       From (Select b.����id, d.����
                              From (Select a.Id
                                     From �ҺŰ��� A
                                     Where a.ͣ������ Is Null And Not Exists
                                      (Select 1 From ������λ���ſ��� Where ����id = a.Id And ������λ = v_������λ)
                                     Union All
                                     Select a.Id
                                     From �ҺŰ��� A, ������λ���ſ��� B
                                     Where a.ͣ������ Is Null And ������λ = v_������λ And a.Id = b.����id And b.���� <> 0) A, �ҺŰ��� B,
                                   �ҺŰ��żƻ� C, ���ű� D
                              Where a.Id = b.Id And Sysdate < d_����ʱ�� And c.����id = a.Id And c.���ʱ�� Is Not Null And
                                    ((c.��Чʱ�� < Sysdate And c.ʧЧʱ�� > Sysdate + n_��ѯ����) Or
                                    (c.��Чʱ�� Between Sysdate And Sysdate + n_��ѯ����) Or
                                    (c.ʧЧʱ�� Between Sysdate And Sysdate + n_��ѯ����)) And Not Exists
                               (Select 1
                                     From ������λ�ƻ�����
                                     Where �ƻ�id = c.Id And ������λ = v_������λ And ���� = 0) And b.����id = d.Id And
                                    (d.վ�� = Nvl(n_վ��, 0) Or d.վ�� Is Null) And
                                    (Sysdate Between d.����ʱ�� And d.����ʱ�� Or Sysdate >= d.����ʱ�� And d.����ʱ�� Is Null)
                              Union All
                              Select b.����id, d.����
                              From (Select a.Id
                                     From �ҺŰ��� A
                                     Where a.ͣ������ Is Null And Not Exists
                                      (Select 1 From ������λ���ſ��� Where ����id = a.Id And ������λ = v_������λ)
                                     Union All
                                     Select a.Id
                                     From �ҺŰ��� A, ������λ���ſ��� B
                                     Where a.ͣ������ Is Null And ������λ = v_������λ And a.Id = b.����id And b.���� <> 0) A, �ҺŰ��� B,
                                   ���ű� D
                              Where a.Id = b.Id And Sysdate < d_����ʱ�� And Not Exists
                               (Select 1 From �ҺŰ��żƻ� Where ����id = a.Id And Rownum < 2) And b.����id = d.Id And
                                    (d.վ�� = Nvl(n_վ��, 0) Or d.վ�� Is Null) And
                                    (Sysdate Between d.����ʱ�� And d.����ʱ�� Or Sysdate >= d.����ʱ�� And d.����ʱ�� Is Null))
                       Union
                       Select Distinct ����id, ����
                       From (Select a.����id, b.����
                              From �ٴ������¼ A, ���ű� B
                              Where a.�������� >= Trunc(Sysdate) And a.��ʼʱ�� > d_����ʱ�� And a.�������� <= Trunc(Sysdate + n_��ѯ����) And
                                    Nvl(a.�Ƿ񷢲�, 0) = 1 And a.����id = b.Id And (b.վ�� = Nvl(n_վ��, 0) Or b.վ�� Is Null) And
                                    Not Exists
                               (Select 1 From �ٴ�����Һſ��Ƽ�¼ Where ��¼id = a.Id And ���� = 1 And ���� = 1)
                              Union
                              Select a.����id, b.����
                              From �ٴ������¼ A, ���ű� B
                              Where a.�������� >= Trunc(Sysdate) And a.��ʼʱ�� > d_����ʱ�� And a.�������� <= Trunc(Sysdate + n_��ѯ����) And
                                    Nvl(a.�Ƿ񷢲�, 0) = 1 And a.����id = b.Id And (b.վ�� = Nvl(n_վ��, 0) Or b.վ�� Is Null) And
                                    Exists
                               (Select 1
                                     From �ٴ�����Һſ��Ƽ�¼
                                     Where ��¼id = a.Id And ���� = v_������λ And ���� = 1 And ���� = 1 And ���Ʒ�ʽ <> 0))) Loop
          If Nvl(n_Add_Lists, 0) = 0 Then
            Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype('<KSLIST></KSLIST>')) Into x_Templet From Dual;
            n_Add_Lists := 1;
          End If;
          v_Temp := '<KS>' || '<ID>' || r_Dept.����id || '</ID>' || '<MC>' || r_Dept.���� || '</MC>' || '</KS>';
          Select Appendchildxml(x_Templet, '/OUTPUT/KSLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
        End Loop;
      End If;
    End If;
  End If;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getdeptlist;
/


Create Or Replace Procedure Zl_Third_Docarrange
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --����:ҽ���Ű�ƻ�
  --���:Xml_In:
  --<IN>
  --   <YSID>870</YSID>    //ҽ��ID
  --   <KDID>870</KSID>    //����ID
  --   <KSSJ>2014-10-29 </KSSJ>    //��ʼʱ��
  --   <CXTS>14</CXTS>    //��ѯ����
  --   <HZDW>֧����</HZDW> //������λ
  --   <HL>����</HL>      //���࣬�ɴ�������ö��ŷָ�����ʽ:��ͨ,ר��,...
  --</IN>

  --����:Xml_Out
  --<OUTPUT>
  --   <PBLIST>       //δ���ظýڵ��ʾû������
  --    <PB>
  --     <RQ>2014-10-29</RQ>     //����
  --     <SYHS>5</SYHS>    //ʣ�����
  --     <SBSJ>ȫ��</SBSJ>             //�ϰ�ʱ��
  --     <YGS>5</YGS>    //�ѹҺ���
  --    </PB>
  --   <PBLIST>
  --   <ERROR><MSG></MSG></ERROR> //�����������
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  d_����         Date;
  v_�Ű�         �ҺŰ���.����%Type;
  n_�޺���       �ҺŰ�������.�޺���%Type;
  n_�ѹ���       �ҺŰ�������.�޺���%Type;
  n_���ѹ���     �ҺŰ�������.�޺���%Type;
  n_��Լ��       �ҺŰ�������.�޺���%Type;
  n_��Լ��       �ҺŰ�������.�޺���%Type;
  n_ʣ����       �ҺŰ�������.�޺���%Type;
  v_�ϰ�ʱ��     Varchar2(300);
  n_ҽ��id       ��Ա��.Id%Type;
  n_����id       ���ű�.Id%Type;
  n_��ѯ����     Number(4);
  n_������λ���� Number(5);
  n_��Լ�ѹ���   Number(4);
  n_��Լ����     Number(3);
  n_���Ŵ���     Number(3);
  v_����         �ҺŰ���.����%Type;
  n_����id       �ҺŰ��żƻ�.����id%Type;
  n_�ƻ�id       �ҺŰ��żƻ�.Id%Type;
  v_������λ     �Һź�����λ.����%Type;
  n_Daycount     Number(4);
  d_��ʼʱ��     Date;
  d_ԭʼʱ��     Date;
  n_����         Number(3);
  v_Temp         Varchar2(32767); --��ʱXML
  x_Templet      Xmltype; --ģ��XML
  v_Err_Msg      Varchar2(200);
  v_����         Varchar2(200);
  n_Exists       Number(2);
  n_�Һ�ģʽ     Number(3);
  n_��Լģʽ     �ٴ�����Һſ��Ƽ�¼.���Ʒ�ʽ%Type;
  v_����ʱ��     Varchar2(500);
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/YSID'), Extractvalue(Value(A), 'IN/KSID'), Extractvalue(Value(A), 'IN/CXTS'),
         To_Date(Extractvalue(Value(A), 'IN/KSSJ'), 'YYYY-MM-DD hh24:mi:ss'), Extractvalue(Value(A), 'IN/HZDW'),
         Extractvalue(Value(A), 'IN/HL')
  Into n_ҽ��id, n_����id, n_��ѯ����, d_��ʼʱ��, v_������λ, v_����
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  n_��ѯ���� := Nvl(n_��ѯ����, 14);
  d_ԭʼʱ�� := Trunc(d_��ʼʱ��);
  d_��ʼʱ�� := Trunc(d_��ʼʱ��);
  n_Daycount := 0;
  n_�Һ�ģʽ := To_Number(Substr(Nvl(zl_GetSysParameter('�Һ��Ű�ģʽ'), '0'), 1, 1));
  v_����ʱ�� := Substr(Nvl(zl_GetSysParameter('�Һ��Ű�ģʽ'), '0'), 3);
  If n_�Һ�ģʽ = 0 Then
    If Nvl(n_����id, 0) = 0 Then
      While (n_Daycount < n_��ѯ����) Loop
        If Trunc(d_��ʼʱ�� + n_Daycount) = Trunc(Sysdate) Then
          d_��ʼʱ�� := Sysdate - n_Daycount;
        Else
          d_��ʼʱ�� := d_ԭʼʱ��;
        End If;
        n_���Ŵ��� := 0;
        v_�ϰ�ʱ�� := Null;
        n_���ѹ��� := 0;
        n_�ѹ���   := 0;
        n_ʣ����   := 0;
        n_�޺���   := 0;
        n_��Լ��   := 0;
        n_��Լ��   := 0;
        For r_�Ű� In (Select d_��ʼʱ�� + n_Daycount As ����, a.�Ű�, a.�޺���, a.��Լ��, Nvl(Hz.�ѹ���, 0) As �ѹ���, Nvl(Hz.��Լ��, 0) As ��Լ��,
                            a.����id, a.�ƻ�id, a.����, a.����
                     
                     From (Select Ap.����id, Ap.����, Bm.���� As ��������, Ap.ҽ������, Nvl(Ap.ҽ��id, 0) As ҽ��id, Ry.רҵ����ְ�� As ְ��, Ap.����,
                                   Ap.����id, Ap.�ƻ�id, Ap.�Ű�, Ap.��Ŀid, Fy.���� As ��Ŀ����, Ap.��ſ���, Ap.�޺���, Ap.��Լ��
                            
                            From (Select Ap.����id, Ap.����, Bm.���� As ��������, Ap.ҽ������, Ap.ҽ��id, Ap.����, Ap.Id As ����id, 0 As �ƻ�id,
                                          Ap.��Ŀid, Nvl(Ap.��ſ���, 0) As ��ſ���,
                                          Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', Ap.����, '2', Ap.��һ, '3', Ap.�ܶ�, '4',
                                                  Ap.����, '5', Ap.����, '6', Ap.����, '7', Ap.����, Null) As �Ű�,
                                          Nvl(Xz.��Լ��, Xz.�޺���) As ��Լ��, Nvl(Xz.�޺���, 0) As �޺���
                                   From �ҺŰ��� Ap, ���ű� Bm, �ҺŰ������� Xz
                                   Where Ap.����id = Bm.Id(+) And Ap.ҽ��id = n_ҽ��id And Ap.ͣ������ Is Null And
                                         d_��ʼʱ�� + n_Daycount Between Nvl(Ap.��ʼʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                         Nvl(Ap.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.����id(+) = Ap.Id And
                                         Xz.������Ŀ(+) =
                                         Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                                '5', '����', '6', '����', '7', '����', Null) And Not Exists
                                    (Select Rownum
                                          From �ҺŰ���ͣ��״̬ Ty
                                          Where Ty.����id = Ap.Id And d_��ʼʱ�� + n_Daycount Between Ty.��ʼֹͣʱ�� And Ty.����ֹͣʱ��) And
                                         Not Exists
                                    (Select Rownum
                                          From �ҺŰ��żƻ� Jh
                                          Where Jh.����id = Ap.Id And Jh.���ʱ�� Is Not Null And
                                                d_��ʼʱ�� + n_Daycount Between Nvl(Jh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                Nvl(Jh.ʧЧʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')))
                                   Union All
                                   Select Ap.����id, Ap.����, Bm.���� As ��������, Jh.ҽ������, Jh.ҽ��id, Ap.����, Ap.Id As ����id,
                                          Jh.Id As �ƻ�id, Jh.��Ŀid, Nvl(Jh.��ſ���, 0) As ��ſ���,
                                          Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', Jh.����, '2', Jh.��һ, '3', Jh.�ܶ�, '4',
                                                  Jh.����, '5', Jh.����, '6', Jh.����, '7', Jh.����, Null) As �Ű�,
                                          Nvl(Xz.��Լ��, Xz.�޺���) As ��Լ��, Nvl(Xz.�޺���, 0) As �޺���
                                   From �ҺŰ��żƻ� Jh, �ҺŰ��� Ap, ���ű� Bm, �Һżƻ����� Xz
                                   Where Jh.����id = Ap.Id And Ap.����id = Bm.Id(+) And Ap.ͣ������ Is Null And Jh.ҽ��id = n_ҽ��id And
                                         d_��ʼʱ�� + n_Daycount Between Nvl(Jh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                         Nvl(Jh.ʧЧʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.�ƻ�id(+) = Jh.Id And
                                         Xz.������Ŀ(+) =
                                         Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                                '5', '����', '6', '����', '7', '����', Null) And Not Exists
                                    (Select Rownum
                                          From �ҺŰ���ͣ��״̬ Ty
                                          Where Ty.����id = Ap.Id And d_��ʼʱ�� + n_Daycount Between Ty.��ʼֹͣʱ�� And Ty.����ֹͣʱ��) And
                                         (Jh.��Чʱ��, Jh.����id) =
                                         (Select Max(Sxjh.��Чʱ��) As ��Чʱ��, ����id
                                          From �ҺŰ��żƻ� Sxjh
                                          Where Sxjh.���ʱ�� Is Not Null And d_��ʼʱ�� + n_Daycount Between
                                                Nvl(Sxjh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                Nvl(Sxjh.ʧЧʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Sxjh.����id = Jh.����id
                                          Group By Sxjh.����id)) Ap, ���ű� Bm, ��Ա�� Ry, �շ���ĿĿ¼ Fy
                            Where Ap.����id = Bm.Id(+) And Ap.ҽ��id = Ry.Id(+) And Ap.��Ŀid = Fy.Id And �Ű� Is Not Null) A,
                          ���˹ҺŻ��� Hz, �շѼ�Ŀ B
                     Where a.���� = Hz.����(+) And Hz.����(+) = Trunc(d_��ʼʱ�� + n_Daycount) And a.��Ŀid = b.�շ�ϸĿid And
                           b.��ֹ���� > d_��ʼʱ�� + n_Daycount And b.ִ������ <= d_��ʼʱ�� + n_Daycount
                     
                     ) Loop
          If v_���� Is Null Or Instr(',' || v_���� || ',', ',' || r_�Ű�.���� || ',') > 0 Then
            v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_�Ű�.�Ű�;
            n_���ѹ��� := n_���ѹ��� + r_�Ű�.�ѹ���;
            n_�ѹ���   := r_�Ű�.�ѹ���;
            n_�޺���   := r_�Ű�.�޺���;
            n_��Լ��   := r_�Ű�.��Լ��;
            n_��Լ��   := r_�Ű�.��Լ��;
            n_����id   := Nvl(r_�Ű�.����id, 0);
            n_�ƻ�id   := Nvl(r_�Ű�.�ƻ�id, 0);
            v_����     := r_�Ű�.����;
            n_���Ŵ��� := 1;
            If v_�ϰ�ʱ�� Is Not Null Then
              If v_������λ Is Not Null Then
                If n_�ƻ�id <> 0 Then
                  Begin
                    Select 1
                    Into n_��Լ����
                    From ������λ�ƻ�����
                    Where �ƻ�id = n_�ƻ�id And ������λ = v_������λ And Rownum < 2;
                  Exception
                    When Others Then
                      n_��Լ���� := 0;
                  End;
                Else
                  Begin
                    Select 1
                    Into n_��Լ����
                    From ������λ���ſ���
                    Where ����id = n_����id And ������λ = v_������λ And Rownum < 2;
                  Exception
                    When Others Then
                      n_��Լ���� := 0;
                  End;
                End If;
              End If;
            
              If v_������λ Is Null Or Nvl(n_��Լ����, 0) = 0 Then
                If n_�ƻ�id <> 0 Then
                  Begin
                    Select Sum(����)
                    Into n_������λ����
                    From ������λ�ƻ�����
                    Where �ƻ�id = n_�ƻ�id And
                          ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                        '5', '����', '6', '����', '7', '����', Null);
                  Exception
                    When Others Then
                      n_������λ���� := 0;
                  End;
                Else
                  Begin
                    Select Sum(����)
                    Into n_������λ����
                    From ������λ���ſ���
                    Where ����id = n_����id And
                          ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                        '5', '����', '6', '����', '7', '����', Null);
                  Exception
                    When Others Then
                      n_������λ���� := 0;
                  End;
                End If;
                Begin
                  Select Count(1)
                  Into n_��Լ�ѹ���
                  From ���˹Һż�¼
                  Where �ű� = v_���� And ��¼״̬ = 1 And ������λ Is Not Null And ����ʱ�� Between Trunc(d_��ʼʱ�� + n_Daycount) And
                        Trunc(d_��ʼʱ�� + n_Daycount + 1) - 1 / 60 / 60 / 24;
                Exception
                  When Others Then
                    n_��Լ�ѹ��� := 0;
                End;
                If n_������λ���� = 0 Then
                  n_������λ���� := Null;
                End If;
                If Trunc(d_��ʼʱ�� + n_Daycount) > Trunc(Sysdate) Then
                  n_ʣ���� := Nvl(n_ʣ����, 0) + n_��Լ�� - n_��Լ�� - Nvl(n_������λ����, n_��Լ�ѹ���) + Nvl(n_��Լ�ѹ���, 0);
                Else
                  n_ʣ���� := Nvl(n_ʣ����, 0) + n_�޺��� - n_�ѹ��� - Nvl(n_������λ����, n_��Լ�ѹ���) + Nvl(n_��Լ�ѹ���, 0);
                End If;
              Else
                --��Լ��λ
                If n_�ƻ�id <> 0 Then
                  Begin
                    Select 1
                    Into n_����
                    From ������λ�ƻ�����
                    Where �ƻ�id = n_�ƻ�id And
                          ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                        '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ And ���� = 0 And
                          Rownum < 2;
                  Exception
                    When Others Then
                      n_���� := 0;
                  End;
                Else
                  Begin
                    Select 1
                    Into n_����
                    From ������λ���ſ���
                    Where ����id = n_����id And
                          ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                        '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ And ���� = 0 And
                          Rownum < 2;
                  Exception
                    When Others Then
                      n_���� := 0;
                  End;
                End If;
                If Nvl(n_����, 0) = 0 Then
                  Begin
                    Select Count(1)
                    Into n_��Լ�ѹ���
                    From ���˹Һż�¼
                    Where �ű� = v_���� And ��¼״̬ = 1 And ������λ = v_������λ And ����ʱ�� Between Trunc(d_��ʼʱ�� + n_Daycount) And
                          Trunc(d_��ʼʱ�� + n_Daycount + 1) - 1 / 60 / 60 / 24;
                  Exception
                    When Others Then
                      n_��Լ�ѹ��� := 0;
                  End;
                  If n_�ƻ�id <> 0 Then
                    Begin
                      Select Sum(����)
                      Into n_������λ����
                      From ������λ�ƻ�����
                      Where �ƻ�id = n_�ƻ�id And
                            ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                          '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ;
                    Exception
                      When Others Then
                        n_������λ���� := 0;
                    End;
                  Else
                    Begin
                      Select Sum(����)
                      Into n_������λ����
                      From ������λ���ſ���
                      Where ����id = n_����id And
                            ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                          '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ;
                    Exception
                      When Others Then
                        n_������λ���� := 0;
                    End;
                  End If;
                  If n_������λ���� = 0 Then
                    n_������λ���� := Null;
                  End If;
                  n_ʣ���� := Nvl(n_ʣ����, 0) + Nvl(n_������λ����, n_��Լ�ѹ���) - Nvl(n_��Լ�ѹ���, 0);
                
                End If;
              End If;
            End If;
            n_������λ���� := 0;
            n_��Լ����     := 0;
            n_����         := 0;
          End If;
        End Loop;
        v_�ϰ�ʱ�� := Substr(v_�ϰ�ʱ��, 2);
        If n_���Ŵ��� = 1 Then
          v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_��ʼʱ�� + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                    '<SYHS>' || n_ʣ���� || '</SYHS>' || '<SBSJ>' || v_�ϰ�ʱ�� || '</SBSJ>' || '<YGS>' || n_���ѹ��� || '</YGS>' ||
                    '</PB>';
        End If;
        n_Daycount := n_Daycount + 1;
      End Loop;
    Else
      While (n_Daycount < n_��ѯ����) Loop
        If Trunc(d_��ʼʱ�� + n_Daycount) = Trunc(Sysdate) Then
          d_��ʼʱ�� := Sysdate - n_Daycount;
        Else
          d_��ʼʱ�� := d_ԭʼʱ��;
        End If;
        v_�ϰ�ʱ�� := Null;
        n_���ѹ��� := 0;
        n_�ѹ���   := 0;
        n_ʣ����   := 0;
        n_�޺���   := 0;
        n_��Լ��   := 0;
        n_��Լ��   := 0;
        n_���Ŵ��� := 0;
        For r_�Ű� In (Select d_��ʼʱ�� + n_Daycount As ����, a.�Ű�, a.�޺���, a.��Լ��, Nvl(Hz.�ѹ���, 0) As �ѹ���, Nvl(Hz.��Լ��, 0) As ��Լ��,
                            a.����id, a.�ƻ�id, a.����, a.����
                     
                     From (Select Ap.����id, Ap.����, Bm.���� As ��������, Ap.ҽ������, Nvl(Ap.ҽ��id, 0) As ҽ��id, Ry.רҵ����ְ�� As ְ��, Ap.����,
                                   Ap.����id, Ap.�ƻ�id, Ap.�Ű�, Ap.��Ŀid, Fy.���� As ��Ŀ����, Ap.��ſ���, Ap.�޺���, Ap.��Լ��
                            
                            From (Select Ap.����id, Ap.����, Bm.���� As ��������, Ap.ҽ������, Ap.ҽ��id, Ap.����, Ap.Id As ����id, 0 As �ƻ�id,
                                          Ap.��Ŀid, Nvl(Ap.��ſ���, 0) As ��ſ���,
                                          Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', Ap.����, '2', Ap.��һ, '3', Ap.�ܶ�, '4',
                                                  Ap.����, '5', Ap.����, '6', Ap.����, '7', Ap.����, Null) As �Ű�,
                                          Nvl(Xz.��Լ��, Xz.�޺���) As ��Լ��, Nvl(Xz.�޺���, 0) As �޺���
                                   From �ҺŰ��� Ap, ���ű� Bm, �ҺŰ������� Xz
                                   Where Ap.����id = Bm.Id(+) And Ap.ҽ��id = n_ҽ��id And Ap.����id = n_����id And Ap.ͣ������ Is Null And
                                         d_��ʼʱ�� + n_Daycount Between Nvl(Ap.��ʼʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                         Nvl(Ap.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.����id(+) = Ap.Id And
                                         Xz.������Ŀ(+) =
                                         Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                                '5', '����', '6', '����', '7', '����', Null) And Not Exists
                                    (Select Rownum
                                          From �ҺŰ���ͣ��״̬ Ty
                                          Where Ty.����id = Ap.Id And d_��ʼʱ�� + n_Daycount Between Ty.��ʼֹͣʱ�� And Ty.����ֹͣʱ��) And
                                         Not Exists
                                    (Select Rownum
                                          From �ҺŰ��żƻ� Jh
                                          Where Jh.����id = Ap.Id And Jh.���ʱ�� Is Not Null And
                                                d_��ʼʱ�� + n_Daycount Between Nvl(Jh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                Nvl(Jh.ʧЧʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')))
                                   Union All
                                   Select Ap.����id, Ap.����, Bm.���� As ��������, Jh.ҽ������, Jh.ҽ��id, Ap.����, Ap.Id As ����id,
                                          Jh.Id As �ƻ�id, Jh.��Ŀid, Nvl(Jh.��ſ���, 0) As ��ſ���,
                                          Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', Jh.����, '2', Jh.��һ, '3', Jh.�ܶ�, '4',
                                                  Jh.����, '5', Jh.����, '6', Jh.����, '7', Jh.����, Null) As �Ű�,
                                          Nvl(Xz.��Լ��, Xz.�޺���) As ��Լ��, Nvl(Xz.�޺���, 0) As �޺���
                                   From �ҺŰ��żƻ� Jh, �ҺŰ��� Ap, ���ű� Bm, �Һżƻ����� Xz
                                   Where Jh.����id = Ap.Id And Ap.����id = Bm.Id(+) And Ap.ͣ������ Is Null And Jh.ҽ��id = n_ҽ��id And
                                         Ap.����id = n_����id And
                                         d_��ʼʱ�� + n_Daycount Between Nvl(Jh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                         Nvl(Jh.ʧЧʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.�ƻ�id(+) = Jh.Id And
                                         Xz.������Ŀ(+) =
                                         Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                                '5', '����', '6', '����', '7', '����', Null) And Not Exists
                                    (Select Rownum
                                          From �ҺŰ���ͣ��״̬ Ty
                                          Where Ty.����id = Ap.Id And d_��ʼʱ�� + n_Daycount Between Ty.��ʼֹͣʱ�� And Ty.����ֹͣʱ��) And
                                         (Jh.��Чʱ��, Jh.����id) =
                                         (Select Max(Sxjh.��Чʱ��) As ��Чʱ��, ����id
                                          From �ҺŰ��żƻ� Sxjh
                                          Where Sxjh.���ʱ�� Is Not Null And d_��ʼʱ�� + n_Daycount Between
                                                Nvl(Sxjh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                Nvl(Sxjh.ʧЧʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Sxjh.����id = Jh.����id
                                          Group By Sxjh.����id)) Ap, ���ű� Bm, ��Ա�� Ry, �շ���ĿĿ¼ Fy
                            Where Ap.����id = Bm.Id(+) And Ap.ҽ��id = Ry.Id(+) And Ap.��Ŀid = Fy.Id And �Ű� Is Not Null) A,
                          ���˹ҺŻ��� Hz, �շѼ�Ŀ B
                     Where a.���� = Hz.����(+) And Hz.����(+) = Trunc(d_��ʼʱ�� + n_Daycount) And a.��Ŀid = b.�շ�ϸĿid And
                           b.��ֹ���� > d_��ʼʱ�� + n_Daycount And b.ִ������ <= d_��ʼʱ�� + n_Daycount
                     
                     ) Loop
          If v_���� Is Null Or Instr(',' || v_���� || ',', ',' || r_�Ű�.���� || ',') > 0 Then
            v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_�Ű�.�Ű�;
            n_���ѹ��� := n_���ѹ��� + r_�Ű�.�ѹ���;
            n_�ѹ���   := r_�Ű�.�ѹ���;
            n_�޺���   := r_�Ű�.�޺���;
            n_��Լ��   := r_�Ű�.��Լ��;
            n_��Լ��   := r_�Ű�.��Լ��;
            n_����id   := Nvl(r_�Ű�.����id, 0);
            n_�ƻ�id   := Nvl(r_�Ű�.�ƻ�id, 0);
            v_����     := r_�Ű�.����;
            n_���Ŵ��� := 1;
          
            If v_�ϰ�ʱ�� Is Not Null Then
              If v_������λ Is Not Null Then
                If n_�ƻ�id <> 0 Then
                  Begin
                    Select 1
                    Into n_��Լ����
                    From ������λ�ƻ�����
                    Where �ƻ�id = n_�ƻ�id And ������λ = v_������λ And Rownum < 2;
                  Exception
                    When Others Then
                      n_��Լ���� := 0;
                  End;
                Else
                  Begin
                    Select 1
                    Into n_��Լ����
                    From ������λ���ſ���
                    Where ����id = n_����id And ������λ = v_������λ And Rownum < 2;
                  Exception
                    When Others Then
                      n_��Լ���� := 0;
                  End;
                End If;
              End If;
            
              If v_������λ Is Null Or Nvl(n_��Լ����, 0) = 0 Then
                If n_�ƻ�id <> 0 Then
                  Begin
                    Select Sum(����)
                    Into n_������λ����
                    From ������λ�ƻ�����
                    Where �ƻ�id = n_�ƻ�id And
                          ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                        '5', '����', '6', '����', '7', '����', Null);
                  Exception
                    When Others Then
                      n_������λ���� := 0;
                  End;
                Else
                  Begin
                    Select Sum(����)
                    Into n_������λ����
                    From ������λ���ſ���
                    Where ����id = n_����id And
                          ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                        '5', '����', '6', '����', '7', '����', Null);
                  Exception
                    When Others Then
                      n_������λ���� := 0;
                  End;
                End If;
                Begin
                  Select Count(1)
                  Into n_��Լ�ѹ���
                  From ���˹Һż�¼
                  Where �ű� = v_���� And ��¼״̬ = 1 And ������λ Is Not Null And ����ʱ�� Between Trunc(d_��ʼʱ�� + n_Daycount) And
                        Trunc(d_��ʼʱ�� + n_Daycount + 1) - 1 / 60 / 60 / 24;
                Exception
                  When Others Then
                    n_��Լ�ѹ��� := 0;
                End;
                If n_������λ���� = 0 Then
                  n_������λ���� := Null;
                End If;
                If Trunc(d_��ʼʱ�� + n_Daycount) > Trunc(Sysdate) Then
                  n_ʣ���� := Nvl(n_ʣ����, 0) + n_��Լ�� - n_��Լ�� - Nvl(n_������λ����, n_��Լ�ѹ���) + Nvl(n_��Լ�ѹ���, 0);
                Else
                  n_ʣ���� := Nvl(n_ʣ����, 0) + n_�޺��� - n_�ѹ��� - Nvl(n_������λ����, n_��Լ�ѹ���) + Nvl(n_��Լ�ѹ���, 0);
                End If;
              Else
                --��Լ��λ
                If n_�ƻ�id <> 0 Then
                  Begin
                    Select 1
                    Into n_����
                    From ������λ�ƻ�����
                    Where �ƻ�id = n_�ƻ�id And
                          ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                        '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ And ���� = 0 And
                          Rownum < 2;
                  Exception
                    When Others Then
                      n_���� := 0;
                  End;
                Else
                  Begin
                    Select 1
                    Into n_����
                    From ������λ���ſ���
                    Where ����id = n_����id And
                          ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                        '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ And ���� = 0 And
                          Rownum < 2;
                  Exception
                    When Others Then
                      n_���� := 0;
                  End;
                End If;
                If Nvl(n_����, 0) = 0 Then
                  Begin
                    Select Count(1)
                    Into n_��Լ�ѹ���
                    From ���˹Һż�¼
                    Where �ű� = v_���� And ��¼״̬ = 1 And ������λ = v_������λ And ����ʱ�� Between Trunc(d_��ʼʱ�� + n_Daycount) And
                          Trunc(d_��ʼʱ�� + n_Daycount + 1) - 1 / 60 / 60 / 24;
                  Exception
                    When Others Then
                      n_��Լ�ѹ��� := 0;
                  End;
                  If n_�ƻ�id <> 0 Then
                    Begin
                      Select Sum(����)
                      Into n_������λ����
                      From ������λ�ƻ�����
                      Where �ƻ�id = n_�ƻ�id And
                            ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                          '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ;
                    Exception
                      When Others Then
                        n_������λ���� := 0;
                    End;
                  Else
                    Begin
                      Select Sum(����)
                      Into n_������λ����
                      From ������λ���ſ���
                      Where ����id = n_����id And
                            ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                          '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ;
                    Exception
                      When Others Then
                        n_������λ���� := 0;
                    End;
                  End If;
                  If n_������λ���� = 0 Then
                    n_������λ���� := Null;
                  End If;
                  n_ʣ���� := Nvl(n_ʣ����, 0) + Nvl(n_������λ����, n_��Լ�ѹ���) - Nvl(n_��Լ�ѹ���, 0);
                
                End If;
              End If;
            End If;
            n_������λ���� := 0;
            n_��Լ����     := 0;
            n_����         := 0;
          End If;
        End Loop;
        v_�ϰ�ʱ�� := Substr(v_�ϰ�ʱ��, 2);
        If n_���Ŵ��� = 1 Then
          v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_��ʼʱ�� + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                    '<SYHS>' || n_ʣ���� || '</SYHS>' || '<SBSJ>' || v_�ϰ�ʱ�� || '</SBSJ>' || '<YGS>' || n_���ѹ��� || '</YGS>' ||
                    '</PB>';
        End If;
        n_Daycount := n_Daycount + 1;
      End Loop;
    End If;
    v_Temp := '<PBLIST>' || v_Temp || '</PBLIST>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Else
    --������Ű�ģʽ
    If Nvl(n_����id, 0) = 0 Then
      --ͨ��ҽ������
      While (n_Daycount < n_��ѯ����) Loop
        If Trunc(d_��ʼʱ�� + n_Daycount) < To_Date(Substr(v_����ʱ��, 1, Instr(v_����ʱ��, ' ') - 1), 'yyyy-mm-dd') Then
          n_���Ŵ��� := 0;
          v_�ϰ�ʱ�� := Null;
          n_���ѹ��� := 0;
          n_�ѹ���   := 0;
          n_ʣ����   := 0;
          n_�޺���   := 0;
          n_��Լ��   := 0;
          n_��Լ��   := 0;
          For r_�Ű� In (Select d_��ʼʱ�� + n_Daycount As ����, a.�Ű�, a.�޺���, a.��Լ��, Nvl(Hz.�ѹ���, 0) As �ѹ���, Nvl(Hz.��Լ��, 0) As ��Լ��,
                              a.����id, a.�ƻ�id, a.����, a.����
                       
                       From (Select Ap.����id, Ap.����, Bm.���� As ��������, Ap.ҽ������, Nvl(Ap.ҽ��id, 0) As ҽ��id, Ry.רҵ����ְ�� As ְ��,
                                     Ap.����, Ap.����id, Ap.�ƻ�id, Ap.�Ű�, Ap.��Ŀid, Fy.���� As ��Ŀ����, Ap.��ſ���, Ap.�޺���, Ap.��Լ��
                              
                              From (Select Ap.����id, Ap.����, Bm.���� As ��������, Ap.ҽ������, Ap.ҽ��id, Ap.����, Ap.Id As ����id, 0 As �ƻ�id,
                                            Ap.��Ŀid, Nvl(Ap.��ſ���, 0) As ��ſ���,
                                            Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', Ap.����, '2', Ap.��һ, '3', Ap.�ܶ�, '4',
                                                    Ap.����, '5', Ap.����, '6', Ap.����, '7', Ap.����, Null) As �Ű�,
                                            Nvl(Xz.��Լ��, Xz.�޺���) As ��Լ��, Nvl(Xz.�޺���, 0) As �޺���
                                     From �ҺŰ��� Ap, ���ű� Bm, �ҺŰ������� Xz
                                     Where Ap.����id = Bm.Id(+) And Ap.ҽ��id = n_ҽ��id And Ap.ͣ������ Is Null And
                                           d_��ʼʱ�� + n_Daycount Between Nvl(Ap.��ʼʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                           Nvl(Ap.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.����id(+) = Ap.Id And
                                           Xz.������Ŀ(+) =
                                           Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4',
                                                  '����', '5', '����', '6', '����', '7', '����', Null) And Not Exists
                                      (Select Rownum
                                            From �ҺŰ���ͣ��״̬ Ty
                                            Where Ty.����id = Ap.Id And d_��ʼʱ�� + n_Daycount Between Ty.��ʼֹͣʱ�� And Ty.����ֹͣʱ��) And
                                           Not Exists (Select Rownum
                                            From �ҺŰ��żƻ� Jh
                                            Where Jh.����id = Ap.Id And Jh.���ʱ�� Is Not Null And
                                                  d_��ʼʱ�� + n_Daycount Between
                                                  Nvl(Jh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                  Nvl(Jh.ʧЧʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')))
                                     Union All
                                     Select Ap.����id, Ap.����, Bm.���� As ��������, Jh.ҽ������, Jh.ҽ��id, Ap.����, Ap.Id As ����id,
                                            Jh.Id As �ƻ�id, Jh.��Ŀid, Nvl(Jh.��ſ���, 0) As ��ſ���,
                                            Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', Jh.����, '2', Jh.��һ, '3', Jh.�ܶ�, '4',
                                                    Jh.����, '5', Jh.����, '6', Jh.����, '7', Jh.����, Null) As �Ű�,
                                            Nvl(Xz.��Լ��, Xz.�޺���) As ��Լ��, Nvl(Xz.�޺���, 0) As �޺���
                                     From �ҺŰ��żƻ� Jh, �ҺŰ��� Ap, ���ű� Bm, �Һżƻ����� Xz
                                     Where Jh.����id = Ap.Id And Ap.����id = Bm.Id(+) And Ap.ͣ������ Is Null And Jh.ҽ��id = n_ҽ��id And
                                           d_��ʼʱ�� + n_Daycount Between Nvl(Jh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                           Nvl(Jh.ʧЧʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.�ƻ�id(+) = Jh.Id And
                                           Xz.������Ŀ(+) =
                                           Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4',
                                                  '����', '5', '����', '6', '����', '7', '����', Null) And Not Exists
                                      (Select Rownum
                                            From �ҺŰ���ͣ��״̬ Ty
                                            Where Ty.����id = Ap.Id And d_��ʼʱ�� + n_Daycount Between Ty.��ʼֹͣʱ�� And Ty.����ֹͣʱ��) And
                                           (Jh.��Чʱ��, Jh.����id) =
                                           (Select Max(Sxjh.��Чʱ��) As ��Чʱ��, ����id
                                            From �ҺŰ��żƻ� Sxjh
                                            Where Sxjh.���ʱ�� Is Not Null And
                                                  d_��ʼʱ�� + n_Daycount Between
                                                  Nvl(Sxjh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                  Nvl(Sxjh.ʧЧʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Sxjh.����id = Jh.����id
                                            Group By Sxjh.����id)) Ap, ���ű� Bm, ��Ա�� Ry, �շ���ĿĿ¼ Fy
                              Where Ap.����id = Bm.Id(+) And Ap.ҽ��id = Ry.Id(+) And Ap.��Ŀid = Fy.Id And �Ű� Is Not Null) A,
                            ���˹ҺŻ��� Hz, �շѼ�Ŀ B
                       Where a.���� = Hz.����(+) And Hz.����(+) = Trunc(d_��ʼʱ�� + n_Daycount) And a.��Ŀid = b.�շ�ϸĿid And
                             b.��ֹ���� > d_��ʼʱ�� + n_Daycount And b.ִ������ <= d_��ʼʱ�� + n_Daycount
                       
                       ) Loop
            If v_���� Is Null Or Instr(',' || v_���� || ',', ',' || r_�Ű�.���� || ',') > 0 Then
              v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_�Ű�.�Ű�;
              n_���ѹ��� := n_���ѹ��� + r_�Ű�.�ѹ���;
              n_�ѹ���   := r_�Ű�.�ѹ���;
              n_�޺���   := r_�Ű�.�޺���;
              n_��Լ��   := r_�Ű�.��Լ��;
              n_��Լ��   := r_�Ű�.��Լ��;
              n_����id   := Nvl(r_�Ű�.����id, 0);
              n_�ƻ�id   := Nvl(r_�Ű�.�ƻ�id, 0);
              v_����     := r_�Ű�.����;
              n_���Ŵ��� := 1;
              If v_�ϰ�ʱ�� Is Not Null Then
                If v_������λ Is Not Null Then
                  If n_�ƻ�id <> 0 Then
                    Begin
                      Select 1
                      Into n_��Լ����
                      From ������λ�ƻ�����
                      Where �ƻ�id = n_�ƻ�id And ������λ = v_������λ And Rownum < 2;
                    Exception
                      When Others Then
                        n_��Լ���� := 0;
                    End;
                  Else
                    Begin
                      Select 1
                      Into n_��Լ����
                      From ������λ���ſ���
                      Where ����id = n_����id And ������λ = v_������λ And Rownum < 2;
                    Exception
                      When Others Then
                        n_��Լ���� := 0;
                    End;
                  End If;
                End If;
              
                If v_������λ Is Null Or Nvl(n_��Լ����, 0) = 0 Then
                  If n_�ƻ�id <> 0 Then
                    Begin
                      Select Sum(����)
                      Into n_������λ����
                      From ������λ�ƻ�����
                      Where �ƻ�id = n_�ƻ�id And
                            ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                          '5', '����', '6', '����', '7', '����', Null);
                    Exception
                      When Others Then
                        n_������λ���� := 0;
                    End;
                  Else
                    Begin
                      Select Sum(����)
                      Into n_������λ����
                      From ������λ���ſ���
                      Where ����id = n_����id And
                            ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                          '5', '����', '6', '����', '7', '����', Null);
                    Exception
                      When Others Then
                        n_������λ���� := 0;
                    End;
                  End If;
                  Begin
                    Select Count(1)
                    Into n_��Լ�ѹ���
                    From ���˹Һż�¼
                    Where �ű� = v_���� And ��¼״̬ = 1 And ������λ Is Not Null And ����ʱ�� Between Trunc(d_��ʼʱ�� + n_Daycount) And
                          Trunc(d_��ʼʱ�� + n_Daycount + 1) - 1 / 60 / 60 / 24;
                  Exception
                    When Others Then
                      n_��Լ�ѹ��� := 0;
                  End;
                  If n_������λ���� = 0 Then
                    n_������λ���� := Null;
                  End If;
                  If Trunc(d_��ʼʱ�� + n_Daycount) > Trunc(Sysdate) Then
                    n_ʣ���� := Nvl(n_ʣ����, 0) + n_��Լ�� - n_��Լ�� - Nvl(n_������λ����, n_��Լ�ѹ���) + Nvl(n_��Լ�ѹ���, 0);
                  Else
                    n_ʣ���� := Nvl(n_ʣ����, 0) + n_�޺��� - n_�ѹ��� - Nvl(n_������λ����, n_��Լ�ѹ���) + Nvl(n_��Լ�ѹ���, 0);
                  End If;
                Else
                  --��Լ��λ
                  If n_�ƻ�id <> 0 Then
                    Begin
                      Select 1
                      Into n_����
                      From ������λ�ƻ�����
                      Where �ƻ�id = n_�ƻ�id And
                            ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                          '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ And ���� = 0 And
                            Rownum < 2;
                    Exception
                      When Others Then
                        n_���� := 0;
                    End;
                  Else
                    Begin
                      Select 1
                      Into n_����
                      From ������λ���ſ���
                      Where ����id = n_����id And
                            ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                          '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ And ���� = 0 And
                            Rownum < 2;
                    Exception
                      When Others Then
                        n_���� := 0;
                    End;
                  End If;
                  If Nvl(n_����, 0) = 0 Then
                    Begin
                      Select Count(1)
                      Into n_��Լ�ѹ���
                      From ���˹Һż�¼
                      Where �ű� = v_���� And ��¼״̬ = 1 And ������λ = v_������λ And ����ʱ�� Between Trunc(d_��ʼʱ�� + n_Daycount) And
                            Trunc(d_��ʼʱ�� + n_Daycount + 1) - 1 / 60 / 60 / 24;
                    Exception
                      When Others Then
                        n_��Լ�ѹ��� := 0;
                    End;
                    If n_�ƻ�id <> 0 Then
                      Begin
                        Select Sum(����)
                        Into n_������λ����
                        From ������λ�ƻ�����
                        Where �ƻ�id = n_�ƻ�id And
                              ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4',
                                            '����', '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ;
                      Exception
                        When Others Then
                          n_������λ���� := 0;
                      End;
                    Else
                      Begin
                        Select Sum(����)
                        Into n_������λ����
                        From ������λ���ſ���
                        Where ����id = n_����id And
                              ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4',
                                            '����', '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ;
                      Exception
                        When Others Then
                          n_������λ���� := 0;
                      End;
                    End If;
                    If n_������λ���� = 0 Then
                      n_������λ���� := Null;
                    End If;
                    n_ʣ���� := Nvl(n_ʣ����, 0) + Nvl(n_������λ����, n_��Լ�ѹ���) - Nvl(n_��Լ�ѹ���, 0);
                  
                  End If;
                End If;
              End If;
              n_������λ���� := 0;
              n_��Լ����     := 0;
              n_����         := 0;
            End If;
          End Loop;
          v_�ϰ�ʱ�� := Substr(v_�ϰ�ʱ��, 2);
          If n_���Ŵ��� = 1 Then
            v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_��ʼʱ�� + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                      '<SYHS>' || n_ʣ���� || '</SYHS>' || '<SBSJ>' || v_�ϰ�ʱ�� || '</SBSJ>' || '<YGS>' || n_���ѹ��� ||
                      '</YGS>' || '</PB>';
          End If;
        Else
          n_���Ŵ��� := 0;
          v_�ϰ�ʱ�� := Null;
          n_���ѹ��� := 0;
          n_�ѹ���   := 0;
          n_ʣ����   := 0;
          n_�޺���   := 0;
          n_��Լ��   := 0;
          n_��Լ��   := 0;
          If v_������λ Is Null Then
            --�Ǻ�����λ
            If Trunc(d_��ʼʱ�� + n_Daycount) = Trunc(Sysdate) Then
              --����Һ�
              For r_���� In (Select �ѹ���, �޺���, �ϰ�ʱ��
                           From �ٴ������¼ A
                           Where �������� = Trunc(d_��ʼʱ�� + n_Daycount) And ҽ��id = n_ҽ��id And
                                 (a.��ʼʱ�� < Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��) Or a.��ֹʱ�� > Nvl(a.ͣ����ֹʱ��, a.��ʼʱ��)) And
                                 Nvl(a.�Ƿ�����, 0) = 0 And Exists
                            (Select 1
                                  From �ٴ����ﰲ�� M, �ٴ������ N
                                  Where m.Id = a.����id And m.����id = n.Id And n.����ʱ�� Is Not Null)) Loop
                n_�ѹ���   := n_�ѹ��� + Nvl(r_����.�ѹ���, 0);
                n_�޺���   := n_�޺��� + r_����.�޺���;
                v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                n_���Ŵ��� := 1;
              End Loop;
              If v_�ϰ�ʱ�� Is Not Null Then
                v_�ϰ�ʱ�� := Substr(v_�ϰ�ʱ��, 2);
              End If;
              If n_���Ŵ��� = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_��ʼʱ�� + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_�޺��� - n_�ѹ��� || '</SYHS>' || '<SBSJ>' || v_�ϰ�ʱ�� || '</SBSJ>' || '<YGS>' || n_�ѹ��� ||
                          '</YGS>' || '</PB>';
              End If;
            Else
              --ԤԼ�Һ�
              For r_���� In (Select ��Լ��, �޺���, ��Լ��, �ϰ�ʱ��
                           From �ٴ������¼ A
                           Where �������� = Trunc(d_��ʼʱ�� + n_Daycount) And ҽ��id = n_ҽ��id And
                                 (a.��ʼʱ�� < Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��) Or a.��ֹʱ�� > Nvl(a.ͣ����ֹʱ��, a.��ʼʱ��)) And
                                 Nvl(a.�Ƿ�����, 0) = 0 And Exists
                            (Select 1
                                  From �ٴ����ﰲ�� M, �ٴ������ N
                                  Where m.Id = a.����id And m.����id = n.Id And n.����ʱ�� Is Not Null)) Loop
                n_�ѹ���   := n_�ѹ��� + Nvl(r_����.��Լ��, 0);
                n_�޺���   := n_�޺��� + Nvl(r_����.��Լ��, r_����.�޺���);
                v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                n_���Ŵ��� := 1;
              End Loop;
              If v_�ϰ�ʱ�� Is Not Null Then
                v_�ϰ�ʱ�� := Substr(v_�ϰ�ʱ��, 2);
              End If;
              If n_���Ŵ��� = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_��ʼʱ�� + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_�޺��� - n_�ѹ��� || '</SYHS>' || '<SBSJ>' || v_�ϰ�ʱ�� || '</SBSJ>' || '<YGS>' || n_�ѹ��� ||
                          '</YGS>' || '</PB>';
              End If;
            End If;
          Else
            --������λ
            If Trunc(d_��ʼʱ�� + n_Daycount) = Trunc(Sysdate) Then
              For r_���� In (Select ID, �ѹ���, �޺���, ��Լ��, �ϰ�ʱ��, �Ƿ���ſ���
                           From �ٴ������¼ A
                           Where �������� = Trunc(d_��ʼʱ�� + n_Daycount) And ҽ��id = n_ҽ��id And
                                 (a.��ʼʱ�� < Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��) Or a.��ֹʱ�� > Nvl(a.ͣ����ֹʱ��, a.��ʼʱ��)) And
                                 Nvl(a.�Ƿ�����, 0) = 0 And Exists
                            (Select 1
                                  From �ٴ����ﰲ�� M, �ٴ������ N
                                  Where m.Id = a.����id And m.����id = n.Id And n.����ʱ�� Is Not Null)) Loop
                Begin
                  Select ���Ʒ�ʽ
                  Into n_��Լģʽ
                  From �ٴ�����Һſ��Ƽ�¼
                  Where ��¼id = r_����.Id And ���� = 1 And ���� = v_������λ And ���� = 1 And Rownum < 2;
                Exception
                  When Others Then
                    n_��Լģʽ := 4;
                End;
                If n_��Լģʽ = 1 Or n_��Լģʽ = 2 Then
                  Select ����
                  Into n_�޺���
                  From �ٴ�����Һſ��Ƽ�¼
                  Where ��¼id = r_����.Id And ���� = 1 And ���� = v_������λ And ���� = 1;
                  If n_��Լģʽ = 1 Then
                    n_�޺��� := n_�޺��� * Nvl(r_����.��Լ��, r_����.�޺���) / 100;
                  End If;
                  Select Count(1)
                  Into n_�ѹ���
                  From ���˹Һż�¼
                  Where �����¼id = r_����.Id And ��¼״̬ = 1 And ������λ = v_������λ;
                  If r_����.�޺��� - r_����.�ѹ��� < n_�޺��� - n_�ѹ��� Then
                    n_ʣ���� := n_ʣ���� + r_����.�޺��� - r_����.�ѹ���;
                  Else
                    n_ʣ���� := n_ʣ���� + n_�޺��� - n_�ѹ���;
                  End If;
                  n_���ѹ��� := n_���ѹ��� + n_�ѹ���;
                  v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                  n_���Ŵ��� := 1;
                End If;
                If n_��Լģʽ = 3 Then
                  If r_����.�Ƿ���ſ��� = 0 Then
                    n_���ѹ��� := n_���ѹ��� + Nvl(r_����.�ѹ���, 0);
                    n_ʣ����   := n_ʣ���� + r_����.�޺��� - Nvl(r_����.�ѹ���, 0);
                    v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                    n_���Ŵ��� := 1;
                  Else
                    For r_���� In (Select ���, ����
                                 From �ٴ�����Һſ��Ƽ�¼
                                 Where ��¼id = r_����.Id And ���� = 1 And ���� = v_������λ And ���� = 1) Loop
                      Begin
                        Select 1
                        Into n_Exists
                        From �ٴ�������ſ���
                        Where ��¼id = r_����.Id And ��� = r_����.��� And Nvl(�Һ�״̬, 0) = 0;
                      Exception
                        When Others Then
                          n_Exists := 0;
                      End;
                      If n_Exists = 1 Then
                        n_ʣ����   := n_ʣ���� + 1;
                        n_���Ŵ��� := 1;
                      End If;
                    End Loop;
                    v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                    n_���ѹ��� := n_���ѹ��� + Nvl(r_����.�ѹ���, 0);
                  End If;
                End If;
                If n_��Լģʽ = 4 Then
                  n_���ѹ��� := n_���ѹ��� + Nvl(r_����.�ѹ���, 0);
                  n_ʣ����   := n_ʣ���� + r_����.�޺��� - Nvl(r_����.�ѹ���, 0);
                  v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                  n_���Ŵ��� := 1;
                End If;
              End Loop;
            
              If v_�ϰ�ʱ�� Is Not Null Then
                v_�ϰ�ʱ�� := Substr(v_�ϰ�ʱ��, 2);
              End If;
              If n_���Ŵ��� = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_��ʼʱ�� + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_ʣ���� || '</SYHS>' || '<SBSJ>' || v_�ϰ�ʱ�� || '</SBSJ>' || '<YGS>' || n_���ѹ��� ||
                          '</YGS>' || '</PB>';
              End If;
            Else
              --�ǵ���
              For r_���� In (Select ID, ��Լ��, �ѹ���, �޺���, ��Լ��, �ϰ�ʱ��, �Ƿ���ſ���
                           From �ٴ������¼ A
                           Where �������� = Trunc(d_��ʼʱ�� + n_Daycount) And ҽ��id = n_ҽ��id And
                                 (a.��ʼʱ�� < Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��) Or a.��ֹʱ�� > Nvl(a.ͣ����ֹʱ��, a.��ʼʱ��)) And
                                 Nvl(a.�Ƿ�����, 0) = 0 And Exists
                            (Select 1
                                  From �ٴ����ﰲ�� M, �ٴ������ N
                                  Where m.Id = a.����id And m.����id = n.Id And n.����ʱ�� Is Not Null)) Loop
                Begin
                  Select ���Ʒ�ʽ
                  Into n_��Լģʽ
                  From �ٴ�����Һſ��Ƽ�¼
                  Where ��¼id = r_����.Id And ���� = 1 And ���� = v_������λ And ���� = 1 And Rownum < 2;
                Exception
                  When Others Then
                    n_��Լģʽ := 4;
                End;
                If n_��Լģʽ = 1 Or n_��Լģʽ = 2 Then
                  Select ����
                  Into n_�޺���
                  From �ٴ�����Һſ��Ƽ�¼
                  Where ��¼id = r_����.Id And ���� = 1 And ���� = v_������λ And ���� = 1;
                  If n_��Լģʽ = 1 Then
                    n_�޺��� := n_�޺��� * Nvl(r_����.��Լ��, r_����.�޺���) / 100;
                  End If;
                  Select Count(1)
                  Into n_�ѹ���
                  From ���˹Һż�¼
                  Where �����¼id = r_����.Id And ��¼״̬ = 1 And ������λ = v_������λ;
                  If Nvl(r_����.��Լ��, r_����.�޺���) - r_����.��Լ�� < n_�޺��� - n_�ѹ��� Then
                    n_ʣ���� := n_ʣ���� + Nvl(r_����.��Լ��, r_����.�޺���) - r_����.��Լ��;
                  Else
                    n_ʣ���� := n_ʣ���� + n_�޺��� - n_�ѹ���;
                  End If;
                  n_���ѹ��� := n_���ѹ��� + n_�ѹ���;
                  v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                  n_���Ŵ��� := 1;
                End If;
                If n_��Լģʽ = 3 Then
                  If r_����.�Ƿ���ſ��� = 0 Then
                    --��ʱ�η���ſ���
                    For r_���� In (Select ���, ����
                                 From �ٴ�����Һſ��Ƽ�¼
                                 Where ��¼id = r_����.Id And ���� = 1 And ���� = v_������λ And ���� = 1) Loop
                      Select Count(1)
                      Into n_Exists
                      From �ٴ�������ſ���
                      Where ԤԼ˳��� Is Not Null And ��¼id = r_����.Id And ��� = r_����.��� And Nvl(�Һ�״̬, 0) <> 0;
                      If r_����.���� - n_Exists > 0 Then
                        n_ʣ����   := n_ʣ���� + r_����.���� - n_Exists;
                        n_���ѹ��� := n_���ѹ��� + n_Exists;
                        n_���Ŵ��� := 1;
                      Else
                        n_���ѹ��� := n_���ѹ��� + n_Exists;
                        n_���Ŵ��� := 1;
                      End If;
                    End Loop;
                    v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                  Else
                    For r_���� In (Select ���
                                 From �ٴ�����Һſ��Ƽ�¼
                                 Where ��¼id = r_����.Id And ���� = 1 And ���� = v_������λ And ���� = 1) Loop
                      Begin
                        Select 1
                        Into n_Exists
                        From �ٴ�������ſ���
                        Where ��¼id = r_����.Id And ��� = r_����.��� And Nvl(�Һ�״̬, 0) = 0;
                      Exception
                        When Others Then
                          n_Exists := 0;
                      End;
                      If n_Exists = 1 Then
                        n_ʣ����   := n_ʣ���� + 1;
                        n_���Ŵ��� := 1;
                      Else
                        n_���ѹ��� := n_���ѹ��� + 1;
                        n_���Ŵ��� := 1;
                      End If;
                    End Loop;
                    v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                  End If;
                End If;
                If n_��Լģʽ = 4 Then
                  n_���ѹ��� := n_���ѹ��� + Nvl(r_����.��Լ��, 0);
                  n_ʣ����   := n_ʣ���� + r_����.��Լ�� - Nvl(r_����.��Լ��, 0);
                  v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                  n_���Ŵ��� := 1;
                End If;
              End Loop;
              If v_�ϰ�ʱ�� Is Not Null Then
                v_�ϰ�ʱ�� := Substr(v_�ϰ�ʱ��, 2);
              End If;
              If n_���Ŵ��� = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_��ʼʱ�� + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_ʣ���� || '</SYHS>' || '<SBSJ>' || v_�ϰ�ʱ�� || '</SBSJ>' || '<YGS>' || n_���ѹ��� ||
                          '</YGS>' || '</PB>';
              End If;
            End If;
          End If;
        End If;
        n_Daycount := n_Daycount + 1;
      End Loop;
    Else
      --ͨ������+ҽ������
      While (n_Daycount < n_��ѯ����) Loop
        If Trunc(d_��ʼʱ�� + n_Daycount) < To_Date(Substr(v_����ʱ��, 1, Instr(v_����ʱ��, ' ') - 1), 'yyyy-mm-dd') Then
          v_�ϰ�ʱ�� := Null;
          n_���ѹ��� := 0;
          n_�ѹ���   := 0;
          n_ʣ����   := 0;
          n_�޺���   := 0;
          n_��Լ��   := 0;
          n_��Լ��   := 0;
          n_���Ŵ��� := 0;
          For r_�Ű� In (Select d_��ʼʱ�� + n_Daycount As ����, a.�Ű�, a.�޺���, a.��Լ��, Nvl(Hz.�ѹ���, 0) As �ѹ���, Nvl(Hz.��Լ��, 0) As ��Լ��,
                              a.����id, a.�ƻ�id, a.����, a.����
                       
                       From (Select Ap.����id, Ap.����, Bm.���� As ��������, Ap.ҽ������, Nvl(Ap.ҽ��id, 0) As ҽ��id, Ry.רҵ����ְ�� As ְ��,
                                     Ap.����, Ap.����id, Ap.�ƻ�id, Ap.�Ű�, Ap.��Ŀid, Fy.���� As ��Ŀ����, Ap.��ſ���, Ap.�޺���, Ap.��Լ��
                              
                              From (Select Ap.����id, Ap.����, Bm.���� As ��������, Ap.ҽ������, Ap.ҽ��id, Ap.����, Ap.Id As ����id, 0 As �ƻ�id,
                                            Ap.��Ŀid, Nvl(Ap.��ſ���, 0) As ��ſ���,
                                            Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', Ap.����, '2', Ap.��һ, '3', Ap.�ܶ�, '4',
                                                    Ap.����, '5', Ap.����, '6', Ap.����, '7', Ap.����, Null) As �Ű�,
                                            Nvl(Xz.��Լ��, Xz.�޺���) As ��Լ��, Nvl(Xz.�޺���, 0) As �޺���
                                     From �ҺŰ��� Ap, ���ű� Bm, �ҺŰ������� Xz
                                     Where Ap.����id = Bm.Id(+) And Ap.ҽ��id = n_ҽ��id And Ap.����id = n_����id And Ap.ͣ������ Is Null And
                                           d_��ʼʱ�� + n_Daycount Between Nvl(Ap.��ʼʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                           Nvl(Ap.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.����id(+) = Ap.Id And
                                           Xz.������Ŀ(+) =
                                           Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4',
                                                  '����', '5', '����', '6', '����', '7', '����', Null) And Not Exists
                                      (Select Rownum
                                            From �ҺŰ���ͣ��״̬ Ty
                                            Where Ty.����id = Ap.Id And d_��ʼʱ�� + n_Daycount Between Ty.��ʼֹͣʱ�� And Ty.����ֹͣʱ��) And
                                           Not Exists (Select Rownum
                                            From �ҺŰ��żƻ� Jh
                                            Where Jh.����id = Ap.Id And Jh.���ʱ�� Is Not Null And
                                                  d_��ʼʱ�� + n_Daycount Between
                                                  Nvl(Jh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                  Nvl(Jh.ʧЧʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')))
                                     Union All
                                     Select Ap.����id, Ap.����, Bm.���� As ��������, Jh.ҽ������, Jh.ҽ��id, Ap.����, Ap.Id As ����id,
                                            Jh.Id As �ƻ�id, Jh.��Ŀid, Nvl(Jh.��ſ���, 0) As ��ſ���,
                                            Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', Jh.����, '2', Jh.��һ, '3', Jh.�ܶ�, '4',
                                                    Jh.����, '5', Jh.����, '6', Jh.����, '7', Jh.����, Null) As �Ű�,
                                            Nvl(Xz.��Լ��, Xz.�޺���) As ��Լ��, Nvl(Xz.�޺���, 0) As �޺���
                                     From �ҺŰ��żƻ� Jh, �ҺŰ��� Ap, ���ű� Bm, �Һżƻ����� Xz
                                     Where Jh.����id = Ap.Id And Ap.����id = Bm.Id(+) And Ap.ͣ������ Is Null And Jh.ҽ��id = n_ҽ��id And
                                           Ap.����id = n_����id And
                                           d_��ʼʱ�� + n_Daycount Between Nvl(Jh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                           Nvl(Jh.ʧЧʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.�ƻ�id(+) = Jh.Id And
                                           Xz.������Ŀ(+) =
                                           Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4',
                                                  '����', '5', '����', '6', '����', '7', '����', Null) And Not Exists
                                      (Select Rownum
                                            From �ҺŰ���ͣ��״̬ Ty
                                            Where Ty.����id = Ap.Id And d_��ʼʱ�� + n_Daycount Between Ty.��ʼֹͣʱ�� And Ty.����ֹͣʱ��) And
                                           (Jh.��Чʱ��, Jh.����id) =
                                           (Select Max(Sxjh.��Чʱ��) As ��Чʱ��, ����id
                                            From �ҺŰ��żƻ� Sxjh
                                            Where Sxjh.���ʱ�� Is Not Null And
                                                  d_��ʼʱ�� + n_Daycount Between
                                                  Nvl(Sxjh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                  Nvl(Sxjh.ʧЧʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Sxjh.����id = Jh.����id
                                            Group By Sxjh.����id)) Ap, ���ű� Bm, ��Ա�� Ry, �շ���ĿĿ¼ Fy
                              Where Ap.����id = Bm.Id(+) And Ap.ҽ��id = Ry.Id(+) And Ap.��Ŀid = Fy.Id And �Ű� Is Not Null) A,
                            ���˹ҺŻ��� Hz, �շѼ�Ŀ B
                       Where a.���� = Hz.����(+) And Hz.����(+) = Trunc(d_��ʼʱ�� + n_Daycount) And a.��Ŀid = b.�շ�ϸĿid And
                             b.��ֹ���� > d_��ʼʱ�� + n_Daycount And b.ִ������ <= d_��ʼʱ�� + n_Daycount
                       
                       ) Loop
            If v_���� Is Null Or Instr(',' || v_���� || ',', ',' || r_�Ű�.���� || ',') > 0 Then
              v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_�Ű�.�Ű�;
              n_���ѹ��� := n_���ѹ��� + r_�Ű�.�ѹ���;
              n_�ѹ���   := r_�Ű�.�ѹ���;
              n_�޺���   := r_�Ű�.�޺���;
              n_��Լ��   := r_�Ű�.��Լ��;
              n_��Լ��   := r_�Ű�.��Լ��;
              n_����id   := Nvl(r_�Ű�.����id, 0);
              n_�ƻ�id   := Nvl(r_�Ű�.�ƻ�id, 0);
              v_����     := r_�Ű�.����;
              n_���Ŵ��� := 1;
            
              If v_�ϰ�ʱ�� Is Not Null Then
                If v_������λ Is Not Null Then
                  If n_�ƻ�id <> 0 Then
                    Begin
                      Select 1
                      Into n_��Լ����
                      From ������λ�ƻ�����
                      Where �ƻ�id = n_�ƻ�id And ������λ = v_������λ And Rownum < 2;
                    Exception
                      When Others Then
                        n_��Լ���� := 0;
                    End;
                  Else
                    Begin
                      Select 1
                      Into n_��Լ����
                      From ������λ���ſ���
                      Where ����id = n_����id And ������λ = v_������λ And Rownum < 2;
                    Exception
                      When Others Then
                        n_��Լ���� := 0;
                    End;
                  End If;
                End If;
              
                If v_������λ Is Null Or Nvl(n_��Լ����, 0) = 0 Then
                  If n_�ƻ�id <> 0 Then
                    Begin
                      Select Sum(����)
                      Into n_������λ����
                      From ������λ�ƻ�����
                      Where �ƻ�id = n_�ƻ�id And
                            ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                          '5', '����', '6', '����', '7', '����', Null);
                    Exception
                      When Others Then
                        n_������λ���� := 0;
                    End;
                  Else
                    Begin
                      Select Sum(����)
                      Into n_������λ����
                      From ������λ���ſ���
                      Where ����id = n_����id And
                            ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                          '5', '����', '6', '����', '7', '����', Null);
                    Exception
                      When Others Then
                        n_������λ���� := 0;
                    End;
                  End If;
                  Begin
                    Select Count(1)
                    Into n_��Լ�ѹ���
                    From ���˹Һż�¼
                    Where �ű� = v_���� And ��¼״̬ = 1 And ������λ Is Not Null And ����ʱ�� Between Trunc(d_��ʼʱ�� + n_Daycount) And
                          Trunc(d_��ʼʱ�� + n_Daycount + 1) - 1 / 60 / 60 / 24;
                  Exception
                    When Others Then
                      n_��Լ�ѹ��� := 0;
                  End;
                  If n_������λ���� = 0 Then
                    n_������λ���� := Null;
                  End If;
                  If Trunc(d_��ʼʱ�� + n_Daycount) > Trunc(Sysdate) Then
                    n_ʣ���� := Nvl(n_ʣ����, 0) + n_��Լ�� - n_��Լ�� - Nvl(n_������λ����, n_��Լ�ѹ���) + Nvl(n_��Լ�ѹ���, 0);
                  Else
                    n_ʣ���� := Nvl(n_ʣ����, 0) + n_�޺��� - n_�ѹ��� - Nvl(n_������λ����, n_��Լ�ѹ���) + Nvl(n_��Լ�ѹ���, 0);
                  End If;
                Else
                  --��Լ��λ
                  If n_�ƻ�id <> 0 Then
                    Begin
                      Select 1
                      Into n_����
                      From ������λ�ƻ�����
                      Where �ƻ�id = n_�ƻ�id And
                            ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                          '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ And ���� = 0 And
                            Rownum < 2;
                    Exception
                      When Others Then
                        n_���� := 0;
                    End;
                  Else
                    Begin
                      Select 1
                      Into n_����
                      From ������λ���ſ���
                      Where ����id = n_����id And
                            ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                          '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ And ���� = 0 And
                            Rownum < 2;
                    Exception
                      When Others Then
                        n_���� := 0;
                    End;
                  End If;
                  If Nvl(n_����, 0) = 0 Then
                    Begin
                      Select Count(1)
                      Into n_��Լ�ѹ���
                      From ���˹Һż�¼
                      Where �ű� = v_���� And ��¼״̬ = 1 And ������λ = v_������λ And ����ʱ�� Between Trunc(d_��ʼʱ�� + n_Daycount) And
                            Trunc(d_��ʼʱ�� + n_Daycount + 1) - 1 / 60 / 60 / 24;
                    Exception
                      When Others Then
                        n_��Լ�ѹ��� := 0;
                    End;
                    If n_�ƻ�id <> 0 Then
                      Begin
                        Select Sum(����)
                        Into n_������λ����
                        From ������λ�ƻ�����
                        Where �ƻ�id = n_�ƻ�id And
                              ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4',
                                            '����', '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ;
                      Exception
                        When Others Then
                          n_������λ���� := 0;
                      End;
                    Else
                      Begin
                        Select Sum(����)
                        Into n_������λ����
                        From ������λ���ſ���
                        Where ����id = n_����id And
                              ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4',
                                            '����', '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ;
                      Exception
                        When Others Then
                          n_������λ���� := 0;
                      End;
                    End If;
                    If n_������λ���� = 0 Then
                      n_������λ���� := Null;
                    End If;
                    n_ʣ���� := Nvl(n_ʣ����, 0) + Nvl(n_������λ����, n_��Լ�ѹ���) - Nvl(n_��Լ�ѹ���, 0);
                  
                  End If;
                End If;
              End If;
              n_������λ���� := 0;
              n_��Լ����     := 0;
              n_����         := 0;
            End If;
          End Loop;
          v_�ϰ�ʱ�� := Substr(v_�ϰ�ʱ��, 2);
          If n_���Ŵ��� = 1 Then
            v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_��ʼʱ�� + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                      '<SYHS>' || n_ʣ���� || '</SYHS>' || '<SBSJ>' || v_�ϰ�ʱ�� || '</SBSJ>' || '<YGS>' || n_���ѹ��� ||
                      '</YGS>' || '</PB>';
          End If;
        Else
          n_���Ŵ��� := 0;
          v_�ϰ�ʱ�� := Null;
          n_���ѹ��� := 0;
          n_�ѹ���   := 0;
          n_ʣ����   := 0;
          n_�޺���   := 0;
          n_��Լ��   := 0;
          n_��Լ��   := 0;
          If v_������λ Is Null Then
            --�Ǻ�����λ
            If Trunc(d_��ʼʱ�� + n_Daycount) = Trunc(Sysdate) Then
              --����Һ�
              For r_���� In (Select �ѹ���, �޺���, �ϰ�ʱ��
                           From �ٴ������¼ A
                           Where �������� = Trunc(d_��ʼʱ�� + n_Daycount) And ҽ��id = n_ҽ��id And ����id = n_����id And
                                 (a.��ʼʱ�� < Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��) Or a.��ֹʱ�� > Nvl(a.ͣ����ֹʱ��, a.��ʼʱ��)) And
                                 Nvl(a.�Ƿ�����, 0) = 0 And Exists
                            (Select 1
                                  From �ٴ����ﰲ�� M, �ٴ������ N
                                  Where m.Id = a.����id And m.����id = n.Id And n.����ʱ�� Is Not Null)) Loop
                n_�ѹ���   := n_�ѹ��� + Nvl(r_����.�ѹ���, 0);
                n_�޺���   := n_�޺��� + r_����.�޺���;
                v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                n_���Ŵ��� := 1;
              End Loop;
              If v_�ϰ�ʱ�� Is Not Null Then
                v_�ϰ�ʱ�� := Substr(v_�ϰ�ʱ��, 2);
              End If;
              If n_���Ŵ��� = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_��ʼʱ�� + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_�޺��� - n_�ѹ��� || '</SYHS>' || '<SBSJ>' || v_�ϰ�ʱ�� || '</SBSJ>' || '<YGS>' || n_�ѹ��� ||
                          '</YGS>' || '</PB>';
              End If;
            Else
              --ԤԼ�Һ�
              For r_���� In (Select ��Լ��, �޺���, ��Լ��, �ϰ�ʱ��
                           From �ٴ������¼ A
                           Where �������� = Trunc(d_��ʼʱ�� + n_Daycount) And ҽ��id = n_ҽ��id And ����id = n_����id And
                                 (a.��ʼʱ�� < Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��) Or a.��ֹʱ�� > Nvl(a.ͣ����ֹʱ��, a.��ʼʱ��)) And
                                 Nvl(a.�Ƿ�����, 0) = 0 And Exists
                            (Select 1
                                  From �ٴ����ﰲ�� M, �ٴ������ N
                                  Where m.Id = a.����id And m.����id = n.Id And n.����ʱ�� Is Not Null)) Loop
                n_�ѹ���   := n_�ѹ��� + Nvl(r_����.��Լ��, 0);
                n_�޺���   := n_�޺��� + Nvl(r_����.��Լ��, r_����.�޺���);
                v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                n_���Ŵ��� := 1;
              End Loop;
              If v_�ϰ�ʱ�� Is Not Null Then
                v_�ϰ�ʱ�� := Substr(v_�ϰ�ʱ��, 2);
              End If;
              If n_���Ŵ��� = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_��ʼʱ�� + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_�޺��� - n_�ѹ��� || '</SYHS>' || '<SBSJ>' || v_�ϰ�ʱ�� || '</SBSJ>' || '<YGS>' || n_�ѹ��� ||
                          '</YGS>' || '</PB>';
              End If;
            End If;
          Else
            --������λ
            If Trunc(d_��ʼʱ�� + n_Daycount) = Trunc(Sysdate) Then
              For r_���� In (Select ID, �ѹ���, �޺���, ��Լ��, �ϰ�ʱ��, �Ƿ���ſ���
                           From �ٴ������¼ A
                           Where �������� = Trunc(d_��ʼʱ�� + n_Daycount) And ҽ��id = n_ҽ��id And ����id = n_����id And
                                 (a.��ʼʱ�� < Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��) Or a.��ֹʱ�� > Nvl(a.ͣ����ֹʱ��, a.��ʼʱ��)) And
                                 Nvl(a.�Ƿ�����, 0) = 0 And Exists
                            (Select 1
                                  From �ٴ����ﰲ�� M, �ٴ������ N
                                  Where m.Id = a.����id And m.����id = n.Id And n.����ʱ�� Is Not Null)) Loop
                Begin
                  Select ���Ʒ�ʽ
                  Into n_��Լģʽ
                  From �ٴ�����Һſ��Ƽ�¼
                  Where ��¼id = r_����.Id And ���� = 1 And ���� = v_������λ And ���� = 1 And Rownum < 2;
                Exception
                  When Others Then
                    n_��Լģʽ := 4;
                End;
                If n_��Լģʽ = 1 Or n_��Լģʽ = 2 Then
                  Select ����
                  Into n_�޺���
                  From �ٴ�����Һſ��Ƽ�¼
                  Where ��¼id = r_����.Id And ���� = 1 And ���� = v_������λ And ���� = 1;
                  If n_��Լģʽ = 1 Then
                    n_�޺��� := n_�޺��� * Nvl(r_����.��Լ��, r_����.�޺���) / 100;
                  End If;
                  Select Count(1)
                  Into n_�ѹ���
                  From ���˹Һż�¼
                  Where �����¼id = r_����.Id And ��¼״̬ = 1 And ������λ = v_������λ;
                  If r_����.�޺��� - r_����.�ѹ��� < n_�޺��� - n_�ѹ��� Then
                    n_ʣ���� := n_ʣ���� + r_����.�޺��� - r_����.�ѹ���;
                  Else
                    n_ʣ���� := n_ʣ���� + n_�޺��� - n_�ѹ���;
                  End If;
                  n_���ѹ��� := n_���ѹ��� + n_�ѹ���;
                  v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                  n_���Ŵ��� := 1;
                End If;
                If n_��Լģʽ = 3 Then
                  If r_����.�Ƿ���ſ��� = 0 Then
                    n_���ѹ��� := n_���ѹ��� + Nvl(r_����.�ѹ���, 0);
                    n_ʣ����   := n_ʣ���� + r_����.�޺��� - Nvl(r_����.�ѹ���, 0);
                    v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                    n_���Ŵ��� := 1;
                  Else
                    For r_���� In (Select ���, ����
                                 From �ٴ�����Һſ��Ƽ�¼
                                 Where ��¼id = r_����.Id And ���� = 1 And ���� = v_������λ And ���� = 1) Loop
                      Begin
                        Select 1
                        Into n_Exists
                        From �ٴ�������ſ���
                        Where ��¼id = r_����.Id And ��� = r_����.��� And Nvl(�Һ�״̬, 0) = 0;
                      Exception
                        When Others Then
                          n_Exists := 0;
                      End;
                      If n_Exists = 1 Then
                        n_ʣ����   := n_ʣ���� + 1;
                        n_���Ŵ��� := 1;
                      End If;
                    End Loop;
                    v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                    n_���ѹ��� := n_���ѹ��� + Nvl(r_����.�ѹ���, 0);
                  End If;
                End If;
                If n_��Լģʽ = 4 Then
                  n_���ѹ��� := n_���ѹ��� + Nvl(r_����.�ѹ���, 0);
                  n_ʣ����   := n_ʣ���� + r_����.�޺��� - Nvl(r_����.�ѹ���, 0);
                  v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                  n_���Ŵ��� := 1;
                End If;
              End Loop;
            
              If v_�ϰ�ʱ�� Is Not Null Then
                v_�ϰ�ʱ�� := Substr(v_�ϰ�ʱ��, 2);
              End If;
              If n_���Ŵ��� = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_��ʼʱ�� + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_ʣ���� || '</SYHS>' || '<SBSJ>' || v_�ϰ�ʱ�� || '</SBSJ>' || '<YGS>' || n_���ѹ��� ||
                          '</YGS>' || '</PB>';
              End If;
            Else
              --�ǵ���
              For r_���� In (Select ID, ��Լ��, �ѹ���, �޺���, ��Լ��, �ϰ�ʱ��, �Ƿ���ſ���
                           From �ٴ������¼ A
                           Where �������� = Trunc(d_��ʼʱ�� + n_Daycount) And ҽ��id = n_ҽ��id And ����id = n_����id And
                                 (a.��ʼʱ�� < Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��) Or a.��ֹʱ�� > Nvl(a.ͣ����ֹʱ��, a.��ʼʱ��)) And
                                 Nvl(a.�Ƿ�����, 0) = 0 And Exists
                            (Select 1
                                  From �ٴ����ﰲ�� M, �ٴ������ N
                                  Where m.Id = a.����id And m.����id = n.Id And n.����ʱ�� Is Not Null)) Loop
                Begin
                  Select ���Ʒ�ʽ
                  Into n_��Լģʽ
                  From �ٴ�����Һſ��Ƽ�¼
                  Where ��¼id = r_����.Id And ���� = 1 And ���� = v_������λ And ���� = 1 And Rownum < 2;
                Exception
                  When Others Then
                    n_��Լģʽ := 4;
                End;
                If n_��Լģʽ = 1 Or n_��Լģʽ = 2 Then
                  Select ����
                  Into n_�޺���
                  From �ٴ�����Һſ��Ƽ�¼
                  Where ��¼id = r_����.Id And ���� = 1 And ���� = v_������λ And ���� = 1;
                  If n_��Լģʽ = 1 Then
                    n_�޺��� := n_�޺��� * Nvl(r_����.��Լ��, r_����.�޺���) / 100;
                  End If;
                  Select Count(1)
                  Into n_�ѹ���
                  From ���˹Һż�¼
                  Where �����¼id = r_����.Id And ��¼״̬ = 1 And ������λ = v_������λ;
                  If Nvl(r_����.��Լ��, r_����.�޺���) - r_����.��Լ�� < n_�޺��� - n_�ѹ��� Then
                    n_ʣ���� := n_ʣ���� + Nvl(r_����.��Լ��, r_����.�޺���) - r_����.��Լ��;
                  Else
                    n_ʣ���� := n_ʣ���� + n_�޺��� - n_�ѹ���;
                  End If;
                  n_���ѹ��� := n_���ѹ��� + n_�ѹ���;
                  v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                  n_���Ŵ��� := 1;
                End If;
                If n_��Լģʽ = 3 Then
                  If r_����.�Ƿ���ſ��� = 0 Then
                    --��ʱ�η���ſ���
                    For r_���� In (Select ���, ����
                                 From �ٴ�����Һſ��Ƽ�¼
                                 Where ��¼id = r_����.Id And ���� = 1 And ���� = v_������λ And ���� = 1) Loop
                      Select Count(1)
                      Into n_Exists
                      From �ٴ�������ſ���
                      Where ԤԼ˳��� Is Not Null And ��¼id = r_����.Id And ��� = r_����.��� And Nvl(�Һ�״̬, 0) <> 0;
                      If r_����.���� - n_Exists > 0 Then
                        n_ʣ����   := n_ʣ���� + r_����.���� - n_Exists;
                        n_���ѹ��� := n_���ѹ��� + n_Exists;
                        n_���Ŵ��� := 1;
                      Else
                        n_���ѹ��� := n_���ѹ��� + n_Exists;
                        n_���Ŵ��� := 1;
                      End If;
                    End Loop;
                    v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                  Else
                    For r_���� In (Select ���
                                 From �ٴ�����Һſ��Ƽ�¼
                                 Where ��¼id = r_����.Id And ���� = 1 And ���� = v_������λ And ���� = 1) Loop
                      Begin
                        Select 1
                        Into n_Exists
                        From �ٴ�������ſ���
                        Where ��¼id = r_����.Id And ��� = r_����.��� And Nvl(�Һ�״̬, 0) = 0;
                      Exception
                        When Others Then
                          n_Exists := 0;
                      End;
                      If n_Exists = 1 Then
                        n_ʣ����   := n_ʣ���� + 1;
                        n_���Ŵ��� := 1;
                      Else
                        n_���ѹ��� := n_���ѹ��� + 1;
                        n_���Ŵ��� := 1;
                      End If;
                    End Loop;
                    v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                  End If;
                End If;
                If n_��Լģʽ = 4 Then
                  n_���ѹ��� := n_���ѹ��� + Nvl(r_����.��Լ��, 0);
                  n_ʣ����   := n_ʣ���� + r_����.��Լ�� - Nvl(r_����.��Լ��, 0);
                  v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                  n_���Ŵ��� := 1;
                End If;
              End Loop;
              If v_�ϰ�ʱ�� Is Not Null Then
                v_�ϰ�ʱ�� := Substr(v_�ϰ�ʱ��, 2);
              End If;
              If n_���Ŵ��� = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_��ʼʱ�� + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_ʣ���� || '</SYHS>' || '<SBSJ>' || v_�ϰ�ʱ�� || '</SBSJ>' || '<YGS>' || n_���ѹ��� ||
                          '</YGS>' || '</PB>';
              End If;
            End If;
          End If;
        End If;
        n_Daycount := n_Daycount + 1;
      End Loop;
    End If;
    v_Temp := '<PBLIST>' || v_Temp || '</PBLIST>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  End If;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Docarrange;
/



Create Or Replace Procedure Zl_Third_Lockno
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --����:HIS����
  --���:Xml_In:
  --<IN>
  --  <HM>5</HM>           //����
  --  <CZJLID>1</CZJLID>       //�����¼ID,������Ű�ģʽ�´���
  --  <RQ>2013-11-21 09:00</RQ>     //ԤԼʱ��
  --  <CZ>1</CZ>           //����
  --  <HX></HX>          //����
  --  <HZDW>֧����</HZDW>   //������λ
  --  <JQM>������</JQM>        //������
  --</IN>

  --����:Xml_Out
  --<OUTPUT>
  -- <HX>����</HX>          //���Ų������ҳɹ�ʱ����
  -- ������Ϣ  //����ʱ����
  --</ OUTPUT>
  --------------------------------------------------------------------------------------------------
  v_����         �ҺŰ���.����%Type;
  d_����         Date;
  n_��������     Number(3);
  n_��ſ���     Number(3);
  n_����         Number(3);
  n_��ʱ��       Number(3);
  n_�޺���       �ҺŰ�������.�޺���%Type;
  n_����id       �ҺŰ���.Id%Type;
  n_�ƻ�id       �ҺŰ��żƻ�.Id%Type;
  n_����         �Һ����״̬.���%Type;
  v_����         �ҺŰ�������.������Ŀ%Type;
  v_����Ա����   �Һ����״̬.����Ա����%Type;
  v_������       �Һ����״̬.������%Type;
  v_��֤����     �Һ����״̬.����Ա����%Type;
  v_��֤������   �Һ����״̬.������%Type;
  n_״̬         �Һ����״̬.״̬%Type;
  v_������λ     ������λ���ſ���.������λ%Type;
  n_��Լģʽ     Number(3);
  n_���ú�����λ Number(3);
  v_Temp         Varchar2(32767); --��ʱXML
  v_Optemp       Varchar2(300);
  x_Templet      Xmltype; --ģ��XML
  v_Err_Msg      Varchar2(200);
  n_Exists       Number(2);
  n_��¼id       �ٴ������¼.Id%Type;
  n_���         �ٴ�������ſ���.���%Type;
  n_����         �ٴ�������ſ���.����%Type;
  n_˳���       �ٴ�������ſ���.ԤԼ˳���%Type;
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/HM'), To_Date(Extractvalue(Value(A), 'IN/RQ'), 'YYYY-MM-DD hh24:mi:ss'),
         Extractvalue(Value(A), 'IN/CZ'), Extractvalue(Value(A), 'IN/HX'), Extractvalue(Value(A), 'IN/HZDW'),
         Extractvalue(Value(A), 'IN/JQM'), Extractvalue(Value(A), 'IN/CZJLID')
  Into v_����, d_����, n_��������, n_����, v_������λ, v_������, n_��¼id
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  If v_������ Is Null Then
    Select Terminal Into v_������ From V$session Where Audsid = Userenv('sessionid');
  End If;
  v_Optemp := Zl_Identity(1);
  Select Substr(v_Optemp, Instr(v_Optemp, ',') + 1) Into v_Optemp From Dual;
  Select Substr(v_Optemp, Instr(v_Optemp, ',') + 1) Into v_����Ա���� From Dual;

  If n_��¼id Is Null Then
    If n_�������� = 0 Then
      --����
      Begin
        Select 1
        Into n_Exists
        From �Һ����״̬
        Where ������ = v_������ And ����Ա���� = v_����Ա���� And ״̬ = 5 And ��� = n_���� And Trunc(����) = Trunc(d_����) And ���� = v_���� And
              Rownum < 2;
      Exception
        When Others Then
          n_Exists := 0;
      End;
      If n_Exists = 1 Then
        Delete �Һ����״̬
        Where ������ = v_������ And ����Ա���� = v_����Ա���� And ״̬ = 5 And ��� = n_���� And Trunc(����) = Trunc(d_����) And ���� = v_����;
        v_Temp := '<HX>' || n_���� || '</HX>';
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
      Else
        v_Temp := 'û�з�����Ҫ���������';
        Raise Err_Item;
      End If;
    End If;
  
    If n_�������� = 1 Then
      --����
      Select Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����',
                     Null)
      Into v_����
      From Dual;
      Begin
        Select ��ſ���, ID
        Into n_��ſ���, n_�ƻ�id
        From (Select ��ſ���, ID
               From �ҺŰ��żƻ�
               Where ���� = v_���� And d_���� Between Nvl(��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                     Nvl(ʧЧʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And ���ʱ�� Is Not Null
               Order By ��Чʱ�� Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          Select ��ſ���, ID Into n_��ſ���, n_����id From �ҺŰ��� Where ���� = v_����;
      End;
      If n_��ſ��� = 1 Then
        If Nvl(n_�ƻ�id, 0) <> 0 Then
          Begin
            Select 1 Into n_��ʱ�� From �Һżƻ�ʱ�� Where ���� = v_���� And �ƻ�id = n_�ƻ�id And Rownum < 2;
          Exception
            When Others Then
              n_��ʱ�� := 0;
          End;
          Begin
            Select 1
            Into n_���ú�����λ
            From ������λ�ƻ�����
            Where ������Ŀ = v_���� And �ƻ�id = n_�ƻ�id And ������λ = v_������λ And Rownum < 2;
          Exception
            When Others Then
              n_���ú�����λ := 0;
          End;
          Begin
            Select 1, a.״̬, a.����Ա����, a.������
            Into n_����, n_״̬, v_��֤����, v_��֤������
            From �Һ����״̬ A, �Һżƻ�ʱ�� B
            Where a.���� = v_���� And Trunc(a.����) = Trunc(d_����) And a.��� = b.��� And b.�ƻ�id = n_�ƻ�id And b.���� = v_���� And
                  To_Char(b.��ʼʱ��, 'hh24:mi') = To_Char(d_����, 'hh24:mi') And Rownum < 2;
          Exception
            When Others Then
              n_���� := 0;
          End;
          If n_���� = 1 Then
            If n_״̬ = 5 And v_��֤���� = v_����Ա���� And v_������ = v_��֤������ Then
              Null;
            Else
              --����ʱ�������Ѿ���ʹ��
              v_Temp := '����ʱ��' || d_���� || '������ѱ�ʹ��';
              Raise Err_Item;
            End If;
          Else
            If n_��ʱ�� = 1 Then
              Begin
                Select ���
                Into n_����
                From �Һżƻ�ʱ��
                Where �ƻ�id = n_�ƻ�id And ���� = v_���� And To_Char(��ʼʱ��, 'hh24:mi') = To_Char(d_����, 'hh24:mi') And
                      Rownum < 2;
              Exception
                When Others Then
                  Select Max(���) + 1
                  Into n_����
                  From (Select Distinct ���
                         From �Һżƻ�ʱ��
                         Where �ƻ�id = n_�ƻ�id And ���� = v_����
                         Union
                         Select Distinct ��� From �Һ����״̬ Where ���� = v_���� And Trunc(����) = Trunc(d_����));
                
              End;
              Begin
                Select 1 Into n_���� From �Һ����״̬ Where ���� = v_���� And ���� = d_���� And ��� = n_����;
              Exception
                When Others Then
                  Insert Into �Һ����״̬
                    (����, ����, ���, ״̬, ����Ա����, ��ע, �Ǽ�ʱ��, ������)
                  Values
                    (v_����, d_����, n_����, 5, v_����Ա����, '�ƶ�����', Sysdate, v_������);
              End;
              v_Temp := '<HX>' || n_���� || '</HX>';
              Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
            Else
              If v_������λ Is Null Or n_���ú�����λ = 0 Then
                If Trunc(d_����) = Trunc(Sysdate) Then
                  n_���� := 1;
                  Select �޺��� Into n_�޺��� From �Һżƻ����� Where �ƻ�id = n_�ƻ�id And ������Ŀ = v_����;
                  For r_��� In (Select ���, ״̬, ����Ա����, ������
                               From �Һ����״̬
                               Where ���� = v_���� And Trunc(����) = Trunc(d_����)
                               Order By ���) Loop
                    Exit When r_���.״̬ = 5 And r_���.����Ա���� = v_����Ա���� And r_���.������ = v_������;
                    If r_���.��� = n_���� Then
                      n_���� := n_���� + 1;
                    End If;
                  End Loop;
                  If n_���� > n_�޺��� Then
                    v_Temp := '����ű�' || v_���� || '����������ѱ�����';
                    Raise Err_Item;
                  Else
                    Begin
                      Select 1
                      Into n_����
                      From �Һ����״̬
                      Where ���� = v_���� And Trunc(����) = Trunc(d_����) And ��� = n_����;
                    Exception
                      When Others Then
                        Insert Into �Һ����״̬
                          (����, ����, ���, ״̬, ����Ա����, ��ע, �Ǽ�ʱ��, ������)
                        Values
                          (v_����, Trunc(d_����), n_����, 5, v_����Ա����, '�ƶ�����', Sysdate, v_������);
                    End;
                    v_Temp := '<HX>' || n_���� || '</HX>';
                    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
                  End If;
                Else
                  n_���� := 1;
                  Select �޺��� Into n_�޺��� From �Һżƻ����� Where �ƻ�id = n_�ƻ�id And ������Ŀ = v_����;
                  For r_��� In (Select ���, ״̬, ����Ա����, ������
                               From �Һ����״̬
                               Where ���� = v_���� And Trunc(����) = Trunc(d_����)
                               
                               Union
                               Select ���, Null, Null, Null
                               From ������λ�ƻ�����
                               Where �ƻ�id = n_�ƻ�id And ������Ŀ = v_���� And ���� <> 0
                               Order By ���) Loop
                    Exit When r_���.״̬ = 5 And r_���.����Ա���� = v_����Ա���� And r_���.������ = v_������;
                    If r_���.��� = n_���� Then
                      n_���� := n_���� + 1;
                    End If;
                  End Loop;
                  If n_���� > n_�޺��� Then
                    v_Temp := '����ű�' || v_���� || '����������ѱ�����';
                    Raise Err_Item;
                  Else
                    Begin
                      Select 1
                      Into n_����
                      From �Һ����״̬
                      Where ���� = v_���� And Trunc(����) = Trunc(d_����) And ��� = n_����;
                    Exception
                      When Others Then
                        Insert Into �Һ����״̬
                          (����, ����, ���, ״̬, ����Ա����, ��ע, �Ǽ�ʱ��, ������)
                        Values
                          (v_����, Trunc(d_����), n_����, 5, v_����Ա����, '�ƶ�����', Sysdate, v_������);
                    End;
                    v_Temp := '<HX>' || n_���� || '</HX>';
                    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
                  End If;
                End If;
              Else
                Select Count(1)
                Into n_��Լģʽ
                From ������λ�ƻ�����
                Where ��� = 0 And �ƻ�id = n_�ƻ�id And ������λ = v_������λ And ������Ŀ = v_���� And ���� <> 0;
                If n_��Լģʽ = 0 Then
                  Begin
                    Select ���
                    Into n_����
                    From (Select ���
                           From ������λ�ƻ����� A
                           Where �ƻ�id = n_�ƻ�id And ������λ = v_������λ And ������Ŀ = v_���� And ���� <> 0 And
                                 (Not Exists
                                  (Select 1
                                   From �Һ����״̬
                                   Where ���� = v_���� And ��� = a.��� And Trunc(����) = Trunc(d_����) And ״̬ <> 5) Or Exists
                                  (Select 1
                                   From �Һ����״̬
                                   Where ���� = v_���� And ��� = a.��� And Trunc(����) = Trunc(d_����) And ״̬ = 5 And ����Ա���� = v_����Ա���� And
                                         ������ = v_������))
                           Order By ���)
                    Where Rownum < 2;
                  Exception
                    When Others Then
                      n_���� := 0;
                  End;
                  If n_���� = 0 Then
                    v_Temp := '����ű�' || v_���� || '����������ѱ�����';
                    Raise Err_Item;
                  Else
                    Begin
                      Select 1
                      Into n_����
                      From �Һ����״̬
                      Where ���� = v_���� And Trunc(����) = Trunc(d_����) And ��� = n_����;
                    Exception
                      When Others Then
                        Insert Into �Һ����״̬
                          (����, ����, ���, ״̬, ����Ա����, ��ע, �Ǽ�ʱ��, ������)
                        Values
                          (v_����, Trunc(d_����), n_����, 5, v_����Ա����, '�ƶ�����', Sysdate, v_������);
                    End;
                    v_Temp := '<HX>' || n_���� || '</HX>';
                    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
                  End If;
                Else
                  n_���� := 1;
                  Select �޺��� Into n_�޺��� From �Һżƻ����� Where �ƻ�id = n_�ƻ�id And ������Ŀ = v_����;
                  For r_��� In (Select ���, ״̬, ����Ա����, ������
                               From �Һ����״̬
                               Where ���� = v_���� And Trunc(����) = Trunc(d_����)
                               
                               Union
                               Select ���, Null, Null, Null
                               From ������λ�ƻ�����
                               Where �ƻ�id = n_�ƻ�id And ������Ŀ = v_���� And ���� <> 0
                               Order By ���) Loop
                    Exit When r_���.״̬ = 5 And r_���.����Ա���� = v_����Ա���� And r_���.������ = v_������;
                    If r_���.��� = n_���� Then
                      n_���� := n_���� + 1;
                    End If;
                  End Loop;
                  If n_���� > n_�޺��� Then
                    v_Temp := '����ű�' || v_���� || '����������ѱ�����';
                    Raise Err_Item;
                  Else
                    Begin
                      Select 1
                      Into n_����
                      From �Һ����״̬
                      Where ���� = v_���� And Trunc(����) = Trunc(d_����) And ��� = n_����;
                    Exception
                      When Others Then
                        Insert Into �Һ����״̬
                          (����, ����, ���, ״̬, ����Ա����, ��ע, �Ǽ�ʱ��, ������)
                        Values
                          (v_����, Trunc(d_����), n_����, 5, v_����Ա����, '�ƶ�����', Sysdate, v_������);
                    End;
                    v_Temp := '<HX>' || n_���� || '</HX>';
                    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
                  End If;
                End If;
              End If;
            End If;
          End If;
        Else
          Begin
            Select 1 Into n_��ʱ�� From �ҺŰ���ʱ�� Where ���� = v_���� And ����id = n_����id And Rownum < 2;
          Exception
            When Others Then
              n_��ʱ�� := 0;
          End;
          Begin
            Select 1
            Into n_���ú�����λ
            From ������λ���ſ���
            Where ������Ŀ = v_���� And ����id = n_����id And ������λ = v_������λ And Rownum < 2;
          Exception
            When Others Then
              n_���ú�����λ := 0;
          End;
          Begin
            Select 1, a.״̬, a.����Ա����, a.������
            Into n_����, n_״̬, v_��֤����, v_��֤������
            From �Һ����״̬ A, �ҺŰ���ʱ�� B
            Where a.���� = v_���� And Trunc(a.����) = Trunc(d_����) And a.��� = b.��� And b.����id = n_����id And b.���� = v_���� And
                  To_Char(b.��ʼʱ��, 'hh24:mi') = To_Char(d_����, 'hh24:mi') And Rownum < 2;
          Exception
            When Others Then
              n_���� := 0;
          End;
          If n_���� = 1 Then
            If n_״̬ = 5 And v_��֤���� = v_����Ա���� And v_������ = v_��֤������ Then
              Null;
            Else
              --����ʱ�������Ѿ���ʹ��
              v_Temp := '����ʱ��' || d_���� || '������ѱ�ʹ��';
              Raise Err_Item;
            End If;
          Else
            If n_��ʱ�� = 1 Then
              Begin
                Select ���
                Into n_����
                From �ҺŰ���ʱ��
                Where ����id = n_����id And ���� = v_���� And To_Char(��ʼʱ��, 'hh24:mi') = To_Char(d_����, 'hh24:mi') And
                      Rownum < 2;
              Exception
                When Others Then
                  Select Max(���) + 1
                  Into n_����
                  From (Select Distinct ���
                         From �ҺŰ���ʱ��
                         Where ����id = n_����id And ���� = v_����
                         Union
                         Select Distinct ��� From �Һ����״̬ Where ���� = v_���� And Trunc(����) = Trunc(d_����));
              End;
              Begin
                Select 1 Into n_���� From �Һ����״̬ Where ���� = v_���� And ���� = d_���� And ��� = n_����;
              Exception
                When Others Then
                  Insert Into �Һ����״̬
                    (����, ����, ���, ״̬, ����Ա����, ��ע, �Ǽ�ʱ��, ������)
                  Values
                    (v_����, d_����, n_����, 5, v_����Ա����, '�ƶ�����', Sysdate, v_������);
              End;
              v_Temp := '<HX>' || n_���� || '</HX>';
              Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
            Else
              If v_������λ Is Null Or n_���ú�����λ = 0 Then
                If Trunc(d_����) = Trunc(Sysdate) Then
                  n_���� := 1;
                  Select �޺��� Into n_�޺��� From �ҺŰ������� Where ����id = n_����id And ������Ŀ = v_����;
                  For r_��� In (Select ���, ״̬, ����Ա����, ������
                               From �Һ����״̬
                               Where ���� = v_���� And Trunc(����) = Trunc(d_����)
                               Order By ���) Loop
                    Exit When r_���.״̬ = 5 And r_���.����Ա���� = v_����Ա���� And r_���.������ = v_������;
                    If r_���.��� = n_���� Then
                      n_���� := n_���� + 1;
                    End If;
                  End Loop;
                  If n_���� > n_�޺��� Then
                    v_Temp := '����ű�' || v_���� || '����������ѱ�����';
                    Raise Err_Item;
                  Else
                    Begin
                      Select 1
                      Into n_����
                      From �Һ����״̬
                      Where ���� = v_���� And Trunc(����) = Trunc(d_����) And ��� = n_����;
                    Exception
                      When Others Then
                        Insert Into �Һ����״̬
                          (����, ����, ���, ״̬, ����Ա����, ��ע, �Ǽ�ʱ��, ������)
                        Values
                          (v_����, Trunc(d_����), n_����, 5, v_����Ա����, '�ƶ�����', Sysdate, v_������);
                    End;
                    v_Temp := '<HX>' || n_���� || '</HX>';
                    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
                  End If;
                Else
                  n_���� := 1;
                  Select �޺��� Into n_�޺��� From �ҺŰ������� Where ����id = n_����id And ������Ŀ = v_����;
                  For r_��� In (Select ���, ״̬, ����Ա����, ������
                               From �Һ����״̬
                               Where ���� = v_���� And Trunc(����) = Trunc(d_����)
                               
                               Union
                               Select ���, Null, Null, Null
                               From ������λ���ſ���
                               Where ����id = n_����id And ������Ŀ = v_���� And ���� <> 0
                               Order By ���) Loop
                    Exit When r_���.״̬ = 5 And r_���.����Ա���� = v_����Ա���� And r_���.������ = v_������;
                    If r_���.��� = n_���� Then
                      n_���� := n_���� + 1;
                    End If;
                  End Loop;
                  If n_���� > n_�޺��� Then
                    v_Temp := '����ű�' || v_���� || '����������ѱ�����';
                    Raise Err_Item;
                  Else
                    Begin
                      Select 1
                      Into n_����
                      From �Һ����״̬
                      Where ���� = v_���� And Trunc(����) = Trunc(d_����) And ��� = n_����;
                    Exception
                      When Others Then
                        Insert Into �Һ����״̬
                          (����, ����, ���, ״̬, ����Ա����, ��ע, �Ǽ�ʱ��, ������)
                        Values
                          (v_����, Trunc(d_����), n_����, 5, v_����Ա����, '�ƶ�����', Sysdate, v_������);
                    End;
                    v_Temp := '<HX>' || n_���� || '</HX>';
                    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
                  End If;
                End If;
              Else
                Select Count(1)
                Into n_��Լģʽ
                From ������λ���ſ���
                Where ��� = 0 And ����id = n_����id And ������λ = v_������λ And ������Ŀ = v_���� And ���� <> 0;
                If n_��Լģʽ = 0 Then
                  Begin
                    Select ���
                    Into n_����
                    From (Select ���
                           From ������λ���ſ��� A
                           Where ����id = n_����id And ������λ = v_������λ And ������Ŀ = v_���� And ���� <> 0 And
                                 (Not Exists
                                  (Select 1
                                   From �Һ����״̬
                                   Where ���� = v_���� And ��� = a.��� And Trunc(����) = Trunc(d_����) And ״̬ <> 5) Or Exists
                                  (Select 1
                                   From �Һ����״̬
                                   Where ���� = v_���� And ��� = a.��� And Trunc(����) = Trunc(d_����) And ״̬ = 5 And ����Ա���� = v_����Ա���� And
                                         ������ = v_������))
                           Order By ���)
                    Where Rownum < 2;
                  Exception
                    When Others Then
                      n_���� := 0;
                  End;
                  If n_���� = 0 Then
                    v_Temp := '����ű�' || v_���� || '����������ѱ�����';
                    Raise Err_Item;
                  Else
                    Begin
                      Select 1
                      Into n_����
                      From �Һ����״̬
                      Where ���� = v_���� And Trunc(����) = Trunc(d_����) And ��� = n_����;
                    Exception
                      When Others Then
                        Insert Into �Һ����״̬
                          (����, ����, ���, ״̬, ����Ա����, ��ע, �Ǽ�ʱ��, ������)
                        Values
                          (v_����, Trunc(d_����), n_����, 5, v_����Ա����, '�ƶ�����', Sysdate, v_������);
                    End;
                    v_Temp := '<HX>' || n_���� || '</HX>';
                    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
                  End If;
                Else
                  n_���� := 1;
                  Select �޺��� Into n_�޺��� From �ҺŰ������� Where ����id = n_����id And ������Ŀ = v_����;
                  For r_��� In (Select ���, ״̬, ����Ա����, ������
                               From �Һ����״̬
                               Where ���� = v_���� And Trunc(����) = Trunc(d_����)
                               
                               Union
                               Select ���, Null, Null, Null
                               From ������λ���ſ���
                               Where ����id = n_����id And ������Ŀ = v_���� And ���� <> 0
                               Order By ���) Loop
                    Exit When r_���.״̬ = 5 And r_���.����Ա���� = v_����Ա���� And r_���.������ = v_������;
                    If r_���.��� = n_���� Then
                      n_���� := n_���� + 1;
                    End If;
                  End Loop;
                  If n_���� > n_�޺��� Then
                    v_Temp := '����ű�' || v_���� || '����������ѱ�����';
                    Raise Err_Item;
                  Else
                    Begin
                      Select 1
                      Into n_����
                      From �Һ����״̬
                      Where ���� = v_���� And Trunc(����) = Trunc(d_����) And ��� = n_����;
                    Exception
                      When Others Then
                        Insert Into �Һ����״̬
                          (����, ����, ���, ״̬, ����Ա����, ��ע, �Ǽ�ʱ��, ������)
                        Values
                          (v_����, Trunc(d_����), n_����, 5, v_����Ա����, '�ƶ�����', Sysdate, v_������);
                    End;
                    v_Temp := '<HX>' || n_���� || '</HX>';
                    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
                  End If;
                End If;
              End If;
            End If;
          End If;
        End If;
      End If;
    End If;
  Else
    --������Ű�ģʽ
    If n_�������� = 0 Then
      --����
      Begin
        Select 1
        Into n_Exists
        From �ٴ�������ſ���
        Where ����վ���� = v_������ And ����Ա���� = v_����Ա���� And �Һ�״̬ = 5 And (��� = n_���� Or ��ע = n_����) And ��¼id = n_��¼id And
              Rownum < 2;
      Exception
        When Others Then
          n_Exists := 0;
      End;
      If n_Exists = 1 Then
        Update �ٴ�������ſ���
        Set �Һ�״̬ = 0
        Where ����վ���� = v_������ And ����Ա���� = v_����Ա���� And �Һ�״̬ = 5 And (��� = n_���� Or ��ע = n_����) And ��¼id = n_��¼id;
        v_Temp := '<HX>' || n_���� || '</HX>';
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
      Else
        v_Temp := 'û�з�����Ҫ���������';
        Raise Err_Item;
      End If;
    End If;
  
    If n_�������� = 1 Then
      --����
      If n_���� Is Null Then
        Select �Ƿ���ſ���, �Ƿ��ʱ�� Into n_��ſ���, n_��ʱ�� From �ٴ������¼ Where ID = n_��¼id;
        Begin
          Select 1
          Into n_���ú�����λ
          From �ٴ�����Һſ��Ƽ�¼
          Where ��¼id = n_��¼id And ���� = v_������λ And ���� = 1 And ���� = 1 And Rownum < 2;
        Exception
          When Others Then
            n_���ú�����λ := 0;
        End;
        If n_��ſ��� = 1 Then
          Begin
            Select 1, �Һ�״̬, ����Ա����, ����վ����
            Into n_����, n_״̬, v_��֤����, v_��֤������
            From �ٴ�������ſ���
            Where ��¼id = n_��¼id And To_Char(��ʼʱ��, 'hh24:mi') = To_Char(d_����, 'hh24:mi') And Nvl(�Һ�״̬, 0) <> 0 And
                  Rownum < 2;
          Exception
            When Others Then
              n_���� := 0;
          End;
          If n_���� = 1 Then
            If n_״̬ = 5 And v_��֤���� = v_����Ա���� And v_������ = v_��֤������ Then
              Null;
            Else
              --����ʱ�������Ѿ���ʹ��
              v_Temp := '����ʱ��' || d_���� || '������ѱ�ʹ��';
              Raise Err_Item;
            End If;
          Else
            If n_��ʱ�� = 1 Then
              Begin
                Select ���
                Into n_���
                From �ٴ�������ſ���
                Where ��¼id = n_��¼id And To_Char(��ʼʱ��, 'hh24:mi') = To_Char(d_����, 'hh24:mi') And Nvl(�Һ�״̬, 0) = 0 And
                      Rownum < 2;
              Exception
                When Others Then
                  Select Max(���) + 1 Into n_��� From �ٴ�������ſ��� Where ��¼id = n_��¼id;
              End;
              Update �ٴ�������ſ���
              Set �Һ�״̬ = 5, ����ʱ�� = Sysdate, ����Ա���� = v_����Ա����, ����վ���� = v_������
              Where ��¼id = n_��¼id And ��� = n_���;
              If Sql%RowCount = 0 Then
                Insert Into �ٴ�������ſ���
                  (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����ʱ��, ����, ����, ����Ա����, ����վ����)
                  Select ��¼id, n_���, To_Date(To_Char(��ʼʱ��, 'yyyy-mm-dd') || ' ' || To_Char(d_����, 'hh24:mi:ss')),
                         To_Date(To_Char(��ʼʱ��, 'yyyy-mm-dd') || ' ' || To_Char(d_����, 'hh24:mi:ss')), 1, �Ƿ�ԤԼ, 5, Sysdate,
                         v_������λ, 1, v_����Ա����, v_������
                  From �ٴ�������ſ���
                  Where ��¼id = n_��¼id And Rownum < 2;
              End If;
              v_Temp := '<HX>' || n_��� || '</HX>';
              Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
            Else
              If v_������λ Is Null Or n_���ú�����λ = 0 Then
                --�Ǻ�����λ
                Select Min(���) Into n_��� From �ٴ�������ſ��� Where ��¼id = n_��¼id And Nvl(�Һ�״̬, 0) = 0;
                If n_��� = 0 Then
                  Select Max(���) + 1 Into n_��� From �ٴ�������ſ��� Where ��¼id = n_��¼id;
                End If;
                Update �ٴ�������ſ���
                Set �Һ�״̬ = 5, ����ʱ�� = Sysdate, ����Ա���� = v_����Ա����, ����վ���� = v_������
                Where ��¼id = n_��¼id And ��� = n_���;
                If Sql%RowCount = 0 Then
                  Insert Into �ٴ�������ſ���
                    (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����ʱ��, ����, ����, ����Ա����, ����վ����)
                    Select ��¼id, n_���, To_Date(To_Char(��ʼʱ��, 'yyyy-mm-dd') || ' ' || To_Char(d_����, 'hh24:mi:ss')),
                           To_Date(To_Char(��ʼʱ��, 'yyyy-mm-dd') || ' ' || To_Char(d_����, 'hh24:mi:ss')), 1, �Ƿ�ԤԼ, 5,
                           Sysdate, v_������λ, 1, v_����Ա����, v_������
                    From �ٴ�������ſ���
                    Where ��¼id = n_��¼id And Rownum < 2;
                End If;
                v_Temp := '<HX>' || n_��� || '</HX>';
                Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
              Else
                --������λ
                Begin
                  Select ���Ʒ�ʽ
                  Into n_��Լģʽ
                  From �ٴ�����Һſ��Ƽ�¼
                  Where ��¼id = n_��¼id And ���� = 1 And ���� = 1 And ���� = v_������λ And Rownum < 2;
                Exception
                  When Others Then
                    n_��Լģʽ := 4;
                End;
                If n_��Լģʽ = 0 Then
                  v_Temp := '���ű��ֹ�ú�����λԤԼ!';
                  Raise Err_Item;
                End If;
                If n_��Լģʽ = 1 Or n_��Լģʽ = 2 Or n_��Լģʽ = 4 Then
                  Select Min(���) Into n_��� From �ٴ�������ſ��� Where ��¼id = n_��¼id And Nvl(�Һ�״̬, 0) = 0;
                  If n_��� = 0 Then
                    Select Max(���) + 1 Into n_��� From �ٴ�������ſ��� Where ��¼id = n_��¼id;
                  End If;
                End If;
                If n_��Լģʽ = 3 Then
                  Select Min(a.���)
                  Into n_���
                  From �ٴ�������ſ��� A, �ٴ�����Һſ��Ƽ�¼ B
                  Where a.��¼id = n_��¼id And a.��¼id = b.��¼id And b.���� = 1 And b.���� = 1 And b.���� = v_������λ And a.��� = b.��� And
                        Nvl(a.�Һ�״̬, 0) = 0;
                  If n_��� = 0 Then
                    v_Temp := '���ű������λ��ԤԼ����Ѿ�ȫ��ʹ����!';
                    Raise Err_Item;
                  End If;
                End If;
                Update �ٴ�������ſ���
                Set �Һ�״̬ = 5, ����ʱ�� = Sysdate, ����Ա���� = v_����Ա����, ����վ���� = v_������
                Where ��¼id = n_��¼id And ��� = n_���;
                If Sql%RowCount = 0 Then
                  Insert Into �ٴ�������ſ���
                    (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����ʱ��, ����, ����, ����Ա����, ����վ����)
                    Select ��¼id, n_���, To_Date(To_Char(��ʼʱ��, 'yyyy-mm-dd') || ' ' || To_Char(d_����, 'hh24:mi:ss')),
                           To_Date(To_Char(��ʼʱ��, 'yyyy-mm-dd') || ' ' || To_Char(d_����, 'hh24:mi:ss')), 1, �Ƿ�ԤԼ, 5,
                           Sysdate, v_������λ, 1, v_����Ա����, v_������
                    From �ٴ�������ſ���
                    Where ��¼id = n_��¼id And Rownum < 2;
                End If;
                v_Temp := '<HX>' || n_��� || '</HX>';
                Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
              End If;
            End If;
          End If;
        Else
          --����ſ���
          If n_��ʱ�� = 1 Then
            If v_������λ Is Null Or n_���ú�����λ = 0 Then
              Begin
                Select ���, ����
                Into n_����, n_����
                From �ٴ�������ſ���
                Where ��¼id = n_��¼id And ԤԼ˳��� Is Null And ��ʼʱ�� = d_����;
                Select Count(1)
                Into n_Exists
                From �ٴ�������ſ���
                Where ��¼id = n_��¼id And ԤԼ˳��� Is Not Null And ��� = n_���� And Nvl(�Һ�״̬, 0) <> 0;
                If n_Exists >= n_���� Then
                  v_Temp := '���ű��������Ѿ�ȫ��ʹ����!';
                  Raise Err_Item;
                Else
                  Select Min(ԤԼ˳���)
                  Into n_˳���
                  From �ٴ�������ſ���
                  Where ��¼id = n_��¼id And ԤԼ˳��� Is Not Null And ��� = n_���� And Nvl(�Һ�״̬, 0) = 0;
                  If n_˳��� = 0 Then
                    Select Max(ԤԼ˳���) + 1
                    Into n_˳���
                    From �ٴ�������ſ���
                    Where ��¼id = n_��¼id And ԤԼ˳��� Is Not Null And ��� = n_����;
                  End If;
                  Update �ٴ�������ſ���
                  Set �Һ�״̬ = 5, ����ʱ�� = Sysdate, ����Ա���� = v_����Ա����, ����վ���� = v_������
                  Where ��¼id = n_��¼id And ��� = n_��� And ԤԼ˳��� = n_˳���;
                  If Sql%RowCount = 0 Then
                    Insert Into �ٴ�������ſ���
                      (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����ʱ��, ����, ����, ����Ա����, ����վ����, ԤԼ˳���)
                      Select ��¼id, n_���, To_Date(To_Char(��ʼʱ��, 'yyyy-mm-dd') || ' ' || To_Char(d_����, 'hh24:mi:ss')),
                             To_Date(To_Char(��ʼʱ��, 'yyyy-mm-dd') || ' ' || To_Char(d_����, 'hh24:mi:ss')), 1, �Ƿ�ԤԼ, 5,
                             Sysdate, v_������λ, 1, v_����Ա����, v_������, n_˳���
                      From �ٴ�������ſ���
                      Where ��¼id = n_��¼id And Rownum < 2;
                  End If;
                  v_Temp := '<HX>' || n_���� || '</HX>';
                  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
                End If;
              Exception
                When Others Then
                  Null;
              End;
            Else
              --������λ
              Begin
                Select ���Ʒ�ʽ
                Into n_��Լģʽ
                From �ٴ�����Һſ��Ƽ�¼
                Where ��¼id = n_��¼id And ���� = 1 And ���� = 1 And ���� = v_������λ And Rownum < 2;
              Exception
                When Others Then
                  n_��Լģʽ := 4;
              End;
              If n_��Լģʽ = 0 Then
                v_Temp := '���ű��ֹ�ú�����λԤԼ!';
                Raise Err_Item;
              End If;
              If n_��Լģʽ = 1 Or n_��Լģʽ = 2 Or n_��Լģʽ = 4 Then
                Begin
                  Select ���, ����
                  Into n_����, n_����
                  From �ٴ�������ſ���
                  Where ��¼id = n_��¼id And ԤԼ˳��� Is Null And ��ʼʱ�� = d_����;
                  Select Count(1)
                  Into n_Exists
                  From �ٴ�������ſ���
                  Where ��¼id = n_��¼id And ԤԼ˳��� Is Not Null And ��� = n_���� And Nvl(�Һ�״̬, 0) <> 0;
                  If n_Exists >= n_���� Then
                    v_Temp := '���ű��������Ѿ�ȫ��ʹ����!';
                    Raise Err_Item;
                  Else
                    Select Min(ԤԼ˳���)
                    Into n_˳���
                    From �ٴ�������ſ���
                    Where ��¼id = n_��¼id And ԤԼ˳��� Is Not Null And ��� = n_���� And Nvl(�Һ�״̬, 0) = 0;
                    If n_˳��� = 0 Then
                      Select Max(ԤԼ˳���) + 1
                      Into n_˳���
                      From �ٴ�������ſ���
                      Where ��¼id = n_��¼id And ԤԼ˳��� Is Not Null And ��� = n_����;
                    End If;
                    Update �ٴ�������ſ���
                    Set �Һ�״̬ = 5, ����ʱ�� = Sysdate, ����Ա���� = v_����Ա����, ����վ���� = v_������
                    Where ��¼id = n_��¼id And ��� = n_��� And ԤԼ˳��� = n_˳���;
                    If Sql%RowCount = 0 Then
                      Insert Into �ٴ�������ſ���
                        (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����ʱ��, ����, ����, ����Ա����, ����վ����, ԤԼ˳���)
                        Select ��¼id, n_���, To_Date(To_Char(��ʼʱ��, 'yyyy-mm-dd') || ' ' || To_Char(d_����, 'hh24:mi:ss')),
                               To_Date(To_Char(��ʼʱ��, 'yyyy-mm-dd') || ' ' || To_Char(d_����, 'hh24:mi:ss')), 1, �Ƿ�ԤԼ, 5,
                               Sysdate, v_������λ, 1, v_����Ա����, v_������, n_˳���
                        From �ٴ�������ſ���
                        Where ��¼id = n_��¼id And Rownum < 2;
                    End If;
                    v_Temp := '<HX>' || n_���� || '</HX>';
                    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
                  End If;
                Exception
                  When Others Then
                    Null;
                End;
              End If;
              If n_��Լģʽ = 3 Then
                Begin
                  Select ���
                  Into n_����
                  From �ٴ�������ſ���
                  Where ��¼id = n_��¼id And ԤԼ˳��� Is Null And ��ʼʱ�� = d_����;
                  Select ����
                  Into n_����
                  From �ٴ�����Һſ��Ƽ�¼
                  Where ��¼id = n_��¼id And ���� = 1 And ���� = 1 And ���� = v_������λ And ��� = n_����;
                  Select Count(1)
                  Into n_Exists
                  From �ٴ�������ſ���
                  Where ��¼id = n_��¼id And ԤԼ˳��� Is Not Null And ��� = n_���� And Nvl(�Һ�״̬, 0) <> 0;
                  If n_Exists >= n_���� Then
                    v_Temp := '���ű��������Ѿ�ȫ��ʹ����!';
                    Raise Err_Item;
                  Else
                    Select Min(ԤԼ˳���)
                    Into n_˳���
                    From �ٴ�������ſ���
                    Where ��¼id = n_��¼id And ԤԼ˳��� Is Not Null And ��� = n_���� And Nvl(�Һ�״̬, 0) = 0;
                    If n_˳��� = 0 Then
                      Select Max(ԤԼ˳���) + 1
                      Into n_˳���
                      From �ٴ�������ſ���
                      Where ��¼id = n_��¼id And ԤԼ˳��� Is Not Null And ��� = n_����;
                    End If;
                    Update �ٴ�������ſ���
                    Set �Һ�״̬ = 5, ����ʱ�� = Sysdate, ����Ա���� = v_����Ա����, ����վ���� = v_������
                    Where ��¼id = n_��¼id And ��� = n_��� And ԤԼ˳��� = n_˳���;
                    If Sql%RowCount = 0 Then
                      Insert Into �ٴ�������ſ���
                        (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����ʱ��, ����, ����, ����Ա����, ����վ����, ԤԼ˳���)
                        Select ��¼id, n_���, To_Date(To_Char(��ʼʱ��, 'yyyy-mm-dd') || ' ' || To_Char(d_����, 'hh24:mi:ss')),
                               To_Date(To_Char(��ʼʱ��, 'yyyy-mm-dd') || ' ' || To_Char(d_����, 'hh24:mi:ss')), 1, �Ƿ�ԤԼ, 5,
                               Sysdate, v_������λ, 1, v_����Ա����, v_������, n_˳���
                        From �ٴ�������ſ���
                        Where ��¼id = n_��¼id And Rownum < 2;
                    End If;
                    v_Temp := '<HX>' || n_���� || '</HX>';
                    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
                  End If;
                Exception
                  When Others Then
                    Null;
                End;
              End If;
            End If;
          End If;
        End If;
      Else
        n_��� := n_����;
        Update �ٴ�������ſ���
        Set �Һ�״̬ = 5, ����Ա���� = v_����Ա����, ����վ���� = v_������, ����ʱ�� = Sysdate
        Where ��¼id = n_��¼id And ��� = n_���;
        If Sql%RowCount = 0 Then
          Begin
            Insert Into �ٴ�������ſ���
              (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����ʱ��, ����, ����, ����Ա����, ����վ����)
              Select ��¼id, n_���, To_Date(To_Char(��ʼʱ��, 'yyyy-mm-dd') || ' ' || To_Char(d_����, 'hh24:mi:ss')),
                     To_Date(To_Char(��ʼʱ��, 'yyyy-mm-dd') || ' ' || To_Char(d_����, 'hh24:mi:ss')), 1, �Ƿ�ԤԼ, 5, Sysdate,
                     v_������λ, 1, v_����Ա����, v_������
              From �ٴ�������ſ���
              Where ��¼id = n_��¼id And Rownum < 2;
          Exception
            When Others Then
              v_Temp := '�������������ѱ�ʹ��!';
              Raise Err_Item;
          End;
        End If;
      End If;
    End If;
  End If;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Temp || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Lockno;
/

Create Or Replace Procedure Zl_Third_Getnolist
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --����:��ȡ��Դ�б�
  --���:Xml_In:
  --<IN>
  --  <RQ>����</RQ>
  --  <KSID>����ID</KSID>
  --  <YSID>ҽ��ID</YSID>
  --  <YSXM>ҽ������</YSXM>
  --  <HZDW>֧����</HZDW>    //������λ�������˵�ʱ��ֻȡ������λ�ĺ�;Ϊ��ʱ��ֻȡ�Ǻ�����λ�ĺ�
  --</IN>

  --����:Xml_Out
  --<OUTPUT>
  --  <GROUP>
  --    <RQ>����</RQ>
  --    <HBLIST>
  --     <HB>
  --        <CZJLID>1</CZJLID>     //�����¼ID
  --        <HM>235</HM>       //����
  --        <YSID>549</YSID>      //ҽ��ID
  --        <YS>����</YS>       //ҽ������
  --        <KSID>123</KSID>   //����ID
  --        <KSMC>�ڿ�</KSMC>   //��������
  --        <ZC>����ҽʦ</ZC> //ְ��
  --        <XMID>10086<XMID> //�Һ���Ŀ��ID
  --        <XMMC>�Һŷ�</XMMC> //�Һ���Ŀ������
  --        <YGHS>0</YGHS>      //�ѹҺ���
  --        <SYHS>99</SYHS>   //ʣ�����
  --        <PRICE>15</PRICE>      //�۸�
  --        <HL>��ͨ</HL>       //�Һ�����
  --        <HCXH>1</HCXH>    //�Ƿ���ڻ������ʱ��Σ�1-���� 0���߿�-������
  --        <FSD>0</FSD>      //�Ƿ��ʱ��
  --        <FWMC>����</FWMC>     //�ű�ʱ��
  --        <HBTIME>(08:00-17:59)</HBTIME> //�ɹ�ʱ��
  --     <SPANLIST>
  --            <SPAN>
  --                  <SJD/>      //ʱ���
  --                  <SL/>      //����
  --            </SPAN>
  --            ����
  --          </SPANLIST>
  --      </HB>
  --      <HB>
  --      ����
  --      </HB>
  --    </HBLIST>
  --  </GROUP>
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  d_����         Date;
  n_����id       ���˹Һż�¼.ִ�в���id%Type;
  n_ҽ��id       ��Ա��.Id%Type;
  v_ҽ������     ��Ա��.����%Type;
  v_����         �ҺŰ�������.������Ŀ%Type;
  v_ʱ���       Varchar2(100);
  v_������λ     �Һź�����λ.����%Type;
  n_��ʱ��       Number(3);
  n_����ʣ��     Number(5);
  n_�ѹ���       Number(5);
  n_��Լ�ѹ���   Number(5);
  n_�ϼƽ��     �շѼ�Ŀ.�ּ�%Type;
  n_��Լ������   Number(5);
  n_��Լʣ������ Number(5);
  n_���������� Number(5);
  n_��Լģʽ     Number(3); --��Լģʽ:1-��Լ��λ������ģʽ 0-��Լ��λָ�����ģʽ
  n_�Ǻ�Լ       Number(3);
  n_�Ƿ�Ԥ��     Number(3);
  d_�Ӻ�ʱ��     Date;
  d_��ʼʱ��     �ٴ������¼.��ʼʱ��%Type;
  d_��ֹʱ��     �ٴ������¼.��ֹʱ��%Type;
  n_�������     Number(3);
  n_ʱ������     Number(5);
  n_��ſ���     �ٴ������¼.�Ƿ���ſ���%Type;
  n_Ԥ������     Number(5);
  n_����ԤԼ     Number(3);
  n_����         Number(3);
  v_ʣ������     Varchar2(100);
  v_Timetemp     Varchar2(100);
  v_Temp         Varchar2(32767); --��ʱXML
  v_Xmlmain      Clob; --��ʱXML
  c_Xmlmain      Clob; --��ʱXML
  x_Templet      Xmltype; --ģ��XML
  v_Err_Msg      Varchar2(200);
  n_Exists       Number(2);
  v_Sql          Varchar2(20000);
  Type c_Main Is Ref Cursor;
  r_����id   �ҺŰ���.����id%Type;
  r_����     �ҺŰ���.����%Type;
  r_�������� ���ű�.����%Type;
  r_ҽ������ �ҺŰ���.ҽ������%Type;
  r_ҽ��id   �ҺŰ���.ҽ��id%Type;
  r_ְ��     ��Ա��.רҵ����ְ��%Type;
  r_����     �ҺŰ���.����%Type;
  r_����id   �ҺŰ���.Id%Type;
  r_�ƻ�id   �ҺŰ��żƻ�.Id%Type;
  r_�Ű�     �ҺŰ���.����%Type;
  r_��Ŀid   �ҺŰ���.��Ŀid%Type;
  r_��Ŀ���� �շ���ĿĿ¼.����%Type;
  r_��ſ��� �ҺŰ���.��ſ���%Type;
  r_�޺���   �ҺŰ�������.�޺���%Type;
  r_��Լ��   �ҺŰ�������.��Լ��%Type;
  n_ʱ���ѹ� Number(5);
  r_�ѹ���   ���˹ҺŻ���.�ѹ���%Type;
  r_��Լ��   ���˹ҺŻ���.��Լ��%Type;
  r_�ѽ���   ���˹ҺŻ���.�����ѽ���%Type;
  r_�۸�     �շѼ�Ŀ.�ּ�%Type;
  r_��ʱ��   �ٴ������¼.�Ƿ��ʱ��%Type;
  r_��ʼʱ�� �ٴ������¼.��ʼʱ��%Type;
  r_��ֹʱ�� �ٴ������¼.��ֹʱ��%Type;
  r_ԤԼ���� �ٴ������¼.ԤԼ����%Type;
  r_No       c_Main;
  n_Curcount Number(3);
  n_�Һ�ģʽ Number(3);
  v_�Һ�ģʽ Varchar2(500);
  v_����ʱ�� Varchar2(500);

  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select To_Date(Extractvalue(Value(A), 'IN/RQ'), 'YYYY-MM-DD hh24:mi:ss'), Extractvalue(Value(A), 'IN/KSID'),
         Extractvalue(Value(A), 'IN/YSID'), Extractvalue(Value(A), 'IN/YSXM'), Extractvalue(Value(A), 'IN/HZDW')
  Into d_����, n_����id, n_ҽ��id, v_ҽ������, v_������λ
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  v_�Һ�ģʽ := zl_GetSysParameter('�Һ��Ű�ģʽ');
  n_�Һ�ģʽ := To_Number(Substr(v_�Һ�ģʽ, 1, 1));
  If n_�Һ�ģʽ = 1 Then
    Begin
      v_����ʱ�� := Substr(v_�Һ�ģʽ, 3);
    Exception
      When Others Then
        Null;
    End;
  End If;
  --���ڽڵ�Ϊ�յ����
  If d_���� Is Null Then
    d_���� := Trunc(Sysdate);
  End If;

  If n_�Һ�ģʽ = 0 Then
    Select Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', Null)
    Into v_����
    From Dual;
    n_��Լʣ������ := 0;
  
    v_Sql := 'Select a.*, Nvl(Hz.�ѹ���, 0) As �ѹ���, Nvl(Hz.��Լ��, 0) As ��Լ��, Nvl(Hz.�����ѽ���, 0) As �ѽ���, b.�ּ� As �۸� ';
    v_Sql := v_Sql ||
             'From (Select Ap.����id, Ap.����, Bm.���� As ��������, Ap.ҽ������, Nvl(Ap.ҽ��id, 0) As ҽ��id, Ry.רҵ����ְ�� As ְ��, Ap.����, ';
    v_Sql := v_Sql || ' Ap.����id, Ap.�ƻ�id, Ap.�Ű�, Ap.��Ŀid, Fy.���� As ��Ŀ����, Ap.��ſ���, Ap.�޺���, Ap.��Լ�� ';
    v_Sql := v_Sql || 'From (Select Ap.����id, Ap.����, Bm.���� As ��������, Ap.ҽ������, Ap.ҽ��id, Ap.����, Ap.Id As ����id, 0 As �ƻ�id, ';
    v_Sql := v_Sql || 'Ap.��Ŀid, Nvl(Ap.��ſ���, 0) As ��ſ���, ';
    v_Sql := v_Sql ||
             'Decode(To_Char(:1, ''D''), ''1'', Ap.����, ''2'', Ap.��һ, ''3'', Ap.�ܶ�, ''4'', Ap.����, ''5'', Ap.����, ';
    v_Sql := v_Sql || ' ''6'', Ap.����, ''7'', Ap.����, Null) As �Ű�, Xz.��Լ��, Xz.�޺��� ';
    v_Sql := v_Sql || 'From �ҺŰ��� Ap, ���ű� Bm, �ҺŰ������� Xz ';
    v_Sql := v_Sql || 'Where Ap.����id = Bm.Id(+) ';
  
    n_Curcount := 2;
    If Nvl(n_����id, 0) <> 0 Then
      v_Sql      := v_Sql || 'And Ap.����id = :2 ';
      n_Curcount := n_Curcount + 1;
    End If;
    If Nvl(n_ҽ��id, 0) <> 0 Then
      If n_Curcount = 2 Then
        v_Sql := v_Sql || 'And Ap.ҽ��id = :2 ';
      Else
        v_Sql := v_Sql || 'And Ap.ҽ��id = :3 ';
      End If;
      n_Curcount := n_Curcount + 1;
    End If;
    If Nvl(v_ҽ������, '_') <> '_' Then
      If n_Curcount = 2 Then
        v_Sql := v_Sql || 'And Ap.ҽ������ = :2 ';
      End If;
      If n_Curcount = 3 Then
        v_Sql := v_Sql || 'And Ap.ҽ������ = :3 ';
      End If;
      If n_Curcount = 4 Then
        v_Sql := v_Sql || 'And Ap.ҽ������ = :4 ';
      End If;
      n_Curcount := n_Curcount + 1;
    End If;
  
    v_Sql      := v_Sql || 'And Ap.ͣ������ Is Null And :' || n_Curcount ||
                  ' Between Nvl(Ap.��ʼʱ��, To_Date(''1900-01-01'', ''YYYY-MM-DD'')) And ';
    n_Curcount := n_Curcount + 1;
    v_Sql      := v_Sql || 'Nvl(Ap.��ֹʱ��, To_Date(''3000 - 01 - 01'', ''YYYY-MM-DD'')) And Xz.����id(+) = Ap.Id And ';
    v_Sql      := v_Sql || ' Xz.������Ŀ(+) = Decode(To_Char(:' || n_Curcount ||
                  ', ''D''), ''1'', ''����'', ''2'', ''��һ'', ''3'', ''�ܶ�'', ''4'', ''����'', ''5'', ';
    n_Curcount := n_Curcount + 1;
    v_Sql      := v_Sql || ' ''����'', ''6'', ''����'', ''7'', ''����'', Null) And Not Exists ';
    v_Sql      := v_Sql || '(Select Rownum ';
    v_Sql      := v_Sql || 'From �ҺŰ���ͣ��״̬ Ty ';
    v_Sql      := v_Sql || 'Where Ty.����id = Ap.Id And :' || n_Curcount ||
                  ' Between Ty.��ʼֹͣʱ�� And Ty.����ֹͣʱ��) And Not Exists ';
    n_Curcount := n_Curcount + 1;
    v_Sql      := v_Sql || '(Select Rownum ';
    v_Sql      := v_Sql || 'From �ҺŰ��żƻ� Jh Where Jh.����id = Ap.Id And Jh.���ʱ�� Is Not Null And ';
    v_Sql      := v_Sql || ':' || n_Curcount ||
                  ' Between Nvl(Jh.��Чʱ��, To_Date(''1900-01-01'', ''YYYY-MM-DD'')) And Nvl(Jh.ʧЧʱ��, To_Date(''3000-01-01'', ''YYYY-MM-DD''))) ';
    n_Curcount := n_Curcount + 1;
    v_Sql      := v_Sql || 'Union All ';
    v_Sql      := v_Sql ||
                  'Select Ap.����id, Ap.����, Bm.���� As ��������, Jh.ҽ������, Jh.ҽ��id, Ap.����, Ap.Id As ����id, Jh.Id As �ƻ�id, ';
    v_Sql      := v_Sql || 'Jh.��Ŀid, Nvl(Jh.��ſ���, 0) As ��ſ���,Decode(To_Char(:' || n_Curcount ||
                  ', ''D''), ''1'', Jh.����, ''2'', Jh.��һ, ''3'', Jh.�ܶ�, ''4'', Jh.����, ''5'', Jh.����, ';
    n_Curcount := n_Curcount + 1;
    v_Sql      := v_Sql || ' ''6'', Jh.����, ''7'', Jh.����, Null) As �Ű�, Xz.��Լ��, Xz.�޺��� ';
    v_Sql      := v_Sql || 'From �ҺŰ��żƻ� Jh, �ҺŰ��� Ap, ���ű� Bm, �Һżƻ����� Xz ';
    v_Sql      := v_Sql || 'Where Jh.����id = Ap.Id And Ap.����id = Bm.Id(+) And Ap.ͣ������ Is Null ';
  
    If Nvl(n_����id, 0) <> 0 Then
      v_Sql      := v_Sql || 'And Ap.����id = :' || n_Curcount || ' ';
      n_Curcount := n_Curcount + 1;
    End If;
    If Nvl(n_ҽ��id, 0) <> 0 Then
      v_Sql      := v_Sql || 'And Ap.ҽ��id = :' || n_Curcount || ' ';
      n_Curcount := n_Curcount + 1;
    End If;
    If Nvl(v_ҽ������, '_') <> '_' Then
      v_Sql      := v_Sql || 'And Ap.ҽ������ = :' || n_Curcount || ' ';
      n_Curcount := n_Curcount + 1;
    End If;
  
    v_Sql      := v_Sql || ' And :' || n_Curcount ||
                  ' Between Nvl(Jh.��Чʱ��, To_Date(''1900-01-01'', ''YYYY-MM-DD'')) And Nvl(Jh.ʧЧʱ��, To_Date(''3000-01-01'', ''YYYY-MM-DD'')) And Xz.�ƻ�id(+) = Jh.Id And ';
    n_Curcount := n_Curcount + 1;
    v_Sql      := v_Sql || 'Xz.������Ŀ(+) = Decode(To_Char(:' || n_Curcount ||
                  ', ''D''), ''1'', ''����'', ''2'', ''��һ'', ''3'', ''�ܶ�'', ''4'', ''����'', ''5'', ''����'', ''6'', ''����'', ''7'', ''����'', Null) And Not Exists ';
    n_Curcount := n_Curcount + 1;
    v_Sql      := v_Sql || '(Select Rownum From �ҺŰ���ͣ��״̬ Ty Where Ty.����id = Ap.Id And :' || n_Curcount ||
                  ' Between Ty.��ʼֹͣʱ�� And Ty.����ֹͣʱ��) And ';
    n_Curcount := n_Curcount + 1;
    v_Sql      := v_Sql || '(Jh.��Чʱ��, Jh.����id) = (Select Max(Sxjh.��Чʱ��) As ��Чʱ��, ����id From �ҺŰ��żƻ� Sxjh ';
    v_Sql      := v_Sql || ' Where Sxjh.���ʱ�� Is Not Null And :' || n_Curcount ||
                  ' Between Nvl(Sxjh.��Чʱ��, To_Date(''1900-01-01'', ''YYYY-MM-DD'')) And ';
    n_Curcount := n_Curcount + 1;
    v_Sql      := v_Sql || 'Nvl(Sxjh.ʧЧʱ��, To_Date(''3000-01-01'', ''YYYY-MM-DD'')) And Sxjh.����id = Jh.����id ';
    v_Sql      := v_Sql || 'Group By Sxjh.����id)) Ap, ���ű� Bm, ��Ա�� Ry, �շ���ĿĿ¼ Fy ';
    v_Sql      := v_Sql ||
                  'Where Ap.����id = Bm.Id(+) And Ap.ҽ��id = Ry.Id(+) And Ap.��Ŀid = Fy.Id And �Ű� Is Not Null) A, ';
    v_Sql      := v_Sql || '���˹ҺŻ��� Hz, �շѼ�Ŀ B ';
    v_Sql      := v_Sql || 'Where a.���� = Hz.����(+) And Hz.����(+) = Trunc(:' || n_Curcount ||
                  ') And a.��Ŀid = b.�շ�ϸĿid And ';
    n_Curcount := n_Curcount + 1;
    v_Sql      := v_Sql || 'Nvl(b.��ֹ����, To_Date(''3000-1-1'', ''YYYY-Mm-DD'')) > :' || n_Curcount || ' ';
    n_Curcount := n_Curcount + 1;
    v_Sql      := v_Sql || 'And b.ִ������ <= :' || n_Curcount || ' ';
    If Nvl(n_����id, 0) <> 0 And Nvl(n_ҽ��id, 0) <> 0 And Nvl(v_ҽ������, '_') <> '_' Then
      Open r_No For v_Sql
        Using d_����, n_����id, n_ҽ��id, v_ҽ������, d_����, d_����, d_����, d_����, d_����, n_����id, n_ҽ��id, v_ҽ������, d_����, d_����, d_����, d_����, d_����, d_����, d_����;
    End If;
    If Nvl(n_����id, 0) <> 0 And Nvl(n_ҽ��id, 0) = 0 And Nvl(v_ҽ������, '_') = '_' Then
      Open r_No For v_Sql
        Using d_����, n_����id, d_����, d_����, d_����, d_����, d_����, n_����id, d_����, d_����, d_����, d_����, d_����, d_����, d_����;
    End If;
    If Nvl(n_����id, 0) = 0 And Nvl(n_ҽ��id, 0) <> 0 And Nvl(v_ҽ������, '_') = '_' Then
      Open r_No For v_Sql
        Using d_����, n_ҽ��id, d_����, d_����, d_����, d_����, d_����, n_ҽ��id, d_����, d_����, d_����, d_����, d_����, d_����, d_����;
    End If;
    If Nvl(n_����id, 0) = 0 And Nvl(n_ҽ��id, 0) = 0 And Nvl(v_ҽ������, '_') <> '_' Then
      Open r_No For v_Sql
        Using d_����, v_ҽ������, d_����, d_����, d_����, d_����, d_����, v_ҽ������, d_����, d_����, d_����, d_����, d_����, d_����, d_����;
    End If;
    If Nvl(n_����id, 0) <> 0 And Nvl(n_ҽ��id, 0) <> 0 And Nvl(v_ҽ������, '_') = '_' Then
      Open r_No For v_Sql
        Using d_����, n_����id, n_ҽ��id, d_����, d_����, d_����, d_����, d_����, n_����id, n_ҽ��id, d_����, d_����, d_����, d_����, d_����, d_����, d_����;
    End If;
    If Nvl(n_����id, 0) <> 0 And Nvl(n_ҽ��id, 0) = 0 And Nvl(v_ҽ������, '_') <> '_' Then
      Open r_No For v_Sql
        Using d_����, n_����id, v_ҽ������, d_����, d_����, d_����, d_����, d_����, n_����id, v_ҽ������, d_����, d_����, d_����, d_����, d_����, d_����, d_����;
    End If;
    If Nvl(n_����id, 0) = 0 And Nvl(n_ҽ��id, 0) <> 0 And Nvl(v_ҽ������, '_') <> '_' Then
      Open r_No For v_Sql
        Using d_����, n_ҽ��id, v_ҽ������, d_����, d_����, d_����, d_����, d_����, n_ҽ��id, v_ҽ������, d_����, d_����, d_����, d_����, d_����, d_����, d_����;
    End If;
    Loop
      Fetch r_No
        Into r_����id, r_����, r_��������, r_ҽ������, r_ҽ��id, r_ְ��, r_����, r_����id, r_�ƻ�id, r_�Ű�, r_��Ŀid, r_��Ŀ����, r_��ſ���, r_�޺���,
             r_��Լ��, r_�ѹ���, r_��Լ��, r_�ѽ���, r_�۸�;
      Exit When r_No%NotFound;
      If r_�ƻ�id <> 0 Then
        Select Sign(Count(Rownum))
        Into n_��ʱ��
        From �ҺŰ��żƻ� Jh, �Һżƻ�ʱ�� Sd
        Where Jh.Id = Sd.�ƻ�id And Jh.Id = r_�ƻ�id And
              Sd.���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7',
                             '����', Null) And Rownum < 2;
      Else
        Select Sign(Count(Rownum))
        Into n_��ʱ��
        From �ҺŰ��� Ap, �ҺŰ���ʱ�� Sd
        Where Ap.Id = Sd.����id And Ap.Id = r_����id And
              Sd.���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7',
                             '����', Null) And Rownum < 2;
      End If;
      If n_��ʱ�� = 0 Then
        v_Temp := '';
        If v_������λ Is Not Null And r_��ſ��� = 1 Then
          If r_�ƻ�id <> 0 Then
            Select Nvl(Sum(����), 0)
            Into n_��Լ������
            From ������λ�ƻ�����
            Where �ƻ�id = r_�ƻ�id And ������λ = v_������λ And
                  ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                                '7', '����', Null);
            Select Count(1)
            Into n_��Լģʽ
            From ������λ�ƻ�����
            Where �ƻ�id = r_�ƻ�id And ������λ = v_������λ And
                  ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                                '7', '����', Null) And ��� = 0;
          Else
            Select Nvl(Sum(����), 0)
            Into n_��Լ������
            From ������λ���ſ���
            Where ����id = r_����id And ������λ = v_������λ And
                  ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                                '7', '����', Null);
            Select Count(1)
            Into n_��Լģʽ
            From ������λ���ſ���
            Where ����id = r_����id And ������λ = v_������λ And
                  ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                                '7', '����', Null) And ��� = 0;
          End If;
          If n_��Լģʽ = 0 Then
            If r_�ƻ�id <> 0 Then
              Select Count(1)
              Into n_��Լ�ѹ���
              From ���˹Һż�¼ A
              Where �ű� = r_���� And ��¼״̬ = 1 And ����ʱ�� Between Trunc(d_����) And Trunc(d_���� + 1) - 1 / 60 / 60 / 24 And
                    Exists (Select 1
                     From ������λ�ƻ�����
                     Where �ƻ�id = r_�ƻ�id And ������λ = v_������λ And
                           ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                         '����', '6', '����', '7', '����', Null) And ��� = a.���� And ���� <> 0);
            Else
              Select Count(1)
              Into n_��Լ�ѹ���
              From ���˹Һż�¼ A
              Where �ű� = r_���� And ��¼״̬ = 1 And ����ʱ�� Between Trunc(d_����) And Trunc(d_���� + 1) - 1 / 60 / 60 / 24 And
                    Exists (Select 1
                     From ������λ���ſ���
                     Where ����id = r_����id And ������λ = v_������λ And
                           ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                         '����', '6', '����', '7', '����', Null) And ��� = a.���� And ���� <> 0);
            End If;
          Else
            Begin
              Select Count(1)
              Into n_��Լ�ѹ���
              From ���˹Һż�¼
              Where �ű� = r_���� And ��¼״̬ = 1 And ������λ = v_������λ And ����ʱ�� Between Trunc(d_����) And
                    Trunc(d_���� + 1) - 1 / 60 / 60 / 24;
            Exception
              When Others Then
                n_��Լ�ѹ��� := 0;
            End;
          End If;
          If n_��Լ������ = 0 Then
            n_��Լʣ������ := 0;
          Else
            n_��Լʣ������ := n_��Լ������ - n_��Լ�ѹ���;
            If n_��Լʣ������ > (Nvl(r_�޺���, 0) - r_�ѹ���) Then
              n_��Լʣ������ := Nvl(r_�޺���, 0) - r_�ѹ���;
            End If;
          End If;
        End If;
      Else
        v_Temp := '<SPANLIST>';
        If r_�ƻ�id <> 0 Then
          Select Max(����ʱ��)
          Into d_�Ӻ�ʱ��
          From �Һżƻ�ʱ��
          Where �ƻ�id = r_�ƻ�id And ���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����',
                                              '6', '����', '7', '����', Null);
          If r_��ſ��� = 1 Then
            If Trunc(d_����) = Trunc(Sysdate) Then
              n_����ԤԼ := 0;
            Else
              Select Nvl(Max(Jh.�Ƿ�ԤԼ), 0)
              Into n_����ԤԼ
              From (Select Sd.�ƻ�id, Sd.���, Sd.����, Jh.����,
                            To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'),
                                     'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                            To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'),
                                     'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Sd.��������, Sd.�Ƿ�ԤԼ
                     From �ҺŰ��żƻ� Jh, �Һżƻ�ʱ�� Sd
                     Where Jh.Id = Sd.�ƻ�id And Jh.Id = r_�ƻ�id And
                           Sd.���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6',
                                          '����', '7', '����', Null)) Jh, �Һ����״̬ Zt
              Where Zt.����(+) = Jh.��ʼʱ�� And Zt.����(+) = Jh.���� And Zt.���(+) = Jh.��� And
                    Decode(Sign(Sysdate - Jh.��ʼʱ��), -1, 0, 1) <> 1;
            End If;
          
            For r_Time In (Select Jh.����, Jh.���, Jh.����, Jh.��ʼʱ��, Jh.����ʱ��, Jh.��������, Jh.�Ƿ�ԤԼ, 0 As ��Լ��,
                                  Decode(Nvl(Zt.���, 0), 0, 1, 0) As ʣ����,
                                  Decode(Sign(Sysdate - Jh.��ʼʱ��), -1, 0, 1) As ʧЧʱ��
                           
                           From (Select Sd.�ƻ�id, Sd.���, Sd.����, Jh.����,
                                         To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'),
                                                  'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                         To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'),
                                                  'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Sd.��������, Sd.�Ƿ�ԤԼ
                                  From �ҺŰ��żƻ� Jh, �Һżƻ�ʱ�� Sd
                                  Where Jh.Id = Sd.�ƻ�id And Jh.Id = r_�ƻ�id And
                                        Sd.���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                       '����', '6', '����', '7', '����', Null)) Jh, �Һ����״̬ Zt
                           Where Zt.����(+) = Jh.��ʼʱ�� And Zt.����(+) = Jh.���� And Zt.���(+) = Jh.��� And
                                 Decode(Sign(Sysdate - Jh.��ʼʱ��), -1, 0, 1) <> 1
                           Order By ���) Loop
              If v_������λ Is Not Null Then
                Begin
                  Select 1
                  Into n_��Լģʽ
                  From ������λ�ƻ�����
                  Where ������Ŀ = r_Time.���� And �ƻ�id = r_�ƻ�id And ��� = 0 And ������λ = v_������λ;
                Exception
                  When Others Then
                    n_��Լģʽ := 0;
                End;
              Else
                n_��Լģʽ := 0;
              End If;
              If r_Time.ʣ���� = 0 Then
                n_����ʣ�� := 0;
              Else
                n_����ʣ�� := r_Time.��������;
              End If;
              If v_������λ Is Null Or n_��Լģʽ = 1 Then
                Begin
                  Select 1
                  Into n_Exists
                  From ������λ�ƻ�����
                  Where ������Ŀ = r_Time.���� And �ƻ�id = r_�ƻ�id And ��� = r_Time.��� And Rownum < 2;
                Exception
                  When Others Then
                    n_Exists := 0;
                End;
                If n_Exists = 0 Then
                  If n_����ԤԼ = 1 And r_Time.�Ƿ�ԤԼ = 0 Then
                    Null;
                  Else
                    Begin
                      Select 1
                      Into n_�Ƿ�Ԥ��
                      From �Һ����״̬
                      Where ״̬ In (3, 4) And ���� = r_���� And ��� = r_Time.��� And Trunc(����) = Trunc(d_����);
                    Exception
                      When Others Then
                        n_�Ƿ�Ԥ�� := 0;
                    End;
                    If n_�Ƿ�Ԥ�� = 0 Then
                      v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                    To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                    '</SPAN>';
                      n_ʱ������ := Nvl(n_ʱ������, 0) + n_����ʣ��;
                    End If;
                  End If;
                End If;
              Else
                Begin
                  Select 1
                  Into n_Exists
                  From ������λ�ƻ�����
                  Where ������Ŀ = r_Time.���� And �ƻ�id = r_�ƻ�id And ��� = r_Time.��� And ������λ = v_������λ And Rownum < 2;
                Exception
                  When Others Then
                    n_Exists := 0;
                End;
                Begin
                  Select 0
                  Into n_�Ǻ�Լ
                  From ������λ�ƻ�����
                  Where �ƻ�id = r_�ƻ�id And ������λ = v_������λ And Rownum < 2;
                Exception
                  When Others Then
                    n_�Ǻ�Լ := 1;
                End;
                If n_Exists = 1 Or n_�Ǻ�Լ = 1 Then
                  If n_����ԤԼ = 1 And r_Time.�Ƿ�ԤԼ = 0 Then
                    Null;
                  Else
                    Begin
                      Select 1
                      Into n_�Ƿ�Ԥ��
                      From �Һ����״̬
                      Where ״̬ In (3, 4) And ���� = r_���� And ��� = r_Time.��� And Trunc(����) = Trunc(d_����);
                    Exception
                      When Others Then
                        n_�Ƿ�Ԥ�� := 0;
                    End;
                    If n_�Ƿ�Ԥ�� = 0 Then
                      v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                        To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                        '</SPAN>';
                      n_��Լʣ������ := n_��Լʣ������ + 1;
                    End If;
                  End If;
                End If;
              End If;
            End Loop;
          Else
            n_���������� := Nvl(r_��Լ��, Nvl(r_�޺���, 0)) - Nvl(r_��Լ��, 0);
            For r_Time In (Select Jh.����, Jh.���, Jh.����, Jh.��ʼʱ��, Jh.����ʱ��, Jh.��������, Jh.�Ƿ�ԤԼ,
                                  Sum(Decode(Nvl(Zt.���, 0), 0, 0, 1)) As ��Լ��,
                                  Jh.�������� - Sum(Decode(Nvl(Zt.���, 0), 0, 0, 1)) As ʣ����,
                                  Decode(Sign(Sysdate - Jh.��ʼʱ��), -1, 0, 1) As ʧЧʱ��
                           From (Select Sd.�ƻ�id, Sd.���, Sd.����, Jh.����,
                                         To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'),
                                                  'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                         To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'),
                                                  'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Sd.��������, Sd.�Ƿ�ԤԼ
                                  From �ҺŰ��żƻ� Jh, �Һżƻ�ʱ�� Sd
                                  Where Jh.Id = Sd.�ƻ�id And Jh.Id = r_�ƻ�id And
                                        Sd.���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                       '����', '6', '����', '7', '����', Null)) Jh, �Һ����״̬ Zt
                           Where Jh.���� = Zt.����(+) And Jh.��ʼʱ�� = Zt.����(+) And
                                 Decode(Sign(Sysdate - Jh.��ʼʱ��), -1, 0, 1) <> 1
                           Group By Jh.����, Jh.���, Jh.����, Jh.��ʼʱ��, Jh.����ʱ��, Jh.��������, Jh.�Ƿ�ԤԼ
                           Order By Jh.���) Loop
              If v_������λ Is Not Null Then
                Begin
                  Select 1
                  Into n_��Լģʽ
                  From ������λ�ƻ�����
                  Where ������Ŀ = r_Time.���� And �ƻ�id = r_�ƻ�id And ��� = 0 And ������λ = v_������λ;
                Exception
                  When Others Then
                    n_��Լģʽ := 0;
                End;
              Else
                n_��Լģʽ := 0;
              End If;
              n_����ʣ�� := r_Time.ʣ����;
              If v_������λ Is Null Or n_��Լģʽ = 1 Then
                Begin
                  Select 1
                  Into n_Exists
                  From ������λ�ƻ�����
                  Where ������Ŀ = r_Time.���� And �ƻ�id = r_�ƻ�id And ��� = r_Time.��� And Rownum < 2;
                Exception
                  When Others Then
                    n_Exists := 0;
                End;
                If n_Exists = 0 Then
                  If n_���������� < n_����ʣ�� Then
                    v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_���������� || '</SL>' ||
                                  '</SPAN>';
                    n_ʱ������ := Nvl(n_ʱ������, 0) + n_����������;
                  Else
                    v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                  '</SPAN>';
                    n_ʱ������ := Nvl(n_ʱ������, 0) + n_����ʣ��;
                  End If;
                End If;
              Else
                Begin
                  Select 1
                  Into n_Exists
                  From ������λ�ƻ�����
                  Where ������Ŀ = r_Time.���� And �ƻ�id = r_�ƻ�id And ��� = r_Time.��� And ������λ = v_������λ And Rownum < 2;
                Exception
                  When Others Then
                    n_Exists := 0;
                End;
                Begin
                  Select 0
                  Into n_�Ǻ�Լ
                  From ������λ�ƻ�����
                  Where �ƻ�id = r_�ƻ�id And ������λ = v_������λ And Rownum < 2;
                Exception
                  When Others Then
                    n_�Ǻ�Լ := 1;
                End;
                If n_Exists = 1 Or n_�Ǻ�Լ = 1 Then
                  If n_���������� < n_����ʣ�� Then
                    v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                      To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_���������� || '</SL>' ||
                                      '</SPAN>';
                    n_��Լʣ������ := n_��Լʣ������ + Nvl(n_����ʣ��, 0);
                  Else
                    v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                      To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                      '</SPAN>';
                    n_��Լʣ������ := n_��Լʣ������ + Nvl(n_����ʣ��, 0);
                  End If;
                End If;
              End If;
            End Loop;
          End If;
        Else
          Select Max(����ʱ��)
          Into d_�Ӻ�ʱ��
          From �ҺŰ���ʱ��
          Where ����id = r_����id And ���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����',
                                              '6', '����', '7', '����', Null);
          If r_��ſ��� = 1 Then
            If Trunc(d_����) = Trunc(Sysdate) Then
              n_����ԤԼ := 0;
            Else
              Select Nvl(Max(Ap.�Ƿ�ԤԼ), 0)
              Into n_����ԤԼ
              From (Select Sd.����id, Sd.���, Sd.����, Ap.����,
                            To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'),
                                     'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                            To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'),
                                     'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Sd.��������, Sd.�Ƿ�ԤԼ
                     From �ҺŰ��� Ap, �ҺŰ���ʱ�� Sd
                     Where Ap.Id = Sd.����id And Ap.Id = r_����id And
                           Sd.���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6',
                                          '����', '7', '����', Null)) Ap, �Һ����״̬ Zt
              Where Zt.����(+) = Ap.��ʼʱ�� And Zt.����(+) = Ap.���� And Zt.���(+) = Ap.��� And
                    Decode(Sign(Sysdate - Ap.��ʼʱ��), -1, 0, 1) <> 1;
            End If;
            For r_Time In (Select Ap.����, Ap.���, Ap.����, Ap.��ʼʱ��, Ap.����ʱ��, Ap.��������, Ap.�Ƿ�ԤԼ, 0 As ��Լ��,
                                  Decode(Nvl(Zt.���, 0), 0, 1, 0) As ʣ����,
                                  Decode(Sign(Sysdate - Ap.��ʼʱ��), -1, 0, 1) As ʧЧʱ��
                           
                           From (Select Sd.����id, Sd.���, Sd.����, Ap.����,
                                         To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'),
                                                  'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                         To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'),
                                                  'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Sd.��������, Sd.�Ƿ�ԤԼ
                                  From �ҺŰ��� Ap, �ҺŰ���ʱ�� Sd
                                  Where Ap.Id = Sd.����id And Ap.Id = r_����id And
                                        Sd.���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                       '����', '6', '����', '7', '����', Null)) Ap, �Һ����״̬ Zt
                           Where Zt.����(+) = Ap.��ʼʱ�� And Zt.����(+) = Ap.���� And Zt.���(+) = Ap.��� And
                                 Decode(Sign(Sysdate - Ap.��ʼʱ��), -1, 0, 1) <> 1
                           Order By ���) Loop
              If v_������λ Is Not Null Then
                Begin
                  Select 1
                  Into n_��Լģʽ
                  From ������λ���ſ���
                  Where ������Ŀ = r_Time.���� And ����id = r_����id And ��� = 0 And ������λ = v_������λ;
                Exception
                  When Others Then
                    n_��Լģʽ := 0;
                End;
              Else
                n_��Լģʽ := 0;
              End If;
              If r_Time.ʣ���� = 0 Then
                n_����ʣ�� := 0;
              Else
                n_����ʣ�� := r_Time.��������;
              End If;
              If v_������λ Is Null Or n_��Լģʽ = 1 Then
                Begin
                  Select 1
                  Into n_Exists
                  From ������λ���ſ���
                  Where ������Ŀ = r_Time.���� And ����id = r_����id And ��� = r_Time.��� And Rownum < 2;
                Exception
                  When Others Then
                    n_Exists := 0;
                End;
                If n_Exists = 0 Then
                  If n_����ԤԼ = 1 And r_Time.�Ƿ�ԤԼ = 0 Then
                    Null;
                  Else
                    Begin
                      Select 1
                      Into n_�Ƿ�Ԥ��
                      From �Һ����״̬
                      Where ״̬ In (3, 4) And ���� = r_���� And ��� = r_Time.��� And Trunc(����) = Trunc(d_����);
                    Exception
                      When Others Then
                        n_�Ƿ�Ԥ�� := 0;
                    End;
                    If n_�Ƿ�Ԥ�� = 0 Then
                      v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                    To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                    '</SPAN>';
                      n_ʱ������ := Nvl(n_ʱ������, 0) + n_����ʣ��;
                    End If;
                  End If;
                End If;
              Else
                Begin
                  Select 1
                  Into n_Exists
                  From ������λ���ſ���
                  Where ������Ŀ = r_Time.���� And ����id = r_����id And ��� = r_Time.��� And ������λ = v_������λ And Rownum < 2;
                Exception
                  When Others Then
                    n_Exists := 0;
                End;
                Begin
                  Select 0
                  Into n_�Ǻ�Լ
                  From ������λ���ſ���
                  Where ����id = r_����id And ������λ = v_������λ And Rownum < 2;
                Exception
                  When Others Then
                    n_�Ǻ�Լ := 1;
                End;
                If n_Exists = 1 Or n_�Ǻ�Լ = 1 Then
                  If n_����ԤԼ = 1 And r_Time.�Ƿ�ԤԼ = 0 Then
                    Null;
                  Else
                    Begin
                      Select 1
                      Into n_�Ƿ�Ԥ��
                      From �Һ����״̬
                      Where ״̬ In (3, 4) And ���� = r_���� And ��� = r_Time.��� And Trunc(����) = Trunc(d_����);
                    Exception
                      When Others Then
                        n_�Ƿ�Ԥ�� := 0;
                    End;
                    If n_�Ƿ�Ԥ�� = 0 Then
                      v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                        To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                        '</SPAN>';
                      n_��Լʣ������ := n_��Լʣ������ + 1;
                    End If;
                  End If;
                End If;
              End If;
            End Loop;
          Else
            n_���������� := Nvl(r_��Լ��, Nvl(r_�޺���, 0)) - Nvl(r_��Լ��, 0);
            For r_Time In (Select Ap.����, Ap.���, Ap.����, Ap.��ʼʱ��, Ap.����ʱ��, Ap.��������, Ap.�Ƿ�ԤԼ,
                                  Sum(Decode(Nvl(Zt.���, 0), 0, 0, 1)) As ��Լ��,
                                  Ap.�������� - Sum(Decode(Nvl(Zt.���, 0), 0, 0, 1)) As ʣ����,
                                  Decode(Sign(Sysdate - Ap.��ʼʱ��), -1, 0, 1) As ʧЧʱ��
                           From (Select Sd.����id, Sd.���, Sd.����, Ap.����,
                                         To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'),
                                                  'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                         To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'),
                                                  'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Sd.��������, Sd.�Ƿ�ԤԼ
                                  From �ҺŰ��� Ap, �ҺŰ���ʱ�� Sd
                                  Where Ap.Id = Sd.����id And Ap.Id = r_����id And
                                        Sd.���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                       '����', '6', '����', '7', '����', Null)) Ap, �Һ����״̬ Zt
                           Where Ap.���� = Zt.����(+) And Ap.��ʼʱ�� = Zt.����(+) And
                                 Decode(Sign(Sysdate - Ap.��ʼʱ��), -1, 0, 1) <> 1
                           Group By Ap.����, Ap.���, Ap.����, Ap.��ʼʱ��, Ap.����ʱ��, Ap.��������, Ap.�Ƿ�ԤԼ
                           Order By Ap.���) Loop
              If v_������λ Is Not Null Then
                Begin
                  Select 1
                  Into n_��Լģʽ
                  From ������λ���ſ���
                  Where ������Ŀ = r_Time.���� And ����id = r_����id And ��� = 0 And ������λ = v_������λ;
                Exception
                  When Others Then
                    n_��Լģʽ := 0;
                End;
              Else
                n_��Լģʽ := 0;
              End If;
              n_����ʣ�� := r_Time.ʣ����;
              If v_������λ Is Null Or n_��Լģʽ = 1 Then
                Begin
                  Select 1
                  Into n_Exists
                  From ������λ���ſ���
                  Where ������Ŀ = r_Time.���� And ����id = r_����id And ��� = r_Time.��� And Rownum < 2;
                Exception
                  When Others Then
                    n_Exists := 0;
                End;
                If n_Exists = 0 Then
                  If n_���������� < n_����ʣ�� Then
                    v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_���������� || '</SL>' ||
                                  '</SPAN>';
                    n_ʱ������ := Nvl(n_ʱ������, 0) + n_����������;
                  Else
                    v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                  '</SPAN>';
                    n_ʱ������ := Nvl(n_ʱ������, 0) + n_����ʣ��;
                  End If;
                End If;
              Else
                Begin
                  Select 1
                  Into n_Exists
                  From ������λ���ſ���
                  Where ������Ŀ = r_Time.���� And ����id = r_����id And ��� = r_Time.��� And ������λ = v_������λ And Rownum < 2;
                Exception
                  When Others Then
                    n_Exists := 0;
                End;
                Begin
                  Select 0
                  Into n_�Ǻ�Լ
                  From ������λ���ſ���
                  Where ����id = r_����id And ������λ = v_������λ And Rownum < 2;
                Exception
                  When Others Then
                    n_�Ǻ�Լ := 1;
                End;
                If n_Exists = 1 Or n_�Ǻ�Լ = 1 Then
                  If n_���������� < n_����ʣ�� Then
                    v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                      To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_���������� || '</SL>' ||
                                      '</SPAN>';
                    n_��Լʣ������ := n_��Լʣ������ + Nvl(n_����ʣ��, 0);
                  Else
                    v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                      To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                      '</SPAN>';
                    n_��Լʣ������ := n_��Լʣ������ + Nvl(n_����ʣ��, 0);
                  End If;
                End If;
              End If;
            End Loop;
          End If;
        End If;
      End If;
      If v_������λ Is Not Null Then
        If Nvl(r_�ƻ�id, 0) <> 0 Then
          Begin
            Select 0
            Into n_�Ǻ�Լ
            From ������λ�ƻ�����
            Where �ƻ�id = r_�ƻ�id And ������λ = v_������λ And Rownum < 2;
          Exception
            When Others Then
              n_�Ǻ�Լ := 1;
          End;
        Else
          Begin
            Select 0
            Into n_�Ǻ�Լ
            From ������λ���ſ���
            Where ����id = r_����id And ������λ = v_������λ And Rownum < 2;
          Exception
            When Others Then
              n_�Ǻ�Լ := 1;
          End;
        End If;
      End If;
      If v_������λ Is Null Or n_�Ǻ�Լ = 1 Then
        If r_�޺��� = 0 Then
          v_ʣ������ := '';
        Else
          If Nvl(r_�ƻ�id, 0) <> 0 Then
            Select Sum(����)
            Into n_��Լ������
            From ������λ�ƻ�����
            Where �ƻ�id = r_�ƻ�id And ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                  '����', '6', '����', '7', '����', Null);
          Else
            Select Sum(����)
            Into n_��Լ������
            From ������λ���ſ���
            Where ����id = r_����id And ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                  '����', '6', '����', '7', '����', Null);
          End If;
          Begin
            Select Count(1)
            Into n_��Լ�ѹ���
            From ���˹Һż�¼
            Where �ű� = r_���� And ��¼״̬ = 1 And ������λ Is Not Null And ����ʱ�� Between Trunc(d_����) And
                  Trunc(d_���� + 1) - 1 / 60 / 60 / 24;
          Exception
            When Others Then
              n_��Լ�ѹ��� := 0;
          End;
          Select Count(1)
          Into n_Ԥ������
          From �Һ����״̬
          Where ״̬ = 3 And ���� = r_���� And Trunc(����) = Trunc(d_����);
          If Trunc(d_����) = Trunc(Sysdate) Then
            If Nvl(n_��Լ������, 0) = 0 Then
              v_ʣ������ := r_�޺��� - r_�ѹ��� - r_��Լ�� + r_�ѽ��� - n_Ԥ������;
            Else
              v_ʣ������ := r_�޺��� - r_�ѹ��� - r_��Լ�� + r_�ѽ��� - n_��Լ������ + n_��Լ�ѹ��� - n_Ԥ������;
            End If;
            n_�ѹ��� := r_�ѹ���;
            If Nvl(n_ʱ������, 0) < v_ʣ������ And n_��ʱ�� <> 0 Then
              n_������� := 1;
              v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_�Ӻ�ʱ��, 'hh24:mi:ss') || '-' || '</SJD>' || '<SL>' ||
                            To_Number(v_ʣ������ - Nvl(n_ʱ������, 0)) || '</SL>' || '</SPAN>';
            Else
              n_������� := 0;
            End If;
          Else
            If Nvl(n_��Լ������, 0) = 0 Then
              v_ʣ������ := r_��Լ�� - r_��Լ�� - n_Ԥ������;
              If v_ʣ������ Is Null Then
                v_ʣ������ := r_�޺��� - r_�ѹ��� - r_��Լ�� + r_�ѽ��� - n_Ԥ������;
              End If;
            Else
              v_ʣ������ := r_��Լ�� - r_��Լ�� - n_��Լ������ + n_��Լ�ѹ��� - n_Ԥ������;
              If v_ʣ������ Is Null Then
                v_ʣ������ := r_�޺��� - r_�ѹ��� - r_��Լ�� + r_�ѽ��� - n_��Լ������ + n_��Լ�ѹ��� - n_Ԥ������;
              End If;
            End If;
            n_�ѹ��� := r_�ѹ���;
          End If;
        End If;
      Else
        If Nvl(r_�ƻ�id, 0) <> 0 Then
          If v_������λ Is Not Null Then
            Begin
              Select 1
              Into n_��Լģʽ
              From ������λ�ƻ�����
              Where ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                                  '7', '����', Null) And �ƻ�id = r_�ƻ�id And ��� = 0 And ������λ = v_������λ;
            Exception
              When Others Then
                n_��Լģʽ := 0;
            End;
          Else
            n_��Լģʽ := 0;
          End If;
          Select Sum(����)
          Into n_��Լ������
          From ������λ�ƻ�����
          Where �ƻ�id = r_�ƻ�id And ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                '����', '6', '����', '7', '����', Null) And ������λ = v_������λ;
        Else
          If v_������λ Is Not Null Then
            Begin
              Select 1
              Into n_��Լģʽ
              From ������λ���ſ���
              Where ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                                  '7', '����', Null) And ����id = r_����id And ��� = 0 And ������λ = v_������λ;
            Exception
              When Others Then
                n_��Լģʽ := 0;
            End;
          Else
            n_��Լģʽ := 0;
          End If;
          Select Sum(����)
          Into n_��Լ������
          From ������λ���ſ���
          Where ����id = r_����id And ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                '����', '6', '����', '7', '����', Null) And ������λ = v_������λ;
        End If;
        If n_��Լģʽ = 0 Then
          v_ʣ������   := n_��Լʣ������;
          n_�ѹ���     := r_�ѹ���;
          n_��Լ�ѹ��� := Nvl(n_��Լ������, 0) - n_��Լʣ������;
        Else
          n_�ѹ��� := r_�ѹ���;
          Begin
            Select Count(1)
            Into n_��Լ�ѹ���
            From ���˹Һż�¼
            Where �ű� = r_���� And ��¼״̬ = 1 And ������λ = v_������λ And ����ʱ�� Between Trunc(d_����) And
                  Trunc(d_���� + 1) - 1 / 60 / 60 / 24;
          Exception
            When Others Then
              n_��Լ�ѹ��� := 0;
          End;
          If Nvl(n_��Լ������, 0) = 0 Then
            v_ʣ������ := '0';
          Else
            v_ʣ������ := n_��Լ������ - n_��Լ�ѹ���;
          End If;
        End If;
      End If;
      Select To_Char(��ʼʱ��, 'hh24:mi') Into v_Timetemp From ʱ��� Where ʱ��� = r_�Ű�;
      v_ʱ��� := v_Timetemp || '-';
      Select To_Char(��ֹʱ��, 'hh24:mi') Into v_Timetemp From ʱ��� Where ʱ��� = r_�Ű�;
      v_ʱ��� := v_ʱ��� || v_Timetemp;
      If v_Temp Is Not Null Then
        v_Temp := v_Temp || '</SPANLIST>';
      End If;
      If v_������λ Is Not Null Then
        If Nvl(r_�ƻ�id, 0) <> 0 Then
          Begin
            Select 1
            Into n_����
            From ������λ�ƻ�����
            Where �ƻ�id = r_�ƻ�id And ������λ = v_������λ And ���� = 0 And Rownum < 2;
          Exception
            When Others Then
              n_���� := 0;
          End;
        Else
          Begin
            Select 1
            Into n_����
            From ������λ���ſ���
            Where ����id = r_����id And ������λ = v_������λ And ���� = 0 And Rownum < 2;
          Exception
            When Others Then
              n_���� := 0;
          End;
        End If;
      End If;
      --��Լ��=0��ԤԼ��ֹ
      If Trunc(d_����) <> Trunc(Sysdate) Then
        If r_��Լ�� = 0 Then
          n_���� := 1;
        End If;
      End If;
      If Nvl(n_����, 0) = 0 Then
        --���������
        n_�ϼƽ�� := r_�۸�;
        For r_Subfee In (Select �ּ�, ��������
                         From �շѴ�����Ŀ A, �շѼ�Ŀ B
                         Where a.����id = r_��Ŀid And a.����id = b.�շ�ϸĿid And Sysdate Between b.ִ������ And
                               Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))) Loop
          n_�ϼƽ�� := n_�ϼƽ�� + r_Subfee.�ּ� * r_Subfee.��������;
        End Loop;
        If Trunc(Sysdate) = Trunc(d_����) Then
          Begin
            Select 1
            Into n_Exists
            From (Select ʱ���
                   From ʱ���
                   Where ('3000-01-10 ' || To_Char(Sysdate, 'HH24:MI:SS') < '3000-01-10 ' || To_Char(��ֹʱ��, 'HH24:MI:SS')) Or
                         ('3000-01-10 ' || To_Char(Sysdate, 'HH24:MI:SS') <
                         Decode(Sign(��ʼʱ�� - ��ֹʱ��), 1, '3000-01-11 ' || To_Char(��ֹʱ��, 'HH24:MI:SS'),
                                 '3000-01-10 ' || To_Char(��ֹʱ��, 'HH24:MI:SS'))))
            Where ʱ��� = r_�Ű�;
          Exception
            When Others Then
              n_Exists := 0;
          End;
        Else
          n_Exists := 1;
        End If;
        If n_Exists = 1 Then
          If v_ʣ������ > 0 Then
            c_Xmlmain := '<HB>' || '<HM>' || r_���� || '</HM>' || '<YSID>' || r_ҽ��id || '</YSID>' || '<YS>' || r_ҽ������ ||
                         '</YS>' || '<KSID>' || r_����id || '</KSID>' || '<KSMC>' || r_�������� || '</KSMC>' || '<ZC>' || r_ְ�� ||
                         '</ZC>' || '<XMID>' || r_��Ŀid || '</XMID>' || '<XMMC>' || r_��Ŀ���� || '</XMMC>' || '<YGHS>' ||
                         n_�ѹ��� || '</YGHS>' || '<SYHS>' || v_ʣ������ || '</SYHS>' || '<PRICE>' || n_�ϼƽ�� || '</PRICE>' ||
                         '<HCXH>' || n_������� || '</HCXH>' || '<HL>' || r_���� || '</HL>' || '<FSD>' || n_��ʱ�� || '</FSD>' ||
                         '<HBTIME>' || v_ʱ��� || '</HBTIME>' || '<FWMC>' || r_�Ű� || '</FWMC>' || v_Temp || '</HB>';
            v_Xmlmain := v_Xmlmain || c_Xmlmain;
          Else
            c_Xmlmain := '<HB>' || '<HM>' || r_���� || '</HM>' || '<YSID>' || r_ҽ��id || '</YSID>' || '<YS>' || r_ҽ������ ||
                         '</YS>' || '<KSID>' || r_����id || '</KSID>' || '<KSMC>' || r_�������� || '</KSMC>' || '<ZC>' || r_ְ�� ||
                         '</ZC>' || '<XMID>' || r_��Ŀid || '</XMID>' || '<XMMC>' || r_��Ŀ���� || '</XMMC>' || '<YGHS>' ||
                         n_�ѹ��� || '</YGHS>' || '<SYHS>' || v_ʣ������ || '</SYHS>' || '<PRICE>' || n_�ϼƽ�� || '</PRICE>' ||
                         '<HL>' || r_���� || '</HL>' || '<FSD>' || n_��ʱ�� || '</FSD>' || '<HBTIME>' || v_ʱ��� ||
                         '</HBTIME>' || '<FWMC>' || r_�Ű� || '</FWMC>' || '</HB>';
            v_Xmlmain := v_Xmlmain || c_Xmlmain;
          End If;
        End If;
      End If;
      n_��Լʣ������ := 0;
      n_��Լ������   := 0;
      n_ʱ������     := 0;
      n_����         := 0;
      n_�Ǻ�Լ       := 0;
    End Loop;
    Close r_No;
    v_Xmlmain := '<GROUP>' || '<RQ>' || To_Char(d_����, 'yyyy-mm-dd') || '</RQ>' || '<HBLIST>' || v_Xmlmain ||
                 '</HBLIST>' || '</GROUP>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Xmlmain)) Into x_Templet From Dual;
  Else
    If Trunc(d_����) < To_Date(Substr(v_����ʱ��, 1, Instr(v_����ʱ��, ' ') - 1), 'yyyy-mm-dd') Then
      Select Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����',
                     Null)
      Into v_����
      From Dual;
      n_��Լʣ������ := 0;
    
      v_Sql := 'Select a.*, Nvl(Hz.�ѹ���, 0) As �ѹ���, Nvl(Hz.��Լ��, 0) As ��Լ��, Nvl(Hz.�����ѽ���, 0) As �ѽ���, b.�ּ� As �۸� ';
      v_Sql := v_Sql ||
               'From (Select Ap.����id, Ap.����, Bm.���� As ��������, Ap.ҽ������, Nvl(Ap.ҽ��id, 0) As ҽ��id, Ry.רҵ����ְ�� As ְ��, Ap.����, ';
      v_Sql := v_Sql || ' Ap.����id, Ap.�ƻ�id, Ap.�Ű�, Ap.��Ŀid, Fy.���� As ��Ŀ����, Ap.��ſ���, Ap.�޺���, Ap.��Լ�� ';
      v_Sql := v_Sql ||
               'From (Select Ap.����id, Ap.����, Bm.���� As ��������, Ap.ҽ������, Ap.ҽ��id, Ap.����, Ap.Id As ����id, 0 As �ƻ�id, ';
      v_Sql := v_Sql || 'Ap.��Ŀid, Nvl(Ap.��ſ���, 0) As ��ſ���, ';
      v_Sql := v_Sql ||
               'Decode(To_Char(:1, ''D''), ''1'', Ap.����, ''2'', Ap.��һ, ''3'', Ap.�ܶ�, ''4'', Ap.����, ''5'', Ap.����, ';
      v_Sql := v_Sql || ' ''6'', Ap.����, ''7'', Ap.����, Null) As �Ű�, Xz.��Լ��, Xz.�޺��� ';
      v_Sql := v_Sql || 'From �ҺŰ��� Ap, ���ű� Bm, �ҺŰ������� Xz ';
      v_Sql := v_Sql || 'Where Ap.����id = Bm.Id(+) ';
    
      n_Curcount := 2;
      If Nvl(n_����id, 0) <> 0 Then
        v_Sql      := v_Sql || 'And Ap.����id = :2 ';
        n_Curcount := n_Curcount + 1;
      End If;
      If Nvl(n_ҽ��id, 0) <> 0 Then
        If n_Curcount = 2 Then
          v_Sql := v_Sql || 'And Ap.ҽ��id = :2 ';
        Else
          v_Sql := v_Sql || 'And Ap.ҽ��id = :3 ';
        End If;
        n_Curcount := n_Curcount + 1;
      End If;
      If Nvl(v_ҽ������, '_') <> '_' Then
        If n_Curcount = 2 Then
          v_Sql := v_Sql || 'And Ap.ҽ������ = :2 ';
        End If;
        If n_Curcount = 3 Then
          v_Sql := v_Sql || 'And Ap.ҽ������ = :3 ';
        End If;
        If n_Curcount = 4 Then
          v_Sql := v_Sql || 'And Ap.ҽ������ = :4 ';
        End If;
        n_Curcount := n_Curcount + 1;
      End If;
    
      v_Sql      := v_Sql || 'And Ap.ͣ������ Is Null And :' || n_Curcount ||
                    ' Between Nvl(Ap.��ʼʱ��, To_Date(''1900-01-01'', ''YYYY-MM-DD'')) And ';
      n_Curcount := n_Curcount + 1;
      v_Sql      := v_Sql || 'Nvl(Ap.��ֹʱ��, To_Date(''3000 - 01 - 01'', ''YYYY-MM-DD'')) And Xz.����id(+) = Ap.Id And ';
      v_Sql      := v_Sql || ' Xz.������Ŀ(+) = Decode(To_Char(:' || n_Curcount ||
                    ', ''D''), ''1'', ''����'', ''2'', ''��һ'', ''3'', ''�ܶ�'', ''4'', ''����'', ''5'', ';
      n_Curcount := n_Curcount + 1;
      v_Sql      := v_Sql || ' ''����'', ''6'', ''����'', ''7'', ''����'', Null) And Not Exists ';
      v_Sql      := v_Sql || '(Select Rownum ';
      v_Sql      := v_Sql || 'From �ҺŰ���ͣ��״̬ Ty ';
      v_Sql      := v_Sql || 'Where Ty.����id = Ap.Id And :' || n_Curcount ||
                    ' Between Ty.��ʼֹͣʱ�� And Ty.����ֹͣʱ��) And Not Exists ';
      n_Curcount := n_Curcount + 1;
      v_Sql      := v_Sql || '(Select Rownum ';
      v_Sql      := v_Sql || 'From �ҺŰ��żƻ� Jh Where Jh.����id = Ap.Id And Jh.���ʱ�� Is Not Null And ';
      v_Sql      := v_Sql || ':' || n_Curcount ||
                    ' Between Nvl(Jh.��Чʱ��, To_Date(''1900-01-01'', ''YYYY-MM-DD'')) And Nvl(Jh.ʧЧʱ��, To_Date(''3000-01-01'', ''YYYY-MM-DD''))) ';
      n_Curcount := n_Curcount + 1;
      v_Sql      := v_Sql || 'Union All ';
      v_Sql      := v_Sql ||
                    'Select Ap.����id, Ap.����, Bm.���� As ��������, Jh.ҽ������, Jh.ҽ��id, Ap.����, Ap.Id As ����id, Jh.Id As �ƻ�id, ';
      v_Sql      := v_Sql || 'Jh.��Ŀid, Nvl(Jh.��ſ���, 0) As ��ſ���,Decode(To_Char(:' || n_Curcount ||
                    ', ''D''), ''1'', Jh.����, ''2'', Jh.��һ, ''3'', Jh.�ܶ�, ''4'', Jh.����, ''5'', Jh.����, ';
      n_Curcount := n_Curcount + 1;
      v_Sql      := v_Sql || ' ''6'', Jh.����, ''7'', Jh.����, Null) As �Ű�, Xz.��Լ��, Xz.�޺��� ';
      v_Sql      := v_Sql || 'From �ҺŰ��żƻ� Jh, �ҺŰ��� Ap, ���ű� Bm, �Һżƻ����� Xz ';
      v_Sql      := v_Sql || 'Where Jh.����id = Ap.Id And Ap.����id = Bm.Id(+) And Ap.ͣ������ Is Null ';
    
      If Nvl(n_����id, 0) <> 0 Then
        v_Sql      := v_Sql || 'And Ap.����id = :' || n_Curcount || ' ';
        n_Curcount := n_Curcount + 1;
      End If;
      If Nvl(n_ҽ��id, 0) <> 0 Then
        v_Sql      := v_Sql || 'And Ap.ҽ��id = :' || n_Curcount || ' ';
        n_Curcount := n_Curcount + 1;
      End If;
      If Nvl(v_ҽ������, '_') <> '_' Then
        v_Sql      := v_Sql || 'And Ap.ҽ������ = :' || n_Curcount || ' ';
        n_Curcount := n_Curcount + 1;
      End If;
    
      v_Sql      := v_Sql || ' And :' || n_Curcount ||
                    ' Between Nvl(Jh.��Чʱ��, To_Date(''1900-01-01'', ''YYYY-MM-DD'')) And Nvl(Jh.ʧЧʱ��, To_Date(''3000-01-01'', ''YYYY-MM-DD'')) And Xz.�ƻ�id(+) = Jh.Id And ';
      n_Curcount := n_Curcount + 1;
      v_Sql      := v_Sql || 'Xz.������Ŀ(+) = Decode(To_Char(:' || n_Curcount ||
                    ', ''D''), ''1'', ''����'', ''2'', ''��һ'', ''3'', ''�ܶ�'', ''4'', ''����'', ''5'', ''����'', ''6'', ''����'', ''7'', ''����'', Null) And Not Exists ';
      n_Curcount := n_Curcount + 1;
      v_Sql      := v_Sql || '(Select Rownum From �ҺŰ���ͣ��״̬ Ty Where Ty.����id = Ap.Id And :' || n_Curcount ||
                    ' Between Ty.��ʼֹͣʱ�� And Ty.����ֹͣʱ��) And ';
      n_Curcount := n_Curcount + 1;
      v_Sql      := v_Sql || '(Jh.��Чʱ��, Jh.����id) = (Select Max(Sxjh.��Чʱ��) As ��Чʱ��, ����id From �ҺŰ��żƻ� Sxjh ';
      v_Sql      := v_Sql || ' Where Sxjh.���ʱ�� Is Not Null And :' || n_Curcount ||
                    ' Between Nvl(Sxjh.��Чʱ��, To_Date(''1900-01-01'', ''YYYY-MM-DD'')) And ';
      n_Curcount := n_Curcount + 1;
      v_Sql      := v_Sql || 'Nvl(Sxjh.ʧЧʱ��, To_Date(''3000-01-01'', ''YYYY-MM-DD'')) And Sxjh.����id = Jh.����id ';
      v_Sql      := v_Sql || 'Group By Sxjh.����id)) Ap, ���ű� Bm, ��Ա�� Ry, �շ���ĿĿ¼ Fy ';
      v_Sql      := v_Sql ||
                    'Where Ap.����id = Bm.Id(+) And Ap.ҽ��id = Ry.Id(+) And Ap.��Ŀid = Fy.Id And �Ű� Is Not Null) A, ';
      v_Sql      := v_Sql || '���˹ҺŻ��� Hz, �շѼ�Ŀ B ';
      v_Sql      := v_Sql || 'Where a.���� = Hz.����(+) And Hz.����(+) = Trunc(:' || n_Curcount ||
                    ') And a.��Ŀid = b.�շ�ϸĿid And ';
      n_Curcount := n_Curcount + 1;
      v_Sql      := v_Sql || 'Nvl(b.��ֹ����, To_Date(''3000-1-1'', ''YYYY-Mm-DD'')) > :' || n_Curcount || ' ';
      n_Curcount := n_Curcount + 1;
      v_Sql      := v_Sql || 'And b.ִ������ <= :' || n_Curcount || ' ';
      If Nvl(n_����id, 0) <> 0 And Nvl(n_ҽ��id, 0) <> 0 And Nvl(v_ҽ������, '_') <> '_' Then
        Open r_No For v_Sql
          Using d_����, n_����id, n_ҽ��id, v_ҽ������, d_����, d_����, d_����, d_����, d_����, n_����id, n_ҽ��id, v_ҽ������, d_����, d_����, d_����, d_����, d_����, d_����, d_����;
      End If;
      If Nvl(n_����id, 0) <> 0 And Nvl(n_ҽ��id, 0) = 0 And Nvl(v_ҽ������, '_') = '_' Then
        Open r_No For v_Sql
          Using d_����, n_����id, d_����, d_����, d_����, d_����, d_����, n_����id, d_����, d_����, d_����, d_����, d_����, d_����, d_����;
      End If;
      If Nvl(n_����id, 0) = 0 And Nvl(n_ҽ��id, 0) <> 0 And Nvl(v_ҽ������, '_') = '_' Then
        Open r_No For v_Sql
          Using d_����, n_ҽ��id, d_����, d_����, d_����, d_����, d_����, n_ҽ��id, d_����, d_����, d_����, d_����, d_����, d_����, d_����;
      End If;
      If Nvl(n_����id, 0) = 0 And Nvl(n_ҽ��id, 0) = 0 And Nvl(v_ҽ������, '_') <> '_' Then
        Open r_No For v_Sql
          Using d_����, v_ҽ������, d_����, d_����, d_����, d_����, d_����, v_ҽ������, d_����, d_����, d_����, d_����, d_����, d_����, d_����;
      End If;
      If Nvl(n_����id, 0) <> 0 And Nvl(n_ҽ��id, 0) <> 0 And Nvl(v_ҽ������, '_') = '_' Then
        Open r_No For v_Sql
          Using d_����, n_����id, n_ҽ��id, d_����, d_����, d_����, d_����, d_����, n_����id, n_ҽ��id, d_����, d_����, d_����, d_����, d_����, d_����, d_����;
      End If;
      If Nvl(n_����id, 0) <> 0 And Nvl(n_ҽ��id, 0) = 0 And Nvl(v_ҽ������, '_') <> '_' Then
        Open r_No For v_Sql
          Using d_����, n_����id, v_ҽ������, d_����, d_����, d_����, d_����, d_����, n_����id, v_ҽ������, d_����, d_����, d_����, d_����, d_����, d_����, d_����;
      End If;
      If Nvl(n_����id, 0) = 0 And Nvl(n_ҽ��id, 0) <> 0 And Nvl(v_ҽ������, '_') <> '_' Then
        Open r_No For v_Sql
          Using d_����, n_ҽ��id, v_ҽ������, d_����, d_����, d_����, d_����, d_����, n_ҽ��id, v_ҽ������, d_����, d_����, d_����, d_����, d_����, d_����, d_����;
      End If;
      Loop
        Fetch r_No
          Into r_����id, r_����, r_��������, r_ҽ������, r_ҽ��id, r_ְ��, r_����, r_����id, r_�ƻ�id, r_�Ű�, r_��Ŀid, r_��Ŀ����, r_��ſ���, r_�޺���,
               r_��Լ��, r_�ѹ���, r_��Լ��, r_�ѽ���, r_�۸�;
        Exit When r_No%NotFound;
        If r_�ƻ�id <> 0 Then
          Select Sign(Count(Rownum))
          Into n_��ʱ��
          From �ҺŰ��żƻ� Jh, �Һżƻ�ʱ�� Sd
          Where Jh.Id = Sd.�ƻ�id And Jh.Id = r_�ƻ�id And
                Sd.���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7',
                               '����', Null) And Rownum < 2;
        Else
          Select Sign(Count(Rownum))
          Into n_��ʱ��
          From �ҺŰ��� Ap, �ҺŰ���ʱ�� Sd
          Where Ap.Id = Sd.����id And Ap.Id = r_����id And
                Sd.���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7',
                               '����', Null) And Rownum < 2;
        End If;
        If n_��ʱ�� = 0 Then
          v_Temp := '';
          If v_������λ Is Not Null And r_��ſ��� = 1 Then
            If r_�ƻ�id <> 0 Then
              Select Nvl(Sum(����), 0)
              Into n_��Լ������
              From ������λ�ƻ�����
              Where �ƻ�id = r_�ƻ�id And ������λ = v_������λ And
                    ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                                  '7', '����', Null);
              Select Count(1)
              Into n_��Լģʽ
              From ������λ�ƻ�����
              Where �ƻ�id = r_�ƻ�id And ������λ = v_������λ And
                    ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                                  '7', '����', Null) And ��� = 0;
            Else
              Select Nvl(Sum(����), 0)
              Into n_��Լ������
              From ������λ���ſ���
              Where ����id = r_����id And ������λ = v_������λ And
                    ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                                  '7', '����', Null);
              Select Count(1)
              Into n_��Լģʽ
              From ������λ���ſ���
              Where ����id = r_����id And ������λ = v_������λ And
                    ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                                  '7', '����', Null) And ��� = 0;
            End If;
            If n_��Լģʽ = 0 Then
              If r_�ƻ�id <> 0 Then
                Select Count(1)
                Into n_��Լ�ѹ���
                From ���˹Һż�¼ A
                Where �ű� = r_���� And ��¼״̬ = 1 And ����ʱ�� Between Trunc(d_����) And Trunc(d_���� + 1) - 1 / 60 / 60 / 24 And
                      Exists
                 (Select 1
                       From ������λ�ƻ�����
                       Where �ƻ�id = r_�ƻ�id And ������λ = v_������λ And
                             ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6',
                                           '����', '7', '����', Null) And ��� = a.���� And ���� <> 0);
              Else
                Select Count(1)
                Into n_��Լ�ѹ���
                From ���˹Һż�¼ A
                Where �ű� = r_���� And ��¼״̬ = 1 And ����ʱ�� Between Trunc(d_����) And Trunc(d_���� + 1) - 1 / 60 / 60 / 24 And
                      Exists
                 (Select 1
                       From ������λ���ſ���
                       Where ����id = r_����id And ������λ = v_������λ And
                             ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6',
                                           '����', '7', '����', Null) And ��� = a.���� And ���� <> 0);
              End If;
            Else
              Begin
                Select Count(1)
                Into n_��Լ�ѹ���
                From ���˹Һż�¼
                Where �ű� = r_���� And ��¼״̬ = 1 And ������λ = v_������λ And ����ʱ�� Between Trunc(d_����) And
                      Trunc(d_���� + 1) - 1 / 60 / 60 / 24;
              Exception
                When Others Then
                  n_��Լ�ѹ��� := 0;
              End;
            End If;
            If n_��Լ������ = 0 Then
              n_��Լʣ������ := 0;
            Else
              n_��Լʣ������ := n_��Լ������ - n_��Լ�ѹ���;
              If n_��Լʣ������ > (Nvl(r_�޺���, 0) - r_�ѹ���) Then
                n_��Լʣ������ := Nvl(r_�޺���, 0) - r_�ѹ���;
              End If;
            End If;
          End If;
        Else
          v_Temp := '<SPANLIST>';
          If r_�ƻ�id <> 0 Then
            Select Max(����ʱ��)
            Into d_�Ӻ�ʱ��
            From �Һżƻ�ʱ��
            Where �ƻ�id = r_�ƻ�id And ���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                '����', '6', '����', '7', '����', Null);
            If r_��ſ��� = 1 Then
              If Trunc(d_����) = Trunc(Sysdate) Then
                n_����ԤԼ := 0;
              Else
                Select Nvl(Max(Jh.�Ƿ�ԤԼ), 0)
                Into n_����ԤԼ
                From (Select Sd.�ƻ�id, Sd.���, Sd.����, Jh.����,
                              To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                              To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Sd.��������, Sd.�Ƿ�ԤԼ
                       From �ҺŰ��żƻ� Jh, �Һżƻ�ʱ�� Sd
                       Where Jh.Id = Sd.�ƻ�id And Jh.Id = r_�ƻ�id And
                             Sd.���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6',
                                            '����', '7', '����', Null)) Jh, �Һ����״̬ Zt
                Where Zt.����(+) = Jh.��ʼʱ�� And Zt.����(+) = Jh.���� And Zt.���(+) = Jh.��� And
                      Decode(Sign(Sysdate - Jh.��ʼʱ��), -1, 0, 1) <> 1;
              End If;
            
              For r_Time In (Select Jh.����, Jh.���, Jh.����, Jh.��ʼʱ��, Jh.����ʱ��, Jh.��������, Jh.�Ƿ�ԤԼ, 0 As ��Լ��,
                                    Decode(Nvl(Zt.���, 0), 0, 1, 0) As ʣ����,
                                    Decode(Sign(Sysdate - Jh.��ʼʱ��), -1, 0, 1) As ʧЧʱ��
                             
                             From (Select Sd.�ƻ�id, Sd.���, Sd.����, Jh.����,
                                           To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                           To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Sd.��������, Sd.�Ƿ�ԤԼ
                                    From �ҺŰ��żƻ� Jh, �Һżƻ�ʱ�� Sd
                                    Where Jh.Id = Sd.�ƻ�id And Jh.Id = r_�ƻ�id And
                                          Sd.���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                                         '5', '����', '6', '����', '7', '����', Null)) Jh, �Һ����״̬ Zt
                             Where Zt.����(+) = Jh.��ʼʱ�� And Zt.����(+) = Jh.���� And Zt.���(+) = Jh.��� And
                                   Decode(Sign(Sysdate - Jh.��ʼʱ��), -1, 0, 1) <> 1
                             Order By ���) Loop
                If v_������λ Is Not Null Then
                  Begin
                    Select 1
                    Into n_��Լģʽ
                    From ������λ�ƻ�����
                    Where ������Ŀ = r_Time.���� And �ƻ�id = r_�ƻ�id And ��� = 0 And ������λ = v_������λ;
                  Exception
                    When Others Then
                      n_��Լģʽ := 0;
                  End;
                Else
                  n_��Լģʽ := 0;
                End If;
                If r_Time.ʣ���� = 0 Then
                  n_����ʣ�� := 0;
                Else
                  n_����ʣ�� := r_Time.��������;
                End If;
                If v_������λ Is Null Or n_��Լģʽ = 1 Then
                  Begin
                    Select 1
                    Into n_Exists
                    From ������λ�ƻ�����
                    Where ������Ŀ = r_Time.���� And �ƻ�id = r_�ƻ�id And ��� = r_Time.��� And Rownum < 2;
                  Exception
                    When Others Then
                      n_Exists := 0;
                  End;
                  If n_Exists = 0 Then
                    If n_����ԤԼ = 1 And r_Time.�Ƿ�ԤԼ = 0 Then
                      Null;
                    Else
                      Begin
                        Select 1
                        Into n_�Ƿ�Ԥ��
                        From �Һ����״̬
                        Where ״̬ In (3, 4) And ���� = r_���� And ��� = r_Time.��� And Trunc(����) = Trunc(d_����);
                      Exception
                        When Others Then
                          n_�Ƿ�Ԥ�� := 0;
                      End;
                      If n_�Ƿ�Ԥ�� = 0 Then
                        v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                      To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                      '</SPAN>';
                        n_ʱ������ := Nvl(n_ʱ������, 0) + n_����ʣ��;
                      End If;
                    End If;
                  End If;
                Else
                  Begin
                    Select 1
                    Into n_Exists
                    From ������λ�ƻ�����
                    Where ������Ŀ = r_Time.���� And �ƻ�id = r_�ƻ�id And ��� = r_Time.��� And ������λ = v_������λ And Rownum < 2;
                  Exception
                    When Others Then
                      n_Exists := 0;
                  End;
                  Begin
                    Select 0
                    Into n_�Ǻ�Լ
                    From ������λ�ƻ�����
                    Where �ƻ�id = r_�ƻ�id And ������λ = v_������λ And Rownum < 2;
                  Exception
                    When Others Then
                      n_�Ǻ�Լ := 1;
                  End;
                  If n_Exists = 1 Or n_�Ǻ�Լ = 1 Then
                    If n_����ԤԼ = 1 And r_Time.�Ƿ�ԤԼ = 0 Then
                      Null;
                    Else
                      Begin
                        Select 1
                        Into n_�Ƿ�Ԥ��
                        From �Һ����״̬
                        Where ״̬ In (3, 4) And ���� = r_���� And ��� = r_Time.��� And Trunc(����) = Trunc(d_����);
                      Exception
                        When Others Then
                          n_�Ƿ�Ԥ�� := 0;
                      End;
                      If n_�Ƿ�Ԥ�� = 0 Then
                        v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                          To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                          '</SPAN>';
                        n_��Լʣ������ := n_��Լʣ������ + 1;
                      End If;
                    End If;
                  End If;
                End If;
              End Loop;
            Else
              n_���������� := Nvl(r_��Լ��, Nvl(r_�޺���, 0)) - Nvl(r_��Լ��, 0);
              For r_Time In (Select Jh.����, Jh.���, Jh.����, Jh.��ʼʱ��, Jh.����ʱ��, Jh.��������, Jh.�Ƿ�ԤԼ,
                                    Sum(Decode(Nvl(Zt.���, 0), 0, 0, 1)) As ��Լ��,
                                    Jh.�������� - Sum(Decode(Nvl(Zt.���, 0), 0, 0, 1)) As ʣ����,
                                    Decode(Sign(Sysdate - Jh.��ʼʱ��), -1, 0, 1) As ʧЧʱ��
                             From (Select Sd.�ƻ�id, Sd.���, Sd.����, Jh.����,
                                           To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                           To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Sd.��������, Sd.�Ƿ�ԤԼ
                                    From �ҺŰ��żƻ� Jh, �Һżƻ�ʱ�� Sd
                                    Where Jh.Id = Sd.�ƻ�id And Jh.Id = r_�ƻ�id And
                                          Sd.���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                                         '5', '����', '6', '����', '7', '����', Null)) Jh, �Һ����״̬ Zt
                             Where Jh.���� = Zt.����(+) And Jh.��ʼʱ�� = Zt.����(+) And
                                   Decode(Sign(Sysdate - Jh.��ʼʱ��), -1, 0, 1) <> 1
                             Group By Jh.����, Jh.���, Jh.����, Jh.��ʼʱ��, Jh.����ʱ��, Jh.��������, Jh.�Ƿ�ԤԼ
                             Order By Jh.���) Loop
                If v_������λ Is Not Null Then
                  Begin
                    Select 1
                    Into n_��Լģʽ
                    From ������λ�ƻ�����
                    Where ������Ŀ = r_Time.���� And �ƻ�id = r_�ƻ�id And ��� = 0 And ������λ = v_������λ;
                  Exception
                    When Others Then
                      n_��Լģʽ := 0;
                  End;
                Else
                  n_��Լģʽ := 0;
                End If;
                n_����ʣ�� := r_Time.ʣ����;
                If v_������λ Is Null Or n_��Լģʽ = 1 Then
                  Begin
                    Select 1
                    Into n_Exists
                    From ������λ�ƻ�����
                    Where ������Ŀ = r_Time.���� And �ƻ�id = r_�ƻ�id And ��� = r_Time.��� And Rownum < 2;
                  Exception
                    When Others Then
                      n_Exists := 0;
                  End;
                  If n_Exists = 0 Then
                    If n_���������� < n_����ʣ�� Then
                      v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                    To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_���������� || '</SL>' ||
                                    '</SPAN>';
                      n_ʱ������ := Nvl(n_ʱ������, 0) + n_����������;
                    Else
                      v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                    To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                    '</SPAN>';
                      n_ʱ������ := Nvl(n_ʱ������, 0) + n_����ʣ��;
                    End If;
                  End If;
                Else
                  Begin
                    Select 1
                    Into n_Exists
                    From ������λ�ƻ�����
                    Where ������Ŀ = r_Time.���� And �ƻ�id = r_�ƻ�id And ��� = r_Time.��� And ������λ = v_������λ And Rownum < 2;
                  Exception
                    When Others Then
                      n_Exists := 0;
                  End;
                  Begin
                    Select 0
                    Into n_�Ǻ�Լ
                    From ������λ�ƻ�����
                    Where �ƻ�id = r_�ƻ�id And ������λ = v_������λ And Rownum < 2;
                  Exception
                    When Others Then
                      n_�Ǻ�Լ := 1;
                  End;
                  If n_Exists = 1 Or n_�Ǻ�Լ = 1 Then
                    If n_���������� < n_����ʣ�� Then
                      v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                        To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_���������� || '</SL>' ||
                                        '</SPAN>';
                      n_��Լʣ������ := n_��Լʣ������ + Nvl(n_����ʣ��, 0);
                    Else
                      v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                        To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                        '</SPAN>';
                      n_��Լʣ������ := n_��Լʣ������ + Nvl(n_����ʣ��, 0);
                    End If;
                  End If;
                End If;
              End Loop;
            End If;
          Else
            Select Max(����ʱ��)
            Into d_�Ӻ�ʱ��
            From �ҺŰ���ʱ��
            Where ����id = r_����id And ���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                '����', '6', '����', '7', '����', Null);
            If r_��ſ��� = 1 Then
              If Trunc(d_����) = Trunc(Sysdate) Then
                n_����ԤԼ := 0;
              Else
                Select Nvl(Max(Ap.�Ƿ�ԤԼ), 0)
                Into n_����ԤԼ
                From (Select Sd.����id, Sd.���, Sd.����, Ap.����,
                              To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                              To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Sd.��������, Sd.�Ƿ�ԤԼ
                       From �ҺŰ��� Ap, �ҺŰ���ʱ�� Sd
                       Where Ap.Id = Sd.����id And Ap.Id = r_����id And
                             Sd.���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6',
                                            '����', '7', '����', Null)) Ap, �Һ����״̬ Zt
                Where Zt.����(+) = Ap.��ʼʱ�� And Zt.����(+) = Ap.���� And Zt.���(+) = Ap.��� And
                      Decode(Sign(Sysdate - Ap.��ʼʱ��), -1, 0, 1) <> 1;
              End If;
              For r_Time In (Select Ap.����, Ap.���, Ap.����, Ap.��ʼʱ��, Ap.����ʱ��, Ap.��������, Ap.�Ƿ�ԤԼ, 0 As ��Լ��,
                                    Decode(Nvl(Zt.���, 0), 0, 1, 0) As ʣ����,
                                    Decode(Sign(Sysdate - Ap.��ʼʱ��), -1, 0, 1) As ʧЧʱ��
                             
                             From (Select Sd.����id, Sd.���, Sd.����, Ap.����,
                                           To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                           To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Sd.��������, Sd.�Ƿ�ԤԼ
                                    From �ҺŰ��� Ap, �ҺŰ���ʱ�� Sd
                                    Where Ap.Id = Sd.����id And Ap.Id = r_����id And
                                          Sd.���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                                         '5', '����', '6', '����', '7', '����', Null)) Ap, �Һ����״̬ Zt
                             Where Zt.����(+) = Ap.��ʼʱ�� And Zt.����(+) = Ap.���� And Zt.���(+) = Ap.��� And
                                   Decode(Sign(Sysdate - Ap.��ʼʱ��), -1, 0, 1) <> 1
                             Order By ���) Loop
                If v_������λ Is Not Null Then
                  Begin
                    Select 1
                    Into n_��Լģʽ
                    From ������λ���ſ���
                    Where ������Ŀ = r_Time.���� And ����id = r_����id And ��� = 0 And ������λ = v_������λ;
                  Exception
                    When Others Then
                      n_��Լģʽ := 0;
                  End;
                Else
                  n_��Լģʽ := 0;
                End If;
                If r_Time.ʣ���� = 0 Then
                  n_����ʣ�� := 0;
                Else
                  n_����ʣ�� := r_Time.��������;
                End If;
                If v_������λ Is Null Or n_��Լģʽ = 1 Then
                  Begin
                    Select 1
                    Into n_Exists
                    From ������λ���ſ���
                    Where ������Ŀ = r_Time.���� And ����id = r_����id And ��� = r_Time.��� And Rownum < 2;
                  Exception
                    When Others Then
                      n_Exists := 0;
                  End;
                  If n_Exists = 0 Then
                    If n_����ԤԼ = 1 And r_Time.�Ƿ�ԤԼ = 0 Then
                      Null;
                    Else
                      Begin
                        Select 1
                        Into n_�Ƿ�Ԥ��
                        From �Һ����״̬
                        Where ״̬ In (3, 4) And ���� = r_���� And ��� = r_Time.��� And Trunc(����) = Trunc(d_����);
                      Exception
                        When Others Then
                          n_�Ƿ�Ԥ�� := 0;
                      End;
                      If n_�Ƿ�Ԥ�� = 0 Then
                        v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                      To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                      '</SPAN>';
                        n_ʱ������ := Nvl(n_ʱ������, 0) + n_����ʣ��;
                      End If;
                    End If;
                  End If;
                Else
                  Begin
                    Select 1
                    Into n_Exists
                    From ������λ���ſ���
                    Where ������Ŀ = r_Time.���� And ����id = r_����id And ��� = r_Time.��� And ������λ = v_������λ And Rownum < 2;
                  Exception
                    When Others Then
                      n_Exists := 0;
                  End;
                  Begin
                    Select 0
                    Into n_�Ǻ�Լ
                    From ������λ���ſ���
                    Where ����id = r_����id And ������λ = v_������λ And Rownum < 2;
                  Exception
                    When Others Then
                      n_�Ǻ�Լ := 1;
                  End;
                  If n_Exists = 1 Or n_�Ǻ�Լ = 1 Then
                    If n_����ԤԼ = 1 And r_Time.�Ƿ�ԤԼ = 0 Then
                      Null;
                    Else
                      Begin
                        Select 1
                        Into n_�Ƿ�Ԥ��
                        From �Һ����״̬
                        Where ״̬ In (3, 4) And ���� = r_���� And ��� = r_Time.��� And Trunc(����) = Trunc(d_����);
                      Exception
                        When Others Then
                          n_�Ƿ�Ԥ�� := 0;
                      End;
                      If n_�Ƿ�Ԥ�� = 0 Then
                        v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                          To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                          '</SPAN>';
                        n_��Լʣ������ := n_��Լʣ������ + 1;
                      End If;
                    End If;
                  End If;
                End If;
              End Loop;
            Else
              n_���������� := Nvl(r_��Լ��, Nvl(r_�޺���, 0)) - Nvl(r_��Լ��, 0);
              For r_Time In (Select Ap.����, Ap.���, Ap.����, Ap.��ʼʱ��, Ap.����ʱ��, Ap.��������, Ap.�Ƿ�ԤԼ,
                                    Sum(Decode(Nvl(Zt.���, 0), 0, 0, 1)) As ��Լ��,
                                    Ap.�������� - Sum(Decode(Nvl(Zt.���, 0), 0, 0, 1)) As ʣ����,
                                    Decode(Sign(Sysdate - Ap.��ʼʱ��), -1, 0, 1) As ʧЧʱ��
                             From (Select Sd.����id, Sd.���, Sd.����, Ap.����,
                                           To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                           To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Sd.��������, Sd.�Ƿ�ԤԼ
                                    From �ҺŰ��� Ap, �ҺŰ���ʱ�� Sd
                                    Where Ap.Id = Sd.����id And Ap.Id = r_����id And
                                          Sd.���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                                         '5', '����', '6', '����', '7', '����', Null)) Ap, �Һ����״̬ Zt
                             Where Ap.���� = Zt.����(+) And Ap.��ʼʱ�� = Zt.����(+) And
                                   Decode(Sign(Sysdate - Ap.��ʼʱ��), -1, 0, 1) <> 1
                             Group By Ap.����, Ap.���, Ap.����, Ap.��ʼʱ��, Ap.����ʱ��, Ap.��������, Ap.�Ƿ�ԤԼ
                             Order By Ap.���) Loop
                If v_������λ Is Not Null Then
                  Begin
                    Select 1
                    Into n_��Լģʽ
                    From ������λ���ſ���
                    Where ������Ŀ = r_Time.���� And ����id = r_����id And ��� = 0 And ������λ = v_������λ;
                  Exception
                    When Others Then
                      n_��Լģʽ := 0;
                  End;
                Else
                  n_��Լģʽ := 0;
                End If;
                n_����ʣ�� := r_Time.ʣ����;
                If v_������λ Is Null Or n_��Լģʽ = 1 Then
                  Begin
                    Select 1
                    Into n_Exists
                    From ������λ���ſ���
                    Where ������Ŀ = r_Time.���� And ����id = r_����id And ��� = r_Time.��� And Rownum < 2;
                  Exception
                    When Others Then
                      n_Exists := 0;
                  End;
                  If n_Exists = 0 Then
                    If n_���������� < n_����ʣ�� Then
                      v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                    To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_���������� || '</SL>' ||
                                    '</SPAN>';
                      n_ʱ������ := Nvl(n_ʱ������, 0) + n_����������;
                    Else
                      v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                    To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                    '</SPAN>';
                      n_ʱ������ := Nvl(n_ʱ������, 0) + n_����ʣ��;
                    End If;
                  End If;
                Else
                  Begin
                    Select 1
                    Into n_Exists
                    From ������λ���ſ���
                    Where ������Ŀ = r_Time.���� And ����id = r_����id And ��� = r_Time.��� And ������λ = v_������λ And Rownum < 2;
                  Exception
                    When Others Then
                      n_Exists := 0;
                  End;
                  Begin
                    Select 0
                    Into n_�Ǻ�Լ
                    From ������λ���ſ���
                    Where ����id = r_����id And ������λ = v_������λ And Rownum < 2;
                  Exception
                    When Others Then
                      n_�Ǻ�Լ := 1;
                  End;
                  If n_Exists = 1 Or n_�Ǻ�Լ = 1 Then
                    If n_���������� < n_����ʣ�� Then
                      v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                        To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_���������� || '</SL>' ||
                                        '</SPAN>';
                      n_��Լʣ������ := n_��Լʣ������ + Nvl(n_����ʣ��, 0);
                    Else
                      v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                        To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                        '</SPAN>';
                      n_��Լʣ������ := n_��Լʣ������ + Nvl(n_����ʣ��, 0);
                    End If;
                  End If;
                End If;
              End Loop;
            End If;
          End If;
        End If;
        If v_������λ Is Not Null Then
          If Nvl(r_�ƻ�id, 0) <> 0 Then
            Begin
              Select 0
              Into n_�Ǻ�Լ
              From ������λ�ƻ�����
              Where �ƻ�id = r_�ƻ�id And ������λ = v_������λ And Rownum < 2;
            Exception
              When Others Then
                n_�Ǻ�Լ := 1;
            End;
          Else
            Begin
              Select 0
              Into n_�Ǻ�Լ
              From ������λ���ſ���
              Where ����id = r_����id And ������λ = v_������λ And Rownum < 2;
            Exception
              When Others Then
                n_�Ǻ�Լ := 1;
            End;
          End If;
        End If;
        If v_������λ Is Null Or n_�Ǻ�Լ = 1 Then
          If r_�޺��� = 0 Then
            v_ʣ������ := '';
          Else
            If Nvl(r_�ƻ�id, 0) <> 0 Then
              Select Sum(����)
              Into n_��Լ������
              From ������λ�ƻ�����
              Where �ƻ�id = r_�ƻ�id And ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                    '����', '6', '����', '7', '����', Null);
            Else
              Select Sum(����)
              Into n_��Լ������
              From ������λ���ſ���
              Where ����id = r_����id And ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                    '����', '6', '����', '7', '����', Null);
            End If;
            Begin
              Select Count(1)
              Into n_��Լ�ѹ���
              From ���˹Һż�¼
              Where �ű� = r_���� And ��¼״̬ = 1 And ������λ Is Not Null And ����ʱ�� Between Trunc(d_����) And
                    Trunc(d_���� + 1) - 1 / 60 / 60 / 24;
            Exception
              When Others Then
                n_��Լ�ѹ��� := 0;
            End;
            Select Count(1)
            Into n_Ԥ������
            From �Һ����״̬
            Where ״̬ = 3 And ���� = r_���� And Trunc(����) = Trunc(d_����);
            If Trunc(d_����) = Trunc(Sysdate) Then
              If Nvl(n_��Լ������, 0) = 0 Then
                v_ʣ������ := r_�޺��� - r_�ѹ��� - r_��Լ�� + r_�ѽ��� - n_Ԥ������;
              Else
                v_ʣ������ := r_�޺��� - r_�ѹ��� - r_��Լ�� + r_�ѽ��� - n_��Լ������ + n_��Լ�ѹ��� - n_Ԥ������;
              End If;
              n_�ѹ��� := r_�ѹ���;
              If Nvl(n_ʱ������, 0) < v_ʣ������ And n_��ʱ�� <> 0 Then
                n_������� := 1;
                v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_�Ӻ�ʱ��, 'hh24:mi:ss') || '-' || '</SJD>' ||
                              '<SL>' || To_Number(v_ʣ������ - Nvl(n_ʱ������, 0)) || '</SL>' || '</SPAN>';
              Else
                n_������� := 0;
              End If;
            Else
              If Nvl(n_��Լ������, 0) = 0 Then
                v_ʣ������ := r_��Լ�� - r_��Լ�� - n_Ԥ������;
                If v_ʣ������ Is Null Then
                  v_ʣ������ := r_�޺��� - r_�ѹ��� - r_��Լ�� + r_�ѽ��� - n_Ԥ������;
                End If;
              Else
                v_ʣ������ := r_��Լ�� - r_��Լ�� - n_��Լ������ + n_��Լ�ѹ��� - n_Ԥ������;
                If v_ʣ������ Is Null Then
                  v_ʣ������ := r_�޺��� - r_�ѹ��� - r_��Լ�� + r_�ѽ��� - n_��Լ������ + n_��Լ�ѹ��� - n_Ԥ������;
                End If;
              End If;
              n_�ѹ��� := r_�ѹ���;
            End If;
          End If;
        Else
          If Nvl(r_�ƻ�id, 0) <> 0 Then
            If v_������λ Is Not Null Then
              Begin
                Select 1
                Into n_��Լģʽ
                From ������λ�ƻ�����
                Where ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                                    '7', '����', Null) And �ƻ�id = r_�ƻ�id And ��� = 0 And ������λ = v_������λ;
              Exception
                When Others Then
                  n_��Լģʽ := 0;
              End;
            Else
              n_��Լģʽ := 0;
            End If;
            Select Sum(����)
            Into n_��Լ������
            From ������λ�ƻ�����
            Where �ƻ�id = r_�ƻ�id And ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                  '����', '6', '����', '7', '����', Null) And ������λ = v_������λ;
          Else
            If v_������λ Is Not Null Then
              Begin
                Select 1
                Into n_��Լģʽ
                From ������λ���ſ���
                Where ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                                    '7', '����', Null) And ����id = r_����id And ��� = 0 And ������λ = v_������λ;
              Exception
                When Others Then
                  n_��Լģʽ := 0;
              End;
            Else
              n_��Լģʽ := 0;
            End If;
            Select Sum(����)
            Into n_��Լ������
            From ������λ���ſ���
            Where ����id = r_����id And ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                  '����', '6', '����', '7', '����', Null) And ������λ = v_������λ;
          End If;
          If n_��Լģʽ = 0 Then
            v_ʣ������   := n_��Լʣ������;
            n_�ѹ���     := r_�ѹ���;
            n_��Լ�ѹ��� := Nvl(n_��Լ������, 0) - n_��Լʣ������;
          Else
            n_�ѹ��� := r_�ѹ���;
            Begin
              Select Count(1)
              Into n_��Լ�ѹ���
              From ���˹Һż�¼
              Where �ű� = r_���� And ��¼״̬ = 1 And ������λ = v_������λ And ����ʱ�� Between Trunc(d_����) And
                    Trunc(d_���� + 1) - 1 / 60 / 60 / 24;
            Exception
              When Others Then
                n_��Լ�ѹ��� := 0;
            End;
            If Nvl(n_��Լ������, 0) = 0 Then
              v_ʣ������ := '0';
            Else
              v_ʣ������ := n_��Լ������ - n_��Լ�ѹ���;
            End If;
          End If;
        End If;
        Select To_Char(��ʼʱ��, 'hh24:mi') Into v_Timetemp From ʱ��� Where ʱ��� = r_�Ű�;
        v_ʱ��� := v_Timetemp || '-';
        Select To_Char(��ֹʱ��, 'hh24:mi') Into v_Timetemp From ʱ��� Where ʱ��� = r_�Ű�;
        v_ʱ��� := v_ʱ��� || v_Timetemp;
        If v_Temp Is Not Null Then
          v_Temp := v_Temp || '</SPANLIST>';
        End If;
        If v_������λ Is Not Null Then
          If Nvl(r_�ƻ�id, 0) <> 0 Then
            Begin
              Select 1
              Into n_����
              From ������λ�ƻ�����
              Where �ƻ�id = r_�ƻ�id And ������λ = v_������λ And ���� = 0 And Rownum < 2;
            Exception
              When Others Then
                n_���� := 0;
            End;
          Else
            Begin
              Select 1
              Into n_����
              From ������λ���ſ���
              Where ����id = r_����id And ������λ = v_������λ And ���� = 0 And Rownum < 2;
            Exception
              When Others Then
                n_���� := 0;
            End;
          End If;
        End If;
        --��Լ��=0��ԤԼ��ֹ
        If Trunc(d_����) <> Trunc(Sysdate) Then
          If r_��Լ�� = 0 Then
            n_���� := 1;
          End If;
        End If;
        If Nvl(n_����, 0) = 0 Then
          --���������
          n_�ϼƽ�� := r_�۸�;
          For r_Subfee In (Select �ּ�, ��������
                           From �շѴ�����Ŀ A, �շѼ�Ŀ B
                           Where a.����id = r_��Ŀid And a.����id = b.�շ�ϸĿid And Sysdate Between b.ִ������ And
                                 Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))) Loop
            n_�ϼƽ�� := n_�ϼƽ�� + r_Subfee.�ּ� * r_Subfee.��������;
          End Loop;
          If Trunc(Sysdate) = Trunc(d_����) Then
            Begin
              Select 1
              Into n_Exists
              From (Select ʱ���
                     From ʱ���
                     Where ('3000-01-10 ' || To_Char(Sysdate, 'HH24:MI:SS') <
                           '3000-01-10 ' || To_Char(��ֹʱ��, 'HH24:MI:SS')) Or
                           ('3000-01-10 ' || To_Char(Sysdate, 'HH24:MI:SS') <
                           Decode(Sign(��ʼʱ�� - ��ֹʱ��), 1, '3000-01-11 ' || To_Char(��ֹʱ��, 'HH24:MI:SS'),
                                   '3000-01-10 ' || To_Char(��ֹʱ��, 'HH24:MI:SS'))))
              Where ʱ��� = r_�Ű�;
            Exception
              When Others Then
                n_Exists := 0;
            End;
          Else
            n_Exists := 1;
          End If;
          If n_Exists = 1 Then
            If v_ʣ������ > 0 Then
              c_Xmlmain := '<HB>' || '<HM>' || r_���� || '</HM>' || '<YSID>' || r_ҽ��id || '</YSID>' || '<YS>' || r_ҽ������ ||
                           '</YS>' || '<KSID>' || r_����id || '</KSID>' || '<KSMC>' || r_�������� || '</KSMC>' || '<ZC>' || r_ְ�� ||
                           '</ZC>' || '<XMID>' || r_��Ŀid || '</XMID>' || '<XMMC>' || r_��Ŀ���� || '</XMMC>' || '<YGHS>' ||
                           n_�ѹ��� || '</YGHS>' || '<SYHS>' || v_ʣ������ || '</SYHS>' || '<PRICE>' || n_�ϼƽ�� || '</PRICE>' ||
                           '<HCXH>' || n_������� || '</HCXH>' || '<HL>' || r_���� || '</HL>' || '<FSD>' || n_��ʱ�� || '</FSD>' ||
                           '<HBTIME>' || v_ʱ��� || '</HBTIME>' || '<FWMC>' || r_�Ű� || '</FWMC>' || v_Temp || '</HB>';
              v_Xmlmain := v_Xmlmain || c_Xmlmain;
            Else
              c_Xmlmain := '<HB>' || '<HM>' || r_���� || '</HM>' || '<YSID>' || r_ҽ��id || '</YSID>' || '<YS>' || r_ҽ������ ||
                           '</YS>' || '<KSID>' || r_����id || '</KSID>' || '<KSMC>' || r_�������� || '</KSMC>' || '<ZC>' || r_ְ�� ||
                           '</ZC>' || '<XMID>' || r_��Ŀid || '</XMID>' || '<XMMC>' || r_��Ŀ���� || '</XMMC>' || '<YGHS>' ||
                           n_�ѹ��� || '</YGHS>' || '<SYHS>' || v_ʣ������ || '</SYHS>' || '<PRICE>' || n_�ϼƽ�� || '</PRICE>' ||
                           '<HL>' || r_���� || '</HL>' || '<FSD>' || n_��ʱ�� || '</FSD>' || '<HBTIME>' || v_ʱ��� ||
                           '</HBTIME>' || '<FWMC>' || r_�Ű� || '</FWMC>' || '</HB>';
              v_Xmlmain := v_Xmlmain || c_Xmlmain;
            End If;
          End If;
        End If;
        n_��Լʣ������ := 0;
        n_��Լ������   := 0;
        n_ʱ������     := 0;
        n_����         := 0;
        n_�Ǻ�Լ       := 0;
      End Loop;
      Close r_No;
      v_Xmlmain := '<GROUP>' || '<RQ>' || To_Char(d_����, 'yyyy-mm-dd') || '</RQ>' || '<HBLIST>' || v_Xmlmain ||
                   '</HBLIST>' || '</GROUP>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Xmlmain)) Into x_Templet From Dual;
    Else
      --������Ű�ģʽ
      n_��Լʣ������ := 0;
      v_Sql          := 'Select a.����id, b.����, c.���� As ��������, a.ҽ������, a.ҽ��id, d.רҵ����ְ�� As ְ��, b.����, a.Id As ��¼id, a.�ϰ�ʱ��, a.��Ŀid, e.���� As ��Ŀ����, ';
      v_Sql          := v_Sql ||
                        'a.�Ƿ���ſ��� As ��ſ���, a.�޺���, Nvl(a.��Լ��,a.�޺���) As ��Լ��, a.�ѹ���, a.��Լ��, a.�����ѽ��� As �ѽ���, a.�Ƿ��ʱ�� As ��ʱ��, a.��ʼʱ��, a.��ֹʱ��, a.ԤԼ����   ';
      v_Sql          := v_Sql || 'From �ٴ������¼ A, �ٴ������Դ B, ���ű� C, ��Ա�� D, �շ���ĿĿ¼ E ';
      v_Sql          := v_Sql ||
                        'Where a.�������� = Trunc(:1) And a.��Դid = b.Id And a.��Ŀid = e.Id And a.ҽ��id = d.Id(+) And b.����id = c.Id And Nvl(a.�Ƿ�����, 0) = 0 And ';
      v_Sql          := v_Sql || '      (a.��ʼʱ�� < Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��) Or a.��ֹʱ�� > Nvl(a.ͣ����ֹʱ��, a.��ʼʱ��)) ';
      v_Sql          := v_Sql || '      And Nvl(a.�Ƿ񷢲�,0) = 1 And a.��ʼʱ�� > To_Date( ' || Chr(39) || v_����ʱ�� || Chr(39) || ',' ||
                        Chr(39) || 'yyyy-mm-dd hh24:mi:ss' || Chr(39) || ')';
      n_Curcount     := 2;
      If Nvl(n_����id, 0) <> 0 Then
        v_Sql      := v_Sql || 'And b.����id = :' || n_Curcount || ' ';
        n_Curcount := n_Curcount + 1;
      End If;
      If Nvl(n_ҽ��id, 0) <> 0 Then
        v_Sql      := v_Sql || 'And a.ҽ��id = :' || n_Curcount || ' ';
        n_Curcount := n_Curcount + 1;
      End If;
      If Nvl(v_ҽ������, '_') <> '_' Then
        v_Sql      := v_Sql || 'And a.ҽ������ = :' || n_Curcount || ' ';
        n_Curcount := n_Curcount + 1;
      End If;
    
      If Nvl(n_����id, 0) <> 0 And Nvl(n_ҽ��id, 0) <> 0 And Nvl(v_ҽ������, '_') <> '_' Then
        Open r_No For v_Sql
          Using d_����, n_����id, n_ҽ��id, v_ҽ������;
      End If;
      If Nvl(n_����id, 0) <> 0 And Nvl(n_ҽ��id, 0) = 0 And Nvl(v_ҽ������, '_') = '_' Then
        Open r_No For v_Sql
          Using d_����, n_����id;
      End If;
      If Nvl(n_����id, 0) = 0 And Nvl(n_ҽ��id, 0) <> 0 And Nvl(v_ҽ������, '_') = '_' Then
        Open r_No For v_Sql
          Using d_����, n_ҽ��id;
      End If;
      If Nvl(n_����id, 0) = 0 And Nvl(n_ҽ��id, 0) = 0 And Nvl(v_ҽ������, '_') <> '_' Then
        Open r_No For v_Sql
          Using d_����, v_ҽ������;
      End If;
      If Nvl(n_����id, 0) <> 0 And Nvl(n_ҽ��id, 0) <> 0 And Nvl(v_ҽ������, '_') = '_' Then
        Open r_No For v_Sql
          Using d_����, n_����id, n_ҽ��id;
      End If;
      If Nvl(n_����id, 0) <> 0 And Nvl(n_ҽ��id, 0) = 0 And Nvl(v_ҽ������, '_') <> '_' Then
        Open r_No For v_Sql
          Using d_����, n_����id, v_ҽ������;
      End If;
      If Nvl(n_����id, 0) = 0 And Nvl(n_ҽ��id, 0) <> 0 And Nvl(v_ҽ������, '_') <> '_' Then
        Open r_No For v_Sql
          Using d_����, n_ҽ��id, v_ҽ������;
      End If;
      Loop
        Fetch r_No
          Into r_����id, r_����, r_��������, r_ҽ������, r_ҽ��id, r_ְ��, r_����, r_����id, r_�Ű�, r_��Ŀid, r_��Ŀ����, r_��ſ���, r_�޺���, r_��Լ��,
               r_�ѹ���, r_��Լ��, r_�ѽ���, r_��ʱ��, r_��ʼʱ��, r_��ֹʱ��, r_ԤԼ����;
        Exit When r_No%NotFound;
        If Trunc(d_����) = Trunc(Sysdate) Then
          --����Һ�
          If v_������λ Is Null Then
            --δ���������λ
            n_�ѹ���   := r_�ѹ���;
            v_ʣ������ := r_�޺��� - Nvl(r_�ѹ���, 0);
            If r_��ʱ�� = 1 And r_��ſ��� = 1 Then
              --��ʱ��
              v_Temp   := '<SPANLIST>';
              n_Exists := 0;
              For r_Time In (Select ���, ��ʼʱ��, ��ֹʱ��, �Һ�״̬ From �ٴ�������ſ��� Where ��¼id = r_����id) Loop
                v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                          To_Char(r_Time.��ֹʱ��, 'hh24:mi:ss') || '</SJD>';
                If Nvl(r_Time.�Һ�״̬, 0) = 0 Then
                  v_Temp   := v_Temp || '<SL>1</SL></SPAN>';
                  n_Exists := n_Exists + 1;
                Else
                  v_Temp := v_Temp || '<SL>0</SL></SPAN>';
                End If;
              End Loop;
              If n_Exists < To_Number(v_ʣ������) Then
                Select Max(��ֹʱ��) Into d_�Ӻ�ʱ�� From �ٴ�������ſ��� Where ��¼id = r_����id;
                v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_�Ӻ�ʱ��, 'hh24:mi:ss') || '-' || '</SJD><SL>' ||
                          v_ʣ������ - n_Exists || '</SL></SPAN>';
              End If;
              v_Temp := v_Temp || '</SPANLIST>';
            End If;
          Else
            --���������λ
            n_�ѹ��� := r_�ѹ���;
            Begin
              Select ���Ʒ�ʽ
              Into n_��Լģʽ
              From �ٴ�����Һſ��Ƽ�¼
              Where ��¼id = r_����id And ���� = 1 And ���� = v_������λ And ���� = 1 And Rownum < 2;
            Exception
              When Others Then
                n_��Լģʽ := 4;
            End;
            If n_��Լģʽ = 0 Then
              n_���� := 1;
            End If;
            If n_��Լģʽ = 1 Or n_��Լģʽ = 2 Then
              Select ����
              Into n_��Լ������
              From �ٴ�����Һſ��Ƽ�¼
              Where ��¼id = r_����id And ���� = 1 And ���� = v_������λ And ���� = 1;
              If n_��Լģʽ = 1 Then
                n_��Լ������ := Round(r_�޺��� * n_��Լ������ / 100);
              End If;
              Begin
                Select Count(1)
                Into n_��Լ�ѹ���
                From ���˹Һż�¼
                Where �ű� = r_���� And ��¼״̬ = 1 And ������λ = v_������λ And ����ʱ�� Between Trunc(d_����) And
                      Trunc(d_���� + 1) - 1 / 60 / 60 / 24;
              Exception
                When Others Then
                  n_��Լ�ѹ��� := 0;
              End;
              n_��Լʣ������ := n_��Լ������ - n_��Լ�ѹ���;
              If r_�޺��� - Nvl(r_�ѹ���, 0) < n_��Լʣ������ Then
                v_ʣ������ := r_�޺��� - Nvl(r_�ѹ���, 0);
              Else
                v_ʣ������ := n_��Լʣ������;
              End If;
              If r_��ʱ�� = 1 And r_��ſ��� = 1 Then
                --��ʱ��
                v_Temp   := '<SPANLIST>';
                n_Exists := 0;
                For r_Time In (Select ���, ��ʼʱ��, ��ֹʱ��, �Һ�״̬ From �ٴ�������ſ��� Where ��¼id = r_����id) Loop
                  v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                            To_Char(r_Time.��ֹʱ��, 'hh24:mi:ss') || '</SJD>';
                  If Nvl(r_Time.�Һ�״̬, 0) = 0 Then
                    v_Temp   := v_Temp || '<SL>1</SL></SPAN>';
                    n_Exists := n_Exists + 1;
                  Else
                    v_Temp := v_Temp || '<SL>0</SL></SPAN>';
                  End If;
                End Loop;
                If n_Exists < To_Number(v_ʣ������) Then
                  Select Max(��ֹʱ��) Into d_�Ӻ�ʱ�� From �ٴ�������ſ��� Where ��¼id = r_����id;
                  v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_�Ӻ�ʱ��, 'hh24:mi:ss') || '-' || '</SJD><SL>' ||
                            v_ʣ������ - n_Exists || '</SL></SPAN>';
                End If;
                v_Temp := v_Temp || '</SPANLIST>';
              End If;
            End If;
            If n_��Լģʽ = 3 Then
              If n_��ſ��� = 0 Then
                n_�ѹ���   := r_�ѹ���;
                v_ʣ������ := r_�޺��� - Nvl(r_�ѹ���, 0);
              Else
                n_�ѹ���   := 0;
                v_ʣ������ := 0;
                For r_���� In (Select ���
                             From �ٴ�����Һſ��Ƽ�¼
                             Where ��¼id = r_����id And ���� = 1 And ���� = 1 And ���� = v_������λ) Loop
                  Begin
                    Select 1, ��ʼʱ��, ��ֹʱ��
                    Into n_Exists, d_��ʼʱ��, d_��ֹʱ��
                    From �ٴ�������ſ���
                    Where ��� = r_����.��� And Nvl(�Һ�״̬, 0) = 0;
                  Exception
                    When Others Then
                      n_Exists := 0;
                  End;
                  If n_Exists = 1 Then
                    v_ʣ������ := v_ʣ������ + 1;
                    If r_��ʱ�� = 1 Then
                      v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                To_Char(d_��ֹʱ��, 'hh24:mi:ss') || '</SJD><SL>1</SL></SPAN>';
                    End If;
                  Else
                    n_�ѹ��� := n_�ѹ��� + 1;
                  End If;
                End Loop;
                If v_Temp Is Not Null Then
                  v_Temp := '<SPANLIST>' || v_Temp || '</SPANLIST>';
                End If;
              End If;
            End If;
            If n_��Լģʽ = 4 Then
              n_�ѹ���   := r_�ѹ���;
              v_ʣ������ := r_�޺��� - Nvl(r_�ѹ���, 0);
              If r_��ʱ�� = 1 And r_��ſ��� = 1 Then
                --��ʱ��
                v_Temp   := '<SPANLIST>';
                n_Exists := 0;
                For r_Time In (Select ���, ��ʼʱ��, ��ֹʱ��, �Һ�״̬ From �ٴ�������ſ��� Where ��¼id = r_����id) Loop
                  v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                            To_Char(r_Time.��ֹʱ��, 'hh24:mi:ss') || '</SJD>';
                  If Nvl(r_Time.�Һ�״̬, 0) = 0 Then
                    v_Temp   := v_Temp || '<SL>1</SL></SPAN>';
                    n_Exists := n_Exists + 1;
                  Else
                    v_Temp := v_Temp || '<SL>0</SL></SPAN>';
                  End If;
                End Loop;
                If n_Exists < To_Number(v_ʣ������) Then
                  Select Max(��ֹʱ��) Into d_�Ӻ�ʱ�� From �ٴ�������ſ��� Where ��¼id = r_����id;
                  v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_�Ӻ�ʱ��, 'hh24:mi:ss') || '-' || '</SJD><SL>' ||
                            v_ʣ������ - n_Exists || '</SL></SPAN>';
                End If;
                v_Temp := v_Temp || '</SPANLIST>';
              End If;
            End If;
          End If;
        Else
          --ԤԼ�Һ�
          If r_ԤԼ���� = 1 Then
            n_���� := 1;
          Else
            --������ԤԼ
            If v_������λ Is Null Then
              If r_��ʱ�� = 0 Then
                n_�ѹ���   := r_��Լ��;
                v_ʣ������ := Nvl(r_��Լ��, r_�޺���) - Nvl(r_��Լ��, 0);
              Else
                --��ʱ��
                n_�ѹ���   := 0;
                v_ʣ������ := 0;
                v_Temp     := '<SPANLIST>';
                If r_��ſ��� = 0 Then
                  --����ſ��Ʒ�ʱ��ԤԼ
                  For r_Time In (Select ���, ��ʼʱ��, ��ֹʱ��, �Һ�״̬, ����
                                 From �ٴ�������ſ���
                                 Where ��¼id = r_����id And ԤԼ˳��� Is Null And �Ƿ�ԤԼ = 1) Loop
                    Select Count(1)
                    Into n_ʱ���ѹ�
                    From �ٴ�������ſ���
                    Where ԤԼ˳��� Is Not Null And Nvl(�Һ�״̬, 0) <> 0;
                    v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.��ֹʱ��, 'hh24:mi:ss') || '</SJD>';
                    v_Temp     := v_Temp || '<SL>' || r_Time.���� - n_ʱ���ѹ� || '</SL></SPAN>';
                    n_�ѹ���   := n_�ѹ��� + n_ʱ���ѹ�;
                    v_ʣ������ := v_ʣ������ + (r_Time.���� - n_ʱ���ѹ�);
                  End Loop;
                Else
                  For r_Time In (Select ���, ��ʼʱ��, ��ֹʱ��, �Һ�״̬ From �ٴ�������ſ��� Where ��¼id = r_����id) Loop
                    v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                              To_Char(r_Time.��ֹʱ��, 'hh24:mi:ss') || '</SJD>';
                    If Nvl(r_Time.�Һ�״̬, 0) = 0 Then
                      v_Temp     := v_Temp || '<SL>1</SL></SPAN>';
                      v_ʣ������ := v_ʣ������ + 1;
                    Else
                      v_Temp   := v_Temp || '<SL>0</SL></SPAN>';
                      n_�ѹ��� := n_�ѹ��� + 1;
                    End If;
                  End Loop;
                End If;
                v_Temp := v_Temp || '</SPANLIST>';
              End If;
            Else
              --������λԤԼ�Һ�
              If r_ԤԼ���� = 2 Then
                n_���� := 1;
              Else
                Begin
                  Select ���Ʒ�ʽ
                  Into n_��Լģʽ
                  From �ٴ�����Һſ��Ƽ�¼
                  Where ��¼id = r_����id And ���� = 1 And ���� = v_������λ And ���� = 1 And Rownum < 2;
                Exception
                  When Others Then
                    n_��Լģʽ := 4;
                End;
                If n_��Լģʽ = 0 Then
                  n_���� := 1;
                End If;
                If n_��Լģʽ = 1 Or n_��Լģʽ = 2 Then
                  Select ����
                  Into n_��Լ������
                  From �ٴ�����Һſ��Ƽ�¼
                  Where ��¼id = r_����id And ���� = 1 And ���� = v_������λ And ���� = 1;
                  If n_��Լģʽ = 1 Then
                    n_��Լ������ := Round(r_�޺��� * n_��Լ������ / 100);
                  End If;
                  Begin
                    Select Count(1)
                    Into n_��Լ�ѹ���
                    From ���˹Һż�¼
                    Where �ű� = r_���� And ��¼״̬ = 1 And ������λ = v_������λ And ����ʱ�� Between Trunc(d_����) And
                          Trunc(d_���� + 1) - 1 / 60 / 60 / 24;
                  Exception
                    When Others Then
                      n_��Լ�ѹ��� := 0;
                  End;
                  n_��Լʣ������ := n_��Լ������ - n_��Լ�ѹ���;
                  If Nvl(r_��Լ��, r_�޺���) - Nvl(r_��Լ��, 0) < n_��Լʣ������ Then
                    v_ʣ������ := Nvl(r_��Լ��, r_�޺���) - Nvl(r_��Լ��, 0);
                  Else
                    v_ʣ������ := n_��Լʣ������;
                  End If;
                  If r_��ʱ�� = 1 Then
                    v_Temp := '<SPANLIST>';
                    If r_��ſ��� = 1 Then
                      --��ʱ��,��ſ���
                      n_Exists := 0;
                      For r_Time In (Select ���, ��ʼʱ��, ��ֹʱ��, �Һ�״̬
                                     From �ٴ�������ſ���
                                     Where ��¼id = r_����id) Loop
                        v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.��ֹʱ��, 'hh24:mi:ss') || '</SJD>';
                        If Nvl(r_Time.�Һ�״̬, 0) = 0 Then
                          v_Temp   := v_Temp || '<SL>1</SL></SPAN>';
                          n_Exists := n_Exists + 1;
                        Else
                          v_Temp := v_Temp || '<SL>0</SL></SPAN>';
                        End If;
                      End Loop;
                      If n_Exists < To_Number(v_ʣ������) Then
                        v_ʣ������ := n_Exists;
                      End If;
                    Else
                      --��ʱ��,����ſ���
                      n_Exists := 0;
                      For r_Time In (Select ���, ��ʼʱ��, ��ֹʱ��, �Һ�״̬, ����
                                     From �ٴ�������ſ���
                                     Where ��¼id = r_����id And ԤԼ˳��� Is Null And �Ƿ�ԤԼ = 1) Loop
                        Select Count(1)
                        Into n_ʱ���ѹ�
                        From �ٴ�������ſ���
                        Where ԤԼ˳��� Is Not Null And Nvl(�Һ�״̬, 0) <> 0;
                        v_Temp   := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                    To_Char(r_Time.��ֹʱ��, 'hh24:mi:ss') || '</SJD>';
                        v_Temp   := v_Temp || '<SL>' || r_Time.���� - n_ʱ���ѹ� || '</SL></SPAN>';
                        n_Exists := n_Exists + (r_Time.���� - n_ʱ���ѹ�);
                      End Loop;
                      If n_Exists < To_Number(v_ʣ������) Then
                        v_ʣ������ := n_Exists;
                      End If;
                    End If;
                    v_Temp := v_Temp || '</SPANLIST>';
                  End If;
                  n_�ѹ��� := r_��Լ��;
                End If;
                If n_��Լģʽ = 3 Then
                  If r_��ʱ�� = 0 Then
                    If r_��ſ��� = 0 Then
                      n_���� := 1;
                    Else
                      n_�ѹ���   := 0;
                      v_ʣ������ := 0;
                      For r_���� In (Select ���
                                   From �ٴ�����Һſ��Ƽ�¼
                                   Where ��¼id = r_����id And ���� = 1 And ���� = 1 And ���� = v_������λ) Loop
                        Begin
                          Select 1
                          Into n_Exists
                          From �ٴ�������ſ���
                          Where ��¼id = r_����id And ��� = r_����.��� And �Ƿ�ԤԼ = 1 And Nvl(�Һ�״̬, 0) = 0;
                        Exception
                          When Others Then
                            n_Exists := 0;
                        End;
                        If n_Exists = 1 Then
                          v_ʣ������ := v_ʣ������ + 1;
                        Else
                          n_�ѹ��� := n_�ѹ��� + 1;
                        End If;
                      End Loop;
                      If Nvl(r_��Լ��, r_�޺���) - Nvl(r_��Լ��, 0) < v_ʣ������ Then
                        v_ʣ������ := Nvl(r_��Լ��, r_�޺���) - Nvl(r_��Լ��, 0);
                      End If;
                    End If;
                  Else
                    If r_��ſ��� = 0 Then
                      --������λ,��ʱ��,����ſ���
                      n_�ѹ���   := 0;
                      v_ʣ������ := 0;
                      v_Temp     := '<SPANLIST>';
                      For r_���� In (Select ���, ����
                                   From �ٴ�����Һſ��Ƽ�¼
                                   Where ��¼id = r_����id And ���� = 1 And ���� = 1 And ���� = v_������λ) Loop
                        Select Count(1), Max(��ʼʱ��), Max(��ֹʱ��)
                        Into n_ʱ���ѹ�, d_��ʼʱ��, d_��ֹʱ��
                        From �ٴ�������ſ���
                        Where ��¼id = r_����id And ԤԼ˳��� Is Not Null And Nvl(�Һ�״̬, 0) <> 0 And ��� = r_����.���;
                        n_�ѹ���   := n_�ѹ��� + n_Exists;
                        v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                      To_Char(d_��ֹʱ��, 'hh24:mi:ss') || '</SJD>';
                        v_Temp     := v_Temp || '<SL>' || r_����.���� - n_ʱ���ѹ� || '</SL></SPAN>';
                        v_ʣ������ := v_ʣ������ + r_����.���� - n_ʱ���ѹ�;
                      End Loop;
                      v_Temp := v_Temp || '</SPANLIST>';
                    Else
                      n_�ѹ���   := 0;
                      v_ʣ������ := 0;
                      v_Temp     := '<SPANLIST>';
                      For r_���� In (Select ���
                                   From �ٴ�����Һſ��Ƽ�¼
                                   Where ��¼id = r_����id And ���� = 1 And ���� = 1 And ���� = v_������λ) Loop
                        Begin
                          Select 1, ��ʼʱ��, ��ֹʱ��
                          Into n_Exists, d_��ʼʱ��, d_��ֹʱ��
                          From �ٴ�������ſ���
                          Where ��¼id = r_����id And Nvl(�Һ�״̬, 0) = 0 And ��� = r_����.���;
                        Exception
                          When Others Then
                            n_Exists := 0;
                        End;
                        If n_Exists = 1 Then
                          v_ʣ������ := v_ʣ������ + 1;
                          v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                        To_Char(d_��ֹʱ��, 'hh24:mi:ss') || '</SJD>';
                          v_Temp     := v_Temp || '<SL>' || 1 || '</SL></SPAN>';
                        Else
                          v_Temp   := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                      To_Char(d_��ֹʱ��, 'hh24:mi:ss') || '</SJD>';
                          v_Temp   := v_Temp || '<SL>' || 0 || '</SL></SPAN>';
                          n_�ѹ��� := n_�ѹ��� + 1;
                        End If;
                      End Loop;
                    End If;
                  End If;
                End If;
                If n_��Լģʽ = 4 Then
                  If r_��ʱ�� = 0 Then
                    n_�ѹ���   := r_��Լ��;
                    v_ʣ������ := Nvl(r_��Լ��, r_�޺���) - Nvl(r_��Լ��, 0);
                  Else
                    --��ʱ��
                    n_�ѹ���   := 0;
                    v_ʣ������ := 0;
                    v_Temp     := '<SPANLIST>';
                    If r_��ſ��� = 0 Then
                      --����ſ��Ʒ�ʱ��ԤԼ
                      For r_Time In (Select ���, ��ʼʱ��, ��ֹʱ��, �Һ�״̬, ����
                                     From �ٴ�������ſ���
                                     Where ��¼id = r_����id And ԤԼ˳��� Is Null And �Ƿ�ԤԼ = 1) Loop
                        Select Count(1)
                        Into n_ʱ���ѹ�
                        From �ٴ�������ſ���
                        Where ԤԼ˳��� Is Not Null And Nvl(�Һ�״̬, 0) <> 0;
                        v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                      To_Char(r_Time.��ֹʱ��, 'hh24:mi:ss') || '</SJD>';
                        v_Temp     := v_Temp || '<SL>' || r_Time.���� - n_ʱ���ѹ� || '</SL></SPAN>';
                        n_�ѹ���   := n_�ѹ��� + n_ʱ���ѹ�;
                        v_ʣ������ := v_ʣ������ + (r_Time.���� - n_ʱ���ѹ�);
                      End Loop;
                    Else
                      For r_Time In (Select ���, ��ʼʱ��, ��ֹʱ��, �Һ�״̬
                                     From �ٴ�������ſ���
                                     Where ��¼id = r_����id) Loop
                        v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.��ֹʱ��, 'hh24:mi:ss') || '</SJD>';
                        If Nvl(r_Time.�Һ�״̬, 0) = 0 Then
                          v_Temp     := v_Temp || '<SL>1</SL></SPAN>';
                          v_ʣ������ := v_ʣ������ + 1;
                        Else
                          v_Temp   := v_Temp || '<SL>0</SL></SPAN>';
                          n_�ѹ��� := n_�ѹ��� + 1;
                        End If;
                      End Loop;
                    End If;
                    v_Temp := v_Temp || '</SPANLIST>';
                  End If;
                End If;
              End If;
            End If;
          End If;
        End If;
      
        If Nvl(n_����, 0) = 0 Then
	  n_�ϼƽ�� := 0;
          For r_Fee In (Select b.�ּ�, a.��������
                        From �շѴ�����Ŀ A, �շѼ�Ŀ B
                        Where a.����id = r_��Ŀid And a.����id = b.�շ�ϸĿid And Sysdate Between b.ִ������ And
                              Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))
                        Union
                        Select b.�ּ�, 1 As ��������
                        From �շ���ĿĿ¼ A, �շѼ�Ŀ B
                        Where a.Id = b.�շ�ϸĿid And a.Id = r_��Ŀid And Sysdate Between b.ִ������ And
                              Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))) Loop
            n_�ϼƽ�� := n_�ϼƽ�� + r_Fee.�ּ� * r_Fee.��������;
          End Loop;
          v_ʱ���  := To_Char(r_��ʼʱ��, 'HH24:MI') || '-' || To_Char(r_��ֹʱ��, 'HH24:MI');
          c_Xmlmain := '<HB>' || '<CZJLID>' || r_����id || '</CZJLID>' || '<HM>' || r_���� || '</HM>' || '<YSID>' || r_ҽ��id ||
                       '</YSID>' || '<YS>' || r_ҽ������ || '</YS>' || '<KSID>' || r_����id || '</KSID>' || '<KSMC>' ||
                       r_�������� || '</KSMC>' || '<ZC>' || r_ְ�� || '</ZC>' || '<XMID>' || r_��Ŀid || '</XMID>' || '<XMMC>' ||
                       r_��Ŀ���� || '</XMMC>' || '<YGHS>' || n_�ѹ��� || '</YGHS>' || '<SYHS>' || v_ʣ������ || '</SYHS>' ||
                       '<PRICE>' || n_�ϼƽ�� || '</PRICE>' || '<HL>' || r_���� || '</HL>' || '<FSD>' || r_��ʱ�� || '</FSD>' ||
                       '<HBTIME>' || v_ʱ��� || '</HBTIME>' || '<FWMC>' || r_�Ű� || '</FWMC>' || v_Temp || '</HB>';
          v_Xmlmain := v_Xmlmain || c_Xmlmain;
        End If;
        n_���� := 0;
      End Loop;
      Close r_No;
      v_Xmlmain := '<GROUP>' || '<RQ>' || To_Char(d_����, 'yyyy-mm-dd') || '</RQ>' || '<HBLIST>' || v_Xmlmain ||
                   '</HBLIST>' || '</GROUP>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Xmlmain)) Into x_Templet From Dual;
    End If;
  End If;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getnolist;
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
  --   <JSLIST>
  --     <JS>            //������Ϣ���Һ�Ŀǰ��֧��һ�����ṹ���շ�һ�£��Ժ����չ
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
  --   <HZDW>������λ</HZDW>        //������λ����
  --   <YYFS>֧����<YYFS>    //ԤԼ��ʽ,����������֧����
  --   <BRID>����ID</BRID>     //����ID
  --   <BRLX></BRLX>             //ҽ����������
  --   <FB>��ͨ</FB>               //���˷ѱ𣬿��Բ���
  --   <JQM>������</JQM>            //������
  --</IN>

  --����:Xml_Out
  --<OUTPUT>
  -- <GHDH>�Һŵ���</GHDH>          //�Һŵ���
  -- <CZSJ>����ʱ��</CZSJ>          //HIS�ĵǼ�ʱ��
  -- <ERROR><MSG>������Ϣ</MSG></ERROR>  //����ʱ����
  --</ OUTPUT>
  --------------------------------------------------------------------------------------------------
  v_����     �ҺŰ���.����%Type;
  d_����ʱ�� Date;
  d_ԭʼʱ�� Date;
  d_�Ǽ�ʱ�� Date;
  v_���     Varchar2(200);

  n_Ӧ�ս��   ������ü�¼.Ӧ�ս��%Type;
  v_��ˮ��     ����Ԥ����¼.������ˮ��%Type;
  v_˵��       ������ü�¼.ժҪ%Type;
  n_����id     ������Ϣ.����id%Type;
  v_ԤԼ��ʽ   ԤԼ��ʽ.����%Type;
  v_��������� ҽ�ƿ����.����%Type;
  v_���㿨��   ����Ԥ����¼.����%Type;
  n_�����     ������ü�¼.��ʶ��%Type;
  v_����       ������ü�¼.����%Type;
  v_�Ա�       ������ü�¼.�Ա�%Type;
  v_����       ������ü�¼.����%Type;
  v_���ʽ   ������ü�¼.���ʽ%Type;
  v_�ѱ�       ������ü�¼.�ѱ�%Type;
  v_No         ���˹Һż�¼.No%Type;
  v_���㷽ʽ   ҽ�ƿ����.���㷽ʽ%Type;
  v_�շ����   ������ü�¼.�շ����%Type;
  n_�շ�ϸĿid ������ü�¼.�շ�ϸĿid%Type;
  n_��׼����   ������ü�¼.��׼����%Type;
  n_������Ŀid ������ü�¼.������Ŀid%Type;
  n_���ηѱ�   �շ���ĿĿ¼.���ηѱ�%Type;
  v_�վݷ�Ŀ   ������ü�¼.�վݷ�Ŀ%Type;
  n_���˿���id ������ü�¼.���˿���id%Type;
  n_��������id ������ü�¼.��������id%Type;
  v_����Ա��� ������ü�¼.����Ա���%Type;
  v_����Ա���� ������ü�¼.����Ա����%Type;
  v_ҽ������   �ҺŰ���.ҽ������%Type;
  n_ҽ��id     �ҺŰ���.ҽ��id%Type;
  n_����id     ������ü�¼.����id%Type;
  n_�����id   ҽ�ƿ����.Id%Type;
  v_�Ű�       �ҺŰ���.����%Type;
  n_����id     �ҺŰ���.Id%Type;
  n_�ƻ�id     �ҺŰ��żƻ�.Id%Type;
  n_Ԥ��id     ����Ԥ����¼.Id%Type;
  n_��ſ���   �ҺŰ���.��ſ���%Type;
  n_����       �Һ����״̬.���%Type;
  v_����       �ҺŰ�������.������Ŀ%Type;
  v_��������   ������Ϣ.��������%Type;
  n_����       Number(3);
  v_�ֽ�       ���㷽ʽ.����%Type;
  n_��ʱ��     Number(3);
  v_��������   Varchar2(3000);
  v_������λ   ���˹Һż�¼.������λ%Type;
  v_������     �Һ����״̬.������%Type;
  n_�ɿʽ   Number(3);
  n_��¼id     �ٴ������¼.Id%Type;
  v_Temp       Varchar2(32767); --��ʱXML
  x_Templet    Xmltype; --ģ��XML
  v_Err_Msg    Varchar2(200);
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/HM'), Extractvalue(Value(A), 'IN/HX'),
         To_Date(Extractvalue(Value(A), 'IN/YYSJ'), 'YYYY-MM-DD hh24:mi:ss'), Extractvalue(Value(A), 'IN/JE'),
         Extractvalue(Value(A), 'IN/YYFS'), Extractvalue(Value(A), 'IN/HZDW'), Extractvalue(Value(A), 'IN/BRID'),
         Extractvalue(Value(A), 'IN/BRLX'), Extractvalue(Value(A), 'IN/FB'), Extractvalue(Value(A), 'IN/JQM'),
         Extractvalue(Value(A), 'IN/JKFS'), Extractvalue(Value(A), 'IN/CZJLID')
  Into v_����, n_����, d_ԭʼʱ��, n_Ӧ�ս��, v_ԤԼ��ʽ, v_������λ, n_����id, v_��������, v_�ѱ�, v_������, n_�ɿʽ, n_��¼id
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  d_�Ǽ�ʱ�� := Sysdate;
  d_����ʱ�� := Trunc(d_ԭʼʱ��);
  If v_�������� Is Not Null Then
    Begin
      Select 1 Into n_���� From �������� Where ���� = v_��������;
    Exception
      When Others Then
        v_Err_Msg := 'û�з���Ϊ(' || v_�������� || ')�Ĳ�������';
        Raise Err_Item;
    End;
    Update ������Ϣ Set �������� = Nvl(��������, v_��������) Where ����id = n_����id;
  End If;

  Select a.�����, a.����, a.�Ա�, a.����, Nvl(b.����, c.����)
  Into n_�����, v_����, v_�Ա�, v_����, v_���ʽ
  From ������Ϣ A, ҽ�Ƹ��ʽ B, (Select ���� From ҽ�Ƹ��ʽ Where ȱʡ��־ = '1' And Rownum < 2) C
  Where a.����id = n_����id And a.ҽ�Ƹ��ʽ = b.����(+);
  v_No   := Nextno(12);
  v_Temp := Zl_Identity(1);
  Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_Temp From Dual;
  Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into v_����Ա��� From Dual;
  Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_����Ա���� From Dual;
  v_Temp := Zl_Identity(2);
  Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into n_��������id From Dual;

  If n_��¼id Is Null Then
    Select Extractvalue(b.Column_Value, '/JS/JSKLB'), Extractvalue(b.Column_Value, '/JS/JSKH'),
           Extractvalue(b.Column_Value, '/JS/JSFS'), Extractvalue(b.Column_Value, '/JS/JYLSH'),
           Extractvalue(b.Column_Value, '/JS/JYSM')
    Into v_���������, v_���㿨��, v_���㷽ʽ, v_��ˮ��, v_˵��
    From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B;
  
    Begin
      Select b.���㷽ʽ, b.Id Into v_���㷽ʽ, n_�����id From ҽ�ƿ���� B Where b.���� = v_��������� And Rownum < 2;
    Exception
      When Others Then
        v_Err_Msg := 'û�з��ָý��㿨�������Ϣ';
        Raise Err_Item;
    End;
    Select Decode(To_Char(d_ԭʼʱ��, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����',
                   Null)
    Into v_����
    From Dual;
    Begin
      Select ID
      Into n_�ƻ�id
      From (Select ID
             From �ҺŰ��żƻ�
             Where ���� = v_���� And d_ԭʼʱ�� Between Nvl(��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                   Nvl(ʧЧʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And ���ʱ�� Is Not Null
             Order By ��Чʱ�� Desc)
      Where Rownum < 2;
    Exception
      When Others Then
        Select ID Into n_����id From �ҺŰ��� Where ���� = v_����;
    End;
    If Nvl(n_�ƻ�id, 0) <> 0 Then
      --�Ӽƻ���ȡ��Ϣ
      Select a.��Ŀid, b.����id, a.ҽ������, a.ҽ��id,
             Decode(To_Char(d_����ʱ��, 'D'), '1', a.����, '2', a.��һ, '3', a.�ܶ�, '4', a.����, '5', a.����, '6', a.����, '7', a.����,
                     Null), Nvl(a.��ſ���, 0)
      Into n_�շ�ϸĿid, n_���˿���id, v_ҽ������, n_ҽ��id, v_�Ű�, n_��ſ���
      From �ҺŰ��żƻ� A, �ҺŰ��� B
      Where a.Id = n_�ƻ�id And b.Id = a.����id;
      Select Count(Rownum) Into n_��ʱ�� From �Һżƻ�ʱ�� Where ���� = v_���� And �ƻ�id = n_�ƻ�id And Rownum < 2;
      --������λ���
      If v_������λ Is Not Null Then
        Begin
          Select 1 Into n_���� From ������λ�ƻ����� Where �ƻ�id = n_�ƻ�id And ���� = 0 And ������λ = v_������λ;
        Exception
          When Others Then
            n_���� := 0;
        End;
      End If;
      If n_���� = 1 Then
        v_Err_Msg := '����ĺ�����λ�ڴ˺����ϱ����ã�';
        Raise Err_Item;
      End If;
      If n_��ʱ�� = 1 And n_��ſ��� = 0 Then
        d_����ʱ�� := d_ԭʼʱ��;
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
              Select To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || '' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
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
      Select b.��Ŀid, b.����id, b.ҽ������, b.ҽ��id,
             Decode(To_Char(d_����ʱ��, 'D'), '1', b.����, '2', b.��һ, '3', b.�ܶ�, '4', b.����, '5', b.����, '6', b.����, '7', b.����,
                     Null), Nvl(b.��ſ���, 0)
      Into n_�շ�ϸĿid, n_���˿���id, v_ҽ������, n_ҽ��id, v_�Ű�, n_��ſ���
      From �ҺŰ��� B
      Where b.Id = n_����id;
      Select Count(Rownum) Into n_��ʱ�� From �ҺŰ���ʱ�� Where ���� = v_���� And ����id = n_����id And Rownum < 2;
      --������λ���
      If v_������λ Is Not Null Then
        Begin
          Select 1 Into n_���� From ������λ���ſ��� Where ����id = n_����id And ���� = 0 And ������λ = v_������λ;
        Exception
          When Others Then
            n_���� := 0;
        End;
      End If;
      If n_���� = 1 Then
        v_Err_Msg := '����ĺ�����λ�ڴ˺����ϱ����ã�';
        Raise Err_Item;
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
              Select To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || '' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
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
  
    Select a.���, b.�ּ�, b.������Ŀid, c.�վݷ�Ŀ, a.���ηѱ�
    Into v_�շ����, n_��׼����, n_������Ŀid, v_�վݷ�Ŀ, n_���ηѱ�
    From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C
    Where a.Id = n_�շ�ϸĿid And b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And Sysdate Between b.ִ������ And
          Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And Rownum < 2;
  
    Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
    Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
  
    If Trunc(d_����ʱ��) <> Trunc(Sysdate) Then
      If Nvl(n_�ɿʽ, 0) = 0 Then
        Zl_���������Һ�_Insert(3, n_����id, v_����, n_����, v_No, Null, v_���㷽ʽ, Null, d_����ʱ��, d_�Ǽ�ʱ��, v_������λ, n_Ӧ�ս��, Null, Null,
                         v_��ˮ��, v_˵��, v_ԤԼ��ʽ, n_Ԥ��id, n_�����id, Null, 1, n_����id, 0, Null, Null, v_���㿨��, 1, v_�ѱ�, Null,
                         v_������, 1);
      Else
        Zl_���������Һ�_Insert(2, n_����id, v_����, n_����, v_No, Null, v_���㷽ʽ, Null, d_����ʱ��, d_�Ǽ�ʱ��, v_������λ, n_Ӧ�ս��, Null, Null,
                         v_��ˮ��, v_˵��, v_ԤԼ��ʽ, n_Ԥ��id, n_�����id, Null, 1, n_����id, 0, Null, Null, v_���㿨��, 1, v_�ѱ�, Null,
                         v_������, 1);
      End If;
    Else
      Zl_���������Һ�_Insert(1, n_����id, v_����, n_����, v_No, Null, v_���㷽ʽ, Null, d_����ʱ��, d_�Ǽ�ʱ��, v_������λ, n_Ӧ�ս��, Null, Null,
                       v_��ˮ��, v_˵��, v_ԤԼ��ʽ, n_Ԥ��id, n_�����id, Null, 1, n_����id, 0, Null, Null, v_���㿨��, 1, v_�ѱ�, Null,
                       v_������, 1);
    End If;
  
    For c_��չ��Ϣ In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As ��������,
                          Extractvalue(b.Column_Value, '/EXPEND/JYLR') As ��������
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS/EXPENDLIST/EXPEND'))) B) Loop
      Zl_�������㽻��_Insert(n_�����id, 0, v_���㿨��, n_����id, c_��չ��Ϣ.�������� || '|' || c_��չ��Ϣ.��������, 0);
    End Loop;
  
    v_Temp := '<GHDH>' || v_No || '</GHDH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<CZSJ>' || To_Char(d_�Ǽ�ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Else
    --������Ű�ģʽ
    For r_���� In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As ���㿨���,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As ���㿨��,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As ���㷽ʽ,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As ������ˮ��,
                        Extractvalue(b.Column_Value, '/JS/JYSM') As ����˵��,
                        Extractvalue(b.Column_Value, '/JS/JSJE') As ������
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
      v_�������� := v_�������� || '|' || r_����.���㷽ʽ || ',' || r_����.������ || ',,';
      If r_����.���㿨��� Is Not Null Then
        v_��������   := v_�������� || '1';
        v_��������� := r_����.���㿨���;
        v_���㿨��   := r_����.���㿨��;
        v_��ˮ��     := r_����.������ˮ��;
        v_˵��       := r_����.����˵��;
      Else
        v_�������� := v_�������� || '0';
      End If;
    End Loop;
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
    Select ��Ŀid, ����id, ҽ������, ҽ��id, �Ƿ���ſ���, �Ƿ��ʱ��
    Into n_�շ�ϸĿid, n_���˿���id, v_ҽ������, n_ҽ��id, n_��ſ���, n_��ʱ��
    From �ٴ������¼
    Where ID = n_��¼id;
  
    Begin
      Select ��ʼʱ�� Into d_����ʱ�� From �ٴ�������ſ��� Where ��¼id = n_��¼id And ��� = n_����;
    Exception
      When Others Then
        d_����ʱ�� := d_ԭʼʱ��;
    End;
  
    Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
    Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
  
    If Trunc(d_����ʱ��) <> Trunc(Sysdate) Then
      If Nvl(n_�ɿʽ, 0) = 0 Then
        Zl_���������Һ�_Insert(3, n_����id, v_����, n_����, v_No, Null, v_��������, Null, d_����ʱ��, d_�Ǽ�ʱ��, v_������λ, n_Ӧ�ս��,
                           Null, Null, v_��ˮ��, v_˵��, v_ԤԼ��ʽ, n_Ԥ��id, n_�����id, Null, 1, n_����id, 0, Null, Null, v_���㿨��, 1,
                           v_�ѱ�, Null, v_������, 1, n_��¼id);
      Else
        Zl_���������Һ�_Insert(2, n_����id, v_����, n_����, v_No, Null, v_��������, Null, d_����ʱ��, d_�Ǽ�ʱ��, v_������λ, n_Ӧ�ս��,
                           Null, Null, v_��ˮ��, v_˵��, v_ԤԼ��ʽ, n_Ԥ��id, n_�����id, Null, 1, n_����id, 0, Null, Null, v_���㿨��, 1,
                           v_�ѱ�, Null, v_������, 1, n_��¼id);
      End If;
    Else
      Zl_���������Һ�_Insert(1, n_����id, v_����, n_����, v_No, Null, v_���㷽ʽ, Null, d_����ʱ��, d_�Ǽ�ʱ��, v_������λ, n_Ӧ�ս��, Null,
                         Null, v_��ˮ��, v_˵��, v_ԤԼ��ʽ, n_Ԥ��id, n_�����id, Null, 1, n_����id, 0, Null, Null, v_���㿨��, 1, v_�ѱ�,
                         Null, v_������, 1, n_��¼id);
    End If;
  
    For c_��չ��Ϣ In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As ��������,
                          Extractvalue(b.Column_Value, '/EXPEND/JYLR') As ��������
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS/EXPENDLIST/EXPEND'))) B) Loop
      Zl_�������㽻��_Insert(n_�����id, 0, v_���㿨��, n_����id, c_��չ��Ϣ.�������� || '|' || c_��չ��Ϣ.��������, 0);
    End Loop;
  
    v_Temp := '<GHDH>' || v_No || '</GHDH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<CZSJ>' || To_Char(d_�Ǽ�ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  End If;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Regist;
/

Create Or Replace Procedure Zl_���������Һ�_Delete
(
  ���ݺ�_In     ������ü�¼.No%Type,
  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type,
  ����˵��_In   ����Ԥ����¼.����˵��%Type,
  �˺�ʱ��_In   ������ü�¼.�Ǽ�ʱ��%Type := Null,
  Ԥ��id_In     ����Ԥ����¼.Id%Type := Null
) As
  v_Error Varchar(255);
  Err_Custom Exception;

  --���α������ж��Ƿ񵥶��ղ�����,���ҺŻ��ܱ���
  Cursor c_Registinfo
  (
    v_״̬     ���˹Һż�¼.��¼״̬%Type,
    v_����     ���˹Һż�¼.��¼����%Type,
    v_��Ч���� Number := 0
  ) Is
    Select a.����ʱ��, a.�Ǽ�ʱ��, b.��Ŀid, b.����id, b.ҽ������, b.ҽ��id, b.����
    From ���˹Һż�¼ A, �ҺŰ��� B
    Where a.��¼���� = Decode(v_��Ч����, 0, v_����, a.��¼����) And a.��¼״̬ = v_״̬ And a.No = ���ݺ�_In And a.�ű� = b.���� And Rownum = 1;

  r_Registrow c_Registinfo%RowType;

  --�ù�����ڴ�����Ա�ɿ�������˵Ĳ�ͬ���㷽ʽ�Ľ��
  Cursor c_Opermoney Is
    Select Distinct b.���㷽ʽ, b.��Ԥ��
    From ������ü�¼ A, ����Ԥ����¼ B
    Where a.����id = b.����id And a.No = ���ݺ�_In And a.��¼���� = 4 And a.��¼״̬ = 3 And b.��¼���� = 4 And b.��¼״̬ = 3 And
          Nvl(b.��Ԥ��, 0) <> 0;

  n_ִ��״̬       ���˹Һż�¼.ִ��״̬%Type;
  n_��ӡid         Ʊ�ݴ�ӡ����.Id%Type;
  n_����id         ������ü�¼.����id%Type;
  n_ԭ����id       ����Ԥ����¼.����id%Type;
  n_����id         ������Ϣ.����id%Type;
  n_����ֵ         �������.Ԥ�����%Type;
  n_����̨ǩ���Ŷ� Number;
  n_Ԥ��id         ����Ԥ����¼.Id%Type;
  n_ԤԼ�Һ�       Number;
  n_��Ч����       Number; --��Ч����û�в������õ���
  n_�Һ����ɶ���   Number;
  n_Count          Number;
  n_��id           ����ɿ����.Id%Type;
  d_�˺�ʱ��       Date;
  v_����Ա���     ��Ա��.���%Type;
  v_����Ա����     ��Ա��.����%Type;
  v_������λ       ������λ�ҺŻ���.������λ%Type;
  n_ԤԼ״̬       ���˹Һż�¼.ԤԼ%Type;
  v_Temp           Varchar2(100);
  d_�Ǽ�ʱ��       ���˹Һż�¼.�Ǽ�ʱ��%Type;
  v_�ű�           ���˹Һż�¼.�ű�%Type;
  n_����           ���˹Һż�¼.����%Type;
  n_���÷�ʱ��     Number;
  d_ԤԼʱ��       ���˹Һż�¼.ԤԼʱ��%Type;
  n_������λ����   Number(18);
  n_ԤԼ���ɶ���   Number;
  n_��¼����       Number;
  n_״̬           Number;
  n_�˺�����       Number(3);
  n_�Һ��Ű�ģʽ   Number;
  n_�Һ�id         ���˹Һż�¼.Id%Type;
  Function Zl_����Ա
  (
    Type_In     Integer,
    Splitstr_In Varchar2
  ) Return Varchar2 As
    n_Step Number(18);
    v_Sub  Varchar2(1000);
    --Type_In:0-��ȡȱʡ����ID;1-��ȡ����Ա���;2-��ȡ����Ա����
    -- SplitStr:��ʽΪ:����ID,��������;��ԱID,��Ա���,��Ա����(��Zl_Identity��ȡ��)
  Begin
    If Type_In = 0 Then
      --ȱʡ����
      n_Step := Instr(Splitstr_In, ',');
      v_Sub  := Substr(Splitstr_In, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 1 Then
      --����Ա����
      n_Step := Instr(Splitstr_In, ';');
      v_Sub  := Substr(Splitstr_In, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 2 Then
      --����Ա����
      n_Step := Instr(Splitstr_In, ';');
      v_Sub  := Substr(Splitstr_In, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      Return v_Sub;
    End If;
  End;

  Procedure Zl_���������Һ�_����_Delete
  (
    ���ݺ�_In     ������ü�¼.No%Type,
    ������ˮ��_In ����Ԥ����¼.������ˮ��%Type,
    ����˵��_In   ����Ԥ����¼.����˵��%Type,
    �˺�ʱ��_In   ������ü�¼.�Ǽ�ʱ��%Type := Null,
    Ԥ��id_In     ����Ԥ����¼.Id%Type := Null
  ) As
    v_Error Varchar(255);
    Err_Custom Exception;
  
    --���α������ж��Ƿ񵥶��ղ�����,���ҺŻ��ܱ���
    Cursor c_Registinfo
    (
      v_״̬     ���˹Һż�¼.��¼״̬%Type,
      v_����     ���˹Һż�¼.��¼����%Type,
      v_��Ч���� Number := 0
    ) Is
      Select a.����ʱ��, a.�Ǽ�ʱ��, b.��Ŀid, b.����id, b.ҽ������, b.ҽ��id, b.Id As ��¼id, a.�ű� As ����
      From ���˹Һż�¼ A, �ٴ������¼ B
      Where a.��¼���� = Decode(v_��Ч����, 0, v_����, a.��¼����) And a.��¼״̬ = v_״̬ And a.No = ���ݺ�_In And a.�����¼id = b.Id And
            Rownum < 2;
  
    r_Registrow c_Registinfo%RowType;
  
    --�ù�����ڴ�����Ա�ɿ�������˵Ĳ�ͬ���㷽ʽ�Ľ��
    Cursor c_Opermoney Is
      Select Distinct b.���㷽ʽ, b.��Ԥ��
      From ������ü�¼ A, ����Ԥ����¼ B
      Where a.����id = b.����id And a.No = ���ݺ�_In And a.��¼���� = 4 And a.��¼״̬ = 3 And b.��¼���� = 4 And b.��¼״̬ = 3 And
            Nvl(b.��Ԥ��, 0) <> 0;
  
    n_ִ��״̬       ���˹Һż�¼.ִ��״̬%Type;
    n_��ӡid         Ʊ�ݴ�ӡ����.Id%Type;
    n_����id         ������ü�¼.����id%Type;
    n_ԭ����id       ����Ԥ����¼.����id%Type;
    n_����id         ������Ϣ.����id%Type;
    n_����ֵ         �������.Ԥ�����%Type;
    n_����̨ǩ���Ŷ� Number;
    n_Ԥ��id         ����Ԥ����¼.Id%Type;
    n_ԤԼ�Һ�       Number;
    n_��Ч����       Number; --��Ч����û�в������õ���
    n_�Һ����ɶ���   Number;
    n_Count          Number;
    n_��id           ����ɿ����.Id%Type;
    d_�˺�ʱ��       Date;
    v_����Ա���     ��Ա��.���%Type;
    v_����Ա����     ��Ա��.����%Type;
    v_������λ       ������λ�ҺŻ���.������λ%Type;
    n_ԤԼ״̬       ���˹Һż�¼.ԤԼ%Type;
    v_Temp           Varchar2(100);
    d_�Ǽ�ʱ��       ���˹Һż�¼.�Ǽ�ʱ��%Type;
    v_�ű�           ���˹Һż�¼.�ű�%Type;
    n_����           ���˹Һż�¼.����%Type;
    n_���÷�ʱ��     Number;
    d_ԤԼʱ��       ���˹Һż�¼.ԤԼʱ��%Type;
    n_������λ����   Number(18);
    n_ԤԼ���ɶ���   Number;
    n_��¼����       Number;
    n_״̬           Number;
    n_�˺�����       Number(3);
    n_�Һ�id         ���˹Һż�¼.Id%Type;
    n_��¼id         �ٴ������¼.Id%Type;
    Function Zl_����Ա
    (
      Type_In     Integer,
      Splitstr_In Varchar2
    ) Return Varchar2 As
      n_Step Number(18);
      v_Sub  Varchar2(1000);
      --Type_In:0-��ȡȱʡ����ID;1-��ȡ����Ա���;2-��ȡ����Ա����
      -- SplitStr:��ʽΪ:����ID,��������;��ԱID,��Ա���,��Ա����(��Zl_Identity��ȡ��)
    Begin
      If Type_In = 0 Then
        --ȱʡ����
        n_Step := Instr(Splitstr_In, ',');
        v_Sub  := Substr(Splitstr_In, 1, n_Step - 1);
        Return v_Sub;
      End If;
      If Type_In = 1 Then
        --����Ա����
        n_Step := Instr(Splitstr_In, ';');
        v_Sub  := Substr(Splitstr_In, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, 1, n_Step - 1);
        Return v_Sub;
      End If;
      If Type_In = 2 Then
        --����Ա����
        n_Step := Instr(Splitstr_In, ';');
        v_Sub  := Substr(Splitstr_In, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        Return v_Sub;
      End If;
    End;
  Begin
    --����ID,��������;��ԱID,��Ա���,��Ա����
    v_Temp := Zl_Identity(0);
    If Nvl(v_Temp, ' ') = ' ' Then
      v_Error := '��ǰ������Աδ���ö�Ӧ����Ա��ϵ,���ܼ�����';
      Raise Err_Custom;
    End If;
    v_����Ա��� := Zl_����Ա(1, v_Temp);
    v_����Ա���� := Zl_����Ա(2, v_Temp);
  
    n_��id := Zl_Get��id(v_����Ա����);
  
    d_�˺�ʱ�� := �˺�ʱ��_In;
    If d_�˺�ʱ�� Is Null Then
      d_�˺�ʱ�� := Sysdate;
    End If;
  
    --�����ж�Ҫ�˺�/ȡ��ԤԼ�ļ�¼�Ƿ����
    Begin
      Select Decode(��¼����, 2, 1, 0), ��¼����, �Ǽ�ʱ��, �ű�, ����, Nvl(ԤԼʱ��, ����ʱ��), Nvl(������λ, ''), Nvl(ԤԼ, 0),
             Decode(��¼״̬, 0, 1, 0), �����¼id
      Into n_ԤԼ�Һ�, n_��¼����, d_�Ǽ�ʱ��, v_�ű�, n_����, d_ԤԼʱ��, v_������λ, n_ԤԼ״̬, n_��Ч����, n_��¼id
      From ���˹Һż�¼
      Where NO = ���ݺ�_In And ��¼״̬ In (0, 1) And Rownum < 2;
    Exception
      When Others Then
        n_ԤԼ�Һ� := -1;
    End;
  
    If n_ԤԼ�Һ� = -1 Then
      v_Error := '���ݿ����Ѿ����˺Ż򵥾��������!';
      Raise Err_Custom;
    End If;
  
    Begin
      Select Nvl(�Ƿ��ʱ��, 0) Into n_���÷�ʱ�� From �ٴ������¼ Where ID = n_��¼id And Rownum < 2;
    Exception
      When Others Then
        n_���÷�ʱ�� := 0;
    End;
  
    --ԤԼ����Ƿ���Ӻ�����λ����
    --��������˺�����λ���� ��
    Select Count(0) Into n_������λ���� From �ٴ�����Һſ��Ƽ�¼ Where ���� = 1 And ���� = 1 And Rownum < 2;
    --���¹Һ����״̬
    n_�˺����� := Zl_To_Number(zl_GetSysParameter('�����������Һ�', 1111));
    If n_�˺����� = 0 Then
      Update �ٴ�������ſ��� Set �Һ�״̬ = 4 Where ��¼id = n_��¼id And (��� = n_���� Or ��ע = n_����);
    Else
      Update �ٴ�������ſ���
      Set �Һ�״̬ = 0, ���� = Null, ���� = Null, ����Ա���� = Null, ����վ���� = Null
      Where ��¼id = n_��¼id And (��� = n_���� Or ��ע = n_����);
    End If;
    If Nvl(n_ԤԼ�Һ�, 0) = 1 Or Nvl(n_��Ч����, 0) = 1 Then
      If Nvl(n_��Ч����, 0) = 0 Then
        --N���ڲ���ȡ��ԤԼ��
        n_Count := Zl_To_Number(zl_GetSysParameter('N���ڲ���ȡ��ԤԼ��', 1111));
        If n_Count <> 0 Then
          If Trunc(Sysdate - n_Count) < d_�Ǽ�ʱ�� Then
            v_Error := '�����˵�ԤԼ��' || To_Char(Trunc(Sysdate - n_Count), 'yyyy-mm-dd') || '��ǰ��ԤԼ��!';
            Raise Err_Custom;
          End If;
        End If;
      End If;
    
      n_״̬ := Case n_��Ч����
                When 1 Then
                 0
                Else
                 1
              End;
      --������Լ��
      Open c_Registinfo(n_״̬, 2, n_��Ч����);
      Fetch c_Registinfo
        Into r_Registrow;
    
      Update ���˹ҺŻ���
      Set ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ״̬, �ѹ��� = Nvl(�ѹ���, 0) - Decode(n_ԤԼ״̬, 0, 1, 0)
      Where ���� = Trunc(r_Registrow.����ʱ��) And ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And
            Nvl(ҽ������, 'ҽ��') = Nvl(r_Registrow.ҽ������, 'ҽ��') And Nvl(ҽ��id, 0) = Nvl(r_Registrow.ҽ��id, 0) And
            (���� = r_Registrow.���� Or ���� Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into ���˹ҺŻ���
          (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, ��Լ��, �ѹ���)
        Values
          (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
           Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -n_ԤԼ״̬, Decode(n_ԤԼ״̬, 0, 1, 0));
      End If;
    
      Update �ٴ������¼
      Set ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ״̬, �ѹ��� = Nvl(�ѹ���, 0) - Decode(n_ԤԼ״̬, 0, 1, 0)
      Where ID = n_��¼id;
      Close c_Registinfo;
    
      If Nvl(n_��Ч����, 0) = 0 Then
        --ɾ��������ü�¼
        Delete From ������ü�¼ Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 0;
        --���ԤԼ���ɶ���ʱ��Ҫ�������
        n_�Һ����ɶ��� := Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113));
        If Nvl(n_�Һ����ɶ���, 0) = 1 Then
          n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
          If Nvl(n_ԤԼ���ɶ���, 0) = 1 Then
            --Ҫɾ������
            For v_�Һ� In (Select ID, ����, ִ�в���id, ִ���� From ���˹Һż�¼ Where NO = ���ݺ�_In) Loop
              Zl_�ŶӽкŶ���_Delete(v_�Һ�.ִ�в���id, v_�Һ�.Id);
            End Loop;
          End If;
        End If;
      End If;
    Else
      Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
    
      --���¹Һ����״̬
    
      --���˾���״̬
      Select ����id
      Into n_����id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And ��� = 1;
    
      If n_����id Is Not Null Then
        Update ������Ϣ Set ����״̬ = 0, �������� = Null Where ����id = n_����id;
        --ɾ���������ش���,ֻ�е�ֻ��һ���Һż�¼���Ҳ��˽���������Һ����ڽ���ʱ�Żᴦ��
      End If;
    
      --������ü�¼
      Insert Into ������ü�¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ,
         ����, ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��,
         �Ǽ�ʱ��, ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
        Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
               �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
               v_����Ա���, v_����Ա����, ����ʱ��, d_�˺�ʱ��, n_����id, -1 * ���ʽ��, ������Ŀ��, ���մ���id, -1 * ͳ����, ժҪ,
               Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
        From ������ü�¼
        Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In;
    
      Update ������ü�¼ Set ��¼״̬ = 3 Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In;
      Select ����id
      Into n_ԭ����id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 3 And NO = ���ݺ�_In And Rownum = 1;
    
      n_Ԥ��id := Ԥ��id_In;
      If Nvl(Ԥ��id_In, 0) = 0 Then
        Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      End If;
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ������ˮ��, ����˵��, ������λ,
         �������, �����id, ��������)
        Select n_Ԥ��id, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_�˺�ʱ��, v_����Ա���, v_����Ա����, -��Ԥ��, n_����id, n_��id,
               ������ˮ��_In, ����˵��_In, ������λ, n_����id, �����id, ��������
        From ����Ԥ����¼
        Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_ԭ����id;
    
      Update ����Ԥ����¼ Set ��¼״̬ = 3 Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_ԭ����id;
    
      --�˿��ջ�Ʊ��(�����ϴιҺ�ʹ��Ʊ��,�����ջ�)
      Begin
        --�����һ�εĴ�ӡ������ȡ
        Select ID
        Into n_��ӡid
        From (Select b.Id
               From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
               Where a.��ӡid = b.Id And a.���� = 1 And b.�������� = 4 And b.No = ���ݺ�_In
               Order By a.ʹ��ʱ�� Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          Null;
      End;
      If n_��ӡid Is Not Null Then
        Insert Into Ʊ��ʹ����ϸ
          (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
          Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_�˺�ʱ��, v_����Ա����
          From Ʊ��ʹ����ϸ
          Where ��ӡid = n_��ӡid And ���� = 1;
      End If;
    
      --��ػ��ܱ�Ĵ���
    
      --���˹ҺŻ���
      Open c_Registinfo(1, 1);
      Fetch c_Registinfo
        Into r_Registrow;
    
      If c_Registinfo%RowCount = 0 Then
        --ֻ�ղ�����ʱ�޺ű�,������
        Close c_Registinfo;
      Else
      
        --��Ҫȷ���Ƿ�ԤԼ�Һ�
        --1.�������ԤԼ�ҺŲ����ĹҺż�¼,����Ҫ���ѹ����������ѽ���
        --2.����������Һ�,��ֻ���ѹ���
        Begin
          Select Decode(ԤԼ, Null, 0, 0, 0, 1), ִ��״̬
          Into n_ԤԼ�Һ�, n_ִ��״̬
          From ���˹Һż�¼
          Where NO = ���ݺ�_In And ��¼״̬ = 1 And Rownum = 1;
        Exception
          When Others Then
            n_ԤԼ�Һ� := 0;
        End;
        --0-�ȴ�����,1-��ɾ���,2-���ھ���,-1���Ϊ������
        If n_ִ��״̬ > 0 Then
          If n_ִ��״̬ = 1 Then
            v_Error := '�ò����Ѿ���ɾ���,�������˺�!';
          Else
            v_Error := '�ò������ھ���, �����˺�!';
          End If;
          Raise Err_Custom;
        End If;
      
        Update ���˹ҺŻ���
        Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - n_ԤԼ״̬, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ״̬
        Where ���� = Trunc(r_Registrow.����ʱ��) And ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And
              Nvl(ҽ������, 'ҽ��') = Nvl(r_Registrow.ҽ������, 'ҽ��') And Nvl(ҽ��id, 0) = Nvl(r_Registrow.ҽ��id, 0) And
              (���� = r_Registrow.���� Or ���� Is Null);
      
        If Sql%RowCount = 0 Then
          Insert Into ���˹ҺŻ���
            (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, �����ѽ���)
          Values
            (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
             Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -1, -1 * n_ԤԼ�Һ�);
        End If;
      
        Update �ٴ������¼
        Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - n_ԤԼ״̬, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ״̬
        Where ID = n_��¼id;
        Close c_Registinfo;
      End If;
    
      --��Ա�ɿ����(���������ʻ��ȵĽ�����,�����˳�Ԥ����)
      For r_Opermoney In c_Opermoney Loop
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + (-1 * r_Opermoney.��Ԥ��)
        Where �տ�Ա = v_����Ա���� And ���㷽ʽ = r_Opermoney.���㷽ʽ And ���� = 1
        Returning ��� Into n_����ֵ;
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ����
            (�տ�Ա, ���㷽ʽ, ����, ���)
          Values
            (v_����Ա����, r_Opermoney.���㷽ʽ, 1, -1 * r_Opermoney.��Ԥ��);
          n_����ֵ := r_Opermoney.��Ԥ��;
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From ��Ա�ɿ����
          Where �տ�Ա = v_����Ա���� And ���㷽ʽ = r_Opermoney.���㷽ʽ And ���� = 1 And Nvl(���, 0) = 0;
        End If;
      End Loop;
    
      n_�Һ����ɶ��� := Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113));
      If n_�Һ����ɶ��� <> 0 Then
        n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113));
        If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Then
          --Ҫɾ������
          For v_�Һ� In (Select ID, ����, ִ�в���id, ִ���� From ���˹Һż�¼ Where NO = ���ݺ�_In) Loop
            Zl_�ŶӽкŶ���_Delete(v_�Һ�.ִ�в���id, v_�Һ�.Id);
          End Loop;
        End If;
      End If;
    
      --ҽ�������ľ���ǼǼ�¼
      Delete From ����ǼǼ�¼
      Where (����id, ��ҳid, ����ʱ��) In (Select ����id, ��ҳid, ����ʱ�� From ���˹Һż�¼ Where NO = ���ݺ�_In);
    End If;
  
    If Nvl(n_��Ч����, 0) = 0 Then
      Update ���˹Һż�¼ Set ��¼״̬ = 3 Where NO = ���ݺ�_In And ��¼״̬ = 1;
      If Sql%NotFound Then
        v_Error := 'δ�ҵ��Һŵ���,����!';
        Raise Err_Custom;
      End If;
      Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
      Insert Into ���˹Һż�¼
        (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
         ����, ����, ԤԼ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, �����¼id)
        Select n_�Һ�id, NO, ��¼����, 2, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, d_�˺�ʱ��, ����ʱ��,
               v_����Ա���, v_����Ա����, ����, ����, ԤԼ, ������, ����ʱ��, ������ˮ��_In, ����˵��_In, ������λ, �����¼id
        From ���˹Һż�¼
        Where NO = ���ݺ�_In And ��¼״̬ = 3;
    End If;
    --��Ϣ����
    Begin
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
        Using 2, ���ݺ�_In;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_003(n_�Һ�id, ���ݺ�_In);
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Error || '+[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End;

Begin
  n_�Һ��Ű�ģʽ := Nvl(zl_GetSysParameter('�Һ��Ű�ģʽ'), 0);
  If n_�Һ��Ű�ģʽ = 1 Then
    --������Ű�ģʽ
    Zl_���������Һ�_����_Delete(���ݺ�_In, ������ˮ��_In, ����˵��_In, �˺�ʱ��_In, Ԥ��id_In);
  Else
    --����ID,��������;��ԱID,��Ա���,��Ա����
    v_Temp := Zl_Identity(0);
    If Nvl(v_Temp, ' ') = ' ' Then
      v_Error := '��ǰ������Աδ���ö�Ӧ����Ա��ϵ,���ܼ�����';
      Raise Err_Custom;
    End If;
    v_����Ա��� := Zl_����Ա(1, v_Temp);
    v_����Ա���� := Zl_����Ա(2, v_Temp);
  
    n_��id := Zl_Get��id(v_����Ա����);
  
    d_�˺�ʱ�� := �˺�ʱ��_In;
    If d_�˺�ʱ�� Is Null Then
      d_�˺�ʱ�� := Sysdate;
    End If;
  
    --�����ж�Ҫ�˺�/ȡ��ԤԼ�ļ�¼�Ƿ����
    Begin
      Select Decode(��¼����, 2, 1, 0), ��¼����, �Ǽ�ʱ��, �ű�, ����, Nvl(ԤԼʱ��, ����ʱ��), Nvl(������λ, ''), Nvl(ԤԼ, 0),
             Decode(��¼״̬, 0, 1, 0)
      Into n_ԤԼ�Һ�, n_��¼����, d_�Ǽ�ʱ��, v_�ű�, n_����, d_ԤԼʱ��, v_������λ, n_ԤԼ״̬, n_��Ч����
      From ���˹Һż�¼
      Where NO = ���ݺ�_In And ��¼״̬ In (0, 1) And Rownum <= 1;
    Exception
      When Others Then
        n_ԤԼ�Һ� := -1;
    End;
  
    If n_ԤԼ�Һ� = -1 Then
      v_Error := '���ݿ����Ѿ����˺Ż򵥾��������!';
      Raise Err_Custom;
    End If;
  
    Begin
      Select 1
      Into n_���÷�ʱ��
      From �ҺŰ��� A, �ҺŰ���ʱ�� B
      Where a.���� = v_�ű� And a.Id = b.����id And Rownum <= 1;
    Exception
      When Others Then
        n_���÷�ʱ�� := 0;
    End;
  
    --ԤԼ����Ƿ���Ӻ�����λ����
    --��������˺�����λ���� ��
    Select Count(0) Into n_������λ���� From ������λ���ſ��� Where Rownum = 1;
    --���¹Һ����״̬
    n_�˺����� := Zl_To_Number(zl_GetSysParameter('�����������Һ�', 1111));
    If n_�˺����� = 0 Then
      Update �Һ����״̬
      Set ״̬ = 4
      Where ���� = v_�ű� And ��� = n_���� And ���� Between Trunc(d_ԤԼʱ��) And Trunc(d_ԤԼʱ�� + 1) - 1 / 24 / 60 / 60;
    Else
      Delete �Һ����״̬
      Where ���� = v_�ű� And ��� = n_���� And ���� Between Trunc(d_ԤԼʱ��) And Trunc(d_ԤԼʱ�� + 1) - 1 / 24 / 60 / 60;
    End If;
    If Nvl(n_ԤԼ�Һ�, 0) = 1 Or Nvl(n_��Ч����, 0) = 1 Then
      If Nvl(n_��Ч����, 0) = 0 Then
        --N���ڲ���ȡ��ԤԼ��
        n_Count := Zl_To_Number(zl_GetSysParameter('N���ڲ���ȡ��ԤԼ��', 1111));
        If n_Count <> 0 Then
          If Trunc(Sysdate - n_Count) < d_�Ǽ�ʱ�� Then
            v_Error := '�����˵�ԤԼ��' || To_Char(Trunc(Sysdate - n_Count), 'yyyy-mm-dd') || '��ǰ��ԤԼ��!';
            Raise Err_Custom;
          End If;
        End If;
      End If;
    
      n_״̬ := Case n_��Ч����
                When 1 Then
                 0
                Else
                 1
              End;
      --������Լ��
      Open c_Registinfo(n_״̬, 2, n_��Ч����);
      Fetch c_Registinfo
        Into r_Registrow;
    
      Update ���˹ҺŻ���
      Set ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ״̬, �ѹ��� = Nvl(�ѹ���, 0) - Decode(n_ԤԼ״̬, 0, 1, 0)
      Where ���� = Trunc(r_Registrow.����ʱ��) And ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And
            Nvl(ҽ������, 'ҽ��') = Nvl(r_Registrow.ҽ������, 'ҽ��') And Nvl(ҽ��id, 0) = Nvl(r_Registrow.ҽ��id, 0) And
            (���� = r_Registrow.���� Or ���� Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into ���˹ҺŻ���
          (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, ��Լ��, �ѹ���)
        Values
          (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
           Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -n_ԤԼ״̬, Decode(n_ԤԼ״̬, 0, 1, 0));
      End If;
    
      If Nvl(n_������λ����, 0) <> 0 And Nvl(v_������λ, '') <> '' And Nvl(n_ԤԼ״̬, 0) <> 0 Then
        Update ������λ�ҺŻ���
        Set ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ״̬, �ѽ��� = Nvl(�ѽ���, 0) - Decode(n_ԤԼ״̬, 0, 1, 0)
        Where ���� = Trunc(r_Registrow.����ʱ��) And (���� = r_Registrow.���� Or ���� Is Null) And ������λ = Nvl(v_������λ, '') And
              ��� = Nvl(n_����, 0);
        If Sql%RowCount = 0 Then
          Insert Into ������λ�ҺŻ���
            (����, ����, ��Լ��, ������λ, ���, �ѽ���)
          Values
            (Trunc(r_Registrow.����ʱ��), r_Registrow.����, -n_ԤԼ״̬, v_������λ, Nvl(n_����, 0), -decode(n_ԤԼ״̬, 0, 1, 0));
        End If;
      End If;
      Close c_Registinfo;
    
      If Nvl(n_��Ч����, 0) = 0 Then
        --ɾ��������ü�¼
        Delete From ������ü�¼ Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 0;
        --���ԤԼ���ɶ���ʱ��Ҫ�������
        n_�Һ����ɶ��� := Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113));
        If Nvl(n_�Һ����ɶ���, 0) = 1 Then
          n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
          If Nvl(n_ԤԼ���ɶ���, 0) = 1 Then
            --Ҫɾ������
            For v_�Һ� In (Select ID, ����, ִ�в���id, ִ���� From ���˹Һż�¼ Where NO = ���ݺ�_In) Loop
              Zl_�ŶӽкŶ���_Delete(v_�Һ�.ִ�в���id, v_�Һ�.Id);
            End Loop;
          End If;
        End If;
      End If;
    Else
      Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
    
      --���¹Һ����״̬
    
      --���˾���״̬
      Select ����id
      Into n_����id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And ��� = 1;
    
      If n_����id Is Not Null Then
        Update ������Ϣ Set ����״̬ = 0, �������� = Null Where ����id = n_����id;
        --ɾ���������ش���,ֻ�е�ֻ��һ���Һż�¼���Ҳ��˽���������Һ����ڽ���ʱ�Żᴦ��
      End If;
    
      --������ü�¼
      Insert Into ������ü�¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ,
         ����, ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��,
         �Ǽ�ʱ��, ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
        Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
               �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
               v_����Ա���, v_����Ա����, ����ʱ��, d_�˺�ʱ��, n_����id, -1 * ���ʽ��, ������Ŀ��, ���մ���id, -1 * ͳ����, ժҪ,
               Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
        From ������ü�¼
        Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In;
    
      Update ������ü�¼ Set ��¼״̬ = 3 Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In;
      Select ����id
      Into n_ԭ����id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 3 And NO = ���ݺ�_In And Rownum = 1;
    
      Select Count(Distinct ���㷽ʽ) Into n_Count From ����Ԥ����¼ Where ����id = n_ԭ����id;
      If n_Count > 1 Then
        v_Error := '���ܴ�����ֽ��㷽ʽ,���鴫����˺ŵ����Ƿ���ȷ!';
        Raise Err_Custom;
      End If;
      n_Ԥ��id := Ԥ��id_In;
      If Nvl(Ԥ��id_In, 0) = 0 Then
        Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      End If;
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ������ˮ��, ����˵��, ������λ,
         �������, �����id, ��������)
        Select n_Ԥ��id, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_�˺�ʱ��, v_����Ա���, v_����Ա����, -��Ԥ��, n_����id, n_��id,
               ������ˮ��_In, ����˵��_In, ������λ, n_����id, �����id, ��������
        From ����Ԥ����¼
        Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_ԭ����id;
    
      Update ����Ԥ����¼ Set ��¼״̬ = 3 Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_ԭ����id;
    
      --�˿��ջ�Ʊ��(�����ϴιҺ�ʹ��Ʊ��,�����ջ�)
      Begin
        --�����һ�εĴ�ӡ������ȡ
        Select ID
        Into n_��ӡid
        From (Select b.Id
               From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
               Where a.��ӡid = b.Id And a.���� = 1 And b.�������� = 4 And b.No = ���ݺ�_In
               Order By a.ʹ��ʱ�� Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          Null;
      End;
      If n_��ӡid Is Not Null Then
        Insert Into Ʊ��ʹ����ϸ
          (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
          Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_�˺�ʱ��, v_����Ա����
          From Ʊ��ʹ����ϸ
          Where ��ӡid = n_��ӡid And ���� = 1;
      End If;
    
      --��ػ��ܱ�Ĵ���
    
      --���˹ҺŻ���
      Open c_Registinfo(1, 1);
      Fetch c_Registinfo
        Into r_Registrow;
    
      If c_Registinfo%RowCount = 0 Then
        --ֻ�ղ�����ʱ�޺ű�,������
        Close c_Registinfo;
      Else
      
        --��Ҫȷ���Ƿ�ԤԼ�Һ�
        --1.�������ԤԼ�ҺŲ����ĹҺż�¼,����Ҫ���ѹ����������ѽ���
        --2.����������Һ�,��ֻ���ѹ���
        Begin
          Select Decode(ԤԼ, Null, 0, 0, 0, 1), ִ��״̬
          Into n_ԤԼ�Һ�, n_ִ��״̬
          From ���˹Һż�¼
          Where NO = ���ݺ�_In And ��¼״̬ = 1 And Rownum = 1;
        Exception
          When Others Then
            n_ԤԼ�Һ� := 0;
        End;
        --0-�ȴ�����,1-��ɾ���,2-���ھ���,-1���Ϊ������
        If n_ִ��״̬ > 0 Then
          If n_ִ��״̬ = 1 Then
            v_Error := '�ò����Ѿ���ɾ���,�������˺�!';
          Else
            v_Error := '�ò������ھ���, �����˺�!';
          End If;
          Raise Err_Custom;
        End If;
      
        Update ���˹ҺŻ���
        Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - n_ԤԼ״̬, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ״̬
        Where ���� = Trunc(r_Registrow.����ʱ��) And ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And
              Nvl(ҽ������, 'ҽ��') = Nvl(r_Registrow.ҽ������, 'ҽ��') And Nvl(ҽ��id, 0) = Nvl(r_Registrow.ҽ��id, 0) And
              (���� = r_Registrow.���� Or ���� Is Null);
      
        If Sql%RowCount = 0 Then
          Insert Into ���˹ҺŻ���
            (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, �����ѽ���)
          Values
            (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
             Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -1, -1 * n_ԤԼ�Һ�);
        End If;
        If Nvl(n_������λ����, 0) <> 0 And Nvl(v_������λ, '') <> '' And Nvl(n_ԤԼ״̬, 0) <> 0 Then
          Update ������λ�ҺŻ���
          Set �ѽ��� = Nvl(�ѽ���, 0) - 1, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ�Һ�
          Where ���� = Trunc(r_Registrow.����ʱ��) And (���� = r_Registrow.���� Or ���� Is Null) And ������λ = Nvl(v_������λ, '') And
                ��� = Nvl(n_����, 0);
          If Sql%RowCount = 0 Then
            Insert Into ������λ�ҺŻ���
              (����, ����, ��Լ��, ������λ, �ѽ���)
            Values
              (Trunc(r_Registrow.����ʱ��), r_Registrow.����, -1, v_������λ, -1 * n_ԤԼ�Һ�);
          End If;
        End If;
        Close c_Registinfo;
      End If;
    
      --��Ա�ɿ����(���������ʻ��ȵĽ�����,�����˳�Ԥ����)
      For r_Opermoney In c_Opermoney Loop
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + (-1 * r_Opermoney.��Ԥ��)
        Where �տ�Ա = v_����Ա���� And ���㷽ʽ = r_Opermoney.���㷽ʽ And ���� = 1
        Returning ��� Into n_����ֵ;
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ����
            (�տ�Ա, ���㷽ʽ, ����, ���)
          Values
            (v_����Ա����, r_Opermoney.���㷽ʽ, 1, -1 * r_Opermoney.��Ԥ��);
          n_����ֵ := r_Opermoney.��Ԥ��;
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From ��Ա�ɿ����
          Where �տ�Ա = v_����Ա���� And ���㷽ʽ = r_Opermoney.���㷽ʽ And ���� = 1 And Nvl(���, 0) = 0;
        End If;
      End Loop;
    
      n_�Һ����ɶ��� := Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113));
      If n_�Һ����ɶ��� <> 0 Then
        n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113));
        If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Then
          --Ҫɾ������
          For v_�Һ� In (Select ID, ����, ִ�в���id, ִ���� From ���˹Һż�¼ Where NO = ���ݺ�_In) Loop
            Zl_�ŶӽкŶ���_Delete(v_�Һ�.ִ�в���id, v_�Һ�.Id);
          End Loop;
        End If;
      End If;
    
      --ҽ�������ľ���ǼǼ�¼
      Delete From ����ǼǼ�¼
      Where (����id, ��ҳid, ����ʱ��) In (Select ����id, ��ҳid, ����ʱ�� From ���˹Һż�¼ Where NO = ���ݺ�_In);
    End If;
  
    If Nvl(n_��Ч����, 0) = 0 Then
      Update ���˹Һż�¼ Set ��¼״̬ = 3 Where NO = ���ݺ�_In And ��¼״̬ = 1;
      If Sql%NotFound Then
        v_Error := 'δ�ҵ��Һŵ���,����!';
        Raise Err_Custom;
      End If;
      Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
      Insert Into ���˹Һż�¼
        (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
         ����, ����, ԤԼ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ)
        Select n_�Һ�id, NO, ��¼����, 2, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, d_�˺�ʱ��, ����ʱ��,
               v_����Ա���, v_����Ա����, ����, ����, ԤԼ, ������, ����ʱ��, ������ˮ��_In, ����˵��_In, ������λ
        From ���˹Һż�¼
        Where NO = ���ݺ�_In And ��¼״̬ = 3;
    End If;
    --��Ϣ����
    Begin
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
        Using 2, ���ݺ�_In;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_003(n_�Һ�id, ���ݺ�_In);
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Error || '+[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���������Һ�_Delete;
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
  --  <LSH>34563</LSH>           //������ˮ��
  --  <JKFS>0</JKFS>             //�ɿʽ,0-�ҺŻ�ԤԼ�ɿ�;1-ԤԼ���ɿ�
  --  <YYFS></YYFS>              //�ɿʽ=1ʱ���룬ԤԼ��ԤԼ��ʽ
  --</IN>

  --����:Xml_Out
  --<OUTPUT>
  -- <CZSJ>����ʱ��</CZSJ>          //HIS�ĵǼ�ʱ��
  -- <ERROR><MSG></MSG></ERROR> //Ϊ�ձ�ʾȡ���Һųɹ�
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  v_�����     Varchar2(100);
  v_No         ���˹Һż�¼.No%Type;
  n_�ҺŽ��   ������ü�¼.ʵ�ս��%Type;
  v_����Ա��� ������ü�¼.����Ա���%Type;
  v_����Ա���� ������ü�¼.����Ա����%Type;
  v_���㷽ʽ   ҽ�ƿ����.���㷽ʽ%Type;
  n_ʵ�ս��   ������ü�¼.ʵ�ս��%Type;
  v_������ˮ�� ����Ԥ����¼.������ˮ��%Type;
  n_����       Number(3);
  v_Type       Varchar2(50);
  v_Temp       Varchar2(32767); --��ʱXML
  x_Templet    Xmltype; --ģ��XML
  v_Err_Msg    Varchar2(200);
  n_�ѿ�ҽ��   Number(2);
  n_��鷢Ʊ   Number(3);
  n_�Ƿ��ӡ   Number(3);
  n_�ɿʽ   Number(3);
  d_�Ǽ�ʱ��   Date;
  n_�Һ�ģʽ   Number(3);
  v_ԤԼ��ʽ   ���˹Һż�¼.ԤԼ��ʽ%Type;
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/GHDH'), Extractvalue(Value(A), 'IN/JSKLB'), Extractvalue(Value(A), 'IN/GHJE'),
         Extractvalue(Value(A), 'IN/LSH'), To_Number(Extractvalue(Value(A), 'IN/JCFP')),
         To_Number(Extractvalue(Value(A), 'IN/JKFS')), Extractvalue(Value(A), 'IN/YYFS')
  Into v_No, v_�����, n_�ҺŽ��, v_������ˮ��, n_��鷢Ʊ, n_�ɿʽ, v_ԤԼ��ʽ
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

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
  
    Select Sum(ʵ�ս��) Into n_ʵ�ս�� From ������ü�¼ Where NO = v_No And ��¼���� = 4;
  
    If Nvl(n_�ɿʽ, 0) = 0 Then
      --Ҫ�˵ĵ��ݲ����Ըý��㿨����ģ����ֹ�˺�
      Begin
        Select 1
        Into n_����
        From ����Ԥ����¼ A,
             (Select Distinct ����id
               From ������ü�¼
               Where NO = v_No And ��¼���� = 4
               Union
               Select Distinct ����id From סԺ���ü�¼ Where NO = v_No And ��¼���� = 5) B
        Where a.����id = b.����id And ���㷽ʽ = v_���㷽ʽ And Rownum < 2;
      Exception
        When Others Then
          n_���� := 0;
      End;
      If n_���� = 0 Then
        v_Err_Msg := '����ĹҺŵ��ݲ���' || v_���㷽ʽ || '�����,�޷��˺�!';
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

  --��������飬�Ѵ��ڲ��������ݵģ������˺�
  Begin
    Select 1
    Into n_����
    From ���ò����¼ A,
         (Select Distinct ����id
           From ������ü�¼
           Where NO = v_No And ��¼���� = 4
           Union
           Select Distinct ����id From סԺ���ü�¼ Where NO = v_No And ��¼���� = 5) B
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
  End If;
  --��ȡ����Ա��Ϣ
  v_Temp := Zl_Identity(1);
  Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_Temp From Dual;
  Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into v_����Ա��� From Dual;
  Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_����Ա���� From Dual;
  d_�Ǽ�ʱ�� := Sysdate;
  n_�Һ�ģʽ := zl_GetSysParameter('�Һ��Ű�ģʽ');

  Zl_���������Һ�_Delete(v_No, v_������ˮ��, '�ƶ�ƽ̨�˺�', d_�Ǽ�ʱ��);

  v_Temp := '<CZSJ>' || To_Char(d_�Ǽ�ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
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

Create Or Replace Procedure Zl_�Һ����״̬_Delete
(
  ������ʽ_In Number := 0,
  �ű�_In     ���˹Һż�¼.�ű�%Type := Null
) As
  n_ԤԼ��Чʱ�� Number(5);
  n_ʧԼ���ڹҺ� Number(2);
  n_�Һ���Ч���� Number(5);
Begin
  If ������ʽ_In = 0 Then
    --�����ʷ��¼
    Delete �Һ����״̬ Where ���� < Trunc(Sysdate);
  Else
    --���ʧԼ��
    n_ԤԼ��Чʱ�� := Nvl(zl_GetSysParameter('ԤԼ��Чʱ��', 1111), 0);
    n_ʧԼ���ڹҺ� := Nvl(zl_GetSysParameter('ʧԼ���ڹҺ�', 1111), 0);
    n_�Һ���Ч���� := Nvl(zl_GetSysParameter('�Һ���Ч����'), 7);
    If n_ԤԼ��Чʱ�� <> 0 And n_ʧԼ���ڹҺ� <> 0 Then
      If �ű�_In Is Null Then
        For c_ʧЧԤԼ In (Select b.����, b.����, b.���
                       From ���˹Һż�¼ A, �Һ����״̬ B
                       Where a.ԤԼʱ�� - 1 / 24 / 60 * n_ԤԼ��Чʱ�� < Sysdate And a.ԤԼʱ�� > Sysdate - n_�Һ���Ч���� And a.��¼���� = 2 And
                             a.�ű� = b.���� And a.���� = b.���) Loop
          Delete From �Һ����״̬
          Where ���� = c_ʧЧԤԼ.���� And ��� = c_ʧЧԤԼ.��� And ״̬ = 2 And ���� = c_ʧЧԤԼ.����;
        End Loop;
      Else
        For c_ʧЧԤԼ In (Select b.����, b.����, b.���
                       From ���˹Һż�¼ A, �Һ����״̬ B
                       Where a.ԤԼʱ�� - 1 / 24 / 60 * n_ԤԼ��Чʱ�� < Sysdate And a.ԤԼʱ�� > Sysdate - n_�Һ���Ч���� And a.��¼���� = 2 And
                             a.�ű� = b.���� And a.���� = b.��� And a.�ű� = �ű�_In) Loop
          Delete From �Һ����״̬
          Where ���� = c_ʧЧԤԼ.���� And ��� = c_ʧЧԤԼ.��� And ״̬ = 2 And ���� = c_ʧЧԤԼ.����;
        End Loop;
      End If;
    End If;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�Һ����״̬_Delete;
/

Create Or Replace Procedure Zl_�Һ����״̬_����_Delete(��¼id_In �ٴ������¼.Id%Type := Null) As
  n_ԤԼ��Чʱ�� Number(5);
  n_ʧԼ���ڹҺ� Number(2);
  n_�Һ���Ч���� Number(5);
Begin

  --���ʧԼ��
  n_ԤԼ��Чʱ�� := Nvl(zl_GetSysParameter('ԤԼ��Чʱ��', 1111), 0);
  n_ʧԼ���ڹҺ� := Nvl(zl_GetSysParameter('ʧԼ���ڹҺ�', 1111), 0);
  n_�Һ���Ч���� := Nvl(zl_GetSysParameter('�Һ���Ч����'), 7);
  If n_ԤԼ��Чʱ�� <> 0 And n_ʧԼ���ڹҺ� <> 0 Then
    If ��¼id_In Is Null Then
      For c_ʧЧԤԼ In (Select b.��¼id, b.���, b.ԤԼ˳���
                     From ���˹Һż�¼ A, �ٴ�������ſ��� B
                     Where a.ԤԼʱ�� - 1 / 24 / 60 * n_ԤԼ��Чʱ�� < Sysdate And a.ԤԼʱ�� > Sysdate - n_�Һ���Ч���� And a.��¼���� = 2 And
                           a.�����¼id = b.��¼id And (a.���� = b.��� Or a.���� = b.��ע) And Nvl(b.�Һ�״̬, 0) = 2) Loop
        Update �ٴ�������ſ���
        Set �Һ�״̬ = 0, ����Ա���� = Null, ����վip = Null, ����վ���� = Null, ��ע = Null
        Where ��¼id = c_ʧЧԤԼ.��¼id And ��� = c_ʧЧԤԼ.��� And ԤԼ˳��� = c_ʧЧԤԼ.ԤԼ˳���;
      End Loop;
    Else
      For c_ʧЧԤԼ In (Select b.��¼id, b.���, b.ԤԼ˳���
                     From ���˹Һż�¼ A, �ٴ�������ſ��� B
                     Where a.ԤԼʱ�� - 1 / 24 / 60 * n_ԤԼ��Чʱ�� < Sysdate And a.ԤԼʱ�� > Sysdate - n_�Һ���Ч���� And a.��¼���� = 2 And
                           a.�����¼id = b.��¼id And (a.���� = b.��� Or a.���� = b.��ע) And b.��¼id = ��¼id_In And
                           Nvl(b.�Һ�״̬, 0) = 2) Loop
        If c_ʧЧԤԼ.ԤԼ˳��� Is Null Then
          Update �ٴ�������ſ���
          Set �Һ�״̬ = 0, ����Ա���� = Null, ����վip = Null, ����վ���� = Null, ��ע = Null
          Where ��¼id = c_ʧЧԤԼ.��¼id And ��� = c_ʧЧԤԼ.���;
        Else
          Update �ٴ�������ſ���
          Set �Һ�״̬ = 0, ����Ա���� = Null, ����վip = Null, ����վ���� = Null, ��ע = Null
          Where ��¼id = c_ʧЧԤԼ.��¼id And ��� = c_ʧЧԤԼ.��� And ԤԼ˳��� = c_ʧЧԤԼ.ԤԼ˳���;
        End If;
      End Loop;
    End If;
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�Һ����״̬_����_Delete;
/


CREATE OR REPLACE Procedure Zl_�Һ����״̬_Lock
(
  ����_In       Number, --1-����,2-�������
  ����Ա����_In �Һ����״̬.����Ա����%Type,
  ����_In       �Һ����״̬.����%Type := Null,
  ����_In       �Һ����״̬.����%Type := Null,
  ���_In       �Һ����״̬.���%Type := Null,
  �����¼ID_In �ٴ������¼.ID%type := Null
) As

  v_����       �Һ����״̬.����Ա����%Type;
  v_״̬       �Һ����״̬.״̬%Type;
  v_������     �Һ����״̬.������%Type;
  v_��֤������ �Һ����״̬.������%Type;
  v_����վIP   �ٴ�������ſ���.����վIP%Type;
  v_Error      Varchar2(255);
  Err_Custom Exception;
Begin
  Select Terminal Into v_������ From V$session Where Audsid = Userenv('sessionid');
  Select SYS_CONTEXT('USERENV','IP_ADDRESS') Into v_����վIP from dual;
  If ����_In = 1 Then
    --�����Һ����״̬
    If �����¼ID_In is Null then
      Begin
        Select ����Ա����, ״̬, ������
        Into v_����, v_״̬, v_��֤������
        From �Һ����״̬
        Where ���� = ����_In And ���� = ����_In And ��� = ���_In;
      Exception
        When Others Then
          Null;
      End;
      If v_���� Is Null Then
        Insert Into �Һ����״̬
          (����, ����, ���, ״̬, ����Ա����, ��ע, �Ǽ�ʱ��, ������)
        Values
          (����_In, ����_In, ���_In, 5, ����Ա����_In, '����������', Sysdate, v_������);
      Else
        v_Error := '���' || ���_In || '�ѱ�����Ա' || v_����;
        If v_״̬ = 1 Then
          v_Error := v_Error || 'ʹ��';
        Elsif v_״̬ = 2 Then
          v_Error := v_Error || 'ԤԼ';
        Elsif v_״̬ = 3 Then
          v_Error := v_Error || 'Ԥ��';
        Elsif v_״̬ = 4 Then
          v_Error := v_Error || '�˺�';
        Elsif v_״̬ = 5 Then
          v_Error := v_Error || '(' || v_��֤������ || ')����';
        End If;
        Raise Err_Custom;
      End If;
    Else
      Begin
        Select ����Ա����, �Һ�״̬, ����վ����
        Into v_����, v_״̬, v_��֤������
        From �ٴ�������ſ���
        Where ��¼ID = �����¼ID_In And ��� = ���_In;
      Exception
        When Others Then
          Null;
      End;
      
      If Nvl(v_״̬,0) = 0 Then
        Update �ٴ�������ſ��� set �Һ�״̬=5,����ʱ��=Sysdate,����Ա����=����Ա����_In,����վIP=v_����վIP,����վ����=v_������,��ע='����������'
        Where ��¼ID=�����¼ID_In  And ���=���_In;
      Else
        v_Error := '���' || ���_In || '�ѱ�����Ա' || v_����;
        If v_״̬ = 1 Then
          v_Error := v_Error || 'ʹ��';
        Elsif v_״̬ = 2 Then
          v_Error := v_Error || 'ԤԼ';
        Elsif v_״̬ = 3 Then
          v_Error := v_Error || 'Ԥ��';
        Elsif v_״̬ = 4 Then
          v_Error := v_Error || '�˺�';
        Elsif v_״̬ = 5 Then
          v_Error := v_Error || '(' || v_��֤������ || ')����';
        End If;
        Raise Err_Custom;
      End If;
    End If;
  Elsif ����_In = 2 Then
    If �����¼ID_In is Null then
       Delete �Һ����״̬ Where ������ = v_������ And ����Ա���� = ����Ա����_In And ״̬ = 5;
    Else
      Update �ٴ�������ſ��� A set A.�Һ�״̬=0,A.����ʱ��=NULL,A.����Ա����=NULL,A.����վIP=NULL,A.����վ����=NULL,A.����=NULL,A.����=NULL,A.��ע=NULL
      Where A.����վ���� =v_������ And A.����վIP=v_����վIP And A.����Ա���� = ����Ա����_In And A.�Һ�״̬ = 5 And A.����ʱ�� > Sysdate -1
        And Exists (Select 1 From �ٴ������¼ B Where A.��¼ID=B.ID And B.�Ƿ���ſ��� = 1);
      
      Update �ٴ�������ſ��� A set A.�Һ�״̬=4,A.����ʱ��=NULL,A.����Ա����=NULL,A.����վIP=NULL,A.����վ����=NULL,A.����=NULL,A.����=NULL,A.��ע=NULL
      Where A.����վ���� =v_������ And A.����վIP=v_����վIP And A.����Ա���� = ����Ա����_In And A.�Һ�״̬ = 5 And A.����ʱ�� > Sysdate -1
        And Exists (Select 1 From �ٴ������¼ B Where A.��¼ID=B.ID And B.�Ƿ���ſ��� = 0);
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�Һ����״̬_Lock;
/

CREATE OR REPLACE Function NextReservationNum
(
  ��¼ID_In         In �ٴ�������ſ���.��¼ID%Type,
  ���_In           In �ٴ�������ſ���.���%Type,
  ����Ա����_In     In �ٴ�������ſ���.����Ա����%Type
) Return Varchar2
 --��ȡ���ԤԼ˳��ţ�ֻ���ԤԼ��ͨ��ʱ��
 Is
  Pragma Autonomous_Transaction;
  v_������    �ٴ�������ſ���.����վ����%Type;
  v_����վIP  �ٴ�������ſ���.����վIP%Type;
  n_����      �ٴ�������ſ���.����%Type;
  n_��Լ��    �ٴ�������ſ���.����%Type;
  n_Maxno  Number;

  v_Error      Varchar2(255);
  Err_Custom Exception;
Begin
  Select Terminal Into v_������ From V$session Where Audsid = Userenv('sessionid');
  Select SYS_CONTEXT('USERENV','IP_ADDRESS') Into v_����վIP from dual;
  Begin
     Select A.����,B.��Լ��
     Into n_����,n_��Լ��
     From �ٴ�������ſ��� A,
      (Select ��¼ID,���,Count(1) as ��Լ�� From �ٴ�������ſ��� Where ��¼ID=��¼ID_In And ���=���_In And �Һ�״̬<>0 and �Һ�״̬<>4 And ԤԼ˳��� is Not Null group by ��¼ID,���) B
     Where A.��¼ID = B.��¼ID(+) And A.��� = B.���(+) And A.��¼ID=��¼ID_In And A.���=���_In And A.ԤԼ˳��� is Null;
  Exception
    When Others Then
      v_Error:='û�ҵ���Ӧ�ĳ��ﰲ�ż�¼';
      Raise Err_Custom;
  End;
  
  If Nvl(n_��Լ��,0)<Nvl(n_����,0) Then
    Select Nvl(Max(ԤԼ˳���),0) Into n_Maxno From �ٴ�������ſ��� WHERE ��¼ID=��¼ID_In  And ���=���_In;
    --If n_�Һ����=0 then
      n_Maxno:=n_Maxno+1;
      Insert Into �ٴ�������ſ���(��¼ID,���,ԤԼ˳���,��ʼʱ��,��ֹʱ��,����,�Ƿ�ԤԼ,�Һ�״̬,����ʱ��,����,����,����Ա����,����վIP,����վ����,��ע)
      Select ��¼ID,���,n_Maxno,��ʼʱ��,��ֹʱ��,1,�Ƿ�ԤԼ,5,Sysdate,����,����,����Ա����_In,v_����վIP,v_������,'����������'
      From �ٴ�������ſ���
      Where ��¼ID=��¼ID_In  And ���=���_In And ԤԼ˳��� is Null;
  Else
      v_Error:='��ǰʱ��ԤԼ�ѳ��������Լ��';
      Raise Err_Custom;
  End If;
  Commit;
  Return n_Maxno;
Exception
  When Err_Custom Then
    Rollback;
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Rollback;
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End NextReservationNum;
/


--����ZL1_INSIDE_1114_1/�̶������
Insert Into zlReports(ID,���,����,˵��,����,��ֽ,��ӡ��,Ʊ��,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��,��ӡ��ʽ,��ֹ��ʼʱ��,��ֹ����ʱ��) Values(zlReports_ID.NextVal,'ZL1_INSIDE_1114_1','�̶������','�̶������','I`;g$oi|}90Fiql4H+LL',15,Null,0,&n_System,1114,'�̶������',Sysdate,Sysdate,0,To_Date('2016-03-01 00:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('2016-03-01 00:00:00','YYYY-MM-DD HH24:MI:SS'));
Insert Into zlRPTFmts(����ID,���,˵��,ͼ��,W,H,ֽ��,ֽ��,��ֽ̬��) Values(zlReports_ID.CurrVal,1,'�̶������',0,11904,16832,9,1,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����,˵��) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'�������','�������,202',User||'.�ٴ������',0,Null);
Insert Into zlRPTSQLs(ԴID,�к�,����)  Select zlRPTDatas_ID.CurrVal,a.* From (
  Select 1,'Select ������� From �ٴ������ Where �Ű෽ʽ=0 And ID=[0]' From Dual) a;
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����,����) Values(zlRPTDatas_ID.CurrVal,Null,0,'����ID',1,Null,0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����,˵��) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'�ٴ������_����','��������,202|��Ŀ����,202|ҽ������,202|��һ,202|�ܶ�,202|����,202|����,202|����,202|����,202|����,202',User||'.�ٴ������,'||User||'.�ٴ����ﰲ��,'||User||'.�ٴ���������,'||User||'.�ٴ������Դ,'||User||'.���ű�,'||User||'.�շ���ĿĿ¼',0,Null);
Insert Into zlRPTSQLs(ԴID,�к�,����)  Select zlRPTDatas_ID.CurrVal,a.* From (
  Select 1,'Select e.���� As ��������, f.���� As ��Ŀ����, b.ҽ������, Max(Decode(c.������Ŀ, ''��һ'', c.�ϰ�ʱ��, Null)) As ��һ,' From Dual Union All
  Select 2,'       Max(Decode(c.������Ŀ, ''��һ'', c.�ϰ�ʱ��, Null)) As �ܶ�, Max(Decode(c.������Ŀ, ''��һ'', c.�ϰ�ʱ��, Null)) As ����,' From Dual Union All
  Select 3,'       Max(Decode(c.������Ŀ, ''��һ'', c.�ϰ�ʱ��, Null)) As ����, Max(Decode(c.������Ŀ, ''��һ'', c.�ϰ�ʱ��, Null)) As ����,' From Dual Union All
  Select 4,'       Max(Decode(c.������Ŀ, ''��һ'', c.�ϰ�ʱ��, Null)) As ����, Max(Decode(c.������Ŀ, ''��һ'', c.�ϰ�ʱ��, Null)) As ����' From Dual Union All
  Select 5,'From �ٴ������ A, �ٴ����ﰲ�� B, �ٴ��������� C, �ٴ������Դ D, ���ű� E, �շ���ĿĿ¼ F' From Dual Union All
  Select 6,'Where a.Id = b.����id And b.Id = c.����id(+) And b.��Դid = d.Id And d.����id = e.Id And b.��Ŀid = f.Id And a.�Ű෽ʽ = 0 And' From Dual Union All
  Select 7,'      a.Id = [0]' From Dual Union All
  Select 8,'Group By e.����, f.����, b.ҽ������' From Dual Union All
  Select 9,'Order By e.����, f.����, b.ҽ������' From Dual) a;
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����,����) Values(zlRPTDatas_ID.CurrVal,Null,0,'����ID',1,Null,0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ2',2,Null,0,'�����1',21,'������:[����Ա����]',Null,150,6460,1710,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ1',2,Null,0,Null,0,'[�������.�������]',Null,3960,435,2895,300,0,0,1,'����',14,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ3',2,Null,0,'�����1',23,'��������:[yyyy-mm-dd]',Null,9825,6460,1890,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'�����1',4,Null,0,Null,0,Null,Null,150,930,11565,5430,255,0,0,'����',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-1,0,Null,Null,'[�ٴ������_����.��������]','4^255^����^0^0',0,0,1320,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-2,1,Null,Null,'[�ٴ������_����.��Ŀ����]','4^255^��Ŀ^0^0',0,0,1800,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-3,2,Null,Null,'[�ٴ������_����.ҽ������]','4^255^ҽ��^0^0',0,0,1110,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-4,3,Null,Null,'[�ٴ������_����.��һ]','4^255^��һ^0^0',0,0,660,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-5,4,Null,Null,'[�ٴ������_����.�ܶ�]','4^255^�ܶ�^0^0',0,0,600,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-6,5,Null,Null,'[�ٴ������_����.����]','4^255^����^0^0',0,0,630,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-7,6,Null,Null,'[�ٴ������_����.����]','4^255^����^0^0',0,0,570,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-8,7,Null,Null,'[�ٴ������_����.����]','4^255^����^0^0',0,0,630,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-9,8,Null,Null,'[�ٴ������_����.����]','4^255^����^0^0',0,0,585,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-10,9,Null,Null,'[�ٴ������_����.����]','4^255^����^0^0',0,0,600,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);

--����ZL1_INSIDE_1114_1/�̶������
Insert into zlProgFuncs(ϵͳ,���,����,˵��) Values(&n_System,1114,'�̶������','�̶������');
Insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
  Select &n_System,1114,'�̶������',User,'���ű�','SELECT' From Dual Union All
  Select &n_System,1114,'�̶������',User,'�ٴ����ﰲ��','SELECT' From Dual Union All
  Select &n_System,1114,'�̶������',User,'�ٴ������','SELECT' From Dual Union All
  Select &n_System,1114,'�̶������',User,'�ٴ������Դ','SELECT' From Dual Union All
  Select &n_System,1114,'�̶������',User,'�ٴ���������','SELECT' From Dual Union All
  Select &n_System,1114,'�̶������',User,'�շ���ĿĿ¼','SELECT' From Dual;
  
--����ZL1_INSIDE_1114_2/�³����
Insert Into zlReports(ID,���,����,˵��,����,��ֽ,��ӡ��,Ʊ��,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��,��ӡ��ʽ,��ֹ��ʼʱ��,��ֹ����ʱ��) Values(zlReports_ID.NextVal,'ZL1_INSIDE_1114_2','�³����','�³����','Wg"|?kw}~8-@sht1V+LL',15,'������ OneNote 2010',0,&n_System,1114,'�³����',Sysdate,Sysdate,0,To_Date('2016-03-01 00:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('2016-03-01 00:00:00','YYYY-MM-DD HH24:MI:SS'));
Insert Into zlRPTFmts(����ID,���,˵��,ͼ��,W,H,ֽ��,ֽ��,��ֽ̬��) Values(zlReports_ID.CurrVal,1,'�³����(31��)',0,21563,11906,256,1,0);
Insert Into zlRPTFmts(����ID,���,˵��,ͼ��,W,H,ֽ��,ֽ��,��ֽ̬��) Values(zlReports_ID.CurrVal,2,'�³����(30��)',0,20843,11906,256,1,0);
Insert Into zlRPTFmts(����ID,���,˵��,ͼ��,W,H,ֽ��,ֽ��,��ֽ̬��) Values(zlReports_ID.CurrVal,3,'�³����(29��)',0,20258,11906,256,1,0);
Insert Into zlRPTFmts(����ID,���,˵��,ͼ��,W,H,ֽ��,ֽ��,��ֽ̬��) Values(zlReports_ID.CurrVal,4,'�³����(28��)',0,19778,11906,256,1,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����,˵��) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'�������','�������,202',User||'.�ٴ������',0,Null);
Insert Into zlRPTSQLs(ԴID,�к�,����)  Select zlRPTDatas_ID.CurrVal,a.* From (
  Select 1,'Select ������� From �ٴ������ Where ID=[0] And �Ű෽ʽ=1' From Dual) a;
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����,����) Values(zlRPTDatas_ID.CurrVal,Null,0,'����ID',1,Null,0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����,˵��) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'�ٴ������_����','��������,202|��Ŀ����,202|ҽ������,202|C1,202|C2,202|C3,202|C4,202|C5,202|C6,202|C7,202|C8,202|C9,202|C10,202|C11,202|C12,202|C13,202|C14,202|C15,202|C16,202|C17,202|C18,202|C19,202|C20,202|C21,202|C22,202|C23,202|C24,202|C25,202|C26,202|C27,202|C28,202|C29,202|C30,202|C31,202',User||'.�ٴ������,'||User||'.�ٴ����ﰲ��,'||User||'.�ٴ������¼,'||User||'.�ٴ������Դ,'||User||'.���ű�,'||User||'.�շ���ĿĿ¼',1,Null);
Insert Into zlRPTSQLs(ԴID,�к�,����)  Select zlRPTDatas_ID.CurrVal,a.* From (
  Select 1,'Select e.���� As ��������, f.���� As ��Ŀ����, b.ҽ������, Max(Decode(To_Char(c.��������, ''DD''), ''01'', c.�ϰ�ʱ��, Null)) As C1,' From Dual Union All
  Select 2,'       Max(Decode(To_Char(c.��������, ''DD''), ''02'', c.�ϰ�ʱ��, Null)) As C2,' From Dual Union All
  Select 3,'       Max(Decode(To_Char(c.��������, ''DD''), ''03'', c.�ϰ�ʱ��, Null)) As C3,' From Dual Union All
  Select 4,'       Max(Decode(To_Char(c.��������, ''DD''), ''04'', c.�ϰ�ʱ��, Null)) As C4,' From Dual Union All
  Select 5,'       Max(Decode(To_Char(c.��������, ''DD''), ''05'', c.�ϰ�ʱ��, Null)) As C5,' From Dual Union All
  Select 6,'       Max(Decode(To_Char(c.��������, ''DD''), ''06'', c.�ϰ�ʱ��, Null)) As C6,' From Dual Union All
  Select 7,'       Max(Decode(To_Char(c.��������, ''DD''), ''07'', c.�ϰ�ʱ��, Null)) As C7,' From Dual Union All
  Select 8,'       Max(Decode(To_Char(c.��������, ''DD''), ''08'', c.�ϰ�ʱ��, Null)) As C8,' From Dual Union All
  Select 9,'       Max(Decode(To_Char(c.��������, ''DD''), ''09'', c.�ϰ�ʱ��, Null)) As C9,' From Dual Union All
  Select 10,'       Max(Decode(To_Char(c.��������, ''DD''), ''10'', c.�ϰ�ʱ��, Null)) As C10,' From Dual Union All
  Select 11,'       Max(Decode(To_Char(c.��������, ''DD''), ''11'', c.�ϰ�ʱ��, Null)) As C11,' From Dual Union All
  Select 12,'       Max(Decode(To_Char(c.��������, ''DD''), ''12'', c.�ϰ�ʱ��, Null)) As C12,' From Dual Union All
  Select 13,'       Max(Decode(To_Char(c.��������, ''DD''), ''13'', c.�ϰ�ʱ��, Null)) As C13,' From Dual Union All
  Select 14,'       Max(Decode(To_Char(c.��������, ''DD''), ''14'', c.�ϰ�ʱ��, Null)) As C14,' From Dual Union All
  Select 15,'       Max(Decode(To_Char(c.��������, ''DD''), ''15'', c.�ϰ�ʱ��, Null)) As C15,' From Dual Union All
  Select 16,'       Max(Decode(To_Char(c.��������, ''DD''), ''16'', c.�ϰ�ʱ��, Null)) As C16,' From Dual Union All
  Select 17,'       Max(Decode(To_Char(c.��������, ''DD''), ''17'', c.�ϰ�ʱ��, Null)) As C17,' From Dual Union All
  Select 18,'       Max(Decode(To_Char(c.��������, ''DD''), ''18'', c.�ϰ�ʱ��, Null)) As C18,' From Dual Union All
  Select 19,'       Max(Decode(To_Char(c.��������, ''DD''), ''19'', c.�ϰ�ʱ��, Null)) As C19,' From Dual Union All
  Select 20,'       Max(Decode(To_Char(c.��������, ''DD''), ''20'', c.�ϰ�ʱ��, Null)) As C20,' From Dual Union All
  Select 21,'       Max(Decode(To_Char(c.��������, ''DD''), ''21'', c.�ϰ�ʱ��, Null)) As C21,' From Dual Union All
  Select 22,'       Max(Decode(To_Char(c.��������, ''DD''), ''22'', c.�ϰ�ʱ��, Null)) As C22,' From Dual Union All
  Select 23,'       Max(Decode(To_Char(c.��������, ''DD''), ''23'', c.�ϰ�ʱ��, Null)) As C23,' From Dual Union All
  Select 24,'       Max(Decode(To_Char(c.��������, ''DD''), ''24'', c.�ϰ�ʱ��, Null)) As C24,' From Dual Union All
  Select 25,'       Max(Decode(To_Char(c.��������, ''DD''), ''25'', c.�ϰ�ʱ��, Null)) As C25,' From Dual Union All
  Select 26,'       Max(Decode(To_Char(c.��������, ''DD''), ''26'', c.�ϰ�ʱ��, Null)) As C26,' From Dual Union All
  Select 27,'       Max(Decode(To_Char(c.��������, ''DD''), ''27'', c.�ϰ�ʱ��, Null)) As C27,' From Dual Union All
  Select 28,'       Max(Decode(To_Char(c.��������, ''DD''), ''28'', c.�ϰ�ʱ��, Null)) As C28,' From Dual Union All
  Select 29,'       Max(Case' From Dual Union All
  Select 30,'             When To_Number(To_Char(Last_Day(c.��������), ''DD'')) < 29 Then' From Dual Union All
  Select 31,'              Null' From Dual Union All
  Select 32,'             Else' From Dual Union All
  Select 33,'              Decode(To_Char(c.��������, ''DD''), ''29'', c.�ϰ�ʱ��, '' '')' From Dual Union All
  Select 34,'           End) As C29,' From Dual Union All
  Select 35,'       Max(Case' From Dual Union All
  Select 36,'             When To_Number(To_Char(Last_Day(c.��������), ''DD'')) < 30 Then' From Dual Union All
  Select 37,'              Null' From Dual Union All
  Select 38,'             Else' From Dual Union All
  Select 39,'              Decode(To_Char(c.��������, ''DD''), ''30'', c.�ϰ�ʱ��, '' '')' From Dual Union All
  Select 40,'           End) As C30,' From Dual Union All
  Select 41,'       Max(Case' From Dual Union All
  Select 42,'             When To_Number(To_Char(Last_Day(c.��������), ''DD'')) < 31 Then' From Dual Union All
  Select 43,'              Null' From Dual Union All
  Select 44,'             Else' From Dual Union All
  Select 45,'              Decode(To_Char(c.��������, ''DD''), ''31'', c.�ϰ�ʱ��, '' '')' From Dual Union All
  Select 46,'           End) As C31' From Dual Union All
  Select 47,'From �ٴ����ﰲ�� B, �ٴ������ A, �ٴ������¼ C, �ٴ������Դ D, ���ű� E, �շ���ĿĿ¼ F' From Dual Union All
  Select 48,'Where b.Id = c.����id(+) And b.��Դid = d.Id And b.����id = a.Id And b.��Ŀid = f.Id And' From Dual Union All
  Select 49,'      d.����id = e.Id And a.�Ű෽ʽ = 1 And a.Id = [0]' From Dual Union All
  Select 50,'Group By e.����, f.����, b.ҽ������' From Dual Union All
  Select 51,'Order By e.����, f.����, b.ҽ������' From Dual Union All
  Select 52,Null From Dual) a;
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����,����) Values(zlRPTDatas_ID.CurrVal,Null,0,'����ID',1,Null,0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ2',2,Null,0,'�����1',21,'������:[����Ա����]',Null,150,6520,1710,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ1',2,Null,0,'�����1',12,'[�������.�������]',Null,9302,165,3150,300,0,0,1,'����',14,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ3',2,Null,0,'�����1',23,'����ʱ��:[yyyy-mm-dd]',Null,19458,6520,1890,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'�����1',4,Null,0,Null,0,Null,Null,150,615,21198,5805,255,0,0,'����',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-1,0,Null,Null,'[�ٴ������_����.��������]','4^300^����^0^0',0,0,1305,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-2,1,Null,Null,'[�ٴ������_����.��Ŀ����]','4^300^��Ŀ^0^0',0,0,1965,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-3,2,Null,Null,'[�ٴ������_����.ҽ������]','4^300^ҽ��^0^0',0,0,1110,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-4,3,Null,Null,'[�ٴ������_����.C1]','4^300^1^0^0',0,0,525,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-5,4,Null,Null,'[�ٴ������_����.C2]','4^300^2^0^0',0,0,555,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-6,5,Null,Null,'[�ٴ������_����.C3]','4^300^3^0^0',0,0,495,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-7,6,Null,Null,'[�ٴ������_����.C4]','4^300^4^0^0',0,0,480,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-8,7,Null,Null,'[�ٴ������_����.C5]','4^300^5^0^0',0,0,495,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-9,8,Null,Null,'[�ٴ������_����.C6]','4^300^6^0^0',0,0,495,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-10,9,Null,Null,'[�ٴ������_����.C7]','4^300^7^0^0',0,0,510,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-11,10,Null,Null,'[�ٴ������_����.C8]','4^300^8^0^0',0,0,465,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-12,11,Null,Null,'[�ٴ������_����.C9]','4^300^9^0^0',0,0,510,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-13,12,Null,Null,'[�ٴ������_����.C10]','4^300^10^0^0',0,0,465,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-14,13,Null,Null,'[�ٴ������_����.C11]','4^300^11^0^0',0,0,480,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-15,14,Null,Null,'[�ٴ������_����.C12]','4^300^12^0^0',0,0,525,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-16,15,Null,Null,'[�ٴ������_����.C13]','4^300^13^0^0',0,0,540,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-17,16,Null,Null,'[�ٴ������_����.C14]','4^300^14^0^0',0,0,525,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-18,17,Null,Null,'[�ٴ������_����.C15]','4^300^15^0^0',0,0,525,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-19,18,Null,Null,'[�ٴ������_����.C16]','4^300^16^0^0',0,0,540,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-20,19,Null,Null,'[�ٴ������_����.C17]','4^300^17^0^0',0,0,540,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-21,20,Null,Null,'[�ٴ������_����.C18]','4^300^18^0^0',0,0,570,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-22,21,Null,Null,'[�ٴ������_����.C19]','4^300^19^0^0',0,0,525,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-23,22,Null,Null,'[�ٴ������_����.C20]','4^300^20^0^0',0,0,525,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-24,23,Null,Null,'[�ٴ������_����.C21]','4^300^21^0^0',0,0,570,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-25,24,Null,Null,'[�ٴ������_����.C22]','4^300^22^0^0',0,0,585,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-26,25,Null,Null,'[�ٴ������_����.C23]','4^300^23^0^0',0,0,600,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-27,26,Null,Null,'[�ٴ������_����.C24]','4^300^24^0^0',0,0,615,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-28,27,Null,Null,'[�ٴ������_����.C25]','4^300^25^0^0',0,0,570,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-29,28,Null,Null,'[�ٴ������_����.C26]','4^300^26^0^0',0,0,540,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-30,29,Null,Null,'[�ٴ������_����.C27]','4^300^27^0^0',0,0,570,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-31,30,Null,Null,'[�ٴ������_����.C28]','4^300^28^0^0',0,0,585,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-32,31,Null,Null,'[�ٴ������_����.C29]','4^300^29^0^0',0,0,525,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-33,32,Null,Null,'[�ٴ������_����.C30]','4^300^30^0^0',0,0,585,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-34,33,Null,Null,'[�ٴ������_����.C31]','4^300^31^0^0',0,0,585,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null);

--����ZL1_INSIDE_1114_2/�³����
Insert into zlProgFuncs(ϵͳ,���,����,˵��) Values(&n_System,1114,'�³����','�³����');
Insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
  Select &n_System,1114,'�³����',User,'���ű�','SELECT' From Dual Union All
  Select &n_System,1114,'�³����',User,'�ٴ����ﰲ��','SELECT' From Dual Union All
  Select &n_System,1114,'�³����',User,'�ٴ������','SELECT' From Dual Union All
  Select &n_System,1114,'�³����',User,'�ٴ������Դ','SELECT' From Dual Union All
  Select &n_System,1114,'�³����',User,'�ٴ������¼','SELECT' From Dual Union All
  Select &n_System,1114,'�³����',User,'�շ���ĿĿ¼','SELECT' From Dual;


--����ZL1_INSIDE_1114_3/�ܳ����
Insert Into zlReports(ID,���,����,˵��,����,��ֽ,��ӡ��,Ʊ��,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��,��ӡ��ʽ,��ֹ��ʼʱ��,��ֹ����ʱ��) Values(zlReports_ID.NextVal,'ZL1_INSIDE_1114_3','�ܳ����','�ܳ����','Tg"}<kw}}8-@pht1T+LL',15,Null,0,&n_System,1114,'�ܳ����',Sysdate,Sysdate,0,To_Date('2016-03-01 00:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('2016-03-01 00:00:00','YYYY-MM-DD HH24:MI:SS'));
Insert Into zlRPTFmts(����ID,���,˵��,ͼ��,W,H,ֽ��,ֽ��,��ֽ̬��) Values(zlReports_ID.CurrVal,1,'�ܳ����',0,11904,16832,9,1,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����,˵��) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'�������','�������,202',User||'.�ٴ������',0,Null);
Insert Into zlRPTSQLs(ԴID,�к�,����)  Select zlRPTDatas_ID.CurrVal,a.* From (
  Select 1,'Select ������� From �ٴ������ Where �Ű෽ʽ=2 And ID=[0]' From Dual) a;
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����,����) Values(zlRPTDatas_ID.CurrVal,Null,0,'����ID',1,Null,0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����,˵��) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'�ٴ������_����','��������,202|��Ŀ����,202|ҽ������,202|C1,202|C2,202|C3,202|C4,202|C5,202|C6,202|C7,202',User||'.�ٴ������,'||User||'.�ٴ����ﰲ��,'||User||'.�ٴ������¼,'||User||'.�ٴ������Դ,'||User||'.���ű�,'||User||'.�շ���ĿĿ¼',0,Null);
Insert Into zlRPTSQLs(ԴID,�к�,����)  Select zlRPTDatas_ID.CurrVal,a.* From (
  Select 1,'Select e.���� As ��������, f.���� As ��Ŀ����, b.ҽ������, ' From Dual Union All
  Select 2,'       Max(Decode(To_Char(c.��������, ''D''), 2, c.�ϰ�ʱ��, Null)) As C1,' From Dual Union All
  Select 3,'       Max(Decode(To_Char(c.��������, ''D''), 3, c.�ϰ�ʱ��, Null)) As C2,' From Dual Union All
  Select 4,'       Max(Decode(To_Char(c.��������, ''D''), 4, c.�ϰ�ʱ��, Null)) As C3,' From Dual Union All
  Select 5,'       Max(Decode(To_Char(c.��������, ''D''), 5, c.�ϰ�ʱ��, Null)) As C4,' From Dual Union All
  Select 6,'       Max(Decode(To_Char(c.��������, ''D''), 6, c.�ϰ�ʱ��, Null)) As C5,' From Dual Union All
  Select 7,'       Max(Decode(To_Char(c.��������, ''D''), 7, c.�ϰ�ʱ��, Null)) As C6,' From Dual Union All
  Select 8,'       Max(Decode(To_Char(c.��������, ''D''), 1, c.�ϰ�ʱ��, Null)) As C7' From Dual Union All
  Select 9,'From �ٴ������ A, �ٴ����ﰲ�� B, �ٴ������¼ C, �ٴ������Դ D, ���ű� E, �շ���ĿĿ¼ F' From Dual Union All
  Select 10,'Where a.Id = b.����id And b.��Դid = d.Id And b.Id = c.����id(+) And d.����id = e.Id And b.��Ŀid = f.Id And a.�Ű෽ʽ = 2 And' From Dual Union All
  Select 11,'      a.Id = [0]' From Dual Union All
  Select 12,'Group By e.����, f.����, b.ҽ������' From Dual Union All
  Select 13,'Order By e.����, f.����, b.ҽ������' From Dual) a;
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����,����) Values(zlRPTDatas_ID.CurrVal,Null,0,'����ID',1,Null,0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����,˵��) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'����','C1,202|C2,202|C3,202|C4,202|C5,202|C6,202|C7,202',User||'.�ٴ������,'||User||'.�ٴ����ﰲ��,'||User||'.�ٴ������¼',0,Null);
Insert Into zlRPTSQLs(ԴID,�к�,����)  Select zlRPTDatas_ID.CurrVal,a.* From (
  Select 1,'Select To_Char(Trunc(��������, ''d'') + 1, ''DD'') As C1, To_Char(Trunc(��������, ''d'') + 2, ''DD'') As C2,' From Dual Union All
  Select 2,'       To_Char(Trunc(��������, ''d'') + 3, ''DD'') As C3, To_Char(Trunc(��������, ''d'') + 4, ''DD'') As C4,' From Dual Union All
  Select 3,'       To_Char(Trunc(��������, ''d'') + 5, ''DD'') As C5, To_Char(Trunc(��������, ''d'') + 6, ''DD'') As C6,' From Dual Union All
  Select 4,'       To_Char(Trunc(��������, ''d'') + 7, ''DD'') As C7' From Dual Union All
  Select 5,'From (Select c.��������' From Dual Union All
  Select 6,'       From �ٴ������ A, �ٴ����ﰲ�� B, �ٴ������¼ C' From Dual Union All
  Select 7,'       Where a.Id = b.����id And b.Id = c.����id(+) And a.�Ű෽ʽ = 2 And a.Id = [0]' From Dual Union All
  Select 8,'			And c.�������� Is Not Null And Rownum < 2)' From Dual Union All
  Select 9,Null From Dual) a;
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����,����) Values(zlRPTDatas_ID.CurrVal,Null,0,'����ID',1,Null,0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ2',2,Null,0,'�����1',21,'������:[����Ա����]',Null,210,4720,1710,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ1',2,Null,0,'�����1',12,'[�������.�������]',Null,4522,195,2895,300,0,0,1,'����',14,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ3',2,Null,0,'�����1',23,'��������:[yyyy-mm-dd]',Null,9840,4720,1890,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'�����1',4,Null,0,Null,0,Null,Null,210,660,11520,3960,255,0,0,'����',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-1,0,Null,Null,'[�ٴ������_����.��������]','4^600^����^0^0',0,0,1350,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-2,1,Null,Null,'[�ٴ������_����.��Ŀ����]','4^600^��Ŀ^0^0',0,0,1710,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-3,2,Null,Null,'[�ٴ������_����.ҽ������]','4^600^ҽ��^0^0',0,0,1350,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-4,3,Null,Null,'[�ٴ������_����.C1]','4^600^��һ
[����.C1]^0^0',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-5,4,Null,Null,'[�ٴ������_����.C2]','4^600^�ܶ�
[����.C2]^0^0',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-6,5,Null,Null,'[�ٴ������_����.C3]','4^600^����
[����.C3]^0^0',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-7,6,Null,Null,'[�ٴ������_����.C4]','4^600^����
[����.C4]^0^0',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-8,7,Null,Null,'[�ٴ������_����.C5]','4^600^����
[����.C5]^0^0',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-9,8,Null,Null,'[�ٴ������_����.C6]','4^600^����
[����.C6]^0^0',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-10,9,Null,Null,'[�ٴ������_����.C7]','4^600^����
[����.C7]^0^0',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);

--����ZL1_INSIDE_1114_3/�ܳ����
Insert into zlProgFuncs(ϵͳ,���,����,˵��) Values(&n_System,1114,'�ܳ����','�ܳ����');
Insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
  Select &n_System,1114,'�ܳ����',User,'���ű�','SELECT' From Dual Union All
  Select &n_System,1114,'�ܳ����',User,'�ٴ����ﰲ��','SELECT' From Dual Union All
  Select &n_System,1114,'�ܳ����',User,'�ٴ������','SELECT' From Dual Union All
  Select &n_System,1114,'�ܳ����',User,'�ٴ������Դ','SELECT' From Dual Union All
  Select &n_System,1114,'�ܳ����',User,'�ٴ������¼','SELECT' From Dual Union All
  Select &n_System,1114,'�ܳ����',User,'�շ���ĿĿ¼','SELECT' From Dual;


--����ZL1_INSIDE_1114_4/����ԤԼ�嵥
Insert Into zlReports(ID,���,����,˵��,����,��ֽ,��ӡ��,Ʊ��,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��,��ӡ��ʽ,��ֹ��ʼʱ��,��ֹ����ʱ��) Values(zlReports_ID.NextVal,'ZL1_INSIDE_1114_4','����ԤԼ�嵥','����ԤԼ�嵥','Lv!a7lom~"'||CHR(38)||'Fhyw*X,T\',15,'������ OneNote 2010',0,&n_System,1114,'����ԤԼ�嵥',Sysdate,Sysdate,0,To_Date('2016-04-01 00:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('2016-04-01 00:00:00','YYYY-MM-DD HH24:MI:SS'));
Insert Into zlRPTFmts(����ID,���,˵��,ͼ��,W,H,ֽ��,ֽ��,��ֽ̬��) Values(zlReports_ID.CurrVal,1,'����ԤԼ�嵥',0,11906,16838,9,2,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����,˵��) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'���˹Һż�¼_����','����,202|�Ա�,202|����,202|��ͥ��ַ,202|��ͥ�绰,202|�ű�,202|����,202|����,202|�շ���Ŀ,202|ҽ��,202|����ҽ��,202|ԤԼ����,202|ԤԼʱ��,135',User||'.���˹Һż�¼,'||User||'.������Ϣ,'||User||'.�ٴ������Դ,'||User||'.�ٴ������¼,'||User||'.���ű�,'||User||'.�շ���ĿĿ¼,'||User||'.���˷�����Ϣ��¼',0,Null);
Insert Into zlRPTSQLs(ԴID,�к�,����)  Select zlRPTDatas_ID.CurrVal,a.* From (
  Select 1,'Select a.����, a.�Ա�, a.����, b.��ͥ��ַ, b.��ͥ�绰, c.���� As �ű�, c.����, e.���� As ����, f.���� As �շ���Ŀ, d.ҽ������ As ҽ��, d.����ҽ������ As ����ҽ��,' From Dual Union All
  Select 2,'       a.No As ԤԼ����, a.����ʱ�� As ԤԼʱ��' From Dual Union All
  Select 3,'From ���˹Һż�¼ A, ������Ϣ B, �ٴ������Դ C, �ٴ������¼ D, ���ű� E, �շ���ĿĿ¼ F, ���˷�����Ϣ��¼ G' From Dual Union All
  Select 4,'Where a.id = g.�Һ�id And g.֪ͨ���� In (1,2) And a.��¼״̬ = 1 And a.����id = b.����id(+) And' From Dual Union All
  Select 5,'      g.��¼id = d.Id And d.��Դid = c.Id And d.����id = e.Id And d.��Ŀid = f.Id And g.��¼id In (Select Column_Value From Table(f_Str2list([0])))' From Dual Union All
  Select 6,Null From Dual) a;
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����,����) Values(zlRPTDatas_ID.CurrVal,Null,0,'�����¼IDS',0,Null,0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ1',2,Null,0,'�����1',12,'����ԤԼ�嵥',Null,7545,195,1800,300,0,0,1,'����',14,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'�����1',4,Null,0,Null,0,Null,Null,180,645,16530,4545,255,0,0,'����',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-1,6,Null,Null,'[���˹Һż�¼_����.����]','4^315^����',0,0,810,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-2,7,Null,Null,'[���˹Һż�¼_����.����]','4^315^����',0,0,1440,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-3,8,Null,Null,'[���˹Һż�¼_����.�շ���Ŀ]','4^315^�շ���Ŀ',0,0,1620,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-4,9,Null,Null,'[���˹Һż�¼_����.ҽ��]','4^315^ҽ��',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-5,10,Null,Null,'[���˹Һż�¼_����.����ҽ��]','4^315^����ҽ��',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-6,11,Null,Null,'[���˹Һż�¼_����.ԤԼ����]','4^315^ԤԼ����',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-7,12,Null,Null,'[���˹Һż�¼_����.ԤԼʱ��]','4^315^ԤԼʱ��',0,0,1665,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-8,0,Null,Null,'[���˹Һż�¼_����.����]','4^315^����',0,0,1140,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-9,1,Null,Null,'[���˹Һż�¼_����.�Ա�]','4^315^�Ա�',0,0,705,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-10,2,Null,Null,'[���˹Һż�¼_����.����]','4^315^����',0,0,855,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-11,3,Null,Null,'[���˹Һż�¼_����.��ͥ��ַ]','4^315^��ͥ��ַ',0,0,2850,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-12,4,Null,Null,'[���˹Һż�¼_����.��ͥ�绰]','4^315^��ϵ�绰',0,0,1500,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-13,5,Null,Null,'[���˹Һż�¼_����.�ű�]','4^315^�ű�',0,0,840,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);

--����ZL1_INSIDE_1114_4/����ԤԼ�嵥
Insert into zlProgFuncs(ϵͳ,���,����,˵��) Values(&n_System,1114,'����ԤԼ�嵥','����ԤԼ�嵥');
Insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
  Select &n_System,1114,'����ԤԼ�嵥',User,'���˷�����Ϣ��¼','SELECT' From Dual Union All
  Select &n_System,1114,'����ԤԼ�嵥',User,'���˹Һż�¼','SELECT' From Dual Union All
  Select &n_System,1114,'����ԤԼ�嵥',User,'������Ϣ','SELECT' From Dual Union All
  Select &n_System,1114,'����ԤԼ�嵥',User,'���ű�','SELECT' From Dual Union All
  Select &n_System,1114,'����ԤԼ�嵥',User,'�ٴ������Դ','SELECT' From Dual Union All
  Select &n_System,1114,'����ԤԼ�嵥',User,'�ٴ������¼','SELECT' From Dual Union All
  Select &n_System,1114,'����ԤԼ�嵥',User,'�շ���ĿĿ¼','SELECT' From Dual;





