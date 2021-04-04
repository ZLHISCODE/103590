
--1.��������Ϣ
Create Table ��������Ϣ(
    ����� VARCHAR2(20),   
    ҽ��ID Number(18),     
    ������� Number(1),
    ��ǰ���� Number(2) default 0,
    �޼����� Varchar2(2048),
    ʣ��λ�� Varchar2(64),
    �걾���� Varchar2(10),
    ��Ƭ���� Varchar2(10))
    TABLESPACE zl9BaseItem; 
    
    
Alter Table ��������Ϣ Add Constraint ��������Ϣ_PK Primary Key (�����) Using Index Tablespace zl9indexhis;    
Alter Table ��������Ϣ Add Constraint ��������Ϣ_FK_ҽ��ID Foreign Key (ҽ��ID) References ����ҽ����¼(ID) On Delete Cascade;  
Create Index ��������Ϣ_IX_ҽ��ID On ��������Ϣ(ҽ��ID) Pctfree 5 Tablespace zl9indexhis;   
Create Sequence ��������Ϣ_����� Start With 1;



--2.����걾��Ϣ
Create Table ����걾��Ϣ(
    �걾ID NUMBER(18),
    ҽ��ID Number(18),
    �걾���� VARCHAR2(64) Not Null,
    ������� NUMBER(1) default 0,
    �걾���� NUMBER(1) default 0,
    �ɼ���λ VARCHAR2(20),
    ԭ�б�� VARCHAR2(20),
    ���� Number(2) Not Null,
    ���λ�� VARCHAR2(64),
    �������� Date,
    ��ע VARCHAR2(1024))
    TABLESPACE zl9BaseItem;    
    
Alter Table ����걾��Ϣ Add Constraint ����걾��Ϣ_PK Primary Key (�걾ID) Using Index Tablespace zl9indexhis;    
Alter Table ����걾��Ϣ Add Constraint ����걾��Ϣ_FK_ҽ��ID Foreign Key (ҽ��ID) References ����ҽ����¼(ID) On Delete Cascade;
Create Index ����걾��Ϣ_IX_ҽ��ID On ����걾��Ϣ(ҽ��ID) Pctfree 5 Tablespace zl9indexhis;       
Create Sequence ����걾��Ϣ_�걾ID Start With 1;   


--3.�����ͼ���Ϣ
Create Table �����ͼ���Ϣ(
    ID NUMBER(18),   
    ҽ��ID NUMBER(18),
    �ͼ쵥λ VARCHAR2(64),
    �ͼ���� VARCHAR2(64),
    �ͼ��� VARCHAR2(64) Not Null,
    �ͼ����� DATE Not Null,
    ��ϵ��ʽ VARCHAR2(64),
    �Ǽ��� VARCHAR2(64) Not Null,
    ����״̬ NUMBER(1) default 1,
    ����ԭ�� VARCHAR2(1024),
    ֪ͨ�� VARCHAR2(64),
    ��ע VARCHAR2(1024))
    TABLESPACE zl9BaseItem;
    

Alter Table �����ͼ���Ϣ Add Constraint �����ͼ���Ϣ_PK Primary Key (ID) Using Index Tablespace zl9indexhis;    
Alter Table �����ͼ���Ϣ Add Constraint �����ͼ���Ϣ_FK_ҽ��ID Foreign Key (ҽ��ID) References ����ҽ����¼(ID) On Delete Cascade;   
Create Index �����ͼ���Ϣ_IX_ҽ��ID On �����ͼ���Ϣ(ҽ��ID) Pctfree 5 Tablespace zl9indexhis; 
Create Sequence �����ͼ���Ϣ_ID Start With 1;  
 
    
--4.����������Ϣ    
Create Table ����������Ϣ(
    ����ID Number(18),  
    ����� Varchar2(20),  
    ������ Varchar2(64) Not Null,
    ����ʱ�� Date,        
    �������� Number(1) default 0,
    ����״̬ Number(1) default 0,
    �������� Varchar2(1024),
    �Ƿ��ӡ Number(1) default 0,
    ���ʱ�� Date)
    TABLESPACE zl9BaseItem;    
    
    
Alter Table ����������Ϣ Add Constraint ����������Ϣ_PK Primary Key (����ID) Using Index Tablespace zl9indexhis;  
Alter Table ����������Ϣ Add Constraint ����������Ϣ_FK_����� Foreign Key (�����) References ��������Ϣ(�����) On Delete Cascade;
Create Index ����������Ϣ_IX_����� On ����������Ϣ(�����) Pctfree 5 Tablespace zl9indexhis;
Create Sequence ����������Ϣ_����ID Start With 1;  
  
    
--5.����ȡ����Ϣ    
Create Table ����ȡ����Ϣ(
    �Ŀ�ID Number(18),
    ��� Number(18),
    ����� Varchar2(20),
    ����ID Number(18),
    �걾ID Number(18),
    �걾���� Varchar2(64),
    ȡ��λ�� Varchar2(64),
    ��״ Varchar2(64),
    ��ɫ Varchar2(20),
    ���� Varchar2(20),
    �걾�� Varchar2(20),
    ������ Number(2) default 1,   
    �Ƿ���� Number(1) default 0,
    ��ȡҽʦ Varchar2(64) Not Null,
    ��ȡҽʦ Varchar2(64),
    ��¼ҽʦ Varchar2(64) Not Null,
    ȡ��ʱ�� Date)
    TABLESPACE zl9BaseItem;   
    
    
Alter Table ����ȡ����Ϣ Add Constraint ����ȡ����Ϣ_PK Primary Key (�Ŀ�ID) Using Index Tablespace zl9indexhis;    
Alter Table ����ȡ����Ϣ Add Constraint ����ȡ����Ϣ_FK_����� Foreign Key (�����) References ��������Ϣ(�����) On Delete Cascade; 
--Alter Table ����ȡ����Ϣ Add Constraint ����ȡ����Ϣ_FK_����ID Foreign Key (����ID) References ����������Ϣ(����ID) On Delete Cascade; --����ȡ��û��������Ϣ
Alter Table ����ȡ����Ϣ Add Constraint ����ȡ����Ϣ_FK_�걾ID Foreign Key (�걾ID) References ����걾��Ϣ(�걾ID) On Delete Cascade; 
Create Index ����ȡ����Ϣ_IX_����� On ����ȡ����Ϣ(�����) Pctfree 5 Tablespace zl9indexhis; 
Create Index ����ȡ����Ϣ_IX_����ID On ����ȡ����Ϣ(����ID) Pctfree 5 Tablespace zl9indexhis; 
Alter Table ����ȡ����Ϣ Add Constraint ����ȡ����Ϣ_CK_�Ƿ���� Check (�Ƿ���� IN(0,1));
Create Sequence ����ȡ����Ϣ_�Ŀ�ID Start With 1; 
  
      
    
--6.�����Ѹ���Ϣ    
Create Table �����Ѹ���Ϣ(
    ID Number(18),   
    �걾ID Number(18),
    ��ʼʱ�� Date,
    ����ʱ�� Number(5),
    ��ǰ�״� Number(2),
    ���״̬ Number(1) default 0,
    ����Ա Varchar2(64))
    TABLESPACE zl9BaseItem;     
    
    
Alter Table �����Ѹ���Ϣ Add Constraint �����Ѹ���Ϣ_PK Primary Key (ID) Using Index Tablespace zl9indexhis;    
Alter Table �����Ѹ���Ϣ Add Constraint �����Ѹ���Ϣ_FK_�걾ID Foreign Key (�걾ID) References ����걾��Ϣ(�걾ID) On Delete Cascade;   
Create Index �����Ѹ���Ϣ_IX_�걾ID On �����Ѹ���Ϣ(�걾ID) Pctfree 5 Tablespace zl9indexhis;    
Create Sequence �����Ѹ���Ϣ_ID Start With 1; 


--7.������Ƭ��Ϣ
Create Table ������Ƭ��Ϣ(
    ID Number(18),  
    ����� Varchar(20), 
    �Ŀ�ID Number(18),
    ����ID Number(18),
    ��Ƭ���� Number(1) default 0,
    ��Ƭ��ʽ Number(1) default 0,
    ��Ƭʱ�� Date,
    ��Ƭ�� Number(2),
    ��Ƭ�� Varchar2(64),       
    ��ǰ״̬ Number(1) default 0,
    �嵥״̬ Number(1) default 0)
    TABLESPACE zl9BaseItem;     
    
    
Alter Table ������Ƭ��Ϣ Add Constraint ������Ƭ��Ϣ_PK Primary Key (ID) Using Index Tablespace zl9indexhis;   
Alter table ������Ƭ��Ϣ Add constraint ������Ƭ��Ϣ_FK_����� Foreign key(�����) References ��������Ϣ(�����) On Delete Cascade;   
Alter Table ������Ƭ��Ϣ Add Constraint ������Ƭ��Ϣ_FK_�Ŀ�ID Foreign Key (�Ŀ�ID) References ����ȡ����Ϣ(�Ŀ�ID) On Delete Cascade;  
Create Index ������Ƭ��Ϣ_IX_�Ŀ�ID On ������Ƭ��Ϣ(�Ŀ�ID) Pctfree 5 Tablespace zl9indexhis;     
Create Index ������Ƭ��Ϣ_IX_����ID On ������Ƭ��Ϣ(����ID, �Ŀ�ID) Pctfree 5 Tablespace zl9indexhis;    
Create Sequence ������Ƭ��Ϣ_ID Start With 1;  


--8.������̱���
Create Table ������̱���(
    ID Number(18),  
    ����� Varchar2(20),
    �걾���� Varchar2(64),
    �������� Number(1),
    ����� Varchar2(2048),
    ������ Varchar2(2048),
    ����ͼ�� Varchar2(2048),
    ����ҽʦ Varchar2(64),        
    �������� Date,       
    ��ǰ״̬ Number(1) default 0,
    ��ע Varchar2(1024))
    TABLESPACE zl9BaseItem;  

Alter Table ������̱��� Add Constraint ������̱���_PK Primary Key (ID) Using Index Tablespace zl9indexhis;
Alter Table ������̱��� Add Constraint ������̱���_FK_����� Foreign Key (�����) References ��������Ϣ(�����) on Delete Cascade;
Create Index ������̱���_IX_����� On ������̱���(�����) Pctfree 5 Tablespace zl9indexhis;
Create Sequence ������̱���_ID Start With 1;


    
--9.��������Ϣ  
Create Table ��������Ϣ(
    ����ID Number(18), 
    �������� VARCHAR2(64) Not Null,
    ʹ���˷� Number(5),
    �����˷� Number(5),
    �������� Date,
    ��Ч�� Number(2),
    �������� Date,
    ��¡�� Number(1),
    ���ö��� Varchar2(20),
    ������ Varchar2(10),
    Ӧ����� Varchar2(1024),
    �Ǽ��� Varchar2(64)  Not Null,
    �Ǽ�ʱ�� Date,
    ʹ��״̬ Number(1) default 1,
    ��ע Varchar2(1024))
    TABLESPACE zl9BaseItem;  
        
    
Alter Table ��������Ϣ Add Constraint ��������Ϣ_PK Primary Key (����ID) Using Index Tablespace zl9indexhis;   
Create Sequence ��������Ϣ_����ID Start With 1;     


--10.�����ؼ���Ϣ    
Create Table �����ؼ���Ϣ(
    ID Number(18),    
    ����� Varchar(20) not null,
    �Ŀ�ID Number(18) not null,
    ����ID Number(18),        
    ����ID Number(18),
    �ؼ����� Number(1) default 0,
    �������� Number(1) default 0,
    ��ǰ״̬ NUMBER(1) default 0,
    ���ʱ�� Date,    
    �ؼ�ҽʦ Varchar2(64),
    �嵥״̬ Number(1) default 0,
    ��Ŀ��� Varchar2(20) null)
    TABLESPACE zl9BaseItem; 
    
    
Alter Table �����ؼ���Ϣ Add Constraint �����ؼ���Ϣ_PK Primary Key (ID) Using Index Tablespace zl9indexhis;   
Alter table �����ؼ���Ϣ Add constraint �����ؼ���Ϣ_FK_����� Foreign key(�����) References ��������Ϣ(�����) On Delete Cascade;        
Alter Table �����ؼ���Ϣ Add Constraint �����ؼ���Ϣ_FK_�Ŀ�ID Foreign Key (�Ŀ�ID) References ����ȡ����Ϣ(�Ŀ�ID) On Delete Cascade; 
Alter Table �����ؼ���Ϣ Add Constraint �����ؼ���Ϣ_FK_����ID Foreign Key (����ID) References ����������Ϣ(����ID) On Delete Cascade; 
Alter Table �����ؼ���Ϣ Add Constraint �����ؼ���Ϣ_FK_����ID Foreign Key (����ID) References ��������Ϣ(����ID) On Delete Cascade; 
Create Index �����ؼ���Ϣ_IX_�Ŀ�ID On �����ؼ���Ϣ(�Ŀ�ID, ����ID) Pctfree 5 Tablespace zl9indexhis;          
Create Index �����ؼ���Ϣ_IX_����ID On �����ؼ���Ϣ(����ID,�Ŀ�ID,����ID) Pctfree 5 Tablespace zl9indexhis; 
Create Sequence �����ؼ���Ϣ_ID Start With 1;   
         

--11.�������ӳ�
Create Table �������ӳ�(
    ID Number(18),    
    ����� Varchar2(20),
    �ӳ�ԭ�� Varchar2(1024) not null,        
    �ӳ����� Number(2) not null,
    ��ʱ��� Varchar2(1024),
    ת���� Varchar2(64),
    �Ǽ��� Varchar2(64),
    �Ǽ�ʱ�� Date,    
    ��ǰ״̬ Number(1) default 0)
    TABLESPACE zl9BaseItem; 
    
Alter Table �������ӳ� Add Constraint �������ӳ�_PK Primary Key(ID) Using Index Tablespace zl9indexhis;
Alter Table �������ӳ� Add Constraint �������ӳ�_FK_����� Foreign Key(�����) References ��������Ϣ(�����) On Delete Cascade;
Create Index �������ӳ�_IX_����� On �������ӳ�(�����) Pctfree 5 Tablespace zl9indexhis;
Create Sequence �������ӳ�_ID Start With 1;


--12.���������Ϣ
Create Table ���������Ϣ(
    ID Number(18),    
    ����� Varchar2(20),
    ����ҽʦ Varchar2(64) not null,
    ����ҽʦ Varchar2(64),
    ���ﵥλ Varchar2(64),         
    ����ʱ�� Date not null,
    ��ֹʱ�� Date not null,
    �������� Number(1) default 0,
    ������� Varchar2(2048),
    ��Ͻ�� Varchar2(2048),
    ������ Varchar2(2048),    
    ���ʱ�� Date,
    ��ע Varchar2(1024),
    ��ǰ״̬ Number(1) default 0)
    TABLESPACE zl9BaseItem; 
        
    
Alter Table ���������Ϣ Add Constraint ���������Ϣ_PK Primary Key(ID) Using Index Tablespace zl9indexhis;
Alter Table ���������Ϣ Add Constraint ���������Ϣ_FK_����� Foreign Key(�����) References ��������Ϣ(�����) On Delete Cascade;
Create Index ���������Ϣ_IX_����� On ���������Ϣ(�����) Pctfree 5 Tablespace zl9indexhis;   
Create Sequence ���������Ϣ_ID Start With 1;

      
    
--13.�����巴��
Create Table �����巴��(
    ID Number(18),   
    ����ID Number(18), 
    �ο������ VARCHAR2(200),
    ʵ������ Number(1) default 0,
    �������� VARCHAR2(10),
    ������� VARCHAR2(1024) Not Null,
    ����ҽ�� VARCHAR2(64) Not Null,
    ����ʱ�� Date)
    TABLESPACE zl9BaseItem;   
    
Alter Table �����巴�� Add Constraint �����巴��_PK Primary Key (ID) Using Index Tablespace zl9indexhis;    
Alter Table �����巴�� Add Constraint �����巴��_FK_����ID Foreign Key (����ID) References ��������Ϣ(����ID) On Delete Cascade; 
Create Index �����巴��_IX_����ID On �����巴��(����ID) Pctfree 5 Tablespace zl9indexhis;   
Create Sequence �����巴��_ID Start With 1;  


--14.�����ײ���Ϣ
Create Table �����ײ���Ϣ(
    �ײ�ID Number(18), 
    �ײ����� VARCHAR2(64) not null,
    �ײ�˵�� VARCHAR2(1024),
    ������ VARCHAR2(64) Not Null,
    ����ʱ�� Date)
    TABLESPACE zl9BaseItem;  
    
Alter Table �����ײ���Ϣ Add Constraint �����ײ���Ϣ_PK Primary Key (�ײ�ID) Using Index Tablespace zl9indexhis;
Create Sequence �����ײ���Ϣ_�ײ�ID Start With 1;


--15.�����ײ͹���
 Create Table �����ײ͹���(
    ID Number(18),    
    �ײ�ID Number(18), 
    ����ID Number(18))
    TABLESPACE zl9BaseItem;  

Alter Table �����ײ͹��� Add Constraint �����ײ͹���_PK Primary Key (ID) Using Index Tablespace zl9indexhis;
Alter Table �����ײ͹��� Add Constraint �����ײ͹���_FK_�ײ�ID Foreign Key (�ײ�ID) References �����ײ���Ϣ(�ײ�ID) On Delete Cascade;
Alter Table �����ײ͹��� Add Constraint �����ײ͹���_FK_����ID Foreign Key (����ID) References ��������Ϣ(����ID) On Delete Cascade;
Create Index �����ײ͹���_IX_�ײ�ID On �����ײ͹���(�ײ�ID,����ID) Pctfree 5 Tablespace zl9indexhis;
Create Sequence �����ײ͹���_ID Start With 1;
  
    
--16.����鵵��Ϣ
Create Table ����鵵��Ϣ(
    ����ID Number(18), 
    ����� Varchar2(20),
    ������� Number(1) default 0,
    ���ϱ�� Varchar2(20) Not Null,
    �������� Number(2) Not null,
    �ɽ����� Number(2) Not null,
    ��ʧ���� Number(2),    
    ���ʱ�� Date,
    ����״̬ Number(1)  default 0,
    ����� Varchar2(64) Not Null,
    ���λ�� Varchar2(64),
    ��ע Varchar2(1024))
    TABLESPACE zl9BaseItem; 
    
    
Alter Table ����鵵��Ϣ Add Constraint ����鵵��Ϣ_PK Primary Key (����ID) Using Index Tablespace zl9indexhis; 
Alter Table ����鵵��Ϣ Add Constraint ����鵵��Ϣ_FK_����� Foreign Key (�����) References ��������Ϣ(�����) On Delete Cascade;
Create Index ����鵵��Ϣ_IX_����� On ����鵵��Ϣ(�����) Pctfree 5 Tablespace zl9indexhis; 
Create Index ����鵵��Ϣ_IX_���ϱ�� On ����鵵��Ϣ(���ϱ��) Pctfree 5 Tablespace zl9indexhis; 
Create Sequence ����鵵��Ϣ_����ID Start With 1;            
    
   
--17.���������Ϣ    
Create Table ���������Ϣ(
    ID Number(18), 
    ����ID Number(18),
    ����ʱ�� Date Not Null,
    ������ Varchar2(64) Not Null,
    ֤������ Number(1) default 0,
    ֤������ Varchar2(20),
    ��ϵ�绰 Varchar2(20),
    ��ϵ��ַ Varchar2(128),
    Ѻ�� Number(16, 5),
    �������� Number(2),
    �������� Number(1) Default 0,
    �������� Number(5),
    ����ԭ�� Varchar2(1024),
    �Ǽ��� Varchar2(64) Not Null,
    �黹״̬ Number(1)  default 1,
    �黹���� Date,
    �˻�Ѻ�� Number(16,5),
    ����ҽԺ Varchar2(64),
    ����ҽʦ Varchar2(64),
    ������� Varchar2(2048))
    TABLESPACE zl9BaseItem;       

Alter Table ���������Ϣ Add Constraint ���������Ϣ_PK Primary Key (ID) Using Index Tablespace zl9indexhis;    
Alter Table ���������Ϣ Add Constraint ���������Ϣ_FK_����ID Foreign Key (����ID) References ����鵵��Ϣ(����ID) On Delete Cascade; 
Create Index ���������Ϣ_IX_����ID On ���������Ϣ(����ID) Pctfree 5 Tablespace zl9indexhis; 
Create Sequence ���������Ϣ_ID Start With 1;
  
