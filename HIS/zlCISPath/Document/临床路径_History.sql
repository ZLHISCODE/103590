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
    PCTFREE 5
    PCTUSED 85;
Alter Table �����ٴ�·�� Add Constraint �����ٴ�·��_PK Primary Key (ID) Using Index Pctfree 5;
Create Index �����ٴ�·��_IX_����ID On �����ٴ�·��(����ID,��ҳID) Pctfree 5
/
Create Index �����ٴ�·��_IX_����ID On �����ٴ�·��(����ID) Pctfree 5
/
Create Index �����ٴ�·��_IX_·��ID On �����ٴ�·��(·��ID,�汾��) Pctfree 5
/
Create Index �����ٴ�·��_IX_����ʱ�� On �����ٴ�·��(����ʱ��) Pctfree 5
/


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
    PCTFREE 5
    PCTUSED 85;
Alter Table ����·��ִ�� Add Constraint ����·��ִ��_PK Primary Key (ID) Using Index Pctfree 5;
Alter Table ����·��ִ�� Add Constraint ����·��ִ��_UQ_��Ŀ���� Unique (·����¼ID,�׶�ID,����,��ĿID,��Ŀ����) Using Index Pctfree 5;
Create Index ����·��ִ��_IX_���� On ����·��ִ��(����) Pctfree 5
/
Create Index ����·��ִ��_IX_·����¼ID On ����·��ִ��(·����¼ID) Pctfree 5
/
Create Index ����·��ִ��_IX_�׶�ID On ����·��ִ��(�׶�ID) Pctfree 5
/
Create Index ����·��ִ��_IX_��ĿID On ����·��ִ��(��ĿID) Pctfree 5
/
Create Index ����·��ִ��_IX_ͼ��ID On ����·��ִ��(ͼ��ID) Pctfree 5
/
Create Index ����·��ִ��_IX_�Ǽ�ʱ�� On ����·��ִ��(�Ǽ�ʱ��) Pctfree 5
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
    PCTFREE 5
    PCTUSED 85;
Alter Table ����·������ Add Constraint ����·������_PK Primary Key (·����¼ID,�׶�ID,����) Using Index Pctfree 5;
Create Index ����·������_IX_���� On ����·������(����) Pctfree 5
/
Create Index ����·������_IX_�Ǽ�ʱ�� On ����·������(�Ǽ�ʱ��) Pctfree 5
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
    PCTFREE 5
    PCTUSED 85;
Alter Table ����·��ָ�� Add Constraint ����·��ָ��_UQ_����ָ�� Unique (·����¼ID,�׶�ID,����,����ָ��) Using Index Pctfree 5;
Create Index ����·��ָ��_IX_���� On ����·��ָ��(����) Pctfree 5
/

CREATE TABLE ����·��ҽ��(
		·��ִ��ID NUMBER(18),
    ����ҽ��ID NUMBER(18))
    PCTFREE 5
    PCTUSED 85;
Alter Table ����·��ҽ�� Add Constraint ����·��ҽ��_PK Primary Key (·��ִ��ID,����ҽ��ID) Using Index Pctfree 5;

--��ԭ���Ӳ�����¼�ĸ���
Alter Table ���Ӳ�����¼ Add ·��ִ��ID Number(18);
Create Index ���Ӳ�����¼_IX_·��ִ��ID On ���Ӳ�����¼(·��ִ��ID) Pctfree 5
/