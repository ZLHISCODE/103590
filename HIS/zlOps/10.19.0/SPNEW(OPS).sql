
Define n_System=100;

--11927 by �¸��� 2007-11-08
Alter Table ����������¼ Add ��̨���� Number(1);

--11928 by �¸��� 2007-11-08
Alter Table ����������¼ Add �����̶� VarChar2(10);

--11929 by �¸��� 2007-11-08
Alter Table ����������¼ Add (��Ⱦ���� Number(1),��Ⱦ���� Number(1));

--11930 by �¸��� 2007-11-08
Alter Table ����������¼ Add (������ VarChar2(10),�Ƶ��� VarChar2(10),�������� VarChar2(10));

--12454��2008-01-17 by cfr
Alter Table ����������¼ Add ����ʽid Number(18);
Alter Table ������������ Drop Constraint ������������_CK_��������;
Alter Table ������������ Add Constraint ������������_CK_�������� Check (�������� IN(1,2,3,4));

--11935 by �¸��ݡ�2007-11-09
Alter Table ����ҽ������ Modify �վ�˵�� VarChar2(1000);

Create Index ����ҽ����¼_IX_��ʼִ��ʱ�� On ����ҽ����¼(��ʼִ��ʱ��) Pctfree 5  Tablespace zl9indexcis
/

--12025 2007-12-04 by cfr
Drop Table ������ҩ����;
Create Table ������ҩ����(
    ���� VARCHAR2(2),
    ���� VARCHAR2(20),
    ���� VARCHAR2(20),
    �Ƿ������ Number(1))
    TABLESPACE zl9BaseItem
    PCTFREE 5  PCTUSED 85;
Alter Table ������ҩ���� Add Constraint ������ҩ����_PK Primary Key (����) Using Index Pctfree 5 Tablespace zl9indexhis;
Alter Table ������ҩ���� Add Constraint ������ҩ����_UQ_���� Unique (����) Using Index Pctfree 5 Tablespace zl9indexhis;


Alter Table ����������ҩ Add ����_Bak Varchar2(20);
Alter Table ������ҩ�ο� Add ����_Bak Varchar2(20);

Update ������ҩ�ο� Set ����_Bak=Decode(����,1,'��ǰ��ҩ',2,'������ҩ',3,'������ҩ',����_Bak) Where ����_Bak Is Null;
Update ����������ҩ Set ����_Bak=Decode(����,1,'��ǰ��ҩ',2,'������ҩ',3,'������ҩ',����_Bak) Where ����_Bak Is Null;

Alter Table ������ҩ�ο� Drop Constraint ������ҩ�ο�_CK_����;
Alter Table ������ҩ�ο� Drop Constraint ������ҩ�ο�_PK;
Alter Table ����������ҩ Drop Constraint ����������ҩ_PK;
Alter Table ������ҩ�ο� Drop Column ����;
Alter Table ����������ҩ Drop Column ����;

Alter Table ����������ҩ Add ���� Varchar2(20);
Alter Table ������ҩ�ο� Add ���� Varchar2(20);

Update ������ҩ�ο� Set ����=����_Bak;
Update ����������ҩ Set ����=����_Bak;

Alter Table ������ҩ�ο� Add Constraint ������ҩ�ο�_PK Primary Key (����id,����,ҩ��id) Using Index Pctfree 0 Tablespace zl9indexhis;
Alter Table ����������ҩ Add Constraint ����������ҩ_PK Primary Key (��¼id,����,ҩƷid) Using Index Pctfree 0 Tablespace zl9indexhis;

--12144 2007-12-10 by cfr
Alter Table ������λ Add �Ƿ�Ψһ Number(1);
Alter Table ������λ Add �Ƿ�ҽ�� Number(1);
Alter Table ������λ Add �Ƿ�ʿ Number(1);

--12177 2007-12-18 by cfr
Alter Table ����������Ա Add �ڼ� Number(5);
Alter Table ����������¼ Add ˵�� VarChar2(255);

Alter Table ����������Ա Drop Constraint ����������Ա_UQ_��¼id;
Alter Table ����������Ա Add Constraint ����������Ա_UQ_��¼id Unique (��¼id,�ڼ�,����id,��λ,����,����) Using Index Pctfree 5 Tablespace zl9indexhis;

--12009 0-����;1-����;2-Ӥ�� 2007-12-07 by cfr
Update �����¼��Ŀ Set ���ò���=Decode(��Ŀ���,-1,2,0) Where ���ò��� Is Null;

--11953 ����С��ʧ����������ɣ�*����Ϊ��������2007-12-10 by cfr
Update ���˻������� Set ��¼����='��' Where ��Ŀ���=10 And ��¼����='*';


--12025 2007-12-04 by cfr
Delete From ������ҩ����;
Insert Into ������ҩ����(����,����,����,�Ƿ������)
		Select '1','��ǰ��ҩ','SQYY',0 From Dual
Union All	Select '2','������ҩ','MZYY',1 From Dual
Union All	Select '3','������ҩ','SZYY',0 From Dual
Union All	Select '9','������ҩ','QTYY',0 From Dual;

Delete From zlBaseCode Where ϵͳ=&n_System And ����='������ҩ����';
Insert into zlBaseCode(ϵͳ,����,�̶�,˵��,����) VALUES(&n_System,'������ҩ����',0,'�������õ���ҩƷ������','ҽ������');

--12144 2007-12-10 by cfr
Update ������λ Set �Ƿ�Ψһ=1 Where ����='����ҽ��';
Update ������λ Set �Ƿ�ҽ��=1 Where ���� Like '%ҽ��';
Update ������λ Set �Ƿ�ʿ=1 Where ���� Like '%��ʿ';
Update zlBaseCode Set �̶�=0 Where ϵͳ=&n_System And ����='������λ';

--12177 2007-12-18 by cfr
Update ����������Ա Set �ڼ�=1 Where �ڼ� Is Null;


--12025 2007-12-04 by cfr
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(&n_System,1801,'����',User,'������ҩ����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(&n_System,1804,'����',User,'������ҩ����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(&n_System,1804,'����',User,'������������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(&n_System,1804,'����',User,'������������¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(&n_System,1804,'����',User,'zl_����������Ա_Insert','EXECUTE');


--����ZL1_INSIDE_1804_2/����ҽ����
Insert Into zlReports(ID,���,����,˵��,����,��ֽ,��ӡ��,Ʊ��,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��) Values(zlReports_ID.NextVal,'ZL1_INSIDE_1804_2','����ҽ����','����ҽ����',']~!d"{vo}?$Xzpj U1LJ',15,'Epson LQ-1600K',1,&n_System,1804,'����ҽ����',Sysdate,Sysdate);
Insert Into zlRPTFmts(����ID,���,˵��,ͼ��,W,H,ֽ��,ֽ��,��ֽ̬��) Values(zlReports_ID.CurrVal,1,'����ҽ����',0,11904,16832,9,1,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ1',2,Null,0,Null,0,'����:[������Ϣ.����]',Null,735,2400,1800,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ2',2,Null,0,Null,0,'�Ա�:[������Ϣ.�Ա�]',Null,2760,2400,1080,180,0,0,0,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ3',2,Null,0,Null,0,'����:[������Ϣ.����]',Null,4020,2400,1050,180,0,0,0,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����',2,Null,0,'�����1',12,'����ҽ����¼��',Null,4152,1530,3495,495,0,1,1,'����',24,1,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��λ����',2,Null,0,'�����1',12,'[��λ����]',Null,5105,1110,1590,315,0,0,1,'����',16,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ4',2,Null,0,Null,0,'����:[������Ϣ.����]',Null,5385,2400,1800,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ5',2,Null,0,Null,0,'����:[������Ϣ.����]',Null,7395,2400,1140,180,0,2,0,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ6',2,Null,0,Null,0,'סԺ��:[������Ϣ.סԺ��]',Null,8940,2400,2160,180,0,2,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'�����1',4,Null,0,Null,0,'����ҽ��',Null,720,3015,10360,12674,420,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,4210816,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-1,0,Null,Null,'[����ҽ��.��������]','4^345^�´�ҽ��|4^345^����',0,0,525,0,255,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-2,1,Null,Null,'[����ҽ��.����ʱ��]','4^345^�´�ҽ��|4^345^ʱ��',0,0,660,0,255,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-3,2,Null,Null,'[����ҽ��.����ҽ��]','4^345^�´�ҽ��|4^345^ҽ��',0,0,870,0,255,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-4,3,Null,Null,'[����ҽ��.ҽ������]','4^345^��  ��  ҽ  ��|4^345^��  ��',0,0,4230,0,255,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-5,4,Null,Null,'[����ҽ��.�÷�]','4^345^��  ��  ҽ  ��|4^345^��  ��',0,0,1785,0,255,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-6,5,Null,Null,'[����ҽ��.У������]','4^345^ִ��ҽ��|4^345^����',0,0,615,0,255,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-7,6,Null,Null,'[����ҽ��.У��ʱ��]','4^345^ִ��ҽ��|4^345^ʱ��',0,0,645,0,255,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-8,7,Null,Null,'[����ҽ��.У�Ի�ʿ]','4^345^ִ��ҽ��|4^345^��ʿ',0,0,930,0,255,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ7',2,Null,0,'�����1',11,'����:[������Ϣ.ҽ������]',Null,720,2715,2160,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ8',2,Null,0,Null,0,'ִ�п���:[������Ϣ.ִ�п���]',Null,5020,2730,2520,180,0,0,1,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����,˵��) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'������Ϣ','����,202|�Ա�,202|����,202|����,202|����,202|סԺ��,131|ҽ������,202|ִ�п���,202',User||'.������Ϣ,'||User||'.������ҳ,'||User||'.����ҽ����¼,'||User||'.����ҽ������,'||User||'.���ű�,'||User||'.������ĿĿ¼',0,Null);
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,1,'Select P.����,P.�Ա�,P.����,D.���� As ����,P.��Ժ���� As ����,P.סԺ��,A.���� As ҽ������,B.���� As ִ�п���');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,2,'From (Select I.����,I.�Ա�,I.����,I.סԺ��,P.��Ժ����,P.��Ժ����id,V.������Ŀid,L.ִ�в���id');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,3,'      From ������Ϣ I,������ҳ P,����ҽ����¼ V,����ҽ������ L');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,4,'      Where I.����id=V.����id And P.����id=V.����id And P.��ҳID=V.��ҳid And V.ID=[0] And L.ҽ��id=V.ID) P,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,5,'      ���ű� D,������ĿĿ¼ A,���ű� B');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,6,'Where P.��Ժ����id=D.Id And A.ID=P.������Ŀid And B.ID=P.ִ�в���id');
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,Null,0,'ҽ��ID',1,'1',0,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����,˵��) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'����ҽ��','����,139|ID,131|��������,202|����ʱ��,202|����ҽ��,202|У������,202|У��ʱ��,202|У�Ի�ʿ,202|ҽ������,202|�÷�,202',User||'.����ҽ����¼,'||User||'.������ĿĿ¼,'||User||'.�շ���ĿĿ¼',0,Null);
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,1,'--ע������Դ������ȡ��ID�ֶ���Ҫ���������ڵ��ó����¼�Ѵ�ӡҽ��');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,2,'--��ЩID���ǿɼ�ҽ���е�ID(�������г�ҩ�⣬������Ϊ"���ID=NULL"��ҽ��ID)');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,3,Null);
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,4,'--"ҽ����ӡ��¼"�е������ɵ��ó�����ʱ����,����������,�ش�,�״�ֹͣʱ��');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,5,'----������ǰ����(Ӥ��)��Ч�Ĵ�ӡҽ��,������ҽ��,����δУ�Ժ����δ�ӡ��ҽ��.');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,6,Null);
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,7,'--Union ALLǰ�棺�������г�ҩƷҽ����������ҩ�䷽(���䷽�÷���Ϊ׼)��һ���ɼ��걾�ļ���(�Բɼ���ʽ��Ϊ׼)��Ƥ�Ե���������ҽ������������¼����ı�ҽ����');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,8,'--Union ALL���棺��ҩ���г�ҩҽ��');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,9,Null);
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,10,'Select 1 As ����, ID, Substr(����ʱ��, 1, 5) As ��������, Substr(����ʱ��, 7) As ����ʱ��, ����ҽ��, Substr(У��ʱ��, 1, 5) As У������,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,11,'       Substr(У��ʱ��, 7) As У��ʱ��, У�Ի�ʿ, ҽ������, �÷�');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,12,'From (Select L.ID, To_Char(L.����ʱ��, ''DD/MM HH24:MI'') As ����ʱ��, L.����ҽ��, To_Char(L.У��ʱ��, ''DD/MM HH24:MI'') As У��ʱ��,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,13,'              L.У�Ի�ʿ,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,14,'              L.ҽ������ ||');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,15,'               Decode(I.��� || I.��������, ''E4'', '''',');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,16,'                      ''  '' ||');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,17,'                       Decode(L.ִ��Ƶ��, ''һ����'', '''', ''������'', '''', ''��Ҫʱ'', ''��Ҫʱ'', ''����ʱ'', ''����ʱ'',');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,18,'                              Decode(L.������Ŀid, Null, Null,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,19,'                                      ''ÿ��'' || L.�������� || I.���㵥λ || '','' || L.ִ��Ƶ�� || '',��'' || L.�ܸ����� || I.���㵥λ))) || L.Ƥ�Խ�� As ҽ������,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,20,'              '''' As �÷�');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,21,'       From ����ҽ����¼ L, ������ĿĿ¼ I');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,22,'       Where L.ǰ��id = [0] And L.������Ŀid = I.ID(+) And');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,23,'             (L.������� Not In (''5'', ''6'', ''7'', ''E'') Or L.������� = ''E'' And I.�������� Not In (''2'', ''3'') Or I.ID Is Null) And');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,24,'             L.���id Is Null');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,25,'       Union All');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,26,'       Select M.ID, Decode(M.���, U.��ʼҩƷ���, To_Char(M.����ʱ��, ''DD/MM HH24:MI''), '''') As ����ʱ��,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,27,'              Decode(M.���, U.��ʼҩƷ���, M.����ҽ��, '''') As ����ҽ��, To_Char(M.У��ʱ��, ''DD/MM HH24:MI'') As У��ʱ��, M.У�Ի�ʿ, M.ҽ������,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,28,'              Decode(M.���, U.��ʼҩƷ���, U.��ҩ, Decode(M.���, U.����ҩƷ���, ''��'', ''��'')) As ��ҩ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,29,'       From (Select L.ID, L.���id, L.���, L.����ʱ��, L.����ҽ��, L.У�Ի�ʿ, L.У��ʱ��,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,30,'                     L.ҽ������ || ''  ÿ��'' || �������� || I.���㵥λ || '',��'' || �ܸ����� || E.���㵥λ As ҽ������');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,31,'              From ����ҽ����¼ L, ������ĿĿ¼ I, �շ���ĿĿ¼ E');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,32,'              Where L.ǰ��id = [0] And L.������Ŀid = I.ID And L.�շ�ϸĿid = E.ID And L.������� In (''5'', ''6'')) M,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,33,'            (Select U.ID, U.ִ��Ƶ�� || '','' || U.���� As ��ҩ, Min(M.���) As ��ʼҩƷ���, Max(M.���) As ����ҩƷ���');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,34,'              From (Select L.ID, L.ִ��Ƶ��, I.����');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,35,'                     From ����ҽ����¼ L, ������ĿĿ¼ I');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,36,'                     Where L.ǰ��id = [0] And L.������Ŀid = I.ID And I.��� = ''E'' And I.�������� = ''2'') U,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,37,'                   (Select L.���, L.���id');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,38,'                     From ����ҽ����¼ L');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,39,'                     Where L.ǰ��id = [0] And L.������� In (''5'', ''6'')) M');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,40,'              Where U.ID = M.���id');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,41,'              Group By U.ID, U.ִ��Ƶ��, U.����) U');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,42,'       Where M.���id = U.ID)');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,43,'Order By ����');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,44,Null);
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,Null,0,'ҽ��ID',1,'1',0,Null,Null,Null,Null,Null,Null);

--����ZL1_INSIDE_1804_2/����ҽ����
Insert into zlProgFuncs(ϵͳ,���,����,˵��) Values(&n_System,1804,'����ҽ����',Null);
Insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(&n_System,1804,'����ҽ����',User,'������ҳ','SELECT');
Insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(&n_System,1804,'����ҽ����',User,'������Ϣ','SELECT');
Insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(&n_System,1804,'����ҽ����',User,'����ҽ������','SELECT');
Insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(&n_System,1804,'����ҽ����',User,'����ҽ����¼','SELECT');
Insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(&n_System,1804,'����ҽ����',User,'���ű�','SELECT');
Insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(&n_System,1804,'����ҽ����',User,'�շ���ĿĿ¼','SELECT');
Insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(&n_System,1804,'����ҽ����',User,'������ĿĿ¼','SELECT');

--12176 2007-12-17 by cfr
--12177 2007-12-18 by cfr
CREATE OR REPLACE PROCEDURE zl_����������Ա_Insert(
	��¼id_IN	IN   ����������Ա.��¼id%TYPE,
	��λ_IN	IN   ����������Ա.��λ%TYPE,
	��Աid_IN	IN   ����������Ա.��Աid%TYPE,
	����_IN	IN   ����������Ա.����%TYPE,
	�ڼ�_In	In    ����������Ա.�ڼ�%TYPE:=1
)
IS
BEGIN
	INSERT INTO ����������Ա(��¼id,��λ,��Աid,����,�ڼ�)
	VALUES (��¼id_IN,��λ_IN,Decode(��Աid_IN,0,Null,��Աid_IN),����_IN,�ڼ�_In);

EXCEPTION
	WHEN OTHERS THEN
		Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_����������Ա_Insert;
/

--11928 by cfr 2007-11-08
--12012 by cfr 2007-11-19
--12177 2007-12-18 by cfr
--12454��2008-01-17 by cfr
Create Or Replace Procedure Zl_����������¼_Aduit
(
  Id_In       In ����������¼.ID%Type,
  ҽ��id_In   In ����ҽ����¼.ID%Type,
  ����ʽ_In In ����������¼.����ʽ%Type,
  ��������_In In ����������¼.��������%Type,
  ������ģ_In In ����������¼.������ģ%Type,
  ����ʽid_In In ����������¼.����ʽid%Type
) Is
Begin
  --��д����������¼(��������ϵͳ������¼)
  -------------------------------------------------------------------------------------------------------------------
  Update ����������¼
  Set ҽ��id = ҽ��id_In, ����ʽ = ����ʽ_In, �������� = ��������_In, ������ģ = ������ģ_In, ����״̬ = 1, ����ʽid=Decode(����ʽid_In,0,Null,����ʽid_In)
  Where ID = Id_In;
  If Sql%Rowcount = 0 Then
  
    --��д����������¼(��������ϵͳ������¼)
    -------------------------------------------------------------------------------------------------------------------
    Insert Into ����������¼
      (ID, ҽ��id, ����id, ��ҳid, ����״̬, ����ʽ, ��������, ������ģ, �����̶�, ����ʽid)
      Select Id_In, ҽ��id_In, A.����id, A.��ҳid, 1, ����ʽ_In, ��������_In, ������ģ_In,
             Decode(A.������־, 1, '��', ''), Decode(����ʽid_In,0,Null,����ʽid_In)
      From ����ҽ����¼ A
      Where A.ID = ҽ��id_In;
  End If;

  --���ʱ��дȱʡ��������λ��Ա
  -------------------------------------------------------------------------------------------------------------------
  Delete From ����������Ա Where ��¼id = Id_In;
  For r_List In (Select ����, �Ƿ�ҽ��, �Ƿ�Ψһ From ������λ) Loop
    If r_List.�Ƿ�ҽ�� = 1 And r_List.�Ƿ�Ψһ = 1 Then
      Insert Into ����������Ա
        (��¼id, ��λ, ����, �ڼ�)
        Select Id_In, r_List.����, ����ҽ��, 1 From ����ҽ����¼ Where ID = ҽ��id_In;
    Else
      Insert Into ����������Ա
        (��¼id, ��λ, ����, �ڼ�)
        Select Id_In, r_List.����, Trim(Substrb(Nvl(����, ' '), 1, 20)), 1
        From ����ҽ������
        Where ҽ��id = ҽ��id_In And ��Ŀ = r_List.����;
    End If;
  End Loop;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����������¼_Aduit;
/

--12177 2007-12-18 by cfr
Create Or Replace Procedure Zl_����������Ա_Delete
(
	��¼id_In In ����������Ա.��¼id%Type,
	�ڼ�_In   In ����������Ա.�ڼ�%Type := 0
) Is
Begin
	If �ڼ�_In = 0 Then
		Delete From ����������Ա Where ��¼id = ��¼id_In;
	Else
		Delete From ����������Ա Where ��¼id = ��¼id_In And �ڼ� = �ڼ�_In;
	End If;
Exception
	When Others Then
		Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����������Ա_Delete;
/

--12012 by cfr 2007-11-19
CREATE OR REPLACE PROCEDURE zl_����������¼_AduitCancel(
	ID_In				In		����������¼.ID%Type
)
IS	
	v_Error		varchar2(250);
	Err_custom	Exception;
BEGIN	
	Delete From  ����������� Where ��¼ID=ID_In And ����=2;
	Update ����������¼ Set ����״̬=Null Where ID=ID_In;

	zl_����������Ա_Delete(ID_In);
EXCEPTION
	When Err_custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_����������¼_AduitCancel;
/

--11927 by cfr 2007-11-08
--11928 by cfr 2007-11-08
--11929 by cfr 2007-11-08
--11930 by cfr 2007-11-08
--11935 by cfr 2007-11-08
--12177 2007-12-18 by cfr
--12454��2008-01-17 by cfr
Create Or Replace Procedure Zl_����������¼_Arrange
(
  Id_In           In ����������¼.ID%Type,
  ������ʼʱ��_In In ����������¼.������ʼʱ��%Type,
  ��������ʱ��_In In ����������¼.��������ʱ��%Type,
  ������_In       In ����������¼.������%Type,
  ������id_In     In ����������¼.������id%Type := Null,
  ������Ա_In     In Varchar2 := Null,
  ��¼����_In     In Number := 2,
  �����̶�_In     In ����������¼.�����̶�%Type := Null,
  ��̨����_In     In ����������¼.��̨����%Type := 0,
  �޾�����_In     In ����������¼.�޾�����%Type := 0,
  ��Ⱦ����_In     In ����������¼.��Ⱦ����%Type := 0,
  ��Ⱦ����_In     In ����������¼.��Ⱦ����%Type := 0
) Is
  v_Tmp    Varchar2(4000);
  v_Tmprow Varchar2(4000);
  v_Svrtmp Varchar2(50);
  n_Pos    Number(18);
  v_��λ   ����������Ա.��λ%Type;
  n_��Աid ����������Ա.��Աid%Type;
  v_����   ����������Ա.����%Type;
  v_����   ����������Ա.����%Type;

  n_��¼��� ����ҽ������.��¼���%Type;
  v_No       ����ҽ������.NO%Type;
  v_����no   ����ҽ������.NO%Type;
  n_���ͺ�   ����ҽ������.���ͺ�%Type;
  n_�Ƽ����� ����ҽ����¼.�Ƽ�����%Type;
  v_Error    Varchar2(255);
  Err_Custom Exception;
  n_Flag Number(1);
Begin
  --���
  ---------------------------------------------------------------------------------------------------------------
  n_Flag := 0;
  Begin
    Select 1 Into n_Flag From ����ҽ����¼ A, ����������¼ B Where A.ID = B.ҽ��id And B.ID = Id_In And ҽ��״̬ <> 4;
  Exception
    When Others Then
      n_Flag := 0;
  End;
  If n_Flag = 0 Then
    v_Error := '��������ҽ����¼�Ѿ������ڻ�ɾ����';
    Raise Err_Custom;
  End If;

  ---------------------------------------------------------------------------------------------------------------
  n_Flag := 0;
  Begin
    Select 1
    Into n_Flag
    From ����ҽ������ A, ����������¼ B
    Where A.ִ��״̬ > 0 And A.ҽ��id = B.ҽ��id And B.ID = Id_In;
  Exception
    When Others Then
      n_Flag := 0;
  End;
  If n_Flag = 1 Then
    v_Error := '����ҽ���Ѿ����Ͳ�������ִ�л��Ѿ�ִ����ɣ�';
    Raise Err_Custom;
  End If;

  --����ʱ��,�ص���д
  ---------------------------------------------------------------------------------------------------------------
  Update ����������¼
  Set ������ʼʱ�� = ������ʼʱ��_In, ��������ʱ�� = ��������ʱ��_In, �������� = Trunc(������ʼʱ��_In),
      ������ = ������_In, ������id = ������id_In, ����״̬ = 2, �����̶� = �����̶�_In, ��̨���� = ��̨����_In,
      �޾����� = �޾�����_In, ��Ⱦ���� = ��Ⱦ����_In, ��Ⱦ���� = ��Ⱦ����_In
  Where ID = Id_In And ����״̬ = 1;
  If Sql%Rowcount = 0 Then
    v_Error := '��ǰ�����Ѿ�ȡ����ˣ����ܼ������Ų�����';
    Raise Err_Custom;
  End If;

  --�޸�ҽ���Ŀ�ʼִ��ʱ��
  ---------------------------------------------------------------------------------------------------------------
  For r_Order In (Select ҽ��id From ����������¼ Where ID = Id_In) Loop
    Update ����ҽ����¼ Set ��ʼִ��ʱ�� = ������ʼʱ��_In Where r_Order.ҽ��id In (ID, ���id);
  End Loop;

  --������Ա����д
  ---------------------------------------------------------------------------------------------------------------
  Delete From ����������Ա Where ��¼id = Id_In;
  v_Tmp := ������Ա_In || ';';
  While v_Tmp Is Not Null Loop
    n_Pos := Instr(v_Tmp, ';');
    If n_Pos > 0 Then
      v_Tmprow := Substr(v_Tmp, 1, n_Pos - 1);
      v_Tmp    := Substr(v_Tmp, n_Pos + 1);
      n_Pos    := Instr(v_Tmprow, ',');
      If n_Pos > 0 Then
        n_��Աid := To_Number(Substr(v_Tmprow, 1, n_Pos - 1));
        v_Tmprow := Substr(v_Tmprow, n_Pos + 1);
        n_Pos    := Instr(v_Tmprow, ',');
        If n_Pos > 0 Then
          v_��λ   := Substr(v_Tmprow, 1, n_Pos - 1);
          v_Tmprow := Substr(v_Tmprow, n_Pos + 1);
          n_Pos    := Instr(v_Tmprow, ',');
          If n_Pos > 0 Then
            v_���� := Substr(v_Tmprow, 1, n_Pos - 1);
            v_���� := Substr(v_Tmprow, n_Pos + 1, 1);
            If v_��λ Is Not Null And v_���� Is Not Null Then
              Zl_����������Ա_Insert(Id_In, v_��λ, n_��Աid, v_����, 1);
            End If;
          End If;
        End If;
      End If;
    End If;
  End Loop;

  --���ҽ��δ����,�����ҽ������
  ---------------------------------------------------------------------------------------------------------------
  n_��¼��� := 0;

  For r_Order In (Select A.ID, A.���id, ִ�п���id, Decode(A.�Ƽ�����, 0, 1, 1, -1, 2, 0) As �Ƽ�����, A.�������
                  From ����ҽ����¼ A, ����������¼ B
                  Where A.ҽ��״̬ Not In (4, 8) And B.ҽ��id In (A.ID, A.���id) And B.ID = Id_In
                  Order By Decode(A.���id, Null, 0, 1)) Loop
  
    n_��¼��� := n_��¼��� + 1;
    If n_��¼��� = 1 Then
      Select Nextno(10), Nextno(Decode(��¼����_In, 1, 13, 14)) Into n_���ͺ�, v_No From Dual;
      n_�Ƽ����� := r_Order.�Ƽ�����;
    End If;
  
    If v_����no Is Null And r_Order.������� = 'G' Then
      Select Nextno(Decode(��¼����_In, 1, 13, 14)) Into v_����no From Dual;
    End If;
  
    If r_Order.���id Is Null Then
      Zl_����ҽ������_Insert(r_Order.ID, n_���ͺ�, ��¼����_In, v_No, n_��¼���, 1, Null, Null,
                             Sysdate + 1 / 24 / 60 / 60, 0, r_Order.ִ�п���id, n_�Ƽ�����, 1);
    Else
      If r_Order.������� = 'G' Then
        Zl_����ҽ������_Insert(r_Order.ID, n_���ͺ�, ��¼����_In, v_����no, n_��¼���, 1, Null, Null,
                               Sysdate + 1 / 24 / 60 / 60, 0, r_Order.ִ�п���id, r_Order.�Ƽ�����, 0);
      Else
        Zl_����ҽ������_Insert(r_Order.ID, n_���ͺ�, ��¼����_In, v_No, n_��¼���, 1, Null, Null,
                               Sysdate + 1 / 24 / 60 / 60, 0, r_Order.ִ�п���id, n_�Ƽ�����, 0);
      End If;
    End If;
  End Loop;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����������¼_Arrange;
/

--12454��2008-01-17���£١��ãƣ�
Create Or Replace Procedure Zl_�����������_Delete
(
  ��¼id_In In ����������¼.ID%Type,
  ����_In   In �����������.����%Type := 0
) Is
Begin
  If ����_In = 0 Then
    Delete From ����������� Where ��¼id = ��¼id_In;
  Else
    Delete From ����������� Where ��¼id = ��¼id_In And ���� = ����_In;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�����������_Delete;
/
--12454��2008-01-17���£١��ãƣ�
Create Or Replace Procedure Zl_����������¼_Updateadvice(Id_In In ����������¼.ID%Type) Is

  Cursor c_Opsrecords Is
    Select * From ����������¼ Where ID = Id_In;
  r_Opsrecord c_Opsrecords%Rowtype;

  v_Tmp    Varchar2(4000);
  v_Tmprow Varchar2(4000);
  v_Svrtmp Varchar2(50);

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin

  Open c_Opsrecords;
  Fetch c_Opsrecords
    Into r_Opsrecord;
  If c_Opsrecords%Rowcount = 0 Then
    Close c_Opsrecords;
    v_Error := '��������ҽ����¼�Ѿ������ڻ�ɾ����';
    Raise Err_Custom;
  End If;

  --��д����������Ϣ��ҽ��վ����ʿվ���Ա���ʾ����������Ϣ��
  ---------------------------------------------------------------------------------------------------------------
  v_Tmp    := ' ';
  v_Tmprow := ' ';
  v_Svrtmp := ' ';
  For r_List In (Select A.��λ, A.����
                 From ����������Ա A, ������λ B
                 Where A.��λ = B.���� And A.��¼id = Id_In And A.�ڼ� = 1
                 Order By B.����) Loop
    If v_Svrtmp <> r_List.��λ Then
    
      If Trim(v_Svrtmp) Is Not Null Then
        v_Tmp    := Trim(v_Tmp || v_Tmprow || Chr(13) || Chr(10));
        v_Tmprow := ' ';
      End If;
      v_Svrtmp := r_List.��λ;
      v_Tmprow := Trim(v_Svrtmp || '��');
      v_Tmprow := v_Tmprow || r_List.����;
    Else
      v_Tmprow := v_Tmprow || ',' || r_List.����;
    End If;
  End Loop;
  v_Tmp := Trim(v_Tmp || v_Tmprow || Chr(13) || Chr(10));

  --���������¼(��������)
  ---------------------------------------------------------------------------------------------------------------
  v_Tmprow := ' ';
  For r_List In (Select A.��������, Rownum As ���
                 From ����������� A
                 Where A.��¼id = Id_In And A.���� = 1
                 Order By Decode(A.ȱʡ, 1, 0, 1)) Loop
    If r_List.��� = 1 Then
      v_Tmprow := '����������' || r_List.��������;
    Else
      v_Tmprow := v_Tmprow || ',' || r_List.��������;
    End If;
  End Loop;
  v_Tmp := Trim(v_Tmp || v_Tmprow || Chr(13) || Chr(10));

  --���������¼(��������)
  ---------------------------------------------------------------------------------------------------------------
  v_Tmprow := ' ';
  For r_List In (Select A.��������, Rownum As ���
                 From ����������� A
                 Where A.��¼id = Id_In And A.���� = 2
                 Order By Decode(A.ȱʡ, 1, 0, 1)) Loop
    If r_List.��� = 1 Then
      v_Tmprow := '����������' || r_List.��������;
    Else
      v_Tmprow := v_Tmprow || ',' || r_List.��������;
    End If;
  End Loop;
  v_Tmp := Trim(v_Tmp || v_Tmprow || Chr(13) || Chr(10));

  --��������˵��
  ---------------------------------------------------------------------------------------------------------------
  Update ����ҽ������
  Set ����ʱ�� = r_Opsrecord.������ʼʱ��, ִ�м� = r_Opsrecord.������, �վ�˵�� = v_Tmp
  Where ҽ��id In (Select A.ID From ����ҽ����¼ A, ����������¼ B Where B.ID = Id_In And B.ҽ��id In (A.ID, A.���id));

  Close c_Opsrecords;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����������¼_Updateadvice;
/

--12176 ȡ������ʱ��ɾ����Ա������� 2007-12-17 by cfr
CREATE OR REPLACE PROCEDURE zl_����������¼_ArrangeCancel(
	ID_In			IN	����������¼.ID%TYPE
)
IS
	v_Error varchar2(255);
	Err_custom    Exception;
BEGIN
	Update ����������¼ Set ������ʼʱ��=Null,
					��������ʱ��=Null,
					��������=Null,
					������=Null,
					������id=Null,
					����״̬=1
	Where ID=ID_In;		
EXCEPTION
	When Err_custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_����������¼_ArrangeCancel;
/

--11927 by �¸��� 2007-11-08
--11928 by �¸��� 2007-11-08
--11929 by �¸��� 2007-11-08
--11930 by �¸��� 2007-11-08
--12177 2007-12-18 by cfr
Create Or Replace Procedure Zl_����������¼_Update
(
	��¼id_In       In ����������¼.Id%Type,
	��������_In     In ����������¼.��������%Type,
	������ʼʱ��_In In ����������¼.������ʼʱ��%Type,
	��������ʱ��_In In ����������¼.��������ʱ��%Type,
	����ʼʱ��_In In ����������¼.����ʼʱ��%Type,
	�������ʱ��_In In ����������¼.�������ʱ��%Type,
	����ʽ_In     In ����������¼.����ʽ%Type,
	����ʽid_In     In ����������¼.����ʽid%Type,
	��������_In     In ����������¼.��������%Type,
	��������_In     In ����������¼.��������%Type,
	��Һ����_In     In ����������¼.��Һ����%Type,
	������ʼʱ��_In In ����������¼.������ʼʱ��%Type,
	��������ʱ��_In In ����������¼.��������ʱ��%Type,
	������_In       In ����������¼.������%Type,
	������id_In     In ����������¼.������id%Type,
	������ģ_In     In ����������¼.������ģ%Type,
	�����̶�_In     In ����������¼.�����̶�%Type := Null,
	������_In       In ����������¼.������%Type := Null,
	�Ƶ���_In       In ����������¼.�Ƶ���%Type := Null,
	��������_In     In ����������¼.��������%Type := Null,
	��̨����_In     In ����������¼.��̨����%Type := 0,
	�޾�����_In     In ����������¼.�޾�����%Type := 0,
	��Ⱦ����_In     In ����������¼.��Ⱦ����%Type := 0,
	��Ⱦ����_In     In ����������¼.��Ⱦ����%Type := 0,
	˵��_In     In ����������¼.˵��%Type := Null
) Is
Begin
	Update ����������¼
	Set �������� = ��������_In, ������ʼʱ�� = ������ʼʱ��_In, ��������ʱ�� = ��������ʱ��_In,
			����ʼʱ�� = ����ʼʱ��_In, �������ʱ�� = �������ʱ��_In, ����ʽ = ����ʽ_In, �������� = ��������_In,
			�������� = ��������_In, ��Һ���� = ��Һ����_In, ������ʼʱ�� = ������ʼʱ��_In, ��������ʱ�� = ��������ʱ��_In,
			������ = ������_In, ������id = ������id_In, ������ģ = ������ģ_In, ������ = ������_In, �Ƶ��� = �Ƶ���_In,
			�������� = ��������_In, �޾����� = �޾�����_In, ��Ⱦ���� = ��Ⱦ����_In, ��Ⱦ���� = ��Ⱦ����_In,�����̶� = �����̶�_In,
			��̨���� = ��̨����_In,˵��=˵��_In, ����ʽid=Decode(����ʽid_In,0,Null,����ʽid_In)
	Where Id = ��¼id_In;
Exception
	When Others Then
		Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����������¼_Update;
/

CREATE OR REPLACE PROCEDURE ZL_�����������_INSERT(
	��¼ID_IN IN ����������¼.ID%TYPE,
	����_IN IN �����������.����%TYPE,
	ȱʡ_IN IN �����������.ȱʡ%TYPE,
	��������_IN IN �����������.��������%TYPE,
	��������ID_IN IN �����������.��������ID%TYPE,
	������ĿID_IN IN �����������.������ĿID%TYPE
)
IS
BEGIN	
	Insert Into �����������
		(��¼ID,����,ȱʡ,��������,��������ID,������ĿID)
		VALUES
		(��¼ID_IN,����_IN,ȱʡ_IN,��������_IN,Decode(��������ID_IN,0,Null,��������ID_IN),Decode(������ĿID_IN,0,Null,������ĿID_IN));
EXCEPTION
	WHEN OTHERS THEN
		Zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_�����������_INSERT;
/
