--��������
Insert Into zlParameters(ID,ϵͳ,ģ��,˽��,����,��Ȩ,�̶�,������,������,����ֵ,ȱʡֵ,����˵��)
Select Rownum+B.ID,A.* From (
  Select ϵͳ,ģ��,˽��,����,��Ȩ,�̶�,������,������,����ֵ,ȱʡֵ,����˵�� From zlParameters Where ID=0 Union All
  Select 100,1294,1,0,0,0,40,'�������', '1','1','���˲��������Ϊ����ļ��.'   From Dual   Union ALL
  Select 100,1294,1,0,0,0,41,'��������', '1','1','���˲��������Ϊ�����ļ��.'   From Dual   Union ALL
  Select 100,1294,1,0,0,0,42,'ϸ������', '1','1','���˲��������Ϊϸ���ļ��.'   From Dual   Union ALL
  Select 100,1294,1,0,0,0,43,'�������', '1','1','���˲��������Ϊ����ļ��.'   From Dual   Union ALL
  Select 100,1294,1,0,0,0,44,'ʬ�����', '1','1','���˲��������Ϊʬ��ļ��.'   From Dual   Union ALL
  Select 100,1294,1,0,0,0,45,'���ι���', '1','1','���˲���걾����Ϊ���εļ��.'   From Dual   Union ALL
  Select 100,1294,1,0,0,0,46,'С�걾����', '1','1','���˲���걾����ΪС�걾�ļ��.'   From Dual   Union ALL
  Select 100,1294,1,0,0,0,47,'���̹���', '1','1','���˲���걾����Ϊ���̵ļ��.'   From Dual   Union ALL
  Select 100,1294,1,0,0,0,48,'�������', '1','1','���˲���걾����Ϊ����ļ��.'   From Dual   Union ALL
  Select 100,1294,1,0,0,0,49,'Һ������', '1','1','���˲���걾����ΪҺ���ļ��.'   From Dual   Union ALL    
  Select 100,1294,1,0,0,0,50,'����ҳ��', '0','0','���õ�ǰ���������ݵĹ���ҳ������.'   From Dual
  ) A,(Select Nvl(Max(ID),0) AS ID From zlParameters) B;




--��Ӳ���������
Insert Into zlPrograms(���,����,˵��,ϵͳ,����) Values(1294,'Ӱ������վ','���ڲ���걾���պ�ȡ�ġ�ͼ��ɼ���������д������ҽ���ͷ��õĵǼ�',100,'zl9PacsWork');


--����ģ��
--Ӱ������վ 1294
Insert Into zlProgFuncs(ϵͳ,���,����) Values(100,1294,'����');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'���п���',1,'���Բ鿴���п���PACS����Ȩ�ޡ�');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'���Ǽ�',2,'��д���Ǽ�,ȡ���ǼǺ��޸ĵǼǡ�');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'�ޱ������',3,'ֱ����ɺͻ����ޱ���ļ�顣');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'�ļ�����',4,'�ɷ���ָ��ͼ���ļ�������Ŀ¼��Ȩ�ޡ�');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'ȡ������',5,'ȡ��Ӱ���鱨��״̬��');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'���ͼ��',6,'ɾ��ͼ��ȡ�����������¹����Լ�Q/R��ȡͼ���Ȩ�ޡ�');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'������',7,'ȷ�ϱ��μ����úͱ��涼�Ѿ�¼����ɡ�');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'ȡ��������',8,'ȡ�����μ����ɵ�״̬��');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'��Ƶ�ɼ�',9,'�ɽ�����ƵӰ��ɼ���Ȩ�ޡ�');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'�洢����',10,'�ɽ������߽�������ͼ�����ݵĹ���');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'��������',11,'���в����趨');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'�ɼ���������',12,'���в����趨');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'Ӱ���ʿ�',13,'����Ӱ�������ȼ�����');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'PACS������д',15,'ʹ��PACS����༭����д����');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'PACS�����޶�',16,'ʹ��PACS����༭���޶�����');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'PACS�����ӡ',17,'ʹ��PACS����༭����ӡ����');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'PACS����ɾ��',18,'ʹ��PACS����༭������ǿ��ɾ������');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'��鱨��',19,'ȷ�ϱ���');
Insert Into zlprogfuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'��ɫͨ��',20,'��ĳ�μ����/ȡ����ɫͨ��');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'PACS���˱���',21,'ʹ��PACS����༭��,�������д�������Աɾ��������д�ı���');
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'�Ŷӽк�',22,'�Ա����Ļ��߽����Ŷ���ʾ����������');
Insert Into zlprogfuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'���',23,'��¼���˵������Ϣ');
Insert Into zlprogfuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'δ�ɷѱ���',24,'ӵ�и�Ȩ�޿��Ա���δ�ɷѵļ���¼');
Insert Into zlprogfuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'��������',25,'ӵ�и�Ȩ�޿������ù�������');
Insert Into zlprogfuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'PACS�������Ʊ���',26,'ӵ�и�Ȩ�޿�����PACS����༭���У�ͨ����ʷ���湦�ܲ鿴�������ҵı���');


Insert Into zlprogfuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'�걾����',28,'���ͼ�걾���к���');
Insert Into zlprogfuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'����ȡ��',29,'��ȡ�����ĲĿ�');
Insert Into zlprogfuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'������Ƭ',30,'�����������Ƭ������ϸ����������ʯ��');
Insert Into zlprogfuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'�����ӳ�',31,'����Ҫ�ӳٵĲ�������еǼǹ���');
Insert Into zlprogfuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'��������',32,'�༭�������̱���');
Insert Into zlprogfuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'���߱���',33,'�༭���߹��̱���');
Insert Into zlprogfuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'���ӱ���',34,'�༭���ӹ��̱���');
Insert Into zlprogfuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'��Ⱦ����',35,'�༭��Ⱦ���̱���');
Insert Into zlprogfuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'�ؼ�����',36,'����������');
Insert Into zlprogfuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'��Ƭ����',37,'������Ƭ');
Insert Into zlprogfuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'��ȡ����',38,'����ȡ��');
Insert Into zlprogfuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'��������',39,'�������');
Insert Into zlprogfuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'���ﷴ��',40,'����������');
Insert Into zlprogfuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'�����黯',41,'�����黯����');
Insert Into zlprogfuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'����Ⱦɫ',42,'����Ⱦɫ����');
Insert Into zlprogfuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'���Ӳ���',43,'���Ӳ������');
Insert Into zlprogfuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'�������',44,'��������Ϣ');
Insert Into zlprogfuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'���巴��',45,'����ʹ���������');
Insert Into zlprogfuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'�ײ�ά��',46,'ά�������ײ���Ϣ');
Insert Into zlprogfuncs(ϵͳ,���,����,����,˵��) Values(100,1294,'�����ؼ챨�����',47,'���ĺͳ��������ؼ챨��');


--����Ȩ��
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','��Ա����˵��',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','��������˵��',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','������Ϣ',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','�Ա�',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','�ѱ�',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','����',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ְҵ',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','����״��',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ҽ�Ƹ��ʽ',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','������ҳ',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','����ҽ����¼',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','����ҽ������',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','����ҽ������',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','���Ӳ�����¼',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','���Ӳ�������',User,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','��������Ӧ��',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','����ҽ������',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','����ҽ��ִ��',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','�շ���ĿĿ¼',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','����ҽ������',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','������ü�¼',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','סԺ���ü�¼',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ҩƷ���',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ҩƷ����',User,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ҩƷ���',User,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ҽ��ִ�з���',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','������Ŀ����',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','����ִ�п���',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','������Ŀ��λ',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','���Ƽ�鲿λ',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','������Ŀ���',User,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','�����÷�����',User,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','���Ʒ���Ŀ¼',User,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','���ƻ�����Ŀ',User,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','���Ƽ���걾',User,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','���Ƹ�����Ŀ',user,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','������ĿĿ¼',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','�����ļ��б�',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','�������ݸ���',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','������Ʊ�',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','����ҽ����¼_ID',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','�������',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','��λ״����¼',User,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','���˹Һż�¼',User,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','������Ŀ�ο�',User,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','���鱨����Ŀ',User,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Ӱ���豸Ŀ¼',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Ӱ�����¼',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Ӱ������Ŀ',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Ӱ�������',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Ӱ����ͼ��',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Ӱ��������',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Ӱ����ʱ��¼',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Ӱ����ʱͼ��',User,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Ӱ����ʱ����',User,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Ӱ����Ļ����',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Ӱ��Ԥ�贰��λ',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Ӱ��ͼ��������',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Ӱ����갴ť����',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Ӱ����������',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Ӱ���ע�洢��',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Ӱ��ͼ����Ϣ��',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Ӱ���ӡ������',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Ӱ��Ƭ���',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Ӱ���ӡ��ʽ',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Ӱ��Ƭ��ӡ����',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Ӱ����ɫ�嵥',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Ӱ����UID���_ID',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','��Ӱ��',User,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','������Ӱ��',User,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Ӱ�����̲���',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','�����ļ��ṹ',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Ӱ�������',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','������������¼',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','���ʱ�����',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','����ģ�����',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ҽ�����˹�����',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ҽ�����˵���',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1294,'����',USER,'�����ʾ�ʾ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1294,'����',USER,'�����ʾ����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1294,'����',USER,'�����ʾ�����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Ӱ�������¼',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Ӱ��ͼ��ע',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','�����ٴ�·��',User,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Ӱ����걾',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Ӱ��걾����ȡ��',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Ӱ����걾��λ',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','�����շѹ�ϵ',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Zl_Ӱ�������¼_Update',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1294,'����',USER,'f_Sentence_Matched','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ZL_Ӱ�񱨸�����_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ZL_Ӱ�񱨸�����_update',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ZL_Ӱ�񱨸��ע_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ZL_Ӱ�񱨸�ǩ��_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ZL_Ӱ�񱨸����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ZL_Ӱ�񱨸�ͼ��_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Zl_Ӱ����_��鼼ʦ',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ZL_Ӱ����_STATE',User,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ZL_Ӱ�����ִ��',User,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ZL_Ӱ��Ԥ�贰��λ_INSERT',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ZL_Ӱ��Ԥ�贰��λ_UPDATE',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ZL_Ӱ��Ԥ�贰��λ_DELETE',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ZL_Ӱ�񴰿�λ_����_UPDATE',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ZL_Ӱ����������_INSERT',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ZL_Ӱ����������_UPDATE',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ZL_Ӱ����Ļ����_����_INSERT',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ZL_Ӱ����Ļ����_����_UPDATE',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ZL_Ӱ����Ļ����_DELETE',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ZL_Ӱ����Ļ����_UPDATE',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ZL_Ӱ��ͼ��������_����_INSERT',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ZL_Ӱ��ͼ��������_����_Update',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ZL_Ӱ��ͼ��������_UPDATE',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Zl_Ӱ��ͼ��������_Delete',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ZL_Ӱ����갴ť����_INSERT',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ZL_Ӱ����갴ť����_UPDATE',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ZL_Ӱ����_���',USER,'EXECUTE');
Insert Into zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Zl_��ɫͨ��_Update',USER,'EXECUTE');
Insert Into Zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) values(100,1294,'����','ZL_Ӱ�񱨸��ӡ_Update',USER,'EXECUTE');
Insert Into Zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) values(100,1294,'����','ZL_Ӱ�񱨸汣��_Update',USER,'EXECUTE');
Insert Into zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) values(100,1294,'����','ZL_Ӱ�񱨸���_Clear',User,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ZL_������Ӱ��_INSERT',User,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ZL_Ӱ�񱨸����_Update',User,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ZL_����ҽ��ִ��_ȡ���ܾ�',User,'EXECUTE');
Insert Into zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Zl_���Ӳ�����¼_Print',User,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','Zl_Ӱ��ͼ��ע_Insert',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ZL_GetNumber',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','ZL_AgeToDays',USER,'EXECUTE');

--�������Ȩ��
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','��������Ϣ',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','����걾��Ϣ',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','�����ͼ���Ϣ',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','����ȡ����Ϣ',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','�����Ѹ���Ϣ',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','������Ƭ��Ϣ',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','�����ؼ���Ϣ',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','������̱���',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','����������Ϣ',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','�������ӳ�',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','���������Ϣ',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','��������Ϣ',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','�����巴��',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','�����ײ���Ϣ',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','�����ײ͹���',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','����鵵��Ϣ',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','���������Ϣ',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','�����ʾ����',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����','�����ʾ����',USER,'SELECT');


Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���Ǽ�','zl_�ҺŲ��˲���_INSERT',User,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���Ǽ�','ZL_����ҽ����¼_Insert',User,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���Ǽ�','ZL_����ҽ������_Insert',User,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���Ǽ�','Zl_����ҽ������_Insert',User,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���Ǽ�','ZL_Ӱ����_SET',User,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���Ǽ�','ZL_Ӱ����_BEGIN',User,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���Ǽ�','ZL_����ҽ��ִ��_�ܾ�ִ��',User,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���Ǽ�','zl_���˷��ü�¼_ҽ��',User,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���Ǽ�','NextNO',User,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���Ǽ�','Zl_������Ϣ_Update',User,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'ȡ������','ZL_Ӱ����_CANCEL',User,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���ͼ��','ZL_Ӱ����_PhotoDelete',User,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���ͼ��','ZL_Ӱ����_SET',User,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���ͼ��','ZL_Ӱ����_PhotoCancel',User,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'��Ƶ�ɼ�','ZL_Ӱ��ͼ��_DELETE',User,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'��Ƶ�ɼ�','ZL_Ӱ�����¼_SET',User,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'��Ƶ�ɼ�','ZL_Ӱ������_INSERT',User,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'��Ƶ�ɼ�','ZL_Ӱ��ͼ��_INSERT',User,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'��Ƶ�ɼ�','ZL_Ӱ���鱨��_ADD',User,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'Ӱ���ʿ�','Zl_Ӱ������_Update',User,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�Ŷӽк�','�ŶӽкŶ���',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�Ŷӽк�','�Ŷ���������',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�Ŷӽк�','�Ŷ�LED��ʾ����',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�Ŷӽк�','ZL_�ŶӽкŶ���_INSERT',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�Ŷӽк�','ZL_�ŶӽкŶ���_UPDATE',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�Ŷӽк�','ZL_�ŶӽкŶ���_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�Ŷӽк�','ZL_�ŶӽкŶ���_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�Ŷӽк�','ZL_�ŶӽкŶ���_DELETE',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�Ŷӽк�','ZL_�Ŷ���������_INSERT',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�Ŷӽк�','ZL_�Ŷ���������_DELETE',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���','Ӱ����Ϸ���',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���','Zl_Ӱ�����_Update',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���','Zl_Ӱ����Ϸ���_Update',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'��������','ZL_Ӱ���������',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'��������','ZL_Ӱ��ȡ����������',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'��鱨��','zl_�ҺŲ��˲���_INSERT',User,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'��鱨��','ZL_Ӱ����_SET',User,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'��鱨��','ZL_Ӱ����_BEGIN',User,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'��鱨��','NextNO',User,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'δ�ɷѱ���','�շ���Ŀ���',USER,'SELECT');




--�걾����
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�걾����','Zl_����걾_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�걾����','Zl_����걾_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�걾����','Zl_����걾_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�걾����','Zl_����걾_����',USER,'EXECUTE');

--�걾ȡ��
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����ȡ��','Zl_�����Ѹ�_��ʼ',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����ȡ��','Zl_�����Ѹ�_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����ȡ��','Zl_�����Ѹ�_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����ȡ��','Zl_�����Ѹ�_���',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����ȡ��','Zl_����ȡ��_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����ȡ��','Zl_����ȡ��_�������',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����ȡ��','Zl_����ȡ��_ϸ��',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����ȡ��','Zl_����ȡ��_ϸ������',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����ȡ��','Zl_����ȡ��_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����ȡ��','Zl_����ȡ��_��������',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����ȡ��','Zl_����ȡ��_ɾ��',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����ȡ��','Zl_����ȡ��_��Ϣ����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����ȡ��','Zl_����ȡ��_ȷ��',USER,'EXECUTE');

--������Ƭ
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'������Ƭ','Zl_������Ƭ_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'������Ƭ','Zl_������Ƭ_�嵥��ӡ',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'������Ƭ','Zl_������Ƭ_ȷ��',USER,'EXECUTE');

--�����ӳ�
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�����ӳ�','Zl_�������ӳ�_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�����ӳ�','Zl_�������ӳ�_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�����ӳ�','Zl_�������ӳ�_ɾ��',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�����ӳ�','Zl_�������ӳ�_��ӡ',USER,'EXECUTE');

--���������ߣ���Ⱦ�����ӹ��̱��棬�����ؼ챨�����
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'��������','Zl_������̱���_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'��������','Zl_������̱���_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'��������','Zl_������̱���_ɾ��',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'��������','Zl_������̱���_״̬',USER,'EXECUTE');

Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���߱���','Zl_������̱���_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���߱���','Zl_������̱���_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���߱���','Zl_������̱���_ɾ��',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���߱���','Zl_������̱���_״̬',USER,'EXECUTE');

Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���ӱ���','Zl_������̱���_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���ӱ���','Zl_������̱���_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���ӱ���','Zl_������̱���_ɾ��',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���ӱ���','Zl_������̱���_״̬',USER,'EXECUTE');

Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'��Ⱦ����','Zl_������̱���_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'��Ⱦ����','Zl_������̱���_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'��Ⱦ����','Zl_������̱���_ɾ��',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'��Ⱦ����','Zl_������̱���_״̬',USER,'EXECUTE');

Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�����ؼ챨�����','Zl_������̱���_״̬',USER,'EXECUTE');

--�ؼ�����
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�ؼ�����','Zl_��������_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�ؼ�����','Zl_��������_ɾ��',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�ؼ�����','Zl_��������_�ؼ���Ŀ_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�ؼ�����','Zl_��������_�ؼ���Ŀ_ɾ��',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�ؼ�����','Zl_��������_�ؼ���Ŀ_����',USER,'EXECUTE');

--��Ƭ����
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'��Ƭ����','Zl_��������_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'��Ƭ����','Zl_��������_ɾ��',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'��Ƭ����','Zl_��������_��Ƭ��Ŀ_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'��Ƭ����','Zl_��������_��Ƭ��Ŀ_ɾ��',USER,'EXECUTE');

--��ȡ����
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'��ȡ����','Zl_��������_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'��ȡ����','Zl_��������_ɾ��',USER,'EXECUTE');

--��������
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'��������','Zl_�������_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'��������','Zl_�������_ɾ��',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'��������','Zl_�������_״̬',USER,'EXECUTE');

--���ﷴ��
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���ﷴ��','Zl_�������_����',USER,'EXECUTE');

--���߼��
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�����黯','Zl_�����ؼ�_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�����黯','Zl_�����ؼ�_�嵥��ӡ',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�����黯','Zl_�����ؼ�_��Ŀ¼��',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�����黯','Zl_�����ؼ�_ȷ��',USER,'EXECUTE');

--���Ӽ��
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���Ӳ���','Zl_�����ؼ�_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���Ӳ���','Zl_�����ؼ�_�嵥��ӡ',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���Ӳ���','Zl_�����ؼ�_��Ŀ¼��',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���Ӳ���','Zl_�����ؼ�_ȷ��',USER,'EXECUTE');

--��Ⱦ���
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����Ⱦɫ','Zl_�����ؼ�_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����Ⱦɫ','Zl_�����ؼ�_�嵥��ӡ',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����Ⱦɫ','Zl_�����ؼ�_��Ŀ¼��',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'����Ⱦɫ','Zl_�����ؼ�_ȷ��',USER,'EXECUTE');

--�������
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�������','Zl_������_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�������','Zl_������_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�������','Zl_������_ɾ��',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�������','Zl_������_ʹ��״̬',USER,'EXECUTE');

--���巴��
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���巴��','Zl_�����巴��_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���巴��','Zl_�����巴��_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'���巴��','Zl_�����巴��_ɾ��',USER,'EXECUTE');

--�ײ�ά��
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�ײ�ά��','Zl_�����ײ�_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�ײ�ά��','Zl_�����ײ�_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�ײ�ά��','Zl_�����ײ�_ɾ��',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�ײ�ά��','Zl_�����ײ͹���_����',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�ײ�ά��','Zl_�����ײ͹���_ɾ��',USER,'EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1294,'�ײ�ά��','Zl_�����ײ͹���_ɾ��1',USER,'EXECUTE');


--����̨�˵�
Insert Into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Values('ȱʡ',zlMenus_id.nextval, '188','Ӱ������վ','������վ','C',230,'���ڲ���걾���պ�ȡ�ġ�ͼ��ɼ���������д������ҽ���ͷ��õĵǼ�',100,1294);
