--���ϵͳ�˵�,��д��ZLSOFT�еĹ�������

--zlComponent����
insert into zlComponent(����,����,���汾,�ΰ汾,���汾,ϵͳ) values ('zl9Medical','��������',10,0,0,100);

----------------------------------------------
--zlPrograms����
----------------------------------------------
Insert Into zlPrograms(���,����,˵��,ϵͳ,����) Values(1850,'�����������','�������������޸�������ͼ���Ӧ�������Ŀ��',100,'zl9Medical');
Insert Into zlPrograms(���,����,˵��,ϵͳ,����) Values(1860,'���ԤԼ����','������ԤԼ�����뼰ȷ�ϡ�',100,'zl9Medical');
Insert Into zlPrograms(���,����,˵��,ϵͳ,����) Values(1861,'��칤������','��ɸ������Ŀ�ı�����д������ܽᡣ',100,'zl9Medical');
Insert Into zlPrograms(���,����,˵��,ϵͳ,����) Values(1862,'����������','ά�������Ա������������Ϣ���ϡ�',100,'zl9Medical');
----------------------------------------------
--zlProgFuncs����
----------------------------------------------
--�����������
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1850,'����','');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1850,'��ɾ��','');
--���ԤԼ����
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1860,'����','');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1860,'���п���','');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1860,'���ԤԼ','');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1860,'ȷ��ԤԼ','');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1860,'ȡ��ԤԼ','');
--��칤������
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1861,'����','');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1861,'���п���','');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1861,'��ʼ���','');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1861,'ȡ����ʼ','');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1861,'������','');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1861,'ȡ�����','');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1861,'�����Ŀ','');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1861,'������Ŀ','');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1861,'��ӳ�Ա','');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1861,'�Ƴ���Ա','');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1861,'��д����','');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1861,'��д�ܽ�','');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1861,'��ӡ����','');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1861,'�ۺϲ�ѯ','');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1861,'���ô���','');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1861,'δ�շ����','');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1861,'����С��','');

--����������
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1862,'����','');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1862,'���п���','');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1862,'������','');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1862,'��������','');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1862,'�����ش�','');

----------------------------------------------
--zlProgPrivs����
----------------------------------------------
--�����������
--����
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1850,'����',user,'�������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1850,'����',user,'�������Ŀ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1850,'����',user,'������ĿĿ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1850,'����',user,'������Ŀ����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1850,'����',user,'���Ʒ���Ŀ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1850,'����',user,'�����շѹ�ϵ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1850,'����',user,'�շѼ�Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1850,'����',user,'�շ���ĿĿ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1850,'����',user,'���鱨����Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1850,'����',user,'������Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1850,'����',user,'����������Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1850,'����',user,'��λ״����¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1850,'����',user,'�����÷�����','SELECT');

--��ɾ��
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1850,'��ɾ��',user,'ZL_�������_INSERT','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1850,'��ɾ��',user,'ZL_�������_UPDATE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1850,'��ɾ��',user,'ZL_�������_DELETE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1850,'��ɾ��',user,'ZL_�������Ŀ¼_INSERT','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1850,'��ɾ��',user,'ZL_�������Ŀ¼_DELETE','EXECUTE');


--���ԤԼ����
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'�����÷�����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'��λ״����¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'�������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'�������Ŀ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'������ĿĿ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'������Ŀ����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'���Ʒ���Ŀ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'����ִ�п���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'������Ŀ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'�����շѹ�ϵ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'�շѼ�Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'�շ���ĿĿ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'��������˵��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'������Ա','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'���ǼǼ�¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'�����Ŀ�嵥','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'�����Ŀ�嵥_ID','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'�����Ա����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'�����Ա����_ID','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'������Ϣ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'���ǼǼ�¼_ID','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'��Լ��λ_id','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'��Լ��λ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'������Ŀ�ο�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'���鱨����Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'������Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'����������Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'������Ʊ�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'������Ʊ�','UPDATE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'ϵͳ������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'�Ա�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'����״��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'ѧ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'ְҵ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'�ѱ�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'ZL_���ǼǼ�¼_STATE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'����',user,'ZL_���ǼǼ�¼_�������','EXECUTE');

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'���ԤԼ',user,'ZL_������_INSERT','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'���ԤԼ',user,'ZL_������_DELETE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'���ԤԼ',user,'ZL_�����Ŀ�嵥_DELETE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'���ԤԼ',user,'ZL_�����Ŀ�嵥_INSERT','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'���ԤԼ',user,'ZL_�����Ա����_DELETE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'���ԤԼ',user,'ZL_�����Ա����_UPDATE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'���ԤԼ',user,'ZL_�����Ա����_INSERT','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'���ԤԼ',user,'ZL_���ǼǼ�¼_INSERT','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'���ԤԼ',user,'ZL_���ǼǼ�¼_UPDATE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'���ԤԼ',user,'ZL_���ǼǼ�¼_DELETE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'���ԤԼ',user,'ZL_�����Ա����_CLASS','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'���ԤԼ',user,'zl_������Ϣ_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'���ԤԼ',user,'zl_������Ϣ_Update','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'���ԤԼ',user,'zl_��Լ��λ_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1860,'���ԤԼ',user,'zl_��Լ��λ_Update','EXECUTE');

--��칤������

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'���ǼǼ�¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'�����Ա����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'�����Ա����_ID','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'�����Ŀ�嵥','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'�����Ŀ�嵥_ID','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'�����Ŀҽ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'�������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'�������Ŀ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'������Ա','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'Ӱ������Ŀ','SELECT');

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'����ģ�����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'����ģ��Ӧ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'����ģ������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'��������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ѧ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'�ѱ�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����ҽ����¼_ID','SELECT');

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'������Ʊ�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'������Ʊ�','UPDATE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'zlGetReference','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'zlGetResult','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'���˲�����¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'������Ϣ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'�������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����ҽ����¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����ҽ������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'������ҳ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'�����ļ�Ŀ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'��������˵��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����걾��¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'������ͨ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'���鱨����Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'������Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'��������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����ҩ�����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'�����ÿ�����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����ϸ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����걾��̬','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����걾��¼_ID','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'������ͨ���_ID','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'������Ŀȡֵ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'������Ŀ�ο�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����������Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'���鿹������ҩ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'���鿹������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����ϸ��������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����ϸ������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'���Ʒ���Ŀ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'������Ŀ����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'������ĿĿ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'���Ƽ���걾','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'ҽ�Ƹ��ʽ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'�շ�ִ�п���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'���Ű���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����ҽ������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����ҽ��ִ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'�����շѹ�ϵ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'�շѼ�Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'�շ���ĿĿ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'������Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'�շ���Ŀ����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'�շ���Ŀ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'ҩƷ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'���ʱ�����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����ģ�����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'�շѴ�����Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'ҩƷ�շ���¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'ҩƷ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'δ��ҩƷ��¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'��������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'ҩƷ��������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����������¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'�����������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����������Ա','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'������ϼ�¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����ҽ��״̬','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'���˹�����¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'���˷��ü�¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'���˲����ı���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'���˲����ⲿͼ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'���˲���������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'���˲�������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'���˲������ͼ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'���˱䶯��¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'������ҳ�ӱ�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����Ԫ��Ŀ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'�������ͼ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����ʾ��Ŀ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'�����ļ����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'ҽ��ִ�з���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'������λ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'ҩƷ���ʷ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'ҩƷ��;����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'ҩƷ����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'ҩƷ������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'�շ�ִ�в���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'�շѷ���Ŀ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'������Ŀ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'������������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'������������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����������Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����ִ�п���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'�����÷�����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'������Ŀ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����Ƶ����Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����Ƶ��ʱ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'���ƻ�����Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'���Ƶ���Ӧ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'�������Ŀ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'������ϱ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'������Ϸ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'�����������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'������϶���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'��������Ŀ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'�����������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'������ϲο�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'�����ο���Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'�����������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'�������ƴ�ʩ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'������Ϲ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'��λ״����¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'��Ա����˵��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����״��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'��Լ��λ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'��ҩ�����ע','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'ְҵ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'Ѫ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'�Ա�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'ϵͳ������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'�������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����ϵ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'��������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����������ģ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'ZL_�����¼��Ŀ_UPDATE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'zl_PatiDayCharge','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'�������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'���ղ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'�����ʻ�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'����֧����Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'����֧������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',User,'���˲����޶���¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',User,'���˹Һż�¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'����ҽ���Ƽ�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'�ѱ���ϸ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'���˹���ҩ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'��������','UPDATE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'����������ҩ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_���˷��ü�¼_����ҽ��','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_���˷��ü�¼_�ϴ�','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_���˼��ʼ�¼_�ϴ�','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_�����Ա����_����','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'���ղ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'������׼��Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'������Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'�ʻ������Ϣ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'���˲����޶���¼_ID','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'�����Ա����','SELECT');

insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'����',USER,'ZL_�����Ա����_REFRESH','EXECUTE');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'����',USER,'ZL_�����Ա����_UPDATE','EXECUTE');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'����',USER,'ZL_����ҽ������_�Ʒ�','EXECUTE');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'����',USER,'zl_���ﻮ�ۼ�¼_Insert','EXECUTE');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'����',USER,'zl_������ʼ�¼_Insert','EXECUTE');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'����',USER,'zl_סԺ���ʼ�¼_Insert','EXECUTE');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'����',USER,'zl_סԺ���ʼ�¼_DELETE','EXECUTE');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'����',USER,'zl_������ʼ�¼_DELETE','EXECUTE');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'����',USER,'zl_���ﻮ�ۼ�¼_DELETE','EXECUTE');

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'���˲�����¼_ID','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'���˲�������_ID','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'���������¼_ID','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_���˲���_INSERT','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_���˲���_�鵵','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_���˲���_����','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_���˲������ͼ_SAVE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_���˲������ӱ�_SAVE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_���˲������ӱ�Ԫ_SAVE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_���˲�������_DELETE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_���˲�������_INSERT','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_���˲���������_SAVE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_���˲����ⲿͼ_���','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_���˲����ı���_SAVE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_���˳�Ժ��ϼ�¼��_INSERT','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_����������ϼ�¼��_INSERT','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_����������Ҫ�����¼_INSERT','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_��������¼_INSERT','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'����ʾ��Ŀ¼_ID','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_����ʾ��Ŀ¼_UPDATE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_����ʾ��Ŀ¼_INSERT','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_����ʾ��Ŀ¼_DELETE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_���˲���_DELETE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'ZL_���˲�����¼����_INSERT','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'ZL_���������¼����_INSERT','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'ZL_���������¼��ҩ_INSERT','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'ZL_���������¼��ע_INSERT','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',USER,'ZL_���������¼��ע_DELETE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_���˲����޶�_INSERT','EXECUTE');

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_���Ƶ���_����','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_���Ƶ���_����','EXECUTE');

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'zl_������Ϣ_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'zl_������Ϣ_Update','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_�����Ա����_UPDATE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_�����Ա����_INSERT','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_���ǼǼ�¼_INSERT','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_���ǼǼ�¼_UPDATE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_���ǼǼ�¼_DELETE','EXECUTE');

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_���ǼǼ�¼_Cancel','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_�����Ŀҽ��_INSERT','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_����ҽ������_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_����ҽ����¼_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_���ǼǼ�¼_STATE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_�����Ա����_�ܽ�','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'zl_���˲���_DELETE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_�����Ա����_����','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_���ǼǼ�¼_Finish','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_���ǼǼ�¼_CancelFinish','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_����ҽ������_Insert','EXECUTE');

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_�����Ա����_DELETE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_���ǼǼ�¼_ItemCancel','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_�����Ŀ�嵥_DELETE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_�����Ŀ�嵥_INSERT','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_���ǼǼ�¼_������д','EXECUTE');

--����������

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1862,'����',user,'�������¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1862,'����',user,'�������¼_ID','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1862,'����',user,'�������嵥','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1862,'����',user,'ϵͳ������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1862,'����',user,'���ǼǼ�¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1862,'����',user,'���˽��ʼ�¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1862,'����',user,'���˷��ü�¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1862,'����',user,'��Լ��λ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1862,'����',user,'������Ʊ�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1862,'����',user,'������Ʊ�','UPDATE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1862,'����',user,'����Ԥ����¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1862,'����',user,'��������˵��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1862,'����',user,'������Ա','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1862,'����',user,'Ʊ��ʹ����ϸ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1862,'����',user,'Ʊ�����ü�¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1862,'����',user,'���㷽ʽӦ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1862,'����',user,'���㷽ʽ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1862,'����',user,'���˽��ʼ�¼_ID','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1862,'����',user,'����ҽ����¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1862,'����',user,'������Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1862,'����',user,'�շ���Ŀ����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1862,'����',user,'�շ���ĿĿ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1862,'��������',user,'zl_���˽��ʼ�¼_Delete','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1862,'������',user,'zl_���˽��ʼ�¼_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1862,'������',user,'zl_���˽���Ʊ��_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1862,'������',user,'zl_���ʽɿ��¼_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1862,'������',user,'zl_���ʷ��ü�¼_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1862,'�����ش�',user,'zl_���˽��ʼ�¼_RePrint','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1862,'������',user,'ZL_�������¼_INSERT','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1862,'������',user,'ZL_�������嵥_INSERT','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1862,'����',user,'ZL_�������¼_Cancel','EXECUTE');


----------------------------------------------
--zlMenus���� 
----------------------------------------------
Insert Into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Values('ȱʡ',zlMenus_id.nextval,null,'������ϵͳ','������','A',99,'',100,NULL);
Insert Into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Values('ȱʡ',zlMenus_id.nextval,zlMenus_id.nextval-1,'�����������','�������','A',99,'�������������޸�������ͼ���Ӧ�������Ŀ��',100,1850);
Insert Into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Values('ȱʡ',zlMenus_id.nextval,zlMenus_id.nextval-2,'���ԤԼ����','���ԤԼ','B',213,'������ԤԼ�����뼰ȷ�ϡ�',100,1860);
Insert Into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Values('ȱʡ',zlMenus_id.nextval,zlMenus_id.nextval-3,'��칤������','��칤��','D',225,'��ɸ������Ŀ�ı�����д������ܽᡣ',100,1861);
Insert Into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Values('ȱʡ',zlMenus_id.nextval,zlMenus_id.nextval-4,'����������','�������','E',99,'ά�������Ա������������Ϣ���ϡ�',100,1862);

--�����ϵͳ����
Alter Table ��Լ��λ Add �����ʼ� varchar(50);
Alter Table ��Լ��λ Add ˵�� varchar(2000);

CREATE INDEX ����ҽ����¼_IX_�Һŵ� ON ����ҽ����¼(�Һŵ�) PCTFREE 10 TABLESPACE zl9CisRec;
Insert Into �������ʷ���(����,����,����,������,˵��) Select 'Q','���','TJ',3,'' From dual;
insert into ������Ʊ�(��Ŀ���,��Ŀ����,������,�Զ���ȱ,��Ź���) values (78,'��쵥��','',1,null);

---------------------------------------------
-- ��"��Լ��λ"�����Ӳ���
----------------------------------------------
CREATE OR REPLACE PROCEDURE zl_��Լ��λ_Insert (
    ID_IN IN ��Լ��λ.ID%TYPE,
    �ϼ�ID_IN IN ��Լ��λ.�ϼ�ID%TYPE,
    ����_IN IN ��Լ��λ.����%TYPE,
    ����_IN IN ��Լ��λ.����%TYPE,
    ����_IN IN ��Լ��λ.����%TYPE := NULL,
    ��ַ_IN IN ��Լ��λ.��ַ%TYPE := NULL,
    �绰_IN IN ��Լ��λ.�绰%TYPE := NULL,
    ��������_IN IN ��Լ��λ.��������%TYPE := NULL,
    �ʺ�_IN IN ��Լ��λ.�ʺ�%TYPE := NULL,
    ��ϵ��_IN IN ��Լ��λ.��ϵ��%TYPE := NULL,
    ĩ��_IN IN ��Լ��λ.ĩ��%TYPE := 1,
    �����ʼ�_IN IN ��Լ��λ.�����ʼ�%TYPE := NULL,
    ˵��_IN IN ��Լ��λ.˵��%TYPE := NULL
)
IS
BEGIN
    --���Ȳ����¼
    Insert INTO ��Լ��λ
                    (
                        ID,
                        ����,
                        ����,
                        ����,
                        ��ַ,
                        �绰,
                        ��������,
                        �ʺ�,
                        ��ϵ��,
                        �ϼ�ID,
                        ����ʱ��,
                        ����ʱ��,
                        ĩ��,
			�����ʼ�,
			˵��
                    )
          VALUES (
              ID_IN,
              ����_IN,
              ����_IN,
              ����_IN,
              ��ַ_IN,
              �绰_IN,
              ��������_IN,
              �ʺ�_IN,
              ��ϵ��_IN,
              �ϼ�ID_IN,
              SYSDATE,
              TO_DATE ('3000-01-01', 'yyyy-mm-dd'),
              ĩ��_IN,
	      �����ʼ�_IN,
	      ˵��_IN
          );
EXCEPTION
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_��Լ��λ_Insert;
/
---------------------------------------------
-- ��"��Լ��λ"���޸Ĳ���
----------------------------------------------
CREATE OR REPLACE PROCEDURE zl_��Լ��λ_UPDATE (
    ID_IN IN ��Լ��λ.ID%TYPE,
    �ϼ�ID_IN IN ��Լ��λ.�ϼ�ID%TYPE,
    ����_IN IN ��Լ��λ.����%TYPE,
    ����_IN IN ��Լ��λ.����%TYPE,
    ����_IN IN ��Լ��λ.����%TYPE,
    ��ַ_IN IN ��Լ��λ.��ַ%TYPE := NULL,
    �绰_IN IN ��Լ��λ.�绰%TYPE := NULL,
    ��������_IN IN ��Լ��λ.��������%TYPE := NULL,
    �ʺ�_IN IN ��Լ��λ.�ʺ�%TYPE := NULL,
    ��ϵ��_IN IN ��Լ��λ.��ϵ��%TYPE := NULL,
    ԭ����_IN IN PLS_INTEGER,
    �����ʼ�_IN IN ��Լ��λ.�����ʼ�%TYPE := NULL,
    ˵��_IN IN ��Լ��λ.˵��%TYPE := NULL
)
IS
BEGIN
    --���Ȳ����޸ļ�¼
    UPDATE ��Լ��λ
        SET ���� = ����_IN,
             ���� = ����_IN,
             ���� = ����_IN,
             ��ַ = ��ַ_IN,
             �绰 = �绰_IN,
             �������� = ��������_IN,
             �ʺ� = �ʺ�_IN,
             ��ϵ�� = ��ϵ��_IN,
             �ϼ�ID = �ϼ�ID_IN,
	     �����ʼ�=�����ʼ�_IN,
	     ˵��=˵��_IN
     WHERE ID = ID_IN;

    --�������¼�ҲҪ�޸ı���
    UPDATE ��Լ��λ
        SET ���� = ����_IN || SUBSTR (����, ԭ����_IN)
     WHERE ID IN (SELECT ID
                         FROM ��Լ��λ
                        START WITH �ϼ�ID = ID_IN
                      CONNECT BY PRIOR ID = �ϼ�ID);
EXCEPTION
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_��Լ��λ_UPDATE;
/

--���ϵͳ���ݱ�
Create Table �������(
	���		NUMBER(18),
	�ϼ����	NUMBER(18),			--new
	����		VARCHAR2(10),	
	����		VARCHAR2(30),
	����		VARCHAR2(30),
	ĩ��		number(1) default 0,		--new
	˵��		VARCHAR2(100))
    TABLESPACE zl9BaseItem
    PCTFREE 10 PCTUSED 85 STORAGE (NEXT 255 PCTINCREASE 0 MAXEXTENTS UNLIMITED);

Create Table �������Ŀ¼(
	���		NUMBER(18),
	������Ŀid	NUMBER(18))	
    TABLESPACE zl9BaseItem
    PCTFREE 10 PCTUSED 85 STORAGE (NEXT 255 PCTINCREASE 0 MAXEXTENTS UNLIMITED);

Create Table ���ǼǼ�¼(
	ID		NUMBER(18),
	����		varchar2(10),		--����:����	
	��¼����	NUMBER(3),		--
	���״̬	NUMBER(3),		--1:�¿�ԤԼ;2:ȷ��ԤԼ;3:����ԤԼȷ��;4:�������;5:������
	��ϵ��		VARCHAR2(20),
	��ϵ�绰	VARCHAR2(30),
	�ƶ��绰	VARCHAR(20),
	��ϵ��ַ	VARCHAR2(50),
	��Լ��λid	NUMBER(18),
	�������	NUMBER(5),
	�����ۿ�	NUMBER(5,2) DEFAULT 1,		--new
	���ʱ��	DATE,
	�������	VARCHAR2(1000),		--����Ŀ��������new
	��첿��id	NUMBER(18),
	����˵��	VARCHAR2(2000),
	�Ƿ�����	NUMBER(1) DEFAULT 0,
	�Ǽ�ʱ��	DATE,
	���ʱ��	DATE)
    TABLESPACE zl9CisRec
    PCTFREE 10 PCTUSED 85 STORAGE (NEXT 255 PCTINCREASE 0 MAXEXTENTS UNLIMITED);

Create Table ������(
	�Ǽ�id		NUMBER(18),
	�������	VARCHAR2(30),
	˵��		VARCHAR2(100))
    TABLESPACE zl9CisRec
    PCTFREE 10 PCTUSED 85 STORAGE (NEXT 255 PCTINCREASE 0 MAXEXTENTS UNLIMITED);

CREATE TABLE �����Ŀ�嵥(
	ID		NUMBER(18),
	�Ǽ�id		NUMBER(18),
	�������	VARCHAR2(30),
	����id		NUMBER(18),		--������ֵʱ����ʾ�˲��˵�˽�������Ŀ
	������Ŀid	NUMBER(18),
	ִ�п���id	NUMBER(18),
	�ɼ���ʽid	NUMBER(18),
	�������	VARCHAR2(30),		--����Ŀ��������new
	����;��	number(1) default 1,	
	����걾	VARCHAR2(50),		--������Ŀ�ı걾����
	��鲿λ	VARCHAR2(4000),		--�ಿλʱ���Զ��ŷָ��������ƣ���
	��鲿λid	VARCHAR2(4000))		--�ಿλʱ���Զ��ŷָ�����id����234,67,9821))		
    TABLESPACE zl9CisRec
    PCTFREE 10 PCTUSED 85 STORAGE (NEXT 255 PCTINCREASE 0 MAXEXTENTS UNLIMITED);

CREATE TABLE �����Ŀҽ��(
	�嵥id		NUMBER(18),	
	����id		NUMBER(18),	
	ҽ��id		NUMBER(18))		--��¼����ҽ������ҽ��id(���idΪNULL)		
    TABLESPACE zl9CisRec
    PCTFREE 10 PCTUSED 85 STORAGE (NEXT 255 PCTINCREASE 0 MAXEXTENTS UNLIMITED);

Create Table �����Ա����(
	ID		NUMBER(18),
	�Ǽ�id		NUMBER(18),
	����id		NUMBER(18),
	���״̬	NUMBER(3) DEFAULT 1,		--1:ԤԼ;4:�������;5:������
	�������	VARCHAR2(30),
	����		VARCHAR2(20),			
	�Ա�		VARCHAR2(10),			
	����		VARCHAR2(20),			
	����״��	VARCHAR2(20),			
	��ϵ�绰	VARCHAR2(30),			
	�ƶ��绰	VARCHAR2(20),			
	��ϵ��ַ	VARCHAR2(50),			
	����ʱ��	DATE,
	��챨��	NUMBER(2) DEFAULT 0,
	��첡��id	number(18),
	���ʱ��	DATE)
    TABLESPACE zl9CisRec
    PCTFREE 10 PCTUSED 85 STORAGE (NEXT 255 PCTINCREASE 0 MAXEXTENTS UNLIMITED);

CREATE TABLE �����Ա����(
	�Ǽ�id		NUMBER(18),
	����id		NUMBER(18),
	����id		NUMBER(18),			--NULL,��ʾ�ܽ�
	����id		NUMBER(18))
    TABLESPACE zl9CisRec
    PCTFREE 10 PCTUSED 85 STORAGE (NEXT 255 PCTINCREASE 0 MAXEXTENTS UNLIMITED);

create table �������¼
(
  ID         NUMBER(18),
  ��¼״̬   NUMBER(1),
  ��Լ��λID NUMBER(18),
  ����ID     NUMBER(18),
  ���㲿��ID NUMBER(18),
  ������   NUMBER(16,5)
)
    TABLESPACE zl9CisRec
    PCTFREE 10 PCTUSED 85 STORAGE (NEXT 255 PCTINCREASE 0 MAXEXTENTS UNLIMITED);

create table �������嵥
(
  ����ID NUMBER(18),
  �Ǽ�ID NUMBER(18)
)
    TABLESPACE zl9CisRec
    PCTFREE 10 PCTUSED 85 STORAGE (NEXT 255 PCTINCREASE 0 MAXEXTENTS UNLIMITED);

Create Sequence ���ǼǼ�¼_ID start with 1;
Create Sequence �����Ա����_ID start with 1;
Create Sequence �����Ŀ�嵥_ID start with 1;
Create Sequence �������¼_ID start with 1;

--���ϵͳ����

CREATE INDEX �������Ŀ¼_IX_��� on �������Ŀ¼(���) PCTFREE 10 TABLESPACE zl9BaseItem;

CREATE INDEX ���ǼǼ�¼_IX_��Լ��λid on ���ǼǼ�¼(��Լ��λid) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX ���ǼǼ�¼_IX_��첿��id on ���ǼǼ�¼(��첿��id) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX ���ǼǼ�¼_IX_���ʱ�� on ���ǼǼ�¼(���ʱ��) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX ������_IX_�Ǽ�id on ������(�Ǽ�id) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX �����Ŀ�嵥_IX_�Ǽ�id on �����Ŀ�嵥(�Ǽ�id) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX �����Ŀ�嵥_IX_ִ�п���id on �����Ŀ�嵥(ִ�п���id) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX �����Ŀ�嵥_IX_������Ŀid on �����Ŀ�嵥(������Ŀid) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX �����Ŀ�嵥_IX_�ɼ���ʽid on �����Ŀ�嵥(�ɼ���ʽid) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX �����Ŀ�嵥_IX_����id on �����Ŀ�嵥(����id) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX �����Ŀҽ��_IX_�嵥id on �����Ŀҽ��(�嵥id) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX �����Ŀҽ��_IX_����id on �����Ŀҽ��(����id) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX �����Ŀҽ��_IX_ҽ��id on �����Ŀҽ��(ҽ��id) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX �����Ա����_IX_�Ǽ�id on �����Ա����(�Ǽ�id) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX �����Ա����_IX_����id on �����Ա����(����id) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX �����Ա����_IX_��첡��id on �����Ա����(��첡��id) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX �������¼_IX_��Լ��λid on �������¼(��Լ��λid) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX �������¼_IX_���㲿��ID on �������¼(���㲿��ID) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX �������¼_IX_����id on �������¼(����id) PCTFREE 10 TABLESPACE zl9CisRec;

--���ϵͳԼ��
ALTER TABLE ������� ADD CONSTRAINT �������_PK PRIMARY KEY (���) USING INDEX PCTFREE 15 TABLESPACE zl9BaseItem;
ALTER TABLE ������� ADD CONSTRAINT �������_UQ_���� UNIQUE (����) USING INDEX PCTFREE 5 TABLESPACE zl9BaseItem;
ALTER TABLE ������� ADD CONSTRAINT �������_UQ_���� UNIQUE (����) USING INDEX PCTFREE 5 TABLESPACE zl9BaseItem;
ALTER TABLE ������� ADD CONSTRAINT �������_CK_ȱʡ CHECK (ȱʡ IN(0,1));
ALTER TABLE ������� ADD CONSTRAINT �������_FK_�ϼ���� FOREIGN KEY (�ϼ����) REFERENCES �������(���) ON DELETE CASCADE;

ALTER TABLE �������Ŀ¼ ADD CONSTRAINT �������Ŀ¼_FK_��� FOREIGN KEY (���) REFERENCES �������(���) ON DELETE CASCADE;
ALTER TABLE �������Ŀ¼ ADD CONSTRAINT �������Ŀ¼_FK_������Ŀid FOREIGN KEY (������Ŀid) REFERENCES ������ĿĿ¼(ID);

ALTER TABLE ���ǼǼ�¼ ADD CONSTRAINT ���ǼǼ�¼_PK PRIMARY KEY (ID) USING INDEX PCTFREE 15 TABLESPACE zl9CisRec;
ALTER TABLE ���ǼǼ�¼ ADD CONSTRAINT ���ǼǼ�¼_UQ_���� UNIQUE (����) USING INDEX PCTFREE 5 TABLESPACE zl9CisRec;
ALTER TABLE ���ǼǼ�¼ ADD CONSTRAINT ���ǼǼ�¼_CK_�Ƿ����� CHECK (�Ƿ����� IN(0,1));
ALTER TABLE ���ǼǼ�¼ ADD CONSTRAINT ���ǼǼ�¼_CK_���״̬ CHECK (���״̬ IN(1,2,3,4,5));
ALTER TABLE ���ǼǼ�¼ ADD CONSTRAINT ���ǼǼ�¼_FK_��Լ��λid FOREIGN KEY (��Լ��λid) REFERENCES ��Լ��λ(ID);
ALTER TABLE ���ǼǼ�¼ ADD CONSTRAINT ���ǼǼ�¼_FK_��첿��id FOREIGN KEY (��첿��id) REFERENCES ���ű�(ID);

ALTER TABLE ������ ADD CONSTRAINT ������_PK PRIMARY KEY (�Ǽ�id,�������) USING INDEX PCTFREE 15 TABLESPACE zl9CisRec;
ALTER TABLE ������ ADD CONSTRAINT ������_FK_�Ǽ�id FOREIGN KEY (�Ǽ�id) REFERENCES ���ǼǼ�¼(ID);

ALTER TABLE �����Ŀ�嵥 ADD CONSTRAINT �����Ŀ�嵥_PK PRIMARY KEY (ID) USING INDEX PCTFREE 15 TABLESPACE zl9CisRec;
ALTER TABLE �����Ŀ�嵥 ADD CONSTRAINT �����Ŀ�嵥_FK_�Ǽ�id FOREIGN KEY (�Ǽ�id) REFERENCES ���ǼǼ�¼(ID);
ALTER TABLE �����Ŀ�嵥 ADD CONSTRAINT �����Ŀ�嵥_FK_������Ŀid FOREIGN KEY (������Ŀid) REFERENCES ������ĿĿ¼(ID);
ALTER TABLE �����Ŀ�嵥 ADD CONSTRAINT �����Ŀ�嵥_FK_ִ�п���id FOREIGN KEY (ִ�п���id) REFERENCES ���ű�(ID);
ALTER TABLE �����Ŀ�嵥 ADD CONSTRAINT �����Ŀ�嵥_FK_�ɼ���ʽid FOREIGN KEY (�ɼ���ʽid) REFERENCES ������ĿĿ¼(ID);
ALTER TABLE �����Ŀ�嵥 ADD CONSTRAINT �����Ŀ�嵥_FK_����id FOREIGN KEY (����id) REFERENCES ������Ϣ(����id);

ALTER TABLE �����Ŀҽ�� ADD CONSTRAINT �����Ŀҽ��_PK PRIMARY KEY (�嵥id,����id,ҽ��id) USING INDEX PCTFREE 15 TABLESPACE zl9CisRec;
ALTER TABLE �����Ŀҽ�� ADD CONSTRAINT �����Ŀҽ��_FK_�嵥id FOREIGN KEY (�嵥id) REFERENCES �����Ŀ�嵥(ID) ON DELETE CASCADE;
ALTER TABLE �����Ŀҽ�� ADD CONSTRAINT �����Ŀҽ��_FK_����id FOREIGN KEY (����id) REFERENCES ������Ϣ(����id);
ALTER TABLE �����Ŀҽ�� ADD CONSTRAINT �����Ŀҽ��_FK_ҽ��id FOREIGN KEY (ҽ��id) REFERENCES ����ҽ����¼(ID);

ALTER TABLE �����Ա���� ADD CONSTRAINT �����Ա����_PK PRIMARY KEY (ID) USING INDEX PCTFREE 15 TABLESPACE zl9CisRec;
ALTER TABLE �����Ա���� ADD CONSTRAINT �����Ա����_FK_�Ǽ�id FOREIGN KEY (�Ǽ�id) REFERENCES ���ǼǼ�¼(ID);
ALTER TABLE �����Ա���� ADD CONSTRAINT �����Ա����_FK_����id FOREIGN KEY (����id) REFERENCES ������Ϣ(����id);
ALTER TABLE �����Ա���� ADD CONSTRAINT �����Ա����_CK_���״̬ CHECK (���״̬ IN(1,4,5));
ALTER TABLE �����Ա���� ADD CONSTRAINT �����Ա����_FK_��첡��id FOREIGN KEY (��첡��id) REFERENCES ���˲�����¼(ID);

ALTER TABLE �����Ա���� ADD CONSTRAINT �����Ա����_PK PRIMARY KEY (�Ǽ�id,����id,����id) USING INDEX PCTFREE 15 TABLESPACE zl9CisRec;
ALTER TABLE �����Ա���� ADD CONSTRAINT �����Ա����_FK_�Ǽ�id FOREIGN KEY (�Ǽ�id) REFERENCES ���ǼǼ�¼(ID);
ALTER TABLE �����Ա���� ADD CONSTRAINT �����Ա����_FK_����id FOREIGN KEY (����id) REFERENCES ������Ϣ(����id);
ALTER TABLE �����Ա���� ADD CONSTRAINT �����Ա����_FK_����id FOREIGN KEY (����id) REFERENCES ���ű�(ID);
ALTER TABLE �����Ա���� ADD CONSTRAINT �����Ա����_FK_����id FOREIGN KEY (����id) REFERENCES ���˲�����¼(ID);

ALTER TABLE �������¼ ADD CONSTRAINT �������¼_PK PRIMARY KEY (ID) USING INDEX PCTFREE 15 TABLESPACE zl9CisRec;
ALTER TABLE �������¼ ADD CONSTRAINT �������¼_CK_��¼״̬ CHECK (��¼״̬ IN(1,2));
ALTER TABLE �������¼ ADD CONSTRAINT �������¼_FK_��Լ��λid FOREIGN KEY (��Լ��λid) REFERENCES ��Լ��λ(ID);
ALTER TABLE �������¼ ADD CONSTRAINT �������¼_FK_���㲿��ID FOREIGN KEY (���㲿��ID) REFERENCES ���ű�(ID);
ALTER TABLE �������¼ ADD CONSTRAINT �������¼_FK_����id FOREIGN KEY (����id) REFERENCES ���˽��ʼ�¼(ID);


--���ϵͳ����
----------------------------------------------------------------------------
---  INSERT   for   �������
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_�������_INSERT(
	���_IN IN �������.���%TYPE,
	����_IN IN �������.����%TYPE,
	����_IN IN �������.����%TYPE,
	����_IN IN �������.����%TYPE,
	˵��_IN IN �������.˵��%TYPE,
	�ϼ����_IN IN �������.�ϼ����%TYPE:=NULL,
	ĩ��_IN IN �������.ĩ��%TYPE:=1,
	ͬ������_IN  NUMBER:=0
)
IS
	v_Extend number(18);
	v_Parent varchar2(30);
BEGIN	
	IF ĩ��_IN=0 THEN
		IF ͬ������_IN=1 THEN
			    --����ͬ������ĳ���
			IF NVL(�ϼ����_IN,0)<>0 THEN
			    SELECT ���� INTO v_Parent FROM ������� WHERE ���=�ϼ����_IN;
			ELSE
			    v_Parent:=NULL;
			END IF;

			BEGIN
			    SELECT length(rtrim(����_IN))-length(rtrim(����)) INTO v_Extend
			    FROM �������
			    WHERE ĩ��=0 AND (�ϼ����=�ϼ����_IN OR �ϼ���� IS NULL AND NVL(�ϼ����_IN,0)=0) AND Rownum=1;
			EXCEPTION
			    WHEN OTHERS THEN v_Extend:=0;
			END;

			IF v_Extend>0 THEN
			    --���䴦��
			    IF v_Parent IS null THEN
				UPDATE ������� SET ����=lpad('0',v_Extend,'0')||���� WHERE ���<>���_IN AND ĩ��=0;
			    ELSE
				UPDATE ������� SET ����=v_Parent||lpad('0',v_Extend,'0')||substr(����,length(v_Parent)+1) WHERE ���� LIKE v_Parent||'_%' AND ĩ��=0;
			    END IF;
			END IF;

			IF v_Extend<0 THEN
			    --ѹ������
			    IF v_Parent IS null THEN
				UPDATE ������� SET ����=substr(����,1+abs(v_Extend)) WHERE ���<>���_IN AND ĩ��=0;
			    ELSE
				UPDATE ������� SET ����=v_Parent||substr(����,length(v_Parent)+1+abs(v_Extend)) WHERE ���� LIKE v_Parent||'_%' AND ĩ��=0;
			    END IF;
			END IF;

		END IF;
	END IF;
	Insert Into �������(���,�ϼ����,ĩ��,����,����,����,˵��) VALUES(���_IN,DECODE(�ϼ����_IN,0,NULL,�ϼ����_IN),ĩ��_IN,����_IN,����_IN,����_IN,˵��_IN);
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_�������_INSERT;
/
----------------------------------------------------------------------------
---  UPDATE   for   �������
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_�������_UPDATE(
	���_IN IN �������.���%TYPE,
	����_IN IN �������.����%TYPE,
	����_IN IN �������.����%TYPE,
	����_IN IN �������.����%TYPE,
	˵��_IN IN �������.˵��%TYPE,
	�ϼ����_IN IN �������.�ϼ����%TYPE:=NULL,
	ͬ������_IN  NUMBER:=0
)
IS
	v_OldCode  VARCHAR2(30);  --ԭ���ı���
	v_Parent  VARCHAR2(30);  --�ϼ�����
	v_Extend  NUMBER(18);    --���䳤��(Ϊ����ʾѹ��)
	Err_NotFind  EXCEPTION;
BEGIN
	
	SELECT rtrim(����) INTO v_OldCode FROM ������� WHERE ���=���_IN;
	IF v_OldCode is null THEN
		RAISE Err_NotFind;
	END IF;

	--�޸���Ŀ����
	Update �������
		Set ����=����_IN,
		    ����=����_IN,
		    ����=����_IN,
		    ˵��=˵��_IN,
		    �ϼ����=DECODE(�ϼ����_IN,0,NULL,�ϼ����_IN)
	WHERE ���=���_IN;    

	--�޸ı�ϵ������������

	UPDATE ������� SET ����=����_IN||substr(����,length(v_OldCode)+1) WHERE ����<>����_IN And ���� LIKE v_OldCode||'_%' And ĩ��=0;

	--����ͬ������ĳ���
	IF ͬ������_IN=1 THEN
		IF NVL(�ϼ����_IN,0)<>0 THEN
		    SELECT ���� INTO v_Parent FROM ������� WHERE ���=�ϼ����_IN;
		ELSE
		    v_Parent:=NULL;
		END IF;

		BEGIN
		    SELECT length(rtrim(����_IN))-length(rtrim(����)) INTO v_Extend FROM ������� WHERE ĩ��=0 AND (�ϼ����=�ϼ����_IN OR �ϼ���� IS NULL AND nvl(�ϼ����_IN,0)=0) AND ���<>���_IN AND Rownum=1;
		EXCEPTION
		    WHEN OTHERS THEN v_Extend:=0;
		END;

		IF v_Extend>0 THEN
		    --���䴦��
		    IF v_Parent IS null THEN
			UPDATE ������� SET ����=lpad('0',v_Extend,'0')||����  WHERE ĩ��=0 and ��� not in (select ��� from ������� WHERE ĩ��=0 start with ���=���_IN connect by prior ���=�ϼ����);
		    ELSE
			UPDATE �������	SET ����=v_Parent||lpad('0',v_Extend,'0')||substr(����,length(v_Parent)+1) WHERE ĩ��=0 AND ���� LIKE v_Parent||'_%' and ��� not in (select ��� from ������� where ĩ��=0 start with ���=���_IN connect by prior ���=�ϼ����);
		    END IF;
		END IF;

		IF v_Extend<0 THEN
		    --ѹ������
		    IF v_Parent IS null THEN
			UPDATE ������� SET ����=substr(����,1+abs(v_Extend)) WHERE ���<>���_IN AND ĩ��=0;
		    ELSE
			UPDATE ������� SET ����=v_Parent||substr(����,length(v_Parent)+1+abs(v_Extend)) WHERE ���� LIKE v_Parent||'_%' AND ���<>���_IN AND ĩ��=0;
		    END IF;
		END IF;
	END IF;
EXCEPTION
	WHEN Err_NotFind THEN Raise_application_error (-20101, '[ZLSOFT]����Ŀ�����ڣ������ѱ������û�ɾ����[ZLSOFT]');
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_�������_UPDATE;
/

----------------------------------------------------------------------------
---  DELETE   for   �������
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_�������_DELETE(
	���_IN IN �������.���%TYPE
)
IS
BEGIN
	DELETE FROM ������� WHERE ���=���_IN;
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_�������_DELETE;
/
----------------------------------------------------------------------------
---  INSERT   for   �������Ŀ¼
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_�������Ŀ¼_INSERT(
	���_IN IN �������Ŀ¼.���%TYPE,
	������Ŀid_IN IN �������Ŀ¼.������Ŀid%TYPE
)
IS
BEGIN
	Insert Into �������Ŀ¼(���,������Ŀid)
		VALUES(���_IN,������Ŀid_IN);
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_�������Ŀ¼_INSERT;
/
----------------------------------------------------------------------------
---  DELETE   for   �������Ŀ¼
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_�������Ŀ¼_DELETE(
	���_IN IN �������Ŀ¼.���%TYPE
)
IS
BEGIN
	DELETE FROM �������Ŀ¼ WHERE ���=���_IN;
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_�������Ŀ¼_DELETE;
/
----------------------------------------------------------------------------
---  INSERT   for   ���ǼǼ�¼
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_���ǼǼ�¼_INSERT(
	ID_IN IN ���ǼǼ�¼.ID%TYPE,
	����_IN IN ���ǼǼ�¼.����%TYPE,
	��¼����_IN IN ���ǼǼ�¼.��¼����%TYPE,
	���״̬_IN IN ���ǼǼ�¼.���״̬%TYPE,
	��ϵ��_IN IN ���ǼǼ�¼.��ϵ��%TYPE,
	��ϵ�绰_IN IN ���ǼǼ�¼.��ϵ�绰%TYPE,
	�ƶ��绰_IN IN ���ǼǼ�¼.�ƶ��绰%TYPE,
	��ϵ��ַ_IN IN ���ǼǼ�¼.��ϵ��ַ%TYPE,
	��Լ��λID_IN IN ���ǼǼ�¼.��Լ��λID%TYPE,
	�������_IN IN ���ǼǼ�¼.�������%TYPE,
	���ʱ��_IN IN ���ǼǼ�¼.���ʱ��%TYPE,
	��첿��ID_IN IN ���ǼǼ�¼.��첿��ID%TYPE,
	����˵��_IN IN ���ǼǼ�¼.����˵��%TYPE,
	�Ǽ�ʱ��_IN IN ���ǼǼ�¼.�Ǽ�ʱ��%TYPE,
	���ʱ��_IN IN ���ǼǼ�¼.���ʱ��%TYPE,
	�Ƿ�����_IN IN ���ǼǼ�¼.�Ƿ�����%TYPE:=0,
	�����ۿ�_IN IN ���ǼǼ�¼.�����ۿ�%TYPE:=1
)
IS
BEGIN
	Insert Into ���ǼǼ�¼
		(ID,����,��¼����,���״̬,��ϵ��,��ϵ�绰,�ƶ��绰,��ϵ��ַ,��Լ��λID,�������,���ʱ��,��첿��ID,����˵��,�Ǽ�ʱ��,���ʱ��,�Ƿ�����,�����ۿ�)
		VALUES
		(ID_IN,����_IN,��¼����_IN,���״̬_IN,��ϵ��_IN,��ϵ�绰_IN,�ƶ��绰_IN,��ϵ��ַ_IN,��Լ��λID_IN,�������_IN,���ʱ��_IN,��첿��ID_IN,����˵��_IN,�Ǽ�ʱ��_IN,���ʱ��_IN,�Ƿ�����_IN,DECODE(�����ۿ�_IN,0,1,�����ۿ�_IN));
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_���ǼǼ�¼_INSERT;
/
----------------------------------------------------------------------------
---  UPDATE   for   ���ǼǼ�¼
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_���ǼǼ�¼_UPDATE(
	ID_IN IN ���ǼǼ�¼.ID%TYPE,
	����_IN IN ���ǼǼ�¼.����%TYPE,
	��¼����_IN IN ���ǼǼ�¼.��¼����%TYPE,
	���״̬_IN IN ���ǼǼ�¼.���״̬%TYPE,
	��ϵ��_IN IN ���ǼǼ�¼.��ϵ��%TYPE,
	��ϵ�绰_IN IN ���ǼǼ�¼.��ϵ�绰%TYPE,
	�ƶ��绰_IN IN ���ǼǼ�¼.�ƶ��绰%TYPE,
	��ϵ��ַ_IN IN ���ǼǼ�¼.��ϵ��ַ%TYPE,
	��Լ��λID_IN IN ���ǼǼ�¼.��Լ��λID%TYPE,
	�������_IN IN ���ǼǼ�¼.�������%TYPE,
	���ʱ��_IN IN ���ǼǼ�¼.���ʱ��%TYPE,
	��첿��ID_IN IN ���ǼǼ�¼.��첿��ID%TYPE,
	����˵��_IN IN ���ǼǼ�¼.����˵��%TYPE,
	�Ǽ�ʱ��_IN IN ���ǼǼ�¼.�Ǽ�ʱ��%TYPE,
	���ʱ��_IN IN ���ǼǼ�¼.���ʱ��%TYPE,
	�����ۿ�_IN IN ���ǼǼ�¼.�����ۿ�%TYPE:=1
)
IS
BEGIN
	Update ���ǼǼ�¼
		Set 		    
		    ����=����_IN,
		    ��¼����=��¼����_IN,
		    ���״̬=���״̬_IN,
		    ��ϵ��=��ϵ��_IN,
		    ��ϵ�绰=��ϵ�绰_IN,
		    �ƶ��绰=�ƶ��绰_IN,
		    ��ϵ��ַ=��ϵ��ַ_IN,
		    ��Լ��λID=��Լ��λID_IN,
		    �������=�������_IN,
		    ���ʱ��=���ʱ��_IN,
		    ��첿��ID=��첿��ID_IN,
		    ����˵��=����˵��_IN,
		    �Ǽ�ʱ��=�Ǽ�ʱ��_IN,
		    ���ʱ��=���ʱ��_IN,
		    �����ۿ�=DECODE(�����ۿ�_IN,0,1,�����ۿ�_IN)
	WHERE ID=ID_IN;
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_���ǼǼ�¼_UPDATE;
/

----------------------------------------------------------------------------
---  GROUP   for   ���ǼǼ�¼
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_���ǼǼ�¼_GROUP(
	ID_IN IN ���ǼǼ�¼.ID%TYPE,
	��Լ��λID_IN IN ���ǼǼ�¼.��Լ��λID%TYPE
)
IS
BEGIN
	Update ���ǼǼ�¼
		Set 
		    ��Լ��λID=��Լ��λID_IN
	WHERE ID=ID_IN;
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_���ǼǼ�¼_GROUP;
/

----------------------------------------------------------------------------
---  DELETE   for   ���ǼǼ�¼
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_���ǼǼ�¼_DELETE(
	ID_IN IN ���ǼǼ�¼.ID%TYPE
)
IS
BEGIN
	Delete from �����Ա���� WHERE �Ǽ�id=ID_IN;
	Delete from �����Ŀ�嵥 WHERE �Ǽ�id=ID_IN;
	Delete from ������ WHERE �Ǽ�id=ID_IN;
	Delete From ���ǼǼ�¼ WHERE ID=ID_IN;
EXCEPTION
	WHEN OTHERS THEN
		zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_���ǼǼ�¼_DELETE;
/
----------------------------------------------------------------------------
---  STATE   for   ���ǼǼ�¼
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_���ǼǼ�¼_STATE(
	ID_IN		IN ���ǼǼ�¼.ID%TYPE,
	���״̬_IN	IN ���ǼǼ�¼.���״̬%TYPE,
	����id_IN	IN �����Ա����.����id%TYPE:=0
)
IS
BEGIN
	IF ����id_IN=0 THEN
		UPDATE ���ǼǼ�¼ SET ���״̬=���״̬_IN WHERE ID=ID_IN;
		IF ���״̬_IN=4 THEN
			UPDATE �����Ա���� SET ���״̬=���״̬_IN WHERE �Ǽ�id=ID_IN;
			UPDATE ���ǼǼ�¼ SET ���ʱ��=SYSDATE WHERE ID=ID_IN;
		END IF;
	ELSE
		UPDATE �����Ա���� SET ���״̬=���״̬_IN WHERE �Ǽ�id=ID_IN AND ����id=����id_IN;		
	END IF;
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_���ǼǼ�¼_STATE;
/

----------------------------------------------------------------------------
---  INSERT   for   ������
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_������_INSERT(
	�Ǽ�id_IN IN ������.�Ǽ�id%TYPE,
	�������_IN IN ������.�������%TYPE,
	˵��_IN IN ������.˵��%TYPE:=null
)
IS
BEGIN
	Insert Into ������(�Ǽ�id,�������,˵��)
		VALUES
		(�Ǽ�id_IN,�������_IN,˵��_IN);
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_������_INSERT;
/

----------------------------------------------------------------------------
---  DELETE   for   ������
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_������_DELETE(
	�Ǽ�id_IN IN ������.�Ǽ�id%TYPE
)
IS
BEGIN
	Delete from ������ where �Ǽ�id=�Ǽ�id_IN;
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_������_DELETE;
/
----------------------------------------------------------------------------
---  INSERT   for   ��������Ŀ
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_�����Ŀ�嵥_INSERT(
	�Ǽ�id_IN IN �����Ŀ�嵥.�Ǽ�id%TYPE,
	�������_IN IN �����Ŀ�嵥.�������%TYPE,
	������Ŀid_IN IN �����Ŀ�嵥.������Ŀid%TYPE,
	�������_IN IN �����Ŀ�嵥.�������%TYPE,
	ִ�п���id_IN IN �����Ŀ�嵥.ִ�п���id%TYPE:=NULL,
	�ɼ���ʽid_IN IN �����Ŀ�嵥.�ɼ���ʽid%TYPE:=NULL,
	����걾_IN IN �����Ŀ�嵥.����걾%TYPE:=NULL,
	��鲿λ_IN IN �����Ŀ�嵥.��鲿λ%TYPE:=NULL,
	��鲿λid_IN IN �����Ŀ�嵥.��鲿λid%TYPE:=NULL,
	����id_IN IN �����Ŀ�嵥.����id%TYPE:=0,
	����;��_IN IN �����Ŀ�嵥.����;��%TYPE:=1
)
IS
BEGIN
	Insert Into �����Ŀ�嵥(ID,�Ǽ�id,�������,������Ŀid,ִ�п���id,�ɼ���ʽid,����걾,��鲿λ,��鲿λid,����id,�������,����;��)
	VALUES(�����Ŀ�嵥_ID.NEXTVAL,�Ǽ�id_IN,�������_IN,������Ŀid_IN,DECODE(ִ�п���id_IN,0,NULL,ִ�п���id_IN),DECODE(�ɼ���ʽid_IN,0,NULL,�ɼ���ʽid_IN),����걾_IN,��鲿λ_IN,��鲿λid_IN,DECODE(����id_IN,0,NULL,����id_IN),�������_IN,����;��_IN);

EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_�����Ŀ�嵥_INSERT;
/

----------------------------------------------------------------------------
---  �������   for   ���ǼǼ�¼
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_���ǼǼ�¼_�������(
	�Ǽ�id_IN IN ���ǼǼ�¼.ID%TYPE
)
IS
	Cursor c_Type is
		SELECT DISTINCT ������� FROM �����Ŀ�嵥 WHERE ������� IS NOT NULL AND �Ǽ�id=�Ǽ�id_IN;

	v_������� VARCHAR2(1000);
BEGIN
	For r_Type IN c_Type Loop
		IF INSTR(';'||v_�������||';',';'||r_Type.�������||';')<=0 THEN
			v_�������:=v_�������||';'||r_Type.�������;		
		END IF;
	end loop;
	IF v_������� IS NOT NULL THEN 
		v_�������:=substr(v_�������,2,length(v_�������)-1);
	END IF;
	UPDATE ���ǼǼ�¼ SET �������=v_������� WHERE ID=�Ǽ�id_IN;
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_���ǼǼ�¼_�������;
/

----------------------------------------------------------------------------
---  DELETE   for   ��������Ŀ
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_�����Ŀ�嵥_DELETE(
	�Ǽ�id_IN IN �����Ŀ�嵥.�Ǽ�id%TYPE,
	�������_IN		varchar2:=null,
	������Ŀid_IN	number:=0,
	����id_IN	number:=0
)
IS
	Cursor c_Items is
		SELECT A.* FROM �����Ŀ�嵥 A,���ǼǼ�¼ B  WHERE A.�Ǽ�id=B.ID AND B.ID=�Ǽ�id_IN AND A.�������=�������_IN AND A.������Ŀid=������Ŀid_IN;

BEGIN

	if ������Ŀid_IN=0 then
		Delete from �����Ŀ�嵥 where �Ǽ�id=�Ǽ�id_IN;
	else
		if �������_IN IS NULL THEN
			Delete from �����Ŀ�嵥 where �Ǽ�id=�Ǽ�id_IN AND ������� IS NULL AND ������Ŀid=������Ŀid_IN and ����id=����id_IN;
		ELSE
			Delete from �����Ŀ�嵥 where �Ǽ�id=�Ǽ�id_IN AND �������=�������_IN AND ������Ŀid=������Ŀid_IN;
		END IF;
	end if;

EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_�����Ŀ�嵥_DELETE;
/

----------------------------------------------------------------------------
---  INSERT   for   �����Ŀ�嵥
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_�����Ŀҽ��_INSERT(
	ID_IN IN �����Ŀ�嵥.ID%TYPE,
	����id_IN IN �����Ա����.����id%TYPE,
	ҽ��id_IN IN �����Ŀҽ��.ҽ��id%TYPE
)
IS
BEGIN
	INSERT INTO �����Ŀҽ��(�嵥id,����id,ҽ��id)
	SELECT ID,����id_IN,ҽ��id_IN FROM �����Ŀ�嵥 WHERE ID=ID_IN;
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_�����Ŀҽ��_INSERT;
/

----------------------------------------------------------------------------
---  INSERT   for   �����Ա����
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_�����Ա����_INSERT(
	ID_IN IN �����Ա����.ID%TYPE,
	�Ǽ�ID_IN IN �����Ա����.�Ǽ�ID%TYPE,
	����ID_IN IN �����Ա����.����ID%TYPE,
	�������_IN IN �����Ա����.�������%TYPE
)
IS
BEGIN
	Insert Into �����Ա����
		(ID,�Ǽ�ID,����ID,�������)
		VALUES
		(ID_IN,�Ǽ�ID_IN,����ID_IN,�������_IN);
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_�����Ա����_INSERT;
/
----------------------------------------------------------------------------
---  UPDATE   for   �����Ա����
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_�����Ա����_UPDATE(
	�Ǽ�ID_IN IN �����Ա����.�Ǽ�ID%TYPE,
	����ID_IN IN �����Ա����.����ID%TYPE,
	�������_IN IN �����Ա����.�������%TYPE,		
	ԭ����ID_IN IN NUMBER:=0
)
IS
BEGIN
	Update �����Ա����
		Set �Ǽ�ID=�Ǽ�ID_IN,
		    ����ID=����ID_IN,
		    �������=�������_IN		    
	WHERE �Ǽ�id=�Ǽ�id_IN AND ����ID=ԭ����ID_IN;
EXCEPTION
	WHEN OTHERS THEN
		zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_�����Ա����_UPDATE;
/
----------------------------------------------------------------------------
---  CLASS   for   �����Ա����
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_�����Ա����_CLASS(
	�Ǽ�ID_IN IN �����Ա����.�Ǽ�ID%TYPE,
	����ID_IN IN �����Ա����.����ID%TYPE,
	�������_IN IN �����Ա����.�������%TYPE
)
IS
BEGIN
	Update �����Ա����
		Set �������=�������_IN		    
	WHERE �Ǽ�id=�Ǽ�id_IN AND ����ID=����ID_IN;
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_�����Ա����_CLASS;
/
----------------------------------------------------------------------------
---  DELETE   for   �����Ա����
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_�����Ա����_DELETE(
	�Ǽ�id_IN IN �����Ա����.�Ǽ�ID%TYPE,
	����id_IN IN �����Ա����.����ID%TYPE:=0,
	ҽ������_IN NUMBER:=0
)
IS
	Cursor c_Advice is	
		SELECT A.ID,A.���id FROM ����ҽ����¼ A,���ǼǼ�¼ B WHERE A.����id=����id_IN AND A.ҽ��״̬ <>4 AND A.�Һŵ�=B.���� AND B.ID=�Ǽ�id_IN;
	
	r_Advice c_Advice%RowType;
	v_Count number(18);

	v_Have		number(1);
	Err_Custom	Exception;
	v_Error		Varchar2(255);
BEGIN
	
	--Ҫ��ҽ�����ϴ���
	IF ҽ������_IN=1 AND ����id_IN>0 THEN
		
		For r_Advice IN c_Advice Loop

			Update ����ҽ������ Set ִ��״̬=0,����id=NULL WHERE ҽ��ID=r_Advice.ID;

			Update ���˷��ü�¼ 
				Set ִ��״̬=0,ִ��ʱ��=NULL,ִ����=NULL
			Where �շ���� Not IN('5','6','7') 
				AND ҽ�����=r_Advice.ID
				And (��¼����,NO) IN(
						Select ��¼����,NO From ����ҽ������ Where ҽ��id=r_Advice.ID
						Union ALL
						Select ��¼����,NO From ����ҽ������ Where ҽ��id=r_Advice.ID);
						
		END LOOP;
		
		DELETE FROM �����Ա���� WHERE �Ǽ�id=�Ǽ�id_IN AND ����id=����id_IN;
		

		For r_Advice IN c_Advice Loop
			IF r_Advice.���id IS NULL THEN

				--�ж��Ƿ������Ч�ĸ���,���۵�\�շѵ�\���ʵ�
				v_Have:=0;
				BEGIN
					SELECT 1 INTO v_Have FROM ���˷��ü�¼ 
					WHERE  ��¼״̬ IN (0,1) 
						AND (ҽ�����,NO) IN 
							(
							SELECT ҽ��id,NO 
							FROM ����ҽ������ 
							WHERE ҽ��id IN (
									SELECT ID FROM ����ҽ����¼ 
									WHERE ID=r_Advice.ID OR ���id=r_Advice.ID
									)
							);

				EXCEPTION
					WHEN OTHERS THEN v_Have:=0;
				END;
				
				IF v_Have=1 THEN
					v_Error:='�������Ŀ�����ڸ���,���ȶԸ��ѽ���ɾ�������ϣ�';
				        Raise Err_Custom;
				END IF;
				
				DELETE FROM ����ҽ������ WHERE ҽ��id IN (
									SELECT ID FROM ����ҽ����¼ 
									WHERE ID=r_Advice.ID OR ���id=r_Advice.ID);	

				ZL_����ҽ����¼_����(r_Advice.ID);
			END IF;
		END LOOP;

		
	END IF;

	IF ����ID_IN=0 THEN
		Delete From �����Ա���� WHERE �Ǽ�id=�Ǽ�id_IN;
		Delete from �����Ŀҽ�� WHERE �嵥id in (SELECT ID FROM �����Ŀ�嵥 WHERE �Ǽ�id=�Ǽ�id_IN);
	ELSE
		Delete From �����Ա���� WHERE �Ǽ�id=�Ǽ�id_IN AND ����id=����ID_IN;
		Delete from �����Ŀҽ�� WHERE  ����id=����ID_IN AND �嵥id in (SELECT ID FROM �����Ŀ�嵥 WHERE �Ǽ�id=�Ǽ�id_IN);
	END IF;

	IF ҽ������_IN=1 THEN
		--����Ƴ���Ա��,û������Ա�����,�Զ��˵�δ��ʼ���״̬
		v_Count:=0;
		BEGIN 
			SELECT COUNT(1) INTO v_Count FROM �����Ա���� WHERE ��챨��=1 AND �Ǽ�id=�Ǽ�id_IN;
		EXCEPTION
			WHEN OTHERS THEN v_Count:=0;
		END;

		IF v_Count=0 THEN
			UPDATE ���ǼǼ�¼ SET ���״̬=2 WHERE ID=�Ǽ�id_IN;
		END IF;		
	END IF;
EXCEPTION
	When Err_Custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_�����Ա����_DELETE;
/

----------------------------------------------------------------------------
---  CANCEL   for   ���ǼǼ�¼
---	ȡ����쿪ʼ
----------------------------------------------------------------------------
CREATE OR REPLACE Procedure ZL_���ǼǼ�¼_Cancel(
	����_IN		varchar2
) IS
	Cursor c_Advice is
		SELECT ID,���id FROM ����ҽ����¼ WHERE ������Դ=4 AND �Һŵ�=����_IN AND ҽ��״̬<>4;
	Cursor c_Advice2 is
		SELECT ID FROM ����ҽ����¼ WHERE ���id IS NULL AND ������Դ=4 AND �Һŵ�=����_IN AND ҽ��״̬<>4;

	Cursor c_Report is
		SELECT ҽ��id,����id FROM ����ҽ������ WHERE ����id>0 AND ҽ��id IN (SELECT ID FROM ����ҽ����¼ WHERE ������Դ=4 AND �Һŵ�=����_IN AND ҽ��״̬<>4);

	Cursor c_Person is
		SELECT ID,��첡��id FROM �����Ա���� WHERE �Ǽ�id=(SELECT ID FROM ���ǼǼ�¼ WHERE ����=����_IN);

	r_Row c_Advice%RowType;

	v_Have number(1);
	Err_Custom	Exception;
	v_Error		Varchar2(255);
Begin
	
	DELETE FROM �����Ա���� WHERE �Ǽ�id=(SELECT ID FROM ���ǼǼ�¼ WHERE ����=����_IN);

	For r_Row IN c_Report Loop				
		UPDATE ����ҽ������ SET ����id=NULL WHERE ҽ��id=r_Row.ҽ��id;
	END LOOP;

	For r_Row IN c_Report Loop				
		DELETE FROM ���˲�����¼ WHERE ID=r_Row.����id;
	END LOOP;

	For r_Row IN c_Advice Loop

		Update ����ҽ������ Set ִ��״̬=0,����id=NULL WHERE ҽ��ID=r_Row.ID;

		Update ���˷��ü�¼ 
			Set ִ��״̬=0,ִ��ʱ��=NULL,ִ����=NULL
		Where �շ���� Not IN('5','6','7') 
			AND ҽ�����=r_Row.ID
			And (��¼����,NO) IN(
					Select ��¼����,NO From ����ҽ������ Where ҽ��id=r_Row.ID
					Union ALL
					Select ��¼����,NO From ����ҽ������ Where ҽ��id=r_Row.ID);

		--DELETE FROM ����ҽ������ WHERE ҽ��id=r_Row.ID;
	END LOOP;
	
	For r_Row IN c_Person Loop
		UPDATE �����Ա���� SET ���״̬=1,��첡��id=NULL,����ʱ��=NULL,��챨��=0 WHERE ID=r_Row.ID;
		DELETE FROM ���˲�����¼ WHERE ID=r_Row.��첡��id;
	end loop;

	UPDATE ���ǼǼ�¼ SET ���״̬=2 WHERE ����=����_IN;
	DELETE FROM �����Ŀҽ�� WHERE �嵥id IN (SELECT ID FROM �����Ŀ�嵥 WHERE �Ǽ�id=(SELECT ID FROM ���ǼǼ�¼ WHERE ����=����_IN));
	DELETE FROM �����Ŀ�嵥 WHERE ����id>0 AND �Ǽ�id=(SELECT ID FROM ���ǼǼ�¼ WHERE ����=����_IN);


	For r_Row IN c_Advice2 Loop		
		--�ж��Ƿ������Ч�ĸ���,���۵�\�շѵ�\���ʵ�
		v_Have:=0;
		BEGIN
			SELECT 1 INTO v_Have FROM ���˷��ü�¼ 
			WHERE  ��¼״̬ IN (0,1) 
				AND (ҽ�����,NO) IN 
					(
					SELECT ҽ��id,NO 
					FROM ����ҽ������ 
					WHERE ҽ��id IN (
							SELECT ID FROM ����ҽ����¼ 
							WHERE ID=r_Row.ID OR ���id=r_Row.ID
							)
					);

		EXCEPTION
			WHEN OTHERS THEN v_Have:=0;
		END;
		
		IF v_Have=1 THEN
			v_Error:='�������Ŀ�����ڸ���,���ȶԸ��ѽ���ɾ�������ϣ�';
			Raise Err_Custom;
		END IF;

		DELETE FROM ����ҽ������ WHERE ҽ��id IN (SELECT ID FROM ����ҽ����¼ WHERE ID=r_Row.ID OR ���id=r_Row.ID);

		ZL_����ҽ����¼_����(r_Row.ID);
	END LOOP;

Exception
    When Err_Custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_���ǼǼ�¼_Cancel;
/


CREATE OR REPLACE Procedure ZL_���ǼǼ�¼_ItemCancel(
	����_IN		varchar2,
	������Ŀid_IN		number:=0,
	�������_IN		varchar2:=NULL,
	����id_IN		number:=0
) IS
	Cursor c_Items is
		SELECT A.* FROM �����Ŀ�嵥 A,���ǼǼ�¼ B  WHERE A.�Ǽ�id=B.ID AND B.����=����_IN AND A.�������=�������_IN AND A.������Ŀid=������Ŀid_IN;

	Cursor c_ItemsMembers is
		SELECT A.* FROM �����Ŀ�嵥 A,���ǼǼ�¼ B  WHERE A.�Ǽ�id=B.ID AND B.����=����_IN AND A.������� IS NULL AND A.������Ŀid=������Ŀid_IN AND A.����id=����id_IN;

	Cursor c_ItemsPerson is
		SELECT A.* FROM �����Ŀ�嵥 A,���ǼǼ�¼ B  WHERE A.�Ǽ�id=B.ID AND B.����=����_IN AND A.������Ŀid=������Ŀid_IN AND A.����id=����id_IN;

	Cursor c_Advice(v_����id number,v_���� varchar2,v_���� number) is
		SELECT DECODE(v_����,1,���id,ID) AS ID FROM ����ҽ����¼ 
		WHERE ������Դ=4 
			AND �Һŵ�=v_���� 
			AND ҽ��״̬<>4 
			AND ������Ŀid=v_����id;

	Cursor c_AdvicePerson(v_����id number,v_���� varchar2,v_���� number) is
		SELECT DECODE(v_����,1,���id,ID) AS ID FROM ����ҽ����¼ 
		WHERE ������Դ=4 
			AND �Һŵ�=v_���� 
			AND ҽ��״̬<>4 
			AND ������Ŀid=v_����id 
			AND ����id=����id_IN;

	Cursor c_Advice2(v_ҽ��id number) is
		SELECT ID FROM ����ҽ����¼ WHERE ���id=v_ҽ��id or ID=v_ҽ��id;

	r_Advice c_Advice%RowType;
	r_AdvicePerson c_AdvicePerson%RowType;
	r_Item c_Items%RowType;
	r_ItemPerson c_ItemsPerson%RowType;

	r_Advice2 c_Advice2%RowType;
	
	v_������id number(18);
	v_Have number(1);
	v_Flag number(1);
	Err_Custom	Exception;
	v_Error		Varchar2(255);
Begin
	IF ����id_IN=0 THEN
		For r_Item IN c_Items Loop
			--if r_Item.�ɼ���ʽid IS NOT NULL AND r_Item.����걾 IS NOT NULL then
			v_Flag:=0;

			v_������id:=r_Item.������Ŀid;
			if r_Item.�ɼ���ʽid IS NOT NULL then
				v_Flag:=1;
			end if;
			
			For r_Advice IN c_Advice(v_������id,����_IN,v_Flag) Loop
				--�ҳ���ҽ��id
				For r_Advice2 IN c_Advice2(r_Advice.ID) Loop

					Update ����ҽ������ Set ִ��״̬=0,����id=NULL WHERE ҽ��ID=r_Advice2.ID;
					
					Update ���˷��ü�¼ 
							Set ִ��״̬=0,ִ��ʱ��=NULL,ִ����=NULL
					Where �շ���� Not IN('5','6','7') 
						AND ҽ�����=r_Advice2.ID
						And (��¼����,NO) IN(
							Select ��¼����,NO From ����ҽ������ Where ҽ��id=r_Advice2.ID
							Union ALL
							Select ��¼����,NO From ����ҽ������ Where ҽ��id=r_Advice2.ID);
				END LOOP;

				--�ж��Ƿ������Ч�ĸ���,���۵�\�շѵ�\���ʵ�
				v_Have:=0;
				BEGIN
					SELECT 1 INTO v_Have FROM ���˷��ü�¼ 
					WHERE  ��¼״̬ IN (0,1) 
						AND (ҽ�����,NO) IN 
							(
							SELECT ҽ��id,NO 
							FROM ����ҽ������ 
							WHERE ҽ��id IN (
									SELECT ID FROM ����ҽ����¼ 
									WHERE ID=r_Advice.ID OR ���id=r_Advice.ID
									)
							);

				EXCEPTION
					WHEN OTHERS THEN v_Have:=0;
				END;
				
				IF v_Have=1 THEN
					v_Error:='��ǰ�����Ŀ�����ڸ���,���ȶԸ��ѽ���ɾ�������ϣ�';
				        Raise Err_Custom;
				END IF;
				
				DELETE FROM ����ҽ������ WHERE ҽ��id IN (SELECT ID FROM ����ҽ����¼ WHERE ID=r_Advice.ID OR ���id=r_Advice.ID);

				ZL_����ҽ����¼_����(r_Advice.ID);
			END LOOP;
		end loop;
	ELSE
		For r_ItemPerson IN c_ItemsMembers Loop
			--if r_Item.�ɼ���ʽid IS NOT NULL AND r_Item.����걾 IS NOT NULL then
			
			v_Flag:=0;
			v_������id:=r_ItemPerson.������Ŀid;

			if r_ItemPerson.�ɼ���ʽid IS NOT NULL then
				v_Flag:=1;
			end if;
			
			For r_AdvicePerson IN c_AdvicePerson(v_������id,����_IN,v_Flag) Loop
				--�ҳ���ҽ��id
				For r_Advice2 IN c_Advice2(r_AdvicePerson.ID) Loop

					Update ����ҽ������ Set ִ��״̬=0,����id=NULL WHERE ҽ��ID=r_Advice2.ID;
					
					Update ���˷��ü�¼ 
							Set ִ��״̬=0,ִ��ʱ��=NULL,ִ����=NULL
					Where �շ���� Not IN('5','6','7') 
						AND ҽ�����=r_Advice2.ID
						And (��¼����,NO) IN(
							Select ��¼����,NO From ����ҽ������ Where ҽ��id=r_Advice2.ID
							Union ALL
							Select ��¼����,NO From ����ҽ������ Where ҽ��id=r_Advice2.ID);
				END LOOP;

				--�ж��Ƿ������Ч�ĸ���,���۵�\�շѵ�\���ʵ�
				v_Have:=0;
				BEGIN
					SELECT 1 INTO v_Have FROM ���˷��ü�¼ 
					WHERE  ��¼״̬ IN (0,1) 
						AND (ҽ�����,NO) IN 
							(
							SELECT ҽ��id,NO 
							FROM ����ҽ������ 
							WHERE ҽ��id IN (
									SELECT ID FROM ����ҽ����¼ 
									WHERE ID=r_AdvicePerson.ID OR ���id=r_AdvicePerson.ID
									)
							);

				EXCEPTION
					WHEN OTHERS THEN v_Have:=0;
				END;
				
				IF v_Have=1 THEN
					v_Error:='��ǰ�����Ŀ�����ڸ���,���ȶԸ��ѽ���ɾ�������ϣ�';
				        Raise Err_Custom;
				END IF;
				
				DELETE FROM ����ҽ������ WHERE ҽ��id IN (SELECT ID FROM ����ҽ����¼ WHERE ID=r_AdvicePerson.ID OR ���id=r_AdvicePerson.ID);
				ZL_����ҽ����¼_����(r_AdvicePerson.ID);
			END LOOP;
		end loop;		
	END IF;
Exception
    When Err_Custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_���ǼǼ�¼_ItemCancel;
/

CREATE OR REPLACE Procedure ZL_���ǼǼ�¼_Finish(
	����_IN		varchar2,
	����id_IN		number:=0
	--����id_INΪ0ʱ��ʾ�����ŵ����в���
) IS
	v_Count NUMBER(18);
Begin

	IF ����id_IN=0 THEN
		UPDATE �����Ա���� A SET A.���״̬=5,
					A.���ʱ��=SYSDATE,
					A.����=(SELECT ���� FROM ������Ϣ WHERE ����id=A.����id),
					A.�Ա�=(SELECT �Ա� FROM ������Ϣ WHERE ����id=A.����id),
					A.����=(SELECT ���� FROM ������Ϣ WHERE ����id=A.����id),
					A.����״��=(SELECT ����״�� FROM ������Ϣ WHERE ����id=A.����id),
					A.��ϵ�绰=(SELECT ��ϵ�绰 FROM ������Ϣ WHERE ����id=A.����id),
					A.��ϵ��ַ=(SELECT ��ϵ��ַ FROM ������Ϣ WHERE ����id=A.����id)					
		WHERE A.�Ǽ�id=(SELECT ID FROM ���ǼǼ�¼ WHERE ����=����_IN);
		UPDATE ���ǼǼ�¼ SET ���״̬=5 WHERE ����=����_IN;
	ELSE
		UPDATE �����Ա���� SET ���״̬=5,
					���ʱ��=SYSDATE,
					����=(SELECT ���� FROM ������Ϣ WHERE ����id=����id_IN),
					�Ա�=(SELECT �Ա� FROM ������Ϣ WHERE ����id=����id_IN),
					����=(SELECT ���� FROM ������Ϣ WHERE ����id=����id_IN),
					����״��=(SELECT ����״�� FROM ������Ϣ WHERE ����id=����id_IN),
					��ϵ�绰=(SELECT ��ϵ�绰 FROM ������Ϣ WHERE ����id=����id_IN),
					��ϵ��ַ=(SELECT ��ϵ��ַ FROM ������Ϣ WHERE ����id=����id_IN)
		WHERE ����id=����id_IN 
			AND �Ǽ�id=(SELECT ID FROM ���ǼǼ�¼ WHERE ����=����_IN);
		
		v_Count:=0;
		BEGIN
			SELECT NVL(COUNT(1),0) INTO v_Count FROM �����Ա���� WHERE ���״̬<5 AND �Ǽ�id=(SELECT ID FROM ���ǼǼ�¼ WHERE ����=����_IN);
		EXCEPTION
			WHEN OTHERS THEN v_Count:=0;
		END;

		IF v_Count<=0 THEN
			UPDATE ���ǼǼ�¼ SET ���״̬=5,���ʱ��=SYSDATE WHERE ����=����_IN;		
		END IF;
	END IF;
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_���ǼǼ�¼_Finish;
/

CREATE OR REPLACE Procedure ZL_���ǼǼ�¼_CancelFinish(
	����_IN		varchar2,
	����id_IN		number:=0
	--����id_INΪ0ʱ��ʾ�����ŵ����в���
) IS
	v_No varchar2(30);
	v_Temp			Varchar2(255);
	v_��Ա���		��Ա��.���%Type;
	v_��Ա����		��Ա��.����%Type;
Begin
	
	--��ǰ������Ա
	v_Temp:=zl_Identity;
	v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
	v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
	v_��Ա���:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
	v_��Ա����:=Substr(v_Temp,Instr(v_Temp,',')+1);

	IF ����id_IN=0 THEN
		UPDATE �����Ա���� SET ���״̬=4,���ʱ��=NULL WHERE �Ǽ�id=(SELECT ID FROM ���ǼǼ�¼ WHERE ����=����_IN);

	ELSE
		UPDATE �����Ա���� SET ���״̬=4,���ʱ��=NULL WHERE ����id=����id_IN AND �Ǽ�id=(SELECT ID FROM ���ǼǼ�¼ WHERE ����=����_IN);
	END IF;

	UPDATE ���ǼǼ�¼ SET ���״̬=4,���ʱ��=NULL WHERE ����=����_IN;		

	--ȡ����������
	BEGIN
		zl_���˽��ʼ�¼_Delete(v_No,v_��Ա���,v_��Ա����,0);
	EXCEPTION
		WHEN OTHERS THEN v_No:='';
	END;
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_���ǼǼ�¼_CancelFinish;
/

CREATE OR REPLACE Procedure ZL_���ǼǼ�¼_������д(
	����id_IN		���˲���������.����id%TYPE,
	������id_IN		���˲���������.������id%TYPE,
	��������_IN		���˲���������.��������%TYPE
) IS
Begin
	UPDATE ���˲��������� SET ��������=��������_IN WHERE ����id=����id_IN AND ������id=������id_IN;	
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_���ǼǼ�¼_������д;
/

CREATE OR REPLACE Procedure ZL_�����Ա����_�ܽ�(
	�Ǽ�id_IN		�����Ա����.�Ǽ�id%TYPE,
	����id_IN		�����Ա����.����id%TYPE,
	��첡��id_IN		�����Ա����.��첡��id%TYPE
) IS
Begin
	UPDATE �����Ա���� SET ��첡��id=��첡��id_IN WHERE �Ǽ�id=�Ǽ�id_IN AND ����id=����id_IN;	
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_�����Ա����_�ܽ�;
/

CREATE OR REPLACE Procedure ZL_�����Ա����_����(
	�Ǽ�id_IN		�����Ա����.�Ǽ�id%TYPE,
	����id_IN		�����Ա����.����id%TYPE,
	����ʱ��_IN		�����Ա����.����ʱ��%TYPE
) IS
Begin
	UPDATE �����Ա���� SET ����ʱ��=����ʱ��_IN WHERE �Ǽ�id=�Ǽ�id_IN AND ����id=����id_IN;	
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_�����Ա����_����;
/

CREATE OR REPLACE PROCEDURE ZL_�����Ա����_REFRESH(
	�Ǽ�id_IN IN �����Ա����.�Ǽ�id%TYPE,
	����id_IN IN �����Ա����.����id%TYPE:=0
)
IS
	Cursor c_PersonAll is	
		SELECT ����id FROM �����Ա���� WHERE �Ǽ�id=�Ǽ�id_IN;

	Cursor c_Person is	
		SELECT ����id FROM �����Ա���� WHERE �Ǽ�id=�Ǽ�id_IN AND ����id=����id_IN;

	Cursor c_List(v_����id number) is	
		SELECT DISTINCT ִ�п���id 
		FROM (       
			SELECT ִ�п���id 
			FROM �����Ŀ�嵥 A,�����Ա���� B
			WHERE B.�Ǽ�id=�Ǽ�id_IN 
			      AND B.����ID=v_����id
			      AND A.�Ǽ�ID=B.�Ǽ�ID
			      AND A.�������=B.�������
			      AND A.����id IS NULL   
			UNION ALL 
			SELECT ִ�п���id 
			FROM �����Ŀ�嵥 A
			WHERE A.�Ǽ�id=�Ǽ�id_IN
			      AND A.����ID=v_����id      
		     );

	v_Have NUMBER(1);

BEGIN
	IF ����id_IN=0 THEN
		For r_Person IN c_PersonAll Loop

			For r_Row IN c_List(r_Person.����id) Loop
				
				v_Have:=0;
				BEGIN
					SELECT 1 INTO v_Have FROM �����Ա���� WHERE �Ǽ�id=�Ǽ�id_IN AND ����id=r_Person.����id AND ����id=r_Row.ִ�п���id;
				EXCEPTION
					WHEN OTHERS THEN v_Have:=0;
				END;

				IF v_Have=0 THEN
					--û��,������
					INSERT INTO �����Ա����(�Ǽ�id,����id,����id,����id) VALUES (�Ǽ�id_IN,r_Person.����id,r_Row.ִ�п���id,NULL);
				END IF;

			END LOOP;
			
			DELETE FROM �����Ա���� 
			WHERE �Ǽ�id=�Ǽ�id_IN 
				AND ����id=r_Person.����id 
				AND ����id NOT IN (
						SELECT DISTINCT ִ�п���id 
						FROM (       
							SELECT ִ�п���id 
							FROM �����Ŀ�嵥 A,�����Ա���� B
							WHERE B.�Ǽ�id=�Ǽ�id_IN 
							      AND B.����ID=r_Person.����id
							      AND A.�Ǽ�ID=B.�Ǽ�ID
							      AND A.�������=B.�������
							      AND A.����id IS NULL   
							UNION ALL 
							SELECT ִ�п���id 
							FROM �����Ŀ�嵥 A
							WHERE A.�Ǽ�id=�Ǽ�id_IN
							      AND A.����ID=r_Person.����id
						));
		end loop;
	ELSE
		For r_Row IN c_List(����id_IN) Loop
			
			v_Have:=0;
			BEGIN
				SELECT 1 INTO v_Have FROM �����Ա���� WHERE �Ǽ�id=�Ǽ�id_IN AND ����id=����id_IN AND ����id=r_Row.ִ�п���id;
			EXCEPTION
				WHEN OTHERS THEN v_Have:=0;
			END;

			IF v_Have=0 THEN
				--û��,������
				INSERT INTO �����Ա����(�Ǽ�id,����id,����id,����id) VALUES (�Ǽ�id_IN,����id_IN,r_Row.ִ�п���id,NULL);
			END IF;

		END LOOP;
		
		DELETE FROM �����Ա���� 
		WHERE �Ǽ�id=�Ǽ�id_IN 
			AND ����id=����id_IN
			AND ����id NOT IN (
					SELECT DISTINCT ִ�п���id 
					FROM (       
						SELECT ִ�п���id 
						FROM �����Ŀ�嵥 A,�����Ա���� B
						WHERE B.�Ǽ�id=�Ǽ�id_IN 
						      AND B.����ID=����id_IN
						      AND A.�Ǽ�ID=B.�Ǽ�ID
						      AND A.�������=B.�������
						      AND A.����id IS NULL   
						UNION ALL 
						SELECT ִ�п���id 
						FROM �����Ŀ�嵥 A
						WHERE A.�Ǽ�id=�Ǽ�id_IN
						      AND A.����ID=����id_IN
					));
	END IF;
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_�����Ա����_REFRESH;
/

CREATE OR REPLACE PROCEDURE ZL_�����Ա����_UPDATE(
	�Ǽ�id_IN IN �����Ա����.�Ǽ�id%TYPE,
	����id_IN IN �����Ա����.����id%TYPE,
	����id_IN IN �����Ա����.����id%TYPE,
	����id_IN IN �����Ա����.����id%TYPE
)
IS
BEGIN
	UPDATE �����Ա���� SET ����id=����id_IN WHERE �Ǽ�id=�Ǽ�id_IN AND ����id=����id_IN AND ����id=����id_IN;
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_�����Ա����_UPDATE;
/

CREATE OR REPLACE Procedure ZL_�����Ա����_����(
	�Ǽ�id_IN		�����Ա����.�Ǽ�id%TYPE,
	����id_IN		�����Ա����.����id%TYPE,
	��챨��_IN		�����Ա����.��챨��%TYPE
) IS
Begin
	UPDATE �����Ա���� SET ��챨��=��챨��_IN WHERE �Ǽ�id=�Ǽ�id_IN AND ����id=����id_IN;	
	IF ��챨��_IN=1 THEN
		ZL_�����Ա����_REFRESH(�Ǽ�id_IN,����id_IN);
	END IF;
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_�����Ա����_����;
/

CREATE OR REPLACE Procedure ZL_�������¼_INSERT(
	ID_IN		IN	�������¼.ID%TYPE,
	��Լ��λid_IN	IN	�������¼.��Լ��λid%TYPE,
	����id_IN	IN	�������¼.����id%TYPE,
	������_IN	IN	�������¼.������%TYPE,
	���㲿��id_IN	IN 	�������¼.���㲿��id%TYPE
) IS
Begin
	INSERT INTO �������¼(ID,��¼״̬,��Լ��λid,����id,������,���㲿��id)
	VALUES (ID_IN,1,��Լ��λid_IN,����id_IN,������_IN,���㲿��id_IN);
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_�������¼_INSERT;
/

CREATE OR REPLACE Procedure ZL_�������嵥_INSERT(
	����id_IN	IN	�������嵥.����id%TYPE,
	�Ǽ�id_IN	IN	�������嵥.�Ǽ�id%TYPE
) IS
Begin
	INSERT INTO �������嵥(����id,�Ǽ�id)	VALUES (����id_IN,�Ǽ�id_IN);
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_�������嵥_INSERT;
/

CREATE OR REPLACE Procedure zl_�������¼_Cancel(
	����id_IN	IN	���˽��ʼ�¼.ID%TYPE
) IS
	Cursor c_Items is
		SELECT A.* FROM ���˽��ʼ�¼ A WHERE A.ID=����id_IN;

	v_Temp			Varchar2(255);
	v_��Ա���		��Ա��.���%Type;
	v_��Ա����		��Ա��.����%Type;
Begin

	--��ǰ������Ա
	v_Temp:=zl_Identity;
	v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
	v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
	v_��Ա���:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
	v_��Ա����:=Substr(v_Temp,Instr(v_Temp,',')+1);

	For r_Row IN c_Items Loop
		zl_���˽��ʼ�¼_Delete(r_Row.No,v_��Ա���,v_��Ա����,0);
	end loop;

	UPDATE �������¼ SET ��¼״̬=2 WHERE ����id=����id_IN;
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_�������¼_Cancel;
/

--����ZL1_BILL_1861/���Ա��Ŀ�嵥
Insert Into zlReports(ID,���,����,˵��,����,W,H,ֽ��,ֽ��,��ֽ,��ӡ��,��ֽ̬��,Ʊ��,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��) Values(zlReports_ID.NextVal,'ZL1_BILL_1861','���Ա��Ŀ�嵥','��Ŀ�嵥','Zn!t_jgnq1<S~aimD0[_',11904,16832,9,1,15,NULL,0,1,100,1861,'��Ŀ�嵥',Sysdate,Sysdate);
Insert Into zlRPTFMTs(����ID,���,˵��,ͼ��) Values(zlReports_ID.CurrVal,1,'���Ա��Ŀ�嵥1',0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ3',2,NULL,0,'�����1',11,'����:[�����Ŀ�嵥_����.����]',NULL,450,1425,2610,180,0,0,1,'����',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ1',2,NULL,0,'�����1',12,'�����Ŀ�嵥',NULL,4478,660,2700,435,0,0,1,'����',22,1,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ2',2,NULL,0,'�����1',13,'��쵥:[�����Ŀ�嵥_����.����]',NULL,8235,1665,2970,180,0,2,1,'����',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'�����1',4,NULL,0,NULL,0,'�����Ŀ�嵥_����',NULL,450,1950,10755,8460,255,0,0,'����',9,0,0,0,0,16777215,1,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-1,0,NULL,NULL,'[�����Ŀ�嵥_����.���]','4^255^���',0,0,1140,0,255,0,0,'����',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-2,1,NULL,NULL,'[�����Ŀ�嵥_����.����]','4^255^����',0,0,6345,0,255,0,0,'����',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-3,2,NULL,NULL,'[�����Ŀ�嵥_����.������]','4^255^������',0,0,2625,0,255,0,0,'����',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ4',2,NULL,0,'�����1',11,'����:[�����Ŀ�嵥_����.����]',NULL,450,1670,2610,180,0,0,1,'����',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'�����Ŀ�嵥_����','���,200|����,200|������,200|����,200|����,200|����,200',USER||'.�����Ŀ�嵥,'||USER||'.������ĿĿ¼,'||USER||'.���ű�,'||USER||'.������Ŀ���,'||USER||'.�����Ŀҽ��,'||USER||'.���ǼǼ�¼,'||USER||'.��Լ��λ,'||USER||'.������Ϣ',0);
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,1,'Select D.���� AS ���,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,2,'       B.����,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,3,'       C.���� AS ������,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,4,'       F.����,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,5,'       H.����,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,6,'       G.���� AS ����');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,7,'From �����Ŀ�嵥 A,������ĿĿ¼ B,���ű� C,������Ŀ��� D,�����Ŀҽ�� E,���ǼǼ�¼ F,��Լ��λ G,������Ϣ H');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,8,'WHERE A.������ĿID=B.ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,9,'      AND C.ID=A.ִ�п���ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,10,'      AND B.���=D.����');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,11,'      AND E.�嵥ID=A.ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,12,'      AND A.�Ǽ�ID=F.ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,13,'      AND E.����ID=H.����ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,14,'      AND F.��Լ��λid=G.ID(+)');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,15,'      AND A.�Ǽ�ID=[0]');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,16,'      AND E.����ID=[1]');
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,NULL,0,'�Ǽ�ID',1,NULL,0,NULL,NULL,NULL,NULL,NULL,NULL);
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,NULL,1,'����ID',1,NULL,0,NULL,NULL,NULL,NULL,NULL,NULL);

--����ZL1_BILL_1861_2/��챨����
Insert Into zlReports(ID,���,����,˵��,����,W,H,ֽ��,ֽ��,��ֽ,��ӡ��,��ֽ̬��,Ʊ��,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��) Values(zlReports_ID.NextVal,'ZL1_BILL_1861_2','��챨����','�������ӡ','Zn:kA}6x|;0Tm=|sW*Q]',11904,16832,9,1,15,NULL,0,1,100,1861,'�������ӡ',Sysdate,Sysdate);
Insert Into zlRPTFMTs(����ID,���,˵��,ͼ��) Values(zlReports_ID.CurrVal,1,'��챨����1',0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ6',2,NULL,0,NULL,0,'��[ҳ��]ҳ ��[ҳ��]ҳ',NULL,345,15960,1965,180,0,1,1,'����',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ5',2,NULL,0,'�����1',11,'����:[�����Ա����_����.����]',NULL,360,1220,2610,180,0,0,1,'����',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ2',2,NULL,0,'�����1',11,'����:[�����Ա����_����.����]',NULL,360,1505,2610,180,0,0,1,'����',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ1',2,NULL,0,'�����1',12,'��챨�浥',NULL,4860,675,2250,435,0,0,1,'����',22,1,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ4',2,NULL,0,'�����1',13,'���ʱ��:[�����Ա����_����.���ʱ��]',NULL,8280,1485,3330,180,0,1,1,'����',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'�����1',4,NULL,0,NULL,0,'���˲���������_����',NULL,360,1785,11250,12570,345,0,0,'����',9,0,0,0,0,16777215,1,NULL,NULL,NULL,1,8421504,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-1,0,NULL,NULL,'[���˲���������_����.��Ŀ]','4^345^��Ŀ',0,0,4455,0,255,0,0,'����',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-2,1,NULL,NULL,'[���˲���������_����.���]','4^345^���',0,0,4110,0,255,0,0,'����',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-3,2,NULL,NULL,'[���˲���������_����.�ο�]','4^345^�ο�',0,0,1860,0,255,0,0,'����',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-4,3,NULL,NULL,'[���˲���������_����.��ʾ]','4^345^��ʾ',0,0,765,0,0,0,0,'����',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'�����2',4,NULL,0,'�����1',1,'����ܼ�_����',NULL,360,14355,11250,1545,570,0,1,'����',9,0,0,0,0,16777215,1,NULL,NULL,NULL,1,8421504,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-1,0,NULL,NULL,'[����ܼ�_����.��Ŀ]','1^345^�ܼ�',0,0,1260,0,255,0,0,'����',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-2,1,NULL,NULL,'[����ܼ�_����.���]','1^345^�ܼ�',0,0,9855,0,255,0,0,'����',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'�����Ա����_����','����,200|����,200|���ʱ��,200',USER||'.�����Ա����,'||USER||'.������Ϣ,'||USER||'.���ǼǼ�¼,'||USER||'.��Լ��λ',0);
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,1,'SELECT 	B.����, ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,2,'	D.���� AS ����,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,3,'	TO_CHAR(C.���ʱ��,'||CHR(39)||'yyyy-mm-dd'||CHR(39)||') AS ���ʱ��');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,4,'FROM 	�����Ա���� A,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,5,'	������Ϣ B,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,6,'	���ǼǼ�¼ C,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,7,'	��Լ��λ D');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,8,'WHERE 	A.����id=B.����id');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,9,'	AND C.ID=A.�Ǽ�id');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,10,'	AND C.��Լ��λid=D.ID(+)');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,11,'	AND C.ID=[0]');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,12,'	AND A.����id=[1]');
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,NULL,0,'�Ǽ�ID',1,NULL,0,NULL,NULL,NULL,NULL,NULL,NULL);
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,NULL,1,'����ID',1,NULL,0,NULL,NULL,NULL,NULL,NULL,NULL);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'����ܼ�_����','��Ŀ,200|���,200',USER||'.���˲���������,'||USER||'.����������Ŀ,'||USER||'.���˲����ı���,'||USER||'.�����Ա����,'||USER||'.���˲�������',0);
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,1,'SELECT 	'||CHR(39)||'    '||CHR(39)||'||��Ŀ AS ��Ŀ,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,2,'	���');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,3,'FROM (');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,4,'select ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,5,'       X.�������,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,6,'       X1.�����,              ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,7,'       X1.��Ŀ,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,8,'       X1.���  ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,9,'from      ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,10,'     �����Ա���� A,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,11,'     ���˲������� X,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,12,'     (');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,13,'     select  ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,14,'             A.����ID,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,15,'             A.�ؼ��� AS �����,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,16,'             B.������ AS ��Ŀ,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,17,'             DECODE(A.��������,NULL,NULL,A.��������||'||CHR(39)||' '||CHR(39)||'||B.��λ) AS ���');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,18,'      from ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,19,'        ���˲��������� A,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,20,'        ����������Ŀ B');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,21,'      where A.������id=B.ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,22,'            and ������id>0');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,23,'      ) X1');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,24,'where X.������¼id=A.��첡��ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,25,'      AND X.ID=X1.����id');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,26,'      AND A.����ID=[1]');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,27,'      AND A.�Ǽ�ID=[0]');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,28,'      AND X.Ԫ������=2       ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,29,NULL);
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,30,'union all ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,31,'      ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,32,'select     ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,33,'       X.�������,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,34,'       X1.�����,        ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,35,'       X.�����ı� AS ��Ŀ,       ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,36,'       X1.���  ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,37,'from ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,38,'     �����Ա���� A,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,39,'     ���˲������� X,     ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,40,'      (');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,41,'      select ����id,0 AS �����,'||CHR(39)||''||CHR(39)||' AS ��Ŀ,���� AS ��� from ���˲����ı���   ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,42,'      ) X1');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,43,'where X.������¼id=A.��첡��ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,44,'      AND X.ID=X1.����id');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,45,'      AND A.����ID=[1]      ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,46,'      AND A.�Ǽ�ID=[0]');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,47,'      AND X.Ԫ������ in (4,-5)          ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,48,') ORDER BY  �������,����� ');
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,NULL,0,'�Ǽ�ID',1,NULL,0,NULL,NULL,NULL,NULL,NULL,NULL);
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,NULL,1,'����ID',1,NULL,0,NULL,NULL,NULL,NULL,NULL,NULL);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'���˲���������_����','��Ŀ,200|���,200|�ο�,200|��ʾ,200|����1,200|����2,200|����3,139',USER||'.�����Ŀҽ��,'||USER||'.����ҽ����¼,'||USER||'.����ҽ������,'||USER||'.���˲�����¼,'||USER||'.�����Ŀ�嵥,'||USER||'.���˲���������,'||USER||'.����������Ŀ,'||USER||'.���˲����ı���,'||USER||'.���˲�������,'||USER||'.���ű�,'||USER||'.������ĿĿ¼,'||USER||'.�����Ա����',0);
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,1,'SELECT * FROM (');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,2,'  SELECT ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,3,'         '||CHR(39)||'        '||CHR(39)||'||��Ŀ AS ��Ŀ,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,4,'         ���||DECODE(��־,NULL,'||CHR(39)||''||CHR(39)||',DECODE(SUBSTR(��־,3,100),'||CHR(39)||'����'||CHR(39)||','||CHR(39)||''||CHR(39)||','||CHR(39)||'�쳣'||CHR(39)||','||CHR(39)||'(+)'||CHR(39)||','||CHR(39)||'ƫ��'||CHR(39)||','||CHR(39)||'��'||CHR(39)||','||CHR(39)||'ƫ��'||CHR(39)||','||CHR(39)||'��'||CHR(39)||')) AS ���,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,5,'         �ο�,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,6,'         DECODE(��־,NULL,'||CHR(39)||''||CHR(39)||',SUBSTR(��־,3,100)) AS ��ʾ,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,7,'         ������ AS ����1,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,8,'         �����Ŀ AS ����2,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,9,'         3 AS ����3       ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,10,'  FROM (');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,11,'  SELECT ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,12,'         U.���� AS ������,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,13,'         T.���� AS �����Ŀ,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,14,'         R.��Ŀ,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,15,'         R.���,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,16,'         DECODE(SIGN(INSTR(R.��־�ο�,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')),1,SUBSTR(R.��־�ο�,1,INSTR(R.��־�ο�,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')-1),'||CHR(39)||''||CHR(39)||') AS ��־,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,17,'         DECODE(SIGN(INSTR(R.��־�ο�,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')),1,SUBSTR(R.��־�ο�,INSTR(R.��־�ο�,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')+1,1000),'||CHR(39)||''||CHR(39)||') AS �ο�');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,18,'  FROM (');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,19,'  select ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,20,'         A.ִ�в���ID,       ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,21,'         A.�����Ŀid,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,22,'         A.ID,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,23,'         X.�������,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,24,'         X1.�����,              ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,25,'         X1.��Ŀ,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,26,'         DECODE(SIGN(INSTR(X1.���,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')),1,SUBSTR(X1.���,1,INSTR(X1.���,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')-1),X1.���) AS ���,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,27,'         DECODE(SIGN(INSTR(X1.���,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')),1,SUBSTR(X1.���,INSTR(X1.���,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')+1,1000),'||CHR(39)||''||CHR(39)||') AS ��־�ο�       ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,28,'  from ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,29,'       (');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,30,'       select DISTINCT A1.ҽ��ID,A3.ִ�в���ID,A4.ID,A5.������Ŀid AS �����Ŀid');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,31,'        from �����Ŀҽ�� A1,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,32,'             ����ҽ����¼ A2,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,33,'             ����ҽ������ A3,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,34,'             ���˲�����¼ A4,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,35,'             �����Ŀ�嵥 A5');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,36,'        where A1.����id=[1]');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,37,'      AND A5.�Ǽ�id=[0]');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,38,'              AND (A1.ҽ��ID=A2.ID OR A1.ҽ��ID=A2.���id)');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,39,'              AND A3.ҽ��ID=A2.ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,40,'              AND A4.ID=A3.����ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,41,'              AND A5.ID=A1.�嵥ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,42,'       ) A,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,43,'       ���˲������� X,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,44,'       (');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,45,'       select  ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,46,'               A.����ID,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,47,'               A.�ؼ��� AS �����,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,48,'               B.������ AS ��Ŀ,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,49,'               DECODE(A.��������,NULL,NULL,A.��������||'||CHR(39)||' '||CHR(39)||'||B.��λ) AS ���');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,50,'        from ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,51,'          ���˲��������� A,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,52,'          ����������Ŀ B');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,53,'        where A.������id=B.ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,54,'              and ������id>0');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,55,'        ) X1');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,56,'  where X.������¼id=A.ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,57,'        AND X.ID=X1.����id');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,58,'  union all     ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,59,'  select ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,60,'         A.ִ�в���ID,       ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,61,'         A.�����Ŀid,      ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,62,'         A.ID,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,63,'         X.�������,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,64,'         X1.�����,        ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,65,'         X.�����ı� AS ��Ŀ,       ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,66,'         X1.���,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,67,'         '||CHR(39)||''||CHR(39)||' AS ��־�ο�   ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,68,'  from ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,69,'       (');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,70,'       select DISTINCT A1.ҽ��ID,A3.ִ�в���ID,A4.ID,A5.������Ŀid AS �����Ŀid');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,71,'        from �����Ŀҽ�� A1,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,72,'             ����ҽ����¼ A2,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,73,'             ����ҽ������ A3,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,74,'             ���˲�����¼ A4,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,75,'             �����Ŀ�嵥 A5');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,76,'        where A1.����id=[1]');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,77,'      AND A5.�Ǽ�id=[0]');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,78,'              AND (A1.ҽ��ID=A2.ID OR A1.ҽ��ID=A2.���id)');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,79,'              AND A3.ҽ��ID=A2.ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,80,'              AND A4.ID=A3.����ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,81,'              AND A5.ID=A1.�嵥ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,82,'       ) A,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,83,'       ���˲������� X,     ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,84,'        (');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,85,'        select ����id,0 AS �����,'||CHR(39)||''||CHR(39)||' AS ��Ŀ,���� AS ��� from ���˲����ı���   ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,86,'        ) X1');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,87,'  where X.������¼id=A.ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,88,'        AND X.ID=X1.����id');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,89,'        AND X.Ԫ������ IN (0,4,-5)');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,90,'        AND X.Ԫ�ر���<>'||CHR(39)||'000009'||CHR(39));
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,91,'  ) R,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,92,'  ���ű� U,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,93,'  ������ĿĿ¼ T');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,94,'  WHERE R.ִ�в���id=U.ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,95,'        AND R.�����Ŀid=T.ID)              ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,96,NULL);
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,97,'UNION ALL ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,98,'	');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,99,'	SELECT ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,100,'	       '||CHR(39)||'    '||CHR(39)||'||T.���� AS ��Ŀ,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,101,'	       '||CHR(39)||'���ʱ��:'||CHR(39)||'||TO_CHAR(R.��д����,'||CHR(39)||'yyyy-mm-dd hh24:mi'||CHR(39)||') AS ���,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,102,'	       '||CHR(39)||'���ҽ��:'||CHR(39)||'||R.��д�� AS �ο�,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,103,'		'||CHR(39)||''||CHR(39)||' AS ��ʾ,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,104,'	       U.���� AS ����1,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,105,'	       T.���� AS ����2,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,106,'	       2 AS ����3       ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,107,'	FROM (');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,108,'	SELECT DISTINCT');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,109,'	       A.ִ�в���ID,       ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,110,'	       A.�����Ŀid,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,111,'	       A.ID,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,112,'	       A.��д��,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,113,'	       A.��д����                 ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,114,'	from ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,115,'	     (');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,116,'	     SELECT DISTINCT A1.ҽ��ID,A3.ִ�в���ID,A4.ID,A5.������Ŀid AS �����Ŀid,A4.��д��,A4.��д����');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,117,'	      from �����Ŀҽ�� A1,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,118,'	           ����ҽ����¼ A2,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,119,'	           ����ҽ������ A3,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,120,'	           ���˲�����¼ A4,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,121,'	           �����Ŀ�嵥 A5');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,122,'	      where A1.����id=[1]');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,123,'		      AND A5.�Ǽ�id=[0]');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,124,'	            AND (A1.ҽ��ID=A2.ID OR A1.ҽ��ID=A2.���id)');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,125,'	            AND A3.ҽ��ID=A2.ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,126,'	            AND A4.ID=A3.����ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,127,'	            AND A5.ID=A1.�嵥ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,128,'	     ) A,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,129,'	     ���˲������� X');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,130,'	WHERE X.������¼id=A.ID     ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,131,') R,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,132,'���ű� U,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,133,'������ĿĿ¼ T');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,134,'WHERE R.ִ�в���id=U.ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,135,'      AND R.�����Ŀid=T.ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,136,'union all ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,137,'SELECT ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,138,'       ������ AS ��Ŀ,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,139,'       '||CHR(39)||''||CHR(39)||' AS ���,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,140,'       '||CHR(39)||''||CHR(39)||' AS �ο�,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,141,'	'||CHR(39)||''||CHR(39)||' AS ��ʾ,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,142,'       ������ AS ����1,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,143,'       '||CHR(39)||' '||CHR(39)||' AS ����2,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,144,'       1 AS ����3       ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,145,'FROM (');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,146,'     select DISTINCT U.���� AS ������');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,147,'      from �����Ŀҽ�� A1,           ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,148,'           �����Ŀ�嵥 A5,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,149,'           ���ű� U');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,150,'      where A1.����id=[1]');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,151,'	      AND A5.�Ǽ�id=[0]');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,152,'            AND A5.ִ�п���ID=U.ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,153,'            AND A5.ID=A1.�嵥ID     ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,154,'     ) R');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,155,NULL);
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,156,'union all ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,157,'SELECT ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,158,'       '||CHR(39)||'    С��'||CHR(39)||' AS ��Ŀ,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,159,'       DECODE(R.��д����,NULL,'||CHR(39)||''||CHR(39)||','||CHR(39)||'С��ʱ��:'||CHR(39)||'||TO_CHAR(R.��д����,'||CHR(39)||'yyyy-mm-dd hh24:mi'||CHR(39)||')) AS ���,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,160,'       DECODE(R.��д��,NULL,'||CHR(39)||''||CHR(39)||','||CHR(39)||'С��ҽ��:'||CHR(39)||'||R.��д��) AS �ο�,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,161,'	'||CHR(39)||''||CHR(39)||' as ��ʾ,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,162,'       ������ AS ����1,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,163,'       '||CHR(39)||'��������������������������������'||CHR(39)||' AS ����2,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,164,'       4 AS ����3       ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,165,'FROM (');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,166,'     select DISTINCT U.���� AS ������,A4.��д��,A4.��д����');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,167,'      from �����Ա���� A1,           ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,168,'           ���˲�����¼ A4,		');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,169,'           ���ű� U');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,170,'      where A1.����id=[1]');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,171,'	      AND A1.�Ǽ�id=[0]');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,172,'            AND A1.����ID=U.ID ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,173,'	    AND A1.����id=A4.ID(+)');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,174,'     ) R         ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,175,'UNION ALL ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,176,'  SELECT ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,177,'         '||CHR(39)||'        '||CHR(39)||'||��Ŀ AS ��Ŀ,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,178,'         ���||DECODE(��־,NULL,'||CHR(39)||''||CHR(39)||',DECODE(SUBSTR(��־,3,100),'||CHR(39)||'����'||CHR(39)||','||CHR(39)||''||CHR(39)||','||CHR(39)||'�쳣'||CHR(39)||','||CHR(39)||'(+)'||CHR(39)||','||CHR(39)||'ƫ��'||CHR(39)||','||CHR(39)||'��'||CHR(39)||','||CHR(39)||'ƫ��'||CHR(39)||','||CHR(39)||'��'||CHR(39)||')) AS ���,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,179,'         �ο�,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,180,'	'||CHR(39)||''||CHR(39)||' as ��ʾ,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,181,'         ������ AS ����1,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,182,'         '||CHR(39)||'��������������������������������'||CHR(39)||' AS ����2,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,183,'         5 AS ����3       ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,184,'  FROM (');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,185,'  SELECT ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,186,'         U.���� AS ������,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,187,'         R.��Ŀ,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,188,'         R.���,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,189,'         DECODE(SIGN(INSTR(R.��־�ο�,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')),1,SUBSTR(R.��־�ο�,1,INSTR(R.��־�ο�,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')-1),'||CHR(39)||''||CHR(39)||') AS ��־,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,190,'         DECODE(SIGN(INSTR(R.��־�ο�,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')),1,SUBSTR(R.��־�ο�,INSTR(R.��־�ο�,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')+1,1000),'||CHR(39)||''||CHR(39)||') AS �ο�');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,191,'  FROM (');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,192,'  select ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,193,'         A.ִ�в���ID,       ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,194,'         A.ID,              ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,195,'         X1.��Ŀ,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,196,'         DECODE(SIGN(INSTR(X1.���,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')),1,SUBSTR(X1.���,1,INSTR(X1.���,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')-1),X1.���) AS ���,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,197,'         DECODE(SIGN(INSTR(X1.���,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')),1,SUBSTR(X1.���,INSTR(X1.���,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')+1,1000),'||CHR(39)||''||CHR(39)||') AS ��־�ο�       ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,198,'  from ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,199,'       (');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,200,'       select ����id AS ִ�в���ID,����id AS ID  from �����Ա���� WHERE ����id=[1] AND �Ǽ�id=[0]');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,201,'       ) A,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,202,'       ���˲������� X,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,203,'       (');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,204,'       select  ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,205,'               A.����ID,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,206,'               A.�ؼ��� AS �����,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,207,'               B.������ AS ��Ŀ,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,208,'               DECODE(A.��������,NULL,NULL,A.��������||'||CHR(39)||' '||CHR(39)||'||B.��λ) AS ���');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,209,'        from ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,210,'          ���˲��������� A,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,211,'          ����������Ŀ B');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,212,'        where A.������id=B.ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,213,'              and ������id>0');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,214,'        ) X1');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,215,'  where X.������¼id=A.ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,216,'        AND X.ID=X1.����id');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,217,'  union all     ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,218,'  select ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,219,'         A.ִ�в���ID,            ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,220,'         A.ID,       ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,221,'         X.�����ı� AS ��Ŀ,       ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,222,'         X1.���,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,223,'         '||CHR(39)||''||CHR(39)||' AS ��־�ο�   ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,224,'  from ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,225,'       (');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,226,'       select ����id AS ִ�в���ID,����id AS ID  from �����Ա���� WHERE ����id=[1] AND �Ǽ�id=[0]');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,227,'       ) A,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,228,'       ���˲������� X,     ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,229,'        (');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,230,'        select ����id,0 AS �����,'||CHR(39)||''||CHR(39)||' AS ��Ŀ,���� AS ��� from ���˲����ı���   ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,231,'        ) X1');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,232,'  where X.������¼id=A.ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,233,'        AND X.ID=X1.����id');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,234,'        AND X.Ԫ������ IN (0,4,-5)');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,235,'  ) R,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,236,'  ���ű� U');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,237,'  WHERE R.ִ�в���id=U.ID)       ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,238,')       ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,239,'ORDER BY ����1,����2,����3');
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,NULL,0,'�Ǽ�ID',1,NULL,0,NULL,NULL,NULL,NULL,NULL,NULL);
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,NULL,1,'����ID',1,NULL,0,NULL,NULL,NULL,NULL,NULL,NULL);

--����ZL1_BILL_1862/�����������վ�
Insert Into zlReports(ID,���,����,˵��,����,W,H,ֽ��,ֽ��,��ֽ,��ӡ��,��ֽ̬��,Ʊ��,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��) Values(zlReports_ID.NextVal,'ZL1_BILL_1862','�����������վ�','�վݴ�ӡ','Zp,fXhpso<0TfvnmI<BD',12191,5443,256,1,7,'Star AR-3200+',0,1,100,1862,'�վݴ�ӡ',Sysdate,Sysdate);
Insert Into zlRPTFMTs(����ID,���,˵��,ͼ��) Values(zlReports_ID.CurrVal,1,'��������վ�',0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ1',2,NULL,0,NULL,0,'[�շѻ���.Ʊ�ݺ�][�շѻ���.��Դ]',NULL,555,780,3645,225,0,0,1,'����',11,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ5',2,NULL,0,NULL,0,'[�շѻ���.����]',NULL,555,4320,1350,180,0,0,1,'����',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ2',2,NULL,0,NULL,0,'[�շѻ���.����]',NULL,915,1125,1710,225,0,0,1,'����',11,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ4',2,NULL,0,NULL,0,'[�շѻ���.��д]',NULL,1395,3945,1710,225,0,0,1,'����',11,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ3',2,NULL,0,NULL,0,'[�շѻ���.�ϼ�]',NULL,1650,3555,1710,225,0,2,1,'����',11,0,0,0,0,16777215,0,NULL,'0.00',NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ6',2,NULL,0,NULL,0,'[�շѻ���.����Ա����]',NULL,2310,4320,1890,180,0,0,1,'����',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ7',2,NULL,0,NULL,0,'[�շѻ���.Ʊ�ݺ�][�շѻ���.��Դ]',NULL,4410,765,3645,225,0,0,1,'����',11,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ8',2,NULL,0,NULL,0,'[�շѻ���.����]',NULL,4415,4310,1350,180,0,0,1,'����',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ9',2,NULL,0,NULL,0,'[�շѻ���.����]',NULL,4770,1110,1710,225,0,0,1,'����',11,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ10',2,NULL,0,NULL,0,'[�շѻ���.��д]',NULL,5255,3935,1710,225,0,0,1,'����',11,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ11',2,NULL,0,NULL,0,'[�շѻ���.�ϼ�]',NULL,5505,3540,1710,225,0,2,1,'����',11,0,0,0,0,16777215,0,NULL,'0.00',NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ12',2,NULL,0,NULL,0,'[�շѻ���.����Ա����]',NULL,6165,4305,1890,180,0,0,1,'����',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ13',2,NULL,0,NULL,0,'[�շѻ���.Ʊ�ݺ�][�շѻ���.��Դ]',NULL,8265,765,3645,225,0,0,1,'����',11,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ14',2,NULL,0,NULL,0,'[�շѻ���.����]',NULL,8275,4315,1350,180,0,0,1,'����',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ15',2,NULL,0,NULL,0,'[�շѻ���.����]',NULL,8625,1110,1710,225,0,0,1,'����',11,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ16',2,NULL,0,NULL,0,'[�շѻ���.��д]',NULL,9115,3940,1710,225,0,0,1,'����',11,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ17',2,NULL,0,NULL,0,'[�շѻ���.�ϼ�]',NULL,9375,3555,1710,225,0,2,1,'����',11,0,0,0,0,16777215,0,NULL,'0.00',NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ18',2,NULL,0,NULL,0,'[�շѻ���.����Ա����]',NULL,10020,4305,1890,180,0,0,1,'����',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'�����1',4,NULL,0,NULL,0,'�շ���ϸ',NULL,556,1938,2858,1485,465,0,0,'����',11,0,0,0,0,16777215,1,NULL,NULL,NULL,1,16777215,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-1,0,NULL,NULL,'[�շ���ϸ.��Ŀ]','4^30^#',0,0,1380,0,255,1,0,'����',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-2,1,NULL,NULL,'[�շ���ϸ.���]','4^30^#',0,0,1425,0,255,1,0,'����',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'�����2',4,NULL,0,NULL,0,'�շ���ϸ',NULL,4416,1928,2858,1485,465,0,0,'����',11,0,0,0,0,16777215,1,NULL,NULL,NULL,1,16777215,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-1,0,NULL,NULL,'[�շ���ϸ.��Ŀ]','4^30^#',300,300,1380,0,255,1,0,'����',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-2,1,NULL,NULL,'[�շ���ϸ.���]','4^30^#',300,300,1425,0,255,1,0,'����',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'�����3',4,NULL,0,NULL,0,'�շ���ϸ',NULL,8276,1933,2858,1485,465,0,0,'����',11,0,0,0,0,16777215,1,NULL,NULL,NULL,1,16777215,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-1,0,NULL,NULL,'[�շ���ϸ.��Ŀ]','4^30^#',600,600,1380,0,255,1,0,'����',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-2,1,NULL,NULL,'[�շ���ϸ.���]','4^30^#',600,600,1425,0,255,1,0,'����',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'�շ���ϸ','��Ŀ,200|���,200',USER||'.���˷��ü�¼',0);
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,1,'--��Ϊ�в����˷��ش�,��˲��ܼ�¼״̬;�൥���շ�ʱ,���������˶��NO');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,2,'Select ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,3,'	�վݷ�Ŀ as ��Ŀ,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,4,'	Ltrim(To_Char(Sum(Nvl(���ʽ��,0)),'||CHR(39)||'999999990.00'||CHR(39)||')) as ���');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,5,'From ���˷��ü�¼');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,6,'Where Mod(��¼����,10)=2 And ����id=[0] and ��¼״̬<>0');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,7,'Group by �վݷ�Ŀ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,8,'Having Sum(Nvl(���ʽ��,0))<>0');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,9,'Order by �վݷ�Ŀ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,10,NULL);
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,NULL,0,'����id',1,'0',0,NULL,NULL,NULL,NULL,NULL,NULL);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'�շѻ���','ҳ��,139|Ʊ�ݺ�,200|NO,200|����,200|��Դ,200|����Ա���,200|����Ա����,200|����,200|�ϼ�,200|��д,200',USER||'.���˷��ü�¼,'||USER||'.ϵͳ������,'||USER||'.�������¼,'||USER||'.��Լ��λ,'||USER||'.Ʊ�ݴ�ӡ����,'||USER||'.���˽��ʼ�¼,'||USER||'.Ʊ��ʹ����ϸ',0);
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,1,'--Ʊ�ݺ�û�й̶����䵽������շ��д���,��˸����վݷ�Ŀ�������');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,2,'--�����վݷ�Ŀʱ,��Ϊ�в����˷��ش�,��˲��ܼ�¼״̬,���Ȱ��վݷ�Ŀ����,�ٰ�Ʊ�ݺ�����');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,3,'--֧�ֶ��ŵ����շ�ͳһ��ӡƱ�ݵķ�ʽ(���������˶��NO),���ŵ��ݵ�Ʊ�ݴ�ӡ����ID��ͬ��');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,4,'--�ӱ�A�������վ��д����ü������е��վݷ�Ŀ,����ÿ��Ʊ�ݵĻ��ܽ��');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,5,'--�ӱ�B�����ص����еĲ�����Ϣ,��Ϊ�൥���շ�ʱ���Ե����޸�,���ȡ�����Ч��¼');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,6,'--�ӱ�C�����ص��ݶ�Ӧ����Ʊ�ݺ�');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,7,'Select ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,8,'	A.ҳ��,C.���� As Ʊ�ݺ�,B.NO,B.����,B.��Դ,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,9,'	B.����Ա���,B.����Ա����,To_Char(B.�Ǽ�ʱ��,'||CHR(39)||'YYYY-MM-DD'||CHR(39)||') As ����,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,10,'	Ltrim(To_Char(A.���,'||CHR(39)||'9999999.00'||CHR(39)||')) As �ϼ�,zlUppMoney(A.���) As ��д');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,11,'From (Select Ceil(A.���/B.�վ��д�) As ҳ��,Sum(A.���) As ���');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,12,'		From (Select Rownum As ���,��Ŀ,���');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,13,'				From (');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,14,'				Select �վݷ�Ŀ As ��Ŀ,Sum(���ʽ��) As ���');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,15,'				From ���˷��ü�¼');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,16,'				Where ����id=[0] And Mod(��¼����,10)=2 and ��¼״̬<>0');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,17,'				Group By �վݷ�Ŀ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,18,'					)');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,19,'			) A,(Select Nvl(Nvl(����ֵ,ȱʡֵ),3) as �վ��д� From ϵͳ������ Where ������=4) B');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,20,'		Group By Ceil(A.���/B.�վ��д�)');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,21,'	) A,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,22,'	(Select ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,23,'		Min(NO)||DeCode(Max(NO),Min(NO),Null,'||CHR(39)||'-'||CHR(39)||' || Max(NO)) As NO,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,24,'		Max(C.����) as ����,Max(A.����Ա���) as ����Ա���,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,25,'		Max(A.����Ա����) as ����Ա����,Max(A.�Ǽ�ʱ��) As �Ǽ�ʱ��,Decode(Max(A.�����־),2,'||CHR(39)||'(סԺ�շ�)'||CHR(39)||',NULL) as ��Դ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,26,'		From ���˷��ü�¼ A,�������¼ B,��Լ��λ C');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,27,'		Where Mod(A.��¼����,10)=2 And A.��¼״̬ In (1,3) And A.���=1 And A.����id=[0] AND A.����id=B.����id AND C.ID=B.��Լ��λid');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,28,'	) B,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,29,'	(Select Rownum As ҳ��,A.����');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,30,'		From Ʊ��ʹ����ϸ A,Ʊ�ݴ�ӡ���� B');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,31,'		Where B.��������=1 And B.ID=(Select Max(A.ID) From Ʊ�ݴ�ӡ���� A,���˽��ʼ�¼ B Where A.��������=1 And A.NO=B.NO AND B.ID=[0])');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,32,'			And A.��ӡID=B.ID And A.Ʊ��=1 And A.����=1');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,33,'	) C');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,34,'Where A.ҳ��=C.ҳ��(+)');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,35,'Order By C.����');
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,NULL,0,'����id',1,'0',0,NULL,NULL,NULL,NULL,NULL,NULL);

--����ZL1_REPORT_1876/���ҹ�����ͳ��
Insert Into zlReports(ID,���,����,˵��,����,W,H,ֽ��,ֽ��,��ֽ,��ӡ��,��ֽ̬��,Ʊ��,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��) Values(zlReports_ID.NextVal,'ZL1_REPORT_1876','���ҹ�����ͳ��','ͳ��һ��ʱ������������칤���������','Ew?vNub{b-<XqdldZ2ZZ',11904,16832,9,1,15,NULL,0,0,100,1876,'����',Sysdate,Sysdate);
Insert Into zlRPTFMTs(����ID,���,˵��,ͼ��) Values(zlReports_ID.CurrVal,1,'���ҹ�����ͳ��1',0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ2',2,NULL,0,'���ܱ�1',11,'ͳ�Ʒ�Χ:[=��ʼ����]��[=��������]',NULL,675,1650,2970,180,0,0,1,'����',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ1',2,NULL,0,'���ܱ�1',12,'���ҹ�����ͳ��',NULL,4260,810,2625,375,0,1,1,'����',18,1,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'���ܱ�1',5,NULL,0,NULL,0,'���ҹ�����',NULL,675,1935,9795,7950,255,0,0,'����',9,0,0,0,0,16777215,1,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,7,zlRPTItems_ID.CurrVal-1,0,NULL,NULL,'������',NULL,0,0,1605,0,255,0,0,'����',0,0,0,0,0,0,0,NULL,NULL,'SUM',1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,7,zlRPTItems_ID.CurrVal-2,1,NULL,NULL,'��Ŀ����',NULL,0,0,6750,0,0,0,0,'����',0,0,0,0,0,0,0,'��Ŀ����',NULL,'SUM',1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,9,zlRPTItems_ID.CurrVal-3,0,NULL,NULL,'�˴�',NULL,0,0,735,0,255,1,0,'����',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'���ҹ�����','������,200|��Ŀ����,200|�˴�,139',USER||'.���ǼǼ�¼,'||USER||'.�����Ŀ�嵥,'||USER||'.�����Ŀҽ��,'||USER||'.����ҽ����¼,'||USER||'.����ҽ������,'||USER||'.������ĿĿ¼,'||USER||'.���ű�',1);
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,1,'select G.���� AS ������,F.���� AS ��Ŀ����, COUNT(1) AS �˴�');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,2,'from ���ǼǼ�¼ A,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,3,'     �����Ŀ�嵥 B,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,4,'     �����Ŀҽ�� C,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,5,'     ����ҽ����¼ D,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,6,'     ����ҽ������ E,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,7,'     ������ĿĿ¼ F,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,8,'     ���ű� G');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,9,'WHERE A.ID=B.�Ǽ�ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,10,'      AND C.�嵥ID=B.ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,11,'      AND (D.ID=C.ҽ��id OR D.���ID=C.ҽ��id)');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,12,'      AND D.������� IN ('||CHR(39)||'C'||CHR(39)||','||CHR(39)||'D'||CHR(39)||')');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,13,'      AND E.ҽ��ID=D.ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,14,'      AND E.����ID>0 ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,15,'	AND A.���״̬=5');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,16,'      AND F.ID=D.������Ŀid');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,17,'      AND G.ID=E.ִ�в���ID	');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,18,'      AND A.���ʱ�� BETWEEN [0] and [1]+1-1/24/60/60 ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,19,'GROUP BY G.����,F.����');
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,NULL,0,'��ʼ����',2,CHR(38)||'ǰһ������',0,NULL,NULL,NULL,NULL,NULL,NULL);
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,NULL,1,'��������',2,CHR(38)||'��ǰ����',0,NULL,NULL,NULL,NULL,NULL,NULL);

--����ZL1_REPORT_1877/ҽ��������ͳ��
Insert Into zlReports(ID,���,����,˵��,����,W,H,ֽ��,ֽ��,��ֽ,��ӡ��,��ֽ̬��,Ʊ��,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��) Values(zlReports_ID.NextVal,'ZL1_REPORT_1877','ҽ��������ͳ��','ͳ��һ��ʱ�������ҽ����칤���������','Ww?vNtbib-<XqddZ2ZZ',11904,16832,9,1,15,NULL,0,0,100,1877,'����',Sysdate,Sysdate);
Insert Into zlRPTFMTs(����ID,���,˵��,ͼ��) Values(zlReports_ID.CurrVal,1,'ҽ��������ͳ��1',0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ2',2,NULL,0,'���ܱ�1',11,'ͳ�Ʒ�Χ:[=��ʼ����]��[=��������]',NULL,675,1650,2970,180,0,0,1,'����',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ1',2,NULL,0,'���ܱ�1',12,'ҽ��������ͳ��',NULL,4245,810,2655,360,0,1,1,'����',18,1,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'���ܱ�1',5,NULL,0,NULL,0,'ҽ��������',NULL,675,1935,9795,7950,255,0,0,'����',9,0,0,0,0,16777215,1,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,7,zlRPTItems_ID.CurrVal-1,0,NULL,NULL,'ҽ��',NULL,0,0,1605,0,255,0,0,'����',0,0,0,0,0,0,0,'ҽ��',NULL,'SUM',1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,7,zlRPTItems_ID.CurrVal-2,1,NULL,NULL,'��Ŀ����',NULL,0,0,6750,0,0,0,0,'����',0,0,0,0,0,0,0,'��Ŀ����',NULL,'SUM',1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,9,zlRPTItems_ID.CurrVal-3,0,NULL,NULL,'�˴�',NULL,0,0,735,0,255,1,0,'����',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'ҽ��������','ҽ��,200|��Ŀ����,200|�˴�,139',USER||'.���ǼǼ�¼,'||USER||'.�����Ŀ�嵥,'||USER||'.�����Ŀҽ��,'||USER||'.����ҽ����¼,'||USER||'.����ҽ������,'||USER||'.������ĿĿ¼,'||USER||'.���˲�����¼',1);
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,1,'select G.��д�� AS ҽ��,F.���� AS ��Ŀ����, COUNT(1) AS �˴�');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,2,'from ���ǼǼ�¼ A,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,3,'     �����Ŀ�嵥 B,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,4,'     �����Ŀҽ�� C,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,5,'     ����ҽ����¼ D,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,6,'     ����ҽ������ E,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,7,'     ������ĿĿ¼ F,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,8,'     ���˲�����¼ G');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,9,'WHERE A.ID=B.�Ǽ�ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,10,'      AND C.�嵥ID=B.ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,11,'      AND (D.ID=C.ҽ��id OR D.���ID=C.ҽ��id)');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,12,'      AND D.������� IN ('||CHR(39)||'C'||CHR(39)||','||CHR(39)||'D'||CHR(39)||')');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,13,'      AND E.ҽ��ID=D.ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,14,'      AND E.����ID>0 ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,15,'      AND F.ID=D.������Ŀid');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,16,'      AND G.ID=E.����ID	');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,17,'	AND A.���״̬=5');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,18,'      AND A.���ʱ�� BETWEEN [0] and [1]+1-1/24/60/60 ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,19,'GROUP BY G.��д��,F.����');
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,NULL,0,'��ʼ����',2,CHR(38)||'ǰһ������',0,NULL,NULL,NULL,NULL,NULL,NULL);
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,NULL,1,'��������',2,CHR(38)||'��ǰ����',0,NULL,NULL,NULL,NULL,NULL,NULL);

--����ZL1_REPORT_1878/�������ͳ�Ʒ���
Insert Into zlReports(ID,���,����,˵��,����,W,H,ֽ��,ֽ��,��ֽ,��ӡ��,��ֽ̬��,Ʊ��,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��) Values(zlReports_ID.NextVal,'ZL1_REPORT_1878','�������ͳ�Ʒ���','ͳ��һ��ʱ���ڸ�����������������','Zn*Venhe 4GqdooI"D]',11904,16832,9,1,15,NULL,0,0,100,1878,'����',Sysdate,Sysdate);
Insert Into zlRPTFMTs(����ID,���,˵��,ͼ��) Values(zlReports_ID.CurrVal,1,'�������ͳ��1',0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ1',2,NULL,0,'�����1',11,'ͳ�Ʒ�Χ:[=��ʼ����]��[=��������]',NULL,330,1965,2970,180,0,0,1,'����',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ2',2,NULL,0,'�����1',12,'�������ͳ��',NULL,4550,1080,2250,375,0,1,1,'����',18,1,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'�����1',4,NULL,0,NULL,0,'���ǼǼ�¼_����',NULL,330,2265,10690,7200,255,0,0,'����',9,0,0,0,0,16777215,1,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-1,0,NULL,NULL,'[���ǼǼ�¼_����.��������]','1^255^��������|1^255^��������',0,0,3735,0,255,0,0,'����',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-2,1,NULL,NULL,'[���ǼǼ�¼_����.��������]','4^255^����|4^255^����',0,0,630,0,255,1,0,'����',0,0,0,0,0,0,0,NULL,NULL,'SUM',1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-3,2,NULL,NULL,'[���ǼǼ�¼_����.Ů������]','4^255^����|4^255^Ů��',0,0,585,0,255,1,0,'����',0,0,0,0,0,0,0,NULL,NULL,'SUM',1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-4,3,NULL,NULL,'[���ǼǼ�¼_����.����]','4^255^����|4^255^�ϼ�',0,0,810,0,255,1,0,'����',0,0,0,0,0,0,0,NULL,NULL,'SUM',1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-5,4,NULL,NULL,'[���ǼǼ�¼_����.�Ѽ���������]','4^255^�Ѽ�����|4^255^����',0,0,630,0,255,1,0,'����',0,0,0,0,0,0,0,NULL,NULL,'SUM',1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-6,5,NULL,NULL,'[���ǼǼ�¼_����.�Ѽ�Ů������]','4^255^�Ѽ�����|4^255^Ů��',0,0,675,0,255,1,0,'����',0,0,0,0,0,0,0,NULL,NULL,'SUM',1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-7,6,NULL,NULL,'[���ǼǼ�¼_����.�Ѽ�����]','4^255^�Ѽ�����|4^255^�ϼ�',0,0,795,0,255,1,0,'����',0,0,0,0,0,0,0,NULL,NULL,'SUM',1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-8,7,NULL,NULL,'[���ǼǼ�¼_����.δ����������]','4^255^δ������|4^255^����',0,0,705,0,255,1,0,'����',0,0,0,0,0,0,0,NULL,NULL,'SUM',1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-9,8,NULL,NULL,'[���ǼǼ�¼_����.δ��Ů������]','4^255^δ������|4^255^Ů��',0,0,690,0,255,1,0,'����',0,0,0,0,0,0,0,NULL,NULL,'SUM',1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-10,9,NULL,NULL,'[���ǼǼ�¼_����.δ������]','4^255^δ������|4^255^�ϼ�',0,0,1005,0,255,1,0,'����',0,0,0,0,0,0,0,NULL,NULL,'SUM',1,0,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'���ǼǼ�¼_����','��������,200|��������,200|Ů������,200|����,200|�Ѽ���������,200|�Ѽ�Ů������,200|�Ѽ�����,200|δ����������,200|δ��Ů������,200|δ������,200',USER||'.���ǼǼ�¼,'||USER||'.�����Ա����,'||USER||'.��Լ��λ',0);
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,1,'SELECT ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,2,'	��������,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,3,'	DECODE(��������,0,NULL,��������) AS ��������,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,4,'	DECODE(Ů������,0,NULL,Ů������) AS Ů������,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,5,'	DECODE(����,0,NULL,����) AS ����,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,6,'	DECODE(�Ѽ���������,0,NULL,�Ѽ���������) AS �Ѽ���������,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,7,'	DECODE(�Ѽ�Ů������,0,NULL,�Ѽ�Ů������) AS �Ѽ�Ů������,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,8,'	DECODE(�Ѽ�����,0,NULL,�Ѽ�����) AS �Ѽ�����,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,9,'	DECODE(δ����������,0,NULL,δ����������) AS δ����������,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,10,'	DECODE(δ��Ů������,0,NULL,δ��Ů������) AS δ��Ů������,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,11,'	DECODE(δ������,0,NULL,δ������) AS δ������              ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,12,'FROM ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,13,'(');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,14,'SELECT B.���� AS ��������,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,15,'       A.��������,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,16,'       A.Ů������,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,17,'       nvl(A.��������,0)+nvl(A.Ů������,0) AS ����,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,18,'       A.�Ѽ���������,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,19,'       A.�Ѽ�Ů������,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,20,'       nvl(A.�Ѽ���������,0)+nvl(A.�Ѽ�Ů������,0) AS �Ѽ�����,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,21,'       nvl(A.��������,0)-nvl(A.�Ѽ���������,0) AS δ����������,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,22,'       nvl(A.Ů������,0)-nvl(A.�Ѽ�Ů������,0) AS δ��Ů������,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,23,'       (nvl(A.��������,0)-nvl(A.�Ѽ���������,0))+(nvl(A.Ů������,0)-nvl(A.�Ѽ�Ů������,0)) AS δ������              ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,24,'FROM ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,25,'(');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,26,'select A.��Լ��λid,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,27,'       SUM(DECODE(sign(instr(B.�Ա�,'||CHR(39)||'Ů'||CHR(39)||')-0),1,0,1)) AS ��������,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,28,'       SUM(DECODE(sign(instr(B.�Ա�,'||CHR(39)||'Ů'||CHR(39)||')-0),1,1,0)) AS Ů������,       ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,29,'       SUM(DECODE(sign(0 - NVL(B.��첡��ID,0)),-1, DECODE(SIGN(instr(B.�Ա�,'||CHR(39)||'Ů'||CHR(39)||')-0),1,0,1),0)) AS �Ѽ���������,       ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,30,'       SUM(DECODE(sign(0 - NVL(B.��첡��ID,0)),-1, DECODE(SIGN(instr(B.�Ա�,'||CHR(39)||'Ů'||CHR(39)||')-0),1,1,0),0)) AS �Ѽ�Ů������');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,31,'from ���ǼǼ�¼ A,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,32,'     �����Ա���� B     ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,33,'WHERE A.ID=B.�Ǽ�ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,34,'      AND A.��Լ��λid>0');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,35,'	AND A.���״̬=5');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,36,'      AND A.���ʱ�� BETWEEN [0] AND [1]+1-1/24/60/60');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,37,'GROUP BY A.��Լ��λid');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,38,') A,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,39,'��Լ��λ B');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,40,'WHERE A.��Լ��λid=B.ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,41,')');
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,NULL,0,'��ʼ����',2,CHR(38)||'ǰһ������',0,NULL,NULL,NULL,NULL,NULL,NULL);
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,NULL,1,'��������',2,CHR(38)||'��ǰ����',0,NULL,NULL,NULL,NULL,NULL,NULL);

--����ZL1_REPORT_1879/������Ա�嵥
Insert Into zlReports(ID,���,����,˵��,����,W,H,ֽ��,ֽ��,��ֽ,��ӡ��,��ֽ̬��,Ʊ��,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��) Values(zlReports_ID.NextVal,'ZL1_REPORT_1879','������Ա�嵥','ͳ��һ��ʱ������Ҫ�������Ա','Hg*uSjnsc37PcmznL,PM',11904,16832,9,1,15,NULL,0,0,100,1879,'����',Sysdate,Sysdate);
Insert Into zlRPTFMTs(����ID,���,˵��,ͼ��) Values(zlReports_ID.CurrVal,1,'������Ա�嵥1',0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ3',2,NULL,0,'�����1',11,'��������:[=��ʼ����]��[=��������]',NULL,600,1755,2970,180,0,0,1,'����',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ1',2,NULL,0,'�����1',12,'������Ա�嵥',NULL,4855,1125,2250,360,0,1,1,'����',18,1,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'�����1',4,NULL,0,NULL,0,'����Դ',NULL,600,2040,10760,5865,255,0,0,'����',9,0,0,0,0,16777215,1,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-1,0,NULL,NULL,'[����Դ.����]','4^255^����',0,0,1005,0,255,1,0,'����',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-2,1,NULL,NULL,'[����Դ.����]','4^255^����',0,0,1005,0,255,1,0,'����',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-3,2,NULL,NULL,'[����Դ.�������]','4^255^�������',0,0,5130,0,255,0,0,'����',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-4,3,NULL,NULL,'[����Դ.���ʱ��]','4^255^���ʱ��',0,0,1785,0,255,1,0,'����',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-5,4,NULL,NULL,'[����Դ.����ʱ��]','4^255^����ʱ��',0,0,1290,0,255,1,0,'����',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'����Դ','����,200|����,131|�������,200|���ʱ��,200|����ʱ��,200',USER||'.���ǼǼ�¼,'||USER||'.�����Ա����,'||USER||'.��Լ��λ',0);
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,1,'select B.����,A.����,F.���� AS �������,TO_CHAR(A.���ʱ��,'||CHR(39)||'yyyy-mm-dd'||CHR(39)||') AS ���ʱ��,to_char(B.����ʱ��,'||CHR(39)||'yyyy-mm-dd'||CHR(39)||') as ����ʱ��');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,2,'from ���ǼǼ�¼ A,�����Ա���� B,��Լ��λ F');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,3,'where A.��Լ��λid=F.ID	');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,4,'      AND A.���״̬=5');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,5,'      AND A.ID=B.�Ǽ�ID	      ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,6,'      AND B.����ʱ�� IS NOT NULL');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,7,'      AND B.����ʱ�� BETWEEN [0] and [1]+1-1/24/60/60 ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,8,NULL);
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,NULL,0,'��ʼ����',2,CHR(38)||'��ǰ����',0,NULL,'select id,decode(�ϼ�id,null,-1,�ϼ�id) as �ϼ�id,
	����,���� from ��Լ��λ
where ĩ��<>1
Start With �ϼ�id is null
Connect By prior id=�ϼ�id','select * from ��Լ��λ
where (���� like '||CHR(39)||'%[*]%'||CHR(39)||'
or ���� like '||CHR(39)||'%[*]%'||CHR(39)||'
or ���� like '||CHR(39)||'%[*]%'||CHR(39)||') AND ĩ��=1','ID,131,'||CHR(38)||'R|�ϼ�ID,139,|����,200,'||CHR(38)||'S|����,200,','ID,131,'||CHR(38)||'B|�ϼ�ID,131,'||CHR(38)||'R|����,200,'||CHR(38)||'S|����,200,'||CHR(38)||'S'||CHR(38)||'D|����,200,'||CHR(38)||'S|ĩ��,131,|��ַ,200,|�绰,200,|��������,200,|�ʺ�,200,|��ϵ��,200,|����ʱ��,135,|����ʱ��,135,|�����ʼ�,200,|˵��,200,',USER||'.��Լ��λ|'||USER||'.��Լ��λ');
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,NULL,1,'��������',2,CHR(38)||'��һ������',0,NULL,NULL,NULL,NULL,NULL,NULL);

--����ZL1_SUB_1875_1/�������������(����)
Insert Into zlReports(ID,���,����,˵��,����,W,H,ֽ��,ֽ��,��ֽ,��ӡ��,��ֽ̬��,Ʊ��,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��) Values(zlReports_ID.NextVal,'ZL1_SUB_1875_1','�������������(����)',NULL,'Zp,fI`z?<,6'||CHR(39)||''||CHR(38)||'pkq[0\L',11904,16832,9,1,15,NULL,0,0,100,NULL,NULL,Sysdate,NULL);
Insert Into zlRPTFMTs(����ID,���,˵��,ͼ��) Values(zlReports_ID.CurrVal,1,'�������������(����)1',0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ3',2,NULL,0,NULL,0,'���ʱ��:[=��ʼʱ��]��[=����ʱ��]',NULL,585,1755,2970,180,0,0,1,'����',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ4',2,NULL,0,NULL,0,'�������:[����Դ.�������]',NULL,585,2040,4230,180,0,0,1,'����',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ1',2,NULL,0,'�����1',12,'�������������(����)',NULL,3712,1125,4140,360,0,1,1,'����',18,1,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'�����1',4,NULL,0,NULL,0,'����Դ',NULL,615,2325,10335,6465,255,0,0,'����',9,0,0,0,0,16777215,1,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-1,0,NULL,NULL,'[����Դ.��������]','1^255^��������(����)',0,0,5865,0,255,0,1,'����',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-2,1,NULL,NULL,'[����Դ.����]','4^255^����',0,0,1740,0,255,1,0,'����',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'����Դ','�������,200|��������,200|����,200',USER||'.���ǼǼ�¼,'||USER||'.�����Ա����,'||USER||'.���˲�������,'||USER||'.������ϼ�¼,'||USER||'.��Լ��λ',0);
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,1,'select �������,��������||'||CHR(39)||'('||CHR(39)||'||TO_CHAR(����)||'||CHR(39)||')'||CHR(39)||' AS ��������,���� from ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,2,'(');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,3,'select D.�������,D.������� AS ��������,B.���� AS ����,F.���� AS �������');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,4,'from ���ǼǼ�¼ A,�����Ա���� B,���˲������� C,������ϼ�¼ D,��Լ��λ F');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,5,'where A.��Լ��λid=F.ID	');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,6,'      AND F.ID=[0]');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,7,'      AND A.���״̬=5');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,8,'      AND A.ID=B.�Ǽ�ID	');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,9,'      AND B.��첡��ID>0');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,10,'      AND C.������¼ID=B.��첡��ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,11,'      AND C.Ԫ������=4');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,12,'      AND D.����ID=C.ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,13,'      AND A.���ʱ�� BETWEEN [1] and [2]+1-1/24/60/60 ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,14,') A,');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,15,'(');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,16,'select D.�������,COUNT(1) AS ����');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,17,'from ���ǼǼ�¼ A,�����Ա���� B,���˲������� C,������ϼ�¼ D,��Լ��λ E');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,18,'where A.��Լ��λid=E.ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,19,'      AND E.ID=[0]');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,20,'      AND A.���״̬=5');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,21,'      AND A.ID=B.�Ǽ�ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,22,'      AND B.��첡��ID>0');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,23,'      AND C.������¼ID=B.��첡��ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,24,'      AND C.Ԫ������=4');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,25,'      AND D.����ID=C.ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,26,'      AND A.���ʱ�� BETWEEN [1] and [2]+1-1/24/60/60 ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,27,'GROUP BY D.�������      ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,28,') B      ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,29,'WHERE A.�������=B.������� ');
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,NULL,0,'�������',1,'ѡ�������塭',0,NULL,NULL,'select * from ��Լ��λ
where (���� like '||CHR(39)||'%[*]%'||CHR(39)||'
or ���� like '||CHR(39)||'%[*]%'||CHR(39)||'
or ���� like '||CHR(39)||'%[*]%'||CHR(39)||') AND ĩ��=1',NULL,'ID,131,'||CHR(38)||'B|�ϼ�ID,131,|����,200,'||CHR(38)||'S|����,200,'||CHR(38)||'S'||CHR(38)||'D|����,200,'||CHR(38)||'S|ĩ��,131,|��ַ,200,|�绰,200,|��������,200,|�ʺ�,200,|��ϵ��,200,|����ʱ��,135,|����ʱ��,135,|�����ʼ�,200,|˵��,200,',USER||'.��Լ��λ|');
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,NULL,1,'��ʼʱ��',2,CHR(38)||'ǰһ������',0,NULL,NULL,NULL,NULL,NULL,NULL);
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,NULL,2,'����ʱ��',2,CHR(38)||'��ǰ����',0,NULL,NULL,'select * from ��Լ��λ
where (���� like '||CHR(39)||'%[*]%'||CHR(39)||'
or ���� like '||CHR(39)||'%[*]%'||CHR(39)||'
or ���� like '||CHR(39)||'%[*]%'||CHR(39)||') AND ĩ��=1',NULL,'ID,131,'||CHR(38)||'B|�ϼ�ID,131,|����,200,'||CHR(38)||'S|����,200,'||CHR(38)||'S'||CHR(38)||'D|����,200,'||CHR(38)||'S|ĩ��,131,|��ַ,200,'||CHR(38)||'S|�绰,200,|��������,200,|�ʺ�,200,|��ϵ��,200,|����ʱ��,135,|����ʱ��,135,|�����ʼ�,200,|˵��,200,',USER||'.��Լ��λ|');

--����ZL1_SUB_1875_2/�������������(��)
Insert Into zlReports(ID,���,����,˵��,����,W,H,ֽ��,ֽ��,��ֽ,��ӡ��,��ֽ̬��,Ʊ��,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��) Values(zlReports_ID.NextVal,'ZL1_SUB_1875_2','�������������(��)',NULL,'Zp,fI`y?$:6'||CHR(39)||'8sfpE:BZ',11904,16832,9,1,15,NULL,0,0,100,NULL,NULL,Sysdate,NULL);
Insert Into zlRPTFMTs(����ID,���,˵��,ͼ��) Values(zlReports_ID.CurrVal,1,'�������������(��)1',0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ3',2,NULL,0,NULL,0,'���ʱ��:[=��ʼʱ��]��[=����ʱ��]',NULL,585,1755,2970,180,0,0,1,'����',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ4',2,NULL,0,NULL,0,'�������:[����Դ.�������]',NULL,585,2040,4230,180,0,0,1,'����',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ1',2,NULL,0,'�����1',12,'�������������(��)',NULL,3942,1125,3780,360,0,1,1,'����',18,1,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'�����1',4,NULL,0,NULL,0,'����Դ',NULL,615,2355,10435,5550,255,0,0,'����',9,0,0,0,0,16777215,1,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-1,0,NULL,NULL,'[����Դ.����]','4^255^����',0,0,1785,0,255,0,1,'����',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-2,1,NULL,NULL,'[����Դ.��������]','4^255^��������',0,0,7320,0,255,0,0,'����',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'����Դ','�������,200|����,200|��������,200',USER||'.���ǼǼ�¼,'||USER||'.�����Ա����,'||USER||'.���˲�������,'||USER||'.������ϼ�¼,'||USER||'.��Լ��λ',0);
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,1,'select �������,����||'||CHR(39)||'('||CHR(39)||'||�Ա�||'||CHR(39)||')'||CHR(39)||' AS ����,�������� from ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,2,'(');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,3,'select D.������� AS ��������,B.����,B.�Ա�,F.���� AS �������');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,4,'from ���ǼǼ�¼ A,�����Ա���� B,���˲������� C,������ϼ�¼ D,��Լ��λ F');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,5,'where A.��Լ��λid=F.ID	');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,6,'      AND F.ID=[0]');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,7,'      AND A.���״̬=5');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,8,'      AND A.ID=B.�Ǽ�ID	');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,9,'      AND B.��첡��ID>0');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,10,'      AND C.������¼ID=B.��첡��ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,11,'      AND C.Ԫ������=4');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,12,'      AND D.����ID=C.ID');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,13,'      AND A.���ʱ�� BETWEEN [1] and [2]+1-1/24/60/60 ');
Insert Into zlRPTSQLs(ԴID,�к�,����) Values(zlRPTDatas_ID.CurrVal,14,') A');
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,NULL,0,'�������',1,'ѡ�������塭',0,NULL,NULL,'select * from ��Լ��λ
where (���� like '||CHR(39)||'%[*]%'||CHR(39)||'
or ���� like '||CHR(39)||'%[*]%'||CHR(39)||'
or ���� like '||CHR(39)||'%[*]%'||CHR(39)||') AND ĩ��=1',NULL,'ID,131,'||CHR(38)||'B|�ϼ�ID,131,|����,200,'||CHR(38)||'S|����,200,'||CHR(38)||'S'||CHR(38)||'D|����,200,'||CHR(38)||'S|ĩ��,131,|��ַ,200,|�绰,200,|��������,200,|�ʺ�,200,|��ϵ��,200,|����ʱ��,135,|����ʱ��,135,|�����ʼ�,200,|˵��,200,',USER||'.��Լ��λ|');
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,NULL,1,'��ʼʱ��',2,CHR(38)||'ǰһ������',0,NULL,NULL,NULL,NULL,NULL,NULL);
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����) Values(zlRPTDatas_ID.CurrVal,NULL,2,'����ʱ��',2,CHR(38)||'��ǰ����',0,NULL,NULL,'select * from ��Լ��λ
where (���� like '||CHR(39)||'%[*]%'||CHR(39)||'
or ���� like '||CHR(39)||'%[*]%'||CHR(39)||'
or ���� like '||CHR(39)||'%[*]%'||CHR(39)||') AND ĩ��=1',NULL,'ID,131,'||CHR(38)||'B|�ϼ�ID,131,|����,200,'||CHR(38)||'S|����,200,'||CHR(38)||'S'||CHR(38)||'D|����,200,'||CHR(38)||'S|ĩ��,131,|��ַ,200,'||CHR(38)||'S|�绰,200,|��������,200,|�ʺ�,200,|��ϵ��,200,|����ʱ��,135,|����ʱ��,135,|�����ʼ�,200,|˵��,200,',USER||'.��Լ��λ|');


--�����飺ZL1_GROUP_1875/�������������
Insert Into zlRPTGroups(ID,���,����,˵��,ϵͳ,����ID,����ʱ��) Values(zlRPTGroups_ID.NextVal,'ZL1_GROUP_1875','�������������','ͳ��һ�����ʱ�䷶Χ���������������ͳ�����',100,1875,Sysdate);
Insert Into zlRPTSubs(��ID,����ID,���,����) Select zlRPTGroups_ID.CurrVal,ID,1,'�������������(����)' From zlReports Where Upper(���)=Upper('ZL1_SUB_1875_1') And ϵͳ=100;
Insert Into zlRPTSubs(��ID,����ID,���,����) Select zlRPTGroups_ID.CurrVal,ID,2,'�������������(��)' From zlReports Where Upper(���)=Upper('ZL1_SUB_1875_2') And ϵͳ=100;

--����ZL1_BILL_1861/���Ա��Ŀ�嵥
insert into zlProgFuncs(ϵͳ,���,����) values (100,1861,'��Ŀ�嵥');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'��Ŀ�嵥',USER,'�����Ŀ�嵥','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'��Ŀ�嵥',USER,'������ĿĿ¼','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'��Ŀ�嵥',USER,'���ű�','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'��Ŀ�嵥',USER,'������Ŀ���','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'��Ŀ�嵥',USER,'�����Ŀҽ��','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'��Ŀ�嵥',USER,'���ǼǼ�¼','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'��Ŀ�嵥',USER,'��Լ��λ','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'��Ŀ�嵥',USER,'������Ϣ','SELECT');

--����ZL1_BILL_1861_2/��챨����
insert into zlProgFuncs(ϵͳ,���,����) values (100,1861,'�������ӡ');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'�������ӡ',USER,'�����Ա����','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'�������ӡ',USER,'������Ϣ','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'�������ӡ',USER,'���ǼǼ�¼','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'�������ӡ',USER,'��Լ��λ','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'�������ӡ',USER,'���˲���������','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'�������ӡ',USER,'����������Ŀ','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'�������ӡ',USER,'���˲����ı���','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'�������ӡ',USER,'���˲�������','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'�������ӡ',USER,'�����Ŀҽ��','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'�������ӡ',USER,'����ҽ����¼','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'�������ӡ',USER,'����ҽ������','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'�������ӡ',USER,'���˲�����¼','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'�������ӡ',USER,'�����Ŀ�嵥','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'�������ӡ',USER,'���ű�','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'�������ӡ',USER,'������ĿĿ¼','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1861,'�������ӡ',USER,'�����Ա����','SELECT');

--����ZL1_BILL_1862/�����������վ�
insert into zlProgFuncs(ϵͳ,���,����) values (100,1862,'�վݴ�ӡ');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1862,'�վݴ�ӡ',USER,'ϵͳ������','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1862,'�վݴ�ӡ',USER,'�������¼','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1862,'�վݴ�ӡ',USER,'��Լ��λ','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1862,'�վݴ�ӡ',USER,'Ʊ�ݴ�ӡ����','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1862,'�վݴ�ӡ',USER,'���˽��ʼ�¼','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1862,'�վݴ�ӡ',USER,'Ʊ��ʹ����ϸ','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1862,'�վݴ�ӡ',USER,'���˷��ü�¼','SELECT');

--����ZL1_REPORT_1876/���ҹ�����ͳ��
insert into zlPrograms(���,����,˵��,ϵͳ,����) values(1876,'���ҹ�����ͳ��','ͳ��һ��ʱ������������칤���������',100,'zl9Report');
insert into zlProgFuncs(ϵͳ,���,����) values (100,1876,'����');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1876,'����',USER,'���ǼǼ�¼','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1876,'����',USER,'�����Ŀ�嵥','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1876,'����',USER,'�����Ŀҽ��','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1876,'����',USER,'����ҽ����¼','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1876,'����',USER,'����ҽ������','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1876,'����',USER,'������ĿĿ¼','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1876,'����',USER,'���ű�','SELECT');
insert into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Select 'ȱʡ',zlMenus_ID.Nextval,ID,'���ҹ�����ͳ��','���ҹ�����ͳ��',NULL,105,'ͳ��һ��ʱ������������칤���������',100,1876 From zlMenus Where ϵͳ=100 And ���='ȱʡ' And ����='������ϵͳ' And ģ�� is NULL;

--����ZL1_REPORT_1877/ҽ��������ͳ��
insert into zlPrograms(���,����,˵��,ϵͳ,����) values(1877,'ҽ��������ͳ��','ͳ��һ��ʱ�������ҽ����칤���������',100,'zl9Report');
insert into zlProgFuncs(ϵͳ,���,����) values (100,1877,'����');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1877,'����',USER,'���ǼǼ�¼','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1877,'����',USER,'�����Ŀ�嵥','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1877,'����',USER,'�����Ŀҽ��','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1877,'����',USER,'����ҽ����¼','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1877,'����',USER,'����ҽ������','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1877,'����',USER,'������ĿĿ¼','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1877,'����',USER,'���˲�����¼','SELECT');
insert into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Select 'ȱʡ',zlMenus_ID.Nextval,ID,'ҽ��������ͳ��','ҽ��������ͳ��',NULL,105,'ͳ��һ��ʱ�������ҽ����칤���������',100,1877 From zlMenus Where ϵͳ=100 And ���='ȱʡ' And ����='������ϵͳ' And ģ�� is NULL;

--����ZL1_REPORT_1878/�������ͳ�Ʒ���
insert into zlPrograms(���,����,˵��,ϵͳ,����) values(1878,'�������ͳ�Ʒ���','ͳ��һ��ʱ���ڸ�����������������',100,'zl9Report');
insert into zlProgFuncs(ϵͳ,���,����) values (100,1878,'����');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1878,'����',USER,'���ǼǼ�¼','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1878,'����',USER,'�����Ա����','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1878,'����',USER,'��Լ��λ','SELECT');
insert into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Select 'ȱʡ',zlMenus_ID.Nextval,ID,'�������ͳ�Ʒ���','�������ͳ�Ʒ���',NULL,105,'ͳ��һ��ʱ���ڸ�����������������',100,1878 From zlMenus Where ϵͳ=100 And ���='ȱʡ' And ����='������ϵͳ' And ģ�� is NULL;

--����ZL1_REPORT_1879/������Ա�嵥
insert into zlPrograms(���,����,˵��,ϵͳ,����) values(1879,'������Ա�嵥','ͳ��һ��ʱ������Ҫ�������Ա',100,'zl9Report');
insert into zlProgFuncs(ϵͳ,���,����) values (100,1879,'����');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1879,'����',USER,'���ǼǼ�¼','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1879,'����',USER,'�����Ա����','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1879,'����',USER,'��Լ��λ','SELECT');
insert into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Select 'ȱʡ',zlMenus_ID.Nextval,ID,'������Ա�嵥','������Ա�嵥',NULL,105,'ͳ��һ��ʱ������Ҫ�������Ա',100,1879 From zlMenus Where ϵͳ=100 And ���='ȱʡ' And ����='������ϵͳ' And ģ�� is NULL;

--�����飺ZL1_GROUP_1875/�������������
insert into zlPrograms(���,����,˵��,ϵͳ,����) values(1875,'�������������','ͳ��һ�����ʱ�䷶Χ���������������ͳ�����',100,'zl9Report');
insert into zlProgFuncs(ϵͳ,���,����) values (100,1875,'�������������(����)');
insert into zlProgFuncs(ϵͳ,���,����) values (100,1875,'�������������(��)');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1875,'�������������(����)',USER,'���ǼǼ�¼','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1875,'�������������(����)',USER,'�����Ա����','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1875,'�������������(����)',USER,'���˲�������','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1875,'�������������(����)',USER,'������ϼ�¼','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1875,'�������������(����)',USER,'��Լ��λ','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1875,'�������������(��)',USER,'���ǼǼ�¼','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1875,'�������������(��)',USER,'�����Ա����','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1875,'�������������(��)',USER,'���˲�������','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1875,'�������������(��)',USER,'������ϼ�¼','SELECT');
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1875,'�������������(��)',USER,'��Լ��λ','SELECT');
insert into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Select 'ȱʡ',zlMenus_ID.Nextval,ID,'�������������','�������������(����)',NULL,105,'ͳ��һ�����ʱ�䷶Χ���������������ͳ�����',100,1875 From zlMenus Where ϵͳ=100 And ���='ȱʡ' And ����='������ϵͳ' And ģ�� is NULL;

commit;

