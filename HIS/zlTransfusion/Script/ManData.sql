--zlComponent
--Insert Into zlComponent(����,����,���汾,�ΰ汾,���汾,ϵͳ) Values('zl9Transfusion','������Һע�䲿��',10,15,0,100);

--zlPrograms
Insert Into zlPrograms(���,����,˵��,ϵͳ,����) Values(1264,'������Һע�����','�������ﻤʿ�Խ������Ƶ����ﲡ���Ŷӹ������ƹ��̵Ǽ�',100,'zl9CISJob');

--1264:��Һ�Ŷ�(����)
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1264,'����',Null);
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1264,'���п���','���п��Ҳ���ִ����Һ�Ŷӹ��ܵ�Ȩ��');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1264,'��λ����','��������ﲡ�˰�����λ');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1264,'��λ����','���ӡ��޸ġ�ɾ��Ȩ��');

Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1264,'�Ŷӹ���','����Ա����ҵĲ��˶��н��е���');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1264,'ҽ���ӵ�','�����ҽ��ִ����Ŀ���в���');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1264,'ҽ��ִ��','�����ҽ��ִ����Ŀ���в���');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1264,'ҩƷ�Ĵ�','�ɷ����ҩƷ�Ĵ����');

--  1264:��Һ�Ŷ�(����)
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1264,'����',User,'ִ�д�ӡ��¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1264,'����',User,'�ݴ�ҩƷ��¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1264,'����',User,'������ĿĿ¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1264,'����',User,'ҩƷ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1264,'����',User,'����ҽ��ִ��','SELECT');

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1264,'����',USER,'���ű�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1264,'����',USER,'���˹Һż�¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1264,'����',USER,'������Ϣ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1264,'����',USER,'����ҽ����¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1264,'����',USER,'����ҽ������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1264,'����',USER,'������ϼ�¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1264,'����',USER,'�ŶӼ�¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1264,'����',USER,'��λ״����¼','SELECT');

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1264,'����',USER,'ZL_�ŶӼ�¼_Addqueue','EXECUTE'); 
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1264,'����',USER,'Zl_�ŶӼ�¼_Update','EXECUTE');

-- 1264:��Һ�Ŷ�(��λ����)
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1264,'��λ����',USER,'ZL_��λ״����¼_Setseating','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1264,'��λ����',USER,'ZL_��λ״����¼_Clear','EXECUTE'); 
-- 1264:��Һ�Ŷ�(��λ����)
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1264,'��λ����',USER,'Zl_��λ״����¼_Update','EXECUTE'); 
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1264,'��λ����',USER,'Zl_��λ״����¼_Insert','EXECUTE'); 
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1264,'��λ����',USER,'Zl_��λ״����¼_Delete','EXECUTE'); 
-- 1264:��Һ�Ŷ�(ҽ��ִ��)
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1264,'ҽ��ִ��',USER,'Zl_����ҽ��ִ��_Transfusion','EXECUTE'); 
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1264,'ҽ��ִ��',USER,'Zl_����ҽ��ִ��_Modify','EXECUTE'); 
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1264,'ҽ��ִ��',USER,'Zl_����ҽ��ִ��_Insert','EXECUTE'); 
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1264,'ҽ��ִ��',USER,'Zl_����ҽ��ִ��_Delete','EXECUTE'); 
-- 1264:��Һ�Ŷ�(ҩƷ�Ĵ�)
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1264,'ҩƷ�Ĵ�',USER,'Zl_�ݴ�ҩƷ��¼_Insert','EXECUTE'); 
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1264,'ҩƷ�Ĵ�',USER,'Zl_�ݴ�ҩƷ��¼_Delete','EXECUTE'); 
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1264,'ҩƷ�Ĵ�',USER,'Zl_�ݴ�ҩƷ��¼_Adviceused','EXECUTE'); 

--- �����Ĺ��̣�Ȩ�޴���
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1264,'����',USER,'����ҽ��ִ��_��ˮ��','SELECT');

--zlMenus
Insert Into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Values('ȱʡ',zlMenus_id.nextval, zlMenus_id.nextval-5,'������Һע�����','������Һ','F',200,'�������ﻤʿ�Խ������Ƶ����ﲡ���Ŷӹ������ƹ��̵Ǽ�',100,1264);

--zlBaseCode

commit;
