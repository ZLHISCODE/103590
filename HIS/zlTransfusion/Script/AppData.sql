-- ������Ʊ�
Insert Into ������Ʊ�(��Ŀ���,��Ŀ����,������,�Զ���ȱ,��Ź���) Values(19,'�ݴ�ҩƷ��',Null,0,0);

-- Ƥ������
INSERT INTO zlTools.zlNotices(���,ϵͳ,��������,���ѱ���,��������,���Ѵ���,����˳��,�������,��������,��ʼʱ��,��ֹʱ��,��������)
SELECT NVL(MAX(���),0)+1,100,'[����][����]ʱ���ѵ�����鿴�����',NULL+0,106,1,'[����];VARCHAR2|[����];VARCHAR2',3,2,SYSDATE,NULL,
'Select e.����, d.����
From ����ҽ��ִ�� a, ����ҽ������ b, ����ҽ����¼ c, ������ĿĿ¼ d, ������Ϣ e
Where a.��� = 1 And a.���� > 0 And a.ҽ��id = b.ҽ��id And a.���ͺ� = b.���ͺ� And a.ҽ��id = c.Id And
			c.������Ŀid = d.Id And d.ִ�з��� = 3 And c.����id = e.����id And Sysdate Between a.ִ��ʱ�� - (a.���� / 86400) And
			a.ִ��ʱ�� And
			b.ִ�в���id In (Select Distinct a.����id
											 From ������Ա a, ��Ա�� b
											 Where a.��Աid = b.Id And a.ȱʡ = 1 And Upper(b.����) = Upper([USER]))'
FROM zltools.zlNotices;
