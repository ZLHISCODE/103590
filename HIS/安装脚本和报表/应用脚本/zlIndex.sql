--����Ŀ¼
--1.��������,2.ҽ������,3.���˲�������,4.���û���,5.ҩƷ���Ļ���
--6.�ٴ�����,7.�ٴ�·������,8.��������,9.�������,10.�������
--11.������,12.ҽ��ҵ��,13.���˲���ҵ��,14.����ҵ��,15.ҩƷ����ҵ��
--16.�ٴ�ҽ��,17.�ٴ�·��,18.����ҵ��,19.����ҵ��,20.����ҵ��,21.���ҵ��

----------------------------------------------------------------------------
--[[1.��������]]
----------------------------------------------------------------------------
Create Index ���ǰ��ע��_IX_���� on ���ǰ��ע��(����) Tablespace zl9Indexhis;

Create Index ����ǰ��ע��_IX_���� on ����ǰ��ע��(����) Tablespace zl9Indexhis;

create index ZLMSG_TODO_IX_CREATE_TIME on ZLMSG_TODO (CREATE_TIME) tablespace ZLMSGDATA;

Create Index ������չ��Ϣ_IX_��Ŀ On ������չ��Ϣ(��Ŀ) Tablespace zl9Indexhis;
Create Index ��Ա��չ��Ϣ_IX_��Ŀ On ��Ա��չ��Ϣ(��Ŀ) Tablespace zl9Indexhis;
Create Index ����_IX_�ϼ����� On ����(�ϼ�����) Tablespace zl9Indexhis;
Create Index ��Ա��_IX_ǩ�� On ��Ա��(ǩ��) Tablespace zl9Indexhis;
Create Index ��Ա����˵��_IX_��Ա���� On ��Ա����˵��(��Ա����) Tablespace zl9Indexhis;
Create Index ��Ա֤���¼_IX_��ԱID On ��Ա֤���¼(��ԱID) Tablespace zl9Indexhis;
Create Index ��������˵��_IX_�������� On ��������˵��(��������) Tablespace zl9Indexhis;
Create Index ������Ա_IX_��ԱID On ������Ա(��ԱID) Tablespace zl9Indexhis;
Create Index �ٴ�����_IX_����ID On �ٴ�����(����ID) Tablespace zl9Indexhis;
Create Index �������Ҷ�Ӧ_IX_����ID On �������Ҷ�Ӧ(����ID) Tablespace zl9Indexhis;

Create Index �ŶӽкŶ���_IX_�������� On �ŶӽкŶ���(��������) Tablespace zl9Indexhis;
Create Index �ŶӽкŶ���_IX_����ID On �ŶӽкŶ���(����id) Tablespace zl9Indexhis;
Create Index �ŶӽкŶ���_IX_����ID On �ŶӽкŶ���(����ID) Tablespace zl9Indexhis;
create index �ŶӽкŶ���_IX_ҵ��ID on �ŶӽкŶ���(ҵ��ID,ҵ������) tablespace zl9indexhis;
create index �Ŷ���������_IX_����ID on �Ŷ���������(����ID,վ��) Tablespace zl9indexhis;
Create Index �ϻ���Ա��_IX_��ԱID On �ϻ���Ա��(��Աid)   Tablespace zl9indexhis;
----------------------------------------------------------------------------
--[[2.ҽ������]]
----------------------------------------------------------------------------
Create Index ���ս����¼_IX_����ID On ���ս����¼(����ID) Tablespace zl9Indexhis;
Create Index ���ս����¼_IX_����ʱ�� On ���ս����¼(����ʱ��) Tablespace zl9Indexhis;
Create Index ���ս����¼_IX_�����ID On ���ս����¼(�����ID) Tablespace zl9Indexhis;
Create Index ����֧����Ŀ_IX_����ID On ����֧����Ŀ(����ID,����) Tablespace zl9Indexhis;
Create Index ����֧����Ŀ_IX_��Ŀ���� On ����֧����Ŀ(��Ŀ����,����) Tablespace zl9Indexhis;
Create Index ������Ŀģ��_IX_��ĿID On ������Ŀģ��(��ĿID) Tablespace zl9Indexhis;

----------------------------------------------------------------------------
--[[3.���˲�������]]
----------------------------------------------------------------------------
create index ����ʵ����Ϣ_IX_����ʱ�� on ����ʵ����Ϣ (����ʱ��)  tablespace ZL9INDEXHIS;
create index ����ʵ����Ϣ_IX_���������֤�� on ����ʵ����Ϣ (���������֤��) tablespace ZL9INDEXHIS;
create index ����ʵ����Ϣ_IX_�ֻ��� on ����ʵ����Ϣ (�ֻ���) tablespace ZL9INDEXHIS;
create index ����ʵ����Ϣ_IX_���� on ����ʵ����Ϣ (����) tablespace ZL9INDEXHIS;
create index ����ʵ��֤��_IX_֤������ on ����ʵ��֤�� (֤������) tablespace ZL9INDEXHIS;
create index ʵ����֤�ӿ���־_IX_ʵ��ID on ʵ����֤�ӿ���־ (ʵ��ID) tablespace ZL9INDEXHIS;
create index ʵ����֤�ӿ���־_IX_�ӿ�ID on ʵ����֤�ӿ���־ (�ӿ�ID) tablespace ZL9INDEXHIS;
create index ʵ����֤�ӿ���־_IX_����ʱ�� on ʵ����֤�ӿ���־ (����ʱ��) tablespace ZL9INDEXHIS;

Create Index ҽ�ƻ���_IX_�ϼ� on ҽ�ƻ���(�ϼ�) Tablespace zl9Indexhis;

Create Index ��Ժת��_IX_�ϼ� on ��Ժת��(�ϼ�) Tablespace zl9Indexhis;

Create Index �����������_IX_�ϼ�ID On �����������(�ϼ�ID) Tablespace zl9Indexhis;
Create Index ��������Ŀ¼_IX_����ID On ��������Ŀ¼(����ID) Tablespace zl9Indexhis;
Create Index ��������Ŀ¼_IX_���� On ��������Ŀ¼(����) Tablespace zl9Indexhis;
Create Index ��������Ŀ¼_IX_���� On ��������Ŀ¼(����) Tablespace zl9Indexhis;
Create Index ��������Ŀ¼_IX_����� On ��������Ŀ¼(�����) Tablespace zl9Indexhis;
Create Index �����������_IX_����ID On �����������(����ID) Tablespace zl9Indexhis;
Create Index �����������_IX_��ԱID On �����������(��ԱID) Tablespace zl9Indexhis;
Create Index ������Ͽ���_IX_����ID On ������Ͽ���(����ID) Tablespace zl9Indexhis;
Create Index ������Ͽ���_IX_��ԱID On ������Ͽ���(��ԱID) Tablespace zl9Indexhis;
Create Index ������Ϸ���_IX_�ϼ�ID On ������Ϸ���(�ϼ�ID) Tablespace zl9Indexcis;
Create Index ������ϱ���_IX_���ID On ������ϱ���(���id) Tablespace zl9Indexcis;
Create Index ������ϱ���_IX_���� On ������ϱ���(����) Tablespace zl9Indexcis;
Create Index ������ϱ���_IX_���� On ������ϱ���(����) Tablespace zl9Indexcis;
Create Index �������ƴ�ʩ_IX_������ĿID On �������ƴ�ʩ(������ĿID) Tablespace zl9Indexcis;
Create Index ������Ϲ���_IX_��ĿID On ������Ϲ���(��ĿID) Tablespace zl9Indexcis;
Create Index ������϶���_IX_���ID On ������϶���(���ID) Tablespace zl9Indexcis;
Create Index ������϶���_IX_����ID On ������϶���(����ID) Tablespace zl9Indexcis;

Create Index ��ѯ�������_IX_��� On ��ѯ�������(���) Tablespace zl9Indexhis;
Create Index ��ѯ�������_IX_ͼƬ��� On ��ѯ�������(ͼƬ���) Tablespace zl9Indexhis;
Create Index ��ѯҳ��Ŀ¼_IX_�������� On ��ѯҳ��Ŀ¼(��������) Tablespace zl9Indexhis;
Create Index ��ѯҳ��Ŀ¼_IX_ҳ�汳�� On ��ѯҳ��Ŀ¼(ҳ�汳��) Tablespace zl9Indexhis;
Create Index ��ѯҳ��Ŀ¼_IX_�ϼ���� On ��ѯҳ��Ŀ¼(�ϼ����) Tablespace zl9Indexhis;
Create Index ��ѯҳ������_IX_ҳ�� On ��ѯҳ������(ҳ��) Tablespace zl9Indexhis;
Create Index ��ѯҳ������_IX_����� On ��ѯҳ������(�����) Tablespace zl9Indexhis;
Create Index ��ѯҳ������_IX_ҳ��ͼ�� On ��ѯҳ������(ҳ��ͼ��) Tablespace zl9Indexhis;
Create Index ��ѯ����Ŀ¼_IX_ҳ����� On ��ѯ����Ŀ¼(ҳ�����) Tablespace zl9Indexhis;
Create Index ��ѯ����Ŀ¼_IX_����ͼ�� On ��ѯ����Ŀ¼(����ͼ��) Tablespace zl9Indexhis;
Create Index ��ѯ����Ŀ¼_IX_������ On ��ѯ����Ŀ¼(������) Tablespace zl9Indexhis;
Create Index ��ѯ����Ŀ¼_IX_��ͼ��� On ��ѯ����Ŀ¼(��ͼ���) Tablespace zl9Indexhis;
Create Index ��ѯ��������_IX_���� On ��ѯ��������(ҳ�����,�������) Tablespace zl9Indexhis;
Create Index ��ѯ��������_IX_����ҳ�� On ��ѯ��������(����ҳ��) Tablespace zl9Indexhis;
Create Index ��ѯר���嵥_IX_��Աid On ��ѯר���嵥(��Աid) Tablespace zl9Indexhis;
Create Index ��ѯר���嵥_IX_����id On ��ѯר���嵥(����id) Tablespace zl9Indexhis;


----------------------------------------------------------------------------
--[[4.���û���]]
----------------------------------------------------------------------------
CREATE INDEX ����Ʊ���쳣��¼_IX_�Ǽ�ʱ�� ON ����Ʊ���쳣��¼(�Ǽ�ʱ��) TABLESPACE zl9Indexhis; 

CREATE INDEX ����Ʊ���쳣��¼_IX_����ID ON ����Ʊ���쳣��¼(����ID) TABLESPACE zl9Indexhis;
CREATE INDEX ����Ʊ���쳣��¼_IX_����Ʊ��id ON ����Ʊ���쳣��¼(����Ʊ��id) TABLESPACE zl9Indexhis;
CREATE INDEX ����Ʊ���쳣��¼_IX_���ݺ� ON ����Ʊ���쳣��¼(���ݺ�,��������) TABLESPACE zl9Indexhis; 
CREATE INDEX ����Ʊ�ݿ�Ʊ��_IX_���� ON ����Ʊ�ݿ�Ʊ��(����) TABLESPACE zl9Indexhis;

CREATE INDEX ����Ʊ�ݿ�Ʊ��_IX_����ID ON ����Ʊ�ݿ�Ʊ��(����ID) TABLESPACE zl9Indexhis;
CREATE INDEX Ʊ�ݿ�Ʊ�����_IX_��ԱID ON Ʊ�ݿ�Ʊ�����(��ԱID) TABLESPACE zl9Indexhis;

CREATE INDEX Ʊ�ݿ�Ʊ�����_IX_�ͻ��� ON Ʊ�ݿ�Ʊ�����(�ͻ���) TABLESPACE zl9Indexhis;
CREATE INDEX ��Լ��λ_IX_���� ON ��Լ��λ(����) TABLESPACE zl9Indexhis;

CREATE INDEX ����Ʊ��ʹ�ü�¼_IX_�Ǽ�ʱ�� ON ����Ʊ��ʹ�ü�¼(�Ǽ�ʱ��) TABLESPACE zl9Indexhis;

CREATE INDEX ����Ʊ��ʹ�ü�¼_IX_����ʱ�� ON ����Ʊ��ʹ�ü�¼(����ʱ��) TABLESPACE zl9Indexhis;
CREATE INDEX ����Ʊ��ʹ�ü�¼_IX_����ID ON ����Ʊ��ʹ�ü�¼(����ID) TABLESPACE zl9Indexhis;
CREATE INDEX ����Ʊ��ʹ�ü�¼_IX_ԭƱ��ID ON ����Ʊ��ʹ�ü�¼(ԭƱ��ID) TABLESPACE zl9Indexhis;
Create Index ����Ʊ��ʹ�ü�¼_IX_��ת�� On ����Ʊ��ʹ�ü�¼(��ת��) Tablespace zl9Indexcis;
Create Index ����Ʊ�ݶ�ά��_IX_��ת�� On ����Ʊ�ݶ�ά��(��ת��) Tablespace zl9Indexcis;

Create Index ���˷����쳣��¼_IX_NO On ���˷����쳣��¼(NO,��¼����) Pctfree 5 Tablespace zl9Indexhis;

Create Index ���˷����쳣��¼_IX_����ID On ���˷����쳣��¼(����ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index ���˽����쳣��¼_IX_�Ǽ�ʱ�� On ���˽����쳣��¼(�Ǽ�ʱ��) Pctfree 5 Tablespace zl9Indexhis;
Create Index ���˽����쳣��¼_IX_����id On ���˽����쳣��¼(����id) Pctfree 5 Tablespace zl9Indexhis;
Create Index ���˽����쳣��¼_IX_Ԥ������ On ���˽����쳣��¼(Ԥ������) Pctfree 5 Tablespace zl9Indexhis;
Create Index ���˽����쳣��¼_IX_ҽ�ƿ����� On ���˽����쳣��¼(ҽ�ƿ�����) Pctfree 5 Tablespace zl9Indexhis;

Create Index ����Ѻ���¼_IX_����ID On ����Ѻ���¼(����ID) Pctfree 5 Tablespace zl9Indexhis;

Create Index ����Ѻ���¼_IX_��ҳID On ����Ѻ���¼(����ID,��ҳID) Pctfree 5 Tablespace zl9Indexhis;
Create Index ����Ѻ���¼_IX_�ɿ���ID On ����Ѻ���¼(�ɿ���ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index ����Ѻ���¼_IX_�տ�ʱ�� On ����Ѻ���¼(�տ�ʱ��) Pctfree 5 Tablespace zl9Indexhis;
Create Index ����Ѻ���¼_IX_����ʱ�� On ����Ѻ���¼(����ʱ��) Pctfree 5 Tablespace zl9Indexhis;
Create Index ����Ѻ���¼_IX_��ת�� On ����Ѻ���¼(��ת��) Pctfree 5 Tablespace zl9Indexhis;
Create Index ���ý������_IX_����ID on ���ý������(����ID) Tablespace zl9Indexhis;
Create Index ���ý������_IX_��ת�� On ���ý������(��ת��) Tablespace Zl9indexhis;

Create Index �������׼�¼_IX_����ʱ�� On �������׼�¼(����ʱ��) Tablespace zl9Indexhis;

Create Index �ѱ���ϸ_IX_�շ�ϸĿid On �ѱ���ϸ(�ѱ�, �շ�ϸĿid) Tablespace zl9Indexhis;
Create Index �շѷ���Ŀ¼_IX_�ϼ�ID On �շѷ���Ŀ¼(�ϼ�ID) Tablespace zl9Indexhis;
Create Index �շ���ĿĿ¼_IX_����ID On �շ���ĿĿ¼(����ID) Tablespace zl9Indexhis;
Create Index �շ���Ŀ����_IX_���� On �շ���Ŀ����(����) Tablespace zl9Indexhis;
Create Index �շ���Ŀ����_IX_���� On �շ���Ŀ����(����) Tablespace zl9Indexhis;
Create Index �շ�ִ�п���_IX_��������ID On �շ�ִ�п���(��������ID) Tablespace zl9Indexhis;
Create Index �շ�ִ�п���_IX_ִ�п���ID On �շ�ִ�п���(ִ�п���ID) Tablespace zl9Indexhis;
Create Index �շѼ�Ŀ_IX_�շ�ϸĿid On �շѼ�Ŀ(�շ�ϸĿid) Tablespace zl9Indexhis;
Create Index �շѼ�Ŀ_IX_�۸�ȼ� on �շѼ�Ŀ(�۸�ȼ�) Tablespace zl9Indexhis;
Create Index �շѼ�Ŀ_IX_�䶯ԭ�� On �շѼ�Ŀ(�䶯ԭ��) Tablespace zl9Indexhis;
Create Index ������Ŀ����_IX_���� On ������Ŀ����(����) Tablespace zl9Indexhis;
Create Index �����շ���Ŀ_IX_ƴ�� On �����շ���Ŀ(ƴ��) Tablespace zl9Indexhis;
Create Index �����շ���Ŀ_IX_��� On �����շ���Ŀ(���) Tablespace zl9Indexhis;
Create Index �����շ���Ŀ_IX_����ID On �����շ���Ŀ(����ID) Tablespace zl9Indexhis;

Create Index �ҺŰ���_IX_ִ�мƻ�ID On �ҺŰ���(ִ�мƻ�ID) Tablespace zl9Indexhis;
Create Index �ҺŰ��żƻ�_IX_����ʱ�� On �ҺŰ��żƻ�(����ʱ��) Tablespace zl9Indexhis;
Create Index �ҺŰ��żƻ�_IX_���ʱ�� On �ҺŰ��żƻ�(���ʱ��) Tablespace zl9Indexhis;
Create Index �ҺŰ��żƻ�_IX_��Чʱ�� On �ҺŰ��żƻ�(��Чʱ��) Tablespace zl9Indexhis;
Create Index �ҺŰ��żƻ�_IX_ʧЧʱ�� On �ҺŰ��żƻ�(ʧЧʱ��) Tablespace zl9Indexhis;
Create Index �ҺŰ��żƻ�_IX_ʵ����Ч On �ҺŰ��żƻ�(ʵ����Ч) Tablespace zl9Indexhis;
Create Index �ҺŰ��żƻ�_IX_����ID On �ҺŰ��żƻ�(����ID) Tablespace zl9Indexhis;
Create Index �ҺŰ��żƻ�_IX_�ϴμƻ�ID on �ҺŰ��żƻ�(�ϴμƻ�ID) Tablespace zl9Indexhis;
Create Index �ҺŰ���ͣ��״̬_IX_��ʼʱ�� On �ҺŰ���ͣ��״̬(��ʼֹͣʱ��) Tablespace zl9Indexhis;
Create Index �ҺŰ���ͣ��״̬_IX_����ʱ�� On �ҺŰ���ͣ��״̬(����ֹͣʱ��) Tablespace zl9Indexhis;
Create Index �����˷�ԭ��_IX_���� On �����˷�ԭ��(����) Tablespace zl9Indexhis;

Create Index ���÷���ԭ��_IX_���� On ���÷���ԭ��(����) Tablespace zl9Indexhis;
Create Index ���ѿ���Ϣ_IX_������� On ���ѿ���Ϣ(�������) Tablespace zl9Indexhis;
Create Index ���ѿ���Ϣ_IX_��Ч�� On ���ѿ���Ϣ(��Ч��) Tablespace zl9Indexhis;
Create Index ���ѿ���Ϣ_IX_����ʱ�� On ���ѿ���Ϣ(����ʱ��) Tablespace zl9Indexhis;
Create Index ���ѿ���Ϣ_IX_����ʱ�� On ���ѿ���Ϣ(����ʱ��) Tablespace zl9Indexhis;
Create Index ���ѿ���Ϣ_IX_��ǰ״̬ On ���ѿ���Ϣ(��ǰ״̬) Tablespace zl9Indexhis;
Create Index ���ѿ���Ϣ_IX_ͣ������ On ���ѿ���Ϣ(ͣ������) Tablespace zl9Indexhis;
Create Index ���ѿ���Ϣ_Ix_����id On ���ѿ���Ϣ(����id) Tablespace Zl9indexhis;
Create Index ���ѿ���Ϣ_Ix_����id On ���ѿ���Ϣ(����id) Tablespace Zl9indexhis;

----------------------------------------------------------------------------
--[[5.ҩƷ���Ļ���]]
----------------------------------------------------------------------------
Create Index ҩƷ�洢�ⷿ_IX_�ⷿID On ҩƷ�洢�ⷿ(�ⷿID) Tablespace zl9Indexhis;
Create Index ҩƷ�洢�ⷿ_IX_����ID On ҩƷ�洢�ⷿ(����ID) Tablespace zl9Indexhis;
Create Index ҩƷ���_IX_ҩ��ID On ҩƷ���(ҩ��ID) Tablespace zl9Indexhis;
Create Index ҩƷ���_IX_��ʶ�� On ҩƷ���(��ʶ��) Tablespace zl9Indexhis;
Create Index ҩƷ���_IX_������Ӧ��ID On ҩƷ���(������Ӧ��ID) Tablespace zl9Indexhis;
Create Index ��������_IX_����ID On ��������(����ID) Tablespace zl9Indexhis;
Create Index ����������Ϣ_IX_��ҳID On ����������Ϣ(����ID,��ҳID) Tablespace zl9Indexhis;
Create Index ����������;_IX_���� On ����������;(����) Tablespace zl9Indexhis;
Create Index ��Ӧ��_IX_�ϼ�ID On ��Ӧ��(�ϼ�ID) Tablespace zl9Indexhis;
Create Index ��Ӧ��_IX_���� On ��Ӧ��(����) Tablespace zl9Indexhis;
Create Index �շѼ�Ŀ_IX_���ۻ��ܺ� On �շѼ�Ŀ(���ۻ��ܺ�) Tablespace zl9Indexhis;
Create Index ���������ӡ��¼_IX_���ʱ�� on ���������ӡ��¼(���ʱ��) Tablespace zl9Indexhis;
Create Index ���������ӡ��¼_IX_����id on ���������ӡ��¼(����id) Tablespace zl9Indexhis;
Create Index ҩƷ�����չ��Ϣ_IX_��Ŀ On ҩƷ�����չ��Ϣ(��Ŀ) Tablespace zl9Indexhis;
Create Index ҩƷ�ⷿ��λ_IX_�ⷿid on ҩƷ�ⷿ��λ(�ⷿid) Tablespace zl9Indexhis;
Create Index ҩƷ�ⷿ��λ_IX_�ϼ�id on ҩƷ�ⷿ��λ(�ϼ�id) Tablespace zl9Indexhis;
Create Index ҩƷ��λ����_IX_ҩƷID on ҩƷ��λ����(ҩƷID) Tablespace zl9Indexhis;
Create Index ҩƷ��λ����_IX_��λID on ҩƷ��λ����(��λID) Tablespace zl9Indexhis;
Create Index ���������_IX_�����id on ���������(�����id) Tablespace zl9Indexhis;
Create Index ���������_IX_�����id on ���������(�����id) Tablespace zl9Indexhis;
Create Index ҩƷ���Ŷ���_IX_��Ӧ��ID on ҩƷ���Ŷ���(��Ӧ��ID) Tablespace zl9Indexhis;

----------------------------------------------------------------------------
--[[6.�ٴ�����]]
----------------------------------------------------------------------------
Create Index ҽ�����Ӱ��¼_IX_�Ӱ࿪ʼʱ�� On ҽ�����Ӱ��¼(�Ӱ࿪ʼʱ��)  Tablespace zl9Indexhis;

Create Index ҽ�����Ӱ��¼_IX_�Ӱ����ʱ�� On ҽ�����Ӱ��¼(�Ӱ����ʱ��)  Tablespace zl9Indexhis;
Create Index ҽ�����Ӱ�����_IX_����ID On ҽ�����Ӱ�����(����ID,��ҳID) Tablespace zl9Indexhis;

Create Index ҽ�����Ӱ�ǩ��_IX_��¼ID On ҽ�����Ӱ�ǩ��(��¼ID)  Tablespace zl9Indexhis;

Create Index ���ﳣ������_IX_�ϼ� on ���ﳣ������(�ϼ�) Tablespace zl9Indexhis;

Create Index ����Ự��_IX_����ʱ�� On ����Ự��(����ʱ��)  Tablespace zl9Indexhis;

Create Index ����Ự��_IX_������ On ����Ự��(������)  Tablespace zl9Indexhis;
Create Index ����Ự��_IX_����ID On ����Ự��(����ID,����ID) Tablespace zl9Indexhis;
Create Index ������Ϣ��_IX_�Ķ�ʱ�� On ������Ϣ��(�Ķ�ʱ��)  Tablespace zl9Indexhis;

Create Index ������Ϣ��_IX_������ On ������Ϣ��(������)  Tablespace zl9Indexhis;
Create Index ������Ϣ��_IX_�Ựid On ������Ϣ��(�Ựid)  Tablespace zl9Indexhis;
Create Index ֤�ͷ�������_IX_����ID on ֤�ͷ�������(����ID) Tablespace zl9indexhis;

Create Index ��������_IX_��ҩID on ��������(��ҩID) Tablespace zl9indexhis;

Create Index ��֢�η�_IX_��֢ID on ��֢�η�(��֢ID) Tablespace zl9indexhis;

Create Index ��֢��ҩ_IX_��ҩID on ��֢��ҩ(��ҩID) Tablespace zl9indexhis;

Create Index ҽ��ִ�����_IX_��ת�� On ҽ��ִ�����(��ת��) Tablespace zl9Indexcis;

Create Index ҽ��ִ�����_Ix_Ҫ��ʱ�� On ҽ��ִ�����(Ҫ��ʱ��) Pctfree 5 Tablespace Zl9indexcis;

Create Index ������ҽ������¼_IX_����ID on ������ҽ������¼(����ID) Tablespace zl9indexhis;
Create Index ������ҽ������¼_IX_�巨ID on ������ҽ������¼(HIS�巨ID) Tablespace zl9indexhis;
Create Index ������ҽ������¼_IX_�÷�ID on ������ҽ������¼(HIS�÷�ID) Tablespace zl9indexhis;
Create Index ������ҽ������¼_IX_ҩ��ID on ������ҽ������¼(HISҩ��ID) Tablespace zl9indexhis;

Create Index ������ҽ��ϼ�¼_IX_����ID on ������ҽ��ϼ�¼(����ID) Tablespace zl9indexhis;
Create Index ������ҽ��ϼ�¼_IX_����ID on ������ҽ��ϼ�¼(����ID) Tablespace zl9indexhis;
Create Index ������ҽ��ϼ�¼_IX_����ID on ������ҽ��ϼ�¼(����ID) Tablespace zl9indexhis;
Create Index ������ҽ��ϼ�¼_IX_֤��ID on ������ҽ��ϼ�¼(֤��ID) Tablespace zl9indexhis;

Create Index ������ҽ��ϼ�¼_IX_�Һŵ� on ������ҽ��ϼ�¼(�Һŵ�) Tablespace zl9indexhis;
Create Index ������ҽ��ϼ�¼_IX_����� on ������ҽ��ϼ�¼(�����) Tablespace zl9indexhis;
Create Index ������ҽ��ϼ�¼_IX_HIS���ID on ������ҽ��ϼ�¼(HIS���ID) Tablespace zl9indexhis;
Create Index ������ҽ��ϼ�¼_IX_HISҽ��ID on ������ҽ��ϼ�¼(HISҽ��ID) Tablespace zl9indexhis;
Create Index ������ҽ��ϼ�¼_IX_����ʱ�� on ������ҽ��ϼ�¼(����ʱ��) Tablespace zl9indexhis;
Create Index ������ҽ������ϸ_IX_������ĿID on ������ҽ������ϸ(HISƷ��ID) Tablespace zl9indexhis;
Create Index ������ҽ������ϸ_IX_��ҩID on ������ҽ������ϸ(��ҩID) Tablespace zl9indexhis;

Create Index ������ҽ������ϸ_IX_���ID on ������ҽ������ϸ(HIS���ID) Tablespace zl9indexhis;
Create Index ��ҩĿ¼_IX_���� on ��ҩĿ¼(����) Tablespace zl9indexhis;

Create Index ��ҩĿ¼_IX_HISƷ��ID on ��ҩĿ¼(HISƷ��ID) Tablespace zl9indexhis;
Create Index ��ҽ֤��_IX_����ID on ��ҽ֤��(����ID) Tablespace zl9indexhis;
Create Index ��ҽ����_IX_�Ʊ� on ��ҽ����(�Ʊ�) Tablespace zl9indexhis;

Create Index ���Ӳ�����������_IX_������ on ���Ӳ�����������(������) Tablespace zl9indexhis;

Create Index ���Ӳ�����������_IX_����ʱ�� on ���Ӳ�����������(����ʱ��) Tablespace zl9indexhis;
Create Index ���Ӳ�����������_IX_����״̬ on ���Ӳ�����������(����״̬) Tablespace zl9indexhis;
Create Index ���Ӳ���������Ȩ_IX_��Ȩ�� on ���Ӳ���������Ȩ(��Ȩ��) Tablespace zl9indexhis;

Create Index ���Ӳ���������Ȩ_IX_����ID on ���Ӳ���������Ȩ(����ID) Tablespace zl9indexhis;
Create Index ���Ӳ���������Ȩ_IX_��Ȩʱ�� on ���Ӳ���������Ȩ(��Ȩʱ��) Tablespace zl9indexhis;
Create Index ���Ӳ���������Ȩ_IX_��ʼʱ�� on ���Ӳ���������Ȩ(���ʿ�ʼʱ��) Tablespace zl9indexhis;
Create Index ���Ӳ���������Ȩ_IX_����ʱ�� on ���Ӳ���������Ȩ(���ʽ���ʱ��) Tablespace zl9indexhis;
Create Index ���Ӳ���������־_IX_������ on ���Ӳ���������־(������) Tablespace zl9indexhis;

Create Index ���Ӳ���������־_IX_����ʱ�� on ���Ӳ���������־(����ʱ��) Tablespace zl9indexhis;
Create Index ���Ӳ�����Ȩ���ʲ���_IX_��Ȩid on ���Ӳ�����Ȩ���ʲ���(��Ȩid) Tablespace zl9indexhis;

Create Index ���Ӳ���������ʲ���_IX_����ID on ���Ӳ���������ʲ���(����ID) Tablespace zl9indexhis;

Create Index ���Ӳ�����Ȩ������Ա_IX_��Ȩid on ���Ӳ�����Ȩ������Ա(��Ȩid) Tablespace zl9indexhis;

Create Index RIS���ԤԼ_IX_��ת�� On RIS���ԤԼ(��ת��) Tablespace zl9Indexcis;
Create Index RIS���ԤԼ_IX_ԤԼ��ʼʱ�� On RIS���ԤԼ(ԤԼ��ʼʱ��) Tablespace zl9Indexcis;
Create Index RIS���ԤԼ_IX_ԤԼ���� On RIS���ԤԼ(ԤԼ����) Tablespace zl9Indexcis;
Create Index ҽ���������_IX_����ID On ҽ���������(����ID) Tablespace zl9Indexcis;
Create Index ҽ���������_IX_���ID On ҽ���������(���ID) Tablespace zl9Indexcis;
Create Index ҽ������ҽ��_IX_����ID On ҽ������ҽ��(����ID) Tablespace zl9Indexcis;
Create Index ҽ������ҽ��_IX_���ID On ҽ������ҽ��(���ID) Tablespace zl9Indexcis;
Create Index ҽ������ҽ��_IX_������ĿID On ҽ������ҽ��(������ĿID) Tablespace zl9Indexcis;
Create Index ҽ������ҽ��_IX_ҩƷID On ҽ������ҽ��(ҩƷID) Tablespace zl9Indexcis;
Create Index ҽ������ҽ��_IX_��ԱID On ҽ������ҽ��(��ԱID) Tablespace zl9Indexcis;
Create Index ��Ѫ�������_IX_������Ŀid On ��Ѫ�������(������Ŀid) Tablespace zl9Indexhis;
Create Index ����ҩ�������¼_IX_����ʱ�� On ����ҩ�������¼(����ʱ��) Tablespace zl9Indexcis;
Create Index ����ҩ�������ϸ_IX_����ID On ����ҩ�������ϸ(����ID,��ҳID) Tablespace zl9Indexcis;
Create Index ����ҩ�������ϸ_IX_�ٴ�֢״ On ����ҩ�������ϸ(�ٴ�֢״) Tablespace zl9Indexcis;
Create Index ����ҩ�������ϸ_IX_��Ⱦ��� On ����ҩ�������ϸ(��Ⱦ���) Tablespace zl9Indexcis;
Create Index ����ҩ���������_IX_����ID On ����ҩ���������(����ID) Tablespace zl9Indexcis;
Create Index ���Ʒ���Ŀ¼_IX_�ϼ�ID On ���Ʒ���Ŀ¼(�ϼ�ID) Tablespace zl9Indexhis;
Create Index ������ĿĿ¼_IX_����ID On ������ĿĿ¼(����ID) Tablespace zl9Indexhis;
Create Index ������Ŀ����_IX_���� On ������Ŀ����(����) Tablespace zl9Indexhis;
Create Index ������Ŀ����_IX_���� On ������Ŀ����(����) Tablespace zl9Indexhis;
Create Index ����ִ�п���_IX_��������ID On ����ִ�п���(��������ID) Tablespace zl9Indexcis;
Create Index ����ִ�п���_IX_ִ�п���ID On ����ִ�п���(ִ�п���ID) Tablespace zl9Indexcis;
Create Index ������Ŀ���_IX_������ĿID On ������Ŀ���(������ĿID) Tablespace zl9Indexcis;
Create Index ������Ŀ���_IX_�䷽ID On ������Ŀ���(�䷽ID) Tablespace zl9Indexcis;
Create Index ������Ŀ��λ_IX_��λ on ������Ŀ��λ(��λ,����) Tablespace zl9indexhis;
Create Index �����շѹ�ϵ_IX_�շ���ĿID On �����շѹ�ϵ(�շ���Ŀid) Tablespace zl9Indexcis;
Create Index ��Ա����ҩ��Ȩ��_Ix_��Աid On ��Ա����ҩ��Ȩ��(��Աid) Tablespace Zl9Indexhis;
Create Index ��Ա����Ȩ������_IX_������ĿID On ��Ա����Ȩ������(������ĿID) Tablespace zl9Indexhis;
Create Index ��Ա����Ȩ������_IX_��Ȩ��ԱID On ��Ա����Ȩ������(��Ȩ��ԱID) Tablespace zl9Indexhis;
Create Index ��Ա����Ȩ��_IX_������ĿID On ��Ա����Ȩ��(������ĿID) Tablespace zl9Indexhis;
Create Index ���þ���ժҪ_IX_��ԱID On ���þ���ժҪ(��ԱID) Tablespace zl9Indexhis;

----------------------------------------------------------------------------
--[[7.�ٴ�·������]]
----------------------------------------------------------------------------
Create Index �ٴ�·����Ŀ_IX_�汾�� On �ٴ�·����Ŀ(·��ID,�汾��) Tablespace zl9Indexcis;
Create Index �ٴ�·����Ŀ_IX_�׶�ID On �ٴ�·����Ŀ(�׶�ID) Tablespace zl9Indexcis;
Create Index �ٴ�·����Ŀ_IX_ͼ��ID On �ٴ�·����Ŀ(ͼ��ID) Tablespace zl9Indexcis;
Create Index ·��ҽ������_IX_���ID On ·��ҽ������(���ID) Tablespace zl9Indexcis;
Create Index ·��ҽ������_IX_������ĿID On ·��ҽ������(������ĿID) Tablespace zl9Indexcis;
Create Index ·��ҽ������_IX_�շ�ϸĿID On ·��ҽ������(�շ�ϸĿID) Tablespace zl9Indexcis;
Create Index ·��ҽ������_IX_ִ�п���ID On ·��ҽ������(ִ�п���ID) Tablespace zl9Indexcis;
Create Index ·��ҽ������_IX_�䷽ID On ·��ҽ������(�䷽ID) Tablespace zl9Indexcis;
Create Index �ٴ�·����֧_IX_ǰһ�׶�ID On �ٴ�·����֧(ǰһ�׶�ID) Tablespace zl9Indexhis;
Create Index �ٴ�·���׶�_IX_��֧ID On �ٴ�·���׶�(��֧ID) Tablespace zl9Indexhis;
Create Index �ٴ�·���׶�_IX_��ID On �ٴ�·���׶�(��ID) Tablespace zl9Indexcis;
Create Index �ٴ�·������_IX_��֧ID On �ٴ�·������(��֧ID) Tablespace zl9Indexhis;
Create Index �ٴ�·����Ŀ_IX_��֧ID On �ٴ�·����Ŀ(��֧ID) Tablespace zl9Indexhis;
Create Index �ٴ�·������_IX_��֧ID On �ٴ�·������(��֧ID) Tablespace zl9Indexhis;
Create Index �ٴ�·������_IX_�׶�ID On �ٴ�·������(�׶�ID) Tablespace zl9Indexcis;
Create Index ·����������_IX_����ID On ·����������(����ID) Tablespace zl9Indexcis;
Create Index ·����������_IX_��ĿID On ·����������(��ĿID) Tablespace zl9Indexcis;
Create Index �ٴ�·��ҽ��_IX_ҽ������ID On �ٴ�·��ҽ��(ҽ������ID) Tablespace zl9Indexcis;

----------------------------------------------------------------------------
--[[8.��������]]
----------------------------------------------------------------------------
Create Index ����������Ŀ_IX_����ID On ����������Ŀ(����ID) Tablespace zl9Indexcis;
Create Index ������ٴʾ�_IX_�ʾ����ID On ������ٴʾ�(�ʾ����ID) Tablespace zl9Indexcis;
Create Index ���������ϵ_IX_���ID On ���������ϵ(���ID) Tablespace zl9Indexcis;
Create Index ����Ӧ�ÿ���_IX_����ID On ����Ӧ�ÿ���(����ID) Tablespace zl9Indexcis;
Create Index ��������ǰ��_IX_����ID On ��������ǰ��(����ID) Tablespace zl9Indexcis;
Create Index ��������ǰ��_IX_���ID On ��������ǰ��(���ID) Tablespace zl9Indexcis;
Create Index ��������Ӧ��_IX_�����ļ�ID On ��������Ӧ��(�����ļ�ID) Tablespace zl9Indexcis;
Create Index ��������ģ��_IX_�����ļ�Id On ��������ģ��(�����ļ�Id,���ݸ���) Tablespace zl9Indexhis;
Create Index �����ļ��ṹ_IX_��ID On �����ļ��ṹ(��ID) Tablespace zl9Indexcis;
Create Index �����ļ��ṹ_IX_Ԥ�����ID On �����ļ��ṹ(Ԥ�����ID) Tablespace zl9Indexcis;
Create Index �����ļ��ṹ_IX_����Ҫ��ID On �����ļ��ṹ(����Ҫ��ID) Tablespace zl9Indexcis;
Create Index �����ʾ�ʾ��_IX_����id On �����ʾ�ʾ��(����id) Tablespace zl9Indexcis;
Create Index �����ʾ�ʾ��_IX_��Աid On �����ʾ�ʾ��(��Աid) Tablespace zl9Indexcis;
Create Index �����ʾ�ʾ��_IX_��� On �����ʾ�ʾ��(���) Tablespace zl9Indexcis;
Create Index �����ʾ�ʾ��_IX_���� On �����ʾ�ʾ��(����) Tablespace zl9Indexcis;
Create Index �����ʾ����_IX_�����ı� On �����ʾ����(�����ı�) Tablespace zl9Indexcis;
Create Index ��������Ŀ¼_IX_����id On ��������Ŀ¼(����id) Tablespace zl9Indexcis;
Create Index ��������Ŀ¼_IX_��Աid On ��������Ŀ¼(��Աid) Tablespace zl9Indexcis;
Create Index ������������_IX_��ID On ������������(��ID) Tablespace zl9Indexcis;
Create Index ������������_IX_Ԥ�����ID On ������������(Ԥ�����ID) Tablespace zl9Indexcis;
Create Index ������������_IX_����Ҫ��ID On ������������(����Ҫ��ID) Tablespace zl9Indexcis;

Create Index ����������_IX_�ϼ�id On ����������(�ϼ�id) Tablespace zl9Indexcis;
Create Index ����������_IX_����id On ����������(����id) Tablespace zl9Indexcis;
Create Index �������Ŀ¼_IX_����id On �������Ŀ¼(����id) Tablespace zl9Indexcis;
----------------------------------------------------------------------------
--[[9.�������]]
----------------------------------------------------------------------------
Create Index �����ص����_IX_�ϼ���� On �����ص����(�ϼ����) Tablespace zl9Indexcis;
Create Index �������ÿ���_IX_����ID On �������ÿ���(����ID) Tablespace zl9Indexcis;
----------------------------------------------------------------------------
--[[10.�������]]
----------------------------------------------------------------------------
Create Index ����ϸ��_IX_���� On ����ϸ��(����) Tablespace zl9Indexcis;
Create Index �����Լ���ϵ_IX_����id On �����Լ���ϵ(����id) Tablespace zl9Indexcis;
Create Index ���鱸ע����_IX_���� On ���鱸ע����(����) Tablespace zl9Indexcis;
Create Index ������������_IX_���� On ������������(����) Tablespace zl9Indexcis;
Create Index ���鱨����Ŀ_IX_ϸ��ID On ���鱨����Ŀ(ϸ��id) Tablespace zl9Indexcis;
Create Index ���鱨����Ŀ_IX_������ĿID On ���鱨����Ŀ(������ĿID) Tablespace zl9Indexcis;

Create Index ����ģ��Ŀ¼_IX_������ĿID On ����ģ��Ŀ¼(������ĿID) Tablespace zl9Indexcis;
Create Index ����ģ������_IX_ģ��ID On ����ģ������(ģ��ID) Tablespace zl9Indexcis;
Create Index ����ģ������_IX_��ĿID On ����ģ������(��ĿID) Tablespace zl9Indexcis;
Create Index ����ģ������_IX_ϸ��ID On ����ģ������(ϸ��ID) Tablespace zl9Indexcis;
Create Index ����ģ��ҩ��_IX_������ID On ����ģ��ҩ��(������ID) Tablespace zl9Indexcis;
Create Index ����ϲ�����_IX_����ĿID On ����ϲ�����(����ĿID) Tablespace zl9Indexcis;
Create Index ����ϲ�����_IX_�ϲ���ĿID On ����ϲ�����(�ϲ���ĿID) Tablespace zl9Indexcis;

Create Index ��������_IX_ʹ��С��ID On ��������(ʹ��С��ID) Tablespace zl9Indexcis;
Create Index ��������������_IX_������ID On ����������Ŀ(������id) Tablespace zl9Indexcis;
Create Index ��������������_IX_��ĿID On ����������Ŀ(��Ŀid) Tablespace zl9Indexcis;
Create Index ��������״̬_IX_��ĿID On ��������״̬(��ĿID) Tablespace zl9Indexcis;
Create Index ������������_IX_�ϼ�ID On ������������(�ϼ�ID) Tablespace zl9Indexcis;
Create Index ������������_IX_����ID On ������������(����ID) Tablespace zl9Indexcis;
Create Index ������������_IX_����ID On ������������(����ID) Tablespace zl9Indexcis;

----------------------------------------------------------------------------
--[[11.������]]
----------------------------------------------------------------------------
Create Index RIS�ӿ���־��¼_IX_ʱ�� On RIS�ӿ���־��¼(ʱ��) Tablespace zl9Indexcis;
Create Index ��������¼_IX_�������ID On ��������¼(�������ID) Tablespace zl9Indexcis;
Create Index ��������¼_IX_�� On ��������¼(��) Pctfree 5 Tablespace zl9Indexcis;
Create Index Ӱ���ѯ����_IX_�������� On Ӱ���ѯ����(��������) Tablespace zl9Indexhis;
Create Index Ӱ���ѯ����_IX_����ID On Ӱ���ѯ����(����ID) Tablespace zl9Indexhis;
Create Index ��ݹ�����Ϣ_IX_ģ��� On ��ݹ�����Ϣ(ģ���,��Ŀ) Tablespace zl9Indexhis;
create index ҽ��ִ�з���_IX_����ID on ҽ��ִ�з���(����ID) Tablespace zl9Indexhis;
create index Ӱ��������_IX_����ID on Ӱ��������(����ID) Tablespace zl9Indexhis;
create index Ӱ��ִ�з���_IX_����ID on Ӱ��ִ�з���(����ID) Tablespace zl9Indexhis;
Create Index Ӱ�����볣�ôʾ�_IX_����ID On Ӱ�����볣�ôʾ�(����ID) Tablespace zl9Indexhis;
Create Index Ӱ�����볣�ôʾ�_IX_������ԱID On Ӱ�����볣�ôʾ�(������ԱID) Tablespace zl9Indexhis;
----------------------------------------------------------------------------
--[[12.ҽ��ҵ��]]
----------------------------------------------------------------------------
Create Index ҽ�����˵���_IX_����ʱ�� On ҽ�����˵���(����ʱ��) Pctfree 5 Tablespace zl9Indexhis;
Create Index ҽ�����˹�����_IX_����ID On ҽ�����˹�����(����ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index ����������Ŀ_IX_��ĿID On ����������Ŀ(��ĿID) Tablespace zl9Indexhis;

----------------------------------------------------------------------------
--[[13.���˲���ҵ��]]
----------------------------------------------------------------------------
Create Index ������������¼_IX_Ӥ������ID On ������������¼(Ӥ������ID,Ӥ����ҳID) Pctfree 5 Tablespace zl9Indexhis;

Create Index ���ò�����Ϊԭ��_IX_���� On ���ò�����Ϊԭ��(����) Tablespace zl9Indexhis;
Create Index ������Ϊ����_IX_��Ϊ��� On ������Ϊ����(��Ϊ���) Tablespace zl9Indexhis;
Create Index ���˲�����¼_IX_����ʱ�� On ���˲�����¼(����ʱ��) Tablespace zl9Indexhis;
Create Index ���˲�����¼_IX_����ʱ�� On ���˲�����¼(����ʱ��) Tablespace zl9Indexhis;
Create Index ���˲�����¼_IX_����ʱ�� On ���˲�����¼(����ʱ��) Tablespace zl9Indexhis;
Create Index ���˲�����¼_IX_����ID On ���˲�����¼(����ID) Tablespace zl9Indexhis;
Create Index ���˲�����¼_IX_��Ϊ��� On ���˲�����¼(��Ϊ���) Tablespace zl9Indexhis;
Create Index �����Զ�����_IX_����ID On �����Զ�����(����ID,��ҳID) Pctfree 5 Tablespace zl9Indexhis;
Create Index �����Զ�����_IX_��ʼʱ�� On �����Զ�����(��ʼʱ��) Pctfree 5 Tablespace zl9Indexhis;
Create Index �����Զ�����_IX_��ֹʱ�� On �����Զ�����(��ֹʱ��) Pctfree 5 Tablespace zl9Indexhis;
Create Index ��λ������¼_IX_����ID On ��λ������¼(����ID) Tablespace zl9Indexhis;
Create Index ��λ״����¼_IX_����ID On ��λ״����¼(����ID) Tablespace zl9Indexhis;
Create Index ��λ״����¼_IX_����ID On ��λ״����¼(����ID) Tablespace zl9Indexhis;

Create Index ������Ϣ_IX_���� On ������Ϣ(����) Pctfree 5 Tablespace zl9Indexhis;
Create Index ������Ϣ_IX_�Ǽ�ʱ�� On ������Ϣ(�Ǽ�ʱ��) Pctfree 5 Tablespace zl9Indexhis;
Create Index ������Ϣ_IX_���֤�� On ������Ϣ(���֤��) Pctfree 5 Tablespace zl9Indexhis;
Create Index ������Ϣ_IX_IC���� On ������Ϣ(IC����) Pctfree 5 Tablespace zl9Indexhis;
Create Index ������Ϣ_IX_ҽ���� On ������Ϣ(ҽ����) Pctfree 5 Tablespace zl9Indexhis;
Create Index ������Ϣ_IX_��ͬ��λid On ������Ϣ(��ͬ��λid) Pctfree 5 Tablespace zl9Indexhis;
Create Index ������Ϣ_IX_��Ժ On ������Ϣ(��Ժ) Pctfree 5 Tablespace zl9Indexhis;
Create Index ������Ϣ_IX_�ֻ��� on ������Ϣ(�ֻ���) Tablespace zl9Indexhis;
Create Index ������Ϣ_IX_��ǰ����ID On ������Ϣ(��ǰ����ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index ������Ϣ_IX_��ϵ�����֤�� On ������Ϣ(��ϵ�����֤�� ) Pctfree 5 Tablespace zl9Indexhis;
Create Index ������ݹ���_IX_����ID On ������ݹ���(����ID) Tablespace zl9Indexhis;
Create Index ��Ժ����_IX_����ID On ��Ժ����(����ID,��ҳID) Tablespace zl9Indexhis;
Create Index ���˺ϲ���¼_IX_����ID On ���˺ϲ���¼(����id) Tablespace zl9Indexhis;
Create Index ���˺ϲ���¼_IX_ԭ����id On ���˺ϲ���¼(ԭ����id) Tablespace Zl9indexhis;
Create Index ���˵�����¼_IX_��ҳID On ���˵�����¼(����ID,��ҳID) Tablespace zl9Indexhis;
Create Index ���˱䶯��¼_IX_����ID On ���˱䶯��¼(����ID,��ҳID) Pctfree 5 Tablespace zl9Indexhis;
Create Index ���˱䶯��¼_IX_ҽ��С��ID On ���˱䶯��¼(ҽ��С��ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index ���˱䶯��¼_IX_��ʼʱ�� On ���˱䶯��¼(��ʼʱ��) Pctfree 5 Tablespace zl9Indexhis;
Create Index ���˱䶯��¼_IX_��ֹʱ�� On ���˱䶯��¼(��ֹʱ��) Pctfree 5 Tablespace zl9Indexhis;
Create Index ������ҳ_IX_��Ժ���� On ������ҳ(��Ժ����) Pctfree 5 Tablespace zl9Indexhis;
Create Index ������ҳ_IX_��Ժ���� On ������ҳ(��Ժ����) Pctfree 5 Tablespace zl9Indexhis;
Create Index ������ҳ_IX_ҽ��С��ID On ������ҳ(ҽ��С��ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index ������ҳ_IX_סԺ�� On ������ҳ(סԺ��) Pctfree 5 Tablespace zl9Indexhis;
Create Index ������ҳ_IX_������ On ������ҳ(������) Pctfree 5 Tablespace zl9Indexhis;
Create Index ������ҳ_IX_���ۺ� On ������ҳ(���ۺ�) Tablespace zl9Indexhis;
Create Index ������ҳ_IX_��ת�� On ������ҳ(��ת��) Tablespace zl9Indexhis;
Create Index ������ҳ_IX_�Һ�ID On ������ҳ(�Һ�ID)  Tablespace zl9Indexcis;

Create Index ���˼���_IX_����ID On ���˼���(����ID) Tablespace zl9Indexhis;

Create Index סԺ������¼_IX_����ID On סԺ������¼(������) PCTFREE 5 Tablespace zl9Indexhis;
Create Index סԺ������¼_IX_������ On סԺ������¼(������) PCTFREE 5 Tablespace zl9Indexhis;

Create Index ���˹�����¼_IX_����ID On ���˹�����¼(����ID,��ҳID) Tablespace zl9Indexcis;
Create Index ���˹�����¼_IX_��ת�� On ���˹�����¼(��ת��) Tablespace zl9Indexcis;
Create Index ������ϼ�¼_IX_����ID On ������ϼ�¼(����ID,��ҳID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ������ϼ�¼_IX_ҽ��id On ������ϼ�¼(ҽ��id) Pctfree 5 Tablespace zl9Indexhis;
Create Index ������ϼ�¼_IX_����ID On ������ϼ�¼(����ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ������ϼ�¼_IX_����ID On ������ϼ�¼(����ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ������ϼ�¼_IX_��ת�� On ������ϼ�¼(��ת��) Tablespace zl9Indexcis;
Create Index ������ϼ�¼_IX_����id On ������ϼ�¼(����id) Pctfree 5 Tablespace zl9Indexcis;
Create Index ������ϼ�¼_IX_���id On ������ϼ�¼(���id) Pctfree 5 Tablespace zl9Indexcis;
Create Index ������ϼ�¼_IX_֤��id On ������ϼ�¼(֤��id) Pctfree 5 Tablespace zl9Indexcis;
Create Index �������ҽ��_IX_ҽ��ID On �������ҽ��(ҽ��ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index �������ҽ��_IX_��ת�� On �������ҽ��(��ת��) Tablespace zl9Indexcis;
Create Index ���������¼_IX_��ҳID On ���������¼(����ID,��ҳID ) Tablespace zl9Indexcis;
Create Index ���������¼_IX_��ת�� On ���������¼(��ת��) Tablespace zl9Indexcis;
Create Index ���˿����ؼ�¼_IX_ҩ��id On ���˿����ؼ�¼(ҩ��id) Tablespace zl9Indexcis;

Create Index �������Ƽ�¼_IX_��ʼ���� On �������Ƽ�¼(��ʼ����) PCTFREE 5 Tablespace zl9Indexcis;
Create Index �������Ƽ�¼_IX_�������� On �������Ƽ�¼(��������) PCTFREE 5 Tablespace zl9Indexcis;
Create Index �������Ƽ�¼_IX_��ʼ���� On �������Ƽ�¼(��ʼ����) PCTFREE 5 Tablespace zl9Indexcis;
Create Index �������Ƽ�¼_IX_�������� On �������Ƽ�¼(��������) PCTFREE 5 Tablespace zl9Indexcis;
Create Index ������������_IX_�Ǽ�ʱ�� On ������������(ҩ������) PCTFREE 5 Tablespace zl9Indexcis;

----------------------------------------------------------------------------
--[[14.����ҵ��]]
----------------------------------------------------------------------------
Create Index �����ID_IX_�����ID on �����˿���Ϣ(�����ID) Tablespace zl9Indexhis;
Create Index �����˿���Ϣ_IX_�����ID On �����˿���Ϣ(�����ID) Tablespace Zl9indexhis;

Create Index Ԥ���������_IX_����ID on Ԥ���������(����ID) Tablespace zl9Indexhis;

Create Index ���ѿ���Ϣ_Ix_����id On ���ѿ���Ϣ(����id) Tablespace Zl9indexhis;
Create Index ���ѿ���Ϣ_Ix_����id On ���ѿ���Ϣ(����id) Tablespace Zl9indexhis;

Create Index ���ѿ�����¼_Ix_�Ǽ��� On ���ѿ�����¼(�Ǽ���) Tablespace Zl9indexhis;
Create Index ���ѿ�����¼_Ix_�Ǽ�ʱ�� On ���ѿ�����¼(�Ǽ�ʱ��) Tablespace Zl9indexhis;
Create Index ���ѿ�����¼_Ix_�Ƿ���ڿ� On ���ѿ�����¼(�Ƿ���ڿ�) Tablespace Zl9indexhis;
Create Index ���ѿ�����¼_IX_���� On ���ѿ�����¼(����) Tablespace zl9Indexhis;

Create Index ���ѿ����ü�¼_Ix_������ On ���ѿ����ü�¼(������) Tablespace Zl9indexhis;
Create Index ���ѿ����ü�¼_Ix_���� On ���ѿ����ü�¼(����) Tablespace Zl9indexhis;
Create Index ���ѿ����ü�¼_Ix_�Ǽ�ʱ�� On ���ѿ����ü�¼(�Ǽ�ʱ��) Tablespace Zl9indexhis;
Create Index ���ѿ����ü�¼_IX_���ID On ���ѿ����ü�¼(���ID) Tablespace zl9Indexhis;

Create Index ���ѿ������¼_Ix_���id On ���ѿ������¼(���id) Tablespace Zl9indexhis;
Create Index ���ѿ������¼_Ix_������ On ���ѿ������¼(������) Tablespace Zl9indexhis;
Create Index ���ѿ������¼_Ix_����ʱ�� On ���ѿ������¼(����ʱ��) Tablespace Zl9indexhis;

Create Index ���ѿ�ʹ�ü�¼_Ix_����id On ���ѿ�ʹ�ü�¼(����id, ����) Tablespace Zl9indexhis;
Create Index ���ѿ�ʹ�ü�¼_Ix_ʹ��ʱ�� On ���ѿ�ʹ�ü�¼(ʹ��ʱ��) Tablespace Zl9indexhis;

Create Index �ʻ��ɿ����_Ix_������� On �ʻ��ɿ����(�������) Tablespace Zl9indexhis;

Create Index ���˽ɿ��¼_IX_����ID On ���˽ɿ��¼(����ID) Tablespace zl9indexhis;
Create Index ���˽ɿ��¼_IX_�Ǽ�ʱ�� On ���˽ɿ��¼(�Ǽ�ʱ��) Tablespace zl9indexhis;
Create Index ���˽ɿ����_IX_����Id On ���˽ɿ����(����Id) Tablespace zl9indexhis;

Create Index ���ñ䶯��¼_Ix_Ŀ��䶯id On ���ñ䶯��¼(Ŀ��䶯id) Tablespace Zl9indexhis;
Create Index ���ñ䶯��¼_IX_��ת�� On ���ñ䶯��¼(��ת��) Tablespace Zl9indexhis;

Create Index ���ñ䶯��¼_Ix_�շ�ϸĿid On ���ñ䶯��¼(�շ�ϸĿid) Tablespace Zl9indexhis;
Create Index ���ñ䶯��¼_Ix_����id On ���ñ䶯��¼(����id) Tablespace Zl9indexhis;
Create Index ���ñ䶯��¼_Ix_����id On ���ñ䶯��¼(����id, ��ҳid) Tablespace Zl9indexhis;
Create Index ���˷�����Ϣ��¼_IX_�Ǽ�ʱ�� on ���˷�����Ϣ��¼(�Ǽ�ʱ��) Tablespace zl9Indexhis;

Create Index ���˷�����Ϣ��¼_IX_����ʱ�� on ���˷�����Ϣ��¼(����ʱ��) Tablespace zl9Indexhis;
Create Index ���˷�����Ϣ��¼_IX_����ID on ���˷�����Ϣ��¼(����ID) Tablespace zl9Indexhis;
Create Index ���˷�����Ϣ��¼_IX_�Һ�ID on ���˷�����Ϣ��¼(�Һ�ID) Tablespace zl9Indexhis;
Create Index ���˷�����Ϣ��¼_IX_����ID on ���˷�����Ϣ��¼(����) Tablespace zl9Indexhis;
Create Index ���˷�����Ϣ��¼_IX_��ԴID on ���˷�����Ϣ��¼(��ԴID) Tablespace zl9Indexhis;
Create Index ���˷�����Ϣ��¼_IX_��¼ID on ���˷�����Ϣ��¼(��¼ID) Tablespace zl9Indexhis;
Create Index ���˷�����Ϣ��¼_IX_��ĿID on ���˷�����Ϣ��¼(��ĿID) Tablespace zl9Indexhis;
Create Index ���˷�����Ϣ��¼_IX_ҽ��ID on ���˷�����Ϣ��¼(ҽ��ID) Tablespace zl9Indexhis;
Create Index �ٴ�����䶯��¼_IX_��¼ID on �ٴ�����䶯��¼(��¼ID) Tablespace zl9Indexhis;

Create Index �ٴ�����䶯��¼_IX_�Ǽ�ʱ�� on �ٴ�����䶯��¼(�Ǽ�ʱ��) Tablespace zl9Indexhis;
Create Index �ٴ�����䶯��¼_IX_ԭ����ID on �ٴ�����䶯��¼(ԭ����ID) Tablespace zl9Indexhis;
Create Index �ٴ�����䶯��¼_IX_������ID on �ٴ�����䶯��¼(������ID) Tablespace zl9Indexhis;
Create Index �ٴ�����䶯��ϸ_IX_����ID on �ٴ�����䶯��ϸ(����ID) Tablespace zl9Indexhis;

Create Index �ٴ�����ͣ���¼_IX_��¼ID on �ٴ�����ͣ���¼(��¼ID) Tablespace zl9Indexhis;

Create Index �ٴ�����ͣ���¼_IX_����ҽ��ID on �ٴ�����ͣ���¼(����ҽ��ID) Tablespace zl9Indexhis;
Create Index �ٴ�����ͣ���¼_IX_����ʱ�� on �ٴ�����ͣ���¼(����ʱ��) Tablespace zl9Indexhis;
Create Index �ٴ�����ͣ���¼_IX_����ʱ�� on �ٴ�����ͣ���¼(����ʱ��) Tablespace zl9Indexhis;
Create Index �����������ÿ���_IX_����id on �����������ÿ���(����id) Tablespace zl9Indexhis;

Create Index �ٴ������¼_IX_����ID on �ٴ������¼(����ID) Tablespace zl9Indexhis;
Create Index �ٴ������¼_IX_����ID on �ٴ������¼(����ID) Tablespace zl9Indexhis;
Create Index �ٴ������¼_IX_��ԴID on �ٴ������¼(��ԴID) Tablespace zl9Indexhis;
Create Index �ٴ������¼_IX_����ҽ��id on �ٴ������¼(����ҽ��id) Tablespace zl9Indexhis;
Create Index �ٴ������¼_IX_ҽ��id on �ٴ������¼(ҽ��id) Tablespace zl9Indexhis;
Create Index �ٴ������¼_IX_��ĿID on �ٴ������¼(��ĿID) Tablespace zl9Indexhis;
Create Index �ٴ������¼_IX_����ID on �ٴ������¼(����ID) Tablespace zl9Indexhis;
Create Index �ٴ������¼_IX_���ID On �ٴ������¼(���ID) Tablespace zl9Indexhis;
Create Index �ٴ������_IX_����ID on �ٴ������(����ID) Tablespace zl9Indexhis;
Create Index �ٴ������_IX_����ID On �ٴ������(����ID) Tablespace zl9Indexhis;
Create Index �ٴ����ﰲ��_IX_��ĿID on �ٴ����ﰲ��(��ĿID) Tablespace zl9Indexhis;
Create Index �ٴ����ﰲ��_IX_ҽ��id on �ٴ����ﰲ��(ҽ��id) Tablespace zl9Indexhis;
Create Index �ٴ����ﰲ��_IX_��ԴID on �ٴ����ﰲ��(��ԴID) Tablespace zl9Indexhis;
Create Index �ٴ����ﰲ��_IX_����ID on �ٴ����ﰲ��(����ID) Tablespace zl9Indexhis;
Create Index �ٴ���������_IX_����ID on �ٴ���������(����ID) Tablespace zl9Indexhis;
Create Index �ٴ��������Ҽ�¼_IX_����ID on �ٴ��������Ҽ�¼(����ID) Tablespace zl9Indexhis;
Create Index �ٴ���������_IX_����ID on �ٴ���������(����ID) Tablespace zl9Indexhis;
create Index �ٴ������Դ����_IX_����ID on �ٴ������Դ����(����ID) Tablespace zl9Indexhis;
create Index �ٴ������Դ����_IX_����ID on �ٴ������Դ����(����ID) Tablespace zl9Indexhis;
Create Index �ٴ������Դ_IX_��ĿID on �ٴ������Դ(��ĿID) Tablespace zl9Indexhis;

Create Index �ٴ������Դ_IX_ҽ��id on �ٴ������Դ(ҽ��id) Tablespace zl9Indexhis;
Create Index �ٴ������Դ_IX_ҽ������ on �ٴ������Դ(ҽ������) Tablespace zl9Indexhis;
Create Index �����˿���Ϣ_IX_��ת�� On �����˿���Ϣ(��ת��) Tablespace zl9Indexhis;
Create Index �����˿���Ϣ_IX_��¼id On �����˿���Ϣ(��¼id) Tablespace Zl9indexhis;
Create Index �����嵥��ӡ_IX_��ҳID On �����嵥��ӡ(����ID,��ҳID) Tablespace zl9Indexhis;

Create Index �����嵥��ӡ_IX_��ӡʱ�� On �����嵥��ӡ(��ӡʱ��) Tablespace zl9Indexhis;
Create Index �����嵥��ӡ_IX_��ת�� On �����嵥��ӡ(��ת��) Tablespace zl9Indexhis;
Create Index ҽ��������ϸ_Ix_��ת�� On ҽ��������ϸ(��ת��) Tablespace Zl9indexhis;
Create Index ҽ��������ϸ_IX_NO On ҽ��������ϸ(NO) Pctfree 5 Tablespace zl9Indexhis;
Create Index ҽ��������ϸ_IX_�����ID On ҽ��������ϸ(�����ID) Tablespace zl9Indexhis;

Create Index ����䶯��¼_IX_�Ǽ�ʱ�� On ����䶯��¼(�Ǽ�ʱ��) Tablespace zl9indexhis;
Create Index ����䶯��¼_IX_����ID On ����䶯��¼(����ID) Tablespace zl9indexhis;

Create Index ���ò����¼_IX_����ID On ���ò����¼(����ID) Tablespace zl9indexhis;
  Create Index ���ò����¼_Ix_�ɿ���id On ���ò����¼(�ɿ���id) Tablespace Zl9indexhis;

Create Index ���ò����¼_IX_�շѽ���ID On ���ò����¼(�շѽ���ID) Tablespace zl9indexhis;
Create Index ���ò����¼_IX_������� On ���ò����¼(�������) Tablespace zl9indexhis;
Create Index ���ò����¼_IX_����״̬ On ���ò����¼(����״̬) Tablespace zl9indexhis;
Create Index ���ò����¼_IX_��ת�� On ���ò����¼(��ת��) Tablespace zl9indexhis;
Create Index ���ò����¼_IX_�Ǽ�ʱ�� On ���ò����¼(�Ǽ�ʱ��) Tablespace zl9indexhis;
Create Index ���ò����¼_IX_����id On ���ò����¼(����id) Tablespace zl9indexhis;
Create Index ƾ����ӡ��¼_IX_��ת�� On ƾ����ӡ��¼(��ת��) Tablespace zl9Indexhis;
Create Index �������㽻��_IX_��ת�� On �������㽻��(��ת��) Tablespace zl9Indexhis;
Create Index �������㽻��_IX_ԭԤ��id On �������㽻��(ԭԤ��id) Tablespace zl9Indexhis;
Create Index ���˿������¼_IX_����ʱ�� On ���˿������¼(����ʱ��) Pctfree 5 Tablespace zl9Indexhis;
Create Index ���˿������¼_IX_��ת�� On ���˿������¼(��ת��) Tablespace zl9Indexhis;
Create Index ���˿������¼_Ix_����id On ���˿������¼(����id) Tablespace Zl9indexhis;
Create Index ���˿������¼_Ix_������� On ���˿������¼(�������) Tablespace Zl9indexhis;
Create Index ���˿������¼_Ix_������� On ���˿������¼(�������) Tablespace Zl9indexhis;
Create Index ���˿������¼_Ix_����id On ���˿������¼(����id) Tablespace Zl9indexhis;
Create Index ���˿������¼_Ix_�Ǽ�ʱ�� On ���˿������¼(�Ǽ�ʱ��) Tablespace Zl9indexhis;
Create Index ����ҽ�ƿ��䶯_IX_�䶯ID On ����ҽ�ƿ��䶯(�䶯ID) Tablespace zl9Indexhis;
Create Index ����ҽ�ƿ��䶯_IX_���� On ����ҽ�ƿ��䶯(����,�����ID,�䶯ʱ��) Tablespace zl9Indexhis;
Create Index ����ҽ�ƿ��䶯_IX_���õ��� On ����ҽ�ƿ��䶯(���õ���) Pctfree 5 Tablespace zl9Indexhis;
Create Index ����ҽ�ƿ���Ϣ_IX_��ʧʱ�� On ����ҽ�ƿ���Ϣ(��ʧʱ��) Tablespace zl9Indexhis;
Create Index ����ҽ�ƿ���Ϣ_IX_�������� On ����ҽ�ƿ���Ϣ(��������) Tablespace zl9Indexhis;
Create Index ����ҽ�ƿ���Ϣ_IX_��ֹʹ��ʱ�� on ����ҽ�ƿ���Ϣ(��ֹʹ��ʱ��) Tablespace zl9Indexhis;
Create Index ����ҽ�ƿ���Ϣ_IX_��ά�� On ����ҽ�ƿ���Ϣ(��ά��) Pctfree 5 Tablespace zl9Indexhis;

Create Index ���˹ҺŻ���_IX_���� On ���˹ҺŻ���(����) Tablespace zl9Indexhis;
Create Index ���˹ҺŻ���_IX_��ĿID On ���˹ҺŻ���(��ĿID) Tablespace zl9Indexhis;
Create Index ���˹ҺŻ���_IX_��ת�� On ���˹ҺŻ���(��ת��) Tablespace zl9Indexhis;
Create Index ���˹Һż�¼_IX_����ID On ���˹Һż�¼(����ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index ���˹Һż�¼_IX_����ʱ�� On ���˹Һż�¼(����ʱ��) Tablespace zl9Indexhis;
Create Index ���˹Һż�¼_IX_�Ǽ�ʱ�� On ���˹Һż�¼(�Ǽ�ʱ��) Pctfree 5 Tablespace zl9Indexhis;
Create Index ���˹Һż�¼_IX_ԤԼʱ�� On ���˹Һż�¼(ԤԼʱ��) Pctfree 5 Tablespace zl9Indexhis;
Create Index ���˹Һż�¼_IX_����ʱ�� On ���˹Һż�¼(����ʱ��) Pctfree 5 Tablespace zl9Indexhis;
Create Index ���˹Һż�¼_IX_ִ��ʱ�� On ���˹Һż�¼(ִ��ʱ��) Pctfree 5 Tablespace zl9Indexhis;
Create Index ���˹Һż�¼_IX_ִ��״̬ On ���˹Һż�¼(ִ��״̬) Pctfree 5 Tablespace zl9Indexhis;
Create Index ���˹Һż�¼_IX_��ת�� On ���˹Һż�¼(��ת��) Tablespace zl9Indexhis;
Create Index ���˹Һż�¼_IX_�����¼ID on ���˹Һż�¼(�����¼ID) Tablespace zl9Indexhis;
Create Index ���˹Һż�¼_IX_�Һ���ĿID on ���˹Һż�¼(�Һ���ĿID) Tablespace zl9IndexHis;
Create Index �Һ����״̬_IX_���� On �Һ����״̬(����) Initrans 20 Tablespace zl9Indexhis;
Create Index �Һ����״̬_IX_�Ǽ�ʱ�� On �Һ����״̬(�Ǽ�ʱ��) Initrans 20 Tablespace zl9indexhis;
Create Index �Һ����״̬_IX_���� On �Һ����״̬(����) Initrans 20 Tablespace zl9Indexhis;
Create Index ����ת���¼_IX_��ת�� On ����ת���¼(��ת��) Tablespace zl9Indexhis;

Create Index ��Ա�սɼ�¼_IX_�տ�Ա On ��Ա�սɼ�¼(�տ�Ա) Tablespace zl9Indexhis;
Create Index ��Ա�սɼ�¼_IX_�ɿ���ID On ��Ա�սɼ�¼(�ɿ���ID) Tablespace zl9Indexhis;
Create Index ��Ա�սɼ�¼_IX_С���տ�ID On ��Ա�սɼ�¼(С���տ�ID) Tablespace zl9Indexhis;
Create Index ��Ա�սɼ�¼_IX_С������ID On ��Ա�սɼ�¼(С������ID) Tablespace zl9Indexhis;
Create Index ��Ա�սɼ�¼_IX_�����տ�ID On ��Ա�սɼ�¼(�����տ�ID) Tablespace zl9Indexhis;
Create Index ��Ա�սɼ�¼_IX_����ʱ�� On ��Ա�սɼ�¼(����ʱ��) Tablespace zl9Indexhis;
Create Index ��Ա�սɼ�¼_IX_�Ǽ�ʱ�� On ��Ա�սɼ�¼(�Ǽ�ʱ��) Tablespace zl9Indexhis;
Create Index ��Ա�սɼ�¼_IX_��ת�� On ��Ա�սɼ�¼(��ת��) Tablespace zl9Indexhis;
Create Index ��Ա�ս���ϸ_IX_��ת�� On ��Ա�ս���ϸ(��ת��) Tablespace zl9Indexhis;
Create Index ��Ա�ս�Ʊ��_IX_��ת�� On ��Ա�ս�Ʊ��(��ת��) Tablespace zl9Indexhis;
Create Index ��Ա�սɶ���_IX_��¼ID On ��Ա�սɶ���(��¼ID, ����) Tablespace zl9Indexhis;
Create Index ��Ա�սɶ���_IX_��ת�� On ��Ա�սɶ���(��ת��) Tablespace zl9Indexhis;
Create Index ��Ա�ݴ��¼_IX_�ս�ID On ��Ա�ݴ��¼(�ս�ID) Tablespace zl9Indexhis;
Create Index ��Ա�ݴ��¼_IX_�ջ�ʱ�� On ��Ա�ݴ��¼(�ջ�ʱ��) Tablespace zl9Indexhis;
Create Index ��Ա�ݴ��¼_IX_�Ǽ�ʱ�� On ��Ա�ݴ��¼(�Ǽ�ʱ��) Tablespace zl9Indexhis;
Create Index ��Ա�ݴ��¼_IX_����ʱ�� On ��Ա�ݴ��¼(����ʱ��) Tablespace zl9Indexhis;
Create Index ��Ա�ݴ��¼_IX_�տ�Ա On ��Ա�ݴ��¼(�տ�Ա) Tablespace zl9Indexhis;
Create Index ��Ա�ݴ��¼_IX_��ת�� On ��Ա�ݴ��¼(��ת��) Tablespace zl9Indexhis;
Create Index ��Ա����¼_IX_����� On ��Ա����¼(�����) Tablespace zl9Indexhis;
Create Index ��Ա����¼_IX_����ʱ�� On ��Ա����¼(����ʱ��) Tablespace zl9Indexhis;
Create Index ��Ա����¼_IX_����� On ��Ա����¼(�����) Tablespace zl9Indexhis;
Create Index ��Ա����¼_IX_���ʱ�� On ��Ա����¼(���ʱ��) Tablespace zl9Indexhis;
Create Index ��Ա����¼_IX_��ת�� On ��Ա����¼(��ת��) Tablespace zl9Indexhis;
Create Index ���˴߿��¼_IX_����ID On ���˴߿��¼(����ID,��ҳID) Pctfree 5 Tablespace zl9Indexhis;
Create Index ���˴߿��¼_IX_��ӡ���� On ���˴߿��¼(��ӡ����) Pctfree 5 Tablespace zl9Indexhis;
Create Index �������鳤����_IX_�鳤ID On �������鳤����(�鳤ID) Pctfree 5 Tablespace zl9indexhis;

Create Index ���˽��ʼ�¼_IX_�շ�ʱ�� On ���˽��ʼ�¼(�շ�ʱ��) Pctfree 5 Tablespace zl9Indexhis;
Create Index ���˽��ʼ�¼_IX_����id On ���˽��ʼ�¼(����id) Pctfree 5 Tablespace zl9Indexhis;
Create Index ���˽��ʼ�¼_IX_��ת�� On ���˽��ʼ�¼(��ת��) Tablespace zl9Indexhis;
Create Index ���˽��ʼ�¼_IX_����״̬ On ���˽��ʼ�¼(����״̬) Tablespace zl9indexhis;
Create Index סԺ���ü�¼_IX_�շ�ϸĿid On סԺ���ü�¼(�շ�ϸĿid) Pctfree 5 Tablespace zl9Indexhis;
Create Index סԺ���ü�¼_IX_������Ŀid On סԺ���ü�¼(������Ŀid) Pctfree 5 Tablespace zl9Indexhis;
Create Index סԺ���ü�¼_IX_ҽ����� On סԺ���ü�¼(ҽ�����) Pctfree 5 Tablespace zl9Indexhis;
Create Index סԺ���ü�¼_IX_����ID On סԺ���ü�¼(����ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index סԺ���ü�¼_IX_�Ǽ�ʱ�� On סԺ���ü�¼(�Ǽ�ʱ��) Pctfree 5 Tablespace zl9Indexhis;
Create Index סԺ���ü�¼_IX_����ʱ�� On סԺ���ü�¼(����ʱ��) Pctfree 5 Tablespace zl9Indexhis;
Create Index סԺ���ü�¼_IX_����id On סԺ���ü�¼(����id,��ҳID) Pctfree 5 Tablespace zl9Indexhis;
Create Index סԺ���ü�¼_IX_���մ���ID On סԺ���ü�¼(���մ���ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index סԺ���ü�¼_IX_��ת�� On סԺ���ü�¼(��ת��) Tablespace zl9Indexhis;

Create Index ������ü�¼_IX_�շ�ϸĿid On ������ü�¼(�շ�ϸĿid) Pctfree 5 Tablespace zl9Indexhis;
Create Index ������ü�¼_IX_������Ŀid On ������ü�¼(������Ŀid) Pctfree 5 Tablespace zl9Indexhis;
Create Index ������ü�¼_IX_ҽ����� On ������ü�¼(ҽ�����) Pctfree 5 Tablespace zl9Indexhis;
Create Index ������ü�¼_IX_����ID On ������ü�¼(����ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index ������ü�¼_IX_�Ǽ�ʱ�� On ������ü�¼(�Ǽ�ʱ��) Pctfree 5 Tablespace zl9Indexhis;
Create Index ������ü�¼_IX_����ʱ�� On ������ü�¼(����ʱ��) Pctfree 5 Tablespace zl9Indexhis;
Create Index ������ü�¼_IX_����id On ������ü�¼(����id,��ҳid) Pctfree 5 Tablespace zl9Indexhis;
Create Index ������ü�¼_IX_���մ���ID On ������ü�¼(���մ���ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index ������ü�¼_IX_�Һ�ID On ������ü�¼(�Һ�ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index ������ü�¼_IX_��ת�� On ������ü�¼(��ת��) Tablespace zl9Indexhis;

Create Index ���˷�������_IX_����ʱ�� On ���˷�������(����ʱ��) Tablespace zl9Indexhis;
Create Index ���˷�������_IX_�˲����� On ���˷�������(�˲�����) Tablespace zl9Indexhis;
Create Index ���˷�������_IX_��ת�� On ���˷�������(��ת��) Tablespace Zl9indexhis;
Create Index ������˼�¼_IX_����ID On ������˼�¼(����ID,��ҳID) Tablespace zl9Indexhis;
Create Index ������˼�¼_IX_������� On ������˼�¼(�������) Tablespace zl9Indexhis;
Create Index ���˷�������_IX_���ʱ�� On ���˷�������(���ʱ��) Tablespace zl9Indexhis;
Create Index ���˷�������_IX_״̬ On ���˷�������(״̬) Tablespace zl9Indexhis;
Create Index �����˷�����_IX_����ʱ�� On �����˷�����(����ʱ��) Tablespace zl9Indexhis;
Create Index �����˷�����_IX_���ʱ�� On �����˷�����(���ʱ��) Tablespace zl9Indexhis;
Create Index ���˷��û���_IX_������Ŀid On ���˷��û���(������Ŀid) Tablespace zl9Indexhis;
Create Index ���˽��ʻ���_IX_����ID On ���˽��ʻ���(����ID) Tablespace zl9Indexhis;
Create Index ���˽��ʻ���_IX_������Ŀid On ���˽��ʻ���(������Ŀid) Tablespace zl9Indexhis;
Create Index ���˽��ʻ���_IX_����id On ���˽��ʻ���(����id,��ҳid) Tablespace zl9Indexhis;
Create Index ҽ���������_IX_ִ���� On ҽ���������(����,ִ����) Tablespace zl9Indexhis;
Create Index ҽ���������_IX_������Ŀid On ҽ���������(������Ŀid) Tablespace zl9Indexhis;
Create Index ����δ�����_IX_����id On ����δ�����(����id,��ҳID) Tablespace zl9Indexhis;
Create Index ����δ�����_IX_������Ŀid On ����δ�����(������Ŀid) Tablespace zl9Indexhis;
Create Index ����Ԥ����¼_IX_��ҳID On ����Ԥ����¼(����ID,��ҳID) Pctfree 5 Tablespace zl9Indexhis;
Create Index ����Ԥ����¼_IX_����id On ����Ԥ����¼(����id) Pctfree 5 Tablespace zl9Indexhis;
Create Index ����Ԥ����¼_IX_�տ�ʱ�� On ����Ԥ����¼(�տ�ʱ��) Pctfree 5 Tablespace zl9Indexhis;
Create Index ����Ԥ����¼_IX_������� On ����Ԥ����¼(�������) Tablespace zl9Indexhis;
Create Index ����Ԥ����¼_IX_��ת�� On ����Ԥ����¼(��ת��) Tablespace zl9Indexhis;
Create Index ����Ԥ����¼_IX_����ʱ�� on ����Ԥ����¼(����ʱ��) Tablespace zl9Indexhis;
Create Index ����Ԥ����¼_IX_��������ID on ����Ԥ����¼(��������ID) Tablespace zl9Indexhis;

Create Index Ʊ������¼_IX_�Ǽ��� On Ʊ������¼(�Ǽ���) Pctfree 5 Tablespace zl9Indexhis;
Create Index Ʊ������¼_IX_�Ǽ�ʱ�� On Ʊ������¼(�Ǽ�ʱ��) Pctfree 5 Tablespace zl9Indexhis;
Create Index Ʊ������¼_IX_����Ʊ�� On Ʊ������¼(����Ʊ��) Pctfree 5 Tablespace zl9Indexhis;
Create Index Ʊ������¼_IX_���� On Ʊ������¼(����) Tablespace zl9Indexhis;
Create Index Ʊ�ݱ����¼_IX_������ On Ʊ�ݱ����¼(������) Pctfree 5 Tablespace zl9Indexhis;
Create Index Ʊ�ݱ����¼_IX_����ʱ�� On Ʊ�ݱ����¼(����ʱ��) Pctfree 5 Tablespace zl9Indexhis;
Create Index Ʊ�ݱ����¼_IX_���ID On Ʊ�ݱ����¼(���ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index Ʊ�����ü�¼_IX_������ On Ʊ�����ü�¼(������) Pctfree 5 Tablespace zl9Indexhis;
Create Index Ʊ�����ü�¼_IX_���� On Ʊ�����ü�¼(����,Ʊ��) Pctfree 5 Tablespace zl9Indexhis;
Create Index Ʊ�����ü�¼_IX_�Ǽ�ʱ�� On Ʊ�����ü�¼(�Ǽ�ʱ��) Pctfree 5 Tablespace zl9Indexhis;
Create Index Ʊ�����ü�¼_IX_��ת�� On Ʊ�����ü�¼(��ת��) Tablespace zl9Indexhis;
Create Index Ʊ�����ü�¼_IX_���ID On Ʊ�����ü�¼(���ID) Tablespace zl9Indexhis;
Create Index Ʊ�ݴ�ӡ����_IX_NO On Ʊ�ݴ�ӡ����(NO) Pctfree 5 Tablespace zl9Indexhis;
Create Index Ʊ��ʹ����ϸ_IX_����ID On Ʊ��ʹ����ϸ(����ID,Ʊ��,����) Pctfree 5 Tablespace zl9Indexhis;
Create Index Ʊ��ʹ����ϸ_IX_ʹ��ʱ�� On Ʊ��ʹ����ϸ(ʹ��ʱ��) Pctfree 5 Tablespace zl9Indexhis;
Create Index Ʊ��ʹ����ϸ_IX_��ӡID On Ʊ��ʹ����ϸ(��ӡID) Pctfree 5 Tablespace zl9Indexhis;
Create Index Ʊ��ʹ����ϸ_IX_��ת�� On Ʊ��ʹ����ϸ(��ת��) Tablespace zl9Indexhis;
CREATE INDEX Ʊ��ʹ����ϸ_IX_����Ʊ��ID ON Ʊ��ʹ����ϸ(����Ʊ��ID) TABLESPACE zl9Indexhis;
Create Index Ʊ�ݴ�ӡ��ϸ_IX_ʹ��ID On Ʊ�ݴ�ӡ��ϸ(ʹ��ID,NO) Pctfree 5 Tablespace zl9Indexhis;
Create Index Ʊ�ݴ�ӡ��ϸ_IX_����Ʊ����� On Ʊ�ݴ�ӡ��ϸ(����Ʊ�����) Pctfree 5 Tablespace zl9Indexhis;
Create Index Ʊ�ݴ�ӡ��ϸ_IX_��ת�� On Ʊ�ݴ�ӡ��ϸ(��ת��) Tablespace zl9Indexhis;
Create Index Ʊ�ݴ�ӡ����_IX_��ת�� On Ʊ�ݴ�ӡ����(��ת��) Tablespace zl9Indexhis;
Create Index �ɿ��Ա���_IX_��ԱID On �ɿ��Ա���(��ԱID) Pctfree 5 Tablespace zl9Indexhis;

----------------------------------------------------------------------------
--[[15.ҩƷ����ҵ��]]
---------------------------------------------------------------------------- 

Create Index δ��ҩƷ��¼_IX_ҩƷid On δ��ҩƷ��¼(ҩƷid) Tablespace zl9Indexhis;   
Create Index δ��ҩƷ��¼_IX_�������� On δ��ҩƷ��¼(��������) Tablespace zl9Indexhis;
Create Index δ��ҩƷ��¼_IX_��ת�� On δ��ҩƷ��¼(��ת��) Tablespace zl9Indexhis;  
Create Index ���Ͻ���¼_IX_�ϴν��id On ���Ͻ���¼(�ϴν��id) Tablespace zl9Indexhis;
Create Index ���Ͻ���¼_IX_�������� On ���Ͻ���¼(��������) Tablespace zl9Indexhis;
Create Index ���Ͻ���¼_IX_������� On ���Ͻ���¼(�������) Tablespace zl9Indexhis;

Create Index ���Ͻ����ϸ_IX_���id On ���Ͻ����ϸ(���id) Tablespace zl9Indexhis;
Create Index ���Ͻ����ϸ_IX_����id On ���Ͻ����ϸ(����id) Tablespace zl9Indexhis;

Create Index ���Ͻ�����_IX_���id On ���Ͻ�����(���id) Tablespace zl9Indexhis;
Create Index ���Ͻ�����_IX_����id On ���Ͻ�����(����id) Tablespace zl9Indexhis;

Create Index ��Һ��������_IX_��ҩ̨id On ��Һ��������(��ҩ̨id) Tablespace ZL9INDEXHIS;

Create Index ҩƷ�շ������־_IX_��ת�� ON ҩƷ�շ������־(��ת��) Tablespace Zl9indexhis;
Create Index ҩƷ�շ�סԺ��־_IX_��ת�� ON ҩƷ�շ�סԺ��־(��ת��) Tablespace Zl9indexhis;

Create Index ����������¼_IX_����id On ����������¼(����id) Tablespace zl9Indexhis;
Create Index ����������¼_IX_����id On ����������¼(����id) Tablespace zl9Indexhis;
Create Index ����������¼_IX_��ҩ��λid On ����������¼(��ҩ��λid) Tablespace zl9Indexhis;

Create Index ҩƷ���ռ�¼_IX_��ҩ��λid On ҩƷ���ռ�¼(��ҩ��λid) Tablespace zl9Indexhis;   
Create Index ҩƷ���ռ�¼_IX_NO On ҩƷ���ռ�¼(NO) Tablespace zl9Indexhis;  
Create Index ҩƷ������ϸ_IX_ҩƷid On ҩƷ������ϸ(ҩƷid) Tablespace zl9Indexhis;  

Create Index ҩƷ�÷�����_IX_�÷�id On ҩƷ�÷�����(�÷�ID) Tablespace zl9Indexhis;
Create Index ���ϴ����޶�_IX_����ID On ���ϴ����޶�(����ID) Tablespace zl9Indexhis;

Create Index �����������_IX_����ID ON �����������(����ID) Tablespace Zl9indexhis;
Create Index �����������_IX_ҽ��ID ON �����������(ҽ��ID) Tablespace Zl9indexhis;
Create Index �����������_IX_���ID ON �����������(���ID) Tablespace Zl9indexhis;
Create Index �����������_IX_����ID ON �����������(����ID) Tablespace Zl9indexhis;
Create Index �����������_IX_ҩ��ID ON �����������(ҩ��ID) Tablespace Zl9indexhis;
Create Index ��������¼_Ix_�Һ�id On ��������¼(�Һ�id) Tablespace Zl9indexhis;

Create Index ��������¼_Ix_����id On ��������¼(����id, ��ҳid) Tablespace Zl9indexhis;
Create Index ��������¼_Ix_���ʱ�� On ��������¼(���ʱ��) Tablespace Zl9indexhis;
Create Index ��������¼_Ix_״̬ On ��������¼(״̬) Tablespace Zl9indexhis;
Create Index ��������¼_Ix_�����û� On ��������¼(�����û�) Tablespace Zl9indexhis;
Create Index ��������¼_Ix_��ת�� On ��������¼(��ת��) Tablespace Zl9indexhis;
Create Index ���������ϸ_Ix_ҽ��id On ���������ϸ(ҽ��id) Tablespace Zl9indexhis;

Create Index ���������ϸ_IX_��ת�� ON ���������ϸ(��ת��) Tablespace Zl9indexhis;
Create Index ���������_Ix_�����Ŀid On ���������(�����Ŀid) Tablespace Zl9indexhis;
Create Index ���������_Ix_ҽ��id On ���������(ҽ��id) Tablespace Zl9indexhis;

Create Index ���������_IX_��ת�� ON ���������(��ת��) Tablespace Zl9indexhis;
Create Index �շѵ��ۼ�¼_IX_�շ�ϸĿid On �շѵ��ۼ�¼(�շ�ϸĿid) Tablespace zl9Indexhis;
Create Index �շѵ��ۼ�¼_IX_�۸�ȼ� on �շѵ��ۼ�¼(�۸�ȼ�) Tablespace zl9Indexhis;

Create Index �շѵ��ۼ�¼_IX_������Ŀid On �շѵ��ۼ�¼(������Ŀid) Tablespace zl9Indexhis;
Create Index �շѵ��ۼ�¼_IX_��˱�־ On �շѵ��ۼ�¼(��˱�־) Tablespace zl9Indexhis;
Create Index �շѵ��ۼ�¼_IX_�������� On �շѵ��ۼ�¼(��������) Tablespace zl9Indexhis;
Create Index ҩƷ�������_IX_������� On ҩƷ�������(�������) Tablespace zl9Indexcis;
Create Index ҩƷ�ɹ��ƻ�_IX_�������� On ҩƷ�ɹ��ƻ�(��������) Pctfree 5 Tablespace zl9Indexhis;
Create Index ҩƷ�ɹ��ƻ�_IX_������� On ҩƷ�ɹ��ƻ�(�������) Pctfree 5 Tablespace zl9Indexhis;
Create Index ҩƷ�ɹ��ƻ�_IX_�������� On ҩƷ�ɹ��ƻ�(��������) Pctfree 5 Tablespace zl9Indexhis;
Create Index ҩƷ�ɹ��ƻ�_IX_�ϲ��ƻ�id On ҩƷ�ɹ��ƻ�(�ϲ��ƻ�id) Tablespace zl9Indexhis;
Create Index ҩƷ�ƻ�����_IX_ҩƷid On ҩƷ�ƻ�����(ҩƷid) Tablespace zl9Indexhis;
Create Index ҩƷ��ҩ�ƻ�_IX_ҩƷid On ҩƷ��ҩ�ƻ�(ҩƷid) Tablespace zl9Indexhis;
Create Index ҩƷ��ҩ�ƻ�_IX_��ҩ��λid On ҩƷ��ҩ�ƻ�(��ҩ��λid) Tablespace zl9Indexhis;
Create Index ҩƷ��ҩ�ƻ�_IX_�������� On ҩƷ��ҩ�ƻ�(��������) Tablespace zl9Indexhis;
Create Index ҩƷ��ҩ�ƻ�_IX_������� On ҩƷ��ҩ�ƻ�(�������) Tablespace zl9Indexhis;
Create Index ���ϲɹ��ƻ�_IX_NO On ���ϲɹ��ƻ�(no) Tablespace zl9Indexhis;
Create Index ���ϼƻ�����_IX_����id On ���ϼƻ�����(����id) Tablespace zl9Indexhis;
Create Index ҩƷ����ƻ�_IX_״̬ On ҩƷ����ƻ�(����ID,״̬) Tablespace zl9Indexhis;
Create Index ҩƷ����ƻ�_IX_����ID On ҩƷ����ƻ�(����ID) Tablespace zl9Indexhis;
Create Index ҩƷ����ƻ�_IX_��ת�� On ҩƷ����ƻ�(��ת��) Tablespace zl9Indexhis;

Create Index ҩƷ���_IX_ҩƷid On ҩƷ���(ҩƷid) Tablespace zl9Indexhis;
Create Index ҩƷ���_IX_��Ʒ���� On ҩƷ���(��Ʒ����) Tablespace zl9Indexhis;
Create Index ҩƷ���_IX_�ڲ����� On ҩƷ���(�ڲ�����) Tablespace zl9Indexhis;
Create Index ҩƷ������_IX_ҩƷid On ҩƷ������(ҩƷid) Tablespace zl9Indexhis;
Create Index ҩƷ������_IX_���id On ҩƷ������(���id) Tablespace zl9Indexhis;
Create Index ҩƷ������_IX_�ⷿid On ҩƷ������(�ⷿid) Tablespace zl9Indexhis;
Create Index ҩƷ������_IX_������id On ҩƷ������(������id) Tablespace zl9Indexhis;
Create Index ҩƷ���_IX_ҩƷid On ҩƷ���(ҩƷid) Tablespace zl9Indexhis;
Create Index ҩƷ����_IX_ҩƷid On ҩƷ����(ҩƷid) Tablespace zl9Indexhis;
Create Index ҩƷ�շ�����_IX_ҩƷid On ҩƷ�շ�����(ҩƷid) Tablespace zl9Indexhis;
Create Index ҩƷ�շ�����_IX_���id On ҩƷ�շ�����(���id) Tablespace zl9Indexhis;
Create Index δ��ҩƷ��¼_IX_�������� On δ��ҩƷ��¼(��������) Tablespace zl9Indexhis;
Create Index δ��ҩƷ��¼_IX_�Է�����ID On δ��ҩƷ��¼(�Է�����ID) Tablespace zl9Indexhis;
Create Index δ��ҩƷ��¼_IX_��ҳID On δ��ҩƷ��¼(����ID,��ҳID) Tablespace zl9Indexhis;
Create Index δ��ҩƷ��¼_IX_�Ŷ�״̬ On δ��ҩƷ��¼(�Ŷ�״̬) Tablespace zl9Indexcis;
Create Index ҩƷ�շ���¼_IX_����id On ҩƷ�շ���¼(����id) Pctfree 5 Tablespace zl9Indexhis;
Create Index ҩƷ�շ���¼_IX_ҩƷid On ҩƷ�շ���¼(ҩƷid) Pctfree 5 Tablespace zl9Indexhis;
Create Index ҩƷ�շ���¼_IX_������id On ҩƷ�շ���¼(������id) Pctfree 5 Tablespace zl9Indexhis;
Create Index ҩƷ�շ���¼_IX_��ҩ��λid On ҩƷ�շ���¼(��ҩ��λid) Pctfree 5 Tablespace zl9Indexhis;
Create Index ҩƷ�շ���¼_IX_�������� On ҩƷ�շ���¼(��������) Pctfree 5 Tablespace zl9Indexhis;
Create Index ҩƷ�շ���¼_IX_������� On ҩƷ�շ���¼(�������) Pctfree 5 Tablespace zl9Indexhis;
Create Index ҩƷ�շ���¼_IX_�۸�ID On ҩƷ�շ���¼(�۸�ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index ҩƷ�շ���¼_IX_���ܷ�ҩ�� On ҩƷ�շ���¼(���ܷ�ҩ��) Pctfree 5 Tablespace zl9Indexhis;
Create Index ҩƷ�շ���¼_IX_��Ʒ���� On ҩƷ�շ���¼(��Ʒ����) Tablespace zl9Indexhis;
Create Index ҩƷ�շ���¼_IX_�ڲ����� On ҩƷ�շ���¼(�ڲ�����) Tablespace zl9Indexhis;
Create Index ҩƷ�շ���¼_IX_��ת�� On ҩƷ�շ���¼(��ת��) Tablespace zl9Indexhis;
Create Index ҩƷ�շ���¼_IX_�ƻ�id On ҩƷ�շ���¼(�ƻ�id) Tablespace zl9Indexhis;
Create Index ҩƷ�շ���¼_IX_����ID On ҩƷ�շ���¼(����ID,��ҳID) Tablespace zl9Indexhis;
Create Index ҩƷ�շ���¼_IX_ҽ��id On ҩƷ�շ���¼(ҽ��id) Pctfree 5 Tablespace zl9Indexhis;
Create Index �շ���¼������Ϣ_IX_��ת�� On �շ���¼������Ϣ(��ת��) Tablespace zl9Indexhis;

Create Index ҩƷǩ����¼_IX_֤��ID On ҩƷǩ����¼(֤��ID) Tablespace zl9Indexhis;
Create Index ҩƷǩ����¼_IX_��ת�� On ҩƷǩ����¼(��ת��) Tablespace zl9Indexhis;
Create Index ҩƷǩ����ϸ_IX_�շ�ID On ҩƷǩ����ϸ(�շ�ID) Tablespace zl9Indexhis;
Create Index ҩƷǩ����ϸ_IX_��ת�� On ҩƷǩ����ϸ(��ת��) Tablespace zl9Indexhis;

Create Index �ɱ��۵�����Ϣ_IX_ִ������ On �ɱ��۵�����Ϣ(ִ������) Tablespace zl9Indexhis;
Create Index �ɱ��۵�����Ϣ_IX_ҩƷID On �ɱ��۵�����Ϣ(ҩƷID) Tablespace zl9Indexhis;
Create Index �ɱ��۵�����Ϣ_IX_�շ�id On �ɱ��۵�����Ϣ(�շ�id) Tablespace zl9Indexhis;
Create Index �ɱ��۵�����Ϣ_IX_��ҩ��λID On �ɱ��۵�����Ϣ(��ҩ��λID) Tablespace zl9Indexhis;

Create Index ҩƷ�۸��¼_IX_ҩƷID On ҩƷ�۸��¼(ҩƷID) Tablespace zl9Indexhis;
Create Index ҩƷ�۸��¼_IX_ԭ��id On ҩƷ�۸��¼(ԭ��id) Tablespace zl9Indexhis;
Create Index ҩƷ�۸��¼_IX_�շ�id On ҩƷ�۸��¼(�շ�id) Tablespace zl9Indexhis;
Create Index ҩƷ�۸��¼_IX_��ҩ��λID On ҩƷ�۸��¼(��ҩ��λID) Tablespace zl9Indexhis;
Create Index ҩƷ�۸��¼_IX_���ۻ��ܺ� On ҩƷ�۸��¼(���ۻ��ܺ�) Tablespace zl9Indexhis;
Create Index ҩƷ�۸��¼_IX_��¼״̬ On ҩƷ�۸��¼(��¼״̬) Tablespace zl9Indexhis;

Create Index ҩƷ������¼_IX_ҩƷid On ҩƷ������¼(ҩƷid) Tablespace zl9Indexhis;
Create Index ҩƷ������¼_IX_��ҩ��λid On ҩƷ������¼(��ҩ��λid) Tablespace zl9Indexhis;
Create Index ҩƷ������¼_IX_�Ǽ�ʱ�� On ҩƷ������¼(�Ǽ�ʱ��) Tablespace zl9Indexhis;
Create Index ҩƷ������¼_IX_����ʱ�� On ҩƷ������¼(����ʱ��) Tablespace zl9Indexhis;

Create Index ҩƷ����¼_IX_�������� On ҩƷ����¼(��������) Tablespace zl9Indexhis;
Create Index ҩƷ����¼_IX_������� On ҩƷ����¼(�������) Tablespace zl9Indexhis;
Create Index ҩƷ�����ϸ_IX_ҩƷid On ҩƷ�����ϸ(ҩƷid) Tablespace zl9Indexhis;
Create Index ҩƷ������_IX_ҩƷid On ҩƷ������(ҩƷid) Tablespace zl9Indexhis;
Create Index ҩƷ������_IX_���id On ҩƷ������(���id) Tablespace zl9Indexhis;
Create Index �ݴ�ҩƷ��¼_IX_����ID On �ݴ�ҩƷ��¼(����ID) Tablespace zl9Indexhis;
Create Index �ݴ�ҩƷ��¼_IX_�Ǽ�ʱ�� On �ݴ�ҩƷ��¼(�Ǽ�ʱ��) Tablespace zl9Indexhis;
Create Index �ݴ�ҩƷ��¼_IX_ҽ��ID On �ݴ�ҩƷ��¼(ҽ��ID, ���ͺ�) Tablespace zl9Indexhis;

Create Index ��Һ��ҩ��¼_IX_ִ��ʱ�� On ��Һ��ҩ��¼(ִ��ʱ��) Tablespace zl9Indexhis;
Create Index ��Һ��ҩ��¼_IX_����ʱ�� On ��Һ��ҩ��¼(����ʱ��,����״̬) Tablespace zl9Indexhis;
Create Index ��Һ��ҩ��¼_IX_��ҩ���� On ��Һ��ҩ��¼(��ҩ����) Pctfree 20 Tablespace zl9Indexhis;
Create Index ��Һ��ҩ��¼_IX_ƿǩ�� On ��Һ��ҩ��¼(ƿǩ��) Tablespace zl9Indexhis;
Create Index ��Һ��ҩ��¼_IX_��ת�� On ��Һ��ҩ��¼(��ת��) Tablespace zl9Indexhis;
Create Index ��Һ��ҩ��¼_IX_��ӡʱ�� On ��Һ��ҩ��¼(��ӡʱ��) Tablespace zl9Indexcis;
Create Index ��Һ��ҩ��¼_IX_����ID On ��Һ��ҩ��¼(����ID,��ҳID) Tablespace zl9Indexhis;
Create Index ��Һ��ҩ��¼_IX_����ʱ�� On ��Һ��ҩ��¼(����ʱ��) Tablespace zl9Indexhis;

Create Index ��Һ��ҩ״̬_IX_����ʱ�� On ��Һ��ҩ״̬(����ʱ��,��������) Tablespace zl9Indexhis;
Create Index ��Һ��ҩ״̬_IX_��ת�� On ��Һ��ҩ״̬(��ת��) Tablespace zl9Indexhis;
Create Index ��Һ��ҩ����_IX_�շ�ID On ��Һ��ҩ����(�շ�ID) Tablespace zl9Indexhis;
Create Index ��Һ��ҩ����_IX_��ת�� On ��Һ��ҩ����(��ת��) Tablespace zl9Indexhis;
Create Index ��Һ��ҩ����_IX_��ת�� On ��Һ��ҩ����(��ת��) Tablespace zl9Indexhis;
Create Index ��Һ��ҩ����_IX_����id On ��Һ��ҩ����(����id) Tablespace zl9Indexcis;

Create Index �����շѷ���_IX_����ID On �����շѷ���(����ID) Tablespace zl9Indexhis;

Create Index Ӧ����¼_IX_�շ�ID On Ӧ����¼(�շ�ID) Tablespace zl9Indexhis;
Create Index Ӧ����¼_IX_��λID On Ӧ����¼(��λID) Tablespace zl9Indexhis;
Create Index Ӧ����¼_IX_������� On Ӧ����¼(�������) Tablespace zl9Indexhis;
Create Index Ӧ����¼_IX_������� On Ӧ����¼(�������) Tablespace zl9Indexhis;
Create Index Ӧ����¼_IX_��Ʊ�� On Ӧ����¼(��Ʊ��) Tablespace zl9Indexhis;
Create Index Ӧ����¼_IX_������� On Ӧ����¼(�������) Tablespace zl9Indexhis;
Create Index Ӧ����¼_IX_��ⵥ�ݺ� On Ӧ����¼(��ⵥ�ݺ�) Tablespace zl9Indexhis;
Create Index �����¼_IX_��λid On �����¼(��λid) Tablespace zl9Indexhis;
Create Index �����¼_IX_�������� On �����¼(��������) Tablespace zl9Indexhis;
Create Index �����¼_IX_Ԥ������ On �����¼(Ԥ������) Tablespace zl9Indexhis;
Create Index �����¼_IX_������� On �����¼(�������) Tablespace zl9Indexhis;
Create Index �����¼_IX_������� On �����¼(�������) Tablespace zl9Indexhis;

Create Index ���ۻ��ܼ�¼_IX_ִ������ On ���ۻ��ܼ�¼(ִ������) Tablespace zl9Indexhis;
Create Index ���ۻ��ܼ�¼_IX_�������� On ���ۻ��ܼ�¼(��������) Tablespace zl9Indexhis;
Create Index �ɱ��۵�����Ϣ_IX_���ۻ��ܺ� On �ɱ��۵�����Ϣ(���ۻ��ܺ�) Tablespace zl9Indexhis;

----------------------------------------------------------------------------
--[[16.�ٴ�ҽ��]]
----------------------------------------------------------------------------
Create Index ��������¼_IX_����ID On ��������¼(����ID) Tablespace zl9IndexCis;
Create Index ��������¼_IX_�Ǽ�ʱ�� On ��������¼(�Ǽ�ʱ��) Tablespace zl9IndexCis;
Create Index ��������¼_IX_��ת�� On ��������¼(��ת��) Tablespace zl9IndexCis;
Create Index ���ﲡ������ָ��_IX_��ת�� On ���ﲡ������ָ��(��ת��) Tablespace zl9IndexCis;
Create Index ���ﲡ������_IX_����ID On ���ﲡ������(����ID) Tablespace zl9IndexCis;
Create Index ���ﲡ������_IX_��ת�� On ���ﲡ������(��ת��) Tablespace zl9IndexCis;
Create Index ��������¼_IX_����ID On ��������¼(����ID) Tablespace zl9IndexCis;
Create Index ��������¼_IX_�Һ�ID On ��������¼(�Һ�ID) Tablespace zl9IndexCis;
Create Index ��������¼_IX_�Ǽ�ʱ�� On ��������¼(�Ǽ�ʱ��) Tablespace zl9IndexCis;
Create Index ��������¼_IX_��ת�� On ��������¼(��ת��) Tablespace zl9IndexCis;
Create Index ·��ͨ��������Ŀ_IX_������ĿID On ·��ͨ��������Ŀ(������ĿID) Tablespace zl9Indexcis;

Create Index ҩ������˵��_IX_ҽ��B On ҩ������˵��(ҽ��B) Tablespace zl9Indexcis;

Create Index ҩ������˵��_IX_��ת�� On ҩ������˵��(��ת��) Tablespace zl9Indexcis;
Create Index ������ҩ�嵥_IX_����ID on ������ҩ�嵥(����ID, ��ҳID) Tablespace zl9indexhis;
Create Index ������ҩ�嵥_IX_�÷�ID on ������ҩ�嵥(�÷�ID) Tablespace zl9indexhis;
Create Index ������ҩ�嵥_IX_�巨ID on ������ҩ�嵥(�巨ID) Tablespace zl9indexhis;

Create Index ������ҩ�嵥_IX_�շ�ϸĿID on ������ҩ�嵥(�շ�ϸĿID) Tablespace zl9indexhis;
Create Index ������ҩ�嵥_IX_������ĿID on ������ҩ�嵥(������ĿID) Tablespace zl9indexhis;
Create Index ������ҩ�嵥_IX_��ʼʱ�� on ������ҩ�嵥(��ʼʱ��) Tablespace zl9indexhis;
Create Index ������ҩ�嵥_IX_��ת�� on ������ҩ�嵥(��ת��) Tablespace zl9indexhis;
Create Index ������ҩ�䷽_IX_�շ�ϸĿID on ������ҩ�䷽(�շ�ϸĿID) Tablespace zl9indexhis;

Create Index ������ҩ�䷽_IX_������ĿID on ������ҩ�䷽(������ĿID) Tablespace zl9indexhis;
Create Index ������ҩ�䷽_IX_��ת�� on ������ҩ�䷽(��ת��) Tablespace zl9indexhis;
Create Index ����Σ��ֵ��¼_IX_����ID On ����Σ��ֵ��¼(����ID,��ҳID)  Tablespace zl9Indexcis;
Create Index ����Σ��ֵ��¼_IX_�Һŵ� On ����Σ��ֵ��¼(�Һŵ�)  Tablespace zl9Indexcis;
Create Index ����Σ��ֵ��¼_IX_ҽ��ID On ����Σ��ֵ��¼(ҽ��ID) Tablespace zl9Indexcis;
Create Index ����Σ��ֵ��¼_IX_����ʱ�� On ����Σ��ֵ��¼(����ʱ��)  Tablespace zl9Indexcis;
Create Index ����Σ��ֵ��¼_IX_��ת�� On ����Σ��ֵ��¼(��ת��) Tablespace zl9Indexcis;
Create Index ����Σ��ֵҽ��_IX_ҽ��ID On ����Σ��ֵҽ��(ҽ��ID) Tablespace zl9Indexcis;

Create Index ����Σ��ֵҽ��_IX_��ת�� On ����Σ��ֵҽ��(��ת��) Tablespace zl9Indexcis;
Create Index ����Σ��ֵ����_IX_��ת�� On ����Σ��ֵ����(��ת��) Tablespace zl9Indexcis;

Create Index ҽ�����뵥�ļ�_IX_��ת�� On ҽ�����뵥�ļ�(��ת��) Tablespace zl9Indexcis;

Create Index ҽ�����뵥�ļ�_IX_�ļ�ID On ҽ�����뵥�ļ�(�ļ�ID) Tablespace zl9Indexcis;
Create Index ҽ����������_IX_��ת�� On ҽ����������(��ת��) Tablespace zl9Indexcis;

Create Index �����걨����_IX_�Ǽ�ʱ�� On �����걨����(�Ǽ�ʱ��) Tablespace zl9Indexcis;

Create Index �����걨����_IX_��ת�� On �����걨����(��ת��) Tablespace zl9Indexcis;
Create Index �������Լ�¼_IX_����ID On �������Լ�¼(����ID,��ҳID)  Tablespace zl9Indexcis;
Create Index �������Լ�¼_IX_�Ǽ�ʱ�� On �������Լ�¼(�Ǽ�ʱ��)  Tablespace zl9Indexcis;
Create Index �������Լ�¼_IX_�Һŵ� On �������Լ�¼(�Һŵ�)  Tablespace zl9Indexcis;
Create Index �������Լ�¼_IX_��ת�� On �������Լ�¼(��ת��) Tablespace zl9Indexcis;
Create Index �������Լ�¼_IX_�ļ�ID On �������Լ�¼(�ļ�ID) Tablespace zl9Indexcis;
Create Index �������Լ�¼_IX_ҽ��ID On �������Լ�¼(ҽ��ID) Tablespace zl9Indexcis;
Create Index �������淴��_IX_��ת�� On �������淴��(��ת��) Tablespace zl9Indexcis;
Create Index �������淴��_IX_�Ǽ�ʱ�� On �������淴��(�Ǽ�ʱ��) Tablespace zl9Indexcis;

Create Index ��Ѫ������_IX_������ĿID On ��Ѫ������(������ĿID) Tablespace zl9Indexcis;
Create Index ��Ѫ������_IX_��ת�� On ��Ѫ������(��ת��) Tablespace zl9Indexcis;
Create Index �ŶӼ�¼_IX_����ID On �ŶӼ�¼(����ID) Tablespace zl9Indexcis;
Create Index �ŶӼ�¼_IX_���� On �ŶӼ�¼(����) Tablespace zl9Indexcis;
Create Index �ŶӼ�¼_IX_���б�־ On �ŶӼ�¼(���б�־) Tablespace zl9Indexcis;
Create Index ��λ״����¼_IX_����ID On ��λ״����¼(����ID) Tablespace zl9Indexcis;
Create Index ��λ״����¼_IX_�շ�ϸĿid On ��λ״����¼(�շ�ϸĿid) Tablespace zl9Indexcis;
Create Index ������Һ������־_IX_ʱ�� On ������Һ������־(ʱ��) Tablespace ZL9INDEXCIS;
Create Index ������Һ������־_IX_����Ա On ������Һ������־(����Ա) Tablespace ZL9INDEXCIS;
Create Index ������Һ������־_IX_�Һŵ� On ������Һ������־(�Һŵ�) Tablespace ZL9INDEXCIS;
Create Index ���ﴩ��̨_Ix_��������id On ���ﴩ��̨(��������id) Pctfree 5 Tablespace Zl9indexcis;

Create Index ����ҽ����¼_IX_���ID On ����ҽ����¼(���ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����ҽ����¼_IX_��ҳID On ����ҽ����¼(����ID,��ҳID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����ҽ����¼_IX_������ĿID On ����ҽ����¼(������Ŀid) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����ҽ����¼_IX_�շ�ϸĿID On ����ҽ����¼(�շ�ϸĿid) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����ҽ����¼_IX_�Һŵ� On ����ҽ����¼(�Һŵ�) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����ҽ����¼_IX_����ʱ�� On ����ҽ����¼(����ʱ��,ҽ��״̬) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����ҽ����¼_IX_��ʼִ��ʱ�� On ����ҽ����¼(��ʼִ��ʱ��) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����ҽ����¼_IX_����ʱ�� On ����ҽ����¼(����ʱ��) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����ҽ����¼_IX_���״̬ On ����ҽ����¼(���״̬) Tablespace zl9Indexcis;
Create Index ����ҽ����¼_IX_������� On ����ҽ����¼(�������) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����ҽ����¼_IX_�䷽ID On ����ҽ����¼(�䷽ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����ҽ����¼_IX_��ת�� On ����ҽ����¼(��ת��) Tablespace zl9Indexcis;
Create Index ����ҽ����¼_IX_������� On ����ҽ����¼(�������) Pctfree 5 Tablespace zl9Indexcis;

Create Index ����ҽ��״̬_IX_����ʱ�� On ����ҽ��״̬(����ʱ��,��������) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����ҽ��״̬_IX_ǩ��ID On ����ҽ��״̬(ǩ��ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����ҽ��״̬_IX_��ת�� On ����ҽ��״̬(��ת��) Tablespace zl9Indexcis;
Create Index ����ҽ���Ƽ�_IX_�շ�ϸĿID On ����ҽ���Ƽ�(�շ�ϸĿID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����ҽ���Ƽ�_IX_��ת�� On ����ҽ���Ƽ�(��ת��) Tablespace zl9Indexcis;
Create Index ����ҽ������_IX_���ͺ� On ����ҽ������(���ͺ�) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����ҽ������_IX_����ʱ�� On ����ҽ������(����ʱ��,ִ��״̬) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����ҽ������_IX_�״�ʱ�� On ����ҽ������(�״�ʱ��) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����ҽ������_IX_����ʱ�� On ����ҽ������(����ʱ��) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����ҽ������_IX_�������� On ����ҽ������(��������) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����ҽ������_IX_�������� On ����ҽ������(��������) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����ҽ������_IX_��ת�� On ����ҽ������(��ת��) Tablespace zl9Indexcis;
Create Index ����ҽ������_IX_����ID On ����ҽ������(����ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����ҽ������_IX_��ת�� On ����ҽ������(��ת��) Tablespace zl9Indexcis;
Create Index ����ҽ������_IX_RISID On ����ҽ������(RISID)  Tablespace zl9Indexcis;
Create Index ����ҽ������_IX_����ID On ����ҽ������(����ID) Tablespace zl9Indexcis;
Create Index ����ҽ������_IX_����ִ�� On ����ҽ������(����ʱ��,ִ�в���id) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����ҽ������_IX_�걾�������� On ����ҽ������(�걾��������) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����ҽ���쳣��¼_IX_����ID On ����ҽ���쳣��¼(����ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����ҽ���쳣��¼_IX_NO On ����ҽ���쳣��¼(NO,��¼����) Pctfree 5 Tablespace zl9Indexcis;

Create Index ����ҽ������_IX_NO	On ����ҽ������(NO,��¼����) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����ҽ������_IX_��ת�� On ����ҽ������(��ת��) Tablespace zl9Indexcis;
Create Index ����ҽ������_IX_��ת�� On ����ҽ������(��ת��) Tablespace zl9Indexcis;
Create Index ҽ��ǩ����¼_IX_֤��ID On ҽ��ǩ����¼(֤��ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ҽ��ǩ����¼_IX_��ת�� On ҽ��ǩ����¼(��ת��) Tablespace zl9Indexcis;
Create Index ����ҽ����ӡ_IX_��ҳID On ����ҽ����ӡ(����ID,��ҳID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����ҽ����ӡ_IX_��ӡʱ�� On ����ҽ����ӡ(��ӡʱ��) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����ҽ����ӡ_IX_��ת�� On ����ҽ����ӡ(��ת��) Tablespace zl9Indexcis;
Create Index ҽ��ִ��ʱ��_Ix_Ҫ��ʱ�� On ҽ��ִ��ʱ��(Ҫ��ʱ��) Pctfree 5 Tablespace Zl9indexcis;
Create Index ҽ��ִ��ʱ��_IX_��ת�� On ҽ��ִ��ʱ��(��ת��) Tablespace zl9Indexcis;
Create Index ҽ��ִ�мƼ�_IX_�շ�ϸĿid On ҽ��ִ�мƼ�(�շ�ϸĿid) Pctfree 5 Tablespace zl9Indexcis;
Create Index ҽ��ִ�мƼ�_IX_��ת�� On ҽ��ִ�мƼ�(��ת��) Tablespace zl9Indexcis;
Create Index ҽ��ִ�д�ӡ_IX_��ת�� On ҽ��ִ�д�ӡ(��ת��) Tablespace zl9Indexcis;
Create Index ִ�д�ӡ��¼_IX_��ˮ�� On ִ�д�ӡ��¼(��ˮ��) Tablespace zl9Indexcis;
Create Index ִ�д�ӡ��¼_IX_��ת�� On ִ�д�ӡ��¼(��ת��) Tablespace zl9Indexcis;

Create Index ����ҽ��ִ��_IX_ִ��ʱ�� On ����ҽ��ִ��(ִ��ʱ��) Tablespace zl9Indexcis;
Create Index ����ҽ��ִ��_IX_��ˮ�� On ����ҽ��ִ��(��ˮ��) Tablespace zl9Indexcis;
Create Index ����ҽ��ִ��_IX_��ת�� On ����ҽ��ִ��(��ת��) Tablespace zl9Indexcis;
Create Index ���Ƶ��ݴ�ӡ_IX_��ת�� On ���Ƶ��ݴ�ӡ(��ת��) Tablespace zl9Indexcis;
Create Index ��Ѫ�����¼_IX_��ת�� On ��Ѫ�����¼(��ת��) Tablespace zl9Indexcis;
Create Index ������ļ�¼_IX_����ID On ������ļ�¼(����ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ������ļ�¼_IX_��ת�� On ������ļ�¼(��ת��) Tablespace zl9Indexcis;
Create index ��Ѫ������Ŀ_IX_������ĿID on ��Ѫ������Ŀ (������ĿID) tablespace ZL9INDEXCIS;
Create index ��Ѫ������Ŀ_IX_��ת�� on ��Ѫ������Ŀ (��ת��) tablespace ZL9INDEXCIS;

Create Index ҵ����Ϣ�嵥_IX_����ID On ҵ����Ϣ�嵥(����ID,����ID) Tablespace zl9Indexcis;
Create Index ҵ����Ϣ�嵥_IX_�Ǽ�ʱ�� On ҵ����Ϣ�嵥(�Ǽ�ʱ��) Tablespace zl9Indexcis;
Create Index ҵ����Ϣ״̬_IX_�Ķ�ʱ�� On ҵ����Ϣ״̬(�Ķ�ʱ��) Tablespace zl9Indexcis;

----------------------------------------------------------------------------
--[[17.�ٴ�·��]]
----------------------------------------------------------------------------
Create Index ����·�������ļ�_IX_·��ID On ����·�������ļ�(·��ID) Tablespace zl9Indexcis;

Create Index ����·���������_IX_����ID On ����·���������(����ID) Tablespace zl9Indexcis;

Create Index ����·������_IX_�ļ�ID On ����·������(�ļ�ID) Tablespace zl9Indexcis;
Create Index ����·���׶�_IX_��ID On ����·���׶�(��ID) Tablespace zl9Indexcis;

Create Index ����·������_IX_�׶�ID On ����·������(�׶�ID) Tablespace zl9Indexcis;

Create Index ����·����Ŀ_IX_�汾�� On ����·����Ŀ(·��ID,�汾��) Tablespace zl9Indexcis;

Create Index ����·����Ŀ_IX_�׶�ID On ����·����Ŀ(�׶�ID) Tablespace zl9Indexcis;
Create Index ����·����Ŀ_IX_ͼ��ID On ����·����Ŀ(ͼ��ID) Tablespace zl9Indexcis;
Create Index ����·����������_IX_����ID On ����·����������(����ID) Tablespace zl9Indexcis;

Create Index ����·����������_IX_��ĿID On ����·����������(��ĿID) Tablespace zl9Indexcis;
Create Index ����·��ҽ������_IX_���ID On ����·��ҽ������(���ID) Tablespace zl9Indexcis;

Create Index ����·��ҽ������_IX_������ĿID On ����·��ҽ������(������ĿID) Tablespace zl9Indexcis;
Create Index ����·��ҽ������_IX_�շ�ϸĿID On ����·��ҽ������(�շ�ϸĿID) Tablespace zl9Indexcis;
Create Index ����·��ҽ������_IX_ִ�п���ID On ����·��ҽ������(ִ�п���ID) Tablespace zl9Indexcis;
Create Index ����·��ҽ������_IX_�䷽ID On ����·��ҽ������(�䷽ID) Tablespace zl9Indexcis;
Create Index ����·��ҽ��_IX_ҽ������ID On ����·��ҽ��(ҽ������ID) Tablespace zl9Indexcis;

Create Index ����·��ҽ���䶯_IX_������ĿID On ����·��ҽ���䶯(������ĿID) Tablespace zl9Indexcis;

Create Index ����·��ҽ���䶯_IX_�շ�ϸĿID On ����·��ҽ���䶯(�շ�ϸĿId)   Tablespace zl9Indexcis;
Create Index ����·��ҽ���䶯_IX_�䷽ID On ����·��ҽ���䶯(�䷽ID)  Tablespace zl9Indexcis;
Create Index ��������·��_IX_����ID On ��������·��(����ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ��������·��_IX_�Һ�ID On ��������·��(�Һ�ID) Pctfree 5 Tablespace zl9Indexcis;

Create Index ��������·��_IX_·��ID On ��������·��(·��ID,�汾��) Pctfree 5 Tablespace zl9Indexcis;
Create Index ��������·��_IX_����ʱ�� On ��������·��(����ʱ��) Pctfree 5 Tablespace zl9Indexcis;
Create Index ��������·��_IX_����ID On ��������·��(����ID) Tablespace zl9Indexcis;
Create Index ��������·��_IX_���ID On ��������·��(���ID) Tablespace zl9Indexcis;
Create Index ��������·��_IX_δ����ԭ�� On ��������·��(δ����ԭ��) Tablespace zl9Indexcis;
Create Index ��������·��_IX_��ת�� On ��������·��(��ת��) Tablespace zl9Indexcis;
Create Index ��������·����¼_IX_��ת�� On ��������·����¼(��ת��) Tablespace zl9Indexcis;
Create Index ��������·����¼_IX_�Һ�ID On ��������·����¼(�Һ�ID) Tablespace zl9Indexcis;

Create Index ��������·������_IX_���� On ��������·������(����) Pctfree 5 Tablespace zl9Indexcis;

Create Index ��������·������_IX_�Ǽ�ʱ�� On ��������·������(�Ǽ�ʱ��) Pctfree 5 Tablespace zl9Indexcis;
Create Index ��������·������_IX_�׶�ID On ��������·������(�׶�ID) Tablespace zl9Indexcis;
Create Index ��������·������_IX_����ԭ�� On ��������·������(����ԭ��) Pctfree 5 Tablespace zl9Indexcis;
Create Index ��������·������_IX_��ת�� On ��������·������(��ת��) Tablespace zl9Indexcis;
Create Index ��������·������_IX_����ԭ�� On ��������·������(����ԭ��) Tablespace zl9Indexcis;

Create Index ��������·������_IX_��ת�� On ��������·������(��ת��) Tablespace zl9Indexcis;
Create Index ��������·��ִ��_IX_���� On ��������·��ִ��(����) Pctfree 5 Tablespace zl9Indexcis;

Create Index ��������·��ִ��_IX_·����¼ID On ��������·��ִ��(·����¼ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ��������·��ִ��_IX_�׶�ID On ��������·��ִ��(�׶�ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ��������·��ִ��_IX_��ĿID On ��������·��ִ��(��ĿID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ��������·��ִ��_IX_ͼ��ID On ��������·��ִ��(ͼ��ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ��������·��ִ��_IX_�Ǽ�ʱ�� On ��������·��ִ��(�Ǽ�ʱ��) Pctfree 5 Tablespace zl9Indexcis;
Create Index ��������·��ִ��_IX_����ԭ�� On ��������·��ִ��(����ԭ��) Pctfree 5 Tablespace zl9Indexcis;
Create Index ��������·��ִ��_IX_��ת�� On ��������·��ִ��(��ת��) Tablespace zl9Indexcis;
Create Index ��������·��ָ��_IX_���� On ��������·��ָ��(����) Pctfree 5 Tablespace zl9Indexcis;

Create Index ��������·��ָ��_IX_�׶�ID On ��������·��ָ��(�׶�ID) Tablespace zl9Indexcis;
Create Index ��������·��ָ��_IX_��ת�� On ��������·��ָ��(��ת��) Tablespace zl9Indexcis;
Create Index ��������·��ҽ��_IX_����ҽ��ID On ��������·��ҽ��(����ҽ��ID) Pctfree 5 Tablespace zl9Indexhis;

Create Index ��������·��ҽ��_IX_��ת�� On ��������·��ҽ��(��ת��) Tablespace zl9Indexcis;
Create Index �������������¼_IX_·����¼ID On �������������¼(·����¼ID) Tablespace zl9Indexhis;

Create Index �������������¼_IX_��ת�� On �������������¼(��ת��) Tablespace zl9Indexhis;
Create Index �������������¼_IX_����ID On �������������¼(����ID) Tablespace zl9Indexcis;
Create Index �������������¼_IX_�Һ�ID On �������������¼(�Һ�ID) Tablespace zl9Indexcis;
Create Index ��������·��ȡ��_IX_����ID On ��������·��ȡ��(����ID) Tablespace zl9Indexcis;

Create Index ��������·��ȡ��_IX_�Һ�ID On ��������·��ȡ��(�Һ�ID) Tablespace zl9Indexcis;
Create Index ·��ҽ���䶯_IX_������ĿID On ·��ҽ���䶯(������ĿID) Tablespace zl9Indexcis;
Create Index ·��ҽ���䶯_IX_�շ�ϸĿID On ·��ҽ���䶯(�շ�ϸĿId)   Tablespace zl9Indexcis;
Create Index ·��ҽ���䶯_IX_�䷽ID On ·��ҽ���䶯(�䷽ID)  Tablespace zl9Indexcis;
Create Index �����ٴ�·��_IX_����ID On �����ٴ�·��(����ID,��ҳID) Pctfree 5 Tablespace zl9Indexcis;
Create Index �����ٴ�·��_IX_·��ID On �����ٴ�·��(·��ID,�汾��) Pctfree 5 Tablespace zl9Indexcis;
Create Index �����ٴ�·��_IX_����ʱ�� On �����ٴ�·��(����ʱ��) Pctfree 5 Tablespace zl9Indexcis;
Create Index �����ٴ�·��_IX_����ID On �����ٴ�·��(����ID) Tablespace zl9Indexcis;
Create Index �����ٴ�·��_IX_���ID On �����ٴ�·��(���ID) Tablespace zl9Indexcis;
Create Index �����ٴ�·��_IX_δ����ԭ�� On �����ٴ�·��(δ����ԭ��) Tablespace zl9Indexcis;
Create Index �����ٴ�·��_IX_��ת�� On �����ٴ�·��(��ת��) Tablespace zl9Indexcis;

Create Index ����·������_IX_����ԭ�� On ����·������(����ԭ��) Tablespace zl9Indexcis;
Create Index ����·������_IX_��ת�� On ����·������(��ת��) Tablespace zl9Indexcis;
Create Index ����·��ҽ������_IX_����ԭ�� On ����·��ҽ������(����ԭ��) Tablespace zl9Indexcis;
Create Index ����·��ҽ������_IX_��ת�� On ����·��ҽ������(��ת��) Tablespace zl9Indexcis;
Create Index ����·��ҽ��_IX_����ҽ��ID On ����·��ҽ��(����ҽ��ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index ����·��ҽ��_IX_��ת�� On ����·��ҽ��(��ת��) Tablespace zl9Indexcis;
Create Index ����·��ִ��_IX_���� On ����·��ִ��(����) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����·��ִ��_IX_·����¼ID On ����·��ִ��(·����¼ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����·��ִ��_IX_�׶�ID On ����·��ִ��(�׶�ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����·��ִ��_IX_��ĿID On ����·��ִ��(��ĿID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����·��ִ��_IX_ͼ��ID On ����·��ִ��(ͼ��ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����·��ִ��_IX_�Ǽ�ʱ�� On ����·��ִ��(�Ǽ�ʱ��) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����·��ִ��_IX_����ԭ�� On ����·��ִ��(����ԭ��) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����·��ִ��_IX_��ת�� On ����·��ִ��(��ת��) Tablespace zl9Indexcis;

Create Index ����·������_IX_���� On ����·������(����) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����·������_IX_�Ǽ�ʱ�� On ����·������(�Ǽ�ʱ��) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����·������_IX_�׶�ID On ����·������(�׶�ID) Tablespace zl9Indexcis;
Create Index ����·������_IX_����ԭ�� On ����·������(����ԭ��) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����·������_IX_��ת�� On ����·������(��ת��) Tablespace zl9Indexcis;
Create Index ���˺ϲ�·��_IX_��ҳID On ���˺ϲ�·��(����ID,��ҳID) Tablespace zl9Indexhis;
Create Index ���˺ϲ�·��_IX_�汾�� On ���˺ϲ�·��(·��ID,�汾��) Tablespace zl9Indexhis;
Create Index ���˺ϲ�·��_IX_����ID On ���˺ϲ�·��(����ID) Tablespace zl9Indexhis;
Create Index ���˺ϲ�·��_IX_��ǰ�׶�ID On ���˺ϲ�·��(��ǰ�׶�ID) Tablespace zl9Indexhis;
Create Index ���˺ϲ�·��_IX_ǰһ�׶�ID On ���˺ϲ�·��(ǰһ�׶�ID) Tablespace zl9Indexhis;
Create Index ���˺ϲ�·��_IX_��Ҫ·���׶�ID On ���˺ϲ�·��(��Ҫ·���׶�ID) Tablespace zl9Indexhis;
Create Index ���˺ϲ�·��_IX_��Ҫ·����¼ID On ���˺ϲ�·��(��Ҫ·����¼ID) Tablespace zl9Indexhis;
Create Index ���˺ϲ�·��_IX_��ת�� On ���˺ϲ�·��(��ת��) Tablespace zl9Indexcis;
Create Index ���˺ϲ�·������_IX_��ת�� On ���˺ϲ�·������(��ת��) Tablespace zl9Indexcis;

Create Index ����·��ִ��_IX_�ϲ�·���׶�ID On ����·��ִ��(�ϲ�·���׶�ID) Tablespace zl9Indexhis;
Create Index ����·��ִ��_IX_�ϲ�·����¼ID On ����·��ִ��(�ϲ�·����¼ID) Tablespace zl9Indexhis;
Create Index ����·��ָ��_IX_�ϲ�·���׶�ID On ����·��ָ��(�ϲ�·����¼ID) Tablespace zl9Indexhis;
Create Index ����·��ָ��_IX_���� On ����·��ָ��(����) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����·��ָ��_IX_�׶�ID On ����·��ָ��(�׶�ID) Tablespace zl9Indexcis;
Create Index ����·��ָ��_IX_��ת�� On ����·��ָ��(��ת��) Tablespace zl9Indexcis;
Create Index ���˳�����¼_IX_·����¼ID On ���˳�����¼(·����¼ID) Tablespace zl9Indexhis;
Create Index ���˳�����¼_IX_��ת�� On ���˳�����¼(��ת��) Tablespace zl9Indexhis;
Create Index ����·��ȡ��_IX_����ID On ����·��ȡ��(����ID,��ҳID) Tablespace zl9Indexcis;

----------------------------------------------------------------------------
--[[18.����ҵ��]]
----------------------------------------------------------------------------
Create Index ���Ӳ�����¼_IX_����ID On ���Ӳ�����¼(����ID,��ҳID) Pctfree 5 Initrans 20 Tablespace zl9Indexcis;
Create Index ���Ӳ�����¼_IX_�ļ�ID On ���Ӳ�����¼(�ļ�ID) Pctfree 5 Initrans 20 Tablespace zl9Indexcis;
Create Index ���Ӳ�����¼_IX_���ʱ�� On ���Ӳ�����¼(���ʱ��,��������,����id) Pctfree 5 Initrans 20 Tablespace zl9Indexcis;
Create Index ���Ӳ�����¼_IX_����ʱ�� On ���Ӳ�����¼(����ʱ��) Pctfree 5 Initrans 20 Tablespace zl9Indexcis;
Create Index ���Ӳ�����¼_IX_·��ִ��ID On ���Ӳ�����¼(·��ִ��ID) Pctfree 5 Initrans 20 Tablespace zl9Indexcis;
Create Index ���Ӳ�����¼_IX_��ת�� On ���Ӳ�����¼(��ת��) Initrans 20 Tablespace zl9Indexcis;
Create Index ���Ӳ�����¼_IX_����·��ִ��ID On ���Ӳ�����¼(����·��ִ��ID) Pctfree 5 Initrans 20 Tablespace zl9Indexcis;

Create Index ���Ӳ�������_IX_��ID On ���Ӳ�������(��ID) Pctfree 5 Initrans 20 Tablespace zl9Indexcis;
Create Index ���Ӳ�������_IX_Ԥ�����ID On ���Ӳ�������(Ԥ�����ID) Pctfree 5 Initrans 20 Tablespace zl9Indexcis;
Create Index ���Ӳ�������_IX_����Ҫ��ID On ���Ӳ�������(����Ҫ��ID) Pctfree 5 Initrans 20 Tablespace zl9Indexcis;
Create Index ���Ӳ�������_IX_��ת�� On ���Ӳ�������(��ת��)  Initrans 20 Tablespace zl9Indexcis;
Create Index �����䶯ԭ��_IX_�����ļ�id On �����䶯ԭ��(�����ļ�id) Pctfree 5 Tablespace zl9Indexhis;
Create Index �����䶯ԭ��_IX_ԭ��Ҫ��id On �����䶯ԭ��(ԭ��Ҫ��id) Pctfree 5 Tablespace zl9Indexhis;
Create Index �����䶯ԭ��_IX_ԭ��Ҫ�� On �����䶯ԭ��(ԭ��Ҫ��) Pctfree 5 Tablespace zl9Indexhis;
Create Index �����䶯���_IX_�䶯ԭ��id On �����䶯���(�䶯ԭ��id) Pctfree 5 Tablespace zl9Indexhis;
Create Index �����䶯���_IX_���Ҫ��id On �����䶯���(���Ҫ��id) Pctfree 5 Tablespace zl9Indexhis;
Create Index �����䶯���_IX_���Ҫ�� On �����䶯���(���Ҫ��) Pctfree 5 Tablespace zl9Indexhis;
Create Index ���Ӳ�����ӡ_IX_����ID On ���Ӳ�����ӡ(����ID,��ҳID) Pctfree 5 Initrans 20 Tablespace zl9Indexcis;
Create Index ���Ӳ���ʱ��_IX_����ID On ���Ӳ���ʱ��(����ID,��ҳID) Pctfree 20 Tablespace zl9Indexcis;
Create Index ���Ӳ���ʱ��_IX_�ļ�ID On ���Ӳ���ʱ��(�ļ�ID) Pctfree 20 Tablespace zl9Indexcis;
Create Index �����걨��¼_IX_�ĵ�ID On �����걨��¼(�ĵ�ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index �����걨��¼_IX_��ת�� On �����걨��¼(��ת��) Tablespace zl9Indexcis;
Create Index �����걨��¼_IX_���� On �����걨��¼(����) Tablespace zl9Indexcis;
Create Index �����걨��¼_IX_����ID On �����걨��¼(����ID,��ҳID) Tablespace zl9Indexcis;

Create Index ���Ӳ�������_IX_��ת�� On ���Ӳ�������(��ת��) Initrans 20 Tablespace zl9Indexcis;
Create Index ���Ӳ�����ʽ_IX_��ת�� On ���Ӳ�����ʽ(��ת��) Initrans 20 Tablespace zl9Indexcis;
Create Index ���Ӳ���ͼ��_IX_��ת�� On ���Ӳ���ͼ��(��ת��) Initrans 20 Tablespace zl9Indexcis;

--��ʱ��,��Ҫָ����ռ�,Pctfree�Ȳ���
Create Index ����ʱ�޼��_IX_����id On ����ʱ�޼��(����ID,��ҳID,������Դ);
Create Index �������ݼ��_IX_����id On �������ݼ��(����ID,��ҳID,������Դ);

--�������鵵
Create Index �����ύ��¼_IX_��ҳID On �����ύ��¼(����ID,��ҳID) Tablespace zl9Indexcis;
Create Index �����ύ��¼_IX_�ύʱ�� On �����ύ��¼(�ύʱ��) Tablespace zl9Indexcis;
Create Index ������ӡ��¼_IX_��ҳID On ������ӡ��¼(����id,��ҳid) Tablespace zl9Indexcis;
Create Index ������ӡ��¼_IX_��ӡʱ�� On ������ӡ��¼(��ӡʱ��) Tablespace zl9Indexcis;
Create Index ����������ǩ_IX_�ύid On ����������ǩ(�ύid) Tablespace zl9Indexcis;
Create Index ����������¼_IX_��ҳID On ����������¼(����ID,��ҳID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����������¼_IX_�ύid On ����������¼(�ύid) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����������¼_IX_���id On ����������¼(���id) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����������¼_IX_����ʱ�� On ����������¼(����ʱ��) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����������¼_IX_����ʱ�� On ����������¼(����ʱ��) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����������¼_IX_ҽ��id On ����������¼(ҽ��id) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����������¼_IX_����id On ����������¼(����id) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����������ʷ_IX_��ҳID On ����������ʷ(����ID,��ҳID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����������ʷ_IX_ҽ��id On ����������ʷ(ҽ��id) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����������ʷ_IX_����id On ����������ʷ(����id) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����������ʷ_IX_�ύid On ����������ʷ(�ύid) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����������ʷ_IX_���id On ����������ʷ(���id) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����������ʷ_IX_����ʱ�� On ����������ʷ(����ʱ��) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����������ʷ_IX_����ʱ�� On ����������ʷ(����ʱ��) Pctfree 5 Tablespace zl9Indexcis;
Create Index ��������¼_IX_��ҳID On ��������¼(����ID,��ҳID) Tablespace zl9Indexcis;
Create Index ������������_IX_��ҳID On ������������(����ID,��ҳID) Tablespace zl9Indexcis;
Create Index ������������_IX_����id On ������������(����id) Tablespace zl9Indexcis;
Create Index ������������_IX_����id On ������������(����id) Tablespace zl9Indexcis;
Create Index ����������Ա_IX_����id On ����������Ա(����id) Tablespace zl9Indexcis;
Create Index �������ֱ�׼_IX_����ID On �������ֱ�׼(����ID) Tablespace zl9Indexcis;
Create Index �������ֱ�׼_IX_�ϼ�ID On �������ֱ�׼(�ϼ�ID) Tablespace zl9Indexcis;
Create Index �������ֽ��_IX_����ID On �������ֽ��(����ID) Tablespace zl9Indexcis;
Create Index ����������ϸ_IX_���ID On ����������ϸ(����ID) Tablespace zl9Indexcis;
Create Index ����������ϸ_IX_���ֱ�׼ID On ����������ϸ(���ֱ�׼ID) Tablespace zl9Indexcis;
Create Index �������ļ�¼_IX_�Ǽ�ʱ�� On �������ļ�¼(�Ǽ�ʱ��) Tablespace zl9Indexcis;


----------------------------------------------------------------------------
--[[19.����ҵ��]]
----------------------------------------------------------------------------
Create Index ���˻����ļ�_IX_��ҳID On ���˻����ļ�(����ID,��ҳID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ���˻����ļ�_IX_��ת�� On ���˻����ļ�(��ת��) Tablespace zl9Indexcis;
Create index ���˻����ļ�_IX_����ID On ���˻����ļ�(����ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ���˻����¼_IX_��ת�� On ���˻����¼(��ת��) Tablespace zl9Indexcis;
Create Index ���˻����¼_IX_��ҳID On ���˻����¼(����ID,��ҳID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ���˻����¼_IX_����ʱ�� On ���˻����¼(����ʱ��) Pctfree 5 Tablespace zl9Indexcis;
Create Index ���˻�������_IX_��ת�� On ���˻�������(��ת��) Tablespace zl9Indexcis;
Create Index ���˻�������_IX_��¼id On ���˻�������(��¼id) Pctfree 5 Tablespace zl9Indexcis;

Create Index ���˻�������_IX_�ļ�ID On ���˻�������(�ļ�ID,����ʱ��) Pctfree 5 Tablespace zl9Indexcis;
Create Index ���˻�������_IX_��ת�� On ���˻�������(��ת��) Tablespace zl9Indexcis;
Create Index ���˻�����ϸ_IX_��¼ID On ���˻�����ϸ(��¼ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ���˻�����ϸ_IX_��ԴID On ���˻�����ϸ(��ԴID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ���˻�����ϸ_IX_��ת�� On ���˻�����ϸ(��ת��) Tablespace zl9Indexcis;

Create Index ���˻����ӡ_IX_�ļ�ID On ���˻����ӡ(�ļ�ID,����ʱ��) Pctfree 5 Tablespace zl9Indexcis;
Create Index ���˻����ӡ_IX_��ת�� On ���˻����ӡ(��ת��) Tablespace zl9Indexcis;
Create Index ������Ǽ�¼_IX_��ҳID On ������Ǽ�¼(����ID,��ҳID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ���˻�����Ŀ_IX_��ת�� On ���˻�����Ŀ(��ת��) Tablespace zl9Indexcis;
Create Index ����Ҫ������_IX_��ת�� On ����Ҫ������(��ת��) Tablespace zl9Indexcis;
Create Index ���˻���Ҫ������_IX_��ת�� On ���˻���Ҫ������(��ת��) Tablespace zl9Indexcis;
Create Index ���˻������_IX_����ID On ���˻������(����ID,��ҳID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ���˻������_IX_�ļ�ID On ���˻������(�ļ�ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ���˻������_IX_��ת�� ON ���˻������ (��ת��) Tablespace zl9Indexcis;


----------------------------------------------------------------------------

--[[20.����ҵ��]]

----------------------------------------------------------------------------
Create Index ������ˮ�߱걾_IX_��ת�� On ������ˮ�߱걾(��ת��) Pctfree 5 Tablespace zl9Indexcis;
Create Index ������ˮ��ָ��_IX_��ת�� On ������ˮ��ָ��(��ת��) Pctfree 5 Tablespace zl9Indexcis;
Create Index ������ˮ��ָ��_IX_��ĿID On ������ˮ��ָ��(��ĿID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ������ˮ�߱걾_IX_�걾ID On ������ˮ�߱걾(�걾ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ������ˮ��ָ��_IX_�걾ID On ������ˮ��ָ��(�걾ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����걾��¼_IX_ҽ��ID On ����걾��¼(ҽ��id) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����걾��¼_IX_����ʱ�� On ����걾��¼(����ʱ��) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����걾��¼_IX_����ʱ�� On ����걾��¼(����ʱ��) Pctfree 5 Tablespace ZL9INDEXCIS;
Create Index ����걾��¼_IX_���ʱ�� On ����걾��¼(���ʱ��) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����걾��¼_IX_�������� On ����걾��¼(��������) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����걾��¼_IX_�Һŵ� On ����걾��¼(�Һŵ�) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����걾��¼_IX_�ϲ�ID On ����걾��¼(�ϲ�ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����걾��¼_IX_��ҳID On ����걾��¼(����ID,��ҳID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����걾��¼_IX_��ʶ�� On ����걾��¼(��ʶ��) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����걾��¼_IX_��ת�� On ����걾��¼(��ת��) Tablespace zl9Indexcis;
Create Index ����걾��¼_IX_NO On ����걾��¼(NO) Tablespace zl9Indexcis;

Create Index ������ͨ���_IX_ϸ��ID On ������ͨ���(ϸ��ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ������ͨ���_IX_����ID On ������ͨ���(����ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ������ͨ���_IX_����걾ID On ������ͨ���(����걾ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ������ͨ���_IX_ҩ����ID On ������ͨ���(ҩ����ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ������ͨ���_IX_ø���ID On ������ͨ���(ø���ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ������ͨ���_IX_��Ŀid On ������ͨ���(������ĿID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ������ͨ���_IX_��ת�� On ������ͨ���(��ת��) Tablespace zl9Indexcis;
Create Index ������Ŀ�ֲ�_IX_�걾id On ������Ŀ�ֲ�(�걾id) Pctfree 5 Tablespace zl9Indexcis;
Create Index ������Ŀ�ֲ�_IX_��Ŀid On ������Ŀ�ֲ�(��Ŀid) Pctfree 5 Tablespace zl9Indexcis;
Create Index ������Ŀ�ֲ�_IX_ҽ��id On ������Ŀ�ֲ�(ҽ��id) Pctfree 5 Tablespace zl9Indexcis;
Create Index ������Ŀ�ֲ�_IX_ϸ��ID On ������Ŀ�ֲ�(ϸ��id) Pctfree 5 Tablespace zl9Indexcis;
Create Index ������Ŀ�ֲ�_IX_��ת�� On ������Ŀ�ֲ�(��ת��) Tablespace zl9Indexcis;

Create Index �����ʿؼ�¼_IX_����ID On �����ʿؼ�¼(����ID) Tablespace zl9Indexcis;
Create Index �����ʿؼ�¼_IX_�ʿ�ƷID On �����ʿؼ�¼(�ʿ�ƷID) Tablespace zl9Indexcis;
Create Index �����ʿؼ�¼_IX_��ת�� On �����ʿؼ�¼(��ת��) Tablespace zl9Indexcis;
Create Index ����ͼ����_IX_�걾id On ����ͼ����(�걾id) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����ø���¼_IX_����ʱ�� On ����ø���¼(����ʱ��) Tablespace zl9Indexhis;
Create Index ���������¼_IX_�걾id On ���������¼(�걾id) Tablespace zl9Indexcis;
Create Index ���������¼_IX_��ת�� On ���������¼(��ת��) Tablespace zl9Indexhis;
Create Index ���������¼_IX_�걾ID On ���������¼(�걾ID) Tablespace zl9Indexhis;
Create Index ���������¼_IX_��; On ���������¼(��;) Tablespace zl9Indexhis;
Create Index ���������¼_IX_��ת�� On ���������¼(��ת��) Tablespace zl9Indexcis;
Create Index ������ռ�¼_IX_ҽ��ID On ������ռ�¼(ҽ��ID) Tablespace zl9Indexcis;
Create Index ������ռ�¼_IX_��ת�� On ������ռ�¼(��ת��) Tablespace zl9Indexcis;

Create Index ����������Ŀ_IX_��ת�� On ����������Ŀ(��ת��) Tablespace zl9Indexcis;
Create Index �����Լ���¼_IX_��ת�� On �����Լ���¼(��ת��) Tablespace zl9Indexcis;
Create Index �����ʿر���_IX_��ת�� On �����ʿر���(��ת��) Tablespace zl9Indexcis;
Create Index ����ҩ�����_IX_��ת�� On ����ҩ�����(��ת��) Tablespace zl9Indexcis;
Create Index ����ǩ����¼_IX_��ת�� On ����ǩ����¼(��ת��) Tablespace zl9Indexhis;

----------------------------------------------------------------------------
--[[21.���ҵ��]]
----------------------------------------------------------------------------
Create Index Ӱ�񱨸沵��_IX_ҽ��ID On Ӱ�񱨸沵��(ҽ��ID,����ID,��鱨��ID,RISID,����ID) Tablespace ZL9INDEXCIS;
Create Index Ӱ�񱨸沵��_IX_��ת�� On Ӱ�񱨸沵��(��ת��) Tablespace zl9Indexcis;
Create Index Ӱ�񱨸沵��_IX_��鱨��ID On Ӱ�񱨸沵��(��鱨��ID) Tablespace zlPacsBaseIndex;
Create Index Ӱ�����¼_IX_���� On Ӱ�����¼(����, Ӱ�����) Tablespace zl9Indexcis;
Create Index Ӱ�����¼_IX_λ��һ On Ӱ�����¼(λ��һ) Tablespace zl9Indexcis;
Create Index Ӱ�����¼_IX_λ�ö� On Ӱ�����¼(λ�ö�) Tablespace zl9Indexcis;
Create Index Ӱ�����¼_IX_λ���� On Ӱ�����¼(λ����) Tablespace zl9Indexcis;
Create Index Ӱ�����¼_IX_�������� On Ӱ�����¼(��������) Pctfree 5 Tablespace zl9Indexcis;
Create Index Ӱ�����¼_Ix_ִ�п���id On Ӱ�����¼(ִ�п���id) Pctfree 5 Tablespace Zl9Indexcis;
Create Index Ӱ�����¼_IX_����ID On Ӱ�����¼(����ID) Pctfree 5 Tablespace Zl9Indexcis;
Create Index Ӱ�����¼_IX_��ת�� On Ӱ�����¼(��ת��) Tablespace zl9Indexcis;
Create Index Ӱ�����¼_IX_У��״̬ On Ӱ�����¼(У��״̬)  Tablespace zl9Indexcis;

Create Index Ӱ����ʱ��¼_IX_���� On Ӱ����ʱ��¼(����, Ӱ�����) Tablespace zl9Indexcis;
Create Index Ӱ����ʱ��¼_IX_λ��һ On Ӱ����ʱ��¼(λ��һ) Tablespace zl9Indexcis;
Create Index Ӱ����ʱ��¼_IX_λ�ö� On Ӱ����ʱ��¼(λ�ö�) Tablespace zl9Indexcis;
Create Index Ӱ����ʱ��¼_IX_λ���� On Ӱ����ʱ��¼(λ����) Tablespace zl9Indexcis;
Create Index Ӱ����ʱ��¼_IX_�������� On Ӱ����ʱ��¼(��������) Tablespace zl9Indexcis;
Create Index ��Ƭ��ӡ��¼_IX_���ID On ��Ƭ��ӡ��¼(���ID) Tablespace zl9Indexhis;
Create Index ��Ƭ��ӡ��¼_IX_��ӡʱ�� On ��Ƭ��ӡ��¼(��ӡʱ��) Tablespace zl9Indexhis;
Create Index Ӱ���ղ����_IX_�ϼ�ID On Ӱ���ղ����(�ϼ�ID) Tablespace zl9Indexcis;
Create Index Ӱ���ղ����_IX_������ID On Ӱ���ղ����(������ID) Tablespace zl9Indexcis;
Create Index Ӱ�����뵥ͼ��_IX_ҽ��ID On Ӱ�����뵥ͼ��(ҽ��ID) Tablespace zl9Indexcis;
Create Index Ӱ�����뵥ͼ��_IX_��ת�� On Ӱ�����뵥ͼ��(��ת��) Tablespace zl9Indexcis;
Create Index Ӱ���ղ�����_IX_ҽ��ID On Ӱ���ղ�����(ҽ��ID) Tablespace zl9Indexcis;
Create Index Ӱ���ղ�����_IX_��ת�� On Ӱ���ղ�����(��ת��) Tablespace zl9Indexcis;

Create Index Ӱ����ͼ��_IX_��ת�� On Ӱ����ͼ��(��ת��) Tablespace zl9Indexcis;
Create Index Ӱ��������_IX_��ת�� On Ӱ��������(��ת��) Tablespace zl9Indexcis;
Create Index Ӱ��Σ��ֵ��¼_IX_��ת�� On Ӱ��Σ��ֵ��¼(��ת��) Tablespace zl9Indexcis;

Create Index Ӱ��ԤԼ��¼_IX_ԤԼ�豸ID On Ӱ��ԤԼ��¼(ԤԼ�豸ID) Tablespace zl9Indexcis;
Create Index Ӱ��ԤԼ��¼_IX_ԤԼ��ʼʱ�� On Ӱ��ԤԼ��¼(ԤԼ��ʼʱ��) Tablespace zl9Indexcis;
Create Index Ӱ��ԤԼ��¼_IX_ҽ��ID On Ӱ��ԤԼ��¼(ҽ��ID) Tablespace zl9Indexcis;
Create Index Ӱ��ԤԼ��¼_IX_��ת�� On Ӱ��ԤԼ��¼(��ת��) Tablespace zl9Indexcis;
Create Index Ӱ��ԤԼ��Ŀ_IX_ԤԼ�豸ID On Ӱ��ԤԼ��Ŀ(ԤԼ�豸ID) Tablespace zl9Indexcis;
Create Index Ӱ��ԤԼ��Ŀ_IX_������ĿID On Ӱ��ԤԼ��Ŀ(������ĿID) Tablespace zl9Indexcis;
Create Index Ӱ��ԤԼ����_IX_ԤԼ�豸ID On Ӱ��ԤԼ����(ԤԼ�豸ID) Tablespace zl9Indexcis;
Create Index Ӱ��ԤԼʱ��ƻ�_IX_ԤԼ����ID On Ӱ��ԤԼʱ��ƻ�(ԤԼ����ID) Tablespace zl9Indexcis;

Create Index ��������Ϣ_IX_�������ID On ��������Ϣ(�������ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ��������Ϣ_IX_ҽ��ID On ��������Ϣ(ҽ��ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ��������Ϣ_IX_����ʱ�� On ��������Ϣ(����ʱ��) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����������Ϣ_IX_����ҽ��ID On ����������Ϣ(����ҽ��ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����걾��Ϣ_IX_ҽ��ID On ����걾��Ϣ(ҽ��ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����걾��Ϣ_IX_�ͼ�ID On ����걾��Ϣ(�ͼ�ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index �����ͼ���Ϣ_IX_ҽ��ID On �����ͼ���Ϣ(ҽ��ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����������Ϣ_IX_����ҽ��ID On ����������Ϣ(����ҽ��ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����������Ϣ_IX_����ʱ�� On ����������Ϣ(����ʱ��) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����ȡ����Ϣ_IX_����ҽ��ID On ����ȡ����Ϣ(����ҽ��ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����ȡ����Ϣ_IX_����ID On ����ȡ����Ϣ(����ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����ȡ����Ϣ_IX_�걾ID On ����ȡ����Ϣ(�걾ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ����ȡ����Ϣ_IX_ȡ��ʱ�� On ����ȡ����Ϣ(ȡ��ʱ��) Pctfree 5 Tablespace zl9Indexcis;
Create Index �����Ѹ���Ϣ_IX_�걾ID On �����Ѹ���Ϣ(�걾ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ������Ƭ��Ϣ_IX_�Ŀ�ID On ������Ƭ��Ϣ(�Ŀ�ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ������Ƭ��Ϣ_IX_����ID On ������Ƭ��Ϣ(����ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ������Ƭ��Ϣ_IX_����ҽ��ID On ������Ƭ��Ϣ(����ҽ��ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ������Ƭ��Ϣ_IX_��Ƭʱ�� On ������Ƭ��Ϣ(��Ƭʱ��) Pctfree 5 Tablespace zl9Indexcis;
Create Index ������̱���_IX_����ҽ��ID On ������̱���(����ҽ��ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index �����ؼ���Ϣ_IX_����ID On �����ؼ���Ϣ(����ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index �����ؼ���Ϣ_IX_����ID On �����ؼ���Ϣ(����ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index �����ؼ���Ϣ_IX_�Ŀ�ID On �����ؼ���Ϣ(�Ŀ�ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index �����ؼ���Ϣ_IX_���ʱ�� On �����ؼ���Ϣ(���ʱ��) Pctfree 5 Tablespace zl9Indexcis;
Create Index �������ӳ�_IX_����ҽ��ID On �������ӳ�(����ҽ��ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ���������Ϣ_IX_����ҽ��ID On ���������Ϣ(����ҽ��ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index �����巴��_IX_����ID On �����巴��(����ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index �����ײ͹���_IX_����ID On �����ײ͹���(����ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index ��������Ϣ_IX_����ID On ��������Ϣ(����ID) Tablespace zl9Indexcis;
Create Index ��������Ϣ_IX_�������� On ��������Ϣ(��������) Tablespace zl9Indexcis;
Create Index ����鵵��Ϣ_IX_�Ŀ�ID On ����鵵��Ϣ(�Ŀ�ID) Pctfree 5 TableSpace zl9Indexcis;
Create Index ����鵵��Ϣ_IX_��ƬID On ����鵵��Ϣ(��ƬID) Pctfree 5 TableSpace zl9Indexcis;
Create Index ����鵵��Ϣ_IX_�ؼ�ID On ����鵵��Ϣ(�ؼ�ID) Pctfree 5 TableSpace zl9Indexcis;
Create Index ����鵵��Ϣ_IX_����ID On ����鵵��Ϣ(����ID) Pctfree 5 TableSpace zl9Indexcis;
Create Index ���������Ϣ_IX_����ʱ�� On ���������Ϣ(����ʱ��) TableSpace zl9Indexcis;
Create Index ���������Ϣ_IX_֤������ On ���������Ϣ(֤������,֤������) TableSpace zl9Indexcis;
Create Index ������ʧ��Ϣ_IX_����ID On ������ʧ��Ϣ(����ID) TableSpace zl9Indexcis;
Create Index ������ʧ��Ϣ_IX_�鵵ID On ������ʧ��Ϣ(�鵵ID) TableSpace zl9Indexcis;
Create Index ������ʧ��Ϣ_IX_��ʧ���� On ������ʧ��Ϣ(��ʧ����) TableSpace zl9Indexcis;
Create Index ����黹��Ϣ_IX_����ID On ����黹��Ϣ(����ID) TableSpace zl9Indexcis;
Create Index ������Ĺ���_IX_����ID On ������Ĺ���(����ID) TableSpace zl9Indexcis;
Create Index ����Ƭ��Ϣ_IX_�Ŀ�Id On ����Ƭ��Ϣ(�Ŀ�Id) Tablespace zl9Indexcis;
Create Index ����Ƭ��Ϣ_IX_��ԴID On ����Ƭ��Ϣ(��ԴID) Tablespace zl9Indexcis;
Create Index ����Ƭ��Ϣ_IX_����ҽ��ID On ����Ƭ��Ϣ(����ҽ��ID) Tablespace zl9Indexcis;


Create Index Ӱ�񱨸�ֵ���嵥_IX_����ID On Ӱ�񱨸�ֵ���嵥(����ID) Tablespace zlPacsBaseIndex;
Create Index Ӱ�񱨸�Ԫ���嵥_IX_����ID On Ӱ�񱨸�Ԫ���嵥(����ID) Tablespace zlPacsBaseIndex;
Create Index Ӱ�񱨸�Ԫ���嵥_IX_ֵ��ID On Ӱ�񱨸�Ԫ���嵥(ֵ��ID) Tablespace zlPacsBaseIndex;
Create Index Ӱ�񱨸�Ƭ���嵥_IX_�ϼ�ID On Ӱ�񱨸�Ƭ���嵥(�ϼ�ID) Tablespace zlPacsBaseIndex;
Create Index Ӱ�񱨸涯��_IX_ԭ��ID On Ӱ�񱨸涯��(ԭ��ID) Tablespace zlPacsBaseIndex;
Create Index Ӱ�񱨸涯��_IX_�¼�ID On Ӱ�񱨸涯��(�¼�ID) Tablespace zlPacsBaseIndex;
Create Index Ӱ�񱨸��¼_IX_ԭ��ID On Ӱ�񱨸��¼(ԭ��ID) Tablespace zlPacsBizIndex;
Create Index Ӱ�񱨸��¼_IX_��ת�� On Ӱ�񱨸��¼(��ת��) Tablespace zlPacsBizIndex;
Create Index Ӱ�񱨸��¼_IX_ҽ��ID On Ӱ�񱨸��¼(ҽ��ID) Tablespace zlPacsBizIndex;
Create Index Ӱ�����˵��_IX_PID On Ӱ�����˵��(PID) Tablespace zlPacsBaseIndex;
Create Index Ӱ�񱨸������¼_IX_����ID On Ӱ�񱨸������¼(����ID) Tablespace zlPacsBaseIndex;
Create Index Ӱ�񱨸������¼_IX_��ת�� On Ӱ�񱨸������¼(��ת��) Tablespace zlPacsBaseIndex;
Create Index Ӱ�񱨸������¼_IX_ҽ��ID On Ӱ�񱨸������¼(ҽ��ID) Tablespace zlPacsBaseIndex;

Create Index Ӱ���ѯ����_IX_����ģ�� On Ӱ���ѯ����(����ģ��) Tablespace zl9Indexhis;
Create Index Ӱ���ѯ����_IX_�û�ID On Ӱ���ѯ����(�û�ID) Tablespace zl9Indexhis;
Create Index Ӱ���ѯ����_IX_��ѯ����ID On Ӱ���ѯ����(��ѯ����ID) Tablespace zl9Indexhis;
Create Index Ӱ���ѯ����_IX_�û�ID On Ӱ���ѯ����(�û�ID) Tablespace zl9Indexhis;
Create Index Ӱ���ѯ����_IX_��ѯ����ID On Ӱ���ѯ����(��ѯ����ID) Tablespace zl9Indexhis;