Create Table ִ�д�ӡ��¼ (
       ҽ��ID     Number(18),
       ���ͺ�         Number(18),
       ��ˮ��     Number(18),
       ��ӡ˵��       Varchar2(1000),
       ��ӡʱ��       Date,
       ��ӡ��         Varchar2(20))
       TABLESPACE zl9CisRec
       Pctfree 5 Pctused 85;

Create Table �ݴ�ҩƷ��¼ (
       NO             VARCHAR2(8),
       ���           NUMBER(5),
       ����ID         Number(18),
       ����ID         Number(18),	
       ҽ��ID         Number(18),	
       ���ͺ�         Number(18),
       ҩƷID         Number(18),	
       ҩƷ����       Varchar2(80),	
       ���           Varchar2(40),
       ִ�з���       Number(2),    -- 0-���������� 1-��Һ�� 2-ע���� 3-Ƥ����
       ʹ��״̬       Number(1),    -- 0-δ��,1-����
       ժҪ           Varchar2(200),	
       ���ϵ��       Number(2),    -- 1-���ݴ�ҩƷ -1-ʹ���ݴ�ҩƷ
       ��λ           varchar2(20), -- Ŀ¼�ڵ�ҩƷ��ҽ��ҩƷΪ���㵥λ ,Ŀ¼��ҩƷΪ���ﵥλ
       ����           Number(16,5),
       ����           Number(16,5), -- ��������,Ŀ¼�ڼ�¼���Ǽ��㵥λ����,Ŀ¼��Ϊ���ﵥλ����
       ����           Number(16,5),	
       ���           Number(16,5),	
       ����Ա         Varchar2(10),	
       �Ǽ�ʱ��       Date,	
       ����ʱ��       Date) --	ʹ��״̬Ϊ1�ļ�¼��������
       TABLESPACE zl9CisRec
       Pctfree 5 Pctused 85;

Create Table ��λ״����¼(
       ����ID         Number(18),
       ����ID         Number(18),
       ���           Varchar2(30), -- ��λ���
       ���           Number(1), -- 0-��ͨ��λ 1-���� 2-����ҩƷ��λ 3-VIP��λ  
       �շ�ϸĿID     Number(18), -- ��Ҫ�շѣ����Ŷ�Ӧ���շ�ϸĿID
       ״̬           Number(1), -- 0-��,1-����,2-������,������ά��
       ��ע           Varchar2(100),
       NO             Varchar2(8))
       TABLESPACE zl9CisRec
       Pctfree 5 Pctused 85;
       

Create Table �ŶӼ�¼(
       ����ID         Number(18),	
       ����ID         Number(18),	
       ����           Date Default Sysdate,	
       ˳���         Number(5), -- �����Ŷӵ�˳���
       ��Ȩ��         Number(10), -- ���ⲡ�������¸ı�˳����
       ״̬           Number(2), -- 0-���� 1-��� 2-���� 3-�˺�
       ��ע           Varchar2(100))
       TABLESPACE zl9CisRec
       Pctfree 5 Pctused 85;         
