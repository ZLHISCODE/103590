Insert Into �����ѽӿ�Ŀ¼(���,����,ϵͳ,���㷽ʽ,����,����,���ƿ�,���ų���,ǰ׺�ı�) Select Max(���)+1,'����POS',1,'POS����','zlZHPOS',1,2,16,Null FROM �����ѽӿ�Ŀ¼;

insert into ���㷽ʽ (����,����, ����, ����, ȱʡ��־) 	SELECT  nvl(max(to_number(����)),0) +1, 'POS����', 'POS', 8, 0 FROM ���㷽ʽ;
insert into ���㷽ʽӦ�� (Ӧ�ó���, ���㷽ʽ, ȱʡ��־) values ('�շ�', 'POS����', 0);
insert into ���㷽ʽӦ�� (Ӧ�ó���, ���㷽ʽ, ȱʡ��־) values ('����', 'POS����', 0);
insert into ���㷽ʽӦ�� (Ӧ�ó���, ���㷽ʽ, ȱʡ��־) values ('Ԥ����', 'POS����', 0);
