alter table �ŶӼ�¼  add constraint �ŶӼ�¼_PK primary key (����, ����ID, ˳���)  using index  Pctfree 5 Tablespace zl9indexcis;

alter table ��λ״����¼ add constraint ��λ״����¼_PK primary key (����ID, ���) Using Index Pctfree 5 Tablespace zl9indexcis;

alter table �ݴ�ҩƷ��¼ add constraint �ݴ�ҩƷ��¼_PK primary key (NO, ���, ���ϵ��, �Ǽ�ʱ��) Using Index Pctfree 5 Tablespace zl9indexcis;
alter table �ݴ�ҩƷ��¼  add constraint �ݴ�ҩƷ��¼_FK_����ID foreign key (����ID) references ������Ϣ (����ID);
alter table �ݴ�ҩƷ��¼  add constraint �ݴ�ҩƷ��¼_FK_����ID foreign key (����ID) references ���ű� (ID);
-----------  
-- alter table �ݴ�ҩƷ��¼  add constraint �ݴ�ҩƷ��¼_FK_ҩƷID foreign key (ҩƷID)  references ҩƷ��� (ҩƷID);
-- alter table �ݴ�ҩƷ��¼  add constraint �ݴ�ҩƷ��¼_FK_ҽ��ID foreign key (ҽ��ID, ���ͺ�)  references ����ҽ������ (ҽ��ID, ���ͺ�);

Alter table ִ�д�ӡ��¼ Add Constraint ִ�д�ӡ��¼_PK Primary Key (ҽ��ID�����ͺš���ӡʱ��) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter table ִ�д�ӡ��¼ Add constraint ִ�д�ӡ��¼_FK_ҽ��ID foreign key (ҽ��ID, ���ͺ�) references ����ҽ������ (ҽ��ID, ���ͺ�);
