alter table 排队记录  add constraint 排队记录_PK primary key (日期, 科室ID, 顺序号)  using index  Pctfree 5 Tablespace zl9indexcis;

alter table 座位状况记录 add constraint 座位状况记录_PK primary key (科室ID, 编号) Using Index Pctfree 5 Tablespace zl9indexcis;

alter table 暂存药品记录 add constraint 暂存药品记录_PK primary key (NO, 序号, 入出系数, 登记时间) Using Index Pctfree 5 Tablespace zl9indexcis;
alter table 暂存药品记录  add constraint 暂存药品记录_FK_病人ID foreign key (病人ID) references 病人信息 (病人ID);
alter table 暂存药品记录  add constraint 暂存药品记录_FK_科室ID foreign key (科室ID) references 部门表 (ID);
-----------  
-- alter table 暂存药品记录  add constraint 暂存药品记录_FK_药品ID foreign key (药品ID)  references 药品规格 (药品ID);
-- alter table 暂存药品记录  add constraint 暂存药品记录_FK_医嘱ID foreign key (医嘱ID, 发送号)  references 病人医嘱发送 (医嘱ID, 发送号);

Alter table 执行打印记录 Add Constraint 执行打印记录_PK Primary Key (医嘱ID、发送号、打印时间) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter table 执行打印记录 Add constraint 执行打印记录_FK_医嘱ID foreign key (医嘱ID, 发送号) references 病人医嘱发送 (医嘱ID, 发送号);
