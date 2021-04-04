Insert Into 卡消费接口目录(编号,名称,系统,结算方式,部件,启用,自制卡,卡号长度,前缀文本) Select Max(编号)+1,'招行POS',1,'POS结算','zlZHPOS',1,2,16,Null FROM 卡消费接口目录;

insert into 结算方式 (编码,名称, 简码, 性质, 缺省标志) 	SELECT  nvl(max(to_number(编码)),0) +1, 'POS结算', 'POS', 8, 0 FROM 结算方式;
insert into 结算方式应用 (应用场合, 结算方式, 缺省标志) values ('收费', 'POS结算', 0);
insert into 结算方式应用 (应用场合, 结算方式, 缺省标志) values ('结帐', 'POS结算', 0);
insert into 结算方式应用 (应用场合, 结算方式, 缺省标志) values ('预交款', 'POS结算', 0);
