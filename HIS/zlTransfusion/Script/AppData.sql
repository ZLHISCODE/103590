-- 号码控制表
Insert Into 号码控制表(项目序号,项目名称,最大号码,自动补缺,编号规则) Values(19,'暂存药品单',Null,0,0);

-- 皮试提醒
INSERT INTO zlTools.zlNotices(序号,系统,提醒内容,提醒报表,提醒声音,提醒窗口,提醒顺序,检查周期,提醒周期,开始时间,终止时间,提醒条件)
SELECT NVL(MAX(序号),0)+1,100,'[姓名][名称]时间已到，请查看结果。',NULL+0,106,1,'[姓名];VARCHAR2|[名称];VARCHAR2',3,2,SYSDATE,NULL,
'Select e.姓名, d.名称
From 病人医嘱执行 a, 病人医嘱发送 b, 病人医嘱记录 c, 诊疗项目目录 d, 病人信息 e
Where a.组次 = 1 And a.提醒 > 0 And a.医嘱id = b.医嘱id And a.发送号 = b.发送号 And a.医嘱id = c.Id And
			c.诊疗项目id = d.Id And d.执行分类 = 3 And c.病人id = e.病人id And Sysdate Between a.执行时间 - (a.提醒 / 86400) And
			a.执行时间 And
			b.执行部门id In (Select Distinct a.部门id
											 From 部门人员 a, 人员表 b
											 Where a.人员id = b.Id And a.缺省 = 1 And Upper(b.姓名) = Upper([USER]))'
FROM zltools.zlNotices;
