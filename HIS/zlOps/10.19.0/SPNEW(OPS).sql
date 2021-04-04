
Define n_System=100;

--11927 by 陈福容 2007-11-08
Alter Table 病人手术记录 Add 接台手术 Number(1);

--11928 by 陈福容 2007-11-08
Alter Table 病人手术记录 Add 紧急程度 VarChar2(10);

--11929 by 陈福容 2007-11-08
Alter Table 病人手术记录 Add (污染手术 Number(1),感染手术 Number(1));

--11930 by 陈福容 2007-11-08
Alter Table 病人手术记录 Add (手术床 VarChar2(10),灯吊塔 VarChar2(10),层流性能 VarChar2(10));

--12454：2008-01-17 by cfr
Alter Table 病人手术记录 Add 麻醉方式id Number(18);
Alter Table 病人手术单据 Drop Constraint 病人手术单据_CK_单据类型;
Alter Table 病人手术单据 Add Constraint 病人手术单据_CK_单据类型 Check (单据类型 IN(1,2,3,4));

--11935 by 陈福容　2007-11-09
Alter Table 病人医嘱发送 Modify 收拒说明 VarChar2(1000);

Create Index 病人医嘱记录_IX_开始执行时间 On 病人医嘱记录(开始执行时间) Pctfree 5  Tablespace zl9indexcis
/

--12025 2007-12-04 by cfr
Drop Table 手术用药类型;
Create Table 手术用药类型(
    编码 VARCHAR2(2),
    名称 VARCHAR2(20),
    简码 VARCHAR2(20),
    是否麻醉剂 Number(1))
    TABLESPACE zl9BaseItem
    PCTFREE 5  PCTUSED 85;
Alter Table 手术用药类型 Add Constraint 手术用药类型_PK Primary Key (编码) Using Index Pctfree 5 Tablespace zl9indexhis;
Alter Table 手术用药类型 Add Constraint 手术用药类型_UQ_名称 Unique (名称) Using Index Pctfree 5 Tablespace zl9indexhis;


Alter Table 病人手术用药 Add 类型_Bak Varchar2(20);
Alter Table 方案用药参考 Add 类型_Bak Varchar2(20);

Update 方案用药参考 Set 类型_Bak=Decode(类型,1,'术前用药',2,'麻醉用药',3,'其他用药',类型_Bak) Where 类型_Bak Is Null;
Update 病人手术用药 Set 类型_Bak=Decode(类型,1,'术前用药',2,'麻醉用药',3,'其他用药',类型_Bak) Where 类型_Bak Is Null;

Alter Table 方案用药参考 Drop Constraint 方案用药参考_CK_类型;
Alter Table 方案用药参考 Drop Constraint 方案用药参考_PK;
Alter Table 病人手术用药 Drop Constraint 病人手术用药_PK;
Alter Table 方案用药参考 Drop Column 类型;
Alter Table 病人手术用药 Drop Column 类型;

Alter Table 病人手术用药 Add 类型 Varchar2(20);
Alter Table 方案用药参考 Add 类型 Varchar2(20);

Update 方案用药参考 Set 类型=类型_Bak;
Update 病人手术用药 Set 类型=类型_Bak;

Alter Table 方案用药参考 Add Constraint 方案用药参考_PK Primary Key (方案id,类型,药名id) Using Index Pctfree 0 Tablespace zl9indexhis;
Alter Table 病人手术用药 Add Constraint 病人手术用药_PK Primary Key (记录id,类型,药品id) Using Index Pctfree 0 Tablespace zl9indexhis;

--12144 2007-12-10 by cfr
Alter Table 手术岗位 Add 是否唯一 Number(1);
Alter Table 手术岗位 Add 是否医生 Number(1);
Alter Table 手术岗位 Add 是否护士 Number(1);

--12177 2007-12-18 by cfr
Alter Table 病人手术人员 Add 期间 Number(5);
Alter Table 病人手术记录 Add 说明 VarChar2(255);

Alter Table 病人手术人员 Drop Constraint 病人手术人员_UQ_记录id;
Alter Table 病人手术人员 Add Constraint 病人手术人员_UQ_记录id Unique (记录id,期间,科室id,岗位,编码,姓名) Using Index Pctfree 5 Tablespace zl9indexhis;

--12009 0-所有;1-病人;2-婴儿 2007-12-07 by cfr
Update 护理记录项目 Set 适用病人=Decode(项目序号,-1,2,0) Where 适用病人 Is Null;

--11953 将大小便失禁特殊符号由＇*＇改为＇※＇　2007-12-10 by cfr
Update 病人护理内容 Set 记录内容='※' Where 项目序号=10 And 记录内容='*';


--12025 2007-12-04 by cfr
Delete From 手术用药类型;
Insert Into 手术用药类型(编码,名称,简码,是否麻醉剂)
		Select '1','术前用药','SQYY',0 From Dual
Union All	Select '2','麻醉用药','MZYY',1 From Dual
Union All	Select '3','术中用药','SZYY',0 From Dual
Union All	Select '9','其他用药','QTYY',0 From Dual;

Delete From zlBaseCode Where 系统=&n_System And 表名='手术用药类型';
Insert into zlBaseCode(系统,表名,固定,说明,分类) VALUES(&n_System,'手术用药类型',0,'手术中用到的药品的类型','医技工作');

--12144 2007-12-10 by cfr
Update 手术岗位 Set 是否唯一=1 Where 名称='主刀医生';
Update 手术岗位 Set 是否医生=1 Where 名称 Like '%医生';
Update 手术岗位 Set 是否护士=1 Where 名称 Like '%护士';
Update zlBaseCode Set 固定=0 Where 系统=&n_System And 表名='手术岗位';

--12177 2007-12-18 by cfr
Update 病人手术人员 Set 期间=1 Where 期间 Is Null;


--12025 2007-12-04 by cfr
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(&n_System,1801,'基本',User,'手术用药类型','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(&n_System,1804,'基本',User,'手术用药类型','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(&n_System,1804,'基本',User,'诊疗麻醉类型','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(&n_System,1804,'基本',User,'病人新生儿记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(&n_System,1804,'基本',User,'zl_病人手术人员_Insert','EXECUTE');


--报表：ZL1_INSIDE_1804_2/术中医嘱单
Insert Into zlReports(ID,编号,名称,说明,密码,进纸,打印机,票据,系统,程序ID,功能,修改时间,发布时间) Values(zlReports_ID.NextVal,'ZL1_INSIDE_1804_2','术中医嘱单','术中医嘱单',']~!d"{vo}?$Xzpj U1LJ',15,'Epson LQ-1600K',1,&n_System,1804,'术中医嘱单',Sysdate,Sysdate);
Insert Into zlRPTFmts(报表ID,序号,说明,图样,W,H,纸张,纸向,动态纸张) Values(zlReports_ID.CurrVal,1,'术中医嘱单',0,11904,16832,9,1,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签1',2,Null,0,Null,0,'姓名:[病人信息.姓名]',Null,735,2400,1800,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签2',2,Null,0,Null,0,'性别:[病人信息.性别]',Null,2760,2400,1080,180,0,0,0,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签3',2,Null,0,Null,0,'年龄:[病人信息.年龄]',Null,4020,2400,1050,180,0,0,0,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标题',2,Null,0,'任意表1',12,'术中医嘱记录单',Null,4152,1530,3495,495,0,1,1,'宋体',24,1,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'单位名称',2,Null,0,'任意表1',12,'[单位名称]',Null,5105,1110,1590,315,0,0,1,'宋体',16,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签4',2,Null,0,Null,0,'科室:[病人信息.科室]',Null,5385,2400,1800,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签5',2,Null,0,Null,0,'床号:[病人信息.床号]',Null,7395,2400,1140,180,0,2,0,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签6',2,Null,0,Null,0,'住院号:[病人信息.住院号]',Null,8940,2400,2160,180,0,2,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'任意表1',4,Null,0,Null,0,'病人医嘱',Null,720,3015,10360,12674,420,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,4210816,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-1,0,Null,Null,'[病人医嘱.开嘱日期]','4^345^下达医嘱|4^345^日期',0,0,525,0,255,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-2,1,Null,Null,'[病人医嘱.开嘱时间]','4^345^下达医嘱|4^345^时间',0,0,660,0,255,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-3,2,Null,Null,'[病人医嘱.开嘱医生]','4^345^下达医嘱|4^345^医生',0,0,870,0,255,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-4,3,Null,Null,'[病人医嘱.医嘱内容]','4^345^术  中  医  嘱|4^345^内  容',0,0,4230,0,255,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-5,4,Null,Null,'[病人医嘱.用法]','4^345^术  中  医  嘱|4^345^内  容',0,0,1785,0,255,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-6,5,Null,Null,'[病人医嘱.校对日期]','4^345^执行医嘱|4^345^日期',0,0,615,0,255,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-7,6,Null,Null,'[病人医嘱.校对时间]','4^345^执行医嘱|4^345^时间',0,0,645,0,255,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-8,7,Null,Null,'[病人医嘱.校对护士]','4^345^执行医嘱|4^345^护士',0,0,930,0,255,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签7',2,Null,0,'任意表1',11,'手术:[病人信息.医嘱内容]',Null,720,2715,2160,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签8',2,Null,0,Null,0,'执行科室:[病人信息.执行科室]',Null,5020,2730,2520,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'病人信息','姓名,202|性别,202|年龄,202|科室,202|床号,202|住院号,131|医嘱内容,202|执行科室,202',User||'.病人信息,'||User||'.病案主页,'||User||'.病人医嘱记录,'||User||'.病人医嘱发送,'||User||'.部门表,'||User||'.诊疗项目目录',0,Null);
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,1,'Select P.姓名,P.性别,P.年龄,D.名称 As 科室,P.出院病床 As 床号,P.住院号,A.名称 As 医嘱内容,B.名称 As 执行科室');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,2,'From (Select I.姓名,I.性别,I.年龄,I.住院号,P.出院病床,P.出院科室id,V.诊疗项目id,L.执行部门id');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,3,'      From 病人信息 I,病案主页 P,病人医嘱记录 V,病人医嘱发送 L');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,4,'      Where I.病人id=V.病人id And P.病人id=V.病人id And P.主页ID=V.主页id And V.ID=[0] And L.医嘱id=V.ID) P,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,5,'      部门表 D,诊疗项目目录 A,部门表 B');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,6,'Where P.出院科室id=D.Id And A.ID=P.诊疗项目id And B.ID=P.执行部门id');
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,Null,0,'医嘱ID',1,'1',0,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'病人医嘱','排序,139|ID,131|开嘱日期,202|开嘱时间,202|开嘱医生,202|校对日期,202|校对时间,202|校对护士,202|医嘱内容,202|用法,202',User||'.病人医嘱记录,'||User||'.诊疗项目目录,'||User||'.收费项目目录',0,Null);
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,1,'--注意数据源中所提取的ID字段需要保留，用于调用程序记录已打印医嘱');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,2,'--这些ID都是可见医嘱行的ID(除西、中成药外，其他都为"相关ID=NULL"的医嘱ID)');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,3,Null);
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,4,'--"医嘱打印记录"中的数据由调用程序临时生成,适用于续打,重打,套打停止时间');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,5,'----包含当前病人(婴儿)有效的打印医嘱,含作废医嘱,不含未校对和屏蔽打印的医嘱.');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,6,Null);
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,7,'--Union ALL前面：非西、中成药品医嘱，包含中药配方(以配方用法行为准)、一并采集标本的检验(以采集方式行为准)、皮试等其他独立医嘱，包括自由录入的文本医嘱。');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,8,'--Union ALL后面：西药、中成药医嘱');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,9,Null);
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,10,'Select 1 As 排序, ID, Substr(开嘱时间, 1, 5) As 开嘱日期, Substr(开嘱时间, 7) As 开嘱时间, 开嘱医生, Substr(校对时间, 1, 5) As 校对日期,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,11,'       Substr(校对时间, 7) As 校对时间, 校对护士, 医嘱内容, 用法');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,12,'From (Select L.ID, To_Char(L.开嘱时间, ''DD/MM HH24:MI'') As 开嘱时间, L.开嘱医生, To_Char(L.校对时间, ''DD/MM HH24:MI'') As 校对时间,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,13,'              L.校对护士,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,14,'              L.医嘱内容 ||');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,15,'               Decode(I.类别 || I.操作类型, ''E4'', '''',');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,16,'                      ''  '' ||');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,17,'                       Decode(L.执行频次, ''一次性'', '''', ''持续性'', '''', ''必要时'', ''必要时'', ''不定时'', ''不定时'',');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,18,'                              Decode(L.诊疗项目id, Null, Null,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,19,'                                      ''每次'' || L.单次用量 || I.计算单位 || '','' || L.执行频次 || '',共'' || L.总给予量 || I.计算单位))) || L.皮试结果 As 医嘱内容,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,20,'              '''' As 用法');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,21,'       From 病人医嘱记录 L, 诊疗项目目录 I');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,22,'       Where L.前提id = [0] And L.诊疗项目id = I.ID(+) And');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,23,'             (L.诊疗类别 Not In (''5'', ''6'', ''7'', ''E'') Or L.诊疗类别 = ''E'' And I.操作类型 Not In (''2'', ''3'') Or I.ID Is Null) And');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,24,'             L.相关id Is Null');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,25,'       Union All');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,26,'       Select M.ID, Decode(M.序号, U.开始药品序号, To_Char(M.开嘱时间, ''DD/MM HH24:MI''), '''') As 开嘱时间,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,27,'              Decode(M.序号, U.开始药品序号, M.开嘱医生, '''') As 开嘱医生, To_Char(M.校对时间, ''DD/MM HH24:MI'') As 校对时间, M.校对护士, M.医嘱内容,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,28,'              Decode(M.序号, U.开始药品序号, U.给药, Decode(M.序号, U.结束药品序号, ''┛'', ''┃'')) As 给药');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,29,'       From (Select L.ID, L.相关id, L.序号, L.开嘱时间, L.开嘱医生, L.校对护士, L.校对时间,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,30,'                     L.医嘱内容 || ''  每次'' || 单次用量 || I.计算单位 || '',共'' || 总给予量 || E.计算单位 As 医嘱内容');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,31,'              From 病人医嘱记录 L, 诊疗项目目录 I, 收费项目目录 E');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,32,'              Where L.前提id = [0] And L.诊疗项目id = I.ID And L.收费细目id = E.ID And L.诊疗类别 In (''5'', ''6'')) M,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,33,'            (Select U.ID, U.执行频次 || '','' || U.名称 As 给药, Min(M.序号) As 开始药品序号, Max(M.序号) As 结束药品序号');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,34,'              From (Select L.ID, L.执行频次, I.名称');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,35,'                     From 病人医嘱记录 L, 诊疗项目目录 I');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,36,'                     Where L.前提id = [0] And L.诊疗项目id = I.ID And I.类别 = ''E'' And I.操作类型 = ''2'') U,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,37,'                   (Select L.序号, L.相关id');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,38,'                     From 病人医嘱记录 L');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,39,'                     Where L.前提id = [0] And L.诊疗类别 In (''5'', ''6'')) M');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,40,'              Where U.ID = M.相关id');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,41,'              Group By U.ID, U.执行频次, U.名称) U');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,42,'       Where M.相关id = U.ID)');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,43,'Order By 排序');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,44,Null);
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,Null,0,'医嘱ID',1,'1',0,Null,Null,Null,Null,Null,Null);

--报表：ZL1_INSIDE_1804_2/术中医嘱单
Insert into zlProgFuncs(系统,序号,功能,说明) Values(&n_System,1804,'术中医嘱单',Null);
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(&n_System,1804,'术中医嘱单',User,'病案主页','SELECT');
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(&n_System,1804,'术中医嘱单',User,'病人信息','SELECT');
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(&n_System,1804,'术中医嘱单',User,'病人医嘱发送','SELECT');
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(&n_System,1804,'术中医嘱单',User,'病人医嘱记录','SELECT');
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(&n_System,1804,'术中医嘱单',User,'部门表','SELECT');
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(&n_System,1804,'术中医嘱单',User,'收费项目目录','SELECT');
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(&n_System,1804,'术中医嘱单',User,'诊疗项目目录','SELECT');

--12176 2007-12-17 by cfr
--12177 2007-12-18 by cfr
CREATE OR REPLACE PROCEDURE zl_病人手术人员_Insert(
	记录id_IN	IN   病人手术人员.记录id%TYPE,
	岗位_IN	IN   病人手术人员.岗位%TYPE,
	人员id_IN	IN   病人手术人员.人员id%TYPE,
	姓名_IN	IN   病人手术人员.姓名%TYPE,
	期间_In	In    病人手术人员.期间%TYPE:=1
)
IS
BEGIN
	INSERT INTO 病人手术人员(记录id,岗位,人员id,姓名,期间)
	VALUES (记录id_IN,岗位_IN,Decode(人员id_IN,0,Null,人员id_IN),姓名_IN,期间_In);

EXCEPTION
	WHEN OTHERS THEN
		Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_病人手术人员_Insert;
/

--11928 by cfr 2007-11-08
--12012 by cfr 2007-11-19
--12177 2007-12-18 by cfr
--12454：2008-01-17 by cfr
Create Or Replace Procedure Zl_病人手术记录_Aduit
(
  Id_In       In 病人手术记录.ID%Type,
  医嘱id_In   In 病人医嘱记录.ID%Type,
  麻醉方式_In In 病人手术记录.麻醉方式%Type,
  麻醉类型_In In 病人手术记录.麻醉类型%Type,
  手术规模_In In 病人手术记录.手术规模%Type,
  麻醉方式id_In In 病人手术记录.麻醉方式id%Type
) Is
Begin
  --填写病人手术记录(手术麻醉系统的主记录)
  -------------------------------------------------------------------------------------------------------------------
  Update 病人手术记录
  Set 医嘱id = 医嘱id_In, 麻醉方式 = 麻醉方式_In, 麻醉类型 = 麻醉类型_In, 手术规模 = 手术规模_In, 手术状态 = 1, 麻醉方式id=Decode(麻醉方式id_In,0,Null,麻醉方式id_In)
  Where ID = Id_In;
  If Sql%Rowcount = 0 Then
  
    --填写病人手术记录(手术麻醉系统的主记录)
    -------------------------------------------------------------------------------------------------------------------
    Insert Into 病人手术记录
      (ID, 医嘱id, 病人id, 主页id, 手术状态, 麻醉方式, 麻醉类型, 手术规模, 紧急程度, 麻醉方式id)
      Select Id_In, 医嘱id_In, A.病人id, A.主页id, 1, 麻醉方式_In, 麻醉类型_In, 手术规模_In,
             Decode(A.紧急标志, 1, '急', ''), Decode(麻醉方式id_In,0,Null,麻醉方式id_In)
      From 病人医嘱记录 A
      Where A.ID = 医嘱id_In;
  End If;

  --审核时填写缺省的手术岗位人员
  -------------------------------------------------------------------------------------------------------------------
  Delete From 病人手术人员 Where 记录id = Id_In;
  For r_List In (Select 名称, 是否医生, 是否唯一 From 手术岗位) Loop
    If r_List.是否医生 = 1 And r_List.是否唯一 = 1 Then
      Insert Into 病人手术人员
        (记录id, 岗位, 姓名, 期间)
        Select Id_In, r_List.名称, 开嘱医生, 1 From 病人医嘱记录 Where ID = 医嘱id_In;
    Else
      Insert Into 病人手术人员
        (记录id, 岗位, 姓名, 期间)
        Select Id_In, r_List.名称, Trim(Substrb(Nvl(内容, ' '), 1, 20)), 1
        From 病人医嘱附件
        Where 医嘱id = 医嘱id_In And 项目 = r_List.名称;
    End If;
  End Loop;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病人手术记录_Aduit;
/

--12177 2007-12-18 by cfr
Create Or Replace Procedure Zl_病人手术人员_Delete
(
	记录id_In In 病人手术人员.记录id%Type,
	期间_In   In 病人手术人员.期间%Type := 0
) Is
Begin
	If 期间_In = 0 Then
		Delete From 病人手术人员 Where 记录id = 记录id_In;
	Else
		Delete From 病人手术人员 Where 记录id = 记录id_In And 期间 = 期间_In;
	End If;
Exception
	When Others Then
		Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病人手术人员_Delete;
/

--12012 by cfr 2007-11-19
CREATE OR REPLACE PROCEDURE zl_病人手术记录_AduitCancel(
	ID_In				In		病人手术记录.ID%Type
)
IS	
	v_Error		varchar2(250);
	Err_custom	Exception;
BEGIN	
	Delete From  病人手术情况 Where 记录ID=ID_In And 性质=2;
	Update 病人手术记录 Set 手术状态=Null Where ID=ID_In;

	zl_病人手术人员_Delete(ID_In);
EXCEPTION
	When Err_custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_病人手术记录_AduitCancel;
/

--11927 by cfr 2007-11-08
--11928 by cfr 2007-11-08
--11929 by cfr 2007-11-08
--11930 by cfr 2007-11-08
--11935 by cfr 2007-11-08
--12177 2007-12-18 by cfr
--12454：2008-01-17 by cfr
Create Or Replace Procedure Zl_病人手术记录_Arrange
(
  Id_In           In 病人手术记录.ID%Type,
  手术开始时间_In In 病人手术记录.手术开始时间%Type,
  手术结束时间_In In 病人手术记录.手术结束时间%Type,
  手术间_In       In 病人手术记录.手术间%Type,
  手术室id_In     In 病人手术记录.手术室id%Type := Null,
  手术人员_In     In Varchar2 := Null,
  记录性质_In     In Number := 2,
  紧急程度_In     In 病人手术记录.紧急程度%Type := Null,
  接台手术_In     In 病人手术记录.接台手术%Type := 0,
  无菌手术_In     In 病人手术记录.无菌手术%Type := 0,
  污染手术_In     In 病人手术记录.污染手术%Type := 0,
  感染手术_In     In 病人手术记录.感染手术%Type := 0
) Is
  v_Tmp    Varchar2(4000);
  v_Tmprow Varchar2(4000);
  v_Svrtmp Varchar2(50);
  n_Pos    Number(18);
  v_岗位   病人手术人员.岗位%Type;
  n_人员id 病人手术人员.人员id%Type;
  v_编码   病人手术人员.编码%Type;
  v_姓名   病人手术人员.姓名%Type;

  n_记录序号 病人医嘱发送.记录序号%Type;
  v_No       病人医嘱发送.NO%Type;
  v_麻醉no   病人医嘱发送.NO%Type;
  n_发送号   病人医嘱发送.发送号%Type;
  n_计价特性 病人医嘱记录.计价特性%Type;
  v_Error    Varchar2(255);
  Err_Custom Exception;
  n_Flag Number(1);
Begin
  --检查
  ---------------------------------------------------------------------------------------------------------------
  n_Flag := 0;
  Begin
    Select 1 Into n_Flag From 病人医嘱记录 A, 病人手术记录 B Where A.ID = B.医嘱id And B.ID = Id_In And 医嘱状态 <> 4;
  Exception
    When Others Then
      n_Flag := 0;
  End;
  If n_Flag = 0 Then
    v_Error := '手术麻醉医嘱记录已经不存在或被删除！';
    Raise Err_Custom;
  End If;

  ---------------------------------------------------------------------------------------------------------------
  n_Flag := 0;
  Begin
    Select 1
    Into n_Flag
    From 病人医嘱发送 A, 病人手术记录 B
    Where A.执行状态 > 0 And A.医嘱id = B.医嘱id And B.ID = Id_In;
  Exception
    When Others Then
      n_Flag := 0;
  End;
  If n_Flag = 1 Then
    v_Error := '手术医嘱已经发送并且正在执行或已经执行完成！';
    Raise Err_Custom;
  End If;

  --手术时间,地点填写
  ---------------------------------------------------------------------------------------------------------------
  Update 病人手术记录
  Set 手术开始时间 = 手术开始时间_In, 手术结束时间 = 手术结束时间_In, 手术日期 = Trunc(手术开始时间_In),
      手术间 = 手术间_In, 手术室id = 手术室id_In, 手术状态 = 2, 紧急程度 = 紧急程度_In, 接台手术 = 接台手术_In,
      无菌手术 = 无菌手术_In, 污染手术 = 污染手术_In, 感染手术 = 感染手术_In
  Where ID = Id_In And 手术状态 = 1;
  If Sql%Rowcount = 0 Then
    v_Error := '当前手术已经取消审核，不能继续安排操作！';
    Raise Err_Custom;
  End If;

  --修改医嘱的开始执行时间
  ---------------------------------------------------------------------------------------------------------------
  For r_Order In (Select 医嘱id From 病人手术记录 Where ID = Id_In) Loop
    Update 病人医嘱记录 Set 开始执行时间 = 手术开始时间_In Where r_Order.医嘱id In (ID, 相关id);
  End Loop;

  --手术人员的填写
  ---------------------------------------------------------------------------------------------------------------
  Delete From 病人手术人员 Where 记录id = Id_In;
  v_Tmp := 手术人员_In || ';';
  While v_Tmp Is Not Null Loop
    n_Pos := Instr(v_Tmp, ';');
    If n_Pos > 0 Then
      v_Tmprow := Substr(v_Tmp, 1, n_Pos - 1);
      v_Tmp    := Substr(v_Tmp, n_Pos + 1);
      n_Pos    := Instr(v_Tmprow, ',');
      If n_Pos > 0 Then
        n_人员id := To_Number(Substr(v_Tmprow, 1, n_Pos - 1));
        v_Tmprow := Substr(v_Tmprow, n_Pos + 1);
        n_Pos    := Instr(v_Tmprow, ',');
        If n_Pos > 0 Then
          v_岗位   := Substr(v_Tmprow, 1, n_Pos - 1);
          v_Tmprow := Substr(v_Tmprow, n_Pos + 1);
          n_Pos    := Instr(v_Tmprow, ',');
          If n_Pos > 0 Then
            v_姓名 := Substr(v_Tmprow, 1, n_Pos - 1);
            v_编码 := Substr(v_Tmprow, n_Pos + 1, 1);
            If v_岗位 Is Not Null And v_姓名 Is Not Null Then
              Zl_病人手术人员_Insert(Id_In, v_岗位, n_人员id, v_姓名, 1);
            End If;
          End If;
        End If;
      End If;
    End If;
  End Loop;

  --如果医嘱未发送,则进行医嘱发送
  ---------------------------------------------------------------------------------------------------------------
  n_记录序号 := 0;

  For r_Order In (Select A.ID, A.相关id, 执行科室id, Decode(A.计价特性, 0, 1, 1, -1, 2, 0) As 计价特性, A.诊疗类别
                  From 病人医嘱记录 A, 病人手术记录 B
                  Where A.医嘱状态 Not In (4, 8) And B.医嘱id In (A.ID, A.相关id) And B.ID = Id_In
                  Order By Decode(A.相关id, Null, 0, 1)) Loop
  
    n_记录序号 := n_记录序号 + 1;
    If n_记录序号 = 1 Then
      Select Nextno(10), Nextno(Decode(记录性质_In, 1, 13, 14)) Into n_发送号, v_No From Dual;
      n_计价特性 := r_Order.计价特性;
    End If;
  
    If v_麻醉no Is Null And r_Order.诊疗类别 = 'G' Then
      Select Nextno(Decode(记录性质_In, 1, 13, 14)) Into v_麻醉no From Dual;
    End If;
  
    If r_Order.相关id Is Null Then
      Zl_病人医嘱发送_Insert(r_Order.ID, n_发送号, 记录性质_In, v_No, n_记录序号, 1, Null, Null,
                             Sysdate + 1 / 24 / 60 / 60, 0, r_Order.执行科室id, n_计价特性, 1);
    Else
      If r_Order.诊疗类别 = 'G' Then
        Zl_病人医嘱发送_Insert(r_Order.ID, n_发送号, 记录性质_In, v_麻醉no, n_记录序号, 1, Null, Null,
                               Sysdate + 1 / 24 / 60 / 60, 0, r_Order.执行科室id, r_Order.计价特性, 0);
      Else
        Zl_病人医嘱发送_Insert(r_Order.ID, n_发送号, 记录性质_In, v_No, n_记录序号, 1, Null, Null,
                               Sysdate + 1 / 24 / 60 / 60, 0, r_Order.执行科室id, n_计价特性, 0);
      End If;
    End If;
  End Loop;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病人手术记录_Arrange;
/

--12454：2008-01-17　ＢＹ　ＣＦＲ
Create Or Replace Procedure Zl_病人手术情况_Delete
(
  记录id_In In 病人手术记录.ID%Type,
  性质_In   In 病人手术情况.性质%Type := 0
) Is
Begin
  If 性质_In = 0 Then
    Delete From 病人手术情况 Where 记录id = 记录id_In;
  Else
    Delete From 病人手术情况 Where 记录id = 记录id_In And 性质 = 性质_In;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病人手术情况_Delete;
/
--12454：2008-01-17　ＢＹ　ＣＦＲ
Create Or Replace Procedure Zl_病人手术记录_Updateadvice(Id_In In 病人手术记录.ID%Type) Is

  Cursor c_Opsrecords Is
    Select * From 病人手术记录 Where ID = Id_In;
  r_Opsrecord c_Opsrecords%Rowtype;

  v_Tmp    Varchar2(4000);
  v_Tmprow Varchar2(4000);
  v_Svrtmp Varchar2(50);

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin

  Open c_Opsrecords;
  Fetch c_Opsrecords
    Into r_Opsrecord;
  If c_Opsrecords%Rowcount = 0 Then
    Close c_Opsrecords;
    v_Error := '手术麻醉医嘱记录已经不存在或被删除！';
    Raise Err_Custom;
  End If;

  --填写手术安排信息到医生站／护士站，以便显示手术安排信息。
  ---------------------------------------------------------------------------------------------------------------
  v_Tmp    := ' ';
  v_Tmprow := ' ';
  v_Svrtmp := ' ';
  For r_List In (Select A.岗位, A.姓名
                 From 病人手术人员 A, 手术岗位 B
                 Where A.岗位 = B.名称 And A.记录id = Id_In And A.期间 = 1
                 Order By B.编码) Loop
    If v_Svrtmp <> r_List.岗位 Then
    
      If Trim(v_Svrtmp) Is Not Null Then
        v_Tmp    := Trim(v_Tmp || v_Tmprow || Chr(13) || Chr(10));
        v_Tmprow := ' ';
      End If;
      v_Svrtmp := r_List.岗位;
      v_Tmprow := Trim(v_Svrtmp || '：');
      v_Tmprow := v_Tmprow || r_List.姓名;
    Else
      v_Tmprow := v_Tmprow || ',' || r_List.姓名;
    End If;
  End Loop;
  v_Tmp := Trim(v_Tmp || v_Tmprow || Chr(13) || Chr(10));

  --手术情况记录(拟行手术)
  ---------------------------------------------------------------------------------------------------------------
  v_Tmprow := ' ';
  For r_List In (Select A.手术名称, Rownum As 序号
                 From 病人手术情况 A
                 Where A.记录id = Id_In And A.性质 = 1
                 Order By Decode(A.缺省, 1, 0, 1)) Loop
    If r_List.序号 = 1 Then
      v_Tmprow := '拟行手术：' || r_List.手术名称;
    Else
      v_Tmprow := v_Tmprow || ',' || r_List.手术名称;
    End If;
  End Loop;
  v_Tmp := Trim(v_Tmp || v_Tmprow || Chr(13) || Chr(10));

  --手术情况记录(已行手术)
  ---------------------------------------------------------------------------------------------------------------
  v_Tmprow := ' ';
  For r_List In (Select A.手术名称, Rownum As 序号
                 From 病人手术情况 A
                 Where A.记录id = Id_In And A.性质 = 2
                 Order By Decode(A.缺省, 1, 0, 1)) Loop
    If r_List.序号 = 1 Then
      v_Tmprow := '已行手术：' || r_List.手术名称;
    Else
      v_Tmprow := v_Tmprow || ',' || r_List.手术名称;
    End If;
  End Loop;
  v_Tmp := Trim(v_Tmp || v_Tmprow || Chr(13) || Chr(10));

  --更新申请说明
  ---------------------------------------------------------------------------------------------------------------
  Update 病人医嘱发送
  Set 安排时间 = r_Opsrecord.手术开始时间, 执行间 = r_Opsrecord.手术间, 收拒说明 = v_Tmp
  Where 医嘱id In (Select A.ID From 病人医嘱记录 A, 病人手术记录 B Where B.ID = Id_In And B.医嘱id In (A.ID, A.相关id));

  Close c_Opsrecords;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病人手术记录_Updateadvice;
/

--12176 取消安排时不删除人员安排情况 2007-12-17 by cfr
CREATE OR REPLACE PROCEDURE zl_病人手术记录_ArrangeCancel(
	ID_In			IN	病人手术记录.ID%TYPE
)
IS
	v_Error varchar2(255);
	Err_custom    Exception;
BEGIN
	Update 病人手术记录 Set 手术开始时间=Null,
					手术结束时间=Null,
					手术日期=Null,
					手术间=Null,
					手术间id=Null,
					手术状态=1
	Where ID=ID_In;		
EXCEPTION
	When Err_custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_病人手术记录_ArrangeCancel;
/

--11927 by 陈福容 2007-11-08
--11928 by 陈福容 2007-11-08
--11929 by 陈福容 2007-11-08
--11930 by 陈福容 2007-11-08
--12177 2007-12-18 by cfr
Create Or Replace Procedure Zl_病人手术记录_Update
(
	记录id_In       In 病人手术记录.Id%Type,
	手术日期_In     In 病人手术记录.手术日期%Type,
	手术开始时间_In In 病人手术记录.手术开始时间%Type,
	手术结束时间_In In 病人手术记录.手术结束时间%Type,
	麻醉开始时间_In In 病人手术记录.麻醉开始时间%Type,
	麻醉结束时间_In In 病人手术记录.麻醉结束时间%Type,
	麻醉方式_In     In 病人手术记录.麻醉方式%Type,
	麻醉方式id_In     In 病人手术记录.麻醉方式id%Type,
	麻醉类型_In     In 病人手术记录.麻醉类型%Type,
	麻醉质量_In     In 病人手术记录.麻醉质量%Type,
	输液总量_In     In 病人手术记录.输液总量%Type,
	输氧开始时间_In In 病人手术记录.输氧开始时间%Type,
	输氧结束时间_In In 病人手术记录.输氧结束时间%Type,
	手术间_In       In 病人手术记录.手术间%Type,
	手术室id_In     In 病人手术记录.手术室id%Type,
	手术规模_In     In 病人手术记录.手术规模%Type,
	紧急程度_In     In 病人手术记录.紧急程度%Type := Null,
	手术床_In       In 病人手术记录.手术床%Type := Null,
	灯吊塔_In       In 病人手术记录.灯吊塔%Type := Null,
	层流性能_In     In 病人手术记录.层流性能%Type := Null,
	接台手术_In     In 病人手术记录.接台手术%Type := 0,
	无菌手术_In     In 病人手术记录.无菌手术%Type := 0,
	污染手术_In     In 病人手术记录.污染手术%Type := 0,
	感染手术_In     In 病人手术记录.感染手术%Type := 0,
	说明_In     In 病人手术记录.说明%Type := Null
) Is
Begin
	Update 病人手术记录
	Set 手术日期 = 手术日期_In, 手术开始时间 = 手术开始时间_In, 手术结束时间 = 手术结束时间_In,
			麻醉开始时间 = 麻醉开始时间_In, 麻醉结束时间 = 麻醉结束时间_In, 麻醉方式 = 麻醉方式_In, 麻醉类型 = 麻醉类型_In,
			麻醉质量 = 麻醉质量_In, 输液总量 = 输液总量_In, 输氧开始时间 = 输氧开始时间_In, 输氧结束时间 = 输氧结束时间_In,
			手术间 = 手术间_In, 手术室id = 手术室id_In, 手术规模 = 手术规模_In, 手术床 = 手术床_In, 灯吊塔 = 灯吊塔_In,
			层流性能 = 层流性能_In, 无菌手术 = 无菌手术_In, 污染手术 = 污染手术_In, 感染手术 = 感染手术_In,紧急程度 = 紧急程度_In,
			接台手术 = 接台手术_In,说明=说明_In, 麻醉方式id=Decode(麻醉方式id_In,0,Null,麻醉方式id_In)
	Where Id = 记录id_In;
Exception
	When Others Then
		Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病人手术记录_Update;
/

CREATE OR REPLACE PROCEDURE ZL_病人手术情况_INSERT(
	记录ID_IN IN 病人手术记录.ID%TYPE,
	性质_IN IN 病人手术情况.性质%TYPE,
	缺省_IN IN 病人手术情况.缺省%TYPE,
	手术名称_IN IN 病人手术情况.手术名称%TYPE,
	手术操作ID_IN IN 病人手术情况.手术操作ID%TYPE,
	诊疗项目ID_IN IN 病人手术情况.诊疗项目ID%TYPE
)
IS
BEGIN	
	Insert Into 病人手术情况
		(记录ID,性质,缺省,手术名称,手术操作ID,诊疗项目ID)
		VALUES
		(记录ID_IN,性质_IN,缺省_IN,手术名称_IN,Decode(手术操作ID_IN,0,Null,手术操作ID_IN),Decode(诊疗项目ID_IN,0,Null,诊疗项目ID_IN));
EXCEPTION
	WHEN OTHERS THEN
		Zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_病人手术情况_INSERT;
/
