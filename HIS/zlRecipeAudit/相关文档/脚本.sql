----------------
--系统脚本
----------------
Insert Into zlBakTables(系统,组号,表名,序号,直接转出,停用触发器)
Select &n_System,2,A.* From (
Select 表名,序号,直接转出,停用触发器 From zlBakTables Where 1 = 0 Union All 
  Select '处方审查明细',14,1,-Null From Dual Union All 
  Select '处方审查结果',15,1,-Null From Dual Union All 
  Select '处方审查记录',18,1,-Null From Dual Union All 
Select 表名,序号,直接转出,停用触发器 From ZLBAKTABLES Where 1 = 0) A;

Insert Into zlComponent(部件,名称,主版本,次版本,附版本,系统,注册产品名称,注册产品简名,注册产品版本) Values('zl9RecipeAudit','处方审查部件',10,35,0,&n_System,'中联医院信息系统','ZLHIS+','10');

Insert Into zlPrograms(序号,标题,说明,系统,部件) Values(1351,'门诊处方审查','药剂师对门诊医生新开的处方进行审查，只有通过审查才能收费和配发药。',&n_System,'zl9RecipeAudit'); 
Insert Into zlPrograms(序号,标题,说明,系统,部件) Values(1352,'住院药嘱审查','药剂师对住院医生新开的药嘱进行审查，只有通过审查才能配发药。',&n_System,'zl9RecipeAudit'); 
Insert Into zlPrograms(序号,标题,说明,系统,部件) Values(1353,'处方审查项目','确定门诊处方和住院药嘱，需要审查哪些项目。',&n_System,'zl9RecipeAudit'); 
Insert Into zlPrograms(序号,标题,说明,系统,部件) Values(1354,'处方审查条件','门诊处方审查的处方，按本窗体条件设置提取和开展审查。',&n_System,'zl9RecipeAudit'); 
Insert Into zlPrograms(序号,标题,说明,系统,部件) Values(1355,'处方审查统计','可分别以门诊处方和住院药嘱来统计数据，并输出报表。',&n_System,'zl9RecipeAudit'); 

------
Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select &n_System,1351,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0 Union All
  Select '基本',-NULL,NULL,1 From Dual Union All
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0) A;

Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select &n_System,1352,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0 Union All
  Select '基本',-NULL,NULL,1 From Dual Union All
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0) A;

Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select &n_System,1353,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0 Union All
  Select '基本',-NULL,NULL,1 From Dual Union All
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0) A;

Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select &n_System,1354,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0 Union All
  Select '基本',-NULL,NULL,1 From Dual Union All
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0) A;

Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select &n_System,1355,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0 Union All
  Select '基本',-NULL,NULL,1 From Dual Union All
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0) A;

------
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1351,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All
  Select '病人信息','SELECT' From Dual Union All
  Select '病人医嘱记录','SELECT' From Dual Union All
  Select '病人医嘱发送','SELECT' From Dual Union All
  Select '部门人员','SELECT' From Dual Union All
  Select '部门性质说明','SELECT' From Dual Union All
  Select '处方审查常用理由','SELECT' From Dual Union All
  Select '处方审查参数','SELECT' From Dual Union All  
  Select '处方审查项目','SELECT' From Dual Union All
  Select '处方审查条件','SELECT' From Dual Union All
  Select '处方审查记录','SELECT' From Dual Union All
  Select '处方审查明细','SELECT' From Dual Union All
  Select '处方审查结果','SELECT' From Dual Union All
  Select '门诊费用记录','SELECT' From Dual Union All
  Select '收费项目目录','SELECT' From Dual Union All
  Select '收费项目别名','SELECT' From Dual Union All
  Select 'ZL_处方审查常用理由_UPDATE','EXECUTE' From Dual Union All
  Select 'ZL_处方审查_AUDIT','EXECUTE' From Dual Union All
  Select 'ZL_处方审查_AUDIT_DETAIL','EXECUTE' From Dual Union All
  Select 'ZL_处方审查参数_SAVE','EXECUTE' From Dual Union All
  Select 'ZL_处方审查记录_LOCK','EXECUTE' From Dual Union All
  Select 'ZL_业务消息清单_INSERT','EXECUTE' From Dual Union All
  Select 'ZL_FUN_PATI_CALORIE','EXECUTE' From Dual Union All
  Select '病人挂号记录','SELECT' From Dual Union All
  Select '病人过敏记录','SELECT' From Dual Union All
  Select '病人诊断记录','SELECT' From Dual Union All
  Select '疾病编码目录','SELECT' From Dual Union All
  Select '疾病诊断目录','SELECT' From Dual Union All
  Select '诊疗项目目录','SELECT' From Dual Union All
  Select '药品规格','SELECT' From Dual Union All
  Select '病人生理情况','SELECT' From Dual Union All
  Select '病人护理记录','SELECT' From Dual Union All
  Select '病人护理内容','SELECT' From Dual Union All
  Select '诊疗频率项目','SELECT' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1352,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All
  Select '病人信息','SELECT' From Dual Union All
  Select '病人医嘱记录','SELECT' From Dual Union All
  Select '病人医嘱发送','SELECT' From Dual Union All
  Select '部门人员','SELECT' From Dual Union All
  Select '部门性质说明','SELECT' From Dual Union All
  Select '处方审查常用理由','SELECT' From Dual Union All
  Select '处方审查参数','SELECT' From Dual Union All  
  Select '处方审查项目','SELECT' From Dual Union All
  Select '处方审查记录','SELECT' From Dual Union All
  Select '处方审查明细','SELECT' From Dual Union All
  Select '处方审查结果','SELECT' From Dual Union All
  Select '收费项目目录','SELECT' From Dual Union All
  Select '收费项目别名','SELECT' From Dual Union All
  Select 'ZL_处方审查常用理由_UPDATE','EXECUTE' From Dual Union All
  Select 'ZL_处方审查_AUDIT','EXECUTE' From Dual Union All
  Select 'ZL_处方审查_AUDIT_DETAIL','EXECUTE' From Dual Union All
  Select 'ZL_处方审查参数_SAVE','EXECUTE' From Dual Union All
  Select 'ZL_处方审查记录_LOCK','EXECUTE' From Dual Union All
  Select 'ZL_业务消息清单_INSERT','EXECUTE' From Dual Union All
  Select 'ZL_FUN_PATI_CALORIE','EXECUTE' From Dual Union All
  Select '病案主页','SELECT' From Dual Union All
  Select '病案主页从表','SELECT' From Dual Union All
  Select '病人过敏记录','SELECT' From Dual Union All
  Select '病人诊断记录','SELECT' From Dual Union All
  Select '疾病编码目录','SELECT' From Dual Union All
  Select '疾病诊断目录','SELECT' From Dual Union All
  Select '诊疗项目目录','SELECT' From Dual Union All
  Select '药品规格','SELECT' From Dual Union All
  Select '病人生理情况','SELECT' From Dual Union All
  Select '病人护理记录','SELECT' From Dual Union All
  Select '病人护理内容','SELECT' From Dual Union All
  Select '诊疗频率项目','SELECT' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1353,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All
  Select '处方审查项目','SELECT' From Dual Union All
  Select '处方审查项目_ID','SELECT' From Dual Union All
  Select 'ZL_处方审查项目_UPDATE','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1354,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All
  Select '部门性质说明','SELECT' From Dual Union All
  Select '人员性质说明','SELECT' From Dual Union All
  Select '诊疗分类目录','SELECT' From Dual Union All
  Select '诊疗项目目录','SELECT' From Dual Union All
  Select '药品特性','SELECT' From Dual Union All
  Select '疾病编码目录','SELECT' From Dual Union All
  Select '疾病诊断目录','SELECT' From Dual Union All
  Select '处方审查条件','SELECT' From Dual Union All
  Select 'ZL_处方审查条件_UPDATE','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1355,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All
  Select '部门性质说明','SELECT' From Dual Union All
  Select '人员性质说明','SELECT' From Dual Union All
  Select '病人医嘱记录','SELECT' From Dual Union All
  Select '收费项目目录','SELECT' From Dual Union All
  Select '处方审查项目','SELECT' From Dual Union All
  Select '处方审查记录','SELECT' From Dual Union All
  Select '处方审查明细','SELECT' From Dual Union All
  Select '处方审查结果','SELECT' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

--[[zlModuleRelas]]
Insert Into zlModuleRelas(系统,模块,功能,相关系统,相关模块,相关类型,相关功能,缺省值)
Select &n_System,1351,A.* From (
Select 功能,相关系统,相关模块,相关类型,相关功能,缺省值 From zlModuleRelas Where 1 = 0 Union All
  Select NULL,&n_System,9001,1,'基本',1 From Dual Union All
Select 功能,相关系统,相关模块,相关类型,相关功能,缺省值 From zlModuleRelas Where 1 = 0) A;

Insert Into zlModuleRelas(系统,模块,功能,相关系统,相关模块,相关类型,相关功能,缺省值)
Select &n_System,1352,A.* From (
Select 功能,相关系统,相关模块,相关类型,相关功能,缺省值 From zlModuleRelas Where 1 = 0 Union All
  Select NULL,&n_System,9001,1,'基本',1 From Dual Union All
Select 功能,相关系统,相关模块,相关类型,相关功能,缺省值 From zlModuleRelas Where 1 = 0) A;

--[[zlProgRelas]]
--无


------
Insert Into zlMenus
  (组别, ID, 上级id, 标题, 快键, 说明, 系统, 模块, 短标题, 图标)
  Select 组别, Zlmenus_Id.Nextval, ID, '处方审查系统', Null, '药剂师对医生新开的门诊处方或住院药嘱审查的系统。', &n_System, -Null, '处方审查系统', 图标
  From zlMenus
  Where 标题 = '医疗管理系统' And 组别 = '缺省' And 系统 = &n_System And 模块 Is Null;

Insert Into zlMenus(组别,ID,上级ID,标题,快键,说明,系统,模块,短标题,图标) Select A.组别,ZlMenus_ID.Nextval,A.ID,B.* From (
Select 组别,ID From zlMenus Where 标题 = '处方审查系统' And 组别 = '缺省' And 系统 = &n_System And 模块 Is Null) A,(Select 标题,快键,说明,系统,模块,短标题,图标 From zlMenus Where 1 = 0 Union All
  Select '门诊处方审查', Null,'药剂师对门诊医生新开的处方进行审查，只有通过审查才能收费和发药。' ,&n_System,1351,'门诊处方审查' ,232 From Dual Union All
  Select '住院药嘱审查', Null,'药剂师对住院医生新开的药嘱进行审查，只有通过审查才能计费和发药。' ,&n_System,1352,'住院药嘱审查' ,234 From Dual Union All
  Select '处方审查项目', Null,'确定门诊处方和住院药嘱，需要审查哪些项目。' ,&n_System,1353,'处方审查项目' ,193 From Dual Union All
  Select '处方审查条件', Null,'门诊处方审查的处方，按本窗体条件设置提取和开展审查。' ,&n_System,1354,'处方审查条件' ,210 From Dual Union All
  Select '处方审查统计', Null,'可分别以门诊处方和住院药嘱来统计数据，并输出报表。' ,&n_System,1355,'处方审查统计' ,179 From Dual Union All
Select 标题,快键,说明,系统,模块,短标题,图标 From zlMenus Where 1 = 0) B;
  
Insert Into zlParameters(ID,系统,模块,私有,本机,授权,固定,部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明)
Select zlParameters_ID.Nextval,&n_System,-Null,A.* From (
Select 私有,本机,授权,固定,部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明 From zlParameters Where 1 = 0 Union All 
  Select 0,0,0,0,0,0,241,'处方审查',Null,'0','是否启用处方审查系统功能与流程控制；','0-门诊和住院都不启用；1-门诊启用，住院不启用；2-门诊不启用，住院启用；3-门诊和住院都启用',Null,Null,Null From Dual Union All
  Select 0,0,0,0,0,0,242,'门诊审方时机',Null,'1','确定门诊药师审方的介入时机；','1-处方发送前；2-药房配发药前',Null,Null,Null From Dual Union All 
  Select 0,0,0,0,0,0,243,'门诊药师离岗时长',Null,'10','门诊药师审方离岗时长（单位分钟），即：超出设定时长值未审查的处方，医师可发送通过，避免病人长时间滞留临床科室或药房。',Null,Null,Null,Null From Dual Union All
  Select 0,0,0,0,0,0,244,'处方审查依据',Null,'1','确定门诊处方/住院药嘱的审查依据什么开展。','1-依据《处方点评管理规范》28项；2-依据《处方管理办法》7项',Null,Null,Null From Dual Union All
  Select 0,0,0,0,0,0,245,'提醒门诊医生不合格医嘱',Null,'0','门诊医生开方保存后，是否开启提醒医生有问题的药嘱。','0-不提醒；1-提醒',Null,Null,Null From Dual Union All
  Select 0,0,0,0,0,0,246,'提醒住院医生不合格医嘱',Null,'0','住院医生开方保存后，是否开启提醒医生有问题的药嘱。','0-不提醒；1-提醒',Null,Null,Null From Dual Union All
Select 私有,本机,授权,固定,部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明 From zlParameters Where 1 = 0) A;

----------------
--数据结构
----------------
Create Sequence 处方审查条件_Id Start With 1;
Create Sequence 处方审查记录_Id Start With 1;
Create Sequence 处方审查项目_Id Start With 1;

Create Table 处方审查参数(
       机器名 Varchar2(15), 
       服务对象 Number(1),
       是否开启审方 Number(1), 
       最后操作时间 Date,
       来源科室 Varchar2(4000)) 
Pctfree 10 Initrans 1 
Tablespace Zl9baseitem;

Create Table 处方审查条件(
       ID Number(18), 
       类别 Number(2), 
       药名id Number(18),
       科室id Number(18), 
       医生id Number(18), 
       诊断id Number(18), 
       疾病id Number(18)) 
Pctfree 10 Initrans 1 
Tablespace Zl9baseitem;

Create Table 处方审查项目(
       ID Number(18), 
       类别 Number(1),
       编码 Varchar2(10), 
       简称 Varchar2(50),
       内容 Varchar2(500), 
       是否门诊启用 Number(1), 
       是否住院启用 Number(1), 
       服务对象 Number(1),
       PASS结果 Varchar2(50),
       操作人 Varchar2(100), 
       操作时间 Date, 
       作废时间 Date)
Pctfree 10 Initrans 1 
Tablespace Zl9baseitem;

Create Table 处方审查常用理由(
       用户名 Varchar2(20),
       内容 Varchar2(500))
Pctfree 10 Initrans 1 
Tablespace Zl9medlst;

Create Table 处方审查记录(
       ID Number(18),
       病人id Number(18), 
       挂号id Number(18), 
       主页id Number(18),
       提交人 Varchar2(100), 
       提交时间 Date, 
       提交科室id Number(18),
       审查结果 Number(1),
       审查人 Varchar2(100), 
       审查时间 Date, 
       发药药房id Number(18), 
       综合理由 Varchar2(500),
       状态 Number(1),
       锁定用户 Varchar2(20),
       锁定时间 Date,
       待转出 Number(3)) 
Pctfree 10 Initrans 20 
Tablespace Zl9medlst;

Create Table 处方审查明细(
       审方id Number(18), 
       医嘱id Number(18), 
       最后提交 Number(1), 
       待转出 Number(3)) 
Pctfree 10 Initrans 20 
Tablespace Zl9medlst;

Create Table 处方审查结果(
       审方id Number(18), 
       医嘱id Number(18), 
       审查项目id Number(18),
       最后提交 Number(1),
       药师审查 Number(1),
       自动审查 Number(1),
       理由 Varchar2(500),
       待转出 Number(3)) 
Pctfree 10 Initrans 20 
Tablespace Zl9medlst;

Alter Table 处方审查参数 Add Constraint 处方审查参数_Pk Primary Key(机器名, 服务对象) Using Index Tablespace Zl9indexhis;
Alter Table 处方审查条件 Add Constraint 处方审查条件_Pk Primary Key(ID) Using Index Tablespace Zl9indexhis;
--Alter Table 处方审查条件 Add Constraint 处方审查条件_Uq_科室id Unique(科室id, 类别) Using Index Tablespace Zl9indexhis;
--Alter Table 处方审查条件 Add Constraint 处方审查条件_Uq_医生id Unique(医生id, 类别) Using Index Tablespace Zl9indexhis;
--Alter Table 处方审查条件 Add Constraint 处方审查条件_Uq_诊断id Unique(诊断id, 类别) Using Index Tablespace Zl9indexhis;
--Alter Table 处方审查条件 Add Constraint 处方审查条件_Uq_疾病id Unique(疾病id, 类别) Using Index Tablespace Zl9indexhis;
--Alter Table 处方审查条件 Add Constraint 处方审查条件_Uq_药名id Unique(药名id, 类别) Using Index Tablespace Zl9indexhis;
Alter Table 处方审查项目 Add Constraint 处方审查项目_Pk Primary Key(ID) Using Index Tablespace Zl9indexhis;
Alter Table 处方审查记录 Add Constraint 处方审查记录_Pk Primary Key(ID) Using Index Tablespace Zl9indexhis;
Alter Table 处方审查明细 Add Constraint 处方审查明细_Pk Primary Key(审方id, 医嘱id) Using Index Tablespace Zl9indexhis;
Alter Table 处方审查常用理由 Add Constraint 处方审查常用理由_Pk Primary Key(用户名, 内容) Using Index Tablespace Zl9indexhis;

Alter Table 处方审查项目 Add Constraint 处方审查项目_Uq_编码 Unique(编码) Using Index Tablespace Zl9indexhis;
Alter Table 处方审查项目 Add Constraint 处方审查项目_Uq_简称 Unique(简称) Using Index Tablespace Zl9indexhis;
Alter Table 处方审查结果 Add Constraint 处方审查结果_Uq_审方Id Unique(审方id, 医嘱id, 审查项目id) Using Index Tablespace Zl9indexhis;
Alter Table 处方审查记录 Add Constraint 处方审查记录_Uq_提交时间 Unique(提交时间, 提交科室id, 病人id, 发药药房id) Using Index Tablespace Zl9indexhis;

Alter Table 处方审查条件 Modify 类别 Constraint 处方审查条件_NN_类别 not null;
Alter Table 处方审查项目 Modify 编码 Constraint 处方审查项目_NN_编码 Not Null;
Alter Table 处方审查项目 Modify 简称 Constraint 处方审查项目_NN_简称 Not Null;
Alter Table 处方审查结果 Modify 审方Id Constraint 处方审查结果_NN_审方Id not null;

Alter Table 处方审查条件 Add Constraint 处方审查条件_Fk_科室ID Foreign Key(科室id) References 部门表(ID) On Delete Cascade enable novalidate;
Alter Table 处方审查条件 Add Constraint 处方审查条件_Fk_医生ID Foreign Key(医生id) References 人员表(ID) On Delete Cascade enable novalidate;
Alter Table 处方审查条件 Add Constraint 处方审查条件_Fk_诊断ID Foreign Key(诊断id) References 疾病诊断目录(ID) On Delete Cascade enable novalidate;
Alter Table 处方审查条件 Add Constraint 处方审查条件_Fk_疾病ID Foreign Key(疾病id) References 疾病编码目录(ID) On Delete Cascade enable novalidate;
Alter Table 处方审查条件 Add Constraint 处方审查条件_Fk_药名ID Foreign Key(药名id) References 诊疗项目目录(ID) On Delete Cascade enable novalidate;
Alter Table 处方审查记录 Add Constraint 处方审查记录_FK_病人Id Foreign Key(病人Id) References 病人信息(病人ID) enable novalidate;
Alter Table 处方审查记录 Add Constraint 处方审查记录_FK_挂号Id Foreign Key(挂号Id) References 病人挂号记录(ID) enable novalidate;
Alter Table 处方审查记录 Add Constraint 处方审查记录_FK_提交科室Id Foreign Key(提交科室Id) References 部门表(ID) enable novalidate;
Alter Table 处方审查记录 Add Constraint 处方审查记录_FK_发药药房Id Foreign Key(发药药房Id) References 部门表(ID) enable novalidate;
Alter Table 处方审查明细 Add Constraint 处方审查明细_Fk_审方id Foreign Key(审方id) References 处方审查记录(ID) On Delete Cascade enable novalidate;
Alter Table 处方审查明细 Add Constraint 处方审查明细_Fk_医嘱id Foreign Key(医嘱id) References 病人医嘱记录(ID) enable novalidate;
Alter Table 处方审查结果 Add Constraint 处方审查结果_Fk_审方id Foreign Key(审方id) References 处方审查记录(ID) On Delete Cascade enable novalidate;
Alter Table 处方审查结果 Add Constraint 处方审查结果_Fk_审查项目id Foreign Key(审查项目id) References 处方审查项目(ID) enable novalidate;
Alter Table 处方审查结果 Add Constraint 处方审查结果_Fk_医嘱id Foreign Key(医嘱id) References 病人医嘱记录(ID) Enable Novalidate;

Create Index 处方审查记录_Ix_挂号id On 处方审查记录(挂号id) Tablespace Zl9indexhis;
Create Index 处方审查记录_Ix_病人id On 处方审查记录(病人id, 主页id) Tablespace Zl9indexhis;
Create Index 处方审查记录_Ix_审查时间 On 处方审查记录(审查时间) Tablespace Zl9indexhis;
Create Index 处方审查记录_Ix_状态 On 处方审查记录(状态) Tablespace Zl9indexhis;
Create Index 处方审查记录_Ix_锁定用户 On 处方审查记录(锁定用户) Tablespace Zl9indexhis;
Create Index 处方审查记录_Ix_待转出 On 处方审查记录(待转出) Tablespace Zl9indexhis;
Create Index 处方审查明细_Ix_医嘱id On 处方审查明细(医嘱id) Tablespace Zl9indexhis;
Create Index 处方审查明细_IX_待转出 ON 处方审查明细(待转出) Tablespace Zl9indexhis;
Create Index 处方审查结果_Ix_审查项目id On 处方审查结果(审查项目id) Tablespace Zl9indexhis;
Create Index 处方审查结果_IX_待转出 ON 处方审查结果(待转出) Tablespace Zl9indexhis;
  




CREATE OR REPLACE Procedure Zl_处方审查条件_Update
(
  类别_In In 处方审查条件.类别%Type,
  序号_In In Number,
  Ids_In  In Varchar2
) Is

  --v_Err_Msg Varchar2(2000);
  --Err_Item Exception;

Begin

  If 序号_In = 1 Then
    Delete 处方审查条件;
  End If;

  If 类别_In = 1 Then
    --科室
    Insert Into 处方审查条件
      (ID, 类别, 科室id)
      Select 处方审查条件_Id.Nextval, 类别_In, Column_Value From Table(f_Num2list(Ids_In));
  Elsif 类别_In = 2 Then
    --医生
    Insert Into 处方审查条件
      (ID, 类别, 医生id)
      Select 处方审查条件_Id.Nextval, 类别_In, Column_Value From Table(f_Num2list(Ids_In));
  Elsif 类别_In = 3 Then
    --诊断
    Insert Into 处方审查条件
      (ID, 类别, 诊断id)
      Select 处方审查条件_Id.Nextval, 类别_In, Column_Value From Table(f_Num2list(Ids_In));
  Elsif 类别_In = 4 Then
    --疾病
    Insert Into 处方审查条件
      (ID, 类别, 疾病id)
      Select 处方审查条件_Id.Nextval, 类别_In, Column_Value From Table(f_Num2list(Ids_In));
  Elsif 类别_In = 5 Then
    --药名
    Insert Into 处方审查条件
      (ID, 类别, 药名id)
      Select 处方审查条件_Id.Nextval, 类别_In, Column_Value From Table(f_Num2list(Ids_In));
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_处方审查条件_Update;
/

Create Or Replace Procedure Zl_处方审查项目_Update
(
  项目id_In       In 处方审查项目.Id%Type,
  类别_In         In 处方审查项目.类别%Type,
  编码_In         In 处方审查项目.编码%Type,
  简称_In         In 处方审查项目.简称%Type,
  内容_In         In 处方审查项目.内容%Type,
  是否门诊启用_In In 处方审查项目.是否门诊启用%Type,
  是否住院启用_In In 处方审查项目.是否住院启用%Type,
  服务对象_In     In 处方审查项目.服务对象%Type,
  Pass结果_In     In 处方审查项目.Pass结果%Type,
  操作人_In       In 处方审查项目.操作人%Type,
  是否作废_In     In Number := Null
) Is

  n_Count Number(18);
  n_Id    Number(18);

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

Begin

  --检查项目ID是否存在
  Select Count(1) Into n_Count From 处方审查项目 Where ID = 项目id_In;

  If n_Count = 0 Then
    If 类别_In = 4 Then
      --类别“4=自定义”
      If Nvl(项目id_In, 0) <= 0 Then
        --产生ID
        Select 处方审查项目_Id.Nextval Into n_Id From Dual;
      Else
        n_Id := 项目id_In;
      End If;
      Insert Into 处方审查项目
        (ID, 类别, 编码, 简称, 内容, 是否门诊启用, 是否住院启用, 服务对象, Pass结果, 操作人, 操作时间, 作废时间)
      Values
        (n_Id, 类别_In, 编码_In, 简称_In, 内容_In, 是否门诊启用_In, 是否住院启用_In, 服务对象_In, Null, 操作人_In, Sysdate, Null);
    Else
      v_Err_Msg := '未找到项目数据！';
      Raise Err_Item;
    End If;
  Else
    If 类别_In = 1 Then
      --1=处方管理办法7项
      Update 处方审查项目
      Set 是否门诊启用 = 是否门诊启用_In, 是否住院启用 = 是否住院启用_In, 操作人 = 操作人_In, 操作时间 = Sysdate
      Where ID = 项目id_In;
    Elsif 类别_In = 2 Then
      --2=处方点评管理规范28项
      Update 处方审查项目
      Set 是否门诊启用 = 是否门诊启用_In, 是否住院启用 = 是否住院启用_In, 操作人 = 操作人_In, 操作时间 = Sysdate
      Where ID = 项目id_In;
    Elsif 类别_In = 3 Then
      --3=固定
      Update 处方审查项目
      Set 是否门诊启用 = 是否门诊启用_In, 是否住院启用 = 是否住院启用_In, Pass结果 = Pass结果_In, 操作人 = 操作人_In, 操作时间 = Sysdate
      Where ID = 项目id_In;
    Elsif 类别_In = 4 Then
      --4=自定义
      If 是否作废_In = 1 Then
        --检查“处方审查结果”是否使用
        Select Count(1) Into n_Count From 处方审查结果 Where 审查项目id = 项目id_In;
        If n_Count <= 0 Then
          Delete 处方审查项目 Where ID = 项目id_In;
        Else
          Update 处方审查项目 Set 作废时间 = Sysdate Where ID = 项目id_In;
        End If;
      Else
        Update 处方审查项目
        Set 编码 = 编码_In, 简称 = 简称_In, 内容 = 内容_In, 是否门诊启用 = 是否门诊启用_In, 是否住院启用 = 是否住院启用_In, 操作人 = 操作人_In, 操作时间 = Sysdate
        Where ID = 项目id_In;
      End If;
    Else
      v_Err_Msg := '项目的类别不正确！';
      Raise Err_Item;
    End If;
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_处方审查项目_Update;
/

CREATE OR REPLACE Procedure Zl_处方审查_Insert
(
  审方id_In     In 处方审查记录.Id%Type,
  病人id_In     In 处方审查记录.病人id%Type,
  挂号id_In     In 处方审查记录.挂号id%Type,
  主页id_In     In 处方审查记录.主页id%Type,
  提交科室id_In In 处方审查记录.提交科室id%Type,
  提交人_In     In 处方审查记录.提交人%Type,
  发药药房id_In In 处方审查记录.发药药房id%Type,
  医嘱id_In     In Varchar2
) Is

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

Begin

  --插入待审查记录
  If Nvl(挂号id_In, 0) > 0 Then
    Insert Into 处方审查记录
      (ID, 病人id, 挂号id, 提交人, 提交时间, 提交科室id, 发药药房id, 状态)
    Values
      (审方id_In, 病人id_In, 挂号id_In, 提交人_In, Sysdate, 提交科室id_In, 发药药房id_In, 0);
  Else
    Insert Into 处方审查记录
      (ID, 病人id, 主页id, 提交人, 提交时间, 提交科室id, 发药药房id, 状态)
    Values
      (审方id_In, 病人id_In, 主页id_In, 提交人_In, Sysdate, 提交科室id_In, 发药药房id_In, 0);
  End If;

  --插入待审查记录对应的医嘱
  For r_Medical In (Select /*+RULE*/
                     ID
                    From 病人医嘱记录 A, Table(f_Num2list(医嘱id_In, ',')) B
                    Where a.Id = b.Column_Value) Loop
    --先修改旧医嘱id的最后提交
    If r_Medical.Id Is Not Null Then
      Update 处方审查明细 Set 最后提交 = Null Where 医嘱id = r_Medical.Id;
    
      Insert Into 处方审查明细 (审方id, 医嘱id, 最后提交) Values (审方id_In, r_Medical.Id, 1);
    End If;
  
  End Loop;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_处方审查_Insert;
/

CREATE OR REPLACE Procedure Zl_处方审查_Auto
(
  审方id_In         In 处方审查结果.审方id%Type,
  自动审查_In       In 处方审查结果.自动审查%Type,
  审查项目与医嘱_In In Varchar2
) Is

  v_Err_Msg Varchar2(255);
  Err_Item Exception;
Begin

  For r_Info In (Select /*+RULE*/
                  C1 审查项目id, C2 医嘱id
                 From Table(f_Num2list2(审查项目与医嘱_In, '|', ','))) Loop
    If r_Info.医嘱id Is Not Null Then
      --先修改旧医嘱id的最后提交
      Update 处方审查结果 Set 最后提交 = Null Where 医嘱id = r_Info.医嘱id;
    
      Insert Into 处方审查结果
        (审方id, 医嘱id, 审查项目id, 最后提交, 自动审查)
      Values
        (审方id_In, Decode(r_Info.医嘱id, 0, Null, r_Info.医嘱id), r_Info.审查项目id, 1, 自动审查_In);
    End If;
  
  End Loop;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_处方审查_Auto;
/

Create Or Replace Procedure Zl_处方审查_Cancel
(
  医嘱id_In  In Varchar2,
  锁定id_Out Out Varchar2
) Is

  n_Count   Number;
  v_Lockid  Varchar2(4000);
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

Begin

  --取医嘱对应的审方ID
  For r_Info In (Select /* +RULE*/
                 Distinct b.Id, b.状态
                 From 处方审查明细 A, 处方审查记录 B, 病人医嘱记录 C, Table(f_Num2list(医嘱id_In, ',')) D
                 Where a.审方id = b.Id And a.医嘱id = c.Id And c.相关id = d.Column_Value And a.最后提交 = 1 And
                       (b.状态 Between 0 And 1 Or b.状态 Is Null) And c.诊疗类别 In ('5', '6', '7')) Loop
  
    Select Count(1) Into n_Count From 处方审查记录 Where ID = r_Info.Id And 锁定用户 Is Not Null;
    If n_Count = 0 Then
      --未锁定
      If Nvl(r_Info.状态, 0) = 0 Then
        --未审查，直接删除记录
        Delete 处方审查记录 Where ID = r_Info.Id And (状态 = 0 Or 状态 Is Null);
      Elsif r_Info.状态 = 1 Then
        --已审查，调整状态
        Update 处方审查记录 Set 状态 = 状态 + 10 Where 状态 = 1;
      End If;
    Else
      --被锁定
      Begin
        Select f_List2str(Cast(Collect(Cast(医嘱id As Varchar2(20))) As t_Strlist), ',')
        Into v_Lockid
        From 处方审查明细
        Where 审方id = r_Info.Id
        Order By 医嘱id;
      Exception
        When Others Then
          v_Lockid := Null;
      End;
    
      If v_Lockid Is Not Null Then
        锁定id_Out := 锁定id_Out || ',' || v_Lockid;
      End If;
    End If;
  
  End Loop;

  If Substr(锁定id_Out, 1, 1) = ',' Then
    锁定id_Out := Substr(锁定id_Out, 2, Length(锁定id_Out));
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_处方审查_Cancel;
/

CREATE OR REPLACE Procedure Zl_处方审查_Update
(
  业务类别_In In Number,
  科室id_In   In 部门表.Id%Type,
  审方id_In   In 处方审查记录.Id%Type
) Is
  --功能：按业务类别，更新处方审查记录的状态
  --参数：
  --  业务类别_In：1-门诊业务；2-住院业务
  --  科室id_In：临床科室ID
  --  审议ID_In：略

  n_Param1  Number(10);
  n_Param2  Number(10);
  n_Count   Number(18);
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

Begin

  Select Nvl(zl_GetSysParameter('门诊审方时机'), '0') Into n_Param1 From Dual;
  Select Nvl(zl_GetSysParameter('门诊药师离岗时长'), '10') Into n_Param2 From Dual;

  If 业务类别_In = 1 Then
    --门诊业务
    Select Count(1)
    Into n_Count
    From (Select Max(最后操作时间) 最后操作时间
           From 处方审查参数
           Where Nvl(服务对象, 0) = 0 And ',' || 来源科室 || ',' Like '%,' || 科室id_In || ',%')
    Where 最后操作时间 <= (Sysdate - n_Param2 / 24 / 60);
  
    If n_Count > 0 Then
      --长时间未审查，标记2-超时免审
      Update 处方审查记录
      Set 状态 = 2, 锁定用户 = Null, 锁定时间 = Null
      Where (状态 = 0 Or 状态 Is Null) And ID = 审方id_In;
    End If;
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_处方审查_Update;
/

CREATE OR REPLACE Procedure Zl_处方审查参数_Save
(
  类别_In         In Number,
  机器名_In       In 处方审查参数.机器名%Type,
  服务对象_In     In 处方审查参数.服务对象%Type,
  是否开启审方_In In 处方审查参数.是否开启审方%Type := Null,
  来源科室_In     In 处方审查参数.来源科室%Type := Null
) Is

  --功能：保存处方审查参数
  --参数：
  --  类别_In：1-保存来源科室；2-保存最后操作时间
  --  服务对象_In：0-门诊；1-住院
  --  是否开启审方_In：类别_In = 2，该参数有用
  --  来源科室_In：类别_In = 1，该参数有有

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

Begin

  If 类别_In = 1 Then
    --保存来源科室
    Update 处方审查参数 Set 来源科室 = 来源科室_In Where 机器名 = 机器名_In And 服务对象 = 服务对象_In;
  
    If Sql%RowCount = 0 Then
      Insert Into 处方审查参数
        (机器名, 服务对象, 是否开启审方, 最后操作时间, 来源科室)
      Values
        (机器名_In, 服务对象_In, 0, Sysdate, 来源科室_In);
    End If;
  Elsif 类别_In = 2 Then
    --保存是否开启审方、最后操作时间
    Update 处方审查参数
    Set 是否开启审方 = 是否开启审方_In, 最后操作时间 = Sysdate
    Where 机器名 = 机器名_In And 服务对象 = 服务对象_In;
  
    If Sql%RowCount = 0 Then
      Insert Into 处方审查参数
        (机器名, 服务对象, 是否开启审方, 最后操作时间, 来源科室)
      Values
        (机器名_In, 服务对象_In, 是否开启审方_In, Sysdate, Null);
    End If;
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_处方审查参数_Save;
/

CREATE OR REPLACE Procedure Zl_处方审查常用理由_Update
(
  功能号_In In Number,
  用户名_In In 处方审查常用理由.用户名%Type,
  内容_In   In 处方审查常用理由.内容%Type
) Is

  --功能：新增、删除处方审查常用理由
  --参数：
  --  功能号_In：1-新增；0-删除

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

Begin

  If 功能号_In = 0 Then
    Delete 处方审查常用理由 Where 用户名 = 用户名_In And 内容 = 内容_In;
  Elsif 功能号_In = 1 Then
    Insert Into 处方审查常用理由 (用户名, 内容) Values (用户名_In, 内容_In);
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_处方审查常用理由_Update;
/

Create Or Replace Procedure Zl_处方审查记录_Lock
(
  Lock_In     In Number,
  机器名_In   In 处方审查参数.机器名%Type,
  服务对象_In In 处方审查参数.服务对象%Type,
  审方id_In   In 处方审查记录.Id%Type := Null
) Is
  --功能：处方审查记录的加锁、解锁切换
  --参数：
  --  Lock_In：0-解锁；1-加锁
  --  服务对象_In：0-门诊；1-住院

  n_Count   Number(2);
  v_Err_Msg Varchar2(255);
  Err_Item Exception;
Begin

  If Lock_In = 1 Then
  
    --先检查
    Select Count(1) Into n_Count From 处方审查记录 Where ID = 审方id_In;
    If n_Count <= 0 Then
      v_Err_Msg := '该处方审查记录已被删除！';
      Raise Err_Item;
    End If;
    
    Select Count(1) Into n_Count From 处方审查记录 Where ID = 审方id_In And 状态 = 0;
    If n_Count <= 0 Then
      v_Err_Msg := '该处方审查记录已被审查！';
      Raise Err_Item;
    End If;

    --加锁
    If Nvl(审方id_In, 0) > 0 Then
      Update 处方审查参数 Set 最后操作时间 = Sysdate Where 机器名 = 机器名_In And 服务对象 = 服务对象_In;
    
      Update 处方审查记录
      Set 锁定用户 = Upper(User), 锁定时间 = Sysdate
      Where ID = 审方id_In And (锁定用户 Is Null Or 锁定用户 = Upper(User));
      If Sql%NotFound Then
        v_Err_Msg := '该处方审查记录已被其他人锁定！';
        Raise Err_Item;
      End If;
    End If;
  
  Else
  
    --解锁
    Update 处方审查参数 Set 最后操作时间 = Sysdate Where 机器名 = 机器名_In And 服务对象 = 服务对象_In;
  
    If Nvl(审方id_In, 0) = 0 Then
      --所有当前用户的锁定记录
      Update 处方审查记录 Set 锁定用户 = Null, 锁定时间 = Null Where 锁定用户 = Upper(User);
    Else
      Update 处方审查记录 Set 锁定用户 = Null, 锁定时间 = Null Where ID = 审方id_In And 锁定用户 = Upper(User);
      If Sql%NotFound Then
        v_Err_Msg := '该处方审查记录已被其他人解锁！';
        Raise Err_Item;
      Else
        --解锁超一小时锁定的记录
        Update 处方审查记录
        Set 锁定用户 = Null, 锁定时间 = Null
        Where 锁定时间 < Sysdate - 1 / 24 And 锁定用户 = Upper(User);
      End If;
    End If;
  
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_处方审查记录_Lock;
/

Create Or Replace Procedure Zl_处方审查_Audit
(
  审方id_In   In 处方审查记录.Id%Type,
  审查结果_In In 处方审查记录.审查结果%Type,
  审查人_In   In 处方审查记录.审查人%Type,
  综合理由_In In 处方审查记录.综合理由%Type
) Is
  --功能：提交处方审查记录信息

  v_Err_Msg Varchar2(255);
  Err_Item Exception;
Begin

  Update 处方审查记录
  Set 审查结果 = 审查结果_In, 审查人 = 审查人_In, 审查时间 = Sysdate, 综合理由 = 综合理由_In, 状态 = 1, 锁定用户 = Null, 锁定时间 = Null
  Where ID = 审方id_In And 审查时间 Is Null;
  If Sql%NotFound Then
    v_Err_Msg := '该病人的记录已被其他人审查！';
    Raise Err_Item;
  End If; 

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_处方审查_Audit;
/

Create Or Replace Procedure Zl_处方审查_Audit_Detail
(
  审方id_In     In 处方审查结果.审方id%Type,
  医嘱id_In     In 处方审查结果.医嘱id%Type,
  审查项目id_In In 处方审查结果.审查项目id%Type,
  药师审查_In   In 处方审查结果.药师审查%Type
) Is
  --功能：提交处方审查结果信息(药师审查信息)

  v_Err_Msg Varchar2(255);
  Err_Item Exception;
Begin

  --更新旧医嘱ID的“最后提交”
  Update 处方审查结果 Set 最后提交 = Null Where 最后提交 Is Not Null And 医嘱id = 医嘱id_In;

  --更新药师审查信息
  Update 处方审查结果
  Set 最后提交 = Decode(Nvl(医嘱id_In, 0), 0, Null, 1), 药师审查 = 药师审查_In
  Where 医嘱id = 医嘱id_In And 审查项目id = 审查项目id_In And 审方id = 审方id_In;
  If Sql%NotFound Then
    Insert Into 处方审查结果
      (审方id, 医嘱id, 审查项目id, 最后提交, 药师审查, 自动审查)
    Values
      (审方id_In, 医嘱id_In, 审查项目id_In, Decode(Nvl(医嘱id_In, 0), 0, Null, 1), 药师审查_In, Null);
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_处方审查_Audit_Detail;
/

CREATE OR REPLACE Function Zl_Fun_Pati_Calorie
(
  病人id_In In 病人信息.病人id%Type,
  主页id_In In 病人信息.主页id%Type,
  挂号id_In In 病人挂号记录.Id%Type
) Return Varchar2 Is

  --功能：通过病人信息，计算出病人的热量需要量
  v_Return  Varchar2(500);
  n_Sex     Number(1);
  n_Age     Number(5);
  n_Age_Var Number(10, 2);
  n_High    Number(5);
  n_Weight  Number(5);
  n_Calorie Number(10);
  n_Err     Number(1) := 1;
  v_Tmp     Varchar2(500);

  --获取年龄字符串的数值
  Function Get_Age(年龄_In In Varchar2) Return Number Is
    v_Tmp Varchar2(100) := '';
    N     Number(3) := 1;
  Begin
    Loop
      If N > Length(年龄_In) Then
        Exit;
      End If;
      If Regexp_Like(Substr(年龄_In, N, 1), '[0-9]') Then
        v_Tmp := v_Tmp || Substr(年龄_In, N, 1);
      Else
        Exit;
      End If;
      N := N + 1;
    End Loop;
  
    Return v_Tmp;
  End;

Begin

  If 主页id_In Is Null And 挂号id_In Is Null Then
    Return Null;
  End If;

  --性别
  Begin
    Select Decode(性别, '男', 1, '女', 2, Null), 年龄 Into n_Sex, v_Tmp From 病人信息 Where 病人id = 病人id_In;
  Exception
    When Others Then
      Select Null, Null Into n_Sex, v_Tmp From Dual;
  End;

  --年龄
  If v_Tmp Is Null Then
    Select 0, 0, 0 Into n_Age, n_Age_Var, n_Err From Dual;
  Else
    If v_Tmp Like '%岁%' Then
      n_Age     := Get_Age(v_Tmp);
      n_Age_Var := 1;
    Elsif v_Tmp Like '%月%' Then
      n_Age     := Get_Age(v_Tmp);
      n_Age_Var := Round(1 / 12, 2);
    Elsif v_Tmp Like '%天%' Or v_Tmp Like '%日%' Then
      n_Age     := Get_Age(v_Tmp);
      n_Age_Var := Round(1 / 365, 2);
    Elsif v_Tmp Like '%小时%' Or v_Tmp Like '%分%' Then
      n_Age     := 1;
      n_Age_Var := Round(1 / 365, 2);
    Else
      Select 0, 0, 0 Into n_Age, n_Age_Var, n_Err From Dual;
    End If;
  End If;

  If 主页id_In Is Not Null Then
    --住院
  
    --身高
    Begin
      Select 身高, 体重 Into n_High, n_Weight From 病案主页 Where 病人id = 病人id_In And 主页id = 主页id_In;
      If Nvl(n_High, 0) = 0 Or Nvl(n_Weight, 0) = 0 Then
        n_Err := 0;
      End If;
    Exception
      When Others Then
        Select 0, 0, 0 Into n_High, n_Weight, n_Err From Dual;
    End;
  
  Else
    --门诊
  
    --身高
    Begin
      Select b.记录内容
      Into n_High
      From 病人护理记录 A, 病人护理内容 B
      Where a.Id = b.记录id And a.病人id = 病人id_In And a.主页id = 挂号id_In And a.病人来源 = 1 And b.项目名称 = '身高';
    Exception
      When Others Then
        Select 0, 0 Into n_High, n_Err From Dual;
    End;
  
    --体重
    Begin
      Select b.记录内容
      Into n_Weight
      From 病人护理记录 A, 病人护理内容 B
      Where a.Id = b.记录id And a.病人id = 病人id_In And a.主页id = 挂号id_In And a.病人来源 = 1 And b.项目名称 = '体重';
    Exception
      When Others Then
        Select 0, 0 Into n_Weight, n_Err From Dual;
    End;
  
  End If;

  --计算需要量
  Select Nvl(n_High, 0), Nvl(n_Weight, 0) Into n_High, n_Weight From Dual;
  If n_Sex = 1 Then
    n_Calorie := 66.5 + 13.8 * n_Weight + 5.0 * n_High - 6.8 * n_Age * n_Age_Var;
    v_Return  := '66.5 + 13.8 * ' || n_Weight || 'KG + 5.0 * ' || n_High || 'CM - 6.8 * ' ||
                 Round(n_Age * n_Age_Var, 2) || '岁 = ' || n_Calorie * n_Err;
  Else
    n_Calorie := 655.1 + 9.6 * n_Weight + 1.8 * n_High - 4.7 * n_Age * n_Age_Var;
    v_Return  := '655.1 + 9.6 * ' || n_Weight || 'KG + 1.8 * ' || n_High || 'CM - 4.7 * ' ||
                 Round(n_Age * n_Age_Var, 2) || '岁 = ' || n_Calorie * n_Err;
  End If;

  Return v_Return;

End Zl_Fun_Pati_Calorie;
/

Create Or Replace Procedure Zl1_Datamove_Reb
(
  System_In    In Number,
  Speedmode_In In Number,
  Func_In      In Number,
  Enable_In    In Number := 0,
  Parallel_In  In Number := 0,
  Rebscope_In  In Number := 0
) As
  --功能：在历史数据转出之前，禁用触发器、自动作业、约束、索引，转出之后启用这些对象，以及重建待转出索引，收回标记转出所用索引的空间
  --参数：
  --System_In:    应用系统编号,100=标准版
  --speedmode_in：数据转出模式，0-在线模式，1-离线模式（在客户端停用时，转出期间禁用转出表的主键、唯一键、外键约束和索引，以加快已转数据的删除操作）
  --func_in:      1=触发器，2=自动作业，3=约束，4=索引，5=重建待转出索引，6-收回标记转出所用索引的空间，7-重组表的存储空间（move），并恢复被禁用的约束和索引
  --Enable_in:    0-禁用，1=启用，对func_in值为1-4有效
  --rebScope_in:   Func_In=6时，指重建索引的范围(0-经济核算类,1-经济核算类及医嘱类,2-全部)，Func_In=7时指Move表的范围(0-经济核算类，1-全部)

  v_Sql Varchar2(1000);
  n_Do  Number(1);
  v_Tbs Varchar2(100);

  --转出标记中的SQL查询所需的索引
  v_Indexeswithtag Varchar2(4000) := '门诊费用记录_IX_结帐ID,住院费用记录_IX_结帐ID,费用补充记录_IX_结算ID,费用补充记录_IX_登记时间,病人预交记录_IX_主页ID,病人预交记录_IX_结帐ID,病人预交记录_IX_收款时间,门诊费用记录_IX_登记时间,门诊费用记录_IX_医嘱序号,住院费用记录_IX_登记时间,病人结帐记录_IX_收费时间,病人结帐记录_IX_病人id' ||
                                     ',药品收发记录_IX_费用ID,收发记录补充信息_IX_收发ID,输液配药内容_IX_收发ID,药品留存计划_IX_留存ID,药品签名明细_IX_收发ID' ||
                                     ',人员借款记录_IX_借出时间,人员收缴记录_IX_登记时间,人员暂存记录_IX_收缴ID,人员暂存记录_IX_登记时间,票据领用记录_IX_登记时间,票据使用明细_IX_领用ID,票据打印明细_IX_使用ID' ||
                                     ',病人挂号记录_IX_登记时间,病人医嘱发送_IX_发送时间,病人医嘱记录_IX_挂号单,病人医嘱记录_IX_主页ID,病人医嘱记录_IX_相关ID' ||
                                     ',病案主页_IX_出院日期,住院费用记录_IX_病人ID,病人过敏记录_IX_病人ID,病人诊断记录_IX_病人ID,病人手麻记录_IX_主页ID' ||
                                     ',病人护理记录_IX_主页ID,病人护理内容_IX_记录id,病人护理文件_IX_主页ID,病人护理数据_IX_文件ID,病人护理明细_IX_记录ID,病人护理打印_IX_文件ID' ||
                                     ',电子病历记录_IX_病人ID,病人医嘱报告_IX_病历ID,影像报告驳回_IX_医嘱ID,报告查阅记录_IX_病历ID,病人诊断记录_IX_病历ID' ||
                                     ',病人临床路径_IX_病人ID,病人合并路径_IX_首要路径记录ID,病人路径执行_IX_路径记录ID,病人出径记录_IX_路径记录ID,病人诊断医嘱_IX_医嘱ID' ||
                                     ',影像申请单图像_IX_医嘱ID,影像收藏内容_IX_医嘱ID,检验标本记录_IX_医嘱ID,检验项目分布_IX_标本ID,检验分析记录_IX_标本ID' ||
                                     ',检验操作记录_IX_标本ID,检验图像结果_IX_标本ID,检验拒收记录_IX_医嘱ID,检验普通结果_IX_检验标本ID,处方审查明细_IX_医嘱ID';

  --转出标记中的SQL查询所需的索引(主键及唯一键对应的索引)
  v_Constraintswithtag Varchar2(4000) := '病人预交记录_UQ_NO,病人结帐记录_UQ_NO,病人结帐记录_PK,门诊费用记录_UQ_NO,住院费用记录_UQ_NO' ||
                                         ',病人卡结算对照_PK,费用补充记录_PK,病人卡结算记录_PK,三方结算交易_PK,输液配药记录_PK,药品签名记录_PK,票据打印内容_PK,病人挂号记录_PK,病人挂号汇总_UQ_日期,病人转诊记录_UQ_NO' ||
                                         ',病人护理活动项目_UQ_页号,病人护理要素内容_UQ_页号,产程要素内容_PK,电子病历记录_PK,电子病历附件_PK,电子病历格式_PK,电子病历内容_UQ_对象序号,电子病历图形_PK,疾病申报记录_PK' ||
                                         ',病人合并路径评估_PK,病人路径评估_PK,病人路径变异_PK,病人路径指标_UQ_评估指标,病人路径医嘱_PK' ||
                                         ',病人医嘱记录_PK,病人医嘱报告_PK,病人医嘱计价_UQ_收费细目ID,病人医嘱附费_PK,病人医嘱附件_PK,病人医嘱执行_PK,医嘱执行时间_PK,医嘱执行打印_PK,病人医嘱打印_UQ_医嘱ID,输血申请记录_PK,输血检验结果_PK' ||
                                         ',病人诊断记录_PK,病人医嘱状态_PK,医嘱签名记录_PK,病人医嘱发送_PK,诊疗单据打印_PK,医嘱执行计价_PK,执行打印记录_PK' ||
                                         ',影像检查记录_PK,影像检查序列_UQ_序列号,影像检查图象_UQ_图像号,影像危急值记录_UQ_医嘱ID' ||
                                         ',检验申请项目_PK,检验质控记录_PK,检验签名记录_PK,检验试剂记录_PK,检验质控报告_PK,检验药敏结果_PK,人员收缴记录_PK,人员收缴明细_PK,人员收缴票据_PK,人员收缴对照_PK' ||
                                         ',处方审查记录_PK,处方审查结果_UQ_审方ID';

  --功能：1.禁用或启用引用转出表主键的他表外键,避免删除主表记录时对子表每行记录执行一次SQL查询或删除
  --      2.禁用或启用主键或唯一键约束（禁用时会自动删除对应的索引，启用时自动创建），以提高数据删除性能
  --例如：病人医嘱发送_FK_医嘱ID，如果这些外键所在的表，数据未转出（未在zlbaktables表中定义），执行前会检查并限制转出。
  Procedure Setconstraintstatus As
  Begin
    --禁用时，先禁用引用转出表主键的他表外键，再禁用转出表的主键
    If Enable_In = 0 Then
      For R In (Select c.Table_Name, c.Constraint_Name, a.Table_Name As Ptable_Name
                From User_Constraints A, User_Constraints C, zlBakTables B
                Where a.Table_Name = b.表名 And b.直接转出 = 1 And b.系统 = System_In And a.Constraint_Type In ('P', 'U') And
                      c.r_Constraint_Name = a.Constraint_Name And c.Constraint_Type = 'R' And c.Status = 'ENABLED'
                Order By a.Table_Name) Loop
        v_Sql := 'Alter Table ' || r.Table_Name || ' Disable Constraint ' || r.Constraint_Name;
        Execute Immediate v_Sql;
      End Loop;
    
      If Speedmode_In = 1 Then
        --禁用主键或唯一键索引(必须删除索引，否则即使skip_unusable_indexes为true，也无法删除存在Unusable状态的唯一性索引的表中的记录)
        --保留转出标记中的SQL查询所需的索引(主键和唯一键对应的索引)
        For R In (Select a.Table_Name, a.Constraint_Name
                  From User_Constraints A, zlBakTables T, User_Tables B
                  Where a.Table_Name = t.表名 And t.直接转出 = 1 And t.系统 = System_In And a.Status = 'ENABLED' And
                        a.Constraint_Type In ('P', 'U') And a.Table_Name = b.Table_Name And b.Iot_Type Is Null And
                        a.Constraint_Name Not In
                        (Select Upper(Column_Value) As Constraint_Name From Table(f_Str2list(v_Constraintswithtag)))
                  Order By Constraint_Name) Loop
          v_Sql := 'Alter Table ' || r.Table_Name || ' Disable Constraint ' || r.Constraint_Name ||
                   ' Cascade Drop Index';
          Execute Immediate v_Sql;
        End Loop;
      End If;
    Else
      --启用时，先启用主键和唯一键，再启用引用转出表主键的他表外键
      If Speedmode_In = 1 Then
        --先重建索引，再启用约束，以便重建索引时利用并行执行缩短时间，并且启用约束时也可以采用novalidate方式
        For R In (Select d.Table_Name, d.Constraint_Name, LTrim(Max(Sys_Connect_By_Path(d.Column_Name, ',')), ',') Colstr
                  From User_Cons_Columns D,
                       (Select a.Table_Name, a.Constraint_Name
                         From User_Constraints A, zlBakTables T
                         Where a.Table_Name = t.表名 And t.直接转出 = 1 And t.系统 = System_In And a.Status = 'DISABLED' And
                               a.Constraint_Type In ('P', 'U')) A
                  Where a.Constraint_Name = d.Constraint_Name And a.Table_Name = d.Table_Name
                  Start With d.Position = 1
                  Connect By Prior d.Position + 1 = d.Position And Prior d.Constraint_Name = d.Constraint_Name
                  Group By d.Table_Name, d.Constraint_Name
                  Order By Constraint_Name) Loop
          Update Zldatamovelog
          Set 当前进度 = '正在恢复约束:' || r.Constraint_Name
          Where 系统 = System_In And 批次 = (Select Max(批次) From Zldatamovelog Where 系统 = System_In);
        
          Select Tablespace_Name Into v_Tbs From User_Indexes Where Table_Name = r.Table_Name And Rownum < 2;
        
          --禁用主键或唯一键时，索引是被删除了的，所以这里要用Create
          v_Sql := 'Create Unique Index ' || r.Constraint_Name || ' On ' || r.Table_Name || '(' || r.Colstr ||
                   ') Tablespace ' || v_Tbs || ' Nologging';
          Begin
            Execute Immediate v_Sql;
          Exception
            When Others Then
              Null; --可能有些主键或唯一键不是本次转出期间被禁用的，之前就存在不唯一数据，创建唯一索引会出错
          End;
        
          --会自动建立约束与索引的关联
          v_Sql := 'Alter Table ' || r.Table_Name || ' Enable Novalidate Constraint ' || r.Constraint_Name;
          Execute Immediate v_Sql;
        End Loop;
      End If;
    
      --启用引用转出表主键的他表外键
      For R In (Select c.Table_Name, c.Constraint_Name, a.Table_Name As Ptable_Name
                From User_Constraints A, User_Constraints C, zlBakTables B
                Where a.Table_Name = b.表名 And b.直接转出 = 1 And b.系统 = System_In And a.Constraint_Type In ('P', 'U') And
                      c.r_Constraint_Name = a.Constraint_Name And c.Constraint_Type = 'R' And c.Status = 'DISABLED'
                Order By a.Table_Name) Loop
        --为了加快速度，采用novalidate，不验证已有数据
        --可能引用转出表主键的他表，在zlbaktables中定义了，但没有编写对应的数据转出脚本，未验证的数据可能有违反约束的情况。
        v_Sql := 'Alter Table ' || r.Table_Name || ' Enable Novalidate Constraint ' || r.Constraint_Name;
        Execute Immediate v_Sql;
      End Loop;
    End If;
  End Setconstraintstatus;

  --功能：高速模式时禁用LOB以外的所有索引，在线模式时仅禁用转出表引用非转出表的外键索引(例如：病人医嘱计价_IX_收费细目ID)
  --说明：禁用索引是为了提高删除数据的性能
  Procedure Setindexstatus As
  Begin
    If Speedmode_In = 1 Then
      --保留转出标记中的SQL查询所需的索引
      For R In (Select /*+ rule*/
                 a.Index_Name
                From User_Indexes A, zlBakTables T
                Where a.Table_Name = t.表名 And t.直接转出 = 1 And t.系统 = System_In And t.直接转出 = 1 And
                      a.Index_Name <> a.Table_Name || '_IX_待转出' And
                      a.Index_Name Not In
                      (Select Upper(Column_Value) As Index_Name From Table(f_Str2list(v_Indexeswithtag))) And
                      a.Status = Decode(Enable_In, 0, 'VALID', 'UNUSABLE') And a.Index_Type = 'NORMAL' And Not Exists
                 (Select 1
                       From User_Constraints C
                       Where c.Index_Name = a.Index_Name And c.Constraint_Type In ('P', 'U'))
                Order By Index_Name) Loop
      
        If Enable_In = 0 Then
          v_Sql := 'Alter Index ' || r.Index_Name || ' Unusable';
          Execute Immediate v_Sql;
        Else
          Update Zldatamovelog
          Set 当前进度 = '正在重建索引:' || r.Index_Name
          Where 系统 = System_In And 批次 = (Select Max(批次) From Zldatamovelog Where 系统 = System_In);
        
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Nologging';
          Begin
            Execute Immediate v_Sql;
            --在线重建比较慢，不在线重建则需要锁表，如果有其他并发事务，则会出错：ORA-00054: 资源正忙, 但指定以 NOWAIT 方式获取资源


          Exception
            When Others Then
              If SQLErrM Like 'ORA-00054%' Then
                v_Sql := Replace(v_Sql, 'Rebuild', 'Rebuild Online');
                Execute Immediate v_Sql;
              End If;
          End;
        End If;
      End Loop;
    Else
      For R In (Select a.Index_Name
                From (Select d.Table_Name, d.Index_Name, LTrim(Max(Sys_Connect_By_Path(d.Column_Name, ',')), ',') Colstr
                       From User_Ind_Columns D, zlBakTables T, User_Indexes C
                       Where c.Table_Name = t.表名 And t.直接转出 = 1 And t.系统 = System_In And c.Uniqueness = 'NONUNIQUE' And
                             c.Index_Type = 'NORMAL' And c.Status = Decode(Enable_In, 0, 'VALID', 'UNUSABLE') And
                             c.Index_Name = d.Index_Name And c.Table_Name = d.Table_Name
                       Start With d.Column_Position = 1
                       Connect By Prior d.Column_Position + 1 = d.Column_Position And Prior d.Index_Name = d.Index_Name
                       Group By d.Table_Name, d.Index_Name) A,
                     (Select e.Table_Name, LTrim(Max(Sys_Connect_By_Path(e.Column_Name, ',')), ',') Colstr
                       From User_Cons_Columns E, User_Constraints F, zlBakTables T, User_Constraints C
                       Where e.Table_Name = t.表名 And t.直接转出 = 1 And t.系统 = System_In And
                             e.Constraint_Name = f.Constraint_Name And f.Constraint_Type = 'R' And
                             c.Constraint_Name = f.r_Constraint_Name And c.Table_Name Not In ('病案主页', '病人信息') And
                             Not Exists
                        (Select 1 From zlBakTables G Where g.表名 = c.Table_Name And g.系统 = System_In)
                       Start With Nvl(e.Position, 1) = 1
                       Connect By Prior Nvl(e.Position, 1) + 1 = Nvl(e.Position, 1) And
                                  Prior e.Constraint_Name = e.Constraint_Name
                       Group By e.Table_Name, e.Constraint_Name) B
                Where a.Table_Name = b.Table_Name And a.Colstr = b.Colstr
                Order By Index_Name) Loop
      
        If Enable_In = 0 Then
          --特殊处理：以下两个索引不禁用，是由于药品目录修改规格，财务缴款需要使用
          If r.Index_Name Not In ('病人医嘱记录_IX_收费细目ID', '人员收缴记录_IX_缴款组ID') Then
            v_Sql := 'Alter Index ' || r.Index_Name || ' Unusable';
            Execute Immediate v_Sql;
          End If;
        Else
          Update Zldatamovelog
          Set 当前进度 = '正在重建索引:' || r.Index_Name
          Where 系统 = System_In And 批次 = (Select Max(批次) From Zldatamovelog Where 系统 = System_In);
        
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Online Nologging';
          Execute Immediate v_Sql;
          --在线重建比较慢，不在线重建则需要锁表，如果有其他并发事务，则会出错：ORA-00054: 资源正忙, 但指定以 NOWAIT 方式获取资源  
        End If;
      End Loop;
    End If;
  End Setindexstatus;

  --功能：转出数据期间，停用转出表上的所有触发器，转出后再恢复
  Procedure Settriggerstatus As
  Begin
    For R In (Select Distinct a.Table_Name, t.停用触发器
              From User_Triggers A, zlBakTables T
              Where a.Status = Decode(Enable_In, 0, 'ENABLED', 'DISABLED') And a.Table_Name = t.表名 And t.直接转出 = 1 And
                    t.系统 = System_In) Loop
      If Enable_In = 0 Then
        v_Sql := 'Alter Table ' || r.Table_Name || ' DISABLE ALL TRIGGERS';
        Update zlBakTables Set 停用触发器 = 1 Where 系统 = System_In And 表名 = r.Table_Name;
      Elsif Nvl(r.停用触发器, 0) = 1 Then
        v_Sql := 'Alter Table ' || r.Table_Name || ' ENABLE ALL TRIGGERS';
        Update zlBakTables Set 停用触发器 = Null Where 系统 = System_In And 表名 = r.Table_Name;
      End If;
      Execute Immediate v_Sql;
    End Loop;
    Commit;
  End Settriggerstatus;

  --功能：转出数据期间，停用当前所有者的所有自动作业，转出后再启用
  Procedure Setjobstatus As
    v_Jobs Varchar2(4000);
  Begin
    --停用
    If Enable_In = 0 Then
      For R In (Select Job From User_Jobs Where Broken = 'N') Loop
        Dbms_Job.Broken(r.Job, True);
        v_Jobs := v_Jobs || ',' || r.Job;
      End Loop;
    
      If v_Jobs Is Not Null Then
        v_Jobs := Substr(v_Jobs, 2);
        Update zlDataMove Set 停用作业号 = v_Jobs Where 系统 = System_In And 组号 = 1;
      End If;
    Else
      --启用
      Select 停用作业号 Into v_Jobs From zlDataMove Where 系统 = System_In And 组号 = 1;
      If v_Jobs Is Not Null Then
        For R In (Select Job
                  From User_Jobs
                  Where Broken = 'Y' And Job In (Select Column_Value From Table(f_Num2list(v_Jobs)))) Loop
          Dbms_Job.Broken(r.Job, False);
        End Loop;
        Update zlDataMove Set 停用作业号 = Null Where 系统 = System_In And 组号 = 1;
      End If;
    End If;
    --作业设置后必须提交事务才生效
    Commit;
  End Setjobstatus;
Begin
  If Parallel_In < 2 Then
    Execute Immediate 'Alter Session DISABLE PARALLEL DDL';
  Else
    If Speedmode_In = 1 And (Func_In In (6, 7) Or Func_In In (3, 4) And Enable_In = 1) Then
      --为重建索引设置并行执行（由于通常受限于IO设备的性能，设置太高的并行度反而会降低性能，如有高性能存储设备，可加大并行度）
      --执行重建索引后会自动为索引加上并行度属性，如果不取消，会影响相关SQL的执行计划(全表扫描+并行查询，巨慢),在后面取消索引的并行度


      Execute Immediate 'Alter Session FORCE PARALLEL DDL PARALLEL ' || Parallel_In;
    End If;
  End If;

  If Func_In = 1 Then
    --1.设置触发器
    Settriggerstatus;
  Elsif Func_In = 2 Then
    --2.设置自动作业
    Setjobstatus;
  Elsif Func_In = 3 Then
    --3.设置约束状态
    Setconstraintstatus;
  Elsif Func_In = 4 Then
    --4.设置索引状态
    Setindexstatus;
  Elsif Func_In = 5 Then
    --5.重建"待转出"索引
    For R In (Select b.Index_Name
              From zlBakTables A, User_Indexes B
              Where a.表名 = b.Table_Name And a.直接转出 = 1 And a.系统 = System_In And b.Index_Name = b.Table_Name || '_IX_待转出'
              Union All
              Select '病案主页_IX_待转出' From Dual Where System_In = 100) Loop
      Update Zldatamovelog
      Set 当前进度 = '正在重建索引:' || r.Index_Name
      Where 系统 = System_In And 批次 = (Select Max(批次) From Zldatamovelog Where 系统 = System_In);
    
      --耗时太短，无须并行DDL
      --在线转出时如果重建索引会锁表，如果有其他并发事务，则会出错：ORA-00054: 资源正忙, 但指定以 NOWAIT 方式获取资源
      If Speedmode_In = 1 Then
        v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Nologging';
      Else
        v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Online Nologging';
      End If;
      Begin
        Execute Immediate v_Sql;
      Exception
        When Others Then
          If SQLErrM Like 'ORA-00054%' Then
            v_Sql := Replace(v_Sql, 'Rebuild', 'Rebuild Online');
            Execute Immediate v_Sql;
          End If;
      End;
    End Loop;
  
  Elsif Func_In = 6 Then
    --6.重建标记转出查询所用到的索引（测试表明重建后最多可缩短一半的查询时间）
    --根据业务的启用阶段来决定重建哪些索引，以避免一些不必要的重建耗时
    For R In (Select b.Index_Name, a.组号
              From User_Indexes B, zlBakTables A
              Where a.系统 = System_In And a.表名 = b.Table_Name And
                    b.Index_Name In
                    (Select Upper(Column_Value)
                     From Table(f_Str2list(v_Indexeswithtag))
                     Union
                     Select Upper(Column_Value) From Table(f_Str2list(v_Constraintswithtag)))
              Order By Index_Name) Loop
      n_Do := 0;
      If Rebscope_In = 0 Then
        If r.组号 < 5 Then
          n_Do := 1; --仅经济核算类
        End If;
      Elsif Rebscope_In = 1 Then
        If r.组号 < 5 Or r.组号 = 8 Then
          n_Do := 1; --仅经济核算类、医嘱类
        End If;
      Else
        n_Do := 1;
      End If;
    
      If n_Do = 1 Then
        Update Zldatamovelog
        Set 当前进度 = '正在重建索引:' || r.Index_Name
        Where 系统 = System_In And 批次 = (Select Max(批次) From Zldatamovelog Where 系统 = System_In);
      
        --v_Sql := 'Alter Index ' || r.Index_Name || ' shrink Space';
        --使用shrink方式不能并行执行,试验表明速度比rebuild PARALLEL 8 慢6倍
        If Speedmode_In = 1 Then
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Nologging';
        Else
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Online Nologging';
        End If;
        Begin
          Execute Immediate v_Sql;
          --在线重建比较慢，不在线重建则需要锁表，如果有其他并发事务，则会出错：ORA-00054: 资源正忙, 但指定以 NOWAIT 方式获取资源


        Exception
          When Others Then
            If SQLErrM Like 'ORA-00054%' Then
              v_Sql := Replace(v_Sql, 'Rebuild', 'Rebuild Online');
              Execute Immediate v_Sql;
            End If;
        End;
      End If;
    End Loop;
  
    --重组表的数据(在线转出时会影响业务的使用，所以不支持)
  Elsif Func_In = 7 And Speedmode_In = 1 Then
    --rebScope_in=0,只重组组号小于5的经济核算类表（费用、药品、票据），否则全部重组
    For R In (Select a.表名 As Table_Name
              From zlBakTables A
              Where a.直接转出 = 1 And (组号 < Decode(Rebscope_In, 0, 5, 100))
              Order By 组号, 序号) Loop
    
      Update Zldatamovelog
      Set 当前进度 = '正在重组表:' || r.Table_Name
      Where 系统 = System_In And 批次 = (Select Max(批次) From Zldatamovelog Where 系统 = System_In);
    
      --如果有空闲的空间，最好移到其他表空间，只有这样才能绝对移动文件尾部的数据块，以便进行表空间文件的收缩
      --在前面设置了会话级的强制并行
      v_Sql := 'Alter Table ' || r.Table_Name || ' Move Nologging';
      Execute Immediate v_Sql;
    
      --单独移动Lob对象
      For L In (Select Column_Name, Tablespace_Name From User_Lobs Where Table_Name = r.Table_Name) Loop
        v_Sql := 'Alter Table ' || r.Table_Name || ' Move Lob(' || l.Column_Name || ') Store as (Tablespace ' ||
                 l.Tablespace_Name || ') Nologging';
        Execute Immediate v_Sql;
      End Loop;
    
      v_Sql := 'Alter Table ' || r.Table_Name || ' Noparallel';
      Execute Immediate v_Sql;
    
      --move后，表相关的索引会全部失效，需要全部重建
      For S In (Select Index_Name
                From User_Indexes
                Where Table_Name = r.Table_Name And Status = 'UNUSABLE' And
                      (Index_Name = r.Table_Name || '_IX_待转出' Or
                      Index_Name In
                      (Select Upper(Column_Value)
                        From Table(f_Str2list(v_Indexeswithtag))
                        Union
                        Select Upper(Column_Value) From Table(f_Str2list(v_Constraintswithtag))))
                Order By Index_Name) Loop
        Update Zldatamovelog
        Set 当前进度 = '正在恢复失效索引:' || s.Index_Name
        Where 系统 = System_In And 批次 = (Select Max(批次) From Zldatamovelog Where 系统 = System_In);
      
        --在前面设置了会话级的强制并行
        v_Sql := 'Alter Index ' || s.Index_Name || ' Rebuild Nologging';
        Execute Immediate v_Sql;
      End Loop;
    End Loop;
  End If;

  --执行重建索引后会自动为索引加上并行度属性，如果不取消，会影响相关SQL的执行计划(全表扫描+并行查询，巨慢)
  ---------------------------------------------------------------------------------------------------
  If Speedmode_In = 1 And Parallel_In > 1 And (Func_In In (6, 7) Or Func_In In (3, 4) And Enable_In = 1) Then
    Execute Immediate 'ALTER Session DISABLE PARALLEL DDL';
  
    For R In (Select Index_Name From User_Indexes Where Degree Not In ('1', '0')) Loop
      v_Sql := 'Alter Index ' || r.Index_Name || ' Noparallel';
      Execute Immediate v_Sql;
    End Loop;
  End If;

  Update Zldatamovelog
  Set 当前进度 = '重建完成'
  Where 系统 = System_In And 批次 = (Select Max(批次) From Zldatamovelog Where 系统 = System_In);
  Commit;
  --本过程不进行错误处理，错误由调用过程处理
End Zl1_Datamove_Reb;
/

Create Or Replace Procedure Zl1_Datamove_Tag
(
  d_End    In Date,
  n_批次   In Number,
  n_System In Number
) As
  --功能：标记待转出的数据
  --说明：为避免Undo表空间膨胀过大，分段提交
Begin
  --1.经济核算（费用,药品,收款和票据等）

  --*****特殊处理遵义医院特殊数据:
  --病人ID为1的"医保病人",为2的"旧病人":不检查是否结清，不限制预交款未冲完的，强制转出
  Update /*+ rule*/ 病人预交记录 L
  Set 待转出 = n_批次
  Where 结帐id In
        (Select Distinct a.结帐id --1.门诊收费和挂号的收费结算记录(排除之后退号和退费的,一张单据中只要其中一行退了)
         From 门诊费用记录 A
         Where a.待转出 Is Null And a.登记时间 < d_End And a.记录性质 In (1, 4) And
               (a.记录状态 In (1, 2) Or a.记录状态 = 3 And Not Exists
                (Select 1
                 From 门诊费用记录 B
                 Where a.No = b.No And a.记录性质 = b.记录性质 And b.记录状态 = 2 And b.登记时间 >= d_End))
         Union All
         Select Distinct a.结算id --2.医保补结算
         From 费用补充记录 A
         Where a.待转出 Is Null And a.登记时间 < d_End And a.记录性质 = 1 And
               (a.记录状态 In (1, 2) Or a.记录状态 = 3 And Not Exists
                (Select 1
                 From 费用补充记录 B
                 Where a.No = b.No And a.记录性质 = b.记录性质 And b.记录状态 In (1, 2) And b.登记时间 >= d_End))
         Union All
         Select Distinct a.结帐id --3.就诊卡的收费结算记录(排除之后退卡费的,一张单据中只要其中一行退了)
         From 住院费用记录 A
         Where a.待转出 Is Null And a.登记时间 < d_End And a.记录性质 = 5 And a.记帐费用 = 0 And
               (a.记录状态 In (1, 2) Or a.记录状态 = 3 And Not Exists
                (Select 1
                 From 住院费用记录 B
                 Where a.No = b.No And a.记录性质 = b.记录性质 And b.记录状态 = 2 And b.登记时间 >= d_End))
         Union All --4.门诊(记帐单)和住院的结帐结算记录
         Select 结帐id
         From (With Settle As (Select Distinct a.Id As 结帐id, a.病人id --3.门诊(记帐单)和住院的结帐结算记录(排除之后结帐作废的)
                               From 病人结帐记录 A
                               Where a.待转出 Is Null And a.收费时间 < d_End And
                                     (a.记录状态 In (1, 2) Or a.记录状态 = 3 And Not Exists
                                      (Select 1
                                       From 病人结帐记录 B
                                       Where a.No = b.No And b.记录状态 = 2 And b.收费时间 >= d_End)))
                Select 结帐id
                From Settle
                Minus
                --1.一张预交款被多笔结帐冲完（结帐ID不同），这些结帐ID要整体排除,避免部分被转出后影响后续的计算是否冲完 
                --2.这些费用单据的结帐ID对应的可能还有其他NO的其他结帐ID(结帐作废后分多次结帐结清，可能部分在转出时间之后)，这些结帐ID要整体排除,避免部分被转出后影响后续的计算是否结清
                --考虑到这情况的复杂性，为简化逻辑，提升查询性能，按病人ID来排除
                Select Distinct d.Id
                From 病人结帐记录 D,
                     (Select Distinct c.病人id --多次住院可以一起结，以及门诊记帐和住院记帐可以一起结且冲同一笔预交，所以这里不加主页ID
                       From 住院费用记录 C,
                            (Select Distinct d.No, d.序号, Mod(d.记录性质, 10) As 记录性质
                              From 住院费用记录 D,
                                   (Select s.结帐id
                                     From Settle S, 病人结帐记录 E
                                     Where s.病人id = e.病人id And
                                           (e.收费时间 > d_End Or Exists (Select 1 From 在院病人 F Where s.病人id = f.病人id))) S --没有结清且之后没有再结过就成了呆帐，这种就不排除
                              Where d.结帐id = s.结帐id) D
                       Where c.No = d.No And Mod(c.记录性质, 10) = d.记录性质 And c.序号 = d.序号 --结帐后作废后，再对包含记帐单销帐的结帐ID为空的记录,一起汇总计算是否结清,这种结帐ID为空的数据转出在后面单独转出                                        
                       Group By c.No, Mod(c.记录性质, 10), c.病人id --一张单据中的一行可部分结帐，以单据为对象来判断，避免一张单据的其中一部分被转出
                       Having Nvl(Sum(c.实收金额), 0) <> Nvl(Sum(c.结帐金额), 0) Or Exists (Select 1 --排除转出时间之后再次结帐的(作废后再次结帐)，避免原始单据转走后，后续结帐时无法正确判断
                                                                                   From 住院费用记录 E, 病人结帐记录 S
                                                                                   Where e.No = c.No And
                                                                                         Mod(e.记录性质, 10) = Mod(c.记录性质, 10) And
                                                                                         e.记录性质 In (12, 13, 15) And
                                                                                         e.结帐id = s.Id And s.收费时间 >= d_End)
                       Union All
                       Select Distinct c.病人id
                       From 门诊费用记录 C,
                            (Select Distinct d.No, d.序号, Mod(d.记录性质, 10) As 记录性质
                              From 门诊费用记录 D, Settle S
                              Where d.结帐id = s.结帐id) D --因为是门诊病人，所以，只要没有结清,该病人的都不转出
                       Where c.No = d.No And Mod(c.记录性质, 10) = d.记录性质 And c.序号 = d.序号
                       Group By c.No, Mod(c.记录性质, 10), c.病人id
                       Having Nvl(Sum(c.实收金额), 0) <> Nvl(Sum(c.结帐金额), 0) Or Exists (Select 1
                                                                                    From 门诊费用记录 E, 病人结帐记录 S
                                                                                    Where e.No = c.No And
                                                                                          Mod(e.记录性质, 10) = Mod(c.记录性质, 10) And
                                                                                          e.记录性质 In (12, 13, 15) And
                                                                                          e.结帐id = s.Id And s.收费时间 >= d_End)) N
                Where d.病人id = n.病人id)
                
         
         );

  --排除预交款未冲完的和转出时间之后发药的记录
  --因为前面的SQL查出的结帐ID可能不全是冲预交的(门诊收费和住院结帐补费等)，所以，需要单独一个SQL来排除
  --由于可能存在数据异常(住院费用结帐冲预交类别为1的门诊预交)，所以没有加预交类别条件限定
  Update /*+ rule*/ 病人预交记录
  Set 待转出 = Null
  Where 待转出 = n_批次 And
        结帐id In
        (Select Distinct d.结帐id
         From 病人预交记录 D,
              --连接D表是为了查冲同一预交单据的其他结帐ID（退预交款，冲预交作废的，再次冲同一预交单据）
              --该病人的所有结帐ID的都不转出，避免部分冲预交的结帐ID被排除后，原始预交单被转走，或者其他结帐ID将费用单据的一部分(原始结帐、结帐作废、再次结一部分、再次结全部)转走
              (Select Distinct l.病人id
                From 病人预交记录 L, 病人预交记录 P --可能本次结帐冲的只是剩余款，所以需要连接L表，查原始交预交的单据，以及记录性质为11的可能还有转出时间之后其他冲剩余款的结帐ID
                Where l.记录性质 In (1, 11) And l.No = p.No And p.记录性质 In (1, 11) And p.待转出 = n_批次
                Group By l.No, l.病人id
                Having Nvl(Sum(l.金额), 0) <> Nvl(Sum(l.冲预交), 0) And (Exists (Select 1
                                                                           From 病人预交记录 E
                                                                           Where l.病人id = e.病人id And e.收款时间 > d_End) Or Exists (Select 1
                                                                                                                               From 在院病人 E
                                                                                                                               Where l.病人id =
                                                                                                                                     e.病人id)) --没有冲完且之后没有再冲过或结算过就成了呆帐（存在用负的结帐补款来表示冲预交当成冲完的清况），这种就不排除 
                Or Nvl(Sum(l.金额), 0) = Nvl(Sum(l.冲预交), 0) And (Exists (Select 1
                                                                      From 病人预交记录 E, 病人结帐记录 F --排除转出时间之后的其他结帐ID冲的
                                                                      Where e.No = l.No And e.记录性质 = 11 And e.结帐id = f.Id And --冲预交时的收款时间填的是交预交款的时间，所以这里需要用其他表的时间
                                                                            f.收费时间 >= d_End) Or Exists (Select 1
                                                                                                       From 病人预交记录 E,
                                                                                                            门诊费用记录 F
                                                                                                       Where e.No = l.No And
                                                                                                             e.记录性质 = 11 And
                                                                                                             e.结帐id =
                                                                                                             f.结帐id And
                                                                                                             f.登记时间 >=
                                                                                                             d_End And
                                                                                                             f.记录性质 In
                                                                                                             (1, 4) And
                                                                                                             Nvl(f.记帐费用, 0) <> 1) Or Exists (Select 1
                                                                                                                                            From 病人预交记录 E,
                                                                                                                                                 住院费用记录 F
                                                                                                                                            Where e.No = l.No And
                                                                                                                                                  e.记录性质 = 11 And
                                                                                                                                                  e.结帐id =
                                                                                                                                                  f.结帐id And
                                                                                                                                                  f.登记时间 >=
                                                                                                                                                  d_End And
                                                                                                                                                  f.记录性质 In (5,
                                                                                                                                                             15) And
                                                                                                                                                  Nvl(f.记帐费用,
                                                                                                                                                      0) <> 1))) N
         
         Where d.病人id = n.病人id);

  --为了降低逻辑的复杂性，不排除在转出时间之后发药或未发药的费用记录对应的结帐ID，将这种情况的结算数据和费用数据强制转走 

  --预交款没有使用就直接退了的记录(结帐ID为空)
  Update /*+ rule*/ 病人预交记录
  Set 待转出 = n_批次
  Where 记录性质 = 1 And
        NO In (Select a.No
               From 病人预交记录 A
               Where a.结帐id Is Null And a.记录性质 = 1 And a.记录状态 In (2, 3) And a.待转出 Is Null And a.收款时间 < d_End
               Group By a.No
               Having Sum(a.金额) = 0);

  --冲预交款作废的记录（记录性质为2），没有结帐ID
  Update /*+ rule*/ 病人预交记录
  Set 待转出 = n_批次
  Where 结帐id Is Null And 记录性质 = 2 And NO In (Select a.No From 病人预交记录 A Where a.待转出 = n_批次 And a.记录性质 = 3);

  Update Zldatamovelog
  Set 当前进度 = '(1/10)结算数据标记完成，正在标记费用数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  Update /*+ rule*/ 病人结帐记录
  Set 待转出 = n_批次
  Where ID In (Select 结帐id From 病人预交记录 Where 待转出 = n_批次);

  --结帐无结算的记录(为了提升性能，不判断费用，只要结了帐且无预交记录就当成是零费用结帐)
  Update /*+ rule*/ 病人结帐记录 L
  Set 待转出 = n_批次
  Where 收费时间 < d_End And 待转出 Is Null And Not Exists (Select 1 From 病人预交记录 P Where l.Id = p.结帐id);

  Update /*+ rule*/ 病人卡结算对照
  Set 待转出 = n_批次
  Where 预交id In (Select a.Id From 病人预交记录 A Where 待转出 = n_批次);

  Update /*+ rule*/ 病人卡结算记录
  Set 待转出 = n_批次
  Where ID In (Select 卡结算id From 病人卡结算对照 Where 待转出 = n_批次);

  Update /*+ rule*/ 三方结算交易
  Set 待转出 = n_批次
  Where 交易id In (Select a.Id From 病人预交记录 A Where 待转出 = n_批次);

  --挂号打折后实收金额为0的(没有对应的预交记录),即使之后有退号费用也不管，因为金额为零不影响计算),而卡费即使为零也有预交记录                 
  --根据挂号记录再找门诊费用，比直接按时间查门诊费用要快
  Update /*+ rule*/ 门诊费用记录
  Set 待转出 = n_批次
  Where NO In (Select NO From 病人挂号记录 Where 待转出 Is Null And 登记时间 < d_End) And 记录性质 = 4 And 实收金额 = 0;

  --没有结帐的已冲销的记帐单或打折后实收金额为零的，且没有其他记帐单的强制转出
  Update /*+ rule*/ 门诊费用记录
  Set 待转出 = n_批次
  Where 记录性质 = 2 And
        NO In
        (Select NO
         From (Select b.挂号单, a.No, a.序号, Sum(a.实收金额)
                From 门诊费用记录 A, 病人医嘱记录 B
                Where a.医嘱序号 = b.Id And a.结帐id Is Null And a.记录性质 = 2 And b.病人来源 <> 4 And a.待转出 Is Null And a.登记时间 < d_End
                Group By a.No, a.序号, b.挂号单
                Having Sum(a.实收金额) = 0 And Not Exists (Select 1
                                                      From 门诊费用记录 C, 病人医嘱记录 D
                                                      Where b.挂号单 = d.挂号单 And d.Id = c.医嘱序号 And d.病人来源 <> 4 And c.记录性质 = 2 And
                                                            c.待转出 Is Null
                                                      Group By c.No, c.序号
                                                      Having Sum(a.实收金额) > 0)));

  --直接收费的和结帐无结算（预交）记录的，Union不加all去掉重复以减少in的数量
  Update /*+ rule*/ 门诊费用记录
  Set 待转出 = n_批次
  Where 结帐id In (Select 结帐id
                 From 病人预交记录
                 Where 待转出 = n_批次
                 Union
                 Select ID From 病人结帐记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 费用补充记录
  Set 待转出 = n_批次
  Where 结算id In (Select 结帐id From 病人预交记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 凭条打印记录
  Set 待转出 = n_批次
  Where (NO, 记录性质) In (Select NO, 记录性质 From 门诊费用记录 Where 待转出 = n_批次);

  --从预交记录读是为了取就诊卡直接收费的（无结帐ID）,再加结帐记录是为了取结帐无结算（预交）记录的
  Update /*+ rule*/ 住院费用记录
  Set 待转出 = n_批次
  Where 结帐id In (Select 结帐id
                 From 病人预交记录
                 Where 待转出 = n_批次
                 Union
                 Select ID From 病人结帐记录 Where 待转出 = n_批次);

  --1.转出结帐作废后，记帐单销帐的记录（记帐状态为2且没有结帐ID且(记录状态为3的有结帐ID的)在最前面已转出）
  --2.未结帐的零费用(已冲销的记帐单)
  --3.没有结帐ID的划价记录处理为转出
  --4.不收费也没有冲预交的零费用处理为转出
  --加条件"待转出 Is Null"是为了处理连续多次标记转出的情况
  Update /*+ rule*/ 门诊费用记录 A
  Set 待转出 = n_批次
  Where ((Exists (Select 1
                  From 门诊费用记录 B
                  Where a.No = b.No And a.记录性质 = b.记录性质 And a.序号 = b.序号 And b.记录状态 = 3 And b.结帐id Is Not Null And
                        b.待转出 + 0 = n_批次) And 记录状态 = 2 Or Exists
         (Select 1
           From 门诊费用记录 B
           Where a.No = b.No And a.记录性质 = b.记录性质 And a.序号 = b.序号
           Group By b.No, b.记录性质, b.序号
           Having Nvl(Sum(b.实收金额), 0) = 0)) And 记录性质 = 2 Or 记录状态 = 0 Or 记录性质 = 1 And 实收金额 = 0 And 结帐金额 = 0) And
        结帐id Is Null And 待转出 Is Null And 登记时间 < d_End;

  --1.转出结帐作废后，记帐单销帐的记录（记帐状态为2且没有结帐ID且(记录状态为3的有结帐ID的)在最前面已转出）
  --2.未结帐的零费用(已冲销的记帐单)
  --3.没有结帐ID的划价记录处理为转出
  Update /*+ rule*/ 住院费用记录 A
  Set 待转出 = n_批次
  Where ((Exists (Select 1
                  From 住院费用记录 B
                  Where a.No = b.No And a.记录性质 = b.记录性质 And a.序号 = b.序号 And b.记录状态 = 3 And b.结帐id Is Not Null And
                        b.待转出 + 0 = n_批次) And 记录状态 = 2 Or Exists
         (Select 1
           From 住院费用记录 B
           Where a.No = b.No And a.记录性质 = b.记录性质 And a.序号 = b.序号
           Group By b.No, b.记录性质, b.序号
           Having Nvl(Sum(b.实收金额), 0) = 0)) And 记录性质 = 2 Or 记录状态 = 0) And 结帐id Is Null And 待转出 Is Null And 登记时间 < d_End;

  --由于存在赖帐病人离院未结的情况，对于很久以前的这些数据，如果预交已冲完，则处理为要转出
  Update /*+ rule*/ 住院费用记录 A
  Set 待转出 = n_批次
  Where 待转出 Is Null And 结帐id Is Null And
        (病人id, 主页id) In (Select 病人id, 主页id
                         From 病案主页 C
                         Where 出院日期 < d_End And 待转出 Is Null And 数据转出 Is Null And Not Exists
                          (Select 1
                                From 病人预交记录 B
                                Where b.病人id = c.病人id And b.预交类别 = 2 And b.记录性质 In (1, 11) Having
                                 Nvl(Sum(b.金额), 0) - Nvl(Sum(b.冲预交), 0) <> 0));

  Update Zldatamovelog
  Set 当前进度 = '(2/10)费用数据标记完成，正在标记药品数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  Update /*+ Rule*/ 药品收发记录 A
  Set 待转出 = n_批次
  Where Rowid In (Select m.Rowid
                  From 药品收发记录 M, 门诊费用记录 E
                  Where m.费用id = e.Id And (e.记录性质 = 1 And m.单据 In (8, 24) Or e.记录性质 = 2 And m.单据 In (9, 25)) And
                        e.收费类别 In ('4', '5', '6', '7') And e.待转出 = n_批次
                  Union All
                  Select m.Rowid
                  From 药品收发记录 M, 住院费用记录 E
                  Where m.费用id = e.Id And m.单据 In (9, 10, 25, 26) And e.记录性质 = 2 And e.收费类别 In ('4', '5', '6', '7') And
                        e.待转出 = n_批次);

  Update /*+ rule*/ 收发记录补充信息
  Set 待转出 = n_批次
  Where 收发id In (Select ID From 药品收发记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 输液配药内容
  Set 待转出 = n_批次
  Where 收发id In (Select ID From 药品收发记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 输液配药记录
  Set 待转出 = n_批次
  Where ID In (Select 记录id From 输液配药内容 Where 待转出 = n_批次);

  Update /*+ rule*/ 输液配药附费
  Set 待转出 = n_批次
  Where 配药id In (Select ID From 输液配药记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 输液配药状态
  Set 待转出 = n_批次
  Where 配药id In (Select ID From 输液配药记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 药品留存计划
  Set 待转出 = n_批次
  Where 留存id In (Select ID From 药品收发记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 药品签名明细
  Set 待转出 = n_批次
  Where 收发id In (Select ID From 药品收发记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 药品签名记录
  Set 待转出 = n_批次
  Where ID In (Select 签名id From 药品签名明细 Where 待转出 = n_批次);

  Update Zldatamovelog
  Set 当前进度 = '(3/10)药品数据标记完成，正在标记缴款与票据数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  Update /*+ rule*/ 人员借款记录 Set 待转出 = n_批次 Where 待转出 Is Null And 借出时间 < d_End;

  Update /*+ rule*/ 人员收缴记录 Set 待转出 = n_批次 Where 待转出 Is Null And 登记时间 < d_End;

  Update /*+ rule*/ 人员收缴对照
  Set 待转出 = n_批次
  Where 收缴id In (Select ID From 人员收缴记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 人员收缴明细
  Set 待转出 = n_批次
  Where 收缴id In (Select ID From 人员收缴记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 人员收缴票据
  Set 待转出 = n_批次
  Where 收缴id In (Select ID From 人员收缴记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 人员暂存记录
  Set 待转出 = n_批次
  Where 收缴id In (Select ID From 人员收缴记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 人员暂存记录 Set 待转出 = n_批次 Where 待转出 Is Null And 记录性质 = 1 And 登记时间 < d_End;

  Update /*+ rule*/ 票据领用记录 A
  Set 待转出 = n_批次
  Where Not Exists
   (Select 1 From 票据使用明细 B Where b.领用id = a.Id And b.使用时间 >= d_End) And 待转出 Is Null And 剩余数量 = 0 And 登记时间 < d_End;

  Update /*+ rule*/ 票据使用明细
  Set 待转出 = n_批次
  Where 领用id In (Select ID From 票据领用记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 票据打印内容
  Set 待转出 = n_批次
  Where ID In (Select 打印id From 票据使用明细 Where 待转出 = n_批次);

  Update /*+ rule*/ 票据打印明细
  Set 待转出 = n_批次
  Where 使用id In (Select ID From 票据使用明细 Where 待转出 = n_批次);

  Update Zldatamovelog
  Set 当前进度 = '(4/10)缴款与票据数据标记完成，正在标记就诊及诊治数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  --2.就诊及诊治数据
  --不转出的条件：挂号费用未转出的，转出时间之后存在医嘱，医嘱对应的费用未转出的
  --即使正在就诊(r.执行状态 <> 2 )的也强制转出
  Update /*+ rule*/ 病人挂号记录 T
  Set 待转出 = n_批次
  Where Rowid In
        (Select Rowid
         From 病人挂号记录 R
         Where Not Exists (Select 1
                From 门诊费用记录 A
                Where r.No = a.No And a.登记时间 < d_End And a.记录性质 = 4 And a.待转出 Is Null) And Not Exists
          (Select 1
                From 病人医嘱记录 A
                Where a.挂号单 = r.No And a.病人来源 <> 4 And Nvl(a.停嘱时间, a.开嘱时间) >= d_End) And Not Exists
          (Select 1
                From 门诊费用记录 E, 病人医嘱记录 A
                Where r.No = a.挂号单 And a.Id = e.医嘱序号 And a.病人来源 <> 4 And e.待转出 Is Null) And r.待转出 Is Null And
               r.登记时间 < d_End);

  --由于有一部分挂号数据未转出，所以，汇总表的数据可能与挂号数据不匹配
  Update 病人挂号汇总 Set 待转出 = n_批次 Where 待转出 Is Null And 日期 < d_End;
  Update /*+ rule*/ 病人转诊记录 Set 待转出 = n_批次 Where NO In (Select NO From 病人挂号记录 Where 待转出 = n_批次);

  --通过"住院费用记录"来查询，而不是"病人结帐记录",因为离院未结的赖帐病人也转出了费用
  --出院日期条件仍然需要，因为可能某次结帐转出了，但病人当时并未出院(一次住院多次结帐)。
  --通过指定索引方式进行特殊优化（缺省采用"病案主页IX_出院日期"索引的效率太低）
  Update /*+ rule*/ 病案主页 P
  Set 待转出 = n_批次
  Where Not Exists (Select 1 From 住院费用记录 A Where a.病人id = p.病人id And a.主页id = p.主页id And a.待转出 Is Null) And 待转出 Is Null And
        数据转出 Is Null And 出院日期 < d_End And
        (病人id, 主页id) In (Select Distinct 病人id, 主页id From 住院费用记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人过敏记录
  Set 待转出 = n_批次
  Where (病人id, 主页id) In (Select 病人id, ID
                         From 病人挂号记录
                         Where 待转出 = n_批次
                         Union All
                         Select 病人id, 主页id From 病案主页 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人诊断记录
  Set 待转出 = n_批次
  Where (病人id, 主页id) In (Select 病人id, ID
                         From 病人挂号记录
                         Where 待转出 = n_批次
                         Union All
                         Select 病人id, 主页id From 病案主页 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人手麻记录
  Set 待转出 = n_批次
  Where (病人id, 主页id) In (Select 病人id, ID
                         From 病人挂号记录
                         Where 待转出 = n_批次
                         Union All
                         Select 病人id, 主页id From 病案主页 Where 待转出 = n_批次);

  Update Zldatamovelog
  Set 当前进度 = '(5/10)就诊及诊治数据标记完成，正在标记护理数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  --3.护理数据
  Update /*+ rule*/ 病人护理文件
  Set 待转出 = n_批次
  Where (病人id, 主页id) In (Select 病人id, 主页id From 病案主页 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人护理数据
  Set 待转出 = n_批次
  Where 文件id In (Select ID From 病人护理文件 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人护理明细
  Set 待转出 = n_批次
  Where 记录id In (Select ID From 病人护理数据 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人护理打印
  Set 待转出 = n_批次
  Where 文件id In (Select ID From 病人护理文件 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人护理活动项目
  Set 待转出 = n_批次
  Where 文件id In (Select ID From 病人护理文件 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人护理要素内容
  Set 待转出 = n_批次
  Where 文件id In (Select ID From 病人护理文件 Where 待转出 = n_批次);
  Update /*+ rule*/ 产程要素内容
  Set 待转出 = n_批次
  Where 文件id In (Select ID From 病人护理文件 Where 待转出 = n_批次);

  --老版护理系统数据
  Update /*+ rule*/ 病人护理记录
  Set 待转出 = n_批次
  Where (病人id, 主页id) In (Select 病人id, 主页id From 病案主页 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人护理内容
  Set 待转出 = n_批次
  Where 记录id In (Select ID From 病人护理记录 Where 待转出 = n_批次);

  Update Zldatamovelog
  Set 当前进度 = '(6/10)护理数据标记完成，正在标记病历数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  --4.病历数据
  Update /*+ rule*/ 电子病历记录
  Set 待转出 = n_批次
  Where 病人来源 <> 4 And (病人id, 主页id) In (Select 病人id, ID
                                       From 病人挂号记录
                                       Where 待转出 = n_批次
                                       Union All
                                       Select 病人id, 主页id From 病案主页 Where 待转出 = n_批次);

  --自登记类病人(无挂号单号)
  --病历ID可能重复是因为检验报告之类的，如肝功、肾功共打一张报告，即在病人医嘱报告表中，多个医嘱id对应同一报告ID
  --不管医嘱发送记录的执行状态，因为可能没有启用执行登记
  Update /*+ rule*/ 电子病历记录
  Set 待转出 = n_批次
  Where ID In (Select c.病历id
               From 病人医嘱发送 A, 病人医嘱记录 B, 病人医嘱报告 C
               Where c.医嘱id = b.Id And b.Id = a.医嘱id And b.相关id Is Null And Nvl(b.主页id, 0) = 0 And b.挂号单 Is Null And
                     a.发送数次 = 1 And a.待转出 Is Null And a.发送时间 < d_End);

  Update /*+ rule*/ 电子病历附件
  Set 待转出 = n_批次
  Where 病历id In (Select ID From 电子病历记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 电子病历格式
  Set 待转出 = n_批次
  Where 文件id In (Select ID From 电子病历记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 电子病历内容
  Set 待转出 = n_批次
  Where 文件id In (Select ID From 电子病历记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 电子病历图形
  Set 待转出 = n_批次
  Where 对象id In (Select ID From 电子病历内容 Where 待转出 = n_批次 And 对象类型 = 5);

  Update /*+ rule*/ 病人医嘱报告
  Set 待转出 = n_批次
  Where 病历id In (Select ID From 电子病历记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 影像报告驳回
  Set 待转出 = n_批次
  Where (医嘱id, 病历id) In (Select 医嘱id, 病历id From 病人医嘱报告 Where 待转出 = n_批次);

  Update /*+ rule*/ 报告查阅记录
  Set 待转出 = n_批次
  Where 病历id In (Select ID From 电子病历记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 疾病申报记录
  Set 待转出 = n_批次
  Where 文件id In (Select ID From 电子病历记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人诊断记录
  Set 待转出 = n_批次
  Where 病历id In (Select ID From 电子病历记录 Where 待转出 = n_批次);

  Update Zldatamovelog
  Set 当前进度 = '(7/10)病历数据标记完成，正在标记临床路径数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  --5.临床路径    
  Update /*+ rule*/ 病人临床路径
  Set 待转出 = n_批次
  Where (病人id, 主页id) In (Select 病人id, 主页id From 病案主页 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人合并路径
  Set 待转出 = n_批次
  Where 首要路径记录id In (Select ID From 病人临床路径 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人合并路径评估
  Set 待转出 = n_批次
  Where 路径记录id In (Select ID From 病人临床路径 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人出径记录
  Set 待转出 = n_批次
  Where 路径记录id In (Select ID From 病人临床路径 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人路径执行
  Set 待转出 = n_批次
  Where 路径记录id In (Select ID From 病人临床路径 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人路径评估
  Set 待转出 = n_批次
  Where 路径记录id In (Select ID From 病人临床路径 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人路径变异
  Set 待转出 = n_批次
  Where 路径记录id In (Select ID From 病人临床路径 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人路径指标
  Set 待转出 = n_批次
  Where 路径记录id In (Select ID From 病人临床路径 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人路径医嘱
  Set 待转出 = n_批次
  Where 路径执行id In (Select ID From 病人路径执行 Where 待转出 = n_批次);

  Update Zldatamovelog
  Set 当前进度 = '(8/10)临床路径数据标记完成，正在标记医嘱数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  --6.医嘱，检验，检查
  Update /*+ rule*/ 病人医嘱记录
  Set 待转出 = n_批次
  Where 挂号单 In (Select NO From 病人挂号记录 Where 待转出 = n_批次) And 病人来源 <> 4;
  Update /*+ rule*/ 病人医嘱记录
  Set 待转出 = n_批次
  Where (病人id, 主页id) In (Select 病人id, 主页id From 病案主页 Where 待转出 = n_批次);

  --自登记类病人(无挂号单)，病人医嘱报告在前面转病历时已转出
  Update /*+ rule*/ 病人医嘱记录
  Set 待转出 = n_批次
  Where Rowid In (Select b.Rowid
                  From 病人医嘱记录 B, 病人医嘱报告 C
                  Where (b.相关id = c.医嘱id Or b.Id = c.医嘱id) And c.待转出 = n_批次);

  Update /*+ rule*/ 病人医嘱计价
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人医嘱附费
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人医嘱附件
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 输血申请记录
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 输血检验结果
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人医嘱执行
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人医嘱打印
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 医嘱执行打印
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人诊断医嘱
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人诊断记录
  Set 待转出 = n_批次
  Where ID In (Select 诊断id From 病人诊断医嘱 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人医嘱状态
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 医嘱签名记录
  Set 待转出 = n_批次
  Where ID In (Select 签名id From 病人医嘱状态 Where 待转出 = n_批次 And 签名id Is Not Null);

  Update /*+ rule*/ 病人医嘱发送
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 诊疗单据打印
  Set 待转出 = n_批次
  Where (NO, 记录性质) In (Select NO, 记录性质 From 病人医嘱发送 Where 待转出 = n_批次);

  Update /*+ rule*/ 医嘱执行时间
  Set 待转出 = n_批次
  Where (医嘱id, 发送号) In (Select 医嘱id, 发送号 From 病人医嘱发送 Where 待转出 = n_批次);

  Update /*+ rule*/ 医嘱执行计价
  Set 待转出 = n_批次
  Where (医嘱id, 发送号) In (Select 医嘱id, 发送号 From 病人医嘱发送 Where 待转出 = n_批次);

  Update /*+ rule*/ 执行打印记录
  Set 待转出 = n_批次
  Where (医嘱id, 发送号) In (Select 医嘱id, 发送号 From 病人医嘱发送 Where 待转出 = n_批次);

  Update /*+ rule*/ 处方审查明细
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 处方审查记录
  Set 待转出 = n_批次
  Where ID In (Select 审方id From 处方审查明细 Where 待转出 = n_批次);

  Update /*+ rule*/ 处方审查结果
  Set 待转出 = n_批次
  Where 审方id In (Select ID From 处方审查记录 Where 待转出 = n_批次);

  Update Zldatamovelog
  Set 当前进度 = '(9/10)医嘱数据标记完成，正在标记检查检验数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  Update /*+ rule*/ 影像检查记录
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 影像检查序列
  Set 待转出 = n_批次
  Where 检查uid In (Select 检查uid From 影像检查记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 影像检查图象
  Set 待转出 = n_批次
  Where 序列uid In (Select 序列uid From 影像检查序列 Where 待转出 = n_批次);

  Update /*+ rule*/ 影像申请单图像
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 影像收藏内容
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 影像危急值记录
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update Zldatamovelog
  Set 当前进度 = '(10/10)影像数据标记完成，正在标记检验数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  Update /*+ rule*/ 检验标本记录
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 检验申请项目
  Set 待转出 = n_批次
  Where 标本id In (Select ID From 检验标本记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 检验项目分布
  Set 待转出 = n_批次
  Where 标本id In (Select ID From 检验标本记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 检验分析记录
  Set 待转出 = n_批次
  Where 标本id In (Select ID From 检验标本记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 检验质控记录
  Set 待转出 = n_批次
  Where 标本id In (Select ID From 检验标本记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 检验操作记录
  Set 待转出 = n_批次
  Where 标本id In (Select ID From 检验标本记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 检验签名记录
  Set 待转出 = n_批次
  Where 检验标本id In (Select ID From 检验标本记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 检验图像结果
  Set 待转出 = n_批次
  Where 标本id In (Select ID From 检验标本记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 检验试剂记录
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 检验拒收记录
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 检验普通结果
  Set 待转出 = n_批次
  Where 检验标本id In (Select ID From 检验标本记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 检验质控报告
  Set 待转出 = n_批次
  Where 结果id In (Select ID From 检验普通结果 Where 待转出 = n_批次);
  Update /*+ rule*/ 检验药敏结果
  Set 待转出 = n_批次
  Where 细菌结果id In (Select ID From 检验普通结果 Where 待转出 = n_批次);
  Update /*+ rule*/ 检验流水线标本
  Set 待转出 = n_批次
  Where 标本id In (Select ID From 检验标本记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 检验流水线指标
  Set 待转出 = n_批次
  Where 标本id In (Select ID From 检验标本记录 Where 待转出 = n_批次);

  Commit;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl1_Datamove_Tag;
/



---------------------
--zlAppData
---------------------
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 1, 'A01', '药品皮试', '规定必须做皮试的药品，处方医师是否注明过敏试验及结果的判定', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 1, 'A02', '用药与临床诊断的相符性', '处方用药与临床诊断的相符性', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 1, 'A03', '剂量、用法的正确性', '剂量、用法的正确性', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 1, 'A04', '剂型与给药途径的合理性', '选用剂型与给药途径的合理性', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 1, 'A05', '是否有重复给药现象', '是否有重复给药现象', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 1, 'A06', '药物相互作用和配伍禁忌', '是否有潜在临床意义的药物相互作用和配伍禁忌', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 1, 'A07', '其它用药不适宜情况', '其它用药不适宜情况', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 2, '1-1', '处方内容缺', '处方的前记、正文、后记内容缺项', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 2, '1-2', '药师签名签章不一致', '医师签名、签章不规范或者与签名、签章的留样不一致的', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 2, '1-3', '药师未对处方进行适宜性审核的', '药师未对处方进行适宜性审核的（处方后记的审核、调配、核对、发药栏目无审核调配药师及核对发药药师签名，或者单人值班调剂未执行双签名规定）', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 2, '1-4', '新生儿、婴幼儿未写明日、月龄', '新生儿、婴幼儿处方未写明日、月龄的', 0, 0, 1, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 2, '1-5', '西药、中成药与中药饮片未分别开具处方', '西药、中成药与中药饮片未分别开具处方的', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 2, '1-6', '未使用药品规范名称开具处方', '未使用药品规范名称开具处方的', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 2, '1-7', '药品书写不规范或不清楚', '药品的剂量、规格、数量、单位等书写不规范或不清楚的', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 2, '1-8', '用法用量使用含糊不清字句', '用法、用量使用“遵医嘱”、“自用”等含糊不清字句的', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 2, '1-9', '处方修改未签名或药品超量未注明原因', '处方修改未签名并注明修改日期，或药品超剂量使用未注明原因和再次签名的', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 2, '1-10', '未写临床诊断或书写不全', '开具处方未写临床诊断或临床诊断书写不全的', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 2, '1-11', '单张门急诊处方超五种药品', '单张门急诊处方超过五种药品的', 0, 0, 0, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 2, '1-12', '延长处方用量未注明理由', '无特殊情况下，门诊处方超过7日用量，急诊处方超过3日用量，慢性病、老年病或特殊情况下需要适当延长处方用量未注明理由的', 0, 0, 0, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 2, '1-13', '开具特殊管理药品未执行国家规定', '开具麻醉药品、精神药品、医疗用毒性药品、放射性药品等特殊管理药品处方未执行国家有关规定的', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 2, '1-14', '未按抗菌药物管理开具', '医师未按照抗菌药物临床应用管理规定开具抗菌药物处方的', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 2, '1-15', '中药饮片未按“君臣佐使”排列', '中药饮片处方药物未按照“君、臣、佐、使”的顺序排列，或未按要求标注药物调剂、煎煮等特殊要求的', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 2, '2-1', '适应证不适宜', '适应证不适宜的', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 2, '2-2', '遴选的药品不适宜', '遴选的药品不适宜的', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 2, '2-3', '药品剂型或给药途径不适宜', '药品剂型或给药途径不适宜的', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 2, '2-4', '无正当理由不首选国家基本药物', '无正当理由不首选国家基本药物的', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 2, '2-5', '用法、用量不适宜', '用法、用量不适宜的', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 2, '2-6', '联合用药不适宜', '联合用药不适宜的', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 2, '2-7', '重复给药', '重复给药的', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 2, '2-8', '有配伍禁忌或者不良相互作用', '有配伍禁忌或者不良相互作用的', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 2, '2-9', '其它用药不适宜', '其它用药不适宜情况的', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 2, '3-1', '无适应证用药', '无适应证用药', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 2, '3-2', '无正当理由开具高价药', '无正当理由开具高价药的', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 2, '3-3', '无正当理由超说明书用药', '无正当理由超说明书用药的', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 2, '3-4', '无正当理由为同一患者开2种以上作用相同药物', '无正当理由为同一患者同时开具2种以上药理作用相同药物的', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 3, 'C01', 'PASS结果', '合理用药监测结果', 0, 0, 2, Null, user, sysdate From Dual;
Insert Into 处方审查项目 (ID,类别,编码,简称,内容,是否门诊启用,是否住院启用,服务对象,PASS结果,操作人,操作时间) Select 处方审查项目_ID.Nextval, 4, 'D01', '中药注射剂两种以上', '中药注射剂两种以上（含两种）', 0, 0, 2, Null, user, sysdate From Dual;

