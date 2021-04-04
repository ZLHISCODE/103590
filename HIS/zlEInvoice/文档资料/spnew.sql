CREATE TABLE 电子票据类别(
	编号   number(3),
	名称   varchar2(50),
	简码   varchar2(20),
	是否启用   number(2),
	部件 varchar2(100),
	包名称 varchar2(100))
 TABLESPACE zl9Expense;

Alter Table 电子票据类别 Add Constraint 电子票据类别_PK Primary Key(编号) Using Index Tablespace zl9Indexhis;
Alter Table 电子票据类别 Add Constraint 电子票据类别_UQ_名称  Unique(名称)  Using Index Tablespace zl9Indexhis;
 
Create Table 电子票据站点控制(
 场合 Number(2),
 站点 varchar2(50))
 TABLESPACE zl9Expense;
Alter Table 电子票据站点控制 Add Constraint 电子票据站点控制_PK Primary Key(站点,场合) Using Index Tablespace zl9Indexhis;


CREATE TABLE 电子票据开票点(
	ID   Number(18),
	上级ID   Number(18),
	编码   varchar2(20),
	名称   varchar2(50),
	简码   varchar2(20),
	院区   varchar2(50),
	客户端   varchar2(50),
	部门ID   number(18),
	位置   varchar2(100),
	末级   number(2),
	建档时间   date,
	撤档时间   date)
 TABLESPACE zl9Expense;

Alter Table 电子票据开票点 Add Constraint 电子票据开票点_PK Primary Key(ID) Using Index Tablespace zl9Indexhis;
Alter Table 电子票据开票点 Add Constraint 电子票据开票点_UQ_编码  Unique(编码, 撤档时间)  Using Index Tablespace zl9Indexhis;
Alter Table 电子票据开票点 Add Constraint 电子票据开票点_FK_上级ID Foreign Key (上级ID) References 电子票据开票点(ID) on delete cascade;
Alter Table 电子票据开票点 Add Constraint 电子票据开票点_FK_部门ID Foreign Key (部门ID) References 部门表(ID) on delete cascade;

CREATE INDEX 电子票据开票点_IX_简码 ON 电子票据开票点(简码) TABLESPACE zl9Indexhis;
CREATE INDEX 电子票据开票点_IX_部门ID ON 电子票据开票点(部门ID) TABLESPACE zl9Indexhis;


CREATE SEQUENCE 电子票据开票点_ID START WITH 1;  


CREATE TABLE 票据开票点对照(
    Id Number(18),
	开票点ID Number(18),
	人员ID Number(18),
	客户端 varchar2(50))
TABLESPACE zl9Expense;

Alter Table 票据开票点对照 Add Constraint 票据开票点对照_PK Primary Key(ID) Using Index Tablespace zl9Indexhis;
Alter Table 票据开票点对照 Add Constraint 票据开票点对照_UQ_开票点ID  Unique(开票点ID, 人员ID,客户端)  Using Index Tablespace zl9Indexhis;
Alter Table 票据开票点对照 Add Constraint 票据开票点对照_FK_人员ID Foreign Key (人员ID) References 人员表(ID) on delete cascade;
CREATE INDEX 票据开票点对照_IX_人员ID ON 票据开票点对照(人员ID) TABLESPACE zl9Indexhis;
CREATE INDEX 票据开票点对照_IX_客户端 ON 票据开票点对照(客户端) TABLESPACE zl9Indexhis;
CREATE SEQUENCE 票据开票点对照_ID START WITH 1;  


ALTER TABLE 病人预交记录 Add  (是否电子票据 number(2),预交电子票据 number(2));
ALTER TABLE 病人结帐记录 Add  是否电子票据 number(2);

ALTER TABLE 合约单位 ADD (社会信用代码 varchar2(50));
ALTER TABLE 保险类别 add(保险机构编码 varchar2(50));
CREATE INDEX 合约单位_IX_名称 ON 合约单位(名称) TABLESPACE zl9Indexhis;

Create Table 电子票据使用记录(
 ID Number(18),
 票种 number(2),
 记录状态 number(2),
 结算ID number(18),
 病人ID number(18),
 姓名 varchar2(100),
 性别 varchar2(4),
 年龄 varchar2(20),
 门诊号 number(18),
 住院号 number(18),
 代码 Varchar2(50),
 号码 Varchar2(50),
 检验码 Varchar2(20),
 凭证代码 Varchar2(50),
 凭证号码 Varchar2(50),
 凭证检验码 Varchar2(20),
 票据金额 number(16,5),
 生成时间 varchar2(30),
 URL内网  varchar2(2000),
 URL外网  varchar2(2000),
 原票据ID number(18),
 是否换开 number(2),
 纸质发票号 Varchar2(50),
 打印ID Number(18),
 退款id number(18),
 备注 varchar2(4000),
 开票点 varchar2(100),
 系统来源 varchar2(100),
 操作员编号 varchar2(6),
 操作员姓名 varchar2(50),
 登记时间 Date,
 待转出 number(3))
 TABLESPACE zl9Expense PCTFREE 5 initrans 20;

Alter Table 电子票据使用记录 Add Constraint 电子票据使用记录_PK Primary Key(ID) Using Index Tablespace zl9Indexhis;
Alter table 电子票据使用记录 Add Constraint 电子票据使用记录_UQ_号码 Unique(号码,票种,记录状态,代码)  Using Index Tablespace zl9Indexhis;
Alter Table 电子票据使用记录 Add Constraint 电子票据使用记录_FK_原票据ID Foreign Key (原票据ID) References 电子票据使用记录(ID);
Alter Table 电子票据使用记录 Add Constraint 电子票据使用记录_FK_病人ID Foreign Key (病人ID) References 病人信息(病人ID);

CREATE INDEX 电子票据使用记录_IX_登记时间 ON 电子票据使用记录(登记时间) TABLESPACE zl9Indexhis;
CREATE INDEX 电子票据使用记录_IX_生成时间 ON 电子票据使用记录(生成时间) TABLESPACE zl9Indexhis;
CREATE INDEX 电子票据使用记录_IX_结算ID ON 电子票据使用记录(结算ID) TABLESPACE zl9Indexhis;
CREATE INDEX 电子票据使用记录_IX_原票据ID ON 电子票据使用记录(原票据ID) TABLESPACE zl9Indexhis;

CREATE SEQUENCE 电子票据使用记录_ID START WITH 1;  

CREATE TABLE 电子票据二维码 (
 使用记录ID number(18),
 二维码 clob,
 待转出 number(3)) 
TABLESPACE zl9Expense PCTFREE 20;

ALTER TABLE 电子票据二维码 ADD CONSTRAINT 电子票据二维码_PK PRIMARY KEY (使用记录ID) USING INDEX TABLESPACE zl9Indexhis;
ALTER TABLE 电子票据二维码 ADD CONSTRAINT 电子票据二维码_FK_使用记录ID  FOREIGN KEY (使用记录ID ) REFERENCES 电子票据使用记录(ID)  On Delete Cascade;


ALTER TABLE 票据入库记录 ADD (是否下载 number(2));
ALTER TABLE 票据领用记录 ADD (是否下载 number(2));

ALTER TABLE 票据使用明细 ADD (电子票据ID number(18));
ALTER TABLE 票据使用明细 ADD CONSTRAINT 票据使用明细_FK_电子票据ID FOREIGN KEY(电子票据ID) REFERENCES 电子票据使用记录(ID);
CREATE INDEX 票据使用明细_IX_电子票据ID ON 票据使用明细(电子票据ID) TABLESPACE zl9Indexhis;

Insert into zlTables(系统,表名,表空间,分类) Values(&n_System,'电子票据类别','ZL9EXPENSE','A2');
Insert into zlTables(系统,表名,表空间,分类) Values(&n_System,'电子票据站点控制','ZL9EXPENSE','A2');
Insert into zlTables(系统,表名,表空间,分类) Values(&n_System,'电子票据开票点','ZL9EXPENSE','A2');
Insert into zlTables(系统,表名,表空间,分类) Values(&n_System,'票据开票点对照','ZL9EXPENSE','A2');
Insert into zlTables(系统,表名,表空间,分类) Values(&n_System,'电子票据使用记录','ZL9EXPENSE','B1');
Insert into zlTables(系统,表名,表空间,分类) Values(&n_System,'电子票据二维码','ZL9EXPENSE','B1');
Insert Into zlBakTables(系统,组号,表名,序号,直接转出,停用触发器)
select &n_system,1,A.* From (
Select 表名,序号,直接转出,停用触发器 From zlBakTables Where 1 = 0
Union All Select '电子票据使用记录',18,1,-NULL From Dual
Union All Select '电子票据二维码',19,1,-NULL From Dual) A;

Insert Into zlParameters
(ID,系统,模块,私有,本机,授权,固定,部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明)
Select Zlparameters_Id.Nextval, &n_System, -1*null, 0, 0, 0, 0, 0, 0,327 , '挂号电子票据控制', '', '0|1|0:','主要控制挂号业务是否启用电子票据'||CHR(13)||'1.启用了电子票据的业务:系统将不会再按本身的票据管理体系进行票据管理和控制，将通过调用中联电子票据接口部件进行电子票据的开具、作废、换开等,因此，客户端控件必须要有部件“zlElectronicInvoice.dll”且与编制了相关电子票据接口的。'||CHR(13)||'2.未启用电子票据的业务:票据、打印等都由HIS系统进行管理和控制。',
'格式:票据启用控制|票据管理控制|医保启用控制'||CHR(13)||'1. 票据启用控制:主要控制电子票据是否启用方式：0-表示未启用电子票据;1-代表启用电子票据;2-代表分站点启用电子票据  '||CHR(13)||'2.票据管理控制:主要是控制是否HIS系统管理票据:0-代表HIS管理票据;1-三方票据平台'||chr(13)||'3.医保启用控制:格式为启用标志:启用险类'||CHR(13)||'   a.启用标志:0-代表未启用;1-代表启用'||CHR(13)||'   b.启用险类:空代表所有医保启用;非空时，代表医保编号，多个医保时用逗号分离', '', '适用于某些医院需要启用电子票据管理业务', '在医院启用电子票据后，一般不调整此参数，如果调整此参数，将会影响到挂号业务的票据使用及打印。'
From Dual;

Insert Into zlParameters
(ID,系统,模块,私有,本机,授权,固定,部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明)
Select Zlparameters_Id.Nextval, &n_System, -1*null, 0, 0, 0, 0, 0, 0,328 , '收费电子票据控制', '', '0|1|0:','主要控制收费业务是否启用电子票据'||CHR(13)||'1.启用了电子票据的业务:系统将不会再按本身的票据管理体系进行票据管理和控制，将通过调用中联电子票据接口部件进行电子票据的开具、作废、换开等,因此，客户端控件必须要有部件“zlElectronicInvoice.dll”且与编制了相关电子票据接口的。'||CHR(13)||'2.未启用电子票据的业务:票据、打印等都由HIS系统进行管理和控制。',
'格式:票据启用控制|票据管理控制|医保启用控制'||CHR(13)||'1. 票据启用控制:主要控制电子票据是否启用方式：0-表示未启用电子票据;1-代表启用电子票据;2-代表分站点启用电子票据  '||CHR(13)||'2.票据管理控制:主要是控制是否HIS系统管理票据:0-代表HIS管理票据;1-三方票据平台'||chr(13)||'3.医保启用控制:格式为启用标志:启用险类'||CHR(13)||'   a.启用标志:0-代表未启用;1-代表启用'||CHR(13)||'   b.启用险类:空代表所有医保启用;非空时，代表医保编号，多个医保时用逗号分离', '', '适用于某些医院需要启用电子票据管理业务', '在医院启用电子票据后，一般不调整此参数，如果调整此参数，将会影响到收费业务的票据使用及打印。'
From Dual;

Insert Into zlParameters
(ID,系统,模块,私有,本机,授权,固定,部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明)
Select Zlparameters_Id.Nextval, &n_System, -1*null, 0, 0, 0, 0, 0, 0,329 , '预交电子票据控制', '', '0|0|1|0:','主要控制预交业务是否启用电子票据'||CHR(13)||'1.启用了电子票据的业务:系统将不会再按本身的票据管理体系进行票据管理和控制，将通过调用中联电子票据接口部件进行电子票据的开具、作废、换开等,因此，客户端控件必须要有部件“zlElectronicInvoice.dll”且与编制了相关电子票据接口的。'||CHR(13)||'2.未启用电子票据的业务:票据、打印等都由HIS系统进行管理和控制。',
'格式:预交类别|票据启用控制|票据管理控制|医保启用控制'||CHR(13)||'1. 预交类别:主要控制启用电子票据的预交类型：0-表示所有预交;1-代表门诊预交;2-代表住院预交  '||CHR(13)||'2. 票据启用控制:主要控制电子票据是否启用方式：0-表示未启用电子票据;1-代表启用电子票据;2-代表分站点启用电子票据  '||CHR(13)||'3.票据管理控制:主要是控制是否HIS系统管理票据:0-代表HIS管理票据;1-三方票据平台'||chr(13)||'4.医保启用控制:格式为启用标志:启用险类'||CHR(13)||'   a.启用标志:0-代表未启用;1-代表启用'||CHR(13)||'   b.启用险类:空代表所有医保启用;非空时，代表医保编号，多个医保时用逗号分离', '', '适用于某些医院需要启用电子票据管理业务', '在医院启用电子票据后，一般不调整此参数，如果调整此参数，将会影响到预交业务的票据使用及打印。'
From Dual;

Insert Into zlParameters
(ID,系统,模块,私有,本机,授权,固定,部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明)
Select Zlparameters_Id.Nextval, &n_System, -1*null, 0, 0, 0, 0, 0, 0,330 , '结帐电子票据控制', '', '0|1|0:','主要控制结帐业务是否启用电子票据'||CHR(13)||'1.启用了电子票据的业务:系统将不会再按本身的票据管理体系进行票据管理和控制，将通过调用中联电子票据接口部件进行电子票据的开具、作废、换开等,因此，客户端控件必须要有部件“zlElectronicInvoice.dll”且与编制了相关电子票据接口的。'||CHR(13)||'2.未启用电子票据的业务:票据、打印等都由HIS系统进行管理和控制。',
'格式:票据启用控制|票据管理控制|医保启用控制'||CHR(13)||'1. 票据启用控制:主要控制电子票据是否启用方式：0-表示未启用电子票据;1-代表启用电子票据;2-代表分站点启用电子票据  '||CHR(13)||'2.票据管理控制:主要是控制是否HIS系统管理票据:0-代表HIS管理票据;1-三方票据平台'||chr(13)||'3.医保启用控制:格式为启用标志:启用险类'||CHR(13)||'   a.启用标志:0-代表未启用;1-代表启用'||CHR(13)||'   b.启用险类:空代表所有医保启用;非空时，代表医保编号，多个医保时用逗号分离', '', '适用于某些医院需要启用电子票据管理业务', '在医院启用电子票据后，一般不调整此参数，如果调整此参数，将会影响到结帐业务的票据使用及打印。'
From Dual;

Insert Into zlParameters
(ID,系统,模块,私有,本机,授权,固定,部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明)
Select Zlparameters_Id.Nextval, &n_System, -1*null, 0, 0, 0, 0, 0, 0,331 , '就诊卡电子票据控制', '', '0|1|0:','主要控制发卡业务是否启用电子票据'||CHR(13)||'1.启用了电子票据的业务:系统将不会再按本身的票据管理体系进行票据管理和控制，将通过调用中联电子票据接口部件进行电子票据的开具、作废、换开等,因此，客户端控件必须要有部件“zlElectronicInvoice.dll”且与编制了相关电子票据接口的。'||CHR(13)||'2.未启用电子票据的业务:票据、打印等都由HIS系统进行管理和控制。',
'格式:票据启用控制|票据管理控制|医保启用控制'||CHR(13)||'1. 票据启用控制:主要控制电子票据是否启用方式：0-表示未启用电子票据;1-代表启用电子票据;2-代表分站点启用电子票据  '||CHR(13)||'2.票据管理控制:主要是控制是否HIS系统管理票据:0-代表HIS管理票据;1-三方票据平台'||chr(13)||'3.医保启用控制:格式为启用标志:启用险类'||CHR(13)||'   a.启用标志:0-代表未启用;1-代表启用'||CHR(13)||'   b.启用险类:空代表所有医保启用;非空时，代表医保编号，多个医保时用逗号分离', '', '适用于某些医院需要启用电子票据管理业务', '在医院启用电子票据后，一般不调整此参数，如果调整此参数，将会影响到发卡业务的票据使用及打印。'
From Dual;


Insert Into zlPrograms(序号,标题,说明,系统,部件) Values(1145,'电子票据操作','主要是主要是针对挂号、收费、预交及结帐等业务的电子票据的开票、打印、换票、退票等操作，有此权限时，才允许对电子票据的开具、打印、换票及退票的操作。',&n_System,'zL9CashBill');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select &n_System,1145,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0
Union All Select '基本',1,'',1 From Dual
Union All Select '参数设置',2,'针对参数进行操作的权限。有该权限时，允许进行本地参数设置',0 From Dual
Union All Select '开具电子票据',3,'主要是控制是否允许开具电子票据权限',0 From Dual
Union All Select '换开纸质票据',4,'主要是控制是否换开纸质票据权限.',0 From Dual
Union All Select '重新换开票据',5,'主要是控制是否重新换开纸质票据权限.',0 From Dual
Union All Select '作废纸质票据',6,'主要是控制是否允许作废已换开的纸质票据.',0 From Dual 
) A;
  
Insert Into zlModuleRelas(相关系统,模块,功能,系统,相关模块,相关类型,相关功能,缺省值)
Select  &n_System,1145,A.* From (
Select 功能,相关系统,相关模块,相关类型,相关功能,缺省值 From zlModuleRelas Where 1 = 0
Union All Select '基本',&n_System,1101,1,'基本',1 From Dual
Union All Select '基本',&n_System,1103,1,'基本',1 From Dual
Union All Select '基本',&n_System,1107,1,'基本',1 From Dual
Union All Select '基本',&n_System,1151,1,'基本',1 From Dual
Union All Select '基本',&n_System,1111,1,'基本',1 From Dual
Union All Select '基本',&n_System,1113,1,'基本',1 From Dual
Union All Select '基本',&n_System,1121,1,'基本',1 From Dual
Union All Select '基本',&n_System,1124,1,'基本',1 From Dual
Union All Select '基本',&n_System,1131,1,'基本',1 From Dual
Union All Select '基本',&n_System,1137,1,'基本',1 From Dual
Union All Select '基本',&n_System,1801,1,'基本',1 From Dual
Union All Select '基本',&n_System,1802,1,'基本',1 From Dual
Union All Select '基本',&n_System,1803,1,'基本',1 From Dual
Union All Select '基本',&n_System,1804,1,'基本',1 From Dual
Union All Select '基本',&n_System,1805,1,'基本',1 From Dual
Union All Select '基本',&n_System,1806,1,'基本',1 From Dual
Union All Select '基本',&n_System,1807,1,'基本',1 From Dual
Union All Select '基本',&n_System,1809,1,'基本',1 From Dual
Union All Select '基本',&n_System,1811,1,'基本',1 From Dual
) A;

Insert Into zlProgPrivs(系统, 序号, 功能, 所有者, 对象, 权限)
Select &n_System, 1145, '基本', User, A.*
From (Select 对象, 权限 From zlProgPrivs Where 1 = 0
Union All Select '电子票据站点控制','SELECT' From Dual
Union All Select '电子票据使用记录','SELECT' From Dual
Union All Select '票据入库记录','SELECT' From Dual
Union All Select '票据领用记录','SELECT' From Dual
Union All Select '票据使用明细','SELECT' From Dual
Union All Select '票据使用类别','SELECT' From Dual
Union All Select '票据打印内容','SELECT' From Dual
Union All Select '电子票据二维码','SELECT' From Dual
Union All Select '电子票据使用记录_ID','SELECT' From Dual
Union All Select '病人预交记录','SELECT' From Dual
Union All Select '门诊费用记录','SELECT' From Dual
Union All Select '费用补充记录','SELECT' From Dual
Union All Select '住院费用记录','SELECT' From Dual
Union All Select '病人挂号记录','SELECT' From Dual
Union All Select '保险结算记录','SELECT' From Dual
Union All Select '病人结帐记录','SELECT' From Dual
Union All Select '病人卡结算记录','SELECT' From Dual
Union All Select '三方结算交易','SELECT' From Dual
Union All Select '三方退款信息','SELECT' From Dual 
Union All Select '保险结算明细','SELECT' From Dual
Union All Select '保险类别','SELECT' From Dual
Union All Select '保险特准项目','SELECT' From Dual
Union All Select '保险支付大类','SELECT' From Dual
Union All Select '保险支付项目','SELECT' From Dual
Union All Select '大类档次比例','SELECT' From Dual
Union All Select '帐户年度信息','SELECT' From Dual
Union All Select '病区科室对应','SELECT' From Dual
Union All Select '病人余额','SELECT' From Dual
Union All Select '结算方式','SELECT' From Dual
Union All Select '结算方式应用','SELECT' From Dual
Union All Select '人员表','SELECT' From Dual
Union All Select '部门表','SELECT' From Dual
Union All Select '部门性质说明','SELECT' From Dual
Union All Select '收费分类目录','SELECT' From Dual
Union All Select '收费特定项目','SELECT' From Dual
Union All Select '收费细目','SELECT' From Dual
Union All Select '收费项目别名','SELECT' From Dual
Union All Select '收费项目类别','SELECT' From Dual
Union All Select '收费项目目录','SELECT' From Dual
Union All Select '收费执行科室','SELECT' From Dual
Union All Select '收据费目','SELECT' From Dual
Union All Select '收入项目','SELECT' From Dual
Union All Select '性别','SELECT' From Dual
Union All Select '费别','SELECT' From Dual
Union All Select '费别适用科室','SELECT' From Dual
Union All Select '材料特性','SELECT' From Dual
Union All Select '药品规格','SELECT' From Dual
Union All Select '药品目录','SELECT' From Dual
Union All Select '药品特性','SELECT' From Dual
Union All Select '药品信息','SELECT' From Dual
Union All Select '医保对照类别','SELECT' From Dual
Union All Select '医保对照明细','SELECT' From Dual
Union All Select '医保核对表','SELECT' From Dual
Union All Select '医疗付款方式','SELECT' From Dual
Union All Select '医疗卡挂失方式','SELECT' From Dual
Union All Select '诊疗分类目录','SELECT' From Dual
Union All Select '诊疗互斥项目','SELECT' From Dual
Union All Select '诊疗收费关系','SELECT' From Dual
Union All Select '诊疗项目目录','SELECT' From Dual
Union All Select '诊疗执行科室','SELECT' From Dual
Union All Select '证件类型','SELECT' From Dual
Union All Select '病人类型','SELECT' From Dual
Union All Select '消费卡类型','SELECT' From Dual
Union All Select '消费卡类别目录','SELECT' From Dual
Union All Select '消费卡信息','SELECT' From Dual
Union All Select '三方接口配置','SELECT' From Dual
Union All Select '电子票据开票点','SELECT' From Dual
Union All Select '票据开票点对照','SELECT' From Dual
Union All Select '电子票据类别','SELECT' From Dual
UNION ALL SELECT 'Zl_电子票据使用记录_Insert','EXECUTE' From dual
UNION ALL SELECT 'Zl_电子票据使用记录_Delete','EXECUTE' From dual
UNION ALL SELECT 'Zl_电子票据使用记录_Update','EXECUTE' From dual
UNION ALL SELECT 'Zl_电子票据二维码_Update','EXECUTE' From dual
UNION ALL SELECT '电子票据二维码','UPDATE' From dual
UNION ALL SELECT 'Zl_纸质票据使用_Update','EXECUTE' From dual
UNION ALL SELECT 'Zl_电子票据站点控制_Update','EXECUTE' From dual
UNION ALL SELECT 'zl_电子票据开票点_insert','EXECUTE' From dual
UNION ALL SELECT 'zl_电子票据开票点_update','EXECUTE' From dual
UNION ALL SELECT 'zl_电子票据开票点_start','EXECUTE' From dual
UNION ALL SELECT 'zl_电子票据开票点_stop','EXECUTE' From dual
UNION ALL SELECT 'zl_电子票据开票点_delete','EXECUTE' From dual
UNION ALL SELECT 'Zl_票据开票点对照_Update','EXECUTE' From dual
UNION ALL SELECT 'Zl_三方接口配置_Set','EXECUTE' From dual
UNION ALL SELECT 'Zl_三方接口配置_Get','EXECUTE' From dual
) A;

Insert Into zlParameters
(ID,系统,模块,私有,本机,授权,固定,部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明)
Select Zlparameters_Id.Nextval, &n_System, 1145, 0, 1,1, 0, 0, 0,1 , '票据换开方式', '', '0','主要控制在开具电子票据后针对纸质票据的换开控制,分三种：不换开，自动换开，提示换开.',
'0-不换开，1-自动换开，2-提示换开.', '1.针对挂号业务：需要设置“挂号电子票据控制”为启用才有效'||CHR(13)||'2.针对收费业务：需要设置“收费电子票据控制”为启用才有效'||CHR(13)||'3.针对预交业务：需要设置“预交电子票据控制”为启用才有效'||CHR(13)||'4.针对结帐业务：需要设置“结帐电子票据控制”为启用才有效', '适用于某些医院需要启用电子票据管理业务时，同步需要换开电子票据业务(主要是过渡期使用).', ''
From Dual;

Insert Into zlParameters
(ID,系统,模块,私有,本机,授权,固定,部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明)
Select Zlparameters_Id.Nextval, &n_System, 1145, 0, 1,1, 0, 0, 0,2 , '告知单打印方式', '', '0','主要控制在开具电子票据后是否打印告知单给患者,分三种：不打印，打印，提示打印.',
'0-不打印，1-自动打印，2-提示打印..', '启用了电子票据业务(挂号，收费，预交及结帐)参数后，本参数有效，调用报表:zl1_INSIDE_1145', '适用于某些医院需要启用电子票据管理业务时，同步需要打印告知道给患者.', ''
From Dual;

Insert Into zlParameters
(ID,系统,模块,私有,本机,授权,固定,部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明)
Select Zlparameters_Id.Nextval, &n_System, 1145, 0, 1,1, 0, 0, 0,3 , '电子票据打印方式', '', '0','主要控制在开具电子票据后是否打印电子票据给患者,分三种：不打印，打印，提示打印.',
'0-不打印，1-自动打印，2-提示打印..', '启用了电子票据业务(挂号，收费，预交及结帐)参数后，本参数有效，调用打印接口进行打印', '适用于某些医院需要启用电子票据管理业务时，同步需要打印电子票据给患者.', ''
From Dual;

Insert Into zlParameters
(ID,系统,模块,私有,本机,授权,固定,部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明)
Select Zlparameters_Id.Nextval, &n_System, 1145, 0, 0,1, 0, 0, 0,4 , '开票点对码方式', '', '1','主要控制电子票据开票点对码方式：0-按客户端对码，1-按收费员对码，2-按客户端和收费员对码.',
'0-按客户端对码，1-按收费员对码，2-按客户端和收费员对码', '启用了电子票据业务(挂号，收费，预交及结帐)参数后，本参数有效', '适用于某些医院需要启用电子票据管理业务时，同步需要打印电子票据给患者.', ''
From Dual;

Insert Into zlPrograms(序号,标题,说明,系统,部件) Values(1144,'电子票据管理','主要是针对电子票据的相关基础项目的对码、纸质票据下发、对帐及电子票据开具等功能的管理。',&n_System,'zL9CashBill');
Insert Into Zlmenus
  (组别, Id, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块)
  Select '缺省', Zlmenus_Id.Nextval, Id, '电子票据管理', '电子票据', 'E', 246, '主要是针对电子票据的相关基础项目的对码、纸质票据下发、对帐及电子票据开具等功能的管理。', &n_System, 1144
  From Zlmenus
  Where 标题 = '运营管理系统' And 组别 = '缺省' And 系统 = &n_System And 模块 Is Null AND ROWNUM <2;

Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select &n_System,1144,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0
Union All Select '基本',1,'',1 From Dual
Union All Select '基础数据管理',2,'主要控制基础数据的一些维护，比如：站点对码、收据费目对码等基础性设置.',0 From Dual
Union All Select '电子票据核对',3,'主要是针对电子票据的核对操作',0 From Dual
Union All Select '挂号票据核对',4,'主要是针对挂号业务发生的电子票据核对.',0 From Dual
Union All Select '收费票据核对',5,'主要是针对收费业务发生的电子票据核对.',0 From Dual
Union All Select '门诊预交票据核对',6,'主要是针对门诊预交业务发生的电子票据核对.',0 From Dual
Union All Select '住院预交票据核对',7,'主要是针对住院预交业务发生的电子票据核对.',0 From Dual
Union All Select '门诊结帐票据核对',8,'主要是针对门诊结帐业务发生的电子票据核对.',0 From Dual
Union All Select '住院结帐票据核对',9,'主要是针对住院结帐业务发生的电子票据核对.',0 From Dual
Union All Select '电子票据管理',10,'主要是针对未开具的电子票据进行批量开具操作.',0 From Dual
Union All Select '开具挂号电子票据',11,'主要是针对挂号未开具电子票据的记录进行电子票据的开具.',0 From Dual
Union All Select '开具收费电子票据',12,'主要是针对门诊收费未开具电子票据的记录进行电子票据的开具.',0 From Dual
Union All Select '开具门诊预交电子票据',13,'主要是针对门诊预交未开具电子票据的记录进行电子票据的开具.',0 From Dual
Union All Select '开具住院预交电子票据',14,'主要是针对住院预交未开具电子票据的记录进行电子票据的开具.',0 From Dual
Union All Select '开具门诊结帐电子票据',15,'主要是针对门诊结帐未开具电子票据的记录进行电子票据的开具.',0 From Dual
Union All Select '开具住院结帐电子票据',16,'主要是针对住院结帐未开具电子票据的记录进行电子票据的开具.',0 From Dual
Union All Select '纸质票据管理',17,'主要是针对未换开纸质票据的电子票据进行换开操作.',0 From Dual
Union All Select '换开挂号票据',18,'主要是针对挂号未换开纸质票据的电子票据记录进行换开.',0 From Dual
Union All Select '换开收费票据',19,'主要是针对收费未换开纸质票据的电子票据记录进行换开.',0 From Dual
Union All Select '换开门诊预交票据',20,'主要是针对门诊预交未换开纸质票据的电子票据记录进行换开.',0 From Dual
Union All Select '换开住院预交票据',21,'主要是针对住院预交未换开纸质票据的电子票据记录进行换开.',0 From Dual
Union All Select '换开门诊结帐票据',22,'主要是针对门诊结帐未换开纸质票据的电子票据记录进行换开.',0 From Dual
Union All Select '换开住院结帐票据',23,'主要是针对住院结帐未换开纸质票据的电子票据记录进行换开.',0 From Dual
) A;

--电子票据核对
Insert Into zlProgRelas(系统,序号,组号,功能,关系,主项,主项关系)
Select &n_System,1144,1,A.* From (
Select 功能,关系,主项,主项关系 From zlProgRelas Where 1 = 0
Union All Select '电子票据核对',2,1,1 From Dual
Union All Select '挂号票据核对',2,0,0 From Dual
Union All Select '收费票据核对',2,0,0 From Dual
Union All Select '门诊预交票据核对',2,0,0 From Dual
Union All Select '住院预交票据核对',2,0,0 From Dual
Union All Select '门诊结帐票据核对',2,0,0 From Dual
Union All Select '住院结帐票据核对',2,0,0 From Dual) A;

--电子票据管理
Insert Into zlProgRelas(系统,序号,组号,功能,关系,主项,主项关系)
Select &n_System,1144,2,A.* From (
Select 功能,关系,主项,主项关系 From zlProgRelas Where 1 = 0
Union All Select '电子票据管理',2,1,1 From Dual
Union All Select '开具挂号电子票据',2,0,0 From Dual
Union All Select '开具收费电子票据',2,0,0 From Dual
Union All Select '开具门诊预交电子票据',2,0,0 From Dual
Union All Select '开具住院预交电子票据',2,0,0 From Dual
Union All Select '开具门诊结帐电子票据',2,0,0 From Dual
Union All Select '开具住院结帐电子票据',2,0,0 From Dual) A;

--纸质票据管理
Insert Into zlProgRelas(系统,序号,组号,功能,关系,主项,主项关系)
Select &n_System,1144,3,A.* From (
Select 功能,关系,主项,主项关系 From zlProgRelas Where 1 = 0
Union All Select '纸质票据管理',2,1,1 From Dual
Union All Select '换开挂号票据',2,0,0 From Dual
Union All Select '换开收费票据',2,0,0 From Dual
Union All Select '换开门诊预交票据',2,0,0 From Dual
Union All Select '换开住院预交票据',2,0,0 From Dual
Union All Select '换开门诊结帐票据',2,0,0 From Dual
Union All Select '换开住院结帐票据',2,0,0 From Dual) A;



Insert Into zlProgPrivs(系统, 序号, 功能, 所有者, 对象, 权限)
Select &n_System, 1144, '基本', User, A.*
From (Select 对象, 权限 From zlProgPrivs Where 1 = 0
Union All Select '电子票据使用记录','SELECT' From Dual
Union All Select '电子票据站点控制','SELECT' From Dual
Union All Select '票据入库记录','SELECT' From Dual
Union All Select '票据领用记录','SELECT' From Dual
Union All Select '票据使用明细','SELECT' From Dual
Union All Select '票据使用类别','SELECT' From Dual
Union All Select '票据打印内容','SELECT' From Dual
Union All Select '电子票据二维码','SELECT' From Dual
Union All Select '电子票据使用记录_ID','SELECT' From Dual
Union All Select '电子票据开票点_ID','SELECT' From Dual
Union All Select '病人预交记录','SELECT' From Dual
Union All Select '门诊费用记录','SELECT' From Dual
Union All Select '费用补充记录','SELECT' From Dual
Union All Select '住院费用记录','SELECT' From Dual
Union All Select '病人挂号记录','SELECT' From Dual
Union All Select '保险结算记录','SELECT' From Dual
Union All Select '病人结帐记录','SELECT' From Dual
Union All Select '病人卡结算记录','SELECT' From Dual
Union All Select '三方结算交易','SELECT' From Dual
Union All Select '三方退款信息','SELECT' From Dual 
Union All Select '保险结算明细','SELECT' From Dual
Union All Select '保险类别','SELECT' From Dual
Union All Select '保险特准项目','SELECT' From Dual
Union All Select '保险支付大类','SELECT' From Dual
Union All Select '保险支付项目','SELECT' From Dual
Union All Select '大类档次比例','SELECT' From Dual
Union All Select '帐户年度信息','SELECT' From Dual
Union All Select '病区科室对应','SELECT' From Dual
Union All Select '病人余额','SELECT' From Dual
Union All Select '结算方式','SELECT' From Dual
Union All Select '结算方式应用','SELECT' From Dual
Union All Select '人员表','SELECT' From Dual
Union All Select '部门表','SELECT' From Dual
Union All Select '部门性质说明','SELECT' From Dual
Union All Select '收费分类目录','SELECT' From Dual
Union All Select '收费特定项目','SELECT' From Dual
Union All Select '收费细目','SELECT' From Dual
Union All Select '收费项目别名','SELECT' From Dual
Union All Select '收费项目类别','SELECT' From Dual
Union All Select '收费项目目录','SELECT' From Dual
Union All Select '收费执行科室','SELECT' From Dual
Union All Select '收据费目','SELECT' From Dual
Union All Select '收入项目','SELECT' From Dual
Union All Select '性别','SELECT' From Dual
Union All Select '费别','SELECT' From Dual
Union All Select '费别适用科室','SELECT' From Dual
Union All Select '材料特性','SELECT' From Dual
Union All Select '药品规格','SELECT' From Dual
Union All Select '药品目录','SELECT' From Dual
Union All Select '药品特性','SELECT' From Dual
Union All Select '药品信息','SELECT' From Dual
Union All Select '医保对照类别','SELECT' From Dual
Union All Select '医保对照明细','SELECT' From Dual
Union All Select '医保核对表','SELECT' From Dual
Union All Select '医疗付款方式','SELECT' From Dual
Union All Select '医疗卡挂失方式','SELECT' From Dual
Union All Select '诊疗分类目录','SELECT' From Dual
Union All Select '诊疗互斥项目','SELECT' From Dual
Union All Select '诊疗收费关系','SELECT' From Dual
Union All Select '诊疗项目目录','SELECT' From Dual
Union All Select '诊疗执行科室','SELECT' From Dual
Union All Select '证件类型','SELECT' From Dual
Union All Select '病人类型','SELECT' From Dual
Union All Select '消费卡类型','SELECT' From Dual
Union All Select '消费卡类别目录','SELECT' From Dual
Union All Select '消费卡信息','SELECT' From Dual
Union All Select '三方接口配置','SELECT' From Dual
Union All Select '电子票据开票点','SELECT' From Dual
Union All Select '票据开票点对照','SELECT' From Dual
Union All Select '电子票据类别','SELECT' From Dual
UNION ALL SELECT 'Zl_电子票据使用记录_Insert','EXECUTE' From dual
UNION ALL SELECT 'Zl_电子票据使用记录_Delete','EXECUTE' From dual
UNION ALL SELECT 'Zl_电子票据使用记录_Update','EXECUTE' From dual
UNION ALL SELECT 'Zl_电子票据二维码_Update','EXECUTE' From dual
UNION ALL SELECT '电子票据二维码','UPDATE' From dual
UNION ALL SELECT 'Zl_纸质票据使用_Update','EXECUTE' From dual
UNION ALL SELECT 'Zl_电子票据站点控制_Update','EXECUTE' From dual
UNION ALL SELECT 'zl_电子票据开票点_insert','EXECUTE' From dual
UNION ALL SELECT 'zl_电子票据开票点_update','EXECUTE' From dual
UNION ALL SELECT 'zl_电子票据开票点_start','EXECUTE' From dual
UNION ALL SELECT 'zl_电子票据开票点_stop','EXECUTE' From dual
UNION ALL SELECT 'zl_电子票据开票点_delete','EXECUTE' From dual
UNION ALL SELECT 'Zl_票据开票点对照_Update','EXECUTE' From dual
UNION ALL SELECT 'Zl_三方接口配置_Set','EXECUTE' From dual
UNION ALL SELECT 'Zl_三方接口配置_Get','EXECUTE' From dual
) A;

Insert Into zlProgPrivs(系统, 序号, 功能, 所有者, 对象, 权限)
Select &n_System, 1006, '基本', User, A.*
From (Select 对象, 权限 From zlProgPrivs Where 1 = 0
Union All Select 'zlClients','SELECT' From Dual
Union All Select '保险类别','SELECT' From Dual
Union All Select '电子票据站点控制','SELECT' From Dual
UNION ALL SELECT 'Zl_电子票据站点控制_Update','EXECUTE' From dual
Union All Select '电子票据类别','SELECT' From Dual
Union All Select 'Zl_电子票据类别_Update','EXECUTE' From Dual
) A;

Create Or Replace Procedure Zl_电子票据使用记录_Insert
(
  Id_In         In 电子票据使用记录.Id%Type,
  票种_In       In 电子票据使用记录.票种%Type,
  结算id_In     In 电子票据使用记录.结算id%Type,
  病人id_In     In 电子票据使用记录.病人id%Type,
  姓名_In       In 电子票据使用记录.姓名%Type,
  性别_In       In 电子票据使用记录.性别%Type,
  年龄_In       In 电子票据使用记录.年龄%Type,
  门诊号_In     In 电子票据使用记录.门诊号%Type,
  住院号_In     In 电子票据使用记录.住院号%Type,
  票据金额_In   In 电子票据使用记录.票据金额%Type,
  开票点_In     In 电子票据使用记录.开票点%Type,
  系统来源_In   In 电子票据使用记录.系统来源%Type,
  生成时间_In   In 电子票据使用记录.生成时间%Type,
  备注_In       In 电子票据使用记录.备注%Type,
  操作员编号_In In 电子票据使用记录.操作员编号%Type,
  操作员姓名_In In 电子票据使用记录.操作员姓名%Type,
  登记时间_In   In 电子票据使用记录.登记时间%Type,
  原票据id_In   In 电子票据使用记录.原票据id%Type := Null,
  退款id_In     In 电子票据使用记录.原票据id%Type := Null
) As
  n_记录状态 电子票据使用记录.记录状态%Type;
Begin
  n_记录状态 := 1;

  Insert Into 电子票据使用记录
    (ID, 票种, 记录状态, 结算id, 病人id, 姓名, 性别, 年龄, 门诊号, 住院号, 票据金额, 生成时间, 原票据id, 退款id, 开票点, 系统来源, 备注, 操作员编号, 操作员姓名, 登记时间)
  Values
    (Id_In, 票种_In, n_记录状态, 结算id_In, Decode(Nvl(病人id_In, 0), 0, Null, 病人id_In), 姓名_In, 性别_In, 年龄_In,
     Decode(Nvl(门诊号_In, 0), 0, Null, 门诊号_In), Decode(Nvl(住院号_In, 0), 0, Null, 住院号_In), 票据金额_In, 生成时间_In, 原票据id_In,
     退款id_In, 开票点_In, 系统来源_In, 备注_In, 操作员编号_In, 操作员姓名_In, Nvl(登记时间_In, Sysdate));
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_电子票据使用记录_Insert;
/

Create Or Replace Procedure Zl_电子票据二维码_Update
(
  使用记录id_In In 电子票据二维码.使用记录id%Type,
  是否删除_In   Number := 0,
  二维码_In     Varchar2 := Null
) As
  -- 是否删除_IN:1-表示删除;0-表示不删除
  n_Count Number(18);
Begin
  If Nvl(是否删除_In, 0) = 1 Then
    Delete 电子票据二维码 Where 使用记录id = 使用记录id_In;
    Return;
  End If;
  Select Count(1) Into n_Count From 电子票据二维码 Where 使用记录id = 使用记录id_In;
  If n_Count = 0 Then
    Insert Into 电子票据二维码 (使用记录id, 二维码) Values (使用记录id_In, 二维码_In);
  Else
    Update 电子票据二维码 Set 二维码 = 二维码_In Where 使用记录id = 使用记录id_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_电子票据二维码_Update;
/


Create Or Replace Procedure Zl_电子票据使用记录_Update
(
  Id_In         In 电子票据使用记录.Id%Type,
  代码_In       In 电子票据使用记录.代码%Type,
  号码_In       In 电子票据使用记录.号码%Type,
  检验码_In     In 电子票据使用记录.检验码%Type,
  生成时间_In   In 电子票据使用记录.生成时间%Type,
  Url内网_In    In 电子票据使用记录.Url内网%Type,
  Url外网_In    In 电子票据使用记录.Url外网%Type,
  备注_In       In 电子票据使用记录.备注%Type,
  开票点_In     In 电子票据使用记录.开票点%Type,
  系统来源_In   In 电子票据使用记录.系统来源%Type,
  票据金额_In   In 电子票据使用记录.票据金额%Type := Null,
  凭证代码_In   In 电子票据使用记录.凭证代码%Type := Null,
  凭证号码_In   In 电子票据使用记录.凭证号码%Type := Null,
  凭证检验码_In In 电子票据使用记录.凭证检验码%Type := Null
) As
Begin

  Update 电子票据使用记录
  Set 代码 = Nvl(代码_In, 代码), 号码 = Nvl(号码_In, 号码), 检验码 = Nvl(检验码_In, 检验码), 生成时间 = 生成时间_In, Url内网 = Nvl(Url内网_In, Url内网),
      Url外网 = Nvl(Url外网_In, Url外网), 备注 = Nvl(备注_In, 备注), 开票点 = Nvl(开票点_In, 开票点), 系统来源 = 系统来源_In,
      票据金额 = Nvl(票据金额_In, 票据金额), 凭证代码 = Nvl(凭证代码_In, 凭证代码), 凭证号码 = Nvl(凭证号码_In, 凭证号码), 凭证检验码 = Nvl(凭证检验码_In, 凭证检验码)
  Where ID = Id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_电子票据使用记录_Update;
/

Create Or Replace Procedure Zl_电子票据使用记录_Delete
(
  Id_In           In 电子票据使用记录.Id%Type,
  开票点_In       In 电子票据使用记录.开票点%Type,
  系统来源_In     In 电子票据使用记录.系统来源%Type,
  生成时间_In     In 电子票据使用记录.生成时间%Type,
  备注_In         In 电子票据使用记录.备注%Type,
  操作员编号_In   In 电子票据使用记录.操作员编号%Type,
  操作员姓名_In   In 电子票据使用记录.操作员姓名%Type,
  登记时间_In     In 电子票据使用记录.登记时间%Type,
  原电子票据id_In In 电子票据使用记录.Id%Type
) As
  v_Err_Msg Varchar2(200);
  Err_Item Exception;
  n_是否换开 电子票据使用记录.是否换开%Type;
Begin

  Update 电子票据使用记录 Set 记录状态 = 3 Where ID = 原电子票据id_In Returning Nvl(是否换开, 0) Into n_是否换开;
  If Sql%NotFound Then
    v_Err_Msg := '未找到原始的电子票据信息，不能作废操作!';
    Raise Err_Item;
  End If;
  If Nvl(n_是否换开, 0) = 1 Then
    --当前电子票据已经换开纸质票据
    v_Err_Msg := '当前电子票据已经换开纸质票据,需要先冲红纸质票据后才能作废电子发票!';
    Raise Err_Item;
  End If;

  Insert Into 电子票据使用记录
    (ID, 票种, 记录状态, 结算id, 病人id, 姓名, 性别, 年龄, 门诊号, 住院号, 代码, 号码, 检验码, 票据金额, Url内网, Url外网, 生成时间, 原票据id, 打印id, 是否换开, 纸质发票号,
     开票点, 系统来源, 备注, 操作员编号, 操作员姓名, 登记时间, 退款id)
    Select Id_In, 票种, 2, 结算id, 病人id, 姓名, 性别, 年龄, 门诊号, 住院号, 代码, 号码, 检验码, 票据金额, Url内网, Url外网, 生成时间_In, 原电子票据id_In, 打印id,
           是否换开, 纸质发票号, Nvl(开票点_In, 开票点) As 开票点, Nvl(系统来源_In, 系统来源) As 系统来源, Nvl(备注_In, 备注) As 备注, 操作员编号_In, 操作员姓名_In,
           登记时间_In, 退款id
    From 电子票据使用记录
    Where ID = 原电子票据id_In;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_电子票据使用记录_Delete;
/

Create Or Replace Procedure Zl_纸质票据使用_Update
(
  数据来源_In   票据打印内容.数据性质%Type,
  票种_In       票据使用明细.票种%Type,
  结算id_In     病人预交记录.结帐id%Type,
  电子票据id_In 电子票据使用记录.Id%Type,
  票据号_In     Varchar2,
  票据金额_In   票据使用明细.票据金额%Type,
  领用id_In     票据使用明细.领用id%Type,
  使用人_In     票据使用明细.使用人%Type,
  使用时间_In   票据使用明细.使用时间%Type,
  操作方式_In   Integer := 0,
  是否补结算_In Number := 0,
  红票类型_In   Number := 0
) As
  --功能：用换开、重开及作废纸质票据
  --参数：
  --     操作方式_In:0-换开;1-重新换开;2-作废票据;3-回收票据
  --     数据来源_IN =1-收费,2-预交(包含押金),3-结帐,4-挂号,5-就诊卡
  --     票种_in: 1-收费,2-预交,3-结帐,4-挂号
  --      结算ID_In:票种_in=2:原预交ID;其他为结帐ID(原结帐ID)
  --      票据号_IN：多个用逗号分离;
  --      领用ID:如果为0或NULL,表示不严格控制票据。
  --      是否补结算_In-0-不是补结算;1-是补结算
  --      红票类型_In-0-不是红票,1-余额退款(或转出预交)产生的红票.目前仅针对预交有效(票种_In=2)

  c_No t_StrList := t_StrList();

  n_收回id   票据打印内容.Id%Type;
  n_打印id   票据打印内容.Id%Type;
  v_发票号   票据使用明细.号码%Type;
  v_最大发票 票据使用明细.号码%Type;
  n_仅换开   Number(2);
  n_Count    Number(18);
  v_Error    Varchar2(255);
  Err_Custom Exception;
  n_原因 Number(2);
Begin

  If Nvl(数据来源_In, 0) = 1 Then
    --收费
    If Nvl(是否补结算_In, 0) = 0 Then
      Select NO Bulk Collect Into c_No From (Select Distinct NO From 门诊费用记录 Where 结帐id = 结算id_In);
    
    Else
      Select NO Bulk Collect Into c_No From (Select Distinct NO From 费用补充记录 Where 结算id = 结算id_In);
    
    End If;
  
  Elsif Nvl(数据来源_In, 0) = 2 Then
    --预交
    Select NO Bulk Collect Into c_No From (Select Distinct NO From 病人预交记录 Where ID = 结算id_In);
  
  Elsif Nvl(数据来源_In, 0) = 3 Then
    --结帐
    Select NO Bulk Collect Into c_No From (Select Distinct NO From 病人结帐记录 Where ID = 结算id_In);
  
  Elsif Nvl(数据来源_In, 0) = 4 Then
    --挂号
    If Nvl(是否补结算_In, 0) = 0 Then
      Select NO Bulk Collect Into c_No From (Select Distinct NO From 门诊费用记录 Where 结帐id = 结算id_In);
    
    Else
      Select NO Bulk Collect Into c_No From (Select Distinct NO From 费用补充记录 Where 结算id = 结算id_In);
    
    End If;
  Elsif Nvl(数据来源_In, 0) = 5 Then
    --就诊卡
    Select NO Bulk Collect Into c_No From (Select Distinct NO From 住院费用记录 Where 结帐id = 结算id_In);
  
  Else
    v_Error := '无效数据来源(' || Nvl(数据来源_In, 0) || '),无法进行换开数据！';
    Raise Err_Custom;
  End If;
  If c_No.Count = 0 Then
    v_Error := '未找到对应的结算数据(' || Nvl(结算id_In, 0) || '),无法进行换开数据！';
    Raise Err_Custom;
  End If;

  --1.先收回票据
  Begin
    If Nvl(红票类型_In, 0) = 1 Then
      If Nvl(操作方式_In, 0) > 0 Then
        Select ID
        Into n_收回id
        From (Select b.Id
               From 票据使用明细 A, 票据打印内容 B
               Where a.打印id = b.Id And a.性质 = 1 And a.原因 = 6 And b.数据性质 = 数据来源_In And b.No = c_No(1) And a.票种 = 票种_In And
                     Not Exists
                (Select 1 From 票据使用明细 B Where a.号码 = b.号码 And a.打印id = b.打印id And 性质 = 2)
               Order By a.使用时间 Desc)
        Where Rownum < 2;
        n_原因 := 7;
      Else
        n_仅换开 := 1;
      End If;
    Else
      Select ID
      Into n_收回id
      From (Select b.Id
             From 票据使用明细 A, 票据打印内容 B
             Where a.打印id = b.Id And a.性质 = 1 And a.原因 In (1, 3) And b.数据性质 = 数据来源_In And b.No = c_No(1) And a.票种 = 票种_In And
                   Not Exists
              (Select 1 From 票据使用明细 B Where a.号码 = b.号码 And a.打印id = b.打印id And 性质 = 2)
             Order By a.使用时间 Desc)
      Where Rownum < 2;
      n_原因 := 2;
    End If;
  Exception
    When Others Then
      n_仅换开 := 1;
  End;

  --收回票据(可能以前未控制票据,无法收回)
  If n_收回id Is Not Null Then
    --Decode(操作方式_In, 2, 5, 2)
    Insert Into 票据使用明细
      (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用人, 使用时间, 票据金额, 电子票据id)
      Select 票据使用明细_Id.Nextval, 票种, 号码, 2, n_原因, 领用id, 打印id, 使用人_In, 使用时间_In, 票据金额, 电子票据id
      From 票据使用明细 A
      Where 打印id = n_收回id And 性质 = 1 And a.票种 = 票种_In And Not Exists
       (Select 1 From 票据使用明细 B Where a.号码 = b.号码 And 打印id = n_收回id And 性质 = 2);
  Else
    n_仅换开 := 1;
  End If;

  --无票据号时,不用处理票据
  If 票据号_In Is Null Or Nvl(操作方式_In, 0) >= 2 Then
  
    If Nvl(数据来源_In, 0) = 1 Then
      --收费
    
      If Nvl(是否补结算_In, 0) = 0 Then
        Update 门诊费用记录
        Set 实际票号 = Null
        Where Mod(记录性质, 10) = 1 And NO In (Select Column_Value From Table(c_No));
      Else
        Update 费用补充记录
        Set 实际票号 = Null
        Where Mod(记录性质, 10) = 1 And NO In (Select Column_Value From Table(c_No));
      End If;
    
    Elsif Nvl(数据来源_In, 0) = 2 Then
      --预交
      If Not Nvl(红票类型_In, 0) = 1 Then
        Update 病人预交记录
        Set 实际票号 = Null
        Where Mod(记录性质, 10) = 1 And NO In (Select Column_Value From Table(c_No));
      End If;
    Elsif Nvl(数据来源_In, 0) = 3 Then
      --结帐
    
      Update 病人结帐记录 Set 实际票号 = v_发票号 Where NO In (Select Column_Value From Table(c_No));
    
    Elsif Nvl(数据来源_In, 0) = 4 Then
      --挂号
      If Nvl(是否补结算_In, 0) = 0 Then
        Update 门诊费用记录
        Set 实际票号 = Null
        Where Mod(记录性质, 10) = 4 And NO In (Select Column_Value From Table(c_No));
      Else
        Update 费用补充记录
        Set 实际票号 = Null
        Where Mod(记录性质, 10) = 1 And NO In (Select Column_Value From Table(c_No));
      End If;
    
    End If;
    Update 电子票据使用记录 Set 打印id = Null, 是否换开 = 0, 纸质发票号 = Null Where ID = Nvl(电子票据id_In, 0);
    Return;
  End If;

  v_发票号 := Substr(票据号_In || ',', 1, Instr(票据号_In || ',', ',') - 1);

  --重新发出票据并填写票据打印内容
  Select 票据打印内容_Id.Nextval Into n_打印id From Dual;

  Insert Into 票据打印内容
    (ID, 数据性质, NO, 打印类型)
    Select Distinct n_打印id, 数据来源_In, NO, 0 From (Select Distinct Column_Value As NO From Table(c_No));

  If Nvl(红票类型_In, 0) = 1 And Nvl(票种_In, 0) = 2 Then
    n_原因 := 6;
  Else
    If n_仅换开 = 1 Then
      n_原因 := 1;
    Else
      n_原因 := 3;
    End If;
  End If;

  Insert Into 票据使用明细
    (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额, 电子票据id)
    Select 票据使用明细_Id.Nextval, 票种_In, Column_Value As 发票号, 1, n_原因, Decode(Nvl(领用id_In, 0), 0, Null, 领用id_In), n_打印id,
           使用时间_In, 使用人_In, 票据金额_In, 电子票据id_In
    From Table(f_Str2List(票据号_In));

  If Nvl(领用id_In, 0) <> 0 Then
    Select Count(*) As n_Count, Max(Column_Value) Into n_Count, v_最大发票 From Table(f_Str2List(票据号_In));
  
    Update 票据领用记录 Set 剩余数量 = Nvl(剩余数量, 0) - n_Count, 当前号码 = v_最大发票 Where ID = 领用id_In;
  
  End If;
  If Nvl(数据来源_In, 0) = 1 Then
    --收费
  
    If Nvl(是否补结算_In, 0) = 0 Then
      Update 门诊费用记录
      Set 实际票号 = v_发票号
      Where Mod(记录性质, 10) = 1 And NO In (Select Column_Value From Table(c_No));
    Else
      Update 费用补充记录
      Set 实际票号 = v_发票号
      Where Mod(记录性质, 10) = 1 And NO In (Select Column_Value From Table(c_No));
    End If;
  
  Elsif Nvl(数据来源_In, 0) = 2 Then
    --预交
    If Not Nvl(红票类型_In, 0) = 1 Then
      Update 病人预交记录
      Set 实际票号 = v_发票号
      Where Mod(记录性质, 10) = 1 And NO In (Select Column_Value From Table(c_No));
    End If;
  Elsif Nvl(数据来源_In, 0) = 3 Then
    --结帐
  
    Update 病人结帐记录 Set 实际票号 = v_发票号 Where NO In (Select Column_Value From Table(c_No));
  
  Elsif Nvl(数据来源_In, 0) = 4 Then
    --挂号
    If Nvl(是否补结算_In, 0) = 0 Then
      Update 门诊费用记录
      Set 实际票号 = v_发票号
      Where Mod(记录性质, 10) = 4 And NO In (Select Column_Value From Table(c_No));
    Else
      Update 费用补充记录
      Set 实际票号 = v_发票号
      Where Mod(记录性质, 10) = 1 And NO In (Select Column_Value From Table(c_No));
    End If;
  
    --ELSIF nvl(数据来源_in,0)=5 THEN  --就诊卡
    --就诊卡无实际票据，待以后扩展
    --UPDATE 住院费用记录 SET 实票票号=v_发票号 WHERE  mod(记录性质,10)=1 AND NO IN (Select Column_value From table(c_NO));
  End If;

  Update 电子票据使用记录 Set 打印id = n_打印id, 是否换开 = 1, 纸质发票号 = v_发票号 Where ID = Nvl(电子票据id_In, 0);

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_纸质票据使用_Update;
/


Create Or Replace Procedure Zl_电子票据站点控制_Update
(
  场合_In In 电子票据站点控制.场合%Type,
  站点_In In Clob := Null
) Is

  --说明：
  --     场合_In：1-收费,2-预交,3-结帐,4-挂号
  --     站点_In：电子票据使用记录.站点,多个用逗号分隔,不传入站点表示仅删除 电子票据站点控制
Begin
  Delete From 电子票据站点控制 A Where a.场合 = 场合_In;

  For r_站点 In (Select Column_Value As 站点 From Table(f_Str2List(站点_In))) Loop
    Insert Into 电子票据站点控制 (场合, 站点) Values (场合_In, r_站点.站点);
  End Loop;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_电子票据站点控制_Update;
/

Create Or Replace Function Zl_Fun_Isstarteinvoice
(
  场合_In     Integer,
  险类_In     保险结算记录.险类%Type := 0,
  检查站点_In Integer := 1,
  类型_In     Integer := Null
) Return Number Is
  ---------------------------------------------------------------------------
  --功能：判断指定声合是否启用了电子票据 
  --参数：场合_In：1-收费,2-预交,3-结帐,4-挂号;5-就诊卡
  --      检查站点_IN-1-表示需要检查站点是否启用;0-不检查，直接判断业务是否启用
  --      类型_In:null-不区分类型;场合=预交和结帐，分别代表1-门诊和2-住院
  --返回:1-启用电子票据;0-未启用电子票据
  ---------------------------------------------------------------------------
  v_机器名       Varchar2(100);
  v_Para         Varchar2(4000);
  v_医保         Varchar2(4000);
  n_电子票据启用 Number(2);
  n_医保启用     Number(2);
  n_类型         Number(2);

  n_Return Number(2);
Begin

  If Nvl(场合_In, 0) = 1 Then
    v_Para := zl_GetSysParameter('收费电子票据控制');
  Elsif Nvl(场合_In, 0) = 2 Then
    v_Para := zl_GetSysParameter('预交电子票据控制');
    --格式：预交类别|票据启用控制|票据管理控制|医保启用控制
    v_Para := Nvl(v_Para, '') || '|||||';
    n_类型 := To_Number(Substr(v_Para, 1, Instr(v_Para, '|') - 1));
    If n_类型 <> 0 And Nvl(类型_In, 0) <> 0 And Nvl(类型_In, 0) <> n_类型 Then
      Return 0;
    End If;
    v_Para := Substr(v_Para, Instr(v_Para, '|') + 1);
  Elsif Nvl(场合_In, 0) = 3 Then
    v_Para := zl_GetSysParameter('结帐电子票据控制');
  Elsif Nvl(场合_In, 0) = 4 Then
    v_Para := zl_GetSysParameter('挂号电子票据控制');
  Elsif Nvl(场合_In, 0) = 5 Then
    v_Para := zl_GetSysParameter('就诊卡电子票据控制');
  Else
    Return 0;
  End If;

  --格式：票据启用控制|票据管理控制|医保启用控制
  v_Para         := Nvl(v_Para, '') || '|||||';
  n_电子票据启用 := To_Number(Substr(v_Para, 1, Instr(v_Para, '|') - 1));
  If Nvl(n_电子票据启用, 0) = 0 Then
    Return 0;
  End If;

  n_Return := 1;
  If Nvl(n_电子票据启用, 0) = 2 And Nvl(检查站点_In, 1) = 1 Then
    --0-表示未启用电子票据;1-代表启用电子票据;2-代表分站点启用电子票据
    Select Terminal Into v_机器名 From V$session Where Audsid = Userenv('sessionid');
    Select Nvl(Max(1), 0) Into n_Return From 电子票据站点控制 Where 场合 = Nvl(场合_In, 0) And 站点 = v_机器名;
  End If;
  If Nvl(险类_In, 0) = 0 Then
    --非医保，直接返回
    Return n_Return;
  End If;

  --医保验证
  v_Para := Substr(v_Para, Instr(v_Para, '|') + 1);
  v_Para := Substr(v_Para, Instr(v_Para, '|') + 1);
  v_Para := Substr(v_Para, 1, Instr(v_Para, '|') - 1);
  If Instr(v_Para, ':') > 0 Then
    --医保相关：启用标志:启用险类
    v_医保     := Substr(v_Para, Instr(v_Para, ':') + 1);
    n_医保启用 := To_Number(Substr(v_Para, 1, Instr(v_Para, ':') - 1));
  End If;
  If Nvl(n_医保启用, 0) = 0 Or v_医保 Is Null Then
    Return 0;
  End If;
  If Instr(',' || v_医保 || ',', ',' || 险类_In || ',') > 0 Then
    Return 1;
  End If;
  Return 0;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Fun_Isstarteinvoice;
/


Create Or Replace Procedure Zl_批量结帐结算_Update
(
  病人id_In       门诊费用记录.病人id%Type,
  结帐id_In       病人预交记录.结帐id%Type,
  保险结算_In     Varchar2,
  保险类别_In     保险类别.名称%Type,
  支付方式_In     结算方式.名称%Type,
  操作员编号_In   病人预交记录.操作员编号%Type,
  操作员姓名_In   病人预交记录.操作员姓名%Type,
  收款时间_In     病人预交记录.收款时间%Type,
  完成结算_In     Number := 0,
  是否电子票据_In 病人预交记录.是否电子票据%Type := Null,
  险类_In         保险结算记录.险类%Type := Null
) As
  ------------------------------------------------------------------------------------------------------------------------------
  --功能:收费结算时,修改结算的相关信息
  -- 保险结算_In:(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额||.."
  -- 是否电子票据_In-为空时，内部根据属性进行判断是否启用
  -- 结帐类型_IN:1-门诊结帐;2-住院结帐
  -- 完成结算_In:1-完成收费;0-未完成收费
  ------------------------------------------------------------------------------------------------------------------------------
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_结算内容     Varchar2(500);
  v_当前结算     Varchar2(50);
  n_预交id       病人预交记录.Id%Type;
  n_主页id       病人预交记录.主页id%Type;
  n_科室id       病人预交记录.科室id%Type;
  v_结算方式     病人预交记录.结算方式%Type;
  n_结算金额     病人预交记录.冲预交%Type;
  n_剩余款       病人预交记录.冲预交%Type;
  n_结帐金额     病人预交记录.冲预交%Type;
  n_返回值       人员缴款余额.余额%Type;
  v_误差费       结算方式.名称%Type;
  n_误差金额     病人预交记录.冲预交%Type;
  n_是否电子票据 病人预交记录.是否电子票据%Type;
  n_Count        Number;
  n_Havenull     Number;
  l_预交id       t_NumList := t_NumList();
  n_缴款组id     病人预交记录.缴款组id%Type;
  n_充值id       病人预交记录.Id%Type;
  d_交易时间     病人预交记录.交易时间%Type;
  v_交易人员     病人预交记录.交易人员%Type;
  n_险类         保险结算记录.险类%Type;
  Cursor c_Balance_Record Is
    Select Max(m.病人id) As 病人id, Max(NO) As NO, Max(Nvl(收款时间_In, m.收费时间)) As 收费时间, Max(Nvl(操作员编号_In, m.操作员编号)) As 操作员编号,
           Max(Nvl(m.操作员姓名, 操作员姓名_In)) As 操作员姓名, Max(Nvl(n_缴款组id, m.缴款组id)) As 缴款组id, Max(结帐类型) As 结帐类型
    
    From 病人结帐记录 M
    Where m.Id = 结帐id_In;
  r_Balance_Record c_Balance_Record%RowType;

  Cursor c_Balancedata Is
    Select 记录性质, NO, 记录状态, 病人id, 主页id, 科室id, 结算方式, Nvl(收款时间_In, 收款时间) As 收款时间, Nvl(操作员编号_In, 操作员编号) As 操作员编号,
           Nvl(操作员姓名_In, 操作员姓名) As 操作员姓名, 冲预交, 结帐id, Nvl(n_缴款组id, 缴款组id) As 缴款组id, 结算序号
    From 病人预交记录
    Where 结帐id = 结帐id_In And 结算方式 Is Null;

  r_Balancedata c_Balancedata%RowType;

  Procedure 病人预交记录_冲预交
  (
    冲预交_In        In Out 病人预交记录.冲预交%Type,
    预交类别_In      病人预交记录.预交类别%Type,
    冲预交病人ids_In Varchar2 := Null,
    结算性质_In      Number := 2
  ) As
    --费用余额检查_In  0-预交余额检查时不减去费用余额，1-减去费用余额；2-根据金额，有多少冲多少。
    --冲预交_In:如果费用余额检查_In=2，则返回未分摊完成的结算金额;其他为NULL或0
    v_冲预交病人ids Varchar2(4000);
    n_返回值        人员缴款余额.余额%Type;
    n_预交金额      病人预交记录.冲预交%Type;
    n_冲预交        病人预交记录.冲预交%Type;
    n_会话号        病人预交记录.会话号%Type; --格式：SID+'_'+SERIAL#
    n_组id          财务缴款分组.Id%Type;
  Begin
    v_冲预交病人ids := Nvl(冲预交病人ids_In, 病人id_In);
    n_组id          := Zl_Get组id(操作员姓名_In);
  
    --预交款处理
    If Nvl(冲预交_In, 0) = 0 Then
      Return;
    End If;
    Select Max(Sid || '_' || Serial#) Into n_会话号 From V$session Where Audsid = Userenv('sessionid');
  
    n_预交金额 := 冲预交_In; --先缴先用，且先用自己的
    --不包含结算方式为代收款项的预交款。
    For c_冲预交 In (Select a.No, b.预交余额 As 金额, Nvl(a.结帐id, 0) As 结帐id, a.病人id, a.记录状态, a.Id, a.收款时间, a.关联交易id
                  From 病人预交记录 A, 预交单据余额 B
                  Where a.Id = b.预交id And b.病人id In (Select Column_Value From Table(f_Num2List(v_冲预交病人ids))) And
                        Nvl(b.预交类别, 2) = 预交类别_In And a.结算方式 Not In (Select 名称 From 结算方式 Where 性质 = 5) And
                        Nvl(a.校对标志, 0) = 0
                  Order By Decode(病人id, Nvl(病人id_In, 0), 0, 1), a.收款时间) Loop
    
      If c_冲预交.金额 - n_预交金额 < 0 Then
        n_冲预交 := c_冲预交.金额;
      Else
        n_冲预交 := n_预交金额;
      End If;
    
      If c_冲预交.结帐id = 0 Then
        --第一次冲预交(填上结帐ID,金额为0)
        Update 病人预交记录
        Set 冲预交 = 0, 结帐id = 结帐id_In, 结算序号 = -1 * 结帐id_In, 结算性质 = 结算性质_In, 会话号 = n_会话号
        Where ID = c_冲预交.Id;
      End If;
      --冲上次剩余额
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
         结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 结算性质, 会话号, 关联交易id, 交易时间, 交易人员, 校对标志)
        Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间_In,
               操作员姓名_In, 操作员编号_In, n_冲预交, 结帐id_In, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, -1 * 结帐id_In,
               结算性质_In, n_会话号, c_冲预交.关联交易id, 收款时间_In, 操作员姓名_In, 0
        From 病人预交记录
        Where NO = c_冲预交.No And 记录状态 = c_冲预交.记录状态 And 记录性质 In (1, 11) And Rownum = 1;
    
      --更新数据(结算方式为NULL的)
      Update 病人预交记录
      Set 冲预交 = 冲预交 - n_冲预交
      Where 结帐id = 结帐id_In And 结算方式 Is Null
      Returning Nvl(冲预交, 0) Into n_返回值;
    
      --更新病人预交余额
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) - n_冲预交
      Where 病人id = c_冲预交.病人id And 性质 = 1 And 类型 = 预交类别_In
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (c_冲预交.病人id, 预交类别_In, -1 * n_冲预交, 1);
        n_返回值 := -1 * n_冲预交;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 病人余额
        Where 病人id = c_冲预交.病人id And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
      End If;
    
      --更新预交单据余额
      Update 预交单据余额
      Set 预交余额 = Nvl(预交余额, 0) - n_冲预交
      Where 病人id = c_冲预交.病人id And 预交id = c_冲预交.Id
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 预交单据余额
          (预交id, 病人id, 预交类别, 预交余额)
        Values
          (c_冲预交.Id, c_冲预交.病人id, 预交类别_In, -1 * n_冲预交);
        n_返回值 := -1 * n_冲预交;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 预交单据余额 Where 预交id = c_冲预交.Id And Nvl(预交余额, 0) = 0;
      End If;
    
      --检查是否已经处理完
      If c_冲预交.金额 < n_预交金额 Then
        n_预交金额 := n_预交金额 - c_冲预交.金额;
      Else
        n_预交金额 := 0;
      End If;
      If n_预交金额 = 0 Then
        Exit;
      End If;
    
    End Loop;
    --检查金额是否足够
    冲预交_In := n_预交金额;
  End 病人预交记录_冲预交;

Begin

  If 操作员姓名_In Is Null Then
    n_缴款组id := Null;
  Else
    n_缴款组id := Zl_Get组id(操作员姓名_In);
  End If;

  --0.正式结算
  Select Max(Decode(结算方式, Null, 1, 0)) Into n_Havenull From 病人预交记录 Where 结帐id = 结帐id_In;

  If Nvl(n_Count, 0) = 0 Then
    --增加结算方式为NULL的记录
    Begin
      Select 名称 Into v_误差费 From 结算方式 Where Nvl(性质, 0) = 9;
    Exception
      When Others Then
        v_误差费 := '误差费';
    End;
  End If;

  --1.增加结算方式为空的结算数据
  If Nvl(n_Havenull, 0) = 0 Then
  
    Open c_Balance_Record;
    Fetch c_Balance_Record
      Into r_Balance_Record;
  
    If c_Balance_Record%RowCount = 0 Then
      Close c_Balance_Record;
      v_Err_Msg := '未找到结帐记录,可能因为并发原因删除了结帐数据,请重新操作结帐!';
      Raise Err_Item;
    End If;
  
    Select Sum(Nvl(结帐金额, 0))
    Into n_结算金额
    From (Select Sum(结帐金额) As 结帐金额
           From 住院费用记录
           Where 结帐id = 结帐id_In
           Union All
           Select Sum(结帐金额) As 结帐金额
           From 门诊费用记录
           Where 结帐id = 结帐id_In);
  
    n_误差金额 := Round(n_结算金额 - Round(Nvl(n_结算金额, 0), 2), 6);
    n_结算金额 := Round(Nvl(n_结算金额, 0), 2);
  
    Select a.主页id, a.出院科室id
    Into n_主页id, n_科室id
    From 病案主页 A, 病人信息 B
    Where a.病人id = 病人id_In And a.病人id = b.病人id And a.主页id = b.主页id;
  
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 科室id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 校对标志, 结算性质)
    Values
      (病人预交记录_Id.Nextval, 2, r_Balance_Record.No, 1, Decode(病人id_In, 0, Null, 病人id_In), Decode(n_主页id, 0, Null, n_主页id),
       Decode(n_科室id, 0, Null, n_科室id), Null, r_Balance_Record.收费时间, r_Balance_Record.操作员编号, r_Balance_Record.操作员姓名,
       n_结算金额, 结帐id_In, r_Balance_Record.缴款组id, 1, 2);
  
    --误差费(先汇总后生成误差费
    If n_误差金额 <> 0 Then
      Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 结算性质, 关联交易id)
      Values
        (n_预交id, 2, r_Balance_Record.No, 1, Decode(病人id_In, 0, Null, 病人id_In), v_误差费, r_Balance_Record.收费时间,
         r_Balance_Record.操作员编号, r_Balance_Record.操作员姓名, n_误差金额, 结帐id_In, r_Balance_Record.缴款组id, -1 * 结帐id_In, 1, 2,
         n_预交id);
    End If;
    Close c_Balance_Record;
  End If;

  Open c_Balancedata;
  Fetch c_Balancedata
    Into r_Balancedata;

  --1.先回退数据
  n_结算金额 := 0;
  n_剩余款   := 0;
  For c_结算 In (Select a.Id, a.No, a.记录性质 As 记录性质, a.结算方式, a.病人id, a.冲预交, a.预交类别, b.性质,
                      Decode(a.记录性质, 1, Decode(a.记录状态, 2, 0, a.Id)) As 预交id
               From 病人预交记录 A, 结算方式 B
               Where a.结帐id = 结帐id_In And a.结算方式 = b.名称(+)) Loop
  
    If c_结算.性质 <> 9 And c_结算.结算方式 Is Not Null Then
      If c_结算.记录性质 = 1 Or c_结算.记录性质 = 11 Then
      
        Update 病人余额
        Set 预交余额 = Nvl(预交余额, 0) + Nvl(c_结算.冲预交, 0)
        Where 病人id = Nvl(c_结算.病人id, 0) And 性质 = 1 And 类型 = Nvl(c_结算.预交类别, 2)
        Returning 预交余额 Into n_返回值;
        If Sql%NotFound Then
          Insert Into 病人余额
            (病人id, 类型, 预交余额, 性质)
          Values
            (c_结算.病人id, Nvl(c_结算.预交类别, 2), Nvl(c_结算.冲预交, 0), 1);
        
          n_返回值 := Nvl(c_结算.冲预交, 0);
        End If;
      
        If Nvl(n_返回值, 0) = 0 Then
          Delete 病人余额 Where 性质 = 1 And 病人id = c_结算.病人id And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
        End If;
      
        --更新预交单据余额
        n_充值id := c_结算.预交id;
        If Nvl(n_充值id, 0) = 0 Then
          Select Max(ID) Into n_充值id From 病人预交记录 Where NO = c_结算.No And 记录性质 = 1 And 记录状态 <> 2;
        End If;
      
        If n_充值id <> 0 Then
        
          Update 预交单据余额
          Set 预交余额 = Nvl(预交余额, 0) + Nvl(c_结算.冲预交, 0)
          Where 病人id = c_结算.病人id And 预交id = n_充值id
          Returning 预交余额 Into n_返回值;
        
          If Sql%RowCount = 0 Then
            Insert Into 预交单据余额
              (预交id, 病人id, 预交类别, 预交余额)
            Values
              (n_充值id, c_结算.病人id, Nvl(c_结算.预交类别, 2), Nvl(c_结算.冲预交, 0));
            n_返回值 := Nvl(c_结算.冲预交, 0);
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 预交单据余额 Where 预交id = n_充值id And Nvl(预交余额, 0) = 0;
          End If;
        
        End If;
      End If;
      n_结算金额 := Nvl(n_结算金额, 0) + Nvl(c_结算.冲预交, 0);
      If c_结算.记录性质 = 11 Then
        l_预交id.Extend;
        l_预交id(l_预交id.Count) := c_结算.Id;
      Else
        Update 病人预交记录 Set 结帐id = Null, 冲预交 = Null Where ID = c_结算.Id;
      End If;
    
    End If;
  
    n_剩余款 := Nvl(n_剩余款, 0) + Nvl(c_结算.冲预交, 0);
  End Loop;
  n_剩余款 := Nvl(n_剩余款, 0) - Nvl(n_误差金额, 0);

  If Nvl(n_结算金额, 0) <> 0 Then
    Update 病人预交记录 Set 冲预交 = Nvl(冲预交, 0) + Nvl(n_结算金额, 0) Where 结帐id = 结帐id_In And 结算方式 Is Null;
    If Sql%NotFound Then
      v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[病人结帐窗口]中重新结帐！';
      Raise Err_Item;
    End If;
  End If;
  If l_预交id.Count <> 0 Then
    Forall I In 1 .. l_预交id.Count
      Delete 病人预交记录 Where ID = l_预交id(I);
  End If;

  --2.再处理医保结算

  If Not 保险结算_In Is Null Then
    n_结算金额 := 0;
    v_结算内容 := 保险结算_In || '||';
    n_预交id   := Null;
    d_交易时间 := Sysdate;
    v_交易人员 := zl_UserName;
  
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      n_结算金额 := To_Number(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
    
      If n_结算金额 <> 0 Then
        n_剩余款 := Nvl(n_剩余款, 0) - Nvl(n_结算金额, 0);
        If Nvl(n_预交id, 0) = 0 Then
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
        End If;
      
        If Nvl(n_充值id, 0) = 0 Then
          n_充值id := n_预交id;
        End If;
      
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 校对标志, 结算性质, 缴款单位,
           关联交易id, 交易时间, 交易人员)
        Values
          (n_预交id, 2, r_Balancedata.No, 1, r_Balancedata.病人id, r_Balancedata.主页id, r_Balancedata.科室id, '保险结算', v_结算方式,
           r_Balancedata.收款时间, r_Balancedata.操作员编号, r_Balancedata.操作员姓名, n_结算金额, r_Balancedata.结帐id, r_Balancedata.缴款组id,
           2, 2, 保险类别_In, n_充值id, d_交易时间, v_交易人员);
      
        --更新数据(结算方式为NULL的)
        Update 病人预交记录
        Set 冲预交 = 冲预交 - n_结算金额
        Where 结帐id = 结帐id_In And 结算方式 Is Null
        Returning Nvl(冲预交, 0) Into n_返回值;
      End If;
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
      n_预交id   := Null;
    End Loop;
  
    If Nvl(完成结算_In, 0) = 1 Then
      --医保相关表的处理
      Update 保险结算明细 Set 标志 = 2 Where 结帐id = 结帐id_In;
    End If;
  End If;

  --预交款处理
  If Nvl(n_剩余款, 0) <> 0 Then
  
    If Nvl(病人id_In, 0) = 0 Then
      v_Err_Msg := '不能确定病人的病人ID,收费不能使用预交款结算,结算操作失败！';
      Raise Err_Item;
    End If;
  
    病人预交记录_冲预交(n_剩余款, 2, Null, 2);
  
    If Nvl(n_剩余款, 0) > 0 Then
    
      If 支付方式_In Is Null Then
        v_Err_Msg := '不能确定缴款方式，请检查缴款方式是否正确,结算操作失败！';
        Raise Err_Item;
      End If;
      n_结帐金额 := n_剩余款;
    
      Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员编号, 操作员姓名, 冲预交,
         结帐id, 缴款, 找补, 缴款组id, 校对标志, 结算性质, 关联交易id)
      Values
        (n_预交id, r_Balancedata.No, Null, 2, 1, 病人id_In, r_Balancedata.主页id, r_Balancedata.科室id, Null, 支付方式_In, Null,
         '结帐缴款', Null, Null, Null, r_Balancedata.收款时间, 操作员编号_In, 操作员姓名_In, n_结帐金额, 结帐id_In, Null, Null,
         r_Balancedata.缴款组id, 2, 2, n_预交id);
      --更新数据(结算方式为NULL的)
      Update 病人预交记录
      Set 冲预交 = 冲预交 - n_结帐金额
      Where 结帐id = 结帐id_In And 结算方式 Is Null
      Returning Nvl(冲预交, 0) Into n_返回值;
    
    End If;
  
  End If;
  If Nvl(完成结算_In, 0) = 0 Then
    Close c_Balancedata;
    Return;
  End If;

  -----------------------------------------------------------------------------------------
  --完成收费,需要处理人员缴款余额,预交记录(结算方式=NULL)
  --1.删除结算方式为NULL的预交记录
  Delete 病人预交记录 Where 结帐id = 结帐id_In And 结算方式 Is Null And Nvl(冲预交, 0) = 0;
  If Sql%NotFound Then
    Select Count(*) Into n_Count From 病人预交记录 A Where 结帐id = 结帐id_In And 结算方式 Is Null;
    If n_Count <> 0 Then
      v_Err_Msg := '还存在未缴款的数据,不能完成结算!';
    Else
      v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！!';
    End If;
    Raise Err_Item;
  End If;

  --结算金额为零时，增加一条金额为0的病人预交记录
  Select Count(*) Into n_Count From 病人预交记录 A Where 结帐id = 结帐id_In;

  If n_Count = 0 Then
    v_结算方式 := 支付方式_In;
    If v_结算方式 Is Null Then
      Select Max(结算方式) Into v_结算方式 From 结算方式应用 Where 应用场合 = '结帐' And Nvl(缺省标志, 0) = 1;
      If v_结算方式 Is Null Then
        Select Nvl(Max(名称), '现金') Into v_结算方式 From 结算方式 Where Nvl(性质, 0) = 1;
      End If;
    End If;
  
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 校对标志, 卡类别id, 结算卡序号, 卡号,
       交易流水号, 交易说明, 结算号码, 结算性质)
    Values
      (病人预交记录_Id.Nextval, 2, r_Balancedata.No, 1, r_Balancedata.病人id, r_Balancedata.主页id, r_Balancedata.科室id, '结帐缴款',
       v_结算方式, r_Balancedata.收款时间, r_Balancedata.操作员编号, r_Balancedata.操作员姓名, 0, r_Balancedata.结帐id, r_Balancedata.缴款组id,
       2, Null, Null, Null, Null, Null, Null, 2);
  End If;

  n_是否电子票据 := 是否电子票据_In;
  If 是否电子票据_In Is Null Then
    n_险类 := Nvl(险类_In, 0);
    If 险类_In Is Null Then
      Select Max(险类) Into n_险类 From 保险结算记录 Where 记录id = 结帐id_In And 性质 = 2;
    End If;
    n_是否电子票据 := Zl_Fun_Isstarteinvoice(3, n_险类);
  End If;
  --2.处理缴款数据和找补数据及校对标志更新为0
  Update 病人预交记录 Set 校对标志 = 0, 是否电子票据 = n_是否电子票据 Where 结帐id = 结帐id_In;

  --3.更新费用状态
  Update 病人结帐记录 Set 结算状态 = Null,是否电子票据 = n_是否电子票据 Where ID = 结帐id_In;

  --4.更新人员缴款数据
  For c_缴款 In (Select 结算方式, 操作员姓名, Sum(Nvl(a.冲预交, 0)) As 冲预交
               From 病人预交记录 A
               Where a.结帐id = 结帐id_In And Mod(a.记录性质, 10) <> 1
               Group By 结算方式, 操作员姓名) Loop
  
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) + Nvl(c_缴款.冲预交, 0)
    Where 收款员 = c_缴款.操作员姓名 And 性质 = 1 And 结算方式 = c_缴款.结算方式;
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额
        (收款员, 结算方式, 性质, 余额)
      Values
        (c_缴款.操作员姓名, c_缴款.结算方式, 1, Nvl(c_缴款.冲预交, 0));
    End If;
  End Loop;
  Close c_Balancedata;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_批量结帐结算_Update;
/


Create Or Replace Procedure Zl_病人结帐结算_Modify
(
  操作类型_In      Number,
  病人id_In        门诊费用记录.病人id%Type,
  结帐id_In        病人预交记录.结帐id%Type,
  结算方式_In      Varchar2,
  冲预交_In        病人预交记录.冲预交%Type := Null,
  退支票额_In      病人预交记录.冲预交%Type := Null,
  卡类别id_In      病人预交记录.卡类别id%Type := Null,
  卡号_In          病人预交记录.卡号%Type := Null,
  交易流水号_In    病人预交记录.交易流水号%Type := Null,
  交易说明_In      病人预交记录.交易说明%Type := Null,
  缴款_In          病人预交记录.缴款%Type := Null,
  找补_In          病人预交记录.找补%Type := Null,
  误差金额_In      门诊费用记录.实收金额%Type := Null,
  结帐类型_In      Number := 2,
  缺省结算方式_In  结算方式.名称%Type := Null,
  操作员编号_In    病人预交记录.操作员编号%Type := Null,
  操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
  收款时间_In      病人预交记录.收款时间%Type := Null,
  冲预交病人ids_In Varchar2 := Null,
  完成结算_In      Number := 0,
  校对标志_In      Number := 2,
  预交id_In        病人预交记录.Id%Type := Null,
  关联交易id_In    病人预交记录.Id%Type := Null,
  清除原交易_In    Number := 0,
  附加标志_In      病人预交记录.Id%Type := Null,
  加入会话_In      Number := 1,
  是否电子票据_In  病人预交记录.是否电子票据%Type := Null
) As
  ------------------------------------------------------------------------------------------------------------------------------
  --功能:收费结算时,修改结算的相关信息
  --操作类型_In:
  --   0-普通收费方式:
  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要|卡号||.." ;也允许传入空.
  --     ②退支票额_In:如果涉及退支票,则传入本次的退支票额,非正常收费时,传入零
  --   1.三方卡结算:
  --     ①结算方式_IN:可以传入多个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
  --     ②退支票额_In:传入零
  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
  --     ④校对标志_IN:可以传入，不传为2:1-待校对的结算;2-接口调用成功或支付成功
  --     ④是否转帐_IN:三方方才传入
  --     @关联交易id_In:三方结算才传入
  --     @预交id_In:三方结算才传入
  --   2-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
  --     ②退支票额_In:传入零
  --   3-消费卡结算:
  --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."  消费卡ID:为零时,根据卡号自动定位
  --     ②冲预交_In: 传入零
  --     ②退支票额_In:传入零
  -- 冲预交_In: 存在冲预交时,传入
  -- 误差金额_In:存在误差费时,传入
  --  结帐类型_IN:1-门诊结帐;2-住院结帐
  --冲预交病人ids_In:多个用逗分分离,冲预交时有效(冲预交类别与业务操作保持一致),主要是使用家属的预交款
  -- 完成结算_In:1-完成收费;0-未完成收费
  -- 附加标志_IN:对于结帐业务(三方卡退款传入)：NULl or 0-普通业务;1-分交易退款,2-调用一次交易接口退款;3-转帐方式退款
  -- 加入会话_In:1-表示加放会话，0-表示不加入会话
  --是否电子票据_In:null-表示过程内部直接判断，非空表示直接以传入的为准
  ------------------------------------------------------------------------------------------------------------------------------
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_结算内容   Varchar2(500);
  v_当前结算   Varchar2(300);
  v_卡号       病人医疗卡信息.卡号%Type;
  n_消费卡id   消费卡信息.Id%Type;
  n_卡类别id   病人预交记录.结算卡序号%Type;
  v_名称       Varchar2(100);
  n_预交id     病人预交记录.Id%Type;
  n_关联交易id 病人预交记录.Id%Type;
  v_结算方式   病人预交记录.结算方式%Type;
  n_结算金额   病人预交记录.冲预交%Type;
  n_返回值     人员缴款余额.余额%Type;
  n_冲预交     病人预交记录.冲预交%Type;
  v_退支票     病人预交记录.结算方式%Type;
  v_结算号码   病人预交记录.结算号码%Type;
  v_结算摘要   病人预交记录.摘要%Type;
  v_误差费     结算方式.名称%Type;
  n_误差金额   病人预交记录.冲预交%Type;
  n_科室id     病人预交记录.科室id%Type;
  n_Count      Number;
  n_Havenull   Number;
  l_预交id     t_NumList := t_NumList();
  n_缴款组id   病人预交记录.缴款组id%Type;
  n_费用金额   门诊费用记录.结帐金额%Type;
  n_结帐金额   病人预交记录.冲预交%Type;
  v_交易人员   病人预交记录.交易人员%Type;
  d_交易时间   病人预交记录.交易时间%Type;

  n_险类         保险结算记录.险类%Type;
  n_校对标志     病人预交记录.校对标志%Type;
  n_是否未退     三方退款信息.是否未退%Type;
  n_退款金额     三方退款信息.金额%Type;
  v_会话号       病人预交记录.会话号%Type; --格式：SID+'_'+SERIAL#
  n_是否电子票据 Number(2);
  Cursor c_Balance_Record Is
    Select 病人id, NO, Nvl(收款时间_In, 收费时间) As 收费时间, Nvl(操作员编号_In, 操作员编号) As 操作员编号, Nvl(操作员姓名_In, 操作员姓名) As 操作员姓名,
           Nvl(n_缴款组id, 缴款组id) As 缴款组id, 结帐类型 As 结帐类型, 主页id
    From 病人结帐记录
    Where ID = 结帐id_In;

  r_Balance_Record c_Balance_Record%RowType;

  Cursor c_Balancedata Is
    Select 记录性质, NO, 记录状态, 病人id, 主页id, 科室id, 结算方式, Nvl(收款时间_In, 收款时间) As 收款时间, Nvl(操作员编号_In, 操作员编号) As 操作员编号,
           Nvl(操作员姓名_In, 操作员姓名) As 操作员姓名, 冲预交, 结帐id, Nvl(n_缴款组id, 缴款组id) As 缴款组id, 结算序号
    From 病人预交记录
    Where 结帐id = 结帐id_In And 结算方式 Is Null;

  r_Balancedata c_Balancedata%RowType;

Begin

  n_误差金额 := 误差金额_In;

  If Nvl(加入会话_In, 0) = 1 Then
    Select Max(Sid || '_' || Serial#) Into v_会话号 From V$session Where Audsid = Userenv('sessionid');
  End If;

  --0.正式结算
  Select Count(1), Max(Decode(结算方式, Null, 1, 0)), Sum(冲预交), Max(Decode(结算方式, Null, 缴款组id, 0))
  Into n_Count, n_Havenull, n_冲预交, n_缴款组id
  From 病人预交记录
  Where 结帐id = 结帐id_In;

  If Nvl(n_缴款组id, 0) = 0 Then
    n_缴款组id := Null;
  End If;

  --1.增加结算方式为空的结算数据
  If Nvl(n_Havenull, 0) = 0 Then
    n_Count := 0;
    Open c_Balance_Record;
    Fetch c_Balance_Record
      Into r_Balance_Record;
  
    If c_Balance_Record%NotFound Then
      Close c_Balance_Record;
      v_Err_Msg := '未找到指定的结帐数据,当前结算操作失败！';
      Raise Err_Item;
    End If;
  
    n_缴款组id := r_Balance_Record.缴款组id;
    If Nvl(n_缴款组id, 0) = 0 Then
      n_缴款组id := Null;
    End If;
    Select Sum(Nvl(结帐金额, 0))
    Into n_结算金额
    From (Select Sum(结帐金额) As 结帐金额
           From 住院费用记录
           Where 结帐id = 结帐id_In
           Union All
           Select Sum(结帐金额) As 结帐金额
           From 门诊费用记录
           Where 结帐id = 结帐id_In);
  
    n_误差金额 := n_结算金额 - Round(Nvl(n_结算金额, 0), 6);
    n_结算金额 := Round(Nvl(n_结算金额, 0) - Nvl(n_冲预交, 0), 6);
  
    n_科室id := Null;
    If Nvl(r_Balance_Record.结帐类型, 0) = 2 Then
      --住院的才有科室ID
      Select Max(当前科室id) Into n_科室id From 病人信息 Where 病人id = 病人id_In;
    End If;
  
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 科室id, 主页id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 校对标志, 结算性质, 会话号)
    Values
      (病人预交记录_Id.Nextval, 2, r_Balance_Record.No, 1, r_Balance_Record.病人id, n_科室id, r_Balance_Record.主页id, Null,
       r_Balance_Record.收费时间, r_Balance_Record.操作员编号, r_Balance_Record.操作员姓名, n_结算金额, 结帐id_In, r_Balance_Record.缴款组id, 1,
       2, v_会话号);
  
    n_误差金额 := Nvl(n_误差金额, 0) + Nvl(误差金额_In, 0);
    Close c_Balance_Record;
  End If;

  Open c_Balancedata;
  Fetch c_Balancedata
    Into r_Balancedata;

  If c_Balancedata%NotFound Then
    Close c_Balancedata;
    v_Err_Msg := '未找到指定的结算数据,结算操作失败！';
    Raise Err_Item;
  End If;

  If Nvl(清除原交易_In, 0) = 1 Then
    --处理校对标志为1的记录
    n_结算金额 := 0;
  
    For c_校对 In (Select ID, 冲预交
                 From 病人预交记录
                 Where 记录性质 = 2 And 结帐id = r_Balancedata.结帐id And Nvl(卡类别id, 0) = 卡类别id_In And
                       关联交易id = Nvl(关联交易id_In, 0)) Loop
      n_结算金额 := Round(n_结算金额 + Nvl(c_校对.冲预交, 0), 5);
    
      l_预交id.Extend;
      l_预交id(l_预交id.Count) := c_校对.Id;
    
    End Loop;
  
    If n_结算金额 <> 0 Then
      Update 病人预交记录
      Set 冲预交 = Nvl(冲预交, 0) + Nvl(n_结算金额, 0)
      Where 结帐id = 结帐id_In And 结算方式 Is Null;
      If Sql%NotFound Then
        v_Err_Msg := ' 结算信息错误, 可能因为并发原因造成结算信息错误, 请在 [ 收费结算窗口 ] 中重新收费！ ';
        Raise Err_Item;
      End If;
    
    End If;
    If l_预交id.Count <> 0 Then
      --预防删除错误，带上结帐ID
      Forall I In 1 .. l_预交id.Count
        Delete 病人预交记录 Where ID = l_预交id(I) And 结帐id + 0 = 结帐id_In;
    End If;
  
  End If;

  --2.处理误差费
  If Nvl(n_误差金额, 0) <> 0 Then
  
    Select Nvl(Max(名称), '误差费') Into v_误差费 From 结算方式 Where Nvl(性质, 0) = 9;
  
    Update 病人预交记录
    Set 冲预交 = Nvl(冲预交, 0) + Nvl(n_误差金额, 0)
    Where 结帐id = 结帐id_In And 结算方式 = v_误差费;
  
    If Sql%NotFound Then
      Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
    
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 科室id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 校对标志, 卡类别id, 结算卡序号, 卡号,
         交易流水号, 交易说明, 结算号码, 结算性质, 关联交易id, 会话号)
      Values
        (n_预交id, 2, r_Balancedata.No, 1, r_Balancedata.病人id, r_Balancedata.科室id, r_Balancedata.主页id, Null, v_误差费,
         r_Balancedata.收款时间, r_Balancedata.操作员编号, r_Balancedata.操作员姓名, n_误差金额, r_Balancedata.结帐id, r_Balancedata.缴款组id,
         2, Null, Null, 卡号_In, 交易流水号_In, 交易说明_In, Null, 2, n_预交id, v_会话号);
    End If;
  
    Update 病人预交记录 Set 冲预交 = 冲预交 - Nvl(n_误差金额, 0) Where 结帐id = 结帐id_In And 结算方式 Is Null;
    If Sql%NotFound Then
      v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！';
      Raise Err_Item;
    End If;
  End If;

  --3.处理冲预款
  If Nvl(冲预交_In, 0) <> 0 Then
    If Nvl(病人id_In, 0) = 0 Then
      v_Err_Msg := '不能确定病人的病人ID,不能使用预交款结算,结算操作失败！';
      Raise Err_Item;
    End If;
    Zl_病人预交记录_冲预交(病人id_In, 结帐id_In, 冲预交_In, 结帐类型_In, r_Balancedata.操作员编号, r_Balancedata.操作员姓名, r_Balancedata.收款时间,
                  冲预交病人ids_In, 2, 1);
  End If;

  n_预交id     := 预交id_In;
  n_关联交易id := 关联交易id_In;

  --4.处理普通结算信息
  If 操作类型_In = 0 Then
  
    If Nvl(退支票额_In, 0) <> 0 Then
      --0-普通收费方式且是支票
      Select Max(b.名称)
      Into v_退支票
      From 结算方式应用 A, 结算方式 B
      Where a.应用场合 = '结帐' And b.名称 = a.结算方式 And Nvl(b.应付款, 0) = 1;
    
      If v_退支票 Is Null Then
        v_Err_Msg := '在结算场合中,不存在结算性质为应付款的结算方式,请在[结算方式]中设置！';
        Raise Err_Item;
      End If;
      Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
    
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 科室id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 校对标志, 卡类别id, 结算卡序号, 卡号,
         交易流水号, 交易说明, 结算号码, 结算性质, 关联交易id, 会话号)
      Values
        (n_预交id, 2, r_Balancedata.No, 1, r_Balancedata.病人id, r_Balancedata.科室id, r_Balancedata.主页id, Null, v_退支票,
         r_Balancedata.收款时间, r_Balancedata.操作员编号, r_Balancedata.操作员姓名, 退支票额_In, r_Balancedata.结帐id, r_Balancedata.缴款组id,
         校对标志_In, Null, Null, 卡号_In, 交易流水号_In, 交易说明_In, Null, 2, n_预交id, v_会话号);
    
      Update 病人预交记录 Set 冲预交 = 冲预交 - 退支票额_In Where 结帐id = 结帐id_In And 结算方式 Is Null;
      If Sql%NotFound Then
        v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！';
        Raise Err_Item;
      End If;
    
    End If;
  
    n_预交id := 预交id_In;
    --各个收费结算 :格式为:"结算方式|结算金额|结算号码|结算摘要|卡号||.."
    v_结算内容 := 结算方式_In || '||';
  
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
    
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
    
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
    
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      v_结算号码 := LTrim(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_卡号     := Null;
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      If Instr(v_当前结算, '|') > 0 Then
      
        v_结算摘要 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
        v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
        v_卡号     := LTrim(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
      Else
        v_结算摘要 := v_当前结算;
      End If;
    
      If Nvl(n_结算金额, 0) <> 0 Then
        If Nvl(n_预交id, 0) = 0 Then
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          n_关联交易id := n_预交id;
        End If;
      
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 科室id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 校对标志, 卡类别id, 结算卡序号, 卡号,
           交易流水号, 交易说明, 结算号码, 结算性质, 关联交易id, 会话号)
        Values
          (病人预交记录_Id.Nextval, 2, r_Balancedata.No, 1, r_Balancedata.病人id, r_Balancedata.科室id, r_Balancedata.主页id,
           v_结算摘要, v_结算方式, r_Balancedata.收款时间, r_Balancedata.操作员编号, r_Balancedata.操作员姓名, n_结算金额, r_Balancedata.结帐id,
           r_Balancedata.缴款组id, 校对标志_In, Null, Null, v_卡号, 交易流水号_In, 交易说明_In, v_结算号码, 2, n_关联交易id, v_会话号);
      
        Update 病人预交记录 Set 冲预交 = 冲预交 - n_结算金额 Where 结帐id = 结帐id_In And 结算方式 Is Null;
        If Sql%NotFound Then
          v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！';
          Raise Err_Item;
        End If;
      End If;
      n_预交id   := Null;
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  --5.三方卡结算交易
  If 操作类型_In = 1 Then
    v_当前结算 := 结算方式_In;
    v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
    v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
    n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
    v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
    v_结算号码 := LTrim(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
    v_结算摘要 := LTrim(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
  
    If Nvl(n_结算金额, 0) <> 0 Then
    
      If Nvl(n_预交id, 0) = 0 Then
        Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      End If;
    
      n_校对标志 := 校对标志_In;
      If Nvl(n_关联交易id, 0) = 0 Then
        n_关联交易id := n_预交id;
      Else
        Select Count(1), Max(a.是否未退), -1 * Nvl(Sum(a.金额), 0)
        Into n_Count, n_是否未退, n_退款金额
        From 三方退款信息 A, 病人预交记录 B
        Where a.记录id = b.Id And a.结帐id = r_Balancedata.结帐id And b.关联交易id = n_关联交易id;
        If n_Count > 1 Then
          --预交款多笔退款时，关联交易ID相同的需要合并；注意只要任一笔还没有退校对标志都是1
          If Nvl(n_是否未退, 0) = 1 Then
            n_校对标志 := 1;
          End If;
          n_结算金额 := n_退款金额;
        End If;
      End If;
    
      v_交易人员 := zl_UserName;
    
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 科室id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 校对标志, 卡类别id, 结算卡序号, 卡号,
         交易流水号, 交易说明, 结算号码, 结算性质, 关联交易id, 附加标志, 交易时间, 交易人员, 会话号)
      Values
        (n_预交id, 2, r_Balancedata.No, 1, r_Balancedata.病人id, r_Balancedata.科室id, r_Balancedata.主页id, v_结算摘要, v_结算方式,
         r_Balancedata.收款时间, r_Balancedata.操作员编号, r_Balancedata.操作员姓名, n_结算金额, r_Balancedata.结帐id, r_Balancedata.缴款组id,
         n_校对标志, 卡类别id_In, Null, 卡号_In, 交易流水号_In, 交易说明_In, v_结算号码, 2, n_关联交易id, 附加标志_In, Sysdate, v_交易人员, v_会话号);
    
      Update 病人预交记录 Set 冲预交 = 冲预交 - n_结算金额 Where 结帐id = 结帐id_In And 结算方式 Is Null;
    
      If Sql%NotFound Then
        v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！';
        Raise Err_Item;
      End If;
    
      --调用其他结算信息更新
      Zl_Custom_Balance_Update(n_预交id);
    
    End If;
  
  End If;

  --6.医保结算(调用此过程,采取平均分摊的方式分摊结算情况):这种情况医保结处后,必须全退
  If 操作类型_In = 2 Then
  
    --2.1检查是否已经存在医保结算数据,存在先删除
    n_结算金额 := 0;
    For v_医保 In (Select ID, 结算方式, 冲预交
                 From 病人预交记录 A
                 Where 结帐id = 结帐id_In And 卡类别id Is Null And Exists
                  (Select 1 From 结算方式 Where a.结算方式 = 名称 And 性质 In (3, 4)) And Mod(记录性质, 10) <> 1) Loop
      n_结算金额 := n_结算金额 + Nvl(v_医保.冲预交, 0);
      l_预交id.Extend;
      l_预交id(l_预交id.Count) := v_医保.Id;
    End Loop;
  
    If Nvl(n_结算金额, 0) <> 0 Then
      Update 病人预交记录
      Set 冲预交 = Nvl(冲预交, 0) + Nvl(n_结算金额, 0)
      Where 结帐id = 结帐id_In And 结算方式 Is Null;
      If Sql%NotFound Then
        v_Err_Msg := ' 结算信息错误, 可能因为并发原因造成结算信息错误, 请在 [ 收费结算窗口 ] 中重新收费！ ';
        Raise Err_Item;
      End If;
    End If;
  
    If l_预交id.Count <> 0 Then
      Forall I In 1 .. l_预交id.Count
        Delete 病人预交记录 Where ID = l_预交id(I);
    End If;
  
    If 结算方式_In Is Not Null Then
      v_结算内容 := 结算方式_In || '||';
    End If;
    n_预交id     := 预交id_In;
    n_关联交易id := 关联交易id_In;
  
    d_交易时间 := Sysdate;
    v_交易人员 := zl_UserName;
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      n_结算金额 := To_Number(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
    
      If n_结算金额 <> 0 Then
      
        If Nvl(n_预交id, 0) = 0 Then
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
        End If;
      
        If Nvl(n_关联交易id, 0) = 0 Then
          n_关联交易id := n_预交id;
        End If;
      
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 科室id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 校对标志, 结算性质, 关联交易id,
           交易时间, 交易人员, 会话号)
        Values
          (n_预交id, 2, r_Balancedata.No, 1, r_Balancedata.病人id, r_Balancedata.科室id, r_Balancedata.主页id, ' 保险结算 ', v_结算方式,
           r_Balancedata.收款时间, r_Balancedata.操作员编号, r_Balancedata.操作员姓名, n_结算金额, r_Balancedata.结帐id, r_Balancedata.缴款组id,
           校对标志_In, 2, n_关联交易id, d_交易时间, v_交易人员, v_会话号);
      
        --更新数据(结算方式为NULL的)
        Update 病人预交记录
        Set 冲预交 = 冲预交 - n_结算金额
        Where 结帐id = 结帐id_In And 结算方式 Is Null
        Returning Nvl(冲预交, 0) Into n_返回值;
        n_预交id := Null;
      End If;
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  
    --医保相关表的处理
    Update 保险结算明细 Set 标志 = 2 Where 结帐id = 结帐id_In;
  
  End If;

  --7-消费卡批量结算
  If 操作类型_In = 3 Then
    v_结算内容   := 结算方式_In || '||';
    n_预交id     := 预交id_In;
    n_关联交易id := 关联交易id_In;
  
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      --卡类别ID|卡号|消费卡ID|消费金额
      n_卡类别id := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      v_卡号     := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_消费卡id := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_结算金额 := To_Number(v_当前结算);
    
      Select Max(名称), Max(结算方式) Into v_名称, v_结算方式 From 消费卡类别目录 Where 编号 = n_卡类别id;
      If v_名称 Is Null Then
        v_Err_Msg := ' 未找到对应的结算卡接口, 本次刷卡消费失败 ! ';
        Raise Err_Item;
      End If;
      If v_结算方式 Is Null Then
        v_Err_Msg := v_名称 || ' 未设置对应的结算方式, 本次刷卡消费失败 ! ';
        Raise Err_Item;
      End If;
    
      If Nvl(n_结算金额, 0) <> 0 Then
        Update 病人预交记录
        Set 冲预交 = Nvl(冲预交, 0) + n_结算金额
        Where 记录性质 = 2 And 结帐id = r_Balancedata. 结帐id And 结算方式 = v_结算方式 And 结算卡序号 = n_卡类别id
        Returning ID Into n_预交id;
        If Sql%NotFound Then
        
          If Nvl(n_预交id, 0) = 0 Then
            Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          End If;
        
          If Nvl(n_关联交易id, 0) = 0 Then
            n_关联交易id := n_预交id;
          End If;
        
          Insert Into 病人预交记录
            (ID, 记录性质, NO, 记录状态, 病人id, 科室id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算卡序号, 校对标志, 结算性质,
             关联交易id, 会话号)
          Values
            (n_预交id, 2, r_Balancedata.No, 1, r_Balancedata. 病人id, r_Balancedata.科室id, r_Balancedata.主页id, Null, v_结算方式,
             r_Balancedata. 收款时间, r_Balancedata. 操作员编号, r_Balancedata. 操作员姓名, n_结算金额, r_Balancedata. 结帐id,
             r_Balancedata. 缴款组id, n_卡类别id, 校对标志_In, 2, n_关联交易id, v_会话号);
        End If;
      
        Zl_病人卡结算记录_支付(n_卡类别id, v_卡号, n_消费卡id, n_结算金额, n_预交id, r_Balancedata. 操作员编号, r_Balancedata. 操作员姓名,
                      r_Balancedata. 收款时间);
      
        --更新数据(结算方式为NULL的)
        Update 病人预交记录
        Set 冲预交 = 冲预交 - n_结算金额
        Where 结帐id = r_Balancedata. 结帐id And 结算方式 Is Null And Nvl(校对标志, 0) = 1
        Returning Nvl(冲预交, 0) Into n_返回值;
        n_预交id := Null;
      End If;
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  If Nvl(完成结算_In, 0) = 0 Then
    Close c_Balancedata;
    Return;
  End If;

  -----------------------------------------------------------------------------------------
  --完成收费,需要处理人员缴款余额,预交记录(结算方式=NULL)

  --清除三方未退款的记录，更新病人预交记录的校对标志
  --存在多笔预交款关联交易ID相同时，合并为了一条病人预交记录，只有全部成功时校对标志才更新为2，否则都是1
  Delete 三方退款信息 Where Nvl(是否未退, 0) = 1 And 结帐id = 结帐id_In;
  For c_记录 In (Select ID, 关联交易id, 冲预交, 卡类别id
               From 病人预交记录
               Where 记录性质 = 2 And 冲预交 < 0 And 结帐id = 结帐id_In And 卡类别id Is Not Null And 校对标志 = 1) Loop
  
    Select -1 * Nvl(Sum(a.金额), 0)
    Into n_结算金额
    From 三方退款信息 A, 病人预交记录 B
    Where a.记录id = b.Id And a.结帐id = 结帐id_In And b.关联交易id = c_记录.关联交易id And a.卡类别id = c_记录.卡类别id;
  
    Update 病人预交记录 Set 冲预交 = n_结算金额, 校对标志 = 2 Where ID = c_记录.Id;
    Update 病人预交记录 Set 冲预交 = 冲预交 + c_记录.冲预交 - n_结算金额 Where 结帐id = 结帐id_In And 结算方式 Is Null;
  End Loop;

  --1.删除结算方式为NULL的预交记录
  Delete 病人预交记录 Where 结帐id = 结帐id_In And 结算方式 Is Null And Nvl(冲预交, 0) = 0;
  If Sql%NotFound Then
    Select Count(*) Into n_Count From 病人预交记录 A Where 结帐id = 结帐id_In And 结算方式 Is Null;
    If n_Count <> 0 Then
      v_Err_Msg := ' 还存在未缴款的数据, 不能完成结算 ! ';
    Else
      v_Err_Msg := ' 结算信息错误, 可能因为并发原因造成结算信息错误, 请在 [ 收费结算窗口 ] 中重新收费! ';
    End If;
    Raise Err_Item;
  End If;

  Select Count(*), Max(b.名称)
  Into n_Count, v_Err_Msg
  From 病人预交记录 A, 医疗卡类别 B
  Where a.卡类别id = b.Id And a.结帐id = 结帐id_In And a.卡类别id Is Not Null And a.交易流水号 Is Null;
  --三方交易需要在最后检查其合法性
  If n_Count <> 0 Then
    v_Err_Msg := v_Err_Msg || '无交易流水号，但交易成功,请与系统管理员联系!';
    Raise Err_Item;
  End If;

  --结算金额为零时，增加一条金额为0的病人预交记录
  Select Count(*) Into n_Count From 病人预交记录 A Where 结帐id = 结帐id_In;

  If n_Count = 0 Then
    v_结算方式 := 缺省结算方式_In;
    If v_结算方式 Is Null Then
      Select Max(结算方式) Into v_结算方式 From 结算方式应用 Where 应用场合 = '结帐' And Nvl(缺省标志, 0) = 1;
      If v_结算方式 Is Null Then
        Select Nvl(Max(名称), ' 现金 ') Into v_结算方式 From 结算方式 Where Nvl(性质, 0) = 1;
      End If;
    End If;
    If Nvl(n_预交id, 0) = 0 Then
      Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
    End If;
  
    If Nvl(n_关联交易id, 0) = 0 Then
      n_关联交易id := n_预交id;
    End If;
  
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 科室id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 校对标志, 卡类别id, 结算卡序号, 卡号,
       交易流水号, 交易说明, 结算号码, 结算性质, 关联交易id, 会话号)
    Values
      (n_预交id, 2, Null, 1, r_Balancedata.病人id, r_Balancedata.科室id, Null, Null, v_结算方式, r_Balancedata.收款时间,
       r_Balancedata.操作员编号, r_Balancedata.操作员姓名, 0, r_Balancedata.结帐id, r_Balancedata.缴款组id, 2, Null, Null, Null, Null,
       交易说明_In, Null, 2, n_关联交易id, v_会话号);
  End If;
  n_是否电子票据 := 是否电子票据_In;
  If 是否电子票据_In Is Null Then
    Select Nvl(Max(险类), 0) Into n_险类 From 保险结算记录 Where 记录id = 结帐id_In And 性质 = 2;
    n_是否电子票据 := Zl_Fun_Isstarteinvoice(3, n_险类);
  End If;

  --2.处理缴款数据和找补数据及校对标志更新为0
  Update 病人预交记录
  Set 缴款 = 缴款_In, 找补 = 找补_In, 校对标志 = 0, 会话号 = Null, 是否电子票据 = n_是否电子票据
  Where 结帐id = 结帐id_In;

  --3.更新费用状态
  Update 病人结帐记录 Set 结算状态 = Null,是否电子票据 = n_是否电子票据 Where ID = 结帐id_In;

  --4.更新人员缴款数据
  For c_缴款 In (Select 结算方式, 操作员姓名, Sum(Nvl(a.冲预交, 0)) As 冲预交
               From 病人预交记录 A
               Where a.结帐id = 结帐id_In And Mod(a.记录性质, 10) <> 1
               Group By 结算方式, 操作员姓名) Loop
  
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) + Nvl(c_缴款.冲预交, 0)
    Where 收款员 = c_缴款.操作员姓名 And 性质 = 1 And 结算方式 = c_缴款.结算方式;
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额
        (收款员, 结算方式, 性质, 余额)
      Values
        (c_缴款.操作员姓名, c_缴款.结算方式, 1, Nvl(c_缴款.冲预交, 0));
    End If;
  End Loop;
  Close c_Balancedata;

  --5.费用记录与预交记录匹配检查
  Select Nvl(Sum(费用金额), 0), Nvl(Sum(结帐金额), 0)
  Into n_费用金额, n_结帐金额
  From (Select Nvl(Sum(结帐金额), 0) As 费用金额, 0 As 结帐金额
         From 门诊费用记录
         Where 结帐id = 结帐id_In
         Union All
         Select Nvl(Sum(结帐金额), 0) As 费用金额, 0 As 结帐金额
         From 住院费用记录
         Where 结帐id = 结帐id_In
         Union All
         Select 0 As 费用金额, Nvl(Sum(冲预交), 0) As 结帐金额
         From 病人预交记录
         Where 结帐id = 结帐id_In);

  If Nvl(n_费用金额, 0) <> Nvl(n_结帐金额, 0) Then
    v_Err_Msg := ' 结算信息与费用信息不匹配, 无法完成结算 ! ';
    Raise Err_Item;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人结帐结算_Modify;
/

Create Or Replace Procedure Zl_病人结帐作废_Modify
(
  操作类型_In      Number,
  病人id_In        病人结帐记录.病人id%Type,
  冲销id_In        病人预交记录.结帐id%Type,
  结算方式_In      Varchar2,
  卡类别id_In      病人预交记录.卡类别id%Type := Null,
  卡号_In          病人预交记录.卡号%Type := Null,
  交易流水号_In    病人预交记录.交易流水号%Type := Null,
  交易说明_In      病人预交记录.交易说明%Type := Null,
  缴款_In          病人预交记录.缴款%Type := Null,
  找补_In          病人预交记录.找补%Type := Null,
  误差金额_In      病人预交记录.冲预交%Type := Null,
  预交金额_In      病人预交记录.冲预交%Type := Null,
  操作员编号_In    病人预交记录.操作员编号%Type := Null,
  操作员姓名_In    病人预交记录.操作员姓名%Type := Null,
  收款时间_In      病人预交记录.收款时间%Type := Null,
  冲预交病人ids_In Varchar2 := Null,
  完成作废_In      Number := 0,
  校对标志_In      Number := 0,
  关联交易id_In    病人预交记录.Id%Type := Null,
  清除原交易_In    Number := 0,
  预交id_In        病人预交记录.Id%Type := Null,
  加入会话_In      Number := 1
) As
  ------------------------------------------------------------------------------------------------------------------------------
  --功能:收费结算时,修改结算的相关信息
  --操作类型_In:
  --   0-原样退:
  --       其他参数不处理
  --   1-普通退费方式:
  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
  --   2.三方卡退费结算:
  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
  --     ②卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
  --     关联交易ID_IN:
  --     清除原交易_In:1-表示在更新数据前，清除原来的交易信息(按结帐ID+关联交易ID来清除);0-表示不清除
  --   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
  --   4-消费卡结算:
  --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
  -- 预交金额_In:如果涉及预交款,则传入本次的退预交或冲预交金额 传入零<0时 表示退预交款;>0 时:表示冲预交款
  -- 误差金额_In:存在误差费时,传入
  -- 完成退费_In:0-未完成退费;1-异常完成退费;2-完成退费
  -- 冲预交病人ids_In:多个用逗分分离,冲预交时有效(冲预交类别与业务操作保持一致),主要是使用家属的预交款
  -- 校对标志_In:0-完成或不需要校对;1-需要校对;2-接口已经调用成功
  --加入会话_In：1-表示加放会话，0-表示不加入会话
  ------------------------------------------------------------------------------------------------------------------------------
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_结算内容 Varchar2(500);
  v_当前结算 Varchar2(500);
  v_卡号     病人医疗卡信息.卡号%Type;
  n_消费卡id 消费卡信息.Id%Type;
  n_卡类别id 病人预交记录.结算卡序号%Type;
  v_名称     Varchar2(100);
  n_预交id   病人预交记录.Id%Type;
  v_结算方式 病人预交记录.结算方式%Type;
  n_结算金额 病人预交记录.冲预交%Type;
  n_返回值   人员缴款余额.余额%Type;
  v_结算号码 病人预交记录.结算号码%Type;
  v_结算摘要 病人预交记录.摘要%Type;
  v_误差费   结算方式.名称%Type;
  n_缴款组id 病人预交记录.缴款组id%Type;
  n_异常作废 Number(3);
  n_校对标志 病人预交记录.校对标志%Type;

  d_交易时间 病人预交记录.交易时间%Type;
  v_交易人员 病人预交记录.交易人员%Type;
  n_Dec      Number; --金额小数位数

  n_Count        Number;
  n_Havenull     Number;
  l_预交id       t_NumList := t_NumList();
  n_原结帐id     病人预交记录.结帐id%Type;
  n_结帐id       病人预交记录.结帐id%Type;
  n_原预交id     病人预交记录.Id%Type;
  n_充值id       病人预交记录.Id%Type;
  v_会话号       病人预交记录.会话号%Type; --格式：SID+'_'+SERIAL#
  n_是否电子票据 病人预交记录.是否电子票据%Type;

  Cursor c_Balance_Record Is
    Select Max(NO) As NO, Max(m.病人id) As 病人id, Max(Nvl(收款时间_In, m.收费时间)) As 登记时间, Max(Nvl(操作员编号_In, m.操作员编号)) As 操作员编号,
           Max(Nvl(操作员姓名_In, m.操作员姓名)) As 操作员姓名, Sum(结帐金额) As 结算金额, Max(Nvl(n_缴款组id, m.缴款组id)) As 缴款组id,
           Max(结帐类型) As 结帐类型
    From 病人结帐记录 M
    Where ID = 冲销id_In;
  r_Balance_Record c_Balance_Record%RowType;

  Cursor c_Balance_Data Is
    Select 记录性质, NO, 记录状态, 病人id, 科室id, 主页id, 结算方式, Nvl(收款时间_In, 收款时间) As 收款时间, Nvl(操作员编号_In, 操作员编号) As 操作员编号,
           Nvl(操作员姓名_In, 操作员姓名) As 操作员姓名, 冲预交, 结帐id, Nvl(n_缴款组id, 缴款组id) As 缴款组id, 结算序号
    From 病人预交记录
    Where 结帐id = 冲销id_In And 结算方式 Is Null;
  r_Balance_Data c_Balance_Data%RowType;
  n_误差费       病人预交记录.冲预交%Type;

Begin

  If 操作员姓名_In Is Not Null Then
    n_缴款组id := Zl_Get组id(操作员姓名_In);
  End If;

  If Nvl(加入会话_In, 0) = 1 Then
    Select Max(Sid || '_' || Serial#) Into v_会话号 From V$session Where Audsid = Userenv('sessionid');
  End If;

  Open c_Balance_Record;
  Fetch c_Balance_Record
    Into r_Balance_Record;

  Open c_Balance_Data;
  Fetch c_Balance_Data
    Into r_Balance_Data;

  If r_Balance_Record.No Is Null Then
    v_Err_Msg := '未找到指定的结帐作废记录！';
    Raise Err_Item;
  End If;

  If 操作类型_In = 0 Then
    --原样作废
    Select Max(ID) Into n_原结帐id From 病人结帐记录 Where 记录状态 In (1, 3) And NO = r_Balance_Data.No;
    If n_原结帐id Is Null Then
      v_Err_Msg := '没有发现要作废的结帐单据,可能已经作废！';
      Raise Err_Item;
    End If;
  
    Delete 病人预交记录 Where 结帐id = 冲销id_In And 结算方式 Is Null;
    -- 退预交
    If Nvl(预交金额_In, 0) <> 0 Then
    
      Zl_病人结帐预交_Cancel(病人id_In, n_原结帐id, 冲销id_In, -1 * 预交金额_In, r_Balance_Data.操作员编号, r_Balance_Data.操作员姓名,
                       r_Balance_Data.收款时间, r_Balance_Record.缴款组id);
    Else
    
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 科室id, 主页id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
         结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算性质, 关联交易id, 会话号, 是否电子票据)
        Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 科室id, 主页id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号,
               r_Balance_Record.登记时间, r_Balance_Record.操作员姓名, r_Balance_Record.操作员编号, -1 * 冲预交, 冲销id_In,
               r_Balance_Record.缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 2, 2, 关联交易id, v_会话号, 是否电子票据
        From 病人预交记录
        Where 结帐id = n_原结帐id And 记录性质 In (1, 11) And Nvl(冲预交, 0) <> 0;
    
      For r_预交 In (Select 病人id, NO, 预交类别, Sum(冲预交) As 冲预交, Max(Decode(记录性质, 1, Decode(记录状态, 2, 0, ID), 0)) As 预交id
                   From 病人预交记录
                   Where 结帐id = 冲销id_In And Mod(记录性质, 10) = 1
                   Group By 病人id, NO, 预交类别) Loop
        --病人余额(预交)
        Update 病人余额
        Set 预交余额 = Nvl(预交余额, 0) - Nvl(r_预交.冲预交, 0) --注:新的结帐ID产生的是负数金额
        Where 病人id = r_预交.病人id And 类型 = Nvl(r_预交.预交类别, 2) And 性质 = 1
        Returning 预交余额 Into n_返回值;
      
        If Sql%RowCount = 0 Then
          Insert Into 病人余额
            (病人id, 性质, 类型, 预交余额, 费用余额)
          Values
            (r_预交.病人id, 1, Nvl(r_预交.预交类别, 2), -1 * r_预交.冲预交, 0);
          n_返回值 := -1 * r_预交.冲预交;
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete 病人余额 Where 性质 = 1 And 病人id = r_预交.病人id And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
        End If;
      
        --更新预交单据余额
        n_充值id := Nvl(r_预交.预交id, 0);
        If n_充值id = 0 Then
          Select Max(ID) Into n_充值id From 病人预交记录 Where NO = r_预交.No And 记录性质 = 1 And 记录状态 <> 2;
        End If;
        If n_充值id <> 0 Then
          Update 预交单据余额
          Set 预交余额 = Nvl(预交余额, 0) - Nvl(r_预交.冲预交, 0)
          Where 病人id = r_预交.病人id And 预交id = n_充值id
          Returning 预交余额 Into n_返回值;
        
          If Sql%RowCount = 0 Then
            Insert Into 预交单据余额
              (预交id, 病人id, 预交类别, 预交余额)
            Values
              (n_充值id, r_预交.病人id, Nvl(r_预交.预交类别, 2), Nvl(-1 * r_预交.冲预交, 0));
            n_返回值 := -1 * Nvl(r_预交.冲预交, 0);
          End If;
        
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 预交单据余额 Where 预交id = n_充值id And Nvl(预交余额, 0) = 0;
          End If;
        
        End If;
      End Loop;
    
    End If;
  
    --退其他
    Insert Into 病人预交记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交, 结帐id,
       缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算性质, 关联交易id, 交易时间, 交易人员, 会话号)
      Select 病人预交记录_Id.Nextval, a.No, 实际票号, 12, a.记录状态, a.病人id, a.主页id, a.科室id, Null, a.结算方式, a.结算号码, a.摘要, a.缴款单位,
             a.单位开户行, a.单位帐号, r_Balance_Record.登记时间, r_Balance_Record.操作员姓名, r_Balance_Record.操作员编号, -1 * 冲预交, 冲销id_In,
             r_Balance_Record.缴款组id, a.预交类别, a.卡类别id, a.结算卡序号, a.卡号, a.交易流水号, a.交易说明, a.合作单位, Nvl(校对标志_In, 0), 2,
             a.关联交易id,
             Decode(a.卡类别id, Null, Decode(Nvl(b.性质, 0), 3, r_Balance_Record.登记时间, 4, r_Balance_Record.登记时间, Null),
                     r_Balance_Record.登记时间),
             Decode(a.卡类别id, Null, Decode(Nvl(b.性质, 0), 3, r_Balance_Record.操作员姓名, 4, r_Balance_Record.操作员姓名, Null),
                     r_Balance_Record.操作员姓名), v_会话号
      From 病人预交记录 A, 结算方式 B
      Where 结帐id = n_原结帐id And a.结算方式 = b.名称(+) And Mod(记录性质, 10) <> 1 And Nvl(冲预交, 0) >= 0;
  
    Update 病人预交记录 Set 缴款 = 缴款_In, 找补 = 找补_In, 校对标志 = 0 Where 结帐id = 冲销id_In;
  
    Select Count(1)
    Into n_异常作废
    From 病人结帐记录
    Where NO = r_Balance_Data.No And 记录状态 = 3 And 结算状态 = 1 And Rownum < 2;
  
    If Nvl(校对标志_In, 0) = 0 And Nvl(n_异常作废, 0) <> 0 Then
      Update 病人结帐记录
      Set 结算状态 = Decode(n_异常作废, 1, 2, Null)
      Where NO = r_Balance_Data.No And 结算状态 Is Not Null;
    
      Update 病人预交记录 Set 会话号 = Null Where 结帐id = 冲销id_In And 会话号 Is Not Null;
    End If;
    Close c_Balance_Record;
  
    --需要调用三方卡其他信息更新过程
    For c_三方结算 In (Select ID From 病人预交记录 Where 结帐id = 冲销id_In And 卡类别id Is Not Null Order By 卡类别id) Loop
      --调用其他结算信息更新
      Zl_Custom_Balance_Update(c_三方结算.Id);
    End Loop;
  
    Return;
  End If;

  --0.正式结算
  Select Count(1), Max(Decode(结算方式, Null, 1, 0))
  Into n_Count, n_Havenull
  From 病人预交记录
  Where 结帐id = 冲销id_In;

  --金额小数位数
  Select zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  --1.增加结算方式为空的结算数据
  n_误差费 := 误差金额_In;

  If Nvl(n_Havenull, 0) = 0 Then
    n_Count := 0;
    Select Sum(结帐金额)
    Into n_结算金额
    From (Select Sum(结帐金额) As 结帐金额
           From 门诊费用记录
           Where 结帐id = 冲销id_In
           Union All
           Select Sum(结帐金额)
           From 住院费用记录
           Where 结帐id = 冲销id_In);
  
    Begin
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 校对标志, 结算性质, 会话号)
      Values
        (病人预交记录_Id.Nextval, 12, r_Balance_Record.No, 1, r_Balance_Record.病人id, Null, r_Balance_Record.登记时间,
         r_Balance_Record.操作员编号, r_Balance_Record.操作员姓名, n_结算金额, 冲销id_In, r_Balance_Record.缴款组id, 2, 2, v_会话号);
    Exception
      When Others Then
        n_Count := -1;
    End;
  
    If n_Count = -1 Then
      v_Err_Msg := '未找到指定的结帐作废数据,退费操作失败！';
      Raise Err_Item;
    End If;
  End If;

  --处理误差费
  If Nvl(n_误差费, 0) <> 0 Then
    Select Nvl(Max(名称), '误差费') Into v_误差费 From 结算方式 Where Nvl(性质, 0) = 9;
  
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 校对标志, 结算性质, 会话号)
    Values
      (病人预交记录_Id.Nextval, 12, r_Balance_Record.No, 1, r_Balance_Record.病人id, v_误差费, r_Balance_Record.登记时间,
       r_Balance_Record.操作员编号, r_Balance_Record.操作员姓名, n_误差费, 冲销id_In, r_Balance_Record.缴款组id, 2, 2, v_会话号);
  
    --更新数据(结算方式为NULL的)
    Update 病人预交记录
    Set 冲预交 = 冲预交 - Nvl(n_误差费, 0)
    Where 结帐id = r_Balance_Data.结帐id And 结算方式 Is Null
    Returning Nvl(冲预交, 0) Into n_返回值;
  
  End If;

  If Nvl(预交金额_In, 0) <> 0 Then
    Select Max(ID) Into n_原结帐id From 病人结帐记录 Where 记录状态 In (1, 3) And NO = r_Balance_Data.No;
    If n_原结帐id Is Null Then
      v_Err_Msg := '没有发现要作废的结帐单据,可能已经作废！';
      Raise Err_Item;
    End If;
  
    Zl_病人结帐预交_Cancel(病人id_In, n_原结帐id, 冲销id_In, -1 * 预交金额_In, r_Balance_Data.操作员编号, r_Balance_Data.操作员姓名,
                     r_Balance_Data.收款时间, r_Balance_Record.缴款组id);
  End If;

  If 操作类型_In = 1 Then
    --   1-普通退费方式:
    --各个收费结算 :格式为:"结算方式|结算金额|结算号码|结算摘要||.."
    v_结算内容 := 结算方式_In || '||';
    n_预交id   := 预交id_In;
  
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
    
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
    
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
    
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      v_结算号码 := LTrim(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_结算摘要 := LTrim(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
    
      --不判断“结算金额”是否为零，有可能已经退完，但这时结算方式为空的重结和冲销记录的冲预交之和为零
      If v_结算方式 Is Not Null Then
        --If Nvl(n_结算金额, 0) <> 0 Then
        n_结算金额 := Nvl(n_结算金额, 0);
        If Nvl(n_结算金额, 0) <> 0 Then
          If Nvl(n_预交id, 0) = 0 Then
            Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          End If;
        
          Insert Into 病人预交记录
            (ID, 记录性质, NO, 记录状态, 病人id, 科室id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 校对标志, 卡类别id, 结算卡序号,
             卡号, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号)
          Values
            (n_预交id, 12, r_Balance_Data.No, 1, r_Balance_Data.病人id, r_Balance_Data.科室id, r_Balance_Data.主页id, v_结算摘要,
             v_结算方式, r_Balance_Data.收款时间, r_Balance_Data.操作员编号, r_Balance_Data.操作员姓名, n_结算金额, r_Balance_Data.结帐id,
             r_Balance_Data.缴款组id, 2, Null, Null, 卡号_In, 交易流水号_In, 交易说明_In, Null, 2, v_会话号);
        
          Update 病人预交记录 Set 冲预交 = 冲预交 - n_结算金额 Where 结帐id = 冲销id_In And 结算方式 Is Null;
          If Sql%NotFound Then
            v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！';
            Raise Err_Item;
          End If;
          n_预交id := Null;
        End If;
      End If;
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  If 操作类型_In = 2 Then
  
    If Nvl(清除原交易_In, 0) = 1 And Nvl(关联交易id_In, 0) <> 0 Then
      --还原结算方式为空的结算金额
      --先锁表，以免并发操作
      Update 病人预交记录
      Set 冲预交 = 冲预交
      Where 结帐id = 冲销id_In And 关联交易id = 关联交易id_In And Mod(记录性质, 10) <> 1;
    
      Select Sum(冲预交)
      Into n_结算金额
      From 病人预交记录
      Where 结帐id = 冲销id_In And 关联交易id = 关联交易id_In And Mod(记录性质, 10) <> 1;
    
      Update 病人预交记录
      Set 冲预交 = Nvl(冲预交, 0) + Nvl(n_结算金额, 0)
      Where 结帐id = 冲销id_In And 结算方式 Is Null;
    
      Delete 病人预交记录 Where 结帐id = 冲销id_In And 关联交易id = 关联交易id_In And Mod(记录性质, 10) <> 1;
    End If;
  
    d_交易时间 := Sysdate;
    v_交易人员 := zl_UserName;
    --   2.三方卡退费结算:
    v_当前结算 := 结算方式_In;
    v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
    v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
    n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
    v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
    v_结算号码 := LTrim(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
    v_结算摘要 := LTrim(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
  
    If Nvl(n_结算金额, 0) <> 0 Then
      n_预交id := 预交id_In;
      If Nvl(n_预交id, 0) = 0 Then
        Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      End If;
    
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 科室id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 校对标志, 卡类别id, 结算卡序号, 卡号,
         交易流水号, 交易说明, 结算号码, 结算性质, 关联交易id, 交易时间, 交易人员, 会话号)
      Values
        (n_预交id, 12, r_Balance_Data.No, 1, r_Balance_Data.病人id, r_Balance_Data.科室id, r_Balance_Data.主页id, v_结算摘要,
         v_结算方式, r_Balance_Data.收款时间, r_Balance_Data.操作员编号, r_Balance_Data.操作员姓名, n_结算金额, r_Balance_Data.结帐id,
         r_Balance_Data.缴款组id, 校对标志_In, 卡类别id_In, Null, 卡号_In, 交易流水号_In, 交易说明_In, v_结算号码, 2,
         Decode(Nvl(关联交易id_In, 0), 0, n_预交id, 关联交易id_In), d_交易时间, v_交易人员, v_会话号);
    
      Update 病人预交记录 Set 冲预交 = 冲预交 - n_结算金额 Where 结帐id = 冲销id_In And 结算方式 Is Null;
      If Sql%NotFound Then
        v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！';
        Raise Err_Item;
      End If;
    
      --调用其他结算信息更新
      Zl_Custom_Balance_Update(n_预交id);
    End If;
  End If;

  If 操作类型_In = 3 Then
    --   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
    --3.1检查是否已经存在医保结算数据,存在先删除
    n_结算金额 := 0;
  
    If 校对标志_In = 0 Then
      n_校对标志 := 2;
    Else
      n_校对标志 := 1;
    End If;
  
    For v_医保 In (Select ID, 结算方式, 冲预交
                 From 病人预交记录 A
                 Where 结帐id = 冲销id_In And 卡类别id Is Null And Exists
                  (Select 1 From 结算方式 Where a.结算方式 = 名称 And 性质 In (3, 4)) And Mod(记录性质, 10) <> 1) Loop
      n_结算金额 := n_结算金额 + Nvl(v_医保.冲预交, 0);
      l_预交id.Extend;
      l_预交id(l_预交id.Count) := v_医保.Id;
    End Loop;
  
    If Nvl(n_结算金额, 0) <> 0 Then
      Update 病人预交记录
      Set 冲预交 = Nvl(冲预交, 0) + Nvl(n_结算金额, 0)
      Where 结帐id = 冲销id_In And 结算方式 Is Null;
      If Sql%NotFound Then
        v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！';
        Raise Err_Item;
      End If;
    End If;
    If l_预交id.Count <> 0 Then
      Forall I In 1 .. l_预交id.Count
        Delete 病人预交记录 Where ID = l_预交id(I);
    End If;
  
    If 结算方式_In Is Not Null Then
      v_结算内容 := 结算方式_In || '||';
    End If;
    d_交易时间 := Sysdate;
    v_交易人员 := zl_UserName;
    n_预交id   := 预交id_In;
  
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      n_结算金额 := To_Number(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
    
      For c_结算信息 In (Select 记录性质, NO, 记录状态, 病人id, 科室id, 主页id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 关联交易id
                     From 病人预交记录
                     Where 结帐id = 冲销id_In And 结算方式 Is Null) Loop
        If Nvl(n_预交id, 0) = 0 Then
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
        End If;
      
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 科室id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 校对标志, 结算性质, 交易时间, 交易人员,
           关联交易id, 会话号)
        Values
          (n_预交id, 12, c_结算信息.No, 1, c_结算信息.病人id, c_结算信息.科室id, c_结算信息.主页id, '保险结算', v_结算方式, c_结算信息.收款时间, c_结算信息.操作员编号,
           c_结算信息.操作员姓名, n_结算金额, c_结算信息.结帐id, c_结算信息.缴款组id, n_校对标志, 2, d_交易时间, v_交易人员,
           Decode(Nvl(关联交易id_In, 0), 0, n_预交id, 关联交易id_In), v_会话号);
        n_预交id := Null;
      End Loop;
    
      --更新数据(结算方式为NULL的)
      Update 病人预交记录
      Set 冲预交 = 冲预交 - n_结算金额
      Where 结帐id = 冲销id_In And 结算方式 Is Null
      Returning Nvl(冲预交, 0) Into n_返回值;
    
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  --4-消费卡批量结算
  If 操作类型_In = 4 Then
    v_结算内容 := 结算方式_In || '||';
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      --卡类别ID|卡号|消费卡ID|消费金额
      n_卡类别id := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      v_卡号     := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_消费卡id := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_结算金额 := To_Number(v_当前结算);
    
      Begin
        Select 名称, 结算方式 Into v_名称, v_结算方式 From 消费卡类别目录 Where 编号 = n_卡类别id;
      Exception
        When Others Then
          v_名称 := Null;
      End;
      If v_名称 Is Null Then
        v_Err_Msg := '未找到对应的结算卡接口,本次刷卡消费失败!';
        Raise Err_Item;
      End If;
      If v_结算方式 Is Null Then
        v_Err_Msg := v_名称 || '未设置对应的结算方式,本次刷卡消费失败!';
        Raise Err_Item;
      End If;
      If Nvl(n_结算金额, 0) <> 0 Then
        n_结帐id := 冲销id_In;
      
        Update 病人预交记录
        Set 冲预交 = Nvl(冲预交, 0) + n_结算金额
        Where 结帐id = n_结帐id And 结算方式 = v_结算方式 And 结算卡序号 = n_卡类别id
        Returning ID Into n_预交id;
        If Sql%NotFound Then
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (ID, 记录性质, NO, 记录状态, 病人id, 科室id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算卡序号, 校对标志, 结算性质,
             会话号)
          Values
            (n_预交id, 12, r_Balance_Data.No, 1, r_Balance_Data. 病人id, r_Balance_Data.科室id, r_Balance_Data.主页id, Null,
             v_结算方式, r_Balance_Data. 收款时间, r_Balance_Data. 操作员编号, r_Balance_Data. 操作员姓名, n_结算金额, n_结帐id,
             r_Balance_Data. 缴款组id, n_卡类别id, 2, 2, v_会话号);
        End If;
      
        Begin
          Select b.Id
          Into n_原预交id
          From 病人结帐记录 A, 病人预交记录 B
          Where a.Id = b.结帐id And a.记录状态 In (1, 3) And a.No = r_Balance_Data.No And b.结算卡序号 = n_卡类别id;
        Exception
          When Others Then
            Begin
              v_Err_Msg := '没有发现' || v_名称 || '的原结算数据！';
              Raise Err_Item;
            End;
        End;
      
        --插入卡结算记录
        Zl_病人卡结算记录_退款(n_卡类别id, v_卡号, n_消费卡id, -1 * n_结算金额, n_原预交id, n_预交id, r_Balance_Data. 操作员编号,
                      r_Balance_Data. 操作员姓名, r_Balance_Data. 收款时间);
      
        --更新数据(结算方式为NULL的)
        Update 病人预交记录
        Set 冲预交 = 冲预交 - n_结算金额
        Where 结帐id = n_结帐id And 结算方式 Is Null
        Returning Nvl(冲预交, 0) Into n_返回值;
      End If;
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  If Nvl(完成作废_In, 0) = 0 Then
    Return;
  End If;

  -----------------------------------------------------------------------------------------
  --完成收费,需要处理人员缴款余额,预交记录(结算方式=NULL)
  If Nvl(完成作废_In, 0) = 1 Then
  
    --异常完成退费
    Update 病人预交记录 Set 缴款 = 缴款_In, 找补 = 找补_In, 校对标志 = 0, 会话号 = Null Where 结帐id = 冲销id_In;
  
    Update 病人结帐记录
    Set 结算状态 = 2
    Where ID In (Select ID From 病人结帐记录 A Where NO = r_Balance_Data.No) And 结算状态 Is Not Null;
  
    Return;
  End If;

  --1.删除结算方式为NULL的预交记录
  Delete 病人预交记录 Where 结帐id = 冲销id_In And 结算方式 Is Null And Nvl(冲预交, 0) = 0;
  If Sql%NotFound Then
    Select Count(*) Into n_Count From 病人预交记录 A Where 结帐id = 冲销id_In And 结算方式 Is Null;
    If n_Count <> 0 Then
      v_Err_Msg := '还存在未退款的数据,不能完成结帐作废操作!';
    Else
      v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[结帐窗口]中重新作废！!';
    End If;
    Raise Err_Item;
  End If;

  --结算金额为零时，增加一条金额为0的病人预交记录
  Select Count(*) Into n_Count From 病人预交记录 A Where 结帐id = 冲销id_In;

  If n_Count = 0 Then
    If v_结算方式 Is Null Then
      Begin
        Select 结算方式 Into v_结算方式 From 结算方式应用 Where 应用场合 = '结帐' And Nvl(缺省标志, 0) = 1;
      Exception
        When Others Then
          v_结算方式 := Null;
      End;
      If v_结算方式 Is Null Then
        Begin
          Select 名称 Into v_结算方式 From 结算方式 Where Nvl(性质, 0) = 1;
        Exception
          When Others Then
            v_结算方式 := '现金';
        End;
      End If;
    End If;
  
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 科室id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 校对标志, 卡类别id, 结算卡序号, 卡号,
       交易流水号, 交易说明, 结算号码, 结算性质)
    Values
      (病人预交记录_Id.Nextval, 12, r_Balance_Data.No, 1, r_Balance_Data.病人id, r_Balance_Data.科室id, r_Balance_Data.主页id, Null,
       v_结算方式, r_Balance_Data.收款时间, r_Balance_Data.操作员编号, r_Balance_Data.操作员姓名, 0, r_Balance_Data.结帐id,
       r_Balance_Data.缴款组id, 2, Null, Null, Null, Null, 交易说明_In, Null, 2);
  End If;

  --更新电子票据
  Select Max(是否电子票据)
  Into n_是否电子票据
  From 病人预交记录
  Where 结帐id In (Select ID From 病人结帐记录 Where 记录状态 In (1, 3) And NO = r_Balance_Data.No);

  --2.处理缴款数据和找补数据及校对标志更新为0
  Update 病人预交记录
  Set 缴款 = 缴款_In, 找补 = 找补_In, 校对标志 = 0, 会话号 = Null, 是否电子票据 = n_是否电子票据
  Where 结帐id = 冲销id_In;

  --3.更新费用状态
  Update 病人结帐记录 Set 结算状态 = Null, 是否电子票据 = n_是否电子票据 Where ID = 冲销id_In;

  --4.票据回收
  -- 可能存在合约单位按病人打印, 所以存在多张票据
  For c_票据 In (Select ID As 打印id
               From (Select b.Id
                      From 票据使用明细 A, 票据打印内容 B
                      Where a.打印id = b.Id And a.性质 = 1 And a.原因 In (1, 3) And b.数据性质 = 3 And b.No = r_Balance_Data.No
                      Order By a.使用时间 Desc)
               Where Rownum < 2) Loop
  
    --作废收回票据(可能以前没有使用票据,无法收回)
    If c_票据.打印id Is Not Null Then
      Insert Into 票据使用明细
        (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
        Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, r_Balance_Data.收款时间, r_Balance_Data.操作员姓名, 票据金额
        From 票据使用明细 A
        Where 打印id = c_票据.打印id And 票种 In (1, 3) And 性质 = 1 And Not Exists
         (Select 1 From 票据使用明细 B Where a.号码 = b.号码 And 打印id = c_票据.打印id And 性质 = 2);
    End If;
  
  End Loop;

  --5.结算总额要与费用信息保持一致
  Select Sum(冲预交), Sum(结帐金额)
  Into n_返回值, n_结算金额
  From (Select Sum(冲预交) As 冲预交, 0 As 结帐金额
         From 病人预交记录
         Where 结帐id = 冲销id_In
         Union All
         Select 0, Sum(结帐金额)
         From 门诊费用记录
         Where 结帐id = 冲销id_In
         Union All
         Select 0, Sum(结帐金额) As 结帐金额
         From 住院费用记录
         Where 结帐id = 冲销id_In);

  If Nvl(n_返回值, 0) <> Nvl(n_结算金额, 0) Then
    v_Err_Msg := '结算总额与费用总额不一致,不能进行作废操作，请与系统管理员联系!';
    Raise Err_Item;
  End If;

  --5.更新人员缴款数据
  For c_缴款 In (Select 结算方式, 操作员姓名, Sum(Nvl(a.冲预交, 0)) As 冲预交
               From 病人预交记录 A
               Where a.结帐id = 冲销id_In And Mod(a.记录性质, 10) <> 1
               Group By 结算方式, 操作员姓名) Loop
  
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) + Nvl(c_缴款.冲预交, 0)
    Where 收款员 = c_缴款.操作员姓名 And 性质 = 1 And 结算方式 = c_缴款.结算方式;
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额
        (收款员, 结算方式, 性质, 余额)
      Values
        (c_缴款.操作员姓名, c_缴款.结算方式, 1, Nvl(c_缴款.冲预交, 0));
    End If;
  End Loop;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人结帐作废_Modify;
/
 

Create Or Replace Procedure Zl_医疗卡结算_Modify
(
  单据号_In       住院费用记录.No%Type,
  结帐id_In       住院费用记录.结帐id%Type,
  结算方式_In     病人预交记录.结算方式%Type := Null,
  结算金额_In     病人预交记录.冲预交%Type := 0,
  完成标志_In     Number := 0,
  卡类别id_In     病人预交记录.卡类别id%Type := Null,
  消费卡_In       Number := 0,
  卡号_In         病人预交记录.卡号%Type := Null,
  交易流水号_In   病人预交记录.交易流水号%Type := Null,
  交易说明_In     病人预交记录.交易说明%Type := Null,
  普通结算_In     Number := 0,
  结算号码_In     病人预交记录.结算号码%Type := Null,
  摘要_In         病人预交记录.摘要%Type := Null,
  校对标志_In     病人预交记录.校对标志%Type := 2,
  关联交易id_In   病人预交记录.关联交易id%Type := Null,
  是否电子票据_In 病人预交记录.是否电子票据%Type := Null
) As
  ----------------------------------------------------------------------------
  --参数:
  --是否电子票据_In:null-表示过程内部直接判断，非空表示直接以传入的为准
  --                （注：针对退费，该参数失效)
  ----------------------------------------------------------------------------

  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  n_结算金额     病人预交记录.冲预交%Type;
  n_冲预交       病人预交记录.冲预交%Type;
  n_预交id       病人预交记录.Id%Type;
  n_卡类别id     病人预交记录.卡类别id%Type;
  n_消费卡id     病人预交记录.卡类别id%Type;
  v_结算方式     病人预交记录.结算方式%Type;
  n_原预交id     病人预交记录.Id%Type;
  n_冲销金额     病人预交记录.冲预交%Type;
  n_收费作废     Number;
  n_退费         Number;
  n_记帐         Number;
  n_Count        Number;
  d_Date         病人预交记录.收款时间%Type;
  v_操作员编号   病人预交记录.操作员编号%Type;
  v_操作员姓名   病人预交记录.操作员姓名%Type;
  n_原结帐id     病人预交记录.结帐id%Type;
  n_是否电子票据 Number(2);
  n_险类         保险结算记录.险类%Type;
Begin

  Select Nvl(Max(记帐费用), 0), Max(结帐id), Max(Decode(记录状态, 3, 费用状态, 0)), Max(Decode(记录状态, 3, 1, 0))
  Into n_记帐, n_原结帐id, n_收费作废, n_退费
  From 住院费用记录
  Where NO = 单据号_In And 记录性质 = 5 And 记录状态 In (1, 3);

  If Nvl(结帐id_In, 0) = 0 Or n_记帐 = 1 Then
    Return;
  End If;

  If n_退费 = 1 Then
    n_结算金额 := -1 * Nvl(结算金额_In, 0);
  Else
    n_结算金额 := Nvl(结算金额_In, 0);
  End If;
  If 普通结算_In = 0 And Nvl(卡类别id_In, 0) <> 0 Then
    If Nvl(消费卡_In, 0) = 0 Then
      n_卡类别id := 卡类别id_In;
    Else
      n_消费卡id := 卡类别id_In;
    End If;
  End If;
  If 结算方式_In Is Null Then
    Update 病人预交记录
    Set 校对标志 = 校对标志_In
    Where 记录性质 = 5 And 结帐id = 结帐id_In And Rownum < 2 Return ID, 卡类别id, 收款时间, 操作员编号, 操作员姓名 Into n_预交id, n_卡类别id, d_Date,
     v_操作员编号, v_操作员姓名;
  Else
    Update 病人预交记录
    Set 结算方式 = 结算方式_In, 冲预交 = n_结算金额, 校对标志 = 校对标志_In, 卡类别id = n_卡类别id, 结算卡序号 = n_消费卡id, 卡号 = 卡号_In, 交易流水号 = 交易流水号_In,
        交易说明 = 交易说明_In, 结算号码 = 结算号码_In, 摘要 = Nvl(摘要_In, 摘要), 关联交易id = Nvl(关联交易id_In, ID)
    Where 记录性质 = 5 And 结帐id = 结帐id_In And Rownum < 2 Return ID, 卡类别id, 收款时间, 操作员编号, 操作员姓名 Into n_预交id, n_卡类别id, d_Date,
     v_操作员编号, v_操作员姓名;
  End If;

  --调用三方自主更新接口信息
  If 校对标志_In = 2 Then
    If Nvl(n_卡类别id, 0) <> 0 Then
      Zl_Custom_Balance_Update(n_预交id);
      Update 病人预交记录
      Set 交易时间 = 收款时间, 交易人员 = 操作员姓名
      Where 记录性质 = 5 And Nvl(卡类别id, 0) > 0 And 结帐id = 结帐id_In;
    End If;
  
    If Nvl(n_消费卡id, 0) <> 0 Then
      If n_退费 = 0 Then
        Zl_病人卡结算记录_支付(n_消费卡id, 卡号_In, 0, n_结算金额, n_预交id, v_操作员编号, v_操作员姓名, d_Date);
      Else
        Select Nvl(ID, 0), -1 * Nvl(冲预交, 0)
        Into n_原预交id, n_冲销金额
        From 病人预交记录
        Where NO = 单据号_In And 记录性质 = 5 And 记录状态 = 3 And 结算方式 = 结算方式_In And 结算卡序号 = 卡类别id_In;
        If n_原预交id = 0 Then
          v_Err_Msg := '未找到原结算记录！';
          Raise Err_Item;
        End If;
        If n_冲销金额 <> n_结算金额 Then
          v_Err_Msg := '消费卡退款金额不一致！';
          Raise Err_Item;
        End If;
        Zl_病人卡结算记录_退款(卡类别id_In, 卡号_In, 0, -1 * n_结算金额, n_原预交id, n_预交id, v_操作员编号, v_操作员姓名, d_Date);
      End If;
    End If;
  End If;

  If Nvl(完成标志_In, 0) = 0 Then
    Return;
  End If;

  --1.先检查金额是否一致
  Select Nvl(Sum(实收金额), 0) Into n_结算金额 From 住院费用记录 Where 结帐id = 结帐id_In;
  Select Nvl(Sum(冲预交), 0), Max(结算方式) Into n_冲预交, v_结算方式 From 病人预交记录 Where 结帐id = 结帐id_In;
  If n_结算金额 <> n_冲预交 Then
    v_Err_Msg := '卡费结算信息有误，实收金额(' || n_结算金额 || ')与结算金额(' || n_冲预交 || ')不一致，不能完成结算！';
    Raise Err_Item;
  End If;

  Select Count(1) Into n_Count From 病人预交记录 Where 结帐id = 结帐id_In And 卡类别id Is Not Null And 校对标志 = 1;
  If n_Count > 0 Then
    v_Err_Msg := '三方卡未调用接口支付，不能完成结算！';
    Raise Err_Item;
  End If;
  If v_结算方式 Is Null And Nvl(n_收费作废, 0) = 0 Then
    v_Err_Msg := '存在未指定的结算方式，不能完成结算！';
    Raise Err_Item;
  End If;
  If Nvl(n_退费, 0) = 1 Then
    Select Max(是否电子票据) Into n_是否电子票据 From 病人预交记录 Where 结帐id = n_原结帐id;
  Else
    n_是否电子票据 := 是否电子票据_In;
    If 是否电子票据_In Is Null Then
      Select Nvl(Max(险类), 0) Into n_险类 From 保险结算记录 Where 记录id = 结帐id_In And 性质 = 2;
      n_是否电子票据 := Zl_Fun_Isstarteinvoice(3, n_险类);
    End If;
  End If;

  --2.处理缴款数据和找补数据及校对标志更新为0
  Update 病人预交记录 Set 校对标志 = 0, 是否电子票据 = n_是否电子票据 Where 结帐id = 结帐id_In;

  If Nvl(n_收费作废, 0) = 1 Then
    Update 病人预交记录 Set 结算方式 = Null Where NO = 单据号_In And 记录性质 = 5 And 记录状态 = 3 And 校对标志 = 1;
  End If;

  --3.更新费用状态
  If Nvl(n_收费作废, 0) = 0 Then
    Update 住院费用记录 Set 费用状态 = 0 Where 结帐id = 结帐id_In;
  End If;

  --4.更新人员缴款数据,Not Exists中主要是针对三方卡发卡作废单原始单据结算成功了的，输入也调退费接口，但不能更新缴款余额
  For c_缴款 In (Select 结算方式, 操作员姓名, Sum(Nvl(a.冲预交, 0)) As 冲预交
               From 病人预交记录 A
               Where a.结帐id = 结帐id_In And Mod(a.记录性质, 10) <> 1 And 结算方式 Is Not Null And Not Exists
                (Select 1
                      From 病人预交记录 B
                      Where b.No = a.No And b.记录性质 = a.记录性质 And b.记录状态 = 3 And b.关联交易id = a.关联交易id And
                            Nvl(b.校对标志, 0) <> 0)
               Group By 结算方式, 操作员姓名) Loop
  
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) + Nvl(c_缴款.冲预交, 0)
    Where 收款员 = c_缴款.操作员姓名 And 性质 = 1 And 结算方式 = c_缴款.结算方式;
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额
        (收款员, 结算方式, 性质, 余额)
      Values
        (c_缴款.操作员姓名, c_缴款.结算方式, 1, Nvl(c_缴款.冲预交, 0));
    End If;
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_医疗卡结算_Modify;
/



Create Or Replace Procedure Zl_病人预交记录_冲预交
(
  病人id_In        病人预交记录.病人id%Type,
  结帐id_In        病人预交记录.结帐id%Type,
  冲预交_In        病人预交记录.冲预交%Type := Null,
  预交类别_In      病人预交记录.预交类别%Type,
  操作员编号_In    病人预交记录.操作员编号%Type,
  操作员姓名_In    病人预交记录.操作员姓名%Type,
  收款时间_In      病人预交记录.收款时间%Type,
  冲预交病人ids_In Varchar2 := Null,
  结算性质_In      Number := 3,
  费用余额检查_In  Number := 0,
  校对标志_In      病人预交记录.校对标志%Type := Null,
  不控制会话_In    病人预交记录.会话号%Type := 0,
  是否电子票据_In  病人预交记录.是否电子票据%Type := Null
) As
  --费用余额检查_In  0-预交余额检查时不减去费用余额，1-减去费用余额
  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  v_冲预交病人ids Varchar2(4000);
  n_返回值        人员缴款余额.余额%Type;
  n_预交金额      病人预交记录.冲预交%Type;
  n_冲预交        病人预交记录.冲预交%Type;
  n_会话号        病人预交记录.会话号%Type; --格式：SID+'_'+SERIAL#
  n_组id          财务缴款分组.Id%Type;
Begin
  v_冲预交病人ids := Nvl(冲预交病人ids_In, 病人id_In);
  n_组id          := Zl_Get组id(操作员姓名_In);

  If Nvl(不控制会话_In, 0) = 0 Then
    Select Max(Sid || '_' || Serial#) Into n_会话号 From V$session Where Audsid = Userenv('sessionid');
  End If;

  --预交款处理
  If Nvl(冲预交_In, 0) <> 0 Then
    If Nvl(病人id_In, 0) = 0 Then
      v_Err_Msg := '不能确定病人的病人ID,收费不能使用预交款结算,结算操作失败！';
      Raise Err_Item;
    End If;
  
    --病人余额检查
    Select Nvl(Sum(Nvl(预交余额, 0) - Decode(费用余额检查_In, 1, Nvl(费用余额, 0), 0)), 0)
    Into n_预交金额
    From 病人余额
    Where 病人id In (Select Column_Value From Table(f_Num2List(v_冲预交病人ids))) And Nvl(性质, 0) = 1 And 类型 = 预交类别_In;
  
    If Nvl(n_预交金额, 0) < Nvl(冲预交_In, 0) Then
      v_Err_Msg := '病人的当前预交余额为 ' || LTrim(To_Char(n_预交金额, '9999999990.00')) || '，小于本次支付金额 ' ||
                   LTrim(To_Char(冲预交_In, '9999999990.00')) || '，支付失败！';
      Raise Err_Item;
    End If;
  
    n_预交金额 := 冲预交_In;
  
    --先缴先用，且先用自己的
    --不包含结算方式为代收款项的预交款。
    For c_冲预交 In (Select a.No, b.预交余额 As 金额, Nvl(a.结帐id, 0) As 结帐id, a.病人id, a.记录状态, a.Id, a.收款时间, a.关联交易id
                  From 病人预交记录 A, 预交单据余额 B
                  Where a.Id = b.预交id And b.病人id In (Select Column_Value From Table(f_Num2List(v_冲预交病人ids))) And
                        Nvl(b.预交类别, 2) = 预交类别_In And a.结算方式 Not In (Select 名称 From 结算方式 Where 性质 = 5) And
                        Nvl(a.校对标志, 0) = 0
                  Order By Decode(病人id, Nvl(病人id_In, 0), 0, 1), a.收款时间) Loop
    
      If c_冲预交.金额 - n_预交金额 < 0 Then
        n_冲预交 := c_冲预交.金额;
      Else
        n_冲预交 := n_预交金额;
      End If;
    
      If c_冲预交.结帐id = 0 Then
        --第一次冲预交(填上结帐ID,金额为0)
        Update 病人预交记录
        Set 冲预交 = 0, 结帐id = 结帐id_In, 结算序号 = -1 * 结帐id_In, 结算性质 = 结算性质_In, 会话号 = n_会话号,
            是否电子票据 = Decode(Nvl(是否电子票据, 0), 1, 1, 是否电子票据_In)
        Where ID = c_冲预交.Id;
      End If;
      --冲上次剩余额
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
         结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 结算性质, 会话号, 关联交易id, 交易时间, 交易人员, 校对标志, 是否电子票据)
        Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间_In,
               操作员姓名_In, 操作员编号_In, n_冲预交, 结帐id_In, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, -1 * 结帐id_In,
               结算性质_In, n_会话号, c_冲预交.关联交易id, 收款时间_In, 操作员姓名_In, 校对标志_In, Decode(Nvl(是否电子票据, 0), 1, 1, 是否电子票据_In)
        From 病人预交记录
        Where NO = c_冲预交.No And 记录状态 = c_冲预交.记录状态 And 记录性质 In (1, 11) And Rownum = 1;
    
      --更新数据(结算方式为NULL的)
      Update 病人预交记录
      Set 冲预交 = 冲预交 - n_冲预交
      Where 结帐id = 结帐id_In And 结算方式 Is Null
      Returning Nvl(冲预交, 0) Into n_返回值;
    
      --更新病人预交余额
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) - n_冲预交
      Where 病人id = c_冲预交.病人id And 性质 = 1 And 类型 = 预交类别_In
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (c_冲预交.病人id, 预交类别_In, -1 * n_冲预交, 1);
        n_返回值 := -1 * n_冲预交;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 病人余额
        Where 病人id = c_冲预交.病人id And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
      End If;
    
      --更新预交单据余额
      Update 预交单据余额
      Set 预交余额 = Nvl(预交余额, 0) - n_冲预交
      Where 病人id = c_冲预交.病人id And 预交id = c_冲预交.Id
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 预交单据余额
          (预交id, 病人id, 预交类别, 预交余额)
        Values
          (c_冲预交.Id, c_冲预交.病人id, 预交类别_In, -1 * n_冲预交);
        n_返回值 := -1 * n_冲预交;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 预交单据余额 Where 预交id = c_冲预交.Id And Nvl(预交余额, 0) = 0;
      End If;
    
      --检查是否已经处理完
      If c_冲预交.金额 < n_预交金额 Then
        n_预交金额 := n_预交金额 - c_冲预交.金额;
      Else
        n_预交金额 := 0;
      End If;
      If n_预交金额 = 0 Then
        Exit;
      End If;
    
    End Loop;
    --检查金额是否足够
    If Abs(n_预交金额) > 0 Then
      v_Err_Msg := '病人的当前预交余额小于本次支付金额 ' || LTrim(To_Char(冲预交_In, '9999999990.00')) || '，不能继续操作！';
      Raise Err_Item;
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人预交记录_冲预交;
/

Create Or Replace Procedure Zl_病人预约挂号_接收
(
  No_In            In 病人挂号记录.No%Type,
  诊室_In          In 病人挂号记录.诊室%Type,
  结帐id_In        In 门诊费用记录.结帐id%Type := Null,
  卡类别id_In      In 病人预交记录.卡类别id%Type := Null,
  卡号_In          In 病人预交记录.卡号%Type := Null,
  交易流水号_In    In 病人预交记录.交易流水号%Type := Null,
  交易说明_In      In 病人预交记录.交易说明%Type := Null,
  接收时间_In      In 病人挂号记录.接收时间%Type := Null,
  冲预交病人ids_In In Varchar2 := Null,
  是否电子票据_In  病人预交记录.是否电子票据%Type := Null
  --该过程用于直接完成预约挂号接收、就诊；主要是医生站使用。
  -- 冲预交病人ids_In:多个用逗分分离,冲预交时有效(冲预交类别与业务操作保持一致),主要是使用家属的预交款
) As
  --号别信息
  Cursor c_Regist Is
    Select b.科室id, b.项目id, b.医生id, b.医生姓名, b.号码
    From 门诊费用记录 A, 挂号安排 B
    Where a.记录性质 = 4 And a.记录状态 = 1 And a.No = No_In And a.序号 = 1 And a.计算单位 = b.号码;
  r_Regist c_Regist%RowType;

  Cursor c_Registnew Is
    Select b.科室id, b.项目id, b.医生id, b.医生姓名, c.号码
    From 病人挂号记录 A, 临床出诊记录 B, 临床出诊号源 C
    Where a.记录性质 = 1 And a.记录状态 = 1 And a.No = No_In And a.出诊记录id = b.Id And b.号源id = c.Id;
  r_Registnew c_Registnew%RowType;

  v_划价no       门诊费用记录.No%Type;
  v_Temp         Varchar2(255);
  v_人员编号     门诊费用记录.操作员编号%Type;
  v_人员姓名     门诊费用记录.操作员姓名%Type;
  v_挂号生成队列 Varchar2(2);
  v_排队号码     排队叫号队列.排队号码%Type;
  v_预约方式     病人挂号记录.预约方式%Type;

  n_病人id   病人挂号记录.病人id%Type;
  n_门诊号   病人挂号记录.门诊号%Type;
  n_挂号金额 门诊费用记录.实收金额%Type;
  n_剩余金额 病人余额.预交余额%Type;
  n_结帐id   门诊费用记录.结帐id%Type;

  d_Date     Date;
  n_当天排队 Number(18);
  n_排队     Number(18);
  v_结算方式 病人预交记录.结算方式%Type;
  v_三方名称 医疗卡类别.名称%Type;

  v_Error Varchar2(255);
  Err_Custom Exception;
  n_组id         财务缴款分组.Id%Type;
  v_排队序号     排队叫号队列.排队序号%Type;
  n_结算模式     病人信息.结算模式%Type;
  v_付款方式     病人挂号记录.医疗付款方式%Type;
  n_出诊记录id   临床出诊记录.Id%Type;
  n_是否电子票据 Number(2);
  n_险类         保险结算记录.险类%Type;
Begin
  Begin
    Select a.病人id, a.标识号, Nvl(b.预交余额, 0) - Nvl(b.费用余额, 0) As 余额, Sum(a.实收金额) As n_挂号金额, Substr(a.结论, 1, 10)
    Into n_病人id, n_门诊号, n_剩余金额, n_挂号金额, v_预约方式
    From 门诊费用记录 A, 病人余额 B
    Where a.病人id = b.病人id(+) And b.性质(+) = 1 And b.类型(+) = 1 And a.No = No_In And a.记录性质 = 4 And a.记录状态 = 0
    Group By a.病人id, a.标识号, Nvl(b.预交余额, 0) - Nvl(b.费用余额, 0), a.结论;
  Exception
    When Others Then
      v_Error := '预约挂号信息不存在，可能该预约挂号已被接收。';
      Raise Err_Custom;
  End;
  If 接收时间_In Is Null Then
    Select Sysdate Into d_Date From Dual;
  Else
    d_Date := 接收时间_In;
  End If;

  Begin
    Select 出诊记录id Into n_出诊记录id From 病人挂号记录 Where NO = No_In;
  Exception
    When Others Then
      n_出诊记录id := Null;
  End;

  n_是否电子票据 := 是否电子票据_In;
  If 是否电子票据_In Is Null Then
    Select Nvl(Max(险类), 0) Into n_险类 From 保险结算记录 Where 记录id = 结帐id_In And 性质 = 1;
    n_是否电子票据 := Zl_Fun_Isstarteinvoice(4, n_险类);
  End If;

  --当前操作人员
  v_Temp     := Zl_Identity;
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_人员编号 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  n_组id     := Zl_Get组id(v_人员姓名);
  n_结算模式 := 0;
  If Nvl(n_病人id, 0) <> 0 Then
    Select Nvl(结算模式, 0) Into n_结算模式 From 病人信息 Where 病人id = n_病人id;
  End If;
  If n_结算模式 = 0 Then
    If Nvl(结帐id_In, 0) = 0 Then
      Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
    Else
      n_结帐id := 结帐id_In;
    End If;
  End If;

  --病人信息的产生
  If n_门诊号 Is Null Then
    Select To_Number(Nextno(3)) Into n_门诊号 From Dual;
  End If;

  If n_病人id Is Null Then
    Select To_Number(Nextno(1)) Into n_病人id From Dual;
    Insert Into 病人信息
      (病人id, 门诊号, 姓名, 性别, 年龄, 费别, 医疗付款方式, 登记时间)
      Select n_病人id, n_门诊号, a.姓名, a.性别, a.年龄, a.费别, b.名称, d_Date
      From 门诊费用记录 A, 医疗付款方式 B
      Where a.付款方式 = b.编码(+) And a.No = No_In And a.记录性质 = 4 And a.记录状态 = 0 And a.序号 = 1;
  End If;

  --更新病人信息，含就诊信息
  Update 病人信息 Set 就诊时间 = d_Date, 就诊状态 = 2, 就诊诊室 = 诊室_In Where 病人id = n_病人id;

  --更新门诊费用记录，含就诊信息
  Update 门诊费用记录
  Set 记录状态 = 1, 结帐id = Decode(n_结算模式, 1, Null, n_结帐id), 结帐金额 = Decode(n_结算模式, 1, Null, 实收金额), 发药窗口 = 诊室_In, 执行人 = v_人员姓名,
      执行状态 = 2, 执行时间 = d_Date, 病人id = Decode(病人id, Null, n_病人id, 病人id), 标识号 = Decode(标识号, Null, n_门诊号, 标识号),
      登记时间 = d_Date, 操作员编号 = v_人员编号, 操作员姓名 = v_人员姓名, 缴款组id = n_组id, 记帐费用 = Decode(n_结算模式, 1, 1, 0)
  Where NO = No_In And 记录性质 = 4 And 记录状态 = 0;

  Update 病人挂号记录
  Set 记录性质 = 1, 接收人 = v_人员姓名, 接收时间 = d_Date, 诊室 = 诊室_In, 执行人 = v_人员姓名, 执行时间 = d_Date, 执行状态 = 2,
      病人id = Decode(病人id, Null, n_病人id, 病人id), 门诊号 = Decode(门诊号, Null, n_门诊号, 门诊号)
  Where NO = No_In And 记录状态 = 1 And 记录性质 = 2;

  b_Message.Zlhis_Cis_008(n_病人id, No_In);

  If Sql%NotFound Then
    --产生病人挂号记录，含就诊信息
    Begin
      Select a.名称
      Into v_付款方式
      From 医疗付款方式 A, 门诊费用记录 B
      Where b.No = No_In And b.记录性质 = 4 And b.记录状态 = 1 And b.序号 = 1 And a.编码 = b.付款方式 And Rownum < 2;
      Insert Into 病人挂号记录
        (ID, NO, 病人id, 记录性质, 记录状态, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 发生时间, 登记时间, 操作员编号, 操作员姓名,
         预约, 预约方式, 接收时间, 接收人, 预约时间, 医疗付款方式, 出诊记录id, 挂号项目id, 费别)
        Select 病人挂号记录_Id.Nextval, No_In, 病人id, 1, 1, 标识号, 姓名, 性别, 年龄, 计算单位, 加班标志, 诊室_In, Null, 执行部门id, v_人员姓名, 2, d_Date,
               发生时间, 登记时间, 操作员编号, 操作员姓名, 1, Substr(结论, 1, 10) As 预约方式, d_Date, v_人员姓名, 发生时间, v_付款方式, n_出诊记录id, 收费细目id,
               费别
        
        From 门诊费用记录
        Where NO = No_In And 记录性质 = 4 And 记录状态 = 1 And 序号 = 1;
    Exception
      When Others Then
        v_Error := '该预约挂号已被接收。';
        Raise Err_Custom;
    End;
  End If;

  v_挂号生成队列 := zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113));
  If v_挂号生成队列 <> 0 Then
    For c_挂号 In (Select ID, 执行部门id, 姓名, 诊室_In As 诊室, 登记时间, 执行人 As 执行人, 病人id, 号别, 号序
                 From 病人挂号记录
                 Where NO = No_In And Rownum = 1) Loop
      Begin
        Select 1,
               Case
                 When 排队时间 < Trunc(Sysdate) Then
                  1
                 Else
                  0
               End
        Into n_排队, n_当天排队
        From 排队叫号队列
        Where 业务类型 = 0 And 业务id = c_挂号.Id And Rownum <= 1;
      Exception
        When Others Then
          n_排队 := 0;
      End;
    
      If n_排队 = 0 Then
        --新增排队
        --产生队列
        --.按”执行部门” 的方式生成队列
        v_排队号码 := Zlgetnextqueue(c_挂号.执行部门id, c_挂号.Id, c_挂号.号别 || '|' || Nvl(c_挂号.号序, 0));
        v_排队序号 := Zlgetsequencenum(0, c_挂号.Id, 0);
        --  队列名称_In , 业务类型_In, 业务id_In,科室id_In,排队号码_In,排队标记_In,患者姓名_In,病人ID_IN, 诊室_In, 医生姓名_In, 排队时间_In
        Zl_排队叫号队列_Insert(c_挂号.执行部门id, 0, c_挂号.Id, c_挂号.执行部门id, v_排队号码, Null, c_挂号.姓名, c_挂号.病人id, c_挂号.诊室, c_挂号.执行人,
                         Sysdate, v_预约方式, Null, v_排队序号);
      Elsif Nvl(n_当天排队, 0) = 1 Then
        --更新队列号
        v_排队号码 := Zlgetnextqueue(c_挂号.执行部门id, c_挂号.Id, c_挂号.号别 || '|' || Nvl(c_挂号.号序, 0));
        v_排队序号 := Zlgetsequencenum(0, c_挂号.Id, 1);
        --新队列名称_IN, 业务类型_In, 业务id_In , 科室id_In , 患者姓名_In , 诊室_In, 医生姓名_In ,排队号码_In
        Zl_排队叫号队列_Update(c_挂号.执行部门id, 0, c_挂号.Id, c_挂号.执行部门id, c_挂号.姓名, c_挂号.诊室, c_挂号.执行人, v_排队号码, v_排队序号);
      
      Else
        --新队列名称_IN, 业务类型_In, 业务id_In , 科室id_In , 患者姓名_In , 诊室_In, 医生姓名_In ,排队号码_In
        Zl_排队叫号队列_Update(c_挂号.执行部门id, 0, c_挂号.Id, c_挂号.执行部门id, c_挂号.姓名, c_挂号.诊室, c_挂号.执行人);
      End If;
      --接收后,变成弃号
      Update 排队叫号队列 Set 排队状态 = 2 Where 业务类型 = 0 And 业务id = c_挂号.Id;
    End Loop;
  End If;

  --挂号费用结算
  If Nvl(n_挂号金额, 0) <> 0 Then
  
    If Nvl(n_剩余金额, 0) >= Nvl(n_挂号金额, 0) And Nvl(卡类别id_In, 0) = 0 And n_结算模式 = 0 Then
      --冲预交方式结算
      Zl_病人预交记录_冲预交(n_病人id, n_结帐id, n_挂号金额, 1, v_人员编号, v_人员姓名, d_Date, 冲预交病人ids_In, 4, 1, Null, 0, n_是否电子票据);
    Elsif Nvl(卡类别id_In, 0) > 0 And n_结算模式 = 0 Then
    
      Begin
        Select 结算方式, 名称 Into v_结算方式, v_三方名称 From 医疗卡类别 Where ID = 卡类别id_In;
      Exception
        When Others Then
          v_三方名称 := Null;
      End;
      If v_三方名称 Is Null Then
        v_Error := '未找到三方接口,请在医疗卡类别中设置.';
        Raise Err_Custom;
      End If;
      If v_结算方式 Is Null Then
        v_Error := v_三方名称 || '未设置对应的结算方式,请在医疗卡类别中设置.';
        Raise Err_Custom;
      End If;
    
      --第三方接口支付
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交, 结帐id,
         缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 预交类别, 结算序号, 结算性质, 是否电子票据)
        Select 病人预交记录_Id.Nextval, NO, Null, 4, 1, 病人id, 病人科室id, Null, v_结算方式, Null, '医生站挂号接收', Null, Null, Null, 登记时间,
               操作员姓名, 操作员编号, n_挂号金额, 结帐id, 缴款组id, 卡类别id_In, Null, 卡号_In, 交易流水号_In, 交易说明_In, v_三方名称, Null, 结帐id, 4,
               n_是否电子票据
        From 门诊费用记录
        Where NO = No_In And 记录性质 = 4 And 记录状态 = 1 And 序号 = 1;
    
      Update 人员缴款余额
      Set 余额 = Nvl(余额, 0) + n_挂号金额
      Where 收款员 = v_人员姓名 And 性质 = 1 And 结算方式 = v_结算方式;
      If Sql%RowCount = 0 Then
        Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (v_人员姓名, v_结算方式, 1, n_挂号金额);
      End If;
    Else
      If n_结算模式 = 1 Then
        --记帐
        For c_费用 In (Select 实收金额, 病人科室id, 开单部门id, 执行部门id, 收入项目id
                     From 门诊费用记录
                     Where 记录性质 = 4 And 记录状态 = 1 And NO = No_In And Nvl(记帐费用, 0) = 1) Loop
          --病人余额
          Update 病人余额
          Set 费用余额 = Nvl(费用余额, 0) + Nvl(c_费用.实收金额, 0)
          Where 病人id = Nvl(n_病人id, 0) And 性质 = 1 And 类型 = 1;
        
          If Sql%RowCount = 0 Then
            Insert Into 病人余额
              (病人id, 性质, 类型, 费用余额, 预交余额)
            Values
              (n_病人id, 1, 1, Nvl(c_费用.实收金额, 0), 0);
          End If;
        
          --病人未结费用
          Update 病人未结费用
          Set 金额 = Nvl(金额, 0) + Nvl(c_费用.实收金额, 0)
          Where 病人id = n_病人id And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(病人科室id, 0) = Nvl(c_费用.病人科室id, 0) And
                Nvl(开单部门id, 0) = Nvl(c_费用.开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(c_费用.执行部门id, 0) And
                收入项目id + 0 = c_费用.收入项目id And 来源途径 + 0 = 1;
        
          If Sql%RowCount = 0 Then
            Insert Into 病人未结费用
              (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
            Values
              (n_病人id, Null, Null, c_费用.病人科室id, c_费用.开单部门id, c_费用.执行部门id, c_费用.收入项目id, 1, Nvl(c_费用.实收金额, 0));
          End If;
        End Loop;
      
      Else
      
        --生成划价单收费(允许的情况下)
        v_Temp := zl_GetSysParameter('挂号模式', 9000);
        If Nvl(v_Temp, '0') = '0' Then
          v_Error := '病人剩余款额' || To_Char(Nvl(n_剩余金额, 0), '0.00') || ' 不足挂号金额' || To_Char(Nvl(n_挂号金额, 0), '0.00') ||
                     '，不能完成预约接收。';
          Raise Err_Custom;
        End If;
      
        Select Nextno(13) Into v_划价no From Dual;
      
        Insert Into 门诊费用记录
          (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 门诊标志, 病人id, 标识号, 付款方式, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数,
           数次, 发药窗口, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 记帐费用, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 摘要,
           结论, 缴款组id)
          Select 病人费用记录_Id.Nextval, 1, v_划价no, 0, a.序号, a.从属父号, a.价格父号, a.门诊标志, a.病人id, a.标识号, a.付款方式, a.姓名, a.性别, a.年龄,
                 a.病人科室id, a.费别, a.收费类别, a.收费细目id, b.计算单位, a.付数, a.数次, Null, Null, Null, a.收入项目id, a.收据费目, a.标准单价,
                 a.应收金额, a.实收金额, 0, v_人员姓名, a.执行部门id, v_人员姓名, d_Date, d_Date, a.执行部门id, 0, '挂号:' || No_In, a.结论, n_组id
          From 门诊费用记录 A, 收费项目目录 B
          Where a.收费细目id = b.Id And a.No = No_In And a.记录性质 = 4 And a.记录状态 = 1;
      
        --挂号本身不收费
        Update 门诊费用记录
        Set 应收金额 = 0, 实收金额 = 0, 结帐金额 = 0
        Where NO = No_In And 记录性质 = 4 And 记录状态 = 1;
      End If;
    End If;
  End If;

  --病人挂号汇总(只处理一次,且单独收取病历费不处理)
  If n_出诊记录id Is Null Then
    Open c_Regist;
    Fetch c_Regist
      Into r_Regist;
    Update 病人挂号汇总
    Set 已挂数 = Nvl(已挂数, 0) + 1, 其中已接收 = Nvl(其中已接收, 0) + 1
    Where 日期 = Trunc(d_Date) And Nvl(科室id, 0) = Nvl(r_Regist.科室id, 0) And Nvl(项目id, 0) = Nvl(r_Regist.项目id, 0) And
          Nvl(医生姓名, '医生') = Nvl(r_Regist.医生姓名, '医生') And Nvl(医生id, 0) = Nvl(r_Regist.医生id, 0) And
          (号码 = r_Regist.号码 Or 号码 Is Null);
    If Sql%RowCount = 0 Then
      Insert Into 病人挂号汇总
        (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数)
      Values
        (Trunc(d_Date), r_Regist.科室id, r_Regist.项目id, r_Regist.医生姓名, r_Regist.医生id, r_Regist.号码, 1);
    End If;
    Close c_Regist;
  Else
    Open c_Registnew;
    Fetch c_Registnew
      Into r_Registnew;
    Update 临床出诊记录 Set 已挂数 = Nvl(已挂数, 0) + 1, 其中已接收 = Nvl(其中已接收, 0) + 1 Where ID = n_出诊记录id;
    Update 病人挂号汇总
    Set 已挂数 = Nvl(已挂数, 0) + 1, 其中已接收 = Nvl(其中已接收, 0) + 1
    Where 日期 = Trunc(d_Date) And Nvl(科室id, 0) = Nvl(r_Registnew.科室id, 0) And Nvl(项目id, 0) = Nvl(r_Registnew.项目id, 0) And
          Nvl(医生姓名, '医生') = Nvl(r_Registnew.医生姓名, '医生') And Nvl(医生id, 0) = Nvl(r_Registnew.医生id, 0) And
          (号码 = r_Registnew.号码 Or 号码 Is Null);
    If Sql%RowCount = 0 Then
      Insert Into 病人挂号汇总
        (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数)
      Values
        (Trunc(d_Date), r_Registnew.科室id, r_Registnew.项目id, r_Registnew.医生姓名, r_Registnew.医生id, r_Registnew.号码, 1);
    End If;
    Close c_Registnew;
  End If;

  --病人担保信息
  If n_病人id Is Not Null Then
    Update 病人信息
    Set 担保人 = Null, 担保额 = Null, 担保性质 = Null
    Where 病人id = n_病人id And Exists (Select 1
           From 病人担保记录
           Where 病人id = n_病人id And 主页id Is Not Null And
                 登记时间 = (Select Max(登记时间) From 病人担保记录 Where 病人id = n_病人id));
    If Sql%RowCount > 0 Then
      Update 病人担保记录
      Set 到期时间 = d_Date
      Where 病人id = n_病人id And 主页id Is Not Null And Nvl(到期时间, d_Date) > d_Date;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人预约挂号_接收;
/

Create Or Replace Procedure Zl_医疗卡记录_Delete
(
  单据号_In     住院费用记录.No%Type,
  操作员编号_In 住院费用记录.操作员编号%Type,
  操作员姓名_In 住院费用记录.操作员姓名%Type,
  退费类型_In   Integer := 0,
  退费方式_In   病人预交记录.结算方式%Type := Null,
  销账id_In     住院费用记录.结帐id%Type := 0
) As
  --退费类型_In:0-退卡不包含病历费；1-退卡包含病历费；2-单独退病历费
  Cursor c_Cardinfo Is
    Select a.Id As 费用id, Nvl(a.记帐费用, 0) As 记帐, a.结帐id, a.实际票号, a.病人id, Nvl(a.主页id, 0) As 主页id,
           Nvl(a.病人病区id, 0) As 病人病区id, Nvl(a.病人科室id, 0) As 病人科室id, Nvl(a.开单部门id, 0) As 开单部门id,
           Nvl(a.执行部门id, 0) As 执行部门id, a.收入项目id, a.实收金额, b.结算方式, b.冲预交, b.卡类别id, b.卡号, b.结算卡序号, b.结算序号, a.结论,
           b.Id As 预交id, a.摘要, a.费用状态
    From 住院费用记录 A, 病人预交记录 B
    Where a.记录性质 = 5 And a.记录状态 = 1 And a.No = 单据号_In And a.结帐id = b.结帐id(+) And a.附加标志 <> 8;
  r_Cardrow c_Cardinfo%RowType;

  Cursor c_Booksinfo Is
    Select a.Id As 费用id, Nvl(a.记帐费用, 0) As 记帐, a.结帐id, a.实际票号, a.病人id, Nvl(a.主页id, 0) As 主页id,
           Nvl(a.病人病区id, 0) As 病人病区id, Nvl(a.病人科室id, 0) As 病人科室id, Nvl(a.开单部门id, 0) As 开单部门id,
           Nvl(a.执行部门id, 0) As 执行部门id, a.收入项目id, a.实收金额, b.结算方式, b.冲预交, b.卡类别id, b.卡号, b.结算卡序号, b.结算序号, a.结论,
           b.Id As 预交id, a.摘要, a.费用状态
    From 住院费用记录 A, 病人预交记录 B
    Where a.记录性质 = 5 And a.记录状态 = 1 And a.No = 单据号_In And a.结帐id = b.结帐id(+) And a.附加标志 = 8;
  r_Booksrow c_Booksinfo%RowType;

  v_费用id     住院费用记录.Id%Type;
  v_结帐id     住院费用记录.结帐id%Type;
  n_返回值     病人余额.费用余额%Type;
  n_卡类别id   Number(18);
  v_划价状态   门诊费用记录.记录状态%Type;
  n_病历id     住院费用记录.Id%Type;
  n_退费金额   病人预交记录.冲预交%Type;
  v_Date       Date;
  n_记账       Number(1);
  n_病人id     住院费用记录.病人id%Type;
  n_主页id     住院费用记录.主页id%Type;
  n_病人科室id 住院费用记录.病人科室id%Type;
  n_病人病区id 住院费用记录.病人病区id%Type;
  n_开单部门id 住院费用记录.开单部门id%Type;
  n_执行部门id 住院费用记录.执行部门id%Type;
  n_收入项目id 住院费用记录.收入项目id%Type;
  n_二次退费   Number; --记录是否是此单据的第二次退费
  n_新预交id   病人预交记录.Id%Type;
  n_校对标志   病人预交记录.校对标志%Type;
  v_Err_Msg    Varchar2(200);
  Err_Item Exception;
  n_组id 财务缴款分组.Id%Type;

Begin
  Begin
    Select 1 Into n_二次退费 From 住院费用记录 Where 记录性质 = 5 And NO = 单据号_In And 记录状态 = 3 And Rownum < 2;
  Exception
    When Others Then
      n_二次退费 := 0;
  End;

  If 退费类型_In <> 2 Then
    Open c_Cardinfo;
    Fetch c_Cardinfo
      Into r_Cardrow;
    n_组id := Zl_Get组id(操作员姓名_In);
  
    --首先判断要退卡的记录是否存在
    If c_Cardinfo%RowCount = 0 Then
      Close c_Cardinfo;
      v_Err_Msg := '[ZLSOFT]没有发现要退卡的记录,该记录可能已经退除！[ZLSOFT]';
      Raise Err_Item;
    Else
      Select Sysdate Into v_Date From Dual;
      Select 病人费用记录_Id.Nextval Into v_费用id From Dual;
    
      If r_Cardrow.记帐 = 0 Then
        If Nvl(销账id_In, 0) = 0 Then
          Select 病人结帐记录_Id.Nextval Into v_结帐id From Dual;
        Else
          v_结帐id := 销账id_In;
        End If;
        n_校对标志 := 1;
      Else
        n_校对标志 := 0;
      End If;
    
      --退除就诊卡费用记录
      Insert Into 住院费用记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 病人id, 主页id, 病人病区id, 病人科室id, 门诊标志, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位, 付数, 数次,
         加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 操作员编号, 操作员姓名, 发生时间, 登记时间, 结帐id,
         结帐金额, 缴款组id, 结论, 摘要, 费用状态)
        Select v_费用id, NO, 实际票号, 记录性质, 2, 序号, 病人id, 主页id, 病人病区id, 病人科室id, 门诊标志, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位, 付数,
               -数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, -应收金额, -实收金额, 开单部门id, 开单人, 执行部门id, 操作员编号_In, 操作员姓名_In,
               发生时间, v_Date, v_结帐id, Decode(v_结帐id, Null, Null, -结帐金额), n_组id, 结论, 摘要, n_校对标志
        From 住院费用记录
        Where ID = r_Cardrow.费用id;
    
      Update 住院费用记录 Set 记录状态 = 3 Where ID = r_Cardrow.费用id;
    
      --如果退病历费，需要同时处理病历费
      If Nvl(退费类型_In, 0) = 1 Then
        Begin
          Select ID
          Into n_病历id
          From 住院费用记录
          Where 记录性质 = 5 And 记录状态 = 1 And NO = 单据号_In And 附加标志 = 8;
        Exception
          When Others Then
            n_病历id := 0;
        End;
        If n_病历id <> 0 Then
          Select 病人费用记录_Id.Nextval Into v_费用id From Dual;
        
          Insert Into 住院费用记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 病人id, 主页id, 病人病区id, 病人科室id, 门诊标志, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位, 付数, 数次,
             加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 操作员编号, 操作员姓名, 发生时间, 登记时间, 结帐id,
             结帐金额, 缴款组id, 结论, 摘要, 费用状态)
            Select v_费用id, NO, 实际票号, 记录性质, 2, 序号, 病人id, 主页id, 病人病区id, 病人科室id, 门诊标志, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位,
                   付数, -数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, -应收金额, -实收金额, 开单部门id, 开单人, 执行部门id, 操作员编号_In,
                   操作员姓名_In, 发生时间, v_Date, v_结帐id, Decode(v_结帐id, Null, Null, -结帐金额), n_组id, 结论, 摘要, n_校对标志
            
            From 住院费用记录
            Where ID = n_病历id;
        
          Update 住院费用记录 Set 记录状态 = 3 Where ID = n_病历id;
        End If;
      End If;
      --处理发卡划价单，如果划价还未收费，直接删除
      Begin
        Select Nvl(记录状态, -1)
        Into v_划价状态
        From 门诊费用记录
        Where 病人id = r_Cardrow.病人id And 记录性质 = 1 And NO = r_Cardrow.摘要;
      Exception
        When Others Then
          v_划价状态 := -1;
      End;
      If v_划价状态 = 0 Then
        Zl_门诊划价记录_Delete(r_Cardrow.摘要);
      End If;
    
      If Nvl(退费类型_In, 0) = 1 Then
        n_退费金额 := -1 * r_Cardrow.冲预交;
      Else
        n_退费金额 := -1 * r_Cardrow.实收金额;
      End If;
    
      --预交款里现收的结算金额
      If r_Cardrow.记帐 = 0 Then
        Select 病人预交记录_Id.Nextval Into n_新预交id From Dual;
        If 退费方式_In Is Null Then
          Insert Into 病人预交记录
            (ID, NO, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号,
             卡号, 交易流水号, 交易说明, 合作单位, 结算性质, 关联交易id, 校对标志, 是否电子票据)
            Select n_新预交id, NO, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 退费方式_In, v_Date, 操作员编号_In, 操作员姓名_In, n_退费金额, v_结帐id,
                   n_组id, 预交类别, Null, Null, Null, Null, Null, 合作单位, 5, n_新预交id, n_校对标志, 是否电子票据
            From 病人预交记录
            Where 记录性质 = 5 And 记录状态 = Decode(n_二次退费, 0, 1, 3) And 结帐id = r_Cardrow.结帐id;
        Else
          Insert Into 病人预交记录
            (ID, NO, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号,
             卡号, 交易流水号, 交易说明, 合作单位, 结算性质, 关联交易id, 校对标志, 是否电子票据)
            Select n_新预交id, NO, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 退费方式_In, v_Date, 操作员编号_In, 操作员姓名_In, n_退费金额, v_结帐id,
                   n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 5, 关联交易id, n_校对标志, 是否电子票据
            From 病人预交记录
            Where 记录性质 = 5 And 记录状态 = Decode(n_二次退费, 0, 1, 3) And 结帐id = r_Cardrow.结帐id;
        End If;
      
        Update 病人预交记录 Set 记录状态 = 3 Where 记录性质 = 5 And 记录状态 = 1 And 结帐id = r_Cardrow.结帐id;
      End If;
    
      --更改医疗卡状态为退卡异常
      If Nvl(r_Cardrow.费用状态, 0) = 0 Then
        n_卡类别id := To_Number(Nvl(r_Cardrow.结论, '0'));
        Update 病人医疗卡信息 Set 状态 = 3 Where 卡类别id = n_卡类别id And 卡号 = r_Cardrow.实际票号;
      End If;
    
      --缓存通用信息
      n_记账       := r_Cardrow.记帐;
      n_病人id     := r_Cardrow.病人id;
      n_主页id     := r_Cardrow.主页id;
      n_病人科室id := r_Cardrow.病人科室id;
      n_病人病区id := r_Cardrow.病人病区id;
      n_开单部门id := r_Cardrow.开单部门id;
      n_执行部门id := r_Cardrow.执行部门id;
      n_收入项目id := r_Cardrow.收入项目id;
    
      Close c_Cardinfo;
    End If;
  Else
    Open c_Booksinfo;
    Fetch c_Booksinfo
      Into r_Booksrow;
    n_组id := Zl_Get组id(操作员姓名_In);
  
    --首先判断要退卡的记录是否存在
    If c_Booksinfo%RowCount = 0 Then
      Close c_Booksinfo;
      v_Err_Msg := '[ZLSOFT]没有发现要退费的记录,该记录可能已经退除！[ZLSOFT]';
      Raise Err_Item;
    Else
      Select Sysdate Into v_Date From Dual;
      Select 病人费用记录_Id.Nextval Into v_费用id From Dual;
    
      If r_Booksrow.记帐 = 0 Then
        If Nvl(销账id_In, 0) = 0 Then
          Select 病人结帐记录_Id.Nextval Into v_结帐id From Dual;
        Else
          v_结帐id := 销账id_In;
        End If;
        n_校对标志 := 1;
      Else
        n_校对标志 := 0;
      End If;
    
      n_退费金额 := -1 * r_Booksrow.实收金额;
    
      --退除病历费费用记录
      Insert Into 住院费用记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 病人id, 主页id, 病人病区id, 病人科室id, 门诊标志, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位, 付数, 数次,
         加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 操作员编号, 操作员姓名, 发生时间, 登记时间, 结帐id,
         结帐金额, 缴款组id, 结论, 摘要, 费用状态)
        Select v_费用id, NO, 实际票号, 记录性质, 2, 序号, 病人id, 主页id, 病人病区id, 病人科室id, 门诊标志, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位, 付数,
               -数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, -应收金额, -实收金额, 开单部门id, 开单人, 执行部门id, 操作员编号_In, 操作员姓名_In,
               发生时间, v_Date, v_结帐id, Decode(v_结帐id, Null, Null, -结帐金额), n_组id, 结论, 摘要, n_校对标志
        From 住院费用记录
        Where ID = r_Booksrow.费用id;
    
      Update 住院费用记录 Set 记录状态 = 3 Where ID = r_Booksrow.费用id;
    
      --预交款里现收的结算金额
      If r_Booksrow.记帐 = 0 Then
        Select 病人预交记录_Id.Nextval Into n_新预交id From Dual;
        If 退费方式_In Is Null Then
          Insert Into 病人预交记录
            (ID, NO, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号,
             卡号, 交易流水号, 交易说明, 合作单位, 结算性质, 关联交易id, 校对标志, 是否电子票据)
            Select n_新预交id, NO, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 退费方式_In, v_Date, 操作员编号_In, 操作员姓名_In, n_退费金额, v_结帐id,
                   n_组id, 预交类别, Null, Null, Null, Null, Null, 合作单位, 5, ID, n_校对标志, 是否电子票据
            From 病人预交记录
            Where 记录性质 = 5 And 记录状态 = Decode(n_二次退费, 0, 1, 3) And 结帐id = r_Booksrow.结帐id;
        Else
          Insert Into 病人预交记录
            (ID, NO, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号,
             卡号, 交易流水号, 交易说明, 合作单位, 结算性质, 关联交易id, 校对标志, 是否电子票据)
            Select n_新预交id, NO, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 退费方式_In, v_Date, 操作员编号_In, 操作员姓名_In, n_退费金额, v_结帐id,
                   n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 5, 关联交易id, n_校对标志, 是否电子票据
            From 病人预交记录
            Where 记录性质 = 5 And 记录状态 = Decode(n_二次退费, 0, 1, 3) And 结帐id = r_Booksrow.结帐id;
        End If;
      
        Update 病人预交记录 Set 记录状态 = 3 Where 记录性质 = 5 And 记录状态 = 1 And 结帐id = r_Booksrow.结帐id;
      End If;
    
      --缓存通用信息
      n_记账       := r_Booksrow.记帐;
      n_病人id     := r_Booksrow.病人id;
      n_主页id     := r_Booksrow.主页id;
      n_病人科室id := r_Booksrow.病人科室id;
      n_病人病区id := r_Booksrow.病人病区id;
      n_开单部门id := r_Booksrow.开单部门id;
      n_执行部门id := r_Booksrow.执行部门id;
      n_收入项目id := r_Booksrow.收入项目id;
      Close c_Booksinfo;
    End If;
  End If;
  ----------------------------------------------------------------------------------------------------------------------------------------

  --相关汇总表的处理
  If n_记账 = 1 Then
    --汇总'病人余额'
    Update 病人余额
    Set 费用余额 = Nvl(费用余额, 0) + n_退费金额
    Where 性质 = 1 And 病人id = n_病人id And Nvl(类型, 2) = Decode(Nvl(n_主页id, 0), 0, 1, 2)
    Returning 费用余额 Into n_返回值;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人余额
        (病人id, 性质, 类型, 预交余额, 费用余额)
      Values
        (n_病人id, 1, Decode(Nvl(n_主页id, 0), 0, 1, 2), 0, n_退费金额);
      n_返回值 := n_退费金额;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete 病人余额 Where 性质 = 1 And 病人id = n_病人id And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
    End If;
  
    --汇总'病人未结费用'
    Update 病人未结费用
    Set 金额 = Nvl(金额, 0) + n_退费金额
    Where 病人id = n_病人id And Nvl(主页id, 0) = n_主页id And Nvl(病人病区id, 0) = n_病人病区id And Nvl(病人科室id, 0) = n_病人科室id And
          Nvl(开单部门id, 0) = n_开单部门id And Nvl(执行部门id, 0) = n_执行部门id And 收入项目id + 0 = n_收入项目id And 来源途径 = 3;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人未结费用
        (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
      Values
        (n_病人id, Decode(n_主页id, 0, Null, n_主页id), Decode(n_病人病区id, 0, Null, n_病人病区id),
         Decode(n_病人科室id, 0, Null, n_病人科室id), Decode(n_开单部门id, 0, Null, n_开单部门id), Decode(n_执行部门id, 0, Null, n_执行部门id),
         n_收入项目id, 3, n_退费金额);
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20999, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_医疗卡记录_Delete;
/

Create Or Replace Procedure Zl_门诊收费结算_Modify
(
  操作类型_In      Number,
  病人id_In        门诊费用记录.病人id%Type,
  结帐id_In        病人预交记录.结帐id%Type,
  结算方式_In      Varchar2,
  冲预交_In        病人预交记录.冲预交%Type := Null,
  退支票额_In      病人预交记录.冲预交%Type := Null,
  卡类别id_In      病人预交记录.卡类别id%Type := Null,
  卡号_In          病人预交记录.卡号%Type := Null,
  交易流水号_In    病人预交记录.交易流水号%Type := Null,
  交易说明_In      病人预交记录.交易说明%Type := Null,
  缴款_In          病人预交记录.缴款%Type := Null,
  找补_In          病人预交记录.找补%Type := Null,
  误差金额_In      门诊费用记录.实收金额%Type := Null,
  完成结算_In      Number := 0,
  缺省结算方式_In  结算方式.名称%Type := Null,
  冲预交病人ids_In Varchar2 := Null,
  更新交款余额_In  Number := 1,
  关联交易id_In    病人预交记录.关联交易id%Type := Null,
  删除原结算_In    Number := 0,
  校对标志_In      病人预交记录.校对标志%Type := 0,
  不控制会话_In    病人预交记录.会话号%Type := 0,
  是否电子票据_In  病人预交记录.是否电子票据%Type := Null
) As
  ------------------------------------------------------------------------------------------------------------------------------ 
  --功能:收费结算时,修改结算的相关信息 
  --操作类型_In: 
  --   0-普通收费方式: 
  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空. 
  --     ②退支票额_In:如果涉及退支票,则传入本次的退支票额,非正常收费时,传入零 
  --   1.三方卡结算: 
  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要" 
  --     ②退支票额_In:传入零 
  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入 
  --   2-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新) 
  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
  --     ②退支票额_In:传入零
  --   3-消费卡结算:
  --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."  消费卡ID:为零时,根据卡号自动定位 
  --     ②退支票额_In:传入零 
  --   4-三方卡结算，多种结算方式: 
  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要|单据号|是否普通结算|卡号" 
  --     ②退支票额_In:传入零 
  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入 

  -- 冲预交_In: 存在冲预交时,传入 
  -- 误差金额_In:存在误差费时,传入 
  -- 完成结算_In:1-完成收费;0-未完成收费 
  -- 冲预交病人ids_In:多个用逗分分离,冲预交时有效(冲预交类别与业务操作保持一致),主要是使用家属的预交款 
  --更新交款余额_In  是否更新人员交款余额，主要是处理统一操作员登录多台自助机的情况 
  --关联交易id_In 操作类型_In 为1,4时必须传入 
  --删除原结算_in 操作类型_In为4时有效，多个结算方式时调用多次该过程 
  --校对标志_In  操作类型_In为4时有效 
    --是否电子票据_In:null-表示过程内部直接判断，非空表示直接以传入的为准
  ------------------------------------------------------------------------------------------------------------------------------ 
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_结算内容 Varchar2(500);
  v_当前结算 Varchar2(50);
  v_卡号     病人医疗卡信息.卡号%Type;
  n_消费卡id 消费卡信息.Id%Type;
  v_名称     消费卡类别目录.名称%Type;
  n_卡类别id 病人预交记录.结算卡序号%Type;
  n_预交id   病人预交记录.Id%Type;

  v_结算方式 病人预交记录.结算方式%Type;
  n_结算金额 病人预交记录.冲预交%Type;
  v_结算号码 病人预交记录.结算号码%Type;
  v_结算摘要 病人预交记录.摘要%Type;
  v_交易人员 病人预交记录.交易人员%Type;

  n_返回值   人员缴款余额.余额%Type;
  n_冲预交   病人预交记录.冲预交%Type;
  v_退支票   病人预交记录.结算方式%Type;
  v_误差费   结算方式.名称%Type;
  n_Count    Number;
  n_Havenull Number;
  l_预交id   t_Numlist := t_Numlist();
  n_会话号   病人预交记录.会话号%Type; --格式：SID+'_'+SERIAL# 
n_是否电子票据 Number(2);
n_险类         保险结算记录.险类%Type;
  Cursor c_Feedata Is
    Select Max(m.病人id) As 病人id, Max(m.登记时间) As 登记时间, Max(m.操作员编号) As 操作员编号, Max(m.操作员姓名) As 操作员姓名, Sum(结帐金额) As 结算金额,
           Max(m.缴款组id) As 缴款组id
    From 门诊费用记录 M
    Where m.结帐id = 结帐id_In;
  r_Feedata c_Feedata%RowType;

  Cursor c_Balancedata Is
    Select 记录性质, NO, 记录状态, 病人id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号
    From 病人预交记录
    Where 结帐id = 结帐id_In And 结算方式 Is Null;
  r_Balancedata c_Balancedata%RowType;

Begin
  IF Nvl(不控制会话_In, 0) = 0 then
    Begin
      Select Sid || '_' || Serial# Into n_会话号 From V$session Where Audsid = Userenv('sessionid');
    Exception
      When Others Then
        n_会话号 := Null;
    End;
  End IF;
  v_交易人员 := zl_UserName;

  Begin
    Select 名称 Into v_误差费 From 结算方式 Where Nvl(性质, 0) = 9;
  Exception
    When Others Then
      v_误差费 := '误差费';
  End;

  --0.正式结算 
  Select Count(1), Max(Decode(结算方式, Null, 1, 0))
  Into n_Count, n_Havenull
  From 病人预交记录
  Where 结帐id = 结帐id_In;

  --1.增加结算方式为空的结算数据 
  n_结算金额 := 0;
  n_Count    := 0;
  Open c_Feedata;
  Begin
    Fetch c_Feedata
      Into r_Feedata;
    --修正或新增结算方式为null的记录 
    Select Nvl(Sum(冲预交), 0) Into n_结算金额 From 病人预交记录 Where 结帐id = 结帐id_In;
    If Nvl(n_Havenull, 0) = 0 Or Round(Nvl(r_Feedata.结算金额, 0), 6) <> Round(Nvl(n_结算金额, 0), 6) Then
      --先删除存在的结算方式为null的记录 
      Delete From 病人预交记录 Where 结帐id = 结帐id_In And 结算方式 Is Null;
      Select Nvl(Sum(冲预交), 0) Into n_结算金额 From 病人预交记录 Where 结帐id = 结帐id_In;
    
      n_结算金额 := Round(Nvl(r_Feedata.结算金额, 0) - n_结算金额, 6);
    
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志, 结算性质, 会话号)
      Values
        (病人预交记录_Id.Nextval, 3, Null, 1, Decode(病人id_In, 0, Null, 病人id_In), Null, r_Feedata.登记时间, r_Feedata.操作员编号,
         r_Feedata.操作员姓名, n_结算金额, 结帐id_In, r_Feedata.缴款组id, Sysdate, v_交易人员, -1 * 结帐id_In, 1, 3, n_会话号);
    End If;
  Exception
    When Others Then
      n_Count := 1;
  End;
  Close c_Feedata;
  If n_Count = 1 Then
    v_Err_Msg := '未找到指定的收费明细数据,结算操作失败！';
    Raise Err_Item;
  End If;

  If 操作类型_In = 0 And Nvl(退支票额_In, 0) <> 0 Then
    Begin
      Select b.名称
      Into v_退支票
      From 结算方式应用 A, 结算方式 B
      Where a.应用场合 = '收费' And b.名称 = a.结算方式 And Nvl(b.应付款, 0) = 1 And Rownum <= 1;
    Exception
      When Others Then
        v_退支票 := '无';
    End;
    If v_退支票 = '无' Then
      v_Err_Msg := '在结算场合中,不存在结算性质为应付款的结算方式,请在[结算方式]中设置！';
      Raise Err_Item;
    End If;
  End If;

  Open c_Balancedata;
  Fetch c_Balancedata
    Into r_Balancedata;

  If Nvl(误差金额_In, 0) <> 0 Then
    Update 病人预交记录
    Set 冲预交 = Nvl(冲预交, 0) + Nvl(误差金额_In, 0)
    Where 结帐id = 结帐id_In And 结算方式 = v_误差费;
    If Sql%NotFound Then
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志, 卡类别id,
         结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号)
      Values
        (病人预交记录_Id.Nextval, 3, Null, 1, r_Balancedata.病人id, Null, Null, v_误差费, r_Balancedata.收款时间, r_Balancedata.操作员编号,
         r_Balancedata.操作员姓名, 误差金额_In, 结帐id_In, r_Balancedata.缴款组id, Sysdate, v_交易人员, r_Balancedata.结算序号, 2, Null, Null,
         卡号_In, 交易流水号_In, 交易说明_In, Null, 3, n_会话号);
    End If;
  
    Update 病人预交记录 Set 冲预交 = 冲预交 - Nvl(误差金额_In, 0) Where 结帐id = 结帐id_In And 结算方式 Is Null;
  End If;

  --预交款处理 
  If Nvl(冲预交_In, 0) <> 0 Then
    If Nvl(病人id_In, 0) = 0 Then
      v_Err_Msg := '不能确定病人的病人ID,收费不能使用预交款结算,结算操作失败！';
      Raise Err_Item;
    End If;
  
    Zl_病人预交记录_冲预交(病人id_In, 结帐id_In, 冲预交_In, 1, r_Balancedata.操作员编号, r_Balancedata.操作员姓名, r_Balancedata.收款时间,
                  冲预交病人ids_In, 3, 1);
  End If;

  If 操作类型_In = 0 Then
    If Nvl(退支票额_In, 0) <> 0 Then
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志, 卡类别id,
         结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号)
      Values
        (病人预交记录_Id.Nextval, 3, Null, 1, r_Balancedata.病人id, Null, Null, v_退支票, r_Balancedata.收款时间, r_Balancedata.操作员编号,
         r_Balancedata.操作员姓名, 退支票额_In, 结帐id_In, r_Balancedata.缴款组id, Sysdate, v_交易人员, r_Balancedata.结算序号, 2, Null, Null,
         卡号_In, 交易流水号_In, 交易说明_In, Null, 3, n_会话号);
    
      Update 病人预交记录 Set 冲预交 = 冲预交 - 退支票额_In Where 结帐id = 结帐id_In And 结算方式 Is Null;
    End If;
  
    --各个收费结算 :格式为:"结算方式|结算金额|结算号码|结算摘要||.." 
    v_结算内容 := 结算方式_In || '||';
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      v_结算号码 := LTrim(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_结算摘要 := LTrim(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
    
      If Nvl(n_结算金额, 0) <> 0 Then
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志,
           卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号)
        Values
          (病人预交记录_Id.Nextval, 3, Null, 1, r_Balancedata.病人id, Null, v_结算摘要, v_结算方式, r_Balancedata.收款时间,
           r_Balancedata.操作员编号, r_Balancedata.操作员姓名, n_结算金额, 结帐id_In, r_Balancedata.缴款组id, Sysdate, v_交易人员,
           r_Balancedata.结算序号, 2, Null, Null, 卡号_In, 交易流水号_In, 交易说明_In, v_结算号码, 3, n_会话号);
      
        Update 病人预交记录 Set 冲预交 = 冲预交 - n_结算金额 Where 结帐id = 结帐id_In And 结算方式 Is Null;
      End If;
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  --1.三方卡结算交易 
  If 操作类型_In = 1 Then
    v_当前结算 := 结算方式_In;
    v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
    v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
    n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
    v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
    v_结算号码 := LTrim(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
    v_结算摘要 := LTrim(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
  
    If Nvl(n_结算金额, 0) <> 0 Then
      Select Count(1) Into n_Count From 病人预交记录 Where ID = 关联交易id_In And Rownum < 2;
      If n_Count = 0 And Nvl(关联交易id_In, 0) <> 0 Then
        n_预交id := 关联交易id_In;
      Else
        Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      End If;
    
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志, 卡类别id,
         结算卡序号, 卡号, 关联交易id, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号)
      Values
        (n_预交id, 3, Null, 1, r_Balancedata.病人id, Null, v_结算摘要, v_结算方式, r_Balancedata.收款时间, r_Balancedata.操作员编号,
         r_Balancedata.操作员姓名, n_结算金额, 结帐id_In, r_Balancedata.缴款组id, Sysdate, v_交易人员, r_Balancedata.结算序号, 2, 卡类别id_In,
         Null, 卡号_In, 关联交易id_In, 交易流水号_In, 交易说明_In, v_结算号码, 3, n_会话号);
    
      Update 病人预交记录 Set 冲预交 = 冲预交 - n_结算金额 Where 结帐id = 结帐id_In And 结算方式 Is Null;
    
      --调用其他结算信息更新
      Zl_Custom_Balance_Update(n_预交id);
    End If;
  End If;

  --2.医保结算(调用此过程,采取平均分摊的方式分摊结算情况):这种情况医保结处后,必须全退 
  If 操作类型_In = 2 Then
    --2.1检查是否已经存在医保结算数据,存在先删除 
    n_结算金额 := 0;
    For v_医保 In (Select ID, 结算方式, 冲预交
                 From 病人预交记录 A
                 Where a.结帐id = 结帐id_In And Exists
                  (Select 1 From 结算方式 Where a.结算方式 = 名称 And 性质 In (3, 4)) And Mod(记录性质, 10) <> 1) Loop
      n_结算金额 := n_结算金额 + Nvl(v_医保.冲预交, 0);
      l_预交id.Extend;
      l_预交id(l_预交id.Count) := v_医保.Id;
    End Loop;
  
    Update 病人预交记录 Set 冲预交 = Nvl(冲预交, 0) + Nvl(n_结算金额, 0) Where 结帐id = 结帐id_In And 结算方式 Is Null;
  
    Forall I In 1 .. l_预交id.Count
      Delete From 病人预交记录 Where ID = l_预交id(I);
  
    If 结算方式_In Is Not Null Then
      v_结算内容 := 结算方式_In || '||';
    End If;
  
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      n_结算金额 := To_Number(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
    
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志, 结算性质,
         会话号)
      Values
        (病人预交记录_Id.Nextval, 3, Null, 1, r_Balancedata.病人id, Null, '保险结算', v_结算方式, r_Balancedata.收款时间,
         r_Balancedata.操作员编号, r_Balancedata.操作员姓名, n_结算金额, 结帐id_In, r_Balancedata.缴款组id, Sysdate, v_交易人员,
         r_Balancedata.结算序号, 1, 3, n_会话号);
    
      --更新数据(结算方式为NULL的) 
      Update 病人预交记录
      Set 冲预交 = 冲预交 - n_结算金额
      Where 结帐id = 结帐id_In And 结算方式 Is Null
      Returning Nvl(冲预交, 0) Into n_返回值;
    
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  --3-消费卡批量结算 
  If 操作类型_In = 3 Then
    v_结算内容 := 结算方式_In || '||';
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      --卡类别ID|卡号|消费卡ID|消费金额 
      n_卡类别id := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      v_卡号     := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_消费卡id := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_结算金额 := To_Number(v_当前结算);
      Begin
        Select 名称, 结算方式 Into v_名称, v_结算方式 From 消费卡类别目录 Where 编号 = 卡类别id_In;
      Exception
        When Others Then
          v_名称 := Null;
      End;
      If v_名称 Is Null Then
        v_Err_Msg := '未找到对应的结算卡接口,本次刷卡消费失败!';
        Raise Err_Item;
      End If;
      If v_结算方式 Is Null Then
        v_Err_Msg := v_名称 || '未设置对应的结算方式,本次刷卡消费失败!';
        Raise Err_Item;
      End If;
    
      If Nvl(n_结算金额, 0) <> 0 Then
        Update 病人预交记录
        Set 冲预交 = Nvl(冲预交, 0) + n_结算金额
        Where 结帐id = r_Balancedata. 结帐id And 结算方式 = v_结算方式 And 结算卡序号 = n_卡类别id
        Returning ID Into n_预交id;
        If Sql%NotFound Then
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 结算卡序号,
             校对标志, 结算性质, 会话号)
          Values
            (n_预交id, 3, Null, 1, r_Balancedata. 病人id, Null, Null, v_结算方式, r_Balancedata. 收款时间, r_Balancedata. 操作员编号,
             r_Balancedata. 操作员姓名, n_结算金额, r_Balancedata. 结帐id, r_Balancedata. 缴款组id, Sysdate, v_交易人员,
             r_Balancedata. 结算序号, n_卡类别id, 2, 3, n_会话号);
        End If;
      
        Zl_病人卡结算记录_支付(n_卡类别id, v_卡号, n_消费卡id, n_结算金额, n_预交id, r_Balancedata. 操作员编号, r_Balancedata. 操作员姓名,
                      r_Balancedata. 收款时间);
      
        --更新数据(结算方式为NULL的) 
        Update 病人预交记录
        Set 冲预交 = 冲预交 - n_结算金额
        Where 结帐id = r_Balancedata. 结帐id And 结算方式 Is Null And Nvl(校对标志, 0) = 1
        Returning Nvl(冲预交, 0) Into n_返回值;
      End If;
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  --4.三方卡结算，多种结算方式 
  If 操作类型_In = 4 Then
    If Nvl(删除原结算_In, 0) = 1 Then
      --1.1检查是否已经存在三方卡结算数据,存在先删除 
      n_结算金额 := 0;
      For c_结算 In (Select ID, 结算方式, 冲预交
                   From 病人预交记录 A
                   Where 结帐id = 结帐id_In And 卡类别id = 卡类别id_In And 关联交易id = 关联交易id_In) Loop
        n_结算金额 := n_结算金额 + Nvl(c_结算.冲预交, 0);
        l_预交id.Extend;
        l_预交id(l_预交id.Count) := c_结算.Id;
      End Loop;
    
      Update 病人预交记录
      Set 冲预交 = Nvl(冲预交, 0) + Nvl(n_结算金额, 0)
      Where 结帐id = 结帐id_In And 结算方式 Is Null;
    
      Forall I In 1 .. l_预交id.Count
        Delete From 病人预交记录 Where ID = l_预交id(I);
    End If;
  
    n_预交id := 0;
    --格式：结算方式|结算金额|结算号码|结算摘要|单据号|是否普通结算|卡号 
    For c_结算 In (Select Max(Decode(序号, 1, 值, Null)) As 结算方式, Zl_To_Number(Max(Decode(序号, 2, 值, ''))) As 结算金额,
                        Trim(Max(Decode(序号, 3, 值, ''))) As 结算号码, Trim(Max(Decode(序号, 4, 值, ''))) As 结算摘要,
                        Trim(Max(Decode(序号, 5, 值, ''))) As 单据号, Zl_To_Number(Max(Decode(序号, 6, 值, ''))) As 是否普通结算,
                        Trim(Max(Decode(序号, 7, 值, ''))) As 卡号
                 From (Select Rownum As 序号, Column_Value As 值 From Table(f_Str2list(结算方式_In, '|')))
                 Having Nvl(Zl_To_Number(Max(Decode(序号, 2, 值, ''))), 0) <> 0) Loop
    
      Update 病人预交记录
      Set 冲预交 = 冲预交 + c_结算.结算金额
      Where 结帐id = 结帐id_In And 结算方式 = c_结算.结算方式 And 关联交易id = 关联交易id_In
      Returning ID Into n_预交id;
      If Sql%NotFound Then
        Select Count(1) Into n_Count From 病人预交记录 Where ID = 关联交易id_In And Rownum < 2;
        If n_Count = 0 Then
          n_预交id := 关联交易id_In;
        Else
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
        End If;
      
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志,
           卡类别id, 卡号, 关联交易id, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号)
        Values
          (n_预交id, 3, Null, 1, r_Balancedata.病人id, Null, c_结算.结算摘要, c_结算.结算方式, r_Balancedata.收款时间, r_Balancedata.操作员编号,
           r_Balancedata.操作员姓名, c_结算.结算金额, 结帐id_In, r_Balancedata.缴款组id, Sysdate, v_交易人员, r_Balancedata.结算序号, 校对标志_In,
           Decode(c_结算.是否普通结算, 1, Null, 卡类别id_In), Decode(c_结算.是否普通结算, 1, Null, Nvl(c_结算.卡号, 卡号_In)), 关联交易id_In,
           交易流水号_In, 交易说明_In, c_结算.结算号码, 3, n_会话号);
      End If;
    
      Update 病人预交记录 Set 冲预交 = 冲预交 - c_结算.结算金额 Where 结帐id = 结帐id_In And 结算方式 Is Null;
    
      If c_结算.单据号 Is Not Null Then
        Insert Into 医保结算明细
          (结帐id, NO, 结算方式, 金额, 卡类别id, 关联交易id, 交易流水号, 交易说明)
        Values
          (结帐id_In, c_结算.单据号, c_结算.结算方式, c_结算.结算金额, Decode(c_结算.是否普通结算, 1, Null, 卡类别id_In), 关联交易id_In, 交易流水号_In,
           交易说明_In);
      End If;
    
      --调用其他结算信息更新
      Zl_Custom_Balance_Update(n_预交id);
    End Loop;
  End If;

  If Nvl(完成结算_In, 0) = 0 Then
    Return;
  End If;

  ----------------------------------------------------------------------------------------- 
  --完成收费,需要处理人员缴款余额,预交记录(结算方式=NULL) 

  --1.删除结算方式为NULL的预交记录 
  Delete 病人预交记录 Where 结帐id = 结帐id_In And 结算方式 Is Null And Nvl(冲预交, 0) = 0;
  If Sql%NotFound Then
    Select Count(*) Into n_Count From 病人预交记录 A Where 结帐id = 结帐id_In And 结算方式 Is Null;
    If n_Count <> 0 Then
      v_Err_Msg := '还存在未缴款的数据,不能完成结算!';
    Else
      v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！!';
    End If;
    Raise Err_Item;
  End If;

  --检查门诊费用记录与病人预交记录的金额是否相等 
  n_结算金额 := 0;
  n_冲预交   := 0;
  Select Nvl(Sum(实收金额), 0) Into n_结算金额 From 门诊费用记录 Where 结帐id = 结帐id_In;
  Select Nvl(Sum(冲预交), 0) Into n_冲预交 From 病人预交记录 Where 结帐id = 结帐id_In;
  If n_结算金额 <> n_冲预交 Then
    v_Err_Msg := '结算信息有误，实收金额(' || n_结算金额 || ')与结算金额(' || n_冲预交 || ')不一致，不能完成结算！';
    Raise Err_Item;
  End If;

  --结算金额为零时，增加一条金额为0的病人预交记录 
  Select Count(1) Into n_Count From 病人预交记录 A Where 结帐id = 结帐id_In;
  If n_Count = 0 Then
    v_结算方式 := 缺省结算方式_In;
    If v_结算方式 Is Null Then
      Begin
        Select 结算方式 Into v_结算方式 From 结算方式应用 Where 应用场合 = '收费' And Nvl(缺省标志, 0) = 1;
      Exception
        When Others Then
          v_结算方式 := Null;
      End;
      If v_结算方式 Is Null Then
        Begin
          Select 名称 Into v_结算方式 From 结算方式 Where Nvl(性质, 0) = 1;
        Exception
          When Others Then
            v_结算方式 := '现金';
        End;
      End If;
    End If;
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志, 卡类别id,
       结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号)
    Values
      (病人预交记录_Id.Nextval, 3, Null, 1, r_Balancedata.病人id, Null, Null, v_结算方式, r_Balancedata.收款时间, r_Balancedata.操作员编号,
       r_Balancedata.操作员姓名, 0, 结帐id_In, r_Balancedata.缴款组id, Sysdate, v_交易人员, r_Balancedata.结算序号, 2, Null, Null, Null,
       Null, 交易说明_In, Null, 3, n_会话号);
  End If;

  n_是否电子票据 := 是否电子票据_In;
  If 是否电子票据_In Is Null Then
    Select Nvl(Max(险类), 0) Into n_险类 From 保险结算记录 Where 记录id = 结帐id_In And 性质 = 1;
    n_是否电子票据 := Zl_Fun_Isstarteinvoice(1, n_险类);
  End If;

  --2.处理缴款数据和找补数据及校对标志更新为0，会话号更新为NULL 
  Update 病人预交记录 Set 缴款 = 缴款_In, 找补 = 找补_In, 校对标志 = 0, 会话号 = Null,是否电子票据=n_是否电子票据 Where 结帐id = 结帐id_In;

  --3.更新费用状态 
  Update 门诊费用记录 Set 费用状态 = 0 Where 结帐id = 结帐id_In;

  --4.更新人员缴款数据 
  If Nvl(更新交款余额_In, 1) = 1 Then
    For c_缴款 In (Select 结算方式, 操作员姓名, Sum(Nvl(a.冲预交, 0)) As 冲预交
                 From 病人预交记录 A
                 Where a.结帐id = 结帐id_In And Mod(a.记录性质, 10) <> 1
                 Group By 结算方式, 操作员姓名) Loop
    
      Update 人员缴款余额
      Set 余额 = Nvl(余额, 0) + Nvl(c_缴款.冲预交, 0)
      Where 收款员 = c_缴款.操作员姓名 And 性质 = 1 And 结算方式 = c_缴款.结算方式;
      If Sql%RowCount = 0 Then
        Insert Into 人员缴款余额
          (收款员, 结算方式, 性质, 余额)
        Values
          (c_缴款.操作员姓名, c_缴款.结算方式, 1, Nvl(c_缴款.冲预交, 0));
      End If;
    End Loop;
  End If;

  --5.相关业务数据处理 
  Zl_门诊收费记录_完成收费(结帐id_In);

  --消息集成处理 
  --结算类型:1-收费结算，2-补充结算 
  --结帐ID:结算id 
  b_Message.Zlhis_Charge_002(1, 结帐id_In);

  --收费后产生导引 
  Begin
    Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
      Using 4, 结帐id_In;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊收费结算_Modify;
/

CREATE OR REPLACE Procedure Zl_电子票据开票点_Insert
(
  Id_In       In 电子票据开票点.Id%Type,
  上级id_In   In 电子票据开票点.上级id%Type,
  编码_In     In 电子票据开票点.编码%Type,
  名称_In     In 电子票据开票点.名称%Type,
  简码_In     In 电子票据开票点.简码%Type,
  院区_In     In 电子票据开票点.院区%Type,
  客户端_In   In 电子票据开票点.客户端%Type,
  部门id_In   In 电子票据开票点.部门id%Type,
  位置_In     In 电子票据开票点.位置%Type,
  末级_In     In 电子票据开票点.末级%Type := Null,
  建档时间_In In 电子票据开票点.建档时间%Type := Null
  
) Is
  d_建档时间 Date;
Begin

  d_建档时间 := 建档时间_In;
  If d_建档时间 Is Null Then
    d_建档时间 := Sysdate;
  End If;

  Insert Into 电子票据开票点
    (ID, 上级id, 编码, 名称, 简码, 院区, 客户端, 部门id, 位置, 末级, 建档时间, 撤档时间)
  Values
    (Id_In, 上级id_In, 编码_In, 名称_In, 简码_In, 院区_In, 客户端_In, 部门id_In, 位置_In, 末级_In, d_建档时间,
     To_Date('3000-01-01', 'yyyy-mm-dd hh24:mi:ss'));

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_电子票据开票点_Insert;
/
CREATE OR REPLACE Procedure Zl_电子票据开票点_Update
(
  Id_In     In 电子票据开票点.Id%Type,
  上级id_In In 电子票据开票点.上级id%Type,
  编码_In   In 电子票据开票点.编码%Type,
  名称_In   In 电子票据开票点.名称%Type,
  简码_In   In 电子票据开票点.简码%Type,
  院区_In   In 电子票据开票点.院区%Type,
  客户端_In In 电子票据开票点.客户端%Type,
  部门id_In In 电子票据开票点.部门id%Type,
  位置_In   In 电子票据开票点.位置%Type
) Is
  n_上级id 电子票据开票点.上级id%Type;
Begin
  Select 上级id Into n_上级id From 电子票据开票点 Where ID = Id_In;

  Update 电子票据开票点
  Set 上级id = 上级id_In, 编码 = 编码_In, 名称 = 名称_In, 简码 = 简码_In, 院区 = 院区_In, 客户端 = 客户端_In, 部门id = 部门id_In, 位置 = 位置_In
  Where ID = Id_In;

  Update 电子票据开票点
  Set 编码 = 编码_In || Substr(编码, Length(编码_In) + 1)
  Where ID In (Select ID From 电子票据开票点 Start With 上级id = Id_In Connect By Prior ID = 上级id);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_电子票据开票点_Update;
/
CREATE OR REPLACE Procedure Zl_电子票据开票点_Start(Id_In In 电子票据开票点.Id%Type) Is

Begin
  Update 电子票据开票点 Set 撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd hh24:mi:ss') Where ID = Id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_电子票据开票点_Start;
/
CREATE OR REPLACE Procedure Zl_电子票据开票点_Stop(Id_In In 电子票据开票点.Id%Type) Is

Begin
  Update 电子票据开票点 Set 撤档时间 = Sysdate Where ID = Id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_电子票据开票点_Stop;
/
Create Or Replace Procedure Zl_电子票据开票点_Delete(Id_In In 电子票据开票点.Id%Type) Is
Begin
  Delete From 电子票据开票点 Where ID = Id_In;
  Delete From 票据开票点对照 Where 开票点id = Id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_电子票据开票点_Delete;
/
Create Or Replace Procedure Zl_票据开票点对照_Update
(
  操作_In     In Number,
  Id_In       In 票据开票点对照.Id%Type := Null,
  开票点id_In In 电子票据开票点.Id%Type := Null,
  人员id_In   In 票据开票点对照.人员id%Type := Null,
  客户端_In   In 票据开票点对照.客户端%Type := Null
) Is
  --说明
  --操作_In:0-新增;1-修改;2-删除;3-删除所有
Begin
  --删除所有
  If Nvl(操作_In, 0) = 3 Then
    Delete From 票据开票点对照;
    Return;
  End If;
  --删除
  If Nvl(操作_In, 0) = 2 Then
    Delete From 票据开票点对照 Where ID = Id_In;
    Return;
  End If;
  --修改
  If Nvl(操作_In, 0) = 1 Then
    Update 票据开票点对照 Set 人员id = 人员id_In, 客户端 = 客户端_In Where 开票点id = Id_In;
    Return;
  End If;

  --新增
  Insert Into 票据开票点对照 (ID, 开票点id, 人员id, 客户端) Values (Id_In, 开票点id_In, 人员id_In, 客户端_In);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_票据开票点对照_Update;
/


Create Or Replace Procedure Zl_病人预交记录_Insert
(
  Id_In           病人预交记录.Id%Type,
  单据号_In       病人预交记录.No%Type,
  票据号_In       票据使用明细.号码%Type,
  病人id_In       病人预交记录.病人id%Type,
  主页id_In       病人预交记录.主页id%Type,
  科室id_In       病人预交记录.科室id%Type,
  金额_In         病人预交记录.金额%Type,
  结算方式_In     病人预交记录.结算方式%Type,
  结算号码_In     病人预交记录.结算号码%Type,
  缴款单位_In     病人预交记录.缴款单位%Type,
  单位开户行_In   病人预交记录.单位开户行%Type,
  单位帐号_In     病人预交记录.单位帐号%Type,
  摘要_In         病人预交记录.摘要%Type,
  操作员编号_In   病人预交记录.操作员编号%Type,
  操作员姓名_In   病人预交记录.操作员姓名%Type,
  领用id_In       票据使用明细.领用id%Type,
  预交类别_In     病人预交记录.预交类别%Type := Null,
  卡类别id_In     病人预交记录.卡类别id%Type := Null,
  结算卡序号_In   病人预交记录.结算卡序号%Type := Null,
  卡号_In         病人预交记录.卡号%Type := Null,
  交易流水号_In   病人预交记录.交易流水号%Type := Null,
  交易说明_In     病人预交记录.交易说明%Type := Null,
  合作单位_In     病人预交记录.合作单位%Type := Null,
  收款时间_In     病人预交记录.收款时间%Type := Null,
  操作类型_In     Integer := 0,
  结帐id_In       病人预交记录.结帐id%Type := Null,
  结算性质_In     病人预交记录.结算性质%Type := Null,
  退款检查_In     Number := 0,
  强制退现_In     Number := 0,
  更新交款余额_In Number := 1,
  是否转账_In     Number := 0,
  校对标志_In     病人预交记录.校对标志%Type := Null,
  操作状态_In     Number := 0,
  预交电子票据_In 病人预交记录.预交电子票据%Type := Null,
  险类_In         保险结算记录.险类%Type := Null
) As
  ----------------------------------------------
  --操作类型_In:0-正常缴预交;1-保存为未生效的预交款;3-余额退款
  --结帐ID_IN:>0时,表示某次结帐时,同步产生的预交记录
  --退款检查_In;0-忽略退款金额是否大于了病人余额；1-检查退款金额
  --更新交款余额_In:0-在 zl_人员缴款余额_Update 中更新；1-在本过程中更新
  --强制退现_In:0-不强制，1-三方卡或消费卡不允许退现但强制退现金给病人
  --是否转账_In:0-原样退或退现，1-转账到支持的三方卡上
  --操作状态_In:0-正常结算，1-保存为异常单据，2-完成异常结算

  v_Err_Msg Varchar2(200);
  Err_Item Exception;

  v_性质     结算方式.性质%Type;
  v_打印id   票据打印内容.Id%Type;
  v_担保     病人信息.担保性质%Type;
  v_Date     Date;
  n_返回值   病人余额.预交余额%Type;
  n_组id     财务缴款分组.Id%Type;
  n_病人余额 病人余额.预交余额%Type;
  n_三方预交 病人余额.预交余额%Type;
  n_退款金额 病人预交记录.金额%Type;
  n_剩余款   病人预交记录.金额%Type;
  n_结帐id   病人结帐记录.Id%Type;
  n_险类     保险结算记录.险类%Type;

  n_预交电子票据 Number(2);
  Cursor c_冲预交 Is
    Select a.Id, a.No, a.病人id, a.预交类别, a.卡类别id, a.卡号, a.交易流水号, a.交易说明, 0 As 序号, a.收款时间, a.金额 As 预交金, a.关联交易id
    From 病人预交记录 A
    Where Rownum < 2;
  r_冲预交 c_冲预交%RowType;

  Type Ty_剩余款 Is Ref Cursor;
  c_剩余款 Ty_剩余款; --动态游标变量
Begin

  n_预交电子票据 := 预交电子票据_In;
  If n_预交电子票据 Is Null Then
    n_险类 := 险类_In;
    If 险类_In Is Null Then
      Select Nvl(Max(险类), 0) Into n_险类 From 保险结算记录 Where 记录id = Id_In And 性质 = 3;
    End If;
    n_预交电子票据 := Zl_Fun_Isstarteinvoice(2, n_险类, 预交类别_In);
  End If;

  v_Date := 收款时间_In;
  If v_Date Is Null Then
    Select Sysdate Into v_Date From Dual;
  End If;
  n_组id := Zl_Get组id(操作员姓名_In);

  If Not (操作类型_In = 3 And 操作状态_In = 2) Then
  
    If 操作状态_In = 0 Or 操作状态_In = 1 Then
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要, 缴款组id,
         预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 冲预交, 结算性质, 校对标志, 关联交易id, 交易时间, 交易人员, 预交电子票据)
      Values
        (Id_In, 单据号_In, Decode(操作状态_In, 0, 票据号_In, Null), 1, Decode(操作状态_In, 1, 0, 1), 病人id_In,
         Decode(主页id_In, 0, Null, 主页id_In), Decode(科室id_In, 0, Null, 科室id_In), 金额_In, 结算方式_In, 结算号码_In, v_Date, 缴款单位_In,
         单位开户行_In, 单位帐号_In, 操作员编号_In, 操作员姓名_In, 摘要_In, n_组id, 预交类别_In, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In,
         合作单位_In, 结帐id_In, Decode(结帐id_In, Null, Null, 0), 结算性质_In, 校对标志_In, Id_In, 收款时间_In, 操作员姓名_In, n_预交电子票据);
    
      If Nvl(卡类别id_In, 0) <> 0 Then
        --自定义过程调用
        Zl_Custom_Balance_Update(Id_In);
      End If;
    End If;
  
    If 操作类型_In = 0 Then
      --更新预交单据余额
      Insert Into 预交单据余额 (预交id, 病人id, 预交类别, 预交余额) Values (Id_In, 病人id_In, 预交类别_In, 金额_In);
    
    Elsif 操作类型_In = 1 Then
      --暂不处理汇总表
      Return;
    Elsif 操作类型_In = 3 Then
      --更新预交单据余额
      Insert Into 预交单据余额 (预交id, 病人id, 预交类别, 预交余额) Values (Id_In, 病人id_In, 预交类别_In, 金额_In);
    
      --生成一条原预交ID的冲销记录，同时也生成一条余额退款的冲销记录
      --代收款项不能进行冲销
      Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
      If Nvl(卡类别id_In, 0) = 0 And Nvl(结算卡序号_In, 0) = 0 Then
        --退现，包括普通结算方式退现、强制退现、三方卡允许退现
        Open c_剩余款 For
          Select Max(Decode(a.记录性质, 1, Decode(a.记录状态, 2, 0, a.Id), 0)) As ID, a.No, a.病人id, a.预交类别, a.卡类别id, a.卡号,
                 a.交易流水号, a.交易说明, Min(Decode(Sign(a.金额), -1, 0, 1)) As 序号, Min(Decode(a.记录性质, 1, a.收款时间, Null)) As 收款时间,
                 Nvl(Sum(a.金额), 0) - Nvl(Sum(a.冲预交), 0) As 预交金, Max(a.关联交易id) As 关联交易id
          From 病人预交记录 A, 医疗卡类别 B, 消费卡类别目录 C
          Where a.病人id = 病人id_In And a.记录性质 In (1, 11) And a.预交类别 = Nvl(预交类别_In, 2) And a.卡类别id = b.Id(+) And
                Decode(强制退现_In, 1, 1, Nvl(b.是否退现, 1)) = 1 And a.卡类别id = c.编号(+) And
                Decode(强制退现_In, 1, 1, Nvl(c.是否退现, 1)) = 1 And a.结算方式 Not In (Select 名称 From 结算方式 Where 性质 = 5)
          Group By a.No, a.病人id, a.预交类别, a.卡类别id, a.卡号, a.交易流水号, a.交易说明
          Having Nvl(Sum(a.金额), 0) - Nvl(Sum(a.冲预交), 0) <> 0
          Order By 序号, 收款时间;
      Elsif Nvl(是否转账_In, 0) = 1 Then
        --转账，三方卡允许退现或者强制退现，传入的卡号可能不是原卡号,金额由同种卡类别的预交缴款分摊
        --目前只支持同一种卡转账
        Open c_剩余款 For
          Select Max(Decode(a.记录性质, 1, Decode(a.记录状态, 2, 0, a.Id), 0)) As ID, a.No, a.病人id, a.预交类别, a.卡类别id, a.卡号,
                 a.交易流水号, a.交易说明, Min(Decode(Sign(a.金额), -1, 0, 1)) As 序号, Min(Decode(a.记录性质, 1, a.收款时间, Null)) As 收款时间,
                 Nvl(Sum(a.金额), 0) - Nvl(Sum(a.冲预交), 0) As 预交金, Max(a.关联交易id) As 关联交易id
          From 病人预交记录 A, 医疗卡类别 B
          Where a.病人id = 病人id_In And a.记录性质 In (1, 11) And a.预交类别 = Nvl(预交类别_In, 2) And a.卡类别id = b.Id(+) And
                Nvl(卡类别id, 0) = Nvl(卡类别id_In, 0) And Nvl(交易流水号, '-') = Nvl(交易流水号_In, '-') And
                a.结算方式 Not In (Select 名称 From 结算方式 Where 性质 = 5)
          Group By a.No, a.病人id, a.预交类别, a.卡类别id, a.卡号, a.交易流水号, a.交易说明
          Having Nvl(Sum(a.金额), 0) - Nvl(Sum(a.冲预交), 0) <> 0
          Order By 序号, 收款时间;
      Else
        --退三方卡或者是消费卡，根据卡类别ID、结算卡序号、卡号、交易流水号缺省原预交记录，如果不能确定唯一则进行分摊
        Open c_剩余款 For
          Select Max(Decode(a.记录性质, 1, Decode(a.记录状态, 2, 0, a.Id), 0)) As ID, a.No, a.病人id, a.预交类别, a.卡类别id, a.卡号,
                 a.交易流水号, a.交易说明, Min(Decode(Sign(a.金额), -1, 0, 1)) As 序号, Min(Decode(a.记录性质, 1, a.收款时间, Null)) As 收款时间,
                 Nvl(Sum(a.金额), 0) - Nvl(Sum(a.冲预交), 0) As 预交金, Max(a.关联交易id) As 关联交易id
          From 病人预交记录 A
          Where a.病人id = 病人id_In And a.记录性质 In (1, 11) And a.预交类别 = Nvl(预交类别_In, 2) And
                Nvl(a.卡类别id, 0) = Nvl(卡类别id_In, 0) And Nvl(a.结算卡序号, 0) = Nvl(结算卡序号_In, 0) And
                Nvl(a.卡号, '-') = Nvl(卡号_In, '-') And Nvl(交易流水号, '-') = Nvl(交易流水号_In, '-') And
                a.结算方式 Not In (Select 名称 From 结算方式 Where 性质 = 5)
          Group By a.No, a.病人id, a.预交类别, a.卡类别id, a.卡号, a.交易流水号, a.交易说明
          Having Nvl(Sum(a.金额), 0) - Nvl(Sum(a.冲预交), 0) <> 0
          Order By 序号, 收款时间;
      End If;
    
      n_剩余款   := -1 * 金额_In;
      n_退款金额 := 0;
      Loop
        Fetch c_剩余款
          Into r_冲预交;
        Exit When c_剩余款%NotFound;
        If r_冲预交.No <> 单据号_In Then
          If n_剩余款 > r_冲预交.预交金 Then
            n_退款金额 := r_冲预交.预交金;
            n_剩余款   := n_剩余款 - n_退款金额;
          Else
            n_退款金额 := n_剩余款;
            n_剩余款   := 0;
          End If;
        
          If Nvl(n_退款金额, 0) <> 0 Then
            Update 病人预交记录 Set 结帐id = n_结帐id Where NO = r_冲预交.No And 记录性质 = 1 And 结帐id Is Null;
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 预交类别, 科室id, 金额, 结算方式, 结算号码, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 收款时间, 操作员姓名,
               摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 冲预交, 结算性质, 关联交易id, 交易时间, 交易人员, 校对标志, 预交电子票据)
              Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 1, 病人id, 主页id, 预交类别, 科室id, Null, 结算方式, 结算号码, 缴款单位, 单位开户行, 单位帐号,
                     操作员编号_In, v_Date, 操作员姓名_In, 摘要, n_组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, n_结帐id, n_退款金额, Null,
                     Nvl(r_冲预交.关联交易id, r_冲预交.Id), 收款时间_In, 操作员姓名_In, 校对标志_In, n_预交电子票据
              From 病人预交记录
              Where NO = r_冲预交.No And 记录性质 In (1, 11) And Rownum < 2;
          
            --更新预交单据余额
            Update 预交单据余额
            Set 预交余额 = Nvl(预交余额, 0) - n_退款金额
            Where 病人id = r_冲预交.病人id And 预交id = r_冲预交.Id
            Returning 预交余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 预交单据余额
                (预交id, 病人id, 预交类别, 预交余额)
              Values
                (r_冲预交.Id, r_冲预交.病人id, 1, -1 * n_退款金额);
              n_返回值 := -1 * n_退款金额;
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 预交单据余额 Where 预交id = r_冲预交.Id And Nvl(预交余额, 0) = 0;
            End If;
          
          End If;
        
          If n_剩余款 = 0 Then
            Exit;
          End If;
        End If;
      End Loop;
    
      If n_剩余款 <> 0 And Nvl(退款检查_In, 0) = 1 Then
        v_Err_Msg := '退款金额大于病人剩余预交余额。';
        Raise Err_Item;
      End If;
    
      n_退款金额 := -1 * (-1 * 金额_In - n_剩余款);
      If n_退款金额 <> 0 Then
        Update 病人预交记录 Set 结帐id = n_结帐id Where ID = Id_In;
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要, 缴款组id,
           预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 冲预交, 结算性质, 关联交易id, 交易时间, 交易人员, 校对标志, 预交电子票据)
        Values
          (病人预交记录_Id.Nextval, 单据号_In, Decode(操作状态_In, 0, 票据号_In, Null), 11, 1, 病人id_In,
           Decode(主页id_In, 0, Null, 主页id_In), Decode(科室id_In, 0, Null, 科室id_In), Null, 结算方式_In, 结算号码_In, v_Date, 缴款单位_In,
           单位开户行_In, 单位帐号_In, 操作员编号_In, 操作员姓名_In, 摘要_In, n_组id, 预交类别_In, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In,
           合作单位_In, n_结帐id, n_退款金额, Null, Id_In, 收款时间_In, 操作员姓名_In, 校对标志_In, n_预交电子票据);
        --更新预交单据余额
        Update 预交单据余额
        Set 预交余额 = Nvl(预交余额, 0) - n_退款金额
        Where 病人id = 病人id_In And 预交id = Id_In
        Returning 预交余额 Into n_返回值;
        If Sql%RowCount = 0 Then
          Insert Into 预交单据余额 (预交id, 病人id, 预交类别, 预交余额) Values (Id_In, 病人id_In, 1, -1 * n_退款金额);
          n_返回值 := -1 * n_退款金额;
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 预交单据余额 Where 预交id = Id_In And Nvl(预交余额, 0) = 0;
        End If;
      End If;
    
      If 金额_In < 0 And Nvl(强制退现_In, 0) = 0 Then
        Begin
          Select Nvl(预交余额, 0) - Nvl(费用余额, 0)
          Into n_病人余额
          From 病人余额
          Where 性质 = 1 And 病人id = 病人id_In And Nvl(类型, 0) = Nvl(预交类别_In, 0);
        Exception
          When Others Then
            Null;
        End;
        --余额退款要考虑三方预交是否支持退现
      
        For c_三方预交 In (Select a.预交id, a.预交类别, a.卡类别id, a.结算卡序号 As 消费接口id, Nvl(b.编码, c.编号) As 编码, Nvl(b.名称, c.名称) As 名称,
                              Decode(b.编码, Null, c.是否全退, b.是否全退) As 是否全退, Decode(b.编码, Null, c.是否退现, b.是否退现) As 是否退现,
                              a.卡号, a.交易流水号, a.交易说明, a.预交余额
                       From (Select a.预交类别, Nvl(a.卡类别id, 0) As 卡类别id, Nvl(a.结算卡序号, 0) As 结算卡序号, a.卡号, a.交易流水号, a.交易说明,
                                     Max(Decode(Sign(金额), -1, Decode(a.记录状态, 1, 0, 2, 0, ID), ID)) As 预交id,
                                     Nvl(Sum(金额), 0) - Nvl(Sum(Nvl(冲预交, 0)), 0) As 预交余额
                              From 病人预交记录 A
                              Where a.病人id = 病人id_In And (Nvl(a.结算卡序号, 0) <> 0 Or Nvl(卡类别id, 0) <> 0)
                              Group By a.预交类别, Nvl(a.卡类别id, 0), Nvl(a.结算卡序号, 0), a.卡号, a.交易流水号, a.交易说明
                              Having Nvl(Sum(金额), 0) - Nvl(Sum(Nvl(冲预交, 0)), 0) <> 0) A, 医疗卡类别 B, 消费卡类别目录 C
                       Where a.预交类别 = Nvl(预交类别_In, 0) And a.卡类别id = b.Id(+) And a.结算卡序号 = c.编号(+) And
                             Nvl(a.预交余额, 0) <> 0
                       Order By 编码, a.卡号, a.交易流水号, a.交易说明) Loop
        
          If Instr(',7,8,', ',' || v_性质 || ',') = 0 And Nvl(c_三方预交.是否退现, 0) = 0 And Nvl(c_三方预交.预交余额, 0) > 0 Then
            n_三方预交 := Nvl(n_三方预交, 0) + Nvl(c_三方预交.预交余额, 0);
          Elsif Instr(',7,8,', ',' || v_性质 || ',') > 0 Then
            If Nvl(c_三方预交.卡号, '0') = Nvl(卡号_In, '0') And Nvl(c_三方预交.交易流水号, '0') = Nvl(交易流水号_In, '0') And
               Nvl(c_三方预交.交易说明, '0') = Nvl(交易说明_In, '0') Then
              n_三方预交 := Nvl(n_三方预交, 0) + Nvl(c_三方预交.预交余额, 0);
            End If;
          End If;
        End Loop;
      
        If Instr(',7,8,', ',' || v_性质 || ',') > 0 And Nvl(n_三方预交, 0) < 0 And 操作类型_In = 3 Then
          v_Err_Msg := '退款金额大于病人三方预交金额。';
          Raise Err_Item;
        Elsif Nvl(n_病人余额, 0) < 0 And 退款检查_In = 1 Then
          v_Err_Msg := '退款金额大于病人剩余预交余额。';
          Raise Err_Item;
        Elsif Instr(',7,8,', ',' || v_性质 || ',') = 0 And Nvl(n_病人余额, 0) - Nvl(n_三方预交, 0) < 0 And 操作类型_In = 3 And
              退款检查_In = 1 Then
          v_Err_Msg := '退款金额大于病人剩余预交余额。';
          Raise Err_Item;
        End If;
      End If;
    End If;
    --病人余额(预交余额现收)
  
    Select Max(性质) Into v_性质 From 结算方式 Where 名称 = 结算方式_In;
  
    If Nvl(v_性质, 1) <> 5 Then
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) + 金额_In
      Where 性质 = 1 And 病人id = 病人id_In And Nvl(类型, 0) = Nvl(预交类别_In, 0)
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 性质, 类型, 预交余额, 费用余额)
        Values
          (病人id_In, 1, Nvl(预交类别_In, 0), 金额_In, 0);
        n_返回值 := 金额_In;
      End If;
      If Nvl(金额_In, 0) = 0 Then
        Delete From 病人余额 Where 病人id = 病人id_In And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
      End If;
    End If;
  End If;
  --结算完成才处理票据，消息等
  If 操作状态_In = 1 Then
    Return;
    --更新异常单据
  Elsif 操作状态_In = 2 Then
    If 操作类型_In = 3 Then
      Update 病人预交记录 Set 记录状态 = 1, 实际票号 = 票据号_In Where ID = Id_In Return 结帐id Into n_结帐id;
      Update 病人预交记录
      Set 实际票号 = 票据号_In
      Where NO = (Select NO From 病人预交记录 Where ID = Id_In) And 记录性质 = 11;
      Update 病人预交记录
      Set 收款时间 = v_Date, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In, 缴款组id = n_组id, 交易时间 = v_Date, 交易人员 = 操作员姓名_In,
          预交电子票据 = n_预交电子票据
      Where 结帐id = n_结帐id And Nvl(校对标志, 0) = 1;
      Update 病人预交记录 Set 校对标志 = Null Where 结帐id = n_结帐id;
      --自定义过程调用
      Zl_Custom_Balance_Update(Id_In);
    Else
      --更新并处理余额
      Update 病人预交记录
      Set 记录状态 = 1, 校对标志 = Null, 实际票号 = 票据号_In, 收款时间 = v_Date, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In, 缴款组id = n_组id,
          交易时间 = v_Date, 交易人员 = 操作员姓名_In, 科室id = Decode(科室id_In, 0, Null, 科室id_In), 金额 = 金额_In, 结算方式 = 结算方式_In,
          结算号码 = 结算号码_In, 缴款单位 = 缴款单位_In, 单位开户行 = 单位开户行_In, 单位帐号 = 单位帐号_In, 摘要 = 摘要_In, 卡类别id = 卡类别id_In,
          结算卡序号 = 结算卡序号_In, 卡号 = 卡号_In, 预交电子票据 = n_预交电子票据
      Where ID = Id_In;
      --自定义过程调用
      Zl_Custom_Balance_Update(Id_In);
    
    End If;
  End If;
  --处理票据
  If 票据号_In Is Not Null Then
    Select 票据打印内容_Id.Nextval Into v_打印id From Dual;
  
    --发出票据
    Insert Into 票据打印内容 (ID, 数据性质, NO) Values (v_打印id, 2, 单据号_In);
  
    Insert Into 票据使用明细
      (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
    Values
      (票据使用明细_Id.Nextval, 2, 票据号_In, 1, 1, 领用id_In, v_打印id, v_Date, 操作员姓名_In, 金额_In);
    --状态改动
    Update 票据领用记录
    Set 当前号码 = 票据号_In, 剩余数量 = Decode(Sign(剩余数量 - 1), -1, 0, 剩余数量 - 1), 使用时间 = Sysdate
    Where ID = Nvl(领用id_In, 0);
  End If;

  --相关汇总表处理

  --人员缴款余额(现收)
  If Nvl(更新交款余额_In, 1) = 1 Then
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) + 金额_In
    Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = 结算方式_In
    Returning 余额 Into n_返回值;
  
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (操作员姓名_In, 结算方式_In, 1, 金额_In);
      n_返回值 := 金额_In;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 人员缴款余额
      Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = 结算方式_In And Nvl(余额, 0) = 0;
    End If;
  End If;
  --对临时担保的处理
  Select Nvl(担保性质, 0) Into v_担保 From 病人信息 Where 病人id = 病人id_In;
  If v_担保 = 1 And Nvl(金额_In, 0) > 0 Then
    Update 病人信息
    Set 担保额 = Decode(Sign(Nvl(担保额, 0) - Nvl(金额_In, 0)), 1, Nvl(担保额, 0) - Nvl(金额_In, 0), Null),
        担保人 = Decode(Sign(Nvl(担保额, 0) - Nvl(金额_In, 0)), 1, 担保人, Null),
        担保性质 = Decode(Sign(Nvl(担保额, 0) - Nvl(金额_In, 0)), 1, 担保性质, Null)
    Where 病人id = 病人id_In;
  End If;
  If 操作类型_In <> 1 And 结帐id_In Is Null Then
    If 金额_In < 0 Then
      b_Message.Zlhis_Charge_006(Id_In, 单据号_In);
    Else
      b_Message.Zlhis_Charge_005(Id_In, 单据号_In);
    End If;
    --消息推送;
    Begin
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 11, Id_In;
    Exception
      When Others Then
        Null;
    End;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人预交记录_Insert;
/


Create Or Replace Procedure Zl_病人预交记录_Modify
(
  Id_In           病人预交记录.Id%Type,
  结算方式_In     病人预交记录.结算方式%Type,
  结算金额_In     病人预交记录.金额%Type,
  结算号码_In     病人预交记录.结算号码%Type,
  卡号_In         病人预交记录.卡号%Type,
  交易流水号_In   病人预交记录.交易流水号%Type := Null,
  交易说明_In     病人预交记录.交易说明%Type := Null,
  结算摘要_In     病人预交记录.摘要%Type,
  操作员姓名_In   病人预交记录.操作员姓名%Type,
  普通结算_In     Number := 0,
  预交电子票据_In 病人预交记录.预交电子票据%Type := Null,
  险类_In         保险结算记录.险类%Type := Null
) As
  --功能:根据三方接口返回信息更新预交记录
  --普通结算_In: 0-保存卡类别ID，1-不保存卡类别ID

  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  v_结算方式     病人预交记录.结算方式%Type;
  n_卡类别id     病人预交记录.卡类别id%Type;
  n_金额         病人预交记录.金额%Type;
  n_返回值       病人余额.预交余额%Type;
  n_差额         病人余额.预交余额%Type;
  n_病人id       病人预交记录.Id%Type;
  n_预交类别     病人预交记录.预交类别%Type;
  n_结帐id       病人预交记录.结帐id%Type;
  n_记录状态     病人预交记录.记录状态%Type;
  n_预交电子票据 Number(2);
  n_险类         保险结算记录.险类%Type;
Begin

  n_预交电子票据 := 预交电子票据_In;
  Begin
    Select 病人id, 结算方式, 卡类别id, 金额, 预交类别, 结帐id, 记录状态
    Into n_病人id, v_结算方式, n_卡类别id, n_金额, n_预交类别, n_结帐id, n_记录状态
    From 病人预交记录
    Where ID = Id_In;
  Exception
    When Others Then
      v_Err_Msg := '未找到结算数据，请检查！';
      Raise Err_Item;
  End;

  If n_预交电子票据 Is Null Then
    n_险类 := 险类_In;
    If 险类_In Is Null Then
      Select Nvl(Max(险类), 0) Into n_险类 From 保险结算记录 Where 记录id = Id_In And 性质 = 3;
    End If;
    n_预交电子票据 := Zl_Fun_Isstarteinvoice(2, n_险类, n_预交类别);
  End If;

  If Nvl(普通结算_In, 0) = 1 Then
    n_卡类别id := Null;
  End If;
  Update 病人预交记录
  Set 结算方式 = Nvl(结算方式_In, 结算方式), 金额 = Nvl(结算金额_In, 金额), 结算号码 = Nvl(结算号码_In, 结算号码), 摘要 = Nvl(结算摘要_In, 摘要),
      卡类别id = n_卡类别id, 卡号 = Nvl(卡号_In, 卡号), 交易流水号 = Nvl(交易流水号_In, 交易流水号), 交易说明 = Nvl(交易说明_In, 交易说明), 预交电子票据 = n_预交电子票据
  Where ID = Id_In;
  --自定义过程调用
  Zl_Custom_Balance_Update(Id_In);

  --病人余额及预交单据余额
  If Nvl(n_记录状态, 0) = 1 And Nvl(n_结帐id, 0) = 0 Then
    If Nvl(结算金额_In, 0) <> Nvl(n_金额, 0) Then
      n_差额 := Nvl(n_金额, 0) - Nvl(结算金额_In, 0);
      --更新预交单据余额
      Update 预交单据余额
      Set 预交余额 = Nvl(预交余额, 0) - n_差额
      Where 病人id = n_病人id And 预交id = Id_In
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 预交单据余额 (预交id, 病人id, 预交类别, 预交余额) Values (Id_In, n_病人id, 1, Nvl(结算金额_In, 0));
        n_返回值 := Nvl(结算金额_In, 0);
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 预交单据余额 Where 预交id = Id_In And Nvl(预交余额, 0) = 0;
      End If;
      --病人余额
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) - n_差额
      Where 性质 = 1 And 病人id = n_病人id And Nvl(类型, 0) = Nvl(n_预交类别, 0)
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 性质, 类型, 预交余额, 费用余额)
        Values
          (n_病人id, 1, Nvl(n_预交类别, 0), Nvl(结算金额_In, 0), 0);
        n_返回值 := Nvl(结算金额_In, 0);
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 病人余额 Where 病人id = n_病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
      End If;
    End If;
  End If;

  --人员缴款余额
  If Nvl(结算方式_In, Nvl(v_结算方式, '-')) <> Nvl(v_结算方式, '-') Then
    --原结算方式
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) - n_金额
    Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = v_结算方式
    Returning 余额 Into n_返回值;
  
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (操作员姓名_In, v_结算方式, 1, -1 * n_金额);
      n_返回值 := -1 * n_金额;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 人员缴款余额
      Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式 And Nvl(余额, 0) = 0;
    End If;
    --新结算方式
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) + 结算金额_In
    Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = 结算方式_In
    Returning 余额 Into n_返回值;
  
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (操作员姓名_In, 结算方式_In, 1, 结算金额_In);
      n_返回值 := 结算金额_In;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 人员缴款余额
      Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = 结算方式_In And Nvl(余额, 0) = 0;
    End If;
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人预交记录_Modify;
/


Create Or Replace Procedure Zl_病人发卡票据_Print
(
  No_In       Varchar2,
  票据号_In   票据使用明细.号码%Type,
  领用id_In   票据使用明细.领用id%Type,
  使用人_In   票据使用明细.使用人%Type,
  使用时间_In 票据使用明细.使用时间%Type,
  操作类型_In Number,
  票据张数_In Number := 1
) As
  --功能：处理医疗卡使用门诊票据
  --参数：
  --      NO_IN       =     发卡收费的单据号。格式为：A0000001
  --      票据号_IN   =     要使用的开始票据号。该票据号应该不为空，为空时不处理数据,退卡时传入原始票据号
  --      领用ID_IN   =     严格控制票据时，为使用票据的领用批次。非严格控制时，为NULL。
  --      票据张数_In =     实际所需的票据打印张数
  --      操作类型_In =     1-发卡；2-退卡；3-重打；4-补打；5-换卡
  --该游标用于票据范围判断
  Cursor c_Fact Is
    Select * From 票据领用记录 Where ID = Nvl(领用id_In, 0);
  r_Factrow c_Fact%RowType;

  v_收回id     票据打印内容.Id%Type;
  v_票据号     票据使用明细.号码%Type;
  v_当前票据号 票据使用明细.号码%Type;
  n_打印id     票据打印内容.Id%Type;

  n_票据金额 票据使用明细.票据金额%Type;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin

  --无票据号时,不用处理票据
  If 票据号_In Is Null Then
    Return;
  End If;

  --退卡
  If 操作类型_In = 2 Then
    Begin
      --从最后一次打印的内容中取
      Select ID
      Into v_收回id
      From (Select b.Id
             From 票据使用明细 A, 票据打印内容 B
             Where a.打印id = b.Id And a.性质 = 1 And a.原因 In (1, 3) And a.票种 = 1 And b.数据性质 = 5 And b.No = No_In And
                   Not Exists
              (Select 1 From 票据使用明细 B Where a.号码 = b.号码 And a.打印id = b.打印id And 性质 = 2)
             Order By a.使用时间 Desc)
      Where Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
  
    If v_收回id Is Not Null Then
      Insert Into 票据使用明细
        (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
        Select 票据使用明细_Id.Nextval, 1, 票据号_In, 2, 2, 领用id, 打印id, 使用时间_In, 使用人_In
        From 票据使用明细
        Where 打印id = v_收回id And 票种 = 1 And 性质 = 1;
    End If;
    Return;
  End If;

  --重打收回原始票据
  If 操作类型_In = 3 Or 操作类型_In = 5 Then
    Begin
      --从最后一次打印的内容中取
      Select ID
      Into v_收回id
      From (Select b.Id
             From 票据使用明细 A, 票据打印内容 B
             Where a.打印id = b.Id And a.性质 = 1 And a.原因 In (1, 3) And a.票种 = 1 And b.数据性质 = 5 And b.No = No_In And
                   Not Exists
              (Select 1 From 票据使用明细 B Where a.号码 = b.号码 And a.打印id = b.打印id And 性质 = 2)
             Order By a.使用时间 Desc)
      Where Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
  
    If v_收回id Is Not Null Then
      Begin
        Insert Into 票据使用明细
          (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
          Select 票据使用明细_Id.Nextval, 票种, 号码, 2, Decode(操作类型_In, 5, 2, 4), 领用id, 打印id, 使用时间_In, 使用人_In, 票据金额
          From 票据使用明细
          Where 打印id = v_收回id And 票种 = 1 And 性质 = 1;
      Exception
        When Others Then
          Null;
      End;
    End If;
  End If;

  --票据打印金额
  Select Nvl(Sum(实收金额), 0) Into n_票据金额 From 住院费用记录 Where 记录性质 = 5 And NO = No_In;

  Select 票据打印内容_Id.Nextval Into n_打印id From Dual;
  --生成单据的票据打印内容
  Insert Into 票据打印内容 (ID, 数据性质, NO) Values (n_打印id, 5, No_In);

  --并发出票据
  v_票据号 := 票据号_In;
  If Nvl(领用id_In, 0) <> 0 Then
    Open c_Fact;
    Fetch c_Fact
      Into r_Factrow;
    If c_Fact%RowCount = 0 Then
      v_Error := '无效的票据领用批次，无法完成挂号票据分配操作。';
      Close c_Fact;
      Raise Err_Custom;
    Elsif Nvl(r_Factrow.剩余数量, 0) < 票据张数_In Then
      v_Error := '当前批次的剩余数量不足' || 票据张数_In || '张，无法完成挂号票据分配操作。';
      Close c_Fact;
      Raise Err_Custom;
    End If;
  End If;
  For I In 1 .. 票据张数_In Loop
    --检查票据范围是否正确
    If Nvl(领用id_In, 0) <> 0 Then
      If Not (Upper(v_票据号) >= Upper(r_Factrow.开始号码) And Upper(v_票据号) <= Upper(r_Factrow.终止号码) And
          Length(v_票据号) = Length(r_Factrow.终止号码)) Then
        v_Error := '该单据需要打印多张票据,但票据号"' || v_票据号 || '"超出票据领用的号码范围！';
        Close c_Fact;
        Raise Err_Custom;
      End If;
    End If;
  
    --发出票据
    Insert Into 票据使用明细
      (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
    Values
      (票据使用明细_Id.Nextval, 1, v_票据号, 1, Decode(操作类型_In, 3, 3, 1), 领用id_In, n_打印id, 使用时间_In, 使用人_In, n_票据金额);
  
    v_当前票据号 := v_票据号;
    --下一个票据号
    v_票据号 := Zl_Incstr(v_票据号);
  End Loop;

  If Nvl(领用id_In, 0) <> 0 Then
    Update 票据领用记录
    Set 使用时间 = 使用时间_In, 当前号码 = v_当前票据号, 剩余数量 = Nvl(剩余数量, 0) - 票据张数_In
    Where ID = 领用id_In;
    Close c_Fact;
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人发卡票据_Print;
/


Create Or Replace Procedure Zl_三方接口配置_Set
(
  接口名_In 三方接口配置.接口名%Type,
  参数_In   三方接口配置.参数名%Type,
  参数值_In 三方接口配置.参数值%Type
) As
  v_Error Varchar2(255);
Begin
  If zl_To_Number(参数_In) <> 0 Then
    Update 三方接口配置 Set 参数值 = 参数值_In Where 接口名 = 接口名_In And 参数号 = zl_To_Number(参数_In);
  Else
    Update 三方接口配置 Set 参数值 = 参数值_In Where 接口名 = 接口名_In And 参数名 = 参数_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_三方接口配置_Set;
/

Create Or Replace Function Zl_三方接口配置_Get
(
  接口名_In 三方接口配置.接口名%Type,
  参数_In   三方接口配置.参数名%Type ) Return Varchar2  
 As 
  v_参数值  三方接口配置.参数值%Type;
Begin
  If zl_To_Number(参数_In) <> 0 Then
    SELECT  参数值 INTO v_参数值 From  三方接口配置 Where 接口名 = 接口名_In And 参数号 = zl_To_Number(参数_In);
  Else
    SELECT  参数值 INTO v_参数值 From  三方接口配置 Where 接口名 = 接口名_In And 参数名 = 参数_In;
  End If;
  return v_参数值;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_三方接口配置_Get;
/
Create Or Replace Procedure Zl_病人预交记录_转预交
(
  票据号_In     票据使用明细.号码%Type,
  病人id_In     病人预交记录.病人id%Type,
  主页id_In     病人预交记录.主页id%Type,
  科室id_In     病人预交记录.科室id%Type,
  金额_In       病人预交记录.金额%Type,
  操作员编号_In 病人预交记录.操作员编号%Type,
  操作员姓名_In 病人预交记录.操作员姓名%Type,
  收款时间_In   病人预交记录.收款时间%Type,
  领用id_In     票据使用明细.领用id%Type,
  预交类别_In   病人预交记录.预交类别%Type,
  摘要_In       病人预交记录.摘要%Type
) As
  ------------------------------------------------------------
  --预交类别_In:1-门诊转住院;2-住院转门诊
  ------------------------------------------------------------
  v_Err_Msg Varchar2(100);
  Err_Item Exception;

  v_打印id 票据打印内容.Id%Type;
  n_返回值 病人余额.预交余额%Type;

  n_Id       病人预交记录.Id%Type;
  v_No       病人预交记录.No%Type;
  n_金额     病人预交记录.金额%Type;
  n_预交     病人预交记录.金额%Type;
  l_票据no   t_StrList := t_StrList();
  n_预交类别 病人预交记录.预交类别%Type;

  n_组id         病人预交记录.缴款组id%Type;
  d_收款时间     病人预交记录.收款时间%Type;
  n_预交电子票据 Number(2);
  n_险类         病人信息.险类%Type;
  n_结算性质     Number(2);

  Procedure 病人预交记录_Insert
  (
    Id_In           病人预交记录.Id%Type,
    单据号_In       病人预交记录.No%Type,
    病人id_In       病人预交记录.病人id%Type,
    主页id_In       病人预交记录.主页id%Type,
    科室id_In       病人预交记录.科室id%Type,
    充值金额_In     病人预交记录.金额%Type,
    结算方式_In     病人预交记录.结算方式%Type,
    结算号码_In     病人预交记录.结算号码%Type,
    缴款单位_In     病人预交记录.缴款单位%Type,
    单位开户行_In   病人预交记录.单位开户行%Type,
    单位帐号_In     病人预交记录.单位帐号%Type,
    摘要_In         病人预交记录.摘要%Type,
    操作员编号_In   病人预交记录.操作员编号%Type,
    操作员姓名_In   病人预交记录.操作员姓名%Type,
    预交类别_In     病人预交记录.预交类别%Type := Null,
    卡类别id_In     病人预交记录.卡类别id%Type := Null,
    结算卡序号_In   病人预交记录.结算卡序号%Type := Null,
    卡号_In         病人预交记录.卡号%Type := Null,
    交易流水号_In   病人预交记录.交易流水号%Type := Null,
    交易说明_In     病人预交记录.交易说明%Type := Null,
    合作单位_In     病人预交记录.合作单位%Type := Null,
    收款时间_In     病人预交记录.收款时间%Type := Null,
    组id_In         病人预交记录.缴款组id%Type := Null,
    关联交易id_In   病人预交记录.关联交易id%Type := Null,
    预交电子票据_In 病人预交记录.预交电子票据%Type
  ) As
    n_返回值 病人余额.预交余额%Type;
  Begin
  
    Insert Into 病人预交记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要, 缴款组id,
       预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 冲预交, 结算性质, 校对标志, 关联交易id, 交易时间, 交易人员, 预交电子票据)
    Values
      (Id_In, 单据号_In, Null, 1, 1, 病人id_In, Decode(主页id_In, 0, Null, 主页id_In), Decode(科室id_In, 0, Null, 科室id_In),
       充值金额_In, 结算方式_In, 结算号码_In, 收款时间_In, 缴款单位_In, 单位开户行_In, 单位帐号_In, 操作员编号_In, 操作员姓名_In, 摘要_In, 组id_In, 预交类别_In,
       卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In, Null, Null, Null, Null, 关联交易id_In, 收款时间_In, 操作员姓名_In,
       预交电子票据_In);
  
    If Nvl(卡类别id_In, 0) <> 0 Then
      --自定义过程调用
      Zl_Custom_Balance_Update(Id_In);
    End If;
    --更新预交单据余额
    Insert Into 预交单据余额 (预交id, 病人id, 预交类别, 预交余额) Values (Id_In, 病人id_In, 预交类别_In, 充值金额_In);
  
    --病人余额(预交余额现收)
    Update 病人余额
    Set 预交余额 = Nvl(预交余额, 0) + 充值金额_In
    Where 性质 = 1 And 病人id = 病人id_In And Nvl(类型, 0) = Nvl(预交类别_In, 0)
    Returning 预交余额 Into n_返回值;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人余额
        (病人id, 性质, 类型, 预交余额, 费用余额)
      Values
        (病人id_In, 1, Nvl(预交类别_In, 0), 充值金额_In, 0);
      n_返回值 := 充值金额_In;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 病人余额 Where 病人id = 病人id_In And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
    End If;
    If 充值金额_In < 0 Then
      b_Message.Zlhis_Charge_006(Id_In, 单据号_In);
    Else
      b_Message.Zlhis_Charge_005(Id_In, 单据号_In);
    End If;
    --消息推送;
    Begin
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 11, Id_In;
    Exception
      When Others Then
        Null;
    End;
  End 病人预交记录_Insert;

  Procedure 病人预交记录_Strict
  (
    Id_In           病人预交记录.Id%Type,
    原预交id_In     病人预交记录.Id%Type,
    单据号_In       病人预交记录.No%Type,
    病人id_In       病人预交记录.病人id%Type,
    主页id_In       病人预交记录.主页id%Type,
    科室id_In       病人预交记录.科室id%Type,
    冲销金额_In     病人预交记录.金额%Type,
    结算方式_In     病人预交记录.结算方式%Type,
    结算号码_In     病人预交记录.结算号码%Type,
    缴款单位_In     病人预交记录.缴款单位%Type,
    单位开户行_In   病人预交记录.单位开户行%Type,
    单位帐号_In     病人预交记录.单位帐号%Type,
    摘要_In         病人预交记录.摘要%Type,
    操作员编号_In   病人预交记录.操作员编号%Type,
    操作员姓名_In   病人预交记录.操作员姓名%Type,
    预交类别_In     病人预交记录.预交类别%Type := Null,
    卡类别id_In     病人预交记录.卡类别id%Type := Null,
    结算卡序号_In   病人预交记录.结算卡序号%Type := Null,
    卡号_In         病人预交记录.卡号%Type := Null,
    交易流水号_In   病人预交记录.交易流水号%Type := Null,
    交易说明_In     病人预交记录.交易说明%Type := Null,
    合作单位_In     病人预交记录.合作单位%Type := Null,
    收款时间_In     病人预交记录.收款时间%Type := Null,
    关联交易id_In   病人预交记录.关联交易id%Type := Null,
    缴款组id_In     病人预交记录.缴款组id%Type := Null,
    预交电子票据_In 病人预交记录.预交电子票据%Type
  ) As
  
    n_返回值 病人余额.预交余额%Type;
    n_结帐id 病人结帐记录.Id%Type;
  Begin
  
    Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
  
    Insert Into 病人预交记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要, 缴款组id,
       预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 冲预交, 结算性质, 校对标志, 关联交易id, 交易时间, 交易人员, 附加标志, 预交电子票据)
    Values
      (Id_In, 单据号_In, Null, 1, 1, 病人id_In, Decode(主页id_In, 0, Null, 主页id_In), Decode(科室id_In, 0, Null, 科室id_In),
       冲销金额_In, 结算方式_In, 结算号码_In, 收款时间_In, 缴款单位_In, 单位开户行_In, 单位帐号_In, 操作员编号_In, 操作员姓名_In, 摘要_In, 缴款组id_In, 预交类别_In,
       卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In, n_结帐id, 冲销金额_In, Null, Null, 关联交易id_In, 收款时间_In, 操作员姓名_In,
       Decode(预交类别_In, 1, 2, 3), 预交电子票据_In);
  
    If Nvl(卡类别id_In, 0) <> 0 Then
      --自定义过程调用
      Zl_Custom_Balance_Update(Id_In);
    End If;
  
    Update 病人预交记录 Set 结帐id = n_结帐id, 冲预交 = 0 Where ID = 原预交id_In And 结帐id Is Null;
  
    Insert Into 病人预交记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要, 缴款组id,
       预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 冲预交, 结算性质, 校对标志, 关联交易id, 交易时间, 交易人员)
      Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号,
             操作员姓名, 摘要, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, n_结帐id, Round(-1 * 冲销金额_In, 2), 结算性质, 校对标志,
             关联交易id, 交易时间, 交易人员
      From 病人预交记录
      Where ID = 原预交id_In;
  
    --更新预交单据余额
    Update 预交单据余额
    Set 预交余额 = Nvl(预交余额, 0) + 冲销金额_In
    Where 病人id = 病人id_In And 预交id = 原预交id_In
    Returning 预交余额 Into n_返回值;
    If Sql%RowCount = 0 Then
      Insert Into 预交单据余额 (预交id, 病人id, 预交类别, 预交余额) Values (原预交id_In, 病人id_In, 1, 冲销金额_In);
      n_返回值 := 冲销金额_In;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 预交单据余额 Where 预交id = 原预交id_In And Nvl(预交余额, 0) = 0;
    End If;
  
    --病人余额(预交余额现收)
    Update 病人余额
    Set 预交余额 = Nvl(预交余额, 0) + 冲销金额_In
    Where 性质 = 1 And 病人id = 病人id_In And Nvl(类型, 0) = Nvl(预交类别_In, 0)
    Returning 预交余额 Into n_返回值;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人余额
        (病人id, 性质, 类型, 预交余额, 费用余额)
      Values
        (病人id_In, 1, Nvl(预交类别_In, 0), 冲销金额_In, 0);
      n_返回值 := 冲销金额_In;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 病人余额 Where 病人id = 病人id_In And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
    End If;
    If 冲销金额_In < 0 Then
      b_Message.Zlhis_Charge_006(Id_In, 单据号_In);
    Else
      b_Message.Zlhis_Charge_005(Id_In, 单据号_In);
    End If;
  
    --消息推送;
    Begin
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 11, Id_In;
    Exception
      When Others Then
        Null;
    End;
  
  End 病人预交记录_Strict;

Begin

  Select Nvl(Sum(Nvl(预交余额, 0)), 0) Into n_返回值 From 病人余额 Where 病人id = 病人id_In And 类型 = 预交类别_In;
  If Nvl(n_返回值, 0) - Nvl(金额_In, 0) < 0 Then
    v_Err_Msg := '[ZLSOFT]' || Case
                   When Nvl(预交类别_In, 0) = 1 Then
                    '门诊预交'
                   Else
                    '住院预交'
                 End || '余额不足![ZLSOFT]';
    Raise Err_Item;
  End If;

  n_组id := Zl_Get组id(操作员姓名_In);

  d_收款时间 := 收款时间_In;
  If d_收款时间 Is Null Then
    Select Sysdate Into d_收款时间 From Dual;
  End If;

  n_金额 := 金额_In;

  For v_预交 In (Select a.Id, a.结算方式, a.结算号码, a.卡类别id, a.结算卡序号, a.卡号, a.交易流水号, a.交易说明, a.合作单位, b.预交余额 As 金额, a.关联交易id,
                      a.缴款单位, a.单位开户行, a.单位帐号, a.预交电子票据
               From 病人预交记录 A, 预交单据余额 B
               Where a.Id = b.预交id And b.病人id = 病人id_In And b.预交类别 = 预交类别_In And Not Exists
                (Select 1 From 结算方式 Where a.结算方式 = 名称 And 性质 = 5)
               Order By a.收款时间) Loop
  
    Select 病人预交记录_Id.Nextval Into n_Id From Dual;
    Select Decode(预交类别_In, 1, 2, 1) Into n_预交类别 From Dual;
    v_No := Nextno(11);
    l_票据no.Extend;
    l_票据no(l_票据no.Count) := v_No;
  
    n_预交 := Nvl(v_预交.金额, 0);
    If n_金额 < Nvl(v_预交.金额, 0) Then
      n_预交 := n_金额;
    End If;
    n_金额 := n_金额 - n_预交;
    Select 性质 Into n_结算性质 From 结算方式 Where 名称 = v_预交.结算方式;
    If n_结算性质 = 3 Then
      Select To_Number(险类) Into n_险类 From 病人信息 Where 病人id = 病人id_In;
    Else
      n_险类 := 0;
    End If;
    n_预交电子票据 := Zl_Fun_Isstarteinvoice(2, n_险类, n_预交类别);
  
    --1.插入一条充值记录
    病人预交记录_Insert(n_Id, v_No, 病人id_In, 主页id_In, 科室id_In, n_预交, v_预交.结算方式, v_预交.结算号码, v_预交.缴款单位, v_预交.单位开户行, v_预交.单位帐号,
                  摘要_In, 操作员编号_In, 操作员姓名_In, n_预交类别, v_预交.卡类别id, v_预交.结算卡序号, v_预交.卡号, v_预交.交易流水号, v_预交.交易说明, v_预交.合作单位,
                  d_收款时间, n_组id, v_预交.关联交易id, n_预交电子票据);
  
    Update 病人预交记录 Set 实际票号 = 票据号_In Where ID = n_Id;
  
    v_No := Nextno(11);
    l_票据no.Extend;
    l_票据no(l_票据no.Count) := v_No;
  
    Select 病人预交记录_Id.Nextval Into n_Id From Dual;
    Select Decode(预交类别_In, 1, 1, 2) Into n_预交类别 From Dual;
    --2.冲原有预交
    病人预交记录_Strict(n_Id, v_预交.Id, v_No, 病人id_In, 主页id_In, 科室id_In, -1 * n_预交, v_预交.结算方式, v_预交.结算号码, v_预交.缴款单位,
                  v_预交.单位开户行, v_预交.单位帐号, 摘要_In, 操作员编号_In, 操作员姓名_In, n_预交类别, v_预交.卡类别id, v_预交.结算卡序号, v_预交.卡号, v_预交.交易流水号,
                  v_预交.交易说明, v_预交.合作单位, d_收款时间, v_预交.关联交易id, n_组id, v_预交.预交电子票据);
  
    Update 病人预交记录
    Set 实际票号 = 票据号_In
    Where NO = (Select NO From 病人预交记录 Where ID = n_Id) And 记录性质 In (1, 11);
  
    If n_金额 <= 0 Then
      Exit;
    End If;
  End Loop;

  If Nvl(n_金额, 0) <> 0 Then
    v_Err_Msg := '[ZLSOFT]' || Case
                   When Nvl(预交类别_In, 0) = 1 Then
                    '门诊预交'
                   Else
                    '住院预交'
                 End || '余额不足![ZLSOFT]';
    Raise Err_Item;
  End If;

  --处理票据
  If 票据号_In Is Not Null Then
    Select 票据打印内容_Id.Nextval Into v_打印id From Dual;
  
    --发出票据
    Insert Into 票据打印内容
      (ID, 数据性质, NO)
      Select v_打印id, 2, Column_Value From Table(l_票据no);
  
    Insert Into 票据使用明细
      (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
    Values
      (票据使用明细_Id.Nextval, 2, 票据号_In, 1, 1, 领用id_In, v_打印id, 收款时间_In, 操作员姓名_In, 金额_In);
    --状态改动
    Update 票据领用记录
    Set 当前号码 = 票据号_In, 剩余数量 = Decode(Sign(剩余数量 - 1), -1, 0, 剩余数量 - 1), 使用时间 = Sysdate
    Where ID = Nvl(领用id_In, 0);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人预交记录_转预交;
/

Create Or Replace Procedure Zl_病人预交记录_余额退款
(
  Id_In           病人预交记录.Id%Type,
  单据号_In       病人预交记录.No%Type,
  票据号_In       票据使用明细.号码%Type,
  病人id_In       病人预交记录.病人id%Type,
  主页id_In       病人预交记录.主页id%Type,
  科室id_In       病人预交记录.科室id%Type,
  金额_In         病人预交记录.金额%Type,
  结算方式_In     病人预交记录.结算方式%Type,
  结算号码_In     病人预交记录.结算号码%Type,
  缴款单位_In     病人预交记录.缴款单位%Type,
  单位开户行_In   病人预交记录.单位开户行%Type,
  单位帐号_In     病人预交记录.单位帐号%Type,
  摘要_In         病人预交记录.摘要%Type,
  操作员编号_In   病人预交记录.操作员编号%Type,
  操作员姓名_In   病人预交记录.操作员姓名%Type,
  领用id_In       票据使用明细.领用id%Type,
  预交类别_In     病人预交记录.预交类别%Type := Null,
  卡类别id_In     病人预交记录.卡类别id%Type := Null,
  结算卡序号_In   病人预交记录.结算卡序号%Type := Null,
  卡号_In         病人预交记录.卡号%Type := Null,
  关联交易id_In   病人预交记录.关联交易id%Type := Null,
  交易流水号_In   病人预交记录.交易流水号%Type := Null,
  交易说明_In     病人预交记录.交易说明%Type := Null,
  合作单位_In     病人预交记录.合作单位%Type := Null,
  收款时间_In     病人预交记录.收款时间%Type := Null,
  校对标志_In     病人预交记录.校对标志%Type := Null,
  结算信息_In     Varchar2 := Null,
  仅更新数据_In   Number := 0,
  操作状态_In     Number := 0,
  结帐id_In       病人预交记录.结帐id%Type := Null,
  结算序号_In     病人预交记录.结算序号%Type := Null,
  预交电子票据_In 病人预交记录.预交电子票据%Type := Null,
  险类_In         保险结算记录.险类%Type := Null
) As
  ----------------------------------------------
  --余额退款操作
  --结算信息_In:原预交ID|金额||....
  --仅更新数据_IN:0-表示需要插入预交记录及更新病人余额;1-表示只更新结算信息中的消费数据
  --操作状态_IN:0-表示完成结算;1-表示未完成结算;
  v_Err_Msg Varchar2(200);
  Err_Item Exception;

  n_打印id       票据打印内容.Id%Type;
  v_担保         病人信息.担保性质%Type;
  d_收款时间     Date;
  n_返回值       病人余额.预交余额%Type;
  n_组id         财务缴款分组.Id%Type;
  n_结帐id       病人预交记录.结帐id%Type;
  n_结算序号     病人预交记录.结算序号%Type;
  n_Count        Number(18);
  n_预交余额     病人余额.预交余额%Type;
  n_预交电子票据 Number(2);
  n_险类         保险结算记录.险类%Type;

Begin

  n_预交电子票据 := 预交电子票据_In;
  If n_预交电子票据 Is Null Then
    n_险类 := 险类_In;
    If 险类_In Is Null Then
      Select Nvl(Max(险类), 0) Into n_险类 From 保险结算记录 Where 记录id = Id_In And 性质 = 3;
    End If;
    n_预交电子票据 := Zl_Fun_Isstarteinvoice(2, n_险类, 预交类别_In);
  End If;
  n_组id := Zl_Get组id(操作员姓名_In);
  If 仅更新数据_In = 0 Then
    d_收款时间 := 收款时间_In;
    If d_收款时间 Is Null Then
      Select Sysdate Into d_收款时间 From Dual;
    End If;
    n_结算序号 := 结算序号_In;
    n_结帐id   := 结帐id_In;
    If Nvl(n_结帐id, 0) = 0 Then
      Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
    End If;
    If Nvl(n_结算序号, 0) = 0 Then
      n_结算序号 := -1 * n_结帐id;
    End If;
    --为了并发，先锁定病人余额(金额_In为负数)
    Update 病人余额
    Set 预交余额 = Nvl(预交余额, 0) + 金额_In
    Where 病人id = 病人id_In And 类型 = 预交类别_In And 性质 = 1
    Returning 预交余额 Into n_预交余额;
    If Sql%RowCount = 0 Then
      Insert Into 病人余额
        (病人id, 性质, 类型, 预交余额, 费用余额)
      Values
        (病人id_In, 1, Nvl(预交类别_In, 0), 金额_In, 0);
      n_预交余额 := 金额_In;
    End If;
  
    Insert Into 病人预交记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要, 缴款组id,
       预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 冲预交, 结算性质, 校对标志, 关联交易id, 交易时间, 交易人员, 附加标志, 预交电子票据)
    Values
      (Id_In, 单据号_In, Decode(操作状态_In, 0, 票据号_In, Null), 1, 0, 病人id_In, Decode(主页id_In, 0, Null, 主页id_In),
       Decode(科室id_In, 0, Null, 科室id_In), 金额_In, 结算方式_In, 结算号码_In, d_收款时间, 缴款单位_In, 单位开户行_In, 单位帐号_In, 操作员编号_In,
       操作员姓名_In, 摘要_In, n_组id, 预交类别_In, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In, n_结帐id, n_结算序号, Null,
       Null, 校对标志_In, Decode(Nvl(关联交易id_In, 0), 0, Id_In, 关联交易id_In), 收款时间_In, 操作员姓名_In, 1, n_预交电子票据);
  
    --更新预交单据余额
    Insert Into 预交单据余额 (预交id, 病人id, 预交类别, 预交余额) Values (Id_In, 病人id_In, 预交类别_In, 金额_In);
  End If;

  If 仅更新数据_In = 1 Then
    Select Max(结帐id), Max(收款时间), Max(1) Into n_结帐id, d_收款时间, n_Count From 病人预交记录 Where ID = Id_In;
    If n_Count = 0 Then
      v_Err_Msg := '未找到退款记录，请检查！';
      Raise Err_Item;
    End If;
  End If;

  If 结算信息_In Is Not Null Then
    Zl_病人预交记录_Relevance(病人id_In, Id_In, 结算信息_In, n_结帐id, 操作员编号_In, 操作员姓名_In, 收款时间_In, 校对标志_In, n_组id);
  End If;

  If 操作状态_In = 1 Then
    Return;
  End If;

  --更新记录状态1
  Update 病人预交记录
  Set 记录状态 = 1, 校对标志 = 0, 实际票号 = 票据号_In
  Where NO = 单据号_In And 记录性质 = 1 And 记录状态 = 0
  Returning 结帐id Into n_结帐id;
  If Sql%NotFound Then
    v_Err_Msg := '未找到指定的单据(' || 单据号_In || ',可能因为并发原因被他人退款，请检查！';
    Raise Err_Item;
  End If;
  Update 病人预交记录 Set 校对标志 = 0 Where 结帐id = n_结帐id And Nvl(校对标志, 0) <> 0;

  --处理票据
  If 票据号_In Is Not Null Then
    Select 票据打印内容_Id.Nextval Into n_打印id From Dual;
  
    --发出票据
    Insert Into 票据打印内容 (ID, 数据性质, NO) Values (n_打印id, 2, 单据号_In);
  
    Insert Into 票据使用明细
      (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
    Values
      (票据使用明细_Id.Nextval, 2, 票据号_In, 1, 1, 领用id_In, n_打印id, d_收款时间, 操作员姓名_In, 金额_In);
    --状态改动
    Update 票据领用记录
    Set 当前号码 = 票据号_In, 剩余数量 = Decode(Sign(剩余数量 - 1), -1, 0, 剩余数量 - 1), 使用时间 = Sysdate
    Where ID = Nvl(领用id_In, 0);
  
    Update 病人预交记录 Set 实际票号 = 票据号_In Where 病人id = 病人id_In And 记录性质 = 11 And NO = 单据号_In;
  
  End If;

  --人员缴款余额(现收)
  Update 人员缴款余额
  Set 余额 = Nvl(余额, 0) + 金额_In
  Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = 结算方式_In
  Returning 余额 Into n_返回值;

  If Sql%RowCount = 0 Then
    Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (操作员姓名_In, 结算方式_In, 1, 金额_In);
    n_返回值 := 金额_In;
  End If;
  If Nvl(n_返回值, 0) = 0 Then
    Delete From 人员缴款余额 Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = 结算方式_In And Nvl(余额, 0) = 0;
  End If;

  --对临时担保的处理
  Select Nvl(担保性质, 0) Into v_担保 From 病人信息 Where 病人id = 病人id_In;

  If v_担保 = 1 And Nvl(金额_In, 0) > 0 Then
    Update 病人信息
    Set 担保额 = Decode(Sign(Nvl(担保额, 0) - Nvl(金额_In, 0)), 1, Nvl(担保额, 0) - Nvl(金额_In, 0), Null),
        担保人 = Decode(Sign(Nvl(担保额, 0) - Nvl(金额_In, 0)), 1, 担保人, Null),
        担保性质 = Decode(Sign(Nvl(担保额, 0) - Nvl(金额_In, 0)), 1, 担保性质, Null)
    Where 病人id = 病人id_In;
  End If;

  If Nvl(n_预交余额, 0) = 0 Then
    Delete From 病人余额
    Where 病人id = 病人id_In And 类型 = 预交类别_In And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0 And 性质 = 1;
  End If;

  If 金额_In < 0 Then
    b_Message.Zlhis_Charge_006(Id_In, 单据号_In);
  Else
    b_Message.Zlhis_Charge_005(Id_In, 单据号_In);
  End If;

  --消息推送;
  Begin
    Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
      Using 11, Id_In;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人预交记录_余额退款;
/

Create Or Replace Package b_Einvoice_Request Is
  ------------------------------------------------------------------
  --电子票据业务处理
  --  1.Einvoice_Start-判断电子票据是否启用(返回:1-启用;0-未启用)
  --  2.EInvoice_Create-电子票据开具(返回1-成功;0-失败)
  --  3.Einvoice_Cancel_Check-电子票据作废前检查(返回:1-合法;0-不合法)
  --  4.Einvoice_Cancel-电子票据作废(返回1-成功;0-失败)
  ------------------------------------------------------------------
  --1.判断电子票据是否启用
  Function Einvoice_Start
  (
    业务场景_In Integer,
    险类_In     保险结算记录.险类%Type,
    类型_In     Integer := Null
  ) Return Number;

  --2.电子票据开具
  Function Einvoice_Create
  (
    业务场景_In  Integer,
    结算id_In    病人预交记录.结帐id%Type,
    冲销id_In    病人预交记录.结帐id%Type,
    错误信息_Out Out Varchar2
  ) Return Number;

  --3.电子票据作废检查
  Function Einvoice_Cancel_Check
  (
    业务场景_In  Integer,
    结算id_In    病人预交记录.结帐id%Type,
    错误信息_Out Out Varchar2
  ) Return Number;

  --4.电子票据作废
  Function Einvoice_Cancel
  (
    业务场景_In  Integer,
    结算id_In    病人预交记录.结帐id%Type,
    错误信息_Out Out Varchar2
  ) Return Number;
End b_Einvoice_Request;
/

Create Or Replace Package Body b_Einvoice_Request Is
  ------------------------------------------------------------------
  --电子票据业务处理
  --  1.Einvoice_Start-判断电子票据是否启用(返回:1-启用;0-未启用)
  --  2.EInvoice_Create-电子票据开具(返回1-成功;0-失败)
  --  3.Einvoice_Cancel_Check-电子票据作废前检查(返回:1-合法;0-不合法)
  --  4.Einvoice_Cancel-电子票据作废(返回1-成功;0-失败)
  ------------------------------------------------------------------

  Function Einvoice_Start
  (
    业务场景_In Integer,
    险类_In     保险结算记录.险类%Type,
    类型_In     Integer := Null
  ) Return Number Is
    ------------------------------------------------------------------
    --功能:判断电子票据是否启用
    --入参:业务场景_In-1-收费,2-预交,3-结帐,4-挂号;5-就诊卡
    --     类型_in:NULL-不区分类型;针对场合为结账及预交:1-门诊;2-住院;
    --出参:错误信息_Out-返回的错误信息
    --返回:1-启用;0-未启用
    -------------------------------------------------------------------
    v_包名称   电子票据类别.包名称%Type;
    v_Sql      Varchar2(1000);
    n_Return   Number(2);
    n_启用     Number(2);
    n_Err_Code Number(18);
    Err_Item   Exception;
    Err_Custom Exception;
  Begin
  
    If Nvl(业务场景_In, 0) = 2 And Nvl(类型_In, 0) = 1 Then
      --门诊预交，暂不支持
      Return 0;
    End If;
  
    Begin
      n_启用 := 1;
      Select 包名称 Into v_包名称 From 电子票据类别 Where 是否启用 = 1;
    Exception
      When Others Then
        n_启用 := 0;
    End;
    If n_启用 = 0 Or v_包名称 Is Null Then
      --未启用或无包名称，直接返回0，表示成功;
      Return 0;
    End If;
  
    n_Err_Code := Null;
    Begin
      v_Sql := 'begin :n_return:=' || v_包名称 || '.Einvoice_Start(:1,:2,:3); end;';
      Execute Immediate v_Sql
        Using Out n_Return, In 业务场景_In, 险类_In, 类型_In;
      Return n_Return;
    Exception
      When Others Then
        n_Err_Code := SQLCode;
    End;
    If n_Err_Code = -6550 Or n_Err_Code Is Null Then
      Return 0;
    End If;
    Return 0;
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Einvoice_Start;

  Function Einvoice_Create
  (
    业务场景_In  Integer,
    结算id_In    病人预交记录.结帐id%Type,
    冲销id_In    病人预交记录.结帐id%Type,
    错误信息_Out Out Varchar2
  ) Return Number Is
    ------------------------------------------------------------------
    --功能:电子票据开具
    --入参:业务场景_In-1-收费,2-预交,3-结帐,4-挂号;5-就诊卡
    --     结算ID_In-业务场景_In=2(预交)时：原预交ID,业务场景_In<>2(预交)时：原结帐ID
    --     冲销ID_In-业务场景_In=2(预交)时：退款的预交ID,业务场景_In<>2(预交)时：当前退费的结帐ID,部分退费时有效;
    --出参:错误信息_Out-返回的错误信息
    --返回:1-成功;0-失败
    -------------------------------------------------------------------
    v_包名称      电子票据类别.包名称%Type;
    v_Sql         Varchar2(1000);
    n_Return      Number(2);
    v_Err_Msg_Out Varchar2(4000);
    n_启用        Number(2);
    n_Err_Code    Number(18);
    Err_Item   Exception;
    Err_Custom Exception;
  Begin
  
    Begin
      n_启用 := 1;
      Select 包名称 Into v_包名称 From 电子票据类别 Where 是否启用 = 1;
    Exception
      When Others Then
        n_启用 := 0;
    End;
    If n_启用 = 0 Or v_包名称 Is Null Then
      --未启用或无包名称，直接返回1，表示成功;
      Return 1;
    End If;
  
    n_Err_Code := Null;
    Begin
      v_Sql := 'begin :n_return:=' || v_包名称 || '.EInvoice_Create(:1,:2,:3,:v_Err_Msg_out); end;';
      Execute Immediate v_Sql
        Using Out n_Return, In 业务场景_In, 结算id_In, 冲销id_In, Out v_Err_Msg_Out;
      错误信息_Out := v_Err_Msg_Out;
      Return n_Return;
    Exception
      When Others Then
        n_Err_Code    := SQLCode;
        v_Err_Msg_Out := SQLErrM;
    End;
    If n_Err_Code = -6550 Or n_Err_Code Is Null Then
      --没有此过程，返回true
      Return 1;
    End If;
    Raise Err_Item;
  
  Exception
    When Err_Item Then
      zl_ErrorCenter(n_Err_Code, v_Err_Msg_Out);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Einvoice_Create;

  Function Einvoice_Cancel_Check
  (
    业务场景_In  Integer,
    结算id_In    病人预交记录.结帐id%Type,
    错误信息_Out Out Varchar2
  ) Return Number Is
    ------------------------------------------------------------------
    --功能:电子票据作废
    --入参:业务场景_In-1-收费,2-预交,3-结帐,4-挂号;5-就诊卡
    --     结算ID_In-业务场景_In=2(预交)时：原预交ID,业务场景_In<>2(预交)时：原结帐ID 
    --出参:错误信息_Out-返回的错误信息
    --返回:1-成功;0-失败
    -------------------------------------------------------------------
    v_包名称      电子票据类别.包名称%Type;
    v_Sql         Varchar2(1000);
    n_Return      Number(2);
    v_Err_Msg_Out Varchar2(4000);
    n_启用        Number(2);
    n_Err_Code    Number(18);
    Err_Item   Exception;
    Err_Custom Exception;
  Begin
  
    If 业务场景_In = 2 Then
      --预交款
      Select Max(Nvl(预交电子票据, 0)) Into n_Return From 病人预交记录 Where 结帐id = 结算id_In;
    Else
      --非预交：收费、结帐、挂号及就诊卡
      Select Max(Nvl(是否电子票据, 0)) Into n_Return From 病人预交记录 Where 结帐id = 结算id_In;
    End If;
    If Nvl(n_Return, 0) = 0 Then
      --该记录是未启用电子票据的，直接返回1;
      Return 1;
    End If;
  
    Begin
      n_启用 := 1;
      Select 包名称 Into v_包名称 From 电子票据类别 Where 是否启用 = 1;
    Exception
      When Others Then
        n_启用 := 0;
    End;
  
    If n_启用 = 0 Or v_包名称 Is Null Then
      错误信息_Out := '电子票据未启用，请在窗口中进行退费或退款，在此处禁止发起电子票据冲红。';
      Return 0;
    End If;
  
    --检查是否换开过电子票据
    For c_电子票据 In (Select ID, 是否换开, 纸质发票号
                   From 电子票据使用记录
                   Where 票种 = 业务场景_In And 记录状态 = 1 And 结算id = 结算id_In) Loop
      --针对电子票据进行处理
      If Nvl(c_电子票据.是否换开, 0) = 1 Then
        --换开纸质发票号，禁止作废操作
        错误信息_Out := '已经换开纸质发票(' || c_电子票据.纸质发票号 || ')，在此处禁止发起电子票据冲红。';
        Return 0;
      End If;
      n_Err_Code := Null;
      Begin
        v_Sql := 'begin :n_return:=' || v_包名称 || '.Einvoice_Cancel_Check(:1,:2,:v_Err_Msg_out); end;';
        Execute Immediate v_Sql
          Using Out n_Return, In 业务场景_In, c_电子票据.Id, Out v_Err_Msg_Out;
        错误信息_Out := v_Err_Msg_Out;
        Return n_Return;
      Exception
        When Others Then
          n_Err_Code    := SQLCode;
          v_Err_Msg_Out := SQLErrM;
      End;
    
      If n_Err_Code = -6550 Then
        错误信息_Out := '电子票据未启用，请在窗口中进行退费或退款，在此处禁止发起电子票据冲红。';
        Return 0;
      End If;
      If Not n_Err_Code Is Null Then
        Raise Err_Item;
      End If;
    
    End Loop;
    Return 1;
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg_Out || '[ZLSOFT]');
    When Err_Item Then
      zl_ErrorCenter(n_Err_Code, v_Err_Msg_Out);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Einvoice_Cancel_Check;

  Function Einvoice_Cancel
  (
    业务场景_In  Integer,
    结算id_In    病人预交记录.结帐id%Type,
    错误信息_Out Out Varchar2
  ) Return Number Is
    ------------------------------------------------------------------
    --功能:电子票据作废
    --入参:业务场景_In-1-收费,2-预交,3-结帐,4-挂号;5-就诊卡
    --     结算ID_In-业务场景_In=2(预交)时：原预交ID,业务场景_In<>2(预交)时：原结帐ID 
    --出参:错误信息_Out-返回的错误信息
    --返回:1-成功;0-失败
    -------------------------------------------------------------------
    v_包名称      电子票据类别.包名称%Type;
    v_Sql         Varchar2(1000);
    n_Return      Number(2);
    v_Err_Msg_Out Varchar2(4000);
    n_启用        Number(2);
    n_Err_Code    Number(18);
    Err_Item   Exception;
    Err_Custom Exception;
  Begin
  
    If 业务场景_In = 2 Then
      --预交款
      Select Max(Nvl(预交电子票据, 0)) Into n_Return From 病人预交记录 Where 结帐id = 结算id_In;
    Else
      --非预交：收费、结帐、挂号及就诊卡
      Select Max(Nvl(是否电子票据, 0)) Into n_Return From 病人预交记录 Where 结帐id = 结算id_In;
    End If;
    If Nvl(n_Return, 0) = 0 Then
      --该记录是未启用电子票据的，直接返回1;
      Return 1;
    End If;
  
    Begin
      n_启用 := 1;
      Select 包名称 Into v_包名称 From 电子票据类别 Where 是否启用 = 1;
    Exception
      When Others Then
        n_启用 := 0;
    End;
  
    If n_启用 = 0 Or v_包名称 Is Null Then
      错误信息_Out := '电子票据未启用，请在窗口中进行退费或退款，在此处禁止发起电子票据冲红。';
      Return 0;
    End If;
  
    --检查是否换开过电子票据
    For c_电子票据 In (Select ID, 是否换开, 纸质发票号
                   From 电子票据使用记录
                   Where 票种 = 业务场景_In And 记录状态 = 1 And 结算id = 结算id_In) Loop
      --针对电子票据进行处理
      If Nvl(c_电子票据.是否换开, 0) = 1 Then
        --换开纸质发票号，禁止作废操作
        错误信息_Out := '已经换开纸质发票(' || c_电子票据.纸质发票号 || ')，在此处禁止发起电子票据冲红。';
        Return 0;
      End If;
      n_Err_Code := Null;
    
      --避免并发原因，还是需要先进行检查电子票据是否允许冲红。
      Begin
        v_Sql := 'begin :n_return:=' || v_包名称 || '.Einvoice_Cancel_Check(:1,:2,:v_Err_Msg_out); end;';
        Execute Immediate v_Sql
          Using Out n_Return, In 业务场景_In, c_电子票据.Id, Out v_Err_Msg_Out;
        错误信息_Out := v_Err_Msg_Out;
        Return n_Return;
      Exception
        When Others Then
          n_Err_Code    := SQLCode;
          v_Err_Msg_Out := SQLErrM;
      End;
    
      If n_Err_Code = -6550 Then
        错误信息_Out := '电子票据未启用，请在窗口中进行退费或退款，在此处禁止发起电子票据冲红。';
        Return 0;
      End If;
      If Not n_Err_Code Is Null Then
        Raise Err_Item;
      End If;
    
      --进行电子票据冲红处理
      Begin
        v_Sql := 'begin :n_return:=' || v_包名称 || '.Einvoice_Cancel(:1,:2,:v_Err_Msg_out); end;';
        Execute Immediate v_Sql
          Using Out n_Return, In 业务场景_In, c_电子票据.Id, Out v_Err_Msg_Out;
        错误信息_Out := v_Err_Msg_Out;
        Return n_Return;
      Exception
        When Others Then
          n_Err_Code    := SQLCode;
          v_Err_Msg_Out := SQLErrM;
      End;
    
      If n_Err_Code = -6550 Then
        错误信息_Out := '电子票据未启用，请在窗口中进行退费或退款，在此处禁止发起电子票据冲红。';
        Return 0;
      End If;
      If Not n_Err_Code Is Null Then
        Raise Err_Item;
      End If;
    End Loop;
    Return 1;
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg_Out || '[ZLSOFT]');
    When Err_Item Then
      zl_ErrorCenter(n_Err_Code, v_Err_Msg_Out);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Einvoice_Cancel;
End b_Einvoice_Request;
/


Create Or Replace Procedure Zl_Third_Regist
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:HIS挂号
  --入参:Xml_In:
  --<IN>
  --   <CZFS>3</CZFS>    //操作方式
  --   <CZJLID>1</CZJLID>    //出诊记录ID
  --   <HM>号码</HM>    //号码
  --   <HX>号序</HX>     //号序
  --   <JKFS>0</JKFS>  //缴款方式,0-挂号或预约缴款;1-预约不缴款
  --   <YYSJ>2014-10-21 </YYSJ>    //预约日期 YYYY-MM-DD,分时段非序号控制需要传入时间
  --   <JE>金额</JE>     //金额
  --   <HZDW>合作单位</HZDW>        //合作单位名称
  --   <YYFS>支付宝<YYFS>    //预约方式,如自助机，支付宝
  --   <BRID>病人ID</BRID>     //病人ID
  --   <SFZH>身份证号</SFZH>     //身份证号
  --   <XM>姓名</XM>            //姓名
  --   <BRLX></BRLX>             //医保病人类型
  --   <FB>普通</FB>               //病人费别，可以不传
  --   <JQM>机器名</JQM>            //机器名
  --   <JSMS>1</JSMS>          //结算模式：0-普通模式，1-异步结算模式
  --   <CZLX>0</CZLX>          //操作类型：结算模式为1时传入，0-开始结算，1-完成结算，2-回退结算
  --   <JZID>1</JZID>          //结帐ID，操作类型为1或2时传入
  --   <ZFBZH>支付宝公众号UserID</ZFBZH>
  --   <ZFBXCY>支付宝小程序UserID</ZFBXCY>
  --   <WXGZHID>微信公众号OpenID</WXGZH>
  --   <WXXCXID>微信小程序OpenID</WXXCXID>
  --   <JSLIST>          //结算列表，操作类型为2时可不传入
  --     <JS>            //结算信息，挂号非医保结算目前仅支持一个，结构与收费一致
  --       <JSKLB>结算卡类别</JSKLB>    //结算卡类别
  --       <JSKH>支付宝帐号</JSKH>           //结算卡号(支付宝帐号)
  --       <JYSM>交易说明</JYSM>            //说明，固定传支付宝
  --       <JYLSH>流水号</JYLSH>           //流水号，传订单号
  --       <JSFS>结算方式</JSFS>            //结算方式:现金、支票，如果是三方卡,可以传空
  --       <JSJE>结算金额</JSJE>            //结算金额
  --       <ZY>摘要</ZY>                  //摘要
  --       <SFCYJ></SFCYJ>              //是否冲预交，挂号目前不传
  --       <SFXFK></SFXFK>              //是否消费卡,挂号目前不传
  --       <EXPENDLIST>                 //扩展信息
  --         <EXPEND>
  --           <JYMC>交易名称</JYMC>        //交易名称
  --           <JYLR>交易内容<JYLR>         //交易内容
  --         </EXPEND>
  --         <EXPEND>
  --           ...
  --         </EXPEND>
  --       </EXPENDLIST>
  --     </JS>
  --   </JSLIST>
  --</IN>

  --出参:Xml_Out
  --<OUTPUT>
  --  <GHDH>挂号单号</GHDH>          //挂号单号
  --  <CZSJ>操作时间</CZSJ>          //HIS的登记时间
  --  <JZID>结帐ID</JZID>          //本次结帐ID
  --  <KPBZ>开票标志</KPBZ> //1-成功开具电子票据;0-未开票成功标志
  --  <URL>H5页面URL</URL>
  --  <NETURL>外网H5页面URL</NETURL>
  --  <FPTT>发票抬头</FPTT>        //病人姓名
  --  <FPH>发票号</FPH>             //发票编号
  --  <FPJE>发票金额</FPJE>        //100.00
  --  <KPRQ>开票日期</KPRQ>   //yyyy-mm-dd
  --  <ERROR><MSG>错误信息</MSG></ERROR>  //出错时返回
  --</ OUTPUT>
  --------------------------------------------------------------------------------------------------
  v_号码     挂号安排.号码%Type;
  n_号序     挂号序号状态.序号%Type;
  d_原始时间 Date;
  n_应收金额 门诊费用记录.应收金额%Type;
  v_预约方式 预约方式.名称%Type;
  v_合作单位 病人挂号记录.合作单位%Type;
  n_病人id   病人信息.病人id%Type;
  v_病人类型 病人信息.病人类型%Type;
  v_费别     门诊费用记录.费别%Type;
  v_机器名   挂号序号状态.机器名%Type;
  n_缴款方式 Number(3);
  n_记录id   临床出诊记录.Id%Type;
  v_身份证号 病人信息.身份证号%Type;
  v_姓名     门诊费用记录.姓名%Type;
  n_结算模式 Number(1); --0-普通模式，1-异步结算模式
  n_操作类型 Number(1); --结算模式为1时传入，0 - 开始结算，1 - 完成结算，2 - 回退结算

  v_Para     Varchar2(2000);
  n_挂号模式 Number(3);
  d_启用时间 Date;
  d_发生时间 Date;
  d_登记时间 Date;

  n_操作方式   Number;
  v_流水号     病人预交记录.交易流水号%Type;
  v_说明       门诊费用记录.摘要%Type;
  v_卡类别名称 医疗卡类别.名称%Type;
  v_结算卡号   病人预交记录.卡号%Type;
  v_No         病人挂号记录.No%Type;
  v_结算方式   医疗卡类别.结算方式%Type;

  n_结帐id   门诊费用记录.结帐id%Type;
  n_卡类别id 医疗卡类别.Id%Type;
  v_排班     挂号安排.周日%Type;
  n_安排id   挂号安排.Id%Type;
  n_计划id   挂号安排计划.Id%Type;
  n_预交id   病人预交记录.Id%Type;
  n_序号控制 挂号安排.序号控制%Type;
  v_星期     挂号安排限制.限制项目%Type;
  v_现金     结算方式.名称%Type;
  n_分时段   Number(3);
  v_结算内容 Varchar2(3000);
  v_保险结算 Varchar2(1000);
  n_Step     Number(2);

  v_卡类别     三方交易记录.类别%Type;
  n_冲预交     病人预交记录.冲预交%Type;
  n_关联交易id 病人预交记录.关联交易id%Type;

  v_Temp    Varchar2(32767); --临时XML
  x_Templet Xmltype; --模板XML

  n_Count     Number(3);
  n_Checkmzlg Number(2);
  v_Err_Msg   Varchar2(200);
  Err_Item    Exception;
  Err_Special Exception;
  n_是否电子票据       病人预交记录.是否电子票据%Type;
  v_支付宝公众号userid Varchar2(100);
  v_支付宝小程序userid Varchar2(100);
  v_微信公众号openid   Varchar2(100);
  v_微信小程序openid   Varchar2(100);
  n_开票标志           Number(2);
  v_患者姓名           电子票据使用记录.姓名%Type;
  v_发票编号           电子票据使用记录.号码%Type;
  v_开票日期           Varchar2(20);
  n_发票金额           电子票据使用记录.票据金额%Type;
  v_Url                电子票据使用记录.Url内网%Type;
  v_Url外网            电子票据使用记录.Url外网%Type;

Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/HM'), Extractvalue(Value(A), 'IN/HX'),
         To_Date(Extractvalue(Value(A), 'IN/YYSJ'), 'YYYY-MM-DD hh24:mi:ss'), Extractvalue(Value(A), 'IN/JE'),
         Extractvalue(Value(A), 'IN/YYFS'), Extractvalue(Value(A), 'IN/HZDW'), Extractvalue(Value(A), 'IN/BRID'),
         Extractvalue(Value(A), 'IN/BRLX'), Extractvalue(Value(A), 'IN/FB'), Extractvalue(Value(A), 'IN/JQM'),
         Extractvalue(Value(A), 'IN/JKFS'), Extractvalue(Value(A), 'IN/CZJLID'), Extractvalue(Value(A), 'IN/SFZH'),
         Extractvalue(Value(A), 'IN/XM'), Extractvalue(Value(A), 'IN/JSMS'), Extractvalue(Value(A), 'IN/CZLX'),
         Extractvalue(Value(A), 'IN/JZID'), Extractvalue(Value(A), 'IN/ZFBZH'), Extractvalue(Value(A), 'IN/ZFBXCY'),
         Extractvalue(Value(A), 'IN/WXGZHID'), Extractvalue(Value(A), 'IN/WXXCXID')
  Into v_号码, n_号序, d_原始时间, n_应收金额, v_预约方式, v_合作单位, n_病人id, v_病人类型, v_费别, v_机器名, n_缴款方式, n_记录id, v_身份证号, v_姓名, n_结算模式,
       n_操作类型, n_结帐id, v_支付宝公众号userid, v_支付宝小程序userid, v_微信公众号openid, v_微信小程序openid
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If Nvl(n_病人id, 0) = 0 And Not v_身份证号 Is Null And Not v_姓名 Is Null Then
    n_病人id := Zl_Third_Getpatiid(v_身份证号, v_姓名);
  End If;
  If Nvl(n_病人id, 0) = 0 Then
    v_Err_Msg := '无法确定病人信息,请检查!';
    Raise Err_Item;
  End If;

  If Not v_支付宝公众号userid Is Null Then
    Zl_病人信息从表_Update(n_病人id, Upper('支付宝公众号UserID'), v_支付宝公众号userid);
  End If;

  If Not v_支付宝小程序userid Is Null Then
    Zl_病人信息从表_Update(n_病人id, Upper('支付宝小程序UserID'), v_支付宝公众号userid);
  End If;

  If Not v_微信公众号openid Is Null Then
    Zl_病人信息从表_Update(n_病人id, Upper('微信公众号OpenID'), v_支付宝公众号userid);
  End If;

  If Not v_微信小程序openid Is Null Then
    Zl_病人信息从表_Update(n_病人id, Upper('微信小程序OpenID'), v_支付宝公众号userid);
  End If;

  If Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) <> 0 Then
    If Nvl(n_结帐id, 0) = 0 Then
      v_Err_Msg := '没有指定相关的结算数据！';
      Raise Err_Item;
    End If;
  
    Begin
      Select NO, 发生时间, 登记时间
      Into v_No, d_发生时间, d_登记时间
      From 门诊费用记录
      Where 记录性质 = 4 And Nvl(费用状态, 0) = 1 And 结帐id = n_结帐id And Rownum < 2;
    Exception
      When Others Then
        v_Err_Msg := '没有找到指定的相关数据，可能已被处理！';
        Raise Err_Item;
    End;
  End If;

  If Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) = 2 Then
    Zl_病人挂号记录_Cancel(n_结帐id);
  
    v_Temp := '<GHDH>' || v_No || '</GHDH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<CZSJ>' || To_Char(d_登记时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<JZID>' || n_结帐id || '</JZID>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    Xml_Out := x_Templet;
    Return;
  End If;

  v_Para     := zl_GetSysParameter(256);
  n_挂号模式 := Substr(v_Para, 1, 1);
  Begin
    d_启用时间 := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
  Exception
    When Others Then
      d_启用时间 := Null;
  End;

  If Nvl(n_结算模式, 0) = 0 Or Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) = 0 Then
    If Sysdate - 10 > Nvl(d_启用时间, Sysdate - 30) Then
      If n_挂号模式 = 1 And Nvl(d_原始时间, Sysdate) > Nvl(d_启用时间, Sysdate - 30) And n_记录id Is Null Then
        v_Err_Msg := '系统当前处于出诊表排班模式，传入的参数无法确定挂号安排，请重试！';
        Raise Err_Item;
      End If;
    Else
      If n_挂号模式 = 1 And Nvl(d_原始时间, Sysdate) > Nvl(d_启用时间, Sysdate - 30) And n_记录id Is Null Then
        Begin
          Select a.Id
          Into n_记录id
          From 临床出诊记录 A, 临床出诊号源 B
          Where a.号源id = b.Id And b.号码 = v_号码 And Nvl(d_原始时间, Sysdate) Between a.开始时间 And a.终止时间;
        Exception
          When Others Then
            v_Err_Msg := '系统当前处于出诊表排班模式，传入的参数无法确定挂号安排，请重试！';
            Raise Err_Item;
        End;
      End If;
    End If;
  
    n_Checkmzlg := To_Number(Nvl(zl_GetSysParameter(323), '0'));
    For c_交易记录 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                          Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                          Extractvalue(b.Column_Value, '/JS/SFCYJ') As 是否冲预交,
                          Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                          Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号, Extractvalue(b.Column_Value, '/JS/ZY') As 摘要
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
    
      --冲预交不需要三方交易锁
      If Nvl(c_交易记录.是否冲预交, 0) = 0 Then
        If c_交易记录.结算卡类别 Is Null Then
          v_卡类别 := c_交易记录.结算方式;
        Else
          Select Decode(Translate(Nvl(c_交易记录.结算卡类别, 'abcd'), '#1234567890', '#'), Null, 1, 0)
          Into n_Count
          From Dual;
          If Nvl(n_Count, 0) = 1 Then
            Select Max(名称) Into v_卡类别 From 医疗卡类别 Where ID = To_Number(c_交易记录.结算卡类别);
          Else
            Select Max(名称) Into v_卡类别 From 医疗卡类别 Where 名称 = c_交易记录.结算卡类别;
          End If;
        End If;
      
        If v_卡类别 Is Null Then
          v_Err_Msg := '不支持的结算方式,请检查！';
          Raise Err_Item;
        End If;
      
        --仅第一个结算方式才检查交易锁
        n_Step := Nvl(n_Step, 0) + 1;
        If Zl_Fun_三方交易记录_Locked(v_卡类别, c_交易记录.交易流水号, c_交易记录.结算卡号, c_交易记录.摘要, 4) = 0 And n_Step = 1 Then
          v_Err_Msg := '交易流水号为:' || c_交易记录.交易流水号 || '的交易正在进行中，不允许再次提交此交易!';
          Raise Err_Special;
        End If;
      Else
        If Nvl(n_Checkmzlg, 0) <> 0 Then
          Select Count(1)
          Into n_Count
          From 病案主页 A, 病人信息 B
          Where a.病人id = n_病人id And a.病人性质 = 1 And a.病人id = b.病人id And a.主页id = b.主页id And Nvl(b.在院, 0) = 1;
          If n_Count <> 0 Then
            v_Err_Msg := '门诊留观病人不能使用门诊预交！';
            Raise Err_Item;
          End If;
        End If;
      End If;
    End Loop;
  
    If v_病人类型 Is Not Null Then
      Select Count(1) Into n_Count From 病人类型 Where 名称 = v_病人类型;
      If n_Count = 0 Then
        v_Err_Msg := '没有发现为(' || v_病人类型 || ')的病人类型！';
        Raise Err_Item;
      End If;
      Update 病人信息 Set 病人类型 = Nvl(病人类型, v_病人类型) Where 病人id = n_病人id;
    End If;
  
    d_登记时间 := Sysdate;
    v_No       := Nextno(12);
    Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
  
    If n_记录id Is Null Then
      Select C2
      Into v_星期
      From Table(f_Str2List2('1:周日,2:周一,3:周二,4:周三,5:周四,6:周五,7:周六'))
      Where C1 = To_Char(d_原始时间, 'D');
    
      Begin
        Select ID
        Into n_计划id
        From (Select ID
               From 挂号安排计划
               Where 号码 = v_号码 And d_原始时间 Between Nvl(生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And 失效时间 And
                     审核时间 Is Not Null
               Order By 生效时间 Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          Select ID Into n_安排id From 挂号安排 Where 号码 = v_号码;
      End;
    
      d_发生时间 := d_原始时间;
      If Nvl(n_计划id, 0) <> 0 Then
        --从计划读取信息
        Select Decode(To_Char(d_原始时间, 'D'), '1', a.周日, '2', a.周一, '3', a.周二, '4', a.周三, '5', a.周四, '6', a.周五, '7', a.周六,
                       Null), Nvl(a.序号控制, 0)
        Into v_排班, n_序号控制
        From 挂号安排计划 A
        Where a.Id = n_计划id;
        Select Count(1) Into n_分时段 From 挂号计划时段 Where 星期 = v_星期 And 计划id = n_计划id And Rownum < 2;
      
        --合作单位检查
        If v_合作单位 Is Not Null Then
          Select Count(1)
          Into n_Count
          From 合作单位计划控制
          Where 计划id = n_计划id And 数量 = 0 And 合作单位 = v_合作单位;
          If n_Count = 1 Then
            v_Err_Msg := '传入的合作单位在此号码上被禁用！';
            Raise Err_Item;
          End If;
        End If;
      
        If n_分时段 = 1 And n_序号控制 = 0 Then
          Select 序号
          Into n_号序
          From 挂号计划时段
          Where 计划id = n_计划id And 星期 = v_星期 And To_Char(开始时间, 'hh24:mi:ss') = To_Char(d_发生时间, 'hh24:mi:ss');
        Else
          Begin
            Select To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || '' || To_Char(开始时间, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
            Into d_发生时间
            From 挂号计划时段
            Where 计划id = n_计划id And 星期 = v_星期 And 序号 = Nvl(n_号序, 0);
          Exception
            When Others Then
              If n_分时段 = 1 And n_序号控制 = 1 Then
                Select To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || '' || To_Char(Max(结束时间), 'hh24:mi:ss'),
                                'YYYY-MM-DD hh24:mi:ss')
                Into d_发生时间
                From 挂号计划时段
                Where 计划id = n_计划id And 星期 = v_星期;
              Else
                Select To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || '' || To_Char(开始时间, 'hh24:mi:ss'),
                                'YYYY-MM-DD hh24:mi:ss')
                Into d_发生时间
                From 时间段
                Where 时间段 = v_排班;
              End If;
              If d_发生时间 < d_登记时间 Then
                d_发生时间 := d_登记时间;
              End If;
          End;
        End If;
      Else
        --从安排读取信息
        Select Decode(To_Char(d_原始时间, 'D'), '1', b.周日, '2', b.周一, '3', b.周二, '4', b.周三, '5', b.周四, '6', b.周五, '7', b.周六,
                       Null), Nvl(b.序号控制, 0)
        Into v_排班, n_序号控制
        From 挂号安排 B
        Where b.Id = n_安排id;
        Select Count(1) Into n_分时段 From 挂号安排时段 Where 星期 = v_星期 And 安排id = n_安排id And Rownum < 2;
      
        --合作单位检查
        If v_合作单位 Is Not Null Then
          Select Count(1)
          Into n_Count
          From 合作单位安排控制
          Where 安排id = n_安排id And 数量 = 0 And 合作单位 = v_合作单位;
          If n_Count = 1 Then
            v_Err_Msg := '传入的合作单位在此号码上被禁用！';
            Raise Err_Item;
          End If;
        End If;
      
        If n_分时段 = 1 And n_序号控制 = 0 Then
          d_发生时间 := d_原始时间;
          Select 序号
          Into n_号序
          From 挂号安排时段
          Where 安排id = n_安排id And 星期 = v_星期 And To_Char(开始时间, 'hh24:mi:ss') = To_Char(d_发生时间, 'hh24:mi:ss');
        Else
          Begin
            Select To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || '' || To_Char(开始时间, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
            Into d_发生时间
            From 挂号安排时段
            Where 安排id = n_安排id And 星期 = v_星期 And 序号 = Nvl(n_号序, 0);
          Exception
            When Others Then
              If n_分时段 = 1 And n_序号控制 = 1 Then
                Select To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || '' || To_Char(Max(结束时间), 'hh24:mi:ss'),
                                'YYYY-MM-DD hh24:mi:ss')
                Into d_发生时间
                From 挂号安排时段
                Where 安排id = n_安排id And 星期 = v_星期;
              Else
                Select To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || '' || To_Char(开始时间, 'hh24:mi:ss'),
                                'YYYY-MM-DD hh24:mi:ss')
                Into d_发生时间
                From 时间段
                Where 时间段 = v_排班;
              End If;
              If d_发生时间 < d_登记时间 Then
                d_发生时间 := d_登记时间;
              End If;
          End;
        End If;
      End If;
    Else
      --出诊表排班模式
      Begin
        Select 开始时间 Into d_发生时间 From 临床出诊序号控制 Where 记录id = n_记录id And 序号 = n_号序;
      Exception
        When Others Then
          d_发生时间 := d_原始时间;
      End;
    End If;
  
    --先产生病人挂号记录和病人费用记录
    If Nvl(n_缴款方式, 0) = 0 Then
      If Trunc(d_发生时间) <> Trunc(Sysdate) Then
        n_操作方式 := 3;
      Else
        n_操作方式 := 1;
      End If;
    Else
      n_操作方式 := 2;
    End If;
    Zl_三方机构挂号_Insert(n_操作方式, n_病人id, v_号码, n_号序, v_No, Null, Null, Null, d_发生时间, d_登记时间, v_合作单位, n_应收金额, Null, Null,
                     Null, Null, v_预约方式, Null, Null, Null, 1, n_结帐id, 0, Null, Null, Null, 1, v_费别, Null, v_机器名, 1, 0,
                     n_记录id, 0, Null, 1, 0);
  End If;

  Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
  For r_结算 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                      Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号,
                      Extractvalue(b.Column_Value, '/JS/SFCYJ') As 是否冲预交,
                      Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                      Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                      Extractvalue(b.Column_Value, '/JS/JYSM') As 交易说明, Extractvalue(b.Column_Value, '/JS/JSJE') As 结算金额
               From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
  
    If Nvl(r_结算.是否冲预交, 0) = 0 Then
      If r_结算.结算方式 Is Null Then
        Begin
          Select b.结算方式, b.Id
          Into v_结算方式, n_卡类别id
          From 医疗卡类别 B
          Where b.名称 = r_结算.结算卡类别 And Rownum < 2;
        Exception
          When Others Then
            v_Err_Msg := '没有发现该结算卡的相关信息';
            Raise Err_Item;
        End;
        v_结算内容 := v_结算内容 || '|' || v_结算方式 || ',' || r_结算.结算金额 || ',,';
      Else
        v_结算方式 := r_结算.结算方式;
        Select Count(1) Into n_Count From 结算方式 Where 名称 = v_结算方式 And 性质 In (3, 4);
        If n_Count = 1 And r_结算.结算卡类别 Is Null Then
          v_保险结算 := v_保险结算 || '||' || v_结算方式 || '|' || r_结算.结算金额;
        Else
          v_结算内容 := v_结算内容 || '|' || v_结算方式 || ',' || r_结算.结算金额 || ',,';
        End If;
      End If;
    
      If r_结算.结算卡类别 Is Not Null Then
        v_结算内容   := v_结算内容 || '1,';
        v_卡类别名称 := r_结算.结算卡类别;
        v_结算卡号   := r_结算.结算卡号;
        v_流水号     := r_结算.交易流水号;
        v_说明       := r_结算.交易说明;
        If n_卡类别id Is Null Then
          Begin
            Select b.Id Into n_卡类别id From 医疗卡类别 B Where b.名称 = r_结算.结算卡类别 And Rownum < 2;
          Exception
            When Others Then
              v_Err_Msg := '没有发现该结算卡的相关信息';
              Raise Err_Item;
          End;
        End If;
      
        Select Decode(Translate(Nvl(r_结算.结算卡类别, 'abcd'), '#1234567890', '#'), Null, 1, 0)
        Into n_Count
        From Dual;
        If Nvl(n_Count, 0) = 1 Then
          Select Max(名称) Into v_卡类别 From 医疗卡类别 Where ID = To_Number(r_结算.结算卡类别);
        Else
          Select Max(名称) Into v_卡类别 From 医疗卡类别 Where 名称 = r_结算.结算卡类别;
        End If;
      
        If Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) = 0 Then
          Select Max(关联交易id)
          Into n_关联交易id
          From 病人预交记录
          Where 记录性质 Not In (1, 11) And 结帐id = n_结帐id And 卡类别id = n_卡类别id And Rownum < 2;
          If Nvl(n_关联交易id, 0) = 0 Then
            n_关联交易id := n_预交id;
          Else
            Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          End If;
        
          Insert Into 病人预交记录
            (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明,
             合作单位, 结算性质, 结算号码, 交易时间, 交易人员, 关联交易id, 校对标志)
            Select n_预交id, 4, 1, v_No, 病人id, v_结算方式, r_结算.结算金额, 登记时间, 操作员编号, 操作员姓名, 结帐id, Null, 缴款组id, n_卡类别id, Null,
                   v_结算卡号, v_流水号, v_说明, v_合作单位, 4, '三方接口挂号', 登记时间, 操作员姓名, n_关联交易id, 1
            From 门诊费用记录
            Where 记录性质 = 4 And 结帐id = n_结帐id And Rownum < 2;
        Else
          If Nvl(n_结算模式, 0) = 1 Then
            Delete From 病人预交记录 Where 记录性质 Not In (1, 11) And 结帐id = n_结帐id And n_卡类别id = n_卡类别id;
          End If;
          v_结算内容 := v_结算内容 || n_预交id;
        End If;
      Else
        v_结算内容 := v_结算内容 || '0,';
        v_卡类别   := r_结算.结算方式;
      End If;
    
      If Nvl(n_结算模式, 0) = 0 Or Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) = 1 Then
        Update 三方交易记录
        Set 业务结算id = n_结帐id
        Where 流水号 = r_结算.交易流水号 And 类别 = v_卡类别 And 业务类型 = 4;
      End If;
    Else
      n_冲预交 := r_结算.结算金额;
    End If;
  End Loop;

  If Nvl(n_结算模式, 0) = 0 Or Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) = 1 Then
    If v_结算内容 Is Not Null Then
      v_结算内容 := Substr(v_结算内容, 2);
    Else
      Begin
        Select 名称 Into v_现金 From 结算方式 Where 性质 = 1;
      Exception
        When Others Then
          v_现金 := '现金';
      End;
      v_结算内容 := v_现金 || ',' || 0 || ',,0';
    End If;
  
    If Nvl(n_缴款方式, 0) = 0 Then
      If Trunc(d_发生时间) <> Trunc(Sysdate) Then
        n_操作方式 := 3;
      Else
        n_操作方式 := 1;
      End If;
    Else
      n_操作方式 := 2;
    End If;
    Zl_三方机构挂号_Insert(n_操作方式, n_病人id, v_号码, n_号序, v_No, Null, v_结算内容, Null, d_发生时间, d_登记时间, v_合作单位, n_应收金额, Null, Null,
                     v_流水号, v_说明, v_预约方式, n_预交id, n_卡类别id, Null, 1, n_结帐id, 0, v_保险结算, n_冲预交, v_结算卡号, 1, v_费别, Null,
                     v_机器名, 1, 0, n_记录id, 0, Null, 1, 1);
  
    If Nvl(n_卡类别id, 0) <> 0 Then
      For c_扩展信息 In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As 交易名称,
                            Extractvalue(b.Column_Value, '/EXPEND/JYLR') As 交易内容
                     From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS/EXPENDLIST/EXPEND'))) B) Loop
        Zl_三方结算交易_Insert(n_卡类别id, 0, v_结算卡号, n_结帐id, c_扩展信息.交易名称 || '|' || c_扩展信息.交易内容, 0);
      End Loop;
    
    End If;
  
    --电子票据处理  
    n_是否电子票据 := b_Einvoice_Request.Einvoice_Start(4, Null);
    Update 病人预交记录 Set 是否电子票据 = n_是否电子票据 Where 结帐id = n_结帐id;
  
    If Nvl(n_是否电子票据, 0) = 1 Then
      --需要开具电子票据
      If b_Einvoice_Request.Einvoice_Create(4, n_结帐id, Null, v_Err_Msg) = 0 Then
        --电子票据开具成功
        Raise Err_Item;
      End If;
    
      Select Max(1), Max(姓名), Max(号码), Max(To_Char(生成时间, 'yyyy-mm-dd')), Max(Url内网), Max(Url外网), Max(票据金额)
      Into n_开票标志, v_患者姓名, v_发票编号, v_开票日期, v_Url, v_Url外网, n_发票金额
      From 电子票据使用记录
      Where 结算id = n_结帐id And 票种 = 4 And 记录状态 = 1;
    
      If v_患者姓名 Is Not Null Then
        v_姓名 := v_患者姓名;
      End If;
    End If;
  End If;
  v_Temp := '<GHDH>' || v_No || '</GHDH>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<CZSJ>' || To_Char(d_登记时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<JZID>' || n_结帐id || '</JZID>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<KPBZ>' || Nvl(n_开票标志, 0) || '</KPBZ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;

  v_Temp := '<URL>' || Nvl(v_Url, '') || '</URL>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;

  v_Temp := '<NETURL>' || Nvl(v_Url外网, '') || '</NETURL>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;

  v_Temp := '<FPTT>' || v_姓名 || '</FPTT>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;

  v_Temp := '<FPH>' || v_发票编号 || '</FPH>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<FPJE>' || Nvl(n_发票金额, 0) || '</FPJE>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<KPRQ>' || v_开票日期 || '</KPRQ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Xml_Out := x_Templet;

Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Err_Special Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20105, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Regist;
/


Create Or Replace Procedure Zl_Third_Registdelcheck
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:HIS退号检查
  --入参:Xml_In:
  --<IN>
  --  <GHDH>A000001</GHDH>    //挂号单号
  --  <JSKLB>支付宝</JSKLB>      //结算卡类别
  --  <JCFP>1</JCFP>            //检查发票
  --  <GHJE>20</GHJE>            //挂号金额
  --  <LSH>34563</LSH>           //交易流水号
  --  <JKFS>0</JKFS>             //缴款方式,0-挂号或预约缴款;1-预约不缴款
  --  <YYFS></YYFS>              //缴款方式=1时传入，预约的预约方式
  --  <XL></XL>                  //险类
  --</IN>

  --出参:Xml_Out
  --<OUTPUT>
  -- <ERROR><MSG></MSG></ERROR> //为空表示检查成功
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  v_卡类别     Varchar2(100);
  v_No         病人挂号记录.No%Type;
  n_挂号金额   门诊费用记录.实收金额%Type;
  v_结算方式   医疗卡类别.结算方式%Type;
  n_实收金额   门诊费用记录.实收金额%Type;
  v_交易流水号 病人预交记录.交易流水号%Type;
  n_存在       Number(3);
  v_Type       Varchar2(50);
  v_Temp       Varchar2(32767); --临时XML
  x_Templet    Xmltype; --模板XML

  n_已开医嘱 Number(2);
  n_检查发票 Number(3);
  n_是否打印 Number(3);
  n_缴款方式 Number(3);
  n_险类     病人信息.险类%Type;
  v_预约方式 病人挂号记录.预约方式%Type;
  v_收费单   门诊费用记录.No%Type;
  n_结帐id   门诊费用记录.结帐id%Type;
  v_Err_Msg  Varchar2(200);
  Err_Item Exception;
  n_Count Number;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/GHDH'), Extractvalue(Value(A), 'IN/JSKLB'), Extractvalue(Value(A), 'IN/GHJE'),
         Extractvalue(Value(A), 'IN/LSH'), To_Number(Extractvalue(Value(A), 'IN/JCFP')),
         To_Number(Extractvalue(Value(A), 'IN/JKFS')), Extractvalue(Value(A), 'IN/YYFS'),
         To_Number(Extractvalue(Value(A), 'IN/XL'))
  Into v_No, v_卡类别, n_挂号金额, v_交易流水号, n_检查发票, n_缴款方式, v_预约方式, n_险类
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  Select Max(收费单) Into v_收费单 From 病人挂号记录 Where NO = v_No;

  n_缴款方式 := Nvl(n_缴款方式, 0);
  If v_卡类别 Is Not Null And n_缴款方式 = 0 Then
    Select Nvl2(Translate(v_卡类别, '\1234567890', '\'), 'Char', 'Num') Into v_Type From Dual;
    If v_Type = 'Num' Then
      --传入的是卡类别ID
      Select 结算方式 Into v_结算方式 From 医疗卡类别 Where ID = To_Number(v_卡类别);
    Else
      --传入的是卡类别名称
      Select 结算方式 Into v_结算方式 From 医疗卡类别 Where 名称 = v_卡类别;
    End If;
    If Nvl(n_缴款方式, 0) = 0 Then
      If Nvl(n_险类, 0) = 0 Then
        Select Nvl(Max(1), 0)
        Into n_存在
        From 病人预交记录 A,
             (Select Distinct 结帐id
               From 门诊费用记录
               Where NO = v_No And 记录性质 = 4
               Union
               Select Distinct 结帐id
               From 住院费用记录
               Where NO = v_No And 记录性质 = 5
               Union
               Select Distinct 结帐id
               From 门诊费用记录
               Where NO = v_收费单 And 记录性质 = 1) B
        Where a.结帐id = b.结帐id And 结算方式 <> v_结算方式 And Mod(记录性质, 10) <> 1 And Rownum < 2;
      Else
        Select Nvl(Max(1), 0)
        Into n_存在
        From 病人预交记录 A,
             (Select Distinct 结帐id
               From 门诊费用记录
               Where NO = v_No And 记录性质 = 4
               Union
               Select Distinct 结帐id
               From 住院费用记录
               Where NO = v_No And 记录性质 = 5
               Union
               Select Distinct 结帐id
               From 门诊费用记录
               Where NO = v_收费单 And 记录性质 = 1) B, 结算方式 C
        Where a.结帐id = b.结帐id And 结算方式 <> v_结算方式 And Mod(记录性质, 10) <> 1 And a.结算方式 = c.名称 And c.性质 Not In (3, 4) And
              Rownum < 2;
        If n_存在 = 0 Then
          Select Nvl(Max(1), 0)
          Into n_存在
          From 保险结算记录 A,
               (Select Distinct 结帐id
                 From 门诊费用记录
                 Where NO = v_No And 记录性质 = 4
                 Union
                 Select Distinct 结帐id
                 From 住院费用记录
                 Where NO = v_No And 记录性质 = 5
                 Union
                 Select Distinct 结帐id
                 From 门诊费用记录
                 Where NO = v_收费单 And 记录性质 = 1) B
          Where a.记录id = b.结帐id And 险类 <> n_险类 And Rownum < 2;
        End If;
      End If;
      If n_存在 = 1 Then
        v_Err_Msg := '传入的挂号单据包含' || v_结算方式 || '以外的结算方式,无法退号!';
        Raise Err_Item;
      End If;
    Else
      Begin
        Select 1 Into n_存在 From 病人挂号记录 A Where a.No = v_No And a.预约方式 = v_预约方式 And Rownum < 2;
      Exception
        When Others Then
          n_存在 := 0;
      End;
      If n_存在 = 0 Then
        v_Err_Msg := '传入的挂号单据不是' || v_预约方式 || '预约的,无法退号!';
        Raise Err_Item;
      End If;
    End If;
  End If;

  If n_缴款方式 = 0 Then
    Select Sum(实收金额), Max(Decode(记录状态, 2, 0, 结帐id))
    Into n_实收金额, n_结帐id
    From 门诊费用记录
    Where NO = v_No And 记录性质 = 4;
    If Not v_收费单 Is Null Then
      Select Sum(实收金额) Into n_实收金额 From 门诊费用记录 Where NO = v_收费单 And 记录性质 = 1;
    End If;
    If n_实收金额 <> n_挂号金额 Then
      v_Err_Msg := '传入的退款金额与实际挂号金额不符，请检查!';
      Raise Err_Item;
    End If;
    --电子票据检查
    If b_Einvoice_Request.Einvoice_Cancel_Check(4, n_结帐id, v_Err_Msg) = 0 Then
      --失败后，直接抛错
      Raise Err_Item;
    End If;
  End If;

  --补充结算检查，已存在补结算数据的，不能退号
  Begin
    Select 1
    Into n_存在
    From 费用补充记录 A,
         (Select Distinct 结帐id
           From 门诊费用记录
           Where NO = v_No And 记录性质 = 4
           Union
           Select Distinct 结帐id
           From 住院费用记录
           Where NO = v_No And 记录性质 = 5
           Union
           Select Distinct 结帐id
           From 门诊费用记录
           Where NO = v_收费单 And 记录性质 = 1) B
    Where a.收费结帐id = b.结帐id And a.记录性质 = 1 And a.附加标志 = 1 And Nvl(a.费用状态, 0) <> 2 And Rownum < 2;
  Exception
    When Others Then
      n_存在 := 0;
  End;
  If n_存在 = 1 Then
    v_Err_Msg := '传入的挂号单据已经进行了二次结算,无法退号!';
    Raise Err_Item;
  End If;
  --医嘱检查，已经开过医嘱的，不能退号
  Begin
    Select Distinct 1 Into n_已开医嘱 From 病人医嘱记录 Where 挂号单 = v_No;
  Exception
    When Others Then
      n_已开医嘱 := 0;
  End;
  If n_已开医嘱 = 1 Then
    v_Err_Msg := '传入的挂号单据已经开过医嘱,无法退号!';
    Raise Err_Item;
  End If;
  If Nvl(n_检查发票, 0) = 1 Then
    Select Max(Decode(a.实际票号, Null, 0, 1)) Into n_是否打印 From 门诊费用记录 A Where NO = v_No And 记录性质 = 4;
    If Nvl(n_是否打印, 0) = 1 Then
      v_Err_Msg := '本次退号的单据已开发票,不能退费!';
      Raise Err_Item;
    End If;
    Select Max(Decode(a.实际票号, Null, 0, 1))
    Into n_是否打印
    From 门诊费用记录 A
    Where NO = v_收费单 And 记录性质 = 1;
    If Nvl(n_是否打印, 0) = 1 Then
      v_Err_Msg := '本次退号的单据已开发票,不能退费!';
      Raise Err_Item;
    End If;
  End If;

  Select Count(1) Into n_Count From 病人挂号记录 Where NO = v_No And 记录标志 = -1;
  If n_Count <> 0 Then
    v_Err_Msg := '本次退号的单据处于交易异常状态,不能再退费!';
    Raise Err_Item;
  End If;

  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Registdelcheck;
/

Create Or Replace Procedure Zl_Third_Registdel
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  -------------------------------------------------------------------------------------------------- 
  --功能:HIS退号 
  --入参:Xml_In: 
  --<IN> 
  --  <GHDH>A000001</GHDH>    //挂号单号 
  --  <JSKLB>支付宝</JSKLB>      //结算卡类别 
  --  <JCFP>1</JCFP>            //检查发票 
  --  <GHJE>20</GHJE>            //挂号金额 
  --  <JSMS>1</JSMS>          //结算模式：0-普通模式，1-异步结算模式
  --  <CZLX>0</CZLX>          //操作类型：结算模式为1时传入，0-开始结算，1-完成结算，2-回退结算
  --  <CXID>1</CXID>          //冲销结帐ID，操作类型为1或2时传入 
  --  <JKFS>0</JKFS>             //缴款方式,0-挂号或预约缴款;1-预约不缴款 
  --  <YYFS></YYFS>              //缴款方式=1时传入，预约的预约方式 
  --  <LSH>34563</LSH>           //交易流水号
  --</IN> 

  --出参:Xml_Out 
  --<OUTPUT> 
  --  <CZSJ>操作时间</CZSJ>          //HIS的登记时间 
  --  <YJZID>原结帐ID</YJZID> 
  --  <CXID>冲销ID</CXID> 
  --  <KPBZ>开票标志</KPBZ> //部分退才有效:1-成功开具电子票据;0-未开票成功标志
  --  <URL>H5页面URL</URL>
  --  <NETURL>外网H5页面URL</NETURL>
  --  <FPTT>发票抬头</FPTT>        //病人姓名
  --  <FPH>发票号</FPH>             //发票编号
  --  <FPJE>发票金额</FPJE>        //100.00
  --  <KPRQ>开票日期</KPRQ>   //yyyy-mm-dd
  --  <ERROR><MSG></MSG></ERROR> //为空表示取消挂号成功 
  --</OUTPUT> 
  -------------------------------------------------------------------------------------------------- 
  v_No         病人挂号记录.No%Type;
  v_卡类别     Varchar2(100);
  n_挂号金额   门诊费用记录.实收金额%Type;
  v_交易流水号 病人预交记录.交易流水号%Type;
  n_检查发票   Number(3);
  n_缴款方式   Number(3);
  v_预约方式   病人挂号记录.预约方式%Type;
  n_结算模式   Number(1); --0-普通模式，1-异步结算模式
  n_操作类型   Number(1); --结算模式为1时传入，0 - 开始结算，1 - 完成结算，2 - 回退结算

  v_操作员编号   门诊费用记录.操作员编号%Type;
  v_操作员姓名   门诊费用记录.操作员姓名%Type;
  v_结算方式     医疗卡类别.结算方式%Type;
  v_Type         Varchar2(50);
  n_已开医嘱     Number(2);
  n_是否打印     Number(3);
  n_结帐id       门诊费用记录.结帐id%Type;
  n_冲销id       门诊费用记录.结帐id%Type;
  n_挂号原结帐id 门诊费用记录.结帐id%Type;
  n_挂号冲销id   门诊费用记录.结帐id%Type;
  n_剩余金额     门诊费用记录.结帐id%Type;
  d_登记时间     Date;
  v_收费单       门诊费用记录.No%Type;
  n_病人id       门诊费用记录.病人id%Type;
  n_卡类别id     医疗卡类别.Id%Type;
  v_退费结算     Varchar2(1000);
  n_Temp         Number(18);
  n_预交id       病人预交记录.Id%Type;
  n_是否电子票据 病人预交记录.是否电子票据%Type;

  n_开票标志 Number(2);
  v_患者姓名 电子票据使用记录.姓名%Type;
  v_发票编号 电子票据使用记录.号码%Type;
  v_开票日期 Varchar2(20);
  n_发票金额 电子票据使用记录.票据金额%Type;
  v_Url      电子票据使用记录.Url内网%Type;
  v_Url外网  电子票据使用记录.Url外网%Type;

  v_Temp    Varchar2(32767); --临时XML 
  x_Templet Xmltype; --模板XML 

  n_Count   Number;
  v_Err_Msg Varchar2(200);
  Err_Item Exception;

Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/GHDH'), Extractvalue(Value(A), 'IN/JSKLB'), Extractvalue(Value(A), 'IN/GHJE'),
         Extractvalue(Value(A), 'IN/LSH'), To_Number(Extractvalue(Value(A), 'IN/JCFP')),
         Nvl(To_Number(Extractvalue(Value(A), 'IN/JKFS')), 0), Extractvalue(Value(A), 'IN/YYFS'),
         Extractvalue(Value(A), 'IN/JSMS'), Extractvalue(Value(A), 'IN/CZLX'), Extractvalue(Value(A), 'IN/CXID')
  Into v_No, v_卡类别, n_挂号金额, v_交易流水号, n_检查发票, n_缴款方式, v_预约方式, n_结算模式, n_操作类型, n_冲销id
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  Select Max(结帐id) Into n_结帐id From 门诊费用记录 Where NO = v_No And 记录性质 = 4 And 记录状态 In (1, 3);

  --先要对电子票据冲红处理
  If b_Einvoice_Request.Einvoice_Cancel(4, n_结帐id, v_Err_Msg) = 0 Then
    --电子票据作废失败 
    Raise Err_Item;
  End If;

  If Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) <> 0 Then
    If Nvl(n_冲销id, 0) = 0 And Nvl(n_结帐id, 0) <> 0 Then
      v_Err_Msg := '没有指定相关的结算数据！';
      Raise Err_Item;
    End If;
  
    Begin
      Select 登记时间 Into d_登记时间 From 病人挂号记录 Where NO = v_No And 记录状态 In (1, 3) And Rownum < 2;
    Exception
      When Others Then
        v_Err_Msg := '没有找到指定的相关结算数据，可能已被处理！';
        Raise Err_Item;
    End;
  End If;

  If Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) = 2 Then
    --删除结算数据
    Zl_病人挂号记录_Cancel(n_冲销id);
  
    v_Temp := '<CZSJ>' || To_Char(d_登记时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<YJZID>' || n_结帐id || '</YJZID>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<CXID>' || n_冲销id || '</CXID>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<KPBZ>' || 0 || '</KPBZ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<URL>' || '' || '</URL>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<NETURL>' || '' || '</NETURL>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<FPTT>' || '' || '</FPTT>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<FPH>' || '' || '</FPH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<FPJE>' || '' || '</FPJE>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<KPRQ>' || '' || '</KPRQ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    Xml_Out := x_Templet;
    Return;
  End If;

  --获取操作员信息
  v_操作员编号 := Zl_操作员信息(1);
  v_操作员姓名 := Zl_操作员信息(2);

  Select Max(收费单) Into v_收费单 From 病人挂号记录 Where NO = v_No;
  If Nvl(n_结算模式, 0) = 0 Or Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) = 0 Then
    If n_缴款方式 = 1 Then
      Select Count(1)
      Into n_Count
      From 门诊费用记录
      Where NO = v_No And 记录性质 = 4 And 结帐id Is Not Null And Rownum < 2;
      If n_Count = 0 Then
        Select Count(1)
        Into n_Count
        From 门诊费用记录
        Where NO In (Select /*+cardinality(B,10)*/
                      Column_Value
                     From Table(f_Str2List(v_收费单)) B) And 记录性质 = 1 And 结帐id Is Not Null And Rownum < 2;
      End If;
      If n_Count <> 0 Then
        v_Err_Msg := '传入的挂号单据不是预约挂号单,无法取消预约!';
        Raise Err_Item;
      End If;
    
      Select Count(1) Into n_Count From 病人挂号记录 A Where a.No = v_No And a.预约方式 = v_预约方式 And Rownum < 2;
      If n_Count = 0 Then
        v_Err_Msg := '传入的挂号单据不是' || v_预约方式 || '预约的,无法取消预约!';
        Raise Err_Item;
      End If;
    End If;
  
    If v_卡类别 Is Not Null And Nvl(n_缴款方式, 0) = 0 Then
      Select Nvl2(Translate(v_卡类别, '\1234567890', '\'), 'Char', 'Num') Into v_Type From Dual;
      If v_Type = 'Num' Then
        --传入的是卡类别ID 
        Select 结算方式, ID Into v_结算方式, n_卡类别id From 医疗卡类别 Where ID = To_Number(v_卡类别);
      Else
        --传入的是卡类别名称 
        Select 结算方式, ID Into v_结算方式, n_卡类别id From 医疗卡类别 Where 名称 = v_卡类别;
      End If;
    
      --要退的单据不是以该结算卡结算的，则禁止退号 
      Select Count(1)
      Into n_Count
      From 病人预交记录 A,
           (Select Distinct 结帐id
             From 门诊费用记录
             Where NO = v_No And 记录性质 = 4
             Union
             Select Distinct 结帐id
             From 住院费用记录
             Where NO = v_No And 记录性质 = 5
             Union
             Select Distinct 结帐id
             From 门诊费用记录
             Where NO In (Select /*+cardinality(B,10)*/
                           Column_Value
                          From Table(f_Str2List(v_收费单)) B) And 记录性质 = 1) B
      Where a.结帐id = b.结帐id And a.卡类别id = n_卡类别id And Rownum < 2;
      If n_Count = 0 Then
        v_Err_Msg := '传入的挂号单据不是' || v_结算方式 || '结算的,无法退号!';
        Raise Err_Item;
      End If;
    End If;
  
    --补充结算检查，已存在补结算数据的，不能退号 
    Select Count(1)
    Into n_Count
    From 费用补充记录 A,
         (Select Distinct 结帐id
           From 门诊费用记录
           Where NO = v_No And 记录性质 = 4
           Union
           Select Distinct 结帐id
           From 住院费用记录
           Where NO = v_No And 记录性质 = 5
           Union
           Select Distinct 结帐id
           From 门诊费用记录
           Where NO In (Select /*+cardinality(B,10)*/
                         Column_Value
                        From Table(f_Str2List(v_收费单)) B) And 记录性质 = 1) B
    Where a.收费结帐id = b.结帐id And a.记录性质 = 1 And a.附加标志 = 1 And Nvl(a.费用状态, 0) <> 2 And Rownum < 2;
    If n_Count = 1 Then
      v_Err_Msg := '传入的挂号单据已经进行了二次结算,无法退号!';
      Raise Err_Item;
    End If;
  
    --医嘱检查，已经开过医嘱的，不能退号 
    Select Count(1) Into n_已开医嘱 From 病人医嘱记录 Where 挂号单 = v_No And Rownum < 2;
    If n_已开医嘱 = 1 Then
      v_Err_Msg := '传入的挂号单据已经开过医嘱,无法退号!';
      Raise Err_Item;
    End If;
  
    If Nvl(n_检查发票, 0) = 1 Then
      Select Max(Decode(a.实际票号, Null, 0, 1)) Into n_是否打印 From 门诊费用记录 A Where NO = v_No And 记录性质 = 4;
      If Nvl(n_是否打印, 0) = 1 Then
        v_Err_Msg := '本次退号的单据已开发票,不能退费!';
        Raise Err_Item;
      End If;
    
      Select Max(Decode(a.实际票号, Null, 0, 1))
      Into n_是否打印
      From 门诊费用记录 A
      Where NO In (Select /*+cardinality(B,10)*/
                    Column_Value
                   From Table(f_Str2List(v_收费单)) B) And 记录性质 = 1;
      If Nvl(n_是否打印, 0) = 1 Then
        v_Err_Msg := '本次退号的单据已开发票,不能退费!';
        Raise Err_Item;
      End If;
    End If;
  
    d_登记时间 := Sysdate;
    Select 病人结帐记录_Id.Nextval Into n_冲销id From Dual;
    Zl_三方机构挂号_Delete(v_No, v_交易流水号, '移动平台退号', d_登记时间, Null, 1, 0, n_冲销id);
  
    If Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) = 0 And Nvl(n_卡类别id, 0) > 0 Then
      For c_记录 In (Select NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, -1 * 冲预交 As 冲预交, 合作单位, 卡类别id, 卡号, 交易流水号, 结算性质,
                          关联交易id
                   From 病人预交记录
                   Where 记录性质 = 4 And 记录状态 In (1, 3) And 结帐id = n_结帐id And 卡类别id = n_卡类别id) Loop
      
        Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易流水号, 交易说明,
           合作单位, 结算序号, 卡类别id, 卡号, 结算性质, 交易时间, 交易人员, 关联交易id, 校对标志)
          Select n_预交id, c_记录.No, c_记录.实际票号, c_记录.记录性质, 2, c_记录.病人id, c_记录.主页id, c_记录.科室id, c_记录.摘要, c_记录.结算方式, d_登记时间,
                 操作员编号, 操作员姓名, c_记录.冲预交, n_冲销id, 缴款组id, c_记录.交易流水号, '移动平台退号', c_记录.合作单位, n_冲销id, c_记录.卡类别id, c_记录.卡号,
                 c_记录.结算性质, d_登记时间, 操作员姓名, c_记录.关联交易id, 1
          From 门诊费用记录
          Where 记录性质 = 4 And 结帐id = n_冲销id And Rownum < 2;
      End Loop;
    End If;
  Else
    If v_卡类别 Is Not Null And Nvl(n_缴款方式, 0) = 0 Then
      Select Nvl2(Translate(v_卡类别, '\1234567890', '\'), 'Char', 'Num') Into v_Type From Dual;
      If v_Type = 'Num' Then
        Select ID Into n_卡类别id From 医疗卡类别 Where ID = To_Number(v_卡类别);
      Else
        Select ID Into n_卡类别id From 医疗卡类别 Where 名称 = v_卡类别;
      End If;
      Delete From 病人预交记录 Where 结帐id = n_冲销id And n_卡类别id = n_卡类别id;
    End If;
  End If;

  If Nvl(n_结算模式, 0) = 0 Or Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) = 1 Then
    Zl_三方机构挂号_Delete(v_No, v_交易流水号, '移动平台退号', d_登记时间, Null, 1, 1, n_冲销id);
    n_挂号原结帐id := n_结帐id;
    n_挂号冲销id   := n_冲销id;
  
    --同步处理划价单 
    If v_收费单 Is Not Null Then
      n_Temp := 0;
      For c_挂号 In (Select NO, Max(记录状态) As 记录状态, Max(病人id) As 病人id, Max(Decode(记录状态, 2, 0, 结帐id)) As 原结帐id,
                          Max(Decode(记录状态, 2, 结帐id, 0)) As 冲销id
                   From 门诊费用记录
                   Where NO In (Select /*+cardinality(B,10)*/
                                 Column_Value
                                From Table(f_Str2List(v_收费单)) B) And 记录性质 = 1) Loop
      
        If Nvl(c_挂号.记录状态, 0) = 0 Then
          Zl_门诊划价记录_Delete(c_挂号.No);
          n_结帐id := c_挂号.原结帐id;
          n_冲销id := c_挂号.冲销id;
        Elsif Nvl(c_挂号.记录状态, 0) = 1 Then
          If v_结算方式 Is Null Then
            v_Err_Msg := '本次挂号单据退款失败,请检查!';
            Raise Err_Item;
          End If;
          --先要对电子票据冲红处理
          If b_Einvoice_Request.Einvoice_Cancel(1, c_挂号.原结帐id, v_Err_Msg) = 0 Then
            --电子票据作废失败 
            Raise Err_Item;
          End If;
        
          Select 病人结帐记录_Id.Nextval Into n_冲销id From Dual;
          Zl_门诊收费记录_销帐(c_挂号.No, v_操作员编号, v_操作员姓名, Null, d_登记时间, Null, n_冲销id);
          v_退费结算 := v_结算方式 || '|' || -1 * n_挂号金额 || '|' || ' |' || ' ';
          Zl_门诊退费结算_Modify(2, n_病人id, n_冲销id, v_退费结算, 0, n_卡类别id, Null, v_交易流水号, Null, 0, 0, 0, 2);
        
          n_结帐id := c_挂号.原结帐id;
          n_冲销id := c_挂号.冲销id;
          n_Temp   := n_Temp + 1;
        Else
          n_结帐id := c_挂号.原结帐id;
          n_冲销id := c_挂号.冲销id;
        End If;
      
      End Loop;
    
      If n_Temp > 1 Then
        v_Err_Msg := '本次挂号存在多次收费，请先退费后再退号!';
        Raise Err_Item;
      End If;
    End If;
  
    --处理电子票据
    n_Count := 0;
    Select Sum(结帐金额) Into n_剩余金额 From 门诊费用记录 Where NO = v_No And 记录性质 = 4;
    Select Max(是否电子票据) Into n_是否电子票据 From 病人预交记录 Where 结帐id = n_挂号原结帐id;
  
    Update 病人预交记录 Set 是否电子票据 = n_是否电子票据 Where 结帐id = n_挂号冲销id;
  
    If Nvl(n_剩余金额, 0) <> 0 And Nvl(n_是否电子票据, 0) = 1 Then
      --部分结帐，需要重新开具电子票据
      If b_Einvoice_Request.Einvoice_Create(4, n_挂号原结帐id, n_挂号冲销id, v_Err_Msg) = 0 Then
        Raise Err_Item;
      End If;
    
      Select Max(1), Max(姓名), Max(号码), Max(To_Char(生成时间, 'yyyy-mm-dd')), Max(Url内网), Max(Url外网), Max(票据金额)
      Into n_开票标志, v_患者姓名, v_发票编号, v_开票日期, v_Url, v_Url外网, n_发票金额
      From 电子票据使用记录
      Where 结算id = n_结帐id And 票种 = 4 And 记录状态 = 1;
    
      n_Count := 1;
    End If;
  
    If Nvl(n_挂号原结帐id, 0) <> Nvl(n_结帐id, 0) Then
      --收费部分的处理
      Select Max(是否电子票据) Into n_是否电子票据 From 病人预交记录 Where 结帐id = n_结帐id;
      If Nvl(n_是否电子票据, 0) = 1 Then
      
        Update 病人预交记录 Set 是否电子票据 = n_是否电子票据 Where 结帐id = n_冲销id;
      
        Select Sum(结帐金额)
        Into n_剩余金额
        From 门诊费用记录
        Where NO In (Select Distinct NO From 门诊费用记录 Where 结帐id = n_结帐id) And Mod(记录性质, 10) = 1;
      
        If Nvl(n_剩余金额, 10) <> 0 Then
          --部分开具
          If b_Einvoice_Request.Einvoice_Create(1, n_结帐id, n_冲销id, v_Err_Msg) = 0 Then
            If Nvl(n_Count, 0) = 0 Then
              --挂号处理了电子票据，则不能抛错
              Raise Err_Item;
            End If;
          Else
            If Nvl(n_开票标志, 0) = 1 Then
              Select Max(1), Max(姓名), Max(号码), Max(To_Char(生成时间, 'yyyy-mm-dd')), Max(Url内网), Max(Url外网), Max(票据金额)
              Into n_开票标志, v_患者姓名, v_发票编号, v_开票日期, v_Url, v_Url外网, n_发票金额
              From 电子票据使用记录
              Where 结算id = n_结帐id And 票种 = 1 And 记录状态 = 1;
            
            End If;
          End If;
        End If;
      End If;
    End If;
  End If;

  v_Temp := '<CZSJ>' || To_Char(d_登记时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<YJZID>' || n_结帐id || '</YJZID>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<CXID>' || n_冲销id || '</CXID>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<KPBZ>' || Nvl(n_开票标志, 0) || '</KPBZ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<URL>' || Nvl(v_Url, '') || '</URL>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<NETURL>' || Nvl(v_Url外网, '') || '</NETURL>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<FPTT>' || v_患者姓名 || '</FPTT>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<FPH>' || v_发票编号 || '</FPH>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<FPJE>' || Nvl(n_发票金额, 0) || '</FPJE>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<KPRQ>' || v_开票日期 || '</KPRQ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Registdel;
/


Create Or Replace Procedure Zl_门诊收费记录_Insert
(
  No_In            门诊费用记录.No%Type,
  序号_In          门诊费用记录.序号%Type,
  病人id_In        门诊费用记录.病人id%Type,
  病人来源_In      Number,
  标识号_In        门诊费用记录.标识号%Type,
  付款方式_In      门诊费用记录.付款方式%Type,
  姓名_In          门诊费用记录.姓名%Type,
  性别_In          门诊费用记录.性别%Type,
  年龄_In          门诊费用记录.年龄%Type,
  费别_In          门诊费用记录.费别%Type,
  加班标志_In      门诊费用记录.加班标志%Type,
  病人科室id_In    门诊费用记录.病人科室id%Type,
  开单部门id_In    门诊费用记录.开单部门id%Type,
  开单人_In        门诊费用记录.开单人%Type,
  从属父号_In      门诊费用记录.从属父号%Type,
  收费细目id_In    门诊费用记录.收费细目id%Type,
  收费类别_In      门诊费用记录.收费类别%Type,
  计算单位_In      门诊费用记录.计算单位%Type,
  保险项目否_In    门诊费用记录.保险项目否%Type,
  保险大类id_In    门诊费用记录.保险大类id%Type,
  发药窗口_In      门诊费用记录.发药窗口%Type,
  付数_In          门诊费用记录.付数%Type,
  数次_In          门诊费用记录.数次%Type,
  附加标志_In      门诊费用记录.附加标志%Type,
  执行部门id_In    门诊费用记录.执行部门id%Type,
  价格父号_In      门诊费用记录.价格父号%Type,
  收入项目id_In    门诊费用记录.收入项目id%Type,
  收据费目_In      门诊费用记录.收据费目%Type,
  标准单价_In      门诊费用记录.标准单价%Type,
  应收金额_In      门诊费用记录.应收金额%Type,
  实收金额_In      门诊费用记录.实收金额%Type,
  统筹金额_In      门诊费用记录.统筹金额%Type,
  发生时间_In      门诊费用记录.发生时间%Type,
  登记时间_In      门诊费用记录.登记时间%Type,
  原no_In          门诊费用记录.No%Type,
  结帐id_In        门诊费用记录.结帐id%Type,
  收费结算_In      Varchar2,
  冲预交额_In      病人预交记录.冲预交%Type,
  保险结算_In      Varchar2,
  操作员编号_In    门诊费用记录.操作员编号%Type,
  操作员姓名_In    门诊费用记录.操作员姓名%Type,
  摘要_In          门诊费用记录.摘要%Type := Null,
  是否急诊_In      门诊费用记录.是否急诊%Type := 0,
  用法_In          药品收发记录.用法%Type := Null, --用法[|煎法]
  缴款_In          病人预交记录.缴款%Type := Null,
  找补_In          病人预交记录.找补%Type := Null,
  中药形态_In      门诊费用记录.结论%Type := Null,
  冲预交病人ids_In Varchar2 := Null,
  是否电子票据_In  病人预交记录.是否电子票据%Type := 0
) As
  --功能：新收一张门诊收费单据
  --参数：
  --  病人来源_IN:1-门诊;2-住院  住院病人收费时用。
  --  原NO_IN:修改保存新单据时用。目前用于存放于药品收发记录的摘要中。
  --         原单据(记录状态=2)记录修改产生的新单据号。
  --         新单据(记录状态=1)记录所修改的原单据号。
  -- 收费结算_IN:格式="结算方式|结算金额|结算号码|结算摘要||.....",注意无结算号码和摘要时要用空格填充
  -- 保险结算_IN:格式="结算方式|结算金额||....."
  -- 冲预交病人ids_In:多个用逗分分离,冲预交时有效(冲预交类别与业务操作保持一致),主要是使用家属的预交款
  -- 是否电子票据_In 是否使用电子票据
  v_费用id 门诊费用记录.Id%Type;

  v_用法 药品收发记录.用法%Type;
  v_煎法 药品收发记录.外观%Type;
  ------------------------------------------------------------
  --结算方式串
  v_结算内容 Varchar2(3000);
  v_当前结算 Varchar2(150);
  v_结算方式 病人预交记录.结算方式%Type;
  v_结算金额 病人预交记录.冲预交%Type;
  v_结算号码 病人预交记录.结算号码%Type;
  v_结算摘要 病人预交记录.摘要%Type;
  n_返回值   病人余额.费用余额%Type;

  v_Dec        Number;
  v_付款方式   医疗付款方式.名称%Type;
  v_费别性质   费别.属性%Type;
  n_新病人模式 Number;

  --临时变量
  Err_Custom Exception;
  v_Error       Varchar2(255);
  n_组id        财务缴款分组.Id%Type;
  n_单价小数    Number;
  v_Strtmpbefor Varchar2(4000);
  v_Msg         Varchar2(4000);

Begin
  n_组id := Zl_Get组id(操作员姓名_In);

  --金额小数位数
  Select zl_To_Number(Nvl(zl_GetSysParameter(9), '2')), zl_To_Number(Nvl(zl_GetSysParameter(157), '5'))
  Into v_Dec, n_单价小数
  From Dual;

  --门诊费用记录
  Select 病人费用记录_Id.Nextval Into v_费用id From Dual;
  Insert Into 门诊费用记录
    (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 门诊标志, 病人id, 标识号, 付款方式, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否,
     保险大类id, 付数, 数次, 发药窗口, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额, 记帐费用, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id,
     执行状态, 结帐id, 结帐金额, 操作员编号, 操作员姓名, 摘要, 是否急诊, 结论, 缴款组id)
  Values
    (v_费用id, 1, No_In, 1, 序号_In, Decode(从属父号_In, 0, Null, 从属父号_In), Decode(价格父号_In, 0, Null, 价格父号_In),
     Decode(病人来源_In, 1, 1, 2), Decode(病人id_In, 0, Null, 病人id_In), Decode(标识号_In, 0, Null, 标识号_In), 付款方式_In, 姓名_In, 性别_In,
     年龄_In, 病人科室id_In, 费别_In, 收费类别_In, 收费细目id_In, 计算单位_In, 保险项目否_In, 保险大类id_In, 付数_In, 数次_In, 发药窗口_In, 加班标志_In, 附加标志_In,
     收入项目id_In, 收据费目_In, 标准单价_In, 应收金额_In, 实收金额_In, 统筹金额_In, 0, 操作员姓名_In, 开单部门id_In, 开单人_In, 发生时间_In, 登记时间_In, 执行部门id_In,
     0, 结帐id_In, 实收金额_In, 操作员编号_In, 操作员姓名_In, 摘要_In, 是否急诊_In, 中药形态_In, n_组id);

  If 序号_In = 1 Then
    --病人预交记录(第一行时处理)
    --正常结算
    If 收费结算_In Is Not Null Then
      --各个收费结算
      v_结算内容 := 收费结算_In || '||';
      While v_结算内容 Is Not Null Loop
      
        v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
        v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      
        v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
        v_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      
        v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
        v_结算号码 := LTrim(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
        v_结算摘要 := LTrim(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
      
        If Nvl(v_结算金额, 0) <> 0 Then
          Insert Into 病人预交记录
            (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 结算号码, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款, 找补, 缴款组id, 结算性质)
          Values
            (病人预交记录_Id.Nextval, 3, No_In, 1, Decode(病人id_In, 0, Null, 病人id_In), Null, v_结算摘要, v_结算方式, v_结算号码, 登记时间_In,
             操作员编号_In, 操作员姓名_In, v_结算金额, 结帐id_In, Decode(v_结算内容, 收费结算_In || '||', 缴款_In, Null),
             Decode(v_结算内容, 收费结算_In || '||', 找补_In, Null), n_组id, 3);
        End If;
      
        v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
      End Loop;
    End If;
  
    --各个保险结算
    If 保险结算_In Is Not Null Then
      v_结算内容 := 保险结算_In || '||';
      While v_结算内容 Is Not Null Loop
        v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      
        v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
        v_结算金额 := To_Number(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
      
        If Nvl(v_结算金额, 0) <> 0 Then
          Insert Into 病人预交记录
            (ID, 记录性质, NO, 记录状态, 病人id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算性质)
          Values
            (病人预交记录_Id.Nextval, 3, No_In, 1, Decode(病人id_In, 0, Null, 病人id_In), '保险结算', v_结算方式, 登记时间_In, 操作员编号_In,
             操作员姓名_In, v_结算金额, 结帐id_In, n_组id, 3);
        End If;
        v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
      End Loop;
    End If;
  
    --预交结算
    If Nvl(冲预交额_In, 0) <> 0 Then
      Zl_病人预交记录_冲预交(病人id_In, 结帐id_In, 冲预交额_In, 1, 操作员编号_In, 操作员姓名_In, 登记时间_In, 冲预交病人ids_In, 3);
    End If;
  End If;

  Update 病人预交记录 Set 是否电子票据 = 是否电子票据_In Where 结帐id = 结帐id_In And 记录性质 <> 1;

  --相关汇总表的处理
  --汇总"人员缴款余额"(注意要处理个人帐户的结算)
  n_返回值 := 0;
  If 序号_In = 1 Then
    --各个收费结算
    If 收费结算_In Is Not Null Then
      v_结算内容 := 收费结算_In || '||';
      While v_结算内容 Is Not Null Loop
        v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
        v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
        v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
        v_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      
        If Nvl(v_结算金额, 0) <> 0 Then
          Update 人员缴款余额
          Set 余额 = Nvl(余额, 0) + Nvl(v_结算金额, 0)
          Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式
          Returning 余额 + n_返回值 Into n_返回值;
        
          If Sql%RowCount = 0 Then
            Insert Into 人员缴款余额
              (收款员, 结算方式, 性质, 余额)
            Values
              (操作员姓名_In, v_结算方式, 1, Nvl(v_结算金额, 0));
            n_返回值 := n_返回值 + Nvl(v_结算金额, 0);
          End If;
        End If;
      
        v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
      End Loop;
    End If;
  
    --各个保险结算
    If 保险结算_In Is Not Null Then
      v_结算内容 := 保险结算_In || '||';
      While v_结算内容 Is Not Null Loop
        v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      
        v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
        v_结算金额 := To_Number(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
      
        If Nvl(v_结算金额, 0) <> 0 Then
          Update 人员缴款余额
          Set 余额 = Nvl(余额, 0) + Nvl(v_结算金额, 0)
          Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式
          Returning 余额 + n_返回值 Into n_返回值;
        
          If Sql%RowCount = 0 Then
            Insert Into 人员缴款余额
              (收款员, 结算方式, 性质, 余额)
            Values
              (操作员姓名_In, v_结算方式, 1, Nvl(v_结算金额, 0));
            n_返回值 := n_返回值 + Nvl(v_结算金额, 0);
          End If;
        End If;
      
        v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
      End Loop;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 人员缴款余额 Where 性质 = 1 And 收款员 = 操作员姓名_In And Nvl(余额, 0) = 0;
    End If;
  End If;

  --药品和卫生材料部分
  If 收费类别_In In ('4', '5', '6', '7') Then
    --药品用法煎法分解
    If 用法_In Is Not Null Then
      If Instr(用法_In, '|') > 0 Then
        v_用法 := Substr(用法_In, 1, Instr(用法_In, '|') - 1);
        v_煎法 := Substr(用法_In, Instr(用法_In, '|') + 1);
      Else
        v_用法 := 用法_In;
      End If;
    End If;
    Zl_药品收发记录_销售出库(v_费用id, 原no_In, Null, Null, v_用法, v_煎法);
  End If;

  --更新部份病人信息
  If 序号_In = 1 And 病人id_In Is Not Null Then
    If 付款方式_In Is Not Null And Nvl(病人来源_In, 1) = 1 Then
      Select 名称 Into v_付款方式 From 医疗付款方式 Where 编码 = 付款方式_In;
    End If;
    Select Max(属性) Into v_费别性质 From 费别 Where 名称 = 费别_In; --2-动态费别不更新
  
    Select Zl_Fun_Checkidentify(0, 病人id_In, v_Strtmpbefor) Into v_Strtmpbefor From Dual;
    Update 病人信息
    Set 性别 = Decode(姓名, '新病人', Nvl(性别_In, 性别), 性别), 年龄 = Decode(姓名, '新病人', Nvl(年龄_In, 年龄), 年龄),
        姓名 = Decode(姓名, '新病人', 姓名_In, 姓名), 医疗付款方式 = Nvl(v_付款方式, 医疗付款方式), 费别 = Decode(v_费别性质, 1, 费别_In, 费别)
    Where 病人id = 病人id_In;
    Select Zl_Fun_Checkidentify(1, 病人id_In, v_Strtmpbefor) Into v_Msg From Dual;
    Select zl_To_Number(Nvl(zl_GetSysParameter('自动产生姓名', '1111'), '0')) Into n_新病人模式 From Dual;
    If n_新病人模式 = 1 Then
      Update 病人挂号记录
      Set 姓名 = 姓名_In, 性别 = 性别_In, 年龄 = 年龄_In
      Where 病人id = 病人id_In And 姓名 = '新病人';
      Update 门诊费用记录
      Set 姓名 = 姓名_In, 性别 = 性别_In, 年龄 = 年龄_In, 付款方式 = 付款方式_In
      Where 病人id = 病人id_In And 姓名 = '新病人';
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊收费记录_Insert;
/
Create Or Replace Procedure Zl_划价收费记录_Insert
(
  No_In            门诊费用记录.No%Type,
  病人id_In        门诊费用记录.病人id%Type,
  病人来源_In      Number,
  付款方式_In      门诊费用记录.付款方式%Type,
  姓名_In          门诊费用记录.姓名%Type,
  性别_In          门诊费用记录.性别%Type,
  年龄_In          门诊费用记录.年龄%Type,
  病人科室id_In    门诊费用记录.病人科室id%Type,
  开单部门id_In    门诊费用记录.开单部门id%Type,
  开单人_In        门诊费用记录.开单人%Type,
  收费结算_In      Varchar2,
  冲预交额_In      病人预交记录.冲预交%Type,
  保险结算_In      Varchar2,
  结帐id_In        门诊费用记录.结帐id%Type,
  发生时间_In      门诊费用记录.发生时间%Type,
  操作员编号_In    门诊费用记录.操作员编号%Type,
  操作员姓名_In    门诊费用记录.操作员姓名%Type,
  发药窗口_In      Varchar2,
  是否急诊_In      门诊费用记录.是否急诊%Type := 0,
  缴款_In          病人预交记录.缴款%Type := Null,
  找补_In          病人预交记录.找补%Type := Null,
  三方卡结算_In    Varchar2 := Null,
  登记时间_In      门诊费用记录.登记时间%Type := Null,
  结算序号_In      病人预交记录.结算序号%Type := Null,
  交易流水号_In    病人预交记录.交易流水号%Type := Null,
  交易说明_In      病人预交记录.交易说明%Type := Null,
  简单收费_In      Number := 0,
  冲预交病人ids_In Varchar2 := Null,
  是否电子票据_In  病人预交记录.是否电子票据%Type := 0
) As
  --功能：用于收费时收取划价单费用
  --参数：
  -- 发药窗口_In:执行部门ID1|发药窗口1;...;执行部门IDn|发药窗口n
  -- 病人来源_IN:1-门诊;2-住院
  -- 收费结算_IN:格式="结算方式|结算金额|结算号码|结算摘要||.....",注意无结算号码和摘要时要用空格填充
  -- 保险结算_IN:格式="结算方式|结算金额||....."
  -- 三方卡结算_In:格式=卡类别Id|是否消费卡|结算金额|卡号|备注||...
  -- 交易流水号_In和交易说明_In:收费结算_IN时有效.
  -- 冲预交病人ids_In:多个用逗分分离,冲预交时有效(冲预交类别与业务操作保持一致),主要是使用家属的预交款
  -- 是否电子票据_In 是否使用电子票据 
  --说明：
  --        1.收取划价费用时,才计算费用相关汇总,在划价时不处理;但药品相关汇总(姓名除外)划价时已经计算。
  --        2.收取划价费用时,目前界面及过程中未处理加收工本费,由划价时直接处理。  

  --=================================
  --备注：该过程目前只有简单收费使用！
  --=================================

  Cursor c_Price Is
    Select ID
    From 门诊费用记录
    Where NO = No_In And 记录性质 = 1 And 记录状态 = 0 And 操作员姓名 Is Null
    Order By 序号;

  n_Array_Size Number := 200;
  t_费用id     t_NumList;
  v_部门名称   部门表.名称%Type;

  --预交与结算相关变量
  v_结算内容 Varchar2(3000);
  v_当前结算 Varchar2(150);
  v_结算方式 病人预交记录.结算方式%Type;
  n_结算金额 病人预交记录.冲预交%Type;
  v_结算号码 病人预交记录.结算号码%Type;
  v_结算摘要 病人预交记录.摘要%Type;

  n_病人id   门诊费用记录.病人id%Type;
  v_标识号   门诊费用记录.标识号%Type;
  v_付款方式 医疗付款方式.名称%Type;
  n_返回值   病人余额.预交余额%Type;

  --临时变量
  n_Count      Number;
  n_新病人模式 Number;
  v_出库no     药品收发记录.No%Type;
  v_Date       Date;
  v_Err_Msg    Varchar2(255);
  Err_Item Exception;
  n_组id        财务缴款分组.Id%Type;
  n_卡类别id    医疗卡类别.Id%Type;
  n_消费卡      Number;
  v_卡号        病人预交记录.卡号%Type;
  v_卡名称      Varchar2(100);
  n_预交id      病人预交记录.Id%Type;
  n_消费卡id    消费卡信息.Id%Type;
  v_Strtmpbefor Varchar2(4000);
  v_Msg         Varchar2(4000);
Begin
  n_组id := Zl_Get组id(操作员姓名_In);

  Select Count(ID), Max(病人id)
  Into n_Count, n_病人id
  From 门诊费用记录
  Where 记录性质 = 1 And 记录状态 = 0 And NO = No_In And 操作员姓名 Is Null;
  If n_Count = 0 Then
    v_Err_Msg := '不能读取划价单内容,该单据可能已经删除或已经收费！';
    Raise Err_Item;
  End If;
  If Nvl(n_病人id, 0) <> 0 And Nvl(n_病人id, 0) <> Nvl(病人id_In, 0) Then
    v_Err_Msg := '单据【' || No_In || '】不是当前病人的费用，不能对其进行收费！';
    Raise Err_Item;
  End If;

  v_Date := 登记时间_In;
  If v_Date Is Null Then
    Select Sysdate Into v_Date From Dual;
  End If;
  If Nvl(病人id_In, 0) <> 0 Then
    Select Decode(当前科室id, Null, 门诊号, 住院号) Into v_标识号 From 病人信息 Where 病人id = 病人id_In;
  End If;

  ------------------------------------------------------------------------------------------------------------------------
  --批量更新
  Open c_Price;
  Loop
    Fetch c_Price Bulk Collect
      Into t_费用id Limit n_Array_Size;
    Exit When t_费用id.Count = 0;
  
    --循环处理门诊费用记录
    Forall I In 1 .. t_费用id.Count
    --执行状态相关字段不处理,在划价时处理;因为可能未收费发药,这种已执行的划价单是允许收费操作的。
    --为保证与预交结算记录的时间相同,重新填写登记时间,但药品部分不变动。
      Update 门诊费用记录
      Set 记录状态 = 1, 病人id = Decode(病人id_In, 0, Null, 病人id_In), 标识号 = v_标识号, 付款方式 = 付款方式_In, 姓名 = 姓名_In, 年龄 = 年龄_In,
          性别 = 性别_In,
          --可能保持医嘱发送的内容
          病人科室id = Nvl(病人科室id_In, 病人科室id), 开单部门id = Nvl(开单部门id_In, 开单部门id), 开单人 = Nvl(开单人_In, 开单人), 结帐金额 = 实收金额,
          结帐id = 结帐id_In, 发生时间 = 发生时间_In, 登记时间 = v_Date, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In, 是否急诊 = 是否急诊_In,
          缴款组id = n_组id
      Where ID = t_费用id(I) And 记录状态 = 0;
  
    If Sql%RowCount <> t_费用id.Count Then
      v_Err_Msg := '由于并发操作,该单据可能已经删除或已经收费！';
      Raise Err_Item;
    End If;
  
  End Loop;
  Close c_Price;
  ------------------------------------------------------------------------------------------------------------------------

  --预交款相关结算
  --收费结算
  If 收费结算_In Is Not Null Then
    v_结算内容 := 收费结算_In || '||';
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      v_结算号码 := LTrim(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_结算摘要 := LTrim(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
    
      If Nvl(n_结算金额, 0) <> 0 Then
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 结算号码, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款, 找补, 缴款组id, 结算序号, 交易流水号,
           交易说明, 结算性质)
        Values
          (病人预交记录_Id.Nextval, 3, No_In, 1, Decode(病人id_In, 0, Null, 病人id_In), Null, v_结算摘要, v_结算方式, v_结算号码, v_Date,
           操作员编号_In, 操作员姓名_In, n_结算金额, 结帐id_In, Decode(v_结算内容, 收费结算_In || '||', 缴款_In, Null),
           Decode(v_结算内容, 收费结算_In || '||', 找补_In, Null), n_组id, 结算序号_In, 交易流水号_In, 交易说明_In, 3);
      End If;
    
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  If 保险结算_In Is Not Null Then
    --各个保险结算
    v_结算内容 := 保险结算_In || '||';
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      n_结算金额 := To_Number(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
    
      If Nvl(n_结算金额, 0) <> 0 Then
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 结算性质)
        Values
          (病人预交记录_Id.Nextval, 3, No_In, 1, Decode(病人id_In, 0, Null, 病人id_In), '保险结算', v_结算方式, v_Date, 操作员编号_In,
           操作员姓名_In, n_结算金额, 结帐id_In, n_组id, 结算序号_In, 3);
      End If;
    
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  If 三方卡结算_In Is Not Null Then
    v_结算内容 := 三方卡结算_In || '||';
    While v_结算内容 Is Not Null Loop
      --卡类别Id|是否消费卡|结算金额|卡号|备注||...
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      n_卡类别id := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_消费卡   := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      v_卡号     := LTrim(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      v_结算摘要 := v_当前结算;
    
      If n_消费卡 = 1 Then
        Select 结算方式, 名称 Into v_结算方式, v_卡名称 From 消费卡类别目录 Where 编号 = n_卡类别id;
        If v_结算方式 Is Null Then
          v_Err_Msg := v_卡名称 || '未设置结算方式对照,请在消费卡中进行设置,结算失败！';
          Raise Err_Item;
        End If;
      Else
        Select 结算方式, 名称 Into v_结算方式, v_卡名称 From 医疗卡类别 Where ID = n_卡类别id;
        If v_结算方式 Is Null Then
          v_Err_Msg := v_卡名称 || '未设置结算方式对照,请在医疗卡管理中进行设置,结算失败！';
          Raise Err_Item;
        End If;
      End If;
      If Nvl(n_结算金额, 0) <> 0 Then
        Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 卡类别id, 结算卡序号, 结算方式, 结算号码, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款, 找补, 缴款组id,
           结算序号, 卡号, 结算性质)
        Values
          (n_预交id, 3, No_In, 1, Decode(病人id_In, 0, Null, 病人id_In), Null, v_结算摘要, Decode(n_消费卡, 1, Null, n_卡类别id),
           Decode(n_消费卡, 0, Null, n_卡类别id), v_结算方式, v_结算号码, v_Date, 操作员编号_In, 操作员姓名_In, n_结算金额, 结帐id_In, Null, Null,
           n_组id, 结算序号_In, v_卡号, 3);
      
        --卡结算对照
        If n_消费卡 = 1 Then
          Zl_病人卡结算记录_支付(n_卡类别id, v_卡号, n_消费卡id, n_结算金额, n_预交id, 操作员编号_In, 操作员姓名_In, v_Date);
        End If;
      End If;
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  --预交结算
  If Nvl(冲预交额_In, 0) <> 0 Then
    Zl_病人预交记录_冲预交(病人id_In, 结帐id_In, 冲预交额_In, 1, 操作员编号_In, 操作员姓名_In, 登记时间_In, 冲预交病人ids_In, 3, 1);
  End If;

  Update 病人预交记录 Set 是否电子票据 = 是否电子票据_In Where 结帐id = 结帐id_In And 记录性质 <> 1;

  --相关汇总表的处理

  --汇总"人员缴款余额"
  --收费结算
  n_返回值 := 0;
  If 收费结算_In Is Not Null Then
    v_结算内容 := 收费结算_In || '||';
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
    
      If Nvl(n_结算金额, 0) <> 0 Then
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) + Nvl(n_结算金额, 0)
        Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式
        Returning Nvl(余额, 0) + n_返回值 Into n_返回值;
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额
            (收款员, 结算方式, 性质, 余额)
          Values
            (操作员姓名_In, v_结算方式, 1, Nvl(n_结算金额, 0));
          n_返回值 := Nvl(n_返回值, 0) + Nvl(n_结算金额, 0);
        End If;
      End If;
    
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  --各个保险结算
  If 保险结算_In Is Not Null Then
    v_结算内容 := 保险结算_In || '||';
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      n_结算金额 := To_Number(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
    
      If Nvl(n_结算金额, 0) <> 0 Then
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) + Nvl(n_结算金额, 0)
        Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式
        Returning Nvl(余额, 0) + n_返回值 Into n_返回值;
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额
            (收款员, 结算方式, 性质, 余额)
          Values
            (操作员姓名_In, v_结算方式, 1, Nvl(n_结算金额, 0));
          n_返回值 := Nvl(n_返回值, 0) + Nvl(n_结算金额, 0);
        End If;
      End If;
    
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  If 三方卡结算_In Is Not Null Then
    v_结算内容 := 三方卡结算_In || '||';
    While v_结算内容 Is Not Null Loop
      --卡类别Id|是否消费卡|结算金额|卡号|备注||...
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      n_卡类别id := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_消费卡   := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      v_卡号     := Substr(v_当前结算, 1, Instr(v_当前结算, '|') + 1);
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      v_结算摘要 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
    
      If n_消费卡 = 1 Then
        Select 结算方式, 名称 Into v_结算方式, v_卡名称 From 消费卡类别目录 Where 编号 = n_卡类别id;
        If v_结算方式 Is Null Then
          v_Err_Msg := v_卡名称 || '未设置结算方式对照,请在消费卡中进行设置,结算失败！';
          Raise Err_Item;
        End If;
      Else
        Select 结算方式, 名称 Into v_结算方式, v_卡名称 From 医疗卡类别 Where ID = n_卡类别id;
        If v_结算方式 Is Null Then
          v_Err_Msg := v_卡名称 || '未设置结算方式对照,请在医疗卡管理中进行设置,结算失败！';
          Raise Err_Item;
        End If;
      End If;
    
      If Nvl(n_结算金额, 0) <> 0 Then
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) + Nvl(n_结算金额, 0)
        Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式
        Returning Nvl(余额, 0) + n_返回值 Into n_返回值;
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额
            (收款员, 结算方式, 性质, 余额)
          Values
            (操作员姓名_In, v_结算方式, 1, Nvl(n_结算金额, 0));
          n_返回值 := Nvl(n_返回值, 0) + Nvl(n_结算金额, 0);
        End If;
      End If;
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  If Nvl(n_返回值, 0) = 0 Then
    Delete From 人员缴款余额 Where 性质 = 1 And 收款员 = 操作员姓名_In And Nvl(余额, 0) = 0;
  End If;

  --药品部分非费用信息的修改
  --药品未发记录(如果已发药则修改不到),分离发药时无库房ID
  --可能存在材料和药品库房相同，但材料无发药窗口
  Update 未发药品记录
  Set 病人id = Decode(病人id_In, 0, Null, 病人id_In), 姓名 = 姓名_In, 对方部门id = 开单部门id_In, 已收费 = 1, 填制日期 = v_Date
  Where 单据 = 24 And NO = No_In And
        Nvl(库房id, 0) In (Select Distinct Nvl(执行部门id, 0)
                         From 门诊费用记录
                         Where 记录性质 = 1 And 记录状态 = 1 And NO = No_In And 收费类别 = '4');

  Update 未发药品记录
  Set 病人id = Decode(病人id_In, 0, Null, 病人id_In), 姓名 = 姓名_In, 对方部门id = 开单部门id_In, 已收费 = 1, 填制日期 = v_Date
  Where 单据 = 8 And NO = No_In And
        Nvl(库房id, 0) In (Select Distinct Nvl(执行部门id, 0)
                         From 门诊费用记录
                         Where 记录性质 = 1 And 记录状态 = 1 And NO = No_In And 收费类别 In ('5', '6', '7'));

  --药品收发记录(可能已经发药或取消发药,所有记录更改)
  Update 药品收发记录
  Set 对方部门id = 开单部门id_In, 填制日期 = Decode(Sign(Nvl(审核日期, v_Date) - v_Date), -1, 填制日期, v_Date)
  Where 单据 = 24 And NO = No_In And
        费用id + 0 In (Select ID From 门诊费用记录 Where 记录性质 = 1 And 记录状态 = 1 And NO = No_In And 收费类别 = '4');

  -------------------------------------------------------------------------------------------
  --处理备货卫材
  n_Count := Null;
  Begin
    Select Count(*), Max(a.No)
    Into n_Count, v_出库no
    From 药品收发记录 A, 门诊费用记录 B
    Where a.费用id = b.Id And b.收费类别 = '4' And b.记录性质 = 1 And b.记录状态 = 1 And
          Instr(',8,9,10,21,24,25,26,', ',' || a.单据 || ',') > 0 And b.No = No_In And Rownum <= 1;
  Exception
    When Others Then
      Null;
  End;
  If Nvl(n_Count, 0) > 0 Then
    If Nvl(病人科室id_In, 0) <> 0 Then
      Select 名称 Into v_部门名称 From 部门表 Where ID = 病人科室id_In;
    End If;
    v_Err_Msg := LPad(' ', 4);
    v_Err_Msg := Substr('病人姓名:' || 姓名_In || v_Err_Msg || '性别:' || 性别_In || v_Err_Msg || '年龄' || 年龄_In || v_Err_Msg ||
                        '门诊号:' || Nvl(v_标识号, '') || v_Err_Msg || '病人科室:' || v_部门名称, 1, 100);
  
    Update 药品收发记录
    Set 对方部门id = 开单部门id_In, 填制日期 = Decode(Sign(Nvl(审核日期, v_Date) - v_Date), -1, 填制日期, v_Date), 摘要 = v_Err_Msg
    Where 单据 = 21 And NO = v_出库no And
          费用id + 0 In (Select ID From 门诊费用记录 Where 记录性质 = 1 And 记录状态 = 1 And NO = No_In And 收费类别 = '4');
  End If;

  Update 药品收发记录
  Set 对方部门id = 开单部门id_In, 填制日期 = Decode(Sign(Nvl(审核日期, v_Date) - v_Date), -1, 填制日期, v_Date)
  Where 单据 = 8 And NO = No_In And
        费用id + 0 In
        (Select ID From 门诊费用记录 Where 记录性质 = 1 And 记录状态 = 1 And NO = No_In And 收费类别 In ('5', '6', '7'));
  If Not 发药窗口_In Is Null Then
    --更新发药窗口
    If Nvl(简单收费_In, 0) <> 0 Then
      Update 门诊费用记录
      Set 发药窗口 = 发药窗口_In
      Where NO = No_In And 记录性质 = 1 And 记录状态 = 1 And 收费类别 = 'Z';
    Else
      For v_窗口 In (Select To_Number(C1) As C1, C2 From Table(f_Str2List2(发药窗口_In, ';', '|'))) Loop
        Update 门诊费用记录
        Set 发药窗口 = Nvl(v_窗口.C2, 发药窗口)
        Where NO = No_In And 记录性质 = 1 And 记录状态 = 1 And 执行部门id = Nvl(v_窗口.C1, 执行部门id) And 收费类别 In ('5', '6', '7');
      
        Update 药品收发记录
        Set 发药窗口 = Nvl(v_窗口.C2, 发药窗口)
        Where 单据 = 8 And NO = No_In And 库房id = Nvl(v_窗口.C1, 库房id) And
              费用id + 0 In (Select ID
                           From 门诊费用记录
                           Where 记录性质 = 1 And 记录状态 = 1 And NO = No_In And 收费类别 In ('5', '6', '7'));
      
        Update 未发药品记录
        Set 发药窗口 = Nvl(v_窗口.C2, 发药窗口)
        Where 单据 = 8 And NO = No_In And 库房id = Nvl(v_窗口.C1, 库房id) And
              Nvl(库房id, 0) In (Select Distinct Nvl(执行部门id, 0)
                               From 门诊费用记录
                               Where 记录性质 = 1 And 记录状态 = 1 And NO = No_In And 收费类别 In ('5', '6', '7'));
      End Loop;
    End If;
  End If;

  --更新部份病人信息
  If 病人id_In Is Not Null Then
    If 付款方式_In Is Not Null And 病人来源_In = 1 Then
      Select 名称 Into v_付款方式 From 医疗付款方式 Where 编码 = 付款方式_In;
    End If;
  
    --通过划价单收费时不允许改费别,因为费用不允许变
  
    Select Zl_Fun_Checkidentify(0, 病人id_In, v_Strtmpbefor) Into v_Strtmpbefor From Dual;
    Update 病人信息
    Set 性别 = Nvl(性别_In, 性别), 年龄 = Nvl(年龄_In, 年龄), 姓名 = Decode(姓名, '新病人', 姓名_In, 姓名), 医疗付款方式 = Nvl(v_付款方式, 医疗付款方式)
    Where 病人id = 病人id_In;
    Select Zl_Fun_Checkidentify(1, 病人id_In, v_Strtmpbefor) Into v_Msg From Dual;
    Select zl_To_Number(Nvl(zl_GetSysParameter('自动产生姓名', '1111'), '0')) Into n_新病人模式 From Dual;
    If n_新病人模式 = 1 Then
    
      Update 病人挂号记录
      Set 姓名 = 姓名_In, 性别 = 性别_In, 年龄 = 年龄_In
      Where 病人id = 病人id_In And 姓名 = '新病人';
    
      Update 门诊费用记录
      Set 姓名 = 姓名_In, 性别 = 性别_In, 年龄 = 年龄_In, 付款方式 = 付款方式_In
      Where 病人id = 病人id_In And 姓名 = '新病人';
    End If;
  End If;
  --医嘱处理
  --场合_In    Integer:=0, --0:门诊;1-住院
  --性质_In    Integer:=1, --1-收费单;2-记帐单
  --操作_In    Integer:=0, --0:删除划价单;1-收费或记帐;2-退费或销帐
  --No_In      门诊费用记录.No%Type,
  --医嘱ids_In varchar2 := Null
  Zl_医嘱发送_计费状态_Update(0, 1, 1, No_In);

  --推送消息
  Begin
    Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
      Using 4, 结帐id_In;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_划价收费记录_Insert;
/
Create Or Replace Procedure Zl_门诊收费结算_Modify
(
  操作类型_In      Number,
  病人id_In        门诊费用记录.病人id%Type,
  结帐id_In        病人预交记录.结帐id%Type,
  结算方式_In      Varchar2,
  冲预交_In        病人预交记录.冲预交%Type := Null,
  退支票额_In      病人预交记录.冲预交%Type := Null,
  卡类别id_In      病人预交记录.卡类别id%Type := Null,
  卡号_In          病人预交记录.卡号%Type := Null,
  交易流水号_In    病人预交记录.交易流水号%Type := Null,
  交易说明_In      病人预交记录.交易说明%Type := Null,
  缴款_In          病人预交记录.缴款%Type := Null,
  找补_In          病人预交记录.找补%Type := Null,
  误差金额_In      门诊费用记录.实收金额%Type := Null,
  完成结算_In      Number := 0,
  缺省结算方式_In  结算方式.名称%Type := Null,
  冲预交病人ids_In Varchar2 := Null,
  更新交款余额_In  Number := 1,
  关联交易id_In    病人预交记录.关联交易id%Type := Null,
  删除原结算_In    Number := 0,
  校对标志_In      病人预交记录.校对标志%Type := 0,
  不控制会话_In    病人预交记录.会话号%Type := 0,
  是否电子票据_In  病人预交记录.是否电子票据%Type := 0
) As
  ------------------------------------------------------------------------------------------------------------------------------ 
  --功能:收费结算时,修改结算的相关信息 
  --操作类型_In: 
  --   0-普通收费方式: 
  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空. 
  --     ②退支票额_In:如果涉及退支票,则传入本次的退支票额,非正常收费时,传入零 
  --   1.三方卡结算: 
  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要" 
  --     ②退支票额_In:传入零 
  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入 
  --   2-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新) 
  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
  --     ②退支票额_In:传入零
  --   3-消费卡结算:
  --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."  消费卡ID:为零时,根据卡号自动定位 
  --     ②退支票额_In:传入零 
  --   4-三方卡结算，多种结算方式: 
  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要|单据号|是否普通结算|卡号" 
  --     ②退支票额_In:传入零 
  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入 

  -- 冲预交_In: 存在冲预交时,传入 
  -- 误差金额_In:存在误差费时,传入 
  -- 完成结算_In:1-完成收费;0-未完成收费 
  -- 冲预交病人ids_In:多个用逗分分离,冲预交时有效(冲预交类别与业务操作保持一致),主要是使用家属的预交款 
  -- 更新交款余额_In  是否更新人员交款余额，主要是处理统一操作员登录多台自助机的情况 
  -- 关联交易id_In 操作类型_In 为1,4时必须传入 
  -- 删除原结算_in 操作类型_In为4时有效，多个结算方式时调用多次该过程 
  -- 校对标志_In  操作类型_In为4时有效 
  -- 是否电子票据_In 是否使用电子票据，完成结算_In=1 时传入
  ------------------------------------------------------------------------------------------------------------------------------ 
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_结算内容 Varchar2(500);
  v_当前结算 Varchar2(50);
  v_卡号     病人医疗卡信息.卡号%Type;
  n_消费卡id 消费卡信息.Id%Type;
  v_名称     消费卡类别目录.名称%Type;
  n_卡类别id 病人预交记录.结算卡序号%Type;
  n_预交id   病人预交记录.Id%Type;

  v_结算方式 病人预交记录.结算方式%Type;
  n_结算金额 病人预交记录.冲预交%Type;
  v_结算号码 病人预交记录.结算号码%Type;
  v_结算摘要 病人预交记录.摘要%Type;
  v_交易人员 病人预交记录.交易人员%Type;

  n_返回值   人员缴款余额.余额%Type;
  n_冲预交   病人预交记录.冲预交%Type;
  v_退支票   病人预交记录.结算方式%Type;
  v_误差费   结算方式.名称%Type;
  n_Count    Number;
  n_Havenull Number;
  l_预交id   t_NumList := t_NumList();
  n_会话号   病人预交记录.会话号%Type; --格式：SID+'_'+SERIAL# 

  Cursor c_Feedata Is
    Select Max(m.病人id) As 病人id, Max(m.登记时间) As 登记时间, Max(m.操作员编号) As 操作员编号, Max(m.操作员姓名) As 操作员姓名, Sum(结帐金额) As 结算金额,
           Max(m.缴款组id) As 缴款组id
    From 门诊费用记录 M
    Where m.结帐id = 结帐id_In;
  r_Feedata c_Feedata%RowType;

  Cursor c_Balancedata Is
    Select 记录性质, NO, 记录状态, 病人id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号
    From 病人预交记录
    Where 结帐id = 结帐id_In And 结算方式 Is Null;
  r_Balancedata c_Balancedata%RowType;

Begin
  If Nvl(不控制会话_In, 0) = 0 Then
    Begin
      Select Sid || '_' || Serial# Into n_会话号 From V$session Where Audsid = Userenv('sessionid');
    Exception
      When Others Then
        n_会话号 := Null;
    End;
  End If;
  v_交易人员 := zl_UserName;

  Begin
    Select 名称 Into v_误差费 From 结算方式 Where Nvl(性质, 0) = 9;
  Exception
    When Others Then
      v_误差费 := '误差费';
  End;

  --0.正式结算 
  Select Count(1), Max(Decode(结算方式, Null, 1, 0))
  Into n_Count, n_Havenull
  From 病人预交记录
  Where 结帐id = 结帐id_In;

  --1.增加结算方式为空的结算数据 
  n_结算金额 := 0;
  n_Count    := 0;
  Open c_Feedata;
  Begin
    Fetch c_Feedata
      Into r_Feedata;
    --修正或新增结算方式为null的记录 
    Select Nvl(Sum(冲预交), 0) Into n_结算金额 From 病人预交记录 Where 结帐id = 结帐id_In;
    If Nvl(n_Havenull, 0) = 0 Or Round(Nvl(r_Feedata.结算金额, 0), 6) <> Round(Nvl(n_结算金额, 0), 6) Then
      --先删除存在的结算方式为null的记录 
      Delete From 病人预交记录 Where 结帐id = 结帐id_In And 结算方式 Is Null;
      Select Nvl(Sum(冲预交), 0) Into n_结算金额 From 病人预交记录 Where 结帐id = 结帐id_In;
    
      n_结算金额 := Round(Nvl(r_Feedata.结算金额, 0) - n_结算金额, 6);
    
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志, 结算性质, 会话号)
      Values
        (病人预交记录_Id.Nextval, 3, Null, 1, Decode(病人id_In, 0, Null, 病人id_In), Null, r_Feedata.登记时间, r_Feedata.操作员编号,
         r_Feedata.操作员姓名, n_结算金额, 结帐id_In, r_Feedata.缴款组id, Sysdate, v_交易人员, -1 * 结帐id_In, 1, 3, n_会话号);
    End If;
  Exception
    When Others Then
      n_Count := 1;
  End;
  Close c_Feedata;
  If n_Count = 1 Then
    v_Err_Msg := '未找到指定的收费明细数据,结算操作失败！';
    Raise Err_Item;
  End If;

  If 操作类型_In = 0 And Nvl(退支票额_In, 0) <> 0 Then
    Begin
      Select b.名称
      Into v_退支票
      From 结算方式应用 A, 结算方式 B
      Where a.应用场合 = '收费' And b.名称 = a.结算方式 And Nvl(b.应付款, 0) = 1 And Rownum <= 1;
    Exception
      When Others Then
        v_退支票 := '无';
    End;
    If v_退支票 = '无' Then
      v_Err_Msg := '在结算场合中,不存在结算性质为应付款的结算方式,请在[结算方式]中设置！';
      Raise Err_Item;
    End If;
  End If;

  Open c_Balancedata;
  Fetch c_Balancedata
    Into r_Balancedata;

  If Nvl(误差金额_In, 0) <> 0 Then
    Update 病人预交记录
    Set 冲预交 = Nvl(冲预交, 0) + Nvl(误差金额_In, 0)
    Where 结帐id = 结帐id_In And 结算方式 = v_误差费;
    If Sql%NotFound Then
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志, 卡类别id,
         结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号)
      Values
        (病人预交记录_Id.Nextval, 3, Null, 1, r_Balancedata.病人id, Null, Null, v_误差费, r_Balancedata.收款时间, r_Balancedata.操作员编号,
         r_Balancedata.操作员姓名, 误差金额_In, 结帐id_In, r_Balancedata.缴款组id, Sysdate, v_交易人员, r_Balancedata.结算序号, 2, Null, Null,
         卡号_In, 交易流水号_In, 交易说明_In, Null, 3, n_会话号);
    End If;
  
    Update 病人预交记录 Set 冲预交 = 冲预交 - Nvl(误差金额_In, 0) Where 结帐id = 结帐id_In And 结算方式 Is Null;
  End If;

  --预交款处理 
  If Nvl(冲预交_In, 0) <> 0 Then
    If Nvl(病人id_In, 0) = 0 Then
      v_Err_Msg := '不能确定病人的病人ID,收费不能使用预交款结算,结算操作失败！';
      Raise Err_Item;
    End If;
  
    Zl_病人预交记录_冲预交(病人id_In, 结帐id_In, 冲预交_In, 1, r_Balancedata.操作员编号, r_Balancedata.操作员姓名, r_Balancedata.收款时间,
                  冲预交病人ids_In, 3, 1);
  End If;

  If 操作类型_In = 0 Then
    If Nvl(退支票额_In, 0) <> 0 Then
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志, 卡类别id,
         结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号)
      Values
        (病人预交记录_Id.Nextval, 3, Null, 1, r_Balancedata.病人id, Null, Null, v_退支票, r_Balancedata.收款时间, r_Balancedata.操作员编号,
         r_Balancedata.操作员姓名, 退支票额_In, 结帐id_In, r_Balancedata.缴款组id, Sysdate, v_交易人员, r_Balancedata.结算序号, 2, Null, Null,
         卡号_In, 交易流水号_In, 交易说明_In, Null, 3, n_会话号);
    
      Update 病人预交记录 Set 冲预交 = 冲预交 - 退支票额_In Where 结帐id = 结帐id_In And 结算方式 Is Null;
    End If;
  
    --各个收费结算 :格式为:"结算方式|结算金额|结算号码|结算摘要||.." 
    v_结算内容 := 结算方式_In || '||';
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      v_结算号码 := LTrim(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_结算摘要 := LTrim(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
    
      If Nvl(n_结算金额, 0) <> 0 Then
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志,
           卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号)
        Values
          (病人预交记录_Id.Nextval, 3, Null, 1, r_Balancedata.病人id, Null, v_结算摘要, v_结算方式, r_Balancedata.收款时间,
           r_Balancedata.操作员编号, r_Balancedata.操作员姓名, n_结算金额, 结帐id_In, r_Balancedata.缴款组id, Sysdate, v_交易人员,
           r_Balancedata.结算序号, 2, Null, Null, 卡号_In, 交易流水号_In, 交易说明_In, v_结算号码, 3, n_会话号);
      
        Update 病人预交记录 Set 冲预交 = 冲预交 - n_结算金额 Where 结帐id = 结帐id_In And 结算方式 Is Null;
      End If;
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  --1.三方卡结算交易 
  If 操作类型_In = 1 Then
    v_当前结算 := 结算方式_In;
    v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
    v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
    n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
    v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
    v_结算号码 := LTrim(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
    v_结算摘要 := LTrim(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
  
    If Nvl(n_结算金额, 0) <> 0 Then
      Select Count(1) Into n_Count From 病人预交记录 Where ID = 关联交易id_In And Rownum < 2;
      If n_Count = 0 And Nvl(关联交易id_In, 0) <> 0 Then
        n_预交id := 关联交易id_In;
      Else
        Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      End If;
    
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志, 卡类别id,
         结算卡序号, 卡号, 关联交易id, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号)
      Values
        (n_预交id, 3, Null, 1, r_Balancedata.病人id, Null, v_结算摘要, v_结算方式, r_Balancedata.收款时间, r_Balancedata.操作员编号,
         r_Balancedata.操作员姓名, n_结算金额, 结帐id_In, r_Balancedata.缴款组id, Sysdate, v_交易人员, r_Balancedata.结算序号, 2, 卡类别id_In,
         Null, 卡号_In, 关联交易id_In, 交易流水号_In, 交易说明_In, v_结算号码, 3, n_会话号);
    
      Update 病人预交记录 Set 冲预交 = 冲预交 - n_结算金额 Where 结帐id = 结帐id_In And 结算方式 Is Null;
    
      --调用其他结算信息更新
      Zl_Custom_Balance_Update(n_预交id);
    End If;
  End If;

  --2.医保结算(调用此过程,采取平均分摊的方式分摊结算情况):这种情况医保结处后,必须全退 
  If 操作类型_In = 2 Then
    --2.1检查是否已经存在医保结算数据,存在先删除 
    n_结算金额 := 0;
    For v_医保 In (Select ID, 结算方式, 冲预交
                 From 病人预交记录 A
                 Where a.结帐id = 结帐id_In And Exists
                  (Select 1 From 结算方式 Where a.结算方式 = 名称 And 性质 In (3, 4)) And Mod(记录性质, 10) <> 1) Loop
      n_结算金额 := n_结算金额 + Nvl(v_医保.冲预交, 0);
      l_预交id.Extend;
      l_预交id(l_预交id.Count) := v_医保.Id;
    End Loop;
  
    Update 病人预交记录 Set 冲预交 = Nvl(冲预交, 0) + Nvl(n_结算金额, 0) Where 结帐id = 结帐id_In And 结算方式 Is Null;
  
    Forall I In 1 .. l_预交id.Count
      Delete From 病人预交记录 Where ID = l_预交id(I);
  
    If 结算方式_In Is Not Null Then
      v_结算内容 := 结算方式_In || '||';
    End If;
  
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      n_结算金额 := To_Number(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
    
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志, 结算性质,
         会话号)
      Values
        (病人预交记录_Id.Nextval, 3, Null, 1, r_Balancedata.病人id, Null, '保险结算', v_结算方式, r_Balancedata.收款时间,
         r_Balancedata.操作员编号, r_Balancedata.操作员姓名, n_结算金额, 结帐id_In, r_Balancedata.缴款组id, Sysdate, v_交易人员,
         r_Balancedata.结算序号, 1, 3, n_会话号);
    
      --更新数据(结算方式为NULL的) 
      Update 病人预交记录
      Set 冲预交 = 冲预交 - n_结算金额
      Where 结帐id = 结帐id_In And 结算方式 Is Null
      Returning Nvl(冲预交, 0) Into n_返回值;
    
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  --3-消费卡批量结算 
  If 操作类型_In = 3 Then
    v_结算内容 := 结算方式_In || '||';
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      --卡类别ID|卡号|消费卡ID|消费金额 
      n_卡类别id := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      v_卡号     := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_消费卡id := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_结算金额 := To_Number(v_当前结算);
      Begin
        Select 名称, 结算方式 Into v_名称, v_结算方式 From 消费卡类别目录 Where 编号 = 卡类别id_In;
      Exception
        When Others Then
          v_名称 := Null;
      End;
      If v_名称 Is Null Then
        v_Err_Msg := '未找到对应的结算卡接口,本次刷卡消费失败!';
        Raise Err_Item;
      End If;
      If v_结算方式 Is Null Then
        v_Err_Msg := v_名称 || '未设置对应的结算方式,本次刷卡消费失败!';
        Raise Err_Item;
      End If;
    
      If Nvl(n_结算金额, 0) <> 0 Then
        Update 病人预交记录
        Set 冲预交 = Nvl(冲预交, 0) + n_结算金额
        Where 结帐id = r_Balancedata. 结帐id And 结算方式 = v_结算方式 And 结算卡序号 = n_卡类别id
        Returning ID Into n_预交id;
        If Sql%NotFound Then
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 结算卡序号,
             校对标志, 结算性质, 会话号)
          Values
            (n_预交id, 3, Null, 1, r_Balancedata. 病人id, Null, Null, v_结算方式, r_Balancedata. 收款时间, r_Balancedata. 操作员编号,
             r_Balancedata. 操作员姓名, n_结算金额, r_Balancedata. 结帐id, r_Balancedata. 缴款组id, Sysdate, v_交易人员,
             r_Balancedata. 结算序号, n_卡类别id, 2, 3, n_会话号);
        End If;
      
        Zl_病人卡结算记录_支付(n_卡类别id, v_卡号, n_消费卡id, n_结算金额, n_预交id, r_Balancedata. 操作员编号, r_Balancedata. 操作员姓名,
                      r_Balancedata. 收款时间);
      
        --更新数据(结算方式为NULL的) 
        Update 病人预交记录
        Set 冲预交 = 冲预交 - n_结算金额
        Where 结帐id = r_Balancedata. 结帐id And 结算方式 Is Null And Nvl(校对标志, 0) = 1
        Returning Nvl(冲预交, 0) Into n_返回值;
      End If;
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  --4.三方卡结算，多种结算方式 
  If 操作类型_In = 4 Then
    If Nvl(删除原结算_In, 0) = 1 Then
      --1.1检查是否已经存在三方卡结算数据,存在先删除 
      n_结算金额 := 0;
      For c_结算 In (Select ID, 结算方式, 冲预交
                   From 病人预交记录 A
                   Where 结帐id = 结帐id_In And 卡类别id = 卡类别id_In And 关联交易id = 关联交易id_In) Loop
        n_结算金额 := n_结算金额 + Nvl(c_结算.冲预交, 0);
        l_预交id.Extend;
        l_预交id(l_预交id.Count) := c_结算.Id;
      End Loop;
    
      Update 病人预交记录
      Set 冲预交 = Nvl(冲预交, 0) + Nvl(n_结算金额, 0)
      Where 结帐id = 结帐id_In And 结算方式 Is Null;
    
      Forall I In 1 .. l_预交id.Count
        Delete From 病人预交记录 Where ID = l_预交id(I);
    End If;
  
    n_预交id := 0;
    --格式：结算方式|结算金额|结算号码|结算摘要|单据号|是否普通结算|卡号 
    For c_结算 In (Select Max(Decode(序号, 1, 值, Null)) As 结算方式, zl_To_Number(Max(Decode(序号, 2, 值, ''))) As 结算金额,
                        Trim(Max(Decode(序号, 3, 值, ''))) As 结算号码, Trim(Max(Decode(序号, 4, 值, ''))) As 结算摘要,
                        Trim(Max(Decode(序号, 5, 值, ''))) As 单据号, zl_To_Number(Max(Decode(序号, 6, 值, ''))) As 是否普通结算,
                        Trim(Max(Decode(序号, 7, 值, ''))) As 卡号
                 From (Select Rownum As 序号, Column_Value As 值 From Table(f_Str2List(结算方式_In, '|')))
                 Having Nvl(zl_To_Number(Max(Decode(序号, 2, 值, ''))), 0) <> 0) Loop
    
      Update 病人预交记录
      Set 冲预交 = 冲预交 + c_结算.结算金额
      Where 结帐id = 结帐id_In And 结算方式 = c_结算.结算方式 And 关联交易id = 关联交易id_In
      Returning ID Into n_预交id;
      If Sql%NotFound Then
        Select Count(1) Into n_Count From 病人预交记录 Where ID = 关联交易id_In And Rownum < 2;
        If n_Count = 0 Then
          n_预交id := 关联交易id_In;
        Else
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
        End If;
      
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志,
           卡类别id, 卡号, 关联交易id, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号)
        Values
          (n_预交id, 3, Null, 1, r_Balancedata.病人id, Null, c_结算.结算摘要, c_结算.结算方式, r_Balancedata.收款时间, r_Balancedata.操作员编号,
           r_Balancedata.操作员姓名, c_结算.结算金额, 结帐id_In, r_Balancedata.缴款组id, Sysdate, v_交易人员, r_Balancedata.结算序号, 校对标志_In,
           Decode(c_结算.是否普通结算, 1, Null, 卡类别id_In), Decode(c_结算.是否普通结算, 1, Null, Nvl(c_结算.卡号, 卡号_In)), 关联交易id_In,
           交易流水号_In, 交易说明_In, c_结算.结算号码, 3, n_会话号);
      End If;
    
      Update 病人预交记录 Set 冲预交 = 冲预交 - c_结算.结算金额 Where 结帐id = 结帐id_In And 结算方式 Is Null;
    
      If c_结算.单据号 Is Not Null Then
        Insert Into 医保结算明细
          (结帐id, NO, 结算方式, 金额, 卡类别id, 关联交易id, 交易流水号, 交易说明)
        Values
          (结帐id_In, c_结算.单据号, c_结算.结算方式, c_结算.结算金额, Decode(c_结算.是否普通结算, 1, Null, 卡类别id_In), 关联交易id_In, 交易流水号_In,
           交易说明_In);
      End If;
    
      --调用其他结算信息更新
      Zl_Custom_Balance_Update(n_预交id);
    End Loop;
  End If;

  If Nvl(完成结算_In, 0) = 0 Then
    Return;
  End If;

  ----------------------------------------------------------------------------------------- 
  --完成收费,需要处理人员缴款余额,预交记录(结算方式=NULL) 

  --1.删除结算方式为NULL的预交记录 
  Delete 病人预交记录 Where 结帐id = 结帐id_In And 结算方式 Is Null And Nvl(冲预交, 0) = 0;
  If Sql%NotFound Then
    Select Count(*) Into n_Count From 病人预交记录 A Where 结帐id = 结帐id_In And 结算方式 Is Null;
    If n_Count <> 0 Then
      v_Err_Msg := '还存在未缴款的数据,不能完成结算!';
    Else
      v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！!';
    End If;
    Raise Err_Item;
  End If;

  --检查门诊费用记录与病人预交记录的金额是否相等 
  n_结算金额 := 0;
  n_冲预交   := 0;
  Select Nvl(Sum(实收金额), 0) Into n_结算金额 From 门诊费用记录 Where 结帐id = 结帐id_In;
  Select Nvl(Sum(冲预交), 0) Into n_冲预交 From 病人预交记录 Where 结帐id = 结帐id_In;
  If n_结算金额 <> n_冲预交 Then
    v_Err_Msg := '结算信息有误，实收金额(' || n_结算金额 || ')与结算金额(' || n_冲预交 || ')不一致，不能完成结算！';
    Raise Err_Item;
  End If;

  --结算金额为零时，增加一条金额为0的病人预交记录 
  Select Count(1) Into n_Count From 病人预交记录 A Where 结帐id = 结帐id_In;
  If n_Count = 0 Then
    v_结算方式 := 缺省结算方式_In;
    If v_结算方式 Is Null Then
      Begin
        Select 结算方式 Into v_结算方式 From 结算方式应用 Where 应用场合 = '收费' And Nvl(缺省标志, 0) = 1;
      Exception
        When Others Then
          v_结算方式 := Null;
      End;
      If v_结算方式 Is Null Then
        Begin
          Select 名称 Into v_结算方式 From 结算方式 Where Nvl(性质, 0) = 1;
        Exception
          When Others Then
            v_结算方式 := '现金';
        End;
      End If;
    End If;
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志, 卡类别id,
       结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号)
    Values
      (病人预交记录_Id.Nextval, 3, Null, 1, r_Balancedata.病人id, Null, Null, v_结算方式, r_Balancedata.收款时间, r_Balancedata.操作员编号,
       r_Balancedata.操作员姓名, 0, 结帐id_In, r_Balancedata.缴款组id, Sysdate, v_交易人员, r_Balancedata.结算序号, 2, Null, Null, Null,
       Null, 交易说明_In, Null, 3, n_会话号);
  End If;

  --2.处理缴款数据和找补数据及校对标志更新为0，会话号更新为NULL 
  Update 病人预交记录
  Set 缴款 = 缴款_In, 找补 = 找补_In, 校对标志 = 0, 会话号 = Null, 是否电子票据 = 是否电子票据_In
  Where 结帐id = 结帐id_In And 记录性质 <> 1;

  --3.更新费用状态 
  Update 门诊费用记录 Set 费用状态 = 0 Where 结帐id = 结帐id_In;

  --4.更新人员缴款数据 
  If Nvl(更新交款余额_In, 1) = 1 Then
    For c_缴款 In (Select 结算方式, 操作员姓名, Sum(Nvl(a.冲预交, 0)) As 冲预交
                 From 病人预交记录 A
                 Where a.结帐id = 结帐id_In And Mod(a.记录性质, 10) <> 1
                 Group By 结算方式, 操作员姓名) Loop
    
      Update 人员缴款余额
      Set 余额 = Nvl(余额, 0) + Nvl(c_缴款.冲预交, 0)
      Where 收款员 = c_缴款.操作员姓名 And 性质 = 1 And 结算方式 = c_缴款.结算方式;
      If Sql%RowCount = 0 Then
        Insert Into 人员缴款余额
          (收款员, 结算方式, 性质, 余额)
        Values
          (c_缴款.操作员姓名, c_缴款.结算方式, 1, Nvl(c_缴款.冲预交, 0));
      End If;
    End Loop;
  End If;

  --5.相关业务数据处理 
  Zl_门诊收费记录_完成收费(结帐id_In);

  --消息集成处理 
  --结算类型:1-收费结算，2-补充结算 
  --结帐ID:结算id 
  b_Message.Zlhis_Charge_002(1, 结帐id_In);

  --收费后产生导引 
  Begin
    Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
      Using 4, 结帐id_In;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊收费结算_Modify;
/
Create Or Replace Procedure Zl_门诊退费结算_Modify
(
  操作类型_In      Number,
  病人id_In        门诊费用记录.病人id%Type,
  冲销id_In        病人预交记录.结帐id%Type,
  结算方式_In      Varchar2,
  冲预交_In        病人预交记录.冲预交%Type := Null,
  卡类别id_In      病人预交记录.卡类别id%Type := Null,
  卡号_In          病人预交记录.卡号%Type := Null,
  交易流水号_In    病人预交记录.交易流水号%Type := Null,
  交易说明_In      病人预交记录.交易说明%Type := Null,
  缴款_In          病人预交记录.缴款%Type := Null,
  找补_In          病人预交记录.找补%Type := Null,
  误差金额_In      门诊费用记录.实收金额%Type := Null,
  完成退费_In      Number := 0,
  原结帐id_In      病人预交记录.结帐id%Type := Null,
  剩余转预交_In    Number := 0,
  缺省结算方式_In  结算方式.名称%Type := Null,
  冲预交病人ids_In Varchar2 := Null,
  关联交易id_In    病人预交记录.关联交易id%Type := Null,
  删除原结算_In    Number := 0,
  校对标志_In      病人预交记录.校对标志%Type := 0
) As
  ------------------------------------------------------------------------------------------------------------------------------ 
  --功能:收费结算时,修改结算的相关信息 
  --操作类型_In: 
  --   0-原样退 
  --      原样结算一起全退,所有校对标志都为1,医保调用成功后,调整为2,完成后变成0 
  --   1-普通退费方式: 
  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空. 
  --   2.三方卡退费结算: 
  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要" 
  --     ②卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入 
  --   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新) 
  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
  --   4-消费卡结算:
  --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||." 
  --   5.三方卡退费结算，多种结算方式: 
  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要|单据号|是否普通结算" 
  --     ②卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入 

  -- 冲预交_In:如果涉及预交款,则传入本次的退预交 传入零<0时 表示退预交款或充值;>0 时:表示冲预交款 
  -- 剩余转预交_In: 1表示将剩余退款额转换为充值金额;0表示退预交 
  -- 误差金额_In:存在误差费时,传入 
  -- 完成退费_In:0-未完成退费;1-异常完成退费;2-完成退费 
  -- 原结帐ID_IN:原样退时,传入(如果原样退未传入时,则以最后一次结帐为准) 
  -- 冲预交病人ids_In:多个用逗分分离,冲预交时有效(冲预交类别与业务操作保持一致),主要是使用家属的预交款 
  -- 关联交易id_In 操作类型_In 为3,5时必须传入 
  -- 删除原结算_in 操作类型_In为5时有效，多个结算方式时调用多次该过程 
  -- 校对标志_In  操作类型_In为5时有效
  ------------------------------------------------------------------------------------------------------------------------------ 
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_结算内容 Varchar2(500);
  v_当前结算 Varchar2(50);
  v_卡号     病人医疗卡信息.卡号%Type;
  n_消费卡id 消费卡信息.Id%Type;
  v_名称     消费卡类别目录.名称%Type;
  n_卡类别id 病人预交记录.结算卡序号%Type;
  n_原预交id 病人预交记录.Id%Type;
  n_预交id   病人预交记录.Id%Type;
  v_结算方式 病人预交记录.结算方式%Type;
  n_结算金额 病人预交记录.冲预交%Type;
  n_返回值   人员缴款余额.余额%Type;
  n_预交金额 病人预交记录.冲预交%Type;
  n_冲预交   病人预交记录.冲预交%Type;
  v_结算号码 病人预交记录.结算号码%Type;
  v_结算摘要 病人预交记录.摘要%Type;
  v_误差费   结算方式.名称%Type;
  n_记录状态 病人预交记录.记录状态%Type;
  n_充值id   病人预交记录.Id%Type;

  v_退费结算 结算方式.名称%Type;
  v_No       病人预交记录.No%Type;
  n_Dec      Number; --金额小数位数 

  n_Count    Number;
  n_Havenull Number;
  l_预交id   t_NumList := t_NumList();
  n_原结帐id 病人预交记录.结帐id%Type;
  n_重结id   病人预交记录.结帐id%Type;
  n_结帐id   病人预交记录.结帐id%Type;
  n_结算序号 病人预交记录.结帐id%Type;
  v_Msg      Varchar2(500);
  n_会话号   病人预交记录.会话号%Type; --格式：SID+'_'+SERIAL# 
  v_交易人员 病人预交记录.交易人员%Type;
  n_异步结算 Number;

  n_门诊转住院退费 Number;
  n_是否电子票据   病人预交记录.是否电子票据%Type;

  Cursor c_Feedata Is
    Select Max(NO) As NO, Max(m.病人id) As 病人id, Max(m.登记时间) As 登记时间, Max(m.操作员编号) As 操作员编号, Max(m.操作员姓名) As 操作员姓名,
           Sum(结帐金额) As 结算金额, Max(m.缴款组id) As 缴款组id
    From 门诊费用记录 M
    Where m.结帐id = 冲销id_In;
  r_Feedata c_Feedata%RowType;

  Cursor c_Balancedata Is
    Select 记录性质, NO, 记录状态, 病人id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号
    From 病人预交记录
    Where 结帐id = 冲销id_In And 结算方式 Is Null;
  r_Balancedata c_Balancedata%RowType;

Begin
  Begin
    Select Sid || '_' || Serial# Into n_会话号 From V$session Where Audsid = Userenv('sessionid');
  Exception
    When Others Then
      n_会话号 := Null;
  End;
  v_交易人员 := zl_UserName;

  Begin
    Select 名称 Into v_退费结算 From 结算方式 Where 性质 = 1;
  Exception
    When Others Then
      v_退费结算 := '现金';
  End;
  Select Count(1) Into n_异步结算 From 费用结算对照 A Where a.门诊标志 = 1 And a.结帐id = 冲销id_In And Rownum < 2;

  --金额小数位数 
  Select zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  --0.正式结算 
  Select Count(1), Max(Decode(结算方式, Null, 1, 0)), Max(结算序号)
  Into n_Count, n_Havenull, n_结算序号
  From 病人预交记录
  Where 结帐id = 冲销id_In;

  If Nvl(n_Count, 0) = 0 Or Nvl(误差金额_In, 0) <> 0 Then
    --增加结算方式为NULL的记录 
    Begin
      Select 名称 Into v_误差费 From 结算方式 Where Nvl(性质, 0) = 9;
    Exception
      When Others Then
        v_误差费 := '误差费';
    End;
  End If;

  --1.增加结算方式为空的结算数据 
  n_Count := 0;
  Open c_Feedata;
  Begin
    Fetch c_Feedata
      Into r_Feedata;
    If Nvl(n_Havenull, 0) = 0 Then
      n_结算金额 := Round(Nvl(r_Feedata.结算金额, 0), n_Dec);
    
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志, 结算性质, 会话号)
      Values
        (病人预交记录_Id.Nextval, 3, Null, 2, Decode(病人id_In, 0, Null, 病人id_In), Null, r_Feedata.登记时间, r_Feedata.操作员编号,
         r_Feedata.操作员姓名, n_结算金额, 冲销id_In, r_Feedata.缴款组id, Sysdate, v_交易人员, -1 * 冲销id_In, 1, 3, n_会话号);
    
      --误差费(先汇总后生成误差费 
      If n_结算金额 <> Nvl(r_Feedata.结算金额, 0) Then
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志, 结算性质, 会话号)
        Values
          (病人预交记录_Id.Nextval, 3, Null, 1, Decode(病人id_In, 0, Null, 病人id_In), v_误差费, r_Feedata.登记时间, r_Feedata.操作员编号,
           r_Feedata.操作员姓名, Nvl(r_Feedata.结算金额, 0) - n_结算金额, 冲销id_In, r_Feedata.缴款组id, Sysdate, v_交易人员, -1 * 冲销id_In, 1,
           3, n_会话号);
      End If;
      n_结算序号 := -1 * 冲销id_In;
    End If;
  Exception
    When Others Then
      n_Count := 1;
  End;
  Close c_Feedata;
  If n_Count = 1 Then
    v_Err_Msg := '未找到指定的收费明细数据,结算操作失败！';
    Raise Err_Item;
  End If;

  Open c_Balancedata;
  Fetch c_Balancedata
    Into r_Balancedata;

  n_原结帐id := 原结帐id_In;
  If Nvl(n_原结帐id, 0) = 0 Then
    Select Max(b.结帐id)
    Into n_原结帐id
    From 门诊费用记录 A, 门诊费用记录 B
    Where a.结帐id = 冲销id_In And a.No = b.No And b.记录性质 = 1 And b.记录状态 In (1, 3);
  End If;

  If Nvl(n_原结帐id, 0) = 0 Then
    v_Err_Msg := '未找到原结帐数据,不能原样退！';
    Raise Err_Item;
  End If;

  If 操作类型_In = 0 Then
    --0.原样退 
    --1.只处理消费卡部分 
    Select Count(1)
    Into n_Count
    From 病人预交记录 A, 病人卡结算记录 B
    Where a.Id = b.结算id And a.记录性质 = 3 And a.结帐id = n_原结帐id And Rownum < 2;
    If n_Count <> 0 Then
      Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
    
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间,
         交易人员, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质, 会话号)
        Select n_预交id, 记录性质, NO, 2, 病人id, 主页id, 摘要, 结算方式, 结算号码, r_Balancedata. 收款时间, 缴款单位, 单位开户行, 单位帐号,
               r_Balancedata.操作员编号, r_Balancedata.操作员姓名, -1 * 冲预交, 冲销id_In, r_Balancedata.缴款组id, Sysdate, v_交易人员, 预交类别,
               卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 2, r_Balancedata.结算序号, Mod(记录性质, 10), n_会话号
        
        From 病人预交记录 A
        Where a.记录性质 = 3 And a.结帐id = n_原结帐id And Exists (Select 1 From 病人卡结算记录 Where 结算id = a.Id);
    
      --收费时可能使用了多张消费卡 
      For c_记录 In (Select a.Id, c.接口编号, c.消费卡id, c.卡号, -1 * Sum(c.应收金额) As 结算金额
                   From 病人预交记录 A, 病人卡结算记录 C
                   Where a.Id = c.结算id And a.记录性质 = 3 And a.记录状态 In (1, 3) And a.结帐id = n_原结帐id
                   Group By a.Id, c.接口编号, c.消费卡id, c.卡号) Loop
      
        Zl_病人卡结算记录_退款(c_记录.接口编号, c_记录.卡号, c_记录.消费卡id, c_记录.结算金额, c_记录.Id, n_预交id, r_Balancedata. 操作员编号,
                      r_Balancedata. 操作员姓名, r_Balancedata. 收款时间);
      End Loop;
    End If;
  
    --2.医保 
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 结算号码, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 关联交易id, 交易时间, 交易人员, 预交类别,
       卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质, 会话号)
      Select 病人预交记录_Id.Nextval, 记录性质, NO, 2, 病人id, 主页id, 摘要, 结算方式, 结算号码, r_Balancedata.收款时间, r_Balancedata.操作员编号,
             r_Balancedata.操作员姓名, -1 * 冲预交, r_Balancedata.结帐id, r_Balancedata.缴款组id, 关联交易id, Sysdate, v_交易人员, 预交类别,
             卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 1 As 校对标志, r_Balancedata.结算序号, Mod(记录性质, 10), n_会话号
      From 病人预交记录 A, (Select 名称 From 结算方式 Where 性质 In (3, 4)) J
      Where a.记录状态 In (1, 3) And a.结算方式 = j.名称 And a.结算方式 Is Not Null And a.结帐id = n_原结帐id And a.卡类别id Is Null;
  
    --更新结算方式为NULL 的记录 
    Select Sum(冲预交) Into n_返回值 From 病人预交记录 Where 结帐id = 冲销id_In And 结算方式 Is Not Null;
    Select Sum(结帐金额) Into n_结算金额 From 门诊费用记录 Where 结帐id = 冲销id_In;
    Update 病人预交记录
    Set 冲预交 = Nvl(n_结算金额, 0) - Nvl(n_返回值, 0)
    Where 结帐id = 冲销id_In And 结算方式 Is Null;
  End If;

  n_重结id := 0;
  If 操作类型_In <> 0 Then
    --不是全退时,检查是否产生了重新收费数据的 
    Begin
      Select 结帐id Into n_重结id From 病人预交记录 Where 结算序号 = n_结算序号 And 结帐id <> 冲销id_In And Rownum < 2;
    Exception
      When Others Then
        n_重结id := 0;
    End;
  End If;

  --需要处理误差金额 
  If Nvl(误差金额_In, 0) <> 0 Then
    --误差费放在重收的结算记录中 
    n_结帐id   := 冲销id_In;
    n_记录状态 := 2;
    If Nvl(n_重结id, 0) <> 0 Then
      n_结帐id   := n_重结id;
      n_记录状态 := 1;
    End If;
  
    Update 病人预交记录
    Set 冲预交 = Nvl(冲预交, 0) + Nvl(误差金额_In, 0)
    Where 结帐id = n_结帐id And 结算方式 = v_误差费;
    If Sql%NotFound Then
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志, 卡类别id,
         结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号)
      Values
        (病人预交记录_Id.Nextval, 3, Null, n_记录状态, r_Balancedata.病人id, Null, Null, v_误差费, r_Balancedata.收款时间,
         r_Balancedata.操作员编号, r_Balancedata.操作员姓名, 误差金额_In, n_结帐id, r_Balancedata.缴款组id, Sysdate, v_交易人员,
         r_Balancedata.结算序号, 2, Null, Null, 卡号_In, 交易流水号_In, 交易说明_In, Null, 3, n_会话号);
    End If;
  
    Update 病人预交记录 Set 冲预交 = 冲预交 - Nvl(误差金额_In, 0) Where 结帐id = n_结帐id And 结算方式 Is Null;
  End If;

  --预交款处理:如果是冲预交,需要先处理冲预交款 
  If Nvl(冲预交_In, 0) <> 0 Then
    If Nvl(r_Balancedata.病人id, 0) = 0 Then
      v_Err_Msg := '不能确定病人信息,不能使用预交款结算！';
      Raise Err_Item;
    End If;
  
    n_预交金额 := 冲预交_In;
    If n_预交金额 < 0 And Nvl(剩余转预交_In, 0) = 1 Then
      --     ②冲预交_In:如果涉及预交款,则传入本次的退预交 传入零<0时 表示退预交款或充值;>0 时:表示冲预交款 
      --     ③剩余转预交_In: 1表示将剩余退款额转换为充值金额;0表示退预交 
    
      --1.先生成冲值预交 
      v_No := Nextno(11);
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 金额, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志, 卡类别id,
         结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 预交类别, 结算性质, 会话号)
      Values
        (病人预交记录_Id.Nextval, 1, v_No, 1, r_Balancedata.病人id, Null, '退费生成预交', v_退费结算, r_Balancedata.收款时间,
         r_Balancedata.操作员编号, r_Balancedata.操作员姓名, -1 * n_预交金额, 冲销id_In, r_Balancedata.缴款组id, Sysdate, v_交易人员,
         r_Balancedata.结算序号, 0, Null, Null, 卡号_In, 交易流水号_In, 交易说明_In, Null, 1, Null, n_会话号);
    
      --预交单据余额 
      Insert Into 预交单据余额
        (预交id, 病人id, 预交类别, 预交余额)
      Values
        (病人预交记录_Id.Currval, r_Balancedata.病人id, 1, -1 * n_预交金额);
    
      --更新病人余额 
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) + (-1 * n_预交金额)
      Where 病人id = 病人id_In And 性质 = 1 And 类型 = 1
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (病人id_In, 1, -1 * n_预交金额, 1);
        n_返回值 := -1 * n_预交金额;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 病人余额
        Where 病人id = r_Balancedata.病人id And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
      End If;
    
      --2.生成退费记录 
      If Nvl(n_重结id, 0) <> 0 Then
        Select Sum(冲预交)
        Into n_返回值
        From 病人预交记录
        Where 结帐id = 冲销id_In And 结算方式 Is Null And Nvl(冲预交, 0) <> 0;
        If Nvl(n_返回值, 0) <> 0 Then
          Insert Into 病人预交记录
            (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志,
             卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号)
          Values
            (病人预交记录_Id.Nextval, 3, Null, 2, r_Balancedata.病人id, Null, v_结算摘要, v_退费结算, r_Balancedata.收款时间,
             r_Balancedata.操作员编号, r_Balancedata.操作员姓名, n_返回值, 冲销id_In, r_Balancedata.缴款组id, Sysdate, v_交易人员,
             r_Balancedata.结算序号, 2, Null, Null, 卡号_In, 交易流水号_In, 交易说明_In, Null, 3, n_会话号);
        
          Update 病人预交记录 Set 冲预交 = 冲预交 - Nvl(n_返回值, 0) Where 结帐id = 冲销id_In And 结算方式 Is Null;
        End If;
        n_结算金额 := n_结算金额 - Nvl(n_返回值, 0);
      
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志,
           卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号)
        Values
          (病人预交记录_Id.Nextval, 3, Null, 1, r_Balancedata.病人id, Null, v_结算摘要, v_退费结算, r_Balancedata.收款时间,
           r_Balancedata.操作员编号, r_Balancedata.操作员姓名, n_结算金额, n_重结id, r_Balancedata.缴款组id, Sysdate, v_交易人员,
           r_Balancedata.结算序号, 2, Null, Null, 卡号_In, 交易流水号_In, 交易说明_In, Null, 3, n_会话号);
      
        Update 病人预交记录 Set 冲预交 = 冲预交 + n_预交金额 Where 结帐id = 冲销id_In And 结算方式 Is Null;
      Else
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志,
           卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号)
        Values
          (病人预交记录_Id.Nextval, 3, Null, 2, r_Balancedata.病人id, Null, v_结算摘要, v_结算方式, r_Balancedata.收款时间,
           r_Balancedata.操作员编号, r_Balancedata.操作员姓名, -1 * n_预交金额, 冲销id_In, r_Balancedata.缴款组id, Sysdate, v_交易人员,
           r_Balancedata.结算序号, 2, Null, Null, 卡号_In, 交易流水号_In, 交易说明_In, Null, 3, n_会话号);
      
        Update 病人预交记录 Set 冲预交 = 冲预交 + n_预交金额 Where 结帐id = 冲销id_In And 结算方式 Is Null;
      End If;
    End If;
  
    If Nvl(n_预交金额, 0) < 0 And Nvl(剩余转预交_In, 0) = 0 Then
      If Nvl(n_重结id, 0) <> 0 Then
        --1.退预交款 
        For v_退预交 In (Select Max(a.Id) As ID, Max(a.No) As NO, a.病人id, Max(a.收款时间) As 收款时间, Sum(Nvl(a.冲预交, 0)) As 金额
                      From 病人预交记录 A,
                           (Select Distinct a.结帐id
                             From 门诊费用记录 A, 门诊费用记录 B
                             Where a.No = b.No And Mod(a.记录性质, 10) = 1 And b.结帐id = n_原结帐id) B
                      Where a.结帐id = b.结帐id And Mod(a.记录性质, 10) = 1 And Nvl(a.预交类别, 0) = 1
                      Group By NO, 病人id
                      Having Sum(Nvl(a.冲预交, 0)) > 0
                      Order By 收款时间 Desc) Loop
          Insert Into 病人预交记录
            (ID, 记录性质, NO, 实际票号, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号,
             校对标志, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 预交类别, 结算性质, 会话号, 关联交易id)
            Select 病人预交记录_Id.Nextval, 11, NO, 实际票号, 记录状态, 病人id, Null, 摘要, 结算方式, r_Balancedata.收款时间, r_Balancedata.操作员编号,
                   r_Balancedata.操作员姓名, -1 * v_退预交.金额, 冲销id_In, r_Balancedata.缴款组id, Sysdate, v_交易人员, r_Balancedata.结算序号,
                   2, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 预交类别, 3, n_会话号, 关联交易id
            From 病人预交记录
            Where ID = v_退预交.Id;
        
          --更新预交单据余额 
          Select Max(ID) Into n_充值id From 病人预交记录 Where NO = v_退预交.No And 记录性质 = 1 And 记录状态 <> 2;
          If Nvl(n_充值id, 0) <> 0 Then
            Update 预交单据余额
            Set 预交余额 = Nvl(预交余额, 0) + Nvl(v_退预交.金额, 0)
            Where 病人id = v_退预交.病人id And 预交id = n_充值id
            Returning 预交余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 预交单据余额
                (预交id, 病人id, 预交类别, 预交余额)
              Values
                (n_充值id, v_退预交.病人id, 1, Nvl(v_退预交.金额, 0));
              n_返回值 := Nvl(v_退预交.金额, 0);
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 预交单据余额 Where 预交id = n_充值id And Nvl(预交余额, 0) = 0;
            End If;
          End If;
        
          --更新病人余额 
          Update 病人余额
          Set 预交余额 = Nvl(预交余额, 0) + v_退预交.金额
          Where 病人id = v_退预交.病人id And 性质 = 1 And 类型 = 1
          Returning 预交余额 Into n_返回值;
          If Sql%RowCount = 0 Then
            Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (v_退预交.病人id, 1, v_退预交.金额, 1);
            n_返回值 := v_退预交.金额;
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 病人余额
            Where 病人id = v_退预交.病人id And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
          End If;
        
          Update 病人预交记录 Set 冲预交 = 冲预交 - (-1 * v_退预交.金额) Where 结帐id = 冲销id_In And 结算方式 Is Null;
        
          n_预交金额 := n_预交金额 - (-1 * v_退预交.金额);
        End Loop;
      
        --2.冲预交款 
        If n_预交金额 <> 0 Then
          For v_退预交 In (Select Max(a.Id) As ID, a.No, a.病人id, a.预交类别, Max(a.收款时间) As 收款时间, Sum(Nvl(a.冲预交, 0)) As 金额
                        From 病人预交记录 A,
                             (Select Distinct a.结帐id
                               From 门诊费用记录 A, 门诊费用记录 B
                               Where a.No = b.No And Mod(a.记录性质, 10) = 1 And b.结帐id = n_原结帐id) B
                        Where a.结帐id = b.结帐id And Mod(a.记录性质, 10) = 1 And Nvl(a.预交类别, 0) = 1 And a.结帐id <> 冲销id_In
                        Group By a.No, a.病人id, a.预交类别
                        Having Sum(Nvl(a.冲预交, 0)) > 0
                        Order By 收款时间 Desc) Loop
          
            If v_退预交.金额 - n_预交金额 < 0 Then
              n_结算金额 := v_退预交.金额;
              n_预交金额 := n_预交金额 - v_退预交.金额;
            Else
              n_结算金额 := n_预交金额;
              n_预交金额 := 0;
            End If;
          
            Insert Into 病人预交记录
              (ID, 记录性质, NO, 实际票号, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号,
               校对标志, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 预交类别, 结算性质, 会话号, 关联交易id)
              Select 病人预交记录_Id.Nextval, 11, NO, 实际票号, 记录状态, 病人id, Null, 摘要, 结算方式, r_Balancedata.收款时间,
                     r_Balancedata.操作员编号, r_Balancedata.操作员姓名, n_结算金额, n_重结id, r_Balancedata.缴款组id, Sysdate, v_交易人员,
                     r_Balancedata.结算序号, 2, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 预交类别, 3, n_会话号, 关联交易id
              From 病人预交记录
              Where ID = v_退预交.Id;
          
            --更新预交单据余额 
            Select Max(ID) Into n_充值id From 病人预交记录 Where NO = v_退预交.No And 记录性质 = 1 And 记录状态 <> 2;
            If Nvl(n_充值id, 0) <> 0 Then
              Update 预交单据余额
              Set 预交余额 = Nvl(预交余额, 0) + (-1 * n_结算金额)
              Where 病人id = v_退预交.病人id And 预交id = n_充值id
              Returning 预交余额 Into n_返回值;
              If Sql%RowCount = 0 Then
                Insert Into 预交单据余额
                  (预交id, 病人id, 预交类别, 预交余额)
                Values
                  (n_充值id, v_退预交.病人id, Nvl(v_退预交.预交类别, 2), -1 * n_结算金额);
                n_返回值 := -1 * n_结算金额;
              End If;
              If Nvl(n_返回值, 0) = 0 Then
                Delete From 预交单据余额 Where 预交id = n_充值id And Nvl(预交余额, 0) = 0;
              End If;
            End If;
          
            --更新病人余额 
            Update 病人余额
            Set 预交余额 = Nvl(预交余额, 0) + (-1 * n_结算金额)
            Where 病人id = v_退预交.病人id And 性质 = 1 And 类型 = 1
            Returning 预交余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (v_退预交.病人id, 1, -1 * n_结算金额, 1);
              n_返回值 := -1 * n_结算金额;
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 病人余额
              Where 病人id = v_退预交.病人id And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
            End If;
          
            Update 病人预交记录 Set 冲预交 = 冲预交 - n_结算金额 Where 结帐id = n_重结id And 结算方式 Is Null;
          
            n_返回值 := 1;
            If n_预交金额 = 0 Then
              Exit;
            End If;
          End Loop;
        End If;
      Else
        --退预交款 
        n_返回值   := 0;
        n_预交金额 := -1 * n_预交金额;
      
        For v_退预交 In (Select Max(a.Id) As ID, a.No, a.病人id, a.预交类别, Max(a.收款时间) As 收款时间, Sum(Nvl(a.冲预交, 0)) As 金额
                      From 病人预交记录 A,
                           (Select Distinct a.结帐id
                             From 门诊费用记录 A, 门诊费用记录 B
                             Where a.No = b.No And Mod(a.记录性质, 10) = 1 And b.结帐id = n_原结帐id) B
                      Where a.结帐id = b.结帐id And Mod(a.记录性质, 10) = 1 And Nvl(a.预交类别, 0) = 1
                      Group By a.No, a.病人id, a.预交类别
                      Having Sum(Nvl(a.冲预交, 0)) > 0
                      Order By 收款时间 Desc) Loop
        
          If v_退预交.金额 - n_预交金额 < 0 Then
            n_结算金额 := -1 * v_退预交.金额;
            n_预交金额 := n_预交金额 - v_退预交.金额;
          Else
            n_结算金额 := -1 * n_预交金额;
            n_预交金额 := 0;
          End If;
        
          Insert Into 病人预交记录
            (ID, 记录性质, NO, 实际票号, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号,
             校对标志, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 预交类别, 结算性质, 会话号, 关联交易id)
            Select 病人预交记录_Id.Nextval, 11, NO, 实际票号, 记录状态, 病人id, 主页id, 摘要, 结算方式, r_Balancedata.收款时间, r_Balancedata.操作员编号,
                   r_Balancedata.操作员姓名, n_结算金额, 冲销id_In, r_Balancedata.缴款组id, Sysdate, v_交易人员, r_Balancedata.结算序号, 2,
                   卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 预交类别, 3, n_会话号, 关联交易id
            From 病人预交记录
            Where ID = v_退预交.Id;
        
          --更新预交单据余额 
          Select Max(ID) Into n_充值id From 病人预交记录 Where NO = v_退预交.No And 记录性质 = 1 And 记录状态 <> 2;
          If Nvl(n_充值id, 0) <> 0 Then
            Update 预交单据余额
            Set 预交余额 = Nvl(预交余额, 0) + (-1 * n_结算金额)
            Where 病人id = v_退预交.病人id And 预交id = n_充值id
            Returning 预交余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 预交单据余额
                (预交id, 病人id, 预交类别, 预交余额)
              Values
                (n_充值id, v_退预交.病人id, Nvl(v_退预交.预交类别, 2), -1 * n_结算金额);
              n_返回值 := -1 * n_结算金额;
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 预交单据余额 Where 预交id = n_充值id And Nvl(预交余额, 0) = 0;
            End If;
          End If;
        
          --更新病人余额 
          Update 病人余额
          Set 预交余额 = Nvl(预交余额, 0) + (-1 * n_结算金额)
          Where 病人id = v_退预交.病人id And 性质 = 1 And 类型 = 1
          Returning 预交余额 Into n_返回值;
          If Sql%RowCount = 0 Then
            Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (v_退预交.病人id, 1, -1 * n_结算金额, 1);
            n_返回值 := -1 * n_结算金额;
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 病人余额
            Where 病人id = v_退预交.病人id And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
          End If;
        
          Update 病人预交记录 Set 冲预交 = 冲预交 - n_结算金额 Where 结帐id = 冲销id_In And 结算方式 Is Null;
        
          n_返回值 := 1;
          If n_预交金额 = 0 Then
            Exit;
          End If;
        End Loop;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        v_Err_Msg := '未找到原始的冲预交记录,不能回退预交款！';
        Raise Err_Item;
      End If;
    
      If Nvl(n_预交金额, 0) <> 0 Then
        v_Err_Msg := '当前退预交超过了收费结算中的冲预交款,不能回退预交款！';
        Raise Err_Item;
      End If;
    End If;
  
    n_预交金额 := 冲预交_In;
    If Nvl(n_预交金额, 0) > 0 Then
      --冲预交款 
      Zl_病人预交记录_冲预交(病人id_In, 冲销id_In, n_预交金额, 1, r_Balancedata.操作员编号, r_Balancedata.操作员姓名, r_Balancedata.收款时间,
                    冲预交病人ids_In, 3, 1);
    End If;
  End If;

  --1-普通退费方式 
  If 操作类型_In = 1 Then
    --各个收费结算 :格式为:"结算方式|结算金额|结算号码|结算摘要||.." 
    v_结算内容 := 结算方式_In || '||';
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      v_结算方式 := Nvl(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1), 缺省结算方式_In);
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      v_结算号码 := LTrim(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_结算摘要 := LTrim(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
    
      If Nvl(n_结算金额, 0) <> 0 Then
        n_结算金额 := Nvl(n_结算金额, 0);
        If Nvl(n_重结id, 0) <> 0 Then
          --1.先按此种方式全退 
          --2.再按此种方式收款 
          --3.本次退款=1+2 
          Select Sum(冲预交)
          Into n_返回值
          From 病人预交记录
          Where 结帐id = 冲销id_In And 结算方式 Is Null And Nvl(冲预交, 0) <> 0;
          If Nvl(n_返回值, 0) <> 0 Then
            Insert Into 病人预交记录
              (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志,
               卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号)
            Values
              (病人预交记录_Id.Nextval, 3, Null, 2, r_Balancedata.病人id, Null, v_结算摘要, v_结算方式, r_Balancedata.收款时间,
               r_Balancedata.操作员编号, r_Balancedata.操作员姓名, n_返回值, 冲销id_In, r_Balancedata.缴款组id, Sysdate, v_交易人员,
               r_Balancedata.结算序号, 2, Null, Null, 卡号_In, 交易流水号_In, 交易说明_In, Null, 3, n_会话号);
          
            Update 病人预交记录 Set 冲预交 = 冲预交 - Nvl(n_返回值, 0) Where 结帐id = 冲销id_In And 结算方式 Is Null;
          End If;
          n_结算金额 := n_结算金额 - Nvl(n_返回值, 0);
        
          If Nvl(n_结算金额, 0) <> 0 Then
            Insert Into 病人预交记录
              (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志,
               卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号)
            Values
              (病人预交记录_Id.Nextval, 3, Null, 1, r_Balancedata.病人id, Null, v_结算摘要, v_结算方式, r_Balancedata.收款时间,
               r_Balancedata.操作员编号, r_Balancedata.操作员姓名, n_结算金额, n_重结id, r_Balancedata.缴款组id, Sysdate, v_交易人员,
               r_Balancedata.结算序号, 2, Null, Null, 卡号_In, 交易流水号_In, 交易说明_In, Null, 3, n_会话号);
          
            Update 病人预交记录 Set 冲预交 = 冲预交 - n_结算金额 Where 结帐id = n_重结id And 结算方式 Is Null;
          End If;
        Else
          If Nvl(n_结算金额, 0) <> 0 Then
            Insert Into 病人预交记录
              (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志,
               卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号)
            Values
              (病人预交记录_Id.Nextval, 3, Null, 2, r_Balancedata.病人id, Null, v_结算摘要, v_结算方式, r_Balancedata.收款时间,
               r_Balancedata.操作员编号, r_Balancedata.操作员姓名, n_结算金额, 冲销id_In, r_Balancedata.缴款组id, Sysdate, v_交易人员,
               r_Balancedata.结算序号, 2, Null, Null, 卡号_In, 交易流水号_In, 交易说明_In, Null, 3, n_会话号);
          
            Update 病人预交记录 Set 冲预交 = 冲预交 - n_结算金额 Where 结帐id = 冲销id_In And 结算方式 Is Null;
          End If;
        End If;
      End If;
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  --2.三方卡退费结算 
  If 操作类型_In = 2 Then
    v_当前结算 := 结算方式_In;
    v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
    v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
    n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
    v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
    v_结算号码 := LTrim(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
    v_结算摘要 := LTrim(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
  
    If Nvl(n_结算金额, 0) <> 0 Then
      If Nvl(n_重结id, 0) <> 0 Then
        --1.先按此种方式全退 
        --2.再按此种方式收款 
        --3.本次退款=1+2 
        Select Sum(冲预交)
        Into n_返回值
        From 病人预交记录
        Where 结帐id = 冲销id_In And 结算方式 Is Null And Nvl(冲预交, 0) <> 0;
        If Nvl(n_返回值, 0) <> 0 Then
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志,
             卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号, 关联交易id)
          Values
            (n_预交id, 3, Null, 2, r_Balancedata.病人id, Null, v_结算摘要, v_结算方式, r_Balancedata.收款时间, r_Balancedata.操作员编号,
             r_Balancedata.操作员姓名, n_返回值, 冲销id_In, r_Balancedata.缴款组id, Sysdate, v_交易人员, r_Balancedata.结算序号, 2, 卡类别id_In,
             Null, 卡号_In, 交易流水号_In, 交易说明_In, Null, 3, n_会话号, 关联交易id_In);
        
          Update 病人预交记录 Set 冲预交 = 冲预交 - Nvl(n_返回值, 0) Where 结帐id = 冲销id_In And 结算方式 Is Null;
        
          --调用其他结算信息更新
          Zl_Custom_Balance_Update(n_预交id);
        End If;
        n_结算金额 := n_结算金额 - Nvl(n_返回值, 0);
      
        Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志,
           卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号, 关联交易id)
        Values
          (n_预交id, 3, Null, 1, r_Balancedata.病人id, Null, v_结算摘要, v_结算方式, r_Balancedata.收款时间, r_Balancedata.操作员编号,
           r_Balancedata.操作员姓名, n_结算金额, n_重结id, r_Balancedata.缴款组id, Sysdate, v_交易人员, r_Balancedata.结算序号, 2, 卡类别id_In,
           Null, 卡号_In, 交易流水号_In, 交易说明_In, Null, 3, n_会话号, 关联交易id_In);
      
        Update 病人预交记录 Set 冲预交 = 冲预交 - n_结算金额 Where 结帐id = n_重结id And 结算方式 Is Null;
      
        --调用其他结算信息更新
        Zl_Custom_Balance_Update(n_预交id);
      Else
        Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志,
           卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号, 关联交易id)
        Values
          (n_预交id, 3, Null, 2, r_Balancedata.病人id, Null, v_结算摘要, v_结算方式, r_Balancedata.收款时间, r_Balancedata.操作员编号,
           r_Balancedata.操作员姓名, n_结算金额, 冲销id_In, r_Balancedata.缴款组id, Sysdate, v_交易人员, r_Balancedata.结算序号, 2, 卡类别id_In,
           Null, 卡号_In, 交易流水号_In, 交易说明_In, v_结算号码, 3, n_会话号, 关联交易id_In);
      
        Update 病人预交记录 Set 冲预交 = 冲预交 - n_结算金额 Where 结帐id = 冲销id_In And 结算方式 Is Null;
      
        --调用其他结算信息更新
        Zl_Custom_Balance_Update(n_预交id);
      End If;
    End If;
  End If;

  --3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新) 
  If 操作类型_In = 3 Then
    --3.1检查是否已经存在医保结算数据,存在先删除 
    n_结算金额 := 0;
    For v_医保 In (Select ID, 结算方式, 冲预交
                 From 病人预交记录 A
                 Where 结帐id = 冲销id_In And Exists
                  (Select 1 From 结算方式 Where a.结算方式 = 名称 And 性质 In (3, 4)) And Mod(记录性质, 10) <> 1) Loop
      n_结算金额 := n_结算金额 + Nvl(v_医保.冲预交, 0);
      l_预交id.Extend;
      l_预交id(l_预交id.Count) := v_医保.Id;
    End Loop;
  
    Update 病人预交记录 Set 冲预交 = Nvl(冲预交, 0) + Nvl(n_结算金额, 0) Where 结帐id = 冲销id_In And 结算方式 Is Null;
  
    Forall I In 1 .. l_预交id.Count
      Delete 病人预交记录 Where ID = l_预交id(I);
  
    If 结算方式_In Is Not Null Then
      v_结算内容 := 结算方式_In || '||';
    End If;
  
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      n_结算金额 := To_Number(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
    
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志, 结算性质,
         会话号)
      Values
        (病人预交记录_Id.Nextval, 3, Null, 2, r_Balancedata.病人id, Null, '保险结算', v_结算方式, r_Balancedata.收款时间,
         r_Balancedata.操作员编号, r_Balancedata.操作员姓名, n_结算金额, 冲销id_In, r_Balancedata.缴款组id, Sysdate, v_交易人员,
         r_Balancedata.结算序号, 1, 3, n_会话号);
    
      --更新数据(结算方式为NULL的) 
      Update 病人预交记录 Set 冲预交 = 冲预交 - n_结算金额 Where 结帐id = 冲销id_In And 结算方式 Is Null;
    
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  --4-消费卡批量结算 
  If 操作类型_In = 4 Then
    Begin
      --获取上一次收款结帐ID 
      Select Max(a.结帐id)
      Into n_原结帐id
      From 门诊费用记录 A, (Select Distinct NO From 门诊费用记录 Where 结帐id = n_原结帐id) M
      Where a.No = m.No And Mod(a.记录性质, 10) = 1 And a.记录状态 In (1, 3) And Nvl(a.费用状态, 0) <> 1 And
            a.登记时间 + 0 =
            (Select Max(m.登记时间)
             From 门诊费用记录 M, (Select Distinct NO From 门诊费用记录 Where 结帐id = n_原结帐id) J
             Where m.No = j.No And Mod(m.记录性质, 10) = 1 And m.记录状态 In (1, 3) And Nvl(m.费用状态, 0) <> 1);
    
    Exception
      When Others Then
        v_Err_Msg := '未找到原结帐数据！';
        Raise Err_Item;
    End;
  
    v_结算内容 := 结算方式_In || '||';
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      --卡类别ID|卡号|消费卡ID|消费金额 
      n_卡类别id := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      v_卡号     := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_消费卡id := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_结算金额 := To_Number(v_当前结算);
    
      Begin
        Select 名称, 结算方式 Into v_名称, v_结算方式 From 消费卡类别目录 Where 编号 = n_卡类别id;
      Exception
        When Others Then
          v_名称 := Null;
      End;
      If v_名称 Is Null Then
        v_Err_Msg := '未找到对应的结算卡接口,本次刷卡消费失败!';
        Raise Err_Item;
      End If;
      If v_结算方式 Is Null Then
        v_Err_Msg := v_名称 || '未设置对应的结算方式,本次刷卡消费失败!';
        Raise Err_Item;
      End If;
      If Nvl(n_结算金额, 0) <> 0 Then
        n_结帐id := 冲销id_In;
      
        If Nvl(n_重结id, 0) <> 0 Then
          For c_记录 In (Select a.Id, c.接口编号, c.消费卡id, c.卡号, c.应收金额 As 结算金额
                       From 病人预交记录 A, 病人卡结算记录 C
                       Where a.Id = c.结算id And a.记录性质 = 3 And a.记录状态 In (1, 3) And a.结帐id = n_原结帐id And c.接口编号 = n_卡类别id And
                             c.消费卡id = n_消费卡id) Loop
          
            If Nvl(c_记录.结算金额, 0) <> 0 Then
              Update 病人预交记录
              Set 冲预交 = Nvl(冲预交, 0) + c_记录.结算金额
              Where 结帐id = 冲销id_In And 结算方式 = v_结算方式 And 结算卡序号 = n_卡类别id
              Returning ID Into n_预交id;
              If Sql%NotFound Then
                Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
                Insert Into 病人预交记录
                  (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号,
                   结算卡序号, 校对标志, 结算性质, 会话号)
                Values
                  (n_预交id, 3, Null, 2, r_Balancedata. 病人id, Null, Null, v_结算方式, r_Balancedata.收款时间, r_Balancedata.操作员编号,
                   r_Balancedata. 操作员姓名, c_记录.结算金额, 冲销id_In, r_Balancedata. 缴款组id, Sysdate, v_交易人员, r_Balancedata.结算序号,
                   n_卡类别id, 2, 3, n_会话号);
              End If;
            
              Update 病人预交记录 Set 冲预交 = 冲预交 - c_记录.结算金额 Where 结帐id = 冲销id_In And 结算方式 Is Null;
            
              --插入卡结算记录 
              Zl_病人卡结算记录_退款(c_记录.接口编号, c_记录.卡号, c_记录.消费卡id, -1 * c_记录.结算金额, c_记录.Id, n_预交id, r_Balancedata. 操作员编号,
                            r_Balancedata. 操作员姓名, r_Balancedata. 收款时间);
            
              n_结算金额 := n_结算金额 - c_记录.结算金额;
            End If;
          End Loop;
          n_结帐id := n_重结id;
        End If;
      
        Update 病人预交记录
        Set 冲预交 = Nvl(冲预交, 0) + n_结算金额
        Where 结帐id = n_结帐id And 结算方式 = v_结算方式 And 结算卡序号 = n_卡类别id
        Returning ID Into n_预交id;
        If Sql%NotFound Then
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 结算卡序号,
             校对标志, 结算性质, 会话号)
          Values
            (n_预交id, 3, Null, Decode(Nvl(n_重结id, 0), 0, 2, 1), r_Balancedata. 病人id, Null, Null, v_结算方式,
             r_Balancedata. 收款时间, r_Balancedata. 操作员编号, r_Balancedata. 操作员姓名, n_结算金额, n_结帐id, r_Balancedata. 缴款组id,
             Sysdate, v_交易人员, r_Balancedata. 结算序号, n_卡类别id, 2, 3, n_会话号);
        End If;
      
        If Nvl(n_重结id, 0) = 0 Then
          Begin
            Select ID Into n_原预交id From 病人预交记录 A Where 结帐id = n_原结帐id And 结算卡序号 = n_卡类别id;
          Exception
            When Others Then
              v_Err_Msg := '未找到原结算记录！';
              Raise Err_Item;
          End;
        
          Zl_病人卡结算记录_退款(n_卡类别id, v_卡号, n_消费卡id, -1 * n_结算金额, n_原预交id, n_预交id, r_Balancedata. 操作员编号,
                        r_Balancedata. 操作员姓名, r_Balancedata. 收款时间);
        Else
          Zl_病人卡结算记录_支付(n_卡类别id, v_卡号, n_消费卡id, n_结算金额, n_预交id, r_Balancedata. 操作员编号, r_Balancedata. 操作员姓名,
                        r_Balancedata. 收款时间);
        End If;
      
        --更新数据(结算方式为NULL的) 
        Update 病人预交记录 Set 冲预交 = 冲预交 - n_结算金额 Where 结帐id = n_结帐id And 结算方式 Is Null;
      End If;
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  --5-三方卡医保结算 
  If 操作类型_In = 5 Then
    If Nvl(删除原结算_In, 0) = 1 Then
      --1.1检查是否已经存在三方卡结算数据,存在先删除 
      n_结算金额 := 0;
      For c_结算 In (Select ID, 结算方式, 冲预交
                   From 病人预交记录 A
                   Where 结帐id = 冲销id_In And 卡类别id = 卡类别id_In And Nvl(关联交易id, 0) = Nvl(关联交易id_In, 0)) Loop
        n_结算金额 := n_结算金额 + Nvl(c_结算.冲预交, 0);
        l_预交id.Extend;
        l_预交id(l_预交id.Count) := c_结算.Id;
      End Loop;
    
      Update 病人预交记录
      Set 冲预交 = Nvl(冲预交, 0) + Nvl(n_结算金额, 0)
      Where 结帐id = 冲销id_In And 结算方式 Is Null;
    
      Forall I In 1 .. l_预交id.Count
        Delete From 病人预交记录 Where ID = l_预交id(I);
    
      Delete From 医保结算明细
      Where 结帐id = 冲销id_In And 卡类别id = 卡类别id_In And Nvl(关联交易id, 0) = Nvl(关联交易id_In, 0);
    End If;
  
    --格式：结算方式|结算金额|结算号码|结算摘要|单据号|是否普通结算 
    For c_结算 In (Select Max(Decode(序号, 1, 值, Null)) As 结算方式, zl_To_Number(Max(Decode(序号, 2, 值, ''))) As 结算金额,
                        Trim(Max(Decode(序号, 3, 值, ''))) As 结算号码, Trim(Max(Decode(序号, 4, 值, ''))) As 结算摘要,
                        Trim(Max(Decode(序号, 5, 值, ''))) As 单据号, zl_To_Number(Max(Decode(序号, 6, 值, ''))) As 是否普通结算
                 From (Select Rownum As 序号, Column_Value As 值 From Table(f_Str2List(结算方式_In, '|')))
                 Having Nvl(zl_To_Number(Max(Decode(序号, 2, 值, ''))), 0) <> 0) Loop
    
      Update 病人预交记录
      Set 冲预交 = 冲预交 + c_结算.结算金额
      Where 结帐id = 冲销id_In And 卡类别id = 卡类别id_In And 结算方式 = c_结算.结算方式 And Nvl(关联交易id, 0) = Nvl(关联交易id_In, 0)
      Returning ID Into n_预交id;
      If Sql%NotFound Then
        Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志,
           卡类别id, 卡号, 关联交易id, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号)
        Values
          (n_预交id, 3, Null, 2, r_Balancedata.病人id, Null, c_结算.结算摘要, c_结算.结算方式, r_Balancedata.收款时间, r_Balancedata.操作员编号,
           r_Balancedata.操作员姓名, c_结算.结算金额, 冲销id_In, r_Balancedata.缴款组id, Sysdate, v_交易人员, r_Balancedata.结算序号, 校对标志_In,
           Decode(c_结算.是否普通结算, 1, Null, 卡类别id_In), Decode(c_结算.是否普通结算, 1, Null, 卡号_In), 关联交易id_In, 交易流水号_In, 交易说明_In,
           c_结算.结算号码, 3, n_会话号);
      End If;
    
      Update 病人预交记录 Set 冲预交 = 冲预交 - c_结算.结算金额 Where 结帐id = 冲销id_In And 结算方式 Is Null;
    
      If c_结算.单据号 Is Not Null Then
        Insert Into 医保结算明细
          (结帐id, NO, 结算方式, 金额, 卡类别id, 关联交易id, 交易流水号, 交易说明)
        Values
          (冲销id_In, c_结算.单据号, c_结算.结算方式, c_结算.结算金额, Decode(c_结算.是否普通结算, 1, Null, 卡类别id_In), 关联交易id_In, 交易流水号_In,
           交易说明_In);
      End If;
    
      --调用其他结算信息更新
      Zl_Custom_Balance_Update(n_预交id);
    End Loop;
  End If;

  If Nvl(完成退费_In, 0) = 0 Then
    Return;
  End If;

  ----------------------------------------------------------------------------------------- 
  --完成收费,需要处理人员缴款余额,预交记录(结算方式=NULL) 
  -- 完成退费_In:0-未完成退费;1-异常完成退费;2-完成退费 
  If Nvl(完成退费_In, 0) = 1 Then
    --必须已全部原样退后才能完成作废 
    Select Count(1)
    Into n_Count
    From 病人预交记录 A
    Where a.结帐id = n_原结帐id And Nvl(a.校对标志, 0) = 2 And Nvl(a.冲预交, 0) <> 0 And Not Exists
     (Select 1
           From 病人预交记录
           Where 结帐id = 冲销id_In And 结算方式 = a.结算方式 And Nvl(关联交易id, 0) = Nvl(a.关联交易id, 0) And Nvl(校对标志, 0) = 2);
    If n_Count <> 0 Then
      v_Err_Msg := '还存在未作废的交易，不能完成作废！';
      Raise Err_Item;
    End If;
  
    Update 病人预交记录 Set 校对标志 = 0, 会话号 = Null Where 结帐id = 冲销id_In;
  
    If Nvl(n_异步结算, 0) <> 0 Then
      --恢复划价单状态 
    
      --将未发药品记录标记为未收费状态 
      For c_No In (Select Distinct NO From 门诊费用记录 Where 记录性质 = 1 And 结帐id = 冲销id_In) Loop
        Update 未发药品记录 A
        Set 已收费 = 0
        Where 单据 In (8, 24) And NO = c_No.No And Exists
         (Select 1
               From 药品收发记录
               Where 单据 = a.单据 And Nvl(库房id, 0) = Nvl(a.库房id, 0) And NO = c_No.No And Mod(记录状态, 3) = 1 And 审核人 Is Null);
      End Loop;
    
      Delete From 门诊费用记录 Where 记录性质 = 1 And 结帐id = 冲销id_In;
    
      Update 门诊费用记录
      Set 记录状态 = 1, 结帐id = Null, 结帐金额 = Null, 操作员编号 = Null, 操作员姓名 = Null, 缴款组id = Null
      Where 记录性质 = 1 And 结帐id = n_原结帐id;
    
      --标记原预交记录 
      Update 病人预交记录 Set 记录状态 = 3 Where 结帐id = n_原结帐id And Mod(记录性质, 10) <> 1;
    End If;
    Return;
  End If;

  Select Max(a.是否电子票据)
  Into n_是否电子票据
  From 病人预交记录 A
  Where a.结帐id = n_原结帐id And a.记录性质 In (11, 3);

  --1.相关业务数据处理 
  --  门诊费用转住院退费异常完成退费时，不处理相关业务数据
  Select Count(1)
  Into n_门诊转住院退费
  From 病人预交记录
  Where 结帐id = 冲销id_In And 结算方式 Is Null And Nvl(附加标志, 0) = 2 And Rownum < 2;
  If Nvl(n_异步结算, 0) = 1 And Nvl(n_门诊转住院退费, 0) = 0 Then
    Zl_门诊收费记录_完成退费(冲销id_In, r_Balancedata.操作员姓名, n_重结id, n_是否电子票据);
  End If;

  --2.删除结算方式为NULL的预交记录 
  --结算方式为NULL的冲销记录和重结记录的金额之和为零，说明已完成全部结算 
  If Nvl(n_重结id, 0) <> 0 Then
    Select Sum(Nvl(冲预交, 0))
    Into n_冲预交
    From 病人预交记录
    Where 结帐id In (冲销id_In, n_重结id) And 结算方式 Is Null;
    If Nvl(n_冲预交, 0) <> 0 Then
      v_Err_Msg := '还存在未缴款的数据,不能完成结算!';
      Raise Err_Item;
    Else
      Delete 病人预交记录 Where 结帐id = 冲销id_In And 结算方式 Is Null And Nvl(冲预交, 0) = 0;
      If Sql%NotFound Then
        Update 病人预交记录 Set 结算方式 = v_退费结算 Where 结帐id = 冲销id_In And 结算方式 Is Null;
        If Sql%NotFound Then
          v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[退费结算窗口]中重新收费！!';
          Raise Err_Item;
        End If;
      End If;
    
      Delete 病人预交记录 Where 结帐id = n_重结id And 结算方式 Is Null And Nvl(冲预交, 0) = 0;
      If Sql%NotFound Then
        Update 病人预交记录 Set 结算方式 = v_退费结算 Where 结帐id = n_重结id And 结算方式 Is Null;
        If Sql%NotFound Then
          v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[退费结算窗口]中重新收费！!';
          Raise Err_Item;
        End If;
      End If;
    End If;
    Update 病人预交记录 Set 缴款 = 缴款_In, 找补 = 找补_In, 校对标志 = 0 Where 结帐id = n_重结id;
  Else
    Delete 病人预交记录 Where 结帐id = 冲销id_In And 结算方式 Is Null And Nvl(冲预交, 0) = 0;
    If Sql%NotFound Then
      Select Count(1) Into n_Count From 病人预交记录 A Where 结帐id = 冲销id_In And 结算方式 Is Null;
      If n_Count <> 0 Then
        v_Err_Msg := '还存在未缴款的数据,不能完成结算!';
      Else
        v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[退费结算窗口]中重新收费！!';
      End If;
      Raise Err_Item;
    End If;
  End If;

  --3.检查门诊费用记录与病人预交记录的金额是否相等 
  n_结算金额 := 0;
  n_冲预交   := 0;
  Select Nvl(Sum(实收金额), 0)
  Into n_结算金额
  From 门诊费用记录
  Where 结帐id In (Select 结帐id From 病人预交记录 Where 结算序号 = n_结算序号);
  Select Nvl(Sum(冲预交), 0) Into n_冲预交 From 病人预交记录 Where 结算序号 = n_结算序号;
  If n_结算金额 <> n_冲预交 Then
    v_Err_Msg := '结算信息有误，实收金额(' || n_结算金额 || ')与结算金额(' || n_冲预交 || ')不一致，不能完成结算！';
    Raise Err_Item;
  End If;

  --4.结算金额为零时，增加一条金额为0的病人预交记录 
  Select Count(1) Into n_Count From 病人预交记录 A Where 结帐id = 冲销id_In;
  If n_Count = 0 Then
    v_结算方式 := 缺省结算方式_In;
    If v_结算方式 Is Null Then
      Begin
        Select 结算方式 Into v_结算方式 From 结算方式应用 Where 应用场合 = '收费' And Nvl(缺省标志, 0) = 1;
      Exception
        When Others Then
          v_结算方式 := Null;
      End;
      If v_结算方式 Is Null Then
        Begin
          Select 名称 Into v_结算方式 From 结算方式 Where Nvl(性质, 0) = 1;
        Exception
          When Others Then
            v_结算方式 := '现金';
        End;
      End If;
    End If;
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志, 卡类别id,
       结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号)
    Values
      (病人预交记录_Id.Nextval, 3, Null, 1, r_Balancedata.病人id, Null, Null, v_结算方式, r_Balancedata.收款时间, r_Balancedata.操作员编号,
       r_Balancedata.操作员姓名, 0, 冲销id_In, r_Balancedata.缴款组id, Sysdate, v_交易人员, r_Balancedata.结算序号, 2, Null, Null, Null,
       Null, 交易说明_In, Null, 3, n_会话号);
  End If;

  --5.处理缴款数据和找补数据及校对标志更新为0，会话号更新为NULL  
  Update 病人预交记录
  Set 缴款 = 缴款_In, 找补 = 找补_In, 校对标志 = 0, 会话号 = Null, 是否电子票据 = n_是否电子票据
  Where 结帐id In (冲销id_In, n_重结id);

  --6.更新费用状态 
  Update 门诊费用记录 Set 费用状态 = 0 Where 结帐id In (冲销id_In, n_重结id);

  --7.更新人员缴款数据 
  For c_缴款 In (Select 结算方式, 操作员姓名, Sum(Nvl(a.冲预交, 0)) As 冲预交
               From 病人预交记录 A
               Where a.结帐id In (冲销id_In, n_重结id) And Mod(a.记录性质, 10) <> 1
               Group By 结算方式, 操作员姓名) Loop
  
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) + Nvl(c_缴款.冲预交, 0)
    Where 收款员 = c_缴款.操作员姓名 And 性质 = 1 And 结算方式 = c_缴款.结算方式;
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额
        (收款员, 结算方式, 性质, 余额)
      Values
        (c_缴款.操作员姓名, c_缴款.结算方式, 1, Nvl(c_缴款.冲预交, 0));
    End If;
  End Loop;

  --8.消息集成处理 
  b_Message.Zlhis_Charge_004(1, 冲销id_In);
  If Nvl(n_重结id, 0) <> 0 Then
    b_Message.Zlhis_Charge_002(1, n_重结id);
  End If;

  --9.消息推送 
  Select 病人id_In || ',' || 冲销id_In || ',' || Decode(完成退费_In, 2, 0, 0, 0, 1) Into v_Msg From Dual;
  Begin
    Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
      Using 5, v_Msg;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊退费结算_Modify;
/
Create Or Replace Procedure Zl_门诊收费记录_完成退费
(
  冲销id_In     病人预交记录.结帐id%Type,
  操作员姓名_In 病人预交记录.操作员姓名%Type,
  重收结帐id_In 病人预交记录.结帐id%Type := Null,
  电子票据_In   病人预交记录.是否电子票据%Type := 0
) As
  --功能：门诊退费完成后，异步结算时处理相关业务数据
  n_原结帐id 病人预交记录.结帐id%Type;
  n_剩余数量 门诊费用记录.数次%Type;
  n_准退数量 门诊费用记录.数次%Type;
  n_执行状态 门诊费用记录.执行状态%Type;

  v_Para         Varchar2(1000);
  n_启用模式     Number(3);
  n_分别打印     Number;
  n_Onepatiprint Number;
  n_打印id       票据打印内容.Id%Type;
  l_使用id       t_NumList := t_NumList();
  n_回收票据     Number;
  n_部分退       Number;

  n_Count Number;
  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin
  --求原结帐ID
  Select Max(结帐id)
  Into n_原结帐id
  From (With c_单据 As (Select NO From 门诊费用记录 Where Mod(记录性质, 10) = 1 And 结帐id = 冲销id_In)
         Select Max(a.结帐id) As 结帐id
         From 门诊费用记录 A, c_单据 M
         Where a.No = m.No And Mod(a.记录性质, 10) = 1 And a.记录状态 In (1, 3) And Nvl(a.费用状态, 0) <> 1 And
               a.登记时间 + 0 =
               (Select Max(m.登记时间)
                From 门诊费用记录 M, c_单据 J
                Where m.No = j.No And Mod(m.记录性质, 10) = 1 And m.记录状态 In (1, 3) And Nvl(m.费用状态, 0) <> 1));


  If Nvl(重收结帐id_In, 0) <> 0 Then
    --存在重收时，原记录是被全部冲销了的
    Update 门诊费用记录 Set 记录状态 = 3 Where Mod(记录性质, 10) = 1 And 记录状态 = 1 And 结帐id = n_原结帐id;
  End If;
  --标记原预交记录
  Update 病人预交记录 Set 记录状态 = 3 Where 结帐id = n_原结帐id And Mod(记录性质, 10) <> 1;

  --标记退款时一卡通的收款记录
  Update 病人预交记录
  Set 记录状态 = 3
  Where 记录性质 Not In (1, 11) And 记录状态 = 1 And 结帐id <> 冲销id_In And
        (卡类别id, 关联交易id) In (Select 卡类别id, 关联交易id
                            From 病人预交记录
                            Where 记录性质 Not In (1, 11) And 结帐id = 冲销id_In And 卡类别id Is Not Null);

  --必须按照“收费细目id”升序排序，防止并发，锁“药品库存”表
  For c_No In (Select NO, 序号
               From 门诊费用记录
               Where Mod(记录性质, 10) = 1 And 结帐id In (冲销id_In, 重收结帐id_In)
               Group By NO, 序号, 收费细目id
               Having Nvl(Sum(Nvl(付数, 1) * 数次), 0) <> 0
               Order By 收费细目id) Loop
  
    For r_Bill In (Select a.Id, a.No, a.收费类别, a.医嘱序号, a.序号, a.价格父号, a.收费细目id, a.执行状态, j.诊疗类别, m.跟踪在用,
                          Nvl(j.医嘱状态, 0) As 医嘱状态
                   From 门诊费用记录 A, 病人医嘱记录 J, 材料特性 M
                   Where a.医嘱序号 = j.Id(+) And a.收费细目id + 0 = m.材料id(+) And a.记录性质 = 1 And a.记录状态 In (1, 3) And
                         Nvl(a.执行状态, 0) <> 1 And a.No = c_No.No And a.序号 = c_No.序号) Loop
    
      --药品卫材相关内容
      If Instr(',4,5,6,7,', r_Bill.收费类别) > 0 Then
        Zl_药品收发记录_销售退费(r_Bill.Id);
      End If;
    
      Select Nvl(Sum(Nvl(付数, 1) * 数次), 0)
      Into n_剩余数量
      From 门诊费用记录
      Where Mod(记录性质, 10) = 1 And NO = r_Bill.No And 序号 = r_Bill.序号;
    
      n_准退数量 := 0;
      --准退数量(非药品项目为剩余数量,原始数量)
      If Instr(',4,5,6,7,', r_Bill.收费类别) = 0 Or (r_Bill.收费类别 = '4' And Nvl(r_Bill.跟踪在用, 0) = 0) Then
        --非药品部分(以具体医嘱执行为准进行检查)
        --: 1.存在医嘱执行计价的,则以医嘱执行为准(但不能包含:检查;检验;手术;麻醉及输血),已执行的不允许退费
        --: 2.不存在医嘱执行计价的,则以剩余数量为准
        --: 3.医嘱作废了的,则以剩余数量为准(病人医嘱记录.医嘱状态=4表示作废医嘱，会删除"病人医嘱发送",门诊药嘱先作废后退药)
        --: 4.病人医嘱发送.执行状态=1（完成执行）时，准退数为0，不再根据医嘱执行计价来统计准退数
        If Instr(',C,D,F,G,K,', ',' || r_Bill.诊疗类别 || ',') = 0 And r_Bill.诊疗类别 Is Not Null And r_Bill.医嘱状态 <> 4 Then
          Select Nvl(Sum(Decode(b.执行状态, 1, 0, 1) * Decode(c.执行状态, 0, 1, 0) * c.数量), 0)
          Into n_准退数量
          From 病人医嘱发送 B, 医嘱执行计价 C
          Where b.医嘱id = r_Bill.医嘱序号 And b.No = r_Bill.No And b.医嘱id = c.医嘱id And b.发送号 = c.发送号 And
                c.收费细目id + 0 = r_Bill.收费细目id And b.记录性质 = 1;
        End If;
      Else
        Select Nvl(Sum(Nvl(付数, 1) * 实际数量), 0)
        Into n_准退数量
        From 药品收发记录
        Where 单据 In (8, 24) And Mod(记录状态, 3) = 1 And 审核人 Is Null And NO = r_Bill.No And 费用id = r_Bill.Id;
      End If;
    
      --标记原费用记录
      n_执行状态 := Case
                  When n_剩余数量 = n_准退数量 Then
                   0
                  When n_准退数量 = 0 Then
                   1
                  Else
                   2
                End;
      Update 门诊费用记录
      Set 执行状态 = n_执行状态
      Where Mod(记录性质, 10) = 1 And 记录状态 In (1, 3) And NO = c_No.No And 序号 = c_No.序号;
    
      Update 门诊费用记录
      Set 记录状态 = 3
      Where Mod(记录性质, 10) = 1 And 记录状态 = 1 And 结帐id = n_原结帐id And NO = c_No.No And 序号 = c_No.序号;
    End Loop;
  End Loop;

  If Nvl(电子票据_In, 0) = 0 Then
    --启用标志||NO;执行科室(条数);收据费目(首页汇总,条数);收费细目(条数)
    v_Para     := Nvl(zl_GetSysParameter('票据分配规则', 1121), '0||0;0;0,0;0');
    n_启用模式 := zl_To_Number(Substr(v_Para, 1, 1));
    n_分别打印 := Nvl(zl_GetSysParameter('多张单据收费分别打印', 1121), '0');
  End If;

  For c_No In (Select NO
               From 门诊费用记录
               Where Mod(记录性质, 10) = 1 And 结帐id In (冲销id_In, 重收结帐id_In)
               Group By NO, 序号
               Having Nvl(Sum(Nvl(付数, 1) * 数次), 0) <> 0) Loop
  
    If Nvl(电子票据_In, 0) = 0 Then
      --是否按病人一次打印的
      Select Count(1)
      Into n_Onepatiprint
      From 票据打印内容 A1, 票据打印内容 A2
      Where A1.Id = A2.Id And A1.数据性质 = A2.数据性质 And A1.No = c_No.No And A1.数据性质 = 1 And Nvl(A2.打印类型, 0) = 1;
    
      --退费票据回收(仅全退时才回收，部分退时在重打过程中回收)
      If n_启用模式 = 0 And n_分别打印 = 1 And n_Onepatiprint = 0 Then
        Select Count(1)
        Into n_部分退
        From (Select 1
               From 门诊费用记录 A
               Where Mod(a.记录性质, 10) = 1 And a.No = c_No.No
               Group By a.No, a.序号
               Having Nvl(Sum(Nvl(a.付数, 1) * a.数次), 0) <> 0);
      Else
        Select Count(1)
        Into n_部分退
        From (Select 1
               From 门诊费用记录 A
               Where a.No In
                     (Select a.No
                      From 门诊费用记录 A, 门诊费用记录 B
                      Where a.结帐id = b.结帐id And Mod(a.记录性质, 10) = 1 And Mod(b.记录性质, 10) = 1 And b.No = c_No.No) And
                     Mod(a.记录性质, 10) = 1
               Group By a.No, a.序号
               Having Nvl(Sum(Nvl(a.付数, 1) * a.数次), 0) <> 0);
      End If;
    
      If n_启用模式 <> 0 And n_Onepatiprint <> 0 And n_部分退 = 0 Then
        n_回收票据 := 0;
      Elsif n_部分退 = 0 Then
        n_回收票据 := 1;
      Else
        n_回收票据 := 0;
      End If;
    
      If n_回收票据 = 1 Then
        If n_启用模式 <> 0 Then
          --收回票据
          Select 使用id
          Bulk Collect
          Into l_使用id
          From (Select Distinct b.使用id From 票据打印明细 B Where b.No = c_No.No And Nvl(b.票种, 0) = 1);
        
          n_启用模式 := l_使用id.Count;
          If l_使用id.Count <> 0 Then
            --插入回收记录
            Forall I In 1 .. l_使用id.Count
              Insert Into 票据使用明细
                (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用人, 使用时间, 票据金额)
                Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, 操作员姓名_In, Sysdate, 票据金额
                From 票据使用明细 A
                Where ID = l_使用id(I) And 性质 = 1 And Not Exists
                 (Select 1 From 票据使用明细 Where 号码 = a.号码 And 票种 = a.票种 And Nvl(性质, 0) <> 1);
          
            Forall I In 1 .. l_使用id.Count
              Update 票据打印明细 Set 是否回收 = 1 Where 使用id = l_使用id(I) And Nvl(是否回收, 0) = 0;
          
          End If;
        End If;
        If n_启用模式 = 0 Then
          --获取单据最后一次的打印ID(可能是多张单据收费打印)
          --性质=1，原因=6为退费打印票据(红票)，不回收
          Select Max(ID)
          Into n_打印id
          From (Select b.Id
                 From 票据使用明细 A, 票据打印内容 B
                 Where a.打印id = b.Id And a.性质 = 1 And a.原因 In (1, 3) And b.数据性质 = 1 And b.No = c_No.No
                 Order By a.使用时间 Desc)
          Where Rownum < 2;
        
          --可能以前没有打印,无收回
          If n_打印id Is Not Null Then
            --a.多张单据循环调用时只能收回一次
            Select Count(1) Into n_Count From 票据使用明细 Where 票种 = 1 And 性质 = 2 And 打印id = n_打印id;
            If n_Count = 0 Then
              Insert Into 票据使用明细
                (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
                Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, Sysdate, 操作员姓名_In, 票据金额
                From 票据使用明细
                Where 打印id = n_打印id And 票种 = 1 And 性质 = 1;
            Else
              --b.部分退费多次收回时,最后一次全退收回要排开已收回的
              Insert Into 票据使用明细
                (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
                Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, Sysdate, 操作员姓名_In, 票据金额
                From 票据使用明细 A
                Where 打印id = n_打印id And 票种 = 1 And 性质 = 1 And Not Exists
                 (Select 1
                       From 票据使用明细 B
                       Where a.号码 = b.号码 And 打印id = n_打印id And 票种 = 1 And 性质 = 2);
            End If;
          End If;
        End If;
      End If;
    End If;
  
    --清除医嘱执行计价.费用ID
    For c_费用 In (Select Distinct a.Id, a.医嘱序号 As 医嘱id, a.收费细目id, b.发送号
                 From 门诊费用记录 A, 病人医嘱发送 B
                 Where a.医嘱序号 = b.医嘱id And a.No = b.No And a.结帐id = 冲销id_In And a.价格父号 Is Null And b.记录性质 = 1) Loop
      Update 医嘱执行计价
      Set 费用id = Null
      Where 医嘱id = c_费用.医嘱id And 发送号 = c_费用.发送号 And 收费细目id = c_费用.收费细目id And 执行状态 = 2 And 费用id = c_费用.Id;
    End Loop;
  
    --删除病人医嘱附费(最后一次删除时)
    For c_医嘱 In (Select Distinct 医嘱序号
                 From 门诊费用记录
                 Where Mod(记录性质, 10) = 1 And 记录状态 In (1, 3) And 医嘱序号 Is Not Null And NO = c_No.No) Loop
    
      Select Count(1)
      Into n_Count
      From (Select 1
             From 门诊费用记录
             Where Mod(记录性质, 10) = 1 And 医嘱序号 + 0 = c_医嘱.医嘱序号 And NO = c_No.No
             Group By 序号
             Having Nvl(Sum(Nvl(付数, 1) * 数次), 0) <> 0);
    
      If n_Count = 0 Then
        Delete From 病人医嘱附费 Where 医嘱id = c_医嘱.医嘱序号 And 记录性质 = 1 And NO = c_No.No;
      End If;
    End Loop;
  
    --场合_In    Integer:=0, --0:门诊;1-住院
    --性质_In    Integer:=1, --1-收费单;2-记帐单
    --操作_In    Integer:=0, --0:删除划价单;1-收费或记帐;2-退费或销帐
    --No_In      门诊费用记录.No%Type,
    --医嘱ids_In Varchar2 := Null
    Zl_医嘱发送_计费状态_Update(0, 1, 2, c_No.No);
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊收费记录_完成退费;
/
Create Or Replace Procedure Zl_门诊简单收费_Delete
(
  No_In         门诊费用记录.No%Type,
  操作员编号_In 门诊费用记录.操作员编号%Type,
  操作员姓名_In 门诊费用记录.操作员姓名%Type
) As
  --功能：删除一张门诊简单收费单据

  --该游标为要退费单据的所有原始记录
  Cursor c_Bill Is
    Select a.Id, a.No, a.附加标志, a.收费细目id, a.序号, a.价格父号, a.执行状态, a.收费类别, a.付数, a.数次, a.医嘱序号, j.诊疗类别, m.跟踪在用
    From 门诊费用记录 A, 病人医嘱记录 J, 材料特性 M
    Where a.医嘱序号 = j.Id(+) And a.No = No_In And a.记录性质 = 1 And a.记录状态 In (1, 3) And a.收费细目id + 0 = m.材料id(+)
    Order By a.收费细目id, a.序号;

  --该光标用于处理人员缴款余额中退的不同结算方式的金额
  Cursor c_Money(冲销id_In 病人预交记录.结帐id%Type) Is
    Select 结算方式, 冲预交
    From 病人预交记录
    Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = 冲销id_In And 结算方式 Is Not Null And Nvl(冲预交, 0) <> 0 And Nvl(校对标志, 0) = 0;

  --该游标用于查找收费时使用过的冲预交款记录
  Cursor c_Deposit(V结帐id 病人预交记录.结帐id%Type) Is
    Select NO, ID, 病人id, 冲预交 As 金额, 预交类别
    From 病人预交记录
    Where 记录性质 In (1, 11) And 记录状态 In (1, 3) And 结帐id = V结帐id And Nvl(冲预交, 0) <> 0
    Order By ID Desc;

  n_病人id   病人信息.病人id%Type;
  n_结帐id   门诊费用记录.结帐id%Type;
  n_结算序号 病人预交记录.结算序号%Type;
  n_打印id   票据打印内容.Id%Type;

  n_预交金额 病人预交记录.冲预交%Type;
  n_返回值   病人预交记录.冲预交%Type;
  n_充值id   病人预交记录.Id%Type;
  --部分退费计算变量
  n_剩余数量 Number;
  n_剩余应收 Number;
  n_剩余实收 Number;
  n_剩余统筹 Number;
  n_准退数量 Number;
  n_退费次数 Number;

  n_应收金额 Number;
  n_实收金额 Number;
  n_统筹金额 Number;
  n_总金额   Number;
  n_费用状态 门诊费用记录.费用状态%Type;
  n_正常退费 Number; --是否第一次退费且全部退费,在每行退费过程中判断得到。
  n_组id     财务缴款分组.Id%Type;

  v_退费结算 结算方式.名称%Type;
  l_使用id   t_NumList := t_NumList();
  n_Dec      Number;
  d_Date     Date;
  n_Count    Number;
  n_原结帐id Number;

  Err_Item Exception;
  v_Err_Msg      Varchar2(255);
  n_启用模式     Number(3);
  v_Para         Varchar2(1000);
  n_医属执行计价 Number;
  n_是否电子票据 病人预交记录.是否电子票据%Type;
Begin
  n_组id := Zl_Get组id(操作员姓名_In);

  --是否已经全部完全执行(只是该单据整张单据的检查)
  Select Nvl(Count(*), 0)
  Into n_Count
  From 门诊费用记录
  Where NO = No_In And 记录性质 = 1 And 记录状态 In (1, 3) And Nvl(执行状态, 0) <> 1;
  If n_Count = 0 Then
    v_Err_Msg := '该单据中的项目已经全部完全执行！';
    Raise Err_Item;
  End If;

  --未完全执行的项目是否有剩余数量(只是整张单据的检查)
  --执行状态在原始记录上判断
  Select Nvl(Count(*), 0)
  Into n_Count
  From (Select 序号, Sum(数量) As 剩余数量
         From (Select 记录状态, Nvl(价格父号, 序号) As 序号, Avg(Nvl(付数, 1) * 数次) As 数量
                From 门诊费用记录
                Where NO = No_In And 记录性质 = 1 And Nvl(附加标志, 0) <> 9 And
                      Nvl(价格父号, 序号) In
                      (Select Nvl(价格父号, 序号)
                       From 门诊费用记录
                       Where NO = No_In And 记录性质 = 1 And 记录状态 In (1, 3) And Nvl(执行状态, 0) <> 1)
                Group By 记录状态, Nvl(价格父号, 序号))
         Group By 序号
         Having Sum(数量) <> 0);
  If n_Count = 0 Then
    v_Err_Msg := '该单据中未完全执行部分项目剩余数量为零,没有可以退费的费用！';
    Raise Err_Item;
  End If;
  --确定是否在医嘱执行计价中存在数据,如果存在数据,则根据医嘱执行计价进行退费,否则按旧方式进行处理
  Select Count(1)
  Into n_医属执行计价
  From 门诊费用记录 A, 医嘱执行计价 B
  Where a.医嘱序号 = b.医嘱id And a.记录性质 = 1 And a.No = No_In And a.记录状态 In (1, 3) And Rownum = 1;

  ---------------------------------------------------------------------------------
  --公用变量
  Select Sysdate Into d_Date From Dual;
  Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
  n_结算序号 := Null;

  --金额小数位数
  Select zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  --获取结算方式名称
  Begin
    Select 名称 Into v_退费结算 From 结算方式 Where 性质 = 1;
  Exception
    When Others Then
      v_退费结算 := '现金';
  End;

  ---------------------------------------------------------------------------------
  --循环处理每行费用(收入项目行)
  n_总金额   := 0;
  n_正常退费 := 1;
  For r_Bill In c_Bill Loop
    If Nvl(r_Bill.执行状态, 0) <> 1 Then
      --求剩余数量,剩余应收,剩余实收
      Select Sum(Nvl(付数, 1) * 数次), Sum(应收金额), Sum(实收金额), Sum(统筹金额)
      Into n_剩余数量, n_剩余应收, n_剩余实收, n_剩余统筹
      From 门诊费用记录
      Where NO = No_In And 记录性质 = 1 And 序号 = r_Bill.序号;
    
      If n_剩余数量 = 0 Then
        --情况：未限定行号,原始单据中的该笔已经全部退费(执行状态=0的一种可能)
        n_正常退费 := 0;
      Else
        --准退数量(非药品项目为剩余数量,原始数量)
        If Instr(',4,5,6,7,', r_Bill.收费类别) = 0 Or (r_Bill.收费类别 = '4' And Nvl(r_Bill.跟踪在用, 0) = 0) Then
          --@@@
          --非药品部分(以具体医嘱执行为准进行检查)
          --: 1.存在医嘱发送的,则以医嘱执行为准(但不能包含:检查;检验;手术;麻醉及输血)
          --: 2.不存在医嘱的,则以剩余数量为准
          n_Count := 0;
          If Instr(',C,D,F,G,K,', ',' || r_Bill.诊疗类别 || ',') = 0 And r_Bill.诊疗类别 Is Not Null Then
            If n_医属执行计价 = 1 Then
              Select Decode(Sign(Sum(数量)), -1, 0, Sum(数量)), Count(*)
              Into n_准退数量, n_Count
              From (Select Max(Decode(a.记录状态, 2, 0, a.Id)) As ID, Max(a.医嘱序号) As 医嘱id, Max(a.收费细目id) As 收费细目id,
                            Sum(Nvl(a.付数, 1) * Nvl(a.数次, 1)) As 数量,
                            Sum(Decode(a.记录状态, 2, 0, Nvl(a.付数, 1) * Nvl(a.数次, 1))) As 原始数量
                     From 门诊费用记录 A, 病人医嘱记录 M
                     Where a.医嘱序号 = m.Id And Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And
                           Instr('5,6,7', a.收费类别) = 0 And a.No = No_In And a.序号 = r_Bill.序号 And a.记录性质 = 1 And
                           a.记录状态 In (1, 2, 3) And a.价格父号 Is Null
                     Group By a.序号
                     Union All
                     Select a.Id, a.医嘱序号 As 医嘱id, a.收费细目id, -1 * b.数量 As 已执行, 0 原始数量
                     From 门诊费用记录 A, 医嘱执行计价 B, 病人医嘱记录 M
                     Where a.医嘱序号 = b.医嘱id And a.收费细目id = b.收费细目id + 0 And a.医嘱序号 = m.Id And
                           Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And Instr('5,6,7', a.收费类别) = 0 And
                           (Exists
                            (Select 1
                             From 病人医嘱执行
                             Where b.医嘱id = 医嘱id And b.发送号 = 发送号 And b.要求时间 = 要求时间 And Nvl(执行结果, 0) = 1) Or Exists
                            (Select 1
                             From 病人医嘱发送
                             Where b.医嘱id = 医嘱id And b.发送号 = 发送号 And Nvl(执行状态, 0) = 1)) And Not Exists
                      (Select 1
                            From 病人医嘱附费
                            Where a.医嘱序号 = 医嘱id And a.No = NO And Mod(a.记录性质, 10) = 记录性质) And a.No = No_In And
                           a.序号 = r_Bill.序号 And a.记录性质 = 1 And a.记录状态 In (1, 3) 　and a.价格父号 Is Null) Q1
              Where Not Exists (Select 1 From 药品收发记录 Where 费用id = Q1.Id) Having Max(ID) <> 0;
            Else
              Select Nvl(Sum(数量), 0), Count(*)
              Into n_准退数量, n_Count
              From (Select a.医嘱id, a.收费细目id, Nvl(a.数量, 1) * Nvl(b.发送数次, 1) As 数量
                     From 病人医嘱计价 A, 病人医嘱发送 B, 门诊费用记录 J, 病人医嘱记录 M
                     Where a.医嘱id = b.医嘱id And a.医嘱id = m.Id And Nvl(b.执行状态, 0) <> 1 And a.医嘱id = j.医嘱序号 And
                           a.收费细目id = j.收费细目id And j.No = No_In And j.记录性质 = 1 And j.序号 = r_Bill.序号 And j.记录状态 In (1, 3) And
                           j.价格父号 Is Null And Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And Exists
                      (Select 1
                            From 病人医嘱计价 A
                            Where a.医嘱id = j.医嘱序号 And a.收费细目id = j.收费细目id And Nvl(a.收费方式, 0) = 0) And Not Exists
                      (Select 1 From 药品收发记录 Where 费用id = j.Id)
                     Union All
                     Select a.医嘱id, a.收费细目id, -1 * Nvl(a.数量, 1) * Nvl(c.本次数次, 1) As 数量
                     From 病人医嘱计价 A, 病人医嘱发送 B, 病人医嘱执行 C, 门诊费用记录 J, 病人医嘱记录 M
                     Where a.医嘱id = b.医嘱id And b.医嘱id = c.医嘱id And b.发送号 = c.发送号 And a.医嘱id = m.Id And Nvl(c.执行结果, 1) = 1 And
                           Nvl(b.执行状态, 0) <> 1 And a.医嘱id = j.医嘱序号 And a.收费细目id = j.收费细目id And j.No = No_In And
                           j.记录性质 = 1 And Nvl(a.收费方式, 0) = 0 And j.序号 = r_Bill.序号 And j.记录状态 In (1, 3) And j.价格父号 Is Null And
                           Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And Not Exists
                      (Select 1 From 药品收发记录 Where 费用id = j.Id) And Not Exists
                      (Select 1 From 材料特性 Where 材料id = j.收费细目id And Nvl(跟踪在用, 0) = 1)
                     Union All
                     Select a.医嘱序号 As 医嘱id, a.收费细目id, Nvl(a.付数, 1) * a.数次 As 数量
                     From 门诊费用记录 A, 病人医嘱记录 M
                     Where a.医嘱序号 = m.Id And Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And a.No = No_In And
                           a.记录性质 = 1 And a.序号 = r_Bill.序号 And a.记录状态 = 2 And a.价格父号 Is Null And Not Exists
                      (Select 1 From 药品收发记录 Where NO = No_In And 单据 In (8, 24) And 药品id = a.收费细目id));
            End If;
          End If;
          If Nvl(n_Count, 0) <> 0 And n_准退数量 = 0 Then
            v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已执行,不允许退费！';
            Raise Err_Item;
          End If;
        
          If Nvl(n_Count, 0) = 0 Then
            n_准退数量 := n_剩余数量;
          End If;
        
        Else
          Select Nvl(Sum(Nvl(付数, 1) * 实际数量), 0), Count(*)
          Into n_准退数量, n_Count
          From 药品收发记录
          Where NO = No_In And 单据 In (8, 24) And Mod(记录状态, 3) = 1 --@@@
                And 审核人 Is Null And 费用id = r_Bill.Id;
        
          --有剩余数量无准退数量的有两种情况：
          --1.不跟踪在用的卫材无对应的收发记录,这时使用剩余数量
          --2.并发操作,此时已发药或发料
          If n_准退数量 = 0 Then
            If r_Bill.收费类别 = '4' Then
              If n_Count > 0 Then
                v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已发料,须退料后再退费！';
                Raise Err_Item;
              Else
                n_准退数量 := n_剩余数量;
              End If;
            Else
              v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已发药,须退药后再退费！';
              Raise Err_Item;
            End If;
          End If;
        End If;
      
        --是否部分退费
        If r_Bill.执行状态 = 2 Or n_准退数量 <> Nvl(r_Bill.付数, 1) * r_Bill.数次 Then
          n_正常退费 := 0;
        End If;
      
        --处理门诊费用记录
        n_费用状态 := 0;
        --该笔项目第几次退费
        Select Nvl(Max(Abs(执行状态)), 0) + 1
        Into n_退费次数
        From 门诊费用记录
        Where NO = No_In And 记录性质 = 1 And 记录状态 = 2 And Nvl(执行状态, 0) < 0 And 序号 = r_Bill.序号;
      
        n_应收金额 := Round(n_剩余应收 * (n_准退数量 / n_剩余数量), n_Dec);
        n_实收金额 := Round(n_剩余实收 * (n_准退数量 / n_剩余数量), n_Dec);
        n_统筹金额 := Round(n_剩余统筹 * (n_准退数量 / n_剩余数量), n_Dec);
        n_总金额   := n_总金额 + n_实收金额;
      
        --插入退费记录
        Insert Into 门诊费用记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id, 收费类别, 收费细目id,
           计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人, 执行状态,
           费用状态, 执行时间, 操作员编号, 操作员姓名, 发生时间, 登记时间, 结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 结论, 缴款组id)
          Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别,
                 病人科室id, 收费类别, 收费细目id, 计算单位, Decode(Sign(n_准退数量 - Nvl(付数, 1) * 数次), 0, 付数, 1), 发药窗口,
                 Decode(Sign(n_准退数量 - Nvl(付数, 1) * 数次), 0, -1 * 数次, -1 * n_准退数量), 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价,
                 -1 * n_应收金额, -1 * n_实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人, -1 * n_退费次数, n_费用状态, 执行时间, 操作员编号_In, 操作员姓名_In,
                 发生时间, d_Date, n_结帐id, -1 * n_实收金额, 保险项目否, 保险大类id, -1 * n_统筹金额, 摘要, Decode(Nvl(附加标志, 0), 9, 1, 0), 保险编码,
                 费用类型, 结论, n_组id
          From 门诊费用记录
          Where ID = r_Bill.Id;
      
        --标记原费用记录
        --执行状态:全部退完(准退数=剩余数)标记为0,否则标记为1,异常收费单,还是标明9
        Update 门诊费用记录
        Set 记录状态 = 3, 执行状态 = Decode(Nvl(执行状态, 0), 9, 9, Decode(Sign(n_准退数量 - n_剩余数量), 0, 0, 1))
        Where ID = r_Bill.Id;
      End If;
    Else
      --情况:没限定行号,原始单据中包括已经完全执行的
      n_正常退费 := 0;
    End If;
  End Loop;

  ---------------------------------------------------------------------------------
  --处理病人预交记录
  --自动产生误差费,默认保留一位
  n_总金额 := Round(n_总金额, 1);
  --原单据的结帐ID
  Select 结帐id, 病人id
  Into n_原结帐id, n_病人id
  From 门诊费用记录
  Where NO = No_In And 记录性质 = 1 And 记录状态 In (1, 3) And Rownum = 1;

  If n_正常退费 = 1 Then
    --单据第一次退费且全部退完
    --冲预交部分记录
    Insert Into 病人预交记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交, 结帐id,
       缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质)
      Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, d_Date,
             操作员姓名_In, 操作员编号_In, -1 * 冲预交, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 0, n_结算序号, 3
      From 病人预交记录
      Where 记录性质 In (1, 11) And 结帐id = n_原结帐id And Nvl(冲预交, 0) <> 0;
    --处理病人预交余额
    For v_预交 In (Select NO, 预交类别, Nvl(Sum(Nvl(冲预交, 0)), 0) As 预交金额, 病人id
                 From 病人预交记录
                 Where 记录性质 In (1, 11) And 结帐id = n_原结帐id
                 Group By NO, 预交类别, 病人id
                 Having Sum(Nvl(冲预交, 0)) <> 0) Loop
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) + Nvl(v_预交.预交金额, 0)
      Where 病人id = v_预交.病人id And 性质 = 1 And 类型 = Nvl(v_预交.预交类别, 2)
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 类型, 预交余额, 性质)
        Values
          (v_预交.病人id, Nvl(v_预交.预交类别, 2), Nvl(v_预交.预交金额, 0), 1);
        n_返回值 := n_预交金额;
      End If;
      If n_返回值 = 0 Then
        Delete From 病人余额
        Where 病人id = v_预交.病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
      End If;
    
      --更新预交单据余额
      Select Max(ID) Into n_充值id From 病人预交记录 Where NO = v_预交.No And 记录性质 = 1 And 记录状态 <> 2;
      If Nvl(n_充值id, 0) <> 0 Then
        Update 预交单据余额
        Set 预交余额 = Nvl(预交余额, 0) + Nvl(v_预交.预交金额, 0)
        Where 病人id = v_预交.病人id And 预交id = n_充值id
        Returning 预交余额 Into n_返回值;
        If Sql%RowCount = 0 Then
          Insert Into 预交单据余额
            (预交id, 病人id, 预交类别, 预交余额)
          Values
            (n_充值id, v_预交.病人id, Nvl(v_预交.预交类别, 2), Nvl(v_预交.预交金额, 0));
          n_返回值 := Nvl(v_预交.预交金额, 0);
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 预交单据余额 Where 预交id = n_充值id And Nvl(预交余额, 0) = 0;
        End If;
      End If;
    End Loop;
  
    --原样退回(冲预交在前面已处理)
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别,
       卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质)
      Select 病人预交记录_Id.Nextval, 记录性质, NO, 2, 病人id, 主页id, 摘要, 结算方式, 结算号码, d_Date, 缴款单位, 单位开户行, 单位帐号, 操作员编号_In, 操作员姓名_In,
             -1 * 冲预交, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 0, n_结算序号, 3
      From 病人预交记录 A, (Select 名称 From 结算方式 Where 性质 In (3, 4)) J,
           (Select m.Id As 预交id From 病人预交记录 M Where m.结帐id = n_原结帐id And m.记录性质 = 3 And m.记录状态 = 1) Q
      Where a.记录性质 = 3 And a.记录状态 = 1 And a.结帐id = n_原结帐id And a.Id = q.预交id(+) And a.结算方式 = j.名称(+);
  Else
    --部分退费直接退为指定结算方式
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
       结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质)
      Select 病人预交记录_Id.Nextval, 3, No_In, 2, 病人id, 主页id, '部分退费结算', v_退费结算, d_Date, Null, Null, Null, 操作员编号_In, 操作员姓名_In,
             -1 * n_总金额, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 0, n_结算序号, 3
      
      From 病人预交记录
      Where 记录性质 = 3 And 记录状态 In (1, 3) And 结帐id = n_原结帐id And Rownum = 1;
  
    --如果收费时只使用了预交款,则要退预交,并且可能有多笔冲预交
    If Sql%RowCount = 0 Then
      n_预交金额 := n_总金额;
    
      For r_Deposit In c_Deposit(n_原结帐id) Loop
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
           结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质)
          Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号,
                 d_Date, 操作员姓名_In, 操作员编号_In, Decode(Sign(r_Deposit.金额 - n_预交金额), -1, -1 * r_Deposit.金额, -1 * n_预交金额),
                 n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 0, n_结算序号, 3
          From 病人预交记录
          Where ID = r_Deposit.Id;
      
        --更新病人预交余额
        Update 病人余额
        Set 预交余额 = Nvl(预交余额, 0) + Nvl(r_Deposit.金额, 0)
        Where 病人id = r_Deposit.病人id And 性质 = 1 And 类型 = 1
        Returning 预交余额 Into n_返回值;
        If Sql%RowCount = 0 Then
          Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (r_Deposit.病人id, 1, Nvl(r_Deposit.金额, 0), 1);
          n_返回值 := Nvl(r_Deposit.金额, 0);
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 病人余额
          Where 病人id = r_Deposit.病人id And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
        End If;
        --更新预交单据余额
        Select Max(ID) Into n_充值id From 病人预交记录 Where NO = r_Deposit.No And 记录性质 = 1 And 记录状态 <> 2;
        If Nvl(n_充值id, 0) <> 0 Then
          Update 预交单据余额
          Set 预交余额 = Nvl(预交余额, 0) + Nvl(r_Deposit.金额, 0)
          Where 病人id = r_Deposit.病人id And 预交id = n_充值id
          Returning 预交余额 Into n_返回值;
          If Sql%RowCount = 0 Then
            Insert Into 预交单据余额
              (预交id, 病人id, 预交类别, 预交余额)
            Values
              (n_充值id, r_Deposit.病人id, Nvl(r_Deposit.预交类别, 2), Nvl(r_Deposit.金额, 0));
            n_返回值 := Nvl(r_Deposit.金额, 0);
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 预交单据余额 Where 预交id = n_充值id And Nvl(预交余额, 0) = 0;
          End If;
        End If;
        --检查是否已经处理完
        If r_Deposit.金额 < n_预交金额 Then
          n_预交金额 := n_预交金额 - r_Deposit.金额;
        Else
          n_预交金额 := 0;
        End If;
        If n_预交金额 = 0 Then
          Exit;
        End If;
      End Loop;
    
    End If;
  End If;
  --更新原记录
  Update 病人预交记录 Set 记录状态 = 3 Where 记录性质 = 3 And 记录状态 In (1, 3) And 结帐id = n_原结帐id;

  Select Nvl(Sum(Nvl(结帐金额, 0)), 0) Into n_实收金额 From 门诊费用记录 Where 结帐id = n_结帐id;
  Select Nvl(Sum(Nvl(冲预交, 0)), 0) Into n_返回值 From 病人预交记录 Where 结帐id = n_结帐id;

  n_实收金额 := n_实收金额 - n_返回值;

  If n_实收金额 <> 0 Then
    --未找到，新产生误差项
    Zl_简单收费误差_Insert(No_In, n_病人id, n_结帐id, n_实收金额, d_Date, 操作员编号_In, 操作员姓名_In, 1);
  End If;

  --更新 是否电子票据 标记
  Select Max(a.是否电子票据)
  Into n_是否电子票据
  From 病人预交记录 A
  Where a.结帐id = n_原结帐id And a.记录性质 In (11, 3);

  Update 病人预交记录 Set 是否电子票据 = n_是否电子票据 Where 结帐id = n_结帐id;

  --人员缴款余额(注意是预交记录处理后才处理，包括个人帐户等的结算金额,不含退冲预交款)
  For r_Moneyrow In c_Money(n_结帐id) Loop
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) + r_Moneyrow.冲预交
    Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_Moneyrow.结算方式
    Returning 余额 Into n_返回值;
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额
        (收款员, 结算方式, 性质, 余额)
      Values
        (操作员姓名_In, r_Moneyrow.结算方式, 1, r_Moneyrow.冲预交);
      n_返回值 := r_Moneyrow.冲预交;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 人员缴款余额
      Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_Moneyrow.结算方式 And Nvl(余额, 0) = 0;
    End If;
  End Loop;

  ---------------------------------------------------------------------------------
  --退费票据回收(仅全退时才回退,部分退是在重打过程中回收)
  If Nvl(n_是否电子票据, 0) = 0 Then
    --启用标志||NO;执行科室(条数);收据费目(首页汇总,条数);收费细目(条数)
    v_Para     := Nvl(zl_GetSysParameter('票据分配规则', 1121), '0||0;0;0,0;0');
    n_启用模式 := zl_To_Number(Substr(v_Para, 1, 1));
    If n_启用模式 <> 0 Then
      --收回票据
      Select 使用id
      Bulk Collect
      Into l_使用id
      From (Select Distinct b.使用id From 票据打印明细 B Where b.No = No_In And Nvl(b.票种, 0) = 1);
    
      n_启用模式 := l_使用id.Count;
      If l_使用id.Count <> 0 Then
        --插入回收记录
        Forall I In 1 .. l_使用id.Count
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用人, 使用时间, 票据金额)
            Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, 操作员姓名_In, d_Date, 票据金额
            From 票据使用明细 A
            Where ID = l_使用id(I) And 性质 = 1 And Not Exists
             (Select 1 From 票据使用明细 Where 号码 = a.号码 And 票种 = a.票种 And Nvl(性质, 0) <> 1);
      
        Forall I In 1 .. l_使用id.Count
          Update 票据打印明细 Set 是否回收 = 1 Where 使用id = l_使用id(I) And Nvl(是否回收, 0) = 0;
      
      End If;
    End If;
  
    If n_启用模式 = 0 Then
      --获取单据最后一次的打印ID(可能是多张单据收费打印)
      Begin
        --性质=1，原因=6为退费打印票据(红票)，不回收
        Select ID
        Into n_打印id
        From (Select b.Id
               From 票据使用明细 A, 票据打印内容 B
               Where a.打印id = b.Id And a.性质 = 1 And a.原因 In (1, 3) And b.数据性质 = 1 And b.No = No_In
               Order By a.使用时间 Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          Null;
      End;
      --可能以前没有打印,无收回
      If n_打印id Is Not Null Then
        --a.多张单据循环调用时只能收回一次
        Select Count(*) Into n_Count From 票据使用明细 Where 票种 = 1 And 性质 = 2 And 打印id = n_打印id;
        If n_Count = 0 Then
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
            Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_Date, 操作员姓名_In, 票据金额
            From 票据使用明细
            Where 打印id = n_打印id And 票种 = 1 And 性质 = 1;
        Else
          --b.部分退费多次收回时,最后一次全退收回要排开已收回的
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
            Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_Date, 操作员姓名_In, 票据金额
            From 票据使用明细 A
            Where 打印id = n_打印id And 票种 = 1 And 性质 = 1 And Not Exists
             (Select 1 From 票据使用明细 B Where a.号码 = b.号码 And 打印id = n_打印id And 票种 = 1 And 性质 = 2);
        End If;
      End If;
    End If;
  End If;

  ---------------------------------------------------------------------------------
  --药品卫材相关内容
  --必须按照“收费细目id”升序排序，防止并发锁“药品库存”表
  For r_Expenses In (Select ID
                     From 门诊费用记录
                     Where NO = No_In And 记录性质 = 1 And 记录状态 In (1, 3) And 收费类别 In ('4', '5', '6', '7')
                     Order By 收费细目id) Loop
    Zl_药品收发记录_销售退费(r_Expenses.Id);
  End Loop;

  --医嘱处理
  --删除病人医嘱附费(最后一次删除时)
  For c_医嘱 In (Select Distinct 医嘱序号
               From 门诊费用记录
               Where NO = No_In And 记录性质 = 1 And 记录状态 = 3 And 医嘱序号 Is Not Null) Loop
    Select Nvl(Count(*), 0)
    Into n_Count
    From (Select 序号, Sum(数量) As 剩余数量
           From (Select 记录状态, 执行状态, Nvl(价格父号, 序号) As 序号, Avg(Nvl(付数, 1) * 数次) As 数量
                  From 门诊费用记录
                  Where 记录性质 = 1 And Nvl(附加标志, 0) <> 9 And 医嘱序号 + 0 = c_医嘱.医嘱序号 And NO = No_In
                  Group By 记录状态, 执行状态, Nvl(价格父号, 序号))
           Group By 序号
           Having Sum(数量) <> 0);
  
    If n_Count = 0 Then
      Delete From 病人医嘱附费 Where 医嘱id = c_医嘱.医嘱序号 And 记录性质 = 1 And NO = No_In;
    End If;
  End Loop;

  --场合_In    Integer:=0, --0:门诊;1-住院
  --性质_In    Integer:=1, --1-收费单;2-记帐单
  --操作_In    Integer:=0, --0:删除划价单;1-收费或记帐;2-退费或销帐
  --No_In      门诊费用记录.No%Type,
  --医嘱ids_In varchar2 := Null
  Zl_医嘱发送_计费状态_Update(0, 1, 2, No_In);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊简单收费_Delete;
/
Create Or Replace Procedure Zl_费用补充结算_Modify
(
  操作类型_In     Number,
  结算id_In       In 费用补充记录.结算id%Type,
  结算方式_In     Varchar2,
  卡类别id_In     病人预交记录.卡类别id%Type := Null,
  卡号_In         病人预交记录.卡号%Type := Null,
  交易流水号_In   病人预交记录.交易流水号%Type := Null,
  交易说明_In     病人预交记录.交易说明%Type := Null,
  误差金额_In     门诊费用记录.实收金额%Type := Null,
  完成结算_In     Number := 0,
  冲预交_In       病人预交记录.冲预交%Type := Null,
  校对标志_In     病人预交记录.校对标志%Type := 0,
  是否电子票据_In 病人预交记录.是否电子票据%Type := Null
) As
  ------------------------------------------------------------------------------------------------------------------------------ 
  --功能:保险补充结算时,修改结算的相关信息 
  --操作类型_In: 
  --   0-普通结算方式: 
  --     结算方式_IN:允许传入多个,格式为:结算方式|结算金额|结算号码|结算摘要||.. ;也允许传入空. 
  --   1.三方卡结算: 
  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:结算方式|结算金额|结算号码|结算摘要 
  --     ②卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入 
  --   2-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新) 
  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.. 
  -- 误差金额_In:存在误差费时,传入 
  -- 完成结算_In:1-完成补充结算;0-未完成补充结算;2-完成了异常作废 
  -- 冲预交_In:冲预交金额，退款时为负，收款时为正 
  -- 校对标志_In  操作类型_In为1时有效 
  --是否电子票据_In:null-表示过程内部直接判断，非空表示直接以传入的为准
  ------------------------------------------------------------------------------------------------------------------------------ 
  v_误差费   结算方式.名称%Type;
  n_Count    Number(18);
  n_会话号   病人预交记录.会话号%Type; --格式：SID+'_'+SERIAL# 
  v_交易人员 病人预交记录.交易人员%Type;

  v_结算内容 Varchar2(4000);
  v_当前结算 Varchar2(4000);
  v_结算方式 病人预交记录.结算方式%Type;
  n_结算金额 病人预交记录.冲预交%Type;
  v_结算号码 病人预交记录.结算号码%Type;
  v_结算摘要 病人预交记录.摘要%Type;

  n_冲预交 病人预交记录.冲预交%Type;
  n_返回值 病人预交记录.冲预交%Type;
  n_预交id 病人预交记录.Id%Type;
  l_预交id t_NumList := t_NumList();

  n_是否电子票据 Number(2);
  n_险类         保险结算记录.险类%Type;

  Cursor c_Balance Is
    Select 记录性质, NO, 记录状态, 实际票号, 结算id, 收费结帐id, 费用状态, 操作员编号, 操作员姓名, 登记时间, 缴款组id, 病人id, 结算序号, 附加标志
    From 费用补充记录 A
    Where 结算id = 结算id_In And 记录性质 = 1 And Rownum < 2;
  r_Balance c_Balance%RowType;

  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin
  Begin
    Select Sid || '_' || Serial# Into n_会话号 From V$session Where Audsid = Userenv('sessionid');
  Exception
    When Others Then
      n_会话号 := Null;
  End;
  v_交易人员 := zl_UserName;

  Select Count(1) Into n_Count From 费用补充记录 Where 结算id = 结算id_In And Rownum < 2 And 记录性质 = 1;
  If n_Count = 0 Then
    v_Err_Msg := '未找到医保补结算数据，不能继续操作!';
    Raise Err_Item;
  End If;

  Open c_Balance;
  Fetch c_Balance
    Into r_Balance;

  If Nvl(误差金额_In, 0) <> 0 Then
    Begin
      Select 名称 Into v_误差费 From 结算方式 Where Nvl(性质, 0) = 9;
    Exception
      When Others Then
        v_误差费 := '误差费';
    End;
  
    Update 病人预交记录
    Set 冲预交 = Nvl(冲预交, 0) + Nvl(误差金额_In, 0)
    Where 结帐id = 结算id_In And 结算方式 = v_误差费;
    If Sql%NotFound Then
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志, 卡类别id,
         结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号)
      Values
        (病人预交记录_Id.Nextval, 6, r_Balance.No, r_Balance.记录状态, r_Balance.病人id, Null, Null, v_误差费, r_Balance.登记时间,
         r_Balance.操作员编号, r_Balance.操作员姓名, 误差金额_In, r_Balance.结算id, r_Balance.缴款组id, Sysdate, v_交易人员, r_Balance.结算序号, 2,
         Null, Null, 卡号_In, 交易流水号_In, 交易说明_In, Null, 6, n_会话号);
      Update 病人预交记录 Set 冲预交 = Nvl(冲预交, 0) - 误差金额_In Where 结帐id = 结算id_In And 结算方式 Is Null;
    End If;
  End If;

  --退预交款 
  If Nvl(冲预交_In, 0) < 0 Then
    n_冲预交 := -1 * 冲预交_In;
    For v_退预交 In (Select Max(ID) As 预交id, NO, 病人id, Max(收款时间) As 收款时间, Sum(Nvl(冲预交, 0)) As 金额
                  From 病人预交记录
                  Where Mod(记录性质, 10) = 1 And Nvl(预交类别, 0) = 1 And
                        结帐id In
                        (Select a.结帐id
                         From 门诊费用记录 A, 门诊费用记录 B
                         Where a.记录性质 = b.记录性质 And a.No = b.No And a.序号 = b.序号 And b.记录状态 <> 2 And
                               b.结帐id In
                               (Select 收费结帐id From 费用补充记录 Where 记录性质 = 1 And 结算序号 = r_Balance.结算序号))
                  Group By NO, 病人id
                  Having Sum(Nvl(冲预交, 0)) > 0
                  Order By 收款时间 Desc) Loop
    
      If v_退预交.金额 - n_冲预交 < 0 Then
        n_结算金额 := -1 * v_退预交.金额;
        n_冲预交   := n_冲预交 - v_退预交.金额;
      Else
        n_结算金额 := -1 * n_冲预交;
        n_冲预交   := 0;
      End If;
    
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 实际票号, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志,
         卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 预交类别, 结算性质, 会话号)
        Select 病人预交记录_Id.Nextval, 11, NO, 实际票号, 记录状态, 病人id, 主页id, 摘要, 结算方式, r_Balance.登记时间, r_Balance.操作员编号,
               r_Balance.操作员姓名, n_结算金额, r_Balance.结算id, r_Balance.缴款组id, Sysdate, v_交易人员, r_Balance.结算序号, 2, 卡类别id,
               结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 预交类别, 6, n_会话号
        From 病人预交记录
        Where ID = v_退预交.预交id;
    
      --更新预交单据余额 
      Select Max(ID) Into n_预交id From 病人预交记录 Where NO = v_退预交.No And 记录性质 = 1 And 记录状态 <> 2;
      If Nvl(n_预交id, 0) <> 0 Then
        Update 预交单据余额
        Set 预交余额 = Nvl(预交余额, 0) + (-1 * n_结算金额)
        Where 病人id = v_退预交.病人id And 预交id = n_预交id
        Returning 预交余额 Into n_返回值;
        If Sql%RowCount = 0 Then
          Insert Into 预交单据余额
            (预交id, 病人id, 预交类别, 预交余额)
          Values
            (n_预交id, v_退预交.病人id, 1, -1 * n_结算金额);
          n_返回值 := -1 * n_结算金额;
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 预交单据余额 Where 预交id = n_预交id And Nvl(预交余额, 0) = 0;
        End If;
      End If;
    
      --更新病人余额 
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) + (-1 * n_结算金额)
      Where 病人id = v_退预交.病人id And 性质 = 1 And 类型 = 1
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (v_退预交.病人id, 1, -1 * n_结算金额, 1);
        n_返回值 := -1 * n_结算金额;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 病人余额
        Where 病人id = v_退预交.病人id And 性质 = 1 And 类型 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
      End If;
    
      Update 病人预交记录 Set 冲预交 = Nvl(冲预交, 0) - n_结算金额 Where 结帐id = 结算id_In And 结算方式 Is Null;
    
      If n_冲预交 = 0 Then
        Exit;
      End If;
    End Loop;
    If n_冲预交 <> 0 Then
      v_Err_Msg := '当前退款金额大于了预交款可退金额，退款失败！';
      Raise Err_Item;
    End If;
  End If;

  --0.普通结算方式 
  If Nvl(操作类型_In, 0) = 0 Then
    --各个收费结算 :格式为:结算方式|结算金额|结算号码|结算摘要||.. 
    v_结算内容 := 结算方式_In || '||';
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      v_结算号码 := LTrim(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_结算摘要 := LTrim(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
    
      If Nvl(n_结算金额, 0) <> 0 Then
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志,
           卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号)
        Values
          (病人预交记录_Id.Nextval, 6, r_Balance.No, r_Balance.记录状态, r_Balance.病人id, Null, v_结算摘要, v_结算方式, r_Balance.登记时间,
           r_Balance.操作员编号, r_Balance.操作员姓名, n_结算金额, r_Balance.结算id, r_Balance.缴款组id, Sysdate, v_交易人员, r_Balance.结算序号, 2,
           Null, Null, 卡号_In, 交易流水号_In, 交易说明_In, v_结算号码, 6, n_会话号);
      
        Update 病人预交记录 Set 冲预交 = Nvl(冲预交, 0) - n_结算金额 Where 结帐id = 结算id_In And 结算方式 Is Null;
      End If;
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  --1.三方卡结算 
  If Nvl(操作类型_In, 0) = 1 Then
    v_当前结算 := 结算方式_In;
    v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
    v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
    n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
    v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
    v_结算号码 := LTrim(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
    v_结算摘要 := LTrim(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
  
    If Nvl(n_结算金额, 0) <> 0 Then
      Update 病人预交记录
      Set 冲预交 = Nvl(冲预交, 0) + n_结算金额
      Where 结帐id = 结算id_In And 结算方式 = v_结算方式 And 卡类别id = 卡类别id_In
      Returning ID Into n_预交id;
      If Sql%NotFound Then
        Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志,
           卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号)
        Values
          (n_预交id, 6, r_Balance.No, r_Balance.记录状态, r_Balance.病人id, Null, v_结算摘要, v_结算方式, r_Balance.登记时间,
           r_Balance.操作员编号, r_Balance.操作员姓名, n_结算金额, r_Balance.结算id, r_Balance.缴款组id, Sysdate, v_交易人员, r_Balance.结算序号,
           校对标志_In, 卡类别id_In, Null, 卡号_In, 交易流水号_In, 交易说明_In, v_结算号码, 6, n_会话号);
      End If;
      Update 病人预交记录 Set 冲预交 = Nvl(冲预交, 0) - n_结算金额 Where 结帐id = 结算id_In And 结算方式 Is Null;
    
      --调用其他结算信息更新
      Zl_Custom_Balance_Update(n_预交id);
    End If;
  End If;

  --2.医保结算 
  If Nvl(操作类型_In, 0) = 2 Then
    --2.1检查是否已经存在医保结算数据,存在先删除 
    n_结算金额 := 0;
    For v_医保 In (Select ID, 结算方式, 冲预交
                 From 病人预交记录 A
                 Where 结帐id = 结算id_In And Exists
                  (Select 1 From 结算方式 Where a.结算方式 = 名称 And 性质 In (3, 4)) And Mod(记录性质, 10) <> 1) Loop
      n_结算金额 := n_结算金额 + Nvl(v_医保.冲预交, 0);
      l_预交id.Extend;
      l_预交id(l_预交id.Count) := v_医保.Id;
    End Loop;
  
    If l_预交id.Count <> 0 Then
      Forall I In 1 .. l_预交id.Count
        Delete 病人预交记录 Where ID = l_预交id(I);
    End If;
  
    --先删除结算方式为空的记录 
    Delete 病人预交记录 Where 结帐id = 结算id_In And 结算方式 Is Null;
    If 结算方式_In Is Not Null Then
      v_结算内容 := 结算方式_In || '||';
    End If;
    n_冲预交 := 0;
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      n_结算金额 := To_Number(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志, 卡类别id,
         结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号)
      Values
        (病人预交记录_Id.Nextval, 6, r_Balance.No, r_Balance.记录状态, r_Balance.病人id, Null, '保险结算', v_结算方式, r_Balance.登记时间,
         r_Balance.操作员编号, r_Balance.操作员姓名, Nvl(n_结算金额, 0), r_Balance.结算id, r_Balance.缴款组id, Sysdate, v_交易人员,
         r_Balance.结算序号, 2, Null, Null, Null, Null, Null, Null, 6, n_会话号);
      n_冲预交 := n_冲预交 + Nvl(n_结算金额, 0);
    
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  
    --处理结算方式为NULL 
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志, 卡类别id,
       结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号)
    Values
      (病人预交记录_Id.Nextval, 6, r_Balance.No, r_Balance.记录状态, r_Balance.病人id, Null, '', Null, r_Balance.登记时间,
       r_Balance.操作员编号, r_Balance.操作员姓名, -1 * n_冲预交, r_Balance.结算id, r_Balance.缴款组id, Sysdate, v_交易人员, r_Balance.结算序号, 1,
       Null, Null, Null, Null, Null, Null, 6, n_会话号);
    --医保相关表的处理 
    Update 保险结算明细 Set 标志 = 2 Where 结帐id = 结算id_In;
  End If;

  If Nvl(完成结算_In, 0) = 0 Then
    Return;
  End If;

  If Nvl(完成结算_In, 0) = 2 Then
    --1.更新校对标志 
    Update 病人预交记录 Set 校对标志 = 0 Where NO = r_Balance.No;
    Update 费用补充记录 Set 费用状态 = 2 Where NO = r_Balance.No;
    If Sql%NotFound Then
      v_Err_Msg := '未找到医保补结算数据，可能被他人进行了作废操作!';
      Raise Err_Item;
    End If;
    Return;
  End If;

  Delete 病人预交记录 Where 结帐id = 结算id_In And Mod(记录性质, 10) <> 1 And 结算方式 Is Null And Nvl(冲预交, 0) = 0;
  If Sql%NotFound Then
    Select Count(1) Into n_Count From 病人预交记录 A Where 结帐id = 结算id_In And 结算方式 Is Null;
    If n_Count <> 0 Then
      v_Err_Msg := '还存在未缴款的数据，不能完成结算！';
    Else
      v_Err_Msg := '结算信息错误，可能因为并发原因造成结算信息错误，请在[保险补充结算]中重新结算！';
    End If;
    Raise Err_Item;
  End If;

  --1.更新异常状态 
  Update 费用补充记录 Set 费用状态 = 0 Where 结算序号 = r_Balance.结算序号;
  If Sql%NotFound Then
    v_Err_Msg := '未找到医保补结算数据，可能被他人进行了退费或作废操作!';
    Raise Err_Item;
  End If;

  n_是否电子票据 := 是否电子票据_In;
  If 是否电子票据_In Is Null Then
    Select Nvl(Max(险类), 0) Into n_险类 From 保险结算记录 Where 记录id = 结算id_In And 性质 = 1;
    If Nvl(r_Balance.附加标志, 0) = 1 Then
      n_是否电子票据 := Zl_Fun_Isstarteinvoice(4, n_险类);
    Else
      n_是否电子票据 := Zl_Fun_Isstarteinvoice(1, n_险类);
    End If;
  End If;

  --2.更新校对标志,会话号 
  Update 病人预交记录 Set 校对标志 = 0, 会话号 = Null, 是否电子票据 = n_是否电子票据 Where 结帐id = 结算id_In;

  --3.更新人员缴款数据 
  For c_缴款 In (Select a.结算方式, a.操作员姓名, Nvl(Sum(a.冲预交), 0) As 冲预交
               From 病人预交记录 A
               Where a.结算序号 = r_Balance.结算序号 And Mod(a.记录性质, 10) <> 1
               Group By a.结算方式, a.操作员姓名
               Having Nvl(Sum(a.冲预交), 0) <> 0) Loop
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) + c_缴款.冲预交
    Where 收款员 = c_缴款.操作员姓名 And 性质 = 1 And 结算方式 = c_缴款.结算方式;
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额
        (收款员, 结算方式, 性质, 余额)
      Values
        (c_缴款.操作员姓名, c_缴款.结算方式, 1, c_缴款.冲预交);
    End If;
  End Loop;

  --消息集成处理 
  --结算类型:1-收费结算，2-补充结算 
  --结帐ID:结算id 
  b_Message.Zlhis_Charge_002(2, 结算id_In);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_费用补充结算_Modify;
/
Create Or Replace Procedure Zl_费用补充结算_完成退费
(
  结算id_In     In 费用补充记录.结算id%Type,
  结算方式_In   Varchar2,
  卡类别id_In   病人预交记录.卡类别id%Type := Null,
  卡号_In       病人预交记录.卡号%Type := Null,
  交易流水号_In 病人预交记录.交易流水号%Type := Null,
  交易说明_In   病人预交记录.交易说明%Type := Null,
  误差金额_In   门诊费用记录.实收金额%Type := Null,
  完成结算_In   Number := 1,
  冲预交_In     病人预交记录.冲预交%Type := Null,
  校对标志_In   病人预交记录.校对标志%Type := 0
) As
  ------------------------------------------------------------------------------------------------------------------------------
  --功能:补结算退费
  --   结算方式_IN:格式为:"结算方式|结算金额|结算号码|结算摘要" ;也允许传入空；结算金额，退款时为负，收款时为正
  --   三方卡结算需传入卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In
  --   误差金额_In 存在误差费时,传入
  --   完成结算_In  1-完成补充结算退费;0-未完成补充结算退费
  --   冲预交_In 冲预交金额，退款时为负，收款时为正
  --   校对标志_In  三方卡退费时有效
  ------------------------------------------------------------------------------------------------------------------------------
  v_误差费   结算方式.名称%Type;
  n_Count    Number(18);
  n_会话号   病人预交记录.会话号%Type; --格式：SID+'_'+SERIAL#
  v_交易人员 病人预交记录.交易人员%Type;

  v_当前结算 Varchar2(4000);
  v_结算内容 Varchar2(4000);
  v_结算方式 病人预交记录.结算方式%Type;
  n_结算金额 病人预交记录.冲预交%Type;
  v_结算号码 病人预交记录.结算号码%Type;
  v_结算摘要 病人预交记录.摘要%Type;

  n_预交金额 病人预交记录.冲预交%Type;
  n_冲预交   病人预交记录.冲预交%Type;
  n_剩余金额 病人预交记录.冲预交%Type;
  n_返回值   病人预交记录.冲预交%Type;
  n_预交id   病人预交记录.Id%Type;
  n_Rowcount Number;
  n_Currrow  Number;

  n_是否电子票据 病人预交记录.是否电子票据%Type;

  Cursor c_Balance Is
    Select 记录性质, NO, 记录状态, 实际票号, 结算id, 收费结帐id, 费用状态, 操作员编号, 操作员姓名, 登记时间, 缴款组id, 病人id, 结算序号, 附加标志
    From 费用补充记录 A
    Where 结算id = 结算id_In And 记录性质 = 1 And Rownum < 2;
  r_Balance c_Balance%RowType;

  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin
  Begin
    Select Sid || '_' || Serial# Into n_会话号 From V$session Where Audsid = Userenv('sessionid');
  Exception
    When Others Then
      n_会话号 := Null;
  End;
  v_交易人员 := zl_UserName;

  Select Count(1) Into n_Count From 费用补充记录 Where 结算id = 结算id_In And 记录性质 = 1 And Rownum < 2;
  If n_Count = 0 Then
    v_Err_Msg := '未找到医保补结算数据，不能继续操作!';
    Raise Err_Item;
  End If;

  Open c_Balance;
  Fetch c_Balance
    Into r_Balance;

  If Nvl(误差金额_In, 0) <> 0 Then
    Begin
      Select 名称 Into v_误差费 From 结算方式 Where Nvl(性质, 0) = 9;
    Exception
      When Others Then
        v_误差费 := '误差费';
    End;
  
    Update 病人预交记录
    Set 冲预交 = Nvl(冲预交, 0) + Nvl(误差金额_In, 0)
    Where 结帐id = 结算id_In And 结算方式 = v_误差费;
    If Sql%NotFound Then
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志, 卡类别id,
         结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号)
      Values
        (病人预交记录_Id.Nextval, 6, r_Balance.No, r_Balance.记录状态, r_Balance.病人id, Null, Null, v_误差费, r_Balance.登记时间,
         r_Balance.操作员编号, r_Balance.操作员姓名, 误差金额_In, r_Balance.结算id, r_Balance.缴款组id, Sysdate, v_交易人员, r_Balance.结算序号, 2,
         Null, Null, 卡号_In, 交易流水号_In, 交易说明_In, Null, 6, n_会话号);
    End If;
  
    Update 病人预交记录
    Set 冲预交 = Nvl(冲预交, 0) - Nvl(误差金额_In, 0)
    Where 结算方式 Is Null And 结帐id = 结算id_In;
    If Sql%NotFound Then
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志, 卡类别id,
         结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号)
      Values
        (病人预交记录_Id.Nextval, 6, r_Balance.No, r_Balance.记录状态, r_Balance.病人id, Null, Null, Null, r_Balance.登记时间,
         r_Balance.操作员编号, r_Balance.操作员姓名, -1 * 误差金额_In, r_Balance.结算id, r_Balance.缴款组id, Sysdate, v_交易人员,
         r_Balance.结算序号, 2, Null, Null, Null, Null, Null, Null, 6, n_会话号);
    End If;
  End If;

  --门诊费用转住院时，退款金额可能大于剩余未退金额
  --计算结算方式为NULL的记录数，退款金额大于未退金额时多余金额全部加到最后一条记录上
  Select Count(1) Into n_Rowcount From 病人预交记录 Where 结算序号 = r_Balance.结算序号 And 结算方式 Is Null;

  --退预交款
  If Nvl(冲预交_In, 0) < 0 Then
    n_预交金额 := -1 * 冲预交_In;
    For v_退预交 In (Select Max(ID) As 预交id, NO, 病人id, Max(收款时间) As 收款时间, Sum(Nvl(冲预交, 0)) As 金额
                  From 病人预交记录
                  Where Mod(记录性质, 10) = 1 And Nvl(预交类别, 0) = 1 And
                        结帐id In
                        (
                         --费用结帐ID
                         Select a.结帐id As 原结帐id
                         From 门诊费用记录 A, 门诊费用记录 B
                         Where a.记录性质 = b.记录性质 And a.No = b.No And a.序号 = b.序号 And b.记录状态 <> 2 And
                               b.结帐id In
                               (Select 收费结帐id From 费用补充记录 Where 记录性质 = 1 And 结算序号 = r_Balance.结算序号)
                         Union All
                         --补充结算结帐ID
                         Select 结帐id
                         From 病人预交记录
                         Where 结算序号 In (Select a.结算序号
                                        From 费用补充记录 A, 费用补充记录 B
                                        Where a.No = b.No And a.记录性质 = b.记录性质 And a.附加标志 = b.附加标志 And b.记录性质 = 1 And
                                              b.结算序号 = r_Balance.结算序号))
                  Group By NO, 病人id
                  Having Sum(Nvl(冲预交, 0)) > 0
                  Order By 收款时间 Desc) Loop
    
      If v_退预交.金额 - n_预交金额 < 0 Then
        n_结算金额 := -1 * v_退预交.金额;
        n_预交金额 := n_预交金额 - v_退预交.金额;
      Else
        n_结算金额 := -1 * n_预交金额;
        n_预交金额 := 0;
      End If;
    
      n_剩余金额 := Nvl(n_结算金额, 0);
      n_Currrow  := 0;
      --需要根据“结算金额”排序，先处理收款（正）的，再处理退款（负）的
      For c_结算 In (Select 病人id, 记录性质, 结帐id, Nvl(Sum(冲预交), 0) As 结算金额
                   From 病人预交记录
                   Where 结算序号 = r_Balance.结算序号 And 结算方式 Is Null
                   Group By 病人id, 结帐id, 记录性质
                   Order By 结算金额 Desc) Loop
      
        n_Currrow := n_Currrow + 1;
        If c_结算.结算金额 < n_剩余金额 Then
          n_冲预交   := n_剩余金额;
          n_剩余金额 := 0;
        Else
          n_冲预交   := c_结算.结算金额;
          n_剩余金额 := n_剩余金额 - c_结算.结算金额;
        End If;
      
        --退款金额大于未退金额时多余金额全部加到最后一条记录上
        If n_Currrow = n_Rowcount And n_剩余金额 <> 0 Then
          n_冲预交   := n_冲预交 + n_剩余金额;
          n_剩余金额 := 0;
        End If;
      
        If n_冲预交 <> 0 Then
          Insert Into 病人预交记录
            (ID, 记录性质, NO, 实际票号, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号,
             校对标志, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 预交类别, 结算性质, 会话号)
            Select 病人预交记录_Id.Nextval, 11, NO, 实际票号, 记录状态, 病人id, 主页id, 摘要, 结算方式, r_Balance.登记时间, r_Balance.操作员编号,
                   r_Balance.操作员姓名, n_冲预交, c_结算.结帐id, r_Balance.缴款组id, Sysdate, v_交易人员, r_Balance.结算序号, 2, 卡类别id, 结算卡序号,
                   卡号, 交易流水号, 交易说明, 结算号码, 预交类别, Mod(c_结算.记录性质, 10), n_会话号
            From 病人预交记录
            Where ID = v_退预交.预交id;
        
          --更新预交单据余额
          Select Max(ID) Into n_预交id From 病人预交记录 Where NO = v_退预交.No And 记录性质 = 1 And 记录状态 <> 2;
          If Nvl(n_预交id, 0) <> 0 Then
            Update 预交单据余额
            Set 预交余额 = Nvl(预交余额, 0) + (-1 * n_冲预交)
            Where 病人id = v_退预交.病人id And 预交id = n_预交id
            Returning 预交余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 预交单据余额
                (预交id, 病人id, 预交类别, 预交余额)
              Values
                (n_预交id, v_退预交.病人id, 1, -1 * n_冲预交);
              n_返回值 := -1 * n_冲预交;
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 预交单据余额 Where 预交id = n_预交id And Nvl(预交余额, 0) = 0;
            End If;
          End If;
        
          --更新病人余额
          Update 病人余额
          Set 预交余额 = Nvl(预交余额, 0) + (-1 * n_冲预交)
          Where 病人id = v_退预交.病人id And 性质 = 1 And 类型 = 1
          Returning 预交余额 Into n_返回值;
          If Sql%RowCount = 0 Then
            Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (c_结算.病人id, 1, -1 * n_冲预交, 1);
            n_返回值 := -1 * n_冲预交;
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 病人余额
            Where 病人id = c_结算.病人id And 性质 = 1 And 类型 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
          End If;
        
          Update 病人预交记录 Set 冲预交 = Nvl(冲预交, 0) - n_冲预交 Where 结算方式 Is Null And 结帐id = c_结算.结帐id;
        End If;
      
        If n_剩余金额 = 0 Then
          Exit;
        End If;
      End Loop;
      If n_剩余金额 <> 0 Then
        v_Err_Msg := '当前退款金额大于了剩余未退金额，退款失败！';
        Raise Err_Item;
      End If;
    
      If n_预交金额 = 0 Then
        Exit;
      End If;
    End Loop;
    If n_预交金额 <> 0 Then
      v_Err_Msg := '当前退款金额大于了预交款可退金额，退款失败！';
      Raise Err_Item;
    End If;
  End If;

  --各个收费结算 :格式为:"结算方式|结算金额|结算号码|结算摘要||.."
  v_结算内容 := 结算方式_In || '||';
  While v_结算内容 Is Not Null Loop
    v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
    v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
    v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
    n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
    v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
    v_结算号码 := LTrim(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
    v_结算摘要 := LTrim(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
  
    n_剩余金额 := Nvl(n_结算金额, 0);
    n_Currrow  := 0;
    -- n_结算金额 为负，需要根据“结算金额”排序，先处理退款（负）的，再处理收款（正）的
    --这样，才能使所有结算方式为空的记录的冲预交金额为零
    For c_结算 In (Select 结帐id, 记录性质, 记录状态, NO, Nvl(Sum(冲预交), 0) As 结算金额
                 From 病人预交记录
                 Where 结算序号 = r_Balance.结算序号 And 结算方式 Is Null
                 Group By 结帐id, 记录性质, NO, 记录状态
                 Order By -1 * 结算金额) Loop
    
      n_Currrow := n_Currrow + 1;
      If c_结算.结算金额 < n_剩余金额 Then
        n_冲预交   := n_剩余金额;
        n_剩余金额 := 0;
      Else
        n_冲预交   := c_结算.结算金额;
        n_剩余金额 := n_剩余金额 - c_结算.结算金额;
      End If;
    
      --退款金额大于未退金额时多余金额全部加到最后一条记录上
      If n_Currrow = n_Rowcount And n_剩余金额 <> 0 Then
        n_冲预交   := n_冲预交 + n_剩余金额;
        n_剩余金额 := 0;
      End If;
    
      If n_冲预交 <> 0 Then
        Update 病人预交记录
        Set 冲预交 = Nvl(冲预交, 0) + n_冲预交
        Where 结帐id = c_结算.结帐id And 结算方式 = v_结算方式 And 卡类别id = 卡类别id_In
        Returning ID Into n_预交id;
        If Sql%NotFound Then
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易时间, 交易人员, 结算序号, 校对标志,
             卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 结算性质, 会话号)
          Values
            (n_预交id, c_结算.记录性质, c_结算.No, c_结算.记录状态, r_Balance.病人id, Null, v_结算摘要, v_结算方式, r_Balance.登记时间,
             r_Balance.操作员编号, r_Balance.操作员姓名, n_冲预交, c_结算.结帐id, r_Balance.缴款组id, Sysdate, v_交易人员, r_Balance.结算序号,
             Decode(Nvl(卡类别id_In, 0), 0, 2, 校对标志_In), 卡类别id_In, Null, 卡号_In, 交易流水号_In, 交易说明_In, v_结算号码,
             Mod(c_结算.记录性质, 10), n_会话号);
        End If;
      
        Update 病人预交记录 Set 冲预交 = Nvl(冲预交, 0) - n_冲预交 Where 结算方式 Is Null And 结帐id = c_结算.结帐id;
      
        If Nvl(卡类别id_In, 0) <> 0 Then
          --调用其他结算信息更新
          Zl_Custom_Balance_Update(n_预交id);
        End If;
      End If;
    
      If n_剩余金额 = 0 Then
        Exit;
      End If;
    End Loop;
    If n_剩余金额 <> 0 Then
      v_Err_Msg := '当前退款金额大于了剩余未退金额，退款失败！';
      Raise Err_Item;
    End If;
  
    v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
  End Loop;

  If Nvl(完成结算_In, 0) = 0 Then
    Return;
  End If;

  Delete From 病人预交记录
  Where 结算序号 = r_Balance.结算序号 And 结算方式 Is Null And Mod(记录性质, 10) <> 1 And Nvl(冲预交, 0) = 0;
  If Sql%NotFound Then
    Select Count(1) Into n_Count From 病人预交记录 A Where 结帐id = 结算id_In And 结算方式 Is Null;
    If n_Count <> 0 Then
      v_Err_Msg := '还存在未缴款的数据，不能完成结算！';
    Else
      v_Err_Msg := '结算信息错误，可能因为并发原因造成结算信息错误，请在[保险补充结算]中重新结算！';
    End If;
    Raise Err_Item;
  End If;

  --一次结算有多条结算方式为NULL的记录
  Select Count(1) Into n_Count From 病人预交记录 A Where 结算序号 = r_Balance.结算序号 And 结算方式 Is Null;
  If n_Count <> 0 Then
    v_Err_Msg := '结算信息错误，请在[保险补充结算]中重新结算！';
    Raise Err_Item;
  End If;

  --1.更新异常状态
  Update 门诊费用记录
  Set 费用状态 = 0
  Where Nvl(费用状态, 0) = 1 And 结帐id In (Select Distinct 结帐id From 病人预交记录 Where 结算序号 = r_Balance.结算序号);

  Update 费用补充记录 Set 费用状态 = 0 Where 结算序号 = r_Balance.结算序号;
  If Sql%NotFound Then
    v_Err_Msg := '未找到医保补结算数据，可能补他人进行了退费或作废操作!';
    Raise Err_Item;
  End If;

  --2.更新校对标志,会话号
  Update 病人预交记录 Set 校对标志 = 0, 会话号 = Null Where 结算序号 = r_Balance.结算序号;

  --3.更新 是否电子票据 标记 
  Select Max(a.是否电子票据)
  Into n_是否电子票据
  From 病人预交记录 A,
       (Select 结算id
         From (Select b.结算id
                From 费用补充记录 B
                Where b.No = r_Balance.No And b.记录性质 = 1 And b.记录状态 In (1, 3)
                Order By b.登记时间)
         Where Rownum < 2) B
  Where a.结帐id = b.结算id;

  Update 病人预交记录 Set 是否电子票据 = n_是否电子票据 Where 结算序号 = r_Balance.结算序号 And 结算性质 = 6;

  --4.更新人员缴款数据
  For c_缴款 In (Select a.结算方式, a.操作员姓名, Nvl(Sum(a.冲预交), 0) As 冲预交
               From 病人预交记录 A
               Where a.结算序号 = r_Balance.结算序号 And Mod(a.记录性质, 10) <> 1
               Group By a.结算方式, a.操作员姓名
               Having Nvl(Sum(a.冲预交), 0) <> 0) Loop
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) + c_缴款.冲预交
    Where 收款员 = c_缴款.操作员姓名 And 性质 = 1 And 结算方式 = c_缴款.结算方式;
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额
        (收款员, 结算方式, 性质, 余额)
      Values
        (c_缴款.操作员姓名, c_缴款.结算方式, 1, c_缴款.冲预交);
    End If;
  End Loop;

  --消息集成处理
  For c_结帐id In (Select Distinct 结帐id From 病人预交记录 Where 记录性质 = 3 And 结算序号 = r_Balance.结算序号) Loop
    b_Message.Zlhis_Charge_004(2, c_结帐id.结帐id);
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_费用补充结算_完成退费;
/

Create Or Replace Procedure Zl_门诊转住院_收费转出
(
  No_In           住院费用记录.No%Type,
  操作员编号_In   住院费用记录.操作员编号%Type,
  操作员姓名_In   住院费用记录.操作员姓名%Type,
  退费时间_In     住院费用记录.发生时间%Type,
  门诊退费_In     Number := 0,
  入院科室id_In   住院费用记录.开单部门id%Type := Null,
  主页id_In       住院费用记录.主页id%Type := Null,
  结算方式_In     病人预交记录.结算方式%Type := Null,
  结帐id_In       病人预交记录.结帐id%Type := Null,
  原结帐id_In     病人预交记录.结帐id%Type := Null,
  误差费_In       病人预交记录.冲预交%Type := Null,
  缴款组id_In     病人预交记录.缴款组id%Type := Null,
  预交电子票据_In 病人预交记录.预交电子票据%Type := 0
) As
  --入参:
  --  门诊退费_In:0-门诊转住院立即销帐;1-门诊退费模式;=1时:入院科室id_In和主页ID_IN可以不传入
  --  缴款组ID_In:NULL表示款传入缴款组ID（需重新读取);0-表示已经读取，不用再读取;>0表示已经读取出具体的缴款组 
  --  预交电子票据_In:预交款是否启用电子票据
  n_Count      Number(5);
  n_原结帐id   住院费用记录.结帐id%Type;
  n_实收金额   门诊费用记录.实收金额%Type;
  n_组id       财务缴款分组.Id%Type;
  n_病人id     病人信息.病人id%Type;
  v_预交no     病人预交记录.No%Type;
  n_预交金额   病人预交记录.冲预交%Type;
  n_打印id     票据使用明细.打印id%Type;
  n_开单部门id 住院费用记录.开单部门id%Type;
  v_开单人     门诊费用记录.开单人%Type;
  n_结帐id     门诊费用记录.结帐id%Type;
  v_误差费     结算方式.名称%Type;
  n_误差费     病人预交记录.冲预交%Type;
  n_返回值     病人余额.费用余额%Type;

  n_剩余预交     病人预交记录.冲预交%Type;
  v_结算方式     结算方式.名称%Type;
  v_缺省结算方式 结算方式.名称%Type;
  v_Nos          Varchar2(3000);
  v_结帐ids      Varchar2(3000);
  v_原结帐ids    Varchar2(3000);
  n_Tempid       病人预交记录.Id%Type;
  n_医保         Number;
  n_存在         Number;
  n_部分退费     Number;
  n_退费条数     Number;
  n_费用状态     门诊费用记录.费用状态%Type;
  n_关联交易id   病人预交记录.关联交易id%Type;
  n_充值id       病人预交记录.Id%Type;
  n_是否存在医保 Number(2);
  Err_Item Exception;
  v_Err_Msg Varchar2(200);

  Procedure 病人预交款_Del
  (
    冲销id_In    病人预交记录.结帐id%Type,
    结帐ids_In   Varchar2,
    退预交款_In  病人预交记录.冲预交%Type,
    退款合计_Out Out 病人预交记录.冲预交%Type
  ) As
    --退预交款_In：传入空时，表示全退,否则按退预交款方式进行退款
    n_全退     Number(2);
    n_退预交款 病人预交记录.冲预交%Type;
    n_冲预交   病人预交记录.冲预交%Type;
  Begin
    n_全退     := 1;
    n_退预交款 := Nvl(退预交款_In, 0);
    If Nvl(退预交款_In, 0) <> 0 Then
      n_全退 := 0;
    End If;
  
    退款合计_Out := 0;
    For r_Prepay In (Select NO, Max(Decode(记录性质, 1, 实际票号, Null)) As 实际票号, 病人id, 主页id, Max(科室id) As 科室id,
                            Max(结算方式) As 结算方式, Max(结算号码) As 结算号码, Max(缴款单位) As 缴款单位, Max(单位开户行) As 单位开户行,
                            Max(单位帐号) As 单位帐号, Min(收款时间) As 收款时间, Sum(冲预交) As 冲预交, 卡类别id, 结算卡序号, Max(卡号) As 卡号, 交易流水号,
                            Max(交易说明) As 交易说明, Max(合作单位) As 合作单位, 结算性质, Decode(Nvl(卡类别id, 0), 0, 0, 关联交易id) As 关联交易id,
                            Max(预交类别) As 预交类别, Max(交易时间) As 交易时间, Max(交易人员) As 交易人员, Max(Decode(记录性质, 1, ID, 0)) As 预交id
                     From 病人预交记录 A
                     Where 记录性质 In (1, 11) And a.结帐id In (Select Column_Value From Table(f_Str2List(结帐ids_In))) And
                           Nvl(冲预交, 0) <> 0
                     Group By NO, 病人id, 主页id, 卡类别id, 结算卡序号, 交易流水号, 结算性质, 关联交易id) Loop
    
      n_冲预交 := Nvl(r_Prepay.冲预交, 0);
      If n_全退 = 0 Then
        If n_退预交款 <> 0 Then
          If n_退预交款 > n_冲预交 Then
            n_退预交款 := Round(n_退预交款 - n_冲预交, 6);
          Else
            n_冲预交   := Nvl(n_退预交款, 0);
            n_退预交款 := 0;
          End If;
        Else
          Exit;
        End If;
      
      End If;
      Select 病人预交记录_Id.Nextval Into n_Tempid From Dual;
    
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
         结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 预交类别, 结算性质, 关联交易id, 交易时间, 交易人员)
        Select n_Tempid, r_Prepay.No, r_Prepay.实际票号, 11, 1, r_Prepay.病人id, r_Prepay.主页id, r_Prepay.科室id, Null,
               r_Prepay.结算方式, r_Prepay.结算号码, Null, r_Prepay.缴款单位, r_Prepay.单位开户行, r_Prepay.单位帐号, 退费时间_In, 操作员姓名_In,
               操作员编号_In, -1 * n_冲预交, 冲销id_In, n_组id, r_Prepay.卡类别id, r_Prepay.结算卡序号, r_Prepay.卡号, r_Prepay.交易流水号,
               r_Prepay.交易说明, r_Prepay.合作单位, -1 * 冲销id_In, Nvl(r_Prepay.预交类别, 1), r_Prepay.结算性质, r_Prepay.关联交易id,
               r_Prepay.交易时间, r_Prepay.交易人员
        From Dual;
    
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) + n_冲预交
      Where 病人id = r_Prepay.病人id And 类型 = Nvl(r_Prepay.预交类别, 1) And 性质 = 1
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 类型, 预交余额, 性质)
        Values
          (r_Prepay.病人id, Nvl(r_Prepay.预交类别, 1), n_冲预交, 1);
        n_返回值 := n_冲预交;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 病人余额
        Where 病人id = r_Prepay.病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
      End If;
    
      n_充值id := r_Prepay.预交id;
      If Nvl(n_充值id, 0) = 0 Then
        Select Max(ID) Into n_充值id From 病人预交记录 Where NO = r_Prepay.No And 记录性质 = 1 And 记录状态 In (1, 3);
      End If;
      If Nvl(n_充值id, 0) <> 0 Then
        Update 预交单据余额
        Set 预交余额 = Nvl(预交余额, 0) + n_冲预交
        Where 病人id = r_Prepay.病人id And 预交id = n_充值id
        Returning 预交余额 Into n_返回值;
      
        If Sql%RowCount = 0 Then
          Insert Into 预交单据余额
            (预交id, 病人id, 预交类别, 预交余额)
          Values
            (n_充值id, r_Prepay.病人id, Nvl(r_Prepay.预交类别, 1), n_冲预交);
          n_返回值 := n_冲预交;
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 预交单据余额 Where 预交id = n_充值id And Nvl(预交余额, 0) = 0;
        End If;
      End If;
      退款合计_Out := 退款合计_Out + n_冲预交;
    End Loop;
  
  End 病人预交款_Del;

  ---------------------------------------------------------------------------------------------
  --生成冲预交及产生预交款
  Procedure 病人结算_Strict
  (
    冲销id_In       病人预交记录.结帐id%Type,
    结帐ids_In      Varchar2,
    剩余款_In       病人预交记录.冲预交%Type := Null,
    是否生成预交_In Number := 1,
    冲销医保_In     Number := 1
  ) As
    --剩余款_In：传入空时，表示全退,否则按剩余款进行退款
    --是否生成预交_In:1-生成新的住院预交;0-不生成新的住院预交
    --冲销医保_In:1-表示对医保进行冲销，否则不冲销这部分数据
    n_退款金额   病人预交记录.冲预交%Type;
    n_冲预交     病人预交记录.冲预交%Type;
    n_全退       Number(2);
    n_关联交易id 病人预交记录.关联交易id%Type;
  Begin
    n_全退     := 1;
    n_退款金额 := Nvl(剩余款_In, 0);
    If Nvl(剩余款_In, 0) <> 0 Then
      n_全退 := 0;
    End If;
  
    For r_Pay In (Select a.结算方式, Sum(a.冲预交) As 冲预交, 2 As 预交类别, a.卡类别id, a.结算卡序号, a.卡号, a.交易流水号 As 交易流水号,
                         Max(Decode(a.记录状态, 2, Null, a.交易说明)) As 交易说明, a.合作单位, b.性质, a.关联交易id,
                         Max(Decode(a.记录状态, 2, 0, a.Id)) As 预交id, Max(结算号码) As 结算号码, Max(摘要) As 摘要,
                         Sign(Sum(a.冲预交)) As 标志
                  From 病人预交记录 A, 结算方式 B
                  Where a.记录性质 = 3 And a.记录状态 In (1, 2, 3) And
                        a.结帐id In (Select Column_Value From Table(f_Str2List(结帐ids_In))) And a.结算方式 = b.名称(+) And
                        a.结算方式 Is Not Null
                  Group By a.结算方式, 预交类别, a.卡类别id, a.结算卡序号, a.卡号, b.性质, a.交易流水号, a.合作单位, a.关联交易id
                  Having Sum(a.冲预交) <> 0
                  Order By 卡类别id, 标志 Desc, 性质) Loop
    
      If (Nvl(冲销医保_In, 0) = 1) Or (Instr('34', Nvl(r_Pay.性质, 0))) = 0 Or
         (冲销医保_In = 0 And Instr('34', Nvl(r_Pay.性质, 0)) > 0 And Nvl(r_Pay.卡类别id, 0) <> 0) Then
      
        n_冲预交 := Nvl(r_Pay.冲预交, 0);
      
        If n_全退 = 0 Then
          If n_退款金额 <> 0 Then
            If n_退款金额 > n_冲预交 Then
              n_退款金额 := Round(n_退款金额 - n_冲预交, 6);
            Else
              n_冲预交   := Nvl(n_退款金额, 0);
              n_退款金额 := 0;
            End If;
          Else
            Exit;
          End If;
        End If;
      
        --4.1产生预交款单据 (不存在部分退费的情况)
        --所有单据,按规则生成预交款单据
        --因为收款后立即缴款,所以人员缴款余额无变化
        --一卡通:只有含有医保或多种结算方式的，才会调用接口处理
        --1.一卡通
        n_Count      := 0;
        n_关联交易id := r_Pay.关联交易id;
        If r_Pay.卡类别id Is Not Null Or Nvl(r_Pay.结算卡序号, 0) <> 0 Then
          If Nvl(r_Pay.关联交易id, 0) = 0 And Nvl(r_Pay.卡类别id, 0) <> 0 Then
          
            n_关联交易id := r_Pay.预交id;
            Update 病人预交记录
            Set 关联交易id = n_关联交易id
            Where 结帐id In (Select Column_Value From Table(f_Str2List(结帐ids_In))) And 记录性质 = 3 And 记录状态 In (2, 3) And
                  卡类别id = r_Pay.卡类别id;
          End If;
        
          --三方卡及消费卡(消费卡在完成时回退)
          If Instr('34', r_Pay.性质) = 0 And Nvl(r_Pay.结算卡序号, 0) = 0 Then
            --三方方，需要检 查是否多种结算方式，如果是多种结算方式，需要调用退款接口
            Select Count(Distinct 结算方式)
            Into n_Count
            From 病人预交记录
            Where 结帐id In (Select Column_Value From Table(f_Str2List(结帐ids_In))) And 卡类别id = r_Pay.卡类别id And
                  Nvl(关联交易id, 0) = Nvl(r_Pay.关联交易id, 0);
          Else
            --消费卡及三方退款带有医保的:需要调用接口
            n_Count := 2;
          End If;
        
          If n_Count > 1 Then
          
            --需要调用接口退的，将附加标志填为1,以便退款
            Update 病人预交记录
            Set 冲预交 = 冲预交 + (-1 * n_冲预交)
            Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = 冲销id_In And 结算方式 = r_Pay.结算方式 And Nvl(关联交易id, 0) = Nvl(n_关联交易id, 0) And
                  Nvl(卡类别id, 0) = Nvl(r_Pay.卡类别id, 0) And Nvl(结算卡序号, 0) = Nvl(r_Pay.结算卡序号, 0);
          
            If Sql%RowCount = 0 Then
              Select 病人预交记录_Id.Nextval Into n_Tempid From Dual;
              Insert Into 病人预交记录
                (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
                 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质, 关联交易id, 交易人员, 交易时间, 附加标志)
              Values
                (n_Tempid, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * n_冲预交, r_Pay.结算方式, Null, 退费时间_In, Null,
                 Null, Null, 操作员编号_In, 操作员姓名_In, '', n_组id, r_Pay.卡类别id, r_Pay.结算卡序号, r_Pay.卡号, r_Pay.交易流水号, r_Pay.交易说明,
                 r_Pay.合作单位, 冲销id_In, -1 * 冲销id_In, 1, 3, Decode(n_关联交易id, 0, Null, n_关联交易id), 操作员姓名_In, 退费时间_In, -1);
            
              --调用其他结算信息更新
              Zl_Custom_Balance_Update(n_Tempid);
            End If;
          
            n_Count    := 2;
            n_费用状态 := 1;
          End If;
        End If;
      
        If n_Count <= 1 Then
          --新生成预交款
          v_结算方式 := Nvl(r_Pay.结算方式, v_缺省结算方式);
        
          --医保，误差费不产生预交单
          If Instr('349', r_Pay.性质) = 0 And Nvl(是否生成预交_In, 0) = 1 And n_冲预交 <> 0 Then
            --一卡通，每一笔都生成一条预交款记录
            --其它，同一种结算方式只生成一条预交款记录
            Update 病人预交记录
            Set 金额 = Nvl(金额, 0) + n_冲预交
            Where 记录性质 = 1 And 记录状态 = 1 And 收款时间 = 退费时间_In And 病人id + 0 = n_病人id And 结算方式 = v_结算方式 And
                  (Nvl(卡类别id, 0) = 0 And Nvl(r_Pay.卡类别id, 0) = 0)
            Returning ID Into n_充值id;
            If Sql%RowCount = 0 Then
              v_预交no := Nextno(11);
              Select 病人预交记录_Id.Nextval Into n_Tempid From Dual;
              Insert Into 病人预交记录
                (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要, 缴款组id,
                 预交类别, 卡类别id, 关联交易id, 交易人员, 交易时间, 卡号, 交易说明, 交易流水号, 结算号码, 预交电子票据)
              Values
                (n_Tempid, v_预交no, Null, 1, 1, n_病人id, 主页id_In, 入院科室id_In, n_冲预交, v_结算方式, 退费时间_In, Null, Null, Null,
                 操作员编号_In, 操作员姓名_In, '门诊转住院预交', n_组id, r_Pay.预交类别, r_Pay.卡类别id, Nvl(r_Pay.关联交易id, n_Tempid), 操作员姓名_In,
                 退费时间_In, r_Pay.卡号, r_Pay.交易说明, r_Pay.交易流水号, r_Pay.结算号码, 预交电子票据_In);
              n_充值id := n_Tempid;
            End If;
            n_费用状态 := 1;
          
            --病人余额
            Update 病人余额
            Set 预交余额 = Nvl(预交余额, 0) + n_冲预交
            Where 性质 = 1 And 病人id = n_病人id And 类型 = 2
            Returning 预交余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 病人余额 (病人id, 性质, 类型, 预交余额, 费用余额) Values (n_病人id, 1, 2, n_冲预交, 0);
              n_返回值 := n_冲预交;
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 病人余额
              Where 病人id = n_病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
            End If;
          
            If Nvl(n_充值id, 0) <> 0 Then
              Update 预交单据余额
              Set 预交余额 = Nvl(预交余额, 0) + Nvl(n_冲预交, 0)
              Where 病人id = n_病人id And 预交id = n_充值id
              Returning 预交余额 Into n_返回值;
            
              If Sql%RowCount = 0 Then
                Insert Into 预交单据余额
                  (预交id, 病人id, 预交类别, 预交余额)
                Values
                  (n_充值id, n_病人id, 2, Nvl(n_冲预交, 0));
                n_返回值 := Nvl(n_冲预交, 0);
              End If;
              If Nvl(n_返回值, 0) = 0 Then
                Delete From 预交单据余额 Where 预交id = n_充值id And Nvl(预交余额, 0) = 0;
              End If;
            End If;
          End If;
        
          Update 病人预交记录
          Set 冲预交 = 冲预交 + (-1 * n_冲预交), 关联交易id = Nvl(关联交易id, n_关联交易id)
          Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = 冲销id_In And 结算方式 = v_结算方式 And Nvl(关联交易id, 0) = Nvl(n_关联交易id, 0) And
                Nvl(卡类别id, 0) = Nvl(r_Pay.卡类别id, 0) And Nvl(结算卡序号, 0) = Nvl(r_Pay.结算卡序号, 0);
          If Sql%RowCount = 0 Then
            Select 病人预交记录_Id.Nextval Into n_Tempid From Dual;
          
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
               缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质, 关联交易id, 交易人员, 交易时间)
            Values
              (n_Tempid, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * n_冲预交, v_结算方式, Null, 退费时间_In, Null, Null,
               Null, 操作员编号_In, 操作员姓名_In, '', n_组id, r_Pay.卡类别id, r_Pay.结算卡序号, r_Pay.卡号, r_Pay.交易流水号, r_Pay.交易说明,
               r_Pay.合作单位, 冲销id_In, -1 * 冲销id_In, 1, 3, Decode(n_关联交易id, 0, Null, n_关联交易id), 操作员姓名_In, 退费时间_In);
          End If;
        
          --4.2缴款数据处理
          --   因为没有实际收病人的钱,所以不处理
          --部分退费情况，退原预交记录
          If r_Pay.性质 In (3, 4) Then
            Update 人员缴款余额
            Set 余额 = Nvl(余额, 0) - n_冲预交
            Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_Pay.结算方式
            Returning 余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 人员缴款余额
                (收款员, 结算方式, 性质, 余额)
              Values
                (操作员姓名_In, r_Pay.结算方式, 1, -1 * n_冲预交);
              n_返回值 := n_冲预交;
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 人员缴款余额
              Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_Pay.结算方式 And Nvl(余额, 0) = 0;
            End If;
          End If;
        End If;
      End If;
    End Loop;
  
  End 病人结算_Strict;

  Procedure 费用记录_Strict
  (
    冲销id_In   门诊费用记录.结帐id%Type,
    Nos_In      Varchar2,
    缴款组id_In 门诊费用记录.缴款组id%Type
  ) As
  
  Begin
  
    --更新费用审核记录
    Update 费用审核记录
    Set 记录状态 = 2
    Where 费用id In (Select a.Id
                   From 门诊费用记录 A
                   Where a.No In (Select Column_Value From Table(f_Str2List(Nos_In))) And Mod(a.记录性质, 10) = 1 And
                         a.记录状态 In (1, 3)) And 性质 = 1;
    --作废门诊记录
    Update 门诊费用记录
    Set 记录状态 = 3
    Where Mod(记录性质, 10) = 1 And 记录状态 = 1 And NO In (Select Column_Value From Table(f_Str2List(Nos_In)));
  
    For r_Clinic In (Select Min(Mod(a.记录性质, 10)) As 记录性质, a.No, a.序号, a.从属父号, a.价格父号, a.医嘱序号, a.病人id, a.姓名, a.性别, a.年龄,
                            a.病人科室id, a.费别, a.收费类别, a.收费细目id, a.计算单位, a.保险项目否, a.保险大类id, a.保险编码, a.费用类型, a.发药窗口, a.付数,
                            Sum(a.数次) As 数次, a.加班标志, a.附加标志, a.收入项目id, a.收据费目, a.标准单价, Sum(a.应收金额) As 应收金额,
                            Sum(a.实收金额) As 实收金额, Sum(a.统筹金额) As 统筹金额, a.开单部门id, a.开单人, a.执行部门id, a.划价人,
                            Max(a.记帐单id) As 记帐单id, Max(a.是否急诊) As 是否急诊, a.发生时间, Min(a.实际票号) As 实际票号,
                            Nvl(Min(Decode(a.记录状态, 2, a.执行状态, 0)), 0) - 1 As 执行状态, 挂号id, 主页id, 病人病区id
                     From 门诊费用记录 A
                     Where a.No In (Select Column_Value From Table(f_Str2List(Nos_In))) And Mod(a.记录性质, 10) = 1 And
                           Nvl(a.附加标志, 0) Not In (8, 9)
                     Group By a.No, a.序号, a.从属父号, a.价格父号, a.医嘱序号, a.病人id, a.姓名, a.性别, a.年龄, a.病人科室id, a.费别, a.收费类别,
                              a.收费细目id, a.计算单位, a.保险项目否, a.保险大类id, a.保险编码, a.费用类型, a.发药窗口, a.付数, a.加班标志, a.附加标志, a.收入项目id,
                              a.收据费目, a.标准单价, a.开单部门id, a.开单人, a.执行部门id, a.划价人, a.发生时间, 挂号id, 主页id, 病人病区id
                     Having Sum(a.数次) <> 0) Loop
      Insert Into 门诊费用记录
        (ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 门诊标志, 医嘱序号, 病人id, 标识号, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位,
         保险项目否, 保险大类id, 保险编码, 费用类型, 发药窗口, 付数, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额, 记帐费用, 开单部门id, 开单人,
         发生时间, 登记时间, 执行部门id, 划价人, 操作员编号, 操作员姓名, 记帐单id, 摘要, 是否急诊, 缴款组id, 结帐id, 结帐金额, 执行状态, 费用状态, 挂号id, 主页id, 病人病区id)
      Values
        (病人费用记录_Id.Nextval, r_Clinic.记录性质, r_Clinic.No, r_Clinic.实际票号, 2, r_Clinic.序号, r_Clinic.从属父号, r_Clinic.价格父号, 1,
         r_Clinic.医嘱序号, r_Clinic.病人id, '', r_Clinic.姓名, r_Clinic.性别, r_Clinic.年龄, r_Clinic.病人科室id, r_Clinic.费别,
         r_Clinic.收费类别, r_Clinic.收费细目id, r_Clinic.计算单位, r_Clinic.保险项目否, r_Clinic.保险大类id, r_Clinic.保险编码, r_Clinic.费用类型,
         r_Clinic.发药窗口, r_Clinic.付数, -1 * r_Clinic.数次, r_Clinic.加班标志, r_Clinic.附加标志, r_Clinic.收入项目id, r_Clinic.收据费目,
         r_Clinic.标准单价, -1 * r_Clinic.应收金额, -1 * r_Clinic.实收金额, -1 * r_Clinic.统筹金额, 0, r_Clinic.开单部门id, r_Clinic.开单人,
         r_Clinic.发生时间, 退费时间_In, r_Clinic.执行部门id, r_Clinic.划价人, 操作员编号_In, 操作员姓名_In, r_Clinic.记帐单id, '', r_Clinic.是否急诊,
         缴款组id_In, 冲销id_In, -1 * r_Clinic.实收金额, r_Clinic.执行状态, 0, r_Clinic.挂号id, r_Clinic.主页id, r_Clinic.病人病区id);
    End Loop;
  End 费用记录_Strict;

  Procedure 医保结算明细_Strict
  (
    冲销id_In    病人预交记录.结帐id%Type,
    结帐ids_In   Varchar2,
    Nos_In       Varchar2,
    缴款组id_In  门诊费用记录.缴款组id%Type,
    退款合计_Out Out 病人预交记录.冲预交%Type,
    仅冲销_In    Number := 0
  ) As
    --仅冲销_In:1-只作销医保明细处理;0-除了冲销，还要处理余额及预交记录
    v_卡号 病人预交记录.卡号%Type;
  Begin
    --医保退款
    退款合计_Out := 0;
    For r_医保 In (Select 结帐id, NO, 结算方式, 金额, 备注, 卡类别id, 关联交易id, 交易流水号, 交易说明
                 From 医保结算明细 A
                 Where Instr(',' || Nos_In || ',', ',' || NO || ',') > 0 And
                       结帐id In (Select Column_Value From Table(f_Str2List(结帐ids_In)))) Loop
    
      If Nvl(仅冲销_In, 0) = 0 Then
      
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) - r_医保.金额
        Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_医保.结算方式
        Returning 余额 Into n_返回值;
      
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额
            (收款员, 结算方式, 性质, 余额)
          Values
            (操作员姓名_In, r_医保.结算方式, 1, -1 * r_医保.金额);
          n_返回值 := r_医保.金额;
        End If;
      
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 人员缴款余额
          Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_医保.结算方式 And Nvl(余额, 0) = 0;
        End If;
      
        Select Max(关联交易id), Max(卡号)
        Into n_关联交易id, v_卡号
        From 病人预交记录
        Where 结帐id = r_医保.结帐id And 结算方式 = r_医保.结算方式 And Nvl(卡类别id, 0) = Nvl(r_医保.卡类别id, 0);
      
        Update 病人预交记录
        Set 冲预交 = 冲预交 + (-1 * r_医保.金额)
        Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = 冲销id_In And 结算方式 = r_医保.结算方式 And Nvl(卡类别id, 0) = Nvl(r_医保.卡类别id, 0) And
              Nvl(关联交易id, 0) = Nvl(n_关联交易id, 0);
        If Sql%RowCount = 0 Then
          Insert Into 病人预交记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
             缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质, 关联交易id, 交易时间, 交易人员, 附加标志)
          Values
            (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * r_医保.金额, r_医保.结算方式, Null, 退费时间_In,
             Null, Null, Null, 操作员编号_In, 操作员姓名_In, r_医保.备注, 缴款组id_In, r_医保.卡类别id, Null, v_卡号, r_医保.交易流水号, r_医保.交易说明,
             Null, 冲销id_In, -1 * 冲销id_In, Decode(Nvl(r_医保.卡类别id, 0), 0, 0, 1), 3, n_关联交易id, 退费时间_In, 操作员姓名_In,
             Decode(Nvl(r_医保.卡类别id, 0), 0, Null, -1));
        End If;
      
        Update 病人预交记录
        Set 记录状态 = 3
        Where 记录性质 = 3 And 记录状态 = 1 And 结帐id In (Select Column_Value From Table(f_Str2List(v_原结帐ids))) And
              结算方式 = r_医保.结算方式;
      End If;
    
      Update 医保结算明细
      Set 金额 = 金额 + (-1 * r_医保.金额)
      Where NO = r_医保.No And 结帐id = 冲销id_In And 结算方式 = r_医保.结算方式 And Nvl(卡类别id, 0) = Nvl(r_医保.卡类别id, 0) And
            Nvl(关联交易id, 0) = Nvl(r_医保.关联交易id, 0);
    
      If Sql%RowCount = 0 Then
        Insert Into 医保结算明细
          (结帐id, NO, 结算方式, 金额, 卡类别id, 关联交易id, 交易流水号, 交易说明)
        Values
          (冲销id_In, r_医保.No, r_医保.结算方式, -1 * r_医保.金额, r_医保.卡类别id, r_医保.关联交易id, r_医保.交易流水号, r_医保.交易说明);
      End If;
      退款合计_Out := 退款合计_Out + Nvl(r_医保.金额, 0);
    End Loop;
  End 医保结算明细_Strict;

  Procedure 门诊结算票据_回收(Nos_In Varchar2) As
    n_打印id 票据打印内容.Id%Type;
  Begin
    --2.票据收回
    --可能以前没有打印,无收回
    For r_Nos In (Select Column_Value As NO From Table(f_Str2List(Nos_In))) Loop
    
      Select Nvl(Max(ID), 0)
      Into n_打印id
      From (Select b.Id
             From 票据使用明细 A, 票据打印内容 B
             Where a.打印id = b.Id And a.性质 = 1 And a.原因 In (1, 3) And b.数据性质 = 1 And b.No = r_Nos.No
             Order By a.使用时间 Desc)
      Where Rownum < 2;
      If n_打印id > 0 Then
        --多张单据循环调用时只能收回一次
        Select Count(打印id) Into n_Count From 票据使用明细 Where 票种 = 1 And 性质 = 2 And 打印id = n_打印id;
        If n_Count = 0 Then
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
            Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, 退费时间_In, 操作员姓名_In, 票据金额
            From 票据使用明细
            Where 打印id = n_打印id And 票种 = 1 And 性质 = 1;
        End If;
      End If;
    End Loop;
  End 门诊结算票据_回收;

  Procedure 门诊转住费用_Over
  (
    冲销id_In   病人预交记录.结帐id%Type,
    原结帐id_In 病人预交记录.结帐id%Type := Null
  ) As
    n_费用状态 Number(2);
  Begin
  
    Delete From 病人预交记录 Where 结帐id = Nvl(原结帐id_In, 0) And 摘要 = '预交临时记录' And 记录性质 = 3;
  
    Delete From 病人预交记录
    Where 结帐id = 冲销id_In And 记录性质 = 3 And 记录状态 = 2 And 冲预交 = 0 And 结算方式 Is Not Null;
  
    --合法性检查，冲预交要与费用结算合计一致，不一致时，直接退出
    Select Nvl(Max(1), 0) Into n_费用状态 From 病人预交记录 Where 结帐id = 冲销id_In And 附加标志 = -1;
  
    Update 门诊费用记录 Set 费用状态 = Nvl(n_费用状态, 0) Where 结帐id = 冲销id_In;
  
    If Nvl(n_费用状态, 0) = 0 Then
      --不存在异常，直接更新
      Update 病人预交记录 Set 校对标志 = 0, 附加标志 = Null Where 结帐id = 冲销id_In;
    Else
      Update 病人预交记录 A
      Set 校对标志 = 2
      Where 结帐id = 冲销id_In And
            ((Nvl(附加标志, 0) <> -1 And 卡类别id Is Not Null) Or
            (Exists (Select 1 From 结算方式 Where a.结算方式 = 名称 And 性质 In (3, 4) And 卡类别id Is Null)));
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 科室id, 主页id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 校对标志, 结算性质, 结算序号)
        Select 病人预交记录_Id.Nextval, 3, Null, 2, n_病人id, Null, Null, Null, 退费时间_In, 操作员编号_In, 操作员姓名_In, Null, 冲销id_In,
               n_组id, 0, 3, -1 * 冲销id_In
        From Dual;
    End If;
  End 门诊转住费用_Over;

Begin
  n_组id := 缴款组id_In;
  If n_组id Is Null Then
    n_组id := Zl_Get组id(操作员姓名_In);
  End If;
  If Nvl(n_组id, 0) = 0 Then
    n_组id := Null;
  End If;

  --误差费
  Select Max(名称) Into v_误差费 From 结算方式 Where 性质 = 9 And Rownum < 2;
  If v_误差费 Is Null Then
    v_Err_Msg := '没有发现误差结算方式，请检查是否正确设置！';
    Raise Err_Item;
  End If;

  n_误差费       := 误差费_In;
  v_缺省结算方式 := 结算方式_In;
  If v_缺省结算方式 Is Null Then
    Select Nvl(Max(名称), '现金') Into v_缺省结算方式 From 结算方式 Where 性质 = 1 And Rownum < 2;
  End If;

  If 原结帐id_In Is Null Then
  
    Select Count(NO), Sum(实收金额)
    Into n_Count, n_实收金额
    From 门诊费用记录
    Where NO = No_In And Mod(记录性质, 10) = 1;
  
    If n_Count = 0 Or n_实收金额 = 0 Then
      v_Err_Msg := '单据' || No_In || '不是收费单据或因并发原因他人操作了该单据,不能转为住院费用.';
      Raise Err_Item;
    End If;
  
    Select 结帐id, 病人id, 开单部门id, 开单人
    Into n_原结帐id, n_病人id, n_开单部门id, v_开单人
    From 门诊费用记录
    Where NO = No_In And Mod(记录性质, 10) = 1 And 记录状态 In (1, 3) And Rownum < 2;
  
    --1.1作废费用记录
    n_结帐id := 结帐id_In;
    If n_结帐id Is Null Then
      Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
    End If;
  
    费用记录_Strict(n_结帐id, No_In, n_组id);
  
    --1.2作废预交记录
    --作废冲预交部分
    For r_结账id In (Select Distinct 结帐id
                   From 门诊费用记录
                   Where NO In (Select Distinct NO
                                From 门诊费用记录
                                Where 结帐id In (Select 结帐id
                                               From 病人预交记录
                                               Where 结算序号 In (Select b.结算序号
                                                              From 门诊费用记录 A, 病人预交记录 B
                                                              Where a.No = No_In And b.结算序号 < 0 And Mod(a.记录性质, 10) = 1 And
                                                                    a.记录状态 <> 0 And a.结帐id = b.结帐id))) And
                         Mod(记录性质, 10) = 1 And 记录状态 <> 0
                   Union
                   Select Distinct 结帐id
                   From 门诊费用记录
                   Where NO In (Select Distinct NO
                                From 门诊费用记录
                                Where 结帐id In (Select a.结帐id
                                               From 门诊费用记录 A, 病人预交记录 B
                                               Where a.No = No_In And b.结算序号 > 0 And Mod(a.记录性质, 10) = 1 And a.记录状态 <> 0 And
                                                     a.结帐id = b.结帐id)) And Mod(记录性质, 10) = 1 And 记录状态 <> 0) Loop
      v_原结帐ids := v_原结帐ids || ',' || r_结账id.结帐id;
    End Loop;
  
    v_原结帐ids := Substr(v_原结帐ids, 2);
  
    Select Nvl(Max(1), 0)
    Into n_医保
    From 保险结算记录
    Where 记录id In (Select Column_Value From Table(f_Str2List(v_原结帐ids))) And Rownum < 2 And 卡类别id Is Not Null;
  
    If n_医保 = 1 Then
    
      Select Nvl(Max(1), 0)
      Into n_存在
      From 医保结算明细
      Where NO = No_In And 结帐id In (Select Column_Value From Table(f_Str2List(v_原结帐ids))) And Rownum < 2 And
            卡类别id Is Not Null;
    
      If n_存在 = 0 Then
        v_Err_Msg := '当前单据' || No_In || '不存在医保结算明细,无法进行门诊转住院!';
        Raise Err_Item;
      End If;
    
    End If;
  
    --先进行误差费处理
    n_返回值 := Round(n_实收金额, 2) - Nvl(n_实收金额, 0);
    If Nvl(n_返回值, 0) <> 0 Then
      n_误差费 := Nvl(n_误差费, 0) + n_返回值;
    End If;
    n_实收金额 := Round(n_实收金额, 2);
  
    --医保明细冲销: 冲销id_In, 结帐ids_In,Nos_In,缴款组id_In 退款合计_In out
    医保结算明细_Strict(n_结帐id, v_原结帐ids, No_In, n_组id, n_预交金额);
    n_实收金额 := n_实收金额 - Nvl(n_预交金额, 0);
  
    If n_实收金额 <> 0 Then
      --退预交款:
      病人预交款_Del(n_结帐id, v_原结帐ids, n_实收金额, n_预交金额);
      n_实收金额 := n_实收金额 - Nvl(n_预交金额, 0);
    End If;
  
    --2.票据收回
    门诊结算票据_回收(No_In);
  
    --3.缴款数据处理(
    --   现有两种情况:
    --    1. 转出过程直接销帐的,则缴款数据不增加;
    --    2. 先转出,再到门诊退款退票,则需要进行缴款数据处理
  
    If Nvl(门诊退费_In, 0) = 1 Or n_实收金额 <> 0 Then
      If Nvl(门诊退费_In, 0) = 1 Then
        --不生成预交:退款转预交(不产生票据,由操作员通过重打进行)
        病人结算_Strict(n_结帐id, v_原结帐ids, n_实收金额, 0, 0);
      Elsif n_实收金额 <> 0 Then
        --生成预交:
        病人结算_Strict(n_结帐id, v_原结帐ids, n_实收金额, 1, 0);
      End If;
    End If;
  
    If n_误差费 Is Not Null Then
      Update 病人预交记录 Set 冲预交 = Nvl(冲预交, 0) + n_误差费 Where 结帐id = n_结帐id And 结算方式 = v_误差费;
      If Sql%NotFound Then
        Select 病人预交记录_Id.Nextval Into n_Tempid From Dual;
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
           缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质, 关联交易id)
        Values
          (n_Tempid, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, n_误差费, v_误差费, Null, 退费时间_In, Null, Null, Null,
           操作员编号_In, 操作员姓名_In, '', n_组id, Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 0, 3, n_Tempid);
      End If;
    End If;
  
    门诊转住费用_Over(n_结帐id, n_原结帐id);
    Return;
  End If;

  --医保按结算转出
  For r_Nos In (Select Distinct a.No
                From 门诊费用记录 A
                Where Mod(a.记录性质, 10) = 1 And a.记录状态 In (1, 3) And a.结帐id = 原结帐id_In) Loop
    v_Nos := v_Nos || ',' || r_Nos.No;
  End Loop;
  v_Nos := Substr(v_Nos, 2);

  For r_结帐ids In (Select Distinct a.结帐id
                  From 门诊费用记录 A
                  Where a.No In (Select Column_Value From Table(f_Str2List(v_Nos))) And Mod(a.记录性质, 10) = 1 And
                        a.记录状态 <> 0) Loop
    v_结帐ids := v_结帐ids || ',' || r_结帐ids.结帐id;
  End Loop;

  v_结帐ids := Substr(v_结帐ids, 2);
  Select Count(a.No), Sum(a.实收金额)
  Into n_Count, n_实收金额
  From 门诊费用记录 A
  Where a.No In (Select Column_Value From Table(f_Str2List(v_Nos))) And Mod(a.记录性质, 10) = 1;

  If n_Count = 0 Or n_实收金额 = 0 Then
    v_Err_Msg := '本次结算不是收费或因并发原因他人操作了该结算,不能转为住院费用.';
    Raise Err_Item;
  End If;

  Select 结帐id, 病人id, 开单部门id, 开单人
  Into n_原结帐id, n_病人id, n_开单部门id, v_开单人
  From 门诊费用记录
  Where 结帐id = 原结帐id_In And Mod(记录性质, 10) = 1 And 记录状态 In (1, 3) And Rownum < 2;

  Select Nvl(Max(1), 0)
  Into n_部分退费
  From 门诊费用记录 A
  Where Mod(a.记录性质, 10) = 1 And a.记录状态 = 2 And a.结帐id In (Select Column_Value From Table(f_Str2List(v_结帐ids))) And
        Rownum < 2;

  Begin
    Select 0
    Into n_部分退费
    From 门诊费用记录 A
    Where 记录性质 = 11 And a.结帐id In (Select Column_Value From Table(f_Str2List(v_结帐ids))) And Rownum < 2;
  Exception
    When Others Then
      Null;
  End;

  n_退费条数 := 0;
  --只有存在部分退时，发需要检 查结算方式中含了多少条结算信息
  Select Count(*) - Max(Decode(性质, 1, 1, 0)) As 统计条数, Sum(Decode(性质, 1, 1, 0) * 冲预交) As 剩余预交, Sum(冲预交) As 剩余退款,
         Max(是否有医保) As 是否有医保
  Into n_退费条数, n_剩余预交, n_返回值, n_是否存在医保
  From (Select Mod(a.记录性质, 10) As 性质, Decode(Mod(a.记录性质, 10), 1, '冲预交', a.结算方式) As 结算方式, Max(1) As 退费条数, Sum(冲预交) As 冲预交,
                Max(Decode(m.性质, 3, 1, 4, 1, 0)) As 是否有医保
         From 病人预交记录 A, 结算方式 M
         Where a.结算方式 = m.名称(+) And a.记录状态 <> 0 And 结帐id In (Select Column_Value From Table(f_Str2List(v_结帐ids)))
         Group By Mod(a.记录性质, 10), Decode(Mod(a.记录性质, 10), 1, '冲预交', a.结算方式)
         Having Sum(冲预交) <> 0);

  If Round(Nvl(n_实收金额, 0), 5) <> Round(Nvl(n_返回值, 0), 5) Then
    v_Err_Msg := '本次结算的剩余费用未退款与结算信息的未退款信息不符,不能转为住院费用.' || Chr(13) || '费用剩余合计:' ||
                 LTrim(To_Char(n_实收金额, '9999999990.99')) || Chr(13) || '结算剩余合计:' ||
                 LTrim(To_Char(n_返回值, '9999999990.99'));
    Raise Err_Item;
  End If;

  --先进行误差费处理
  n_返回值 := Round(n_实收金额, 2) - Nvl(n_实收金额, 0);
  If Nvl(n_返回值, 0) <> 0 Then
    n_误差费 := Nvl(n_误差费, 0) + n_返回值;
  End If;
  n_实收金额 := Round(n_实收金额, 2);
  --1.1作废费用记录
  n_结帐id := 结帐id_In;
  If n_结帐id Is Null Then
    Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
  End If;

  费用记录_Strict(n_结帐id, v_Nos, n_组id);

  --作废医保(以前是原结帐ID冲销，应该是有问题 ，现调整冲销ID
  医保结算明细_Strict(n_结帐id, v_结帐ids, v_Nos, n_组id, n_预交金额, 1);

  --1.2作废预交记录
  --作废冲预交部分
  If (n_部分退费 = 0 Or n_退费条数 = 0) And Nvl(门诊退费_In, 0) = 0 Then
    --1.预交原样退
    病人预交款_Del(n_结帐id, v_结帐ids, Null, n_返回值);
    n_实收金额 := n_实收金额 - Nvl(n_返回值, 0);
  Elsif n_退费条数 >= 1 And n_剩余预交 >= n_实收金额 And n_剩余预交 <> 0 And Nvl(n_是否存在医保, 0) = 0 Then
    --2.部门退且预交足够:在部分退时，未退预交金，而本次预交金大于了剩余的金额，直接返回预交金
    --预交按指定金额退
    病人预交款_Del(n_结帐id, v_结帐ids, n_实收金额, n_返回值);
    n_实收金额 := 0;
  Elsif n_退费条数 = 1 And n_剩余预交 < n_实收金额 And n_剩余预交 <> 0 And Nvl(n_是否存在医保, 0) = 0 Then
    --3.只有一种结算且部分退:剩余只有一种结算方式且存在预交小于了剩余款,则全退预交金
    病人预交款_Del(n_结帐id, v_结帐ids, Null, n_预交金额);
    n_实收金额 := n_实收金额 - Nvl(n_预交金额, 0);
  Else
    --4.存在多条记录时，需要先退预交进行排除，然后在统计，将这部分转为缺省结算方式
    Select 病人预交记录_Id.Nextval Into n_Tempid From Dual;
    Insert Into 病人预交记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交, 结帐id,
       卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 结算性质, 关联交易id)
      Select n_Tempid, Max(NO), Max(实际票号), 3, 3, 病人id, Max(主页id) As 主页id, Max(科室id) As 科室id, Null, v_缺省结算方式, Max(结算号码),
             '预交临时记录', Null, Null, Null, Max(收款时间), 操作员姓名_In, 操作员编号_In, Sum(冲预交), n_原结帐id, Null, Null, Null, Null, Null,
             Null, -1 * n_原结帐id, 3, Null
      From 病人预交记录 A
      Where 记录性质 In (1, 11) And a.结帐id In (Select Column_Value From Table(f_Str2List(v_结帐ids))) And Nvl(冲预交, 0) <> 0
      Group By 病人id;
  End If;

  --作废门诊缴费及医保部分(含老一卡通):校对标志，统计更改为1
  If n_实收金额 <> 0 Then
  
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别,
       卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 结算性质, 校对标志, 关联交易id)
      Select 病人预交记录_Id.Nextval, 记录性质, NO, 2, 病人id, 主页id, 摘要, 结算方式, 结算号码, 退费时间_In, 缴款单位, 单位开户行, 单位帐号, 操作员编号_In, 操作员姓名_In,
             0, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, -1 * n_结帐id, 结算性质, 1, 关联交易id
      From 病人预交记录 A
      Where a.记录性质 = 3 And a.记录状态 = 1 And a.结帐id In (Select Column_Value From Table(f_Str2List(v_结帐ids))) And
            (a.卡类别id Is Null And a.结算卡序号 Is Null);
    Update 病人预交记录
    Set 记录状态 = 3
    Where 记录性质 = 3 And 记录状态 = 1 And 结帐id In (Select Column_Value From Table(f_Str2List(v_结帐ids)));
  
  End If;

  --2.票据收回
  门诊结算票据_回收(v_Nos);

  --3.缴款数据处理(
  --   现有两种情况:
  --    1. 转出过程直接销帐的,则缴款数据不增加;
  --    2. 先转出,再到门诊退款退票,则需要进行缴款数据处理
  If Nvl(门诊退费_In, 0) = 1 Then
    --结算冲销
    病人结算_Strict(n_结帐id, v_结帐ids, Null, 0);
  Elsif n_实收金额 <> 0 Then
    If n_部分退费 = 0 Then
      n_实收金额 := Null; --不是部分退，则全退
    End If;
    --4.退款转预交(不产生票据,由操作员通过重打进行)
    病人结算_Strict(n_结帐id, v_结帐ids, n_实收金额);
  End If;

  If n_误差费 Is Not Null Then
  
    Update 病人预交记录
    Set 冲预交 = 冲预交 - n_误差费
    Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 = v_缺省结算方式;
  
    Update 病人预交记录
    Set 冲预交 = 冲预交 + n_误差费
    Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 = v_误差费;
  
    If Sql%RowCount = 0 Then
      Select 病人预交记录_Id.Nextval Into n_Tempid From Dual;
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要, 缴款组id,
         卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质, 关联交易id)
      Values
        (n_Tempid, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, n_误差费, v_误差费, Null, 退费时间_In, Null, Null, Null,
         操作员编号_In, 操作员姓名_In, '', n_组id, Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 0, 3, n_Tempid);
    End If;
  End If;

  门诊转住费用_Over(n_结帐id, n_原结帐id);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊转住院_收费转出;
/
Create Or Replace Procedure Zl_门诊转住院_补结算转出
(
  No_In           费用补充记录.No%Type,
  费用冲销id_In   病人预交记录.结帐id%Type,
  结算冲销id_In   病人预交记录.结帐id%Type,
  结算序号_In     病人预交记录.结算序号%Type,
  退费时间_In     住院费用记录.发生时间%Type,
  操作员编号_In   住院费用记录.操作员编号%Type,
  操作员姓名_In   住院费用记录.操作员姓名%Type,
  主页id_In       病人预交记录.主页id%Type,
  入院科室id_In   病人预交记录.科室id%Type,
  结算方式_In     病人预交记录.结算方式%Type := Null,
  误差费_In       病人预交记录.冲预交%Type := Null,
  预交电子票据_In 病人预交记录.预交电子票据%Type := 0
) As
  --功能：对费用补充结算的门诊费用进行转住院费用处理 
  --入参： 
  --  结算方式_In 不为空，表示所有除预交款的非医保金额全部退为指定的结算方式； 
  --              为空，表示所有除预交款的非医保金额全部转为住院预交款  
  --  预交电子票据_In:预交款是否启用电子票据
  Err_Item Exception;
  v_Err_Msg Varchar2(200);
  n_返回值  病人预交记录.冲预交%Type;

  n_组id     财务缴款分组.Id%Type;
  v_误差费   结算方式.名称%Type;
  n_误差费   病人预交记录.冲预交%Type;
  n_Dec      Number; --金额小数位数 
  n_异步结算 Number;

  v_Nos    Varchar2(4000);
  n_病人id 病人预交记录.病人id%Type;

  n_已退金额 病人预交记录.冲预交%Type;
  n_未退金额 病人预交记录.冲预交%Type;
  n_冲预交   病人预交记录.冲预交%Type;
  v_结算方式 Varchar2(4000);
  v_预交no   病人预交记录.No%Type;

  n_是否电子票据 病人预交记录.是否电子票据%Type;

  --保存预交款单据 
  Procedure 病人预交记录_Insert
  (
    病人id_In     病人预交记录.病人id%Type,
    金额_In       病人预交记录.金额%Type,
    结算方式_In   病人预交记录.结算方式%Type,
    收款时间_In   病人预交记录.收款时间%Type,
    结算号码_In   病人预交记录.结算号码%Type,
    卡类别id_In   病人预交记录.卡类别id%Type := Null,
    卡号_In       病人预交记录.卡号%Type := Null,
    交易流水号_In 病人预交记录.交易流水号%Type := Null,
    交易说明_In   病人预交记录.交易说明%Type := Null,
    关联交易id_In 病人预交记录.关联交易id%Type := Null
  ) As
    n_充值id 病人预交记录.Id%Type;
    v_预交no 病人预交记录.No%Type;
    n_返回值 病人预交记录.金额%Type;
  Begin
    If Nvl(金额_In, 0) = 0 Or 结算方式_In Is Null Then
      Return;
    End If;
  
    --一卡通，每一笔都生成一条预交款记录 
    --其它，同一种结算方式只生成一条预交款记录 
    Update 病人预交记录
    Set 金额 = Nvl(金额, 0) + 金额_In
    Where 记录性质 = 1 And 记录状态 = 1 And 收款时间 = 收款时间_In And 病人id + 0 = 病人id_In And 结算方式 = 结算方式_In And
          (Nvl(卡类别id, 0) = 0 And Nvl(卡类别id_In, 0) = 0)
    Returning ID Into n_充值id;
    If Sql%RowCount = 0 Then
      v_预交no := Nextno(11);
      Select 病人预交记录_Id.Nextval Into n_充值id From Dual;
    
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要, 缴款组id, 预交类别,
         卡类别id, 卡号, 交易说明, 交易流水号, 结算号码, 关联交易id, 交易人员, 交易时间, 预交电子票据)
      Values
        (n_充值id, v_预交no, Null, 1, 1, 病人id_In, 主页id_In, 入院科室id_In, 金额_In, 结算方式_In, 收款时间_In, Null, Null, Null, 操作员编号_In,
         操作员姓名_In, '门诊转住院预交', n_组id, 2, 卡类别id_In, 卡号_In, 交易说明_In, 交易流水号_In, 结算号码_In, Nvl(关联交易id_In, n_充值id), 操作员姓名_In,
         收款时间_In, 预交电子票据_In);
    End If;
  
    Update 预交单据余额
    Set 预交余额 = Nvl(预交余额, 0) + 金额_In
    Where 病人id = 病人id_In And 预交id = n_充值id
    Returning 预交余额 Into n_返回值;
    If Sql%RowCount = 0 Then
      Insert Into 预交单据余额 (预交id, 病人id, 预交类别, 预交余额) Values (n_充值id, 病人id_In, 2, 金额_In);
      n_返回值 := 金额_In;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 预交单据余额 Where 预交id = n_充值id And Nvl(预交余额, 0) = 0;
    End If;
  
    Update 病人余额
    Set 预交余额 = Nvl(预交余额, 0) + 金额_In
    Where 性质 = 1 And 病人id = 病人id_In And 类型 = 2
    Returning 预交余额 Into n_返回值;
    If Sql%RowCount = 0 Then
      Insert Into 病人余额 (病人id, 性质, 类型, 预交余额, 费用余额) Values (病人id_In, 1, 2, 金额_In, 0);
      n_返回值 := 金额_In;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 病人余额 Where 病人id = 病人id_In And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
    End If;
  End;
Begin
  n_组id := Zl_Get组id(操作员姓名_In);
  --误差费 
  Begin
    Select 名称 Into v_误差费 From 结算方式 Where 性质 = 9 And Rownum < 2;
  Exception
    When Others Then
      v_Err_Msg := '没有发现误差结算方式，请检查是否正确设置！';
      Raise Err_Item;
  End;
  n_误差费 := Nvl(误差费_In, 0);

  --金额小数位数 
  Select zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  Select f_List2Str(Cast(Collect(a.No) As t_StrList), ',', 1), Max(a.病人id)
  Into v_Nos, n_病人id
  From 门诊费用记录 A, 费用补充记录 B
  Where a.结帐id = b.收费结帐id And b.记录性质 = 1 And b.附加标志 = 0 And b.No = No_In;
  If v_Nos Is Null Then
    v_Err_Msg := '未找到原医保补结算数据，费用转出失败!';
    Raise Err_Item;
  End If;

  --1.更新费用审核记录 
  Update 费用审核记录
  Set 记录状态 = 2
  Where 性质 = 1 And 费用id In (Select /*+cardinality(b,10)*/
                             a.Id
                            From 门诊费用记录 A, (Select Column_Value As NO From Table(f_Str2List(v_Nos))) B
                            Where a.No = b.No And Mod(a.记录性质, 10) = 1 And a.记录状态 In (1, 3));

  --2.作废门诊费用记录 
  Update 门诊费用记录
  Set 记录状态 = 3
  Where Mod(记录性质, 10) = 1 And 记录状态 = 1 And NO In (Select Column_Value As NO From Table(f_Str2List(v_Nos)));

  --根据原收费记录是否生成对照数据来确定作废记录是否也生成对照数据 
  Select /*+cardinality(c,10)*/
   Count(1)
  Into n_异步结算
  From 费用结算对照 A, 门诊费用记录 B, (Select Column_Value As NO From Table(f_Str2List(v_Nos))) C
  Where a.费用id = b.Id And a.门诊标志 = 1 And b.记录性质 = 1 And b.No = c.No And Rownum < 2;

  For c_费用 In (Select /*+cardinality(b,10)*/
                a.No, a.实际票号, a.序号, a.从属父号, a.价格父号, a.病人id, a.医嘱序号, a.门诊标志, a.姓名, a.性别, a.年龄, a.标识号, a.付款方式, a.病人科室id,
                a.费别, a.收费类别, a.收费细目id, a.计算单位, a.发药窗口, Sum(Nvl(a.付数, 1) * a.数次) As 数次, a.加班标志, a.附加标志, a.婴儿费, a.收入项目id,
                a.收据费目, a.标准单价, Sum(a.应收金额) As 应收金额, Sum(a.实收金额) As 实收金额, a.划价人, a.开单部门id, a.开单人, a.发生时间, a.执行部门id, a.执行人,
                Nvl(Min(Decode(a.记录状态, 2, a.执行状态, 0)), 0) - 1 As 执行状态, a.结论, Sum(a.结帐金额) As 结帐金额, Max(保险大类id) As 保险大类id,
                Max(保险项目否) As 保险项目否, Max(保险编码) As 保险编码, Max(费用类型) As 费用类型, Sum(a.统筹金额) As 统筹金额, 是否急诊, a.挂号id, a.主页id,
                a.病人病区id
               From 门诊费用记录 A, (Select Column_Value As NO From Table(f_Str2List(v_Nos))) B
               Where a.No = b.No And a.记录性质 In (1, 11)
               Group By a.No, a.实际票号, a.序号, a.从属父号, a.价格父号, a.病人id, a.医嘱序号, a.门诊标志, a.姓名, a.性别, a.年龄, a.标识号, a.付款方式,
                        a.病人科室id, a.费别, a.收费类别, a.收费细目id, a.计算单位, a.发药窗口, a.加班标志, a.附加标志, a.婴儿费, a.收入项目id, a.收据费目, a.标准单价,
                        a.划价人, a.开单部门id, a.开单人, a.发生时间, a.执行部门id, a.执行人, a.结论, 是否急诊, a.挂号id, a.主页id, a.病人病区id
               Having Nvl(Sum(Nvl(a.付数, 1) * a.数次), 0) <> 0) Loop
  
    Insert Into 门诊费用记录
      (ID, 记录性质, NO, 记录状态, 实际票号, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 病人科室id, 费别, 收费类别, 收费细目id,
       计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人,
       执行状态, 执行时间, 结论, 操作员编号, 操作员姓名, 结帐id, 结帐金额, 保险大类id, 保险项目否, 保险编码, 费用类型, 统筹金额, 摘要, 是否急诊, 缴款组id, 费用状态, 挂号id, 主页id,
       病人病区id)
    Values
      (病人费用记录_Id.Nextval, 1, c_费用.No, 2, c_费用.实际票号, c_费用.序号, c_费用.从属父号, c_费用.价格父号, c_费用.病人id, c_费用.医嘱序号, c_费用.门诊标志,
       c_费用.姓名, c_费用.性别, c_费用.年龄, c_费用.标识号, c_费用.付款方式, c_费用.病人科室id, c_费用.费别, c_费用.收费类别, c_费用.收费细目id, c_费用.计算单位, 1,
       c_费用.发药窗口, -1 * c_费用.数次, c_费用.加班标志, c_费用.附加标志, c_费用.婴儿费, c_费用.收入项目id, c_费用.收据费目, c_费用.标准单价, -1 * c_费用.应收金额,
       -1 * c_费用.实收金额, c_费用.划价人, c_费用.开单部门id, c_费用.开单人, c_费用.发生时间, 退费时间_In, c_费用.执行部门id, c_费用.执行人, c_费用.执行状态, Null,
       c_费用.结论, 操作员编号_In, 操作员姓名_In, 费用冲销id_In, -1 * c_费用.结帐金额, c_费用.保险大类id, c_费用.保险项目否, c_费用.保险编码, c_费用.费用类型,
       -1 * c_费用.统筹金额, '', c_费用.是否急诊, n_组id, 0, c_费用.挂号id, c_费用.主页id, c_费用.病人病区id);
  
    If n_异步结算 = 1 Then
      Insert Into 费用结算对照
        (门诊标志, 费用id, 是否重收, 结帐id, 结帐金额, 操作员编号, 操作员姓名)
      Values
        (1, 病人费用记录_Id.Currval, 0, 费用冲销id_In, -1 * c_费用.结帐金额, 操作员编号_In, 操作员姓名_In);
    End If;
  End Loop;
  Zl_门诊退费结算_Modify(1, n_病人id, 费用冲销id_In, Null);

  --3.作废补充结算记录（同时已进行了票据回收和医保原样退） 
  Zl_费用补充记录_Delete(No_In, 结算冲销id_In, Null, 结算序号_In, 费用冲销id_In, 操作员编号_In, 操作员姓名_In, 退费时间_In);
  Update 费用补充记录 Set 费用状态 = 0 Where 结算序号 = 结算序号_In;
  --处理为医保接口已调用成功 
  Update 病人预交记录
  Set 校对标志 = 2
  Where 记录性质 = 6 And 结帐id = 结算冲销id_In And 结算方式 In (Select 名称 From 结算方式 Where 性质 In (3, 4));

  --4.结算数据处理 
  Select -1 * Nvl(Sum(a.冲预交), 0)
  Into n_未退金额
  From 病人预交记录 A
  Where a.结算序号 = 结算序号_In And a.结算方式 Is Null;
  If Nvl(n_误差费, 0) = 0 Then
    n_误差费 := Round(n_未退金额, n_Dec) - n_未退金额;
  End If;
  n_未退金额 := n_未退金额 - n_误差费;

  For r_预交 In (Select Case
                        When Mod(a.记录性质, 10) = 1 Then
                         1
                        When Nvl(a.卡类别id, 0) <> 0 Then
                         2
                        Else
                         0
                      End As 类型, a.结帐id, Nvl(a.冲预交, 0) As 冲预交, a.No, a.病人id, a.结算方式, a.结算号码, a.卡类别id, a.卡号, a.交易流水号,
                      a.交易说明, a.关联交易id
               From 病人预交记录 A, 结算方式 B
               Where a.结算方式 = b.名称 And a.记录状态 In (1, 3) And b.性质 Not In (3, 4, 9) And
                     a.结帐id In (Select 收费结帐id From 费用补充记录 Where 记录性质 = 1 And 附加标志 = 0 And NO = No_In)) Loop
  
    --都是单种结算方式 
    If r_预交.类型 = 1 Then
      --预交款 
      Zl_费用补充结算_完成退费(结算冲销id_In, Null, Null, Null, Null, Null, n_误差费, 0, -1 * n_未退金额);
      Exit;
    Elsif r_预交.类型 = 2 Then
      --一卡通 
      Select Nvl(Sum(金额), 0) Into n_已退金额 From 三方退款信息 Where 记录id = r_预交.结帐id;
      If r_预交.冲预交 - n_已退金额 > 0 Then
        If r_预交.冲预交 - n_已退金额 > n_未退金额 Then
          n_冲预交 := n_未退金额;
        Else
          n_冲预交 := r_预交.冲预交 - n_已退金额;
        End If;
      
        v_结算方式 := r_预交.结算方式 || '|' || -1 * n_冲预交 || '| | ';
        Zl_费用补充结算_完成退费(结算冲销id_In, v_结算方式, r_预交.卡类别id, r_预交.卡号, r_预交.交易流水号, r_预交.交易说明, n_误差费, 0, 0, 2);
        Zl_三方退款信息_Insert(结算序号_In, r_预交.结帐id, n_冲预交, r_预交.卡号, r_预交.交易流水号, r_预交.交易说明, 0, 0, 0, r_预交.卡类别id, r_预交.交易流水号,
                         r_预交.交易说明);
      
        --转为住院预交款 
        病人预交记录_Insert(r_预交.病人id, n_冲预交, r_预交.结算方式, 退费时间_In, r_预交.结算号码, r_预交.卡类别id, r_预交.卡号, r_预交.交易流水号, r_预交.交易说明,
                      r_预交.关联交易id);
      
        n_未退金额 := n_未退金额 - n_冲预交;
        n_误差费   := 0;
      End If;
      If n_未退金额 = 0 Then
        Exit;
      End If;
    Else
      --其它非医保结算方式 
      --结算方式|结算金额|结算号码|结算摘要 
      v_结算方式 := r_预交.结算方式 || '|' || -1 * n_未退金额 || '| | ';
      Zl_费用补充结算_完成退费(结算冲销id_In, v_结算方式, Null, Null, Null, Null, n_误差费, 0);
    
      --转为住院预交款 
      病人预交记录_Insert(r_预交.病人id, n_未退金额, r_预交.结算方式, 退费时间_In, r_预交.结算号码);
      Exit;
    End If;
  End Loop;

  --5.转出完成处理 
  Delete From 病人预交记录 Where 结帐id = 结算冲销id_In And 结算方式 Is Null And Nvl(冲预交, 0) = 0;
  If Sql%NotFound Then
    v_Err_Msg := '还存在未缴款的数据，不能完成结算！';
    Raise Err_Item;
  End If;
  Delete From 病人预交记录 Where 结帐id = 费用冲销id_In And 结算方式 Is Null And Nvl(冲预交, 0) = 0;
  If Sql%NotFound Then
    v_Err_Msg := '还存在未缴款的数据，不能完成结算！';
    Raise Err_Item;
  End If;
  Update 病人预交记录 Set 校对标志 = 0, 会话号 = Null Where 结算序号 = 结算序号_In;

  --6.更新 是否电子票据 标记 
  Select Max(a.是否电子票据)
  Into n_是否电子票据
  From 病人预交记录 A,
       (Select 结算id
         From (Select b.结算id
                From 费用补充记录 B
                Where b.No = No_In And b.记录性质 = 1 And b.记录状态 In (1, 3)
                Order By b.登记时间)
         Where Rownum < 2) B
  Where a.结帐id = b.结算id;

  Update 病人预交记录 Set 是否电子票据 = n_是否电子票据 Where 结算序号 = 结算序号_In And 结算性质 = 6;

  --人员缴款余额（主要是医保） 
  For c_预交 In (Select a.结算方式, a.操作员姓名, Nvl(Sum(a.冲预交), 0) As 冲预交
               From 病人预交记录 A, 结算方式 B
               Where a.结算方式 = b.名称 And b.性质 In (3, 4) And a.结算序号 = 结算序号_In
               Group By a.结算方式, a.操作员姓名) Loop
  
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) + c_预交.冲预交
    Where 收款员 = c_预交.操作员姓名 And 性质 = 1 And 结算方式 = c_预交.结算方式
    Returning 余额 Into n_返回值;
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额
        (收款员, 结算方式, 性质, 余额)
      Values
        (c_预交.操作员姓名, c_预交.结算方式, 1, c_预交.冲预交);
      n_返回值 := c_预交.冲预交;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 人员缴款余额
      Where 收款员 = c_预交.操作员姓名 And 性质 = 1 And 结算方式 = c_预交.结算方式 And Nvl(余额, 0) = 0;
    End If;
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊转住院_补结算转出;
/
Create Or Replace Procedure Zl_门诊转住院_结帐作废
(
  No_In           病人结帐记录.No%Type,
  冲销id_In       病人结帐记录.Id%Type,
  主页id_In       病人预交记录.主页id%Type,
  入院科室id_In   病人预交记录.科室id%Type,
  完成作废_In     Number := 0,
  预交电子票据_In 病人预交记录.预交电子票据%Type := 0
) As
  --功能：门诊费用转住院作废结帐结算数据，立即销账模式调用 
  --入参： 
  --  完成作废_In:0-开始结帐作废;1-完成结帐作废 
  --  预交电子票据_In:预交款是否启用电子票据，完成作废时传入
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_返回值   病人预交记录.冲预交%Type;
  v_结算方式 病人预交记录.结算方式%Type;
  n_结算金额 病人预交记录.冲预交%Type;
  n_原结帐id 病人结帐记录.Id%Type;
  n_结帐金额 病人结帐记录.结帐金额%Type;

  n_冲预交       病人预交记录.冲预交%Type;
  n_是否转为预交 Number(2);
  n_预交id       病人预交记录.Id%Type;
  n_误差费       病人预交记录.冲预交%Type;
  n_存在退支票   Number(2);

  Cursor c_Balance_Data Is
    Select NO, 病人id, 科室id, 主页id, 冲预交, 收款时间, 操作员编号, 操作员姓名, 缴款组id
    From 病人预交记录
    Where 结帐id = 冲销id_In And 结算方式 Is Null;
  r_Balance_Data c_Balance_Data%RowType;

  Procedure 人员缴款余额_Update
  (
    收款员_In   人员缴款余额.收款员%Type,
    结算方式_In 人员缴款余额.结算方式%Type,
    金额_In     人员缴款余额.余额%Type
  ) As
    --功能：更新 人员缴款余额 
    n_返回值 人员缴款余额.余额%Type;
  Begin
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) + Nvl(金额_In, 0)
    Where 收款员 = 收款员_In And 性质 = 1 And 结算方式 = 结算方式_In
    Returning 余额 Into n_返回值;
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (收款员_In, 结算方式_In, 1, Nvl(金额_In, 0));
      n_返回值 := Nvl(金额_In, 0);
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 人员缴款余额 Where 收款员 = 收款员_In And 性质 = 1 And 结算方式 = 结算方式_In And Nvl(余额, 0) = 0;
    End If;
  End 人员缴款余额_Update;

  Procedure 住院预交款_Insert
  (
    病人id_In     病人预交记录.病人id%Type,
    金额_In       病人预交记录.金额%Type,
    结算方式_In   病人预交记录.结算方式%Type,
    收款时间_In   病人预交记录.收款时间%Type,
    结算号码_In   病人预交记录.结算号码%Type,
    操作员编号_In 病人预交记录.操作员编号%Type,
    操作员姓名_In 病人预交记录.操作员姓名%Type,
    缴款组id_In   病人预交记录.缴款组id%Type,
    卡类别id_In   病人预交记录.卡类别id%Type := Null,
    卡号_In       病人预交记录.卡号%Type := Null,
    交易流水号_In 病人预交记录.交易流水号%Type := Null,
    交易说明_In   病人预交记录.交易说明%Type := Null,
    关联交易id_In 病人预交记录.关联交易id%Type := Null,
    交易人员_In   病人预交记录.交易人员%Type := Null,
    交易时间_In   病人预交记录.交易时间%Type := Null
  ) As
    --功能：新增 住院预交款 
    n_预交id 病人预交记录.Id%Type;
    v_预交no 病人预交记录.No%Type;
    n_返回值 病人预交记录.金额%Type;
  Begin
    If Nvl(金额_In, 0) = 0 Or 结算方式_In Is Null Then
      Return;
    End If;
  
    --一卡通，每一笔都生成一条预交款记录 
    --其它，同一种结算方式只生成一条预交款记录 
    Update 病人预交记录
    Set 金额 = Nvl(金额, 0) + 金额_In
    Where 记录性质 = 1 And 记录状态 = 1 And 收款时间 = 收款时间_In And 病人id + 0 = 病人id_In And 结算方式 = 结算方式_In And
          (Nvl(卡类别id, 0) = 0 And Nvl(卡类别id_In, 0) = 0)
    Returning ID Into n_预交id;
    If Sql%RowCount = 0 Then
      v_预交no := Nextno(11);
      Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
    
      Insert Into 病人预交记录
        (ID, NO, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 收款时间, 操作员编号, 操作员姓名, 摘要, 缴款组id, 预交类别, 卡类别id, 卡号, 交易说明, 交易流水号,
         结算号码, 关联交易id, 交易人员, 交易时间, 预交电子票据)
      Values
        (n_预交id, v_预交no, 1, 1, 病人id_In, 主页id_In, 入院科室id_In, 金额_In, 结算方式_In, 收款时间_In, 操作员编号_In, 操作员姓名_In, '门诊转住院预交',
         缴款组id_In, 2, 卡类别id_In, 卡号_In, 交易说明_In, 交易流水号_In, 结算号码_In, Nvl(关联交易id_In, n_预交id), Nvl(交易人员_In, 操作员姓名_In),
         Nvl(交易时间_In, 收款时间_In), 预交电子票据_In);
    End If;
  
    Update 预交单据余额
    Set 预交余额 = Nvl(预交余额, 0) + 金额_In
    Where 病人id = 病人id_In And 预交id = n_预交id
    Returning 预交余额 Into n_返回值;
    If Sql%RowCount = 0 Then
      Insert Into 预交单据余额 (预交id, 病人id, 预交类别, 预交余额) Values (n_预交id, 病人id_In, 2, 金额_In);
      n_返回值 := 金额_In;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 预交单据余额 Where 预交id = n_预交id And Nvl(预交余额, 0) = 0;
    End If;
  
    Update 病人余额
    Set 预交余额 = Nvl(预交余额, 0) + 金额_In
    Where 性质 = 1 And 病人id = 病人id_In And 类型 = 2
    Returning 预交余额 Into n_返回值;
    If Sql%RowCount = 0 Then
      Insert Into 病人余额 (病人id, 性质, 类型, 预交余额, 费用余额) Values (病人id_In, 1, 2, 金额_In, 0);
      n_返回值 := 金额_In;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 病人余额 Where 病人id = 病人id_In And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
    End If;
  
    --更新人员缴款余额 
    人员缴款余额_Update(操作员姓名_In, 结算方式_In, 金额_In);
  End 住院预交款_Insert;
Begin
  Open c_Balance_Data;
  Fetch c_Balance_Data
    Into r_Balance_Data;
  If r_Balance_Data.No Is Null Then
    v_Err_Msg := '未找到指定的结帐作废结算数据！';
    Raise Err_Item;
  End If;

  Select Max(ID), Nvl(Sum(结帐金额), 0)
  Into n_原结帐id, n_结帐金额
  From 病人结帐记录
  Where 记录状态 In (1, 3) And NO = No_In;
  If Nvl(n_原结帐id, 0) = 0 Then
    v_Err_Msg := '没有发现要作废的结帐单据，可能已经作废！';
    Raise Err_Item;
  End If;

  If Nvl(完成作废_In, 0) = 0 Then
    --类型：0-普通结算;1-预交款;2-医保,3-一卡通,4-一卡通(老),5-消费卡 
    --按原样作废处理
    For r_Pay In (Select Case
                            When Mod(a.记录性质, 10) = 1 Then
                             1
                            When Instr(',3,4,', ',' || b.性质 || ',') > 0 And a.卡类别id Is Null Then
                             2
                            When Nvl(a.卡类别id, 0) <> 0 Then
                             3
                            When j.结算方式 Is Not Null Then
                             4
                            When a.结算卡序号 Is Not Null Then
                             5
                            Else
                             0
                          End As 类型, a.Id, a.No, a.病人id, a.科室id, a.主页id, a.结算方式,
                         Nvl(Decode(a.结算卡序号, Null, a.冲预交, -1 * p.应收金额), 0) As 冲预交, a.结算号码, a.卡类别id,
                         Decode(a.结算卡序号, Null, a.卡号, p.卡号) As 卡号, a.交易流水号, a.交易说明, a.合作单位, a.结算卡序号, p.消费卡id,
                         Nvl(b.性质, 1) As 结算性质, a.关联交易id, a.交易人员, a.交易时间, b.应付款
                  From 病人预交记录 A, 结算方式 B, 一卡通目录 J, 病人卡结算记录 P
                  Where a.结算方式 = b.名称(+) And a.结算方式 = j.结算方式(+) And a.Id = p.结算id(+) And a.结算卡序号 = p.接口编号(+) And
                        a.结帐id = n_原结帐id
                  Order By 冲预交) Loop
      n_冲预交 := Nvl(r_Pay.冲预交, 0);
    
      --1-预交款,在完成时再处理 
      --2-医保,程序中处理 
      If r_Pay.类型 = 1 Or r_Pay.类型 = 2 Then
        n_是否转为预交 := 0;
      Else
        n_是否转为预交 := 1;
      End If;
    
      --3-一卡通 
      If r_Pay.类型 = 3 Then
        --需要检查是否多种结算方式或含医保，如果是，则需要调用接口退款 
        Select Count(Distinct 结算方式) + Max(Decode(b.性质, 3, 2, 4, 2, 0))
        Into n_是否转为预交
        From 病人预交记录 A, 结算方式 B
        Where a.结算方式 = b.名称(+) And 结帐id = n_原结帐id And 卡类别id = r_Pay.卡类别id And Nvl(关联交易id, 0) = Nvl(r_Pay.关联交易id, 0);
      
        If Nvl(n_是否转为预交, 0) = 1 And n_冲预交 < 0 Then
          --单种结算方式且不是医保，同时结帐时是退款/转账，则不处理 
          n_是否转为预交 := 0;
        End If;
      End If;
    
      --4-一卡通(老),原样退回 
    
      --5-消费卡,原样退回 
      If r_Pay.类型 = 5 Then
        Update 病人预交记录
        Set 冲预交 = Nvl(冲预交, 0) - n_冲预交
        Where 结帐id = 冲销id_In And 结算方式 = r_Pay.结算方式 And 结算卡序号 = r_Pay.结算卡序号
        Returning ID Into n_预交id;
        If Sql%NotFound Then
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (ID, 记录性质, NO, 记录状态, 病人id, 科室id, 主页id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算卡序号, 校对标志, 结算性质)
          Values
            (n_预交id, 12, r_Pay.No, 1, r_Pay.病人id, r_Pay.科室id, r_Pay.主页id, r_Pay.结算方式, r_Balance_Data.收款时间,
             r_Balance_Data.操作员编号, r_Balance_Data.操作员姓名, -1 * n_冲预交, 冲销id_In, r_Balance_Data.缴款组id, r_Pay.结算卡序号, 2, 2);
        End If;
      
        --插入卡结算记录 
        Zl_病人卡结算记录_退款(r_Pay.结算卡序号, r_Pay.卡号, r_Pay.消费卡id, n_冲预交, r_Pay.Id, n_预交id, r_Balance_Data.操作员编号,
                      r_Balance_Data.操作员姓名, r_Balance_Data.收款时间);
      
        Update 病人预交记录 Set 冲预交 = 冲预交 + n_冲预交 Where 结帐id = 冲销id_In And 结算方式 Is Null;
        n_是否转为预交 := 0;
      End If;
    
      --0-普通结算,转为住院预交 
      If r_Pay.类型 = 0 Then
        If Nvl(r_Pay.结算性质, 1) = 1 Or Nvl(r_Pay.结算性质, 1) = 9 Or Nvl(r_Pay.应付款, 0) = 1 Then
          --现金和误差费、应付款(退支票)先不退款,在完成时再转为住院预交 
          n_是否转为预交 := 0;
        End If;
        If Nvl(r_Pay.结算性质, 1) = 2 And Instr(r_Pay.结算方式, '支票') > 0 Then
          --存在退支票的直接按退现转为住院预交 
          Select Count(1)
          Into n_存在退支票
          From 病人预交记录 A, 结算方式 B
          Where a.结算方式 = b.名称 And 结帐id = n_原结帐id And Nvl(b.应付款, 0) = 1 And Rownum < 2;
          If Nvl(n_存在退支票, 0) > 0 Then
            n_是否转为预交 := 0;
          End If;
        End If;
      End If;
    
      If n_是否转为预交 > 0 Then
        Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 操作员编号, 操作员姓名, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号,
           交易说明, 合作单位, 结帐id, 校对标志, 结算性质, 关联交易id, 交易人员, 交易时间)
        Values
          (n_预交id, 12, r_Pay.No, 1, r_Pay.病人id, r_Pay.主页id, r_Pay.科室id, -1 * n_冲预交, r_Pay.结算方式, r_Pay.结算号码,
           r_Balance_Data.收款时间, r_Balance_Data.操作员编号, r_Balance_Data.操作员姓名, r_Balance_Data.缴款组id, r_Pay.卡类别id,
           r_Pay.结算卡序号, r_Pay.卡号, r_Pay.交易流水号, r_Pay.交易说明, r_Pay.合作单位, 冲销id_In, Decode(n_是否转为预交, 1, 2, 1), 2,
           r_Pay.关联交易id, r_Pay.交易人员, r_Pay.交易时间);
      
        --转为住院预交 
        If n_是否转为预交 = 1 Then
          住院预交款_Insert(r_Pay.病人id, n_冲预交, r_Pay.结算方式, r_Balance_Data.收款时间, r_Pay.结算号码, r_Balance_Data.操作员编号,
                       r_Balance_Data.操作员姓名, r_Balance_Data.缴款组id, r_Pay.卡类别id, r_Pay.卡号, r_Pay.交易流水号, r_Pay.交易说明,
                       r_Pay.关联交易id, r_Pay.交易人员, r_Pay.交易时间);
        End If;
      
        Update 病人预交记录 Set 冲预交 = 冲预交 + n_冲预交 Where 结帐id = 冲销id_In And 结算方式 Is Null;
      End If;
    End Loop;
    Return;
  End If;

  -------------------------------------------------------------------------------------------------------------- 
  --完成门诊费用转住院结帐作废 
  Select -1 * Nvl(冲预交, 0) Into n_结算金额 From 病人预交记录 Where 结帐id = 冲销id_In And 结算方式 Is Null;

  --1.按剩余结算金额退预交款 
  If Nvl(n_结算金额, 0) > 0 Then
    For r_预交 In (Select NO, 实际票号, 记录状态, 病人id, 主页id, 科室id, 结算方式, Max(结算号码) As 结算号码, Max(摘要) As 摘要, Max(缴款单位) As 缴款单位,
                        Max(单位开户行) As 单位开户行, Max(单位帐号) As 单位帐号, Sum(冲预交) As 冲预交, Max(预交类别) As 预交类别, 卡类别id, 结算卡序号,
                        Max(卡号) As 卡号, Max(关联交易id) As 关联交易id, Max(交易流水号) As 交易流水号, Max(交易说明) As 交易说明, Max(合作单位) As 合作单位,
                        Nvl(Max(是否转帐及代扣), 0) As 是否转帐及代扣, Max(预交id) As 预交id, Max(交易时间) As 交易时间, Max(交易人员) As 交易人员
                 From (Select a.No, a.实际票号, a.记录状态, 病人id, 主页id, 科室id, a.结算方式, a.结算号码, a.摘要, a.缴款单位, a.单位开户行, a.单位帐号, a.冲预交,
                               a.预交类别, a.卡类别id, a.结算卡序号, a.关联交易id, a.卡号, a.交易流水号, a.交易说明, a.合作单位, b.是否转帐及代扣,
                               Decode(a.记录性质, 1, Decode(a.记录状态, 2, 0, a.Id), 0) As 预交id, a.交易时间 As 交易时间, a.交易人员 As 交易人员
                        From 病人预交记录 A, 医疗卡类别 B
                        Where a.结帐id = n_原结帐id And a.记录性质 In (1, 11) And Nvl(a.冲预交, 0) <> 0 And a.卡类别id = b.Id(+)
                        Union All
                        Select a.No, a.实际票号, a.记录状态, a.病人id, 主页id, a.科室id, a.结算方式, '' || 结算号码 As 结算号码, '' As 摘要, '' As 缴款单位,
                               '' As 单位开户行, '' As 单位帐号, -1 * b.金额 As 冲预交, a.预交类别, a.卡类别id, a.结算卡序号, a.关联交易id, '' As 卡号,
                               '' As 交易流水号, '' As 交易说明, '' As 合作单位, 0 As 是否转帐及代扣,
                               Decode(a.记录性质, 1, Decode(a.记录状态, 2, 0, a.Id), 0) As 预交id, a.交易时间 As 交易时间, a.交易人员 As 交易人员
                        From 病人预交记录 A, 三方退款信息 B
                        Where b.结帐id = n_原结帐id And a.Id = b.记录id And Nvl(b.是否未退, 0) <> 1)
                 Group By NO, 实际票号, 记录状态, 病人id, 主页id, 科室id, 结算方式, 卡类别id, 结算卡序号
                 Having Nvl(Sum(冲预交), 0) <> 0
                 Order By 预交类别 Desc, 是否转帐及代扣 Desc, 交易时间 Desc) Loop
      n_冲预交 := Nvl(r_预交.冲预交, 0);
    
      If n_结算金额 > n_冲预交 Then
        n_结算金额 := Round(n_结算金额 - n_冲预交, 6);
      Else
        n_冲预交   := Nvl(n_结算金额, 0);
        n_结算金额 := 0;
      End If;
    
      Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 科室id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 校对标志, 卡类别id, 结算卡序号, 卡号,
         交易流水号, 交易说明, 结算号码, 结算性质, 关联交易id, 交易时间, 交易人员)
      Values
        (n_预交id, 12, r_Balance_Data.No, 1, r_Balance_Data.病人id, r_Balance_Data.科室id, r_Balance_Data.主页id, Null,
         r_预交.结算方式, r_Balance_Data.收款时间, r_Balance_Data.操作员编号, r_Balance_Data.操作员姓名, -1 * n_冲预交, 冲销id_In,
         r_Balance_Data.缴款组id, 2, r_预交.卡类别id, r_预交.结算卡序号, r_预交.卡号, r_预交.交易流水号, r_预交.交易说明, r_预交.结算号码, 2, r_预交.关联交易id,
         r_预交.交易时间, r_预交.交易人员);
    
      --转为住院预交 
      住院预交款_Insert(r_预交.病人id, n_冲预交, r_预交.结算方式, r_Balance_Data.收款时间, r_预交.结算号码, r_Balance_Data.操作员编号,
                   r_Balance_Data.操作员姓名, r_Balance_Data.缴款组id, r_预交.卡类别id, r_预交.卡号, r_预交.交易流水号, r_预交.交易说明, r_预交.关联交易id,
                   r_预交.交易人员, r_预交.交易时间);
    
      Update 病人预交记录 Set 冲预交 = 冲预交 + n_冲预交 Where 结帐id = 冲销id_In And 结算方式 Is Null;
    
      If n_结算金额 = 0 Then
        Exit;
      End If;
    End Loop;
  End If;

  n_结算金额 := -1 * n_结算金额;
  --2.将未退金额全部按"现金"转为住院预交进行退款 
  n_冲预交 := Zl_Cent_Money(n_结算金额);
  n_误差费 := Round(n_结算金额 - n_冲预交, 6);
  If n_冲预交 <> 0 Then
    Select Nvl(Max(名称), '现金') Into v_结算方式 From 结算方式 Where Nvl(性质, 0) = 1 And Rownum < 1;
  
    Update 病人预交记录
    Set 冲预交 = Nvl(冲预交, 0) + n_冲预交
    Where 结帐id = 冲销id_In And 结算方式 = v_结算方式 And Rownum < 2;
    If Sql%NotFound Then
      Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 科室id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 校对标志, 结算性质)
      Values
        (n_预交id, 12, r_Balance_Data.No, 1, r_Balance_Data.病人id, r_Balance_Data.科室id, r_Balance_Data.主页id, Null, v_结算方式,
         r_Balance_Data.收款时间, r_Balance_Data.操作员编号, r_Balance_Data.操作员姓名, n_冲预交, 冲销id_In, r_Balance_Data.缴款组id, 2, 2);
    End If;
  
    --转为住院预交 
    住院预交款_Insert(r_Balance_Data.病人id, -1 * n_冲预交, v_结算方式, r_Balance_Data.收款时间, Null, r_Balance_Data.操作员编号,
                 r_Balance_Data.操作员姓名, r_Balance_Data.缴款组id);
  
    Update 病人预交记录 Set 冲预交 = 冲预交 - n_冲预交 Where 结帐id = 冲销id_In And 结算方式 Is Null;
  End If;

  --3.完成结帐作废 
  Zl_病人结帐作废_Modify(1, r_Balance_Data.病人id, 冲销id_In, Null, Null, Null, Null, Null, Null, Null, n_误差费, Null,
                   r_Balance_Data.操作员编号, r_Balance_Data.操作员姓名, r_Balance_Data.收款时间, Null, 2);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊转住院_结帐作废;
/
Create Or Replace Procedure Zl_门诊费用转住院_Modify
(
  操作类型_In   Number,
  冲销id_In     病人预交记录.结帐id%Type,
  病人id_In     病人结帐记录.病人id%Type,
  结算方式_In   Varchar2,
  操作员编号_In 病人预交记录.操作员编号%Type := Null,
  操作员姓名_In 病人预交记录.操作员姓名%Type := Null,
  完成退费_In   Number := 0,
  关联交易id_In 病人预交记录.Id%Type := Null,
  退款时间_In   病人预交记录.收款时间%Type := Null,
  校对标志_In   病人预交记录.校对标志%Type := Null,
  误差金额_In   病人预交记录.冲预交%Type := Null,
  卡类别id_In   病人预交记录.卡类别id%Type := Null,
  卡号_In       病人预交记录.卡号%Type := Null,
  交易流水号_In 病人预交记录.交易流水号%Type := Null,
  交易说明_In   病人预交记录.交易说明%Type := Null,
  清除原交易_In Number := 0
) As
  ------------------------------------------------------------------------------------------------------------------------------
  --功能:收费结算时,修改结算的相关信息
  --操作类型_In:
  --   0-仅更新校对标志:只更新关联交易ID的校对标志
  --   1-普通退费方式:
  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
  --   2.三方卡退费结算:
  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
  --     ②卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
  --     关联交易ID_IN:
  --     清除原交易_In:1-表示在更新数据前，清除原来的交易信息(按结帐ID+关联交易ID来清除);0-表示不清除
  --   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
  --   4-消费卡结算:
  --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
  -- 完成退费_In:0-未完成退费;1-异常完成退费;2-完成退费
  -- 冲预交病人ids_In:多个用逗分分离,冲预交时有效(冲预交类别与业务操作保持一致),主要是使用家属的预交款
  -- 校对标志_In:0-完成或不需要校对;1-需要校对;2-接口已经调用成功
  ------------------------------------------------------------------------------------------------------------------------------
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_结算内容  Varchar2(500);
  v_当前结算  Varchar2(500);
  v_原结帐ids Varchar2(500);
  v_结算方式  病人预交记录.结算方式%Type;
  n_结算金额  病人预交记录.冲预交%Type;
  n_返回值    人员缴款余额.余额%Type;
  v_结算号码  病人预交记录.结算号码%Type;
  v_结算摘要  病人预交记录.摘要%Type;
  v_误差费    结算方式.名称%Type;
  n_预交id    病人预交记录.Id%Type;
  n_缴款组id  病人预交记录.缴款组id%Type;
  n_校对标志  病人预交记录.校对标志%Type;
  n_冲销金额  病人预交记录.冲预交%Type;
  d_交易时间  病人预交记录.交易时间%Type;
  v_交易人员  病人预交记录.交易人员%Type;
  n_Dec       Number; --金额小数位数

  n_Count  Number;
  l_预交id t_NumList := t_NumList();
  n_误差费 病人预交记录.冲预交%Type;

  n_是否电子票据 病人预交记录.是否电子票据%Type;

  Procedure Zl_Square_Update
  (
    结帐ids_In    Varchar2,
    预交id_In     病人预交记录.Id%Type,
    现结帐id_In   病人预交记录.结帐id%Type,
    缴款组id_In   病人预交记录.缴款组id%Type,
    退款时间_In   病人预交记录.收款时间%Type,
    结算序号_In   病人预交记录.结算序号%Type,
    退费金额_In   病人预交记录.冲预交%Type := Null,
    结算卡序号_In 病人预交记录.结算卡序号%Type := Null
  ) As
    n_预交id   病人预交记录.Id%Type;
    n_结算金额 病人预交记录.冲预交%Type;
    n_本次金额 病人预交记录.冲预交%Type;
  Begin
  
    n_结算金额 := Nvl(退费金额_In, 0);
  
    n_预交id := 预交id_In;
    --处理消费卡,结算卡在上面就已经处理了
    For v_校对 In (Select Min(a.Id) As 预交id, c.消费卡id, -1 * Nvl(Sum(c.应收金额), 0) As 结算金额, c.接口编号, c.卡号
                 From 病人预交记录 A, 病人卡结算记录 C
                 Where a.Id = c.结算id And a.结算卡序号 = 结算卡序号_In And a.记录性质 = 3 And
                       a.结帐id In (Select Column_Value From Table(f_Str2List(结帐ids_In)))
                 Group By c.消费卡id, c.接口编号, c.卡号) Loop
    
      If v_校对.结算金额 < n_结算金额 Then
        n_本次金额 := v_校对.结算金额;
        n_结算金额 := n_结算金额 - v_校对.结算金额;
      Else
        n_本次金额 := n_结算金额;
        n_结算金额 := 0;
      End If;
    
      --多条时,只更新一条
      If n_预交id = 0 Then
        Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id,
           预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质)
          Select n_预交id, 记录性质, NO, 2, 病人id, 主页id, 摘要, 结算方式, 结算号码, 退款时间_In, 缴款单位, 单位开户行, 单位帐号, 操作员编号_In, 操作员姓名_In,
                 -1 * 退费金额_In, 现结帐id_In, 缴款组id_In, 预交类别, 卡类别id, Nvl(结算卡序号, v_校对.接口编号), 卡号, 交易流水号, 交易说明, 合作单位, 2, 结算序号_In,
                 结算性质
          From 病人预交记录 A
          Where ID = v_校对.预交id;
      End If;
    
      Zl_病人卡结算记录_退款(v_校对.接口编号, v_校对.卡号, v_校对.消费卡id, n_本次金额, v_校对.预交id, n_预交id, 操作员编号_In, 操作员姓名_In, 退款时间_In);
    
      If n_结算金额 = 0 Then
        Exit;
      End If;
    End Loop;
  End;
Begin

  If 操作员姓名_In Is Null Then
    n_缴款组id := Null;
  Else
    n_缴款组id := Zl_Get组id(操作员姓名_In);
  End If;

  Select Count(1) Into n_Count From 门诊费用记录 Where 结帐id = 冲销id_In And Rownum < 2;
  If n_Count = 0 Then
    v_Err_Msg := '未找到指定的门诊收费的退费记录,请检查！';
    Raise Err_Item;
  End If;

  Select Count(1) Into n_Count From 病人预交记录 Where 结帐id = 冲销id_In And 结算方式 Is Null;

  If n_Count = 0 Then
    --插入结算方式为NULL的记录，以便退款
    Select Sum(Nvl(结帐金额, 0)) - Sum(冲预交)
    Into n_结算金额
    From (Select Sum(结帐金额) As 结帐金额, 0 As 冲预交
           From 住院费用记录
           Where 结帐id = 冲销id_In
           Union All
           Select Sum(结帐金额) As 结帐金额, 0 As 冲预交
           From 门诊费用记录
           Where 结帐id = 冲销id_In
           Union All
           Select 0 As 结帐金额, Sum(冲预交) As 冲预交
           From 病人预交记录
           Where 结帐id = 冲销id_In);
  
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 科室id, 主页id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 校对标志, 结算性质)
      Select 病人预交记录_Id.Nextval, 3, Null, 2, 病人id_In, Null, Null, Null, 退款时间_In, 操作员编号_In, 操作员姓名_In, n_结算金额, 冲销id_In,
             n_缴款组id, 0, 3
      From Dual;
  End If;

  If 操作类型_In = 0 Then
    --仅更新校对标志:只更新关联交易ID的校对标志
    Update 病人预交记录
    Set 校对标志 = 校对标志_In
    Where 结帐id = 冲销id_In And Nvl(关联交易id, 0) = Nvl(关联交易id_In, 0);
    Return;
  End If;

  --金额小数位数
  Select zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  --1.增加结算方式为空的结算数据
  n_误差费 := 误差金额_In;
  --处理误差费
  If Nvl(n_误差费, 0) <> 0 Then
    Select Nvl(Max(名称), '误差费') Into v_误差费 From 结算方式 Where Nvl(性质, 0) = 9;
  
    Update 病人预交记录 Set 冲预交 = 冲预交 + Nvl(n_误差费, 0) Where 结帐id = 冲销id_In And 结算方式 = v_误差费;
    If Sql%NotFound Then
    
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 结算序号, 缴款组id, 校对标志, 结算性质)
      Values
        (病人预交记录_Id.Nextval, 3, Null, 1, 病人id_In, v_误差费, 退款时间_In, 操作员编号_In, 操作员姓名_In, n_误差费, 冲销id_In, -1 * 冲销id_In,
         n_缴款组id, 2, 3);
    End If;
    --更新数据(结算方式为NULL的)
    Update 病人预交记录
    Set 冲预交 = 冲预交 - Nvl(n_误差费, 0)
    Where 结帐id = 冲销id_In And 结算方式 Is Null
    Returning Nvl(冲预交, 0) Into n_返回值;
  End If;

  If 操作类型_In = 1 Then
    --   1-普通退费方式:
    --各个收费结算 :格式为:"结算方式|结算金额|结算号码|结算摘要||.."
    v_结算内容 := 结算方式_In || '||';
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
    
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
    
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
    
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      v_结算号码 := LTrim(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_结算摘要 := LTrim(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
    
      --不判断“结算金额”是否为零，有可能已经退完，但这时结算方式为空的重结和冲销记录的冲预交之和为零
      If v_结算方式 Is Not Null Then
        --If Nvl(n_结算金额, 0) <> 0 Then
        n_结算金额 := Nvl(n_结算金额, 0);
        If Nvl(n_结算金额, 0) <> 0 Then
          Insert Into 病人预交记录
            (ID, 记录性质, NO, 记录状态, 病人id, 科室id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 结算序号, 缴款组id, 校对标志, 卡类别id,
             结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 结算性质, 关联交易id, 附加标志)
          Values
            (病人预交记录_Id.Nextval, 3, Null, 1, 病人id_In, Null, Null, v_结算摘要, v_结算方式, 退款时间_In, 操作员编号_In, 操作员姓名_In, n_结算金额,
             冲销id_In, -1 * 冲销id_In, n_缴款组id, 校对标志_In, Null, Null, 卡号_In, 交易流水号_In, 交易说明_In, Null, 3, 关联交易id_In, -1);
        
          Update 病人预交记录 Set 冲预交 = 冲预交 - n_结算金额 Where 结帐id = 冲销id_In And 结算方式 Is Null;
          If Sql%NotFound Then
            v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！';
            Raise Err_Item;
          End If;
        End If;
      End If;
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  If 操作类型_In = 2 And 结算方式_In Is Not Null Then
    --三方卡退费结算
    If Nvl(清除原交易_In, 0) = 1 And Nvl(关联交易id_In, 0) <> 0 Then
      --还原结算方式为空的结算金额
      --先锁表，以免并发操作
      Update 病人预交记录
      Set 冲预交 = 冲预交
      Where 结帐id = 冲销id_In And 关联交易id = 关联交易id_In And Mod(记录性质, 10) <> 1;
    
      Select Sum(冲预交)
      Into n_结算金额
      From 病人预交记录
      Where 结帐id = 冲销id_In And 关联交易id = 关联交易id_In And Mod(记录性质, 10) <> 1;
    
      Update 病人预交记录
      Set 冲预交 = Nvl(冲预交, 0) + Nvl(n_结算金额, 0)
      Where 结帐id = 冲销id_In And 结算方式 Is Null;
    
      Delete 病人预交记录 Where 结帐id = 冲销id_In And 关联交易id = 关联交易id_In And Mod(记录性质, 10) <> 1;
    End If;
  
    d_交易时间 := Sysdate;
    v_交易人员 := zl_UserName;
    --   2.三方卡退费结算:
    v_当前结算 := 结算方式_In;
    v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
    v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
    n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
    v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
    v_结算号码 := LTrim(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
    v_结算摘要 := LTrim(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
  
    If Nvl(n_结算金额, 0) <> 0 Then
      --先更新：
      Update 病人预交记录
      Set 冲预交 = Nvl(冲预交, 0) + n_结算金额
      Where 结帐id = 冲销id_In And 卡类别id = 卡类别id_In And 关联交易id = 关联交易id_In And 结算方式 = v_结算方式;
    
      If Sql%NotFound Then
        Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 科室id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 结算序号, 缴款组id, 校对标志, 卡类别id,
           结算卡序号, 卡号, 交易流水号, 交易说明, 结算号码, 结算性质, 关联交易id, 交易时间, 交易人员, 附加标志)
        Values
          (n_预交id, 3, Null, 2, 病人id_In, Null, Null, v_结算摘要, v_结算方式, 退款时间_In, 操作员编号_In, 操作员姓名_In, n_结算金额, 冲销id_In,
           -1 * 冲销id_In, n_缴款组id, 校对标志_In, 卡类别id_In, Null, 卡号_In, 交易流水号_In, 交易说明_In, v_结算号码, 3, 关联交易id_In, d_交易时间,
           v_交易人员, -1);
      
        --调用其他结算信息更新
        Zl_Custom_Balance_Update(n_预交id);
      
      End If;
    
      Update 病人预交记录 Set 冲预交 = 冲预交 - n_结算金额 Where 结帐id = 冲销id_In And 结算方式 Is Null;
      If Sql%NotFound Then
        v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！';
        Raise Err_Item;
      End If;
    End If;
  End If;

  If 操作类型_In = 3 Then
    --   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
    --3.1检查是否已经存在医保结算数据,存在先删除
    n_结算金额 := 0;
  
    If 校对标志_In = 0 Then
      n_校对标志 := 2;
    Else
      n_校对标志 := 1;
    End If;
  
    For v_医保 In (Select ID, 结算方式, 冲预交
                 From 病人预交记录 A
                 Where 结帐id = 冲销id_In And 卡类别id Is Null And Exists
                  (Select 1 From 结算方式 Where a.结算方式 = 名称 And 性质 In (3, 4)) And Mod(记录性质, 10) <> 1 And 卡类别id Is Null) Loop
      n_结算金额 := n_结算金额 + Nvl(v_医保.冲预交, 0);
      l_预交id.Extend;
      l_预交id(l_预交id.Count) := v_医保.Id;
    End Loop;
  
    If Nvl(n_结算金额, 0) <> 0 Then
      Update 病人预交记录
      Set 冲预交 = Nvl(冲预交, 0) + Nvl(n_结算金额, 0)
      Where 结帐id = 冲销id_In And 结算方式 Is Null;
      If Sql%NotFound Then
        v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[收费结算窗口]中重新收费！';
        Raise Err_Item;
      End If;
    End If;
  
    If l_预交id.Count <> 0 Then
      Forall I In 1 .. l_预交id.Count
        Delete 病人预交记录 Where ID = l_预交id(I);
    End If;
  
    If 结算方式_In Is Not Null Then
      v_结算内容 := 结算方式_In || '||';
    End If;
    d_交易时间 := Sysdate;
    v_交易人员 := zl_UserName;
  
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      n_结算金额 := To_Number(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
    
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 科室id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 结算序号, 缴款组id, 校对标志, 结算性质, 关联交易id,
         交易时间, 交易人员, 附加标志)
      Values
        (病人预交记录_Id.Nextval, 3, Null, 2, 病人id_In, Null, Null, '保险结算', v_结算方式, 退款时间_In, 操作员编号_In, 操作员姓名_In, n_结算金额,
         冲销id_In, -1 * 冲销id_In, n_缴款组id, n_校对标志, 3, 关联交易id_In, d_交易时间, v_交易人员, -1);
    
      --更新数据(结算方式为NULL的)
      Update 病人预交记录
      Set 冲预交 = 冲预交 - n_结算金额
      Where 结帐id = 冲销id_In And 结算方式 Is Null
      Returning Nvl(冲预交, 0) Into n_返回值;
    
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;

  --4-消费卡批量结算
  If 操作类型_In = 4 Then
    Null;
  End If;

  If Nvl(完成退费_In, 0) = 0 Then
    Return;
  End If;

  -----------------------------------------------------------------------------------------
  --完成收费,需要处理人员缴款余额,预交记录(结算方式=NULL)
  If Nvl(完成退费_In, 0) = 1 Then
    Update 病人预交记录 Set 校对标志 = 0 Where 结帐id = 冲销id_In;
    Return;
  End If;

  --处理消费卡
  v_原结帐ids := Null;
  For c_原结帐 In (Select Distinct a.结帐id
                From 病人预交记录 A,
                     (Select Distinct 结帐id
                       From 门诊费用记录
                       Where NO In (Select Distinct NO From 门诊费用记录 Where 结帐id = 冲销id_In And Mod(记录性质, 10) = 1)) B
                Where a.结帐id = b.结帐id And Mod(a.记录性质, 10) <> 1 And 记录状态 In (3, 1) And a.结算卡序号 Is Not Null) Loop
    v_原结帐ids := Nvl(v_原结帐ids, '') || ',' || c_原结帐.结帐id;
  End Loop;

  If v_原结帐ids Is Not Null Then
    v_原结帐ids := Substr(v_原结帐ids, 2);
  End If;

  For c_消费卡 In (Select ID, 结算卡序号, 结算方式, 冲预交, 关联交易id
                From 病人预交记录
                Where 结帐id = 冲销id_In And 附加标志 = -1 And Nvl(结算卡序号, 0) <> 0) Loop
    n_冲销金额 := Nvl(c_消费卡.冲预交, 0);
    If n_冲销金额 <> 0 Then
      Zl_Square_Update(v_原结帐ids, c_消费卡.Id, 冲销id_In, n_缴款组id, 退款时间_In, -1 * 冲销id_In, -1 * n_冲销金额, c_消费卡.结算卡序号);
    End If;
  End Loop;

  --1.删除结算方式为NULL的预交记录
  Delete 病人预交记录 Where 结帐id = 冲销id_In And 结算方式 Is Null And Nvl(冲预交, 0) = 0;
  If Sql%NotFound Then
    Select Count(*) Into n_Count From 病人预交记录 A Where 结帐id = 冲销id_In And 结算方式 Is Null;
    If n_Count <> 0 Then
      v_Err_Msg := '还存在未退款的数据,不能完成结帐作废操作!';
    Else
      v_Err_Msg := '结算信息错误,可能因为并发原因造成结算信息错误,请在[结帐窗口]中重新作废！!';
    End If;
    Raise Err_Item;
  End If;

  --结算金额为零时，增加一条金额为0的病人预交记录
  Select Count(*) Into n_Count From 病人预交记录 A Where 结帐id = 冲销id_In;

  If n_Count = 0 Then
    If v_结算方式 Is Null Then
      Select Max(结算方式) Into v_结算方式 From 结算方式应用 Where 应用场合 = '收费' And Nvl(缺省标志, 0) = 1;
      If v_结算方式 Is Null Then
        Select Nvl(Max(名称), '现金') Into v_结算方式 From 结算方式 Where Nvl(性质, 0) = 1;
      End If;
    End If;
  
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 科室id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 结算序号, 缴款组id, 校对标志, 卡类别id, 结算卡序号,
       卡号, 交易流水号, 交易说明, 结算号码, 结算性质)
    Values
      (病人预交记录_Id.Nextval, 3, Null, 2, 病人id_In, Null, Null, Null, v_结算方式, 退款时间_In, 操作员编号_In, 操作员姓名_In, 0, 冲销id_In,
       -1 * 冲销id_In, n_缴款组id, 2, Null, Null, Null, Null, 交易说明_In, Null, 2);
  End If;

  --4.结算总额要与费用信息保持一致
  Select Sum(冲预交), Sum(结帐金额)
  Into n_返回值, n_结算金额
  From (Select Sum(冲预交) As 冲预交, 0 As 结帐金额
         From 病人预交记录
         Where 结帐id = 冲销id_In
         Union All
         Select 0, Sum(结帐金额)
         From 门诊费用记录
         Where 结帐id = 冲销id_In
         Union All
         Select 0, Sum(结帐金额) As 结帐金额
         From 住院费用记录
         Where 结帐id = 冲销id_In);

  If Nvl(n_返回值, 0) <> Nvl(n_结算金额, 0) Then
    v_Err_Msg := '结算总额与费用总额不一致,不能进行作废操作，请与系统管理员联系!';
    Raise Err_Item;
  End If;

  --5.更新人员缴款数据
  For c_汇总 In (Select 操作员姓名, 结算方式, -1 * Sum(冲预交) As 冲预交
               From 病人预交记录
               Where 结帐id = 冲销id_In And 附加标志 = -1
               Group By 操作员姓名, 结算方式) Loop
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) - Nvl(c_汇总.冲预交, 0)
    Where 收款员 = c_汇总.操作员姓名 And 性质 = 1 And 结算方式 = c_汇总.结算方式
    Returning 余额 Into n_返回值;
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额
        (收款员, 结算方式, 性质, 余额)
      Values
        (c_汇总.操作员姓名, c_汇总.结算方式, 1, -1 * Nvl(c_汇总.冲预交, 0));
      n_返回值 := Nvl(c_汇总.冲预交, 0);
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 人员缴款余额
      Where 收款员 = c_汇总.操作员姓名 And 性质 = 1 And 结算方式 = c_汇总.结算方式 And Nvl(余额, 0) = 0;
    End If;
  End Loop;

  --2.处理缴款数据和找补数据及校对标志更新为0  
  Select Max(a.是否电子票据)
  Into n_是否电子票据
  From 病人预交记录 A,
       (Select Max(b.结帐id) As 结帐id
         From 门诊费用记录 A, 门诊费用记录 B
         Where a.结帐id = 冲销id_In And a.No = b.No And b.记录性质 = 1 And b.记录状态 In (1, 3)) B
  Where a.结帐id = b.结帐id And a.记录性质 In (11, 3);

  Update 病人预交记录 Set 校对标志 = 0, 附加标志 = Null, 是否电子票据 = n_是否电子票据 Where 结帐id = 冲销id_In;

  --3.更新费用状态
  Update 门诊费用记录 Set 费用状态 = Null Where 结帐id = 冲销id_In;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊费用转住院_Modify;
/

Create Or Replace Procedure Zl_病人挂号收费_Modify
(
  单据号_In     门诊费用记录.No%Type,
  结帐id_In     门诊费用记录.结帐id%Type,
  结算信息_In   Varchar2,
  结算类型_In   Number := 0,
  完成标志_In   Number := 0,
  生成队列_In   Number := 0,
  退号重用_In   Number := 1,
  收回票据号_In Varchar2 := Null,
  连续更新_In   Number := 0,
  关联交易id_In 病人预交记录.关联交易id%Type := Null,
  卡类别id_In   病人预交记录.卡类别id%Type := Null,
  卡号_In       病人预交记录.卡号%Type := Null,
  交易流水号_In 病人预交记录.交易流水号%Type := Null,
  交易说明_In   病人预交记录.交易说明%Type := Null,
  普通结算_In   Number := 0,
  校对标志_In   Number := 2,
  电子票据_In   病人预交记录.预交电子票据%Type := 0
) As
  --功能:重新更新挂号结算信息,分摊到指定的单据上
  --结算信息_In:为空时,表示只更新预交的标志(以预结算和结算一样时,才会使用此方式)
  --结算类型_In：
  -- 0-普通方式结算：允许传入多个结算方式,格式为:"结算方式,结算金额,结算号码,结算摘要|.." ;
  --                 也允许传入空.为空时只更新性质1，2且卡类别ID=null的校对标志
  -- 1-三方卡：只能传入一个结算方式,格式为:"结算方式,结算金额,结算号码,结算摘要"
  --           也允许传入空.为空时只更新性质7,8且卡类别ID=卡类别ID_In的校对标志
  -- 2-消费卡；只能传入一个结算方式,格式为:"结算方式,结算金额"
  -- 3-预交支付：必传且只能传入一个预交病人IDs 格式为:"结算金额|冲预交病人ids"
  -- 4-医保结算：允许传入多个,格式为:"结算方式,结算金额|.."
  --             也允许传入空.为空时只更新性质3，4且卡类别ID=BULL的校对标志
  --                                   第二个结算方式及以后传入连续更新_In = 1
  --连续更新_In：三方卡混合结算后传入，第一个结算方式在原结算记录上更新，并删除其他相同关联交易ID记录,连续更新标识不传；后续更新需要传入
  --完成标志_In：更新所有完成的收费
  --普通结算_In: 三方接口返回的结算方式是否保存卡类别ID
  --校对标志_In:医保和三方卡结算时有效
  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  n_组id     病人预交记录.缴款组id%Type;
  v_结算内容 Varchar2(3000);
  v_当前结算 Varchar2(200);

  v_现金          病人预交记录.结算方式%Type;
  v_结算方式      病人预交记录.结算方式%Type;
  n_结算金额      病人预交记录.冲预交%Type;
  v_结算号码      病人预交记录.结算号码%Type;
  v_结算摘要      病人预交记录.摘要%Type;
  n_冲预交        病人预交记录.冲预交%Type;
  n_冲销金额      病人预交记录.冲预交%Type;
  v_冲预交病人ids Varchar2(1000);
  n_预交id        病人预交记录.Id%Type;
  n_原预交id      病人预交记录.Id%Type;
  n_关联id        病人预交记录.Id%Type;
  n_充值id        病人预交记录.Id%Type;
  v_操作员编号    病人预交记录.操作员编号%Type;
  v_操作员姓名    病人预交记录.操作员姓名%Type;
  d_Date          门诊费用记录.登记时间%Type;
  v_No            门诊费用记录.No%Type;
  n_病人id        门诊费用记录.病人id%Type;
  n_卡类别id      病人预交记录.卡类别id%Type;
  n_返回值        病人余额.预交余额%Type;
  n_收费作废      Number;
  n_Count         Number;
  l_预交id        t_Numlist := t_Numlist();

  n_预约生成队列   Number;
  n_分诊台签到排队 Number;
  n_记账           Number;
  n_退号           Number; --0-挂号；1-退号
  n_退费           Number;
  n_预约挂号       Number; --0-挂号；1-预约挂号
  n_预约标志       Number; --0-挂号；1-预约挂号
  n_挂号id         病人挂号记录.Id%Type;
  n_记录id         病人挂号记录.出诊记录id%Type;
  v_星期           挂号安排限制.限制项目%Type;
  d_发生时间       病人挂号记录.发生时间%Type;
  n_序号           门诊费用记录.序号%Type;
  n_安排id         挂号安排.Id%Type;
  n_计划id         挂号安排计划.Id%Type;
  v_号码           病人挂号记录.号别%Type;
  n_号序           病人挂号记录.号序%Type;
  n_打印id         票据打印内容.Id%Type;

  n_排队       Number;
  n_当天排队   Number;
  n_分时点显示 Number;
  n_分时段     Number;
  v_排队号码   排队叫号队列.排队号码%Type;
  v_排队序号   排队叫号队列.排队序号%Type;
  v_队列名称   排队叫号队列.队列名称%Type;
  d_排队时间   排队叫号队列.排队时间%Type;
  n_是否电子票据 Number(2);
  n_险类         保险结算记录.险类%Type;

  Cursor c_Registinfo(v_状态 门诊费用记录.记录状态%Type) Is
    Select a.发生时间, a.登记时间, c.接收时间, Nvl(c.挂号项目id, a.收费细目id) As 项目id, c.执行部门id As 科室id, c.执行人 As 医生姓名, d.Id As 医生id,
           c.号别 As 号码, c.号序, a.操作员姓名
    From 门诊费用记录 a, 病人挂号记录 c, 人员表 d
    Where a.记录性质 = 4 And a.No = 单据号_In And a.No = c.No And a.记录状态 = v_状态 And c.执行人 = d.姓名(+) And Rownum < 2;
  r_Registrow c_Registinfo%Rowtype;
Begin
  --n_退号：0-挂号更新；1-退号更新
  Select Max(Id), Decode(Nvl(Max(记录状态), 0), 3, 1, 0), Decode(Nvl(Max(记录性质), 0), 2, 1, 0),
         Decode(Nvl(Max(预约), 0), 1, 1, 0), Nvl(Max(出诊记录id), 0), Nvl(Max(号别), '0'), Max(发生时间), Nvl(Max(号序), 0)
  Into n_挂号id, n_退号, n_预约挂号, n_预约标志, n_记录id, v_号码, d_发生时间, n_号序
  From 病人挂号记录
  Where No = 单据号_In And 记录状态 In (1, 3) And Rownum < 2;

  Select Nvl(Max(记帐费用), 0), Min(序号)
  Into n_记账, n_序号
  From 门诊费用记录
  Where No = 单据号_In And 记录性质 = 4 And
        登记时间 = (Select Max(登记时间) From 门诊费用记录 Where No = 单据号_In And 记录性质 = 4)
  Order By 序号, 登记时间 Desc;

  Select Nvl(Max(名称), '现金') Into v_现金 From 结算方式 Where 性质 = 1;
  If Nvl(n_序号, 0) <> 1 Then
    n_挂号id := 0;
    n_退号   := 0;
  End If;
  If 普通结算_In = 0 And Nvl(卡类别id_In, 0) <> 0 Then
    n_卡类别id := 卡类别id_In;
  End If;

  If n_预约挂号 = 0 And n_记账 = 0 Then
    Select Count(1)
    Into n_收费作废
    From 门诊费用记录
    Where No = 单据号_In And 记录性质 = 4 And 记录状态 = 3 And 费用状态 = 1 And Rownum < 2;
    Select Count(1) Into n_退费 From 门诊费用记录 Where 结帐id = 结帐id_In And 记录状态 = 2 And Rownum < 2;
  
    Select No, 操作员编号, 操作员姓名, 病人id, 登记时间, 缴款组id
    Into v_No, v_操作员编号, v_操作员姓名, n_病人id, d_Date, n_组id
    From 门诊费用记录
    Where 结帐id = 结帐id_In And Rownum < 2;
  
    If Nvl(结算类型_In, 0) = 0 Then
      --0.普通结算
      If 结算信息_In Is Null Then
        Update 病人预交记录 a
        Set a.校对标志 = 2
        Where a.结帐id = 结帐id_In And Nvl(卡类别id, 0) = 0 And a.结算方式 Is Not Null And Exists
         (Select 1 From 结算方式 Where a.结算方式 = 名称 And 性质 In (1, 2));
      Else
        n_Count    := 0;
        v_结算内容 := 结算信息_In || '|';
        While v_结算内容 Is Not Null Loop
          v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '|') - 1) || ',,,';
          v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
        
          v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
          n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1));
          If n_退费 = 1 Then
            n_结算金额 := -1 * n_结算金额;
          Else
            v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
            v_结算号码 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
          
            v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
            v_结算摘要 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
          End If;
        
          If v_结算方式 Is Null Or n_结算金额 Is Null Then
            v_Err_Msg := '结算方式不正确！';
            Raise Err_Item;
          End If;
        
          If n_退费 = 0 And Nvl(n_结算金额, 0) = 0 Then
            Delete 病人预交记录
            Where Nvl(卡类别id, 0) = 0 And 记录性质 = 4 And 结帐id = 结帐id_In And 结算方式 = v_结算方式 And
                  Nvl(关联交易id, 0) = Nvl(关联交易id_In, 0) Return Nvl(Sum(冲预交), 0) Into n_结算金额;
            n_冲预交 := Nvl(n_冲预交, 0) + n_结算金额;
          Else
            Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
            Insert Into 病人预交记录
              (Id, 记录性质, No, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 结算号码, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款, 找补, 缴款组id, 预交类别,
               合作单位, 结算序号, 交易说明, 校对标志, 待转出, 结算性质, 会话号, 关联交易id)
              Select n_预交id, 记录性质, No, 记录状态, 病人id, 主页id, 科室id, Nvl(v_结算摘要, '挂号收费'), v_结算方式, v_结算号码, 收款时间, 操作员编号, 操作员姓名,
                     n_结算金额, 结帐id, 缴款, 找补, 缴款组id, 预交类别, 合作单位, 结算序号, 交易说明_In, 2, 待转出, 结算性质, 会话号, n_预交id
              From 病人预交记录
              Where Nvl(卡类别id, 0) = Nvl(n_卡类别id, 0) And 结帐id = 结帐id_In And 记录性质 = 4 And Rownum < 2;
            n_冲预交 := Nvl(n_冲预交, 0) - n_结算金额;
          End If;
        
          v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '|') + 1);
        End Loop;
      End If;
    End If;
  
    If Nvl(结算类型_In, 0) = 1 Then
      --1.三方卡
      If 结算信息_In Is Null Then
        Update 病人预交记录
        Set 交易流水号 = Nvl(交易流水号_In, 交易流水号), 交易说明 = Nvl(交易说明_In, 交易说明), 校对标志 = 校对标志_In, 卡号 = Nvl(卡号_In, 卡号)
        Where 卡类别id = 卡类别id_In And 结帐id = 结帐id_In
        Returning Id Bulk Collect Into l_预交id;
      
        If 校对标志_In = 2 Then
          --调用三方自主更新接口信息
          For i In 1 .. l_预交id.Count Loop
            Zl_Custom_Balance_Update(l_预交id(i));
          End Loop;
        End If;
      Else
        n_Count    := 0;
        v_结算内容 := 结算信息_In || '|';
        While v_结算内容 Is Not Null Loop
          v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '|') - 1);
          v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
        
          v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
          n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1));
        
          If n_退费 = 1 Then
            n_结算金额 := -1 * n_结算金额;
          End If;
        
          v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
          v_结算号码 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
        
          v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
          v_结算摘要 := v_当前结算;
        
          If v_结算方式 Is Null Or n_结算金额 Is Null Then
            v_Err_Msg := '结算方式不正确！';
            Raise Err_Item;
          End If;
        
          n_Count := n_Count + 1;
          If Nvl(连续更新_In, 0) = 1 Then
            Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
            Insert Into 病人预交记录
              (Id, 记录性质, No, 记录状态, 病人id, 主页id, 科室id, 缴款单位, 单位开户行, 单位帐号, 摘要, 结算方式, 结算号码, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id,
               缴款, 找补, 缴款组id, 预交类别, 卡类别id, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 校对标志, 待转出, 结算性质, 会话号, 关联交易id)
              Select n_预交id, 记录性质, No, 记录状态, 病人id, 主页id, 科室id, 缴款单位, 单位开户行, 单位帐号,
                     Decode(n_退费, 1, Null, Nvl(v_结算摘要, 摘要)), v_结算方式, Decode(n_退费, 1, Null, Nvl(v_结算号码, 结算号码)), 收款时间,
                     操作员编号, 操作员姓名, n_结算金额, 结帐id, 缴款, 找补, 缴款组id, 预交类别, n_卡类别id, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位, 结算序号,
                     校对标志_In, 待转出, 结算性质, 会话号, 关联交易id_In
              From 病人预交记录
              Where 结帐id = 结帐id_In And 记录性质 = 4 And Rownum < 2;
          Else
            If n_Count = 1 Then
              --第一次保存同一关联交易ID的总金额，更新其中一条，删除其他的记录，以便后续直接更新
              Select Nvl(Sum(冲预交), 0)
              Into n_冲预交
              From 病人预交记录
              Where Nvl(关联交易id, 0) = Nvl(关联交易id_In, 0) And Nvl(卡类别id, 0) = Nvl(卡类别id_In, 0) And 结帐id = 结帐id_In And
                    校对标志 = 1;
              Update 病人预交记录
              Set 冲预交 = n_结算金额, 卡类别id = Decode(结算类型_In, 2, Null, n_卡类别id),
                  摘要 = Decode(n_退费, 1, Null, Nvl(v_结算摘要, 摘要)), 结算号码 = Decode(n_退费, 1, Null, Nvl(v_结算号码, 结算号码)),
                  结算卡序号 = Decode(结算类型_In, 2, n_卡类别id, Null), 卡号 = 卡号_In, 交易流水号 = 交易流水号_In, 交易说明 = 交易说明_In,
                  校对标志 = 校对标志_In
              Where Nvl(关联交易id, 0) = Nvl(关联交易id_In, 0) And 结算方式 = v_结算方式 And 结帐id = 结帐id_In And 校对标志 = 1 And Rownum < 2
               Return Id Into n_预交id;
            
              If Sql%Notfound Then
                Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
                Insert Into 病人预交记录
                  (Id, 记录性质, No, 记录状态, 病人id, 主页id, 科室id, 缴款单位, 单位开户行, 单位帐号, 摘要, 结算方式, 结算号码, 收款时间, 操作员编号, 操作员姓名, 冲预交,
                   结帐id, 缴款, 找补, 缴款组id, 预交类别, 卡类别id, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 校对标志, 待转出, 结算性质, 会话号, 关联交易id)
                  Select n_预交id, 记录性质, No, 记录状态, 病人id, 主页id, 科室id, 缴款单位, 单位开户行, 单位帐号,
                         Decode(n_退费, 1, Null, Nvl(v_结算摘要, 摘要)), v_结算方式,
                         Decode(n_退费, 1, Null, Nvl(v_结算号码, 结算号码)), 收款时间, 操作员编号, 操作员姓名, n_结算金额, 结帐id, 缴款, 找补, 缴款组id, 预交类别,
                         n_卡类别id, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位, 结算序号, 校对标志_In, 待转出, 结算性质, 会话号, 关联交易id_In
                  From 病人预交记录
                  Where 结帐id = 结帐id_In And 记录性质 = 4 And Rownum < 2;
              End If;
              Delete From 病人预交记录
              Where Id <> Nvl(n_预交id, 0) And Nvl(关联交易id, 0) = Nvl(关联交易id_In, 0) And Nvl(卡类别id, 0) = Nvl(卡类别id_In, 0) And
                    结帐id = 结帐id_In And 校对标志 = 1;
            Else
              Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
              Insert Into 病人预交记录
                (Id, 记录性质, No, 记录状态, 病人id, 主页id, 科室id, 缴款单位, 单位开户行, 单位帐号, 摘要, 金额, 结算方式, 结算号码, 收款时间, 操作员编号, 操作员姓名, 冲预交,
                 结帐id, 缴款, 找补, 缴款组id, 预交类别, 卡类别id, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 校对标志, 待转出, 结算性质, 会话号, 关联交易id)
                Select n_预交id, 记录性质, No, 记录状态, 病人id, 主页id, 科室id, 缴款单位, 单位开户行, 单位帐号,
                       Decode(n_退费, 1, Null, Nvl(v_结算摘要, 摘要)), 金额, v_结算方式,
                       Decode(n_退费, 1, Null, Nvl(v_结算号码, 结算号码)), 收款时间, 操作员编号, 操作员姓名, n_结算金额, 结帐id, 缴款, 找补, 缴款组id, 预交类别,
                       卡类别id_In, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位, 结算序号, 校对标志_In, 待转出, 结算性质, 会话号, 关联交易id_In
                From 病人预交记录
                Where 结帐id = 结帐id_In And 记录性质 = 4 And Rownum < 2;
            End If;
          End If;
          v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '|') + 1);
          n_冲预交   := Nvl(n_冲预交, 0) - n_结算金额;
          If 校对标志_In = 2 Then
            --调用三方自主更新接口信息
            Zl_Custom_Balance_Update(n_预交id);
          End If;
        End Loop;
      End If;
      If 校对标志_In = 2 Then
        Update 病人预交记录
        Set 交易时间 = 收款时间, 交易人员 = 操作员姓名
        Where 记录性质 = 4 And Nvl(卡类别id, 0) > 0 And 结帐id = 结帐id_In;
      End If;
    End If;
  
    If Nvl(结算类型_In, 0) = 2 Then
      --2.消费卡
      v_结算内容 := 结算信息_In || '|';
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '|') - 1);
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
    
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
      n_结算金额 := To_Number(v_当前结算);
      If n_退费 = 1 Then
        n_结算金额 := -1 * n_结算金额;
      End If;
    
      If v_结算方式 Is Null Or n_结算金额 Is Null Then
        v_Err_Msg := '结算方式不正确！';
        Raise Err_Item;
      End If;
    
      If n_结算金额 <> 0 Then
        If n_退费 = 0 Then
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (Id, 记录性质, No, 记录状态, 病人id, 主页id, 科室id, 缴款单位, 单位开户行, 单位帐号, 摘要, 结算方式, 结算号码, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款,
             找补, 缴款组id, 预交类别, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 校对标志, 待转出, 结算性质, 会话号, 关联交易id)
            Select n_预交id, 记录性质, No, 记录状态, 病人id, 主页id, 科室id, 缴款单位, 单位开户行, 单位帐号, 摘要, v_结算方式, 结算号码, 收款时间, 操作员编号, 操作员姓名,
                   n_结算金额, 结帐id, 缴款, 找补, 缴款组id, 预交类别, 卡类别id_In, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位, 结算序号, 2, 待转出, 结算性质, 会话号,
                   n_预交id
            From 病人预交记录
            Where 结帐id = 结帐id_In And 记录性质 = 4 And Rownum < 2;
        
          Zl_病人卡结算记录_支付(卡类别id_In, 卡号_In, 0, n_结算金额, n_预交id, v_操作员编号, v_操作员姓名, d_Date);
          n_冲预交 := Nvl(n_冲预交, 0) - n_结算金额;
        Else
          Select Nvl(Id, 0), -1 * Nvl(冲预交, 0)
          Into n_原预交id, n_冲销金额
          From 病人预交记录
          Where No = 单据号_In And 记录性质 = 4 And 记录状态 = 3 And 结算方式 = v_结算方式 And 结算卡序号 = 卡类别id_In;
          If n_原预交id = 0 Then
            v_Err_Msg := '未找到原结算记录！';
            Raise Err_Item;
          End If;
          If n_冲销金额 <> n_结算金额 Then
            v_Err_Msg := '消费卡退款金额不一致！';
            Raise Err_Item;
          End If;
        
          Update 病人预交记录
          Set 校对标志 = 2
          Where 结帐id = 结帐id_In And 结算方式 = v_结算方式 And 结算卡序号 = 卡类别id_In
          Returning Id Into n_预交id;
          If Sql%Notfound Then
            Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
            Insert Into 病人预交记录
              (Id, 记录性质, No, 记录状态, 病人id, 主页id, 科室id, 缴款单位, 单位开户行, 单位帐号, 摘要, 结算方式, 结算号码, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id,
               缴款, 找补, 缴款组id, 预交类别, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 校对标志, 待转出, 结算性质, 会话号, 关联交易id)
              Select n_预交id, 记录性质, No, 记录状态, 病人id, 主页id, 科室id, 缴款单位, 单位开户行, 单位帐号, 摘要, v_结算方式, 结算号码, 收款时间, 操作员编号, 操作员姓名,
                     n_结算金额, 结帐id, 缴款, 找补, 缴款组id, 预交类别, 卡类别id_In, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位, 结算序号, 2, 待转出, 结算性质,
                     会话号, n_预交id
              From 病人预交记录
              Where 结帐id = 结帐id_In And 记录性质 = 4 And Rownum < 2;
            n_冲预交 := Nvl(n_冲预交, 0) - n_结算金额;
          End If;
          Zl_病人卡结算记录_退款(卡类别id_In,
                        卡号_In,
                        0,
                        -1 * n_结算金额,
                        n_原预交id,
                        n_预交id,
                        v_操作员编号,
                        v_操作员姓名,
                        d_Date);
        End If;
      End If;
    End If;
  
    If Nvl(结算类型_In, 0) = 3 Then
      --3.预交
      v_当前结算      := 结算信息_In || '|';
      n_冲预交        := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_当前结算      := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      v_冲预交病人ids := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
    
      If Nvl(n_冲预交, 0) <> 0 Then
        If n_退费 = 1 Then
          n_冲销金额 := n_冲预交;
          For c_预交 In (Select Max(Id) As 预交id, No, 病人id, 预交类别, Sum(冲预交) As 冲预交
                       From 病人预交记录
                       Where 记录性质 In (1, 11) And
                             结帐id In (Select Distinct 结帐id From 门诊费用记录 Where No = 单据号_In And 记录性质 = 4)
                       Group By No, 病人id, 预交类别, 结帐id
                       Having Sum(冲预交) > 0) Loop
            If n_冲销金额 > Nvl(c_预交.冲预交, 0) Then
              n_结算金额 := Nvl(c_预交.冲预交, 0);
              n_冲销金额 := n_冲销金额 - Nvl(c_预交.冲预交, 0);
            Else
              n_结算金额 := n_冲销金额;
              n_冲销金额 := 0;
            End If;
          
            Insert Into 病人预交记录
              (Id, No, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,
               冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质, 关联交易id, 交易人员, 交易时间, 校对标志)
              Select 病人预交记录_Id.Nextval, No, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号,
                     d_Date, v_操作员姓名, v_操作员编号, -1 * n_结算金额, 结帐id_In, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4,
                     关联交易id, v_操作员姓名, d_Date, 2
              From 病人预交记录
              Where Id = c_预交.预交id And Rownum < 2;
          
            --更新预交单据余额
            Select Max(Id) Into n_充值id From 病人预交记录 Where No = c_预交.No And 记录性质 = 1 And 记录状态 <> 2;
            If Nvl(n_充值id, 0) <> 0 Then
              Update 预交单据余额
              Set 预交余额 = Nvl(预交余额, 0) + n_结算金额
              Where 病人id = c_预交.病人id And 预交id = n_充值id
              Returning 预交余额 Into n_返回值;
              If Sql%Rowcount = 0 Then
                Insert Into 预交单据余额
                  (预交id, 病人id, 预交类别, 预交余额)
                Values
                  (n_充值id, c_预交.病人id, Nvl(c_预交.预交类别, 2), n_结算金额);
                n_返回值 := n_结算金额;
              End If;
              If Nvl(n_返回值, 0) = 0 Then
                Delete From 预交单据余额 Where 预交id = n_充值id And Nvl(预交余额, 0) = 0;
              End If;
            End If;
            If n_冲销金额 = 0 Then
              Exit;
            End If;
          End Loop;
          If Nvl(n_冲销金额, 0) <> 0 Then
            v_Err_Msg := '退预交金额超过了支付的预交金额，请检查！';
            Raise Err_Item;
          End If;
        
          Update 病人余额
          Set 预交余额 = Nvl(预交余额, 0) + n_冲预交
          Where 病人id = n_病人id And 类型 = 1 And 性质 = 1
          Returning 预交余额 Into n_返回值;
          If Sql%Rowcount = 0 Then
            Insert Into 病人余额 (病人id, 预交余额, 性质, 类型) Values (n_病人id, n_冲预交, 1, 1);
            n_返回值 := n_冲预交;
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 病人余额
            Where 病人id = n_病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
          End If;
        Else
          Zl_病人预交记录_冲预交(n_病人id,
                        结帐id_In,
                        n_冲预交,
                        1,
                        v_操作员编号,
                        v_操作员姓名,
                        d_Date,
                        v_冲预交病人ids,
                        4,
                        1,
                        2,
                        1);
          n_冲预交 := 0; --Zl_病人预交记录_冲预交 中冲减了NULL的金额
        End If;
      End If;
    End If;
  
    If Nvl(结算类型_In, 0) = 4 Then
      --4.医保
      If 结算信息_In Is Null Then
        --预结算和结算一致时,才会只更新标志
        Update 病人预交记录 a
        Set a.校对标志 = 校对标志_In
        Where a.结帐id = 结帐id_In And Nvl(卡类别id, 0) = 0 And Exists
         (Select 1 From 结算方式 Where a.结算方式 = 名称 And 性质 In (3, 4));
        --医保相关表的处理
        Update 保险结算明细 Set 标志 = 校对标志_In Where 结帐id = 结帐id_In;
      Else
        --删除所有医保结算数据(其他结算方式不删除)
        Select Nvl(Sum(a.冲预交), 0)
        Into n_冲预交
        From 病人预交记录 a
        Where a.结帐id = 结帐id_In And a.记录性质 = 4 And Nvl(a.卡类别id, 0) = 0 And a.结算方式 Is Not Null And Exists
         (Select 1 From 结算方式 Where 性质 In (3, 4) And a.结算方式 = 名称);
      
        n_Count    := 0;
        v_结算内容 := 结算信息_In || '|';
        While v_结算内容 Is Not Null Loop
          v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '|') - 1) || ',,,';
          v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
        
          v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
          n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1));
          If n_退费 = 1 Then
            n_结算金额 := -1 * n_结算金额;
          End If;
        
          If v_结算方式 Is Null Or n_结算金额 Is Null Then
            v_Err_Msg := '结算方式不正确！';
            Raise Err_Item;
          End If;
        
          n_Count := n_Count + 1;
          If n_Count = 1 Then
            --第一次保存同一关联交易ID的总金额，更新其中一条，删除其他的记录，以便后续直接更新
            Update 病人预交记录
            Set 冲预交 = n_结算金额, 摘要 = '医保挂号', 校对标志 = 校对标志_In, 关联交易id = Id
            Where Nvl(卡类别id, 0) = 0 And 结算方式 = v_结算方式 And 结帐id = 结帐id_In And Rownum < 2 Return Id Into n_关联id;
          
            If Sql%Notfound Then
              Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
              n_关联id := n_预交id;
              Insert Into 病人预交记录
                (Id, 记录性质, No, 记录状态, 病人id, 主页id, 科室id, 缴款单位, 单位开户行, 单位帐号, 摘要, 结算方式, 结算号码, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id,
                 缴款, 找补, 缴款组id, 预交类别, 合作单位, 结算序号, 校对标志, 待转出, 结算性质, 会话号, 关联交易id)
                Select n_预交id, 记录性质, No, 记录状态, 病人id, 主页id, 科室id, 缴款单位, 单位开户行, 单位帐号, '医保挂号', v_结算方式, Null, 收款时间, 操作员编号,
                       操作员姓名, n_结算金额, 结帐id, 缴款, 找补, 缴款组id, 预交类别, 合作单位, 结算序号, 校对标志_In, 待转出, 结算性质, 会话号, n_关联id
                From 病人预交记录
                Where 结帐id = 结帐id_In And 记录性质 = 4 And Rownum < 2;
            End If;
          
            Delete 病人预交记录 a
            Where a.结帐id = 结帐id_In And a.记录性质 = 4 And Id <> Nvl(n_关联id, 0) And Nvl(a.卡类别id, 0) = 0 And
                  a.结算方式 Is Not Null And Exists (Select 1 From 结算方式 Where 性质 In (3, 4) And a.结算方式 = 名称);
            n_冲预交 := Nvl(n_冲预交, 0) - n_结算金额;
          Else
            Insert Into 病人预交记录
              (Id, 记录性质, No, 记录状态, 病人id, 主页id, 科室id, 缴款单位, 单位开户行, 单位帐号, 摘要, 结算方式, 结算号码, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id,
               缴款, 找补, 缴款组id, 预交类别, 卡类别id, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 校对标志, 待转出, 结算性质, 会话号, 关联交易id)
              Select 病人预交记录_Id.Nextval, 记录性质, No, 记录状态, 病人id, 主页id, 科室id, 缴款单位, 单位开户行, 单位帐号, '医保挂号', v_结算方式, Null, 收款时间,
                     操作员编号, 操作员姓名, n_结算金额, 结帐id, 缴款, 找补, 缴款组id, 预交类别, n_卡类别id, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位, 结算序号,
                     校对标志_In, 待转出, 结算性质, 会话号, n_关联id
              From 病人预交记录
              Where Nvl(关联交易id, 0) = Nvl(n_关联id, 0) And Nvl(卡类别id, 0) = Nvl(n_卡类别id, 0) And 结帐id = 结帐id_In And 记录性质 = 4 And
                    Rownum < 2;
            n_冲预交 := Nvl(n_冲预交, 0) - n_结算金额;
          End If;
          v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '|') + 1);
        End Loop;
        Update 保险结算明细 Set 标志 = 校对标志_In Where 结帐id = 结帐id_In;
      End If;
    End If;
  
    If Nvl(n_冲预交, 0) <> 0 Then
      --有未更新完的金额，累计到结算方式为NULL的记录中；或者新增了结算记录，从NULL中扣除
      Update 病人预交记录
      Set 冲预交 = Nvl(冲预交, 0) + n_冲预交
      Where 记录性质 = 4 And 结算方式 Is Null And 结帐id = 结帐id_In And 校对标志 = 1;
      If Sql%Notfound Then
        Insert Into 病人预交记录
          (Id, 记录性质, No, 记录状态, 病人id, 主页id, 科室id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 校对标志, 结算性质)
          Select 病人预交记录_Id.Nextval, 记录性质, No, 记录状态, 病人id, 主页id, 科室id, Null, 收款时间, 操作员编号, 操作员姓名, n_冲预交, 结帐id, 缴款组id, 1,
                 结算性质
          From 病人预交记录
          Where 记录性质 = 4 And 结帐id = 结帐id_In And Rownum < 2;
      End If;
      n_冲预交 := 0;
    End If;
  
    If Nvl(完成标志_In, 0) = 0 Then
      Return;
    End If;
  
    --1.先检查金额是否一致
    Select Nvl(Sum(实收金额), 0) Into n_结算金额 From 门诊费用记录 Where 结帐id = 结帐id_In;
    If n_结算金额 = 0 Then
      --0费用特殊处理
      Update 病人预交记录
      Set 结算方式 = v_现金, 校对标志 = 0
      Where 结算方式 Is Null And Nvl(冲预交, 0) = 0 And 结帐id = 结帐id_In;
    End If;
    Select Nvl(Sum(冲预交), 0) Into n_冲预交 From 病人预交记录 Where 结帐id = 结帐id_In;
    If n_结算金额 <> n_冲预交 Then
      v_Err_Msg := '结算信息有误，实收金额(' || n_结算金额 || ')与结算金额(' || n_冲预交 || ')不一致，不能完成结算！';
      Raise Err_Item;
    End If;
    Delete From 病人预交记录 Where 结算方式 Is Null And Nvl(冲预交, 0) = 0 And 结帐id = 结帐id_In;
    --2.检查是否存在未校对的记录
    If Nvl(n_收费作废, 0) = 1 Then
      Update 病人预交记录 Set 校对标志 = 0 Where 结帐id = 结帐id_In And 结算方式 Is Null;
    Else
      Select Count(1) Into n_Count From 病人预交记录 Where 结帐id = 结帐id_In And 结算方式 Is Null;
      If n_Count > 0 Then
        v_Err_Msg := '结算数据中还存在未结算的数据，不能完成结算！';
        Raise Err_Item;
      End If;
    End If;
    Select Count(1) Into n_Count From 病人预交记录 Where 结帐id = 结帐id_In And 校对标志 = 1 And Nvl(冲预交, 0) <> 0;
    If n_Count > 0 Then
      v_Err_Msg := '结算数据中还存在未校对支付方式，不能完成结算！';
      Raise Err_Item;
    End If;
  
    --3.处理预交记录的校对标志
    If Nvl(n_退费, 0) = 1 Then
      Select Max(是否电子票据) Into n_是否电子票据 From 病人预交记录 
       Where 结帐id In (Select 结帐ID From 门诊费用记录 Where no = 单据号_In And 记录性质 = 4 And 记录状态 = 3);
    Else
      n_是否电子票据 := 电子票据_In;
      If 电子票据_In Is Null Then
        Select Nvl(Max(险类), 0) Into n_险类 From 保险结算记录 Where 记录id = 结帐id_In And 性质 = 2;
        n_是否电子票据 := Zl_Fun_Isstarteinvoice(4, n_险类);
      End If;
    End If;
    Update 病人预交记录 Set 校对标志 = 0, 是否电子票据 = n_是否电子票据 Where 结帐id = 结帐id_In;
    If Nvl(n_收费作废, 0) = 1 Then
      Update 病人预交记录 Set 结算方式 = Null Where No = 单据号_In And 记录性质 = 4 And 记录状态 = 3 And 校对标志 = 1;
    End If;
  
    --4.更新费用状态,如果是对异常单据作废，则原始记录和退费记录的费用状态都应该为异常
    If Nvl(n_收费作废, 0) = 0 Then
      Update 门诊费用记录 Set 费用状态 = 0 Where 结帐id = 结帐id_In;
    End If;
  
    --5.更新挂号记录标志
    If Nvl(n_收费作废, 0) = 0 And Nvl(n_序号, 0) = 1 Then
      Update 病人挂号记录 Set 记录标志 = 0 Where 记录状态 = Decode(n_退号, 1, 2, 1) And No = 单据号_In;
    End If;
  
    --6.更新人员缴款数据,Not Exists中主要是针对三方卡挂号作废单原始单据结算成功了的，输入也调退费接口，但不能更新缴款余额
    If Nvl(n_收费作废, 0) = 0 Then
      For c_缴款 In (Select 结算方式, 操作员姓名, Sum(Nvl(a.冲预交, 0)) As 冲预交
                   From 病人预交记录 a
                   Where a.结帐id = 结帐id_In And Mod(a.记录性质, 10) <> 1 And 结算方式 Is Not Null And Not Exists
                    (Select 1
                          From 病人预交记录 b
                          Where b.No = a.No And b.记录性质 = a.记录性质 And b.记录状态 = 3 And b.关联交易id = a.关联交易id And
                                Nvl(b.校对标志, 0) <> 0)
                   Group By 结算方式, 操作员姓名) Loop
      
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) + Nvl(c_缴款.冲预交, 0)
        Where 收款员 = c_缴款.操作员姓名 And 性质 = 1 And 结算方式 = c_缴款.结算方式;
        If Sql%Rowcount = 0 Then
          Insert Into 人员缴款余额
            (收款员, 结算方式, 性质, 余额)
          Values
            (c_缴款.操作员姓名, c_缴款.结算方式, 1, Nvl(c_缴款.冲预交, 0));
        End If;
      End Loop;
    End If;
  Else
    If Nvl(完成标志_In, 0) = 0 Then
      Return;
    End If;
    Update 门诊费用记录 Set 费用状态 = 0 Where 记录性质 = 4 And No = 单据号_In;
    Update 病人挂号记录 Set 记录标志 = 0 Where No = 单据号_In;
  End If;

  If n_退号 = 1 Then
    --挂号必须在完成挂号时生成队列，退号在开始退号时就取消队列
    Open c_Registinfo(3);
    Fetch c_Registinfo
      Into r_Registrow;
  
    Update 病人挂号汇总
    Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - n_预约标志, 已约数 = Nvl(已约数, 0) - n_预约标志
    Where 日期 = Trunc(r_Registrow.发生时间) And 科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And
          (号码 = r_Registrow.号码 Or 号码 Is Null);
  
    If Sql%Rowcount = 0 Then
      Insert Into 病人挂号汇总
        (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 其中已接收, 已约数)
      Values
        (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
         Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, -1, -1 * n_预约标志, -1 * n_预约标志);
    End If;
  
    If 退号重用_In = 1 Or (退号重用_In = 2 And Trunc(r_Registrow.发生时间) <> Trunc(Sysdate)) Then
      Delete 挂号序号状态
      Where 状态 = 1 And
            (号码, 序号, 日期) = (Select 号别, 号序, Trunc(发生时间) From 病人挂号记录 Where No = 单据号_In And Rownum < 2) Or
            (号码, 序号, 日期) = (Select 号别, 号序, 发生时间 From 病人挂号记录 Where No = 单据号_In And Rownum < 2);
    Else
      Update 挂号序号状态
      Set 状态 = 4
      Where 状态 = 1 And
            (号码, 序号, 日期) = (Select 号别, 号序, Trunc(发生时间) From 病人挂号记录 Where No = 单据号_In And Rownum < 2) Or
            (号码, 序号, 日期) = (Select 号别, 号序, 发生时间 From 病人挂号记录 Where No = 单据号_In And Rownum < 2);
    End If;
  
    If n_记录id <> 0 Then
      If 退号重用_In = 1 Or (退号重用_In = 2 And Trunc(r_Registrow.发生时间) <> Trunc(Sysdate)) Then
        Update 临床出诊序号控制
        Set 挂号状态 = 0, 操作员姓名 = Null
        Where 挂号状态 = 1 And 记录id = n_记录id And 序号 = r_Registrow.号序;
      
        Update 临床出诊序号控制
        Set 挂号状态 = 4, 操作员姓名 = Null
        Where 挂号状态 = 1 And 记录id = n_记录id And 备注 = To_Char(r_Registrow.号序);
      Else
        Update 临床出诊序号控制
        Set 挂号状态 = 4, 操作员姓名 = r_Registrow.操作员姓名
        Where 挂号状态 = 1 And 记录id = n_记录id And (序号 = r_Registrow.号序 Or 备注 = To_Char(r_Registrow.号序));
      End If;
    
      Update 临床出诊记录
      Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - n_预约标志, 已约数 = Nvl(已约数, 0) - n_预约标志
      Where Id = n_记录id;
    End If;
  
    --医保产生的就诊登记记录
    Begin
      Delete From 就诊登记记录 Where 病人id = n_病人id And 就诊时间 = d_发生时间 And 主页id Is Null;
    Exception
      When Others Then
        Null;
    End;
  Elsif Nvl(n_退费, 0) = 0 Then
    --0-不产生队列;1-按医生或分诊台排队;2-先分诊,后医生站
    If Nvl(n_预约挂号, 0) = 1 Then
      n_预约生成队列 := Zl_To_Number(Zl_Getsysparameter('预约生成队列', 1113));
    End If;
  
    If Nvl(生成队列_In, 0) <> 0 And Nvl(n_预约挂号, 0) = 0 Or Nvl(n_预约生成队列, 0) = 1 Then
      For v_挂号 In (Select Id, 姓名, 诊室, 执行人, 执行部门id, 发生时间, 号别, 号序, 预约方式
                   From 病人挂号记录
                   Where No = 单据号_In) Loop
        n_分诊台签到排队 := Zl_To_Number(Zl_Getsysparameter('分诊台签到排队', 1113, 1, Nvl(v_挂号.执行部门id, 0)));
        If Nvl(n_分诊台签到排队, 0) = 0 Or Nvl(n_预约生成队列, 0) = 1 Then
          Begin
            Select 1,
                   Case
                     When 排队时间 < Trunc(Sysdate) Then
                      1
                     Else
                      0
                   End
            Into n_排队, n_当天排队
            From 排队叫号队列
            Where 业务类型 = 0 And 业务id = v_挂号.Id And Rownum <= 1;
          Exception
            When Others Then
              n_排队 := 0;
          End;
          If n_排队 = 0 Then
            If n_预约生成队列 = 1 Then
              If Nvl(n_记录id, 0) = 0 Then
                Select Decode(To_Char(d_发生时间, 'D'),
                               '1',
                               '周日',
                               '2',
                               '周一',
                               '3',
                               '周二',
                               '4',
                               '周三',
                               '5',
                               '周四',
                               '6',
                               '周五',
                               '7',
                               '周六',
                               Null)
                Into v_星期
                From Dual;
                Select Max(Id) Into n_安排id From 挂号安排 Where 号码 = v_号码;
                Select Max(Id)
                Into n_计划id
                From 挂号安排计划
                Where 安排id = n_安排id And 审核时间 Is Not Null And
                      Nvl(生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) =
                      (Select Max(a.生效时间) As 生效
                       From 挂号安排计划 a
                       Where a.审核时间 Is Not Null And d_发生时间 Between Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                             a.失效时间 And a.安排id = n_安排id) And
                      d_发生时间 Between Nvl(生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And 失效时间;
              
                If Nvl(n_计划id, 0) = 0 Then
                  Select Count(Rownum)
                  Into n_分时段
                  From 挂号安排时段
                  Where 星期 = v_星期 And 安排id = n_安排id And Rownum <= 1;
                Else
                  Select Count(Rownum)
                  Into n_分时段
                  From 挂号计划时段
                  Where 星期 = v_星期 And 计划id = n_计划id And Rownum <= 1;
                End If;
              Else
                Select Nvl(是否分时段, 0) Into n_分时段 From 临床出诊记录 Where Id = n_记录id;
              End If;
              n_分时点显示 := Nvl(Zl_To_Number(Zl_Getsysparameter(270)), 0);
              If n_分时点显示 = 1 And n_分时段 = 1 Then
                n_分时点显示 := 1;
              Else
                n_分时点显示 := Null;
              End If;
            End If;
            --产生队列
            --按”执行部门”产生队列
            n_挂号id   := v_挂号.Id;
            v_队列名称 := v_挂号.执行部门id;
            v_排队号码 := Zlgetnextqueue(v_挂号.执行部门id, n_挂号id, v_挂号.号别 || '|' || v_挂号.号序);
            v_排队序号 := Zlgetsequencenum(0, n_挂号id, 0);
          
            --挂号id_In,号码_In,号序_In,缺省日期_In,扩展_In(暂无用)
            d_排队时间 := Zl_Get_Queuedate(n_挂号id, v_挂号.号别, v_挂号.号序, v_挂号.发生时间);
            --   队列名称_In , 业务类型_In, 业务id_In,科室id_In,排队号码_In,排队标记_In,患者姓名_In,病人ID_IN, 诊室_In, 医生姓名_In,
            Zl_排队叫号队列_Insert(v_队列名称,
                             0,
                             n_挂号id,
                             v_挂号.执行部门id,
                             v_排队号码,
                             Null,
                             v_挂号.姓名,
                             n_病人id,
                             v_挂号.诊室,
                             v_挂号.执行人,
                             d_排队时间,
                             v_挂号.预约方式,
                             n_分时点显示,
                             v_排队序号);
          
            --挂号立即排队
            If Nvl(n_分诊台签到排队, 0) = 0 Then
              Update 病人挂号记录 Set 记录标志 = 1 Where Id = n_挂号id;
            End If;
          
          Elsif Nvl(n_当天排队, 0) = 1 Then
            --更新队列号
            v_排队号码 := Zlgetnextqueue(v_挂号.执行部门id, v_挂号.Id, v_挂号.号别 || '|' || Nvl(v_挂号.号序, 0));
            v_排队序号 := Zlgetsequencenum(0, v_挂号.Id, 1);
            --新队列名称_IN, 业务类型_In, 业务id_In , 科室id_In , 患者姓名_In , 诊室_In, 医生姓名_In ,排队号码_In
            Zl_排队叫号队列_Update(v_挂号.执行部门id,
                             0,
                             v_挂号.Id,
                             v_挂号.执行部门id,
                             v_挂号.姓名,
                             v_挂号.诊室,
                             v_挂号.执行人,
                             v_排队号码,
                             v_排队序号);
          
          Else
            --新队列名称_IN, 业务类型_In, 业务id_In , 科室id_In , 患者姓名_In , 诊室_In, 医生姓名_In ,排队号码_In
            Zl_排队叫号队列_Update(v_挂号.执行部门id,
                             0,
                             v_挂号.Id,
                             v_挂号.执行部门id,
                             v_挂号.姓名,
                             v_挂号.诊室,
                             v_挂号.执行人);
          End If;
        End If;
      End Loop;
    End If;
  End If;

  If 收回票据号_In Is Not Null And Nvl(n_退费, 0) = 1 Then
    --光退挂号费,不回收票据
    --退卡收回票据(可能上次挂号使用票据,不能收回)
    Begin
      --从最后一次打印的内容中取
      Select Id
      Into n_打印id
      From (Select b.Id
             From 票据使用明细 a, 票据打印内容 b
             Where a.打印id = b.Id And a.性质 = 1 And a.原因 In (1, 3) And b.数据性质 = 4 And b.No = 单据号_In
             Order By a.使用时间 Desc)
      Where Rownum < 2;
    Exception
      When Others Then
        n_打印id := Null;
    End;
  
    --先收回原票据
    If n_打印id Is Not Null Then
      Begin
        Insert Into 票据使用明细
          (Id, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
          Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_Date, v_操作员姓名, 票据金额
          From 票据使用明细
          Where 打印id = n_打印id And 性质 = 1 And Instr(',' || 收回票据号_In || ',', ',' || 号码 || ',') > 0;
      Exception
        When Others Then
          Delete From 票据使用明细 Where 打印id = n_打印id And 性质 = 2 And 原因 = 2;
          Insert Into 票据使用明细
            (Id, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
            Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_Date, v_操作员姓名, 票据金额
            From 票据使用明细
            Where 打印id = n_打印id And 性质 = 1 And Instr(',' || 收回票据号_In || ',', ',' || 号码 || ',') > 0;
      End;
    End If;
  End If;

  --消息推送
  If Nvl(n_退费, 0) = 0 Then
    Begin
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 1, n_挂号id;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_001(n_挂号id, 单据号_In);
  Else
    Begin
      Select Id Into n_挂号id From 病人挂号记录 Where No = 单据号_In And 记录状态 = 2 And Rownum < 2;
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 2, 单据号_In;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_003(n_挂号id, 单据号_In);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病人挂号收费_Modify;
/

Create Or Replace Procedure Zl_Third_Payment
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is

  --------------------------------------------------------------------------------------------------
  --功能:三方接口支付
  --入参:Xml_In:
  --<IN>
  --        <NO></NO>                       //收费单据号串,逗号分隔多个单据号
  --        <JE></JE>                       //总金额
  --        <BRID>病人ID</BRID>             //病人ID
  --        <XM>姓名</XM>                   //姓名
  --        <SFZH>身份证号</SFZH>           //身份证号
  --        <SFGH></SFGH>                   //是否挂号单
  --        <WCJE>误差额</WCJE>             //误差项不传时,以总金额-本次结算费用总额为准
  --        <JSMS>1</JSMS>          //结算模式：0-普通模式，1-异步结算模式
  --        <CZLX>0</CZLX>          //操作类型：结算模式为1时传入，0-开始结算，1-完成结算，2-回退结算
  --        <JZID>1</JZID>          //结帐ID，操作类型为1或2时传入
  --    <ZFBZH>支付宝公众号UserID</ZFBZH>
  --    <ZFBXCY>支付宝小程序UserID</ZFBXCY>
  --    <WXGZHID>微信公众号OpenID</WXGZH>
  --    <WXXCXID>微信小程序OpenID</WXXCXID>
  --        <JSLIST>          //结算列表，操作类型为2时可不传入
  --         <JS>
  --              <JSKLB>支付卡类别</JSKLB >
  --              <JSKH>支付卡号</ JSKH >
  --              <JSFS>支付方式</JSFS> //支付方式:现金;支票,如果是三方卡,可以传空
  --              <JSJE>支付金额</JSJE>
  --              <JYLSH>交易流水号</JYLSH>
  --              <JYSM>交易说明</JYSM>
  --              <ZY>摘要</ZY>
  --              <SFCYJ>是否冲预交</SFCYJ>  //允冲预交时,只填JSJE节点:1-冲预交
  --              <SFXFK>是否消费卡</SFXFK>  //(1-是消费卡),消费卡时,传入结算卡类别,结算卡号,结算金额等接点
  --              <DJH>S0000001</DJH> //分单据结算时传入
  --              <EXPENDLIST>  //扩展交易信息
  --                  <EXPEND>
  --                        <JYMC >交易名称</交易名称>
  --                        <JYLR>交易内容</JYLR>
  --                  </EXPEND>
  --              </EXPENDLIST>
  --         </JS>
  --       </JSLIST >
  --</IN>

  --出参:Xml_Out
  --  <OUTPUT>
  --    <JZID>结帐ID</JZID>       //结帐ID
  --    <CZSJ>操作时间</CZSJ>          //HIS的登记时间
  --  <KPBZ>开票标志</KPBZ> //1-成功开具电子票据;0-未开票成功标志
  --  <URL>H5页面URL</URL>
  --  <NETURL>外网H5页面URL</NETURL>
  --  <FPTT>发票抬头</FPTT>        //病人姓名
  --  <FPH>发票号</FPH>             //发票编号
  --  <FPJE>发票金额</FPJE>        //100.00
  --  <KPRQ>开票日期</KPRQ>   //yyyy-mm-dd
  --    ――如无下列错误结点则说明正确执行
  --    <ERROR>
  --      <MSG>错误信息</MSG>
  --    </ERROR>
  --  </OUTPUT>
  --------------------------------------------------------------------------------------------------
  v_Nos      Varchar2(4000);
  n_结算金额 门诊费用记录.实收金额%Type;

  n_卡类别id   医疗卡类别.Id%Type;
  n_结算卡序号 病人预交记录.结算卡序号%Type;
  v_结算方式   Varchar2(2000);
  n_病人id     门诊费用记录.病人id%Type;
  v_身份证号   病人信息.身份证号%Type;
  v_姓名       门诊费用记录.姓名%Type;
  v_性别       门诊费用记录.性别%Type;
  v_年龄       门诊费用记录.年龄%Type;
  n_结算模式   Number(1); --0-普通模式，1-异步结算模式
  n_操作类型   Number(1); --结算模式为1时传入，0 - 开始结算，1 - 完成结算，2 - 回退结算

  n_关联交易id 病人预交记录.关联交易id%Type;
  n_预交id     病人预交记录.Id%Type;
  n_消费卡     Number;
  n_删除原结算 Number;

  v_医疗付款方式编码 医疗付款方式.编码%Type;
  v_付款方式         医疗付款方式.名称%Type;
  v_操作员编码       门诊费用记录.操作员编号%Type;
  v_操作员姓名       门诊费用记录.操作员姓名%Type;
  n_结帐id           门诊费用记录.结帐id%Type;
  n_结帐金额         门诊费用记录.结帐金额%Type;
  d_收费时间         病人预交记录.收款时间%Type;
  n_消费卡id         消费卡信息.Id%Type;
  v_收费结算         Varchar2(2000);
  v_普通结算         Varchar2(4000);
  n_是否挂号         Number(3);
  n_预交支付         门诊费用记录.实收金额%Type;
  n_普通支付         门诊费用记录.实收金额%Type;
  v_结算卡号         病人预交记录.卡号%Type;
  v_交易流水号       病人预交记录.交易流水号%Type;
  v_交易说明         病人预交记录.交易说明%Type;
  v_摘要             病人预交记录.摘要%Type;
  n_科室id           挂号安排.科室id%Type;
  n_项目id           挂号安排.项目id%Type;
  n_医生id           挂号安排.医生id%Type;
  v_医生姓名         挂号安排.医生姓名%Type;
  v_号码             挂号安排.号码%Type;
  n_门诊号           病人信息.门诊号%Type;
  d_发生时间         病人挂号记录.发生时间%Type;
  v_费别             病人信息.费别%Type;
  n_号序             病人挂号记录.号序%Type;
  v_Para             Varchar2(500);
  n_挂号模式         Number(3);
  d_启用时间         Date;
  v_临时结算方式     病人预交记录.结算方式%Type;
  n_出诊记录id       临床出诊记录.Id%Type;
  n_序号             门诊费用记录.序号%Type;
  v_附加项目id       Varchar2(500);
  v_附加内容         Varchar2(500);
  v_附加值           Varchar2(100);
  n_Cursor           Number(3);
  n_实收金额         门诊费用记录.实收金额%Type;
  v_实收             Varchar2(500);
  n_从属父号         门诊费用记录.从属父号%Type;
  n_病人科室id       门诊费用记录.病人科室id%Type;
  n_执行部门id       门诊费用记录.执行部门id%Type;
  v_No               门诊费用记录.No%Type;
  v_普通等级         Varchar2(100);
  v_Pricegrade       Varchar2(500);
  n_医保支付         病人预交记录.冲预交%Type;
  n_Exists           Number;
  v_站点             部门表.站点%Type;
  n_结算序号         病人预交记录.结算序号%Type;
  n_业务类型         三方交易记录.业务类型%Type;
  v_Temp             Varchar2(32767); --临时XML
  x_Templet          Xmltype; --模板XML
  v_卡类别           三方交易记录.类别%Type;
  v_操作员           门诊费用记录.操作员姓名%Type;
  v_发药窗口         Varchar2(4000);
  n_误差额           病人预交记录.冲预交%Type;
  n_连续更新         Number;
  n_实名制           Number(3);
  n_认证             Number(3);
  n_Step             Number(2);
  n_Checkmzlg        Number(2);
  n_Count            Number(2);

  n_是否电子票据       病人预交记录.是否电子票据%Type;
  v_支付宝公众号userid Varchar2(100);
  v_支付宝小程序userid Varchar2(100);
  v_微信公众号openid   Varchar2(100);
  v_微信小程序openid   Varchar2(100);
  n_开票标志           Number(2);

  v_开票日期 Varchar2(20);
  v_患者姓名 电子票据使用记录.姓名%Type;
  v_发票编号 电子票据使用记录.号码%Type;
  n_发票金额 电子票据使用记录.票据金额%Type;
  v_Url      电子票据使用记录.Url内网%Type;
  v_Url外网  电子票据使用记录.Url外网%Type;

  Type Price_Type Is Record(
    项目id 门诊费用记录.收费细目id%Type,
    数次   门诊费用记录.数次%Type,
    单价   门诊费用记录.标准单价%Type,
    应收   门诊费用记录.应收金额%Type,
    实收   门诊费用记录.实收金额%Type); --定义Price记录类型 
  Type Price_Type_Array Is Table Of Price_Type Index By Binary_Integer; --定义存放Price记录的数组类型 
  Price_Rec       Price_Type; --声明变量，类型：Price记录类型
  Price_Rec_Array Price_Type_Array; --声明变量，类型：存放Price记录的数组类型

  v_Err_Msg Varchar2(200);
  Err_Item    Exception;
  Err_Special Exception;

  --获取卡类别名称
  Function Get_Cardname
  (
    卡类别_In Varchar2,
    消费卡_In Number
  ) Return Varchar2 As
    v_名称       医疗卡类别.名称%Type;
    n_By_Id_Find Number;
  
    v_Err_Msg Varchar2(200);
    Err_Item Exception;
  Begin
    If 卡类别_In Is Null Then
      Return Null;
    End If;
  
    Select Decode(Translate(Nvl(卡类别_In, 'abcd'), '#1234567890', '#'), Null, 1, 0) Into n_By_Id_Find From Dual;
  
    --1.按卡类别ID查找医疗卡
    If Nvl(消费卡_In, 0) = 0 And Nvl(n_By_Id_Find, 0) = 1 Then
      Begin
        Select 名称 Into v_名称 From 医疗卡类别 Where ID = To_Number(卡类别_In);
      Exception
        When Others Then
          v_Err_Msg := '未找到指定的结算支付信息！';
          Raise Err_Item;
      End;
    End If;
  
    --2.按卡名称查找医疗卡
    If Nvl(消费卡_In, 0) = 0 And Nvl(n_By_Id_Find, 0) = 0 Then
      Begin
        Select 名称 Into v_名称 From 医疗卡类别 Where 名称 = 卡类别_In;
      Exception
        When Others Then
          v_Err_Msg := 卡类别_In || '不存在!';
          Raise Err_Item;
      End;
    End If;
  
    --3.按卡类别ID查找消费卡
    If Nvl(消费卡_In, 0) = 1 And Nvl(n_By_Id_Find, 0) = 1 Then
      Begin
        Select 名称 Into v_名称 From 消费卡类别目录 Where 编号 = To_Number(卡类别_In);
      Exception
        When Others Then
          v_Err_Msg := '未找到指定的结算支付信息！';
          Raise Err_Item;
      End;
    End If;
  
    --3.按卡名称查找消费卡
    If Nvl(消费卡_In, 0) = 1 And Nvl(n_By_Id_Find, 0) = 0 Then
      Begin
        Select 名称 Into v_名称 From 消费卡类别目录 Where 名称 = 卡类别_In;
      Exception
        When Others Then
          v_Err_Msg := 卡类别_In || '不存在!';
          Raise Err_Item;
      End;
    End If;
  
    Return v_名称;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  End;

  --获取卡类别ID
  Function Get_Cardtypeid
  (
    卡类别_In    Varchar2,
    消费卡_In    Number,
    结算方式_Out In Out 医疗卡类别.结算方式%Type
  ) Return Number As
    n_卡类别id 医疗卡类别.Id%Type;
    v_名称     医疗卡类别.名称%Type;
    n_启用     医疗卡类别.是否启用%Type;
    v_结算方式 医疗卡类别.结算方式%Type;
  
    n_By_Id_Find Number;
    v_Err_Msg    Varchar2(200);
    Err_Item Exception;
  Begin
    If 卡类别_In Is Null Then
      Return 0;
    End If;
  
    Select Decode(Translate(Nvl(卡类别_In, 'abcd'), '#1234567890', '#'), Null, 1, 0) Into n_By_Id_Find From Dual;
  
    --1.按卡类别ID查找医疗卡
    If Nvl(消费卡_In, 0) = 0 And Nvl(n_By_Id_Find, 0) = 1 Then
      Begin
        Select ID, 结算方式, 名称, 是否启用
        Into n_卡类别id, v_名称, v_结算方式, n_启用
        From 医疗卡类别
        Where ID = To_Number(卡类别_In);
      Exception
        When Others Then
          v_Err_Msg := '未找到指定的结算支付信息！';
          Raise Err_Item;
      End;
    End If;
  
    --2.按卡名称查找医疗卡
    If Nvl(消费卡_In, 0) = 0 And Nvl(n_By_Id_Find, 0) = 0 Then
      Begin
        Select ID, 结算方式, 名称, 是否启用
        Into n_卡类别id, v_名称, v_结算方式, n_启用
        From 医疗卡类别
        Where 名称 = 卡类别_In;
      Exception
        When Others Then
          v_Err_Msg := 卡类别_In || '不存在!';
          Raise Err_Item;
      End;
    End If;
  
    --3.按卡类别ID查找消费卡
    If Nvl(消费卡_In, 0) = 1 And Nvl(n_By_Id_Find, 0) = 1 Then
      Begin
        Select 编号, 结算方式, 名称, 启用
        Into n_卡类别id, v_名称, v_结算方式, n_启用
        From 消费卡类别目录
        Where 编号 = To_Number(卡类别_In);
      Exception
        When Others Then
          v_Err_Msg := '未找到指定的结算支付信息！';
          Raise Err_Item;
      End;
    End If;
  
    --3.按卡名称查找消费卡
    If Nvl(消费卡_In, 0) = 1 And Nvl(n_By_Id_Find, 0) = 0 Then
      Begin
        Select 编号, 结算方式, 名称, 启用
        Into n_卡类别id, v_名称, v_结算方式, n_启用
        From 消费卡类别目录
        Where 名称 = 卡类别_In;
      Exception
        When Others Then
          v_Err_Msg := 卡类别_In || '不存在!';
          Raise Err_Item;
      End;
    End If;
  
    If Nvl(n_启用, 0) = 0 Then
      v_Err_Msg := v_名称 || '未启用，不允许进行缴费！';
      Raise Err_Item;
    End If;
  
    If 结算方式_Out Is Null Then
      If v_结算方式 Is Null Then
        v_Err_Msg := v_名称 || '未设置结算方式，不允许进行缴费！';
        Raise Err_Item;
      End If;
    
      结算方式_Out := v_结算方式;
    End If;
  
    Return n_卡类别id;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  End;

  Procedure Thirdcard_Balance
  (
    病人id_In     病人预交记录.病人id%Type,
    结帐id_In     病人预交记录.结帐id%Type,
    结算方式_In   病人预交记录.结算方式%Type,
    卡类别_In     病人预交记录.卡类别id%Type,
    卡号_In       病人预交记录.卡号%Type,
    支付金额_In   病人预交记录.冲预交%Type,
    交易流水号_In 病人预交记录.交易流水号%Type,
    交易说明_In   病人预交记录.交易说明%Type,
    No_In         病人预交记录.No%Type,
    关联交易id_In 病人预交记录.关联交易id%Type,
    Xmlexpned_In  Xmltype,
    结算模式_In   Number := 0,
    操作类型_In   Number := 0,
    删除原结算_In Number := 0
  ) Is
    --入参：
    --         结算模式_in   0-普通模式，1-异步结算模式
    --        操作类型_in   结算模式为1时传入，0 - 开始结算，1 - 完成结算，2 - 回退结算
    --        删除原结算_in 操作类型_In为1时有效，多个结算方式时调用多次该过程
    v_收费结算 Varchar2(2000);
    n_校对标志 病人预交记录.校对标志%Type;
  Begin
    If Nvl(结算模式_In, 0) = 1 And Nvl(操作类型_In, 0) = 0 Then
      n_校对标志 := 1;
    Else
      n_校对标志 := 2;
    End If;
  
    --结算方式|结算金额|结算号码|结算摘要|单据号|是否普通结算
    v_收费结算 := 结算方式_In || '|' || 支付金额_In || '| | |' || No_In || '|0';
    Zl_门诊收费结算_Modify(4, 病人id_In, 结帐id_In, v_收费结算, 0, 0, 卡类别_In, 卡号_In, 交易流水号_In, 交易说明_In, 0, 0, 0, 0, Null, Null, 0,
                     关联交易id_In, 删除原结算_In, n_校对标志);
  
    If Nvl(结算模式_In, 0) = 1 And Nvl(操作类型_In, 0) = 0 Then
      Return;
    End If;
  
    --保存扩展结算信息 
    For c_扩展 In (Select Extractvalue(j.Column_Value, '/EXPEND/JYMC') As Jymc,
                        Extractvalue(j.Column_Value, '/EXPEND/JYLR') As Jylr
                 From Table(Xmlsequence(Extract(Xmlexpned_In, '/EXPENDLIST/EXPEND'))) J) Loop
      Zl_三方结算交易_Insert(卡类别_In, 0, 卡号_In, 结帐id_In, c_扩展.Jymc || '|' || c_扩展.Jylr);
    End Loop;
  End Thirdcard_Balance;

  Procedure Squarecard_Balance
  (
    病人id_In     病人预交记录.病人id%Type,
    结帐id_In     病人预交记录.结帐id%Type,
    卡类别_In     病人预交记录.卡类别id%Type,
    卡号_In       病人预交记录.卡号%Type,
    支付金额_In   病人预交记录.冲预交%Type,
    交易流水号_In 病人预交记录.交易流水号%Type,
    交易说明_In   病人预交记录.交易说明%Type,
    Xmlexpned_In  Xmltype
  ) Is
    v_收费结算 Varchar2(2000);
    n_消费卡id 消费卡信息.Id%Type;
  Begin
    Select ID
    Into n_消费卡id
    From 消费卡信息
    Where 接口编号 = 卡类别_In And 卡号 = 卡号_In And
          序号 = (Select Max(序号) From 消费卡信息 Where 接口编号 = 卡类别_In And 卡号 = 卡号_In);
  
    --结算方式_IN格式为:卡类别ID|卡号|消费卡ID|消费金额||....
    v_收费结算 := 卡类别_In || '|' || 卡号_In || '|' || n_消费卡id || '|' || 支付金额_In;
    Zl_门诊收费结算_Modify(3, 病人id_In, 结帐id_In, v_收费结算, 0, 0, 卡类别_In, 卡号_In, 交易流水号_In, 交易说明_In, 0, 0, 0, 0);
  
    --保存扩展结算信息
    For c_扩展 In (Select Extractvalue(j.Column_Value, '/EXPEND/JYMC') As Jymc,
                        Extractvalue(j.Column_Value, '/EXPEND/JYLR') As Jylr
                 From Table(Xmlsequence(Extract(Xmlexpned_In, '/EXPENDLIST/EXPEND'))) J) Loop
      Zl_三方结算交易_Insert(卡类别_In, 1, 卡号_In, 结帐id_In, c_扩展.Jymc || '|' || c_扩展.Jylr, 0);
    End Loop;
  End Squarecard_Balance;

Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/NO'), To_Number(Extractvalue(Value(A), 'IN/BRID')),
         To_Number(Extractvalue(Value(A), 'IN/JE')), To_Number(Extractvalue(Value(A), 'IN/WCJE')),
         To_Number(Extractvalue(Value(A), 'IN/SFGH')), Extractvalue(Value(A), 'IN/ZD'),
         Extractvalue(Value(A), 'IN/SFZH'), Extractvalue(Value(A), 'IN/XM'), Extractvalue(Value(A), 'IN/JSMS'),
         Extractvalue(Value(A), 'IN/CZLX'), Extractvalue(Value(A), 'IN/JZID'), Extractvalue(Value(A), 'IN/ZFBZH'),
         Extractvalue(Value(A), 'IN/ZFBXCY'), Extractvalue(Value(A), 'IN/WXGZHID'), Extractvalue(Value(A), 'IN/WXXCXID')
  Into v_Nos, n_病人id, n_结算金额, n_误差额, n_是否挂号, v_站点, v_身份证号, v_姓名, n_结算模式, n_操作类型, n_结帐id, v_支付宝公众号userid, v_支付宝小程序userid,
       v_微信公众号openid, v_微信小程序openid
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If Nvl(n_病人id, 0) = 0 And Not v_身份证号 Is Null And Not v_姓名 Is Null Then
    n_病人id := Zl_Third_Getpatiid(v_身份证号, v_姓名);
  End If;

  --0.相关检查
  If Nvl(n_病人id, 0) = 0 Then
    v_Err_Msg := '不能有效识别病人身份,不允许缴费!';
    Raise Err_Item;
  End If;

  If Not v_支付宝公众号userid Is Null Then
    Zl_病人信息从表_Update(n_病人id, Upper('支付宝公众号UserID'), v_支付宝公众号userid);
  End If;

  If Not v_支付宝小程序userid Is Null Then
    Zl_病人信息从表_Update(n_病人id, Upper('支付宝小程序UserID'), v_支付宝公众号userid);
  End If;

  If Not v_微信公众号openid Is Null Then
    Zl_病人信息从表_Update(n_病人id, Upper('微信公众号OpenID'), v_支付宝公众号userid);
  End If;

  If Not v_微信小程序openid Is Null Then
    Zl_病人信息从表_Update(n_病人id, Upper('微信小程序OpenID'), v_支付宝公众号userid);
  End If;

  If v_Nos Is Null Then
    v_Err_Msg := '没有指定相关的收费单据,不允许缴费!';
    Raise Err_Item;
  End If;

  If Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) <> 0 Then
    If Nvl(n_结帐id, 0) = 0 Then
      v_Err_Msg := '没有指定相关的结算数据！';
      Raise Err_Item;
    End If;
  
    Begin
      Select 收款时间
      Into d_收费时间
      From 病人预交记录
      Where 结帐id = n_结帐id And Nvl(校对标志, 0) = 1 And Rownum < 2;
    Exception
      When Others Then
        v_Err_Msg := '没有找到指定的相关结算数据，可能已被处理！';
        Raise Err_Item;
    End;
  End If;

  If Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) = 2 Then
    If Nvl(n_是否挂号, 0) = 0 Then
      --删除结算数据，恢复划价单
      Zl_病人结算记录_Delete(n_结帐id);
      Zl_门诊收费结算_Cancel(n_结帐id);
    Else
      Zl_病人挂号记录_Cancel(n_结帐id);
    End If;
  
    v_Temp := '<CZSJ>' || To_Char(d_收费时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<JZID>' || n_结帐id || '</JZID>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<KPBZ>' || 0 || '</KPBZ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<URL>' || '' || '</URL>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<NETURL>' || '' || '</NETURL>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<FPTT>' || v_姓名 || '</FPTT>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<FPH>' || '' || '</FPH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<FPJE>' || '' || '</FPJE>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<KPRQ>' || '' || '</KPRQ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    Xml_Out := x_Templet;
    Return;
  End If;

  --人员id,人员编号,人员姓名
  v_Temp := Zl_Identity(1);
  If Nvl(v_Temp, '0') = '0' Or Nvl(v_Temp, '_') = '_' Then
    v_Err_Msg := '系统不能认别有效的操作员,不允许缴费!';
    Raise Err_Item;
  End If;
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_操作员编码 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_操作员姓名 := v_Temp;

  Begin
    Select b.编码, b.名称, a.姓名, a.性别, a.年龄
    Into v_医疗付款方式编码, v_付款方式, v_姓名, v_性别, v_年龄
    From 病人信息 A, 医疗付款方式 B
    Where a.医疗付款方式 = b.名称(+) And a.病人id = n_病人id;
  Exception
    When Others Then
      v_Err_Msg := '指定的缴费单据中不能有效识别病人,不允许缴费!';
      Raise Err_Item;
  End;

  n_Checkmzlg := To_Number(Nvl(zl_GetSysParameter(323), '0'));
  Select Decode(Nvl(n_是否挂号, 0), 0, 3, 4) Into n_业务类型 From Dual;
  If Nvl(n_结算模式, 0) = 0 Or Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) = 0 Then
    For c_交易记录 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                          Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                          Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                          Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号, Extractvalue(b.Column_Value, '/JS/ZY') As 摘要,
                          Extractvalue(b.Column_Value, '/JS/SFXFK') As 是否消费卡,
                          Extractvalue(b.Column_Value, '/JS/SFCYJ') As 是否冲预交
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
    
      If Nvl(c_交易记录.是否冲预交, 0) = 0 Then
        If c_交易记录.结算卡类别 Is Null Then
          v_卡类别 := c_交易记录.结算方式;
        Else
          v_卡类别 := Get_Cardname(c_交易记录.结算卡类别, c_交易记录.是否消费卡);
        End If;
      
        If v_卡类别 Is Null Then
          v_Err_Msg := '不支持的结算方式,请检查！';
          Raise Err_Item;
        End If;
      
        --仅第一个结算方式才检查交易锁
        n_Step := Nvl(n_Step, 0) + 1;
        If Zl_Fun_三方交易记录_Locked(v_卡类别, c_交易记录.交易流水号, c_交易记录.结算卡号, c_交易记录.摘要, n_业务类型) = 0 And n_Step = 1 Then
          v_Err_Msg := '交易流水号为:' || c_交易记录.交易流水号 || '的交易正在进行中，不允许再次提交此交易!';
          Raise Err_Special;
        End If;
      Else
        If Nvl(n_Checkmzlg, 0) <> 0 Then
          Select Count(1)
          Into n_Count
          From 病案主页 A, 病人信息 B
          Where a.病人id = n_病人id And a.病人性质 = 1 And a.病人id = b.病人id And a.主页id = b.主页id And Nvl(b.在院, 0) = 1;
          If n_Count <> 0 Then
            v_Err_Msg := '门诊留观病人不能使用门诊预交！';
            Raise Err_Item;
          End If;
        End If;
      End If;
    End Loop;
  End If;

  --费用单据
  If Nvl(n_是否挂号, 0) = 0 Then
    If Nvl(n_结算模式, 0) = 0 Or Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) = 0 Then
      --1.进行费用收费处理
      --获取发药窗口
      v_发药窗口 := Zl_Getclinicchargepaywins(v_Nos);
    
      Select 病人结帐记录_Id.Nextval, Sysdate Into n_结帐id, d_收费时间 From Dual;
    
      n_结帐金额 := 0;
      For c_缴费单 In (Select /*+ rule */
                     a.No, Max(a.开单部门id) As 开单部门id, Max(a.病人科室id) As 病人科室id, Max(a.病人id) As 病人id, Sum(实收金额) As 实收金额,
                     Max(a.开单人) As 开单人
                    From 门诊费用记录 A, Table(f_Str2List(v_Nos)) J
                    Where a.记录性质 = 1 And a.No = j.Column_Value
                    Group By a.No) Loop
        If Nvl(c_缴费单.病人id, 0) <> n_病人id Then
          v_Err_Msg := '缴费单据:' || c_缴费单.No || '与当前病人身份不符,不允许缴费!';
          Raise Err_Item;
        End If;
      
        n_结帐金额 := n_结帐金额 + Nvl(c_缴费单.实收金额, 0);
        Zl_病人划价收费_Insert(c_缴费单.No, n_病人id, 1, v_医疗付款方式编码, v_姓名, v_性别, v_年龄, c_缴费单.病人科室id, c_缴费单.开单部门id, c_缴费单.开单人,
                         n_结帐id, d_收费时间, v_操作员编码, v_操作员姓名, v_发药窗口, 0, d_收费时间);
      End Loop;
    
      --检查总金额是否正确
      If Nvl(n_误差额, 0) = 0 Then
        n_误差额 := Nvl(n_结帐金额, 0) - Nvl(n_结算金额, 0);
        If Abs(n_误差额) > 1.00 Then
          v_Err_Msg := '单据缴款金额与实际结算的误差太大!';
          Raise Err_Item;
        End If;
      End If;
      If Nvl(n_结帐金额, 0) <> Nvl(n_结算金额, 0) + Nvl(n_误差额, 0) Then
        v_Err_Msg := '指定的缴费单据的总金额不对,请重新选择缴费单据!';
        Raise Err_Item;
      End If;
    End If;
  
    --2.确定支付方式 
    n_结算序号   := -1 * n_结帐id;
    n_删除原结算 := 0;
    If Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) = 1 Then
      n_删除原结算 := 1;
    End If;
    For c_结算方式 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                          Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号,
                          Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                          Extractvalue(b.Column_Value, '/JS/JSJE') As 结算金额,
                          Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                          Extractvalue(b.Column_Value, '/JS/JYSM') As 交易说明, Extractvalue(b.Column_Value, '/JS/ZY') As 摘要,
                          Extractvalue(b.Column_Value, '/JS/SFCYJ') As 是否冲预交,
                          Extractvalue(b.Column_Value, '/JS/SFXFK') As 是否消费卡,
                          Extractvalue(b.Column_Value, '/JS/DJH') As 单据号,
                          Extract(b.Column_Value, '/JS/EXPENDLIST') As Expend
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
    
      n_消费卡   := Nvl(c_结算方式.是否消费卡, 0);
      v_结算方式 := c_结算方式.结算方式;
      --1.三方卡结算
      If c_结算方式.结算卡类别 Is Not Null And n_消费卡 = 0 Then
        n_卡类别id := Get_Cardtypeid(c_结算方式.结算卡类别, 0, v_结算方式);
        Select Max(关联交易id)
        Into n_关联交易id
        From 病人预交记录
        Where 结帐id = n_结帐id And 卡类别id = n_卡类别id And Rownum < 2;
        If Nvl(n_关联交易id, 0) = 0 Then
          Select 病人预交记录_Id.Nextval Into n_关联交易id From Dual;
        End If;
      
        Thirdcard_Balance(n_病人id, n_结帐id, v_结算方式, n_卡类别id, c_结算方式.结算卡号, c_结算方式.结算金额, c_结算方式.交易流水号, c_结算方式.交易说明,
                          c_结算方式.单据号, n_关联交易id, c_结算方式.Expend, n_结算模式, n_操作类型, n_删除原结算);
      
        n_删除原结算 := 0;
        If Nvl(n_结算模式, 0) = 0 Or Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) = 1 Then
          If c_结算方式.结算卡类别 Is Null Then
            v_卡类别 := v_结算方式;
          Else
            v_卡类别 := Get_Cardname(c_结算方式.结算卡类别, n_消费卡);
          End If;
          Update 三方交易记录
          Set 业务结算id = n_结算序号
          Where 流水号 = c_结算方式.交易流水号 And 类别 = v_卡类别 And 业务类型 = n_业务类型;
        End If;
      
        --完成结算时才处理非三方卡结算的
      Elsif Nvl(n_结算模式, 0) = 0 Or Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) = 1 Then
        --2.消费卡结算                           
        If c_结算方式.结算卡类别 Is Not Null And n_消费卡 = 1 Then
          n_卡类别id := Get_Cardtypeid(c_结算方式.结算卡类别, 1, v_结算方式);
          Squarecard_Balance(n_病人id, n_结帐id, n_卡类别id, c_结算方式.结算卡号, c_结算方式.结算金额, c_结算方式.交易流水号, c_结算方式.交易说明,
                             c_结算方式.Expend);
        
          --3.冲预交款
        Elsif Nvl(c_结算方式.是否冲预交, 0) = 1 Then
          Zl_门诊收费结算_Modify(0, n_病人id, n_结帐id, Null, c_结算方式.结算金额, 0, Null, Null, Null, Null, 0, 0, 0, 0);
        
          --4.普通结算
        Else
          If v_结算方式 Is Null Then
            v_Err_Msg := '未指定支付方式，不允许缴款!';
            Raise Err_Item;
          End If;
        
          --结算方式|结算金额|结算号码|结算摘要||..
          v_收费结算 := v_结算方式 || '|' || c_结算方式.结算金额 || '| | ';
          v_普通结算 := v_普通结算 || '||' || v_收费结算;
        End If;
      
        If c_结算方式.结算卡类别 Is Null Then
          v_卡类别 := v_结算方式;
        Else
          v_卡类别 := Get_Cardname(c_结算方式.结算卡类别, n_消费卡);
        End If;
        Update 三方交易记录
        Set 业务结算id = n_结算序号
        Where 流水号 = c_结算方式.交易流水号 And 类别 = v_卡类别 And 业务类型 = n_业务类型;
      End If;
    End Loop;
  
    If Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) = 0 Then
      v_Temp := '<CZSJ>' || To_Char(d_收费时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
      v_Temp := '<JZID>' || n_结帐id || '</JZID>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
      v_Temp := '<KPBZ>' || 0 || '</KPBZ>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    
      v_Temp := '<URL>' || '' || '</URL>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    
      v_Temp := '<NETURL>' || '' || '</NETURL>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    
      v_Temp := '<FPTT>' || v_姓名 || '</FPTT>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    
      v_Temp := '<FPH>' || '' || '</FPH>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
      v_Temp := '<FPJE>' || '' || '</FPJE>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
      v_Temp := '<KPRQ>' || '' || '</KPRQ>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    
      Xml_Out := x_Templet;
      Return;
    End If;
  
    --5.普通结算及完成结算
    If v_普通结算 Is Not Null Then
      v_普通结算 := Substr(v_普通结算, 3);
    End If;
  
    --6.电子票据处理
    n_是否电子票据 := b_Einvoice_Request.Einvoice_Start(1, Null);
    Zl_门诊收费结算_Modify(0, n_病人id, n_结帐id, v_普通结算, Null, 0, Null, Null, Null, Null, 0, 0, n_误差额, 1, Null, Null, 1, Null, 0,
                     0, 0, n_是否电子票据);
    If Nvl(n_是否电子票据, 0) = 1 Then
      If b_Einvoice_Request.Einvoice_Create(1, n_结帐id, Null, v_Err_Msg) = 0 Then
        --电子票据开具成功
        Raise Err_Item;
      End If;
    
      Select Max(1), Max(姓名), Max(号码), Max(To_Char(生成时间, 'yyyy-mm-dd')), Max(Url内网), Max(Url外网), Max(票据金额)
      Into n_开票标志, v_患者姓名, v_发票编号, v_开票日期, v_Url, v_Url外网, n_发票金额
      From 电子票据使用记录
      Where 结算id = n_结帐id And 票种 = 1 And 记录状态 = 1;
    
      If v_患者姓名 Is Not Null Then
        v_姓名 := v_患者姓名;
      End If;
    
    End If;
  
    v_Temp := '<CZSJ>' || To_Char(d_收费时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<JZID>' || n_结帐id || '</JZID>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<KPBZ>' || Nvl(n_开票标志, 0) || '</KPBZ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<URL>' || Nvl(v_Url, '') || '</URL>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<NETURL>' || Nvl(v_Url外网, '') || '</NETURL>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<FPTT>' || v_姓名 || '</FPTT>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<FPH>' || v_发票编号 || '</FPH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<FPJE>' || Nvl(n_发票金额, 0) || '</FPJE>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<KPRQ>' || v_开票日期 || '</KPRQ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    Xml_Out := x_Templet;
    Return;
  End If;

  --=================================================================================
  --挂号单据 
  n_结帐金额 := 0;
  v_Para     := zl_GetSysParameter(256);
  n_挂号模式 := Substr(v_Para, 1, 1);
  Begin
    d_启用时间 := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
  Exception
    When Others Then
      d_启用时间 := Null;
  End;

  --实名制检查
  n_实名制 := To_Number(Nvl(zl_GetSysParameter(319), '0'));
  If n_实名制 = 1 Then
    Select Count(1) Into n_认证 From 病人实名信息 Where 病人id = n_病人id And Rownum < 2;
    If n_认证 = 0 Then
      v_Err_Msg := '病人未实名认证，不能挂号。';
      Raise Err_Item;
    End If;
  End If;

  Begin
    Select a.执行部门id, a.收费细目id, c.Id, a.执行人, b.号别, b.门诊号, b.发生时间, a.费别, b.号序, b.出诊记录id
    Into n_科室id, n_项目id, n_医生id, v_医生姓名, v_号码, n_门诊号, d_发生时间, v_费别, n_号序, n_出诊记录id
    From 门诊费用记录 A, 病人挂号记录 B, 人员表 C
    Where a.No = v_Nos And a.记录性质 = 4 And a.序号 = 1 And a.No = b.No And a.执行人 = c.姓名(+);
  Exception
    When Others Then
      v_Err_Msg := '没有找到指定的单据数据！';
      Raise Err_Item;
  End;

  --预约接收
  If n_挂号模式 = 1 Then
    If d_启用时间 > d_发生时间 And n_出诊记录id Is Null Then
      n_挂号模式 := 0;
    End If;
  End If;

  If Nvl(n_结算模式, 0) = 0 Or Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) = 0 Then
    v_Pricegrade := Zl_Get_Pricegrade(v_站点);
    v_普通等级   := Substr(v_Pricegrade, 1, Instr(v_Pricegrade, '|') - 1);
  
    Select 病人结帐记录_Id.Nextval, Sysdate Into n_结帐id, d_收费时间 From Dual;
  
    For c_费用 In (Select 1 As 顺序号, b.No, b.收据费目, b.结帐id, b.执行部门id, b.病人科室id, b.开单人, b.收费类别, b.收入项目id, b.附加标志,
                        To_Char(b.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间, b.价格父号, b.从属父号, b.序号, b.收费细目id, b.计算单位,
                        Max(m.名称) As 名称, Max(m.规格) As 规格, Sum(b.标准单价) As 单价, Avg(Nvl(b.付数, 1) * b.数次) As 数量,
                        Sum(b.应收金额) As 应收金额, Sum(b.实收金额) As 实收金额, Max(j.名称) As 开单科室, Max(q.名称) As 执行科室
                 From 门诊费用记录 B, 收费项目目录 M, 部门表 J, 部门表 Q
                 Where b.No = v_Nos And b.记录性质 = 4 And Nvl(b.费用状态, 0) = 0 And b.收费细目id = m.Id And b.开单部门id = j.Id(+) And
                       b.执行部门id = q.Id(+)
                 Group By b.No, b.收据费目, b.结帐id, b.执行部门id, b.病人科室id, b.开单人, b.收入项目id, b.收费类别, b.登记时间, b.价格父号, b.从属父号, b.序号,
                          b.收费细目id, b.计算单位, b.附加标志
                 Order By 序号) Loop
    
      Zl_病人预约挂号记录_Update(c_费用.No, c_费用.序号, c_费用.价格父号, c_费用.从属父号, c_费用.收费类别, c_费用.收费细目id, c_费用.数量, c_费用.单价, c_费用.收入项目id,
                         c_费用.收据费目, c_费用.应收金额, c_费用.实收金额, c_费用.附加标志, Null, Null, Null, Null, c_费用.病人科室id, c_费用.执行部门id);
    
      n_结帐金额   := n_结帐金额 + c_费用.实收金额;
      n_序号       := c_费用.序号;
      n_病人科室id := c_费用.病人科室id;
      n_执行部门id := c_费用.执行部门id;
      v_No         := c_费用.No;
    End Loop;
  
    Begin
      Select Zl_Fun_Customregexpenses(n_病人id, 0, v_号码, v_姓名, v_性别, v_年龄, v_身份证号, v_费别, v_付款方式)
      Into v_附加项目id
      From Dual;
    Exception
      When Others Then
        v_附加项目id := Null;
    End;
    If v_附加项目id Is Not Null Then
      If Instr(v_附加项目id, '|') > 0 Then
        v_附加内容   := v_附加项目id || ','; --以空格分开以|结尾,没有结算号码的
        v_附加项目id := '';
        n_Cursor     := 0;
        While v_附加内容 Is Not Null Loop
          v_附加值         := Substr(v_附加内容, 1, Instr(v_附加内容, ',') - 1);
          Price_Rec.项目id := To_Number(Substr(v_附加值, 1, Instr(v_附加值, '|') - 1));
          v_附加值         := Substr(v_附加值, Instr(v_附加值, '|') + 1);
          Price_Rec.数次   := To_Number(Substr(v_附加值, 1, Instr(v_附加值, '|') - 1));
          v_附加值         := Substr(v_附加值, Instr(v_附加值, '|') + 1);
          Price_Rec.单价   := To_Number(Substr(v_附加值, 1, Instr(v_附加值, '|') - 1));
          v_附加值         := Substr(v_附加值, Instr(v_附加值, '|') + 1);
          Price_Rec.应收   := To_Number(Substr(v_附加值, 1, Instr(v_附加值, '|') - 1));
          v_附加值         := Substr(v_附加值, Instr(v_附加值, '|') + 1);
          Price_Rec.实收   := To_Number(v_附加值);
        
          n_Cursor := n_Cursor + 1;
          Price_Rec_Array(n_Cursor) := Price_Rec;
          v_附加内容 := Substr(v_附加内容, Instr(v_附加内容, ',') + 1);
          v_附加项目id := v_附加项目id || ',' || Price_Rec_Array(n_Cursor).项目id;
        End Loop;
      
        If v_附加项目id Is Not Null Then
          v_附加项目id := Substr(v_附加项目id, 2);
        End If;
      
        For c_附加项目 In (Select /*+cardinality(D,10)*/
                        5 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, c.Id As 收入项目id,
                        c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, -1 As 执行科室类型
                       From 收费项目目录 A, 收费价目 B, 收入项目 C, Table(f_Str2List(v_附加项目id)) D
                       Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = d.Column_Value And Sysdate Between b.执行日期 And
                             Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And
                             (b.价格等级 = v_普通等级 Or
                             (b.价格等级 Is Null And Not Exists
                              (Select 1
                                From 收费价目
                                Where b.收费细目id = 收费细目id And 价格等级 = v_普通等级 And Sysdate Between 执行日期 And
                                      Nvl(终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')))))) Loop
        
          n_序号 := n_序号 + 1;
          For n_Cursor In 1 .. Price_Rec_Array.Count Loop
            If c_附加项目.项目id = Price_Rec_Array(n_Cursor).项目id Then
              Zl_病人预约挂号记录_Update(v_No, n_序号, Null, Null, c_附加项目.类别, c_附加项目.项目id, Price_Rec_Array(n_Cursor).数次,
                                 Price_Rec_Array(n_Cursor).单价, c_附加项目.收入项目id, c_附加项目.收据费目, Price_Rec_Array(n_Cursor).应收,
                                 Price_Rec_Array(n_Cursor).实收, Null, Null, Null, Null, Null, n_病人科室id, n_执行部门id);
            
              n_实收金额 := Price_Rec_Array(n_Cursor).实收;
              n_结帐金额 := n_结帐金额 + n_实收金额;
              Exit;
            End If;
          End Loop;
        End Loop;
      Else
        For c_附加项目 In (Select /*+cardinality(D,10)*/
                        5 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, 1 As 数次, c.Id As 收入项目id,
                        c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, -1 As 执行科室类型
                       From 收费项目目录 A, 收费价目 B, 收入项目 C, Table(f_Str2List(v_附加项目id)) D
                       Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = d.Column_Value And Sysdate Between b.执行日期 And
                             Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And
                             (b.价格等级 = v_普通等级 Or
                             (b.价格等级 Is Null And Not Exists
                              (Select 1
                                From 收费价目
                                Where b.收费细目id = 收费细目id And 价格等级 = v_普通等级 And Sysdate Between 执行日期 And
                                      Nvl(终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')))))
                       Union All
                       Select /*+cardinality(E,10)*/
                        6 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, d.从项数次 As 数次,
                        c.Id As 收入项目id, c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, -1 As 执行科室类型
                       From 收费项目目录 A, 收费价目 B, 收入项目 C, 收费从属项目 D, Table(f_Str2List(v_附加项目id)) E
                       Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = d.从项id And d.主项id = e.Column_Value And
                             Sysdate Between b.执行日期 And Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And
                             (b.价格等级 = v_普通等级 Or
                             (b.价格等级 Is Null And Not Exists
                              (Select 1
                                From 收费价目
                                Where b.收费细目id = 收费细目id And 价格等级 = v_普通等级 And Sysdate Between 执行日期 And
                                      Nvl(终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')))))) Loop
          n_序号 := n_序号 + 1;
          If c_附加项目.性质 = 5 Then
            n_从属父号 := n_序号;
          End If;
        
          v_实收     := Zl_Actualmoney(v_费别, c_附加项目.项目id, c_附加项目.收入项目id, c_附加项目.数次 * c_附加项目.单价);
          n_实收金额 := To_Number(Substr(v_实收, Instr(v_实收, ':') + 1));
        
          If c_附加项目.性质 = 5 Then
            Zl_病人预约挂号记录_Update(v_No, n_序号, Null, Null, c_附加项目.类别, c_附加项目.项目id, c_附加项目.数次, c_附加项目.单价, c_附加项目.收入项目id,
                               c_附加项目.收据费目, c_附加项目.数次 * c_附加项目.单价, n_实收金额, Null, Null, Null, Null, Null, n_病人科室id,
                               n_执行部门id);
          Else
            Zl_病人预约挂号记录_Update(v_No, n_序号, Null, n_从属父号, c_附加项目.类别, c_附加项目.项目id, c_附加项目.数次, c_附加项目.单价, c_附加项目.收入项目id,
                               c_附加项目.收据费目, c_附加项目.数次 * c_附加项目.单价, n_实收金额, Null, Null, Null, Null, Null, n_病人科室id,
                               n_执行部门id);
          End If;
          n_结帐金额 := n_结帐金额 + n_实收金额;
        End Loop;
      End If;
    End If;
  
    --检查总金额是否正确
    If Nvl(n_误差额, 0) = 0 Then
      n_误差额 := Nvl(n_结帐金额, 0) - Nvl(n_结算金额, 0);
      If Abs(n_误差额) > 1.00 Then
        v_Err_Msg := '单据缴款金额与实际结算的误差太大!';
        Raise Err_Item;
      End If;
    End If;
    If Nvl(n_结帐金额, 0) <> Nvl(n_结算金额, 0) + Nvl(n_误差额, 0) Then
      Select Max(操作员姓名) Into v_操作员 From 门诊费用记录 Where 记录性质 = 4 And NO = v_Nos;
      If v_操作员 = v_操作员姓名 Then
        v_Err_Msg := '指定的缴费单据的总金额不对,请重新选择缴费单据!';
        Raise Err_Special;
      Else
        v_Err_Msg := '指定的缴费单据的总金额不对,请重新选择缴费单据!';
        Raise Err_Item;
      End If;
    End If;
  
    n_关联交易id := 0;
    For c_结算方式 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                          Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号,
                          Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                          Extractvalue(b.Column_Value, '/JS/JSJE') As 结算金额,
                          Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                          Extractvalue(b.Column_Value, '/JS/JYSM') As 交易说明, Extractvalue(b.Column_Value, '/JS/ZY') As 摘要,
                          Extractvalue(b.Column_Value, '/JS/SFCYJ') As 是否冲预交,
                          Extractvalue(b.Column_Value, '/JS/SFXFK') As 是否消费卡
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
    
      n_消费卡       := Nvl(c_结算方式.是否消费卡, 0);
      v_临时结算方式 := c_结算方式.结算方式;
    
      If Nvl(c_结算方式.是否冲预交, 0) = 1 Then
        n_预交支付 := c_结算方式.结算金额;
      Else
        n_普通支付 := Nvl(n_普通支付, 0) + c_结算方式.结算金额;
      
        If c_结算方式.结算卡类别 Is Not Null Then
          --三方卡结算方式
          If n_消费卡 = 0 Then
            n_卡类别id := Get_Cardtypeid(c_结算方式.结算卡类别, 0, v_临时结算方式);
          Else
            n_结算卡序号 := Get_Cardtypeid(c_结算方式.结算卡类别, 1, v_临时结算方式);
          End If;
          v_结算卡号   := c_结算方式.结算卡号;
          v_交易流水号 := c_结算方式.交易流水号;
          v_交易说明   := c_结算方式.交易说明;
          v_摘要       := c_结算方式.摘要;
        
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          If Nvl(n_关联交易id, 0) = 0 Then
            n_关联交易id := n_预交id;
          End If;
          v_结算方式 := v_结算方式 || '|' || v_临时结算方式 || ',' || c_结算方式.结算金额 || ',,1' || ',' || n_预交id || ',' || n_关联交易id;
        Else
          Select Nvl(Max(1), 0) Into n_Exists From 结算方式 Where 名称 = Nvl(v_临时结算方式, '-') And 性质 In (3, 4);
        
          If n_Exists = 1 Then
            n_医保支付 := c_结算方式.结算金额;
          Else
            --其他结算方式
            v_结算方式 := v_结算方式 || '|' || v_临时结算方式 || ',' || c_结算方式.结算金额 || ',,0';
          End If;
        End If;
      End If;
    End Loop;
  
    If v_结算方式 Is Not Null Then
      v_结算方式 := Substr(v_结算方式, 2);
    End If;
  
    If n_挂号模式 = 0 Then
      Zl_预约挂号接收_Insert(v_Nos, Null, Null, n_结帐id, Zl_Get_出诊诊室(v_号码), n_病人id, n_门诊号, v_姓名, v_性别, v_年龄, v_医疗付款方式编码, v_费别,
                       v_结算方式, n_普通支付, n_预交支付, n_医保支付, d_发生时间, v_操作员编码, v_操作员姓名, d_收费时间, n_卡类别id, n_结算卡序号, v_结算卡号,
                       v_交易流水号, v_交易说明, Null, 0, 0, Null, 1);
    Else
      Zl_预约挂号接收_出诊_Insert(v_Nos, Null, Null, n_结帐id, Zl_Get_出诊诊室(v_号码, n_出诊记录id), n_病人id, n_门诊号, v_姓名, v_性别, v_年龄,
                          v_医疗付款方式编码, v_费别, v_结算方式, n_普通支付, n_预交支付, Null, d_发生时间, v_操作员编码, v_操作员姓名, d_收费时间, n_卡类别id,
                          n_结算卡序号, v_结算卡号, v_交易流水号, v_交易说明, Null, 0, 0, Null, 1);
    End If;
  
  End If;

  If Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) = 0 Then
    v_Temp := '<CZSJ>' || To_Char(d_收费时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<JZID>' || n_结帐id || '</JZID>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    Xml_Out := x_Templet;
    Return;
  End If;

  Zl_预约挂号接收_序号更新(v_Nos, Null, v_操作员编码, v_操作员姓名, d_发生时间, d_收费时间);
  n_连续更新 := 0;
  For c_结算方式 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                        Extractvalue(b.Column_Value, '/JS/JSJE') As 结算金额,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                        Extractvalue(b.Column_Value, '/JS/JYSM') As 交易说明, Extractvalue(b.Column_Value, '/JS/ZY') As 摘要,
                        Extractvalue(b.Column_Value, '/JS/SFCYJ') As 是否冲预交,
                        Extractvalue(b.Column_Value, '/JS/SFXFK') As 是否消费卡
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
  
    v_结算方式 := c_结算方式.结算方式;
    If c_结算方式.结算卡类别 Is Null Then
      v_卡类别 := v_结算方式;
    Else
      v_卡类别 := Get_Cardname(c_结算方式.结算卡类别, c_结算方式.是否消费卡);
      If Nvl(c_结算方式.是否消费卡, 0) = 0 Then
        n_卡类别id := Get_Cardtypeid(c_结算方式.结算卡类别, 0, v_结算方式);
      
        If Nvl(n_卡类别id, 0) <> 0 Then
          v_结算卡号 := c_结算方式.结算卡号;
        
          Select Max(关联交易id)
          Into n_关联交易id
          From 病人预交记录
          Where 结帐id = n_结帐id And 卡类别id = n_卡类别id And Rownum < 2;
          If Nvl(n_关联交易id, 0) = 0 Then
            Select 病人预交记录_Id.Nextval Into n_关联交易id From Dual;
          End If;
        
          Zl_病人挂号收费_Modify(v_Nos, n_结帐id, v_结算方式 || ',' || c_结算方式.结算金额 || ',,' || c_结算方式.摘要, 1, 0, 0, 1, Null, n_连续更新,
                           n_关联交易id, n_卡类别id, v_结算卡号, c_结算方式.交易流水号, c_结算方式.交易说明);
        
          For c_扩展信息 In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As 交易名称,
                                Extractvalue(b.Column_Value, '/EXPEND/JYLR') As 交易内容
                         From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS/EXPENDLIST/EXPEND'))) B) Loop
            Zl_三方结算交易_Insert(n_卡类别id, 0, v_结算卡号, n_结帐id, c_扩展信息.交易名称 || '|' || c_扩展信息.交易内容, 0);
          End Loop;
        End If;
      End If;
    End If;
    n_连续更新 := 1;
  
    Update 三方交易记录
    Set 业务结算id = n_结帐id
    Where 流水号 = c_结算方式.交易流水号 And 类别 = v_卡类别 And 业务类型 = n_业务类型;
  End Loop;

  --6.电子票据处理
  n_是否电子票据 := b_Einvoice_Request.Einvoice_Start(4, Null);
  Zl_病人挂号收费_Modify(v_Nos, n_结帐id, Null, 0, 1, 0, 1, Null, 0, Null, Null, Null, Null, Null, 0, 2, n_是否电子票据);

  --处理汇总
  Zl_病人挂号汇总_Update(v_医生姓名, n_医生id, n_项目id, n_科室id, d_发生时间, 2, v_号码, 1, n_出诊记录id);
  If Nvl(n_是否电子票据, 0) = 1 Then
    If b_Einvoice_Request.Einvoice_Create(4, n_结帐id, Null, v_Err_Msg) = 0 Then
      --电子票据开具成功
      Raise Err_Item;
    End If;
  
    Select Max(1), Max(姓名), Max(号码), Max(To_Char(生成时间, 'yyyy-mm-dd')), Max(Url内网), Max(Url外网), Max(票据金额)
    Into n_开票标志, v_患者姓名, v_发票编号, v_开票日期, v_Url, v_Url外网, n_发票金额
    From 电子票据使用记录
    Where 结算id = n_结帐id And 票种 = 4 And 记录状态 = 1;
  
    If v_患者姓名 Is Not Null Then
      v_姓名 := v_患者姓名;
    End If;
  
  End If;
  v_Temp := '<CZSJ>' || To_Char(d_收费时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<JZID>' || n_结帐id || '</JZID>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;

  v_Temp := '<KPBZ>' || Nvl(n_开票标志, 0) || '</KPBZ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;

  v_Temp := '<URL>' || Nvl(v_Url, '') || '</URL>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;

  v_Temp := '<NETURL>' || Nvl(v_Url外网, '') || '</NETURL>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;

  v_Temp := '<FPTT>' || v_姓名 || '</FPTT>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;

  v_Temp := '<FPH>' || v_发票编号 || '</FPH>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<FPJE>' || Nvl(n_发票金额, 0) || '</FPJE>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<KPRQ>' || v_开票日期 || '</KPRQ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;

  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Err_Special Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20105, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Payment;
/


Create Or Replace Procedure Zl_Third_Charge_Delcheck
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  -------------------------------------------------------------------------------------------------- 
  --功能:三方退费检查 
  --入参:Xml_In: 
  --<IN> 
  --    <BRID>病人ID</BRID> 
  --    <XM>姓名</XM> 
  --    <SFZH>身份证号</SFZH> 
  --    <JE></JE> //退款总金额 
  --    <JSKLB></JSKLB>     //结算卡类别 
  --    <TFZY>退费摘要</TFZY> 
  --    <JCFP>1</JCFP>      //检查发票,0-不检查;1-检查;为1时，打印了发票的单据不能退费 
  --    <XL>险类</XL>         //医保病人险类,空代表普通病人 
  --    <FYLIST> 
  --        <FY> 
  --           <DJH>退款单据号</DJH> 
  --           <XH>退款序号(格式:1,2,3..为空代表退剩余数量)</DJH> 
  --        <FY> 
  --    </FYLIST> 
  --    <TKLIST> 
  --        <TK> 
  --            <TKKLB>退款卡类别</TKKLB> 
  --            <TKKH>退款卡号</TKKH> 
  --            <TKFS>退款方式</TKFS> //退款方式:现金;支票,如果是三方卡,可以传空 
  --            <TKJE>支付金额</TKJE> 
  --            <JYLSH>交易流水号</JYLSH> 
  --            <JYSM>交易说明</JYSM> 
  --            <TKZY>摘要</TKZY> 
  --            <TYJK>退回预交款</TYJK> //允冲预交时,只填JSJE节点:1-冲预交 
  --            <SFXFK>是否消费卡</SFXFK>   //(1-是消费卡),消费卡时,传入结算卡类别,结算卡号,结算金额等接点 
  --            <DJH>S0000001</DJH> //分单据结算时传入 
  --            <EXPENDLIST>  //扩展交易信息 
  --                <EXPEND> 
  --                    <JYMC>交易名称</JYMC> 
  --                    <JYLR>交易内容</JYLR> 
  --                </EXPEND> 
  --            </EXPENDLIST> 
  --        </TK> 
  --    </TKLIST> 
  --</IN> 

  --出参:Xml_Out 
  --  <OUT> 
  --    ――如无下列错误结点则说明通过检查 
  --    <ERROR> 
  --      <MSG>错误信息</MSG> 
  --    </ERROR> 
  --  </OUT> 
  -------------------------------------------------------------------------------------------------- 
  n_退款总额 门诊费用记录.实收金额%Type;

  n_病人id     门诊费用记录.病人id%Type;
  v_姓名       病人信息.姓名%Type;
  v_身份证号   病人信息.身份证号%Type;
  n_单据病人id 门诊费用记录.病人id%Type;
  n_结帐id     门诊费用记录.结帐id%Type;
  n_原结算序号 病人预交记录.结算序号%Type;
  v_结算卡类别 Varchar2(100);
  v_结算方式   医疗卡类别.结算方式%Type;
  n_卡类别id   医疗卡类别.Id%Type;
  v_卡名称     医疗卡类别.名称%Type;

  v_摘要     门诊费用记录.摘要%Type;
  n_Count    Number(18);
  n_Temp     Number(18);
  n_检查发票 Number(3);
  n_是否打印 Number(3);
  n_退费模式 Number(3);
  n_状态     Number(3);
  n_险类     病人信息.险类%Type;

  v_Temp    Varchar2(32767); --临时XML 
  x_Templet Xmltype; --模板XML 

  v_Err_Msg Varchar2(200);
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  --0.获取入参中的病人ID等信息 
  Select To_Number(Extractvalue(Value(A), 'IN/BRID')), To_Number(Extractvalue(Value(A), 'IN/JE')),
         Extractvalue(Value(A), 'IN/TFZY'), To_Number(Extractvalue(Value(A), 'IN/JCFP')),
         Extractvalue(Value(A), 'IN/JSKLB'), To_Number(Extractvalue(Value(A), 'IN/XL')),
         Extractvalue(Value(A), 'IN/SFZH'), Extractvalue(Value(A), 'IN/XM')
  Into n_病人id, n_退款总额, v_摘要, n_检查发票, v_结算卡类别, n_险类, v_身份证号, v_姓名
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  If Nvl(n_病人id, 0) = 0 And Not v_身份证号 Is Null And Not v_姓名 Is Null Then
    n_病人id := Zl_Third_Getpatiid(v_身份证号, v_姓名);
  End If;
  --0.相关检查 
  If Nvl(n_病人id, 0) = 0 Then
    v_Err_Msg := '不能有效识别病人身份,不允许退费操作!';
    Raise Err_Item;
  End If;

  n_退费模式 := zl_GetSysParameter('门诊退费须先申请');

  If v_结算卡类别 Is Not Null Then
    Begin
      n_卡类别id := To_Number(v_结算卡类别);
    Exception
      When Others Then
        n_卡类别id := 0;
    End;
    If n_卡类别id = 0 Then
      Begin
        Select ID, 名称 Into n_卡类别id, v_卡名称 From 医疗卡类别 Where 名称 = v_结算卡类别;
      Exception
        When Others Then
          v_Err_Msg := '无法确认传入的结算卡！';
          Raise Err_Item;
      End;
    Else
      Begin
        Select 名称 Into v_卡名称 From 医疗卡类别 Where ID = n_卡类别id;
      Exception
        When Others Then
          v_Err_Msg := '无法确认传入的结算卡！';
          Raise Err_Item;
      End;
    End If;
  Else
    n_卡类别id := 0;
  End If;

  If Nvl(n_卡类别id, 0) <> 0 Then
    Select 结算方式 Into v_结算方式 From 医疗卡类别 Where ID = n_卡类别id;
  End If;

  --人员id,人员编号,人员姓名 
  v_Temp := Zl_Identity(1);
  If Nvl(v_Temp, '0') = '0' Or Nvl(v_Temp, '_') = '_' Then
    v_Err_Msg := '系统不能认别有效的操作员,不允许退费!';
    Raise Err_Item;
  End If;

  --1.退费检查 
  n_Count      := 0;
  n_原结算序号 := 0;
  For c_费用 In (Select Extractvalue(b.Column_Value, '/FY/DJH') As 单据号, Extractvalue(b.Column_Value, '/FY/XH') As 退款序号
               From Table(Xmlsequence(Extract(Xml_In, '/IN/FYLIST/FY'))) B) Loop
  
    If c_费用.单据号 Is Null Then
      v_Err_Msg := '未确定指定退费的单据号,不能退费!';
      Raise Err_Item;
    End If;
  
    If Nvl(n_退费模式, 0) = 1 Then
      Begin
        Select Nvl(状态, 0) Into n_状态 From 病人退费申请 Where NO = c_费用.单据号 And Mod(记录性质, 10) = 1;
      Exception
        When Others Then
          n_状态 := 0;
      End;
      If n_状态 <> 1 Then
        v_Err_Msg := '当前为退费申请模式,退费之前需申请并审核通过该单据!';
        Raise Err_Item;
      End If;
    End If;
  
    Begin
      Select a.结算序号, a.结帐id, a.病人id
      Into n_Temp, n_结帐id, n_单据病人id
      From 病人预交记录 A, 门诊费用记录 B
      Where a.结帐id = b.结帐id And b.No = c_费用.单据号 And b.记录性质 = 1 And Nvl(b.费用状态, 0) = 0 And b.记录状态 In (1, 3) And
            Rownum < 2;
    Exception
      When Others Then
        n_Temp := Null;
    End;
  
    If n_Temp Is Null Then
      v_Err_Msg := '指定的单据号:' || c_费用.单据号 || '未找到,不能退费!';
      Raise Err_Item;
    End If;
  
    If Nvl(n_单据病人id, 0) = 0 Then
      Begin
        Select 病人id
        Into n_单据病人id
        From 门诊费用记录
        Where NO = c_费用.单据号 And 记录性质 = 1 And Nvl(费用状态, 0) = 0 And 记录状态 In (1, 3) And Rownum < 2;
      Exception
        When Others Then
          n_单据病人id := 0;
      End;
    End If;
  
    If Nvl(n_病人id, 0) <> Nvl(n_单据病人id, 0) Then
      v_Err_Msg := '本次退费的收费单:' || c_费用.单据号 || '不是当前病人的收费单,不能退费!';
      Raise Err_Item;
    End If;
  
    If n_原结算序号 <> 0 And n_原结算序号 <> n_Temp Then
      v_Err_Msg := '本次退费的单据号不是一次收费结算,不能退费!';
      Raise Err_Item;
    End If;
    n_原结算序号 := n_Temp;
  
    Select Count(1) Into n_Temp From 费用补充记录 Where 收费结帐id = n_结帐id;
    If Nvl(n_Temp, 0) <> 0 Then
      v_Err_Msg := '本次退费的单据号已经进行了保险补充结算,不能退费!';
      Raise Err_Item;
    End If;
  
    If Nvl(n_卡类别id, 0) <> 0 Then
      If Nvl(n_险类, 0) = 0 Then
        Select Count(1) Into n_Temp From 病人预交记录 Where 结帐id = n_结帐id And 卡类别id <> n_卡类别id;
      Else
        Select Count(1)
        Into n_Temp
        From 病人预交记录 A, 结算方式 B
        Where a.结帐id = n_结帐id And 卡类别id <> n_卡类别id And a.结算方式 = b.名称 And b.性质 Not In (3, 4);
        If n_Temp = 0 Then
          Select Nvl(Max(1), 0)
          Into n_Temp
          From 保险结算记录 A
          Where a.记录id = n_结帐id And 险类 <> n_险类 And Rownum < 2;
        End If;
      End If;
      If Nvl(n_Temp, 0) > 0 Then
        v_Err_Msg := '本次退费的单据包含' || v_结算方式 || '以外的结算方式,不能退费!';
        Raise Err_Item;
      End If;
    End If;
  
    If Nvl(n_检查发票, 0) = 1 Then
      Select Max(Decode(a.实际票号, Null, 0, 1))
      Into n_是否打印
      From 门诊费用记录 A
      Where NO = c_费用.单据号 And 记录性质 = 1;
      If Nvl(n_是否打印, 0) = 1 Then
        v_Err_Msg := '本次退费的单据号已开发票,不能退费!';
        Raise Err_Item;
      End If;
    End If;
  
    --电子票据检查 
    If b_Einvoice_Request.Einvoice_Cancel_Check(1, n_结帐id, v_Err_Msg) = 0 Then
      --失败后，直接抛错
      Raise Err_Item;
    End If;
  
    n_Count := n_Count + 1;
  End Loop;

  If n_Count = 0 Then
    v_Err_Msg := '未确定本次需要退费的单据,不能退费!';
    Raise Err_Item;
  End If;

  --2.支付方式检查 
  n_Count := 0;
  For c_结算方式 In (Select Extractvalue(b.Column_Value, '/TK/TKKLB') As 卡类别, Extractvalue(b.Column_Value, '/TK/TKKH') As 卡号,
                        Extractvalue(b.Column_Value, '/TK/TKFS') As 结算方式,
                        -1 * To_Number(Extractvalue(b.Column_Value, '/TK/TKJE')) As 退款金额,
                        Extractvalue(b.Column_Value, '/TK/JYLSH') As 交易流水号,
                        Extractvalue(b.Column_Value, '/TK/JYSM') As 交易说明, Extractvalue(b.Column_Value, '/TK/TKZY') As 摘要,
                        Extractvalue(b.Column_Value, '/TK/TYJK') As 是否退预交,
                        Extractvalue(b.Column_Value, '/TK/SFXFK') As 是否消费卡,
                        Extract(b.Column_Value, '/TK/EXPENDLIST') As Expend
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/TKLIST/TK'))) B) Loop
  
    --1.退回三方卡 
    If c_结算方式.卡类别 Is Not Null And Nvl(c_结算方式.是否消费卡, 0) = 0 Then
      --1.三方卡结算 
      Null;
    Elsif c_结算方式.卡类别 Is Not Null And Nvl(c_结算方式.是否消费卡, 0) = 1 Then
      --2.消费卡结算 
      Null;
    Elsif Nvl(c_结算方式.是否退预交, 0) = 1 Then
      --3.退预交款 
      Null;
    Else
      --4.普通结算 
      If c_结算方式.结算方式 Is Null Then
        v_Err_Msg := '未指定支付方式,不能退费!';
        Raise Err_Item;
      End If;
    End If;
    n_Count := n_Count + 1;
  End Loop;

  If n_Count = 0 Then
    v_Err_Msg := '不能有效确认当前的支付方式,不能退费!';
    Raise Err_Item;
  End If;

  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Charge_Delcheck;
/

Create Or Replace Procedure Zl_Third_Charge_Del
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is

  -------------------------------------------------------------------------------------------------- 
  --功能:三方退费交易 
  --入参:Xml_In: 
  --<IN> 
  --    <BRID>病人ID</BRID> 
  --    <XM>姓名</XM> 
  --    <SFZH>身份证号</SFZH> 
  --    <JE></JE> //退款总金额 
  --    <JSKLB></JSKLB>     //结算卡类别 
  --    <TFZY>退费摘要</TFZY> 
  --    <JCFP>1</JCFP>      //检查发票 
  --    <JSMS>1</JSMS>          //结算模式：0-普通模式，1-异步结算模式 
  --    <CZLX>0</CZLX>          //操作类型：结算模式为1时传入，0-开始结算，1-完成结算，2-回退结算 
  --    <CXID>1</CXID>          //冲销结帐ID，操作类型为1或2时传入 
  --    <FYLIST> 
  --        <FY> 
  --           <DJH>退款单据号</DJH> 
  --           <XH>退款序号(格式:1,2,3..为空代表退剩余数量)</DJH> 
  --        <FY> 
  --    </FYLIST> 
  --    <TKLIST>          //结算列表，操作类型为2时可不传入 
  --        <TK> 
  --            <TKKLB>退款卡类别</TKKLB> 
  --            <TKKH>退款卡号</TKKH> 
  --            <TKFS>退款方式</TKFS> //退款方式:现金;支票,如果是三方卡,可以传空 
  --            <TKJE>支付金额</TKJE> 
  --            <JYLSH>交易流水号</JYLSH> 
  --            <TKZY>摘要</TKZY> 
  --            <TYJK>退回预交款</TYJK> //允冲预交时,只填JSJE节点:1-冲预交 
  --            <SFXFK>是否消费卡</SFXFK>   //(1-是消费卡),消费卡时,传入结算卡类别,结算卡号,结算金额等接点 
  --            <DJH>S0000001</DJH> //分单据结算时传入 
  --            <EXPENDLIST>  //扩展交易信息 
  --                <EXPEND> 
  --                    <JYMC>交易名称</JYMC> 
  --                    <JYLR>交易内容</JYLR> 
  --                </EXPEND> 
  --            </EXPENDLIST> 
  --        </TK> 
  --    </TKLIST> 
  --</IN> 

  --出参:Xml_Out 
  --  <OUTPUT> 
  --    <CZSJ>操作时间</CZSJ>          //HIS的登记时间 
  --    <YJZID>原结帐ID</YJZID>       //原结帐ID 
  --    <CXID>冲销ID</CXID>          //冲销ID 
  --  <KPBZ>开票标志</KPBZ> //1-成功开具电子票据;0-未开票成功标志
  --  <URL>H5页面URL</URL>
  --  <NETURL>外网H5页面URL</NETURL>
  --  <FPTT>发票抬头</FPTT>        //病人姓名
  --  <FPH>发票号</FPH>             //发票编号
  --  <FPJE>发票金额</FPJE>        //100.00
  --  <KPRQ>开票日期</KPRQ>   //yyyy-mm-dd
  --    ――如无下列错误结点则说明正确执行 
  --    <ERROR> 
  --      <MSG>错误信息</MSG> 
  --    </ERROR> 
  --  </OUTPUT> 
  -------------------------------------------------------------------------------------------------- 
  n_退款总额 门诊费用记录.实收金额%Type;
  n_卡类别id 医疗卡类别.Id%Type;
  v_结算方式 Varchar2(2000);
  n_结算模式 Number(1); --0-普通模式，1-异步结算模式 
  n_操作类型 Number(1); --结算模式为1时传入，0 - 开始结算，1 - 完成结算，2 - 回退结算 

  n_病人id     门诊费用记录.病人id%Type;
  v_姓名       病人信息.姓名%Type;
  v_身份证号   病人信息.身份证号%Type;
  n_单据病人id 门诊费用记录.病人id%Type;
  v_操作员编码 门诊费用记录.操作员编号%Type;
  v_操作员姓名 门诊费用记录.操作员姓名%Type;
  n_冲销id     门诊费用记录.结帐id%Type;
  n_结帐id     门诊费用记录.结帐id%Type;
  n_结帐金额   门诊费用记录.结帐金额%Type;
  n_误差额     病人预交记录.冲预交%Type;
  n_原结算序号 病人预交记录.结算序号%Type;
  l_挂号单     t_StrList := t_StrList();
  n_结算序号   病人预交记录.结算序号%Type;
  n_检查发票   Number(3);
  n_是否打印   Number(3);
  v_结算卡类别 Varchar2(100);
  v_结帐ids    Varchar2(1000);

  n_消费卡     Number;
  n_消费卡id   消费卡信息.Id%Type;
  v_摘要       门诊费用记录.摘要%Type;
  n_Count      Number(18);
  n_Billcount  Number(18);
  n_关联交易id 病人预交记录.关联交易id%Type;
  n_删除原结算 Number;

  d_退费时间 病人预交记录.收款时间%Type;
  v_挂号单   病人挂号记录.No%Type;
  v_收费单   门诊费用记录.No%Type;

  v_退费结算 Varchar2(2000);
  v_普通结算 Varchar2(4000);
  n_剩余金额 门诊费用记录.结帐金额%Type;

  n_是否电子票据 病人预交记录.是否电子票据%Type;
  n_开票标志     Number(2);
  v_患者姓名     电子票据使用记录.姓名%Type;
  v_发票编号     电子票据使用记录.号码%Type;
  v_开票日期     Varchar2(20);
  n_发票金额     电子票据使用记录.票据金额%Type;
  v_Url          电子票据使用记录.Url内网%Type;
  v_Url外网      电子票据使用记录.Url外网%Type;

  v_Temp    Varchar2(32767); --临时XML 
  v_Err_Msg Varchar2(200);
  Err_Item Exception;

  --获取卡类别名称 
  Function Get_Cardname
  (
    卡类别_In Varchar2,
    消费卡_In Number
  ) Return Varchar2 As
    v_名称       医疗卡类别.名称%Type;
    n_By_Id_Find Number;
  
    v_Err_Msg Varchar2(200);
    Err_Item Exception;
  Begin
    If 卡类别_In Is Null Then
      Return Null;
    End If;
  
    Select Decode(Translate(Nvl(卡类别_In, 'abcd'), '#1234567890', '#'), Null, 1, 0) Into n_By_Id_Find From Dual;
  
    --1.按卡类别ID查找医疗卡 
    If Nvl(消费卡_In, 0) = 0 And Nvl(n_By_Id_Find, 0) = 1 Then
      Begin
        Select 名称 Into v_名称 From 医疗卡类别 Where ID = To_Number(卡类别_In);
      Exception
        When Others Then
          v_Err_Msg := '未找到指定的结算支付信息！';
          Raise Err_Item;
      End;
    End If;
  
    --2.按卡名称查找医疗卡 
    If Nvl(消费卡_In, 0) = 0 And Nvl(n_By_Id_Find, 0) = 0 Then
      Begin
        Select 名称 Into v_名称 From 医疗卡类别 Where 名称 = 卡类别_In;
      Exception
        When Others Then
          v_Err_Msg := 卡类别_In || '不存在!';
          Raise Err_Item;
      End;
    End If;
  
    --3.按卡类别ID查找消费卡 
    If Nvl(消费卡_In, 0) = 1 And Nvl(n_By_Id_Find, 0) = 1 Then
      Begin
        Select 名称 Into v_名称 From 消费卡类别目录 Where 编号 = To_Number(卡类别_In);
      Exception
        When Others Then
          v_Err_Msg := '未找到指定的结算支付信息！';
          Raise Err_Item;
      End;
    End If;
  
    --3.按卡名称查找消费卡 
    If Nvl(消费卡_In, 0) = 1 And Nvl(n_By_Id_Find, 0) = 0 Then
      Begin
        Select 名称 Into v_名称 From 消费卡类别目录 Where 名称 = 卡类别_In;
      Exception
        When Others Then
          v_Err_Msg := 卡类别_In || '不存在!';
          Raise Err_Item;
      End;
    End If;
  
    Return v_名称;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  End;

  --获取卡类别ID 
  Function Get_Cardtypeid
  (
    卡类别_In    Varchar2,
    消费卡_In    Number,
    结算方式_Out In Out 医疗卡类别.结算方式%Type
  ) Return Number As
    n_卡类别id 医疗卡类别.Id%Type;
    v_名称     医疗卡类别.名称%Type;
    n_启用     医疗卡类别.是否启用%Type;
    v_结算方式 医疗卡类别.结算方式%Type;
  
    n_By_Id_Find Number;
    v_Err_Msg    Varchar2(200);
    Err_Item Exception;
  Begin
    If 卡类别_In Is Null Then
      Return 0;
    End If;
  
    Select Decode(Translate(Nvl(卡类别_In, 'abcd'), '#1234567890', '#'), Null, 1, 0) Into n_By_Id_Find From Dual;
  
    --1.按卡类别ID查找医疗卡 
    If Nvl(消费卡_In, 0) = 0 And Nvl(n_By_Id_Find, 0) = 1 Then
      Begin
        Select ID, 结算方式, 名称, 是否启用
        Into n_卡类别id, v_名称, v_结算方式, n_启用
        From 医疗卡类别
        Where ID = To_Number(卡类别_In);
      Exception
        When Others Then
          v_Err_Msg := '未找到指定的结算支付信息！';
          Raise Err_Item;
      End;
    End If;
  
    --2.按卡名称查找医疗卡 
    If Nvl(消费卡_In, 0) = 0 And Nvl(n_By_Id_Find, 0) = 0 Then
      Begin
        Select ID, 结算方式, 名称, 是否启用
        Into n_卡类别id, v_名称, v_结算方式, n_启用
        From 医疗卡类别
        Where 名称 = 卡类别_In;
      Exception
        When Others Then
          v_Err_Msg := 卡类别_In || '不存在!';
          Raise Err_Item;
      End;
    End If;
  
    --3.按卡类别ID查找消费卡 
    If Nvl(消费卡_In, 0) = 1 And Nvl(n_By_Id_Find, 0) = 1 Then
      Begin
        Select 编号, 结算方式, 名称, 启用
        Into n_卡类别id, v_名称, v_结算方式, n_启用
        From 消费卡类别目录
        Where 编号 = To_Number(卡类别_In);
      Exception
        When Others Then
          v_Err_Msg := '未找到指定的结算支付信息！';
          Raise Err_Item;
      End;
    End If;
  
    --3.按卡名称查找消费卡 
    If Nvl(消费卡_In, 0) = 1 And Nvl(n_By_Id_Find, 0) = 0 Then
      Begin
        Select 编号, 结算方式, 名称, 启用
        Into n_卡类别id, v_名称, v_结算方式, n_启用
        From 消费卡类别目录
        Where 名称 = 卡类别_In;
      Exception
        When Others Then
          v_Err_Msg := 卡类别_In || '不存在!';
          Raise Err_Item;
      End;
    End If;
  
    If Nvl(n_启用, 0) = 0 Then
      v_Err_Msg := v_名称 || '未启用，不允许进行缴费！';
      Raise Err_Item;
    End If;
  
    If 结算方式_Out Is Null Then
      If v_结算方式 Is Null Then
        v_Err_Msg := v_名称 || '未设置结算方式，不允许进行缴费！';
        Raise Err_Item;
      End If;
    
      结算方式_Out := v_结算方式;
    End If;
  
    Return n_卡类别id;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  End;

  Procedure Thirdcard_Balance
  (
    病人id_In     病人预交记录.病人id%Type,
    冲销id_In     病人预交记录.结帐id%Type,
    结算方式_In   病人预交记录.结算方式%Type,
    卡类别_In     病人预交记录.卡类别id%Type,
    卡号_In       病人预交记录.卡号%Type,
    退款金额_In   病人预交记录.冲预交%Type,
    交易流水号_In 病人预交记录.交易流水号%Type,
    交易说明_In   病人预交记录.交易说明%Type,
    No_In         病人预交记录.No%Type,
    关联交易id_In 病人预交记录.关联交易id%Type,
    摘要_In       病人预交记录.摘要%Type,
    Xmlexpned_In  Xmltype,
    结算模式_In   Number := 0,
    操作类型_In   Number := 0,
    删除原结算_In Number := 0
  ) Is
    --入参： 
    --         结算模式_in   0-普通模式，1-异步结算模式 
    --        操作类型_in   结算模式为1时传入，0 - 开始结算，1 - 完成结算，2 - 回退结算 
    --        删除原结算_in 操作类型_In为1时有效，多个结算方式时调用多次该过程 
    v_退费结算 Varchar2(2000);
    n_校对标志 病人预交记录.校对标志%Type;
  Begin
    If Nvl(结算模式_In, 0) = 1 And Nvl(操作类型_In, 0) = 0 Then
      n_校对标志 := 1;
    Else
      n_校对标志 := 2;
    End If;
  
    --结算方式|结算金额|结算号码|结算摘要|单据号|是否普通结算 
    v_退费结算 := 结算方式_In || '|' || 退款金额_In || '| |' || 摘要_In || '|' || No_In || '|0';
    Zl_门诊退费结算_Modify(5, 病人id_In, 冲销id_In, v_退费结算, 0, 卡类别_In, 卡号_In, 交易流水号_In, 交易说明_In, 0, 0, 0, 0, Null, 0, Null, Null,
                     关联交易id_In, 删除原结算_In, n_校对标志);
  
    If Nvl(结算模式_In, 0) = 1 And Nvl(操作类型_In, 0) = 0 Then
      Return;
    End If;
  
    --保存扩展结算信息 
    For c_扩展 In (Select Extractvalue(j.Column_Value, '/EXPEND/JYMC') As Jymc,
                        Extractvalue(j.Column_Value, '/EXPEND/JYLR') As Jylr
                 From Table(Xmlsequence(Extract(Xmlexpned_In, '/EXPENDLIST/EXPEND'))) J) Loop
      Zl_三方结算交易_Insert(卡类别_In, 0, 卡号_In, 冲销id_In, c_扩展.Jymc || '|' || c_扩展.Jylr);
    End Loop;
  End Thirdcard_Balance;

  Procedure Squarecard_Balance
  (
    病人id_In     病人预交记录.病人id%Type,
    冲销id_In     病人预交记录.结帐id%Type,
    卡类别_In     病人预交记录.卡类别id%Type,
    卡号_In       病人预交记录.卡号%Type,
    退款金额_In   病人预交记录.冲预交%Type,
    交易流水号_In 病人预交记录.交易流水号%Type,
    交易说明_In   病人预交记录.交易说明%Type,
    Xmlexpned_In  Xmltype
  ) Is
    v_退费结算 Varchar2(2000);
    n_消费卡id 消费卡信息.Id%Type;
  Begin
    Select ID
    Into n_消费卡id
    From 消费卡信息
    Where 接口编号 = 卡类别_In And 卡号 = 卡号_In And
          序号 = (Select Max(序号) From 消费卡信息 Where 接口编号 = 卡类别_In And 卡号 = 卡号_In);
  
    --卡类别ID|卡号|消费卡ID|消费金额||. 
    v_退费结算 := 卡类别_In || '|' || 卡号_In || '|' || n_消费卡id || '|' || 退款金额_In;
    Zl_门诊退费结算_Modify(4, 病人id_In, 冲销id_In, v_退费结算, 0, Null, 卡号_In, 交易流水号_In, 交易说明_In, 0, 0, 0, 0);
  
    --保存扩展结算信息 
    For c_扩展 In (Select Extractvalue(j.Column_Value, '/EXPEND/JYMC') As Jymc,
                        Extractvalue(j.Column_Value, '/EXPEND/JYLR') As Jylr
                 From Table(Xmlsequence(Extract(Xmlexpned_In, '/EXPENDLIST/EXPEND'))) J) Loop
      Zl_三方结算交易_Insert(卡类别_In, 1, 卡号_In, 冲销id_In, c_扩展.Jymc || '|' || c_扩展.Jylr, 0);
    End Loop;
  End Squarecard_Balance;

Begin
  --0.获取入参中的病人ID等信息 
  Select To_Number(Extractvalue(Value(A), 'IN/BRID')), To_Number(Extractvalue(Value(A), 'IN/JE')),
         Extractvalue(Value(A), 'IN/TFZY'), To_Number(Extractvalue(Value(A), 'IN/JCFP')),
         Extractvalue(Value(A), 'IN/JSKLB'), Extractvalue(Value(A), 'IN/SFZH'), Extractvalue(Value(A), 'IN/XM'),
         Extractvalue(Value(A), 'IN/JSMS'), Extractvalue(Value(A), 'IN/CZLX'), Extractvalue(Value(A), 'IN/CXID')
  Into n_病人id, n_退款总额, v_摘要, n_检查发票, v_结算卡类别, v_身份证号, v_姓名, n_结算模式, n_操作类型, n_冲销id
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If Nvl(n_病人id, 0) = 0 And Not v_身份证号 Is Null And Not v_姓名 Is Null Then
    n_病人id := Zl_Third_Getpatiid(v_身份证号, v_姓名);
  End If;

  If Nvl(n_病人id, 0) = 0 Then
    v_Err_Msg := '无法确定病人信息,请检查!';
    Raise Err_Item;
  End If;

  If Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) <> 0 Then
    --n_结算模式 --0-普通模式，1-异步结算模式 
    --n_操作类型 :结算模式为1时传入，0 - 开始结算，1 - 完成结算，2 - 回退结算 
  
    If Nvl(n_冲销id, 0) = 0 Then
      v_Err_Msg := '没有指定相关的结算数据！';
      Raise Err_Item;
    End If;
  
    Begin
      Select 收款时间
      Into d_退费时间
      From 病人预交记录
      Where 结帐id = n_冲销id And Nvl(校对标志, 0) = 1 And Rownum < 2;
    Exception
      When Others Then
        v_Err_Msg := '没有找到指定的相关结算数据，可能已被处理！';
        Raise Err_Item;
    End;
  
    Select f_List2Str(Cast(Collect(To_Char(a.结帐id)) As t_StrList), ',', 1)
    Into v_结帐ids
    From 门诊费用记录 A, 门诊费用记录 B
    Where a.No = b.No And a.记录性质 = b.记录性质 And a.记录性质 = 1 And a.记录状态 In (1, 3) And b.结帐id = n_冲销id;
  End If;

  If Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) = 2 Then
    --删除结算数据，恢复划价单 
    --n_结算模式 --0-普通模式，1-异步结算模式 
    --n_操作类型 :结算模式为1时传入，0 - 开始结算，1 - 完成结算，2 - 回退结算 
  
    Zl_病人结算记录_Delete(n_冲销id);
    Zl_门诊退费结算_Cancel(n_冲销id);
  
    v_Temp  := '<CZSJ>' || To_Char(d_退费时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    v_Temp  := v_Temp || '<YJZID>' || v_结帐ids || '</YJZID>';
    v_Temp  := v_Temp || '<CXID>' || n_冲销id || '</CXID>';
    v_Temp  := v_Temp || '<KPBZ>' || 0 || '</KPBZ>';
    v_Temp  := v_Temp || '<URL>' || '' || '</URL>';
    v_Temp  := v_Temp || '<NETURL>' || '' || '</NETURL>';
    v_Temp  := v_Temp || '<FPTT>' || v_姓名 || '</FPTT>';
    v_Temp  := v_Temp || '<FPH>' || '' || '</FPH>';
    v_Temp  := v_Temp || '<FPJE>' || '' || '</FPJE>';
    v_Temp  := v_Temp || '<KPRQ>' || '' || '</KPRQ>';
    Xml_Out := Xmltype('<OUTPUT>' || v_Temp || '</OUTPUT>');
    Return;
  End If;

  --人员id,人员编号,人员姓名 
  v_Temp       := Zl_Identity(1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_操作员编码 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_操作员姓名 := v_Temp;

  If v_结算卡类别 Is Not Null Then
    n_卡类别id := Get_Cardtypeid(v_结算卡类别, 0, v_结算方式);
    If Nvl(n_卡类别id, 0) = 0 Then
      v_Err_Msg := '无法确认传入的结算卡！';
      Raise Err_Item;
    End If;
  End If;

  If Nvl(n_结算模式, 0) = 0 Or Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) = 0 Then
    --n_结算模式 --0-普通模式，1-异步结算模式 
    --n_操作类型 :结算模式为1时传入，0 - 开始结算，1 - 完成结算，2 - 回退结算 
    --1.先进行退费 
    Select 病人结帐记录_Id.Nextval, Sysdate Into n_冲销id, d_退费时间 From Dual;
    n_Billcount  := 0;
    n_原结算序号 := 0;
    For c_费用 In (Select Extractvalue(b.Column_Value, '/FY/DJH') As 单据号, Extractvalue(b.Column_Value, '/FY/XH') As 退款序号
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/FYLIST/FY'))) B) Loop
      Begin
        Select 结算序号, 结帐id, 病人id
        Into n_结算序号, n_结帐id, n_单据病人id
        From 病人预交记录
        Where 结帐id In (Select 结帐id
                       From 门诊费用记录
                       Where NO = c_费用.单据号 And 记录性质 = 1 And Nvl(费用状态, 0) = 0 And 记录状态 In (1, 3) And Rownum < 2) And
              Rownum < 2;
      Exception
        When Others Then
          n_结算序号 := Null;
      End;
    
      If Instr(',' || v_结帐ids || ',', ',' || n_结帐id || ',') = 0 Then
        v_结帐ids := v_结帐ids || ',' || n_结帐id;
        --先要对电子票据冲红处理
        If b_Einvoice_Request.Einvoice_Cancel(1, n_结帐id, v_Err_Msg) = 0 Then
          --电子票据作废失败 
          Raise Err_Item;
        End If;
      End If;
      If n_结算序号 Is Null Then
        v_Err_Msg := '指定的单据号:' || c_费用.单据号 || '未找到,不能退费!';
        Raise Err_Item;
      End If;
    
      --挂号划价模式
      Select Max(Substr(a.摘要, 4))
      Into v_挂号单
      From 门诊费用记录 A
      Where a.No = c_费用.单据号 And a.记录性质 = 1 And a.摘要 Like '挂号:%' And Rownum < 2;
      If v_挂号单 Is Not Null Then
        Select Max(收费单) Into v_收费单 From 病人挂号记录 Where NO = v_挂号单;
        If v_收费单 Is Not Null Then
          Select Count(1)
          Into n_Count
          From 门诊费用记录 A
          Where a.记录性质 = 1 And a.记录状态 = 1 And a.No <> c_费用.单据号 And a.序号 = 1 And
                a.No In (Select /*+ cardinality(b, 10) */
                          Column_Value
                         From Table(f_Str2List(v_收费单)) B);
          If n_Count = 0 Then
            l_挂号单.Extend;
            l_挂号单(l_挂号单.Count) := v_挂号单;
          End If;
        End If;
      End If;
    
      If Nvl(n_单据病人id, 0) = 0 Then
        Select Nvl(Max(病人id), 0)
        Into n_单据病人id
        From 门诊费用记录
        Where NO = c_费用.单据号 And 记录性质 = 1 And Nvl(费用状态, 0) = 0 And 记录状态 In (1, 3) And Rownum < 2;
      End If;
    
      If Nvl(n_病人id, 0) <> Nvl(n_单据病人id, 0) Then
        v_Err_Msg := '本次退费的收费单:' || c_费用.单据号 || '不是当前病人的收费单,不能退费!';
        Raise Err_Item;
      End If;
    
      If n_原结算序号 <> 0 And n_原结算序号 <> n_结算序号 Then
        v_Err_Msg := '本次退费的单据不是一次收费结算,不能退费!';
        Raise Err_Item;
      End If;
    
      n_原结算序号 := n_结算序号;
      Select Count(1) Into n_Count From 费用补充记录 Where 收费结帐id = n_结帐id And Rownum < 2;
      If n_Count <> 0 Then
        v_Err_Msg := '本次退费的单据已经进行了保险补充结算,不能退费!';
        Raise Err_Item;
      End If;
    
      If v_结算卡类别 Is Not Null Then
        Select Count(1) Into n_Count From 病人预交记录 Where 结帐id = n_结帐id And 卡类别id = n_卡类别id;
        If n_Count = 0 Then
          v_Err_Msg := '本次退费的单据不是' || v_结算方式 || '结算的,不能退费!';
          Raise Err_Item;
        End If;
      End If;
    
      If Nvl(n_检查发票, 0) = 1 Then
        Select Max(Decode(a.实际票号, Null, 0, 1))
        Into n_是否打印
        From 门诊费用记录 A
        Where NO = c_费用.单据号 And 记录性质 = 1;
        If Nvl(n_是否打印, 0) = 1 Then
          v_Err_Msg := '本次退费的单据号已开发票,不能退费!';
          Raise Err_Item;
        End If;
      End If;
    
      Zl_门诊收费记录_销帐(c_费用.单据号, v_操作员编码, v_操作员姓名, c_费用.退款序号, d_退费时间, v_摘要, n_冲销id);
      n_Billcount := n_Billcount + 1;
    End Loop;
    If n_Billcount = 0 Then
      v_Err_Msg := '未确定本次需要退费的单据,不能退费!';
      Raise Err_Item;
    End If;
  
    --检查总金额是否正确 
    Select Sum(结帐金额) Into n_结帐金额 From 门诊费用记录 Where 结帐id = n_冲销id;
    n_误差额 := -1 * Nvl(n_结帐金额, 0) - Nvl(n_退款总额, 0);
    If Abs(n_误差额) > 1.00 Then
      v_Err_Msg := '单据缴款金额与实际结算的误差太大!';
      Raise Err_Item;
    End If;
  End If;

  --2.处理退费的结算信息 
  --2.确定支付方式 
  n_删除原结算 := 0;
  If Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) = 1 Then
    n_删除原结算 := 1;
  End If;
  For c_结算方式 In (Select Extractvalue(b.Column_Value, '/TK/TKKLB') As 卡类别, Extractvalue(b.Column_Value, '/TK/TKKH') As 卡号,
                        Extractvalue(b.Column_Value, '/TK/TKFS') As 结算方式,
                        -1 * To_Number(Extractvalue(b.Column_Value, '/TK/TKJE')) As 退款金额,
                        Extractvalue(b.Column_Value, '/TK/JYLSH') As 交易流水号,
                        Extractvalue(b.Column_Value, '/TK/JYSM') As 交易说明, Extractvalue(b.Column_Value, '/TK/TKZY') As 摘要,
                        Extractvalue(b.Column_Value, '/TK/TYJK') As 是否退预交,
                        Extractvalue(b.Column_Value, '/TK/SFXFK') As 是否消费卡,
                        Extractvalue(b.Column_Value, '/TK/DJH') As 单据号,
                        Extract(b.Column_Value, '/TK/EXPENDLIST') As Expend
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/TKLIST/TK'))) B) Loop
  
    n_消费卡 := Nvl(c_结算方式.是否消费卡, 0);
    --1.三方卡结算 
    If c_结算方式.卡类别 Is Not Null And n_消费卡 = 0 Then
      n_卡类别id := Get_Cardtypeid(c_结算方式.卡类别, 0, v_结算方式);
      Select Max(关联交易id)
      Into n_关联交易id
      From 病人预交记录 A,
           (Select a.结帐id
             From 门诊费用记录 A, 门诊费用记录 B
             Where a.No = b.No And Mod(a.记录性质, 10) = 1 And a.记录状态 In (1, 3) And b.结帐id = n_冲销id) B
      Where a.结帐id = b.结帐id And Mod(a.记录性质, 10) <> 1 And a.卡类别id = n_卡类别id;
    
      Thirdcard_Balance(n_病人id, n_冲销id, Nvl(c_结算方式.结算方式, v_结算方式), n_卡类别id, c_结算方式.卡号, c_结算方式.退款金额, c_结算方式.交易流水号,
                        c_结算方式.交易说明, c_结算方式.单据号, n_关联交易id, c_结算方式.摘要, c_结算方式.Expend, n_结算模式, n_操作类型, n_删除原结算);
      n_删除原结算 := 0;
    
      --完成结算时才处理非三方卡结算的 
    Elsif Nvl(n_结算模式, 0) = 0 Or Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) = 1 Then
      --2.消费卡结算 
      If c_结算方式.卡类别 Is Not Null And n_消费卡 = 1 Then
        n_卡类别id := Get_Cardtypeid(c_结算方式.卡类别, 1, v_结算方式);
        Squarecard_Balance(n_病人id, n_冲销id, n_卡类别id, c_结算方式.卡号, c_结算方式.退款金额, c_结算方式.交易流水号, c_结算方式.交易说明, c_结算方式.Expend);
      
        --3.退预交款 
      Elsif Nvl(c_结算方式.是否退预交, 0) = 1 Then
        Zl_门诊退费结算_Modify(1, n_病人id, n_冲销id, Null, c_结算方式.退款金额, Null, Null, Null, Null, 0, 0, 0, 0);
      
        --4.普通结算 
      Else
        If c_结算方式.结算方式 Is Null Then
          v_Err_Msg := '未指定指付方式，不允缴款!';
          Raise Err_Item;
        End If;
      
        --结算方式|结算金额|结算号码|结算摘要||.. 
        v_退费结算 := c_结算方式.结算方式 || '|' || c_结算方式.退款金额 || '| |' || Nvl(c_结算方式.摘要, '  ');
        v_普通结算 := Nvl(v_普通结算, '') || '||' || v_退费结算;
      End If;
    End If;
  End Loop;

  If Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) = 0 Then
    --n_结算模式 --0-普通模式，1-异步结算模式 
    --n_操作类型 :结算模式为1时传入，0 - 开始结算，1 - 完成结算，2 - 回退结算 
  
    v_Temp  := '<CZSJ>' || To_Char(d_退费时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    v_Temp  := v_Temp || '<YJZID>' || v_结帐ids || '</YJZID>';
    v_Temp  := v_Temp || '<CXID>' || n_冲销id || '</CXID>';
    v_Temp  := v_Temp || '<KPBZ>' || 0 || '</KPBZ>';
    v_Temp  := v_Temp || '<URL>' || '' || '</URL>';
    v_Temp  := v_Temp || '<NETURL>' || '' || '</NETURL>';
    v_Temp  := v_Temp || '<FPTT>' || v_姓名 || '</FPTT>';
    v_Temp  := v_Temp || '<FPH>' || '' || '</FPH>';
    v_Temp  := v_Temp || '<FPJE>' || '' || '</FPJE>';
    v_Temp  := v_Temp || '<KPRQ>' || '' || '</KPRQ>';
    Xml_Out := Xmltype('<OUTPUT>' || v_Temp || '</OUTPUT>');
  
    Return;
  End If;

  --5.普通结算及完成结算 
  If v_普通结算 Is Not Null Then
    v_普通结算 := Substr(v_普通结算, 3);
  End If;

  Zl_门诊退费结算_Modify(1, n_病人id, n_冲销id, v_普通结算, 0, Null, Null, Null, Null, 0, 0, n_误差额, 2);

  If v_结帐ids Is Not Null Then
    v_结帐ids := Substr(v_结帐ids, 2);
  End If;

  If l_挂号单.Count <> 0 Then
    For I In 0 .. l_挂号单.Count Loop
      v_Temp := '<GHDH>' || l_挂号单(I) || '</GHDH>';
      v_Temp := v_Temp || '<JSKLB>' || v_结算卡类别 || '</JSKLB>';
      v_Temp := v_Temp || '<GHJE>' || 0 || '</GHJE>';
      Zl_Third_Registdel(Xmltype('<IN>' || v_Temp || '</IN>'), Xml_Out);
    End Loop;
  End If;

  If Nvl(n_结算模式, 0) = 0 Or Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) = 1 Then
    --n_结算模式 --0-普通模式，1-异步结算模式 
    --n_操作类型 :结算模式为1时传入，0 - 开始结算，1 - 完成结算，2 - 回退结算 
    --电子票据处理
  
    For c_原结帐 In (Select Distinct a.结帐id
                  From 门诊费用记录 A, 门诊费用记录 B
                  Where a.No = b.No And a.记录性质 = b.记录性质 And a.记录性质 = 1 And a.记录状态 In (1, 3) And b.结帐id = n_冲销id) Loop
      Select Sum(结帐金额)
      Into n_剩余金额
      From 门诊费用记录
      Where NO In (Select Distinct NO From 门诊费用记录 Where 结帐id = c_原结帐.结帐id) And Mod(记录性质, 10) = 1;
    
      Select Max(是否电子票据) Into n_是否电子票据 From 病人预交记录 Where 结帐id = c_原结帐.结帐id;
    
      If Nvl(n_剩余金额, 0) <> 0 And Nvl(n_是否电子票据, 0) = 1 Then
        --部分退，需要重新开具电子票据
        If b_Einvoice_Request.Einvoice_Create(1, c_原结帐.结帐id, n_冲销id, v_Err_Msg) = 0 Then
          Raise Err_Item;
        End If;
        Select Max(1), Max(姓名), Max(号码), Max(To_Char(生成时间, 'yyyy-mm-dd')), Max(Url内网), Max(Url外网), Max(票据金额)
        Into n_开票标志, v_患者姓名, v_发票编号, v_开票日期, v_Url, v_Url外网, n_发票金额
        From 电子票据使用记录
        Where 结算id = n_结帐id And 票种 = 1 And 记录状态 = 1;
      
        If v_患者姓名 Is Not Null Then
          v_姓名 := v_患者姓名;
        End If;
      End If;
    End Loop;
  End If;
  v_Temp  := '<CZSJ>' || To_Char(d_退费时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
  v_Temp  := v_Temp || '<YJZID>' || v_结帐ids || '</YJZID>';
  v_Temp  := v_Temp || '<CXID>' || n_冲销id || '</CXID>';
  v_Temp  := v_Temp || '<KPBZ>' || Nvl(n_开票标志, 0) || '</KPBZ>';
  v_Temp  := v_Temp || '<URL>' || Nvl(v_Url, '') || '</URL>';
  v_Temp  := v_Temp || '<NETURL>' || Nvl(v_Url外网, '') || '</NETURL>';
  v_Temp  := v_Temp || '<FPTT>' || v_姓名 || '</FPTT>';
  v_Temp  := v_Temp || '<FPH>' || v_发票编号 || '</FPH>';
  v_Temp  := v_Temp || '<FPJE>' || Nvl(n_发票金额, 0) || '</FPJE>';
  v_Temp  := v_Temp || '<KPRQ>' || v_开票日期 || '</KPRQ>';
  Xml_Out := Xmltype('<OUTPUT>' || v_Temp || '</OUTPUT>');

Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Charge_Del;
/

Create Or Replace Procedure Zl_Third_Deposit_Recharge
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:预交款充值
  --入参:Xml_In:
  --    <IN>
  --        <BRID>病人ID</BRID>
  --        <ZYID>主页ID</ZYID>
  --        <XM>姓名</XM>
  --        <SFZH>身份证号</SFZH>
  --        <SFMZ>是否门诊</SFMZ>   //1-是门诊,0-住院
  --        <SFYJ>是否押金</SFYJ>   //是否为押金：0-预交缴款，1-押金缴款
  --        <JSMS>0</JSMS>          //结算模式：0-普通模式，1-异步结算模式
  --        <CZLX>0</CZLX>          //操作类型：结算模式为1时传入，0-开始结算，1-完成结算，2-回退结算
  --        <YJDH></YJDH>           //预交单号，操作类型为1或2时传入
  --    <ZFBZH>支付宝公众号UserID</ZFBZH>
  --    <ZFBXCY>支付宝小程序UserID</ZFBXCY>
  --    <WXGZHID>微信公众号OpenID</WXGZH>
  --    <WXXCXID>微信小程序OpenID</WXXCXID>
  --        <JSLIST>                //结算列表，操作类型为2时可不传入
  --            <JS>
  --              <JSKLB>支付卡类别</JSKLB >
  --              <JSKH>支付卡号</JSKH>
  --              <JYLSH>交易流水号</JYLSH>
  --              <JYSM>交易说明</JYSM>
  --              <JSFS>支付方式</JSFS> //充值方式:现金;支票,如果是三方卡,可以传空
  --              <JSJE>交易金额</JSJE> //充值金额
  --              <ZY>摘要</ZY>
  --              <SFXFK>是否消费卡</SFXFK>
  --              <JSHM>结算号码(可以不传)</JSHM>
  --              <JKDW>缴款单位(可以不传)</JKDW>
  --              <DWKFH>单位开户行(可以不传)</DWKFH>
  --              <DWZH>单位帐号(可以不传)</DWZH>
  --              <HZDW>合作单位(可以不传)</HZDW>
  --              <EXPENDLIST>         //扩展交易信息
  --                   <EXPEND>
  --                        <JYMC>交易名称</JYMC>
  --                        <JYLR>交易内容</JYLR>
  --                   </EXPEND>
  --              </EXPENDLIST >
  --            </JS>
  --         </JSLIST>
  --    </IN>
  --出参:Xml_Out
  --  <OUTPUT>
  --    <CZSJ>操作时间</CZSJ>          //HIS的登记时间
  --    <YJDH>预交单号(多个逗号分隔)</YJDH>
  --  <KPBZ>开票标志</KPBZ> //1-成功开具电子票据;0-未开票成功标志
  --  <URL>H5页面URL</URL>
  --  <NETURL>外网H5页面URL</NETURL>
  --  <FPTT>发票抬头</FPTT>        //病人姓名
  --  <FPH>发票号</FPH>             //发票编号
  --  <FPJE>发票金额</FPJE>        //100.00
  --  <KPRQ>开票日期</KPRQ>   //yyyy-mm-dd
  --    ――如无下列错误结点则说明正确执行
  --    <ERROR>
  --      <MSG>错误信息</MSG>
  --    </ERROR>
  --  </OUTPUT>
  --------------------------------------------------------------------------------------------------
  v_结算方式   Varchar2(2000);
  v_No         病人预交记录.No%Type;
  v_操作员编码 病人预交记录.操作员编号%Type;
  v_操作员姓名 病人预交记录.操作员姓名%Type;
  n_结算模式   Number(2); --0-普通模式，1-异步结算模式
  n_操作类型   Number(2); --结算模式为1时传入，0-开始结算，1-完成结算，2-回退结算 

  n_卡类别id           医疗卡类别.Id%Type;
  n_病人id             门诊费用记录.病人id%Type;
  v_姓名               病人信息.姓名%Type;
  v_身份证号           病人信息.身份证号%Type;
  n_主页id             病人预交记录.主页id%Type;
  n_科室id             病人预交记录.科室id%Type;
  n_预交id             病人预交记录.Id%Type;
  n_结算卡序号         病人预交记录.结算卡序号%Type;
  n_预交类别           病人预交记录.预交类别%Type;
  n_消费卡             Number(2);
  n_门诊预存           Number(2);
  v_卡类别             三方交易记录.类别%Type;
  n_Step               Number(2);
  d_登记时间           Date;
  n_类型               Number(1);
  n_状态               Number(1);
  n_校对标志           病人预交记录.校对标志%Type;
  n_Billcount          Number;
  n_是否押金           Number(1);
  n_预交电子票据       病人预交记录.预交电子票据%Type;
  v_支付宝公众号userid Varchar2(100);
  v_支付宝小程序userid Varchar2(100);
  v_微信公众号openid   Varchar2(100);
  v_微信小程序openid   Varchar2(100);
  n_开票标志           Number(2);
  v_患者姓名           电子票据使用记录.姓名%Type;
  v_发票编号           电子票据使用记录.号码%Type;
  v_开票日期           Varchar2(20);
  n_发票金额           电子票据使用记录.票据金额%Type;
  v_Url                电子票据使用记录.Url内网%Type;
  v_Url外网            电子票据使用记录.Url外网%Type;

  v_Temp    Varchar2(32767); --临时XML
  v_Err_Msg Varchar2(200);
  Err_Special Exception;
  Err_Item    Exception;

  Function Get卡名称
  (
    卡类别_In Varchar2,
    消费卡_In Number
  ) Return Varchar2 As
    v_卡类别 Varchar2(200);
    n_Num    Number(1);
  Begin
    Select Decode(Translate(卡类别_In, '#1234567890', '#'), Null, 1, 0) Into n_Num From Dual;
    If Nvl(消费卡_In, 0) = 1 Then
      If Nvl(n_Num, 0) = 1 Then
        Select Max(名称) Into v_卡类别 From 消费卡类别目录 Where 编号 = To_Number(卡类别_In);
      Else
        Select Max(名称) Into v_卡类别 From 消费卡类别目录 Where 名称 = 卡类别_In;
      End If;
    Else
      If Nvl(n_Num, 0) = 1 Then
        Select Max(名称) Into v_卡类别 From 医疗卡类别 Where ID = To_Number(卡类别_In);
      Else
        Select Max(名称) Into v_卡类别 From 医疗卡类别 Where 名称 = 卡类别_In;
      End If;
    End If;
    Return v_卡类别;
  End Get卡名称;

  Function Get结算方式
  (
    卡类别_In    Varchar2,
    消费卡_In    Number,
    卡类别id_Out Out 病人预交记录.卡类别id%Type
  ) Return Varchar2 As
    --卡类别_In 卡类别名称
  Begin
    If Nvl(消费卡_In, 0) = 1 Then
      Begin
        Select 编号, 结算方式, Decode(Nvl(启用, 0), 1, Null, 名称 || '未启用，不允许进行缴费！')
        Into 卡类别id_Out, v_结算方式, v_Err_Msg
        From 消费卡类别目录
        Where 名称 = 卡类别_In;
      Exception
        When Others Then
          v_Err_Msg := 卡类别_In || '不存在！';
      End;
    
      If v_Err_Msg Is Not Null Then
        Raise Err_Item;
      End If;
      If v_结算方式 Is Null Then
        v_Err_Msg := 卡类别_In || '未设置结算方式，请在消费卡管理中设置结算方式！';
        Raise Err_Item;
      End If;
    Else
      Begin
        Select ID, 结算方式, Decode(Nvl(是否启用, 0), 1, Null, 名称 || '未启用，不允许进行缴费！')
        Into 卡类别id_Out, v_结算方式, v_Err_Msg
        From 医疗卡类别
        Where 名称 = 卡类别_In;
      Exception
        When Others Then
          v_Err_Msg := 卡类别_In || '不存在！';
      End;
    
      If v_Err_Msg Is Not Null Then
        Raise Err_Item;
      End If;
      If v_结算方式 Is Null Then
        v_Err_Msg := 卡类别_In || '未设置结算方式，请在医疗卡类别中设置结算方式！';
        Raise Err_Item;
      End If;
    End If;
  
    Return v_结算方式;
  End Get结算方式;
Begin
  Select Extractvalue(Value(A), 'IN/BRID'), To_Number(Extractvalue(Value(A), 'IN/ZYID')),
         To_Number(Extractvalue(Value(A), 'IN/SFMZ')), Extractvalue(Value(A), 'IN/SFZH'),
         Extractvalue(Value(A), 'IN/XM'), Extractvalue(Value(A), 'IN/JSMS'), Extractvalue(Value(A), 'IN/CZLX'),
         Extractvalue(Value(A), 'IN/YJDH'), Extractvalue(Value(A), 'IN/SFYJ'), Extractvalue(Value(A), 'IN/ZFBZH'),
         Extractvalue(Value(A), 'IN/ZFBXCY'), Extractvalue(Value(A), 'IN/WXGZHID'), Extractvalue(Value(A), 'IN/WXXCXID')
  Into n_病人id, n_主页id, n_门诊预存, v_身份证号, v_姓名, n_结算模式, n_操作类型, v_No, n_是否押金, v_支付宝公众号userid, v_支付宝小程序userid, v_微信公众号openid,
       v_微信小程序openid
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If Nvl(n_门诊预存, 0) = 1 And Nvl(n_病人id, 0) = 0 Then
    If Not v_身份证号 Is Null And Not v_姓名 Is Null Then
      n_病人id := Zl_Third_Getpatiid(v_身份证号, v_姓名);
    End If;
  End If;

  --0.相关检查
  If Nvl(n_病人id, 0) = 0 Then
    v_Err_Msg := '不能有效识别病人身份，不允许充值！';
    Raise Err_Item;
  End If;

  If Not v_支付宝公众号userid Is Null Then
    Zl_病人信息从表_Update(n_病人id, Upper('支付宝公众号UserID'), v_支付宝公众号userid);
  End If;

  If Not v_支付宝小程序userid Is Null Then
    Zl_病人信息从表_Update(n_病人id, Upper('支付宝小程序UserID'), v_支付宝公众号userid);
  End If;

  If Not v_微信公众号openid Is Null Then
    Zl_病人信息从表_Update(n_病人id, Upper('微信公众号OpenID'), v_支付宝公众号userid);
  End If;

  If Not v_微信小程序openid Is Null Then
    Zl_病人信息从表_Update(n_病人id, Upper('微信小程序OpenID'), v_支付宝公众号userid);
  End If;

  Begin
    Select Nullif(Nvl(a.当前科室id, b.出院科室id), 0)
    Into n_科室id
    From 病人信息 A, 病案主页 B
    Where a.病人id = b.病人id(+) And a.主页id = b.主页id(+) And a.病人id = n_病人id;
  Exception
    When Others Then
      v_Err_Msg := '不能有效识别病人身份，不允许充值！';
      Raise Err_Item;
  End;

  If Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) <> 0 Then
    If v_No Is Null Then
      v_Err_Msg := '没有指定相关的结算数据！';
      Raise Err_Item;
    End If;
  
    Begin
      If n_是否押金 = 1 Then
        Select 收款时间, ID
        Into d_登记时间, n_预交id
        From 病人押金记录
        Where 记录状态 = 0 And NO = v_No And Nvl(校对标志, 0) = 1 And Rownum < 2;
      Else
        Select 收款时间, ID
        Into d_登记时间, n_预交id
        From 病人预交记录
        Where 记录性质 = 1 And 记录状态 = 0 And NO = v_No And Nvl(校对标志, 0) = 1 And Rownum < 2;
      End If;
    Exception
      When Others Then
        v_Err_Msg := '没有找到指定的相关结算数据，可能已被处理！';
        Raise Err_Item;
    End;
  End If;

  If Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) = 2 Then
    --删除结算数据，恢复划价单
    If n_是否押金 = 1 Then
      Zl_病人押金异常记录_Delete(v_No);
    Else
      Zl_病人预交异常记录_Delete(v_No);
    End If;
    v_Temp  := '<YJDH>' || v_No || '</YJDH>';
    v_Temp  := v_Temp || '<CZSJ>' || To_Char(d_登记时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    v_Temp  := v_Temp || '<KPBZ>' || 0 || '</KPBZ>';
    v_Temp  := v_Temp || '<URL>' || '' || '</URL>';
    v_Temp  := v_Temp || '<NETURL>' || '' || '</NETURL>';
    v_Temp  := v_Temp || '<FPTT>' || v_姓名 || '</FPTT>';
    v_Temp  := v_Temp || '<FPH>' || '' || '</FPH>';
    v_Temp  := v_Temp || '<FPJE>' || '' || '</FPJE>';
    v_Temp  := v_Temp || '<KPRQ>' || '' || '</KPRQ>';
    Xml_Out := Xmltype('<OUTPUT>' || v_Temp || '</OUTPUT>');
    Return;
  End If;

  --操作员信息:部门ID,部门名称;人员ID,人员编号,人员姓名
  v_Temp := Zl_Identity;
  If Nvl(v_Temp, '0') = '0' Or Nvl(v_Temp, '_') = '_' Then
    v_Err_Msg := '不能识别有效的操作员，不允许缴费！';
    Raise Err_Item;
  End If;
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_操作员编码 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_操作员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);

  If Nvl(n_结算模式, 0) = 0 Or Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) = 0 Then
    For c_交易记录 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                          Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                          Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                          Extractvalue(b.Column_Value, '/JS/SFXFK') As 是否消费卡,
                          Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号, Extractvalue(b.Column_Value, '/JS/ZY') As 摘要
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
    
      If c_交易记录.结算卡类别 Is Null Then
        v_卡类别 := c_交易记录.结算方式;
      Else
        v_卡类别 := Get卡名称(c_交易记录.结算卡类别, c_交易记录.是否消费卡);
      End If;
      If v_卡类别 Is Null Then
        v_Err_Msg := '不支持的结算方式，请检查！';
        Raise Err_Item;
      End If;
    
      --仅第一个结算方式才检查交易锁
      n_Step := Nvl(n_Step, 0) + 1;
      If Zl_Fun_三方交易记录_Locked(v_卡类别, c_交易记录.交易流水号, c_交易记录.结算卡号, c_交易记录.摘要, 1) = 0 And n_Step = 1 Then
        v_Err_Msg := '交易流水号为:' || c_交易记录.交易流水号 || '的交易正在进行中，不允许再次提交此交易！';
        Raise Err_Special;
      End If;
    End Loop;
  End If;

  --2.确定支付方式 
  If Nvl(n_门诊预存, 0) = 0 Then
    n_预交类别 := 2;
  Else
    n_预交类别 := 1;
  End If;
  d_登记时间  := Sysdate;
  n_Billcount := 0;
  For c_结算方式 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                        Extractvalue(b.Column_Value, '/JS/JSJE') As 结算金额,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                        Extractvalue(b.Column_Value, '/JS/JYSM') As 交易说明, Extractvalue(b.Column_Value, '/JS/ZY') As 摘要,
                        Extractvalue(b.Column_Value, '/JS/SFXFK') As 是否消费卡,
                        Extractvalue(b.Column_Value, '/JS/JSHM') As 结算号码,
                        Extractvalue(b.Column_Value, '/JS/JKDW') As 缴款单位,
                        Extractvalue(b.Column_Value, '/JS/DWKFH') As 单位开户行,
                        Extractvalue(b.Column_Value, '/JS/DWZH') As 单位帐号,
                        Extractvalue(b.Column_Value, '/JS/HZDW') As 合作单位,
                        Extract(b.Column_Value, '/JS/EXPENDLIST') As Expend
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
  
    If Nvl(c_结算方式.结算金额, 0) = 0 Then
      v_Err_Msg := '传入的充值金额为零，没必要进行充值处理，请检查充值金额是否传入错误！';
      Raise Err_Item;
    End If;
  
    n_消费卡     := Nvl(c_结算方式.是否消费卡, 0);
    n_结算卡序号 := Null;
    n_卡类别id   := Null;
    If c_结算方式.结算卡类别 Is Null Then
      v_卡类别   := c_结算方式.结算方式;
      v_结算方式 := c_结算方式.结算方式;
    Else
      v_卡类别   := Get卡名称(c_结算方式.结算卡类别, n_消费卡);
      v_结算方式 := Get结算方式(v_卡类别, n_消费卡, n_卡类别id);
      If Nvl(n_消费卡, 0) = 1 Then
        n_结算卡序号 := n_卡类别id;
        n_卡类别id   := Null;
      End If;
    End If;
    If v_结算方式 Is Null Then
      v_Err_Msg := '未确定本次充值的支付方式，请检查支付方式是否传入错误！';
      Raise Err_Item;
    End If;
  
    If Nvl(n_结算模式, 0) = 0 Or Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) = 0 Then
      If Nvl(n_Billcount, 0) > 0 Then
        v_Err_Msg := '目前只支持一种支付方式，请检查充值信息是否传入错误！';
        Raise Err_Item;
      End If;
    
      v_No := Nextno(11);
      Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
    End If;
  
    --操作类型_In:0-正常缴预交;1-保存为未生效的预交款;3-余额退款
    --操作状态_In:0-正常结算，1-保存为异常单据，2-完成异常结算 
    If Nvl(n_结算模式, 0) = 0 Then
      n_类型     := 0;
      n_状态     := 0;
      n_校对标志 := 0;
    Else
      If Nvl(n_操作类型, 0) = 0 Then
        n_类型     := 1;
        n_状态     := 1;
        n_校对标志 := 1;
      Else
        n_类型     := 0;
        n_状态     := 2;
        n_校对标志 := 0;
      End If;
    End If;
    If n_是否押金 = 1 Then
      Zl_病人押金记录_Insert(n_预交id, v_No, Null, n_病人id, n_主页id, n_科室id, c_结算方式.缴款单位, c_结算方式.单位开户行, c_结算方式.单位帐号, c_结算方式.摘要,
                       c_结算方式.结算金额, v_结算方式, c_结算方式.结算号码, n_预交类别, Null, v_操作员编码, v_操作员姓名, c_结算方式.结算卡号, n_卡类别id,
                       c_结算方式.交易流水号, c_结算方式.交易说明, n_校对标志, n_状态);
    Else
      Zl_病人预交记录_Insert(n_预交id, v_No, Null, n_病人id, n_主页id, n_科室id, c_结算方式.结算金额, v_结算方式, c_结算方式.结算号码, c_结算方式.缴款单位,
                       c_结算方式.单位开户行, c_结算方式.单位帐号, c_结算方式.摘要, v_操作员编码, v_操作员姓名, Null, n_预交类别, n_卡类别id, n_结算卡序号,
                       c_结算方式.结算卡号, c_结算方式.交易流水号, c_结算方式.交易说明, c_结算方式.合作单位, d_登记时间, n_类型, Null, Null, 0, 0, 1, 0, n_校对标志,
                       n_状态, 0);
    
    End If;
    If Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) = 0 Then
      v_Temp := '<YJDH>' || v_No || '</YJDH>';
      v_Temp := v_Temp || '<CZSJ>' || To_Char(d_登记时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
      v_Temp := v_Temp || '<KPBZ>' || 0 || '</KPBZ>';
      v_Temp := v_Temp || '<URL>' || '' || '</URL>';
      v_Temp := v_Temp || '<NETURL>' || '' || '</NETURL>';
      v_Temp := v_Temp || '<FPTT>' || v_姓名 || '</FPTT>';
      v_Temp := v_Temp || '<FPH>' || '' || '</FPH>';
      v_Temp := v_Temp || '<FPJE>' || '' || '</FPJE>';
      v_Temp := v_Temp || '<KPRQ>' || '' || '</KPRQ>';
    
      Xml_Out := Xmltype('<OUTPUT>' || v_Temp || '</OUTPUT>');
      Return;
    End If;
  
    If Nvl(n_消费卡, 0) = 1 Then
      Zl_三方接口更新_Update(n_结算卡序号, 1, c_结算方式.结算卡号, n_预交id, c_结算方式.交易流水号, c_结算方式.交易说明, 1, 0);
    End If;
  
    --保存扩展结算信息
    For c_扩展 In (Select Extractvalue(j.Column_Value, '/EXPEND/JYMC') As Jymc,
                        Extractvalue(j.Column_Value, '/EXPEND/JYLR') As Jylr
                 From Table(Xmlsequence(Extract(c_结算方式.Expend, '/EXPENDLIST/EXPEND'))) J) Loop
    
      If Nvl(n_消费卡, 0) = 1 Then
        Zl_三方结算交易_Insert(n_结算卡序号, 1, c_结算方式.结算卡号, n_预交id, c_扩展.Jymc || '|' || c_扩展.Jylr, 1);
      Else
        Zl_三方结算交易_Insert(n_卡类别id, 0, c_结算方式.结算卡号, n_预交id, c_扩展.Jymc || '|' || c_扩展.Jylr, 1 + Nvl(n_是否押金, 0));
      End If;
    End Loop;
    Update 三方交易记录
    Set 业务结算id = n_预交id
    Where 流水号 = c_结算方式.交易流水号 And 类别 = v_卡类别 And 业务类型 = 1;
  
    n_Billcount := n_Billcount + 1;
  End Loop;

  If Nvl(n_Billcount, 0) = 0 Then
    v_Err_Msg := '不能有效确认当前充值的支付方式！';
    Raise Err_Item;
  End If;
  If Nvl(n_是否押金, 0) = 0 Then
    --电子票据处理  
    n_预交电子票据 := b_Einvoice_Request.Einvoice_Start(2, Null,n_门诊预存);
    Update 病人预交记录 Set 预交电子票据 = n_预交电子票据 Where ID = n_预交id;
    --需要开具电子票据
    If Nvl(n_预交电子票据, 0) = 1 Then
      If b_Einvoice_Request.Einvoice_Create(2, n_预交id, Null, v_Err_Msg) = 0 Then
        --电子票据开具成功
        Raise Err_Item;
      End If;
    
      Select Max(1), Max(姓名), Max(号码), Max(To_Char(生成时间, 'yyyy-mm-dd')), Max(Url内网), Max(Url外网), Max(票据金额)
      Into n_开票标志, v_患者姓名, v_发票编号, v_开票日期, v_Url, v_Url外网, n_发票金额
      From 电子票据使用记录
      Where 结算id = n_预交id And 票种 = 2 And 记录状态 = 1;
    
      If v_患者姓名 Is Not Null Then
        v_姓名 := v_患者姓名;
      End If;
    
    End If;
  End If;
  v_Temp  := '<YJDH>' || v_No || '</YJDH>';
  v_Temp  := v_Temp || '<CZSJ>' || To_Char(d_登记时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
  v_Temp  := v_Temp || '<KPBZ>' || Nvl(n_开票标志, 0) || '</KPBZ>';
  v_Temp  := v_Temp || '<URL>' || Nvl(v_Url, '') || '</URL>';
  v_Temp  := v_Temp || '<NETURL>' || Nvl(v_Url外网, '') || '</NETURL>';
  v_Temp  := v_Temp || '<FPTT>' || v_姓名 || '</FPTT>';
  v_Temp  := v_Temp || '<FPH>' || v_发票编号 || '</FPH>';
  v_Temp  := v_Temp || '<FPJE>' || Nvl(n_发票金额, 0) || '</FPJE>';
  v_Temp  := v_Temp || '<KPRQ>' || v_开票日期 || '</KPRQ>';
  Xml_Out := Xmltype('<OUTPUT>' || v_Temp || '</OUTPUT>');
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Err_Special Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20105, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Deposit_Recharge;
/

Create Or Replace Procedure Zl_Third_Settlement
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:三方接口支付
  --入参:Xml_In:
  --<IN>
  --  <BRID>病人ID</BRID>         //病人ID
  --  <XM>姓名</XM>               //姓名
  --  <SFZH>身份证号</SFZH>       //身份证号
  --  <ZYID>主页ID</ZYID>         //主页ID
  --  <JSLX>2</JSLX>         //结算类型,1-门诊,2-住院，默认为 2
  --  <JE></JE>         //本次结算总金额
  --  <NO></NO>         //结帐的费用单据号(门诊记帐单),目前仅结算类型=1时候使用
  --  <JZKNO></JZKNO>   //结帐的就诊卡单据号,目前仅结算类型=1时候使用
  --  <JZSJ></JZSJ>     //结帐时间
  --  <JSMS>1</JSMS>    //结算模式：0-普通模式，1-异步结算模式
  --  <CZLX>0</CZLX>    //操作类型：结算模式为1时传入，0-开始结算，1-完成结算，2-回退结算
  --  <JZID>1</JZID>    //结帐ID，操作类型为1或2时传入
  --  <ZFBZH>支付宝公众号UserID</ZFBZH>
  --  <ZFBXCY>支付宝小程序UserID</ZFBXCY>
  --  <WXGZHID>微信公众号OpenID</WXGZH>
  --  <WXXCXID>微信小程序OpenID</WXXCXID>
  --  <JSLIST>          //结算列表，操作类型为2时可不传入 
  --    <JS>
  --      <JSKLB>支付卡类别</JSKLB >
  --      <JSKH>支付卡号</ JSKH >
  --      <JSFS>支付方式</JSFS> //支付方式:现金;支票,如果是三方卡,可以传空
  --      <JSJE>结算金额</JSJE> //结算金额，均为正金额；SFCYJ为1时为总的冲预交金额
  --      <JYLSH>交易流水号</JYLSH>
  --      <JYSM>交易说明</JYSM>
  --      <ZY>摘要</ZY>
  --      <SFXFK>是否消费卡</SFXFK>  //(1-是消费卡),消费卡时,传入结算卡类别,结算卡号,结算金额等接点
  --      <SFCYJ>是否冲预交</SFCYJ>  //是否冲预交，0-结算，1-冲预交.允冲预交时,只填JSJE节点
  --      <CYJLIST>  //冲预交款集合：是否冲预交=1时传入，此时以传入的单据进行冲预交款；如果不传入该节点，则按HIS的先缴先用规则进行冲预交款，此时不能进行退预交款（如果存在退预交款，则只能使用预交款进行结帐）
  --        <TKFS>退款方式<TKFS>  //退款方式，存在退款时传入：0-分交易退款;1-调用一次交易接口退款;2-转帐方式退款(暂不支持)(应以线下一致，主要用于异步模式下，线下窗口处理异常结算)
  --        <ITEM>  //说明：有退必有冲；如单据A001，预交款金额1000，本次结帐800，则传入两条记录①冲1000，②退200，Sum(冲金额-退金额)=结帐金额
  --          <DJH>预交款单据号</DJH>
  --          <JE>交易金额</JE>
  --          <SFTK>是否退预交款</SFTK>  //是否退预交款：0-冲预交款;1-退预交款
  --          <JYLSH>交易流水号</JYLSH>  //退款交易流水号
  --          <JYSM>交易说明</JYSM>  //退款交易说明
  --          <EXPENDLIST>  //退款交易的扩展信息，退款方式=0时传入
  --            <EXPEND>
  --              <JYMC>交易名称</JYMC> //交易名称
  --              <JYLR>交易内容</JYLR> //交易内容
  --            </EXPEND>
  --          </EXPENDLIST>
  --        </ITEM>
  --        <EXPENDLIST>  //退款交易的扩展信息，退款方式=1、2时传入
  --          <EXPEND>      
  --            <JYMC>交易名称</JYMC> //交易名称
  --            <JYLR>交易内容</JYLR> //交易内容
  --          </EXPEND>
  --        </EXPENDLIST>
  --      </CYJLIST>
  --      <EXPENDLIST>  //扩展交易信息
  --        <EXPEND>
  --          <JYMC>交易名称</JYMC> //交易名称
  --          <JYLR>交易内容</JYLR> //交易内容
  --        </EXPEND>
  --      </EXPENDLIST>
  --    </JS>
  --  </JSLIST >
  --</IN>

  --出参:Xml_Out
  --<OUT>
  --  <CZSJ>操作时间</CZSJ>          //HIS的登记时间
  --  <JZID>结帐ID</JZID>
  --  <KPBZ>开票标志</KPBZ> //1-成功开具电子票据;0-未开票成功标志
  --  <URL>H5页面URL</URL>
  --  <NETURL>外网H5页面URL</NETURL>
  --  <FPTT>发票抬头</FPTT>        //病人姓名
  --  <FPH>发票号</FPH>             //发票编号
  --  <FPJE>发票金额</FPJE>        //100.00
  --  <KPRQ>开票日期</KPRQ>   //yyyy-mm-dd
  --  <ERROR>  //如无该错误结点则说明正确执行
  --    <MSG>错误信息</MSG>
  --  </ERROR>
  --</OUT>
  --------------------------------------------------------------------------------------------------
  n_主页id       病案主页.主页id%Type;
  n_病人id       病案主页.病人id%Type;
  v_姓名         病人信息.姓名%Type;
  v_身份证号     病人信息.身份证号%Type;
  n_结帐总额     病人预交记录.冲预交%Type;
  n_结算类型     Number(3);
  v_就诊卡单据号 Varchar2(20000);
  d_结帐时间     Date;
  v_单据号       Varchar2(20000);
  n_结算模式     Number(1); --0-普通模式，1-异步结算模式
  n_操作类型     Number(1); --结算模式为1时传入，0 - 开始结算，1 - 完成结算，2 - 回退结算

  v_操作员编码 病人结帐记录.操作员编号%Type;
  v_操作员姓名 病人结帐记录.操作员姓名%Type;
  n_结帐id     病人结帐记录.Id%Type;
  n_待结帐金额 病人预交记录.冲预交%Type;
  d_开始日期   Date;
  d_结束日期   Date;
  d_最小日期   Date;
  d_最大日期   Date;
  n_关联交易id 病人预交记录.关联交易id%Type;
  n_预交id     病人预交记录.Id%Type;
  n_删除原结算 Number;

  v_结帐单号 病人结帐记录.No%Type;
  d_作废时间 病人结帐记录.收费时间%Type;
  n_冲销id   病人结帐记录.Id%Type;

  n_结算卡序号 消费卡类别目录.编号%Type;
  n_时间类型   Number(3);
  v_No         病人结帐记录.No%Type;
  n_卡类别id   医疗卡类别.Id%Type;
  v_结算方式   病人预交记录.结算方式%Type;
  v_Temp       Varchar2(500);
  v_Ids        Varchar2(20000);
  x_Templet    Xmltype; --模板XML
  v_消费卡结算 Varchar2(20000);
  n_结算金额   病人预交记录.冲预交%Type;
  n_Step       Number(2);

  v_Err_Msg Varchar2(200);
  Err_Item    Exception;
  Err_Special Exception;
  v_卡类别 三方交易记录.类别%Type;
  n_Number Number(2);

  n_费用id   门诊费用记录.Id%Type;
  n_记录性质 门诊费用记录.记录性质%Type;
  v_费用no   门诊费用记录.No%Type;
  n_序号     门诊费用记录.序号%Type;
  n_记录状态 门诊费用记录.记录状态%Type;
  n_执行状态 门诊费用记录.执行状态%Type;
  n_未结金额 门诊费用记录.实收金额%Type;
  n_结帐金额 门诊费用记录.实收金额%Type;
  n_误差费   门诊费用记录.实收金额%Type;
  Type t_费用结算明细 Is Ref Cursor;
  c_费用结算明细 t_费用结算明细;

  Type Ty_预交款 Is Record(
    退款方式   Number(1), --0-分交易退款;1-调用一次交易接口退款;2-转帐方式退款(暂不支持)
    单据号     病人预交记录.No%Type,
    冲预交     病人预交记录.冲预交%Type,
    交易流水号 病人预交记录.交易流水号%Type,
    交易说明   病人预交记录.交易说明%Type,
    
    预交id       病人预交记录.Id%Type,
    结算方式     病人预交记录.结算方式%Type,
    卡类别id     病人预交记录.卡类别id%Type,
    卡号         病人预交记录.卡号%Type,
    原交易流水号 病人预交记录.交易流水号%Type,
    原交易说明   病人预交记录.交易说明%Type,
    关联交易id   病人预交记录.关联交易id%Type,
    是否转账     Number(1),
    交易扩展信息 Xmltype, --分交易退款的扩展信息
    
    是否退款 Number(1),
    扩展信息 Xmltype);
  Type t_预交款 Is Table Of Ty_预交款;
  l_预交款 t_预交款;

  n_非预交款结算       Number(1);
  n_冲预交金额         病人预交记录.冲预交%Type;
  n_退预交金额         病人预交记录.冲预交%Type;
  n_存在退预交         Number(1);
  n_是否电子票据       病人预交记录.是否电子票据%Type;
  v_支付宝公众号userid Varchar2(100);
  v_支付宝小程序userid Varchar2(100);
  v_微信公众号openid   Varchar2(100);
  v_微信小程序openid   Varchar2(100);
  n_开票标志           Number(2);
  v_患者姓名           电子票据使用记录.姓名%Type;
  v_发票编号           电子票据使用记录.号码%Type;
  v_开票日期           Varchar2(20);
  n_发票金额           电子票据使用记录.票据金额%Type;
  v_Url                电子票据使用记录.Url内网%Type;
  v_Url外网            电子票据使用记录.Url外网%Type;

  Procedure 病人预交款_List
  (
    预交款信息_In  Xmltype,
    预交款支付_In  Number,
    预交款_Out     In Out t_预交款,
    存在退预交_Out In Out Number
  ) As
    n_原预交id 病人预交记录.Id%Type;
    n_剩余金额 病人预交记录.冲预交%Type;
  
    v_结算方式     病人预交记录.结算方式%Type;
    n_卡类别id     病人预交记录.卡类别id%Type;
    v_卡号         病人预交记录.卡号%Type;
    v_原交易流水号 病人预交记录.交易流水号%Type;
    v_原交易说明   病人预交记录.交易说明%Type;
    n_关联交易id   病人预交记录.关联交易id%Type;
  
    n_退款方式 Number(1);
    x_扩展信息 Xmltype;
  
    I         Number(18);
    v_Err_Msg Varchar2(200);
    Err_Item Exception;
  Begin
    If 预交款_Out Is Null Then
      预交款_Out := t_预交款();
    End If;
  
    --退款方式:0-分交易退款;1-调用一次交易接口退款;2-转帐方式退款(暂不支持)
    Select Extractvalue(Value(A), 'CYJLIST/TKFS'), Extract(a.Column_Value, 'CYJLIST/EXPENDLIST')
    Into n_退款方式, x_扩展信息
    From Table(Xmlsequence(Extract(预交款信息_In, 'CYJLIST'))) A;
  
    For r_预交款 In (Select Extractvalue(b.Column_Value, 'ITEM/DJH') As 单据号, Extractvalue(b.Column_Value, 'ITEM/JE') As 冲预交,
                         Extractvalue(b.Column_Value, 'ITEM/SFTK') As 是否退款,
                         Extractvalue(b.Column_Value, 'ITEM/JYLSH') As 交易流水号,
                         Extractvalue(b.Column_Value, 'ITEM/JYSM') As 交易说明,
                         Extract(b.Column_Value, 'ITEM/EXPENDLIST') As 扩展信息
                  From Table(Xmlsequence(Extract(预交款信息_In, 'CYJLIST/ITEM'))) B
                  Order By Nvl(是否退款, 0)) Loop
    
      Select Max(Decode(记录性质, 1, ID, 0)), Sum(Nvl(金额, 0) - Nvl(冲预交, 0))
      Into n_原预交id, n_剩余金额
      From 病人预交记录
      Where 记录性质 In (1, 11) And NO = r_预交款.单据号;
      If Nvl(n_原预交id, 0) = 0 Then
        v_Err_Msg := '预交款单据[' || r_预交款.单据号 || ']不存在，结算失败！';
        Raise Err_Item;
      End If;
    
      If Nvl(r_预交款.是否退款, 0) = 0 Then
        If Nvl(n_剩余金额, 0) < Nvl(r_预交款.冲预交, 0) And Nvl(预交款支付_In, 0) = 1 Then
          v_Err_Msg := '预交款单据[' || r_预交款.单据号 || ']余额不足，结算失败！';
          Raise Err_Item;
        End If;
      End If;
    
      Begin
        Select ID, 结算方式, 卡类别id, 卡号, 交易流水号, 交易说明, 关联交易id
        Into n_原预交id, v_结算方式, n_卡类别id, v_卡号, v_原交易流水号, v_原交易说明, n_关联交易id
        From 病人预交记录
        Where 记录性质 = 1 And NO = r_预交款.单据号;
      Exception
        When Others Then
          v_Err_Msg := '预交款单据[' || r_预交款.单据号 || ']不存在，结算失败！';
          Raise Err_Item;
      End;
    
      If Nvl(r_预交款.是否退款, 0) = 1 Then
        存在退预交_Out := 1;
      End If;
    
      预交款_Out.Extend();
      I := 预交款_Out.Count;
      预交款_Out(I).单据号 := r_预交款.单据号;
      预交款_Out(I).冲预交 := r_预交款.冲预交;
      预交款_Out(I).是否退款 := r_预交款.是否退款;
      预交款_Out(I).交易流水号 := r_预交款.交易流水号;
      预交款_Out(I).交易说明 := r_预交款.交易说明;
    
      预交款_Out(I).预交id := n_原预交id;
      预交款_Out(I).结算方式 := v_结算方式;
      预交款_Out(I).卡类别id := n_卡类别id;
      预交款_Out(I).卡号 := v_卡号;
      预交款_Out(I).原交易流水号 := v_原交易流水号;
      预交款_Out(I).原交易说明 := v_原交易说明;
      预交款_Out(I).关联交易id := n_关联交易id;
      If Nvl(n_退款方式, 0) = 2 Then
        预交款_Out(I).是否转账 := 1;
      Else
        预交款_Out(I).是否转账 := 0;
      End If;
      预交款_Out(I).交易扩展信息 := r_预交款.扩展信息;
    
      预交款_Out(I).退款方式 := n_退款方式;
      预交款_Out(I).扩展信息 := x_扩展信息;
    End Loop;
  
    If Nvl(存在退预交_Out, 0) = 1 And Nvl(n_退款方式, 0) = 2 Then
      v_Err_Msg := '预交款暂不支持转账退款，结算失败！';
      Raise Err_Item;
    End If;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  End;

  Procedure 三方结算交易_Save
  (
    结帐id_In   病人预交记录.结帐id%Type,
    卡类别id_In 病人预交记录.卡类别id%Type,
    卡号_In     病人预交记录.卡号%Type,
    扩展信息_In Xmltype
  ) As
  Begin
    If 扩展信息_In Is Null Then
      Return;
    End If;
  
    For c_扩展 In (Select Extractvalue(j.Column_Value, 'EXPEND/JYMC') As 名称,
                        Extractvalue(j.Column_Value, 'EXPEND/JYLR') As 内容
                 From Table(Xmlsequence(Extract(扩展信息_In, 'EXPENDLIST/EXPEND'))) J) Loop
      Zl_三方结算交易_Insert(卡类别id_In, 0, 卡号_In, 结帐id_In, c_扩展.名称 || '|' || c_扩展.内容);
    End Loop;
  End;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/ZYID'), To_Number(Extractvalue(Value(A), 'IN/BRID')),
         To_Number(Extractvalue(Value(A), 'IN/JE')), Nvl(To_Number(Extractvalue(Value(A), 'IN/JSLX')), 2),
         Extractvalue(Value(A), 'IN/NO'), To_Number(Extractvalue(Value(A), 'IN/JZKNO')),
         To_Date(Extractvalue(Value(A), 'IN/JZSJ'), 'yyyy-mm-dd hh24:mi:ss'), Extractvalue(Value(A), 'IN/SFZH'),
         Extractvalue(Value(A), 'IN/XM'), Extractvalue(Value(A), 'IN/JSMS'), Extractvalue(Value(A), 'IN/CZLX'),
         Extractvalue(Value(A), 'IN/JZID'), Extractvalue(Value(A), 'IN/ZFBZH'), Extractvalue(Value(A), 'IN/ZFBXCY'),
         Extractvalue(Value(A), 'IN/WXGZHID'), Extractvalue(Value(A), 'IN/WXXCXID')
  Into n_主页id, n_病人id, n_结帐总额, n_结算类型, v_单据号, v_就诊卡单据号, d_结帐时间, v_身份证号, v_姓名, n_结算模式, n_操作类型, n_结帐id, v_支付宝公众号userid,
       v_支付宝小程序userid, v_微信公众号openid, v_微信小程序openid
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If n_结算类型 = 1 And Nvl(n_病人id, 0) = 0 And Not v_身份证号 Is Null And Not v_姓名 Is Null Then
    n_病人id := Zl_Third_Getpatiid(v_身份证号, v_姓名);
  End If;

  --0.相关检查
  If Nvl(n_病人id, 0) = 0 Then
    v_Err_Msg := '不能有效识别病人身份,不允许结算!';
    Raise Err_Item;
  End If;

  If Not v_支付宝公众号userid Is Null Then
    Zl_病人信息从表_Update(n_病人id, Upper('支付宝公众号UserID'), v_支付宝公众号userid);
  End If;

  If Not v_支付宝小程序userid Is Null Then
    Zl_病人信息从表_Update(n_病人id, Upper('支付宝小程序UserID'), v_支付宝公众号userid);
  End If;

  If Not v_微信公众号openid Is Null Then
    Zl_病人信息从表_Update(n_病人id, Upper('微信公众号OpenID'), v_支付宝公众号userid);
  End If;

  If Not v_微信小程序openid Is Null Then
    Zl_病人信息从表_Update(n_病人id, Upper('微信小程序OpenID'), v_支付宝公众号userid);
  End If;

  If Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) <> 0 Then
    If Nvl(n_结帐id, 0) = 0 Then
      v_Err_Msg := '没有指定相关的结算数据！';
      Raise Err_Item;
    End If;
  
    Begin
      Select 收款时间, 关联交易id, 卡类别id
      Into d_结帐时间, n_关联交易id, n_卡类别id
      From 病人预交记录
      Where 结帐id = n_结帐id And Nvl(校对标志, 0) = 1 And 卡类别id Is Not Null And Rownum < 2;
    Exception
      When Others Then
        v_Err_Msg := '没有找到指定的相关结算数据，可能已被处理！';
        Raise Err_Item;
    End;
  End If;

  v_操作员编码 := Zl_操作员信息(1);
  v_操作员姓名 := Zl_操作员信息(2);

  --【1】异步模式回退交易
  If Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) = 2 Then
    --删除结算数据
    Zl_病人结帐结算_Delete(n_结帐id, n_卡类别id, n_关联交易id);
    --作废原结帐
    Begin
      Select NO Into v_结帐单号 From 病人结帐记录 Where ID = n_结帐id And Nvl(结算状态, 0) = 1 And Rownum < 2;
    Exception
      When Others Then
        v_Err_Msg := '没有找到原结帐数据，可能已被处理！';
        Raise Err_Item;
    End;
  
    d_作废时间 := Sysdate;
    Select 病人结帐记录_Id.Nextval Into n_冲销id From Dual;
    Zl_病人结帐记录_Cancel(v_结帐单号, n_冲销id, v_操作员编码, v_操作员姓名, d_作废时间);
    Zl_病人结帐作废_Modify(0, n_病人id, n_冲销id, Null, Null, Null, Null, Null, Null, Null, Null, Null, v_操作员编码, v_操作员姓名, d_作废时间,
                     Null, 1);
  
    v_Temp := '<CZSJ>' || To_Char(d_结帐时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<JZID>' || n_结帐id || '</JZID>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<KPBZ>' || 0 || '</KPBZ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<URL>' || '' || '</URL>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<NETURL>' || '' || '</NETURL>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<FPTT>' || v_姓名 || '</FPTT>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<FPH>' || '' || '</FPH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<FPJE>' || '' || '</FPJE>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<KPRQ>' || '' || '</KPRQ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    Xml_Out := x_Templet;
    Return;
  End If;

  --【2】非异步模式或异步模式开始交易
  If Nvl(n_结算模式, 0) = 0 Or Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) = 0 Then
    --【2.1】锁定三方卡支付交易
    For c_交易记录 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                          Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                          Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                          Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号, Extractvalue(b.Column_Value, '/JS/ZY') As 摘要,
                          Extractvalue(b.Column_Value, '/JS/SFCYJ') As 是否冲预交,
                          Extractvalue(b.Column_Value, '/JS/SFXFK') As 是否消费卡
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
    
      If Not (c_交易记录.结算卡类别 Is Null Or Nvl(c_交易记录.是否消费卡, '0') = '1' Or Nvl(c_交易记录.是否冲预交, 0) = 1) Then
        Select Decode(Translate(Nvl(c_交易记录.结算卡类别, 'abcd'), '#1234567890', '#'), Null, 1, 0)
        Into n_Number
        From Dual;
      
        If Nvl(n_Number, 0) = 1 Then
          Select Max(名称) Into v_卡类别 From 医疗卡类别 Where ID = To_Number(c_交易记录.结算卡类别);
        Else
          Select Max(名称) Into v_卡类别 From 医疗卡类别 Where 名称 = c_交易记录.结算卡类别;
        End If;
        If v_卡类别 Is Null Then
          v_Err_Msg := '不支持的结算方式,请检查！';
          Raise Err_Item;
        End If;
      
        --仅第一个结算方式才检查交易锁
        n_Step := Nvl(n_Step, 0) + 1;
        If Zl_Fun_三方交易记录_Locked(v_卡类别, c_交易记录.交易流水号, c_交易记录.结算卡号, c_交易记录.摘要, 2) = 0 And n_Step = 1 Then
          v_Err_Msg := '交易流水号为:' || c_交易记录.交易流水号 || '的交易正在进行中，不允许再次提交此交易!';
          Raise Err_Special;
        End If;
      End If;
    End Loop;
  
    --【2.2】费用结帐记录
    Select Nvl(zl_GetSysParameter('结帐费用时间', 1137), 0) Into n_时间类型 From Dual;
    If n_结算类型 = 2 Then
      Open c_费用结算明细 For
        Select Max(Decode(结帐id, Null, ID, Null)) As ID, Mod(记录性质, 10) As 记录性质, NO, 序号, 记录状态, 执行状态,
               Trunc(Min(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最小时间, Trunc(Max(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最大时间,
               Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) As 金额, Sum(Nvl(结帐金额, 0)) As 结帐金额
        From 住院费用记录
        Where 病人id = n_病人id And 记录状态 <> 0 And 主页id = n_主页id And 记帐费用 = 1
        Group By Mod(记录性质, 10), NO, 序号, 记录状态, 执行状态
        Having(Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) <> 0) Or (Sum(Nvl(实收金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Sum(Nvl(结帐金额, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(结帐金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Mod(Count(*), 2) = 0
        Order By NO, 序号;
    Else
      If v_单据号 Is Null And v_就诊卡单据号 Is Null Then
        Open c_费用结算明细 For
          Select Min(Decode(结帐id, Null, ID, Null)) As ID, Mod(记录性质, 10) As 记录性质, NO, 序号, 记录状态, 执行状态,
                 Trunc(Min(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最小时间, Trunc(Max(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最大时间,
                 Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) As 金额, Sum(Nvl(结帐金额, 0)) As 结帐金额
          From 门诊费用记录
          Where 病人id = n_病人id And 记录状态 <> 0 And 记帐费用 = 1
          Group By Mod(记录性质, 10), NO, 序号, 记录状态, 执行状态
          Having(Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) <> 0) Or (Sum(Nvl(实收金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Sum(Nvl(结帐金额, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(结帐金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Mod(Count(*), 2) = 0
          Union All
          Select Min(Decode(结帐id, Null, ID, Null)) As ID, Mod(记录性质, 10) As 记录性质, NO, 序号, 记录状态, 执行状态,
                 Trunc(Min(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最小时间, Trunc(Max(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最大时间,
                 Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) As 金额, Sum(Nvl(结帐金额, 0)) As 结帐金额
          From 住院费用记录
          Where 病人id = n_病人id And 记录状态 <> 0 And 记帐费用 = 1 And Mod(记录性质, 10) = 5
          Group By Mod(记录性质, 10), NO, 序号, 记录状态, 执行状态
          Having(Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) <> 0) Or (Sum(Nvl(实收金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Sum(Nvl(结帐金额, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(结帐金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Mod(Count(*), 2) = 0
          Order By NO, 序号;
      Elsif v_单据号 Is Not Null And v_就诊卡单据号 Is Not Null Then
        Open c_费用结算明细 For
          Select Min(Decode(结帐id, Null, ID, Null)) As ID, Mod(记录性质, 10) As 记录性质, NO, 序号, 记录状态, 执行状态,
                 Trunc(Min(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最小时间, Trunc(Max(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最大时间,
                 Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) As 金额, Sum(Nvl(结帐金额, 0)) As 结帐金额
          From 门诊费用记录
          Where 病人id + 0 = n_病人id And 记录状态 <> 0 And Mod(记录性质, 10) = 2 And 记帐费用 = 1 And
                NO In (Select /*+cardinality(b,10) */
                        Column_Value
                       From Table(f_Str2List(v_单据号)) B)
          Group By Mod(记录性质, 10), NO, 序号, 记录状态, 执行状态
          Having(Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) <> 0) Or (Sum(Nvl(实收金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Sum(Nvl(结帐金额, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(结帐金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Mod(Count(*), 2) = 0
          Union All
          Select Min(Decode(结帐id, Null, ID, Null)) As ID, Mod(记录性质, 10) As 记录性质, NO, 序号, 记录状态, 执行状态,
                 Trunc(Min(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最小时间, Trunc(Max(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最大时间,
                 Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) As 金额, Sum(Nvl(结帐金额, 0)) As 结帐金额
          From 住院费用记录
          Where 病人id + 0 = n_病人id And 记录状态 <> 0 And 记帐费用 = 1 And Mod(记录性质, 10) = 5 And
                NO In (Select /*+cardinality(b,10) */
                        Column_Value
                       From Table(f_Str2List(v_就诊卡单据号)) B)
          Group By Mod(记录性质, 10), NO, 序号, 记录状态, 执行状态
          Having(Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) <> 0) Or (Sum(Nvl(实收金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Sum(Nvl(结帐金额, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(结帐金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Mod(Count(*), 2) = 0
          Order By 记录性质, NO, 序号;
      Elsif v_单据号 Is Not Null Then
        Open c_费用结算明细 For
          Select Min(Decode(结帐id, Null, ID, Null)) As ID, Mod(记录性质, 10) As 记录性质, NO, 序号, 记录状态, 执行状态,
                 Trunc(Min(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最小时间, Trunc(Max(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最大时间,
                 Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) As 金额, Sum(Nvl(结帐金额, 0)) As 结帐金额
          From 门诊费用记录
          Where 病人id + 0 = n_病人id And 记录状态 <> 0 And Mod(记录性质, 10) = 2 And 记帐费用 = 1 And
                NO In (Select /*+cardinality(b,10) */
                        Column_Value
                       From Table(f_Str2List(v_单据号)) B)
          Group By Mod(记录性质, 10), NO, 序号, 记录状态, 执行状态
          Having(Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) <> 0) Or (Sum(Nvl(实收金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Sum(Nvl(结帐金额, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(结帐金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Mod(Count(*), 2) = 0
          Order By NO, 序号;
      Else
        Open c_费用结算明细 For
          Select Min(Decode(结帐id, Null, ID, Null)) As ID, Mod(记录性质, 10) As 记录性质, NO, 序号, 记录状态, 执行状态,
                 Trunc(Min(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最小时间, Trunc(Max(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最大时间,
                 Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) As 金额, Sum(Nvl(结帐金额, 0)) As 结帐金额
          From 住院费用记录
          Where 病人id + 0 = n_病人id And 记录状态 <> 0 And 记帐费用 = 1 And Mod(记录性质, 10) = 5 And
                NO In (Select /*+cardinality(b,10) */
                        Column_Value
                       From Table(f_Str2List(v_就诊卡单据号)) B)
          Group By Mod(记录性质, 10), NO, 序号, 记录状态, 执行状态
          Having(Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) <> 0) Or (Sum(Nvl(实收金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Sum(Nvl(结帐金额, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(结帐金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Mod(Count(*), 2) = 0
          Order By NO, 序号;
      End If;
    End If;
  
    Select 病人结帐记录_Id.Nextval, Sysdate, Nextno(15) Into n_结帐id, d_结帐时间, v_No From Dual;
    n_待结帐金额 := 0;
    Loop
      Fetch c_费用结算明细
        Into n_费用id, n_记录性质, v_费用no, n_序号, n_记录状态, n_执行状态, d_最小日期, d_最大日期, n_未结金额, n_结帐金额;
      Exit When c_费用结算明细%NotFound;
    
      n_待结帐金额 := n_待结帐金额 + Nvl(n_未结金额, 0);
      If d_开始日期 Is Null Then
        d_开始日期 := d_最小日期;
      Elsif d_开始日期 > d_最小日期 Then
        d_开始日期 := d_最小日期;
      End If;
      If d_结束日期 Is Null Then
        d_结束日期 := d_最大日期;
      Elsif d_结束日期 < d_最大日期 Then
        d_结束日期 := d_最大日期;
      End If;
    
      If Nvl(n_结帐金额, 0) = 0 Then
        If n_费用id Is Not Null Then
          If Length(v_Ids || ',' || n_费用id) > 4000 Then
            v_Ids := Substr(v_Ids, 2);
            Zl_结帐费用记录_Batch(v_Ids, n_病人id, n_结帐id);
            v_Ids := '';
          End If;
          v_Ids := v_Ids || ',' || n_费用id;
        Else
          Zl_结帐费用记录_Insert(0, v_费用no, n_记录性质, n_记录状态, n_执行状态, n_序号, n_未结金额, n_结帐id);
        End If;
      Else
        Zl_结帐费用记录_Insert(0, v_费用no, n_记录性质, n_记录状态, n_执行状态, n_序号, n_未结金额, n_结帐id);
      End If;
    End Loop;
  
    If v_Ids Is Not Null Then
      v_Ids := Substr(v_Ids, 2);
      Zl_结帐费用记录_Batch(v_Ids, n_病人id, n_结帐id);
    End If;
  
    If Round(n_待结帐金额, 6) <> Nvl(n_结帐总额, 0) Then
      v_Err_Msg := '传入的结帐金额与实际结帐金额不符,不允许结算!';
      Raise Err_Item;
    End If;
  
    Zl_病人结帐记录_Insert(n_结帐id, v_No, n_病人id, d_结帐时间, d_开始日期, d_结束日期, 0, 0, n_主页id, Null, n_结算类型, Null, n_结算类型, 1, n_主页id,
                     n_结帐总额);
  
    --【2.3】结算数据预先保存,三方卡支付和存在退款的预交款支付
    n_非预交款结算 := 0;
    n_存在退预交   := 0;
    l_预交款       := t_预交款();
    For r_结算方式 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                          Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号,
                          Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                          Extractvalue(b.Column_Value, '/JS/JSJE') As 结算金额,
                          Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                          Extractvalue(b.Column_Value, '/JS/JYSM') As 交易说明, Extractvalue(b.Column_Value, '/JS/ZY') As 摘要,
                          Extractvalue(b.Column_Value, '/JS/SFXFK') As 是否消费卡,
                          Extractvalue(b.Column_Value, '/JS/SFCYJ') As 是否冲预交,
                          Extract(b.Column_Value, '/JS/CYJLIST') As 冲预交列表
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
    
      If Nvl(r_结算方式.是否冲预交, 0) = 0 Then
        n_非预交款结算 := 1;
        n_卡类别id     := Null;
        If r_结算方式.结算卡类别 Is Not Null And Nvl(r_结算方式.是否消费卡, 0) = 0 Then
          Select Decode(Translate(Nvl(r_结算方式.结算卡类别, 'abcd'), '#1234567890', '#'), Null, 1, 0)
          Into n_Number
          From Dual;
        
          If Nvl(n_Number, 0) = 1 Then
            Select Max(ID) Into n_卡类别id From 医疗卡类别 Where ID = r_结算方式.结算卡类别 And Nvl(是否启用, 0) = 1;
          Else
            Select Max(ID) Into n_卡类别id From 医疗卡类别 Where 名称 = r_结算方式.结算卡类别 And Nvl(是否启用, 0) = 1;
          End If;
        
          If n_卡类别id Is Null Then
            v_Err_Msg := '未找到对应的医疗卡信息!';
            Raise Err_Item;
          End If;
        
          Select Max(关联交易id)
          Into n_关联交易id
          From 病人预交记录
          Where 结帐id = n_结帐id And 卡类别id = n_卡类别id And Rownum < 2;
          If Nvl(n_关联交易id, 0) = 0 Then
            Select 病人预交记录_Id.Nextval Into n_关联交易id From Dual;
            n_预交id := n_关联交易id;
          Else
            n_预交id := Null;
          End If;
        
          v_结算方式 := r_结算方式.结算方式 || '|' || r_结算方式.结算金额 || '|';
          Zl_病人结帐结算_Modify(1, n_病人id, n_结帐id, v_结算方式, Null, 0, n_卡类别id, r_结算方式.结算卡号, r_结算方式.交易流水号, r_结算方式.交易说明, 0, 0, 0,
                           n_结算类型, Null, v_操作员编码, v_操作员姓名, d_结帐时间, Null, 0, 1, n_预交id, n_关联交易id);
        End If;
      
      Elsif r_结算方式.冲预交列表 Is Not Null Then
        --【*】按指定单据进行预交款支付
        病人预交款_List(r_结算方式.冲预交列表, 1, l_预交款, n_存在退预交);
      End If;
    End Loop;
  
    --【*】按指定单据进行预交款支付
    If l_预交款.Count > 0 And Nvl(n_存在退预交, 0) = 1 Then
      --  1.如果存在退预交款，则只能使用预交款进行结帐
      If Nvl(n_非预交款结算, 0) = 1 Then
        v_Err_Msg := '存在对预交款进行退款时，全部结帐金额都必须使用预交款进行支付！';
        Raise Err_Item;
      End If;
    
      --先保存预交款数据
      n_冲预交金额 := 0;
      n_退预交金额 := 0;
      For I In 1 .. l_预交款.Count Loop
        If Nvl(l_预交款(I).是否退款, 0) = 0 Then
          --冲预交款  
          Zl_结帐预交记录_Insert(l_预交款(I).预交id, l_预交款(I).单据号, 1, l_预交款(I).冲预交, n_结帐id, n_病人id, v_操作员编码, v_操作员姓名, d_结帐时间);
        
          n_冲预交金额 := n_冲预交金额 + Nvl(l_预交款(I).冲预交, 0);
        Else
          --退预交款  
          Zl_三方退款信息_Insert(n_结帐id, l_预交款(I).预交id, l_预交款(I).冲预交, l_预交款(I).卡号, Null, Null, 0, 1, l_预交款(I).是否转账,
                           l_预交款(I).卡类别id, l_预交款(I).原交易流水号, l_预交款(I).原交易说明);
          Zl_病人结帐结算_Modify(1, n_病人id, n_结帐id, l_预交款(I).结算方式 || '|' || -1 * l_预交款(I).冲预交 || '| | ', Null, Null,
                           l_预交款(I).卡类别id, l_预交款(I).卡号, l_预交款(I).原交易流水号, l_预交款(I).原交易说明, Null, Null, Null, n_结算类型, Null,
                           v_操作员编码, v_操作员姓名, d_结帐时间, Null, 0, 1, Null, l_预交款(I).关联交易id, 0, Nvl(l_预交款(I).退款方式, 0) + 1);
        
          n_退预交金额 := n_退预交金额 + Nvl(l_预交款(I).冲预交, 0);
        End If;
      End Loop;
    
      --  2.有退必有冲；如单据A001，预交款金额1000，本次结帐800，则传入两条记录①冲1000，②退200，Sum(冲金额-退金额)=结帐金额
      --  说明：允许有误差金额
      If Abs(Nvl(n_结帐总额, 0) - (Nvl(n_冲预交金额, 0) - Nvl(n_退预交金额, 0))) >= 1.00 Then
        v_Err_Msg := '预交款支付金额与结帐金额不相等，不允许结算！';
        Raise Err_Item;
      End If;
    End If;
  End If;

  --【3】异步模式开始交易已处理完，返回结果
  If Nvl(n_结算模式, 0) = 1 And Nvl(n_操作类型, 0) = 0 Then
    v_Temp := '<CZSJ>' || To_Char(d_结帐时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<JZID>' || n_结帐id || '</JZID>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<KPBZ>' || 0 || '</KPBZ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<URL>' || '' || '</URL>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<NETURL>' || '' || '</NETURL>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<FPTT>' || v_姓名 || '</FPTT>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    v_Temp := '<FPH>' || '' || '</FPH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<FPJE>' || '' || '</FPJE>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<KPRQ>' || '' || '</KPRQ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    Xml_Out := x_Templet;
    Return;
  End If;

  --【4】非异步模式或异步模式完成交易，结算数据修正
  n_结算金额     := 0;
  n_删除原结算   := 1;
  n_结帐金额     := 0;
  n_非预交款结算 := 0;
  n_存在退预交   := 0;
  l_预交款       := t_预交款();
  For r_结算方式 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                        Extractvalue(b.Column_Value, '/JS/JSJE') As 结算金额,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                        Extractvalue(b.Column_Value, '/JS/JYSM') As 交易说明, Extractvalue(b.Column_Value, '/JS/ZY') As 摘要,
                        Extractvalue(b.Column_Value, '/JS/SFXFK') As 是否消费卡,
                        Extractvalue(b.Column_Value, '/JS/SFCYJ') As 是否冲预交,
                        Extract(b.Column_Value, '/JS/CYJLIST') As 冲预交列表,
                        Extract(b.Column_Value, '/JS/EXPENDLIST') As Expend
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
  
    v_卡类别   := r_结算方式.结算方式;
    n_结帐金额 := n_结帐金额 + Nvl(r_结算方式.结算金额, 0);
    If Nvl(r_结算方式.是否冲预交, 0) = 0 Then
      n_非预交款结算 := 1;
      n_卡类别id     := Null;
      If r_结算方式.结算卡类别 Is Not Null Then
        Select Decode(Translate(Nvl(r_结算方式.结算卡类别, 'abcd'), '#1234567890', '#'), Null, 1, 0)
        Into n_Number
        From Dual;
      
        If Nvl(r_结算方式.是否消费卡, 0) = 1 Then
          If Nvl(n_Number, 0) = 1 Then
            Select Max(编号), Max(名称)
            Into n_结算卡序号, v_卡类别
            From 消费卡类别目录
            Where 编号 = r_结算方式.结算卡类别 And Nvl(启用, 0) = 1;
          Else
            Select Max(编号), Max(名称)
            Into n_结算卡序号, v_卡类别
            From 消费卡类别目录
            Where 名称 = r_结算方式.结算卡类别 And Nvl(启用, 0) = 1;
          End If;
        
          If n_结算卡序号 Is Null Then
            v_Err_Msg := '未找到对应的消费卡信息';
            Raise Err_Item;
          End If;
        Else
          If Nvl(n_Number, 0) = 1 Then
            Select Max(ID), Max(名称)
            Into n_卡类别id, v_卡类别
            From 医疗卡类别
            Where ID = r_结算方式.结算卡类别 And Nvl(是否启用, 0) = 1;
          Else
            Select Max(ID), Max(名称)
            Into n_卡类别id, v_卡类别
            From 医疗卡类别
            Where 名称 = r_结算方式.结算卡类别 And Nvl(是否启用, 0) = 1;
          End If;
        
          If n_卡类别id Is Null Then
            v_Err_Msg := '未找到对应的医疗卡信息!';
            Raise Err_Item;
          End If;
        End If;
      End If;
    
      If n_卡类别id Is Not Null Then
        --三方卡
        Select Max(关联交易id)
        Into n_关联交易id
        From 病人预交记录
        Where 结帐id = n_结帐id And 卡类别id = n_卡类别id And Rownum < 2;
        If n_删除原结算 = 1 Then
          n_预交id := n_关联交易id;
        Else
          n_预交id := Null;
        End If;
      
        v_结算方式 := r_结算方式.结算方式 || '|' || r_结算方式.结算金额 || '|';
        Zl_病人结帐结算_Modify(1, n_病人id, n_结帐id, v_结算方式, Null, 0, n_卡类别id, r_结算方式.结算卡号, r_结算方式.交易流水号, r_结算方式.交易说明, 0, 0, 0,
                         n_结算类型, Null, v_操作员编码, v_操作员姓名, d_结帐时间, Null, 0, 1, n_预交id, n_关联交易id, n_删除原结算);
        n_删除原结算 := 0;
      
        三方结算交易_Save(n_结帐id, n_卡类别id, r_结算方式.结算卡号, r_结算方式.Expend);
      Else
        If n_结算卡序号 Is Not Null Then
          --消费卡
          v_消费卡结算 := Nvl(v_消费卡结算, '') || '||' || n_结算卡序号 || '|' || r_结算方式.结算卡号 || '|0|' || r_结算方式.结算金额;
        Else
          --其他结算
          v_结算方式 := r_结算方式.结算方式 || '|' || r_结算方式.结算金额 || '||';
          Zl_病人结帐结算_Modify(0, n_病人id, n_结帐id, v_结算方式, Null, 0, n_卡类别id, r_结算方式.结算卡号, r_结算方式.交易流水号, r_结算方式.交易说明, 0, 0, 0,
                           n_结算类型, Null, v_操作员编码, v_操作员姓名, d_结帐时间, Null, 0);
        End If;
      End If;
    
      n_结算金额 := n_结算金额 + Nvl(r_结算方式.结算金额, 0);
    Elsif r_结算方式.冲预交列表 Is Null Then
      --【**】冲预交,目前默认全冲
      n_冲预交金额 := r_结算方式.结算金额;
      Zl_病人结帐结算_Modify(0, n_病人id, n_结帐id, Null, n_冲预交金额, 0, Null, Null, Null, Null, 0, 0, 0, n_结算类型, Null, v_操作员编码,
                       v_操作员姓名, d_结帐时间, Null, 0);
    
      n_结算金额 := n_结算金额 + Nvl(r_结算方式.结算金额, 0);
    Else
      --【*】按指定单据进行预交款支付
      病人预交款_List(r_结算方式.冲预交列表, 0, l_预交款, n_存在退预交);
    End If;
  
    Update 三方交易记录
    Set 业务结算id = n_结帐id
    Where 流水号 = Nvl(r_结算方式.交易流水号, '-') And 类别 = v_卡类别 And 业务类型 = 2;
  End Loop;

  --消费卡处理
  If v_消费卡结算 Is Not Null Then
    v_消费卡结算 := Substr(v_消费卡结算, 3);
    Zl_病人结帐结算_Modify(3, n_病人id, n_结帐id, v_消费卡结算, Null, 0, Null, Null, Null, Null, 0, 0, 0, n_结算类型, Null, v_操作员编码,
                     v_操作员姓名, d_结帐时间, Null, 0);
  End If;

  --【*】按指定单据进行预交款支付
  If l_预交款.Count > 0 Then
    n_冲预交金额 := 0;
    n_退预交金额 := 0;
    If Nvl(n_存在退预交, 0) = 0 Then
      --新增预交款结算数据
      For I In 1 .. l_预交款.Count Loop
        Zl_结帐预交记录_Insert(l_预交款(I).预交id, l_预交款(I).单据号, 1, l_预交款(I).冲预交, n_结帐id, n_病人id, v_操作员编码, v_操作员姓名, d_结帐时间);
      
        n_冲预交金额 := n_冲预交金额 + Nvl(l_预交款(I).冲预交, 0);
      End Loop;
    Else
      --修正预交款结算数据
      --  1.如果存在退预交款，则只能使用预交款进行结帐
      If Nvl(n_非预交款结算, 0) = 1 Then
        v_Err_Msg := '存在对预交款进行退款时，所有全部结帐金额都必须使用预交款进行支付！';
        Raise Err_Item;
      End If;
    
      For I In 1 .. l_预交款.Count Loop
        If Nvl(l_预交款(I).是否退款, 0) = 0 Then
          n_冲预交金额 := n_冲预交金额 + Nvl(l_预交款(I).冲预交, 0);
        Else
          Zl_三方退款信息_Insert(n_结帐id, l_预交款(I).预交id, l_预交款(I).冲预交, l_预交款(I).卡号, l_预交款(I).交易流水号, l_预交款(I).交易说明, 1, 0,
                           l_预交款(I).是否转账, l_预交款(I).卡类别id, l_预交款(I).原交易流水号, l_预交款(I).原交易说明);
          Zl_病人结帐结算_Modify(1, n_病人id, n_结帐id, l_预交款(I).结算方式 || '|' || -1 * l_预交款(I).冲预交 || '| | ', Null, Null,
                           l_预交款(I).卡类别id, l_预交款(I).卡号, l_预交款(I).原交易流水号, l_预交款(I).原交易说明, Null, Null, Null, n_结算类型, Null,
                           v_操作员编码, v_操作员姓名, d_结帐时间, Null, 0, 2, Null, l_预交款(I).关联交易id, 1, Nvl(l_预交款(I).退款方式, 0) + 1);
        
          --保存扩展结算信息
          If Nvl(l_预交款(I).退款方式, 0) = 0 Then
            三方结算交易_Save(n_结帐id, l_预交款(I).卡类别id, l_预交款(I).卡号, l_预交款(I).交易扩展信息);
          End If;
        
          n_退预交金额 := n_退预交金额 + Nvl(l_预交款(I).冲预交, 0);
        End If;
      End Loop;
    
      --  2.有退必有冲；如单据A001，预交款金额1000，本次结帐800，则传入两条记录①冲1000，②退200，Sum(冲金额-退金额)=结帐金额
      If Abs(Nvl(n_结帐总额, 0) - (Nvl(n_冲预交金额, 0) - Nvl(n_退预交金额, 0))) > 1.00 Then
        v_Err_Msg := '预交款支付金额与结帐金额不相等，不允许结算！';
        Raise Err_Item;
      End If;
    End If;
  
    If Nvl(l_预交款(1).退款方式, 0) <> 0 Then
      For I In 1 .. l_预交款.Count Loop
        If Nvl(l_预交款(I).是否退款, 0) = 1 Then
          三方结算交易_Save(n_结帐id, l_预交款(I).卡类别id, l_预交款(I).卡号, l_预交款(I).扩展信息);
          Exit;
        End If;
      End Loop;
    End If;
  
    n_结算金额 := n_结算金额 + (Nvl(n_冲预交金额, 0) - Nvl(n_退预交金额, 0));
  End If;

  n_误差费 := Round(Nvl(n_结帐总额, 0) - Nvl(n_结算金额, 0), 6);
  If Abs(Nvl(n_误差费, 0)) > 1 Then
    v_Err_Msg := '计算的误差金额大于了1.00或小于-1.00元,不允许结帐操作,请检查!';
    Raise Err_Item;
  End If;

  --【5】完成结算，返回结果

  --电子票据处理  
  n_是否电子票据 := b_Einvoice_Request.Einvoice_Start(3, Null, n_结算类型);
  Zl_病人结帐结算_Modify(0, n_病人id, n_结帐id, '', Null, 0, Null, Null, Null, Null, 0, 0, n_误差费, n_结算类型, Null, v_操作员编码, v_操作员姓名,
                   d_结帐时间, Null, 1, 2, Null, Null, 0, Null, 0, n_是否电子票据);
  --需要开具电子票据
  If Nvl(n_是否电子票据, 0) = 1 Then
    If b_Einvoice_Request.Einvoice_Create(3, n_结帐id, Null, v_Err_Msg) = 0 Then
      --电子票据开具成功
      Raise Err_Item;
    End If;
  
    Select Max(1), Max(姓名), Max(号码), Max(To_Char(生成时间, 'yyyy-mm-dd')), Max(Url内网), Max(Url外网), Max(票据金额)
    Into n_开票标志, v_患者姓名, v_发票编号, v_开票日期, v_Url, v_Url外网, n_发票金额
    From 电子票据使用记录
    Where 结算id = n_结帐id And 票种 = 3 And 记录状态 = 1;
  
    If v_患者姓名 Is Not Null Then
      v_姓名 := v_患者姓名;
    End If;
  End If;
  v_Temp := '<CZSJ>' || To_Char(d_结帐时间, 'YYYY-MM-DD hh23:mi:ss') || '</CZSJ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<JZID>' || n_结帐id || '</JZID>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<KPBZ>' || Nvl(n_开票标志, 0) || '</KPBZ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<URL>' || Nvl(v_Url, '') || '</URL>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<NETURL>' || Nvl(v_Url外网, '') || '</NETURL>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<FPTT>' || v_姓名 || '</FPTT>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<FPH>' || v_发票编号 || '</FPH>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<FPJE>' || Nvl(n_发票金额, 0) || '</FPJE>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<KPRQ>' || v_开票日期 || '</KPRQ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Err_Special Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20105, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Settlement;
/

Create Or Replace Procedure Zl_电子票据类别_Update(编号_In 电子票据类别.编号%Type) As
  --功能：修改电子票据类别的启用、停用
  v_Error Varchar2(255);
  Err_Item Exception;
Begin

  If 编号_In Is Null Then
    v_Error := '传入电子票据类别的编号,请检查！';
    Raise Err_Item;
  End If;

  --先停用原电子票据接口
  Update 电子票据类别 Set 是否启用 = Null Where Nvl(是否启用, 0) = 1;

  --再启用现电子票据接口
  Update 电子票据类别 Set 是否启用 = 1 Where 编号 = 编号_In;
  If Sql%NotFound Then
    v_Error := '传入编号未找到对应的电子票据类别数据,请检查！';
    Raise Err_Item;
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_电子票据类别_Update;
/
Create Or Replace Procedure Zl_合约单位_Insert
(
  Id_In           In 合约单位.Id%Type,
  上级id_In       In 合约单位.上级id%Type,
  编码_In         In 合约单位.编码%Type,
  名称_In         In 合约单位.名称%Type,
  简码_In         In 合约单位.简码%Type := Null,
  地址_In         In 合约单位.地址%Type := Null,
  电话_In         In 合约单位.电话%Type := Null,
  开户银行_In     In 合约单位.开户银行%Type := Null,
  帐号_In         In 合约单位.帐号%Type := Null,
  联系人_In       In 合约单位.联系人%Type := Null,
  末级_In         In 合约单位.末级%Type := 1,
  电子邮件_In     In 合约单位.电子邮件%Type := Null,
  说明_In         In 合约单位.说明%Type := Null,
  站点_In         In 合约单位.站点%Type := Null,
  社会信用代码_In In 合约单位.社会信用代码%Type := Null
) Is
Begin
  --首先插入记录 
  Insert Into 合约单位
    (ID, 编码, 名称, 简码, 地址, 电话, 开户银行, 帐号, 联系人, 上级id, 建档时间, 撤档时间, 末级, 电子邮件, 说明, 站点, 社会信用代码)
  Values
    (Id_In, 编码_In, 名称_In, 简码_In, 地址_In, 电话_In, 开户银行_In, 帐号_In, 联系人_In, 上级id_In, Sysdate,
     To_Date('3000-01-01', 'yyyy-mm-dd'), 末级_In, 电子邮件_In, 说明_In, 站点_In, 社会信用代码_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_合约单位_Insert;
/

Create Or Replace Procedure Zl_合约单位_Update
(
  Id_In           In 合约单位.Id%Type,
  上级id_In       In 合约单位.上级id%Type,
  编码_In         In 合约单位.编码%Type,
  名称_In         In 合约单位.名称%Type,
  简码_In         In 合约单位.简码%Type,
  地址_In         In 合约单位.地址%Type := Null,
  电话_In         In 合约单位.电话%Type := Null,
  开户银行_In     In 合约单位.开户银行%Type := Null,
  帐号_In         In 合约单位.帐号%Type := Null,
  联系人_In       In 合约单位.联系人%Type := Null,
  原长度_In       In Number,
  电子邮件_In     In 合约单位.电子邮件%Type := Null,
  说明_In         In 合约单位.说明%Type := Null,
  站点_In         In 合约单位.站点%Type := Null,
  社会信用代码_In In 合约单位.社会信用代码%Type := Null
) Is
Begin
  --首先插入修改记录 
  Update 合约单位
  Set 编码 = 编码_In, 名称 = 名称_In, 简码 = 简码_In, 地址 = 地址_In, 电话 = 电话_In, 开户银行 = 开户银行_In, 帐号 = 帐号_In, 联系人 = 联系人_In,
      上级id = 上级id_In, 电子邮件 = 电子邮件_In, 说明 = 说明_In, 站点 = 站点_In, 社会信用代码 = 社会信用代码_In
  Where ID = Id_In;

  --对它的下级也要修改编码 
  Update 合约单位
  Set 编码 = 编码_In || Substr(编码, 原长度_In)
  Where ID In (Select ID From 合约单位 Start With 上级id = Id_In Connect By Prior ID = 上级id);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_合约单位_Update;
/
Create Or Replace Procedure Zl_病人预交记录_Delete
(
  Id_In         病人预交记录.Id%Type,
  摘要_In       病人预交记录.摘要%Type,
  操作员编号_In 病人预交记录.操作员编号%Type,
  操作员姓名_In 病人预交记录.操作员姓名%Type,
  帐户退费_In   Number := 1,
  冲预交id_In   病人预交记录.Id%Type := Null,
  票据号_In     病人预交记录.实际票号%Type := Null,
  领用id_In     票据领用记录.Id%Type := Null,
  校对标志_In   病人预交记录.校对标志%Type := Null,
  结算模式_In   Number := 0,
  结算状态_In   Number := 0,
  三方退现_In   Number := 0,
  退现方式_In   病人押金记录.结算方式%Type := Null
) As
  --校对标志_In  三方卡退款时，先传入1，生成校对标志为1的记录，再传入空或0更新为0的正常退款记录。
  --结算模式_In  0-同步完成，1-异步完成
  --结算状态_In  结算模式_In=1时，0-异常状态，1-完成结算
  --三方退现_In:三方支付的预交款是否退现  0：不退现，1：退现
  Cursor c_Moneyinfo Is
    Select ID, NO, 金额, 结算方式, 病人id, 预交类别
    From 病人预交记录
    Where ID = Id_In And 记录性质 = 1 And (记录状态 = 1 Or 记录状态 = 3);
  r_Moneyrow c_Moneyinfo%RowType;

  v_打印id   票据打印内容.Id%Type;
  v_性质     结算方式.性质%Type;
  v_现金     结算方式.名称%Type;
  v_个人帐户 结算方式.名称%Type;
  n_返回值   病人余额.预交余额%Type;
  n_预交id   病人预交记录.Id%Type;
  v_No       病人预交记录.No%Type;
  n_卡类别id 病人预交记录.卡类别id%Type;
  v_Date     Date;
  Err_Custom Exception;
  n_组id         财务缴款分组.Id%Type;
  v_Msg          Varchar2(500);
  n_预交电子票据 Number(1);
Begin
  n_预交id := 冲预交id_In;

  --获取结算方式名称
  Select Max(Nvl(名称, '现金')) Into v_现金 From 结算方式 Where 性质 = 1;
  Select Max(Nvl(名称, '个人帐户')) Into v_个人帐户 From 结算方式 Where 性质 = 3;

  Open c_Moneyinfo;
  Fetch c_Moneyinfo
    Into r_Moneyrow;

  --首先判断要退款的记录是否存在
  If c_Moneyinfo%RowCount = 0 Then
    Close c_Moneyinfo;
    Raise Err_Custom;
  End If;
  Select Sysdate Into v_Date From Dual;
  If n_预交id Is Null Then
    Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
  End If;
  n_组id := Zl_Get组id(操作员姓名_In);

  If Not (结算模式_In = 1 And 结算状态_In = 1) Then
  
    If Nvl(帐户退费_In, 0) = 1 Then
      --支持个人帐户退费,正常处理
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 金额, 结算方式, 结算号码, 收款时间, 操作员编号, 操作员姓名, 缴款单位, 单位开户行, 单位帐号, 缴款组id,
         预交类别, 卡类别id, 卡号, 交易流水号, 交易说明, 合作单位, 结算卡序号, 校对标志, 关联交易id, 交易时间, 交易人员, 预交电子票据)
        Select n_预交id, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要_In, -1 * 金额, Decode(三方退现_In, 1, 退现方式_In, 结算方式),
               Decode(三方退现_In, 1, Null, 结算号码), v_Date, 操作员编号_In, 操作员姓名_In, 缴款单位, 单位开户行, 单位帐号, n_组id, 预交类别,
               Decode(三方退现_In, 1, Null, 卡类别id), Decode(三方退现_In, 1, Null, 卡号), Decode(三方退现_In, 1, Null, 交易流水号),
               Decode(三方退现_In, 1, Null, 交易说明), 合作单位, 结算卡序号, 校对标志_In, 关联交易id, Decode(三方退现_In, 1, Null, v_Date),
               Decode(三方退现_In, 1, Null, 操作员姓名_In), 预交电子票据
        From 病人预交记录
        Where ID = Id_In;
    Else
      --不支持时,处理成现金,记录性质为2的摘要填标志,为3的更新新输入的摘要
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 金额, 结算方式, 结算号码, 收款时间, 操作员编号, 操作员姓名, 缴款单位, 单位开户行, 单位帐号, 缴款组id,
         预交类别, 卡类别id, 卡号, 交易流水号, 交易说明, 合作单位, 结算卡序号, 校对标志, 关联交易id, 交易时间, 交易人员, 预交电子票据)
        Select n_预交id, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, Nvl(摘要_In, '个人帐户退款'), -1 * 金额,
               Decode(结算方式, v_个人帐户, v_现金, Decode(三方退现_In, 1, 退现方式_In, 结算方式)), Decode(三方退现_In, 1, Null, 结算号码), v_Date,
               操作员编号_In, 操作员姓名_In, Decode(结算方式, v_个人帐户, Null, 缴款单位), Decode(结算方式, v_个人帐户, Null, 单位开户行),
               Decode(结算方式, v_个人帐户, Null, 单位帐号), n_组id, 预交类别, Decode(三方退现_In, 1, Null, 卡类别id),
               Decode(三方退现_In, 1, Null, 卡号), Decode(三方退现_In, 1, Null, 交易流水号), Decode(三方退现_In, 1, Null, 交易说明), 合作单位,
               结算卡序号, 校对标志_In, 关联交易id, Decode(三方退现_In, 1, Null, v_Date), Decode(三方退现_In, 1, Null, 操作员姓名_In), 预交电子票据
        From 病人预交记录
        Where ID = Id_In;
    End If;
    Select 卡类别id Into n_卡类别id From 病人预交记录 Where ID = Id_In;
    If Nvl(n_卡类别id, 0) <> 0 Then
      --自定义过程调用
      Zl_Custom_Balance_Update(n_预交id);
    End If;
    Update 病人预交记录 Set 记录状态 = 3 Where ID = Id_In;
    --病人(预交)余额(不管是退现金还是个人帐户都应该减少)
    --判断要退款的性质
    Select b.性质 Into v_性质 From 病人预交记录 A, 结算方式 B Where a.结算方式 = b.名称(+) And a.Id = Id_In;
    If Nvl(v_性质, 1) <> 5 Then
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) - r_Moneyrow.金额
      Where 性质 = 1 And 病人id = r_Moneyrow.病人id And Nvl(类型, 2) = Nvl(r_Moneyrow.预交类别, 2)
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 性质, 类型, 预交余额, 费用余额)
        Values
          (r_Moneyrow.病人id, 1, Nvl(r_Moneyrow.预交类别, 2), -r_Moneyrow.金额, 0);
        n_返回值 := -r_Moneyrow.金额;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 病人余额
        Where 病人id = r_Moneyrow.病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
      End If;
    End If;
  
    --预交单据余额
    Update 预交单据余额
    Set 预交余额 = Nvl(预交余额, 0) - r_Moneyrow.金额
    Where 预交id = Id_In
    Returning 预交余额 Into n_返回值;
    If Sql%RowCount = 0 Then
      Insert Into 预交单据余额
        (预交id, 病人id, 预交类别, 预交余额)
      Values
        (r_Moneyrow.Id, r_Moneyrow.病人id, Nvl(r_Moneyrow.预交类别, 2), -r_Moneyrow.金额);
      n_返回值 := -r_Moneyrow.金额;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 预交单据余额
      Where 预交id = r_Moneyrow.Id And Nvl(预交类别, 2) = Nvl(r_Moneyrow.预交类别, 2) And Nvl(预交余额, 0) = 0;
    End If;
  End If;

  --异步操作完成时才执行此内容
  If 结算模式_In = 1 Then
    If 结算状态_In = 1 Then
      --对三方卡退款，先生成校对标志为0的记录更新
      Update 病人预交记录
      Set 校对标志 = 校对标志_In, 收款时间 = v_Date, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In, 缴款组id = n_组id, 交易时间 = v_Date,
          交易人员 = 操作员姓名_In
      Where ID = n_预交id;
      --自定义过程调用
      Zl_Custom_Balance_Update(n_预交id);
    Else
      Return;
    End If;
  End If;

  --处理相关汇总表
  --人员缴款余额(注意包括处理个人帐户的结算方式)
  If Nvl(帐户退费_In, 0) = 1 Then
    --支持退个人帐户时的处理
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) - r_Moneyrow.金额
    Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = Decode(三方退现_In, 1, 退现方式_In, r_Moneyrow.结算方式)
    Returning 余额 Into n_返回值;
  
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额
        (收款员, 结算方式, 性质, 余额)
      Values
        (操作员姓名_In, Decode(三方退现_In, 1, 退现方式_In, r_Moneyrow.结算方式), 1, -r_Moneyrow.金额);
      n_返回值 := -r_Moneyrow.金额;
    
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 人员缴款余额
      Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = Decode(三方退现_In, 1, 退现方式_In, r_Moneyrow.结算方式) And Nvl(余额, 0) = 0;
    End If;
  Else
    --不支持时的处理
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) - r_Moneyrow.金额
    Where 性质 = 1 And 收款员 = 操作员姓名_In And
          结算方式 = Decode(r_Moneyrow.结算方式, v_个人帐户, v_现金, Decode(三方退现_In, 1, 退现方式_In, r_Moneyrow.结算方式))
    Returning 余额 Into n_返回值;
  
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额
        (收款员, 结算方式, 性质, 余额)
      Values
        (操作员姓名_In, Decode(r_Moneyrow.结算方式, v_个人帐户, v_现金, Decode(三方退现_In, 1, 退现方式_In, r_Moneyrow.结算方式)), 1,
         -r_Moneyrow.金额);
      n_返回值 := -r_Moneyrow.金额;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 人员缴款余额
      Where 收款员 = 操作员姓名_In And 性质 = 1 And
            结算方式 = Decode(r_Moneyrow.结算方式, v_个人帐户, v_现金, Decode(三方退现_In, 1, 退现方式_In, r_Moneyrow.结算方式)) And
            Nvl(余额, 0) = 0;
    End If;
  End If;
  --作废收回票据(可能以前没有使用票据,无法收回)
  Begin
    Select Nvl(预交电子票据, 0) Into n_预交电子票据 From 病人预交记录 Where ID = Id_In;
    If n_预交电子票据 = 0 Then
      Select ID
      Into v_打印id
      From (Select b.Id
             From 票据使用明细 A, 票据打印内容 B
             Where a.打印id = b.Id And a.性质 = 1 And b.数据性质 = 2 And b.No = r_Moneyrow.No
             Order By a.使用时间 Desc)
      Where Rownum < 2;
    End If;
  Exception
    When Others Then
      Null;
  End;

  If v_打印id Is Not Null Then
    Insert Into 票据使用明细
      (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
      Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, v_Date, 操作员姓名_In, 票据金额
      From 票据使用明细
      Where 打印id = v_打印id And 票种 = 2 And 性质 = 1;
  End If;

  --处理票据
  If 票据号_In Is Not Null Then
    Select 票据打印内容_Id.Nextval Into v_打印id From Dual;
  
    --发出票据
    Insert Into 票据打印内容 (ID, 数据性质, NO) Values (v_打印id, 2, r_Moneyrow.No);
  
    Insert Into 票据使用明细
      (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
    Values
      (票据使用明细_Id.Nextval, 2, 票据号_In, 1, 6, 领用id_In, v_打印id, v_Date, 操作员姓名_In, -1 * r_Moneyrow.金额);
  
    --状态改动
    Update 票据领用记录
    Set 当前号码 = 票据号_In, 剩余数量 = Decode(Sign(剩余数量 - 1), -1, 0, 剩余数量 - 1), 使用时间 = Sysdate
    Where ID = Nvl(领用id_In, 0);
  
  End If;

  Close c_Moneyinfo;

  --消息推送;
  Select NO Into v_No From 病人预交记录 Where ID = n_预交id;
  b_Message.Zlhis_Charge_006(n_预交id, v_No);
  Select Id_In || ',' || 帐户退费_In Into v_Msg From Dual;
  Begin
    Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
      Using 12, v_Msg;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20999, '[ZLSOFT]没有发现要退款的预交记录,该记录可能已经退除！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人预交记录_Delete;
/

Create Or Replace Procedure Zl1_Datamove_Tag
(
  d_End            In Date,
  n_批次           In Number,
  n_System         In Number,
  n_预交剩余款上限 In 病人预交记录.金额%Type := 10 --当病人不存在未结费用，也不是在院病人时，允许未冲完的预交款在指定值以下的数据强制转出，避免大量呆帐未转出从而影响转出速度
  
) As
  --功能：标记待转出的数据
  --说明：为避免Undo表空间膨胀过大，分段提交
  d_Lastend Date; --最终转出截止时间（d_End为本批转出截止时间）

  --递归取消“一张预交款单据中的一部分被标记为待转出”的数据
  Procedure Datamove_Tag_Update
  (
    结帐id_In t_NumList,
    d_End     In Date,
    n_批次    In Number
  ) As
  
    c_结帐id t_NumList := t_NumList();
    c_No     t_StrList := t_StrList();
  Begin
    --1.1一张预交单据被多个结帐ID冲了，找出其中的一部分被标记为待转出的数据，如：
    --   NO=A001 记录性质=11 结帐ID=10 待转出=1
    --   NO=A001 记录性质=11 结帐ID=11 待转出=NULL
    If 结帐id_In Is Null Then
      Select Distinct a.No Bulk Collect
      Into c_No
      From 病人预交记录 A
      Where a.记录性质 In (1, 11) And a.待转出 = n_批次 And Exists
       (Select 1 From 病人预交记录 Where NO = a.No And 记录性质 In (1, 11) And 待转出 Is Null);
    Else
      Select Distinct a.No Bulk Collect
      Into c_No
      From 病人预交记录 A
      Where a.结帐id In (Select /*+cardinality(b,10) */
                        Column_Value
                       From Table(结帐id_In) B) And a.记录性质 In (1, 11) And a.待转出 Is Null And Exists
       (Select 1 From 病人预交记录 Where NO = a.No And 记录性质 In (1, 11) And 待转出 + 0 = n_批次);
    End If;
  
    If c_No.Count = 0 Then
      Return;
    End If;
  
    --1.2取消标记
    Forall I In 1 .. c_No.Count
      Update 病人预交记录 Set 待转出 = Null Where NO = c_No(I) And 记录性质 In (1, 11);
  
    --------------------------------------------------------------------------------------------------------
    --2.1一个结帐ID冲了多张预交单据，找出其中的一部分被标记为待转出的数据，如：
    --   NO=A001 记录性质=11 结帐ID=20 待转出=1
    --   NO=A002 记录性质=11 结帐ID=20 待转出=NULL
    Select Distinct a.结帐id Bulk Collect
    Into c_结帐id
    From 病人预交记录 A
    Where a.No In (Select /*+cardinality(b,10) */
                    Column_Value
                   From Table(c_No) B) And a.记录性质 In (1, 11) And a.待转出 Is Null And a.收款时间 + 0 < d_End And Exists
     (Select 1 From 病人预交记录 Where 结帐id = a.结帐id And 待转出 + 0 = n_批次);
  
    If c_结帐id.Count = 0 Then
      Return;
    End If;
  
    --2.2取消标记(包括一次结帐的其他结算方式的记录)
    Forall I In 1 .. c_结帐id.Count
      Update 病人预交记录 Set 待转出 = Null Where 结帐id = c_结帐id(I);
  
    --递归调用
    Datamove_Tag_Update(c_结帐id, d_End, n_批次);
  End Datamove_Tag_Update;
Begin
  Select 本次最终日期 Into d_Lastend From zlDataMove Where 系统 = n_System And 组号 = 1;
  If d_Lastend Is Null Then
    Return;
  End If;
  --新加子查询注意性能优化，把能够将数据过滤到最小的条件放到最后，Exists类条件放前面

  --1.经济核算（费用,药品,收款和票据等）
  --冲销业务与原始业务的发生时间相同，登记时间不同，所以要按发生时间来查询.
  --以下情况，可能有多个结帐ID，或涉及多个费用单据，这些数据要一起转出或排除转出，否则影响后续判断是否结清
  --1.一张费用单据的一行费用或多行费用可能分多次结帐（有多个不同的结帐ID）
  --2.结帐作废后也可能分多次结清(一张单据多个不同的结帐ID)
  --3.结帐作废后可能与其他费用单据一起结(一张单据的多个结帐ID，涉及多个费用NO，这些NO可能之前结帐作废过，有其他结帐ID)
  --考虑到这情况的复杂性，为简化逻辑，提升查询性能，按病人ID来排除(该病人的结帐数据都不转出)

  Update /*+ rule*/ 病人预交记录 L
  Set 待转出 = n_批次
  Where 结帐id In
        (Select Distinct a.结帐id --1.门诊收费和挂号的收费结算记录
         From 门诊费用记录 A
         Where (a.记录状态 In (1, 2) Or a.记录状态 = 3 And Not Exists
                (Select 1
                 From 门诊费用记录 B
                 Where a.No = b.No And a.记录性质 = b.记录性质 And b.记录状态 = 2 And b.登记时间 >= d_Lastend)) And a.待转出 Is Null And
               a.记录性质 In (1, 4) And a.发生时间 < d_End And a.登记时间 < d_Lastend
         Union All
         Select Distinct b.结算id --2.医保补结算(没有发生时间字段,作废记录的登记时间不同，为了把收费和作废的一次性转出，所以要连接B表)
         From 费用补充记录 A, 费用补充记录 B
         Where a.待转出 Is Null And a.No = b.No And a.记录性质 = b.记录性质 And a.登记时间 < d_End
         Union All
         Select Distinct a.结帐id --3.就诊卡的收费结算记录(排除之后退卡费的,一张单据中只要其中一行退了)
         From 住院费用记录 A
         Where (a.记录状态 In (1, 2) Or a.记录状态 = 3 And Not Exists
                (Select 1
                 From 住院费用记录 B
                 Where a.No = b.No And a.记录性质 = b.记录性质 And b.记录状态 = 2 And b.登记时间 >= d_Lastend)) And a.待转出 Is Null And
               a.记帐费用 = 0 And a.记录性质 = 5 And a.发生时间 < d_End
         Union All --4.住院记帐费用的结帐结算记录
         Select 结帐id
         From (With Settle As (Select Distinct c.结帐id
                               From (Select Distinct b.No, b.序号, Mod(b.记录性质, 10) As 记录性质
                                      From (Select Distinct b.Id
                                             From 病人结帐记录 A, 病人结帐记录 B --作废的结帐单的收费时间可能在指定时间之后，所以要连接B表
                                             Where (a.记录状态 In (1, 2) Or a.记录状态 = 3 And Not Exists
                                                    (Select 1
                                                     From 病人结帐记录 C
                                                     Where a.No = c.No And c.记录状态 = 2 And c.收费时间 >= d_Lastend)) And
                                                   a.待转出 Is Null And a.No = b.No And (a.结帐类型 = 2 Or Nvl(a.结帐类型, 0) = 0) And
                                                   a.收费时间 < d_End) A, 住院费用记录 B
                                      Where a.Id = b.结帐id) B, 住院费用记录 C --通过C表找到这些费用单据的所有结帐ID一起转(可能在转出时间之后)
                               Where c.No = b.No And Mod(c.记录性质, 10) = b.记录性质 And c.序号 = b.序号)
                Select 结帐id
                From Settle
                Minus
                Select Distinct a.Id
                From 病人结帐记录 A,
                     (Select Distinct 病人id
                       From (Select c.病人id, c.No, Mod(c.记录性质, 10) As 记录性质, Nvl(Sum(c.实收金额), 0) As 实收金额,
                                     Nvl(Sum(c.结帐金额), 0) As 结帐金额
                              From 住院费用记录 C, Settle S
                              Where c.结帐id = s.结帐id
                              Group By c.No, Mod(c.记录性质, 10), c.病人id) C
                       Where c.实收金额 <> c.结帐金额 And Exists (Select 1 From 在院病人 F Where c.病人id = f.病人id) --出院病人没有结清的也转走（在需要时再抽回），否则排除的数据量太大
                             Or Exists (Select 1
                              From 住院费用记录 E, 病人结帐记录 S
                              Where e.No = c.No And Mod(e.记录性质, 10) = c.记录性质 And e.结帐id = s.Id And
                                    s.待转出 Is Null And s.收费时间 >= d_Lastend)) N --即使是在本批转出时间之后结清，只要不是在最终转出时间之后，就不排除
                
                Where a.病人id = n.病人id And (a.结帐类型 = 2 Or Nvl(a.结帐类型, 0) = 0))
                Union All --5.门诊记帐费用的结帐结算记录
                Select 结帐id
                From (With Settle As (Select Distinct c.结帐id
                                      From (Select Distinct b.No, b.序号, Mod(b.记录性质, 10) As 记录性质
                                             From (Select Distinct b.Id
                                                    From 病人结帐记录 A, 病人结帐记录 B
                                                    Where a.待转出 Is Null And a.No = b.No And (a.结帐类型 = 1 Or Nvl(a.结帐类型, 0) = 0) And
                                                          a.收费时间 < d_End) A, 门诊费用记录 B
                                             Where a.Id = b.结帐id) B, 门诊费用记录 C
                                      Where c.No = b.No And Mod(c.记录性质, 10) = b.记录性质 And c.序号 = b.序号)
                       Select 结帐id
                       From Settle
                       Minus
                       Select Distinct a.Id
                       From 病人结帐记录 A,
                            (Select Distinct c.病人id
                              From (Select c.病人id, c.No, Mod(c.记录性质, 10) As 记录性质, Nvl(Sum(c.实收金额), 0) As 实收金额,
                                            Nvl(Sum(c.结帐金额), 0) As 结帐金额
                                     From 门诊费用记录 C, Settle S
                                     Where c.结帐id = s.结帐id
                                     Group By c.No, Mod(c.记录性质, 10), c.病人id) C
                              Where c.实收金额 <> c.结帐金额 --门诊病人没有结清的不转走
                                    Or Exists (Select 1
                                     From 门诊费用记录 E, 病人结帐记录 S
                                     Where e.No = c.No And Mod(e.记录性质, 10) = c.记录性质 And e.结帐id = s.Id And
                                           s.待转出 Is Null And s.收费时间 >= d_Lastend)) N
                       Where a.病人id = n.病人id And (a.结帐类型 = 1 Or Nvl(a.结帐类型, 0) = 0))
                       
         
         
         );

  --排除预交款未冲完的
  --为了降低逻辑的复杂性，不排除在转出时间之后发药或未发药的费用记录对应的结帐ID，将这种情况的结算数据和费用数据强制转走
  --因为前面的SQL查出的结帐ID可能不全是冲预交的(门诊收费和住院结帐补费等)，所以，需要单独一个SQL来排除
  Update /*+ rule*/ 病人预交记录
  Set 待转出 = Null
  Where 待转出 = n_批次 And
        结帐id In
        (Select Distinct d.结帐id --该单据相关的所有冲预交的结帐ID都不转出
         From 病人预交记录 D,
              (Select Distinct l.No
                From (Select l.No, l.病人id, l.预交类别, Nvl(Sum(l.金额), 0) As 金额, Nvl(Sum(l.冲预交), 0) As 冲预交,
                              Sum(Decode(l.待转出, Null, Decode(结帐id, Null, Decode(记录状态, 2, 0, 1), 1), 0)) As 未转出
                       From 病人预交记录 L --可能按结帐ID确认本次待转出的冲的只是剩余款，所以需要连接L表，查原始交预交的单据，以及记录性质为11的可能还有转出时间之后其他冲剩余款的结帐ID
                       Where l.记录性质 In (1, 11) And
                             l.No In
                             (Select Distinct p.No From 病人预交记录 P Where p.记录性质 In (1, 11) And p.待转出 = n_批次)
                       Group By l.No, l.病人id, l.预交类别) L --多次住院可以一次结清，所以，不能加主页ID
                Where 未转出 > 0 --只要该预交单据还有未转出的预交或冲预交记录，则不转出，避免转出一部分导致后续判断错误
                      Or
                      l.金额 <> l.冲预交 And
                      (Exists (Select 1
                               From 病人预交记录 E --剩的预交款，一般用负数交预交来退款（NO号不同），这种相当于是冲完了，不排除
                               Where e.病人id = l.病人id And e.预交类别 = l.预交类别 And e.记录性质 In (1, 11) And
                                     (e.待转出 = n_批次 Or e.待转出 Is Null And e.结帐id Is Null And e.记录性质 = 1 And 收款时间 < d_End)
                                Having Abs(Nvl(Sum(e.金额), 0) - Nvl(Sum(e.冲预交), 0)) > n_预交剩余款上限) --余额小于等于n不排除，与下面第3种结帐ID为空的要保持一致
                       Or l.预交类别 = 2 And Exists (Select 1 From 在院病人 E Where l.病人id = e.病人id) Or Exists
                       (Select 1
                        From 病人未结费用 E
                        Where l.病人id = e.病人id And (l.预交类别 = 1 And e.主页id Is Null Or l.预交类别 = 2 And e.主页id Is Not Null)))) N
         Where d.No = n.No And d.记录性质 In (1, 11));

  --单独处理3种结帐ID为空的预交记录
  --1.预交款没有使用就直接退了的记录(结帐ID为空)
  Update /*+ rule*/ 病人预交记录
  Set 待转出 = n_批次
  Where 记录性质 = 1 And
        NO In (Select a.No
               From 病人预交记录 A
               Where a.结帐id Is Null And a.记录性质 = 1 And a.记录状态 In (2, 3) And a.待转出 Is Null And a.收款时间 < d_End
               Group By a.No
               Having Sum(a.金额) = 0);

  --2.交预交款后退款的记录（结帐ID为空，记录状态为2）
  Update /*+ rule*/ 病人预交记录
  Set 待转出 = n_批次
  Where 结帐id Is Null And 记录性质 = 1 And 记录状态 = 2 And
        NO In (Select a.No From 病人预交记录 A Where a.记录性质 = 1 And a.记录状态 = 3 And a.待转出 = n_批次);

  --排除同一张预交款单据部分记录被标记为转出的,只要有不转出的，则整张单据都不转出
  --跟第2种有关联影响，所以要放在它之后执行
  --要影响第3种情况的判断，所以要放在它之前执行
  Datamove_Tag_Update(Null, d_End, n_批次);

  --3.预交款未用完时用交负数预交来退款(结帐ID为空，并且跟原始的冲预交的NO没有关联关系)
  --不加条件"金额 < 0"，因为存在预交款没有使用过，就直接用交负数预交来退款的情况
  Update /*+ rule*/ 病人预交记录 L
  Set 待转出 = n_批次
  Where Exists (Select 1
         From 病人预交记录 E
         Where e.病人id = l.病人id And e.预交类别 = l.预交类别 And e.记录性质 In (1, 11) And
               (e.待转出 = n_批次 Or e.待转出 Is Null And e.结帐id Is Null And e.记录性质 = 1 And 记录状态 = 1 And 收款时间 < d_End)
         Group By e.病人id
         Having Abs(Nvl(Sum(e.金额), 0) - Nvl(Sum(e.冲预交), 0)) <= n_预交剩余款上限) --余额小于等于n要转出，与前面“排除预交款未冲完的”要保持一致
       
        And Exists (Select 1
         From 病人预交记录 E
         Where e.病人id = l.病人id And e.预交类别 = l.预交类别 And e.记录性质 In (1, 11) And e.待转出 = n_批次) And
        待转出 Is Null And 结帐id Is Null And 记录性质 = 1 And 记录状态 = 1 And 收款时间 < d_End;

  Update /*+ rule*/ 病人押金记录 Set 待转出 = n_批次 Where 记录状态 In (2, 3) And 待转出 Is Null And 收款时间 < d_End;

  Update /*+ rule*/ 三方结算交易
  Set 待转出 = n_批次
  Where 交易id In (Select a.Id From 病人押金记录 A Where 待转出 = n_批次 And Nvl(性质, 0) = 2);

  Update /*+ rule*/ 电子票据使用记录
  Set 待转出 = n_批次
  Where 票种 = 2 And
        结算id In
        (Select ID From 病人预交记录 Where 待转出 = n_批次 And Nvl(预交电子票据, 0) = 1 And Mod(记录性质, 10) = 1);

  Update /*+ rule*/ 电子票据使用记录
  Set 待转出 = n_批次
  Where 票种 <> 2 And
        结算id In
        (Select ID From 病人预交记录 Where 待转出 = n_批次 And Nvl(是否电子票据, 0) = 1 And Mod(记录性质, 10) <> 1);

  Update /*+ rule*/ 电子票据二维码
  Set 待转出 = n_批次
  Where 使用记录id In (Select ID From 电子票据使用记录 Where 待转出 = n_批次);

  --预交票据（不严格控制）
  Update /*+ rule*/ 票据使用明细
  Set 待转出 = n_批次
  Where 票种 = 2 And 号码 In (Select Distinct 实际票号
                          From 病人预交记录
                          Where Mod(记录性质, 0) = 1 And Nvl(校对标志, 0) = 0 And 待转出 = n_批次) And Nvl(领用id, 0) = 0;

  Update zlDataMovelog
  Set 当前进度 = '(1/11)结算数据标记完成，正在标记费用数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  Update /*+ rule*/ 病人结帐记录
  Set 待转出 = n_批次
  Where ID In (Select 结帐id From 病人预交记录 Where 待转出 = n_批次);

  --结帐无结算的记录(为了提升性能，不判断费用，只要结了帐且无预交记录就当成是零费用结帐)
  Update /*+ rule*/ 病人结帐记录 L
  Set 待转出 = n_批次
  Where 收费时间 < d_End And 待转出 Is Null And Not Exists (Select 1 From 病人预交记录 P Where l.Id = p.结帐id);

  --结帐票据（不严格控制）
  Update /*+ rule*/ 票据使用明细
  Set 待转出 = n_批次
  Where 票种 = 3 And 号码 In (Select Distinct 实际票号 From 病人结帐记录 Where Nvl(结算状态, 0) = 0 And 待转出 = n_批次) And Nvl(领用id, 0) = 0;

  Update /*+ rule*/ 病人卡结算记录
  Set 待转出 = n_批次
  Where 记录性质 = 4 And 结算id In (Select a.Id From 病人预交记录 A Where 待转出 = n_批次);

  Update /*+ rule*/ 三方结算交易
  Set 待转出 = n_批次
  Where 交易id In (Select a.Id From 病人预交记录 A Where 待转出 = n_批次) And Nvl(性质, 0) = 0;

  Update /*+ rule*/ 三方退款信息
  Set 待转出 = n_批次
  Where (记录id, 结帐id) In (Select a.Id, a.结帐id From 病人预交记录 A Where 待转出 = n_批次);

  --1.挂号费用异常数据
  --a.结帐ID为空（实收金额可能不为零）
  --b.结帐ID不为空，打折后实收金额为0（应收金额正负冲销）的挂号费用，没有挂号记录，也没有预交记录
  --按发生时间转出，因为收和退的发生时间相同，登记时间不同。
  Update /*+ rule*/ 门诊费用记录
  Set 待转出 = n_批次
  Where 待转出 Is Null And 发生时间 < d_End And 记录性质 = 4 And (实收金额 = 0 Or 结帐id Is Null);

  --2.直接收费的和结帐无结算（预交）记录的，Union不加all去掉重复以减少in的数量
  Update /*+ rule*/ 门诊费用记录
  Set 待转出 = n_批次
  Where 结帐id In (Select 结帐id
                 From 病人预交记录
                 Where 待转出 = n_批次
                 Union
                 Select ID From 病人结帐记录 Where 待转出 = n_批次);

  --3.没有结帐id的数据(按发生时间)
  --a.未结帐的划价记录
  --b.未收费的零费用
  --加条件"待转出 Is Null"是为了处理连续多次标记转出的情况
  Update /*+ rule*/ 门诊费用记录 A
  Set 待转出 = n_批次
  Where (记录状态 = 0 Or 记录性质 = 1 And 实收金额 = 0 And 结帐金额 = 0) And 结帐id Is Null And 待转出 Is Null And 发生时间 < d_End;

  --4.没有结帐id的数据(按发生时间)
  --未结帐的门诊记帐费用(赖账)，该病人没有预交余额，并且病人在最终转出时间之后无未结门诊记帐费用
  Update /*+ rule*/ 门诊费用记录 A
  Set 待转出 = n_批次
  Where Not Exists (Select 1
         From 病人预交记录 B
         Where b.病人id = a.病人id And b.待转出 Is Null And b.预交类别 = 1 And b.记录性质 In (1, 11) Having
          Nvl(Sum(b.金额), 0) <> Nvl(Sum(b.冲预交), 0)) And Not Exists
   (Select 1
         From 门诊费用记录 B
         Where a.病人id = b.病人id And b.记录性质 = 2 And b.结帐id Is Null And b.待转出 Is Null And b.登记时间 > = d_Lastend) And
        记录性质 = 2 And 结帐id Is Null And 待转出 Is Null And 发生时间 < d_End;

  --5.没有结帐id的数据(按发生时间)
  --冲销产生的记帐记录（记录状态为2），登记时间可能在当前指定转出时间之后，而原始记帐记录（记录状态为3），登记时间在指定转出时间之前。前后两者的发生时间是相同的。
  --a.未结帐的零记帐费用或打折后实收金额为零的（结帐模块参数没有勾选对零费用结帐）
  --b.结帐作废后，记帐单销帐的记录（结帐ID为空且记录状态为2的），记录状态为3的且有结帐ID的在最前面已转出.
  Update /*+ rule*/ 门诊费用记录 A
  Set 待转出 = n_批次
  Where (Exists (Select 1
                 From 门诊费用记录 B
                 Where a.No = b.No And a.记录性质 = b.记录性质 And a.序号 = b.序号 And b.记录状态 = 3 And b.结帐id Is Not Null And
                       b.待转出 + 0 = n_批次) And 记录状态 = 2 Or Exists
         (Select 1
          From 门诊费用记录 B
          Where a.No = b.No And a.记录性质 = b.记录性质 And a.序号 = b.序号 And b.结帐id Is Null
          Group By b.No, b.记录性质, b.序号
          Having Nvl(Sum(b.实收金额), 0) = 0)) And 记录性质 = 2 And 结帐id Is Null And 待转出 Is Null And 发生时间 < d_End;

  --6.有结帐id的零费用(按发生时间)
  --a.按费别打折后结帐金额为零的收费记录,
  --b.一张单据相同结帐ID的结帐金额之和为0(冲销后为零)
  --即使在转出时间之后发药的，也强制转出（为了减少逻辑复杂性，提高查询性能）
  Update /*+ rule*/ 门诊费用记录 A
  Set 待转出 = n_批次
  Where (结帐金额 = 0 Or Exists
         (Select 1 From 门诊费用记录 C Where a.结帐id = c.结帐id Group By c.结帐id, c.No Having Sum(c.结帐金额) = 0)) And Not Exists
   (Select 1 From 病人预交记录 B Where a.结帐id = b.结帐id And b.待转出 Is Null) And 记录性质 = 1 And 结帐id Is Not Null And
        待转出 Is Null And 发生时间 < d_End;

  --收费票据（不严格控制）
  Update /*+ rule*/ 票据使用明细
  Set 待转出 = n_批次
  Where 票种 = 1 And
        号码 In
        (Select 实际票号 From 门诊费用记录 Where 待转出 = n_批次 And Mod(记录性质, 10) = 1 And Nvl(费用状态, 0) = 0) And Nvl(领用id, 0) = 0;

  --挂号票据（不严格控制）
  Update /*+ rule*/ 票据使用明细
  Set 待转出 = n_批次
  Where 票种 = 4 And
        号码 In (Select 实际票号 From 门诊费用记录 Where 待转出 = n_批次 And 记录性质 = 4 And Nvl(费用状态, 0) = 0) And Nvl(领用id, 0) = 0;

  Update /*+ rule*/ 费用结算对照
  Set 待转出 = n_批次
  Where 结帐id In (Select 结帐id From 病人预交记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 医保结算明细
  Set 待转出 = n_批次
  Where 结帐id In (Select 结帐id From 病人预交记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 费用补充记录
  Set 待转出 = n_批次
  Where 结算id In (Select 结帐id From 病人预交记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 凭条打印记录
  Set 待转出 = n_批次
  Where (NO, 记录性质) In (Select NO, 记录性质 From 门诊费用记录 Where 待转出 = n_批次);

  --1.从预交记录读是为了取就诊卡直接收费的（无结帐ID）,再加结帐记录是为了取结帐无结算（预交）记录的
  Update /*+ rule*/ 住院费用记录
  Set 待转出 = n_批次
  Where 结帐id In (Select 结帐id
                 From 病人预交记录
                 Where 待转出 = n_批次
                 Union
                 Select ID From 病人结帐记录 Where 待转出 = n_批次);

  --2.没有结帐id的数据(按发生时间)
  --冲销产生的记帐记录（记录状态为2），原始记录和冲销记录的发生时间是相同的。
  --1)转出结帐作废后，记帐单销帐的记录（记录状态为2，且没有结帐ID，且(记录状态为3的有结帐ID的)在最前面已转出）
  --2)未结帐的零费用(已冲销的记帐单或打折后实收金额为零)
  --3)没有结帐ID的划价记录处理为转出
  Update /*+ rule*/ 住院费用记录 A
  Set 待转出 = n_批次
  Where ((Exists (Select 1
                  From 住院费用记录 B
                  Where a.No = b.No And a.记录性质 = b.记录性质 And a.序号 = b.序号 And b.记录状态 = 3 And b.结帐id Is Not Null And
                        b.待转出 + 0 = n_批次) And 记录状态 = 2 Or Exists
         (Select 1
           From 住院费用记录 B
           Where a.No = b.No And a.记录性质 = b.记录性质 And a.序号 = b.序号 And b.结帐id Is Null
           Group By b.No, b.记录性质, b.序号
           Having Nvl(Sum(b.实收金额), 0) = 0)) And a.记录性质 In (2, 3, 5) Or a.记录状态 = 0) And a.结帐id Is Null And a.待转出 Is Null And
        a.发生时间 < d_End;

  --3.离院未结帐的（赖帐病人），因为是很久以前的这些数据，如果预交已冲完，则处理为要转出
  --去掉病案主页中的"数据转出 is null"的条件，是因为一些病人可能在之前的批次中已转出了
  Update /*+ rule*/ 住院费用记录 A
  Set 待转出 = n_批次
  Where 待转出 Is Null And 结帐id Is Null And
        (病人id, 主页id) In (Select 病人id, 主页id
                         From 病案主页 C
                         Where 出院日期 < d_End And 待转出 Is Null And Not Exists
                          (Select 1
                                From 病人预交记录 B
                                Where b.病人id = c.病人id And b.待转出 Is Null And b.预交类别 = 2 And b.记录性质 In (1, 11) Having
                                 Nvl(Sum(b.金额), 0) <> Nvl(Sum(b.冲预交), 0)));
  --就诊卡票据（不严格控制）
  Update /*+ rule*/ 票据使用明细
  Set 待转出 = n_批次
  Where 票种 = 5 And
        号码 In
        (Select 实际票号 From 住院费用记录 Where 待转出 = n_批次 And Mod(记录性质, 10) = 5 And Nvl(费用状态, 0) = 0) And Nvl(领用id, 0) = 0;

  Update /*+ rule*/ 费用清单打印
  Set 待转出 = n_批次
  Where (NO, Mod(记录性质, 10), Decode(记录状态, 3, 1, 记录状态), 序号) In
        (Select NO, Mod(记录性质, 10) As 记录性质, Decode(记录状态, 3, 1, 记录状态) As 记录状态, 序号
         From 门诊费用记录
         Where 待转出 = n_批次
         Union
         Select NO, Mod(记录性质, 10) As 记录性质, Decode(记录状态, 3, 1, 记录状态) As 记录状态, 序号
         From 住院费用记录
         Where 待转出 = n_批次);

  Update /*+ rule*/ 费用变动记录
  Set 待转出 = n_批次
  Where 费用id In (Select ID From 住院费用记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人费用销帐
  Set 待转出 = n_批次
  Where 费用id In (Select ID From 住院费用记录 Where 待转出 = n_批次);

  Update zlDataMovelog
  Set 当前进度 = '(2/11)费用数据标记完成，正在标记药品数据'
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

  Update /*+ rule*/ 药品收发门诊标志 A
  Set 待转出 = n_批次
  Where (a.处方号, a.单据) In (Select b.No, b.单据 From 药品收发记录 B Where b.待转出 = n_批次);

  Update /*+ rule*/ 药品收发住院标志
  Set 待转出 = n_批次
  Where 收发id In (Select ID From 药品收发记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 未审药品记录
  Set 待转出 = n_批次
  Where 收发id In (Select ID From 药品收发记录 Where 待转出 = n_批次);

  Update zlDataMovelog
  Set 当前进度 = '(3/11)药品数据标记完成，正在标记缴款与票据数据'
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
  Where Not Exists (Select 1 From 票据使用明细 B Where b.领用id = a.Id And b.使用时间 >= d_Lastend) And 待转出 Is Null And 剩余数量 = 0 And
        登记时间 < d_End;

  Update /*+ rule*/ 票据使用明细
  Set 待转出 = n_批次
  Where 领用id In (Select ID From 票据领用记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 票据打印内容
  Set 待转出 = n_批次
  Where ID In (Select 打印id From 票据使用明细 Where 待转出 = n_批次);

  Update /*+ rule*/ 票据打印明细
  Set 待转出 = n_批次
  Where 使用id In (Select ID From 票据使用明细 Where 待转出 = n_批次);

  Update /*+ rule*/ 急诊就诊记录
  Set 待转出 = n_批次
  Where 挂号id In (Select ID From 病人挂号记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 急诊分诊记录
  Set 待转出 = n_批次
  Where 就诊id In (Select ID From 急诊就诊记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 急诊病人评分
  Set 待转出 = n_批次
  Where 分诊id In (Select ID From 急诊分诊记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 急诊病人评分指标
  Set 待转出 = n_批次
  Where 评分id In (Select ID From 急诊病人评分 Where 待转出 = n_批次);

  Update zlDataMovelog
  Set 当前进度 = '(4/11)缴款与票据数据标记完成，正在标记就诊及诊治数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  --2.就诊及诊治数据
  --不转出的条件：挂号费用未转出的，最终转出时间之后存在医嘱（这些医嘱因为时间没有到，不应转出），医嘱对应的费用未转出的
  --即使正在就诊(r.执行状态 <> 2 )的也强制转出(医生可能没有使用完成就诊功能)
  Update /*+ rule*/ 病人挂号记录 T
  Set 待转出 = n_批次
  Where Rowid In
        (Select Rowid
         From 病人挂号记录 R
         Where Not Exists (Select 1 From 门诊费用记录 A Where r.No = a.No And a.记录性质 = 4 And a.待转出 Is Null) And Not Exists
          (Select 1
                From 病人医嘱记录 A
                Where a.挂号单 = r.No And a.待转出 Is Null And a.病人来源 <> 4 And Nvl(a.停嘱时间, a.开嘱时间) >= d_Lastend) And
               Not Exists (Select 1
                From 门诊费用记录 E, 病人医嘱记录 A
                Where r.No = a.挂号单 And a.Id = e.医嘱序号 And a.病人来源 <> 4 And e.待转出 Is Null) And
               r.待转出 Is Null And r.发生时间 < d_End);

  --由于有一部分挂号数据未转出，所以，汇总表的数据可能与挂号数据不匹配
  Update 病人挂号汇总 Set 待转出 = n_批次 Where 待转出 Is Null And 日期 < d_End;
  Update /*+ rule*/ 病人转诊记录 Set 待转出 = n_批次 Where NO In (Select NO From 病人挂号记录 Where 待转出 = n_批次);

  --通过"住院费用记录"来查询，而不是"病人结帐记录",因为离院未结的赖帐病人也转出了费用
  --出院日期条件仍然需要，因为可能某次结帐转出了，但病人在最终转出截止时间之前并未出院(一次住院多次结帐)。
  --通过指定索引方式进行特殊优化（缺省采用"病案主页IX_出院日期"索引的效率太低）
  --不加"数据转出 is null"的条件，因为一次住院多次结帐时，如果跨不同的转出批次(转出截止时间)，该字段将会被更新多次。
  Update /*+ rule*/ 病案主页 P
  Set 待转出 = n_批次
  Where Not Exists
   (Select 1 From 住院费用记录 A Where a.病人id = p.病人id And a.主页id = p.主页id And a.待转出 Is Null) And 待转出 Is Null And
        出院日期 < d_Lastend And (病人id, 主页id) In (Select Distinct 病人id, 主页id From 住院费用记录 Where 待转出 = n_批次);

  --已出院，但没有费用的，也标记为转出，以便转出病历数据
  Update /*+ rule*/ 病案主页 P
  Set 待转出 = n_批次
  Where Not Exists (Select 1 From 住院费用记录 A Where a.病人id = p.病人id And a.主页id = p.主页id) And 待转出 Is Null And 数据转出 Is Null And
        出院日期 < d_End;

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

  Update zlDataMovelog
  Set 当前进度 = '(5/11)就诊及诊治数据标记完成，正在标记护理数据'
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

  Update /*+ rule*/ 病人护理诊断
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

  Update zlDataMovelog
  Set 当前进度 = '(6/11)护理数据标记完成，正在标记病历数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  --4.病历数据
  Update /*+ rule*/ 电子病历记录
  Set 待转出 = n_批次
  Where 病人来源 = 1 And (病人id, 主页id) In (Select 病人id, ID From 病人挂号记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 电子病历记录
  Set 待转出 = n_批次
  Where 病人来源 = 2 And (病人id, 主页id) In (Select 病人id, 主页id From 病案主页 Where 待转出 = n_批次);

  --自登记类病人(无挂号单号)
  --病历ID可能重复是因为检验报告之类的，如肝功、肾功共打一张报告，即在病人医嘱报告表中，多个医嘱id对应同一报告ID
  --为提升性能，不从医嘱发送记录的发送时间查询，不采用精确的时间，因为直接登记的检验医嘱，一般开嘱时间与发送时间相差不大
  --有些特殊（错误）数据，挂号单为空的医嘱，除了来源为3的（直接登记的检查检验医嘱），还可能有来源为1或4的（门诊或体检医嘱），主页ID可能不是0
  Update /*+ rule*/ 电子病历记录
  Set 待转出 = n_批次
  Where 待转出 Is Null And 病历种类 = 7 And ID In (Select c.病历id
                                            From 病人医嘱记录 B, 病人医嘱报告 C
                                            Where c.医嘱id = b.Id And b.病人来源 <> 2 And b.挂号单 Is Null And b.相关id Is Null And
                                                  b.待转出 Is Null And b.开嘱时间 < d_End);

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
  Where 病历id In (Select ID From 电子病历记录 Where 病历种类 = 7 And 待转出 = n_批次);

  Update /*+ rule*/ 影像报告驳回
  Set 待转出 = n_批次
  Where (医嘱id, 病历id) In (Select 医嘱id, 病历id From 病人医嘱报告 Where 待转出 = n_批次);

  Update /*+ rule*/ 医嘱报告内容
  Set 待转出 = n_批次
  Where ID In (Select 报告id From 病人医嘱报告 Where 待转出 = n_批次);

  Update /*+ rule*/ 报告查阅记录
  Set 待转出 = n_批次
  Where 病历id In (Select ID From 电子病历记录 Where 病历种类 = 7 And 待转出 = n_批次);

  Update /*+ rule*/ 疾病申报记录
  Set 待转出 = n_批次
  Where 文件id In (Select ID From 电子病历记录 Where 病历种类 = 5 And 待转出 = n_批次);

  Update /*+ rule*/ 疾病报告反馈
  Set 待转出 = n_批次
  Where 文件id In (Select ID From 电子病历记录 Where 病历种类 = 5 And 待转出 = n_批次);

  Update /*+ rule*/ 疾病申报反馈
  Set 待转出 = n_批次
  Where 申报id In (Select ID From 电子病历记录 Where 病历种类 = 5 And 待转出 = n_批次);

  Update /*+ rule*/ 病人诊断记录
  Set 待转出 = n_批次
  Where 病历id In (Select ID From 电子病历记录 Where 待转出 = n_批次);

  Update zlDataMovelog
  Set 当前进度 = '(7/11)病历数据标记完成，正在标记临床路径数据'
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
  Update /*+ rule*/ 病人路径医嘱变异
  Set 待转出 = n_批次
  Where 路径执行id In (Select ID From 病人路径执行 Where 待转出 = n_批次);

  Update zlDataMovelog
  Set 当前进度 = '(8/11)临床路径数据标记完成，正在标记医嘱数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  --6.医嘱，检验，检查
  --加上病人来源，避免来源为3的自登记类病人误填了挂号单后，医嘱被转走了而医嘱报告没有转出
  Update /*+ rule*/ 病人医嘱记录
  Set 待转出 = n_批次
  Where 挂号单 In (Select NO From 病人挂号记录 Where 待转出 = n_批次) And 病人来源 = 1;

  --加上病人来源，避免 后，医嘱被转走了而医嘱报告没有转出
  Update /*+ rule*/ 病人医嘱记录
  Set 待转出 = n_批次
  Where (病人id, 主页id) In (Select 病人id, 主页id From 病案主页 Where 待转出 = n_批次) And 病人来源 = 2;

  --自登记类病人(无挂号单)，病人医嘱报告在前面转病历时已转出
  Update /*+ rule*/ 病人医嘱记录
  Set 待转出 = n_批次
  Where 待转出 Is Null And Rowid In (Select b.Rowid
                                  From 病人医嘱记录 B, 病人医嘱报告 C
                                  Where (b.相关id = c.医嘱id Or b.Id = c.医嘱id) And c.待转出 = n_批次);

  --自登记类病人(无挂号单)，没有医嘱报告
  Update /*+ rule*/ 病人医嘱记录 A
  Set 待转出 = n_批次
  Where Not Exists (Select 1 From 病人医嘱报告 B Where a.Id = b.医嘱id) And Not Exists
   (Select 1 From 病人医嘱报告 B Where a.相关id = b.医嘱id) And 挂号单 Is Null And 病人来源 = 3 And 待转出 Is Null And 开嘱时间 < d_End;

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
  Update /*+ rule*/ 输血申请项目
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

  Update /*+ rule*/ 医嘱执行组合
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

  Update /*+ rule*/ Ris检查预约
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 疾病阳性记录
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 医嘱申请单文件
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人危急值记录
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人危急值病历
  Set 待转出 = n_批次
  Where 危急值id In (Select ID From 病人危急值记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人危急值医嘱
  Set 待转出 = n_批次
  Where 危急值id In (Select ID From 病人危急值记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 药嘱禁忌说明
  Set 待转出 = n_批次
  Where 医嘱a In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 药嘱禁忌说明
  Set 待转出 = n_批次
  Where 医嘱b In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update zlDataMovelog
  Set 当前进度 = '(9/11)医嘱数据标记完成，正在标记检查检验数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  Update /*+ rule*/ 影像检查记录
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 影像报告记录
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 影像报告操作记录
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

  Update /*+ rule*/ 影像预约记录
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update zlDataMovelog
  Set 当前进度 = '(10/11)影像数据标记完成，正在标记检验数据'
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
  Update zlDataMovelog
  Set 当前进度 = '(11/11)检验数据标记完成，正在标记门诊临床路径数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  --11.门诊临床路径数据
  Update /*+ rule*/ 病人门诊路径记录
  Set 待转出 = n_批次
  Where 挂号id In (Select ID From 病人挂号记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人门诊路径记录 A
  Set 待转出 = Null
  Where Exists (Select 1 From 病人门诊路径记录 C Where a.路径记录id = c.路径记录id And c.待转出 Is Null) And 待转出 = n_批次;

  Update /*+ rule*/ 病人门诊路径
  Set 待转出 = n_批次
  Where ID In (Select 路径记录id From 病人门诊路径记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人门诊出径记录
  Set 待转出 = n_批次
  Where 路径记录id In (Select ID From 病人门诊路径 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人门诊路径执行
  Set 待转出 = n_批次
  Where 路径记录id In (Select ID From 病人门诊路径 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人门诊路径指标
  Set 待转出 = n_批次
  Where 路径记录id In (Select ID From 病人门诊路径 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人门诊路径评估
  Set 待转出 = n_批次
  Where 路径记录id In (Select ID From 病人门诊路径 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人门诊路径变异
  Set 待转出 = n_批次
  Where 路径记录id In (Select ID From 病人门诊路径 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人门诊路径医嘱
  Set 待转出 = n_批次
  Where 路径执行id In (Select ID From 病人门诊路径执行 Where 待转出 = n_批次);

  --12.病人用药清单
  Update /*+ rule*/ 病人用药清单
  Set 待转出 = n_批次
  Where (病人id, 主页id) In (Select 病人id, 主页id From 病案主页 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人用药配方
  Set 待转出 = n_批次
  Where 配方id In (Select ID From 病人用药清单 Where 待转出 = n_批次);

  Commit;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl1_Datamove_Tag;
/

--报表：ZL1_INSIDE_1145/电子票据告知单
Insert Into zlReports(ID,分类ID,编号,名称,说明,密码,打印机,进纸,票据,打印方式,系统,程序ID,功能,修改时间,发布时间,禁止开始时间,禁止结束时间) Values(zlReports_ID.NextVal,Null,'ZL1_INSIDE_1145','电子票据告知单','电子票据专用','J~(f^sl{}+=EpjkvM"QT','Microsoft XPS Document Writer',15,0,0,&n_system,Null,Null,Sysdate,Sysdate,To_Date('2020-05-01 00:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('2020-05-01 00:00:00','YYYY-MM-DD HH24:MI:SS'));
Insert Into zlRPTPuts(报表ID,系统,程序ID,功能) Values(zlReports_ID.CurrVal,&n_system,1103,'电子票据告知单');
Insert Into zlRPTPuts(报表ID,系统,程序ID,功能) Values(zlReports_ID.CurrVal,&n_system,1107,'电子票据告知单');
Insert Into zlRPTPuts(报表ID,系统,程序ID,功能) Values(zlReports_ID.CurrVal,&n_system,1111,'电子票据告知单');
Insert Into zlRPTPuts(报表ID,系统,程序ID,功能) Values(zlReports_ID.CurrVal,&n_system,1121,'电子票据告知单');
Insert Into zlRPTPuts(报表ID,系统,程序ID,功能) Values(zlReports_ID.CurrVal,&n_system,1124,'电子票据告知单');
Insert Into zlRPTPuts(报表ID,系统,程序ID,功能) Values(zlReports_ID.CurrVal,&n_system,1137,'电子票据告知单');
Insert Into zlRPTPuts(报表ID,系统,程序ID,功能) Values(zlReports_ID.CurrVal,&n_system,1144,'电子票据告知单');
Insert Into zlRPTPuts(报表ID,系统,程序ID,功能) Values(zlReports_ID.CurrVal,&n_system,1145,'电子票据告知单');
Insert Into zlRPTFmts(报表ID,序号,说明,W,H,纸张,纸向,动态纸张,图样,是否停用,停用原因) Values(zlReports_ID.CurrVal,1,'挂号告知单',5874,3855,256,1,0,0,Null,Null);
Insert Into zlRPTFmts(报表ID,序号,说明,W,H,纸张,纸向,动态纸张,图样,是否停用,停用原因) Values(zlReports_ID.CurrVal,2,'收费告知单',7749,4305,256,1,0,0,Null,Null);
Insert Into zlRPTFmts(报表ID,序号,说明,W,H,纸张,纸向,动态纸张,图样,是否停用,停用原因) Values(zlReports_ID.CurrVal,3,'结帐告知单',4494,4742,256,1,0,0,Null,Null);
Insert Into zlRPTFmts(报表ID,序号,说明,W,H,纸张,纸向,动态纸张,图样,是否停用,停用原因) Values(zlReports_ID.CurrVal,4,'预交告知单',4464,4455,256,1,0,0,Null,Null);
Insert Into zlRPTFmts(报表ID,序号,说明,W,H,纸张,纸向,动态纸张,图样,是否停用,停用原因) Values(zlReports_ID.CurrVal,5,'医疗卡告知单',6039,4487,256,1,0,0,Null,Null);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明,数据连接编号) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'二维码','二维码,205',User||'.电子票据二维码',0,Null,Null);
Insert Into zlRPTSQLs(源ID,行号,内容)  Select zlRPTDatas_ID.CurrVal,a.* From (
Select 1,'select  zltools.Zlbase64.Decode(二维码) as 二维码 from 电子票据二维码 where 使用记录id=[0]' From Dual) a;
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,0,'电子票据ID',1,Null,0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明,数据连接编号) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'挂号','NO,202|序号,139|数量,139|单价,139|应收金额,139|实收金额,139|号码,202|外网,202|姓名,202|性别,202|年龄,202|名称,202|规格,202|单位,202|标识号,139',User||'.费用补充记录,'||User||'.门诊费用记录,'||User||'.电子票据使用记录,'||User||'.收费项目目录',0,Null,Null);
Insert Into zlRPTSQLs(源ID,行号,内容)  Select zlRPTDatas_ID.CurrVal,a.* From (
Select 1,'Select a.No, a.序号, Sum(a.数次) As 数量, Avg(a.标准单价) As 单价, Sum(a.应收金额) As 应收金额, Sum(a.实收金额) As 实收金额, Max(a.号码) As 号码,' From Dual
Union All Select 2,'       Max(a.外网) As 外网, Max(a.姓名) As 姓名, Max(a.性别) As 性别, Max(a.年龄) As 年龄, Max(a.名称) As 名称, Max(a.规格) As 规格,' From Dual
Union All Select 3,'       Max(a.单位) As 单位,Max(标识号) As 标识号' From Dual
Union All Select 4,'From (' From Dual
Union All Select 5,'       --2.求单价' From Dual
Union All Select 6,'       Select a.No, Nvl(a.价格父号, a.序号) As 序号, Avg(Nvl(a.付数, 1) * a.数次) As 数次, Sum(a.标准单价) As 标准单价, Sum(a.应收金额) As 应收金额,' From Dual
Union All Select 7,'               Sum(a.实收金额) As 实收金额, Max(b.号码) As 号码, Max(b.外网) As 外网, Max(a.标识号) As 标识号, Max(a.姓名) As 姓名, Max(a.性别) As 性别,' From Dual
Union All Select 8,'               Max(a.年龄) As 年龄, Max(c.名称) As 名称, Max(c.规格) As 规格, Max(c.计算单位) As 单位' From Dual
Union All Select 9,'       From 门诊费用记录 A,' From Dual
Union All Select 10,'             (' From Dual
Union All Select 11,'               --1.找到原始结算的所有单据' From Dual
Union All Select 12,'               Select Distinct a.记录性质, a.No, a.序号, d.号码, d.Url外网 As 外网' From Dual
Union All Select 13,'               From 门诊费用记录 A, 电子票据使用记录 D' From Dual
Union All Select 14,'               Where a.结帐id = d.结算id And d.Id = [0] And Not Exists (Select 1 From 费用补充记录 Where 收费结帐id = a.结帐id) And d.票种 = 4' From Dual
Union All Select 15,'               Union All' From Dual
Union All Select 16,'               Select Distinct a.记录性质, a.No, a.序号, d.号码, d.Url外网 As 外网' From Dual
Union All Select 17,'               From 门诊费用记录 A, 费用补充记录 B, 电子票据使用记录 D' From Dual
Union All Select 18,'               Where a.结帐id = b.收费结帐id And b.结算id = d.结算id And d.Id = [0] And d.票种 = 4) B, 收费项目目录 C' From Dual
Union All Select 19,'       Where a.No = b.No And Mod(a.记录性质, 10) = b.记录性质 And a.序号 = b.序号 And a.收费细目id = c.Id' From Dual
Union All Select 20,'       Group By a.记录性质, a.记录状态, a.No, Nvl(a.价格父号, a.序号)) A' From Dual
Union All Select 21,'Group By a.No, a.序号' From Dual
Union All Select 22,'Having Nvl(Sum(a.数次), 0) <> 0' From Dual) a;
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,0,'电子票据ID',1,Null,0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明,数据连接编号) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'结帐','收据费目,202|姓名,202|性别,202|年龄,202|门诊号,139|住院号,139|金额,139|外网,202|号码,202',User||'.住院费用记录,'||User||'.病人结帐记录,'||User||'.电子票据使用记录,'||User||'.门诊费用记录,'||User||'.病人信息',0,Null,Null);
Insert Into zlRPTSQLs(源ID,行号,内容)  Select zlRPTDatas_ID.CurrVal,a.* From (
Select 1,'Select a.收据费目,b.姓名, b.性别, b.年龄,max(b.门诊号) as 门诊号,max(b.住院号) as 住院号, Sum(a.金额) As 金额, ' From Dual
Union All Select 2,'       Max(外网) As 外网,  Max(a.号码) As 号码' From Dual
Union All Select 3,'From (Select b.病人id, a.收据费目, Sum(a.结帐金额) As 金额, Max(c.Url外网) As 外网, Max(c.号码) As 号码' From Dual
Union All Select 4,'       From 住院费用记录 A, 病人结帐记录 B, 电子票据使用记录 C' From Dual
Union All Select 5,'       Where a.结帐id = b.Id And b.Id = c.结算id And c.票种 = 3 And c.记录状态 = 1 And b.记录状态 In (1, 3) And c.Id = [0]' From Dual
Union All Select 6,'       Group By b.病人id, a.收据费目' From Dual
Union All Select 7,'       Union All' From Dual
Union All Select 8,'       Select b.病人id, a.收据费目, Sum(a.结帐金额) As 金额, Max(c.Url外网) As 外网, Max(c.号码) As 号码' From Dual
Union All Select 9,'       From 门诊费用记录 A, 病人结帐记录 B, 电子票据使用记录 C' From Dual
Union All Select 10,'       Where a.结帐id = b.Id And b.Id = c.结算id And c.票种 = 3 And c.记录状态 = 1 And b.记录状态 In (1, 3) And c.Id = [0]' From Dual
Union All Select 11,'       Group By b.病人id, a.收据费目) A, 病人信息 B' From Dual
Union All Select 12,'Where a.病人id = b.病人id' From Dual
Union All Select 13,'Group By  a.收据费目,b.姓名, b.性别, b.年龄' From Dual) a;
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,0,'电子票据ID',1,Null,0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明,数据连接编号) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'收费','NO,202|序号,139|数量,139|单价,139|应收金额,139|实收金额,139|号码,202|外网,202|姓名,202|性别,202|年龄,202|名称,202|规格,202|单位,202|标识号,139',User||'.费用补充记录,'||User||'.门诊费用记录,'||User||'.电子票据使用记录,'||User||'.收费项目目录',0,Null,Null);
Insert Into zlRPTSQLs(源ID,行号,内容)  Select zlRPTDatas_ID.CurrVal,a.* From (
Select 1,'Select a.No, a.序号, Sum(a.数次) As 数量, Avg(a.标准单价) As 单价, Sum(a.应收金额) As 应收金额, Sum(a.实收金额) As 实收金额, Max(a.号码) As 号码,' From Dual
Union All Select 2,'       Max(a.外网) As 外网, Max(a.姓名) As 姓名, Max(a.性别) As 性别, Max(a.年龄) As 年龄, Max(a.名称) As 名称, Max(a.规格) As 规格,' From Dual
Union All Select 3,'       Max(a.单位) As 单位,Max(标识号) As 标识号' From Dual
Union All Select 4,'From (' From Dual
Union All Select 5,'       --2.求单价' From Dual
Union All Select 6,'       Select a.No, Nvl(a.价格父号, a.序号) As 序号, Avg(Nvl(a.付数, 1) * a.数次) As 数次, Sum(a.标准单价) As 标准单价, Sum(a.应收金额) As 应收金额,' From Dual
Union All Select 7,'               Sum(a.实收金额) As 实收金额, Max(b.号码) As 号码, Max(b.外网) As 外网, Max(a.标识号) As 标识号, Max(a.姓名) As 姓名, Max(a.性别) As 性别,' From Dual
Union All Select 8,'               Max(a.年龄) As 年龄, Max(c.名称) As 名称, Max(c.规格) As 规格, Max(c.计算单位) As 单位' From Dual
Union All Select 9,'       From 门诊费用记录 A,' From Dual
Union All Select 10,'             (' From Dual
Union All Select 11,'               --1.找到原始结算的所有单据' From Dual
Union All Select 12,'               Select Distinct a.记录性质, a.No, a.序号, d.号码, d.Url外网 As 外网' From Dual
Union All Select 13,'               From 门诊费用记录 A, 电子票据使用记录 D' From Dual
Union All Select 14,'               Where a.结帐id = d.结算id And d.Id = [0] And Not Exists (Select 1 From 费用补充记录 Where 收费结帐id = a.结帐id) And d.票种 = 1' From Dual
Union All Select 15,'               Union All' From Dual
Union All Select 16,'               Select Distinct a.记录性质, a.No, a.序号, d.号码, d.Url外网 As 外网' From Dual
Union All Select 17,'               From 门诊费用记录 A, 费用补充记录 B, 电子票据使用记录 D' From Dual
Union All Select 18,'               Where a.结帐id = b.收费结帐id And b.结算id = d.结算id And d.Id = [0] And d.票种 = 1) B, 收费项目目录 C' From Dual
Union All Select 19,'       Where a.No = b.No And Mod(a.记录性质, 10) = b.记录性质 And a.序号 = b.序号 And a.收费细目id = c.Id' From Dual
Union All Select 20,'       Group By a.记录性质, a.记录状态, a.No, Nvl(a.价格父号, a.序号)) A' From Dual
Union All Select 21,'Group By a.No, a.序号' From Dual
Union All Select 22,'Having Nvl(Sum(a.数次), 0) <> 0' From Dual) a;
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,0,'电子票据ID',1,Null,0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明,数据连接编号) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'医疗卡','标识号,131|NO,202|姓名,202|性别,202|年龄,202|名称,202|单价,131|数量,139|实收金额,131|号码,202|外网,202',User||'.住院费用记录,'||User||'.收费项目目录,'||User||'.电子票据使用记录',0,Null,Null);
Insert Into zlRPTSQLs(源ID,行号,内容)  Select zlRPTDatas_ID.CurrVal,a.* From (
Select 1,'Select a.标识号, a.No, a.姓名, a.性别, a.年龄, b.名称, a.标准单价 As 单价, Nvl(a.付数, 1) * a.数次 As 数量, a.实收金额, ' From Dual
Union All Select 2,'       c.号码,c.Url外网 As 外网' From Dual
Union All Select 3,'From 住院费用记录 A, 收费项目目录 B, 电子票据使用记录 C' From Dual
Union All Select 4,'Where a.结帐Id = c.结算id And c.票种 = 5 And c.记录状态 = 1 And a.收费细目id = b.Id And a.记录性质 = 5 And a.记录状态 = 1 And c.Id = [0]' From Dual) a;
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,0,'电子票据ID',1,Null,0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明,数据连接编号) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'预交','NO,202|金额,131|结算方式,202|号码,202|外网,202|姓名,202|性别,202|年龄,202|门诊号,131|住院号,131',User||'.病人预交记录,'||User||'.电子票据使用记录',0,Null,Null);
Insert Into zlRPTSQLs(源ID,行号,内容)  Select zlRPTDatas_ID.CurrVal,a.* From (
Select 1,'Select a.No, a.金额, a.结算方式, b.号码, b.Url外网 As 外网, b.姓名, b.性别, b.年龄,b.门诊号,b.住院号' From Dual
Union All Select 2,'From 病人预交记录 A, 电子票据使用记录 B' From Dual
Union All Select 3,'Where a.记录状态 = 1 And a.记录性质 = 1 And a.Id = b.结算id And b.票种 = 2 And b.Id = [0]' From Dual) a;
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,0,'电子票据ID',1,Null,0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签6',2,Null,0,Null,0,'电子发票号:[挂号.号码]',Null,450,720,1980,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签1',2,Null,0,Null,0,'标识号:[挂号.标识号]',Null,450,1035,2160,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签3',2,Null,0,Null,0,'性别:[挂号.性别]',Null,450,1380,1800,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签5',2,Null,0,Null,0,'挂号告知单',Null,1995,255,1650,330,0,0,1,'宋体',16,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签2',2,Null,0,Null,0,'姓名:[挂号.姓名]',Null,3450,1035,1800,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签4',2,Null,0,Null,0,'年龄:[挂号.年龄]',Null,3465,1380,1800,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签7',2,Null,0,Null,0,'[二维码.二维码]',Null,4215,30,1390,1200,0,0,1,'宋体',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'任意表1',4,Null,0,Null,0,'挂号',Null,465,1755,5030,1395,255,0,0,'宋体',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-1,0,Null,Null,'[挂号.NO]','4^225^NO',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-2,1,Null,Null,'[挂号.名称]','4^225^名称',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-3,2,Null,Null,'[挂号.单价]','4^225^单价',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-4,3,Null,Null,'[挂号.数量]','4^225^数量',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-5,4,Null,Null,'[挂号.实收金额]','4^225^实收金额',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,'标签5',2,Null,0,Null,0,'电子发票号:[收费.号码]',Null,435,825,1980,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,'标签2',2,Null,0,Null,0,'姓名:[收费.姓名]',Null,450,1200,1440,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,'收费告知单',2,Null,0,Null,0,'收费告知单',Null,2520,210,1650,330,0,0,1,'宋体',16,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,'标签1',2,Null,0,Null,0,'标志号:[收费.标识号]',Null,3000,810,1800,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,'标签3',2,Null,0,Null,0,'性别:[收费.性别]',Null,3030,1140,1440,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,'标签4',2,Null,0,Null,0,'年龄:[收费.年龄]',Null,5790,1155,1440,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,'标签6',2,Null,0,Null,0,'[二维码.二维码]',Null,5835,30,1375,1125,0,0,0,'宋体',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,'任意表1',4,Null,0,Null,0,'收费',Null,420,1590,7030,2475,255,0,0,'宋体',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,Null,6,zlRPTItems_ID.CurrVal-1,0,Null,Null,'[收费.NO]','4^300^NO',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,Null,6,zlRPTItems_ID.CurrVal-2,1,Null,Null,'[收费.名称]','4^300^名称',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,Null,6,zlRPTItems_ID.CurrVal-3,2,Null,Null,'[收费.规格]','4^300^规格',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,Null,6,zlRPTItems_ID.CurrVal-4,3,Null,Null,'[收费.单位]','4^300^单位',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,Null,6,zlRPTItems_ID.CurrVal-5,4,Null,Null,'[收费.单价]','4^300^单价',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,Null,6,zlRPTItems_ID.CurrVal-6,5,Null,Null,'[收费.数量]','4^300^数量',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,2,Null,6,zlRPTItems_ID.CurrVal-7,6,Null,Null,'[收费.实收金额]','4^300^实收金额',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,3,'标签7',2,Null,0,Null,0,'电子发票号:[结帐.号码]',Null,375,855,1980,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,3,'标签3',2,Null,0,Null,0,'姓名:[结帐.姓名]',Null,375,1245,2160,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,3,'标签5',2,Null,0,Null,0,'年龄:[结帐.年龄]',Null,375,1605,1440,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,3,'标签2',2,Null,0,Null,0,'门诊号:[结帐.门诊号]',Null,375,1995,1800,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,3,'标签1',2,Null,0,Null,0,'结帐告知单',Null,1215,255,1650,330,0,0,1,'宋体',16,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,3,'标签4',2,Null,0,Null,0,'性别:[结帐.性别]',Null,2265,1230,1440,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,3,'标签6',2,Null,0,Null,0,'住院号:[结帐.住院号]',Null,2315,1995,1800,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,3,'标签8',2,Null,0,Null,0,'[二维码.二维码]',Null,2940,150,1345,1080,0,0,0,'宋体',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,3,'任意表1',4,Null,0,Null,0,'结帐',Null,405,2385,3601,1935,255,0,0,'宋体',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,3,Null,6,zlRPTItems_ID.CurrVal-1,0,Null,Null,'[结帐.收据费目]','4^330^收据费目',0,0,1860,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,3,Null,6,zlRPTItems_ID.CurrVal-2,1,Null,Null,'[结帐.金额]','4^330^金额',0,0,1485,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,4,'标签8',2,Null,0,Null,0,'电子发票号:[预交.号码]',Null,435,855,1980,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,4,'标签3',2,Null,0,Null,0,'姓名:[预交.姓名]',Null,435,1200,1440,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,4,'标签5',2,Null,0,Null,0,'年龄:[预交.年龄]',Null,435,1530,1440,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,4,'标签2',2,Null,0,Null,0,'门诊号:[预交.门诊号]',Null,435,1845,1800,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,4,'标签1',2,Null,0,Null,0,'预交告知单',Null,1305,315,1650,330,0,0,1,'宋体',16,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,4,'标签6',2,Null,0,Null,0,'住院号:[预交.住院号]',Null,2375,1845,1800,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,4,'标签4',2,Null,0,Null,0,'性别:[预交.性别]',Null,2400,1200,1440,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,4,'标签9',2,Null,0,Null,0,'[二维码.二维码]',Null,2910,75,1300,1245,0,0,0,'宋体',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,4,'任意表1',4,Null,0,Null,0,'预交',Null,450,2235,3480,1770,255,0,0,'宋体',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,4,Null,6,zlRPTItems_ID.CurrVal-1,0,Null,Null,'[预交.NO]','4^270^NO',0,0,1110,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,4,Null,6,zlRPTItems_ID.CurrVal-2,1,Null,Null,'[预交.金额]','4^270^金额',0,0,1230,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,4,Null,6,zlRPTItems_ID.CurrVal-3,2,Null,Null,'[预交.结算方式]','4^270^结算方式',0,0,1050,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,5,'标签6',2,Null,0,Null,0,'电子发票号:[医疗卡.号码]',Null,495,890,2160,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,5,'标签2',2,Null,0,Null,0,'标识号:[医疗卡.标识号]',Null,495,1320,1980,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,5,'标签4',2,Null,0,Null,0,'性别:[医疗卡.性别]',Null,495,1785,1620,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,5,'标签1',2,Null,0,Null,0,'医疗卡告知单',Null,1935,285,1980,330,0,0,1,'宋体',16,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,5,'标签3',2,Null,0,Null,0,'姓名:[医疗卡.姓名]',Null,3705,1320,1620,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,5,'标签5',2,Null,0,Null,0,'年龄:[医疗卡.年龄]',Null,3705,1785,1620,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,5,'标签7',2,Null,0,Null,0,'[二维码.二维码]',Null,4125,150,1300,1110,0,0,0,'宋体',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,5,'任意表1',4,Null,0,Null,0,'医疗卡',Null,495,2280,5030,1545,255,0,0,'宋体',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,5,Null,6,zlRPTItems_ID.CurrVal-1,0,Null,Null,'[医疗卡.NO]','4^225^NO',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,5,Null,6,zlRPTItems_ID.CurrVal-2,1,Null,Null,'[医疗卡.名称]','4^225^名称',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,5,Null,6,zlRPTItems_ID.CurrVal-3,2,Null,Null,'[医疗卡.单价]','4^225^单价',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,5,Null,6,zlRPTItems_ID.CurrVal-4,3,Null,Null,'[医疗卡.数量]','4^225^数量',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转,拆分单元格,自动填充,关联表格) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,5,Null,6,zlRPTItems_ID.CurrVal-5,4,Null,Null,'[医疗卡.实收金额]','4^225^实收金额',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null,Null,Null,Null);

--报表：ZL1_INSIDE_1145/电子票据告知单
Insert into zlProgFuncs(系统,序号,功能,说明) Values(n_System,1103,'电子票据告知单','电子票据专用');
Insert into zlProgFuncs(系统,序号,功能,说明) Values(n_System,1107,'电子票据告知单','电子票据专用');
Insert into zlProgFuncs(系统,序号,功能,说明) Values(n_System,1111,'电子票据告知单','电子票据专用');
Insert into zlProgFuncs(系统,序号,功能,说明) Values(n_System,1121,'电子票据告知单','电子票据专用');
Insert into zlProgFuncs(系统,序号,功能,说明) Values(n_System,1124,'电子票据告知单','电子票据专用');
Insert into zlProgFuncs(系统,序号,功能,说明) Values(n_System,1137,'电子票据告知单','电子票据专用');
Insert into zlProgFuncs(系统,序号,功能,说明) Values(n_System,1144,'电子票据告知单','电子票据专用');
Insert into zlProgFuncs(系统,序号,功能,说明) Values(n_System,1145,'电子票据告知单','电子票据专用');
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
select &n_system,1103,'电子票据告知单',User,'病人结帐记录','SELECT' From Dual
Union All select &n_system,1103,'电子票据告知单',User,'病人预交记录','SELECT' From Dual
Union All select &n_system,1103,'电子票据告知单',User,'电子票据二维码','SELECT' From Dual
Union All select &n_system,1103,'电子票据告知单',User,'电子票据使用记录','SELECT' From Dual
Union All select &n_system,1103,'电子票据告知单',User,'费用补充记录','SELECT' From Dual
Union All select &n_system,1103,'电子票据告知单',User,'门诊费用记录','SELECT' From Dual
Union All select &n_system,1103,'电子票据告知单',User,'收费项目目录','SELECT' From Dual
Union All select &n_system,1103,'电子票据告知单',User,'住院费用记录','SELECT' From Dual
Union All select &n_system,1107,'电子票据告知单',User,'病人结帐记录','SELECT' From Dual
Union All select &n_system,1107,'电子票据告知单',User,'病人预交记录','SELECT' From Dual
Union All select &n_system,1107,'电子票据告知单',User,'电子票据二维码','SELECT' From Dual
Union All select &n_system,1107,'电子票据告知单',User,'电子票据使用记录','SELECT' From Dual
Union All select &n_system,1107,'电子票据告知单',User,'费用补充记录','SELECT' From Dual
Union All select &n_system,1107,'电子票据告知单',User,'门诊费用记录','SELECT' From Dual
Union All select &n_system,1107,'电子票据告知单',User,'住院费用记录','SELECT' From Dual
Union All select &n_system,1111,'电子票据告知单',User,'病人结帐记录','SELECT' From Dual
Union All select &n_system,1111,'电子票据告知单',User,'病人预交记录','SELECT' From Dual
Union All select &n_system,1111,'电子票据告知单',User,'电子票据二维码','SELECT' From Dual
Union All select &n_system,1111,'电子票据告知单',User,'电子票据使用记录','SELECT' From Dual
Union All select &n_system,1111,'电子票据告知单',User,'费用补充记录','SELECT' From Dual
Union All select &n_system,1111,'电子票据告知单',User,'门诊费用记录','SELECT' From Dual
Union All select &n_system,1111,'电子票据告知单',User,'住院费用记录','SELECT' From Dual
Union All select &n_system,1121,'电子票据告知单',User,'病人结帐记录','SELECT' From Dual
Union All select &n_system,1121,'电子票据告知单',User,'病人预交记录','SELECT' From Dual
Union All select &n_system,1121,'电子票据告知单',User,'电子票据二维码','SELECT' From Dual
Union All select &n_system,1121,'电子票据告知单',User,'电子票据使用记录','SELECT' From Dual
Union All select &n_system,1121,'电子票据告知单',User,'费用补充记录','SELECT' From Dual
Union All select &n_system,1121,'电子票据告知单',User,'门诊费用记录','SELECT' From Dual
Union All select &n_system,1121,'电子票据告知单',User,'住院费用记录','SELECT' From Dual
Union All select &n_system,1124,'电子票据告知单',User,'病人结帐记录','SELECT' From Dual
Union All select &n_system,1124,'电子票据告知单',User,'病人预交记录','SELECT' From Dual
Union All select &n_system,1124,'电子票据告知单',User,'电子票据二维码','SELECT' From Dual
Union All select &n_system,1124,'电子票据告知单',User,'电子票据使用记录','SELECT' From Dual
Union All select &n_system,1124,'电子票据告知单',User,'费用补充记录','SELECT' From Dual
Union All select &n_system,1124,'电子票据告知单',User,'门诊费用记录','SELECT' From Dual
Union All select &n_system,1124,'电子票据告知单',User,'住院费用记录','SELECT' From Dual
Union All select &n_system,1137,'电子票据告知单',User,'病人结帐记录','SELECT' From Dual
Union All select &n_system,1137,'电子票据告知单',User,'病人预交记录','SELECT' From Dual
Union All select &n_system,1137,'电子票据告知单',User,'电子票据二维码','SELECT' From Dual
Union All select &n_system,1137,'电子票据告知单',User,'电子票据使用记录','SELECT' From Dual
Union All select &n_system,1137,'电子票据告知单',User,'费用补充记录','SELECT' From Dual
Union All select &n_system,1137,'电子票据告知单',User,'门诊费用记录','SELECT' From Dual
Union All select &n_system,1137,'电子票据告知单',User,'住院费用记录','SELECT' From Dual
Union All select &n_system,1144,'电子票据告知单',User,'病人结帐记录','SELECT' From Dual
Union All select &n_system,1144,'电子票据告知单',User,'病人信息','SELECT' From Dual
Union All select &n_system,1144,'电子票据告知单',User,'病人预交记录','SELECT' From Dual
Union All select &n_system,1144,'电子票据告知单',User,'电子票据二维码','SELECT' From Dual
Union All select &n_system,1144,'电子票据告知单',User,'电子票据使用记录','SELECT' From Dual
Union All select &n_system,1144,'电子票据告知单',User,'费用补充记录','SELECT' From Dual
Union All select &n_system,1144,'电子票据告知单',User,'门诊费用记录','SELECT' From Dual
Union All select &n_system,1144,'电子票据告知单',User,'收费项目目录','SELECT' From Dual
Union All select &n_system,1144,'电子票据告知单',User,'住院费用记录','SELECT' From Dual
Union All select &n_system,1145,'电子票据告知单',User,'病人结帐记录','SELECT' From Dual
Union All select &n_system,1145,'电子票据告知单',User,'病人信息','SELECT' From Dual
Union All select &n_system,1145,'电子票据告知单',User,'病人预交记录','SELECT' From Dual
Union All select &n_system,1145,'电子票据告知单',User,'电子票据二维码','SELECT' From Dual
Union All select &n_system,1145,'电子票据告知单',User,'电子票据使用记录','SELECT' From Dual
Union All select &n_system,1145,'电子票据告知单',User,'费用补充记录','SELECT' From Dual
Union All select &n_system,1145,'电子票据告知单',User,'门诊费用记录','SELECT' From Dual
Union All select &n_system,1145,'电子票据告知单',User,'收费项目目录','SELECT' From Dual
Union All select &n_system,1145,'电子票据告知单',User,'住院费用记录','SELECT' From Dual;


