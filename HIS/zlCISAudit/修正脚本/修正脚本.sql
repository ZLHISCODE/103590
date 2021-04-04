
-----------------------------------------------------------------------------------------------------------------------
--    电子病案审查归档   表空间:zl9CISAduit
-----------------------------------------------------------------------------------------------------------------------
--电子病案审查归档
Create Sequence 病案提交记录_ID start with 1;
Create Sequence 病案反馈记录_ID start with 1;
Create Sequence 病案借阅记录_ID start with 1;
Create Sequence 病案封存记录_ID start with 1;
Create Sequence 病案评分方案_ID start with 1;
Create Sequence 病案评分标准_ID start with 1;
Create Sequence 病案评分结果_ID start with 1;
Create Sequence 病案评分明细_ID start with 1;

Create Table 病案提交记录(
    ID			Number(18),
    病人id		Number(18),
    主页id		Number(5),
    记录状态	Number(3),
    提交人		Varchar2(20),
    提交时间	Date,
    接收人		Varchar2(20),
    接收时间	Date,
    归档人		Varchar2(20),
    归档时间	Date,
    拒审人		Varchar2(20),
    拒审时间	Date,
    拒审理由	Varchar2(255))
    TableSpace zl9CISAduit
    PCTFREE 5 PCTUSED 85;

Create Table 病案审阅书签(
    提交id		Number(18),
    审阅对象	Number(3),
    文件id		Number(18),
    审阅时间	Date)
    TableSpace zl9CISAduit
    PCTFREE 5 PCTUSED 85;

Create Table 病案反馈记录(
    ID			Number(18),
    相关id		Number(18),
    提交id		Number(18),
    病人id		Number(18),
    主页id		Number(5),
    反馈对象	Number(3),
    文件id		Number(18),
    记录性质	Number(3),
    记录状态	Number(3),
    反馈意见	Varchar2(255),
    反馈项目id	Number(18),
    反馈人		Varchar2(20),
    反馈时间	Date,
    处理期限	Date,
    处理说明	Varchar2(255),
    处理人		Varchar2(20),
    处理时间	Date)
    TableSpace zl9CISAduit
    PCTFREE 5 PCTUSED 85;

Create Table 病案反馈历史(
    ID			Number(18),
    相关id		Number(18),
    提交id		Number(18),
    病人id		Number(18),
    主页id		Number(5),
    反馈对象	Number(3),
    文件id		Number(18),
    记录性质	Number(3),
    记录状态	Number(3),
    反馈意见	Varchar2(255),
    反馈项目id	Number(18),
    反馈人		Varchar2(20),
    反馈时间	Date,
    处理期限	Date,
    处理说明	Varchar2(255),
    处理人		Varchar2(20),
    处理时间	Date)
    TableSpace zl9CISAduit
    PCTFREE 5 PCTUSED 85;

Create Table 病案借阅记录(
	ID		Number(18),
	No		Varchar2(10),
	记录状态	Number(3),   
	申请人	Varchar2(20),	
	申请理由	Varchar2(255),
	申请时间	Date,
	申请期限	Date,
	借阅时间	Date,
	借阅期限	Date,
	批准人	Varchar2(20),
	批准时间	Date,
	拒借理由	Varchar2(255),
	拒借人	Varchar2(20),
	拒借时间	Date,
	登记时间	Date)
	TABLESPACE zl9CISAduit
	PCTFREE 5 PCTUSED 85;

Create Table 病案借阅内容(
    借阅id		Number(18),
    病人id		Number(18),
    主页id		Number(5))
    TableSpace zl9CISAduit
    PCTFREE 5 PCTUSED 85;

Create Table 病案借阅人员(
    借阅id		Number(18),
    人员id		Number(18))
    TableSpace zl9CISAduit
    PCTFREE 5 PCTUSED 85;

Create Table 病案封存记录(
    ID			Number(18),
    病人id		Number(18),
    主页id		Number(5),
    记录状态	Number(3),
    封存人		Varchar2(20),
    封存时间	Date,
    封存理由	Varchar2(255),
    解封人		Varchar2(20),
    解封时间	Date)
    TableSpace zl9CISAduit
    PCTFREE 5 PCTUSED 85;

--病案评分部分
Create Table 病案评分方案(
	ID number(18) not null,
	名称 varchar2(50),
	总分 number(8,2) default 100,
	上值 number(8,2),
	下值 number(8,2),
	类型 varchar2(10),
	分制 varchar2(10),
	选用 number(1) default 0,
	启用时间 Date,
	停用时间 Date)
    	TABLESPACE zl9CISAduit
    	PCTFREE 5  PCTUSED 85;

Create Table 病案评分标准(
	ID number(18) not null,
	上级ID number(18),
	方案ID number(18),
	名称 varchar2(50),
	描述 varchar2(4000),
	标准分值 number(8,2),
	缺陷等级 varchar2(2),
	评分单位 varchar2(8),
	上级序号 NUMBER(18),
	序号 NUMBER(18))
    	TABLESPACE zl9CISAduit
    	PCTFREE 5  PCTUSED 85;

Create Table 病案评分结果(
	ID number(18) not null,
	病人ID number(18),
	主页ID number(5),
	方案ID number(18),
	总分 number(8,2),
	等级 varchar2(2),
	返回修改 number(1),
	备注	varchar(50),
	评分人 varchar2(20),
	评分时间 Date,
	审核人 varchar2(20),
	审核时间 Date)
	TABLESPACE zl9CISAduit
	PCTFREE 10 PCTUSED 80;

Create Table 病案评分明细(
	ID number(18) not null,
	主表ID number(18),
	评分标准ID number(18),
	单项分数 number(8,2),
	缺陷等级 varchar2(2),
	可否修改 Number(1) Default 0,
	备注	varchar(50))
	TABLESPACE zl9CISAduit
	PCTFREE 10 PCTUSED 80;


--修改已有的表
Alter Table 病案主页 Add 病案状态 Number(3);
Alter Table 病案主页 Add 封存时间 Date;

--电子病案审查归档
Alter Table 病案提交记录 Add Constraint 病案提交记录_PK Primary Key (ID) Using Index Pctfree 0 Tablespace zl9indexcis;
Alter Table 病案提交记录 Add Constraint 病案提交记录_CK_记录状态 Check (记录状态 IN(1,2,3,4,5));
Alter Table 病案提交记录 Add Constraint 病案提交记录_FK_病人ID Foreign Key (病人ID) References 病人信息(病人ID);
Alter Table 病案提交记录 Add Constraint 病案提交记录_FK_主页ID Foreign Key (病人ID,主页ID) References 病案主页(病人ID,主页ID);
Alter Table 病案审阅书签 Add Constraint 病案审阅书签_PK Primary Key (提交ID,审阅对象,文件ID) Using Index Pctfree 0 Tablespace zl9indexcis;
Alter Table 病案审阅书签 Add Constraint 病案审阅书签_FK_提交ID Foreign Key (提交ID) References 病案提交记录(ID) On Delete Cascade;
Alter Table 病案审阅书签 Add Constraint 病案审阅书签_CK_审阅对象 Check (审阅对象 IN(1,2,3,4,5,6,7,8));
Alter Table 病案反馈记录 Add Constraint 病案反馈记录_PK Primary Key (ID) Using Index Pctfree 0 Tablespace zl9indexcis;
Alter Table 病案反馈记录 Add Constraint 病案反馈记录_FK_相关ID Foreign Key (相关ID) References 病案反馈记录(ID) On Delete Cascade;
Alter Table 病案反馈记录 Add Constraint 病案反馈记录_FK_提交ID Foreign Key (提交ID) References 病案提交记录(ID);
Alter Table 病案反馈记录 Add Constraint 病案反馈记录_CK_记录性质 Check (记录性质 IN(1,2));
Alter Table 病案反馈记录 Add Constraint 病案反馈记录_CK_记录状态 Check (记录状态 IN(1,2,3));
Alter Table 病案反馈记录 Add Constraint 病案反馈记录_CK_反馈对象 Check (反馈对象 IN(1,2,3,4,5,6,7,8));
Alter Table 病案反馈记录 Add Constraint 病案反馈记录_FK_病人ID Foreign Key (病人ID) References 病人信息(病人ID);
Alter Table 病案反馈记录 Add Constraint 病案反馈记录_FK_主页ID Foreign Key (病人ID,主页ID) References 病案主页(病人ID,主页ID);
Alter Table 病案反馈历史 Add Constraint 病案反馈历史_PK Primary Key (ID) Using Index Pctfree 0 Tablespace zl9indexcis;
Alter Table 病案反馈历史 Add Constraint 病案反馈历史_FK_相关ID Foreign Key (相关ID) References 病案反馈历史(ID) On Delete Cascade;
Alter Table 病案反馈历史 Add Constraint 病案反馈历史_FK_提交ID Foreign Key (提交ID) References 病案提交记录(ID);
Alter Table 病案反馈历史 Add Constraint 病案反馈历史_CK_记录性质 Check (记录性质 IN(1,2));
Alter Table 病案反馈历史 Add Constraint 病案反馈历史_CK_记录状态 Check (记录状态 IN(1,2,3));
Alter Table 病案反馈历史 Add Constraint 病案反馈历史_CK_反馈对象 Check (反馈对象 IN(1,2,3,4,5,6,7,8));
Alter Table 病案反馈历史 Add Constraint 病案反馈历史_FK_病人ID Foreign Key (病人ID) References 病人信息(病人ID);
Alter Table 病案反馈历史 Add Constraint 病案反馈历史_FK_主页ID Foreign Key (病人ID,主页ID) References 病案主页(病人ID,主页ID);

Alter Table 病案借阅记录 Add Constraint 病案借阅记录_PK Primary Key (ID) Using Index Pctfree 0 Tablespace zl9indexcis;
Alter Table 病案借阅记录 Add Constraint 病案借阅记录_CK_记录状态 Check (记录状态 IN(1,2,3));
Alter Table 病案借阅内容 Add Constraint 病案借阅内容_PK Primary Key (借阅ID,病人ID,主页ID) Using Index Pctfree 0 Tablespace zl9indexcis;
Alter Table 病案借阅内容 Add Constraint 病案借阅内容_FK_借阅ID Foreign Key (借阅ID) References 病案借阅记录(ID) On Delete Cascade;
Alter Table 病案借阅内容 Add Constraint 病案借阅内容_FK_病人ID Foreign Key (病人ID) References 病人信息(病人ID);
Alter Table 病案借阅内容 Add Constraint 病案借阅内容_FK_主页ID Foreign Key (病人ID,主页ID) References 病案主页(病人ID,主页ID);
Alter Table 病案借阅人员 Add Constraint 病案借阅人员_PK Primary Key (借阅ID,人员ID) Using Index Pctfree 0 Tablespace zl9indexcis;
Alter Table 病案借阅人员 Add Constraint 病案借阅人员_FK_借阅ID Foreign Key (借阅ID) References 病案借阅记录(ID) On Delete Cascade;
Alter Table 病案封存记录 Add Constraint 病案封存记录_PK Primary Key (ID) Using Index Pctfree 0 Tablespace zl9indexcis;
Alter Table 病案封存记录 Add Constraint 病案封存记录_CK_记录状态 Check (记录状态 IN(1,2));
Alter Table 病案封存记录 Add Constraint 病案封存记录_FK_病人ID Foreign Key (病人ID) References 病人信息(病人ID);
Alter Table 病案封存记录 Add Constraint 病案封存记录_FK_主页ID Foreign Key (病人ID,主页ID) References 病案主页(病人ID,主页ID);

--病案评分部分
ALTER TABLE 病案评分方案 ADD CONSTRAINT 病案评分方案_PK PRIMARY KEY (ID) USING INDEX PCTFREE 5 TABLESPACE zl9indexcis;
ALTER TABLE 病案评分方案 Add CONSTRAINT 病案评分方案_CK_选用 CHECK (选用 IN(0,1));
ALTER TABLE 病案评分标准 ADD CONSTRAINT 病案评分标准_PK PRIMARY KEY (ID) USING INDEX PCTFREE 5 TABLESPACE zl9indexcis;
ALTER TABLE 病案评分标准 ADD CONSTRAINT 病案评分标准_FK_上级ID FOREIGN KEY (上级ID) REFERENCES 病案评分标准(ID) ON DELETE CASCADE;
ALTER TABLE 病案评分标准 ADD CONSTRAINT 病案评分标准_FK_方案ID FOREIGN KEY (方案ID) REFERENCES 病案评分方案(ID) ON DELETE CASCADE;
ALTER TABLE 病案评分结果 ADD CONSTRAINT 病案评分结果_PK PRIMARY KEY (ID) USING INDEX PCTFREE 5 TABLESPACE zl9indexcis;
ALTER TABLE 病案评分结果 Add CONSTRAINT 病案评分结果_UQ_病人ID_主页ID UNIQUE (病人ID,主页ID) USING INDEX PCTFREE 5 TABLESPACE zl9indexcis;
ALTER TABLE 病案评分结果 ADD CONSTRAINT 病案评分结果_FK_病人ID_主页ID FOREIGN KEY (病人ID,主页ID) REFERENCES 病案主页(病人ID,主页ID) ON DELETE CASCADE;
ALTER TABLE 病案评分结果 ADD CONSTRAINT 病案评分结果_FK_方案ID FOREIGN KEY (方案ID) REFERENCES 病案评分方案(ID) ON DELETE CASCADE;
ALTER TABLE 病案评分明细 ADD CONSTRAINT 病案评分明细_PK PRIMARY KEY (ID) USING INDEX PCTFREE 5 TABLESPACE zl9indexcis;
ALTER TABLE 病案评分明细 ADD CONSTRAINT 病案评分明细_FK_评分标准ID FOREIGN KEY (评分标准ID) REFERENCES 病案评分标准(ID) ON DELETE CASCADE;
ALTER TABLE 病案评分明细 ADD CONSTRAINT 病案评分明细_FK_主表ID FOREIGN KEY (主表ID) REFERENCES 病案评分结果(ID) ON DELETE CASCADE;
ALTER TABLE 病案评分明细 Add CONSTRAINT 病案评分明细_CK_可否修改 CHECK (可否修改 IN(0,1));

-----------------------------------------------------------------------------------------------------------------------
---电子病案审查归档
-----------------------------------------------------------------------------------------------------------------------
Create Index 病案提交记录_IX_病人id On 病案提交记录(病人id) Pctfree 5 Tablespace zl9indexcis
/
Create Index 病案审阅书签_IX_提交id On 病案审阅书签(提交id) Pctfree 5 Tablespace zl9indexcis
/
Create Index 病案反馈记录_IX_提交id On 病案反馈记录(提交id) Pctfree 5 Tablespace zl9indexcis
/
Create Index 病案反馈记录_IX_相关id On 病案反馈记录(相关id) Pctfree 5 Tablespace zl9indexcis
/
Create Index 病案反馈记录_IX_反馈时间 On 病案反馈记录(反馈时间) Pctfree 5 Tablespace zl9indexcis
/
Create Index 病案反馈记录_IX_处理时间 On 病案反馈记录(处理时间) Pctfree 5 Tablespace zl9indexcis
/
Create Index 病案反馈历史_IX_提交id On 病案反馈历史(提交id) Pctfree 5 Tablespace zl9indexcis
/
Create Index 病案反馈历史_IX_相关id On 病案反馈历史(相关id) Pctfree 5 Tablespace zl9indexcis
/
Create Index 病案反馈历史_IX_反馈时间 On 病案反馈历史(反馈时间) Pctfree 5 Tablespace zl9indexcis
/
Create Index 病案反馈历史_IX_处理时间 On 病案反馈历史(处理时间) Pctfree 5 Tablespace zl9indexcis
/
Create Index 病案借阅内容_IX_借阅id On 病案借阅内容(借阅id) Pctfree 5 Tablespace zl9indexcis
/
Create Index 病案借阅内容_IX_病人id On 病案借阅内容(病人id) Pctfree 5 Tablespace zl9indexcis
/
Create Index 病案借阅人员_IX_借阅id On 病案借阅人员(借阅id) Pctfree 5 Tablespace zl9indexcis
/

--病案评分部分
Create Index 病案评分标准_IX_方案ID on 病案评分标准(方案ID) PCTFREE 5 TABLESPACE zl9indexcis
/
Create Index 病案评分标准_IX_上级ID on 病案评分标准(上级ID) PCTFREE 5 TABLESPACE zl9indexcis
/
Create Index 病案评分结果_IX_方案ID on 病案评分结果(方案ID) PCTFREE 5 TABLESPACE zl9indexcis
/
Create Index 病案评分明细_IX_结果ID on 病案评分明细(主表ID) PCTFREE 5 TABLESPACE zl9indexcis
/
Create Index 病案评分明细_IX_评分标准ID on 病案评分明细(评分标准ID) PCTFREE 5 TABLESPACE zl9indexcis
/


--视图数据
create or replace view 病案评分标准视图 as
select decode(T.上级序号,null,序号,T.上级序号) as 上级序号, decode(T.序号,null,T.ID,T.序号) as 序号,T.ID,T.上级ID,T.方案ID,T.项目,T.标准分值,T.基本要求,T.缺陷内容,T.扣分标准,decode(T.子项个数,0,'否','是') as 隐藏
from
(
  select B.上级序号,A.序号,A.方案ID,
  A.ID,
  A.上级ID,
  decode(A.子项个数,0,decode(A.上级ID,Null,A.名称,B.名称),A.名称) as 项目,
  decode(A.子项个数,0,decode(A.上级ID,Null,A.标准分值,B.标准分值),B.标准分值) as 标准分值,
  decode(A.子项个数,0,decode(A.上级ID,Null,A.描述,B.描述),A.描述) as 基本要求,
  A.描述 as 缺陷内容,
  DECODE(A.缺陷等级,NULL,decode(sign(A.标准分值-1),-1,To_CHAR(A.标准分值,'0.9'),To_Char(A.标准分值))||decode(A.评分单位,NULL,'','/'||A.评分单位),A.缺陷等级) as 扣分标准,
  A.子项个数
  from
      (
          select AA.序号,AA.ID,AA.方案ID,AA.上级ID,AA.名称,AA.描述,AA.标准分值,AA.缺陷等级,AA.评分单位,count(BB.ID) as 子项个数
          from 病案评分标准 AA,病案评分标准 BB
          where AA.ID=BB.上级ID(+)
          group by AA.序号,AA.ID,AA.方案ID,AA.上级ID,AA.名称,AA.描述,AA.标准分值,AA.缺陷等级,AA.评分单位
      ) A,
      (
          select 序号 as 上级序号,ID,名称,标准分值,描述 from 病案评分标准
      ) B
  where A.上级ID=B.ID(+)
) T
order by decode(T.上级序号,null,序号,T.上级序号),decode(T.序号,null,T.ID,T.序号);


create or replace view 病案质量报表视图 as
Select Tb.住院号, Tb.姓名, Tb.性别, Ta.*
From (Select T1.病人id, T1.主页id, T1.入院日期, T1.出院日期, T2.名称 As 入院科室, T3.名称 As 出院科室, T1.门诊医师,
              T1.责任护士, T1.住院医师, T1.编目日期, T1.结果id, T1.方案id, T1.总分, T1.等级, T1.评分人,
              To_Char(T1.评分时间, 'YYYY-MM-DD') As 评分时间, T1.审核人, To_Char(T1.审核时间, 'YYYY-MM-DD') As 审核时间,
              T1.返回修改, T1.备注
       From (Select A.病人id, A.主页id, A.入院科室id, A.出院科室id, A.入院日期, A.出院日期, A.门诊医师, A.责任护士,
                     A.住院医师, A.编目日期, B.ID As 结果id, B.方案id, B.总分, B.等级, B.评分人, B.评分时间, B.审核人,
                     B.审核时间, B.返回修改, B.备注
              From 病案主页 A, 病案评分结果 B
              Where A.病人id = B.病人id(+) And A.主页id = B.主页id(+)) T1, 部门表 T2, 部门表 T3
       Where T1.入院科室id = T2.ID And T1.出院科室id = T3.ID) Ta, 病人信息 Tb
Where Ta.病人id = Tb.病人id;



--zlStreamTabs数据
-----------------------------------------------------------------------------------------------------------------------



--zlBakTables数据
-----------------------------------------------------------------------------------------------------------------------


--zlComponent数据
-----------------------------------------------------------------------------------------------------------------------
Insert Into zlComponent(部件,名称,主版本,次版本,附版本,系统) Values('zl9CISAduit','电子病案审查归档',10,20,0,100);

--zlPrograms数据
-----------------------------------------------------------------------------------------------------------------------
Insert into zlPrograms(序号,标题,说明,系统,部件) Values(1550,'评分标准维护','完成病案评分标准增删改和选用。',100,'zl9CISAduit');
Insert Into zlPrograms(序号,标题,说明,系统,部件) Values(1560,'电子病案审查','出院病人的电子病案的审查和归档以及在院病人的电子病案的抽查及封存。',100,'zl9CISAduit');
Insert Into zlPrograms(序号,标题,说明,系统,部件) Values(1561,'电子病案借阅','归档病人的电子病案查阅的申请和查阅。',100,'zl9CISAduit');
Insert into zlPrograms(序号,标题,说明,系统,部件) Values(1562,'电子病案评分','按照录入的评分标准对已有病案进行评分和审核。',100,'zl9CISAduit');


--zlProgFuncs数据
-----------------------------------------------------------------------------------------------------------------------
--评分标准维护
Insert Into zlProgFuncs(系统,序号,功能,说明)  values (100,1550,'基本','');
Insert Into zlProgFuncs(系统,序号,功能,说明)  values (100,1550,'增删改','有此权限的用户可以进行病案评分标准的增删改及选用操作。');

--电子病案审查
Insert Into zlProgFuncs(系统,序号,功能,说明)  values (100,1560,'基本','');
Insert Into zlProgFuncs(系统,序号,功能,说明)  values (100,1560,'参数设置','设置本模块相关的全局参数');
Insert Into zlProgFuncs(系统,序号,功能,说明)  values (100,1560,'审查接收','开始审查病人的电子病案');
Insert Into zlProgFuncs(系统,序号,功能,说明)  values (100,1560,'拒绝审查','拒绝审查病人的电子病案');
Insert Into zlProgFuncs(系统,序号,功能,说明)  values (100,1560,'审查病案','对病人的电子病案进行审查处理');
Insert Into zlProgFuncs(系统,序号,功能,说明)  values (100,1560,'归档病案','对病人的电子病案进行归档处理');
Insert Into zlProgFuncs(系统,序号,功能,说明)  values (100,1560,'封存病案','对在院病人的电子病案进行封存处理');
Insert Into zlProgFuncs(系统,序号,功能,说明)  values (100,1560,'解封病案','对在院病人的电子病案进行解封处理');
Insert Into zlProgFuncs(系统,序号,功能,说明)  values (100,1560,'回退接收','回退已经开始审查的病人电子病案');
Insert Into zlProgFuncs(系统,序号,功能,说明)  values (100,1560,'回退拒绝','回退被拒绝的病人电子病案');
Insert Into zlProgFuncs(系统,序号,功能,说明)  values (100,1560,'回退归档','回退已经归档的病人电子病案');

--电子病案借阅
Insert Into zlProgFuncs(系统,序号,功能,说明)  values (100,1561,'基本','');
Insert Into zlProgFuncs(系统,序号,功能,说明)  values (100,1561,'参数设置','设置本模块相关的全局参数');
Insert Into zlProgFuncs(系统,序号,功能,说明)  values (100,1561,'登记申请','登记、修改和删除电子病案的借阅申请单');
Insert Into zlProgFuncs(系统,序号,功能,说明)  values (100,1561,'审批申请','批准或拒绝新登记的借阅申请单');
Insert Into zlProgFuncs(系统,序号,功能,说明)  values (100,1561,'查阅病案','查阅已经批准的病人的电子病案内容');
--病案评分审核
Insert Into zlProgFuncs(系统,序号,功能,说明)  values (100,1562,'基本','');
Insert Into zlProgFuncs(系统,序号,功能,说明)  values (100,1562,'评分','有此权限的用户可以进行病案评分操作。');
Insert Into zlProgFuncs(系统,序号,功能,说明)  values (100,1562,'审核','有此权限的用户可以进行病案评分结果的审核操作。');
Insert Into zlProgFuncs(系统,序号,功能,说明)  values (100,1562,'取消审核','有此权限的用户可以对已经审核病案进行取消审核的操作。');
Insert Into zlProgFuncs(系统,序号,功能,说明)  values (100,1562,'所有科室','有此权限的用户可以对所有科室进行评分，否则只能对本科室病案进行评分。');
Insert Into zlProgFuncs(系统,序号,功能,说明)  values (100,1562,'修改他人评分','有此权限的用户可以修改他人评分结果，否则只能修改本人评分结果。');


--zlProgPrivs数据
-----------------------------------------------------------------------------------------------------------------------
--评分标准维护
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1550,'基本',user,'病案评分方案','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1550,'基本',user,'病案评分结果','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1550,'基本',user,'病案评分方案_ID','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1550,'基本',user,'病案评分标准','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1550,'基本',user,'病案评分标准_ID','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1550,'基本',user,'病案评分标准视图','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1550,'增删改',user,'ZL_病案评分方案_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1550,'增删改',user,'ZL_病案评分方案_Update','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1550,'增删改',user,'ZL_病案评分方案_Delete','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1550,'增删改',user,'ZL_病案评分标准_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1550,'增删改',user,'ZL_病案评分标准_Update','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1550,'增删改',user,'ZL_病案评分标准_Delete','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1550,'增删改',user,'ZL_病案评分方案_选用','EXECUTE');

--电子病案审查
--基本
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'病案提交记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'病案反馈记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'病案反馈历史','SELECT');

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'病人信息','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'病案主页','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'病案主页从表','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'诊断符合情况','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'病人变动记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'床位状况记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'病人过敏记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'病人诊断记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'病人手麻记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'病人新生儿记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'病区科室对应','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'病人医嘱状态','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'病人医嘱发送','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'病人费用记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'药品特性','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'药品规格','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'诊疗项目目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'病历单据应用','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'电子病历附件','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'Zl_Lob_Read','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'电子病历内容','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'隐私保护项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'诊治所见项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'病历文件结构','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'病人护理内容','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'护理记录项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'体温记录项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'电子病历记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'病人医嘱报告','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'病人医嘱记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'病历文件列表','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'病历页面格式','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'病历应用科室','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'基本',user,'病人护理记录','SELECT');
--审查接收
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'审查接收',user,'zl_病案提交记录_Receive','EXECUTE');
--拒绝审查
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'拒绝审查',user,'zl_病案提交记录_Refuse','EXECUTE');
--归档病案
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'归档病案',user,'zl_病案提交记录_Archive','EXECUTE');
--回退接收
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'回退接收',user,'zl_病案提交记录_UnReceive','EXECUTE');
--回退拒绝
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'回退拒绝',user,'zl_病案提交记录_UnRefuse','EXECUTE');
--回退归档
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'回退归档',user,'zl_病案提交记录_UnArchive','EXECUTE');
--封存病案
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'封存病案',user,'zl_病案封存记录_Lock','EXECUTE');
--解封病案
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'解封病案',user,'zl_病案封存记录_UnLock','EXECUTE');
--审查病案
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'审查病案',user,'病案评分方案','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'审查病案',user,'病案评分标准','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'审查病案',user,'病案反馈记录_ID','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'审查病案',user,'病历时限要求','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'审查病案',user,'病人挂号记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'审查病案',user,'病历书写事件','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'审查病案',user,'病历内容监测','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'审查病案',user,'病历时限监测','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'审查病案',user,'zl_病案反馈记录_Update','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'审查病案',user,'zl_病案反馈记录_Delete','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'审查病案',user,'zl_病案反馈记录_Finish','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'审查病案',user,'zl_病案反馈记录_RollBackFinish','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'审查病案',user,'Zl_病历内容监测_Neaten','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1560,'审查病案',user,'Zl_病历时限监测_Neaten','EXECUTE');

--电子病案借阅
--基本
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'病人信息','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'病案主页','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'病案主页从表','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'诊断符合情况','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'病人变动记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'床位状况记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'病人过敏记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'病人诊断记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'病人手麻记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'病人新生儿记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'病区科室对应','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'病人医嘱状态','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'病人医嘱发送','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'病人费用记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'药品特性','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'药品规格','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'诊疗项目目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'病历单据应用','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'电子病历附件','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'Zl_Lob_Read','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'电子病历内容','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'隐私保护项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'诊治所见项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'病历文件结构','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'病人护理内容','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'护理记录项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'体温记录项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'电子病历记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'病人医嘱报告','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'病人医嘱记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'病历文件列表','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'病历页面格式','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'病历应用科室','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'病人护理记录','SELECT');

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'病案借阅记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'病案借阅人员','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'基本',user,'病案借阅内容','SELECT');
--登记申请
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'登记申请',user,'病案借阅记录_ID','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'登记申请',user,'性别','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'登记申请',user,'婚姻状况','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'登记申请',user,'疾病编码分类','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'登记申请',user,'疾病编码目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'登记申请',user,'zl_病案借阅人员_Update','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'登记申请',user,'zl_病案借阅内容_Update','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'登记申请',user,'zl_病案借阅记录_Update','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'登记申请',user,'zl_病案借阅记录_Delete','EXECUTE');
--审批申请
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'审批申请',user,'zl_病案借阅记录_Authorize','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'审批申请',user,'zl_病案借阅记录_Refuse','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1561,'审批申请',user,'zl_病案借阅记录_Rollback','EXECUTE');

--病案评分审核
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'基本',user,'病案评分结果','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'基本',user,'病案评分明细','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'基本',user,'病案评分标准视图','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'基本',user,'病案质量报表视图','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'基本',user,'病案评分方案','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'基本',user,'病案评分标准','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'基本',user,'病案主页','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'基本',user,'病案主页从表','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'基本',user,'病人过敏药物','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'基本',user,'病人信息','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'基本',user,'部门表','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'基本',user,'部门人员','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'基本',user,'部门性质说明','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'基本',user,'疾病编码目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'基本',user,'人员表','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'基本',user,'上机人员表','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'基本',user,'病人手麻记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'基本',user,'诊断符合情况','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'基本',user,'病人诊断记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'基本',user,'住院病案记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'基本',user,'临床部门','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'基本',user,'婚姻状况','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'基本',user,'职业','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'基本',user,'病情','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'基本',user,'血型','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'基本',user,'医疗付款方式','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'基本',user,'疾病编码分类','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'基本',user,'病案费目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'基本',user,'地区','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'评分',user,'病案评分结果_ID','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'评分',user,'病案评分明细_ID','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'评分',user,'ZL_病案评分结果_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'评分',user,'ZL_病案评分结果_Update','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'评分',user,'ZL_病案评分结果_Delete','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'评分',user,'ZL_病案评分明细_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'审核',user,'ZL_病案评分结果_审核','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'审核',user,'病案主页从表','INSERT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1562,'取消审核',user,'ZL_病案评分结果_取消审核','EXECUTE');

--zlMenus数据
-----------------------------------------------------------------------------------------------------------------------
Insert Into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) Values('缺省',zlMenus_id.nextval,null,'电子病案管理','病案管理','A',99,'电子病案的审查归档处理',100,NULL);
Insert Into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) Values('缺省',zlMenus_id.nextval,   zlMenus_id.nextval-1,'评分标准维护','评分标准','A',231,'完成病案评分标准增删改和选用。',100,1550);
Insert Into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) Values('缺省',zlMenus_id.nextval,   zlMenus_id.nextval-2,'电子病案审查','病案审查','B',232,'出院病人的电子病案的审查和归档以及在院病人的电子病案的抽查及封存。',100,1560);
Insert Into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) Values('缺省',zlMenus_id.nextval,   zlMenus_id.nextval-3,'电子病案借阅','病案借阅','C',141,'归档病人的电子病案查阅的申请和查阅。',100,1561);
Insert Into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) Values('缺省',zlMenus_id.nextval,   zlMenus_id.nextval-4,'电子病案评分','病案评分','D',136,'按照录入的评分标准对已有病案进行评分和审核。',100,1562);
Insert Into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) Values('缺省',zlMenus_id.nextval,   zlMenus_id.nextval-5,'病案报表分析','报表分析','E',99,'电子病案审查归档的相关报表分析',100,NULL);

--zlBaseCode数据
-----------------------------------------------------------------------------------------------------------------------

--zlParameters数据
--1560:电子病案审查
Insert Into zlParameters(ID,系统,模块,私有,参数号,参数名,参数值,缺省值,参数说明)
Select Rownum+B.ID,A.* From (
	Select 系统,模块,私有,参数号,参数名,参数值,缺省值,参数说明 From zlParameters Where ID=0 Union ALL
	Select 100,1560,1,1,'定位依据','姓名','姓名','查找定位病人的方式' From Dual Union ALL
	Select 100,1560,1,2,'上次状态','0','0','操作员上次选择的病人列表序号' From Dual Union ALL
	Select 100,1560,1,3,'审查缺省范围','今  天','今  天','显示未归档病人的缺省时间范围' From Dual Union ALL
	Select 100,1560,1,4,'归档缺省范围','今  天','今  天','显示已归档病人的缺省时间范围' From Dual Union ALL
	Select 100,1560,0,5,'反馈处理期限','7','7','反馈的问题要求临床科室处理的最晚天数' From Dual Union ALL
	Select 100,1560,0,6,'未复查刷新频率','5','5','自动刷新未复查反馈问题的时间间隔，单位：分钟' From Dual Union ALL
	Select 100,1560,1,7,'完成问题范围','今  天','今  天','显示已完成的问题的时间范围' From Dual
  ) A,(Select Nvl(Max(ID),0) AS ID From zlParameters) B;

--1561:电子病案借阅
Insert Into zlParameters(ID,系统,模块,私有,参数号,参数名,参数值,缺省值,参数说明)
Select Rownum+B.ID,A.* From (
	Select 系统,模块,私有,参数号,参数名,参数值,缺省值,参数说明 From zlParameters Where ID=0 Union ALL
	Select 100,1561,1,1,'定位依据','No','No','查找定位借阅申请的方式' From Dual Union ALL
	Select 100,1561,1,2,'上次状态','0','0','操作员上次选择的列表序号' From Dual Union ALL
	Select 100,1561,1,3,'登记缺省范围','今  天','今  天','显示登记申请的缺省时间范围' From Dual Union ALL
	Select 100,1561,0,4,'病案借阅期限','7','7','病案借阅的最晚天数' From Dual
  ) A,(Select Nvl(Max(ID),0) AS ID From zlParameters) B;

--最后调整zlParameters的序列
Select zlParameters_ID.Nextval From zlParameters;

--号码控制表数据
Insert Into 号码控制表(项目序号,项目名称,最大号码,自动补缺,编号规则)
Select 86,'病案借阅申请单','',1,0 From Dual;

------------------------------------------------------------------------------------------------------------------------------------------
--		过程清单
------------------------------------------------------------------------------------------------------------------------------------------
--病案评分方案
CREATE OR REPLACE PROCEDURE ZL_病案评分方案_Insert
(	ID_in IN 病案评分方案.ID%TYPE,
	名称_in IN 病案评分方案.名称%TYPE,
	总分_in IN 病案评分方案.总分%TYPE,
	上值_in IN 病案评分方案.上值%TYPE,
	下值_in IN 病案评分方案.下值%TYPE,
	类型_in IN 病案评分方案.类型%TYPE,
	分制_in IN 病案评分方案.分制%TYPE,
	选用_in IN 病案评分方案.选用%TYPE,
	启用时间_in IN 病案评分方案.启用时间%TYPE,
	停用时间_in IN 病案评分方案.停用时间%TYPE
)
IS
BEGIN
  if 选用_in=1 then
     update 病案评分方案 
     set 选用=0
     where 类型=类型_in;  
  end if;
  
	INSERT INTO 病案评分方案
		(ID,名称,总分,上值,下值,类型,分制,选用,启用时间,停用时间)
	VALUES
		(ID_in,名称_in,总分_in,上值_in,下值_in,类型_in,分制_in,选用_in,启用时间_in,停用时间_in);
    
EXCEPTION
	WHEN OTHERS THEN
		ZL_ErrorCenter (SQLCODE, SQLERRM);
END  ZL_病案评分方案_Insert;
/

CREATE OR REPLACE PROCEDURE ZL_病案评分方案_Update
(	ID_in IN 病案评分方案.ID%TYPE,
	名称_in IN 病案评分方案.名称%TYPE,
	总分_in IN 病案评分方案.总分%TYPE,
	上值_in IN 病案评分方案.上值%TYPE,
	下值_in IN 病案评分方案.下值%TYPE,
	类型_in IN 病案评分方案.类型%TYPE,
	分制_in IN 病案评分方案.分制%TYPE,
	选用_in IN 病案评分方案.选用%TYPE,
	启用时间_in IN 病案评分方案.启用时间%TYPE,
	停用时间_in IN 病案评分方案.停用时间%TYPE
)
IS
BEGIN
  if 选用_in=1 then
     update 病案评分方案 
     set 选用=0
     where 类型=类型_in;  
  end if;

	Update 病案评分方案
	set	名称=名称_in,总分=总分_in,上值=上值_in,下值=下值_in,类型=类型_in,分制=分制_in,选用=选用_in,启用时间=启用时间_in,停用时间=停用时间_in
	where ID=ID_in;
	
	IF SQL%NOTFOUND THEN
		---如果没有更新到，那么新增一条
		INSERT INTO 病案评分方案
			(ID,名称,总分,上值,下值,类型,分制,选用,启用时间,停用时间)
		VALUES
			(ID_in,名称_in,总分_in,上值_in,下值_in,类型_in,分制_in,选用_in,启用时间_in,停用时间_in);
	END IF;

EXCEPTION
	WHEN OTHERS THEN
		ZL_ErrorCenter (SQLCODE, SQLERRM);
END  ZL_病案评分方案_Update;
/

CREATE OR REPLACE PROCEDURE ZL_病案评分方案_Delete
(
	ID_in IN 病案评分方案.ID%TYPE
)
IS
BEGIN
	DELETE FROM 病案评分方案
		WHERE ID = ID_in;
EXCEPTION
	WHEN OTHERS THEN
		ZL_ErrorCenter (SQLCODE, SQLERRM);
END  ZL_病案评分方案_Delete;
/

CREATE OR REPLACE PROCEDURE ZL_病案评分方案_选用
(	ID_in IN 病案评分方案.ID%TYPE,
  选用_in IN 病案评分方案.选用%TYPE
)
IS
BEGIN
  if 选用_in=1 then
     update 病案评分方案
     set 选用=0
     where 类型=(select 类型 from 病案评分方案 where ID=ID_in);
  end if;

	update 病案评分方案
	set 选用=选用_in
  where ID=ID_in; 

EXCEPTION
	WHEN OTHERS THEN
		ZL_ErrorCenter (SQLCODE, SQLERRM);
END  ZL_病案评分方案_选用;
/

--病案评分标准

CREATE OR REPLACE PROCEDURE ZL_病案评分标准_Insert
(	ID_in IN 病案评分标准.ID%TYPE,
	上级ID_in IN 病案评分标准.上级ID%TYPE,
	方案ID_in IN 病案评分标准.方案ID%TYPE,
	名称_in IN 病案评分标准.名称%TYPE,
	描述_in IN 病案评分标准.描述%TYPE,
	标准分值_in IN 病案评分标准.标准分值%TYPE,
	缺陷等级_in IN 病案评分标准.缺陷等级%TYPE,
	评分单位_in IN 病案评分标准.评分单位%TYPE,
  基准ID_IN IN 病案评分标准.序号%TYPE
)
IS
  基准序号 NUMBER;
BEGIN
  if 基准ID_IN=0 or 基准ID_IN is null then
     if 上级ID_IN=0 or 上级ID_IN is null then
        select decode(max(序号),null,0,max(序号)+1) into 基准序号 from 病案评分标准 where 方案ID=方案ID_IN and 上级ID is null;
     else
        select decode(max(序号),null,0,max(序号)+1) into 基准序号 from 病案评分标准 where 方案ID=方案ID_IN and 上级ID=上级ID_IN;
     end if;
     --添加评分标准
     INSERT INTO 病案评分标准
       (ID,上级ID,方案ID,名称,描述,标准分值,缺陷等级,评分单位,序号)
     VALUES
       (ID_IN,上级ID_IN,方案ID_IN,名称_IN,描述_IN,标准分值_IN,缺陷等级_IN,评分单位_IN,基准序号);       
  else
     --插入评分标准
    if 上级ID_IN=0 or 上级ID_IN is null then   --为评分项目
       select 序号 into 基准序号 from 病案评分标准 where ID=基准ID_IN;
       update 病案评分标准 set 序号=序号+1 where 上级ID is null and 序号>=基准序号 and 方案ID=方案ID_IN;
       INSERT INTO 病案评分标准
      	 (ID,上级ID,方案ID,名称,描述,标准分值,缺陷等级,评分单位,序号)
       VALUES
      	 (ID_IN,上级ID_IN,方案ID_IN,名称_IN,描述_IN,标准分值_IN,缺陷等级_IN,评分单位_IN,基准序号);     
    else                        --为评分标准
       select 序号 into 基准序号 from 病案评分标准 where ID=基准ID_IN;
       update 病案评分标准 set 序号=序号+1 where 上级ID=上级ID_in  and 序号>=基准序号 and 方案ID=方案ID_IN;
       INSERT INTO 病案评分标准
      	 (ID,上级ID,方案ID,名称,描述,标准分值,缺陷等级,评分单位,序号)
       VALUES
      	 (ID_IN,上级ID_IN,方案ID_IN,名称_IN,描述_IN,标准分值_IN,缺陷等级_IN,评分单位_IN,基准序号);     
    end if;
  end if;
EXCEPTION
	WHEN OTHERS THEN
		ZL_ErrorCenter (SQLCODE, SQLERRM);
END  ZL_病案评分标准_Insert;
/

CREATE OR REPLACE PROCEDURE ZL_病案评分标准_Update
(	ID_in IN 病案评分标准.ID%TYPE,
	上级ID_in IN 病案评分标准.上级ID%TYPE,
	方案ID_in IN 病案评分标准.方案ID%TYPE,
	名称_in IN 病案评分标准.名称%TYPE,
	描述_in IN 病案评分标准.描述%TYPE,
	标准分值_in IN 病案评分标准.标准分值%TYPE,
	缺陷等级_in IN 病案评分标准.缺陷等级%TYPE,
	评分单位_in IN 病案评分标准.评分单位%TYPE
)
IS
BEGIN
	Update 病案评分标准
	set 上级ID=上级ID_IN,方案ID=方案ID_IN,名称=名称_IN,描述=描述_IN,标准分值=标准分值_IN,缺陷等级=缺陷等级_IN,评分单位=评分单位_IN
	where ID=ID_IN;
	
	IF SQL%NOTFOUND THEN
		---如果没有更新到，那么新增一条
		INSERT INTO 病案评分标准
			(ID,上级ID,方案ID,名称,描述,标准分值,缺陷等级,评分单位)
		VALUES
			(ID_IN,上级ID_IN,方案ID_IN,名称_IN,描述_IN,标准分值_IN,缺陷等级_IN,评分单位_IN);
	END IF;

EXCEPTION
	WHEN OTHERS THEN
		ZL_ErrorCenter (SQLCODE, SQLERRM);
END  ZL_病案评分标准_Update;
/

CREATE OR REPLACE PROCEDURE ZL_病案评分标准_Delete
(
	ID_in IN 病案评分标准.ID%TYPE,
  删除独立项目_in IN NUMBER
)
IS
  lng上级ID NUMBER;
  lng独立评分项 NUMBER;
BEGIN
  if 删除独立项目_in=1 then
    select decode(上级ID,null,0,上级ID) into lng上级ID
      from 病案评分标准 where ID=ID_IN;
  end if;
  
	DELETE FROM 病案评分标准
		WHERE ID = ID_in;
    
  if 删除独立项目_in=1 then
 
    select decode(隐藏,'否',1,0) into lng独立评分项
      from 病案评分标准视图 where ID=lng上级ID;
    if lng独立评分项=1 then
       DELETE FROM 病案评分标准	WHERE ID = lng上级ID;
    end if;
  end if;

EXCEPTION
	WHEN OTHERS THEN
		ZL_ErrorCenter (SQLCODE, SQLERRM);
END  ZL_病案评分标准_Delete;
/


--病案评分结果
Create Or Replace Procedure Zl_病案评分结果_Insert
(
  Id_In       In 病案评分结果.ID%Type,
  病人id_In   In 病案评分结果.病人id%Type,
  主页id_In   In 病案评分结果.主页id%Type,
  方案id_In   In 病案评分结果.方案id%Type,
  总分_In     In 病案评分结果.总分%Type,
  等级_In     In 病案评分结果.等级%Type,
  备注_In     In 病案评分结果.备注%Type,
  评分人_In   In 病案评分结果.评分人%Type,
  评分时间_In In 病案评分结果.评分时间%Type,
  审核人_In   In 病案评分结果.审核人%Type,
  审核时间_In In 病案评分结果.审核时间%Type,
  返回修改_In In 病案评分结果.返回修改%Type
) Is
Begin
  Insert Into 病案评分结果
    (ID, 病人id, 主页id, 方案id, 总分, 等级, 备注, 评分人, 评分时间, 审核人, 审核时间, 返回修改)
  Values
    (Id_In, 病人id_In, 主页id_In, 方案id_In, 总分_In, 等级_In, 备注_In, 评分人_In, 评分时间_In, 审核人_In, 审核时间_In,
     返回修改_In);
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病案评分结果_Insert;
/

Create Or Replace Procedure Zl_病案评分结果_Update
(
  Id_In       In 病案评分结果.ID%Type,
  病人id_In   In 病案评分结果.病人id%Type,
  主页id_In   In 病案评分结果.主页id%Type,
  方案id_In   In 病案评分结果.方案id%Type,
  总分_In     In 病案评分结果.总分%Type,
  等级_In     In 病案评分结果.等级%Type,
  备注_In     In 病案评分结果.备注%Type,
  评分人_In   In 病案评分结果.评分人%Type,
  评分时间_In In 病案评分结果.评分时间%Type,
  审核人_In   In 病案评分结果.审核人%Type,
  审核时间_In In 病案评分结果.审核时间%Type,
  返回修改_In In 病案评分结果.返回修改%Type
) Is
Begin
  Update 病案评分结果
  Set 病人id = 病人id_In, 主页id = 主页id_In, 方案id = 方案id_In, 总分 = 总分_In, 等级 = 等级_In, 备注 = 备注_In,
      评分人 = 评分人_In, 评分时间 = 评分时间_In, 审核人 = 审核人_In, 审核时间 = 审核时间_In, 返回修改 = 返回修改_In
  Where ID = Id_In;

  If Sql%NotFound Then
    ---如果没有更新到，那么新增一条
    Insert Into 病案评分结果
      (ID, 病人id, 主页id, 方案id, 总分, 等级, 备注, 评分人, 评分时间, 审核人, 审核时间, 返回修改)
    Values
      (Id_In, 病人id_In, 主页id_In, 方案id_In, 总分_In, 等级_In, 备注_In, 评分人_In, 评分时间_In, 审核人_In,
       审核时间_In, 返回修改_In);
  End If;

Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病案评分结果_Update;
/


CREATE OR REPLACE PROCEDURE ZL_病案评分结果_Delete
(
	ID_in IN 病案评分结果.ID%TYPE
)
IS
BEGIN
	DELETE FROM 病案评分结果
		WHERE ID = ID_in;
EXCEPTION
	WHEN OTHERS THEN
		ZL_ErrorCenter (SQLCODE, SQLERRM);
END  ZL_病案评分结果_Delete;
/

  --病案评分明细

Create Or Replace Procedure Zl_病案评分明细_Insert
(
  Id_In         In 病案评分明细.ID%Type,
  主表id_In     In 病案评分明细.主表id%Type,
  评分标准id_In In 病案评分明细.评分标准id%Type,
  单项分数_In   In 病案评分明细.单项分数%Type,
  缺陷等级_In   In 病案评分明细.缺陷等级%Type,
  可否修改_In   In 病案评分明细.可否修改%Type,
  备注_In       In 病案评分结果.备注%Type
  
) Is
Begin
  Insert Into 病案评分明细
    (ID, 主表id, 评分标准id, 单项分数, 缺陷等级, 可否修改, 备注)
  Values
    (Id_In, 主表id_In, 评分标准id_In, 单项分数_In, 缺陷等级_In, 可否修改_In, 备注_In);
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病案评分明细_Insert;
/

Create Or Replace Procedure Zl_病案评分明细_Update
(
  Id_In         In 病案评分明细.ID%Type,
  主表id_In     In 病案评分明细.主表id%Type,
  评分标准id_In In 病案评分明细.评分标准id%Type,
  单项分数_In   In 病案评分明细.单项分数%Type,
  缺陷等级_In   In 病案评分明细.缺陷等级%Type,
  可否修改_In   In 病案评分明细.可否修改%Type,
  备注_In       In 病案评分结果.备注%Type
) Is
Begin
  Update 病案评分明细
  Set 主表id = 主表id_In, 评分标准id = 评分标准id_In, 单项分数 = 单项分数_In, 缺陷等级 = 缺陷等级_In,
      可否修改 = 可否修改_In, 备注 = 备注_In
  Where ID = Id_In;

  If Sql%NotFound Then
    ---如果没有更新到，那么新增一条 
    Insert Into 病案评分明细
      (ID, 主表id, 评分标准id, 单项分数, 缺陷等级, 可否修改, 备注)
    Values
      (Id_In, 主表id_In, 评分标准id_In, 单项分数_In, 缺陷等级_In, 可否修改_In, 备注_In);
  End If;

Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病案评分明细_Update;
/


CREATE OR REPLACE PROCEDURE ZL_病案评分明细_Delete
(
	ID_in IN 病案评分明细.ID%TYPE
)
IS
BEGIN
	DELETE FROM 病案评分明细
		WHERE ID = ID_in;
EXCEPTION
	WHEN OTHERS THEN
		ZL_ErrorCenter (SQLCODE, SQLERRM);
END  ZL_病案评分明细_Delete;
/

  --病案评分结果-审核与取消审核

Create Or Replace Procedure Zl_病案评分结果_审核
(
  Id_In     In 病案评分结果.ID%Type,
  审核人_In In 病案评分结果.审核人%Type
) Is
  n_数目         Number;
  n_自动写入从表 Number;
  n_病人id       Number;
  n_主页id       Number;
  v_等级         Varchar2(2);
  v_旧等级值     Varchar2(2);
Begin
  Update 病案评分结果 Set 审核人 = 审核人_In, 审核时间 = Sysdate Where ID = Id_In;

  Select 病人id Into n_病人id From 病案评分结果 Where ID = Id_In;
  Select 主页id Into n_主页id From 病案评分结果 Where ID = Id_In;
  Select 等级 Into v_等级 From 病案评分结果 Where ID = Id_In;

  Select Count(*) Into n_数目 From 病案主页从表 Where 病人id = n_病人id And 主页id = n_主页id And 信息名 = '病案质量';

  n_自动写入从表 := To_Number(Zl_Getsysparameter(90, 0, 0), '9999999');

  If n_数目 = 0 And n_自动写入从表 = 1 And v_等级 <> '否' Then
    Insert Into 病案主页从表 (病人id, 主页id, 信息名, 信息值) Values (n_病人id, n_主页id, '病案质量', v_等级);
  Else
    If n_数目 = 1 And n_自动写入从表 = 1 And v_等级 <> '否' Then
      Select 信息值
      Into v_旧等级值
      From 病案主页从表
      Where 病人id = n_病人id And 主页id = n_主页id And 信息名 = '病案质量';
      If v_旧等级值 Is Null Then
        Update 病案主页从表 Set 信息值 = v_等级 Where 病人id = n_病人id And 主页id = n_主页id And 信息名 = '病案质量';
      End If;
    End If;
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病案评分结果_审核;
/

CREATE OR REPLACE PROCEDURE ZL_病案评分结果_取消审核
(	ID_in IN 病案评分结果.ID%TYPE
)
IS
BEGIN
	Update 病案评分结果
	set 审核人=NULL,审核时间=NULL
	where ID=ID_IN;


EXCEPTION
	WHEN OTHERS THEN
		ZL_ErrorCenter (SQLCODE, SQLERRM);
END  ZL_病案评分结果_取消审核;
/
----------------------------------------------------------------------------
---  UPDATE   for   病案借阅记录
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_病案借阅记录_Update(
	ID_IN		IN	病案借阅记录.ID%TYPE,
	No_IN		IN	病案借阅记录.No%TYPE,
	申请人_IN		IN	病案借阅记录.申请人%TYPE,
	申请理由_IN	IN	病案借阅记录.申请理由%TYPE,
	申请时间_IN	IN	病案借阅记录.申请时间%TYPE,
	申请期限_IN	IN	病案借阅记录.申请期限%TYPE,
	登记时间_IN	IN	病案借阅记录.登记时间%TYPE:=Sysdate
)
IS
BEGIN
	Update 病案借阅记录 Set No=No_IN,
				申请人=申请人_IN,
				申请理由=申请理由_IN,
				申请时间=申请时间_IN,
				申请期限=申请期限_IN
	Where ID=ID_IN And 记录状态=1;
	
	If SQL%RowCount=0 Then
		Insert Into 病案借阅记录(ID,No,记录状态,申请人,申请理由,申请时间,申请期限,登记时间) 
		VALUES (ID_IN,No_IN,1,申请人_IN,申请理由_IN,申请时间_IN,申请期限_IN,登记时间_IN);
	End If;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_病案借阅记录_Update;
/

----------------------------------------------------------------------------
---  Update   for   病案借阅人员
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_病案借阅人员_Update(
	借阅id_IN		IN	病案借阅人员.借阅id%TYPE,
	人员id_In		IN	Varchar2:=Null
)
IS
	strTmp			Varchar2(4000);
	intPos			Number(18);
	n_人员id			病案借阅人员.人员id%TYPE;
BEGIN
	If 人员id_IN Is Null Then
		Delete From 病案借阅人员 Where 借阅id=借阅id_IN;
	Else
		strTmp := 人员id_In||';';
		WHILE strTmp IS NOT NULL LOOP
			intPos := INSTR (strTmp, ';');
			IF intPos >0 Then
				n_人员id := To_Number(SUBSTR (strTmp, 1, intPos - 1));
				strTmp := SUBSTR (strTmp, intPos + 1);
				If n_人员id>0 Then
					Insert Into 病案借阅人员(借阅id,人员id) values (借阅id_IN,n_人员id);
				End If;
			End If;       
		END LOOP;
	End If;

EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_病案借阅人员_Update;
/

----------------------------------------------------------------------------
---  Update   for   病案借阅内容
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_病案借阅内容_Update(
	借阅id_IN		IN	病案借阅内容.借阅id%TYPE,
	病人id_In		IN	Varchar2:=Null
)
IS
	strTmp			Varchar2(4000);
	str病人			Varchar2(4000);
	intPos			Number(18);
	n_病人id			病案借阅内容.病人id%TYPE;
	n_主页id			病案借阅内容.主页id%TYPE;
BEGIN
	If 病人id_IN Is Null Then
		Delete From 病案借阅内容 Where 借阅id=借阅id_IN;
	Else
		strTmp := 病人id_In||';';
		WHILE strTmp IS NOT NULL LOOP
			intPos := INSTR (strTmp, ';');

			IF intPos >0 Then
				
				str病人 := SUBSTR (strTmp, 1, intPos - 1)||':';
				strTmp := SUBSTR (strTmp, intPos + 1);
				
				If str病人 Is Not Null Then
					intPos := INSTR (str病人, ':');
					n_病人id := To_Number(SUBSTR (str病人, 1, intPos - 1));				
					str病人 := SUBSTR (str病人, intPos + 1);

					intPos := INSTR (str病人, ':');
					n_主页id := To_Number(SUBSTR (str病人, 1, intPos - 1));
					str病人 := SUBSTR (str病人,intPos + 1);

					If n_病人id>0 And n_主页id>0 Then
						Insert Into 病案借阅内容(借阅id,病人id,主页id) values (借阅id_IN,n_病人id,n_主页id);
					End If;
				End If;
			End If;      
		END LOOP;
	End If;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_病案借阅内容_Update;
/

----------------------------------------------------------------------------
---  Delete   for   病案借阅记录
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_病案借阅记录_Delete(
	ID_IN		IN	病案借阅记录.ID%TYPE
)
IS
BEGIN
	Delete From 病案借阅记录 Where ID=ID_IN And 记录状态=1;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_病案借阅记录_Delete;
/

----------------------------------------------------------------------------
---  Authorize   for   病案借阅记录
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_病案借阅记录_Authorize(
	ID_IN		IN	病案借阅记录.ID%TYPE,
	借阅时间_In	In	病案借阅记录.借阅时间%TYPE,
	借阅期限_In	In	病案借阅记录.借阅期限%TYPE,
	批准人_In		In	病案借阅记录.批准人%TYPE,	
	批准时间_In	In	病案借阅记录.批准时间%TYPE:=Sysdate
)
IS
BEGIN
	Update 病案借阅记录 Set 借阅时间=借阅时间_In,借阅期限=借阅期限_In,批准人=批准人_In,批准时间=批准时间_In,记录状态=2 Where ID=ID_IN And 记录状态=1;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_病案借阅记录_Authorize;
/

----------------------------------------------------------------------------
---  Refuse   for   病案借阅记录
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_病案借阅记录_Refuse(
	ID_IN		IN	病案借阅记录.ID%TYPE,
	拒借人_In		In	病案借阅记录.拒借人%TYPE,
	拒借理由_In	In	病案借阅记录.拒借理由%TYPE,
	拒借时间_In	In	病案借阅记录.拒借时间%TYPE:=Sysdate
)
IS
BEGIN
	Update 病案借阅记录 Set 拒借人=拒借人_In,拒借理由=拒借理由_In,拒借时间=拒借时间_In,记录状态=3 Where ID=ID_IN And 记录状态=1;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_病案借阅记录_Refuse;
/

----------------------------------------------------------------------------
---  Rollback   for   病案借阅记录
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_病案借阅记录_Rollback(
	ID_IN		IN	病案借阅记录.ID%TYPE,
	回退性质_In	In	Number
)
IS
BEGIN
	If 回退性质_In=1 Then
		Update 病案借阅记录 Set 借阅时间=Null,借阅期限=Null,批准人=Null,批准时间=Null,记录状态=1 Where ID=ID_IN And 记录状态=2;
	ElsIf 回退性质_In=2 Then
		Update 病案借阅记录 Set 拒借人=Null,拒借理由=Null,拒借时间=Null,记录状态=1 Where ID=ID_IN And 记录状态=3;
	End If;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_病案借阅记录_Rollback;
/

----------------------------------------------------------------------------
---  Update   for   病案反馈记录
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_病案反馈记录_Update(
	ID_In			In	病案反馈记录.ID%TYPE,
	相关ID_In		In	病案反馈记录.相关ID%TYPE,
	提交ID_In		In	病案反馈记录.提交ID%TYPE,
	病人ID_In		In	病案反馈记录.病人ID%TYPE,
	主页ID_In		In	病案反馈记录.主页ID%TYPE,
	反馈对象_In	In	病案反馈记录.反馈对象%TYPE,
	文件ID_In		In	病案反馈记录.文件ID%TYPE,
	反馈意见_In	In	病案反馈记录.反馈意见%TYPE,
	反馈项目ID_In	In	病案反馈记录.反馈项目ID%TYPE,
	反馈人_In		In	病案反馈记录.反馈人%TYPE,
	反馈时间_In	In	病案反馈记录.反馈时间%TYPE,
	处理期限_In	In	病案反馈记录.处理期限%TYPE
)
IS
BEGIN
	
	Update 病案反馈记录 Set 	提交ID=Decode(提交ID_In,0,Null,提交ID_In),
						病人ID=病人ID_In,
						主页ID=主页ID_In,
						反馈对象=反馈对象_In,
						记录性质=Decode(提交ID_In,Null,1,0,1,2),
						记录状态=1,
						反馈意见=反馈意见_In,
						反馈项目ID=Decode(反馈项目ID_In,0,Null,反馈项目ID_In),
						文件ID=Decode(文件ID_In,0,Null,文件ID_In),
						反馈人=反馈人_In,
						反馈时间=反馈时间_In,
						处理期限=处理期限_In
	Where ID=ID_In;

	If SQL%RowCount=0 Then
		Insert Into 病案反馈记录(ID,相关ID,提交ID,病人ID,主页ID,反馈对象,文件ID,记录性质,记录状态,反馈意见,反馈项目ID,反馈人,反馈时间,处理期限)
		Values (ID_In,Decode(相关ID_In,0,Null,相关ID_In),Decode(提交ID_In,0,Null,提交ID_In),病人ID_In,主页ID_In,反馈对象_In,Decode(文件ID_In,0,Null,文件ID_In),Decode(提交ID_In,Null,1,0,1,2),1,反馈意见_In,Decode(反馈项目ID_In,0,Null,反馈项目ID_In),反馈人_In,反馈时间_In,处理期限_In);
		Update 病案主页 Set 病案状态=4 Where 病人id=病人ID_In And 主页ID=主页ID_In;
	End If;

EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_病案反馈记录_Update;
/
----------------------------------------------------------------------------
---  Finish   for   病案反馈记录
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_病案反馈记录_Finish(
	ID_In			In	病案反馈记录.ID%TYPE
)
IS
BEGIN	
	Update 病案反馈记录 Set 记录状态=3 Where ID=ID_In And 记录状态<>3;
	
	If SQL%RowCount>0 Then
		Update 病案主页 a Set 病案状态=3 Where (a.病人id,a.主页ID) In (Select 病人id,主页id From 病案反馈记录 Where ID=ID_In) And Not Exists (Select 1 From 病案反馈记录 b Where a.病人id=b.病人id And a.主页id=b.主页id And 记录状态=1);
	End If;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_病案反馈记录_Finish;
/

----------------------------------------------------------------------------
---  RollBackFinish   for   病案反馈记录
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_病案反馈记录_RollBackFinish(
	ID_In			In	病案反馈记录.ID%TYPE
)
IS
BEGIN	
	Update 病案反馈记录 Set 记录状态=Decode(处理人,Null,1,2) Where ID=ID_In;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_病案反馈记录_RollBackFinish;
/

----------------------------------------------------------------------------
---  Delete   for   病案反馈记录
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_病案反馈记录_Delete(
	ID_In			In	病案反馈记录.ID%TYPE
)
IS
BEGIN	
	Delete 病案反馈记录 Where ID=ID_In And 记录状态=1;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_病案反馈记录_Delete;
/
----------------------------------------------------------------------------
---  Commit   for   病案提交记录
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_病案提交记录_Commit(
	提交id_In		In	Varchar2,
	记录状态_In	In	病案提交记录.记录状态%Type,
	处理人_In		In	病案提交记录.接收人%Type,
	处理时间_In	In	病案提交记录.接收时间%Type:=Sysdate,
	拒审理由_In	In	病案提交记录.拒审理由%Type:=Null
)
IS
	v_Tmp			Varchar2(4000);
	n_提交id			Number(18);
	n_Pos			Number(18);
BEGIN		
	If 提交id_In Is Not Null Then
		v_Tmp := 提交id_In||',';
		WHILE v_Tmp IS NOT NULL LOOP
			n_Pos := INSTR (v_Tmp, ',');
			IF n_Pos >0 Then
				n_提交id := To_Number(SUBSTR (v_Tmp, 1, n_Pos - 1));
				v_Tmp := SUBSTR (v_Tmp, n_Pos + 1);

				If n_提交id>0 Then

					For r_List In (Select ID,接收人 From 病案提交记录 Where ID=n_提交id And 记录状态<>记录状态_In) Loop
						If 记录状态_In=3 Then
							--接收处理
							Update 病案主页 Set 病案状态=3 Where (病人id,主页id) In (Select 病人id,主页id From 病案提交记录 Where ID=r_List.ID) And Nvl(病案状态,0)<>3;
							Update 病案提交记录 Set 记录状态=3,接收人=处理人_In,接收时间=处理时间_In Where ID=r_List.ID And 记录状态<>3;
						ElsIf 记录状态_In=2 Then
							--拒绝审查
							Update 病案主页 Set 病案状态=2 Where (病人id,主页id) In (Select 病人id,主页id From 病案提交记录 Where ID=r_List.ID) And Nvl(病案状态,0)<>2;
							Update 病案提交记录 Set 记录状态=2,拒审人=处理人_In,拒审时间=处理时间_In,拒审理由=拒审理由_In Where ID=r_List.ID And 记录状态<>2;
						ElsIf 记录状态_In=5 Then
							--审查归档
							Update 病案主页 Set 病案状态=5 Where (病人id,主页id) In (Select 病人id,主页id From 病案提交记录 Where ID=r_List.ID) And Nvl(病案状态,0)<>5;
							Update 病案提交记录 Set 记录状态=5,归档人=处理人_In,归档时间=处理时间_In Where ID=r_List.ID And 记录状态<>5;
						End If;
					End Loop;
				End If;
			End If;
		End Loop;
	End If;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_病案提交记录_Commit;
/
----------------------------------------------------------------------------
---  Receive   for   病案提交记录
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_病案提交记录_Receive(
	提交id_In		In	Varchar2,
	接收人_In		In	病案提交记录.接收人%Type,
	接收时间_In	In	病案提交记录.接收时间%Type:=Sysdate
)
IS
BEGIN	

	If 提交id_In Is Not Null Then
		zl_病案提交记录_Commit(提交id_In,3,接收人_In,接收时间_In);
	End If;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_病案提交记录_Receive;
/
----------------------------------------------------------------------------
---  Refuse   for   病案提交记录
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_病案提交记录_Refuse(
	提交id_In		In	Varchar2,
	拒审人_In		In	病案提交记录.拒审人%Type,
	拒审时间_In	In	病案提交记录.拒审时间%Type:=Sysdate,
	拒审理由_In	In	病案提交记录.拒审理由%Type:=Null
)
IS
BEGIN	

	If 提交id_In Is Not Null Then
		zl_病案提交记录_Commit(提交id_In,2,拒审人_In,拒审时间_In,拒审理由_In);
	End If;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_病案提交记录_Refuse;
/
----------------------------------------------------------------------------
---  Archive   for   病案提交记录
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_病案提交记录_Archive(
	提交id_In		In	Varchar2,
	归档人_In		In	病案提交记录.归档人%Type,
	归档时间_In	In	病案提交记录.归档时间%Type:=Sysdate
)
IS
	n_Count			Number(18);
	v_Error			Varchar2(255);
	Err_Custom		Exception;
BEGIN	
	--检查是否是所有的反馈问题都已经完成了
	Select Count(1) Into n_Count From 病案反馈记录 Where 提交id=提交id_In And 记录状态 In (1,2);
	If n_Count>0 Then
		v_Error:='当前病人还有未完结的反馈问题。';
		Raise Err_Custom;
	End If;

	If 提交id_In Is Not Null Then		
		zl_病案提交记录_Commit(提交id_In,5,归档人_In,归档时间_In);
	End If;
EXCEPTION
	When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_病案提交记录_Archive;
/
----------------------------------------------------------------------------
---  RollBack   for   病案提交记录
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_病案提交记录_RollBack(
	提交id_In			In	Varchar2,
	回退状态_In	In	Number:=1
)
IS	
	v_Tmp			Varchar2(4000);
	n_提交id			Number(18);
	n_Pos			Number(18);
BEGIN		
	If 提交id_In Is Not Null Then
		v_Tmp := 提交id_In||',';
		WHILE v_Tmp IS NOT NULL LOOP
			n_Pos := INSTR (v_Tmp, ',');
			IF n_Pos >0 Then
				n_提交id := To_Number(SUBSTR (v_Tmp, 1, n_Pos - 1));
				v_Tmp := SUBSTR (v_Tmp, n_Pos + 1);

				If n_提交id>0 Then

					For r_List In (Select ID,接收人 From 病案提交记录 Where ID=n_提交id And 记录状态=回退状态_In) Loop
						If 回退状态_In=3 Then
							--回退接收
							Update 病案主页 Set 病案状态=1 Where (病人id,主页id) In (Select 病人id,主页id From 病案提交记录 Where ID=r_List.ID) And Nvl(病案状态,0)<>1;
							Update 病案提交记录 Set 记录状态=1,接收人=Null,接收时间=Null Where ID=r_List.ID And 记录状态<>1;
						ElsIf 回退状态_In=2 Then
							--回退拒绝
							Update 病案主页 Set 病案状态=1 Where (病人id,主页id) In (Select 病人id,主页id From 病案提交记录 Where ID=r_List.ID) And Nvl(病案状态,0)<>1;
							Update 病案提交记录 Set 记录状态=1,拒审人=Null,拒审时间=Null,拒审理由=Null Where ID=r_List.ID And 记录状态<>1;
						ElsIf 回退状态_In=5 Then
							--回退归档
							If r_List.接收人 Is Null Then
								Update 病案主页 Set 病案状态=1 Where (病人id,主页id) In (Select 病人id,主页id From 病案提交记录 Where ID=r_List.ID) And Nvl(病案状态,0)<>1;
								Update 病案提交记录 Set 记录状态=1,归档人=Null,归档时间=Null Where ID=r_List.ID And 记录状态<>1;
							Else
								Update 病案主页 Set 病案状态=3 Where (病人id,主页id) In (Select 病人id,主页id From 病案提交记录 Where ID=r_List.ID) And Nvl(病案状态,0)<>3;
								Update 病案提交记录 Set 记录状态=3,归档人=Null,归档时间=Null Where ID=r_List.ID And 记录状态<>3;
							End If;
						End If;
					End Loop;
				End If;
			End If;
		End Loop;
	End If;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_病案提交记录_RollBack;
/
----------------------------------------------------------------------------
---  UnReceive   for   病案提交记录
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_病案提交记录_UnReceive(
	提交id_In		In	Varchar2
)
IS
BEGIN	
	If 提交id_In Is Not Null Then
		zl_病案提交记录_RollBack(提交id_In,3);
	End If;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_病案提交记录_UnReceive;
/
----------------------------------------------------------------------------
---  UnRefuse   for   病案提交记录
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_病案提交记录_UnRefuse(
	提交id_In		In	Varchar2
)
IS
BEGIN	

	If 提交id_In Is Not Null Then
		zl_病案提交记录_RollBack(提交id_In,2);
	End If;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_病案提交记录_UnRefuse;
/
----------------------------------------------------------------------------
---  UnArchive   for   病案提交记录
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_病案提交记录_UnArchive(
	提交id_In		In	Varchar2
)
IS
BEGIN	
	If 提交id_In Is Not Null Then		
		zl_病案提交记录_RollBack(提交id_In,5);
	End If;
EXCEPTION
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_病案提交记录_UnArchive;
/

----------------------------------------------------------------------------
---  Lock   for   病案封存记录
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_病案封存记录_Lock(
	病人ID_In		In	病案封存记录.病人ID%Type,
	主页ID_In		In	病案封存记录.主页ID%Type,
	封存人_In		In	病案封存记录.封存人%Type,
	封存时间_In	In	病案封存记录.封存时间%Type:=Sysdate,
	封存理由_In	In	病案封存记录.封存理由%Type:=Null
)
IS
	v_Error			Varchar2(255);
	Err_Custom		Exception;
BEGIN	
	Update 病案主页 Set 封存时间=封存时间_In Where 病人id=病人ID_In And 主页ID=主页ID_In And 封存时间 Is Null;
	If SQL%RowCount=0 Then
		v_Error:='当前病人已经被封存或不存在的病人。';
		Raise Err_Custom;
	End If;

	Insert Into 病案封存记录(ID,病人ID,主页ID,记录状态,封存人,封存时间,封存理由)
	Select  病案封存记录_ID.NextVal ,病人ID_In,主页ID_In,1,封存人_In,封存时间_In,封存理由_In From Dual;

EXCEPTION
	When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_病案封存记录_Lock;
/

----------------------------------------------------------------------------
---  UnLock   for   病案封存记录
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_病案封存记录_UnLock(
	病人ID_In		In	病案封存记录.病人ID%Type,
	主页ID_In		In	病案封存记录.主页ID%Type
)
IS
	v_Error			Varchar2(255);
	Err_Custom		Exception;
BEGIN		
	Update 病案主页 Set 封存时间=Null Where 病人id=病人ID_In And 主页ID=主页ID_In And 封存时间 Is Not Null;
	If SQL%RowCount=0 Then
		v_Error:='当前病人已经解除封存或不存在的病人。';
		Raise Err_Custom;
	End If;

	Update 病案封存记录 Set 记录状态=2 Where 病人id=病人ID_In And 主页ID=主页ID_In And 记录状态=1;
EXCEPTION
	When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
	WHEN OTHERS THEN Zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_病案封存记录_UnLock;
/

------------------------------------------------------------------------------------------------------------------------------------------
--		报表部份
------------------------------------------------------------------------------------------------------------------------------------------
--报表：ZL1_INSIDE_1562_1/单病案评分结果表
Insert Into zlReports(ID,编号,名称,说明,密码,打印机,进纸,票据,系统,程序ID,功能,修改时间,发布时间) Values(zlReports_ID.NextVal,'ZL1_INSIDE_1562_1','单病案评分结果表','打印评分结果表','Jf9m dmyc96Rhfo1H*W^','Microsoft Office Document Image Writer',1,0,100,1562,'打印评分结果表',Sysdate,Sysdate);
Insert Into zlRPTFmts(报表ID,序号,说明,W,H,纸张,纸向,动态纸张,图样) Values(zlReports_ID.CurrVal,1,'单病案评分结果统计表1',11904,16838,9,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签2',2,Null,0,'任意表1',11,'填报单位:[单位名称]',Null,795,1305,1710,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,1);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签7',2,Null,0,'任意表1',11,'出院科室:[病案评分结果_数据.出院科室]',Null,795,2105,3330,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签14',2,Null,0,'任意表1',21,'评分人:[病案评分结果_数据.评分人]',Null,795,15008,2970,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签16',2,Null,0,'任意表1',21,'审核人:[病案评分结果_数据.审核人]',Null,795,15300,2970,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签3',2,Null,0,'任意表1',21,'制表人:[操作员姓名]',Null,795,15592,1710,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签15',2,Null,0,'任意表1',22,'评分时间:[病案评分结果_数据.评分时间]',Null,4482,15008,3330,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签17',2,Null,0,'任意表1',22,'审核时间:[病案评分结果_数据.审核时间]',Null,4482,15300,3330,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签9',2,Null,0,'任意表1',11,'住 院 号:[病案评分结果_数据.住院号]',Null,795,1590,3150,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签6',2,Null,0,'任意表1',12,'住院次数:[病案评分结果_数据.住院次数]',Null,4482,1845,3330,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签10',2,Null,0,'任意表1',12,'住院医师:[病案评分结果_数据.住院医师]',Null,4482,2105,3330,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签1',2,Null,0,'任意表1',12,'病案评分结果表',Null,4835,675,2625,360,0,0,1,'楷体_GB2312',18,1,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签4',2,Null,0,'任意表1',22,'第[页号]页',Null,5698,15592,900,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签13',2,Null,0,'任意表1',13,'等    级:[病案评分结果_数据.等级]',Null,8530,1590,2970,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签12',2,Null,0,'任意表1',13,'总    分:[病案评分结果_数据.总分]',Null,8530,1845,2970,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签11',2,Null,0,'任意表1',13,'出院日期:[病案评分结果_数据.出院日期]',Null,8170,2105,3330,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签5',2,Null,0,'任意表1',23,'报表日期：[YYYY-MM-DD]',Null,9520,15592,1980,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'任意表1',4,Null,0,Null,0,'病案评分明细_数据',Null,795,2385,10705,12478,255,0,1,'宋体',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-1,0,Null,Null,'[病案评分明细_数据.项目]','4^435^项目',0,0,930,0,255,1,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-2,1,Null,Null,'[病案评分明细_数据.标准分值]','4^435^标准分值',0,0,555,0,255,1,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-3,2,Null,Null,'[病案评分明细_数据.基本要求]','4^435^基本要求',0,0,2145,0,255,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-4,3,Null,Null,'[病案评分明细_数据.缺陷内容]','4^435^缺陷内容',0,0,3540,0,255,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-5,4,Null,Null,'[病案评分明细_数据.扣分标准]','4^435^扣分标准',0,0,1140,0,255,1,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-6,5,Null,Null,'[病案评分明细_数据.评分]','4^435^评分',0,0,1140,0,255,1,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签18',2,Null,0,'任意表1',12,'返回修改:[病案评分结果_数据.返回修改]',Null,4482,1590,3330,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签8',2,Null,0,'任意表1',11,'姓    名:[病案评分结果_数据.姓名]',Null,795,1845,2970,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'病案评分结果_数据','ID,131|住院次数,131|出院科室,200|姓名,200|住院号,131|住院医师,200|出院日期,200|总分,200|等级,200|返回修改,200|评分人,200|评分时间,200|审核人,200|审核时间,200',User||'.部门表,'||User||'.病案评分结果,'||User||'.病案主页,'||User||'.病人信息',0,Null);
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,1,'select A.ID,A.主页ID as 住院次数,(select 部门表.名称 from 部门表 where 部门表.id=B.出院科室ID) as 出院科室,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,2,'C.姓名,C.住院号,B.住院医师,TO_CHAR(B.出院日期,''YYYY-MM-DD'') as 出院日期,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,3,'decode(A.等级,''否'',''-     '',A.总分) as 总分,decode(A.等级,''否'',''不合格'',A.等级) as 等级,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,4,'decode(A.返回修改,null,''否'',0,''否'',1,''是'') as 返回修改,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,5,'A.评分人,TO_CHAR(A.评分时间,''YYYY-MM-DD'') as 评分时间,A.审核人,TO_CHAR(A.审核时间,''YYYY-MM-DD'') as 审核时间 ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,6,'from 病案评分结果 A, 病案主页 B,病人信息 C');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,7,'where A.病人ID=C.病人ID and A.病人ID= B.病人ID and A.主页ID=B.主页ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,8,'and ID=[0]');
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,Null,0,'结果ID',1,Null,0,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'病案评分明细_数据','ID,131|项目,200|标准分值,200|基本要求,200|缺陷内容,200|扣分标准,200|评分,200',User||'.病案评分明细,'||User||'.病案评分结果,'||User||'.病案评分标准视图',0,Null);
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,1,'select A.ID,A.项目,TO_CHAR(A.标准分值)||''分'' as 标准分值,A.基本要求,A.缺陷内容,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,2,'decode(A.扣分标准,''甲'',''甲级'',''乙'',''乙级'',''丙'',''丙级'',''否'',''单项否决'',A.扣分标准) as 扣分标准,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,3,'(');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,4,'select decode(缺陷等级,null,to_CHAR(单项分数),''否'',''单项否决'',缺陷等级||''级'') ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,5,'from 病案评分明细 where 评分标准ID=A.ID and 主表ID=[0]');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,6,') as 评分  ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,7,'from 病案评分标准视图 A  ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,8,'where A.隐藏=''否'' and A.方案ID=(select B.方案ID from 病案评分结果 B where B.ID=[0])  ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,9,'order by A.上级ID,A.ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,10,Null);
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,Null,0,'结果ID',1,Null,0,Null,Null,Null,Null,Null,Null);

--报表：ZL1_REPORT_1570/医师病案质量统计表
Insert Into zlReports(ID,编号,名称,说明,密码,打印机,进纸,票据,系统,程序ID,功能,修改时间,发布时间) Values(zlReports_ID.NextVal,'ZL1_REPORT_1570','医师病案质量统计表','医师病案质量统计表','Ww:mXk|ws35VitmcW*O]','Microsoft Office Document Image Writer',1,0,100,1570,'基本',Sysdate,Sysdate);
Insert Into zlRPTFmts(报表ID,序号,说明,W,H,纸张,纸向,动态纸张,图样) Values(zlReports_ID.CurrVal,1,'医师病案质量统计表1',11904,16838,9,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签2',2,Null,0,'汇总表1',11,'填报单位:[单位名称]',Null,270,935,1710,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,1);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签3',2,Null,0,'汇总表1',21,'制表人:[操作员姓名]',Null,270,15325,1710,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签1',2,Null,0,Null,0,'住院医师病案质量统计表',Null,3822,395,4125,360,0,0,1,'楷体_GB2312',18,1,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签5',2,Null,0,'汇总表1',22,'第[页号]页',Null,5295,15325,900,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签6',2,Null,0,'汇总表1',13,'统计时间:[=开始日期] 至 [=结束日期]',Null,8070,935,3150,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签4',2,Null,0,'汇总表1',23,'报表日期：[YYYY-MM-DD]',Null,9240,15325,1980,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'汇总表1',5,Null,0,Null,0,'病案质量报表_数据',Null,270,1245,10950,13980,255,0,0,'宋体',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,7,zlRPTItems_ID.CurrVal-1,0,Null,Null,'住院医师',Null,0,0,795,0,255,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-2,0,Null,Null,'病案总数',Null,0,0,945,0,255,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-3,1,Null,Null,'已评病案',Null,0,0,930,0,255,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-4,2,Null,Null,'已审病案',Null,0,0,885,0,255,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-5,3,Null,Null,'甲等',Null,0,0,855,0,255,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-6,4,Null,Null,'乙等',Null,0,0,810,0,255,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-7,5,Null,Null,'丙等',Null,0,0,810,0,255,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-8,6,Null,Null,'返回修改数',Null,0,0,930,0,255,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-9,7,Null,Null,'甲等率',Null,0,0,915,0,255,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-10,8,Null,Null,'乙等率',Null,0,0,930,0,255,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-11,9,Null,Null,'丙等率',Null,0,0,900,0,255,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-12,10,Null,Null,'返回修改率',Null,0,0,1005,0,255,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'病案质量报表_数据','住院医师,200|病案总数,139|已评病案,139|已审病案,139|甲等,139|乙等,139|丙等,139|返回修改数,139|甲等率,139|乙等率,139|丙等率,139|返回修改率,139',User||'.病案质量报表视图',1,Null);
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,1,'Select 住院医师,count(*) as 病案总数,count(评分时间) as 已评病案,count(审核时间) as 已审病案,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,2,'       count(decode(等级,''甲'',审核时间,null)) as 甲等,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,3,'       count(decode(等级,''乙'',审核时间,null)) as 乙等,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,4,'       count(decode(等级,''丙'',审核时间,null)) as 丙等,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,5,'       count(decode(返回修改,1,审核时间,null)) as 返回修改数,       ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,6,'       round(decode(count(审核时间),0,0,count(decode(等级,''甲'',审核时间,null))/count(审核时间))*100,1) as 甲等率,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,7,'       round(decode(count(审核时间),0,0,count(decode(等级,''乙'',审核时间,null))/count(审核时间))*100,1) as 乙等率,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,8,'       round(decode(count(审核时间),0,0,count(decode(等级,''丙'',审核时间,null))/count(审核时间))*100,1) as 丙等率,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,9,'       round(decode(count(审核时间),0,0,count(decode(返回修改,1,审核时间,null))/count(审核时间))*100,1) as 返回修改率');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,10,'From 病案质量报表视图');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,11,'where 出院日期 between [0] and [1]+1 ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,12,'group by 住院医师');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,13,Null);
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,14,'union all');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,15,'Select ''合计'',count(*) as 病案总数,count(评分时间) as 已评病案,count(审核时间) as 已审病案,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,16,'       count(decode(等级,''甲'',审核时间,null)) as 甲等,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,17,'       count(decode(等级,''乙'',审核时间,null)) as 乙等,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,18,'       count(decode(等级,''丙'',审核时间,null)) as 丙等,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,19,'       count(decode(返回修改,1,审核时间,null)) as 返回修改数,       ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,20,'       round(decode(count(审核时间),0,0,count(decode(等级,''甲'',审核时间,null))/count(审核时间))*100,1) as 甲等率,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,21,'       round(decode(count(审核时间),0,0,count(decode(等级,''乙'',审核时间,null))/count(审核时间))*100,1) as 乙等率,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,22,'       round(decode(count(审核时间),0,0,count(decode(等级,''丙'',审核时间,null))/count(审核时间))*100,1) as 丙等率,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,23,'       round(decode(count(审核时间),0,0,count(decode(返回修改,1,审核时间,null))/count(审核时间))*100,1) as 返回修改率');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,24,'From 病案质量报表视图');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,25,'where 出院日期 between [0] and [1]+1 ');
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,Null,0,'开始日期',2,CHR(38)||'前一月日期',0,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,Null,1,'结束日期',2,CHR(38)||'当前日期',0,Null,Null,Null,Null,Null,Null);

--报表：ZL1_REPORT_1571/科室病案质量统计表
Insert Into zlReports(ID,编号,名称,说明,密码,打印机,进纸,票据,系统,程序ID,功能,修改时间,发布时间) Values(zlReports_ID.NextVal,'ZL1_REPORT_1571','科室病案质量统计表','科室病案质量统计表','Ew:mXj|ws!5VitlcW*]]','Microsoft Office Document Image Writer',1,0,100,1571,'基本',Sysdate,Sysdate);
Insert Into zlRPTFmts(报表ID,序号,说明,W,H,纸张,纸向,动态纸张,图样) Values(zlReports_ID.CurrVal,1,'科室病案质量统计表1',11904,16838,9,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签2',2,Null,0,'汇总表1',11,'填报单位:[单位名称]',Null,255,1190,1710,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,1);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签3',2,Null,0,'汇总表1',21,'制表人:[操作员姓名]',Null,255,15295,1710,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签1',2,Null,0,Null,0,'科室病案质量统计表',Null,4205,645,3375,360,0,1,1,'楷体_GB2312',18,1,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签5',2,Null,0,'汇总表1',22,'第[页号]页',Null,5385,15295,900,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签6',2,Null,0,'汇总表1',13,'统计时间:[=开始日期] 至 [=结束日期]',Null,8265,1190,3150,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签4',2,Null,0,'汇总表1',23,'报表日期：[YYYY-MM-DD]',Null,9435,15295,1980,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'汇总表1',5,Null,0,Null,0,'病案质量报表_数据',Null,255,1500,11160,13695,255,0,0,'宋体',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,7,zlRPTItems_ID.CurrVal-1,0,Null,Null,'科室',Null,0,0,840,0,255,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-2,0,Null,Null,'病案总数',Null,0,0,855,0,255,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-3,1,Null,Null,'已评病案',Null,0,0,930,0,255,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-4,2,Null,Null,'已审病案',Null,0,0,900,0,255,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-5,3,Null,Null,'甲等',Null,0,0,840,0,255,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-6,4,Null,Null,'乙等',Null,0,0,870,0,255,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-7,5,Null,Null,'丙等',Null,0,0,825,0,255,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-8,6,Null,Null,'返回修改数',Null,0,0,1020,0,255,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-9,7,Null,Null,'甲等率',Null,0,0,1005,0,255,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-10,8,Null,Null,'乙等率',Null,0,0,1005,0,255,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-11,9,Null,Null,'丙等率',Null,0,0,1005,0,255,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-12,10,Null,Null,'返回修改率',Null,0,0,1005,0,255,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'病案质量报表_数据','科室,200|病案总数,139|已评病案,139|已审病案,139|甲等,139|乙等,139|丙等,139|返回修改数,139|甲等率,139|乙等率,139|丙等率,139|返回修改率,139',User||'.病案质量报表视图',1,Null);
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,1,'Select 出院科室 as 科室,count(*) as 病案总数,count(评分时间) as 已评病案,count(审核时间) as 已审病案,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,2,'       count(decode(等级,''甲'',审核时间,null)) as 甲等,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,3,'       count(decode(等级,''乙'',审核时间,null)) as 乙等,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,4,'       count(decode(等级,''丙'',审核时间,null)) as 丙等,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,5,'       count(decode(返回修改,1,审核时间,null)) as 返回修改数,       ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,6,'       round(decode(count(审核时间),0,0,count(decode(等级,''甲'',审核时间,null))/count(审核时间))*100,1) as 甲等率,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,7,'       round(decode(count(审核时间),0,0,count(decode(等级,''乙'',审核时间,null))/count(审核时间))*100,1) as 乙等率,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,8,'       round(decode(count(审核时间),0,0,count(decode(等级,''丙'',审核时间,null))/count(审核时间))*100,1) as 丙等率,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,9,'       round(decode(count(审核时间),0,0,count(decode(返回修改,1,审核时间,null))/count(审核时间))*100,1) as 返回修改率');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,10,'From 病案质量报表视图');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,11,'where 出院日期 between [0] and [1]+1 ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,12,'group by 出院科室');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,13,Null);
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,14,'union all');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,15,'Select ''合计'',count(*) as 病案总数,count(评分时间) as 已评病案,count(审核时间) as 已审病案,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,16,'       count(decode(等级,''甲'',审核时间,null)) as 甲等,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,17,'       count(decode(等级,''乙'',审核时间,null)) as 乙等,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,18,'       count(decode(等级,''丙'',审核时间,null)) as 丙等,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,19,'       count(decode(返回修改,1,审核时间,null)) as 返回修改数,       ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,20,'       round(decode(count(审核时间),0,0,count(decode(等级,''甲'',审核时间,null))/count(审核时间))*100,1) as 甲等率,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,21,'       round(decode(count(审核时间),0,0,count(decode(等级,''乙'',审核时间,null))/count(审核时间))*100,1) as 乙等率,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,22,'       round(decode(count(审核时间),0,0,count(decode(等级,''丙'',审核时间,null))/count(审核时间))*100,1) as 丙等率,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,23,'       round(decode(count(审核时间),0,0,count(decode(返回修改,1,审核时间,null))/count(审核时间))*100,1) as 返回修改率');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,24,'From 病案质量报表视图');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,25,'where 出院日期 between [0] and [1]+1 ');
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,Null,0,'开始日期',2,CHR(38)||'前一月日期',0,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,Null,1,'结束日期',2,CHR(38)||'当前日期',0,Null,Null,Null,Null,Null,Null);

--报表：ZL1_REPORT_1572/病案评分结果清单
Insert Into zlReports(ID,编号,名称,说明,密码,打印机,进纸,票据,系统,程序ID,功能,修改时间,发布时间) Values(zlReports_ID.NextVal,'ZL1_REPORT_1572','病案评分结果清单','病案评分结果清单','Le(jHbyys+6Rbirs_)FH','Microsoft Office Document Image Writer',1,0,100,1572,'基本',Sysdate,Sysdate);
Insert Into zlRPTFmts(报表ID,序号,说明,W,H,纸张,纸向,动态纸张,图样) Values(zlReports_ID.CurrVal,1,'11',11904,16838,9,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签2',2,Null,0,'任意表1',11,'填报单位:[单位名称]',Null,345,995,1710,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,1);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签3',2,Null,0,'任意表1',21,'制表人:[操作员姓名]',Null,345,15370,1710,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签1',2,Null,0,Null,0,'病案评分结果清单',Null,4015,525,3000,360,0,0,1,'楷体_GB2312',18,1,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签5',2,Null,0,'任意表1',22,'第[页号]页',Null,5452,15370,900,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签6',2,Null,0,'任意表1',13,'统计时间:[=开始日期] 至 [=结束日期]',Null,8309,995,3150,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签4',2,Null,0,'任意表1',23,'报表日期：[YYYY-MM-DD]',Null,9479,15370,1980,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'任意表1',4,Null,0,Null,0,'病案质量报表_数据',Null,345,1305,11114,13965,255,0,0,'宋体',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-1,0,Null,Null,'[病案质量报表_数据.住院号]','4^255^住院号',0,0,960,0,255,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-2,1,Null,Null,'[病案质量报表_数据.姓名]','4^255^姓名',0,0,825,0,255,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-3,2,Null,Null,'[病案质量报表_数据.出院科室]','4^255^出院科室',0,0,1005,0,255,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-4,3,Null,Null,'[病案质量报表_数据.责任护士]','4^255^责任护士',0,0,1005,0,255,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-5,4,Null,Null,'[病案质量报表_数据.住院医师]','4^255^住院医师',0,0,900,0,255,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-6,5,Null,Null,'[病案质量报表_数据.总分]','4^255^总分',0,0,675,0,255,2,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-7,6,Null,Null,'[病案质量报表_数据.等级]','4^255^等级',0,0,585,0,255,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-8,7,Null,Null,'[病案质量报表_数据.返回修改]','4^255^返回修改',0,0,885,0,255,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-9,8,Null,Null,'[病案质量报表_数据.评分人]','4^255^评分人',0,0,1005,0,255,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-10,9,Null,Null,'[病案质量报表_数据.评分时间]','4^255^评分时间',0,0,1005,0,255,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-11,10,Null,Null,'[病案质量报表_数据.审核人]','4^255^审核人',0,0,1005,0,255,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-12,11,Null,Null,'[病案质量报表_数据.审核时间]','4^255^审核时间',0,0,1005,0,255,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'病案质量报表_数据','住院号,131|姓名,200|出院科室,200|责任护士,200|住院医师,200|总分,131|等级,200|返回修改,200|评分人,200|评分时间,200|审核人,200|审核时间,200',User||'.病案质量报表视图',0,Null);
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,1,'Select 住院号,姓名,出院科室,责任护士,住院医师,总分,等级,decode(返回修改,1,''是'','''') as 返回修改,评分人,评分时间,审核人,审核时间');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,2,'From 病案质量报表视图');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,3,'where 评分时间 is not null and 出院日期 between [0] and [1]+1 ');
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,Null,0,'开始日期',2,CHR(38)||'前一月日期',0,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,Null,1,'结束日期',2,CHR(38)||'当前日期',0,Null,Null,Null,Null,Null,Null);

--报表：ZL1_INSIDE_1562_1/单病案评分结果表
Insert into zlProgFuncs(系统,序号,功能,说明) Values(100,1562,'打印评分结果表',Null);
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1562,'打印评分结果表',User,'病案评分标准视图','SELECT');
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1562,'打印评分结果表',User,'病案评分结果','SELECT');
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1562,'打印评分结果表',User,'病案评分明细','SELECT');
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1562,'打印评分结果表',User,'病案主页','SELECT');
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1562,'打印评分结果表',User,'病人信息','SELECT');
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1562,'打印评分结果表',User,'部门表','SELECT');
--报表：ZL1_REPORT_1570/医师病案质量统计表
Insert into zlPrograms(序号,标题,说明,系统,部件) Values(1570,'医师病案质量统计表','医师病案质量统计表',100,'zl9Report');
Insert into zlProgFuncs(系统,序号,功能,说明) Values(100,1570,'基本',Null);
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1570,'基本',User,'病案质量报表视图','SELECT');
Insert into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) Select '缺省',zlMenus_ID.Nextval,ID,'医师病案质量统计表','医师病案质量统计表',Null,105,'医师病案质量统计表',100,1570 From zlMenus Where 系统=100 And 组别='缺省' And 标题='病案报表分析' And 模块 is NULL;

--报表：ZL1_REPORT_1571/科室病案质量统计表
Insert into zlPrograms(序号,标题,说明,系统,部件) Values(1571,'科室病案质量统计表','科室病案质量统计表',100,'zl9Report');
Insert into zlProgFuncs(系统,序号,功能,说明) Values(100,1571,'基本',Null);
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1571,'基本',User,'病案质量报表视图','SELECT');
Insert into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) Select '缺省',zlMenus_ID.Nextval,ID,'科室病案质量统计表','科室病案质量统计表',Null,105,'科室病案质量统计表',100,1571 From zlMenus Where 系统=100 And 组别='缺省' And 标题='病案报表分析' And 模块 is NULL;

--报表：ZL1_REPORT_1572/病案评分结果清单
Insert into zlPrograms(序号,标题,说明,系统,部件) Values(1572,'病案评分结果清单','病案评分结果清单',100,'zl9Report');
Insert into zlProgFuncs(系统,序号,功能,说明) Values(100,1572,'基本',Null);
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1572,'基本',User,'病案质量报表视图','SELECT');
Insert into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) Select '缺省',zlMenus_ID.Nextval,ID,'病案评分结果清单','病案评分结果清单',Null,105,'病案评分结果清单',100,1572 From zlMenus Where 系统=100 And 组别='缺省' And 标题='病案报表分析' And 模块 is NULL;
