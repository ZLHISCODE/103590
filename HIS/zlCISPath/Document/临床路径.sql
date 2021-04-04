--------------------------------------------------------------------------------------------------------------------------
--临床路径数据结构部分
--------------------------------------------------------------------------------------------------------------------------
Create Table 临床病例分型(
    编码 VARCHAR2(1),
    名称 VARCHAR2(20),
    简码 VARCHAR2(10))
    TABLESPACE zl9BaseItem
    PCTFREE 5  
		PCTUSED 85;
Alter Table 临床病例分型 Add Constraint 临床病例分型_PK Primary Key (编码) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table 临床病例分型 Add Constraint 临床病例分型_UQ_名称 Unique (名称) Using Index Pctfree 5 Tablespace zl9indexcis;


Create Table 路径常见结果(
    编码 VARCHAR2(5),
    名称 VARCHAR2(100),
    简码 VARCHAR2(10),
		上级 VARCHAR2(5),
		末级 NUMBER(1))
    TABLESPACE zl9BaseItem
    PCTFREE 5  
		PCTUSED 85;
Alter Table 路径常见结果 Add Constraint 路径常见结果_PK Primary Key (编码) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table 路径常见结果 Add Constraint 路径常见结果_UQ_名称 Unique (上级,名称) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table 路径常见结果 Add Constraint 路径常见结果_CK_末级 Check (末级 in(0,1));


Create Table 变异常见原因(
    编码 VARCHAR2(5),
    名称 VARCHAR2(100),
    简码 VARCHAR2(10),
		上级 VARCHAR2(5),
		末级 NUMBER(1))
    TABLESPACE zl9BaseItem
    PCTFREE 5  
		PCTUSED 85;
Alter Table 变异常见原因 Add Constraint 变异常见原因_PK Primary Key (编码) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table 变异常见原因 Add Constraint 变异常见原因_UQ_名称 Unique (上级,名称) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table 变异常见原因 Add Constraint 变异常见原因_CK_末级 Check (末级 in(0,1));


Create Sequence 临床路径图标_ID Start With 1;
CREATE TABLE 临床路径图标(
		ID NUMBER(18),
		图标 BLOB,
		性质 NUMBER(1))
		LOB(图标) Store as (Cache)
    TABLESPACE zl9BaseItem
    PCTFREE 20
    PCTUSED 70;
Alter Table 临床路径图标 Add Constraint 临床路径图标_PK Primary Key (ID) Using Index Pctfree 5 Tablespace zl9indexcis;

Create Sequence 临床路径目录_ID Start With 1;
CREATE TABLE 临床路径目录(
    ID NUMBER(18),
		分类 VARCHAR2(50),
    编码 VARCHAR2(5),
    名称 VARCHAR2(100),
    通用 NUMBER(1),
    最新版本 NUMBER(3),
    病例分型 VARCHAR2(20),
    适用病情 VARCHAR2(20),
		适用性别 NUMBER(1),
		适用年龄 VARCHAR2(10),
    说明 VARCHAR2(200))
    TABLESPACE zl9BaseItem
    PCTFREE 5
    PCTUSED 85;
Alter Table 临床路径目录 Add Constraint 临床路径目录_PK Primary Key (ID) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table 临床路径目录 Add Constraint 临床路径目录_UQ_编码 Unique (编码,分类) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table 临床路径目录 Add Constraint 临床路径目录_UQ_名称 Unique (名称,分类) Using Index Pctfree 5 Tablespace zl9indexcis;


CREATE TABLE 临床路径病种(
    路径ID NUMBER(18),
    疾病ID NUMBER(18),
		诊断ID NUMBER(18))
    TABLESPACE zl9BaseItem
    PCTFREE 5
    PCTUSED 85;
Alter Table 临床路径病种 Add Constraint 临床路径病种_UQ_病种ID Unique (路径ID,疾病ID,诊断ID) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table 临床路径病种 Add Constraint 临床路径病种_FK_路径ID Foreign Key (路径ID) References 临床路径目录(ID) On Delete Cascade;
Alter Table 临床路径病种 Add Constraint 临床路径病种_FK_疾病ID Foreign Key (疾病ID) References 疾病编码目录(ID) On Delete Cascade;
Alter Table 临床路径病种 Add Constraint 临床路径病种_FK_诊断ID Foreign Key (诊断ID) References 疾病诊断目录(ID) On Delete Cascade;


CREATE TABLE 临床路径科室(
    路径ID NUMBER(18),
    科室ID NUMBER(18))
    TABLESPACE zl9BaseItem
    PCTFREE 5
    PCTUSED 85;
Alter Table 临床路径科室 Add Constraint 临床路径科室_PK Primary Key (路径ID,科室ID) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table 临床路径科室 Add Constraint 临床路径科室_FK_路径ID Foreign Key (路径ID) References 临床路径目录(ID) On Delete Cascade;
Alter Table 临床路径科室 Add Constraint 临床路径科室_FK_科室ID Foreign Key (科室ID) References 部门表(ID) On Delete Cascade;


CREATE TABLE 临床路径文件(
    路径ID NUMBER(18),
		文件名 VARCHAR2(200),
    内容 BLOB,
		创建人 VARCHAR2(20),
		创建时间 DATE)
    TABLESPACE zl9BaseItem
    PCTFREE 20
    PCTUSED 70;
Alter Table 临床路径文件 Add Constraint 临床路径文件_PK Primary Key (路径ID,文件名) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table 临床路径文件 Add Constraint 临床路径文件_FK_路径ID Foreign Key (路径ID) References 临床路径目录(ID) On Delete Cascade;

CREATE TABLE 临床路径版本(
    路径ID NUMBER(18),
    版本号 NUMBER(3),
    标准住院日 VARCHAR2(10),
    标准费用 VARCHAR2(20),
    版本说明 VARCHAR2(200),
    创建人 VARCHAR2(20),
    创建时间 DATE,
    审核人 VARCHAR2(20),
    审核时间 DATE,
		停用人 VARCHAR2(20),
    停用时间 DATE)
    TABLESPACE zl9BaseItem
    PCTFREE 5
    PCTUSED 85;
Alter Table 临床路径版本 Add Constraint 临床路径版本_PK Primary Key (路径ID,版本号) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table 临床路径版本 Add Constraint 临床路径版本_FK_路径ID Foreign Key (路径ID) References 临床路径目录(ID) On Delete Cascade;


Create Sequence 临床路径阶段_ID Start With 1;
CREATE TABLE 临床路径阶段(
		ID NUMBER(18),
    路径ID NUMBER(18),
    版本号 NUMBER(3),
		父ID NUMBER(18),
    序号 NUMBER(5),
    名称 VARCHAR2(50),
    开始天数 NUMBER(3),
    结束天数 NUMBER(3),
    标志 VARCHAR2(10),
    说明 VARCHAR2(200))
    TABLESPACE zl9BaseItem
    PCTFREE 5
    PCTUSED 85;
Alter Table 临床路径阶段 Add Constraint 临床路径阶段_PK Primary Key (ID) Using Index Pctfree 5 Tablespace zl9indexcis;
--序号使用延迟约束为了路径表调整保存过程完之前序号不重复检查
Alter Table 临床路径阶段 Add Constraint 临床路径阶段_UQ_序号 Unique (路径ID,版本号,父ID,序号) Deferrable Initially Deferred Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table 临床路径阶段 Add Constraint 临床路径阶段_FK_版本号 Foreign Key (路径ID,版本号) References 临床路径版本(路径ID,版本号) On Delete Cascade;
Alter Table 临床路径阶段 Add Constraint 临床路径阶段_FK_父ID Foreign Key (父ID) References 临床路径阶段(ID) On Delete Cascade;


CREATE TABLE 临床路径分类(
    路径ID NUMBER(18),
    版本号 NUMBER(3),
    序号 NUMBER(5),
		名称 VARCHAR2(50))
    TABLESPACE zl9BaseItem
    PCTFREE 5
    PCTUSED 85;
Alter Table 临床路径分类 Add Constraint 临床路径分类_PK Primary Key (路径ID,版本号,序号) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table 临床路径分类 Add Constraint 临床路径分类_UQ_名称 Unique (路径ID,版本号,名称) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table 临床路径分类 Add Constraint 临床路径分类_FK_版本号 Foreign Key (路径ID,版本号) References 临床路径版本(路径ID,版本号) On Delete Cascade;

Create Sequence 临床路径项目_ID Start With 1;
CREATE TABLE 临床路径项目(
		ID NUMBER(18),
    路径ID NUMBER(18),
    版本号 NUMBER(3),
    阶段ID NUMBER(18),
		分类 VARCHAR2(50),
		项目序号 NUMBER(5),
		项目内容 VARCHAR2(1000),
		执行方式 NUMBER(1),
		执行者 NUMBER(1),
		项目结果 VARCHAR2(500),
		图标ID NUMBER(18))
    TABLESPACE zl9BaseItem
    PCTFREE 5
    PCTUSED 85;
Alter Table 临床路径项目 Add Constraint 临床路径项目_PK Primary Key (ID) Using Index Pctfree 5 Tablespace zl9indexcis;
--序号使用延迟约束为了路径表调整保存过程完之前序号不重复检查
Alter Table 临床路径项目 Add Constraint 临床路径项目_UQ_项目序号 Unique (阶段ID,分类,项目序号) Deferrable Initially Deferred Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table 临床路径项目 Add Constraint 临床路径项目_UQ_项目内容 Unique (阶段ID,项目内容) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table 临床路径项目 Add Constraint 临床路径项目_FK_版本号 Foreign Key (路径ID,版本号) References 临床路径版本(路径ID,版本号) On Delete Cascade;
Alter Table 临床路径项目 Add Constraint 临床路径项目_FK_阶段ID Foreign Key (阶段ID) References 临床路径阶段(ID) On Delete Cascade;
Alter Table 临床路径项目 Add Constraint 临床路径项目_FK_图标ID Foreign Key (图标ID) References 临床路径图标(ID);
Create Index 临床路径项目_IX_版本号 On 临床路径项目(路径ID,版本号) Pctfree 5 Tablespace zl9indexcis
/
Create Index 临床路径项目_IX_阶段ID On 临床路径项目(阶段ID) Pctfree 5 Tablespace zl9indexcis
/
Create Index 临床路径项目_IX_图标ID On 临床路径项目(图标ID) Pctfree 5 Tablespace zl9indexcis
/


Create Sequence 路径医嘱内容_ID Start With 1;
CREATE TABLE 路径医嘱内容(
		ID NUMBER(18),
    相关ID NUMBER(18),
    序号 NUMBER(5),
    期效 NUMBER(1),
    诊疗项目ID NUMBER(18),
		收费细目ID NUMBER(18),
		医嘱内容 VARCHAR2(1000),
		单次用量 NUMBER(16,5),
		总给予量 NUMBER(16,5),
		标本部位 VARCHAR2(60),
		检查方法 VARCHAR2(30),
		医生嘱托 VARCHAR2(1000),
		执行频次 VARCHAR2(20),
		频率次数 NUMBER(3),
		频率间隔 NUMBER(3),
		间隔单位 VARCHAR2(4),
		执行性质 NUMBER(1),
		执行科室ID NUMBER(18),
		时间方案 VARCHAR2(50))
    TABLESPACE zl9BaseItem
    PCTFREE 5
    PCTUSED 85;
Alter Table 路径医嘱内容 Add Constraint 路径医嘱内容_PK Primary Key (ID) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table 路径医嘱内容 Add Constraint 路径医嘱内容_FK_相关ID Foreign Key (相关ID) References 路径医嘱内容(ID) Deferrable Initially Deferred;
Alter Table 路径医嘱内容 Add Constraint 路径医嘱内容_FK_诊疗项目ID Foreign Key (诊疗项目ID) References 诊疗项目目录(ID);
Alter Table 路径医嘱内容 Add Constraint 路径医嘱内容_FK_收费细目ID Foreign Key (收费细目ID) References 收费项目目录(ID);
Alter Table 路径医嘱内容 Add Constraint 路径医嘱内容_FK_执行科室ID Foreign Key (执行科室ID) References 部门表(ID);
Create Index 路径医嘱内容_IX_相关ID On 路径医嘱内容(相关ID) Pctfree 5 Tablespace zl9indexcis
/
Create Index 路径医嘱内容_IX_诊疗项目ID On 路径医嘱内容(诊疗项目ID) Pctfree 5 Tablespace zl9indexcis
/
Create Index 路径医嘱内容_IX_收费细目ID On 路径医嘱内容(收费细目ID) Pctfree 5 Tablespace zl9indexcis
/
Create Index 路径医嘱内容_IX_执行科室ID On 路径医嘱内容(执行科室ID) Pctfree 5 Tablespace zl9indexcis
/


CREATE TABLE 临床路径医嘱(
		路径项目ID NUMBER(18),
    医嘱内容ID NUMBER(18))
    TABLESPACE zl9BaseItem
    PCTFREE 5
    PCTUSED 85;
Alter Table 临床路径医嘱 Add Constraint 临床路径医嘱_PK Primary Key (路径项目ID,医嘱内容ID) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table 临床路径医嘱 Add Constraint 临床路径医嘱_FK_路径项目ID Foreign Key (路径项目ID) References 临床路径项目(ID) On Delete Cascade;
Alter Table 临床路径医嘱 Add Constraint 临床路径医嘱_FK_医嘱内容ID Foreign Key (医嘱内容ID) References 路径医嘱内容(ID) On Delete Cascade;


CREATE TABLE 临床路径病历(
		项目ID NUMBER(18),
    文件ID NUMBER(18))
    TABLESPACE zl9BaseItem
    PCTFREE 5
    PCTUSED 85;
Alter Table 临床路径病历 Add Constraint 临床路径病历_PK Primary Key (项目ID,文件ID) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table 临床路径病历 Add Constraint 临床路径病历_FK_项目ID Foreign Key (项目ID) References 临床路径项目(ID) On Delete Cascade;
Alter Table 临床路径病历 Add Constraint 临床路径病历_FK_文件ID Foreign Key (文件ID) References 病历文件列表(ID) On Delete Cascade;

Create Sequence 临床路径评估_ID Start With 1;
CREATE TABLE 临床路径评估(
		ID NUMBER(18),
    路径ID NUMBER(18),
    版本号 NUMBER(3),
		阶段ID NUMBER(18),
		评估类型 NUMBER(1))
    TABLESPACE zl9BaseItem
    PCTFREE 5
    PCTUSED 85;
Alter Table 临床路径评估 Add Constraint 临床路径评估_PK Primary Key (ID) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table 临床路径评估 Add Constraint 临床路径评估_UQ_评估类型 Unique (路径ID,版本号,阶段ID,评估类型) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table 临床路径评估 Add Constraint 临床路径评估_FK_版本号 Foreign Key (路径ID,版本号) References 临床路径版本(路径ID,版本号) On Delete Cascade;
Alter Table 临床路径评估 Add Constraint 临床路径评估_FK_阶段ID Foreign Key (阶段ID) References 临床路径阶段(ID) On Delete Cascade;


Create Sequence 路径评估指标_ID Start With 1;
CREATE TABLE 路径评估指标(
		ID NUMBER(18),
    评估ID NUMBER(18),
    序号 NUMBER(5),
		评估指标 VARCHAR2(200),
		指标类型 NUMBER(1),
		指标结果 VARCHAR2(500))
    TABLESPACE zl9BaseItem
    PCTFREE 5
    PCTUSED 85;
Alter Table 路径评估指标 Add Constraint 路径评估指标_PK Primary Key (ID) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table 路径评估指标 Add Constraint 路径评估指标_UQ_序号 Unique (评估ID,序号) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table 路径评估指标 Add Constraint 路径评估指标_UQ_评估指标 Unique (评估ID,评估指标) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table 路径评估指标 Add Constraint 路径评估指标_FK_评估ID Foreign Key (评估ID) References 临床路径评估(ID) On Delete Cascade;


CREATE TABLE 路径评估条件(
		评估ID NUMBER(18),
    指标ID NUMBER(18),
    项目ID NUMBER(18),
		关系式 VARCHAR2(5),
		条件值 VARCHAR2(50),
		条件组合 NUMBER(1))
    TABLESPACE zl9BaseItem
    PCTFREE 5
    PCTUSED 85;
Alter Table 路径评估条件 Add Constraint 路径评估条件_UQ_条件 Unique (指标ID,项目ID,关系式,条件值) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table 路径评估条件 Add Constraint 路径评估条件_FK_评估ID Foreign Key (评估ID) References 临床路径评估(ID) On Delete Cascade;
Alter Table 路径评估条件 Add Constraint 路径评估条件_FK_指标ID Foreign Key (指标ID) References 路径评估指标(ID) On Delete Cascade;
Alter Table 路径评估条件 Add Constraint 路径评估条件_FK_项目ID Foreign Key (项目ID) References 临床路径项目(ID) On Delete Cascade;
Create Index 路径评估条件_IX_评估ID On 路径评估条件(评估ID) Pctfree 5 Tablespace zl9indexcis
/


Create Sequence 病人临床路径_ID Start With 1;
CREATE TABLE 病人临床路径(
		ID NUMBER(18),
    病人ID NUMBER(18),
    主页ID NUMBER(5),
		科室ID NUMBER(18),
		路径ID NUMBER(18),
		版本号 NUMBER(3),
		导入人 VARCHAR2(20),
		导入时间 DATE,
		导入说明 VARCHAR2(1000),
		结束时间 DATE,
		状态 NUMBER(1),
		当前天数   NUMBER(18),
		当前阶段ID NUMBER(18),
		前一阶段ID NUMBER(18))
    TABLESPACE zl9CISRec
    PCTFREE 5
    PCTUSED 85;
Alter Table 病人临床路径 Add Constraint 病人临床路径_PK Primary Key (ID) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table 病人临床路径 Add Constraint 病人临床路径_FK_主页ID Foreign Key (病人ID,主页ID) References 病案主页(病人ID,主页ID);
Alter Table 病人临床路径 Add Constraint 病人临床路径_FK_科室ID Foreign Key (科室ID) References 部门表(ID);
Alter Table 病人临床路径 Add Constraint 病人临床路径_FK_版本号 Foreign Key (路径ID,版本号) References 临床路径版本(路径ID,版本号);
Create Index 病人临床路径_IX_病人ID On 病人临床路径(病人ID,主页ID) Pctfree 5 Tablespace zl9indexcis
/
Create Index 病人临床路径_IX_科室ID On 病人临床路径(科室ID) Pctfree 5 Tablespace zl9indexcis
/
Create Index 病人临床路径_IX_路径ID On 病人临床路径(路径ID,版本号) Pctfree 5 Tablespace zl9indexcis
/
Create Index 病人临床路径_IX_导入时间 On 病人临床路径(导入时间) Pctfree 5 Tablespace zl9indexcis
/


Create Sequence 病人路径执行_ID Start With 1;
CREATE TABLE 病人路径执行(
		ID NUMBER(18),
		路径记录ID NUMBER(18),
		阶段ID NUMBER(18),		
		日期 DATE,
		天数 NUMBER(5),
		分类 VARCHAR2(50),
		项目ID NUMBER(18),
		项目序号 NUMBER(5),
		项目内容 VARCHAR2(1000),
		执行者 NUMBER(1),
		项目结果 VARCHAR2(500),
		添加原因 VARCHAR2(1000),
		图标ID NUMBER(18),
		执行人 VARCHAR2(20),
		执行时间 DATE,
		执行结果 VARCHAR2(50),
		执行说明 VARCHAR2(200),
		登记人 VARCHAR2(20),
		登记时间 DATE)
    TABLESPACE zl9CISRec
    PCTFREE 5
    PCTUSED 85;
Alter Table 病人路径执行 Add Constraint 病人路径执行_PK Primary Key (ID) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table 病人路径执行 Add Constraint 病人路径执行_UQ_项目内容 Unique (路径记录ID,阶段ID,日期,项目ID,项目内容) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table 病人路径执行 Add Constraint 病人路径执行_FK_路径记录ID Foreign Key (路径记录ID) References 病人临床路径(ID);
Alter Table 病人路径执行 Add Constraint 病人路径执行_FK_阶段ID Foreign Key (阶段ID) References 临床路径阶段(ID);
Alter Table 病人路径执行 Add Constraint 病人路径执行_FK_项目ID Foreign Key (项目ID) References 临床路径项目(ID);
Alter Table 病人路径执行 Add Constraint 病人路径执行_FK_图标ID Foreign Key (图标ID) References 临床路径图标(ID);
Create Index 病人路径执行_IX_日期 On 病人路径执行(日期) Pctfree 5 Tablespace zl9indexcis
/
Create Index 病人路径执行_IX_路径记录ID On 病人路径执行(路径记录ID) Pctfree 5 Tablespace zl9indexcis
/
Create Index 病人路径执行_IX_阶段ID On 病人路径执行(阶段ID) Pctfree 5 Tablespace zl9indexcis
/
Create Index 病人路径执行_IX_项目ID On 病人路径执行(项目ID) Pctfree 5 Tablespace zl9indexcis
/
Create Index 病人路径执行_IX_图标ID On 病人路径执行(图标ID) Pctfree 5 Tablespace zl9indexcis
/
Create Index 病人路径执行_IX_登记时间 On 病人路径执行(登记时间) Pctfree 5 Tablespace zl9indexcis
/


CREATE TABLE 病人路径评估(
		路径记录ID NUMBER(18),
		阶段ID NUMBER(18),
		日期 DATE,
    天数 NUMBER(5),
		评估人 VARCHAR2(50),
		评估时间 DATE,
		评估结果 NUMBER(2),
		评估说明 VARCHAR2(1000),
		登记人 VARCHAR2(20),
		登记时间 DATE)
    TABLESPACE zl9CISRec
    PCTFREE 5
    PCTUSED 85;
Alter Table 病人路径评估 Add Constraint 病人路径评估_PK Primary Key (路径记录ID,阶段ID,日期) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table 病人路径评估 Add Constraint 病人路径评估_FK_阶段ID Foreign Key (阶段ID) References 临床路径阶段(ID);
Alter Table 病人路径评估 Add Constraint 病人路径评估_FK_路径记录ID Foreign Key (路径记录ID) References 病人临床路径(ID);
Create Index 病人路径评估_IX_日期 On 病人路径评估(日期) Pctfree 5 Tablespace zl9indexcis
/
Create Index 病人路径评估_IX_登记时间 On 病人路径评估(登记时间) Pctfree 5 Tablespace zl9indexcis
/


CREATE TABLE 病人路径指标(
		路径记录ID NUMBER(18),
		阶段ID NUMBER(18),
		日期 DATE,
    天数 NUMBER(5),
		评估类型 NUMBER(1),
		评估指标 VARCHAR2(50),
		指标类型 NUMBER(1),
		指标结果 VARCHAR2(50))
    TABLESPACE zl9CISRec
    PCTFREE 5
    PCTUSED 85;
Alter Table 病人路径指标 Add Constraint 病人路径指标_UQ_评估指标 Unique (路径记录ID,阶段ID,日期,评估指标) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table 病人路径指标 Add Constraint 病人路径指标_FK_阶段ID Foreign Key (阶段ID) References 临床路径阶段(ID);
Alter Table 病人路径指标 Add Constraint 病人路径指标_FK_路径记录ID Foreign Key (路径记录ID) References 病人临床路径(ID);
Create Index 病人路径指标_IX_日期 On 病人路径指标(日期) Pctfree 5 Tablespace zl9indexcis
/

CREATE TABLE 病人路径医嘱(
		路径执行ID NUMBER(18),
    病人医嘱ID NUMBER(18))
    TABLESPACE zl9CISRec
    PCTFREE 5
    PCTUSED 85;
Alter Table 病人路径医嘱 Add Constraint 病人路径医嘱_PK Primary Key (路径执行ID,病人医嘱ID) Using Index Pctfree 5 Tablespace zl9indexcis;
Alter Table 病人路径医嘱 Add Constraint 病人路径医嘱_FK_路径执行ID Foreign Key (路径执行ID) References 病人路径执行(ID);
Alter Table 病人路径医嘱 Add Constraint 病人路径医嘱_FK_病人医嘱ID Foreign Key (病人医嘱ID) References 病人医嘱记录(ID);


--对原电子病历记录的更改
Alter Table 电子病历记录 Add 路径执行ID Number(18);
Alter Table 电子病历记录 Add Constraint 电子病历记录_FK_路径执行ID Foreign Key (路径执行ID) References 病人路径执行(ID);
Create Index 电子病历记录_IX_路径执行ID On 电子病历记录(路径执行ID) Pctfree 5 Tablespace zl9indexcis
/

--------------------------------------------------------------------------------------------------------------------------
--临床路径数据内容部分
--------------------------------------------------------------------------------------------------------------------------
Insert Into zlBaseCode(系统,表名,固定,说明,分类) Values(100,'临床病例分型',1,'临床路径根据其对应病情的复杂、紧急程度，以及实施路径的难易程度进行的一个划分标准。','医疗工作');
Insert Into zlBaseCode(系统,表名,固定,说明,分类) Values(100,'路径常见结果',0,'临床路径项目执行时的常见结果','医疗工作');
Insert Into zlBaseCode(系统,表名,固定,说明,分类) Values(100,'变异常见原因',0,'临床路径常见的变异原因','医疗工作');

Insert Into 临床病例分型(编码,名称,简码)
	Select 'A','单纯普通型','DCPTX' From Dual Union ALL
	Select 'B','单纯急症型','DCJZX' From Dual Union ALL
	Select 'C','复杂疑难型','FZYNX' From Dual Union ALL
	Select 'D','复杂危重型','FZWZX' From Dual;

Insert Into zlStreamTabs(System_NO,Table_Name,Dml_Handle,Repeat_Way,Fixation)
Select 100,'临床病例分型',0,2,1 From Dual Union All
Select 100,'路径常见结果',0,2,1 From Dual Union All
Select 100,'变异常见原因',0,2,1 From Dual Union All
Select 100,'临床路径图标',0,2,1 From Dual Union All
Select 100,'临床路径目录',0,2,1 From Dual Union All
Select 100,'临床路径文件',0,2,1 From Dual Union All
Select 100,'临床路径病种',0,2,1 From Dual Union All
Select 100,'临床路径科室',0,2,1 From Dual Union All
Select 100,'临床路径版本',0,2,1 From Dual Union All
Select 100,'临床路径阶段',0,2,1 From Dual Union All
Select 100,'临床路径分类',0,2,1 From Dual Union All
Select 100,'临床路径项目',0,2,1 From Dual Union All
Select 100,'路径医嘱内容',0,2,1 From Dual Union All
Select 100,'临床路径医嘱',0,2,1 From Dual Union All
Select 100,'临床路径病历',0,2,1 From Dual Union All
Select 100,'临床路径评估',0,2,1 From Dual Union All
Select 100,'路径评估指标',0,2,1 From Dual Union All
Select 100,'路径评估条件',0,2,1 From Dual;

Insert Into zlStreamTabs(System_NO,Table_Name,Dml_Handle,Repeat_Way,Fixation)
Select 100,'病人临床路径',0,3,1 From Dual Union All
Select 100,'病人路径执行',0,3,1 From Dual Union All
Select 100,'病人路径评估',0,3,1 From Dual Union ALL
Select 100,'病人路径指标',0,3,1 From Dual Union ALL
Select 100,'病人路径医嘱',0,3,1 From Dual;

Insert Into zlBakTables(系统,表名)
Select 100,'病人临床路径' From Dual Union ALL
Select 100,'病人路径执行' From Dual Union ALL
Select 100,'病人路径评估' From Dual Union ALL
Select 100,'病人路径指标' From Dual Union ALL
Select 100,'病人路径医嘱' From Dual;

Insert Into zlPrograms(序号,标题,说明,系统,部件) Values(1078,'临床路径管理','对临床路径的基本信息、路径表信息，及版本变化进行定义、维护',100,'zl9CISJob');
Insert Into zlPrograms(序号,标题,说明,系统,部件) Values(1256,'临床路径应用','病人临床路径的导入，生成，执行，评估等功能应用',100,'zl9CISJob');
Insert Into zlPrograms(序号,标题,说明,系统,部件) Values(1275,'临床路径跟踪','对各个临床路径的应用情况和明细进行查阅，跟踪',100,'zl9CISJob');

Insert Into zlProgFuncs(系统,序号,功能) Values(100,1078,'基本');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1078,'增删改',1,'对临床路径进行增加、修改、删除，应用范围等基本信息维护的权限');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1078,'导入XML',2,'从XML文件导入临床路径的权限');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1078,'导出XML',3,'将临床路径导出到XML文件的权限');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1078,'审核',4,'对制定好的临床路径，进行审核生效的权限；具有该权限同时可以取消审核');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1078,'停用',5,'对已经审核应用的临床路径进行停用的权限；具有该权限同时可以启用已经停用的路径');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1078,'路径表设计',6,'对临床路径表的分类，时间阶段，项目，版本等信息进行设计定义的权限');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1078,'评估表设计',7,'对临床路径的导入评估表单，阶段评估表单，结束评估表单进行设计定义的权限');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1078,'全院路径',8,'对全院的临床路径进行管理的权限，不具有该权限只能管理本科的临床路径');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1078,'图标设置',9,'对临床路径表中项目可以对应的图标进行设置增删的权限');

Insert Into zlProgFuncs(系统,序号,功能) Values(100,1256,'基本');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1256,'导入路径',1,'对新入院或新入科病人进行评估并导入适用临床路径的权限');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1256,'生成路径',2,'对病人进行路径项目生成的权限，必须具备医嘱下达和病历生成的权限才能生成对应项目');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1256,'执行路径',3,'对病人临床路径的内容进行执行的权限，必须具备医嘱停止权限才能进行批量执行');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1256,'阶段评估',4,'对病人临床路径的每个阶段进行评估的权限');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1256,'结束路径',5,'自动或人为完成临床路径执行的权限');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1256,'路径外项目',6,'在临床路径设计定义好的范围之外添加路径项目的权限');

Insert Into zlProgFuncs(系统,序号,功能) Values(100,1275,'基本');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1275,'全院路径',1,'对全院的临床路径进行跟踪的权限，不具有该权限只能跟踪本科的临床路径');

--- 临床路径管理［1］
Insert Into zlProgRelas(系统,序号,功能,组号,关系,主项,主项关系) Values(100,1078,'路径表设计',1,2,1,0);
Insert Into zlProgRelas(系统,序号,功能,组号,关系,主项,主项关系) Values(100,1078,'评估表设计',1,2,0,0);

--1078:临床路径管理(基本)
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'部门表','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'部门性质说明','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'人员表','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'人员性质说明','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'部门人员','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'上机人员表','SELECT');

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'病情','SELECT');

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'临床病例分型','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'路径常见结果','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'变异常见原因','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'临床路径图标','SELECT');

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'临床路径目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'临床路径病种','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'临床路径科室','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'临床路径文件','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'临床路径版本','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'临床路径分类','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'临床路径阶段','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'临床路径项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'临床路径评估','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'路径评估指标','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'路径评估条件','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'临床路径医嘱','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'路径医嘱内容','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'临床路径病历','SELECT');

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'病历文件列表','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'疾病编码目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'疾病诊断目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'疾病诊断别名','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'诊疗项目目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'收费项目目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'药品规格','SELECT');

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'诊疗用法用量','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'诊疗频率项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'诊疗项目别名','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'诊疗分类目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'诊疗执行科室','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'诊疗个人项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'诊疗项目组合','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'诊疗检查部位','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'诊疗检验标本','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'收费项目别名','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'收费执行科室','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'材料特性','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'药品特性','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'常用嘱托','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'医嘱内容定义','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'中药煎服脚注','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'检验项目参考','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'检验报告项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'病人信息','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'病案主页','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'病人挂号记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'病人医嘱记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'病人新生儿记录','SELECT');

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'基本',user,'Zl_Lob_Read','EXECUTE');

--1078:临床路径管理(增删改)
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'增删改',user,'Zl_Lob_Append','EXECUTE');

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'增删改',user,'Zl_临床路径目录_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'增删改',user,'Zl_临床路径目录_Update','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'增删改',user,'Zl_临床路径目录_Delete','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'增删改',user,'Zl_临床路径文件_Delete','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'增删改',user,'Zl_临床路径文件_Insert','EXECUTE');

--1078:临床路径管理(导入XML)
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'导入XML',user,'zl_临床路径目录_Update','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'导入XML',user,'zl_临床路径目录_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'导入XML',user,'Zl_临床路径版本_Delete','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'导入XML',user,'Zl_临床路径版本_Update','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'导入XML',user,'Zl_路径评估指标_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'导入XML',user,'Zl_路径评估条件_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'导入XML',user,'Zl_路径医嘱内容_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'导入XML',user,'Zl_临床路径分类_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'导入XML',user,'Zl_临床路径阶段_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'导入XML',user,'Zl_临床路径项目_Insert','EXECUTE');

--1078:临床路径管理(审核)
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'审核',user,'Zl_临床路径版本_Audit','EXECUTE');

--1078:临床路径管理(停用)
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'停用',user,'Zl_临床路径版本_Stop','EXECUTE');

--1078:临床路径管理(路径表设计)
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'路径表设计',user,'Zl_临床路径版本_Delete','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'路径表设计',user,'Zl_临床路径版本_Copy','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'路径表设计',user,'Zl_临床路径版本_Update','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'路径表设计',user,'Zl_路径评估指标_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'路径表设计',user,'Zl_路径评估条件_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'路径表设计',user,'Zl_临床路径阶段_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'路径表设计',user,'Zl_临床路径阶段_Delete','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'路径表设计',user,'Zl_临床路径阶段_Update','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'路径表设计',user,'Zl_临床路径分类_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'路径表设计',user,'Zl_路径医嘱内容_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'路径表设计',user,'Zl_临床路径项目_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'路径表设计',user,'Zl_临床路径项目_Delete','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'路径表设计',user,'Zl_临床路径项目_Update','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'路径表设计',user,'Zl_GetPathCharge','EXECUTE');

--1078:临床路径管理(图标设置)
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'图标设置',user,'Zl_Lob_Append','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'图标设置',user,'Zl_临床路径图标_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1078,'图标设置',user,'Zl_临床路径图标_Delete','EXECUTE');

--1275:临床路径跟踪(基本)
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'部门表','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'部门性质说明','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'人员表','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'人员性质说明','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'部门人员','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'上机人员表','SELECT');

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'临床路径图标','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'临床路径目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'临床路径病种','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'临床路径科室','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'临床路径文件','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'临床路径版本','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'临床路径分类','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'临床路径阶段','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'临床路径项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'临床路径评估','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'路径评估指标','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'路径评估条件','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'临床路径医嘱','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'路径医嘱内容','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'临床路径病历','SELECT');

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'病历文件列表','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'诊疗项目目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'收费项目目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'药品规格','SELECT');

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'病人信息','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'病案主页','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'病人临床路径','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'病人路径执行','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'病人路径评估','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'病人路径指标','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'病人路径医嘱','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'病人医嘱记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'病人医嘱状态','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'病人医嘱报告','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'病人医嘱发送','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'电子病历记录','SELECT');

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1275,'基本',user,'Zl_Lob_Read','EXECUTE');


--1256:临床路径应用(基本)
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'基本',user,'临床路径目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'基本',user,'临床路径版本','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'基本',user,'临床路径阶段','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'基本',user,'临床路径分类','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'基本',user,'临床路径项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'基本',user,'病人路径执行','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'基本',user,'病人路径评估','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'基本',user,'临床路径医嘱','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'基本',user,'临床路径病历','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'基本',user,'路径医嘱内容','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'基本',user,'病历文件列表','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'基本',user,'临床路径图标','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'基本',user,'临床路径评估','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'基本',user,'路径评估指标','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'基本',user,'路径评估条件','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'基本',user,'病人临床路径','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'基本',user,'病人路径指标','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'基本',user,'病案主页','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'基本',user,'病人信息','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'基本',user,'病人变动记录','SELECT');


--1256:临床路径应用(导入路径)
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'导入路径',user,'临床路径科室','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'导入路径',user,'临床路径病种','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'导入路径',user,'病人诊断记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'导入路径',user,'Zl_病人路径导入_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'导入路径',user,'Zl_病人路径导入_Delete','EXECUTE');

--1256:临床路径应用(生成路径)
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'生成路径',user,'病人路径医嘱','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'生成路径',user,'电子病历记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'生成路径',user,'病人新生儿记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'生成路径',user,'Zl_病人路径生成_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'生成路径',user,'Zl_病人路径生成_Delete','EXECUTE');

--1256:临床路径应用(路径外项目)
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'路径外项目',user,'路径常见结果','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'路径外项目',user,'Zl_病人路径生成_Insert','EXECUTE');

--1256:临床路径应用(阶段评估)
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'阶段评估',user,'变异常见原因','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'阶段评估',user,'保险模拟结算','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'阶段评估',user,'病人余额','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'阶段评估',user,'Zl_病人路径评估_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'阶段评估',user,'Zl_病人路径评估_Delete','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'阶段评估',user,'Zl_Getpathcharge','EXECUTE');

--1256:临床路径应用(执行路径)
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'执行路径',user,'病人路径医嘱','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'执行路径',user,'Zl_病人路径执行_Update','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'执行路径',user,'Zl_病人路径执行_Delete','EXECUTE');

--1256:临床路径应用(结束路径)
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'结束路径',user,'Zl_病人路径结束_Update','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1256,'结束路径',user,'Zl_病人路径结束_Delete','EXECUTE');

--Insert Into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) Values('缺省',zlMenus_id.nextval,   zlMenus_id.nextval-20,'医护基础项目','医护基础','D',99,'建立疾病诊断与诊疗措施应用的相关基础。',100,NULL);
--Insert Into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) Values('缺省',zlMenus_id.nextval,      zlMenus_id.nextval-10,'临床路径管理','路径管理','J',99,'对临床路径的基本信息、路径表信息，及版本变化进行定义、维护',100,1078);

--Insert Into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) Values('缺省',zlMenus_id.nextval,null,'病历资料检索','病历检索','E',99,'相关资料与病历检索查询',100,NULL);
--Insert Into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) Values('缺省',zlMenus_id.nextval,   zlMenus_id.nextval-5,'临床路径跟踪','路径跟踪','G',129,'对各个临床路径的应用情况和明细进行查阅，跟踪',100,1275);

Insert Into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) 
	Select '缺省',zlMenus_id.nextval,ID,'临床路径管理','路径管理','J',99,'对临床路径的基本信息、路径表信息，及版本变化进行定义、维护',100,1078 
	From zlMenus Where 组别='缺省' And 标题='医护基础项目' And 系统=100 And 模块 Is Null;

Insert Into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) 
	Select '缺省',zlMenus_id.nextval,ID,'临床路径跟踪','路径跟踪','G',129,'对各个临床路径的应用情况和明细进行查阅，跟踪',100,1275
	From zlMenus Where 组别='缺省' And 标题='病历资料检索' And 系统=100 And 模块 Is Null;

--------------------------------------------------------------------------------------------------------------------------
--临床路径存储过程部分
--------------------------------------------------------------------------------------------------------------------------
Create Or Replace Procedure Zl_临床路径目录_Insert
(
  分类_In     临床路径目录.分类%Type,
  编码_In     临床路径目录.编码%Type,
  名称_In     临床路径目录.名称%Type,
  说明_In     临床路径目录.说明%Type,
  病例分型_In 临床路径目录.病例分型%Type,
  适用病情_In 临床路径目录.适用病情%Type,
  适用性别_In 临床路径目录.适用性别%Type,
  适用年龄_In 临床路径目录.适用年龄%Type,
  通用_In     临床路径目录.通用%Type,
  科室ids_In  Varchar2 := Null,
  病种ids_In  Varchar2 := Null,
  路径id_In   临床路径目录.Id%Type := Null
  --参数：
  --科室IDs_IN：当为指定科室应用时传入，格式为"科室ID1,科室ID2,..."
  --病种IDs_IN：传入格式为"疾病ID1,疾病ID2,...;诊断ID1,诊断ID2,..."
  --路径id_In：是否由外部确定新的ID
) Is
  v_路径id 临床路径目录.Id%Type;

  v_参数串 Varchar2(4000);
  v_当前id Number(18);

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  If 路径id_In Is Not Null Then
    v_路径id := 路径id_In;
  Else
    Select 临床路径目录_Id.Nextval Into v_路径id From Dual;
  End If;

  Insert Into 临床路径目录
    (ID, 分类, 编码, 名称, 说明, 病例分型, 适用病情, 适用性别, 适用年龄, 通用)
  Values
    (v_路径id, 分类_In, 编码_In, 名称_In, 说明_In, 病例分型_In, 适用病情_In, 适用性别_In, 适用年龄_In, 通用_In);

  If 通用_In = 2 And 科室ids_In Is Not Null Then
    v_参数串 := 科室ids_In || ',';
    While v_参数串 Is Not Null Loop
      v_当前id := To_Number(Substr(v_参数串, 1, Instr(v_参数串, ',') - 1));
      v_参数串 := Substr(v_参数串, Instr(v_参数串, ',') + 1);
    
      Insert Into 临床路径科室 (路径id, 科室id) Values (v_路径id, v_当前id);
    End Loop;
  End If;

  If 病种ids_In Is Not Null Then
    v_参数串 := Substr(病种ids_In, 1, Instr(病种ids_In, ';') - 1);
    If v_参数串 Is Not Null Then
      v_参数串 := v_参数串 || ',';
      While v_参数串 Is Not Null Loop
        v_当前id := To_Number(Substr(v_参数串, 1, Instr(v_参数串, ',') - 1));
        v_参数串 := Substr(v_参数串, Instr(v_参数串, ',') + 1);
      
        Insert Into 临床路径病种 (路径id, 疾病id) Values (v_路径id, v_当前id);
      End Loop;
    End If;
  
    v_参数串 := Substr(病种ids_In, Instr(病种ids_In, ';') + 1);
    If v_参数串 Is Not Null Then
      v_参数串 := v_参数串 || ',';
      While v_参数串 Is Not Null Loop
        v_当前id := To_Number(Substr(v_参数串, 1, Instr(v_参数串, ',') - 1));
        v_参数串 := Substr(v_参数串, Instr(v_参数串, ',') + 1);
      
        Insert Into 临床路径病种 (路径id, 诊断id) Values (v_路径id, v_当前id);
      End Loop;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床路径目录_Insert;
/

Create Or Replace Procedure Zl_临床路径目录_Update
(
  路径id_In   临床路径目录.Id%Type,
  分类_In     临床路径目录.分类%Type,
  编码_In     临床路径目录.编码%Type,
  名称_In     临床路径目录.名称%Type,
  说明_In     临床路径目录.说明%Type,
  病例分型_In 临床路径目录.病例分型%Type,
  适用病情_In 临床路径目录.适用病情%Type,
  适用性别_In 临床路径目录.适用性别%Type,
  适用年龄_In 临床路径目录.适用年龄%Type,
  通用_In     临床路径目录.通用%Type,
  科室ids_In  Varchar2 := Null,
  病种ids_In  Varchar2 := Null
  --参数：
  --科室IDs_IN：当为指定科室应用时传入，格式为"科室ID1,科室ID2,..."
  --病种IDs_IN：传入格式为"疾病ID1,疾病ID2,...;诊断ID1,诊断ID2,..."
) Is
  v_参数串 Varchar2(4000);
  v_当前id Number(18);

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  Update 临床路径目录
  Set 分类 = 分类_In, 编码 = 编码_In, 名称 = 名称_In, 说明 = 说明_In, 病例分型 = 病例分型_In, 适用病情 = 适用病情_In, 适用性别 = 适用性别_In, 适用年龄 = 适用年龄_In,
      通用 = 通用_In
  Where ID = 路径id_In;

  Delete From 临床路径科室 Where 路径id = 路径id_In;
  If 通用_In = 2 And 科室ids_In Is Not Null Then
    v_参数串 := 科室ids_In || ',';
    While v_参数串 Is Not Null Loop
      v_当前id := To_Number(Substr(v_参数串, 1, Instr(v_参数串, ',') - 1));
      v_参数串 := Substr(v_参数串, Instr(v_参数串, ',') + 1);
    
      Insert Into 临床路径科室 (路径id, 科室id) Values (路径id_In, v_当前id);
    End Loop;
  End If;

  Delete From 临床路径病种 Where 路径id = 路径id_In;
  If 病种ids_In Is Not Null Then
    v_参数串 := Substr(病种ids_In, 1, Instr(病种ids_In, ';') - 1);
    If v_参数串 Is Not Null Then
      v_参数串 := v_参数串 || ',';
      While v_参数串 Is Not Null Loop
        v_当前id := To_Number(Substr(v_参数串, 1, Instr(v_参数串, ',') - 1));
        v_参数串 := Substr(v_参数串, Instr(v_参数串, ',') + 1);
      
        Insert Into 临床路径病种 (路径id, 疾病id) Values (路径id_In, v_当前id);
      End Loop;
    End If;
  
    v_参数串 := Substr(病种ids_In, Instr(病种ids_In, ';') + 1);
    If v_参数串 Is Not Null Then
      v_参数串 := v_参数串 || ',';
      While v_参数串 Is Not Null Loop
        v_当前id := To_Number(Substr(v_参数串, 1, Instr(v_参数串, ',') - 1));
        v_参数串 := Substr(v_参数串, Instr(v_参数串, ',') + 1);
      
        Insert Into 临床路径病种 (路径id, 诊断id) Values (路径id_In, v_当前id);
      End Loop;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床路径目录_Update;
/

Create Or Replace Procedure Zl_临床路径目录_Delete(路径id_In 临床路径目录.Id%Type) Is
  v_Count Number;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --如果存在已审核的版本，则不允许删除
  Select Count(*) Into v_Count From 临床路径版本 Where 路径id = 路径id_In And 审核时间 Is Not Null;
  If Nvl(v_Count, 0) > 0 Then
    v_Error := '该临床路径存在已经审核的路径表版本，不允许删除。';
    Raise Err_Custom;
  End If;

  Delete From 临床路径目录 Where ID = 路径id_In;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床路径目录_Delete;
/

Create Or Replace Procedure Zl_临床路径文件_Delete
(
  路径id_In 临床路径文件.路径id%Type,
  文件名_In 临床路径文件.文件名%Type
) Is
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  Delete From 临床路径文件 Where 路径id = 路径id_In And 文件名 = 文件名_In;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床路径文件_Delete;
/

Create Or Replace Procedure Zl_临床路径文件_Insert
(
  路径id_In 临床路径文件.路径id%Type,
  文件名_In 临床路径文件.文件名%Type
) Is
  v_Temp     Varchar2(255);
  v_人员姓名 病人医嘱状态.操作人员%Type;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  v_Temp     := Zl_Identity;
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);

  Insert Into 临床路径文件 (路径id, 文件名, 创建人, 创建时间) Values (路径id_In, 文件名_In, v_人员姓名, Sysdate);
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床路径文件_Insert;
/

Create Or Replace Procedure Zl_临床路径图标_Insert(图标id_In 临床路径图标.Id%Type) Is
Begin
  Insert Into 临床路径图标 (ID, 性质) Values (图标id_In, 0);
End Zl_临床路径图标_Insert;
/

Create Or Replace Procedure Zl_临床路径图标_Delete(图标id_In 临床路径图标.Id%Type) Is
Begin
  Delete From 临床路径图标 Where ID = 图标id_In;
End Zl_临床路径图标_Delete;
/

Create Or Replace Function Zl_Lob_Read
(
  Tab_In   In Number,
  Key_In   In Varchar2,
  Pos_In   In Number,
  Moved_In In Number := 0
  --参数说明： 
  --Tab_In：包含LOB的数据表
  --        0-病历标记图形;1-病历文件格式;2-病历文件图形;3-病历范文格式;4-病历范文图形; 
  --        5-电子病历格式;6-电子病历图形;7-病历页面格式；8-电子病历附件;9-体温重叠标记 
  --        10-临床路径文件,11-临床路径图标
  --Key_In：数据记录的关键字 
  --Pos_In：从0开始不断读取，直到返回为空
  --Moved_In: 0正常记录,1读取转储后备表记录 
) Return Varchar2 Is
  l_Blob   Blob;
  v_Buffer Varchar2(32767);
  n_Amount Number := 2000;
  n_Offset Number := 1;
Begin
  If Tab_In = 0 Then
    Select 图形 Into l_Blob From 病历标记图形 Where 编码 = Key_In;
  Elsif Tab_In = 1 Then
    Select 内容 Into l_Blob From 病历文件格式 Where 文件id = To_Number(Key_In);
  Elsif Tab_In = 2 Then
    Select 图形 Into l_Blob From 病历文件图形 Where 对象id = To_Number(Key_In);
  Elsif Tab_In = 3 Then
    Select 内容 Into l_Blob From 病历范文格式 Where 文件id = To_Number(Key_In);
  Elsif Tab_In = 4 Then
    Select 图形 Into l_Blob From 病历范文图形 Where 对象id = To_Number(Key_In);
  Elsif Tab_In = 5 Then
    If Moved_In = 0 Then
      Select 内容 Into l_Blob From 电子病历格式 Where 文件id = To_Number(Key_In);
    Else
      Select 内容 Into l_Blob From H电子病历格式 Where 文件id = To_Number(Key_In);
    End If;
  Elsif Tab_In = 6 Then
    If Moved_In = 0 Then
      Select 图形 Into l_Blob From 电子病历图形 Where 对象id = To_Number(Key_In);
    Else
      Select 图形 Into l_Blob From H电子病历图形 Where 对象id = To_Number(Key_In);
    End If;
  Elsif Tab_In = 7 Then
    Select 图形
    Into l_Blob
    From 病历页面格式
    Where 种类 = To_Number(Substr(Key_In, 1, 1)) And 编号 = Substr(Key_In, 3);
  Elsif Tab_In = 8 Then
    If Moved_In = 0 Then
      Select 内容
      Into l_Blob
      From 电子病历附件
      Where 病历id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And 序号 = Substr(Key_In, Instr(Key_In, ',') + 1);
    Else
      Select 内容
      Into l_Blob
      From H电子病历附件
      Where 病历id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And 序号 = Substr(Key_In, Instr(Key_In, ',') + 1);
    End If;
  Elsif Tab_In = 9 Then
    Select 标记图形 Into l_Blob From 体温重叠标记 Where 序号 = To_Number(Key_In);
  Elsif Tab_In = 10 Then
    Select 内容
    Into l_Blob
    From 临床路径文件
    Where 路径id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And 文件名 = Substr(Key_In, Instr(Key_In, ',') + 1);
  Elsif Tab_In = 11 Then
    Select 图标 Into l_Blob From 临床路径图标 Where ID = To_Number(Key_In);
  End If;

  n_Offset := n_Offset + Pos_In * n_Amount;
  Dbms_Lob.Read(l_Blob, n_Amount, n_Offset, v_Buffer);
  Return v_Buffer;
Exception
  When No_Data_Found Then
    Return Null;
End Zl_Lob_Read;
/

Create Or Replace Procedure Zl_Lob_Append
(
  Tab_In In Number,
  Key_In In Varchar2,
  Txt_In In Varchar2, --16进制的文件片段或文字片段
  Cls_In In Number := 0 --是否清除原来的内容，第一片段传递时为1，以后为0
  --参数说明：
  --Tab_In：包含LOB的数据表
  --        0-病历标记图形;1-病历文件格式;2-病历文件图形;3-病历范文格式;4-病历范文图形;
  --        5-电子病历格式;6-电子病历图形;7-病历页面格式；8-电子病历附件;9-体温重叠标记
  --        10-临床路径文件,11-临床路径图标
  --Key_In：数据记录的关键字
  --Txt_In：16进制的文件片段或文字片段
  --Cls_In：是否清除原来的内容，第一片段传递时为1，以后为0
) Is
  l_Blob Blob;
Begin
  If Tab_In = 0 Then
    If Cls_In = 1 Then
      Update 病历标记图形 Set 图形 = Empty_Blob() Where 编码 = Key_In;
    End If;
    Select 图形 Into l_Blob From 病历标记图形 Where 编码 = Key_In For Update;
  Elsif Tab_In = 1 Then
    If Cls_In = 1 Then
      Update 病历文件格式 Set 内容 = Empty_Blob() Where 文件id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into 病历文件格式 (文件id, 内容) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select 内容 Into l_Blob From 病历文件格式 Where 文件id = To_Number(Key_In) For Update;
  Elsif Tab_In = 2 Then
    If Cls_In = 1 Then
      Update 病历文件图形 Set 图形 = Empty_Blob() Where 对象id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into 病历文件图形 (对象id, 图形) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select 图形 Into l_Blob From 病历文件图形 Where 对象id = To_Number(Key_In) For Update;
  Elsif Tab_In = 3 Then
    If Cls_In = 1 Then
      Update 病历范文格式 Set 内容 = Empty_Blob() Where 文件id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into 病历范文格式 (文件id, 内容) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select 内容 Into l_Blob From 病历范文格式 Where 文件id = To_Number(Key_In) For Update;
  Elsif Tab_In = 4 Then
    If Cls_In = 1 Then
      Update 病历范文图形 Set 图形 = Empty_Blob() Where 对象id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into 病历范文图形 (对象id, 图形) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select 图形 Into l_Blob From 病历范文图形 Where 对象id = To_Number(Key_In) For Update;
  Elsif Tab_In = 5 Then
    If Cls_In = 1 Then
      Update 电子病历格式 Set 内容 = Empty_Blob() Where 文件id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into 电子病历格式 (文件id, 内容) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select 内容 Into l_Blob From 电子病历格式 Where 文件id = To_Number(Key_In) For Update;
  Elsif Tab_In = 6 Then
    If Cls_In = 1 Then
      Update 电子病历图形 Set 图形 = Empty_Blob() Where 对象id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into 电子病历图形 (对象id, 图形) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select 图形 Into l_Blob From 电子病历图形 Where 对象id = To_Number(Key_In) For Update;
  Elsif Tab_In = 7 Then
    If Cls_In = 1 Then
      Update 病历页面格式
      Set 图形 = Empty_Blob()
      Where 种类 = To_Number(Substr(Key_In, 1, 1)) And 编号 = Substr(Key_In, 3);
    End If;
    Select 图形
    Into l_Blob
    From 病历页面格式
    Where 种类 = To_Number(Substr(Key_In, 1, 1)) And 编号 = Substr(Key_In, 3)
    For Update;
  Elsif Tab_In = 8 Then
    If Cls_In = 1 Then
      Update 电子病历附件
      Set 内容 = Empty_Blob()
      Where 病历id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And 序号 = Substr(Key_In, Instr(Key_In, ',') + 1);
    End If;
    Select 内容
    Into l_Blob
    From 电子病历附件
    Where 病历id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And 序号 = Substr(Key_In, Instr(Key_In, ',') + 1)
    For Update;
  Elsif Tab_In = 9 Then
    If Cls_In = 1 Then
      Update 体温重叠标记 Set 标记图形 = Empty_Blob() Where 序号 = To_Number(Key_In);
    End If;
    Select 标记图形 Into l_Blob From 体温重叠标记 Where 序号 = To_Number(Key_In) For Update;
  Elsif Tab_In = 10 Then
    If Cls_In = 1 Then
      Update 临床路径文件
      Set 内容 = Empty_Blob()
      Where 路径id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And
            文件名 = Substr(Key_In, Instr(Key_In, ',') + 1);
    End If;
    Select 内容
    Into l_Blob
    From 临床路径文件
    Where 路径id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And 文件名 = Substr(Key_In, Instr(Key_In, ',') + 1)
    For Update;
  Elsif Tab_In = 11 Then
    If Cls_In = 1 Then
      Update 临床路径图标 Set 图标 = Empty_Blob() Where ID = To_Number(Key_In);
    End If;
    Select 图标 Into l_Blob From 临床路径图标 Where ID = To_Number(Key_In) For Update;
  End If;
  Dbms_Lob.Writeappend(l_Blob, Length(Hextoraw(Txt_In)) / 2, Hextoraw(Txt_In));
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Lob_Append;
/

Create Or Replace Procedure Zl_临床路径版本_Update
(
  路径id_In     临床路径版本.路径id%Type,
  版本号_In     临床路径版本.版本号%Type,
  标准住院日_In 临床路径版本.标准住院日%Type,
  标准费用_In   临床路径版本.标准费用%Type,
  版本说明_In   临床路径版本.版本说明%Type
) Is
Begin
  Update 临床路径版本
  Set 标准住院日 = 标准住院日_In, 标准费用 = 标准费用_In, 版本说明 = 版本说明_In
  Where 路径id = 路径id_In And 版本号 = 版本号_In;
  If Sql%RowCount = 0 Then
    Insert Into 临床路径版本
      (路径id, 版本号, 标准住院日, 标准费用, 版本说明, 创建人, 创建时间)
    Values
      (路径id_In, 版本号_In, 标准住院日_In, 标准费用_In, 版本说明_In, zl_UserName, Sysdate);
  Else
    --删除对应的导入评估信息,后面重新保存
    Delete From 临床路径评估 Where 路径id = 路径id_In And 版本号 = 版本号_In And 评估类型 = 1;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床路径版本_Update;
/

Create Or Replace Procedure Zl_临床路径分类_Insert
(
  路径id_In 临床路径分类.路径id%Type,
  版本号_In 临床路径分类.版本号%Type,
  序号_In   临床路径分类.序号%Type,
  名称_In   临床路径分类.名称%Type,
  Clear_In  Number := 0
  --参数：
  --  Clear_IN：插入前是否清除当前版本路径的所有分类
) Is
Begin
  If Nvl(Clear_In, 0) = 1 And 序号_In = 1 Then
    Delete From 临床路径分类 Where 路径id = 路径id_In And 版本号 = 版本号_In;
  End If;
  Insert Into 临床路径分类 (路径id, 版本号, 序号, 名称) Values (路径id_In, 版本号_In, 序号_In, 名称_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床路径分类_Insert;
/

Create Or Replace Procedure Zl_临床路径阶段_Delete(阶段id_In Varchar2) Is
  --参数：
  --  阶段ID_IN：ID1,ID2,...
Begin
  Delete /*+ Rule*/
  From 临床路径阶段
  Where ID In (Select * From Table(Cast(f_Num2list(阶段id_In) As Zltools.t_Numlist)));
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床路径阶段_Delete;
/

Create Or Replace Procedure Zl_临床路径阶段_Insert
(
  Id_In       临床路径阶段.Id%Type,
  路径id_In   临床路径阶段.路径id%Type,
  版本号_In   临床路径阶段.版本号%Type,
  父id_In     临床路径阶段.父id%Type,
  序号_In     临床路径阶段.序号%Type,
  名称_In     临床路径阶段.名称%Type,
  开始天数_In 临床路径阶段.开始天数%Type,
  结束天数_In 临床路径阶段.结束天数%Type,
  标志_In     临床路径阶段.标志%Type,
  说明_In     临床路径阶段.说明%Type
) Is
Begin
  Insert Into 临床路径阶段
    (ID, 路径id, 版本号, 父id, 序号, 名称, 开始天数, 结束天数, 标志, 说明)
  Values
    (Id_In, 路径id_In, 版本号_In, 父id_In, 序号_In, 名称_In, 开始天数_In, 结束天数_In, 标志_In, 说明_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床路径阶段_Insert;
/

Create Or Replace Procedure Zl_临床路径阶段_Update
(
  Id_In       临床路径阶段.Id%Type,
  路径id_In   临床路径阶段.路径id%Type,
  版本号_In   临床路径阶段.版本号%Type,
  序号_In     临床路径阶段.序号%Type,
  名称_In     临床路径阶段.名称%Type,
  开始天数_In 临床路径阶段.开始天数%Type,
  结束天数_In 临床路径阶段.结束天数%Type,
  标志_In     临床路径阶段.标志%Type,
  说明_In     临床路径阶段.说明%Type
) Is
Begin
  Update 临床路径阶段
  Set 序号 = 序号_In, 名称 = 名称_In, 开始天数 = 开始天数_In, 结束天数 = 结束天数_In, 标志 = 标志_In, 说明 = 说明_In
  Where ID = Id_In And 路径id = 路径id_In And 版本号 = 版本号_In;

  --删除对应的阶段评估信息,后面重新保存
  Delete From 临床路径评估 Where 路径id = 路径id_In And 版本号 = 版本号_In And 阶段id = Id_In And 评估类型 = 2;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床路径阶段_Update;
/

Create Or Replace Procedure Zl_路径医嘱内容_Insert
(
  Id_In         路径医嘱内容.Id%Type,
  相关id_In     路径医嘱内容.相关id%Type,
  序号_In       路径医嘱内容.序号%Type,
  期效_In       路径医嘱内容.期效%Type,
  诊疗项目id_In 路径医嘱内容.诊疗项目id%Type,
  医嘱内容_In   路径医嘱内容.医嘱内容%Type,
  单次用量_In   路径医嘱内容.单次用量%Type,
  总给予量_In   路径医嘱内容.总给予量%Type,
  收费细目id_In 路径医嘱内容.收费细目id%Type,
  标本部位_In   路径医嘱内容.标本部位%Type,
	检查方法_In   路径医嘱内容.检查方法%Type,
  执行频次_In   路径医嘱内容.执行频次%Type,
  频率次数_In   路径医嘱内容.频率次数%Type,
  频率间隔_In   路径医嘱内容.频率间隔%Type,
  间隔单位_In   路径医嘱内容.间隔单位%Type,
  医生嘱托_In   路径医嘱内容.医生嘱托%Type,
  执行性质_In   路径医嘱内容.执行性质%Type,
  执行科室id_In 路径医嘱内容.执行科室id%Type,
  时间方案_In   路径医嘱内容.时间方案%Type,
  路径id_In     临床路径项目.路径id%Type := Null,
  版本号_In     临床路径项目.版本号%Type := Null
) Is
  --参数：
  --  路径ID_IN,版本号_IN：当传入时，表示要清除指定版本的路径表中的所有医嘱内容和关联数据
Begin
  If 路径id_In Is Not Null And 版本号_In Is Not Null Then
    --会级联删除
    --Delete From 临床路径医嘱
    --Where 路径项目id In (Select ID From 临床路径项目 Where 路径id = 路径id_In And 版本号 = 版本号_In);
  
    Delete From 路径医嘱内容
    Where ID In (Select 医嘱内容id
                 From 临床路径项目 A, 临床路径医嘱 B
                 Where a.Id = b.路径项目id And a.路径id = 路径id_In And a.版本号 = 版本号_In);
  End If;

  Insert Into 路径医嘱内容
    (ID, 相关id, 序号, 期效, 诊疗项目id, 医嘱内容, 单次用量, 总给予量, 收费细目id, 标本部位, 检查方法, 执行频次, 频率次数, 频率间隔, 间隔单位, 医生嘱托, 执行性质, 执行科室id, 时间方案)
  Values
    (Id_In, 相关id_In, 序号_In, 期效_In, 诊疗项目id_In, 医嘱内容_In, 单次用量_In, 总给予量_In, 收费细目id_In, 标本部位_In, 检查方法_In, 执行频次_In, 频率次数_In,
     频率间隔_In, 间隔单位_In, 医生嘱托_In, 执行性质_In, 执行科室id_In, 时间方案_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_路径医嘱内容_Insert;
/

Create Or Replace Procedure Zl_临床路径项目_Insert
(
  Id_In       临床路径项目.Id%Type,
  路径id_In   临床路径项目.路径id%Type,
  版本号_In   临床路径项目.版本号%Type,
  阶段id_In   临床路径项目.阶段id%Type,
  分类_In     临床路径项目.分类%Type,
  项目序号_In 临床路径项目.项目序号%Type,
  项目内容_In 临床路径项目.项目内容%Type,
  执行方式_In 临床路径项目.执行方式%Type,
  执行者_In   临床路径项目.执行者%Type,
  项目结果_In 临床路径项目.项目结果%Type,
  图标id_In   临床路径项目.图标id%Type,
  医嘱id_In   Varchar2,
  病历id_In   Varchar2
  --参数：
  --   医嘱ID_IN：对应路径医嘱内容的ID，格式为ID1,ID2,....
  --   病历ID_IN：对应病历文件列表的ID，格式为ID1,ID2,...
) Is
Begin
  Insert Into 临床路径项目
    (ID, 路径id, 版本号, 阶段id, 分类, 项目序号, 项目内容, 执行方式, 执行者, 项目结果, 图标id)
  Values
    (Id_In, 路径id_In, 版本号_In, 阶段id_In, 分类_In, 项目序号_In, 项目内容_In, 执行方式_In, 执行者_In, 项目结果_In, 图标id_In);

  --处理医嘱关联
  If 医嘱id_In Is Not Null Then
    For r_Row In (Select /*+ Rule*/
                   Column_Value As ID
                  From Table(Cast(f_Num2list(医嘱id_In) As Zltools.t_Numlist))) Loop
      Insert Into 临床路径医嘱 (路径项目id, 医嘱内容id) Values (Id_In, r_Row.Id);
    End Loop;
  End If;

  --处理病历关联
  If 病历id_In Is Not Null Then
    For r_Row In (Select /*+ Rule*/
                   Column_Value As ID
                  From Table(Cast(f_Num2list(病历id_In) As Zltools.t_Numlist))) Loop
      Insert Into 临床路径病历 (项目id, 文件id) Values (Id_In, r_Row.Id);
    End Loop;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床路径项目_Insert;
/

Create Or Replace Procedure Zl_临床路径项目_Update
(
  Id_In       临床路径项目.Id%Type,
  路径id_In   临床路径项目.路径id%Type,
  版本号_In   临床路径项目.版本号%Type,
  项目序号_In 临床路径项目.项目序号%Type,
  项目内容_In 临床路径项目.项目内容%Type,
  执行方式_In 临床路径项目.执行方式%Type,
  执行者_In   临床路径项目.执行者%Type,
  项目结果_In 临床路径项目.项目结果%Type,
  图标id_In   临床路径项目.图标id%Type,
  医嘱id_In   Varchar2,
  病历id_In   Varchar2
) Is
Begin
  Update 临床路径项目
  Set 项目序号 = 项目序号_In, 项目内容 = 项目内容_In, 执行方式 = 执行方式_In, 执行者 = 执行者_In, 项目结果 = 项目结果_In, 图标id = 图标id_In
  Where ID = Id_In And 路径id = 路径id_In And 版本号 = 版本号_In;

  --处理医嘱关联
  Delete From 临床路径医嘱 Where 路径项目id = Id_In;
  If 医嘱id_In Is Not Null Then
    For r_Row In (Select /*+ Rule*/
                   Column_Value As ID
                  From Table(Cast(f_Num2list(医嘱id_In) As Zltools.t_Numlist))) Loop
      Insert Into 临床路径医嘱 (路径项目id, 医嘱内容id) Values (Id_In, r_Row.Id);
    End Loop;
  End If;

  --处理病历关联
  Delete From 临床路径病历 Where 项目id = Id_In;
  If 病历id_In Is Not Null Then
    For r_Row In (Select /*+ Rule*/
                   Column_Value As ID
                  From Table(Cast(f_Num2list(病历id_In) As Zltools.t_Numlist))) Loop
      Insert Into 临床路径病历 (项目id, 文件id) Values (Id_In, r_Row.Id);
    End Loop;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床路径项目_Update;
/

Create Or Replace Procedure Zl_临床路径项目_Delete(项目id_In Varchar2) Is
  --参数：
  --  项目id_In：ID1,ID2,...
Begin
  Delete /*+ Rule*/
  From 临床路径项目
  Where ID In (Select * From Table(Cast(f_Num2list(项目id_In) As Zltools.t_Numlist)));
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床路径项目_Delete;
/

Create Or Replace Procedure Zl_临床路径版本_Audit
(
  路径id_In 临床路径项目.路径id%Type,
  版本号_In 临床路径项目.版本号%Type,
  审核_In   Number
  --参数：
  --   审核_IN：1=通过审核，-1=取消审核
) Is
  v_Date  Date;
  v_Count Number;
  v_User  人员表.姓名%Type;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  If 审核_In = 1 Then
    --审核
    Select Sysdate Into v_Date From Dual;
    Select zl_UserName Into v_User From Dual;
  
    Update 临床路径版本
    Set 审核人 = v_User, 审核时间 = v_Date
    Where 路径id = 路径id_In And 版本号 = 版本号_In And 审核时间 Is Null;
    If Sql%RowCount > 0 Then
      --自动停用之前的版本
      Update 临床路径版本
      Set 停用人 = v_User, 停用时间 = v_Date
      Where 路径id = 路径id_In And 版本号 < 版本号_In And 停用时间 Is Null;
    
      Update 临床路径目录 Set 最新版本 = 版本号_In Where ID = 路径id_In;
    End If;
  Elsif 审核_In = -1 Then
    --取消审核
    Select Count(*) Into v_Count From 病人临床路径 Where 路径id = 路径id_In And 版本号 = 版本号_In And Rownum = 1;
    If Nvl(v_Count, 0) > 0 Then
      v_Error := '该版本的临床路径已经在使用，不能取消审核。';
      Raise Err_Custom;
    End If;
  
    Select Count(*) Into v_Count From 临床路径版本 Where 路径id = 路径id_In And 版本号 > 版本号_In;
    If Nvl(v_Count, 0) > 0 Then
      v_Error := '该版本后面存在其他新的版本，不能取消审核。';
      Raise Err_Custom;
    End If;
  
    Select 审核人, 审核时间 Into v_User, v_Date From 临床路径版本 Where 路径id = 路径id_In And 版本号 = 版本号_In;
  
    Update 临床路径版本
    Set 审核人 = Null, 审核时间 = Null
    Where 路径id = 路径id_In And 版本号 = 版本号_In And 审核时间 Is Not Null;
    If Sql%RowCount > 0 Then
      --恢复之前审核时自动停用的版本(手工停用的不处理)
      Select Max(版本号)
      Into v_Count
      From 临床路径版本
      Where 路径id = 路径id_In And 版本号 < 版本号_In And 停用人 = v_User And 停用时间 = v_Date;
      If Nvl(v_Count, 0) > 0 Then
        Update 临床路径版本 Set 停用人 = Null, 停用时间 = Null Where 路径id = 路径id_In And 版本号 = v_Count;
      End If;
    
      --更新最新版本
      Select Max(版本号)
      Into v_Count
      From 临床路径版本
      Where 路径id = 路径id_In And 审核时间 Is Not Null And 停用时间 Is Null;
      If Nvl(v_Count, 0) > 0 Then
        Update 临床路径目录 Set 最新版本 = v_Count Where ID = 路径id_In;
      Else
        Update 临床路径目录 Set 最新版本 = Null Where ID = 路径id_In;
      End If;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床路径版本_Audit;
/

Create Or Replace Procedure Zl_临床路径版本_Stop
(
  路径id_In 临床路径项目.路径id%Type,
  版本号_In 临床路径项目.版本号%Type,
  停用_In   Number
  --参数：
  --   停用_In：1=停用，-1=取消停用
) Is
  v_Date  Date;
  v_Count Number;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  If 停用_In = 1 Then
    Select 审核时间 Into v_Date From 临床路径版本 Where 路径id = 路径id_In And 版本号 = 版本号_In;
    If v_Date Is Null Then
      v_Error := '该版本的临床路径尚未审核，不需要停用。';
      Raise Err_Custom;
    End If;
  
    Update 临床路径版本
    Set 停用人 = zl_UserName, 停用时间 = Sysdate
    Where 路径id = 路径id_In And 版本号 = 版本号_In And 停用时间 Is Null;
  Elsif 停用_In = -1 Then
    Select Count(*)
    Into v_Count
    From 临床路径版本
    Where 路径id = 路径id_In And 版本号 > 版本号_In And (停用时间 Is Not Null Or 审核时间 Is Not Null);
    If Nvl(v_Count, 0) > 0 Then
      v_Error := '该版本后面存在其他已经审核或者停用的版本，不能取消停用。';
      Raise Err_Custom;
    End If;
  
    Update 临床路径版本
    Set 停用人 = Null, 停用时间 = Null
    Where 路径id = 路径id_In And 版本号 = 版本号_In And 停用时间 Is Not Null;
  End If;

  --更新最新版本
  Select Max(版本号)
  Into v_Count
  From 临床路径版本
  Where 路径id = 路径id_In And 审核时间 Is Not Null And 停用时间 Is Null;
  If Nvl(v_Count, 0) > 0 Then
    Update 临床路径目录 Set 最新版本 = v_Count Where ID = 路径id_In;
  Else
    Update 临床路径目录 Set 最新版本 = Null Where ID = 路径id_In;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床路径版本_Stop;
/

Create Or Replace Procedure Zl_临床路径版本_Delete
(
  路径id_In 临床路径版本.路径id%Type,
  版本号_In 临床路径版本.版本号%Type
) Is
  v_Count Number;
Begin
  Delete From 路径医嘱内容
  Where ID In (Select 医嘱内容id
               From 临床路径项目 A, 临床路径医嘱 B
               Where a.Id = b.路径项目id And a.路径id = 路径id_In And a.版本号 = 版本号_In);

  --关联表自动级联删除
  Delete From 临床路径版本 Where 路径id = 路径id_In And 版本号 = 版本号_In;

  --更新最新版本
  Select Max(版本号)
  Into v_Count
  From 临床路径版本
  Where 路径id = 路径id_In And 审核时间 Is Not Null And 停用时间 Is Null;
  If Nvl(v_Count, 0) > 0 Then
    Update 临床路径目录 Set 最新版本 = v_Count Where ID = 路径id_In;
  Else
    Update 临床路径目录 Set 最新版本 = Null Where ID = 路径id_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床路径版本_Delete;
/

Create Or Replace Procedure Zl_路径评估指标_Insert
(
  路径id_In   临床路径评估.路径id%Type,
  版本号_In   临床路径评估.版本号%Type,
  阶段id_In   临床路径评估.阶段id%Type,
  评估类型_In 临床路径评估.评估类型%Type,
  指标id_In   路径评估指标.Id%Type,
  序号_In     路径评估指标.序号%Type,
  评估指标_In 路径评估指标.评估指标%Type,
  指标类型_In 路径评估指标.指标类型%Type,
  指标结果_In 路径评估指标.指标结果%Type
) Is
  v_评估id 临床路径评估.Id%Type;
Begin
  Begin
    Select ID
    Into v_评估id
    From 临床路径评估
    Where 路径id = 路径id_In And 版本号 = 版本号_In And Nvl(阶段id, 0) = Nvl(阶段id_In, 0) And 评估类型 = 评估类型_In;
  Exception
    When Others Then
      Null;
  End;

  If v_评估id Is Null Then
    Select 临床路径评估_Id.Nextval Into v_评估id From Dual;
    Insert Into 临床路径评估
      (ID, 路径id, 版本号, 阶段id, 评估类型)
    Values
      (v_评估id, 路径id_In, 版本号_In, 阶段id_In, 评估类型_In);
  End If;

  Insert Into 路径评估指标
    (ID, 评估id, 序号, 评估指标, 指标类型, 指标结果)
  Values
    (指标id_In, v_评估id, 序号_In, 评估指标_In, 指标类型_In, 指标结果_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_路径评估指标_Insert;
/

Create Or Replace Procedure Zl_路径评估条件_Insert
(
  路径id_In   临床路径评估.路径id%Type,
  版本号_In   临床路径评估.版本号%Type,
  阶段id_In   临床路径评估.阶段id%Type,
  评估类型_In 临床路径评估.评估类型%Type,
  指标id_In   路径评估条件.指标id%Type,
  项目id_In   路径评估条件.项目id%Type,
  关系式_In   路径评估条件.关系式%Type,
  条件值_In   路径评估条件.条件值%Type,
  条件组合_In 路径评估条件.条件组合%Type
) Is
  v_评估id 临床路径评估.Id%Type;
Begin
  Begin
    Select ID
    Into v_评估id
    From 临床路径评估
    Where 路径id = 路径id_In And 版本号 = 版本号_In And Nvl(阶段id, 0) = Nvl(阶段id_In, 0) And 评估类型 = 评估类型_In;
  Exception
    When Others Then
      Null;
  End;

  If v_评估id Is Null Then
    Select 临床路径评估_Id.Nextval Into v_评估id From Dual;
    Insert Into 临床路径评估
      (ID, 路径id, 版本号, 阶段id, 评估类型)
    Values
      (v_评估id, 路径id_In, 版本号_In, 阶段id_In, 评估类型_In);
  End If;

  Insert Into 路径评估条件
    (评估id, 指标id, 项目id, 关系式, 条件值, 条件组合)
  Values
    (v_评估id, 指标id_In, 项目id_In, 关系式_In, 条件值_In, 条件组合_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_路径评估条件_Insert;
/

Create Or Replace Procedure Zl_临床路径版本_Copy
(
  源路径id_In   临床路径版本.路径id%Type,
  源版本号_In   临床路径版本.版本号%Type,
  目标路径id_In 临床路径版本.路径id%Type,
  目标版本号_In 临床路径版本.版本号%Type
  --功能：复制产生新的临床路径版本
  --参数：
  --  源版本号_In：如果未指定(0或NULL)，则取最新有效的版本号
  --  目标本号_In：如果未指定(0或NULL)，则产生新的版本号
) Is
  v_源版本号   临床路径版本.版本号%Type;
  v_目标版本号 临床路径版本.版本号%Type;

  v_Advice_Id Number;
  v_Step_Id   Number;
  v_Item_Id   Number;
  v_Eval_Id   Number;
  v_Mark_Id   Number;

  v_Error Varchar2(255);
  Err_Custom Exception;

  --调整序列相关函数
  Type t_Id_Table Is Table Of Number;
  Arr_Id t_Id_Table;

  Procedure Adjuest_Sequence_Advice(n_Count Number) Is
  Begin
    Select 路径医嘱内容_Id.Nextval Bulk Collect Into Arr_Id From Dual Connect By Rownum <= n_Count;
  End;
  Procedure Adjuest_Sequence_Step(n_Count Number) Is
  Begin
    Select 临床路径阶段_Id.Nextval Bulk Collect Into Arr_Id From Dual Connect By Rownum <= n_Count;
  End;
  Procedure Adjuest_Sequence_Item(n_Count Number) Is
  Begin
    Select 临床路径项目_Id.Nextval Bulk Collect Into Arr_Id From Dual Connect By Rownum <= n_Count;
  End;
  Procedure Adjuest_Sequence_Eval(n_Count Number) Is
  Begin
    Select 临床路径评估_Id.Nextval Bulk Collect Into Arr_Id From Dual Connect By Rownum <= n_Count;
  End;
  Procedure Adjuest_Sequence_Mark(n_Count Number) Is
  Begin
    Select 路径评估指标_Id.Nextval Bulk Collect Into Arr_Id From Dual Connect By Rownum <= n_Count;
  End;
Begin
  --确定源路径版本号
  v_源版本号 := Nvl(源版本号_In, 0);
  If v_源版本号 = 0 Then
    Select 最新版本 Into v_源版本号 From 临床路径目录 Where ID = 源路径id_In;
    If Nvl(v_源版本号, 0) = 0 Then
      v_Error := '要复制的来源临床路径中没有可用的有效版本。';
      Raise Err_Custom;
    End If;
  End If;

  --确定目标路径版本号
  v_目标版本号 := Nvl(目标版本号_In, 0);
  If v_目标版本号 = 0 Then
    Select Nvl(Max(版本号), 0) + 1 Into v_目标版本号 From 临床路径版本 Where 路径id = 目标路径id_In;
  Else
    Zl_临床路径版本_Delete(目标路径id_In, 目标版本号_In);
  End If;

  --路径医嘱内容
  Select 路径医嘱内容_Id.Currval, 路径医嘱内容_Id.Nextval Into v_Advice_Id, v_Advice_Id From Dual;

  Select v_Advice_Id - Nvl(Min(ID), 0) + 1
  Into v_Advice_Id
  From 路径医嘱内容
  Where ID In (Select b.医嘱内容id
               From 临床路径项目 A, 临床路径医嘱 B
               Where a.Id = b.路径项目id And a.路径id = 源路径id_In And a.版本号 = v_源版本号);

  Insert Into 路径医嘱内容
    (ID, 相关id, 序号, 期效, 诊疗项目id, 医嘱内容, 单次用量, 总给予量, 收费细目id, 标本部位, 检查方法, 执行频次, 频率次数, 频率间隔, 间隔单位, 医生嘱托, 执行性质, 执行科室id, 时间方案)
    Select ID + v_Advice_Id, 相关id + v_Advice_Id, 序号, 期效, 诊疗项目id, 医嘱内容, 单次用量, 总给予量, 收费细目id, 标本部位, 检查方法, 执行频次, 频率次数, 频率间隔,
           间隔单位, 医生嘱托, 执行性质, 执行科室id, 时间方案
    From 路径医嘱内容
    Where ID In (Select b.医嘱内容id
                 From 临床路径项目 A, 临床路径医嘱 B
                 Where a.Id = b.路径项目id And a.路径id = 源路径id_In And a.版本号 = v_源版本号);
  Adjuest_Sequence_Advice(v_Advice_Id); --调整序列

  --临床路径版本
  Insert Into 临床路径版本
    (路径id, 版本号, 标准住院日, 标准费用, 版本说明, 创建人, 创建时间)
    Select 目标路径id_In, v_目标版本号, 标准住院日, 标准费用, 版本说明, 创建人, 创建时间
    From 临床路径版本
    Where 路径id = 源路径id_In And 版本号 = v_源版本号;

  --临床路径分类
  Insert Into 临床路径分类
    (路径id, 版本号, 序号, 名称)
    Select 目标路径id_In, v_目标版本号, 序号, 名称
    From 临床路径分类
    Where 路径id = 源路径id_In And 版本号 = v_源版本号;

  --临床路径阶段
  Select 临床路径阶段_Id.Currval, 临床路径阶段_Id.Nextval Into v_Step_Id, v_Step_Id From Dual;
  Select v_Step_Id - Nvl(Min(ID), 0) + 1
  Into v_Step_Id
  From 临床路径阶段
  Where 路径id = 源路径id_In And 版本号 = v_源版本号;

  Insert Into 临床路径阶段
    (ID, 路径id, 版本号, 父id, 序号, 名称, 开始天数, 结束天数, 标志, 说明)
    Select ID + v_Step_Id, 目标路径id_In, v_目标版本号, 父id + v_Step_Id, 序号, 名称, 开始天数, 结束天数, 标志, 说明
    From 临床路径阶段
    Where 路径id = 源路径id_In And 版本号 = v_源版本号;
  Adjuest_Sequence_Step(v_Step_Id); --调整序列

  --临床路径项目
  Select 临床路径项目_Id.Currval, 临床路径项目_Id.Nextval Into v_Item_Id, v_Item_Id From Dual;
  Select v_Item_Id - Nvl(Min(ID), 0) + 1
  Into v_Item_Id
  From 临床路径项目
  Where 路径id = 源路径id_In And 版本号 = v_源版本号;

  Insert Into 临床路径项目
    (ID, 路径id, 版本号, 阶段id, 分类, 项目序号, 项目内容, 执行方式, 执行者, 项目结果, 图标id)
    Select ID + v_Item_Id, 目标路径id_In, v_目标版本号, 阶段id + v_Step_Id, 分类, 项目序号, 项目内容, 执行方式, 执行者, 项目结果, 图标id
    From 临床路径项目
    Where 路径id = 源路径id_In And 版本号 = v_源版本号;
  Adjuest_Sequence_Item(v_Item_Id); --调整序列

  --临床路径医嘱
  Insert Into 临床路径医嘱
    (路径项目id, 医嘱内容id)
    Select b.路径项目id + v_Item_Id, b.医嘱内容id + v_Advice_Id
    From 临床路径项目 A, 临床路径医嘱 B
    Where a.Id = b.路径项目id And a.路径id = 源路径id_In And a.版本号 = v_源版本号;

  --临床路径病历
  Insert Into 临床路径病历
    (项目id, 文件id)
    Select b.项目id + v_Item_Id, b.文件id
    From 临床路径项目 A, 临床路径病历 B
    Where a.Id = b.项目id And a.路径id = 源路径id_In And a.版本号 = v_源版本号;

  --临床路径评估
  Select 临床路径评估_Id.Currval, 临床路径评估_Id.Nextval Into v_Eval_Id, v_Eval_Id From Dual;
  Select v_Eval_Id - Nvl(Min(ID), 0) + 1
  Into v_Eval_Id
  From 临床路径评估
  Where 路径id = 源路径id_In And 版本号 = v_源版本号;

  Insert Into 临床路径评估
    (ID, 路径id, 版本号, 阶段id, 评估类型)
    Select ID + v_Eval_Id, 目标路径id_In, v_目标版本号, 阶段id + v_Step_Id, 评估类型
    From 临床路径评估
    Where 路径id = 源路径id_In And 版本号 = v_源版本号;
  Adjuest_Sequence_Eval(v_Eval_Id); --调整序列

  --路径评估指标
  Select 路径评估指标_Id.Currval, 路径评估指标_Id.Nextval Into v_Mark_Id, v_Mark_Id From Dual;
  Select v_Mark_Id - Nvl(Min(ID), 0) + 1
  Into v_Mark_Id
  From 路径评估指标
  Where 评估id In (Select ID From 临床路径评估 Where 路径id = 源路径id_In And 版本号 = v_源版本号);

  Insert Into 路径评估指标
    (ID, 评估id, 序号, 评估指标, 指标类型, 指标结果)
    Select ID + v_Mark_Id, 评估id + v_Eval_Id, 序号, 评估指标, 指标类型, 指标结果
    From 路径评估指标
    Where 评估id In (Select ID From 临床路径评估 Where 路径id = 源路径id_In And 版本号 = v_源版本号);
  Adjuest_Sequence_Mark(v_Mark_Id); --调整序列

  --路径评估条件
  Insert Into 路径评估条件
    (评估id, 指标id, 项目id, 关系式, 条件值, 条件组合)
    Select 评估id + v_Eval_Id, 指标id + v_Mark_Id, 项目id + v_Item_Id, 关系式, 条件值, 条件组合
    From 路径评估条件
    Where 评估id In (Select ID From 临床路径评估 Where 路径id = 源路径id_In And 版本号 = v_源版本号);
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床路径版本_Copy;
/












--临床路径应用相关过程
---------------------------------------------------------------------------------------------------------------------------

Create Or Replace Procedure Zl_病人路径导入_Insert
(
  病人id_In   病人临床路径.病人id%Type,
  主页id_In   病人临床路径.主页id%Type,
  科室id_In   病人临床路径.科室id%Type,
  路径id_In   病人临床路径.路径id%Type,
  版本号_In   病人临床路径.版本号%Type,
  路径记录_In 病人临床路径.Id%Type,
  导入人_In   病人临床路径.导入人%Type,
  导入说明_In 病人临床路径.导入说明%Type,
  符合导入_In 病人临床路径.状态%Type, --0=不符合,1=符合
  指标评估_In Varchar2, --指标名称|指标结果|指标类型||...,末尾带||,允许为空
  序号_In     Number
) Is
  v_Str   Varchar2(4000);
  v_Tmp   Varchar2(1000);
  v_Index Number;
  I       Number(5) := 1;

  l_指标名称 t_Strlist := t_Strlist();
  l_指标结果 t_Strlist := t_Strlist();
  l_指标类型 t_Numlist := t_Numlist();

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  If 序号_In = 1 Then
    Insert Into 病人临床路径
      (ID, 病人id, 主页id, 科室id, 路径id, 版本号, 导入人, 导入时间, 导入说明, 状态)
    Values
      (路径记录_In, 病人id_In, 主页id_In, 科室id_In, 路径id_In, 版本号_In, 导入人_In, Sysdate, 导入说明_In, 符合导入_In);
  End If;

  If Not 指标评估_In Is Null Then
    v_Str := 指标评估_In;
    Loop
      v_Index := Instr(v_Str, '||');
      Exit When(Nvl(v_Index, 0) = 0);
      l_指标名称.Extend;
      l_指标结果.Extend;
      l_指标类型.Extend;
    
      v_Tmp := Substr(v_Str, 1, v_Index - 1);
      l_指标名称(I) := Substr(v_Tmp, 1, Instr(v_Tmp, '|') - 1);
      v_Tmp := Substr(v_Tmp, Instr(v_Tmp, '|') + 1);
      l_指标结果(I) := Trim(Substr(v_Tmp, 1, Instr(v_Tmp, '|') - 1));
      v_Tmp := Substr(v_Tmp, Instr(v_Tmp, '|') + 1);
      l_指标类型(I) := To_Number(v_Tmp);
    
      v_Str := Substr(v_Str, v_Index + 2);
      I     := I + 1;
    End Loop;
  
    Forall I In 1 .. l_指标名称.Count
      Insert Into 病人路径指标
        (路径记录id, 评估类型, 评估指标, 指标结果, 指标类型)
      Values
        (路径记录_In, 1, l_指标名称(I), l_指标结果(I), l_指标类型(I));
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病人路径导入_Insert;
/


Create Or Replace Procedure Zl_病人路径导入_Delete(病人路径id_In 病人临床路径.Id%Type) Is
  v_Count 病人临床路径.Id%Type;
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  Select Nvl(Max(路径记录id), 0) Into v_Count From 病人路径执行 Where 路径记录id = 病人路径id_In;

  If v_Count = 0 Then
    Delete 病人路径指标 Where 路径记录id = 病人路径id_In;
    Delete 病人临床路径 Where ID = 病人路径id_In;
  Else
    v_Error := '该病人的路径已生成了路径项目,不能取消导入。';
    Raise Err_Custom;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病人路径导入_Delete;
/

Create Or Replace Procedure Zl_病人路径评估_Insert
(
  功能_In       Number, --1=新增,2=修改
  路径记录id_In 病人临床路径.Id%Type,
  阶段id_In     临床路径阶段.Id%Type,
  日期_In       病人路径评估.日期%Type,
  天数_In       病人路径评估.天数%Type,
  评估人_In     病人路径评估.评估人%Type,
  评估结果_In   病人路径评估.评估结果%Type,
  评估说明_In   病人路径评估.评估说明%Type,
  登记人_In     病人路径评估.登记人%Type,
  指标评估_In   Varchar2, --指标名称|指标结果|指标类型||...,末尾带||,允许为空
  序号_In       Number
) Is
  v_Str   Varchar2(4000);
  v_Tmp   Varchar2(1000);
  v_Index Number;
  I       Number(5) := 1;

  l_指标名称 t_Strlist := t_Strlist();
  l_指标结果 t_Strlist := t_Strlist();
  l_指标类型 t_Numlist := t_Numlist();

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin

  If 序号_In = 1 Then
    If 功能_In = 1 Then
      Insert Into 病人路径评估
        (路径记录id, 阶段id, 日期, 天数, 评估人, 评估时间, 评估结果, 评估说明, 登记人, 登记时间)
      Values
        (路径记录id_In, 阶段id_In, 日期_In, 天数_In, 评估人_In, Sysdate, 评估结果_In, 评估说明_In, 登记人_In, Sysdate);
    Else
      Update 病人路径评估
      Set 评估人 = 评估人_In, 评估时间 = Sysdate, 评估结果 = 评估结果_In, 评估说明 = 评估说明_In, 登记人 = 登记人_In, 登记时间 = Sysdate
      Where 路径记录id = 路径记录id_In And 阶段id = 阶段id_In And 日期 = 日期_In;
    End If;
  End If;

  If Not 指标评估_In Is Null Then
    v_Str := 指标评估_In;
    Loop
      v_Index := Instr(v_Str, '||');
      Exit When(Nvl(v_Index, 0) = 0);
      l_指标名称.Extend;
      l_指标结果.Extend;
      l_指标类型.Extend;
    
      v_Tmp := Substr(v_Str, 1, v_Index - 1);
      l_指标名称(I) := Substr(v_Tmp, 1, Instr(v_Tmp, '|') - 1);
      v_Tmp := Substr(v_Tmp, Instr(v_Tmp, '|') + 1);
      l_指标结果(I) := Trim(Substr(v_Tmp, 1, Instr(v_Tmp, '|') - 1));
      v_Tmp := Substr(v_Tmp, Instr(v_Tmp, '|') + 1);
      l_指标类型(I) := To_Number(v_Tmp);
    
      v_Str := Substr(v_Str, v_Index + 2);
      I     := I + 1;
    End Loop;
  
    If 功能_In = 1 Then
      Forall I In 1 .. l_指标名称.Count
      
        Insert Into 病人路径指标
          (路径记录id, 阶段id, 日期, 天数, 评估类型, 评估指标, 指标结果, 指标类型)
        Values
          (路径记录id_In, 阶段id_In, 日期_In, 天数_In, 2, l_指标名称(I), l_指标结果(I), l_指标类型(I));
    Else
      Forall I In 1 .. l_指标名称.Count
        Update 病人路径指标
        Set 指标结果 = l_指标结果(I)
        Where 路径记录id = 路径记录id_In And 阶段id = 阶段id_In And 日期 = 日期_In And 评估指标 = l_指标名称(I);
    End If;
  End If;

  If 评估结果_In = -1 Then
    If 功能_In = 2 Then
      v_Index := 0;
      Select 当前阶段id Into v_Index From 病人临床路径 Where ID = 路径记录id_In;
      If v_Index <> 阶段id_In Then
        v_Error := '该病人已生成了次日的路径项目,不能修改评估结果来结束路径。';
        Raise Err_Custom;
      End If;
    End If;
    Update 病人临床路径
    Set 结束时间 = Sysdate, 状态 = 3, 前一阶段id = 阶段id_In, 当前阶段id = Null, 当前天数 = Null
    Where ID = 路径记录id_In;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病人路径评估_Insert;
/

Create Or Replace Procedure Zl_病人路径评估_Delete
(
  路径记录id_In 病人路径执行.Id%Type,
  阶段id_In     病人路径执行.阶段id%Type,
  日期_In       病人路径执行.日期%Type
) Is
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin

  --评估结果为变异时自动结束的,取消结束自动取消评估
  Delete 病人路径评估 Where 路径记录id = 路径记录id_In And 阶段id = 阶段id_In And 日期 = 日期_In;
  Delete 病人路径指标 Where 路径记录id = 路径记录id_In And 阶段id = 阶段id_In And 日期 = 日期_In;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病人路径评估_Delete;
/


Create Or Replace Procedure Zl_病人路径生成_Insert
(
  序号_In        Number,
  病人id_In      病人临床路径.病人id%Type,
  主页id_In      病人临床路径.主页id%Type,
  婴儿_In        电子病历记录.婴儿%Type,
  科室id_In      病人临床路径.科室id%Type,
  路径记录id_In  病人路径执行.路径记录id%Type,
  阶段id_In      病人路径执行.阶段id%Type,
  日期_In        病人路径执行.日期%Type,
  天数_In        病人路径执行.天数%Type,
  分类_In        病人路径执行.分类%Type,
  项目id_In      病人路径执行.项目id%Type,
  医嘱ids_In     Varchar2,
  病历文件ids_In Varchar2,
  病人病历ids_In Varchar2,
  登记人_In      病人路径执行.登记人%Type,
  登记时间_In    病人路径执行.登记时间%Type,
  项目内容_In    病人路径执行.项目内容%Type := Null,
  执行者_In      病人路径执行.执行者%Type := Null,
  项目结果_In    病人路径执行.项目结果%Type := Null,
  图标id_In      病人路径执行.图标id%Type := Null,
  添加原因_In    病人路径执行.添加原因%Type := Null
) Is
  v_当前阶段id 病人临床路径.当前阶段id%Type;
  v_路径执行id 病人路径执行.Id%Type;
  v_病历id     电子病历记录.Id%Type;
  t_Advice     t_Numlist;
  t_File       t_Numlist;
  t_Doc        t_Numlist;

  v_Id         电子病历内容.Id%Type;
  v_父id       电子病历内容.父id%Type;
  v_当前父id   电子病历内容.父id%Type;
  v_对象序号   电子病历内容.对象序号%Type;
  v_原对象序号 电子病历内容.父id%Type;
  v_内容文本   电子病历内容.内容文本%Type;
  n_预制提纲id 电子病历内容.预制提纲id%Type;

  v_项目序号 病人路径执行.项目序号%Type;
  v_Error    Varchar2(255);
  Err_Custom Exception;
Begin
  If 序号_In = 1 And 项目内容_In Is Null Then
    Select Nvl(当前阶段id, 0) Into v_当前阶段id From 病人临床路径 Where ID = 路径记录id_In;
    If v_当前阶段id <> 阶段id_In Then
      Update 病人临床路径 Set 前一阶段id = 当前阶段id, 当前阶段id = 阶段id_In Where ID = 路径记录id_In;
    End If;
    Update 病人临床路径 Set 当前天数 = 天数_In Where ID = 路径记录id_In;
  End If;

  --添加的路径外项目
  If 项目内容_In Is Not Null Then
    Select Max(项目序号) + 1
    Into v_项目序号
    From 病人路径执行
    Where 路径记录id = 路径记录id_In And 阶段id = 阶段id_In And 天数 = 天数_In And 分类 = 分类_In;
    If v_项目序号 Is Null Then
      --排在所有路径项目之后,即使有可选的项目可能还未生成(可以补充生成,所以预留序号)
      Select Nvl(Max(项目序号), 0) + 1
      Into v_项目序号
      From 临床路径项目 A, 病人临床路径 B
      Where a.路径id = b.路径id And a.版本号 = b.版本号 And b.Id = 路径记录id_In And a.阶段id = 阶段id_In And 分类 = 分类_In;
    End If;
  End If;

  Select 病人路径执行_Id.Nextval Into v_路径执行id From Dual;
  Insert Into 病人路径执行
    (ID, 路径记录id, 阶段id, 日期, 天数, 分类, 项目id, 登记人, 登记时间, 项目序号, 项目内容, 执行者, 项目结果, 图标id, 添加原因)
  Values
    (v_路径执行id, 路径记录id_In, 阶段id_In, 日期_In, 天数_In, 分类_In, 项目id_In, 登记人_In, 登记时间_In, v_项目序号, 项目内容_In, 执行者_In, 项目结果_In,
     图标id_In, 添加原因_In);

  If 医嘱ids_In Is Not Null Then
    Select Column_Value Bulk Collect Into t_Advice From Table(f_Num2list(医嘱ids_In));
    Forall I In 1 .. t_Advice.Count
      Insert Into 病人路径医嘱 (路径执行id, 病人医嘱id) Values (v_路径执行id, t_Advice(I));
  End If;

  If 病人病历ids_In Is Not Null Then
    Select Column_Value Bulk Collect Into t_Doc From Table(f_Num2list(病人病历ids_In));
    Select Column_Value Bulk Collect Into t_File From Table(f_Num2list(病历文件ids_In));
    For I In 1 .. t_Doc.Count Loop
      v_病历id := t_Doc(I);
    
      Insert Into 电子病历记录
        (ID, 病人来源, 病人id, 主页id, 婴儿, 科室id, 病历种类, 文件id, 病历名称, 创建人, 创建时间, 保存人, 保存时间, 最后版本, 签名级别, 编辑方式, 路径执行id)
        Select v_病历id, 2, 病人id_In, 主页id_In, 婴儿_In, 科室id_In, 种类, ID, 名称, 登记人_In, 登记时间_In, 登记人_In, 登记时间_In, 1, 0, 0,
               v_路径执行id
        From 病历文件列表
        Where ID = t_File(I);
    
      v_对象序号 := 0;
      For Rs In (Select ID, 文件id, Nvl(父id, 0) As 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 预制提纲id, 复用提纲, 使用时机,
                        诊治要素id, 替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域
                 From 病历文件结构
                 Where 文件id = t_File(I)
                 Order By 对象序号) Loop
      
        v_对象序号 := v_对象序号 + 1;
        Select 电子病历内容_Id.Nextval Into v_Id From Dual;
      
        If Rs.父id = 0 Then
          v_当前父id := v_Id;
          v_父id     := Null;
          If Rs.对象类型 = 1 And Not Rs.预制提纲id Is Null Then
            n_预制提纲id := Rs.预制提纲id;
          Else
            n_预制提纲id := Null;
          End If;
        Else
          --对象序号为空的时候，父ID就不是按照顺序的了，需要重新查找
          If Rs.对象序号 Is Null Then
            n_预制提纲id := Null;
            Select 对象序号 Into v_原对象序号 From 病历文件结构 Where ID = Rs.父id;
            If v_原对象序号 Is Null Then
              v_父id := Null;
            Else
              Select ID Into v_父id From 电子病历内容 Where 文件id = v_病历id And 对象序号 = v_原对象序号;
            End If;
          Else
            v_父id := v_当前父id;
          End If;
        End If;
      
        If Rs.对象类型 = 4 And Rs.替换域 = 1 Then
          v_内容文本 := Zl_Replace_Element_Value(Rs.要素名称, 病人id_In, 主页id_In, 2, Null, 婴儿_In);
        Else
          v_内容文本 := Rs.内容文本;
        End If;
      
        Insert Into 电子病历内容
          (ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 预制提纲id, 复用提纲, 使用时机, 诊治要素id, 替换域,
           要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域)
        Values
          (v_Id, v_病历id, 1, 0, v_父id, v_对象序号, Rs.对象类型, Rs.对象标记, Rs.保留对象, Null, Rs.内容行次, v_内容文本, Rs.是否换行, Rs.预制提纲id,
           Rs.复用提纲, Rs.使用时机, Rs.诊治要素id, Rs.替换域, Rs.要素名称, Rs.要素类型, Rs.要素长度, Rs.要素小数, Rs.要素单位, Rs.要素表示, Rs.输入形态, Rs.要素值域);
      
        If Rs.对象类型 = 5 Then
          Insert Into 电子病历图形 (对象id, 图形) Values (v_Id, (Select 图形 From 病历文件图形 Where 对象id = Rs.Id));
        End If;
      
      End Loop;
    
      Insert Into 电子病历格式
        (文件id, 内容)
      Values
        (v_病历id, (Select 内容 From 病历文件格式 Where 文件id = t_File(I)));
    End Loop;
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病人路径生成_Insert;
/



Create Or Replace Procedure Zl_病人路径生成_Delete(执行记录id_In 病人路径执行.Id%Type) Is
  t_Id t_Numlist;

  --长期医嘱,其它阶段存在时不删除
  Cursor c_Advice Is
    Select 病人医嘱id
    From 病人路径医嘱 A
    Where 路径执行id = 执行记录id_In And Not Exists
     (Select 1 From 病人路径医嘱 B Where a.病人医嘱id = b.病人医嘱id And a.路径执行id <> b.路径执行id);

  Cursor c_Doc Is
    Select ID From 电子病历记录 Where 路径执行id = 执行记录id_In;

  v_阶段id     病人路径执行.阶段id%Type;
  v_路径记录id 病人路径执行.路径记录id%Type;
  v_登记时间   病人路径执行.登记时间%Type;
  v_天数       病人路径执行.天数%Type;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --是否允许取消的逻辑规则在界面程序中检查
  Delete 病人路径医嘱 Where 路径执行id = 执行记录id_In;
  Open c_Advice;
  Fetch c_Advice Bulk Collect
    Into t_Id;
  Close c_Advice;
  If t_Id.Count > 0 Then
    For I In 1 .. t_Id.Count Loop
      Zl_病人医嘱记录_Delete(t_Id(I), 0);
    End Loop;
  End If;

  Open c_Doc;
  Fetch c_Doc Bulk Collect
    Into t_Id;
  Close c_Doc;
  If t_Id.Count > 0 Then
    For I In 1 .. t_Id.Count Loop
      Zl_电子病历记录_Delete(t_Id(I));
    End Loop;
  End If;
  Delete 病人路径执行 Where ID = 执行记录id_In Returning 路径记录id, 阶段id Into v_路径记录id, v_阶段id;

  Select Max(天数) Into v_天数 From 病人路径执行 Where 路径记录id = v_路径记录id And 阶段id = v_阶段id;
  --如果当前阶段的最后一个执行记录被删除(全部都是非必须执行的情况下)
  If v_天数 Is Null Then
    --a.如果当前没有任何执行记录
    Select Max(天数) Into v_天数 From 病人路径执行 Where 路径记录id = v_路径记录id;
    If v_天数 Is Null Then
      Update 病人临床路径 Set 前一阶段id = Null, 当前阶段id = Null, 当前天数 = Null, 状态 = 1 Where ID = v_路径记录id;
    Else
      --b.回退到前一个阶段
      Select Max(阶段id)
      Into v_阶段id
      From 病人路径执行
      Where 路径记录id = v_路径记录id And
            登记时间 = (Select Max(登记时间)
                    From 病人路径执行
                    Where 路径记录id = v_路径记录id And 阶段id <> (Select 前一阶段id From 病人临床路径 Where ID = v_路径记录id));
    
      Update 病人临床路径
      Set 当前阶段id = 前一阶段id, 前一阶段id = v_阶段id, 当前天数 = v_天数, 状态 = 1
      Where ID = v_路径记录id;
    End If;
  Else
    Update 病人临床路径 Set 当前天数 = v_天数 Where ID = v_路径记录id And 当前天数 <> v_天数;
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病人路径生成_Delete;
/


Create Or Replace Procedure Zl_病人路径执行_Update
(
  执行人_In   病人路径执行.执行人%Type,
  执行时间_In 病人路径执行.执行时间%Type,
  执行内容_In Varchar2 --ID|执行结果|执行说明||...末尾带||,当执行说明为空时,要加空格,避免和||粘上
) Is
  v_Str   Varchar2(4000);
  v_Tmp   Varchar2(1000);
  v_Index Number;
  I       Number(5) := 1;

  v_Id       t_Numlist := t_Numlist();
  v_执行结果 t_Strlist := t_Strlist();
  v_执行说明 t_Strlist := t_Strlist();

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  v_Str := 执行内容_In;
  Loop
    v_Index := Instr(v_Str, '||');
    Exit When(Nvl(v_Index, 0) = 0);
    v_Id.Extend;
    v_执行结果.Extend;
    v_执行说明.Extend;
  
    v_Tmp := Substr(v_Str, 1, v_Index - 1);
    v_Id(I) := Substr(v_Tmp, 1, Instr(v_Tmp, '|') - 1);
    v_Tmp := Substr(v_Tmp, Instr(v_Tmp, '|') + 1);
    v_执行结果(I) := Trim(Substr(v_Tmp, 1, Instr(v_Tmp, '|') - 1));
    v_Tmp := Substr(v_Tmp, Instr(v_Tmp, '|') + 1);
    v_执行说明(I) := Trim(v_Tmp);
  
    v_Str := Substr(v_Str, v_Index + 2);
    I     := I + 1;
  End Loop;

  Forall I In 1 .. v_Id.Count
    Update 病人路径执行
    Set 执行人 = 执行人_In, 执行时间 = 执行时间_In, 执行结果 = v_执行结果(I), 执行说明 = v_执行说明(I)
    Where ID = v_Id(I);

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病人路径执行_Update;
/

Create Or Replace Procedure Zl_病人路径执行_Delete(路径执行id_In 病人路径执行.Id%Type) Is
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin

  Update 病人路径执行 Set 执行人 = Null, 执行时间 = Null, 执行结果 = Null, 执行说明 = Null Where ID = 路径执行id_In;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病人路径执行_Delete;
/


Create Or Replace Procedure Zl_病人路径结束_Update(路径记录id_In 病人临床路径.Id%Type) Is
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin

  Update 病人临床路径
  Set 结束时间 = Sysdate, 状态 = 2, 前一阶段id = 当前阶段id, 当前阶段id = Null, 当前天数 = Null
  Where ID = 路径记录id_In;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病人路径结束_Update;
/


Create Or Replace Procedure Zl_病人路径结束_Delete
(
  路径记录id_In 病人临床路径.Id%Type,
  结束类型_In   病人临床路径.状态%Type
) Is
  v_阶段id     病人路径评估.阶段id%Type;
  v_前一阶段id 病人路径评估.阶段id%Type;
  v_日期       病人路径评估.日期%Type;
  v_天数       病人路径评估.天数%Type;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin

  Select 前一阶段id Into v_阶段id From 病人临床路径 Where ID = 路径记录id_In;
  Select Max(日期), Max(天数)
  Into v_日期, v_天数
  From 病人路径执行
  Where 路径记录id = 路径记录id_In And 阶段id = v_阶段id;

  If 结束类型_In = 3 Then
    --评估结果为变异时自动结束的,取消结束自动取消评估
    Delete 病人路径评估 Where 路径记录id = 路径记录id_In And 阶段id = v_阶段id And 日期 = v_日期;
    Delete 病人路径指标 Where 路径记录id = 路径记录id_In And 阶段id = v_阶段id And 日期 = v_日期;
  End If;

  --b.回退到前一个阶段
  Select Max(阶段id)
  Into v_前一阶段id
  From 病人路径执行
  Where 路径记录id = 路径记录id_In And
        登记时间 = (Select Max(登记时间) From 病人路径执行 Where 路径记录id = 路径记录id_In And 阶段id <> v_阶段id);

  Update 病人临床路径
  Set 结束时间 = Null, 状态 = 1, 前一阶段id = v_前一阶段id, 当前阶段id = v_阶段id, 当前天数 = v_天数
  Where ID = 路径记录id_In;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病人路径结束_Delete;
/

Create Or Replace Function Zl_Getpathcharge
(
  病人id_In   病案主页.病人id%Type, --不确定病人时传入0
  主页id_In   病案主页.主页id%Type, --不确定病人时传入0
  路径id_In   临床路径项目.路径id%Type,
  版本号_In   临床路径项目.版本号%Type,
  阶段id_In   临床路径项目.阶段id%Type, --没有指定阶段时,根据当前天数来确定缺省的阶段
  天数_In     病人路径执行.天数%Type, --当前阶段正在执行的天数
  入院时间_In Date --病人入院或入科时间,用于计算上次执行时间(频率为每n天m次时),不确定病人时传入当前系统时间
) Return Number As
  v_Error Varchar2(255);
  Err_Custom Exception;

  v_Tmp Varchar2(1000);
  n_Tmp Number(8);

  v_费别         病案主页.费别%Type;
  n_实收金额     Number(16, 5);
  n_应收金额     Number(16, 5);
  n_主项金额     Number(16, 5);
  n_实收合计     Number(16, 5);
  n_计费总量     Number(16, 5);
  n_总量         Number(16, 5);
  n_汇总计算折扣 Number(1);
  n_主收入id     Number(8);
  n_次数         Number(8);
  n_Day          Number(8); --当前天数的星期数
  n_Lastday      Number(8);

  l_采集方法 Boolean;
  l_中药煎法 Boolean;
  l_中药用法 Boolean;
  l_给药途径 Boolean;
  l_输血途径 Boolean;

  v_Lasttype   诊疗项目目录.类别%Type;
  n_Lastsum    路径医嘱内容.总给予量%Type;
  n_Last相关id 路径医嘱内容.相关id%Type;
  n_Lastid     路径医嘱内容.Id%Type;
  n_Lastamount Number(8);
  n_Last付数   Number(8);
  l_Last煎法   Boolean;
  l_Do         Boolean;
  l_Firstday   Boolean;

  n_阶段id     临床路径阶段.Id%Type;
  n_前一阶段id 临床路径阶段.Id%Type;
  l_Rate       t_Strlist;

  --取药品相关信息(未明确规格时,取其中一个规格)
  Cursor Mediinfo
  (
    诊疗项目id_In Number,
    收费细目id_In Number
  ) Is
    Select g.Id As 收费细目id, Nvl(f.剂量系数, 1) As 剂量系数, f.可否分零, Nvl(g.是否变价, 0) 是否变价, h.缺省价格, h.现价, g.屏蔽费别, h.收入项目id,
           Nvl(h.附术收费率, 1) 附术收费率
    From 药品规格 F, 收费项目目录 G, 收费价目 H
    Where f.药名id = 诊疗项目id_In And f.药品id = Nvl(收费细目id_In, f.药品id) And f.药品id = g.Id And g.Id = h.收费细目id And
          Sysdate Between h.执行日期 And Nvl(h.终止日期, Sysdate + 1)
    Order By g.编码;
  r_Medi Mediinfo%Rowtype;

  --功能:获取指天数所属的缺省时间阶段ID  
  Function Getphaseid(n_Day Number) Return Number As
    n_Id 临床路径阶段.Id%Type;
  Begin
    For R In (Select ID
              From 临床路径阶段
              Where n_Day Between Nvl(开始天数, n_Day) And Nvl(结束天数, Nvl(开始天数, n_Day)) And 路径id = 路径id_In And 版本号 = 版本号_In
              Order By Decode(父id, Null, 0, 1), 序号) Loop
      n_Id := r.Id;
      Exit;
    End Loop;
    Return n_Id;
  End Getphaseid;

  --功能:获取指时间阶段(不能是分支)的前一时间阶段id 
  Function Getprephaseid(n_阶段id 临床路径阶段.Id%Type) Return Number As
    n_Id 临床路径阶段.Id%Type;
  Begin
    Select Nvl(Max(b.Id), 0)
    Into n_Id
    From 临床路径阶段 A, 临床路径阶段 B
    Where a.路径id = b.路径id And a.版本号 = b.版本号 And a.Id = n_阶段id And b.序号 = a.序号 - 1;
    Return n_Id;
  End Getprephaseid;

  --功能:获取指定路径项目的开始执行天数(入院时间为第一天)
  Function Getitembeginday(n_路径项目id 临床路径项目.Id%Type) Return Number As
    n_Preday Number(8);
    n_Id     临床路径阶段.Id%Type;
    n_Preid  临床路径阶段.Id%Type;
    n_Tmp    Number(8);
    n_Return Number(8);
  Begin
    n_Preday := 天数_In - 1;
    If n_Preday = 0 Or n_前一阶段id = 0 Then
      --当前是第一天或第一个阶段
      n_Return := 1;
    Else
      n_Id    := n_阶段id;
      n_Preid := n_前一阶段id;
      Loop
        --检查前一阶段是否有相同的路径项目
        Select Nvl(Count(p.Id), 0)
        Into n_Tmp
        From 临床路径项目 T, 临床路径项目 P
        Where t.路径id = p.路径id And t.版本号 = p.版本号 And t.项目内容 = p.项目内容 And t.Id = n_路径项目id And p.阶段id = n_Preid;
        If n_Tmp = 0 Then
          Exit;
        End If;
      
        n_Id := n_Preid; --如果有,继续往前找
        Select Nvl(Max(b.Id), 0)
        Into n_Preid
        From 临床路径阶段 A, 临床路径阶段 B
        Where a.路径id = b.路径id And a.版本号 = b.版本号 And a.Id = n_Id And b.序号 = a.序号 - 1;
        If n_Preid = 0 Then
          Exit;
        End If;
      End Loop;
    
      Select Nvl(开始天数, 0) Into n_Tmp From 临床路径阶段 Where ID = n_Id;
      If n_Tmp = 0 Then
        --不定期间的两个阶段不可能连续,所以取前一个阶段的结束天数+1        
        Select Nvl(Nvl(结束天数, 开始天数), 0) + 1 Into n_Tmp From 临床路径阶段 Where ID = n_Preid;
      End If;
      If n_Tmp <= 1 Then
        n_Return := 1;
      Else
        n_Return := n_Tmp;
      End If;
    End If;
    Return n_Return;
  End Getitembeginday;

  --功能:获取时价药品的应收金额(因为是估算,不管出库模式  )
  Function Get时价药品金额
  (
    n_总数量     In Number,
    n_执行科室id In Number,
    n_收费细目id In Number
  ) Return Number As
    n_总金额   Number(16, 5);
    n_可用数量 Number(16, 5);
    n_本次数量 Number(16, 5);
  Begin
    n_总金额   := 0;
    n_可用数量 := n_总数量;
    For D In (Select Nvl(可用数量, 0) As 库存, Nvl(零售价, Nvl(Decode(Nvl(实际数量, 0), 0, 0, 实际金额 / 实际数量), 0)) As 时价
              From 药品库存
              Where 库房id = n_执行科室id And 药品id = n_收费细目id And Nvl(可用数量, 0) > 0 And 性质 = 1 And
                    (Nvl(批次, 0) = 0 Or 效期 Is Null Or 效期 > Trunc(Sysdate))
              Order By Nvl(批次, 0)) Loop
      If n_可用数量 <= d.库存 Then
        n_本次数量 := n_可用数量;
      Else
        n_本次数量 := d.库存;
      End If;
      n_总金额   := n_总金额 + n_本次数量 * d.时价;
      n_可用数量 := n_可用数量 - n_本次数量;
      If n_可用数量 = 0 Then
        Exit;
      End If;
    End Loop;
    Return n_总金额;
  End Get时价药品金额;
Begin
  Select To_Char(Sysdate, 'D') - 1 Into n_Day From Dual;
  If n_Day = 0 Then
    n_Day := 7;
  End If;

  If Nvl(阶段id_In, 0) = 0 Then
    n_阶段id := Getphaseid(天数_In);
    n_Tmp    := n_阶段id;
  Else
    n_阶段id := 阶段id_In; --如果当前是分支,则求缺省分支    
    Select Nvl(父id, ID) Into n_Tmp From 临床路径阶段 Where ID = n_阶段id;
  End If;
  n_前一阶段id := Getprephaseid(n_Tmp);

  l_Firstday := False;
  Select Nvl(开始天数, 0) Into n_Tmp From 临床路径阶段 Where ID = n_阶段id;
  If n_Tmp = 0 Then
    --不定期间的两个阶段不可能连续,所以取前一个阶段的结束天数+1 
    Select Nvl(Nvl(结束天数, 开始天数), 0) + 1 Into n_Tmp From 临床路径阶段 Where ID = n_前一阶段id;
  End If;
  If n_Tmp = 天数_In Then
    l_Firstday := True;
  End If;

  If Nvl(病人id_In, 0) <> 0 Then
    Select 费别 Into v_费别 From 病案主页 Where 病人id = 病人id_In And 主页id = 主页id_In;
  End If;
  n_汇总计算折扣 := Nvl(Zl_Getsysparameter(93), 0);
  n_实收合计     := 0;

  --院外执行和无执行的叮嘱除外,不计价的除外
  For R In (Select Nvl(c.相关id, c.Id) As 组id, Nvl(e.序号, c.序号) As 组号, c.序号, a.Id, c.相关id, c.期效, d.类别, c.诊疗项目id, c.收费细目id,
                   c.执行科室id, c.标本部位, c.检查方法, Nvl(c.单次用量, 1) As 单次用量, c.总给予量, c.执行频次, c.频率次数, c.频率间隔, c.间隔单位, c.执行性质,
                   c.时间方案, Nvl(d.计算规则, 0) As 计算规则, a.执行方式
            From 临床路径项目 A, 临床路径医嘱 B, 路径医嘱内容 C, 诊疗项目目录 D, 路径医嘱内容 E
            Where a.路径id = 路径id_In And a.版本号 = 版本号_In And a.阶段id = n_阶段id And a.执行方式 Not In (0, 3) And a.Id = b.路径项目id And
                  b.医嘱内容id = c.Id And c.诊疗项目id = d.Id And c.执行性质 Not In (0, 5) And d.计价性质 <> 1 And c.相关id = e.Id(+)
            Order By a.项目序号, 组号, 组id, c.序号) Loop
    l_Do := True;
    If r.执行方式 = 2 Then
      l_Do := l_Firstday; --至少执行一次的项目,仅在本阶段的第一天时计算
    End If;
    If l_Do Then
      --1.计算总量
      n_次数     := 0;
      l_中药煎法 := False;
      l_输血途径 := False;
      l_中药用法 := False;
      l_采集方法 := False;
      l_给药途径 := False;
      If r.类别 = 'E' And r.相关id Is Not Null And r.相关id = n_Last相关id Then
        If v_Lasttype = '7' Then
          l_中药煎法 := True;
        Elsif v_Lasttype = 'K' Then
          l_输血途径 := True;
        End If;
      Elsif r.类别 = 'E' And r.相关id Is Null And r.Id = n_Last相关id Then
        If l_Last煎法 Or v_Lasttype = '7' Then
          l_中药用法 := True;
        Elsif v_Lasttype = 'C' Then
          l_采集方法 := True;
        Elsif v_Lasttype In ('5', '6') Then
          l_给药途径 := True;
        End If;
      End If;
      If r.类别 In ('5', '6', '7') Then
        Open Mediinfo(r.诊疗项目id, r.收费细目id);
        Fetch Mediinfo
          Into r_Medi;
        Close Mediinfo;
      End If;
    
      --长嘱 
      If r.期效 = 0 Then
        --a.主要医嘱或一并采集的检验项目,或药品(因为是估算,不考虑动态分零对数量的影响)
        If (r.相关id Is Null And Not l_采集方法 And Not l_中药用法 And Not l_给药途径) Or (r.相关id Is Not Null And r.类别 = 'C') Or
           r.类别 In ('5', '6', '7') Then
        
          If r.时间方案 Is Null And (Nvl(r.频率次数, 0) = 0 Or Nvl(r.频率间隔, 0) = 0 Or r.间隔单位 Is Null) Then
            n_次数 := 1; --持续性项目 --因为是估算,简化为不考虑起止时间,按每天一次算
          Else
            Select Column_Value Bulk Collect Into l_Rate From Table(f_Str2list(r.时间方案, '-')); --执行频率为"可选频率"的项目             
            Case r.间隔单位
              When '周' Then
                --例:每周三次： 1/8:00-3/8:00-5/8:00或1/8-3/8-5/8                  
                For I In 1 .. l_Rate.Count Loop
                  If n_Day = Substr(l_Rate(I), 1, Instr(l_Rate(I), '/') - 1) Then
                    n_次数 := n_次数 + 1;
                  End If;
                End Loop;
              When '天' Then
                --例:每天三次：8:00-12:00-16:00 或 8:12:16                     
                If r.频率间隔 = 1 Then
                  If 天数_In = 1 Then
                    For I In 1 .. l_Rate.Count Loop
                      If 入院时间_In <= To_Date(To_Char(入院时间_In, 'yyyy-mm-dd') || ' ' || l_Rate(I), 'yyyy-mm-dd hh24:mi') Then
                        n_次数 := n_次数 + 1; --入院当天 
                      End If;
                    End Loop;
                  Else
                    n_次数 := r.频率次数;
                  End If;
                Else
                  n_Lastday := 天数_In - Getitembeginday(r.Id); --例:两天一次： 1/8 或 1/8:00                
                  For I In 1 .. l_Rate.Count Loop
                    If n_Lastday = Substr(l_Rate(I), 1, Instr(l_Rate(I), '/') - 1) Then
                      n_次数 := n_次数 + 1;
                    End If;
                  End Loop;
                End If;
              When '小时' Then
                If 天数_In = 1 Then
                  n_次数 := Trunc(Trunc((Trunc(入院时间_In + 1) - 入院时间_In) * 24) / r.频率间隔); --入院当天          
                Else
                  n_次数 := Trunc(24 / r.频率间隔);
                End If;
              When '分钟' Then
                If 天数_In = 1 Then
                  n_次数 := Trunc(Trunc((Trunc(入院时间_In + 1) - 入院时间_In) * 24 * 60) / r.频率间隔); --入院当天     
                Else
                  n_次数 := Trunc((24 * 60) / r.频率间隔);
                End If;
            End Case;
          End If;
          If r.类别 = '7' Then
            n_次数 := r.总给予量 * n_次数;
            If r_Medi.可否分零 = 0 Then
              n_总量 := n_次数 * r.单次用量 / r_Medi.剂量系数;
            Else
              n_总量 := n_次数 * Ceil(r.单次用量 / r_Medi.剂量系数); --总给予量=付数    
            End If;
            n_Last付数 := n_次数; --中药煎法、用法的总给予量为付数 
          Elsif r.类别 = '5' Or r.类别 = '6' Then
            If r_Medi.可否分零 = 0 Then
              n_总量 := n_次数 * r.单次用量 / r_Medi.剂量系数; --可分零
            Elsif r_Medi.可否分零 = 1 Then
              n_总量 := Ceil(n_次数 * r.单次用量 / r_Medi.剂量系数); --不分零
            Elsif r_Medi.可否分零 = 2 Then
              n_总量 := n_次数 * Ceil(r.单次用量 / r_Medi.剂量系数); --一次性
            Else
              n_总量 := Ceil(n_次数 * r.单次用量 / r_Medi.剂量系数); --可否分零<0  :n天内分零使用有效,计算太复杂,按不可分零处理
            End If;
          Else
            If r.计算规则 = 1 Then
              n_总量 := Ceil(r.单次用量 * n_次数); --取最大整数--取整计算，适用于可选频率的计量、计时、计次长期医嘱。
            Else
              n_总量 := r.单次用量 * n_次数;
            End If;
          End If;
        Elsif l_中药煎法 Or l_中药用法 Then
          n_总量 := n_Last付数; --b.中药煎法、用法为付数 
          n_次数 := n_Lastamount;
        Elsif l_给药途径 Then
          n_总量 := n_Lastamount; --c.给药途径     
          n_次数 := n_Lastamount;
        Elsif l_输血途径 Then
          n_总量 := n_Lastamount; --d.输血途径的执行次数      
          n_次数 := n_Lastamount;
        Elsif r.相关id Is Not Null Or l_采集方法 Then
          n_总量 := n_Lastsum; --e.附加医嘱或标本采集方法(检查组合和手术组合不可能为长嘱,所以此段不会执行)   
          n_次数 := n_Lastamount;
        End If;
      Else
        --临嘱
        If r.类别 = '7' Then
          n_次数 := r.总给予量;
          If r_Medi.可否分零 = 0 Then
            n_总量 := r.总给予量 * r.单次用量 / r_Medi.剂量系数;
          Else
            n_总量 := r.总给予量 * Ceil(r.单次用量 / r_Medi.剂量系数); --总给予量=付数    
          End If;
        Elsif r.类别 In ('5', '6') Then
          If Nvl(r.频率次数, 0) = 0 Or Nvl(r.频率间隔, 0) = 0 Then
            n_次数 := 1; --一次性的临嘱药品
            --因为没有天数,所以不按频率计算
          Elsif r_Medi.可否分零 = 0 And Nvl(r.单次用量, 0) <> 0 Then
            --可分零药品时,按总量对单量的倍数计算给药途径的次数,否则按一个频率周期的次数计算
            n_次数 := Trunc(r.总给予量 * r_Medi.剂量系数 / r.单次用量);
          Else
            n_次数 := r.频率次数;
          End If;
          n_总量 := r.总给予量;
        Elsif l_中药煎法 Or l_中药用法 Or l_给药途径 Then
          n_总量 := n_Lastamount; --给药途径,中药用法,煎法的次数
          n_次数 := n_Lastamount;
        Elsif (r.相关id Is Null And Not l_采集方法) Or (r.相关id Is Not Null And r.类别 = 'C') Then
          --主要医嘱或一并采集的检验项目
          n_总量 := Nvl(r.总给予量, 1);
          n_次数 := Ceil(n_总量 / r.单次用量);
        Elsif l_输血途径 Then
          n_总量 := n_Lastamount; --d.输血途径的执行次数  
          n_次数 := n_Lastamount;
        Elsif r.相关id Is Not Null Or l_采集方法 Then
          n_总量 := n_Lastsum; --e.附加医嘱或标本采集方法    
          n_次数 := n_Lastamount;
        End If;
      End If;
      n_Lastamount := n_次数;
      n_Lastsum    := n_总量;
      v_Lasttype   := r.类别;
      n_Last相关id := r.相关id;
      n_Lastid     := r.Id;
      l_Last煎法   := l_中药煎法;
    
      --2.计算实收金额(不考虑加班加价)    
      n_实收金额 := 0;
      n_应收金额 := 0;
      n_主收入id := 0;
      n_主项金额 := 0;
      If r.类别 In ('4', '5', '6', '7') Then
      
        If r_Medi.是否变价 = 0 Then
          n_应收金额 := r_Medi.现价 * n_总量;
        Else
          n_应收金额 := Get时价药品金额(n_总量, r.执行科室id, r_Medi.收费细目id);
        End If;
        If Not (v_费别 Is Null Or r_Medi.屏蔽费别 = 1) Then
          v_Tmp      := Zl_Actualmoney(v_费别, r_Medi.收费细目id, r_Medi.收入项目id, n_应收金额, n_总量, r.执行科室id);
          n_实收金额 := Substr(v_Tmp, Instr(v_Tmp, ':') + 1);
        Else
          n_实收金额 := n_应收金额;
        End If;
        n_实收合计 := n_实收合计 + n_实收金额;
      
      Else
        For D In (Select c.类别, c.Id As 收费细目id, a.收费数量, b.收入项目id, Decode(c.是否变价, 1, b.缺省价格, b.现价) As 单价, c.是否变价,
                         Nvl(a.从属项目, 0) As 从项, d.跟踪在用, c.屏蔽费别, Nvl(a.费用性质, 0) As 费用性质, Nvl(a.收费方式, 0) As 收费方式, b.附术收费率
                  From 诊疗收费关系 A, 收费价目 B, 收费项目目录 C, 材料特性 D
                  Where a.诊疗项目id = r.诊疗项目id And (r.类别 <> 'D' Or r.类别 = 'D' And a.检查部位 = r.标本部位 And a.检查方法 = r.检查方法) And
                        a.收费项目id = b.收费细目id And a.收费项目id = c.Id And a.收费项目id = d.材料id(+) And c.服务对象 In (2, 3) And
                        (c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.撤档时间 Is Null) And Sysdate Between b.执行日期 And
                        Nvl(b.终止日期, Sysdate + 1) And
                        (a.收费方式 = 1 And c.类别 = '4' And a.收费项目id = r_Medi.收费细目id Or Not (a.收费方式 = 1 And c.类别 = '4'))
                  Order By 费用性质, 从项, a.收费项目id) Loop
          n_计费总量 := n_总量 * d.收费数量;
        
          If d.是否变价 = 1 And (d.类别 In ('5', '6', '7') Or (d.类别 = '4' And d.跟踪在用 = 1)) Then
            n_应收金额 := Get时价药品金额(n_计费总量, r.执行科室id, d.收费细目id); --时价非药嘱药品或跟踪在用的卫材    
          Elsif r.类别 = 'F' And r.相关id Is Not Null Then
            n_应收金额 := d.单价 * Nvl(d.附术收费率, 100) / 100 * n_计费总量;
          Else
            n_应收金额 := d.单价 * n_计费总量;
          End If;
          If n_应收金额 <> 0 Then
            If d.从项 = 0 And n_汇总计算折扣 = 1 And n_主收入id = 0 Then
              n_主收入id := d.收入项目id; --SQL中主项排在前面,只取主项目的第一个收入
            End If;
          
            If n_主收入id <> 0 Then
              n_主项金额 := n_主项金额 + n_应收金额;
              n_实收金额 := n_应收金额;
            Elsif v_费别 Is Null Or d.屏蔽费别 = 1 Then
              n_实收金额 := n_应收金额;
            Else
              v_Tmp      := Zl_Actualmoney(v_费别, d.收费细目id, d.收入项目id, n_应收金额, n_计费总量, r.执行科室id);
              n_实收金额 := Substr(v_Tmp, Instr(v_Tmp, ':') + 1);
            End If;
            n_实收合计 := n_实收合计 + n_实收金额;
          End If;
        End Loop;
        If n_主收入id <> 0 And n_主项金额 <> 0 Then
          v_Tmp      := Zl_Actualmoney(v_费别, 0, n_主收入id, n_主项金额);
          n_实收金额 := Substr(v_Tmp, Instr(v_Tmp, ':') + 1);
          n_实收合计 := n_实收合计 + n_实收金额;
        End If;
      End If;
    End If;
  End Loop;
  Return n_实收合计;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Getpathcharge;
/


Create Or Replace Procedure Zl_病人变动记录_Change
(
  病人id_In     病案主页.病人id%Type,
  主页id_In     病案主页.主页id%Type,
  转入科室id_In 病人变动记录.科室id%Type,
  操作员编号_In 病人变动记录.操作员编号%Type,
  操作员姓名_In 病人变动记录.操作员姓名%Type
) As
  -----------------------------------------------------------
  --说明：病人转科登记
  -----------------------------------------------------------    
  v_Count Number;
  v_年龄  病人信息.年龄%Type;
  Err_Custom Exception;
  v_Error Varchar2(255);
Begin
  --首先判断该病人是否处于等待转科或入科状态
  Select Count(*)
  Into v_Count
  From 病案主页
  Where 病人id = 病人id_In And 主页id = 主页id_In And 出院日期 Is Null And Nvl(状态, 0) In (0, 3);
  If v_Count = 0 Then
    v_Error := '操作失败,该病人正处于转科状态或尚未入科,不能转科。';
    Raise Err_Custom;
  End If;
  
   --临床路径正在执行时不允许转科
  Select Max(b.状态)
  Into v_Count
  From 病案主页 A, 病人临床路径 B
  Where a.病人id = 病人id_In And a.主页id = 主页id_In And a.病人id = b.病人id And a.主页id = b.主页id And a.出院科室id = b.科室id;
  If v_Count = 1 Then
    v_Error := '该病人的临床路径正在执行中,不能转科。';
    Raise Err_Custom;
  End If;

  --不填写开始时间和终止时间
  Insert Into 病人变动记录
    (ID, 病人id, 主页id, 开始时间, 开始原因, 病区id, 科室id, 操作员编号, 操作员姓名)
  Values
    (病人变动记录_Id.Nextval, 病人id_In, 主页id_In, Null, 3, Null, 转入科室id_In, 操作员编号_In, 操作员姓名_In);

  Update 病案主页
  Set 状态 = 2, 年龄 = Zl_Age_Calc(病人id)
  Where 病人id = 病人id_In And 主页id = 主页id_In Returning 年龄 Into v_年龄;

  Update 病人信息 Set 年龄 = v_年龄 Where 病人id = 病人id_In;

  --并发操作检查
  Select Count(*)
  Into v_Count
  From 病人变动记录
  Where 病人id = 病人id_In And 主页id = 主页id_In And 开始时间 Is Null And 终止时间 Is Null;
  If v_Count > 1 Then
    v_Error := '发现病人存在非法的变动记录,当前操作不能继续！' || Chr(13) || Chr(10) ||
               '这可能是由于网络并发操作引起的,请刷新病人状态后再试！';
    Raise Err_Custom;
  End If;

  Select Count(*) Into v_Count From 病案主页 Where 病人id = 病人id_In And 主页id = 主页id_In And 出院日期 Is Null;
  If v_Count = 0 Then
    v_Error := '操作失败,该病人已出院,不能进行当前操作.' || Chr(13) || Chr(10) ||
               '这可能是由于网络并发操作引起的,请刷新病人状态！';
    Raise Err_Custom;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病人变动记录_Change;
/

Create Or Replace Procedure Zl_病人变动记录_Out
(
  病人id_In       病案主页.病人id%Type,
  主页id_In       病案主页.主页id%Type,
  疾病id_In       病人诊断记录.疾病id%Type,
  诊断id_In       病人诊断记录.诊断id%Type,
  出院诊断_In     病人诊断记录.诊断描述%Type,
  出院情况_In     病人诊断记录.出院情况%Type,
  中医疾病id_In   病人诊断记录.疾病id%Type,
  中医诊断id_In   病人诊断记录.诊断id%Type,
  中医诊断_In     病人诊断记录.诊断描述%Type,
  中医出院情况_In 病人诊断记录.出院情况%Type,
  是否疑诊_In     病案主页.是否确诊%Type, --同时作为西医的是否疑诊
  出院方式_In     病案主页.出院方式%Type,
  出院时间_In     病案主页.出院日期%Type,
  随诊标志_In     病案主页.随诊标志%Type, --0/NULL-不随诊，1-月，2-年，3-周，4-天，9-终身
  随诊期限_In     病案主页.随诊期限%Type,
  尸检标志_In     病案主页.尸检标志%Type,
  操作员编号_In   病人变动记录.操作员编号%Type,
  操作员姓名_In   病人变动记录.操作员姓名%Type
) As
  -----------------------------------------------------------
  --说明：病人出院
  -----------------------------------------------------------
  Cursor c_Bedinfo Is
    Select 病区id, 床号 From 床位状况记录 Where 病人id = 病人id_In;

  v_Count Number;
  Err_Custom Exception;
  v_Error    Varchar2(255);
  v_随诊期限 Date;
  v_应发时间 Date;
  v_共享号   zlSystems.共享号%Type;
  v_年龄     病人信息.年龄%Type;
  v_Sql      Varchar2(1000);
  v_出院科室 Number;
Begin
  --首先判断该病人是否已出院
  Select Count(*) Into v_Count From 病案主页 Where 病人id = 病人id_In And 主页id = 主页id_In And 出院日期 Is Null;

  If v_Count = 0 Then
    v_Error := '操作失败,该病人可能已经出院！';
    Raise Err_Custom;
  End If;

  --首先判断该病人是否处于等待转科或入科状态
  Select Count(*)
  Into v_Count
  From 病案主页
  Where 病人id = 病人id_In And 主页id = 主页id_In And 出院日期 Is Null And Nvl(状态, 0) In (0, 3);
  If v_Count = 0 Then
    v_Error := '操作失败,该病人正处于转科状态或尚未入科,不能出院！';
    Raise Err_Custom;
  End If;
  
   --临床路径正在执行时不允许出院
  Select Max(b.状态)
  Into v_Count
  From 病案主页 A, 病人临床路径 B
  Where a.病人id = 病人id_In And a.主页id = 主页id_In And a.病人id = b.病人id And a.主页id = b.主页id And a.出院科室id = b.科室id;
  If v_Count = 1 Then
    v_Error := '该病人的临床路径正在执行中,不能出院。';
    Raise Err_Custom;
  End If;

  --判断是否产生住院日报
  Select Nvl(出院科室id, 入院科室id) Into v_出院科室 From 病案主页 Where 病人id = 病人id_In And 主页id = 主页id_In;
  Select Zl_住院日报_Count(v_出院科室, 出院时间_In) Into v_Count From Dual;
  If v_Count > 0 Then
    v_Error := '已产生业务时间内的住院日报,不能办理该业务!';
    Raise Err_Custom;
  End If;

  --判断是否与病案系统共享
  Begin
    Select 共享号 Into v_共享号 From zlSystems Where Floor(编号 / 100) = 3;
  Exception
    When Others Then
      Null;
  End;
  --出院变动
  Update 病人变动记录
  Set 终止时间 = 出院时间_In, 终止原因 = 1, 终止人员 = 操作员姓名_In
  Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null;

  --床位记录
  For r_Bedrow In c_Bedinfo Loop
    Update 床位状况记录
    Set 状态 = '空床', 病人id = Null, 科室id = Decode(共用, 1, Null, 科室id)
    Where 病区id = r_Bedrow.病区id And 床号 = r_Bedrow.床号;
  End Loop;

  --病案主页
  Update 病案主页
  Set 状态 = 0, 出院日期 = 出院时间_In, 出院方式 = 出院方式_In,
      住院天数 = Decode(Trunc(出院时间_In) - Trunc(入院日期), 0, 1, Trunc(出院时间_In) - Trunc(入院日期)), 随诊标志 = 随诊标志_In,
      随诊期限 = Decode(随诊期限_In, 0, Null, 随诊期限_In), 尸检标志 = 尸检标志_In, 是否确诊 = Decode(Nvl(是否疑诊_In, 0), 0, 1, 0),
      年龄 = Zl_Age_Calc(病人id), 病案状态 = Null
  Where 病人id = 病人id_In And 主页id = 主页id_In
  Returning 年龄 Into v_年龄;

  --增加随诊记录
  If v_共享号 = 100 Then
    If Nvl(随诊期限_In, 0) <> 0 Then
      If 随诊标志_In = 1 Then
        v_随诊期限 := Add_Months(出院时间_In, 随诊期限_In);
      Elsif 随诊标志_In = 2 Then
        v_随诊期限 := Add_Months(出院时间_In, 12 * 随诊期限_In);
      Elsif 随诊标志_In = 3 Then
        v_随诊期限 := 出院时间_In + 7 * 随诊期限_In;
      Elsif 随诊标志_In = 4 Then
        v_随诊期限 := 出院时间_In + 随诊期限_In;
      End If;
    Else
      v_随诊期限 := To_Date('3000-1-1', 'YYYY-MM-DD');
    End If;
  
    If 随诊标志_In = 1 Or 随诊标志_In = 2 Or 随诊标志_In = 3 Or 随诊标志_In = 4 Then
      If v_随诊期限 > Add_Months(出院时间_In, 3) Then
        v_应发时间 := Trunc(Add_Months(出院时间_In, 3));
      Else
        v_应发时间 := v_随诊期限;
      End If;
      v_Sql := 'Insert Into 随诊记录 (ID, 病人id, 主页id, 随诊期限, 应发时间) Values (随诊记录_Id.Nextval,' || 病人id_In || ',' || 主页id_In ||
               ',TO_DATE(''' || To_Char(v_随诊期限, 'YYYY-MM-DD') || ''',' || '''YYYY-MM-DD ''' || '),' || 'TO_DATE(''' ||
               To_Char(v_应发时间, 'YYYY-MM-DD') || ''',' || '''YYYY-MM-DD ''' || '))';
      Execute Immediate v_Sql;
    End If;
  End If;
  --病人信息
  Update 病人信息
  Set 当前科室id = Null, 当前病区id = Null, 当前床号 = Null, 出院时间 = 出院时间_In, 年龄 = v_年龄, 在院 = Null
  Where 病人id = 病人id_In;

  --出院诊断
  If 出院诊断_In Is Not Null Or 疾病id_In Is Not Null Or 诊断id_In Is Not Null Then
    Delete From 病人诊断记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 诊断类型 = 3 And 诊断次序 = 1 And 记录来源 = 2;
    Insert Into 病人诊断记录
      (ID, 病人id, 主页id, 记录来源, 诊断类型, 诊断次序, 疾病id, 诊断id, 诊断描述, 是否疑诊, 出院情况, 记录日期, 记录人)
    Values
      (病人诊断记录_Id.Nextval, 病人id_In, 主页id_In, 2, 3, 1, 疾病id_In, 诊断id_In, 出院诊断_In, 是否疑诊_In, 出院情况_In, Sysdate, 操作员姓名_In);
  End If;

  --中医出院诊断
  If 中医诊断_In Is Not Null Or 中医疾病id_In Is Not Null Or 中医诊断id_In Is Not Null Then
    Delete From 病人诊断记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 诊断类型 = 13 And 诊断次序 = 1 And 记录来源 = 2;
    Insert Into 病人诊断记录
      (ID, 病人id, 主页id, 记录来源, 诊断类型, 诊断次序, 疾病id, 诊断id, 诊断描述, 是否疑诊, 出院情况, 记录日期, 记录人)
    Values
      (病人诊断记录_Id.Nextval, 病人id_In, 主页id_In, 2, 13, 1, 中医疾病id_In, 中医诊断id_In, 中医诊断_In, Null, 中医出院情况_In, Sysdate, 操作员姓名_In);
  End If;

  --PDA同步日志写入
  If Zl_Pda_Enabled > 0 Then
    Zl_Pdasynch_Log(1, 病人id_In, 2);
  End IF;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人变动记录_Out;
/

