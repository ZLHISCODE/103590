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
    PCTFREE 5
    PCTUSED 85;
Alter Table 病人临床路径 Add Constraint 病人临床路径_PK Primary Key (ID) Using Index Pctfree 5;
Create Index 病人临床路径_IX_病人ID On 病人临床路径(病人ID,主页ID) Pctfree 5
/
Create Index 病人临床路径_IX_科室ID On 病人临床路径(科室ID) Pctfree 5
/
Create Index 病人临床路径_IX_路径ID On 病人临床路径(路径ID,版本号) Pctfree 5
/
Create Index 病人临床路径_IX_导入时间 On 病人临床路径(导入时间) Pctfree 5
/


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
    PCTFREE 5
    PCTUSED 85;
Alter Table 病人路径执行 Add Constraint 病人路径执行_PK Primary Key (ID) Using Index Pctfree 5;
Alter Table 病人路径执行 Add Constraint 病人路径执行_UQ_项目内容 Unique (路径记录ID,阶段ID,日期,项目ID,项目内容) Using Index Pctfree 5;
Create Index 病人路径执行_IX_日期 On 病人路径执行(日期) Pctfree 5
/
Create Index 病人路径执行_IX_路径记录ID On 病人路径执行(路径记录ID) Pctfree 5
/
Create Index 病人路径执行_IX_阶段ID On 病人路径执行(阶段ID) Pctfree 5
/
Create Index 病人路径执行_IX_项目ID On 病人路径执行(项目ID) Pctfree 5
/
Create Index 病人路径执行_IX_图标ID On 病人路径执行(图标ID) Pctfree 5
/
Create Index 病人路径执行_IX_登记时间 On 病人路径执行(登记时间) Pctfree 5
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
    PCTFREE 5
    PCTUSED 85;
Alter Table 病人路径评估 Add Constraint 病人路径评估_PK Primary Key (路径记录ID,阶段ID,日期) Using Index Pctfree 5;
Create Index 病人路径评估_IX_日期 On 病人路径评估(日期) Pctfree 5
/
Create Index 病人路径评估_IX_登记时间 On 病人路径评估(登记时间) Pctfree 5
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
    PCTFREE 5
    PCTUSED 85;
Alter Table 病人路径指标 Add Constraint 病人路径指标_UQ_评估指标 Unique (路径记录ID,阶段ID,日期,评估指标) Using Index Pctfree 5;
Create Index 病人路径指标_IX_日期 On 病人路径指标(日期) Pctfree 5
/

CREATE TABLE 病人路径医嘱(
		路径执行ID NUMBER(18),
    病人医嘱ID NUMBER(18))
    PCTFREE 5
    PCTUSED 85;
Alter Table 病人路径医嘱 Add Constraint 病人路径医嘱_PK Primary Key (路径执行ID,病人医嘱ID) Using Index Pctfree 5;

--对原电子病历记录的更改
Alter Table 电子病历记录 Add 路径执行ID Number(18);
Create Index 电子病历记录_IX_路径执行ID On 电子病历记录(路径执行ID) Pctfree 5
/