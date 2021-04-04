ALTER TABLE 预约方式 add 预约天数 number(5);

ALTER  Table 时间段 modify 时间段 varchar2(10);

ALTER  Table 时间段 ADD (
   站点 varchar2(1),
   号类 varchar2(10),
   出诊预留时间 number(18),
   休息时段 varchar2(200));
Alter Table 时间段  drop Constraint 时间段_PK   Cascade Drop Index;
Alter Table 时间段 Add Constraint 时间段_UQ_时间段 Unique (时间段,号类,站点) Using Index Tablespace zl9Indexhis;
Alter Table 时间段 Modify 时间段 Constraint 时间段_NN_时间段 Not Null;

Create Table 常用停诊原因(
   编码 varchar2(5),
   名称 varchar2(50),
   简码 varchar2(20),
   缺省标志 number(1) default 0)
TABLESPACE zl9BaseItem ;

Alter Table 常用停诊原因  Add Constraint 常用停诊原因_PK  Primary Key (编码) Using Index Tablespace zl9Indexhis;
Alter Table 常用停诊原因 Add Constraint 常用停诊原因_UQ_名称 Unique (名称) Using Index Tablespace zl9Indexhis;

Create Sequence 门诊诊室_ID start with 1;
ALTER TABLE 门诊诊室 add(ID number(18));

 
Declare
  Cursor c_门诊诊室 Is
    Select 门诊诊室_Id.Nextval ID, 名称 From 门诊诊室 Where ID Is Null;
  n_Array_Size Number := 200;

  t_Id       t_Numlist;
  t_门诊诊室 t_Strlist;
Begin

  Open c_门诊诊室;

  Loop
    Fetch c_门诊诊室 Bulk Collect
      Into t_Id, t_门诊诊室 Limit n_Array_Size;
    Exit When t_门诊诊室.Count = 0;
  
    --循环处理门诊费用记录
    Forall I In 1 .. t_门诊诊室.Count
      Update 门诊诊室 Set ID = t_Id(I) Where 名称 = t_门诊诊室(I);
  End Loop;
  COMMIT ;
  Close c_门诊诊室;
End;
/

Alter Table 门诊诊室  drop Constraint 门诊诊室_PK   Cascade Drop Index;
Alter Table 门诊诊室  Add Constraint 门诊诊室_PK  Primary Key (ID) Using Index Tablespace zl9Indexhis;
Alter Table 门诊诊室 Add Constraint 门诊诊室_UQ_编码 Unique (编码) Using Index Tablespace zl9Indexhis;


CREATE TABLE 门诊诊室适用科室 (
	诊室ID number(18),
	科室ID number(18),
	缺省标志 number(2)) 
TABLESPACE zl9BaseItem ;

Alter Table 门诊诊室适用科室  Add Constraint 门诊诊室适用科室_PK  Primary Key (诊室ID,科室ID) Using Index Tablespace zl9Indexhis;
Alter Table 门诊诊室适用科室 Add Constraint 门诊诊室适用科室_FK_诊室ID Foreign Key (诊室ID) References 门诊诊室( ID) ;
Alter Table 门诊诊室适用科室 Add Constraint 门诊诊室适用科室_FK_科室ID Foreign Key (科室ID) References 部门表( ID) ;

Create Index 门诊诊室适用科室_IX_科室id on 门诊诊室适用科室(科室id) Tablespace zl9Indexhis;


Create Table 法定假日表(
   年份 number(18),
   节日名称 varchar2(50),
   性质 number(18),
   开始日期 Date,
   终止日期 Date,
   备注 varchar2(1000),
   允许挂号 varchar2(500),
   允许预约 varchar2(500))
TABLESPACE zl9BaseItem ;

Alter Table 法定假日表  Add Constraint 法定假日表_PK  Primary Key (开始日期,年份,节日名称,性质) Using Index Tablespace zl9Indexhis;

ALTER TABLE 挂号合作单位 ADD 锁号时间 number(18) ;

   

Create Sequence 临床出诊号源_ID start with 1;
Create Table 临床出诊号源(
   ID number(18) not null,
   号类 varchar2(10),
   号码 varchar2(5),
   科室id number(18),
   项目ID number(18),
   医生id number(18),
   医生姓名 varchar2(50),
   是否建病案 number(2) default 0,
   预约天数 number(3),
   出诊频次 number(3),
   假日控制状态 number(2) ,
   是否假日换休 number(2) default 0,
   是否临床排班 number(2) default 0,
   排班方式 number(2),
   是否删除 number(2) default 0,
   建档时间 Date,
   撤档时间 Date)
TABLESPACE zl9BaseItem ;

Alter Table 临床出诊号源 Add Constraint 临床出诊号源_PK  Primary Key (ID) Using Index Tablespace zl9Indexhis;
Alter Table 临床出诊号源 Add Constraint 临床出诊号源_UQ_号码 Unique (号码,撤档时间) Using Index Tablespace zl9Indexhis;
Alter Table 临床出诊号源 Add Constraint 临床出诊号源_UQ_科室项目 Unique (科室ID,项目ID,医生姓名,医生ID,撤档时间) Using Index Tablespace zl9Indexhis;
Alter Table 临床出诊号源 Add Constraint 临床出诊号源_FK_科室ID Foreign Key (科室ID) References 部门表( ID) ;
Alter Table 临床出诊号源 Add Constraint 临床出诊号源_FK_项目ID Foreign Key (项目ID) References 收费项目目录(ID) ;
Alter Table 临床出诊号源 Add Constraint 临床出诊号源_FK_医生id Foreign Key (医生id) References 人员表(ID) ;
 
Create Index 临床出诊号源_IX_项目ID on 临床出诊号源(项目ID) Tablespace zl9Indexhis;
Create Index 临床出诊号源_IX_医生id on 临床出诊号源(医生id) Tablespace zl9Indexhis;
Create Index 临床出诊号源_IX_医生姓名 on 临床出诊号源(医生姓名) Tablespace zl9Indexhis;


Create Sequence 临床出诊号源限制_ID start with 1;
Create Table 临床出诊号源限制(
   ID number(18) not null,
   号源ID number(18),
   上班时段 varchar2(10),
   限号数 number(10),
   限约数 number(10),
   是否序号控制 number(2) default 0,
   是否分时段  NUMBER(2),
   预约控制 number(2),
   是否独占 number(2) default 0,   
   分诊方式 number(3),
   诊室ID number(18))
TABLESPACE zl9BaseItem ;

Alter Table 临床出诊号源限制  Add Constraint 临床出诊号源限制_PK  Primary Key (ID) Using Index Tablespace zl9Indexhis;
Alter Table 临床出诊号源限制 Add Constraint 临床出诊号源限制_FK_号源ID Foreign Key (号源ID) References 临床出诊号源( ID) ;
Alter Table 临床出诊号源限制  Add Constraint 临床出诊号源限制_UQ_号源ID  Unique (号源ID,上班时段) Using Index Tablespace zl9Indexhis; 
Alter Table 临床出诊号源限制 Add Constraint 临床出诊号源限制_FK_诊室ID Foreign Key (诊室ID) References 门诊诊室( ID) ;
create Index 临床出诊号源限制_IX_诊室ID on 临床出诊号源限制(诊室ID) Tablespace zl9Indexhis;




Create Table 临床出诊号源诊室(
   限制ID number(18),
   诊室ID number(18))
TABLESPACE zl9BaseItem ;

Alter Table 临床出诊号源诊室  Add Constraint 临床出诊号源诊室_PK  Primary Key (限制ID,诊室ID) Using Index Tablespace zl9Indexhis;
Alter Table 临床出诊号源诊室 Add Constraint 临床出诊号源诊室_FK_限制ID Foreign Key (限制ID) References 临床出诊号源限制( ID) ;
Alter Table 临床出诊号源诊室 Add Constraint 临床出诊号源诊室_FK_诊室ID Foreign Key (诊室ID) References 门诊诊室( ID) ;
create Index 临床出诊号源诊室_IX_诊室ID on 临床出诊号源诊室(诊室ID) Tablespace zl9Indexhis;



Create Table 临床出诊号源时段(
   限制ID number(18),
   序号 number(18),
   开始时间 Date,
   终止时间 Date,
   限制数量 number(10),
   是否预约 number(2))
TABLESPACE zl9BaseItem;

Alter Table 临床出诊号源时段  Add Constraint 临床出诊号源时段_PK  Primary Key (限制ID,序号) Using Index Tablespace zl9Indexhis;
Alter Table 临床出诊号源时段 Add Constraint 临床出诊号源时段_FK_限制ID Foreign Key (限制ID) References 临床出诊号源限制( ID) ;

Create Table 临床出诊号源控制(
   限制ID number(18),
   类型 number(2),
   性质 number(2),
   名称 varchar2(50),
   序号 number(18),
   控制方式 number(2),
   数量 number(16,5))
TABLESPACE zl9BaseItem ;

Alter Table 临床出诊号源控制  Add Constraint 临床出诊号源控制_PK  Primary Key (限制ID,类型,性质,名称,序号) Using Index Tablespace zl9Indexhis;
Alter Table 临床出诊号源控制 Add Constraint 临床出诊号源控制_FK_限制ID Foreign Key (限制ID) References 临床出诊号源限制(ID);

Create Sequence 临床出诊表_ID start with 1;

Create Table 临床出诊表(
   ID number(18) not null,
   排班方式 number(18),
   出诊表名 varchar2(50),
   年份 number(4),
   月份 number(2),
   周数 number(2),
   应用范围 number(2),
   科室ID number(18),
   备注 varchar2(100),
   发布人 varchar2(50),
   发布时间 Date)
TABLESPACE zl9BaseItem ;

Alter Table 临床出诊表  Add Constraint 临床出诊表_PK  Primary Key (ID) Using Index Tablespace zl9Indexhis;
Alter Table 临床出诊表  Add Constraint 临床出诊表_UQ_出诊表名  Unique (年份,月份,周数,出诊表名,排班方式) Using Index Tablespace zl9Indexhis;
Alter Table 临床出诊表 Add Constraint 临床出诊表_FK_科室ID Foreign Key (科室ID) References 部门表(ID) ;

Create Index 临床出诊表_IX_科室ID on 临床出诊表(科室ID) Tablespace zl9Indexhis;



Create Sequence 临床出诊安排_ID start with 1;
Create Table 临床出诊安排(
   ID number(18) not null,
   出诊ID number(18),
   号源ID number(18),
   项目ID number(18),
   医生id number(18),
   医生姓名 varchar2(50),
   排班规则 number(2),
   是否周六出诊 number(2),
   是否周日出诊 number(2),
   开始时间 Date,
   终止时间 Date,
   操作员姓名 varchar2(50),
   登记时间 Date,
   原终止时间 Date)
TABLESPACE zl9BaseItem ;

Alter Table 临床出诊安排 Add Constraint 临床出诊安排_PK  Primary Key (ID) Using Index Tablespace zl9Indexhis;
Alter Table 临床出诊安排  Add Constraint 临床出诊安排_UQ_出诊ID  Unique (出诊ID,号源ID,开始时间) Using Index Tablespace zl9Indexhis;

Alter Table 临床出诊安排 Add Constraint 临床出诊安排_FK_号源ID Foreign Key (号源ID) References 临床出诊号源( ID) ;
Alter Table 临床出诊安排 Add Constraint 临床出诊安排_FK_出诊ID Foreign Key (出诊ID) References 临床出诊表( ID) ;
Alter Table 临床出诊安排 Add Constraint 临床出诊安排_FK_项目ID Foreign Key (项目ID) References 收费项目目录(ID) ;
Alter Table 临床出诊安排 Add Constraint 临床出诊安排_FK_医生id Foreign Key (医生id) References 人员表(ID);



Create Index 临床出诊安排_IX_项目ID on 临床出诊安排(项目ID) Tablespace zl9Indexhis;
Create Index 临床出诊安排_IX_医生id on 临床出诊安排(医生id) Tablespace zl9Indexhis;
Create Index 临床出诊安排_IX_号源ID on 临床出诊安排(号源ID) Tablespace zl9Indexhis;


Create Sequence 临床出诊限制_ID start with 1;
Create Table 临床出诊限制(
   ID     number(18),
   安排ID number(18),
   限制项目 varchar2(20),
   上班时段 varchar2(10),
   限号数 number(10),
   限约数 number(10),
   是否序号控制 number(2),
   是否分时段 NUMBER(2),
   预约控制 number(2),
   分诊方式 number(2),
   诊室ID number(18),
   是否独占 number(2))
TABLESPACE zl9BaseItem ;

Alter Table 临床出诊限制  Add Constraint 临床出诊限制_PK  Primary Key (ID) Using Index Tablespace zl9Indexhis;
Alter Table 临床出诊限制  Add Constraint 临床出诊限制_UQ_安排ID  Unique (安排ID,限制项目,上班时段) Using Index Tablespace zl9Indexhis;
Alter Table 临床出诊限制 Add Constraint 临床出诊限制_FK_安排ID Foreign Key (安排ID) References 临床出诊安排( ID) ;
Alter Table 临床出诊限制 Add Constraint 临床出诊限制_FK_诊室id Foreign Key (诊室id) References 门诊诊室(ID) ;
Create Index 临床出诊限制_IX_诊室ID on 临床出诊限制(诊室ID) Tablespace zl9Indexhis;



Create Table 临床出诊时段(
   限制ID number(18),
   序号 number(18),
   开始时间 Date,
   终止时间 Date,
   限制数量 number(10),
   是否预约 number(2))
TABLESPACE zl9BaseItem ;

Alter Table 临床出诊时段  Add Constraint 临床出诊时段_PK  Primary Key (限制ID,序号) Using Index Tablespace zl9Indexhis;
Alter Table 临床出诊时段 Add Constraint 临床出诊时段_FK_限制ID Foreign Key (限制ID) References 临床出诊限制( ID) ;


Create Table 临床出诊诊室(
   限制ID number(18),
   诊室ID number(18))
TABLESPACE zl9BaseItem ;
Alter Table 临床出诊诊室  Add Constraint 临床出诊诊室_PK  Primary Key (限制ID,诊室id) Using Index Tablespace zl9Indexhis;
Alter Table 临床出诊诊室 Add Constraint 临床出诊诊室_FK_限制ID Foreign Key (限制ID) References 临床出诊限制( ID) ;

Alter Table 临床出诊诊室 Add Constraint 临床出诊诊室_FK_诊室id Foreign Key (诊室id) References 门诊诊室(ID) ;
Create Index 临床出诊诊室_IX_诊室ID on 临床出诊诊室(诊室ID) Tablespace zl9Indexhis;


Create Table 临床出诊挂号控制(
   限制ID number(18),
   类型 number(2),
   性质 number(2),
   名称 varchar2(50),
   序号 number(18),
   控制方式 number(2),
   数量 number(16,5))
TABLESPACE zl9BaseItem ;

Alter Table 临床出诊挂号控制  Add Constraint 临床出诊挂号控制_PK  Primary Key (限制ID,序号,名称,类型,性质) Using Index Tablespace zl9Indexhis;
Alter Table 临床出诊挂号控制 Add Constraint 临床出诊挂号控制_FK_限制ID Foreign Key (限制ID) References 临床出诊限制( ID) ;
 

Create Sequence 临床出诊记录_ID start with 1;
Create Table 临床出诊记录(
   ID number(18) not null,
   安排ID number(18),
   号源ID number(18),
   出诊日期 Date,
   上班时段 varchar2(10),
   开始时间 Date,
   终止时间 Date,
   停诊开始时间 Date,
   停诊终止时间 Date,
   停诊原因 varchar2(50),
   缺省预约时间 Date,
   提前挂号时间 Date,
   限号数 number(10),
   已挂数 number(10),
   限约数 number(10),
   已约数 number(10),
   其中已接收 number(10),
   是否序号控制 number(2) default 0,
   是否分时段 number(2) default 0,
   预约控制 number(2),
   是否独占 number(2),
   项目ID number(18),
   科室ID number(18),
   医生id number(18),
   医生姓名 varchar2(50),
   替诊医生id number(18),
   替诊医生姓名 varchar2(50),
   分诊方式 number(2),
   诊室ID Number(18),
   是否锁定 number(2) default 0,
   是否临时出诊 number(2) default 0,
   登记人 varchar2(50),
   登记时间 Date,
   是否发布 number(2) default 0)
TABLESPACE zl9BaseItem;

Alter Table 临床出诊记录  Add Constraint 临床出诊记录_PK  Primary Key (ID) Using Index Tablespace zl9Indexhis;

Alter Table 临床出诊记录  Add Constraint 临床出诊记录_UQ_出诊日期  Unique (出诊日期,号源ID,上班时段) Using Index Tablespace zl9Indexhis;

Alter Table 临床出诊记录 Add Constraint 临床出诊记录_FK_安排ID Foreign Key (安排ID) References 临床出诊安排( ID) ;
Alter Table 临床出诊记录 Add Constraint 临床出诊记录_FK_号源ID Foreign Key (号源ID) References 临床出诊号源( ID) ;

Alter Table 临床出诊记录 Add Constraint 临床出诊记录_FK_项目ID Foreign Key (项目ID) References 收费项目目录(ID) ;
Alter Table 临床出诊记录 Add Constraint 临床出诊记录_FK_科室ID Foreign Key (科室ID) References 部门表(ID) ;
Alter Table 临床出诊记录 Add Constraint 临床出诊记录_FK_医生id Foreign Key (医生id) References 人员表(ID) ;
Alter Table 临床出诊记录 Add Constraint 临床出诊记录_FK_替诊医生id Foreign Key (替诊医生id) References 人员表(ID) ;
Alter Table 临床出诊记录 Add Constraint 临床出诊记录_FK_诊室id Foreign Key (诊室id) References 门诊诊室(ID) ;
Create Index 临床出诊记录_IX_诊室ID on 临床出诊记录(诊室ID) Tablespace zl9Indexhis;

Create Index 临床出诊记录_IX_安排ID on 临床出诊记录(安排ID) Tablespace zl9Indexhis;
Create Index 临床出诊记录_IX_号源ID on 临床出诊记录(号源ID) Tablespace zl9Indexhis;
Create Index 临床出诊记录_IX_替诊医生id on 临床出诊记录(替诊医生id) Tablespace zl9Indexhis;

Create Index 临床出诊记录_IX_医生id on 临床出诊记录(医生id) Tablespace zl9Indexhis;
Create Index 临床出诊记录_IX_项目ID on 临床出诊记录(项目ID) Tablespace zl9Indexhis;
Create Index 临床出诊记录_IX_科室ID on 临床出诊记录(科室ID) Tablespace zl9Indexhis;
Create Index 临床出诊记录_IX_开始时间 on 临床出诊记录(开始时间,号源ID) Tablespace zl9Indexhis;


Create Sequence 临床出诊停诊记录_ID start with 1;
Create Table 临床出诊停诊记录(
   ID number(18) not null,
   记录ID number(18),
   开始时间 Date,
   终止时间 Date,
   停诊原因 varchar2(50),
   替诊医生ID number(18),
   替诊医生姓名 varchar2(50),
   申请人 varchar2(50),
   申请时间 Date,
   审批人 varchar2(50),
   审批时间 Date,
   取消人 varchar2(50),
   取消时间 Date)
TABLESPACE zl9BaseItem ;

Alter Table 临床出诊停诊记录  Add Constraint 临床出诊停诊记录_PK  Primary Key (ID) Using Index Tablespace zl9Indexhis;
Alter Table 临床出诊停诊记录 Add Constraint 临床出诊停诊记录_FK_记录ID Foreign Key (记录ID) References 临床出诊记录( ID) ;
Alter Table 临床出诊停诊记录 Add Constraint 临床出诊停诊记录_FK_替诊医生ID Foreign Key (替诊医生ID) References 人员表( ID) ;

Create Index 临床出诊停诊记录_IX_记录ID on 临床出诊停诊记录(记录ID) Tablespace zl9Indexhis;
Create Index 临床出诊停诊记录_IX_替诊医生ID on 临床出诊停诊记录(替诊医生ID) Tablespace zl9Indexhis;
Create Index 临床出诊停诊记录_IX_申请时间 on 临床出诊停诊记录(申请时间) Tablespace zl9Indexhis;
Create Index 临床出诊停诊记录_IX_审批时间 on 临床出诊停诊记录(审批时间) Tablespace zl9Indexhis;

Create Table 临床出诊序号控制(
   记录ID number(18),
   序号 number(18),
   预约顺序号 number(18),
   开始时间 Date,
   终止时间 Date,
   数量 number(10),
   是否预约 number(2),
   挂号状态 number(2),
   锁号时间 Date,
   类型   number(2),
   名称 varchar2(50),
   操作员姓名 varchar2(50),
   工作站IP varchar2(20),
   工作站名称 varchar2(200),
   备注 varchar2(100))
TABLESPACE zl9BaseItem ;

Alter Table 临床出诊序号控制  Add Constraint 临床出诊序号控制_UQ_记录ID  Unique (记录ID,序号,预约顺序号) Using Index Tablespace zl9Indexhis;
Alter Table 临床出诊序号控制 Add Constraint 临床出诊序号控制_FK_记录ID Foreign Key (记录ID) References 临床出诊记录( ID) ;
Alter Table 临床出诊序号控制 Modify 记录ID Constraint 临床出诊序号控制_NN_记录ID Not Null;


Create Table 临床出诊诊室记录(
   记录ID number(18),
   诊室ID Number(18),
   当前分配 number(1) default 0)
TABLESPACE zl9BaseItem ;

Alter Table 临床出诊诊室记录 Add Constraint 临床出诊诊室记录_PK  Primary Key (记录ID,诊室ID) Using Index Tablespace zl9Indexhis;
Alter Table 临床出诊诊室记录 Add Constraint 临床出诊诊室记录_FK_记录ID Foreign Key (记录ID) References 临床出诊记录( ID) ;
Alter Table 临床出诊诊室记录 Add Constraint 临床出诊诊室记录_FK_诊室id Foreign Key (诊室id) References 门诊诊室(ID) ;
Create Index 临床出诊诊室记录_IX_诊室ID on 临床出诊诊室记录(诊室ID) Tablespace zl9Indexhis;




Create Table 临床出诊挂号控制记录(
   记录ID number(18),
   类型 number(2),
   性质 number(2),
   名称 varchar2(50),
   序号 number(18),
   控制方式 number(2),
   数量 number(16,5))
TABLESPACE zl9BaseItem ;

Alter Table 临床出诊挂号控制记录  Add Constraint 临床出诊挂号控制记录_PK  Primary Key (记录ID,名称,序号,类型,性质) Using Index Tablespace zl9Indexhis;
Alter Table 临床出诊挂号控制记录 Add Constraint 临床出诊挂号控制记录_FK_记录ID Foreign Key (记录ID) References 临床出诊记录(ID) ;


Create Sequence 病人服务信息记录_ID start with 1;
Create Table 病人服务信息记录(
   ID number(18) not null,
   通知类型 number(18),
   记录ID number(18),
   挂号ID number(18),
   号源ID number(18),
   号码 varchar2(10),
   科室ID number(18),
   项目ID number(18),
   医生ID number(18),
   医生姓名 varchar2(50),
   病人ID number(18),
   复诊方式 number(2),
   数量 number(10),
   开始时间 Date,
   终止时间 Date,
   通知原因 varchar2(100),
   登记人 varchar2(50),
   登记时间 Date,
   处理说明 varchar2(100),
   处理人 varchar2(50),
   处理时间 Date)
TABLESPACE zl9BaseItem ;

Alter Table 病人服务信息记录  Add Constraint 病人服务信息记录_PK  Primary Key (ID) Using Index Tablespace zl9Indexhis;
Alter Table 病人服务信息记录 Add Constraint 病人服务信息记录_FK_号源ID Foreign Key (号源ID) References 临床出诊号源( ID) ;
Alter Table 病人服务信息记录 Add Constraint 病人服务信息记录_FK_记录ID Foreign Key (记录ID) References 临床出诊记录( ID) ;

Alter Table 病人服务信息记录 Add Constraint 病人服务信息记录_FK_项目ID Foreign Key (项目ID) References 收费项目目录(ID) ;
Alter Table 病人服务信息记录 Add Constraint 病人服务信息记录_FK_医生id Foreign Key (医生id) References 人员表(ID) ;
Alter Table 病人服务信息记录 Add Constraint 病人服务信息记录_FK_科室ID Foreign Key (科室ID) References 部门表(ID) ;
Alter Table 病人服务信息记录 Add Constraint 病人服务信息记录_FK_病人ID Foreign Key (病人ID) References 病人信息(病人ID) ;

Create Index 病人服务信息记录_IX_登记时间 on 病人服务信息记录(登记时间) Tablespace zl9Indexhis;
Create Index 病人服务信息记录_IX_处理时间 on 病人服务信息记录(处理时间) Tablespace zl9Indexhis;

Create Index 病人服务信息记录_IX_病人ID on 病人服务信息记录(病人ID) Tablespace zl9Indexhis;
Create Index 病人服务信息记录_IX_挂号ID on 病人服务信息记录(挂号ID) Tablespace zl9Indexhis;
Create Index 病人服务信息记录_IX_号码ID on 病人服务信息记录(号码) Tablespace zl9Indexhis;
Create Index 病人服务信息记录_IX_号源ID on 病人服务信息记录(号源ID) Tablespace zl9Indexhis;
Create Index 病人服务信息记录_IX_记录ID on 病人服务信息记录(记录ID) Tablespace zl9Indexhis;
Create Index 病人服务信息记录_IX_科室ID on 病人服务信息记录(科室ID) Tablespace zl9Indexhis;
Create Index 病人服务信息记录_IX_项目ID on 病人服务信息记录(项目ID) Tablespace zl9Indexhis;
Create Index 病人服务信息记录_IX_医生ID on 病人服务信息记录(医生ID) Tablespace zl9Indexhis;



Create Sequence 临床出诊变动记录_ID start with 1;
Create Table 临床出诊变动记录(
   ID number(18) not null,
   记录ID number(18),
   变动类型 number(2),
   原预约控制 number(2),
   现预约控制 number(2),
   原数量 number(10),
   现数量 number(10),
   原分诊方式 number(2),
   原门诊诊室 varchar2(20),
   原诊室ID number(18),
   现分诊方式 number(2),
   现门诊诊室 varchar2(20),
   现诊室ID number(18),
   操作员姓名 varchar2(50),
   登记时间 Date)
TABLESPACE zl9BaseItem ;

Alter Table 临床出诊变动记录  Add Constraint 临床出诊变动记录_PK  Primary Key (ID) Using Index Tablespace zl9Indexhis;
Alter Table 临床出诊变动记录 Add Constraint 临床出诊变动记录_FK_记录ID Foreign Key (记录ID) References 临床出诊记录( ID) ;
Create Index 临床出诊变动记录_IX_记录ID on 临床出诊变动记录(记录ID) Tablespace zl9Indexhis;
Create Index 临床出诊变动记录_IX_登记时间 on 临床出诊变动记录(登记时间) Tablespace zl9Indexhis;

Alter Table 临床出诊变动记录 Add Constraint 临床出诊变动记录_FK_原诊室id Foreign Key (原诊室id) References 门诊诊室(ID) ;
Create Index 临床出诊变动记录_IX_原诊室ID on 临床出诊变动记录(原诊室ID) Tablespace zl9Indexhis;

Alter Table 临床出诊变动记录 Add Constraint 临床出诊变动记录_FK_现诊室id Foreign Key (现诊室id) References 门诊诊室(ID) ;
Create Index 临床出诊变动记录_IX_现诊室ID on 临床出诊变动记录(现诊室ID) Tablespace zl9Indexhis;


Create Table 临床出诊变动明细(
   变动ID number(18),
   变动性质 number(2),
   类型 number(2),
   名称 varchar2(50),
   序号 number(18),
   控制方式 number(2),
   数量 number(10),
   诊室ID number(18),
   门诊诊室 varchar2(20))
TABLESPACE zl9BaseItem ;

Alter Table 临床出诊变动明细  Add Constraint 临床出诊变动明细_PK  Primary Key (变动ID,名称,变动性质,序号) Using Index Tablespace zl9Indexhis;
Alter Table 临床出诊变动明细 Add Constraint 临床出诊变动明细_FK_变动ID Foreign Key (变动ID) References 临床出诊变动记录( ID) ;

Alter Table 临床出诊变动明细 Add Constraint 临床出诊变动明细_FK_诊室id Foreign Key (诊室id) References 门诊诊室(ID) ;
Create Index 临床出诊变动明细_IX_诊室ID on 临床出诊变动明细(诊室ID) Tablespace zl9Indexhis;


ALTER TABLE 病人挂号记录 ADD (出诊记录ID number(18));
Alter Table 病人挂号记录 Add Constraint 病人挂号记录_FK_出诊记录ID Foreign Key (出诊记录ID) References 临床出诊记录( ID) ;
Create Index 病人挂号记录_出诊记录ID on 病人挂号记录(出诊记录ID) Tablespace zl9Indexhis;


--常用停诊原因数据
insert into 常用停诊原因(编码,名称,简码,缺省标志) values ('01','手术','SS',0);
insert into 常用停诊原因(编码,名称,简码,缺省标志) values ('02','会诊','HZ',0);
insert into 常用停诊原因(编码,名称,简码,缺省标志) values ('03','公休','GX',0);
insert into 常用停诊原因(编码,名称,简码,缺省标志) values ('04','病假','BJ',0);
insert into 常用停诊原因(编码,名称,简码,缺省标志) values ('05','事假','SJ',0);
insert into 常用停诊原因(编码,名称,简码,缺省标志) values ('06','其他','QT',0);


--门诊诊室适用科室升级
Insert Into 门诊诊室适用科室
  (诊室id, 科室id, 缺省标志)
  Select *
  From (Select Distinct q.Id, m.科室id, 0 As 缺省标志
         From (Select Distinct b.科室id, 门诊诊室
                From 挂号安排诊室 A, 挂号安排 B
                Where a.号表id = b.Id
                Union All
                Select Distinct c.科室id, 门诊诊室
                From 挂号计划诊室 A, 挂号安排计划 B, 挂号安排 C
                Where a.计划id = b.Id And b.安排id = c.Id) M, 门诊诊室 Q
         Where m.门诊诊室 = q.名称)；


--模板相关处理
--1114:临床出诊安排
Insert Into zlPrograms(序号,标题,说明,系统,部件) Values( 1114,'临床出诊安排','对本单位临床科室的出诊安排进行管理。',&n_System,'zl9RegEvent'); 
Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select &n_System,1114,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0 Union All 
    Select '基本',-Null,NULL,1 From Dual Union All 
    Select '时间段设置',1,'增加、删除、修改的挂号时间的操作权限。有该权限时，允许对挂号项目的各时间段进行定义',1 From Dual Union All 
    Select '节假日设置',2,'增加、删除、修改法定节假日的操作权限。有该权限时，允许对各法定节假日进行定义',1 From Dual Union All 
    Select '门诊诊室设置',3,'增加、删除、修改门诊诊室的操作权限。有该权限时，允许对各门诊诊室进行设置',1 From Dual Union All 
    Select '出诊号源设置',4,'增加、删除、修改、停用及启用出诊号源的操作权限。有该权限时，允许针对各号源进行增加，修改，删除,停用及启用',1 From Dual Union All 
    Select '模板管理',5,'增加、删除、修改出诊模板的操作权限。有该权限时，允许针对模板进行增加，修改及删除。',1 From Dual Union All 
    Select '出诊安排',6,'针对各号源的出诊进行安排的操作权限，有该权限时，允许针对各号源的出诊进行安排。',1 From Dual Union All 
    Select '发布安排',7,'针对指定的安排表进行发布操作，有该权限时，允许针对出诊表进行发布操作。',1 From Dual Union All 
    Select '取消发布',8,'针对已经发布的安排表进行取消操作，有该权限时，允许针对出诊表进行取消发布操作。',1 From Dual Union All 
    Select '临时出诊安排',9,'对已经发布的安排但又未进行出诊安排的号源进行临时出诊安排操作，有该权限时，允许针对未出诊的号源进行临时出诊安排。',1 From Dual Union All 
    Select '停诊',10,'对已经发布的安排进行停诊操作，有该权限时，允许针对号源进行停诊操作。',1 From Dual Union All 
    Select '替诊',11,'对已经发布的安排进行替诊操作，有该权限时，允许针对号源进行替诊操作。',1 From Dual Union All 
    Select '加号',12,'对已经发布的安排进行加号操作，有该权限时，允许针对号源进行加号操作。',1 From Dual Union All 
    Select '减号',13,'对已经发布的安排进行减号操作，有该权限时，允许针对号源进行减号操作。',1 From Dual Union All 
    Select '调整分诊诊室',14,'对已经发布的安排进行诊室调整操作，有该权限时，允许针对号源进行诊室调整操作。',1 From Dual Union All 
    Select '调整预约挂号',15,'对已经发布的安排进行合作单位和预约方式的预约控制调整操作，有该权限时，允许针对号源进行合作单位和预约方式的预约控制调整。',1 From Dual Union All 
    Select '停诊申请',16,'对医生较长时间停诊的申请操作，有此权限时，允许进行申请操作。',1 From Dual Union All 
    Select '停诊审批',17,'针对停诊申请进行审批操作，有此权限时，允许对申请进行审批操作。',1 From Dual Union All 
    Select '允许代他人停诊申请',18,'允许代他人停诊申请操作,有此权限时，允许代其他人填写申请单的操作。',1 From Dual Union All 
    Select '所有科室',19,'当不具有该权限时，只能查看和处理门诊部下放给本科相关的号源。',1 From Dual Union All 
    Select '参数设置',20,'对临床出诊安排的参数设置进行操作的权限。有该权限时，允许进行本地参数设置',1 From Dual Union All 
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0) A;


--1115:患者服务中心
Insert Into zlPrograms(序号,标题,说明,系统,部件) Values( 1115,'患者服务中心','因停诊、替诊或复诊原估，客服人员通过电话通知病人以及对病人的预约信息进行退号、换诊及替诊等操作',&n_System,'zl9RegEvent'); 
Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select &n_System,1115,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0 Union All 
    Select '基本',-Null,NULL,1 From Dual Union All 
    Select '停诊信息处理',1,'实现因发生停诊或替诊操作时，需要对对应的预约进行取消、换诊或替诊的操作权限。有该权限时，允许操作员对停诊或替诊所对应的预约进行取消，换诊或替诊操作。',1 From Dual Union All 
    Select '预约登记信息处理',2,'实现对复诊病人到期的预约登记信息进行提醒复诊或给病人预约挂号的操作权限。有该权限时，允许操作员对到期的预约登记的信息通知患者或给病人预约挂号的操作权限。',1 From Dual Union All 
    Select '预约挂号登记',3,'实现预约挂号的操作权限。有该权限时，允许操作员进行预约挂号操作。',1 From Dual Union All 
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0) A;


Insert Into zlMenus(组别,ID,上级ID,标题,快键,说明,系统,模块,短标题,图标) Select A.组别,ZlMenus_ID.Nextval,A.ID,B.* 
From (
	Select 组别,ID From zlMenus Where 标题 = '门急诊挂号系统' And 组别 = '缺省' And 系统 = &n_System And 模块 Is Null) A,
	(	Select 标题,快键,说明,系统,模块,短标题,图标 From zlMenus Where 1 = 0 Union All
		Select '临床出诊安排' ,'A' ,'对本单位临床科室的出诊安排进行管理。' ,&n_System,1114,'出诊安排' ,236 From Dual Union All
		Select '患者服务中心' ,'B' ,'实现预约挂号登记,取消预约，换诊，替诊的操作权限。有该权限时，允许操作员进行预约登记,取消预约，换诊或替诊操作。' ,&n_System,1115,'服务中心' ,220 From Dual Union All
		Select 标题,快键,说明,系统,模块,短标题,图标 From zlMenus Where 1 = 0
          ) B;

--参数相关处理
Insert Into zlParameters(ID,系统,模块,私有,本机,授权,固定,部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明)
Select zlParameters_ID.Nextval,&n_System,-Null,-Null,-Null,-Null,-Null,A.* From (
Select 部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明 From zlParameters Where 1 = 0 Union All 
     Select 0,0,256,'挂号排班模式','0','0','1. 影响挂号（窗口，三方平台，自助等）的取数规则:' || chr(10) || 'a.如果设置为按”计划排班模式”,将根据”安排+计划”的方式进行排班，即在挂号业务读取有效号源时是从“挂号安排”等表中取数' || chr(10) || 'b.如果设置为按”出诊表排班模式”, 将根据”出诊表”（固定出诊表，月出诊表，周出诊表）的方式进行排班，即在挂号业务读取有效号源时是从“临诊出诊记录”等表中取数。' || chr(10) || '2.影响挂号窗口的展现形式:' || chr(10) || '    a.如果设置为按”计划排班模式”,挂号窗口左边的挂号安排数据将以周一至周日的方式展现.' || chr(10) || '    b.如果设置”出诊表排班模式”，挂号窗口左边只展现当天或指定日期的挂号安排数据。','0-计划排班模式,1-出诊表排班模式',Null,'1.如果医院业务较简单（一般在三级以下医院）,临床科室出诊相对固定，只有较少的出诊变化时，使用”计划排班模式”' || chr(10) || '2.如果医院业务较复杂（一般在三级医院），临床科室出诊临时变化较大时或按月或按周排班时，使用”出诊表排班模式”', '在医院启用HIS系统后，一般不调整此参数，如果调整此参数，将会直接影响到挂号业务(窗口，三方平台，自助系统等）。' 
     From Dual Union All 
Select 部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明 From zlParameters Where 1 = 0) A;

--参数脚本
Insert Into zlParameters(ID, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
Select Zlparameters_Id.Nextval,&n_System,1114,A.* From (
  Select 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明 From zlParameters Where 1 = 0 Union All 
  Select 1, -null, -null, 1, -null, -null, 1, '显示停用号源', '0', '0', '查看临床出诊号源时是否显示已停用号源。', '0-不显示，1-显示。', Null, '适用于用户在查看号源时停用号源的个性化显示方式', Null From Dual Union All
  Select -null, -null, 1, 1, -null, -null, 2, '只允许选院内医生', '0', '0', '进行号源设置时，选择的医生是否只能选择院内的医生。', '1-仅院内医生;0-允许选择外援医院或院内医生。', Null, '适用于没有外援医生的用户。', Null From Dual Union All
  Select -null, -null, 0, 0, -null, -null, 3, '预约清单控制方式', '0', '0', '停诊时是否将预约清单输出到Excel中。', '0-不输出到Excel，1-自动输出到Excel,2-选择输出到Excel。', Null, '适用于在停诊挂号安排时需要选择性的输出预约清单的业务', Null From Dual Union All
  Select  -null, -null, 0, 0, -null, -null, 4, '预约清单打印方式', '0', '0', '0-不打印,1-自动打印,2-选择是否打印。', '0-不打印,1-自动打印,2-选择是否打印。', Null, '适用于在停诊挂号安排时需要选择性的打印预约清单的业务', Null From Dual Union All
  Select  -null, -null, 0, 0, -null, -null, 5, '出诊表打印方式', '0', '0', '0-不打印,1-自动打印,2-选择是否打印。', '0-不打印,1-自动打印,2-选择是否打印。', Null, '适用于在发布出诊表时需要选择性的打印出诊表的业务', Null From Dual Union All
  Select  -null, -null, 0, 1, -null, -null, 6, '按替诊医生同步更新预约挂号单', '0', '0', '在替诊时，将根据替诊医生同步更新所有涉及替诊的预约挂号单。', '1-替诊同步更新,0-不更新。', Null, '适用于在替诊某个出诊号源的医生业务时，选择是否同步跟新预约挂号单的医生信息', Null From Dual Union All
  Select 1, -null, -null, 1, -null, -null, 7, '显示缺省控制信息', '', '1', '在号源管理中控制在选择号源时，是否在下方显示控制的相关信息，比如：缺省的序号信息、诊室信息、三方预约控制信息等，。', '0-不显示，1-显示。', Null, '适用于用户在查看号源时需要方便查看已设置的出诊上班时段安排信息的业务', Null From Dual Union All
  Select 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明 From zlParameters Where 1 = 0) A;
  
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
  Select Zlparameters_Id.Nextval, &n_System, 1111, 0, 0, 0, 0, 68, '病人同科限挂N个号', '0',
         '0', '控制一个病人同一天在一科室中,只能挂N个号。', '0-不限制;N-限制数量', Null, Null, Null
  From Dual
  Where Not Exists
   (Select 1
         From zlParameters
         Where 参数名 = '病人同科限挂N个号' And Nvl(模块, 0) = 1111 And Nvl(系统, 0) = &n_System);

Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
  Select Zlparameters_Id.Nextval, &n_System, 1111, 0, 0, 0, 0, 69, '病人挂号科室限制', '0',
         '0', '同一病人在同一时间能否挂多个科室。', '0-不限制,>=1表示科室限制的数量', Null, Null, Null
  From Dual
  Where Not Exists
   (Select 1
         From zlParameters
         Where 参数名 = '病人挂号科室限制' And Nvl(模块, 0) = 1111 And Nvl(系统, 0) = &n_System);	

Update zlParameters
Set 参数名 = '病人同科限约N个号', 影响控制说明 = '在预约时,同一病人在同一时间及同一科室的数量限制', 参数值含义= '0-不限制;N-限制预约数量'
Where 参数名 = '病人同科限约一个号' And 模块 = 1111 And 系统 = &n_System And Not Exists
 (Select 1
       From zlParameters
       Where 参数名 = '病人同科限约N个号' And Nvl(模块, 0) = 1111 And Nvl(系统, 0) = &n_System);

Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
  Select Zlparameters_Id.Nextval, &n_System, 1802, 0, 0, 0, 0, 0, 0, 39, '挂号时选择时间', '1',
         '1', '自助挂号挂专家号分时段的号别时，是否提供时段选择界面让用户选择挂号时间', '0-不启用；1-启用。', Null, Null, Null
  From Dual
  Where Not Exists
   (Select 1
         From zlParameters
         Where 参数名 = '挂号时选择时间' And Nvl(模块, 0) = 1802 And Nvl(系统, 0) = &n_System);
         
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
  Select Zlparameters_Id.Nextval, &n_System, 1803, 0, 0, 0, 0, 0, 0, 39, '预约时选择时间', '1',
         '1', '自助预约挂分时段的号别时，是否提供时段选择界面让用户选择预约时间', '0-不启用；1-启用。', Null, Null, Null
  From Dual
  Where Not Exists
   (Select 1
         From zlParameters
         Where 参数名 = '预约时选择时间' And Nvl(模块, 0) = 1803 And Nvl(系统, 0) = &n_System);


Insert Into zlProgFuncs
  (系统, 序号, 功能, 排列, 说明, 缺省值)
  Select &n_System, 9000, '预约登记', 7, '对门诊医生工作或住院医生工作等的预约登记的操作权限。有该权限时，允许进行预约登记操作。', 1
  From Dual
  Where Not Exists (Select 1 From zlProgFuncs Where 系统 = &n_System And 序号 = 9000 And 功能 = '预约登记');


--权限脚本
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1114,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All
Select '临床出诊表_ID','SELECT' From Dual Union All
Select '临床出诊安排_ID','SELECT' From Dual Union All
Select '临床出诊号源_ID','SELECT' From Dual Union All
Select '临床出诊号源限制_ID','SELECT' From Dual Union All
Select '临床出诊记录_ID','SELECT' From Dual Union All
Select '临床出诊变动记录_ID','SELECT' From Dual Union All
Select '临床出诊限制_ID','SELECT' From Dual Union All
Select '临床出诊号源限制','SELECT' From Dual Union All
Select '临床出诊号源时段','SELECT' From Dual Union All
Select '临床出诊号源控制','SELECT' From Dual Union All
Select '病人挂号记录','SELECT' From Dual Union All
Select '部门表','SELECT' From Dual Union All
Select '部门性质说明','SELECT' From Dual Union All
Select '常用停诊原因','SELECT' From Dual Union All
Select '法定假日表','SELECT' From Dual Union All
Select '挂号安排','SELECT' From Dual Union All
Select '挂号合作单位','SELECT' From Dual Union All
Select '号类','SELECT' From Dual Union All
Select '临床出诊安排','SELECT' From Dual Union All
Select '临床出诊变动记录','SELECT' From Dual Union All
Select '临床出诊变动明细','SELECT' From Dual Union All
Select '临床出诊表','SELECT' From Dual Union All
Select '临床出诊挂号控制','SELECT' From Dual Union All
Select '临床出诊挂号控制记录','SELECT' From Dual Union All
Select '临床出诊号源','SELECT' From Dual Union All
Select '临床出诊号源诊室','SELECT' From Dual Union All
Select '临床出诊记录','SELECT' From Dual Union All
Select '临床出诊时段','SELECT' From Dual Union All
Select '临床出诊停诊记录','SELECT' From Dual Union All
Select '临床出诊限制','SELECT' From Dual Union All
Select '临床出诊序号控制','SELECT' From Dual Union All
Select '临床出诊诊室','SELECT' From Dual Union All
Select '临床出诊诊室记录','SELECT' From Dual Union All
Select '门诊诊室','SELECT' From Dual Union All
Select '门诊诊室适用科室','SELECT' From Dual Union All
Select '上机人员表','SELECT' From Dual Union All
Select '时间段','SELECT' From Dual Union All
Select '收费项目目录','SELECT' From Dual Union All
Select '预约方式','SELECT' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,所有者,功能,对象,权限)
Select &n_System,1114,User,A.* From (
Select 功能,对象,权限 From zlProgPrivs Where 1 = 0 Union All
Select '时间段设置','Zl_上班时段_Modify','EXECUTE' From Dual Union All
Select '时间段设置','Zl_上班时段_Delete','EXECUTE' From Dual Union All
Select '节假日设置','Zl_法定假日表_Modify','EXECUTE' From Dual Union All
Select '节假日设置','Zl_法定假日表_Delete','EXECUTE' From Dual Union All
Select '门诊诊室设置','Zl_门诊诊室_Modify','EXECUTE' From Dual Union All
Select '门诊诊室设置','Zl_门诊诊室_Delete','EXECUTE' From Dual Union All
Select '出诊号源设置','Zl_临床出诊号源_Stopandstart','EXECUTE' From Dual Union All
Select '出诊号源设置','Zl_临床出诊号源_Modify','EXECUTE' From Dual Union All
Select '出诊号源设置','Zl_临床出诊号源_Delete','EXECUTE' From Dual Union All
Select '出诊号源设置','Zl_临床出诊号源限制_Modify','EXECUTE' From Dual Union All
Select '模板管理','Zl_临床出诊表_Add','EXECUTE' From Dual Union All
Select '模板管理','Zl_临床出诊表_Update','EXECUTE' From Dual Union All
Select '模板管理','Zl_临床出诊表_Delete','EXECUTE' From Dual Union All
Select '模板管理','Zl_临床出诊安排_Delete','EXECUTE' From Dual Union All
Select '模板管理','Zl_临床出诊上班时段_Delete','EXECUTE' From Dual Union All
Select '模板管理','Zl_临床出诊安排_Insert','EXECUTE' From Dual Union All
Select '模板管理','Zl_临床出诊限制_Insert','EXECUTE' From Dual Union All
Select '模板管理','Zl_临床出诊挂号控制_Insert','EXECUTE' From Dual Union All
Select '出诊安排','Zl_临床出诊表_Delete','EXECUTE' From Dual Union All
Select '出诊安排','Zl_临床出诊表_导入','EXECUTE' From Dual Union All
Select '出诊安排','Zl1_Auto_Buildingregisterplan','EXECUTE' From Dual Union All
Select '出诊安排','Zl_临床出诊表_Add','EXECUTE' From Dual Union All
Select '出诊安排','Zl_临床出诊表_Update','EXECUTE' From Dual Union All
Select '出诊安排','Zl_临床出诊安排_Delete','EXECUTE' From Dual Union All
Select '出诊安排','Zl_临床出诊上班时段_Delete','EXECUTE' From Dual Union All
Select '出诊安排','Zl_临床出诊安排_Insert','EXECUTE' From Dual Union All
Select '出诊安排','Zl_临床出诊限制_Insert','EXECUTE' From Dual Union All
Select '出诊安排','Zl_临床出诊挂号控制_Insert','EXECUTE' From Dual Union All
Select '出诊安排','Zl_临床出诊记录_Insert','EXECUTE' From Dual Union All
Select '出诊安排','Zl_临床出诊挂号控制记录_Insert','EXECUTE' From Dual Union All
Select '出诊安排','Zl_临床出诊安排_Applyto','EXECUTE' From Dual Union All
Select '出诊安排','Zl_Buildregisterplanbyrecord','EXECUTE' From Dual Union All
Select '出诊安排','Zl_临床出诊安排_BatchDelete','EXECUTE' From Dual Union All
Select '出诊安排','Zl_Buildregisterfixedrule','EXECUTE' From Dual Union All
Select '出诊安排','Zl_Buildregisterplanbytemplet','EXECUTE' From Dual Union All
Select '出诊安排','Zl_临床出诊安排_序号控制','EXECUTE' From Dual Union All
Select '出诊安排','Zl_临床出诊记录_Batchlock','EXECUTE' From Dual Union All
Select '发布安排','Zl_临床出诊安排_Publish','EXECUTE' From Dual Union All
Select '取消发布','Zl_临床出诊安排_Publish','EXECUTE' From Dual Union All
Select '临时出诊安排','Zl_临床出诊安排_Delete','EXECUTE' From Dual Union All
Select '临时出诊安排','Zl_临床出诊安排_Insert','EXECUTE' From Dual Union All
Select '临时出诊安排','Zl_临床出诊上班时段_Delete','EXECUTE' From Dual Union All
Select '临时出诊安排','Zl_临床出诊记录_Insert','EXECUTE' From Dual Union All
Select '临时出诊安排','Zl_临床出诊挂号控制记录_Insert','EXECUTE' From Dual Union All
Select '停诊','Zl_临床出诊记录_Stopvisit','EXECUTE' From Dual Union All
Select '替诊','Zl_临床出诊记录_Replacedoctor','EXECUTE' From Dual Union All
Select '替诊','Zl1_Ex_Isdoctorsamelevel','EXECUTE' From Dual Union All
Select '加号','Zl_临床出诊序号控制变动','EXECUTE' From Dual Union All
Select '减号','Zl_临床出诊序号控制变动','EXECUTE' From Dual Union All
Select '调整分诊诊室','Zl_临床出诊诊室_Update','EXECUTE' From Dual Union All
Select '调整预约挂号','Zl_临床出诊预约控制变动','EXECUTE' From Dual Union All
Select '调整预约挂号','Zl_临床出诊序号控制_Update','EXECUTE' From Dual Union All
Select '停诊申请','Zl_临床出诊停诊_Apply','EXECUTE' From Dual Union All
Select '停诊审批','Zl_临床出诊停诊_Audit','EXECUTE' From Dual Union All
Select 功能,对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,所有者,功能,对象,权限)
Select &n_System,1115,User,A.* From (
Select 功能,对象,权限 From zlProgPrivs Where 1 = 0 Union All
Select '停诊信息处理','Zl_患者服务中心_换诊','EXECUTE' From Dual Union All
Select '停诊信息处理','Zl_患者服务中心_替诊','EXECUTE' From Dual Union All
Select '停诊信息处理','Zl_患者服务中心_更新','EXECUTE' From Dual Union All
Select '停诊信息处理','zl_病人挂号记录_出诊_DELETE','EXECUTE' From Dual Union All
Select '预约登记信息处理','Zl_患者服务中心_换诊','EXECUTE' From Dual Union All
Select '预约登记信息处理','Zl_患者服务中心_替诊','EXECUTE' From Dual Union All
Select '预约登记信息处理','Zl_患者服务中心_更新','EXECUTE' From Dual Union All
Select '预约登记信息处理','zl_病人挂号记录_出诊_DELETE','EXECUTE' From Dual Union All
Select '基本','Zl1_Fun_Getreturnvisit','EXECUTE' From Dual Union All
Select '基本','Zl_患者服务中心_更新','EXECUTE' From Dual Union All
Select '基本','门诊费用记录','SELECT' From Dual Union All
Select '基本','病人预交记录','SELECT' From Dual Union All
Select '基本','结算方式','SELECT' From Dual Union All
Select '基本','病人挂号记录','SELECT' From Dual Union All
Select '基本','医疗付款方式','SELECT' From Dual Union All
Select '基本','病人服务信息记录','SELECT' From Dual Union All
Select '基本','病人信息','SELECT' From Dual Union All
Select '基本','部门表','SELECT' From Dual Union All
Select '基本','收费项目目录','SELECT' From Dual Union All
Select '基本','临床出诊记录','SELECT' From Dual Union All
Select '基本','临床出诊号源','SELECT' From Dual Union All
Select '基本','收费价目','SELECT' From Dual Union All
Select '基本','收入项目','SELECT' From Dual Union All
Select '基本','收费从属项目','SELECT' From Dual Union All
Select '基本','临床出诊序号控制','SELECT' From Dual Union All
Select '基本','收费特定项目','SELECT' From Dual Union All
Select '基本','就诊登记记录','SELECT' From Dual Union All
Select '基本','人员表','SELECT' From Dual Union All
Select '基本','就诊变动记录','SELECT' From Dual Union All
Select 功能,对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,所有者,功能,对象,权限)
Select &n_System,1111,User,A.* From (
Select 功能,对象,权限 From zlProgPrivs Where 1 = 0 Union All
Select '取消预约','Zl_病人挂号记录_出诊_Delete','EXECUTE' From Dual Union All
Select '退号','Zl_病人挂号记录_出诊_Delete','EXECUTE' From Dual Union All
Select '挂号','Zl_病人挂号记录_出诊_Insert','EXECUTE' From Dual Union All
Select '预约挂号','Zl_病人挂号记录_出诊_Insert','EXECUTE' From Dual Union All
Select '基本','Zl_挂号序号状态_出诊_Delete','EXECUTE' From Dual Union All
Select 功能,对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 9000, '预约登记', User, 'Zl_病人预约登记_Insert', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 9000 And 功能 = '预约登记' And Upper(对象) = Upper('Zl_病人预约登记_Insert'));

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 9000, '预约登记', User, 'ZL_患者服务中心_更新', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 9000 And 功能 = '预约登记' And Upper(对象) = Upper('ZL_患者服务中心_更新'));

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 9000, '基本', User, 'Zl_挂号序号状态_出诊_Delete', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 9000 And 功能 = '基本' And Upper(对象) = Upper('Zl_挂号序号状态_出诊_Delete'));

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1111,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All 
Select '临床出诊记录','SELECT' From Dual Union All
Select '临床出诊号源','SELECT' From Dual Union All
Select '临床出诊挂号控制记录','SELECT' From Dual Union All
Select '临床出诊诊室记录','SELECT' From Dual Union All
Select '临床出诊序号控制','SELECT' From Dual Union All
Select '临床出诊安排','SELECT' From Dual Union All
Select '临床出诊表','SELECT' From Dual Union All
Select 'Zl1_Auto_Buildingregisterplan','EXECUTE' From Dual Union All
Select 'Zl_预约方式_Check','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1113,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All 
Select '临床出诊记录','SELECT' From Dual Union All
Select '临床出诊号源','SELECT' From Dual Union All
Select '临床出诊挂号控制记录','SELECT' From Dual Union All
Select '临床出诊诊室记录','SELECT' From Dual Union All
Select '临床出诊序号控制','SELECT' From Dual Union All
Select '临床出诊安排','SELECT' From Dual Union All
Select '临床出诊表','SELECT' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,9000,'挂号',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All 
Select '临床出诊记录','SELECT' From Dual Union All
Select '临床出诊号源','SELECT' From Dual Union All
Select '临床出诊挂号控制记录','SELECT' From Dual Union All
Select '临床出诊诊室记录','SELECT' From Dual Union All
Select '临床出诊序号控制','SELECT' From Dual Union All
Select '临床出诊安排','SELECT' From Dual Union All
Select '临床出诊表','SELECT' From Dual Union All
Select 'Zl1_Auto_Buildingregisterplan','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,9000,'预约',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All 
Select '临床出诊记录','SELECT' From Dual Union All
Select '临床出诊号源','SELECT' From Dual Union All
Select '临床出诊挂号控制记录','SELECT' From Dual Union All
Select '临床出诊诊室记录','SELECT' From Dual Union All
Select '临床出诊序号控制','SELECT' From Dual Union All
Select '临床出诊安排','SELECT' From Dual Union All
Select '临床出诊表','SELECT' From Dual Union All
Select 'Zl1_Auto_Buildingregisterplan','EXECUTE' From Dual Union All
Select 'Zl_预约方式_Check','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1539,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All 
Select '临床出诊记录','SELECT' From Dual Union All
Select '临床出诊号源','SELECT' From Dual Union All
Select '号类','SELECT' From Dual Union All
Select '临床出诊诊室记录','SELECT' From Dual Union All
Select '临床出诊序号控制','SELECT' From Dual Union All
Select '收费项目类别','SELECT' From Dual Union All
Select '收费项目目录','SELECT' From Dual Union All
Select 'Zl1_Auto_Buildingregisterplan','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1802,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All 
Select '临床出诊记录','SELECT' From Dual Union All
Select '临床出诊号源','SELECT' From Dual Union All
Select '临床出诊诊室记录','SELECT' From Dual Union All
Select '临床出诊序号控制','SELECT' From Dual Union All
Select 'Zl1_Auto_Buildingregisterplan','EXECUTE' From Dual Union All
Select 'Zl_病人挂号记录_出诊_Insert','EXECUTE' From Dual Union All
Select 'Zl_预约挂号接收_出诊_Insert','EXECUTE' From Dual Union All
Select 'NextReservationNum','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1803,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All 
Select '临床出诊记录','SELECT' From Dual Union All
Select '临床出诊号源','SELECT' From Dual Union All
Select '临床出诊诊室记录','SELECT' From Dual Union All
Select '临床出诊序号控制','SELECT' From Dual Union All
Select 'Zl1_Auto_Buildingregisterplan','EXECUTE' From Dual Union All
Select 'Zl_病人挂号记录_出诊_Insert','EXECUTE' From Dual Union All
Select 'Zl_预约挂号接收_出诊_Insert','EXECUTE' From Dual Union All
Select 'NextReservationNum','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into Zlprocedure
  (ID, 类型, 名称, 状态, 所有者, 说明)
  Select Zlprocedure_Id.Nextval, 2, 'Zl1_Ex_Isdoctorsamelevel', 3, User, '比较两个医生的职务大小' From Dual;
  
Insert Into zlAutoJobs
  (系统, 类型, 序号, 名称, 说明, 内容, 参数, 执行时间, 间隔时间)
  Select &n_System, 1, 11, '出诊记录自动生成', '按固定时间完成对挂号固定安排出诊记录自动生成。', 'Zl1_Auto_BuildingRegisterPlan', Null,
         Trunc(Sysdate) + 1 / 24, 1
  From Dual
  Where Not Exists (Select 1 From zlAutoJobs Where 系统 = &n_System And 类型 = 1 And 序号 = 11);

Insert Into zlBaseCode
  (系统, 表名, 固定, 说明, 分类)
Values
  (&n_System, '常用停诊原因', 0, '临床出诊安排的常用停诊原因。', '医疗工作');
  
--数据处理
Insert Into zlProgFuncs
  (系统, 序号, 功能, 排列, 说明, 缺省值)
  Select 系统, 序号, '挂收费号', 排列, '给病人挂收费号的操作权限。有该权限时，允许对病人进行挂费用不为0的号', 缺省值
  From zlProgFuncs
  Where 序号 = 1111 And 系统 = &n_System And 功能 = '挂号';

Update Zlprogrelas Set 功能 = '挂收费号' Where 序号 = 1111 And 系统 = &n_System And 功能 = '挂号';

Update zlProgPrivs Set 功能 = '挂收费号' Where 序号 = 1111 And 系统 = &n_System And 功能 = '挂号';

Insert Into zlProgPrivs
  (系统, 序号, 功能, 对象, 所有者, 权限)
  Select 系统, 序号, '挂免费号', 对象, 所有者, 权限
  From zlProgPrivs
  Where 序号 = 1111 And 系统 = &n_System And 功能 = '挂收费号';

Update zlRoleGrant Set 功能 = '挂收费号' Where 序号 = 1111 And 系统 = &n_System And 功能 = '挂号';

Delete From zlProgFuncs Where 序号 = 1111 And 系统 = &n_System And 功能 = '挂号';


  --过程脚本
Create Or Replace Procedure Zl_上班时段_Modify
(
  操作类型_In     Number,
  站点_In         时间段.站点%Type,
  号类_In         时间段.号类%Type,
  时间段_In       时间段.时间段%Type,
  开始时间_In     时间段.开始时间%Type,
  终止时间_In     时间段.终止时间%Type,
  休息时段_In     时间段.休息时段%Type,
  缺省时间_In     时间段.缺省时间%Type,
  提前时间_In     时间段.提前时间%Type,
  出诊预留时间_In 时间段.出诊预留时间%Type,
  原站点_In       时间段.站点%Type := Null,
  原号类_In       时间段.号类%Type := Null,
  原时间段_In     时间段.时间段%Type := Null
) As
  --新增、修改上班时段
  --操作类型_In 0-新增，1-修改
  --原站点_In、原号类_In、原时间段_In 修改时传入
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_Count Number;
Begin
  If Nvl(操作类型_In, 0) = 0 Then
    --新增上班时段
    Begin
      Select 1
      Into n_Count
      From 时间段
      Where Nvl(站点, '-') = Nvl(站点_In, '-') And Nvl(号类, '-') = Nvl(号类_In, '-') And 时间段 = 时间段_In;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count > 0 Then
      v_Err_Msg := '当前站点已存在相同号类的上班时间段“' || 时间段_In || '”！';
      Raise Err_Item;
    End If;
  
    Insert Into 时间段
      (站点, 号类, 时间段, 开始时间, 终止时间, 休息时段, 缺省时间, 提前时间, 出诊预留时间)
    Values
      (站点_In, 号类_In, 时间段_In, 开始时间_In, 终止时间_In, 休息时段_In, Nvl(缺省时间_In, 开始时间_In), Nvl(提前时间_In, 开始时间_In), 出诊预留时间_In);
    Return;
  End If;

  --修改时，检查原上班时段是否被使用，被使用的不能修改站点、号类、时间段
  --不能删除被使用的范围最广的那一个,被使用的时段只要有一个即可（不同站点，不同号类可能会有多个同名的时间段）

  If Nvl(原站点_In, '-') <> Nvl(站点_In, '-') Or Nvl(原号类_In, '-') <> Nvl(号类_In, '-') Or 原时间段_In <> 时间段_In Then
    --临床出诊号源限制
    Begin
      Select 1
      Into n_Count
      From (Select b.上班时段, c.站点, a.号类, Row_Number() Over(Partition By b.上班时段 Order By b.上班时段, c.站点 Desc, a.号类 Desc) As 组号
             From 临床出诊号源 A, 临床出诊号源限制 B, 部门表 C
             Where a.Id = b.号源id And a.科室id = c.Id)
      Where 组号 = 1 And Nvl(站点, '-') = Nvl(原站点_In, '-') And Nvl(号类, '-') = Nvl(原号类_In, '-') And 上班时段 = 原时间段_In And
            Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    --临床出诊限制(固定规则、模板)
    If Nvl(n_Count, 0) = 0 Then
      Begin
        Select 1
        Into n_Count
        From (Select a.上班时段, c.站点, b.号类,
                      Row_Number() Over(Partition By a.上班时段 Order By a.上班时段, c.站点 Desc, b.号类 Desc) As 组号
               From 临床出诊限制 A, 临床出诊安排 D, 临床出诊号源 B, 部门表 C
               Where a.安排id = d.Id And d.号源id = b.Id And b.科室id = c.Id)
        Where 组号 = 1 And Nvl(站点, '-') = Nvl(原站点_In, '-') And Nvl(号类, '-') = Nvl(原号类_In, '-') And 上班时段 = 原时间段_In And
              Rownum < 2;
      Exception
        When Others Then
          n_Count := 0;
      End;
    End If;
    --临床出诊记录
    --不检查，因为该表太大，其次上班时段的信息都保存在了这个表中，没有找到上班时段时可由这个表的数据来提取
    If n_Count > 0 Then
      v_Err_Msg := '上班时间段“' || 原时间段_In || '”已被使用，不能修改其站点、号类及时间段名称！';
      Raise Err_Item;
    End If;
  
    Begin
      Select 1
      Into n_Count
      From 时间段
      Where Nvl(站点, '-') = Nvl(站点_In, '-') And Nvl(号类, '-') = Nvl(号类_In, '-') And 时间段 = 时间段_In;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count > 0 Then
      v_Err_Msg := '当前站点已存在相同号类的上班时间段“' || 时间段_In || '”！';
      Raise Err_Item;
    End If;
  End If;

  Update 时间段
  Set 站点 = 站点_In, 号类 = 号类_In, 时间段 = 时间段_In, 开始时间 = 开始时间_In, 终止时间 = 终止时间_In, 休息时段 = 休息时段_In, 缺省时间 = Nvl(缺省时间_In, 开始时间_In),
      提前时间 = Nvl(提前时间_In, 开始时间_In), 出诊预留时间 = 出诊预留时间_In
  Where Nvl(站点, '-') = Nvl(原站点_In, '-') And Nvl(号类, '-') = Nvl(原号类_In, '-') And 时间段 = 原时间段_In;
  If Sql%NotFound Then
    Insert Into 时间段
      (站点, 号类, 时间段, 开始时间, 终止时间, 休息时段, 缺省时间, 提前时间, 出诊预留时间)
    Values
      (站点_In, 号类_In, 时间段_In, 开始时间_In, 终止时间_In, 休息时段_In, Nvl(缺省时间_In, 开始时间_In), Nvl(提前时间_In, 开始时间_In), 出诊预留时间_In);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_上班时段_Modify;
/
Create Or Replace Procedure Zl_上班时段_Delete
(
  站点_In   时间段.站点%Type,
  号类_In   时间段.号类%Type,
  时间段_In 时间段.时间段%Type
) As
  -- 删除上班时段
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_Count Number;
Begin
  --数据检查，上班时间段已被使用则不能删除
  --不能删除被使用的范围最广的那一个,被使用的时段只要有一个即可（不同站点，不同号类可能会有多个同名的时间段）

  --临床出诊号源限制
  Begin
    Select 1
    Into n_Count
    From (Select b.上班时段, c.站点, a.号类, Row_Number() Over(Partition By b.上班时段 Order By b.上班时段, c.站点 Desc, a.号类 Desc) As 组号
           From 临床出诊号源 A, 临床出诊号源限制 B, 部门表 C
           Where a.Id = b.号源id And a.科室id = c.Id)
    Where 组号 = 1 And Nvl(站点, '-') = Nvl(站点_In, '-') And Nvl(号类, '-') = Nvl(号类_In, '-') And 上班时段 = 时间段_In And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;
  --临床出诊限制(固定规则、模板)
  If Nvl(n_Count, 0) = 0 Then
    Begin
      Select 1
      Into n_Count
      From (Select a.上班时段, c.站点, b.号类, Row_Number() Over(Partition By a.上班时段 Order By a.上班时段, c.站点 Desc, b.号类 Desc) As 组号
             From 临床出诊限制 A, 临床出诊安排 D, 临床出诊号源 B, 部门表 C
             Where a.安排id = d.Id And d.号源id = b.Id And b.科室id = c.Id)
      Where 组号 = 1 And Nvl(站点, '-') = Nvl(站点_In, '-') And Nvl(号类, '-') = Nvl(号类_In, '-') And 上班时段 = 时间段_In And
            Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
  End If;
  --临床出诊记录
  --不检查，因为该表太大，其次上班时段的信息都保存在了这个表中，没有找到上班时段时可由这个表的数据来提取

  If n_Count > 0 Then
    v_Err_Msg := '当前上班时间段已被使用，不能删除！';
    Raise Err_Item;
  End If;

  Delete From 时间段
  Where Nvl(站点, '-') = Nvl(站点_In, '-') And Nvl(号类, '-') = Nvl(号类_In, '-') And 时间段 = 时间段_In;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_上班时段_Delete;
/

Create Or Replace Procedure Zl_法定假日表_Modify
(
  操作类型_In Number,
  年份_In     法定假日表.年份%Type,
  节日名称_In 法定假日表.节日名称%Type,
  开始日期_In 法定假日表.开始日期%Type,
  终止日期_In 法定假日表.终止日期%Type,
  备注_In     法定假日表.备注%Type,
  换休情况_In Varchar2 := Null,
  允许预约_In Varchar2 := Null,
  允许挂号_In Varchar2 := Null
) As
  --新增、修改法定节假日
  --      操作类型_In 0-新增，1-修改
  --      换休情况_In 格式：调休时间1~ 原上班时间1;调休时间2~ 原上班时间2;
  --      允许预约_in 允许预约的日期,格式：yyyy-mm-dd;yyyy-mm-dd;...
  --      允许挂号_in 允许挂号的日期,格式：yyyy-mm-dd;yyyy-mm-dd;...
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_Count Number;

  v_换休情况 Varchar2(4000);
  v_当前项目 Varchar2(4000);
  d_开始日期 Date;
  d_终止日期 Date;
Begin
  If 操作类型_In = 0 Then
    --新增
    Begin
      Select 1
      Into n_Count
      From 法定假日表
      Where 性质 = 0 And 年份 = 年份_In And 节日名称 = 节日名称_In And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count > 0 Then
      v_Err_Msg := 年份_In || '年已存在“' || 节日名称_In || '”！';
      Raise Err_Item;
    End If;
  
    Begin
      Select 1
      Into n_Count
      From 临床出诊记录
      Where 出诊日期 Between 开始日期_In And 终止日期_In And Nvl(是否发布, 0) = 1 And (Nvl(已约数, 0) <> 0 Or Nvl(已挂数, 0) <> 0) And
            Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count > 0 Then
      v_Err_Msg := '当前节假日的时间范围内已有预约挂号病人，不能设置！';
      Raise Err_Item;
    End If;
  
    Insert Into 法定假日表
      (年份, 节日名称, 性质, 开始日期, 终止日期, 备注, 允许预约, 允许挂号)
    Values
      (年份_In, 节日名称_In, 0, 开始日期_In, 终止日期_In, 备注_In, 允许预约_In, 允许挂号_In);
  
    If 换休情况_In Is Not Null Then
      v_换休情况 := 换休情况_In || ';';
    End If;
    While v_换休情况 Is Not Null Loop
      v_当前项目 := Substr(v_换休情况, 0, Instr(v_换休情况, ';') - 1);
      d_开始日期 := To_Date(Substr(v_当前项目, 0, Instr(v_当前项目, '~') - 1), 'yyyy-mm-dd');
      d_终止日期 := To_Date(Substr(v_当前项目, Instr(v_当前项目, '~') + 1), 'yyyy-mm-dd');
    
      Insert Into 法定假日表
        (年份, 节日名称, 性质, 开始日期, 终止日期, 备注)
      Values
        (年份_In, 节日名称_In, 1, d_开始日期, d_终止日期, Null);
    
      v_换休情况 := Substr(v_换休情况, Instr(v_换休情况, ';') + 1);
    End Loop;
  
  Elsif 操作类型_In = 1 Then
    --修改
    Begin
      Select 开始日期
      Into d_开始日期
      From 法定假日表
      Where 性质 = 0 And 年份 = 年份_In And 节日名称 = 节日名称_In And Rownum < 2;
    Exception
      When Others Then
        d_开始日期 := Null;
    End;
    If d_开始日期 Is Null Then
      v_Err_Msg := 年份_In || '年不存在“' || 节日名称_In || '”！';
      Raise Err_Item;
    End If;
  
    If Sysdate > d_开始日期 Then
      v_Err_Msg := '当前时间已经大于了节假日开始时间，不能修改！';
      Raise Err_Item;
    End If;
  
    Begin
      Select 1
      Into n_Count
      From 临床出诊记录
      Where 出诊日期 Between 开始日期_In And 终止日期_In And Nvl(是否发布, 0) = 1 And (Nvl(已约数, 0) <> 0 Or Nvl(已挂数, 0) <> 0) And
            Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count > 0 Then
      v_Err_Msg := '当前节假日的时间范围内已有预约挂号病人，不能修改！';
      Raise Err_Item;
    End If;
  
    Update 法定假日表
    Set 开始日期 = 开始日期_In, 终止日期 = 终止日期_In, 备注 = 备注_In, 允许预约 = 允许预约_In, 允许挂号 = 允许挂号_In
    Where 年份 = 年份_In And Nvl(性质, 0) = 0 And 节日名称 = 节日名称_In;
  
    --先删除换休数据
    Delete From 法定假日表 Where 年份 = 年份_In And Nvl(性质, 0) = 1 And 节日名称 = 节日名称_In;
    If 换休情况_In Is Not Null Then
      v_换休情况 := 换休情况_In || ';';
    End If;
    While v_换休情况 Is Not Null Loop
      v_当前项目 := Substr(v_换休情况, 0, Instr(v_换休情况, ';') - 1);
      d_开始日期 := To_Date(Substr(v_当前项目, 0, Instr(v_当前项目, '~') - 1), 'yyyy-mm-dd');
      d_终止日期 := To_Date(Substr(v_当前项目, Instr(v_当前项目, '~') + 1), 'yyyy-mm-dd');
    
      Insert Into 法定假日表
        (年份, 节日名称, 性质, 开始日期, 终止日期, 备注)
      Values
        (年份_In, 节日名称_In, 1, d_开始日期, d_终止日期, Null);
    
      v_换休情况 := Substr(v_换休情况, Instr(v_换休情况, ';') + 1);
    End Loop;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_法定假日表_Modify;
/
Create Or Replace Procedure Zl_法定假日表_Delete
(
  年份_In     法定假日表.年份%Type,
  节日名称_In 法定假日表.节日名称%Type
) As
  --删除法定节假日
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  d_开始日期 Date;
Begin
  Begin
    Select 开始日期
    Into d_开始日期
    From 法定假日表
    Where 性质 = 0 And 年份 = 年份_In And 节日名称 = 节日名称_In And Rownum < 2;
  Exception
    When Others Then
      d_开始日期 := Null;
  End;
  If d_开始日期 Is Null Then
    v_Err_Msg := 年份_In || '年不存在“' || 节日名称_In || '”！';
    Raise Err_Item;
  End If;

  If Sysdate > d_开始日期 Then
    v_Err_Msg := '当前时间已经大于了节假日开始时间，不能修改！';
    Raise Err_Item;
  End If;

  Delete From 法定假日表 Where 年份 = 年份_In And 节日名称 = 节日名称_In;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_法定假日表_Delete;
/
Create Or Replace Procedure Zl_门诊诊室_Modify
(
  操作类型_In Number,
  Id_In       门诊诊室.Id%Type,
  编码_In     门诊诊室.编码%Type := Null,
  名称_In     门诊诊室.名称%Type := Null,
  简码_In     门诊诊室.简码%Type := Null,
  位置_In     门诊诊室.位置%Type := Null,
  站点_In     门诊诊室.站点%Type := Null,
  适用科室_In Varchar2 := Null
) As
  --新增、修改门诊诊室
  --操作类型_In 0-新增，1-修改
  --适用科室_In 格式：科室ID，格式：科室1;科室2;科室3;...
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_Id    门诊诊室.Id%Type;
  n_Count Number;
Begin
  If 操作类型_In = 0 Then
    --新增
    Begin
      Select 1 Into n_Count From 门诊诊室 Where 名称 = 名称_In And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count > 0 Then
      v_Err_Msg := 名称_In || ' 已存在！';
      Raise Err_Item;
    End If;
  
    Select 门诊诊室_Id.Nextval Into n_Id From Dual;
    Insert Into 门诊诊室 (ID, 编码, 名称, 简码, 位置, 站点) Values (n_Id, 编码_In, 名称_In, 简码_In, 位置_In, 站点_In);
  
    --插入门诊诊室适用科室
    If Not 适用科室_In Is Null Then
      Insert Into 门诊诊室适用科室
        (诊室id, 科室id)
        Select n_Id, Column_Value As 科室id From Table(f_Num2list(适用科室_In, ';'));
    End If;
  
    Return;
  End If;

  --修改
  Begin
    Select 1 Into n_Count From 门诊诊室 Where 名称 = 名称_In And ID <> Id_In And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count > 0 Then
    v_Err_Msg := 名称_In || ' 已存在！';
    Raise Err_Item;
  End If;

  Update 门诊诊室 Set 编码 = 编码_In, 名称 = 名称_In, 简码 = 简码_In, 位置 = 位置_In, 站点 = 站点_In Where ID = Id_In;

  --先删除
  Delete From 门诊诊室适用科室 Where 诊室id = Id_In;
  --插入门诊诊室适用科室
  If Not 适用科室_In Is Null Then
    Insert Into 门诊诊室适用科室
      (诊室id, 科室id)
      Select Id_In, Column_Value As 科室id From Table(f_Num2list(适用科室_In, ';'));
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊诊室_Modify;
/
Create Or Replace Procedure Zl_门诊诊室_Delete(Id_In 门诊诊室.Id%Type) As
  --删除门诊诊室
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_Count Number(2);
Begin
  Begin
    Select 1 Into n_Count From 临床出诊号源诊室 Where 诊室id = Id_In And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;

  If Nvl(n_Count, 0) = 0 Then
    Begin
      Select 1 Into n_Count From 临床出诊诊室 Where 诊室id = Id_In And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
  End If;

  If Nvl(n_Count, 0) = 0 Then
    Begin
      Select 1 Into n_Count From 临床出诊诊室记录 Where 诊室id = Id_In And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
  End If;
  If Nvl(n_Count, 0) > 0 Then
    v_Err_Msg := '当前诊室已被使用，不能删除！';
    Raise Err_Item;
  End If;

  Delete From 门诊诊室适用科室 Where 诊室id = Id_In;
  Delete From 门诊诊室 Where ID = Id_In;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊诊室_Delete;
/
Create Or Replace Procedure Zl_临床出诊号源_Stopandstart
(
  Id_In   临床出诊号源.Id%Type,
  停用_In Number := 0
) As
Begin
  If Nvl(停用_In, 0) = 1 Then
    Update 临床出诊号源 Set 撤档时间 = Sysdate Where ID = Id_In;
  Else
    Update 临床出诊号源 Set 撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'), 是否删除 = 0 Where ID = Id_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊号源_Stopandstart;
/
Create Or Replace Procedure Zl_临床出诊号源_Modify
(
  操作类型_In     Number,
  Id_In           临床出诊号源.Id%Type,
  号类_In         临床出诊号源.号类%Type := Null,
  号码_In         临床出诊号源.号码%Type := Null,
  科室id_In       临床出诊号源.科室id%Type := 0,
  项目id_In       临床出诊号源.项目id%Type := 0,
  医生id_In       临床出诊号源.医生id%Type := Null,
  医生姓名_In     临床出诊号源.医生姓名%Type := Null,
  是否建病案_In   临床出诊号源.是否建病案%Type := 0,
  预约天数_In     临床出诊号源.预约天数%Type := 0,
  出诊频次_In     临床出诊号源.出诊频次%Type := 0,
  假日控制状态_In 临床出诊号源.假日控制状态%Type := 0,
  是否假日换休_In 临床出诊号源.是否假日换休%Type := 0,
  是否临床排班_In 临床出诊号源.是否临床排班%Type := 0,
  排班方式_In     临床出诊号源.排班方式%Type := 0
) As
  --操作类型_In 0-新增，1-修改，2-删除
  --分诊诊室_In 诊室ID，格式：诊室ID1;诊室ID2;诊室ID13;...
  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  n_号源id 临床出诊号源.Id%Type;
  n_Count  Number;
Begin

  If 操作类型_In = 0 Then
    --增加号源
    n_号源id := Id_In;
  
    If Nvl(n_号源id, 0) = 0 Then
      Select 临床出诊号源_Id.Nextval Into n_号源id From Dual;
    End If;
    Insert Into 临床出诊号源
      (ID, 号类, 号码, 科室id, 项目id, 医生id, 医生姓名, 是否建病案, 预约天数, 出诊频次, 假日控制状态, 是否假日换休, 是否临床排班, 排班方式, 是否删除, 建档时间, 撤档时间)
    Values
      (n_号源id, 号类_In, 号码_In, 科室id_In, 项目id_In, 医生id_In, 医生姓名_In, 是否建病案_In, 预约天数_In, 出诊频次_In, 假日控制状态_In, 是否假日换休_In,
       是否临床排班_In, 排班方式_In, 0, Sysdate, To_Date('3000-01-01', 'yyyy-mm-dd'));
  
    Return;
  End If;

  --修改号源
  Update 临床出诊号源
  Set 号类 = 号类_In, 号码 = 号码_In, 科室id = 科室id_In, 项目id = 项目id_In, 医生id = 医生id_In, 医生姓名 = 医生姓名_In, 是否建病案 = 是否建病案_In,
      预约天数 = 预约天数_In, 出诊频次 = 出诊频次_In, 假日控制状态 = 假日控制状态_In, 是否假日换休 = 是否假日换休_In, 是否临床排班 = 是否临床排班_In, 排班方式 = 排班方式_In
  Where ID = Id_In And Nvl(是否删除, 0) = 0 And Nvl(撤档时间, Sysdate) >= Sysdate;
  If Sql%NotFound Then
    v_Err_Msg := '当前号源可能已被他人删除或停用，不能对该号源信息进行修改!';
    Raise Err_Item;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊号源_Modify;
/

Create Or Replace Procedure Zl_临床出诊号源_Delete(Id_In 临床出诊号源.Id%Type) As
  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  n_Count  Number;
  l_限制id t_Numlist := t_Numlist();
Begin
  Select Count(1) Into n_Count From 临床出诊安排 Where 号源id = Id_In;

  If n_Count = 0 Then
  
    Select ID Bulk Collect Into l_限制id From 临床出诊号源限制 Where 号源id = Id_In;
  
    Forall I In 1 .. l_限制id.Count
      Delete 临床出诊号源时段 Where 限制id = l_限制id(I);
  
    Forall I In 1 .. l_限制id.Count
      Delete 临床出诊号源控制 Where 限制id = l_限制id(I);
  
    Forall I In 1 .. l_限制id.Count
      Delete 临床出诊号源诊室 Where 限制id = l_限制id(I);
  
    Delete 临床出诊号源限制 Where 号源id = Id_In;
    --假删除
  
    Delete From 临床出诊号源 Where ID = Id_In;
    If Sql%NotFound Then
      v_Err_Msg := '当前号源可能已被他人删除，不能再删除!';
      Raise Err_Item;
    End If;
    Return;
  End If;
  Update 临床出诊号源 Set 是否删除 = 1, 撤档时间 = Sysdate Where ID = Id_In And Nvl(是否删除, 0) = 0;
  If Sql%NotFound Then
    v_Err_Msg := '当前号源可能已被他人删除，不能再删除!';
    Raise Err_Item;
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊号源_Delete;
/

Create Or Replace Procedure Zl_临床出诊号源限制_Modify
(
  Id_In           临床出诊号源限制.Id%Type,
  号源id_In       临床出诊号源限制.号源id%Type,
  上班时段_In     临床出诊号源限制.上班时段%Type,
  限号数_In       临床出诊号源限制.限号数%Type,
  限约数_In       临床出诊号源限制.限约数%Type,
  是否序号控制_In 临床出诊号源限制.是否序号控制%Type,
  是否分时段_In   临床出诊号源限制.是否分时段%Type,
  预约控制_In     临床出诊号源限制.预约控制%Type,
  是否独占_In     临床出诊号源限制.是否独占%Type,
  分诊方式_In     临床出诊号源限制.分诊方式%Type,
  诊室id_In       临床出诊号源限制.诊室id%Type,
  号源诊室_In     Varchar2 := Null,
  号源时段_In     Varchar2 := Null,
  号源控制_In     Varchar2 := Null,
  删除号源限制_In Integer := 0
  
) As
  --号源时段_IN:序号,开始时间(HH:MM:SS),终止时(HH:MM:SS)间,数量,是否预约|...
  --号源诊室_IN:诊室id1,诊室id2,....
  --号源控制_IN:类型,性质,名称,控制方式,序号,数量|
  --删除号源限制_in:1-插入数据前，先删除号源限制,0-不删除数据，直接插入

  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  l_限制id   t_Numlist := t_Numlist();
  n_Count    Number;
  v_开始时间 Varchar2(20);
  v_终止时间 Varchar2(20);

  n_序号     临床出诊号源时段.序号%Type;
  d_开始时间 临床出诊号源时段.开始时间%Type;
  d_终止时间 临床出诊号源时段.终止时间%Type;
  n_数量     临床出诊号源时段.限制数量%Type;
  n_是否预约 临床出诊号源时段.是否预约%Type;
  n_类型     临床出诊号源控制.类型%Type;
  n_性质     临床出诊号源控制.性质%Type;
  v_名称     临床出诊号源控制.名称%Type;
  n_控制方式 临床出诊号源控制.控制方式%Type;
  n_限制数量 临床出诊号源控制.数量%Type;
Begin
  If Nvl(删除号源限制_In, 0) = 1 Then
    Select ID Bulk Collect Into l_限制id From 临床出诊号源限制 Where 号源id = 号源id_In;
    Forall I In 1 .. l_限制id.Count
      Delete 临床出诊号源时段 Where 限制id = l_限制id(I);
  
    Forall I In 1 .. l_限制id.Count
      Delete 临床出诊号源控制 Where 限制id = l_限制id(I);
  
    Forall I In 1 .. l_限制id.Count
      Delete 临床出诊号源诊室 Where 限制id = l_限制id(I);
  
    Delete 临床出诊号源限制 Where 号源id = 号源id_In;
    Delete From 临床出诊号源限制 Where 号源id = 号源id_In;
  
  End If;

  Select Count(1) Into n_Count From 临床出诊号源限制 Where ID = Id_In;
  If n_Count = 0 Then
    Insert Into 临床出诊号源限制
      (ID, 号源id, 上班时段, 限号数, 限约数, 是否序号控制, 是否分时段, 预约控制, 是否独占, 分诊方式, 诊室id)
    Values
      (Id_In, 号源id_In, 上班时段_In, 限号数_In, 限约数_In, 是否序号控制_In, 是否分时段_In, 预约控制_In, 是否独占_In, 分诊方式_In, 诊室id_In);
  
  End If;

  If 号源时段_In Is Not Null Then
    --插入号源缺省时间段
    For c_时间段集 In (Select Rownum As 序号, Column_Value As 值 From Table(f_Str2list(号源时段_In, '|'))) Loop
      n_序号     := Null;
      v_开始时间 := Null;
      v_终止时间 := Null;
      n_数量     := Null;
      n_是否预约 := Null;
      For c_时间段 In (Select Rownum As 序号, Column_Value As 值 From Table(f_Str2list(c_时间段集.值)) Order By 序号) Loop
        If c_时间段.序号 = 1 Then
          n_序号 := To_Number(c_时间段.值);
        End If;
      
        If c_时间段.序号 = 2 Then
          v_开始时间 := c_时间段.值;
        End If;
      
        If c_时间段.序号 = 3 Then
          v_终止时间 := c_时间段.值;
        End If;
      
        If c_时间段.序号 = 4 Then
          n_数量 := To_Number(c_时间段.值);
        End If;
      
        If c_时间段.序号 = 5 Then
          n_是否预约 := To_Number(c_时间段.值);
        End If;
      
      End Loop;
      d_开始时间 := To_Date('3000-01-01 ' || Nvl(v_开始时间, ''), 'yyyy-mm-dd hh24:mi:ss');
      d_终止时间 := To_Date('3000-01-01 ' || Nvl(v_终止时间, ''), 'yyyy-mm-dd hh24:mi:ss');
    
      If d_开始时间 >= d_终止时间 Then
        d_终止时间 := d_终止时间 + 1;
      End If;
    
      If Nvl(n_序号, 0) <> 0 Then
        Insert Into 临床出诊号源时段
          (限制id, 序号, 开始时间, 终止时间, 限制数量, 是否预约)
        Values
          (Id_In, n_序号, d_开始时间, d_终止时间, n_数量, n_是否预约);
      End If;
    End Loop;
  
  End If;

  --插入号源的缺省控制
  --号源控制_IN:类型,性质,名称,控制方式,序号,数量|
  If 号源控制_In Is Not Null Then
    For c_时间段集 In (Select Rownum As 序号, Column_Value As 值 From Table(f_Str2list(号源控制_In, '|'))) Loop
      n_类型     := Null;
      n_性质     := Null;
      v_名称     := Null;
      n_序号     := Null;
      n_控制方式 := Null;
      n_限制数量 := Null;
    
      --类型,性质,名称,控制方式,序号,数量|
      For c_时间段 In (Select Rownum As 序号, Column_Value As 值 From Table(f_Str2list(c_时间段集.值)) Order By 序号) Loop
        If c_时间段.序号 = 1 Then
          n_类型 := To_Number(c_时间段.值);
        End If;
      
        If c_时间段.序号 = 2 Then
          n_性质 := To_Number(c_时间段.值);
        End If;
      
        If c_时间段.序号 = 3 Then
          v_名称 := c_时间段.值;
        End If;
      
        If c_时间段.序号 = 4 Then
          n_控制方式 := To_Number(c_时间段.值);
        End If;
      
        If c_时间段.序号 = 5 Then
          n_序号 := To_Number(c_时间段.值);
        End If;
      
        If c_时间段.序号 = 6 Then
          n_限制数量 := To_Number(c_时间段.值);
        End If;
      
      End Loop;
    
      If v_名称 Is Not Null Then
        Insert Into 临床出诊号源控制
          (限制id, 类型, 性质, 名称, 序号, 控制方式, 数量)
        Values
          (Id_In, n_类型, n_性质, v_名称, n_序号, n_控制方式, n_限制数量);
      
      End If;
    End Loop;
  End If;
  --插入号源诊室
  If 号源诊室_In Is Not Null Then
    Insert Into 临床出诊号源诊室
      (限制id, 诊室id)
      Select Id_In As 限制id, Column_Value As 科室id From Table(f_Num2list(号源诊室_In));
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊号源限制_Modify;
/

Create Or Replace Procedure Zl_临床出诊表_导入(号码_In 挂号安排.号码%Type) As
  -------------------------------------------------------------------------
  --功能说明：导放临床出诊表,主要是根据挂号安排，挂号计划安排等表进行数据导入,规则如下:
  -------------------------------------------------------------------------

  Err_Item Exception;
  v_Err_Msg Varchar2(200);

  l_限制id t_Numlist := t_Numlist();
  n_Count  Number(18);

  v_时间段 Varchar2(4000);

  Procedure Zl_Register_Import(号码_In 挂号安排.号码%Type) As
    n_号源id   临床出诊号源.Id%Type;
    d_开始时间 临床出诊号源.建档时间%Type;
    n_出诊id   临床出诊表.Id%Type;
    n_安排id   临床出诊安排.Id%Type;
    d_终止时间 临床出诊号源.建档时间%Type;
    n_诊室id   门诊诊室.Id%Type;
    n_预约控制 临床出诊限制.预约控制%Type;
  
    l_限制id t_Numlist := t_Numlist();
    n_限制id 临床出诊号源限制.Id%Type;
  Begin
    --科室、项目、医生相同的只导入一个
    For c_号源 In (Select ID, 号类, 号码, 科室id, 项目id, 医生姓名, Decode(医生id, 0, Null, 医生id) As 医生id, 序号, 周日, 周一, 周二, 周三, 周四, 周五, 周六,
                        病案必须, 分诊方式, 序号控制, 开始时间, 终止时间, 停用日期, 执行时间, 执行计划id, 是否删除, 默认时段间隔, 预约天数
                 From 挂号安排 A
                 Where 号码 = 号码_In And Not Exists (Select 1
                        From 临床出诊号源
                        Where 科室id = a.科室id And 项目id = a.项目id And 医生姓名 = a.医生姓名 And
                              医生id = Decode(a.医生id, 0, Null, a.医生id))) Loop
    
      Select 临床出诊号源_Id.Nextval Into n_号源id From Dual;
    
      Select Nvl(Min(开始时间), Sysdate)
      Into d_开始时间
      From (Select Min(开始时间) As 开始时间
             From 挂号安排时段
             Where 安排id = c_号源.Id
             Union All
             Select Min(a.生效时间) As 开始时间 From 挂号安排计划 A Where a.安排id = c_号源.Id);
    
      --1.处理临床出诊号源
      Insert Into 临床出诊号源
        (ID, 号类, 号码, 科室id, 项目id, 医生id, 医生姓名, 是否建病案, 预约天数, 出诊频次, 假日控制状态, 是否临床排班, 排班方式, 是否删除, 建档时间, 撤档时间)
      Values
        (n_号源id, c_号源.号类, c_号源.号码, c_号源.科室id, c_号源.项目id, c_号源.医生id, c_号源.医生姓名, c_号源.病案必须, c_号源.预约天数, c_号源.默认时段间隔, 2, 0,
         0, c_号源.是否删除, d_开始时间, Nvl(c_号源.停用日期, To_Date('3000-01-01', 'yyyy-mm-dd')));
    
      --2.处理临床出诊停诊记录
      Insert Into 临床出诊停诊记录
        (ID, 记录id, 开始时间, 终止时间, 停诊原因, 替诊医生id, 替诊医生姓名, 申请人, 申请时间, 审批人, 审批时间, 取消人, 取消时间)
        Select 临床出诊停诊记录_Id.Nextval As ID, Null As 记录id, a.开始停止时间, a.结束停止时间, a.备注, Null As 替诊医生id, Null As 替诊医生姓名, b.医生姓名,
               a.制订日期, a.制订人, a.制订日期, Null As 取消人, Null As 取消时间
        From 挂号安排停用状态 A, 挂号安排 B
        Where a.安排id = b.Id And b.Id = c_号源.Id And b.医生id Is Not Null And Not Exists
         (Select 1
               From 临床出诊停诊记录
               Where 记录id Is Null And 申请人 = b.医生姓名 And 开始时间 = a.开始停止时间 And 终止时间 = a.结束停止时间);
    
      --3.处理相关的出诊表数据
      --3.1 固定出诊表
      --    导入时的年+固定出诊表,比如：2015年固定出诊表
      Begin
        Select ID Into n_出诊id From 临床出诊表 Where 备注 = '系统导入' And 发布人 Is Null;
      Exception
        When Others Then
          n_出诊id := Null;
      End;
      If Nvl(n_出诊id, 0) = 0 Then
        Select 临床出诊表_Id.Nextval Into n_出诊id From Dual;
        Insert Into 临床出诊表
          (ID, 排班方式, 出诊表名, 年份, 月份, 周数, 应用范围, 科室id, 备注, 发布人, 发布时间)
        Values
          (n_出诊id, 0, To_Char(Sysdate, 'yyyy') || '年固定出诊表', To_Number(To_Char(Sysdate, 'yyyy')), Null, Null, Null, Null,
           '系统导入', Null, Null);
      End If;
    
      --3.2导入临床出诊安排
      d_开始时间 := Sysdate;
      d_终止时间 := Sysdate;
      For c_详情 In (Select *
                   From (Select ID As 安排id, -1 * Null As 计划id, 科室id, 项目id, 医生姓名, Decode(医生id, 0, Null, 医生id) As 医生id, 周日,
                                 周一, 周二, 周三, 周四, 周五, 周六, 分诊方式, 序号控制, Nvl(开始时间, Sysdate - 3) As 开始时间,
                                 Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间
                          From 挂号安排
                          Where ID = c_号源.Id And Not Exists
                           (Select 1 From 挂号安排计划 Where 安排id = c_号源.Id And 停用日期 Is Null) And
                                Not (周日 Is Null And 周一 Is Null And 周二 Is Null And 周三 Is Null And 周四 Is Null And 周五 Is Null And
                                 周六 Is Null)
                          Union All
                          Select a.安排id As 安排id, a.Id As 计划id, b.科室id, a.项目id, a.医生姓名,
                                 Decode(a.医生id, 0, Null, a.医生id) As 医生id, a.周日, a.周一, a.周二, a.周三, a.周四, a.周五, a.周六, a.分诊方式,
                                 a.序号控制, Nvl(a.生效时间, Sysdate - 3) As 开始时间,
                                 Nvl(a.失效时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间
                          From 挂号安排计划 A, 挂号安排 B
                          Where a.安排id = b.Id And b.Id = c_号源.Id And b.停用日期 Is Null And
                                Not (a.周日 Is Null And a.周一 Is Null And a.周二 Is Null And a.周三 Is Null And a.周四 Is Null And
                                 a.周五 Is Null And a.周六 Is Null))
                   Order By 开始时间) Loop
        If Nvl(n_安排id, 0) <> 0 Then
          If c_详情.开始时间 < Sysdate Then
            --不导入失效安排
            Select ID Bulk Collect Into l_限制id From 临床出诊限制 Where 安排id = n_安排id;
          
            Forall I In 1 .. l_限制id.Count
              Delete From 临床出诊时段 Where 限制id = l_限制id(I);
          
            Forall I In 1 .. l_限制id.Count
              Delete From 临床出诊挂号控制 Where 限制id = l_限制id(I);
          
            Forall I In 1 .. l_限制id.Count
              Delete From 临床出诊诊室 Where 限制id = l_限制id(I);
          
            Forall I In 1 .. l_限制id.Count
              Delete From 临床出诊限制 Where ID = l_限制id(I);
          
            Delete From 临床出诊安排 Where ID = n_安排id;
          Else
            --将上次的开始时间作为本次的终止时间
            Update 临床出诊安排
            Set 终止时间 = c_详情.开始时间 - 1 / 24 / 60 / 60, 原终止时间 = c_详情.开始时间 - 1 / 24 / 60 / 60
            Where ID = n_安排id;
          End If;
        End If;
      
        If c_详情.终止时间 > d_终止时间 Then
          d_终止时间 := c_详情.终止时间;
        End If;
      
        Select 临床出诊安排_Id.Nextval Into n_安排id From Dual;
        n_诊室id := Null;
        If Nvl(c_详情.分诊方式, 0) = 1 Then
          Begin
            If Nvl(c_详情.计划id, 0) <> 0 Then
              Select a.Id
              Into n_诊室id
              From 门诊诊室 A, 挂号计划诊室 B
              Where a.名称 = b.门诊诊室 And b.计划id = c_详情.计划id And Rownum < 2;
            Else
              Select a.Id
              Into n_诊室id
              From 门诊诊室 A, 挂号安排诊室 B
              Where a.名称 = b.门诊诊室 And b.号表id = c_详情.安排id And Rownum < 2;
            End If;
          Exception
            When Others Then
              n_诊室id := Null;
          End;
        End If;
      
        --a.临床出诊安排
        Insert Into 临床出诊安排
          (ID, 出诊id, 号源id, 项目id, 医生id, 医生姓名, 排班规则, 是否周六出诊, 是否周日出诊, 开始时间, 终止时间, 操作员姓名, 登记时间, 原终止时间)
        Values
          (n_安排id, n_出诊id, n_号源id, c_详情.项目id, c_详情.医生id, c_详情.医生姓名, Null, Null, Null, c_详情.开始时间, d_终止时间, Zl_Username,
           Sysdate, d_终止时间);
      
        --b.临床出诊限制
        If Nvl(c_详情.计划id, 0) <> 0 Then
          n_预约控制 := 0;
          Begin
            Select 1 Into n_预约控制 From 挂号计划限制 Where 计划id = c_详情.计划id And 限约数 = 0 And Rownum < 2;
          Exception
            When Others Then
              Null;
          End;
        
          Select Count(1) Into n_Count From 挂号计划限制 Where 计划id = c_详情.计划id And Rownum < 2;
          If n_Count = 0 Then
            Insert Into 临床出诊限制
              (ID, 安排id, 限号数, 限约数, 是否序号控制, 是否分时段, 预约控制, 限制项目, 上班时段, 分诊方式, 诊室id)
              Select 临床出诊限制_Id.Nextval, n_安排id, Null, Null, c_详情.序号控制, Decode(b.星期, Null, 0, 1) As 是否分时段, n_预约控制, a.限制项目,
                     a.上班时段, c_详情.分诊方式, n_诊室id
              From (Select '周日' As 限制项目, c_详情.周日 As 上班时段
                     From Dual
                     Where c_详情.周日 Is Not Null
                     Union All
                     Select '周一', c_详情.周一
                     From Dual
                     Where c_详情.周一 Is Not Null
                     Union All
                     Select '周二', c_详情.周二
                     From Dual
                     Where c_详情.周二 Is Not Null
                     Union All
                     Select '周三', c_详情.周三
                     From Dual
                     Where c_详情.周三 Is Not Null
                     Union All
                     Select '周四', c_详情.周四
                     From Dual
                     Where c_详情.周四 Is Not Null
                     Union All
                     Select '周五', c_详情.周五
                     From Dual
                     Where c_详情.周五 Is Not Null
                     Union All
                     Select '周六', c_详情.周六 From Dual Where c_详情.周六 Is Not Null) A,
                   (Select Distinct 星期 From 挂号计划时段 Where 计划id = c_详情.计划id) B
              Where a.限制项目 = b.星期(+);
          Else
            Insert Into 临床出诊限制
              (ID, 安排id, 限制项目, 上班时段, 限号数, 限约数, 是否序号控制, 是否分时段, 预约控制, 分诊方式, 诊室id)
              Select 临床出诊限制_Id.Nextval, n_安排id, 限制项目,
                     Decode(限制项目, '周日', c_详情.周日, '周一', c_详情.周一, '周二', c_详情.周二, '周三', c_详情.周三, '周四', c_详情.周四, '周五',
                             c_详情.周五, '周六', c_详情.周六, Null), 限号数, 限约数, c_详情.序号控制, Decode(b.星期, Null, 0, 1) As 是否分时段,
                     n_预约控制, c_详情.分诊方式, n_诊室id
              From 挂号计划限制 A, (Select Distinct 星期 From 挂号计划时段 Where 计划id = c_详情.计划id) B
              Where a.限制项目 = b.星期(+) And 计划id = c_详情.计划id;
          End If;
        Else
          n_预约控制 := 0;
          Begin
            Select 1 Into n_预约控制 From 挂号安排限制 Where 安排id = c_详情.安排id And 限约数 = 0 And Rownum < 2;
          Exception
            When Others Then
              Null;
          End;
        
          Select Count(1) Into n_Count From 挂号安排限制 Where 安排id = c_详情.安排id And Rownum < 2;
          If n_Count = 0 Then
            Insert Into 临床出诊限制
              (ID, 安排id, 限号数, 限约数, 是否序号控制, 是否分时段, 预约控制, 限制项目, 上班时段, 分诊方式, 诊室id)
              Select 临床出诊限制_Id.Nextval, n_安排id, Null, Null, c_详情.序号控制, Decode(b.星期, Null, 0, 1) As 是否分时段, n_预约控制, a.限制项目,
                     a.上班时段, c_详情.分诊方式, n_诊室id
              From (Select '周日' As 限制项目, c_详情.周日 As 上班时段
                     From Dual
                     Where c_详情.周日 Is Not Null
                     Union All
                     Select '周一', c_详情.周一
                     From Dual
                     Where c_详情.周一 Is Not Null
                     Union All
                     Select '周二', c_详情.周二
                     From Dual
                     Where c_详情.周二 Is Not Null
                     Union All
                     Select '周三', c_详情.周三
                     From Dual
                     Where c_详情.周三 Is Not Null
                     Union All
                     Select '周四', c_详情.周四
                     From Dual
                     Where c_详情.周四 Is Not Null
                     Union All
                     Select '周五', c_详情.周五
                     From Dual
                     Where c_详情.周五 Is Not Null
                     Union All
                     Select '周六', c_详情.周六 From Dual Where c_详情.周六 Is Not Null) A,
                   (Select Distinct 星期 From 挂号安排时段 Where 安排id = c_详情.安排id) B
              Where a.限制项目 = b.星期(+);
          Else
            Insert Into 临床出诊限制
              (ID, 安排id, 限制项目, 上班时段, 限号数, 限约数, 是否序号控制, 是否分时段, 预约控制, 分诊方式, 诊室id)
              Select 临床出诊限制_Id.Nextval, n_安排id, 限制项目,
                     Decode(限制项目, '周日', c_详情.周日, '周一', c_详情.周一, '周二', c_详情.周二, '周三', c_详情.周三, '周四', c_详情.周四, '周五',
                             c_详情.周五, '周六', c_详情.周六, Null), 限号数, 限约数, c_详情.序号控制, Decode(b.星期, Null, 0, 1) As 是否分时段,
                     n_预约控制, c_详情.分诊方式, n_诊室id
              From 挂号安排限制 A, (Select Distinct 星期 From 挂号安排时段 Where 安排id = c_详情.安排id) B
              Where a.限制项目 = b.星期(+) And 安排id = c_详情.安排id;
          End If;
        End If;
      
        --c.临床出诊诊室
        If Nvl(c_详情.分诊方式, 0) > 0 Then
          If Nvl(c_详情.计划id, 0) <> 0 Then
            Insert Into 临床出诊诊室
              (限制id, 诊室id)
              Select a.Id, b.诊室id
              From 临床出诊限制 A,
                   (Select Distinct a.Id As 诊室id
                     From 门诊诊室 A, 挂号计划诊室 B
                     Where a.名称 = b.门诊诊室 And b.计划id = c_详情.计划id) B
              Where a.安排id = n_安排id;
          Else
            Insert Into 临床出诊诊室
              (限制id, 诊室id)
              Select a.Id, b.诊室id
              From 临床出诊限制 A,
                   (Select Distinct a.Id As 诊室id
                     From 门诊诊室 A, 挂号安排诊室 B
                     Where a.名称 = b.门诊诊室 And b.号表id = c_详情.安排id) B
              Where a.安排id = n_安排id;
          End If;
        End If;
      
        --D.临床出诊时段
        If Nvl(c_详情.计划id, 0) <> 0 Then
          Insert Into 临床出诊时段
            (限制id, 序号, 开始时间, 终止时间, 限制数量, 是否预约)
            Select a.Id, b.序号, b.开始时间, b.结束时间, b.限制数量, b.是否预约
            From 临床出诊限制 A,
                 (Select n_安排id As 安排id, 星期,
                          Decode(星期, '周日', c_详情.周日, '周一', c_详情.周一, '周二', c_详情.周二, '周三', c_详情.周三, '周四', c_详情.周四, '周五',
                                  c_详情.周五, '周六', c_详情.周六, Null) As 上班时段, 序号, 开始时间, 结束时间, 限制数量, 是否预约
                   From 挂号计划时段
                   Where 计划id = c_详情.计划id) B
            Where a.安排id = b.安排id And a.限制项目 = b.星期 And a.上班时段 = b.上班时段;
        
        Else
          Insert Into 临床出诊时段
            (限制id, 序号, 开始时间, 终止时间, 限制数量, 是否预约)
            Select a.Id, b.序号, b.开始时间, b.结束时间, b.限制数量, b.是否预约
            From 临床出诊限制 A,
                 (Select n_安排id As 安排id, 星期,
                          Decode(星期, '周日', c_详情.周日, '周一', c_详情.周一, '周二', c_详情.周二, '周三', c_详情.周三, '周四', c_详情.周四, '周五',
                                  c_详情.周五, '周六', c_详情.周六, Null) As 上班时段, 序号, 开始时间, 结束时间, 限制数量, 是否预约
                   From 挂号安排时段
                   Where 安排id = c_详情.安排id) B
            Where a.安排id = b.安排id And a.限制项目 = b.星期 And a.上班时段 = b.上班时段;
        End If;
      
        --不分时段的序号控制号先生成序号
        For c_限制项目 In (Select ID, 限号数
                       From 临床出诊限制
                       Where 安排id = n_安排id And Nvl(限号数, 0) <> 0 And Nvl(是否序号控制, 0) = 1 And Nvl(是否分时段, 0) = 0) Loop
          For I In 1 .. c_限制项目.限号数 Loop
            Insert Into 临床出诊时段 (限制id, 序号, 限制数量, 是否预约) Values (c_限制项目.Id, I, 1, 1);
          End Loop;
        End Loop;
      
        --任何一个都不允许预约时表示全部允许预约
        Update 临床出诊时段 A
        Set a.是否预约 = 1
        Where 限制id In (Select ID From 临床出诊限制 Where 安排id = n_安排id) And Not Exists
         (Select 1 From 临床出诊时段 B Where a.限制id = b.限制id And Nvl(b.是否预约, 0) = 1);
      
        --E.合作单位挂号控制
        If Nvl(c_详情.计划id, 0) <> 0 Then
        
          Insert Into 临床出诊挂号控制
            (限制id, 类型, 性质, 名称, 序号, 控制方式, 数量)
            Select a.Id, b.类型, b.性质, b.合作单位, b.序号, b.控制方式, b.数量
            From 临床出诊限制 A,
                 (Select 1 As 类型, 1 As 性质, 合作单位, n_安排id As 安排id, 限制项目,
                          Decode(限制项目, '周日', c_详情.周日, '周一', c_详情.周一, '周二', c_详情.周二, '周三', c_详情.周三, '周四', c_详情.周四, '周五',
                                  c_详情.周五, '周六', c_详情.周六, Null) As 上班时段, 序号,
                          Case
                             When Nvl(序号, 0) = 0 And Nvl(数量, 0) = 0 Then
                              0
                             When 序号 = 0 And Nvl(数量, 0) <> 0 Then
                              2
                             When Nvl(序号, 0) <> 0 And Nvl(数量, 0) <> 0 Then
                              3
                             Else
                              4
                           End As 控制方式, 数量
                   From 合作单位计划控制
                   Where 计划id = c_详情.计划id And
                         Decode(限制项目, '周日', c_详情.周日, '周一', c_详情.周一, '周二', c_详情.周二, '周三', c_详情.周三, '周四', c_详情.周四, '周五',
                                c_详情.周五, '周六', c_详情.周六, Null) Is Not Null) B
            Where a.安排id = b.安排id And a.限制项目 = b.限制项目 And a.上班时段 = b.上班时段;
        
        Else
          Insert Into 临床出诊挂号控制
            (限制id, 类型, 性质, 名称, 序号, 控制方式, 数量)
            Select a.Id, b.类型, b.性质, b.合作单位, b.序号, b.控制方式, b.数量
            From 临床出诊限制 A,
                 (Select 1 As 类型, 1 As 性质, 合作单位, n_安排id As 安排id, 限制项目,
                          Decode(限制项目, '周日', c_详情.周日, '周一', c_详情.周一, '周二', c_详情.周二, '周三', c_详情.周三, '周四', c_详情.周四, '周五',
                                  c_详情.周五, '周六', c_详情.周六, Null) As 上班时段, 序号,
                          Case
                            When Nvl(序号, 0) = 0 And Nvl(数量, 0) = 0 Then
                             0
                            When 序号 = 0 And Nvl(数量, 0) <> 0 Then
                             2
                            When Nvl(序号, 0) <> 0 And Nvl(数量, 0) <> 0 Then
                             3
                            Else
                             4
                          End As 控制方式, 数量
                   From 合作单位安排控制
                   Where 安排id = c_详情.安排id And
                         Decode(限制项目, '周日', c_详情.周日, '周一', c_详情.周一, '周二', c_详情.周二, '周三', c_详情.周三, '周四', c_详情.周四, '周五',
                                c_详情.周五, '周六', c_详情.周六, Null) Is Not Null) B
            Where a.安排id = b.安排id And a.限制项目 = b.限制项目 And a.上班时段 = b.上班时段;
        End If;
      End Loop;
    
      --4.拷贝一份出诊信息作为号源控制信息
      --说明：1.同一号源多个安排/计划时只导入最后一个安排对应的出诊信息
      --      2.上班时段按星期排序(周一到周日)取第一个
      For c_限制 In (Select ID, 上班时段, 限号数, 限约数, 是否序号控制, 是否分时段, 预约控制, 是否独占, 分诊方式, 诊室id
                   From (Select ID, 上班时段, 限号数, 限约数, 是否序号控制, 是否分时段, 预约控制, 是否独占, 分诊方式, 诊室id,
                                 Row_Number() Over(Partition By 上班时段 Order By Decode(限制项目, '周一', 1, '周二', 2, '周三', 3, '周四', 4, '周五', 5, '周六', 6, '周日', 7)) As 组号
                          From 临床出诊限制
                          Where 安排id = n_安排id)
                   Where 组号 = 1) Loop
        --a.临床出诊号源限制
        Select 临床出诊号源限制_Id.Nextval Into n_限制id From Dual;
        Insert Into 临床出诊号源限制
          (ID, 号源id, 上班时段, 限号数, 限约数, 是否序号控制, 是否分时段, 预约控制, 是否独占, 分诊方式, 诊室id)
        Values
          (n_限制id, n_号源id, c_限制.上班时段, c_限制.限号数, c_限制.限约数, c_限制.是否序号控制, c_限制.是否分时段, c_限制.预约控制, c_限制.是否独占, c_限制.分诊方式,
           c_限制.诊室id);
        --b.临床出诊号源诊室
        Insert Into 临床出诊号源诊室
          (限制id, 诊室id)
          Select n_限制id, 诊室id From 临床出诊诊室 Where 限制id = c_限制.Id;
        --c.临床出诊号源时段
        Insert Into 临床出诊号源时段
          (限制id, 序号, 开始时间, 终止时间, 限制数量, 是否预约)
          Select n_限制id, 序号, 开始时间, 终止时间, 限制数量, 是否预约 From 临床出诊时段 Where 限制id = c_限制.Id;
        --d.临床出诊号源控制
        Insert Into 临床出诊号源控制
          (限制id, 类型, 性质, 名称, 序号, 控制方式, 数量)
          Select n_限制id, 类型, 性质, 名称, 序号, 控制方式, 数量 From 临床出诊挂号控制 Where 限制id = c_限制.Id;
      End Loop;
    End Loop;
  End;
Begin
  Select Count(1) Into n_Count From 临床出诊表 Where Rownum < 2;
  If n_Count <> 0 Then
    v_Err_Msg := '当前已经存在出诊表了，不允许再导入！';
    Raise Err_Item;
  End If;

  Begin
    Select f_List2str(Cast(Collect(s.时间段) As t_Strlist))
    Into v_时间段
    From (Select 时间段, Row_Number() Over(Partition By 时间段 Order By 时间段) As 组号
           From (Select Decode(b.行号, 1, a.周一, 2, a.周二, 3, a.周三, 4, a.周四, 5, a.周五, 6, a.周六, a.周日) As 时间段
                  From 挂号安排 A, (Select Level As 行号 From Dual Connect By Level <= 7) B
                  Where a.停用日期 Is Null And Nvl(a.终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate And
                        (号码_In = a.号码 Or 号码_In Is Null)
                  Union All
                  Select Decode(n.行号, 1, m.周一, 2, m.周二, 3, m.周三, 4, m.周四, 5, m.周五, 6, m.周六, m.周日) As 时间段
                  From 挂号安排计划 M, (Select Level As 行号 From Dual Connect By Level <= 7) N
                  Where Nvl(m.失效时间, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate And (号码_In = m.号码 Or 号码_In Is Null))) S,
         时间段 T
    Where s.时间段 = t.时间段(+) And t.时间段 Is Null And s.组号 = 1;
  Exception
    When Others Then
      v_时间段 := Null;
  End;

  If v_时间段 Is Not Null Then
    v_Err_Msg := '原挂号安排中的上班时间段【' || v_时间段 || '】不存在，请先在“基础设置>上班时间管理”中添加！';
    Raise Err_Item;
  End If;

  If Not 号码_In Is Null Then
    Zl_Register_Import(号码_In);
    Return;
  End If;

  --删除现有所有号源，在调用之前已进行了提示
  Select ID Bulk Collect Into l_限制id From 临床出诊号源限制;

  Forall I In 1 .. l_限制id.Count
    Delete From 临床出诊号源诊室 Where 限制id = l_限制id(I);

  Forall I In 1 .. l_限制id.Count
    Delete From 临床出诊号源时段 Where 限制id = l_限制id(I);

  Forall I In 1 .. l_限制id.Count
    Delete From 临床出诊号源控制 Where 限制id = l_限制id(I);

  Forall I In 1 .. l_限制id.Count
    Delete From 临床出诊号源限制 Where ID = l_限制id(I);

  Delete From 临床出诊号源;

  For c_号源 In (Select 号码
               From 挂号安排
               Where Nvl(是否删除, 0) = 0 And
                     Nvl(停用日期, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd')) Loop
  
    Zl_Register_Import(c_号源.号码);
  End Loop;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊表_导入;
/
Create Or Replace Procedure Zl1_Auto_Buildingregisterplan(挂号时间_In In Date := Null) As

  -------------------------------------------------------------------------
  --功能说明：自动生成临床出诊记录
  --          1、根据号源自动生成预约数内的临床出诊记录;
  --          2、预约天数的确定:号源预约天数-->预约方式的天数（取最大)-->系统预约天数
  --入参:挂号时间_IN:NULL时，自动生成;否则只检查指定日期是否生成了出诊记录没有

  -------------------------------------------------------------------------
  n_缺省预约天数 Number(10);
  v_操作员姓名   临床出诊安排.操作员姓名%Type;
  n_记录id       临床出诊记录.Id%Type;
  n_安排id       临床出诊安排.Id%Type;

  d_出诊日期 Date;
  d_登记日期 Date;
  d_换休日期 Date;
  d_当前日期 Date;
  d_开始日期 Date;
  d_终止日期 Date;
  v_停诊原因 临床出诊记录.停诊原因%Type;
  v_限制项目 临床出诊限制.限制项目%Type;

  n_节假日   Number(2);
  v_节日名称 法定假日表.节日名称%Type;
  n_是否出诊 Number(2);
  n_Count    Number(18);
Begin

  Select Max(预约天数) Into n_缺省预约天数 From 预约方式;
  If Nvl(n_缺省预约天数, 0) = 0 Then
    n_缺省预约天数 := To_Number(Nvl(zl_GetSysParameter('挂号允许预约天数'), '0'));
  End If;
  If Nvl(n_缺省预约天数, 0) = 0 Then
    n_缺省预约天数 := 7;
  End If;

  d_当前日期   := Trunc(Nvl(挂号时间_In, Sysdate));
  d_登记日期   := Sysdate;
  v_操作员姓名 := Zl_Username;
  For c_号源 In (Select c.Id, c.号类, c.号码, c.科室id, c.医生姓名, Decode(Nvl(c.预约天数, 0), 0, n_缺省预约天数, c.预约天数) As 预约天数,
                      Nvl(b.站点, '-') As 站点, Nvl(c.是否假日换休, 0) As 是否假日换休, Nvl(c.假日控制状态, 0) As 假日控制状态
               From 临床出诊号源 C, 部门表 B
               Where c.科室id = b.Id And Nvl(c.是否删除, 0) = 0 And
                     (c.撤档时间 Is Null Or c.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd')) And Exists
                (Select 1
                      From 临床出诊安排 M, 临床出诊表 N
                      Where m.出诊id = n.Id And Nvl(n.排班方式, 0) = 0 And n.发布时间 Is Not Null And m.号源id = c.Id And
                            m.终止时间 >= d_当前日期) And Not Exists
                (Select 1
                      From 临床出诊记录
                      Where 号源id = c.Id And 出诊日期 = d_当前日期 + Decode(Nvl(c.预约天数, 0), 0, n_缺省预约天数, c.预约天数))) Loop
  
    For c_日期信息 In (Select a.日期, b.安排id, b.是否周六出诊, b.是否周日出诊,
                          Decode(To_Char(a.日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                                  '7', '周六', Null) As 限制项目
                   From (Select Trunc(Sysdate) + 天数 As 日期
                          From (Select Level - 1 As 天数 From Dual Connect By Level <= c_号源.预约天数 + 1)
                          Minus
                          Select Trunc(出诊日期) As 日期
                          From 临床出诊记录
                          Where 出诊日期 Between Trunc(Sysdate) And Trunc(Sysdate) + c_号源.预约天数 And 号源id = c_号源.Id) A,
                        (Select m.Id As 安排id, m.开始时间, m.终止时间, m.是否周六出诊, m.是否周日出诊
                          From 临床出诊安排 M, 临床出诊表 N
                          Where m.号源id = c_号源.Id And m.出诊id = n.Id And Nvl(n.排班方式, 0) = 0 And n.发布时间 Is Not Null) B
                   Where a.日期 Between b.开始时间 And b.终止时间) Loop
      d_出诊日期 := c_日期信息.日期;
      v_限制项目 := c_日期信息.限制项目;
      n_安排id   := c_日期信息.安排id;
    
      n_节假日   := 0;
      n_是否出诊 := 1;
      d_开始日期 := Null;
      d_终止日期 := Null;
      v_停诊原因 := Null;
      Begin
        --需要确定
        Select 1, 节日名称
        Into n_节假日, v_节日名称
        From 法定假日表
        Where d_出诊日期 Between 开始日期 And 终止日期 And 性质 = 0;
      Exception
        When Others Then
          Null;
      End;
      --假日控制状态：0-不上班;1-上班且开放预约;2-上班但不开放预约
      If Nvl(c_号源.假日控制状态, 0) = 0 And n_节假日 = 1 Then
        n_是否出诊 := 0;
        v_停诊原因 := v_节日名称;
      End If;
    
      d_换休日期 := Null;
      If Nvl(c_号源.是否假日换休, 0) = 1 And n_节假日 = 0 Then
        Begin
          --需要确定当前日期是否由某一天换休过来的
          --开始日期：原本休息日(即调休日) ， 终止日期：原本上班日(即被调休日)
          Select 终止日期 Into d_换休日期 From 法定假日表 Where 开始日期 = d_出诊日期 And 性质 = 1;
        Exception
          When Others Then
            Null;
        End;
        --当前是换休日，不管是周六，周日都应该上班
        If Not d_换休日期 Is Null Then
          n_是否出诊 := 1;
          Select Decode(To_Char(d_换休日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7',
                         '周六', Null)
          Into v_限制项目
          From Dual;
        End If;
      End If;
    
      If Nvl(n_安排id, 0) = 0 Then
        n_是否出诊 := 0;
      End If;
      --检查这天是否出诊
      If n_是否出诊 = 1 Then
        Select Count(*) Into n_Count From 临床出诊限制 Where 安排id = n_安排id And 限制项目 = v_限制项目;
        If Nvl(n_Count, 0) = 0 Then
          n_是否出诊 := 0;
        End If;
      End If;
    
      If Nvl(n_是否出诊, 0) = 0 Then
        --增加临床出诊记录(时间段为NULL 的空记录)
        Insert Into 临床出诊记录
          (ID, 安排id, 号源id, 出诊日期, 登记人, 登记时间)
          Select 临床出诊记录_Id.Nextval, c_日期信息.安排id, a.Id As ID, c_日期信息.日期, v_操作员姓名, d_登记日期 As 登记时间
          From 临床出诊号源 A, 临床出诊安排 B
          Where a.Id = b.号源id And b.Id = c_日期信息.安排id;
      Else
        --处理请假数据
        Begin
          Select Min(开始时间), Max(终止时间), Max(停诊原因)
          Into d_开始日期, d_开始日期, v_停诊原因
          From 临床出诊停诊记录
          Where 记录id Is Null And c_日期信息.日期 Between 开始时间 And 终止时间 And 申请人 = c_号源.医生姓名 And 审批人 Is Not Null And
                取消人 Is Null And Rownum < 2;
        Exception
          When Others Then
            d_开始日期 := Null;
            d_终止日期 := Null;
            v_停诊原因 := Null;
        End;
      
        For c_记录 In (With c_时间段 As
                        (Select 时间段, 开始时间, 终止时间, 号类, 站点, 缺省时间, 提前时间
                        From (Select 时间段, 开始时间, 终止时间, 号类, 站点, 缺省时间, 提前时间,
                                      Row_Number() Over(Partition By 时间段 Order By 时间段, 站点 Asc, 号类 Asc) As 组号
                               From 时间段
                               Where Nvl(站点, c_号源.站点) = c_号源.站点 And Nvl(号类, c_号源.号类) = c_号源.号类)
                        Where 组号 = 1)
                       Select c_日期信息.安排id As 安排id, B1.号源id, c_日期信息.日期 As 出诊日期, m.上班时段, m.Id As 限制id,
                              To_Date(To_Char(c_日期信息.日期, 'yyyy-mm-dd ') || To_Char(j.开始时间, 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                              To_Date(To_Char(c_日期信息.日期, 'yyyy-mm-dd ') || To_Char(j.终止时间, 'hh24:mi:ss'),
                                      'yyyy-mm-dd hh24:mi:ss') + Case
                                When j.终止时间 <= j.开始时间 Then
                                 1
                                Else
                                 0
                              End As 终止时间, d_开始日期 As 停诊开始时间, d_终止日期 As 停诊终止时间, v_停诊原因 As 停诊原因,
                              To_Date(To_Char(c_日期信息.日期, 'yyyy-mm-dd ') || To_Char(Nvl(j.缺省时间, j.开始时间), 'hh24:mi:ss'),
                                      'yyyy-mm-dd hh24:mi:ss') + Case
                                When j.缺省时间 < j.开始时间 Then
                                 1
                                Else
                                 0
                              End As 缺省预约时间,
                              To_Date(To_Char(c_日期信息.日期, 'yyyy-mm-dd ') || To_Char(Nvl(j.提前时间, j.开始时间), 'hh24:mi:ss'),
                                      'yyyy-mm-dd hh24:mi:ss') + Case
                                When j.开始时间 < j.提前时间 Then
                                 -1
                                Else
                                 0
                              End As 提前挂号时间, m.限号数, 0 As 已挂数, m.限约数, 0 As 已约数, 0 As 其中已接收, m.是否序号控制, m.是否分时段, m.预约控制,
                              B1.项目id, B1.医生id, B1.医生姓名, Null As 替诊医生id, Null As 替诊医生姓名, m.分诊方式, m.诊室id, 0 As 是否锁定,
                              0 As 是否临时出诊, v_操作员姓名 As 操作员姓名, d_登记日期 As 登记时间, v_限制项目 As 限制项目
                       From 临床出诊安排 B1, 临床出诊限制 M, c_时间段 J
                       Where B1.Id = n_安排id And B1.Id = m.安排id And m.限制项目 = v_限制项目 And m.上班时段 = j.时间段 And
                             To_Date(To_Char(c_日期信息.日期, 'yyyy-mm-dd ') || To_Char(j.开始时间, 'hh24:mi:ss'),
                                     'yyyy-mm-dd hh24:mi:ss') >= B1.开始时间) Loop
        
          Select 临床出诊记录_Id.Nextval Into n_记录id From Dual;
          Insert Into 临床出诊记录
            (ID, 安排id, 号源id, 出诊日期, 上班时段, 开始时间, 终止时间, 停诊开始时间, 停诊终止时间, 停诊原因, 缺省预约时间, 提前挂号时间, 限号数, 已挂数, 限约数, 已约数, 其中已接收,
             是否序号控制, 是否分时段, 预约控制, 项目id, 科室id, 医生id, 医生姓名, 替诊医生id, 替诊医生姓名, 分诊方式, 诊室id, 是否锁定, 是否临时出诊, 登记人, 登记时间, 是否发布)
          Values
            (n_记录id, c_记录.安排id, c_记录.号源id, c_记录.出诊日期, c_记录.上班时段, c_记录.开始时间, c_记录.终止时间,
             Case When c_记录.停诊开始时间 Is Null Then Null When c_记录.停诊开始时间 < c_记录.开始时间 Then c_记录.开始时间 Else c_记录.停诊开始时间 End,
             Case When c_记录.停诊终止时间 Is Null Then Null When c_记录.停诊终止时间 > c_记录.终止时间 Then c_记录.终止时间 Else c_记录.停诊终止时间 End,
             c_记录.停诊原因, c_记录.缺省预约时间, c_记录.提前挂号时间, c_记录.限号数, c_记录.已挂数, c_记录.限约数, c_记录.已约数, c_记录.其中已接收, c_记录.是否序号控制,
             c_记录.是否分时段, c_记录.预约控制, c_记录.项目id, c_号源.科室id, c_记录.医生id, c_记录.医生姓名, c_记录.替诊医生id, c_记录.替诊医生姓名, c_记录.分诊方式,
             c_记录.诊室id, c_记录.是否锁定, c_记录.是否临时出诊, c_记录.操作员姓名, d_登记日期, 1);
        
          --插入临床出诊序号控制
          Insert Into 临床出诊序号控制
            (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约)
            Select n_记录id, 序号,
                   To_Date(To_Char(c_记录.出诊日期, 'yyyy-mm-dd ') || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
                   To_Date(To_Char(c_记录.出诊日期, 'yyyy-mm-dd ') || To_Char(终止时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') + Case
                     When 终止时间 <= 开始时间 Then
                      1
                     Else
                      0
                   End, 限制数量, 是否预约
            From 临床出诊时段
            Where 限制id = c_记录.限制id;
        
          --插入合作单位挂号控制记录
          Insert Into 临床出诊挂号控制记录
            (类型, 性质, 名称, 记录id, 序号, 控制方式, 数量)
            Select 类型, 性质, 名称, n_记录id, 序号, 控制方式, 数量 From 临床出诊挂号控制 Where 限制id = c_记录.限制id;
        
          --插入临床出诊诊室记录
          Insert Into 临床出诊诊室记录
            (记录id, 诊室id)
            Select n_记录id, 诊室id From 临床出诊诊室 Where 限制id = c_记录.限制id;
        
        End Loop;
      End If;
      --一天一提交
      Commit;
    End Loop;
  End Loop;
End Zl1_Auto_Buildingregisterplan;
/
Create Or Replace Procedure Zl_临床出诊表_Add
(
  操作类型_In Number,
  出诊id_In   临床出诊表.Id%Type,
  出诊表名_In 临床出诊表.出诊表名%Type,
  站点_In     部门表.站点%Type,
  操作员_In   临床出诊安排.操作员姓名%Type,
  操作时间_In 临床出诊安排.登记时间%Type,
  开始时间_In 临床出诊安排.开始时间%Type := Null,
  终止时间_In 临床出诊安排.终止时间%Type := Null,
  年份_In     临床出诊表.年份%Type := Null,
  月份_In     临床出诊表.月份%Type := Null,
  周数_In     临床出诊表.周数%Type := Null,
  应用范围_In 临床出诊表.应用范围%Type := Null,
  科室id_In   临床出诊表.科室id%Type := Null,
  备注_In     临床出诊表.备注%Type := Null,
  人员id_In   人员表.Id%Type := Null,
  删除安排_In Number := 0
) As
  --功能：增加出诊表或模板
  --参数：
  --        操作类型_In 1-模板，2-固定安排, 3-月安排，4-周安排
  --        人员id_In 除固定安排外有效，不为0或null表示临床科室人员在添加
  --        删除安排_In 固定排班转为月排班/周排班时，在制定月排班/周排班时是否删除新出诊表时间内未使用的出诊记录
  --说明：
  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  n_Count Number(8);

  n_出诊id 临床出诊表.Id%Type;
  l_记录id t_Numlist := t_Numlist();
  l_安排id t_Numlist := t_Numlist();
Begin
  n_出诊id := 出诊id_In;
  If Nvl(n_出诊id, 0) = 0 Then
    Select 临床出诊表_Id.Nextval Into n_出诊id From Dual;
  End If;

  --排班方式：0-固定排班;1-按月排班;2-按周排班;3-模板
  --============================================================================================================================================
  --1.模板
  If Nvl(操作类型_In, 0) = 1 Then
    Begin
      Select 1 Into n_Count From 临床出诊表 Where 出诊表名 = 出诊表名_In And 排班方式 = 3 And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If Nvl(n_Count, 0) > 0 Then
      v_Err_Msg := '当前已存在名为“' || 出诊表名_In || '”的模板！';
      Raise Err_Item;
    End If;
  
    --检查是否有可操作的有效号源
    Begin
      Select 1
      Into n_Count
      From 临床出诊号源 A, 部门表 D
      Where a.科室id = d.Id And a.排班方式 In (1, 2) And Nvl(a.是否删除, 0) = 0 And
            (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'))
           --当前人员可操作的号源
            And (Nvl(人员id_In, 0) = 0 Or
            (Nvl(a.是否临床排班, 0) = 1 And Exists (Select 1 From 部门人员 Where 部门id = a.科室id And 人员id = 人员id_In)))
           --站点
            And (d.站点 Is Null Or d.站点 = 站点_In) And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If Nvl(n_Count, 0) = 0 Then
      v_Err_Msg := '当前无可按月或按周排班的号源，不能新增模板，请先到“基础设置>临床号源管理”中添加出诊号源！';
      Raise Err_Item;
    End If;
  
    --模板，肯定是新出诊表
    Insert Into 临床出诊表
      (ID, 排班方式, 出诊表名, 应用范围, 科室id, 备注, 发布人, 发布时间)
    Values
      (n_出诊id, 3, 出诊表名_In, 应用范围_In, 科室id_In, 备注_In, 操作员_In, 操作时间_In);
  
    --临床出诊安排
    For c_号源 In (Select 临床出诊安排_Id.Nextval As 安排id, n_出诊id As 出诊id, a.Id As 号源id, a.项目id, a.医生id, a.医生姓名
                 From 临床出诊号源 A, 部门表 D
                 Where a.科室id = d.Id And a.排班方式 In (1, 2) And Nvl(a.是否删除, 0) = 0 And
                       (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'))
                      --当前人员可操作的号源
                       And (Nvl(人员id_In, 0) = 0 Or (Nvl(a.是否临床排班, 0) = 1 And Exists
                        (Select 1 From 部门人员 Where 部门id = a.科室id And 人员id = 人员id_In)))
                      --站点
                       And (d.站点 Is Null Or d.站点 = 站点_In)) Loop
    
      Insert Into 临床出诊安排
        (ID, 出诊id, 号源id, 项目id, 医生id, 医生姓名, 操作员姓名, 登记时间)
      Values
        (c_号源.安排id, c_号源.出诊id, c_号源.号源id, c_号源.项目id, c_号源.医生id, c_号源.医生姓名, 操作员_In, 操作时间_In);
    End Loop;
    Return;
  End If;

  --============================================================================================================================================
  --2.固定排班
  If Nvl(操作类型_In, 0) = 2 Then
    Begin
      Select 1 Into n_Count From 临床出诊表 Where 出诊表名 = 出诊表名_In And 排班方式 = 0 And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If Nvl(n_Count, 0) > 0 Then
      v_Err_Msg := '当前已存在名为“' || 出诊表名_In || '”的固定出诊表！';
      Raise Err_Item;
    End If;
  
    Begin
      Select 1
      Into n_Count
      From 临床出诊安排 A, 临床出诊表 B
      Where a.出诊id = b.Id And b.排班方式 = 0 And a.开始时间 = 开始时间_In And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If Nvl(n_Count, 0) > 0 Then
      v_Err_Msg := '已存在为当前开始时间的固定安排！';
      Raise Err_Item;
    End If;
  
    --检查是否有有效号源
    Begin
      Select 1
      Into n_Count
      From 临床出诊号源 A, 部门表 D
      Where a.科室id = d.Id And a.排班方式 = 0 And Nvl(a.是否删除, 0) = 0 And
            (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'))
           --站点
            And (d.站点 Is Null Or d.站点 = 站点_In) And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If Nvl(n_Count, 0) = 0 Then
      v_Err_Msg := '当前无可按固定排班的号源，不能新增固定安排，请先到“基础设置>临床号源管理”中添加出诊号源！';
      Raise Err_Item;
    End If;
  
    --固定安排，肯定是新出诊表,只有有"所有科室"权限的人才能新增
    Insert Into 临床出诊表
      (ID, 排班方式, 出诊表名, 年份)
    Values
      (n_出诊id, 0, 出诊表名_In, To_Number(To_Char(开始时间_In, 'yyyy')));
  
    --缺省加入上一次有效的出诊安排
    For c_号源 In (Select 临床出诊安排_Id.Nextval As 安排id, n_出诊id As 出诊id, 原安排id, 号源id, 项目id, 医生id, 医生姓名
                 From (Select a.Id As 原安排id, b.Id As 号源id, b.项目id, b.医生id, b.医生姓名,
                               Row_Number() Over(Partition By b.Id Order By a.开始时间 Desc) As 组号
                        From 临床出诊安排 A, 临床出诊号源 B, 临床出诊表 C, 部门表 D
                        Where a.号源id = b.Id And a.出诊id = c.Id And b.科室id = d.Id
                             --号源限制
                              And b.排班方式 = 0 And Nvl(b.是否删除, 0) = 0 And
                              (b.撤档时间 Is Null Or b.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'))
                             --上一次出诊安排限制
                              And c.发布人 Is Not Null And c.排班方式 = 0 And (d.站点 Is Null Or d.站点 = 站点_In)) M
                 Where 组号 = 1) Loop
    
      Insert Into 临床出诊安排
        (ID, 出诊id, 号源id, 项目id, 医生id, 医生姓名, 开始时间, 终止时间, 操作员姓名, 登记时间, 原终止时间)
      Values
        (c_号源.安排id, c_号源.出诊id, c_号源.号源id, c_号源.项目id, c_号源.医生id, c_号源.医生姓名, 开始时间_In, 终止时间_In, 操作员_In, 操作时间_In, 终止时间_In);
    
      --复制出诊安排
      For c_限制 In (Select ID From 临床出诊限制 Where 安排id = c_号源.原安排id) Loop
        Zl_临床出诊限制_Copy(c_限制.Id, c_号源.安排id);
      End Loop;
    End Loop;
  
    --加入无上一次有效出诊安排的号源
    For c_号源 In (Select 临床出诊安排_Id.Nextval As 安排id, n_出诊id As 出诊id, a.Id As 号源id, a.项目id, a.医生id, a.医生姓名
                 From 临床出诊号源 A, 部门表 D
                 Where a.科室id = d.Id And a.排班方式 = 0 And Nvl(a.是否删除, 0) = 0 And
                       (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'))
                      --站点
                       And (d.站点 Is Null Or d.站点 = 站点_In)
                      
                       And Not Exists (Select 1 From 临床出诊安排 Where 出诊id = n_出诊id And 号源id = a.Id)) Loop
    
      Insert Into 临床出诊安排
        (ID, 出诊id, 号源id, 项目id, 医生id, 医生姓名, 开始时间, 终止时间, 操作员姓名, 登记时间, 原终止时间)
      Values
        (c_号源.安排id, c_号源.出诊id, c_号源.号源id, c_号源.项目id, c_号源.医生id, c_号源.医生姓名, 开始时间_In, 终止时间_In, 操作员_In, 操作时间_In, 终止时间_In);
    End Loop;
    Return;
  End If;

  --============================================================================================================================================
  --月排班、周排班
  --检查是否有有效号源
  Begin
    Select 1
    Into n_Count
    From 临床出诊号源 A, 部门表 B
    Where a.科室id = b.Id
         --有效号源
          And Nvl(a.是否删除, 0) = 0 And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd')) And
          (
          --月排班
           Nvl(操作类型_In, 0) = 3 And a.排班方式 = 1
          --周排班
           Or Nvl(操作类型_In, 0) = 4 And
           (
           --当前出诊表所在时间范围内不能有月排班
            a.排班方式 = 2 And Not Exists
            (Select 1
                From 临床出诊安排 P, 临床出诊表 Q
                Where p.出诊id = q.Id And p.号源id = a.Id And
                      Not (p.终止时间 < Trunc(开始时间_In, 'MONTH') Or p.开始时间 > Last_Day(开始时间_In)) And q.排班方式 = 1)
           --当前已调整为了月排班,但是本月又用了周排班，则本月剩下的部分将继续按周进行排班
            Or a.排班方式 = 1 And Exists
            (Select 1
                From 临床出诊安排 P, 临床出诊表 Q
                Where p.出诊id = q.Id And p.号源id = a.Id And
                      Not (p.终止时间 < Trunc(开始时间_In, 'MONTH') Or p.开始时间 > Last_Day(开始时间_In)) And q.排班方式 = 2)))
         --号源在该出诊表时间范围内无出诊记录
          And Not Exists
     (Select 1
           From 临床出诊记录 O, 临床出诊安排 P, 临床出诊表 Q
           Where o.安排id = p.Id And p.出诊id = q.Id And p.号源id = a.Id And o.出诊日期 Between 开始时间_In And 终止时间_In And
                 (q.排班方式 In (1, 2)
                 --原来为固定出诊安排
                 Or q.排班方式 = 0 And (Nvl(删除安排_In, 0) = 0 Or Nvl(删除安排_In, 0) = 1 And Exists
                  (Select 1 From 病人挂号记录 Where 出诊记录id = a.Id))))
         --当前人员可操作的号源
          And (Nvl(人员id_In, 0) = 0 Or
          (Nvl(a.是否临床排班, 0) = 1 And Exists (Select 1 From 部门人员 Where 部门id = a.科室id And 人员id = 人员id_In)))
         --站点
          And (b.站点 Is Null Or b.站点 = 站点_In) And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If Nvl(n_Count, 0) = 0 Then
    If Nvl(操作类型_In, 0) = 3 Then
      v_Err_Msg := '当前无可按月排班的号源，不能新增月出诊表，请先到“基础设置>临床号源管理”中添加出诊号源！';
    Else
      v_Err_Msg := '当前无可按周排班的号源，不能新增周出诊表，请先到“基础设置>临床号源管理”中添加出诊号源！';
    End If;
    Raise Err_Item;
  End If;

  --出诊表存在，则不再新增出诊表，直接向该出诊表添加上次有效号源安排即可
  --涉及到临床排班，当前操作员可能只能操作某一部分号源
  Begin
    Select 1 Into n_Count From 临床出诊表 Where ID = n_出诊id;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If Nvl(n_Count, 0) = 0 Then
    Insert Into 临床出诊表
      (ID, 排班方式, 出诊表名, 年份, 月份, 周数)
    Values
      (n_出诊id,
       Case
          When Nvl(操作类型_In, 0) = 3 Then
           1
          Else
           2
        End, 出诊表名_In, 年份_In, 月份_In, 周数_In);
  End If;

  --如果当前出诊表时间范围内无挂号且无预约的出诊记录(固定安排)，则删除这部分出诊记录(在删除出诊表时可恢复)，
  --并修改固定安排的终止时间，程序中已询问
  If Nvl(删除安排_In, 0) = 1 Then
    For c_安排 In (Select b.Id As 安排id
                 From 临床出诊安排 B, 临床出诊表 C, 临床出诊号源 D
                 Where b.出诊id = c.Id And b.号源id = d.Id
                      --号源
                       And Nvl(d.是否删除, 0) = 0 And (d.撤档时间 Is Null Or d.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd')) And
                       Nvl(d.排班方式, 0) = Decode(Nvl(操作类型_In, 0), 3, 1, 2)
                      --安排有被使用了的出诊记录
                       And c.排班方式 = 0 And b.终止时间 >= 开始时间_In And Not Exists
                  (Select 1
                        From 临床出诊记录 M, 病人挂号记录 N
                        Where m.安排id = b.Id And m.Id = n.出诊记录id And m.出诊日期 >= 开始时间_In)
                      --当前人员可操作的号源
                       And (Nvl(人员id_In, 0) = 0 Or (Nvl(d.是否临床排班, 0) = 1 And Exists
                        (Select 1 From 部门人员 Where 部门id = d.科室id And 人员id = 人员id_In)))) Loop
      l_安排id.Extend();
      l_安排id(l_安排id.Count) := c_安排.安排id;
    
      For c_记录 In (Select ID As 记录id From 临床出诊记录 Where 安排id = c_安排.安排id And 出诊日期 >= 开始时间_In) Loop
        l_记录id.Extend();
        l_记录id(l_记录id.Count) := c_记录.记录id;
      End Loop;
    End Loop;
    Zl_临床出诊记录_Batchdelete(l_记录id);
    Forall I In 1 .. l_安排id.Count
      Update 临床出诊安排 A
      Set a.终止时间 = 开始时间_In - 1 / 24 / 60 / 60
      Where a.Id = l_安排id(I) And Not Exists (Select 1 From 临床出诊记录 Where 安排id = a.Id And 出诊日期 >= 开始时间_In);
  End If;

  --缺省加入上一次有效的出诊安排
  For c_号源 In (Select 临床出诊安排_Id.Nextval As 安排id, n_出诊id As 出诊id, 原安排id, 号源id, 项目id, 医生id, 医生姓名
               From (Select a.Id As 原安排id, b.Id As 号源id, b.项目id, b.医生id, b.医生姓名,
                             Row_Number() Over(Partition By b.Id Order By a.开始时间 Desc) As 组号
                      From 临床出诊安排 A, 临床出诊号源 B, 临床出诊表 C, 部门表 D
                      Where a.号源id = b.Id And a.出诊id = c.Id And b.科室id = d.Id
                           --有效号源
                            And Nvl(b.是否删除, 0) = 0 And
                            Nvl(b.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                            (
                            --月排班
                             Nvl(操作类型_In, 0) = 3 And b.排班方式 = 1
                            --周排班
                             Or
                             Nvl(操作类型_In, 0) = 4 And
                             (
                             --当前出诊表所在时间范围内不能有月排班
                              b.排班方式 = 2 And Not Exists
                              (Select 1
                               From 临床出诊安排 P, 临床出诊表 Q
                               Where p.出诊id = q.Id And p.号源id = b.Id And
                                     Not (p.终止时间 < Trunc(开始时间_In, 'MONTH') Or p.开始时间 > Last_Day(开始时间_In)) And q.排班方式 = 1)
                             --当前已调整为了月排班,但是本月又用了周排班，则本月剩下的部分将继续按周进行排班
                              Or b.排班方式 = 1 And Exists
                              (Select 1
                               From 临床出诊安排 P, 临床出诊表 Q
                               Where p.出诊id = q.Id And p.号源id = b.Id And
                                     Not (p.终止时间 < Trunc(开始时间_In, 'MONTH') Or p.开始时间 > Last_Day(开始时间_In)) And q.排班方式 = 2)))
                           --上一次有效出诊安排
                            And c.发布人 Is Not Null And c.排班方式 = Decode(Nvl(操作类型_In, 0), 3, 1, 2)
                           --号源在该出诊表时间范围内无出诊记录
                            And Not Exists (Select 1
                             From 临床出诊记录 P
                             Where p.号源id = b.Id And p.出诊日期 Between 开始时间_In And 终止时间_In)
                           --当前人员可操作的号源
                            And (Nvl(人员id_In, 0) = 0 Or
                            (Nvl(b.是否临床排班, 0) = 1 And Exists
                             (Select 1 From 部门人员 Where 部门id = b.科室id And 人员id = 人员id_In)))
                           --站点
                            And (d.站点 Is Null Or d.站点 = 站点_In))
               Where 组号 = 1) Loop
  
    Insert Into 临床出诊安排
      (ID, 出诊id, 号源id, 项目id, 医生id, 医生姓名, 开始时间, 终止时间, 操作员姓名, 登记时间, 原终止时间)
    Values
      (c_号源.安排id, c_号源.出诊id, c_号源.号源id, c_号源.项目id, c_号源.医生id, c_号源.医生姓名, 开始时间_In, 终止时间_In, 操作员_In, 操作时间_In, 终止时间_In);
  
    --复制出诊安排
    For c_记录 In (Select a.Id, b.日期
                 From 临床出诊记录 A,
                      (Select Trunc(开始时间_In) + Level - 1 As 日期
                        From Dual
                        Connect By Level <= Trunc(终止时间_In) - Trunc(开始时间_In) + 1) B
                 Where a.安排id = c_号源.原安排id
                      --月排班
                       And (Nvl(操作类型_In, 0) = 3 And To_Char(a.出诊日期, 'dd') = To_Char(b.日期, 'dd')
                       --周排班
                       Or Nvl(操作类型_In, 0) = 4 And To_Char(a.出诊日期, 'D') = To_Char(b.日期, 'D'))) Loop
      Zl_临床出诊记录_Copy(c_记录.Id, c_号源.安排id, c_记录.日期, 操作员_In, 操作时间_In);
    End Loop;
  End Loop;

  --加入无上一次有效出诊安排的号源
  For c_号源 In (Select 临床出诊安排_Id.Nextval As 安排id, n_出诊id As 出诊id, a.Id As 号源id, a.项目id, a.医生id, a.医生姓名
               From 临床出诊号源 A, 部门表 D
               Where a.科室id = d.Id
                    --有效号源
                     And Nvl(a.是否删除, 0) = 0 And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd')) And
                     (
                     --月排班
                      Nvl(操作类型_In, 0) = 3 And a.排班方式 = 1
                     --周排班
                      Or Nvl(操作类型_In, 0) = 4 And
                      (
                      --当前出诊表所在时间范围内不能有月排班
                       a.排班方式 = 2 And Not Exists
                       (Select 1
                           From 临床出诊安排 P, 临床出诊表 Q
                           Where p.出诊id = q.Id And p.号源id = a.Id And
                                 Not (p.终止时间 < Trunc(开始时间_In, 'MONTH') Or p.开始时间 > Last_Day(开始时间_In)) And q.排班方式 = 1)
                      --当前已调整为了月排班,但是本月又用了周排班，则本月剩下的部分将继续按周进行排班
                       Or a.排班方式 = 1 And Exists
                       (Select 1
                           From 临床出诊安排 P, 临床出诊表 Q
                           Where p.出诊id = q.Id And p.号源id = a.Id And
                                 Not (p.终止时间 < Trunc(开始时间_In, 'MONTH') Or p.开始时间 > Last_Day(开始时间_In)) And q.排班方式 = 2)))
                    --号源在该出诊表时间范围内无出诊记录
                     And Not Exists
                (Select 1
                      From 临床出诊记录 P
                      Where p.号源id = a.Id And p.出诊日期 Between 开始时间_In And 终止时间_In)
                    --当前人员可操作的号源
                     And (Nvl(人员id_In, 0) = 0 Or (Nvl(a.是否临床排班, 0) = 1 And Exists
                      (Select 1 From 部门人员 Where 部门id = a.科室id And 人员id = 人员id_In)))
                    --站点
                     And (d.站点 Is Null Or d.站点 = 站点_In)
                    
                     And Not Exists (Select 1 From 临床出诊安排 Where 出诊id = n_出诊id And 号源id = a.Id)) Loop
  
    Insert Into 临床出诊安排
      (ID, 出诊id, 号源id, 项目id, 医生id, 医生姓名, 开始时间, 终止时间, 操作员姓名, 登记时间, 原终止时间)
    Values
      (c_号源.安排id, c_号源.出诊id, c_号源.号源id, c_号源.项目id, c_号源.医生id, c_号源.医生姓名, 开始时间_In, 终止时间_In, 操作员_In, 操作时间_In, 终止时间_In);
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊表_Add;
/
Create Or Replace Procedure Zl_临床出诊限制_Copy
(
  原限制id_In 临床出诊限制.Id%Type,
  安排id_In   临床出诊限制.安排id%Type
) As
  --复制临床出诊限制
  n_限制id 临床出诊限制.Id%Type;
Begin
  Select 临床出诊限制_Id.Nextval Into n_限制id From Dual;

  Insert Into 临床出诊限制
    (ID, 安排id, 限制项目, 上班时段, 限号数, 限约数, 是否序号控制, 是否分时段, 预约控制, 分诊方式, 诊室id, 是否独占)
    Select n_限制id, 安排id_In, a.限制项目, a.上班时段, a.限号数, a.限约数, a.是否序号控制, a.是否分时段, a.预约控制, a.分诊方式, a.诊室id, a.是否独占
    From 临床出诊限制 A
    Where a.Id = 原限制id_In;

  Insert Into 临床出诊诊室
    (限制id, 诊室id)
    Select n_限制id, 诊室id From 临床出诊诊室 Where 限制id = 原限制id_In;

  Insert Into 临床出诊时段
    (限制id, 序号, 开始时间, 终止时间, 限制数量, 是否预约)
    Select n_限制id, 序号, 开始时间, 终止时间, 限制数量, 是否预约 From 临床出诊时段 Where 限制id = 原限制id_In;

  Insert Into 临床出诊挂号控制
    (限制id, 类型, 性质, 名称, 序号, 控制方式, 数量)
    Select n_限制id, 类型, 性质, 名称, 序号, 控制方式, 数量 From 临床出诊挂号控制 Where 限制id = 原限制id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊限制_Copy;
/
Create Or Replace Procedure Zl_临床出诊记录_Copy
(
  原记录id_In   临床出诊记录.Id%Type,
  安排id_In     临床出诊限制.安排id%Type,
  出诊日期_In   临床出诊记录.出诊日期%Type,
  操作员姓名_In 临床出诊记录.登记人%Type,
  登记时间_In   临床出诊记录.登记时间%Type
) As
  --复制临床出诊记录
  n_记录id 临床出诊记录.Id%Type;

  d_开始时间 临床出诊记录.开始时间%Type;
Begin
  Select 临床出诊记录_Id.Nextval Into n_记录id From Dual;
  Begin
    Select a.开始时间 Into d_开始时间 From 临床出诊记录 A Where a.Id = 原记录id_In;
  Exception
    When Others Then
      Return;
  End;

  Insert Into 临床出诊记录
    (ID, 安排id, 号源id, 出诊日期, 上班时段, 开始时间, 终止时间, 停诊开始时间, 停诊终止时间, 停诊原因, 缺省预约时间, 提前挂号时间, 限号数, 已挂数, 限约数, 已约数, 其中已接收, 是否序号控制,
     是否分时段, 预约控制, 是否独占, 项目id, 科室id, 医生id, 医生姓名, 替诊医生id, 替诊医生姓名, 分诊方式, 诊室id, 是否锁定, 是否临时出诊, 登记人, 登记时间)
    Select n_记录id, 安排id_In, a.号源id, 出诊日期_In, a.上班时段,
           To_Date(To_Char(出诊日期_In, 'yyyy-mm-dd ') || To_Char(a.开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
           To_Date(To_Char(出诊日期_In, 'yyyy-mm-dd ') || To_Char(a.终止时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') + Case
             When Trunc(终止时间) > Trunc(d_开始时间) Then
              1
             Else
              0
           End, Null As 停诊开始时间, Null As 停诊终止时间, Null As 停诊原因,
           To_Date(To_Char(出诊日期_In, 'yyyy-mm-dd ') || To_Char(a.缺省预约时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') + Case
             When Trunc(缺省预约时间) > Trunc(d_开始时间) Then
              1
             Else
              0
           End,
           To_Date(To_Char(出诊日期_In, 'yyyy-mm-dd ') || To_Char(a.提前挂号时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') + Case
             When Trunc(提前挂号时间) < Trunc(d_开始时间) Then
              -1
             Else
              0
           End, a.限号数, 0 As 已挂数, a.限约数, 0 As 已约数, 0 As 其中已接收, a.是否序号控制, a.是否分时段, a.预约控制, a.是否独占, a.项目id, a.科室id, a.医生id,
           a.医生姓名, Null As 替诊医生id, Null As 替诊医生姓名, a.分诊方式, a.诊室id, 0 As 是否锁定, 0 As 是否临时出诊, 操作员姓名_In, 登记时间_In
    From 临床出诊记录 A
    Where a.Id = 原记录id_In;

  Insert Into 临床出诊诊室记录
    (记录id, 诊室id)
    Select n_记录id, 诊室id From 临床出诊诊室记录 Where 记录id = 原记录id_In;

  --分时段不分序号的，在预约挂号时会新增记录，填写预约顺序号
  Insert Into 临床出诊序号控制
    (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约)
    Select n_记录id, 序号,
           To_Date(To_Char(出诊日期_In, 'yyyy-mm-dd ') || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') + Case
             When Trunc(开始时间) > Trunc(d_开始时间) Then
              1
             Else
              0
           End,
           To_Date(To_Char(出诊日期_In, 'yyyy-mm-dd ') || To_Char(终止时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') + Case
             When Trunc(终止时间) > Trunc(d_开始时间) Then
              1
             Else
              0
           End, 数量, 是否预约
    From 临床出诊序号控制
    Where 预约顺序号 Is Null And 记录id = 原记录id_In;

  Insert Into 临床出诊挂号控制记录
    (记录id, 类型, 性质, 名称, 序号, 控制方式, 数量)
    Select n_记录id, 类型, 性质, 名称, 序号, 控制方式, 数量 From 临床出诊挂号控制记录 Where 记录id = 原记录id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊记录_Copy;
/
Create Or Replace Procedure Zl_临床出诊记录_Batchdelete(记录id_In t_Numlist) As
  --删除临床出诊记录
Begin
  Forall I In 1 .. 记录id_In.Count
    Delete From 临床出诊变动明细 Where 变动id In (Select ID From 临床出诊变动记录 Where 记录id = 记录id_In(I));

  Forall I In 1 .. 记录id_In.Count
    Delete From 临床出诊变动记录 Where 记录id = 记录id_In(I);

  Forall I In 1 .. 记录id_In.Count
    Delete From 临床出诊停诊记录 Where 记录id = 记录id_In(I);

  Forall I In 1 .. 记录id_In.Count
    Delete From 临床出诊序号控制 Where 记录id = 记录id_In(I);

  Forall I In 1 .. 记录id_In.Count
    Delete From 临床出诊诊室记录 Where 记录id = 记录id_In(I);

  Forall I In 1 .. 记录id_In.Count
    Delete From 临床出诊挂号控制记录 Where 记录id = 记录id_In(I);

  Forall I In 1 .. 记录id_In.Count
    Delete From 临床出诊记录 Where ID = 记录id_In(I);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊记录_Batchdelete;
/
Create Or Replace Procedure Zl_Buildregisterfixedrule
(
  Id_In         临床出诊表.Id%Type,
  Newid_In      临床出诊表.Id%Type,
  出诊表名_In   临床出诊表.出诊表名%Type,
  开始时间_In   临床出诊安排.开始时间%Type,
  终止时间_In   临床出诊安排.终止时间%Type,
  操作员姓名_In 临床出诊安排.操作员姓名%Type := Null,
  登记时间_In   临床出诊安排.登记时间%Type := Null,
  站点_In       部门表.站点%Type
) As
  -------------------------------------------------------------------------
  --功能：根据现有固定出诊表规则生成成新的固定出诊表
  -------------------------------------------------------------------------
  n_Count Number;

  n_出诊id 临床出诊表.Id%Type;

  v_操作员   临床出诊安排.操作员姓名%Type;
  d_登记时间 Date;
  v_Err_Msg  Varchar2(255);
  Err_Item Exception;
Begin
  Begin
    Select 1 Into n_Count From 临床出诊表 Where ID = Id_In;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count = 0 Then
    v_Err_Msg := '未发现原出诊表信息！';
    Raise Err_Item;
  End If;

  --检查是否有有效号源
  Begin
    Select 1
    Into n_Count
    From 临床出诊号源 A, 部门表 B
    Where a.科室id = b.Id And a.排班方式 = 0 And Nvl(a.是否删除, 0) = 0 And
          (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'))
         --站点
          And (b.站点 Is Null Or b.站点 = 站点_In) And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If Nvl(n_Count, 0) = 0 Then
    v_Err_Msg := '当前出诊表中已无可按固定排班的号源，不能生成新的固定安排！';
    Raise Err_Item;
  End If;

  Begin
    Select 1
    Into n_Count
    From 临床出诊安排 A, 临床出诊表 B
    Where a.出诊id = b.Id And b.排班方式 = 0 And a.开始时间 = 开始时间_In And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If Nvl(n_Count, 0) <> 0 Then
    v_Err_Msg := '已存在为当前开始时间的固定安排！';
    Raise Err_Item;
  End If;

  n_出诊id := Newid_In;
  If Nvl(n_出诊id, 0) = 0 Then
    Select 临床出诊表_Id.Nextval Into n_出诊id From Dual;
  End If;

  Insert Into 临床出诊表
    (ID, 排班方式, 出诊表名, 年份)
  Values
    (n_出诊id, 0, 出诊表名_In, To_Number(To_Char(开始时间_In, 'yyyy')));

  d_登记时间 := Nvl(登记时间_In, Sysdate);
  v_操作员   := Nvl(操作员姓名_In, Zl_Username);

  For c_号源 In (Select 临床出诊安排_Id.Nextval As 安排id, n_出诊id As 出诊id, 原安排id, 号源id, 项目id, 医生id, 医生姓名
               From (Select b.Id As 原安排id, b.号源id, c.项目id, c.医生id, c.医生姓名,
                             Row_Number() Over(Partition By c.Id Order By b.开始时间 Desc) As 组号
                      From 临床出诊安排 B, 临床出诊号源 C, 部门表 D
                      Where b.号源id = c.Id And c.科室id = d.Id And b.出诊id = Id_In
                           --号源限制
                            And c.排班方式 = 0 And Nvl(c.是否删除, 0) = 0 And
                            (c.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd') Or c.撤档时间 Is Null)
                           --站点
                            And (d.站点 Is Null Or d.站点 = 站点_In)) M
               Where 组号 = 1) Loop
  
    Insert Into 临床出诊安排
      (ID, 出诊id, 号源id, 项目id, 医生id, 医生姓名, 开始时间, 终止时间, 操作员姓名, 登记时间, 原终止时间)
    Values
      (c_号源.安排id, c_号源.出诊id, c_号源.号源id, c_号源.项目id, c_号源.医生id, c_号源.医生姓名, 开始时间_In, 终止时间_In, v_操作员, d_登记时间, 终止时间_In);
  
    --出诊限制
    For c_限制 In (Select ID, 安排id, 限制项目, 上班时段, 限号数, 限约数, 是否序号控制, 是否分时段, 预约控制, 分诊方式, 诊室id, 是否独占
                 From 临床出诊限制
                 Where 安排id = c_号源.原安排id) Loop
    
      Zl_临床出诊限制_Copy(c_限制.Id, c_号源.安排id);
    End Loop;
  End Loop;

  --加入没有的出诊安排的号源
  For c_号源 In (Select 临床出诊安排_Id.Nextval As 安排id, n_出诊id As 出诊id, a.Id As 号源id, a.项目id, a.医生id, a.医生姓名
               From 临床出诊号源 A, 部门表 D
               Where a.科室id = d.Id And a.排班方式 = 0 And Nvl(a.是否删除, 0) = 0 And
                     (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'))
                    --站点
                     And (d.站点 Is Null Or d.站点 = 站点_In)
                    
                     And Not Exists (Select 1 From 临床出诊安排 Where 出诊id = n_出诊id And 号源id = a.Id)) Loop
  
    Insert Into 临床出诊安排
      (ID, 出诊id, 号源id, 项目id, 医生id, 医生姓名, 开始时间, 终止时间, 操作员姓名, 登记时间, 原终止时间)
    Values
      (c_号源.安排id, c_号源.出诊id, c_号源.号源id, c_号源.项目id, c_号源.医生id, c_号源.医生姓名, 开始时间_In, 终止时间_In, v_操作员, d_登记时间, 终止时间_In);
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Buildregisterfixedrule;
/
Create Or Replace Procedure Zl_Buildregisterplanbyrecord
(
  原出诊id_In   临床出诊表.Id%Type,
  新出诊id_In   临床出诊表.Id%Type,
  排班方式_In   临床出诊表.排班方式%Type,
  出诊表名_In   临床出诊表.出诊表名%Type,
  年份_In       临床出诊表.年份%Type,
  月份_In       临床出诊表.月份%Type,
  周数_In       临床出诊表.周数%Type,
  开始时间_In   临床出诊安排.开始时间%Type,
  终止时间_In   临床出诊安排.终止时间%Type,
  操作员姓名_In 临床出诊安排.操作员姓名%Type,
  登记时间_In   临床出诊安排.登记时间%Type,
  站点_In       部门表.站点%Type,
  人员id_In     人员表.Id%Type := Null,
  删除安排_In   Number := 0
) As
  -------------------------------------------------------------------------
  --功能：根据出诊记录生成新的出诊记录（月安排/周安排）
  --参数：
  --        人员id_In 除固定安排外有效，不为0或null表示临床科室人员在添加
  --        删除安排_In 固定排班转为月排班/周排班时，在制定月排班/周排班时是否删除新出诊表时间内未使用的出诊记录
  --说明：
  -------------------------------------------------------------------------
  n_Count Number;

  l_记录id t_Numlist := t_Numlist();
  l_安排id t_Numlist := t_Numlist();

  v_Err_Msg Varchar2(255);
  Err_Item Exception;
Begin
  Begin
    Select 1
    Into n_Count
    From 临床出诊号源 A, 部门表 B
    Where a.科室id = b.Id
         --有效号源
          And Nvl(a.是否删除, 0) = 0 And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd')) And
          (
          --月排班
           Nvl(排班方式_In, 0) = 1 And a.排班方式 = 1
          --周排班
           Or Nvl(排班方式_In, 0) = 2 And
           (
           --当前出诊表所在时间范围内不能有月排班
            a.排班方式 = 2 And Not Exists
            (Select 1
                From 临床出诊安排 P, 临床出诊表 Q
                Where p.出诊id = q.Id And p.号源id = a.Id And
                      Not (p.终止时间 < Trunc(开始时间_In, 'MONTH') Or p.开始时间 > Last_Day(开始时间_In)) And q.排班方式 = 1)
           --当前已调整为了月排班,但是本月又用了周排班，则本月剩下的部分将继续按周进行排班
            Or a.排班方式 = 1 And Exists
            (Select 1
                From 临床出诊安排 P, 临床出诊表 Q
                Where p.出诊id = q.Id And p.号源id = a.Id And
                      Not (p.终止时间 < Trunc(开始时间_In, 'MONTH') Or p.开始时间 > Last_Day(开始时间_In)) And q.排班方式 = 2)))
         --号源在该出诊表时间范围内无出诊记录
          And Not Exists
     (Select 1
           From 临床出诊记录 O, 临床出诊安排 P, 临床出诊表 Q
           Where o.安排id = p.Id And p.出诊id = q.Id And p.号源id = a.Id And o.出诊日期 Between 开始时间_In And 终止时间_In And
                 (q.排班方式 In (1, 2)
                 --原来为固定出诊安排
                 Or q.排班方式 = 0 And (Nvl(删除安排_In, 0) = 0 Or Nvl(删除安排_In, 0) = 1 And Exists
                  (Select 1 From 病人挂号记录 Where 出诊记录id = a.Id))))
         --当前人员可操作的号源
          And (Nvl(人员id_In, 0) = 0 Or
          (Nvl(a.是否临床排班, 0) = 1 And Exists (Select 1 From 部门人员 Where 部门id = a.科室id And 人员id = 人员id_In)))
         --站点
          And (b.站点 Is Null Or b.站点 = 站点_In) And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count = 0 Then
    If Nvl(排班方式_In, 0) = 1 Then
      v_Err_Msg := '当前出诊表中已无可按月排班的号源，不能生成新的出诊表！';
    Else
      v_Err_Msg := '当前出诊表中已无可按周排班的号源，不能生成新的出诊表！';
    End If;
    Raise Err_Item;
  End If;

  --检查出诊表是否存在
  Begin
    Select 1 Into n_Count From 临床出诊表 Where ID = 新出诊id_In;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If Nvl(n_Count, 0) = 0 Then
    Insert Into 临床出诊表
      (ID, 排班方式, 出诊表名, 年份, 月份, 周数)
    Values
      (新出诊id_In, 排班方式_In, 出诊表名_In, 年份_In, 月份_In, 周数_In);
  End If;

  --如果当前出诊表时间范围内无挂号且无预约的出诊记录(固定安排)，则删除这部分出诊记录(在删除出诊表时可恢复)，
  --并修改固定安排的终止时间，程序中已询问
  If Nvl(删除安排_In, 0) = 1 Then
    For c_安排 In (Select b.Id As 安排id
                 From 临床出诊安排 B, 临床出诊表 C, 临床出诊号源 D
                 Where b.出诊id = c.Id And b.号源id = d.Id
                      --号源
                       And Nvl(d.是否删除, 0) = 0 And (d.撤档时间 Is Null Or d.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd')) And
                       Nvl(d.排班方式, 0) = 排班方式_In
                      --安排有被使用了的出诊记录
                       And c.排班方式 = 0 And b.终止时间 >= 开始时间_In And Not Exists
                  (Select 1
                        From 临床出诊记录 M, 病人挂号记录 N
                        Where m.安排id = b.Id And m.Id = n.出诊记录id And m.出诊日期 >= 开始时间_In)
                      --当前人员可操作的号源
                       And (Nvl(人员id_In, 0) = 0 Or (Nvl(d.是否临床排班, 0) = 1 And Exists
                        (Select 1 From 部门人员 Where 部门id = d.科室id And 人员id = 人员id_In)))) Loop
      l_安排id.Extend();
      l_安排id(l_安排id.Count) := c_安排.安排id;
    
      For c_记录 In (Select ID As 记录id From 临床出诊记录 Where 安排id = c_安排.安排id And 出诊日期 >= 开始时间_In) Loop
        l_记录id.Extend();
        l_记录id(l_记录id.Count) := c_记录.记录id;
      End Loop;
    End Loop;
  
    Zl_临床出诊记录_Batchdelete(l_记录id);
    Forall I In 1 .. l_安排id.Count
      Update 临床出诊安排 A
      Set a.终止时间 = 开始时间_In - 1 / 24 / 60 / 60
      Where a.Id = l_安排id(I) And Not Exists (Select 1 From 临床出诊记录 Where 安排id = a.Id And 出诊日期 >= 开始时间_In);
  End If;

  For c_号源 In (Select 临床出诊安排_Id.Nextval As 安排id, 新出诊id_In As 出诊id, b.Id As 原安排id, b.号源id, c.项目id, c.医生id, c.医生姓名
               From 临床出诊安排 B, 临床出诊号源 C, 部门表 D
               Where b.号源id = c.Id And c.科室id = d.Id And b.出诊id = 原出诊id_In
                    --有效号源
                     And Nvl(c.是否删除, 0) = 0 And (c.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd') Or c.撤档时间 Is Null) And
                     (
                     --月排班
                      Nvl(排班方式_In, 0) = 1 And c.排班方式 = 1
                     -- 周排班
                      Or Nvl(排班方式_In, 0) = 2 And
                      (
                      --当前出诊表所在时间范围内不能有月排班
                       c.排班方式 = 2 And Not Exists
                       (Select 1
                           From 临床出诊安排 P, 临床出诊表 Q
                           Where p.出诊id = q.Id And p.号源id = c.Id And
                                 Not (p.终止时间 < Trunc(开始时间_In, 'MONTH') Or p.开始时间 > Last_Day(开始时间_In)) And q.排班方式 = 1)
                      --当前已调整为了月排班,但是本月又用了周排班，则本月剩下的部分将继续按周进行排班
                       Or c.排班方式 = 1 And Exists
                       (Select 1
                           From 临床出诊安排 P, 临床出诊表 Q
                           Where p.出诊id = q.Id And p.号源id = c.Id And
                                 Not (p.终止时间 < Trunc(开始时间_In, 'MONTH') Or p.开始时间 > Last_Day(开始时间_In)) And q.排班方式 = 2)))
                    --号源在该出诊表时间范围内无出诊记录
                     And Not Exists
                (Select 1
                      From 临床出诊记录 P
                      Where p.号源id = c.Id And p.出诊日期 Between 开始时间_In And 终止时间_In)
                    --当前人员可操作的号源
                     And (Nvl(人员id_In, 0) = 0 Or (Nvl(c.是否临床排班, 0) = 1 And Exists
                      (Select 1 From 部门人员 Where 部门id = c.科室id And 人员id = 人员id_In)))
                    --站点
                     And (d.站点 Is Null Or d.站点 = 站点_In)) Loop
  
    Insert Into 临床出诊安排
      (ID, 出诊id, 号源id, 项目id, 医生id, 医生姓名, 开始时间, 终止时间, 操作员姓名, 登记时间, 原终止时间)
    Values
      (c_号源.安排id, c_号源.出诊id, c_号源.号源id, c_号源.项目id, c_号源.医生id, c_号源.医生姓名, 开始时间_In, 终止时间_In, 操作员姓名_In, 登记时间_In, 终止时间_In);
  
    --出诊记录
    For c_记录 In (Select a.Id, b.日期
                 From 临床出诊记录 A,
                      (Select Trunc(开始时间_In) + Level - 1 As 日期
                        From Dual
                        Connect By Level <= Trunc(终止时间_In) - Trunc(开始时间_In) + 1) B
                 Where a.安排id = c_号源.原安排id
                      --月排班
                       And (Nvl(排班方式_In, 0) = 1 And To_Char(a.出诊日期, 'dd') = To_Char(b.日期, 'dd')
                       --周排班
                       Or Nvl(排班方式_In, 0) = 2 And To_Char(a.出诊日期, 'D') = To_Char(b.日期, 'D'))) Loop
    
      Zl_临床出诊记录_Copy(c_记录.Id, c_号源.安排id, c_记录.日期, 操作员姓名_In, 登记时间_In);
    End Loop;
  End Loop;

  --加入没有的出诊安排的号源
  For c_号源 In (Select 临床出诊安排_Id.Nextval As 安排id, 新出诊id_In As 出诊id, a.Id As 号源id, a.项目id, a.医生id, a.医生姓名
               From 临床出诊号源 A, 部门表 D
               Where a.科室id = d.Id
                    --有效号源
                     And Nvl(a.是否删除, 0) = 0 And (a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd') Or a.撤档时间 Is Null) And
                     (
                     --月排班
                      Nvl(排班方式_In, 0) = 1 And a.排班方式 = 1
                     -- 周排班
                      Or Nvl(排班方式_In, 0) = 2 And
                      (
                      --当前出诊表所在时间范围内不能有月排班
                       a.排班方式 = 2 And Not Exists
                       (Select 1
                           From 临床出诊安排 P, 临床出诊表 Q
                           Where p.出诊id = q.Id And p.号源id = a.Id And
                                 Not (p.终止时间 < Trunc(开始时间_In, 'MONTH') Or p.开始时间 > Last_Day(开始时间_In)) And q.排班方式 = 1)
                      --当前已调整为了月排班,但是本月又用了周排班，则本月剩下的部分将继续按周进行排班
                       Or a.排班方式 = 1 And Exists
                       (Select 1
                           From 临床出诊安排 P, 临床出诊表 Q
                           Where p.出诊id = q.Id And p.号源id = a.Id And
                                 Not (p.终止时间 < Trunc(开始时间_In, 'MONTH') Or p.开始时间 > Last_Day(开始时间_In)) And q.排班方式 = 2)))
                    --号源在该出诊表时间范围内无出诊记录
                     And Not Exists
                (Select 1
                      From 临床出诊记录 P
                      Where p.号源id = a.Id And p.出诊日期 Between 开始时间_In And 终止时间_In)
                    --当前人员可操作的号源
                     And (Nvl(人员id_In, 0) = 0 Or (Nvl(a.是否临床排班, 0) = 1 And Exists
                      (Select 1 From 部门人员 Where 部门id = a.科室id And 人员id = 人员id_In)))
                    --站点
                     And (d.站点 Is Null Or d.站点 = 站点_In)
                    
                     And Not Exists (Select 1 From 临床出诊安排 Where 出诊id = 新出诊id_In And 号源id = a.Id)) Loop
  
    Insert Into 临床出诊安排
      (ID, 出诊id, 号源id, 项目id, 医生id, 医生姓名, 开始时间, 终止时间, 操作员姓名, 登记时间, 原终止时间)
    Values
      (c_号源.安排id, c_号源.出诊id, c_号源.号源id, c_号源.项目id, c_号源.医生id, c_号源.医生姓名, 开始时间_In, 终止时间_In, 操作员姓名_In, 登记时间_In, 终止时间_In);
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Buildregisterplanbyrecord;
/
Create Or Replace Procedure Zl_Buildregisterplanbytemplet
(
  模板id_In   临床出诊表.Id%Type,
  人员id_In   人员表.Id%Type,
  出诊id_In   临床出诊表.Id%Type,
  排班方式_In 临床出诊表.排班方式%Type,
  出诊表名_In 临床出诊表.出诊表名%Type,
  年份_In     临床出诊表.年份%Type,
  月份_In     临床出诊表.月份%Type,
  周数_In     临床出诊表.周数%Type,
  开始时间_In 临床出诊安排.开始时间%Type,
  终止时间_In 临床出诊安排.终止时间%Type,
  操作员_In   临床出诊安排.操作员姓名%Type,
  登记时间_In 临床出诊安排.登记时间%Type,
  站点_In     部门表.站点%Type,
  删除安排_In Number := 0
) As
  -------------------------------------------------------------------------
  --功能说明：根据模板自动生成临床出诊记录
  --参数：
  --        人员id_In 除固定安排外有效，不为0或null表示临床科室人员在添加
  --        删除安排_In 固定排班转为月排班/周排班时，在制定月排班/周排班时是否删除新出诊表时间内未使用的出诊记录
  --说明：
  -------------------------------------------------------------------------
  Err_Item Exception;
  v_Err_Msg Varchar2(200);
  n_Count   Number(18);

  d_轮询日期 Date;
  n_轮询天数 Number;
  v_限制项目 临床出诊限制.限制项目%Type;

  n_是否出诊 Number(2);

  l_记录id t_Numlist := t_Numlist();
  l_安排id t_Numlist := t_Numlist();

  Procedure Isvisit
  (
    安排id_In       临床出诊安排.Id%Type,
    排班规则_In     临床出诊安排.排班规则%Type,
    出诊日期_In     临床出诊记录.出诊日期%Type,
    轮询开始时间_In 临床出诊安排.开始时间%Type,
    限制项目_In     Out 临床出诊限制.限制项目%Type,
    是否出诊_In     Out Number
  ) As
    --判断是否出诊，并获取出诊项目
    d_轮询日期 Date;
    n_轮询天数 Number;
  Begin
    是否出诊_In := 1;
    --检查这天是否出诊
    If 排班规则_In = 1 Then
      --星期排班
      Select Decode(To_Char(出诊日期_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六',
                     Null)
      Into 限制项目_In
      From Dual;
      Select Count(1) Into n_Count From 临床出诊限制 Where 安排id = 安排id_In And 限制项目 = 限制项目_In;
      If Nvl(n_Count, 0) = 0 Then
        是否出诊_In := 0;
      End If;
    Elsif 排班规则_In = 2 Then
      --单日排班
      限制项目_In := '单日';
      If Mod(To_Number(To_Char(出诊日期_In, 'dd')), 2) <> 1 Then
        是否出诊_In := 0;
      End If;
    Elsif 排班规则_In = 3 Then
      --双日排班
      限制项目_In := '双日';
      If Mod(To_Number(To_Char(出诊日期_In, 'dd')), 2) <> 0 Then
        是否出诊_In := 0;
      End If;
    Elsif 排班规则_In = 4 Or 排班规则_In = 5 Then
      --4-月内轮循,5-轮循不限制
      If 排班规则_In = 4 Then
        d_轮询日期 := To_Date(To_Char(出诊日期_In, 'yyyy-mm') || To_Char(轮询开始时间_In, '-dd'), 'yyyy-mm-dd');
      Else
        d_轮询日期 := 轮询开始时间_In;
      End If;
      Begin
        Select To_Number(Substr(限制项目, 1, Instr(限制项目, '天') - 1))
        Into n_轮询天数
        From 临床出诊限制
        Where 安排id = 安排id_In And Rownum < 2;
      Exception
        When Others Then
          n_轮询天数 := 0;
      End;
      If Nvl(n_轮询天数, 0) > 0 Then
        限制项目_In := n_轮询天数 || '天';
        If Mod(Trunc(出诊日期_In) - Trunc(d_轮询日期), n_轮询天数 + 1) <> 0 Then
          是否出诊_In := 0;
        End If;
      End If;
    Elsif 排班规则_In = 6 Then
      --特定日期
      限制项目_In := To_Number(To_Char(出诊日期_In, 'dd')) || '日';
      Select Count(1) Into n_Count From 临床出诊限制 Where 安排id = 安排id_In And 限制项目 = 限制项目_In;
      If Nvl(n_Count, 0) = 0 Then
        是否出诊_In := 0;
      End If;
    End If;
  End;
Begin
  Begin
    Select 1
    Into n_Count
    From 临床出诊号源 A, 部门表 B
    Where a.科室id = b.Id
         --有效号源
          And Nvl(a.是否删除, 0) = 0 And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd')) And
          (
          --月排班
           Nvl(排班方式_In, 0) = 1 And a.排班方式 = 1
          --周排班
           Or Nvl(排班方式_In, 0) = 2 And
           (
           --当前出诊表所在时间范围内不能有月排班
            a.排班方式 = 2 And Not Exists
            (Select 1
                From 临床出诊安排 P, 临床出诊表 Q
                Where p.出诊id = q.Id And p.号源id = a.Id And
                      Not (p.终止时间 < Trunc(开始时间_In, 'MONTH') Or p.开始时间 > Last_Day(开始时间_In)) And q.排班方式 = 1)
           --当前已调整为了月排班,但是本月又用了周排班，则本月剩下的部分将继续按周进行排班
            Or a.排班方式 = 1 And Exists
            (Select 1
                From 临床出诊安排 P, 临床出诊表 Q
                Where p.出诊id = q.Id And p.号源id = a.Id And
                      Not (p.终止时间 < Trunc(开始时间_In, 'MONTH') Or p.开始时间 > Last_Day(开始时间_In)) And q.排班方式 = 2)))
         --号源在该出诊表时间范围内无出诊记录
          And Not Exists
     (Select 1
           From 临床出诊记录 O, 临床出诊安排 P, 临床出诊表 Q
           Where o.安排id = p.Id And p.出诊id = q.Id And p.号源id = a.Id And o.出诊日期 Between 开始时间_In And 终止时间_In And
                 (q.排班方式 In (1, 2)
                 --原来为固定出诊安排
                 Or q.排班方式 = 0 And (Nvl(删除安排_In, 0) = 0 Or Nvl(删除安排_In, 0) = 1 And Exists
                  (Select 1 From 病人挂号记录 Where 出诊记录id = a.Id))))
         --当前人员可操作的号源
          And (Nvl(人员id_In, 0) = 0 Or
          (Nvl(a.是否临床排班, 0) = 1 And Exists (Select 1 From 部门人员 Where 部门id = a.科室id And 人员id = 人员id_In)))
         --站点
          And (b.站点 Is Null Or b.站点 = 站点_In) And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count = 0 Then
    If Nvl(排班方式_In, 0) = 1 Then
      v_Err_Msg := '当前出诊表中已无可按月排班的号源，不能生成新的出诊表！';
    Else
      v_Err_Msg := '当前出诊表中已无可按周排班的号源，不能生成新的出诊表！';
    End If;
    Raise Err_Item;
  End If;

  --检查出诊表是否存在
  Begin
    Select 1 Into n_Count From 临床出诊表 Where ID = 出诊id_In;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If Nvl(n_Count, 0) = 0 Then
    Insert Into 临床出诊表
      (ID, 排班方式, 出诊表名, 年份, 月份, 周数)
    Values
      (出诊id_In, 排班方式_In, 出诊表名_In, 年份_In, 月份_In, 周数_In);
  End If;

  --如果当前出诊表时间范围内无挂号且无预约的出诊记录(固定安排)，则删除这部分出诊记录(在删除出诊表时可恢复)，
  --并修改固定安排的终止时间，程序中已询问
  If Nvl(删除安排_In, 0) = 1 Then
    For c_安排 In (Select b.Id As 安排id
                 From 临床出诊安排 B, 临床出诊表 C, 临床出诊号源 D
                 Where b.出诊id = c.Id And b.号源id = d.Id
                      --号源
                       And Nvl(d.是否删除, 0) = 0 And (d.撤档时间 Is Null Or d.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd')) And
                       Nvl(d.排班方式, 0) = 排班方式_In
                      --安排有被使用了的出诊记录
                       And c.排班方式 = 0 And b.终止时间 >= 开始时间_In And Not Exists
                  (Select 1
                        From 临床出诊记录 M, 病人挂号记录 N
                        Where m.安排id = b.Id And m.Id = n.出诊记录id And m.出诊日期 >= 开始时间_In)
                      --当前人员可操作的号源
                       And (Nvl(人员id_In, 0) = 0 Or (Nvl(d.是否临床排班, 0) = 1 And Exists
                        (Select 1 From 部门人员 Where 部门id = d.科室id And 人员id = 人员id_In)))) Loop
      l_安排id.Extend();
      l_安排id(l_安排id.Count) := c_安排.安排id;
    
      For c_记录 In (Select ID As 记录id From 临床出诊记录 Where 安排id = c_安排.安排id And 出诊日期 >= 开始时间_In) Loop
        l_记录id.Extend();
        l_记录id(l_记录id.Count) := c_记录.记录id;
      End Loop;
    End Loop;
  
    Zl_临床出诊记录_Batchdelete(l_记录id);
    Forall I In 1 .. l_安排id.Count
      Update 临床出诊安排 A
      Set a.终止时间 = 开始时间_In - 1 / 24 / 60 / 60
      Where a.Id = l_安排id(I) And Not Exists (Select 1 From 临床出诊记录 Where 安排id = a.Id And 出诊日期 >= 开始时间_In);
  End If;

  For c_号源 In (Select 临床出诊安排_Id.Nextval As 安排id, 出诊id_In As 出诊id, b.Id As 原安排id, b.号源id, c.科室id, c.项目id, c.医生id, c.医生姓名,
                      b.排班规则, b.是否周六出诊, b.是否周日出诊, b.开始时间, c.号类, Nvl(d.站点, '-') As 站点
               From 临床出诊安排 B, 临床出诊号源 C, 部门表 D
               Where b.号源id = c.Id And c.科室id = d.Id And b.出诊id = 模板id_In
                    --有效号源
                     And Nvl(c.是否删除, 0) = 0 And (c.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd') Or c.撤档时间 Is Null) And
                     (
                     --月排班
                      Nvl(排班方式_In, 0) = 1 And c.排班方式 = 1
                     -- 周排班
                      Or Nvl(排班方式_In, 0) = 2 And
                      (
                      --当前出诊表所在时间范围内不能有月排班
                       c.排班方式 = 2 And Not Exists
                       (Select 1
                           From 临床出诊安排 P, 临床出诊表 Q
                           Where p.出诊id = q.Id And p.号源id = c.Id And
                                 Not (p.终止时间 < Trunc(开始时间_In, 'MONTH') Or p.开始时间 > Last_Day(开始时间_In)) And q.排班方式 = 1)
                      --当前已调整为了月排班,但是本月又用了周排班，则本月剩下的部分将继续按周进行排班
                       Or c.排班方式 = 1 And Exists
                       (Select 1
                           From 临床出诊安排 P, 临床出诊表 Q
                           Where p.出诊id = q.Id And p.号源id = c.Id And
                                 Not (p.终止时间 < Trunc(开始时间_In, 'MONTH') Or p.开始时间 > Last_Day(开始时间_In)) And q.排班方式 = 2)))
                    --号源在该出诊表时间范围内无出诊记录
                     And Not Exists
                (Select 1
                      From 临床出诊记录 P
                      Where p.号源id = c.Id And p.出诊日期 Between 开始时间_In And 终止时间_In)
                    --当前人员可操作的号源
                     And (Nvl(人员id_In, 0) = 0 Or (Nvl(c.是否临床排班, 0) = 1 And Exists
                      (Select 1 From 部门人员 Where 部门id = c.科室id And 人员id = 人员id_In)))
                    --站点
                     And (d.站点 Is Null Or d.站点 = 站点_In)) Loop
  
    Insert Into 临床出诊安排
      (ID, 出诊id, 号源id, 项目id, 医生id, 医生姓名, 开始时间, 终止时间, 操作员姓名, 登记时间, 原终止时间)
    Values
      (c_号源.安排id, c_号源.出诊id, c_号源.号源id, c_号源.项目id, c_号源.医生id, c_号源.医生姓名, 开始时间_In, 终止时间_In, 操作员_In, 登记时间_In, 终止时间_In);
  
    --临床出诊记录
    For c_日期 In (Select Trunc(开始时间_In) + Level - 1 As 日期,
                        Decode(To_Char(Trunc(开始时间_In) + Level - 1, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                '周四', '6', '周五', '7', '周六', Null) As 星期
                 From Dual
                 Connect By Level <= Trunc(终止时间_In) - Trunc(开始时间_In) + 1) Loop
    
      Isvisit(c_号源.原安排id, c_号源.排班规则, c_日期.日期, c_号源.开始时间, v_限制项目, n_是否出诊);
    
      --是否周六、周日不出诊
      --排班规则:1-星期排班;2-单日排班;3-双日排班;4-月内轮循;5-轮循不限制;6-特定日期
      If Instr(',2,3,4,5,', c_号源.排班规则) > 0 And
         (Nvl(c_号源.是否周六出诊, 0) = 0 And c_日期.星期 = '周六' Or Nvl(c_号源.是否周日出诊, 0) = 0 And c_日期.星期 = '周日') Then
        n_是否出诊 := 0;
      End If;
    
      If Nvl(n_是否出诊, 0) = 1 Then
        For c_记录 In (With c_时间段 As
                        (Select 时间段, 开始时间, 终止时间, 号类, 站点, 缺省时间, 提前时间
                        From (Select 时间段, 开始时间, 终止时间, 号类, 站点, 缺省时间, 提前时间,
                                      Row_Number() Over(Partition By 时间段 Order By 时间段, 站点 Asc, 号类 Asc) As 组号
                               From 时间段
                               Where Nvl(站点, c_号源.站点) = c_号源.站点 And Nvl(号类, c_号源.号类) = c_号源.号类)
                        Where 组号 = 1)
                       Select 临床出诊记录_Id.Nextval As 记录id, m.Id As 限制id, m.上班时段,
                              To_Date(To_Char(c_日期.日期, 'yyyy-mm-dd ') || To_Char(j.开始时间, 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                              To_Date(To_Char(c_日期.日期, 'yyyy-mm-dd ') || To_Char(j.终止时间, 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') + Case
                                 When j.终止时间 <= j.开始时间 Then
                                  1
                                 Else
                                  0
                               End As 终止时间,
                              To_Date(To_Char(c_日期.日期, 'yyyy-mm-dd ') || To_Char(Nvl(j.缺省时间, j.开始时间), 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') + Case
                                 When j.缺省时间 < j.开始时间 Then
                                  1
                                 Else
                                  0
                               End As 缺省预约时间,
                              To_Date(To_Char(c_日期.日期, 'yyyy-mm-dd ') || To_Char(Nvl(j.提前时间, j.开始时间), 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') + Case
                                 When j.开始时间 < j.提前时间 Then
                                  -1
                                 Else
                                  0
                               End As 提前挂号时间, m.限号数, m.限约数, m.是否序号控制, m.是否分时段, m.预约控制, a.项目id, a.医生id, a.医生姓名, m.分诊方式,
                              m.诊室id, m.是否独占
                       From 临床出诊安排 A, 临床出诊限制 M, c_时间段 J
                       Where a.Id = m.安排id And m.上班时段 = j.时间段 And a.Id = c_号源.原安排id And m.限制项目 = v_限制项目) Loop
        
          Insert Into 临床出诊记录
            (ID, 安排id, 号源id, 出诊日期, 上班时段, 开始时间, 终止时间, 缺省预约时间, 提前挂号时间, 限号数, 限约数, 是否序号控制, 是否分时段, 预约控制, 项目id, 科室id, 医生id,
             医生姓名, 分诊方式, 诊室id, 登记人, 登记时间, 是否独占)
          Values
            (c_记录.记录id, c_号源.安排id, c_号源.号源id, c_日期.日期, c_记录.上班时段, c_记录.开始时间, c_记录.终止时间, c_记录.缺省预约时间, c_记录.提前挂号时间,
             c_记录.限号数, c_记录.限约数, c_记录.是否序号控制, c_记录.是否分时段, c_记录.预约控制, c_记录.项目id, c_号源.科室id, c_记录.医生id, c_记录.医生姓名,
             c_记录.分诊方式, c_记录.诊室id, 操作员_In, 登记时间_In, c_记录.是否独占);
        
          --插入临床出诊序号控制
          Insert Into 临床出诊序号控制
            (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约)
            Select c_记录.记录id, 序号,
                   To_Date(To_Char(c_日期.日期, 'yyyy-mm-dd ') || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
                   To_Date(To_Char(c_日期.日期, 'yyyy-mm-dd ') || To_Char(终止时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') + Case
                     When 终止时间 <= 开始时间 Then
                      1
                     Else
                      0
                   End, 限制数量, 是否预约
            From 临床出诊时段
            Where 限制id = c_记录.限制id;
        
          --插入合作单位挂号控制记录
          Insert Into 临床出诊挂号控制记录
            (类型, 性质, 名称, 记录id, 序号, 控制方式, 数量)
            Select 类型, 性质, 名称, c_记录.记录id, 序号, 控制方式, 数量
            From 临床出诊挂号控制
            Where 限制id = c_记录.限制id;
        
          --插入临床出诊诊室记录
          Insert Into 临床出诊诊室记录
            (记录id, 诊室id)
            Select c_记录.记录id, 诊室id From 临床出诊诊室 Where 限制id = c_记录.限制id;
        End Loop;
      End If;
    End Loop;
  End Loop;

  --加入没有的出诊安排的号源
  For c_号源 In (Select 临床出诊安排_Id.Nextval As 安排id, 出诊id_In As 出诊id, a.Id As 号源id, a.项目id, a.医生id, a.医生姓名
               From 临床出诊号源 A, 部门表 D
               Where a.科室id = d.Id
                    --有效号源
                     And Nvl(a.是否删除, 0) = 0 And (a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd') Or a.撤档时间 Is Null) And
                     (
                     --月排班
                      Nvl(排班方式_In, 0) = 1 And a.排班方式 = 1
                     -- 周排班
                      Or Nvl(排班方式_In, 0) = 2 And
                      (
                      --当前出诊表所在时间范围内不能有月排班
                       a.排班方式 = 2 And Not Exists
                       (Select 1
                           From 临床出诊安排 P, 临床出诊表 Q
                           Where p.出诊id = q.Id And p.号源id = a.Id And
                                 Not (p.终止时间 < Trunc(开始时间_In, 'MONTH') Or p.开始时间 > Last_Day(开始时间_In)) And q.排班方式 = 1)
                      --当前已调整为了月排班,但是本月又用了周排班，则本月剩下的部分将继续按周进行排班
                       Or a.排班方式 = 1 And Exists
                       (Select 1
                           From 临床出诊安排 P, 临床出诊表 Q
                           Where p.出诊id = q.Id And p.号源id = a.Id And
                                 Not (p.终止时间 < Trunc(开始时间_In, 'MONTH') Or p.开始时间 > Last_Day(开始时间_In)) And q.排班方式 = 2)))
                    --号源在该出诊表时间范围内无出诊记录
                     And Not Exists
                (Select 1
                      From 临床出诊记录 P
                      Where p.号源id = a.Id And p.出诊日期 Between 开始时间_In And 终止时间_In)
                    --当前人员可操作的号源
                     And (Nvl(人员id_In, 0) = 0 Or (Nvl(a.是否临床排班, 0) = 1 And Exists
                      (Select 1 From 部门人员 Where 部门id = a.科室id And 人员id = 人员id_In)))
                    --站点
                     And (d.站点 Is Null Or d.站点 = 站点_In)
                    
                     And Not Exists (Select 1 From 临床出诊安排 Where 出诊id = 出诊id_In And 号源id = a.Id)) Loop
  
    Insert Into 临床出诊安排
      (ID, 出诊id, 号源id, 项目id, 医生id, 医生姓名, 开始时间, 终止时间, 操作员姓名, 登记时间, 原终止时间)
    Values
      (c_号源.安排id, c_号源.出诊id, c_号源.号源id, c_号源.项目id, c_号源.医生id, c_号源.医生姓名, 开始时间_In, 终止时间_In, 操作员_In, 登记时间_In, 终止时间_In);
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Buildregisterplanbytemplet;
/
Create Or Replace Procedure Zl_临床出诊表_Delete
(
  Id_In     临床出诊表.Id%Type,
  人员id_In 人员表.Id%Type := Null,
  站点_In   部门表.站点%Type
) As
  --功能：删除临床出诊表
  --参数：
  --        人员id_In 除固定安排外有效，不为0或null表示临床科室人员在删除
  n_Count    Number;
  n_排班方式 临床出诊表.排班方式%Type;
  n_出诊id   临床出诊表.Id%Type;

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  l_记录id t_Numlist := t_Numlist();
  l_限制id t_Numlist := t_Numlist();
Begin
  Begin
    Select 1 Into n_Count From 临床出诊表 Where 排班方式 <> 3 And 发布人 Is Not Null And ID = Id_In;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count <> 0 Then
    v_Err_Msg := '当前出诊表已发布，不能删除！';
    Raise Err_Item;
  End If;

  Begin
    Select 排班方式 Into n_排班方式 From 临床出诊表 Where ID = Id_In;
  Exception
    When Others Then
      v_Err_Msg := '出诊表信息未找到！';
      Raise Err_Item;
  End;

  If Nvl(n_排班方式, 0) = 0 Or Nvl(n_排班方式, 0) = 3 Then
    --固定安排/模板
    --删除临床出诊限制
    Select b.Id Bulk Collect
    Into l_限制id
    From 临床出诊安排 A, 临床出诊限制 B
    Where a.Id = b.安排id And a.出诊id = Id_In;
  
    Forall I In 1 .. l_限制id.Count
      Delete From 临床出诊时段 Where 限制id = l_限制id(I);
  
    Forall I In 1 .. l_限制id.Count
      Delete From 临床出诊诊室 Where 限制id = l_限制id(I);
  
    Forall I In 1 .. l_限制id.Count
      Delete From 临床出诊挂号控制 Where 限制id = l_限制id(I);
  
    Forall I In 1 .. l_限制id.Count
      Delete From 临床出诊限制 Where ID = l_限制id(I);
  
    --删除临床出诊安排
    Delete From 临床出诊安排 Where 出诊id = Id_In;
  
    --删除临床出诊表
    Delete 临床出诊表 Where ID = Id_In;
  
    Return;
  End If;

  --========================================================================================================
  --月出诊表/周出诊表
  --只能从最后一个开始删除
  Begin
    Select ID
    Into n_出诊id
    From (Select a.Id
           From 临床出诊表 A, 临床出诊安排 B, 临床出诊号源 C, 部门表 D
           Where a.排班方式 = n_排班方式 And a.Id = b.出诊id And b.号源id = c.Id And c.科室id = d.Id
                --当前人员可操作的号源
                 And (Nvl(人员id_In, 0) = 0 Or (Nvl(c.是否临床排班, 0) = 1 And Exists
                  (Select 1 From 部门人员 Where 部门id = c.科室id And 人员id = 人员id_In)))
                --站点
                 And (d.站点 Is Null Or d.站点 = 站点_In)
           Order By a.年份 Desc, a.月份 Desc, a.周数 Desc)
    Where Rownum < 2;
  Exception
    When Others Then
      n_出诊id := 0;
  End;
  If Nvl(n_出诊id, 0) <> 0 And Nvl(n_出诊id, 0) <> Id_In Then
    v_Err_Msg := '必须从最后一个出诊表开始删除！';
    Raise Err_Item;
  End If;

  --恢复固定安排的终止时间
  For c_安排 In (Select a.Id, a.原终止时间
               From 临床出诊安排 A, 临床出诊安排 B, 临床出诊表 C
               Where a.号源id = b.号源id And a.终止时间 = b.开始时间 - 1 / 24 / 60 / 60 And a.出诊id = c.Id And c.排班方式 = 0 And
                     b.出诊id = Id_In) Loop
    Update 临床出诊安排 Set 终止时间 = c_安排.原终止时间 Where ID = c_安排.Id;
  End Loop;

  --删除临床出诊记录
  Select a.Id Bulk Collect
  Into l_记录id
  From 临床出诊记录 A, 临床出诊安排 B, 临床出诊号源 C, 部门表 D
  Where a.安排id = b.Id And a.号源id = c.Id And c.科室id = d.Id And b.出诊id = Id_In
       --当前人员可操作的号源
        And (Nvl(人员id_In, 0) = 0 Or
        (Nvl(c.是否临床排班, 0) = 1 And Exists (Select 1 From 部门人员 Where c.科室id = 部门id And 人员id = 人员id_In)))
       --站点
        And (d.站点 Is Null Or d.站点 = 站点_In);

  Zl_临床出诊记录_Batchdelete(l_记录id);

  --删除临床出诊安排
  Delete From 临床出诊安排 A
  Where a.出诊id = Id_In And Exists
   (Select 1
         From 临床出诊号源 B, 部门表 D
         Where a.号源id = b.Id And b.科室id = d.Id
              --当前人员可操作的号源
               And (Nvl(人员id_In, 0) = 0 Or (Nvl(b.是否临床排班, 0) = 1 And Exists
                (Select 1 From 部门人员 Where b.科室id = 部门id And 人员id = 人员id_In)))
              --站点
               And (d.站点 Is Null Or d.站点 = 站点_In));

  --删除临床出诊表
  Delete 临床出诊表 A
  Where a.Id = Id_In And Not Exists (Select 1 From 临床出诊安排 Where 出诊id = a.Id And 号源id Is Not Null);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊表_Delete;
/

Create Or Replace Procedure Zl_临床出诊安排_Applyto
(
  应用类型_In     Number,
  原id_In         临床出诊安排.Id%Type,
  原项目_In       Varchar2,
  新id_In         临床出诊安排.Id%Type,
  新项目_In       Varchar2,
  是否临时出诊_In Number := 0
) As
  -------------------------------------------------------------------------
  --功能：将某个日期的安排应用于其他日期
  --参数：
  --     原Id_In 被应用的安排ID
  --     原项目_in 被应用的项目
  --           1.模板或固定出诊表，限制项目，如"周三"
  --           2.出诊记录，出诊日期，如"2016-01-02"
  --     新id_In 应用于的安排ID
  --     新项目_In 应用于的项目（多个用"|"分隔）
  --           1.模板或固定出诊表，限制项目：项目1|项目2|...，如"周三|周五"
  --           2.出诊记录，出诊日期：日期1|日期2|...，如"2016-01-02|2016-01-05"
  --     应用类型_In 0-模板或固定出诊表,1-出诊记录
  --说明：
  -------------------------------------------------------------------------
  n_Count    Number;
  n_限制id   临床出诊限制.Id%Type;
  n_记录id   临床出诊记录.Id%Type;
  d_出诊日期 临床出诊记录.出诊日期%Type;

  v_Err_Msg Varchar2(255);
  Err_Item Exception;
Begin
  --检查被应用的安排是否有效
  If Nvl(应用类型_In, 0) = 0 Then
    Select Count(1)
    Into n_Count
    From 临床出诊安排 A, 临床出诊限制 B
    Where a.Id = b.安排id And a.Id = 原id_In And b.限制项目 = 原项目_In;
  Else
    Select Count(1)
    Into n_Count
    From 临床出诊安排 A, 临床出诊记录 B
    Where a.Id = b.安排id And a.Id = 原id_In And b.出诊日期 = To_Date(原项目_In, 'yyyy-mm-dd');
  End If;
  If n_Count = 0 Then
    v_Err_Msg := '被应用的安排未设置有效的上班时段，不能应用于其它安排！';
    Raise Err_Item;
  End If;

  Select Count(1) Into n_Count From 临床出诊安排 Where ID = 新id_In;
  If n_Count = 0 Then
    v_Err_Msg := '未发现你将要应用于的临床出诊安排记录！';
    Raise Err_Item;
  End If;
  If 新项目_In Is Null Then
    v_Err_Msg := '无应用于的项目！';
    Raise Err_Item;
  End If;

  If Nvl(应用类型_In, 0) = 0 Then
    --模板或固定出诊表
    For c_限制项目 In (Select Column_Value As 项目 From Table(f_Str2list(新项目_In, '|'))) Loop
      --先删除已有时段
      Zl_临床出诊上班时段_Delete(新id_In, c_限制项目.项目, 0);
    
      For c_时段 In (Select ID, 安排id, 限制项目, 上班时段, 限号数, 限约数, 是否序号控制, 是否分时段, 预约控制, 分诊方式, 诊室id, 是否独占
                   From 临床出诊限制
                   Where 安排id = 原id_In And 限制项目 = 原项目_In) Loop
      
        Select 临床出诊限制_Id.Nextval Into n_限制id From Dual;
        Insert Into 临床出诊限制
          (ID, 安排id, 限制项目, 上班时段, 限号数, 限约数, 是否序号控制, 是否分时段, 预约控制, 分诊方式, 诊室id, 是否独占)
        Values
          (n_限制id, 新id_In, c_限制项目.项目, c_时段.上班时段, c_时段.限号数, c_时段.限约数, c_时段.是否序号控制, c_时段.是否分时段, c_时段.预约控制, c_时段.分诊方式,
           c_时段.诊室id, c_时段.是否独占);
      
        Insert Into 临床出诊诊室
          (限制id, 诊室id)
          Select n_限制id, 诊室id From 临床出诊诊室 Where 限制id = c_时段.Id;
      
        Insert Into 临床出诊时段
          (限制id, 序号, 开始时间, 终止时间, 限制数量, 是否预约)
          Select n_限制id, 序号, 开始时间, 终止时间, 限制数量, 是否预约 From 临床出诊时段 Where 限制id = c_时段.Id;
      
        Insert Into 临床出诊挂号控制
          (限制id, 类型, 性质, 名称, 序号, 控制方式, 数量)
          Select n_限制id, 类型, 性质, 名称, 序号, 控制方式, 数量 From 临床出诊挂号控制 Where 限制id = c_时段.Id;
      End Loop;
    End Loop;
  Else
    --出诊记录
    For c_出诊日期 In (Select Column_Value As 日期 From Table(f_Str2list(新项目_In, '|'))) Loop
      d_出诊日期 := To_Date(c_出诊日期.日期, 'yyyy-mm-dd');
      --不能对历史的安排进行出诊安排操作
      If Trunc(Sysdate + 1) > d_出诊日期 Then
        v_Err_Msg := '不能对当前日期及以前的日期进行出诊安排！';
        Raise Err_Item;
      End If;
    
      --检查当前日期是否已由其它出诊表生成
      --一个号源某一天的安排只能由一个出诊表设置
      Begin
        Select 1
        Into n_Count
        From 临床出诊记录 A, 临床出诊安排 B, 临床出诊安排 C
        Where a.安排id = b.Id And a.号源id = c.号源id And a.出诊日期 = d_出诊日期 And c.Id = 新id_In And b.Id <> 新id_In And Rownum < 2;
      Exception
        When Others Then
          n_Count := 0;
      End;
      If Nvl(n_Count, 0) = 1 Then
        v_Err_Msg := '日期(' || To_Char(d_出诊日期, 'yyyy-mm-dd') || ')已在其它出诊表中进行了安排，不能重复安排！';
        Raise Err_Item;
      End If;
    
      --先删除已有时段
      Zl_临床出诊上班时段_Delete(新id_In, To_Char(d_出诊日期, 'yyyy-mm-dd'), 1);
    
      For c_时段 In (Select ID, 安排id, 号源id, 出诊日期, 上班时段, 开始时间, 终止时间, 缺省预约时间, 提前挂号时间, 限号数, 限约数, 是否序号控制, 是否分时段, 预约控制, 是否独占,
                          项目id, 科室id, 医生id, 医生姓名, 分诊方式, 诊室id
                   From 临床出诊记录
                   Where 安排id = 原id_In And 出诊日期 = To_Date(原项目_In, 'yyyy-mm-dd')) Loop
      
        Select 临床出诊记录_Id.Nextval Into n_记录id From Dual;
        Insert Into 临床出诊记录
          (ID, 安排id, 号源id, 出诊日期, 上班时段, 开始时间, 终止时间, 缺省预约时间, 提前挂号时间, 限号数, 限约数, 是否序号控制, 是否分时段, 预约控制, 是否独占, 项目id, 科室id,
           医生id, 医生姓名, 分诊方式, 诊室id, 是否临时出诊, 登记人, 登记时间)
          Select n_记录id, a.Id, a.号源id, d_出诊日期, c_时段.上班时段,
                 To_Date(To_Char(d_出诊日期, 'yyyy-mm-dd') || ' ' || To_Char(c_时段.开始时间, 'hh24:mi:ss'),
                          'yyyy-mm-dd hh24:mi:ss'),
                 To_Date(To_Char(d_出诊日期, 'yyyy-mm-dd') || ' ' || To_Char(c_时段.终止时间, 'hh24:mi:ss'),
                          'yyyy-mm-dd hh24:mi:ss'),
                 To_Date(To_Char(d_出诊日期, 'yyyy-mm-dd') || ' ' || To_Char(c_时段.缺省预约时间, 'hh24:mi:ss'),
                          'yyyy-mm-dd hh24:mi:ss'),
                 To_Date(To_Char(d_出诊日期, 'yyyy-mm-dd') || ' ' || To_Char(c_时段.提前挂号时间, 'hh24:mi:ss'),
                          'yyyy-mm-dd hh24:mi:ss'), c_时段.限号数, c_时段.限约数, c_时段.是否序号控制, c_时段.是否分时段, c_时段.预约控制, c_时段.是否独占,
                 a.项目id, b.科室id, a.医生id, a.医生姓名, c_时段.分诊方式, c_时段.诊室id, Nvl(是否临时出诊_In, 0), Zl_Username, Sysdate
          From 临床出诊安排 A, 临床出诊号源 B
          Where a.Id = 新id_In And a.号源id = b.Id;
      
        Insert Into 临床出诊诊室记录
          (记录id, 诊室id)
          Select n_记录id, 诊室id From 临床出诊诊室记录 Where 记录id = c_时段.Id;
      
        Insert Into 临床出诊序号控制
          (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约)
          Select n_记录id, 序号,
                 To_Date(To_Char(d_出诊日期, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
                 To_Date(To_Char(d_出诊日期, 'yyyy-mm-dd') || ' ' || To_Char(终止时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
                 数量, 是否预约
          From 临床出诊序号控制
          Where 预约顺序号 Is Null And 记录id = c_时段.Id;
      
        Insert Into 临床出诊挂号控制记录
          (记录id, 类型, 性质, 名称, 序号, 控制方式, 数量)
          Select n_记录id, 类型, 性质, 名称, 序号, 控制方式, 数量 From 临床出诊挂号控制记录 Where 记录id = c_时段.Id;
      
      End Loop;
    End Loop;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊安排_Applyto;
/

Create Or Replace Procedure Zl_临床出诊安排_Batchdelete
(
  出诊id_In 临床出诊表.Id%Type,
  人员id_In 人员表.Id%Type := 0,
  站点_In   部门表.站点%Type := Null,
  号源id_In 临床出诊安排.号源id%Type := 0
) As
  --功能：批量删除临床出诊安排
  --参数：
  --      人员id_In 不等于0则删除人员所在科室的所有号源安排
  --      号源id_In 不等于0则删除该号源的所有安排
  --说明：如果人员id_In=0且号源id_In=0 则删除该出诊表的所有号源的所有安排
  n_Count    Number(8);
  n_出诊记录 Number(1);

  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  l_限制id t_Numlist := t_Numlist();
  l_记录id t_Numlist := t_Numlist();
Begin
  Begin
    Select 1
    Into n_Count
    From 临床出诊表 A
    Where a.Id = 出诊id_In And a.发布人 Is Not Null And a.排班方式 <> 3 And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count <> 0 Then
    v_Err_Msg := '当前出诊表已发布，不允许修改安排！';
    Raise Err_Item;
  End If;

  Begin
    Select 1 Into n_出诊记录 From 临床出诊表 A Where a.Id = 出诊id_In And a.排班方式 In (1, 2) And Rownum < 2;
  Exception
    When Others Then
      n_出诊记录 := 0;
  End;

  If Nvl(n_出诊记录, 0) = 0 Then
    --删除临床出诊规则/模板
    Select a.Id Bulk Collect
    Into l_限制id
    From 临床出诊限制 A, 临床出诊安排 B, 临床出诊号源 C, 部门表 D
    Where a.安排id = b.Id And b.号源id = c.Id And c.科室id = d.Id And b.出诊id = 出诊id_In And
          (
          --删除该出诊表的所有号源的所有安排
           (Nvl(号源id_In, 0) = 0 And Nvl(人员id_In, 0) = 0)
          --删除该号源的所有安排
           Or (Nvl(号源id_In, 0) <> 0 And b.号源id = 号源id_In)
          --删除人员所在科室的所有号源安排
           Or (Nvl(人员id_In, 0) <> 0 And Exists (Select 1 From 部门人员 Where 部门id = c.科室id And 人员id = 人员id_In)))
         --站点
          And (d.站点 Is Null Or d.站点 = 站点_In);
  
    Forall I In 1 .. l_限制id.Count
      Delete From 临床出诊挂号控制 Where 限制id = l_限制id(I);
  
    Forall I In 1 .. l_限制id.Count
      Delete From 临床出诊时段 Where 限制id = l_限制id(I);
  
    Forall I In 1 .. l_限制id.Count
      Delete From 临床出诊诊室 Where 限制id = l_限制id(I);
  
    Forall I In 1 .. l_限制id.Count
      Delete From 临床出诊限制 Where ID = l_限制id(I);
  Else
    --删除临床出诊记录
    Select a.Id Bulk Collect
    Into l_记录id
    From 临床出诊记录 A, 临床出诊安排 B, 临床出诊号源 C, 部门表 D
    Where a.安排id = b.Id And b.号源id = c.Id And c.科室id = d.Id And b.出诊id = 出诊id_In And
          (
          --删除该出诊表的所有号源的所有安排
           (Nvl(号源id_In, 0) = 0 And Nvl(人员id_In, 0) = 0)
          --删除该号源的所有安排
           Or (Nvl(号源id_In, 0) <> 0 And b.号源id = 号源id_In)
          --删除人员所在科室的所有号源安排
           Or (Nvl(人员id_In, 0) <> 0 And Exists (Select 1 From 部门人员 Where 部门id = c.科室id And 人员id = 人员id_In)))
         --站点
          And (d.站点 Is Null Or d.站点 = 站点_In);
  
    Zl_临床出诊记录_Batchdelete(l_记录id);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊安排_Batchdelete;
/

Create Or Replace Procedure Zl_临床出诊安排_Publish
(
  Id_In       临床出诊表.Id%Type,
  发布人_In   临床出诊表.发布人%Type := Null,
  发布时间_In 临床出诊表.发布时间%Type := Null,
  取消发布_In Number := 0
) As
  --发布和取消发布安排
  --参数：
  --        取消发布_In 是否取消发布
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_Count    Number(2);
  n_排班方式 临床出诊表.排班方式%Type;
  l_记录id   t_Numlist := t_Numlist();

  d_停诊开始时间 临床出诊记录.停诊开始时间%Type;
  d_停诊终止时间 临床出诊记录.停诊终止时间%Type;
  v_停诊原因     临床出诊记录.停诊原因%Type;
  d_原上班日期   临床出诊记录.出诊日期%Type;
  d_调休日期     临床出诊记录.出诊日期%Type;
  n_返回         Number(2);

  n_挂号排班模式 Number(2);
  v_挂号排班模式 Varchar2(255);

  Function Fun_Isclinicvisit
  (
    号源id_In       In 临床出诊号源.Id%Type,
    开始时间_In     In 临床出诊记录.开始时间%Type,
    终止时间_In     In 临床出诊记录.终止时间%Type,
    停诊开始时间_In Out 临床出诊记录.停诊开始时间%Type,
    停诊终止时间_In Out 临床出诊记录.停诊终止时间%Type,
    停诊原因_In     Out 临床出诊记录.停诊原因%Type,
    原上班日期_In   Out 临床出诊记录.出诊日期%Type,
    调休日期_In     Out 临床出诊记录.出诊日期%Type
  ) Return Number As
    --功能：判断医生在某个时间范围是否出诊
    --     按每一个时间段进行检查(出诊日期+上班时间段)
    --入参：
    --     号源id_In：临床出诊号源ID
    --     开始时间_In：上班时段的开始时间
    --     终止时间_In：上班时段的终止时间
    --出参：
    --     停诊原因_In：不出诊时返回停诊原因（多个以第一个为准），否则返回空
    --返回：
    --     0-检查出错
    --     1-在法定节假日内，同时临床出诊号源.假日控制状态=0(0-不上班;1-上班且开放预约;2-上班但不开放预约)
    --     2-在停诊安排时间范围内
    --     else-正常出诊
    --说明：
    --     1)全部都不在不出诊时间范围内-->出诊
    --     2)全部都在不出诊时间范围内-->不出诊
    --     3)部分在不出诊时间范围内-->不出诊
  
    n_假日控制状态 临床出诊号源.假日控制状态%Type;
    n_是否假日换休 临床出诊号源.是否假日换休%Type;
  Begin
    --法定节假日
    Begin
      Select Nvl(b.假日控制状态, 0), Nvl(是否假日换休, 0)
      Into n_假日控制状态, n_是否假日换休
      From 临床出诊号源 B
      Where b.Id = 号源id_In;
    Exception
      When Others Then
        n_假日控制状态 := 0;
        n_是否假日换休 := 0;
    End;
  
    If Nvl(n_假日控制状态, 0) = 0 Then
      --假日控制状态:0-不上班;1-上班且开放预约;2-上班但不开放预约
      Begin
        Select a.开始日期, a.终止日期, a.节日名称
        Into 停诊开始时间_In, 停诊终止时间_In, 停诊原因_In
        From 法定假日表 A
        Where a.性质 = 0 And 开始时间_In < a.终止日期 And 终止时间_In > a.开始日期 And Rownum < 2;
      
        If 停诊开始时间_In < 开始时间_In Then
          停诊开始时间_In := 开始时间_In;
        End If;
        If 停诊终止时间_In > 终止时间_In Then
          停诊终止时间_In := 终止时间_In;
        End If;
      
        --确定是否需要换休
        If Nvl(n_是否假日换休, 0) = 1 Then
          --1.前面的换到后面
          Begin
            --开始日期：原本休息日(即调休日) ， 终止日期：原本上班日(即被调休日)
            Select a.终止日期
            Into 原上班日期_In
            From 法定假日表 A
            Where a.性质 = 1 And 开始时间_In < a.开始日期 + 1 - 1 / 24 / 60 / 60 And 终止时间_In > a.开始日期 And Rownum < 2;
          Exception
            When Others Then
              原上班日期_In := Null;
          End;
        
          --2.后面的换到前面，可能后面的还没有发布，在发布前面的出诊表时是没有换的
          Begin
            --开始日期：原本休息日(即调休日) ， 终止日期：原本上班日(即被调休日)
            Select a.开始日期
            Into 调休日期_In
            From 法定假日表 A
            Where a.性质 = 1 And 开始时间_In < a.终止日期 + 1 - 1 / 24 / 60 / 60 And 终止时间_In > a.终止日期 And Rownum < 2;
          Exception
            When Others Then
              调休日期_In := Null;
          End;
        End If;
      
        Return 1;
      Exception
        When Others Then
          停诊开始时间_In := Null;
          停诊终止时间_In := Null;
          停诊原因_In     := Null;
      End;
    End If;
  
    --停诊安排
    Begin
      Select a.开始时间, a.终止时间, a.停诊原因
      Into 停诊开始时间_In, 停诊终止时间_In, 停诊原因_In
      From 临床出诊停诊记录 A, 临床出诊号源 B
      Where a.申请人 = b.医生姓名 And a.记录id Is Null And a.审批时间 Is Not Null And a.取消人 Is Null And b.医生id Is Not Null And
            b.Id = 号源id_In And Not (开始时间_In >= a.终止时间 Or 终止时间_In <= a.开始时间) And Rownum < 2;
    
      If 停诊开始时间_In < 开始时间_In Then
        停诊开始时间_In := 开始时间_In;
      End If;
      If 停诊终止时间_In > 终止时间_In Then
        停诊终止时间_In := 终止时间_In;
      End If;
    
      Return 2;
    Exception
      When Others Then
        停诊开始时间_In := Null;
        停诊终止时间_In := Null;
        停诊原因_In     := Null;
    End;
  
    Return - 1;
  Exception
    When Others Then
      Return 0;
  End Fun_Isclinicvisit;
Begin
  Begin
    Select Nvl(排班方式, 0) Into n_排班方式 From 临床出诊表 Where ID = Id_In;
  Exception
    When Others Then
      v_Err_Msg := '出诊表信息未找到！';
      Raise Err_Item;
  End;

  If Nvl(取消发布_In, 0) = 0 Then
    --发布安排
    If Nvl(n_排班方式, 0) = 0 Then
      Begin
        Select 1
        Into n_Count
        From 临床出诊安排 A, 临床出诊限制 B, 临床出诊表 C
        Where a.Id = b.安排id And a.出诊id = c.Id And c.排班方式 = 0 And c.Id = Id_In And Rownum < 2;
      Exception
        When Others Then
          n_Count := 0;
      End;
      If n_Count = 0 Then
        v_Err_Msg := '当前出诊表无有效的安排，不能发布！';
        Raise Err_Item;
      End If;
    Else
      Begin
        Select 1
        Into n_Count
        From 临床出诊安排 A, 临床出诊记录 B, 临床出诊表 C
        Where a.Id = b.安排id And a.出诊id = c.Id And c.排班方式 In (1, 2) And c.Id = Id_In And Rownum < 2;
      Exception
        When Others Then
          n_Count := 0;
      End;
      If n_Count = 0 Then
        v_Err_Msg := '当前出诊表无有效的安排，不能发布！';
        Raise Err_Item;
      End If;
    
      Begin
        Select 1
        Into n_Count
        From 临床出诊记录 A, 临床出诊安排 B
        Where a.号源id = b.号源id And a.出诊日期 Between b.开始时间 And b.终止时间 And a.安排id <> b.Id And b.出诊id = Id_In And Rownum < 2;
      Exception
        When Others Then
          n_Count := 0;
      End;
      If n_Count <> 0 Then
        v_Err_Msg := '当前出诊表中的部分号源在当前出诊表的生效时间范围内已经存在有效的安排，不能发布！';
        Raise Err_Item;
      End If;
    
      Begin
        Select 1 Into n_Count From 临床出诊安排 A Where a.出诊id = Id_In And a.开始时间 < Sysdate And Rownum < 2;
      Exception
        When Others Then
          n_Count := 0;
      End;
      If n_Count <> 0 Then
        v_Err_Msg := '当前时间大于了出诊表的开始时间，不能发布！';
        Raise Err_Item;
      End If;
    End If;
  
    --如果存在多个未发布的安排表，则不允许发布后面日期的安排，必须按最小有效时间进行发布
    Begin
      Select 1
      Into n_Count
      From 临床出诊安排 A, 临床出诊表 B, 临床出诊安排 C
      Where a.出诊id = b.Id And b.排班方式 = Nvl(n_排班方式, 0) And a.号源id = c.号源id And a.开始时间 < c.开始时间 And b.Id <> c.出诊id And
            b.发布人 Is Null And c.出诊id = Id_In And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count <> 0 Then
      v_Err_Msg := '当前存在多个未发布的安排表，必须按最小开始时间进行发布！';
      Raise Err_Item;
    End If;
  
    Update 临床出诊表 Set 发布人 = 发布人_In, 发布时间 = 发布时间_In Where ID = Id_In;
  
    --删除发布时有安排，但是号源已被停用的记录
    For c_安排 In (Select a.Id
                 From 临床出诊安排 A, 临床出诊号源 B
                 Where a.号源id = b.Id And a.出诊id = Id_In And
                       Not (Nvl(b.是否删除, 0) = 0 And (b.撤档时间 Is Null Or b.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd')))) Loop
      Zl_临床出诊安排_Delete(c_安排.Id, Nvl(n_排班方式, 0));
    End Loop;
  
    --固定安排修改当前有效安排的终止时间，同一时间同一号源有效固定安排只会有一个
    If Nvl(n_排班方式, 0) = 0 Then
      For c_安排 In (Select a.Id, b.开始时间
                   From 临床出诊安排 A, 临床出诊安排 B, 临床出诊表 C
                   Where a.号源id = b.号源id And a.终止时间 > b.开始时间 And a.出诊id = c.Id And Nvl(c.排班方式, 0) = 0 And
                         a.出诊id <> b.出诊id And c.发布人 Is Not Null And b.出诊id = Id_In) Loop
        Update 临床出诊安排 Set 终止时间 = c_安排.开始时间 - 1 Where ID = c_安排.Id;
      End Loop;
    
      --"月排班"/"周排班"调整过来的号源可能在当前固定安排的有效时间内已有出诊记录,需要调整固定安排的生效时间
      For c_安排 In (Select a.Id, b.终止时间
                   From 临床出诊安排 A, 临床出诊安排 B, 临床出诊表 C
                   Where a.号源id = b.号源id And b.出诊id = c.Id And c.排班方式 In (1, 2) And a.出诊id = Id_In And a.开始时间 < b.终止时间) Loop
        Update 临床出诊安排 Set 开始时间 = c_安排.终止时间 + 1 Where ID = c_安排.Id;
      End Loop;
    Else
      --月安排/周安排产生停诊信息
      For c_记录 In (Select a.安排id, a.出诊日期, a.Id, a.号源id, a.开始时间, a.终止时间
                   From 临床出诊记录 A, 临床出诊安排 B
                   Where a.安排id = b.Id And b.出诊id = Id_In
                   Order By a.出诊日期, a.上班时段) Loop
      
        n_返回 := Fun_Isclinicvisit(c_记录.号源id, c_记录.开始时间, c_记录.终止时间, d_停诊开始时间, d_停诊终止时间, v_停诊原因, d_原上班日期, d_调休日期);
        If Nvl(n_返回, 0) = 1 Or Nvl(n_返回, 0) = 2 Then
          --节假日或者停诊安排
          Update 临床出诊记录
          Set 停诊开始时间 = d_停诊开始时间, 停诊终止时间 = d_停诊终止时间, 停诊原因 = v_停诊原因
          Where ID = c_记录.Id;
        
          --产生停诊记录
          Insert Into 临床出诊停诊记录
            (ID, 记录id, 开始时间, 终止时间, 停诊原因, 申请人, 申请时间, 审批人, 审批时间)
          Values
            (临床出诊停诊记录_Id.Nextval, c_记录.Id, d_停诊开始时间, d_停诊终止时间, v_停诊原因, 发布人_In, 发布时间_In, 发布人_In, 发布时间_In);
        End If;
      
        If Nvl(n_返回, 0) = 1 Then
          --进行换休处理
          If d_原上班日期 Is Not Null Then
            --由休息日换到上班日期
            Begin
              Select 1
              Into n_Count
              From 临床出诊记录
              Where 号源id = c_记录.号源id And 出诊日期 = d_原上班日期 And Nvl(是否发布, 0) = 1 And Rownum < 2;
            Exception
              When Others Then
                n_Count := 0;
            End;
            If n_Count > 0 Then
              --先删除现有的
              Select ID Bulk Collect Into l_记录id From 临床出诊记录 Where ID = c_记录.Id;
              Zl_临床出诊记录_Batchdelete(l_记录id);
            
              For c_换休记录 In (Select ID
                             From 临床出诊记录
                             Where 号源id = c_记录.号源id And 出诊日期 = d_原上班日期 And Nvl(是否发布, 0) = 1) Loop
                Zl_临床出诊记录_Copy(c_换休记录.Id, c_记录.安排id, c_记录.出诊日期, 发布人_In, 发布时间_In);
              End Loop;
            End If;
          End If;
        
          If d_调休日期 Is Not Null Then
            --由上班日期换到休息日，如果调休日期已有的安排已被使用则不换休
            Begin
              Select 1
              Into n_Count
              From 临床出诊记录
              Where 号源id = c_记录.号源id And 出诊日期 = c_记录.出诊日期 And Rownum < 2;
            Exception
              When Others Then
                n_Count := 0;
            End;
            If n_Count > 0 Then
              Begin
                Select 1
                Into n_Count
                From 临床出诊记录 A, 病人挂号记录 B
                Where a.Id = b.出诊记录id And a.号源id = c_记录.号源id And a.出诊日期 = d_调休日期 And Rownum < 2;
              Exception
                When Others Then
                  n_Count := 0;
              End;
              If n_Count = 0 Then
                --先删除现有的
                Select ID Bulk Collect
                Into l_记录id
                From 临床出诊记录
                Where 号源id = c_记录.号源id And 出诊日期 = d_调休日期;
                Zl_临床出诊记录_Batchdelete(l_记录id);
              
                For c_换休记录 In (Select ID From 临床出诊记录 Where 号源id = c_记录.号源id And 出诊日期 = c_记录.出诊日期) Loop
                  Zl_临床出诊记录_Copy(c_换休记录.Id, c_记录.安排id, d_调休日期, 发布人_In, 发布时间_In);
                End Loop;
              End If;
            End If;
          End If;
        End If;
      End Loop;
    
      --修改临床出诊记录中的"是否发布"
      Select a.Id Bulk Collect
      Into l_记录id
      From 临床出诊记录 A, 临床出诊安排 B
      Where a.安排id = b.Id And b.出诊id = Id_In;
    
      Forall I In 1 .. l_记录id.Count
        Update 临床出诊记录 Set 是否发布 = 1 Where ID = l_记录id(I);
    End If;
    Return;
  End If;

  --==================================================================================================================
  --取消发布
  Begin
    Select 1
    Into n_Count
    From 临床出诊安排 A, 临床出诊安排 B, 临床出诊表 C, 临床出诊表 D
    Where a.开始时间 > b.开始时间 And a.号源id = b.号源id And a.出诊id = c.Id And b.出诊id = d.Id And c.排班方式 = d.排班方式 And
          c.发布时间 Is Not Null And b.出诊id = Id_In And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count <> 0 Then
    v_Err_Msg := '当前出诊表开始时间之后还存在已发布了的出诊表，不允许对当前出诊表进行取消发布！';
    Raise Err_Item;
  End If;

  Begin
    Select 1
    Into n_Count
    From 病人挂号记录 C, 临床出诊记录 A, 临床出诊安排 B
    Where c.出诊记录id = a.Id And a.安排id = b.Id And b.出诊id = Id_In And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count <> 0 Then
    v_Err_Msg := '当前出诊表的安排已被使用，不允许取消发布！';
    Raise Err_Item;
  End If;

  Begin
    Select 1 Into n_Count From 临床出诊安排 A Where a.出诊id = Id_In And Sysdate > a.开始时间 And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count <> 0 Then
    Select Nvl(zl_GetSysParameter(256), '|') Into v_挂号排班模式 From Dual;
    n_挂号排班模式 := To_Number(Substr(v_挂号排班模式 || '|', 1, Instr(v_挂号排班模式 || '|', '|') - 1));
    If Nvl(n_挂号排班模式, 0) = 1 Then
      --没切换挂号排班模式时可以取消发布
      v_Err_Msg := '当前日期已经在当前安排的有效时间范围内或者大于了当前安排的终止时间，不允许取消发布！';
      Raise Err_Item;
    End If;
  End If;

  Update 临床出诊表 Set 发布人 = Null, 发布时间 = Null Where ID = Id_In;
  If Sql%NotFound Then
    v_Err_Msg := '出诊表信息未找到！';
    Raise Err_Item;
  End If;

  --固定安排取消发布时删除出诊记录并恢复原安排
  If Nvl(n_排班方式, 0) = 0 Then
  
    --还原上一个有效安排的终止时间
    For c_安排 In (Select Distinct a.Id, a.原终止时间
                 From 临床出诊安排 A, 临床出诊安排 B, 临床出诊表 C
                 Where a.号源id = b.号源id And a.终止时间 = b.开始时间 - 1 And a.出诊id = c.Id And c.排班方式 = 0 And c.发布人 Is Not Null And
                       b.出诊id = Id_In And a.出诊id <> Id_In) Loop
      Update 临床出诊安排
      Set 终止时间 = Nvl(c_安排.原终止时间, To_Date('3000-01-01', 'yyyy-mm-dd'))
      Where ID = c_安排.Id;
    End Loop;
  
    --删除出诊记录
    Select a.Id Bulk Collect
    Into l_记录id
    From 临床出诊记录 A, 临床出诊安排 B
    Where a.安排id = b.Id And b.出诊id = Id_In;
  
    Zl_临床出诊记录_Batchdelete(l_记录id);
  Else
    --月安排/周安排清除停诊信息，并修改是否发布
    Select a.Id Bulk Collect
    Into l_记录id
    From 临床出诊记录 A, 临床出诊安排 B
    Where a.安排id = b.Id And b.出诊id = Id_In And a.停诊开始时间 Is Not Null;
  
    Forall I In 1 .. l_记录id.Count
      Delete From 临床出诊停诊记录 Where 记录id = l_记录id(I);
  
    --修改临床出诊记录中的"是否发布"
    Select a.Id Bulk Collect
    Into l_记录id
    From 临床出诊记录 A, 临床出诊安排 B
    Where a.安排id = b.Id And b.出诊id = Id_In;
  
    Forall I In 1 .. l_记录id.Count
      Update 临床出诊记录
      Set 停诊开始时间 = Null, 停诊终止时间 = Null, 停诊原因 = Null, 是否发布 = 0
      Where ID = l_记录id(I);
  
    --换休的不再恢复
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊安排_Publish;
/

Create Or Replace Procedure Zl_临床出诊表_Update
(
  操作类型_In Number,
  Id_In       临床出诊表.Id%Type,
  出诊表名_In 临床出诊表.出诊表名%Type := Null,
  开始时间_In 临床出诊安排.开始时间%Type := Null,
  终止时间_In 临床出诊安排.终止时间%Type := Null,
  应用范围_In 临床出诊表.应用范围%Type := Null,
  科室id_In   临床出诊表.科室id%Type := Null,
  备注_In     临床出诊表.备注%Type := Null
) As
  --调整出诊表信息，针对模板和固定安排
  --操作类型_In 1-模板，2-固定安排
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_Count Number;
Begin
  --模板
  If Nvl(操作类型_In, 0) = 1 Then
    Update 临床出诊表
    Set 出诊表名 = 出诊表名_In, 应用范围 = 应用范围_In, 科室id = 科室id_In, 备注 = 备注_In
    Where ID = Id_In;
    If Sql%NotFound Then
      v_Err_Msg := '出诊表信息未找到！';
      Raise Err_Item;
    End If;
    Return;
  End If;

  --固定安排
  Begin
    Select 1 Into n_Count From 临床出诊表 Where 发布人 Is Not Null And ID = Id_In And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count <> 0 Then
    v_Err_Msg := '当前出诊表已发布，不允许进行调整！';
    Raise Err_Item;
  End If;

  Update 临床出诊表 Set 出诊表名 = 出诊表名_In Where ID = Id_In;
  If Sql%NotFound Then
    v_Err_Msg := '出诊表信息未找到！';
    Raise Err_Item;
  End If;

  Update 临床出诊安排
  Set 开始时间 = Nvl(开始时间_In, 开始时间), 终止时间 = Nvl(终止时间_In, 终止时间), 操作员姓名 = Nvl(操作员姓名, Zl_Username), 登记时间 = Nvl(登记时间, Sysdate),
      原终止时间 = Nvl(终止时间_In, 原终止时间)
  Where 出诊id = Id_In;
  If Sql%NotFound Then
    --插入一条无号源的出诊安排，用于记录出诊表的信息
    Insert Into 临床出诊安排
      (ID, 出诊id, 号源id, 开始时间, 终止时间, 操作员姓名, 登记时间, 原终止时间)
    Values
      (临床出诊安排_Id.Nextval, Id_In, Null, 开始时间_In, 终止时间_In, Zl_Username, Sysdate, 终止时间_In);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊表_Update;
/
Create Or Replace Procedure Zl_临床出诊安排_Delete
(
  Id_In       临床出诊安排.Id%Type,
  出诊记录_In Number := 0
) As
  --功能：删除临床出诊安排
  --参数：

  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  l_限制id t_Numlist := t_Numlist();
  l_记录id t_Numlist := t_Numlist();
Begin

  If Nvl(出诊记录_In, 0) = 0 Then
    --删除临床出诊规则/模板
    Select ID Bulk Collect Into l_限制id From 临床出诊限制 Where 安排id = Id_In;
  
    Forall I In 1 .. l_限制id.Count
      Delete From 临床出诊挂号控制 Where 限制id = l_限制id(I);
  
    Forall I In 1 .. l_限制id.Count
      Delete From 临床出诊时段 Where 限制id = l_限制id(I);
  
    Forall I In 1 .. l_限制id.Count
      Delete From 临床出诊诊室 Where 限制id = l_限制id(I);
  
    Delete From 临床出诊限制 Where 安排id = Id_In;
  Else
    --删除临床出诊记录
    Select ID Bulk Collect Into l_记录id From 临床出诊记录 Where 安排id = Id_In;
  
    Zl_临床出诊记录_Batchdelete(l_记录id);
  End If;
  Delete From 临床出诊安排 Where ID = Id_In;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊安排_Delete;
/
Create Or Replace Procedure Zl_临床出诊安排_序号控制
(
  出诊id_In   临床出诊表.Id%Type,
  序号控制_In 临床出诊限制.是否序号控制%Type,
  站点_In     部门表.站点%Type := Null,
  人员id_In   人员表.Id%Type := 0
) As
  --全部启用序号控制或者全部取消序号控制
  --参数：
  --      人员id_In 不等于0则修改人员所在科室的所有号源安排，否则修改所有号源的安排
  n_Count    Number(2);
  n_出诊记录 Number(2);

  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  l_安排id t_Numlist := t_Numlist();

  --该游标用于读取所有临床出诊安排的ID
  Cursor c_安排
  (
    出诊id_In 临床出诊表.Id%Type,
    人员id_In 人员表.Id%Type := 0
  ) Is
    Select b.Id
    From 临床出诊安排 B, 临床出诊号源 C
    Where b.号源id = c.Id And b.出诊id = 出诊id_In And
          (Nvl(人员id_In, 0) = 0 Or
          (Nvl(人员id_In, 0) <> 0 And Exists (Select 1 From 部门人员 Where 部门id = c.科室id And 人员id = 人员id_In))) And Exists
     (Select 1 From 部门表 Where ID = c.科室id And (站点_In Is Null Or (站点 Is Null Or 站点 = 站点_In)));
Begin
  Select Count(1)
  Into n_Count
  From 临床出诊表 A
  Where a.Id = 出诊id_In And a.发布人 Is Not Null And a.排班方式 <> 3 And Rownum < 2;
  If n_Count <> 0 Then
    v_Err_Msg := '当前出诊表已发布，不允许修改！';
    Raise Err_Item;
  End If;

  Select Count(1) Into n_Count From 临床出诊表 A Where a.Id = 出诊id_In And a.排班方式 In (1, 2) And Rownum < 2;
  If n_Count <> 0 Then
    n_出诊记录 := 1;
  End If;

  Open c_安排(出诊id_In, 人员id_In);
  Fetch c_安排 Bulk Collect
    Into l_安排id;
  Close c_安排;

  If Nvl(n_出诊记录, 0) = 0 Then
    --临床出诊限制或模板
    Forall I In 1 .. l_安排id.Count
      Update 临床出诊限制
      Set 是否序号控制 = 序号控制_In
      Where (限号数 Is Not Null Or 限约数 Is Not Null) And 安排id = l_安排id(I);
  
  Else
    --临床出诊记录
    Forall I In 1 .. l_安排id.Count
      Update 临床出诊记录
      Set 是否序号控制 = 序号控制_In
      Where (限号数 Is Not Null Or 限约数 Is Not Null) And 安排id = l_安排id(I);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊安排_序号控制;
/
Create Or Replace Procedure Zl_临床出诊上班时段_Delete
(
  安排id_In   临床出诊限制.安排id%Type,
  项目_In     临床出诊限制.限制项目%Type,
  出诊记录_In Number := 0
) As
  --功能：删除临床出诊规则/记录
  --参数：
  --      出诊记录_In:是否是对出诊记录进行删除
  --      删除出诊安排_In:删除出诊时段时是否删除安排记录
  l_限制id t_Numlist := t_Numlist();
  l_记录id t_Numlist := t_Numlist();

  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin
  If Nvl(出诊记录_In, 0) = 0 Then
    --删除临床出诊规则/模板
    Select ID Bulk Collect Into l_限制id From 临床出诊限制 Where 安排id = 安排id_In And 限制项目 = 项目_In;
  
    Forall I In 1 .. l_限制id.Count
      Delete From 临床出诊挂号控制 Where 限制id = l_限制id(I);
  
    Forall I In 1 .. l_限制id.Count
      Delete From 临床出诊时段 Where 限制id = l_限制id(I);
  
    Forall I In 1 .. l_限制id.Count
      Delete From 临床出诊诊室 Where 限制id = l_限制id(I);
  
    Forall I In 1 .. l_限制id.Count
      Delete From 临床出诊限制 Where ID = l_限制id(I);
  Else
    --删除临床出诊记录
    Select ID Bulk Collect
    Into l_记录id
    From 临床出诊记录
    Where 安排id = 安排id_In And 出诊日期 = To_Date(项目_In, 'yyyy-mm-dd');
  
    Zl_临床出诊记录_Batchdelete(l_记录id);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊上班时段_Delete;
/
Create Or Replace Procedure Zl_临床出诊安排_Insert
(
  Id_In           临床出诊安排.Id%Type,
  出诊id_In       临床出诊安排.出诊id%Type,
  号源id_In       临床出诊安排.号源id%Type,
  项目id_In       临床出诊安排.项目id%Type,
  医生id_In       临床出诊安排.医生id%Type,
  医生姓名_In     临床出诊安排.医生姓名%Type,
  排班规则_In     临床出诊安排.排班规则%Type,
  是否周六出诊_In 临床出诊安排.是否周六出诊%Type,
  是否周日出诊_In 临床出诊安排.是否周日出诊%Type,
  开始时间_In     临床出诊安排.开始时间%Type,
  终止时间_In     临床出诊安排.终止时间%Type,
  操作员姓名_In   临床出诊安排.操作员姓名%Type,
  登记时间_In     临床出诊安排.登记时间%Type
) As
  --功能：插入或更新临床出诊安排
  --参数：
  n_Count Number;

  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin

  Update 临床出诊安排
  Set 出诊id = 出诊id_In, 号源id = 号源id_In, 项目id = 项目id_In, 医生id = 医生id_In, 医生姓名 = 医生姓名_In, 排班规则 = 排班规则_In, 是否周六出诊 = 是否周六出诊_In,
      是否周日出诊 = 是否周日出诊_In, 开始时间 = 开始时间_In, 终止时间 = 终止时间_In, 操作员姓名 = 操作员姓名_In, 登记时间 = 登记时间_In, 原终止时间 = 终止时间_In
  Where ID = Id_In;
  If Sql% NotFound Then
    Insert Into 临床出诊安排
      (ID, 出诊id, 号源id, 项目id, 医生id, 医生姓名, 排班规则, 是否周六出诊, 是否周日出诊, 开始时间, 终止时间, 操作员姓名, 登记时间, 原终止时间)
    Values
      (Id_In, 出诊id_In, 号源id_In, 项目id_In, 医生id_In, 医生姓名_In, 排班规则_In, 是否周六出诊_In, 是否周日出诊_In, 开始时间_In, 终止时间_In, 操作员姓名_In,
       登记时间_In, 终止时间_In);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊安排_Insert;
/
Create Or Replace Procedure Zl_临床出诊限制_Insert
(
  Id_In           临床出诊限制.Id%Type,
  安排id_In       临床出诊限制.安排id%Type,
  限制项目_In     临床出诊限制.限制项目%Type,
  上班时段_In     临床出诊限制.上班时段%Type,
  限号数_In       临床出诊限制.限号数%Type,
  限约数_In       临床出诊限制.限约数%Type,
  是否分时段_In   临床出诊限制.是否分时段%Type,
  是否序号控制_In 临床出诊限制.是否序号控制%Type,
  预约控制_In     临床出诊限制.预约控制%Type,
  是否独占_In     临床出诊限制.是否独占%Type,
  分诊方式_In     临床出诊限制.分诊方式%Type := Null,
  诊室_In         Varchar2 := Null,
  时段_In         Varchar2 := Null,
  删除序号_In     Number := 0
) As
  --功能：插入或更新临床出诊限制
  --参数：
  --     诊室_In:诊室1,诊室2,...
  --     时段_In:序号,开始时间,终止时间,限制数量,预约标志|...
  --     删除序号_In:是否删除现有序号时段
  v_诊室 Varchar2(100);
  n_诊室 临床出诊诊室.诊室id%Type;

  v_时段     Varchar2(5000);
  n_序号     临床出诊时段.序号%Type;
  d_开始时间 临床出诊时段.开始时间%Type;
  d_终止时间 临床出诊时段.终止时间%Type;
  n_限制数量 临床出诊时段.限制数量%Type;
  n_是否预约 临床出诊时段.是否预约%Type;
  v_当前序号 Varchar2(100);

  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin

  Update 临床出诊限制
  Set 限号数 = 限号数_In, 限约数 = 限约数_In, 是否分时段 = 是否分时段_In, 是否序号控制 = 是否序号控制_In, 预约控制 = 预约控制_In, 是否独占 = 是否独占_In, 分诊方式 = 分诊方式_In,
      诊室id = Null
  Where ID = Id_In;
  If Sql% NotFound Then
    Insert Into 临床出诊限制
      (ID, 安排id, 限制项目, 上班时段, 限号数, 限约数, 是否分时段, 是否序号控制, 预约控制, 是否独占, 分诊方式)
    Values
      (Id_In, 安排id_In, 限制项目_In, 上班时段_In, 限号数_In, 限约数_In, 是否分时段_In, 是否序号控制_In, 预约控制_In, 是否独占_In, 分诊方式_In);
  End If;

  Delete From 临床出诊诊室 Where 限制id = Id_In;
  --出诊诊室
  If 诊室_In Is Not Null Then
    v_诊室 := 诊室_In || ',';
  End If;
  While v_诊室 Is Not Null Loop
    n_诊室 := To_Number(Substr(v_诊室, 1, Instr(v_诊室, ',') - 1));
    If Nvl(分诊方式_In, 0) = 1 Then
      Update 临床出诊限制 Set 诊室id = n_诊室 Where ID = Id_In;
    End If;
    Insert Into 临床出诊诊室 (限制id, 诊室id) Values (Id_In, n_诊室);
    v_诊室 := Substr(v_诊室, Instr(v_诊室, ',') + 1);
  End Loop;

  --出诊时段
  If Nvl(删除序号_In, 0) = 1 Then
    --删除现有序号时段
    Delete 临床出诊时段 Where 限制id = Id_In;
  End If;
  If 时段_In Is Not Null Then
    v_时段 := 时段_In || '|';
  End If;
  While v_时段 Is Not Null Loop
    v_当前序号 := Substr(v_时段, 1, Instr(v_时段, '|') - 1);
    n_序号     := To_Number(Substr(v_当前序号, 1, Instr(v_当前序号, ',') - 1));
    v_当前序号 := Substr(v_当前序号, Instr(v_当前序号, ',') + 1);
    d_开始时间 := To_Date(Substr(v_当前序号, 1, Instr(v_当前序号, ',') - 1), 'yyyy-mm-dd hh24:mi:ss');
    v_当前序号 := Substr(v_当前序号, Instr(v_当前序号, ',') + 1);
    d_终止时间 := To_Date(Substr(v_当前序号, 1, Instr(v_当前序号, ',') - 1), 'yyyy-mm-dd hh24:mi:ss');
    v_当前序号 := Substr(v_当前序号, Instr(v_当前序号, ',') + 1);
    n_限制数量 := To_Number(Substr(v_当前序号, 1, Instr(v_当前序号, ',') - 1));
    v_当前序号 := Substr(v_当前序号, Instr(v_当前序号, ',') + 1);
    n_是否预约 := To_Number(v_当前序号);
    If Nvl(n_序号, 0) <> 0 Then
      Insert Into 临床出诊时段
        (限制id, 序号, 开始时间, 终止时间, 是否预约, 限制数量)
      Values
        (Id_In, n_序号, d_开始时间, d_终止时间, n_是否预约, n_限制数量);
    End If;
    v_时段 := Substr(v_时段, Instr(v_时段, '|') + 1);
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊限制_Insert;
/
Create Or Replace Procedure Zl_临床出诊挂号控制_Insert
(
  限制id_In   临床出诊挂号控制.限制id%Type,
  类型_In     临床出诊挂号控制.类型%Type,
  性质_In     临床出诊挂号控制.性质%Type,
  名称_In     临床出诊挂号控制.名称%Type,
  控制方式_In 临床出诊挂号控制.控制方式%Type,
  是否独占_In 临床出诊限制.是否独占%Type,
  安排控制_In Varchar2,
  删除_In     Number := 0
) As
  --功能:插入或更新临床出诊挂号控制
  --参数:
  --    类型_In:1-三方机构;2-预约方式
  --    安排控制_in:序号1,数量|序号2,数量|...
  --    删除_in:是否删除现有的
  v_序号     Varchar2(5000);
  v_当前项目 Varchar2(5000);
  n_序号     临床出诊挂号控制.序号%Type;
  n_数量     临床出诊挂号控制.数量%Type;
  n_Count    Number;

  Err_Item Exception;
  v_Err_Msg Varchar2(100);
Begin
  If Nvl(类型_In, 0) = 1 Then
    --合作单位
    Select Count(1) Into n_Count From 挂号合作单位 Where 名称 = 名称_In;
    If n_Count = 0 Then
      v_Err_Msg := '[ZLSOFT]挂号合作单位未找到，请检查！[ZLSOFT]';
      Raise Err_Item;
    End If;
  Elsif Nvl(类型_In, 0) = 2 Then
    --预约方式
    Select Count(1) Into n_Count From 预约方式 Where 名称 = 名称_In;
    If n_Count = 0 Then
      v_Err_Msg := '[ZLSOFT]挂号预约方式未找到，请检查！[ZLSOFT]';
      Raise Err_Item;
    End If;
  End If;

  Update 临床出诊限制 Set 是否独占 = 是否独占_In Where ID = 限制id_In;

  If Nvl(删除_In, 0) = 1 Then
    --删除已有的
    Delete From 临床出诊挂号控制 Where 限制id = 限制id_In And 类型 = 类型_In And 性质 = 性质_In And 名称 = 名称_In;
  End If;

  v_序号 := 安排控制_In || '|';
  While v_序号 Is Not Null Loop
    v_当前项目 := Substr(v_序号, 1, Instr(v_序号, '|') - 1);
    n_序号     := To_Number(Substr(v_当前项目, 1, Instr(v_当前项目, ',') - 1));
    v_当前项目 := Substr(v_当前项目, Instr(v_当前项目, ',') + 1);
    n_数量     := To_Number(v_当前项目);
    If Nvl(n_数量, 0) <> 0 Then
      Insert Into 临床出诊挂号控制
        (限制id, 类型, 性质, 名称, 序号, 数量, 控制方式)
      Values
        (限制id_In, 类型_In, 性质_In, 名称_In, n_序号, n_数量, 控制方式_In);
    End If;
    v_序号 := Substr(v_序号, Instr(v_序号, '|') + 1);
  End Loop;

  --每一个合作单位或者预约方式至少得有一条记录
  Select Count(1)
  Into n_Count
  From 临床出诊挂号控制
  Where 限制id = 限制id_In And 类型 = 类型_In And 性质 = 性质_In And 名称 = 名称_In;
  If n_Count = 0 Then
    Insert Into 临床出诊挂号控制
      (限制id, 类型, 性质, 名称, 序号, 数量, 控制方式)
    Values
      (限制id_In, 类型_In, 性质_In, 名称_In, 0, 0, 控制方式_In);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊挂号控制_Insert;
/
Create Or Replace Procedure Zl_临床出诊记录_Insert
(
  Id_In           临床出诊记录.Id%Type,
  安排id_In       临床出诊记录.安排id%Type,
  号源id_In       临床出诊记录.号源id%Type,
  出诊日期_In     临床出诊记录.出诊日期%Type,
  上班时段_In     临床出诊记录.上班时段%Type,
  开始时间_In     临床出诊记录.开始时间%Type,
  终止时间_In     临床出诊记录.终止时间%Type,
  缺省预约时间_In 临床出诊记录.缺省预约时间%Type,
  提前挂号时间_In 临床出诊记录.提前挂号时间%Type,
  限号数_In       临床出诊记录.限号数%Type,
  限约数_In       临床出诊记录.限约数%Type,
  是否序号控制_In 临床出诊记录.是否序号控制%Type,
  是否分时段_In   临床出诊记录.是否分时段%Type,
  预约控制_In     临床出诊记录.预约控制%Type,
  是否独占_In     临床出诊记录.是否独占%Type,
  项目id_In       临床出诊记录.项目id%Type,
  科室id_In       临床出诊记录.科室id%Type,
  医生id_In       临床出诊记录.医生id%Type,
  医生姓名_In     临床出诊记录.医生姓名%Type,
  分诊方式_In     临床出诊记录.分诊方式%Type,
  是否临时出诊_In 临床出诊记录.是否临时出诊%Type,
  登记人_In       临床出诊记录.登记人%Type,
  登记时间_In     临床出诊记录.登记时间%Type,
  是否发布_In     临床出诊记录.是否发布%Type,
  诊室_In         Varchar2 := Null,
  时段_In         Varchar2 := Null,
  删除序号_In     Number := 0
) As
  --功能：插入或更新临床出诊记录
  --参数：
  --     诊室_In:诊室1,诊室2,...
  --     时段_In:序号,开始时间,终止时间,限制数量,预约标志|...
  --     删除序号_In:是否删除现有序号时段
  v_时段     Varchar2(5000);
  n_序号     临床出诊序号控制.序号%Type;
  d_开始时间 临床出诊序号控制.开始时间%Type;
  d_终止时间 临床出诊序号控制.终止时间%Type;
  n_数量     临床出诊序号控制.数量%Type;
  n_是否预约 临床出诊序号控制.是否预约%Type;
  v_当前序号 Varchar2(100);

  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin
  Update 临床出诊记录
  Set 安排id = 安排id_In, 号源id = 号源id_In, 出诊日期 = 出诊日期_In, 上班时段 = 上班时段_In, 开始时间 = 开始时间_In, 终止时间 = 终止时间_In, 缺省预约时间 = 缺省预约时间_In,
      提前挂号时间 = 提前挂号时间_In, 限号数 = 限号数_In, 限约数 = 限约数_In, 是否序号控制 = 是否序号控制_In, 是否分时段 = 是否分时段_In, 预约控制 = 预约控制_In,
      是否独占 = 是否独占_In, 项目id = 项目id_In, 科室id = 科室id_In, 医生id = 医生id_In, 医生姓名 = 医生姓名_In, 分诊方式 = 分诊方式_In, 是否临时出诊 = 是否临时出诊_In,
      登记人 = 登记人_In, 登记时间 = 登记时间_In, 诊室id = Null, 是否发布 = 是否发布_In
  Where ID = Id_In;
  If Sql% NotFound Then
    Insert Into 临床出诊记录
      (ID, 安排id, 号源id, 出诊日期, 上班时段, 开始时间, 终止时间, 缺省预约时间, 提前挂号时间, 限号数, 限约数, 是否序号控制, 是否分时段, 预约控制, 是否独占, 项目id, 科室id, 医生id,
       医生姓名, 分诊方式, 是否临时出诊, 登记人, 登记时间, 是否发布)
    Values
      (Id_In, 安排id_In, 号源id_In, 出诊日期_In, 上班时段_In, 开始时间_In, 终止时间_In, 缺省预约时间_In, 提前挂号时间_In, 限号数_In, 限约数_In, 是否序号控制_In,
       是否分时段_In, 预约控制_In, 是否独占_In, 项目id_In, 科室id_In, 医生id_In, 医生姓名_In, 分诊方式_In, 是否临时出诊_In, 登记人_In, 登记时间_In, 是否发布_In);
  End If;

  Delete From 临床出诊诊室记录 Where 记录id = Id_In;
  --出诊诊室
  If 诊室_In Is Not Null Then
    Insert Into 临床出诊诊室记录
      (记录id, 诊室id)
      Select Id_In, Column_Value From Table(f_Str2list(诊室_In));
  
    If Nvl(分诊方式_In, 0) = 1 Then
      Update 临床出诊记录 Set 诊室id = To_Number(诊室_In) Where ID = Id_In;
    End If;
  End If;

  --出诊时段
  If Nvl(删除序号_In, 0) = 1 Then
    --删除现有序号时段
    Delete 临床出诊序号控制 Where 记录id = Id_In;
  End If;
  If 时段_In Is Not Null Then
    v_时段 := 时段_In || '|';
  End If;
  While v_时段 Is Not Null Loop
    v_当前序号 := Substr(v_时段, 1, Instr(v_时段, '|') - 1);
    n_序号     := To_Number(Substr(v_当前序号, 1, Instr(v_当前序号, ',') - 1));
    v_当前序号 := Substr(v_当前序号, Instr(v_当前序号, ',') + 1);
    d_开始时间 := To_Date(Substr(v_当前序号, 1, Instr(v_当前序号, ',') - 1), 'yyyy-mm-dd hh24:mi:ss');
    v_当前序号 := Substr(v_当前序号, Instr(v_当前序号, ',') + 1);
    d_终止时间 := To_Date(Substr(v_当前序号, 1, Instr(v_当前序号, ',') - 1), 'yyyy-mm-dd hh24:mi:ss');
    v_当前序号 := Substr(v_当前序号, Instr(v_当前序号, ',') + 1);
    n_数量     := To_Number(Substr(v_当前序号, 1, Instr(v_当前序号, ',') - 1));
    v_当前序号 := Substr(v_当前序号, Instr(v_当前序号, ',') + 1);
    n_是否预约 := To_Number(v_当前序号);
    If Nvl(n_序号, 0) > 0 Then
      Insert Into 临床出诊序号控制
        (记录id, 序号, 开始时间, 终止时间, 是否预约, 数量)
      Values
        (Id_In, n_序号, d_开始时间, d_终止时间, n_是否预约, n_数量);
    End If;
    v_时段 := Substr(v_时段, Instr(v_时段, '|') + 1);
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊记录_Insert;
/
Create Or Replace Procedure Zl_临床出诊挂号控制记录_Insert
(
  记录id_In   临床出诊挂号控制记录.记录id%Type,
  类型_In     临床出诊挂号控制记录.类型%Type,
  性质_In     临床出诊挂号控制记录.性质%Type,
  名称_In     临床出诊挂号控制记录.名称%Type,
  控制方式_In 临床出诊挂号控制记录.控制方式%Type,
  是否独占_In 临床出诊记录.是否独占%Type,
  安排控制_In Varchar2,
  删除_In     Number := 0
) As
  --功能:插入或更新临床出诊挂号控制记录
  --参数:
  --    类型_In:1-三方机构;2-预约方式
  --    安排控制_in:序号1,数量|序号2,数量|...
  --    删除_in:是否删除现有的
  v_序号     Varchar2(5000);
  v_当前项目 Varchar2(5000);
  n_序号     临床出诊挂号控制记录.序号%Type;
  n_数量     临床出诊挂号控制记录.数量%Type;
  n_Count    Number;

  Err_Item Exception;
  v_Err_Msg Varchar2(100);
Begin
  If Nvl(类型_In, 0) = 1 Then
    --合作单位
    Select Count(1) Into n_Count From 挂号合作单位 Where 名称 = 名称_In;
    If n_Count = 0 Then
      v_Err_Msg := '[ZLSOFT]挂号合作单位未找到，请检查！[ZLSOFT]';
      Raise Err_Item;
    End If;
  Elsif Nvl(类型_In, 0) = 2 Then
    --预约方式
    Select Count(1) Into n_Count From 预约方式 Where 名称 = 名称_In;
    If n_Count = 0 Then
      v_Err_Msg := '[ZLSOFT]挂号预约方式未找到，请检查！[ZLSOFT]';
      Raise Err_Item;
    End If;
  End If;

  Update 临床出诊记录 Set 是否独占 = 是否独占_In Where ID = 记录id_In;

  If Nvl(删除_In, 0) = 1 Then
    --删除已有的
    Delete From 临床出诊挂号控制记录 Where 记录id = 记录id_In And 类型 = 类型_In And 性质 = 性质_In And 名称 = 名称_In;
  End If;

  v_序号 := 安排控制_In || '|';
  While v_序号 Is Not Null Loop
    v_当前项目 := Substr(v_序号, 1, Instr(v_序号, '|') - 1);
    n_序号     := To_Number(Substr(v_当前项目, 1, Instr(v_当前项目, ',') - 1));
    v_当前项目 := Substr(v_当前项目, Instr(v_当前项目, ',') + 1);
    n_数量     := To_Number(v_当前项目);
    If Nvl(n_数量, 0) <> 0 Then
      Insert Into 临床出诊挂号控制记录
        (记录id, 类型, 性质, 名称, 序号, 数量, 控制方式)
      Values
        (记录id_In, 类型_In, 性质_In, 名称_In, n_序号, n_数量, 控制方式_In);
    End If;
    v_序号 := Substr(v_序号, Instr(v_序号, '|') + 1);
  End Loop;

  --每一个合作单位或者预约方式至少得有一条记录
  Select Count(1)
  Into n_Count
  From 临床出诊挂号控制记录
  Where 记录id = 记录id_In And 类型 = 类型_In And 性质 = 性质_In And 名称 = 名称_In;
  If n_Count = 0 Then
    Insert Into 临床出诊挂号控制记录
      (记录id, 类型, 性质, 名称, 序号, 数量, 控制方式)
    Values
      (记录id_In, 类型_In, 性质_In, 名称_In, 0, 0, 控制方式_In);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊挂号控制记录_Insert;
/
Create Or Replace Procedure Zl_临床出诊诊室_Update
(
  Id_In       临床出诊限制.Id%Type,
  分诊方式_In 临床出诊限制.分诊方式%Type := Null,
  诊室_In     Varchar2 := Null,
  出诊记录_In Number := 0
) As
  --功能：更新临床出诊诊室
  --参数：
  --     诊室_In:诊室1,诊室2,...
  --     出诊记录_In:是否是对出诊记录进行删除
  n_Count  Number;
  n_变动id 临床出诊变动记录.Id%Type;
  v_诊室   临床出诊变动记录.现门诊诊室%Type;

  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin
  If Nvl(出诊记录_In, 0) = 0 Then
    Update 临床出诊限制 Set 分诊方式 = 分诊方式_In Where ID = Id_In;
  
    Delete From 临床出诊诊室 Where 限制id = Id_In;
    --出诊诊室
    If 诊室_In Is Not Null Then
    
      Insert Into 临床出诊诊室
        (限制id, 诊室id)
        Select Id_In, Column_Value From Table(f_Str2list(诊室_In, ','));
    
      If Nvl(分诊方式_In, 0) = 1 Then
        Update 临床出诊限制 Set 诊室id = To_Number(诊室_In) Where ID = Id_In;
      End If;
    End If;
    Return;
  End If;

  --临床出诊变动信息
  Select Count(1)
  Into n_Count
  From 临床出诊表 A, 临床出诊安排 B, 临床出诊记录 C
  Where a.Id = b.出诊id And b.Id = c.安排id And a.发布人 Is Not Null And c.Id = Id_In;
  If Nvl(n_Count, 0) = 0 Then
    v_Err_Msg := '出诊记录不存在！';
    Raise Err_Item;
  End If;

  Select 临床出诊变动记录_Id.Nextval Into n_变动id From Dual;
  Insert Into 临床出诊变动记录
    (ID, 记录id, 变动类型, 原分诊方式, 原诊室id, 原门诊诊室, 现分诊方式, 操作员姓名, 登记时间)
    Select n_变动id, a.Id, 3, a.分诊方式, a.诊室id, b.名称, 分诊方式_In, Zl_Username, Sysdate
    From 临床出诊记录 A, 门诊诊室 B
    Where a.诊室id = b.Id(+) And a.Id = Id_In;

  Insert Into 临床出诊变动明细
    (变动id, 变动性质, 序号, 诊室id, 门诊诊室, 名称)
    Select n_变动id, 1, 序号, 诊室id, 名称, '-'
    From (Select Rownum As 序号, a.诊室id, b.名称
           From 临床出诊诊室记录 A, 门诊诊室 B
           Where a.诊室id = b.Id(+) And a.记录id = Id_In);

  Update 临床出诊记录 Set 分诊方式 = 分诊方式_In Where ID = Id_In;
  Delete From 临床出诊诊室记录 Where 记录id = Id_In;

  --临床出诊变动后信息
  If 诊室_In Is Not Null Then
    Insert Into 临床出诊诊室记录
      (记录id, 诊室id)
      Select Id_In, Column_Value From Table(f_Str2list(诊室_In, ','));
  
    Insert Into 临床出诊变动明细
      (变动id, 变动性质, 序号, 诊室id, 门诊诊室, 名称)
      Select n_变动id, 2, Rownum, a.Id, a.名称, '-'
      From 门诊诊室 A, (Select Column_Value As 诊室id From Table(f_Str2list(诊室_In, ','))) B
      Where a.Id = b.诊室id;
  
    If Nvl(分诊方式_In, 0) = 1 Then
      Update 临床出诊记录 Set 诊室id = To_Number(诊室_In) Where ID = Id_In;
    
      Update 临床出诊变动记录
      Set 现诊室id = To_Number(诊室_In),
          现门诊诊室 =
           (Select 名称 From 门诊诊室 Where ID = To_Number(诊室_In))
      Where ID = n_变动id
      Returning 现门诊诊室 Into v_诊室;
      --病人挂号记录
      Update 病人挂号记录 Set 诊室 = v_诊室 Where 出诊记录id = Id_In;
      --门诊费用记录
      Update 门诊费用记录
      Set 发药窗口 = v_诊室
      Where 记录性质 = 4 And NO In (Select NO From 病人挂号记录 Where 记录性质 = 1 And 出诊记录id = Id_In);
    End If;
  
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊诊室_Update;
/
Create Or Replace Procedure Zl_临床出诊记录_Batchlock
(
  Ids_In      Varchar2,
  取消锁定_In Number := 0
) As
  -- Ids_In 批量加锁或解锁，多个用逗号分隔
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_Count Number;
Begin
  If Nvl(取消锁定_In, 0) = 1 Then
    Update 临床出诊记录 Set 是否锁定 = 0 Where ID In (Select Column_Value From Table(f_Str2list(Ids_In, ',')));
  Else
    Update 临床出诊记录 Set 是否锁定 = 1 Where ID In (Select Column_Value From Table(f_Str2list(Ids_In, ',')));
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊记录_Batchlock;
/

Create Or Replace Procedure Zl_临床出诊记录_Stopvisit
(
  记录id_In   临床出诊停诊记录.记录id%Type,
  开始时间_In 临床出诊停诊记录.开始时间%Type := Null,
  终止时间_In 临床出诊停诊记录.终止时间%Type := Null,
  停诊原因_In 临床出诊停诊记录.停诊原因%Type := Null,
  操作员_In   临床出诊停诊记录.申请人%Type := Null,
  操作时间_In 临床出诊停诊记录.申请时间%Type := Null,
  取消停诊_In Number := 0
) As
  --功能：停诊或者取消停诊
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_Count Number;
  d_Cur   Date;
  v_号码  临床出诊号源.号码%Type;
Begin
  If Nvl(取消停诊_In, 0) = 0 Then
    --停诊
    If 开始时间_In <= Sysdate Then
      v_Err_Msg := '停诊时间的开始时间小于了当前时间，不能进行停诊操作！';
      Raise Err_Item;
    End If;
  
    Insert Into 临床出诊停诊记录
      (ID, 记录id, 开始时间, 终止时间, 停诊原因, 申请人, 申请时间, 审批人, 审批时间)
      Select 临床出诊停诊记录_Id.Nextval, 记录id_In, 开始时间_In, 终止时间_In, 停诊原因_In, Nvl(a.医生姓名, 操作员_In), 操作时间_In, 操作员_In, 操作时间_In
      From 临床出诊记录 A
      Where ID = 记录id_In;
  
    Update 临床出诊记录
    Set 停诊开始时间 = 开始时间_In, 停诊终止时间 = 终止时间_In, 停诊原因 = 停诊原因_In
    Where ID = 记录id_In;
  
    Insert Into 病人服务信息记录
      (ID, 通知类型, 记录id, 挂号id, 号源id, 号码, 科室id, 项目id, 医生id, 医生姓名, 病人id, 登记人, 登记时间, 通知原因)
      Select 病人服务信息记录_Id.Nextval, 1, 记录id_In, 挂号id, 号源id, 号码, 科室id, 项目id, 医生id, 医生姓名, 病人id, 操作员_In, 操作时间_In,
             '医生' || 停诊原因_In || '，已停诊'
      From (Select b.Id As 挂号id, c.Id As 号源id, c.号码, c.科室id, a.项目id, a.医生id, a.医生姓名, b.病人id
             From 临床出诊记录 A, 病人挂号记录 B, 临床出诊号源 C
             Where a.Id = b.出诊记录id And a.号源id = c.Id And b.记录状态 = 1 And a.Id = 记录id_In And
                   (b.记录性质 = 1 And b.发生时间 Between a.停诊开始时间 And a.停诊终止时间 Or
                   b.记录性质 = 2 And b.预约时间 Between a.停诊开始时间 And a.停诊终止时间));
  
    --消息推送
    -- 停诊类型(1-停诊,2-取消停诊),出诊记录ID,停诊号码
    Begin
      Select b.号码 Into v_号码 From 临床出诊记录 A, 临床出诊号源 B Where a.号源id = b.Id And a.Id = 记录id_In;
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 17, 1 || ',' || 记录id_In || ',' || v_号码;
    Exception
      When Others Then
        Null;
    End;
  Else
    --取消停诊
    --数据检查
    Select 停诊开始时间 Into d_Cur From 临床出诊记录 Where ID = 记录id_In And 停诊开始时间 Is Not Null;
    If d_Cur <= Sysdate Then
      v_Err_Msg := '停诊时间的开始时间已小于了当前时间，不能进行取消停诊操作！';
      Raise Err_Item;
    End If;
    Select Count(1)
    Into n_Count
    From 病人服务信息记录
    Where 记录id = 记录id_In And 通知类型 = 1 And 处理人 Is Not Null;
    If n_Count <> 0 Then
      v_Err_Msg := '该出诊记录存在病人服务信息信息记录，且已被处理，不允许取消停诊操作！';
      Raise Err_Item;
    End If;
  
    Update 临床出诊停诊记录
    Set 取消人 = 操作员_In, 取消时间 = 操作时间_In
    Where 记录id = 记录id_In And 替诊医生姓名 Is Null And 取消人 Is Null;
  
    Update 临床出诊记录
    Set 停诊开始时间 = Null, 停诊终止时间 = Null, 停诊原因 = Null
    Where ID = 记录id_In And 停诊开始时间 Is Not Null;
  
    Delete 病人服务信息记录 Where 记录id = 记录id_In And 通知类型 = 1 And 处理人 Is Null;
  
    --消息推送
    -- 停诊类型(1-停诊,2-取消停诊),出诊记录ID,停诊号码
    Begin
      Select b.号码 Into v_号码 From 临床出诊记录 A, 临床出诊号源 B Where a.号源id = b.Id And a.Id = 记录id_In;
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 17, 2 || ',' || 记录id_In || ',' || v_号码;
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
End Zl_临床出诊记录_Stopvisit;
/


Create Or Replace Function Zl1_Ex_Isdoctorsamelevel
(
  甲医生id_In   In 人员表.Id%Type,
  甲医生姓名_In In 人员表.姓名%Type,
  乙医生id_In   In 人员表.Id%Type,
  乙医生姓名_In In 人员表.姓名%Type
) Return Number
--功能说明：比较两个医生的职务大小。
  --适用说明：挂号安排替诊时调用，检查替诊医生的职务是否大于等于原医生职务。
  --入参说明：
  --     甲医生ID：人员ID,院外医生传入NULL
  --     甲医生姓名：人员姓名
  --     乙医生ID：人员ID,院外医生传入NULL
  --     乙医生姓名：人员姓名
  --函数返回：
  --     -1 - 甲医生的职务大于乙医生的职务
  --     0 - 甲医生的职务等于乙医生的职务
  --     1 - 甲医生的职务小于乙医生的职务
  --说明：根据“专业技术职务”来判断,编码小的表示职务越大,没有设置专业技术职务的医生表示职务最低
 Is
  n_a Number;
  n_b Number;
Begin
  If Nvl(甲医生id_In, 0) = 0 Then
    --院外医生
    n_a := -1;
  Else
    Begin
      Select To_Number(Nvl(b.编码, -1))
      Into n_a
      From 人员表 A, 专业技术职务 B
      Where a.专业技术职务 = b.名称(+) And a.Id = 甲医生id_In;
    Exception
      When Others Then
        n_a := -1;
    End;
  End If;
  If Nvl(乙医生id_In, 0) = 0 Then
    --院外医生
    n_b := -1;
  Else
    Begin
      Select To_Number(Nvl(b.编码, -1))
      Into n_b
      From 人员表 A, 专业技术职务 B
      Where a.专业技术职务 = b.名称(+) And a.Id = 乙医生id_In;
    Exception
      When Others Then
        n_b := -1;
    End;
  End If;

  If n_a = -1 And n_b = -1 Then
    Return 0;
  Elsif n_a = -1 Then
    Return 1;
  Elsif n_b = -1 Then
    Return - 1;
  Else
    If n_a = n_b Then
      Return 0;
    Elsif n_a > n_b Then
      Return 1;
    Else
      Return - 1;
    End If;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl1_Ex_Isdoctorsamelevel;
/

Create Or Replace Procedure Zl_临床出诊记录_Replacedoctor
(
  记录id_In       临床出诊停诊记录.记录id%Type,
  开始时间_In     临床出诊停诊记录.开始时间%Type := Null,
  终止时间_In     临床出诊停诊记录.终止时间%Type := Null,
  停诊原因_In     临床出诊停诊记录.停诊原因%Type := Null,
  替诊医生id_In   临床出诊停诊记录.替诊医生id%Type := Null,
  替诊医生姓名_In 临床出诊停诊记录.替诊医生姓名%Type := Null,
  操作员姓名_In   临床出诊停诊记录.申请人%Type := Null,
  操作员编号_In   人员表.编号%Type := Null,
  操作时间_In     临床出诊停诊记录.申请时间%Type := Null,
  取消替诊_In     Number := 0
) As
  --功能：替诊或者取消替诊
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_Count        Number;
  d_Cur          Date;
  n_Updatedoctor Number(2);
  v_号码         临床出诊号源.号码%Type;
Begin
  If Nvl(取消替诊_In, 0) = 0 Then
    --替诊
    Begin
      Select 1 Into n_Count From 临床出诊记录 A Where ID = 记录id_In And 停诊开始时间 Is Not Null;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If Nvl(n_Count, 0) <> 0 Then
      v_Err_Msg := '当前出诊记录已被停诊，不允许替诊！';
      Raise Err_Item;
    End If;
  
    If 开始时间_In <= Sysdate Then
      v_Err_Msg := '停诊时间的开始时间小于了当前时间，不能进行替诊操作！';
      Raise Err_Item;
    End If;
  
    If Nvl(替诊医生id_In, 0) <> 0 Then
      Begin
        Select 1
        Into n_Count
        From 临床出诊记录 A
        Where ID = 记录id_In And Nvl(医生id, 替诊医生id) = 替诊医生id_In And Rownum < 2;
      Exception
        When Others Then
          n_Count := 0;
      End;
      If Nvl(n_Count, 0) <> 0 Then
        v_Err_Msg := '替诊医生不能为原安排医生，请选择其它医生！';
        Raise Err_Item;
      End If;
    End If;
  
    --在该时段内，替诊医生不能存在其他的出诊安排
    --若A[A1,A2],B[B1,B2],且B为空或完全包含于A中(A1<=B1,A2>=B2).那么X[X1,X2]与A-B有交集，则
    --(X1>=A1 And X1<=NVL(B1,A2)) Or (X2>=A1 And X2<=NVL(B1,A2)) Or (X1>=NVL(B2,A1) And X1<=A2) Or (X2>=NVL(B2,A1) And X2<=A2)
    If Nvl(替诊医生id_In, 0) = 0 Then
      Select Count(1)
      Into n_Count
      From 临床出诊记录 A, 临床出诊记录 B
      Where a.出诊日期 = b.出诊日期 And Nvl(a.替诊医生姓名, a.医生姓名) = 替诊医生姓名_In And Nvl(a.替诊医生id, a.医生id) Is Null And b.Id = 记录id_In And
            ((开始时间_In Between a.开始时间 And Nvl(a.停诊开始时间, a.终止时间)) Or (终止时间_In Between a.开始时间 And Nvl(a.停诊开始时间, a.终止时间)) Or
            (开始时间_In Between Nvl(a.停诊终止时间, a.开始时间) And a.终止时间) Or (终止时间_In Between Nvl(a.停诊终止时间, a.开始时间) And a.终止时间));
    Else
      Select Count(1)
      Into n_Count
      From 临床出诊记录 A, 临床出诊记录 B
      Where a.出诊日期 = b.出诊日期 And Nvl(a.替诊医生id, a.医生id) = 替诊医生id_In And b.Id = 记录id_In And
            ((开始时间_In Between a.开始时间 And Nvl(a.停诊开始时间, a.终止时间)) Or (终止时间_In Between a.开始时间 And Nvl(a.停诊开始时间, a.终止时间)) Or
            (开始时间_In Between Nvl(a.停诊终止时间, a.开始时间) And a.终止时间) Or (终止时间_In Between Nvl(a.停诊终止时间, a.开始时间) And a.终止时间));
    End If;
    If n_Count <> 0 Then
      v_Err_Msg := '替诊医生在替诊时间范围内已存在其它出诊安排，请选择其它医生！';
      Raise Err_Item;
    End If;
    --必须为同级别以上的医生
    Select Zl1_Ex_Isdoctorsamelevel(a.医生id, a.医生姓名, 替诊医生id_In, 替诊医生姓名_In)
    Into n_Count
    From 临床出诊记录 A
    Where ID = 记录id_In;
    If n_Count = -1 Then
      v_Err_Msg := '替诊医生的级别小于了原出诊医生的级别，不允许替诊，请选择其它医生！';
      Raise Err_Item;
    End If;
  
    Insert Into 临床出诊停诊记录
      (ID, 记录id, 开始时间, 终止时间, 停诊原因, 替诊医生id, 替诊医生姓名, 申请人, 申请时间, 审批人, 审批时间)
      Select 临床出诊停诊记录_Id.Nextval, 记录id_In, 开始时间_In, 终止时间_In, 停诊原因_In, 替诊医生id_In, 替诊医生姓名_In, Nvl(a.医生姓名, 操作员姓名_In),
             操作时间_In, 操作员姓名_In, 操作时间_In
      From 临床出诊记录 A
      Where ID = 记录id_In;
  
    Update 临床出诊记录 Set 替诊医生id = 替诊医生id_In, 替诊医生姓名 = 替诊医生姓名_In Where ID = 记录id_In;
  
    Insert Into 病人服务信息记录
      (ID, 通知类型, 记录id, 挂号id, 号源id, 号码, 科室id, 项目id, 医生id, 医生姓名, 病人id, 登记人, 登记时间, 通知原因)
      Select 病人服务信息记录_Id.Nextval, 2, 记录id_In, 挂号id, 号源id, 号码, 科室id, 项目id, 医生id, 医生姓名, 病人id, 操作员姓名_In, 操作时间_In,
             '医生' || 停诊原因_In || '，已替诊'
      From (Select b.Id As 挂号id, c.Id As 号源id, c.号码, c.科室id, a.项目id, a.医生id, a.医生姓名, b.病人id
             From 临床出诊记录 A, 病人挂号记录 B, 临床出诊号源 C
             Where a.Id = b.出诊记录id And a.号源id = c.Id And b.记录状态 = 1 And a.Id = 记录id_In And
                   (b.记录性质 = 1 And b.发生时间 Between a.开始时间 And a.终止时间 Or b.记录性质 = 2 And b.预约时间 Between a.开始时间 And a.终止时间));
  
    --消息推送
    -- 替诊类型(1-替诊,2-取消替诊),出诊记录ID,替诊号码
    Begin
      Select b.号码 Into v_号码 From 临床出诊记录 A, 临床出诊号源 B Where a.号源id = b.Id And a.Id = 记录id_In;
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 18, 1 || ',' || 记录id_In || ',' || v_号码;
    Exception
      When Others Then
        Null;
    End;
  
    --按替诊医生同步更新预约挂号单
    n_Updatedoctor := zl_GetSysParameter('按替诊医生同步更新预约挂号单', 1114);
    If Nvl(n_Updatedoctor, 0) = 1 Then
      For c_记录 In (Select a.Id, b.No
                   From 病人服务信息记录 A, 病人挂号记录 B
                   Where a.挂号id = b.Id And a.记录id = 记录id_In And a.通知类型 = 2 And b.记录性质 In (1, 2) And b.记录状态 = 1) Loop
        Zl_患者服务中心_替诊(c_记录.Id, c_记录.No, '按替诊医生同步更新预约挂号单', 操作员姓名_In, 操作员编号_In);
      End Loop;
    End If;
  Else
    --数据检查
    Select 终止时间
    Into d_Cur
    From 临床出诊记录
    Where ID = 记录id_In And 替诊医生姓名 Is Not Null And 停诊开始时间 Is Null;
    If d_Cur <= Sysdate Then
      v_Err_Msg := '终止时间已小于了当前时间，不能进行取消替诊操作！';
      Raise Err_Item;
    End If;
    Select Count(1)
    Into n_Count
    From 病人服务信息记录
    Where 记录id = 记录id_In And 通知类型 = 2 And 处理人 Is Not Null;
    If n_Count <> 0 Then
      v_Err_Msg := '该出诊记录存在病人服务信息信息记录，且已被处理，不允许取消替诊操作！';
      Raise Err_Item;
    End If;
  
    Update 临床出诊停诊记录
    Set 取消人 = 操作员姓名_In, 取消时间 = 操作时间_In
    Where 记录id = 记录id_In And 替诊医生姓名 Is Not Null And 取消人 Is Null;
  
    Update 临床出诊记录
    Set 替诊医生id = Null, 替诊医生姓名 = Null
    Where ID = 记录id_In And 替诊医生姓名 Is Not Null And 停诊开始时间 Is Null;
  
    Delete 病人服务信息记录 Where 记录id = 记录id_In And 通知类型 = 2 And 处理人 Is Null;
  
    --消息推送
    -- 替诊类型(1-替诊,2-取消替诊),出诊记录ID,替诊号码
    Begin
      Select b.号码 Into v_号码 From 临床出诊记录 A, 临床出诊号源 B Where a.号源id = b.Id And a.Id = 记录id_In;
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 18, 2 || ',' || 记录id_In || ',' || v_号码;
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
End Zl_临床出诊记录_Replacedoctor;
/
Create Or Replace Procedure Zl_临床出诊停诊_Apply
(
  操作类型_In Number,
  Id_In       临床出诊停诊记录.Id%Type,
  开始时间_In 临床出诊停诊记录.开始时间%Type,
  终止时间_In 临床出诊停诊记录.终止时间%Type,
  停诊原因_In 临床出诊停诊记录.停诊原因%Type,
  申请人_In   临床出诊停诊记录.申请人%Type,
  申请时间_In 临床出诊停诊记录.申请时间%Type
) As
  --功能：退费申请以及取消申请
  --参数：
  --        操作类型_In：0-申请，else-取消申请
  --说明：
  n_Id    临床出诊停诊记录.Id%Type;
  n_Count Number;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  If 操作类型_In = 0 Then
    --申请
    Begin
      Select 1
      Into n_Count
      From 临床出诊停诊记录
      Where 记录id Is Null And Not (开始时间 > 终止时间_In Or 终止时间 < 开始时间_In) And 申请人 = 申请人_In And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If Nvl(n_Count, 0) > 0 Then
      v_Error := '医生 ' || 申请人_In || ' 在当前停诊时间范围内已存在停诊安排，不能重复申请！';
      Raise Err_Custom;
    End If;
  
    If Nvl(Id_In, 0) = 0 Then
      Select 临床出诊停诊记录_Id.Nextval Into n_Id From Dual;
    End If;
  
    Insert Into 临床出诊停诊记录
      (ID, 开始时间, 终止时间, 停诊原因, 申请人, 申请时间)
    Values
      (n_Id, 开始时间_In, 终止时间_In, 停诊原因_In, 申请人_In, 申请时间_In);
  Else
    --取消申请
    Select Count(1) Into n_Count From 临床出诊停诊记录 Where ID = Id_In;
    If Nvl(n_Count, 0) = 0 Then
      v_Error := '该申请已被取消申请，请刷新后查看...';
      Raise Err_Custom;
    End If;
  
    --审核通过，不允许取消申请
    Select Count(1) Into n_Count From 临床出诊停诊记录 Where ID = Id_In And (审批人 Is Null Or 取消人 Is Not Null);
    If Nvl(n_Count, 0) = 0 Then
      v_Error := '该申请已被审核，不能取消申请。';
      Raise Err_Custom;
    End If;
  
    Delete 临床出诊停诊记录 Where ID = Id_In;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊停诊_Apply;
/
Create Or Replace Procedure Zl_临床出诊停诊_Audit
(
  操作类型_In Number,
  Id_In       临床出诊停诊记录.Id%Type,
  审批人_In   临床出诊停诊记录.审批人%Type,
  审批时间_In 临床出诊停诊记录.审批时间%Type
) As
  --功能：审核停诊安排
  --参数：
  --       状态_In：1-审核，2-取消审核
  n_Count Number;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  If Nvl(操作类型_In, 0) = 1 Then
    --审核
    Select Count(1) Into n_Count From 临床出诊停诊记录 Where ID = Id_In;
    If Nvl(n_Count, 0) = 0 Then
      v_Error := '该申请已被取消申请，请刷新后查看...';
      Raise Err_Custom;
    End If;
  
    Select Count(1) Into n_Count From 临床出诊停诊记录 Where ID = Id_In And (审批人 Is Null Or 取消人 Is Not Null);
    If Nvl(n_Count, 0) = 0 Then
      v_Error := '该申请已被审核，不能再次审核！';
      Raise Err_Custom;
    End If;
  
    Update 临床出诊停诊记录
    Set 审批人 = 审批人_In, 审批时间 = 审批时间_In, 取消人 = Null, 取消时间 = Null
    Where ID = Id_In;
  
    --对出诊记录进行停诊标记
    For c_记录 In (Select a.Id,
                        Case
                          When a.开始时间 < b.开始时间 Then
                           b.开始时间
                          Else
                           a.开始时间
                        End As 停诊开始时间,
                        Case
                          When a.终止时间 > b.终止时间 Then
                           b.终止时间
                          Else
                           a.终止时间
                        End As 停诊终止时间, b.停诊原因, c.号码
                 From 临床出诊记录 A, 临床出诊停诊记录 B, 临床出诊号源 C
                 Where ((a.替诊医生姓名 Is Null And a.医生id Is Not Null And a.医生姓名 = b.申请人) Or
                       (a.替诊医生姓名 Is Not Null And a.替诊医生id Is Not Null And a.替诊医生姓名 = b.申请人)) And a.号源id = c.Id And
                       b.Id = Id_In And Not (a.开始时间 > b.终止时间 Or a.终止时间 < b.开始时间)
                      --只处理已发布了的
                       And Exists (Select 1
                        From 临床出诊安排 C, 临床出诊表 D
                        Where c.出诊id = d.Id And c.Id = a.安排id And d.发布时间 Is Not Null)) Loop
    
      Update 临床出诊记录
      Set 停诊开始时间 = c_记录.停诊开始时间, 停诊终止时间 = c_记录.停诊终止时间, 停诊原因 = c_记录.停诊原因
      Where ID = c_记录.Id;
    
      Insert Into 病人服务信息记录
        (ID, 通知类型, 记录id, 挂号id, 号源id, 号码, 科室id, 项目id, 医生id, 医生姓名, 病人id, 登记人, 登记时间)
        Select 病人服务信息记录_Id.Nextval, 1, a.Id, b.Id, c.Id, c.号码, c.科室id, a.项目id, a.医生id, a.医生姓名, b.病人id, 审批人_In, 审批时间_In
        From 临床出诊记录 A, 病人挂号记录 B, 临床出诊号源 C
        Where a.Id = b.出诊记录id And a.号源id = c.Id And b.记录状态 = 1 And a.Id = c_记录.Id And
              (b.记录性质 = 1 And b.发生时间 Between a.停诊开始时间 And a.停诊终止时间 Or
              b.记录性质 = 2 And b.预约时间 Between a.停诊开始时间 And a.停诊终止时间);
    
      --消息推送
      -- 停诊类型(1-停诊,2-取消停诊),出诊记录ID,停诊号码
      Begin
        Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
          Using 17, 1 || ',' || c_记录.Id || ',' || c_记录.号码;
      Exception
        When Others Then
          Null;
      End;
    End Loop;
  Else
    --取消审核
    Select Count(1) Into n_Count From 临床出诊停诊记录 Where ID = Id_In And 审批人 Is Not Null And 取消人 Is Null;
    If Nvl(n_Count, 0) = 0 Then
      v_Error := '原审核记录未找到，请刷新后查看...';
      Raise Err_Custom;
    End If;
  
    Select Count(1)
    Into n_Count
    From 病人服务信息记录
    Where 记录id In (Select a.Id
                   From 临床出诊记录 A, 临床出诊停诊记录 B
                   Where Nvl(a.替诊医生姓名, a.医生姓名) = b.申请人 And b.Id = Id_In And
                         (a.开始时间 Between b.开始时间 And b.终止时间 Or a.终止时间 Between b.开始时间 And b.终止时间)) And 处理人 Is Not Null;
    If Nvl(n_Count, 0) <> 0 Then
      v_Error := '该停诊安排的部分停诊信息已被处理，不能取消审批！';
      Raise Err_Custom;
    End If;
  
    Update 临床出诊停诊记录 Set 取消人 = 审批人_In, 取消时间 = 审批时间_In Where ID = Id_In;
  
    For c_记录 In (Select a.Id, c.号码
                 From 临床出诊记录 A, 临床出诊停诊记录 B, 临床出诊号源 C
                 Where ((a.替诊医生姓名 Is Null And a.医生id Is Not Null And a.医生姓名 = b.申请人) Or
                       (a.替诊医生姓名 Is Not Null And a.替诊医生id Is Not Null And a.替诊医生姓名 = b.申请人)) And a.号源id = c.Id And
                       b.Id = Id_In And (a.开始时间 Between b.开始时间 And b.终止时间 Or a.终止时间 Between b.开始时间 And b.终止时间) And
                       Exists (Select 1
                        From 临床出诊安排 C, 临床出诊表 D
                        Where c.出诊id = d.Id And c.Id = a.安排id And d.发布时间 Is Not Null)) Loop
    
      Update 临床出诊记录 Set 停诊开始时间 = Null, 停诊终止时间 = Null, 停诊原因 = Null Where ID = c_记录.Id;
    
      Delete 病人服务信息记录 Where 记录id = c_记录.Id And 通知类型 = 1 And 处理人 Is Null;
    
      --消息推送
      -- 停诊类型(1-停诊,2-取消停诊),出诊记录ID,停诊号码
      Begin
        Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
          Using 17, 2 || ',' || c_记录.Id || ',' || c_记录.号码;
      Exception
        When Others Then
          Null;
      End;
    End Loop;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊停诊_Audit;
/

Create Or Replace Procedure Zl_临床出诊预约控制变动
(
  变动性质_In   临床出诊变动明细.变动性质%Type,
  Id_In         临床出诊变动记录.Id%Type,
  记录id_In     临床出诊变动记录.记录id%Type,
  现预约控制_In 临床出诊变动记录.现预约控制%Type := Null
) As
  --功能:修改预约控制时，插入临床出诊变动记录/明细
  --参数:
  --     变动性质_In  1-变动前;2-变动后
  Err_Item Exception;
  v_Err_Msg Varchar2(100);
Begin
  If Nvl(变动性质_In, 0) = 1 Then
    --变动前
    Insert Into 临床出诊变动记录
      (ID, 记录id, 变动类型, 原预约控制, 现预约控制, 操作员姓名, 登记时间)
      Select Id_In, 记录id_In, 4, 预约控制, Nvl(现预约控制_In, 预约控制), Zl_Username, Sysdate
      From 临床出诊记录
      Where ID = 记录id_In;
  
    Insert Into 临床出诊变动明细
      (变动id, 变动性质, 类型, 名称, 序号, 控制方式, 数量)
      Select Id_In, 1, 类型, 名称, 序号, 控制方式, 数量 From 临床出诊挂号控制记录 Where 记录id = 记录id_In;
  Else
    --变动后
    Insert Into 临床出诊变动明细
      (变动id, 变动性质, 类型, 名称, 序号, 控制方式, 数量)
      Select Id_In, 2, 类型, 名称, 序号, 控制方式, 数量 From 临床出诊挂号控制记录 Where 记录id = 记录id_In;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊预约控制变动;
/

Create Or Replace Procedure Zl_临床出诊序号控制变动
(
  记录id_In     临床出诊变动记录.记录id%Type,
  限号数_In     临床出诊记录.限号数%Type,
  限约数_In     临床出诊记录.限约数%Type,
  操作员姓名_In 临床出诊变动记录.操作员姓名%Type := Null,
  登记时间_In   临床出诊变动记录.登记时间%Type := Null
) As
  --功能:修改临床出诊序号控制时，插入临床出诊变动记录/明细
  --参数:
  n_原限号数 临床出诊记录.限号数%Type;
  n_原限约数 临床出诊记录.限约数%Type;

  v_操作员姓名 临床出诊变动记录.操作员姓名%Type := Null;
  d_登记时间   临床出诊变动记录.登记时间%Type := Null;

  Err_Item Exception;
  v_Err_Msg Varchar2(100);
Begin
  Begin
    Select 限号数, 限约数 Into n_原限号数, n_原限约数 From 临床出诊记录 Where ID = 记录id_In;
  Exception
    When Others Then
      v_Err_Msg := '未发现出诊记录！';
      Raise Err_Item;
  End;

  --调整限约，限号数，且限约数为零表示禁止预约
  Update 临床出诊记录
  Set 限约数 = 限约数_In, 限号数 = 限号数_In, 预约控制 = Decode(Nvl(限约数_In, 0), 0, 1, 预约控制)
  Where ID = 记录id_In;

  v_操作员姓名 := Nvl(操作员姓名_In, Zl_Username);
  d_登记时间   := Nvl(登记时间_In, Sysdate);
  If Nvl(n_原限号数, 0) <> Nvl(限号数_In, 0) Then
    Insert Into 临床出诊变动记录
      (ID, 记录id, 变动类型, 原数量, 现数量, 操作员姓名, 登记时间)
    Values
      (临床出诊变动记录_Id.Nextval, 记录id_In, 1, n_原限号数, 限号数_In, v_操作员姓名, d_登记时间);
  End If;
  If Nvl(n_原限约数, 0) <> Nvl(限约数_In, 0) Then
    Insert Into 临床出诊变动记录
      (ID, 记录id, 变动类型, 原数量, 现数量, 操作员姓名, 登记时间)
    Values
      (临床出诊变动记录_Id.Nextval, 记录id_In, 2, n_原限约数, 限约数_In, v_操作员姓名, d_登记时间);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊序号控制变动;
/
Create Or Replace Procedure Zl_临床出诊序号控制_Update
(
  记录id_In   临床出诊记录.Id%Type,
  时段_In     Varchar2 := Null,
  删除序号_In Number := 0
) As
  --功能：更新临床出诊序号
  --参数：
  --     时段_In:序号,开始时间,终止时间,限制数量,预约标志|...
  --     删除序号_In:是否删除现有序号时段
  n_序号     临床出诊序号控制.序号%Type;
  d_开始时间 临床出诊序号控制.开始时间%Type;
  d_终止时间 临床出诊序号控制.终止时间%Type;
  n_数量     临床出诊序号控制.数量%Type;
  n_是否预约 临床出诊序号控制.是否预约%Type;

  n_时段序号 t_Numlist := t_Numlist();

  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin
  If Nvl(删除序号_In, 0) = 1 Then
    --删除本次没有的现有序号时段
    If 时段_In Is Not Null Then
      For c_时段集 In (Select Rownum As 序号, Column_Value As 值 From Table(f_Str2list(时段_In, '|'))) Loop
        --序号,开始时间,终止时间,限制数量,预约标志
        For c_时段 In (Select Column_Value As 值 From Table(f_Str2list(c_时段集.值)) Where Rownum = 1) Loop
        
          n_时段序号.Extend();
          n_时段序号(n_时段序号.Count) := To_Number(c_时段.值);
        End Loop;
      End Loop;
    End If;
  
    Delete 临床出诊序号控制 Where 记录id = 记录id_In And 序号 Not In (Select Column_Value From Table(n_时段序号));
    v_Err_Msg := n_时段序号.Count;
  End If;

  If 时段_In Is Not Null Then
    For c_时段集 In (Select Rownum As 序号, Column_Value As 值 From Table(f_Str2list(时段_In, '|'))) Loop
      --序号,开始时间,终止时间,限制数量,预约标志
      For c_时段 In (Select Rownum As 序号, Column_Value As 值 From Table(f_Str2list(c_时段集.值))) Loop
        If c_时段.序号 = 1 Then
          n_序号 := To_Number(c_时段.值);
        End If;
      
        If c_时段.序号 = 2 Then
          d_开始时间 := To_Date(c_时段.值, 'yyyy-mm-dd hh24:mi:ss');
        End If;
      
        If c_时段.序号 = 3 Then
          d_终止时间 := To_Date(c_时段.值, 'yyyy-mm-dd hh24:mi:ss');
        End If;
      
        If c_时段.序号 = 4 Then
          n_数量 := To_Number(c_时段.值);
        End If;
      
        If c_时段.序号 = 5 Then
          n_是否预约 := To_Number(c_时段.值);
        End If;
      End Loop;
    
      If Nvl(n_序号, 0) <> 0 Then
        Update 临床出诊序号控制
        Set 开始时间 = d_开始时间, 终止时间 = d_终止时间, 是否预约 = n_是否预约, 数量 = n_数量
        Where 记录id = 记录id_In And 序号 = n_序号;
        If Sql%NotFound Then
          Insert Into 临床出诊序号控制
            (记录id, 序号, 开始时间, 终止时间, 是否预约, 数量)
          Values
            (记录id_In, n_序号, d_开始时间, d_终止时间, n_是否预约, n_数量);
        End If;
      End If;
    End Loop;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊序号控制_Update;
/


Create Or Replace Procedure Zl_患者服务中心_替诊
(
  消息id_In     病人服务信息记录.Id%Type,
  No_In         病人挂号记录.No%Type,
  处理说明_In   病人服务信息记录.处理说明%Type,
  操作员姓名_In 病人服务信息记录.处理人%Type,
  操作员编号_In 病人挂号记录.操作员编号%Type
) As
  v_原执行人   病人挂号记录.执行人%Type;
  v_执行人     病人挂号记录.执行人%Type;
  n_原执行人id 临床出诊记录.医生id%Type;
  n_执行人id   临床出诊记录.替诊医生id%Type;
  d_出诊日期   临床出诊记录.出诊日期%Type;
  n_换诊序号   病人挂号记录.号序%Type;
  n_项目id     临床出诊记录.项目id%Type;
  n_科室id     临床出诊号源.科室id%Type;
  n_挂号状态   Number(3); --1=挂号,2=预约未收款,3=预约已收款
  v_号码       临床出诊号源.号码%Type;
  n_变动id     就诊变动记录.Id%Type;
  v_Err_Msg    Varchar2(500);
  Err_Item Exception;
Begin
  --获取替诊医生
  Select b.替诊医生姓名, b.医生姓名, b.医生id, b.替诊医生id, 出诊日期, b.项目id, c.科室id, c.号码
  Into v_执行人, v_原执行人, n_原执行人id, n_执行人id, d_出诊日期, n_项目id, n_科室id, v_号码
  From 病人服务信息记录 A, 临床出诊记录 B, 临床出诊号源 C
  Where a.Id = 消息id_In And a.记录id = b.Id And b.号源id = c.Id;

  Select Decode(Nvl(预约, 0), 0, 1, Decode(接收时间, Null, 2, 3)), 号序
  Into n_挂号状态, n_换诊序号
  From 病人挂号记录
  Where NO = No_In And 记录状态 = 1;

  --就诊变动记录
  Select 就诊变动记录_Id.Nextval Into n_变动id From Dual;
  b_Message.Zlhis_Regist_005(No_In, n_变动id, 1);
  Zl_就诊变动记录_Insert(No_In, 5, '患者服务中心替诊', 操作员姓名_In, 操作员编号_In, v_号码, n_科室id, n_项目id, n_执行人id, v_执行人, Null, n_换诊序号,
                   Sysdate, n_变动id);

  --更新病人挂号记录
  Update 病人挂号记录 Set 执行人 = v_执行人 Where NO = No_In And 记录状态 = 1;
  --更新门诊费用记录
  Update 门诊费用记录 Set 执行人 = v_执行人 Where NO = No_In And 记录性质 = 4;
  --更新患者服务记录
  Update 病人服务信息记录 Set 处理人 = 操作员姓名_In, 处理时间 = Sysdate, 处理说明 = 处理说明_In Where ID = 消息id_In;
  --更新病人挂号汇总
  If n_挂号状态 = 1 Then
    Update 病人挂号汇总
    Set 已挂数 = Nvl(已挂数, 0) - 1
    Where 日期 = Trunc(d_出诊日期) And Nvl(医生姓名, '-') = Nvl(v_原执行人, '-') And Nvl(医生id, 0) = Nvl(n_原执行人id, 0) And
          Nvl(科室id, 0) = Nvl(n_科室id, 0) And Nvl(项目id, 0) = Nvl(n_项目id, 0) And (号码 = v_号码 Or 号码 Is Null);
    Update 病人挂号汇总
    Set 已挂数 = Nvl(已挂数, 0) + 1
    Where 日期 = Trunc(d_出诊日期) And Nvl(医生姓名, '-') = Nvl(v_执行人, '-') And Nvl(医生id, 0) = Nvl(n_执行人id, 0) And
          Nvl(科室id, 0) = Nvl(n_科室id, 0) And Nvl(项目id, 0) = Nvl(n_项目id, 0) And (号码 = v_号码 Or 号码 Is Null);
    If Sql%RowCount = 0 Then
      Insert Into 病人挂号汇总
        (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 已约数, 其中已接收)
      Values
        (Trunc(d_出诊日期), n_科室id, n_项目id, v_执行人, Decode(n_执行人id, 0, Null, n_执行人id), v_号码, 1, 0, 0);
    End If;
  End If;
  If n_挂号状态 = 2 Then
    Update 病人挂号汇总
    Set 已约数 = Nvl(已约数, 0) - 1
    Where 日期 = Trunc(d_出诊日期) And Nvl(医生姓名, '-') = Nvl(v_原执行人, '-') And Nvl(医生id, 0) = Nvl(n_原执行人id, 0) And
          Nvl(科室id, 0) = Nvl(n_科室id, 0) And Nvl(项目id, 0) = Nvl(n_项目id, 0) And (号码 = v_号码 Or 号码 Is Null);
    Update 病人挂号汇总
    Set 已约数 = Nvl(已约数, 0) + 1
    Where 日期 = Trunc(d_出诊日期) And Nvl(医生姓名, '-') = Nvl(v_执行人, '-') And Nvl(医生id, 0) = Nvl(n_执行人id, 0) And
          Nvl(科室id, 0) = Nvl(n_科室id, 0) And Nvl(项目id, 0) = Nvl(n_项目id, 0) And (号码 = v_号码 Or 号码 Is Null);
    If Sql%RowCount = 0 Then
      Insert Into 病人挂号汇总
        (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 已约数, 其中已接收)
      Values
        (Trunc(d_出诊日期), n_科室id, n_项目id, v_执行人, Decode(n_执行人id, 0, Null, n_执行人id), v_号码, 0, 1, 0);
    End If;
  End If;
  If n_挂号状态 = 3 Then
    Update 病人挂号汇总
    Set 已约数 = Nvl(已约数, 0) - 1, 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - 1
    Where 日期 = Trunc(d_出诊日期) And Nvl(医生姓名, '-') = Nvl(v_原执行人, '-') And Nvl(医生id, 0) = Nvl(n_原执行人id, 0) And
          Nvl(科室id, 0) = Nvl(n_科室id, 0) And Nvl(项目id, 0) = Nvl(n_项目id, 0) And (号码 = v_号码 Or 号码 Is Null);
    Update 病人挂号汇总
    Set 已约数 = Nvl(已约数, 0) + 1, 已挂数 = Nvl(已挂数, 0) + 1, 其中已接收 = Nvl(其中已接收, 0) + 1
    Where 日期 = Trunc(d_出诊日期) And Nvl(医生姓名, '-') = Nvl(v_执行人, '-') And Nvl(医生id, 0) = Nvl(n_执行人id, 0) And
          Nvl(科室id, 0) = Nvl(n_科室id, 0) And Nvl(项目id, 0) = Nvl(n_项目id, 0) And (号码 = v_号码 Or 号码 Is Null);
    If Sql%RowCount = 0 Then
      Insert Into 病人挂号汇总
        (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 已约数, 其中已接收)
      Values
        (Trunc(d_出诊日期), n_科室id, n_项目id, v_执行人, Decode(n_执行人id, 0, Null, n_执行人id), v_号码, 1, 1, 1);
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_患者服务中心_替诊;
/


Create Or Replace Procedure Zl_患者服务中心_换诊
(
  消息id_In     病人服务信息记录.Id%Type,
  No_In         病人挂号记录.No%Type,
  换诊序号_In   病人挂号记录.号序%Type,
  换诊时间_In   病人挂号记录.预约时间%Type,
  换诊id_In     临床出诊记录.Id%Type,
  处理说明_In   病人服务信息记录.处理说明%Type,
  操作员姓名_In 病人服务信息记录.处理人%Type,
  操作员编号_In 病人挂号记录.操作员编号%Type
  
) As
  Cursor c_Registinfo Is
    Select a.发生时间, a.登记时间, c.接收时间, a.收费细目id As 项目id, c.执行部门id As 科室id, c.执行人 As 医生姓名, d.Id As 医生id, c.号别 As 号码, c.号序
    From 门诊费用记录 A, 挂号安排 B, 病人挂号记录 C, 人员表 D
    Where a.记录性质 = 4 And c.No = a.No And c.执行人 = d.姓名(+) And a.No = No_In And Nvl(a.计算单位, '号别') = c.号别 And Rownum < 2;
  r_Registrow   c_Registinfo%RowType;
  v_号别        病人挂号记录.号别%Type;
  n_执行部门id  病人挂号记录.执行部门id%Type;
  n_项目id      临床出诊记录.项目id%Type;
  v_执行人      病人挂号记录.执行人%Type;
  n_执行人id    人员表.Id%Type;
  n_病历费      Number(2);
  n_收费        Number(2);
  n_Exists      Number(3);
  v_Temp        Varchar2(500);
  v_收费项目ids Varchar2(500);
  v_Err_Msg     Varchar2(500);
  n_病历费id    收费项目目录.Id%Type;
  n_序号        门诊费用记录.序号%Type;
  n_预约        病人挂号记录.预约%Type;
  n_实收金额    门诊费用记录.实收金额%Type;
  n_应收金额    门诊费用记录.应收金额%Type;
  v_费别        门诊费用记录.费别%Type;
  n_病人id      病人挂号记录.病人id%Type;
  n_已用        临床出诊记录.已挂数%Type;
  n_限制        临床出诊记录.限号数%Type;
  n_原号序      病人挂号记录.号序%Type;
  n_原记录id    临床出诊记录.Id%Type;
  n_原挂号状态  临床出诊序号控制.挂号状态%Type;
  v_原操作员    临床出诊序号控制.操作员姓名%Type;
  v_原备注      临床出诊序号控制.备注%Type;
  n_序号控制    临床出诊记录.是否序号控制%Type;
  n_预约顺序号  临床出诊序号控制.预约顺序号%Type;
  n_实际序号    临床出诊序号控制.序号%Type;
  n_变动id      就诊变动记录.Id%Type;
  Err_Item Exception;
Begin
  Begin
    Select 1, 病人id Into n_Exists, n_病人id From 病人挂号记录 Where NO = No_In And 记录状态 = 1;
  Exception
    When Others Then
      v_Err_Msg := '单据号为' || No_In || '的预约记录不存在,无法换诊!';
      Raise Err_Item;
  End;
  Begin
    Select 费别 Into v_费别 From 门诊费用记录 Where NO = No_In And 记录性质 = 4;
  Exception
    When Others Then
      Begin
        Select 费别 Into v_费别 From 病人信息 Where 病人id = n_病人id;
      Exception
        When Others Then
          v_费别 := Null;
      End;
  End;

  Select b.号码, b.科室id, Nvl(c.姓名, a.医生姓名), a.项目id, c.Id, Nvl(a.是否序号控制, 0)
  Into v_号别, n_执行部门id, v_执行人, n_项目id, n_执行人id, n_序号控制
  From 临床出诊记录 A, 临床出诊号源 B, 人员表 C
  Where a.Id = 换诊id_In And a.号源id = b.Id And a.医生id = c.Id(+);

  Select Max(1) Into n_收费 From 门诊费用记录 Where NO = No_In And 记录性质 = 4 And 结帐金额 Is Not Null;
  Select Max(1)
  Into n_病历费
  From 门诊费用记录 A, 收费特定项目 B
  Where a.No = No_In And a.记录性质 = 4 And a.收费细目id = b.收费细目id And b.特定项目 = '病历费';

  --就诊变动记录
  Select 就诊变动记录_Id.Nextval Into n_变动id From Dual;
  b_Message.Zlhis_Regist_005(No_In, n_变动id, 2);
  Zl_就诊变动记录_Insert(No_In, 4, '患者服务中心换诊', 操作员姓名_In, 操作员编号_In, v_号别, n_执行部门id, n_项目id, n_执行人id, v_执行人, Null, 换诊序号_In,
                   换诊时间_In, n_变动id);

  --更新患者服务记录
  Update 病人服务信息记录 Set 处理人 = 操作员姓名_In, 处理时间 = Sysdate, 处理说明 = 处理说明_In Where ID = 消息id_In;

  --更新病人挂号汇总(减少)
  Select 预约, 出诊记录id Into n_预约, n_原记录id From 病人挂号记录 Where NO = No_In And 记录状态 = 1;

  --检查换诊记录是否数量足够
  If n_预约 = 0 Then
    Select 已挂数, 限号数 Into n_已用, n_限制 From 临床出诊记录 Where ID = 换诊id_In;
    If Not n_限制 Is Null Then
      If Nvl(n_已用, 0) >= n_限制 Then
        v_Err_Msg := '要换诊的记录已经超过最大限制数量' || n_限制 || ',无法换诊!';
        Raise Err_Item;
      End If;
    End If;
  Else
    If n_收费 = 1 Then
      Select 已挂数, 限号数 Into n_已用, n_限制 From 临床出诊记录 Where ID = 换诊id_In;
      If Not n_限制 Is Null Then
        If Nvl(n_已用, 0) >= n_限制 Then
          v_Err_Msg := '要换诊的记录已经超过最大限制数量' || n_限制 || ',无法换诊!';
          Raise Err_Item;
        End If;
      End If;
    Else
      Select 已约数, 限约数 Into n_已用, n_限制 From 临床出诊记录 Where ID = 换诊id_In;
      If Not n_限制 Is Null Then
        If Nvl(n_已用, 0) >= n_限制 Then
          v_Err_Msg := '要换诊的记录已经超过最大限制数量' || n_限制 || ',无法换诊!';
          Raise Err_Item;
        End If;
      End If;
    End If;
  End If;

  If n_预约 = 0 Then
    Open c_Registinfo;
    Fetch c_Registinfo
      Into r_Registrow;
  
    n_原号序 := r_Registrow.号序;
    Update 临床出诊记录 Set 已挂数 = Nvl(已挂数, 0) - 1 Where ID = n_原记录id;
    Update 病人挂号汇总
    Set 已挂数 = Nvl(已挂数, 0) - 1
    Where 日期 = Trunc(r_Registrow.发生时间) And 医生姓名 = r_Registrow.医生姓名 And 医生id = r_Registrow.医生id And
          科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And (号码 = r_Registrow.号码 Or 号码 Is Null);
  
    If Sql%RowCount = 0 Then
      Insert Into 病人挂号汇总
        (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数)
      Values
        (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
         Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, -1);
    End If;
    Close c_Registinfo;
  Else
    If n_收费 = 1 Then
      Open c_Registinfo;
      Fetch c_Registinfo
        Into r_Registrow;
    
      n_原号序 := r_Registrow.号序;
    
      Update 临床出诊记录
      Set 已约数 = Nvl(已约数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - 1, 已挂数 = Nvl(已挂数, 0) - 1
      Where ID = n_原记录id;
    
      Update 病人挂号汇总
      Set 已约数 = Nvl(已约数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - 1, 已挂数 = Nvl(已挂数, 0) - 1
      Where 日期 = Trunc(r_Registrow.发生时间) And 医生姓名 = r_Registrow.医生姓名 And 医生id = r_Registrow.医生id And
            科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And (号码 = r_Registrow.号码 Or 号码 Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into 病人挂号汇总
          (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已约数, 其中已接收, 已挂数)
        Values
          (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
           Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, -1, -1, -1);
      End If;
      Close c_Registinfo;
    Else
      Open c_Registinfo;
      Fetch c_Registinfo
        Into r_Registrow;
    
      n_原号序 := r_Registrow.号序;
    
      Update 临床出诊记录 Set 已约数 = Nvl(已约数, 0) - 1 Where ID = n_原记录id;
      Update 病人挂号汇总
      Set 已约数 = Nvl(已约数, 0) - 1
      Where 日期 = Trunc(r_Registrow.发生时间) And 医生姓名 = r_Registrow.医生姓名 And 医生id = r_Registrow.医生id And
            科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And (号码 = r_Registrow.号码 Or 号码 Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into 病人挂号汇总
          (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已约数)
        Values
          (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
           Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, -1);
      End If;
      Close c_Registinfo;
    End If;
  End If;

  If n_序号控制 = 0 And Nvl(换诊序号_In, 0) <> 0 Then
    Select Max(预约顺序号)
    Into n_预约顺序号
    From 临床出诊序号控制
    Where 记录id = 换诊id_In And 序号 = 换诊序号_In And 预约顺序号 Is Not Null;
    If n_预约顺序号 Is Null Then
      n_预约顺序号 := 1;
    Else
      n_预约顺序号 := n_预约顺序号 + 1;
    End If;
    n_实际序号 := To_Number(换诊序号_In || n_预约顺序号);
  Else
    n_实际序号 := 换诊序号_In;
  End If;
  --更新病人挂号记录
  Update 病人挂号记录
  Set 号别 = v_号别, 执行部门id = n_执行部门id, 执行人 = v_执行人, 号序 = n_实际序号, 发生时间 = 换诊时间_In, 预约时间 = 换诊时间_In, 出诊记录id = 换诊id_In
  Where NO = No_In And 记录状态 = 1;

  --更新门诊费用记录
  If n_病历费 = 1 Then
    Select 收费细目id Into n_病历费id From 收费特定项目 Where 特定项目 = '病历费';
    v_收费项目ids := n_项目id || ',' || n_病历费id;
  Else
    v_收费项目ids := n_项目id;
  End If;
  Update 门诊费用记录
  Set 病人科室id = n_执行部门id, 计算单位 = v_号别, 发药窗口 = n_实际序号, 执行部门id = n_执行部门id, 执行人 = v_执行人, 发生时间 = 换诊时间_In
  Where NO = No_In And 记录性质 = 4;
  n_序号 := 1;
  For c_Item In (Select 1 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, 1 As 数次, c.Id As 收入项目id,
                        c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, Null As 从属父号
                 From 收费项目目录 A, 收费价目 B, 收入项目 C
                 Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = n_项目id And Sysdate Between b.执行日期 And
                       Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))
                 Union All
                 Select 2 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, 1 As 数次, c.Id As 收入项目id,
                        c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, Null As 从属父号
                 From 收费项目目录 A, 收费价目 B, 收入项目 C
                 Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = n_病历费id And Sysdate Between b.执行日期 And
                       Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))
                 Union All
                 Select 3 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, d.从项数次 As 数次,
                        c.Id As 收入项目id, c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, 1 As 从属父号
                 From 收费项目目录 A, 收费价目 B, 收入项目 C, 收费从属项目 D
                 Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = d.从项id And
                       d.主项id In (Select Column_Value From Table(f_Str2list(v_收费项目ids))) And Sysdate Between b.执行日期 And
                       Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))
                 Order By 性质, 项目编码, 收入编码) Loop
    n_应收金额 := c_Item.单价 * c_Item.数次;
  
    If Nvl(c_Item.屏蔽费别, 0) <> 1 Then
      --打折:
      v_Temp     := Zl_Actualmoney(v_费别, c_Item.项目id, c_Item.收入项目id, n_应收金额);
      v_Temp     := Substr(v_Temp, Instr(v_Temp, ':') + 1);
      n_实收金额 := Zl_To_Number(v_Temp);
    Else
      n_实收金额 := n_应收金额;
    End If;
  
    If n_收费 = 1 Then
      Update 门诊费用记录
      Set 收费类别 = c_Item.类别, 收费细目id = c_Item.项目id, 收入项目id = c_Item.收入项目id, 收据费目 = c_Item.收据费目, 数次 = c_Item.数次,
          标准单价 = c_Item.单价, 应收金额 = n_应收金额, 实收金额 = n_实收金额, 结帐金额 = n_实收金额
      Where 序号 = n_序号 And 记录性质 = 4 And NO = No_In;
      If Sql%RowCount = 0 Then
        Insert Into 门诊费用记录
          (ID, 记录性质, 记录状态, 序号, 价格父号, 从属父号, NO, 实际票号, 门诊标志, 加班标志, 附加标志, 发药窗口, 病人id, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 病人科室id,
           收费类别, 计算单位, 收费细目id, 收入项目id, 收据费目, 付数, 数次, 标准单价, 应收金额, 实收金额, 结帐金额, 结帐id, 记帐费用, 开单部门id, 开单人, 划价人, 执行部门id, 执行人,
           操作员编号, 操作员姓名, 发生时间, 登记时间, 保险大类id, 保险项目否, 保险编码, 统筹金额, 摘要, 结论, 缴款组id)
          Select 病人费用记录_Id.Nextval, 4, 记录状态, n_序号, Null, Null, NO, 实际票号, 门诊标志, Null, Null, 发药窗口, 病人id, 标识号, 付款方式, 姓名, 性别,
                 年龄, 费别, 病人科室id, c_Item.类别, 计算单位, c_Item.项目id, c_Item.收入项目id, c_Item.收据费目, 1, c_Item.数次, c_Item.单价,
                 n_应收金额, n_实收金额, n_实收金额, 结帐id, 0, 开单部门id, 开单人, 划价人, 执行部门id, 执行人, 操作员编号, 操作员姓名, 发生时间, 登记时间, 保险大类id, 保险项目否,
                 保险编码, 统筹金额, 摘要, 结论, 缴款组id
          From 门诊费用记录
          Where 记录性质 = 4 And NO = No_In And 序号 = 1;
      End If;
    Else
      Update 门诊费用记录
      Set 收费细目id = c_Item.项目id, 收入项目id = c_Item.收入项目id, 收据费目 = c_Item.收据费目, 标准单价 = c_Item.单价, 应收金额 = c_Item.单价,
          实收金额 = c_Item.单价
      Where 序号 = n_序号 And 记录性质 = 4 And NO = No_In;
      If Sql%RowCount = 0 Then
        Insert Into 门诊费用记录
          (ID, 记录性质, 记录状态, 序号, 价格父号, 从属父号, NO, 实际票号, 门诊标志, 加班标志, 附加标志, 发药窗口, 病人id, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 病人科室id,
           收费类别, 计算单位, 收费细目id, 收入项目id, 收据费目, 付数, 数次, 标准单价, 应收金额, 实收金额, 结帐金额, 结帐id, 记帐费用, 开单部门id, 开单人, 划价人, 执行部门id, 执行人,
           操作员编号, 操作员姓名, 发生时间, 登记时间, 保险大类id, 保险项目否, 保险编码, 统筹金额, 摘要, 结论, 缴款组id)
          Select 病人费用记录_Id.Nextval, 4, 记录状态, n_序号, Null, Null, NO, 实际票号, 门诊标志, Null, Null, 发药窗口, 病人id, 标识号, 付款方式, 姓名, 性别,
                 年龄, 费别, 病人科室id, c_Item.类别, 计算单位, c_Item.项目id, c_Item.收入项目id, c_Item.收据费目, 1, c_Item.数次, c_Item.单价,
                 n_应收金额, n_实收金额, Null, 结帐id, 0, 开单部门id, 开单人, 划价人, 执行部门id, 执行人, 操作员编号, 操作员姓名, 发生时间, 登记时间, 保险大类id, 保险项目否,
                 保险编码, 统筹金额, 摘要, 结论, 缴款组id
          From 门诊费用记录
          Where 记录性质 = 4 And NO = No_In And 序号 = 1;
      End If;
    End If;
    n_序号 := n_序号 + 1;
  End Loop;

  --更新病人挂号汇总(增加)
  If n_预约 = 0 Then
    Open c_Registinfo;
    Fetch c_Registinfo
      Into r_Registrow;
  
    Update 临床出诊记录 Set 已挂数 = Nvl(已挂数, 0) + 1 Where ID = 换诊id_In;
  
    Update 病人挂号汇总
    Set 已挂数 = Nvl(已挂数, 0) + 1
    Where 日期 = Trunc(r_Registrow.发生时间) And 医生姓名 = r_Registrow.医生姓名 And 医生id = r_Registrow.医生id And
          科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And (号码 = r_Registrow.号码 Or 号码 Is Null);
  
    If Sql%RowCount = 0 Then
      Insert Into 病人挂号汇总
        (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数)
      Values
        (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
         Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, 1);
    End If;
    Close c_Registinfo;
  Else
    If n_收费 = 1 Then
      Open c_Registinfo;
      Fetch c_Registinfo
        Into r_Registrow;
    
      Update 临床出诊记录
      Set 已约数 = Nvl(已约数, 0) + 1, 其中已接收 = Nvl(其中已接收, 0) + 1, 已挂数 = Nvl(已挂数, 0) + 1
      Where ID = 换诊id_In;
    
      Update 病人挂号汇总
      Set 已约数 = Nvl(已约数, 0) + 1, 其中已接收 = Nvl(其中已接收, 0) + 1, 已挂数 = Nvl(已挂数, 0) + 1
      Where 日期 = Trunc(r_Registrow.发生时间) And 医生姓名 = r_Registrow.医生姓名 And 医生id = r_Registrow.医生id And
            科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And (号码 = r_Registrow.号码 Or 号码 Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into 病人挂号汇总
          (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已约数, 其中已接收, 已挂数)
        Values
          (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
           Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, 1, 1, 1);
      End If;
      Close c_Registinfo;
    Else
      Open c_Registinfo;
      Fetch c_Registinfo
        Into r_Registrow;
    
      Update 临床出诊记录 Set 已约数 = Nvl(已约数, 0) + 1 Where ID = 换诊id_In;
    
      Update 病人挂号汇总
      Set 已约数 = Nvl(已约数, 0) + 1
      Where 日期 = Trunc(r_Registrow.发生时间) And 医生姓名 = r_Registrow.医生姓名 And 医生id = r_Registrow.医生id And
            科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And (号码 = r_Registrow.号码 Or 号码 Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into 病人挂号汇总
          (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已约数)
        Values
          (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
           Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, 1);
      End If;
      Close c_Registinfo;
    End If;
  End If;
  --更新序号
  Begin
    Select 挂号状态, 操作员姓名, 备注
    Into n_原挂号状态, v_原操作员, v_原备注
    From 临床出诊序号控制
    Where 记录id = n_原记录id And (序号 = n_原号序 Or 备注 = n_原号序);
  
    If n_序号控制 = 0 And Nvl(换诊序号_In, 0) <> 0 Then
      Insert Into 临床出诊序号控制
        (记录id, 序号, 预约顺序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 操作员姓名, 备注)
        Select 记录id, 序号, n_预约顺序号, 开始时间, 终止时间, 1, 是否预约, n_原挂号状态, v_原操作员, n_实际序号
        From 临床出诊序号控制
        Where 记录id = 换诊id_In And 序号 = 换诊序号_In And 预约顺序号 Is Null;
    Else
      Update 临床出诊序号控制
      Set 挂号状态 = n_原挂号状态, 操作员姓名 = v_原操作员, 备注 = v_原备注
      Where 记录id = 换诊id_In And 序号 = 换诊序号_In;
    End If;
  
    Update 临床出诊序号控制
    Set 挂号状态 = Null, 操作员姓名 = Null, 备注 = Null
    Where 记录id = n_原记录id And (序号 = n_原号序 Or 备注 = n_原号序);
  Exception
    When Others Then
      Null;
  End;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_患者服务中心_换诊;
/


Create Or Replace Procedure Zl_患者服务中心_更新
(
  消息id_In     病人服务信息记录.Id%Type,
  处理说明_In   病人服务信息记录.处理说明%Type,
  操作员姓名_In 病人服务信息记录.处理人%Type,
  操作员编号_In 病人挂号记录.操作员编号%Type,
  挂号id_In     病人挂号记录.Id%Type := Null,
  操作方式_In   Number := 0
) As
  --操作方式_IN:0-正常更新,1-取消预约登记
  v_Err_Msg Varchar2(200);
  Err_Item Exception;
Begin
  --更新患者服务记录
  If 操作方式_In = 0 Then
    If 挂号id_In Is Null Then
      Update 病人服务信息记录
      Set 处理人 = 操作员姓名_In, 处理时间 = Sysdate, 处理说明 = 处理说明_In
      Where ID = 消息id_In;
    Else
      Update 病人服务信息记录
      Set 处理人 = 操作员姓名_In, 处理时间 = Sysdate, 处理说明 = 处理说明_In, 挂号id = 挂号id_In
      Where ID = 消息id_In;
    End If;
  Else
    Delete From 病人服务信息记录 Where ID = 消息id_In;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_患者服务中心_更新;
/


Create Or Replace Procedure Zl_病人挂号汇总_Update
(
  医生姓名_In   挂号安排.医生姓名%Type,
  医生id_In     挂号安排.医生id%Type,
  收费细目id_In 门诊费用记录.收费细目id%Type,
  执行部门id_In 门诊费用记录.执行部门id%Type,
  发生时间_In   门诊费用记录.发生时间%Type,
  预约标志_In   Number := 0, --是否为预约接收:0-非预约挂号; 1-预约挂号,2-预约接收;3-收费预约
  号码_In       挂号安排.号码%Type := Null,
  三方调用_In   Number := 0, --是否接口调用
  出诊记录id_In 临床出诊记录.Id%Type := Null
) As
  --发生时间_In:预约时,为预约时间;否则为登记时间
  v_Date    Date;
  n_预约数  病人挂号汇总.已约数%Type;
  n_时段    Number := 0;
  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  n_接收模式 Number := 0;
  n_号源id   临床出诊记录.号源id%Type;
Begin
  If 出诊记录id_In Is Null Then
    Begin
      Select 1
      Into n_时段
      From Dual
      Where Exists (Select 1
             From 挂号安排时段 A, 挂号安排 B
             Where a.安排id = b.Id And b.号码 = 号码_In And Rownum <= 1
             Union All
             Select 1
             From 挂号计划时段 C, 挂号安排计划 D 　
             Where c.计划id = d.Id And d.号码 = 号码_In And d.生效时间 > Sysdate And Rownum <= 1);
    Exception
      When Others Then
        n_时段 := 0;
    End;
    n_接收模式 := Nvl(zl_GetSysParameter('预约接收模式', 1111), 0);
    --分时段的号别，只能当天接收
    If n_时段 = 1 And Nvl(预约标志_In, 0) = 2 And 三方调用_In = 0 And n_接收模式 = 0 Then
      If Trunc(发生时间_In) <> Trunc(Sysdate) Then
        v_Err_Msg := '分时段的预约挂号单只能当天接收！';
        Raise Err_Item;
      End If;
    End If;
    If Nvl(预约标志_In, 0) <> 2 Or 三方调用_In = 1 Then
      v_Date := Trunc(发生时间_In);
    Else
      If n_接收模式 = 0 Then
        v_Date := Trunc(Sysdate);
      Else
        v_Date := Trunc(发生时间_In);
      End If;
    End If;
  
    n_预约数 := 0;
    If Nvl(预约标志_In, 0) <> 1 Then
      --非预约挂号;或预约接收
      If Nvl(预约标志_In, 0) = 2 And v_Date <> Trunc(发生时间_In) Then
        --1.减去预约日期的预约数;
        --2-加上当前预约日期的挂号数;
        Update 病人挂号汇总
        Set 已约数 = Nvl(已约数, 0) - 1
        Where 日期 = Trunc(发生时间_In) And Nvl(科室id, 0) = 执行部门id_In And Nvl(项目id, 0) = 收费细目id_In And
              (号码 = 号码_In Or 号码 Is Null)
        Returning 已约数 Into n_预约数;
      
        If n_预约数 < 0 Then
          Update 病人挂号汇总
          Set 已约数 = 0
          Where 日期 = Trunc(发生时间_In) And Nvl(科室id, 0) = 执行部门id_In And Nvl(项目id, 0) = 收费细目id_In And
                (号码 = 号码_In Or 号码 Is Null)
          Returning 已约数 Into n_预约数;
        End If;
        n_预约数 := 1;
      Elsif Nvl(预约标志_In, 0) = 3 Then
        n_预约数 := 1;
      End If;
      Update 病人挂号汇总
      Set 已挂数 = Nvl(已挂数, 0) + 1, 其中已接收 = Nvl(其中已接收, 0) + Decode(预约标志_In, 0, 0, 1), 已约数 = Nvl(已约数, 0) + Nvl(n_预约数, 0)
      Where 日期 = Decode(预约标志_In, 2, Trunc(v_Date), Trunc(发生时间_In)) And Nvl(科室id, 0) = 执行部门id_In And
            Nvl(项目id, 0) = 收费细目id_In And (号码 = 号码_In Or 号码 Is Null);
      If Sql%RowCount = 0 Then
        Insert Into 病人挂号汇总
          (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 已约数, 其中已接收)
        Values
          (Decode(预约标志_In, 2, Trunc(v_Date), Trunc(发生时间_In)), 执行部门id_In, 收费细目id_In, 医生姓名_In,
           Decode(医生id_In, 0, Null, 医生id_In), 号码_In, 1, Nvl(n_预约数, 0), Decode(预约标志_In, 0, 0, 1));
      End If;
    Else
      --预约挂号
      Update 病人挂号汇总
      Set 已约数 = Nvl(已约数, 0) + 1
      Where 日期 = Trunc(v_Date) And Nvl(科室id, 0) = 执行部门id_In And Nvl(项目id, 0) = 收费细目id_In And (号码 = 号码_In Or 号码 Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into 病人挂号汇总
          (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已约数)
        Values
          (Trunc(v_Date), 执行部门id_In, 收费细目id_In, 医生姓名_In, Decode(医生id_In, 0, Null, 医生id_In), 号码_In, 1);
      End If;
    End If;
  Else
    --出诊表排班模式
    Begin
      Select Nvl(是否分时段, 0) Into n_时段 From 临床出诊记录 Where ID = 出诊记录id_In;
    Exception
      When Others Then
        n_时段 := 0;
    End;
    Select 号源id Into n_号源id From 临床出诊记录 Where ID = 出诊记录id_In;
    n_接收模式 := Nvl(zl_GetSysParameter('预约接收模式', 1111), 0);
    --分时段的号别，只能当天接收
    If n_时段 = 1 And Nvl(预约标志_In, 0) = 2 And 三方调用_In = 0 And n_接收模式 = 0 Then
      If Trunc(发生时间_In) <> Trunc(Sysdate) Then
        v_Err_Msg := '分时段的预约挂号单只能当天接收！';
        Raise Err_Item;
      End If;
    End If;
    If Nvl(预约标志_In, 0) <> 2 Or 三方调用_In = 1 Then
      v_Date := Trunc(发生时间_In);
    Else
      If n_接收模式 = 0 Then
        v_Date := Trunc(Sysdate);
      Else
        v_Date := Trunc(发生时间_In);
      End If;
    End If;
  
    n_预约数 := 0;
    If Nvl(预约标志_In, 0) <> 1 Then
      --非预约挂号;或预约接收
      If Nvl(预约标志_In, 0) = 2 And v_Date <> Trunc(发生时间_In) Then
        --1.减去预约日期的预约数;
        --2-加上当前预约日期的挂号数;
        Update 临床出诊记录 Set 已约数 = Nvl(已约数, 0) - 1 Where ID = 出诊记录id_In Returning 已约数 Into n_预约数;
        If n_预约数 < 0 Then
          Update 临床出诊记录 Set 已约数 = 0 Where ID = 出诊记录id_In Returning 已约数 Into n_预约数;
        End If;
        Update 病人挂号汇总
        Set 已约数 = Nvl(已约数, 0) - 1
        Where 日期 = Trunc(发生时间_In) And Nvl(医生id, 0) = Nvl(医生id_In, 0) And Nvl(医生姓名, '-') = Nvl(医生姓名_In, '-') And
              Nvl(科室id, 0) = 执行部门id_In And Nvl(项目id, 0) = 收费细目id_In And (号码 = 号码_In Or 号码 Is Null)
        Returning 已约数 Into n_预约数;
        If n_预约数 < 0 Then
          Update 病人挂号汇总
          Set 已约数 = 0
          Where 日期 = Trunc(发生时间_In) And Nvl(医生id, 0) = Nvl(医生id_In, 0) And Nvl(医生姓名, '-') = Nvl(医生姓名_In, '-') And
                Nvl(科室id, 0) = 执行部门id_In And Nvl(项目id, 0) = 收费细目id_In And (号码 = 号码_In Or 号码 Is Null)
          Returning 已约数 Into n_预约数;
        End If;
        n_预约数 := 1;
      Elsif Nvl(预约标志_In, 0) = 3 Then
        n_预约数 := 1;
      End If;
    
      Update 临床出诊记录
      Set 已挂数 = Nvl(已挂数, 0) + 1, 其中已接收 = Nvl(其中已接收, 0) + Decode(预约标志_In, 0, 0, 1), 已约数 = Nvl(已约数, 0) + Nvl(n_预约数, 0)
      Where ID = 出诊记录id_In;
    
      Update 病人挂号汇总
      Set 已挂数 = Nvl(已挂数, 0) + 1, 其中已接收 = Nvl(其中已接收, 0) + Decode(预约标志_In, 0, 0, 1), 已约数 = Nvl(已约数, 0) + Nvl(n_预约数, 0)
      Where 日期 = Decode(预约标志_In, 2, Trunc(v_Date), Trunc(发生时间_In)) And Nvl(科室id, 0) = 执行部门id_In And
            Nvl(项目id, 0) = 收费细目id_In And Nvl(医生id, 0) = Nvl(医生id_In, 0) And Nvl(医生姓名, '-') = Nvl(医生姓名_In, '-') And
            (号码 = 号码_In Or 号码 Is Null);
      If Sql%RowCount = 0 Then
        Insert Into 病人挂号汇总
          (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 已约数, 其中已接收)
        Values
          (Decode(预约标志_In, 2, Trunc(v_Date), Trunc(发生时间_In)), 执行部门id_In, 收费细目id_In, 医生姓名_In,
           Decode(医生id_In, 0, Null, 医生id_In), 号码_In, 1, Nvl(n_预约数, 0), Decode(预约标志_In, 0, 0, 1));
      End If;
    Else
      --预约挂号
      Update 临床出诊记录 Set 已约数 = Nvl(已约数, 0) + 1 Where ID = 出诊记录id_In;
      Update 病人挂号汇总
      Set 已约数 = Nvl(已约数, 0) + 1
      Where 日期 = Trunc(v_Date) And Nvl(医生id, 0) = Nvl(医生id_In, 0) And Nvl(医生姓名, '-') = Nvl(医生姓名_In, '-') And
            Nvl(科室id, 0) = 执行部门id_In And Nvl(项目id, 0) = 收费细目id_In And (号码 = 号码_In Or 号码 Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into 病人挂号汇总
          (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已约数)
        Values
          (Trunc(v_Date), 执行部门id_In, 收费细目id_In, 医生姓名_In, Decode(医生id_In, 0, Null, 医生id_In), 号码_In, 1);
      End If;
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人挂号汇总_Update;
/


Create Or Replace Procedure Zl_病人挂号记录_Insert
(
  病人id_In        门诊费用记录.病人id%Type,
  门诊号_In        门诊费用记录.标识号%Type,
  姓名_In          门诊费用记录.姓名%Type,
  性别_In          门诊费用记录.性别%Type,
  年龄_In          门诊费用记录.年龄%Type,
  付款方式_In      门诊费用记录.付款方式%Type, --用于存放病人的医疗付款方式编号
  费别_In          门诊费用记录.费别%Type,
  单据号_In        门诊费用记录.No%Type,
  票据号_In        门诊费用记录.实际票号%Type,
  序号_In          门诊费用记录.序号%Type,
  价格父号_In      门诊费用记录.价格父号%Type,
  从属父号_In      门诊费用记录.从属父号%Type,
  收费类别_In      门诊费用记录.收费类别%Type,
  收费细目id_In    门诊费用记录.收费细目id%Type,
  数次_In          门诊费用记录.数次%Type,
  标准单价_In      门诊费用记录.标准单价%Type,
  收入项目id_In    门诊费用记录.收入项目id%Type,
  收据费目_In      门诊费用记录.收据费目%Type,
  结算方式_In      病人预交记录.结算方式%Type, --现金的结算名称
  应收金额_In      门诊费用记录.应收金额%Type,
  实收金额_In      门诊费用记录.实收金额%Type,
  病人科室id_In    门诊费用记录.病人科室id%Type,
  开单部门id_In    门诊费用记录.开单部门id%Type,
  执行部门id_In    门诊费用记录.执行部门id%Type,
  操作员编号_In    门诊费用记录.操作员编号%Type,
  操作员姓名_In    门诊费用记录.操作员姓名%Type,
  发生时间_In      门诊费用记录.发生时间%Type,
  登记时间_In      门诊费用记录.登记时间%Type,
  医生姓名_In      挂号安排.医生姓名%Type,
  医生id_In        挂号安排.医生id%Type,
  病历费_In        Number, --该条记录是否病历工本费
  急诊_In          Number,
  号别_In          挂号安排.号码%Type,
  诊室_In          门诊费用记录.发药窗口%Type,
  结帐id_In        门诊费用记录.结帐id%Type,
  领用id_In        票据使用明细.领用id%Type,
  预交支付_In      病人预交记录.冲预交%Type, --刷卡挂号时使用的预交金额,序号为1传入.
  现金支付_In      病人预交记录.冲预交%Type, --挂号时现金支付部份金额,序号为1传入.
  个帐支付_In      病人预交记录.冲预交%Type, --挂号时个人帐户支付金额,,序号为1传入.
  保险大类id_In    门诊费用记录.保险大类id%Type,
  保险项目否_In    门诊费用记录.保险项目否%Type,
  统筹金额_In      门诊费用记录.统筹金额%Type,
  摘要_In          门诊费用记录.摘要%Type, --预约挂号摘要信息
  预约挂号_In      Number := 0, --预约挂号时用(记录状态=0,发生时间为预约时间),此时不需要传入结算相关参数
  收费票据_In      Number := 0, --挂号是否使用收费票据
  保险编码_In      门诊费用记录.保险编码%Type,
  复诊_In          病人挂号记录.复诊%Type := 0,
  号序_In          挂号序号状态.序号%Type := Null, --预约时填入费用记录的发药窗口字段,挂号时填入挂号记录
  社区_In          病人挂号记录.社区%Type := Null,
  预约接收_In      Number := 0,
  预约方式_In      预约方式.名称%Type := Null,
  生成队列_In      Number := 0,
  卡类别id_In      病人预交记录.卡类别id%Type := Null,
  结算卡序号_In    病人预交记录.结算卡序号%Type := Null,
  卡号_In          病人预交记录.卡号%Type := Null,
  交易流水号_In    病人预交记录.交易流水号%Type := Null,
  交易说明_In      病人预交记录.交易说明%Type := Null,
  合作单位_In      病人预交记录.合作单位%Type := Null,
  操作类型_In      Number := 0,
  险类_In          病人挂号记录.险类%Type := Null,
  结算模式_In      Number := 0,
  记帐费用_In      Number := 0,
  退号重用_In      Number := 1,
  冲预交病人ids_In Varchar2 := Null,
  修正病人费别_In  Number := 0
) As
  ---------------------------------------------------------------------------
  --
  --参数:
  --     操作类型_in:0-正常挂号或者预约 1-操作员拥有加号权限加号
  --     修正病人费别_In:0-不修改病人费别 1-修改病人费别
  ----------------------------------------------------------------------------
  --该游标用于收费冲预交的可用预交列表
  --以ID排序，优先冲上次未冲完的。
  Cursor c_Deposit
  (
    v_病人id        病人信息.病人id%Type,
    v_冲预交病人ids Varchar2
  ) Is
    Select *
    From (Select a.Id, a.病人id, a.记录状态, a.预交类别, a.No, Nvl(a.金额, 0) As 金额
           From 病人预交记录 A,
                (Select NO, Sum(Nvl(a.金额, 0)) As 金额
                  From 病人预交记录 A
                  Where a.结帐id Is Null And Nvl(a.金额, 0) <> 0 And
                        a.病人id In (Select Column_Value From Table(f_Num2list(v_冲预交病人ids))) And Nvl(a.预交类别, 2) = 1
                  Group By NO
                  Having Sum(Nvl(a.金额, 0)) <> 0) B
           Where a.结帐id Is Null And Nvl(a.金额, 0) <> 0 And a.结算方式 Not In (Select 名称 From 结算方式 Where 性质 = 5) And
                 a.No = b.No And a.病人id In (Select Column_Value From Table(f_Num2list(v_冲预交病人ids))) And
                 Nvl(a.预交类别, 2) = 1
           Union All
           Select 0 As ID, Max(病人id) As 病人id, 记录状态, 预交类别, NO, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额
           From 病人预交记录
           Where 记录性质 In (1, 11) And 结帐id Is Not Null And Nvl(金额, 0) <> Nvl(冲预交, 0) And
                 病人id In (Select Column_Value From Table(f_Num2list(v_冲预交病人ids))) And Nvl(预交类别, 2) = 1 Having
            Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
           Group By 记录状态, NO, 预交类别)
    Order By Decode(病人id, Nvl(v_病人id, 0), 0, 1), ID, NO;
  --功能：新增一行病人挂号费用，并将结算汇总到病人预交记录
  --       同时汇总相关的汇总表(病人挂号汇总、费用汇总)
  --       第一行费用处理票据使用情况(领用ID_IN>0)

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_排队号码 排队叫号队列.排队号码%Type;
  v_现金     结算方式.名称%Type;
  v_个人帐户 结算方式.名称%Type;
  v_队列名称 排队叫号队列.队列名称%Type;

  n_分时段       Number;
  n_时段限号     Number;
  n_时段限约     Number;
  d_时段时间     Date;
  d_最大序号时间 Date;
  n_追加号       Number := 0; --处理时段过期 追加挂号的情况 
  n_已约数       病人挂号汇总.已约数%Type;
  n_已接收       病人挂号汇总.其中已接收%Type;
  n_预约有效时间 Number;
  n_失效数       Number;
  n_失约挂号     Number := 0;
  n_已用数量     Number;
  n_锁定         Number := 0;

  n_打印id        票据打印内容.Id%Type;
  n_费用id        门诊费用记录.Id%Type;
  n_预交金额      病人预交记录.金额%Type;
  n_当前金额      病人预交记录.金额%Type;
  n_返回值        病人预交记录.金额%Type;
  n_预交id        病人预交记录.Id%Type;
  n_消费卡id      消费卡目录.Id%Type;
  n_挂号id        病人挂号记录.Id%Type;
  v_冲预交病人ids Varchar2(4000);

  n_组id           财务缴款分组.Id%Type;
  n_门诊号         病人信息.门诊号%Type;
  n_序号           挂号序号状态.序号%Type;
  n_已用序号       挂号序号状态.序号%Type;
  n_序号控制       挂号安排.序号控制%Type;
  n_分诊台签到排队 Number;
  n_Count          Number;
  n_限号数         Number(18);
  n_自制卡         Number;
  d_排队时间       Date;
  d_序号时间       Date;
  v_费别           费别.名称%Type;
  n_时段序号       Number := -1;
  n_预约生成队列   Number;
  n_安排id         挂号安排.Id%Type;
  n_计划id         挂号安排计划.Id%Type := 0;
  v_星期           挂号安排限制.限制项目%Type;
  n_限约数         Number(18);
  n_已挂数         Number(4) := 0;

  n_挂出的最大序号 Number(4) := 0;
  n_结算模式       病人信息.结算模式%Type;
  v_排队序号       排队叫号队列.排队序号%Type;
  v_机器名         挂号序号状态.机器名%Type;
  v_序号操作员     挂号序号状态.操作员姓名%Type;
  v_序号机器名     挂号序号状态.机器名%Type;
  v_付款方式       病人挂号记录.医疗付款方式%Type;
  v_Temp           Varchar2(3000);
  d_启用时间       Date;
Begin
  --获取当前机器名称
  Select Terminal Into v_机器名 From V$session Where Audsid = Userenv('sessionid');
  n_组id          := Zl_Get组id(操作员姓名_In);
  v_冲预交病人ids := Nvl(冲预交病人ids_In, 病人id_In);

  If 费别_In Is Null Then
    Begin
      Select 名称 Into v_费别 From 费别 Where 缺省标志 = 1 And Rownum < 2;
      Update 病人信息 Set 费别 = v_费别 Where 病人id = 病人id_In;
    Exception
      When Others Then
        v_Err_Msg := '无法确定病人费别，请检查缺省费别是否正确设置！';
        Raise Err_Item;
    End;
  Else
    v_费别 := 费别_In;
    If Nvl(修正病人费别_In, 0) = 1 Then
      Begin
        Update 病人信息 Set 费别 = v_费别 Where 病人id = 病人id_In;
      Exception
        When Others Then
          v_Err_Msg := '没有找到对应的病人！';
          Raise Err_Item;
      End;
    End If;
  End If;

  If 门诊号_In Is Not Null Then
    Begin
      Select Nvl(门诊号, 0) Into n_门诊号 From 病人信息 Where 病人id = 病人id_In;
    Exception
      When Others Then
        n_门诊号 := 0;
    End;
    If n_门诊号 = 0 Then
      Update 病人信息 Set 门诊号 = 门诊号_In Where 病人id = 病人id_In;
    End If;
  End If;

  Begin
    Delete From 挂号序号状态
    Where 号码 = 号别_In And 日期 = 发生时间_In And 序号 = 号序_In And 状态 = 3 And 操作员姓名 = 操作员姓名_In;
  Exception
    When Others Then
      Null;
  End;

  v_Temp := zl_GetSysParameter(256);
  If v_Temp Is Null Or Substr(v_Temp, 1, 1) = '0' Then
    Null;
  Else
    Begin
      d_启用时间 := To_Date(Substr(v_Temp, 3), 'YYYY-MM-DD hh24:mi:ss');
      If 发生时间_In > d_启用时间 Then
        v_Err_Msg := '当前挂号的发生时间' || To_Char(发生时间_In, 'yyyy-mm-dd hh24:mi:ss') || '已经启用了出诊表排班模式,不能再使用计划排班模式挂号!';
        Raise Err_Item;
      End If;
    Exception
      When Others Then
        Null;
    End;
  End If;

  If Nvl(预约挂号_In, 0) = 0 Then
    --挂号或者预约接收
    --因为现在有按日编单据号规则,日挂号量不能超过10000次,所以要检查唯一约束。
    Select Count(*)
    Into n_Count
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 In (1, 3) And 序号 = 序号_In And NO = 单据号_In;
    If n_Count <> 0 Then
      v_Err_Msg := '挂号单据号重复,不能保存！' || Chr(13) || '如果使用了按日顺序编号,当日挂号量不能超过10000人次。';
      Raise Err_Item;
    End If;
  
    --获取结算方式名称
    Begin
      Select 名称 Into v_现金 From 结算方式 Where 性质 = 1;
    Exception
      When Others Then
        v_现金 := '现金';
    End;
    Begin
      Select 名称 Into v_个人帐户 From 结算方式 Where 性质 = 3;
    Exception
      When Others Then
        v_个人帐户 := '个人帐户';
    End;
  End If;

  n_序号 := 号序_In;
  Select Decode(To_Char(发生时间_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null)
  Into v_星期
  From Dual;

  --挂号获取安排
  Begin
    Select a.Id, a.序号控制, Nvl(b.限号数, 0), Nvl(b.限约数, 0)
    Into n_安排id, n_序号控制, n_限号数, n_限约数
    From 挂号安排 A, 挂号安排限制 B
    Where a.Id = b.安排id(+) And b.限制项目(+) = v_星期 And a.号码 = 号别_In;
  
  Exception
    When Others Then
      n_安排id := -1;
  End;

  --如果是病历费或者号别为空时不检查
  If Nvl(病历费_In, 0) = 0 Or 号别_In Is Not Null Then
    If n_安排id = -1 Then
      v_Err_Msg := '不存相应的挂号安排数据,请检查';
      Raise Err_Item;
    End If;
  End If;

  If Nvl(预约挂号_In, 0) = 1 Then
    --首先获取计划
    Begin
      Select ID
      Into n_计划id
      From 挂号安排计划
      Where 安排id = n_安排id And 审核时间 Is Not Null And
            Nvl(生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) =
            (Select Max(a.生效时间) As 生效
             From 挂号安排计划 A
             Where a.审核时间 Is Not Null And 发生时间_In Between Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                   Nvl(a.失效时间, To_Date('3000-01-01', 'yyyy-mm-dd')) And a.安排id = n_安排id) And
            发生时间_In Between Nvl(生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
            Nvl(失效时间, To_Date('3000-01-01', 'yyyy-mm-dd'));
    
    Exception
      When Others Then
        n_计划id := 0;
    End;
    If Nvl(n_计划id, 0) <> 0 Then
      Begin
        --获取计划的限制
        Select a.Id, a.序号控制, Nvl(b.限号数, 0) As 限号数, Nvl(b.限约数, 0) As 限约数
        Into n_计划id, n_序号控制, n_限号数, n_限约数
        From 挂号安排计划 A, 挂号计划限制 B
        Where a.号码 = 号别_In And a.Id = n_计划id And a.审核时间 Is Not Null And a.Id = b.计划id(+) And b.限制项目(+) = v_星期;
      Exception
        When Others Then
          v_Err_Msg := '不存相应的挂号安排或计划数据,请检查';
          Raise Err_Item;
      End;
    End If;
  End If;

  --获取是否分时段
  If Nvl(n_计划id, 0) = 0 Then
    Select Count(Rownum) Into n_分时段 From 挂号安排时段 Where 星期 = v_星期 And 安排id = n_安排id And Rownum <= 1;
  Else
    Select Count(Rownum) Into n_分时段 From 挂号计划时段 Where 星期 = v_星期 And 计划id = n_计划id And Rownum <= 1;
  End If;

  --分时段挂号时判断是否是过期挂号 也就是追加号的情况 目前只针对专家号分时段进行处理
  If Nvl(预约挂号_In, 0) = 0 And n_分时段 > 0 And Nvl(n_序号控制, 0) = 1 And 号序_In Is Null And Nvl(操作类型_In, 0) = 0 Then
    --发生时间_in>Sysdate 发生时间>最大的时段时间--号序_in is null
    Begin
      Select Max(To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'))
      Into d_最大序号时间
      From 挂号安排时段
      Where 安排id = n_安排id And 星期 = v_星期 And Nvl(限制数量, 0) <> 0;
      n_追加号 := Case Sign(发生时间_In - d_最大序号时间)
                 When -1 Then
                  0
                 Else
                  1
               End;
    Exception
      When Others Then
        n_追加号 := 0;
    End;
  End If;
  d_时段时间 := 发生时间_In;

  If 序号_In = 1 And Nvl(预约挂号_In, 0) = 0 And n_分时段 > 0 Then
    --挂号时检查 是否分了时段,分了时段,把时段的限制条件给取出来
    Begin
      Select Nvl(序号, 0),
             To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
             限制数量, Decode(Nvl(是否预约, 0), 0, 0, 限制数量)
      Into n_时段序号, d_时段时间, n_时段限号, n_时段限约
      From 挂号安排时段
      Where 安排id = n_安排id And 星期 = v_星期 And
            (序号, 安排id, 星期) In (Select Nvl(Max(序号), -1), 安排id, 星期
                               From 挂号安排时段
                               Where 安排id = n_安排id And 星期 = v_星期 And
                                     Decode(操作类型_In + n_追加号, 0, To_Char(发生时间_In, 'hh24:mi'),
                                            To_Char(Nvl(开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')) =
                                     To_Char(Nvl(开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
                               Group By 安排id, 星期);
    Exception
      When Others Then
        n_时段序号 := -1;
        n_分时段   := 0;
        d_时段时间 := 发生时间_In;
        n_时段限号 := 0;
        n_时段限约 := 0;
    End;
  End If;

  If 序号_In = 1 And Nvl(预约挂号_In, 0) = 1 And n_分时段 > 0 Then
    --预约号,取计划
    Begin
      If Nvl(n_计划id, 0) = 0 Then
        --没计划生效,取安排的数据
        Select Nvl(序号, 0),
               To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(c.开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 时段时间,
               限制数量, Decode(Nvl(是否预约, 0), 0, 0, 限制数量)
        Into n_时段序号, d_时段时间, n_时段限号, n_时段限约
        From 挂号安排时段 C
        Where 安排id = n_安排id And 星期 = v_星期 And
              (序号, 安排id, 星期) In
              (Select Nvl(Max(c.序号), -1), 安排id, 星期
               From 挂号安排时段 C
               Where 安排id = n_安排id And c.星期 = v_星期 And
                     Decode(操作类型_In, 1, To_Char(Nvl(c.开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi'),
                            To_Char(发生时间_In, 'hh24:mi')) =
                     To_Char(Nvl(c.开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
               Group By 安排id, 星期);
      Else
        --有计划生效取计划
        --没生效，代表是从挂号计划时段查询      
        Select Nvl(序号, -1),
               To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(c.开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 时段时间,
               限制数量, Decode(Nvl(是否预约, 0), 0, 0, 限制数量)
        Into n_时段序号, d_时段时间, n_时段限号, n_时段限约
        From 挂号计划时段 C
        Where 计划id = n_计划id And 星期 = v_星期 And
              (序号, 计划id, 星期) In
              (Select Nvl(Max(c.序号), -1), 计划id, 星期
               From 挂号计划时段 C
               Where 计划id = n_计划id And c.星期 = v_星期 And
                     Decode(操作类型_In, 1, To_Char(Nvl(c.开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi'),
                            To_Char(发生时间_In, 'hh24:mi')) =
                     To_Char(Nvl(c.开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
               Group By 计划id, 星期);
      End If;
    Exception
      When Others Then
        n_时段序号 := -1;
        n_分时段   := 0;
        d_时段时间 := 发生时间_In;
        n_时段限号 := 0;
        n_时段限约 := 0;
    End;
  End If;

  If 序号_In = 1 Then
  
    --获取当前未使用的序号
    If Nvl(n_序号控制, 0) = 1 And n_分时段 = 0 Then
      --<序号控制 未设置时段 获取可用的最大序号,以及已经使用的数量>     
      Begin
        --最大序号
        If 退号重用_In = 1 Then
          Select Max(Nvl(序号, 0)), Sum(Decode(序号, Nvl(号序_In, 0), 0, 1)), Sum(Nvl(预约, 0))
          Into n_已用序号, n_已用数量, n_已约数
          From 挂号序号状态
          Where 号码 = 号别_In And 日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 状态 Not In (4, 5);
        Else
          Select Max(Nvl(序号, 0)), Sum(Decode(序号, Nvl(号序_In, 0), 0, 1)), Sum(Nvl(预约, 0))
          Into n_已用序号, n_已用数量, n_已约数
          From 挂号序号状态
          Where 号码 = 号别_In And 日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 状态 <> 5;
        End If;
      Exception
        When Others Then
          n_已用序号 := 0;
          n_已用数量 := 0;
      End;
      If n_序号 Is Null Then
        n_序号 := Nvl(n_已用序号, 0) + 1;
      End If;
      --<序号控制 未设置时段 获取可用的最大序号 以及已经使用的数量 --end>
    
      --非加号的情况需要检查是否超过了限制
      If 操作类型_In = 0 Then
        If Nvl(预约挂号_In, 0) = 0 Then
          --挂号时检查       
          --启用序号控制未分时段 达到了限制
          If n_限号数 <= n_已用数量 And n_限号数 > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(Trunc(发生时间_In), 'yyyy-mm-dd ') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        
        Else
          --预约时检查
          If n_限约数 = 0 Then
            n_限约数 := n_限号数;
          End If;
        
          If n_限约数 <= n_已约数 And n_限约数 > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(Trunc(发生时间_In), 'yyyy-mm-dd') || '已达到最大限约数！';
            Raise Err_Item;
          End If;
        End If;
      
      Else
        Null;
        --启用序号控制,未分时段 加号情况   不处理,如果以后有限制条件以后补充
      End If;
    
    Elsif Nvl(n_分时段, 0) > 0 And Nvl(n_序号控制, 0) = 0 Then
      --<--普通号分时段 这里只有预约一种情况-->
      If 操作类型_In = 0 Then
        --<正常预约挂号-->
        Begin
          Select Count(0) As 已挂数, Nvl(Sum(Decode(Nvl(Sign(a.日期 - d_时段时间), 0), 0, 1, 0)), 0) As 已约数
          Into n_已挂数, n_已约数
          From 挂号序号状态 A
          Where a.号码 = 号别_In And 日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And
                状态 Not In (4, 5);
        Exception
          When Others Then
            n_已挂数 := 0;
            n_已约数 := 0;
        End;
      
        n_时段限约 := n_时段限号; --普通号分时段的情况,29中n_时段限约始终是0 这里特殊处理
        --检查限制数量  
        If n_限约数 = 0 Then
          n_限约数 := n_限号数;
        End If;
        If n_时段限约 <= n_已约数 Or n_限约数 <= n_已约数 Then
          v_Err_Msg := '号别' || 号别_In || '在时段' || To_Char(d_时段时间, 'hh24:mi') || '已达到最大限约数！';
          Raise Err_Item;
        End If;
        If n_限号数 <= n_已挂数 Then
          v_Err_Msg := '号别' || 号别_In || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
          Raise Err_Item;
        End If;
      End If;
    
      --没有达到时段的限号数 号码在当前时段往后追加
    
      --获取当天挂出的最大号序
      If 退号重用_In = 1 Then
        Select Nvl(Max(序号), 0)
        Into n_挂出的最大序号
        From 挂号序号状态 A
        Where a.日期 = d_时段时间 And 号码 = 号别_In And 状态 Not In (4, 5);
      Else
        Select Nvl(Max(序号), 0)
        Into n_挂出的最大序号
        From 挂号序号状态 A
        Where a.日期 = d_时段时间 And 号码 = 号别_In And 状态 <> 5;
      End If;
    
      --设置序号
      n_序号 := RPad(Nvl(n_时段序号, 0), Length(n_限号数) + Length(Nvl(n_时段序号, 0)), 0) + n_已约数 + 1;
      If n_序号 <= Nvl(n_挂出的最大序号, 0) Then
        n_序号 := Nvl(n_挂出的最大序号, 0) + 1;
      End If;
    
      --<--普通号分时段--End>
    Elsif Nvl(n_分时段, 0) > 0 And Nvl(n_序号控制, 0) = 1 Then
      --<启用序号控制 设置时段
      --专家号分时段
      Begin
        If 退号重用_In = 1 Then
          Select Max(序号), Sum(Decode(序号, Nvl(号序_In, 0), 0, 1)), Sum(Decode(Sign(日期 - d_时段时间), 0, 1, 0))
          Into n_已用序号, n_已挂数, n_已用数量
          From 挂号序号状态
          Where 号码 = 号别_In And 日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 状态 Not In (4, 5);
        Else
          Select Max(序号), Sum(Decode(序号, Nvl(号序_In, 0), 0, 1)), Sum(Decode(Sign(日期 - d_时段时间), 0, 1, 0))
          Into n_已用序号, n_已挂数, n_已用数量
          From 挂号序号状态
          Where 号码 = 号别_In And 日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 状态 <> 5;
        End If;
      Exception
        When Others Then
          n_已用序号 := 0;
          n_已用数量 := 0;
          n_已挂数   := 0;
      End;
    
      n_失效数 := 0;
      If Nvl(预约挂号_In, 0) = 0 Then
        n_预约有效时间 := Zl_To_Number(zl_GetSysParameter('预约有效时间', 1111));
        n_失约挂号     := Zl_To_Number(zl_GetSysParameter('失约用于挂号', 1111));
        If Nvl(n_预约有效时间, 0) <> 0 And Nvl(n_失约挂号, 0) > 0 Then
          Begin
            Select Sum(Decode(Sign((Sysdate - n_预约有效时间 / 24 / 60) - 日期), 1, 1, 0))
            Into n_失效数
            From 挂号序号状态
            Where 号码 = 号别_In And 日期 Between Trunc(Sysdate) And Sysdate And Nvl(预约, 0) = 1 And 状态 = 2;
          Exception
            When Others Then
              n_失效数 := 0;
          End;
        End If;
      End If;
    
      If 操作类型_In = 0 Then
        If Nvl(预约挂号_In, 0) = 0 Then
        
          --挂号 追加号码时不检查时段限号数
          If n_时段限号 <= n_已用数量 And Nvl(n_追加号, 0) = 0 Then
            v_Err_Msg := '号别' || 号别_In || '在时段' || To_Char(d_时段时间, 'hh24:mi') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
          If n_限号数 <= n_已挂数 - n_失效数 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        
        Else
          --挂号
          If n_限约数 = 0 Then
            n_限约数 := n_限号数;
          End If;
          If n_限约数 <= n_已用数量 Then
            v_Err_Msg := '号别' || 号别_In || '在时段' || To_Char(d_时段时间, 'hh24:mi') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
          If n_限号数 <= n_已挂数 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        End If;
      End If;
      If n_序号 Is Null Then
        --设置序号
        If Nvl(n_已用序号, 0) < Nvl(n_时段序号, 0) Then
          n_已用序号 := Nvl(n_时段序号, 0);
        End If;
        n_序号 := Nvl(n_已用序号, 0) + 1;
      End If;
    Elsif Nvl(n_分时段, 0) = 0 And Nvl(n_序号控制, 0) = 0 And Nvl(病历费_In, 0) = 0 And Nvl(号别_In, 0) > 0 Then
      ---<--普通号  -->
      Begin
        Select 已挂数, 已约数
        Into n_已用数量, n_已约数
        From 病人挂号汇总
        Where 日期 = Trunc(发生时间_In) And 号码 = 号别_In;
      Exception
        When Others Then
          n_已用数量 := 0;
          n_已约数   := 0;
      End;
      If Nvl(操作类型_In, 0) = 0 Then
        If Nvl(预约挂号_In, 0) = 0 Then
          --挂号
          If Nvl(n_已用数量, 0) >= Nvl(n_限号数, 0) And Nvl(n_限号数, 0) > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        Else
          --预约
          If (Nvl(n_限约数, 0) > 0 And Nvl(n_已约数, 0) >= Nvl(n_限约数, 0)) Or
             (Nvl(n_已用数量, 0) >= Nvl(n_限号数, 0) And Nvl(n_限号数, 0) > 0) Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        End If;
      End If;
      ---<--普通号  -->
    End If;
  End If;

  --更新挂号序号状态
  If 序号_In = 1 And Not n_序号 Is Null Then
    If n_分时段 = 1 Then
      d_序号时间 := 发生时间_In;
    Else
      d_序号时间 := Trunc(发生时间_In);
    End If;
    --锁定序号的处理
    Begin
      Select 操作员姓名, 机器名
      Into v_序号操作员, v_序号机器名
      From 挂号序号状态
      Where 状态 = 5 And 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号;
      n_锁定 := 1;
    Exception
      When Others Then
        v_序号操作员 := Null;
        v_序号机器名 := Null;
        n_锁定       := 0;
    End;
    If n_锁定 = 0 Then
      Update 挂号序号状态
      Set 状态 = Decode(预约挂号_In, 1, 2, 1), 预约 = Decode(预约接收_In, 1, 1, 0), 登记时间 = Sysdate
      Where 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号 And 状态 = 3 And 操作员姓名 = 操作员姓名_In;
      If Sql%RowCount = 0 Then
        Begin
          If Nvl(n_分时段, 0) = 0 Or Nvl(预约挂号_In, 0) = 1 Or (Nvl(n_序号控制, 0) = 0 And Nvl(号序_In, 0) = 0) Then
            Insert Into 挂号序号状态
              (号码, 日期, 序号, 状态, 操作员姓名, 预约, 登记时间)
            Values
              (号别_In, d_序号时间, n_序号, Decode(预约挂号_In, 1, 2, 1), 操作员姓名_In, Decode(预约接收_In, 1, 1, 0), Sysdate);
            If Nvl(预约挂号_In, 0) = 0 And 预约方式_In Is Not Null Then
              Update 挂号序号状态 Set 预约 = 1 Where 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号;
            End If;
          Elsif Nvl(n_分时段, 0) > 0 Then
            --分时段后专家号 失约的预约号允许挂号
            Update 挂号序号状态
            Set 状态 = Decode(预约挂号_In, 1, 2, 1), 操作员姓名 = 操作员姓名_In, 预约 = Decode(预约接收_In, 1, 1, 0), 登记时间 = Sysdate
            Where 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号 And 状态 = 2;
            If Sql%NotFound Then
              Insert Into 挂号序号状态
                (号码, 日期, 序号, 状态, 操作员姓名, 预约, 登记时间)
              Values
                (号别_In, d_序号时间, n_序号, Decode(预约挂号_In, 1, 2, 1), 操作员姓名_In, Decode(预约接收_In, 1, 1, 0), Sysdate);
            End If;
            If Nvl(预约挂号_In, 0) = 0 And 预约方式_In Is Not Null Then
              Update 挂号序号状态 Set 预约 = 1 Where 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号;
            End If;
          End If;
        Exception
          When Others Then
            v_Err_Msg := '序号' || n_序号 || '已被使用,请重新选择一个序号.';
            Raise Err_Item;
        End;
      End If;
    Else
      If 操作员姓名_In <> v_序号操作员 Or v_机器名 <> v_序号机器名 Then
        v_Err_Msg := '序号' || n_序号 || '已被自助机' || v_机器名 || '锁定,请重新选择一个序号.';
        Raise Err_Item;
      Else
        Update 挂号序号状态
        Set 状态 = Decode(预约挂号_In, 1, 2, 1), 预约 = Decode(预约接收_In, 1, 1, 0), 登记时间 = Sysdate
        Where 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号 And 状态 = 5 And 操作员姓名 = 操作员姓名_In And 机器名 = v_机器名;
        If Nvl(预约挂号_In, 0) = 0 And 预约方式_In Is Not Null Then
          Update 挂号序号状态 Set 预约 = 1 Where 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号;
        End If;
      End If;
    End If;
  End If;

  --产生病人挂号费用(可能单独是或包括病历费用)
  Select 病人费用记录_Id.Nextval Into n_费用id From Dual; --应该通过程序得到

  Insert Into 门诊费用记录
    (ID, 记录性质, 记录状态, 序号, 价格父号, 从属父号, NO, 实际票号, 门诊标志, 加班标志, 附加标志, 发药窗口, 病人id, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 病人科室id, 收费类别,
     计算单位, 收费细目id, 收入项目id, 收据费目, 付数, 数次, 标准单价, 应收金额, 实收金额, 结帐金额, 结帐id, 记帐费用, 开单部门id, 开单人, 划价人, 执行部门id, 执行人, 操作员编号, 操作员姓名,
     发生时间, 登记时间, 保险大类id, 保险项目否, 保险编码, 统筹金额, 摘要, 结论, 缴款组id)
  Values
    (n_费用id, 4, Decode(预约挂号_In, 1, 0, 1), 序号_In, Decode(价格父号_In, 0, Null, 价格父号_In), 从属父号_In, 单据号_In, 票据号_In, 1, 急诊_In,
     病历费_In, Decode(预约挂号_In, 1, To_Char(n_序号), 诊室_In), Decode(病人id_In, 0, Null, 病人id_In),
     Decode(门诊号_In, 0, Null, 门诊号_In), 付款方式_In, 姓名_In, Decode(姓名_In, Null, Null, 性别_In), Decode(姓名_In, Null, Null, 年龄_In),
     v_费别, 病人科室id_In, 收费类别_In, 号别_In, 收费细目id_In, 收入项目id_In, 收据费目_In, 1, 数次_In, 标准单价_In, 应收金额_In, 实收金额_In,
     Decode(预约挂号_In, 1, Null, Decode(Nvl(记帐费用_In, 0), 1, Null, 实收金额_In)),
     Decode(预约挂号_In, 1, Null, Decode(Nvl(记帐费用_In, 0), 1, Null, 结帐id_In)), Decode(Nvl(记帐费用_In, 0), 1, 1, 0), 开单部门id_In,
     操作员姓名_In, Decode(预约挂号_In, 1, 操作员姓名_In, Null), 执行部门id_In, 医生姓名_In, 操作员编号_In, 操作员姓名_In, 发生时间_In, 登记时间_In, 保险大类id_In,
     保险项目否_In, 保险编码_In, 统筹金额_In, 摘要_In, 预约方式_In, Decode(预约挂号_In, 1, Null, n_组id));

  --汇总结算到病人预交记录
  If Nvl(预约挂号_In, 0) = 0 And Nvl(记帐费用_In, 0) = 0 Then
  
    If (Nvl(现金支付_In, 0) <> 0 Or (Nvl(现金支付_In, 0) = 0 And Nvl(个帐支付_In, 0) = 0 And Nvl(预交支付_In, 0) = 0)) And 序号_In = 1 Then
      Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      Insert Into 病人预交记录
        (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
         结算性质)
      Values
        (n_预交id, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(结算方式_In, v_现金), Nvl(现金支付_In, 0), 登记时间_In,
         操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费', n_组id, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In, 4);
    
      If Nvl(结算卡序号_In, 0) <> 0 And Nvl(现金支付_In, 0) <> 0 Then
      
        n_消费卡id := Null;
        Begin
          Select Nvl(自制卡, 0), 1 Into n_自制卡, n_Count From 卡消费接口目录 Where 编号 = 结算卡序号_In;
        Exception
          When Others Then
            n_Count := 0;
        End;
        If n_Count = 0 Then
          v_Err_Msg := '没有发现原结算卡的相应类别,不能继续操作！';
          Raise Err_Item;
        End If;
        If n_自制卡 = 1 Then
          Select ID
          Into n_消费卡id
          From 消费卡目录
          Where 接口编号 = 结算卡序号_In And 卡号 = 卡号_In And
                序号 = (Select Max(序号) From 消费卡目录 Where 接口编号 = 结算卡序号_In And 卡号 = 卡号_In);
        End If;
        Zl_病人卡结算记录_Insert(结算卡序号_In, n_消费卡id, 结算方式_In, 现金支付_In, 卡号_In, Null, 登记时间_In, Null, 结帐id_In, n_预交id);
      End If;
    
    End If;
  
    --对于医保挂号
    If Nvl(个帐支付_In, 0) <> 0 And 序号_In = 1 Then
      Insert Into 病人预交记录
        (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 结算性质)
      Values
        (病人预交记录_Id.Nextval, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), v_个人帐户, 个帐支付_In, 登记时间_In, 操作员编号_In,
         操作员姓名_In, 结帐id_In, '医保挂号', n_组id, 4);
    End If;
  
    --对于就诊卡通过预交金挂号
    If Nvl(预交支付_In, 0) <> 0 And 序号_In = 1 Then
      n_预交金额 := 预交支付_In;
      For r_Deposit In c_Deposit(病人id_In, v_冲预交病人ids) Loop
        n_当前金额 := Case
                    When r_Deposit.金额 - n_预交金额 < 0 Then
                     r_Deposit.金额
                    Else
                     n_预交金额
                  End;
      
        If r_Deposit.Id <> 0 Then
          --第一次冲预交(填上结帐ID,金额为0)
          Update 病人预交记录 Set 冲预交 = 0, 结帐id = 结帐id_In, 结算性质 = 4 Where ID = r_Deposit.Id;
        
        End If;
        --冲上次剩余额
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 预交类别, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,
           冲预交, 结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
          Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 预交类别, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号,
                 登记时间_In, 操作员姓名_In, 操作员编号_In, n_当前金额, 结帐id_In, n_组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
          From 病人预交记录
          Where NO = r_Deposit.No And 记录状态 = r_Deposit.记录状态 And 记录性质 In (1, 11) And Rownum = 1;
      
        --更新病人预交余额
        Update 病人余额
        Set 预交余额 = Nvl(预交余额, 0) - n_当前金额
        Where 病人id = r_Deposit.病人id And 性质 = 1 And 类型 = Nvl(r_Deposit.预交类别, 2);
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
      If n_预交金额 > 0 Then
        v_Err_Msg := '预交余不够支付本次支付金额,不能继续操作！';
        Raise Err_Item;
      
      End If;
      Delete From 病人余额 Where 病人id = 病人id_In And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
    End If;
  
    --相关汇总表的处理
    --人员缴款余额
    If 序号_In = 1 Then
      If Nvl(现金支付_In, 0) <> 0 Then
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) + 现金支付_In
        Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = Nvl(结算方式_In, v_现金)
        Returning 余额 Into n_返回值;
      
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额
            (收款员, 结算方式, 性质, 余额)
          Values
            (操作员姓名_In, Nvl(结算方式_In, v_现金), 1, 现金支付_In);
          n_返回值 := 现金支付_In;
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 人员缴款余额
          Where 收款员 = 操作员姓名_In And 结算方式 = Nvl(结算方式_In, v_现金) And 性质 = 1 And Nvl(余额, 0) = 0;
        End If;
      End If;
    
      If Nvl(个帐支付_In, 0) <> 0 Then
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) + 个帐支付_In
        Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = v_个人帐户
        Returning 余额 Into n_返回值;
      
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (操作员姓名_In, v_个人帐户, 1, 个帐支付_In);
          n_返回值 := 个帐支付_In;
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 人员缴款余额
          Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_个人帐户 And Nvl(余额, 0) = 0;
        End If;
      End If;
    End If;
  End If;

  --病人挂号汇总(只处理一次,且单独收取病历费不处理)
  If Nvl(预约挂号_In, 0) = 0 Then
    If Nvl(记帐费用_In, 0) = 0 Then
      --处理票据使用情况
      If 序号_In = 1 And 票据号_In Is Not Null Then
        Select 票据打印内容_Id.Nextval Into n_打印id From Dual;
      
        --发出票据
        Insert Into 票据打印内容 (ID, 数据性质, NO) Values (n_打印id, 4, 单据号_In);
      
        Insert Into 票据使用明细
          (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
        Values
          (票据使用明细_Id.Nextval, Decode(收费票据_In, 1, 1, 4), 票据号_In, 1, 1, 领用id_In, n_打印id, 登记时间_In, 操作员姓名_In);
      
        --状态改动
        Update 票据领用记录
        Set 当前号码 = 票据号_In, 剩余数量 = Decode(Sign(剩余数量 - 1), -1, 0, 剩余数量 - 1), 使用时间 = Sysdate
        Where ID = Nvl(领用id_In, 0);
      End If;
    End If;
    --病人本次就诊(以发生时间为准)
    If Nvl(病人id_In, 0) <> 0 And 序号_In = 1 Then
      Update 病人信息 Set 就诊时间 = 发生时间_In, 就诊状态 = 1, 就诊诊室 = 诊室_In Where 病人id = 病人id_In;
    End If;
  End If;

  If Nvl(预约挂号_In, 0) = 0 And Nvl(记帐费用_In, 0) = 1 Then
    --记帐
    If Nvl(病人id_In, 0) = 0 Then
      v_Err_Msg := '要针对病人的挂号费进行记帐，必须是建档病人才能记帐挂号。';
      Raise Err_Item;
    End If;
  
    --病人余额
    Update 病人余额
    Set 费用余额 = Nvl(费用余额, 0) + Nvl(实收金额_In, 0)
    Where 病人id = Nvl(病人id_In, 0) And 性质 = 1 And 类型 = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人余额 (病人id, 性质, 类型, 费用余额, 预交余额) Values (病人id_In, 1, 1, Nvl(实收金额_In, 0), 0);
    End If;
  
    --病人未结费用
    Update 病人未结费用
    Set 金额 = Nvl(金额, 0) + Nvl(实收金额_In, 0)
    Where 病人id = 病人id_In And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(病人科室id, 0) = Nvl(病人科室id_In, 0) And
          Nvl(开单部门id, 0) = Nvl(开单部门id_In, 0) And Nvl(执行部门id, 0) = Nvl(执行部门id_In, 0) And 收入项目id + 0 = 收入项目id_In And
          来源途径 + 0 = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人未结费用
        (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
      Values
        (病人id_In, Null, Null, 病人科室id_In, 开单部门id_In, 执行部门id_In, 收入项目id_In, 1, Nvl(实收金额_In, 0));
    End If;
  End If;

  --病人挂号记录
  If 号别_In Is Not Null And 序号_In = 1 Then
    --And Nvl(预约挂号_In, 0) = 0
    Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
    Begin
      Select 名称 Into v_付款方式 From 医疗付款方式 Where 编码 = 付款方式_In And Rownum < 2;
    Exception
      When Others Then
        v_付款方式 := Null;
    End;
    Insert Into 病人挂号记录
      (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 预约时间, 操作员编号,
       操作员姓名, 复诊, 号序, 社区, 预约, 预约方式, 摘要, 交易流水号, 交易说明, 合作单位, 接收时间, 接收人, 预约操作员, 预约操作员编号, 险类, 医疗付款方式)
    Values
      (n_挂号id, 单据号_In, Decode(Nvl(预约挂号_In, 0), 1, 2, 1), 1, 病人id_In, 门诊号_In, 姓名_In, 性别_In, 年龄_In, 号别_In, 急诊_In, 诊室_In,
       Null, 执行部门id_In, 医生姓名_In, 0, Null, 登记时间_In, 发生时间_In, Decode(Nvl(预约挂号_In, 0), 1, 发生时间_In, Null), 操作员编号_In,
       操作员姓名_In, 复诊_In, n_序号, 社区_In, Decode(预约接收_In, 1, 1, 0), 预约方式_In, 摘要_In, 交易流水号_In, 交易说明_In, 合作单位_In,
       Decode(Nvl(预约挂号_In, 0), 0, 登记时间_In, Null), Decode(Nvl(预约挂号_In, 0), 0, 操作员姓名_In, Null),
       Decode(Nvl(预约挂号_In, 0), 1, 操作员姓名_In, Null), Decode(Nvl(预约挂号_In, 0), 1, 操作员编号_In, Null),
       Decode(Nvl(险类_In, 0), 0, Null, 险类_In), v_付款方式);
    If Nvl(预约挂号_In, 0) = 0 And 预约方式_In Is Not Null Then
      Update 病人挂号记录
      Set 预约 = 1, 预约时间 = 发生时间_In, 预约操作员 = 操作员姓名_In, 预约操作员编号 = 操作员编号_In
      Where ID = n_挂号id;
    End If;
    n_预约生成队列 := 0;
    If Nvl(预约挂号_In, 0) = 1 Then
      n_预约生成队列 := Zl_To_Number(zl_GetSysParameter('预约生成队列', 1113));
    End If;
  
    --0-不产生队列;1-按医生或分诊台排队;2-先分诊,后医生站
    If Nvl(生成队列_In, 0) <> 0 And Nvl(预约挂号_In, 0) = 0 Or n_预约生成队列 = 1 Then
      n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113));
      If Nvl(n_分诊台签到排队, 0) = 0 Or n_预约生成队列 = 1 Then
      
        --产生队列
        --.按”执行部门” 的方式生成队列
        v_队列名称 := 执行部门id_In;
        v_排队号码 := Zlgetnextqueue(执行部门id_In, n_挂号id, 号别_In || '|' || n_序号);
        v_排队序号 := Zlgetsequencenum(0, n_挂号id, 0);
        d_排队时间 := Zl_Get_Queuedate(n_挂号id, 号别_In, n_序号, 发生时间_In);
        --  队列名称_In , 业务类型_In, 业务id_In,科室id_In,排队号码_In,排队标记_In,患者姓名_In,病人ID_IN, 诊室_In, 医生姓名_In, 排队时间_In
        Zl_排队叫号队列_Insert(v_队列名称, 0, n_挂号id, 执行部门id_In, v_排队号码, Null, 姓名_In, 病人id_In, 诊室_In, 医生姓名_In, d_排队时间, 预约方式_In,
                         Null, v_排队序号);
      End If;
    End If;
  End If;
  --病人担保信息
  If 病人id_In Is Not Null And 序号_In = 1 Then
    --取参数:
    If Nvl(n_结算模式, 0) <> Nvl(结算模式_In, 0) Then
      --结算模式的确定
    
      v_Err_Msg := Null;
      Begin
        Select Nvl(结算模式, 0) Into n_结算模式 From 病人信息 Where 病人id = 病人id_In;
      Exception
        When Others Then
          v_Err_Msg := '未找到指定的病人信息,不允许挂号';
      End;
    
      If v_Err_Msg Is Not Null Then
        Raise Err_Item;
      End If;
      If n_结算模式 = 1 And Nvl(结算模式_In, 0) = 0 Then
        --病人已经是"先诊疗后结算的",本次是"先结算后诊疗的",则检查是否存在未结数据
        Select Count(1)
        Into n_Count
        From 病人未结费用
        Where 病人id = 病人id_In And (来源途径 In (1, 4) Or 来源途径 = 3 And Nvl(主页id, 0) = 0) And Nvl(金额, 0) <> 0 And Rownum < 2;
        If Nvl(n_Count, 0) <> 0 Then
          --存在未结算数据，必须先结算后才允许执行
          v_Err_Msg := '当前病人的就诊模式为先诊疗后结算且存在未结费用，不允许调整该病人的就诊模式,你可以先对未结费用结帐后再挂号或不调整病人的就诊模式!';
          Raise Err_Item;
        End If;
        --检查
        --未发生医嘱业务的（即当时就挂号的,需要保证同一次的就诊模式是一至的(程序已经检查，不用再处理)
      End If;
      Update 病人信息 Set 结算模式 = 结算模式_In Where 病人id = 病人id_In;
    End If;
  
    Update 病人信息
    Set 担保人 = Null, 担保额 = Null, 担保性质 = Null
    Where 病人id = 病人id_In And Exists (Select 1
           From 病人担保记录
           Where 病人id = 病人id_In And 主页id Is Not Null And
                 登记时间 = (Select Max(登记时间) From 病人担保记录 Where 病人id = 病人id_In));
  
    If Sql%RowCount > 0 Then
      Update 病人担保记录
      Set 到期时间 = Sysdate
      Where 病人id = 病人id_In And 主页id Is Not Null And Nvl(到期时间, Sysdate) > Sysdate;
    End If;
  End If;
  If 序号_In = 1 Then
    --消息推送
    Begin
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 1, n_挂号id;
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
End Zl_病人挂号记录_Insert;
/

Create Or Replace Procedure Zl_病人挂号记录_出诊_Delete
(
  单据号_In       门诊费用记录.No%Type,
  操作员编号_In   门诊费用记录.操作员编号%Type,
  操作员姓名_In   门诊费用记录.操作员姓名%Type,
  摘要_In         门诊费用记录.摘要%Type := Null, --预约取消时 填写 存放预约取消原因
  删除门诊号_In   Number := 0,
  非原样退结算_In Varchar2 := Null,
  退费类型_In     In Number := 0, --0-全退 1-退挂号费 2-退病历费
  退指定结算_In   病人预交记录.结算方式%Type := Null,
  退号重用_In     Number := 1,
  结算方式_In     Varchar2 := Null
) As
  --退费类型_In,在一下几种情况下不准进行部分退费
  --    2.三方接口,暂时不支持
  -- 挂号费病历费分开退,规则
  --    普通结算方式:原结算方式退部分费用
  --    预交款:预交款,退部分
  --    预交款与普通结算方式混合:退款按照普通结算方式部分退
  --    消费卡:原样将费用部分退入消费卡
  --非原样退结算_In:指不能退还给原样结算方式(如医保的个人账户,三方账户的退现等),多个用逗分离
  --退指定结算_IN:指非原样退结算部分,应该退给哪种结算方式,为空时缺省退给现金,否则退给指定的结算方式

  --该游标用于判断是否单独收病历费,及挂号汇总表处理
  Cursor c_Registinfo(v_状态 门诊费用记录.记录状态%Type) Is
    Select a.发生时间, a.登记时间, c.接收时间, a.收费细目id As 项目id, c.执行部门id As 科室id, c.执行人 As 医生姓名, d.Id As 医生id, c.号别 As 号码
    From 门诊费用记录 A, 病人挂号记录 C, 人员表 D
    Where a.记录性质 = 4 And a.记录状态 = v_状态 And c.No = a.No And c.执行人 = d.姓名(+) And a.No = 单据号_In And
          Nvl(a.计算单位, '号别') = c.号别 And Rownum < 2;
  r_Registrow c_Registinfo%RowType;

  --该游标用于判断记录是否存在,及费用汇总表处理
  Cursor c_Moneyinfo(v_状态 门诊费用记录.记录状态%Type) Is
    Select 病人科室id, 开单部门id, 执行部门id, 收入项目id, Nvl(Sum(应收金额), 0) As 应收, Nvl(Sum(实收金额), 0) As 实收, Nvl(Sum(结帐金额), 0) As 结帐
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 = v_状态 And NO = 单据号_In
    Group By 病人科室id, 开单部门id, 执行部门id, 收入项目id;
  r_Moneyrow c_Moneyinfo%RowType;

  --该光标用于处理人员缴款余额中退的不同结算方式的金额
  Cursor c_Opermoney Is
    Select Distinct b.结算方式, -1 * Nvl(b.冲预交, 0) As 冲预交
    From 门诊费用记录 A, 病人预交记录 B
    Where a.结帐id = b.结帐id And a.No = 单据号_In And a.记录性质 = 4 And a.记录状态 = 2 And b.记录性质 = 4 And b.记录状态 = 2 And
          Nvl(b.冲预交, 0) <> 0 And
          Nvl(a.附加标志, 0) =
          Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(a.附加标志, 0), 1, -1, Nvl(a.附加标志, 0)), Nvl(a.附加标志, 0));

  v_Err_Msg Varchar(255);
  Err_Item Exception;

  n_结帐id 病人预交记录.结帐id%Type;
  n_销帐id 门诊费用记录.结帐id%Type;

  v_退指定结算方式 病人预交记录.结算方式%Type;
  n_退款金额       病人预交记录.冲预交%Type;
  n_打印id         票据打印内容.Id%Type;
  n_病人id         病人信息.病人id%Type;
  n_退费金额       病人预交记录.冲预交%Type;
  n_预交金额       病人预交记录.冲预交%Type; --原记录 预交缴款金额
  n_返回值         病人余额.预交余额%Type;
  n_挂号id         病人挂号记录.Id%Type;
  n_组id           财务缴款分组.Id%Type;

  n_二次退费       Number; --记录是否是此单据的第二次退费
  n_分诊台签到排队 Number;
  n_预约生成队列   Number;
  n_预约挂号       Number;
  n_挂号生成队列   Number;
  d_Date           Date;
  n_记帐           门诊费用记录.记帐费用%Type;
  n_病人id1        病人信息.病人id%Type;
  n_返回额         门诊费用记录.实收金额%Type;
  n_已结帐         Number;
  n_序号           病人挂号记录.号序%Type;
  n_就诊病人id     病人信息.病人id%Type;
  d_就诊时间       就诊登记记录.就诊时间%Type;
  n_出诊记录id     临床出诊记录.Id%Type;
  v_结算内容       Varchar2(5000);
  v_当前结算       Varchar2(1000);
  v_结算方式       病人预交记录.结算方式%Type;
  n_三方卡标志     Number;
  n_结算金额       病人预交记录.冲预交%Type;
Begin
  n_组id           := Zl_Get组id(操作员姓名_In);
  v_退指定结算方式 := 退指定结算_In;

  Select 出诊记录id, 号序 Into n_出诊记录id, n_序号 From 病人挂号记录 Where NO = 单据号_In And Rownum < 2;

  --首先判断要退号/取消预约的记录是否存在
  Open c_Moneyinfo(1);
  Fetch c_Moneyinfo
    Into r_Moneyrow;
  If c_Moneyinfo%RowCount = 0 Then
    Close c_Moneyinfo;
    Open c_Moneyinfo(0);
    Fetch c_Moneyinfo
      Into r_Moneyrow;
    If c_Moneyinfo%RowCount = 0 Then
      v_Err_Msg := '要处理的单据不存在。';
      Raise Err_Item;
    End If;
    n_预约挂号 := 1;
  End If;
  Close c_Moneyinfo;

  --1.预约处理
  If Nvl(n_预约挂号, 0) = 1 Then
    --减少已约数
    Open c_Registinfo(0);
    Fetch c_Registinfo
      Into r_Registrow;
  
    Update 临床出诊记录 Set 已约数 = Nvl(已约数, 0) - 1 Where ID = n_出诊记录id;
    Update 病人挂号汇总
    Set 已约数 = Nvl(已约数, 0) - 1
    Where 日期 = Trunc(r_Registrow.发生时间) And Nvl(医生id, 0) = Nvl(r_Registrow.医生id, 0) And
          Nvl(医生姓名, '-') = Nvl(r_Registrow.医生姓名, '-') And 科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And
          (号码 = r_Registrow.号码 Or 号码 Is Null);
  
    If Sql%RowCount = 0 Then
      Insert Into 病人挂号汇总
        (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已约数)
      Values
        (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
         Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, -1);
    End If;
  
    Close c_Registinfo;
  
    --更新挂号序号状态
    Update 临床出诊序号控制
    Set 挂号状态 = 0, 操作员姓名 = Null
    Where 挂号状态 = 2 And 记录id = n_出诊记录id And 序号 = n_序号;
  
    Update 临床出诊序号控制
    Set 挂号状态 = 4, 操作员姓名 = Null
    Where 挂号状态 = 2 And 记录id = n_出诊记录id And 备注 = n_序号;
  
    --添加病人挂号记录的 冲销记录
    Select 病人挂号记录_Id.Nextval, Sysdate Into n_挂号id, d_Date From Dual;
  
    Update 病人挂号记录 Set 记录状态 = 3 Where NO = 单据号_In And 记录状态 = 1 And 记录性质 = 2;
    If Sql%NotFound Then
      v_Err_Msg := '预约单【' || 单据号_In || '】不存在或由于并发原因已经被取消预约';
      Raise Err_Item;
    End If;
  
    Insert Into 病人挂号记录
      (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 操作员编号, 操作员姓名,
       复诊, 号序, 社区, 预约, 摘要, 预约方式, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 医疗付款方式, 出诊记录id, 预约操作员, 预约操作员编号)
      Select n_挂号id, NO, 记录性质, 2, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, d_Date, 发生时间,
             操作员编号_In, 操作员姓名_In, 复诊, 号序, 社区, 预约, Nvl(摘要_In, 摘要) As 摘要, 预约方式, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 医疗付款方式,
             n_出诊记录id, 预约操作员, 预约操作员编号
      From 病人挂号记录
      Where NO = 单据号_In;
  
    Update 门诊费用记录 Set 记录状态 = 3 Where NO = 单据号_In And 记录性质 = 4 And 记录状态 = 0;
    Insert Into 门诊费用记录
      (ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 记帐单id, 病人id, 医嘱序号, 门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号, 付款方式, 病人科室id, 费别, 收费类别,
       收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id,
       执行人, 执行状态, 执行时间, 结论, 操作员编号, 操作员姓名, 结帐id, 结帐金额, 保险大类id, 保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要, 是否急诊, 缴款组id, 费用状态, 待转出,
       挂号id, 主页id)
      Select 病人费用记录_Id.Nextval, 记录性质, NO, 实际票号, 2, 序号, 从属父号, 价格父号, 记帐单id, 病人id, 医嘱序号, 门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号, 付款方式,
             病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, -1 * 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, -1 * 应收金额,
             -1 * 实收金额, 划价人, 开单部门id, 开单人, 发生时间, d_Date, 执行部门id, 执行人, -1, 执行时间, 结论, 操作员编号_In, 操作员姓名_In, Null, Null,
             保险大类id, 保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要, 是否急诊, 缴款组id, 费用状态, 待转出, 挂号id, 主页id
      From 门诊费用记录
      Where NO = 单据号_In And 记录性质 = 4 And 记录状态 = 3;
  
    --如果预约生成队列时需要清除队列
  
    n_预约生成队列 := Zl_To_Number(zl_GetSysParameter('预约生成队列', 1113));
    If Nvl(n_预约生成队列, 0) = 1 Then
      --要删除队列
      For v_挂号 In (Select ID, 诊室, 执行部门id, 执行人 From 病人挂号记录 Where NO = 单据号_In) Loop
        Zl_排队叫号队列_Delete(v_挂号.执行部门id, v_挂号.Id);
      End Loop;
    End If;
    Return;
  End If;

  Select Nvl(记帐费用, 0), 病人id, Decode(Sign(Nvl(结帐id, 0)), 0, 0, 1)
  Into n_记帐, n_病人id, n_已结帐
  From 门诊费用记录
  Where 记录性质 = 4 And NO = 单据号_In And 记录状态 In (1, 3) And Rownum < 2;

  --2.挂号处理
  n_已结帐 := Nvl(n_已结帐, 0);

  If n_已结帐 = 1 And n_记帐 = 1 Then
    Select Sysdate, Null Into d_Date, n_销帐id From Dual;
  Else
    Select Sysdate, 病人结帐记录_Id.Nextval Into d_Date, n_销帐id From Dual;
  End If;

  ----0-全退 1-退挂号费 2-退病历费
  If Nvl(退费类型_In, 0) <> 2 Then
    --不是光退病历费时处理
    --更新挂号序号状态
    If 退号重用_In = 1 Then
      Update 临床出诊序号控制
      Set 挂号状态 = 0, 操作员姓名 = Null
      Where 挂号状态 = 1 And 记录id = n_出诊记录id And 序号 = n_序号;
    
      Update 临床出诊序号控制
      Set 挂号状态 = 4, 操作员姓名 = Null
      Where 挂号状态 = 1 And 记录id = n_出诊记录id And 备注 = n_序号;
    Else
      Update 临床出诊序号控制
      Set 挂号状态 = 4, 操作员姓名 = 操作员姓名_In
      Where 挂号状态 = 1 And 记录id = n_出诊记录id And (序号 = n_序号 Or 备注 = n_序号);
    End If;
  
    --病人就诊状态
    If n_病人id Is Not Null Then
      Update 病人信息 Set 就诊状态 = 0, 就诊诊室 = Null Where 病人id = n_病人id;
    
      --删除门诊号相关处理,只有当只有一条挂号记录并且病人建档日期与挂号日期近似时才会处理
      If 删除门诊号_In = 1 Then
        Delete 门诊病案记录 Where 病人id = n_病人id;
        Update 病人信息 Set 门诊号 = Null Where 病人id = n_病人id;
        --费用记录包括挂号及病案、就诊卡费用,以及病人交费后退费或销帐的费用,挂号记录在最后处理
        Update 门诊费用记录 Set 标识号 = Null Where 门诊标志 = 1 And 病人id = n_病人id;
      End If;
    End If;
  
    --如果挂时收了就诊卡费,退费时清除就诊卡号,在非光退病历费时
    n_病人id1 := Null;
    Begin
      Select 病人id
      Into n_病人id1
      From 门诊费用记录
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And 附加标志 = 2 And Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
  
    If n_病人id1 Is Not Null And Nvl(退费类型_In, 0) <> 2 Then
      Update 病人信息
      Set 就诊卡号 = Null, 卡验证码 = Null, Ic卡号 = Decode(Ic卡号, 就诊卡号, Null, Ic卡号)
      Where 病人id = n_病人id1;
    End If;
  
  End If;

  --检查前面是否已经部分退过费用
  Begin
    Select 1 Into n_二次退费 From 门诊费用记录 Where 记录性质 = 4 And NO = 单据号_In And 记录状态 = 3 And Rownum < 2;
  Exception
    When Others Then
      n_二次退费 := 0;
  End;

  --门诊费用记录
  --冲销记录
  Insert Into 门诊费用记录
    (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位, 付数,
     数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 执行人, 操作员编号, 操作员姓名, 发生时间, 登记时间,
     结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 缴款组id)
    Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别,
           收费细目id, 计算单位, 付数, -数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, -应收金额, -实收金额, 开单部门id, 开单人, 执行部门id, 执行人,
           操作员编号_In, 操作员姓名_In, 发生时间, d_Date, n_销帐id,
           Decode(n_记帐, 1, Decode(Nvl(n_已结帐, 0), 0, -1 * 实收金额, Null), -1 * 结帐金额), 保险项目否, 保险大类id, -1 * 统筹金额,
           Nvl(摘要_In, 摘要) As 摘要, Decode(Nvl(附加标志, 0), 9, 1, 0), 保险编码, 费用类型, n_组id
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And
          Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));

  --原始记录
  If n_记帐 = 1 And Nvl(n_已结帐, 0) = 0 Then
    Update 门诊费用记录
    Set 记录状态 = 3, 结帐id = n_销帐id, 结帐金额 = 实收金额
    Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And
          Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
  Else
    Update 门诊费用记录
    Set 记录状态 = 3
    Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And
          Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
  End If;

  n_结帐id := 0;
  If n_记帐 = 0 Then
    --获取结帐ID
    Select Nvl(结帐id, 0)
    Into n_结帐id
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 = 3 And NO = 单据号_In And
          Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0)) And
          Rownum = 1;
  End If;

  If n_记帐 = 1 Then
    --记帐
    For c_费用 In (Select 实收金额, 病人科室id, 开单部门id, 执行部门id, 收入项目id
                 From 门诊费用记录
                 Where 记录性质 = 4 And 记录状态 = 3 And NO = 单据号_In And
                       Nvl(附加标志, 0) =
                       Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0)) And
                       Nvl(记帐费用, 0) = 1) Loop
      --病人余额
      Update 病人余额
      Set 费用余额 = Nvl(费用余额, 0) - Nvl(c_费用.实收金额, 0)
      Where 病人id = Nvl(n_病人id, 0) And 性质 = 1 And 类型 = 1
      Returning 费用余额 Into n_返回额;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 性质, 类型, 费用余额, 预交余额)
        Values
          (n_病人id, 1, 1, -1 * Nvl(c_费用.实收金额, 0), 0);
        n_返回额 := Nvl(c_费用.实收金额, 0);
      End If;
      If Nvl(n_返回额, 0) = 0 Then
        Delete 病人余额
        Where 病人id = Nvl(n_病人id, 0) And 性质 = 1 And 类型 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
      End If;
      --病人未结费用
      Update 病人未结费用
      Set 金额 = Nvl(金额, 0) - Nvl(c_费用.实收金额, 0)
      Where 病人id = n_病人id And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(病人科室id, 0) = Nvl(c_费用.病人科室id, 0) And
            Nvl(开单部门id, 0) = Nvl(c_费用.开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(c_费用.执行部门id, 0) And 收入项目id + 0 = c_费用.收入项目id And
            来源途径 + 0 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人未结费用
          (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
        Values
          (n_病人id, Null, Null, c_费用.病人科室id, c_费用.开单部门id, c_费用.执行部门id, c_费用.收入项目id, 1, -1 * Nvl(c_费用.实收金额, 0));
      End If;
    End Loop;
    Delete 病人未结费用
    Where 病人id = n_病人id And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(金额, 0) = 0 And 来源途径 + 0 = 1;
  End If;

  If n_记帐 = 0 Then
    --1.退费
    --病人挂号结算:现金和个人帐户部份
    If 结算方式_In Is Null Then
      If 非原样退结算_In Is Not Null Then
        --退款金额获取
        If Nvl(退费类型_In, 0) <> 0 Or Nvl(n_二次退费, 0) <> 0 Then
          --如果是单独退病历费,或者只退挂号费,先获取退费金额
          Begin
            --获取本次退款金额
            Select Sum(Nvl(实收金额, 0)) As 收款金额
            Into n_退款金额
            From 门诊费用记录
            Where NO = 单据号_In And 记录性质 = 4 And 记录状态 = 3 And
                  Nvl(附加标志, 0) =
                  Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
          
          Exception
            When Others Then
              v_Err_Msg := '单据【' || 单据号_In || '】的' || Case Nvl(退费类型_In, 0)
                             When 1 Then
                              '挂号费用'
                             When 2 Then
                              '病历费'
                           End || '可能由于并发原因已经进行了退费或者单据不存在!';
              Raise Err_Item;
          End;
          Begin
            Select 冲预交
            Into n_退费金额
            From 病人预交记录
            Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id And Instr(',' || 非原样退结算_In || ',', ',' || 结算方式 || ',') = 0;
          Exception
            When Others Then
              n_退费金额 := 0;
          End;
        
          --a.允许的结算方式
        
          Insert Into 病人预交记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
             结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
            Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In, -n_退款金额,
                   n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
            From 病人预交记录
            Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id And Instr(',' || 非原样退结算_In || ',', ',' || 结算方式 || ',') = 0;
        
          If n_退费金额 = 0 Then
            --b.不允许的退现金
            If n_退款金额 <> 0 Then
              If v_退指定结算方式 Is Null Then
                --退给现金
                Begin
                  Select 名称 Into v_退指定结算方式 From 结算方式 Where 性质 = 1;
                Exception
                  When Others Then
                    v_退指定结算方式 := '现金';
                End;
              End If;
              Update 病人预交记录
              Set 冲预交 = 冲预交 + (-1 * n_退款金额)
              Where 记录性质 = 4 And 记录状态 = 2 And 结帐id = n_销帐id And 结算方式 = v_退指定结算方式;
              If Sql%RowCount = 0 Then
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别,
                   卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
                  Select 病人预交记录_Id.Nextval, a.No, a.实际票号, a.记录性质, 2, a.病人id, a.主页id, a.科室id, a.摘要, v_退指定结算方式, d_Date,
                         操作员编号_In, 操作员姓名_In, -1 * n_退款金额, n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
                  From 病人预交记录 A
                  Where a.记录性质 = 4 And a.记录状态 = 1 And a.结帐id = n_结帐id And Rownum < 2;
              End If;
            End If;
          End If;
        Else
          --a.允许的结算方式原样退
          Insert Into 病人预交记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
             结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
            Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In, -冲预交,
                   n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
            From 病人预交记录
            Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id And Instr(',' || 非原样退结算_In || ',', ',' || 结算方式 || ',') = 0;
        
          --b.不允许的退现金
          Begin
            Select Sum(冲预交)
            Into n_退费金额
            From 病人预交记录
            Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id And Instr(',' || 非原样退结算_In || ',', ',' || 结算方式 || ',') > 0;
          Exception
            When Others Then
              n_退费金额 := 0;
          End;
          If n_退费金额 <> 0 Then
            If v_退指定结算方式 Is Null Then
              --退给现金
              Begin
                Select 结算方式
                Into v_退指定结算方式
                From 病人预交记录
                Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id And
                      Instr(',' || 非原样退结算_In || ',', ',' || 结算方式 || ',') = 0;
              
              Exception
                When Others Then
                  Begin
                    Select 名称 Into v_退指定结算方式 From 结算方式 Where 性质 = 1;
                  Exception
                    When Others Then
                      v_退指定结算方式 := '现金';
                  End;
              End;
            End If;
            Update 病人预交记录
            Set 冲预交 = 冲预交 + (-1 * n_退费金额)
            Where 记录性质 = 4 And 记录状态 = 2 And 结帐id = n_销帐id And 结算方式 = v_退指定结算方式;
            If Sql%RowCount = 0 Then
              Insert Into 病人预交记录
                (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别,
                 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
                Select 病人预交记录_Id.Nextval, a.No, a.实际票号, a.记录性质, 2, a.病人id, a.主页id, a.科室id, a.摘要, v_退指定结算方式, d_Date,
                       操作员编号_In, 操作员姓名_In, -1 * n_退费金额, n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
                From 病人预交记录 A
                Where a.记录性质 = 4 And a.记录状态 = 1 And a.结帐id = n_结帐id And Rownum < 2;
            End If;
          End If;
        End If;
      Else
        --退款金额获取
        If Nvl(退费类型_In, 0) <> 0 Or Nvl(n_二次退费, 0) <> 0 Then
          --如果是单独退病历费,或者只退挂号费,先获取退费金额
          Begin
            --获取本次退款金额
            Select Sum(Nvl(实收金额, 0)) As 收款金额
            Into n_退款金额
            From 门诊费用记录
            Where NO = 单据号_In And 记录性质 = 4 And 记录状态 = 3 And
                  Nvl(附加标志, 0) =
                  Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
          Exception
            When Others Then
              v_Err_Msg := '单据【' || 单据号_In || '】的' || Case Nvl(退费类型_In, 0)
                             When 1 Then
                              '挂号费用'
                             When 2 Then
                              '病历费'
                           End || '可能由于并发原因已经进行了退费或者单据不存在!';
              Raise Err_Item;
          End;
        End If;
        If Nvl(n_二次退费, 0) = 0 And Nvl(退费类型_In, 0) = 0 Then
          --首次全退
          Insert Into 病人预交记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
             结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
            Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In,
                   -1 * 冲预交, n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
            From 病人预交记录
            Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id;
        Else
          --二次退费,或者本次单退一部分
          --二次退费时,记录状态=3 ,首次部分退,记录状态为1
          Insert Into 病人预交记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
             结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
            Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In,
                   -1 * n_退款金额, n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
            From 病人预交记录
            Where 记录性质 = 4 And 记录状态 = Decode(Nvl(n_二次退费, 0), 0, 1, 3) And 结帐id = n_结帐id And 摘要 = '医保挂号' And
                  冲预交 = n_退款金额 And Rownum < 2;
          If Sql%RowCount = 0 Then
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
               结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
              Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In,
                     -1 * n_退款金额, n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
              From 病人预交记录
              Where 记录性质 = 4 And 记录状态 = Decode(Nvl(n_二次退费, 0), 0, 1, 3) And 结帐id = n_结帐id And 冲预交 = n_退款金额 And
                    Rownum < 2;
          End If;
          If Sql%RowCount = 0 Then
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
               结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
              Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In,
                     -1 * n_退款金额, n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
              From 病人预交记录
              Where 记录性质 = 4 And 记录状态 = Decode(Nvl(n_二次退费, 0), 0, 1, 3) And 结帐id = n_结帐id And Rownum < 2;
            If Sql%RowCount = 0 Then
              --部分退费,并且全部使用预交款缴费时才存在此种情况
              n_预交金额 := n_退款金额;
            End If;
          End If;
        
        End If;
      End If;
    Else
      --按结算方式退
      v_结算内容 := 结算方式_In || '|'; --以空格分开以|结尾,没有结算号码的
      While v_结算内容 Is Not Null Loop
        v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '|') - 1);
        v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
      
        v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1));
      
        v_当前结算   := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        n_三方卡标志 := To_Number(v_当前结算);
      
        If n_三方卡标志 = 0 Then
          Insert Into 病人预交记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
             结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质, 结算号码)
            Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, v_结算方式, d_Date, 操作员编号_In, 操作员姓名_In,
                   -1 * n_结算金额, n_销帐id, n_组id, 预交类别, Null, Null, Null, Null, Null, 合作单位, 4, 结算号码
            From 病人预交记录
            Where 记录性质 = 4 And 记录状态 = Decode(Nvl(n_二次退费, 0), 0, 1, 3) And 结帐id = n_结帐id And Rownum < 2;
        Else
          Insert Into 病人预交记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
             结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质, 结算号码)
            Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, v_结算方式, d_Date, 操作员编号_In, 操作员姓名_In,
                   -1 * n_结算金额, n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4, 结算号码
            From 病人预交记录
            Where 记录性质 = 4 And 记录状态 = Decode(Nvl(n_二次退费, 0), 0, 1, 3) And 结帐id = n_结帐id And
                  (卡类别id Is Not Null Or 结算卡序号 Is Not Null) And Rownum < 2;
        End If;
        v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '|') + 1);
      End Loop;
    End If;
    --首次退费时,记录状态便调整为了3
    Update 病人预交记录
    Set 记录状态 = 3
    Where 记录性质 = 4 And 记录状态 = Decode(Nvl(n_二次退费, 0), 0, 1, 3) And 结帐id = n_结帐id;
  
    --冲预交 1-全退 2-部分退,部分退时当全部使用预交进行缴款
    If Nvl(退费类型_In, 0) = 0 Or (Nvl(退费类型_In, 0) <> 0 And n_预交金额 <> 0) Then
      --病人挂号结算:冲预交款部份
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
         结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
        Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, d_Date,
               操作员姓名_In, 操作员编号_In, -1 * Decode(Nvl(退费类型_In, 0) + Nvl(n_二次退费, 0), 0, 冲预交, n_预交金额), n_销帐id, n_组id, 预交类别,
               卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
        From 病人预交记录
        Where 记录性质 In (1, 11) And 结帐id = n_结帐id And Nvl(冲预交, 0) <> 0 And
              Rownum = Decode(Nvl(退费类型_In, 0) + Nvl(n_二次退费, 0), 0, Rownum, 1);
    End If;
  
    --处理病人预交余额
    For c_预交 In (Select 病人id, 预交类别, -1 * Sum(Nvl(冲预交, 0)) As 冲预交
                 From 病人预交记录
                 Where 记录性质 In (1, 11) And 结帐id = n_销帐id
                 Group By 病人id, 预交类别) Loop
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) + Nvl(c_预交.冲预交, 0)
      Where 病人id = c_预交.病人id And 类型 = Nvl(c_预交.预交类别, 2) And 性质 = 1
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 预交余额, 性质, 类型)
        Values
          (c_预交.病人id, Nvl(c_预交.冲预交, 0), 1, Nvl(c_预交.预交类别, 2));
        n_返回值 := Nvl(c_预交.冲预交, 0);
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 病人余额
        Where 病人id = c_预交.病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
      End If;
    End Loop;
  
    If Nvl(退费类型_In, 0) <> 2 Then
      --光退挂号费,不回收票据
      --退卡收回票据(可能上次挂号使用票据,不能收回)
      Begin
        --从最后一次的打印内容中取
        Select ID
        Into n_打印id
        From (Select b.Id
               From 票据使用明细 A, 票据打印内容 B
               Where a.打印id = b.Id And a.性质 = 1 And b.数据性质 = 4 And b.No = 单据号_In
               Order By a.使用时间 Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          Null;
      End;
    
      If n_打印id Is Not Null Then
        Insert Into 票据使用明细
          (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
          Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_Date, 操作员姓名_In
          From 票据使用明细
          Where 打印id = n_打印id And 性质 = 1;
      End If;
    End If;
  End If;

  --单独退病历费用,不处理汇总记录
  --相关汇总表的处理

  --病人挂号汇总
  Open c_Registinfo(3);
  Fetch c_Registinfo
    Into r_Registrow;

  If c_Registinfo%RowCount = 0 Then
    --只收病历费时无号别,不处理
    Close c_Registinfo;
  Else
  
    --需要确定是否预约挂号
    --1.如果是退预约挂号产生的挂号记录,则需要减已挂数和其中已接数
    --2.如果是正常挂号,则只减已挂数
  
    Begin
      Select Decode(预约, Null, 0, 0, 0, 1) Into n_预约挂号 From 病人挂号记录 Where NO = 单据号_In And Rownum = 1;
    Exception
      When Others Then
        n_预约挂号 := 0;
    End;
  
    Update 临床出诊记录
    Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - n_预约挂号, 已约数 = Nvl(已约数, 0) - n_预约挂号
    Where ID = n_出诊记录id;
  
    Update 病人挂号汇总
    Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - n_预约挂号, 已约数 = Nvl(已约数, 0) - n_预约挂号
    Where 日期 = Trunc(r_Registrow.发生时间) And Nvl(医生id, 0) = Nvl(r_Registrow.医生id, 0) And
          Nvl(医生姓名, '-') = Nvl(r_Registrow.医生姓名, '-') And 科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And
          (号码 = r_Registrow.号码 Or 号码 Is Null);
  
    If Sql%RowCount = 0 Then
      Insert Into 病人挂号汇总
        (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 其中已接收, 已约数)
      Values
        (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
         Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, -1, -1 * n_预约挂号, -1 * n_预约挂号);
    End If;
  
    Close c_Registinfo;
  End If;

  If n_记帐 = 0 Then
    --人员缴款余额(包括个人帐户等的结算金额,不含退冲预交款)
    For r_Opermoney In c_Opermoney Loop
      Update 人员缴款余额
      Set 余额 = Nvl(余额, 0) + (-1 * r_Opermoney.冲预交)
      Where 收款员 = 操作员姓名_In And 结算方式 = r_Opermoney.结算方式 And 性质 = 1
      Returning 余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 人员缴款余额
          (收款员, 结算方式, 性质, 余额)
        Values
          (操作员姓名_In, r_Opermoney.结算方式, 1, -1 * r_Opermoney.冲预交);
        n_返回值 := r_Opermoney.冲预交;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 人员缴款余额
        Where 收款员 = 操作员姓名_In And 结算方式 = r_Opermoney.结算方式 And 性质 = 1 And Nvl(余额, 0) = 0;
      End If;
    End Loop;
  End If;
  If Nvl(退费类型_In, 0) <> 2 Then
    n_挂号生成队列 := Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113));
    If n_挂号生成队列 <> 0 Then
      n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113));
      If Nvl(n_分诊台签到排队, 0) = 0 Then
      
        --要删除队列
        For v_挂号 In (Select ID, 诊室, 执行部门id, 执行人 From 病人挂号记录 Where NO = 单据号_In) Loop
          Zl_排队叫号队列_Delete(v_挂号.执行部门id, v_挂号.Id);
        End Loop;
      End If;
    End If;
  
    --医保产生的就诊登记记录
    Begin
      Select 病人id, 发生时间 Into n_就诊病人id, d_就诊时间 From 病人挂号记录 Where NO = 单据号_In;
      Delete From 就诊登记记录 Where 病人id = n_就诊病人id And 就诊时间 = d_就诊时间 And 主页id Is Null;
    Exception
      When Others Then
        Null;
    End;
  
    --病人挂号记录
    Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
  
    Update 病人挂号记录 Set 记录状态 = 3 Where NO = 单据号_In And 记录性质 = 1 And 记录状态 = 1;
    If Sql%NotFound Then
      v_Err_Msg := '挂号单【' || 单据号_In || '】不存在或由于并发原因已经被退号';
      Raise Err_Item;
    End If;
    Insert Into 病人挂号记录
      (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 操作员编号, 操作员姓名,
       复诊, 号序, 社区, 预约, 摘要, 预约方式, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 险类, 医疗付款方式, 出诊记录id)
      Select n_挂号id, NO, 记录性质, 2, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, d_Date, 发生时间,
             操作员编号_In, 操作员姓名_In, 复诊, 号序, 社区, 预约, Nvl(摘要_In, 摘要) As 摘要, 预约方式, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 险类, 医疗付款方式,
             n_出诊记录id
      From 病人挂号记录
      Where NO = 单据号_In;
  End If;
  --消息推送
  Begin
    Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
      Using 2, 单据号_In;
  Exception
    When Others Then
      Null;
  End;
  b_Message.Zlhis_Regist_003(n_挂号id, 单据号_In);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人挂号记录_出诊_Delete;
/

Create Or Replace Procedure Zl_病人挂号记录_出诊_Insert
(
  出诊记录id_In    临床出诊记录.Id%Type,
  病人id_In        门诊费用记录.病人id%Type,
  门诊号_In        门诊费用记录.标识号%Type,
  姓名_In          门诊费用记录.姓名%Type,
  性别_In          门诊费用记录.性别%Type,
  年龄_In          门诊费用记录.年龄%Type,
  付款方式_In      门诊费用记录.付款方式%Type, --用于存放病人的医疗付款方式编号
  费别_In          门诊费用记录.费别%Type,
  单据号_In        门诊费用记录.No%Type,
  票据号_In        门诊费用记录.实际票号%Type,
  序号_In          门诊费用记录.序号%Type,
  价格父号_In      门诊费用记录.价格父号%Type,
  从属父号_In      门诊费用记录.从属父号%Type,
  收费类别_In      门诊费用记录.收费类别%Type,
  收费细目id_In    门诊费用记录.收费细目id%Type,
  数次_In          门诊费用记录.数次%Type,
  标准单价_In      门诊费用记录.标准单价%Type,
  收入项目id_In    门诊费用记录.收入项目id%Type,
  收据费目_In      门诊费用记录.收据费目%Type,
  结算方式_In      Varchar2,
  应收金额_In      门诊费用记录.应收金额%Type,
  实收金额_In      门诊费用记录.实收金额%Type,
  病人科室id_In    门诊费用记录.病人科室id%Type,
  开单部门id_In    门诊费用记录.开单部门id%Type,
  执行部门id_In    门诊费用记录.执行部门id%Type,
  操作员编号_In    门诊费用记录.操作员编号%Type,
  操作员姓名_In    门诊费用记录.操作员姓名%Type,
  发生时间_In      门诊费用记录.发生时间%Type,
  登记时间_In      门诊费用记录.登记时间%Type,
  医生姓名_In      挂号安排.医生姓名%Type,
  医生id_In        挂号安排.医生id%Type,
  病历费_In        Number, --该条记录是否病历工本费
  急诊_In          Number,
  号别_In          挂号安排.号码%Type,
  诊室_In          门诊费用记录.发药窗口%Type,
  结帐id_In        门诊费用记录.结帐id%Type,
  领用id_In        票据使用明细.领用id%Type,
  预交支付_In      病人预交记录.冲预交%Type, --刷卡挂号时使用的预交金额,序号为1传入.
  现金支付_In      病人预交记录.冲预交%Type, --挂号时现金支付部份金额,序号为1传入.
  个帐支付_In      病人预交记录.冲预交%Type, --挂号时个人帐户支付金额,,序号为1传入.
  保险大类id_In    门诊费用记录.保险大类id%Type,
  保险项目否_In    门诊费用记录.保险项目否%Type,
  统筹金额_In      门诊费用记录.统筹金额%Type,
  摘要_In          门诊费用记录.摘要%Type, --预约挂号摘要信息
  预约挂号_In      Number := 0, --预约挂号时用(记录状态=0,发生时间为预约时间),此时不需要传入结算相关参数
  收费票据_In      Number := 0, --挂号是否使用收费票据
  保险编码_In      门诊费用记录.保险编码%Type,
  复诊_In          病人挂号记录.复诊%Type := 0,
  号序_In          挂号序号状态.序号%Type := Null, --预约时填入费用记录的发药窗口字段,挂号时填入挂号记录
  社区_In          病人挂号记录.社区%Type := Null,
  预约接收_In      Number := 0,
  预约方式_In      预约方式.名称%Type := Null,
  生成队列_In      Number := 0,
  卡类别id_In      病人预交记录.卡类别id%Type := Null,
  结算卡序号_In    病人预交记录.结算卡序号%Type := Null,
  卡号_In          病人预交记录.卡号%Type := Null,
  交易流水号_In    病人预交记录.交易流水号%Type := Null,
  交易说明_In      病人预交记录.交易说明%Type := Null,
  合作单位_In      病人预交记录.合作单位%Type := Null,
  操作类型_In      Number := 0,
  险类_In          病人挂号记录.险类%Type := Null,
  结算模式_In      Number := 0,
  记帐费用_In      Number := 0,
  退号重用_In      Number := 1,
  冲预交病人ids_In Varchar2 := Null,
  修正病人费别_In  Number := 0,
  预约顺序号_In    临床出诊序号控制.预约顺序号%Type := Null
) As
  ---------------------------------------------------------------------------
  --
  --参数:
  --     操作类型_in:0-正常挂号或者预约 1-操作员拥有加号权限加号
  --     修正病人费别_In:0-不修改病人费别 1-修改病人费别
  ----------------------------------------------------------------------------
  --该游标用于收费冲预交的可用预交列表
  --以ID排序，优先冲上次未冲完的。
  Cursor c_Deposit
  (
    v_病人id        病人信息.病人id%Type,
    v_冲预交病人ids Varchar2
  ) Is
    Select *
    From (Select a.Id, a.病人id, a.记录状态, a.预交类别, a.No, Nvl(a.金额, 0) As 金额
           From 病人预交记录 A,
                (Select NO, Sum(Nvl(a.金额, 0)) As 金额
                  From 病人预交记录 A
                  Where a.结帐id Is Null And Nvl(a.金额, 0) <> 0 And
                        a.病人id In (Select Column_Value From Table(f_Num2list(v_冲预交病人ids))) And Nvl(a.预交类别, 2) = 1
                  Group By NO
                  Having Sum(Nvl(a.金额, 0)) <> 0) B
           Where a.结帐id Is Null And Nvl(a.金额, 0) <> 0 And a.结算方式 Not In (Select 名称 From 结算方式 Where 性质 = 5) And
                 a.No = b.No And a.病人id In (Select Column_Value From Table(f_Num2list(v_冲预交病人ids))) And
                 Nvl(a.预交类别, 2) = 1
           Union All
           Select 0 As ID, Max(病人id) As 病人id, 记录状态, 预交类别, NO, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额
           From 病人预交记录
           Where 记录性质 In (1, 11) And 结帐id Is Not Null And Nvl(金额, 0) <> Nvl(冲预交, 0) And
                 病人id In (Select Column_Value From Table(f_Num2list(v_冲预交病人ids))) And Nvl(预交类别, 2) = 1 Having
            Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
           Group By 记录状态, NO, 预交类别)
    Order By Decode(病人id, Nvl(v_病人id, 0), 0, 1), ID, NO;
  --功能：新增一行病人挂号费用，并将结算汇总到病人预交记录
  --       同时汇总相关的汇总表(病人挂号汇总、费用汇总)
  --       第一行费用处理票据使用情况(领用ID_IN>0)

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_排队号码 排队叫号队列.排队号码%Type;
  v_现金     结算方式.名称%Type;
  v_个人帐户 结算方式.名称%Type;
  v_队列名称 排队叫号队列.队列名称%Type;

  n_分时段       Number;
  n_时段限号     Number;
  n_时段限约     Number;
  d_时段时间     Date;
  d_最大序号时间 Date;
  n_追加号       Number := 0; --处理时段过期 追加挂号的情况
  n_已约数       病人挂号汇总.已约数%Type;
  n_已接收       病人挂号汇总.其中已接收%Type;
  n_预约有效时间 Number;
  n_失效数       Number;
  n_失约挂号     Number := 0;
  n_已用数量     Number;
  n_锁定         Number := 0;

  n_打印id        票据打印内容.Id%Type;
  n_费用id        门诊费用记录.Id%Type;
  n_预交金额      病人预交记录.金额%Type;
  n_当前金额      病人预交记录.金额%Type;
  n_返回值        病人预交记录.金额%Type;
  n_预交id        病人预交记录.Id%Type;
  n_消费卡id      消费卡目录.Id%Type;
  n_挂号id        病人挂号记录.Id%Type;
  v_冲预交病人ids Varchar2(4000);

  n_组id           财务缴款分组.Id%Type;
  n_门诊号         病人信息.门诊号%Type;
  n_序号           挂号序号状态.序号%Type;
  n_已用序号       挂号序号状态.序号%Type;
  n_序号控制       挂号安排.序号控制%Type;
  n_分诊台签到排队 Number;
  n_Count          Number;
  n_限号数         Number(18);
  n_自制卡         Number;
  d_排队时间       Date;
  v_结算方式记录   Varchar2(1000);
  d_序号时间       Date;
  v_费别           费别.名称%Type;
  n_时段序号       Number := -1;
  n_预约生成队列   Number;
  v_结算方式       结算方式.名称%Type;
  v_结算内容       Varchar2(1000);
  v_当前结算       Varchar2(200);
  v_结算号码       病人预交记录.结算号码%Type;
  n_结算金额       病人预交记录.冲预交%Type;
  n_三方卡标志     Number(2);
  n_安排id         挂号安排.Id%Type;
  n_预约顺序号     临床出诊序号控制.预约顺序号%Type;
  n_计划id         挂号安排计划.Id%Type := 0;
  v_星期           挂号安排限制.限制项目%Type;
  n_限约数         Number(18);
  n_已挂数         Number(4) := 0;

  n_挂出的最大序号 Number(4) := 0;
  n_结算模式       病人信息.结算模式%Type;
  v_排队序号       排队叫号队列.排队序号%Type;
  v_机器名         挂号序号状态.机器名%Type;
  v_序号操作员     挂号序号状态.操作员姓名%Type;
  v_序号机器名     挂号序号状态.机器名%Type;
  v_付款方式       病人挂号记录.医疗付款方式%Type;
  n_状态           临床出诊序号控制.挂号状态%Type;
Begin
  --获取当前机器名称
  Select Terminal Into v_机器名 From V$session Where Audsid = Userenv('sessionid');
  n_组id          := Zl_Get组id(操作员姓名_In);
  v_冲预交病人ids := Nvl(冲预交病人ids_In, 病人id_In);

  If 费别_In Is Null Then
    Begin
      Select 名称 Into v_费别 From 费别 Where 缺省标志 = 1 And Rownum < 2;
      Update 病人信息 Set 费别 = v_费别 Where 病人id = 病人id_In;
    Exception
      When Others Then
        v_Err_Msg := '无法确定病人费别，请检查缺省费别是否正确设置！';
        Raise Err_Item;
    End;
  Else
    v_费别 := 费别_In;
    If Nvl(修正病人费别_In, 0) = 1 Then
      Begin
        Update 病人信息 Set 费别 = v_费别 Where 病人id = 病人id_In;
      Exception
        When Others Then
          v_Err_Msg := '没有找到对应的病人！';
          Raise Err_Item;
      End;
    End If;
  End If;

  If 门诊号_In Is Not Null Then
    Begin
      Select Nvl(门诊号, 0) Into n_门诊号 From 病人信息 Where 病人id = 病人id_In;
    Exception
      When Others Then
        n_门诊号 := 0;
    End;
    If n_门诊号 = 0 Then
      Update 病人信息 Set 门诊号 = 门诊号_In Where 病人id = 病人id_In;
    End If;
  End If;

  Begin
    Update 临床出诊序号控制
    Set 挂号状态 = 0
    Where 记录id = 出诊记录id_In And 序号 = 号序_In And Nvl(挂号状态, 0) = 3 And 操作员姓名 = 操作员姓名_In;
  Exception
    When Others Then
      Null;
  End;

  If Nvl(预约挂号_In, 0) = 0 Then
    --挂号或者预约接收
    --因为现在有按日编单据号规则,日挂号量不能超过10000次,所以要检查唯一约束。
    Select Count(*)
    Into n_Count
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 In (1, 3) And 序号 = 序号_In And NO = 单据号_In;
    If n_Count <> 0 Then
      v_Err_Msg := '挂号单据号重复,不能保存！' || Chr(13) || '如果使用了按日顺序编号,当日挂号量不能超过10000人次。';
      Raise Err_Item;
    End If;
  
    --获取结算方式名称
    Begin
      Select 名称 Into v_现金 From 结算方式 Where 性质 = 1;
    Exception
      When Others Then
        v_现金 := '现金';
    End;
    Begin
      Select 名称 Into v_个人帐户 From 结算方式 Where 性质 = 3;
    Exception
      When Others Then
        v_个人帐户 := '个人帐户';
    End;
  End If;

  n_序号 := 号序_In;

  --获取是否分时段
  Begin
    Select Nvl(是否分时段, 0), Nvl(是否序号控制, 0), 限号数, 限约数
    Into n_分时段, n_序号控制, n_限号数, n_限约数
    From 临床出诊记录
    Where ID = 出诊记录id_In;
  Exception
    When Others Then
      n_分时段   := 0;
      n_序号控制 := 0;
      n_限号数   := Null;
      n_限约数   := Null;
  End;

  If n_序号 Is Null And n_分时段 = 1 And n_序号控制 = 0 Then
    Begin
      Select 序号 Into n_序号 From 临床出诊序号控制 Where 记录id = 出诊记录id_In And 开始时间 = 发生时间_In;
    Exception
      When Others Then
        n_序号 := Null;
    End;
  End If;

  --分时段挂号时判断是否是过期挂号 也就是追加号的情况 目前只针对专家号分时段进行处理
  If Nvl(预约挂号_In, 0) = 0 And n_分时段 > 0 And Nvl(n_序号控制, 0) = 1 And 号序_In Is Null And Nvl(操作类型_In, 0) = 0 Then
    Begin
      Select Max(To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'))
      Into d_最大序号时间
      From 临床出诊序号控制
      Where 记录id = 出诊记录id_In And Nvl(数量, 0) <> 0;
    
      n_追加号 := Case Sign(发生时间_In - d_最大序号时间)
                 When -1 Then
                  0
                 Else
                  1
               End;
    Exception
      When Others Then
        n_追加号 := 0;
    End;
  End If;
  d_时段时间 := 发生时间_In;

  If 序号_In = 1 And n_分时段 > 0 Then
    If Nvl(n_序号控制, 0) = 1 Then
      Begin
        Select Nvl(序号, 0),
               To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
               数量, Decode(Nvl(是否预约, 0), 0, 0, 数量)
        Into n_时段序号, d_时段时间, n_时段限号, n_时段限约
        From 临床出诊序号控制
        Where 记录id = 出诊记录id_In And 序号 = n_序号;
      Exception
        When Others Then
          n_时段序号 := -1;
          n_分时段   := 0;
          d_时段时间 := 发生时间_In;
          n_时段限号 := 0;
          n_时段限约 := 0;
      End;
    Else
      --挂号时检查 是否分了时段,分了时段,把时段的限制条件给取出来
      Begin
        Select Nvl(序号, 0),
               To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
               数量, Decode(Nvl(是否预约, 0), 0, 0, 数量)
        Into n_时段序号, d_时段时间, n_时段限号, n_时段限约
        From 临床出诊序号控制
        Where 记录id = 出诊记录id_In And 序号 = n_序号 And 预约顺序号 Is Null;
      Exception
        When Others Then
          n_时段序号 := -1;
          n_分时段   := 0;
          d_时段时间 := 发生时间_In;
          n_时段限号 := 0;
          n_时段限约 := 0;
      End;
    End If;
  End If;

  If 序号_In = 1 Then
    --获取当前未使用的序号
    If Nvl(n_序号控制, 0) = 1 And n_分时段 = 0 Then
      --<序号控制 未设置时段 获取可用的最大序号,以及已经使用的数量>
      Begin
        --最大序号
        Select Count(1) Into n_已用数量 From 病人挂号记录 Where 出诊记录id = 出诊记录id_In And 记录状态 = 1;
        Select Max(序号) Into n_已用序号 From 临床出诊序号控制 Where 记录id = 出诊记录id_In;
      Exception
        When Others Then
          n_已用序号 := 0;
          n_已用数量 := 0;
      End;
      Begin
        --最大序号
        Select Sum(Nvl(数量, 0))
        
        Into n_已约数
        From 临床出诊序号控制
        Where 记录id = 出诊记录id_In And Nvl(挂号状态, 0) = 2;
      Exception
        When Others Then
          n_已约数 := 0;
      End;
    
      If n_序号 Is Null Then
        n_序号 := Nvl(n_已用序号, 0) + 1;
      End If;
      --<序号控制 未设置时段 获取可用的最大序号 以及已经使用的数量 --end>
    
      --非加号的情况需要检查是否超过了限制
      If 操作类型_In = 0 Then
        If Nvl(预约挂号_In, 0) = 0 Then
          --挂号时检查
          --启用序号控制未分时段 达到了限制
          If n_限号数 <= n_已用数量 And n_限号数 > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(Trunc(发生时间_In), 'yyyy-mm-dd ') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        
        Else
          --预约时检查
          If n_限约数 = 0 Then
            n_限约数 := n_限号数;
          End If;
        
          If n_限约数 <= n_已约数 And n_限约数 > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(Trunc(发生时间_In), 'yyyy-mm-dd') || '已达到最大限约数！';
            Raise Err_Item;
          End If;
        End If;
      
      Else
        Null;
        --启用序号控制,未分时段 加号情况   不处理,如果以后有限制条件以后补充
      End If;
    
    Elsif Nvl(n_分时段, 0) > 0 And Nvl(n_序号控制, 0) = 0 Then
      --<--普通号分时段 这里只有预约一种情况-->
      If 操作类型_In = 0 Then
        --<正常预约挂号-->
        Begin
          Select Count(0) As 已挂数, Nvl(Sum(Decode(Nvl(Sign(a.开始时间 - d_时段时间), 0), 0, 1, 0)), 0) As 已约数
          Into n_已挂数, n_已约数
          From 临床出诊序号控制 A
          Where 记录id = 出诊记录id_In And 挂号状态 Not In (0, 4, 5);
        Exception
          When Others Then
            n_已挂数 := 0;
            n_已约数 := 0;
        End;
      
        n_时段限约 := n_时段限号; --普通号分时段的情况,29中n_时段限约始终是0 这里特殊处理
        --检查限制数量
        If n_限约数 = 0 Then
          n_限约数 := n_限号数;
        End If;
        If n_时段限约 <= n_已约数 Or n_限约数 <= n_已约数 Then
          v_Err_Msg := '号别' || 号别_In || '在时段' || To_Char(d_时段时间, 'hh24:mi') || '已达到最大限约数！';
          Raise Err_Item;
        End If;
        If n_限号数 <= n_已挂数 Then
          v_Err_Msg := '号别' || 号别_In || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
          Raise Err_Item;
        End If;
      End If;
    
      --没有达到时段的限号数 号码在当前时段往后追加
    
      --获取当天挂出的最大号序
      Select Nvl(Max(序号), 0)
      Into n_挂出的最大序号
      From 临床出诊序号控制 A
      Where 记录id = 出诊记录id_In And 预约顺序号 Is Null And 挂号状态 Not In (0, 5);
      If 预约顺序号_In Is Not Null Then
        n_预约顺序号 := 预约顺序号_In;
      Else
        Begin
          Select Nvl(Max(预约顺序号), 0) + 1
          Into n_预约顺序号
          From 临床出诊序号控制
          Where 记录id = 出诊记录id_In And 序号 = n_时段序号 And 预约顺序号 Is Not Null;
        Exception
          When Others Then
            n_预约顺序号 := Null;
        End;
      End If;
      --设置序号
      n_序号 := RPad(Nvl(n_时段序号, 0), Length(n_限号数) + Length(Nvl(n_时段序号, 0)), 0) + n_预约顺序号;
      If n_预约顺序号 Is Null Then
        n_序号 := Nvl(n_挂出的最大序号, 0) + 1;
      End If;
    
      --<--普通号分时段--End>
    Elsif Nvl(n_分时段, 0) > 0 And Nvl(n_序号控制, 0) = 1 Then
      --<启用序号控制 设置时段
      --专家号分时段
      Begin
        Select Max(序号), Sum(Decode(序号, Nvl(号序_In, 0), 0, 1)), Sum(Decode(Sign(开始时间 - d_时段时间), 0, 1, 0))
        Into n_已用序号, n_已挂数, n_已用数量
        From 临床出诊序号控制
        Where 记录id = 出诊记录id_In And 挂号状态 Not In (0, 4, 5);
      Exception
        When Others Then
          n_已用序号 := 0;
          n_已用数量 := 0;
          n_已挂数   := 0;
      End;
    
      n_失效数 := 0;
      If Nvl(预约挂号_In, 0) = 0 Then
        n_预约有效时间 := Zl_To_Number(zl_GetSysParameter('预约有效时间', 1111));
        n_失约挂号     := Zl_To_Number(zl_GetSysParameter('失约用于挂号', 1111));
        If Nvl(n_预约有效时间, 0) <> 0 And Nvl(n_失约挂号, 0) > 0 Then
          Begin
            Select Sum(Decode(Sign((Sysdate - n_预约有效时间 / 24 / 60) - 开始时间), 1, 1, 0))
            Into n_失效数
            From 临床出诊序号控制
            Where 记录id = 出诊记录id_In And 开始时间 Between Trunc(Sysdate) And Sysdate And Nvl(挂号状态, 0) = 2;
          Exception
            When Others Then
              n_失效数 := 0;
          End;
        End If;
      End If;
    
      If 操作类型_In = 0 Then
        If Nvl(预约挂号_In, 0) = 0 Then
        
          --挂号 追加号码时不检查时段限号数
          If n_时段限号 <= n_已用数量 And Nvl(n_追加号, 0) = 0 Then
            v_Err_Msg := '号别' || 号别_In || '在时段' || To_Char(d_时段时间, 'hh24:mi') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
          If n_限号数 <= n_已挂数 - n_失效数 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        
        Else
          --挂号
          If n_限约数 = 0 Then
            n_限约数 := n_限号数;
          End If;
          If n_限约数 <= n_已用数量 Then
            v_Err_Msg := '号别' || 号别_In || '在时段' || To_Char(d_时段时间, 'hh24:mi') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
          If n_限号数 <= n_已挂数 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        End If;
      End If;
      If n_序号 Is Null Then
        --设置序号
        If Nvl(n_已用序号, 0) < Nvl(n_时段序号, 0) Then
          n_已用序号 := Nvl(n_时段序号, 0);
        End If;
        n_序号 := Nvl(n_已用序号, 0) + 1;
      End If;
    Elsif Nvl(n_分时段, 0) = 0 And Nvl(n_序号控制, 0) = 0 And Nvl(病历费_In, 0) = 0 And Nvl(号别_In, 0) > 0 Then
      ---<--普通号  -->
      Begin
        Select 已挂数, 已约数 Into n_已用数量, n_已约数 From 临床出诊记录 Where ID = 出诊记录id_In;
      Exception
        When Others Then
          n_已用数量 := 0;
          n_已约数   := 0;
      End;
      If Nvl(操作类型_In, 0) = 0 Then
        If Nvl(预约挂号_In, 0) = 0 Then
          --挂号
          If Nvl(n_已用数量, 0) >= Nvl(n_限号数, 0) And Nvl(n_限号数, 0) > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        Else
          --预约
          If (Nvl(n_限约数, 0) > 0 And Nvl(n_已约数, 0) >= Nvl(n_限约数, 0)) Or
             (Nvl(n_已用数量, 0) >= Nvl(n_限号数, 0) And Nvl(n_限号数, 0) > 0) Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        End If;
      End If;
      ---<--普通号  -->
    End If;
  End If;

  --更新挂号序号状态
  If 序号_In = 1 And Not n_序号 Is Null Then
    If n_分时段 = 1 Then
      d_序号时间 := 发生时间_In;
    Else
      d_序号时间 := Trunc(发生时间_In);
    End If;
    --锁定序号的处理
    Begin
      If n_预约顺序号 Is Null Then
        Select 操作员姓名, 工作站名称
        Into v_序号操作员, v_序号机器名
        From 临床出诊序号控制
        Where Nvl(挂号状态, 0) = 5 And 记录id = 出诊记录id_In And 序号 = n_序号;
      Else
        Select 操作员姓名, 工作站名称
        Into v_序号操作员, v_序号机器名
        From 临床出诊序号控制
        Where Nvl(挂号状态, 0) = 5 And 记录id = 出诊记录id_In And 序号 = n_时段序号 And 预约顺序号 = n_预约顺序号;
      End If;
      n_锁定 := 1;
    Exception
      When Others Then
        v_序号操作员 := Null;
        v_序号机器名 := Null;
        n_锁定       := 0;
    End;
    If n_锁定 = 0 Then
      If n_预约顺序号 Is Null Then
        Update 临床出诊序号控制
        Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1)
        Where 记录id = 出诊记录id_In And 序号 = n_序号 And Nvl(挂号状态, 0) = 3 And 操作员姓名 = 操作员姓名_In;
      Else
        Update 临床出诊序号控制
        Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1)
        Where 记录id = 出诊记录id_In And 序号 = n_序号 And 预约顺序号 = n_预约顺序号 And Nvl(挂号状态, 0) = 3 And 操作员姓名 = 操作员姓名_In;
      End If;
      If Sql%RowCount = 0 Then
        Begin
          If Nvl(n_分时段, 0) > 0 Then
            If Nvl(n_序号控制, 0) = 1 Then
              --分时段后专家号 失约的预约号允许挂号
              Update 临床出诊序号控制
              Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1), 操作员姓名 = 操作员姓名_In
              Where 记录id = 出诊记录id_In And 序号 = n_序号 And Nvl(挂号状态, 0) In (0, 2);
              If Sql%NotFound Then
                Begin
                  Select 挂号状态 Into n_状态 From 临床出诊序号控制 Where 记录id = 出诊记录id_In And 序号 = n_序号;
                  v_Err_Msg := '序号' || n_序号 || '已被使用,请重新选择一个序号.';
                  Raise Err_Item;
                Exception
                  When Others Then
                    Insert Into 临床出诊序号控制
                      (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 锁号时间, 类型, 名称, 操作员姓名, 备注)
                      Select 出诊记录id_In, n_序号, d_序号时间, d_序号时间, 1, Decode(预约挂号_In, 1, 1, 0), Decode(预约挂号_In, 1, 2, 1),
                             Null, Null, Null, 操作员姓名_In, '追加号'
                      From Dual;
                End;
              End If;
            Else
              If Nvl(预约接收_In, 0) = 1 Then
                Insert Into 临床出诊序号控制
                  (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 锁号时间, 类型, 名称, 操作员姓名, 备注, 预约顺序号)
                  Select 记录id, 序号, 开始时间, 终止时间, 1, 1, Decode(预约挂号_In, 1, 2, 1), Null, Null, Null, 操作员姓名_In, n_序号, n_预约顺序号
                  From 临床出诊序号控制
                  Where 记录id = 出诊记录id_In And 序号 = n_时段序号 And 预约顺序号 Is Null;
              End If;
            End If;
          Else
            If Nvl(n_序号控制, 0) = 1 Then
              Update 临床出诊序号控制
              Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1)
              Where 记录id = 出诊记录id_In And 序号 = n_序号 And Nvl(挂号状态, 0) = 0;
            
              If Sql%RowCount = 0 Then
                Insert Into 临床出诊序号控制
                  (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 锁号时间, 类型, 名称, 操作员姓名, 备注)
                  Select 出诊记录id_In, n_序号, 发生时间_In, 发生时间_In, 1, Decode(预约挂号_In, 1, 1, 0), Decode(预约挂号_In, 1, 2, 1), Null,
                         Null, Null, 操作员姓名_In, '追加号'
                  From Dual;
              End If;
            End If;
          End If;
        Exception
          When Others Then
            v_Err_Msg := '序号' || n_序号 || '已被使用,请重新选择一个序号.';
            Raise Err_Item;
        End;
      End If;
    Else
      If 操作员姓名_In <> v_序号操作员 Or v_机器名 <> v_序号机器名 Then
        v_Err_Msg := '序号' || n_序号 || '已被自助机' || v_机器名 || '锁定,请重新选择一个序号.';
        Raise Err_Item;
      Else
        If n_预约顺序号 Is Null Then
          Update 临床出诊序号控制
          Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1)
          Where 记录id = 出诊记录id_In And 序号 = n_序号 And Nvl(挂号状态, 0) = 5 And 操作员姓名 = 操作员姓名_In And 工作站名称 = v_机器名;
        Else
          Update 临床出诊序号控制
          Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1)
          Where 记录id = 出诊记录id_In And 序号 = n_时段序号 And 预约顺序号 = n_预约顺序号 And Nvl(挂号状态, 0) = 5 And 操作员姓名 = 操作员姓名_In And
                工作站名称 = v_机器名;
        End If;
      End If;
    End If;
  End If;

  --产生病人挂号费用(可能单独是或包括病历费用)
  Select 病人费用记录_Id.Nextval Into n_费用id From Dual; --应该通过程序得到

  Insert Into 门诊费用记录
    (ID, 记录性质, 记录状态, 序号, 价格父号, 从属父号, NO, 实际票号, 门诊标志, 加班标志, 附加标志, 发药窗口, 病人id, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 病人科室id, 收费类别,
     计算单位, 收费细目id, 收入项目id, 收据费目, 付数, 数次, 标准单价, 应收金额, 实收金额, 结帐金额, 结帐id, 记帐费用, 开单部门id, 开单人, 划价人, 执行部门id, 执行人, 操作员编号, 操作员姓名,
     发生时间, 登记时间, 保险大类id, 保险项目否, 保险编码, 统筹金额, 摘要, 结论, 缴款组id)
  Values
    (n_费用id, 4, Decode(预约挂号_In, 1, 0, 1), 序号_In, Decode(价格父号_In, 0, Null, 价格父号_In), 从属父号_In, 单据号_In, 票据号_In, 1, 急诊_In,
     病历费_In, Decode(预约挂号_In, 1, To_Char(n_序号), 诊室_In), Decode(病人id_In, 0, Null, 病人id_In),
     Decode(门诊号_In, 0, Null, 门诊号_In), 付款方式_In, 姓名_In, Decode(姓名_In, Null, Null, 性别_In), Decode(姓名_In, Null, Null, 年龄_In),
     v_费别, 病人科室id_In, 收费类别_In, 号别_In, 收费细目id_In, 收入项目id_In, 收据费目_In, 1, 数次_In, 标准单价_In, 应收金额_In, 实收金额_In,
     Decode(预约挂号_In, 1, Null, Decode(Nvl(记帐费用_In, 0), 1, Null, 实收金额_In)),
     Decode(预约挂号_In, 1, Null, Decode(Nvl(记帐费用_In, 0), 1, Null, 结帐id_In)), Decode(Nvl(记帐费用_In, 0), 1, 1, 0), 开单部门id_In,
     操作员姓名_In, Decode(预约挂号_In, 1, 操作员姓名_In, Null), 执行部门id_In, 医生姓名_In, 操作员编号_In, 操作员姓名_In, 发生时间_In, 登记时间_In, 保险大类id_In,
     保险项目否_In, 保险编码_In, 统筹金额_In, 摘要_In, 预约方式_In, Decode(预约挂号_In, 1, Null, n_组id));

  --汇总结算到病人预交记录
  If Nvl(预约挂号_In, 0) = 0 And Nvl(记帐费用_In, 0) = 0 Then
    If Nvl(现金支付_In, 0) = 0 And Nvl(个帐支付_In, 0) = 0 And Nvl(预交支付_In, 0) = 0 And 序号_In = 1 Then
      Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      Insert Into 病人预交记录
        (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
         结算性质)
      Values
        (n_预交id, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), v_现金, 0, 登记时间_In, 操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费',
         n_组id, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In, 4);
    End If;
    If Nvl(现金支付_In, 0) <> 0 And 序号_In = 1 Then
      v_结算内容     := 结算方式_In || '|'; --以空格分开以|结尾,没有结算号码的
      v_结算方式记录 := '';
      While v_结算内容 Is Not Null Loop
        v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '|') - 1);
        v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
      
        v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1));
      
        v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        v_结算号码 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
      
        v_当前结算   := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        n_三方卡标志 := To_Number(v_当前结算);
      
        If Instr('|' || v_结算方式记录 || '|', '|' || Nvl(v_结算方式, v_现金) || '|') <> 0 Then
          v_Err_Msg := '使用了重复的结算方式,请检查!';
          Raise Err_Item;
        Else
          v_结算方式记录 := v_结算方式记录 || '|' || Nvl(v_结算方式, v_现金);
        End If;
      
        If n_三方卡标志 = 0 Then
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明,
             合作单位, 结算性质, 结算号码)
          Values
            (n_预交id, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(v_结算方式, v_现金), Nvl(n_结算金额, 0), 登记时间_In,
             操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费', n_组id, Null, Null, Null, Null, Null, 合作单位_In, 4, v_结算号码);
        Else
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明,
             合作单位, 结算性质, 结算号码)
          Values
            (n_预交id, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(v_结算方式, v_现金), Nvl(n_结算金额, 0), 登记时间_In,
             操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费', n_组id, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In, 4,
             v_结算号码);
          If Nvl(结算卡序号_In, 0) <> 0 Then
            n_消费卡id := Null;
            Begin
              Select Nvl(自制卡, 0), 1 Into n_自制卡, n_Count From 卡消费接口目录 Where 编号 = 结算卡序号_In;
            Exception
              When Others Then
                n_Count := 0;
            End;
            If n_Count = 0 Then
              v_Err_Msg := '没有发现原结算卡的相应类别,不能继续操作！';
              Raise Err_Item;
            End If;
            If n_自制卡 = 1 Then
              Select ID
              Into n_消费卡id
              From 消费卡目录
              Where 接口编号 = 结算卡序号_In And 卡号 = 卡号_In And
                    序号 = (Select Max(序号) From 消费卡目录 Where 接口编号 = 结算卡序号_In And 卡号 = 卡号_In);
            End If;
            Zl_病人卡结算记录_Insert(结算卡序号_In, n_消费卡id, Nvl(v_结算方式, v_现金), Nvl(n_结算金额, 0), 卡号_In, Null, 登记时间_In, Null, 结帐id_In,
                              n_预交id);
          End If;
        End If;
      
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) + n_结算金额
        Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = Nvl(v_结算方式, v_现金)
        Returning 余额 Into n_返回值;
      
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额
            (收款员, 结算方式, 性质, 余额)
          Values
            (操作员姓名_In, Nvl(v_结算方式, v_现金), 1, n_结算金额);
          n_返回值 := n_结算金额;
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 人员缴款余额
          Where 收款员 = 操作员姓名_In And 结算方式 = Nvl(v_结算方式, v_现金) And 性质 = 1 And Nvl(余额, 0) = 0;
        End If;
      
        v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '|') + 1);
      End Loop;
    End If;
  
    --对于医保挂号
    If Nvl(个帐支付_In, 0) <> 0 And 序号_In = 1 Then
      Insert Into 病人预交记录
        (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 结算性质)
      Values
        (病人预交记录_Id.Nextval, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), v_个人帐户, 个帐支付_In, 登记时间_In, 操作员编号_In,
         操作员姓名_In, 结帐id_In, '医保挂号', n_组id, 4);
    End If;
  
    --对于就诊卡通过预交金挂号
    If Nvl(预交支付_In, 0) <> 0 And 序号_In = 1 Then
      n_预交金额 := 预交支付_In;
      For r_Deposit In c_Deposit(病人id_In, v_冲预交病人ids) Loop
        n_当前金额 := Case
                    When r_Deposit.金额 - n_预交金额 < 0 Then
                     r_Deposit.金额
                    Else
                     n_预交金额
                  End;
      
        If r_Deposit.Id <> 0 Then
          --第一次冲预交(填上结帐ID,金额为0)
          Update 病人预交记录 Set 冲预交 = 0, 结帐id = 结帐id_In, 结算性质 = 4 Where ID = r_Deposit.Id;
        
        End If;
        --冲上次剩余额
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 预交类别, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,
           冲预交, 结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
          Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 预交类别, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号,
                 登记时间_In, 操作员姓名_In, 操作员编号_In, n_当前金额, 结帐id_In, n_组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
          From 病人预交记录
          Where NO = r_Deposit.No And 记录状态 = r_Deposit.记录状态 And 记录性质 In (1, 11) And Rownum = 1;
      
        --更新病人预交余额
        Update 病人余额
        Set 预交余额 = Nvl(预交余额, 0) - n_当前金额
        Where 病人id = r_Deposit.病人id And 性质 = 1 And 类型 = Nvl(r_Deposit.预交类别, 2);
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
      If n_预交金额 > 0 Then
        v_Err_Msg := '预交余不够支付本次支付金额,不能继续操作！';
        Raise Err_Item;
      
      End If;
      Delete From 病人余额 Where 病人id = 病人id_In And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
    End If;
  
    --相关汇总表的处理
    --人员缴款余额
    If 序号_In = 1 Then
      If Nvl(个帐支付_In, 0) <> 0 Then
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) + 个帐支付_In
        Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = v_个人帐户
        Returning 余额 Into n_返回值;
      
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (操作员姓名_In, v_个人帐户, 1, 个帐支付_In);
          n_返回值 := 个帐支付_In;
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 人员缴款余额
          Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_个人帐户 And Nvl(余额, 0) = 0;
        End If;
      End If;
    End If;
  End If;

  --病人挂号汇总(只处理一次,且单独收取病历费不处理)
  If Nvl(预约挂号_In, 0) = 0 Then
    If Nvl(记帐费用_In, 0) = 0 Then
      --处理票据使用情况
      If 序号_In = 1 And 票据号_In Is Not Null Then
        Select 票据打印内容_Id.Nextval Into n_打印id From Dual;
      
        --发出票据
        Insert Into 票据打印内容 (ID, 数据性质, NO) Values (n_打印id, 4, 单据号_In);
      
        Insert Into 票据使用明细
          (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
        Values
          (票据使用明细_Id.Nextval, Decode(收费票据_In, 1, 1, 4), 票据号_In, 1, 1, 领用id_In, n_打印id, 登记时间_In, 操作员姓名_In);
      
        --状态改动
        Update 票据领用记录
        Set 当前号码 = 票据号_In, 剩余数量 = Decode(Sign(剩余数量 - 1), -1, 0, 剩余数量 - 1), 使用时间 = Sysdate
        Where ID = Nvl(领用id_In, 0);
      End If;
    End If;
    --病人本次就诊(以发生时间为准)
    If Nvl(病人id_In, 0) <> 0 And 序号_In = 1 Then
      Update 病人信息 Set 就诊时间 = 发生时间_In, 就诊状态 = 1, 就诊诊室 = 诊室_In Where 病人id = 病人id_In;
    End If;
  End If;

  If Nvl(预约挂号_In, 0) = 0 And Nvl(记帐费用_In, 0) = 1 Then
    --记帐
    If Nvl(病人id_In, 0) = 0 Then
      v_Err_Msg := '要针对病人的挂号费进行记帐，必须是建档病人才能记帐挂号。';
      Raise Err_Item;
    End If;
  
    --病人余额
    Update 病人余额
    Set 费用余额 = Nvl(费用余额, 0) + Nvl(实收金额_In, 0)
    Where 病人id = Nvl(病人id_In, 0) And 性质 = 1 And 类型 = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人余额 (病人id, 性质, 类型, 费用余额, 预交余额) Values (病人id_In, 1, 1, Nvl(实收金额_In, 0), 0);
    End If;
  
    --病人未结费用
    Update 病人未结费用
    Set 金额 = Nvl(金额, 0) + Nvl(实收金额_In, 0)
    Where 病人id = 病人id_In And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(病人科室id, 0) = Nvl(病人科室id_In, 0) And
          Nvl(开单部门id, 0) = Nvl(开单部门id_In, 0) And Nvl(执行部门id, 0) = Nvl(执行部门id_In, 0) And 收入项目id + 0 = 收入项目id_In And
          来源途径 + 0 = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人未结费用
        (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
      Values
        (病人id_In, Null, Null, 病人科室id_In, 开单部门id_In, 执行部门id_In, 收入项目id_In, 1, Nvl(实收金额_In, 0));
    End If;
  End If;

  --病人挂号记录
  If 号别_In Is Not Null And 序号_In = 1 Then
    --And Nvl(预约挂号_In, 0) = 0
    Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
    Begin
      Select 名称 Into v_付款方式 From 医疗付款方式 Where 编码 = 付款方式_In And Rownum < 2;
    Exception
      When Others Then
        v_付款方式 := Null;
    End;
    Insert Into 病人挂号记录
      (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 预约时间, 操作员编号,
       操作员姓名, 复诊, 号序, 社区, 预约, 预约方式, 摘要, 交易流水号, 交易说明, 合作单位, 接收时间, 接收人, 预约操作员, 预约操作员编号, 险类, 医疗付款方式, 出诊记录id)
    Values
      (n_挂号id, 单据号_In, Decode(Nvl(预约挂号_In, 0), 1, 2, 1), 1, 病人id_In, 门诊号_In, 姓名_In, 性别_In, 年龄_In, 号别_In, 急诊_In, 诊室_In,
       Null, 执行部门id_In, 医生姓名_In, 0, Null, 登记时间_In, 发生时间_In, Decode(Nvl(预约挂号_In, 0), 1, 发生时间_In, Null), 操作员编号_In,
       操作员姓名_In, 复诊_In, n_序号, 社区_In, Decode(预约接收_In, 1, 1, 0), 预约方式_In, 摘要_In, 交易流水号_In, 交易说明_In, 合作单位_In,
       Decode(Nvl(预约挂号_In, 0), 0, 登记时间_In, Null), Decode(Nvl(预约挂号_In, 0), 0, 操作员姓名_In, Null),
       Decode(Nvl(预约挂号_In, 0), 1, 操作员姓名_In, Null), Decode(Nvl(预约挂号_In, 0), 1, 操作员编号_In, Null),
       Decode(Nvl(险类_In, 0), 0, Null, 险类_In), v_付款方式, 出诊记录id_In);
  
    n_预约生成队列 := 0;
    If Nvl(预约挂号_In, 0) = 1 Then
      n_预约生成队列 := Zl_To_Number(zl_GetSysParameter('预约生成队列', 1113));
    End If;
  
    --0-不产生队列;1-按医生或分诊台排队;2-先分诊,后医生站
    If Nvl(生成队列_In, 0) <> 0 And Nvl(预约挂号_In, 0) = 0 Or n_预约生成队列 = 1 Then
      n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113));
      If Nvl(n_分诊台签到排队, 0) = 0 Or n_预约生成队列 = 1 Then
      
        --产生队列
        --.按”执行部门” 的方式生成队列
        v_队列名称 := 执行部门id_In;
        v_排队号码 := Zlgetnextqueue(执行部门id_In, n_挂号id, 号别_In || '|' || n_序号);
        v_排队序号 := Zlgetsequencenum(0, n_挂号id, 0);
        d_排队时间 := Zl_Get_Queuedate(n_挂号id, 号别_In, n_序号, 发生时间_In);
        --  队列名称_In , 业务类型_In, 业务id_In,科室id_In,排队号码_In,排队标记_In,患者姓名_In,病人ID_IN, 诊室_In, 医生姓名_In, 排队时间_In
        Zl_排队叫号队列_Insert(v_队列名称, 0, n_挂号id, 执行部门id_In, v_排队号码, Null, 姓名_In, 病人id_In, 诊室_In, 医生姓名_In, d_排队时间, 预约方式_In,
                         Null, v_排队序号);
      End If;
    End If;
  End If;
  --病人担保信息
  If 病人id_In Is Not Null And 序号_In = 1 Then
    --取参数:
    If Nvl(n_结算模式, 0) <> Nvl(结算模式_In, 0) Then
      --结算模式的确定
    
      v_Err_Msg := Null;
      Begin
        Select Nvl(结算模式, 0) Into n_结算模式 From 病人信息 Where 病人id = 病人id_In;
      Exception
        When Others Then
          v_Err_Msg := '未找到指定的病人信息,不允许挂号';
      End;
    
      If v_Err_Msg Is Not Null Then
        Raise Err_Item;
      End If;
      If n_结算模式 = 1 And Nvl(结算模式_In, 0) = 0 Then
        --病人已经是"先诊疗后结算的",本次是"先结算后诊疗的",则检查是否存在未结数据
        Select Count(1)
        Into n_Count
        From 病人未结费用
        Where 病人id = 病人id_In And (来源途径 In (1, 4) Or 来源途径 = 3 And Nvl(主页id, 0) = 0) And Nvl(金额, 0) <> 0 And Rownum < 2;
        If Nvl(n_Count, 0) <> 0 Then
          --存在未结算数据，必须先结算后才允许执行
          v_Err_Msg := '当前病人的就诊模式为先诊疗后结算且存在未结费用，不允许调整该病人的就诊模式,你可以先对未结费用结帐后再挂号或不调整病人的就诊模式!';
          Raise Err_Item;
        End If;
        --检查
        --未发生医嘱业务的（即当时就挂号的,需要保证同一次的就诊模式是一至的(程序已经检查，不用再处理)
      End If;
      Update 病人信息 Set 结算模式 = 结算模式_In Where 病人id = 病人id_In;
    End If;
  
    Update 病人信息
    Set 担保人 = Null, 担保额 = Null, 担保性质 = Null
    Where 病人id = 病人id_In And Exists (Select 1
           From 病人担保记录
           Where 病人id = 病人id_In And 主页id Is Not Null And
                 登记时间 = (Select Max(登记时间) From 病人担保记录 Where 病人id = 病人id_In));
  
    If Sql%RowCount > 0 Then
      Update 病人担保记录
      Set 到期时间 = Sysdate
      Where 病人id = 病人id_In And 主页id Is Not Null And Nvl(到期时间, Sysdate) > Sysdate;
    End If;
  End If;
  If 序号_In = 1 Then
    --消息推送
    Begin
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 1, n_挂号id;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_001(n_挂号id, 单据号_In);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人挂号记录_出诊_Insert;
/



Create Or Replace Procedure Zl_病人挂号记录_更新诊室
(
  No_In       病人挂号记录.No%Type := Null,
  病人id_In   病人挂号记录.病人id%Type := Null,
  诊室_In     病人挂号记录.诊室%Type := Null,
  医生_In     病人挂号记录.执行人%Type := Null,
  分诊时间_In 病人挂号记录.分诊时间%Type := Null,
  更新诊室_In Integer := 1,
  预约方式_In 预约方式.名称%Type := Null
) As
  v_Id           门诊费用记录.Id%Type := Null;
  v_挂号生成队列 Varchar2(2);
  v_排队号码     排队叫号队列.排队号码%Type;
  v_排队序号     排队叫号队列.排队序号%Type;
  n_当天排队     Number(18);
  n_单据性质     病人挂号记录.记录性质%Type;
  n_诊室id       门诊诊室.Id%Type;
  n_排队         Number(18);
  v_Error        Varchar2(255);
  Err_Custom Exception;
Begin
  If Nvl(更新诊室_In, 0) = 2 Then
    Begin
      Select ID Into v_Id From 病人挂号记录 Where NO = No_In;
      Update 排队叫号队列 Set 诊室 = 诊室_In, 医生姓名 = 医生_In Where 业务id = v_Id;
    Exception
      When Others Then
        Null;
    End;
  Else
    Begin
      Select ID, 记录性质 Into v_Id, n_单据性质 From 病人挂号记录 Where NO = No_In And Nvl(执行状态, 0) = 0;
      Select ID Into n_诊室id From 门诊诊室 Where 名称 = 诊室_In;
    Exception
      When Others Then
        Null;
    End;
    If v_Id Is Null Then
      v_Error := '病人已经接诊或已经退号，不能再分诊。';
      Raise Err_Custom;
    End If;
    If Nvl(更新诊室_In, 0) = 1 Then
      If Nvl(病人id_In, 0) <> 0 Then
        --更新病人信息
        Update 病人信息 Set 就诊诊室 = 诊室_In Where 病人id = 病人id_In And 就诊状态 = 1;
      End If;
    
      --更新费用记录
      Update 门诊费用记录
      Set 发药窗口 = 诊室_In, 执行人 = 医生_In
      Where 记录性质 = 4 And 记录状态 = Decode(Nvl(n_单据性质, 0), 2, 0, 1) And NO = No_In;
      --更新病人挂号记录
      Update 病人挂号记录
      Set 诊室 = 诊室_In, 执行人 = 医生_In, 分诊时间 = Decode(分诊时间_In, Null, 分诊时间, 分诊时间_In)
      Where NO = No_In;
    End If;
    v_挂号生成队列 := Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113));
    If v_挂号生成队列 <> 0 Then
      For c_挂号 In (Select ID, 执行部门id, 姓名, 诊室_In As 诊室, 登记时间, 医生_In As 执行人, 病人id, 号别, 号序
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
                           Nvl(分诊时间_In, Sysdate), 预约方式_In, Null, v_排队序号);
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
      End Loop;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人挂号记录_更新诊室;
/


Create Or Replace Procedure Zl_病人挂号记录_换号
(
  No_In         病人挂号记录.No%Type,
  号别_In       病人挂号记录.号别%Type,
  诊室_In       病人挂号记录.诊室%Type,
  科室id_In     病人挂号记录.执行部门id%Type,
  原医生_In     病人挂号记录.执行人%Type,
  原医生id_In   病人挂号汇总.医生id%Type,
  新医生_In     病人挂号记录.执行人%Type,
  新医生id_In   病人挂号汇总.医生id%Type,
  出诊记录id_In 临床出诊记录.Id%Type := Null
  --功能：完成病人换号功能，在挂号项目ID相同的情况下。
) As
  Cursor c_Bill Is
    Select ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 记帐单id, 病人id, 医嘱序号, 门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号, 付款方式, 病人科室id, 费别,
           收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人, 开单部门id, 开单人, 发生时间,
           登记时间, 执行部门id, 执行人, 执行状态, 执行时间, 结论, 操作员编号, 操作员姓名, 结帐id, 结帐金额, 保险大类id, 保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要, 是否急诊
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 = 1 And NO = No_In
    Order By 序号;

  v_病人id       门诊费用记录.Id%Type;
  v_现队列名称   排队叫号队列.队列名称%Type;
  v_挂号生成队列 Varchar2(2);
  v_预约挂号     Number(2);
  n_业务id       病人挂号记录.Id%Type;
  v_排队号码     排队叫号队列.排队号码%Type;
  v_号别         病人挂号记录.号别%Type;
  n_号序         病人挂号记录.号序%Type;
  v_排队序号     排队叫号队列.排队序号%Type;
  v_Temp         Varchar2(500);
  v_操作员编号   就诊变动记录.操作员编号%Type;
  v_操作员姓名   就诊变动记录.操作员姓名%Type;
  n_医生id       人员表.Id%Type;
  n_诊室id       门诊诊室.Id%Type;
  n_原出诊记录id 临床出诊记录.Id%Type;
  n_变动id       就诊变动记录.Id%Type;
  v_Error        Varchar2(255);
  Err_Custom Exception;
Begin
  v_病人id := 0;
  If 出诊记录id_In Is Null Then
    Begin
      Select 病人id Into v_病人id From 病人挂号记录 Where NO = No_In And 记录性质 = 1 And 记录状态 = 1;
    Exception
      When Others Then
        Null;
    End;
    If v_病人id = 0 Then
      v_Error := '没有找到病人的挂号信息。';
      Raise Err_Custom;
    Elsif v_病人id Is Null Then
      v_Error := '没有找到病人信息。';
      Raise Err_Custom;
    End If;
  
    ---先更新病人信息的就诊诊室和状态
    Update 病人信息 Set 就诊诊室 = 诊室_In, 就诊状态 = 1 Where 病人id = v_病人id And 就诊状态 In (1, 2);
  
    For r_Bill In c_Bill Loop
      If r_Bill.序号 = 1 Then
      
        --需要确定是否预约挂号
        --1.如果是退预约挂号产生的挂号记录,则需要减已挂数和其中已接数
        --2.如果是正常挂号,则只减已挂数
        Begin
          Select Decode(预约, Null, 0, 0, 0, 1) Into v_预约挂号 From 病人挂号记录 Where NO = r_Bill.No And Rownum = 1;
        Exception
          When Others Then
            v_预约挂号 := 0;
        End;
      
        --恢复以前的挂号汇总
        Update 病人挂号汇总
        Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - v_预约挂号, 已约数 = Nvl(已约数, 0) - v_预约挂号
        Where 日期 = Trunc(r_Bill.登记时间) And Nvl(科室id, 0) = Nvl(r_Bill.执行部门id, 0) And Nvl(项目id, 0) = Nvl(r_Bill.收费细目id, 0) And
              (号码 = r_Bill.计算单位 Or 号码 Is Null);
        If Sql%RowCount = 0 Then
          Insert Into 病人挂号汇总
            (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 已约数, 其中已接收)
          Values
            (Trunc(r_Bill.登记时间), r_Bill.执行部门id, r_Bill.收费细目id, 原医生_In, Decode(原医生id_In, 0, Null, 原医生id_In), r_Bill.计算单位,
             -1, -1 * v_预约挂号, -1 * v_预约挂号);
        End If;
      
        ----然后再更新挂号汇总
        Update 病人挂号汇总
        Set 已挂数 = Nvl(已挂数, 0) + 1, 其中已接收 = Nvl(其中已接收, 0) + v_预约挂号, 已约数 = Nvl(已约数, 0) + v_预约挂号
        Where 日期 = Trunc(r_Bill.登记时间) And Nvl(科室id, 0) = 科室id_In And Nvl(项目id, 0) = Nvl(r_Bill.收费细目id, 0) And
              (号码 = 号别_In Or 号码 Is Null);
        If Sql%RowCount = 0 Then
          Insert Into 病人挂号汇总
            (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 已约数, 其中已接收)
          Values
            (Trunc(r_Bill.登记时间), 科室id_In, r_Bill.收费细目id, 新医生_In, Decode(新医生id_In, 0, Null, 新医生id_In), 号别_In, 1, v_预约挂号,
             v_预约挂号);
        End If;
      End If;
    
      ---更新挂号记录
      Update 门诊费用记录
      Set 执行部门id = 科室id_In, 病人科室id = 科室id_In, 计算单位 = 号别_In, 发药窗口 = 诊室_In,
          --病人病区id = 科室id_In,
          执行人 = 新医生_In, 执行状态 = 0, 执行时间 = Null
      Where ID = r_Bill.Id;
    
      --更新病人挂号记录
      If r_Bill.序号 = 1 Then
        v_Temp := Zl_Identity(1);
        Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_Temp From Dual;
        Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into v_操作员编号 From Dual;
        Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_操作员姓名 From Dual;
        Begin
          Select ID Into n_医生id From 人员表 Where 姓名 = 新医生_In And Rownum < 2;
        Exception
          When Others Then
            n_医生id := Null;
        End;
        Select 就诊变动记录_Id.Nextval Into n_变动id From Dual;
        b_Message.Zlhis_Regist_005(r_Bill.No, n_变动id, 2);
        Zl_就诊变动记录_Insert(r_Bill.No, 2, '分诊换号', v_操作员姓名, v_操作员编号, 号别_In, 科室id_In, Null, n_医生id, 新医生_In, 诊室_In, n_号序,
                         Null, n_变动id);
        v_挂号生成队列 := Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113));
        If v_挂号生成队列 <> 0 Then
          v_现队列名称 := 科室id_In;
          Select ID, 号别, Nvl(号序, 0)
          Into n_业务id, v_号别, n_号序
          From 病人挂号记录
          Where NO = r_Bill.No And Rownum = 1;
          --Zlgetnextqueue(执行部门id_In Number,业务id_In     Number := Null)
          v_排队号码 := Zlgetnextqueue(科室id_In, n_业务id, v_号别 || '|' || n_号序);
          v_排队序号 := Zlgetsequencenum(0, n_业务id, 1);
          --新队列名称_IN, 业务类型_In, 业务id_In , 科室id_In , 患者姓名_In , 诊室_In , 医生姓名_In
          Zl_排队叫号队列_Update(v_现队列名称, 0, n_业务id, 科室id_In, r_Bill.姓名, 诊室_In, 新医生_In, v_排队号码, v_排队序号);
        End If;
        Update 病人挂号记录
        Set 执行部门id = 科室id_In, 号别 = 号别_In, 诊室 = 诊室_In, 执行人 = 新医生_In, 执行状态 = 0, 执行时间 = Null
        Where NO = r_Bill.No;
      End If;
    End Loop;
  Else
    --出诊表排班模式
    Begin
      Select 病人id, 出诊记录id
      Into v_病人id, n_原出诊记录id
      From 病人挂号记录
      Where NO = No_In And 记录性质 = 1 And 记录状态 = 1;
      Select ID Into n_诊室id From 门诊诊室 Where 名称 = 诊室_In;
    Exception
      When Others Then
        Null;
    End;
    If v_病人id = 0 Then
      v_Error := '没有找到病人的挂号信息。';
      Raise Err_Custom;
    Elsif v_病人id Is Null Then
      v_Error := '没有找到病人信息。';
      Raise Err_Custom;
    End If;
  
    ---先更新病人信息的就诊诊室和状态
    Update 病人信息 Set 就诊诊室 = 诊室_In, 就诊状态 = 1 Where 病人id = v_病人id And 就诊状态 In (1, 2);
  
    For r_Bill In c_Bill Loop
      If r_Bill.序号 = 1 Then
      
        --需要确定是否预约挂号
        --1.如果是退预约挂号产生的挂号记录,则需要减已挂数和其中已接数
        --2.如果是正常挂号,则只减已挂数
        Begin
          Select Decode(预约, Null, 0, 0, 0, 1) Into v_预约挂号 From 病人挂号记录 Where NO = r_Bill.No And Rownum = 1;
        Exception
          When Others Then
            v_预约挂号 := 0;
        End;
      
        --恢复以前的挂号汇总
        Update 病人挂号汇总
        Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - v_预约挂号, 已约数 = Nvl(已约数, 0) - v_预约挂号
        Where 日期 = Trunc(r_Bill.登记时间) And Nvl(医生id, 0) = Nvl(原医生id_In, 0) And Nvl(医生姓名, '-') = Nvl(原医生_In, '-') And
              Nvl(科室id, 0) = Nvl(r_Bill.执行部门id, 0) And Nvl(项目id, 0) = Nvl(r_Bill.收费细目id, 0) And
              (号码 = r_Bill.计算单位 Or 号码 Is Null);
        If Sql%RowCount = 0 Then
          Insert Into 病人挂号汇总
            (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 已约数, 其中已接收)
          Values
            (Trunc(r_Bill.登记时间), r_Bill.执行部门id, r_Bill.收费细目id, 原医生_In, Decode(原医生id_In, 0, Null, 原医生id_In), r_Bill.计算单位,
             -1, -1 * v_预约挂号, -1 * v_预约挂号);
        End If;
        Update 临床出诊记录
        Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - v_预约挂号, 已约数 = Nvl(已约数, 0) - v_预约挂号
        Where ID = n_原出诊记录id;
      
        ----然后再更新挂号汇总
        Update 病人挂号汇总
        Set 已挂数 = Nvl(已挂数, 0) + 1, 其中已接收 = Nvl(其中已接收, 0) + v_预约挂号, 已约数 = Nvl(已约数, 0) + v_预约挂号
        Where 日期 = Trunc(r_Bill.登记时间) And Nvl(科室id, 0) = 科室id_In And Nvl(项目id, 0) = Nvl(r_Bill.收费细目id, 0) And
              (号码 = 号别_In Or 号码 Is Null);
        If Sql%RowCount = 0 Then
          Insert Into 病人挂号汇总
            (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 已约数, 其中已接收)
          Values
            (Trunc(r_Bill.登记时间), 科室id_In, r_Bill.收费细目id, 新医生_In, Decode(新医生id_In, 0, Null, 新医生id_In), 号别_In, 1, v_预约挂号,
             v_预约挂号);
        End If;
        Update 临床出诊记录
        Set 已挂数 = Nvl(已挂数, 0) + 1, 其中已接收 = Nvl(其中已接收, 0) + v_预约挂号, 已约数 = Nvl(已约数, 0) + v_预约挂号
        Where ID = 出诊记录id_In;
      End If;
    
      ---更新挂号记录
      Update 门诊费用记录
      Set 执行部门id = 科室id_In, 病人科室id = 科室id_In, 计算单位 = 号别_In, 发药窗口 = 诊室_In,
          --病人病区id = 科室id_In,
          执行人 = 新医生_In, 执行状态 = 0, 执行时间 = Null
      Where ID = r_Bill.Id;
    
      --更新病人挂号记录
      If r_Bill.序号 = 1 Then
        v_Temp := Zl_Identity(1);
        Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_Temp From Dual;
        Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into v_操作员编号 From Dual;
        Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_操作员姓名 From Dual;
        Begin
          Select ID Into n_医生id From 人员表 Where 姓名 = 新医生_In And Rownum < 2;
        Exception
          When Others Then
            n_医生id := Null;
        End;
        Select 就诊变动记录_Id.Nextval Into n_变动id From Dual;
        b_Message.Zlhis_Regist_005(r_Bill.No, n_变动id, 2);
        Zl_就诊变动记录_Insert(r_Bill.No, 2, '分诊换号', v_操作员姓名, v_操作员编号, 号别_In, 科室id_In, Null, n_医生id, 新医生_In, 诊室_In, n_号序,
                         Null, n_变动id);
        v_挂号生成队列 := Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113));
        If v_挂号生成队列 <> 0 Then
          v_现队列名称 := 科室id_In;
          Select ID, 号别, Nvl(号序, 0)
          Into n_业务id, v_号别, n_号序
          From 病人挂号记录
          Where NO = r_Bill.No And Rownum = 1;
          --Zlgetnextqueue(执行部门id_In Number,业务id_In     Number := Null)
          v_排队号码 := Zlgetnextqueue(科室id_In, n_业务id, v_号别 || '|' || n_号序);
          v_排队序号 := Zlgetsequencenum(0, n_业务id, 1);
          --新队列名称_IN, 业务类型_In, 业务id_In , 科室id_In , 患者姓名_In , 诊室_In , 医生姓名_In
          Zl_排队叫号队列_Update(v_现队列名称, 0, n_业务id, 科室id_In, r_Bill.姓名, 诊室_In, 新医生_In, v_排队号码, v_排队序号);
        End If;
        Update 病人挂号记录
        Set 执行部门id = 科室id_In, 号别 = 号别_In, 诊室 = 诊室_In, 执行人 = 新医生_In, 执行状态 = 0, 执行时间 = Null, 出诊记录id = 出诊记录id_In
        Where NO = r_Bill.No;
      End If;
    End Loop;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人挂号记录_换号;
/


Create Or Replace Procedure Zl_病人挂号记录_批量换号
(
  Nos_In        In Varchar2 := Null,
  新号别_In     In 病人挂号记录.号别%Type := Null,
  新医生姓名_In In 挂号安排.医生姓名%Type := Null,
  新医生id_In   In 挂号安排.医生id%Type := Null,
  新科室id_In   In 挂号安排.科室id%Type := Null,
  原医生姓名_In In 挂号安排.医生姓名%Type := Null,
  原医生id_In   In 挂号安排.医生id%Type := Null,
  原号别_In     In 病人挂号记录.号别%Type := Null,
  操作员姓名_In In 挂号序号状态.操作员姓名%Type := Null,
  原出诊id_In   In 临床出诊记录.Id%Type := Null,
  新出诊id_In   In 临床出诊记录.Id%Type := Null
  --功能: 完成病人批量换号功能,在挂号项目相同,限号数相同,限约数相同,科室相同的情况下。
  --参数说明:  Nos_In :需要跟换排班的病人挂号记录单据集:格式: M000001|M000002|..........
) As
  --获取对应挂号记录的门诊费用记录信息
  Cursor c_Bill(c_No 病人挂号记录.No%Type) Is
    Select ID, 序号, NO, 发生时间, 执行部门id, 收费细目id, 计算单位
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 In (0, 1) And NO = c_No
    Order By 序号;
  --获取相应排班的分诊诊室
  Cursor c_平均分诊(c_指定分诊号表id 挂号安排.Id%Type) Is
    Select 号表id, 门诊诊室, 当前分配 From 挂号安排诊室 Where 号表id = c_指定分诊号表id;
  Cursor c_出诊平均分诊 Is
    Select 记录id, 诊室id, 当前分配 From 临床出诊诊室记录 Where 记录id = 新出诊id_In;

  --变量定义
  r_平均分诊         挂号安排诊室%RowType;
  r_出诊平均分诊     临床出诊诊室记录%RowType;
  r_Bill             c_Bill%RowType;
  v_Nos              Varchar(2000);
  v_No               病人挂号记录.No%Type;
  n_病人id           病人挂号记录.病人id%Type;
  n_原序号           病人挂号记录.号序%Type;
  d_原就诊日期       病人挂号记录.预约时间%Type;
  n_是否已被挂出     Number(1);
  n_记录性质         Number(1);
  n_预约             Number(1);
  n_挂号状态         Number(1); --0-正常挂号:1-预约挂号;2-预约挂号接收
  v_新就诊诊室       病人信息.就诊诊室%Type;
  n_分诊方式         Number(1); --0-不分诊:1-指定分诊:2-动态分诊:3-平均分诊
  n_指定分诊号表id   Number(10);
  n_分诊诊室数量     Number(3);
  n_是否找到分诊诊室 Number(1); --0:未找到:1-找到但分配标识未更改:2-修改第一条数据标识
  n_Index            Number(1); --当前记录集的索引值
  v_现队列名称       排队叫号队列.队列名称%Type;
  n_挂号生成队列     Number;
  n_预约生成队列     Number;
  v_排队号码         排队叫号队列.排队号码%Type;
  n_业务id           病人挂号记录.Id%Type;
  v_Temp             Varchar2(500);
  v_操作员编号       就诊变动记录.操作员编号%Type;
  v_操作员姓名       就诊变动记录.操作员姓名%Type;
  n_医生id           人员表.Id%Type;
  v_Error            Varchar2(255);
  n_挂号排班模式     Number(3);
  n_出诊记录id       临床出诊记录.Id%Type;
  n_变动id           就诊变动记录.Id%Type;
  Err_Custom Exception;
Begin
  n_挂号排班模式 := Zl_To_Number(Nvl(zl_GetSysParameter('挂号排班模式'), 0));
  If n_挂号排班模式 = 0 Then
    --计划排班模式
    --检查是否存在该挂号记录
    If Nos_In Is Not Null Then
      v_Nos := Nos_In || '|';
      While v_Nos Is Not Null Loop
        --初始化变量
        n_病人id           := 0;
        n_原序号           := 0;
        d_原就诊日期       := Null;
        n_是否已被挂出     := 0;
        n_记录性质         := 0;
        n_预约             := 0;
        n_挂号状态         := 0;
        v_新就诊诊室       := '';
        n_分诊方式         := 0;
        n_指定分诊号表id   := 0;
        v_现队列名称       := '';
        v_排队号码         := '';
        n_业务id           := 0;
        n_分诊诊室数量     := 0;
        n_挂号生成队列     := 0;
        n_预约生成队列     := 0;
        n_是否找到分诊诊室 := 0;
        n_Index            := 0;
      
        v_No  := Substr(v_Nos, 1, Instr(v_Nos, '|') - 1);
        v_Nos := Substr(v_Nos, Instr(v_Nos, '|') + 1);
        --检查是否存在该挂号记录
        Begin
          Select a.Id, a.病人id, a.号序, Nvl(b.日期, Nvl(a.预约时间, a.发生时间)), a.记录性质, Nvl(a.预约, 0)
          Into n_业务id, n_病人id, n_原序号, d_原就诊日期, n_记录性质, n_预约
          From 病人挂号记录 A, 挂号序号状态 B
          Where NO = v_No And 记录性质 In (1, 2) And 记录状态 = 1 And a.号别 = b.号码(+) And
                Trunc(Nvl(预约时间, 发生时间)) = Trunc(b.日期(+)) And a.号序 = b.序号(+);
        Exception
          When Others Then
            Null;
        End;
        If n_病人id = 0 Then
          v_Error := '没有找到病人的挂号信息';
          Raise Err_Custom;
        End If;
        --判断当前挂号状态
        n_挂号状态 := 0; --正常挂号
        If n_记录性质 = 1 And n_预约 = 1 Then
          n_挂号状态 := 2; --预约接收
        End If;
        If n_记录性质 = 2 And n_预约 = 1 Then
          n_挂号状态 := 1; --预约
        End If;
      
        --检查换号的新号别是否已被挂出
        Begin
          Select a.状态
          Into n_是否已被挂出
          From 挂号序号状态 A
          Where a.日期 = d_原就诊日期 And a.号码 = 新号别_In And a.序号 = n_原序号;
        Exception
          When Others Then
            n_是否已被挂出 := 0;
        End;
        If n_是否已被挂出 > 0 Then
          v_Error := '要换的号别已被挂出';
          Raise Err_Custom;
        End If;
        --预约接收的情况下进行分诊诊室的获取
        If n_挂号状态 = 2 Then
          --获取新号别诊室
          --说明:预约的情况下，不需要分诊，因此不用获取就诊诊室
          --     接收的情况下,需要进行分诊,因此需要获取接诊诊室
          --获取分诊方式
          Begin
            Select ID, Nvl(分诊方式, 0) Into n_指定分诊号表id, n_分诊方式 From 挂号安排 Where 号码 = 新号别_In;
          Exception
            When Others Then
              n_分诊方式       := 0;
              n_指定分诊号表id := 0;
          End;
        
          Begin
            If n_分诊方式 = 0 Then
              --不分诊
              v_新就诊诊室 := '';
            End If;
            If n_分诊方式 = 1 Then
              --指定分诊
              Select 门诊诊室 Into v_新就诊诊室 From 挂号安排诊室 Where 号表id = n_指定分诊号表id;
            End If;
            If n_分诊方式 = 2 Then
              --动态分诊
              Select 门诊诊室
              Into v_新就诊诊室
              From (Select 门诊诊室, Sum(Num) As Num
                     From (Select 门诊诊室, 0 As Num
                            From 挂号安排诊室
                            Where 号表id = n_指定分诊号表id
                            Union All
                            Select 诊室, Count(诊室) As Num
                            From 病人挂号记录
                            Where Nvl(执行状态, 0) = 0 And 记录性质 = 1 And 记录状态 = 1 And 发生时间 Between Trunc(Sysdate) And Sysdate And
                                  号别 = 新号别_In And 诊室 In (Select 门诊诊室 From 挂号安排诊室 Where 号表id = n_指定分诊号表id)
                            Group By 诊室)
                     Group By 门诊诊室
                     Order By Num)
              Where Rownum = 1;
            End If;
            If n_分诊方式 = 3 Then
              --平均分诊
              --获取当前安排下的诊室数量
              Select Count(1) Into n_分诊诊室数量 From 挂号安排诊室 Where 号表id = n_指定分诊号表id;
            
              Open c_平均分诊(n_指定分诊号表id);
              Loop
                Fetch c_平均分诊
                  Into r_平均分诊;
                Exit When c_平均分诊%NotFound;
                n_Index := n_Index + 1;
                --找到了对应的分诊诊室,需要修改下一个诊室的当前分配为1(代表该诊室是下一次的分诊诊室)
                If n_是否找到分诊诊室 = 1 Then
                  Update 挂号安排诊室
                  Set 当前分配 = 1
                  Where 号表id = r_平均分诊.号表id And 门诊诊室 = r_平均分诊.门诊诊室;
                  Exit;
                End If;
              
                If Nvl(r_平均分诊.当前分配, 0) = 1 Then
                  v_新就诊诊室 := r_平均分诊.门诊诊室;
                  Update 挂号安排诊室
                  Set 当前分配 = 0
                  Where 号表id = r_平均分诊.号表id And 门诊诊室 = r_平均分诊.门诊诊室;
                  n_是否找到分诊诊室 := 1;
                End If;
              
                If n_分诊诊室数量 = 1 And n_是否找到分诊诊室 = 1 Then
                  n_是否找到分诊诊室 := 2;
                  Exit;
                End If;
                If n_分诊诊室数量 > 1 And n_是否找到分诊诊室 = 1 Then
                  --游标已经到了最后,所以需从第一条数据开始修改标识
                  If n_Index >= n_分诊诊室数量 Then
                    n_是否找到分诊诊室 := 2;
                    Exit;
                  End If;
                End If;
              End Loop;
              Close c_平均分诊;
              --重置索引值
              n_Index := 0;
              --第一次分诊
              If Nvl(v_新就诊诊室, ' ') = ' ' Or v_新就诊诊室 Is Null Then
                Open c_平均分诊(n_指定分诊号表id);
                Loop
                  Fetch c_平均分诊
                    Into r_平均分诊;
                  Exit When c_平均分诊%NotFound;
                  n_Index := n_Index + 1;
                
                  If n_是否找到分诊诊室 = 1 Then
                    Update 挂号安排诊室
                    Set 当前分配 = 1
                    Where 号表id = r_平均分诊.号表id And 门诊诊室 = r_平均分诊.门诊诊室;
                    Exit;
                  End If;
                
                  Update 挂号安排诊室
                  Set 当前分配 = 0
                  Where 号表id = r_平均分诊.号表id And 门诊诊室 = r_平均分诊.门诊诊室;
                  v_新就诊诊室 := r_平均分诊.门诊诊室;
                
                  n_是否找到分诊诊室 := 1;
                  If n_分诊诊室数量 = 1 And n_是否找到分诊诊室 = 1 Then
                    n_是否找到分诊诊室 := 2;
                    Exit;
                  End If;
                
                  If n_分诊诊室数量 > 1 And n_是否找到分诊诊室 = 1 Then
                    --游标已经到了最后,所以需从第一条数据开始修改标识
                    If n_Index >= n_分诊诊室数量 Then
                      n_是否找到分诊诊室 := 2;
                      Exit;
                    End If;
                  End If;
                
                End Loop;
                Close c_平均分诊;
              End If;
            
              If n_是否找到分诊诊室 = 2 Then
                Open c_平均分诊(n_指定分诊号表id);
                Loop
                  Fetch c_平均分诊
                    Into r_平均分诊;
                  Exit When c_平均分诊%NotFound;
                  Update 挂号安排诊室
                  Set 当前分配 = 1
                  Where 号表id = r_平均分诊.号表id And 门诊诊室 = r_平均分诊.门诊诊室;
                  Exit;
                End Loop;
                Close c_平均分诊;
              End If;
            End If;
          Exception
            When Others Then
              v_新就诊诊室 := '';
          End;
        End If;
      
        --更新病人信息的就诊诊室和状态
        Update 病人信息 Set 就诊诊室 = v_新就诊诊室, 就诊状态 = 1 Where 病人id = n_病人id And 就诊状态 In (1, 2);
      
        --打开游标
        Open c_Bill(v_No);
        Loop
          Fetch c_Bill
            Into r_Bill;
          Exit When c_Bill%NotFound;
          If r_Bill.序号 = 1 Then
            --需要确定是否预约挂号
            --1.如果是预约挂号产生的挂号记录,则需要减已约数
            --2.如果是预约挂号为接收的挂号记录,则需要减已挂数和其中已接收数
            --3.如果是正常挂号,则只减已挂数
            --恢复以前的挂号汇总
            Update 病人挂号汇总
            Set 已挂数 = Nvl(已挂数, 0) + Decode(n_挂号状态, 0, -1, 2, -1, 0), 其中已接收 = Nvl(其中已接收, 0) + Decode(n_挂号状态, 2, -1, 0),
                已约数 = Nvl(已约数, 0) + Decode(n_挂号状态, 0, 0, -1)
            Where 日期 = Trunc(r_Bill.发生时间) And Nvl(科室id, 0) = Nvl(r_Bill.执行部门id, 0) And
                  Nvl(项目id, 0) = Nvl(r_Bill.收费细目id, 0) And (号码 = r_Bill.计算单位 Or 号码 Is Null);
            If Sql%RowCount = 0 Then
              Insert Into 病人挂号汇总
                (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 已约数, 其中已接收)
              Values
                (Trunc(r_Bill.发生时间), r_Bill.执行部门id, r_Bill.收费细目id, 原医生姓名_In, Decode(原医生id_In, 0, Null, 原医生id_In),
                 r_Bill.计算单位, 0, 0, 0);
            End If;
          
            ----然后再更新挂号汇总
            Update 病人挂号汇总
            Set 已挂数 = Nvl(已挂数, 0) + Decode(n_挂号状态, 0, 1, 2, 1, 0), 其中已接收 = Nvl(其中已接收, 0) + Decode(n_挂号状态, 2, 1, 0),
                已约数 = Nvl(已约数, 0) + Decode(n_挂号状态, 0, 0, 1)
            Where 日期 = Trunc(r_Bill.发生时间) And Nvl(科室id, 0) = 新科室id_In And Nvl(项目id, 0) = Nvl(r_Bill.收费细目id, 0) And
                 
                  (号码 = 新号别_In Or 号码 Is Null);
            If Sql%RowCount = 0 Then
              Insert Into 病人挂号汇总
                (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 已约数, 其中已接收)
              Values
                (Trunc(r_Bill.发生时间), 新科室id_In, r_Bill.收费细目id, 新医生姓名_In, Decode(新医生id_In, 0, Null, 新医生id_In), 新号别_In,
                 Decode(n_挂号状态, 0, 1, 2, 1, 0), Decode(n_挂号状态, 0, 0, 1), Decode(n_挂号状态, 2, 1, 0));
            End If;
          End If;
        
          ---更新挂号记录
          If n_挂号状态 = 1 Then
            --预约
            Update 门诊费用记录
            Set 执行部门id = 新科室id_In, 病人科室id = 新科室id_In, 计算单位 = 新号别_In, 发药窗口 = n_原序号,
                --病人病区id = 科室id_In,
                执行人 = 新医生姓名_In, 执行状态 = 0, 执行时间 = Null
            Where ID = r_Bill.Id;
          Else
            --挂号或接收
            Update 门诊费用记录
            Set 执行部门id = 新科室id_In, 病人科室id = 新科室id_In, 计算单位 = 新号别_In, 发药窗口 = v_新就诊诊室,
                --病人病区id = 科室id_In,
                执行人 = 新医生姓名_In, 执行状态 = 0, 执行时间 = Null
            Where ID = r_Bill.Id;
          End If;
        
          --更新病人挂号记录
          If r_Bill.序号 = 1 Then
            v_Temp := Zl_Identity(1);
            Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_Temp From Dual;
            Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into v_操作员编号 From Dual;
            Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_操作员姓名 From Dual;
            Begin
              Select ID Into n_医生id From 人员表 Where 姓名 = 新医生姓名_In And Rownum < 2;
            Exception
              When Others Then
                n_医生id := Null;
            End;
            Select 就诊变动记录_Id.Nextval Into n_变动id From Dual;
            b_Message.Zlhis_Regist_005(r_Bill.No, n_变动id, 2);
            Zl_就诊变动记录_Insert(r_Bill.No, 1, '批量换号', v_操作员姓名, v_操作员编号, 新号别_In, 新科室id_In, Null, n_医生id, 新医生姓名_In, v_新就诊诊室,
                             n_原序号, Null, n_变动id);
            --修改队列信息
            Update 排队叫号队列
            Set 医生姓名 = 新医生姓名_In, 诊室 = v_新就诊诊室
            Where 业务id = n_业务id And 业务类型 = 0;
          
            Update 病人挂号记录
            Set 执行部门id = 新科室id_In, 号别 = 新号别_In, 诊室 = v_新就诊诊室, 执行人 = 新医生姓名_In, 执行状态 = 0, 执行时间 = Null, 号序 = n_原序号
            Where NO = r_Bill.No;
            --修改挂号序号状态
            If n_原序号 Is Not Null Then
              --1.恢复以前挂号序号状态
              Delete 挂号序号状态 Where 日期 = d_原就诊日期 And 序号 = n_原序号 And 号码 = 原号别_In;
              --2.新增换号后挂号序号状态
              Insert Into 挂号序号状态
                (号码, 日期, 序号, 操作员姓名, 状态, 预约, 登记时间)
              Values
                (新号别_In, d_原就诊日期, n_原序号, 操作员姓名_In, Decode(n_挂号状态, 1, 2, 1), Decode(n_挂号状态, 0, 0, 1), Sysdate);
            End If;
          End If;
        End Loop;
        Close c_Bill;
      End Loop;
    End If;
  Else
    --出诊表排班模式
    --检查是否存在该挂号记录
    If Nos_In Is Not Null Then
      v_Nos := Nos_In || '|';
      While v_Nos Is Not Null Loop
        --初始化变量
        n_病人id           := 0;
        n_原序号           := 0;
        d_原就诊日期       := Null;
        n_是否已被挂出     := 0;
        n_记录性质         := 0;
        n_预约             := 0;
        n_挂号状态         := 0;
        v_新就诊诊室       := '';
        n_分诊方式         := 0;
        n_指定分诊号表id   := 0;
        v_现队列名称       := '';
        v_排队号码         := '';
        n_业务id           := 0;
        n_分诊诊室数量     := 0;
        n_挂号生成队列     := 0;
        n_预约生成队列     := 0;
        n_是否找到分诊诊室 := 0;
        n_Index            := 0;
      
        v_No  := Substr(v_Nos, 1, Instr(v_Nos, '|') - 1);
        v_Nos := Substr(v_Nos, Instr(v_Nos, '|') + 1);
        --检查是否存在该挂号记录
        Begin
          Select a.Id, a.病人id, a.号序, Nvl(b.开始时间, Nvl(a.预约时间, a.发生时间)), a.记录性质, Nvl(a.预约, 0), a.出诊记录id
          Into n_业务id, n_病人id, n_原序号, d_原就诊日期, n_记录性质, n_预约, n_出诊记录id
          From 病人挂号记录 A, 临床出诊序号控制 B
          Where NO = v_No And 记录性质 In (1, 2) And 记录状态 = 1 And a.号序 = b.序号(+) And a.出诊记录id = b.记录id(+);
        Exception
          When Others Then
            Null;
        End;
        If n_病人id = 0 Then
          v_Error := '没有找到病人的挂号信息';
          Raise Err_Custom;
        End If;
        --判断当前挂号状态
        n_挂号状态 := 0; --正常挂号
        If n_记录性质 = 1 And n_预约 = 1 Then
          n_挂号状态 := 2; --预约接收
        End If;
        If n_记录性质 = 2 And n_预约 = 1 Then
          n_挂号状态 := 1; --预约
        End If;
      
        --检查换号的新号别是否已被挂出
        Begin
          Select a.挂号状态
          Into n_是否已被挂出
          From 临床出诊序号控制 A
          Where a.记录id = 新出诊id_In And a.序号 = n_原序号;
        Exception
          When Others Then
            n_是否已被挂出 := 0;
        End;
        If n_是否已被挂出 > 0 Then
          v_Error := '要换的号别已被挂出';
          Raise Err_Custom;
        End If;
        --预约接收的情况下进行分诊诊室的获取
        If n_挂号状态 = 2 Then
          --获取新号别诊室
          --说明:预约的情况下，不需要分诊，因此不用获取就诊诊室
          --     接收的情况下,需要进行分诊,因此需要获取接诊诊室
          --获取分诊方式
          Begin
            Select Nvl(分诊方式, 0) Into n_分诊方式 From 临床出诊记录 Where ID = 新出诊id_In;
          Exception
            When Others Then
              n_分诊方式 := 0;
          End;
        
          Begin
            If n_分诊方式 = 0 Then
              --不分诊
              v_新就诊诊室 := '';
            End If;
            If n_分诊方式 = 1 Then
              --指定分诊
              Select b.名称
              Into v_新就诊诊室
              From 临床出诊诊室记录 A, 门诊诊室 B
              Where a.诊室id = b.Id And a.记录id = 新出诊id_In;
            End If;
            If n_分诊方式 = 2 Then
              --动态分诊
              Select 门诊诊室
              Into v_新就诊诊室
              From (Select 门诊诊室, Sum(Num) As Num
                     From (Select b.名称 As 门诊诊室, 0 As Num
                            From 临床出诊诊室记录 A, 门诊诊室 B
                            Where a.记录id = 新出诊id_In
                            Union All
                            Select 诊室, Count(诊室) As Num
                            From 病人挂号记录
                            Where Nvl(执行状态, 0) = 0 And 记录性质 = 1 And 记录状态 = 1 And 发生时间 Between Trunc(Sysdate) And Sysdate And
                                  号别 = 新号别_In And 诊室 In (Select b.名称
                                                         From 临床出诊诊室记录 A, 门诊诊室 B
                                                         Where a.诊室id = b.Id And a.记录id = 新出诊id_In)
                            Group By 诊室)
                     Group By 门诊诊室
                     Order By Num)
              Where Rownum = 1;
            End If;
            If n_分诊方式 = 3 Then
              --平均分诊
              --获取当前安排下的诊室数量
              Select Count(1) Into n_分诊诊室数量 From 临床出诊诊室记录 Where 记录id = 新出诊id_In;
            
              Open c_出诊平均分诊;
              Loop
                Fetch c_出诊平均分诊
                  Into r_出诊平均分诊;
                Exit When c_出诊平均分诊%NotFound;
                n_Index := n_Index + 1;
                --找到了对应的分诊诊室,需要修改下一个诊室的当前分配为1(代表该诊室是下一次的分诊诊室)
                If n_是否找到分诊诊室 = 1 Then
                  Update 临床出诊诊室记录
                  Set 当前分配 = 1
                  Where 记录id = r_出诊平均分诊.记录id And 诊室id = r_出诊平均分诊.诊室id;
                  Exit;
                End If;
              
                If Nvl(r_平均分诊.当前分配, 0) = 1 Then
                  v_新就诊诊室 := r_平均分诊.门诊诊室;
                  Update 临床出诊诊室记录
                  Set 当前分配 = 0
                  Where 记录id = r_出诊平均分诊.记录id And 诊室id = r_出诊平均分诊.诊室id;
                  n_是否找到分诊诊室 := 1;
                End If;
              
                If n_分诊诊室数量 = 1 And n_是否找到分诊诊室 = 1 Then
                  n_是否找到分诊诊室 := 2;
                  Exit;
                End If;
                If n_分诊诊室数量 > 1 And n_是否找到分诊诊室 = 1 Then
                  --游标已经到了最后,所以需从第一条数据开始修改标识
                  If n_Index >= n_分诊诊室数量 Then
                    n_是否找到分诊诊室 := 2;
                    Exit;
                  End If;
                End If;
              End Loop;
              Close c_出诊平均分诊;
              --重置索引值
              n_Index := 0;
              --第一次分诊
              If Nvl(v_新就诊诊室, ' ') = ' ' Or v_新就诊诊室 Is Null Then
                Open c_出诊平均分诊;
                Loop
                  Fetch c_出诊平均分诊
                    Into r_出诊平均分诊;
                  Exit When c_出诊平均分诊%NotFound;
                  n_Index := n_Index + 1;
                
                  If n_是否找到分诊诊室 = 1 Then
                    Update 临床出诊诊室记录
                    Set 当前分配 = 1
                    Where 记录id = r_出诊平均分诊.记录id And 诊室id = r_出诊平均分诊.诊室id;
                    Exit;
                  End If;
                
                  Update 临床出诊诊室记录
                  Set 当前分配 = 0
                  Where 记录id = r_出诊平均分诊.记录id And 诊室id = r_出诊平均分诊.诊室id;
                  v_新就诊诊室 := r_出诊平均分诊.诊室id;
                
                  n_是否找到分诊诊室 := 1;
                  If n_分诊诊室数量 = 1 And n_是否找到分诊诊室 = 1 Then
                    n_是否找到分诊诊室 := 2;
                    Exit;
                  End If;
                
                  If n_分诊诊室数量 > 1 And n_是否找到分诊诊室 = 1 Then
                    --游标已经到了最后,所以需从第一条数据开始修改标识
                    If n_Index >= n_分诊诊室数量 Then
                      n_是否找到分诊诊室 := 2;
                      Exit;
                    End If;
                  End If;
                
                End Loop;
                Close c_出诊平均分诊;
              End If;
            
              If n_是否找到分诊诊室 = 2 Then
                Open c_出诊平均分诊;
                Loop
                  Fetch c_出诊平均分诊
                    Into r_出诊平均分诊;
                  Exit When c_出诊平均分诊%NotFound;
                  Update 临床出诊诊室记录
                  Set 当前分配 = 1
                  Where 记录id = r_出诊平均分诊.记录id And 诊室id = r_出诊平均分诊.诊室id;
                  Exit;
                End Loop;
                Close c_出诊平均分诊;
              End If;
            End If;
          Exception
            When Others Then
              v_新就诊诊室 := '';
          End;
        End If;
      
        --更新病人信息的就诊诊室和状态
        Update 病人信息
        Set 就诊诊室 =
             (Select 名称 From 门诊诊室 Where ID = v_新就诊诊室), 就诊状态 = 1
        Where 病人id = n_病人id And 就诊状态 In (1, 2);
      
        --打开游标
        Open c_Bill(v_No);
        Loop
          Fetch c_Bill
            Into r_Bill;
          Exit When c_Bill%NotFound;
          If r_Bill.序号 = 1 Then
            --需要确定是否预约挂号
            --1.如果是预约挂号产生的挂号记录,则需要减已约数
            --2.如果是预约挂号为接收的挂号记录,则需要减已挂数和其中已接收数
            --3.如果是正常挂号,则只减已挂数
            --恢复以前的挂号汇总
            Update 病人挂号汇总
            Set 已挂数 = Nvl(已挂数, 0) + Decode(n_挂号状态, 0, -1, 2, -1, 0), 其中已接收 = Nvl(其中已接收, 0) + Decode(n_挂号状态, 2, -1, 0),
                已约数 = Nvl(已约数, 0) + Decode(n_挂号状态, 0, 0, -1)
            Where 日期 = Trunc(r_Bill.发生时间) And Nvl(医生id, 0) = Nvl(原医生id_In, 0) And Nvl(医生姓名, '-') = Nvl(原医生姓名_In, '-') And
                  Nvl(科室id, 0) = Nvl(r_Bill.执行部门id, 0) And Nvl(项目id, 0) = Nvl(r_Bill.收费细目id, 0) And
                  (号码 = r_Bill.计算单位 Or 号码 Is Null);
            If Sql%RowCount = 0 Then
              Insert Into 病人挂号汇总
                (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 已约数, 其中已接收)
              Values
                (Trunc(r_Bill.发生时间), r_Bill.执行部门id, r_Bill.收费细目id, 原医生姓名_In, Decode(原医生id_In, 0, Null, 原医生id_In),
                 r_Bill.计算单位, 0, 0, 0);
            End If;
          
            Update 临床出诊记录
            Set 已挂数 = Nvl(已挂数, 0) + Decode(n_挂号状态, 0, -1, 2, -1, 0), 其中已接收 = Nvl(其中已接收, 0) + Decode(n_挂号状态, 2, -1, 0),
                已约数 = Nvl(已约数, 0) + Decode(n_挂号状态, 0, 0, -1)
            Where ID = 原出诊id_In;
          
            ----然后再更新挂号汇总
            Update 病人挂号汇总
            Set 已挂数 = Nvl(已挂数, 0) + Decode(n_挂号状态, 0, 1, 2, 1, 0), 其中已接收 = Nvl(其中已接收, 0) + Decode(n_挂号状态, 2, 1, 0),
                已约数 = Nvl(已约数, 0) + Decode(n_挂号状态, 0, 0, 1)
            Where 日期 = Trunc(r_Bill.发生时间) And Nvl(医生id, 0) = Nvl(新医生id_In, 0) And Nvl(医生姓名, '-') = Nvl(新医生姓名_In, '-') And
                  Nvl(科室id, 0) = 新科室id_In And Nvl(项目id, 0) = Nvl(r_Bill.收费细目id, 0) And (号码 = 新号别_In Or 号码 Is Null);
            If Sql%RowCount = 0 Then
              Insert Into 病人挂号汇总
                (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 已约数, 其中已接收)
              Values
                (Trunc(r_Bill.发生时间), 新科室id_In, r_Bill.收费细目id, 新医生姓名_In, Decode(新医生id_In, 0, Null, 新医生id_In), 新号别_In,
                 Decode(n_挂号状态, 0, 1, 2, 1, 0), Decode(n_挂号状态, 0, 0, 1), Decode(n_挂号状态, 2, 1, 0));
            End If;
            Update 临床出诊记录
            Set 已挂数 = Nvl(已挂数, 0) + Decode(n_挂号状态, 0, 1, 2, 1, 0), 其中已接收 = Nvl(其中已接收, 0) + Decode(n_挂号状态, 2, 1, 0),
                已约数 = Nvl(已约数, 0) + Decode(n_挂号状态, 0, 0, 1)
            Where ID = 新出诊id_In;
          End If;
        
          ---更新挂号记录
          If n_挂号状态 = 1 Then
            --预约
            Update 门诊费用记录
            Set 执行部门id = 新科室id_In, 病人科室id = 新科室id_In, 计算单位 = 新号别_In, 发药窗口 = n_原序号,
                --病人病区id = 科室id_In,
                执行人 = 新医生姓名_In, 执行状态 = 0, 执行时间 = Null
            Where ID = r_Bill.Id;
          Else
            --挂号或接收
            Update 门诊费用记录
            Set 执行部门id = 新科室id_In, 病人科室id = 新科室id_In, 计算单位 = 新号别_In, 发药窗口 = v_新就诊诊室,
                --病人病区id = 科室id_In,
                执行人 = 新医生姓名_In, 执行状态 = 0, 执行时间 = Null
            Where ID = r_Bill.Id;
          End If;
        
          --更新病人挂号记录
          If r_Bill.序号 = 1 Then
            v_Temp := Zl_Identity(1);
            Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_Temp From Dual;
            Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into v_操作员编号 From Dual;
            Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_操作员姓名 From Dual;
            Begin
              Select ID Into n_医生id From 人员表 Where 姓名 = 新医生姓名_In And Rownum < 2;
            Exception
              When Others Then
                n_医生id := Null;
            End;
            Select 就诊变动记录_Id.Nextval Into n_变动id From Dual;
            b_Message.Zlhis_Regist_005(r_Bill.No, n_变动id, 2);
            Zl_就诊变动记录_Insert(r_Bill.No, 1, '批量换号', v_操作员姓名, v_操作员编号, 新号别_In, 新科室id_In, Null, n_医生id, 新医生姓名_In, v_新就诊诊室,
                             n_原序号, Null, n_变动id);
            --修改队列信息
            Update 排队叫号队列
            Set 医生姓名 = 新医生姓名_In, 诊室 = v_新就诊诊室
            Where 业务id = n_业务id And 业务类型 = 0;
          
            Update 病人挂号记录
            Set 执行部门id = 新科室id_In, 号别 = 新号别_In, 诊室 = v_新就诊诊室, 执行人 = 新医生姓名_In, 执行状态 = 0, 执行时间 = Null, 号序 = n_原序号,
                出诊记录id = 新出诊id_In
            Where NO = r_Bill.No;
            --修改挂号序号状态
            If n_原序号 Is Not Null Then
              --1.恢复以前挂号序号状态
              Update 临床出诊序号控制
              Set 挂号状态 = 0, 操作员姓名 = Null, 工作站ip = Null, 工作站名称 = Null
              Where 记录id = 原出诊id_In And 序号 = n_原序号;
              --2.新增换号后挂号序号状态
              Update 临床出诊序号控制
              Set 挂号状态 = Decode(n_挂号状态, 1, 2, 1), 操作员姓名 = v_操作员姓名, 工作站ip = Null, 工作站名称 = Null
              Where 记录id = 新出诊id_In And 序号 = n_原序号;
            End If;
          End If;
        End Loop;
        Close c_Bill;
      End Loop;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人挂号记录_批量换号;
/



Create Or Replace Procedure Zl_病人接诊完成
(
  病人id_In 病人信息.病人id%Type,
  No_In     病人挂号记录.No%Type,
  诊室_In   病人挂号记录.诊室%Type := Null,
  执行人_In 病人挂号记录.执行人%Type := Null,
  摘要_In   病人挂号记录.摘要%Type := Null,
  护士_In   病人挂号记录.附加标志%Type := Null
) As
  v_挂号id     病人挂号记录.Id%Type;
  v_执行部门id 病人挂号记录.执行部门id%Type;
  v_接诊时间   病人挂号记录.执行时间%Type;
  v_完成时间   病人挂号记录.执行时间%Type;
  v_类别       Varchar2(100);
  n_诊室id     门诊诊室.Id%Type;
Begin
  v_完成时间 := Sysdate;

  Update 病人信息 Set 就诊状态 = 0 Where 病人id = 病人id_In And 就诊状态 In (1, 2); --1-等待就诊,2-正在就诊;
  Begin
    Select ID Into n_诊室id From 门诊诊室 Where 名称 = 诊室_In;
  Exception
    When Others Then
      Null;
  End;
  --执行时间保持了挂号记录一致
  Update 门诊费用记录
  Set 执行人 = Decode(执行人_In, Null, 执行人, 执行人_In), 执行状态 = 1, 发药窗口 = 诊室_In, 结论 = Decode(摘要_In, Null, 结论, 摘要_In), 婴儿费 = 护士_In
  Where NO = No_In And 记录性质 = 4 And 记录状态 In (1, 3) And Nvl(执行状态, 0) In (0, 2);

  --病人挂号记录
  Update 病人挂号记录
  Set 执行人 = Decode(执行人_In, Null, 执行人, 执行人_In), 执行状态 = 1, 诊室 = 诊室_In, 完成时间 = v_完成时间, 摘要 = Decode(摘要_In, Null, 摘要, 摘要_In),
      附加标志 = 护士_In
  Where NO = No_In And Nvl(执行状态, 0) In (0, 2) And 记录状态 = 1 And 记录性质 = 1
  Returning ID, 执行部门id, 执行时间, Decode(复诊, 1, '复诊', Decode(急诊, 1, '急诊', '门诊')) Into v_挂号id, v_执行部门id, v_接诊时间, v_类别;

  If v_挂号id Is Not Null Then
    --并发操作时，可能本次调用没有进行Update操作，就没有返回值
    --接诊后,排队叫号更新为完成
    Update 排队叫号队列 Set 排队状态 = 4 Where 业务类型 = 0 And 业务id = v_挂号id;
  
    --病历时机处理
    Zl_电子病历时机_Insert(病人id_In, v_挂号id, 1, v_类别, v_执行部门id, 执行人_In, v_接诊时间, v_完成时间);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人接诊完成;
/



Create Or Replace Procedure Zl_病人预约挂号_Clear As
  v_预约天数     Number;
  v_操作员姓名   人员表.姓名 %Type;
  v_操作员编号   人员表.编号%Type;
  n_挂号id       病人挂号记录.Id%Type;
  n_挂号排班模式 Number(3);
  n_出诊记录id   临床出诊记录.Id%Type;
  Cursor c_Clear Is
    Select a.No, a.发生时间, b.科室id, b.项目id, b.医生姓名, b.医生id, b.号码
    From 门诊费用记录 A, 挂号安排 B
    Where a.计算单位 = b.号码 And a.记录性质 = 4 And a.记录状态 = 0 And a.序号 = 1 And 登记时间 >= Sysdate - v_预约天数 And
          发生时间 < Trunc(Sysdate);
  Cursor c_出诊clear Is
    Select b.Id, a.No, a.发生时间, d.科室id, b.项目id, b.医生姓名, b.医生id, d.号码, c.号序
    From 门诊费用记录 A, 临床出诊记录 B, 病人挂号记录 C, 临床出诊号源 D
    Where a.No = c.No And c.出诊记录id = b.Id And b.号源id = d.Id And a.记录性质 = 4 And a.记录状态 = 0 And a.序号 = 1 And
          a.登记时间 >= Sysdate - v_预约天数 And a.发生时间 < Trunc(Sysdate);
Begin
  Select Zl_To_Number(Nvl(zl_GetSysParameter(66), '15')) Into v_预约天数 From Dual;
  Begin
    Select b.姓名, b.编号
    Into v_操作员姓名, v_操作员编号
    From 上机人员表 A, 人员表 B
    Where a.人员id = b.Id And a.用户名 = Upper(User);
  Exception
    When Others Then
      Null;
  End;
  n_挂号排班模式 := Zl_To_Number(Nvl(zl_GetSysParameter(253), 0));
  If n_挂号排班模式 = 0 Then
    For r_Clear In c_Clear Loop
      Update 病人挂号汇总
      Set 已约数 = Nvl(已约数, 0) - 1
      Where 日期 = Trunc(r_Clear.发生时间) And 科室id = r_Clear.科室id And 项目id = r_Clear.项目id And
            Nvl(医生姓名, '医生') = Nvl(r_Clear.医生姓名, '医生') And Nvl(医生id, 0) = Nvl(r_Clear.医生id, 0) And
            (号码 = r_Clear.号码 Or 号码 Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into 病人挂号汇总
          (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已约数)
        Values
          (Trunc(r_Clear.发生时间), r_Clear.科室id, r_Clear.项目id, r_Clear.医生姓名, Decode(r_Clear.医生id, 0, Null, r_Clear.医生id),
           r_Clear.号码, -1);
      End If;
      --删除门诊费用记录
      Delete From 门诊费用记录 Where NO = r_Clear.No And 记录性质 = 4 And 记录状态 = 0;
      Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
      Update 病人挂号记录 Set 记录状态 = 3 Where NO = r_Clear.No;
      Insert Into 病人挂号记录
        (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 预约时间, 登记时间, 发生时间, 操作员编号,
         操作员姓名, 复诊, 号序, 社区, 预约, 摘要, 交易流水号, 交易说明, 合作单位, 医疗付款方式)
        Select n_挂号id, NO, 记录性质, 2, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 预约时间, 登记时间, 发生时间,
               v_操作员编号, v_操作员姓名, 复诊, 号序, 社区, 预约, 摘要, 交易流水号, 交易说明, 合作单位, 医疗付款方式
        From 病人挂号记录
        Where NO = r_Clear.No And 记录状态 = 3;
    End Loop;
  Else
    --出诊表排班模式
    For r_出诊clear In c_出诊clear Loop
      Update 病人挂号汇总
      Set 已约数 = Nvl(已约数, 0) - 1
      Where 日期 = Trunc(r_出诊clear.发生时间) And 科室id = r_出诊clear.科室id And 项目id = r_出诊clear.项目id And
            Nvl(医生姓名, '医生') = Nvl(r_出诊clear.医生姓名, '医生') And Nvl(医生id, 0) = Nvl(r_出诊clear.医生id, 0) And
            (号码 = r_出诊clear.号码 Or 号码 Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into 病人挂号汇总
          (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已约数)
        Values
          (Trunc(r_出诊clear.发生时间), r_出诊clear.科室id, r_出诊clear.项目id, r_出诊clear.医生姓名,
           Decode(r_出诊clear.医生id, 0, Null, r_出诊clear.医生id), r_出诊clear.号码, -1);
      End If;
      Update 临床出诊记录 Set 已约数 = Nvl(已约数, 0) - 1 Where ID = r_出诊clear.Id;
      Update 临床出诊序号控制
      Set 挂号状态 = 0, 操作员姓名 = Null, 工作站ip = Null, 工作站名称 = Null
      Where 记录id = r_出诊clear.Id And (序号 = r_出诊clear.号序 Or 备注 = r_出诊clear.号序);
      --门诊费用记录
      Update 门诊费用记录 Set 记录状态 = 3 Where NO = r_出诊clear.No And 记录性质 = 4 And 记录状态 = 0;
      Insert Into 门诊费用记录
        (ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 记帐单id, 病人id, 医嘱序号, 门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号, 付款方式, 病人科室id, 费别,
         收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人, 开单部门id, 开单人, 发生时间, 登记时间,
         执行部门id, 执行人, 执行状态, 执行时间, 结论, 操作员编号, 操作员姓名, 结帐id, 结帐金额, 保险大类id, 保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要, 是否急诊, 缴款组id,
         费用状态, 待转出, 挂号id, 主页id)
        Select 病人费用记录_Id.Nextval, 记录性质, NO, 实际票号, 2, 序号, 从属父号, 价格父号, 记帐单id, 病人id, 医嘱序号, 门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号,
               付款方式, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, -1 * 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, -1 * 应收金额,
               -1 * 实收金额, 划价人, 开单部门id, 开单人, 发生时间, Sysdate, 执行部门id, 执行人, -1, 执行时间, 结论, v_操作员编号, v_操作员姓名, Null, Null,
               保险大类id, 保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要, 是否急诊, 缴款组id, 费用状态, 待转出, 挂号id, 主页id
        From 门诊费用记录
        Where NO = r_出诊clear.No And 记录性质 = 4 And 记录状态 = 3;
    
      Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
      Update 病人挂号记录 Set 记录状态 = 3 Where NO = r_出诊clear.No;
      Insert Into 病人挂号记录
        (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 预约时间, 登记时间, 发生时间, 操作员编号,
         操作员姓名, 复诊, 号序, 社区, 预约, 摘要, 交易流水号, 交易说明, 合作单位, 医疗付款方式)
        Select n_挂号id, NO, 记录性质, 2, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 预约时间, 登记时间, 发生时间,
               v_操作员编号, v_操作员姓名, 复诊, 号序, 社区, 预约, 摘要, 交易流水号, 交易说明, 合作单位, 医疗付款方式
        From 病人挂号记录
        Where NO = r_出诊clear.No And 记录状态 = 3;
    End Loop;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人预约挂号_Clear;
/



Create Or Replace Procedure Zl_病人预约挂号_Defer
(
  号别_In       门诊费用记录.发药窗口%Type,
  预约日期_In   门诊费用记录.发生时间%Type,
  延期日期_In   门诊费用记录.发生时间%Type,
  操作员姓名_In 挂号序号状态.操作员姓名%Type,
  记录id_In     临床出诊记录.Id%Type := Null
) As
  v_Do     Number(1);
  v_医生   挂号安排.医生姓名%Type;
  v_医生id 挂号安排.医生id%Type;
  v_天数   Number;
  n_记录id 临床出诊记录.Id%Type;
Begin
  If 记录id_In Is Null Then
    v_天数 := Trunc(延期日期_In) - Trunc(预约日期_In);
    For c_Fee In (Select Distinct NO, 发药窗口 号序, 执行部门id, 收费细目id
                  From 门诊费用记录
                  Where 记录性质 = 4 And 记录状态 = 0 And 序号 = 1 And 计算单位 = 号别_In And 发生时间 Between Trunc(预约日期_In) And
                        Trunc(预约日期_In + 1) - 1 / 24 / 60 / 60) Loop
      v_Do := 1;
      --挂号序号状态
      If Not c_Fee.号序 Is Null Then
        Begin
          Update 挂号序号状态
          Set 日期 = 日期 + v_天数, 登记时间 = Sysdate
          Where 号码 = 号别_In And Trunc(日期) = Trunc(预约日期_In) And 序号 = c_Fee.号序 And 状态 = 2 And 操作员姓名 = 操作员姓名_In;
        Exception
          --如果延期那天的序号已使用,则该预约挂号不延期
          When Others Then
            --如果有预留的,则允许直接使用
            Update 挂号序号状态
            Set 状态 = 2, 登记时间 = Sysdate
            Where 号码 = 号别_In And Trunc(日期) = Trunc(延期日期_In) And 序号 = c_Fee.号序 And 状态 = 3 And 操作员姓名 = 操作员姓名_In;
            If Sql%RowCount = 0 Then
              v_Do := 0;
            Else
              Delete 挂号序号状态
              Where 号码 = 号别_In And Trunc(日期) = Trunc(预约日期_In) And 序号 = c_Fee.号序 And 状态 = 2 And 操作员姓名 = 操作员姓名_In;
            End If;
        End;
      End If;
    
      If v_Do = 1 Then
        --预约记录
        Update 门诊费用记录
        Set 发生时间 = To_Date(To_Char(延期日期_In, 'yyyy-mm-dd') || To_Char(发生时间, ' hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
        Where 记录性质 = 4 And 记录状态 = 0 And NO = c_Fee.No;
        Update 病人挂号记录
        Set 发生时间 = To_Date(To_Char(延期日期_In, 'yyyy-mm-dd') || To_Char(发生时间, ' hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
        Where 记录性质 = 2 And 记录状态 = 1 And NO = c_Fee.No;
        --病人挂号汇总
        Begin
          Select 医生姓名, 医生id Into v_医生, v_医生id From 挂号安排 Where 号码 = 号别_In;
        Exception
          When Others Then
            Null;
        End;
        Update 病人挂号汇总
        Set 已约数 = Nvl(已约数, 0) - 1
        Where 日期 = Trunc(预约日期_In) And Nvl(科室id, 0) = c_Fee.执行部门id And Nvl(项目id, 0) = c_Fee.收费细目id And
              Nvl(医生姓名, '医生') = Nvl(v_医生, '医生') And Nvl(医生id, 0) = Nvl(v_医生id, 0) And (号码 = 号别_In Or 号码 Is Null);
      
        Update 病人挂号汇总
        Set 已约数 = Nvl(已约数, 0) + 1
        Where 日期 = Trunc(延期日期_In) And Nvl(科室id, 0) = c_Fee.执行部门id And Nvl(项目id, 0) = c_Fee.收费细目id And
              Nvl(医生姓名, '医生') = Nvl(v_医生, '医生') And Nvl(医生id, 0) = Nvl(v_医生id, 0) And (号码 = 号别_In Or 号码 Is Null);
        If Sql%RowCount = 0 Then
          Insert Into 病人挂号汇总
            (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已约数)
          Values
            (Trunc(延期日期_In), c_Fee.执行部门id, c_Fee.收费细目id, v_医生, Decode(v_医生id, 0, Null, v_医生id), 号别_In, 1);
        End If;
      End If;
    End Loop;
  Else
    --出诊表排班模式
    v_天数 := Trunc(延期日期_In) - Trunc(预约日期_In);
    For c_Fee In (Select Distinct NO, 发药窗口 号序, 执行部门id, 收费细目id
                  From 门诊费用记录
                  Where 记录性质 = 4 And 记录状态 = 0 And 序号 = 1 And 计算单位 = 号别_In And 发生时间 Between Trunc(预约日期_In) And
                        Trunc(预约日期_In + 1) - 1 / 24 / 60 / 60) Loop
      v_Do := 1;
      --挂号序号状态
      If Not c_Fee.号序 Is Null Then
        Select c.Id
        Into n_记录id
        From 临床出诊记录 A, 临床出诊号源 B, 临床出诊记录 C
        Where a.Id = 记录id_In And a.号源id = b.Id And b.Id = c.号源id And c.出诊日期 = a.出诊日期 + v_天数;
      
        Update 临床出诊序号控制
        Set 挂号状态 = 2, 操作员姓名 = 操作员姓名_In
        Where 记录id = n_记录id And (序号 = c_Fee.号序 Or 备注 = c_Fee.号序) And 挂号状态 = 0;
      
        If Sql%RowCount = 0 Then
          --如果有预留的,则允许直接使用
          Update 临床出诊序号控制
          Set 挂号状态 = 2, 操作员姓名 = 操作员姓名_In
          Where 记录id = n_记录id And (序号 = c_Fee.号序 Or 备注 = c_Fee.号序) And 挂号状态 = 3 And 操作员姓名 = 操作员姓名_In;
          If Sql%RowCount = 0 Then
            v_Do := 0;
          End If;
        End If;
      
        Update 临床出诊序号控制
        Set 挂号状态 = 0, 操作员姓名 = Null, 工作站ip = Null, 工作站名称 = Null
        Where 记录id = 记录id_In And (序号 = c_Fee.号序 Or 备注 = c_Fee.号序);
      End If;
    
      If v_Do = 1 Then
        --预约记录
        Update 门诊费用记录
        Set 发生时间 = To_Date(To_Char(延期日期_In, 'yyyy-mm-dd') || To_Char(发生时间, ' hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
        Where 记录性质 = 4 And 记录状态 = 0 And NO = c_Fee.No;
        Update 病人挂号记录
        Set 发生时间 = To_Date(To_Char(延期日期_In, 'yyyy-mm-dd') || To_Char(发生时间, ' hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
        Where 记录性质 = 2 And 记录状态 = 1 And NO = c_Fee.No;
        --病人挂号汇总
        Begin
          Select 医生姓名, 医生id Into v_医生, v_医生id From 临床出诊记录 Where ID = 记录id_In;
        Exception
          When Others Then
            Null;
        End;
        Update 临床出诊记录 Set 已约数 = Nvl(已约数, 0) - 1 Where ID = 记录id_In;
        Update 临床出诊记录 Set 已约数 = Nvl(已约数, 0) + 1 Where ID = n_记录id;
      
        Update 病人挂号汇总
        Set 已约数 = Nvl(已约数, 0) - 1
        Where 日期 = Trunc(预约日期_In) And Nvl(科室id, 0) = c_Fee.执行部门id And Nvl(项目id, 0) = c_Fee.收费细目id And
              Nvl(医生姓名, '医生') = Nvl(v_医生, '医生') And Nvl(医生id, 0) = Nvl(v_医生id, 0) And (号码 = 号别_In Or 号码 Is Null);
      
        Update 病人挂号汇总
        Set 已约数 = Nvl(已约数, 0) + 1
        Where 日期 = Trunc(延期日期_In) And Nvl(科室id, 0) = c_Fee.执行部门id And Nvl(项目id, 0) = c_Fee.收费细目id And
              Nvl(医生姓名, '医生') = Nvl(v_医生, '医生') And Nvl(医生id, 0) = Nvl(v_医生id, 0) And (号码 = 号别_In Or 号码 Is Null);
        If Sql%RowCount = 0 Then
          Insert Into 病人挂号汇总
            (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已约数)
          Values
            (Trunc(延期日期_In), c_Fee.执行部门id, c_Fee.收费细目id, v_医生, Decode(v_医生id, 0, Null, v_医生id), 号别_In, 1);
        End If;
      End If;
    End Loop;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人预约挂号_Defer;
/



Create Or Replace Procedure Zl_挂号序号状态_Update
(
  号码_In       挂号序号状态.号码%Type,
  日期_In       挂号序号状态.日期%Type,
  序号_In       挂号序号状态.序号%Type,
  状态_In       挂号序号状态.状态%Type,
  操作员姓名_In 挂号序号状态.操作员姓名%Type,
  操作_In       Number, --1-新增,0-删除
  备注_In       挂号序号状态.备注%Type := Null,
  出诊id_In     临床出诊记录.Id%Type := Null,
  预约顺序号_In 临床出诊序号控制.预约顺序号%Type := Null
) As

  v_姓名         挂号序号状态.操作员姓名%Type;
  v_状态         挂号序号状态.状态%Type;
  n_挂号排班模式 Number(3);
  v_Error        Varchar2(255);
  Err_Custom Exception;
Begin
  n_挂号排班模式 := Zl_To_Number(Nvl(zl_GetSysParameter('挂号排班模式'), 0));
  If n_挂号排班模式 = 0 Then
    If 操作_In = 1 Then
      --新增挂号序号状态
      Begin
        Select 操作员姓名, 状态
        Into v_姓名, v_状态
        From 挂号序号状态
        Where 号码 = 号码_In And 日期 = 日期_In And 序号 = 序号_In;
      Exception
        When Others Then
          Null;
      End;
    
      If v_姓名 Is Null Then
        Insert Into 挂号序号状态
          (号码, 日期, 序号, 状态, 操作员姓名, 备注, 登记时间)
        Values
          (号码_In, 日期_In, 序号_In, 状态_In, 操作员姓名_In, 备注_In, Sysdate);
      Else
        v_Error := '序号' || 序号_In || '已被操作员' || v_姓名;
        If v_状态 = 1 Then
          v_Error := v_Error || '使用';
        Elsif v_状态 = 2 Then
          v_Error := v_Error || '预约';
        Elsif v_状态 = 3 Then
          v_Error := v_Error || '预留';
        End If;
        Raise Err_Custom;
      End If;
    Else
      Begin
        Select 操作员姓名, 状态
        Into v_姓名, v_状态
        From 挂号序号状态
        Where 号码 = 号码_In And 日期 = 日期_In And 序号 = 序号_In;
      Exception
        When Others Then
          Null;
      End;
    
      If v_姓名 <> 操作员姓名_In And v_状态 = 3 Then
        --取消预留序号
        v_Error := '序号' || 序号_In || '是由操作员' || v_姓名 || '预留的,不允许取消!';
        Raise Err_Custom;
      Else
        Delete 挂号序号状态 Where 号码 = 号码_In And 日期 = 日期_In And 序号 = 序号_In;
      End If;
    End If;
  Else
    --出诊表排班模式
    If 操作_In = 1 Then
      --新增挂号序号状态
      Begin
        If 预约顺序号_In Is Null Then
          Select 操作员姓名, 挂号状态
          Into v_姓名, v_状态
          From 临床出诊序号控制
          Where 记录id = 出诊id_In And 序号 = 序号_In;
        Else
          Select 操作员姓名, 挂号状态
          Into v_姓名, v_状态
          From 临床出诊序号控制
          Where 记录id = 出诊id_In And 序号 = 序号_In And 预约顺序号 = 预约顺序号_In;
        End If;
      Exception
        When Others Then
          Null;
      End;
    
      If v_姓名 Is Null Then
        If 预约顺序号_In Is Null Then
          Update 临床出诊序号控制
          Set 挂号状态 = 状态_In, 操作员姓名 = 操作员姓名_In
          Where 记录id = 出诊id_In And 序号 = 序号_In;
        Else
          Update 临床出诊序号控制
          Set 挂号状态 = 状态_In, 操作员姓名 = 操作员姓名_In
          Where 记录id = 出诊id_In And 序号 = 序号_In And 预约顺序号 = 预约顺序号_In;
        End If;
      Else
        v_Error := '序号' || 序号_In || '已被操作员' || v_姓名;
        If v_状态 = 1 Then
          v_Error := v_Error || '使用';
        Elsif v_状态 = 2 Then
          v_Error := v_Error || '预约';
        Elsif v_状态 = 3 Then
          v_Error := v_Error || '预留';
        End If;
        Raise Err_Custom;
      End If;
    Else
      Begin
        If 预约顺序号_In Is Null Then
          Select 操作员姓名, 挂号状态
          Into v_姓名, v_状态
          From 临床出诊序号控制
          Where 记录id = 出诊id_In And 序号 = 序号_In;
        Else
          Select 操作员姓名, 挂号状态
          Into v_姓名, v_状态
          From 临床出诊序号控制
          Where 记录id = 出诊id_In And 序号 = 序号_In And 预约顺序号 = 预约顺序号_In;
        End If;
      Exception
        When Others Then
          Null;
      End;
    
      If v_姓名 <> 操作员姓名_In And v_状态 = 3 Then
        --取消预留序号
        v_Error := '序号' || 序号_In || '是由操作员' || v_姓名 || '预留的,不允许取消!';
        Raise Err_Custom;
      Else
        If 预约顺序号_In Is Null Then
          Update 临床出诊序号控制
          Set 挂号状态 = 0, 操作员姓名 = Null, 工作站ip = Null, 工作站名称 = Null
          Where 记录id = 出诊id_In And 序号 = 序号_In;
        Else
          Update 临床出诊序号控制
          Set 挂号状态 = 0, 操作员姓名 = Null, 工作站ip = Null, 工作站名称 = Null
          Where 记录id = 出诊id_In And 序号 = 序号_In And 预约顺序号 = 预约顺序号_In;
        End If;
      End If;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_挂号序号状态_Update;
/



Create Or Replace Procedure Zl_预约挂号接收_出诊_Insert
(
  No_In            门诊费用记录.No%Type,
  票据号_In        门诊费用记录.实际票号%Type,
  领用id_In        票据使用明细.领用id%Type,
  结帐id_In        门诊费用记录.结帐id%Type,
  诊室_In          门诊费用记录.发药窗口%Type,
  病人id_In        门诊费用记录.病人id%Type,
  门诊号_In        门诊费用记录.标识号%Type,
  姓名_In          门诊费用记录.姓名%Type,
  性别_In          门诊费用记录.性别%Type,
  年龄_In          门诊费用记录.年龄%Type,
  付款方式_In      门诊费用记录.付款方式%Type, --用于存放病人的医疗付款方式编号
  费别_In          门诊费用记录.费别%Type,
  结算方式_In      Varchar2, --现金的结算名称
  现金支付_In      病人预交记录.冲预交%Type, --挂号时现金支付部份金额
  预交支付_In      病人预交记录.冲预交%Type, --挂号时使用的预交金额
  个帐支付_In      病人预交记录.冲预交%Type, --挂号时个人帐户支付金额
  发生时间_In      门诊费用记录.发生时间%Type,
  号序_In          挂号序号状态.序号%Type,
  操作员编号_In    门诊费用记录.操作员编号%Type,
  操作员姓名_In    门诊费用记录.操作员姓名%Type,
  生成队列_In      Number := 0,
  登记时间_In      门诊费用记录.登记时间%Type := Null,
  卡类别id_In      病人预交记录.卡类别id%Type := Null,
  结算卡序号_In    病人预交记录.结算卡序号%Type := Null,
  卡号_In          病人预交记录.卡号%Type := Null,
  交易流水号_In    病人预交记录.交易流水号%Type := Null,
  交易说明_In      病人预交记录.交易说明%Type := Null,
  险类_In          病人挂号记录.险类%Type := Null,
  结算模式_In      Number := 0,
  记帐费用_In      Number := 0,
  冲预交病人ids_In Varchar2 := Null,
  三方调用_In      Number := 0
) As
  --该游标用于收费冲预交的可用预交列表
  --以ID排序，优先冲上次未冲完的。
  Cursor c_Deposit
  (
    v_病人id        病人信息.病人id%Type,
    v_冲预交病人ids Varchar2
  ) Is
    Select *
    From (Select a.Id, a.病人id, a.记录状态, a.预交类别, a.No, Nvl(a.金额, 0) As 金额
           From 病人预交记录 A,
                (Select NO, Sum(Nvl(a.金额, 0)) As 金额
                  From 病人预交记录 A
                  Where a.结帐id Is Null And Nvl(a.金额, 0) <> 0 And
                        a.病人id In (Select Column_Value From Table(f_Num2list(v_冲预交病人ids))) And Nvl(a.预交类别, 2) = 1
                  Group By NO
                  Having Sum(Nvl(a.金额, 0)) <> 0) B
           Where a.结帐id Is Null And Nvl(a.金额, 0) <> 0 And a.结算方式 Not In (Select 名称 From 结算方式 Where 性质 = 5) And
                 a.No = b.No And a.病人id In (Select Column_Value From Table(f_Num2list(v_冲预交病人ids))) And
                 Nvl(a.预交类别, 2) = 1
           Union All
           Select 0 As ID, Max(病人id) As 病人id, 记录状态, 预交类别, NO, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额
           From 病人预交记录
           Where 记录性质 In (1, 11) And 结帐id Is Not Null And Nvl(金额, 0) <> Nvl(冲预交, 0) And
                 病人id In (Select Column_Value From Table(f_Num2list(v_冲预交病人ids))) And Nvl(预交类别, 2) = 1 Having
            Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
           Group By 记录状态, 预交类别, NO)
    Order By Decode(病人id, Nvl(v_病人id, 0), 0, 1), ID, NO, 预交类别;

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_现金     结算方式.名称%Type;
  v_个人帐户 结算方式.名称%Type;
  v_队列名称 排队叫号队列.队列名称%Type;
  v_号别     门诊费用记录.计算单位%Type;
  v_号序     门诊费用记录.发药窗口%Type;
  v_排队号码 排队叫号队列.排队号码 %Type;
  v_预约方式 病人挂号记录.预约方式 %Type;

  n_打印id        票据打印内容.Id%Type;
  n_预交金额      病人预交记录.金额%Type;
  n_返回值        病人预交记录.金额%Type;
  v_冲预交病人ids Varchar2(4000);

  n_挂号id         病人挂号记录.Id%Type;
  n_分诊台签到排队 Number;
  n_组id           财务缴款分组.Id%Type;
  n_Count          Number(18);
  n_排队           Number;
  n_当天排队       Number;
  n_当前金额       病人预交记录.金额%Type;
  n_预交id         病人预交记录.Id%Type;
  n_消费卡id       消费卡目录.Id%Type;
  n_自制卡         Number;

  d_Date         Date;
  d_预约时间     门诊费用记录.发生时间%Type;
  d_发生时间     Date;
  d_排队时间     Date;
  n_时段         Number := 0;
  n_存在         Number := 0;
  v_结算内容     Varchar2(2000);
  v_当前结算     Varchar2(500);
  n_结算金额     病人预交记录.冲预交%Type;
  v_结算号码     病人预交记录.结算号码%Type;
  v_结算方式     病人预交记录.结算方式%Type;
  n_三方卡标志   Number(3);
  v_排队序号     排队叫号队列.排队序号%Type;
  n_结算模式     病人信息.结算模式%Type;
  n_票种         票据使用明细.票种%Type;
  v_付款方式     病人挂号记录.医疗付款方式%Type;
  n_接收模式     Number := 0;
  n_出诊记录id   病人挂号记录.出诊记录id%Type;
  n_新出诊记录id 病人挂号记录.出诊记录id%Type;
  n_号源id       临床出诊记录.号源id%Type;
  n_预约顺序号   临床出诊序号控制.预约顺序号%Type;
  v_Registtemp   Varchar2(500);
  n_检查         Number(3);
Begin
  n_组id          := Zl_Get组id(操作员姓名_In);
  v_冲预交病人ids := Nvl(冲预交病人ids_In, 病人id_In);
  n_接收模式      := Nvl(zl_GetSysParameter('预约接收模式', 1111), 0);

  --获取结算方式名称
  Begin
    Select 名称 Into v_现金 From 结算方式 Where 性质 = 1;
  Exception
    When Others Then
      v_现金 := '现金';
  End;
  Begin
    Select 名称 Into v_个人帐户 From 结算方式 Where 性质 = 3;
  Exception
    When Others Then
      v_个人帐户 := '个人帐户';
  End;
  If 登记时间_In Is Null Then
    Select Sysdate Into d_Date From Dual;
  Else
    d_Date := 登记时间_In;
  End If;

  --更新挂号序号状态
  Begin
    Select 号别, 号序, Trunc(发生时间), 发生时间, 预约方式, 出诊记录id
    Into v_号别, v_号序, d_预约时间, d_发生时间, v_预约方式, n_出诊记录id
    From 病人挂号记录
    Where 记录性质 = 2 And 记录状态 = 1 And Rownum = 1 And NO = No_In;
  Exception
    When Others Then
      v_Err_Msg := '当前预约挂号单已被其它人接收';
      Raise Err_Item;
  End;

  --判断是否分时段
  Select Nvl(是否分时段, 0), 号源id Into n_时段, n_号源id From 临床出诊记录 Where ID = n_出诊记录id;

  If n_时段 = 0 And 三方调用_In = 0 Then
    If n_接收模式 = 0 Then
      If Trunc(发生时间_In) = Trunc(Sysdate) Then
        d_发生时间 := 发生时间_In;
      Else
        d_发生时间 := Sysdate;
      End If;
    Else
      d_发生时间 := 发生时间_In;
    End If;
  Else
    If Not 发生时间_In Is Null Then
      d_发生时间 := 发生时间_In;
    End If;
  End If;
  If Not v_号序 Is Null Then
    If 号序_In Is Null Then
      Update 临床出诊序号控制 Set 挂号状态 = 0 Where (序号 = v_号序 Or 备注 = v_号序) And 记录id = n_出诊记录id;
    Else
      If Trunc(d_预约时间) <> Trunc(Sysdate) And n_接收模式 = 0 Then
        If n_时段 = 0 And 三方调用_In = 0 Then
          --提前接收或延迟接收
          Update 临床出诊序号控制 Set 挂号状态 = 0 Where 序号 = v_号序 And 记录id = n_出诊记录id;
        
          Begin
            Select ID
            Into n_新出诊记录id
            From 临床出诊记录
            Where 号源id = n_号源id And 出诊日期 = Trunc(Sysdate) And Rownum < 2;
          Exception
            When Others Then
              v_Err_Msg := '接收当天没有出诊安排,无法接收!';
              Raise Err_Item;
          End;
        
          Begin
            Select 1
            Into n_存在
            From 临床出诊序号控制
            Where 记录id = n_新出诊记录id And 序号 = v_号序 And Nvl(挂号状态, 0) = 0;
          Exception
            When Others Then
              n_存在 := 0;
          End;
        
          If n_存在 = 0 Then
            Update 临床出诊序号控制
            Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In
            Where 记录id = n_新出诊记录id And 序号 = v_号序 And Nvl(挂号状态, 0) = 0;
          Else
            --号码已被使用的情况
            Select Min(序号) Into v_号序 From 临床出诊序号控制 Where 记录id = n_新出诊记录id And Nvl(挂号状态, 0) = 0;
            If v_号序 Is Null Then
              v_Err_Msg := '接收当天没有可用序号,无法接收!';
              Raise Err_Item;
            End If;
            Update 临床出诊序号控制
            Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In
            Where 记录id = n_新出诊记录id And 序号 = v_号序 And Nvl(挂号状态, 0) = 0;
          End If;
        Else
          Begin
            Select ID
            Into n_新出诊记录id
            From 临床出诊记录
            Where 号源id = n_号源id And 出诊日期 = Trunc(Sysdate) And Rownum < 2;
          Exception
            When Others Then
              v_Err_Msg := '接收当天没有出诊安排,无法接收!';
              Raise Err_Item;
          End;
          Update 临床出诊序号控制
          Set 挂号状态 = 0, 操作员姓名 = 操作员姓名_In
          Where (序号 = v_号序 Or 备注 = v_号序) And 记录id = n_出诊记录id And Nvl(挂号状态, 0) = 2
          Returning 预约顺序号 Into n_预约顺序号;
        
          Update 临床出诊序号控制
          Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In, 预约顺序号 = n_预约顺序号
          Where 序号 = v_号序 And 记录id = n_新出诊记录id And Nvl(挂号状态, 0) = 0;
          If Sql% RowCount = 0 Then
            v_Err_Msg := '序号' || v_号序 || '已被其它人使用,请重新选择一个序号.';
            Raise Err_Item;
          End If;
        End If;
      Else
        Update 临床出诊序号控制
        Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In
        Where (序号 = v_号序 Or 备注 = v_号序) And 记录id = n_出诊记录id;
        If Sql%RowCount = 0 Then
          v_Err_Msg := '序号' || v_号序 || '已被其它人使用,请重新选择一个序号.';
          Raise Err_Item;
        End If;
      End If;
    End If;
  Else
    If Not 号序_In Is Null Then
      If Trunc(d_预约时间) <> Trunc(Sysdate) And n_接收模式 = 0 Then
        Begin
          Select ID
          Into n_新出诊记录id
          From 临床出诊记录
          Where 号源id = n_号源id And 出诊日期 = Trunc(Sysdate) And Rownum < 2;
        Exception
          When Others Then
            v_Err_Msg := '接收当天没有出诊安排,无法接收!';
            Raise Err_Item;
        End;
        Update 临床出诊序号控制
        Set 挂号状态 = 0, 操作员姓名 = 操作员姓名_In
        Where (序号 = 号序_In Or 备注 = 号序_In) And 记录id = n_出诊记录id And Nvl(挂号状态, 0) = 2
        Returning 预约顺序号 Into n_预约顺序号;
        Update 临床出诊序号控制
        Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In, 预约顺序号 = n_预约顺序号
        Where 序号 = 号序_In And 记录id = n_新出诊记录id And Nvl(挂号状态, 0) = 0;
        If Sql%RowCount = 0 Then
          v_Err_Msg := '序号' || 号序_In || '已被其它人使用,请重新选择一个序号.';
          Raise Err_Item;
        End If;
      Else
        Update 临床出诊序号控制
        Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In
        Where (序号 = 号序_In Or 备注 = 号序_In) And 记录id = n_出诊记录id;
      
      End If;
      v_号序 := 号序_In;
    Else
      v_号序 := Null;
    End If;
  End If;

  --更新门诊费用记录
  Update 门诊费用记录
  Set 记录状态 = 1, 实际票号 = Decode(Nvl(记帐费用_In, 0), 1, Null, 票据号_In), 结帐id = Decode(Nvl(记帐费用_In, 0), 1, Null, 结帐id_In),
      结帐金额 = Decode(Nvl(记帐费用_In, 0), 1, Null, 实收金额), 发药窗口 = 诊室_In, 病人id = 病人id_In, 标识号 = 门诊号_In, 姓名 = 姓名_In, 年龄 = 年龄_In,
      性别 = 性别_In, 付款方式 = 付款方式_In, 费别 = 费别_In, 发生时间 = d_发生时间, 登记时间 = d_Date, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In,
      缴款组id = n_组id, 记帐费用 = Decode(Nvl(记帐费用_In, 0), 1, 1, 0)
  Where 记录性质 = 4 And 记录状态 = 0 And NO = No_In;

  v_Registtemp := zl_GetSysParameter('挂号排班模式');
  If Substr(v_Registtemp, 1, 1) = 1 Then
    Begin
      If To_Date(Substr(v_Registtemp, 3), 'yyyy-mm-dd hh24:mi:ss') > d_发生时间 Then
        v_Err_Msg := '接收时间' || To_Char(d_发生时间, 'yyyy-mm-dd hh24:mi:ss') || '未启用出诊表排班模式,目前无法接收!';
        Raise Err_Item;
      End If;
    Exception
      When Others Then
        Null;
    End;
    Begin
      Select 1
      Into n_检查
      From 临床出诊记录
      Where ID = Nvl(n_新出诊记录id, n_出诊记录id) And d_发生时间 Between 停诊开始时间 And 停诊终止时间;
    Exception
      When Others Then
        n_检查 := 0;
    End;
    If n_检查 = 1 Then
      v_Err_Msg := '接收时间' || To_Char(d_发生时间, 'yyyy-mm-dd hh24:mi:ss') || '的安排已经被停诊,无法接收!';
      Raise Err_Item;
    End If;
  End If;

  --病人挂号记录
  Update 病人挂号记录
  Set 接收人 = 操作员姓名_In, 接收时间 = d_Date, 记录性质 = 1, 病人id = 病人id_In, 门诊号 = 门诊号_In, 发生时间 = d_发生时间, 姓名 = 姓名_In, 性别 = 性别_In,
      年龄 = 年龄_In, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In, 险类 = Decode(Nvl(险类_In, 0), 0, Null, 险类_In), 号序 = v_号序, 诊室 = 诊室_In,
      出诊记录id = Nvl(n_新出诊记录id, n_出诊记录id)
  Where 记录状态 = 1 And NO = No_In And 记录性质 = 2
  Returning ID Into n_挂号id;
  If Sql%NotFound Then
    Begin
      Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
      Begin
        Select 名称 Into v_付款方式 From 医疗付款方式 Where 编码 = 付款方式_In And Rownum < 2;
      Exception
        When Others Then
          v_付款方式 := Null;
      End;
      Insert Into 病人挂号记录
        (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 操作员编号, 操作员姓名,
         摘要, 号序, 预约, 预约方式, 接收人, 接收时间, 预约时间, 险类, 医疗付款方式, 出诊记录id)
        Select n_挂号id, No_In, 1, 1, 病人id_In, 门诊号_In, 姓名_In, 性别_In, 年龄_In, 计算单位, 加班标志, 诊室_In, Null, 执行部门id, 执行人, 0, Null,
               登记时间, 发生时间, 操作员编号, 操作员姓名, 摘要, v_号序, 1, Substr(结论, 1, 10) As 预约方式, 操作员姓名_In, Nvl(登记时间_In, Sysdate), 发生时间,
               Decode(Nvl(险类_In, 0), 0, Null, 险类_In), v_付款方式, Nvl(n_新出诊记录id, n_出诊记录id)
        From 门诊费用记录
        Where 记录性质 = 4 And 记录状态 = 1 And Rownum = 1 And NO = No_In;
    Exception
      When Others Then
        v_Err_Msg := '由于并发原因,单据号为【' || No_In || '】的病人' || 姓名_In || '已经被接收';
        Raise Err_Item;
    End;
  End If;

  --0-不产生队列;1-按医生或分诊台排队;2-先分诊,后医生站
  If Nvl(生成队列_In, 0) <> 0 Then
    n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113));
    If Nvl(n_分诊台签到排队, 0) = 0 Then
      For v_挂号 In (Select ID, 姓名, 诊室, 执行人, 执行部门id, 发生时间, 号别, 号序 From 病人挂号记录 Where NO = No_In) Loop
      
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
          --产生队列
          --按”执行部门”产生队列
          n_挂号id   := v_挂号.Id;
          v_队列名称 := v_挂号.执行部门id;
          v_排队号码 := Zlgetnextqueue(v_挂号.执行部门id, n_挂号id, v_挂号.号别 || '|' || v_挂号.号序);
          v_排队序号 := Zlgetsequencenum(0, n_挂号id, 0);
        
          --挂号id_In,号码_In,号序_In,缺省日期_In,扩展_In(暂无用)
          d_排队时间 := Zl_Get_Queuedate(n_挂号id, v_挂号.号别, v_挂号.号序, v_挂号.发生时间);
          --   队列名称_In , 业务类型_In, 业务id_In,科室id_In,排队号码_In,排队标记_In,患者姓名_In,病人ID_IN, 诊室_In, 医生姓名_In,
          Zl_排队叫号队列_Insert(v_队列名称, 0, n_挂号id, v_挂号.执行部门id, v_排队号码, Null, 姓名_In, 病人id_In, v_挂号.诊室, v_挂号.执行人, d_排队时间,
                           v_预约方式, Null, v_排队序号);
        Elsif Nvl(n_当天排队, 0) = 1 Then
          --更新队列号
          v_排队号码 := Zlgetnextqueue(v_挂号.执行部门id, v_挂号.Id, v_挂号.号别 || '|' || Nvl(v_挂号.号序, 0));
          v_排队序号 := Zlgetsequencenum(0, v_挂号.Id, 1);
          --新队列名称_IN, 业务类型_In, 业务id_In , 科室id_In , 患者姓名_In , 诊室_In, 医生姓名_In ,排队号码_In
          Zl_排队叫号队列_Update(v_挂号.执行部门id, 0, v_挂号.Id, v_挂号.执行部门id, v_挂号.姓名, v_挂号.诊室, v_挂号.执行人, v_排队号码, v_排队序号);
        
        Else
          --新队列名称_IN, 业务类型_In, 业务id_In , 科室id_In , 患者姓名_In , 诊室_In, 医生姓名_In ,排队号码_In
          Zl_排队叫号队列_Update(v_挂号.执行部门id, 0, v_挂号.Id, v_挂号.执行部门id, v_挂号.姓名, v_挂号.诊室, v_挂号.执行人);
        End If;
      End Loop;
    End If;
  End If;

  --汇总结算到病人预交记录
  If Nvl(记帐费用_In, 0) = 0 Then
    If Nvl(现金支付_In, 0) = 0 And Nvl(个帐支付_In, 0) = 0 And Nvl(预交支付_In, 0) = 0 Then
      Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      Insert Into 病人预交记录
        (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
         结算性质)
      Values
        (n_预交id, 4, 1, No_In, Decode(病人id_In, 0, Null, 病人id_In), v_现金, 0, 登记时间_In, 操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费',
         n_组id, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, Null, 4);
    End If;
    If Nvl(现金支付_In, 0) <> 0 Then
      v_结算内容 := 结算方式_In || '|'; --以空格分开以|结尾,没有结算号码的
      While v_结算内容 Is Not Null Loop
        v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '|') - 1);
        v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
      
        v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1));
      
        v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        v_结算号码 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
      
        v_当前结算   := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        n_三方卡标志 := To_Number(v_当前结算);
      
        If n_三方卡标志 = 0 Then
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明,
             合作单位, 结算性质, 结算号码)
          Values
            (n_预交id, 4, 1, No_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(v_结算方式, v_现金), Nvl(n_结算金额, 0), 登记时间_In,
             操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费', n_组id, Null, Null, Null, Null, Null, Null, 4, v_结算号码);
        Else
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明,
             合作单位, 结算性质, 结算号码)
          Values
            (n_预交id, 4, 1, No_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(v_结算方式, v_现金), Nvl(n_结算金额, 0), 登记时间_In,
             操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费', n_组id, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, Null, 4, v_结算号码);
          If Nvl(结算卡序号_In, 0) <> 0 Then
            n_消费卡id := Null;
            Begin
              Select Nvl(自制卡, 0), 1 Into n_自制卡, n_Count From 卡消费接口目录 Where 编号 = 结算卡序号_In;
            Exception
              When Others Then
                n_Count := 0;
            End;
            If n_Count = 0 Then
              v_Err_Msg := '没有发现原结算卡的相应类别,不能继续操作！';
              Raise Err_Item;
            End If;
            If n_自制卡 = 1 Then
              Select ID
              Into n_消费卡id
              From 消费卡目录
              Where 接口编号 = 结算卡序号_In And 卡号 = 卡号_In And
                    序号 = (Select Max(序号) From 消费卡目录 Where 接口编号 = 结算卡序号_In And 卡号 = 卡号_In);
            End If;
            Zl_病人卡结算记录_Insert(结算卡序号_In, n_消费卡id, 结算方式_In, 现金支付_In, 卡号_In, Null, 登记时间_In, Null, 结帐id_In, n_预交id);
          End If;
        End If;
      
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) + n_结算金额
        Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = Nvl(v_结算方式, v_现金)
        Returning 余额 Into n_返回值;
      
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额
            (收款员, 结算方式, 性质, 余额)
          Values
            (操作员姓名_In, Nvl(v_结算方式, v_现金), 1, n_结算金额);
          n_返回值 := n_结算金额;
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 人员缴款余额
          Where 收款员 = 操作员姓名_In And 结算方式 = Nvl(v_结算方式, v_现金) And 性质 = 1 And Nvl(余额, 0) = 0;
        End If;
      
        v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '|') + 1);
      End Loop;
    End If;
  End If;

  --对于就诊卡通过预交金挂号
  If Nvl(预交支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 Then
    n_预交金额 := 预交支付_In;
    For r_Deposit In c_Deposit(病人id_In, v_冲预交病人ids) Loop
      n_当前金额 := Case
                  When r_Deposit.金额 - n_预交金额 < 0 Then
                   r_Deposit.金额
                  Else
                   n_预交金额
                End;
      If r_Deposit.Id <> 0 Then
        --第一次冲预交(填上结帐ID,金额为0)
        Update 病人预交记录 Set 冲预交 = 0, 结帐id = 结帐id_In, 结算性质 = 4 Where ID = r_Deposit.Id;
      End If;
      --冲上次剩余额
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
         结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 预交类别, 结算序号, 结算性质)
        Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 登记时间_In,
               操作员姓名_In, 操作员编号_In, n_当前金额, 结帐id_In, n_组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 预交类别, 结帐id_In, 4
        From 病人预交记录
        Where NO = r_Deposit.No And 记录状态 = r_Deposit.记录状态 And 记录性质 In (1, 11) And Rownum = 1;
    
      --更新病人预交余额
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) - n_当前金额
      Where 病人id = r_Deposit.病人id And 性质 = 1 And 类型 = Nvl(r_Deposit.预交类别, 2)
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 类型, 预交余额, 性质)
        Values
          (r_Deposit.病人id, Nvl(r_Deposit.预交类别, 2), -1 * n_当前金额, 1);
        n_返回值 := -1 * n_当前金额;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 病人余额
        Where 病人id = r_Deposit.病人id And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
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

  --对于医保挂号
  If Nvl(个帐支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 Then
    Insert Into 病人预交记录
      (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
       预交类别, 结算序号, 结算性质)
    Values
      (病人预交记录_Id.Nextval, 4, 1, No_In, 病人id_In, v_个人帐户, 个帐支付_In, d_Date, 操作员编号_In, 操作员姓名_In, 结帐id_In, '医保挂号', n_组id,
       卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, Null, Null, 结帐id_In, 4);
  End If;

  --相关汇总表的处理
  --人员缴款余额
  If Nvl(个帐支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 Then
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) + 个帐支付_In
    Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = v_个人帐户
    Returning 余额 Into n_返回值;
  
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (操作员姓名_In, v_个人帐户, 1, 个帐支付_In);
      n_返回值 := 个帐支付_In;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 人员缴款余额 Where 收款员 = 操作员姓名_In And 性质 = 1 And Nvl(余额, 0) = 0;
    End If;
  End If;

  --处理票据使用情况
  If 票据号_In Is Not Null And Nvl(记帐费用_In, 0) = 0 Then
    Select 票据打印内容_Id.Nextval Into n_打印id From Dual;
  
    --当前票据的票种
    Select 票种 Into n_票种 From 票据领用记录 Where ID = Nvl(领用id_In, 0);
    --发出票据
    Insert Into 票据打印内容 (ID, 数据性质, NO) Values (n_打印id, 4, No_In);
  
    Insert Into 票据使用明细
      (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
    Values
      (票据使用明细_Id.Nextval, n_票种, 票据号_In, 1, 1, 领用id_In, n_打印id, d_Date, 操作员姓名_In);
  
    --状态改动
    Update 票据领用记录
    Set 当前号码 = 票据号_In, 剩余数量 = Decode(Sign(剩余数量 - 1), -1, 0, 剩余数量 - 1), 使用时间 = d_Date
    Where ID = Nvl(领用id_In, 0);
  End If;

  If Nvl(记帐费用_In, 0) = 1 Then
    --记帐
    If Nvl(病人id_In, 0) = 0 Then
      v_Err_Msg := '要针对病人的挂号费进行记帐，必须是建档病人才能记帐挂号。';
      Raise Err_Item;
    End If;
    For c_费用 In (Select 实收金额, 病人科室id, 开单部门id, 执行部门id, 收入项目id
                 From 门诊费用记录
                 Where 记录性质 = 4 And 记录状态 = 1 And NO = No_In And Nvl(记帐费用, 0) = 1) Loop
      --病人余额
      Update 病人余额
      Set 费用余额 = Nvl(费用余额, 0) + Nvl(c_费用.实收金额, 0)
      Where 病人id = Nvl(病人id_In, 0) And 性质 = 1 And 类型 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 性质, 类型, 费用余额, 预交余额)
        Values
          (病人id_In, 1, 1, Nvl(c_费用.实收金额, 0), 0);
      End If;
    
      --病人未结费用
      Update 病人未结费用
      Set 金额 = Nvl(金额, 0) + Nvl(c_费用.实收金额, 0)
      Where 病人id = 病人id_In And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(病人科室id, 0) = Nvl(c_费用.病人科室id, 0) And
            Nvl(开单部门id, 0) = Nvl(c_费用.开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(c_费用.执行部门id, 0) And 收入项目id + 0 = c_费用.收入项目id And
            来源途径 + 0 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人未结费用
          (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
        Values
          (病人id_In, Null, Null, c_费用.病人科室id, c_费用.开单部门id, c_费用.执行部门id, c_费用.收入项目id, 1, Nvl(c_费用.实收金额, 0));
      End If;
    End Loop;
  End If;
  If Nvl(病人id_In, 0) <> 0 Then
    n_结算模式 := 0;
    Update 病人信息
    Set 就诊时间 = d_发生时间, 就诊状态 = 1, 就诊诊室 = 诊室_In
    Where 病人id = 病人id_In
    Returning Nvl(结算模式, 0) Into n_结算模式;
    --取参数:
    If Nvl(n_结算模式, 0) <> Nvl(结算模式_In, 0) Then
      --结算模式的确定
      If n_结算模式 = 1 And Nvl(结算模式_In, 0) = 0 Then
        --病人已经是"先诊疗后结算的",本次是"先结算后诊疗的",则检查是否存在未结数据
        Select Count(1)
        Into n_Count
        From 病人未结费用
        Where 病人id = 病人id_In And (来源途径 In (1, 4) Or 来源途径 = 3 And Nvl(主页id, 0) = 0) And Nvl(金额, 0) <> 0 And Rownum < 2;
        If Nvl(n_Count, 0) <> 0 Then
          --存在未结算数据，必须先结算后才允许执行
          v_Err_Msg := '当前病人的就诊模式为先诊疗后结算且存在未结费用，不允许调整该病人的就诊模式,你可以先对未结费用结帐后再挂号或不调整病人的就诊模式!';
          Raise Err_Item;
        End If;
        --检查
        --未发生医嘱业务的（即当时就挂号的,需要保证同一次的就诊模式是一至的(程序已经检查，不用再处理)
      End If;
      Update 病人信息 Set 结算模式 = 结算模式_In Where 病人id = 病人id_In;
    End If;
  End If;

  --病人担保信息
  If 病人id_In Is Not Null Then
    Update 病人信息
    Set 担保人 = Null, 担保额 = Null, 担保性质 = Null
    Where 病人id = 病人id_In And Exists (Select 1
           From 病人担保记录
           Where 病人id = 病人id_In And 主页id Is Not Null And
                 登记时间 = (Select Max(登记时间) From 病人担保记录 Where 病人id = 病人id_In));
    If Sql%RowCount > 0 Then
      Update 病人担保记录
      Set 到期时间 = d_Date
      Where 病人id = 病人id_In And 主页id Is Not Null And Nvl(到期时间, d_Date) > d_Date;
    End If;
  End If;
  --消息推送
  Begin
    Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
      Using 1, n_挂号id;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_预约挂号接收_出诊_Insert;
/




Create Or Replace Procedure Zl_病人预约登记_Insert
(
  病人id_In     病人信息.病人id%Type,
  号源id_In     临床出诊号源.Id%Type,
  复诊方式_In   病人服务信息记录.复诊方式%Type,
  数量_In       病人服务信息记录.数量%Type,
  说明_In       病人服务信息记录.通知原因%Type,
  提醒时间_In   病人服务信息记录.开始时间%Type,
  提醒天数_In   Number,
  操作员姓名_In 病人挂号记录.操作员姓名%Type,
  操作员编号_In 病人挂号记录.操作员编号%Type
) As
  d_开始时间 病人服务信息记录.开始时间%Type;
  d_结束时间 病人服务信息记录.终止时间%Type;
  v_号码     病人服务信息记录.号码%Type;
  n_科室id   病人服务信息记录.科室id%Type;
  n_项目id   病人服务信息记录.项目id%Type;
  n_医生id   病人服务信息记录.医生id%Type;
  v_医生姓名 病人服务信息记录.医生姓名%Type;
  v_Err_Msg  Varchar2(200);
  Err_Item Exception;
Begin
  Begin
    Select 号码, 科室id, 项目id, 医生id, 医生姓名
    Into v_号码, n_科室id, n_项目id, n_医生id, v_医生姓名
    From 临床出诊号源
    Where ID = 号源id_In;
  Exception
    When Others Then
      v_Err_Msg := '没有找到号源信息,预约登记失败';
      Raise Err_Item;
  End;
  d_开始时间 := Trunc(提醒时间_In);
  d_结束时间 := d_开始时间 + Nvl(提醒天数_In, 1);
  Insert Into 病人服务信息记录
    (ID, 通知类型, 号源id, 号码, 科室id, 项目id, 医生id, 医生姓名, 病人id, 复诊方式, 数量, 开始时间, 终止时间, 通知原因, 登记人, 登记时间)
  Values
    (病人服务信息记录_Id.Nextval, 3, 号源id_In, v_号码, n_科室id, n_项目id, n_医生id, v_医生姓名, 病人id_In, 复诊方式_In, 数量_In, d_开始时间, d_结束时间,
     说明_In, 操作员姓名_In, Sysdate);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人预约登记_Insert;
/


CREATE OR REPLACE Function Zl_预约方式_Check
(
  记录id_In   临床出诊记录.Id%Type,
  序号_In     临床出诊序号控制.序号%Type,
  预约方式_In 预约方式.名称%Type
) Return Number Is
  --功能:预约时检查相应的预约方式是否可用
  --返回:0-检查未通过,1-检查通过
  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  n_控制方式 临床出诊挂号控制记录.控制方式%Type;
  n_限约数   临床出诊记录.限约数%Type;
  n_数量     临床出诊挂号控制记录.数量%Type;
  n_已约数   临床出诊记录.已约数%Type;
  n_分时段   临床出诊记录.是否分时段%Type;
  n_序号控制 临床出诊记录.是否序号控制%Type;
Begin
  Begin
    Select 控制方式
    Into n_控制方式
    From 临床出诊挂号控制记录
    Where 类型 = 2 And 性质 = 1 And 记录id = 记录id_In And Rownum < 2;
  Exception
    When Others Then
      Return 1;
  End;
  If n_控制方式 = 0 Then
    Return 0;
  End If;
  If n_控制方式 = 1 Or n_控制方式 = 2 Then
    Select Nvl(限约数, 限号数) Into n_限约数 From 临床出诊记录 Where ID = 记录id_In;
    Select 数量
    Into n_数量
    From 临床出诊挂号控制记录
    Where 类型 = 2 And 性质 = 1 And 名称 = 预约方式_In And 记录id = 记录id_In;
    If n_控制方式 = 1 Then
      n_限约数 := Round(n_限约数 * n_数量 / 100);
    Else
      n_限约数 := n_数量;
    End If;
    Select Count(1)
    Into n_已约数
    From 病人挂号记录
    Where 出诊记录id = 记录id_In And 记录状态 = 1 And 预约方式 = 预约方式_In;
    If n_已约数 >= n_限约数 Then
      Return 0;
    End If;
  End If;
  If n_控制方式 = 3 Then
    Select 数量
    Into n_限约数
    From 临床出诊挂号控制记录
    Where 类型 = 2 And 性质 = 1 And 名称 = 预约方式_In And 记录id = 记录id_In And 序号 = 序号_In;
    Select 是否分时段, 是否序号控制 Into n_分时段, n_序号控制 From 临床出诊记录 Where ID = 记录id_In;
    If n_序号控制 = 1 Then
      Select Nvl(Max(1), 0) Into n_已约数 From 病人挂号记录 Where 出诊记录id = 记录id_In And 号序 = 序号_In;
    Else
      Select Count(1)
      Into n_已约数
      From 临床出诊序号控制 A, 病人挂号记录 B
      Where a.记录id = 记录id_In And a.预约顺序号 Is Not Null And Nvl(a.挂号状态, 0) <> 0 And a.备注 = b.号序 And b.预约方式 = 预约方式_In And
            b.记录状态 = 1;
    End If;
    If n_已约数 >= n_限约数 Then
      Return 0;
    End If;
  End If;
  If n_控制方式 = 4 Then
    Return 1;
  End If;
  Return 1;
Exception
  When Err_Item Then
    Return 0;
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_预约方式_Check;
/


Create Or Replace Procedure Zl_三方机构挂号_Insert
(
  操作方式_In      Integer,
  病人id_In        门诊费用记录.病人id%Type,
  号码_In          挂号安排.号码%Type,
  号序_In          挂号序号状态.序号%Type,
  单据号_In        门诊费用记录.No%Type,
  票据号_In        门诊费用记录.实际票号%Type,
  结算方式_In      Varchar2,
  摘要_In          门诊费用记录.摘要%Type, --预约挂号摘要信息
  发生时间_In      门诊费用记录.发生时间%Type,
  登记时间_In      门诊费用记录.登记时间%Type,
  合作单位_In      挂号合作单位.名称%Type,
  挂号金额合计_In  门诊费用记录.实收金额%Type,
  领用id_In        票据使用明细.领用id%Type,
  收费票据_In      Number := 0, --挂号是否使用收费票据
  交易流水号_In    病人预交记录.交易流水号%Type,
  交易说明_In      病人预交记录.交易说明%Type,
  预约方式_In      预约方式.名称%Type := Null,
  预交id_In        病人预交记录.Id%Type := Null,
  卡类别id_In      病人预交记录.卡类别id%Type := Null,
  加入序号状态_In  Number := 0,
  是否自助设备_In  Number := 0,
  结帐id_In        门诊费用记录.结帐id%Type := Null,
  锁定类型_In      Number := 0,
  保险结算_In      Varchar2 := Null,
  冲预交_In        Number := Null,
  支付卡号_In      病人预交记录.卡号%Type := Null,
  退号重用_In      Number := 1,
  费别_In          门诊费用记录.费别%Type := Null,
  冲预交病人ids_In Varchar2 := Null,
  机器名_In        挂号序号状态.机器名%Type := Null,
  更新年龄_In      Number := 0,
  购买病历_In      Number := 0,
  出诊记录id_In    临床出诊记录.Id%Type := Null
) As
  --功能：三方机构进行挂号(包含预约;预约挂号不扣款;预约挂号扣款)
  --入参:操作方式_IN:1-表示挂号,2-表示预约挂号不扣款,3-表示预约挂号,扣款
  --      结算方式_IN:支持多种结算方式,多种结算方式时，传入格式如下:结算方式名称1,金额,结算号码,三方卡标志|结算方式名称2,金额,结算号码,三方卡标志|...
  --      加入序号状态_In:1-表示强制加入挂号序号状态表中;否则根据启用序号或启用时段时才加入.
  --      是否自助设备_In:1-表示是医院的自助设备进行此函数的调用,自助设备调用此函数 允许加号,否则不允许
  --      锁定类型_In :0-产生正常数据 1-表示对单据进行锁定,产生未生效的单据信息;2-对锁定的记录进行解锁-生成正常数据:未生效的单据在银行扣款完成后进行解锁
  --      保险结算_IN:格式="结算方式|结算金额||....."
  --      冲预交病人ids_In:多个用逗分分离,冲预交时有效(冲预交类别与业务操作保持一致),主要是使用家属的预交款
  Err_Item Exception;
  v_Err_Msg  Varchar2(255);
  n_打印id   票据打印内容.Id%Type;
  n_返回值   病人预交记录.金额%Type;
  v_排队号码 Varchar2(20);
  v_队列名称 排队叫号队列.队列名称%Type;
  n_预交id   病人预交记录.Id%Type;
  n_挂号id   病人挂号记录.Id%Type;
  v_结算内容 Varchar2(3000);
  v_当前结算 Varchar2(150);

  v_结算方式       病人预交记录.结算方式%Type;
  n_结算金额       病人预交记录.冲预交%Type;
  n_结算合计       Number(16, 5);
  n_预交金额       病人预交记录.冲预交%Type;
  n_组id           财务缴款分组.Id%Type;
  d_排队时间       Date;
  n_锁定           Number;
  n_同科限约一个号 Number(18);
  n_病人预约科室数 Number(18);
  n_已约科室       Number(18);

  n_合作单位限制       Number(18);
  n_是否开放           Number(1);
  n_Count              Number(18);
  n_行号               Number(18);
  n_序号               病人挂号记录.号序%Type;
  n_费用id             门诊费用记录.Id%Type;
  n_价格父号           Number(18);
  n_原项目id           收费项目目录.Id%Type;
  n_原收入项目id       收费项目目录.Id%Type;
  v_诊室               病人挂号记录.诊室%Type;
  n_安排id             挂号安排.Id%Type;
  n_实收金额合计       门诊费用记录.实收金额%Type;
  n_开单部门id         门诊费用记录.开单部门id%Type;
  n_实收金额           门诊费用记录.实收金额%Type;
  n_应收金额           门诊费用记录.实收金额%Type;
  n_结帐id             病人结帐记录.Id%Type;
  v_Temp               Varchar2(500);
  n_预约时段序号       Number;
  n_预约总数           Number;
  d_时段开始时间       Date;
  v_收费项目ids        Varchar2(300);
  n_预约数量           合作单位挂号汇总.已约数%Type;
  n_号序               病人挂号记录.号序%Type;
  d_登记时间           Date;
  v_操作员编号         人员表.编号%Type;
  v_操作员姓名         人员表.姓名%Type;
  n_预约               Integer;
  v_星期               挂号安排时段.星期%Type;
  n_启用分时段         Integer;
  n_已挂数             病人挂号汇总.已挂数%Type;
  n_已约数             病人挂号汇总.已约数%Type;
  n_其中已接收         病人挂号汇总.已约数%Type;
  n_预约生成队列       Number;
  d_Date               Date;
  n_挂号序号           Number;
  v_排队标记           排队叫号队列.排队标记%Type;
  v_排队序号           排队叫号队列.排队序号%Type;
  v_机器名             挂号序号状态.机器名%Type;
  v_序号操作员         挂号序号状态.操作员姓名%Type;
  v_序号机器名         挂号序号状态.机器名%Type;
  n_序号锁定           Number := 0;
  n_病历费id           收费特定项目.收费细目id%Type;
  v_付款方式           病人挂号记录.医疗付款方式%Type;
  v_费别               门诊费用记录.费别%Type;
  n_屏蔽费别           Number(3) := 0;
  v_冲预交病人ids      Varchar2(4000);
  n_Tmp安排id          挂号安排.Id%Type;
  n_计划id             挂号安排计划.Id%Type;
  v_年龄               病人信息.年龄%Type;
  n_合作单位限数量模式 Number;
  n_挂号排班模式       Number;
  d_启用时间           Date;

  Cursor c_Pati(n_病人id 病人信息.病人id%Type) Is
    Select a.病人id, a.姓名, a.性别, a.年龄, a.住院号, a.门诊号, a.费别, a.险类, c.编码 As 付款方式
    From 病人信息 A, 医疗付款方式 C
    Where a.病人id = n_病人id And a.医疗付款方式 = c.名称(+);

  r_Pati c_Pati%RowType;

  --该游标用于收费冲预交的可用预交列表
  --以ID排序，优先冲上次未冲完的。
  Cursor c_Deposit
  (
    n_病人id        病人信息.病人id%Type,
    v_冲预交病人ids Varchar2
  ) Is
    Select *
    From (Select a.Id, a.病人id, a.记录状态, a.预交类别, a.No, Nvl(a.金额, 0) As 金额
           From 病人预交记录 A,
                (Select NO, Sum(Nvl(a.金额, 0)) As 金额
                  From 病人预交记录 A
                  Where a.结帐id Is Null And Nvl(a.金额, 0) <> 0 And
                        a.病人id In (Select Column_Value From Table(f_Num2list(v_冲预交病人ids))) And Nvl(a.预交类别, 2) = 1
                  Group By NO
                  Having Sum(Nvl(a.金额, 0)) <> 0) B
           Where a.结帐id Is Null And Nvl(a.金额, 0) <> 0 And a.结算方式 Not In (Select 名称 From 结算方式 Where 性质 = 5) And
                 a.No = b.No And a.病人id In (Select Column_Value From Table(f_Num2list(v_冲预交病人ids))) And
                 Nvl(a.预交类别, 2) = 1
           Union All
           Select 0 As ID, Max(病人id) As 病人id, 记录状态, 预交类别, NO, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额
           From 病人预交记录
           Where 记录性质 In (1, 11) And 结帐id Is Not Null And Nvl(金额, 0) <> Nvl(冲预交, 0) And
                 病人id In (Select Column_Value From Table(f_Num2list(v_冲预交病人ids))) And Nvl(预交类别, 2) = 1 Having
            Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
           Group By 记录状态, NO, 预交类别)
    Order By Decode(病人id, Nvl(n_病人id, 0), 0, 1), ID, NO;

  Cursor c_安排
  (
    v_号码        挂号安排.号码%Type,
    d_发生时间_In Date
  ) Is
    Select *
    From (With 安排时间段 As (Select 时间段
                         From (Select 时间段,
                                       To_Date(Decode(Sign(开始时间 - 终止时间), 1, '3000-01-09 ' || To_Char(开始时间, 'HH24:MI:SS'),
                                                       '3000-01-10 ' || To_Char(开始时间, 'HH24:MI:SS')), 'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                                       To_Date(Decode(Sign(开始时间 - 终止时间), 1, '3000-01-11 ' || To_Char(终止时间, 'HH24:MI:SS'),
                                                       '3000-01-10 ' || To_Char(终止时间, 'HH24:MI:SS')), 'yyyy-mm-dd hh24:mi:ss') As 终止时间,
                                       To_Date('3000-01-10 ' || To_Char(d_发生时间_In, 'HH24:MI:SS'), 'yyyy-mm-dd hh24:mi:ss') As 当前时间,
                                       To_Date('3000-01-10 ' || To_Char(开始时间, 'HH24:MI:SS'), 'yyyy-mm-dd hh24:mi:ss') As 开始时间1,
                                       To_Date('3000-01-10 ' || To_Char(终止时间, 'HH24:MI:SS'), 'yyyy-mm-dd hh24:mi:ss') As 终止时间1
                                From 时间段)
                         Where 当前时间 Between 开始时间 And 终止时间1 Or 当前时间 Between 开始时间1 And 终止时间)
           Select Distinct p.Id, p.号类, p.号码, p.科室id, b.编码 As 科室编码, b.名称 As 科室名称, p.项目id, c.编码 As 项目编码, c.名称 As 项目名称,
                           p.医生id, d.编号 As 医生编号, p.医生姓名, p.限号数, p.限约数, p.周日 As 日, p.周一 As 一, p.周二 As 二, p.周三 As 三,
                           p.周四 As 四, p.周五 As 五, p.周六 As 六, p.序号控制, p.计划id
           From (Select p.Id, p.号码, p.号类, p.科室id, p.项目id, p.医生id, p.医生姓名, b.限号数, b.限约数, Nvl(p.病案必须, 0) As 病案必须, p.周日, p.周一,
                         p.周二, p.周三, p.周四, p.周五, p.周六, p.分诊方式, p.序号控制,
                         Decode(To_Char(d_发生时间_In, 'D'), '1', p.周日, '2', p.周一, '3', p.周二, '4', p.周三, '5', p.周四, '6', p.周五,
                                 '7', p.周六, Null) As 排班, Null As 计划id
                  From 挂号安排 P, 挂号安排限制 B
                  Where p.停用日期 Is Null And p.Id = b.安排id(+) And
                        b.限制项目(+) = Decode(To_Char(d_发生时间_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四',
                                           '6', '周五', '7', '周六', Null) And
                        d_发生时间_In Between Nvl(p.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                        Nvl(p.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Not Exists
                   (Select 1
                         From 挂号安排计划
                         Where 安排id = p.Id And (d_发生时间_In Between 生效时间 And 失效时间) And 审核时间 Is Not Null) And Not Exists
                   (Select 1
                         From 挂号安排停用状态
                         Where 安排id = p.Id And d_发生时间_In Between 开始停止时间 And 结束停止时间) And p.号码 = v_号码
                  Union All
                  Select c.Id, c.号码, c.号类, c.科室id, p.项目id, p.医生id, p.医生姓名, b.限号数, b.限约数, Nvl(c.病案必须, 0) As 病案必须, p.周日, p.周一,
                         p.周二, p.周三, p.周四, p.周五, p.周六, p.分诊方式, p.序号控制,
                         Decode(To_Char(d_发生时间_In, 'D'), '1', p.周日, '2', p.周一, '3', p.周二, '4', p.周三, '5', p.周四, '6', p.周五,
                                 '7', p.周六, Null) As 排班, p.Id As 计划id
                  From 挂号安排计划 P, 挂号安排 C, 挂号计划限制 B,
                       (Select Max(a.生效时间) As 生效, 安排id
                         From 挂号安排计划 A, 挂号安排 B
                         Where a.安排id = b.Id And a.审核时间 Is Not Null And
                               发生时间_In Between Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                               Nvl(a.失效时间, To_Date('3000-01-01', 'yyyy-mm-dd')) And b.号码 = 号码_In
                         Group By 安排id) E
                  Where p.安排id = c.Id And p.Id = b.计划id(+) And p.生效时间 = e.生效 And p.安排id = e.安排id And
                        Nvl(p.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) = Nvl(e.生效, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                        b.限制项目(+) = Decode(To_Char(d_发生时间_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四',
                                           '6', '周五', '7', '周六', Null) And (d_发生时间_In Between p.生效时间 And p.失效时间) And
                        p.审核时间 Is Not Null And Not Exists
                   (Select 1
                         From 挂号安排停用状态
                         Where 安排id = c.Id And d_发生时间_In Between 开始停止时间 And 结束停止时间) And p.号码 = v_号码) P, 部门表 B, 收费项目目录 C,
                人员表 D
           Where p.科室id = b.Id And p.医生id = d.Id(+) And p.项目id = c.Id And
                 (c.撤档时间 Is Null Or c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And
                 (Nvl(p.医生id, 0) = 0 Or Exists
                  (Select 1
                   From 人员表 Q
                   Where p.医生id = q.Id And (q.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or q.撤档时间 Is Null))) And Exists
            (Select 1 From 安排时间段 Where 时间段 = p.排班))
           Order By 号码;


  r_安排 c_安排%RowType;

  Function Zl_诊室(号码_In 挂号安排.号码%Type) Return Varchar2 As
    n_分诊方式 挂号安排.分诊方式%Type;
    n_安排id   挂号安排.Id%Type;
    v_诊室     病人挂号记录.诊室%Type;
    v_Rowid    Varchar2(500);
    n_Next     Integer;
    n_First    Integer;
  Begin
  
    If 锁定类型_In = 2 Then
      --对单据进行解锁,首先检查是否存在锁定
      Select Count(Rowid) Into n_锁定 From 病人挂号记录 Where　记录状态 = 0 And NO = 单据号_In And 病人id = 病人id_In;
      If n_锁定 = 0 Then
        v_Err_Msg := '单据号为(' || 单据号_In || ')的单据,不存在或者已经被解锁!';
        Raise Err_Item;
      End If;
      Select Max(号序) Into n_号序 From 病人挂号记录 Where　记录状态 = 0 And NO = 单据号_In And 病人id = 病人id_In;
    End If;
  
    Begin
      Select ID, Nvl(分诊方式, 0) Into n_安排id, n_分诊方式 From 挂号安排 Where 号码 = 号码_In;
    Exception
      When Others Then
        n_安排id := -1;
    End;
  
    If n_安排id = -1 Then
      v_Err_Msg := '号码(' || 号码_In || ')未找到!';
      Raise Err_Item;
    End If;
    --0-不分诊、1-指定诊室、2-动态分诊、3-平均分诊,对应门诊诊室设置
    v_诊室 := Null;
    If n_分诊方式 = 1 Then
      --1-指定诊室
      Begin
        Select 门诊诊室 Into v_诊室 From 挂号安排诊室 Where 号表id = n_安排id;
      Exception
        When Others Then
          v_诊室 := Null;
      End;
    End If;
    If n_分诊方式 = 2 Then
      --2-动态分诊:该个号别当天挂号未诊数最少的诊室
      For c_诊室 In (Select 门诊诊室, Sum(Num) As Num
                   From (Select 门诊诊室, 0 As Num
                          From 挂号安排诊室
                          Where 号表id = n_安排id
                          Union All
                          Select 诊室, Count(诊室) As Num
                          From 病人挂号记录
                          Where Nvl(执行状态, 0) = 0 And 发生时间 Between Trunc(Sysdate) And Sysdate And 号别 = 号码_In And
                                诊室 In (Select 门诊诊室 From 挂号安排诊室 Where 号表id = n_安排id)
                          Group By 诊室)
                   Group By 门诊诊室
                   Order By Num) Loop
        v_诊室 := c_诊室.门诊诊室;
        Exit;
      End Loop;
    End If;
    If n_分诊方式 = 3 Then
    
      --平均分诊：当前分配=1表示下次应取的当前诊室
      n_Next  := 0;
      n_First := 1;
      For c_诊室 In (Select Rowid As Rid, 号表id, 门诊诊室, 当前分配 From 挂号安排诊室 Where 号表id = n_安排id) Loop
        If n_First = 1 Then
          v_Rowid := c_诊室.Rid;
        End If;
        If n_Next = 1 Then
          v_诊室 := c_诊室.门诊诊室;
          Update 挂号安排诊室 Set 当前分配 = 1 Where Rowid = c_诊室.Rid;
          Exit;
        End If;
        If Nvl(c_诊室.当前分配, 0) = 1 Then
          Update 挂号安排诊室 Set 当前分配 = 0 Where Rowid = c_诊室.Rid;
          n_Next := 1;
        End If;
      End Loop;
      If v_诊室 Is Null Then
        Update 挂号安排诊室 Set 当前分配 = 1 Where Rowid = v_Rowid Returning 门诊诊室 Into v_诊室;
      End If;
    End If;
  
    Return v_诊室;
  End;

  Function Zl_操作员
  (
    Type_In     Integer,
    Splitstr_In Varchar2
  ) Return Varchar2 As
    n_Step Number(18);
    v_Sub  Varchar2(1000);
    --Type_In:0-获取缺省部门ID;1-获取操作员编号;2-获取操作员姓名
    -- SplitStr:格式为:部门ID,部门名称;人员ID,人员编号,人员姓名(用Zl_Identity获取的)
  Begin
    If Type_In = 0 Then
      --缺省部门
      n_Step := Instr(Splitstr_In, ',');
      v_Sub  := Substr(Splitstr_In, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 1 Then
      --操作员编码
      n_Step := Instr(Splitstr_In, ';');
      v_Sub  := Substr(Splitstr_In, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 2 Then
      --操作员姓名
      n_Step := Instr(Splitstr_In, ';');
      v_Sub  := Substr(Splitstr_In, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      Return v_Sub;
    End If;
  End;

  Procedure Zl_三方机构挂号_出诊_Insert
  (
    记录id_In        临床出诊记录.Id%Type,
    操作方式_In      Integer,
    病人id_In        门诊费用记录.病人id%Type,
    号码_In          挂号安排.号码%Type,
    号序_In          挂号序号状态.序号%Type,
    单据号_In        门诊费用记录.No%Type,
    票据号_In        门诊费用记录.实际票号%Type,
    结算方式_In      Varchar2,
    摘要_In          门诊费用记录.摘要%Type, --预约挂号摘要信息
    发生时间_In      门诊费用记录.发生时间%Type,
    登记时间_In      门诊费用记录.登记时间%Type,
    合作单位_In      挂号合作单位.名称%Type,
    挂号金额合计_In  门诊费用记录.实收金额%Type,
    领用id_In        票据使用明细.领用id%Type,
    收费票据_In      Number := 0, --挂号是否使用收费票据
    交易流水号_In    病人预交记录.交易流水号%Type,
    交易说明_In      病人预交记录.交易说明%Type,
    预约方式_In      预约方式.名称%Type := Null,
    预交id_In        病人预交记录.Id%Type := Null,
    卡类别id_In      病人预交记录.卡类别id%Type := Null,
    加入序号状态_In  Number := 0,
    是否自助设备_In  Number := 0,
    结帐id_In        门诊费用记录.结帐id%Type := Null,
    锁定类型_In      Number := 0,
    保险结算_In      Varchar2 := Null,
    冲预交_In        Number := Null,
    支付卡号_In      病人预交记录.卡号%Type := Null,
    退号重用_In      Number := 1,
    费别_In          门诊费用记录.费别%Type := Null,
    冲预交病人ids_In Varchar2 := Null,
    机器名_In        挂号序号状态.机器名%Type := Null,
    更新年龄_In      Number := 0,
    购买病历_In      Number := 0
  ) As
    --功能：三方机构进行挂号(包含预约;预约挂号不扣款;预约挂号扣款),出诊表排班模式下使用
    --入参: 操作方式_IN:1-表示挂号,2-表示预约挂号不扣款,3-表示预约挂号,扣款
    --      加入序号状态_In:1-表示强制加入挂号序号状态表中;否则根据启用序号或启用时段时才加入.
    --      是否自助设备_In:1-表示是医院的自助设备进行此函数的调用,自助设备调用此函数 允许加号,否则不允许
    --      锁定类型_In :0-产生正常数据 1-表示对单据进行锁定,产生未生效的单据信息;2-对锁定的记录进行解锁-生成正常数据:未生效的单据在银行扣款完成后进行解锁
    --      保险结算_IN:格式="结算方式|结算金额||....."
    --      冲预交病人ids_In:多个用逗分分离,冲预交时有效(冲预交类别与业务操作保持一致),主要是使用家属的预交款
    Err_Item Exception;
    v_Err_Msg  Varchar2(255);
    n_打印id   票据打印内容.Id%Type;
    n_返回值   病人预交记录.金额%Type;
    v_排队号码 Varchar2(20);
    v_队列名称 排队叫号队列.队列名称%Type;
    n_预交id   病人预交记录.Id%Type;
    n_挂号id   病人挂号记录.Id%Type;
    v_结算内容 Varchar2(3000);
    v_当前结算 Varchar2(150);
  
    v_结算方式       病人预交记录.结算方式%Type;
    n_结算金额       病人预交记录.冲预交%Type;
    n_结算合计       Number(16, 5);
    n_预交金额       病人预交记录.冲预交%Type;
    n_组id           财务缴款分组.Id%Type;
    d_排队时间       Date;
    n_锁定           Number;
    n_同科限约一个号 Number(18);
    n_病人预约科室数 Number(18);
    n_已约科室       Number(18);
  
    n_合作单位限制       Number(18);
    n_是否开放           Number(1);
    n_Count              Number(18);
    n_行号               Number(18);
    n_序号               病人挂号记录.号序%Type;
    n_费用id             门诊费用记录.Id%Type;
    n_价格父号           Number(18);
    n_原项目id           收费项目目录.Id%Type;
    n_原收入项目id       收费项目目录.Id%Type;
    v_诊室               病人挂号记录.诊室%Type;
    n_安排id             挂号安排.Id%Type;
    n_实收金额合计       门诊费用记录.实收金额%Type;
    n_开单部门id         门诊费用记录.开单部门id%Type;
    n_实收金额           门诊费用记录.实收金额%Type;
    n_应收金额           门诊费用记录.实收金额%Type;
    n_结帐id             病人结帐记录.Id%Type;
    v_Temp               Varchar2(500);
    v_结算方式记录       Varchar2(1000);
    n_预约时段序号       Number;
    n_序号控制           临床出诊记录.是否序号控制%Type;
    n_限约数             临床出诊记录.限约数%Type;
    n_项目id             临床出诊记录.项目id%Type;
    n_科室id             临床出诊记录.科室id%Type;
    d_终止时间           临床出诊记录.终止时间%Type;
    v_医生姓名           临床出诊记录.医生姓名%Type;
    n_医生id             临床出诊记录.医生id%Type;
    n_预约顺序号         临床出诊序号控制.预约顺序号%Type;
    n_预约总数           Number;
    d_时段开始时间       Date;
    d_时段终止时间       Date;
    v_收费项目ids        Varchar2(300);
    n_三方卡标志         Number;
    n_预约数量           合作单位挂号汇总.已约数%Type;
    n_号序               病人挂号记录.号序%Type;
    d_登记时间           Date;
    n_单笔金额           病人预交记录.冲预交%Type;
    v_结算号码           病人预交记录.结算号码%Type;
    v_操作员编号         人员表.编号%Type;
    v_操作员姓名         人员表.姓名%Type;
    n_预约               Integer;
    v_现金               病人预交记录.结算方式%Type;
    v_星期               挂号安排时段.星期%Type;
    n_启用分时段         Integer;
    n_已挂数             病人挂号汇总.已挂数%Type;
    n_已约数             病人挂号汇总.已约数%Type;
    n_其中已接收         病人挂号汇总.已约数%Type;
    n_预约生成队列       Number;
    n_限号数             临床出诊记录.限号数%Type;
    d_Date               Date;
    n_挂号序号           Number;
    v_排队标记           排队叫号队列.排队标记%Type;
    v_排队序号           排队叫号队列.排队序号%Type;
    v_机器名             挂号序号状态.机器名%Type;
    v_序号操作员         挂号序号状态.操作员姓名%Type;
    v_序号机器名         挂号序号状态.机器名%Type;
    n_序号锁定           Number := 0;
    n_病历费id           收费特定项目.收费细目id%Type;
    v_付款方式           病人挂号记录.医疗付款方式%Type;
    v_费别               门诊费用记录.费别%Type;
    n_屏蔽费别           Number(3) := 0;
    v_冲预交病人ids      Varchar2(4000);
    v_年龄               病人信息.年龄%Type;
    n_合作单位限数量模式 Number;
    n_同科限号数         Number;
    n_同科限约数         Number;
    n_病人挂号科室数     Number;
    n_Exists             Number(5);
  
    Cursor c_Pati(n_病人id 病人信息.病人id%Type) Is
      Select a.病人id, a.姓名, a.性别, a.年龄, a.住院号, a.门诊号, a.费别, a.险类, c.编码 As 付款方式
      From 病人信息 A, 医疗付款方式 C
      Where a.病人id = n_病人id And a.医疗付款方式 = c.名称(+);
  
    r_Pati c_Pati%RowType;
  
    --该游标用于收费冲预交的可用预交列表
    --以ID排序，优先冲上次未冲完的。
    Cursor c_Deposit
    (
      n_病人id        病人信息.病人id%Type,
      v_冲预交病人ids Varchar2
    ) Is
      Select *
      From (Select a.Id, a.病人id, a.记录状态, a.预交类别, a.No, Nvl(a.金额, 0) As 金额
             From 病人预交记录 A,
                  (Select NO, Sum(Nvl(a.金额, 0)) As 金额
                    From 病人预交记录 A
                    Where a.结帐id Is Null And Nvl(a.金额, 0) <> 0 And
                          a.病人id In (Select Column_Value From Table(f_Num2list(v_冲预交病人ids))) And Nvl(a.预交类别, 2) = 1
                    Group By NO
                    Having Sum(Nvl(a.金额, 0)) <> 0) B
             Where a.结帐id Is Null And Nvl(a.金额, 0) <> 0 And a.结算方式 Not In (Select 名称 From 结算方式 Where 性质 = 5) And
                   a.No = b.No And a.病人id In (Select Column_Value From Table(f_Num2list(v_冲预交病人ids))) And
                   Nvl(a.预交类别, 2) = 1
             Union All
             Select 0 As ID, Max(病人id) As 病人id, 记录状态, 预交类别, NO, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额
             From 病人预交记录
             Where 记录性质 In (1, 11) And 结帐id Is Not Null And Nvl(金额, 0) <> Nvl(冲预交, 0) And
                   病人id In (Select Column_Value From Table(f_Num2list(v_冲预交病人ids))) And Nvl(预交类别, 2) = 1 Having
              Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
             Group By 记录状态, NO, 预交类别)
      Order By Decode(病人id, Nvl(n_病人id, 0), 0, 1), ID, NO;
  
    Function Zl_诊室(记录id_In 临床出诊记录.Id%Type) Return Varchar2 As
      n_分诊方式 临床出诊记录.分诊方式%Type;
      v_诊室     病人挂号记录.诊室%Type;
      v_Rowid    Varchar2(500);
      n_Next     Integer;
      n_First    Integer;
    Begin
    
      If 锁定类型_In = 2 Then
        --对单据进行解锁,首先检查是否存在锁定
        Select Count(Rowid)
        Into n_锁定
        From 病人挂号记录 Where　记录状态 = 0 And NO = 单据号_In And 病人id = 病人id_In;
        If n_锁定 = 0 Then
          v_Err_Msg := '单据号为(' || 单据号_In || ')的单据,不存在或者已经被解锁!';
          Raise Err_Item;
        End If;
        Select Max(号序) Into n_号序 From 病人挂号记录 Where　记录状态 = 0 And NO = 单据号_In And 病人id = 病人id_In;
      End If;
    
      Begin
        Select Nvl(分诊方式, 0) Into n_分诊方式 From 临床出诊记录 Where ID = 记录id_In;
      Exception
        When Others Then
          v_Err_Msg := '出诊记录(' || 记录id_In || ')未找到!';
          Raise Err_Item;
      End;
    
      --0-不分诊、1-指定诊室、2-动态分诊、3-平均分诊,对应门诊诊室设置
      v_诊室 := Null;
      If n_分诊方式 = 1 Then
        --1-指定诊室
        Begin
          Select b.名称 Into v_诊室 From 临床出诊诊室记录 A, 门诊诊室 B Where a.诊室id = b.Id And a.记录id = 记录id_In;
        Exception
          When Others Then
            v_诊室 := Null;
        End;
      End If;
      If n_分诊方式 = 2 Then
        --2-动态分诊:该个号别当天挂号未诊数最少的诊室
        For c_诊室 In (Select 门诊诊室, Sum(Num) As Num
                     From (Select b.名称 As 门诊诊室, 0 As Num
                            From 临床出诊诊室记录 A, 门诊诊室 B
                            Where a.诊室id = b.Id And a.记录id = 记录id_In
                            Union All
                            Select 诊室, Count(诊室) As Num
                            From 病人挂号记录
                            Where Nvl(执行状态, 0) = 0 And 发生时间 Between Trunc(Sysdate) And Sysdate And 号别 = 号码_In And
                                  诊室 In (Select d.名称
                                         From 临床出诊诊室记录 C, 门诊诊室 D
                                         Where c.诊室id = d.Id And c.记录id = 记录id_In)
                            Group By 诊室)
                     Group By 门诊诊室
                     Order By Num) Loop
          v_诊室 := c_诊室.门诊诊室;
          Exit;
        End Loop;
      End If;
      If n_分诊方式 = 3 Then
        --平均分诊：当前分配=1表示下次应取的当前诊室
        n_Next  := 0;
        n_First := 1;
        For c_诊室 In (Select a.Rowid As Rid, b.名称 As 门诊诊室, a.当前分配
                     From 临床出诊诊室记录 A, 门诊诊室 B
                     Where a.诊室id = b.Id And a.记录id = 记录id_In) Loop
          If n_First = 1 Then
            v_Rowid := c_诊室.Rid;
          End If;
          If n_Next = 1 Then
            v_诊室 := c_诊室.门诊诊室;
            Update 临床出诊诊室记录 Set 当前分配 = 1 Where Rowid = c_诊室.Rid;
            Exit;
          End If;
          If Nvl(c_诊室.当前分配, 0) = 1 Then
            Update 临床出诊诊室记录 Set 当前分配 = 0 Where Rowid = c_诊室.Rid;
            n_Next := 1;
          End If;
        End Loop;
        If v_诊室 Is Null Then
          Update 临床出诊诊室记录 Set 当前分配 = 1 Where Rowid = v_Rowid Returning 诊室id Into v_诊室;
          Select 名称 Into v_诊室 From 门诊诊室 Where ID = v_诊室;
        End If;
      End If;
      Return v_诊室;
    End;
  
    Function Zl_操作员
    (
      Type_In     Integer,
      Splitstr_In Varchar2
    ) Return Varchar2 As
      n_Step Number(18);
      v_Sub  Varchar2(1000);
      --Type_In:0-获取缺省部门ID;1-获取操作员编号;2-获取操作员姓名
      -- SplitStr:格式为:部门ID,部门名称;人员ID,人员编号,人员姓名(用Zl_Identity获取的)
    Begin
      If Type_In = 0 Then
        --缺省部门
        n_Step := Instr(Splitstr_In, ',');
        v_Sub  := Substr(Splitstr_In, 1, n_Step - 1);
        Return v_Sub;
      End If;
      If Type_In = 1 Then
        --操作员编码
        n_Step := Instr(Splitstr_In, ';');
        v_Sub  := Substr(Splitstr_In, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, 1, n_Step - 1);
        Return v_Sub;
      End If;
      If Type_In = 2 Then
        --操作员姓名
        n_Step := Instr(Splitstr_In, ';');
        v_Sub  := Substr(Splitstr_In, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        Return v_Sub;
      End If;
    End;
  
  Begin
    v_冲预交病人ids := Nvl(冲预交病人ids_In, 病人id_In);
  
    Begin
      Select 名称 Into v_现金 From 结算方式 Where 性质 = 1;
    Exception
      When Others Then
        v_现金 := '现金';
    End;
  
    If 费别_In Is Null Then
      Select Zl_Custom_Getpatifeetype(1, 病人id_In) Into v_费别 From Dual;
    Else
      v_费别 := 费别_In;
    End If;
    If v_费别 Is Null Then
      n_屏蔽费别 := 1;
      Select 名称 Into v_费别 From 费别 Where 缺省标志 = 1 And Rownum < 2;
    End If;
    Update 病人信息 Set 费别 = v_费别 Where 病人id = 病人id_In;
  
    If 更新年龄_In = 1 Then
      Select Zl_Age_Calc(病人id_In) Into v_年龄 From Dual;
      If v_年龄 Is Not Null Then
        Update 病人信息 Set 年龄 = v_年龄 Where 病人id = 病人id_In;
      End If;
    End If;
    --获取当前机器名称
    If 机器名_In Is Not Null Then
      v_机器名 := 机器名_In;
    Else
      Select Terminal Into v_机器名 From V$session Where Audsid = Userenv('sessionid');
    End If;
    n_实收金额合计 := 0;
  
    Select Count(*) + 1
    Into n_挂号序号
    From 病人挂号记录
    Where 出诊记录id = 记录id_In And 登记时间 Between Trunc(发生时间_In) And Trunc(发生时间_In + 1) - 1 / 24 / 60 / 60;
  
    --部门ID,部门名称;人员ID,人员编号,人员姓名
    v_Temp := Zl_Identity(0);
    If Nvl(v_Temp, ' ') = ' ' Then
      v_Err_Msg := '当前操作人员未设置对应的人员关系,不能继续。';
      Raise Err_Item;
    End If;
  
    If 登记时间_In Is Null Then
      d_登记时间 := Sysdate;
    Else
      d_登记时间 := 登记时间_In;
    End If;
    If Trunc(Sysdate) > Trunc(发生时间_In) Then
      v_Err_Msg := '不能挂以前的号(' || To_Char(发生时间_In, 'yyyy-mm-dd') || ')。';
      Raise Err_Item;
    End If;
  
    v_Temp           := Nvl(zl_GetSysParameter('病人同科限制', 1111), '0,0|0,0');
    n_同科限号数     := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
    v_Temp           := Substr(v_Temp, Instr(v_Temp, '|') + 1);
    n_同科限约数     := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
    v_Temp           := Nvl(zl_GetSysParameter('病人预约科室数', 1111), '0|0');
    n_病人预约科室数 := Substr(v_Temp, Instr(v_Temp, '|') + 1);
    n_病人挂号科室数 := Substr(v_Temp, 1, Instr(v_Temp, '|') - 1);
    n_开单部门id     := To_Number(Zl_操作员(0, v_Temp));
    v_操作员编号     := Zl_操作员(1, v_Temp);
    v_操作员姓名     := Zl_操作员(2, v_Temp);
    n_组id           := Zl_Get组id(v_操作员姓名);
  
    If 操作方式_In <> 1 Then
      --预约检查是否添加合作单位控制
      --如果设置了合作单位控制 则
      Begin
        Select 1
        Into n_合作单位限制
        From 临床出诊挂号控制记录
        Where 记录id = 记录id_In And 类型 = 1 And 性质 = 1 And 控制方式 <> 4 And Rownum < 2;
      Exception
        When Others Then
          n_合作单位限制 := 0;
      End;
    End If;
  
    If 操作方式_In <> 2 Then
      v_诊室 := Zl_诊室(记录id_In);
    End If;
  
    --因为现在有按日编单据号规则,日挂号量不能超过10000次,所以要检查唯一约束。
    Select Count(*) Into n_Count From 门诊费用记录 Where 记录性质 = 4 And 记录状态 In (1, 3) And NO = 单据号_In;
    If n_Count <> 0 Then
      v_Err_Msg := '挂号单据号重复,不能保存！' || Chr(13) || '如果使用了按日顺序编号,当日挂号量不能超过10000人次。';
      Raise Err_Item;
    End If;
  
    Open c_Pati(病人id_In);
    n_Count := 0;
    Begin
      Fetch c_Pati
        Into r_Pati;
    Exception
      When Others Then
        n_Count := -1;
    End;
    If n_Count = -1 Then
      v_Err_Msg := '病人未找到，不能继续。';
      Raise Err_Item;
    End If;
  
    Begin
      Select Nvl(是否分时段, 0), 限号数, 已挂数, 其中已接收, 已约数, 是否序号控制, 限约数, 项目id, 科室id, 医生id, 医生姓名
      Into n_启用分时段, n_限号数, n_已挂数, n_其中已接收, n_已约数, n_序号控制, n_限约数, n_项目id, n_科室id, n_医生id, v_医生姓名
      From 临床出诊记录
      Where ID = 记录id_In;
    Exception
      When Others Then
        n_Count := -1;
    End;
    If n_Count = -1 Then
      v_Err_Msg := '该号别没有在' || To_Char(发生时间_In, 'yyyy-mm-dd hh24:mi:ss') || '中进行安排。';
      Raise Err_Item;
    End If;
  
    --对参数控制进行检查
    --仅在预约不扣款时进行检查
    If 操作方式_In = 2 Then
      If Nvl(n_同科限约数, 0) <> 0 Or Nvl(n_病人预约科室数, 0) <> 0 Then
        n_已约科室 := 0;
        For c_Chkitem In (Select Distinct 执行部门id
                          From 病人挂号记录
                          Where 记录状态 = 1 And 记录性质 = 2 And 预约时间 Between Trunc(发生时间_In) And
                                Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60) Loop
          n_已约科室 := n_已约科室 + 1;
        End Loop;
        If n_已约科室 >= Nvl(n_病人预约科室数, 0) And Nvl(n_病人预约科室数, 0) > 0 Then
          v_Err_Msg := '同一病人最多同时能预约[' || Nvl(n_病人预约科室数, 0) || ']个科室,不能再预约！';
          Raise Err_Item;
        End If;
      
        Select Count(1)
        Into n_Count
        From 病人挂号记录
        Where 记录状态 = 1 And 记录性质 = 2 And 预约时间 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And
              执行部门id = n_科室id;
        If n_Count >= Nvl(n_同科限约数, 0) And Nvl(n_同科限约数, 0) > 0 Then
          v_Err_Msg := '该病人已经在该科室预约了' || n_Count || '次,不能再预约！';
          Raise Err_Item;
        End If;
      End If;
    Else
      If Nvl(n_同科限号数, 0) <> 0 Or Nvl(n_病人挂号科室数, 0) <> 0 Then
        n_已约科室 := 0;
        For c_Chkitem In (Select Distinct 执行部门id
                          From 病人挂号记录
                          Where 记录状态 = 1 And 记录性质 = 1 And 发生时间 Between Trunc(发生时间_In) And
                                Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60) Loop
          n_已约科室 := n_已约科室 + 1;
        End Loop;
        If n_已约科室 >= Nvl(n_病人挂号科室数, 0) And Nvl(n_病人挂号科室数, 0) > 0 Then
          v_Err_Msg := '同一病人最多同时能挂号[' || Nvl(n_病人挂号科室数, 0) || ']个科室,不能再挂号！';
          Raise Err_Item;
        End If;
      
        Select Count(1)
        Into n_Count
        From 病人挂号记录
        Where 记录状态 = 1 And 记录性质 = 1 And 发生时间 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And
              执行部门id = n_科室id;
        If n_Count >= Nvl(n_同科限号数, 0) And Nvl(n_同科限号数, 0) > 0 Then
          v_Err_Msg := '该病人已经在该科室挂号了' || n_Count || '次,不能再挂号！';
          Raise Err_Item;
        End If;
      End If;
    End If;
  
    d_Date         := Null;
    d_时段开始时间 := Null;
  
    If Nvl(n_限号数, 0) >= 0 Or n_限号数 Is Null Then
      If n_启用分时段 = 1 Then
        If Nvl(n_序号控制, 0) = 1 Then
          If Nvl(是否自助设备_In, 0) = 0 Then
            Select Count(*), Max(开始时间)
            Into n_Count, d_时段开始时间
            From 临床出诊序号控制
            Where 记录id = 记录id_In And 序号 = Nvl(号序_In, 0);
          
            v_Temp := '挂号';
            If 操作方式_In > 1 Then
              v_Temp := '预约挂号';
            End If;
          
            If n_Count = 0 Then
              v_Err_Msg := '号别为' || 号码_In || '的挂号安排中不存在序号为' || Nvl(号序_In, 0) || '的安排,不能再' || v_Temp || '！';
              Raise Err_Item;
            End If;
          End If;
        
          --过点的,不能选择挂号
          If Trunc(Sysdate) = Trunc(发生时间_In) Then
            --挂当天的号
            v_Temp := To_Char(Sysdate, 'yyyy-mm-dd') || ' ';
            For v_时段 In (Select To_Date(v_Temp || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                                To_Date(To_Char(Sysdate + Decode(Sign(开始时间 - 终止时间), 1, 1, 0), 'yyyy-mm-dd') || ' ' ||
                                         To_Char(终止时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 终止时间, 数量, 是否预约
                         From 临床出诊序号控制
                         Where 记录id = 记录id_In And 序号 = Nvl(号序_In, 0)) Loop
              If Sysdate > v_时段.终止时间 Then
                v_Err_Msg := '号别为' || 号码_In || '及号序为' || Nvl(号序_In, 0) || '的安排,已经超过时点,不能再' || v_Temp || '！';
                Raise Err_Item;
              End If;
            End Loop;
          End If;
        Elsif 操作方式_In > 1 Then
          --未启用序号的,需要检查预约的情况
          n_Count := 0;
          For v_时段 In (Select 序号, 开始时间, 终止时间, 数量, 是否预约
                       From 临床出诊序号控制
                       Where 记录id = 记录id_In And
                             (('3000-01-10 ' || To_Char(发生时间_In, 'HH24:MI:SS') Between
                             Decode(Sign(开始时间 - 终止时间 - 1 / 24 / 60 / 60), 1,
                                      '3000-01-09 ' || To_Char(开始时间, 'HH24:MI:SS'),
                                      '3000-01-10 ' || To_Char(开始时间, 'HH24:MI:SS')) And
                             '3000-01-10 ' || To_Char(终止时间 - 1 / 24 / 60 / 60, 'HH24:MI:SS')) Or
                             ('3000-01-10 ' || To_Char(发生时间_In, 'HH24:MI:SS') Between
                             '3000-01-10 ' || To_Char(开始时间, 'HH24:MI:SS') And
                             Decode(Sign(开始时间 - 终止时间 - 1 / 24 / 60 / 60), 1,
                                      '3000-01-11 ' || To_Char(终止时间 - 1 / 24 / 60 / 60, 'HH24:MI:SS'),
                                      '3000-01-10 ' || To_Char(终止时间 - 1 / 24 / 60 / 60, 'HH24:MI:SS'))))) Loop
            n_预约时段序号 := v_时段.序号;
            d_时段开始时间 := v_时段.开始时间;
            d_时段终止时间 := v_时段.终止时间;
          
            Select Count(*), Max(序号), Max(预约顺序号) + 1
            Into n_Count, n_预约总数, n_预约顺序号
            From 临床出诊序号控制
            Where 记录id = 记录id_In And Nvl(挂号状态, 0) Not In (0, 4, 5);
          
            If Nvl(n_Count, 0) > Nvl(v_时段.数量, 0) And 锁定类型_In <> 2 Then
              v_Err_Msg := '号别为' || 号码_In || '的挂号安排中在' || To_Char(v_时段.开始时间, 'hh24:mi:ss') || '至' ||
                           To_Char(v_时段.终止时间, 'hh24:mi:ss') || '最多只能预约' || Nvl(v_时段.数量, 0) || '人,不能再进行预约挂号！';
              Raise Err_Item;
            End If;
            n_Count := 1;
          End Loop;
        
          If n_Count = 0 Then
            v_Err_Msg := '号别为' || 号码_In || '的挂号安排中没有相关的安排计划(' || To_Char(发生时间_In, 'YYYY-mm-dd HH24:MI:SS') ||
                         '),不能进行预约挂号！';
            Raise Err_Item;
          End If;
        End If;
      End If;
    End If;
  
    If 操作方式_In = 1 And 锁定类型_In <> 2 Then
      --挂号规则:
      --  已挂数不能大于限号数
      If n_已挂数 >= Nvl(n_限号数, 0) And n_限号数 Is Not Null Then
        v_Err_Msg := '该号别今天已达到限号数 ' || Nvl(n_限号数, 0) || '不能再挂号！';
        Raise Err_Item;
      End If;
    End If;
  
    If 操作方式_In > 1 Then
      --预约的相关检查
      --规则:
      --   1.已限约不能超过限约数
      --   2.检查是否启用时段的
      If n_已约数 >= Nvl(n_限约数, 0) And Nvl(n_限约数, 0) <> 0 And n_限约数 Is Not Null And 锁定类型_In <> 2 Then
        v_Err_Msg := '该号别已达到限约数 ' || Nvl(n_限约数, 0) || '不能再预约挂号！';
        Raise Err_Item;
      End If;
      If 预约方式_In Is Not Null Then
        Select Zl_预约方式_Check(记录id_In, 号序_In, 预约方式_In) Into n_Exists From Dual;
        If n_Exists = 0 Then
          v_Err_Msg := '传入的预约方式' || 预约方式_In || '预约数量达到上限,不能继续。';
          Raise Err_Item;
        End If;
      End If;
    End If;
    If n_合作单位限制 > 0 And 操作方式_In <> 1 And 合作单位_In Is Not Null Then
      If Nvl(n_序号控制, 0) = 1 And Nvl(号序_In, 0) = 0 Then
        v_Err_Msg := '当前安排使用了序号控制,请确认所需要预约的序号,不能继续。';
        Raise Err_Item;
      End If; --Nvl(r_安排.序号控制, 0) =0
    
      n_序号 := Case
                When Nvl(n_序号控制, 0) = 1 Or n_启用分时段 = 1 And 操作方式_In > 1 Then
                 Nvl(号序_In, 0)
                Else
                 0
              End;
    
      --合作单位控制模式
      Select Nvl(控制方式, 0)
      Into n_合作单位限数量模式
      From 临床出诊挂号控制记录
      Where 记录id = 记录id_In And 名称 = 合作单位_In And 类型 = 1 And 性质 = 1 And Rownum < 2;
    
      If n_合作单位限数量模式 = 0 Then
        v_Err_Msg := '当前号码(' || Nvl(号码_In, 0) || '未开放' || 合作单位_In || '的预约,不能继续。';
        Raise Err_Item;
      End If;
      If n_合作单位限数量模式 = 1 Or n_合作单位限数量模式 = 2 Then
        Select 数量
        Into n_Count
        From 临床出诊挂号控制记录
        Where 记录id = 记录id_In And 名称 = 合作单位_In And 类型 = 1 And 性质 = 1;
        If n_合作单位限数量模式 = 1 Then
          n_Count := Round(Nvl(n_限约数, n_限号数) * n_Count / 100);
        End If;
        Select Count(1)
        Into n_Exists
        From 病人挂号记录
        Where 记录状态 = 1 And 出诊记录id = 记录id_In And 合作单位 = 合作单位_In;
        If n_Exists >= n_Count Then
          v_Err_Msg := '当前号码(' || Nvl(号码_In, 0) || '已经超过' || 合作单位_In || '的预约限制,不能继续。';
          Raise Err_Item;
        End If;
      End If;
      --开放序号检查
      If n_合作单位限数量模式 = 3 Then
        For c_合作单位 In (Select 序号, 数量
                       From 临床出诊挂号控制记录
                       Where 记录id = 记录id_In And 名称 = 合作单位_In And 类型 = 1 And 性质 = 1 And 序号 = 号序_In) Loop
          If n_序号控制 = 1 Then
            Begin
              Select 1
              Into n_Count
              From 临床出诊序号控制
              Where 记录id = 记录id_In And 序号 = 号序_In And Nvl(挂号状态, 0) = 0;
            Exception
              When Others Then
                n_Count := 0;
            End;
            If n_Count = 1 Then
              n_是否开放 := 1;
            Else
              v_Err_Msg := '当前序号(' || Nvl(号序_In, 0) || '已经超过' || 合作单位_In || '的预约限制,不能继续。';
              Raise Err_Item;
            End If;
          Else
            Select Count(1)
            Into n_Count
            From 临床出诊序号控制
            Where 记录id = 记录id_In And 序号 = 号序_In And 预约顺序号 Is Not Null And Nvl(挂号状态, 0) <> 0;
            If n_Count >= c_合作单位.数量 Then
              v_Err_Msg := '当前序号(' || Nvl(号序_In, 0) || '已经超过' || 合作单位_In || '的预约限制,不能继续。';
              Raise Err_Item;
            Else
              n_是否开放 := 1;
            End If;
          End If;
        End Loop;
      
        If Nvl(n_是否开放, 0) = 0 Then
          v_Err_Msg := '当前序号(' || Nvl(号序_In, 0) || '未开放,不能继续。';
          Raise Err_Item;
        End If;
      End If;
    End If;
  
    --检查限号数和限约数
    n_行号         := 1;
    n_原项目id     := 0;
    n_原收入项目id := 0;
    n_实收金额合计 := 0;
    If 锁定类型_In <> 1 Then
      If 操作方式_In <> 2 Then
        If Nvl(结帐id_In, 0) = 0 Then
          --这里应该程序传入
          Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
        Else
          n_结帐id := 结帐id_In;
        End If;
      Else
        n_结帐id := Null;
      End If;
    End If;
  
    If Nvl(购买病历_In, 0) = 1 Then
      Begin
        Select 收费细目id Into n_病历费id From 收费特定项目 Where 特定项目 = '病历费';
        v_收费项目ids := n_项目id || ',' || n_病历费id;
      Exception
        When Others Then
          v_Err_Msg := '不能确定病历费,挂号失败!';
          Raise Err_Item;
      End;
    Else
      v_收费项目ids := n_项目id;
    End If;
  
    For c_Item In (Select 1 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, 1 As 数次,
                          c.Id As 收入项目id, c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, Null As 从属父号
                   From 收费项目目录 A, 收费价目 B, 收入项目 C
                   Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = n_项目id And Sysdate Between b.执行日期 And
                         Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))
                   Union All
                   Select 2 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, 1 As 数次,
                          c.Id As 收入项目id, c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, Null As 从属父号
                   From 收费项目目录 A, 收费价目 B, 收入项目 C
                   Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = n_病历费id And Sysdate Between b.执行日期 And
                         Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))
                   Union All
                   Select 3 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, d.从项数次 As 数次,
                          c.Id As 收入项目id, c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, 1 As 从属父号
                   From 收费项目目录 A, 收费价目 B, 收入项目 C, 收费从属项目 D
                   Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = d.从项id And
                         d.主项id In (Select Column_Value From Table(f_Str2list(v_收费项目ids))) And Sysdate Between b.执行日期 And
                         Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))
                   Order By 性质, 项目编码, 收入编码) Loop
      n_价格父号 := Null;
      If n_原项目id = c_Item.项目id Then
        If n_原收入项目id <> c_Item.收入项目id Then
          n_价格父号 := n_行号;
        End If;
        n_原收入项目id := c_Item.收入项目id;
      End If;
      n_原项目id := c_Item.项目id;
      n_应收金额 := Round(c_Item.数次 * c_Item.单价, 5);
      n_实收金额 := n_应收金额;
      If Nvl(c_Item.屏蔽费别, 0) <> 1 And n_屏蔽费别 = 0 Then
        --打折:
        v_Temp     := Zl_Actualmoney(r_Pati.费别, c_Item.项目id, c_Item.收入项目id, n_应收金额);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ':') + 1);
        n_实收金额 := Zl_To_Number(v_Temp);
      End If;
      n_实收金额合计 := Nvl(n_实收金额合计, 0) + n_实收金额;
    
      --锁定单据不产生费用
      If 锁定类型_In <> 1 Then
        --产生病人挂号费用(可能单独是或包括病历费用)
        Select 病人费用记录_Id.Nextval Into n_费用id From Dual;
        --:操作方式_IN:1-表示挂号,2-表示预约挂号不扣款,3-表示预约挂号,扣款
        Insert Into 门诊费用记录
          (ID, 记录性质, 记录状态, 序号, 价格父号, 从属父号, NO, 实际票号, 门诊标志, 加班标志, 附加标志, 发药窗口, 病人id, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 病人科室id,
           收费类别, 计算单位, 收费细目id, 收入项目id, 收据费目, 付数, 数次, 标准单价, 应收金额, 实收金额, 结帐金额, 结帐id, 记帐费用, 开单部门id, 开单人, 划价人, 执行部门id, 执行人,
           操作员编号, 操作员姓名, 发生时间, 登记时间, 保险大类id, 保险项目否, 保险编码, 统筹金额, 摘要, 结论, 缴款组id)
        Values
          (n_费用id, 4, Decode(操作方式_In, 2, 0, 1), n_行号, n_价格父号, c_Item.从属父号, 单据号_In, 票据号_In, 1, Null, Null,
           Decode(操作方式_In, 2, To_Char(号序_In), v_诊室), r_Pati.病人id, r_Pati.门诊号, r_Pati.付款方式, r_Pati.姓名, r_Pati.性别,
           r_Pati.年龄, r_Pati.费别, n_科室id, c_Item.类别, 号码_In, c_Item.项目id, c_Item.收入项目id, c_Item.收据费目, 1, c_Item.数次,
           c_Item.单价, n_应收金额, n_实收金额, Decode(操作方式_In, 2, Null, n_实收金额), n_结帐id, 0, n_开单部门id, v_操作员姓名,
           Decode(操作方式_In, 2, v_操作员姓名, Null), n_科室id, v_医生姓名, v_操作员编号, v_操作员姓名, 发生时间_In, d_登记时间, Null, 0, Null, Null,
           摘要_In, 预约方式_In, Decode(操作方式_In, 2, Null, n_组id));
      End If;
      n_行号 := n_行号 + 1;
    
    End Loop;
  
    If Round(Nvl(挂号金额合计_In, 0), 5) <> Round(Nvl(n_实收金额合计, 0), 5) Then
      v_Err_Msg := '本次挂号金额不正确,可能是因为医院调整了价格,请重新获取挂号收费项目的价格,不能继续。';
      Raise Err_Item;
    End If;
  
    If n_启用分时段 = 1 Then
      d_Date := To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(d_时段开始时间, 'hh24:mi:ss'),
                        'yyyy-mm-dd hh24:mi:ss');
    Else
      d_Date := Trunc(发生时间_In);
    End If;
  
    --更新挂号序号状态
    If 锁定类型_In <> 2 Then
      n_号序 := 号序_In;
    End If;
    Begin
      Select 1
      Into n_Count
      From 临床出诊序号控制
      Where 记录id = 记录id_In And 序号 = n_号序 And Nvl(挂号状态, 0) Not In (0, 5);
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count = 1 Then
      If n_启用分时段 = 0 And Nvl(n_序号控制, 0) = 1 Then
        n_号序 := Null;
      End If;
      If n_启用分时段 = 1 And Nvl(n_序号控制, 0) = 1 Then
        v_Err_Msg := '当前序号已被使用，请重新选择一个序号！';
        Raise Err_Item;
      End If;
    End If;
    n_Count := 0;
    If n_启用分时段 = 0 And Nvl(n_序号控制, 0) = 1 And n_号序 Is Null And 锁定类型_In <> 2 Then
      Select Nvl(Min(序号), 0)
      Into n_号序
      From 临床出诊序号控制
      Where 记录id = 记录id_In And Nvl(挂号状态, 0) = 5 And 操作员姓名 = v_操作员姓名 And 工作站名称 = v_机器名;
      If n_号序 = 0 Then
        Select Nvl(Max(序号), 0) Into n_号序 From 临床出诊序号控制 Where 记录id = 记录id_In And Nvl(挂号状态, 0) = 0;
        If n_号序 = 0 Then
          Select Nvl(Max(序号), 0) + 1 Into n_号序 From 临床出诊序号控制 Where 记录id = 记录id_In;
        End If;
      End If;
    End If;
  
    If n_启用分时段 = 1 And 锁定类型_In <> 2 Then
      If 操作方式_In > 1 And Nvl(n_序号控制, 0) = 0 Then
        --规则:预约时段序号||预约数
        If Nvl(n_预约总数, 0) = 0 Then
          v_Temp := Nvl(n_限约数, 0);
          v_Temp := LTrim(RTrim(v_Temp));
          v_Temp := LPad(Nvl(n_预约总数, 0) + 1, Length(v_Temp), '0');
          v_Temp := n_预约时段序号 || v_Temp;
          n_号序 := To_Number(v_Temp);
        Else
          n_号序 := n_预约总数 + 1;
        End If;
      End If;
    End If;
  
    If Nvl(n_序号控制, 0) = 1 Or (操作方式_In > 1 And n_启用分时段 = 1) Or 加入序号状态_In = 1 Then
      --锁定序号的处理
      Begin
        Select 操作员姓名, 工作站名称
        Into v_序号操作员, v_序号机器名
        From 临床出诊序号控制
        Where 挂号状态 = 5 And 记录id = 记录id_In And 序号 = n_号序;
        n_序号锁定 := 1;
      Exception
        When Others Then
          v_序号操作员 := Null;
          v_序号机器名 := Null;
          n_序号锁定   := 0;
      End;
      If n_序号锁定 = 0 Then
        If n_启用分时段 = 1 And n_序号控制 = 0 Then
          Insert Into 临床出诊序号控制
            (记录id, 序号, 预约顺序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 类型, 名称, 操作员姓名, 备注)
            Select 记录id_In, n_预约时段序号, n_预约顺序号, d_时段开始时间, d_时段终止时间, 1, Decode(操作方式_In, 1, 0, 1), Decode(操作方式_In, 2, 2, 1),
                   1, 合作单位_In, v_操作员姓名, n_号序
            From Dual;
        Else
          Update 临床出诊序号控制
          Set 挂号状态 = Decode(操作方式_In, 2, 2, 1), 操作员姓名 = v_操作员姓名
          Where 记录id = 记录id_In And 序号 = n_号序;
        End If;
        If Sql%RowCount = 0 Then
          Begin
            If n_启用分时段 = 1 Then
              --分时段
              If n_序号控制 = 1 Then
                --序号控制
                Select Max(终止时间) Into d_终止时间 From 临床出诊序号控制 Where 记录id = 记录id_In;
                If Sysdate > d_终止时间 Then
                  d_终止时间 := Sysdate;
                End If;
                Insert Into 临床出诊序号控制
                  (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 类型, 名称, 操作员姓名)
                  Select 记录id_In, n_号序, d_终止时间, d_终止时间, 1, Decode(操作方式_In, 1, 0, 1), Decode(操作方式_In, 2, 2, 1), 1,
                         合作单位_In, v_操作员姓名
                  From Dual;
              Else
                --分时段,非序号控制
                Null;
              End If;
            Else
              --不分时段
              Insert Into 临床出诊序号控制
                (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 类型, 名称, 操作员姓名)
                Select 记录id_In, n_号序, 开始时间, 终止时间, 1, Decode(操作方式_In, 1, 0, 1), Decode(操作方式_In, 2, 2, 1), 1, 合作单位_In,
                       v_操作员姓名
                From 临床出诊序号控制
                Where 记录id = 记录id_In And 序号 = 1;
            End If;
          Exception
            When Others Then
              v_Err_Msg := '序号' || n_号序 || '已被使用,请重新选择一个序号.';
              Raise Err_Item;
          End;
        End If;
      Else
        If v_操作员姓名 <> v_序号操作员 Or v_机器名 <> v_序号机器名 Then
          v_Err_Msg := '序号' || n_号序 || '已被机器' || v_机器名 || '锁定,请重新选择一个序号.';
          Raise Err_Item;
        Else
          Update 临床出诊序号控制
          Set 挂号状态 = Decode(操作方式_In, 2, 2, 1), 锁号时间 = Null
          Where 记录id = 记录id_In And 序号 = n_号序 And 挂号状态 = 5 And 操作员姓名 = v_操作员姓名 And 工作站名称 = v_机器名;
        End If;
      End If;
    End If;
  
    --锁定单据不产生任何 费用
    If 操作方式_In <> 2 And 锁定类型_In <> 1 Then
      --挂号,预约挂号已经扣款部分
      n_预交id := 预交id_In;
      If Nvl(n_预交id, 0) = 0 Then
        Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      End If;
      n_结算合计 := 0;
      If 保险结算_In Is Not Null Then
        --各个保险结算
        v_结算内容 := 保险结算_In || '||';
        n_结算合计 := 0;
        While v_结算内容 Is Not Null Loop
          v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
          v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
          n_结算金额 := To_Number(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
          If Nvl(n_结算金额, 0) <> 0 Then
            Insert Into 病人预交记录
              (ID, 记录性质, NO, 记录状态, 病人id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 结算性质)
            Values
              (n_预交id, 4, 单据号_In, 1, Decode(病人id_In, 0, Null, 病人id_In), '保险结算', v_结算方式, d_登记时间, v_操作员编号, v_操作员姓名,
               n_结算金额, n_结帐id, n_组id, n_结帐id, 4);
            Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          End If;
          n_结算合计 := Nvl(n_结算合计, 0) + Nvl(n_结算金额, 0);
          v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
        End Loop;
      End If;
    
      If Nvl(冲预交_In, 0) <> 0 Then
        --处理总预交
        n_结算合计 := n_结算合计 + Nvl(冲预交_In, 0);
        n_预交金额 := 冲预交_In;
        For r_Deposit In c_Deposit(病人id_In, v_冲预交病人ids) Loop
          n_结算金额 := Case
                      When r_Deposit.金额 - n_预交金额 < 0 Then
                       r_Deposit.金额
                      Else
                       n_预交金额
                    End;
          If r_Deposit.Id <> 0 Then
            --第一次冲预交(填上结帐ID,金额为0)
            Update 病人预交记录 Set 冲预交 = 0, 结帐id = n_结帐id, 结算性质 = 4 Where ID = r_Deposit.Id;
          End If;
          --冲上次剩余额
          Insert Into 病人预交记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 预交类别, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名,
             操作员编号, 冲预交, 结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 结算性质)
            Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 预交类别, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行,
                   单位帐号, d_登记时间, v_操作员姓名, v_操作员编号, n_结算金额, n_结帐id, n_组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, n_结帐id, 4
            From 病人预交记录
            Where NO = r_Deposit.No And 记录状态 = r_Deposit.记录状态 And 记录性质 In (1, 11) And Rownum = 1;
        
          --更新病人预交余额
          Update 病人余额
          Set 预交余额 = Nvl(预交余额, 0) - n_结算金额
          Where 病人id = r_Deposit.病人id And 性质 = 1 And 类型 = Nvl(r_Deposit.预交类别, 2)
          Returning 预交余额 Into n_返回值;
          If Sql%RowCount = 0 Then
            Insert Into 病人余额 (病人id, 预交余额, 性质, 类型) Values (r_Deposit.病人id, -1 * n_结算金额, 1, 1);
            n_返回值 := -1 * n_结算金额;
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 病人余额
            Where 病人id = r_Deposit.病人id And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
          End If;
        
          --检查是否已经处理完
          If r_Deposit.金额 <= n_结算金额 Then
            n_预交金额 := n_预交金额 - r_Deposit.金额;
          Else
            n_预交金额 := 0;
          End If;
          If n_预交金额 = 0 Then
            Exit;
          End If;
        End Loop;
        If n_预交金额 > 0 Then
          v_Err_Msg := '预交余额不够支付本次支付金额,不能继续操作！';
          Raise Err_Item;
        End If;
      End If;
      --剩余款项,用指定结算方支付
      n_结算金额 := Nvl(n_实收金额合计, 0) - Nvl(n_结算合计, 0);
      If Nvl(n_结算金额, 0) < 0 Then
        v_Err_Msg := '挂号的相关结算金额超出了当前实结金额,不能继续操作！';
        Raise Err_Item;
      End If;
    
      If Nvl(n_结算金额, 0) <> 0 Or (Nvl(n_结算金额, 0) = 0 And Nvl(冲预交_In, 0) = 0) Then
        If 结算方式_In Is Null Then
          v_Err_Msg := '未传入指定的结算方式,不能继续操作！';
          Raise Err_Item;
        End If;
      
        If Nvl(预交id_In, 0) <> 0 Then
          --传入的预交ID_In主要是为了解决三方交易,如果医保结算站用了该ID,需要用新的ID进行更新,三方交易用转入的ID
          Update 病人预交记录 Set ID = n_预交id Where ID = Nvl(预交id_In, 0);
          n_预交id := Nvl(预交id_In, 0);
        End If;
        If Instr(结算方式_In, ',') = 0 Then
          --只传入一种结算方式的
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明,
             合作单位, 结算性质, 结算号码)
          Values
            (n_预交id, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(v_结算方式, v_现金), Nvl(n_结算金额, 0), 登记时间_In,
             v_操作员编号, v_操作员姓名, n_结帐id, '挂号收费', n_组id, 卡类别id_In, Null, 支付卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In, 4, v_结算号码);
        Else
          v_结算内容     := 结算方式_In || '|'; --以空格分开以|结尾,没有结算号码的
          n_Exists       := 0;
          v_结算方式记录 := '';
          While v_结算内容 Is Not Null Loop
            v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '|') - 1);
            v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
          
            v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
            n_单笔金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1));
          
            v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
            v_结算号码 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
          
            v_当前结算   := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
            n_三方卡标志 := To_Number(v_当前结算);
          
            If Instr('|' || v_结算方式记录 || '|', '|' || Nvl(v_结算方式, v_现金) || '|') <> 0 Then
              v_Err_Msg := '使用了重复的结算方式,请检查!';
              Raise Err_Item;
            Else
              v_结算方式记录 := v_结算方式记录 || '|' || Nvl(v_结算方式, v_现金);
            End If;
          
            If n_三方卡标志 = 0 Then
              Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
              Insert Into 病人预交记录
                (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号,
                 交易说明, 合作单位, 结算性质, 结算号码)
              Values
                (n_预交id, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(v_结算方式, v_现金), Nvl(n_单笔金额, 0), 登记时间_In,
                 v_操作员编号, v_操作员姓名, n_结帐id, '挂号收费', n_组id, Null, Null, Null, Null, Null, 合作单位_In, 4, v_结算号码);
            Else
              If n_Exists = 1 Then
                v_Err_Msg := '目前挂号仅支持一种三方结算方式,不能继续操作！';
                Raise Err_Item;
              End If;
              Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
              Insert Into 病人预交记录
                (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号,
                 交易说明, 合作单位, 结算性质, 结算号码)
              Values
                (n_预交id, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(v_结算方式, v_现金), Nvl(n_单笔金额, 0), 登记时间_In,
                 v_操作员编号, v_操作员姓名, n_结帐id, '挂号收费', n_组id, 卡类别id_In, Null, 支付卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In, 4, v_结算号码);
              n_Exists := 1;
            End If;
          
            v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '|') + 1);
          End Loop;
        End If;
      End If;
    
      --更新人员缴款数据
    
      For v_缴款 In (Select 结算方式, Sum(Nvl(a.冲预交, 0)) As 冲预交
                   From 病人预交记录 A
                   Where a.结帐id = n_结帐id And Mod(a.记录性质, 10) <> 1 And Nvl(病人id, 0) = Nvl(病人id_In, 0)
                   Group By 结算方式) Loop
      
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) + Nvl(v_缴款.冲预交, 0)
        Where 收款员 = v_操作员姓名 And 性质 = 1 And 结算方式 = v_缴款.结算方式
        Returning 余额 Into n_返回值;
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额
            (收款员, 结算方式, 性质, 余额)
          Values
            (v_操作员姓名, v_缴款.结算方式, 1, Nvl(v_缴款.冲预交, 0));
          n_返回值 := Nvl(v_缴款.冲预交, 0);
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 人员缴款余额
          Where 收款员 = v_操作员姓名 And 结算方式 = v_缴款.结算方式 And 性质 = 1 And Nvl(余额, 0) = 0;
        End If;
      End Loop;
    End If;
  
    --处理挂号记录
    If 锁定类型_In = 2 Then
      Begin
        Select ID Into n_挂号id From 病人挂号记录 Where　记录状态 = 0 And NO = 单据号_In And 病人id = 病人id_In;
      Exception
        When Others Then
          Null;
      End;
    Else
      Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
    End If;
  
    Update 病人挂号记录
    Set 记录性质 = Decode(操作方式_In, 2, 2, 1), 记录状态 = Decode(锁定类型_In, 1, 0, 1), 门诊号 = r_Pati.门诊号, 操作员姓名 = v_操作员姓名,
        操作员编号 = v_操作员编号, 预约 = Decode(操作方式_In, 1, 0, 1),
        接收人 = Decode(锁定类型_In, 1, Null, Decode(操作方式_In, 2, Null, v_操作员姓名)),
        接收时间 = Case 锁定类型_In
                  When 1 Then
                   Null
                  Else
                   Case 操作方式_In
                     When 2 Then
                      Null
                     Else
                      d_登记时间
                   End
                End, 交易流水号 = Nvl(交易流水号_In, 交易流水号), 交易说明 = Nvl(交易说明_In, 交易说明), 合作单位 = Nvl(合作单位_In, 合作单位),
        预约操作员 = Decode(操作方式_In, 1, Nvl(预约操作员, Null), Nvl(预约操作员, v_操作员姓名)),
        预约操作员编号 = Decode(操作方式_In, 1, Nvl(预约操作员编号, Null), Nvl(预约操作员编号, v_操作员编号)), 出诊记录id = 记录id_In
    Where ID = n_挂号id;
    If Sql%NotFound Then
      Begin
        Select 名称 Into v_付款方式 From 医疗付款方式 Where 编码 = r_Pati.付款方式 And Rownum < 2;
      Exception
        When Others Then
          v_付款方式 := Null;
      End;
      Insert Into 病人挂号记录
        (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 预约时间, 操作员编号,
         操作员姓名, 复诊, 号序, 预约, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 医疗付款方式, 预约操作员, 预约操作员编号, 出诊记录id)
      Values
        (n_挂号id, 单据号_In, Decode(操作方式_In, 2, 2, 1), Decode(锁定类型_In, 1, 0, 1), r_Pati.病人id, r_Pati.门诊号, r_Pati.姓名,
         r_Pati.性别, r_Pati.年龄, 号码_In, 0, v_诊室, Null, n_科室id, v_医生姓名, 0, Null, d_登记时间, 发生时间_In,
         Case When(Nvl(操作方式_In, 0)) = 1 Then Null Else 发生时间_In End, v_操作员编号, v_操作员姓名, 0, n_号序, Decode(操作方式_In, 1, 0, 1),
         Decode(操作方式_In, 2, Null, v_操作员姓名), Decode(操作方式_In, 2, To_Date(Null), d_登记时间), 交易流水号_In, 交易说明_In, 合作单位_In,
         v_付款方式, Decode(操作方式_In, 1, Null, v_操作员姓名), Decode(操作方式_In, 1, Null, v_操作员编号), 记录id_In);
    End If;
    --锁定单据不能产生队列
    If 锁定类型_In <> 1 Then
      n_预约生成队列 := 0;
      If 操作方式_In > 1 Then
        n_预约生成队列 := Zl_To_Number(zl_GetSysParameter('预约生成队列', 1113));
      End If;
      --挂号和收费的预约都直接进入队列(收费预约缺少接收过程,所以直接和挂号一样直接进入队列)
      If 操作方式_In <> 2 Or n_预约生成队列 = 1 Then
        If Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113)) <> 0 Then
          --排队叫号模式:-0-不产生队列;1-按医生或分诊台排队;2-先分诊,后医生站
          If Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113)) = 0 Or n_预约生成队列 = 1 Then
            --产生队列
            --.按”执行部门” 的方式生成队列
            v_队列名称 := n_科室id;
            v_排队号码 := Zlgetnextqueue(n_科室id, n_挂号id, 号码_In || '|' || 号序_In);
            --挂号id_In,号码_In,号序_In,缺省日期_In,扩展_In(暂无用)
            v_排队序号 := Zlgetsequencenum(0, n_挂号id, 0);
            --挂号id_In,号码_In,号序_In,缺省日期_In,扩展_In(暂无用)
            d_排队时间 := Zl_Get_Queuedate(n_挂号id, 号码_In, 号序_In, d_Date);
            --  队列名称_In , 业务类型_In, 业务id_In,科室id_In,排队号码_In,v_排队标记,患者姓名_In,病人ID_IN, 诊室_In, 医生姓名_In, 排队时间_In
            Zl_排队叫号队列_Insert(v_队列名称, 0, n_挂号id, n_科室id, v_排队号码, Null, r_Pati.姓名, r_Pati.病人id, v_诊室, v_医生姓名, d_排队时间,
                             预约方式_In, n_启用分时段, v_排队序号);
          End If;
        End If;
      End If;
    
      If Nvl(操作方式_In, 0) = 1 Then
        --处理票据使用情况
        If 票据号_In Is Not Null Then
          Select 票据打印内容_Id.Nextval Into n_打印id From Dual;
          --发出票据
          Insert Into 票据打印内容 (ID, 数据性质, NO) Values (n_打印id, 4, 单据号_In);
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
          Values
            (票据使用明细_Id.Nextval, Decode(收费票据_In, 1, 1, 4), 票据号_In, 1, 1, 领用id_In, n_打印id, d_登记时间, v_操作员姓名);
          --状态改动
          Update 票据领用记录
          Set 当前号码 = 票据号_In, 剩余数量 = Decode(Sign(剩余数量 - 1), -1, 0, 剩余数量 - 1), 使用时间 = Sysdate
          Where ID = Nvl(领用id_In, 0);
        End If;
        --病人本次就诊(以发生时间为准)
        If Nvl(r_Pati.病人id, 0) <> 0 Then
          Update 病人信息 Set 就诊时间 = 发生时间_In, 就诊状态 = 1, 就诊诊室 = v_诊室 Where 病人id = r_Pati.病人id;
        End If;
      End If;
    End If;
    --病人挂号汇总
    --解锁单据时不用再对汇总单据进行统计了 在锁定单据时已经进行了汇总
    If 锁定类型_In <> 2 Then
      --操作方式_IN:1-表示挂号,2-表示预约挂号不扣款,3-表示预约挂号,扣款
      --是否为预约接收:0-非预约挂号; 1-预约挂号,2-预约接收;3-收费预约
      n_预约 := Case
                When Nvl(操作方式_In, 0) = 1 Then
                 0
                When Nvl(操作方式_In, 0) = 2 Then
                 1
                When Nvl(操作方式_In, 0) = 3 Then
                 3
                Else
                 0
              End;
      Zl_病人挂号汇总_Update(v_医生姓名, n_医生id, n_项目id, n_科室id, 发生时间_In, n_预约, 号码_In, 0, 记录id_In);
    End If;
    --消息推送
    Begin
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 1, n_挂号id;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_001(n_挂号id, 单据号_In);
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Err_Msg || '+[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End;

Begin
  n_挂号排班模式 := Nvl(zl_GetSysParameter('挂号排班模式'), 0);
  If n_挂号排班模式 = 1 Then
    --出诊表排班模式
    Zl_三方机构挂号_出诊_Insert(出诊记录id_In, 操作方式_In, 病人id_In, 号码_In, 号序_In, 单据号_In, 票据号_In, 结算方式_In, 摘要_In, 发生时间_In, 登记时间_In,
                        合作单位_In, 挂号金额合计_In, 领用id_In, 收费票据_In, 交易流水号_In, 交易说明_In, 预约方式_In, 预交id_In, 卡类别id_In, 加入序号状态_In,
                        是否自助设备_In, 结帐id_In, 锁定类型_In, 保险结算_In, 冲预交_In, 支付卡号_In, 退号重用_In, 费别_In, 冲预交病人ids_In, 机器名_In,
                        更新年龄_In, 购买病历_In);
  Else
  
    v_冲预交病人ids := Nvl(冲预交病人ids_In, 病人id_In);
    v_Temp          := zl_GetSysParameter(256);
    If v_Temp Is Null Or Substr(v_Temp, 1, 1) = '0' Then
      Null;
    Else
      Begin
        d_启用时间 := To_Date(Substr(v_Temp, 3), 'YYYY-MM-DD hh24:mi:ss');
        If 发生时间_In > d_启用时间 Then
          v_Err_Msg := '当前挂号的发生时间' || To_Char(发生时间_In, 'yyyy-mm-dd hh24:mi:ss') || '已经启用了出诊表排班模式,不能再使用计划排班模式挂号!';
          Raise Err_Item;
        End If;
      Exception
        When Others Then
          Null;
      End;
    End If;
    If 费别_In Is Null Then
      Select Zl_Custom_Getpatifeetype(1, 病人id_In) Into v_费别 From Dual;
    Else
      v_费别 := 费别_In;
    End If;
    If v_费别 Is Null Then
      n_屏蔽费别 := 1;
      Select 名称 Into v_费别 From 费别 Where 缺省标志 = 1 And Rownum < 2;
    End If;
    Update 病人信息 Set 费别 = v_费别 Where 病人id = 病人id_In;
  
    If 更新年龄_In = 1 Then
      Select Zl_Age_Calc(病人id_In) Into v_年龄 From Dual;
      If v_年龄 Is Not Null Then
        Update 病人信息 Set 年龄 = v_年龄 Where 病人id = 病人id_In;
      End If;
    End If;
    --获取当前机器名称
    If 机器名_In Is Not Null Then
      v_机器名 := 机器名_In;
    Else
      Select Terminal Into v_机器名 From V$session Where Audsid = Userenv('sessionid');
    End If;
    n_实收金额合计 := 0;
    Select Count(*) + 1
    Into n_挂号序号
    From 病人挂号记录
    Where 号别 = 号码_In And 登记时间 Between Trunc(发生时间_In) And Trunc(发生时间_In + 1) - 1 / 24 / 60 / 60;
    --Begin
    --部门ID,部门名称;人员ID,人员编号,人员姓名
    v_Temp := Zl_Identity(0);
    If Nvl(v_Temp, ' ') = ' ' Then
      v_Err_Msg := '当前操作人员未设置对应的人员关系,不能继续。';
      Raise Err_Item;
    End If;
  
    If 登记时间_In Is Null Then
      d_登记时间 := Sysdate;
    Else
      d_登记时间 := 登记时间_In;
    End If;
    If Trunc(Sysdate) > Trunc(发生时间_In) Then
      v_Err_Msg := '不能挂以前的号(' || To_Char(发生时间_In, 'yyyy-mm-dd') || ')。';
      Raise Err_Item;
    End If;
    n_同科限约一个号 := Nvl(zl_GetSysParameter('病人同科限约一个号', 1111), 0);
    n_病人预约科室数 := Nvl(zl_GetSysParameter('病人预约科室数', 1111), 0);
    n_开单部门id     := To_Number(Zl_操作员(0, v_Temp));
    v_操作员编号     := Zl_操作员(1, v_Temp);
    v_操作员姓名     := Zl_操作员(2, v_Temp);
    n_组id           := Zl_Get组id(v_操作员姓名);
  
    If 操作方式_In <> 1 Then
      --预约检查是否添加合作单位控制
      --如果设置了合作单位控制 则
      Begin
        Select ID
        Into n_计划id
        From 挂号安排计划
        Where 号码 = 号码_In And 发生时间_In Between Nvl(生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
              Nvl(失效时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Rownum < 2
        Order By 生效时间 Desc;
      Exception
        When Others Then
          Select ID Into n_Tmp安排id From 挂号安排 Where 号码 = 号码_In;
      End;
      If Nvl(n_计划id, 0) <> 0 Then
        Select Count(0)
        Into n_合作单位限制
        From 合作单位计划控制
        Where 合作单位 = 合作单位_In And 计划id = n_计划id And Rownum < 2;
      Else
        Select Count(0)
        Into n_合作单位限制
        From 合作单位安排控制
        Where 合作单位 = 合作单位_In And 安排id = n_Tmp安排id And Rownum < 2;
      End If;
    End If;
  
    If 操作方式_In <> 2 Then
      v_诊室 := Zl_诊室(号码_In);
    End If;
    If 操作方式_In <> 2 And 结算方式_In Is Not Null Then
      --检查结算方式是否完备
      Select Count(*) Into n_Count From 结算方式 Where 名称 = Nvl(结算方式_In, 'Lxh') And 性质 In (2, 7, 8);
      If Nvl(卡类别id_In, 0) <> 0 And n_Count = 0 Then
        Select Count(1)
        Into n_Count
        From 医疗卡类别
        Where ID = Nvl(卡类别id_In, 0) And 结算方式 = Nvl(结算方式_In, 'lxh');
      End If;
      If n_Count = 0 Then
        v_Err_Msg := '结算方式(' || 结算方式_In || ')未设置,请在结算方式管理中设置。';
        Raise Err_Item;
      End If;
    End If;
  
    --因为现在有按日编单据号规则,日挂号量不能超过10000次,所以要检查唯一约束。
    Select Count(*) Into n_Count From 门诊费用记录 Where 记录性质 = 4 And 记录状态 In (1, 3) And NO = 单据号_In;
    If n_Count <> 0 Then
      v_Err_Msg := '挂号单据号重复,不能保存！' || Chr(13) || '如果使用了按日顺序编号,当日挂号量不能超过10000人次。';
      Raise Err_Item;
    End If;
  
    Open c_Pati(病人id_In);
    n_Count := 0;
    Begin
      Fetch c_Pati
        Into r_Pati;
    Exception
      When Others Then
        n_Count := -1;
    End;
    If n_Count = -1 Then
      v_Err_Msg := '病人未找到，不能继续。';
      Raise Err_Item;
    End If;
  
    Open c_安排(号码_In, 发生时间_In);
    Begin
      Fetch c_安排
        Into r_安排;
    Exception
      When Others Then
        n_Count := -1;
    End;
    If n_Count = -1 Then
      v_Err_Msg := '该号别没有在' || To_Char(发生时间_In, 'yyyy-mm-dd hh24:mi:ss') || '中进行安排。';
      Raise Err_Item;
    End If;
  
    Select Decode(To_Char(发生时间_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六',
                   '周日')
    Into v_星期
    From Dual;
    Begin
      If r_安排.计划id Is Null Then
        Select 1 Into n_启用分时段 From 挂号安排时段 Where 安排id = r_安排.Id And 星期 = v_星期 And Rownum < 2;
      Else
        Select 1 Into n_启用分时段 From 挂号计划时段 Where 计划id = r_安排.计划id And 星期 = v_星期 And Rownum < 2;
      End If;
    Exception
      When Others Then
        n_启用分时段 := 0;
    End;
  
    --对参数控制进行检查
    --仅在预约不扣款时进行检查
    If 操作方式_In = 2 Then
      If Nvl(n_同科限约一个号, 0) <> 0 Or Nvl(n_病人预约科室数, 0) <> 0 Then
        n_已约科室 := 0;
        For c_Chkitem In (Select Count(1) As 已约, a.执行部门id As 科室id, Nvl(k.名称, '') As 科室
                          From 病人挂号记录 A, 病人信息 B, 部门表 K
                          Where a.病人id = b.病人id And a.病人id = 病人id_In And a.执行部门id = k.Id(+) And a.记录性质 = 2 And 记录状态 = 1 And
                                a.预约时间 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60
                          Group By a.执行部门id, k.名称) Loop
          If Nvl(n_同科限约一个号, 0) <> 0 And c_Chkitem.科室id = r_安排.科室id Then
          
            v_Err_Msg := '该病人已经在科室[' || c_Chkitem.科室 || ']进行了预约,不能再预约！';
            Raise Err_Item;
          
            If Nvl(n_病人预约科室数, 0) > 0 And c_Chkitem.科室id <> r_安排.科室id Then
              n_已约科室 := n_已约科室 + 1;
            End If;
          End If;
        End Loop;
        If n_已约科室 >= Nvl(n_病人预约科室数, 0) And Nvl(n_病人预约科室数, 0) > 0 Then
          v_Err_Msg := '同一病人在最多同时预约[' || Nvl(n_病人预约科室数, 0) || ']个科室,不能再预约！';
          Raise Err_Item;
        End If;
      End If;
    End If;
  
    d_Date         := Null;
    d_时段开始时间 := Null;
  
    If Nvl(r_安排.限号数, 0) >= 0 Or r_安排.限号数 Is Null Then
    
      Select Nvl(Sum(Nvl(b.已挂数, 0)), 0), Nvl(Sum(Nvl(b.其中已接收, 0)), 0), Nvl(Sum(Nvl(b.已约数, 0)), 0)
      Into n_已挂数, n_其中已接收, n_已约数
      From 挂号安排 A, 病人挂号汇总 B
      Where a.科室id = b.科室id And a.项目id = b.项目id And a.号码 = 号码_In And b.日期 Between Trunc(发生时间_In) And
            Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And (a.号码 = b.号码 Or b.号码 Is Null) And Nvl(a.医生id, 0) = Nvl(b.医生id, 0) And
            Nvl(a.医生姓名, '医生') = Nvl(b.医生姓名, '医生');
    
      If n_启用分时段 = 1 Then
        If Nvl(r_安排.序号控制, 0) = 1 Then
          If Nvl(是否自助设备_In, 0) = 0 Then
            If r_安排.计划id Is Null Then
              Select Count(*), Max(开始时间)
              Into n_Count, d_时段开始时间
              From 挂号安排时段
              Where 安排id = r_安排.Id And 星期 = v_星期 And 序号 = Nvl(号序_In, 0);
            Else
              Select Count(*), Max(开始时间)
              Into n_Count, d_时段开始时间
              From 挂号计划时段
              Where 计划id = r_安排.计划id And 星期 = v_星期 And 序号 = Nvl(号序_In, 0);
            End If;
            v_Temp := '挂号';
            If 操作方式_In > 1 Then
              v_Temp := '预约挂号';
            End If;
          
            If n_Count = 0 Then
              v_Err_Msg := '号别为' || 号码_In || '的挂号安排中不存在序号为' || Nvl(号序_In, 0) || '的安排,不能再' || v_Temp || '！';
              Raise Err_Item;
            End If;
          End If;
          --过点的,不能选择挂号
          If Trunc(Sysdate) = Trunc(发生时间_In) Then
            --挂当天的号
            v_Temp := To_Char(Sysdate, 'yyyy-mm-dd') || ' ';
            If r_安排.计划id Is Null Then
              For v_时段 In (Select To_Date(v_Temp || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                                  To_Date(To_Char(Sysdate + Decode(Sign(开始时间 - 结束时间), 1, 1, 0), 'yyyy-mm-dd') || ' ' ||
                                           To_Char(结束时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 结束时间, 限制数量, 是否预约
                           From 挂号安排时段
                           Where 安排id = r_安排.Id And 星期 = v_星期 And 序号 = Nvl(号序_In, 0)) Loop
                If Sysdate > v_时段.结束时间 Then
                  v_Err_Msg := '号别为' || 号码_In || '及号序为' || Nvl(号序_In, 0) || '的安排,已经超过时点,不能再' || v_Temp || '！';
                  Raise Err_Item;
                End If;
              End Loop;
            Else
              For v_时段 In (Select To_Date(v_Temp || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                                  To_Date(To_Char(Sysdate + Decode(Sign(开始时间 - 结束时间), 1, 1, 0), 'yyyy-mm-dd') || ' ' ||
                                           To_Char(结束时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 结束时间, 限制数量, 是否预约
                           From 挂号计划时段
                           Where 计划id = r_安排.计划id And 星期 = v_星期 And 序号 = Nvl(号序_In, 0)) Loop
                If Sysdate > v_时段.结束时间 Then
                  v_Err_Msg := '号别为' || 号码_In || '及号序为' || Nvl(号序_In, 0) || '的安排,已经超过时点,不能再' || v_Temp || '！';
                  Raise Err_Item;
                End If;
              End Loop;
            End If;
          End If;
        Elsif 操作方式_In > 1 Then
          --未启用序号的,需要检查预约的情况
          n_Count := 0;
          If r_安排.计划id Is Null Then
            For v_时段 In (Select 序号, 开始时间, 结束时间, 限制数量, 是否预约
                         From 挂号安排时段
                         Where 安排id = r_安排.Id And 星期 = v_星期 And
                               (('3000-01-10 ' || To_Char(发生时间_In, 'HH24:MI:SS') Between
                               Decode(Sign(开始时间 - 结束时间 - 1 / 24 / 60 / 60), 1,
                                        '3000-01-09 ' || To_Char(开始时间, 'HH24:MI:SS'),
                                        '3000-01-10 ' || To_Char(开始时间, 'HH24:MI:SS')) And
                               '3000-01-10 ' || To_Char(结束时间 - 1 / 24 / 60 / 60, 'HH24:MI:SS')) Or
                               ('3000-01-10 ' || To_Char(发生时间_In, 'HH24:MI:SS') Between
                               '3000-01-10 ' || To_Char(开始时间, 'HH24:MI:SS') And
                               Decode(Sign(开始时间 - 结束时间 - 1 / 24 / 60 / 60), 1,
                                        '3000-01-11 ' || To_Char(结束时间 - 1 / 24 / 60 / 60, 'HH24:MI:SS'),
                                        '3000-01-10 ' || To_Char(结束时间 - 1 / 24 / 60 / 60, 'HH24:MI:SS'))))) Loop
              n_预约时段序号 := v_时段.序号;
              d_时段开始时间 := v_时段.开始时间;
            
              Select Count(*), Max(序号)
              Into n_Count, n_预约总数
              From 挂号序号状态
              Where 号码 = 号码_In And 日期 = 发生时间_In And 状态 Not In (4, 5);
            
              If Nvl(n_Count, 0) > Nvl(v_时段.限制数量, 0) And 锁定类型_In <> 2 Then
                v_Err_Msg := '号别为' || 号码_In || '的挂号安排中在' || To_Char(v_时段.开始时间, 'hh24:mi:ss') || '至' ||
                             To_Char(v_时段.结束时间, 'hh24:mi:ss') || '最多只能预约' || Nvl(v_时段.限制数量, 0) || '人,不能再进行预约挂号！';
                Raise Err_Item;
              End If;
              n_Count := 1;
            End Loop;
          Else
            For v_时段 In (Select 序号, 开始时间, 结束时间, 限制数量, 是否预约
                         From 挂号计划时段
                         Where 计划id = r_安排.计划id And 星期 = v_星期 And
                               (('3000-01-10 ' || To_Char(发生时间_In, 'HH24:MI:SS') Between
                               Decode(Sign(开始时间 - 结束时间 - 1 / 24 / 60 / 60), 1,
                                        '3000-01-09 ' || To_Char(开始时间, 'HH24:MI:SS'),
                                        '3000-01-10 ' || To_Char(开始时间, 'HH24:MI:SS')) And
                               '3000-01-10 ' || To_Char(结束时间 - 1 / 24 / 60 / 60, 'HH24:MI:SS')) Or
                               ('3000-01-10 ' || To_Char(发生时间_In, 'HH24:MI:SS') Between
                               '3000-01-10 ' || To_Char(开始时间, 'HH24:MI:SS') And
                               Decode(Sign(开始时间 - 结束时间 - 1 / 24 / 60 / 60), 1,
                                        '3000-01-11 ' || To_Char(结束时间 - 1 / 24 / 60 / 60, 'HH24:MI:SS'),
                                        '3000-01-10 ' || To_Char(结束时间 - 1 / 24 / 60 / 60, 'HH24:MI:SS'))))) Loop
              n_预约时段序号 := v_时段.序号;
              d_时段开始时间 := v_时段.开始时间;
            
              Select Count(*), Max(序号)
              Into n_Count, n_预约总数
              From 挂号序号状态
              Where 号码 = 号码_In And 日期 = 发生时间_In And 状态 Not In (4, 5);
            
              If Nvl(n_Count, 0) > Nvl(v_时段.限制数量, 0) And 锁定类型_In <> 2 Then
                v_Err_Msg := '号别为' || 号码_In || '的挂号安排中在' || To_Char(v_时段.开始时间, 'hh24:mi:ss') || '至' ||
                             To_Char(v_时段.结束时间, 'hh24:mi:ss') || '最多只能预约' || Nvl(v_时段.限制数量, 0) || '人,不能再进行预约挂号！';
                Raise Err_Item;
              End If;
              n_Count := 1;
            End Loop;
          End If;
        
          If n_Count = 0 Then
            v_Err_Msg := '号别为' || 号码_In || '的挂号安排中没有相关的安排计划(' || To_Char(发生时间_In, 'YYYY-mm-dd HH24:MI:SS') ||
                         '),不能进行预约挂号！';
            Raise Err_Item;
          End If;
        End If;
      End If;
    End If;
  
    If 操作方式_In = 1 And 锁定类型_In <> 2 Then
      --挂号规则:
      --  已挂数不能大于限号数
      If n_已挂数 >= Nvl(r_安排.限号数, 0) And r_安排.限号数 Is Not Null Then
        v_Err_Msg := '该号别今天已达到限号数 ' || Nvl(r_安排.限号数, 0) || '不能再挂号！';
        Raise Err_Item;
      End If;
    End If;
  
    If 操作方式_In > 1 Then
      --预约的相关检查
      --规则:
      --   1.已限约不能超过限约数
      --   2.检查是否启用时段的
      If n_已约数 >= Nvl(r_安排.限约数, 0) And Nvl(r_安排.限约数, 0) <> 0 And r_安排.限约数 Is Not Null And 锁定类型_In <> 2 Then
        v_Err_Msg := '该号别已达到限约数 ' || Nvl(r_安排.限约数, 0) || '不能再预约挂号！';
        Raise Err_Item;
      End If;
    End If;
    If n_合作单位限制 > 0 And 操作方式_In <> 1 And 合作单位_In Is Not Null Then
    
      If Nvl(r_安排.序号控制, 0) = 1 And Nvl(号序_In, 0) = 0 Then
        v_Err_Msg := '当前安排使用了序号控制,请确认所需要预约的序号,不能继续。';
        Raise Err_Item;
      End If; --Nvl(r_安排.序号控制, 0) =0
    
      n_序号 := Case
                When Nvl(r_安排.序号控制, 0) = 1 Or n_启用分时段 = 1 And 操作方式_In > 1 Then
                 Nvl(号序_In, 0)
                Else
                 0
              End;
    
      --合作单位限数量模式
      Begin
        If Nvl(n_计划id, 0) <> 0 Then
          Select 0
          Into n_序号
          From 合作单位计划控制
          Where 合作单位 = 合作单位_In And 计划id = n_计划id And
                限制项目 = Decode(To_Char(发生时间_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                              '7', '周六', Null) And 数量 <> 0 And 序号 = 0 And Rownum < 2;
        Else
          Select 0
          Into n_序号
          From 合作单位安排控制
          Where 合作单位 = 合作单位_In And 安排id = n_Tmp安排id And
                限制项目 = Decode(To_Char(发生时间_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                              '7', '周六', Null) And 数量 <> 0 And 序号 = 0 And Rownum < 2;
        End If;
        n_合作单位限数量模式 := 1;
      Exception
        When Others Then
          n_合作单位限数量模式 := 0;
      End;
      --开放序号检查
      For c_合作单位 In (Select c.序号, 数量
                     From 挂号安排 A, 合作单位安排控制 C
                     Where a.号码 = 号码_In And Decode(To_Char(发生时间_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                                   '周四', '6', '周五', '7', '周六', Null) = c.限制项目(+) And a.Id = c.安排id And
                           c.合作单位 = 合作单位_In And c.序号 = n_序号 And Not Exists
                      (Select 1
                            From 挂号安排计划 D
                            Where d.安排id = a.Id And d.审核时间 Is Not Null And
                                  发生时间_In Between Nvl(d.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                                  Nvl(d.失效时间, To_Date('3000-01-01', 'yyyy-mm-dd')))
                     Union All
                     Select c.序号, 数量
                     From 挂号安排计划 A, 挂号安排 D, 合作单位计划控制 C,
                          (Select Max(a.生效时间) As 生效, 安排id
                            From 挂号安排计划 A, 挂号安排 B
                            Where a.安排id = b.Id And a.审核时间 Is Not Null And
                                  发生时间_In Between Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                                  Nvl(a.失效时间, To_Date('3000-01-01', 'yyyy-mm-dd')) And b.号码 = 号码_In
                            Group By 安排id) E
                     Where a.安排id = d.Id And a.审核时间 Is Not Null And d.号码 = 号码_In And a.安排id = e.安排id And
                           Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) =
                           Nvl(e.生效, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                           Decode(To_Char(发生时间_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                                  '7', '周六', Null) = c.限制项目(+) And a.Id = c.计划id And c.合作单位 = 合作单位_In And c.序号 = n_序号 And
                           发生时间_In Between Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                           Nvl(a.失效时间, To_Date('3000-01-01', 'yyyy-mm-dd'))) Loop
      
        If Nvl(r_安排.序号控制, 0) = 1 And c_合作单位.序号 = n_序号 And n_合作单位限数量模式 = 0 Then
          n_是否开放 := 1;
          Exit;
        Elsif (Nvl(r_安排.序号控制, 0) = 0 And c_合作单位.序号 = n_序号) Or n_合作单位限数量模式 = 1 Then
          Begin
            Select Nvl(已约数, 0)
            Into n_预约数量
            From 合作单位挂号汇总
            Where 合作单位 = 合作单位_In And 日期 = Trunc(发生时间_In) And 号码 = 号码_In;
          Exception
            When Others Then
              n_预约数量 := 0;
          End;
          If c_合作单位.数量 <= n_预约数量 And Nvl(c_合作单位.数量, 0) > 0 And 锁定类型_In <> 2 Then
            v_Err_Msg := '该号别已达到限约数 ' || Nvl(c_合作单位.数量, 0) || '不能再预约挂号！';
            Raise Err_Item;
          End If;
          n_是否开放 := 1;
          Exit;
        End If;
      
      End Loop;
    
      If Nvl(n_是否开放, 0) = 0 Then
        v_Err_Msg := '当前序号(' || Nvl(号序_In, 0) || '未开放,不能继续。';
        Raise Err_Item;
      End If;
    End If;
  
    --检查限号数和限约数
    n_行号         := 1;
    n_原项目id     := 0;
    n_原收入项目id := 0;
    n_实收金额合计 := 0;
    If 锁定类型_In <> 1 Then
      If 操作方式_In <> 2 Then
        If Nvl(结帐id_In, 0) = 0 Then
          --这里应该程序传入
          Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
        Else
          n_结帐id := 结帐id_In;
        End If;
      Else
        n_结帐id := Null;
      End If;
    End If;
  
    If Nvl(购买病历_In, 0) = 1 Then
      Begin
        Select 收费细目id Into n_病历费id From 收费特定项目 Where 特定项目 = '病历费';
        v_收费项目ids := r_安排.项目id || ',' || n_病历费id;
      Exception
        When Others Then
          v_Err_Msg := '不能确定病历费,挂号失败!';
          Raise Err_Item;
      End;
    Else
      v_收费项目ids := r_安排.项目id;
    End If;
  
    For c_Item In (Select 1 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, 1 As 数次,
                          c.Id As 收入项目id, c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, Null As 从属父号
                   From 收费项目目录 A, 收费价目 B, 收入项目 C
                   Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = r_安排.项目id And Sysdate Between b.执行日期 And
                         Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))
                   Union All
                   Select 2 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, 1 As 数次,
                          c.Id As 收入项目id, c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, Null As 从属父号
                   From 收费项目目录 A, 收费价目 B, 收入项目 C
                   Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = n_病历费id And Sysdate Between b.执行日期 And
                         Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))
                   Union All
                   Select 3 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, d.从项数次 As 数次,
                          c.Id As 收入项目id, c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, 1 As 从属父号
                   From 收费项目目录 A, 收费价目 B, 收入项目 C, 收费从属项目 D
                   Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = d.从项id And
                         d.主项id In (Select Column_Value From Table(f_Str2list(v_收费项目ids))) And Sysdate Between b.执行日期 And
                         Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))
                   Order By 性质, 项目编码, 收入编码) Loop
      n_价格父号 := Null;
      If n_原项目id = c_Item.项目id Then
        If n_原收入项目id <> c_Item.收入项目id Then
          n_价格父号 := n_行号;
        End If;
        n_原收入项目id := c_Item.收入项目id;
      End If;
      n_原项目id := c_Item.项目id;
      n_应收金额 := Round(c_Item.数次 * c_Item.单价, 5);
      n_实收金额 := n_应收金额;
      If Nvl(c_Item.屏蔽费别, 0) <> 1 And n_屏蔽费别 = 0 Then
        --打折:
        v_Temp     := Zl_Actualmoney(r_Pati.费别, c_Item.项目id, c_Item.收入项目id, n_应收金额);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ':') + 1);
        n_实收金额 := Zl_To_Number(v_Temp);
      End If;
      n_实收金额合计 := Nvl(n_实收金额合计, 0) + n_实收金额;
    
      --锁定单据不产生费用
      If 锁定类型_In <> 1 Then
        --产生病人挂号费用(可能单独是或包括病历费用)
        Select 病人费用记录_Id.Nextval Into n_费用id From Dual;
        --:操作方式_IN:1-表示挂号,2-表示预约挂号不扣款,3-表示预约挂号,扣款
        Insert Into 门诊费用记录
          (ID, 记录性质, 记录状态, 序号, 价格父号, 从属父号, NO, 实际票号, 门诊标志, 加班标志, 附加标志, 发药窗口, 病人id, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 病人科室id,
           收费类别, 计算单位, 收费细目id, 收入项目id, 收据费目, 付数, 数次, 标准单价, 应收金额, 实收金额, 结帐金额, 结帐id, 记帐费用, 开单部门id, 开单人, 划价人, 执行部门id, 执行人,
           操作员编号, 操作员姓名, 发生时间, 登记时间, 保险大类id, 保险项目否, 保险编码, 统筹金额, 摘要, 结论, 缴款组id)
        Values
          (n_费用id, 4, Decode(操作方式_In, 2, 0, 1), n_行号, n_价格父号, c_Item.从属父号, 单据号_In, 票据号_In, 1, Null, Null,
           Decode(操作方式_In, 2, To_Char(号序_In), v_诊室), r_Pati.病人id, r_Pati.门诊号, r_Pati.付款方式, r_Pati.姓名, r_Pati.性别,
           r_Pati.年龄, r_Pati.费别, r_安排.科室id, c_Item.类别, 号码_In, c_Item.项目id, c_Item.收入项目id, c_Item.收据费目, 1, c_Item.数次,
           c_Item.单价, n_应收金额, n_实收金额, Decode(操作方式_In, 2, Null, n_实收金额), n_结帐id, 0, n_开单部门id, v_操作员姓名,
           Decode(操作方式_In, 2, v_操作员姓名, Null), r_安排.科室id, r_安排.医生姓名, v_操作员编号, v_操作员姓名, 发生时间_In, d_登记时间, Null, 0, Null,
           Null, 摘要_In, 预约方式_In, Decode(操作方式_In, 2, Null, n_组id));
      End If;
      n_行号 := n_行号 + 1;
    
    End Loop;
  
    If Round(Nvl(挂号金额合计_In, 0), 5) <> Round(Nvl(n_实收金额合计, 0), 5) Then
      v_Err_Msg := '本次挂号金额不正确,可能是因为医院调整了价格,请重新获取挂号收费项目的价格,不能继续。';
      Raise Err_Item;
    End If;
  
    If n_启用分时段 = 1 Then
      d_Date := To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(d_时段开始时间, 'hh24:mi:ss'),
                        'yyyy-mm-dd hh24:mi:ss');
    Else
      d_Date := Trunc(发生时间_In);
    End If;
  
    --更新挂号序号状态
    If 锁定类型_In <> 2 Then
      n_号序 := 号序_In;
    End If;
    Begin
      Select 1
      Into n_Count
      From 挂号序号状态
      Where Trunc(日期) = Trunc(发生时间_In) And 号码 = 号码_In And 序号 = n_号序 And 状态 <> 5;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count = 1 Then
      If n_启用分时段 = 0 And Nvl(r_安排.序号控制, 0) = 1 Then
        n_号序 := Null;
      End If;
      If n_启用分时段 = 1 And Nvl(r_安排.序号控制, 0) = 1 Then
        v_Err_Msg := '当前序号已被使用，请重新选择一个序号！';
        Raise Err_Item;
      End If;
    End If;
    n_Count := 0;
    If n_启用分时段 = 0 And Nvl(r_安排.序号控制, 0) = 1 And n_号序 Is Null And 锁定类型_In <> 2 Then
      If 退号重用_In = 1 Then
        Select Nvl(Max(序号), 0) + 1
        Into n_号序
        From 挂号序号状态
        Where 日期 = Trunc(发生时间_In) And 号码 = r_安排.号码 And 状态 Not In (4, 5);
      Else
        Select Nvl(Max(序号), 0) + 1
        Into n_号序
        From 挂号序号状态
        Where 日期 = Trunc(发生时间_In) And 号码 = r_安排.号码 And 状态 <> 5;
      End If;
    End If;
    If n_启用分时段 = 1 And 锁定类型_In <> 2 Then
    
      If 操作方式_In > 1 And Nvl(r_安排.序号控制, 0) = 0 Then
        --规则:预约时段序号||预约数
        If Nvl(n_预约总数, 0) = 0 Then
          v_Temp := Nvl(r_安排.限约数, 0);
          v_Temp := LTrim(RTrim(v_Temp));
          v_Temp := LPad(Nvl(n_预约总数, 0) + 1, Length(v_Temp), '0');
          v_Temp := n_预约时段序号 || v_Temp;
          n_号序 := To_Number(v_Temp);
        Else
          n_号序 := n_预约总数 + 1;
        End If;
      End If;
    End If;
  
    If Nvl(r_安排.序号控制, 0) = 1 Or (操作方式_In > 1 And n_启用分时段 = 1) Or 加入序号状态_In = 1 Then
      --锁定序号的处理
      Begin
        Select 操作员姓名, 机器名
        Into v_序号操作员, v_序号机器名
        From 挂号序号状态
        Where 状态 = 5 And 号码 = 号码_In And Trunc(日期) = Trunc(d_Date) And 序号 = n_号序;
        n_序号锁定 := 1;
      Exception
        When Others Then
          v_序号操作员 := Null;
          v_序号机器名 := Null;
          n_序号锁定   := 0;
      End;
      If n_序号锁定 = 0 Then
        Update 挂号序号状态
        Set 状态 = Decode(操作方式_In, 2, 2, 1), 预约 = Decode(操作方式_In, 1, 0, 1), 登记时间 = Sysdate
        Where 号码 = 号码_In And 日期 = d_Date And 序号 = n_号序 And 操作员姓名 = v_操作员姓名;
        If Sql%RowCount = 0 Then
          Begin
            Insert Into 挂号序号状态
              (号码, 日期, 序号, 状态, 操作员姓名, 预约, 登记时间)
            Values
              (号码_In, d_Date, n_号序, Decode(操作方式_In, 2, 2, 1), v_操作员姓名, Decode(操作方式_In, 1, 0, 1), Sysdate);
          
            If n_合作单位限制 > 0 And 操作方式_In > 1 And Nvl(n_是否开放, 0) = 1 Then
              Update 合作单位挂号汇总
              Set 已约数 = 已约数 + Decode(操作方式_In, 2, 1, 0), 已接数 = 已接数 + Decode(操作方式_In, 3, 1, 0)
              Where 号码 = 号码_In And 日期 = d_Date And 序号 = n_号序 And 合作单位 = 合作单位_In;
              If Sql%NotFound Then
                Insert Into 合作单位挂号汇总
                  (号码, 日期, 序号, 合作单位, 已约数, 已接数)
                Values
                  (号码_In, d_Date, n_号序, 合作单位_In, Decode(操作方式_In, 1, 0, 1), Decode(操作方式_In, 3, 1, 0));
              End If;
            End If;
          Exception
            When Others Then
              v_Err_Msg := '序号' || n_号序 || '已被使用,请重新选择一个序号.';
              Raise Err_Item;
          End;
        End If;
      Else
        If v_操作员姓名 <> v_序号操作员 Or v_机器名 <> v_序号机器名 Then
          v_Err_Msg := '序号' || n_号序 || '已被自助机' || v_机器名 || '锁定,请重新选择一个序号.';
          Raise Err_Item;
        Else
          Update 挂号序号状态
          Set 状态 = Decode(操作方式_In, 2, 2, 1), 预约 = Decode(操作方式_In, 1, 0, 1), 登记时间 = Sysdate
          Where 号码 = 号码_In And Trunc(日期) = Trunc(d_Date) And 序号 = n_号序 And 状态 = 5 And 操作员姓名 = v_操作员姓名 And 机器名 = v_机器名;
        End If;
      End If;
    End If;
  
    --锁定单据不产生任何 费用
    If 操作方式_In <> 2 And 锁定类型_In <> 1 Then
      --挂号,预约挂号已经扣款部分
      n_预交id := 预交id_In;
      If Nvl(n_预交id, 0) = 0 Then
        Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      End If;
      n_结算合计 := 0;
      If 保险结算_In Is Not Null Then
        --各个保险结算
        v_结算内容 := 保险结算_In || '||';
        n_结算合计 := 0;
        While v_结算内容 Is Not Null Loop
          v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
          v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
          n_结算金额 := To_Number(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
          If Nvl(n_结算金额, 0) <> 0 Then
            Insert Into 病人预交记录
              (ID, 记录性质, NO, 记录状态, 病人id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 结算性质)
            Values
              (n_预交id, 4, 单据号_In, 1, Decode(病人id_In, 0, Null, 病人id_In), '保险结算', v_结算方式, d_登记时间, v_操作员编号, v_操作员姓名,
               n_结算金额, n_结帐id, n_组id, n_结帐id, 4);
            Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          End If;
          n_结算合计 := Nvl(n_结算合计, 0) + Nvl(n_结算金额, 0);
          v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
        End Loop;
      End If;
    
      If Nvl(冲预交_In, 0) <> 0 Then
        --处理总预交
        n_结算合计 := n_结算合计 + Nvl(冲预交_In, 0);
        n_预交金额 := 冲预交_In;
        For r_Deposit In c_Deposit(病人id_In, v_冲预交病人ids) Loop
          n_结算金额 := Case
                      When r_Deposit.金额 - n_预交金额 < 0 Then
                       r_Deposit.金额
                      Else
                       n_预交金额
                    End;
          If r_Deposit.Id <> 0 Then
            --第一次冲预交(填上结帐ID,金额为0)
            Update 病人预交记录 Set 冲预交 = 0, 结帐id = n_结帐id, 结算性质 = 4 Where ID = r_Deposit.Id;
          End If;
          --冲上次剩余额
          Insert Into 病人预交记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 预交类别, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名,
             操作员编号, 冲预交, 结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 结算性质)
            Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 预交类别, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行,
                   单位帐号, d_登记时间, v_操作员姓名, v_操作员编号, n_结算金额, n_结帐id, n_组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, n_结帐id, 4
            From 病人预交记录
            Where NO = r_Deposit.No And 记录状态 = r_Deposit.记录状态 And 记录性质 In (1, 11) And Rownum = 1;
        
          --更新病人预交余额
          Update 病人余额
          Set 预交余额 = Nvl(预交余额, 0) - n_结算金额
          Where 病人id = r_Deposit.病人id And 性质 = 1 And 类型 = Nvl(r_Deposit.预交类别, 2)
          Returning 预交余额 Into n_返回值;
          If Sql%RowCount = 0 Then
            Insert Into 病人余额 (病人id, 预交余额, 性质, 类型) Values (r_Deposit.病人id, -1 * n_结算金额, 1, 1);
            n_返回值 := -1 * n_结算金额;
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 病人余额
            Where 病人id = r_Deposit.病人id And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
          End If;
        
          --检查是否已经处理完
          If r_Deposit.金额 <= n_结算金额 Then
            n_预交金额 := n_预交金额 - r_Deposit.金额;
          Else
            n_预交金额 := 0;
          End If;
          If n_预交金额 = 0 Then
            Exit;
          End If;
        End Loop;
        If n_预交金额 > 0 Then
          v_Err_Msg := '预交余额不够支付本次支付金额,不能继续操作！';
          Raise Err_Item;
        End If;
      End If;
      --剩余款项,用指定结算方支付
      n_结算金额 := Nvl(n_实收金额合计, 0) - Nvl(n_结算合计, 0);
      If Nvl(n_结算金额, 0) < 0 Then
        v_Err_Msg := '挂号的相关结算金额超出了当前实结金额,不能继续操作！';
        Raise Err_Item;
      End If;
      If Nvl(n_结算金额, 0) <> 0 Or (Nvl(n_结算金额, 0) = 0 And Nvl(冲预交_In, 0) = 0) Then
        If 结算方式_In Is Null Then
          v_Err_Msg := '未传入指定的结算方式,不能继续操作！';
          Raise Err_Item;
        End If;
      
        If Nvl(预交id_In, 0) <> 0 Then
          --传入的预交ID_In主要是为了解决三方交易,如果医保结算站用了该ID,需要用新的ID进行更新,三方交易用转入的ID
          Update 病人预交记录 Set ID = n_预交id Where ID = Nvl(预交id_In, 0);
          n_预交id := Nvl(预交id_In, 0);
        End If;
      
        Insert Into 病人预交记录
          (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 交易流水号, 交易说明, 结算序号, 合作单位, 卡类别id, 卡号,
           结算性质)
        Values
          (n_预交id, 4, 1, 单据号_In, r_Pati.病人id, 结算方式_In, Nvl(n_结算金额, 0), d_登记时间, v_操作员编号, v_操作员姓名, n_结帐id,
           合作单位_In || '缴款', n_组id, 交易流水号_In, 交易说明_In, n_结帐id, 合作单位_In, 卡类别id_In, 支付卡号_In, 4);
      End If;
    
      --更新人员缴款数据
    
      For v_缴款 In (Select 结算方式, Sum(Nvl(a.冲预交, 0)) As 冲预交
                   From 病人预交记录 A
                   Where a.结帐id = n_结帐id And Mod(a.记录性质, 10) <> 1 And Nvl(病人id, 0) = Nvl(病人id_In, 0)
                   Group By 结算方式) Loop
      
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) + Nvl(v_缴款.冲预交, 0)
        Where 收款员 = v_操作员姓名 And 性质 = 1 And 结算方式 = v_缴款.结算方式
        Returning 余额 Into n_返回值;
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额
            (收款员, 结算方式, 性质, 余额)
          Values
            (v_操作员姓名, v_缴款.结算方式, 1, Nvl(v_缴款.冲预交, 0));
          n_返回值 := Nvl(v_缴款.冲预交, 0);
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 人员缴款余额
          Where 收款员 = v_操作员姓名 And 结算方式 = 结算方式_In And 性质 = 1 And Nvl(余额, 0) = 0;
        End If;
      
      End Loop;
    
    End If;
  
    --处理挂号记录
    If 锁定类型_In = 2 Then
      Begin
        Select ID Into n_挂号id From 病人挂号记录 Where　记录状态 = 0 And NO = 单据号_In And 病人id = 病人id_In;
      Exception
        When Others Then
          Null;
      End;
    Else
      Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
    End If;
  
    Update 病人挂号记录
    Set 记录性质 = Decode(操作方式_In, 2, 2, 1), 记录状态 = Decode(锁定类型_In, 1, 0, 1), 门诊号 = r_Pati.门诊号, 操作员姓名 = v_操作员姓名,
        操作员编号 = v_操作员编号, 预约 = Decode(操作方式_In, 1, 0, 1),
        接收人 = Decode(锁定类型_In, 1, Null, Decode(操作方式_In, 2, Null, v_操作员姓名)),
        接收时间 = Case 锁定类型_In
                  When 1 Then
                   Null
                  Else
                   Case 操作方式_In
                     When 2 Then
                      Null
                     Else
                      d_登记时间
                   End
                End, 交易流水号 = Nvl(交易流水号_In, 交易流水号), 交易说明 = Nvl(交易说明_In, 交易说明), 合作单位 = Nvl(合作单位_In, 合作单位),
        预约操作员 = Decode(操作方式_In, 1, Nvl(预约操作员, Null), Nvl(预约操作员, v_操作员姓名)),
        预约操作员编号 = Decode(操作方式_In, 1, Nvl(预约操作员编号, Null), Nvl(预约操作员编号, v_操作员编号))
    Where ID = n_挂号id;
    If Sql%NotFound Then
      Begin
        Select 名称 Into v_付款方式 From 医疗付款方式 Where 编码 = r_Pati.付款方式 And Rownum < 2;
      Exception
        When Others Then
          v_付款方式 := Null;
      End;
      Insert Into 病人挂号记录
        (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 预约时间, 操作员编号,
         操作员姓名, 复诊, 号序, 预约, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 医疗付款方式, 预约操作员, 预约操作员编号)
      Values
        (n_挂号id, 单据号_In, Decode(操作方式_In, 2, 2, 1), Decode(锁定类型_In, 1, 0, 1), r_Pati.病人id, r_Pati.门诊号, r_Pati.姓名,
         r_Pati.性别, r_Pati.年龄, 号码_In, 0, v_诊室, Null, r_安排.科室id, r_安排.医生姓名, 0, Null, d_登记时间, 发生时间_In,
         Case When(Nvl(操作方式_In, 0)) = 1 Then Null Else 发生时间_In End, v_操作员编号, v_操作员姓名, 0, n_号序, Decode(操作方式_In, 1, 0, 1),
         Decode(操作方式_In, 2, Null, v_操作员姓名), Decode(操作方式_In, 2, To_Date(Null), d_登记时间), 交易流水号_In, 交易说明_In, 合作单位_In,
         v_付款方式, Decode(操作方式_In, 1, Null, v_操作员姓名), Decode(操作方式_In, 1, Null, v_操作员编号));
    End If;
    --锁定单据不能产生队列
    If 锁定类型_In <> 1 Then
      n_预约生成队列 := 0;
      If 操作方式_In > 1 Then
        n_预约生成队列 := Zl_To_Number(zl_GetSysParameter('预约生成队列', 1113));
      End If;
      --挂号和收费的预约都直接进入队列(收费预约缺少接收过程,所以直接和挂号一样直接进入队列)
      If 操作方式_In <> 2 Or n_预约生成队列 = 1 Then
        If Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113)) <> 0 Then
          --排队叫号模式:-0-不产生队列;1-按医生或分诊台排队;2-先分诊,后医生站
          If Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113)) = 0 Or n_预约生成队列 = 1 Then
            --产生队列
            --.按”执行部门” 的方式生成队列
            v_队列名称 := r_安排.科室id;
            v_排队号码 := Zlgetnextqueue(r_安排.科室id, n_挂号id, 号码_In || '|' || 号序_In);
            --挂号id_In,号码_In,号序_In,缺省日期_In,扩展_In(暂无用)
            v_排队序号 := Zlgetsequencenum(0, n_挂号id, 0);
            --挂号id_In,号码_In,号序_In,缺省日期_In,扩展_In(暂无用)
            d_排队时间 := Zl_Get_Queuedate(n_挂号id, 号码_In, 号序_In, d_Date);
            --  队列名称_In , 业务类型_In, 业务id_In,科室id_In,排队号码_In,v_排队标记,患者姓名_In,病人ID_IN, 诊室_In, 医生姓名_In, 排队时间_In
            Zl_排队叫号队列_Insert(v_队列名称, 0, n_挂号id, r_安排.科室id, v_排队号码, Null, r_Pati.姓名, r_Pati.病人id, v_诊室, r_安排.医生姓名,
                             d_排队时间, 预约方式_In, n_启用分时段, v_排队序号);
          End If;
        End If;
      End If;
    
      If Nvl(操作方式_In, 0) = 1 Then
        --处理票据使用情况
        If 票据号_In Is Not Null Then
          Select 票据打印内容_Id.Nextval Into n_打印id From Dual;
          --发出票据
          Insert Into 票据打印内容 (ID, 数据性质, NO) Values (n_打印id, 4, 单据号_In);
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
          Values
            (票据使用明细_Id.Nextval, Decode(收费票据_In, 1, 1, 4), 票据号_In, 1, 1, 领用id_In, n_打印id, d_登记时间, v_操作员姓名);
          --状态改动
          Update 票据领用记录
          Set 当前号码 = 票据号_In, 剩余数量 = Decode(Sign(剩余数量 - 1), -1, 0, 剩余数量 - 1), 使用时间 = Sysdate
          Where ID = Nvl(领用id_In, 0);
        End If;
        --病人本次就诊(以发生时间为准)
        If Nvl(r_Pati.病人id, 0) <> 0 Then
          Update 病人信息 Set 就诊时间 = 发生时间_In, 就诊状态 = 1, 就诊诊室 = v_诊室 Where 病人id = r_Pati.病人id;
        End If;
      End If;
    End If;
    --病人挂号汇总
    --解锁单据时不用再对汇总单据进行统计了 在锁定单据时已经进行了汇总
    If 锁定类型_In <> 2 Then
      --操作方式_IN:1-表示挂号,2-表示预约挂号不扣款,3-表示预约挂号,扣款
      --是否为预约接收:0-非预约挂号; 1-预约挂号,2-预约接收;3-收费预约
      n_预约 := Case
                When Nvl(操作方式_In, 0) = 1 Then
                 0
                When Nvl(操作方式_In, 0) = 2 Then
                 1
                When Nvl(操作方式_In, 0) = 3 Then
                 3
                Else
                 0
              End;
      Zl_病人挂号汇总_Update(r_安排.医生姓名, r_安排.医生id, r_安排.项目id, r_安排.科室id, 发生时间_In, n_预约, 号码_In);
    End If;
    --消息推送
    Begin
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 1, n_挂号id;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_001(n_挂号id, 单据号_In);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Err_Msg || '+[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_三方机构挂号_Insert;
/


Create Or Replace Procedure Zl_Third_Getdeptlist
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:获取可挂号科室

  --入参:Xml_In:
  --<IN>
  --  <CXTS>14</CXTS>        //查询天数
  --  <HZDW>支付宝</HZDW>    //合作单位
  --  <ZD></ZD>              //站点
  --</IN>

  --出参:Xml_Out
  --<OUTPUT>
  -- <KSLIST>
  --  <KS>
  --    <ID>科室ID</ID>       //科室ID
  --    <MC>科室名称</MC>     //科室名称
  --  </KS>
  --  <KS>
  --    ...
  --  </KS>
  -- </KSLIST>
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  v_Temp      Varchar(5000); --临时XML
  x_Templet   Xmltype; --模板XML
  v_Para      Varchar2(4000);
  n_查询天数  Number(5);
  n_预约天数  Number(5);
  n_Add_Lists Number(3);
  v_合作单位  合作单位安排控制.合作单位%Type;
  n_站点      部门表.站点%Type;
  v_Err_Msg   Varchar2(200);
  d_启用时间  Date;
  Err_Item Exception;
  n_挂号模式 Number(3);
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');
  Select Extractvalue(Value(A), 'IN/CXTS'), Extractvalue(Value(A), 'IN/HZDW'), Extractvalue(Value(A), 'IN/ZD')
  Into n_查询天数, v_合作单位, n_站点
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  v_Para     := zl_GetSysParameter('挂号排班模式');
  n_预约天数 := zl_GetSysParameter(66);
  n_挂号模式 := To_Number(Substr(v_Para, 1, 1));
  If n_挂号模式 = 1 Then
    Begin
      d_启用时间 := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
    Exception
      When Others Then
        d_启用时间 := Null;
    End;
  End If;
  If n_挂号模式 = 0 Then
    If n_查询天数 Is Null Then
      If v_合作单位 Is Null Then
        For r_Dept In (Select Distinct a.科室id, b.名称
                       From 挂号安排 A, 部门表 B
                       Where a.停用日期 Is Null And a.科室id = b.Id And (b.站点 = Nvl(n_站点, 0) Or b.站点 Is Null)) Loop
        
          If Nvl(n_Add_Lists, 0) = 0 Then
            --增加DJList节点
            Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype('<KSLIST></KSLIST>')) Into x_Templet From Dual;
            n_Add_Lists := 1;
          End If;
          v_Temp := '<KS>' || '<ID>' || r_Dept.科室id || '</ID>' || '<MC>' || r_Dept.名称 || '</MC>' || '</KS>';
          Select Appendchildxml(x_Templet, '/OUTPUT/KSLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
        End Loop;
      Else
        For r_Dept In (Select Distinct 科室id, 名称
                       From (Select b.科室id, d.名称
                              From (Select a.Id, Nvl(a.预约天数, n_预约天数) As 预约天数
                                     From 挂号安排 A
                                     Where a.停用日期 Is Null And Not Exists
                                      (Select 1 From 合作单位安排控制 Where 安排id = a.Id And 合作单位 = v_合作单位)
                                     Union All
                                     Select a.Id, Nvl(a.预约天数, n_预约天数) As 预约天数
                                     From 挂号安排 A, 合作单位安排控制 B
                                     Where a.停用日期 Is Null And 合作单位 = v_合作单位 And a.Id = b.安排id And b.数量 <> 0) A, 挂号安排 B,
                                   挂号安排计划 C, 部门表 D
                              Where a.Id = b.Id And c.安排id = a.Id And c.审核时间 Is Not Null And
                                    ((c.生效时间 < Sysdate And c.失效时间 > Sysdate + a.预约天数) Or
                                    (c.生效时间 Between Sysdate And Sysdate + a.预约天数) Or
                                    (c.失效时间 Between Sysdate And Sysdate + a.预约天数)) And Not Exists
                               (Select 1
                                     From 合作单位计划控制
                                     Where 计划id = c.Id And 合作单位 = v_合作单位 And 数量 = 0) And b.科室id = d.Id And
                                    (d.站点 = Nvl(n_站点, 0) Or d.站点 Is Null) And
                                    (Sysdate Between d.建档时间 And d.撤档时间 Or Sysdate >= d.建档时间 And d.撤档时间 Is Null)
                              Union All
                              Select b.科室id, d.名称
                              From (Select a.Id
                                     From 挂号安排 A
                                     Where a.停用日期 Is Null And Not Exists
                                      (Select 1 From 合作单位安排控制 Where 安排id = a.Id And 合作单位 = v_合作单位)
                                     Union All
                                     Select a.Id
                                     From 挂号安排 A, 合作单位安排控制 B
                                     Where a.停用日期 Is Null And 合作单位 = v_合作单位 And a.Id = b.安排id And b.数量 <> 0) A, 挂号安排 B,
                                   部门表 D
                              Where a.Id = b.Id And Not Exists
                               (Select 1 From 挂号安排计划 Where 安排id = a.Id And Rownum < 2) And b.科室id = d.Id And
                                    (d.站点 = Nvl(n_站点, 0) Or d.站点 Is Null) And
                                    (Sysdate Between d.建档时间 And d.撤档时间 Or Sysdate >= d.建档时间 And d.撤档时间 Is Null))) Loop
          If Nvl(n_Add_Lists, 0) = 0 Then
            --增加DJList节点
            Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype('<KSLIST></KSLIST>')) Into x_Templet From Dual;
            n_Add_Lists := 1;
          End If;
          v_Temp := '<KS>' || '<ID>' || r_Dept.科室id || '</ID>' || '<MC>' || r_Dept.名称 || '</MC>' || '</KS>';
          Select Appendchildxml(x_Templet, '/OUTPUT/KSLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
        End Loop;
      End If;
    Else
      If v_合作单位 Is Null Then
        For r_Dept In (Select Distinct a.科室id, b.名称
                       From 挂号安排 A, 部门表 B
                       Where a.停用日期 Is Null And a.科室id = b.Id And (b.站点 = Nvl(n_站点, 0) Or b.站点 Is Null)) Loop
        
          If Nvl(n_Add_Lists, 0) = 0 Then
            --增加DJList节点
            Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype('<KSLIST></KSLIST>')) Into x_Templet From Dual;
            n_Add_Lists := 1;
          End If;
          v_Temp := '<KS>' || '<ID>' || r_Dept.科室id || '</ID>' || '<MC>' || r_Dept.名称 || '</MC>' || '</KS>';
          Select Appendchildxml(x_Templet, '/OUTPUT/KSLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
        End Loop;
      Else
        For r_Dept In (Select Distinct 科室id, 名称
                       From (Select b.科室id, d.名称
                              From (Select a.Id
                                     From 挂号安排 A
                                     Where a.停用日期 Is Null And Not Exists
                                      (Select 1 From 合作单位安排控制 Where 安排id = a.Id And 合作单位 = v_合作单位)
                                     Union All
                                     Select a.Id
                                     From 挂号安排 A, 合作单位安排控制 B
                                     Where a.停用日期 Is Null And 合作单位 = v_合作单位 And a.Id = b.安排id And b.数量 <> 0) A, 挂号安排 B,
                                   挂号安排计划 C, 部门表 D
                              Where a.Id = b.Id And c.安排id = a.Id And c.审核时间 Is Not Null And
                                    ((c.生效时间 < Sysdate And c.失效时间 > Sysdate + n_查询天数) Or
                                    (c.生效时间 Between Sysdate And Sysdate + n_查询天数) Or
                                    (c.失效时间 Between Sysdate And Sysdate + n_查询天数)) And Not Exists
                               (Select 1
                                     From 合作单位计划控制
                                     Where 计划id = c.Id And 合作单位 = v_合作单位 And 数量 = 0) And b.科室id = d.Id And
                                    (d.站点 = Nvl(n_站点, 0) Or d.站点 Is Null) And
                                    (Sysdate Between d.建档时间 And d.撤档时间 Or Sysdate >= d.建档时间 And d.撤档时间 Is Null)
                              Union All
                              Select b.科室id, d.名称
                              From (Select a.Id
                                     From 挂号安排 A
                                     Where a.停用日期 Is Null And Not Exists
                                      (Select 1 From 合作单位安排控制 Where 安排id = a.Id And 合作单位 = v_合作单位)
                                     Union All
                                     Select a.Id
                                     From 挂号安排 A, 合作单位安排控制 B
                                     Where a.停用日期 Is Null And 合作单位 = v_合作单位 And a.Id = b.安排id And b.数量 <> 0) A, 挂号安排 B,
                                   部门表 D
                              Where a.Id = b.Id And Not Exists
                               (Select 1 From 挂号安排计划 Where 安排id = a.Id And Rownum < 2) And b.科室id = d.Id And
                                    (d.站点 = Nvl(n_站点, 0) Or d.站点 Is Null) And
                                    (Sysdate Between d.建档时间 And d.撤档时间 Or Sysdate >= d.建档时间 And d.撤档时间 Is Null))) Loop
          If Nvl(n_Add_Lists, 0) = 0 Then
            --增加DJList节点
            Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype('<KSLIST></KSLIST>')) Into x_Templet From Dual;
            n_Add_Lists := 1;
          End If;
          v_Temp := '<KS>' || '<ID>' || r_Dept.科室id || '</ID>' || '<MC>' || r_Dept.名称 || '</MC>' || '</KS>';
          Select Appendchildxml(x_Templet, '/OUTPUT/KSLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
        End Loop;
      End If;
    End If;
  Else
    --出诊表排班模式
    If n_查询天数 Is Null Then
      If v_合作单位 Is Null Then
        For r_Dept In (Select Distinct a.科室id, b.名称
                       From 挂号安排 A, 部门表 B
                       Where a.停用日期 Is Null And a.科室id = b.Id And (b.站点 = Nvl(n_站点, 0) Or b.站点 Is Null) And
                             Sysdate < d_启用时间
                       Union
                       Select Distinct a.科室id, b.名称
                       From 临床出诊记录 A, 部门表 B, 临床出诊号源 C
                       Where a.号源id = c.Id And a.开始时间 > d_启用时间 And a.出诊日期 >= Trunc(Sysdate) And
                             a.出诊日期 <= Trunc(Sysdate + Nvl(c.预约天数, n_预约天数)) And Nvl(a.是否发布, 0) = 1 And a.科室id = b.Id And
                             (b.站点 = Nvl(n_站点, 0) Or b.站点 Is Null)) Loop
          If Nvl(n_Add_Lists, 0) = 0 Then
            Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype('<KSLIST></KSLIST>')) Into x_Templet From Dual;
            n_Add_Lists := 1;
          End If;
          v_Temp := '<KS>' || '<ID>' || r_Dept.科室id || '</ID>' || '<MC>' || r_Dept.名称 || '</MC>' || '</KS>';
          Select Appendchildxml(x_Templet, '/OUTPUT/KSLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
        End Loop;
      Else
        For r_Dept In (Select Distinct 科室id, 名称
                       From (Select b.科室id, d.名称
                              From (Select a.Id, Nvl(a.预约天数, n_预约天数) As 预约天数
                                     From 挂号安排 A
                                     Where a.停用日期 Is Null And Not Exists
                                      (Select 1 From 合作单位安排控制 Where 安排id = a.Id And 合作单位 = v_合作单位)
                                     Union All
                                     Select a.Id, Nvl(a.预约天数, n_预约天数) As 预约天数
                                     From 挂号安排 A, 合作单位安排控制 B
                                     Where a.停用日期 Is Null And 合作单位 = v_合作单位 And a.Id = b.安排id And b.数量 <> 0) A, 挂号安排 B,
                                   挂号安排计划 C, 部门表 D
                              Where a.Id = b.Id And Sysdate < d_启用时间 And c.安排id = a.Id And c.审核时间 Is Not Null And
                                    ((c.生效时间 < Sysdate And c.失效时间 > Sysdate + a.预约天数) Or
                                    (c.生效时间 Between Sysdate And Sysdate + a.预约天数) Or
                                    (c.失效时间 Between Sysdate And Sysdate + a.预约天数)) And Not Exists
                               (Select 1
                                     From 合作单位计划控制
                                     Where 计划id = c.Id And 合作单位 = v_合作单位 And 数量 = 0) And b.科室id = d.Id And
                                    (d.站点 = Nvl(n_站点, 0) Or d.站点 Is Null) And
                                    (Sysdate Between d.建档时间 And d.撤档时间 Or Sysdate >= d.建档时间 And d.撤档时间 Is Null)
                              Union All
                              Select b.科室id, d.名称
                              From (Select a.Id
                                     From 挂号安排 A
                                     Where a.停用日期 Is Null And Not Exists
                                      (Select 1 From 合作单位安排控制 Where 安排id = a.Id And 合作单位 = v_合作单位)
                                     Union All
                                     Select a.Id
                                     From 挂号安排 A, 合作单位安排控制 B
                                     Where a.停用日期 Is Null And 合作单位 = v_合作单位 And a.Id = b.安排id And b.数量 <> 0) A, 挂号安排 B,
                                   部门表 D
                              Where a.Id = b.Id And Sysdate < d_启用时间 And Not Exists
                               (Select 1 From 挂号安排计划 Where 安排id = a.Id And Rownum < 2) And b.科室id = d.Id And
                                    (d.站点 = Nvl(n_站点, 0) Or d.站点 Is Null) And
                                    (Sysdate Between d.建档时间 And d.撤档时间 Or Sysdate >= d.建档时间 And d.撤档时间 Is Null))
                       Union
                       Select Distinct 科室id, 名称
                       From (Select a.科室id, b.名称
                              From 临床出诊记录 A, 部门表 B, 临床出诊号源 C
                              Where a.号源id = c.Id And a.开始时间 > d_启用时间 And a.出诊日期 >= Trunc(Sysdate) And
                                    a.出诊日期 <= Trunc(Sysdate + Nvl(c.预约天数, n_预约天数)) And Nvl(a.是否发布, 0) = 1 And
                                    a.科室id = b.Id And (b.站点 = Nvl(n_站点, 0) Or b.站点 Is Null) And Not Exists
                               (Select 1 From 临床出诊挂号控制记录 Where 记录id = a.Id And 性质 = 1 And 类型 = 1)
                              Union
                              Select a.科室id, b.名称
                              From 临床出诊记录 A, 部门表 B, 临床出诊号源 C
                              Where a.号源id = c.Id And a.开始时间 > d_启用时间 And a.出诊日期 >= Trunc(Sysdate) And
                                    a.出诊日期 <= Trunc(Sysdate + Nvl(c.预约天数, n_预约天数)) And Nvl(a.是否发布, 0) = 1 And
                                    a.科室id = b.Id And (b.站点 = Nvl(n_站点, 0) Or b.站点 Is Null) And Exists
                               (Select 1
                                     From 临床出诊挂号控制记录
                                     Where 记录id = a.Id And 名称 = v_合作单位 And 性质 = 1 And 类型 = 1 And 控制方式 <> 0))) Loop
          If Nvl(n_Add_Lists, 0) = 0 Then
            Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype('<KSLIST></KSLIST>')) Into x_Templet From Dual;
            n_Add_Lists := 1;
          End If;
          v_Temp := '<KS>' || '<ID>' || r_Dept.科室id || '</ID>' || '<MC>' || r_Dept.名称 || '</MC>' || '</KS>';
          Select Appendchildxml(x_Templet, '/OUTPUT/KSLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
        End Loop;
      End If;
    Else
      If v_合作单位 Is Null Then
        For r_Dept In (Select Distinct a.科室id, b.名称
                       From 挂号安排 A, 部门表 B
                       Where a.停用日期 Is Null And a.科室id = b.Id And (b.站点 = Nvl(n_站点, 0) Or b.站点 Is Null) And
                             Sysdate < d_启用时间
                       Union
                       Select Distinct a.科室id, b.名称
                       From 临床出诊记录 A, 部门表 B
                       Where a.出诊日期 >= Trunc(Sysdate) And a.开始时间 > d_启用时间 And a.出诊日期 <= Trunc(Sysdate + n_查询天数) And
                             Nvl(a.是否发布, 0) = 1 And a.科室id = b.Id And (b.站点 = Nvl(n_站点, 0) Or b.站点 Is Null)) Loop
          If Nvl(n_Add_Lists, 0) = 0 Then
            Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype('<KSLIST></KSLIST>')) Into x_Templet From Dual;
            n_Add_Lists := 1;
          End If;
          v_Temp := '<KS>' || '<ID>' || r_Dept.科室id || '</ID>' || '<MC>' || r_Dept.名称 || '</MC>' || '</KS>';
          Select Appendchildxml(x_Templet, '/OUTPUT/KSLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
        End Loop;
      Else
        For r_Dept In (Select Distinct 科室id, 名称
                       From (Select b.科室id, d.名称
                              From (Select a.Id
                                     From 挂号安排 A
                                     Where a.停用日期 Is Null And Not Exists
                                      (Select 1 From 合作单位安排控制 Where 安排id = a.Id And 合作单位 = v_合作单位)
                                     Union All
                                     Select a.Id
                                     From 挂号安排 A, 合作单位安排控制 B
                                     Where a.停用日期 Is Null And 合作单位 = v_合作单位 And a.Id = b.安排id And b.数量 <> 0) A, 挂号安排 B,
                                   挂号安排计划 C, 部门表 D
                              Where a.Id = b.Id And Sysdate < d_启用时间 And c.安排id = a.Id And c.审核时间 Is Not Null And
                                    ((c.生效时间 < Sysdate And c.失效时间 > Sysdate + n_查询天数) Or
                                    (c.生效时间 Between Sysdate And Sysdate + n_查询天数) Or
                                    (c.失效时间 Between Sysdate And Sysdate + n_查询天数)) And Not Exists
                               (Select 1
                                     From 合作单位计划控制
                                     Where 计划id = c.Id And 合作单位 = v_合作单位 And 数量 = 0) And b.科室id = d.Id And
                                    (d.站点 = Nvl(n_站点, 0) Or d.站点 Is Null) And
                                    (Sysdate Between d.建档时间 And d.撤档时间 Or Sysdate >= d.建档时间 And d.撤档时间 Is Null)
                              Union All
                              Select b.科室id, d.名称
                              From (Select a.Id
                                     From 挂号安排 A
                                     Where a.停用日期 Is Null And Not Exists
                                      (Select 1 From 合作单位安排控制 Where 安排id = a.Id And 合作单位 = v_合作单位)
                                     Union All
                                     Select a.Id
                                     From 挂号安排 A, 合作单位安排控制 B
                                     Where a.停用日期 Is Null And 合作单位 = v_合作单位 And a.Id = b.安排id And b.数量 <> 0) A, 挂号安排 B,
                                   部门表 D
                              Where a.Id = b.Id And Sysdate < d_启用时间 And Not Exists
                               (Select 1 From 挂号安排计划 Where 安排id = a.Id And Rownum < 2) And b.科室id = d.Id And
                                    (d.站点 = Nvl(n_站点, 0) Or d.站点 Is Null) And
                                    (Sysdate Between d.建档时间 And d.撤档时间 Or Sysdate >= d.建档时间 And d.撤档时间 Is Null))
                       Union
                       Select Distinct 科室id, 名称
                       From (Select a.科室id, b.名称
                              From 临床出诊记录 A, 部门表 B
                              Where a.出诊日期 >= Trunc(Sysdate) And a.开始时间 > d_启用时间 And a.出诊日期 <= Trunc(Sysdate + n_查询天数) And
                                    Nvl(a.是否发布, 0) = 1 And a.科室id = b.Id And (b.站点 = Nvl(n_站点, 0) Or b.站点 Is Null) And
                                    Not Exists
                               (Select 1 From 临床出诊挂号控制记录 Where 记录id = a.Id And 性质 = 1 And 类型 = 1)
                              Union
                              Select a.科室id, b.名称
                              From 临床出诊记录 A, 部门表 B
                              Where a.出诊日期 >= Trunc(Sysdate) And a.开始时间 > d_启用时间 And a.出诊日期 <= Trunc(Sysdate + n_查询天数) And
                                    Nvl(a.是否发布, 0) = 1 And a.科室id = b.Id And (b.站点 = Nvl(n_站点, 0) Or b.站点 Is Null) And
                                    Exists
                               (Select 1
                                     From 临床出诊挂号控制记录
                                     Where 记录id = a.Id And 名称 = v_合作单位 And 性质 = 1 And 类型 = 1 And 控制方式 <> 0))) Loop
          If Nvl(n_Add_Lists, 0) = 0 Then
            Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype('<KSLIST></KSLIST>')) Into x_Templet From Dual;
            n_Add_Lists := 1;
          End If;
          v_Temp := '<KS>' || '<ID>' || r_Dept.科室id || '</ID>' || '<MC>' || r_Dept.名称 || '</MC>' || '</KS>';
          Select Appendchildxml(x_Templet, '/OUTPUT/KSLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
        End Loop;
      End If;
    End If;
  End If;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getdeptlist;
/


Create Or Replace Procedure Zl_Third_Docarrange
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:医生排班计划
  --入参:Xml_In:
  --<IN>
  --   <YSID>870</YSID>    //医生ID
  --   <KDID>870</KSID>    //科室ID
  --   <KSSJ>2014-10-29 </KSSJ>    //开始时间
  --   <CXTS>14</CXTS>    //查询天数
  --   <HZDW>支付宝</HZDW> //合作单位
  --   <HL>号类</HL>      //号类，可传多个，用逗号分隔，格式:普通,专家,...
  --</IN>

  --出参:Xml_Out
  --<OUTPUT>
  --   <PBLIST>       //未返回该节点表示没有数据
  --    <PB>
  --     <RQ>2014-10-29</RQ>     //日期
  --     <SYHS>5</SYHS>    //剩余号数
  --     <SBSJ>全日</SBSJ>             //上班时间
  --     <YGS>5</YGS>    //已挂号数
  --    </PB>
  --   <PBLIST>
  --   <ERROR><MSG></MSG></ERROR> //错误情况返回
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  d_日期         Date;
  v_排班         挂号安排.周日%Type;
  n_限号数       挂号安排限制.限号数%Type;
  n_已挂数       挂号安排限制.限号数%Type;
  n_总已挂数     挂号安排限制.限号数%Type;
  n_限约数       挂号安排限制.限号数%Type;
  n_已约数       挂号安排限制.限号数%Type;
  n_剩余数       挂号安排限制.限号数%Type;
  v_上班时间     Varchar2(300);
  n_医生id       人员表.Id%Type;
  n_科室id       部门表.Id%Type;
  n_查询天数     Number(4);
  n_合作单位数量 Number(5);
  n_合约已挂数   Number(4);
  n_合约存在     Number(3);
  n_安排存在     Number(3);
  v_号码         挂号安排.号码%Type;
  n_安排id       挂号安排计划.安排id%Type;
  n_计划id       挂号安排计划.Id%Type;
  v_合作单位     挂号合作单位.名称%Type;
  n_Daycount     Number(4);
  d_开始时间     Date;
  d_原始时间     Date;
  n_禁用         Number(3);
  v_Temp         Varchar2(32767); --临时XML
  x_Templet      Xmltype; --模板XML
  v_Err_Msg      Varchar2(200);
  v_号类         Varchar2(200);
  n_Exists       Number(2);
  n_挂号模式     Number(3);
  n_合约模式     临床出诊挂号控制记录.控制方式%Type;
  v_启用时间     Varchar2(500);
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/YSID'), Extractvalue(Value(A), 'IN/KSID'), Extractvalue(Value(A), 'IN/CXTS'),
         To_Date(Extractvalue(Value(A), 'IN/KSSJ'), 'YYYY-MM-DD hh24:mi:ss'), Extractvalue(Value(A), 'IN/HZDW'),
         Extractvalue(Value(A), 'IN/HL')
  Into n_医生id, n_科室id, n_查询天数, d_开始时间, v_合作单位, v_号类
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  n_查询天数 := Nvl(n_查询天数, 14);
  d_原始时间 := Trunc(d_开始时间);
  d_开始时间 := Trunc(d_开始时间);
  n_Daycount := 0;
  n_挂号模式 := To_Number(Substr(Nvl(zl_GetSysParameter('挂号排班模式'), '0'), 1, 1));
  v_启用时间 := Substr(Nvl(zl_GetSysParameter('挂号排班模式'), '0'), 3);
  If n_挂号模式 = 0 Then
    If Nvl(n_科室id, 0) = 0 Then
      While (n_Daycount < n_查询天数) Loop
        If Trunc(d_开始时间 + n_Daycount) = Trunc(Sysdate) Then
          d_开始时间 := Sysdate - n_Daycount;
        Else
          d_开始时间 := d_原始时间;
        End If;
        n_安排存在 := 0;
        v_上班时间 := Null;
        n_总已挂数 := 0;
        n_已挂数   := 0;
        n_剩余数   := 0;
        n_限号数   := 0;
        n_已约数   := 0;
        n_限约数   := 0;
        For r_排班 In (Select d_开始时间 + n_Daycount As 日期, a.排班, a.限号数, a.限约数, Nvl(Hz.已挂数, 0) As 已挂数, Nvl(Hz.已约数, 0) As 已约数,
                            a.安排id, a.计划id, a.号码, a.号类
                     
                     From (Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Ap.医生姓名, Nvl(Ap.医生id, 0) As 医生id, Ry.专业技术职务 As 职称, Ap.号码,
                                   Ap.安排id, Ap.计划id, Ap.排班, Ap.项目id, Fy.名称 As 项目名称, Ap.序号控制, Ap.限号数, Ap.限约数
                            
                            From (Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Ap.医生姓名, Ap.医生id, Ap.号码, Ap.Id As 安排id, 0 As 计划id,
                                          Ap.项目id, Nvl(Ap.序号控制, 0) As 序号控制,
                                          Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', Ap.周日, '2', Ap.周一, '3', Ap.周二, '4',
                                                  Ap.周三, '5', Ap.周四, '6', Ap.周五, '7', Ap.周六, Null) As 排班,
                                          Nvl(Xz.限约数, Xz.限号数) As 限约数, Nvl(Xz.限号数, 0) As 限号数
                                   From 挂号安排 Ap, 部门表 Bm, 挂号安排限制 Xz
                                   Where Ap.科室id = Bm.Id(+) And Ap.医生id = n_医生id And Ap.停用日期 Is Null And
                                         d_开始时间 + n_Daycount Between Nvl(Ap.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                         Nvl(Ap.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.安排id(+) = Ap.Id And
                                         Xz.限制项目(+) =
                                         Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                                '5', '周四', '6', '周五', '7', '周六', Null) And Not Exists
                                    (Select Rownum
                                          From 挂号安排停用状态 Ty
                                          Where Ty.安排id = Ap.Id And d_开始时间 + n_Daycount Between Ty.开始停止时间 And Ty.结束停止时间) And
                                         Not Exists
                                    (Select Rownum
                                          From 挂号安排计划 Jh
                                          Where Jh.安排id = Ap.Id And Jh.审核时间 Is Not Null And
                                                d_开始时间 + n_Daycount Between Nvl(Jh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                Nvl(Jh.失效时间, To_Date('3000-01-01', 'YYYY-MM-DD')))
                                   Union All
                                   Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Jh.医生姓名, Jh.医生id, Ap.号码, Ap.Id As 安排id,
                                          Jh.Id As 计划id, Jh.项目id, Nvl(Jh.序号控制, 0) As 序号控制,
                                          Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', Jh.周日, '2', Jh.周一, '3', Jh.周二, '4',
                                                  Jh.周三, '5', Jh.周四, '6', Jh.周五, '7', Jh.周六, Null) As 排班,
                                          Nvl(Xz.限约数, Xz.限号数) As 限约数, Nvl(Xz.限号数, 0) As 限号数
                                   From 挂号安排计划 Jh, 挂号安排 Ap, 部门表 Bm, 挂号计划限制 Xz
                                   Where Jh.安排id = Ap.Id And Ap.科室id = Bm.Id(+) And Ap.停用日期 Is Null And Jh.医生id = n_医生id And
                                         d_开始时间 + n_Daycount Between Nvl(Jh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                         Nvl(Jh.失效时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.计划id(+) = Jh.Id And
                                         Xz.限制项目(+) =
                                         Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                                '5', '周四', '6', '周五', '7', '周六', Null) And Not Exists
                                    (Select Rownum
                                          From 挂号安排停用状态 Ty
                                          Where Ty.安排id = Ap.Id And d_开始时间 + n_Daycount Between Ty.开始停止时间 And Ty.结束停止时间) And
                                         (Jh.生效时间, Jh.安排id) =
                                         (Select Max(Sxjh.生效时间) As 生效时间, 安排id
                                          From 挂号安排计划 Sxjh
                                          Where Sxjh.审核时间 Is Not Null And d_开始时间 + n_Daycount Between
                                                Nvl(Sxjh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                Nvl(Sxjh.失效时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Sxjh.安排id = Jh.安排id
                                          Group By Sxjh.安排id)) Ap, 部门表 Bm, 人员表 Ry, 收费项目目录 Fy
                            Where Ap.科室id = Bm.Id(+) And Ap.医生id = Ry.Id(+) And Ap.项目id = Fy.Id And 排班 Is Not Null) A,
                          病人挂号汇总 Hz, 收费价目 B
                     Where a.号码 = Hz.号码(+) And Hz.日期(+) = Trunc(d_开始时间 + n_Daycount) And a.项目id = b.收费细目id And
                           b.终止日期 > d_开始时间 + n_Daycount And b.执行日期 <= d_开始时间 + n_Daycount
                     
                     ) Loop
          If v_号类 Is Null Or Instr(',' || v_号类 || ',', ',' || r_排班.号类 || ',') > 0 Then
            v_上班时间 := v_上班时间 || '+' || r_排班.排班;
            n_总已挂数 := n_总已挂数 + r_排班.已挂数;
            n_已挂数   := r_排班.已挂数;
            n_限号数   := r_排班.限号数;
            n_已约数   := r_排班.已约数;
            n_限约数   := r_排班.限约数;
            n_安排id   := Nvl(r_排班.安排id, 0);
            n_计划id   := Nvl(r_排班.计划id, 0);
            v_号码     := r_排班.号码;
            n_安排存在 := 1;
            If v_上班时间 Is Not Null Then
              If v_合作单位 Is Not Null Then
                If n_计划id <> 0 Then
                  Begin
                    Select 1
                    Into n_合约存在
                    From 合作单位计划控制
                    Where 计划id = n_计划id And 合作单位 = v_合作单位 And Rownum < 2;
                  Exception
                    When Others Then
                      n_合约存在 := 0;
                  End;
                Else
                  Begin
                    Select 1
                    Into n_合约存在
                    From 合作单位安排控制
                    Where 安排id = n_安排id And 合作单位 = v_合作单位 And Rownum < 2;
                  Exception
                    When Others Then
                      n_合约存在 := 0;
                  End;
                End If;
              End If;
            
              If v_合作单位 Is Null Or Nvl(n_合约存在, 0) = 0 Then
                If n_计划id <> 0 Then
                  Begin
                    Select Sum(数量)
                    Into n_合作单位数量
                    From 合作单位计划控制
                    Where 计划id = n_计划id And
                          限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                        '5', '周四', '6', '周五', '7', '周六', Null);
                  Exception
                    When Others Then
                      n_合作单位数量 := 0;
                  End;
                Else
                  Begin
                    Select Sum(数量)
                    Into n_合作单位数量
                    From 合作单位安排控制
                    Where 安排id = n_安排id And
                          限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                        '5', '周四', '6', '周五', '7', '周六', Null);
                  Exception
                    When Others Then
                      n_合作单位数量 := 0;
                  End;
                End If;
                Begin
                  Select Count(1)
                  Into n_合约已挂数
                  From 病人挂号记录
                  Where 号别 = v_号码 And 记录状态 = 1 And 合作单位 Is Not Null And 发生时间 Between Trunc(d_开始时间 + n_Daycount) And
                        Trunc(d_开始时间 + n_Daycount + 1) - 1 / 60 / 60 / 24;
                Exception
                  When Others Then
                    n_合约已挂数 := 0;
                End;
                If n_合作单位数量 = 0 Then
                  n_合作单位数量 := Null;
                End If;
                If Trunc(d_开始时间 + n_Daycount) > Trunc(Sysdate) Then
                  n_剩余数 := Nvl(n_剩余数, 0) + n_限约数 - n_已约数 - Nvl(n_合作单位数量, n_合约已挂数) + Nvl(n_合约已挂数, 0);
                Else
                  n_剩余数 := Nvl(n_剩余数, 0) + n_限号数 - n_已挂数 - Nvl(n_合作单位数量, n_合约已挂数) + Nvl(n_合约已挂数, 0);
                End If;
              Else
                --合约单位
                If n_计划id <> 0 Then
                  Begin
                    Select 1
                    Into n_禁用
                    From 合作单位计划控制
                    Where 计划id = n_计划id And
                          限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                        '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位 And 数量 = 0 And
                          Rownum < 2;
                  Exception
                    When Others Then
                      n_禁用 := 0;
                  End;
                Else
                  Begin
                    Select 1
                    Into n_禁用
                    From 合作单位安排控制
                    Where 安排id = n_安排id And
                          限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                        '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位 And 数量 = 0 And
                          Rownum < 2;
                  Exception
                    When Others Then
                      n_禁用 := 0;
                  End;
                End If;
                If Nvl(n_禁用, 0) = 0 Then
                  Begin
                    Select Count(1)
                    Into n_合约已挂数
                    From 病人挂号记录
                    Where 号别 = v_号码 And 记录状态 = 1 And 合作单位 = v_合作单位 And 发生时间 Between Trunc(d_开始时间 + n_Daycount) And
                          Trunc(d_开始时间 + n_Daycount + 1) - 1 / 60 / 60 / 24;
                  Exception
                    When Others Then
                      n_合约已挂数 := 0;
                  End;
                  If n_计划id <> 0 Then
                    Begin
                      Select Sum(数量)
                      Into n_合作单位数量
                      From 合作单位计划控制
                      Where 计划id = n_计划id And
                            限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                          '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位;
                    Exception
                      When Others Then
                        n_合作单位数量 := 0;
                    End;
                  Else
                    Begin
                      Select Sum(数量)
                      Into n_合作单位数量
                      From 合作单位安排控制
                      Where 安排id = n_安排id And
                            限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                          '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位;
                    Exception
                      When Others Then
                        n_合作单位数量 := 0;
                    End;
                  End If;
                  If n_合作单位数量 = 0 Then
                    n_合作单位数量 := Null;
                  End If;
                  n_剩余数 := Nvl(n_剩余数, 0) + Nvl(n_合作单位数量, n_合约已挂数) - Nvl(n_合约已挂数, 0);
                
                End If;
              End If;
            End If;
            n_合作单位数量 := 0;
            n_合约存在     := 0;
            n_禁用         := 0;
          End If;
        End Loop;
        v_上班时间 := Substr(v_上班时间, 2);
        If n_安排存在 = 1 Then
          v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_开始时间 + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                    '<SYHS>' || n_剩余数 || '</SYHS>' || '<SBSJ>' || v_上班时间 || '</SBSJ>' || '<YGS>' || n_总已挂数 || '</YGS>' ||
                    '</PB>';
        End If;
        n_Daycount := n_Daycount + 1;
      End Loop;
    Else
      While (n_Daycount < n_查询天数) Loop
        If Trunc(d_开始时间 + n_Daycount) = Trunc(Sysdate) Then
          d_开始时间 := Sysdate - n_Daycount;
        Else
          d_开始时间 := d_原始时间;
        End If;
        v_上班时间 := Null;
        n_总已挂数 := 0;
        n_已挂数   := 0;
        n_剩余数   := 0;
        n_限号数   := 0;
        n_已约数   := 0;
        n_限约数   := 0;
        n_安排存在 := 0;
        For r_排班 In (Select d_开始时间 + n_Daycount As 日期, a.排班, a.限号数, a.限约数, Nvl(Hz.已挂数, 0) As 已挂数, Nvl(Hz.已约数, 0) As 已约数,
                            a.安排id, a.计划id, a.号码, a.号类
                     
                     From (Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Ap.医生姓名, Nvl(Ap.医生id, 0) As 医生id, Ry.专业技术职务 As 职称, Ap.号码,
                                   Ap.安排id, Ap.计划id, Ap.排班, Ap.项目id, Fy.名称 As 项目名称, Ap.序号控制, Ap.限号数, Ap.限约数
                            
                            From (Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Ap.医生姓名, Ap.医生id, Ap.号码, Ap.Id As 安排id, 0 As 计划id,
                                          Ap.项目id, Nvl(Ap.序号控制, 0) As 序号控制,
                                          Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', Ap.周日, '2', Ap.周一, '3', Ap.周二, '4',
                                                  Ap.周三, '5', Ap.周四, '6', Ap.周五, '7', Ap.周六, Null) As 排班,
                                          Nvl(Xz.限约数, Xz.限号数) As 限约数, Nvl(Xz.限号数, 0) As 限号数
                                   From 挂号安排 Ap, 部门表 Bm, 挂号安排限制 Xz
                                   Where Ap.科室id = Bm.Id(+) And Ap.医生id = n_医生id And Ap.科室id = n_科室id And Ap.停用日期 Is Null And
                                         d_开始时间 + n_Daycount Between Nvl(Ap.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                         Nvl(Ap.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.安排id(+) = Ap.Id And
                                         Xz.限制项目(+) =
                                         Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                                '5', '周四', '6', '周五', '7', '周六', Null) And Not Exists
                                    (Select Rownum
                                          From 挂号安排停用状态 Ty
                                          Where Ty.安排id = Ap.Id And d_开始时间 + n_Daycount Between Ty.开始停止时间 And Ty.结束停止时间) And
                                         Not Exists
                                    (Select Rownum
                                          From 挂号安排计划 Jh
                                          Where Jh.安排id = Ap.Id And Jh.审核时间 Is Not Null And
                                                d_开始时间 + n_Daycount Between Nvl(Jh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                Nvl(Jh.失效时间, To_Date('3000-01-01', 'YYYY-MM-DD')))
                                   Union All
                                   Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Jh.医生姓名, Jh.医生id, Ap.号码, Ap.Id As 安排id,
                                          Jh.Id As 计划id, Jh.项目id, Nvl(Jh.序号控制, 0) As 序号控制,
                                          Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', Jh.周日, '2', Jh.周一, '3', Jh.周二, '4',
                                                  Jh.周三, '5', Jh.周四, '6', Jh.周五, '7', Jh.周六, Null) As 排班,
                                          Nvl(Xz.限约数, Xz.限号数) As 限约数, Nvl(Xz.限号数, 0) As 限号数
                                   From 挂号安排计划 Jh, 挂号安排 Ap, 部门表 Bm, 挂号计划限制 Xz
                                   Where Jh.安排id = Ap.Id And Ap.科室id = Bm.Id(+) And Ap.停用日期 Is Null And Jh.医生id = n_医生id And
                                         Ap.科室id = n_科室id And
                                         d_开始时间 + n_Daycount Between Nvl(Jh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                         Nvl(Jh.失效时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.计划id(+) = Jh.Id And
                                         Xz.限制项目(+) =
                                         Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                                '5', '周四', '6', '周五', '7', '周六', Null) And Not Exists
                                    (Select Rownum
                                          From 挂号安排停用状态 Ty
                                          Where Ty.安排id = Ap.Id And d_开始时间 + n_Daycount Between Ty.开始停止时间 And Ty.结束停止时间) And
                                         (Jh.生效时间, Jh.安排id) =
                                         (Select Max(Sxjh.生效时间) As 生效时间, 安排id
                                          From 挂号安排计划 Sxjh
                                          Where Sxjh.审核时间 Is Not Null And d_开始时间 + n_Daycount Between
                                                Nvl(Sxjh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                Nvl(Sxjh.失效时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Sxjh.安排id = Jh.安排id
                                          Group By Sxjh.安排id)) Ap, 部门表 Bm, 人员表 Ry, 收费项目目录 Fy
                            Where Ap.科室id = Bm.Id(+) And Ap.医生id = Ry.Id(+) And Ap.项目id = Fy.Id And 排班 Is Not Null) A,
                          病人挂号汇总 Hz, 收费价目 B
                     Where a.号码 = Hz.号码(+) And Hz.日期(+) = Trunc(d_开始时间 + n_Daycount) And a.项目id = b.收费细目id And
                           b.终止日期 > d_开始时间 + n_Daycount And b.执行日期 <= d_开始时间 + n_Daycount
                     
                     ) Loop
          If v_号类 Is Null Or Instr(',' || v_号类 || ',', ',' || r_排班.号类 || ',') > 0 Then
            v_上班时间 := v_上班时间 || '+' || r_排班.排班;
            n_总已挂数 := n_总已挂数 + r_排班.已挂数;
            n_已挂数   := r_排班.已挂数;
            n_限号数   := r_排班.限号数;
            n_已约数   := r_排班.已约数;
            n_限约数   := r_排班.限约数;
            n_安排id   := Nvl(r_排班.安排id, 0);
            n_计划id   := Nvl(r_排班.计划id, 0);
            v_号码     := r_排班.号码;
            n_安排存在 := 1;
          
            If v_上班时间 Is Not Null Then
              If v_合作单位 Is Not Null Then
                If n_计划id <> 0 Then
                  Begin
                    Select 1
                    Into n_合约存在
                    From 合作单位计划控制
                    Where 计划id = n_计划id And 合作单位 = v_合作单位 And Rownum < 2;
                  Exception
                    When Others Then
                      n_合约存在 := 0;
                  End;
                Else
                  Begin
                    Select 1
                    Into n_合约存在
                    From 合作单位安排控制
                    Where 安排id = n_安排id And 合作单位 = v_合作单位 And Rownum < 2;
                  Exception
                    When Others Then
                      n_合约存在 := 0;
                  End;
                End If;
              End If;
            
              If v_合作单位 Is Null Or Nvl(n_合约存在, 0) = 0 Then
                If n_计划id <> 0 Then
                  Begin
                    Select Sum(数量)
                    Into n_合作单位数量
                    From 合作单位计划控制
                    Where 计划id = n_计划id And
                          限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                        '5', '周四', '6', '周五', '7', '周六', Null);
                  Exception
                    When Others Then
                      n_合作单位数量 := 0;
                  End;
                Else
                  Begin
                    Select Sum(数量)
                    Into n_合作单位数量
                    From 合作单位安排控制
                    Where 安排id = n_安排id And
                          限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                        '5', '周四', '6', '周五', '7', '周六', Null);
                  Exception
                    When Others Then
                      n_合作单位数量 := 0;
                  End;
                End If;
                Begin
                  Select Count(1)
                  Into n_合约已挂数
                  From 病人挂号记录
                  Where 号别 = v_号码 And 记录状态 = 1 And 合作单位 Is Not Null And 发生时间 Between Trunc(d_开始时间 + n_Daycount) And
                        Trunc(d_开始时间 + n_Daycount + 1) - 1 / 60 / 60 / 24;
                Exception
                  When Others Then
                    n_合约已挂数 := 0;
                End;
                If n_合作单位数量 = 0 Then
                  n_合作单位数量 := Null;
                End If;
                If Trunc(d_开始时间 + n_Daycount) > Trunc(Sysdate) Then
                  n_剩余数 := Nvl(n_剩余数, 0) + n_限约数 - n_已约数 - Nvl(n_合作单位数量, n_合约已挂数) + Nvl(n_合约已挂数, 0);
                Else
                  n_剩余数 := Nvl(n_剩余数, 0) + n_限号数 - n_已挂数 - Nvl(n_合作单位数量, n_合约已挂数) + Nvl(n_合约已挂数, 0);
                End If;
              Else
                --合约单位
                If n_计划id <> 0 Then
                  Begin
                    Select 1
                    Into n_禁用
                    From 合作单位计划控制
                    Where 计划id = n_计划id And
                          限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                        '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位 And 数量 = 0 And
                          Rownum < 2;
                  Exception
                    When Others Then
                      n_禁用 := 0;
                  End;
                Else
                  Begin
                    Select 1
                    Into n_禁用
                    From 合作单位安排控制
                    Where 安排id = n_安排id And
                          限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                        '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位 And 数量 = 0 And
                          Rownum < 2;
                  Exception
                    When Others Then
                      n_禁用 := 0;
                  End;
                End If;
                If Nvl(n_禁用, 0) = 0 Then
                  Begin
                    Select Count(1)
                    Into n_合约已挂数
                    From 病人挂号记录
                    Where 号别 = v_号码 And 记录状态 = 1 And 合作单位 = v_合作单位 And 发生时间 Between Trunc(d_开始时间 + n_Daycount) And
                          Trunc(d_开始时间 + n_Daycount + 1) - 1 / 60 / 60 / 24;
                  Exception
                    When Others Then
                      n_合约已挂数 := 0;
                  End;
                  If n_计划id <> 0 Then
                    Begin
                      Select Sum(数量)
                      Into n_合作单位数量
                      From 合作单位计划控制
                      Where 计划id = n_计划id And
                            限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                          '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位;
                    Exception
                      When Others Then
                        n_合作单位数量 := 0;
                    End;
                  Else
                    Begin
                      Select Sum(数量)
                      Into n_合作单位数量
                      From 合作单位安排控制
                      Where 安排id = n_安排id And
                            限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                          '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位;
                    Exception
                      When Others Then
                        n_合作单位数量 := 0;
                    End;
                  End If;
                  If n_合作单位数量 = 0 Then
                    n_合作单位数量 := Null;
                  End If;
                  n_剩余数 := Nvl(n_剩余数, 0) + Nvl(n_合作单位数量, n_合约已挂数) - Nvl(n_合约已挂数, 0);
                
                End If;
              End If;
            End If;
            n_合作单位数量 := 0;
            n_合约存在     := 0;
            n_禁用         := 0;
          End If;
        End Loop;
        v_上班时间 := Substr(v_上班时间, 2);
        If n_安排存在 = 1 Then
          v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_开始时间 + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                    '<SYHS>' || n_剩余数 || '</SYHS>' || '<SBSJ>' || v_上班时间 || '</SBSJ>' || '<YGS>' || n_总已挂数 || '</YGS>' ||
                    '</PB>';
        End If;
        n_Daycount := n_Daycount + 1;
      End Loop;
    End If;
    v_Temp := '<PBLIST>' || v_Temp || '</PBLIST>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Else
    --出诊表排班模式
    If Nvl(n_科室id, 0) = 0 Then
      --通过医生查找
      While (n_Daycount < n_查询天数) Loop
        If Trunc(d_开始时间 + n_Daycount) < To_Date(Substr(v_启用时间, 1, Instr(v_启用时间, ' ') - 1), 'yyyy-mm-dd') Then
          n_安排存在 := 0;
          v_上班时间 := Null;
          n_总已挂数 := 0;
          n_已挂数   := 0;
          n_剩余数   := 0;
          n_限号数   := 0;
          n_已约数   := 0;
          n_限约数   := 0;
          For r_排班 In (Select d_开始时间 + n_Daycount As 日期, a.排班, a.限号数, a.限约数, Nvl(Hz.已挂数, 0) As 已挂数, Nvl(Hz.已约数, 0) As 已约数,
                              a.安排id, a.计划id, a.号码, a.号类
                       
                       From (Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Ap.医生姓名, Nvl(Ap.医生id, 0) As 医生id, Ry.专业技术职务 As 职称,
                                     Ap.号码, Ap.安排id, Ap.计划id, Ap.排班, Ap.项目id, Fy.名称 As 项目名称, Ap.序号控制, Ap.限号数, Ap.限约数
                              
                              From (Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Ap.医生姓名, Ap.医生id, Ap.号码, Ap.Id As 安排id, 0 As 计划id,
                                            Ap.项目id, Nvl(Ap.序号控制, 0) As 序号控制,
                                            Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', Ap.周日, '2', Ap.周一, '3', Ap.周二, '4',
                                                    Ap.周三, '5', Ap.周四, '6', Ap.周五, '7', Ap.周六, Null) As 排班,
                                            Nvl(Xz.限约数, Xz.限号数) As 限约数, Nvl(Xz.限号数, 0) As 限号数
                                     From 挂号安排 Ap, 部门表 Bm, 挂号安排限制 Xz
                                     Where Ap.科室id = Bm.Id(+) And Ap.医生id = n_医生id And Ap.停用日期 Is Null And
                                           d_开始时间 + n_Daycount Between Nvl(Ap.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                           Nvl(Ap.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.安排id(+) = Ap.Id And
                                           Xz.限制项目(+) =
                                           Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4',
                                                  '周三', '5', '周四', '6', '周五', '7', '周六', Null) And Not Exists
                                      (Select Rownum
                                            From 挂号安排停用状态 Ty
                                            Where Ty.安排id = Ap.Id And d_开始时间 + n_Daycount Between Ty.开始停止时间 And Ty.结束停止时间) And
                                           Not Exists (Select Rownum
                                            From 挂号安排计划 Jh
                                            Where Jh.安排id = Ap.Id And Jh.审核时间 Is Not Null And
                                                  d_开始时间 + n_Daycount Between
                                                  Nvl(Jh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                  Nvl(Jh.失效时间, To_Date('3000-01-01', 'YYYY-MM-DD')))
                                     Union All
                                     Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Jh.医生姓名, Jh.医生id, Ap.号码, Ap.Id As 安排id,
                                            Jh.Id As 计划id, Jh.项目id, Nvl(Jh.序号控制, 0) As 序号控制,
                                            Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', Jh.周日, '2', Jh.周一, '3', Jh.周二, '4',
                                                    Jh.周三, '5', Jh.周四, '6', Jh.周五, '7', Jh.周六, Null) As 排班,
                                            Nvl(Xz.限约数, Xz.限号数) As 限约数, Nvl(Xz.限号数, 0) As 限号数
                                     From 挂号安排计划 Jh, 挂号安排 Ap, 部门表 Bm, 挂号计划限制 Xz
                                     Where Jh.安排id = Ap.Id And Ap.科室id = Bm.Id(+) And Ap.停用日期 Is Null And Jh.医生id = n_医生id And
                                           d_开始时间 + n_Daycount Between Nvl(Jh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                           Nvl(Jh.失效时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.计划id(+) = Jh.Id And
                                           Xz.限制项目(+) =
                                           Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4',
                                                  '周三', '5', '周四', '6', '周五', '7', '周六', Null) And Not Exists
                                      (Select Rownum
                                            From 挂号安排停用状态 Ty
                                            Where Ty.安排id = Ap.Id And d_开始时间 + n_Daycount Between Ty.开始停止时间 And Ty.结束停止时间) And
                                           (Jh.生效时间, Jh.安排id) =
                                           (Select Max(Sxjh.生效时间) As 生效时间, 安排id
                                            From 挂号安排计划 Sxjh
                                            Where Sxjh.审核时间 Is Not Null And
                                                  d_开始时间 + n_Daycount Between
                                                  Nvl(Sxjh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                  Nvl(Sxjh.失效时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Sxjh.安排id = Jh.安排id
                                            Group By Sxjh.安排id)) Ap, 部门表 Bm, 人员表 Ry, 收费项目目录 Fy
                              Where Ap.科室id = Bm.Id(+) And Ap.医生id = Ry.Id(+) And Ap.项目id = Fy.Id And 排班 Is Not Null) A,
                            病人挂号汇总 Hz, 收费价目 B
                       Where a.号码 = Hz.号码(+) And Hz.日期(+) = Trunc(d_开始时间 + n_Daycount) And a.项目id = b.收费细目id And
                             b.终止日期 > d_开始时间 + n_Daycount And b.执行日期 <= d_开始时间 + n_Daycount
                       
                       ) Loop
            If v_号类 Is Null Or Instr(',' || v_号类 || ',', ',' || r_排班.号类 || ',') > 0 Then
              v_上班时间 := v_上班时间 || '+' || r_排班.排班;
              n_总已挂数 := n_总已挂数 + r_排班.已挂数;
              n_已挂数   := r_排班.已挂数;
              n_限号数   := r_排班.限号数;
              n_已约数   := r_排班.已约数;
              n_限约数   := r_排班.限约数;
              n_安排id   := Nvl(r_排班.安排id, 0);
              n_计划id   := Nvl(r_排班.计划id, 0);
              v_号码     := r_排班.号码;
              n_安排存在 := 1;
              If v_上班时间 Is Not Null Then
                If v_合作单位 Is Not Null Then
                  If n_计划id <> 0 Then
                    Begin
                      Select 1
                      Into n_合约存在
                      From 合作单位计划控制
                      Where 计划id = n_计划id And 合作单位 = v_合作单位 And Rownum < 2;
                    Exception
                      When Others Then
                        n_合约存在 := 0;
                    End;
                  Else
                    Begin
                      Select 1
                      Into n_合约存在
                      From 合作单位安排控制
                      Where 安排id = n_安排id And 合作单位 = v_合作单位 And Rownum < 2;
                    Exception
                      When Others Then
                        n_合约存在 := 0;
                    End;
                  End If;
                End If;
              
                If v_合作单位 Is Null Or Nvl(n_合约存在, 0) = 0 Then
                  If n_计划id <> 0 Then
                    Begin
                      Select Sum(数量)
                      Into n_合作单位数量
                      From 合作单位计划控制
                      Where 计划id = n_计划id And
                            限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                          '5', '周四', '6', '周五', '7', '周六', Null);
                    Exception
                      When Others Then
                        n_合作单位数量 := 0;
                    End;
                  Else
                    Begin
                      Select Sum(数量)
                      Into n_合作单位数量
                      From 合作单位安排控制
                      Where 安排id = n_安排id And
                            限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                          '5', '周四', '6', '周五', '7', '周六', Null);
                    Exception
                      When Others Then
                        n_合作单位数量 := 0;
                    End;
                  End If;
                  Begin
                    Select Count(1)
                    Into n_合约已挂数
                    From 病人挂号记录
                    Where 号别 = v_号码 And 记录状态 = 1 And 合作单位 Is Not Null And 发生时间 Between Trunc(d_开始时间 + n_Daycount) And
                          Trunc(d_开始时间 + n_Daycount + 1) - 1 / 60 / 60 / 24;
                  Exception
                    When Others Then
                      n_合约已挂数 := 0;
                  End;
                  If n_合作单位数量 = 0 Then
                    n_合作单位数量 := Null;
                  End If;
                  If Trunc(d_开始时间 + n_Daycount) > Trunc(Sysdate) Then
                    n_剩余数 := Nvl(n_剩余数, 0) + n_限约数 - n_已约数 - Nvl(n_合作单位数量, n_合约已挂数) + Nvl(n_合约已挂数, 0);
                  Else
                    n_剩余数 := Nvl(n_剩余数, 0) + n_限号数 - n_已挂数 - Nvl(n_合作单位数量, n_合约已挂数) + Nvl(n_合约已挂数, 0);
                  End If;
                Else
                  --合约单位
                  If n_计划id <> 0 Then
                    Begin
                      Select 1
                      Into n_禁用
                      From 合作单位计划控制
                      Where 计划id = n_计划id And
                            限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                          '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位 And 数量 = 0 And
                            Rownum < 2;
                    Exception
                      When Others Then
                        n_禁用 := 0;
                    End;
                  Else
                    Begin
                      Select 1
                      Into n_禁用
                      From 合作单位安排控制
                      Where 安排id = n_安排id And
                            限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                          '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位 And 数量 = 0 And
                            Rownum < 2;
                    Exception
                      When Others Then
                        n_禁用 := 0;
                    End;
                  End If;
                  If Nvl(n_禁用, 0) = 0 Then
                    Begin
                      Select Count(1)
                      Into n_合约已挂数
                      From 病人挂号记录
                      Where 号别 = v_号码 And 记录状态 = 1 And 合作单位 = v_合作单位 And 发生时间 Between Trunc(d_开始时间 + n_Daycount) And
                            Trunc(d_开始时间 + n_Daycount + 1) - 1 / 60 / 60 / 24;
                    Exception
                      When Others Then
                        n_合约已挂数 := 0;
                    End;
                    If n_计划id <> 0 Then
                      Begin
                        Select Sum(数量)
                        Into n_合作单位数量
                        From 合作单位计划控制
                        Where 计划id = n_计划id And
                              限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4',
                                            '周三', '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位;
                      Exception
                        When Others Then
                          n_合作单位数量 := 0;
                      End;
                    Else
                      Begin
                        Select Sum(数量)
                        Into n_合作单位数量
                        From 合作单位安排控制
                        Where 安排id = n_安排id And
                              限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4',
                                            '周三', '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位;
                      Exception
                        When Others Then
                          n_合作单位数量 := 0;
                      End;
                    End If;
                    If n_合作单位数量 = 0 Then
                      n_合作单位数量 := Null;
                    End If;
                    n_剩余数 := Nvl(n_剩余数, 0) + Nvl(n_合作单位数量, n_合约已挂数) - Nvl(n_合约已挂数, 0);
                  
                  End If;
                End If;
              End If;
              n_合作单位数量 := 0;
              n_合约存在     := 0;
              n_禁用         := 0;
            End If;
          End Loop;
          v_上班时间 := Substr(v_上班时间, 2);
          If n_安排存在 = 1 Then
            v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_开始时间 + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                      '<SYHS>' || n_剩余数 || '</SYHS>' || '<SBSJ>' || v_上班时间 || '</SBSJ>' || '<YGS>' || n_总已挂数 ||
                      '</YGS>' || '</PB>';
          End If;
        Else
          n_安排存在 := 0;
          v_上班时间 := Null;
          n_总已挂数 := 0;
          n_已挂数   := 0;
          n_剩余数   := 0;
          n_限号数   := 0;
          n_已约数   := 0;
          n_限约数   := 0;
          If v_合作单位 Is Null Then
            --非合作单位
            If Trunc(d_开始时间 + n_Daycount) = Trunc(Sysdate) Then
              --当天挂号
              For r_出诊 In (Select 已挂数, 限号数, 上班时段
                           From 临床出诊记录 A
                           Where 出诊日期 = Trunc(d_开始时间 + n_Daycount) And 医生id = n_医生id And
                                 (a.开始时间 < Nvl(a.停诊开始时间, a.终止时间) Or a.终止时间 > Nvl(a.停诊终止时间, a.开始时间)) And
                                 Nvl(a.是否锁定, 0) = 0 And Exists
                            (Select 1
                                  From 临床出诊安排 M, 临床出诊表 N
                                  Where m.Id = a.安排id And m.出诊id = n.Id And n.发布时间 Is Not Null)) Loop
                n_已挂数   := n_已挂数 + Nvl(r_出诊.已挂数, 0);
                n_限号数   := n_限号数 + r_出诊.限号数;
                v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                n_安排存在 := 1;
              End Loop;
              If v_上班时间 Is Not Null Then
                v_上班时间 := Substr(v_上班时间, 2);
              End If;
              If n_安排存在 = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_开始时间 + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_限号数 - n_已挂数 || '</SYHS>' || '<SBSJ>' || v_上班时间 || '</SBSJ>' || '<YGS>' || n_已挂数 ||
                          '</YGS>' || '</PB>';
              End If;
            Else
              --预约挂号
              For r_出诊 In (Select 已约数, 限号数, 限约数, 上班时段
                           From 临床出诊记录 A
                           Where 出诊日期 = Trunc(d_开始时间 + n_Daycount) And 医生id = n_医生id And
                                 (a.开始时间 < Nvl(a.停诊开始时间, a.终止时间) Or a.终止时间 > Nvl(a.停诊终止时间, a.开始时间)) And
                                 Nvl(a.是否锁定, 0) = 0 And Exists
                            (Select 1
                                  From 临床出诊安排 M, 临床出诊表 N
                                  Where m.Id = a.安排id And m.出诊id = n.Id And n.发布时间 Is Not Null)) Loop
                n_已挂数   := n_已挂数 + Nvl(r_出诊.已约数, 0);
                n_限号数   := n_限号数 + Nvl(r_出诊.限约数, r_出诊.限号数);
                v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                n_安排存在 := 1;
              End Loop;
              If v_上班时间 Is Not Null Then
                v_上班时间 := Substr(v_上班时间, 2);
              End If;
              If n_安排存在 = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_开始时间 + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_限号数 - n_已挂数 || '</SYHS>' || '<SBSJ>' || v_上班时间 || '</SBSJ>' || '<YGS>' || n_已挂数 ||
                          '</YGS>' || '</PB>';
              End If;
            End If;
          Else
            --合作单位
            If Trunc(d_开始时间 + n_Daycount) = Trunc(Sysdate) Then
              For r_出诊 In (Select ID, 已挂数, 限号数, 限约数, 上班时段, 是否序号控制
                           From 临床出诊记录 A
                           Where 出诊日期 = Trunc(d_开始时间 + n_Daycount) And 医生id = n_医生id And
                                 (a.开始时间 < Nvl(a.停诊开始时间, a.终止时间) Or a.终止时间 > Nvl(a.停诊终止时间, a.开始时间)) And
                                 Nvl(a.是否锁定, 0) = 0 And Exists
                            (Select 1
                                  From 临床出诊安排 M, 临床出诊表 N
                                  Where m.Id = a.安排id And m.出诊id = n.Id And n.发布时间 Is Not Null)) Loop
                Begin
                  Select 控制方式
                  Into n_合约模式
                  From 临床出诊挂号控制记录
                  Where 记录id = r_出诊.Id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1 And Rownum < 2;
                Exception
                  When Others Then
                    n_合约模式 := 4;
                End;
                If n_合约模式 = 1 Or n_合约模式 = 2 Then
                  Select 数量
                  Into n_限号数
                  From 临床出诊挂号控制记录
                  Where 记录id = r_出诊.Id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1;
                  If n_合约模式 = 1 Then
                    n_限号数 := n_限号数 * Nvl(r_出诊.限约数, r_出诊.限号数) / 100;
                  End If;
                  Select Count(1)
                  Into n_已挂数
                  From 病人挂号记录
                  Where 出诊记录id = r_出诊.Id And 记录状态 = 1 And 合作单位 = v_合作单位;
                  If r_出诊.限号数 - r_出诊.已挂数 < n_限号数 - n_已挂数 Then
                    n_剩余数 := n_剩余数 + r_出诊.限号数 - r_出诊.已挂数;
                  Else
                    n_剩余数 := n_剩余数 + n_限号数 - n_已挂数;
                  End If;
                  n_总已挂数 := n_总已挂数 + n_已挂数;
                  v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                  n_安排存在 := 1;
                End If;
                If n_合约模式 = 3 Then
                  If r_出诊.是否序号控制 = 0 Then
                    n_总已挂数 := n_总已挂数 + Nvl(r_出诊.已挂数, 0);
                    n_剩余数   := n_剩余数 + r_出诊.限号数 - Nvl(r_出诊.已挂数, 0);
                    v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                    n_安排存在 := 1;
                  Else
                    For r_合作 In (Select 序号, 数量
                                 From 临床出诊挂号控制记录
                                 Where 记录id = r_出诊.Id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1) Loop
                      Begin
                        Select 1
                        Into n_Exists
                        From 临床出诊序号控制
                        Where 记录id = r_出诊.Id And 序号 = r_合作.序号 And Nvl(挂号状态, 0) = 0;
                      Exception
                        When Others Then
                          n_Exists := 0;
                      End;
                      If n_Exists = 1 Then
                        n_剩余数   := n_剩余数 + 1;
                        n_安排存在 := 1;
                      End If;
                    End Loop;
                    v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                    n_总已挂数 := n_总已挂数 + Nvl(r_出诊.已挂数, 0);
                  End If;
                End If;
                If n_合约模式 = 4 Then
                  n_总已挂数 := n_总已挂数 + Nvl(r_出诊.已挂数, 0);
                  n_剩余数   := n_剩余数 + r_出诊.限号数 - Nvl(r_出诊.已挂数, 0);
                  v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                  n_安排存在 := 1;
                End If;
              End Loop;
            
              If v_上班时间 Is Not Null Then
                v_上班时间 := Substr(v_上班时间, 2);
              End If;
              If n_安排存在 = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_开始时间 + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_剩余数 || '</SYHS>' || '<SBSJ>' || v_上班时间 || '</SBSJ>' || '<YGS>' || n_总已挂数 ||
                          '</YGS>' || '</PB>';
              End If;
            Else
              --非当天
              For r_出诊 In (Select ID, 已约数, 已挂数, 限号数, 限约数, 上班时段, 是否序号控制
                           From 临床出诊记录 A
                           Where 出诊日期 = Trunc(d_开始时间 + n_Daycount) And 医生id = n_医生id And
                                 (a.开始时间 < Nvl(a.停诊开始时间, a.终止时间) Or a.终止时间 > Nvl(a.停诊终止时间, a.开始时间)) And
                                 Nvl(a.是否锁定, 0) = 0 And Exists
                            (Select 1
                                  From 临床出诊安排 M, 临床出诊表 N
                                  Where m.Id = a.安排id And m.出诊id = n.Id And n.发布时间 Is Not Null)) Loop
                Begin
                  Select 控制方式
                  Into n_合约模式
                  From 临床出诊挂号控制记录
                  Where 记录id = r_出诊.Id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1 And Rownum < 2;
                Exception
                  When Others Then
                    n_合约模式 := 4;
                End;
                If n_合约模式 = 1 Or n_合约模式 = 2 Then
                  Select 数量
                  Into n_限号数
                  From 临床出诊挂号控制记录
                  Where 记录id = r_出诊.Id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1;
                  If n_合约模式 = 1 Then
                    n_限号数 := n_限号数 * Nvl(r_出诊.限约数, r_出诊.限号数) / 100;
                  End If;
                  Select Count(1)
                  Into n_已挂数
                  From 病人挂号记录
                  Where 出诊记录id = r_出诊.Id And 记录状态 = 1 And 合作单位 = v_合作单位;
                  If Nvl(r_出诊.限约数, r_出诊.限号数) - r_出诊.已约数 < n_限号数 - n_已挂数 Then
                    n_剩余数 := n_剩余数 + Nvl(r_出诊.限约数, r_出诊.限号数) - r_出诊.已约数;
                  Else
                    n_剩余数 := n_剩余数 + n_限号数 - n_已挂数;
                  End If;
                  n_总已挂数 := n_总已挂数 + n_已挂数;
                  v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                  n_安排存在 := 1;
                End If;
                If n_合约模式 = 3 Then
                  If r_出诊.是否序号控制 = 0 Then
                    --分时段非序号控制
                    For r_合作 In (Select 序号, 数量
                                 From 临床出诊挂号控制记录
                                 Where 记录id = r_出诊.Id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1) Loop
                      Select Count(1)
                      Into n_Exists
                      From 临床出诊序号控制
                      Where 预约顺序号 Is Not Null And 记录id = r_出诊.Id And 序号 = r_合作.序号 And Nvl(挂号状态, 0) <> 0;
                      If r_合作.数量 - n_Exists > 0 Then
                        n_剩余数   := n_剩余数 + r_合作.数量 - n_Exists;
                        n_总已挂数 := n_总已挂数 + n_Exists;
                        n_安排存在 := 1;
                      Else
                        n_总已挂数 := n_总已挂数 + n_Exists;
                        n_安排存在 := 1;
                      End If;
                    End Loop;
                    v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                  Else
                    For r_合作 In (Select 序号
                                 From 临床出诊挂号控制记录
                                 Where 记录id = r_出诊.Id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1) Loop
                      Begin
                        Select 1
                        Into n_Exists
                        From 临床出诊序号控制
                        Where 记录id = r_出诊.Id And 序号 = r_合作.序号 And Nvl(挂号状态, 0) = 0;
                      Exception
                        When Others Then
                          n_Exists := 0;
                      End;
                      If n_Exists = 1 Then
                        n_剩余数   := n_剩余数 + 1;
                        n_安排存在 := 1;
                      Else
                        n_总已挂数 := n_总已挂数 + 1;
                        n_安排存在 := 1;
                      End If;
                    End Loop;
                    v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                  End If;
                End If;
                If n_合约模式 = 4 Then
                  n_总已挂数 := n_总已挂数 + Nvl(r_出诊.已约数, 0);
                  n_剩余数   := n_剩余数 + r_出诊.限约数 - Nvl(r_出诊.已约数, 0);
                  v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                  n_安排存在 := 1;
                End If;
              End Loop;
              If v_上班时间 Is Not Null Then
                v_上班时间 := Substr(v_上班时间, 2);
              End If;
              If n_安排存在 = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_开始时间 + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_剩余数 || '</SYHS>' || '<SBSJ>' || v_上班时间 || '</SBSJ>' || '<YGS>' || n_总已挂数 ||
                          '</YGS>' || '</PB>';
              End If;
            End If;
          End If;
        End If;
        n_Daycount := n_Daycount + 1;
      End Loop;
    Else
      --通过科室+医生查找
      While (n_Daycount < n_查询天数) Loop
        If Trunc(d_开始时间 + n_Daycount) < To_Date(Substr(v_启用时间, 1, Instr(v_启用时间, ' ') - 1), 'yyyy-mm-dd') Then
          v_上班时间 := Null;
          n_总已挂数 := 0;
          n_已挂数   := 0;
          n_剩余数   := 0;
          n_限号数   := 0;
          n_已约数   := 0;
          n_限约数   := 0;
          n_安排存在 := 0;
          For r_排班 In (Select d_开始时间 + n_Daycount As 日期, a.排班, a.限号数, a.限约数, Nvl(Hz.已挂数, 0) As 已挂数, Nvl(Hz.已约数, 0) As 已约数,
                              a.安排id, a.计划id, a.号码, a.号类
                       
                       From (Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Ap.医生姓名, Nvl(Ap.医生id, 0) As 医生id, Ry.专业技术职务 As 职称,
                                     Ap.号码, Ap.安排id, Ap.计划id, Ap.排班, Ap.项目id, Fy.名称 As 项目名称, Ap.序号控制, Ap.限号数, Ap.限约数
                              
                              From (Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Ap.医生姓名, Ap.医生id, Ap.号码, Ap.Id As 安排id, 0 As 计划id,
                                            Ap.项目id, Nvl(Ap.序号控制, 0) As 序号控制,
                                            Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', Ap.周日, '2', Ap.周一, '3', Ap.周二, '4',
                                                    Ap.周三, '5', Ap.周四, '6', Ap.周五, '7', Ap.周六, Null) As 排班,
                                            Nvl(Xz.限约数, Xz.限号数) As 限约数, Nvl(Xz.限号数, 0) As 限号数
                                     From 挂号安排 Ap, 部门表 Bm, 挂号安排限制 Xz
                                     Where Ap.科室id = Bm.Id(+) And Ap.医生id = n_医生id And Ap.科室id = n_科室id And Ap.停用日期 Is Null And
                                           d_开始时间 + n_Daycount Between Nvl(Ap.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                           Nvl(Ap.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.安排id(+) = Ap.Id And
                                           Xz.限制项目(+) =
                                           Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4',
                                                  '周三', '5', '周四', '6', '周五', '7', '周六', Null) And Not Exists
                                      (Select Rownum
                                            From 挂号安排停用状态 Ty
                                            Where Ty.安排id = Ap.Id And d_开始时间 + n_Daycount Between Ty.开始停止时间 And Ty.结束停止时间) And
                                           Not Exists (Select Rownum
                                            From 挂号安排计划 Jh
                                            Where Jh.安排id = Ap.Id And Jh.审核时间 Is Not Null And
                                                  d_开始时间 + n_Daycount Between
                                                  Nvl(Jh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                  Nvl(Jh.失效时间, To_Date('3000-01-01', 'YYYY-MM-DD')))
                                     Union All
                                     Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Jh.医生姓名, Jh.医生id, Ap.号码, Ap.Id As 安排id,
                                            Jh.Id As 计划id, Jh.项目id, Nvl(Jh.序号控制, 0) As 序号控制,
                                            Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', Jh.周日, '2', Jh.周一, '3', Jh.周二, '4',
                                                    Jh.周三, '5', Jh.周四, '6', Jh.周五, '7', Jh.周六, Null) As 排班,
                                            Nvl(Xz.限约数, Xz.限号数) As 限约数, Nvl(Xz.限号数, 0) As 限号数
                                     From 挂号安排计划 Jh, 挂号安排 Ap, 部门表 Bm, 挂号计划限制 Xz
                                     Where Jh.安排id = Ap.Id And Ap.科室id = Bm.Id(+) And Ap.停用日期 Is Null And Jh.医生id = n_医生id And
                                           Ap.科室id = n_科室id And
                                           d_开始时间 + n_Daycount Between Nvl(Jh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                           Nvl(Jh.失效时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.计划id(+) = Jh.Id And
                                           Xz.限制项目(+) =
                                           Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4',
                                                  '周三', '5', '周四', '6', '周五', '7', '周六', Null) And Not Exists
                                      (Select Rownum
                                            From 挂号安排停用状态 Ty
                                            Where Ty.安排id = Ap.Id And d_开始时间 + n_Daycount Between Ty.开始停止时间 And Ty.结束停止时间) And
                                           (Jh.生效时间, Jh.安排id) =
                                           (Select Max(Sxjh.生效时间) As 生效时间, 安排id
                                            From 挂号安排计划 Sxjh
                                            Where Sxjh.审核时间 Is Not Null And
                                                  d_开始时间 + n_Daycount Between
                                                  Nvl(Sxjh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                  Nvl(Sxjh.失效时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Sxjh.安排id = Jh.安排id
                                            Group By Sxjh.安排id)) Ap, 部门表 Bm, 人员表 Ry, 收费项目目录 Fy
                              Where Ap.科室id = Bm.Id(+) And Ap.医生id = Ry.Id(+) And Ap.项目id = Fy.Id And 排班 Is Not Null) A,
                            病人挂号汇总 Hz, 收费价目 B
                       Where a.号码 = Hz.号码(+) And Hz.日期(+) = Trunc(d_开始时间 + n_Daycount) And a.项目id = b.收费细目id And
                             b.终止日期 > d_开始时间 + n_Daycount And b.执行日期 <= d_开始时间 + n_Daycount
                       
                       ) Loop
            If v_号类 Is Null Or Instr(',' || v_号类 || ',', ',' || r_排班.号类 || ',') > 0 Then
              v_上班时间 := v_上班时间 || '+' || r_排班.排班;
              n_总已挂数 := n_总已挂数 + r_排班.已挂数;
              n_已挂数   := r_排班.已挂数;
              n_限号数   := r_排班.限号数;
              n_已约数   := r_排班.已约数;
              n_限约数   := r_排班.限约数;
              n_安排id   := Nvl(r_排班.安排id, 0);
              n_计划id   := Nvl(r_排班.计划id, 0);
              v_号码     := r_排班.号码;
              n_安排存在 := 1;
            
              If v_上班时间 Is Not Null Then
                If v_合作单位 Is Not Null Then
                  If n_计划id <> 0 Then
                    Begin
                      Select 1
                      Into n_合约存在
                      From 合作单位计划控制
                      Where 计划id = n_计划id And 合作单位 = v_合作单位 And Rownum < 2;
                    Exception
                      When Others Then
                        n_合约存在 := 0;
                    End;
                  Else
                    Begin
                      Select 1
                      Into n_合约存在
                      From 合作单位安排控制
                      Where 安排id = n_安排id And 合作单位 = v_合作单位 And Rownum < 2;
                    Exception
                      When Others Then
                        n_合约存在 := 0;
                    End;
                  End If;
                End If;
              
                If v_合作单位 Is Null Or Nvl(n_合约存在, 0) = 0 Then
                  If n_计划id <> 0 Then
                    Begin
                      Select Sum(数量)
                      Into n_合作单位数量
                      From 合作单位计划控制
                      Where 计划id = n_计划id And
                            限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                          '5', '周四', '6', '周五', '7', '周六', Null);
                    Exception
                      When Others Then
                        n_合作单位数量 := 0;
                    End;
                  Else
                    Begin
                      Select Sum(数量)
                      Into n_合作单位数量
                      From 合作单位安排控制
                      Where 安排id = n_安排id And
                            限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                          '5', '周四', '6', '周五', '7', '周六', Null);
                    Exception
                      When Others Then
                        n_合作单位数量 := 0;
                    End;
                  End If;
                  Begin
                    Select Count(1)
                    Into n_合约已挂数
                    From 病人挂号记录
                    Where 号别 = v_号码 And 记录状态 = 1 And 合作单位 Is Not Null And 发生时间 Between Trunc(d_开始时间 + n_Daycount) And
                          Trunc(d_开始时间 + n_Daycount + 1) - 1 / 60 / 60 / 24;
                  Exception
                    When Others Then
                      n_合约已挂数 := 0;
                  End;
                  If n_合作单位数量 = 0 Then
                    n_合作单位数量 := Null;
                  End If;
                  If Trunc(d_开始时间 + n_Daycount) > Trunc(Sysdate) Then
                    n_剩余数 := Nvl(n_剩余数, 0) + n_限约数 - n_已约数 - Nvl(n_合作单位数量, n_合约已挂数) + Nvl(n_合约已挂数, 0);
                  Else
                    n_剩余数 := Nvl(n_剩余数, 0) + n_限号数 - n_已挂数 - Nvl(n_合作单位数量, n_合约已挂数) + Nvl(n_合约已挂数, 0);
                  End If;
                Else
                  --合约单位
                  If n_计划id <> 0 Then
                    Begin
                      Select 1
                      Into n_禁用
                      From 合作单位计划控制
                      Where 计划id = n_计划id And
                            限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                          '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位 And 数量 = 0 And
                            Rownum < 2;
                    Exception
                      When Others Then
                        n_禁用 := 0;
                    End;
                  Else
                    Begin
                      Select 1
                      Into n_禁用
                      From 合作单位安排控制
                      Where 安排id = n_安排id And
                            限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                          '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位 And 数量 = 0 And
                            Rownum < 2;
                    Exception
                      When Others Then
                        n_禁用 := 0;
                    End;
                  End If;
                  If Nvl(n_禁用, 0) = 0 Then
                    Begin
                      Select Count(1)
                      Into n_合约已挂数
                      From 病人挂号记录
                      Where 号别 = v_号码 And 记录状态 = 1 And 合作单位 = v_合作单位 And 发生时间 Between Trunc(d_开始时间 + n_Daycount) And
                            Trunc(d_开始时间 + n_Daycount + 1) - 1 / 60 / 60 / 24;
                    Exception
                      When Others Then
                        n_合约已挂数 := 0;
                    End;
                    If n_计划id <> 0 Then
                      Begin
                        Select Sum(数量)
                        Into n_合作单位数量
                        From 合作单位计划控制
                        Where 计划id = n_计划id And
                              限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4',
                                            '周三', '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位;
                      Exception
                        When Others Then
                          n_合作单位数量 := 0;
                      End;
                    Else
                      Begin
                        Select Sum(数量)
                        Into n_合作单位数量
                        From 合作单位安排控制
                        Where 安排id = n_安排id And
                              限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4',
                                            '周三', '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位;
                      Exception
                        When Others Then
                          n_合作单位数量 := 0;
                      End;
                    End If;
                    If n_合作单位数量 = 0 Then
                      n_合作单位数量 := Null;
                    End If;
                    n_剩余数 := Nvl(n_剩余数, 0) + Nvl(n_合作单位数量, n_合约已挂数) - Nvl(n_合约已挂数, 0);
                  
                  End If;
                End If;
              End If;
              n_合作单位数量 := 0;
              n_合约存在     := 0;
              n_禁用         := 0;
            End If;
          End Loop;
          v_上班时间 := Substr(v_上班时间, 2);
          If n_安排存在 = 1 Then
            v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_开始时间 + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                      '<SYHS>' || n_剩余数 || '</SYHS>' || '<SBSJ>' || v_上班时间 || '</SBSJ>' || '<YGS>' || n_总已挂数 ||
                      '</YGS>' || '</PB>';
          End If;
        Else
          n_安排存在 := 0;
          v_上班时间 := Null;
          n_总已挂数 := 0;
          n_已挂数   := 0;
          n_剩余数   := 0;
          n_限号数   := 0;
          n_已约数   := 0;
          n_限约数   := 0;
          If v_合作单位 Is Null Then
            --非合作单位
            If Trunc(d_开始时间 + n_Daycount) = Trunc(Sysdate) Then
              --当天挂号
              For r_出诊 In (Select 已挂数, 限号数, 上班时段
                           From 临床出诊记录 A
                           Where 出诊日期 = Trunc(d_开始时间 + n_Daycount) And 医生id = n_医生id And 科室id = n_科室id And
                                 (a.开始时间 < Nvl(a.停诊开始时间, a.终止时间) Or a.终止时间 > Nvl(a.停诊终止时间, a.开始时间)) And
                                 Nvl(a.是否锁定, 0) = 0 And Exists
                            (Select 1
                                  From 临床出诊安排 M, 临床出诊表 N
                                  Where m.Id = a.安排id And m.出诊id = n.Id And n.发布时间 Is Not Null)) Loop
                n_已挂数   := n_已挂数 + Nvl(r_出诊.已挂数, 0);
                n_限号数   := n_限号数 + r_出诊.限号数;
                v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                n_安排存在 := 1;
              End Loop;
              If v_上班时间 Is Not Null Then
                v_上班时间 := Substr(v_上班时间, 2);
              End If;
              If n_安排存在 = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_开始时间 + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_限号数 - n_已挂数 || '</SYHS>' || '<SBSJ>' || v_上班时间 || '</SBSJ>' || '<YGS>' || n_已挂数 ||
                          '</YGS>' || '</PB>';
              End If;
            Else
              --预约挂号
              For r_出诊 In (Select 已约数, 限号数, 限约数, 上班时段
                           From 临床出诊记录 A
                           Where 出诊日期 = Trunc(d_开始时间 + n_Daycount) And 医生id = n_医生id And 科室id = n_科室id And
                                 (a.开始时间 < Nvl(a.停诊开始时间, a.终止时间) Or a.终止时间 > Nvl(a.停诊终止时间, a.开始时间)) And
                                 Nvl(a.是否锁定, 0) = 0 And Exists
                            (Select 1
                                  From 临床出诊安排 M, 临床出诊表 N
                                  Where m.Id = a.安排id And m.出诊id = n.Id And n.发布时间 Is Not Null)) Loop
                n_已挂数   := n_已挂数 + Nvl(r_出诊.已约数, 0);
                n_限号数   := n_限号数 + Nvl(r_出诊.限约数, r_出诊.限号数);
                v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                n_安排存在 := 1;
              End Loop;
              If v_上班时间 Is Not Null Then
                v_上班时间 := Substr(v_上班时间, 2);
              End If;
              If n_安排存在 = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_开始时间 + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_限号数 - n_已挂数 || '</SYHS>' || '<SBSJ>' || v_上班时间 || '</SBSJ>' || '<YGS>' || n_已挂数 ||
                          '</YGS>' || '</PB>';
              End If;
            End If;
          Else
            --合作单位
            If Trunc(d_开始时间 + n_Daycount) = Trunc(Sysdate) Then
              For r_出诊 In (Select ID, 已挂数, 限号数, 限约数, 上班时段, 是否序号控制
                           From 临床出诊记录 A
                           Where 出诊日期 = Trunc(d_开始时间 + n_Daycount) And 医生id = n_医生id And 科室id = n_科室id And
                                 (a.开始时间 < Nvl(a.停诊开始时间, a.终止时间) Or a.终止时间 > Nvl(a.停诊终止时间, a.开始时间)) And
                                 Nvl(a.是否锁定, 0) = 0 And Exists
                            (Select 1
                                  From 临床出诊安排 M, 临床出诊表 N
                                  Where m.Id = a.安排id And m.出诊id = n.Id And n.发布时间 Is Not Null)) Loop
                Begin
                  Select 控制方式
                  Into n_合约模式
                  From 临床出诊挂号控制记录
                  Where 记录id = r_出诊.Id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1 And Rownum < 2;
                Exception
                  When Others Then
                    n_合约模式 := 4;
                End;
                If n_合约模式 = 1 Or n_合约模式 = 2 Then
                  Select 数量
                  Into n_限号数
                  From 临床出诊挂号控制记录
                  Where 记录id = r_出诊.Id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1;
                  If n_合约模式 = 1 Then
                    n_限号数 := n_限号数 * Nvl(r_出诊.限约数, r_出诊.限号数) / 100;
                  End If;
                  Select Count(1)
                  Into n_已挂数
                  From 病人挂号记录
                  Where 出诊记录id = r_出诊.Id And 记录状态 = 1 And 合作单位 = v_合作单位;
                  If r_出诊.限号数 - r_出诊.已挂数 < n_限号数 - n_已挂数 Then
                    n_剩余数 := n_剩余数 + r_出诊.限号数 - r_出诊.已挂数;
                  Else
                    n_剩余数 := n_剩余数 + n_限号数 - n_已挂数;
                  End If;
                  n_总已挂数 := n_总已挂数 + n_已挂数;
                  v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                  n_安排存在 := 1;
                End If;
                If n_合约模式 = 3 Then
                  If r_出诊.是否序号控制 = 0 Then
                    n_总已挂数 := n_总已挂数 + Nvl(r_出诊.已挂数, 0);
                    n_剩余数   := n_剩余数 + r_出诊.限号数 - Nvl(r_出诊.已挂数, 0);
                    v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                    n_安排存在 := 1;
                  Else
                    For r_合作 In (Select 序号, 数量
                                 From 临床出诊挂号控制记录
                                 Where 记录id = r_出诊.Id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1) Loop
                      Begin
                        Select 1
                        Into n_Exists
                        From 临床出诊序号控制
                        Where 记录id = r_出诊.Id And 序号 = r_合作.序号 And Nvl(挂号状态, 0) = 0;
                      Exception
                        When Others Then
                          n_Exists := 0;
                      End;
                      If n_Exists = 1 Then
                        n_剩余数   := n_剩余数 + 1;
                        n_安排存在 := 1;
                      End If;
                    End Loop;
                    v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                    n_总已挂数 := n_总已挂数 + Nvl(r_出诊.已挂数, 0);
                  End If;
                End If;
                If n_合约模式 = 4 Then
                  n_总已挂数 := n_总已挂数 + Nvl(r_出诊.已挂数, 0);
                  n_剩余数   := n_剩余数 + r_出诊.限号数 - Nvl(r_出诊.已挂数, 0);
                  v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                  n_安排存在 := 1;
                End If;
              End Loop;
            
              If v_上班时间 Is Not Null Then
                v_上班时间 := Substr(v_上班时间, 2);
              End If;
              If n_安排存在 = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_开始时间 + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_剩余数 || '</SYHS>' || '<SBSJ>' || v_上班时间 || '</SBSJ>' || '<YGS>' || n_总已挂数 ||
                          '</YGS>' || '</PB>';
              End If;
            Else
              --非当天
              For r_出诊 In (Select ID, 已约数, 已挂数, 限号数, 限约数, 上班时段, 是否序号控制
                           From 临床出诊记录 A
                           Where 出诊日期 = Trunc(d_开始时间 + n_Daycount) And 医生id = n_医生id And 科室id = n_科室id And
                                 (a.开始时间 < Nvl(a.停诊开始时间, a.终止时间) Or a.终止时间 > Nvl(a.停诊终止时间, a.开始时间)) And
                                 Nvl(a.是否锁定, 0) = 0 And Exists
                            (Select 1
                                  From 临床出诊安排 M, 临床出诊表 N
                                  Where m.Id = a.安排id And m.出诊id = n.Id And n.发布时间 Is Not Null)) Loop
                Begin
                  Select 控制方式
                  Into n_合约模式
                  From 临床出诊挂号控制记录
                  Where 记录id = r_出诊.Id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1 And Rownum < 2;
                Exception
                  When Others Then
                    n_合约模式 := 4;
                End;
                If n_合约模式 = 1 Or n_合约模式 = 2 Then
                  Select 数量
                  Into n_限号数
                  From 临床出诊挂号控制记录
                  Where 记录id = r_出诊.Id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1;
                  If n_合约模式 = 1 Then
                    n_限号数 := n_限号数 * Nvl(r_出诊.限约数, r_出诊.限号数) / 100;
                  End If;
                  Select Count(1)
                  Into n_已挂数
                  From 病人挂号记录
                  Where 出诊记录id = r_出诊.Id And 记录状态 = 1 And 合作单位 = v_合作单位;
                  If Nvl(r_出诊.限约数, r_出诊.限号数) - r_出诊.已约数 < n_限号数 - n_已挂数 Then
                    n_剩余数 := n_剩余数 + Nvl(r_出诊.限约数, r_出诊.限号数) - r_出诊.已约数;
                  Else
                    n_剩余数 := n_剩余数 + n_限号数 - n_已挂数;
                  End If;
                  n_总已挂数 := n_总已挂数 + n_已挂数;
                  v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                  n_安排存在 := 1;
                End If;
                If n_合约模式 = 3 Then
                  If r_出诊.是否序号控制 = 0 Then
                    --分时段非序号控制
                    For r_合作 In (Select 序号, 数量
                                 From 临床出诊挂号控制记录
                                 Where 记录id = r_出诊.Id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1) Loop
                      Select Count(1)
                      Into n_Exists
                      From 临床出诊序号控制
                      Where 预约顺序号 Is Not Null And 记录id = r_出诊.Id And 序号 = r_合作.序号 And Nvl(挂号状态, 0) <> 0;
                      If r_合作.数量 - n_Exists > 0 Then
                        n_剩余数   := n_剩余数 + r_合作.数量 - n_Exists;
                        n_总已挂数 := n_总已挂数 + n_Exists;
                        n_安排存在 := 1;
                      Else
                        n_总已挂数 := n_总已挂数 + n_Exists;
                        n_安排存在 := 1;
                      End If;
                    End Loop;
                    v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                  Else
                    For r_合作 In (Select 序号
                                 From 临床出诊挂号控制记录
                                 Where 记录id = r_出诊.Id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1) Loop
                      Begin
                        Select 1
                        Into n_Exists
                        From 临床出诊序号控制
                        Where 记录id = r_出诊.Id And 序号 = r_合作.序号 And Nvl(挂号状态, 0) = 0;
                      Exception
                        When Others Then
                          n_Exists := 0;
                      End;
                      If n_Exists = 1 Then
                        n_剩余数   := n_剩余数 + 1;
                        n_安排存在 := 1;
                      Else
                        n_总已挂数 := n_总已挂数 + 1;
                        n_安排存在 := 1;
                      End If;
                    End Loop;
                    v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                  End If;
                End If;
                If n_合约模式 = 4 Then
                  n_总已挂数 := n_总已挂数 + Nvl(r_出诊.已约数, 0);
                  n_剩余数   := n_剩余数 + r_出诊.限约数 - Nvl(r_出诊.已约数, 0);
                  v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                  n_安排存在 := 1;
                End If;
              End Loop;
              If v_上班时间 Is Not Null Then
                v_上班时间 := Substr(v_上班时间, 2);
              End If;
              If n_安排存在 = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_开始时间 + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_剩余数 || '</SYHS>' || '<SBSJ>' || v_上班时间 || '</SBSJ>' || '<YGS>' || n_总已挂数 ||
                          '</YGS>' || '</PB>';
              End If;
            End If;
          End If;
        End If;
        n_Daycount := n_Daycount + 1;
      End Loop;
    End If;
    v_Temp := '<PBLIST>' || v_Temp || '</PBLIST>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  End If;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Docarrange;
/



Create Or Replace Procedure Zl_Third_Lockno
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:HIS锁号
  --入参:Xml_In:
  --<IN>
  --  <HM>5</HM>           //号码
  --  <CZJLID>1</CZJLID>       //出诊记录ID,出诊表排班模式下传入
  --  <RQ>2013-11-21 09:00</RQ>     //预约时间
  --  <CZ>1</CZ>           //操作
  --  <HX></HX>          //号序
  --  <HZDW>支付宝</HZDW>   //合作单位
  --  <JQM>机器名</JQM>        //机器名
  --</IN>

  --出参:Xml_Out
  --<OUTPUT>
  -- <HX>号序</HX>          //锁号操作并且成功时返回
  -- 错误信息  //出错时返回
  --</ OUTPUT>
  --------------------------------------------------------------------------------------------------
  v_号码         挂号安排.号码%Type;
  d_日期         Date;
  n_操作类型     Number(3);
  n_序号控制     Number(3);
  n_存在         Number(3);
  n_分时段       Number(3);
  n_限号数       挂号安排限制.限号数%Type;
  n_安排id       挂号安排.Id%Type;
  n_计划id       挂号安排计划.Id%Type;
  n_号序         挂号序号状态.序号%Type;
  v_星期         挂号安排限制.限制项目%Type;
  v_操作员姓名   挂号序号状态.操作员姓名%Type;
  v_机器名       挂号序号状态.机器名%Type;
  v_验证姓名     挂号序号状态.操作员姓名%Type;
  v_验证机器名   挂号序号状态.机器名%Type;
  n_状态         挂号序号状态.状态%Type;
  v_合作单位     合作单位安排控制.合作单位%Type;
  n_合约模式     Number(3);
  n_启用合作单位 Number(3);
  v_Temp         Varchar2(32767); --临时XML
  v_Optemp       Varchar2(300);
  x_Templet      Xmltype; --模板XML
  v_Err_Msg      Varchar2(200);
  n_Exists       Number(2);
  n_记录id       临床出诊记录.Id%Type;
  n_序号         临床出诊序号控制.序号%Type;
  n_数量         临床出诊序号控制.数量%Type;
  n_顺序号       临床出诊序号控制.预约顺序号%Type;
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/HM'), To_Date(Extractvalue(Value(A), 'IN/RQ'), 'YYYY-MM-DD hh24:mi:ss'),
         Extractvalue(Value(A), 'IN/CZ'), Extractvalue(Value(A), 'IN/HX'), Extractvalue(Value(A), 'IN/HZDW'),
         Extractvalue(Value(A), 'IN/JQM'), Extractvalue(Value(A), 'IN/CZJLID')
  Into v_号码, d_日期, n_操作类型, n_号序, v_合作单位, v_机器名, n_记录id
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  If v_机器名 Is Null Then
    Select Terminal Into v_机器名 From V$session Where Audsid = Userenv('sessionid');
  End If;
  v_Optemp := Zl_Identity(1);
  Select Substr(v_Optemp, Instr(v_Optemp, ',') + 1) Into v_Optemp From Dual;
  Select Substr(v_Optemp, Instr(v_Optemp, ',') + 1) Into v_操作员姓名 From Dual;

  If n_记录id Is Null Then
    If n_操作类型 = 0 Then
      --解锁
      Begin
        Select 1
        Into n_Exists
        From 挂号序号状态
        Where 机器名 = v_机器名 And 操作员姓名 = v_操作员姓名 And 状态 = 5 And 序号 = n_号序 And Trunc(日期) = Trunc(d_日期) And 号码 = v_号码 And
              Rownum < 2;
      Exception
        When Others Then
          n_Exists := 0;
      End;
      If n_Exists = 1 Then
        Delete 挂号序号状态
        Where 机器名 = v_机器名 And 操作员姓名 = v_操作员姓名 And 状态 = 5 And 序号 = n_号序 And Trunc(日期) = Trunc(d_日期) And 号码 = v_号码;
        v_Temp := '<HX>' || n_号序 || '</HX>';
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
      Else
        v_Temp := '没有发现需要解锁的序号';
        Raise Err_Item;
      End If;
    End If;
  
    If n_操作类型 = 1 Then
      --锁号
      Select Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六',
                     Null)
      Into v_星期
      From Dual;
      Begin
        Select 序号控制, ID
        Into n_序号控制, n_计划id
        From (Select 序号控制, ID
               From 挂号安排计划
               Where 号码 = v_号码 And d_日期 Between Nvl(生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                     Nvl(失效时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And 审核时间 Is Not Null
               Order By 生效时间 Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          Select 序号控制, ID Into n_序号控制, n_安排id From 挂号安排 Where 号码 = v_号码;
      End;
      If n_序号控制 = 1 Then
        If Nvl(n_计划id, 0) <> 0 Then
          Begin
            Select 1 Into n_分时段 From 挂号计划时段 Where 星期 = v_星期 And 计划id = n_计划id And Rownum < 2;
          Exception
            When Others Then
              n_分时段 := 0;
          End;
          Begin
            Select 1
            Into n_启用合作单位
            From 合作单位计划控制
            Where 限制项目 = v_星期 And 计划id = n_计划id And 合作单位 = v_合作单位 And Rownum < 2;
          Exception
            When Others Then
              n_启用合作单位 := 0;
          End;
          Begin
            Select 1, a.状态, a.操作员姓名, a.机器名
            Into n_存在, n_状态, v_验证姓名, v_验证机器名
            From 挂号序号状态 A, 挂号计划时段 B
            Where a.号码 = v_号码 And Trunc(a.日期) = Trunc(d_日期) And a.序号 = b.序号 And b.计划id = n_计划id And b.星期 = v_星期 And
                  To_Char(b.开始时间, 'hh24:mi') = To_Char(d_日期, 'hh24:mi') And Rownum < 2;
          Exception
            When Others Then
              n_存在 := 0;
          End;
          If n_存在 = 1 Then
            If n_状态 = 5 And v_验证姓名 = v_操作员姓名 And v_机器名 = v_验证机器名 Then
              Null;
            Else
              --传入时间的序号已经被使用
              v_Temp := '传入时间' || d_日期 || '的序号已被使用';
              Raise Err_Item;
            End If;
          Else
            If n_分时段 = 1 Then
              Begin
                Select 序号
                Into n_号序
                From 挂号计划时段
                Where 计划id = n_计划id And 星期 = v_星期 And To_Char(开始时间, 'hh24:mi') = To_Char(d_日期, 'hh24:mi') And
                      Rownum < 2;
              Exception
                When Others Then
                  Select Max(序号) + 1
                  Into n_号序
                  From (Select Distinct 序号
                         From 挂号计划时段
                         Where 计划id = n_计划id And 星期 = v_星期
                         Union
                         Select Distinct 序号 From 挂号序号状态 Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期));
                
              End;
              Begin
                Select 1 Into n_存在 From 挂号序号状态 Where 号码 = v_号码 And 日期 = d_日期 And 序号 = n_号序;
              Exception
                When Others Then
                  Insert Into 挂号序号状态
                    (号码, 日期, 序号, 状态, 操作员姓名, 备注, 登记时间, 机器名)
                  Values
                    (v_号码, d_日期, n_号序, 5, v_操作员姓名, '移动锁号', Sysdate, v_机器名);
              End;
              v_Temp := '<HX>' || n_号序 || '</HX>';
              Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
            Else
              If v_合作单位 Is Null Or n_启用合作单位 = 0 Then
                If Trunc(d_日期) = Trunc(Sysdate) Then
                  n_号序 := 1;
                  Select 限号数 Into n_限号数 From 挂号计划限制 Where 计划id = n_计划id And 限制项目 = v_星期;
                  For r_序号 In (Select 序号, 状态, 操作员姓名, 机器名
                               From 挂号序号状态
                               Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期)
                               Order By 序号) Loop
                    Exit When r_序号.状态 = 5 And r_序号.操作员姓名 = v_操作员姓名 And r_序号.机器名 = v_机器名;
                    If r_序号.序号 = n_号序 Then
                      n_号序 := n_号序 + 1;
                    End If;
                  End Loop;
                  If n_号序 > n_限号数 Then
                    v_Temp := '传入号别' || v_号码 || '的所有序号已被用完';
                    Raise Err_Item;
                  Else
                    Begin
                      Select 1
                      Into n_存在
                      From 挂号序号状态
                      Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期) And 序号 = n_号序;
                    Exception
                      When Others Then
                        Insert Into 挂号序号状态
                          (号码, 日期, 序号, 状态, 操作员姓名, 备注, 登记时间, 机器名)
                        Values
                          (v_号码, Trunc(d_日期), n_号序, 5, v_操作员姓名, '移动锁号', Sysdate, v_机器名);
                    End;
                    v_Temp := '<HX>' || n_号序 || '</HX>';
                    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
                  End If;
                Else
                  n_号序 := 1;
                  Select 限号数 Into n_限号数 From 挂号计划限制 Where 计划id = n_计划id And 限制项目 = v_星期;
                  For r_序号 In (Select 序号, 状态, 操作员姓名, 机器名
                               From 挂号序号状态
                               Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期)
                               
                               Union
                               Select 序号, Null, Null, Null
                               From 合作单位计划控制
                               Where 计划id = n_计划id And 限制项目 = v_星期 And 数量 <> 0
                               Order By 序号) Loop
                    Exit When r_序号.状态 = 5 And r_序号.操作员姓名 = v_操作员姓名 And r_序号.机器名 = v_机器名;
                    If r_序号.序号 = n_号序 Then
                      n_号序 := n_号序 + 1;
                    End If;
                  End Loop;
                  If n_号序 > n_限号数 Then
                    v_Temp := '传入号别' || v_号码 || '的所有序号已被用完';
                    Raise Err_Item;
                  Else
                    Begin
                      Select 1
                      Into n_存在
                      From 挂号序号状态
                      Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期) And 序号 = n_号序;
                    Exception
                      When Others Then
                        Insert Into 挂号序号状态
                          (号码, 日期, 序号, 状态, 操作员姓名, 备注, 登记时间, 机器名)
                        Values
                          (v_号码, Trunc(d_日期), n_号序, 5, v_操作员姓名, '移动锁号', Sysdate, v_机器名);
                    End;
                    v_Temp := '<HX>' || n_号序 || '</HX>';
                    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
                  End If;
                End If;
              Else
                Select Count(1)
                Into n_合约模式
                From 合作单位计划控制
                Where 序号 = 0 And 计划id = n_计划id And 合作单位 = v_合作单位 And 限制项目 = v_星期 And 数量 <> 0;
                If n_合约模式 = 0 Then
                  Begin
                    Select 序号
                    Into n_号序
                    From (Select 序号
                           From 合作单位计划控制 A
                           Where 计划id = n_计划id And 合作单位 = v_合作单位 And 限制项目 = v_星期 And 数量 <> 0 And
                                 (Not Exists
                                  (Select 1
                                   From 挂号序号状态
                                   Where 号码 = v_号码 And 序号 = a.序号 And Trunc(日期) = Trunc(d_日期) And 状态 <> 5) Or Exists
                                  (Select 1
                                   From 挂号序号状态
                                   Where 号码 = v_号码 And 序号 = a.序号 And Trunc(日期) = Trunc(d_日期) And 状态 = 5 And 操作员姓名 = v_操作员姓名 And
                                         机器名 = v_机器名))
                           Order By 序号)
                    Where Rownum < 2;
                  Exception
                    When Others Then
                      n_号序 := 0;
                  End;
                  If n_号序 = 0 Then
                    v_Temp := '传入号别' || v_号码 || '的所有序号已被用完';
                    Raise Err_Item;
                  Else
                    Begin
                      Select 1
                      Into n_存在
                      From 挂号序号状态
                      Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期) And 序号 = n_号序;
                    Exception
                      When Others Then
                        Insert Into 挂号序号状态
                          (号码, 日期, 序号, 状态, 操作员姓名, 备注, 登记时间, 机器名)
                        Values
                          (v_号码, Trunc(d_日期), n_号序, 5, v_操作员姓名, '移动锁号', Sysdate, v_机器名);
                    End;
                    v_Temp := '<HX>' || n_号序 || '</HX>';
                    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
                  End If;
                Else
                  n_号序 := 1;
                  Select 限号数 Into n_限号数 From 挂号计划限制 Where 计划id = n_计划id And 限制项目 = v_星期;
                  For r_序号 In (Select 序号, 状态, 操作员姓名, 机器名
                               From 挂号序号状态
                               Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期)
                               
                               Union
                               Select 序号, Null, Null, Null
                               From 合作单位计划控制
                               Where 计划id = n_计划id And 限制项目 = v_星期 And 数量 <> 0
                               Order By 序号) Loop
                    Exit When r_序号.状态 = 5 And r_序号.操作员姓名 = v_操作员姓名 And r_序号.机器名 = v_机器名;
                    If r_序号.序号 = n_号序 Then
                      n_号序 := n_号序 + 1;
                    End If;
                  End Loop;
                  If n_号序 > n_限号数 Then
                    v_Temp := '传入号别' || v_号码 || '的所有序号已被用完';
                    Raise Err_Item;
                  Else
                    Begin
                      Select 1
                      Into n_存在
                      From 挂号序号状态
                      Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期) And 序号 = n_号序;
                    Exception
                      When Others Then
                        Insert Into 挂号序号状态
                          (号码, 日期, 序号, 状态, 操作员姓名, 备注, 登记时间, 机器名)
                        Values
                          (v_号码, Trunc(d_日期), n_号序, 5, v_操作员姓名, '移动锁号', Sysdate, v_机器名);
                    End;
                    v_Temp := '<HX>' || n_号序 || '</HX>';
                    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
                  End If;
                End If;
              End If;
            End If;
          End If;
        Else
          Begin
            Select 1 Into n_分时段 From 挂号安排时段 Where 星期 = v_星期 And 安排id = n_安排id And Rownum < 2;
          Exception
            When Others Then
              n_分时段 := 0;
          End;
          Begin
            Select 1
            Into n_启用合作单位
            From 合作单位安排控制
            Where 限制项目 = v_星期 And 安排id = n_安排id And 合作单位 = v_合作单位 And Rownum < 2;
          Exception
            When Others Then
              n_启用合作单位 := 0;
          End;
          Begin
            Select 1, a.状态, a.操作员姓名, a.机器名
            Into n_存在, n_状态, v_验证姓名, v_验证机器名
            From 挂号序号状态 A, 挂号安排时段 B
            Where a.号码 = v_号码 And Trunc(a.日期) = Trunc(d_日期) And a.序号 = b.序号 And b.安排id = n_安排id And b.星期 = v_星期 And
                  To_Char(b.开始时间, 'hh24:mi') = To_Char(d_日期, 'hh24:mi') And Rownum < 2;
          Exception
            When Others Then
              n_存在 := 0;
          End;
          If n_存在 = 1 Then
            If n_状态 = 5 And v_验证姓名 = v_操作员姓名 And v_机器名 = v_验证机器名 Then
              Null;
            Else
              --传入时间的序号已经被使用
              v_Temp := '传入时间' || d_日期 || '的序号已被使用';
              Raise Err_Item;
            End If;
          Else
            If n_分时段 = 1 Then
              Begin
                Select 序号
                Into n_号序
                From 挂号安排时段
                Where 安排id = n_安排id And 星期 = v_星期 And To_Char(开始时间, 'hh24:mi') = To_Char(d_日期, 'hh24:mi') And
                      Rownum < 2;
              Exception
                When Others Then
                  Select Max(序号) + 1
                  Into n_号序
                  From (Select Distinct 序号
                         From 挂号安排时段
                         Where 安排id = n_安排id And 星期 = v_星期
                         Union
                         Select Distinct 序号 From 挂号序号状态 Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期));
              End;
              Begin
                Select 1 Into n_存在 From 挂号序号状态 Where 号码 = v_号码 And 日期 = d_日期 And 序号 = n_号序;
              Exception
                When Others Then
                  Insert Into 挂号序号状态
                    (号码, 日期, 序号, 状态, 操作员姓名, 备注, 登记时间, 机器名)
                  Values
                    (v_号码, d_日期, n_号序, 5, v_操作员姓名, '移动锁号', Sysdate, v_机器名);
              End;
              v_Temp := '<HX>' || n_号序 || '</HX>';
              Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
            Else
              If v_合作单位 Is Null Or n_启用合作单位 = 0 Then
                If Trunc(d_日期) = Trunc(Sysdate) Then
                  n_号序 := 1;
                  Select 限号数 Into n_限号数 From 挂号安排限制 Where 安排id = n_安排id And 限制项目 = v_星期;
                  For r_序号 In (Select 序号, 状态, 操作员姓名, 机器名
                               From 挂号序号状态
                               Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期)
                               Order By 序号) Loop
                    Exit When r_序号.状态 = 5 And r_序号.操作员姓名 = v_操作员姓名 And r_序号.机器名 = v_机器名;
                    If r_序号.序号 = n_号序 Then
                      n_号序 := n_号序 + 1;
                    End If;
                  End Loop;
                  If n_号序 > n_限号数 Then
                    v_Temp := '传入号别' || v_号码 || '的所有序号已被用完';
                    Raise Err_Item;
                  Else
                    Begin
                      Select 1
                      Into n_存在
                      From 挂号序号状态
                      Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期) And 序号 = n_号序;
                    Exception
                      When Others Then
                        Insert Into 挂号序号状态
                          (号码, 日期, 序号, 状态, 操作员姓名, 备注, 登记时间, 机器名)
                        Values
                          (v_号码, Trunc(d_日期), n_号序, 5, v_操作员姓名, '移动锁号', Sysdate, v_机器名);
                    End;
                    v_Temp := '<HX>' || n_号序 || '</HX>';
                    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
                  End If;
                Else
                  n_号序 := 1;
                  Select 限号数 Into n_限号数 From 挂号安排限制 Where 安排id = n_安排id And 限制项目 = v_星期;
                  For r_序号 In (Select 序号, 状态, 操作员姓名, 机器名
                               From 挂号序号状态
                               Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期)
                               
                               Union
                               Select 序号, Null, Null, Null
                               From 合作单位安排控制
                               Where 安排id = n_安排id And 限制项目 = v_星期 And 数量 <> 0
                               Order By 序号) Loop
                    Exit When r_序号.状态 = 5 And r_序号.操作员姓名 = v_操作员姓名 And r_序号.机器名 = v_机器名;
                    If r_序号.序号 = n_号序 Then
                      n_号序 := n_号序 + 1;
                    End If;
                  End Loop;
                  If n_号序 > n_限号数 Then
                    v_Temp := '传入号别' || v_号码 || '的所有序号已被用完';
                    Raise Err_Item;
                  Else
                    Begin
                      Select 1
                      Into n_存在
                      From 挂号序号状态
                      Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期) And 序号 = n_号序;
                    Exception
                      When Others Then
                        Insert Into 挂号序号状态
                          (号码, 日期, 序号, 状态, 操作员姓名, 备注, 登记时间, 机器名)
                        Values
                          (v_号码, Trunc(d_日期), n_号序, 5, v_操作员姓名, '移动锁号', Sysdate, v_机器名);
                    End;
                    v_Temp := '<HX>' || n_号序 || '</HX>';
                    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
                  End If;
                End If;
              Else
                Select Count(1)
                Into n_合约模式
                From 合作单位安排控制
                Where 序号 = 0 And 安排id = n_安排id And 合作单位 = v_合作单位 And 限制项目 = v_星期 And 数量 <> 0;
                If n_合约模式 = 0 Then
                  Begin
                    Select 序号
                    Into n_号序
                    From (Select 序号
                           From 合作单位安排控制 A
                           Where 安排id = n_安排id And 合作单位 = v_合作单位 And 限制项目 = v_星期 And 数量 <> 0 And
                                 (Not Exists
                                  (Select 1
                                   From 挂号序号状态
                                   Where 号码 = v_号码 And 序号 = a.序号 And Trunc(日期) = Trunc(d_日期) And 状态 <> 5) Or Exists
                                  (Select 1
                                   From 挂号序号状态
                                   Where 号码 = v_号码 And 序号 = a.序号 And Trunc(日期) = Trunc(d_日期) And 状态 = 5 And 操作员姓名 = v_操作员姓名 And
                                         机器名 = v_机器名))
                           Order By 序号)
                    Where Rownum < 2;
                  Exception
                    When Others Then
                      n_号序 := 0;
                  End;
                  If n_号序 = 0 Then
                    v_Temp := '传入号别' || v_号码 || '的所有序号已被用完';
                    Raise Err_Item;
                  Else
                    Begin
                      Select 1
                      Into n_存在
                      From 挂号序号状态
                      Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期) And 序号 = n_号序;
                    Exception
                      When Others Then
                        Insert Into 挂号序号状态
                          (号码, 日期, 序号, 状态, 操作员姓名, 备注, 登记时间, 机器名)
                        Values
                          (v_号码, Trunc(d_日期), n_号序, 5, v_操作员姓名, '移动锁号', Sysdate, v_机器名);
                    End;
                    v_Temp := '<HX>' || n_号序 || '</HX>';
                    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
                  End If;
                Else
                  n_号序 := 1;
                  Select 限号数 Into n_限号数 From 挂号安排限制 Where 安排id = n_安排id And 限制项目 = v_星期;
                  For r_序号 In (Select 序号, 状态, 操作员姓名, 机器名
                               From 挂号序号状态
                               Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期)
                               
                               Union
                               Select 序号, Null, Null, Null
                               From 合作单位安排控制
                               Where 安排id = n_安排id And 限制项目 = v_星期 And 数量 <> 0
                               Order By 序号) Loop
                    Exit When r_序号.状态 = 5 And r_序号.操作员姓名 = v_操作员姓名 And r_序号.机器名 = v_机器名;
                    If r_序号.序号 = n_号序 Then
                      n_号序 := n_号序 + 1;
                    End If;
                  End Loop;
                  If n_号序 > n_限号数 Then
                    v_Temp := '传入号别' || v_号码 || '的所有序号已被用完';
                    Raise Err_Item;
                  Else
                    Begin
                      Select 1
                      Into n_存在
                      From 挂号序号状态
                      Where 号码 = v_号码 And Trunc(日期) = Trunc(d_日期) And 序号 = n_号序;
                    Exception
                      When Others Then
                        Insert Into 挂号序号状态
                          (号码, 日期, 序号, 状态, 操作员姓名, 备注, 登记时间, 机器名)
                        Values
                          (v_号码, Trunc(d_日期), n_号序, 5, v_操作员姓名, '移动锁号', Sysdate, v_机器名);
                    End;
                    v_Temp := '<HX>' || n_号序 || '</HX>';
                    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
                  End If;
                End If;
              End If;
            End If;
          End If;
        End If;
      End If;
    End If;
  Else
    --出诊表排班模式
    If n_操作类型 = 0 Then
      --解锁
      Begin
        Select 1
        Into n_Exists
        From 临床出诊序号控制
        Where 工作站名称 = v_机器名 And 操作员姓名 = v_操作员姓名 And 挂号状态 = 5 And (序号 = n_号序 Or 备注 = n_号序) And 记录id = n_记录id And
              Rownum < 2;
      Exception
        When Others Then
          n_Exists := 0;
      End;
      If n_Exists = 1 Then
        Update 临床出诊序号控制
        Set 挂号状态 = 0
        Where 工作站名称 = v_机器名 And 操作员姓名 = v_操作员姓名 And 挂号状态 = 5 And (序号 = n_号序 Or 备注 = n_号序) And 记录id = n_记录id;
        v_Temp := '<HX>' || n_号序 || '</HX>';
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
      Else
        v_Temp := '没有发现需要解锁的序号';
        Raise Err_Item;
      End If;
    End If;
  
    If n_操作类型 = 1 Then
      --锁号
      If n_号序 Is Null Then
        Select 是否序号控制, 是否分时段 Into n_序号控制, n_分时段 From 临床出诊记录 Where ID = n_记录id;
        Begin
          Select 1
          Into n_启用合作单位
          From 临床出诊挂号控制记录
          Where 记录id = n_记录id And 名称 = v_合作单位 And 类型 = 1 And 性质 = 1 And Rownum < 2;
        Exception
          When Others Then
            n_启用合作单位 := 0;
        End;
        If n_序号控制 = 1 Then
          Begin
            Select 1, 挂号状态, 操作员姓名, 工作站名称
            Into n_存在, n_状态, v_验证姓名, v_验证机器名
            From 临床出诊序号控制
            Where 记录id = n_记录id And To_Char(开始时间, 'hh24:mi') = To_Char(d_日期, 'hh24:mi') And Nvl(挂号状态, 0) <> 0 And
                  Rownum < 2;
          Exception
            When Others Then
              n_存在 := 0;
          End;
          If n_存在 = 1 Then
            If n_状态 = 5 And v_验证姓名 = v_操作员姓名 And v_机器名 = v_验证机器名 Then
              Null;
            Else
              --传入时间的序号已经被使用
              v_Temp := '传入时间' || d_日期 || '的序号已被使用';
              Raise Err_Item;
            End If;
          Else
            If n_分时段 = 1 Then
              Begin
                Select 序号
                Into n_序号
                From 临床出诊序号控制
                Where 记录id = n_记录id And To_Char(开始时间, 'hh24:mi') = To_Char(d_日期, 'hh24:mi') And Nvl(挂号状态, 0) = 0 And
                      Rownum < 2;
              Exception
                When Others Then
                  Select Max(序号) + 1 Into n_序号 From 临床出诊序号控制 Where 记录id = n_记录id;
              End;
              Update 临床出诊序号控制
              Set 挂号状态 = 5, 锁号时间 = Sysdate, 操作员姓名 = v_操作员姓名, 工作站名称 = v_机器名
              Where 记录id = n_记录id And 序号 = n_序号;
              If Sql%RowCount = 0 Then
                Insert Into 临床出诊序号控制
                  (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 锁号时间, 名称, 类型, 操作员姓名, 工作站名称)
                  Select 记录id, n_序号, To_Date(To_Char(开始时间, 'yyyy-mm-dd') || ' ' || To_Char(d_日期, 'hh24:mi:ss')),
                         To_Date(To_Char(开始时间, 'yyyy-mm-dd') || ' ' || To_Char(d_日期, 'hh24:mi:ss')), 1, 是否预约, 5, Sysdate,
                         v_合作单位, 1, v_操作员姓名, v_机器名
                  From 临床出诊序号控制
                  Where 记录id = n_记录id And Rownum < 2;
              End If;
              v_Temp := '<HX>' || n_序号 || '</HX>';
              Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
            Else
              If v_合作单位 Is Null Or n_启用合作单位 = 0 Then
                --非合作单位
                Select Min(序号) Into n_序号 From 临床出诊序号控制 Where 记录id = n_记录id And Nvl(挂号状态, 0) = 0;
                If n_序号 = 0 Then
                  Select Max(序号) + 1 Into n_序号 From 临床出诊序号控制 Where 记录id = n_记录id;
                End If;
                Update 临床出诊序号控制
                Set 挂号状态 = 5, 锁号时间 = Sysdate, 操作员姓名 = v_操作员姓名, 工作站名称 = v_机器名
                Where 记录id = n_记录id And 序号 = n_序号;
                If Sql%RowCount = 0 Then
                  Insert Into 临床出诊序号控制
                    (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 锁号时间, 名称, 类型, 操作员姓名, 工作站名称)
                    Select 记录id, n_序号, To_Date(To_Char(开始时间, 'yyyy-mm-dd') || ' ' || To_Char(d_日期, 'hh24:mi:ss')),
                           To_Date(To_Char(开始时间, 'yyyy-mm-dd') || ' ' || To_Char(d_日期, 'hh24:mi:ss')), 1, 是否预约, 5,
                           Sysdate, v_合作单位, 1, v_操作员姓名, v_机器名
                    From 临床出诊序号控制
                    Where 记录id = n_记录id And Rownum < 2;
                End If;
                v_Temp := '<HX>' || n_序号 || '</HX>';
                Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
              Else
                --合作单位
                Begin
                  Select 控制方式
                  Into n_合约模式
                  From 临床出诊挂号控制记录
                  Where 记录id = n_记录id And 类型 = 1 And 性质 = 1 And 名称 = v_合作单位 And Rownum < 2;
                Exception
                  When Others Then
                    n_合约模式 := 4;
                End;
                If n_合约模式 = 0 Then
                  v_Temp := '本号别禁止该合作单位预约!';
                  Raise Err_Item;
                End If;
                If n_合约模式 = 1 Or n_合约模式 = 2 Or n_合约模式 = 4 Then
                  Select Min(序号) Into n_序号 From 临床出诊序号控制 Where 记录id = n_记录id And Nvl(挂号状态, 0) = 0;
                  If n_序号 = 0 Then
                    Select Max(序号) + 1 Into n_序号 From 临床出诊序号控制 Where 记录id = n_记录id;
                  End If;
                End If;
                If n_合约模式 = 3 Then
                  Select Min(a.序号)
                  Into n_序号
                  From 临床出诊序号控制 A, 临床出诊挂号控制记录 B
                  Where a.记录id = n_记录id And a.记录id = b.记录id And b.类型 = 1 And b.性质 = 1 And b.名称 = v_合作单位 And a.序号 = b.序号 And
                        Nvl(a.挂号状态, 0) = 0;
                  If n_序号 = 0 Then
                    v_Temp := '本号别合作单位可预约序号已经全部使用完!';
                    Raise Err_Item;
                  End If;
                End If;
                Update 临床出诊序号控制
                Set 挂号状态 = 5, 锁号时间 = Sysdate, 操作员姓名 = v_操作员姓名, 工作站名称 = v_机器名
                Where 记录id = n_记录id And 序号 = n_序号;
                If Sql%RowCount = 0 Then
                  Insert Into 临床出诊序号控制
                    (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 锁号时间, 名称, 类型, 操作员姓名, 工作站名称)
                    Select 记录id, n_序号, To_Date(To_Char(开始时间, 'yyyy-mm-dd') || ' ' || To_Char(d_日期, 'hh24:mi:ss')),
                           To_Date(To_Char(开始时间, 'yyyy-mm-dd') || ' ' || To_Char(d_日期, 'hh24:mi:ss')), 1, 是否预约, 5,
                           Sysdate, v_合作单位, 1, v_操作员姓名, v_机器名
                    From 临床出诊序号控制
                    Where 记录id = n_记录id And Rownum < 2;
                End If;
                v_Temp := '<HX>' || n_序号 || '</HX>';
                Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
              End If;
            End If;
          End If;
        Else
          --非序号控制
          If n_分时段 = 1 Then
            If v_合作单位 Is Null Or n_启用合作单位 = 0 Then
              Begin
                Select 序号, 数量
                Into n_号序, n_数量
                From 临床出诊序号控制
                Where 记录id = n_记录id And 预约顺序号 Is Null And 开始时间 = d_日期;
                Select Count(1)
                Into n_Exists
                From 临床出诊序号控制
                Where 记录id = n_记录id And 预约顺序号 Is Not Null And 序号 = n_号序 And Nvl(挂号状态, 0) <> 0;
                If n_Exists >= n_数量 Then
                  v_Temp := '本号别可用序号已经全部使用完!';
                  Raise Err_Item;
                Else
                  Select Min(预约顺序号)
                  Into n_顺序号
                  From 临床出诊序号控制
                  Where 记录id = n_记录id And 预约顺序号 Is Not Null And 序号 = n_号序 And Nvl(挂号状态, 0) = 0;
                  If n_顺序号 = 0 Then
                    Select Max(预约顺序号) + 1
                    Into n_顺序号
                    From 临床出诊序号控制
                    Where 记录id = n_记录id And 预约顺序号 Is Not Null And 序号 = n_号序;
                  End If;
                  Update 临床出诊序号控制
                  Set 挂号状态 = 5, 锁号时间 = Sysdate, 操作员姓名 = v_操作员姓名, 工作站名称 = v_机器名
                  Where 记录id = n_记录id And 序号 = n_序号 And 预约顺序号 = n_顺序号;
                  If Sql%RowCount = 0 Then
                    Insert Into 临床出诊序号控制
                      (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 锁号时间, 名称, 类型, 操作员姓名, 工作站名称, 预约顺序号)
                      Select 记录id, n_序号, To_Date(To_Char(开始时间, 'yyyy-mm-dd') || ' ' || To_Char(d_日期, 'hh24:mi:ss')),
                             To_Date(To_Char(开始时间, 'yyyy-mm-dd') || ' ' || To_Char(d_日期, 'hh24:mi:ss')), 1, 是否预约, 5,
                             Sysdate, v_合作单位, 1, v_操作员姓名, v_机器名, n_顺序号
                      From 临床出诊序号控制
                      Where 记录id = n_记录id And Rownum < 2;
                  End If;
                  v_Temp := '<HX>' || n_号序 || '</HX>';
                  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
                End If;
              Exception
                When Others Then
                  Null;
              End;
            Else
              --合作单位
              Begin
                Select 控制方式
                Into n_合约模式
                From 临床出诊挂号控制记录
                Where 记录id = n_记录id And 类型 = 1 And 性质 = 1 And 名称 = v_合作单位 And Rownum < 2;
              Exception
                When Others Then
                  n_合约模式 := 4;
              End;
              If n_合约模式 = 0 Then
                v_Temp := '本号别禁止该合作单位预约!';
                Raise Err_Item;
              End If;
              If n_合约模式 = 1 Or n_合约模式 = 2 Or n_合约模式 = 4 Then
                Begin
                  Select 序号, 数量
                  Into n_号序, n_数量
                  From 临床出诊序号控制
                  Where 记录id = n_记录id And 预约顺序号 Is Null And 开始时间 = d_日期;
                  Select Count(1)
                  Into n_Exists
                  From 临床出诊序号控制
                  Where 记录id = n_记录id And 预约顺序号 Is Not Null And 序号 = n_号序 And Nvl(挂号状态, 0) <> 0;
                  If n_Exists >= n_数量 Then
                    v_Temp := '本号别可用序号已经全部使用完!';
                    Raise Err_Item;
                  Else
                    Select Min(预约顺序号)
                    Into n_顺序号
                    From 临床出诊序号控制
                    Where 记录id = n_记录id And 预约顺序号 Is Not Null And 序号 = n_号序 And Nvl(挂号状态, 0) = 0;
                    If n_顺序号 = 0 Then
                      Select Max(预约顺序号) + 1
                      Into n_顺序号
                      From 临床出诊序号控制
                      Where 记录id = n_记录id And 预约顺序号 Is Not Null And 序号 = n_号序;
                    End If;
                    Update 临床出诊序号控制
                    Set 挂号状态 = 5, 锁号时间 = Sysdate, 操作员姓名 = v_操作员姓名, 工作站名称 = v_机器名
                    Where 记录id = n_记录id And 序号 = n_序号 And 预约顺序号 = n_顺序号;
                    If Sql%RowCount = 0 Then
                      Insert Into 临床出诊序号控制
                        (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 锁号时间, 名称, 类型, 操作员姓名, 工作站名称, 预约顺序号)
                        Select 记录id, n_序号, To_Date(To_Char(开始时间, 'yyyy-mm-dd') || ' ' || To_Char(d_日期, 'hh24:mi:ss')),
                               To_Date(To_Char(开始时间, 'yyyy-mm-dd') || ' ' || To_Char(d_日期, 'hh24:mi:ss')), 1, 是否预约, 5,
                               Sysdate, v_合作单位, 1, v_操作员姓名, v_机器名, n_顺序号
                        From 临床出诊序号控制
                        Where 记录id = n_记录id And Rownum < 2;
                    End If;
                    v_Temp := '<HX>' || n_号序 || '</HX>';
                    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
                  End If;
                Exception
                  When Others Then
                    Null;
                End;
              End If;
              If n_合约模式 = 3 Then
                Begin
                  Select 序号
                  Into n_号序
                  From 临床出诊序号控制
                  Where 记录id = n_记录id And 预约顺序号 Is Null And 开始时间 = d_日期;
                  Select 数量
                  Into n_数量
                  From 临床出诊挂号控制记录
                  Where 记录id = n_记录id And 类型 = 1 And 性质 = 1 And 名称 = v_合作单位 And 序号 = n_号序;
                  Select Count(1)
                  Into n_Exists
                  From 临床出诊序号控制
                  Where 记录id = n_记录id And 预约顺序号 Is Not Null And 序号 = n_号序 And Nvl(挂号状态, 0) <> 0;
                  If n_Exists >= n_数量 Then
                    v_Temp := '本号别可用序号已经全部使用完!';
                    Raise Err_Item;
                  Else
                    Select Min(预约顺序号)
                    Into n_顺序号
                    From 临床出诊序号控制
                    Where 记录id = n_记录id And 预约顺序号 Is Not Null And 序号 = n_号序 And Nvl(挂号状态, 0) = 0;
                    If n_顺序号 = 0 Then
                      Select Max(预约顺序号) + 1
                      Into n_顺序号
                      From 临床出诊序号控制
                      Where 记录id = n_记录id And 预约顺序号 Is Not Null And 序号 = n_号序;
                    End If;
                    Update 临床出诊序号控制
                    Set 挂号状态 = 5, 锁号时间 = Sysdate, 操作员姓名 = v_操作员姓名, 工作站名称 = v_机器名
                    Where 记录id = n_记录id And 序号 = n_序号 And 预约顺序号 = n_顺序号;
                    If Sql%RowCount = 0 Then
                      Insert Into 临床出诊序号控制
                        (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 锁号时间, 名称, 类型, 操作员姓名, 工作站名称, 预约顺序号)
                        Select 记录id, n_序号, To_Date(To_Char(开始时间, 'yyyy-mm-dd') || ' ' || To_Char(d_日期, 'hh24:mi:ss')),
                               To_Date(To_Char(开始时间, 'yyyy-mm-dd') || ' ' || To_Char(d_日期, 'hh24:mi:ss')), 1, 是否预约, 5,
                               Sysdate, v_合作单位, 1, v_操作员姓名, v_机器名, n_顺序号
                        From 临床出诊序号控制
                        Where 记录id = n_记录id And Rownum < 2;
                    End If;
                    v_Temp := '<HX>' || n_号序 || '</HX>';
                    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
                  End If;
                Exception
                  When Others Then
                    Null;
                End;
              End If;
            End If;
          End If;
        End If;
      Else
        n_序号 := n_号序;
        Update 临床出诊序号控制
        Set 挂号状态 = 5, 操作员姓名 = v_操作员姓名, 工作站名称 = v_机器名, 锁号时间 = Sysdate
        Where 记录id = n_记录id And 序号 = n_序号;
        If Sql%RowCount = 0 Then
          Begin
            Insert Into 临床出诊序号控制
              (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 锁号时间, 名称, 类型, 操作员姓名, 工作站名称)
              Select 记录id, n_序号, To_Date(To_Char(开始时间, 'yyyy-mm-dd') || ' ' || To_Char(d_日期, 'hh24:mi:ss')),
                     To_Date(To_Char(开始时间, 'yyyy-mm-dd') || ' ' || To_Char(d_日期, 'hh24:mi:ss')), 1, 是否预约, 5, Sysdate,
                     v_合作单位, 1, v_操作员姓名, v_机器名
              From 临床出诊序号控制
              Where 记录id = n_记录id And Rownum < 2;
          Exception
            When Others Then
              v_Temp := '传入的锁号序号已被使用!';
              Raise Err_Item;
          End;
        End If;
      End If;
    End If;
  End If;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Temp || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Lockno;
/

Create Or Replace Procedure Zl_Third_Getnolist
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:获取号源列表
  --入参:Xml_In:
  --<IN>
  --  <RQ>日期</RQ>
  --  <KSID>科室ID</KSID>
  --  <YSID>医生ID</YSID>
  --  <YSXM>医生姓名</YSXM>
  --  <HZDW>支付宝</HZDW>    //合作单位，传入了的时候，只取合作单位的号;为空时，只取非合作单位的号
  --</IN>

  --出参:Xml_Out
  --<OUTPUT>
  --  <GROUP>
  --    <RQ>日期</RQ>
  --    <HBLIST>
  --     <HB>
  --        <CZJLID>1</CZJLID>     //出诊记录ID
  --        <HM>235</HM>       //号码
  --        <YSID>549</YSID>      //医生ID
  --        <YS>张锐</YS>       //医生姓名
  --        <KSID>123</KSID>   //科室ID
  --        <KSMC>内科</KSMC>   //科室名称
  --        <ZC>主治医师</ZC> //职称
  --        <XMID>10086<XMID> //挂号项目的ID
  --        <XMMC>挂号费</XMMC> //挂号项目的名称
  --        <YGHS>0</YGHS>      //已挂号数
  --        <SYHS>99</SYHS>   //剩余号数
  --        <PRICE>15</PRICE>      //价格
  --        <HL>普通</HL>       //挂号类型
  --        <HCXH>1</HCXH>    //是否存在缓冲序号时间段，1-存在 0或者空-不存在
  --        <FSD>0</FSD>      //是否分时段
  --        <FWMC>白天</FWMC>     //号别时段
  --        <HBTIME>(08:00-17:59)</HBTIME> //可挂时间
  --     <SPANLIST>
  --            <SPAN>
  --                  <SJD/>      //时间段
  --                  <SL/>      //数量
  --            </SPAN>
  --            ……
  --          </SPANLIST>
  --      </HB>
  --      <HB>
  --      ……
  --      </HB>
  --    </HBLIST>
  --  </GROUP>
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  d_日期         Date;
  n_科室id       病人挂号记录.执行部门id%Type;
  n_医生id       人员表.Id%Type;
  v_医生姓名     人员表.姓名%Type;
  v_星期         挂号安排限制.限制项目%Type;
  v_时间段       Varchar2(100);
  v_合作单位     挂号合作单位.名称%Type;
  n_分时段       Number(3);
  n_单个剩余     Number(5);
  n_已挂数       Number(5);
  n_合约已挂数   Number(5);
  n_合计金额     收费价目.现价%Type;
  n_合约总数量   Number(5);
  n_合约剩余数量 Number(5);
  n_最大可用数量 Number(5);
  n_合约模式     Number(3); --合约模式:1-合约单位限数量模式 0-合约单位指定序号模式
  n_非合约       Number(3);
  n_是否预留     Number(3);
  d_加号时间     Date;
  d_开始时间     临床出诊记录.开始时间%Type;
  d_终止时间     临床出诊记录.终止时间%Type;
  n_缓冲序号     Number(3);
  n_时段数量     Number(5);
  n_序号控制     临床出诊记录.是否序号控制%Type;
  n_预留数量     Number(5);
  n_特殊预约     Number(3);
  n_禁用         Number(3);
  v_剩余数量     Varchar2(100);
  v_Timetemp     Varchar2(100);
  v_Temp         Varchar2(32767); --临时XML
  v_Xmlmain      Clob; --临时XML
  c_Xmlmain      Clob; --临时XML
  x_Templet      Xmltype; --模板XML
  v_Err_Msg      Varchar2(200);
  n_Exists       Number(2);
  v_Sql          Varchar2(20000);
  Type c_Main Is Ref Cursor;
  r_科室id   挂号安排.科室id%Type;
  r_号类     挂号安排.号类%Type;
  r_科室名称 部门表.名称%Type;
  r_医生姓名 挂号安排.医生姓名%Type;
  r_医生id   挂号安排.医生id%Type;
  r_职称     人员表.专业技术职务%Type;
  r_号码     挂号安排.号码%Type;
  r_安排id   挂号安排.Id%Type;
  r_计划id   挂号安排计划.Id%Type;
  r_排班     挂号安排.周日%Type;
  r_项目id   挂号安排.项目id%Type;
  r_项目名称 收费项目目录.名称%Type;
  r_序号控制 挂号安排.序号控制%Type;
  r_限号数   挂号安排限制.限号数%Type;
  r_限约数   挂号安排限制.限约数%Type;
  n_时段已挂 Number(5);
  r_已挂数   病人挂号汇总.已挂数%Type;
  r_已约数   病人挂号汇总.已约数%Type;
  r_已接收   病人挂号汇总.其中已接收%Type;
  r_价格     收费价目.现价%Type;
  r_分时段   临床出诊记录.是否分时段%Type;
  r_开始时间 临床出诊记录.开始时间%Type;
  r_终止时间 临床出诊记录.终止时间%Type;
  r_预约控制 临床出诊记录.预约控制%Type;
  r_No       c_Main;
  n_Curcount Number(3);
  n_挂号模式 Number(3);
  v_挂号模式 Varchar2(500);
  v_启用时间 Varchar2(500);

  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select To_Date(Extractvalue(Value(A), 'IN/RQ'), 'YYYY-MM-DD hh24:mi:ss'), Extractvalue(Value(A), 'IN/KSID'),
         Extractvalue(Value(A), 'IN/YSID'), Extractvalue(Value(A), 'IN/YSXM'), Extractvalue(Value(A), 'IN/HZDW')
  Into d_日期, n_科室id, n_医生id, v_医生姓名, v_合作单位
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  v_挂号模式 := zl_GetSysParameter('挂号排班模式');
  n_挂号模式 := To_Number(Substr(v_挂号模式, 1, 1));
  If n_挂号模式 = 1 Then
    Begin
      v_启用时间 := Substr(v_挂号模式, 3);
    Exception
      When Others Then
        Null;
    End;
  End If;
  --日期节点为空的情况
  If d_日期 Is Null Then
    d_日期 := Trunc(Sysdate);
  End If;

  If n_挂号模式 = 0 Then
    Select Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null)
    Into v_星期
    From Dual;
    n_合约剩余数量 := 0;
  
    v_Sql := 'Select a.*, Nvl(Hz.已挂数, 0) As 已挂数, Nvl(Hz.已约数, 0) As 已约数, Nvl(Hz.其中已接收, 0) As 已接收, b.现价 As 价格 ';
    v_Sql := v_Sql ||
             'From (Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Ap.医生姓名, Nvl(Ap.医生id, 0) As 医生id, Ry.专业技术职务 As 职称, Ap.号码, ';
    v_Sql := v_Sql || ' Ap.安排id, Ap.计划id, Ap.排班, Ap.项目id, Fy.名称 As 项目名称, Ap.序号控制, Ap.限号数, Ap.限约数 ';
    v_Sql := v_Sql || 'From (Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Ap.医生姓名, Ap.医生id, Ap.号码, Ap.Id As 安排id, 0 As 计划id, ';
    v_Sql := v_Sql || 'Ap.项目id, Nvl(Ap.序号控制, 0) As 序号控制, ';
    v_Sql := v_Sql ||
             'Decode(To_Char(:1, ''D''), ''1'', Ap.周日, ''2'', Ap.周一, ''3'', Ap.周二, ''4'', Ap.周三, ''5'', Ap.周四, ';
    v_Sql := v_Sql || ' ''6'', Ap.周五, ''7'', Ap.周六, Null) As 排班, Xz.限约数, Xz.限号数 ';
    v_Sql := v_Sql || 'From 挂号安排 Ap, 部门表 Bm, 挂号安排限制 Xz ';
    v_Sql := v_Sql || 'Where Ap.科室id = Bm.Id(+) ';
  
    n_Curcount := 2;
    If Nvl(n_科室id, 0) <> 0 Then
      v_Sql      := v_Sql || 'And Ap.科室id = :2 ';
      n_Curcount := n_Curcount + 1;
    End If;
    If Nvl(n_医生id, 0) <> 0 Then
      If n_Curcount = 2 Then
        v_Sql := v_Sql || 'And Ap.医生id = :2 ';
      Else
        v_Sql := v_Sql || 'And Ap.医生id = :3 ';
      End If;
      n_Curcount := n_Curcount + 1;
    End If;
    If Nvl(v_医生姓名, '_') <> '_' Then
      If n_Curcount = 2 Then
        v_Sql := v_Sql || 'And Ap.医生姓名 = :2 ';
      End If;
      If n_Curcount = 3 Then
        v_Sql := v_Sql || 'And Ap.医生姓名 = :3 ';
      End If;
      If n_Curcount = 4 Then
        v_Sql := v_Sql || 'And Ap.医生姓名 = :4 ';
      End If;
      n_Curcount := n_Curcount + 1;
    End If;
  
    v_Sql      := v_Sql || 'And Ap.停用日期 Is Null And :' || n_Curcount ||
                  ' Between Nvl(Ap.开始时间, To_Date(''1900-01-01'', ''YYYY-MM-DD'')) And ';
    n_Curcount := n_Curcount + 1;
    v_Sql      := v_Sql || 'Nvl(Ap.终止时间, To_Date(''3000 - 01 - 01'', ''YYYY-MM-DD'')) And Xz.安排id(+) = Ap.Id And ';
    v_Sql      := v_Sql || ' Xz.限制项目(+) = Decode(To_Char(:' || n_Curcount ||
                  ', ''D''), ''1'', ''周日'', ''2'', ''周一'', ''3'', ''周二'', ''4'', ''周三'', ''5'', ';
    n_Curcount := n_Curcount + 1;
    v_Sql      := v_Sql || ' ''周四'', ''6'', ''周五'', ''7'', ''周六'', Null) And Not Exists ';
    v_Sql      := v_Sql || '(Select Rownum ';
    v_Sql      := v_Sql || 'From 挂号安排停用状态 Ty ';
    v_Sql      := v_Sql || 'Where Ty.安排id = Ap.Id And :' || n_Curcount ||
                  ' Between Ty.开始停止时间 And Ty.结束停止时间) And Not Exists ';
    n_Curcount := n_Curcount + 1;
    v_Sql      := v_Sql || '(Select Rownum ';
    v_Sql      := v_Sql || 'From 挂号安排计划 Jh Where Jh.安排id = Ap.Id And Jh.审核时间 Is Not Null And ';
    v_Sql      := v_Sql || ':' || n_Curcount ||
                  ' Between Nvl(Jh.生效时间, To_Date(''1900-01-01'', ''YYYY-MM-DD'')) And Nvl(Jh.失效时间, To_Date(''3000-01-01'', ''YYYY-MM-DD''))) ';
    n_Curcount := n_Curcount + 1;
    v_Sql      := v_Sql || 'Union All ';
    v_Sql      := v_Sql ||
                  'Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Jh.医生姓名, Jh.医生id, Ap.号码, Ap.Id As 安排id, Jh.Id As 计划id, ';
    v_Sql      := v_Sql || 'Jh.项目id, Nvl(Jh.序号控制, 0) As 序号控制,Decode(To_Char(:' || n_Curcount ||
                  ', ''D''), ''1'', Jh.周日, ''2'', Jh.周一, ''3'', Jh.周二, ''4'', Jh.周三, ''5'', Jh.周四, ';
    n_Curcount := n_Curcount + 1;
    v_Sql      := v_Sql || ' ''6'', Jh.周五, ''7'', Jh.周六, Null) As 排班, Xz.限约数, Xz.限号数 ';
    v_Sql      := v_Sql || 'From 挂号安排计划 Jh, 挂号安排 Ap, 部门表 Bm, 挂号计划限制 Xz ';
    v_Sql      := v_Sql || 'Where Jh.安排id = Ap.Id And Ap.科室id = Bm.Id(+) And Ap.停用日期 Is Null ';
  
    If Nvl(n_科室id, 0) <> 0 Then
      v_Sql      := v_Sql || 'And Ap.科室id = :' || n_Curcount || ' ';
      n_Curcount := n_Curcount + 1;
    End If;
    If Nvl(n_医生id, 0) <> 0 Then
      v_Sql      := v_Sql || 'And Ap.医生id = :' || n_Curcount || ' ';
      n_Curcount := n_Curcount + 1;
    End If;
    If Nvl(v_医生姓名, '_') <> '_' Then
      v_Sql      := v_Sql || 'And Ap.医生姓名 = :' || n_Curcount || ' ';
      n_Curcount := n_Curcount + 1;
    End If;
  
    v_Sql      := v_Sql || ' And :' || n_Curcount ||
                  ' Between Nvl(Jh.生效时间, To_Date(''1900-01-01'', ''YYYY-MM-DD'')) And Nvl(Jh.失效时间, To_Date(''3000-01-01'', ''YYYY-MM-DD'')) And Xz.计划id(+) = Jh.Id And ';
    n_Curcount := n_Curcount + 1;
    v_Sql      := v_Sql || 'Xz.限制项目(+) = Decode(To_Char(:' || n_Curcount ||
                  ', ''D''), ''1'', ''周日'', ''2'', ''周一'', ''3'', ''周二'', ''4'', ''周三'', ''5'', ''周四'', ''6'', ''周五'', ''7'', ''周六'', Null) And Not Exists ';
    n_Curcount := n_Curcount + 1;
    v_Sql      := v_Sql || '(Select Rownum From 挂号安排停用状态 Ty Where Ty.安排id = Ap.Id And :' || n_Curcount ||
                  ' Between Ty.开始停止时间 And Ty.结束停止时间) And ';
    n_Curcount := n_Curcount + 1;
    v_Sql      := v_Sql || '(Jh.生效时间, Jh.安排id) = (Select Max(Sxjh.生效时间) As 生效时间, 安排id From 挂号安排计划 Sxjh ';
    v_Sql      := v_Sql || ' Where Sxjh.审核时间 Is Not Null And :' || n_Curcount ||
                  ' Between Nvl(Sxjh.生效时间, To_Date(''1900-01-01'', ''YYYY-MM-DD'')) And ';
    n_Curcount := n_Curcount + 1;
    v_Sql      := v_Sql || 'Nvl(Sxjh.失效时间, To_Date(''3000-01-01'', ''YYYY-MM-DD'')) And Sxjh.安排id = Jh.安排id ';
    v_Sql      := v_Sql || 'Group By Sxjh.安排id)) Ap, 部门表 Bm, 人员表 Ry, 收费项目目录 Fy ';
    v_Sql      := v_Sql ||
                  'Where Ap.科室id = Bm.Id(+) And Ap.医生id = Ry.Id(+) And Ap.项目id = Fy.Id And 排班 Is Not Null) A, ';
    v_Sql      := v_Sql || '病人挂号汇总 Hz, 收费价目 B ';
    v_Sql      := v_Sql || 'Where a.号码 = Hz.号码(+) And Hz.日期(+) = Trunc(:' || n_Curcount ||
                  ') And a.项目id = b.收费细目id And ';
    n_Curcount := n_Curcount + 1;
    v_Sql      := v_Sql || 'Nvl(b.终止日期, To_Date(''3000-1-1'', ''YYYY-Mm-DD'')) > :' || n_Curcount || ' ';
    n_Curcount := n_Curcount + 1;
    v_Sql      := v_Sql || 'And b.执行日期 <= :' || n_Curcount || ' ';
    If Nvl(n_科室id, 0) <> 0 And Nvl(n_医生id, 0) <> 0 And Nvl(v_医生姓名, '_') <> '_' Then
      Open r_No For v_Sql
        Using d_日期, n_科室id, n_医生id, v_医生姓名, d_日期, d_日期, d_日期, d_日期, d_日期, n_科室id, n_医生id, v_医生姓名, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期;
    End If;
    If Nvl(n_科室id, 0) <> 0 And Nvl(n_医生id, 0) = 0 And Nvl(v_医生姓名, '_') = '_' Then
      Open r_No For v_Sql
        Using d_日期, n_科室id, d_日期, d_日期, d_日期, d_日期, d_日期, n_科室id, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期;
    End If;
    If Nvl(n_科室id, 0) = 0 And Nvl(n_医生id, 0) <> 0 And Nvl(v_医生姓名, '_') = '_' Then
      Open r_No For v_Sql
        Using d_日期, n_医生id, d_日期, d_日期, d_日期, d_日期, d_日期, n_医生id, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期;
    End If;
    If Nvl(n_科室id, 0) = 0 And Nvl(n_医生id, 0) = 0 And Nvl(v_医生姓名, '_') <> '_' Then
      Open r_No For v_Sql
        Using d_日期, v_医生姓名, d_日期, d_日期, d_日期, d_日期, d_日期, v_医生姓名, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期;
    End If;
    If Nvl(n_科室id, 0) <> 0 And Nvl(n_医生id, 0) <> 0 And Nvl(v_医生姓名, '_') = '_' Then
      Open r_No For v_Sql
        Using d_日期, n_科室id, n_医生id, d_日期, d_日期, d_日期, d_日期, d_日期, n_科室id, n_医生id, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期;
    End If;
    If Nvl(n_科室id, 0) <> 0 And Nvl(n_医生id, 0) = 0 And Nvl(v_医生姓名, '_') <> '_' Then
      Open r_No For v_Sql
        Using d_日期, n_科室id, v_医生姓名, d_日期, d_日期, d_日期, d_日期, d_日期, n_科室id, v_医生姓名, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期;
    End If;
    If Nvl(n_科室id, 0) = 0 And Nvl(n_医生id, 0) <> 0 And Nvl(v_医生姓名, '_') <> '_' Then
      Open r_No For v_Sql
        Using d_日期, n_医生id, v_医生姓名, d_日期, d_日期, d_日期, d_日期, d_日期, n_医生id, v_医生姓名, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期;
    End If;
    Loop
      Fetch r_No
        Into r_科室id, r_号类, r_科室名称, r_医生姓名, r_医生id, r_职称, r_号码, r_安排id, r_计划id, r_排班, r_项目id, r_项目名称, r_序号控制, r_限号数,
             r_限约数, r_已挂数, r_已约数, r_已接收, r_价格;
      Exit When r_No%NotFound;
      If r_计划id <> 0 Then
        Select Sign(Count(Rownum))
        Into n_分时段
        From 挂号安排计划 Jh, 挂号计划时段 Sd
        Where Jh.Id = Sd.计划id And Jh.Id = r_计划id And
              Sd.星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7',
                             '周六', Null) And Rownum < 2;
      Else
        Select Sign(Count(Rownum))
        Into n_分时段
        From 挂号安排 Ap, 挂号安排时段 Sd
        Where Ap.Id = Sd.安排id And Ap.Id = r_安排id And
              Sd.星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7',
                             '周六', Null) And Rownum < 2;
      End If;
      If n_分时段 = 0 Then
        v_Temp := '';
        If v_合作单位 Is Not Null And r_序号控制 = 1 Then
          If r_计划id <> 0 Then
            Select Nvl(Sum(数量), 0)
            Into n_合约总数量
            From 合作单位计划控制
            Where 计划id = r_计划id And 合作单位 = v_合作单位 And
                  限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                                '7', '周六', Null);
            Select Count(1)
            Into n_合约模式
            From 合作单位计划控制
            Where 计划id = r_计划id And 合作单位 = v_合作单位 And
                  限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                                '7', '周六', Null) And 序号 = 0;
          Else
            Select Nvl(Sum(数量), 0)
            Into n_合约总数量
            From 合作单位安排控制
            Where 安排id = r_安排id And 合作单位 = v_合作单位 And
                  限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                                '7', '周六', Null);
            Select Count(1)
            Into n_合约模式
            From 合作单位安排控制
            Where 安排id = r_安排id And 合作单位 = v_合作单位 And
                  限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                                '7', '周六', Null) And 序号 = 0;
          End If;
          If n_合约模式 = 0 Then
            If r_计划id <> 0 Then
              Select Count(1)
              Into n_合约已挂数
              From 病人挂号记录 A
              Where 号别 = r_号码 And 记录状态 = 1 And 发生时间 Between Trunc(d_日期) And Trunc(d_日期 + 1) - 1 / 60 / 60 / 24 And
                    Exists (Select 1
                     From 合作单位计划控制
                     Where 计划id = r_计划id And 合作单位 = v_合作单位 And
                           限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                         '周四', '6', '周五', '7', '周六', Null) And 序号 = a.号序 And 数量 <> 0);
            Else
              Select Count(1)
              Into n_合约已挂数
              From 病人挂号记录 A
              Where 号别 = r_号码 And 记录状态 = 1 And 发生时间 Between Trunc(d_日期) And Trunc(d_日期 + 1) - 1 / 60 / 60 / 24 And
                    Exists (Select 1
                     From 合作单位安排控制
                     Where 安排id = r_安排id And 合作单位 = v_合作单位 And
                           限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                         '周四', '6', '周五', '7', '周六', Null) And 序号 = a.号序 And 数量 <> 0);
            End If;
          Else
            Begin
              Select Count(1)
              Into n_合约已挂数
              From 病人挂号记录
              Where 号别 = r_号码 And 记录状态 = 1 And 合作单位 = v_合作单位 And 发生时间 Between Trunc(d_日期) And
                    Trunc(d_日期 + 1) - 1 / 60 / 60 / 24;
            Exception
              When Others Then
                n_合约已挂数 := 0;
            End;
          End If;
          If n_合约总数量 = 0 Then
            n_合约剩余数量 := 0;
          Else
            n_合约剩余数量 := n_合约总数量 - n_合约已挂数;
            If n_合约剩余数量 > (Nvl(r_限号数, 0) - r_已挂数) Then
              n_合约剩余数量 := Nvl(r_限号数, 0) - r_已挂数;
            End If;
          End If;
        End If;
      Else
        v_Temp := '<SPANLIST>';
        If r_计划id <> 0 Then
          Select Max(结束时间)
          Into d_加号时间
          From 挂号计划时段
          Where 计划id = r_计划id And 星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四',
                                              '6', '周五', '7', '周六', Null);
          If r_序号控制 = 1 Then
            If Trunc(d_日期) = Trunc(Sysdate) Then
              n_特殊预约 := 0;
            Else
              Select Nvl(Max(Jh.是否预约), 0)
              Into n_特殊预约
              From (Select Sd.计划id, Sd.序号, Sd.星期, Jh.号码,
                            To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'),
                                     'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                            To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'),
                                     'yyyy-mm-dd hh24:mi:ss') As 结束时间, Sd.限制数量, Sd.是否预约
                     From 挂号安排计划 Jh, 挂号计划时段 Sd
                     Where Jh.Id = Sd.计划id And Jh.Id = r_计划id And
                           Sd.星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6',
                                          '周五', '7', '周六', Null)) Jh, 挂号序号状态 Zt
              Where Zt.日期(+) = Jh.开始时间 And Zt.号码(+) = Jh.号码 And Zt.序号(+) = Jh.序号 And
                    Decode(Sign(Sysdate - Jh.开始时间), -1, 0, 1) <> 1;
            End If;
          
            For r_Time In (Select Jh.号码, Jh.序号, Jh.星期, Jh.开始时间, Jh.结束时间, Jh.限制数量, Jh.是否预约, 0 As 已约数,
                                  Decode(Nvl(Zt.序号, 0), 0, 1, 0) As 剩余数,
                                  Decode(Sign(Sysdate - Jh.开始时间), -1, 0, 1) As 失效时段
                           
                           From (Select Sd.计划id, Sd.序号, Sd.星期, Jh.号码,
                                         To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'),
                                                  'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                                         To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'),
                                                  'yyyy-mm-dd hh24:mi:ss') As 结束时间, Sd.限制数量, Sd.是否预约
                                  From 挂号安排计划 Jh, 挂号计划时段 Sd
                                  Where Jh.Id = Sd.计划id And Jh.Id = r_计划id And
                                        Sd.星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                                       '周四', '6', '周五', '7', '周六', Null)) Jh, 挂号序号状态 Zt
                           Where Zt.日期(+) = Jh.开始时间 And Zt.号码(+) = Jh.号码 And Zt.序号(+) = Jh.序号 And
                                 Decode(Sign(Sysdate - Jh.开始时间), -1, 0, 1) <> 1
                           Order By 序号) Loop
              If v_合作单位 Is Not Null Then
                Begin
                  Select 1
                  Into n_合约模式
                  From 合作单位计划控制
                  Where 限制项目 = r_Time.星期 And 计划id = r_计划id And 序号 = 0 And 合作单位 = v_合作单位;
                Exception
                  When Others Then
                    n_合约模式 := 0;
                End;
              Else
                n_合约模式 := 0;
              End If;
              If r_Time.剩余数 = 0 Then
                n_单个剩余 := 0;
              Else
                n_单个剩余 := r_Time.限制数量;
              End If;
              If v_合作单位 Is Null Or n_合约模式 = 1 Then
                Begin
                  Select 1
                  Into n_Exists
                  From 合作单位计划控制
                  Where 限制项目 = r_Time.星期 And 计划id = r_计划id And 序号 = r_Time.序号 And Rownum < 2;
                Exception
                  When Others Then
                    n_Exists := 0;
                End;
                If n_Exists = 0 Then
                  If n_特殊预约 = 1 And r_Time.是否预约 = 0 Then
                    Null;
                  Else
                    Begin
                      Select 1
                      Into n_是否预留
                      From 挂号序号状态
                      Where 状态 In (3, 4) And 号码 = r_号码 And 序号 = r_Time.序号 And Trunc(日期) = Trunc(d_日期);
                    Exception
                      When Others Then
                        n_是否预留 := 0;
                    End;
                    If n_是否预留 = 0 Then
                      v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                    To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                    '</SPAN>';
                      n_时段数量 := Nvl(n_时段数量, 0) + n_单个剩余;
                    End If;
                  End If;
                End If;
              Else
                Begin
                  Select 1
                  Into n_Exists
                  From 合作单位计划控制
                  Where 限制项目 = r_Time.星期 And 计划id = r_计划id And 序号 = r_Time.序号 And 合作单位 = v_合作单位 And Rownum < 2;
                Exception
                  When Others Then
                    n_Exists := 0;
                End;
                Begin
                  Select 0
                  Into n_非合约
                  From 合作单位计划控制
                  Where 计划id = r_计划id And 合作单位 = v_合作单位 And Rownum < 2;
                Exception
                  When Others Then
                    n_非合约 := 1;
                End;
                If n_Exists = 1 Or n_非合约 = 1 Then
                  If n_特殊预约 = 1 And r_Time.是否预约 = 0 Then
                    Null;
                  Else
                    Begin
                      Select 1
                      Into n_是否预留
                      From 挂号序号状态
                      Where 状态 In (3, 4) And 号码 = r_号码 And 序号 = r_Time.序号 And Trunc(日期) = Trunc(d_日期);
                    Exception
                      When Others Then
                        n_是否预留 := 0;
                    End;
                    If n_是否预留 = 0 Then
                      v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                        To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                        '</SPAN>';
                      n_合约剩余数量 := n_合约剩余数量 + 1;
                    End If;
                  End If;
                End If;
              End If;
            End Loop;
          Else
            n_最大可用数量 := Nvl(r_限约数, Nvl(r_限号数, 0)) - Nvl(r_已约数, 0);
            For r_Time In (Select Jh.号码, Jh.序号, Jh.星期, Jh.开始时间, Jh.结束时间, Jh.限制数量, Jh.是否预约,
                                  Sum(Decode(Nvl(Zt.序号, 0), 0, 0, 1)) As 已约数,
                                  Jh.限制数量 - Sum(Decode(Nvl(Zt.序号, 0), 0, 0, 1)) As 剩余数,
                                  Decode(Sign(Sysdate - Jh.开始时间), -1, 0, 1) As 失效时段
                           From (Select Sd.计划id, Sd.序号, Sd.星期, Jh.号码,
                                         To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'),
                                                  'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                                         To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'),
                                                  'yyyy-mm-dd hh24:mi:ss') As 结束时间, Sd.限制数量, Sd.是否预约
                                  From 挂号安排计划 Jh, 挂号计划时段 Sd
                                  Where Jh.Id = Sd.计划id And Jh.Id = r_计划id And
                                        Sd.星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                                       '周四', '6', '周五', '7', '周六', Null)) Jh, 挂号序号状态 Zt
                           Where Jh.号码 = Zt.号码(+) And Jh.开始时间 = Zt.日期(+) And
                                 Decode(Sign(Sysdate - Jh.开始时间), -1, 0, 1) <> 1
                           Group By Jh.号码, Jh.序号, Jh.星期, Jh.开始时间, Jh.结束时间, Jh.限制数量, Jh.是否预约
                           Order By Jh.序号) Loop
              If v_合作单位 Is Not Null Then
                Begin
                  Select 1
                  Into n_合约模式
                  From 合作单位计划控制
                  Where 限制项目 = r_Time.星期 And 计划id = r_计划id And 序号 = 0 And 合作单位 = v_合作单位;
                Exception
                  When Others Then
                    n_合约模式 := 0;
                End;
              Else
                n_合约模式 := 0;
              End If;
              n_单个剩余 := r_Time.剩余数;
              If v_合作单位 Is Null Or n_合约模式 = 1 Then
                Begin
                  Select 1
                  Into n_Exists
                  From 合作单位计划控制
                  Where 限制项目 = r_Time.星期 And 计划id = r_计划id And 序号 = r_Time.序号 And Rownum < 2;
                Exception
                  When Others Then
                    n_Exists := 0;
                End;
                If n_Exists = 0 Then
                  If n_最大可用数量 < n_单个剩余 Then
                    v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_最大可用数量 || '</SL>' ||
                                  '</SPAN>';
                    n_时段数量 := Nvl(n_时段数量, 0) + n_最大可用数量;
                  Else
                    v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                  '</SPAN>';
                    n_时段数量 := Nvl(n_时段数量, 0) + n_单个剩余;
                  End If;
                End If;
              Else
                Begin
                  Select 1
                  Into n_Exists
                  From 合作单位计划控制
                  Where 限制项目 = r_Time.星期 And 计划id = r_计划id And 序号 = r_Time.序号 And 合作单位 = v_合作单位 And Rownum < 2;
                Exception
                  When Others Then
                    n_Exists := 0;
                End;
                Begin
                  Select 0
                  Into n_非合约
                  From 合作单位计划控制
                  Where 计划id = r_计划id And 合作单位 = v_合作单位 And Rownum < 2;
                Exception
                  When Others Then
                    n_非合约 := 1;
                End;
                If n_Exists = 1 Or n_非合约 = 1 Then
                  If n_最大可用数量 < n_单个剩余 Then
                    v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                      To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_最大可用数量 || '</SL>' ||
                                      '</SPAN>';
                    n_合约剩余数量 := n_合约剩余数量 + Nvl(n_单个剩余, 0);
                  Else
                    v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                      To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                      '</SPAN>';
                    n_合约剩余数量 := n_合约剩余数量 + Nvl(n_单个剩余, 0);
                  End If;
                End If;
              End If;
            End Loop;
          End If;
        Else
          Select Max(结束时间)
          Into d_加号时间
          From 挂号安排时段
          Where 安排id = r_安排id And 星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四',
                                              '6', '周五', '7', '周六', Null);
          If r_序号控制 = 1 Then
            If Trunc(d_日期) = Trunc(Sysdate) Then
              n_特殊预约 := 0;
            Else
              Select Nvl(Max(Ap.是否预约), 0)
              Into n_特殊预约
              From (Select Sd.安排id, Sd.序号, Sd.星期, Ap.号码,
                            To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'),
                                     'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                            To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'),
                                     'yyyy-mm-dd hh24:mi:ss') As 结束时间, Sd.限制数量, Sd.是否预约
                     From 挂号安排 Ap, 挂号安排时段 Sd
                     Where Ap.Id = Sd.安排id And Ap.Id = r_安排id And
                           Sd.星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6',
                                          '周五', '7', '周六', Null)) Ap, 挂号序号状态 Zt
              Where Zt.日期(+) = Ap.开始时间 And Zt.号码(+) = Ap.号码 And Zt.序号(+) = Ap.序号 And
                    Decode(Sign(Sysdate - Ap.开始时间), -1, 0, 1) <> 1;
            End If;
            For r_Time In (Select Ap.号码, Ap.序号, Ap.星期, Ap.开始时间, Ap.结束时间, Ap.限制数量, Ap.是否预约, 0 As 已约数,
                                  Decode(Nvl(Zt.序号, 0), 0, 1, 0) As 剩余数,
                                  Decode(Sign(Sysdate - Ap.开始时间), -1, 0, 1) As 失效时段
                           
                           From (Select Sd.安排id, Sd.序号, Sd.星期, Ap.号码,
                                         To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'),
                                                  'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                                         To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'),
                                                  'yyyy-mm-dd hh24:mi:ss') As 结束时间, Sd.限制数量, Sd.是否预约
                                  From 挂号安排 Ap, 挂号安排时段 Sd
                                  Where Ap.Id = Sd.安排id And Ap.Id = r_安排id And
                                        Sd.星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                                       '周四', '6', '周五', '7', '周六', Null)) Ap, 挂号序号状态 Zt
                           Where Zt.日期(+) = Ap.开始时间 And Zt.号码(+) = Ap.号码 And Zt.序号(+) = Ap.序号 And
                                 Decode(Sign(Sysdate - Ap.开始时间), -1, 0, 1) <> 1
                           Order By 序号) Loop
              If v_合作单位 Is Not Null Then
                Begin
                  Select 1
                  Into n_合约模式
                  From 合作单位安排控制
                  Where 限制项目 = r_Time.星期 And 安排id = r_安排id And 序号 = 0 And 合作单位 = v_合作单位;
                Exception
                  When Others Then
                    n_合约模式 := 0;
                End;
              Else
                n_合约模式 := 0;
              End If;
              If r_Time.剩余数 = 0 Then
                n_单个剩余 := 0;
              Else
                n_单个剩余 := r_Time.限制数量;
              End If;
              If v_合作单位 Is Null Or n_合约模式 = 1 Then
                Begin
                  Select 1
                  Into n_Exists
                  From 合作单位安排控制
                  Where 限制项目 = r_Time.星期 And 安排id = r_安排id And 序号 = r_Time.序号 And Rownum < 2;
                Exception
                  When Others Then
                    n_Exists := 0;
                End;
                If n_Exists = 0 Then
                  If n_特殊预约 = 1 And r_Time.是否预约 = 0 Then
                    Null;
                  Else
                    Begin
                      Select 1
                      Into n_是否预留
                      From 挂号序号状态
                      Where 状态 In (3, 4) And 号码 = r_号码 And 序号 = r_Time.序号 And Trunc(日期) = Trunc(d_日期);
                    Exception
                      When Others Then
                        n_是否预留 := 0;
                    End;
                    If n_是否预留 = 0 Then
                      v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                    To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                    '</SPAN>';
                      n_时段数量 := Nvl(n_时段数量, 0) + n_单个剩余;
                    End If;
                  End If;
                End If;
              Else
                Begin
                  Select 1
                  Into n_Exists
                  From 合作单位安排控制
                  Where 限制项目 = r_Time.星期 And 安排id = r_安排id And 序号 = r_Time.序号 And 合作单位 = v_合作单位 And Rownum < 2;
                Exception
                  When Others Then
                    n_Exists := 0;
                End;
                Begin
                  Select 0
                  Into n_非合约
                  From 合作单位安排控制
                  Where 安排id = r_安排id And 合作单位 = v_合作单位 And Rownum < 2;
                Exception
                  When Others Then
                    n_非合约 := 1;
                End;
                If n_Exists = 1 Or n_非合约 = 1 Then
                  If n_特殊预约 = 1 And r_Time.是否预约 = 0 Then
                    Null;
                  Else
                    Begin
                      Select 1
                      Into n_是否预留
                      From 挂号序号状态
                      Where 状态 In (3, 4) And 号码 = r_号码 And 序号 = r_Time.序号 And Trunc(日期) = Trunc(d_日期);
                    Exception
                      When Others Then
                        n_是否预留 := 0;
                    End;
                    If n_是否预留 = 0 Then
                      v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                        To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                        '</SPAN>';
                      n_合约剩余数量 := n_合约剩余数量 + 1;
                    End If;
                  End If;
                End If;
              End If;
            End Loop;
          Else
            n_最大可用数量 := Nvl(r_限约数, Nvl(r_限号数, 0)) - Nvl(r_已约数, 0);
            For r_Time In (Select Ap.号码, Ap.序号, Ap.星期, Ap.开始时间, Ap.结束时间, Ap.限制数量, Ap.是否预约,
                                  Sum(Decode(Nvl(Zt.序号, 0), 0, 0, 1)) As 已约数,
                                  Ap.限制数量 - Sum(Decode(Nvl(Zt.序号, 0), 0, 0, 1)) As 剩余数,
                                  Decode(Sign(Sysdate - Ap.开始时间), -1, 0, 1) As 失效时段
                           From (Select Sd.安排id, Sd.序号, Sd.星期, Ap.号码,
                                         To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'),
                                                  'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                                         To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'),
                                                  'yyyy-mm-dd hh24:mi:ss') As 结束时间, Sd.限制数量, Sd.是否预约
                                  From 挂号安排 Ap, 挂号安排时段 Sd
                                  Where Ap.Id = Sd.安排id And Ap.Id = r_安排id And
                                        Sd.星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                                       '周四', '6', '周五', '7', '周六', Null)) Ap, 挂号序号状态 Zt
                           Where Ap.号码 = Zt.号码(+) And Ap.开始时间 = Zt.日期(+) And
                                 Decode(Sign(Sysdate - Ap.开始时间), -1, 0, 1) <> 1
                           Group By Ap.号码, Ap.序号, Ap.星期, Ap.开始时间, Ap.结束时间, Ap.限制数量, Ap.是否预约
                           Order By Ap.序号) Loop
              If v_合作单位 Is Not Null Then
                Begin
                  Select 1
                  Into n_合约模式
                  From 合作单位安排控制
                  Where 限制项目 = r_Time.星期 And 安排id = r_安排id And 序号 = 0 And 合作单位 = v_合作单位;
                Exception
                  When Others Then
                    n_合约模式 := 0;
                End;
              Else
                n_合约模式 := 0;
              End If;
              n_单个剩余 := r_Time.剩余数;
              If v_合作单位 Is Null Or n_合约模式 = 1 Then
                Begin
                  Select 1
                  Into n_Exists
                  From 合作单位安排控制
                  Where 限制项目 = r_Time.星期 And 安排id = r_安排id And 序号 = r_Time.序号 And Rownum < 2;
                Exception
                  When Others Then
                    n_Exists := 0;
                End;
                If n_Exists = 0 Then
                  If n_最大可用数量 < n_单个剩余 Then
                    v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_最大可用数量 || '</SL>' ||
                                  '</SPAN>';
                    n_时段数量 := Nvl(n_时段数量, 0) + n_最大可用数量;
                  Else
                    v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                  '</SPAN>';
                    n_时段数量 := Nvl(n_时段数量, 0) + n_单个剩余;
                  End If;
                End If;
              Else
                Begin
                  Select 1
                  Into n_Exists
                  From 合作单位安排控制
                  Where 限制项目 = r_Time.星期 And 安排id = r_安排id And 序号 = r_Time.序号 And 合作单位 = v_合作单位 And Rownum < 2;
                Exception
                  When Others Then
                    n_Exists := 0;
                End;
                Begin
                  Select 0
                  Into n_非合约
                  From 合作单位安排控制
                  Where 安排id = r_安排id And 合作单位 = v_合作单位 And Rownum < 2;
                Exception
                  When Others Then
                    n_非合约 := 1;
                End;
                If n_Exists = 1 Or n_非合约 = 1 Then
                  If n_最大可用数量 < n_单个剩余 Then
                    v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                      To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_最大可用数量 || '</SL>' ||
                                      '</SPAN>';
                    n_合约剩余数量 := n_合约剩余数量 + Nvl(n_单个剩余, 0);
                  Else
                    v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                      To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                      '</SPAN>';
                    n_合约剩余数量 := n_合约剩余数量 + Nvl(n_单个剩余, 0);
                  End If;
                End If;
              End If;
            End Loop;
          End If;
        End If;
      End If;
      If v_合作单位 Is Not Null Then
        If Nvl(r_计划id, 0) <> 0 Then
          Begin
            Select 0
            Into n_非合约
            From 合作单位计划控制
            Where 计划id = r_计划id And 合作单位 = v_合作单位 And Rownum < 2;
          Exception
            When Others Then
              n_非合约 := 1;
          End;
        Else
          Begin
            Select 0
            Into n_非合约
            From 合作单位安排控制
            Where 安排id = r_安排id And 合作单位 = v_合作单位 And Rownum < 2;
          Exception
            When Others Then
              n_非合约 := 1;
          End;
        End If;
      End If;
      If v_合作单位 Is Null Or n_非合约 = 1 Then
        If r_限号数 = 0 Then
          v_剩余数量 := '';
        Else
          If Nvl(r_计划id, 0) <> 0 Then
            Select Sum(数量)
            Into n_合约总数量
            From 合作单位计划控制
            Where 计划id = r_计划id And 限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                                  '周四', '6', '周五', '7', '周六', Null);
          Else
            Select Sum(数量)
            Into n_合约总数量
            From 合作单位安排控制
            Where 安排id = r_安排id And 限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                                  '周四', '6', '周五', '7', '周六', Null);
          End If;
          Begin
            Select Count(1)
            Into n_合约已挂数
            From 病人挂号记录
            Where 号别 = r_号码 And 记录状态 = 1 And 合作单位 Is Not Null And 发生时间 Between Trunc(d_日期) And
                  Trunc(d_日期 + 1) - 1 / 60 / 60 / 24;
          Exception
            When Others Then
              n_合约已挂数 := 0;
          End;
          Select Count(1)
          Into n_预留数量
          From 挂号序号状态
          Where 状态 = 3 And 号码 = r_号码 And Trunc(日期) = Trunc(d_日期);
          If Trunc(d_日期) = Trunc(Sysdate) Then
            If Nvl(n_合约总数量, 0) = 0 Then
              v_剩余数量 := r_限号数 - r_已挂数 - r_已约数 + r_已接收 - n_预留数量;
            Else
              v_剩余数量 := r_限号数 - r_已挂数 - r_已约数 + r_已接收 - n_合约总数量 + n_合约已挂数 - n_预留数量;
            End If;
            n_已挂数 := r_已挂数;
            If Nvl(n_时段数量, 0) < v_剩余数量 And n_分时段 <> 0 Then
              n_缓冲序号 := 1;
              v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_加号时间, 'hh24:mi:ss') || '-' || '</SJD>' || '<SL>' ||
                            To_Number(v_剩余数量 - Nvl(n_时段数量, 0)) || '</SL>' || '</SPAN>';
            Else
              n_缓冲序号 := 0;
            End If;
          Else
            If Nvl(n_合约总数量, 0) = 0 Then
              v_剩余数量 := r_限约数 - r_已约数 - n_预留数量;
              If v_剩余数量 Is Null Then
                v_剩余数量 := r_限号数 - r_已挂数 - r_已约数 + r_已接收 - n_预留数量;
              End If;
            Else
              v_剩余数量 := r_限约数 - r_已约数 - n_合约总数量 + n_合约已挂数 - n_预留数量;
              If v_剩余数量 Is Null Then
                v_剩余数量 := r_限号数 - r_已挂数 - r_已约数 + r_已接收 - n_合约总数量 + n_合约已挂数 - n_预留数量;
              End If;
            End If;
            n_已挂数 := r_已挂数;
          End If;
        End If;
      Else
        If Nvl(r_计划id, 0) <> 0 Then
          If v_合作单位 Is Not Null Then
            Begin
              Select 1
              Into n_合约模式
              From 合作单位计划控制
              Where 限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                                  '7', '周六', Null) And 计划id = r_计划id And 序号 = 0 And 合作单位 = v_合作单位;
            Exception
              When Others Then
                n_合约模式 := 0;
            End;
          Else
            n_合约模式 := 0;
          End If;
          Select Sum(数量)
          Into n_合约总数量
          From 合作单位计划控制
          Where 计划id = r_计划id And 限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                                '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位;
        Else
          If v_合作单位 Is Not Null Then
            Begin
              Select 1
              Into n_合约模式
              From 合作单位安排控制
              Where 限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                                  '7', '周六', Null) And 安排id = r_安排id And 序号 = 0 And 合作单位 = v_合作单位;
            Exception
              When Others Then
                n_合约模式 := 0;
            End;
          Else
            n_合约模式 := 0;
          End If;
          Select Sum(数量)
          Into n_合约总数量
          From 合作单位安排控制
          Where 安排id = r_安排id And 限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                                '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位;
        End If;
        If n_合约模式 = 0 Then
          v_剩余数量   := n_合约剩余数量;
          n_已挂数     := r_已挂数;
          n_合约已挂数 := Nvl(n_合约总数量, 0) - n_合约剩余数量;
        Else
          n_已挂数 := r_已挂数;
          Begin
            Select Count(1)
            Into n_合约已挂数
            From 病人挂号记录
            Where 号别 = r_号码 And 记录状态 = 1 And 合作单位 = v_合作单位 And 发生时间 Between Trunc(d_日期) And
                  Trunc(d_日期 + 1) - 1 / 60 / 60 / 24;
          Exception
            When Others Then
              n_合约已挂数 := 0;
          End;
          If Nvl(n_合约总数量, 0) = 0 Then
            v_剩余数量 := '0';
          Else
            v_剩余数量 := n_合约总数量 - n_合约已挂数;
          End If;
        End If;
      End If;
      Select To_Char(开始时间, 'hh24:mi') Into v_Timetemp From 时间段 Where 时间段 = r_排班;
      v_时间段 := v_Timetemp || '-';
      Select To_Char(终止时间, 'hh24:mi') Into v_Timetemp From 时间段 Where 时间段 = r_排班;
      v_时间段 := v_时间段 || v_Timetemp;
      If v_Temp Is Not Null Then
        v_Temp := v_Temp || '</SPANLIST>';
      End If;
      If v_合作单位 Is Not Null Then
        If Nvl(r_计划id, 0) <> 0 Then
          Begin
            Select 1
            Into n_禁用
            From 合作单位计划控制
            Where 计划id = r_计划id And 合作单位 = v_合作单位 And 数量 = 0 And Rownum < 2;
          Exception
            When Others Then
              n_禁用 := 0;
          End;
        Else
          Begin
            Select 1
            Into n_禁用
            From 合作单位安排控制
            Where 安排id = r_安排id And 合作单位 = v_合作单位 And 数量 = 0 And Rownum < 2;
          Exception
            When Others Then
              n_禁用 := 0;
          End;
        End If;
      End If;
      --限约数=0的预约禁止
      If Trunc(d_日期) <> Trunc(Sysdate) Then
        If r_限约数 = 0 Then
          n_禁用 := 1;
        End If;
      End If;
      If Nvl(n_禁用, 0) = 0 Then
        --从项金额计算
        n_合计金额 := r_价格;
        For r_Subfee In (Select 现价, 从项数次
                         From 收费从属项目 A, 收费价目 B
                         Where a.主项id = r_项目id And a.从项id = b.收费细目id And Sysdate Between b.执行日期 And
                               Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))) Loop
          n_合计金额 := n_合计金额 + r_Subfee.现价 * r_Subfee.从项数次;
        End Loop;
        If Trunc(Sysdate) = Trunc(d_日期) Then
          Begin
            Select 1
            Into n_Exists
            From (Select 时间段
                   From 时间段
                   Where ('3000-01-10 ' || To_Char(Sysdate, 'HH24:MI:SS') < '3000-01-10 ' || To_Char(终止时间, 'HH24:MI:SS')) Or
                         ('3000-01-10 ' || To_Char(Sysdate, 'HH24:MI:SS') <
                         Decode(Sign(开始时间 - 终止时间), 1, '3000-01-11 ' || To_Char(终止时间, 'HH24:MI:SS'),
                                 '3000-01-10 ' || To_Char(终止时间, 'HH24:MI:SS'))))
            Where 时间段 = r_排班;
          Exception
            When Others Then
              n_Exists := 0;
          End;
        Else
          n_Exists := 1;
        End If;
        If n_Exists = 1 Then
          If v_剩余数量 > 0 Then
            c_Xmlmain := '<HB>' || '<HM>' || r_号码 || '</HM>' || '<YSID>' || r_医生id || '</YSID>' || '<YS>' || r_医生姓名 ||
                         '</YS>' || '<KSID>' || r_科室id || '</KSID>' || '<KSMC>' || r_科室名称 || '</KSMC>' || '<ZC>' || r_职称 ||
                         '</ZC>' || '<XMID>' || r_项目id || '</XMID>' || '<XMMC>' || r_项目名称 || '</XMMC>' || '<YGHS>' ||
                         n_已挂数 || '</YGHS>' || '<SYHS>' || v_剩余数量 || '</SYHS>' || '<PRICE>' || n_合计金额 || '</PRICE>' ||
                         '<HCXH>' || n_缓冲序号 || '</HCXH>' || '<HL>' || r_号类 || '</HL>' || '<FSD>' || n_分时段 || '</FSD>' ||
                         '<HBTIME>' || v_时间段 || '</HBTIME>' || '<FWMC>' || r_排班 || '</FWMC>' || v_Temp || '</HB>';
            v_Xmlmain := v_Xmlmain || c_Xmlmain;
          Else
            c_Xmlmain := '<HB>' || '<HM>' || r_号码 || '</HM>' || '<YSID>' || r_医生id || '</YSID>' || '<YS>' || r_医生姓名 ||
                         '</YS>' || '<KSID>' || r_科室id || '</KSID>' || '<KSMC>' || r_科室名称 || '</KSMC>' || '<ZC>' || r_职称 ||
                         '</ZC>' || '<XMID>' || r_项目id || '</XMID>' || '<XMMC>' || r_项目名称 || '</XMMC>' || '<YGHS>' ||
                         n_已挂数 || '</YGHS>' || '<SYHS>' || v_剩余数量 || '</SYHS>' || '<PRICE>' || n_合计金额 || '</PRICE>' ||
                         '<HL>' || r_号类 || '</HL>' || '<FSD>' || n_分时段 || '</FSD>' || '<HBTIME>' || v_时间段 ||
                         '</HBTIME>' || '<FWMC>' || r_排班 || '</FWMC>' || '</HB>';
            v_Xmlmain := v_Xmlmain || c_Xmlmain;
          End If;
        End If;
      End If;
      n_合约剩余数量 := 0;
      n_合约总数量   := 0;
      n_时段数量     := 0;
      n_禁用         := 0;
      n_非合约       := 0;
    End Loop;
    Close r_No;
    v_Xmlmain := '<GROUP>' || '<RQ>' || To_Char(d_日期, 'yyyy-mm-dd') || '</RQ>' || '<HBLIST>' || v_Xmlmain ||
                 '</HBLIST>' || '</GROUP>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Xmlmain)) Into x_Templet From Dual;
  Else
    If Trunc(d_日期) < To_Date(Substr(v_启用时间, 1, Instr(v_启用时间, ' ') - 1), 'yyyy-mm-dd') Then
      Select Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六',
                     Null)
      Into v_星期
      From Dual;
      n_合约剩余数量 := 0;
    
      v_Sql := 'Select a.*, Nvl(Hz.已挂数, 0) As 已挂数, Nvl(Hz.已约数, 0) As 已约数, Nvl(Hz.其中已接收, 0) As 已接收, b.现价 As 价格 ';
      v_Sql := v_Sql ||
               'From (Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Ap.医生姓名, Nvl(Ap.医生id, 0) As 医生id, Ry.专业技术职务 As 职称, Ap.号码, ';
      v_Sql := v_Sql || ' Ap.安排id, Ap.计划id, Ap.排班, Ap.项目id, Fy.名称 As 项目名称, Ap.序号控制, Ap.限号数, Ap.限约数 ';
      v_Sql := v_Sql ||
               'From (Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Ap.医生姓名, Ap.医生id, Ap.号码, Ap.Id As 安排id, 0 As 计划id, ';
      v_Sql := v_Sql || 'Ap.项目id, Nvl(Ap.序号控制, 0) As 序号控制, ';
      v_Sql := v_Sql ||
               'Decode(To_Char(:1, ''D''), ''1'', Ap.周日, ''2'', Ap.周一, ''3'', Ap.周二, ''4'', Ap.周三, ''5'', Ap.周四, ';
      v_Sql := v_Sql || ' ''6'', Ap.周五, ''7'', Ap.周六, Null) As 排班, Xz.限约数, Xz.限号数 ';
      v_Sql := v_Sql || 'From 挂号安排 Ap, 部门表 Bm, 挂号安排限制 Xz ';
      v_Sql := v_Sql || 'Where Ap.科室id = Bm.Id(+) ';
    
      n_Curcount := 2;
      If Nvl(n_科室id, 0) <> 0 Then
        v_Sql      := v_Sql || 'And Ap.科室id = :2 ';
        n_Curcount := n_Curcount + 1;
      End If;
      If Nvl(n_医生id, 0) <> 0 Then
        If n_Curcount = 2 Then
          v_Sql := v_Sql || 'And Ap.医生id = :2 ';
        Else
          v_Sql := v_Sql || 'And Ap.医生id = :3 ';
        End If;
        n_Curcount := n_Curcount + 1;
      End If;
      If Nvl(v_医生姓名, '_') <> '_' Then
        If n_Curcount = 2 Then
          v_Sql := v_Sql || 'And Ap.医生姓名 = :2 ';
        End If;
        If n_Curcount = 3 Then
          v_Sql := v_Sql || 'And Ap.医生姓名 = :3 ';
        End If;
        If n_Curcount = 4 Then
          v_Sql := v_Sql || 'And Ap.医生姓名 = :4 ';
        End If;
        n_Curcount := n_Curcount + 1;
      End If;
    
      v_Sql      := v_Sql || 'And Ap.停用日期 Is Null And :' || n_Curcount ||
                    ' Between Nvl(Ap.开始时间, To_Date(''1900-01-01'', ''YYYY-MM-DD'')) And ';
      n_Curcount := n_Curcount + 1;
      v_Sql      := v_Sql || 'Nvl(Ap.终止时间, To_Date(''3000 - 01 - 01'', ''YYYY-MM-DD'')) And Xz.安排id(+) = Ap.Id And ';
      v_Sql      := v_Sql || ' Xz.限制项目(+) = Decode(To_Char(:' || n_Curcount ||
                    ', ''D''), ''1'', ''周日'', ''2'', ''周一'', ''3'', ''周二'', ''4'', ''周三'', ''5'', ';
      n_Curcount := n_Curcount + 1;
      v_Sql      := v_Sql || ' ''周四'', ''6'', ''周五'', ''7'', ''周六'', Null) And Not Exists ';
      v_Sql      := v_Sql || '(Select Rownum ';
      v_Sql      := v_Sql || 'From 挂号安排停用状态 Ty ';
      v_Sql      := v_Sql || 'Where Ty.安排id = Ap.Id And :' || n_Curcount ||
                    ' Between Ty.开始停止时间 And Ty.结束停止时间) And Not Exists ';
      n_Curcount := n_Curcount + 1;
      v_Sql      := v_Sql || '(Select Rownum ';
      v_Sql      := v_Sql || 'From 挂号安排计划 Jh Where Jh.安排id = Ap.Id And Jh.审核时间 Is Not Null And ';
      v_Sql      := v_Sql || ':' || n_Curcount ||
                    ' Between Nvl(Jh.生效时间, To_Date(''1900-01-01'', ''YYYY-MM-DD'')) And Nvl(Jh.失效时间, To_Date(''3000-01-01'', ''YYYY-MM-DD''))) ';
      n_Curcount := n_Curcount + 1;
      v_Sql      := v_Sql || 'Union All ';
      v_Sql      := v_Sql ||
                    'Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Jh.医生姓名, Jh.医生id, Ap.号码, Ap.Id As 安排id, Jh.Id As 计划id, ';
      v_Sql      := v_Sql || 'Jh.项目id, Nvl(Jh.序号控制, 0) As 序号控制,Decode(To_Char(:' || n_Curcount ||
                    ', ''D''), ''1'', Jh.周日, ''2'', Jh.周一, ''3'', Jh.周二, ''4'', Jh.周三, ''5'', Jh.周四, ';
      n_Curcount := n_Curcount + 1;
      v_Sql      := v_Sql || ' ''6'', Jh.周五, ''7'', Jh.周六, Null) As 排班, Xz.限约数, Xz.限号数 ';
      v_Sql      := v_Sql || 'From 挂号安排计划 Jh, 挂号安排 Ap, 部门表 Bm, 挂号计划限制 Xz ';
      v_Sql      := v_Sql || 'Where Jh.安排id = Ap.Id And Ap.科室id = Bm.Id(+) And Ap.停用日期 Is Null ';
    
      If Nvl(n_科室id, 0) <> 0 Then
        v_Sql      := v_Sql || 'And Ap.科室id = :' || n_Curcount || ' ';
        n_Curcount := n_Curcount + 1;
      End If;
      If Nvl(n_医生id, 0) <> 0 Then
        v_Sql      := v_Sql || 'And Ap.医生id = :' || n_Curcount || ' ';
        n_Curcount := n_Curcount + 1;
      End If;
      If Nvl(v_医生姓名, '_') <> '_' Then
        v_Sql      := v_Sql || 'And Ap.医生姓名 = :' || n_Curcount || ' ';
        n_Curcount := n_Curcount + 1;
      End If;
    
      v_Sql      := v_Sql || ' And :' || n_Curcount ||
                    ' Between Nvl(Jh.生效时间, To_Date(''1900-01-01'', ''YYYY-MM-DD'')) And Nvl(Jh.失效时间, To_Date(''3000-01-01'', ''YYYY-MM-DD'')) And Xz.计划id(+) = Jh.Id And ';
      n_Curcount := n_Curcount + 1;
      v_Sql      := v_Sql || 'Xz.限制项目(+) = Decode(To_Char(:' || n_Curcount ||
                    ', ''D''), ''1'', ''周日'', ''2'', ''周一'', ''3'', ''周二'', ''4'', ''周三'', ''5'', ''周四'', ''6'', ''周五'', ''7'', ''周六'', Null) And Not Exists ';
      n_Curcount := n_Curcount + 1;
      v_Sql      := v_Sql || '(Select Rownum From 挂号安排停用状态 Ty Where Ty.安排id = Ap.Id And :' || n_Curcount ||
                    ' Between Ty.开始停止时间 And Ty.结束停止时间) And ';
      n_Curcount := n_Curcount + 1;
      v_Sql      := v_Sql || '(Jh.生效时间, Jh.安排id) = (Select Max(Sxjh.生效时间) As 生效时间, 安排id From 挂号安排计划 Sxjh ';
      v_Sql      := v_Sql || ' Where Sxjh.审核时间 Is Not Null And :' || n_Curcount ||
                    ' Between Nvl(Sxjh.生效时间, To_Date(''1900-01-01'', ''YYYY-MM-DD'')) And ';
      n_Curcount := n_Curcount + 1;
      v_Sql      := v_Sql || 'Nvl(Sxjh.失效时间, To_Date(''3000-01-01'', ''YYYY-MM-DD'')) And Sxjh.安排id = Jh.安排id ';
      v_Sql      := v_Sql || 'Group By Sxjh.安排id)) Ap, 部门表 Bm, 人员表 Ry, 收费项目目录 Fy ';
      v_Sql      := v_Sql ||
                    'Where Ap.科室id = Bm.Id(+) And Ap.医生id = Ry.Id(+) And Ap.项目id = Fy.Id And 排班 Is Not Null) A, ';
      v_Sql      := v_Sql || '病人挂号汇总 Hz, 收费价目 B ';
      v_Sql      := v_Sql || 'Where a.号码 = Hz.号码(+) And Hz.日期(+) = Trunc(:' || n_Curcount ||
                    ') And a.项目id = b.收费细目id And ';
      n_Curcount := n_Curcount + 1;
      v_Sql      := v_Sql || 'Nvl(b.终止日期, To_Date(''3000-1-1'', ''YYYY-Mm-DD'')) > :' || n_Curcount || ' ';
      n_Curcount := n_Curcount + 1;
      v_Sql      := v_Sql || 'And b.执行日期 <= :' || n_Curcount || ' ';
      If Nvl(n_科室id, 0) <> 0 And Nvl(n_医生id, 0) <> 0 And Nvl(v_医生姓名, '_') <> '_' Then
        Open r_No For v_Sql
          Using d_日期, n_科室id, n_医生id, v_医生姓名, d_日期, d_日期, d_日期, d_日期, d_日期, n_科室id, n_医生id, v_医生姓名, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期;
      End If;
      If Nvl(n_科室id, 0) <> 0 And Nvl(n_医生id, 0) = 0 And Nvl(v_医生姓名, '_') = '_' Then
        Open r_No For v_Sql
          Using d_日期, n_科室id, d_日期, d_日期, d_日期, d_日期, d_日期, n_科室id, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期;
      End If;
      If Nvl(n_科室id, 0) = 0 And Nvl(n_医生id, 0) <> 0 And Nvl(v_医生姓名, '_') = '_' Then
        Open r_No For v_Sql
          Using d_日期, n_医生id, d_日期, d_日期, d_日期, d_日期, d_日期, n_医生id, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期;
      End If;
      If Nvl(n_科室id, 0) = 0 And Nvl(n_医生id, 0) = 0 And Nvl(v_医生姓名, '_') <> '_' Then
        Open r_No For v_Sql
          Using d_日期, v_医生姓名, d_日期, d_日期, d_日期, d_日期, d_日期, v_医生姓名, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期;
      End If;
      If Nvl(n_科室id, 0) <> 0 And Nvl(n_医生id, 0) <> 0 And Nvl(v_医生姓名, '_') = '_' Then
        Open r_No For v_Sql
          Using d_日期, n_科室id, n_医生id, d_日期, d_日期, d_日期, d_日期, d_日期, n_科室id, n_医生id, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期;
      End If;
      If Nvl(n_科室id, 0) <> 0 And Nvl(n_医生id, 0) = 0 And Nvl(v_医生姓名, '_') <> '_' Then
        Open r_No For v_Sql
          Using d_日期, n_科室id, v_医生姓名, d_日期, d_日期, d_日期, d_日期, d_日期, n_科室id, v_医生姓名, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期;
      End If;
      If Nvl(n_科室id, 0) = 0 And Nvl(n_医生id, 0) <> 0 And Nvl(v_医生姓名, '_') <> '_' Then
        Open r_No For v_Sql
          Using d_日期, n_医生id, v_医生姓名, d_日期, d_日期, d_日期, d_日期, d_日期, n_医生id, v_医生姓名, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期, d_日期;
      End If;
      Loop
        Fetch r_No
          Into r_科室id, r_号类, r_科室名称, r_医生姓名, r_医生id, r_职称, r_号码, r_安排id, r_计划id, r_排班, r_项目id, r_项目名称, r_序号控制, r_限号数,
               r_限约数, r_已挂数, r_已约数, r_已接收, r_价格;
        Exit When r_No%NotFound;
        If r_计划id <> 0 Then
          Select Sign(Count(Rownum))
          Into n_分时段
          From 挂号安排计划 Jh, 挂号计划时段 Sd
          Where Jh.Id = Sd.计划id And Jh.Id = r_计划id And
                Sd.星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7',
                               '周六', Null) And Rownum < 2;
        Else
          Select Sign(Count(Rownum))
          Into n_分时段
          From 挂号安排 Ap, 挂号安排时段 Sd
          Where Ap.Id = Sd.安排id And Ap.Id = r_安排id And
                Sd.星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7',
                               '周六', Null) And Rownum < 2;
        End If;
        If n_分时段 = 0 Then
          v_Temp := '';
          If v_合作单位 Is Not Null And r_序号控制 = 1 Then
            If r_计划id <> 0 Then
              Select Nvl(Sum(数量), 0)
              Into n_合约总数量
              From 合作单位计划控制
              Where 计划id = r_计划id And 合作单位 = v_合作单位 And
                    限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                                  '7', '周六', Null);
              Select Count(1)
              Into n_合约模式
              From 合作单位计划控制
              Where 计划id = r_计划id And 合作单位 = v_合作单位 And
                    限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                                  '7', '周六', Null) And 序号 = 0;
            Else
              Select Nvl(Sum(数量), 0)
              Into n_合约总数量
              From 合作单位安排控制
              Where 安排id = r_安排id And 合作单位 = v_合作单位 And
                    限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                                  '7', '周六', Null);
              Select Count(1)
              Into n_合约模式
              From 合作单位安排控制
              Where 安排id = r_安排id And 合作单位 = v_合作单位 And
                    限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                                  '7', '周六', Null) And 序号 = 0;
            End If;
            If n_合约模式 = 0 Then
              If r_计划id <> 0 Then
                Select Count(1)
                Into n_合约已挂数
                From 病人挂号记录 A
                Where 号别 = r_号码 And 记录状态 = 1 And 发生时间 Between Trunc(d_日期) And Trunc(d_日期 + 1) - 1 / 60 / 60 / 24 And
                      Exists
                 (Select 1
                       From 合作单位计划控制
                       Where 计划id = r_计划id And 合作单位 = v_合作单位 And
                             限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6',
                                           '周五', '7', '周六', Null) And 序号 = a.号序 And 数量 <> 0);
              Else
                Select Count(1)
                Into n_合约已挂数
                From 病人挂号记录 A
                Where 号别 = r_号码 And 记录状态 = 1 And 发生时间 Between Trunc(d_日期) And Trunc(d_日期 + 1) - 1 / 60 / 60 / 24 And
                      Exists
                 (Select 1
                       From 合作单位安排控制
                       Where 安排id = r_安排id And 合作单位 = v_合作单位 And
                             限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6',
                                           '周五', '7', '周六', Null) And 序号 = a.号序 And 数量 <> 0);
              End If;
            Else
              Begin
                Select Count(1)
                Into n_合约已挂数
                From 病人挂号记录
                Where 号别 = r_号码 And 记录状态 = 1 And 合作单位 = v_合作单位 And 发生时间 Between Trunc(d_日期) And
                      Trunc(d_日期 + 1) - 1 / 60 / 60 / 24;
              Exception
                When Others Then
                  n_合约已挂数 := 0;
              End;
            End If;
            If n_合约总数量 = 0 Then
              n_合约剩余数量 := 0;
            Else
              n_合约剩余数量 := n_合约总数量 - n_合约已挂数;
              If n_合约剩余数量 > (Nvl(r_限号数, 0) - r_已挂数) Then
                n_合约剩余数量 := Nvl(r_限号数, 0) - r_已挂数;
              End If;
            End If;
          End If;
        Else
          v_Temp := '<SPANLIST>';
          If r_计划id <> 0 Then
            Select Max(结束时间)
            Into d_加号时间
            From 挂号计划时段
            Where 计划id = r_计划id And 星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                                '周四', '6', '周五', '7', '周六', Null);
            If r_序号控制 = 1 Then
              If Trunc(d_日期) = Trunc(Sysdate) Then
                n_特殊预约 := 0;
              Else
                Select Nvl(Max(Jh.是否预约), 0)
                Into n_特殊预约
                From (Select Sd.计划id, Sd.序号, Sd.星期, Jh.号码,
                              To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                              To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') As 结束时间, Sd.限制数量, Sd.是否预约
                       From 挂号安排计划 Jh, 挂号计划时段 Sd
                       Where Jh.Id = Sd.计划id And Jh.Id = r_计划id And
                             Sd.星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6',
                                            '周五', '7', '周六', Null)) Jh, 挂号序号状态 Zt
                Where Zt.日期(+) = Jh.开始时间 And Zt.号码(+) = Jh.号码 And Zt.序号(+) = Jh.序号 And
                      Decode(Sign(Sysdate - Jh.开始时间), -1, 0, 1) <> 1;
              End If;
            
              For r_Time In (Select Jh.号码, Jh.序号, Jh.星期, Jh.开始时间, Jh.结束时间, Jh.限制数量, Jh.是否预约, 0 As 已约数,
                                    Decode(Nvl(Zt.序号, 0), 0, 1, 0) As 剩余数,
                                    Decode(Sign(Sysdate - Jh.开始时间), -1, 0, 1) As 失效时段
                             
                             From (Select Sd.计划id, Sd.序号, Sd.星期, Jh.号码,
                                           To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                                           To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As 结束时间, Sd.限制数量, Sd.是否预约
                                    From 挂号安排计划 Jh, 挂号计划时段 Sd
                                    Where Jh.Id = Sd.计划id And Jh.Id = r_计划id And
                                          Sd.星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                                         '5', '周四', '6', '周五', '7', '周六', Null)) Jh, 挂号序号状态 Zt
                             Where Zt.日期(+) = Jh.开始时间 And Zt.号码(+) = Jh.号码 And Zt.序号(+) = Jh.序号 And
                                   Decode(Sign(Sysdate - Jh.开始时间), -1, 0, 1) <> 1
                             Order By 序号) Loop
                If v_合作单位 Is Not Null Then
                  Begin
                    Select 1
                    Into n_合约模式
                    From 合作单位计划控制
                    Where 限制项目 = r_Time.星期 And 计划id = r_计划id And 序号 = 0 And 合作单位 = v_合作单位;
                  Exception
                    When Others Then
                      n_合约模式 := 0;
                  End;
                Else
                  n_合约模式 := 0;
                End If;
                If r_Time.剩余数 = 0 Then
                  n_单个剩余 := 0;
                Else
                  n_单个剩余 := r_Time.限制数量;
                End If;
                If v_合作单位 Is Null Or n_合约模式 = 1 Then
                  Begin
                    Select 1
                    Into n_Exists
                    From 合作单位计划控制
                    Where 限制项目 = r_Time.星期 And 计划id = r_计划id And 序号 = r_Time.序号 And Rownum < 2;
                  Exception
                    When Others Then
                      n_Exists := 0;
                  End;
                  If n_Exists = 0 Then
                    If n_特殊预约 = 1 And r_Time.是否预约 = 0 Then
                      Null;
                    Else
                      Begin
                        Select 1
                        Into n_是否预留
                        From 挂号序号状态
                        Where 状态 In (3, 4) And 号码 = r_号码 And 序号 = r_Time.序号 And Trunc(日期) = Trunc(d_日期);
                      Exception
                        When Others Then
                          n_是否预留 := 0;
                      End;
                      If n_是否预留 = 0 Then
                        v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                      To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                      '</SPAN>';
                        n_时段数量 := Nvl(n_时段数量, 0) + n_单个剩余;
                      End If;
                    End If;
                  End If;
                Else
                  Begin
                    Select 1
                    Into n_Exists
                    From 合作单位计划控制
                    Where 限制项目 = r_Time.星期 And 计划id = r_计划id And 序号 = r_Time.序号 And 合作单位 = v_合作单位 And Rownum < 2;
                  Exception
                    When Others Then
                      n_Exists := 0;
                  End;
                  Begin
                    Select 0
                    Into n_非合约
                    From 合作单位计划控制
                    Where 计划id = r_计划id And 合作单位 = v_合作单位 And Rownum < 2;
                  Exception
                    When Others Then
                      n_非合约 := 1;
                  End;
                  If n_Exists = 1 Or n_非合约 = 1 Then
                    If n_特殊预约 = 1 And r_Time.是否预约 = 0 Then
                      Null;
                    Else
                      Begin
                        Select 1
                        Into n_是否预留
                        From 挂号序号状态
                        Where 状态 In (3, 4) And 号码 = r_号码 And 序号 = r_Time.序号 And Trunc(日期) = Trunc(d_日期);
                      Exception
                        When Others Then
                          n_是否预留 := 0;
                      End;
                      If n_是否预留 = 0 Then
                        v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                          To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                          '</SPAN>';
                        n_合约剩余数量 := n_合约剩余数量 + 1;
                      End If;
                    End If;
                  End If;
                End If;
              End Loop;
            Else
              n_最大可用数量 := Nvl(r_限约数, Nvl(r_限号数, 0)) - Nvl(r_已约数, 0);
              For r_Time In (Select Jh.号码, Jh.序号, Jh.星期, Jh.开始时间, Jh.结束时间, Jh.限制数量, Jh.是否预约,
                                    Sum(Decode(Nvl(Zt.序号, 0), 0, 0, 1)) As 已约数,
                                    Jh.限制数量 - Sum(Decode(Nvl(Zt.序号, 0), 0, 0, 1)) As 剩余数,
                                    Decode(Sign(Sysdate - Jh.开始时间), -1, 0, 1) As 失效时段
                             From (Select Sd.计划id, Sd.序号, Sd.星期, Jh.号码,
                                           To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                                           To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As 结束时间, Sd.限制数量, Sd.是否预约
                                    From 挂号安排计划 Jh, 挂号计划时段 Sd
                                    Where Jh.Id = Sd.计划id And Jh.Id = r_计划id And
                                          Sd.星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                                         '5', '周四', '6', '周五', '7', '周六', Null)) Jh, 挂号序号状态 Zt
                             Where Jh.号码 = Zt.号码(+) And Jh.开始时间 = Zt.日期(+) And
                                   Decode(Sign(Sysdate - Jh.开始时间), -1, 0, 1) <> 1
                             Group By Jh.号码, Jh.序号, Jh.星期, Jh.开始时间, Jh.结束时间, Jh.限制数量, Jh.是否预约
                             Order By Jh.序号) Loop
                If v_合作单位 Is Not Null Then
                  Begin
                    Select 1
                    Into n_合约模式
                    From 合作单位计划控制
                    Where 限制项目 = r_Time.星期 And 计划id = r_计划id And 序号 = 0 And 合作单位 = v_合作单位;
                  Exception
                    When Others Then
                      n_合约模式 := 0;
                  End;
                Else
                  n_合约模式 := 0;
                End If;
                n_单个剩余 := r_Time.剩余数;
                If v_合作单位 Is Null Or n_合约模式 = 1 Then
                  Begin
                    Select 1
                    Into n_Exists
                    From 合作单位计划控制
                    Where 限制项目 = r_Time.星期 And 计划id = r_计划id And 序号 = r_Time.序号 And Rownum < 2;
                  Exception
                    When Others Then
                      n_Exists := 0;
                  End;
                  If n_Exists = 0 Then
                    If n_最大可用数量 < n_单个剩余 Then
                      v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                    To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_最大可用数量 || '</SL>' ||
                                    '</SPAN>';
                      n_时段数量 := Nvl(n_时段数量, 0) + n_最大可用数量;
                    Else
                      v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                    To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                    '</SPAN>';
                      n_时段数量 := Nvl(n_时段数量, 0) + n_单个剩余;
                    End If;
                  End If;
                Else
                  Begin
                    Select 1
                    Into n_Exists
                    From 合作单位计划控制
                    Where 限制项目 = r_Time.星期 And 计划id = r_计划id And 序号 = r_Time.序号 And 合作单位 = v_合作单位 And Rownum < 2;
                  Exception
                    When Others Then
                      n_Exists := 0;
                  End;
                  Begin
                    Select 0
                    Into n_非合约
                    From 合作单位计划控制
                    Where 计划id = r_计划id And 合作单位 = v_合作单位 And Rownum < 2;
                  Exception
                    When Others Then
                      n_非合约 := 1;
                  End;
                  If n_Exists = 1 Or n_非合约 = 1 Then
                    If n_最大可用数量 < n_单个剩余 Then
                      v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                        To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_最大可用数量 || '</SL>' ||
                                        '</SPAN>';
                      n_合约剩余数量 := n_合约剩余数量 + Nvl(n_单个剩余, 0);
                    Else
                      v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                        To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                        '</SPAN>';
                      n_合约剩余数量 := n_合约剩余数量 + Nvl(n_单个剩余, 0);
                    End If;
                  End If;
                End If;
              End Loop;
            End If;
          Else
            Select Max(结束时间)
            Into d_加号时间
            From 挂号安排时段
            Where 安排id = r_安排id And 星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                                '周四', '6', '周五', '7', '周六', Null);
            If r_序号控制 = 1 Then
              If Trunc(d_日期) = Trunc(Sysdate) Then
                n_特殊预约 := 0;
              Else
                Select Nvl(Max(Ap.是否预约), 0)
                Into n_特殊预约
                From (Select Sd.安排id, Sd.序号, Sd.星期, Ap.号码,
                              To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                              To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') As 结束时间, Sd.限制数量, Sd.是否预约
                       From 挂号安排 Ap, 挂号安排时段 Sd
                       Where Ap.Id = Sd.安排id And Ap.Id = r_安排id And
                             Sd.星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6',
                                            '周五', '7', '周六', Null)) Ap, 挂号序号状态 Zt
                Where Zt.日期(+) = Ap.开始时间 And Zt.号码(+) = Ap.号码 And Zt.序号(+) = Ap.序号 And
                      Decode(Sign(Sysdate - Ap.开始时间), -1, 0, 1) <> 1;
              End If;
              For r_Time In (Select Ap.号码, Ap.序号, Ap.星期, Ap.开始时间, Ap.结束时间, Ap.限制数量, Ap.是否预约, 0 As 已约数,
                                    Decode(Nvl(Zt.序号, 0), 0, 1, 0) As 剩余数,
                                    Decode(Sign(Sysdate - Ap.开始时间), -1, 0, 1) As 失效时段
                             
                             From (Select Sd.安排id, Sd.序号, Sd.星期, Ap.号码,
                                           To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                                           To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As 结束时间, Sd.限制数量, Sd.是否预约
                                    From 挂号安排 Ap, 挂号安排时段 Sd
                                    Where Ap.Id = Sd.安排id And Ap.Id = r_安排id And
                                          Sd.星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                                         '5', '周四', '6', '周五', '7', '周六', Null)) Ap, 挂号序号状态 Zt
                             Where Zt.日期(+) = Ap.开始时间 And Zt.号码(+) = Ap.号码 And Zt.序号(+) = Ap.序号 And
                                   Decode(Sign(Sysdate - Ap.开始时间), -1, 0, 1) <> 1
                             Order By 序号) Loop
                If v_合作单位 Is Not Null Then
                  Begin
                    Select 1
                    Into n_合约模式
                    From 合作单位安排控制
                    Where 限制项目 = r_Time.星期 And 安排id = r_安排id And 序号 = 0 And 合作单位 = v_合作单位;
                  Exception
                    When Others Then
                      n_合约模式 := 0;
                  End;
                Else
                  n_合约模式 := 0;
                End If;
                If r_Time.剩余数 = 0 Then
                  n_单个剩余 := 0;
                Else
                  n_单个剩余 := r_Time.限制数量;
                End If;
                If v_合作单位 Is Null Or n_合约模式 = 1 Then
                  Begin
                    Select 1
                    Into n_Exists
                    From 合作单位安排控制
                    Where 限制项目 = r_Time.星期 And 安排id = r_安排id And 序号 = r_Time.序号 And Rownum < 2;
                  Exception
                    When Others Then
                      n_Exists := 0;
                  End;
                  If n_Exists = 0 Then
                    If n_特殊预约 = 1 And r_Time.是否预约 = 0 Then
                      Null;
                    Else
                      Begin
                        Select 1
                        Into n_是否预留
                        From 挂号序号状态
                        Where 状态 In (3, 4) And 号码 = r_号码 And 序号 = r_Time.序号 And Trunc(日期) = Trunc(d_日期);
                      Exception
                        When Others Then
                          n_是否预留 := 0;
                      End;
                      If n_是否预留 = 0 Then
                        v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                      To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                      '</SPAN>';
                        n_时段数量 := Nvl(n_时段数量, 0) + n_单个剩余;
                      End If;
                    End If;
                  End If;
                Else
                  Begin
                    Select 1
                    Into n_Exists
                    From 合作单位安排控制
                    Where 限制项目 = r_Time.星期 And 安排id = r_安排id And 序号 = r_Time.序号 And 合作单位 = v_合作单位 And Rownum < 2;
                  Exception
                    When Others Then
                      n_Exists := 0;
                  End;
                  Begin
                    Select 0
                    Into n_非合约
                    From 合作单位安排控制
                    Where 安排id = r_安排id And 合作单位 = v_合作单位 And Rownum < 2;
                  Exception
                    When Others Then
                      n_非合约 := 1;
                  End;
                  If n_Exists = 1 Or n_非合约 = 1 Then
                    If n_特殊预约 = 1 And r_Time.是否预约 = 0 Then
                      Null;
                    Else
                      Begin
                        Select 1
                        Into n_是否预留
                        From 挂号序号状态
                        Where 状态 In (3, 4) And 号码 = r_号码 And 序号 = r_Time.序号 And Trunc(日期) = Trunc(d_日期);
                      Exception
                        When Others Then
                          n_是否预留 := 0;
                      End;
                      If n_是否预留 = 0 Then
                        v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                          To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                          '</SPAN>';
                        n_合约剩余数量 := n_合约剩余数量 + 1;
                      End If;
                    End If;
                  End If;
                End If;
              End Loop;
            Else
              n_最大可用数量 := Nvl(r_限约数, Nvl(r_限号数, 0)) - Nvl(r_已约数, 0);
              For r_Time In (Select Ap.号码, Ap.序号, Ap.星期, Ap.开始时间, Ap.结束时间, Ap.限制数量, Ap.是否预约,
                                    Sum(Decode(Nvl(Zt.序号, 0), 0, 0, 1)) As 已约数,
                                    Ap.限制数量 - Sum(Decode(Nvl(Zt.序号, 0), 0, 0, 1)) As 剩余数,
                                    Decode(Sign(Sysdate - Ap.开始时间), -1, 0, 1) As 失效时段
                             From (Select Sd.安排id, Sd.序号, Sd.星期, Ap.号码,
                                           To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                                           To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As 结束时间, Sd.限制数量, Sd.是否预约
                                    From 挂号安排 Ap, 挂号安排时段 Sd
                                    Where Ap.Id = Sd.安排id And Ap.Id = r_安排id And
                                          Sd.星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                                         '5', '周四', '6', '周五', '7', '周六', Null)) Ap, 挂号序号状态 Zt
                             Where Ap.号码 = Zt.号码(+) And Ap.开始时间 = Zt.日期(+) And
                                   Decode(Sign(Sysdate - Ap.开始时间), -1, 0, 1) <> 1
                             Group By Ap.号码, Ap.序号, Ap.星期, Ap.开始时间, Ap.结束时间, Ap.限制数量, Ap.是否预约
                             Order By Ap.序号) Loop
                If v_合作单位 Is Not Null Then
                  Begin
                    Select 1
                    Into n_合约模式
                    From 合作单位安排控制
                    Where 限制项目 = r_Time.星期 And 安排id = r_安排id And 序号 = 0 And 合作单位 = v_合作单位;
                  Exception
                    When Others Then
                      n_合约模式 := 0;
                  End;
                Else
                  n_合约模式 := 0;
                End If;
                n_单个剩余 := r_Time.剩余数;
                If v_合作单位 Is Null Or n_合约模式 = 1 Then
                  Begin
                    Select 1
                    Into n_Exists
                    From 合作单位安排控制
                    Where 限制项目 = r_Time.星期 And 安排id = r_安排id And 序号 = r_Time.序号 And Rownum < 2;
                  Exception
                    When Others Then
                      n_Exists := 0;
                  End;
                  If n_Exists = 0 Then
                    If n_最大可用数量 < n_单个剩余 Then
                      v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                    To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_最大可用数量 || '</SL>' ||
                                    '</SPAN>';
                      n_时段数量 := Nvl(n_时段数量, 0) + n_最大可用数量;
                    Else
                      v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                    To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                    '</SPAN>';
                      n_时段数量 := Nvl(n_时段数量, 0) + n_单个剩余;
                    End If;
                  End If;
                Else
                  Begin
                    Select 1
                    Into n_Exists
                    From 合作单位安排控制
                    Where 限制项目 = r_Time.星期 And 安排id = r_安排id And 序号 = r_Time.序号 And 合作单位 = v_合作单位 And Rownum < 2;
                  Exception
                    When Others Then
                      n_Exists := 0;
                  End;
                  Begin
                    Select 0
                    Into n_非合约
                    From 合作单位安排控制
                    Where 安排id = r_安排id And 合作单位 = v_合作单位 And Rownum < 2;
                  Exception
                    When Others Then
                      n_非合约 := 1;
                  End;
                  If n_Exists = 1 Or n_非合约 = 1 Then
                    If n_最大可用数量 < n_单个剩余 Then
                      v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                        To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_最大可用数量 || '</SL>' ||
                                        '</SPAN>';
                      n_合约剩余数量 := n_合约剩余数量 + Nvl(n_单个剩余, 0);
                    Else
                      v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                        To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                        '</SPAN>';
                      n_合约剩余数量 := n_合约剩余数量 + Nvl(n_单个剩余, 0);
                    End If;
                  End If;
                End If;
              End Loop;
            End If;
          End If;
        End If;
        If v_合作单位 Is Not Null Then
          If Nvl(r_计划id, 0) <> 0 Then
            Begin
              Select 0
              Into n_非合约
              From 合作单位计划控制
              Where 计划id = r_计划id And 合作单位 = v_合作单位 And Rownum < 2;
            Exception
              When Others Then
                n_非合约 := 1;
            End;
          Else
            Begin
              Select 0
              Into n_非合约
              From 合作单位安排控制
              Where 安排id = r_安排id And 合作单位 = v_合作单位 And Rownum < 2;
            Exception
              When Others Then
                n_非合约 := 1;
            End;
          End If;
        End If;
        If v_合作单位 Is Null Or n_非合约 = 1 Then
          If r_限号数 = 0 Then
            v_剩余数量 := '';
          Else
            If Nvl(r_计划id, 0) <> 0 Then
              Select Sum(数量)
              Into n_合约总数量
              From 合作单位计划控制
              Where 计划id = r_计划id And 限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                                    '周四', '6', '周五', '7', '周六', Null);
            Else
              Select Sum(数量)
              Into n_合约总数量
              From 合作单位安排控制
              Where 安排id = r_安排id And 限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                                    '周四', '6', '周五', '7', '周六', Null);
            End If;
            Begin
              Select Count(1)
              Into n_合约已挂数
              From 病人挂号记录
              Where 号别 = r_号码 And 记录状态 = 1 And 合作单位 Is Not Null And 发生时间 Between Trunc(d_日期) And
                    Trunc(d_日期 + 1) - 1 / 60 / 60 / 24;
            Exception
              When Others Then
                n_合约已挂数 := 0;
            End;
            Select Count(1)
            Into n_预留数量
            From 挂号序号状态
            Where 状态 = 3 And 号码 = r_号码 And Trunc(日期) = Trunc(d_日期);
            If Trunc(d_日期) = Trunc(Sysdate) Then
              If Nvl(n_合约总数量, 0) = 0 Then
                v_剩余数量 := r_限号数 - r_已挂数 - r_已约数 + r_已接收 - n_预留数量;
              Else
                v_剩余数量 := r_限号数 - r_已挂数 - r_已约数 + r_已接收 - n_合约总数量 + n_合约已挂数 - n_预留数量;
              End If;
              n_已挂数 := r_已挂数;
              If Nvl(n_时段数量, 0) < v_剩余数量 And n_分时段 <> 0 Then
                n_缓冲序号 := 1;
                v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_加号时间, 'hh24:mi:ss') || '-' || '</SJD>' ||
                              '<SL>' || To_Number(v_剩余数量 - Nvl(n_时段数量, 0)) || '</SL>' || '</SPAN>';
              Else
                n_缓冲序号 := 0;
              End If;
            Else
              If Nvl(n_合约总数量, 0) = 0 Then
                v_剩余数量 := r_限约数 - r_已约数 - n_预留数量;
                If v_剩余数量 Is Null Then
                  v_剩余数量 := r_限号数 - r_已挂数 - r_已约数 + r_已接收 - n_预留数量;
                End If;
              Else
                v_剩余数量 := r_限约数 - r_已约数 - n_合约总数量 + n_合约已挂数 - n_预留数量;
                If v_剩余数量 Is Null Then
                  v_剩余数量 := r_限号数 - r_已挂数 - r_已约数 + r_已接收 - n_合约总数量 + n_合约已挂数 - n_预留数量;
                End If;
              End If;
              n_已挂数 := r_已挂数;
            End If;
          End If;
        Else
          If Nvl(r_计划id, 0) <> 0 Then
            If v_合作单位 Is Not Null Then
              Begin
                Select 1
                Into n_合约模式
                From 合作单位计划控制
                Where 限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                                    '7', '周六', Null) And 计划id = r_计划id And 序号 = 0 And 合作单位 = v_合作单位;
              Exception
                When Others Then
                  n_合约模式 := 0;
              End;
            Else
              n_合约模式 := 0;
            End If;
            Select Sum(数量)
            Into n_合约总数量
            From 合作单位计划控制
            Where 计划id = r_计划id And 限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                                  '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位;
          Else
            If v_合作单位 Is Not Null Then
              Begin
                Select 1
                Into n_合约模式
                From 合作单位安排控制
                Where 限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                                    '7', '周六', Null) And 安排id = r_安排id And 序号 = 0 And 合作单位 = v_合作单位;
              Exception
                When Others Then
                  n_合约模式 := 0;
              End;
            Else
              n_合约模式 := 0;
            End If;
            Select Sum(数量)
            Into n_合约总数量
            From 合作单位安排控制
            Where 安排id = r_安排id And 限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                                  '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位;
          End If;
          If n_合约模式 = 0 Then
            v_剩余数量   := n_合约剩余数量;
            n_已挂数     := r_已挂数;
            n_合约已挂数 := Nvl(n_合约总数量, 0) - n_合约剩余数量;
          Else
            n_已挂数 := r_已挂数;
            Begin
              Select Count(1)
              Into n_合约已挂数
              From 病人挂号记录
              Where 号别 = r_号码 And 记录状态 = 1 And 合作单位 = v_合作单位 And 发生时间 Between Trunc(d_日期) And
                    Trunc(d_日期 + 1) - 1 / 60 / 60 / 24;
            Exception
              When Others Then
                n_合约已挂数 := 0;
            End;
            If Nvl(n_合约总数量, 0) = 0 Then
              v_剩余数量 := '0';
            Else
              v_剩余数量 := n_合约总数量 - n_合约已挂数;
            End If;
          End If;
        End If;
        Select To_Char(开始时间, 'hh24:mi') Into v_Timetemp From 时间段 Where 时间段 = r_排班;
        v_时间段 := v_Timetemp || '-';
        Select To_Char(终止时间, 'hh24:mi') Into v_Timetemp From 时间段 Where 时间段 = r_排班;
        v_时间段 := v_时间段 || v_Timetemp;
        If v_Temp Is Not Null Then
          v_Temp := v_Temp || '</SPANLIST>';
        End If;
        If v_合作单位 Is Not Null Then
          If Nvl(r_计划id, 0) <> 0 Then
            Begin
              Select 1
              Into n_禁用
              From 合作单位计划控制
              Where 计划id = r_计划id And 合作单位 = v_合作单位 And 数量 = 0 And Rownum < 2;
            Exception
              When Others Then
                n_禁用 := 0;
            End;
          Else
            Begin
              Select 1
              Into n_禁用
              From 合作单位安排控制
              Where 安排id = r_安排id And 合作单位 = v_合作单位 And 数量 = 0 And Rownum < 2;
            Exception
              When Others Then
                n_禁用 := 0;
            End;
          End If;
        End If;
        --限约数=0的预约禁止
        If Trunc(d_日期) <> Trunc(Sysdate) Then
          If r_限约数 = 0 Then
            n_禁用 := 1;
          End If;
        End If;
        If Nvl(n_禁用, 0) = 0 Then
          --从项金额计算
          n_合计金额 := r_价格;
          For r_Subfee In (Select 现价, 从项数次
                           From 收费从属项目 A, 收费价目 B
                           Where a.主项id = r_项目id And a.从项id = b.收费细目id And Sysdate Between b.执行日期 And
                                 Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))) Loop
            n_合计金额 := n_合计金额 + r_Subfee.现价 * r_Subfee.从项数次;
          End Loop;
          If Trunc(Sysdate) = Trunc(d_日期) Then
            Begin
              Select 1
              Into n_Exists
              From (Select 时间段
                     From 时间段
                     Where ('3000-01-10 ' || To_Char(Sysdate, 'HH24:MI:SS') <
                           '3000-01-10 ' || To_Char(终止时间, 'HH24:MI:SS')) Or
                           ('3000-01-10 ' || To_Char(Sysdate, 'HH24:MI:SS') <
                           Decode(Sign(开始时间 - 终止时间), 1, '3000-01-11 ' || To_Char(终止时间, 'HH24:MI:SS'),
                                   '3000-01-10 ' || To_Char(终止时间, 'HH24:MI:SS'))))
              Where 时间段 = r_排班;
            Exception
              When Others Then
                n_Exists := 0;
            End;
          Else
            n_Exists := 1;
          End If;
          If n_Exists = 1 Then
            If v_剩余数量 > 0 Then
              c_Xmlmain := '<HB>' || '<HM>' || r_号码 || '</HM>' || '<YSID>' || r_医生id || '</YSID>' || '<YS>' || r_医生姓名 ||
                           '</YS>' || '<KSID>' || r_科室id || '</KSID>' || '<KSMC>' || r_科室名称 || '</KSMC>' || '<ZC>' || r_职称 ||
                           '</ZC>' || '<XMID>' || r_项目id || '</XMID>' || '<XMMC>' || r_项目名称 || '</XMMC>' || '<YGHS>' ||
                           n_已挂数 || '</YGHS>' || '<SYHS>' || v_剩余数量 || '</SYHS>' || '<PRICE>' || n_合计金额 || '</PRICE>' ||
                           '<HCXH>' || n_缓冲序号 || '</HCXH>' || '<HL>' || r_号类 || '</HL>' || '<FSD>' || n_分时段 || '</FSD>' ||
                           '<HBTIME>' || v_时间段 || '</HBTIME>' || '<FWMC>' || r_排班 || '</FWMC>' || v_Temp || '</HB>';
              v_Xmlmain := v_Xmlmain || c_Xmlmain;
            Else
              c_Xmlmain := '<HB>' || '<HM>' || r_号码 || '</HM>' || '<YSID>' || r_医生id || '</YSID>' || '<YS>' || r_医生姓名 ||
                           '</YS>' || '<KSID>' || r_科室id || '</KSID>' || '<KSMC>' || r_科室名称 || '</KSMC>' || '<ZC>' || r_职称 ||
                           '</ZC>' || '<XMID>' || r_项目id || '</XMID>' || '<XMMC>' || r_项目名称 || '</XMMC>' || '<YGHS>' ||
                           n_已挂数 || '</YGHS>' || '<SYHS>' || v_剩余数量 || '</SYHS>' || '<PRICE>' || n_合计金额 || '</PRICE>' ||
                           '<HL>' || r_号类 || '</HL>' || '<FSD>' || n_分时段 || '</FSD>' || '<HBTIME>' || v_时间段 ||
                           '</HBTIME>' || '<FWMC>' || r_排班 || '</FWMC>' || '</HB>';
              v_Xmlmain := v_Xmlmain || c_Xmlmain;
            End If;
          End If;
        End If;
        n_合约剩余数量 := 0;
        n_合约总数量   := 0;
        n_时段数量     := 0;
        n_禁用         := 0;
        n_非合约       := 0;
      End Loop;
      Close r_No;
      v_Xmlmain := '<GROUP>' || '<RQ>' || To_Char(d_日期, 'yyyy-mm-dd') || '</RQ>' || '<HBLIST>' || v_Xmlmain ||
                   '</HBLIST>' || '</GROUP>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Xmlmain)) Into x_Templet From Dual;
    Else
      --出诊表排班模式
      n_合约剩余数量 := 0;
      v_Sql          := 'Select a.科室id, b.号类, c.名称 As 科室名称, a.医生姓名, a.医生id, d.专业技术职务 As 职称, b.号码, a.Id As 记录id, a.上班时段, a.项目id, e.名称 As 项目名称, ';
      v_Sql          := v_Sql ||
                        'a.是否序号控制 As 序号控制, a.限号数, Nvl(a.限约数,a.限号数) As 限约数, a.已挂数, a.已约数, a.其中已接收 As 已接收, a.是否分时段 As 分时段, a.开始时间, a.终止时间, a.预约控制   ';
      v_Sql          := v_Sql || 'From 临床出诊记录 A, 临床出诊号源 B, 部门表 C, 人员表 D, 收费项目目录 E ';
      v_Sql          := v_Sql ||
                        'Where a.出诊日期 = Trunc(:1) And a.号源id = b.Id And a.项目id = e.Id And a.医生id = d.Id(+) And b.科室id = c.Id And Nvl(a.是否锁定, 0) = 0 And ';
      v_Sql          := v_Sql || '      (a.开始时间 < Nvl(a.停诊开始时间, a.终止时间) Or a.终止时间 > Nvl(a.停诊终止时间, a.开始时间)) ';
      v_Sql          := v_Sql || '      And Nvl(a.是否发布,0) = 1 And a.开始时间 > To_Date( ' || Chr(39) || v_启用时间 || Chr(39) || ',' ||
                        Chr(39) || 'yyyy-mm-dd hh24:mi:ss' || Chr(39) || ')';
      n_Curcount     := 2;
      If Nvl(n_科室id, 0) <> 0 Then
        v_Sql      := v_Sql || 'And b.科室id = :' || n_Curcount || ' ';
        n_Curcount := n_Curcount + 1;
      End If;
      If Nvl(n_医生id, 0) <> 0 Then
        v_Sql      := v_Sql || 'And a.医生id = :' || n_Curcount || ' ';
        n_Curcount := n_Curcount + 1;
      End If;
      If Nvl(v_医生姓名, '_') <> '_' Then
        v_Sql      := v_Sql || 'And a.医生姓名 = :' || n_Curcount || ' ';
        n_Curcount := n_Curcount + 1;
      End If;
    
      If Nvl(n_科室id, 0) <> 0 And Nvl(n_医生id, 0) <> 0 And Nvl(v_医生姓名, '_') <> '_' Then
        Open r_No For v_Sql
          Using d_日期, n_科室id, n_医生id, v_医生姓名;
      End If;
      If Nvl(n_科室id, 0) <> 0 And Nvl(n_医生id, 0) = 0 And Nvl(v_医生姓名, '_') = '_' Then
        Open r_No For v_Sql
          Using d_日期, n_科室id;
      End If;
      If Nvl(n_科室id, 0) = 0 And Nvl(n_医生id, 0) <> 0 And Nvl(v_医生姓名, '_') = '_' Then
        Open r_No For v_Sql
          Using d_日期, n_医生id;
      End If;
      If Nvl(n_科室id, 0) = 0 And Nvl(n_医生id, 0) = 0 And Nvl(v_医生姓名, '_') <> '_' Then
        Open r_No For v_Sql
          Using d_日期, v_医生姓名;
      End If;
      If Nvl(n_科室id, 0) <> 0 And Nvl(n_医生id, 0) <> 0 And Nvl(v_医生姓名, '_') = '_' Then
        Open r_No For v_Sql
          Using d_日期, n_科室id, n_医生id;
      End If;
      If Nvl(n_科室id, 0) <> 0 And Nvl(n_医生id, 0) = 0 And Nvl(v_医生姓名, '_') <> '_' Then
        Open r_No For v_Sql
          Using d_日期, n_科室id, v_医生姓名;
      End If;
      If Nvl(n_科室id, 0) = 0 And Nvl(n_医生id, 0) <> 0 And Nvl(v_医生姓名, '_') <> '_' Then
        Open r_No For v_Sql
          Using d_日期, n_医生id, v_医生姓名;
      End If;
      Loop
        Fetch r_No
          Into r_科室id, r_号类, r_科室名称, r_医生姓名, r_医生id, r_职称, r_号码, r_安排id, r_排班, r_项目id, r_项目名称, r_序号控制, r_限号数, r_限约数,
               r_已挂数, r_已约数, r_已接收, r_分时段, r_开始时间, r_终止时间, r_预约控制;
        Exit When r_No%NotFound;
        If Trunc(d_日期) = Trunc(Sysdate) Then
          --当天挂号
          If v_合作单位 Is Null Then
            --未传入合作单位
            n_已挂数   := r_已挂数;
            v_剩余数量 := r_限号数 - Nvl(r_已挂数, 0);
            If r_分时段 = 1 And r_序号控制 = 1 Then
              --分时段
              v_Temp   := '<SPANLIST>';
              n_Exists := 0;
              For r_Time In (Select 序号, 开始时间, 终止时间, 挂号状态 From 临床出诊序号控制 Where 记录id = r_安排id) Loop
                v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                          To_Char(r_Time.终止时间, 'hh24:mi:ss') || '</SJD>';
                If Nvl(r_Time.挂号状态, 0) = 0 Then
                  v_Temp   := v_Temp || '<SL>1</SL></SPAN>';
                  n_Exists := n_Exists + 1;
                Else
                  v_Temp := v_Temp || '<SL>0</SL></SPAN>';
                End If;
              End Loop;
              If n_Exists < To_Number(v_剩余数量) Then
                Select Max(终止时间) Into d_加号时间 From 临床出诊序号控制 Where 记录id = r_安排id;
                v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_加号时间, 'hh24:mi:ss') || '-' || '</SJD><SL>' ||
                          v_剩余数量 - n_Exists || '</SL></SPAN>';
              End If;
              v_Temp := v_Temp || '</SPANLIST>';
            End If;
          Else
            --传入合作单位
            n_已挂数 := r_已挂数;
            Begin
              Select 控制方式
              Into n_合约模式
              From 临床出诊挂号控制记录
              Where 记录id = r_安排id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1 And Rownum < 2;
            Exception
              When Others Then
                n_合约模式 := 4;
            End;
            If n_合约模式 = 0 Then
              n_禁用 := 1;
            End If;
            If n_合约模式 = 1 Or n_合约模式 = 2 Then
              Select 数量
              Into n_合约总数量
              From 临床出诊挂号控制记录
              Where 记录id = r_安排id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1;
              If n_合约模式 = 1 Then
                n_合约总数量 := Round(r_限号数 * n_合约总数量 / 100);
              End If;
              Begin
                Select Count(1)
                Into n_合约已挂数
                From 病人挂号记录
                Where 号别 = r_号码 And 记录状态 = 1 And 合作单位 = v_合作单位 And 发生时间 Between Trunc(d_日期) And
                      Trunc(d_日期 + 1) - 1 / 60 / 60 / 24;
              Exception
                When Others Then
                  n_合约已挂数 := 0;
              End;
              n_合约剩余数量 := n_合约总数量 - n_合约已挂数;
              If r_限号数 - Nvl(r_已挂数, 0) < n_合约剩余数量 Then
                v_剩余数量 := r_限号数 - Nvl(r_已挂数, 0);
              Else
                v_剩余数量 := n_合约剩余数量;
              End If;
              If r_分时段 = 1 And r_序号控制 = 1 Then
                --分时段
                v_Temp   := '<SPANLIST>';
                n_Exists := 0;
                For r_Time In (Select 序号, 开始时间, 终止时间, 挂号状态 From 临床出诊序号控制 Where 记录id = r_安排id) Loop
                  v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                            To_Char(r_Time.终止时间, 'hh24:mi:ss') || '</SJD>';
                  If Nvl(r_Time.挂号状态, 0) = 0 Then
                    v_Temp   := v_Temp || '<SL>1</SL></SPAN>';
                    n_Exists := n_Exists + 1;
                  Else
                    v_Temp := v_Temp || '<SL>0</SL></SPAN>';
                  End If;
                End Loop;
                If n_Exists < To_Number(v_剩余数量) Then
                  Select Max(终止时间) Into d_加号时间 From 临床出诊序号控制 Where 记录id = r_安排id;
                  v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_加号时间, 'hh24:mi:ss') || '-' || '</SJD><SL>' ||
                            v_剩余数量 - n_Exists || '</SL></SPAN>';
                End If;
                v_Temp := v_Temp || '</SPANLIST>';
              End If;
            End If;
            If n_合约模式 = 3 Then
              If n_序号控制 = 0 Then
                n_已挂数   := r_已挂数;
                v_剩余数量 := r_限号数 - Nvl(r_已挂数, 0);
              Else
                n_已挂数   := 0;
                v_剩余数量 := 0;
                For r_合作 In (Select 序号
                             From 临床出诊挂号控制记录
                             Where 记录id = r_安排id And 类型 = 1 And 性质 = 1 And 名称 = v_合作单位) Loop
                  Begin
                    Select 1, 开始时间, 终止时间
                    Into n_Exists, d_开始时间, d_终止时间
                    From 临床出诊序号控制
                    Where 序号 = r_合作.序号 And Nvl(挂号状态, 0) = 0;
                  Exception
                    When Others Then
                      n_Exists := 0;
                  End;
                  If n_Exists = 1 Then
                    v_剩余数量 := v_剩余数量 + 1;
                    If r_分时段 = 1 Then
                      v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_开始时间, 'hh24:mi:ss') || '-' ||
                                To_Char(d_终止时间, 'hh24:mi:ss') || '</SJD><SL>1</SL></SPAN>';
                    End If;
                  Else
                    n_已挂数 := n_已挂数 + 1;
                  End If;
                End Loop;
                If v_Temp Is Not Null Then
                  v_Temp := '<SPANLIST>' || v_Temp || '</SPANLIST>';
                End If;
              End If;
            End If;
            If n_合约模式 = 4 Then
              n_已挂数   := r_已挂数;
              v_剩余数量 := r_限号数 - Nvl(r_已挂数, 0);
              If r_分时段 = 1 And r_序号控制 = 1 Then
                --分时段
                v_Temp   := '<SPANLIST>';
                n_Exists := 0;
                For r_Time In (Select 序号, 开始时间, 终止时间, 挂号状态 From 临床出诊序号控制 Where 记录id = r_安排id) Loop
                  v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                            To_Char(r_Time.终止时间, 'hh24:mi:ss') || '</SJD>';
                  If Nvl(r_Time.挂号状态, 0) = 0 Then
                    v_Temp   := v_Temp || '<SL>1</SL></SPAN>';
                    n_Exists := n_Exists + 1;
                  Else
                    v_Temp := v_Temp || '<SL>0</SL></SPAN>';
                  End If;
                End Loop;
                If n_Exists < To_Number(v_剩余数量) Then
                  Select Max(终止时间) Into d_加号时间 From 临床出诊序号控制 Where 记录id = r_安排id;
                  v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_加号时间, 'hh24:mi:ss') || '-' || '</SJD><SL>' ||
                            v_剩余数量 - n_Exists || '</SL></SPAN>';
                End If;
                v_Temp := v_Temp || '</SPANLIST>';
              End If;
            End If;
          End If;
        Else
          --预约挂号
          If r_预约控制 = 1 Then
            n_禁用 := 1;
          Else
            --不限制预约
            If v_合作单位 Is Null Then
              If r_分时段 = 0 Then
                n_已挂数   := r_已约数;
                v_剩余数量 := Nvl(r_限约数, r_限号数) - Nvl(r_已约数, 0);
              Else
                --分时段
                n_已挂数   := 0;
                v_剩余数量 := 0;
                v_Temp     := '<SPANLIST>';
                If r_序号控制 = 0 Then
                  --非序号控制分时段预约
                  For r_Time In (Select 序号, 开始时间, 终止时间, 挂号状态, 数量
                                 From 临床出诊序号控制
                                 Where 记录id = r_安排id And 预约顺序号 Is Null And 是否预约 = 1) Loop
                    Select Count(1)
                    Into n_时段已挂
                    From 临床出诊序号控制
                    Where 预约顺序号 Is Not Null And Nvl(挂号状态, 0) <> 0;
                    v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.终止时间, 'hh24:mi:ss') || '</SJD>';
                    v_Temp     := v_Temp || '<SL>' || r_Time.数量 - n_时段已挂 || '</SL></SPAN>';
                    n_已挂数   := n_已挂数 + n_时段已挂;
                    v_剩余数量 := v_剩余数量 + (r_Time.数量 - n_时段已挂);
                  End Loop;
                Else
                  For r_Time In (Select 序号, 开始时间, 终止时间, 挂号状态 From 临床出诊序号控制 Where 记录id = r_安排id) Loop
                    v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                              To_Char(r_Time.终止时间, 'hh24:mi:ss') || '</SJD>';
                    If Nvl(r_Time.挂号状态, 0) = 0 Then
                      v_Temp     := v_Temp || '<SL>1</SL></SPAN>';
                      v_剩余数量 := v_剩余数量 + 1;
                    Else
                      v_Temp   := v_Temp || '<SL>0</SL></SPAN>';
                      n_已挂数 := n_已挂数 + 1;
                    End If;
                  End Loop;
                End If;
                v_Temp := v_Temp || '</SPANLIST>';
              End If;
            Else
              --合作单位预约挂号
              If r_预约控制 = 2 Then
                n_禁用 := 1;
              Else
                Begin
                  Select 控制方式
                  Into n_合约模式
                  From 临床出诊挂号控制记录
                  Where 记录id = r_安排id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1 And Rownum < 2;
                Exception
                  When Others Then
                    n_合约模式 := 4;
                End;
                If n_合约模式 = 0 Then
                  n_禁用 := 1;
                End If;
                If n_合约模式 = 1 Or n_合约模式 = 2 Then
                  Select 数量
                  Into n_合约总数量
                  From 临床出诊挂号控制记录
                  Where 记录id = r_安排id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1;
                  If n_合约模式 = 1 Then
                    n_合约总数量 := Round(r_限号数 * n_合约总数量 / 100);
                  End If;
                  Begin
                    Select Count(1)
                    Into n_合约已挂数
                    From 病人挂号记录
                    Where 号别 = r_号码 And 记录状态 = 1 And 合作单位 = v_合作单位 And 发生时间 Between Trunc(d_日期) And
                          Trunc(d_日期 + 1) - 1 / 60 / 60 / 24;
                  Exception
                    When Others Then
                      n_合约已挂数 := 0;
                  End;
                  n_合约剩余数量 := n_合约总数量 - n_合约已挂数;
                  If Nvl(r_限约数, r_限号数) - Nvl(r_已约数, 0) < n_合约剩余数量 Then
                    v_剩余数量 := Nvl(r_限约数, r_限号数) - Nvl(r_已约数, 0);
                  Else
                    v_剩余数量 := n_合约剩余数量;
                  End If;
                  If r_分时段 = 1 Then
                    v_Temp := '<SPANLIST>';
                    If r_序号控制 = 1 Then
                      --分时段,序号控制
                      n_Exists := 0;
                      For r_Time In (Select 序号, 开始时间, 终止时间, 挂号状态
                                     From 临床出诊序号控制
                                     Where 记录id = r_安排id) Loop
                        v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.终止时间, 'hh24:mi:ss') || '</SJD>';
                        If Nvl(r_Time.挂号状态, 0) = 0 Then
                          v_Temp   := v_Temp || '<SL>1</SL></SPAN>';
                          n_Exists := n_Exists + 1;
                        Else
                          v_Temp := v_Temp || '<SL>0</SL></SPAN>';
                        End If;
                      End Loop;
                      If n_Exists < To_Number(v_剩余数量) Then
                        v_剩余数量 := n_Exists;
                      End If;
                    Else
                      --分时段,非序号控制
                      n_Exists := 0;
                      For r_Time In (Select 序号, 开始时间, 终止时间, 挂号状态, 数量
                                     From 临床出诊序号控制
                                     Where 记录id = r_安排id And 预约顺序号 Is Null And 是否预约 = 1) Loop
                        Select Count(1)
                        Into n_时段已挂
                        From 临床出诊序号控制
                        Where 预约顺序号 Is Not Null And Nvl(挂号状态, 0) <> 0;
                        v_Temp   := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                    To_Char(r_Time.终止时间, 'hh24:mi:ss') || '</SJD>';
                        v_Temp   := v_Temp || '<SL>' || r_Time.数量 - n_时段已挂 || '</SL></SPAN>';
                        n_Exists := n_Exists + (r_Time.数量 - n_时段已挂);
                      End Loop;
                      If n_Exists < To_Number(v_剩余数量) Then
                        v_剩余数量 := n_Exists;
                      End If;
                    End If;
                    v_Temp := v_Temp || '</SPANLIST>';
                  End If;
                  n_已挂数 := r_已约数;
                End If;
                If n_合约模式 = 3 Then
                  If r_分时段 = 0 Then
                    If r_序号控制 = 0 Then
                      n_禁用 := 1;
                    Else
                      n_已挂数   := 0;
                      v_剩余数量 := 0;
                      For r_合作 In (Select 序号
                                   From 临床出诊挂号控制记录
                                   Where 记录id = r_安排id And 类型 = 1 And 性质 = 1 And 名称 = v_合作单位) Loop
                        Begin
                          Select 1
                          Into n_Exists
                          From 临床出诊序号控制
                          Where 记录id = r_安排id And 序号 = r_合作.序号 And 是否预约 = 1 And Nvl(挂号状态, 0) = 0;
                        Exception
                          When Others Then
                            n_Exists := 0;
                        End;
                        If n_Exists = 1 Then
                          v_剩余数量 := v_剩余数量 + 1;
                        Else
                          n_已挂数 := n_已挂数 + 1;
                        End If;
                      End Loop;
                      If Nvl(r_限约数, r_限号数) - Nvl(r_已约数, 0) < v_剩余数量 Then
                        v_剩余数量 := Nvl(r_限约数, r_限号数) - Nvl(r_已约数, 0);
                      End If;
                    End If;
                  Else
                    If r_序号控制 = 0 Then
                      --合作单位,分时段,非序号控制
                      n_已挂数   := 0;
                      v_剩余数量 := 0;
                      v_Temp     := '<SPANLIST>';
                      For r_合作 In (Select 序号, 数量
                                   From 临床出诊挂号控制记录
                                   Where 记录id = r_安排id And 类型 = 1 And 性质 = 1 And 名称 = v_合作单位) Loop
                        Select Count(1), Max(开始时间), Max(终止时间)
                        Into n_时段已挂, d_开始时间, d_终止时间
                        From 临床出诊序号控制
                        Where 记录id = r_安排id And 预约顺序号 Is Not Null And Nvl(挂号状态, 0) <> 0 And 序号 = r_合作.序号;
                        n_已挂数   := n_已挂数 + n_Exists;
                        v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_开始时间, 'hh24:mi:ss') || '-' ||
                                      To_Char(d_终止时间, 'hh24:mi:ss') || '</SJD>';
                        v_Temp     := v_Temp || '<SL>' || r_合作.数量 - n_时段已挂 || '</SL></SPAN>';
                        v_剩余数量 := v_剩余数量 + r_合作.数量 - n_时段已挂;
                      End Loop;
                      v_Temp := v_Temp || '</SPANLIST>';
                    Else
                      n_已挂数   := 0;
                      v_剩余数量 := 0;
                      v_Temp     := '<SPANLIST>';
                      For r_合作 In (Select 序号
                                   From 临床出诊挂号控制记录
                                   Where 记录id = r_安排id And 类型 = 1 And 性质 = 1 And 名称 = v_合作单位) Loop
                        Begin
                          Select 1, 开始时间, 终止时间
                          Into n_Exists, d_开始时间, d_终止时间
                          From 临床出诊序号控制
                          Where 记录id = r_安排id And Nvl(挂号状态, 0) = 0 And 序号 = r_合作.序号;
                        Exception
                          When Others Then
                            n_Exists := 0;
                        End;
                        If n_Exists = 1 Then
                          v_剩余数量 := v_剩余数量 + 1;
                          v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_开始时间, 'hh24:mi:ss') || '-' ||
                                        To_Char(d_终止时间, 'hh24:mi:ss') || '</SJD>';
                          v_Temp     := v_Temp || '<SL>' || 1 || '</SL></SPAN>';
                        Else
                          v_Temp   := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_开始时间, 'hh24:mi:ss') || '-' ||
                                      To_Char(d_终止时间, 'hh24:mi:ss') || '</SJD>';
                          v_Temp   := v_Temp || '<SL>' || 0 || '</SL></SPAN>';
                          n_已挂数 := n_已挂数 + 1;
                        End If;
                      End Loop;
                    End If;
                  End If;
                End If;
                If n_合约模式 = 4 Then
                  If r_分时段 = 0 Then
                    n_已挂数   := r_已约数;
                    v_剩余数量 := Nvl(r_限约数, r_限号数) - Nvl(r_已约数, 0);
                  Else
                    --分时段
                    n_已挂数   := 0;
                    v_剩余数量 := 0;
                    v_Temp     := '<SPANLIST>';
                    If r_序号控制 = 0 Then
                      --非序号控制分时段预约
                      For r_Time In (Select 序号, 开始时间, 终止时间, 挂号状态, 数量
                                     From 临床出诊序号控制
                                     Where 记录id = r_安排id And 预约顺序号 Is Null And 是否预约 = 1) Loop
                        Select Count(1)
                        Into n_时段已挂
                        From 临床出诊序号控制
                        Where 预约顺序号 Is Not Null And Nvl(挂号状态, 0) <> 0;
                        v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                      To_Char(r_Time.终止时间, 'hh24:mi:ss') || '</SJD>';
                        v_Temp     := v_Temp || '<SL>' || r_Time.数量 - n_时段已挂 || '</SL></SPAN>';
                        n_已挂数   := n_已挂数 + n_时段已挂;
                        v_剩余数量 := v_剩余数量 + (r_Time.数量 - n_时段已挂);
                      End Loop;
                    Else
                      For r_Time In (Select 序号, 开始时间, 终止时间, 挂号状态
                                     From 临床出诊序号控制
                                     Where 记录id = r_安排id) Loop
                        v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.终止时间, 'hh24:mi:ss') || '</SJD>';
                        If Nvl(r_Time.挂号状态, 0) = 0 Then
                          v_Temp     := v_Temp || '<SL>1</SL></SPAN>';
                          v_剩余数量 := v_剩余数量 + 1;
                        Else
                          v_Temp   := v_Temp || '<SL>0</SL></SPAN>';
                          n_已挂数 := n_已挂数 + 1;
                        End If;
                      End Loop;
                    End If;
                    v_Temp := v_Temp || '</SPANLIST>';
                  End If;
                End If;
              End If;
            End If;
          End If;
        End If;
      
        If Nvl(n_禁用, 0) = 0 Then
	  n_合计金额 := 0;
          For r_Fee In (Select b.现价, a.从项数次
                        From 收费从属项目 A, 收费价目 B
                        Where a.主项id = r_项目id And a.从项id = b.收费细目id And Sysdate Between b.执行日期 And
                              Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))
                        Union
                        Select b.现价, 1 As 从项数次
                        From 收费项目目录 A, 收费价目 B
                        Where a.Id = b.收费细目id And a.Id = r_项目id And Sysdate Between b.执行日期 And
                              Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD'))) Loop
            n_合计金额 := n_合计金额 + r_Fee.现价 * r_Fee.从项数次;
          End Loop;
          v_时间段  := To_Char(r_开始时间, 'HH24:MI') || '-' || To_Char(r_终止时间, 'HH24:MI');
          c_Xmlmain := '<HB>' || '<CZJLID>' || r_安排id || '</CZJLID>' || '<HM>' || r_号码 || '</HM>' || '<YSID>' || r_医生id ||
                       '</YSID>' || '<YS>' || r_医生姓名 || '</YS>' || '<KSID>' || r_科室id || '</KSID>' || '<KSMC>' ||
                       r_科室名称 || '</KSMC>' || '<ZC>' || r_职称 || '</ZC>' || '<XMID>' || r_项目id || '</XMID>' || '<XMMC>' ||
                       r_项目名称 || '</XMMC>' || '<YGHS>' || n_已挂数 || '</YGHS>' || '<SYHS>' || v_剩余数量 || '</SYHS>' ||
                       '<PRICE>' || n_合计金额 || '</PRICE>' || '<HL>' || r_号类 || '</HL>' || '<FSD>' || r_分时段 || '</FSD>' ||
                       '<HBTIME>' || v_时间段 || '</HBTIME>' || '<FWMC>' || r_排班 || '</FWMC>' || v_Temp || '</HB>';
          v_Xmlmain := v_Xmlmain || c_Xmlmain;
        End If;
        n_禁用 := 0;
      End Loop;
      Close r_No;
      v_Xmlmain := '<GROUP>' || '<RQ>' || To_Char(d_日期, 'yyyy-mm-dd') || '</RQ>' || '<HBLIST>' || v_Xmlmain ||
                   '</HBLIST>' || '</GROUP>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Xmlmain)) Into x_Templet From Dual;
    End If;
  End If;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getnolist;
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
  --   <JSLIST>
  --     <JS>            //结算信息，挂号目前仅支持一个，结构与收费一致，以后可扩展
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
  --   <HZDW>合作单位</HZDW>        //合作单位名称
  --   <YYFS>支付宝<YYFS>    //预约方式,如自助机，支付宝
  --   <BRID>病人ID</BRID>     //病人ID
  --   <BRLX></BRLX>             //医保病人类型
  --   <FB>普通</FB>               //病人费别，可以不传
  --   <JQM>机器名</JQM>            //机器名
  --</IN>

  --出参:Xml_Out
  --<OUTPUT>
  -- <GHDH>挂号单号</GHDH>          //挂号单号
  -- <CZSJ>操作时间</CZSJ>          //HIS的登记时间
  -- <ERROR><MSG>错误信息</MSG></ERROR>  //出错时返回
  --</ OUTPUT>
  --------------------------------------------------------------------------------------------------
  v_号码     挂号安排.号码%Type;
  d_发生时间 Date;
  d_原始时间 Date;
  d_登记时间 Date;
  v_金额     Varchar2(200);

  n_应收金额   门诊费用记录.应收金额%Type;
  v_流水号     病人预交记录.交易流水号%Type;
  v_说明       门诊费用记录.摘要%Type;
  n_病人id     病人信息.病人id%Type;
  v_预约方式   预约方式.名称%Type;
  v_卡类别名称 医疗卡类别.名称%Type;
  v_结算卡号   病人预交记录.卡号%Type;
  n_门诊号     门诊费用记录.标识号%Type;
  v_姓名       门诊费用记录.姓名%Type;
  v_性别       门诊费用记录.性别%Type;
  v_年龄       门诊费用记录.年龄%Type;
  v_付款方式   门诊费用记录.付款方式%Type;
  v_费别       门诊费用记录.费别%Type;
  v_No         病人挂号记录.No%Type;
  v_结算方式   医疗卡类别.结算方式%Type;
  v_收费类别   门诊费用记录.收费类别%Type;
  n_收费细目id 门诊费用记录.收费细目id%Type;
  n_标准单价   门诊费用记录.标准单价%Type;
  n_收入项目id 门诊费用记录.收入项目id%Type;
  n_屏蔽费别   收费项目目录.屏蔽费别%Type;
  v_收据费目   门诊费用记录.收据费目%Type;
  n_病人科室id 门诊费用记录.病人科室id%Type;
  n_开单部门id 门诊费用记录.开单部门id%Type;
  v_操作员编号 门诊费用记录.操作员编号%Type;
  v_操作员姓名 门诊费用记录.操作员姓名%Type;
  v_医生姓名   挂号安排.医生姓名%Type;
  n_医生id     挂号安排.医生id%Type;
  n_结帐id     门诊费用记录.结帐id%Type;
  n_卡类别id   医疗卡类别.Id%Type;
  v_排班       挂号安排.周日%Type;
  n_安排id     挂号安排.Id%Type;
  n_计划id     挂号安排计划.Id%Type;
  n_预交id     病人预交记录.Id%Type;
  n_序号控制   挂号安排.序号控制%Type;
  n_号序       挂号序号状态.序号%Type;
  v_星期       挂号安排限制.限制项目%Type;
  v_病人类型   病人信息.病人类型%Type;
  n_存在       Number(3);
  v_现金       结算方式.名称%Type;
  n_分时段     Number(3);
  v_结算内容   Varchar2(3000);
  v_合作单位   病人挂号记录.合作单位%Type;
  v_机器名     挂号序号状态.机器名%Type;
  n_缴款方式   Number(3);
  n_记录id     临床出诊记录.Id%Type;
  v_Temp       Varchar2(32767); --临时XML
  x_Templet    Xmltype; --模板XML
  v_Err_Msg    Varchar2(200);
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/HM'), Extractvalue(Value(A), 'IN/HX'),
         To_Date(Extractvalue(Value(A), 'IN/YYSJ'), 'YYYY-MM-DD hh24:mi:ss'), Extractvalue(Value(A), 'IN/JE'),
         Extractvalue(Value(A), 'IN/YYFS'), Extractvalue(Value(A), 'IN/HZDW'), Extractvalue(Value(A), 'IN/BRID'),
         Extractvalue(Value(A), 'IN/BRLX'), Extractvalue(Value(A), 'IN/FB'), Extractvalue(Value(A), 'IN/JQM'),
         Extractvalue(Value(A), 'IN/JKFS'), Extractvalue(Value(A), 'IN/CZJLID')
  Into v_号码, n_号序, d_原始时间, n_应收金额, v_预约方式, v_合作单位, n_病人id, v_病人类型, v_费别, v_机器名, n_缴款方式, n_记录id
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  d_登记时间 := Sysdate;
  d_发生时间 := Trunc(d_原始时间);
  If v_病人类型 Is Not Null Then
    Begin
      Select 1 Into n_存在 From 病人类型 Where 名称 = v_病人类型;
    Exception
      When Others Then
        v_Err_Msg := '没有发现为(' || v_病人类型 || ')的病人类型';
        Raise Err_Item;
    End;
    Update 病人信息 Set 病人类型 = Nvl(病人类型, v_病人类型) Where 病人id = n_病人id;
  End If;

  Select a.门诊号, a.姓名, a.性别, a.年龄, Nvl(b.编码, c.编码)
  Into n_门诊号, v_姓名, v_性别, v_年龄, v_付款方式
  From 病人信息 A, 医疗付款方式 B, (Select 编码 From 医疗付款方式 Where 缺省标志 = '1' And Rownum < 2) C
  Where a.病人id = n_病人id And a.医疗付款方式 = b.名称(+);
  v_No   := Nextno(12);
  v_Temp := Zl_Identity(1);
  Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_Temp From Dual;
  Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into v_操作员编号 From Dual;
  Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_操作员姓名 From Dual;
  v_Temp := Zl_Identity(2);
  Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into n_开单部门id From Dual;

  If n_记录id Is Null Then
    Select Extractvalue(b.Column_Value, '/JS/JSKLB'), Extractvalue(b.Column_Value, '/JS/JSKH'),
           Extractvalue(b.Column_Value, '/JS/JSFS'), Extractvalue(b.Column_Value, '/JS/JYLSH'),
           Extractvalue(b.Column_Value, '/JS/JYSM')
    Into v_卡类别名称, v_结算卡号, v_结算方式, v_流水号, v_说明
    From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B;
  
    Begin
      Select b.结算方式, b.Id Into v_结算方式, n_卡类别id From 医疗卡类别 B Where b.名称 = v_卡类别名称 And Rownum < 2;
    Exception
      When Others Then
        v_Err_Msg := '没有发现该结算卡的相关信息';
        Raise Err_Item;
    End;
    Select Decode(To_Char(d_原始时间, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六',
                   Null)
    Into v_星期
    From Dual;
    Begin
      Select ID
      Into n_计划id
      From (Select ID
             From 挂号安排计划
             Where 号码 = v_号码 And d_原始时间 Between Nvl(生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                   Nvl(失效时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And 审核时间 Is Not Null
             Order By 生效时间 Desc)
      Where Rownum < 2;
    Exception
      When Others Then
        Select ID Into n_安排id From 挂号安排 Where 号码 = v_号码;
    End;
    If Nvl(n_计划id, 0) <> 0 Then
      --从计划读取信息
      Select a.项目id, b.科室id, a.医生姓名, a.医生id,
             Decode(To_Char(d_发生时间, 'D'), '1', a.周日, '2', a.周一, '3', a.周二, '4', a.周三, '5', a.周四, '6', a.周五, '7', a.周六,
                     Null), Nvl(a.序号控制, 0)
      Into n_收费细目id, n_病人科室id, v_医生姓名, n_医生id, v_排班, n_序号控制
      From 挂号安排计划 A, 挂号安排 B
      Where a.Id = n_计划id And b.Id = a.安排id;
      Select Count(Rownum) Into n_分时段 From 挂号计划时段 Where 星期 = v_星期 And 计划id = n_计划id And Rownum < 2;
      --合作单位检查
      If v_合作单位 Is Not Null Then
        Begin
          Select 1 Into n_存在 From 合作单位计划控制 Where 计划id = n_计划id And 数量 = 0 And 合作单位 = v_合作单位;
        Exception
          When Others Then
            n_存在 := 0;
        End;
      End If;
      If n_存在 = 1 Then
        v_Err_Msg := '传入的合作单位在此号码上被禁用！';
        Raise Err_Item;
      End If;
      If n_分时段 = 1 And n_序号控制 = 0 Then
        d_发生时间 := d_原始时间;
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
              Select To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || '' || To_Char(开始时间, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
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
      Select b.项目id, b.科室id, b.医生姓名, b.医生id,
             Decode(To_Char(d_发生时间, 'D'), '1', b.周日, '2', b.周一, '3', b.周二, '4', b.周三, '5', b.周四, '6', b.周五, '7', b.周六,
                     Null), Nvl(b.序号控制, 0)
      Into n_收费细目id, n_病人科室id, v_医生姓名, n_医生id, v_排班, n_序号控制
      From 挂号安排 B
      Where b.Id = n_安排id;
      Select Count(Rownum) Into n_分时段 From 挂号安排时段 Where 星期 = v_星期 And 安排id = n_安排id And Rownum < 2;
      --合作单位检查
      If v_合作单位 Is Not Null Then
        Begin
          Select 1 Into n_存在 From 合作单位安排控制 Where 安排id = n_安排id And 数量 = 0 And 合作单位 = v_合作单位;
        Exception
          When Others Then
            n_存在 := 0;
        End;
      End If;
      If n_存在 = 1 Then
        v_Err_Msg := '传入的合作单位在此号码上被禁用！';
        Raise Err_Item;
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
              Select To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || '' || To_Char(开始时间, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
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
  
    Select a.类别, b.现价, b.收入项目id, c.收据费目, a.屏蔽费别
    Into v_收费类别, n_标准单价, n_收入项目id, v_收据费目, n_屏蔽费别
    From 收费项目目录 A, 收费价目 B, 收入项目 C
    Where a.Id = n_收费细目id And b.收费细目id = a.Id And b.收入项目id = c.Id And Sysdate Between b.执行日期 And
          Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And Rownum < 2;
  
    Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
    Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
  
    If Trunc(d_发生时间) <> Trunc(Sysdate) Then
      If Nvl(n_缴款方式, 0) = 0 Then
        Zl_三方机构挂号_Insert(3, n_病人id, v_号码, n_号序, v_No, Null, v_结算方式, Null, d_发生时间, d_登记时间, v_合作单位, n_应收金额, Null, Null,
                         v_流水号, v_说明, v_预约方式, n_预交id, n_卡类别id, Null, 1, n_结帐id, 0, Null, Null, v_结算卡号, 1, v_费别, Null,
                         v_机器名, 1);
      Else
        Zl_三方机构挂号_Insert(2, n_病人id, v_号码, n_号序, v_No, Null, v_结算方式, Null, d_发生时间, d_登记时间, v_合作单位, n_应收金额, Null, Null,
                         v_流水号, v_说明, v_预约方式, n_预交id, n_卡类别id, Null, 1, n_结帐id, 0, Null, Null, v_结算卡号, 1, v_费别, Null,
                         v_机器名, 1);
      End If;
    Else
      Zl_三方机构挂号_Insert(1, n_病人id, v_号码, n_号序, v_No, Null, v_结算方式, Null, d_发生时间, d_登记时间, v_合作单位, n_应收金额, Null, Null,
                       v_流水号, v_说明, v_预约方式, n_预交id, n_卡类别id, Null, 1, n_结帐id, 0, Null, Null, v_结算卡号, 1, v_费别, Null,
                       v_机器名, 1);
    End If;
  
    For c_扩展信息 In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As 交易名称,
                          Extractvalue(b.Column_Value, '/EXPEND/JYLR') As 交易内容
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS/EXPENDLIST/EXPEND'))) B) Loop
      Zl_三方结算交易_Insert(n_卡类别id, 0, v_结算卡号, n_结帐id, c_扩展信息.交易名称 || '|' || c_扩展信息.交易内容, 0);
    End Loop;
  
    v_Temp := '<GHDH>' || v_No || '</GHDH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<CZSJ>' || To_Char(d_登记时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Else
    --出诊表排班模式
    For r_结算 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                        Extractvalue(b.Column_Value, '/JS/JYSM') As 交易说明,
                        Extractvalue(b.Column_Value, '/JS/JSJE') As 结算金额
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
      v_结算内容 := v_结算内容 || '|' || r_结算.结算方式 || ',' || r_结算.结算金额 || ',,';
      If r_结算.结算卡类别 Is Not Null Then
        v_结算内容   := v_结算内容 || '1';
        v_卡类别名称 := r_结算.结算卡类别;
        v_结算卡号   := r_结算.结算卡号;
        v_流水号     := r_结算.交易流水号;
        v_说明       := r_结算.交易说明;
      Else
        v_结算内容 := v_结算内容 || '0';
      End If;
    End Loop;
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
    Select 项目id, 科室id, 医生姓名, 医生id, 是否序号控制, 是否分时段
    Into n_收费细目id, n_病人科室id, v_医生姓名, n_医生id, n_序号控制, n_分时段
    From 临床出诊记录
    Where ID = n_记录id;
  
    Begin
      Select 开始时间 Into d_发生时间 From 临床出诊序号控制 Where 记录id = n_记录id And 序号 = n_号序;
    Exception
      When Others Then
        d_发生时间 := d_原始时间;
    End;
  
    Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
    Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
  
    If Trunc(d_发生时间) <> Trunc(Sysdate) Then
      If Nvl(n_缴款方式, 0) = 0 Then
        Zl_三方机构挂号_Insert(3, n_病人id, v_号码, n_号序, v_No, Null, v_结算内容, Null, d_发生时间, d_登记时间, v_合作单位, n_应收金额,
                           Null, Null, v_流水号, v_说明, v_预约方式, n_预交id, n_卡类别id, Null, 1, n_结帐id, 0, Null, Null, v_结算卡号, 1,
                           v_费别, Null, v_机器名, 1, n_记录id);
      Else
        Zl_三方机构挂号_Insert(2, n_病人id, v_号码, n_号序, v_No, Null, v_结算内容, Null, d_发生时间, d_登记时间, v_合作单位, n_应收金额,
                           Null, Null, v_流水号, v_说明, v_预约方式, n_预交id, n_卡类别id, Null, 1, n_结帐id, 0, Null, Null, v_结算卡号, 1,
                           v_费别, Null, v_机器名, 1, n_记录id);
      End If;
    Else
      Zl_三方机构挂号_Insert(1, n_病人id, v_号码, n_号序, v_No, Null, v_结算方式, Null, d_发生时间, d_登记时间, v_合作单位, n_应收金额, Null,
                         Null, v_流水号, v_说明, v_预约方式, n_预交id, n_卡类别id, Null, 1, n_结帐id, 0, Null, Null, v_结算卡号, 1, v_费别,
                         Null, v_机器名, 1, n_记录id);
    End If;
  
    For c_扩展信息 In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As 交易名称,
                          Extractvalue(b.Column_Value, '/EXPEND/JYLR') As 交易内容
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS/EXPENDLIST/EXPEND'))) B) Loop
      Zl_三方结算交易_Insert(n_卡类别id, 0, v_结算卡号, n_结帐id, c_扩展信息.交易名称 || '|' || c_扩展信息.交易内容, 0);
    End Loop;
  
    v_Temp := '<GHDH>' || v_No || '</GHDH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<CZSJ>' || To_Char(d_登记时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  End If;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Regist;
/

Create Or Replace Procedure Zl_三方机构挂号_Delete
(
  单据号_In     门诊费用记录.No%Type,
  交易流水号_In 病人预交记录.交易流水号%Type,
  交易说明_In   病人预交记录.交易说明%Type,
  退号时间_In   门诊费用记录.登记时间%Type := Null,
  预交id_In     病人预交记录.Id%Type := Null
) As
  v_Error Varchar(255);
  Err_Custom Exception;

  --该游标用于判断是否单独收病历费,及挂号汇总表处理
  Cursor c_Registinfo
  (
    v_状态     病人挂号记录.记录状态%Type,
    v_性质     病人挂号记录.记录性质%Type,
    v_无效单据 Number := 0
  ) Is
    Select a.发生时间, a.登记时间, b.项目id, b.科室id, b.医生姓名, b.医生id, b.号码
    From 病人挂号记录 A, 挂号安排 B
    Where a.记录性质 = Decode(v_无效单据, 0, v_性质, a.记录性质) And a.记录状态 = v_状态 And a.No = 单据号_In And a.号别 = b.号码 And Rownum = 1;

  r_Registrow c_Registinfo%RowType;

  --该光标用于处理人员缴款余额中退的不同结算方式的金额
  Cursor c_Opermoney Is
    Select Distinct b.结算方式, b.冲预交
    From 门诊费用记录 A, 病人预交记录 B
    Where a.结帐id = b.结帐id And a.No = 单据号_In And a.记录性质 = 4 And a.记录状态 = 3 And b.记录性质 = 4 And b.记录状态 = 3 And
          Nvl(b.冲预交, 0) <> 0;

  n_执行状态       病人挂号记录.执行状态%Type;
  n_打印id         票据打印内容.Id%Type;
  n_结帐id         门诊费用记录.结帐id%Type;
  n_原结帐id       病人预交记录.结帐id%Type;
  n_病人id         病人信息.病人id%Type;
  n_返回值         病人余额.预交余额%Type;
  n_分诊台签到排队 Number;
  n_预交id         病人预交记录.Id%Type;
  n_预约挂号       Number;
  n_无效单据       Number; --无效单据没有产生费用单据
  n_挂号生成队列   Number;
  n_Count          Number;
  n_组id           财务缴款分组.Id%Type;
  d_退号时间       Date;
  v_操作员编号     人员表.编号%Type;
  v_操作员姓名     人员表.姓名%Type;
  v_合作单位       合作单位挂号汇总.合作单位%Type;
  n_预约状态       病人挂号记录.预约%Type;
  v_Temp           Varchar2(100);
  d_登记时间       病人挂号记录.登记时间%Type;
  v_号别           病人挂号记录.号别%Type;
  n_号序           病人挂号记录.号序%Type;
  n_启用分时段     Number;
  d_预约时间       病人挂号记录.预约时间%Type;
  n_合作单位限制   Number(18);
  n_预约生成队列   Number;
  n_记录性质       Number;
  n_状态           Number;
  n_退号重用       Number(3);
  n_挂号排班模式   Number;
  n_挂号id         病人挂号记录.Id%Type;
  Function Zl_操作员
  (
    Type_In     Integer,
    Splitstr_In Varchar2
  ) Return Varchar2 As
    n_Step Number(18);
    v_Sub  Varchar2(1000);
    --Type_In:0-获取缺省部门ID;1-获取操作员编号;2-获取操作员姓名
    -- SplitStr:格式为:部门ID,部门名称;人员ID,人员编号,人员姓名(用Zl_Identity获取的)
  Begin
    If Type_In = 0 Then
      --缺省部门
      n_Step := Instr(Splitstr_In, ',');
      v_Sub  := Substr(Splitstr_In, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 1 Then
      --操作员编码
      n_Step := Instr(Splitstr_In, ';');
      v_Sub  := Substr(Splitstr_In, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 2 Then
      --操作员姓名
      n_Step := Instr(Splitstr_In, ';');
      v_Sub  := Substr(Splitstr_In, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      Return v_Sub;
    End If;
  End;

  Procedure Zl_三方机构挂号_出诊_Delete
  (
    单据号_In     门诊费用记录.No%Type,
    交易流水号_In 病人预交记录.交易流水号%Type,
    交易说明_In   病人预交记录.交易说明%Type,
    退号时间_In   门诊费用记录.登记时间%Type := Null,
    预交id_In     病人预交记录.Id%Type := Null
  ) As
    v_Error Varchar(255);
    Err_Custom Exception;
  
    --该游标用于判断是否单独收病历费,及挂号汇总表处理
    Cursor c_Registinfo
    (
      v_状态     病人挂号记录.记录状态%Type,
      v_性质     病人挂号记录.记录性质%Type,
      v_无效单据 Number := 0
    ) Is
      Select a.发生时间, a.登记时间, b.项目id, b.科室id, b.医生姓名, b.医生id, b.Id As 记录id, a.号别 As 号码
      From 病人挂号记录 A, 临床出诊记录 B
      Where a.记录性质 = Decode(v_无效单据, 0, v_性质, a.记录性质) And a.记录状态 = v_状态 And a.No = 单据号_In And a.出诊记录id = b.Id And
            Rownum < 2;
  
    r_Registrow c_Registinfo%RowType;
  
    --该光标用于处理人员缴款余额中退的不同结算方式的金额
    Cursor c_Opermoney Is
      Select Distinct b.结算方式, b.冲预交
      From 门诊费用记录 A, 病人预交记录 B
      Where a.结帐id = b.结帐id And a.No = 单据号_In And a.记录性质 = 4 And a.记录状态 = 3 And b.记录性质 = 4 And b.记录状态 = 3 And
            Nvl(b.冲预交, 0) <> 0;
  
    n_执行状态       病人挂号记录.执行状态%Type;
    n_打印id         票据打印内容.Id%Type;
    n_结帐id         门诊费用记录.结帐id%Type;
    n_原结帐id       病人预交记录.结帐id%Type;
    n_病人id         病人信息.病人id%Type;
    n_返回值         病人余额.预交余额%Type;
    n_分诊台签到排队 Number;
    n_预交id         病人预交记录.Id%Type;
    n_预约挂号       Number;
    n_无效单据       Number; --无效单据没有产生费用单据
    n_挂号生成队列   Number;
    n_Count          Number;
    n_组id           财务缴款分组.Id%Type;
    d_退号时间       Date;
    v_操作员编号     人员表.编号%Type;
    v_操作员姓名     人员表.姓名%Type;
    v_合作单位       合作单位挂号汇总.合作单位%Type;
    n_预约状态       病人挂号记录.预约%Type;
    v_Temp           Varchar2(100);
    d_登记时间       病人挂号记录.登记时间%Type;
    v_号别           病人挂号记录.号别%Type;
    n_号序           病人挂号记录.号序%Type;
    n_启用分时段     Number;
    d_预约时间       病人挂号记录.预约时间%Type;
    n_合作单位限制   Number(18);
    n_预约生成队列   Number;
    n_记录性质       Number;
    n_状态           Number;
    n_退号重用       Number(3);
    n_挂号id         病人挂号记录.Id%Type;
    n_记录id         临床出诊记录.Id%Type;
    Function Zl_操作员
    (
      Type_In     Integer,
      Splitstr_In Varchar2
    ) Return Varchar2 As
      n_Step Number(18);
      v_Sub  Varchar2(1000);
      --Type_In:0-获取缺省部门ID;1-获取操作员编号;2-获取操作员姓名
      -- SplitStr:格式为:部门ID,部门名称;人员ID,人员编号,人员姓名(用Zl_Identity获取的)
    Begin
      If Type_In = 0 Then
        --缺省部门
        n_Step := Instr(Splitstr_In, ',');
        v_Sub  := Substr(Splitstr_In, 1, n_Step - 1);
        Return v_Sub;
      End If;
      If Type_In = 1 Then
        --操作员编码
        n_Step := Instr(Splitstr_In, ';');
        v_Sub  := Substr(Splitstr_In, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, 1, n_Step - 1);
        Return v_Sub;
      End If;
      If Type_In = 2 Then
        --操作员姓名
        n_Step := Instr(Splitstr_In, ';');
        v_Sub  := Substr(Splitstr_In, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        Return v_Sub;
      End If;
    End;
  Begin
    --部门ID,部门名称;人员ID,人员编号,人员姓名
    v_Temp := Zl_Identity(0);
    If Nvl(v_Temp, ' ') = ' ' Then
      v_Error := '当前操作人员未设置对应的人员关系,不能继续。';
      Raise Err_Custom;
    End If;
    v_操作员编号 := Zl_操作员(1, v_Temp);
    v_操作员姓名 := Zl_操作员(2, v_Temp);
  
    n_组id := Zl_Get组id(v_操作员姓名);
  
    d_退号时间 := 退号时间_In;
    If d_退号时间 Is Null Then
      d_退号时间 := Sysdate;
    End If;
  
    --首先判断要退号/取消预约的记录是否存在
    Begin
      Select Decode(记录性质, 2, 1, 0), 记录性质, 登记时间, 号别, 号序, Nvl(预约时间, 发生时间), Nvl(合作单位, ''), Nvl(预约, 0),
             Decode(记录状态, 0, 1, 0), 出诊记录id
      Into n_预约挂号, n_记录性质, d_登记时间, v_号别, n_号序, d_预约时间, v_合作单位, n_预约状态, n_无效单据, n_记录id
      From 病人挂号记录
      Where NO = 单据号_In And 记录状态 In (0, 1) And Rownum < 2;
    Exception
      When Others Then
        n_预约挂号 := -1;
    End;
  
    If n_预约挂号 = -1 Then
      v_Error := '单据可能已经被退号或单据输入错误!';
      Raise Err_Custom;
    End If;
  
    Begin
      Select Nvl(是否分时段, 0) Into n_启用分时段 From 临床出诊记录 Where ID = n_记录id And Rownum < 2;
    Exception
      When Others Then
        n_启用分时段 := 0;
    End;
  
    --预约检查是否添加合作单位控制
    --如果设置了合作单位控制 则
    Select Count(0) Into n_合作单位限制 From 临床出诊挂号控制记录 Where 类型 = 1 And 性质 = 1 And Rownum < 2;
    --更新挂号序号状态
    n_退号重用 := Zl_To_Number(zl_GetSysParameter('已退序号允许挂号', 1111));
    If n_退号重用 = 0 Then
      Update 临床出诊序号控制 Set 挂号状态 = 4 Where 记录id = n_记录id And (序号 = n_号序 Or 备注 = n_号序);
    Else
      Update 临床出诊序号控制
      Set 挂号状态 = 0, 类型 = Null, 名称 = Null, 操作员姓名 = Null, 工作站名称 = Null
      Where 记录id = n_记录id And (序号 = n_号序 Or 备注 = n_号序);
    End If;
    If Nvl(n_预约挂号, 0) = 1 Or Nvl(n_无效单据, 0) = 1 Then
      If Nvl(n_无效单据, 0) = 0 Then
        --N天内不能取消预约号
        n_Count := Zl_To_Number(zl_GetSysParameter('N天内不能取消预约号', 1111));
        If n_Count <> 0 Then
          If Trunc(Sysdate - n_Count) < d_登记时间 Then
            v_Error := '不能退掉预约在' || To_Char(Trunc(Sysdate - n_Count), 'yyyy-mm-dd') || '以前的预约单!';
            Raise Err_Custom;
          End If;
        End If;
      End If;
    
      n_状态 := Case n_无效单据
                When 1 Then
                 0
                Else
                 1
              End;
      --减少已约数
      Open c_Registinfo(n_状态, 2, n_无效单据);
      Fetch c_Registinfo
        Into r_Registrow;
    
      Update 病人挂号汇总
      Set 已约数 = Nvl(已约数, 0) - n_预约状态, 已挂数 = Nvl(已挂数, 0) - Decode(n_预约状态, 0, 1, 0)
      Where 日期 = Trunc(r_Registrow.发生时间) And 科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And
            Nvl(医生姓名, '医生') = Nvl(r_Registrow.医生姓名, '医生') And Nvl(医生id, 0) = Nvl(r_Registrow.医生id, 0) And
            (号码 = r_Registrow.号码 Or 号码 Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into 病人挂号汇总
          (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已约数, 已挂数)
        Values
          (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
           Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, -n_预约状态, Decode(n_预约状态, 0, 1, 0));
      End If;
    
      Update 临床出诊记录
      Set 已约数 = Nvl(已约数, 0) - n_预约状态, 已挂数 = Nvl(已挂数, 0) - Decode(n_预约状态, 0, 1, 0)
      Where ID = n_记录id;
      Close c_Registinfo;
    
      If Nvl(n_无效单据, 0) = 0 Then
        --删除门诊费用记录
        Delete From 门诊费用记录 Where NO = 单据号_In And 记录性质 = 4 And 记录状态 = 0;
        --如果预约生成队列时需要清除队列
        n_挂号生成队列 := Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113));
        If Nvl(n_挂号生成队列, 0) = 1 Then
          n_预约生成队列 := Zl_To_Number(zl_GetSysParameter('预约生成队列', 1113));
          If Nvl(n_预约生成队列, 0) = 1 Then
            --要删除队列
            For v_挂号 In (Select ID, 诊室, 执行部门id, 执行人 From 病人挂号记录 Where NO = 单据号_In) Loop
              Zl_排队叫号队列_Delete(v_挂号.执行部门id, v_挂号.Id);
            End Loop;
          End If;
        End If;
      End If;
    Else
      Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
    
      --更新挂号序号状态
    
      --病人就诊状态
      Select 病人id
      Into n_病人id
      From 门诊费用记录
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And 序号 = 1;
    
      If n_病人id Is Not Null Then
        Update 病人信息 Set 就诊状态 = 0, 就诊诊室 = Null Where 病人id = n_病人id;
        --删除门诊号相关处理,只有当只有一条挂号记录并且病人建档日期与挂号日期近似时才会处理
      End If;
    
      --门诊费用记录
      Insert Into 门诊费用记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位,
         付数, 数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 执行人, 操作员编号, 操作员姓名, 发生时间,
         登记时间, 结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 缴款组id)
        Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别,
               收费细目id, 计算单位, 付数, -数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, -应收金额, -实收金额, 开单部门id, 开单人, 执行部门id, 执行人,
               v_操作员编号, v_操作员姓名, 发生时间, d_退号时间, n_结帐id, -1 * 结帐金额, 保险项目否, 保险大类id, -1 * 统筹金额, 摘要,
               Decode(Nvl(附加标志, 0), 9, 1, 0), 保险编码, 费用类型, n_组id
        From 门诊费用记录
        Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In;
    
      Update 门诊费用记录 Set 记录状态 = 3 Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In;
      Select 结帐id
      Into n_原结帐id
      From 门诊费用记录
      Where 记录性质 = 4 And 记录状态 = 3 And NO = 单据号_In And Rownum = 1;
    
      n_预交id := 预交id_In;
      If Nvl(预交id_In, 0) = 0 Then
        Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      End If;
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易流水号, 交易说明, 合作单位,
         结算序号, 卡类别id, 结算性质)
        Select n_预交id, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_退号时间, v_操作员编号, v_操作员姓名, -冲预交, n_结帐id, n_组id,
               交易流水号_In, 交易说明_In, 合作单位, n_结帐id, 卡类别id, 结算性质
        From 病人预交记录
        Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_原结帐id;
    
      Update 病人预交记录 Set 记录状态 = 3 Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_原结帐id;
    
      --退卡收回票据(可能上次挂号使用票据,不能收回)
      Begin
        --从最后一次的打印内容中取
        Select ID
        Into n_打印id
        From (Select b.Id
               From 票据使用明细 A, 票据打印内容 B
               Where a.打印id = b.Id And a.性质 = 1 And b.数据性质 = 4 And b.No = 单据号_In
               Order By a.使用时间 Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          Null;
      End;
      If n_打印id Is Not Null Then
        Insert Into 票据使用明细
          (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
          Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_退号时间, v_操作员姓名
          From 票据使用明细
          Where 打印id = n_打印id And 性质 = 1;
      End If;
    
      --相关汇总表的处理
    
      --病人挂号汇总
      Open c_Registinfo(1, 1);
      Fetch c_Registinfo
        Into r_Registrow;
    
      If c_Registinfo%RowCount = 0 Then
        --只收病历费时无号别,不处理
        Close c_Registinfo;
      Else
      
        --需要确定是否预约挂号
        --1.如果是退预约挂号产生的挂号记录,则需要减已挂数和其中已接数
        --2.如果是正常挂号,则只减已挂数
        Begin
          Select Decode(预约, Null, 0, 0, 0, 1), 执行状态
          Into n_预约挂号, n_执行状态
          From 病人挂号记录
          Where NO = 单据号_In And 记录状态 = 1 And Rownum = 1;
        Exception
          When Others Then
            n_预约挂号 := 0;
        End;
        --0-等待接诊,1-完成就诊,2-正在就诊,-1标记为不就诊
        If n_执行状态 > 0 Then
          If n_执行状态 = 1 Then
            v_Error := '该病人已经完成就诊,不能再退号!';
          Else
            v_Error := '该病人正在就诊, 不能退号!';
          End If;
          Raise Err_Custom;
        End If;
      
        Update 病人挂号汇总
        Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - n_预约状态, 已约数 = Nvl(已约数, 0) - n_预约状态
        Where 日期 = Trunc(r_Registrow.发生时间) And 科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And
              Nvl(医生姓名, '医生') = Nvl(r_Registrow.医生姓名, '医生') And Nvl(医生id, 0) = Nvl(r_Registrow.医生id, 0) And
              (号码 = r_Registrow.号码 Or 号码 Is Null);
      
        If Sql%RowCount = 0 Then
          Insert Into 病人挂号汇总
            (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 其中已接收)
          Values
            (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
             Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, -1, -1 * n_预约挂号);
        End If;
      
        Update 临床出诊记录
        Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - n_预约状态, 已约数 = Nvl(已约数, 0) - n_预约状态
        Where ID = n_记录id;
        Close c_Registinfo;
      End If;
    
      --人员缴款余额(包括个人帐户等的结算金额,不含退冲预交款)
      For r_Opermoney In c_Opermoney Loop
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) + (-1 * r_Opermoney.冲预交)
        Where 收款员 = v_操作员姓名 And 结算方式 = r_Opermoney.结算方式 And 性质 = 1
        Returning 余额 Into n_返回值;
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额
            (收款员, 结算方式, 性质, 余额)
          Values
            (v_操作员姓名, r_Opermoney.结算方式, 1, -1 * r_Opermoney.冲预交);
          n_返回值 := r_Opermoney.冲预交;
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 人员缴款余额
          Where 收款员 = v_操作员姓名 And 结算方式 = r_Opermoney.结算方式 And 性质 = 1 And Nvl(余额, 0) = 0;
        End If;
      End Loop;
    
      n_挂号生成队列 := Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113));
      If n_挂号生成队列 <> 0 Then
        n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113));
        If Nvl(n_分诊台签到排队, 0) = 0 Then
          --要删除队列
          For v_挂号 In (Select ID, 诊室, 执行部门id, 执行人 From 病人挂号记录 Where NO = 单据号_In) Loop
            Zl_排队叫号队列_Delete(v_挂号.执行部门id, v_挂号.Id);
          End Loop;
        End If;
      End If;
    
      --医保产生的就诊登记记录
      Delete From 就诊登记记录
      Where (病人id, 主页id, 就诊时间) In (Select 病人id, 主页id, 发生时间 From 病人挂号记录 Where NO = 单据号_In);
    End If;
  
    If Nvl(n_无效单据, 0) = 0 Then
      Update 病人挂号记录 Set 记录状态 = 3 Where NO = 单据号_In And 记录状态 = 1;
      If Sql%NotFound Then
        v_Error := '未找到挂号单据,请检查!';
        Raise Err_Custom;
      End If;
      Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
      Insert Into 病人挂号记录
        (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 操作员编号, 操作员姓名,
         复诊, 号序, 预约, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 出诊记录id)
        Select n_挂号id, NO, 记录性质, 2, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, d_退号时间, 发生时间,
               v_操作员编号, v_操作员姓名, 复诊, 号序, 预约, 接收人, 接收时间, 交易流水号_In, 交易说明_In, 合作单位, 出诊记录id
        From 病人挂号记录
        Where NO = 单据号_In And 记录状态 = 3;
    End If;
    --消息推送
    Begin
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 2, 单据号_In;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_003(n_挂号id, 单据号_In);
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Error || '+[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End;

Begin
  n_挂号排班模式 := Nvl(zl_GetSysParameter('挂号排班模式'), 0);
  If n_挂号排班模式 = 1 Then
    --出诊表排班模式
    Zl_三方机构挂号_出诊_Delete(单据号_In, 交易流水号_In, 交易说明_In, 退号时间_In, 预交id_In);
  Else
    --部门ID,部门名称;人员ID,人员编号,人员姓名
    v_Temp := Zl_Identity(0);
    If Nvl(v_Temp, ' ') = ' ' Then
      v_Error := '当前操作人员未设置对应的人员关系,不能继续。';
      Raise Err_Custom;
    End If;
    v_操作员编号 := Zl_操作员(1, v_Temp);
    v_操作员姓名 := Zl_操作员(2, v_Temp);
  
    n_组id := Zl_Get组id(v_操作员姓名);
  
    d_退号时间 := 退号时间_In;
    If d_退号时间 Is Null Then
      d_退号时间 := Sysdate;
    End If;
  
    --首先判断要退号/取消预约的记录是否存在
    Begin
      Select Decode(记录性质, 2, 1, 0), 记录性质, 登记时间, 号别, 号序, Nvl(预约时间, 发生时间), Nvl(合作单位, ''), Nvl(预约, 0),
             Decode(记录状态, 0, 1, 0)
      Into n_预约挂号, n_记录性质, d_登记时间, v_号别, n_号序, d_预约时间, v_合作单位, n_预约状态, n_无效单据
      From 病人挂号记录
      Where NO = 单据号_In And 记录状态 In (0, 1) And Rownum <= 1;
    Exception
      When Others Then
        n_预约挂号 := -1;
    End;
  
    If n_预约挂号 = -1 Then
      v_Error := '单据可能已经被退号或单据输入错误!';
      Raise Err_Custom;
    End If;
  
    Begin
      Select 1
      Into n_启用分时段
      From 挂号安排 A, 挂号安排时段 B
      Where a.号码 = v_号别 And a.Id = b.安排id And Rownum <= 1;
    Exception
      When Others Then
        n_启用分时段 := 0;
    End;
  
    --预约检查是否添加合作单位控制
    --如果设置了合作单位控制 则
    Select Count(0) Into n_合作单位限制 From 合作单位安排控制 Where Rownum = 1;
    --更新挂号序号状态
    n_退号重用 := Zl_To_Number(zl_GetSysParameter('已退序号允许挂号', 1111));
    If n_退号重用 = 0 Then
      Update 挂号序号状态
      Set 状态 = 4
      Where 号码 = v_号别 And 序号 = n_号序 And 日期 Between Trunc(d_预约时间) And Trunc(d_预约时间 + 1) - 1 / 24 / 60 / 60;
    Else
      Delete 挂号序号状态
      Where 号码 = v_号别 And 序号 = n_号序 And 日期 Between Trunc(d_预约时间) And Trunc(d_预约时间 + 1) - 1 / 24 / 60 / 60;
    End If;
    If Nvl(n_预约挂号, 0) = 1 Or Nvl(n_无效单据, 0) = 1 Then
      If Nvl(n_无效单据, 0) = 0 Then
        --N天内不能取消预约号
        n_Count := Zl_To_Number(zl_GetSysParameter('N天内不能取消预约号', 1111));
        If n_Count <> 0 Then
          If Trunc(Sysdate - n_Count) < d_登记时间 Then
            v_Error := '不能退掉预约在' || To_Char(Trunc(Sysdate - n_Count), 'yyyy-mm-dd') || '以前的预约单!';
            Raise Err_Custom;
          End If;
        End If;
      End If;
    
      n_状态 := Case n_无效单据
                When 1 Then
                 0
                Else
                 1
              End;
      --减少已约数
      Open c_Registinfo(n_状态, 2, n_无效单据);
      Fetch c_Registinfo
        Into r_Registrow;
    
      Update 病人挂号汇总
      Set 已约数 = Nvl(已约数, 0) - n_预约状态, 已挂数 = Nvl(已挂数, 0) - Decode(n_预约状态, 0, 1, 0)
      Where 日期 = Trunc(r_Registrow.发生时间) And 科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And
            Nvl(医生姓名, '医生') = Nvl(r_Registrow.医生姓名, '医生') And Nvl(医生id, 0) = Nvl(r_Registrow.医生id, 0) And
            (号码 = r_Registrow.号码 Or 号码 Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into 病人挂号汇总
          (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已约数, 已挂数)
        Values
          (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
           Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, -n_预约状态, Decode(n_预约状态, 0, 1, 0));
      End If;
    
      If Nvl(n_合作单位限制, 0) <> 0 And Nvl(v_合作单位, '') <> '' And Nvl(n_预约状态, 0) <> 0 Then
        Update 合作单位挂号汇总
        Set 已约数 = Nvl(已约数, 0) - n_预约状态, 已接数 = Nvl(已接数, 0) - Decode(n_预约状态, 0, 1, 0)
        Where 日期 = Trunc(r_Registrow.发生时间) And (号码 = r_Registrow.号码 Or 号码 Is Null) And 合作单位 = Nvl(v_合作单位, '') And
              序号 = Nvl(n_号序, 0);
        If Sql%RowCount = 0 Then
          Insert Into 合作单位挂号汇总
            (日期, 号码, 已约数, 合作单位, 序号, 已接数)
          Values
            (Trunc(r_Registrow.发生时间), r_Registrow.号码, -n_预约状态, v_合作单位, Nvl(n_号序, 0), -decode(n_预约状态, 0, 1, 0));
        End If;
      End If;
      Close c_Registinfo;
    
      If Nvl(n_无效单据, 0) = 0 Then
        --删除门诊费用记录
        Delete From 门诊费用记录 Where NO = 单据号_In And 记录性质 = 4 And 记录状态 = 0;
        --如果预约生成队列时需要清除队列
        n_挂号生成队列 := Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113));
        If Nvl(n_挂号生成队列, 0) = 1 Then
          n_预约生成队列 := Zl_To_Number(zl_GetSysParameter('预约生成队列', 1113));
          If Nvl(n_预约生成队列, 0) = 1 Then
            --要删除队列
            For v_挂号 In (Select ID, 诊室, 执行部门id, 执行人 From 病人挂号记录 Where NO = 单据号_In) Loop
              Zl_排队叫号队列_Delete(v_挂号.执行部门id, v_挂号.Id);
            End Loop;
          End If;
        End If;
      End If;
    Else
      Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
    
      --更新挂号序号状态
    
      --病人就诊状态
      Select 病人id
      Into n_病人id
      From 门诊费用记录
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And 序号 = 1;
    
      If n_病人id Is Not Null Then
        Update 病人信息 Set 就诊状态 = 0, 就诊诊室 = Null Where 病人id = n_病人id;
        --删除门诊号相关处理,只有当只有一条挂号记录并且病人建档日期与挂号日期近似时才会处理
      End If;
    
      --门诊费用记录
      Insert Into 门诊费用记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位,
         付数, 数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 执行人, 操作员编号, 操作员姓名, 发生时间,
         登记时间, 结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 缴款组id)
        Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别,
               收费细目id, 计算单位, 付数, -数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, -应收金额, -实收金额, 开单部门id, 开单人, 执行部门id, 执行人,
               v_操作员编号, v_操作员姓名, 发生时间, d_退号时间, n_结帐id, -1 * 结帐金额, 保险项目否, 保险大类id, -1 * 统筹金额, 摘要,
               Decode(Nvl(附加标志, 0), 9, 1, 0), 保险编码, 费用类型, n_组id
        From 门诊费用记录
        Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In;
    
      Update 门诊费用记录 Set 记录状态 = 3 Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In;
      Select 结帐id
      Into n_原结帐id
      From 门诊费用记录
      Where 记录性质 = 4 And 记录状态 = 3 And NO = 单据号_In And Rownum = 1;
    
      Select Count(Distinct 结算方式) Into n_Count From 病人预交记录 Where 结帐id = n_原结帐id;
      If n_Count > 1 Then
        v_Error := '不能处理多种结算方式,请检查传入的退号单据是否正确!';
        Raise Err_Custom;
      End If;
      n_预交id := 预交id_In;
      If Nvl(预交id_In, 0) = 0 Then
        Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      End If;
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易流水号, 交易说明, 合作单位,
         结算序号, 卡类别id, 结算性质)
        Select n_预交id, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_退号时间, v_操作员编号, v_操作员姓名, -冲预交, n_结帐id, n_组id,
               交易流水号_In, 交易说明_In, 合作单位, n_结帐id, 卡类别id, 结算性质
        From 病人预交记录
        Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_原结帐id;
    
      Update 病人预交记录 Set 记录状态 = 3 Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_原结帐id;
    
      --退卡收回票据(可能上次挂号使用票据,不能收回)
      Begin
        --从最后一次的打印内容中取
        Select ID
        Into n_打印id
        From (Select b.Id
               From 票据使用明细 A, 票据打印内容 B
               Where a.打印id = b.Id And a.性质 = 1 And b.数据性质 = 4 And b.No = 单据号_In
               Order By a.使用时间 Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          Null;
      End;
      If n_打印id Is Not Null Then
        Insert Into 票据使用明细
          (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
          Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_退号时间, v_操作员姓名
          From 票据使用明细
          Where 打印id = n_打印id And 性质 = 1;
      End If;
    
      --相关汇总表的处理
    
      --病人挂号汇总
      Open c_Registinfo(1, 1);
      Fetch c_Registinfo
        Into r_Registrow;
    
      If c_Registinfo%RowCount = 0 Then
        --只收病历费时无号别,不处理
        Close c_Registinfo;
      Else
      
        --需要确定是否预约挂号
        --1.如果是退预约挂号产生的挂号记录,则需要减已挂数和其中已接数
        --2.如果是正常挂号,则只减已挂数
        Begin
          Select Decode(预约, Null, 0, 0, 0, 1), 执行状态
          Into n_预约挂号, n_执行状态
          From 病人挂号记录
          Where NO = 单据号_In And 记录状态 = 1 And Rownum = 1;
        Exception
          When Others Then
            n_预约挂号 := 0;
        End;
        --0-等待接诊,1-完成就诊,2-正在就诊,-1标记为不就诊
        If n_执行状态 > 0 Then
          If n_执行状态 = 1 Then
            v_Error := '该病人已经完成就诊,不能再退号!';
          Else
            v_Error := '该病人正在就诊, 不能退号!';
          End If;
          Raise Err_Custom;
        End If;
      
        Update 病人挂号汇总
        Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - n_预约状态, 已约数 = Nvl(已约数, 0) - n_预约状态
        Where 日期 = Trunc(r_Registrow.发生时间) And 科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And
              Nvl(医生姓名, '医生') = Nvl(r_Registrow.医生姓名, '医生') And Nvl(医生id, 0) = Nvl(r_Registrow.医生id, 0) And
              (号码 = r_Registrow.号码 Or 号码 Is Null);
      
        If Sql%RowCount = 0 Then
          Insert Into 病人挂号汇总
            (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 其中已接收)
          Values
            (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
             Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, -1, -1 * n_预约挂号);
        End If;
        If Nvl(n_合作单位限制, 0) <> 0 And Nvl(v_合作单位, '') <> '' And Nvl(n_预约状态, 0) <> 0 Then
          Update 合作单位挂号汇总
          Set 已接数 = Nvl(已接数, 0) - 1, 已约数 = Nvl(已约数, 0) - n_预约挂号
          Where 日期 = Trunc(r_Registrow.发生时间) And (号码 = r_Registrow.号码 Or 号码 Is Null) And 合作单位 = Nvl(v_合作单位, '') And
                序号 = Nvl(n_号序, 0);
          If Sql%RowCount = 0 Then
            Insert Into 合作单位挂号汇总
              (日期, 号码, 已约数, 合作单位, 已接数)
            Values
              (Trunc(r_Registrow.发生时间), r_Registrow.号码, -1, v_合作单位, -1 * n_预约挂号);
          End If;
        End If;
        Close c_Registinfo;
      End If;
    
      --人员缴款余额(包括个人帐户等的结算金额,不含退冲预交款)
      For r_Opermoney In c_Opermoney Loop
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) + (-1 * r_Opermoney.冲预交)
        Where 收款员 = v_操作员姓名 And 结算方式 = r_Opermoney.结算方式 And 性质 = 1
        Returning 余额 Into n_返回值;
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额
            (收款员, 结算方式, 性质, 余额)
          Values
            (v_操作员姓名, r_Opermoney.结算方式, 1, -1 * r_Opermoney.冲预交);
          n_返回值 := r_Opermoney.冲预交;
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 人员缴款余额
          Where 收款员 = v_操作员姓名 And 结算方式 = r_Opermoney.结算方式 And 性质 = 1 And Nvl(余额, 0) = 0;
        End If;
      End Loop;
    
      n_挂号生成队列 := Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113));
      If n_挂号生成队列 <> 0 Then
        n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113));
        If Nvl(n_分诊台签到排队, 0) = 0 Then
          --要删除队列
          For v_挂号 In (Select ID, 诊室, 执行部门id, 执行人 From 病人挂号记录 Where NO = 单据号_In) Loop
            Zl_排队叫号队列_Delete(v_挂号.执行部门id, v_挂号.Id);
          End Loop;
        End If;
      End If;
    
      --医保产生的就诊登记记录
      Delete From 就诊登记记录
      Where (病人id, 主页id, 就诊时间) In (Select 病人id, 主页id, 发生时间 From 病人挂号记录 Where NO = 单据号_In);
    End If;
  
    If Nvl(n_无效单据, 0) = 0 Then
      Update 病人挂号记录 Set 记录状态 = 3 Where NO = 单据号_In And 记录状态 = 1;
      If Sql%NotFound Then
        v_Error := '未找到挂号单据,请检查!';
        Raise Err_Custom;
      End If;
      Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
      Insert Into 病人挂号记录
        (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 操作员编号, 操作员姓名,
         复诊, 号序, 预约, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位)
        Select n_挂号id, NO, 记录性质, 2, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, d_退号时间, 发生时间,
               v_操作员编号, v_操作员姓名, 复诊, 号序, 预约, 接收人, 接收时间, 交易流水号_In, 交易说明_In, 合作单位
        From 病人挂号记录
        Where NO = 单据号_In And 记录状态 = 3;
    End If;
    --消息推送
    Begin
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 2, 单据号_In;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_003(n_挂号id, 单据号_In);
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Error || '+[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_三方机构挂号_Delete;
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
  --  <LSH>34563</LSH>           //交易流水号
  --  <JKFS>0</JKFS>             //缴款方式,0-挂号或预约缴款;1-预约不缴款
  --  <YYFS></YYFS>              //缴款方式=1时传入，预约的预约方式
  --</IN>

  --出参:Xml_Out
  --<OUTPUT>
  -- <CZSJ>操作时间</CZSJ>          //HIS的登记时间
  -- <ERROR><MSG></MSG></ERROR> //为空表示取消挂号成功
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  v_卡类别     Varchar2(100);
  v_No         病人挂号记录.No%Type;
  n_挂号金额   门诊费用记录.实收金额%Type;
  v_操作员编号 门诊费用记录.操作员编号%Type;
  v_操作员姓名 门诊费用记录.操作员姓名%Type;
  v_结算方式   医疗卡类别.结算方式%Type;
  n_实收金额   门诊费用记录.实收金额%Type;
  v_交易流水号 病人预交记录.交易流水号%Type;
  n_存在       Number(3);
  v_Type       Varchar2(50);
  v_Temp       Varchar2(32767); --临时XML
  x_Templet    Xmltype; --模板XML
  v_Err_Msg    Varchar2(200);
  n_已开医嘱   Number(2);
  n_检查发票   Number(3);
  n_是否打印   Number(3);
  n_缴款方式   Number(3);
  d_登记时间   Date;
  n_挂号模式   Number(3);
  v_预约方式   病人挂号记录.预约方式%Type;
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/GHDH'), Extractvalue(Value(A), 'IN/JSKLB'), Extractvalue(Value(A), 'IN/GHJE'),
         Extractvalue(Value(A), 'IN/LSH'), To_Number(Extractvalue(Value(A), 'IN/JCFP')),
         To_Number(Extractvalue(Value(A), 'IN/JKFS')), Extractvalue(Value(A), 'IN/YYFS')
  Into v_No, v_卡类别, n_挂号金额, v_交易流水号, n_检查发票, n_缴款方式, v_预约方式
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

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
  
    Select Sum(实收金额) Into n_实收金额 From 门诊费用记录 Where NO = v_No And 记录性质 = 4;
  
    If Nvl(n_缴款方式, 0) = 0 Then
      --要退的单据不是以该结算卡结算的，则禁止退号
      Begin
        Select 1
        Into n_存在
        From 病人预交记录 A,
             (Select Distinct 结帐id
               From 门诊费用记录
               Where NO = v_No And 记录性质 = 4
               Union
               Select Distinct 结帐id From 住院费用记录 Where NO = v_No And 记录性质 = 5) B
        Where a.结帐id = b.结帐id And 结算方式 = v_结算方式 And Rownum < 2;
      Exception
        When Others Then
          n_存在 := 0;
      End;
      If n_存在 = 0 Then
        v_Err_Msg := '传入的挂号单据不是' || v_结算方式 || '结算的,无法退号!';
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

  --补充结算检查，已存在补结算数据的，不能退号
  Begin
    Select 1
    Into n_存在
    From 费用补充记录 A,
         (Select Distinct 结帐id
           From 门诊费用记录
           Where NO = v_No And 记录性质 = 4
           Union
           Select Distinct 结帐id From 住院费用记录 Where NO = v_No And 记录性质 = 5) B
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
  End If;
  --获取操作员信息
  v_Temp := Zl_Identity(1);
  Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_Temp From Dual;
  Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into v_操作员编号 From Dual;
  Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_操作员姓名 From Dual;
  d_登记时间 := Sysdate;
  n_挂号模式 := zl_GetSysParameter('挂号排班模式');

  Zl_三方机构挂号_Delete(v_No, v_交易流水号, '移动平台退号', d_登记时间);

  v_Temp := '<CZSJ>' || To_Char(d_登记时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
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

Create Or Replace Procedure Zl_挂号序号状态_Delete
(
  操作方式_In Number := 0,
  号别_In     病人挂号记录.号别%Type := Null
) As
  n_预约有效时间 Number(5);
  n_失约用于挂号 Number(2);
  n_挂号有效天数 Number(5);
Begin
  If 操作方式_In = 0 Then
    --清除历史记录
    Delete 挂号序号状态 Where 日期 < Trunc(Sysdate);
  Else
    --清除失约号
    n_预约有效时间 := Nvl(zl_GetSysParameter('预约有效时间', 1111), 0);
    n_失约用于挂号 := Nvl(zl_GetSysParameter('失约用于挂号', 1111), 0);
    n_挂号有效天数 := Nvl(zl_GetSysParameter('挂号有效天数'), 7);
    If n_预约有效时间 <> 0 And n_失约用于挂号 <> 0 Then
      If 号别_In Is Null Then
        For c_失效预约 In (Select b.号码, b.日期, b.序号
                       From 病人挂号记录 A, 挂号序号状态 B
                       Where a.预约时间 - 1 / 24 / 60 * n_预约有效时间 < Sysdate And a.预约时间 > Sysdate - n_挂号有效天数 And a.记录性质 = 2 And
                             a.号别 = b.号码 And a.号序 = b.序号) Loop
          Delete From 挂号序号状态
          Where 日期 = c_失效预约.日期 And 序号 = c_失效预约.序号 And 状态 = 2 And 号码 = c_失效预约.号码;
        End Loop;
      Else
        For c_失效预约 In (Select b.号码, b.日期, b.序号
                       From 病人挂号记录 A, 挂号序号状态 B
                       Where a.预约时间 - 1 / 24 / 60 * n_预约有效时间 < Sysdate And a.预约时间 > Sysdate - n_挂号有效天数 And a.记录性质 = 2 And
                             a.号别 = b.号码 And a.号序 = b.序号 And a.号别 = 号别_In) Loop
          Delete From 挂号序号状态
          Where 日期 = c_失效预约.日期 And 序号 = c_失效预约.序号 And 状态 = 2 And 号码 = c_失效预约.号码;
        End Loop;
      End If;
    End If;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_挂号序号状态_Delete;
/

Create Or Replace Procedure Zl_挂号序号状态_出诊_Delete(记录id_In 临床出诊记录.Id%Type := Null) As
  n_预约有效时间 Number(5);
  n_失约用于挂号 Number(2);
  n_挂号有效天数 Number(5);
Begin

  --清除失约号
  n_预约有效时间 := Nvl(zl_GetSysParameter('预约有效时间', 1111), 0);
  n_失约用于挂号 := Nvl(zl_GetSysParameter('失约用于挂号', 1111), 0);
  n_挂号有效天数 := Nvl(zl_GetSysParameter('挂号有效天数'), 7);
  If n_预约有效时间 <> 0 And n_失约用于挂号 <> 0 Then
    If 记录id_In Is Null Then
      For c_失效预约 In (Select b.记录id, b.序号, b.预约顺序号
                     From 病人挂号记录 A, 临床出诊序号控制 B
                     Where a.预约时间 - 1 / 24 / 60 * n_预约有效时间 < Sysdate And a.预约时间 > Sysdate - n_挂号有效天数 And a.记录性质 = 2 And
                           a.出诊记录id = b.记录id And (a.号序 = b.序号 Or a.号序 = b.备注) And Nvl(b.挂号状态, 0) = 2) Loop
        Update 临床出诊序号控制
        Set 挂号状态 = 0, 操作员姓名 = Null, 工作站ip = Null, 工作站名称 = Null, 备注 = Null
        Where 记录id = c_失效预约.记录id And 序号 = c_失效预约.序号 And 预约顺序号 = c_失效预约.预约顺序号;
      End Loop;
    Else
      For c_失效预约 In (Select b.记录id, b.序号, b.预约顺序号
                     From 病人挂号记录 A, 临床出诊序号控制 B
                     Where a.预约时间 - 1 / 24 / 60 * n_预约有效时间 < Sysdate And a.预约时间 > Sysdate - n_挂号有效天数 And a.记录性质 = 2 And
                           a.出诊记录id = b.记录id And (a.号序 = b.序号 Or a.号序 = b.备注) And b.记录id = 记录id_In And
                           Nvl(b.挂号状态, 0) = 2) Loop
        If c_失效预约.预约顺序号 Is Null Then
          Update 临床出诊序号控制
          Set 挂号状态 = 0, 操作员姓名 = Null, 工作站ip = Null, 工作站名称 = Null, 备注 = Null
          Where 记录id = c_失效预约.记录id And 序号 = c_失效预约.序号;
        Else
          Update 临床出诊序号控制
          Set 挂号状态 = 0, 操作员姓名 = Null, 工作站ip = Null, 工作站名称 = Null, 备注 = Null
          Where 记录id = c_失效预约.记录id And 序号 = c_失效预约.序号 And 预约顺序号 = c_失效预约.预约顺序号;
        End If;
      End Loop;
    End If;
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_挂号序号状态_出诊_Delete;
/


CREATE OR REPLACE Procedure Zl_挂号序号状态_Lock
(
  操作_In       Number, --1-锁定,2-清除锁定
  操作员姓名_In 挂号序号状态.操作员姓名%Type,
  号码_In       挂号序号状态.号码%Type := Null,
  日期_In       挂号序号状态.日期%Type := Null,
  序号_In       挂号序号状态.序号%Type := Null,
  出诊记录ID_In 临床出诊记录.ID%type := Null
) As

  v_姓名       挂号序号状态.操作员姓名%Type;
  v_状态       挂号序号状态.状态%Type;
  v_机器名     挂号序号状态.机器名%Type;
  v_验证机器名 挂号序号状态.机器名%Type;
  v_工作站IP   临床出诊序号控制.工作站IP%Type;
  v_Error      Varchar2(255);
  Err_Custom Exception;
Begin
  Select Terminal Into v_机器名 From V$session Where Audsid = Userenv('sessionid');
  Select SYS_CONTEXT('USERENV','IP_ADDRESS') Into v_工作站IP from dual;
  If 操作_In = 1 Then
    --锁定挂号序号状态
    If 出诊记录ID_In is Null then
      Begin
        Select 操作员姓名, 状态, 机器名
        Into v_姓名, v_状态, v_验证机器名
        From 挂号序号状态
        Where 号码 = 号码_In And 日期 = 日期_In And 序号 = 序号_In;
      Exception
        When Others Then
          Null;
      End;
      If v_姓名 Is Null Then
        Insert Into 挂号序号状态
          (号码, 日期, 序号, 状态, 操作员姓名, 备注, 登记时间, 机器名)
        Values
          (号码_In, 日期_In, 序号_In, 5, 操作员姓名_In, '自助机锁号', Sysdate, v_机器名);
      Else
        v_Error := '序号' || 序号_In || '已被操作员' || v_姓名;
        If v_状态 = 1 Then
          v_Error := v_Error || '使用';
        Elsif v_状态 = 2 Then
          v_Error := v_Error || '预约';
        Elsif v_状态 = 3 Then
          v_Error := v_Error || '预留';
        Elsif v_状态 = 4 Then
          v_Error := v_Error || '退号';
        Elsif v_状态 = 5 Then
          v_Error := v_Error || '(' || v_验证机器名 || ')锁定';
        End If;
        Raise Err_Custom;
      End If;
    Else
      Begin
        Select 操作员姓名, 挂号状态, 工作站名称
        Into v_姓名, v_状态, v_验证机器名
        From 临床出诊序号控制
        Where 记录ID = 出诊记录ID_In And 序号 = 序号_In;
      Exception
        When Others Then
          Null;
      End;
      
      If Nvl(v_状态,0) = 0 Then
        Update 临床出诊序号控制 set 挂号状态=5,锁号时间=Sysdate,操作员姓名=操作员姓名_In,工作站IP=v_工作站IP,工作站名称=v_机器名,备注='自助机锁号'
        Where 记录ID=出诊记录ID_In  And 序号=序号_In;
      Else
        v_Error := '序号' || 序号_In || '已被操作员' || v_姓名;
        If v_状态 = 1 Then
          v_Error := v_Error || '使用';
        Elsif v_状态 = 2 Then
          v_Error := v_Error || '预约';
        Elsif v_状态 = 3 Then
          v_Error := v_Error || '预留';
        Elsif v_状态 = 4 Then
          v_Error := v_Error || '退号';
        Elsif v_状态 = 5 Then
          v_Error := v_Error || '(' || v_验证机器名 || ')锁定';
        End If;
        Raise Err_Custom;
      End If;
    End If;
  Elsif 操作_In = 2 Then
    If 出诊记录ID_In is Null then
       Delete 挂号序号状态 Where 机器名 = v_机器名 And 操作员姓名 = 操作员姓名_In And 状态 = 5;
    Else
      Update 临床出诊序号控制 A set A.挂号状态=0,A.锁号时间=NULL,A.操作员姓名=NULL,A.工作站IP=NULL,A.工作站名称=NULL,A.类型=NULL,A.名称=NULL,A.备注=NULL
      Where A.工作站名称 =v_机器名 And A.工作站IP=v_工作站IP And A.操作员姓名 = 操作员姓名_In And A.挂号状态 = 5 And A.锁号时间 > Sysdate -1
        And Exists (Select 1 From 临床出诊记录 B Where A.记录ID=B.ID And B.是否序号控制 = 1);
      
      Update 临床出诊序号控制 A set A.挂号状态=4,A.锁号时间=NULL,A.操作员姓名=NULL,A.工作站IP=NULL,A.工作站名称=NULL,A.类型=NULL,A.名称=NULL,A.备注=NULL
      Where A.工作站名称 =v_机器名 And A.工作站IP=v_工作站IP And A.操作员姓名 = 操作员姓名_In And A.挂号状态 = 5 And A.锁号时间 > Sysdate -1
        And Exists (Select 1 From 临床出诊记录 B Where A.记录ID=B.ID And B.是否序号控制 = 0);
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_挂号序号状态_Lock;
/

CREATE OR REPLACE Function NextReservationNum
(
  记录ID_In         In 临床出诊序号控制.记录ID%Type,
  序号_In           In 临床出诊序号控制.序号%Type,
  操作员姓名_In     In 临床出诊序号控制.操作员姓名%Type
) Return Varchar2
 --获取最大预约顺序号，只针对预约普通分时段
 Is
  Pragma Autonomous_Transaction;
  v_机器名    临床出诊序号控制.工作站名称%Type;
  v_工作站IP  临床出诊序号控制.工作站IP%Type;
  n_数量      临床出诊序号控制.数量%Type;
  n_已约数    临床出诊序号控制.数量%Type;
  n_Maxno  Number;

  v_Error      Varchar2(255);
  Err_Custom Exception;
Begin
  Select Terminal Into v_机器名 From V$session Where Audsid = Userenv('sessionid');
  Select SYS_CONTEXT('USERENV','IP_ADDRESS') Into v_工作站IP from dual;
  Begin
     Select A.数量,B.已约数
     Into n_数量,n_已约数
     From 临床出诊序号控制 A,
      (Select 记录ID,序号,Count(1) as 已约数 From 临床出诊序号控制 Where 记录ID=记录ID_In And 序号=序号_In And 挂号状态<>0 and 挂号状态<>4 And 预约顺序号 is Not Null group by 记录ID,序号) B
     Where A.记录ID = B.记录ID(+) And A.序号 = B.序号(+) And A.记录ID=记录ID_In And A.序号=序号_In And A.预约顺序号 is Null;
  Exception
    When Others Then
      v_Error:='没找到对应的出诊安排记录';
      Raise Err_Custom;
  End;
  
  If Nvl(n_已约数,0)<Nvl(n_数量,0) Then
    Select Nvl(Max(预约顺序号),0) Into n_Maxno From 临床出诊序号控制 WHERE 记录ID=记录ID_In  And 序号=序号_In;
    --If n_挂号序号=0 then
      n_Maxno:=n_Maxno+1;
      Insert Into 临床出诊序号控制(记录ID,序号,预约顺序号,开始时间,终止时间,数量,是否预约,挂号状态,锁号时间,类型,名称,操作员姓名,工作站IP,工作站名称,备注)
      Select 记录ID,序号,n_Maxno,开始时间,终止时间,1,是否预约,5,Sysdate,类型,名称,操作员姓名_In,v_工作站IP,v_机器名,'自助机锁号'
      From 临床出诊序号控制
      Where 记录ID=记录ID_In  And 序号=序号_In And 预约顺序号 is Null;
  Else
      v_Error:='当前时段预约已超过最大限约数';
      Raise Err_Custom;
  End If;
  Commit;
  Return n_Maxno;
Exception
  When Err_Custom Then
    Rollback;
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Rollback;
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End NextReservationNum;
/


--报表：ZL1_INSIDE_1114_1/固定出诊表
Insert Into zlReports(ID,编号,名称,说明,密码,进纸,打印机,票据,系统,程序ID,功能,修改时间,发布时间,打印方式,禁止开始时间,禁止结束时间) Values(zlReports_ID.NextVal,'ZL1_INSIDE_1114_1','固定出诊表','固定出诊表','I`;g$oi|}90Fiql4H+LL',15,Null,0,&n_System,1114,'固定出诊表',Sysdate,Sysdate,0,To_Date('2016-03-01 00:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('2016-03-01 00:00:00','YYYY-MM-DD HH24:MI:SS'));
Insert Into zlRPTFmts(报表ID,序号,说明,图样,W,H,纸张,纸向,动态纸张) Values(zlReports_ID.CurrVal,1,'固定出诊表',0,11904,16832,9,1,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'出诊表名','出诊表名,202',User||'.临床出诊表',0,Null);
Insert Into zlRPTSQLs(源ID,行号,内容)  Select zlRPTDatas_ID.CurrVal,a.* From (
  Select 1,'Select 出诊表名 From 临床出诊表 Where 排班方式=0 And ID=[0]' From Dual) a;
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,0,'出诊ID',1,Null,0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'临床出诊表_数据','科室名称,202|项目名称,202|医生姓名,202|周一,202|周二,202|周三,202|周四,202|周五,202|周六,202|周日,202',User||'.临床出诊表,'||User||'.临床出诊安排,'||User||'.临床出诊限制,'||User||'.临床出诊号源,'||User||'.部门表,'||User||'.收费项目目录',0,Null);
Insert Into zlRPTSQLs(源ID,行号,内容)  Select zlRPTDatas_ID.CurrVal,a.* From (
  Select 1,'Select e.名称 As 科室名称, f.名称 As 项目名称, b.医生姓名, Max(Decode(c.限制项目, ''周一'', c.上班时段, Null)) As 周一,' From Dual Union All
  Select 2,'       Max(Decode(c.限制项目, ''周一'', c.上班时段, Null)) As 周二, Max(Decode(c.限制项目, ''周一'', c.上班时段, Null)) As 周三,' From Dual Union All
  Select 3,'       Max(Decode(c.限制项目, ''周一'', c.上班时段, Null)) As 周四, Max(Decode(c.限制项目, ''周一'', c.上班时段, Null)) As 周五,' From Dual Union All
  Select 4,'       Max(Decode(c.限制项目, ''周一'', c.上班时段, Null)) As 周六, Max(Decode(c.限制项目, ''周一'', c.上班时段, Null)) As 周日' From Dual Union All
  Select 5,'From 临床出诊表 A, 临床出诊安排 B, 临床出诊限制 C, 临床出诊号源 D, 部门表 E, 收费项目目录 F' From Dual Union All
  Select 6,'Where a.Id = b.出诊id And b.Id = c.安排id(+) And b.号源id = d.Id And d.科室id = e.Id And b.项目id = f.Id And a.排班方式 = 0 And' From Dual Union All
  Select 7,'      a.Id = [0]' From Dual Union All
  Select 8,'Group By e.名称, f.名称, b.医生姓名' From Dual Union All
  Select 9,'Order By e.名称, f.名称, b.医生姓名' From Dual) a;
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,0,'出诊ID',1,Null,0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签2',2,Null,0,'任意表1',21,'发布人:[操作员姓名]',Null,150,6460,1710,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签1',2,Null,0,Null,0,'[出诊表名.出诊表名]',Null,3960,435,2895,300,0,0,1,'宋体',14,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签3',2,Null,0,'任意表1',23,'发布日期:[yyyy-mm-dd]',Null,9825,6460,1890,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'任意表1',4,Null,0,Null,0,Null,Null,150,930,11565,5430,255,0,0,'宋体',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-1,0,Null,Null,'[临床出诊表_数据.科室名称]','4^255^科室^0^0',0,0,1320,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-2,1,Null,Null,'[临床出诊表_数据.项目名称]','4^255^项目^0^0',0,0,1800,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-3,2,Null,Null,'[临床出诊表_数据.医生姓名]','4^255^医生^0^0',0,0,1110,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-4,3,Null,Null,'[临床出诊表_数据.周一]','4^255^周一^0^0',0,0,660,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-5,4,Null,Null,'[临床出诊表_数据.周二]','4^255^周二^0^0',0,0,600,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-6,5,Null,Null,'[临床出诊表_数据.周三]','4^255^周三^0^0',0,0,630,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-7,6,Null,Null,'[临床出诊表_数据.周四]','4^255^周四^0^0',0,0,570,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-8,7,Null,Null,'[临床出诊表_数据.周五]','4^255^周五^0^0',0,0,630,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-9,8,Null,Null,'[临床出诊表_数据.周六]','4^255^周六^0^0',0,0,585,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-10,9,Null,Null,'[临床出诊表_数据.周日]','4^255^周日^0^0',0,0,600,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);

--报表：ZL1_INSIDE_1114_1/固定出诊表
Insert into zlProgFuncs(系统,序号,功能,说明) Values(&n_System,1114,'固定出诊表','固定出诊表');
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
  Select &n_System,1114,'固定出诊表',User,'部门表','SELECT' From Dual Union All
  Select &n_System,1114,'固定出诊表',User,'临床出诊安排','SELECT' From Dual Union All
  Select &n_System,1114,'固定出诊表',User,'临床出诊表','SELECT' From Dual Union All
  Select &n_System,1114,'固定出诊表',User,'临床出诊号源','SELECT' From Dual Union All
  Select &n_System,1114,'固定出诊表',User,'临床出诊限制','SELECT' From Dual Union All
  Select &n_System,1114,'固定出诊表',User,'收费项目目录','SELECT' From Dual;
  
--报表：ZL1_INSIDE_1114_2/月出诊表
Insert Into zlReports(ID,编号,名称,说明,密码,进纸,打印机,票据,系统,程序ID,功能,修改时间,发布时间,打印方式,禁止开始时间,禁止结束时间) Values(zlReports_ID.NextVal,'ZL1_INSIDE_1114_2','月出诊表','月出诊表','Wg"|?kw}~8-@sht1V+LL',15,'发送至 OneNote 2010',0,&n_System,1114,'月出诊表',Sysdate,Sysdate,0,To_Date('2016-03-01 00:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('2016-03-01 00:00:00','YYYY-MM-DD HH24:MI:SS'));
Insert Into zlRPTFmts(报表ID,序号,说明,图样,W,H,纸张,纸向,动态纸张) Values(zlReports_ID.CurrVal,1,'月出诊表(31日)',0,21563,11906,256,1,0);
Insert Into zlRPTFmts(报表ID,序号,说明,图样,W,H,纸张,纸向,动态纸张) Values(zlReports_ID.CurrVal,2,'月出诊表(30日)',0,20843,11906,256,1,0);
Insert Into zlRPTFmts(报表ID,序号,说明,图样,W,H,纸张,纸向,动态纸张) Values(zlReports_ID.CurrVal,3,'月出诊表(29日)',0,20258,11906,256,1,0);
Insert Into zlRPTFmts(报表ID,序号,说明,图样,W,H,纸张,纸向,动态纸张) Values(zlReports_ID.CurrVal,4,'月出诊表(28日)',0,19778,11906,256,1,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'出诊表名','出诊表名,202',User||'.临床出诊表',0,Null);
Insert Into zlRPTSQLs(源ID,行号,内容)  Select zlRPTDatas_ID.CurrVal,a.* From (
  Select 1,'Select 出诊表名 From 临床出诊表 Where ID=[0] And 排班方式=1' From Dual) a;
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,0,'出诊ID',1,Null,0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'临床出诊表_数据','科室名称,202|项目名称,202|医生姓名,202|C1,202|C2,202|C3,202|C4,202|C5,202|C6,202|C7,202|C8,202|C9,202|C10,202|C11,202|C12,202|C13,202|C14,202|C15,202|C16,202|C17,202|C18,202|C19,202|C20,202|C21,202|C22,202|C23,202|C24,202|C25,202|C26,202|C27,202|C28,202|C29,202|C30,202|C31,202',User||'.临床出诊表,'||User||'.临床出诊安排,'||User||'.临床出诊记录,'||User||'.临床出诊号源,'||User||'.部门表,'||User||'.收费项目目录',1,Null);
Insert Into zlRPTSQLs(源ID,行号,内容)  Select zlRPTDatas_ID.CurrVal,a.* From (
  Select 1,'Select e.名称 As 科室名称, f.名称 As 项目名称, b.医生姓名, Max(Decode(To_Char(c.出诊日期, ''DD''), ''01'', c.上班时段, Null)) As C1,' From Dual Union All
  Select 2,'       Max(Decode(To_Char(c.出诊日期, ''DD''), ''02'', c.上班时段, Null)) As C2,' From Dual Union All
  Select 3,'       Max(Decode(To_Char(c.出诊日期, ''DD''), ''03'', c.上班时段, Null)) As C3,' From Dual Union All
  Select 4,'       Max(Decode(To_Char(c.出诊日期, ''DD''), ''04'', c.上班时段, Null)) As C4,' From Dual Union All
  Select 5,'       Max(Decode(To_Char(c.出诊日期, ''DD''), ''05'', c.上班时段, Null)) As C5,' From Dual Union All
  Select 6,'       Max(Decode(To_Char(c.出诊日期, ''DD''), ''06'', c.上班时段, Null)) As C6,' From Dual Union All
  Select 7,'       Max(Decode(To_Char(c.出诊日期, ''DD''), ''07'', c.上班时段, Null)) As C7,' From Dual Union All
  Select 8,'       Max(Decode(To_Char(c.出诊日期, ''DD''), ''08'', c.上班时段, Null)) As C8,' From Dual Union All
  Select 9,'       Max(Decode(To_Char(c.出诊日期, ''DD''), ''09'', c.上班时段, Null)) As C9,' From Dual Union All
  Select 10,'       Max(Decode(To_Char(c.出诊日期, ''DD''), ''10'', c.上班时段, Null)) As C10,' From Dual Union All
  Select 11,'       Max(Decode(To_Char(c.出诊日期, ''DD''), ''11'', c.上班时段, Null)) As C11,' From Dual Union All
  Select 12,'       Max(Decode(To_Char(c.出诊日期, ''DD''), ''12'', c.上班时段, Null)) As C12,' From Dual Union All
  Select 13,'       Max(Decode(To_Char(c.出诊日期, ''DD''), ''13'', c.上班时段, Null)) As C13,' From Dual Union All
  Select 14,'       Max(Decode(To_Char(c.出诊日期, ''DD''), ''14'', c.上班时段, Null)) As C14,' From Dual Union All
  Select 15,'       Max(Decode(To_Char(c.出诊日期, ''DD''), ''15'', c.上班时段, Null)) As C15,' From Dual Union All
  Select 16,'       Max(Decode(To_Char(c.出诊日期, ''DD''), ''16'', c.上班时段, Null)) As C16,' From Dual Union All
  Select 17,'       Max(Decode(To_Char(c.出诊日期, ''DD''), ''17'', c.上班时段, Null)) As C17,' From Dual Union All
  Select 18,'       Max(Decode(To_Char(c.出诊日期, ''DD''), ''18'', c.上班时段, Null)) As C18,' From Dual Union All
  Select 19,'       Max(Decode(To_Char(c.出诊日期, ''DD''), ''19'', c.上班时段, Null)) As C19,' From Dual Union All
  Select 20,'       Max(Decode(To_Char(c.出诊日期, ''DD''), ''20'', c.上班时段, Null)) As C20,' From Dual Union All
  Select 21,'       Max(Decode(To_Char(c.出诊日期, ''DD''), ''21'', c.上班时段, Null)) As C21,' From Dual Union All
  Select 22,'       Max(Decode(To_Char(c.出诊日期, ''DD''), ''22'', c.上班时段, Null)) As C22,' From Dual Union All
  Select 23,'       Max(Decode(To_Char(c.出诊日期, ''DD''), ''23'', c.上班时段, Null)) As C23,' From Dual Union All
  Select 24,'       Max(Decode(To_Char(c.出诊日期, ''DD''), ''24'', c.上班时段, Null)) As C24,' From Dual Union All
  Select 25,'       Max(Decode(To_Char(c.出诊日期, ''DD''), ''25'', c.上班时段, Null)) As C25,' From Dual Union All
  Select 26,'       Max(Decode(To_Char(c.出诊日期, ''DD''), ''26'', c.上班时段, Null)) As C26,' From Dual Union All
  Select 27,'       Max(Decode(To_Char(c.出诊日期, ''DD''), ''27'', c.上班时段, Null)) As C27,' From Dual Union All
  Select 28,'       Max(Decode(To_Char(c.出诊日期, ''DD''), ''28'', c.上班时段, Null)) As C28,' From Dual Union All
  Select 29,'       Max(Case' From Dual Union All
  Select 30,'             When To_Number(To_Char(Last_Day(c.出诊日期), ''DD'')) < 29 Then' From Dual Union All
  Select 31,'              Null' From Dual Union All
  Select 32,'             Else' From Dual Union All
  Select 33,'              Decode(To_Char(c.出诊日期, ''DD''), ''29'', c.上班时段, '' '')' From Dual Union All
  Select 34,'           End) As C29,' From Dual Union All
  Select 35,'       Max(Case' From Dual Union All
  Select 36,'             When To_Number(To_Char(Last_Day(c.出诊日期), ''DD'')) < 30 Then' From Dual Union All
  Select 37,'              Null' From Dual Union All
  Select 38,'             Else' From Dual Union All
  Select 39,'              Decode(To_Char(c.出诊日期, ''DD''), ''30'', c.上班时段, '' '')' From Dual Union All
  Select 40,'           End) As C30,' From Dual Union All
  Select 41,'       Max(Case' From Dual Union All
  Select 42,'             When To_Number(To_Char(Last_Day(c.出诊日期), ''DD'')) < 31 Then' From Dual Union All
  Select 43,'              Null' From Dual Union All
  Select 44,'             Else' From Dual Union All
  Select 45,'              Decode(To_Char(c.出诊日期, ''DD''), ''31'', c.上班时段, '' '')' From Dual Union All
  Select 46,'           End) As C31' From Dual Union All
  Select 47,'From 临床出诊安排 B, 临床出诊表 A, 临床出诊记录 C, 临床出诊号源 D, 部门表 E, 收费项目目录 F' From Dual Union All
  Select 48,'Where b.Id = c.安排id(+) And b.号源id = d.Id And b.出诊id = a.Id And b.项目id = f.Id And' From Dual Union All
  Select 49,'      d.科室id = e.Id And a.排班方式 = 1 And a.Id = [0]' From Dual Union All
  Select 50,'Group By e.名称, f.名称, b.医生姓名' From Dual Union All
  Select 51,'Order By e.名称, f.名称, b.医生姓名' From Dual Union All
  Select 52,Null From Dual) a;
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,0,'出诊ID',1,Null,0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签2',2,Null,0,'任意表1',21,'发布人:[操作员姓名]',Null,150,6520,1710,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签1',2,Null,0,'任意表1',12,'[出诊表名.出诊表名]',Null,9302,165,3150,300,0,0,1,'宋体',14,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签3',2,Null,0,'任意表1',23,'发布时间:[yyyy-mm-dd]',Null,19458,6520,1890,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'任意表1',4,Null,0,Null,0,Null,Null,150,615,21198,5805,255,0,0,'宋体',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-1,0,Null,Null,'[临床出诊表_数据.科室名称]','4^300^科室^0^0',0,0,1305,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-2,1,Null,Null,'[临床出诊表_数据.项目名称]','4^300^项目^0^0',0,0,1965,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-3,2,Null,Null,'[临床出诊表_数据.医生姓名]','4^300^医生^0^0',0,0,1110,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-4,3,Null,Null,'[临床出诊表_数据.C1]','4^300^1^0^0',0,0,525,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-5,4,Null,Null,'[临床出诊表_数据.C2]','4^300^2^0^0',0,0,555,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-6,5,Null,Null,'[临床出诊表_数据.C3]','4^300^3^0^0',0,0,495,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-7,6,Null,Null,'[临床出诊表_数据.C4]','4^300^4^0^0',0,0,480,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-8,7,Null,Null,'[临床出诊表_数据.C5]','4^300^5^0^0',0,0,495,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-9,8,Null,Null,'[临床出诊表_数据.C6]','4^300^6^0^0',0,0,495,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-10,9,Null,Null,'[临床出诊表_数据.C7]','4^300^7^0^0',0,0,510,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-11,10,Null,Null,'[临床出诊表_数据.C8]','4^300^8^0^0',0,0,465,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-12,11,Null,Null,'[临床出诊表_数据.C9]','4^300^9^0^0',0,0,510,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-13,12,Null,Null,'[临床出诊表_数据.C10]','4^300^10^0^0',0,0,465,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-14,13,Null,Null,'[临床出诊表_数据.C11]','4^300^11^0^0',0,0,480,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-15,14,Null,Null,'[临床出诊表_数据.C12]','4^300^12^0^0',0,0,525,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-16,15,Null,Null,'[临床出诊表_数据.C13]','4^300^13^0^0',0,0,540,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-17,16,Null,Null,'[临床出诊表_数据.C14]','4^300^14^0^0',0,0,525,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-18,17,Null,Null,'[临床出诊表_数据.C15]','4^300^15^0^0',0,0,525,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-19,18,Null,Null,'[临床出诊表_数据.C16]','4^300^16^0^0',0,0,540,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-20,19,Null,Null,'[临床出诊表_数据.C17]','4^300^17^0^0',0,0,540,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-21,20,Null,Null,'[临床出诊表_数据.C18]','4^300^18^0^0',0,0,570,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-22,21,Null,Null,'[临床出诊表_数据.C19]','4^300^19^0^0',0,0,525,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-23,22,Null,Null,'[临床出诊表_数据.C20]','4^300^20^0^0',0,0,525,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-24,23,Null,Null,'[临床出诊表_数据.C21]','4^300^21^0^0',0,0,570,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-25,24,Null,Null,'[临床出诊表_数据.C22]','4^300^22^0^0',0,0,585,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-26,25,Null,Null,'[临床出诊表_数据.C23]','4^300^23^0^0',0,0,600,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-27,26,Null,Null,'[临床出诊表_数据.C24]','4^300^24^0^0',0,0,615,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-28,27,Null,Null,'[临床出诊表_数据.C25]','4^300^25^0^0',0,0,570,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-29,28,Null,Null,'[临床出诊表_数据.C26]','4^300^26^0^0',0,0,540,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-30,29,Null,Null,'[临床出诊表_数据.C27]','4^300^27^0^0',0,0,570,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-31,30,Null,Null,'[临床出诊表_数据.C28]','4^300^28^0^0',0,0,585,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-32,31,Null,Null,'[临床出诊表_数据.C29]','4^300^29^0^0',0,0,525,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-33,32,Null,Null,'[临床出诊表_数据.C30]','4^300^30^0^0',0,0,585,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-34,33,Null,Null,'[临床出诊表_数据.C31]','4^300^31^0^0',0,0,585,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null);

--报表：ZL1_INSIDE_1114_2/月出诊表
Insert into zlProgFuncs(系统,序号,功能,说明) Values(&n_System,1114,'月出诊表','月出诊表');
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
  Select &n_System,1114,'月出诊表',User,'部门表','SELECT' From Dual Union All
  Select &n_System,1114,'月出诊表',User,'临床出诊安排','SELECT' From Dual Union All
  Select &n_System,1114,'月出诊表',User,'临床出诊表','SELECT' From Dual Union All
  Select &n_System,1114,'月出诊表',User,'临床出诊号源','SELECT' From Dual Union All
  Select &n_System,1114,'月出诊表',User,'临床出诊记录','SELECT' From Dual Union All
  Select &n_System,1114,'月出诊表',User,'收费项目目录','SELECT' From Dual;


--报表：ZL1_INSIDE_1114_3/周出诊表
Insert Into zlReports(ID,编号,名称,说明,密码,进纸,打印机,票据,系统,程序ID,功能,修改时间,发布时间,打印方式,禁止开始时间,禁止结束时间) Values(zlReports_ID.NextVal,'ZL1_INSIDE_1114_3','周出诊表','周出诊表','Tg"}<kw}}8-@pht1T+LL',15,Null,0,&n_System,1114,'周出诊表',Sysdate,Sysdate,0,To_Date('2016-03-01 00:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('2016-03-01 00:00:00','YYYY-MM-DD HH24:MI:SS'));
Insert Into zlRPTFmts(报表ID,序号,说明,图样,W,H,纸张,纸向,动态纸张) Values(zlReports_ID.CurrVal,1,'周出诊表',0,11904,16832,9,1,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'出诊表名','出诊表名,202',User||'.临床出诊表',0,Null);
Insert Into zlRPTSQLs(源ID,行号,内容)  Select zlRPTDatas_ID.CurrVal,a.* From (
  Select 1,'Select 出诊表名 From 临床出诊表 Where 排班方式=2 And ID=[0]' From Dual) a;
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,0,'出诊ID',1,Null,0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'临床出诊表_数据','科室名称,202|项目名称,202|医生姓名,202|C1,202|C2,202|C3,202|C4,202|C5,202|C6,202|C7,202',User||'.临床出诊表,'||User||'.临床出诊安排,'||User||'.临床出诊记录,'||User||'.临床出诊号源,'||User||'.部门表,'||User||'.收费项目目录',0,Null);
Insert Into zlRPTSQLs(源ID,行号,内容)  Select zlRPTDatas_ID.CurrVal,a.* From (
  Select 1,'Select e.名称 As 科室名称, f.名称 As 项目名称, b.医生姓名, ' From Dual Union All
  Select 2,'       Max(Decode(To_Char(c.出诊日期, ''D''), 2, c.上班时段, Null)) As C1,' From Dual Union All
  Select 3,'       Max(Decode(To_Char(c.出诊日期, ''D''), 3, c.上班时段, Null)) As C2,' From Dual Union All
  Select 4,'       Max(Decode(To_Char(c.出诊日期, ''D''), 4, c.上班时段, Null)) As C3,' From Dual Union All
  Select 5,'       Max(Decode(To_Char(c.出诊日期, ''D''), 5, c.上班时段, Null)) As C4,' From Dual Union All
  Select 6,'       Max(Decode(To_Char(c.出诊日期, ''D''), 6, c.上班时段, Null)) As C5,' From Dual Union All
  Select 7,'       Max(Decode(To_Char(c.出诊日期, ''D''), 7, c.上班时段, Null)) As C6,' From Dual Union All
  Select 8,'       Max(Decode(To_Char(c.出诊日期, ''D''), 1, c.上班时段, Null)) As C7' From Dual Union All
  Select 9,'From 临床出诊表 A, 临床出诊安排 B, 临床出诊记录 C, 临床出诊号源 D, 部门表 E, 收费项目目录 F' From Dual Union All
  Select 10,'Where a.Id = b.出诊id And b.号源id = d.Id And b.Id = c.安排id(+) And d.科室id = e.Id And b.项目id = f.Id And a.排班方式 = 2 And' From Dual Union All
  Select 11,'      a.Id = [0]' From Dual Union All
  Select 12,'Group By e.名称, f.名称, b.医生姓名' From Dual Union All
  Select 13,'Order By e.名称, f.名称, b.医生姓名' From Dual) a;
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,0,'出诊ID',1,Null,0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'日期','C1,202|C2,202|C3,202|C4,202|C5,202|C6,202|C7,202',User||'.临床出诊表,'||User||'.临床出诊安排,'||User||'.临床出诊记录',0,Null);
Insert Into zlRPTSQLs(源ID,行号,内容)  Select zlRPTDatas_ID.CurrVal,a.* From (
  Select 1,'Select To_Char(Trunc(出诊日期, ''d'') + 1, ''DD'') As C1, To_Char(Trunc(出诊日期, ''d'') + 2, ''DD'') As C2,' From Dual Union All
  Select 2,'       To_Char(Trunc(出诊日期, ''d'') + 3, ''DD'') As C3, To_Char(Trunc(出诊日期, ''d'') + 4, ''DD'') As C4,' From Dual Union All
  Select 3,'       To_Char(Trunc(出诊日期, ''d'') + 5, ''DD'') As C5, To_Char(Trunc(出诊日期, ''d'') + 6, ''DD'') As C6,' From Dual Union All
  Select 4,'       To_Char(Trunc(出诊日期, ''d'') + 7, ''DD'') As C7' From Dual Union All
  Select 5,'From (Select c.出诊日期' From Dual Union All
  Select 6,'       From 临床出诊表 A, 临床出诊安排 B, 临床出诊记录 C' From Dual Union All
  Select 7,'       Where a.Id = b.出诊id And b.Id = c.安排id(+) And a.排班方式 = 2 And a.Id = [0]' From Dual Union All
  Select 8,'			And c.出诊日期 Is Not Null And Rownum < 2)' From Dual Union All
  Select 9,Null From Dual) a;
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,0,'出诊ID',1,Null,0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签2',2,Null,0,'任意表1',21,'发布人:[操作员姓名]',Null,210,4720,1710,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签1',2,Null,0,'任意表1',12,'[出诊表名.出诊表名]',Null,4522,195,2895,300,0,0,1,'宋体',14,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签3',2,Null,0,'任意表1',23,'发布日期:[yyyy-mm-dd]',Null,9840,4720,1890,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'任意表1',4,Null,0,Null,0,Null,Null,210,660,11520,3960,255,0,0,'宋体',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-1,0,Null,Null,'[临床出诊表_数据.科室名称]','4^600^科室^0^0',0,0,1350,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-2,1,Null,Null,'[临床出诊表_数据.项目名称]','4^600^项目^0^0',0,0,1710,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-3,2,Null,Null,'[临床出诊表_数据.医生姓名]','4^600^医生^0^0',0,0,1350,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-4,3,Null,Null,'[临床出诊表_数据.C1]','4^600^周一
[日期.C1]^0^0',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-5,4,Null,Null,'[临床出诊表_数据.C2]','4^600^周二
[日期.C2]^0^0',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-6,5,Null,Null,'[临床出诊表_数据.C3]','4^600^周三
[日期.C3]^0^0',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-7,6,Null,Null,'[临床出诊表_数据.C4]','4^600^周四
[日期.C4]^0^0',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-8,7,Null,Null,'[临床出诊表_数据.C5]','4^600^周五
[日期.C5]^0^0',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-9,8,Null,Null,'[临床出诊表_数据.C6]','4^600^周六
[日期.C6]^0^0',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-10,9,Null,Null,'[临床出诊表_数据.C7]','4^600^周日
[日期.C7]^0^0',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);

--报表：ZL1_INSIDE_1114_3/周出诊表
Insert into zlProgFuncs(系统,序号,功能,说明) Values(&n_System,1114,'周出诊表','周出诊表');
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
  Select &n_System,1114,'周出诊表',User,'部门表','SELECT' From Dual Union All
  Select &n_System,1114,'周出诊表',User,'临床出诊安排','SELECT' From Dual Union All
  Select &n_System,1114,'周出诊表',User,'临床出诊表','SELECT' From Dual Union All
  Select &n_System,1114,'周出诊表',User,'临床出诊号源','SELECT' From Dual Union All
  Select &n_System,1114,'周出诊表',User,'临床出诊记录','SELECT' From Dual Union All
  Select &n_System,1114,'周出诊表',User,'收费项目目录','SELECT' From Dual;


--报表：ZL1_INSIDE_1114_4/病人预约清单
Insert Into zlReports(ID,编号,名称,说明,密码,进纸,打印机,票据,系统,程序ID,功能,修改时间,发布时间,打印方式,禁止开始时间,禁止结束时间) Values(zlReports_ID.NextVal,'ZL1_INSIDE_1114_4','病人预约清单','病人预约清单','Lv!a7lom~"'||CHR(38)||'Fhyw*X,T\',15,'发送至 OneNote 2010',0,&n_System,1114,'病人预约清单',Sysdate,Sysdate,0,To_Date('2016-04-01 00:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('2016-04-01 00:00:00','YYYY-MM-DD HH24:MI:SS'));
Insert Into zlRPTFmts(报表ID,序号,说明,图样,W,H,纸张,纸向,动态纸张) Values(zlReports_ID.CurrVal,1,'病人预约清单',0,11906,16838,9,2,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'病人挂号记录_数据','姓名,202|性别,202|年龄,202|家庭地址,202|家庭电话,202|号别,202|号码,202|科室,202|收费项目,202|医生,202|替诊医生,202|预约单号,202|预约时间,135',User||'.病人挂号记录,'||User||'.病人信息,'||User||'.临床出诊号源,'||User||'.临床出诊记录,'||User||'.部门表,'||User||'.收费项目目录,'||User||'.病人服务信息记录',0,Null);
Insert Into zlRPTSQLs(源ID,行号,内容)  Select zlRPTDatas_ID.CurrVal,a.* From (
  Select 1,'Select a.姓名, a.性别, a.年龄, b.家庭地址, b.家庭电话, c.号类 As 号别, c.号码, e.名称 As 科室, f.名称 As 收费项目, d.医生姓名 As 医生, d.替诊医生姓名 As 替诊医生,' From Dual Union All
  Select 2,'       a.No As 预约单号, a.发生时间 As 预约时间' From Dual Union All
  Select 3,'From 病人挂号记录 A, 病人信息 B, 临床出诊号源 C, 临床出诊记录 D, 部门表 E, 收费项目目录 F, 病人服务信息记录 G' From Dual Union All
  Select 4,'Where a.id = g.挂号id And g.通知类型 In (1,2) And a.记录状态 = 1 And a.病人id = b.病人id(+) And' From Dual Union All
  Select 5,'      g.记录id = d.Id And d.号源id = c.Id And d.科室id = e.Id And d.项目id = f.Id And g.记录id In (Select Column_Value From Table(f_Str2list([0])))' From Dual Union All
  Select 6,Null From Dual) a;
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,0,'出诊记录IDS',0,Null,0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签1',2,Null,0,'任意表1',12,'病人预约清单',Null,7545,195,1800,300,0,0,1,'宋体',14,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'任意表1',4,Null,0,Null,0,Null,Null,180,645,16530,4545,255,0,0,'宋体',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-1,6,Null,Null,'[病人挂号记录_数据.号码]','4^315^号码',0,0,810,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-2,7,Null,Null,'[病人挂号记录_数据.科室]','4^315^科室',0,0,1440,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-3,8,Null,Null,'[病人挂号记录_数据.收费项目]','4^315^收费项目',0,0,1620,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-4,9,Null,Null,'[病人挂号记录_数据.医生]','4^315^医生',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-5,10,Null,Null,'[病人挂号记录_数据.替诊医生]','4^315^替诊医生',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-6,11,Null,Null,'[病人挂号记录_数据.预约单号]','4^315^预约单号',0,0,1005,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-7,12,Null,Null,'[病人挂号记录_数据.预约时间]','4^315^预约时间',0,0,1665,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-8,0,Null,Null,'[病人挂号记录_数据.姓名]','4^315^姓名',0,0,1140,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-9,1,Null,Null,'[病人挂号记录_数据.性别]','4^315^性别',0,0,705,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-10,2,Null,Null,'[病人挂号记录_数据.年龄]','4^315^年龄',0,0,855,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-11,3,Null,Null,'[病人挂号记录_数据.家庭地址]','4^315^家庭地址',0,0,2850,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-12,4,Null,Null,'[病人挂号记录_数据.家庭电话]','4^315^联系电话',0,0,1500,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-13,5,Null,Null,'[病人挂号记录_数据.号别]','4^315^号别',0,0,840,0,0,0,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null);

--报表：ZL1_INSIDE_1114_4/病人预约清单
Insert into zlProgFuncs(系统,序号,功能,说明) Values(&n_System,1114,'病人预约清单','病人预约清单');
Insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
  Select &n_System,1114,'病人预约清单',User,'病人服务信息记录','SELECT' From Dual Union All
  Select &n_System,1114,'病人预约清单',User,'病人挂号记录','SELECT' From Dual Union All
  Select &n_System,1114,'病人预约清单',User,'病人信息','SELECT' From Dual Union All
  Select &n_System,1114,'病人预约清单',User,'部门表','SELECT' From Dual Union All
  Select &n_System,1114,'病人预约清单',User,'临床出诊号源','SELECT' From Dual Union All
  Select &n_System,1114,'病人预约清单',User,'临床出诊记录','SELECT' From Dual Union All
  Select &n_System,1114,'病人预约清单',User,'收费项目目录','SELECT' From Dual;





