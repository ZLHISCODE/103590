------------------------------------------------------------------------------
--结构修正部份
------------------------------------------------------------------------------
--145003:张永康,2019-10-14,新增模块急诊医生站
create table 急诊病人来源 
(
   编码                   varchar2(10),
   名称                   varchar2(50),
   缺省标志               NUMBER(1) default 0
)
tablespace zl9BaseItem;

create table 急诊意识状态 
(
   编码                   varchar2(10),
   名称                   varchar2(50),
   缺省标志               NUMBER(1) default 0
)
tablespace zl9BaseItem;

create table 急诊陪同人员 
(
   编码                   varchar2(10),
   名称                   varchar2(50),
   缺省标志               NUMBER(1) default 0
)
tablespace zl9BaseItem;

create table 急诊常见既往史 
(
   编码                   varchar2(10),
   名称                   varchar2(50)
)
tablespace zl9BaseItem;

create table 急诊常用主诉 
(
   编码                   varchar2(4),
   上级                   varchar2(4),
   名称                   varchar2(50),
   简码                  varchar2(20),  
   末级                  number(1) Default 0
)
tablespace zl9BaseItem;

create table 急诊病情级别 
(
   序号                 number(1),
   名称                 varchar2(5),
   严重程度                 varchar2(20),
   级别描述                 varchar2(400),
   响应要求说明               varchar2(400),
   再次评估时限               number(3),
   患者标识颜色               varchar2(6)
)
tablespace zl9BaseItem;

create table 急诊评分方法 
(
   ID                   number(18),
   英文名                  varchar2(100),
   中文名                  varchar2(100),
   说明                   varchar2(1000)
)
tablespace zl9BaseItem;

create table 急诊人工评定规则 
(
   ID                   number(18),
   分类                   varchar2(20),
   指标名称                 varchar2(200),
   适用人群                 varchar2(10),
   病情级别                 number(1)
)
tablespace zl9BaseItem;

create table 急诊评分方法分级 
(
   ID                       number(18),
   方法ID             number(18),
   分值上限                 number(5),
   分值下限                 number(5),
   运算符                  number(2),--1-等于 2-大于 3-小于 4-大于等于 5-小于等于  6-包含
   评分结果描述               varchar2(200),
   病情级别                 number(1)
)
tablespace zl9BaseItem;

create table 急诊评分方法规则 
(
   ID                   number(18),
   方法ID             number(18),
   指标ID                 number(18),
   指标年龄ID               number(18),
   指标值上限                number(10,2),
   指标值下限                number(10,2),
   运算符                    number(2),--1-等于 2-大于 3-小于 4-大于等于 5-小于等于  6-包含
   指标结果分值               number(5),
   指标结果描述               varchar2(100),
   病情级别                 number(1)
)
tablespace zl9BaseItem;

create table 急诊评分指标 
(
   ID                   number(18),
   方法ID               number(18),
   指标名称                 varchar2(100),
   值域类型                 number(1),
   值域范围                 varchar2(4000),
   值域单位                 varchar2(20)
)
tablespace zl9BaseItem;

create table 急诊评分指标年龄 
(
   ID                   number(18),
   指标ID               number(18),
   年龄上限                 number(4,1),
   年龄下限                 number(4,1),
   运算符                  number(1),
   年龄单位                 varchar2(4),
   年龄段描述                varchar2(20)
)
tablespace zl9BaseItem;

create table 急诊就诊记录 
(
   ID                   number(18),
   病人ID                 number(18),
   病人年龄                 VARCHAR2(20),
   年龄数值                 number(3),
   年龄单位                 VARCHAR2(4),
   挂号ID                 number(18),
   分诊科室ID               number(18),
   保险类别               varchar2(50),
   病情级别                 number(1),
   分诊病情级别             number(1),
   修订说明                 varchar2(50), 
   修订时间                date,        
   修订人员                varchar2(100),    
   到院时间                 date,
   主诉                     varchar2(50),
   是否三无人员               number(1),
   是否绿色通道             number(1),
   陪同人员                 varchar2(10),
   病人来源                 varchar2(50),
   既往病史                 varchar2(50),
   意识状态                 varchar2(50),
   是否成批就诊               number(1),
   成批就诊人数               number(5),
   是否复合伤                number(1),
   备注                   varchar2(500),
   登记人                  varchar2(100),
   登记时间                 DATE,
   待转出                   Number(3)
)
tablespace zl9CisRec;

create table 急诊分诊记录 
(
   ID                   number(18),
   就诊ID                 number(18),
   分诊次数                 number(2),
   自动病情级别                 number(1),
   人工病情级别                 number(1),
   人工评级说明             varchar2(100),
   修改说明                 varchar2(100),
   分诊科室ID               number(18),
   分诊科室名称             varchar2(100),
   收缩压                  number(3),
   舒张压                  number(3),
   心率                   number(3),
   呼吸频率               number(3),
   指氧饱和度             number(3,1),
   体温                   number(3,1),
   血糖                   number(5,2),
   血钾                   number(5,2),
   体征测量时间           date,
   登记人                  varchar2(100),
   登记时间                 date,
   待转出                   Number(3)
)
tablespace zl9CisRec;

create table 急诊病人评分 
(
   ID                   number(18),
   分诊ID             number(18),
   方法ID             number(18),
   评分方法分值       number(5),
   评分结果描述       varchar2(100),
   病情级别           number(1),
   待转出                   Number(3)
)
tablespace zl9CisRec;

create table 急诊病人评分指标 
(
   评分ID               number(18),
   指标ID                 number(18),
   指标结果文本               varchar2(50),
   待转出                   Number(3)
)
tablespace zl9CisRec;

create sequence 急诊分诊记录_ID Start With 1;

create sequence 急诊就诊记录_ID Start With 1;

create sequence 急诊病人评分_ID Start With 1;

alter table 急诊病人来源 add constraint 急诊病人来源_PK primary key (编码) Using Index Tablespace zl9Indexhis;

Alter table 急诊病人来源 Add Constraint 急诊病人来源_UQ_名称 Unique (名称) Using Index Tablespace zl9Indexhis;

alter table 急诊意识状态 add constraint 急诊意识状态_PK primary key (编码) Using Index Tablespace zl9Indexhis;

Alter table 急诊意识状态 Add Constraint 急诊意识状态_UQ_名称 Unique (名称) Using Index Tablespace zl9Indexhis;

alter table 急诊陪同人员 add constraint 急诊陪同人员_PK primary key (编码) Using Index Tablespace zl9Indexhis;

Alter table 急诊陪同人员 Add Constraint 急诊陪同人员_UQ_名称 Unique (名称) Using Index Tablespace zl9Indexhis;

alter table 急诊常见既往史 add constraint 急诊常见既往史_PK primary key (编码) Using Index Tablespace zl9Indexhis;

Alter table 急诊常见既往史 Add Constraint 急诊常见既往史_UQ_名称 Unique (名称) Using Index Tablespace zl9Indexhis;

alter table 急诊常用主诉 add constraint 急诊常用主诉_PK primary key (编码) Using Index Tablespace zl9Indexhis;

Alter table 急诊常用主诉 Add Constraint 急诊常用主诉_UQ_名称 Unique (名称)  Using Index Tablespace zl9Indexhis;

alter table 急诊病情级别 add constraint 急诊病情级别_PK primary key (序号) Using Index Tablespace zl9Indexhis;

Alter table 急诊病情级别 Add Constraint 急诊病情级别_UQ_名称 Unique (名称) Using Index Tablespace zl9Indexhis;

alter table 急诊评分方法 add constraint 急诊评分方法_PK primary key (ID) Using Index Tablespace zl9Indexhis;

Alter table 急诊评分方法 Add Constraint 急诊评分方法_UQ_中文名 Unique (中文名) Using Index Tablespace zl9Indexhis;

alter table 急诊人工评定规则 add constraint 急诊人工评定规则_PK primary key(ID) using index tablespace zl9Indexhis;

Alter table 急诊人工评定规则 Add Constraint 急诊人工评定规则_UQ_指标名称 Unique (指标名称,适用人群) Using Index Tablespace zl9Indexhis;

alter table 急诊评分方法分级 add constraint 急诊评分方法分级_PK primary key (ID) using index tablespace zl9Indexhis;

alter table 急诊评分方法规则 add constraint 急诊评分方法规则_PK primary key (ID) using index tablespace zl9Indexhis;

alter table 急诊评分指标 add constraint 急诊评分指标_PK primary key(ID) using index tablespace zl9Indexhis;

Alter table 急诊评分指标 Add Constraint 急诊评分指标_UQ_指标名称 Unique(指标名称) Using Index Tablespace zl9Indexhis;

alter table 急诊评分指标年龄 add constraint 急诊评分指标年龄_PK primary key (ID) using index tablespace zl9Indexhis;

alter table 急诊就诊记录 add constraint 急诊就诊记录_PK primary key (ID) using index tablespace zl9IndexCis;

alter table 急诊分诊记录 add constraint 急诊分诊记录_PK primary key (ID) using index tablespace zl9IndexCis;

alter table 急诊病人评分 add constraint 急诊病人评分_PK primary key (ID) using index tablespace zl9IndexCis;

alter table 急诊病人评分指标 add constraint 急诊病人评分指标_PK primary key (评分ID,指标ID) using index tablespace zl9IndexCis;

Create Index 急诊常用主诉_IX_上级 on 急诊常用主诉(上级) Tablespace zl9Indexhis;

Create Index 急诊就诊记录_IX_病人ID On 急诊就诊记录(病人ID) Tablespace zl9IndexCis;

Create Index 急诊就诊记录_IX_挂号ID On 急诊就诊记录(挂号ID) Tablespace zl9IndexCis;

Create Index 急诊就诊记录_IX_登记时间 On 急诊就诊记录(登记时间) Tablespace zl9IndexCis;

Create Index 急诊分诊记录_IX_就诊ID On 急诊分诊记录(就诊ID) Tablespace zl9IndexCis;

Create Index 急诊分诊记录_IX_登记时间 On 急诊分诊记录(登记时间) Tablespace zl9IndexCis;

Create Index 急诊病人评分_IX_分诊ID On 急诊病人评分(分诊ID) Tablespace zl9IndexCis;

Create Index 急诊就诊记录_IX_待转出 On 急诊就诊记录(待转出) Tablespace zl9IndexCis;
Create Index 急诊分诊记录_IX_待转出 On 急诊分诊记录(待转出) Tablespace zl9IndexCis;
Create Index 急诊病人评分_IX_待转出 On 急诊病人评分(待转出) Tablespace zl9IndexCis;
Create Index 急诊病人评分指标_IX_待转出 On 急诊病人评分指标(待转出) Tablespace zl9IndexCis;


--145003:张永康,2019-10-14,新增模块急诊医生站
Alter Table 急诊常用主诉 Add Constraint 急诊常用主诉_FK_上级 Foreign Key (上级) References 急诊常用主诉(编码);
alter table 急诊人工评定规则 add constraint 急诊人工评定规则_FK_病情级别 foreign key (病情级别) references 急诊病情级别(序号) On Delete Cascade;
alter table 急诊评分方法分级 add constraint 急诊评分方法分级_FK_方法ID foreign key (方法ID) references 急诊评分方法(ID) On Delete Cascade;
alter table 急诊评分方法分级 add constraint 急诊评分方法分级_FK_病情级别 foreign key (病情级别) references 急诊病情级别(序号) On Delete Cascade;
alter table 急诊评分方法规则 add constraint 急诊评分方法规则_FK_方法ID foreign key (方法ID) references 急诊评分方法(ID) On Delete Cascade;
alter table 急诊评分方法规则 add constraint 急诊评分方法规则_FK_指标ID foreign key (指标ID) references 急诊评分指标(ID) On Delete Cascade;
alter table 急诊评分方法规则 add constraint 急诊评分方法规则_FK_指标年龄ID foreign key (指标年龄ID) references 急诊评分指标年龄(ID) On Delete Cascade;
alter table 急诊评分方法规则 add constraint 急诊评分方法规则_FK_病情级别 foreign key (病情级别) references 急诊病情级别(序号) On Delete Cascade;
alter table 急诊评分指标年龄 add constraint 急诊评分指标年龄_FK_指标ID foreign key(指标ID) references 急诊评分指标(ID) On Delete Cascade;
alter table 急诊就诊记录 add constraint 急诊就诊记录_FK_病人ID foreign key(病人ID) references 病人信息(病人ID);
alter table 急诊就诊记录 add constraint 急诊就诊记录_FK_挂号ID foreign key(挂号ID) references 病人挂号记录(ID);
alter table 急诊分诊记录 add constraint 急诊分诊记录_FK_就诊ID foreign key (就诊ID) references 急诊就诊记录(ID) On Delete Cascade;

alter table 急诊病人评分 add constraint 急诊病人评分_FK_分诊ID foreign key (分诊ID) references 急诊分诊记录(ID) On Delete Cascade;
alter table 急诊病人评分 add constraint 急诊病人评分_FK_方法ID foreign key (方法ID) references 急诊评分方法(ID);
alter table 急诊病人评分指标 add constraint 急诊病人评分指标_FK_评分ID foreign key (评分ID) references 急诊病人评分(ID) On Delete Cascade;
alter table 急诊病人评分指标 add constraint 急诊病人评分指标_FK_指标ID foreign key (指标ID) references 急诊评分指标(ID);


------------------------------------------------------------------------------
--数据修正部份
------------------------------------------------------------------------------
--145003:张永康,2019-10-14,新增模块急诊医生站
Insert Into zlBaseCode(系统,表名,固定,说明,分类) Values( 100,'急诊病人来源',0,'适用于急诊预检分诊登记','医疗工作' ); 

Insert Into zlBaseCode(系统,表名,固定,说明,分类) Values( 100,'急诊意识状态',0,'适用于急诊预检分诊登记','医疗工作' ); 

Insert Into zlBaseCode(系统,表名,固定,说明,分类) Values( 100,'急诊陪同人员',0,'适用于急诊预检分诊登记','医疗工作' ); 

Insert Into zlBaseCode(系统,表名,固定,说明,分类) Values( 100,'急诊常见既往史',0,'适用于急诊预检分诊登记','医疗工作' ); 

Insert Into zlBaseCode(系统,表名,固定,说明,分类) Values( 100,'急诊常用主诉',0,'适用于急诊预检分诊登记','医疗工作' ); 

Insert into zlTables(系统,表名,表空间,分类) Values(100,'急诊病人来源','ZL9BASEITEM','A1');

Insert into zlTables(系统,表名,表空间,分类) Values(100,'急诊意识状态','ZL9BASEITEM','A1');

Insert into zlTables(系统,表名,表空间,分类) Values(100,'急诊陪同人员','ZL9BASEITEM','A1');

Insert into zlTables(系统,表名,表空间,分类) Values(100,'急诊常见既往史','ZL9BASEITEM','A1');

Insert into zlTables(系统,表名,表空间,分类) Values(100,'急诊常用主诉','ZL9BASEITEM','A1');

Insert into zlTables(系统,表名,表空间,分类) Values(100,'急诊病情级别','ZL9BASEITEM','A1');

Insert into zlTables(系统,表名,表空间,分类) Values(100,'急诊人工评定规则','ZL9BASEITEM','A1');

Insert into zlTables(系统,表名,表空间,分类) Values(100,'急诊评分指标年龄','ZL9BASEITEM','A1');

Insert into zlTables(系统,表名,表空间,分类) Values(100,'急诊评分方法','ZL9BASEITEM','A1');

Insert into zlTables(系统,表名,表空间,分类) Values(100,'急诊评分方法分级','ZL9BASEITEM','A1');

Insert into zlTables(系统,表名,表空间,分类) Values(100,'急诊评分方法规则','ZL9BASEITEM','A1');

Insert into zlTables(系统,表名,表空间,分类) Values(100,'急诊评分指标','ZL9BASEITEM','A1');

Insert into zlTables(系统,表名,表空间,分类) Values(100,'急诊就诊记录','ZL9CISREC','B1');

Insert into zlTables(系统,表名,表空间,分类) Values(100,'急诊分诊记录','ZL9CISREC','B1');

Insert into zlTables(系统,表名,表空间,分类) Values(100,'急诊病人评分指标','ZL9CISREC','B1');

Insert into zlTables(系统,表名,表空间,分类) Values(100,'急诊病人评分','ZL9CISREC','B1');

Insert Into zlBakTables(系统,组号,表名,序号,直接转出,停用触发器)
Select 100,4,A.* From (
Select 表名,序号,直接转出,停用触发器 From zlBakTables Where 1 = 0
Union All Select '急诊病人评分指标',6,1,-NULL From Dual
Union All Select '急诊病人评分',7,1,-NULL From Dual
Union All Select '急诊分诊记录',8,1,-NULL From Dual
Union All Select '急诊就诊记录',9,1,-NULL From Dual) A;

Insert Into zlPrograms(序号,标题,说明,系统,部件) Values( 1243,'急诊医生工作站','急诊医生接诊病人，以及写病历和下医嘱等相关工作的处理',100,'zl9CISJob');

Insert Into zlPrograms(序号,标题,说明,系统,部件) Values( 1244,'急诊预检分诊工作站','急诊护士对急诊病人建档、病情评级、分诊等相关工作的处理',100,'zl9CISJob');

Insert Into zlMenus(组别,ID,上级ID,标题,快键,说明,系统,模块,短标题,图标) Select A.组别,ZlMenus_ID.Nextval,A.ID,B.* From (
Select 组别,ID From zlMenus Where 标题 = '临床信息系统' And 组别 = '缺省' And 系统 = 100 And 模块 Is Null) A,
(Select 标题,快键,说明,系统,模块,短标题,图标 From zlMenus Where 1 = 0
Union All Select '急诊医生工作站' ,'D' ,'急诊医生接诊病人，以及书写病历和下达医嘱等相关工作的处理' ,100,1243,'急诊医生' ,243 From Dual
Union All Select '急诊预检分诊工作站' ,'N' ,'急诊护士对急诊病人建档、病情评级、分诊等相关工作的处理' ,100,1244,'急诊分诊' ,244 From Dual
) B;

Insert Into zlParameters(ID,系统,模块,私有,本机,授权,固定,部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明)
Select zlParameters_ID.Nextval,100,1243,A.* From (
Select 私有,本机,授权,固定,部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明 From zlParameters Where 1 = 0
Union All Select 0,0,0,0,0,0,1,'候诊刷新间隔',NULL,'180','急诊医生站候诊列表轮询刷新的频率，N秒刷新一次。','间隔秒钟数','如果启用消息平台，此参数不生效。',NULL,NULL From Dual
Union All Select 1,0,0,0,0,0,3,'急诊诊断输入',NULL,'1','记录急诊医生站首页、医嘱下达页面填写诊断时的诊断输入方式，下次进入时恢复上次选择的输入方式。','0-根据诊断标准输入,1-根据疾病编码输入',NULL,NULL,NULL From Dual
Union All Select 1,0,0,0,0,1,4,'病人查找方式',NULL,'0','记录急诊医生站查找病人的方式，下次进入时恢复上次选择的方式。','0-就诊卡,1-门诊号,2-挂号单,3-姓名查找,4-身份证,5-IC卡，6以后为医院其他医疗卡',NULL,NULL,NULL From Dual
Union All Select 1,0,0,0,0,0,6,'找到病人后自动接诊',NULL,'0','如果启用此参数，急诊医生站查找到候诊病人时，自动对此病人进行接诊操作，否则只是定位此病人。',NULL,'预约病人查找到后也会进行预约接收，但不受此参数影响。','为方便医生操作，如果医院未启用排队叫号时，可开启此参数，通过查找病人自动完成接诊。',NULL From Dual
Union All Select 1,0,0,0,0,0,7,'接诊后自动进行',NULL,'0','对病人接诊后，程序程序进行的处理。','0-不做任何操作，1-医嘱下达，切换到病历时病人没有病历再自动新增病历，2-新增病历，切换到医嘱时，如果病人还没有医嘱则自动新增病历。',NULL,'如果医生接诊后立即书写医嘱，则建议选择1-下达医嘱，如果立即书写病历，则建议选择2-新增病历，否则选择0-不做任何操作。',NULL From Dual
Union All Select 0,0,0,0,0,0,8,'只接收已经分诊的病人',NULL,'0','如果启用此参数，急诊医生站候诊列表中不显示未分诊的病人。',NULL,NULL,'如果医院允许不分诊直接就诊，则不启用此参数，否则启用。',NULL From Dual
Union All Select 0,1,0,0,0,0,9,'本地诊室',NULL,NULL,'当前急诊医生站如果设置接诊范围为：本诊室 时，则急诊医生站候诊列表只显示设置的分诊到该诊室的病人。','诊室名称','须设置关联参数：接诊范围=2-本诊室 时才有用。',NULL,NULL From Dual
Union All Select 0,1,0,0,0,0,10,'接诊范围',NULL,'2','1、如果设置为1-本人号，则医生站候诊列表只显示挂当前操作员的号的病人。'||CHR(13)||'2、如果设置为2-本诊室，则医生站候诊列表只显示分诊到当前诊室的病人。'||CHR(13)||'3、如果设置为3-本科室，则医生站候诊列表显示挂本科号的病人。','1-本人号,2-本诊室,3-本科室','如果选择2-本诊室，则跟参数：本地诊室 相关。'||CHR(13)||'如果选择3-本科室，则跟参数：接诊科室 相关。',NULL,NULL From Dual
Union All Select 0,1,0,0,0,0,11,'接诊科室',NULL,NULL,'当前急诊医生站如果设置接诊范围为：本科室时，则急诊医生站候诊列表只显示设置的挂号到该科室的病人。','科室ID','须设置关联参数：接诊范围=2-本诊室或3-本科室 时才有用。',NULL,NULL From Dual
Union All Select 1,1,0,0,0,0,12,'接诊医生',NULL,NULL,'1、急诊医生站就诊、回诊列表的只显示本参数设置的就诊医生正在就诊的病人。'||CHR(13)||'2、如果当前操作员不具有：急诊医生站.续诊病人 的权限，则此参数无效，就诊、回诊列表只显示当前操作员的病人。','医生姓名','与 急诊医生站.续诊病人 权限关联，没有此权限，则此参数无效。',NULL,NULL From Dual
Union All Select 1,1,0,1,0,1,13,'医护功能',NULL,NULL,'记录急诊医生站当前最后一次选择的功能卡片名称，例如：医嘱信息、病历信息，下次进入时恢复上次选择的状态。','卡片名称',NULL,NULL,NULL From Dual
Union All Select 1,1,0,0,0,0,14,'已诊病人结束间隔',NULL,'0','记录急诊医生站最后一次查找已诊病人的结束时间和当前时间的间隔天数，用于医嘱下达时复制病人医嘱页面显示相同时间范围内的已诊病人，再次进入医生站时不恢复。','间隔天数',NULL,'科室不同，或者目的不同，就会导致查找已诊病人设置时间范围不同，医生慢慢的就形成了个人喜好，在这里就增加了参数',NULL From Dual
Union All Select 1,1,0,0,0,0,15,'已诊病人开始间隔',NULL,'7','记录急诊医生站最后一次查找已诊病人的开始时间和当前时间的间隔天数，用于医嘱下达时复制病人医嘱页面显示相同时间范围内的已诊病人，再次进入医生站时不恢复。','间隔天数',NULL,NULL,NULL From Dual
Union All Select 1,1,0,1,0,1,17,'显示病历快捷输入',NULL,NULL,'如果启用此参数，急诊医生站在病人信息下方显示快捷病历快捷输入的输入框。',NULL,NULL,'有的医院，急诊使用了电子病历，然后以后又比较习惯使用快捷电子病历的方式写病历，而有的医院急诊没有使用电子病历，如果把电子病历快捷方式显示在主界面上就会比较占位置，不清爽，所以增加了参数，医院根据自身情况，设置参数',NULL From Dual
Union All Select 0,1,1,0,0,0,18,'医生就诊人数',NULL,'3','1、急诊医生站启用排队叫号之后，如果队列中已呼叫且正在就诊的病人超过此参数设置后，呼叫时提示已超过N个候诊病人，不允许再叫号;'||CHR(13)||'2、设置为0表示不显示候诊病人人数。','人数','1、此参数需要配合分诊台模块的排队叫号模式为１（医生主动呼叫）并且启用了参数：排队呼叫站点时有效。'||CHR(13)||'2、与参数：就诊人数含回诊，有关，如果启用含回诊，计算当前正在就诊的人数中包含回诊的病人。','避免同时就诊的人数过多的问题。',NULL From Dual
Union All Select 0,0,0,0,0,0,21,'就诊人数含回诊',NULL,'1','参数：医生就诊人数 的子参数，当设置了医生就诊人数时，判断当前正在就诊的人数时，回诊的病人也算是正在就诊的病人。',NULL,'本参数是参数：医生就诊人数 的子参数，当设置了医生就诊人数时，此参数时有效。','有的医院回诊病人不需要从新排队，做完检查检验后，直接将报告拿回找医生，也有的医院病人多，回诊病人多，初诊病人长时间看不到病产生矛盾，所以就将回诊病人纳入排队列表，还有遵义医院让贵阳公司的同事做了一个回诊病人跟初诊病人交叉呼叫的模式，设置参数，医院按照自身情况，选择处理方式',NULL From Dual
Union All Select 0,0,0,0,0,0,22,'医生主动呼叫后才允许接诊',NULL,'1','当启用了此参数，并且启用了排队叫号医生主动呼叫时，医生站工具栏中的接诊按钮不显示，只能通过排队叫号中的呼叫列表进行接诊。',NULL,'与分诊台参数：呼叫模式，有关，设置为医生主动呼叫时有效。','用于控制是否允许医生不通过排队呼叫就接诊病人（挂专家号插队的情况）',NULL From Dual
Union All Select 1,0,0,0,0,1,23,'字体','0','0','记录急诊医生站的字体大小，下次进入时恢复上次的大小。','9-小字体，12-大字体',NULL,'有的医生年龄大一些，就喜欢窗体的字显示的大一些，一些视力好的医生就喜欢字体小一些，界面看起来清爽一些，而且字体大小的喜好也关乎医生个人，所以此参数为私有全局参数，即医生设置一次后，不管在什么电脑上登录，都显示设置字体',NULL From Dual
Union All Select 1,0,0,0,0,0,24,'启用屏幕键盘','0','0','如果启用了屏幕键盘，急诊医生站 强制续诊、医嘱录入时，显示快捷屏幕键盘，可通过鼠标进行输入。',NULL,NULL,'急诊医生站有些老专家对电脑键盘使用不熟悉，则可启用屏幕键盘，实现鼠标化输入。',NULL From Dual
Union All Select 0,0,0,0,0,0,25,'病人接诊控制','0','0|0','如果挂号号别都是分时段挂号设置了预约时间的，急诊医生站接诊病人时，控制是否允许提前接收病人，如果设置为1或2，则可控制允许提前接收的时间范围。','0-不禁止;1-禁止;2-提示)|分钟数',NULL,'如果医院的挂号是分时段设置预约时间的，则可启用此参数，控制病人按预约时间进行就诊，未到预约时间不允许就诊，或提示医生。',NULL From Dual
Union All Select 1,0,0,0,0,0,27,'过敏输入来源','0','0','如果系统参数启用了：太元通合理用药接口，并且设置过敏输入来源=0-医生输入时决定，则急诊首页录入过敏信息时，可以选择按药品目录录入还是按过敏源录入，如果按药品目录录入，则使用HIS的药品目录，如果是过敏源录入，则调用太元通接口选择过敏源。','0-按药品目录输入,1-按过敏源输入','如果系统参数合理用药接口启用了：太元通合理用药接口，并且设置过敏输入来源=0-医生输入时决定时，此参数才有效。','根据医院对过敏信息录入的要求决定按哪种方式录入。',NULL From Dual
Union All Select 1,0,0,0,0,0,28,'自动刷新病历审阅间隔','0','0','如果未启用消息平台，则控制急诊医生站每N分钟刷新一次病历审阅提醒消息。','间隔分钟','当未启用了消息平台，此参数才有效。',NULL,NULL From Dual
Union All Select 1,0,0,0,0,0,29,'自动刷新内容','0','0','用于控制急诊医生站消息提醒列表中显示的需要提醒的消息类型。','每位数分别代表不同消息类型：1危急值、2医嘱安排、3处方审查、4传染病报告、5备血完成、6用血审核、7输血反应',NULL,'医生根据自己的需要，勾选需要提醒自己的消息。',NULL From Dual
Union All Select 1,0,0,0,0,0,30,'自动刷新病历审阅天数',NULL,'1','显示最近N天的消息息提醒','天数',NULL,NULL,NULL From Dual
Union All Select 1,0,0,0,0,0,31,'接诊时自动处理完成就诊',NULL,NULL,'急诊医生接诊病人时自动处理上一个病人完成就诊或需回诊。',NULL,NULL,NULL,NULL From Dual
Union All Select 1,1,0,0,0,0,33,'启用语音提示',NULL,'1','急诊医生站消息列表中的消息是否启用语音提醒。','0－不启用，1－启用。',NULL,'只有能在消息列表中显示出来的消息才会被播报。',NULL From Dual
Union All Select 1,1,0,0,0,0,34,'危机值消息语音配置',NULL,'1<sTab>0<sTab>iif([床号]<>"",[床号],"家庭床")+"有危机值消息。"<sTab>2','急诊医生站危机值消息语音配置信息。','固定格式：状态<sTab>提示方式<sTab>内容<sTab>播放次数。状态0/1是否启用，提示方式0/1 0读文本，1wav音频,内容－文本或音频文件，播放次数',NULL,'适用于急诊医生站危机值消息。',NULL From Dual
Union All Select 1,1,0,0,0,0,35,'安排消息语音配置',NULL,'1<sTab>0<sTab>iif([床号]<>"",[床号],"家庭床")+"有安排消息。"<sTab>2','急诊医生站安排消息语音配置信息。','固定格式：状态<sTab>提示方式<sTab>内容<sTab>播放次数。状态0/1是否启用，提示方式0/1 0读文本，1wav音频,内容－文本或音频文件，播放次数',NULL,'适用于急诊医生站安排消息。',NULL From Dual
Union All Select 1,1,0,0,0,0,36,'处方审查消息语音配置',NULL,'1<sTab>0<sTab>iif([床号]<>"",[床号],"家庭床")+"有处方审查消息。"<sTab>2','急诊医生站处方审查消息语音配置信息。','固定格式：状态<sTab>提示方式<sTab>内容<sTab>播放次数。状态0/1是否启用，提示方式0/1 0读文本，1wav音频,内容－文本或音频文件，播放次数',NULL,'适用于急诊医生站处方审查消息。',NULL From Dual
Union All Select 1,1,0,0,0,0,37,'传染病消息语音配置',NULL,'1<sTab>0<sTab>iif([床号]<>"",[床号],"家庭床")+"有传染病消息。"<sTab>2','急诊医生站传染病消息语音配置信息。','固定格式：状态<sTab>提示方式<sTab>内容<sTab>播放次数。状态0/1是否启用，提示方式0/1 0读文本，1wav音频,内容－文本或音频文件，播放次数',NULL,'适用于急诊医生站传染病消息。',NULL From Dual
Union All Select 1,1,0,0,0,0,38,'显示预约病人',NULL,'1','控制急诊医生工作站界面候诊列表中预约病人的显示，','0-不显示，1-显示。','急诊医生工作站病人列表显示','急诊医生工作站主界面候诊列表显示','当参数设为不显示时，预约病人要收费取挂号费后或者经分诊台处理变为正常挂号病人后才会显示到候诊表中。' From Dual
Union All Select 1,1,0,0,0,0,39,'急诊危急值弹窗提醒','1','1','控制急诊危急值提醒是否弹窗，','0-控制急诊危急值提醒不弹窗；1-控制急诊危急值弹窗提醒','','适用于用户想要控制急诊危急值弹窗提醒',Null From Dual
Union All Select 1,1,0,1,0,1,40,'急诊就诊信息折叠显示',NULL,NULL,'如果启用此参数，急诊医生站病人信息页签下就诊信息可以折叠显示。',NULL,NULL,'适用于急诊医生填写病人就诊信息时',NULL From Dual
Union All Select 1,1,0,1,0,1,41,'急诊基本信息折叠显示',NULL,NULL,'如果启用此参数，急诊医生站病人信息页签下基本信息可以折叠显示。',NULL,NULL,'适用于急诊医生填写病人基本信息时',NULL From Dual) A;

--急诊病人来源
Insert Into 急诊病人来源(编码,名称,缺省标志) Values('1','急救车首诊',0);
Insert Into 急诊病人来源(编码,名称,缺省标志) Values('2','急诊车转诊',0);
Insert Into 急诊病人来源(编码,名称,缺省标志) Values('3','非急诊救车首诊',1);
Insert Into 急诊病人来源(编码,名称,缺省标志) Values('4','门诊转诊',0);
Insert Into 急诊病人来源(编码,名称,缺省标志) Values('5','住院转诊',0);
Insert Into 急诊病人来源(编码,名称,缺省标志) Values('6','外院转诊',0);
Insert Into 急诊病人来源(编码,名称,缺省标志) Values('7','社区转诊',0);

--急诊常见既往史
Insert Into 急诊常见既往史(编码,名称) Values('1','糖尿病');
Insert Into 急诊常见既往史(编码,名称) Values('2','慢性阻塞性肺病');
Insert Into 急诊常见既往史(编码,名称) Values('3','冠心病');
Insert Into 急诊常见既往史(编码,名称) Values('4','高血压');
Insert Into 急诊常见既往史(编码,名称) Values('5','哮喘');

--急诊常用主诉
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('01',NULL,'呼吸系统','HXXT',0);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('02','01','呼吸短促','HXDC',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('03','01','呼吸停止','HXTZ',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('04','01','咳嗽','KS',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('05','01','换气过度','HQGD',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('06','01','咳血','KX',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('07','01','呼吸道内异物','HXDNYW',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('08','01','过敏反应','GMFY',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('09',NULL,'心血管系统','XXGXT',0);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('10','09','心跳停止','XTTZ',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('11','09','胸痛/胸闷','XT/XM',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('12','09','心慌（心悸）/不规则心跳','XHXJ/BGZXT',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('13','09','高血压','GXY',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('14','09','全身虚弱/无力','QSXR/WL',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('15','09','晕厥','YJ',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('16','09','全身性水肿','QSXSZ',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('17','09','肢体水肿','ZTSZ',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('18','09','冰冷无脉搏的肢体','BLWMBDZT',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('19','09','单侧肢体红热','DCZTHR',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('20',NULL,'消化系统','XHXT',0);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('21','20','腹痛','FT',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('22','20','厌食','YS',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('23','20','便秘','BM',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('24','20','腹泻','FX',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('25','20','直肠内异物','ZCNYW',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('26','20','腹股沟肿块','FGGZK',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('27','20','恶心呕吐','EXOT',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('28','20','直肠会阴疼痛','ZCHYTT',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('29','20','呕血','OX',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('30','20','血便/黑便','XB/HB',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('31','20','黄疸','HD',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('32','20','打嗝','DG',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('33','20','腹部肿块/腹胀','FBZK/FZ',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('34',NULL,'神经系统','SJXT',0);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('35','34','意识程度改变','YSCDGB',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('36','34','混乱（谵妄）','HLZW',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('37','34','眩晕/头晕','XY/TY',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('38','34','头痛','TT',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('39','34','抽搐','CC',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('40','34','步态失调/运动失调','BTSD/YDSD',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('41','34','震颤','ZC',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('42','34','肢体无力/中风症','ZTWL/ZFZ',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('43','34','知觉丧失/异常','ZJSS/YC',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('44',NULL,'骨骼系统','GGXT',0);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('45','44','背痛','BT',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('46','44','上肢疼痛','SZTT',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('47','44','下肢疼痛','XZTT',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('48','44','关节肿痛','GJZT',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('49',NULL,'泌尿系统','MNXT',0);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('50','49','腰痛','YT',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('51','49','血尿','XN',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('52','49','生殖器官分泌物/病变','SZQGFMW/BB',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('53','49','阴茎肿胀/疼痛','YJZZ/TT',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('54','49','尿潴留','NZL',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('55','49','泌尿道感染症状（尿频、尿急、尿痛）','MNDGRZZNPN',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('56','49','少尿','SN',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('57','49','多尿','DN',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('58',NULL,'一般和其它','YBHQT',0);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('59','58','发烧/畏寒','FS/WH',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('60','58','高血糖','GXT',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('61','58','低血糖','DXT',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('62','58','换药','HY',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('63','58','拆线','CX',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('64','58','要求开疾病诊断书','YQKJBZDS',1);
Insert Into 急诊常用主诉(编码,上级,名称,简码,末级) Values('65','58','要求开药','YQKY',1);

--急诊陪同人员
Insert Into 急诊陪同人员(编码,名称,缺省标志) Values('1','自行',0);
Insert Into 急诊陪同人员(编码,名称,缺省标志) Values('2','家属',1);
Insert Into 急诊陪同人员(编码,名称,缺省标志) Values('3','朋友',0);
Insert Into 急诊陪同人员(编码,名称,缺省标志) Values('4','民众',0);
Insert Into 急诊陪同人员(编码,名称,缺省标志) Values('5','110',0);
Insert Into 急诊陪同人员(编码,名称,缺省标志) Values('6','120(本院)',0);
Insert Into 急诊陪同人员(编码,名称,缺省标志) Values('7','120(他院)',0);
Insert Into 急诊陪同人员(编码,名称,缺省标志) Values('8','肇事者',0);
Insert Into 急诊陪同人员(编码,名称,缺省标志) Values('9','其他',0);


--急诊意识状态
Insert Into 急诊意识状态(编码,名称,缺省标志) Values('1','清醒',1);
Insert Into 急诊意识状态(编码,名称,缺省标志) Values('2','谵妄',0);
Insert Into 急诊意识状态(编码,名称,缺省标志) Values('3','嗜睡',0);
Insert Into 急诊意识状态(编码,名称,缺省标志) Values('4','昏睡',0);
Insert Into 急诊意识状态(编码,名称,缺省标志) Values('5','昏迷',0);

--------------------------------------------------------------------------------------------------------------------------
--急诊病情级别
Insert Into 急诊病情级别(序号,名称,严重程度,级别描述,响应要求说明,再次评估时限,患者标识颜色) Values(1,'1级','濒危','正在或即将发生的生命威胁 或病情恶化，需要立即进行积极 干预','立即进行评估和救治，安排患者进入复苏室或抢救室',NULL,'FF0000');
Insert Into 急诊病情级别(序号,名称,严重程度,级别描述,响应要求说明,再次评估时限,患者标识颜色) Values(2,'2级','危重','病情危重或迅速恶化，如短时间内不能进行治疗则危及生命或造成严重的器官功能衰竭；或者短时间内进行治疗可对预后产生重大影响','立即监护生命体征，10分钟内得到救治，安排患者进入抢救室',NULL,'FF6600');
Insert Into 急诊病情级别(序号,名称,严重程度,级别描述,响应要求说明,再次评估时限,患者标识颜色) Values(3,'3级','急症','存在潜在的生命威胁，如短时间内不进行干预，病情可进展至威胁生命或产生十分不利的结局；或者存在潜在的严重性，如患者一定时间内没有给予治疗，患者情况可能会恶化或出现不利的结局；以及症状将会加重或持续时间延长','先于4级患者优先诊治，安排患者在普通诊疗区候诊；若候诊时间大于30分钟，需再次评估',30,'FFFF00');
Insert Into 急诊病情级别(序号,名称,严重程度,级别描述,响应要求说明,再次评估时限,患者标识颜色) Values(4,'4级','非急症','慢性或非常轻微的症状，即便 等待一段时间再进行治疗也不会对结局产生大的影响','顺序就诊，除非病情变化，否则候诊时间较长；若候诊时间大于4小时，可再次评估',240,'00FF00');

--急诊评分方法
Insert Into 急诊评分方法(ID,英文名,中文名,说明) Values(1,NULL,'客观评估指标(成人)',NULL);
Insert Into 急诊评分方法(ID,英文名,中文名,说明) Values(2,NULL,'客观评估指标(儿童)',NULL);
Insert Into 急诊评分方法(ID,英文名,中文名,说明) Values(3,'GCS','格拉斯哥昏迷评分',NULL);
Insert Into 急诊评分方法(ID,英文名,中文名,说明) Values(4,'NRS','疼痛数字评分',NULL);
Insert Into 急诊评分方法(ID,英文名,中文名,说明) Values(5,NULL,'人工评定指标(成人)',NULL);
Insert Into 急诊评分方法(ID,英文名,中文名,说明) Values(6,NULL,'人工评定指标(儿童)',NULL);

--急诊评分方法分级
Insert Into 急诊评分方法分级(ID,方法ID,分值上限,分值下限,运算符,评分结果描述,病情级别) Values(1,3,15,15,1,'正常',4);
Insert Into 急诊评分方法分级(ID,方法ID,分值上限,分值下限,运算符,评分结果描述,病情级别) Values(2,3,14,12,6,'轻度意识障碍',4);
Insert Into 急诊评分方法分级(ID,方法ID,分值上限,分值下限,运算符,评分结果描述,病情级别) Values(3,3,11,9,6,'中度意识障碍',3);
Insert Into 急诊评分方法分级(ID,方法ID,分值上限,分值下限,运算符,评分结果描述,病情级别) Values(4,3,8,4,6,'昏迷',2);
Insert Into 急诊评分方法分级(ID,方法ID,分值上限,分值下限,运算符,评分结果描述,病情级别) Values(5,3,NULL,3,5,'深昏迷',1);
Insert Into 急诊评分方法分级(ID,方法ID,分值上限,分值下限,运算符,评分结果描述,病情级别) Values(6,4,0,0,1,'无痛',4);
Insert Into 急诊评分方法分级(ID,方法ID,分值上限,分值下限,运算符,评分结果描述,病情级别) Values(7,4,3,1,6,'轻度疼痛',4);
Insert Into 急诊评分方法分级(ID,方法ID,分值上限,分值下限,运算符,评分结果描述,病情级别) Values(8,4,6,4,6,'中度疼痛',3);
Insert Into 急诊评分方法分级(ID,方法ID,分值上限,分值下限,运算符,评分结果描述,病情级别) Values(9,4,10,7,6,'重度疼痛',2);

--急诊评分指标
Insert Into 急诊评分指标(ID,指标名称,值域类型,值域范围,值域单位,方法ID) Values(1,'睁眼反应',1,'自主|声音刺激|疼痛刺激|无',NULL,3);
Insert Into 急诊评分指标(ID,指标名称,值域类型,值域范围,值域单位,方法ID) Values(2,'语言反应',1,'定向良好|嗜睡|答非所问|不能理解|无',NULL,3);
Insert Into 急诊评分指标(ID,指标名称,值域类型,值域范围,值域单位,方法ID) Values(3,'活动反应',1,'活动自如|定位疼痛|躲避疼痛|痛刺激屈曲|痛刺激伸张|无',NULL,3);
Insert Into 急诊评分指标(ID,指标名称,值域类型,值域范围,值域单位,方法ID) Values(4,'收缩压',0,NULL,'mmHg',NULL);
Insert Into 急诊评分指标(ID,指标名称,值域类型,值域范围,值域单位,方法ID) Values(5,'舒张压',0,NULL,'mmHg',NULL);
Insert Into 急诊评分指标(ID,指标名称,值域类型,值域范围,值域单位,方法ID) Values(6,'心率',0,NULL,'次/分',NULL);
Insert Into 急诊评分指标(ID,指标名称,值域类型,值域范围,值域单位,方法ID) Values(7,'指氧饱和度',0,NULL,'%',NULL);
Insert Into 急诊评分指标(ID,指标名称,值域类型,值域范围,值域单位,方法ID) Values(8,'体温',0,NULL,'℃',NULL);
Insert Into 急诊评分指标(ID,指标名称,值域类型,值域范围,值域单位,方法ID) Values(9,'血糖',0,NULL,'mmol/L',NULL);
Insert Into 急诊评分指标(ID,指标名称,值域类型,值域范围,值域单位,方法ID) Values(10,'血钾',0,NULL,'mmol/L',NULL);
Insert Into 急诊评分指标(ID,指标名称,值域类型,值域范围,值域单位,方法ID) Values(11,'呼吸频率',0,NULL,'次/分',NULL);
Insert Into 急诊评分指标(ID,指标名称,值域类型,值域范围,值域单位,方法ID) Values(12,'疼痛描述',1,'无痛|安静平卧不痛，翻身咳嗽时疼痛|咳嗽疼痛，深呼吸不痛|安静平卧不痛，咳嗽深呼吸疼痛|安静平卧时，间歇疼痛|安静平卧时，持续疼痛|安静平卧时疼痛较重|疼痛较重，翻转不安，无法入睡|持续疼痛难忍，全身大汗|剧烈疼痛，无法忍受|最疼痛，生不如死',NULL,4);

--急诊评分指标年龄
Insert Into 急诊评分指标年龄(ID,指标ID,年龄上限,年龄下限,运算符,年龄单位,年龄段描述) Values(1,6,3,0,6,'月','0～3月');
Insert Into 急诊评分指标年龄(ID,指标ID,年龄上限,年龄下限,运算符,年龄单位,年龄段描述) Values(2,6,6,3.1,6,'月','～6月');
Insert Into 急诊评分指标年龄(ID,指标ID,年龄上限,年龄下限,运算符,年龄单位,年龄段描述) Values(3,6,12,6.1,6,'月','～12月');
Insert Into 急诊评分指标年龄(ID,指标ID,年龄上限,年龄下限,运算符,年龄单位,年龄段描述) Values(4,6,3,1.1,6,'岁','～3岁');
Insert Into 急诊评分指标年龄(ID,指标ID,年龄上限,年龄下限,运算符,年龄单位,年龄段描述) Values(5,6,6,3.1,6,'岁','～6岁');
Insert Into 急诊评分指标年龄(ID,指标ID,年龄上限,年龄下限,运算符,年龄单位,年龄段描述) Values(6,6,10,6.1,6,'岁','～10岁');
Insert Into 急诊评分指标年龄(ID,指标ID,年龄上限,年龄下限,运算符,年龄单位,年龄段描述) Values(7,11,3,0,6,'月','0～3月');
Insert Into 急诊评分指标年龄(ID,指标ID,年龄上限,年龄下限,运算符,年龄单位,年龄段描述) Values(8,11,6,3.1,6,'月','～6月');
Insert Into 急诊评分指标年龄(ID,指标ID,年龄上限,年龄下限,运算符,年龄单位,年龄段描述) Values(9,11,12,6.1,6,'月','～12月');
Insert Into 急诊评分指标年龄(ID,指标ID,年龄上限,年龄下限,运算符,年龄单位,年龄段描述) Values(10,11,3,1.1,6,'岁','～3岁');
Insert Into 急诊评分指标年龄(ID,指标ID,年龄上限,年龄下限,运算符,年龄单位,年龄段描述) Values(11,11,6,3.1,6,'岁','～6岁');
Insert Into 急诊评分指标年龄(ID,指标ID,年龄上限,年龄下限,运算符,年龄单位,年龄段描述) Values(12,11,10,6.1,6,'岁','～10岁');
Insert Into 急诊评分指标年龄(ID,指标ID,年龄上限,年龄下限,运算符,年龄单位,年龄段描述) Values(13,4,31,0,6,'天','0～31天');
Insert Into 急诊评分指标年龄(ID,指标ID,年龄上限,年龄下限,运算符,年龄单位,年龄段描述) Values(14,4,12,1,6,'月','1～12 月');
Insert Into 急诊评分指标年龄(ID,指标ID,年龄上限,年龄下限,运算符,年龄单位,年龄段描述) Values(15,4,2,1,6,'岁','1岁');
Insert Into 急诊评分指标年龄(ID,指标ID,年龄上限,年龄下限,运算符,年龄单位,年龄段描述) Values(16,4,3,2.1,6,'岁','2岁');
Insert Into 急诊评分指标年龄(ID,指标ID,年龄上限,年龄下限,运算符,年龄单位,年龄段描述) Values(17,4,4,3.1,6,'岁','3岁');
Insert Into 急诊评分指标年龄(ID,指标ID,年龄上限,年龄下限,运算符,年龄单位,年龄段描述) Values(18,4,5,4.1,6,'岁','4岁');
Insert Into 急诊评分指标年龄(ID,指标ID,年龄上限,年龄下限,运算符,年龄单位,年龄段描述) Values(19,4,6,5.1,6,'岁','5岁');
Insert Into 急诊评分指标年龄(ID,指标ID,年龄上限,年龄下限,运算符,年龄单位,年龄段描述) Values(20,4,7,6.1,6,'岁','6岁');
Insert Into 急诊评分指标年龄(ID,指标ID,年龄上限,年龄下限,运算符,年龄单位,年龄段描述) Values(21,4,8,7.1,6,'岁','7岁');
Insert Into 急诊评分指标年龄(ID,指标ID,年龄上限,年龄下限,运算符,年龄单位,年龄段描述) Values(22,4,9,8.1,6,'岁','8岁');
Insert Into 急诊评分指标年龄(ID,指标ID,年龄上限,年龄下限,运算符,年龄单位,年龄段描述) Values(23,4,10,9.1,6,'岁','9岁');
Insert Into 急诊评分指标年龄(ID,指标ID,年龄上限,年龄下限,运算符,年龄单位,年龄段描述) Values(24,4,NULL,10,3,'岁','10岁');
Insert Into 急诊评分指标年龄(ID,指标ID,年龄上限,年龄下限,运算符,年龄单位,年龄段描述) Values(25,8,12,3,6,'月','3～12个月');
Insert Into 急诊评分指标年龄(ID,指标ID,年龄上限,年龄下限,运算符,年龄单位,年龄段描述) Values(26,8,3,1,6,'月','1～3个月');
Insert Into 急诊评分指标年龄(ID,指标ID,年龄上限,年龄下限,运算符,年龄单位,年龄段描述) Values(27,8,12,1,6,'岁','1～12岁');
Insert Into 急诊评分指标年龄(ID,指标ID,年龄上限,年龄下限,运算符,年龄单位,年龄段描述) Values(28,8,31,0,6,'天','0～31天');
Insert Into 急诊评分指标年龄(ID,指标ID,年龄上限,年龄下限,运算符,年龄单位,年龄段描述) Values(29,6,31,0,6,'天','0～31天');
Insert Into 急诊评分指标年龄(ID,指标ID,年龄上限,年龄下限,运算符,年龄单位,年龄段描述) Values(30,11,31,0,6,'天','0～31天');

--急诊评分方法规则
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(1,4,12,NULL,NULL,NULL,NULL,0,'无痛',NULL);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(2,4,12,NULL,NULL,NULL,NULL,1,'安静平卧不痛，翻身咳嗽时疼痛',NULL);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(3,4,12,NULL,NULL,NULL,NULL,2,'咳嗽疼痛，深呼吸不痛',NULL);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(4,4,12,NULL,NULL,NULL,NULL,3,'安静平卧不痛，咳嗽深呼吸疼痛',NULL);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(5,4,12,NULL,NULL,NULL,NULL,4,'安静平卧时，间歇疼痛',NULL);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(6,4,12,NULL,NULL,NULL,NULL,5,'安静平卧时，持续疼痛',NULL);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(7,4,12,NULL,NULL,NULL,NULL,6,'安静平卧时疼痛较重',NULL);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(8,4,12,NULL,NULL,NULL,NULL,7,'疼痛较重，翻转不安，无法入睡',NULL);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(9,4,12,NULL,NULL,NULL,NULL,8,'持续疼痛难忍，全身大汗',NULL);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(10,4,12,NULL,NULL,NULL,NULL,9,'剧烈疼痛，无法忍受',NULL);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(11,4,12,NULL,NULL,NULL,NULL,10,'最疼痛，生不如死',NULL);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(12,3,1,NULL,NULL,NULL,NULL,4,'自主',NULL);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(13,3,1,NULL,NULL,NULL,NULL,3,'声音刺激',NULL);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(14,3,1,NULL,NULL,NULL,NULL,2,'疼痛刺激',NULL);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(15,3,1,NULL,NULL,NULL,NULL,1,'无',NULL);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(16,3,2,NULL,NULL,NULL,NULL,5,'定向良好',NULL);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(17,3,2,NULL,NULL,NULL,NULL,4,'嗜睡',NULL);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(18,3,2,NULL,NULL,NULL,NULL,3,'答非所问',NULL);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(19,3,2,NULL,NULL,NULL,NULL,2,'不能理解',NULL);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(20,3,2,NULL,NULL,NULL,NULL,1,'无',NULL);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(21,3,3,NULL,NULL,NULL,NULL,6,'活动自如',NULL);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(22,3,3,NULL,NULL,NULL,NULL,5,'定位疼痛',NULL);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(23,3,3,NULL,NULL,NULL,NULL,4,'躲避疼痛',NULL);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(24,3,3,NULL,NULL,NULL,NULL,3,'痛刺激屈曲',NULL);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(25,3,3,NULL,NULL,NULL,NULL,2,'痛刺激伸张',NULL);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(26,3,3,NULL,NULL,NULL,NULL,1,'无',NULL);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(27,1,4,NULL,NULL,70,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(28,1,6,NULL,180,NULL,2,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(29,1,6,NULL,NULL,40,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(30,1,7,NULL,NULL,80,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(31,1,8,NULL,41,NULL,2,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(32,1,9,NULL,NULL,3.33,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(33,1,10,NULL,7,NULL,2,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(34,1,4,NULL,200,NULL,2,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(35,1,4,NULL,80,70,6,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(36,1,6,NULL,180,150,6,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(37,1,6,NULL,50,40,6,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(38,1,7,NULL,90,80,6,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(39,1,4,NULL,200,180,6,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(40,1,4,NULL,90,80,6,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(41,1,6,NULL,150,100,6,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(42,1,6,NULL,55,50,6,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(43,1,7,NULL,94,90,6,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(44,1,NULL,NULL,NULL,NULL,NULL,NULL,'生命体征平稳',4);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(45,2,4,13,NULL,60,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(46,2,4,14,NULL,70,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(47,2,4,15,NULL,72,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(48,2,4,16,NULL,74,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(49,2,4,17,NULL,76,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(50,2,4,18,NULL,78,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(51,2,4,19,NULL,80,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(52,2,4,20,NULL,82,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(53,2,4,21,NULL,84,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(54,2,4,22,NULL,86,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(55,2,4,23,NULL,88,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(56,2,4,24,NULL,90,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(57,2,4,25,NULL,90,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(58,2,7,NULL,NULL,85,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(59,2,8,NULL,NULL,35,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(60,2,7,NULL,89,85,6,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(61,2,8,25,40,NULL,2,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(62,2,8,26,39,NULL,2,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(63,2,8,27,40,NULL,2,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(64,2,8,28,39,NULL,2,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(65,2,4,NULL,140,NULL,2,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(66,2,5,NULL,90,NULL,2,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(67,2,7,NULL,94,90,6,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(68,2,NULL,NULL,NULL,NULL,NULL,NULL,'生命体征平稳',4);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(69,2,6,29,NULL,40,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(70,2,6,1,NULL,40,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(71,2,6,2,NULL,40,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(72,2,6,3,NULL,40,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(73,2,6,4,NULL,40,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(74,2,6,5,NULL,40,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(75,2,6,6,NULL,30,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(76,2,6,1,70,40,6,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(77,2,6,2,65,40,6,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(78,2,6,3,63,40,6,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(79,2,6,4,60,40,6,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(80,2,6,5,55,40,6,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(81,2,6,6,45,30,6,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(82,2,6,1,90,70,6,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(83,2,6,2,80,65,6,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(84,2,6,3,80,63,6,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(85,2,6,4,75,60,6,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(86,2,6,5,70,55,6,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(87,2,6,6,60,45,6,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(88,2,6,1,180,90,6,NULL,NULL,4);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(89,2,6,2,160,80,6,NULL,NULL,4);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(90,2,6,3,140,80,6,NULL,NULL,4);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(91,2,6,4,130,75,6,NULL,NULL,4);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(92,2,6,5,110,70,6,NULL,NULL,4);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(93,2,6,6,90,60,6,NULL,NULL,4);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(94,2,6,1,205,180,6,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(95,2,6,2,180,160,6,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(96,2,6,3,160,140,6,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(97,2,6,4,145,130,6,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(98,2,6,5,125,110,6,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(99,2,6,6,110,90,6,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(100,2,6,1,230,205,6,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(101,2,6,2,210,180,6,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(102,2,6,3,180,160,6,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(103,2,6,4,165,145,6,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(104,2,6,5,140,125,6,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(105,2,6,6,130,110,6,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(106,2,6,1,230,NULL,2,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(107,2,6,29,230,NULL,2,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(108,2,6,2,210,NULL,2,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(109,2,6,3,180,NULL,2,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(110,2,6,4,165,NULL,2,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(111,2,6,5,140,NULL,2,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(112,2,6,6,130,NULL,2,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(113,2,11,30,NULL,10,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(114,2,11,7,NULL,10,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(115,2,11,8,NULL,10,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(116,2,11,9,NULL,10,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(117,2,11,10,NULL,10,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(118,2,11,11,NULL,8,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(119,2,11,12,NULL,8,3,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(120,2,11,7,20,10,6,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(121,2,11,8,20,10,6,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(122,2,11,9,17,10,6,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(123,2,11,10,15,10,6,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(124,2,11,11,12,8,6,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(125,2,11,12,10,8,6,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(126,2,11,7,30,20,6,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(127,2,11,8,30,20,6,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(128,2,11,9,25,17,6,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(129,2,11,10,20,15,6,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(130,2,11,11,16,12,6,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(131,2,11,12,14,10,6,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(132,2,11,7,60,30,6,NULL,NULL,4);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(133,2,11,8,60,30,6,NULL,NULL,4);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(134,2,11,9,45,25,6,NULL,NULL,4);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(135,2,11,10,35,20,6,NULL,NULL,4);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(136,2,11,11,25,16,6,NULL,NULL,4);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(137,2,11,12,20,14,6,NULL,NULL,4);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(138,2,11,7,70,60,6,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(139,2,11,8,70,60,6,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(140,2,11,9,55,45,6,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(141,2,11,10,40,35,6,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(142,2,11,11,30,25,6,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(143,2,11,12,25,20,6,NULL,NULL,3);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(144,2,11,7,80,70,6,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(145,2,11,8,80,70,6,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(146,2,11,9,60,55,6,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(147,2,11,10,45,40,6,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(148,2,11,11,35,30,6,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(149,2,11,12,30,25,6,NULL,NULL,2);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(150,2,11,30,80,NULL,2,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(151,2,11,7,80,NULL,2,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(152,2,11,8,80,NULL,2,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(153,2,11,9,60,NULL,2,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(154,2,11,10,45,NULL,2,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(155,2,11,11,35,NULL,2,NULL,NULL,1);
Insert Into 急诊评分方法规则(ID,方法ID,指标ID,指标年龄ID,指标值上限,指标值下限,运算符,指标结果分值,指标结果描述,病情级别) Values(156,2,11,12,30,NULL,2,NULL,NULL,1);


--急诊人工评定规则
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(1,'呼吸','气道不能维持','成人',1);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(2,'呼吸','气道风险:严重呼吸困难/气道不能保护','成人',2);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(3,'呼吸','急性哮喘，但血压、脉搏稳定','成人',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(4,'呼吸','吸入异物，无呼吸困难','成人',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(5,'呼吸','吞咽困难，无呼吸困难','成人',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(6,'循环','心博/呼吸停止或节律不稳定','成人',1);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(7,'循环','明确心肌梗死','成人',1);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(8,'循环','循环障碍，皮肤湿冷花斑，灌注差/怀疑脓毒症','成人',2);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(9,'循环','急性脑卒中','成人',2);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(10,'循环','类似心脏因素的胸痛','成人',2);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(11,'循环','活动性或严重失血','成人',2);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(12,'循环','轻微出血','成人',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(13,'神经','休克','成人',1);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(14,'神经','急性意识障碍/无反应或仅有疼痛刺激反应（GCS < 9）','成人',1);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(15,'神经','癫痫持续状态','成人',1);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(16,'神经','间断癫痫发作','成人',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(17,'神经','严重的精神行为异常，正在进行的自伤或他伤行为，需立即药物控制者','成人',1);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(18,'神经','昏睡（强烈刺激下有防御反应）','成人',1);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(19,'神经','严重的精神行为异常（暴力或攻击），直接威胁自身或他人， 需要被约束','成人',2);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(20,'神经','嗜睡（可唤醒，无刺激情况下转入睡眠）','成人',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(21,'神经','精神行为异常：有自残风险/急性精神错乱或思维混乱/焦虑/抑郁/潜在的攻击性','成人',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(22,'神经','轻微的精神行为异常','成人',4);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(23,'创伤','复合伤（需要快速团队应对）','成人',1);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(24,'创伤','严重的局部创伤-大的骨折、截肢','成人',2);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(25,'创伤','头外伤','成人',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(26,'创伤','中等程度外伤，肢体感觉运动异常','成人',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(27,'创伤','无肋骨疼痛或呼吸困难的胸部损伤','成人',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(28,'创伤','轻微头部损伤，无意识丧失','成人',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(29,'创伤','小的肢体创伤，生命体征正常，轻中度疼痛','成人',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(30,'创伤','微小伤口-不需要缝合的小的擦伤、裂伤','成人',4);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(31,'中毒','过量接触或摄入药物、毒物、化学物质、放射物质等','成人',2);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(32,'疼痛','所有原因所致严重疼痛（7~10分）','成人',1);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(33,'疼痛','不明原因的严重疼痛伴大汗（脐以上）','成人',2);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(34,'疼痛','胸腹疼痛，已有证据表明或高度怀疑以下疾病：急性心梗、急性肺栓塞、主动脉夹层、主动脉瘤、急性心肌炎/心包炎、心包积液、异位妊娠、消化道穿孔、睾丸扭转','成人',2);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(35,'疼痛','中等程度的非心源性胸痛','成人',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(36,'疼痛','中等程度或年龄>65岁无高危因素的腹痛','成人',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(37,'疼痛','任何原因出现的中重度疼痛，需要止疼（4～6分）','成人',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(38,'疼痛','中等程度疼痛，有一些危险特征','成人',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(39,'疼痛','非特异性轻度腹痛','成人',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(40,'疼痛','关节热胀，轻度肿痛','成人',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(41,'疼痛','无危险特征的微疼痛','成人',4);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(42,'其他','其他危及生命、需要紧急抢救的情况','成人',1);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(43,'其他','急性药物过量','成人',1);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(44,'其他','其他存在高风险、可能进展至危及生命或致残的情况','成人',2);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(45,'其他','持续呕吐/脱水','成人',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(46,'其他','病情稳定，症状轻微','成人',4);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(47,'其他','低危病史且目前无症状或症状轻微','成人',4);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(48,'其他','熟悉的有慢性症状患者','成人',4);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(49,'其他','稳定恢复期或无症状患者复诊/仅开药','成人',4);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(50,'其他','仅开具医疗证明','成人',4);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(51,'呼吸','严重呼吸窘迫','儿童',1);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(52,'呼吸','气道阻塞/窒息需要紧急气管插管或气管切开','儿童',1);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(53,'呼吸','气道风险:严重呼吸困难/气道不能保护','儿童',2);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(54,'呼吸','急性哮喘重度发作','儿童',2);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(55,'呼吸','胸痛/胸闷（疑张力性气胸、肺栓塞）','儿童',2);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(56,'呼吸','急性哮喘发作','儿童',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(57,'呼吸','发热、咳嗽、咽痛等轻微症状','儿童',4);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(58,'循环','心博/呼吸停止或节律不稳定','儿童',1);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(59,'循环','严重心律失常','儿童',1);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(60,'循环','急性溶血性贫血（重度）','儿童',2);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(61,'循环','高血压危象','儿童',2);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(62,'循环','暴发性心肌炎','儿童',2);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(63,'循环','严重的活动性失血','儿童',2);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(64,'循环','低血糖发作','儿童',2);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(65,'循环','任何原因导致的中度失血','儿童',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(66,'循环','出血性疾病或凝血功能异常','儿童',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(67,'神经','休克','儿童',1);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(68,'神经','急性意识障碍（GCS＜9分）','儿童',2);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(69,'神经','意识状态改变（GCS≥9分/非急性）','儿童',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(70,'神经','昏迷/昏睡','儿童',1);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(71,'神经','惊厥持续状态','儿童',1);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(72,'神经','抽搐发作','儿童',2);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(73,'神经','间断抽搐发作','儿童',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(74,'神经','中枢神经系统感染','儿童',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(75,'创伤','严重复合伤、大面积烧伤（需要快速团队应对）','儿童',1);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(76,'创伤','严重的局部创伤','儿童',2);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(77,'创伤','急性脊髓损伤','儿童',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(78,'创伤','肢体感觉、运动异常','儿童',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(79,'创伤','中等程度外伤','儿童',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(80,'创伤','微小伤口-不需要缝合的小的擦伤、裂伤','儿童',4);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(81,'中毒','急性中毒危及生命','儿童',1);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(82,'中毒','不危及生命的急性中毒','儿童',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(83,'中毒','糖尿病酮症酸中毒','儿童',2);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(84,'中毒','吸入或经消化道摄入过量药物、毒物、化学物质等','儿童',2);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(85,'疼痛','非特异性的轻度腹痛','儿童',4);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(86,'疼痛','生命体征异常的腹痛，已有证据表明或高度怀疑以下疾病：消化道穿孔、睾丸扭转等','儿童',2);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(87,'其他','其他危及生命、需要紧急抢救的情况','儿童',1);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(88,'其他','其他存在高风险、可能迅速进展至危及生命或致残的情况','儿童',2);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(89,'其他','重度脱水','儿童',2);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(90,'其他','中度脱水','儿童',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(91,'其他','急性出血性皮疹（暴发性紫癜）','儿童',2);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(92,'其他','严重电解质紊乱','儿童',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(93,'其他','生命体征平稳的新生儿','儿童',3);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(94,'其他','不伴脱水/轻度脱水的腹泻、呕吐','儿童',4);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(95,'其他','病情稳定，症状轻微','儿童',4);
Insert Into 急诊人工评定规则(ID,分类,指标名称,适用人群,病情级别) Values(96,'其他','低危病史且目前无症状或症状轻微','儿童',4);

-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------

--145003:蒋廷中,2019-10-14,新增模块急诊预检分诊工作站
Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select 100,1244,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0
Union All Select '基本',1,'急诊预检分诊工作站基本权限',1 From Dual) A;


Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select 100,1244,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All
Select 'Pkg_Pretriage_Dml','EXECUTE' From Dual Union All 
Select 'Pkg_Pretriage_Dql','EXECUTE' From Dual Union All 
Select '病人信息','SELECT' From Dual Union All 
Select '急诊就诊记录','SELECT' From Dual Union All 
Select '急诊分诊记录','SELECT' From Dual Union All 
Select '急诊陪同人员','SELECT' From Dual Union All 
Select '临床部门','SELECT' From Dual Union All 
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;



--145003:张永康,2019-10-14,新增模块急诊医生站
Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select 100,1243,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0
Union All Select '基本',-Null,NULL,1 From Dual
Union All Select '病人接诊',3,'接待已挂号的病人的操作权限。有该权限时，允许接诊已挂号的病人',1 From Dual
Union All Select '病人转诊',4,'把待诊或就诊的病人转诊到其它科室的权限。有该权限时，允许把待诊或就诊的病人转诊到其它科室',1 From Dual
Union All Select '续诊病人',5,'允许对本科内其他医生接诊的病人进行续诊的权限。',1 From Dual
Union All Select '急诊首页',6,'急诊病案首页信息(病人信息和过敏药物)登记。有该权限时，允许登记急诊病案首页信息',1 From Dual
Union All Select '参数设置',8,'设置急诊医生工作站运行参数的权限。有该权限时，允许进行本地参数设置',1 From Dual
Union All Select '全院病人续诊',11,'允许对全院其他科室的病人进行强制续诊的权限',1 From Dual
Union All Select '所有操作员',12,'允许对所有医生的就诊病人进行查看的权限',1 From Dual
Union All Select '已下医嘱转诊',13,'允许对已下达医嘱的病人进行转诊操作。',1 From Dual
Union All Select '诊疗一览',17,'有该权限时，急诊医生站才显示诊疗一览页签。',1 From Dual
Union All Select '修改医疗付款方式',18,'没有该权限，在急诊医生工作站医疗付款方式不能修改。',0 From Dual
Union All Select '修改费别',19,'没有该权限，在急诊医生工作站费别不能修改。',0 From Dual
Union All Select '操作其他医生的病人',20,'有权限时，在急诊医生工作站当前医生可以操作由其它医生接诊的病人。',0 From Dual
Union All Select '允许强制续诊正在就诊的病人',21,'有权限时使用强制续诊功能时，可以续诊正在就诊的病人。',0 From Dual
Union All Select '代办人信息允许自由录入',22,'有该权限时，代办人信息允许自由录入。',1 From Dual
Union All Select '危急值处理',23,'有该权限时，急诊医生站才处理病人危急值记录。',1 From Dual
Union All Select '允许设置接诊医生',24,'有该权限且拥有续诊病人权限时，允许操作员修改急诊工作站的接诊医生',1 From Dual) A;

Insert Into zlModuleRelas(系统,模块,功能,相关系统,相关模块,相关类型,相关功能,缺省值)
Select 100,1243,A.* From (
Select 功能,相关系统,相关模块,相关类型,相关功能,缺省值 From zlModuleRelas Where 1 = 0
Union All Select NULL,100,1070,0,'签名权',1 From Dual
Union All Select NULL,100,1153,1,'基本',1 From Dual
Union All Select NULL,100,1160,0,'基本',1 From Dual
Union All Select NULL,100,1250,0,NULL,1 From Dual
Union All Select NULL,100,1252,0,NULL,1 From Dual
Union All Select NULL,100,1259,0,NULL,1 From Dual
Union All Select NULL,100,1266,0,NULL,0 From Dual
Union All Select NULL,100,1268,0,NULL,0 From Dual
Union All Select NULL,100,1270,0,NULL,1 From Dual
Union All Select NULL,100,1271,0,NULL,1 From Dual
Union All Select NULL,100,2251,0,NULL,1 From Dual
Union All Select NULL,100,9000,0,'挂号',1 From Dual
Union All Select NULL,100,9000,0,'挂号病人建档',1 From Dual
Union All Select NULL,100,9000,0,'挂号费别打折',1 From Dual
Union All Select NULL,100,9000,0,'挂号选项设置',1 From Dual
Union All Select NULL,100,9000,0,'预约',1 From Dual
Union All Select NULL,100,9003,0,'基本信息调整',1 From Dual
Union All Select NULL,100,9003,1,'基本',1 From Dual) A;
Insert Into zlProgRelas(系统,序号,组号,功能,关系,主项,主项关系)
Select 100,1243,2,A.* From (
Select 功能,关系,主项,主项关系 From zlProgRelas Where 1 = 0
Union All Select '病人接诊',2,1,0 From Dual
Union All Select '续诊病人',2,0,0 From Dual) A;
Insert Into zlProgRelas(系统,序号,组号,功能,关系,主项,主项关系)
Select 100,1243,3,A.* From (
Select 功能,关系,主项,主项关系 From zlProgRelas Where 1 = 0
Union All Select '全院病人续诊',2,0,0 From Dual
Union All Select '续诊病人',2,1,0 From Dual) A;
Insert Into zlProgRelas(系统,序号,组号,功能,关系,主项,主项关系)
Select 100,1243,4,A.* From (
Select 功能,关系,主项,主项关系 From zlProgRelas Where 1 = 0
Union All Select '病人转诊',2,1,0 From Dual
Union All Select '已下医嘱转诊',2,0,0 From Dual) A;
Insert Into zlProgRelas(系统,序号,组号,功能,关系,主项,主项关系)
Select 100,1243,5,A.* From (
Select 功能,关系,主项,主项关系 From zlProgRelas Where 1 = 0 
Union All Select '允许设置接诊医生',2,0,0 From Dual
Union All Select '续诊病人',2,1,0 From Dual) A;




--1243:急诊医生工作站(基本)
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select 100,1243,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0
Union All Select 'ZL_FUN_GET检验图像','EXECUTE' From Dual
Union All Select 'ZL_病人发卡记录_换补卡','EXECUTE' From Dual
Union All Select 'ZL_病人发卡记录_上传','EXECUTE' From Dual
Union All Select 'ZL_病人危急值记录_UPDATE','EXECUTE' From Dual
Union All Select 'ZL_病人信息_Insert','EXECUTE' From Dual
Union All Select 'ZL_病人信息_UPDATE','EXECUTE' From Dual
Union All Select 'ZL_病人信息_更新信息','EXECUTE' From Dual
Union All Select 'ZL_病人信息从表_UPDATE','EXECUTE' From Dual
Union All Select 'Zl_Age_Calc','EXECUTE' From Dual
Union All Select 'Zl_Fun_Getsignpar','EXECUTE' From Dual
Union All Select 'Zl_Lob_Append','EXECUTE' From Dual
Union All Select 'Zl_Lob_Read','EXECUTE' From Dual
Union All Select 'Zl_Paticardcheck','Execute' From Dual
Union All Select 'Zl_Regist_AutoIntoblacklist','Execute' From Dual
Union All Select 'Zl_病人挂号记录_更新费别','EXECUTE' From Dual
Union All Select 'Zl_病人挂号记录_状态','EXECUTE' From Dual
Union All Select 'Zl_病人过敏药物_DELETE','EXECUTE' From Dual
Union All Select 'Zl_病人过敏药物_UPDATE','EXECUTE' From Dual
Union All Select 'Zl_病人危急值记录_DELETE','EXECUTE' From Dual
Union All Select 'Zl_病人危急值记录_Insert','EXECUTE' From Dual
Union All Select 'Zl_病人危急值记录_处理','EXECUTE' From Dual
Union All Select 'Zl_病人危急值医嘱_Update','EXECUTE' From Dual
Union All Select 'Zl_疾病申报记录_Update','EXECUTE' From Dual
Union All Select 'Zl_疾病阳性检测记录_Update','EXECUTE' From Dual
Union All Select 'Zl_门诊生命体征_UPDATE','EXECUTE' From Dual
Union All Select 'ZL_病人地址信息_UPDATE','EXECUTE' From Dual
Union All Select '保险病种','SELECT' From Dual
Union All Select '保险参数','SELECT' From Dual
Union All Select '保险结算记录','SELECT' From Dual
Union All Select '保险类别','SELECT' From Dual
Union All Select '保险特准项目','SELECT' From Dual
Union All Select '保险项目','SELECT' From Dual
Union All Select '保险帐户','SELECT' From Dual
Union All Select '保险支付大类','SELECT' From Dual
Union All Select '保险支付项目','SELECT' From Dual
Union All Select '保险中心目录','SELECT' From Dual
Union All Select '病案主页从表','SELECT' From Dual
Union All Select '病历单据应用','SELECT' From Dual
Union All Select '病历范文目录','SELECT' From Dual
Union All Select '病历范文内容','SELECT' From Dual
Union All Select '病历文件列表','SELECT' From Dual
Union All Select '病人发卡记录','SELECT' From Dual
Union All Select '病人挂号汇总','SELECT' From Dual
Union All Select '病人挂号记录','SELECT' From Dual
Union All Select '病人过敏记录','SELECT' From Dual
Union All Select '病人过敏药物','SELECT' From Dual
Union All Select '病人护理记录','SELECT' From Dual
Union All Select '病人护理内容','SELECT' From Dual
Union All Select '病人来源','SELECT' From Dual
Union All Select '病人类型','SELECT' From Dual
Union All Select '病人去向','SELECT' From Dual
Union All Select '病人社区信息','SELECT' From Dual
Union All Select '病人危急值病历','SELECT' From Dual
Union All Select '病人危急值记录','SELECT' From Dual
Union All Select '病人危急值记录_ID','SELECT' From Dual
Union All Select '病人危急值医嘱','SELECT' From Dual
Union All Select '病人信息','SELECT' From Dual
Union All Select '病人信息从表','SELECT' From Dual
Union All Select '病人医疗卡属性','SELECT' From Dual
Union All Select '病人医疗卡信息','SELECT' From Dual
Union All Select '病人医嘱记录','SELECT' From Dual
Union All Select '病人余额','SELECT' From Dual
Union All Select '病人预交记录','SELECT' From Dual
Union All Select '病人照片','Select' From Dual
Union All Select '病人诊断记录','SELECT' From Dual
Union All Select '部门性质说明','SELECT' From Dual
Union All Select '材料特性','SELECT' From Dual
Union All Select '常用挂号摘要','SELECT' From Dual
Union All Select '常用就诊摘要','SELECT' From Dual
Union All Select '床位状况记录','SELECT' From Dual
Union All Select '大类档次比例','SELECT' From Dual
Union All Select '单据操作控制','SELECT' From Dual
Union All Select '地区','SELECT' From Dual
Union All Select '电子病历记录','SELECT' From Dual
Union All Select '电子病历内容','SELECT' From Dual
Union All Select '费别','SELECT' From Dual
Union All Select '费别明细','SELECT' From Dual
Union All Select '挂号安排','SELECT' From Dual
Union All Select '挂号安排计划','SELECT' From Dual
Union All Select '挂号安排时段','SELECT' From Dual
Union All Select '挂号安排停用状态','SELECT' From Dual
Union All Select '挂号安排限制','SELECT' From Dual
Union All Select '挂号安排诊室','SELECT' From Dual
Union All Select '挂号计划时段','SELECT' From Dual
Union All Select '挂号计划限制','SELECT' From Dual
Union All Select '挂号项目','SELECT' From Dual
Union All Select '挂号序号状态','SELECT' From Dual
Union All Select '国籍','SELECT' From Dual
Union All Select '过敏源','SELECT' From Dual
Union All Select '号码控制表','SELECT' From Dual
Union All Select '合约单位','SELECT' From Dual
Union All Select '合作单位安排控制','SELECT' From Dual
Union All Select '合作单位计划控制','SELECT' From Dual
Union All Select '婚姻状况','SELECT' From Dual
Union All Select '疾病报告反馈','SELECT' From Dual
Union All Select '疾病报告前提','SELECT' From Dual
Union All Select '疾病编码分类','SELECT' From Dual
Union All Select '疾病编码科室','SELECT' From Dual
Union All Select '疾病编码类别','SELECT' From Dual
Union All Select '疾病编码目录','SELECT' From Dual
Union All Select '疾病申报反馈','SELECT' From Dual
Union All Select '疾病申报记录','SELECT' From Dual
Union All Select '疾病阳性记录','SELECT' From Dual
Union All Select '疾病诊断别名','SELECT' From Dual
Union All Select '疾病诊断参考','SELECT' From Dual
Union All Select '疾病诊断对照','SELECT' From Dual
Union All Select '疾病诊断分类','SELECT' From Dual
Union All Select '疾病诊断科室','SELECT' From Dual
Union All Select '疾病诊断目录','SELECT' From Dual
Union All Select '疾病诊断属类','SELECT' From Dual
Union All Select '检验报告项目','SELECT' From Dual
Union All Select '检验标本记录','SELECT' From Dual
Union All Select '检验普通结果','SELECT' From Dual
Union All Select '检验图像结果','SELECT' From Dual
Union All Select '检验细菌','SELECT' From Dual
Union All Select '检验项目','SELECT' From Dual
Union All Select '检验项目选项','SELECT' From Dual
Union All Select '检验药敏结果','SELECT' From Dual
Union All Select '检验仪器项目','SELECT' From Dual
Union All Select '检验用抗生素','SELECT' From Dual
Union All Select '结算方式','SELECT' From Dual
Union All Select '结算方式应用','SELECT' From Dual
Union All Select '就诊登记记录','SELECT' From Dual
Union All Select '临床部门','SELECT' From Dual
Union All Select '门诊病案记录','SELECT' From Dual
Union All Select '门诊费用记录','SELECT' From Dual
Union All Select '门诊诊室','SELECT' From Dual
Union All Select '门诊诊室','UPDATE' From Dual
Union All Select '门诊诊室适用科室','SELECT' From Dual
Union All Select '民族','SELECT' From Dual
Union All Select '排队叫号队列','SELECT' From Dual
Union All Select '票据打印内容','SELECT' From Dual
Union All Select '票据领用记录','SELECT' From Dual
Union All Select '票据使用明细','SELECT' From Dual
Union All Select '区域','SELECT' From Dual
Union All Select '人员抗菌药物权限','SELECT' From Dual
Union All Select '社区参数','SELECT' From Dual
Union All Select '社区目录','SELECT' From Dual
Union All Select '身份证未录原因','SELECT' From Dual
Union All Select '时间段','SELECT' From Dual
Union All Select '收费从属项目','SELECT' From Dual
Union All Select '收费价目','SELECT' From Dual
Union All Select '收费特定项目','SELECT' From Dual
Union All Select '收费细目','SELECT' From Dual
Union All Select '收入项目','SELECT' From Dual
Union All Select '消费卡类别目录','SELECT' From Dual
Union All Select '消费卡信息','SELECT' From Dual
Union All Select '性别','SELECT' From Dual
Union All Select '学历','SELECT' From Dual
Union All Select '药品剂型','SELECT' From Dual
Union All Select '药品价格记录','SELECT' From Dual
Union All Select '药品特性','SELECT' From Dual
Union All Select '一卡通目录','SELECT' From Dual
Union All Select '医保对照类别','SELECT' From Dual
Union All Select '医保对照明细','SELECT' From Dual
Union All Select '医疗付款方式','SELECT' From Dual
Union All Select '医疗卡挂失方式','SELECT' From Dual
Union All Select '医疗卡类别','SELECT' From Dual
Union All Select '医学警示','SELECT' From Dual
Union All Select '预约方式','SELECT' From Dual
Union All Select '帐户年度信息','SELECT' From Dual
Union All Select '诊疗分类目录','SELECT' From Dual
Union All Select '诊疗项目别名','SELECT' From Dual
Union All Select '诊疗项目目录','SELECT' From Dual
Union All Select '诊治所见项目','SELECT' From Dual
Union All Select '证件类型','SELECT' From Dual
Union All Select '职业','SELECT' From Dual
Union All Select '执业类别','SELECT' From Dual
Union All Select '临床出诊记录','SELECT' From Dual
Union All Select '临床出诊号源','SELECT' From Dual
Union All Select '病人身份关联','SELECT' From Dual
Union All Select '病人地址信息','SELECT' From Dual
Union All Select '疾病编码章节','SELECT' From Dual
Union All Select '血型','SELECT' From Dual
Union All Select '社会关系','SELECT' From Dual
Union All Select '急诊就诊记录','SELECT' From Dual
Union All Select '急诊病情级别','SELECT' From Dual
) A;

--1243:急诊医生工作站(病人接诊)
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select 100,1243,'病人接诊',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0
Union All Select 'ZL_病人接诊','EXECUTE' From Dual
Union All Select 'ZL_病人接诊_CANCEL','EXECUTE' From Dual
Union All Select 'ZL_病人接诊完成','EXECUTE' From Dual
Union All Select 'ZL_病人接诊完成_CANCEL','EXECUTE' From Dual
Union All Select 'ZL_挂号病人病案_Insert','EXECUTE' From Dual
Union All Select 'Zl_QueuedateCheck','EXECUTE' From Dual
Union All Select 'Zl_病人挂号记录_回诊','EXECUTE' From Dual
Union All Select 'Zl_病人挂号记录_取消回诊','EXECUTE' From Dual
Union All Select 'Zl_病人挂号记录_社区验证','EXECUTE' From Dual
Union All Select 'Zl_病人挂号记录_暂停操作','EXECUTE' From Dual
Union All Select 'Zl_病人挂号记录_转诊','EXECUTE' From Dual
Union All Select 'Zl_病人社区信息_Insert','EXECUTE' From Dual
Union All Select 'Zl_病人预约挂号_接收','EXECUTE' From Dual
Union All Select 'Zl_就诊变动记录_Insert','EXECUTE' From Dual
Union All Select 'Zl_Fun_Getblacklistinfor','EXECUTE' From Dual
Union All Select 'Zl_急诊绿色通道_Edit','EXECUTE' From Dual
Union All Select 'Zl_急诊病情级别_Edit','EXECUTE' From Dual) A;

--1243:急诊医生工作站(病人转诊)
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select 100,1243,'病人转诊',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0
Union All Select 'Zl_病人挂号记录_转诊','EXECUTE' From Dual) A;

--1243:急诊医生工作站(急诊首页)
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select 100,1243,'急诊首页',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0
Union All Select 'ZL_病人过敏记录_DELETE','EXECUTE' From Dual
Union All Select 'ZL_病人过敏记录_Insert','EXECUTE' From Dual
Union All Select 'ZL_病人信息_首页整理','EXECUTE' From Dual
Union All Select 'ZL_病人照片_DELETE','EXECUTE' From Dual
Union All Select 'ZL_病人诊断记录_DELETE','EXECUTE' From Dual
Union All Select 'ZL_病人诊断记录_Insert','EXECUTE' From Dual
Union All Select 'ZL_疾病编码科室_DELETE','EXECUTE' From Dual
Union All Select 'ZL_疾病编码科室_Insert','EXECUTE' From Dual
Union All Select 'ZL_疾病诊断科室_DELETE','EXECUTE' From Dual
Union All Select 'ZL_疾病诊断科室_Insert','EXECUTE' From Dual
Union All Select 'Zl_病人过敏记录_Update','EXECUTE' From Dual
Union All Select 'Zl_病人诊断记录_Update','EXECUTE' From Dual
Union All Select '病人照片','INSERT' From Dual) A;

--1243:急诊医生工作站(诊疗一览)
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select 100,1243,'诊疗一览',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0
Union All Select '病人医嘱报告','SELECT' From Dual
Union All Select '药品规格','SELECT' From Dual) A;



-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------
--145003:蒋廷中,2019-10-16,新增模块急诊预检分诊工作站
--报表：ZL1_REPORT_1244_1/急诊预检分诊单
Insert Into zlReports(ID,编号,名称,说明,密码,进纸,打印机,票据,系统,程序ID,功能,修改时间,发布时间,打印方式,禁止开始时间,禁止结束时间,分类ID) Values(zlReports_ID.NextVal,'ZL1_REPORT_1244_1','急诊预检分诊单','急诊预检分诊单','D~!w;yopk3,Ruio2U"PT',15,Null,0,100,1244,'基本',Sysdate,Sysdate,0,To_Date('2019-10-01 00:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('2019-10-01 00:00:00','YYYY-MM-DD HH24:MI:SS'),Null);
Insert Into zlRPTFmts(报表ID,序号,说明,图样,W,H,纸张,纸向,动态纸张,是否停用,停用原因) Values(zlReports_ID.CurrVal,1,'预检分诊单1',0,11904,16832,9,1,0,Null,Null);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明,数据连接编号) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'预检分诊信息','分诊流水号,131|患者姓名,202|性别,202|年龄,202|联系电话,202|分诊科室,202|地址,202|到院时间,135|病人来源,202|主诉,202|病情级别,202|分诊次数,131|病人ID,131|陪同人员,202|病情情况,202',User||'.病人信息,'||User||'.急诊就诊记录,'||User||'.急诊分诊记录,'||User||'.急诊陪同人员',0,Null,Null);
Insert Into zlRPTSQLs(源ID,行号,内容)  Select zlRPTDatas_ID.CurrVal,a.* From (
Select 1,'Select c.Id 分诊流水号, a.姓名 患者姓名, a.性别, a.年龄, a.联系人电话 联系电话, c.分诊科室名称 分诊科室, a.联系人地址 地址, b.到院时间, b.病人来源, b.主诉,' From Dual
Union All Select 2,'       Decode(b.病情级别, Null, Null, b.病情级别 || ''级'') As 病情级别, c.分诊次数,  b.病人id, d.名称 陪同人员,' From Dual
Union All Select 3,'       ''第'' || c.分诊次数 || ''次分诊  自动（'' || c.自动病情级别 || ''级）'' ||' From Dual
Union All Select 4,'         Decode(c.人工病情级别, '''', '''', ''   人工（'' || c.人工病情级别 || ''级）'')||Decode(b.病情级别-b.分诊病情级别, 0, '''', ''   修订（'' || b.病情级别 || ''级）'') As 病情情况' From Dual
Union All Select 5,'From 病人信息 a, 急诊就诊记录 b, 急诊分诊记录 c, 急诊陪同人员 d' From Dual
Union All Select 6,'Where a.病人id = b.病人id And b.Id = c.就诊id And b.陪同人员 = d.编码 And c.Id = []' From Dual) a;
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,0,'分诊ID',1,Null,0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条1',1,Null,0,Null,0,Null,Null,1215,2520,9120,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条2',1,Null,0,Null,0,Null,Null,1245,3210,9120,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'线条3',1,Null,0,Null,0,Null,Null,1245,3870,9120,0,0,0,1,'宋体',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签3',2,Null,0,Null,0,'分诊序号:[预检分诊信息.分诊流水号]',Null,1305,2235,3855,225,0,0,1,'宋体',10.5,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签5',2,Null,0,Null,0,'患者姓名:[预检分诊信息.患者姓名]',Null,1305,2595,3630,225,0,0,1,'宋体',10.5,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签10',2,Null,0,Null,0,'分诊科室:[预检分诊信息.分诊科室]',Null,1305,2910,3630,225,0,0,1,'宋体',10.5,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签11',2,Null,0,Null,0,'到院时间:[预检分诊信息.到院时间]',Null,1305,3270,3360,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签14',2,Null,0,Null,0,'主诉:[预检分诊信息.主诉]',Null,1305,3585,2520,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签6',2,Null,0,Null,0,'性别:[预检分诊信息.性别]',Null,3615,2595,2730,225,0,0,1,'宋体',10.5,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签1',2,Null,0,Null,0,'X X 市X X 医院',Null,4470,1140,2430,330,0,0,1,'宋体',16,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签15',2,Null,0,Null,0,'[预检分诊信息.病情情况]',Null,4740,2235,2610,225,0,0,1,'宋体',10.5,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签7',2,Null,0,Null,0,'年龄:[预检分诊信息.年龄]',Null,4740,2595,2730,225,0,0,1,'宋体',10.5,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签9',2,Null,0,Null,0,'地址:[预检分诊信息.地址]',Null,4755,2910,2730,225,0,0,1,'宋体',10.5,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签12',2,Null,0,Null,0,'病人来源:[预检分诊信息.病人来源]',Null,4755,3285,3360,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签2',2,Null,0,Null,0,'预检分诊单',Null,4785,1725,1425,300,0,0,1,'宋体',14,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签8',2,Null,0,Null,0,'联系电话:[预检分诊信息.联系电话]',Null,6585,2595,3630,225,0,0,1,'宋体',10.5,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签13',2,Null,0,Null,0,'陪同人员:[预检分诊信息.陪同人员]',Null,7185,3300,3360,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'条码1',13,Null,3,Null,0,'[预检分诊信息.病人ID]','000',8895,465,1304,705,2,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'条码2',13,Null,10,Null,0,'[预检分诊信息.病人ID]','100',1680,405,690,690,2,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null);

--报表：ZL1_REPORT_1244_2/急诊腕带打印
Insert Into zlReports(ID,编号,名称,说明,密码,进纸,打印机,票据,系统,程序ID,功能,修改时间,发布时间,打印方式,禁止开始时间,禁止结束时间,分类ID) Values(zlReports_ID.NextVal,'ZL1_REPORT_1244_2','急诊腕带打印','急诊腕带打印','D~/z9za`f."Mewr,K1\T',15,Null,0,100,1244,'基本',Sysdate,Sysdate,0,To_Date('2019-09-01 00:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('2019-09-01 00:00:00','YYYY-MM-DD HH24:MI:SS'),Null);
Insert Into zlRPTFmts(报表ID,序号,说明,图样,W,H,纸张,纸向,动态纸张,是否停用,停用原因) Values(zlReports_ID.CurrVal,1,'腕带打印1',0,11904,16832,9,1,0,Null,Null);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明,数据连接编号) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'病人信息_数据','姓名,202|性别,202|分诊科室名称,202|年龄,202|分诊序号,131|病人ID,131',User||'.病人信息,'||User||'.急诊就诊记录,'||User||'.急诊分诊记录',0,Null,Null);
Insert Into zlRPTSQLs(源ID,行号,内容)  Select zlRPTDatas_ID.CurrVal,a.* From (
Select 1,'Select a.姓名, a.性别, c.分诊科室名称, a.年龄, c.Id 分诊序号, a.病人id' From Dual
Union All Select 2,'From 病人信息 a, 急诊就诊记录 b, 急诊分诊记录 c' From Dual
Union All Select 3,'Where a.病人id = b.病人id And b.Id = c.就诊id And c.Id = [ ]' From Dual) a;
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,0,'分诊ID',1,Null,0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'卡片1',14,Null,0,Null,0,Null,Null,2790,1575,5085,1260,0,0,0,'宋体',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签8',2,Null,0,Null,0,'分诊科室:[病人信息_数据.分诊科室名称]',Null,150,555,3885,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,zlRPTItems_ID.CurrVal-1,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签6',2,Null,0,Null,0,'姓名:[病人信息_数据.姓名]',Null,165,210,2625,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,zlRPTItems_ID.CurrVal-2,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签7',2,Null,0,Null,0,'性别:[病人信息_数据.性别]',Null,2595,210,2145,210,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,zlRPTItems_ID.CurrVal-3,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签9',2,Null,0,Null,0,'年龄:[病人信息_数据.年龄]',Null,2595,555,2130,210,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,zlRPTItems_ID.CurrVal-4,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'条码1',13,Null,10,Null,0,'[病人信息_数据.病人ID]','100',3995,155,690,690,2,0,1,'宋体',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,zlRPTItems_ID.CurrVal-5,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签11',2,Null,0,Null,0,'分诊序号:[病人信息_数据.分诊序号]',Null,150,885,3465,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,zlRPTItems_ID.CurrVal-6,Null,0,0,0,0,0,0,Null,Null);

--报表：ZL1_REPORT_1244_2/急诊腕带打印

--报表：ZL1_REPORT_1244_3/急诊分诊病情统计
Insert Into zlReports(ID,编号,名称,说明,密码,进纸,打印机,票据,系统,程序ID,功能,修改时间,发布时间,打印方式,禁止开始时间,禁止结束时间,分类ID) Values(zlReports_ID.NextVal,'ZL1_REPORT_1244_3','急诊分诊病情统计','急诊分诊病情统计','D~>e?rp{0 Nm|q"D2PT',15,Null,0,100,1244,'基本',Sysdate,Sysdate,0,To_Date('2019-10-01 00:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('2019-10-01 00:00:00','YYYY-MM-DD HH24:MI:SS'),Null);
Insert Into zlRPTFmts(报表ID,序号,说明,图样,W,H,纸张,纸向,动态纸张,是否停用,停用原因) Values(zlReports_ID.CurrVal,1,'分诊病情统计1',0,11904,16832,9,1,0,Null,Null);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明,数据连接编号) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'时间病情级别汇总','病情级别,130|1级,139|2级,139|3级,139|4级,139',User||'.急诊就诊记录',0,Null,Null);
Insert Into zlRPTSQLs(源ID,行号,内容)  Select zlRPTDatas_ID.CurrVal,a.* From (
Select 1,'Select ''就诊人数'' 病情级别,' From Dual
Union All Select 2,'       (Select Count(*) 就诊人数 From 急诊就诊记录 a Where a.病情级别 = 1 And a.登记时间 >= [ ] And a.登记时间 < [ 1 ]) "1级",' From Dual
Union All Select 3,'       (Select Count(*) From 急诊就诊记录 a Where a.病情级别 = 2 And a.登记时间 >= [ ] And a.登记时间 < [ 1 ]) "2级",' From Dual
Union All Select 4,'       (Select Count(*) From 急诊就诊记录 a Where a.病情级别 = 3 And a.登记时间 >= [ ] And a.登记时间 < [ 1 ]) "3级",' From Dual
Union All Select 5,'       (Select Count(*) From 急诊就诊记录 a Where a.病情级别 = 4 And a.登记时间 >= [ ] And a.登记时间 < [ 1 ]) "4级"' From Dual
Union All Select 6,'From Dual' From Dual) a;
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,0,'开始时间',2,CHR(38)||'当天开始时间',0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,1,'结束时间',2,CHR(38)||'当天结束时间',0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签5',2,Null,0,Null,0,'操作人:[操作员姓名]',Null,1080,1245,1995,210,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签1',2,Null,0,Null,0,'日期范围：[=开始时间]至[=结束时间]',Null,4005,1245,3570,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签2',2,Null,0,Null,0,'分诊病情统计',Null,4470,735,1890,330,0,0,1,'宋体',16,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'任意表1',4,Null,0,Null,0,'时间病情级别汇总',Null,1080,1530,9230,3660,255,0,0,'宋体',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-1,0,Null,Null,'[时间病情级别汇总.病情级别]','4^225^病情级别^0^0',0,0,1680,0,0,1,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-2,1,Null,Null,'[时间病情级别汇总.1级]','4^225^1级^0^0',0,0,1830,0,0,1,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-3,2,Null,Null,'[时间病情级别汇总.2级]','4^225^2级^0^0',0,0,1860,0,0,1,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-4,3,Null,Null,'[时间病情级别汇总.3级]','4^225^3级^0^0',0,0,1980,0,0,1,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-5,4,Null,Null,'[时间病情级别汇总.4级]','4^225^4级^0^0',0,0,1845,0,0,1,0,'宋体',0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null);

--报表：ZL1_REPORT_1244_3/急诊分诊病情统计

--报表：ZL1_REPORT_1244_4/急诊分诊患者流量统计
Insert Into zlReports(ID,编号,名称,说明,密码,进纸,打印机,票据,系统,程序ID,功能,修改时间,发布时间,打印方式,禁止开始时间,禁止结束时间,分类ID) Values(zlReports_ID.NextVal,'ZL1_REPORT_1244_4','急诊分诊患者流量统计','急诊分诊患者流量统计','D~>b5ygve ,Niwm2E$BD',15,Null,0,100,1244,'基本',Sysdate,Sysdate,0,To_Date('2019-10-01 00:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('2019-10-01 00:00:00','YYYY-MM-DD HH24:MI:SS'),Null);
Insert Into zlRPTFmts(报表ID,序号,说明,图样,W,H,纸张,纸向,动态纸张,是否停用,停用原因) Values(zlReports_ID.CurrVal,1,'分诊患者流量统计1',0,11904,16832,9,1,0,Null,Null);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型,说明,数据连接编号) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'部门数据汇总','名称,202|就诊人数,139',User||'.临床部门,'||User||'.部门表,'||User||'.急诊就诊记录',1,Null,Null);
Insert Into zlRPTSQLs(源ID,行号,内容)  Select zlRPTDatas_ID.CurrVal,a.* From (
Select 1,'Select a.名称, Nvl(b.人数, 0) 就诊人数' From Dual
Union All Select 2,'From (Select a.部门id, b.名称 From 临床部门 a, 部门表 b Where a.工作性质 = ''20'' And a.部门id = b.Id) a,' From Dual
Union All Select 3,'     (Select a.分诊科室id, Count(1) 人数' From Dual
Union All Select 4,'       From 急诊就诊记录 a' From Dual
Union All Select 5,'       Where a.登记时间 >=[] And a.登记时间 <[1]' From Dual
Union All Select 6,'       Group By a.分诊科室id) b' From Dual
Union All Select 7,'Where a.部门id = b.分诊科室id(+)' From Dual
Union All Select 8,'Order By 名称' From Dual) a;
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,0,'开始时间',2,CHR(38)||'当天开始时间',0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定) Values(zlRPTDatas_ID.CurrVal,Null,1,'结束时间',2,CHR(38)||'当天结束时间',0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签4',2,Null,0,Null,0,'操作人:[操作员姓名]',Null,1425,1395,1995,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签2',2,Null,0,Null,0,'日期范围：[=开始时间]至[=结束时间]',Null,4050,1395,3570,225,0,0,1,'宋体',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签3',2,Null,0,Null,0,'分诊流量统计',Null,4470,615,1890,330,0,0,1,'宋体',16,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'汇总表1',5,Null,0,Null,0,'部门数据汇总',Null,1410,1695,4520,5085,255,0,0,'宋体',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,0,Null,Null,0,0,0,0,0,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,7,zlRPTItems_ID.CurrVal-1,0,Null,Null,'名称',Null,0,0,1005,0,255,0,0,Null,0,0,0,0,0,0,0,Null,Null,'SUM',0,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统,父ID,源ID,上下间距,左右间距,纵向分栏,横向分栏,源行号,表格线加粗,自适应行高,水平反转) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-2,0,Null,Null,'就诊人数',Null,0,0,1005,0,255,1,0,Null,0,0,0,0,0,0,0,Null,Null,Null,0,0,0,Null,Null,Null,Null,Null,Null,Null,0,Null,Null);

--报表：ZL1_REPORT_1244_4/急诊分诊患者流量统计



-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--145003:张永康,2019-10-14,新增模块急诊医生站
Create Or Replace Procedure Zl_Retu_Clinic
(
  n_Patiid In Number,
  v_Times  In Varchar2,
  n_Flag   In Number
) As
  --------------------------------------------
  --参数:n_Patiid,病人id
  --     v_Times,挂号单号或住院主页id（体检时，挂号单是体检单号）
  --     n_Flag,门诊或住院标志:0-门诊,1-住院,2-体检（此时，只有n_Patiid参数无效）
  --------------------------------------------
  Err_Item Exception;
  v_Err_Msg    Varchar2(100);
  n_System     Number(5);
  n_Opersystem Number(5);
  n_只读       Number(2);
  n_Count      Number(5);

  v_Table    Varchar2(100);
  v_Subtable Varchar2(100);
  v_Field    Varchar2(100);
  v_Subfield Varchar2(100);
  v_Sql      Varchar2(4000);
  v_Sqlchild Varchar2(4000);
  v_Fields   Varchar2(4000);

  v_Dblink Varchar2(30);

  Type t_Tab_Col Is Table Of Varchar2(4000) Index By Varchar2(32);
  Arr_Tab_Col t_Tab_Col;

  ---------------------------------------------
  --功能：获取表的字段字符串
  Function Getfields(v_Table In Varchar2) Return Varchar2 As
    v_Colstr Varchar2(4000);
  Begin
    If Arr_Tab_Col.Exists(v_Table) Then
      v_Colstr := Arr_Tab_Col(v_Table);
    Else
      Select f_List2str(Cast(Collect(Column_Name) As t_Strlist)) As Colsstr
      Into v_Colstr
      From (Select Column_Name From User_Tab_Columns Where Table_Name = v_Table Order By Column_Id);
    
      Arr_Tab_Col(v_Table) := v_Colstr;
    End If;
  
    Return v_Colstr;
  End Getfields;

  --------------------------------------------
  --返回指定病人ID和主页的相关表的子过程
  --------------------------------------------
  Procedure Zl_Retu_Other
  (
    n_Pati_Id 病案主页.病人id%Type,
    n_Page_Id 病案主页.主页id%Type
  ) As
  
  Begin
  
    For R In (Select Column_Value From Table(f_Str2list('病人过敏记录,病人诊断记录,病人手麻记录'))) Loop
      v_Table  := r.Column_Value;
      v_Fields := Getfields(v_Table);
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' ||
                  Replace(v_Fields, '待转出', 'Null as 待转出') || ' From H' || v_Table || ' Where 病人id = :1 And 主页id = :2';
      Execute Immediate v_Sql
        Using n_Pati_Id, n_Page_Id;
    
      v_Sql := 'Delete From H' || v_Table || ' Where 病人id = :1 And 主页id = :2';
      Execute Immediate v_Sql
        Using n_Pati_Id, n_Page_Id;
    End Loop;
  End Zl_Retu_Other;

  --------------------------------------------
  --返回指定病人ID和主页的用药清单表的子过程
  --------------------------------------------
  Procedure Zl_Retu_Drug
  (
    n_Pati_Id 病案主页.病人id%Type,
    n_Page_Id 病案主页.主页id%Type
  ) As
  Begin
    v_Table  := '病人用药清单';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                ' From H' || v_Table || ' Where 病人id = :1 And 主页id = :2 ';
    Execute Immediate v_Sql
      Using n_Pati_Id, n_Page_Id;
  
    --病人用药配方，在病人用药清单转出之后执行
    For P In (Select ID From H病人用药清单 Where 病人id = n_Pati_Id And 主页id = n_Page_Id) Loop
    
      v_Table := '病人用药配方';
      v_Field := '配方id';
    
      v_Fields := Getfields(v_Table);
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' ||
                  Replace(v_Fields, '待转出', 'Null as 待转出') || ' From H' || v_Table || ' Where ' || v_Field || ' = :1';
      Execute Immediate v_Sql
        Using p.Id;
      v_Sql := 'Delete H' || v_Table || ' Where ' || v_Field || ' = :1';
      Execute Immediate v_Sql
        Using p.Id;
    End Loop;
  
    Delete H病人用药清单 Where 病人id = n_Pati_Id And 主页id = n_Page_Id;
  End Zl_Retu_Drug;

  --------------------------------------------
  --返回指定病人ID和主页的临床路径相关表的子过程
  --------------------------------------------
  Procedure Zl_Retu_Path
  (
    n_Pati_Id 病案主页.病人id%Type,
    n_Page_Id 病案主页.主页id%Type
  ) As
  Begin
    v_Table  := '病人临床路径';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                ' From H' || v_Table || ' Where 病人id = :1 And 主页id = :2 ';
    Execute Immediate v_Sql
      Using n_Pati_Id, n_Page_Id;
  
    --病人路径医嘱，在病人医嘱记录转出之后执行
    For P In (Select ID As 路径记录id From H病人临床路径 Where 病人id = n_Pati_Id And 主页id = n_Page_Id) Loop
      For r_Exe In (Select ID As 路径执行id From H病人路径执行 Where 路径记录id = p.路径记录id) Loop
        For R In (Select Column_Value From Table(f_Str2list('病人路径医嘱变异'))) Loop
          v_Table := r.Column_Value;
          v_Field := '路径执行ID';
        
          v_Fields := Getfields(v_Table);
          v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' ||
                      Replace(v_Fields, '待转出', 'Null as 待转出') || ' From H' || v_Table || ' Where ' || v_Field ||
                      ' = :1';
          Execute Immediate v_Sql
            Using r_Exe.路径执行id;
        
          v_Sql := 'Delete H' || v_Table || ' Where ' || v_Field || ' = :1';
          Execute Immediate v_Sql
            Using r_Exe.路径执行id;
        End Loop;
      End Loop;
      For R In (Select Column_Value
                From Table(f_Str2list('病人路径执行,病人合并路径,病人路径评估,病人路径变异,病人路径指标,病人合并路径评估,病人出径记录'))) Loop
        v_Table := r.Column_Value;
        If v_Table = '病人合并路径' Then
          v_Field := '首要路径记录id';
        Else
          v_Field := '路径记录id';
        End If;
      
        v_Fields := Getfields(v_Table);
        v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' ||
                    Replace(v_Fields, '待转出', 'Null as 待转出') || ' From H' || v_Table || ' Where ' || v_Field || ' = :1';
        Execute Immediate v_Sql
          Using p.路径记录id;
      
        v_Sql := 'Delete H' || v_Table || ' Where ' || v_Field || ' = :1';
        Execute Immediate v_Sql
          Using p.路径记录id;
      End Loop;
    End Loop;
  
    Delete H病人临床路径 Where 病人id = n_Pati_Id And 主页id = n_Page_Id;
  End Zl_Retu_Path;
  --------------------------------------------
  --返回指定挂号ID的门诊临床路径相关表的子过程
  --------------------------------------------
  Procedure Zl_Retu_Pathout(n_Pati_Visit_Id 病人挂号记录.Id%Type) As
  Begin
    --病人路径医嘱，在病人医嘱记录转出之后执行
    For P In (Select Distinct ID
              From H病人门诊路径
              Where ID In (Select 路径记录id From H病人门诊路径记录 Where 挂号id = n_Pati_Visit_Id)) Loop
      v_Table  := '病人门诊路径';
      v_Fields := Getfields(v_Table);
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' ||
                  Replace(v_Fields, '待转出', 'Null as 待转出') || ' From H' || v_Table || ' Where ID = :1';
      Execute Immediate v_Sql
        Using p.Id;
      For R In (Select Column_Value
                From Table(f_Str2list('病人门诊路径记录,病人门诊路径执行,病人门诊路径评估,病人门诊路径变异,病人门诊路径指标,病人门诊出径记录'))) Loop
        v_Table := r.Column_Value;
        v_Field := '路径记录id';
      
        v_Fields := Getfields(v_Table);
        v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' ||
                    Replace(v_Fields, '待转出', 'Null as 待转出') || ' From H' || v_Table || ' Where ' || v_Field || ' = :1';
        Execute Immediate v_Sql
          Using p.Id;
      
        v_Sql := 'Delete H' || v_Table || ' Where ' || v_Field || ' = :1';
        Execute Immediate v_Sql
          Using p.Id;
      End Loop;
      Delete H病人门诊路径 Where ID = p.Id;
    End Loop;
  
  End Zl_Retu_Pathout;
  --------------------------------------------
  --返回指定病人ID和主页的护理相关表的子过程
  --------------------------------------------
  Procedure Zl_Retu_Tend
  (
    n_Pati_Id 病案主页.病人id%Type,
    n_Page_Id 病案主页.主页id%Type
  ) As
  Begin
  
    v_Table  := '病人护理文件';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                ' From H' || v_Table || ' Where 病人id = :1 And 主页id = :2 ';
    Execute Immediate v_Sql
      Using n_Pati_Id, n_Page_Id;
  
    For P In (Select ID As 文件id From H病人护理文件 Where 病人id = n_Pati_Id And 主页id = n_Page_Id) Loop
      For R In (Select Column_Value
                From Table(f_Str2list('病人护理数据,病人护理打印,病人护理诊断,病人护理活动项目,病人护理要素内容,产程要素内容'))) Loop
        v_Table  := r.Column_Value;
        v_Fields := Getfields(v_Table);
        v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' ||
                    Replace(v_Fields, '待转出', 'Null as 待转出') || ' From H' || v_Table || ' Where 文件id = :1';
        Execute Immediate v_Sql
          Using p.文件id;
      
        If v_Table = '病人护理数据' Then
          v_Fields := Getfields('病人护理明细');
          v_Sql    := 'Insert Into 病人护理明细(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                      ' From H病人护理明细 Where 记录id In (Select ID From H病人护理数据 Where 文件id = :1)';
          Execute Immediate v_Sql
            Using p.文件id;
        
          v_Sql := 'Delete H病人护理明细 Where 记录id In (Select ID From H病人护理数据 Where 文件id = :1)';
          Execute Immediate v_Sql
            Using p.文件id;
        End If;
      
        v_Sql := 'Delete H' || v_Table || ' Where 文件id = :1';
        Execute Immediate v_Sql
          Using p.文件id;
      End Loop;
    End Loop;
  
    Delete H病人护理文件 Where 病人id = n_Pati_Id And 主页id = n_Page_Id;
  
    --老版护理系统数据
    ------------------------------------------------------------------------
    v_Table  := '病人护理记录';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                ' From H' || v_Table || ' Where 病人id = :1 And 主页id = :2 ';
    Execute Immediate v_Sql
      Using n_Pati_Id, n_Page_Id;
  
    For P In (Select ID From H病人护理记录 Where 病人id = n_Pati_Id And 主页id = n_Page_Id) Loop
      v_Table  := '病人护理内容';
      v_Fields := Getfields(v_Table);
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' ||
                  Replace(v_Fields, '待转出', 'Null as 待转出') || ' From H' || v_Table || ' Where 记录ID = :1';
      Execute Immediate v_Sql
        Using p.Id;
    
      v_Sql := 'Delete H' || v_Table || ' Where 记录ID = :1';
      Execute Immediate v_Sql
        Using p.Id;
    End Loop;
  
    Delete H病人护理记录 Where 病人id = n_Pati_Id And 主页id = n_Page_Id;
  End Zl_Retu_Tend;

  --------------------------------------------
  --返回指定ID的病人新版电子病历记录子过程
  --------------------------------------------
  Procedure Zl_Retu_Epr(n_Rec_Id H电子病历记录.Id%Type) As
    v_Field Varchar2(100);
  Begin
    v_Table  := '电子病历记录';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                ' From H' || v_Table || ' Where id = :1';
    Execute Immediate v_Sql
      Using n_Rec_Id;
  
    --病人诊断记录在Zl_Retu_Other中已转回（无病历ID外键）
    --影像报告驳回,病人医嘱报告,报告查阅记录,这几张表的数据在Zl_Retu_Order中转回医嘱后再处理
    For R In (Select Column_Value
              From Table(f_Str2list('电子病历附件,电子病历格式,电子病历内容,疾病申报记录,疾病报告反馈,疾病申报反馈'))) Loop
      v_Table := r.Column_Value;
      If v_Table = '电子病历附件' Then
        v_Field := '病历id';
      Elsif v_Table = '疾病申报反馈' Then
        v_Field := '申报id';
      Else
        v_Field := '文件id';
      End If;
      v_Fields := Getfields(v_Table);
    
      --含LOB字段的表(电子病历图形,电子病历格式,电子病历附件)，其H表是临时表，所以需直接指定dblink
      If v_Dblink Is Not Null And (v_Table = '电子病历附件' Or v_Table = '电子病历格式') Then
        v_Sql := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                 ' From ' || v_Table || '@' || v_Dblink || ' Where ' || v_Field || ' = :1';
      Else
        v_Sql := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                 ' From H' || v_Table || ' Where ' || v_Field || ' = :1';
      End If;
      Execute Immediate v_Sql
        Using n_Rec_Id;
    
      If v_Table = '电子病历内容' Then
        v_Fields := Getfields('电子病历图形');
      
        If v_Dblink Is Not Null Then
          v_Sql := 'Insert Into 电子病历图形(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                   ' From 电子病历图形@' || v_Dblink ||
                   ' a Where 对象id In (Select ID From H电子病历内容 Where 文件id = :1 And 对象类型 = 5)';
        Else
          v_Sql := 'Insert Into 电子病历图形(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                   ' From H电子病历图形 Where 对象id In (Select ID From H电子病历内容 Where 文件id = :1 And 对象类型 = 5)';
        End If;
        Execute Immediate v_Sql
          Using n_Rec_Id;
      
        If v_Dblink Is Not Null Then
          v_Sql := 'Delete 电子病历图形@' || v_Dblink ||
                   ' Where 对象id In (Select ID From H电子病历内容 Where 文件id = :1 And 对象类型 = 5)';
          Execute Immediate v_Sql
            Using n_Rec_Id;
        Else
          Delete H电子病历图形 Where 对象id In (Select ID From H电子病历内容 Where 文件id = n_Rec_Id And 对象类型 = 5);
        End If;
      End If;
    
      If v_Dblink Is Not Null And (v_Table = '电子病历附件' Or v_Table = '电子病历格式') Then
        v_Sql := 'Delete ' || v_Table || '@' || v_Dblink || ' Where ' || v_Field || ' = :1';
      Else
        v_Sql := 'Delete H' || v_Table || ' Where ' || v_Field || ' = :1';
      End If;
      Execute Immediate v_Sql
        Using n_Rec_Id;
    End Loop;
  
    Delete H电子病历记录 Where ID = n_Rec_Id;
  End Zl_Retu_Epr;
  --------------------------------------------
  --返回指定ID的病人医嘱记录子过程，必须在病历、临床路径转出之后执行(病人医嘱报告,影像报告驳回，病人路径医嘱,病人门诊路径医嘱)
  --在Zl_Retu_Other中已转回了"病人诊断记录",转回"病人诊断医嘱"时不用再转
  --------------------------------------------
  Procedure Zl_Retu_Order(n_Rec_Id H病人医嘱记录.Id%Type) As
  Begin
    v_Table  := '病人医嘱记录';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                ' From H' || v_Table || ' Where id = :1';
    Execute Immediate v_Sql
      Using n_Rec_Id;
  
    --以"医嘱ID,发送号"为外键的，都按医嘱ID直接转回，只需要排在"病人医嘱发送"之后即可
    --由于外键关系，"报告查阅记录"须在"病人医嘱报告"后面
    For P In (Select Column_Value
              From Table(f_Str2list('病人医嘱计价,病人医嘱状态,病人医嘱发送,病人医嘱附费,病人医嘱附件,病人医嘱执行,病人医嘱打印,输血申请记录,输血检验结果,输血申请项目,' ||
                                     '医嘱执行打印,医嘱执行时间,医嘱执行计价,执行打印记录,病人诊断医嘱,病人路径医嘱,病人门诊路径医嘱,病人医嘱报告,报告查阅记录,' ||
                                     '影像报告驳回,影像报告记录,影像报告操作记录,影像检查记录,影像申请单图像,影像收藏内容,影像危急值记录,影像预约记录,检验标本记录,' ||
                                     '检验试剂记录,检验拒收记录,RIS检查预约,疾病阳性记录,医嘱申请单文件,病人危急值记录,药嘱禁忌说明,医嘱执行组合'))) Loop
      v_Table := p.Column_Value;
      If Instr('病人路径医嘱,病人门诊路径医嘱', v_Table) > 0 Then
        v_Field := '病人医嘱ID';
      Else
        v_Field := '医嘱ID';
      End If;
    
      v_Fields := Getfields(v_Table);
    
      If v_Dblink Is Not Null And (v_Table = '影像报告记录' Or v_Table = '医嘱申请单文件') Then
        v_Sql := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                 ' From ' || v_Table || '@' || v_Dblink || ' Where ' || v_Field || ' = :1';
      Else
        v_Sql := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                 ' From H' || v_Table || ' Where ' || v_Field || ' = :1';
      End If;
    
      If v_Table = '病人医嘱状态' Or v_Table = '病人医嘱报告' Then
        v_Sqlchild := v_Sql;
      Else
        Execute Immediate v_Sql
          Using n_Rec_Id;
      End If;
    
      If v_Table = '病人医嘱状态' Then
        v_Fields := Getfields('医嘱签名记录');
        v_Sql    := 'Insert Into 医嘱签名记录(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                    ' From H医嘱签名记录 Where ID In (Select 签名id From H病人医嘱状态 Where 医嘱id = :1 And 签名id Is Not Null)';
        Execute Immediate v_Sql
          Using n_Rec_Id;
      
        Delete H医嘱签名记录
        Where ID In (Select 签名id From H病人医嘱状态 Where 医嘱id = n_Rec_Id And 签名id Is Not Null);
      
        Execute Immediate v_Sqlchild
          Using n_Rec_Id;
      
      Elsif v_Table = '病人医嘱发送' Then
        v_Fields := Getfields('诊疗单据打印');
        v_Sql    := 'Insert Into 诊疗单据打印(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                    ' From H诊疗单据打印 Where (NO, 记录性质) In (Select NO, 记录性质 From H病人医嘱发送 Where 医嘱id = :1)';
        Execute Immediate v_Sql
          Using n_Rec_Id;
        Delete H诊疗单据打印 Where (NO, 记录性质) In (Select NO, 记录性质 From H病人医嘱发送 Where 医嘱id = n_Rec_Id);
      
      Elsif v_Table = '影像检查记录' Then
        v_Fields := Getfields('影像检查序列');
        v_Sql    := 'Insert Into 影像检查序列(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                    ' From H影像检查序列 Where 检查uid In (Select 检查uid From H影像检查记录 Where 医嘱id = :1)';
        Execute Immediate v_Sql
          Using n_Rec_Id;
      
        v_Fields := Getfields('影像检查图象');
        v_Sql    := 'Insert Into 影像检查图象(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                    ' From H影像检查图象 Where 序列uid In (Select b.序列uid From H影像检查记录 A, H影像检查序列 B Where a.医嘱id = :1 And a.检查uid = b.检查uid)';
        Execute Immediate v_Sql
          Using n_Rec_Id;
      
        Delete H影像检查图象
        Where 序列uid In (Select b.序列uid
                        From H影像检查记录 A, H影像检查序列 B
                        Where a.医嘱id = n_Rec_Id And a.检查uid = b.检查uid);
        Delete H影像检查序列 Where 检查uid In (Select 检查uid From H影像检查记录 Where 医嘱id = n_Rec_Id);
      
      Elsif v_Table = '检验标本记录' Then
        For R In (Select Column_Value
                  From Table(f_Str2list('检验申请项目,检验分析记录,检验项目分布,检验质控记录,检验操作记录,检验签名记录,检验图像结果'))) Loop
          v_Subtable := r.Column_Value;
          If v_Subtable = '检验签名记录' Then
            v_Subfield := '检验标本ID';
          Else
            v_Subfield := '标本ID';
          End If;
          v_Fields := Getfields(v_Subtable);
          If v_Subtable = '检验项目分布' Then
            v_Sql := 'Insert Into ' || v_Subtable || '(' || v_Fields || ') Select ' ||
                     Replace(v_Fields, '待转出', 'Null as 待转出') || ' From H' || v_Subtable || ' Where ' || v_Subfield ||
                     ' In (Select ID From H检验标本记录 Where 医嘱id = :1) And 医嘱id=:2';
            Execute Immediate v_Sql
              Using n_Rec_Id, n_Rec_Id;
          
            v_Sql := 'Delete H' || v_Subtable || ' Where ' || v_Subfield ||
                     ' In (Select ID From H检验标本记录 Where 医嘱id = :1)  And 医嘱id=:2';
            Execute Immediate v_Sql
              Using n_Rec_Id, n_Rec_Id;
          Elsif v_Dblink Is Not Null And v_Subtable = '检验图像结果' Then
            v_Sql := 'Insert Into ' || v_Subtable || '(' || v_Fields || ') Select ' ||
                     Replace(v_Fields, '待转出', 'Null as 待转出') || ' From ' || v_Subtable || '@' || v_Dblink || ' Where ' ||
                     v_Subfield || ' In (Select ID From H检验标本记录 Where 医嘱id = :1)';
            Execute Immediate v_Sql
              Using n_Rec_Id;
          
            v_Sql := 'Delete ' || v_Subtable || '@' || v_Dblink || ' Where ' || v_Subfield ||
                     ' In (Select ID From H检验标本记录 Where 医嘱id = :1)';
            Execute Immediate v_Sql
              Using n_Rec_Id;
          Else
            v_Sql := 'Insert Into ' || v_Subtable || '(' || v_Fields || ') Select ' ||
                     Replace(v_Fields, '待转出', 'Null as 待转出') || ' From H' || v_Subtable || ' Where ' || v_Subfield ||
                     ' In (Select ID From H检验标本记录 Where 医嘱id = :1)';
            Execute Immediate v_Sql
              Using n_Rec_Id;
          
            v_Sql := 'Delete H' || v_Subtable || ' Where ' || v_Subfield ||
                     ' In (Select ID From H检验标本记录 Where 医嘱id = :1)';
            Execute Immediate v_Sql
              Using n_Rec_Id;
          End If;
        End Loop;
      
        v_Fields := Getfields('检验普通结果');
        v_Sql    := 'Insert Into 检验普通结果(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                    ' From H检验普通结果 Where 检验标本id In (Select ID From H检验标本记录 Where 医嘱id = :1)';
        Execute Immediate v_Sql
          Using n_Rec_Id;
      
        v_Fields := Getfields('检验药敏结果');
        v_Sql    := 'Insert Into 检验药敏结果(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                    ' From H检验药敏结果 Where 细菌结果id In (Select ID From H检验普通结果 Where 检验标本id In (Select ID From H检验标本记录 Where 医嘱id = :1))';
        Execute Immediate v_Sql
          Using n_Rec_Id;
      
        v_Fields := Getfields('检验质控报告');
        v_Sql    := 'Insert Into 检验质控报告(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                    ' From H检验质控报告 Where 结果ID In (Select ID From H检验普通结果 Where 检验标本id In (Select ID From H检验标本记录 Where 医嘱id = :1))';
        Execute Immediate v_Sql
          Using n_Rec_Id;
      
        Delete H检验药敏结果
        Where 细菌结果id In
              (Select ID From H检验普通结果 Where 检验标本id In (Select ID From H检验标本记录 Where 医嘱id = n_Rec_Id));
        Delete H检验质控报告
        Where 结果id In
              (Select ID From H检验普通结果 Where 检验标本id In (Select ID From H检验标本记录 Where 医嘱id = n_Rec_Id));
      
        Delete H检验普通结果 Where 检验标本id In (Select ID From H检验标本记录 Where 医嘱id = n_Rec_Id);
      Elsif v_Table = '病人医嘱报告' Then
        v_Fields := Getfields('医嘱报告内容');
        If v_Dblink Is Not Null Then
          v_Sql := 'Insert Into 医嘱报告内容(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                   ' From 医嘱报告内容@' || v_Dblink || ' Where ID In (Select 报告id From 病人医嘱报告@' || v_Dblink ||
                   ' Where 医嘱id = :1 And 报告id Is Not Null)';
        Else
          v_Sql := 'Insert Into 医嘱报告内容(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                   ' From H医嘱报告内容 Where ID In (Select 报告id From H病人医嘱报告 Where 医嘱id = :1 And 报告id Is Not Null)';
        End If;
        Execute Immediate v_Sql
          Using n_Rec_Id;
      
        If v_Dblink Is Not Null Then
          v_Sql := 'Delete 医嘱报告内容@' || v_Dblink || ' Where ID In (Select 报告id From 病人医嘱报告@' || v_Dblink ||
                   ' Where 医嘱id = :1 And 报告id Is Not Null);';
          Execute Immediate v_Sql
            Using n_Rec_Id;
        Else
          Delete H医嘱报告内容
          Where ID In (Select 报告id From H病人医嘱报告 Where 医嘱id = n_Rec_Id And 报告id Is Not Null);
        End If;
      
        Execute Immediate v_Sqlchild
          Using n_Rec_Id;
      Elsif v_Table = '病人危急值记录' Then
      
        v_Fields := Getfields('病人危急值病历');
        v_Sql    := 'Insert Into 病人危急值病历(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                    ' From H病人危急值病历 Where 危急值id In (Select ID From H病人危急值记录 Where 医嘱id = :1)';
        Execute Immediate v_Sql
          Using n_Rec_Id;
        Delete H病人危急值病历 Where 危急值id In (Select ID From H病人危急值记录 Where 医嘱id = n_Rec_Id);
      
        v_Fields := Getfields('病人危急值医嘱');
        v_Sql    := 'Insert Into 病人危急值医嘱(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                    ' From H病人危急值医嘱 Where 危急值id In (Select ID From H病人危急值记录 Where 医嘱id = :1)';
        Execute Immediate v_Sql
          Using n_Rec_Id;
      
        Delete H病人危急值医嘱 Where 危急值id In (Select ID From H病人危急值记录 Where 医嘱id = n_Rec_Id);
      Elsif v_Table = '药嘱禁忌说明' Then
        v_Fields := Getfields('药嘱禁忌说明');
        v_Sql    := 'Insert Into 药嘱禁忌说明(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                    ' From H药嘱禁忌说明 Where 医嘱A = :1 OR 医嘱B = :2';
        Execute Immediate v_Sql
          Using n_Rec_Id, n_Rec_Id;
      
        Delete H药嘱禁忌说明 Where 医嘱a = n_Rec_Id Or 医嘱b = n_Rec_Id;        
      End If;
    
      If v_Dblink Is Not Null And (v_Table = '影像报告记录' Or v_Table = '医嘱申请单文件') Then
        v_Sql := 'Delete ' || v_Table || '@' || v_Dblink || ' Where ' || v_Field || ' = :1';
      Else
        v_Sql := 'Delete H' || v_Table || ' Where ' || v_Field || ' = :1';
      End If;
    
      Execute Immediate v_Sql
        Using n_Rec_Id;
    End Loop;
  
    --手麻数据
    If n_Opersystem > 0 Then
      Execute Immediate 'Call zl24_Retu_Oper(:1)'
        Using n_Rec_Id;
    End If;
  
    Delete H病人医嘱记录 Where ID = n_Rec_Id;
  End Zl_Retu_Order;
  --------------------------------------------
  --返回指定挂号单的(病人挂号记录\病人医嘱记录\电子病历记录\病人转诊记录)
  --------------------------------------------
  Procedure Zl_Retu_Outclinic(v_Times H病人挂号记录.No%Type) As
  Begin
    v_Table  := '病人挂号记录';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                ' From H' || v_Table || ' Where NO =:1 ';
    Execute Immediate v_Sql
      Using v_Times;
  
    For r_Other In (Select ID, 病人id From H病人挂号记录 Where NO = v_Times) Loop
      Zl_Retu_Other(r_Other.病人id, r_Other.Id);
    End Loop;
  
    For r_Epr In (Select b.Id
                  From H病人挂号记录 A, H电子病历记录 B
                  Where a.No = v_Times And a.病人id = n_Patiid And b.病人id = a.病人id And b.主页id = a.Id) Loop
      Zl_Retu_Epr(r_Epr.Id);
    End Loop;
  
    For r_Order In (Select ID From H病人医嘱记录 Where 病人来源 <> 4 And 病人id = n_Patiid And 挂号单 = v_Times) Loop
      Zl_Retu_Order(r_Order.Id);
    End Loop;
  
    --转诊记录
    v_Table  := '病人转诊记录';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                ' From H' || v_Table || ' Where NO =:1';
    Execute Immediate v_Sql
      Using v_Times;
      
    Delete H病人转诊记录 Where NO = v_Times;
      
    --急诊数据  
    v_Table  := '急诊就诊记录';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                ' From H' || v_Table || ' Where 挂号ID in(Select id From H病人挂号记录 Where NO =:1)';
    Execute Immediate v_Sql
      Using v_Times;
      
      
    v_Table  := '急诊分诊记录';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                ' From H' || v_Table || ' Where 就诊id in(Select b.ID From H病人挂号记录 A,H急诊就诊记录 B Where a.id = b.挂号ID And a.NO =:1)';
    Execute Immediate v_Sql
      Using v_Times;
      
      
    v_Table  := '急诊病人评分';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                ' From H' || v_Table || ' Where 分诊id in(Select c.id From H病人挂号记录 A,H急诊就诊记录 B,H急诊分诊记录 C Where b.id=c.就诊id And a.id = b.挂号ID And a.NO =:1)';
    Execute Immediate v_Sql
      Using v_Times;
      
    v_Table  := '急诊病人评分指标';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '待转出', 'Null as 待转出') ||
                ' From H' || v_Table || ' Where 评分id in(Select d.id From H病人挂号记录 A,H急诊就诊记录 B,H急诊分诊记录 C,H急诊病人评分 D Where c.id=d.分诊id And b.id=c.就诊id And a.id = b.挂号ID And a.NO =:1)';
    Execute Immediate v_Sql
      Using v_Times;
      
    Delete H急诊病人评分指标
    Where 评分id In (Select d.Id
                   From H病人挂号记录 A, H急诊就诊记录 B, H急诊分诊记录 C, H急诊病人评分 D
                   Where c.Id = d.分诊id And b.Id = c.就诊id And a.Id = b.挂号id And a.No = v_Times);
                   
    Delete H急诊病人评分
    Where 分诊id In (Select c.Id
                   From H病人挂号记录 A, H急诊就诊记录 B, H急诊分诊记录 C
                   Where b.Id = c.就诊id And a.Id = b.挂号id And a.No = v_Times);
                   
    Delete H急诊分诊记录
    Where 就诊id In (Select b.Id From H病人挂号记录 A, H急诊就诊记录 B Where a.Id = b.挂号id And a.No = v_Times);
    
    Delete H急诊就诊记录 Where 挂号id In (Select ID From H病人挂号记录 Where NO = v_Times);
    
    Delete H病人挂号记录 Where NO = v_Times;
  End Zl_Retu_Outclinic;
  --------------------------------------------
  --以下为主程序体
  --------------------------------------------
Begin
  ----------------------------------------------------------------------------------------------------------
  --对基于视图的转储方案进行了只读判断.
  Select 编号 Into n_System From zlSystems Where Upper(所有者) = Zl_Owner And 编号 Like '1%';
  Begin
    Select Nvl(只读, 0) Into n_只读 From zlBakSpaces Where 系统 = n_System And 当前 = 1;
  Exception
    When Others Then
      v_Err_Msg := '[ZLSOFT]当前没有可用的历史数据空间,不能继续![ZLSOFT]';
      Raise Err_Item;
  End;
  If n_只读 = 1 Then
    v_Err_Msg := '[ZLSOFT]历史数据空间目前的状态为只读,不能继续![ZLSOFT]';
    Raise Err_Item;
  End If;

  Select Max(Db连接) Into v_Dblink From zlBakSpaces Where 系统 = 100 And 当前 = 1;

  --对基于视图的转储方案进行了只读判断.
  n_Opersystem := 0;
  Select Nvl(max(编号),0) Into n_Opersystem From zlSystems Where Upper(所有者) = Zl_Owner And 编号 Like '24%';
  If n_Opersystem > 0 Then
    Begin
      Select Nvl(只读, 0) Into n_只读 From zlBakSpaces Where 系统 = n_Opersystem And 当前 = 1;
    Exception
      When Others Then
        v_Err_Msg := '[ZLSOFT]当前没有可用的手麻子系统历史数据空间,不能继续![ZLSOFT]';
        Raise Err_Item;
    End;
    If n_只读 = 1 Then
      v_Err_Msg := '[ZLSOFT]手麻子系统历史数据空间目前的状态为只读,不能继续![ZLSOFT]';
      Raise Err_Item;
    End If;
  End If;

  --1.门诊病人，按挂号单抽回
  If n_Flag = 0 Then
    --抽回未结记帐费用
    Zl_Retu_Exes(n_Patiid, 8);
  
    --存在门诊临床路径数据，将所有就诊数据抽回
    Select Count(1) Into n_Count From H病人门诊路径记录 A, H病人挂号记录 B Where b.Id = a.挂号id And b.No = v_Times;
    If n_Count = 0 Then
      Zl_Retu_Outclinic(v_Times);
    Else
      For R In (Select Distinct b.No
                From H病人门诊路径记录 A, H病人挂号记录 B
                Where a.挂号id = b.Id And
                      a.路径记录id In (Select a.路径记录id
                                   From H病人门诊路径记录 A, H病人挂号记录 B
                                   Where b.Id = a.挂号id And b.No = v_Times)) Loop
        Zl_Retu_Outclinic(r.No);
      End Loop;
      For R In (Select ID From 病人挂号记录 Where NO = v_Times) Loop
        Zl_Retu_Pathout(r.Id);
      End Loop;
    End If;
  
    --2.住院病人，按病人ID和主页ID抽回
  Elsif n_Flag = 1 Then
    --抽回未结记帐费用
    Zl_Retu_Exes(n_Patiid || ',' || v_Times, 8);
  
    Zl_Retu_Other(n_Patiid, To_Number(v_Times));
    Zl_Retu_Path(n_Patiid, To_Number(v_Times));
    Zl_Retu_Drug(n_Patiid, To_Number(v_Times));
  
    --先转病历，再转医嘱（影像报告驳回，病人医嘱报告这类又有病历又有医嘱的子表，在医嘱转回后处理）
    For r_Epr In (Select ID From H电子病历记录 Where 病人id = n_Patiid And 主页id = To_Number(v_Times)) Loop
      Zl_Retu_Epr(r_Epr.Id);
    End Loop;
  
    Zl_Retu_Tend(n_Patiid, To_Number(v_Times));
  
    For r_Order In (Select ID From H病人医嘱记录 Where 病人id = n_Patiid And 主页id = To_Number(v_Times)) Loop
      Zl_Retu_Order(r_Order.Id);
    End Loop;
  
    Update 病案主页 Set 数据转出 = 0 Where 病人id = n_Patiid And 主页id = To_Number(v_Times);
  
    --3.体检病人
  Elsif n_Flag = 2 Then
    Zl_Retu_Other(n_Patiid, v_Times);
  
    For r_Cpr In (Select ID From H病人医嘱记录 Where 病人来源 = 4 And 挂号单 = v_Times) Loop
      Zl_Retu_Order(r_Cpr.Id);
    End Loop;
  
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM || ':' || v_Sql);
End Zl_Retu_Clinic;
/

--145003:张永康,2019-10-14,新增急诊医生站
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
    结帐id_In t_Numlist,
    d_End     In Date,
    n_批次    In Number
  ) As
  
    c_结帐id t_Numlist := t_Numlist();
    c_No     t_Strlist := t_Strlist();
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

  Update Zldatamovelog
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

  Update Zldatamovelog
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

  Update Zldatamovelog
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

  Update Zldatamovelog
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

  Update Zldatamovelog
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

  Update Zldatamovelog
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

  Update Zldatamovelog
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

  Update Zldatamovelog
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

  Update Zldatamovelog
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

  Update Zldatamovelog
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
  Update Zldatamovelog
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

--145003:张永康,2019-10-14,新增模块急诊医生站
Create Or Replace Procedure Zl_急诊绿色通道_Edit
(
  挂号id_In 病人挂号记录.Id%Type,
  标记_In   急诊就诊记录.是否绿色通道%Type
) As
  --功能：标记或取消急诊绿色通道
Begin
  Update 急诊就诊记录 Set 是否绿色通道 = 标记_In Where 挂号id = 挂号id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_急诊绿色通道_Edit;
/

--145003:张永康,2019-10-14,新增模块急诊医生站
Create Or Replace Procedure Zl_急诊病情级别_Edit
(
  挂号id_In       病人挂号记录.Id%Type,
  病情级别_In     急诊就诊记录.病情级别%Type,
  修订说明_In     急诊就诊记录.修订说明%Type,
  修订人员_In     急诊就诊记录.修订人员%Type
) As
  --功能：用于急诊医生对急诊病情级别修订
Begin
  Update 急诊就诊记录
  Set 病情级别 = 病情级别_In, 修订说明 = 修订说明_In, 修订人员 = 修订人员_In, 修订时间 = Sysdate
  Where 挂号id = 挂号id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_急诊病情级别_Edit;
/

--145003:蒋廷中,2019-10-14,新增模块急诊预检分诊工作站
CREATE OR REPLACE Function Zl_Get_出诊诊室
(
  号码_In   挂号安排.号码%Type,
  记录id_In    临床出诊记录.Id%Type := Null,
  安排ID_In    挂号安排.Id%Type := Null,
  计划ID_In    挂号安排计划.Id%Type := Null

) Return Varchar2 As
  n_分诊方式 挂号安排.分诊方式%Type;
  v_号码     挂号安排.号码%Type;
  n_安排id   挂号安排.Id%Type;
  n_计划id   挂号安排计划.Id%Type;
  v_诊室     病人挂号记录.诊室%Type;
  v_Rowid    Varchar2(500);
  n_Next     Integer;
  n_First    Integer;

  v_Err_Msg Varchar2(200);
  Err_Item Exception;
Begin
  --安排计划
  If Nvl(记录id_In, 0) = 0 Then
    IF Nvl(安排ID_In, 0) = 0 And Nvl(计划ID_In, 0) = 0 Then
       Select Nvl(Max(号码),'-'), Nvl(Max(Id),0), Nvl(Max(分诊方式), 0) Into v_号码, n_安排id, n_分诊方式 From 挂号安排 Where 号码 = 号码_In;
       n_计划id := 0;
    else
      IF Nvl(计划ID_In, 0) = 0 Then
        Select Nvl(Max(号码),'-'), Nvl(Max(分诊方式), 0) Into v_号码, n_分诊方式 From 挂号安排 Where ID = 安排ID_In;
      Else
        Select Nvl(Max(号码),'-'), Nvl(Max(分诊方式), 0) Into v_号码, n_分诊方式 From 挂号安排计划 Where ID = 计划ID_In;
      End if;
      n_安排id := Nvl(安排ID_In, 0);
      n_计划id := Nvl(计划ID_In, 0);
    End IF;

    If v_号码 = '-' Then
      v_Err_Msg := '挂号安排未找到!';
      Raise Err_Item;
    End If;

    --0-不分诊、1-指定诊室、2-动态分诊、3-平均分诊,对应门诊诊室设置
    v_诊室 := Null;
    If n_分诊方式 = 1 Then
      --1-指定诊室
      IF n_计划id = 0 Then
        Select Max(门诊诊室) Into v_诊室 From 挂号安排诊室 Where 号表id = n_安排id;
      Else
        Select Max(门诊诊室) Into v_诊室 From 挂号计划诊室 Where 计划id = n_计划id;
      End if;
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
                          Where Nvl(执行状态, 0) = 0 And 发生时间 Between Trunc(Sysdate) And Sysdate And 号别 = v_号码 And
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
  End If;

  --==============================================================================================
  --临床出诊安排
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
                        Where Nvl(执行状态, 0) = 0 And 发生时间 Between Trunc(Sysdate) And Sysdate And 出诊记录id = 记录id_In And
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
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Get_出诊诊室;
/

--145003:蒋廷中,2019-10-14,新增模块急诊预检分诊工作站
CREATE OR REPLACE Procedure Zl_病人挂号记录_换号
( 
  No_In         病人挂号记录.No%Type, 
  号别_In       病人挂号记录.号别%Type, 
  诊室_In       病人挂号记录.诊室%Type, 
  科室id_In     病人挂号记录.执行部门id%Type, 
  原医生_In     病人挂号记录.执行人%Type, 
  原医生id_In   病人挂号汇总.医生id%Type, 
  新医生_In     病人挂号记录.执行人%Type, 
  新医生id_In   病人挂号汇总.医生id%Type, 
  出诊记录id_In 临床出诊记录.Id%Type := Null,
  操作类别_In   就诊变动记录.类别%Type := 2
  --功能：完成病人换号功能，在挂号项目ID相同的情况下。 
  --参数：
  --      操作类别_In:1-批量换号;2-分诊换号;3-强制续诊换号;4-预检分诊换号
) As 
  Cursor c_Bill Is 
    Select a.Id, a.记录性质, a.No, a.实际票号, a.记录状态, b.号序, a.序号, a.从属父号, a.价格父号, a.记帐单id, a.病人id, a.医嘱序号, a.门诊标志, a.记帐费用, a.姓名, 
           a.性别, a.年龄, a.标识号, a.付款方式, a.病人科室id, a.费别, 收费类别, a.收费细目id, a.计算单位, a.付数, a.发药窗口, a.数次, a.加班标志, a.附加标志, a.婴儿费, 
           a.收入项目id, a.收据费目, a.标准单价, a.应收金额, a.实收金额, a.划价人, a.开单部门id, a.开单人, b.发生时间, a.登记时间, a.执行部门id, a.执行人, a.执行状态, 
           a.执行时间, a.结论, a.操作员编号, a.操作员姓名, a.结帐id, a.结帐金额, a.保险大类id, a.保险项目否, a.保险编码, a.费用类型, a.统筹金额, a.是否上传, a.摘要, 
           a.是否急诊 
    From 门诊费用记录 A, 病人挂号记录 B 
    Where a.记录性质 = 4 And a.记录状态 = 1 And a.No = No_In And a.No = b.No 
    Order By a.序号; 
 
  v_病人id           门诊费用记录.Id%Type; 
  v_队列名称         排队叫号队列.队列名称%Type; 
  v_现队列名称       排队叫号队列.队列名称%Type; 
  v_挂号生成队列     Varchar2(2); 
  n_分诊台签到排队   Number; 
  n_再次签到重新排队 Number; 
  v_预约挂号         Number(2); 
  n_业务id           病人挂号记录.Id%Type; 
  v_排队号码         排队叫号队列.排队号码%Type; 
  v_号别             病人挂号记录.号别%Type; 
  n_号序             病人挂号记录.号序%Type; 
  v_排队序号         排队叫号队列.排队序号%Type; 
  d_排队时间         排队叫号队列.排队时间%Type; 
  v_Temp             Varchar2(500); 
  v_操作员编号       就诊变动记录.操作员编号%Type; 
  v_操作员姓名       就诊变动记录.操作员姓名%Type; 
  n_医生id           人员表.Id%Type; 
  n_诊室id           门诊诊室.Id%Type; 
  n_原出诊记录id     临床出诊记录.Id%Type; 
  n_变动id           就诊变动记录.Id%Type; 
  v_Error            Varchar2(255); 
  n_Exists           Number(3); 
  n_原序号           临床出诊序号控制.序号%Type; 
  n_原预约顺序号     临床出诊序号控制.预约顺序号%Type; 
  n_原挂号状态       临床出诊序号控制.挂号状态%Type; 
  v_原操作员         临床出诊序号控制.操作员姓名%Type; 
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
        Where 日期 = Trunc(r_Bill.发生时间) And Nvl(科室id, 0) = Nvl(r_Bill.执行部门id, 0) And Nvl(项目id, 0) = Nvl(r_Bill.收费细目id, 0) And 
              (号码 = r_Bill.计算单位 Or 号码 Is Null); 
        If Sql%RowCount = 0 Then 
          Insert Into 病人挂号汇总 
            (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 已约数, 其中已接收) 
          Values 
            (Trunc(r_Bill.发生时间), r_Bill.执行部门id, r_Bill.收费细目id, 原医生_In, Decode(原医生id_In, 0, Null, 原医生id_In), r_Bill.计算单位, 
             -1, -1 * v_预约挂号, -1 * v_预约挂号); 
        End If; 
 
        ----然后再更新挂号汇总 
        Update 病人挂号汇总 
        Set 已挂数 = Nvl(已挂数, 0) + 1, 其中已接收 = Nvl(其中已接收, 0) + v_预约挂号, 已约数 = Nvl(已约数, 0) + v_预约挂号 
        Where 日期 = Trunc(r_Bill.发生时间) And Nvl(科室id, 0) = 科室id_In And Nvl(项目id, 0) = Nvl(r_Bill.收费细目id, 0) And 
              (号码 = 号别_In Or 号码 Is Null); 
        If Sql%RowCount = 0 Then 
          Insert Into 病人挂号汇总 
            (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 已约数, 其中已接收) 
          Values 
            (Trunc(r_Bill.发生时间), 科室id_In, r_Bill.收费细目id, 新医生_In, Decode(新医生id_In, 0, Null, 新医生id_In), 号别_In, 1, v_预约挂号, 
             v_预约挂号); 
        End If; 
 
        --更新序号状态 
        Select Count(1) 
        Into n_Exists 
        From 挂号序号状态 
        Where 号码 = 号别_In And Trunc(日期) = Trunc(r_Bill.发生时间) And 序号 = r_Bill.号序 And Nvl(状态, 0) <> 0; 
 
        If n_Exists = 0 Then 
          Update 挂号序号状态 
          Set 号码 = 号别_In 
          Where Trunc(日期) = Trunc(r_Bill.发生时间) And 号码 = r_Bill.计算单位 And 序号 = r_Bill.号序; 
        Else 
          Delete From 挂号序号状态 
          Where Trunc(日期) = Trunc(r_Bill.发生时间) And 号码 = r_Bill.计算单位 And 序号 = r_Bill.号序; 
          Update 病人挂号记录 Set 号序 = Null Where NO = r_Bill.No; 
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
        Zl_就诊变动记录_Insert(r_Bill.No, Nvl(操作类别_In, 2), '分诊换号', v_操作员姓名, v_操作员编号, 号别_In, 科室id_In, Null, n_医生id, 新医生_In, 诊室_In, n_号序, 
                         Null, n_变动id); 
        v_挂号生成队列     := Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113)); 
        n_分诊台签到排队   := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113, 1, Nvl(科室id_In, 0))); 
        n_再次签到重新排队 := Zl_To_Number(zl_GetSysParameter('再次签到需重新排队', 1113)); 
 
        Select ID, 号别, Nvl(号序, 0) 
        Into n_业务id, v_号别, n_号序 
        From 病人挂号记录 
        Where NO = r_Bill.No And Rownum = 1; 
 
        If v_挂号生成队列 <> 0 Then 
          If Nvl(n_分诊台签到排队, 0) = 1 Then 
            Select 队列名称 Into v_队列名称 From 排队叫号队列 Where 业务id = n_业务id; 
            If Nvl(v_队列名称, 0) <> 0 And Nvl(n_再次签到重新排队, 0) = 1 Then 
              --删除原来排队记录重新排队：队列名称_IN，业务ID_IN 
              Zl_排队叫号队列_Delete(v_队列名称, n_业务id); 
            Else 
              Update 排队叫号队列 Set 排队状态 = 2 Where 业务id = n_业务id And 业务类型 = 0; 
            End If; 
            Update 病人挂号记录 Set 记录标志 = 0 Where ID = n_业务id; 
          Else 
            v_现队列名称 := 科室id_In; 
            --Zlgetnextqueue(执行部门id_In Number,业务id_In     Number := Null) 
            v_排队号码 := Zlgetnextqueue(科室id_In, n_业务id, v_号别 || '|' || n_号序); 
            v_排队序号 := Zlgetsequencenum(0, n_业务id, 1); 
            d_排队时间 := Zl_Get_Requeuedate(3, n_业务id, 科室id_In, 新医生_In, 诊室_In); 
            --新队列名称_IN, 业务类型_In, 业务id_In , 科室id_In , 患者姓名_In , 诊室_In , 医生姓名_In 
            Zl_排队叫号队列_Update(v_现队列名称, 0, n_业务id, 科室id_In, r_Bill.姓名, 诊室_In, 新医生_In, v_排队号码, v_排队序号, d_排队时间); 
            --换号后更新队列信息，排队状态也更新为排队中 
            Update 排队叫号队列 Set 排队状态 = 0 Where 业务id = n_业务id And 业务类型 = 0; 
          End If; 
        End If; 
        --删除转诊信息 
        Update 病人挂号记录 
        Set 执行部门id = 科室id_In, 号别 = 号别_In, 诊室 = 诊室_In, 执行人 = 新医生_In, 执行状态 = 0, 执行时间 = Null, 转诊号别 = Null, 转诊科室id = Null, 
            转诊诊室 = Null, 转诊医生 = Null, 转诊状态 = Null 
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
        Where 日期 = Trunc(r_Bill.发生时间) And Nvl(医生id, 0) = Nvl(原医生id_In, 0) And Nvl(医生姓名, '-') = Nvl(原医生_In, '-') And 
              Nvl(科室id, 0) = Nvl(r_Bill.执行部门id, 0) And Nvl(项目id, 0) = Nvl(r_Bill.收费细目id, 0) And 
              (号码 = r_Bill.计算单位 Or 号码 Is Null); 
        If Sql%RowCount = 0 Then 
          Insert Into 病人挂号汇总 
            (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 已约数, 其中已接收) 
          Values 
            (Trunc(r_Bill.发生时间), r_Bill.执行部门id, r_Bill.收费细目id, 原医生_In, Decode(原医生id_In, 0, Null, 原医生id_In), r_Bill.计算单位, 
             -1, -1 * v_预约挂号, -1 * v_预约挂号); 
        End If; 
        Update 临床出诊记录 
        Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - v_预约挂号, 已约数 = Nvl(已约数, 0) - v_预约挂号 
        Where ID = n_原出诊记录id; 
 
        ----然后再更新挂号汇总 
        Update 病人挂号汇总 
        Set 已挂数 = Nvl(已挂数, 0) + 1, 其中已接收 = Nvl(其中已接收, 0) + v_预约挂号, 已约数 = Nvl(已约数, 0) + v_预约挂号 
        Where 日期 = Trunc(r_Bill.发生时间) And Nvl(科室id, 0) = 科室id_In And Nvl(项目id, 0) = Nvl(r_Bill.收费细目id, 0) And 
              (号码 = 号别_In Or 号码 Is Null); 
        If Sql%RowCount = 0 Then 
          Insert Into 病人挂号汇总 
            (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 已约数, 其中已接收) 
          Values 
            (Trunc(r_Bill.发生时间), 科室id_In, r_Bill.收费细目id, 新医生_In, Decode(新医生id_In, 0, Null, 新医生id_In), 号别_In, 1, v_预约挂号, 
             v_预约挂号); 
        End If; 
        Update 临床出诊记录 
        Set 已挂数 = Nvl(已挂数, 0) + 1, 其中已接收 = Nvl(其中已接收, 0) + v_预约挂号, 已约数 = Nvl(已约数, 0) + v_预约挂号 
        Where ID = 出诊记录id_In; 
 
        --更新序号控制 
        Select Max(序号), Max(预约顺序号), Max(挂号状态), Max(操作员姓名) 
        Into n_原序号, n_原预约顺序号, n_原挂号状态, v_原操作员 
        From 临床出诊序号控制 
        Where 记录id = n_原出诊记录id And (序号 = r_Bill.号序 Or 备注 = To_Char(r_Bill.号序)); 
 
        If n_原序号 Is Not Null Then 
          Select Count(1) 
          Into n_Exists 
          From 临床出诊序号控制 
          Where 记录id = 出诊记录id_In And 序号 = n_原序号 And Nvl(预约顺序号, 0) = Nvl(n_原预约顺序号, 0) And Nvl(挂号状态, 0) = 0; 
          If n_Exists = 1 Then 
            Update 临床出诊序号控制 
            Set 挂号状态 = n_原挂号状态, 操作员姓名 = v_原操作员 
            Where 记录id = 出诊记录id_In And 序号 = n_原序号 And Nvl(预约顺序号, 0) = Nvl(n_原预约顺序号, 0) And Nvl(挂号状态, 0) = 0; 
          Else 
            Update 病人挂号记录 Set 号序 = Null Where NO = r_Bill.No; 
          End If; 
          Update 临床出诊序号控制 
          Set 挂号状态 = 0, 操作员姓名 = Null 
          Where 记录id = n_原出诊记录id And 序号 = n_原序号 And Nvl(预约顺序号, 0) = Nvl(n_原预约顺序号, 0); 
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
        Zl_就诊变动记录_Insert(r_Bill.No, 2, '分诊换号', v_操作员姓名, v_操作员编号, 号别_In, 科室id_In, Null, n_医生id, 新医生_In, 诊室_In, n_号序, 
                         Null, n_变动id); 
        v_挂号生成队列     := Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113)); 
        n_分诊台签到排队   := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113, 1, Nvl(科室id_In, 0))); 
        n_再次签到重新排队 := Zl_To_Number(zl_GetSysParameter('再次签到需重新排队', 1113)); 
        Select ID, 号别, Nvl(号序, 0) 
        Into n_业务id, v_号别, n_号序 
        From 病人挂号记录 
        Where NO = r_Bill.No And Rownum = 1; 
        If v_挂号生成队列 <> 0 Then 
          If Nvl(n_分诊台签到排队, 0) = 1 Then 
            Select 队列名称 Into v_队列名称 From 排队叫号队列 Where 业务id = n_业务id; 
            If Nvl(v_队列名称, 0) <> 0 And Nvl(n_再次签到重新排队, 0) = 1 Then 
              --删除原来排队记录重新排队：队列名称_IN，业务ID_IN 
              Zl_排队叫号队列_Delete(v_队列名称, n_业务id); 
            Else 
              Update 排队叫号队列 Set 排队状态 = 2 Where 业务id = n_业务id And 业务类型 = 0; 
            End If; 
            Update 病人挂号记录 Set 记录标志 = 0 Where ID = n_业务id; 
          Else 
            v_现队列名称 := 科室id_In; 
            --Zlgetnextqueue(执行部门id_In Number,业务id_In     Number := Null) 
            v_排队号码 := Zlgetnextqueue(科室id_In, n_业务id, v_号别 || '|' || n_号序); 
            v_排队序号 := Zlgetsequencenum(0, n_业务id, 1); 
            d_排队时间 := Zl_Get_Requeuedate(3, n_业务id, 科室id_In, 新医生_In, 诊室_In); 
            --新队列名称_IN, 业务类型_In, 业务id_In , 科室id_In , 患者姓名_In , 诊室_In , 医生姓名_In 
            Zl_排队叫号队列_Update(v_现队列名称, 0, n_业务id, 科室id_In, r_Bill.姓名, 诊室_In, 新医生_In, v_排队号码, v_排队序号, d_排队时间); 
            --换号后更新队列信息，排队状态也更新为排队中 
            Update 排队叫号队列 Set 排队状态 = 0 Where 业务id = n_业务id And 业务类型 = 0; 
          End If; 
        End If; 
        Update 病人挂号记录 
        Set 执行部门id = 科室id_In, 号别 = 号别_In, 诊室 = 诊室_In, 执行人 = 新医生_In, 执行状态 = 0, 执行时间 = Null, 出诊记录id = 出诊记录id_In, 
            转诊号别 = Null, 转诊科室id = Null, 转诊诊室 = Null, 转诊医生 = Null, 转诊状态 = Null 
        Where NO = r_Bill.No; 
      End If; 
    End Loop; 
  End If; 
  b_Message.Zlhis_Regist_005(No_In, 2, n_变动id); 
Exception 
  When Err_Custom Then 
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]'); 
  When Others Then 
    zl_ErrorCenter(SQLCode, SQLErrM); 
End Zl_病人挂号记录_换号;
/

--145003:蒋廷中,2019-10-14,新增模块急诊预检分诊工作站
CREATE OR REPLACE Procedure Zl_门诊分诊取号_Insert
(
  病人id_In     病人信息.病人id%Type,
  记录id_In     临床出诊记录.Id%Type,
  安排id_In     挂号安排.Id%Type,
  单据号_In     病人挂号记录.No%Type,
  诊室_In       门诊诊室.名称%Type,
  医生姓名_In   挂号安排.医生姓名%Type,
  医生id_In     挂号安排.医生id%Type,
  开单部门id_In 门诊费用记录.开单部门id%Type,
  操作员编号_In 门诊费用记录.操作员编号%Type,
  操作员姓名_In 门诊费用记录.操作员姓名%Type,
  退号重用_In   Integer := 0,
  站点_In       Varchar2 := Null,
  记帐费用_In   Number := 0
) As
  ---------------------------------------------------------------------------
  --功能:主要应用于免挂号模式,即只在分诊取号，然后接诊生成划价单进行收费方式
  --参数:
       --记帐费用_In:急诊挂号系统使用：产生记账挂号单
  ----------------------------------------------------------------------------

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_费别       费别.名称%Type;
  v_姓名       病人信息.姓名%Type;
  v_年龄       病人信息.年龄%Type;
  v_性别       病人信息.性别%Type;
  n_门诊号     病人信息.门诊号%Type;
  d_出生日期   病人信息.出生日期%Type;
  v_身份证号   病人信息.身份证号%Type;

  v_号码       临床出诊号源.号码%Type;
  n_科室id     临床出诊号源.科室id%Type;
  n_项目id     临床出诊号源.项目id%Type;
  n_是否分时段 临床出诊记录.是否分时段%Type;
  n_已挂数     临床出诊记录.已挂数%Type;
  n_已约数     临床出诊记录.已约数%Type;
  n_结帐id     病人结帐记录.Id%Type;
  n_限号数     临床出诊记录.限号数%Type;
  n_限约数     临床出诊记录.限约数%Type;
  n_序号_Out   挂号序号状态.序号%Type;
  n_急诊       病人挂号记录.急诊%Type;

  n_应收金额 门诊费用记录.应收金额%Type;
  n_实收金额 门诊费用记录.实收金额%Type;
  n_序号     门诊费用记录.序号%Type;
  n_价格父号 门诊费用记录.价格父号%Type;
  n_从属父号 门诊费用记录.从属父号%Type;

  n_更新项目id     挂号安排.项目id%Type;
  n_挂号项目id     病人挂号记录.挂号项目id%Type;
  d_开始时间 Date;
  n_生成队列 Number(3);
  n_取号模式 Number(2);
  v_Temp     Varchar2(4000);

  v_现金     结算方式.名称%Type;
  v_机器名   挂号序号状态.机器名%Type;
  v_排队号码 排队叫号队列.排队号码%Type;
  v_排队序号 排队叫号队列.排队序号%Type;
  v_队列名称 排队叫号队列.队列名称%Type;

  n_费用id 门诊费用记录.Id%Type;
  n_预交id 病人预交记录.Id%Type;
  n_挂号id 病人挂号记录.Id%Type;

  n_组id 财务缴款分组.Id%Type;

  n_序号控制       挂号安排.序号控制%Type;
  n_Count          Number;
  d_排队时间       Date;
  v_星期           挂号安排限制.限制项目%Type;

  v_付款方式编码 医疗付款方式.编码%Type;
  v_付款方式名称 医疗付款方式.名称%Type;

  d_登记时间 Date;
  d_发生时间 Date;

  v_药品等级 收费价格等级.名称%Type;
  v_卫材等级 收费价格等级.名称%Type;
  v_普通等级 收费价格等级.名称%Type;
  n_金额小数 Number;
  n_单价小数 Number;
  v_传入     Varchar2(100);
Begin
  --获取当前机器名称

  Select Terminal Into v_机器名 From V$session Where Audsid = Userenv('sessionid');
  Begin
    n_Count := 1;
    Select a.姓名, a.年龄, a.性别, a.费别, a.门诊号, b.编码, b.名称, a.出生日期, a.身份证号
    Into v_姓名, v_年龄, v_性别, v_费别, n_门诊号, v_付款方式编码, v_付款方式名称, d_出生日期, v_身份证号
    From 病人信息 A, 医疗付款方式 B
    Where a.病人id = 病人id_In And a.医疗付款方式 = b.名称(+);
  Exception
    When Others Then
      n_Count := 0;
  End;

  If n_Count = 0 Then
    v_Err_Msg := '无法确定病人信息，不能取号操作！';
    Raise Err_Item;
  End If;

  If v_费别 Is Null Then
    Select Max(名称) Into v_费别 From 费别 Where 缺省标志 = 1 And Rownum < 2;
    If v_费别 Is Null Then
      v_Err_Msg := '无法确定病人费别，请检查缺省费别是否正确设置！';
      Raise Err_Item;
    End If;
  End If;
  v_Temp := Zl_Get_Pricegrade(站点_In, 病人id_In, 0, v_付款方式名称);

  For c_价格等级 In (Select Rownum As 序号, Column_Value As 价格等级 From Table(f_Str2list(v_Temp, '|'))) Loop
    If c_价格等级.序号 = 1 Then
      v_普通等级 := c_价格等级.价格等级;
    End If;
    If c_价格等级.序号 = 2 Then
      v_药品等级 := c_价格等级.价格等级;
    End If;
    If c_价格等级.序号 = 3 Then
      v_卫材等级 := c_价格等级.价格等级;
    End If;
  End Loop;

  n_组id     := Zl_Get组id(操作员姓名_In);
  d_登记时间 := Sysdate;
  d_发生时间 := Sysdate;

  --金额小数位数
  Select Zl_To_Number(Nvl(Zl_Getsysparameter(9), '2')), Zl_To_Number(Nvl(Zl_Getsysparameter(157), '5')), Zl_To_Number(Nvl(Zl_Getsysparameter(290), '0'))
  Into n_金额小数, n_单价小数, n_取号模式
  From Dual;

  v_现金 := '现金';
  Select Decode(To_Char(d_发生时间, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null)
  Into v_星期
  From Dual;

  --挂号获取安排
  Begin
    n_Count := 1;
    If 记录id_In <> 0 Then
      --避免并发锁
      Select a.是否序号控制, Nvl(a.限号数, 0), Nvl(a.限约数, 0), b.号码, a.科室id, a.项目id, 是否分时段, 已挂数, 已约数
      Into n_序号控制, n_限号数, n_限约数, v_号码, n_科室id, n_项目id, n_是否分时段, n_已挂数, n_已约数
      From 临床出诊记录 A, 临床出诊号源 B
      Where a.Id = 记录id_In And a.号源id = b.Id
      For Update;
    Else

      Select Max(1) Into n_是否分时段 From 挂号安排时段 Where 安排id = 安排id_In And 星期 = v_星期 And Rownum < 2;
      Select a.序号控制, Nvl(b.限号数, 0), Nvl(b.限约数, 0), a.号码, a.科室id, a.项目id
      Into n_序号控制, n_限号数, n_限约数, v_号码, n_科室id, n_项目id
      From 挂号安排 A, 挂号安排限制 B
      Where a.Id = b.安排id(+) And b.限制项目(+) = v_星期 And a.Id = 安排id_In;

      Begin
        Select 已挂数, 已约数
        Into n_已挂数, n_已约数
        From 病人挂号汇总
        Where 日期 = Trunc(d_发生时间) And 科室id = n_科室id And 项目id = n_项目id And 号码 = v_号码 And Rownum < 2
        For Update; --避免并发锁
      Exception
        When Others Then
          n_已挂数 := 0;
      End;
    End If;
  Exception
    When Others Then
      n_Count := 0;
  End;

  If n_Count = 0 Then
    v_Err_Msg := '不存相应的挂号安排数据,请检查';
    Raise Err_Item;
  End If;

  --检查数量是否充足
  If n_限号数 <= n_已挂数 And n_限号数 > 0 Then
    v_Err_Msg := '号别' || v_号码 || '在' || To_Char(Trunc(d_发生时间), 'yyyy-mm-dd ') || '已达到最大限制数！';
    Raise Err_Item;
  End If;

  --锁号操作
  If Nvl(n_序号控制, 0) = 1 Then
    If Nvl(记录id_In, 0) = 0 Then

      --1.传统模式
      Zl_挂号安排_传统_Lockno(2, v_号码, d_发生时间, Null, n_序号_Out, v_机器名, 操作员姓名_In, 安排id_In, Null, 0, 操作员姓名_In || '锁号', Null,
                        Null);
      If Nvl(n_序号_Out, 0) = 0 Then
        v_Err_Msg := '未找到有效的号序,可能该号源已经使用完,请检查';
        Raise Err_Item;
      End If;

      Update 挂号序号状态
      Set 状态 = 1
      Where 号码 = v_号码 And 序号 = n_序号_Out And 日期 Between Trunc(d_发生时间) And (Trunc(d_发生时间) + 1 - 1 / 24 / 60 / 60) And
            Nvl(状态, 0) = 5;

      If Sql%NotFound Then
        v_Err_Msg := '未找到序号' || n_序号_Out || '的相关记录,请检查';
        Raise Err_Item;
      End If;

      If n_是否分时段 = 1 Then

        Begin
          Select To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'HH24:mi:ss'), 'yyyy-mm-dd HH24:mi:ss'),
                 1
          Into d_开始时间, n_Count
          From 挂号安排时段
          Where 安排id = 安排id_In And 星期 = v_星期 And 序号 = n_序号_Out;
        Exception
          When Others Then
            n_Count := 0;
        End;
        If n_Count = 0 Then

          Select To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || ' ' || To_Char(Max(结束时间), 'HH24:mi:ss'),
                          'yyyy-mm-dd HH24:mi:ss')
          Into d_开始时间
          From 挂号安排时段
          Where 安排id = 安排id_In And 星期 = v_星期;
          If d_发生时间 > d_开始时间 Then
            d_开始时间 := d_发生时间;
          End If;
        End If;

        d_发生时间 := d_开始时间;

      End If;
    Else
      --2.临床出诊模式

      Zl_挂号安排_临床出诊_Lockno(2, 记录id_In, d_发生时间, Null, n_序号_Out, 0, 操作员姓名_In || '锁号', v_机器名, 操作员姓名_In, Null, Null, Null,
                          v_号码);
      If Nvl(n_序号_Out, 0) = 0 Then
        v_Err_Msg := '未找到有效的号序,可能该号源已经使用完,请检查';
        Raise Err_Item;
      End If;

      Update 临床出诊序号控制
      Set 挂号状态 = 1
      Where 记录id = 记录id_In And 序号 = n_序号_Out
      Returning 开始时间 Into d_开始时间;

      If Sql%NotFound Then
        v_Err_Msg := '未找到序号' || n_序号_Out || '的相关记录,请检查';
        Raise Err_Item;
      End If;

      If To_Char(d_开始时间, 'HH:MM:DD') <> '00:00:00' And Nvl(n_是否分时段, 0) = 1 Then
        d_发生时间 := d_开始时间;
      End If;

    End If;
  End If;

  --产生挂号费为零的记录
  If Nvl(记帐费用_In, 0)  = 0 then
    Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
  End If;
  v_传入 := '3|' || v_号码;
  If 记录id_In Is Null Then
    If Nvl(安排id_In, 0) <> 0 Then
      v_传入 := '0|' || 安排id_In;
    End If;
  Else
    v_传入 := '2|' || 记录id_In;
  End If;
  n_挂号项目id := n_项目id;
  n_更新项目id := Zl_Custom_Getregeventitem(病人id_In, v_姓名, v_身份证号, d_出生日期, v_性别, v_年龄, v_传入);
  If Nvl(n_更新项目id, 0) <> 0 Then
    n_项目id := n_更新项目id;
  End If;

  n_序号     := 1;
  n_价格父号 := Null;
  n_从属父号 := Null;
  n_费用id   := Null;

  For c_价格 In (Select 1 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, 1 As 数次, c.Id As 收入项目id,
                      c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, -1 As 执行科室类型, a.项目特性
               From 收费项目目录 A, 收费价目 B, 收入项目 C
               Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = n_项目id And Sysdate Between b.执行日期 And
                     Nvl(b.终止日期, Sysdate + 1) And
                     ((b.价格等级 Is Null And Nvl(v_普通等级, '-') = '-') Or b.价格等级 = Nvl(v_普通等级, '-') Or
                     (b.价格等级 Is Null And Not Exists
                      (Select 1
                        From 收费价目
                        Where b.收费细目id = 收费细目id And 价格等级 = Nvl(v_普通等级, '-') And Sysdate Between 执行日期 And
                              Nvl(终止日期, Sysdate + 1))))
               Union All
               Select 2 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, d.从项数次 As 数次,
                      c.Id As 收入项目id, c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, -1 As 执行科室类型, a.项目特性
               From 收费项目目录 A, 收费价目 B, 收入项目 C, 收费从属项目 D
               Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = d.从项id And d.主项id = n_项目id And
                     Sysdate Between b.执行日期 And Nvl(b.终止日期, Sysdate + 1) And
                     ((b.价格等级 Is Null And Nvl(v_普通等级, '-') = '-') Or b.价格等级 = Nvl(v_普通等级, '-') Or
                     (b.价格等级 Is Null And Not Exists
                      (Select 1
                        From 收费价目
                        Where b.收费细目id = 收费细目id And 价格等级 = Nvl(v_普通等级, '-') And Sysdate Between 执行日期 And
                              Nvl(终止日期, Sysdate + 1))))
               Order By 性质, 项目id, 收入项目id) Loop
    Select 病人费用记录_Id.Nextval Into n_费用id From Dual;
    n_应收金额 := Round(c_价格.数次 * c_价格.单价, n_金额小数);
    
    If Nvl(记帐费用_In, 0) = 1 And Nvl(c_价格.屏蔽费别, 0) <> 1 Then
      v_Temp     := Zl_Actualmoney(v_费别, c_价格.项目id, c_价格.收入项目id, n_应收金额);
      n_实收金额 := To_Number(Substr(v_Temp, Instr(v_Temp, ':') + 1));
      n_实收金额 := Round(n_实收金额, n_金额小数);
    Else
      n_实收金额 := 0;
    End If;

    If n_序号 = 1 Then
      n_急诊 := c_价格.项目特性;
    End If;
    
    If c_价格.性质 = 1 And n_价格父号 Is Null And n_序号 <> 1 Then
      n_价格父号 := 1;
      n_从属父号 := Null;
    end if;

    If c_价格.性质 = 2 And n_从属父号 Is Null And n_序号 <> 1 Then
      n_价格父号 := Null;
      n_从属父号 := 1;
    End If;
    Insert Into 门诊费用记录
      (ID, 记录性质, 记录状态, 序号, 价格父号, 从属父号, NO, 实际票号, 门诊标志, 加班标志, 附加标志, 发药窗口, 病人id, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 病人科室id, 收费类别,
       计算单位, 收费细目id, 收入项目id, 收据费目, 付数, 数次, 标准单价, 应收金额, 实收金额, 结帐金额, 结帐id, 记帐费用, 开单部门id, 开单人, 划价人, 执行部门id, 执行人, 操作员编号,
       操作员姓名, 发生时间, 登记时间, 保险大类id, 保险项目否, 保险编码, 统筹金额, 摘要, 结论, 缴款组id)
    Values
      (n_费用id, 4, 1, n_序号, n_价格父号, n_从属父号, 单据号_In, Null, 1, c_价格.项目特性, 0, 诊室_In, Decode(病人id_In, 0, Null, 病人id_In),
       Decode(n_门诊号, 0, Null, n_门诊号), v_付款方式编码, v_姓名, v_性别, v_年龄, v_费别, n_科室id, c_价格.类别, v_号码, c_价格.项目id, c_价格.收入项目id,
       c_价格.收据费目, 1, c_价格.数次, c_价格.单价, n_应收金额, n_实收金额, Decode(Nvl(记帐费用_In, 0), 1, Null, 0), Decode(Nvl(记帐费用_In, 0), 1, Null, n_结帐id), Nvl(记帐费用_In, 0), 开单部门id_In, 
       操作员姓名_In, 操作员姓名_In, n_科室id, 医生姓名_In, 操作员编号_In, 操作员姓名_In, d_发生时间, d_登记时间, Null, Null, Null, Null, Null, Null, n_组id);

    n_序号 := n_序号 + 1;
  End Loop;

  If n_费用id Is Null Then
    v_Err_Msg := '未找到相关的费用数据,请设置对应的挂号费用!';
    Raise Err_Item;
  End If;

  If Nvl(记帐费用_In, 0) = 0 Then
    Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
    Insert Into 病人预交记录
      (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
       结算性质)
    Values
      (n_预交id, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), v_现金, 0, d_登记时间, 操作员编号_In, 操作员姓名_In, n_结帐id, Null, n_组id,
       Null, Null, Null, Null, Null, Null, 4);
  End If;
  
  Update 病人信息 Set 就诊时间 = d_发生时间, 就诊状态 = 1, 就诊诊室 = 诊室_In Where 病人id = 病人id_In;

  Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;

  Insert Into 病人挂号记录
    (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 预约时间, 操作员编号,
     操作员姓名, 复诊, 号序, 社区, 预约, 预约方式, 摘要, 交易流水号, 交易说明, 合作单位, 接收时间, 接收人, 预约操作员, 预约操作员编号, 险类, 医疗付款方式, 收费单, 取号标志,
     记录标志, 出诊记录id, 挂号项目ID, 费别)
  Values
    (n_挂号id, 单据号_In, 1, 1, Decode(病人id_In, 0, Null, 病人id_In), n_门诊号, v_姓名, v_性别, v_年龄, v_号码, n_急诊, 诊室_In, Null, n_科室id,
     医生姓名_In, 0, Null, d_登记时间, d_发生时间, Null, 操作员编号_In, 操作员姓名_In, 0, n_序号_Out, Null, 0, Null, '', Null, Null, Null,
     d_登记时间, 操作员姓名_In, Null, Null, Null, v_付款方式名称, Null, n_取号模式, 0, Decode(Nvl(记录id_In, 0), 0, Null, 记录id_In), n_挂号项目id, v_费别);

  n_生成队列 := Zl_To_Number(Zl_Getsysparameter('排队叫号模式', 1113));

  --0-不产生队列;1-按医生或分诊台排队;2-先分诊,后医生站
  If Nvl(n_生成队列, 0) <> 0 Then

    --产生队列
    --.按”执行部门” 的方式生成队列
    v_队列名称 := n_科室id;
    v_排队号码 := Zlgetnextqueue(n_科室id, n_挂号id, v_号码 || '|' || n_序号_Out);
    v_排队序号 := Zlgetsequencenum(0, n_挂号id, 0);
    d_排队时间 := Zl_Get_Queuedate(n_挂号id, v_号码, n_序号_Out, d_发生时间);
    --  队列名称_In , 业务类型_In, 业务id_In,科室id_In,排队号码_In,排队标记_In,患者姓名_In,病人ID_IN, 诊室_In, 医生姓名_In, 排队时间_In
    Zl_排队叫号队列_Insert(v_队列名称, 0, n_挂号id, n_科室id, v_排队号码, Null, v_姓名, 病人id_In, 诊室_In, 医生姓名_In, d_排队时间, Null, Null, v_排队序号);
    
    Update 病人挂号记录 Set 记录标志 = 1 Where id = n_挂号id;
  End If;

  --生成汇总数据Zl_门诊分诊取号_Insert
  If Nvl(记录id_In, 0) <> 0 Then
    Update 临床出诊记录 Set 已挂数 = Nvl(已挂数, 0) + 1 Where ID = 记录id_In;
  End If;

  Zl_病人挂号汇总_Update(医生姓名_In, 医生id_In, n_挂号项目id, n_科室id, d_发生时间, 0, v_号码, 0, 记录id_In);

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
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_门诊分诊取号_Insert;
/

--145003:蒋廷中,2019-10-14,新增模块急诊预检分诊工作站
CREATE OR REPLACE Function Zl_EmergencyRegist
(
  病人id_In         病人信息.病人id%Type,
  科室id_In         病人挂号记录.执行部门ID%Type,
  站点_In           Varchar2 := Null,
  记帐费用_In       Number := 0
) Return number As
  ---------------------------------------------------------------------------
  --功能：His急诊挂号
  --2019-10-12:排除挂号金额为0的挂号项目
  --返回: 成功时返回挂号id
  ---------------------------------------------------------------------------
  v_Para       Varchar2(2000);
  n_挂号模式   Number(3);
  d_启用时间   Date;

  v_付款方式名称 病人信息.医疗付款方式%Type;
  
  v_号码       挂号安排.号码%Type;
  n_安排id     挂号安排.Id%Type;
  n_记录id     临床出诊记录.Id%Type;
  n_挂号id     病人挂号记录.id%Type;
  v_No         病人挂号记录.No%Type;
  v_医生姓名   挂号安排.医生姓名%Type;
  n_医生id     挂号安排.医生id%Type;
  n_开单部门id 门诊费用记录.开单部门ID%Type;
  v_操作员编号 人员表.编号%Type;
  v_操作员姓名 人员表.姓名%Type;
  v_药品等级   收费价格等级.名称%Type;
  v_卫材等级   收费价格等级.名称%Type;
  v_普通等级   收费价格等级.名称%Type;
  n_实名制     Number(3);
  n_认证       Number(3);
  n_Count      Number(3);
  xmlIn        xmlType;
  xmlOut       xmlType;

  Err_Item     Exception;
  v_Err_Msg    Varchar2(255);
Begin
  Select Count(1), Max(v_付款方式名称) Into n_Count, v_付款方式名称 From 病人信息 Where 病人id = 病人id_In And Rownum < 2;
  If Nvl(n_Count, 0) = 0 Then
    v_Err_Msg := '无法确定病人信息,请检查！';
    Raise Err_Item;
  End If;

  Select Count(1) Into n_Count From 部门表 Where id = 科室id_In And Rownum < 2;
  If Nvl(n_Count, 0) = 0 Then
    v_Err_Msg := '无法确定科室信息,请检查！';
    Raise Err_Item;
  End If;
  
  v_Para := Zl_Get_Pricegrade(站点_In, 病人id_In, 0, v_付款方式名称);

  For c_价格等级 In (Select Rownum As 序号, Column_Value As 价格等级 From Table(f_Str2list(v_Para, '|'))) Loop
    If c_价格等级.序号 = 1 Then
      v_普通等级 := c_价格等级.价格等级;
    End If;
    If c_价格等级.序号 = 2 Then
      v_药品等级 := c_价格等级.价格等级;
    End If;
    If c_价格等级.序号 = 3 Then
      v_卫材等级 := c_价格等级.价格等级;
    End If;
  End Loop;

  v_Para     := zl_GetSysParameter(256);
  n_挂号模式 := Substr(v_Para, 1, 1);
  Begin
    d_启用时间 := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
  Exception
    When Others Then
      d_启用时间 := Null;
  End;

  If n_挂号模式 = 1 And Sysdate < Nvl(d_启用时间, Sysdate - 30) Then
    n_挂号模式 := 0;
  End If;

  v_No       := Nextno(12);
  If n_挂号模式 = 0 Then
    Begin
      If Nvl(n_安排id, 0) = 0 Then
        Select a.Id, a.号码, a.医生姓名, a.医生id
        Into n_安排id, v_号码, v_医生姓名, n_医生ID
        from(Select a.Id, a.号类, a.号码, a.科室id, a.项目id, a.医生姓名, a.医生id, a.序号控制, a.分诊方式,
                 To_Date(To_Char(Sysdate, 'yyyy-mm-dd') || ' ' || To_Char(c.开始时间, 'hh24:mi:ss'),
                          'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                 To_Date(To_Char(Sysdate, 'yyyy-mm-dd') || ' ' || To_Char(c.终止时间, 'hh24:mi:ss'),
                         'yyyy-mm-dd hh24:mi:ss') + Case
                   When To_Char(c.开始时间, 'hh24:mi:ss') >= To_Char(c.终止时间, 'hh24:mi:ss') Then
                    1
                   Else
                    0
                 End As 终止时间
          From 挂号安排 a, 收费项目目录 b,
               (Select 时间段, Decode(Sign(开始时间1 - 当前时间), 1, 开始时间, 开始时间1) As 开始时间,
                        Decode(Sign(终止时间1 - 当前时间), 1, 终止时间1, 终止时间) As 终止时间
                 From (Select 时间段, 号类, 站点,
                               To_Date(Decode(Sign(开始时间 - 终止时间),
                                               1,
                                               To_Char(Sysdate - 1, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'HH24:MI:SS'),
                                               To_Char(Sysdate, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'HH24:MI:SS')),
                                        'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                               To_Date(Decode(Sign(开始时间 - 终止时间),
                                               1,
                                               To_Char(Sysdate + 1, 'yyyy-mm-dd') || ' ' || To_Char(终止时间, 'HH24:MI:SS'),
                                               To_Char(Sysdate, 'yyyy-mm-dd') || ' ' || To_Char(终止时间, 'HH24:MI:SS')),
                                        'yyyy-mm-dd hh24:mi:ss') As 终止时间, Sysdate As 当前时间,
                               To_Date(To_Char(Sysdate, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'HH24:MI:SS'),
                                        'yyyy-mm-dd hh24:mi:ss') As 开始时间1,
                               To_Date(To_Char(Sysdate, 'yyyy-mm-dd') || ' ' || To_Char(终止时间, 'HH24:MI:SS'),
                                        'yyyy-mm-dd hh24:mi:ss') As 终止时间1
                        From 时间段
                        Where 站点 Is Null And 号类 Is Null)
                 Where 当前时间 Between 开始时间 And 终止时间1 Or 当前时间 Between 开始时间1 And 终止时间) c
          Where a.科室id = 科室id_In And Sysdate Between c.开始时间 And c.终止时间 And a.停用日期 Is Null And a.项目id = b.Id And
                Nvl(b.项目特性, 0) = 1 And Decode(To_Char(Sysdate, 'D'), '1', a.周日, '2', a.周一, '3', a.周二, '4', a.周三, '5', a.周四, '6', a.周五, '7', a.周六,  Null) = c.时间段 And Not Exists
           (Select 1
                 From 挂号安排停用状态 t
                 Where t.安排id = a.Id And a.开始时间 Between t.开始停止时间 And t.结束停止时间 And a.终止时间 Between t.开始停止时间 And t.结束停止时间) And Exists
           (Select 1
             From 收费项目目录 e, 收费价目 f
             Where f.收费细目id = e.Id and e.Id = a.项目id And f.现价 <> 0 and  Sysdate between f.执行日期 and
                   Nvl(f.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) and
                   ((f.价格等级 Is Null and Nvl(v_普通等级, '-') = '-') Or f.价格等级 = Nvl(v_普通等级, '-') Or
                           (f.价格等级 Is Null and Not Exists
                            (Select 1
                              From 收费价目
                              Where f.收费细目id = 收费细目id and 价格等级 = Nvl(v_普通等级, '-') and Sysdate between 执行日期 and
                                    Nvl(终止日期, Sysdate + 1))))
             Union all
             Select 1
             From 收费项目目录 e, 收费价目 f, 收费从属项目 g
             Where f.收费细目id = e.Id and e.Id = g.从项id and g.主项id = a.项目id And f.现价 <> 0 and Sysdate between f.执行日期 and
                   Nvl(f.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) and
                   ((f.价格等级 Is Null and Nvl(v_普通等级, '-') = '-') Or f.价格等级 = Nvl(v_普通等级, '-') Or
                           (f.价格等级 Is Null and Not Exists
                            (Select 1
                              From 收费价目
                              Where f.收费细目id = 收费细目id and 价格等级 = Nvl(v_普通等级, '-') and Sysdate between 执行日期 and
                                    Nvl(终止日期, Sysdate + 1)))))
          Order By a.号码) a
        Where rownum < 2;
      Else
        Select a.号码, a.医生姓名, a.医生id Into v_号码, v_医生姓名, n_医生id From 挂号安排 a Where Id = n_安排id;
      End If;
    Exception
      When Others Then
        v_Err_Msg := '当前科室无急诊挂号安排。';
        Raise Err_Item;
    End;
  Else
    --出诊表排班模式
    Begin
      If Nvl(n_记录id, 0) = 0 Then
        Select a.Id, a.号码, a.医生姓名, a.医生id
        Into n_记录id, v_号码, v_医生姓名, n_医生id
        From (Select a.Id, b.号码,
                     Case When Sysdate Between Nvl(a.替诊开始时间, a.终止时间) And Nvl(a.替诊终止时间, a.开始时间) Then a.替诊医生姓名 Else a.医生姓名 End As 医生姓名,
                     Case When Sysdate Between Nvl(a.替诊开始时间, a.终止时间) And Nvl(a.替诊终止时间, a.开始时间) Then a.替诊医生id Else a.医生id End As 医生id
               From 临床出诊记录 a, 临床出诊号源 b, 收费项目目录 d
               Where a.号源id = b.Id And a.科室id = 科室id_In And (a.出诊日期 = Trunc(Sysdate) Or a.出诊日期 = Trunc(Sysdate) - 1) And
                     Sysdate Between Nvl(a.提前挂号时间, a.开始时间) And a.终止时间 And
                     (a.开始时间 < Nvl(a.停诊开始时间, a.终止时间) Or a.终止时间 > Nvl(a.停诊终止时间, a.开始时间) Or Exists
                      (Select 1
                       From 临床出诊序号控制 c, 临床出诊记录 d
                       Where d.Id = a.Id And c.记录id = d.Id And Nvl(c.是否停诊, 0) = 0 And d.是否序号控制 = 1 And d.是否分时段 = 1 And
                             c.开始时间 <> c.终止时间)) And Sysdate Not Between Nvl(a.停诊开始时间, a.终止时间) And
                     Nvl(a.停诊终止时间, a.开始时间) And Nvl(a.是否发布, 0) = 1 And Nvl(a.是否锁定, 0) = 0 And a.项目id = d.Id And
                     Nvl(d.项目特性, 0) = 1 And Exists
                     (Select 1
                       From 收费项目目录 e, 收费价目 f
                       Where f.收费细目id = e.Id and e.Id = a.项目id And f.现价 <> 0 and  Sysdate between f.执行日期 and
                             Nvl(f.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) and
                             ((f.价格等级 Is Null and Nvl(v_普通等级, '-') = '-') Or f.价格等级 = Nvl(v_普通等级, '-') Or
                                     (f.价格等级 Is Null and Not Exists
                                      (Select 1
                                        From 收费价目
                                        Where f.收费细目id = 收费细目id and 价格等级 = Nvl(v_普通等级, '-') and Sysdate between 执行日期 and
                                              Nvl(终止日期, Sysdate + 1))))
                       Union all
                       Select 1
                       From 收费项目目录 e, 收费价目 f, 收费从属项目 g
                       Where f.收费细目id = e.Id and e.Id = g.从项id and g.主项id = a.项目id And f.现价 <> 0 and Sysdate between f.执行日期 and
                             Nvl(f.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) and
                             ((f.价格等级 Is Null and Nvl(v_普通等级, '-') = '-') Or f.价格等级 = Nvl(v_普通等级, '-') Or
                                     (f.价格等级 Is Null and Not Exists
                                      (Select 1
                                        From 收费价目
                                        Where f.收费细目id = 收费细目id and 价格等级 = Nvl(v_普通等级, '-') and Sysdate between 执行日期 and
                                              Nvl(终止日期, Sysdate + 1)))))
               Order By b.号码) a
        Where Rownum < 2;
      Else
        Select a.医生姓名, a.医生id Into v_医生姓名, n_医生id From 临床出诊记录 a Where Id = n_记录id;
      End If;
    Exception
      When Others Then
        v_Err_Msg := '当前科室无急诊挂号安排。';
        Raise Err_Item;
    End;
  End If;

  n_开单部门id := To_Number(Zl_操作员信息(0));
  v_操作员编号 := Zl_操作员信息(1);
  v_操作员姓名 := Zl_操作员信息(2);

  xmlIn := xmlType('<IN><BRID>' || 病人id_In || '</BRID>' || '<HM>' || v_号码 || '</HM>' || '<CZJLID>' || n_记录id || '</CZJLID>' ||
                  '<GHSJ>'|| to_char(Sysdate, 'yyyy-mm-dd hh24:mi:ss') ||'</GHSJ>' || '<KSID>' || 科室id_In || '</KSID>' || '<YSXM>' || v_医生姓名 || '</YSXM>' ||
                  '<JCHBBR>' || 0 || '</JCHBBR></IN>');

  --1.挂号检查
  --1.1实名制检查
  n_实名制 := To_Number(Nvl(zl_GetSysParameter(319), '0'));
  If n_实名制 = 1 Then
    Select Count(1) Into n_认证 From 病人实名信息 Where 病人id = 病人id_In and RowNum < 2;
    IF n_认证 = 0 Then
      v_Err_Msg := '病人未实名认证，不能挂号。';
      Raise Err_Item;
    End if;
  End if;
  
  --1.2挂号限制检查
  Zl_Third_Registercheck(xmlIn, xmlOut);

  Select Extractvalue(Value(A), 'OUTPUT/ERROR/MSG')
  Into v_Err_Msg
  From Table(Xmlsequence(Extract(XmlOut, 'OUTPUT'))) A;
  If Not v_Err_Msg Is Null Then
    Raise Err_Item;
  End If;

  --2.先产生0费用挂号记录
  Zl_门诊分诊取号_Insert(病人id_In,n_记录id,n_安排id,v_no,Zl_Get_出诊诊室(v_号码,n_记录id,n_安排id),
                  v_医生姓名, n_医生id,n_开单部门id,v_操作员编号,v_操作员姓名,0,站点_In, 记帐费用_In);
  --3.再产生对应的门诊
  If Nvl(记帐费用_In, 0) = 0 Then
    Zl_门诊划价记录_Buliding(v_no);
  End If;

  Select id Into n_挂号id From 病人挂号记录 Where no = v_no And 记录状态 In (0,1,3);
  Return n_挂号id;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Err_Msg || '+[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_EmergencyRegist;
/

--145003:蒋廷中,2019-10-14,新增模块急诊预检分诊工作站
CREATE OR REPLACE Procedure Zl_EmergencyRegistDel
(
  挂号id_In     病人挂号记录.id%Type
) As
  ---------------------------------------------------------------------------
  --功能：His急诊退号
  --以下情况不能退号：1.非记账挂号但有挂号金额
  --                  2.结过帐或者挂号划价单已收费
  --                  3.挂号单进行了补结算或者产生了医嘱
  --                  4.挂号单补打了发票
  ---------------------------------------------------------------------------
  n_记录id     临床出诊记录.Id%Type;
  v_No         病人挂号记录.No%Type;
  n_记录状态   病人挂号记录.记录状态%Type;
  v_收费单     病人挂号记录.收费单%Type;
  n_销帐id     门诊费用记录.结帐id%Type;
  n_结帐id     门诊费用记录.结帐id%Type;
  n_结帐金额   门诊费用记录.结帐金额%Type;
  n_记帐费用   门诊费用记录.记帐费用%Type;
  v_票据号     门诊费用记录.实际票号%Type;
  v_操作员编号 人员表.编号%Type;
  v_操作员姓名 人员表.姓名%Type;
  n_Count      Number(3);
  n_退号重用   Number(2);

  Err_Item     Exception;
  v_Err_Msg    Varchar2(255);
Begin
  --1退号检查
  --1.1单据状态检查
  Select Max(a.no),Max(a.记录状态),Max(b.记帐费用),Max(a.收费单),Max(b.结帐id),max(结帐金额),Max(出诊记录ID),
         Max(b.实际票号)
  Into v_No,n_记录状态,n_记帐费用,v_收费单,n_结帐id, n_结帐金额, n_记录Id,v_票据号
  From 病人挂号记录 a, 门诊费用记录 b Where a.id = 挂号id_In And a.no = b.no And b.记录性质 = 4
  And b.序号 = 1 And Rownum < 2;

  If v_No Is Null Then
    v_Err_Msg := '未找到对应的挂号单，请检查单据是否存在。';
    Raise Err_Item;
  End If;

  If nvl(n_记录状态, 0) <> 1 Then
    v_Err_Msg := '挂号单'|| v_No ||'已退号。';
    Raise Err_Item;
  End If;

  If Nvl(n_结帐金额, 0) <> 0 Then
    v_Err_Msg := '挂号单'|| v_No ||'需要退费，请到窗口退号。';
    Raise Err_Item;
  End If;

  If Nvl(n_记帐费用, 0) = 1 And Nvl(n_结帐id, 0) <> 0 Then
    v_Err_Msg := '挂号单'|| v_No ||'已结帐，请到窗口退号。';
    Raise Err_Item;
  End if;

  If v_收费单 Is Not Null Then
    Select Max(结帐id)
    Into n_结帐id
    From 门诊费用记录 Where No = v_收费单 And 记录性质 = 1 And 记录状态 = 1 And 序号 = 1;

    If Nvl(n_结帐id, 0) <> 0 Then
      v_Err_Msg := '挂号划价单'|| v_收费单 ||'已收费，不能退号。';
      Raise Err_Item;
    End If;
  End If;

  If v_票据号 Is Not Null Then
    v_Err_Msg := '挂号单'|| v_No ||'已打印了发票，请到窗口退号。';
    Raise Err_Item;
  End If;

  --1.2补充结算检查，已存在补结算数据的，不能退号
  Begin
    Select 1
    Into n_Count
    From 费用补充记录 A,
         (Select Distinct 结帐id
           From 门诊费用记录
           Where NO = v_No And 记录性质 = 4
           Union
           Select Distinct 结帐id
           From 门诊费用记录
           Where NO = v_收费单 And 记录性质 = 1) B
    Where a.收费结帐id = b.结帐id And a.记录性质 = 1 And a.附加标志 = 1 And Nvl(a.费用状态, 0) <> 2 And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count = 1 Then
    v_Err_Msg := '挂号单'|| v_No ||'已经进行了二次结算, 不能退号!';
    Raise Err_Item;
  End If;

  --1.3医嘱检查，已经开过医嘱的，不能退号
  Select Count(1) Into n_Count From 病人医嘱记录 Where 挂号单 = v_No And Rownum < 2;
  If Nvl(n_Count, 0) <>  0 Then
    v_Err_Msg := '挂号单'|| v_No ||'已经开过医嘱, 不能退号!';
    Raise Err_Item;
  End If;

  v_操作员编号 := Zl_操作员信息(1);
  v_操作员姓名 := Zl_操作员信息(2);
  --2退号
  --2.1如果有划价单，先退划价
  If v_收费单 Is Not Null Then
    zl_门诊划价记录_delete(v_收费单);
  End If;
  --2.2作废业务数据
  If Nvl(n_记帐费用, 0) = 0 Then
    Select 病人结帐记录_Id.Nextval Into n_销帐id From Dual;
  End If;
  If n_记录id Is Null Then
    zl_病人挂号记录_delete(v_No,v_操作员编号,v_操作员姓名,Null,Null,Null,Null,Null,Null,n_销帐id);
  Else
    zl_病人挂号记录_出诊_delete(v_No,v_操作员编号,v_操作员姓名,Null,Null,Null,Null,Null,Null,n_销帐id);
  End If;
  --2.3更新汇总数据
  n_退号重用 := Zl_To_Number(zl_GetSysParameter('已退序号控制', 1111));
  zl_病人挂号收费_modify(v_No,n_销帐id,Null,0,1,Null,n_退号重用);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Err_Msg || '+[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_EmergencyRegistDel;
/

--145003:蒋廷中,2019-10-14,新增模块急诊预检分诊工作站
CREATE OR REPLACE Procedure Zl_EmergencyRegistRedo
(
  挂号id_In         病人挂号记录.id%Type,
  科室id_In         病人挂号记录.执行部门ID%Type,
  站点_In           Varchar2 := Null
) As
  ---------------------------------------------------------------------------
  --功能：His急诊挂号换号
  --      相同科室直接返回成功
  --      已退号或者产生了医嘱了不能换号
  ---------------------------------------------------------------------------
  v_Para       Varchar2(2000);

  n_病人id       病人信息.病人ID%Type;
  v_付款方式名称 病人信息.医疗付款方式%Type;

  v_号码       挂号安排.号码%Type;
  n_安排id     挂号安排.Id%Type;
  n_记录id     临床出诊记录.Id%Type;
  v_No         病人挂号记录.No%Type;
  n_记录状态   病人挂号记录.记录状态%Type;
  v_收费单     病人挂号记录.收费单%Type;
  v_原医生姓名 挂号安排.医生姓名%Type;
  n_原医生id   挂号安排.医生id%Type;
  n_原科室Id   挂号安排.科室ID%Type;
  v_医生姓名   挂号安排.医生姓名%Type;
  n_医生id     挂号安排.医生id%Type;
  v_药品等级   收费价格等级.名称%Type;
  v_卫材等级   收费价格等级.名称%Type;
  v_普通等级   收费价格等级.名称%Type;
  n_Count      Number(3);
  xmlIn        xmlType;
  xmlOut       xmlType;

  Err_Item     Exception;
  v_Err_Msg    Varchar2(255);
Begin
  Select Count(1) Into n_Count From 部门表 Where id = 科室id_In And Rownum < 2;
  If Nvl(n_Count, 0) = 0 Then
    v_Err_Msg := '无法确定科室信息,请检查！';
    Raise Err_Item;
  End If;

  Select Max(a.病人id), Max(a.no),Max(a.记录状态),Max(a.收费单),Max(出诊记录ID),
         Max(c.名称),Max(a.执行人),Max(d.id),Max(a.执行部门id)
  Into n_病人id,v_No,n_记录状态,v_收费单,n_记录Id,v_付款方式名称,v_原医生姓名,n_原医生id,n_原科室id
  From 病人挂号记录 a, 门诊费用记录 b, 医疗付款方式 c,人员表 d
  Where a.id = 挂号id_In And a.no = b.no And b.记录性质 = 4 And b.付款方式 = c.编码(+) And a.执行人 = d.id(+)
  And b.序号 = 1 And Rownum < 2;

  If v_No Is Null Then
    v_Err_Msg := '未找到对应的挂号单，请检查单据是否存在。';
    Raise Err_Item;
  End If;

  If n_原科室id = 科室id_In Then--相同科室不需要换号
    Return;
  End If;

  If nvl(n_记录状态, 0) <> 1 Then
    v_Err_Msg := '挂号单'|| v_No ||'已退号。';
    Raise Err_Item;
  End If;

  Select Count(1) Into n_Count From 病人医嘱记录 Where 挂号单 = v_No And Rownum < 2;
  If Nvl(n_Count, 0) <>  0 Then
    v_Err_Msg := '挂号单'|| v_No ||'已经开过医嘱, 不能换号!';
    Raise Err_Item;
  End If;
  
  v_Para := Zl_Get_Pricegrade(站点_In, n_病人id, 0, v_付款方式名称);

  For c_价格等级 In (Select Rownum As 序号, Column_Value As 价格等级 From Table(f_Str2list(v_Para, '|'))) Loop
    If c_价格等级.序号 = 1 Then
      v_普通等级 := c_价格等级.价格等级;
    End If;
    If c_价格等级.序号 = 2 Then
      v_药品等级 := c_价格等级.价格等级;
    End If;
    If c_价格等级.序号 = 3 Then
      v_卫材等级 := c_价格等级.价格等级;
    End If;
  End Loop;

  If n_记录Id Is Null Then
    Begin
      If Nvl(n_安排id, 0) = 0 Then
        Select a.Id, a.号码, a.医生姓名, a.医生id
        Into n_安排id, v_号码, v_医生姓名, n_医生ID
        from(Select a.Id, a.号类, a.号码, a.科室id, a.项目id, a.医生姓名, a.医生id, a.序号控制, a.分诊方式,
                 To_Date(To_Char(Sysdate, 'yyyy-mm-dd') || ' ' || To_Char(c.开始时间, 'hh24:mi:ss'),
                          'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                 To_Date(To_Char(Sysdate, 'yyyy-mm-dd') || ' ' || To_Char(c.终止时间, 'hh24:mi:ss'),
                         'yyyy-mm-dd hh24:mi:ss') + Case
                   When To_Char(c.开始时间, 'hh24:mi:ss') >= To_Char(c.终止时间, 'hh24:mi:ss') Then
                    1
                   Else
                    0
                 End As 终止时间
          From 挂号安排 a, 收费项目目录 b,
               (Select 时间段, Decode(Sign(开始时间1 - 当前时间), 1, 开始时间, 开始时间1) As 开始时间,
                        Decode(Sign(终止时间1 - 当前时间), 1, 终止时间1, 终止时间) As 终止时间
                 From (Select 时间段, 号类, 站点,
                               To_Date(Decode(Sign(开始时间 - 终止时间),
                                               1,
                                               To_Char(Sysdate - 1, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'HH24:MI:SS'),
                                               To_Char(Sysdate, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'HH24:MI:SS')),
                                        'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                               To_Date(Decode(Sign(开始时间 - 终止时间),
                                               1,
                                               To_Char(Sysdate + 1, 'yyyy-mm-dd') || ' ' || To_Char(终止时间, 'HH24:MI:SS'),
                                               To_Char(Sysdate, 'yyyy-mm-dd') || ' ' || To_Char(终止时间, 'HH24:MI:SS')),
                                        'yyyy-mm-dd hh24:mi:ss') As 终止时间, Sysdate As 当前时间,
                               To_Date(To_Char(Sysdate, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'HH24:MI:SS'),
                                        'yyyy-mm-dd hh24:mi:ss') As 开始时间1,
                               To_Date(To_Char(Sysdate, 'yyyy-mm-dd') || ' ' || To_Char(终止时间, 'HH24:MI:SS'),
                                        'yyyy-mm-dd hh24:mi:ss') As 终止时间1
                        From 时间段
                        Where 站点 Is Null And 号类 Is Null)
                 Where 当前时间 Between 开始时间 And 终止时间1 Or 当前时间 Between 开始时间1 And 终止时间) c
          Where a.科室id = 科室id_In And Sysdate Between c.开始时间 And c.终止时间 And a.停用日期 Is Null And a.项目id = b.Id And
                Nvl(b.项目特性, 0) = 1 And Decode(To_Char(Sysdate, 'D'), '1', a.周日, '2', a.周一, '3', a.周二, '4', a.周三, '5', a.周四, '6', a.周五, '7', a.周六,  Null) = c.时间段 And Not Exists
           (Select 1
                 From 挂号安排停用状态 t
                 Where t.安排id = a.Id And a.开始时间 Between t.开始停止时间 And t.结束停止时间 And a.终止时间 Between t.开始停止时间 And t.结束停止时间) And Exists
           (Select 1
             From 收费项目目录 e, 收费价目 f
             Where f.收费细目id = e.Id and e.Id = a.项目id And f.现价 <> 0 and  Sysdate between f.执行日期 and
                   Nvl(f.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) and
                   ((f.价格等级 Is Null and Nvl(v_普通等级, '-') = '-') Or f.价格等级 = Nvl(v_普通等级, '-') Or
                           (f.价格等级 Is Null and Not Exists
                            (Select 1
                              From 收费价目
                              Where f.收费细目id = 收费细目id and 价格等级 = Nvl(v_普通等级, '-') and Sysdate between 执行日期 and
                                    Nvl(终止日期, Sysdate + 1))))
             Union all
             Select 1
             From 收费项目目录 e, 收费价目 f, 收费从属项目 g
             Where f.收费细目id = e.Id and e.Id = g.从项id and g.主项id = a.项目id And f.现价 <> 0 and Sysdate between f.执行日期 and
                   Nvl(f.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) and
                   ((f.价格等级 Is Null and Nvl(v_普通等级, '-') = '-') Or f.价格等级 = Nvl(v_普通等级, '-') Or
                           (f.价格等级 Is Null and Not Exists
                            (Select 1
                              From 收费价目
                              Where f.收费细目id = 收费细目id and 价格等级 = Nvl(v_普通等级, '-') and Sysdate between 执行日期 and
                                    Nvl(终止日期, Sysdate + 1)))))
          Order By a.号码) a
        Where rownum < 2;
      Else
        Select a.号码, a.医生姓名, a.医生id Into v_号码, v_医生姓名, n_医生id From 挂号安排 a Where Id = n_安排id;
      End If;
    Exception
      When Others Then
        v_Err_Msg := '当前科室无急诊挂号安排。';
        Raise Err_Item;
    End;
  Else
    --出诊表排班模式
    Begin
      If Nvl(n_记录id, 0) = 0 Then
        Select a.Id, a.号码, a.医生姓名, a.医生id
        Into n_记录id, v_号码, v_医生姓名, n_医生id
        From (Select a.Id, b.号码,
                     Case When Sysdate Between Nvl(a.替诊开始时间, a.终止时间) And Nvl(a.替诊终止时间, a.开始时间) Then a.替诊医生姓名 Else a.医生姓名 End As 医生姓名,
                     Case When Sysdate Between Nvl(a.替诊开始时间, a.终止时间) And Nvl(a.替诊终止时间, a.开始时间) Then a.替诊医生id Else a.医生id End As 医生id
               From 临床出诊记录 a, 临床出诊号源 b, 收费项目目录 d
               Where a.号源id = b.Id And a.科室id = 科室id_In And (a.出诊日期 = Trunc(Sysdate) Or a.出诊日期 = Trunc(Sysdate) - 1) And
                     Sysdate Between Nvl(a.提前挂号时间, a.开始时间) And a.终止时间 And
                     (a.开始时间 < Nvl(a.停诊开始时间, a.终止时间) Or a.终止时间 > Nvl(a.停诊终止时间, a.开始时间) Or Exists
                      (Select 1
                       From 临床出诊序号控制 c, 临床出诊记录 d
                       Where d.Id = a.Id And c.记录id = d.Id And Nvl(c.是否停诊, 0) = 0 And d.是否序号控制 = 1 And d.是否分时段 = 1 And
                             c.开始时间 <> c.终止时间)) And Sysdate Not Between Nvl(a.停诊开始时间, a.终止时间) And
                     Nvl(a.停诊终止时间, a.开始时间) And Nvl(a.是否发布, 0) = 1 And Nvl(a.是否锁定, 0) = 0 And a.项目id = d.Id And
                     Nvl(d.项目特性, 0) = 1 And Exists
                     (Select 1
                       From 收费项目目录 e, 收费价目 f
                       Where f.收费细目id = e.Id and e.Id = a.项目id And f.现价 <> 0 and  Sysdate between f.执行日期 and
                             Nvl(f.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) and
                             ((f.价格等级 Is Null and Nvl(v_普通等级, '-') = '-') Or f.价格等级 = Nvl(v_普通等级, '-') Or
                                     (f.价格等级 Is Null and Not Exists
                                      (Select 1
                                        From 收费价目
                                        Where f.收费细目id = 收费细目id and 价格等级 = Nvl(v_普通等级, '-') and Sysdate between 执行日期 and
                                              Nvl(终止日期, Sysdate + 1))))
                       Union all
                       Select 1
                       From 收费项目目录 e, 收费价目 f, 收费从属项目 g
                       Where f.收费细目id = e.Id and e.Id = g.从项id and g.主项id = a.项目id And f.现价 <> 0 and Sysdate between f.执行日期 and
                             Nvl(f.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) and
                             ((f.价格等级 Is Null and Nvl(v_普通等级, '-') = '-') Or f.价格等级 = Nvl(v_普通等级, '-') Or
                                     (f.价格等级 Is Null and Not Exists
                                      (Select 1
                                        From 收费价目
                                        Where f.收费细目id = 收费细目id and 价格等级 = Nvl(v_普通等级, '-') and Sysdate between 执行日期 and
                                              Nvl(终止日期, Sysdate + 1)))))
               Order By b.号码) a
        Where Rownum < 2;
      Else
        Select a.医生姓名, a.医生id Into v_医生姓名, n_医生id From 临床出诊记录 a Where Id = n_记录id;
      End If;
    Exception
      When Others Then
        v_Err_Msg := '当前科室无急诊挂号安排。';
        Raise Err_Item;
    End;
  End If;
  
  xmlIn := xmlType('<IN><BRID>' || n_病人id || '</BRID>' || '<HM>' || v_号码 || '</HM>' || '<CZJLID>' || n_记录id || '</CZJLID>' ||
                  '<GHSJ>'|| to_char(Sysdate, 'yyyy-mm-dd hh24:mi:ss') ||'</GHSJ>' || '<KSID>' || 科室id_In || '</KSID>' || '<YSXM>' || v_医生姓名 || '</YSXM>' ||
                  '<JCHBBR>' || 0 || '</JCHBBR></IN>');

  --1.挂号检查
  --1.1挂号限制检查
  Zl_Third_Registercheck(xmlIn, xmlOut);

  Select Extractvalue(Value(A), 'OUTPUT/ERROR/MSG')
  Into v_Err_Msg
  From Table(Xmlsequence(Extract(XmlOut, 'OUTPUT'))) A;
  If Not v_Err_Msg Is Null Then
    Raise Err_Item;
  End If;

  --换号
  Zl_病人挂号记录_换号(v_no,v_号码,Zl_Get_出诊诊室(v_号码,n_记录id,n_安排id),科室id_In,v_原医生姓名,n_原医生id,v_医生姓名, n_医生id,n_记录id,4);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Err_Msg || '+[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_EmergencyRegistRedo;
/

--145003:蒋廷中,2019-10-14,新增模块急诊预检分诊工作站
Create Or Replace Package b_Emergency_Rating Is
  --疼痛分级方法
  --入参数量1①描述文本
  --调用形式：b_emergency_rating.Is_Pain_num_rating(4,5)
  --返回值类型：varchar2   返回结果形式 病人等级:分数：描述  例如：2:9:重度疼痛
  Function Is_Pain_Num_Rating(Describe Varchar2) Return Varchar2;
  --昏迷评分分级方法
  --入参数量3②睁眼反应指标id：指标结果描述 ③ 语言反应指标id：指标结果描述 ④活动反应指标id：指标结果描述
  --调用形式： b_emergency_rating.Is_coma_rating('1:声音刺激','2:定向良好','3:痛刺激屈曲')
  --返回值类型：varchar2   返回结果形式 病人等级:总分数：描述  例如：3:11:中度意识障碍
  Function Is_Coma_Rating
  (
    Open_Reaction     Varchar2,
    Language_Reaction Varchar2,
    Activity_Reaction Varchar2
  ) Return Varchar2;
  --判断客观评估为儿童还是成人函数
  Function Is_Judgement_Function
  (
    Agenum  Number,
    Ageunit Varchar2
  ) Return Varchar2;
  --客观评价分级方法成人和儿童方法
  --入参数量3①年龄 ②年龄单位 ③ 指标id：指标结果描述（可多个）
  --调用形式： b_emergency_rating.Is_objective_rating(5,'岁','11:9,6:100,4:20')
  --返回值类型：varchar2   返回结果形式 病人等级 1
  Function Is_Objective_Rating
  (
    Agenum           Number,
    Ageunit          Varchar2,
    Indexid_Describe Varchar2
  ) Return Varchar2;
End b_Emergency_Rating;
/
Create Or Replace Package Body b_Emergency_Rating Is

  Function Is_Pain_Num_Rating(Describe Varchar2) --疼痛等级规则
   Return Varchar2 As
    State_Level  Varchar2(10); --病人级别
    Score_Result Varchar2(100); --评分结果描述
    Score        Number; --分数
  Begin
    Select 指标结果分值 Into Score From 急诊评分方法规则 Where 指标结果描述 = Describe;
  
    Select Min(病情级别), Min(评分结果描述)
    Into State_Level, Score_Result
    From 急诊评分方法分级
    Where 运算符 = 2 And Score > 分值上限 And 方法id = 4 Or 运算符 = 3 And 分值下限 < Score And 方法id = 4 Or
          运算符 = 6 And Score Between 分值下限 And 分值上限 And 方法id = 4 Or 运算符 = 1 And 分值上限 = Score And 方法id = 4 Or
          运算符 = 4 And 分值上限 >= Score And 方法id = 4 Or 运算符 = 5 And Score <= 分值下限 And 方法id = 4;
    Return State_Level || ':' || Score || ':' || Score_Result;
  Exception
    --异常处理语句段
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Is_Pain_Num_Rating;

  Function Is_Coma_Rating( --昏迷等级规则
                          Open_Reaction     Varchar2,
                          Language_Reaction Varchar2,
                          Activity_Reaction Varchar2) Return Varchar2 As
  
    Coma_Score_All Number; --昏迷总分数
    Coma_Level     Varchar2(10); --昏迷等级
    Score_Result   Varchar2(100); --评分结果描述
  
    Coma_Id1    Varchar2(10); --昏迷-睁眼指标ID
    Coma_Text1  Varchar2(100); --昏迷-睁眼描述
    Coma_Score1 Number; --昏迷-睁眼分数
  
    Coma_Id2    Varchar2(10); --昏迷-语言指标ID
    Coma_Text2  Varchar2(100); --昏迷-语言描述
    Coma_Score2 Number; --昏迷-语言分数
  
    Coma_Id3    Varchar2(10); --昏迷-活动指标ID
    Coma_Text3  Varchar2(100); --昏迷-活动描述
    Coma_Score3 Number; --昏迷-活动分数
  Begin
    Select C1, C2 Into Coma_Id1, Coma_Text1 From Table(f_Str2list2(Open_Reaction));
    Select C1, C2 Into Coma_Id2, Coma_Text2 From Table(f_Str2list2(Language_Reaction));
    Select C1, C2 Into Coma_Id3, Coma_Text3 From Table(f_Str2list2(Activity_Reaction));
  
    Select 指标结果分值
    Into Coma_Score1
    From 急诊评分方法规则
    Where 方法id = 3 And 指标结果描述 = Coma_Text1 And 指标id = Coma_Id1;
  
    Select 指标结果分值
    Into Coma_Score2
    From 急诊评分方法规则
    Where 方法id = 3 And 指标结果描述 = Coma_Text2 And 指标id = Coma_Id2;
  
    Select 指标结果分值
    Into Coma_Score3
    From 急诊评分方法规则
    Where 方法id = 3 And 指标结果描述 = Coma_Text3 And 指标id = Coma_Id3;
    Coma_Score_All := Coma_Score1 + Coma_Score2 + Coma_Score3;
  
    Select Min(病情级别), Min(评分结果描述)
    Into Coma_Level, Score_Result
    From 急诊评分方法分级
    Where 运算符 = 2 And Coma_Score_All > 分值上限 And 方法id = 3 Or 运算符 = 3 And 分值下限 < Coma_Score_All And 方法id = 3 Or
          运算符 = 6 And Coma_Score_All Between 分值下限 And 分值上限 And 方法id = 3 Or
          运算符 = 1 And 分值上限 = Coma_Score_All And 方法id = 3 Or 运算符 = 4 And 分值上限 >= Coma_Score_All And 方法id = 3 Or
          运算符 = 5 And Coma_Score_All <= 分值下限 And 方法id = 3;
    Return Coma_Level || ':' || Coma_Score_All || ':' || Score_Result;
  
  Exception
    --异常处理语句段
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Is_Coma_Rating;

  Function Is_Judgement_Function
  (
    Agenum  Number,
    Ageunit Varchar2
  ) Return Varchar2 As
    --判断成人或儿童规则
    Children_Age Varchar2(100);
  Begin
    If Ageunit Is Null Then
      Return '1'; --成人
    End If;
  
    If Ageunit = '岁' Then
      Select 参数值 Into Children_Age From zlParameters Where 参数名 = '儿童年龄界定上限';
    
      If Agenum <= To_Number(Children_Age) Then
        Return '2'; --儿童
      Else
        Return '1'; --成人
      End If;
    End If;
    Return '2'; --年龄单位不为岁返回儿童0-1岁
  Exception
    --异常处理语句段
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Is_Judgement_Function;

  Function Is_Objective_Rating --客观评价评分规则
  (
    Agenum           Number,
    Ageunit          Varchar2,
    Indexid_Describe Varchar2
  ) Return Varchar2 As
    Person        Varchar2(2); --儿童或者成人
    o_Indexid     t_Numlist; --指标ID
    o_Describe    t_Numlist; --传入指标参数
    Level_Max     Number; --病情最大值
    Illness_Level Number; --病情级别
    Age_Id        Number; --儿童年龄id
  Begin
    Select b_Emergency_Rating.Is_Judgement_Function(Agenum, Ageunit) Into Person From Dual;
    If Person = '1' Then
      --成人的规则
      Select Max(病情级别) Into Level_Max From 急诊评分方法规则;
      Select C1, C2 Bulk Collect Into o_Indexid, o_Describe From Table(f_Str2list2(Indexid_Describe));
      For I In 1 .. o_Indexid.Count Loop
        Select Min(病情级别)
        Into Illness_Level
        From 急诊评分方法规则
        Where 运算符 = 2 And o_Describe(I) > 指标值上限 And 指标id = o_Indexid(I) And 方法id = 1 Or
              运算符 = 3 And o_Describe(I) < 指标值下限 And 方法id = 1 And 指标id = o_Indexid(I) Or
              运算符 = 6 And o_Describe(I) >= 指标值下限 And o_Describe(I) <= 指标值上限 And 方法id = 1 And 指标id = o_Indexid(I) Or
              运算符 = 1 And 指标值上限 = o_Describe(I) And 方法id = 1 And 指标id = o_Indexid(I) Or
              运算符 = 4 And o_Describe(I) >= 指标值上限 And 方法id = 1 And 指标id = o_Indexid(I) Or
              运算符 = 5 And o_Describe(I) <= 指标值下限 And 方法id = 1 And 指标id = o_Indexid(I);
        If Illness_Level < Level_Max Then
          Level_Max := Illness_Level;
        End If;
      End Loop;
      Return Level_Max;
    End If;
  
    If Person = '2' Then
      --儿童规则
      Select Max(病情级别) Into Level_Max From 急诊评分方法规则;
      --程序逻辑根据传入年龄和单位和指标id找到相应的年龄id，根据年龄id和指标id和指标值找到级别
      --如果没有找到，抛弃年龄条件寻找没有年龄值的级别，如果级别还为空将找到的最小级别赋给它
      If Ageunit = '岁' Then
        Select C1, C2 Bulk Collect Into o_Indexid, o_Describe From Table(f_Str2list2(Indexid_Describe));
        For I In 1 .. o_Indexid.Count Loop
          Select Max(ID)
          Into Age_Id
          From 急诊评分指标年龄
          Where 运算符 = 2 And Agenum > 年龄上限 And 指标id = o_Indexid(I) And 年龄单位 = '岁' Or
                运算符 = 3 And Agenum < 年龄下限 And 指标id = o_Indexid(I) And 年龄单位 = '岁' And 年龄单位 = '岁' Or
                运算符 = 6 And Agenum Between 年龄下限 And 年龄上限 And 指标id = o_Indexid(I) And 年龄单位 = '岁' Or
                运算符 = 1 And 年龄上限 = Agenum And 指标id = o_Indexid(I) And 年龄单位 = '岁' Or
                运算符 = 4 And Agenum >= 年龄上限 And 指标id = o_Indexid(I) And 年龄单位 = '岁' Or
                运算符 = 5 And Agenum <= 年龄下限 And 指标id = o_Indexid(I) And 年龄单位 = '岁';
        
          Select Min(病情级别)
          Into Illness_Level
          From 急诊评分方法规则
          Where 运算符 = 2 And o_Describe(I) > 指标值上限 And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And 方法id = 2 Or
                运算符 = 3 And o_Describe(I) < 指标值下限 And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And 方法id = 2 Or
                运算符 = 6 And o_Describe(I) Between 指标值下限 And 指标值上限 And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And
                方法id = 2 Or 运算符 = 1 And 指标值上限 = o_Describe(I) And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And 方法id = 2 Or
                运算符 = 4 And o_Describe(I) >= 指标值上限 And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And 方法id = 2 Or
                运算符 = 5 And o_Describe(I) <= 指标值下限 And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And 方法id = 2;
          If Illness_Level Is Null Then
            Select Nvl(Min(病情级别), Level_Max)
            Into Illness_Level
            From 急诊评分方法规则
            Where 运算符 = 2 And o_Describe(I) > 指标值上限 And 指标id = o_Indexid(I) And 方法id = 2 And 指标年龄id Is Null Or
                  运算符 = 3 And o_Describe(I) < 指标值下限 And 指标id = o_Indexid(I) And 方法id = 2 And 指标年龄id Is Null Or
                  运算符 = 6 And o_Describe(I) Between 指标值下限 And 指标值上限 And 指标id = o_Indexid(I) And 方法id = 2 And
                  指标年龄id Is Null Or
                  运算符 = 1 And 指标值上限 = o_Describe(I) And 指标id = o_Indexid(I) And 方法id = 2 And 指标年龄id Is Null Or
                  运算符 = 4 And o_Describe(I) >= 指标值上限 And 指标id = o_Indexid(I) And 方法id = 2 And 指标年龄id Is Null Or
                  运算符 = 5 And o_Describe(I) <= 指标值下限 And 指标id = o_Indexid(I) And 方法id = 2 And 指标年龄id Is Null;
          
          End If;
          If Illness_Level < Level_Max Then
            Level_Max := Illness_Level;
          End If;
        
        End Loop;
        Return Level_Max;
      End If;
    
      If Ageunit = '月' Then
        Select C1, C2 Bulk Collect Into o_Indexid, o_Describe From Table(f_Str2list2(Indexid_Describe));
        For I In 1 .. o_Indexid.Count Loop
          Select Max(ID)
          Into Age_Id
          From 急诊评分指标年龄
          Where 运算符 = 2 And Agenum > 年龄上限 And 指标id = o_Indexid(I) And 年龄单位 = '月' Or
                运算符 = 3 And Agenum < 年龄下限 And 指标id = o_Indexid(I) And 年龄单位 = '月' And 年龄单位 = '月' Or
                运算符 = 6 And Agenum Between 年龄下限 And 年龄上限 And 指标id = o_Indexid(I) And 年龄单位 = '月' Or
                运算符 = 1 And 年龄上限 = Agenum And 指标id = o_Indexid(I) And 年龄单位 = '月' Or
                运算符 = 4 And Agenum >= 年龄上限 And 指标id = o_Indexid(I) And 年龄单位 = '月' Or
                运算符 = 5 And Agenum <= 年龄下限 And 指标id = o_Indexid(I) And 年龄单位 = '月';
          Select Min(病情级别)
          Into Illness_Level
          From 急诊评分方法规则
          Where 运算符 = 2 And o_Describe(I) > 指标值上限 And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And 方法id = 2 Or
                运算符 = 3 And o_Describe(I) < 指标值下限 And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And 方法id = 2 Or
                运算符 = 6 And o_Describe(I) Between 指标值下限 And 指标值上限 And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And
                方法id = 2 Or 运算符 = 1 And 指标值上限 = o_Describe(I) And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And 方法id = 2 Or
                运算符 = 4 And o_Describe(I) >= 指标值上限 And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And 方法id = 2 Or
                运算符 = 5 And o_Describe(I) <= 指标值下限 And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And 方法id = 2;
        
          If Illness_Level Is Null Then
            Select Nvl(Min(病情级别), Level_Max)
            Into Illness_Level
            From 急诊评分方法规则
            Where 运算符 = 2 And o_Describe(I) > 指标值上限 And 指标id = o_Indexid(I) And 方法id = 2 And 指标年龄id Is Null Or
                  运算符 = 3 And o_Describe(I) < 指标值下限 And 指标id = o_Indexid(I) And 方法id = 2 And 指标年龄id Is Null Or
                  运算符 = 6 And o_Describe(I) Between 指标值下限 And 指标值上限 And 指标id = o_Indexid(I) And 方法id = 2 And
                  指标年龄id Is Null Or
                  运算符 = 1 And 指标值上限 = o_Describe(I) And 指标id = o_Indexid(I) And 方法id = 2 And 指标年龄id Is Null Or
                  运算符 = 4 And o_Describe(I) >= 指标值上限 And 指标id = o_Indexid(I) And 方法id = 2 And 指标年龄id Is Null Or
                  运算符 = 5 And o_Describe(I) <= 指标值下限 And 指标id = o_Indexid(I) And 方法id = 2 And 指标年龄id Is Null;
          
          End If;
          If Illness_Level < Level_Max Then
            Level_Max := Illness_Level;
          End If;
        
        End Loop;
        Return Level_Max;
      End If;
    
      If Ageunit = '天' Then
        Select C1, C2 Bulk Collect Into o_Indexid, o_Describe From Table(f_Str2list2(Indexid_Describe));
        For I In 1 .. o_Indexid.Count Loop
          Select Max(ID)
          Into Age_Id
          From 急诊评分指标年龄
          Where 运算符 = 2 And Agenum > 年龄上限 And 指标id = o_Indexid(I) And 年龄单位 = '天' Or
                运算符 = 3 And Agenum < 年龄下限 And 指标id = o_Indexid(I) And 年龄单位 = '天' And 年龄单位 = '天' Or
                运算符 = 6 And Agenum Between 年龄下限 And 年龄上限 And 指标id = o_Indexid(I) And 年龄单位 = '天' Or
                运算符 = 1 And 年龄上限 = Agenum And 指标id = o_Indexid(I) And 年龄单位 = '天' Or
                运算符 = 4 And Agenum >= 年龄上限 And 指标id = o_Indexid(I) And 年龄单位 = '天' Or
                运算符 = 5 And Agenum <= 年龄下限 And 指标id = o_Indexid(I) And 年龄单位 = '天';
          Select Min(病情级别)
          Into Illness_Level
          From 急诊评分方法规则
          Where 运算符 = 2 And o_Describe(I) > 指标值上限 And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And 方法id = 2 Or
                运算符 = 3 And o_Describe(I) < 指标值下限 And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And 方法id = 2 Or
                运算符 = 6 And o_Describe(I) Between 指标值下限 And 指标值上限 And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And
                方法id = 2 Or 运算符 = 1 And 指标值上限 = o_Describe(I) And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And 方法id = 2 Or
                运算符 = 4 And o_Describe(I) >= 指标值上限 And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And 方法id = 2 Or
                运算符 = 5 And o_Describe(I) <= 指标值下限 And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And 方法id = 2;
          If Illness_Level Is Null Then
            Select Nvl(Min(病情级别), Level_Max)
            Into Illness_Level
            From 急诊评分方法规则
            Where 运算符 = 2 And o_Describe(I) > 指标值上限 And 指标id = o_Indexid(I) And 方法id = 2 And 指标年龄id Is Null Or
                  运算符 = 3 And o_Describe(I) < 指标值下限 And 指标id = o_Indexid(I) And 方法id = 2 And 指标年龄id Is Null Or
                  运算符 = 6 And o_Describe(I) Between 指标值下限 And 指标值上限 And 指标id = o_Indexid(I) And 方法id = 2 And
                  指标年龄id Is Null Or
                  运算符 = 1 And 指标值上限 = o_Describe(I) And 指标id = o_Indexid(I) And 方法id = 2 And 指标年龄id Is Null Or
                  运算符 = 4 And o_Describe(I) >= 指标值上限 And 指标id = o_Indexid(I) And 方法id = 2 And 指标年龄id Is Null Or
                  运算符 = 5 And o_Describe(I) <= 指标值下限 And 指标id = o_Indexid(I) And 方法id = 2 And 指标年龄id Is Null;
          End If;
        
          If Illness_Level < Level_Max Then
            Level_Max := Illness_Level;
          End If;
        
        End Loop;
        Return Level_Max;
      End If;
    
    End If;
    Return Level_Max;
  Exception
    --异常处理语句段
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Is_Objective_Rating;

End b_Emergency_Rating;
/


--145003:蒋廷中,2019-10-15,新增模块急诊预检分诊工作站
Create Or Replace Package Pkg_Pretriage_Dql As
  -----------------------------------------------------
  --获取报表列表
  -----------------------------------------------------
  Procedure Get_Reportlist
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );
  -----------------------------------------------------
  --检查身份证录入是否正确
  -----------------------------------------------------
  Procedure Checkidcard
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  );

  -----------------------------------------------------
  --检查年龄录入是否正确
  -----------------------------------------------------
  Procedure Checkage
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  );
  -----------------------------------------------------
  --通过医保号读取病人信息列表清单
  -----------------------------------------------------
  Procedure Get_Patlistbymedical
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --通过身份证号读取病人信息列表清单
  -----------------------------------------------------
  Procedure Get_Patlistbyidcard
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --通过输入姓名匹配病人信息列表清单
  -----------------------------------------------------
  Procedure Get_Patlistbyname
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取最新的就诊状态
  -----------------------------------------------------
  Procedure Getvisitstate
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取病人分诊评分信息
  -----------------------------------------------------
  Procedure Load_Levelinfo
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取病人分诊指标信息
  -----------------------------------------------------
  Procedure Load_Rulesinfo
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取病人分诊记录内容
  -----------------------------------------------------
  Procedure Load_Pretriage
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取单个病人分诊信息
  -----------------------------------------------------
  Procedure Get_Patidetail
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );
  -----------------------------------------------------
  --获取病人列表清单
  -----------------------------------------------------
  Procedure Get_Patlist
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --疼痛分级方法
  -----------------------------------------------------
  Procedure Get_Pain_Num_Rating
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --昏迷评分分级方法
  -----------------------------------------------------
  Procedure Get_Coma_Rating
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --客观评价分级方法成人和儿童方法
  -----------------------------------------------------
  Procedure Get_Objective_Rating
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取儿童年龄上限
  -----------------------------------------------------
  Procedure Get_Childmaxage
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  );
  -----------------------------------------------------
  --获取急诊等级
  -----------------------------------------------------
  Procedure Get_Level
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取急诊科室
  -----------------------------------------------------
  Procedure Get_Dept
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取人工评估规则
  -----------------------------------------------------
  Procedure Get_Rules
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );
  -----------------------------------------------------
  --根据出生日期返回年龄
  -----------------------------------------------------
  Procedure Get_Datetoage
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  );
  -----------------------------------------------------
  --获取性别基础数据
  -----------------------------------------------------
  Procedure Get_Sexbase
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取民族基础数据
  -----------------------------------------------------
  Procedure Get_Nationbase
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取急诊评分指标
  -----------------------------------------------------
  Procedure Get_Scorebase
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );
  -----------------------------------------------------
  --获取急诊主诉
  -----------------------------------------------------
  Procedure Get_Paticc
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取病人来源
  -----------------------------------------------------
  Procedure Get_Patifrom
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取急诊意识状态
  -----------------------------------------------------
  Procedure Get_Patistate
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取急诊陪同人员
  -----------------------------------------------------
  Procedure Get_Entourage
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取急诊常见既往史
  -----------------------------------------------------
  Procedure Get_Dishistory
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取数据库系统时间
  -----------------------------------------------------
  Procedure Get_Now_Time
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  );

End Pkg_Pretriage_Dql;
/
Create Or Replace Package Body Pkg_Pretriage_Dql As
  -----------------------------------------------------
  --获取报表列表
  -----------------------------------------------------
  Procedure Get_Reportlist
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select 标志, 系统, 编号, 名称, Nvl(是否停用, 0) 是否停用
      From (Select 1 As 标志, a.系统, a.编号, a.名称, a.是否停用
             From zlReports A, zlPrograms B
             Where a.系统 = b.系统 And a.程序id = b.序号 And Not Upper(a.编号) Like '%BILL%' And Upper(b.部件) <> Upper('zl9Report') And
                   b.系统 = 100 And b.序号 = 1244
             Union All
             Select Decode(a.系统, Null, 2, 1) As 标志, a.系统, a.编号, a.名称, a.是否停用
             From zlReports A, zlRPTPuts B, zlPrograms C
             Where a.Id = b.报表id And b.系统 = c.系统 And b.程序id = c.序号 And (Not Upper(a.编号) Like '%BILL%' Or a.系统 Is Null) And
                   c.系统 = 100 And c.序号 = 1244)
      Where Instr(',ZL1_REPORT_1244_1,ZL1_REPORT_1244_2,', ',' || 编号 || ',') = 0 And Nvl(是否停用, 0) = 0
      Order By 标志, 编号;
  End Get_Reportlist;

  -----------------------------------------------------
  --检查身份证录入是否正确
  -----------------------------------------------------
  Procedure Checkidcard
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  ) Is
    --入参：Json_In:格式
    --  input
    --    idcard        C 1 录入的身份证号
    -- 返回值：固定格式XML串
    --<OUTPUT>
    --       <BIRTHDAY></BIRTHDAY>                //出生日期
    --       <SEX></SEX>                  //性别
    --       <AGE></AGE>                //年龄
    --     <MSG></MSG>         //空串-身份证号有效(可从身份证号中获取出生日期和性别)，非空串-返回错误信息
    --</OUTPUT>
  
    Jsonobj  Pljson;
    v_录入项 Varchar2(50);
  Begin
    Jsonobj    := Pljson(Input_In);
    v_录入项   := Pljson_Ext.Get_String(Jsonobj, 'input.idcard');
    Output_Out := Zl_Fun_Checkidcard(v_录入项);
  End Checkidcard;

  -----------------------------------------------------
  --检查年龄录入是否正确
  -----------------------------------------------------
  Procedure Checkage
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  ) Is
    --入参：Json_In:格式
    --  input
    --    age        C 1 年龄
    Jsonobj Pljson;
    v_年龄  Varchar2(50);
  Begin
    Jsonobj    := Pljson(Input_In);
    v_年龄     := Pljson_Ext.Get_String(Jsonobj, 'input.age');
    Output_Out := Zl_Age_Check(v_年龄);
  End Checkage;

  -----------------------------------------------------
  --通过医保号读取病人信息列表清单
  -----------------------------------------------------
  Procedure Get_Patlistbymedical
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    --入参：Json_In:格式
    Jsonobj    Pljson;
    v_医保号   Varchar2(200);
    v_医保类型 Varchar2(200);
  Begin
    Jsonobj    := Pljson(Input_In);
    v_医保号   := Pljson_Ext.Get_String(Jsonobj, 'input.医保号');
    v_医保类型 := Pljson_Ext.Get_String(Jsonobj, 'input.医保类型');
    Open Output_Out For
      Select /*+Rule */
      Distinct a.病人id, a.姓名, a.性别, a.年龄, a.门诊号, To_Char(a.出生日期, 'yyyy-MM-dd') As 出生日期, a.民族, a.身份证号, a.手机号, a.医保号,
               b.名称 As 保险类别, a.家庭地址, Nvl(就诊时间, 登记时间) As 就诊时间
      From 病人信息 A, 保险类别 B
      Where a.险类 = b.序号(+) And a.停用时间 Is Null And a.医保号 = v_医保号 And b.名称 = v_医保类型
      Order By 病人id Desc;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Patlistbymedical;

  -----------------------------------------------------
  --通过身份证号读取病人信息列表清单
  -----------------------------------------------------
  Procedure Get_Patlistbyidcard
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    --入参：Json_In:格式
    Jsonobj    Pljson;
    v_身份证号 Varchar2(200);
  Begin
    Jsonobj    := Pljson(Input_In);
    v_身份证号 := Pljson_Ext.Get_String(Jsonobj, 'input.身份证号');
    Open Output_Out For
      Select /*+Rule */
      Distinct a.病人id, a.姓名, a.性别, a.年龄, a.门诊号, To_Char(a.出生日期, 'yyyy-MM-dd') As 出生日期, a.民族, a.身份证号, a.手机号, a.医保号,
               b.名称 As 保险类别, a.家庭地址, Nvl(就诊时间, 登记时间) As 就诊时间
      From 病人信息 A, 保险类别 B
      Where a.险类 = b.序号(+) And a.停用时间 Is Null And a.身份证号 = v_身份证号
      Order By 病人id Desc;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Patlistbyidcard;

  -----------------------------------------------------
  --通过输入姓名匹配病人信息列表清单
  -----------------------------------------------------
  Procedure Get_Patlistbyname
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    --入参：Json_In:格式
    Jsonobj Pljson;
    v_姓名  Varchar2(200);
  Begin
    Jsonobj := Pljson(Input_In);
    v_姓名  := Pljson_Ext.Get_String(Jsonobj, 'input.姓名输入');
    Open Output_Out For
      Select 1 As 排序id, a.病人id, a.姓名, a.性别, a.年龄, a.门诊号, To_Char(a.出生日期, 'yyyy-MM-dd') As 出生日期, a.民族, a.身份证号, a.手机号,
             a.医保号, b.名称 As 保险类别, a.家庭地址, Nvl(就诊时间, 登记时间) As 就诊时间
      From 病人信息 A, 保险类别 B
      Where a.险类 = b.序号(+) And a.停用时间 Is Null And (a.身份证号 = v_姓名)
      Union All
      Select 1 As 排序id, a.*
      From (Select /*+Rule */
             Distinct a.病人id, a.姓名, a.性别, a.年龄, a.门诊号, To_Char(a.出生日期, 'yyyy-MM-dd') As 出生日期, a.民族, a.身份证号, a.手机号, a.医保号,
                      b.名称 As 保险类别, a.家庭地址, Nvl(就诊时间, 登记时间) As 就诊时间
             From 病人信息 A, 保险类别 B
             Where a.险类 = b.序号(+) And a.停用时间 Is Null And a.姓名 Like v_姓名 || '%'
             Order By 就诊时间 Desc) A
      Where Rownum < 101
      Union All
      Select 0 As 排序id, -null, '[新病人]', Null, Null, -null, Null, Null, Null, Null, Null, Null, Null, To_Date(Null)
      From Dual
      Order By 排序id;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Patlistbyname;

  -----------------------------------------------------
  --获取最新的就诊状态
  -----------------------------------------------------
  Procedure Getvisitstate
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj Pljson;
    n_Id    急诊分诊记录.Id%Type; --分诊ID
  Begin
    Jsonobj := Pljson(Input_In);
    n_Id    := To_Number(Pljson_Ext.Get_String(Jsonobj, 'input.id'));
    Open Output_Out For
      Select Max(Decode(Nvl(c.执行状态, 0), 0, 0, 1)) As 就诊状态
      From 急诊就诊记录 A, 急诊分诊记录 B, 病人挂号记录 C
      Where a.Id = b.就诊id And a.挂号id = c.Id And b.Id = n_Id;
  End Getvisitstate;

  -----------------------------------------------------
  --获取病人分诊评分信息
  -----------------------------------------------------
  Procedure Load_Levelinfo
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj Pljson;
    n_Id    急诊分诊记录.Id%Type; --分诊ID
  Begin
    Jsonobj := Pljson(Input_In);
    n_Id    := To_Number(Pljson_Ext.Get_String(Jsonobj, 'input.id'));
    Open Output_Out For
      Select ID, 分诊id, 方法id, 评分方法分值, 评分结果描述, 病情级别 From 急诊病人评分 Where 分诊id = n_Id;
  End Load_Levelinfo;

  -----------------------------------------------------
  --获取病人分诊指标信息
  -----------------------------------------------------
  Procedure Load_Rulesinfo
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj Pljson;
    n_Id    急诊分诊记录.Id%Type; --分诊ID
  Begin
    Jsonobj := Pljson(Input_In);
    n_Id    := To_Number(Pljson_Ext.Get_String(Jsonobj, 'input.id'));
    Open Output_Out For
      Select a.评分id, b.方法id, a.指标id, a.指标结果文本
      From 急诊病人评分指标 A, 急诊病人评分 B
      Where a.评分id = b.Id And b.分诊id = n_Id;
  End Load_Rulesinfo;

  -----------------------------------------------------
  --获取病人分诊记录内容
  -----------------------------------------------------
  Procedure Load_Pretriage
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj Pljson;
    n_Id    急诊分诊记录.Id%Type; --分诊ID
  Begin
    Jsonobj := Pljson(Input_In);
    n_Id    := To_Number(Pljson_Ext.Get_String(Jsonobj, 'input.id'));
    Open Output_Out For
      Select a.病人id, a.姓名, a.性别, a.年龄, To_Char(a.出生日期, 'yyyy-MM-dd') As 出生日期, a.身份证号, a.民族, a.家庭地址, d.名称 As 保险类别, a.医保号,
             a.手机号, b. ID As 就诊id, b. 病人id, b. 病人年龄, b. 年龄数值, b. 年龄单位, b. 挂号id, b. 病情级别,
             To_Char(b. 到院时间, 'yyyy-MM-dd HH24:mi') As 到院时间, b. 主诉, b. 是否三无人员, b. 陪同人员, b. 病人来源, b. 既往病史, b. 意识状态,
             b. 是否成批就诊, b. 成批就诊人数, b. 是否复合伤, b. 备注, b. 登记人 As 就诊登记人, b. 登记时间 As 就诊登记时间, c.修改说明, c.Id As 分诊id, c.分诊次数,
             c.自动病情级别, c.分诊科室id, c.分诊科室名称, c.收缩压, c.舒张压, c.心率, c.指氧饱和度, c.体温, c.血糖, c.血钾,
             To_Char(c.体征测量时间, 'yyyy-MM-dd HH24:mi') As 体征测量时间, c.登记人, c.登记时间, c.人工病情级别, c.人工评级说明, c.呼吸频率, b. 是否绿色通道
      From 病人信息 A, 急诊就诊记录 B, 急诊分诊记录 C, 保险类别 D
      Where a.病人id = b.病人id And b.Id = c.就诊id And a.险类 = d.序号(+) And c.Id = n_Id;
  End Load_Pretriage;

  -----------------------------------------------------
  --获取单个病人分诊信息
  -----------------------------------------------------
  Procedure Get_Patidetail
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    --入参：Json_In:格式
    Jsonobj   Pljson;
    n_Id      急诊分诊记录.就诊id%Type;
    n_Max序号 急诊分诊记录.分诊次数%Type;
  Begin
    Jsonobj := Pljson(Input_In);
    n_Id    := To_Number(Pljson_Ext.Get_String(Jsonobj, 'input.id'));
  
    Select Max(分诊次数) Into n_Max序号 From 急诊分诊记录 Where 就诊id = n_Id;
  
    Open Output_Out For
      Select a.Id 分诊id, a.分诊次数, a.自动病情级别 || '级' As 自动病情级别, a.人工病情级别 || '级' As 人工病情级别,
             '第' || a.分诊次数 || '次分诊    自动评级（' || a.自动病情级别 || '级）' ||
              Decode(a.人工病情级别, '', '', '    人工评级（' || a.人工病情级别 || '级）') ||
              Decode(n_Max序号, a.分诊次数,
                     Decode(Nvl(b.病情级别 || '', '0'), Nvl(b.分诊病情级别 || '', '0'), '',
                             '    修订病情级别（' || Nvl(b.病情级别 || '', '0') || '级）')) || '    分诊时间：' ||
              To_Char(a.登记时间, 'yyyy-MM-dd HH24:mi') As 病情情况
      From 急诊分诊记录 A, 急诊就诊记录 B
      Where a.就诊id = b.Id And 就诊id = n_Id
      Order By 分诊次数 Desc;
  End Get_Patidetail;

  -----------------------------------------------------
  --获取病人列表清单
  -----------------------------------------------------
  Procedure Get_Patlist
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    --入参：Json_In:格式
    Jsonobj    Pljson;
    d_开始时间 急诊就诊记录.登记时间%Type;
    d_结束时间 急诊就诊记录.登记时间%Type;
    v_分诊状态 Varchar2(10);
    n_已超时   Number(2); -- =1 仅过滤已超时病人
  Begin
    Jsonobj    := Pljson(Input_In);
    d_开始时间 := To_Date(Pljson_Ext.Get_String(Jsonobj, 'input.begin'), 'yyyy-mm-dd hh24:mi:ss');
    d_结束时间 := To_Date(Pljson_Ext.Get_String(Jsonobj, 'input.end'), 'yyyy-mm-dd hh24:mi:ss');
    v_分诊状态 := Pljson_Ext.Get_String(Jsonobj, 'input.state');
    n_已超时   := Nvl(To_Number(Pljson_Ext.Get_String(Jsonobj, 'input.timeout')), 0);
  
    If n_已超时 = 1 Then
      Open Output_Out For
        Select b.病人id, b.Id 就诊序号, a.姓名, a.性别, a.年龄, To_Char(b.登记时间, 'yyyy-MM-dd HH24:mi') As 登记时间, b.登记人 分诊护士,
               b.病情级别 || '级' As 病情级别, Decode(Nvl(d.执行状态, 0), 0, 0, 1) As 就诊状态, b.是否绿色通道
        From 病人信息 A, 急诊就诊记录 B, 急诊病情级别 C, 病人挂号记录 D
        Where a.病人id = b.病人id And b.挂号id = d.Id(+) And b.病情级别 = c.序号 And b.登记时间 >= d_开始时间 And b.登记时间 < d_结束时间 And
              Decode(Nvl(d.执行状态, 0), 0, 0, 1) In
              (Select Column_Value From Table(Cast(f_Str2list(v_分诊状态) As t_Strlist))) And
              (c.再次评估时限 Is Not Null And (b.登记时间 + (Nvl(c.再次评估时限, 0) / 24 / 60)) < Sysdate);
    Else
      Open Output_Out For
        Select b.病人id, b.Id 就诊序号, a.姓名, a.性别, a.年龄, To_Char(b.登记时间, 'yyyy-MM-dd HH24:mi') As 登记时间, b.登记人 分诊护士,
               b.病情级别 || '级' As 病情级别, Decode(Nvl(d.执行状态, 0), 0, 0, 1) As 就诊状态, b.是否绿色通道
        From 病人信息 A, 急诊就诊记录 B, 病人挂号记录 D
        Where a.病人id = b.病人id And b.挂号id = d.Id(+) And b.登记时间 >= d_开始时间 And b.登记时间 < d_结束时间 And
              Decode(Nvl(d.执行状态, 0), 0, 0, 1) In
              (Select Column_Value From Table(Cast(f_Str2list(v_分诊状态) As t_Strlist)));
    End If;
  End Get_Patlist;

  -----------------------------------------------------
  --疼痛分级方法
  -----------------------------------------------------
  Procedure Get_Pain_Num_Rating
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    --入参：Json_In:格式
    --  input
    --    pain        C 1 疼痛等级
    Jsonobj    Pljson;
    v_疼痛等级 Varchar2(200);
    v_Out      Varchar2(200);
    v_病人等级 Varchar2(200);
    v_病人分数 Varchar2(200);
    v_描述     Varchar2(200);
  Begin
    Jsonobj    := Pljson(Input_In);
    v_疼痛等级 := Pljson_Ext.Get_String(Jsonobj, 'input.pain');
  
    Select b_Emergency_Rating.Is_Pain_Num_Rating(v_疼痛等级) Into v_Out From Dual;
  
    v_病人等级 := Substr(v_Out, 1, Instr(v_Out, ':', 1, 1) - 1);
    v_病人分数 := Substr(v_Out, Instr(v_Out, ':', 1, 1) + 1, Instr(v_Out, ':', 1, 2) - Instr(v_Out, ':', 1, 1) - 1);
    v_描述     := Substr(v_Out, Instr(v_Out, ':', 1, 2) + 1);
  
    Open Output_Out For
      Select v_病人等级 As 病人等级, v_病人分数 As 病人分数, v_描述 As 描述 From Dual;
  End Get_Pain_Num_Rating;

  -----------------------------------------------------
  --昏迷评分分级方法
  -----------------------------------------------------
  Procedure Get_Coma_Rating
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    --入参：Json_In:格式
    --  input
    --    open_reaction            C 1 睁眼反应指标id：指标结果描述
    --    language_reaction        C 1 语言反应指标id：指标结果描述
    --    activity_reaction        C 1 活动反应指标id：指标结果描述
    Jsonobj    Pljson;
    v_睁眼反应 Varchar2(200);
    v_语言反应 Varchar2(200);
    v_活动反应 Varchar2(200);
    v_Out      Varchar2(200);
  
    v_病人等级 Varchar2(200);
    v_病人分数 Varchar2(200);
    v_描述     Varchar2(200);
  Begin
    Jsonobj    := Pljson(Input_In);
    v_睁眼反应 := Pljson_Ext.Get_String(Jsonobj, 'input.open_reaction');
    v_语言反应 := Pljson_Ext.Get_String(Jsonobj, 'input.language_reaction');
    v_活动反应 := Pljson_Ext.Get_String(Jsonobj, 'input.activity_reaction');
  
    Select b_Emergency_Rating.Is_Coma_Rating(v_睁眼反应, v_语言反应, v_活动反应) Into v_Out From Dual;
  
    v_病人等级 := Substr(v_Out, 1, Instr(v_Out, ':', 1, 1) - 1);
    v_病人分数 := Substr(v_Out, Instr(v_Out, ':', 1, 1) + 1, Instr(v_Out, ':', 1, 2) - Instr(v_Out, ':', 1, 1) - 1);
    v_描述     := Substr(v_Out, Instr(v_Out, ':', 1, 2) + 1);
  
    Open Output_Out For
      Select v_病人等级 As 病人等级, v_病人分数 As 病人分数, v_描述 As 描述 From Dual;
  End Get_Coma_Rating;

  -----------------------------------------------------
  --客观评价分级方法成人和儿童方法
  -----------------------------------------------------
  Procedure Get_Objective_Rating
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    --入参：Json_In:格式
    --  input
    --    agenum                   C 1 年龄
    --    ageunit                  C 1 年龄单位
    --    indexid_describe         C 1 指标id：指标结果描述（可多个）
    Jsonobj    Pljson;
    n_年龄     Number;
    v_年龄单位 Varchar2(200);
    v_指标信息 Varchar2(200);
    v_Out      Varchar2(200);
  Begin
    Jsonobj    := Pljson(Input_In);
    n_年龄     := Nvl(Zl_To_Number(Pljson_Ext.Get_String(Jsonobj, 'input.agenum')), 0);
    v_年龄单位 := Pljson_Ext.Get_String(Jsonobj, 'input.ageunit');
    v_指标信息 := Pljson_Ext.Get_String(Jsonobj, 'input.indexid_describe');
  
    Select b_Emergency_Rating.Is_Objective_Rating(n_年龄, v_年龄单位, v_指标信息) Into v_Out From Dual;
  
    Open Output_Out For
      Select v_Out As 病人等级 From Dual;
  End Get_Objective_Rating;
  -----------------------------------------------------
  --获取儿童年龄上限
  -----------------------------------------------------
  Procedure Get_Childmaxage
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  ) Is
  Begin
    Output_Out := Nvl(zl_GetSysParameter('儿童年龄界定上限'), 0);
  End Get_Childmaxage;

  -----------------------------------------------------
  --获取急诊等级
  -----------------------------------------------------
  Procedure Get_Level
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select a.序号, a.名称, a.严重程度, a.再次评估时限, a.患者标识颜色, Null As 缺省
      From 急诊病情级别 A
      Order By a.序号;
  End Get_Level;

  -----------------------------------------------------
  --获取急诊科室
  -----------------------------------------------------
  Procedure Get_Dept
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select a.Id, a.编码, a.名称, a.简码, Null As 缺省
      From 部门表 A, 临床部门 B
      Where a.Id = b.部门id And b.工作性质 = '20' And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)
      Order By a.编码;
  End Get_Dept;

  -----------------------------------------------------
  --获取人工评估规则
  -----------------------------------------------------
  Procedure Get_Rules
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select a.Id, a.分类, a.指标名称, a.适用人群, a.病情级别 From 急诊人工评定规则 A Order By ID, 病情级别;
  End Get_Rules;

  -----------------------------------------------------
  --根据出生日期返回年龄
  -----------------------------------------------------
  Procedure Get_Datetoage
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  ) Is
    --入参：Json_In:格式
    --  input
    --    birthday        C 1 出生日期 yyyy-mm-dd
    Jsonobj    Pljson;
    d_出生日期 Date;
    v_年龄     Varchar2(50);
  Begin
    Jsonobj    := Pljson(Input_In);
    d_出生日期 := To_Date(Pljson_Ext.Get_String(Jsonobj, 'input.birthday'), 'yyyy-mm-dd hh24:mi:ss');
    Select Zl_Age_Calc(0, d_出生日期, Sysdate) Into v_年龄 From Dual;
  
    Output_Out := v_年龄;
  End Get_Datetoage;

  -----------------------------------------------------
  --获取性别基础数据
  -----------------------------------------------------
  Procedure Get_Sexbase
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select 编码, 名称, 简码, Nvl(缺省标志, 0) As 缺省 From 性别 Order By Nvl(缺省标志, 0) Desc, 编码;
  End Get_Sexbase;

  -----------------------------------------------------
  --获取民族基础数据
  -----------------------------------------------------
  Procedure Get_Nationbase
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select 编码, 名称, 简码, Nvl(缺省标志, 0) As 缺省 From 民族 Order By Nvl(缺省标志, 0) Desc, 编码;
  End Get_Nationbase;

  -----------------------------------------------------
  --获取急诊评分指标
  -----------------------------------------------------
  Procedure Get_Scorebase
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select ID, 指标名称, 值域范围, 方法id, 值域单位 From 急诊评分指标 Order By ID;
  End Get_Scorebase;

  -----------------------------------------------------
  --获取急诊主诉
  -----------------------------------------------------
  Procedure Get_Paticc
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select b.名称 分类, a.编码, a.名称, a.简码
      From 急诊常用主诉 A, 急诊常用主诉 B
      Where a.上级 = b.编码 And a.上级 Is Not Null And b.上级 Is Null
      Order By b.编码;
  End Get_Paticc;

  -----------------------------------------------------
  --获取病人来源
  -----------------------------------------------------
  Procedure Get_Patifrom
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select 编码, 名称, 缺省标志 As 缺省 From 急诊病人来源 Order By Nvl(缺省标志, 0) Desc, 名称;
  End Get_Patifrom;

  -----------------------------------------------------
  --获取急诊意识状态
  -----------------------------------------------------
  Procedure Get_Patistate
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select 编码, 名称, 缺省标志 As 缺省 From 急诊意识状态 Order By Nvl(缺省标志, 0) Desc, 名称;
  End Get_Patistate;

  -----------------------------------------------------
  --获取急诊陪同人员
  -----------------------------------------------------
  Procedure Get_Entourage
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select 编码, 名称, 缺省标志 As 缺省 From 急诊陪同人员 Order By Nvl(缺省标志, 0) Desc, 名称;
  End Get_Entourage;

  -----------------------------------------------------
  --获取急诊常见既往史
  -----------------------------------------------------
  Procedure Get_Dishistory
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select 编码, 名称, 0 As 缺省 From 急诊常见既往史 Order By 名称;
  End Get_Dishistory;

  -----------------------------------------------------
  --获取数据库系统时间
  -----------------------------------------------------
  Procedure Get_Now_Time
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  ) Is
  Begin
    Output_Out := To_Char(Sysdate, 'yyyy-MM-dd HH24:mi');
  End Get_Now_Time;
End Pkg_Pretriage_Dql;
/

--145003:蒋廷中,2019-10-15,新增模块急诊预检分诊工作站
Create Or Replace Package Pkg_Pretriage_Dml As

  -----------------------------------------------------
  --变更病人就诊记录的绿色通道状态
  -----------------------------------------------------
  Procedure Change_Greenchannel
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  );
  -----------------------------------------------------
  --删除病人就诊记录
  -----------------------------------------------------
  Procedure Del_Pretriage
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  );

  -----------------------------------------------------
  --清除挂号事务锁定
  -----------------------------------------------------
  Procedure Register_Unlock
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  );

  -----------------------------------------------------
  --更新最新的挂号安排
  -----------------------------------------------------
  Procedure Register_Update
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  );

  -----------------------------------------------------
  --保存病人分诊信息
  -----------------------------------------------------
  Procedure Save_Pretriage
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );
End Pkg_Pretriage_Dml;
/
Create Or Replace Package Body Pkg_Pretriage_Dml As
  -----------------------------------------------------
  --变更病人就诊记录的绿色通道状态
  -----------------------------------------------------
  Procedure Change_Greenchannel
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  ) As
    --功能：标记或取消急诊绿色通道
    Jsonobj        Pljson;
    n_Id           急诊就诊记录.Id%Type; --就诊ID
    n_是否绿色通道 急诊就诊记录.是否绿色通道%Type; --是否绿色通道
    n_挂号id       急诊就诊记录.挂号id %Type;
  Begin
    Jsonobj        := Pljson(Input_In);
    n_Id           := To_Number(Pljson_Ext.Get_String(Jsonobj, 'input.id'));
    n_是否绿色通道 := To_Number(Pljson_Ext.Get_String(Jsonobj, 'input.是否绿色通道'));
  
    Select Max(挂号id) Into n_挂号id From 急诊就诊记录 Where ID = n_Id;
  
    Zl_急诊绿色通道_Edit(n_挂号id, n_是否绿色通道);
    Output_Out := '成功';
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Change_Greenchannel;

  -----------------------------------------------------
  --删除病人就诊记录
  -----------------------------------------------------
  Procedure Del_Pretriage
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  ) Is
    Jsonobj  Pljson;
    n_Id     急诊就诊记录.Id%Type; --就诊ID
    n_病人id 急诊就诊记录.病人id%Type;
    n_挂号id 急诊就诊记录.挂号id%Type;
  Begin
    Jsonobj := Pljson(Input_In);
    n_Id    := To_Number(Pljson_Ext.Get_String(Jsonobj, 'input.id'));
  
    Delete From 急诊就诊记录 Where ID = n_Id Return 病人id, 挂号id Into n_病人id, n_挂号id;
  
    Zl_Emergencyregistdel(n_挂号id);
  
    Delete From 病人信息从表
    Where 病人id = n_病人id And 就诊id = n_挂号id And 信息名 In ('体温', '呼吸', '脉搏', '收缩压', '舒张压', '血糖');
  
    Output_Out := '成功';
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Del_Pretriage;

  -----------------------------------------------------
  --清除挂号事务锁定
  -----------------------------------------------------
  Procedure Register_Unlock
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  ) Is
    v_人员姓名 Varchar2(200);
    v_Temp     Varchar2(4000);
  Begin
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    Zl_挂号序号状态_Lock(2, v_人员姓名);
    Output_Out := '成功';
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Register_Unlock;

  -----------------------------------------------------
  --更新最新的挂号安排
  -----------------------------------------------------
  Procedure Register_Update
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  ) Is
  Begin
    Zl_挂号安排_Autoupdate();
    Output_Out := '成功';
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Register_Update;

  -----------------------------------------------------
  --保存病人分诊信息
  -----------------------------------------------------
  Procedure Save_Pretriage
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    --入参：Json_In:格式
    Jsonobj Pljson;
    n_Type  Number; --1  新增,2  修改
  
    --病人信息
    v_姓名     病人信息.姓名%Type;
    v_性别     病人信息.性别%Type;
    d_出生日期 病人信息.出生日期%Type;
    v_身份证号 病人信息.身份证号%Type;
    v_联系电话 病人信息.联系人电话%Type;
    v_民族     病人信息.民族%Type;
    v_医保卡号 病人信息.医保号%Type;
    v_保险类别 保险类别.名称%Type;
    v_家庭地址 病人信息.家庭地址%Type;
  
    n_就诊id 急诊就诊记录.Id%Type;
    n_病人id 急诊就诊记录.病人id%Type;
    n_挂号id 急诊就诊记录.挂号id%Type;
    n_分诊id 急诊分诊记录.Id%Type;
  
    --就诊记录
    v_病人年龄 急诊就诊记录.病人年龄%Type;
    n_年龄数值 急诊就诊记录.年龄数值%Type;
    v_年龄单位 急诊就诊记录.年龄单位%Type;
  
    d_到院时间     急诊就诊记录.到院时间%Type;
    n_是否三无人员 急诊就诊记录.是否三无人员%Type;
    n_是否复合伤   急诊就诊记录.是否复合伤%Type;
    n_是否绿色通道 急诊就诊记录.是否绿色通道%Type;
  
    n_是否成批就诊 急诊就诊记录.是否成批就诊%Type;
    n_成批就诊人数 急诊就诊记录.成批就诊人数%Type;
    v_病人来源     急诊就诊记录.病人来源%Type;
    v_陪同人员     急诊就诊记录.陪同人员%Type;
    v_意识状态     急诊就诊记录.意识状态%Type;
    v_既往病史     急诊就诊记录.既往病史%Type;
    v_主诉         急诊就诊记录.主诉%Type;
    n_病情级别     急诊就诊记录.病情级别%Type;
    v_登记人       急诊就诊记录.登记人%Type;
    d_登记时间     急诊就诊记录.登记时间%Type;
    v_备注         急诊就诊记录.备注%Type;
  
    --分诊记录
    n_分诊次数 急诊分诊记录.分诊次数%Type;
  
    n_分诊科室id   急诊分诊记录.分诊科室id%Type;
    v_分诊科室名称 急诊分诊记录.分诊科室名称%Type;
  
    d_体征测量时间 急诊分诊记录.体征测量时间%Type;
    n_舒张压       急诊分诊记录.舒张压%Type;
    n_收缩压       急诊分诊记录.收缩压%Type;
    n_血糖         急诊分诊记录.血糖%Type;
    n_指氧饱和度   急诊分诊记录.指氧饱和度%Type;
    n_心率         急诊分诊记录.心率%Type;
    n_血钾         急诊分诊记录.血钾%Type;
    n_体温         急诊分诊记录.体温%Type;
    n_呼吸频率     急诊分诊记录.呼吸频率%Type;
  
    n_自动病情级别  急诊分诊记录.自动病情级别%Type;
    n_人工病情级别  急诊分诊记录.人工病情级别%Type;
    v_人工评级说明  急诊分诊记录.人工评级说明%Type;
    v_修改说明      急诊分诊记录.修改说明%Type;
    v_站点          Varchar2(10);
    n_分诊科室idold 急诊分诊记录.分诊科室id%Type;
  
    d_Now Date;
  
    n_门诊号     Number(18);
    n_险类       Number(5);
    v_登记人编号 Varchar2(6);
    n_Count      Number(5);
  
    n_评分id       Number(18);
    n_方法id       Number(18);
    n_评分方法分值 Number(5);
    v_评分结果描述 Varchar2(100);
    n_评分等级     Number(1);
  
    Jsonlist评分指标 Pljson_List;
    Jsonlist病人评分 Pljson_List;
    Jsonlistitem     Pljson;
    Jsonlistitem指标 Pljson;
  
    n_Edittmp Number(5); --0  新增  1  修改
  Begin
    Jsonobj := Pljson(Input_In);
  
    n_Type           := Pljson_Ext.Get_String(Jsonobj, 'input.type');
    n_就诊id         := Pljson_Ext.Get_String(Jsonobj, 'input.就诊id');
    n_病人id         := Nvl(To_Number(Pljson_Ext.Get_String(Jsonobj, 'input.病人id')), 0);
    n_门诊号         := Nvl(To_Number(Pljson_Ext.Get_String(Jsonobj, 'input.门诊号')), 0);
    v_姓名           := Pljson_Ext.Get_String(Jsonobj, 'input.姓名');
    v_性别           := Pljson_Ext.Get_String(Jsonobj, 'input.性别');
    d_出生日期       := To_Date(Pljson_Ext.Get_String(Jsonobj, 'input.出生日期'), 'yyyy-mm-dd');
    v_身份证号       := Pljson_Ext.Get_String(Jsonobj, 'input.身份证号');
    v_联系电话       := Pljson_Ext.Get_String(Jsonobj, 'input.联系电话');
    v_民族           := Pljson_Ext.Get_String(Jsonobj, 'input.民族');
    v_医保卡号       := Pljson_Ext.Get_String(Jsonobj, 'input.医保卡号');
    v_保险类别       := Pljson_Ext.Get_String(Jsonobj, 'input.保险类别');
    v_家庭地址       := Pljson_Ext.Get_String(Jsonobj, 'input.家庭地址');
    v_病人年龄       := Pljson_Ext.Get_String(Jsonobj, 'input.病人年龄');
    n_年龄数值       := To_Number(Pljson_Ext.Get_String(Jsonobj, 'input.年龄数值'));
    v_年龄单位       := Pljson_Ext.Get_String(Jsonobj, 'input.年龄单位');
    d_到院时间       := To_Date(Pljson_Ext.Get_String(Jsonobj, 'input.到院时间'), 'yyyy-mm-dd hh24:mi:ss');
    n_是否三无人员   := To_Number(Pljson_Ext.Get_String(Jsonobj, 'input.是否三无人员'));
    n_是否复合伤     := To_Number(Pljson_Ext.Get_String(Jsonobj, 'input.是否复合伤'));
    n_是否绿色通道   := To_Number(Pljson_Ext.Get_String(Jsonobj, 'input.是否绿色通道'));
    n_是否成批就诊   := To_Number(Pljson_Ext.Get_String(Jsonobj, 'input.是否成批就诊'));
    n_成批就诊人数   := To_Number(Pljson_Ext.Get_String(Jsonobj, 'input.成批就诊人数'));
    v_病人来源       := Pljson_Ext.Get_String(Jsonobj, 'input.病人来源');
    v_陪同人员       := Pljson_Ext.Get_String(Jsonobj, 'input.陪同人员');
    v_意识状态       := Pljson_Ext.Get_String(Jsonobj, 'input.意识状态');
    v_既往病史       := Pljson_Ext.Get_String(Jsonobj, 'input.既往病史');
    v_主诉           := Pljson_Ext.Get_String(Jsonobj, 'input.主诉');
    n_病情级别       := To_Number(Pljson_Ext.Get_String(Jsonobj, 'input.病情级别'));
    v_登记人         := Pljson_Ext.Get_String(Jsonobj, 'input.登记人');
    v_备注           := Pljson_Ext.Get_String(Jsonobj, 'input.备注');
    n_分诊科室id     := To_Number(Pljson_Ext.Get_String(Jsonobj, 'input.分诊科室id'));
    v_分诊科室名称   := Pljson_Ext.Get_String(Jsonobj, 'input.分诊科室名称');
    d_体征测量时间   := To_Date(Pljson_Ext.Get_String(Jsonobj, 'input.体征测量时间'), 'yyyy-mm-dd hh24:mi:ss');
    n_舒张压         := To_Number(Pljson_Ext.Get_String(Jsonobj, 'input.舒张压'));
    n_收缩压         := To_Number(Pljson_Ext.Get_String(Jsonobj, 'input.收缩压'));
    n_血糖           := To_Number(Pljson_Ext.Get_String(Jsonobj, 'input.血糖'));
    n_指氧饱和度     := To_Number(Pljson_Ext.Get_String(Jsonobj, 'input.指氧饱和度'));
    n_心率           := To_Number(Pljson_Ext.Get_String(Jsonobj, 'input.心率'));
    n_血钾           := To_Number(Pljson_Ext.Get_String(Jsonobj, 'input.血钾'));
    n_体温           := To_Number(Pljson_Ext.Get_String(Jsonobj, 'input.体温'));
    n_呼吸频率       := To_Number(Pljson_Ext.Get_String(Jsonobj, 'input.呼吸频率'));
    n_自动病情级别   := To_Number(Pljson_Ext.Get_String(Jsonobj, 'input.自动病情级别'));
    n_人工病情级别   := To_Number(Pljson_Ext.Get_String(Jsonobj, 'input.人工病情级别'));
    v_人工评级说明   := Pljson_Ext.Get_String(Jsonobj, 'input.人工评级说明');
    v_修改说明       := Pljson_Ext.Get_String(Jsonobj, 'input.修改说明');
    v_登记人编号     := Pljson_Ext.Get_String(Jsonobj, 'input.登记人编号');
    v_站点           := Pljson_Ext.Get_String(Jsonobj, 'input.站点');
    Jsonlist评分指标 := Pljson_Ext.Get_Json_List(Jsonobj, 'input.评分指标');
    Jsonlist病人评分 := Pljson_Ext.Get_Json_List(Jsonobj, 'input.病人评分');
  
    n_Edittmp := 0;
  
    --获取登记人编号
    If v_登记人编号 Is Null Then
      Select Max(编号) Into v_登记人编号 From 人员表 Where 姓名 = v_登记人;
    End If;
    --获取保险类别
    If v_保险类别 Is Not Null Then
      Select Max(序号) Into n_险类 From 保险类别 Where 名称 = v_保险类别;
    End If;
  
    Select Sysdate Into d_Now From Dual;
    d_登记时间 := d_Now;
  
    --新增时重新产生
    If n_Type = 1 Then
      Select 急诊就诊记录_Id.Nextval Into n_就诊id From Dual;
    End If;
  
    --分诊ID都是重新产生
    Select 急诊分诊记录_Id.Nextval Into n_分诊id From Dual;
  
    --产生门诊号
    --等待处理身份信息
    If n_Type = 1 Then
      If n_病人id > 0 Then
        n_Edittmp := 1;
      Else
        If v_身份证号 Is Not Null Then
          n_Count := Nvl(zl_GetSysParameter(279), 0);
          If n_Count = 1 Then
            Select Max(病人id) Into n_病人id From 病人信息 Where 身份证号 = v_身份证号;
            If n_病人id > 0 Then
              n_Edittmp := 1;
            End If;
          End If;
        End If;
      End If;
    
      If n_Edittmp = 0 Then
        Select 病人信息_Id.Nextval Into n_病人id From Dual;
        n_门诊号 := Nextno(3);
        Zl_病人信息_Insert(n_病人id, n_门诊号, Null, Null, v_姓名, v_性别, v_病人年龄, d_出生日期, Null, v_身份证号, Null, Null, v_民族, Null,
                       Null, Null, v_家庭地址, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null,
                       Null, n_险类, Sysdate, Null, Null, v_登记人编号, v_登记人, v_医保卡号, Null, Null, Null, Null, Null, Null, Null,
                       v_联系电话);
      Else
        If n_病人id = 0 Then
          Select Max(门诊号) Into n_门诊号 From 病人信息 Where 病人id = n_病人id;
          If n_门诊号 Is Null Then
            n_门诊号 := Nextno(3);
          End If;
        End If;
        Update 病人信息
        Set 门诊号 = n_门诊号, 姓名 = Nvl(v_姓名, 姓名), 性别 = Nvl(v_性别, 性别), 年龄 = Nvl(v_病人年龄, 年龄), 出生日期 = Nvl(d_出生日期, 出生日期),
            身份证号 = Nvl(v_身份证号, 身份证号), 民族 = Nvl(v_民族, 民族), 家庭地址 = Nvl(v_家庭地址, 家庭地址), 险类 = Nvl(n_险类, 险类),
            医保号 = Nvl(v_医保卡号, 医保号), 手机号 = Nvl(v_联系电话, 手机号)
        Where 病人id = n_病人id;
      End If;
    Else
      If n_就诊id Is Not Null Then
        Select Max(Nvl(病人id, 0)), Max(Nvl(挂号id, 0)), Max(Nvl(分诊科室id, 0))
        Into n_病人id, n_挂号id, n_分诊科室idold
        From 急诊就诊记录
        Where ID = n_就诊id;
      
        --修改不处理病人信息
        /*Select Max(门诊号) Into n_门诊号 From 病人信息 Where 病人id = n_病人id;
        Update 病人信息
        Set 门诊号 = n_门诊号, 姓名 = Nvl(v_姓名, 姓名), 性别 = Nvl(v_性别, 性别), 年龄 = Nvl(v_病人年龄, 年龄), 出生日期 = Nvl(d_出生日期, 出生日期),
            身份证号 = Nvl(v_身份证号, 身份证号), 民族 = Nvl(v_民族, 民族), 家庭地址 = Nvl(v_家庭地址, 家庭地址), 险类 = Nvl(n_险类, 险类),
            医保号 = Nvl(v_医保卡号, 医保号), 手机号 = Nvl(v_联系电话, 手机号)
        Where 病人id = n_病人id;*/
      End If;
    End If;
  
    If n_Type = 1 Then
      --处理挂号id
    
      n_挂号id := Zl_Emergencyregist(n_病人id, n_分诊科室id, v_站点, n_是否绿色通道);
    
      Insert Into 急诊就诊记录
        (ID, 病人id, 病人年龄, 年龄数值, 年龄单位, 挂号id, 病情级别, 到院时间, 主诉, 是否三无人员, 陪同人员, 病人来源, 既往病史, 意识状态, 是否成批就诊, 成批就诊人数, 是否复合伤, 备注,
         登记人, 登记时间, 分诊病情级别, 是否绿色通道, 分诊科室id)
      Values
        (n_就诊id, n_病人id, v_病人年龄, n_年龄数值, v_年龄单位, n_挂号id, n_病情级别, d_到院时间, v_主诉, n_是否三无人员, v_陪同人员, v_病人来源, v_既往病史, v_意识状态,
         n_是否成批就诊, n_成批就诊人数, n_是否复合伤, v_备注, v_登记人, d_登记时间, n_病情级别, n_是否绿色通道, n_分诊科室id);
    Else
      If n_分诊科室idold <> n_分诊科室id Then
        Zl_Emergencyregistredo(n_挂号id, n_分诊科室id, v_站点);
      End If;
      Update 急诊就诊记录
      Set 病人年龄 = v_病人年龄, 年龄数值 = n_年龄数值, 年龄单位 = v_年龄单位, 挂号id = n_挂号id, 病情级别 = n_病情级别, 到院时间 = d_到院时间, 主诉 = v_主诉,
          是否三无人员 = n_是否三无人员, 陪同人员 = v_陪同人员, 病人来源 = v_病人来源, 既往病史 = v_既往病史, 意识状态 = v_意识状态, 是否成批就诊 = n_是否成批就诊,
          成批就诊人数 = n_成批就诊人数, 是否复合伤 = n_是否复合伤, 备注 = v_备注, 分诊病情级别 = n_病情级别, 是否绿色通道 = n_是否绿色通道, 登记时间 = d_登记时间,
          分诊科室id = n_分诊科室id
      Where ID = n_就诊id;
    End If;
  
    If n_Type = 1 Then
      n_分诊次数 := 1;
    Else
      Select Max(分诊次数) + 1 Into n_分诊次数 From 急诊分诊记录 Where 就诊id = n_就诊id;
    End If;
  
    Insert Into 急诊分诊记录
      (ID, 就诊id, 分诊次数, 自动病情级别, 分诊科室id, 分诊科室名称, 收缩压, 舒张压, 心率, 指氧饱和度, 体温, 血糖, 血钾, 体征测量时间, 登记人, 登记时间, 人工病情级别, 人工评级说明, 呼吸频率,
       修改说明)
    Values
      (n_分诊id, n_就诊id, n_分诊次数, n_自动病情级别, n_分诊科室id, v_分诊科室名称, n_收缩压, n_舒张压, n_心率, n_指氧饱和度, n_体温, n_血糖, n_血钾, d_体征测量时间,
       v_登记人, d_登记时间, n_人工病情级别, v_人工评级说明, n_呼吸频率, v_修改说明);
  
    Delete From 病人信息从表
    Where 病人id = n_病人id And 就诊id = n_挂号id And 信息名 In ('体温', '呼吸', '脉搏', '收缩压', '舒张压', '血糖');
  
    If n_体温 Is Not Null Then
      Insert Into 病人信息从表
        (病人id, 就诊id, 信息名, 信息值)
        Select n_病人id, n_挂号id, '体温', To_Char(n_体温) From Dual;
    End If;
  
    If n_呼吸频率 Is Not Null Then
      Insert Into 病人信息从表
        (病人id, 就诊id, 信息名, 信息值)
        Select n_病人id, n_挂号id, '呼吸', To_Char(n_呼吸频率) From Dual;
    End If;
  
    If n_心率 Is Not Null Then
      Insert Into 病人信息从表
        (病人id, 就诊id, 信息名, 信息值)
        Select n_病人id, n_挂号id, '脉搏', To_Char(n_心率) From Dual;
    End If;
  
    If n_收缩压 Is Not Null Then
      Insert Into 病人信息从表
        (病人id, 就诊id, 信息名, 信息值)
        Select n_病人id, n_挂号id, '收缩压', To_Char(n_收缩压) From Dual;
    End If;
  
    If n_舒张压 Is Not Null Then
      Insert Into 病人信息从表
        (病人id, 就诊id, 信息名, 信息值)
        Select n_病人id, n_挂号id, '舒张压', To_Char(n_舒张压) From Dual;
    End If;
  
    If n_血糖 Is Not Null Then
      Insert Into 病人信息从表
        (病人id, 就诊id, 信息名, 信息值)
        Select n_病人id, n_挂号id, '血糖', To_Char(n_血糖) From Dual;
    End If;
  
    For I In 1 .. Jsonlist病人评分.Count Loop
      Jsonlistitem   := Pljson(Jsonlist病人评分.Get(I));
      n_方法id       := To_Number(Pljson_Ext.Get_String(Jsonlistitem, '方法ID'));
      n_评分方法分值 := To_Number(Pljson_Ext.Get_String(Jsonlistitem, '评分方法分值'));
      v_评分结果描述 := Pljson_Ext.Get_String(Jsonlistitem, '评分结果描述');
      n_评分等级     := To_Number(Pljson_Ext.Get_String(Jsonlistitem, '评分等级'));
      Select 急诊病人评分_Id.Nextval Into n_评分id From Dual;
    
      Insert Into 急诊病人评分
        (ID, 分诊id, 方法id, 评分方法分值, 评分结果描述, 病情级别)
      Values
        (n_评分id, n_分诊id, n_方法id, n_评分方法分值, v_评分结果描述, n_评分等级);
    
      For I In 1 .. Jsonlist评分指标.Count Loop
        Jsonlistitem指标 := Pljson(Jsonlist评分指标.Get(I));
        If n_方法id = To_Number(Pljson_Ext.Get_String(Jsonlistitem指标, '方法ID')) Then
          Insert Into 急诊病人评分指标
            (评分id, 指标id, 指标结果文本)
          Values
            (n_评分id, To_Number(Pljson_Ext.Get_String(Jsonlistitem指标, '指标ID')),
             Pljson_Ext.Get_String(Jsonlistitem指标, '指标结果文本'));
        End If;
      End Loop;
    End Loop;
  
    Open Output_Out For
      Select n_病人id As 病人id, n_就诊id As 就诊id, n_分诊id As 分诊id From Dual;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Save_Pretriage;

End Pkg_Pretriage_Dml;
/