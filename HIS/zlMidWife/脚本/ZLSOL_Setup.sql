--以zlsol登录执行以下脚本；conn zlsol/his@ORA_SOL
--ZLHIS_DBL的dblink连接；新生儿同步到HIS需要
create database link ZLHIS_DBL connect to ZLHIS identified by &zlhis连接HIS库的密码 using '(DESCRIPTION =
    (ADDRESS = (PROTOCOL = TCP)(HOST = &HIS库IP)(PORT = &HIS库端口))
    (CONNECT_DATA =
      (SERVICE_NAME = &HIS库实例名)
    )
  )';
create table ZLSOL.SOL_INF_PUERPERA
(
  mid          NUMBER(18) generated always as identity,
  pid          NUMBER(18),
  tid          NUMBER(5),
  name         VARCHAR2(50),
  old          VARCHAR2(20),
  bedno        VARCHAR2(10),
  pno          NUMBER(18),
  diagnosis    VARCHAR2(100),
  status       NUMBER(1),
  outtime      DATE,
  outroomtime  DATE,
  expectant    NUMBER(1),
  checkinroom  NUMBER(1),
  birth        NUMBER(1),
  druglabor    NUMBER(1),
  delivery     NUMBER(1),
  newborns     NUMBER(1),
  postpartum   NUMBER(1),
  checkoutroom NUMBER(1),
  equipment    NUMBER(1)
)
tablespace ZLSOL_DATA
  pctfree 10
  initrans 1
  maxtrans 255
  storage
  (
    initial 64K
    next 1M
    minextents 1
    maxextents unlimited
  );
comment on table ZLSOL.SOL_INF_PUERPERA
  is '产妇信息';
create index ZLSOL.SOL_INF_PUERPERA_IX_OUTROOMTIME on ZLSOL.SOL_INF_PUERPERA (OUTROOMTIME)
  tablespace ZLSOL_DATA
  pctfree 10
  initrans 2
  maxtrans 255
  storage
  (
    initial 64K
    next 1M
    minextents 1
    maxextents unlimited
  );
create index ZLSOL.SOL_INF_PUERPERA_IX_OUTTIME on ZLSOL.SOL_INF_PUERPERA (OUTTIME)
  tablespace ZLSOL_DATA
  pctfree 10
  initrans 2
  maxtrans 255
  storage
  (
    initial 64K
    next 1M
    minextents 1
    maxextents unlimited
  );
create index ZLSOL.SOL_INF_PUERPERA_IX_STATUS on ZLSOL.SOL_INF_PUERPERA (STATUS)
  tablespace ZLSOL_DATA
  pctfree 10
  initrans 2
  maxtrans 255
  storage
  (
    initial 64K
    next 1M
    minextents 1
    maxextents unlimited
  );
alter table ZLSOL.SOL_INF_PUERPERA
  add constraint SOL_INF_PUERPERA_PK primary key (MID)
  using index 
  tablespace ZLSOL_DATA
  pctfree 10
  initrans 2
  maxtrans 255
  storage
  (
    initial 64K
    next 1M
    minextents 1
    maxextents unlimited
  );

create table ZLSOL.SOL_INF_CHECKINROOM
(
  mid      NUMBER(18) not null,
  content  CLOB,
  recorder VARCHAR2(50),
  addtime  DATE
)
tablespace ZLSOL_DATA
  pctfree 10
  initrans 1
  maxtrans 255
  storage
  (
    initial 64K
    next 1M
    minextents 1
    maxextents unlimited
  );
comment on table ZLSOL.SOL_INF_CHECKINROOM
  is '系统用户信息';
alter table ZLSOL.SOL_INF_CHECKINROOM
  add constraint SOL_INF_CHECKINROOM_PK primary key (MID)
  using index 
  tablespace ZLSOL_DATA
  pctfree 10
  initrans 2
  maxtrans 255
  storage
  (
    initial 64K
    next 1M
    minextents 1
    maxextents unlimited
  );
alter table ZLSOL.SOL_INF_CHECKINROOM
  add constraint SOL_INF_CHECKINROOM_FK_MID foreign key (MID)
  references ZLSOL.SOL_INF_PUERPERA (MID);
alter table ZLSOL.SOL_INF_CHECKINROOM
  add check ("CONTENT" IS JSON (LAX));

create table ZLSOL.SOL_INF_CHECKOUTROOM
(
  mid      NUMBER(18) not null,
  content  CLOB,
  recorder VARCHAR2(50),
  addtime  DATE
)
tablespace ZLSOL_DATA
  pctfree 10
  initrans 1
  maxtrans 255
  storage
  (
    initial 64K
    next 1M
    minextents 1
    maxextents unlimited
  );
comment on table ZLSOL.SOL_INF_CHECKOUTROOM
  is '出房信息';
alter table ZLSOL.SOL_INF_CHECKOUTROOM
  add constraint SOL_INF_CHECKOUTROOM_PK primary key (MID)
  using index 
  tablespace ZLSOL_DATA
  pctfree 10
  initrans 2
  maxtrans 255
  storage
  (
    initial 64K
    next 1M
    minextents 1
    maxextents unlimited
  );
alter table ZLSOL.SOL_INF_CHECKOUTROOM
  add constraint SOL_INF_CHECKOUTROOM_FK_MID foreign key (MID)
  references ZLSOL.SOL_INF_PUERPERA (MID);
alter table ZLSOL.SOL_INF_CHECKOUTROOM
  add check ("CONTENT" IS JSON (LAX));

create table ZLSOL.SOL_INF_DELIVERY
(
  mid           NUMBER(18) not null,
  deliveryinf   CLOB,
  newborndetail CLOB,
  newbornscore  CLOB,
  otherinf      CLOB,
  delivertime   DATE
)
tablespace ZLSOL_DATA
  pctfree 10
  initrans 1
  maxtrans 255
  storage
  (
    initial 64K
    next 1M
    minextents 1
    maxextents unlimited
  );
comment on table ZLSOL.SOL_INF_DELIVERY
  is '分娩信息';
alter table ZLSOL.SOL_INF_DELIVERY
  add constraint SOL_INF_DELIVERY_PK primary key (MID)
  using index 
  tablespace ZLSOL_DATA
  pctfree 10
  initrans 2
  maxtrans 255
  storage
  (
    initial 64K
    next 1M
    minextents 1
    maxextents unlimited
  );
alter table ZLSOL.SOL_INF_DELIVERY
  add constraint SOL_INF_DELIVERY_FK_MID foreign key (MID)
  references ZLSOL.SOL_INF_PUERPERA (MID);
alter table ZLSOL.SOL_INF_DELIVERY
  add check (DELIVERYINF IS JSON);
alter table ZLSOL.SOL_INF_DELIVERY
  add check (NEWBORNDETAIL IS JSON);
alter table ZLSOL.SOL_INF_DELIVERY
  add check (NEWBORNSCORE IS JSON);
alter table ZLSOL.SOL_INF_DELIVERY
  add check (OTHERINF IS JSON);

create table ZLSOL.SOL_INF_EQUIPMENT
(
  mid       NUMBER(18),
  content   CLOB,
  befortime DATE,
  deliver   VARCHAR2(50),
  inspector VARCHAR2(50),
  aftertime DATE
)
tablespace ZLSOL_DATA
  pctfree 10
  initrans 1
  maxtrans 255
  storage
  (
    initial 64K
    next 1M
    minextents 1
    maxextents unlimited
  );
comment on table ZLSOL.SOL_INF_EQUIPMENT
  is '器械清点记录';

create table ZLSOL.SOL_INF_NEWBORNS
(
  bid          NUMBER(18) generated always as identity,
  mid          NUMBER(18),
  babyno       NUMBER(5),
  sex          VARCHAR2(10),
  newborninf   CLOB,
  newbornscore CLOB,
  otherinf     CLOB,
  recorder     VARCHAR2(50),
  addtime      DATE
)
tablespace ZLSOL_DATA
  pctfree 10
  initrans 1
  maxtrans 255
  storage
  (
    initial 64K
    next 1M
    minextents 1
    maxextents unlimited
  );
comment on table ZLSOL.SOL_INF_NEWBORNS
  is '新生儿信息';
alter table ZLSOL.SOL_INF_NEWBORNS
  add constraint SOL_INF_NEWBORNS_PK primary key (BID)
  using index 
  tablespace ZLSOL_DATA
  pctfree 10
  initrans 2
  maxtrans 255
  storage
  (
    initial 64K
    next 1M
    minextents 1
    maxextents unlimited
  );
alter table ZLSOL.SOL_INF_NEWBORNS
  add constraint SOL_INF_NEWBORNS_UQ unique (MID, BABYNO)
  using index 
  tablespace ZLSOL_DATA
  pctfree 10
  initrans 2
  maxtrans 255
  storage
  (
    initial 64K
    next 1M
    minextents 1
    maxextents unlimited
  );
alter table ZLSOL.SOL_INF_NEWBORNS
  add check (NEWBORNINF IS JSON);
alter table ZLSOL.SOL_INF_NEWBORNS
  add check (NEWBORNSCORE IS JSON);
alter table ZLSOL.SOL_INF_NEWBORNS
  add check (OTHERINF IS JSON);

create table ZLSOL.SOL_RS_BIRTH
(
  mid     NUMBER(18) not null,
  content CLOB
)
tablespace ZLSOL_DATA
  pctfree 10
  initrans 1
  maxtrans 255
  storage
  (
    initial 64K
    next 1M
    minextents 1
    maxextents unlimited
  );
comment on table ZLSOL.SOL_RS_BIRTH
  is '产前检查信息';
alter table ZLSOL.SOL_RS_BIRTH
  add constraint SOL_RS_BIRTH_PK primary key (MID)
  using index 
  tablespace ZLSOL_DATA
  pctfree 10
  initrans 2
  maxtrans 255
  storage
  (
    initial 64K
    next 1M
    minextents 1
    maxextents unlimited
  );
alter table ZLSOL.SOL_RS_BIRTH
  add constraint SOL_RS_BIRTH_FK_MID foreign key (MID)
  references ZLSOL.SOL_INF_PUERPERA (MID);
alter table ZLSOL.SOL_RS_BIRTH
  add check ("CONTENT" IS JSON (LAX));

create table ZLSOL.SOL_RS_BIRTH_COURSE
(
  courseid NUMBER(18) generated always as identity,
  mid      NUMBER(18),
  content  CLOB
)
tablespace ZLSOL_DATA
  pctfree 10
  initrans 1
  maxtrans 255
  storage
  (
    initial 64K
    next 1M
    minextents 1
    maxextents unlimited
  );
comment on table ZLSOL.SOL_RS_BIRTH_COURSE
  is '产程经过';
alter table ZLSOL.SOL_RS_BIRTH_COURSE
  add constraint SOL_RS_BIRTH_COURSE_PK primary key (COURSEID)
  using index 
  tablespace ZLSOL_DATA
  pctfree 10
  initrans 2
  maxtrans 255
  storage
  (
    initial 64K
    next 1M
    minextents 1
    maxextents unlimited
  );
alter table ZLSOL.SOL_RS_BIRTH_COURSE
  add constraint SOL_RS_BIRTH_COURSE_FK_MID foreign key (MID)
  references ZLSOL.SOL_INF_PUERPERA (MID);
alter table ZLSOL.SOL_RS_BIRTH_COURSE
  add check (CONTENT IS JSON);

create table ZLSOL.SOL_RS_DRUGLABOR
(
  mid  NUMBER(18) not null,
  日期   DATE,
  引产指征 VARCHAR2(100),
  引产方法 VARCHAR2(100),
  检查者 VARCHAR2(100)
)
tablespace ZLSOL_DATA
  pctfree 10
  initrans 1
  maxtrans 255
  storage
  (
    initial 64K
    next 1M
    minextents 1
    maxextents unlimited
  );
comment on table ZLSOL.SOL_RS_DRUGLABOR
  is '药物引产信息';
alter table ZLSOL.SOL_RS_DRUGLABOR
  add constraint SOL_RS_DRUGLABOR_PK primary key (MID)
  using index 
  tablespace ZLSOL_DATA
  pctfree 10
  initrans 2
  maxtrans 255
  storage
  (
    initial 64K
    next 1M
    minextents 1
    maxextents unlimited
  );
alter table ZLSOL.SOL_RS_DRUGLABOR
  add constraint SOL_RS_DRUGLABOR_FK_MID foreign key (MID)
  references ZLSOL.SOL_INF_PUERPERA (MID);

create table ZLSOL.SOL_RS_DRUGLABOR_LIST
(
  courseid NUMBER(18) generated always as identity,
  mid      NUMBER(18),
  content  CLOB
)
tablespace ZLSOL_DATA
  pctfree 10
  initrans 1
  maxtrans 255
  storage
  (
    initial 64K
    next 1M
    minextents 1
    maxextents unlimited
  );
comment on table ZLSOL.SOL_RS_DRUGLABOR_LIST
  is '药物引产记录';
alter table ZLSOL.SOL_RS_DRUGLABOR_LIST
  add constraint SOL_RS_DRUGLABOR_LIST_PK primary key (COURSEID)
  using index 
  tablespace ZLSOL_DATA
  pctfree 10
  initrans 2
  maxtrans 255
  storage
  (
    initial 64K
    next 1M
    minextents 1
    maxextents unlimited
  );
alter table ZLSOL.SOL_RS_DRUGLABOR_LIST
  add constraint SOL_RS_DRUGLABOR_LIST_FK_MID foreign key (MID)
  references ZLSOL.SOL_INF_PUERPERA (MID);
alter table ZLSOL.SOL_RS_DRUGLABOR_LIST
  add check (CONTENT IS JSON);

create table ZLSOL.SOL_RS_EXPECTANT
(
  courseid NUMBER(18) generated always as identity,
  mid      NUMBER(18),
  content  CLOB
)
tablespace ZLSOL_DATA
  pctfree 10
  initrans 1
  maxtrans 255
  storage
  (
    initial 64K
    next 1M
    minextents 1
    maxextents unlimited
  );
comment on table ZLSOL.SOL_RS_EXPECTANT
  is '待产记录';
alter table ZLSOL.SOL_RS_EXPECTANT
  add constraint SOL_RS_EXPECTANT_PK primary key (COURSEID)
  using index 
  tablespace ZLSOL_DATA
  pctfree 10
  initrans 2
  maxtrans 255
  storage
  (
    initial 64K
    next 1M
    minextents 1
    maxextents unlimited
  );
alter table ZLSOL.SOL_RS_EXPECTANT
  add constraint SOL_RS_EXPECTANT_FK_MID foreign key (MID)
  references ZLSOL.SOL_INF_PUERPERA (MID);
alter table ZLSOL.SOL_RS_EXPECTANT
  add check (CONTENT IS JSON);

create table ZLSOL.SOL_RS_POSTPARTUM
(
  mid     NUMBER(18) not null,
  content CLOB
)
tablespace ZLSOL_DATA
  pctfree 10
  initrans 1
  maxtrans 255
  storage
  (
    initial 64K
    next 1M
    minextents 1
    maxextents unlimited
  );
comment on table ZLSOL.SOL_RS_POSTPARTUM
  is '产后观察信息';
alter table ZLSOL.SOL_RS_POSTPARTUM
  add constraint SOL_RS_POSTPARTUM_PK primary key (MID)
  using index 
  tablespace ZLSOL_DATA
  pctfree 10
  initrans 2
  maxtrans 255
  storage
  (
    initial 64K
    next 1M
    minextents 1
    maxextents unlimited
  );
alter table ZLSOL.SOL_RS_POSTPARTUM
  add constraint SOL_RS_POSTPARTUM_FK_MID foreign key (MID)
  references ZLSOL.SOL_INF_PUERPERA (MID);
alter table ZLSOL.SOL_RS_POSTPARTUM
  add check (CONTENT IS JSON);

create table ZLSOL.SOL_RS_POSTPARTUM_LIST
(
  courseid NUMBER(18) generated always as identity,
  mid      NUMBER(18),
  content  CLOB
)
tablespace ZLSOL_DATA
  pctfree 10
  initrans 1
  maxtrans 255
  storage
  (
    initial 64K
    next 1M
    minextents 1
    maxextents unlimited
  );
comment on table ZLSOL.SOL_RS_POSTPARTUM_LIST
  is '产后观察记录';
alter table ZLSOL.SOL_RS_POSTPARTUM_LIST
  add constraint SOL_RS_POSTPARTUM_LIST_PK primary key (COURSEID)
  using index 
  tablespace ZLSOL_DATA
  pctfree 10
  initrans 2
  maxtrans 255
  storage
  (
    initial 64K
    next 1M
    minextents 1
    maxextents unlimited
  );
alter table ZLSOL.SOL_RS_POSTPARTUM_LIST
  add constraint SOL_RS_POSTPARTUM_LIST_FK_MID foreign key (MID)
  references ZLSOL.SOL_INF_PUERPERA (MID);
alter table ZLSOL.SOL_RS_POSTPARTUM_LIST
  add check (CONTENT IS JSON);

create or replace force view ZLSOL.newborn as
Select a.mid, b.SEX,b.宫内窘迫,b.初生时处理舒救方法,b.脐带处理,b.评分1分钟,b.评分5分钟,b.评分10分钟,b.眼睛滴药,b.一般情况,b.皮肤,b.胎脂,
b.体重,b.身长,b.坐高,b.头部产伤,b.变形,b.水肿,b.血肿,b.五官,b.唇,b.口腔,b.胸部,b.心,b.肺,b.脐出血,
b.肝,b.脾,b.包块,b.四肢,b.指,b.趾,b.生殖器,b.肛门
From Sol_Inf_Newborns a,JSON_TABLE(a.newborninf,'$' columns(
SEX   Varchar2(50) PATH '$.SEX',
宫内窘迫            Varchar2(50) PATH '$.宫内窘迫',
初生时处理舒救方法          Varchar2(50) PATH '$.初生时处理舒救方法',
脐带处理      Varchar2(50) PATH '$.脐带处理',
评分1分钟 Number(2) PATH '$.评分1分钟',
评分5分钟 Number(2) PATH '$.评分5分钟',
评分10分钟 Number(2) PATH '$.评分10分钟',
眼睛滴药      Varchar2(50) PATH '$.眼睛滴药',
一般情况       Varchar2(50) PATH '$.一般情况',
皮肤       Varchar2(50) PATH '$.皮肤',
胎脂       Varchar2(10) PATH '$.胎脂',
体重       Varchar2(10) PATH '$.体重',
身长      Varchar2(10) PATH '$.身长',
坐高   Varchar2(20) PATH '$.坐高',
头部产伤        Varchar2(10) PATH '$.头部产伤',
变形       Varchar2(10) PATH '$.变形',
水肿       Varchar2(10) PATH '$.水肿',
血肿       Varchar2(10) PATH '$.血肿',
五官         Varchar2(50) PATH '$.五官',
唇        Varchar2(10) PATH '$.唇',
口腔       Varchar2(10) PATH '$.口腔',
胸部       Varchar2(50) PATH '$.胸部',
心       Varchar2(50) PATH '$.心',
肺         Varchar2(50) PATH '$.肺',
脐出血        Varchar2(50) PATH '$.脐出血',
肝               Varchar2(50) PATH '$.肝',
脾               Varchar2(50) PATH '$.脾',
包块              Varchar2(50) PATH '$.包块',
四肢              Varchar2(50) PATH '$.四肢',
指               Varchar2(50) PATH '$.指',
趾               Varchar2(50) PATH '$.趾',
生殖器             Varchar2(50) PATH '$.生殖器',
肛门              Varchar2(50) PATH '$.肛门'
)) as b;

create or replace force view ZLSOL.sol_userlist as
Select User_Name User_Code, Last_Name || First_Name  User_Name
  From Apex_200100.Wwv_Flow_Fnd_User
  Where First_Name Is Not Null And Last_Name Is Not Null;

create or replace force view ZLSOL.v_delivery as
Select a.Mid, b.隐藏1,b.产程开始时间,b.宫口全开时间,b.胎儿娩出时间,b.胎盘娩出时间,b.第一产程,b.第二产程,b.第三产程,b.宫缩情况,b.出产房宫高脐下距离,b.结扎,b.破膜方式,b.破膜时间,b.羊水性状,b.羊水量,b.羊水颜色,b.胎盘娩出方式,b.胎盘剥离方式,b.胎盘完整度,b.胎盘胎膜残留,b.胎盘体积,b.胎盘形态,b.胎盘大小,b.胎盘重量,b.脐带附着,b.脐带长度,b.脐带绕颈,b.脐带真假结,b.脐带脱垂,b.娩出方式,b.娩出胎方位,b.产瘤大小,b.产瘤部位,b.会阴裂伤程度,b.会阴裂伤切口,b.会阴裂伤缝合,b.会阴裂伤麻醉,b.宫颈裂伤长度,b.宫颈裂伤部位,b.宫颈裂伤状况,b.阴道裂伤部位大小,b.阴道裂伤血肿大小,b.新生儿性别,b.新生儿体重,b.新生儿身长,b.新生儿抢救吸氧,b.新生儿抢救吸出物,b.新生儿抢救吸出物性状,b.新生儿抢救抢救药物,b.新生儿抢救畸形,b.新生儿抢救死胎,b.新生儿抢救死产,b.产后血压,b.产后流血,b.产时用药,b.产后用药,b.特殊情况, d.隐藏3,d.出产房时间,d.护送人,d.接生人,d.记录人
From Sol_Inf_Delivery A,
     Json_Table(Nvl(a.Newborndetail, '{"隐藏1":"1"}'),'$' Columns(隐藏1 Varchar2(50) Path '$.隐藏1',
                  产程开始时间 Varchar2(19) Path '$.产程开始时间',宫口全开时间 Varchar2(19) Path '$.宫口全开时间',
                          胎儿娩出时间 Varchar2(19) Path '$.胎儿娩出时间',胎盘娩出时间 Varchar2(19) Path '$.胎盘娩出时间',
                          第一产程 Varchar2(50) Path '$.第一产程', 第二产程 Varchar2(50) Path '$.第二产程',第三产程 Varchar2(50) Path '$.第三产程',
                          宫缩情况 Varchar2(50) Path '$.宫缩情况', 出产房宫高脐下距离 Varchar2(50) Path '$.出产房宫高脐下距离', 结扎 Varchar2(50) Path '$.结扎',
                          破膜方式 Varchar2(50) Path '$.破膜方式', 破膜时间 Varchar2(19) Path '$.破膜时间',
                          羊水性状 Varchar2(50) Path '$.羊水性状', 羊水量 Varchar2(50) Path '$.羊水量',羊水颜色 Varchar2(50) Path '$.羊水颜色',
                          胎盘娩出方式 Varchar2(50) Path '$.胎盘娩出方式',胎盘剥离方式 Varchar2(50) Path '$.胎盘剥离方式',
                          胎盘完整度 Varchar2(50) Path '$.胎盘完整度', 胎盘胎膜残留 Varchar2(50) Path '$.胎盘胎膜残留',
                          胎盘体积 Varchar2(50) Path '$.胎盘体积', 胎盘形态 Varchar2(50) Path '$.胎盘形态',
                          胎盘大小 Varchar2(50) Path '$.胎盘大小',胎盘重量 Varchar2(50) Path '$.胎盘重量',
                          脐带附着 Varchar2(50) Path '$.脐带附着', 脐带长度 Varchar2(50) Path '$.脐带长度',
                          脐带绕颈 Varchar2(50) Path '$.脐带绕颈', 脐带真假结 Varchar2(50) Path '$.脐带真假结', 脐带脱垂 Varchar2(50) Path '$.脐带脱垂',
                          娩出方式 Varchar2(50) Path '$.娩出方式',娩出胎方位 Varchar2(50) Path '$.娩出胎方位',
                          产瘤大小 Varchar2(50) Path '$.产瘤大小',产瘤部位 Varchar2(50) Path '$.产瘤部位',
                          会阴裂伤程度 Varchar2(50) Path '$.会阴裂伤程度',
                          会阴裂伤切口 Varchar2(50) Path '$.会阴裂伤切口', 会阴裂伤缝合 Varchar2(50) Path '$.会阴裂伤缝合',
                          会阴裂伤麻醉 Varchar2(50) Path '$.会阴裂伤麻醉', 宫颈裂伤长度 Varchar2(50) Path '$.宫颈裂伤长度',
                          宫颈裂伤部位 Varchar2(50) Path '$.宫颈裂伤部位', 宫颈裂伤状况 Varchar2(50) Path '$.宫颈裂伤状况',
                          阴道裂伤部位大小 Varchar2(50) Path '$.阴道裂伤部位大小', 阴道裂伤血肿大小 Varchar2(50) Path '$.阴道裂伤血肿大小',
                          新生儿性别 Varchar2(50) Path '$.新生儿性别', 新生儿体重 Varchar2(50) Path '$.新生儿体重',
                          新生儿身长 Varchar2(50) Path '$.新生儿身长', 新生儿抢救吸氧 Varchar2(50) Path '$.新生儿抢救吸氧',
                          新生儿抢救吸出物 Varchar2(50) Path '$.新生儿抢救吸出物', 新生儿抢救吸出物性状 Varchar2(50) Path '$.新生儿抢救吸出物性状',
                          新生儿抢救抢救药物 Varchar2(50) Path '$.新生儿抢救抢救药物', 新生儿抢救畸形 Varchar2(50) Path '$.新生儿抢救畸形',
                          新生儿抢救死胎 Varchar2(50) Path '$.新生儿抢救死胎', 新生儿抢救死产 Varchar2(50) Path '$.新生儿抢救死产',
                          产后血压 Varchar2(50) Path '$.产后血压', 产后流血 Varchar2(50) Path '$.产后流血', 产时用药 Varchar2(50) Path '$.产时用药',
                          产后用药 Varchar2(50) Path '$.产后用药', 特殊情况 Varchar2(50) Path '$.特殊情况')) As B,
      Json_Table(Nvl(a.Deliveryinf, '{"隐藏3":"1"}'),
                 '$' Columns(隐藏3 Varchar2(1) Path '$.隐藏3', 出产房时间 Varchar2(50) Path '$.出产房时间',
                          护送人 Varchar2(50) Path '$.护送人', 接生人 Varchar2(50) Path '$.接生人', 记录人 Varchar2(50) Path '$.记录人')) As D;

create or replace force view ZLSOL.v_newborn as
Select a.Mid, b.Bid, b.Babyno, b.Addtime, b.Recorder, a.Name, a.Old, a.Bedno, a.Pno, b.Sex, d.隐藏2,d.BOUTT,d.身长,d.体重,d.血型,d.胎儿状况,d.头围,d.胸围,
d.一般情况反应,d.一般情况面色,d.一般情况皮肤,d.一般情况毳毛,d.头部变形,d.颅骨重叠,d.胎头水肿血肿,d.胎头水肿大小,d.前囟,d.张力,d.眼神,d.口腔,
d.心,d.乳结,d.肝,d.脾,d.四肢,d.外展试验,d.肛门,d.生殖器,d.咽喉部吸出物量,d.咽喉部吸出物性状,d.气管插管吸出物量,d.气管插管吸出物性状,
d.新生儿抢救状况,d.抢救药物,d.早吸吮,d.皮肤接触,d.吸氧方式,
e.隐藏3,e.心率1分钟,e.心率5分钟,e.心率10分钟,e.呼吸1分钟,e.呼吸5分钟,
e.呼吸10分钟,e.喉反射1分钟,e.喉反射5分钟,e.喉反射10分钟,e.肌张力1分钟,e.肌张力5分钟,e.肌张力10分钟,e.肤色1分钟,e.肤色5分钟,e.肤色10分钟,
e.总分1分钟,e.总分5分钟,e.总分10分钟, f.隐藏4,f.出孕期产时合并症及用药情况,f.出生前胎儿情况,f.婴儿出生时抢救情况,f.出生缺陷,
f.诊断,f.死亡时间
From Sol_Inf_Puerpera A, Sol_Inf_Newborns B,
     Json_Table(Nvl(b.Newborninf, '{"隐藏2":"2"}'),
                 '$' Columns(隐藏2 Varchar2(50) Path '$.隐藏2',BOUTT Varchar2(50) Path '$.BOUTT', 身长 Varchar2(50) Path '$.身长', 体重 Varchar2(50) Path '$.体重',
                          头围 Varchar2(50) Path '$.头围', 胸围 Varchar2(50) Path '$.胸围', 一般情况反应 Varchar2(50) Path '$.一般情况反应',
                          血型 Varchar2(50) Path '$.血型', 胎儿状况 Varchar2(50) Path '$.胎儿状况',
                          一般情况面色 Varchar2(50) Path '$.一般情况面色', 一般情况皮肤 Varchar2(50) Path '$.一般情况皮肤',
                          一般情况毳毛 Varchar2(50) Path '$.一般情况毳毛', 头部变形 Varchar2(50) Path '$.头部变形',
                          颅骨重叠 Varchar2(50) Path '$.颅骨重叠', 胎头水肿血肿 Varchar2(50) Path '$.胎头水肿血肿',
                          胎头水肿大小 Varchar2(50) Path '$.胎头水肿大小', 前囟 Varchar2(50) Path '$.前囟', 张力 Varchar2(50) Path '$.张力',
                          眼神 Varchar2(50) Path '$.眼神', 口腔 Varchar2(50) Path '$.口腔', 心 Varchar2(50) Path '$.心',
                          乳结 Varchar2(50) Path '$.乳结', 肝 Varchar2(50) Path '$.肝', 脾 Varchar2(50) Path '$.脾',
                          四肢 Varchar2(50) Path '$.四肢', 外展试验 Varchar2(50) Path '$.外展试验', 肛门 Varchar2(50) Path '$.肛门',
                          生殖器 Varchar2(50) Path '$.生殖器',咽喉部吸出物量 Varchar2(50) Path '$.咽喉部吸出物量',咽喉部吸出物性状 Varchar2(50) Path '$.咽喉部吸出物性状',
                          气管插管吸出物量 Varchar2(50) Path '$.气管插管吸出物量',气管插管吸出物性状 Varchar2(50) Path '$.气管插管吸出物性状',
                          吸氧方式 Varchar2(50) Path '$.吸氧方式',新生儿抢救状况 Varchar2(50) Path '$.新生儿抢救状况',抢救药物 Varchar2(50) Path '$.抢救药物',
                          早吸吮 Varchar2(50) Path '$.早吸吮',皮肤接触 Varchar2(50) Path '$.皮肤接触')) As D,
     Json_Table(Nvl(b.Newbornscore, '{"隐藏3":"3"}'),
                 '$' Columns(隐藏3 Varchar2(50) Path '$.隐藏3', 心率1分钟 Varchar2(50) Path '$.心率1分钟',
                          心率5分钟 Varchar2(50) Path '$.心率5分钟', 心率10分钟 Varchar2(50) Path '$.心率10分钟',
                          呼吸1分钟 Varchar2(50) Path '$.呼吸1分钟', 呼吸5分钟 Varchar2(50) Path '$.呼吸5分钟',
                          呼吸10分钟 Varchar2(50) Path '$.呼吸10分钟', 喉反射1分钟 Varchar2(50) Path '$.喉反射1分钟',
                          喉反射5分钟 Varchar2(50) Path '$.喉反射5分钟', 喉反射10分钟 Varchar2(50) Path '$.喉反射10分钟',
                          肌张力1分钟 Varchar2(50) Path '$.肌张力1分钟', 肌张力5分钟 Varchar2(50) Path '$.肌张力5分钟',
                          肌张力10分钟 Varchar2(50) Path '$.肌张力10分钟', 肤色1分钟 Varchar2(50) Path '$.肤色1分钟',
                          肤色5分钟 Varchar2(50) Path '$.肤色5分钟', 肤色10分钟 Varchar2(50) Path '$.肤色10分钟',
                          总分1分钟 Varchar2(50) Path '$.总分1分钟', 总分5分钟 Varchar2(50) Path '$.总分5分钟',
                          总分10分钟 Varchar2(50) Path '$.总分10分钟')) As E,
     Json_Table(Nvl(b.Otherinf, '{"隐藏4":"4"}'),
                 '$' Columns(隐藏4 Varchar2(50) Path '$.隐藏4', 出孕期产时合并症及用药情况 Varchar2(50) Path '$.出孕期产时合并症及用药情况',
                          出生前胎儿情况 Varchar2(50) Path '$.出生前胎儿情况', 婴儿出生时抢救情况 Varchar2(50) Path '$.婴儿出生时抢救情况',
                          出生缺陷 Varchar2(50) Path '$.出生缺陷', 死亡时间 Varchar2(50) Path '$.死亡时间',
                          诊断 Varchar2(50) Path '$.诊断')) As F
Where a.Mid = b.Mid
Order By Addtime;

create or replace force view ZLSOL.v_sol_inf_checkinroom as
Select a.mid, b.入房目的,b.入房时间,b.医疗病历,b.护理病历,b.分娩知情通知书,b.宫缩规律性,b.胎心率,b.胎心次数,b.破膜情况,b.是否有合并症,b.种类,b.输液单,b.静脉通道,b.局部情况,b.特殊药物,b.其他,b.交班者,b.接班者
From SOL_INF_CHECKINROOM a,JSON_TABLE(a.Content,'$' columns(
入房目的       Varchar2(50) PATH '$.入房目的',
入房时间       Varchar2(50) PATH '$.入房时间',
医疗病历       Varchar2(10) PATH '$.医疗病历',
护理病历       Varchar2(10) PATH '$.护理病历',
分娩知情通知书      Varchar2(10) PATH '$.分娩知情通知书',
宫缩规律性   Varchar2(20) PATH '$.宫缩规律性',
胎心率        Varchar2(10) PATH '$.胎心率',
胎心次数       Varchar2(10) PATH '$.胎心次数',
破膜情况       Varchar2(10) PATH '$.破膜情况',
是否有合并症       Varchar2(10) PATH '$.是否有合并症',
种类         Varchar2(50) PATH '$.种类',
输液单        Varchar2(10) PATH '$.输液单',
静脉通道       Varchar2(10) PATH '$.静脉通道',
局部情况       Varchar2(50) PATH '$.局部情况',
特殊药物       Varchar2(50) PATH '$.特殊药物',
其他         Varchar2(50) PATH '$.其他',
交班者        Varchar2(50) PATH '$.交班者',
接班者               Varchar2(50) PATH '$.接班者'
)) as b;

create or replace force view ZLSOL.v_sol_inf_checkoutroom as
Select a.mid, b.OUTROOMTIME,b.出房状态,b.医疗病历,b.护理病历,b.静脉通道,b.局部情况,b.会阴裂伤,b.会阴切开术,b.会阴切口缝合,b.会阴水肿,b.会阴血肿,b.产后出血,b.出血量,b.特殊药物,b.交班者,b.接班者,b.药物,b.备注
From SOL_INF_CheckOutRoom a,JSON_TABLE(a.Content,'$' columns(
OUTROOMTIME     Varchar2(50) PATH '$.OUTROOMTIME',
出房状态     Varchar2(50) PATH '$.出房状态',
医疗病历     Varchar2(10) PATH '$.医疗病历',
护理病历     Varchar2(10) PATH '$.护理病历',
静脉通道     Varchar2(10) PATH '$.静脉通道',
局部情况     Varchar2(50) PATH '$.局部情况',
会阴裂伤     Varchar2(20) PATH '$.会阴裂伤',
会阴切开术   Varchar2(20) PATH '$.会阴切开术',
会阴切口缝合 Varchar2(10) PATH '$.会阴切口缝合',
会阴水肿     Varchar2(10) PATH '$.会阴水肿',
会阴血肿     Varchar2(10) PATH '$.会阴血肿',
产后出血     Varchar2(10) PATH '$.产后出血',
出血量       Number(5) PATH '$.出血量',
特殊药物     Varchar2(50) PATH '$.特殊药物',
交班者       Varchar2(20) PATH '$.交班者',
接班者       Varchar2(20) PATH '$.接班者',
药物         Varchar2(50) PATH '$.药物',
备注         Varchar2(50) PATH '$.备注'
)) as b;

create or replace force view ZLSOL.v_sol_inf_delivery as
Select a.Mid, b.隐藏1, b.BEGINT, b.ALLT, b.OUTT, b.ALLOUTT, b.第一产程, b.第二产程, b.第三产程, b.宫缩情况,
       b.结扎, b.破膜方式, b.破膜时间, b.羊水性状, b.羊水量, b.羊水颜色, b.胎膜清理方式, b.胎盘剥离方式, b.胎盘完整度,
       b.胎盘胎膜残留, b.胎盘体积, b.胎盘形态, b.胎盘大小, b.胎盘重量, b.脐带附着, b.脐带长度, b.脐带真假结, b.脐带,b.绕颈周数, b.娩出方式,
       b.娩出胎方位, b.产瘤大小, b.产瘤部位, b.会阴裂伤程度, b.会阴裂伤切口, b.会阴裂伤缝合, b.会阴裂伤麻醉, b.宫颈裂伤长度, b.宫颈裂伤部位, b.宫颈裂伤状况,
       b.阴道裂伤部位大小, b.阴道裂伤血肿大小, b.产后即刻收缩压,b.产后即刻舒张压,b.产后1小时收缩压,b.产后1小时舒张压,b.产后2小时收缩压,b.产后2小时舒张压,
       b.DNOW,b.DONE,b.DTWO,b.产后出血总量,b.产后即刻脉搏,b.产后1小时脉搏,b.产后2小时脉搏, b.产时用药, b.产后用药,b.无痛分娩用药,
       b.产后诊断, b.特殊情况, d.隐藏3,d.出产房时间,d.出产房宫高脐下, d.护送人, d.接生人, d.记录人
From Sol_Inf_Delivery A,
     Json_Table(Nvl(a.Newborndetail, '{"隐藏1":"1"}'),
                 '$'Columns(隐藏1 Varchar2(50) Path '$.隐藏1', BEGINT Varchar2(19) Path '$.BEGINT',
                          ALLT Varchar2(19) Path '$.ALLT', OUTT Varchar2(19) Path '$.OUTT',
                          ALLOUTT Varchar2(19) Path '$.ALLOUTT', 第一产程 Varchar2(50) Path '$.第一产程',
                          第二产程 Varchar2(50) Path '$.第二产程', 第三产程 Varchar2(50) Path '$.第三产程',结扎 Varchar2(50) Path '$.结扎',
                          破膜方式 Varchar2(50) Path '$.破膜方式', 破膜时间 Varchar2(19) Path '$.破膜时间', 羊水性状 Varchar2(50) Path '$.羊水性状',
                          羊水量 Varchar2(50) Path '$.羊水量', 羊水颜色 Varchar2(50) Path '$.羊水颜色',
                          胎膜清理方式 Varchar2(50) Path '$.胎膜清理方式', 胎盘剥离方式 Varchar2(50) Path '$.胎盘剥离方式',
                          胎盘完整度 Varchar2(50) Path '$.胎盘完整度', 胎盘胎膜残留 Varchar2(50) Path '$.胎盘胎膜残留',
                          胎盘体积 Varchar2(50) Path '$.胎盘体积', 胎盘形态 Varchar2(50) Path '$.胎盘形态', 胎盘大小 Varchar2(50) Path '$.胎盘大小',
                          胎盘重量 Varchar2(50) Path '$.胎盘重量', 脐带附着 Varchar2(50) Path '$.脐带附着', 脐带长度 Varchar2(50) Path '$.脐带长度',
                          脐带真假结 Varchar2(50) Path '$.脐带真假结',脐带 Varchar2(50) Path '$.脐带', 绕颈周数 Varchar2(50) Path '$.绕颈周数',
                          娩出方式 Varchar2(50) Path '$.娩出方式',娩出胎方位 Varchar2(50) Path '$.娩出胎方位', 产瘤大小 Varchar2(50) Path '$.产瘤大小',
                          产瘤部位 Varchar2(50) Path '$.产瘤部位', 会阴裂伤程度 Varchar2(50) Path '$.会阴裂伤程度',
                          会阴裂伤切口 Varchar2(50) Path '$.会阴裂伤切口', 会阴裂伤缝合 Varchar2(50) Path '$.会阴裂伤缝合',
                          会阴裂伤麻醉 Varchar2(50) Path '$.会阴裂伤麻醉', 宫颈裂伤长度 Varchar2(50) Path '$.宫颈裂伤长度',
                          宫颈裂伤部位 Varchar2(50) Path '$.宫颈裂伤部位', 宫颈裂伤状况 Varchar2(50) Path '$.宫颈裂伤状况',
                          阴道裂伤部位大小 Varchar2(50) Path '$.阴道裂伤部位大小', 阴道裂伤血肿大小 Varchar2(50) Path '$.阴道裂伤血肿大小',
                          宫缩情况 Varchar2(50) Path '$.宫缩情况',产后即刻收缩压 Varchar2(50) Path '$.产后即刻收缩压',产后即刻舒张压 Varchar2(50) Path '$.产后即刻舒张压',
                          产后1小时收缩压 Varchar2(50) Path '$.产后1小时收缩压',产后1小时舒张压 Varchar2(50) Path '$.产后1小时舒张压',
                          产后2小时收缩压 Varchar2(50) Path '$.产后2小时收缩压',产后2小时舒张压 Varchar2(50) Path '$.产后2小时舒张压',
                          产后即刻脉搏 Varchar2(50) Path '$.产后即刻脉搏',产后1小时脉搏 Varchar2(50) Path '$.产后1小时脉搏',产后2小时脉搏 Varchar2(50) Path '$.产后2小时脉搏',
                          DNOW Varchar2(50) Path '$.DNOW',DONE Varchar2(50) Path '$.DONE',DTWO Varchar2(50) Path '$.DTWO',产后出血总量 Varchar2(50) Path '$.产后出血总量',
                          产时用药 Varchar2(50) Path '$.产时用药', 产后用药 Varchar2(50) Path '$.产后用药',无痛分娩用药 Varchar2(50) Path '$.无痛分娩用药',
                          产后诊断 Varchar2(50) Path '$.产后诊断',特殊情况 Varchar2(50) Path '$.特殊情况')) As B,
     Json_Table(Nvl(a.Deliveryinf, '{"隐藏3":"1"}'),
                 '$' Columns(隐藏3 Varchar2(1) Path '$.隐藏3', 出产房时间 Varchar2(50) Path '$.出产房时间',出产房宫高脐下 Varchar2(50) Path '$.出产房宫高脐下',
                          护送人 Varchar2(50) Path '$.护送人', 接生人 Varchar2(50) Path '$.接生人', 记录人 Varchar2(50) Path '$.记录人')) As D;

create or replace force view ZLSOL.v_sol_inf_equipment as
Select a.mid, b.侧切剪产前,b.侧切剪术中,b.侧切剪产后,b.脐带剪产前,b.脐带剪术中,b.脐带剪产后,b.止血钳产前,b.止血钳术中,b.止血钳产后,b.牙镊产前,b.牙镊术中,b.牙镊产后,b.持针器产前,b.持针器术中,b.持针器产后,b.穿刺针产前,b.穿刺针术中,b.穿刺针产后,b.洗耳球产前,b.洗耳球术中,b.洗耳球产后,b.缝合针产前,b.缝合针术中,b.缝合针产后,b.拉钩产前,b.拉钩术中,b.拉钩产后,b.宫颈钳产前,b.宫颈钳术中,b.宫颈钳产后,b.窥器产前,b.窥器术中,b.窥器产后,b.刮匙产前,b.刮匙术中,b.刮匙产后,b.艾利斯产前,b.艾利斯术中,b.艾利斯产后,b.产钳产前,b.产钳术中,b.产钳产后,b.纱布产前,b.纱布术中,b.纱布产后,b.卵圆钳产前,b.卵圆钳术中,b.卵圆钳产后
From SOL_INF_Equipment a,JSON_TABLE(a.Content,'$' columns(
侧切剪产前   Number(2) PATH '$.侧切剪产前',
侧切剪术中   Number(2) PATH '$.侧切剪术中',
侧切剪产后   Number(2) PATH '$.侧切剪产后',
脐带剪产前   Number(2) PATH '$.脐带剪产前',
脐带剪术中   Number(2) PATH '$.脐带剪术中',
脐带剪产后   Number(2) PATH '$.脐带剪产后',
止血钳产前   Number(2) PATH '$.止血钳产前',
止血钳术中   Number(2) PATH '$.止血钳术中',
止血钳产后   Number(2) PATH '$.止血钳产后',
牙镊产前   Number(2) PATH '$.牙镊产前',
牙镊术中   Number(2) PATH '$.牙镊术中',
牙镊产后   Number(2) PATH '$.牙镊产后',
持针器产前   Number(2) PATH '$.持针器产前',
持针器术中   Number(2) PATH '$.持针器术中',
持针器产后   Number(2) PATH '$.持针器产后',
穿刺针产前   Number(2) PATH '$.穿刺针产前',
穿刺针术中   Number(2) PATH '$.穿刺针术中',
穿刺针产后   Number(2) PATH '$.穿刺针产后',
洗耳球产前   Number(2) PATH '$.洗耳球产前',
洗耳球术中   Number(2) PATH '$.洗耳球术中',
洗耳球产后   Number(2) PATH '$.洗耳球产后',
缝合针产前   Number(2) PATH '$.缝合针产前',
缝合针术中   Number(2) PATH '$.缝合针术中',
缝合针产后   Number(2) PATH '$.缝合针产后',
拉钩产前   Number(2) PATH '$.拉钩产前',
拉钩术中   Number(2) PATH '$.拉钩术中',
拉钩产后   Number(2) PATH '$.拉钩产后',
宫颈钳产前   Number(2) PATH '$.宫颈钳产前',
宫颈钳术中   Number(2) PATH '$.宫颈钳术中',
宫颈钳产后   Number(2) PATH '$.宫颈钳产后',
窥器产前   Number(2) PATH '$.窥器产前',
窥器术中   Number(2) PATH '$.窥器术中',
窥器产后   Number(2) PATH '$.窥器产后',
刮匙产前   Number(2) PATH '$.刮匙产前',
刮匙术中   Number(2) PATH '$.刮匙术中',
刮匙产后   Number(2) PATH '$.刮匙产后',
艾利斯产前   Number(2) PATH '$.艾利斯产前',
艾利斯术中   Number(2) PATH '$.艾利斯术中',
艾利斯产后   Number(2) PATH '$.艾利斯产后',
产钳产前   Number(2) PATH '$.产钳产前',
产钳术中   Number(2) PATH '$.产钳术中',
产钳产后   Number(2) PATH '$.产钳产后',
纱布产前   Number(2) PATH '$.纱布产前',
纱布术中   Number(2) PATH '$.纱布术中',
纱布产后   Number(2) PATH '$.纱布产后',
卵圆钳产前   Number(2) PATH '$.卵圆钳产前',
卵圆钳术中   Number(2) PATH '$.卵圆钳术中',
卵圆钳产后   Number(2) PATH '$.卵圆钳产后'
)) as b;

create or replace force view ZLSOL.v_sol_inf_newborns as
Select a.Mid, b.Bid, b.Babyno, b.Addtime, b.Recorder, a.Name, a.Old, a.Bedno, a.Pno, b.Sex, d.隐藏2,d.身长,d.体重,d.头围,d.胸围,d.一般情况反应,d.一般情况面色,d.一般情况皮肤,d.一般情况毳毛,d.头部变形,d.颅骨重叠,d.胎头水肿血肿,d.胎头水肿大小,d.前囟,d.张力,d.眼神,d.口腔,d.心,d.乳结,d.肝,d.脾,d.四肢,d.外展试验,d.肛门,d.生殖器, e.隐藏3,e.心率1分钟,e.心率5分钟,e.心率10分钟,e.呼吸1分钟,e.呼吸5分钟,e.呼吸10分钟,e.喉反射1分钟,e.喉反射5分钟,e.喉反射10分钟,e.肌张力1分钟,e.肌张力5分钟,e.肌张力10分钟,e.肤色1分钟,e.肤色5分钟,e.肤色10分钟,e.总分1分钟,e.总分5分钟,e.总分10分钟, f.隐藏4,f.出孕期产时合并症及用药情况,f.出生前胎儿情况,f.婴儿出生时抢救情况,f.出生缺陷,f.母乳喂养指导,f.诊断
From Sol_Inf_Puerpera A, Sol_Inf_Newborns B,
     Json_Table(Nvl(b.Newborninf, '{"隐藏2":"2"}'),
                 '$' Columns(隐藏2 Varchar2(50) Path '$.隐藏2', 身长 Varchar2(50) Path '$.身长', 体重 Varchar2(50) Path '$.体重',
                          头围 Varchar2(50) Path '$.头围', 胸围 Varchar2(50) Path '$.胸围', 一般情况反应 Varchar2(50) Path '$.一般情况反应',
                          一般情况面色 Varchar2(50) Path '$.一般情况面色', 一般情况皮肤 Varchar2(50) Path '$.一般情况皮肤',
                          一般情况毳毛 Varchar2(50) Path '$.一般情况毳毛', 头部变形 Varchar2(50) Path '$.头部变形',
                          颅骨重叠 Varchar2(50) Path '$.颅骨重叠', 胎头水肿血肿 Varchar2(50) Path '$.胎头水肿血肿',
                          胎头水肿大小 Varchar2(50) Path '$.胎头水肿大小', 前囟 Varchar2(50) Path '$.前囟', 张力 Varchar2(50) Path '$.张力',
                          眼神 Varchar2(50) Path '$.眼神', 口腔 Varchar2(50) Path '$.口腔', 心 Varchar2(50) Path '$.心',
                          乳结 Varchar2(50) Path '$.乳结', 肝 Varchar2(50) Path '$.肝', 脾 Varchar2(50) Path '$.脾',
                          四肢 Varchar2(50) Path '$.四肢', 外展试验 Varchar2(50) Path '$.外展试验', 肛门 Varchar2(50) Path '$.肛门',
                          生殖器 Varchar2(50) Path '$.生殖器')) As D,
     Json_Table(Nvl(b.Newbornscore, '{"隐藏3":"3"}'),
                 '$' Columns(隐藏3 Varchar2(50) Path '$.隐藏3', 心率1分钟 Varchar2(50) Path '$.心率1分钟',
                          心率5分钟 Varchar2(50) Path '$.心率5分钟', 心率10分钟 Varchar2(50) Path '$.心率10分钟',
                          呼吸1分钟 Varchar2(50) Path '$.呼吸1分钟', 呼吸5分钟 Varchar2(50) Path '$.呼吸5分钟',
                          呼吸10分钟 Varchar2(50) Path '$.呼吸10分钟', 喉反射1分钟 Varchar2(50) Path '$.喉反射1分钟',
                          喉反射5分钟 Varchar2(50) Path '$.喉反射5分钟', 喉反射10分钟 Varchar2(50) Path '$.喉反射10分钟',
                          肌张力1分钟 Varchar2(50) Path '$.肌张力1分钟', 肌张力5分钟 Varchar2(50) Path '$.肌张力5分钟',
                          肌张力10分钟 Varchar2(50) Path '$.肌张力10分钟', 肤色1分钟 Varchar2(50) Path '$.肤色1分钟',
                          肤色5分钟 Varchar2(50) Path '$.肤色5分钟', 肤色10分钟 Varchar2(50) Path '$.肤色10分钟',
                          总分1分钟 Varchar2(50) Path '$.总分1分钟', 总分5分钟 Varchar2(50) Path '$.总分5分钟',
                          总分10分钟 Varchar2(50) Path '$.总分10分钟')) As E,
     Json_Table(Nvl(b.Otherinf, '{"隐藏4":"4"}'),
                 '$' Columns(隐藏4 Varchar2(50) Path '$.隐藏4', 出孕期产时合并症及用药情况 Varchar2(50) Path '$.出孕期产时合并症及用药情况  ',
                          出生前胎儿情况 Varchar2(50) Path '$.出生前胎儿情况  ', 婴儿出生时抢救情况 Varchar2(50) Path '$.婴儿出生时抢救情况  ',
                          出生缺陷 Varchar2(50) Path '$.出生缺陷  ', 母乳喂养指导 Varchar2(50) Path '$.母乳喂养指导  ',
                          诊断 Varchar2(50) Path '$.诊断  ')) As F
Where a.Mid = b.Mid
Order By Addtime;

create or replace force view ZLSOL.v_sol_inf_puerpera as
Select Name, Mid, Old, LPad(Bedno, 10) Bedno, Pno, Diagnosis, Status, Decode(Expectant, 1, '√', '') 待产,
       Decode(Checkinroom, 1, '√', '') 入房, Decode(Birth, 1, '√', '') 临产, Decode(Druglabor, 1, '√', '') 引产,
       Decode(Delivery, 1, '√', '') 分娩, Decode(Newborns, 1, '√', '') 新生儿, Decode(Postpartum, 1, '√', '') 产后,
       Decode(Checkoutroom, 1, '√', '') 出房,Decode(Equipment, 1, '√', '') 器械,
       Outtime,pid,tid
From Sol_Inf_Puerpera;

create or replace force view ZLSOL.v_sol_rs_birth as
Select a.mid,b.妊次,b.产次,b.血型,b.既往妊娠史,b.末次月经,b.预产期,b.髂前上棘间径,b.髂嵴间径,b.坐骨结节间径,b.骶耻外径,b.骶骨弧度,b.骶骨关节,b.坐骨切迹,b.坐骨髂,b.并发症,b.产前记录特征,b.检查时间,b.收缩压,b.舒张压,b.体温,b.脉搏,b.胎心率,b.胎儿大小,b.宫缩规律性,b.胎位,b.破膜情况,b.先露,b.宫口,b.检查者,b.宫缩开始时间,b.破膜时间,b.入院处理
From SOL_RS_BIRTH a,JSON_TABLE(Nvl(a.CONTENT,'{隐藏:1}'),'$' columns(
妊次            Number(3)    PATH '$.妊次',
产次            Number(3)    PATH '$.产次',
血型            Varchar2(10) PATH '$.血型',
既往妊娠史      Varchar2(50) PATH '$.既往妊娠史',
末次月经        Varchar2(20) PATH '$.末次月经',
预产期          Varchar2(20) PATH '$.预产期',
髂前上棘间径    Number(5) PATH '$.髂前上棘间径',
髂嵴间径        Number(5) PATH '$.髂嵴间径',
坐骨结节间径    Number(5) PATH '$.坐骨结节间径',
骶耻外径        Number(5) PATH '$.骶耻外径',
骶骨弧度        Varchar2(10) PATH '$.骶骨弧度',
骶骨关节        Varchar2(10) PATH '$.骶骨关节',
坐骨切迹        Varchar2(10) PATH '$.坐骨切迹',
坐骨髂          Varchar2(10) PATH '$.坐骨髂',
并发症          Varchar2(100) PATH '$.并发症',
产前记录特征    Varchar2(100) PATH '$.产前记录特征',
检查时间        Varchar2(20) PATH '$.检查时间',
收缩压        Number(3) PATH '$.收缩压',
舒张压        Number(3) PATH '$.舒张压',
体温            Number(4,2)  PATH '$.体温',
脉搏            Varchar2(10) PATH '$.脉搏',
胎心率            Varchar2(10) PATH '$.胎心率',
胎儿大小        Number(5,2) PATH '$.胎儿大小',
宫缩规律性      Varchar2(10) PATH '$.宫缩规律性',
胎位            Varchar2(10) PATH '$.胎位',
破膜情况        Varchar2(10) PATH '$.破膜情况',
先露            Varchar2(2) PATH '$.先露',
宫口            Number(4,2) PATH '$.宫口',
检查者          Varchar2(50) PATH '$.检查者',
宫缩开始时间    Varchar2(20) PATH '$.宫缩开始时间',
破膜时间        Varchar2(20) PATH '$.破膜时间',
入院处理        Varchar2(100) PATH '$.入院处理'
)) as b;

create or replace force view ZLSOL.v_sol_rs_birth_course as
Select  a.courseid,a.mid,b.检查时间,b.是否剖宫产,b.胎方位,b.收缩压,b.舒张压,b.体温,b.脉搏,b.胎心率,b.宫缩强度,b.宫缩持续,b.宫缩间隔,b.宫颈厚薄,b.宫口,b.破膜情况,b.先露,b.处理,b.检查者
From SOL_RS_BIRTH_COURSE a,JSON_TABLE(a.CONTENT,'$' columns(
检查时间        Varchar2(20)  PATH '$.检查时间',
是否剖宫产    Varchar2(20)  PATH '$.是否剖宫产',
胎方位 Varchar2(20)  PATH '$.胎方位',
收缩压        Number(3) PATH '$.收缩压',
舒张压        Number(3) PATH '$.舒张压',
体温        Number(4,2)  PATH '$.体温',
脉搏        Varchar2(10) PATH '$.脉搏',
胎心率        Varchar2(10) PATH '$.胎心率',
宫缩强度    Varchar2(10) PATH '$.宫缩强度',
宫缩持续  Varchar2(10) PATH '$.宫缩持续',
宫缩间隔  Varchar2(10) PATH '$.宫缩间隔',
宫颈厚薄    Varchar2(10) PATH '$.宫颈厚薄',
宫口        Number(4,2) PATH '$.宫口',
破膜情况    Varchar2(10) PATH '$.破膜情况',
先露        Number(2) PATH '$.先露'，
处理        Varchar2(200) PATH '$.处理'，
检查者      Varchar2(50) PATH '$.检查者'
)) as b;

create or replace force view ZLSOL.v_sol_rs_druglabor as
Select Mid, To_Char(日期, 'YYYY-MM-DD HH24:MI') 日期, 引产指征, 引产方法,检查者 from Sol_Rs_Druglabor;

create or replace force view ZLSOL.v_sol_rs_druglabor_list as
Select a.Mid, a.Courseid ID, b.记录时间,b.收缩压,b.舒张压,b.脉搏,b.胎心率,b.宫缩强度,b.宫缩持续,b.宫缩间隔,b.宫口,b.先露,b.羊水量,b.羊水性状,b.处理,b.记录人,b.剂量,b.滴速
From ZLSOL.Sol_Rs_Druglabor_List a,
     Json_Table(a.Content,'$' Columns(
     记录时间 Varchar2(20) Path '$.记录时间',
     剂量 Number(3,1) Path '$.剂量',
     滴速 Number(3) Path '$.滴速',
     收缩压        Number(3) PATH '$.收缩压',
     舒张压        Number(3) PATH '$.舒张压',
     脉搏 Number(3) Path '$.脉搏',
     胎心率 Number(3) Path '$.胎心率',
     宫缩强度 Varchar2(10) Path '$.宫缩强度',
     宫缩持续 Number(3) Path '$.宫缩持续',
     宫缩间隔 Number(2) Path '$.宫缩间隔',
     宫口 Number(3) Path '$.宫口',
     先露 Varchar2(10) Path '$.先露',
     羊水量 Number(4) Path '$.羊水量',
     羊水性状 Varchar2(10) Path '$.羊水性状',
     处理 Varchar2(100) Path '$.处理',
     记录人 Varchar2(100) Path '$.记录人')) b;

create or replace force view ZLSOL.v_sol_rs_expectant as
Select a.mid,a.courseid,b.记录时间,b.胎方位,b.收缩压,b.舒张压,b.宫高,b.腹围,b.胎动计数早,b.胎动计数中,b.胎动计数晚,b.胎心率,b.先露,b.宫口,b.破膜情况,b.羊水性状,b.宫缩强度,b.宫缩持续,b.宫缩间隔,b.处理,b.检查者
From SOL_RS_EXPECTANT a,JSON_TABLE(a.Content,'$' columns(
记录时间    Varchar2(50) PATH '$.记录时间',
胎方位  Varchar2(20) PATH '$.胎方位',
收缩压  Number(3)  PATH '$.收缩压',
舒张压  Number(3)  PATH '$.舒张压',
宫高     Number(4,2) PATH '$.宫高',
腹围     Varchar2(20) PATH '$.腹围',
胎动计数早     Number(3) PATH '$.胎动计数早',
胎动计数中     Number(3) PATH '$.胎动计数中',
胎动计数晚   Number(3) PATH '$.胎动计数晚',
胎心率 varchar2(50) PATH '$.胎心率',
先露     Varchar2(20) PATH '$.先露',
宫口     Varchar2(20) PATH '$.宫口',
破膜情况     Varchar2(20) PATH '$.破膜情况',
羊水性状      Varchar2(20) PATH '$.羊水性状',
宫缩强度     Varchar2(20) PATH '$.宫缩强度',
宫缩持续       Varchar2(20) PATH '$.宫缩持续',
宫缩间隔       Varchar2(20) PATH '$.宫缩间隔',
处理     Varchar2(500) PATH '$.处理',
检查者       Varchar2(20) PATH '$.检查者'
)) as b;

create or replace force view ZLSOL.v_sol_rs_postpartum as
Select a.Mid, 分娩日期, 入产房时间, 分娩方式, 出产房时间, 出产房时bp, 出产房时脉搏, 出产房时宫高脐下, 出产房时阴道流血, 出产房时一般情况, 会阴,  拆线
From ZLSOL.Sol_Rs_Postpartum A,
     Json_Table(a.Content,
                 '$' Columns(分娩日期 varchar2(20) Path '$.分娩日期', 入产房时间 varchar2(20) Path '$.入产房时间', 分娩方式 Varchar2(20) Path '$.分娩方式',
                          出产房时间 varchar2(20) Path '$.出产房时间', 出产房时bp varchar2(7) Path '$.出产房时BP', 出产房时脉搏 Number(3) Path '$.出产房时脉搏',
                          出产房时宫高脐下 Number(2) Path '$.出产房时宫高脐下', 出产房时阴道流血 Number(3) Path '$.出产房时阴道流血',
                          出产房时一般情况 Varchar2(10) Path '$.出产房时一般情况', 会阴 Varchar2(20) Path '$.会阴', 拆线 Varchar2(10) Path '$.拆线'));

create or replace force view ZLSOL.v_sol_rs_postpartum_list as
Select a.Mid, a.Courseid ID, 记录时间, 乳量, 乳房红肿, 乳头, 子宫宫高, 子宫压痛, 恶露量, 恶露颜色, 恶露臭味, 会阴正常, 会阴红肿, 会阴其他, 小便, 大便, 特殊情况, 签名
From ZLSOL.Sol_Rs_Postpartum_List A,
     Json_Table(a.Content,
                 '$' Columns(记录时间 Varchar2(20) Path '$.记录时间', 乳量 Number(4) Path '$.乳量', 乳房红肿 Varchar2(10) Path '$.乳房红肿',
                          乳头 Varchar2(50) Path '$.乳头', 子宫宫高 Number(3) Path '$.子宫宫高', 子宫压痛 Varchar2(50) Path '$.子宫压痛',
                          恶露量 Number(4) Path '$.恶露量', 恶露颜色 Varchar2(20) Path '$.恶露颜色', 恶露臭味 Varchar2(20) Path '$.恶露臭味',
                          会阴正常 Varchar2(10) Path '$.会阴正常', 会阴红肿 Varchar2(10) Path '$.会阴红肿', 会阴其他 Varchar2(50) Path '$.会阴其他',
                          小便 Varchar2(50) Path '$.小便', 大便 Varchar2(50) Path '$.大便', 特殊情况 Varchar2(100) Path '$.特殊情况',
                          签名 Varchar2(100) Path '$.签名'));

CREATE OR REPLACE Function ZLSOL.Sol_Getsdate
(
  Dbegin_In In Varchar2,
  Dend_In   In Varchar2
) Return Varchar2 Is
  v_Temp Varchar2(100);
Begin
  If Dbegin_In Is Not Null And Dend_In Is Not Null Then
    Select '|' || Extract(Day From(To_Date(Dbegin_In, 'YYYY-MM-DD hh24:mi') - To_Date(Dend_In, 'YYYY-MM-DD hh24:mi')) Day To
                           Second) || '天' || '|' ||
            Extract(Hour From(To_Date(Dbegin_In, 'YYYY-MM-DD hh24:mi') - To_Date(Dend_In, 'YYYY-MM-DD hh24:mi')) Day To
                    Second) || '时' || '|' ||
            Extract(Minute From(To_Date(Dbegin_In, 'YYYY-MM-DD hh24:mi') - To_Date(Dend_In, 'YYYY-MM-DD hh24:mi')) Day To
                    Second) || '分'
    Into v_Temp
    From Dual;
    v_Temp := Replace(v_Temp, '|0天', '');
    v_Temp := Replace(v_Temp, '|0时', '');
    v_Temp := Replace(v_Temp, '|0分', '');
    v_Temp := Replace(v_Temp, '|', '');
  End If;
  Return v_Temp;
End Sol_Getsdate;
/
create table ZLSOL.HIS_病人新生儿记录
(
  病人id NUMBER(18) not null,
  主页id NUMBER(18) not null,
  序号   NUMBER(3) not null,
  婴儿姓名 VARCHAR2(100),
  婴儿性别 VARCHAR2(4),
  分娩次数 NUMBER(3),
  分娩方式 VARCHAR2(20),
  胎儿状况 VARCHAR2(20),
  出生时间 DATE,
  身长   NUMBER(16,5),
  体重   NUMBER(16,5),
  血型   VARCHAR2(10),
  备注说明 VARCHAR2(100),
  死亡时间 DATE,
  登记时间 DATE,
  登记人  VARCHAR2(20)
)
tablespace ZLSOL_DATA
  pctfree 5
  initrans 1
  maxtrans 255
  storage
  (
    initial 64K
    next 1M
    minextents 1
    maxextents unlimited
  );

create or replace force view zlsol.v_his_病人新生儿记录 as
select d.pid 病人id,d.tid 住院次数,b.Babyno 序号,d.name||decode(b.Sex,'男','之子','之女')||t.顺序 as 婴儿姓名,
b.Sex as 婴儿性别,
c.妊次 as 分娩次数,
a.娩出方式,b.胎儿状况,
b.身长,b.体重,b.血型,
b.boutt as 出生时间,
b.死亡时间,'' as 备注说明,
b.Recorder 登记人,
b.Addtime 登记时间
from V_SOL_INF_DELIVERY a,v_newborn b,v_sol_rs_birth c,sol_inf_puerpera d,
(select mid,sex,babyno,row_number() over(partition by mid,sex order by babyno) 顺序 from SOL_INF_NEWBORNS t  ) t
where a.Mid=b.Mid and a.mid=d.mid and a.mid=t.mid and b.Babyno=t.babyno
and b.Mid=c.mid(+);

CREATE OR REPLACE Procedure ZLSOL.his_病人新生儿登记_revise
(
  mid_In   SOL_INF_NEWBORNS.mid%Type,
  bid_In   SOL_INF_NEWBORNS.bid%Type,
  state_in number    -----2增加，修改 ，3删除
) As
  n_病人id Number(20);
  n_主页id   Number(20);
  n_count number(2);
  babyno_In number(2);
Begin

  select pid,tid into n_病人id,n_主页id from sol_inf_puerpera where mid=mid_in;
  select babyno into babyno_In  from SOL_INF_NEWBORNS where bid=bid_in;
  select count(1) into n_count from 病人新生儿记录@ZLHIS_DBL where 病人id=n_病人id and 主页id = n_主页id and 序号=babyno_In;
  --新生儿新增修改
  if  state_in=2 then
   if n_count=0 then  ----新增
      insert into   his_病人新生儿记录
      (病人id,主页id,序号,婴儿姓名,婴儿性别,分娩次数,分娩方式,胎儿状况,出生时间,身长,体重,血型,备注说明,死亡时间,登记时间,登记人)
      select 病人id,住院次数,序号,婴儿姓名,婴儿性别,分娩次数,娩出方式,胎儿状况,出生时间,身长,体重,血型,备注说明,死亡时间,登记时间,登记人
       from   ( select  d.pid 病人id,d.tid 住院次数,b.Babyno 序号,d.name||decode(b.Sex,'男','之子','之女')||t.顺序 as 婴儿姓名,
              b.Sex as 婴儿性别,c.妊次 as 分娩次数,a.娩出方式,b.胎儿状况,b.身长,b.体重,b.血型,to_date(b.boutt,'yyyy-mm-dd hh24:mi:ss') as 出生时间,
              b.死亡时间,'' as 备注说明,b.Recorder 登记人,b.Addtime 登记时间
              from V_SOL_INF_DELIVERY a,v_newborn b,v_sol_rs_birth c,sol_inf_puerpera d,
              (select mid,sex,babyno,row_number() over(partition by mid,sex order by babyno) 顺序 from SOL_INF_NEWBORNS t  ) t
              where a.Mid=b.Mid and a.mid=d.mid and a.mid=t.mid and b.Babyno=t.babyno
              and b.Mid=c.mid(+) and a.mid=mid_In
                ) where 序号=babyno_In;
      select count(*) into n_count from his_病人新生儿记录;
      dbms_output.put_line(n_count);
        insert into 病人新生儿记录@ZLHIS_DBL(病人id,主页id,序号,婴儿姓名,婴儿性别,分娩次数,分娩方式,胎儿状况,出生时间,身长,体重,血型,备注说明,死亡时间,登记时间,登记人)
         select  病人id,主页id,序号,婴儿姓名,婴儿性别,分娩次数,分娩方式,胎儿状况,出生时间,身长,体重,血型,备注说明,死亡时间,登记时间,登记人 from his_病人新生儿记录 ;
      Zl_病区自动标记_Update@ZLHIS_DBL(n_病人id, n_主页id);
      b_Message.Zlhis_Patient_011@ZLHIS_DBL(n_病人id, n_主页id, babyno_In);
      delete from his_病人新生儿记录;
    else  ----修改
      insert into   his_病人新生儿记录
      (病人id,主页id,序号,婴儿姓名,婴儿性别,分娩次数,分娩方式,胎儿状况,出生时间,身长,体重,血型,备注说明,死亡时间,登记时间,登记人)
      select 病人id,住院次数,序号,婴儿姓名,婴儿性别,分娩次数,娩出方式,胎儿状况,出生时间,身长,体重,血型,备注说明,死亡时间,登记时间,登记人
       from   ( select  d.pid 病人id,d.tid 住院次数,b.Babyno 序号,d.name||decode(b.Sex,'男','之子','之女')||t.顺序 as 婴儿姓名,
              b.Sex as 婴儿性别,c.妊次 as 分娩次数,a.娩出方式,b.胎儿状况,b.身长,b.体重,b.血型,to_date(b.boutt,'yyyy-mm-dd hh24:mi:ss') as 出生时间,
              b.死亡时间,'' as 备注说明,b.Recorder 登记人,b.Addtime 登记时间
              from V_SOL_INF_DELIVERY a,v_newborn b,v_sol_rs_birth c,sol_inf_puerpera d,
              (select mid,sex,babyno,row_number() over(partition by mid,sex order by babyno) 顺序 from SOL_INF_NEWBORNS t  ) t
              where a.Mid=b.Mid and a.mid=d.mid and a.mid=t.mid and b.Babyno=t.babyno
              and b.Mid=c.mid(+) and a.mid=mid_In
                ) where 序号=babyno_In;
        delete from 病人新生儿记录@ZLHIS_DBL where 病人id=n_病人id and 主页id=n_主页id and 序号=babyno_In;
        insert into 病人新生儿记录@ZLHIS_DBL(病人id,主页id,序号,婴儿姓名,婴儿性别,分娩次数,分娩方式,胎儿状况,出生时间,身长,体重,血型,备注说明,死亡时间,登记时间,登记人)
        select 病人id,主页id,序号,婴儿姓名,婴儿性别,分娩次数,分娩方式,胎儿状况,出生时间,身长,体重,血型,备注说明,死亡时间,登记时间,登记人 from his_病人新生儿记录 ;
        Zl_病区自动标记_Update@ZLHIS_DBL(n_病人id, n_主页id);
      b_Message.Zlhis_Patient_011@ZLHIS_DBL(n_病人id, n_主页id, babyno_In);
     delete from his_病人新生儿记录;
      end if;
     --新生儿登记删除
   elsif state_in=3 then
     delete from 病人新生儿记录@ZLHIS_DBL where 病人id=n_病人id and 主页id=n_主页id and 序号=babyno_In;
     Zl_病区自动标记_Update@ZLHIS_DBL(n_病人id,n_主页id);

     b_Message.ZLHIS_PATIENT_013@ZLHIS_DBL(n_病人id,n_主页id,babyno_In);
  End If;
End his_病人新生儿登记_revise;
/
CREATE OR REPLACE Function ZLSOL.f_List2str
( 
  p_Strlist   In t_Strlist, 
  p_Delimiter In Varchar2 Default ',', 
  p_Distinct  In Number Default 1, 
  p_Maxlength In Number Default 4000 
) Return Varchar2 Is 
  l_String Long; 
  l_Add    Number; 
  --功能：将一个列表集合转换为一个缺省以逗号分隔的字符串。 
  --例： 
  --Select 科室, f_List2str(Cast(Collect(人员 Order By 编号) As t_Strlist)) 人员列表 
  --From (Select a.名称 As 科室, c.姓名 As 人员,c.编号 
  --      From 部门表 A, 部门人员 B, 人员表 C 
  --      Where a.Id = b.部门id And b.人员id = c.Id 
  --      Order By 科室, 人员) 
  --Group By 科室 
 
  --此函数不支持with方式构造的临时内存表，这将会报错：ORA-00932: 数据类型不一致: 应为 -, 但却获得 -。 
  --例如：With Test As (Select '内科' As 科室,'张三' As 人员 From Dual Union All......) 
  --     Select 科室,f_List2str(cast(COLLECT(人员) as t_Strlist)) tt From Test Group By 科室 
Begin 
  If p_Strlist.Count > 0 Then 
    For I In p_Strlist.First .. p_Strlist.Last Loop 
      l_Add := 0; 
      If p_Distinct = 1 Then 
        If Instr(',' || l_String || ',', ',' || p_Strlist(I) || ',') = 0 Then 
          l_Add := 1; 
        End If; 
      Else 
        l_Add := 1; 
      End If; 
      If l_Add = 1 Then 
        If I != p_Strlist.First Then 
          l_String := l_String || p_Delimiter; 
        End If; 
        l_String := l_String || p_Strlist(I); 
        If Lengthb(l_String) > p_Maxlength Then 
          l_String := Substr(l_String, 1, p_Maxlength); 
          Return l_String; 
        End If; 
      End If; 
    End Loop; 
  End If; 
  Return l_String; 
End f_List2str;
/
CREATE OR REPLACE TYPE ZLSOL."T_STRLIST" as Table of Varchar2(4000)
/
--用户同步所需脚本
create table ZLSOL.sol_user
(
code varchar2(20),
name varchar2(20),
state number(1)
);
alter table ZLSOL.sol_user add constraint sol_user_code primary key (code) Using Index Tablespace ZLSOL_DATA;

Create Or Replace Trigger t_Apex_User
  After Insert Or Delete Or Update On Sol_User
  For Each Row
Declare
  n_Group_Id Number(18);
  n_User_Id  Number(18);
  ----新增人员；修改人员；删除人员；启用人员；停用人员
Begin
  n_Group_Id := Apex_Util.Find_Security_Group_Id(p_Workspace => 'ZLSOL');
  If Inserting Then
    --新增人员
    Apex_Util.Set_Security_Group_Id(p_Security_Group_Id => n_Group_Id);
    Apex_Util.Create_User(p_User_Name => :New.Code, p_First_Name => Substr(:New.Name, 2),
                          p_Last_Name => Substr(:New.Name, 1, 1), p_Web_Password => '123',
                          p_Change_Password_On_First_Use => 'N');
  Elsif Deleting Then
    --删除人员
    Apex_Util.Set_Security_Group_Id(p_Security_Group_Id => n_Group_Id);
    Apex_Util.Remove_User(p_User_Name => :Old.Code);
  Elsif Updating Then
    --修改人员姓名
    If :New.Name <> :Old.Name Then
      Apex_Util.Set_Security_Group_Id(p_Security_Group_Id => n_Group_Id);
      n_User_Id := Apex_Util.Get_User_Id(p_Username => :Old.Code);
      Apex_Util.Set_First_Name(p_Userid => n_User_Id, p_First_Name => Substr(:New.Name, 2));
      Apex_Util.Set_Last_Name(p_Userid => n_User_Id, p_Last_Name => Substr(:New.Name, 1, 1));
    End If;
    --启用、停用人员对应启用和锁定账户
    If :New.State <> :Old.State Then
      Apex_Util.Set_Security_Group_Id(p_Security_Group_Id => n_Group_Id);
      If :New.State = 0 Then
        Apex_Util.Lock_Account(p_User_Name => :New.Code);
      Else
        Apex_Util.Unlock_Account(p_User_Name => :New.Code);
      End If;
    End If;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End t_Apex_User;
/
--基础数据
create table ZLSOL.SOL_STD_FetalPosition--胎方位
(
code varchar2(10),
name varchar2(50),
Description varchar2(100)
);
create table ZLSOL.SOL_STD_Delivery--分娩方式
(
code varchar2(10),
name varchar2(50),
Description varchar2(100)
);
create table ZLSOL.SOL_STD_PerinealLaceration--会阴裂伤情况
(
code varchar2(10),
name varchar2(50),
Description varchar2(100)
);
create table ZLSOL.SOL_STD_Anesthesia--麻醉方式
(
code varchar2(10),
name varchar2(50),
Description varchar2(500)
);
create table ZLSOL.SOL_STD_FetalPresentation--胎先露
(
code varchar2(10),
name varchar2(50),
Description varchar2(100)
);
create table ZLSOL.SOL_STD_NeonatalAbnormality--新生儿异常情况
(
code varchar2(10),
name varchar2(50),
Description varchar2(100)
);

--胎方位
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('01','左枕前(LOA)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('02','右枕前(ROA)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('03','左枕后(LOP)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('04','右枕后(ROP)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('05','左枕横(LOT)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('06','右枕横(ROT)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('07','左颏前(LMA)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('08','右颏前(RMA)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('09','左颏后(LMP)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('10','右颏后(RMP)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('11','左颏横(LMT)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('12','右颏横(RMT)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('13','左骶前(LSA)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('14','右骶前(RSA)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('15','左骶后(LSP)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('16','右骶后(RSP)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('17','左骶横(LST)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('18','右骶横(RST)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('19','左肩前(LScA)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('20','右肩前(RscA)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('21','左肩后(LScP)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('22','右肩后(RScP)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('99','不祥','');
--分娩方式
Insert Into ZLSOL.SOL_STD_Delivery(code,name,Description) Values('1','阴道自然分娩','');
Insert Into ZLSOL.SOL_STD_Delivery(code,name,Description) Values('11','会阴切开','');
Insert Into ZLSOL.SOL_STD_Delivery(code,name,Description) Values('12','会阴未切','');
Insert Into ZLSOL.SOL_STD_Delivery(code,name,Description) Values('2','阴道手术助产','');
Insert Into ZLSOL.SOL_STD_Delivery(code,name,Description) Values('21','产钳助产','');
Insert Into ZLSOL.SOL_STD_Delivery(code,name,Description) Values('22','臀位助产','');
Insert Into ZLSOL.SOL_STD_Delivery(code,name,Description) Values('23','胎头吸引','');
Insert Into ZLSOL.SOL_STD_Delivery(code,name,Description) Values('3','剖宫产','');
Insert Into ZLSOL.SOL_STD_Delivery(code,name,Description) Values('31','子宫下段横切口剖宫产','');
Insert Into ZLSOL.SOL_STD_Delivery(code,name,Description) Values('32','子宫体剖宫产','');
Insert Into ZLSOL.SOL_STD_Delivery(code,name,Description) Values('33','腹膜外剖宫产','');
Insert Into ZLSOL.SOL_STD_Delivery(code,name,Description) Values('9','其他','');
--会阴裂伤情况
Insert Into ZLSOL.SOL_STD_PerinealLaceration(code,name,Description) Values('1','无裂伤','');
Insert Into ZLSOL.SOL_STD_PerinealLaceration(code,name,Description) Values('2','Ⅰ°裂伤','');
Insert Into ZLSOL.SOL_STD_PerinealLaceration(code,name,Description) Values('3','Ⅱ°裂伤','');
Insert Into ZLSOL.SOL_STD_PerinealLaceration(code,name,Description) Values('4','Ⅲ°裂伤','');
Insert Into ZLSOL.SOL_STD_PerinealLaceration(code,name,Description) Values('5','会阴切开','');
--麻醉方式
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('1','全身麻醉','用麻醉剂使全身处于麻醉状态');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('11','吸入麻醉','用吸入麻醉剂的方法使全身处于麻醉状态');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('12','静脉麻醉','经静脉注入麻醉剂使全身处于麻醉状态');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('13','基础麻醉','麻醉前先使患者神志消失的方法');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('2','椎管内麻醉','将麻醉药注入椎管内达到局部麻醉效果的方法');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('21','蛛网膜下腔阻滞麻醉','将麻醉药注入蛛网膜下腔达到局部麻醉效果的方法');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('22','硬脊膜外腔阻滞麻醉','将麻醉药注入硬脊膜外腔产生局部麻醉效果的方法');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('3','局部麻醉','将麻醉药直接注入施行手术的组织内或手术部位周围的麻醉方法');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('31','神经丛阻滞麻醉','将局部麻醉药注射于神经丛附近，使通过神经丛的神经及其所分布的区域产生局部麻醉的方法');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('32','神经节阻滞麻醉','将局部麻醉药注射于神经节附近，使通过神经节的神经及其所分布的区域产生局部麻醉的方法');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('33','神经阻滞麻醉','将局麻药物注射于神经干的周围，使该神经分布的区域产生麻醉作用的方法');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('34','区域阻滞麻醉','将局麻药注射于手术野外周，使通往手术野以及由手术野传出的神经末梢皆受到阻滞的局部麻醉方法');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('35','局部浸润麻醉','将局麻药沿手术切口线分层注入组织内，以阻滞组织中的神经末梢的麻醉方法');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('36','表面麻醉','将麻醉药直接与粘膜或皮肤接触，使支配该部分粘膜或皮肤内的神经末梢被阻滞的麻醉方法');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('4','复合麻醉','用一种以上药物或采用多种麻醉方法以增强麻醉效果');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('41','静吸复合全麻','静脉麻醉和吸入麻醉共同作用产生麻醉效果');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('42','针药复合麻醉','针刺麻醉和药物麻醉共同作用产生麻醉效果');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('43','神经丛与硬膜外阻滞复合麻醉','神经丛阻滞麻醉和硬脊膜外腔阻滞麻醉共同作用产生麻醉效果');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('44','全麻复合全身降温','在全身麻醉的同时主动降低患者血压');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('45','全麻复合控制性降压','在全身麻醉的同时降低患者的体温');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('9','其他麻醉方法','以上未提及的其他麻醉方法');
--胎先露
Insert Into ZLSOL.SOL_STD_FetalPresentation(code,name,Description) Values('1','头先露','');
Insert Into ZLSOL.SOL_STD_FetalPresentation(code,name,Description) Values('2','臀先露','');
Insert Into ZLSOL.SOL_STD_FetalPresentation(code,name,Description) Values('3','肩先露','');
Insert Into ZLSOL.SOL_STD_FetalPresentation(code,name,Description) Values('4','足先露','');
Insert Into ZLSOL.SOL_STD_FetalPresentation(code,name,Description) Values('9','不详','');
--新生儿异常情况
Insert Into ZLSOL.SOL_STD_NeonatalAbnormality(code,name,Description) Values('1','无','');
Insert Into ZLSOL.SOL_STD_NeonatalAbnormality(code,name,Description) Values('2','早期新生儿死亡','');
Insert Into ZLSOL.SOL_STD_NeonatalAbnormality(code,name,Description) Values('3','畸形','');
Insert Into ZLSOL.SOL_STD_NeonatalAbnormality(code,name,Description) Values('4','早产','');
Insert Into ZLSOL.SOL_STD_NeonatalAbnormality(code,name,Description) Values('5','窒息','');
Insert Into ZLSOL.SOL_STD_NeonatalAbnormality(code,name,Description) Values('6','低出生体重','');
Insert Into ZLSOL.SOL_STD_NeonatalAbnormality(code,name,Description) Values('9','其他','');