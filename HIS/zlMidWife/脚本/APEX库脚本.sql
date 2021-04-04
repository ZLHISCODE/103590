--在APEX库执行（修改密码，ip，实例名[SERVICE_NAME]）
create database link ZLHIS_DBL  connect to zlhis identified by zlhis_password  using '(DESCRIPTION =
    (ADDRESS = (PROTOCOL = TCP)(HOST = 192.168.0.60)(PORT = 1521))
    (CONNECT_DATA =
      (SERVICE_NAME = dyey)
    )
  )';
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
  引产方法 VARCHAR2(100)
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

create table ZLSOL.SOL_STD_ANESTHESIA
(
  code        VARCHAR2(10),
  name        VARCHAR2(50),
  description VARCHAR2(500)
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

create table ZLSOL.SOL_STD_DELIVERY
(
  code        VARCHAR2(10),
  name        VARCHAR2(50),
  description VARCHAR2(100)
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

create table ZLSOL.SOL_STD_FETALPOSITION
(
  code        VARCHAR2(10),
  name        VARCHAR2(50),
  description VARCHAR2(100)
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

create table ZLSOL.SOL_STD_FETALPRESENTATION
(
  code        VARCHAR2(10),
  name        VARCHAR2(50),
  description VARCHAR2(100)
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

create table ZLSOL.SOL_STD_NEONATALABNORMALITY
(
  code        VARCHAR2(10),
  name        VARCHAR2(50),
  description VARCHAR2(100)
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

create table ZLSOL.SOL_STD_PERINEALLACERATION
(
  code        VARCHAR2(10),
  name        VARCHAR2(50),
  description VARCHAR2(100)
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

create table ZLSOL.SOL_USERLIST
(
  user_code VARCHAR2(20) not null,
  user_name VARCHAR2(50)
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
alter table ZLSOL.SOL_USERLIST
  add constraint SOL_USERLIST_PK primary key (USER_CODE)
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

create or replace force view ZLSOL.v_delivery as
Select a.Mid, b."隐藏1",b."产程开始时间",b."宫口全开时间",b."胎儿娩出时间",b."胎盘娩出时间",b."第一产程",b."第二产程",b."第三产程",b."宫缩情况",b."出产房宫高脐下距离",b."结扎",b."破膜方式",b."破膜时间",b."羊水性状",b."羊水量",b."羊水颜色",b."胎盘娩出方式",b."胎盘剥离方式",b."胎盘完整度",b."胎盘胎膜残留",b."胎盘体积",b."胎盘形态",b."胎盘大小",b."胎盘重量",b."脐带附着",b."脐带长度",b."脐带绕颈",b."脐带真假结",b."脐带脱垂",b."娩出方式",b."娩出胎方位",b."产瘤大小",b."产瘤部位",b."会阴裂伤程度",b."会阴裂伤切口",b."会阴裂伤缝合",b."会阴裂伤麻醉",b."宫颈裂伤长度",b."宫颈裂伤部位",b."宫颈裂伤状况",b."阴道裂伤部位大小",b."阴道裂伤血肿大小",b."新生儿性别",b."新生儿体重",b."新生儿身长",b."新生儿抢救吸氧",b."新生儿抢救吸出物",b."新生儿抢救吸出物性状",b."新生儿抢救抢救药物",b."新生儿抢救畸形",b."新生儿抢救死胎",b."新生儿抢救死产",b."产后血压",b."产后流血",b."产时用药",b."产后用药",b."特殊情况", d."隐藏3",d."出产房时间",d."护送人",d."接生人",d."记录人"
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
Select a.Mid, b.Bid, b.Babyno, b.Addtime, b.Recorder, a.Name, a.Old, a.Bedno, a.Pno, b.Sex, d."隐藏2",d."身长",d."体重",d."头围",d."胸围",d."一般情况反应",d."一般情况面色",d."一般情况皮肤",d."一般情况毳毛",d."头部变形",d."颅骨重叠",d."胎头水肿血肿",d."胎头水肿大小",d."前囟",d."张力",d."眼神",d."口腔",d."心",d."乳结",d."肝",d."脾",d."四肢",d."外展试验",d."肛门",d."生殖器", e."隐藏3",e."心率1分钟",e."心率5分钟",e."心率10分钟",e."呼吸1分钟",e."呼吸5分钟",e."呼吸10分钟",e."喉反射1分钟",e."喉反射5分钟",e."喉反射10分钟",e."肌张力1分钟",e."肌张力5分钟",e."肌张力10分钟",e."肤色1分钟",e."肤色5分钟",e."肤色10分钟",e."总分1分钟",e."总分5分钟",e."总分10分钟", f."隐藏4",f."出孕期产时合并症及用药情况",f."出生前胎儿情况",f."婴儿出生时抢救情况",f."出生缺陷",f."母乳喂养指导",f."诊断"
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

create or replace force view ZLSOL.v_sol_inf_checkinroom as
Select a.mid, b."入房目的",b."入房时间",b."医疗病历",b."护理病历",b."分娩知情通知书",b."宫缩规律性",b."胎心率",b."胎心次数",b."破膜情况",b."是否有合并症",b."种类",b."输液单",b."静脉通道",b."局部情况",b."特殊药物",b."其他",b."交班者",b."接班者"
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
Select a.mid, b."OUTROOMTIME",b."出房状态",b."医疗病历",b."护理病历",b."静脉通道",b."局部情况",b."会阴裂伤",b."会阴切开术",b."会阴切口缝合",b."会阴水肿",b."会阴血肿",b."产后出血",b."出血量",b."特殊药物",b."交班者",b."接班者",b."药物",b."备注"
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
Select a.Mid, b."隐藏1", b."产程开始时间", b."宫口全开时间", b."胎儿娩出时间", b."胎盘娩出时间", b."第一产程", b."第二产程", b."第三产程", b."宫缩情况",
       b."出产房宫高脐下", b."结扎", b."破膜方式", b."破膜时间", b."羊水性状", b."羊水量", b."羊水颜色", b."胎盘娩出方式", b."胎盘剥离方式", b."胎盘完整度",
       b."胎盘胎膜残留", b."胎盘体积", b."胎盘形态", b."胎盘大小", b."胎盘重量", b."脐带附着", b."脐带长度", b."脐带真假结", b."脐带",b."绕颈周数", b."娩出方式",
       b."娩出胎方位", b."产瘤大小", b."产瘤部位", b."会阴裂伤程度", b."会阴裂伤切口", b."会阴裂伤缝合", b."会阴裂伤麻醉", b."宫颈裂伤长度", b."宫颈裂伤部位", b."宫颈裂伤状况",
       b."阴道裂伤部位大小", b."阴道裂伤血肿大小", b."新生儿抢救吸氧", b."新生儿抢救吸出物", b."新生儿抢救吸出物性状",
       b."新生儿抢救抢救药物", b."新生儿抢救畸形", b."新生儿抢救死胎", b."新生儿抢救死产",b."母婴早接触早吸吮时间", b."产后血压", b."产后流血", b."产时用药", b."产后用药",
       b."产后诊断", b."特殊情况", d."隐藏3",d."出产房时间", d."护送人", d."接生人", d."记录人"
From Sol_Inf_Delivery A,
     Json_Table(Nvl(a.Newborndetail, '{"隐藏1":"1"}'),
                 '$'
                  Columns(隐藏1 Varchar2(50) Path '$.隐藏1', 产程开始时间 Varchar2(19) Path '$.产程开始时间',
                          宫口全开时间 Varchar2(19) Path '$.宫口全开时间', 胎儿娩出时间 Varchar2(19) Path '$.胎儿娩出时间',
                          胎盘娩出时间 Varchar2(19) Path '$.胎盘娩出时间', 第一产程 Varchar2(50) Path '$.第一产程',
                          第二产程 Varchar2(50) Path '$.第二产程', 第三产程 Varchar2(50) Path '$.第三产程',
                          出产房宫高脐下 Varchar2(50) Path '$.出产房宫高脐下', 结扎 Varchar2(50) Path '$.结扎',
                          破膜方式 Varchar2(50) Path '$.破膜方式', 破膜时间 Varchar2(19) Path '$.破膜时间', 羊水性状 Varchar2(50) Path '$.羊水性状',
                          羊水量 Varchar2(50) Path '$.羊水量', 羊水颜色 Varchar2(50) Path '$.羊水颜色',
                          胎盘娩出方式 Varchar2(50) Path '$.胎盘娩出方式', 胎盘剥离方式 Varchar2(50) Path '$.胎盘剥离方式',
                          胎盘完整度 Varchar2(50) Path '$.胎盘完整度', 胎盘胎膜残留 Varchar2(50) Path '$.胎盘胎膜残留',
                          胎盘体积 Varchar2(50) Path '$.胎盘体积', 胎盘形态 Varchar2(50) Path '$.胎盘形态', 胎盘大小 Varchar2(50) Path '$.胎盘大小',
                          胎盘重量 Varchar2(50) Path '$.胎盘重量', 脐带附着 Varchar2(50) Path '$.脐带附着', 脐带长度 Varchar2(50) Path '$.脐带长度',
                         脐带真假结 Varchar2(50) Path '$.脐带真假结',脐带 Varchar2(50) Path '$.脐带', 绕颈周数 Varchar2(50) Path '$.绕颈周数',
                          娩出方式 Varchar2(50) Path '$.娩出方式',
                          娩出胎方位 Varchar2(50) Path '$.娩出胎方位', 产瘤大小 Varchar2(50) Path '$.产瘤大小',
                          产瘤部位 Varchar2(50) Path '$.产瘤部位', 会阴裂伤程度 Varchar2(50) Path '$.会阴裂伤程度',
                          会阴裂伤切口 Varchar2(50) Path '$.会阴裂伤切口', 会阴裂伤缝合 Varchar2(50) Path '$.会阴裂伤缝合',
                          会阴裂伤麻醉 Varchar2(50) Path '$.会阴裂伤麻醉', 宫颈裂伤长度 Varchar2(50) Path '$.宫颈裂伤长度',
                          宫颈裂伤部位 Varchar2(50) Path '$.宫颈裂伤部位', 宫颈裂伤状况 Varchar2(50) Path '$.宫颈裂伤状况',
                          阴道裂伤部位大小 Varchar2(50) Path '$.阴道裂伤部位大小', 阴道裂伤血肿大小 Varchar2(50) Path '$.阴道裂伤血肿大小',
                          新生儿抢救吸氧 Varchar2(50) Path '$.新生儿抢救吸氧', 新生儿抢救吸出物 Varchar2(50) Path '$.新生儿抢救吸出物',
                          新生儿抢救吸出物性状 Varchar2(50) Path '$.新生儿抢救吸出物性状', 新生儿抢救抢救药物 Varchar2(50) Path '$.新生儿抢救抢救药物',
                          新生儿抢救畸形 Varchar2(50) Path '$.新生儿抢救畸形', 新生儿抢救死胎 Varchar2(50) Path '$.新生儿抢救死胎',
                          新生儿抢救死产 Varchar2(50) Path '$.新生儿抢救死产',宫缩情况 Varchar2(50) Path '$.宫缩情况', 母婴早接触早吸吮时间 Varchar2(50) Path '$.母婴早接触早吸吮时间',
                          产后血压 Varchar2(50) Path '$.产后血压',
                          产后流血 Varchar2(50) Path '$.产后流血', 产时用药 Varchar2(50) Path '$.产时用药', 产后用药 Varchar2(50) Path '$.产后用药',
                          产后诊断 Varchar2(50) Path '$.产后诊断',特殊情况 Varchar2(50) Path '$.特殊情况')) As B,
     Json_Table(Nvl(a.Deliveryinf, '{"隐藏3":"1"}'),
                 '$' Columns(隐藏3 Varchar2(1) Path '$.隐藏3', 出产房时间 Varchar2(50) Path '$.出产房时间',
                          护送人 Varchar2(50) Path '$.护送人', 接生人 Varchar2(50) Path '$.接生人', 记录人 Varchar2(50) Path '$.记录人')) As D;

create or replace force view ZLSOL.v_sol_inf_equipment as
Select a.mid, b."侧切剪产前",b."侧切剪术中",b."侧切剪产后",b."脐带剪产前",b."脐带剪术中",b."脐带剪产后",b."止血钳产前",b."止血钳术中",b."止血钳产后",b."牙镊产前",b."牙镊术中",b."牙镊产后",b."持针器产前",b."持针器术中",b."持针器产后",b."穿刺针产前",b."穿刺针术中",b."穿刺针产后",b."洗耳球产前",b."洗耳球术中",b."洗耳球产后",b."缝合针产前",b."缝合针术中",b."缝合针产后",b."拉钩产前",b."拉钩术中",b."拉钩产后",b."宫颈钳产前",b."宫颈钳术中",b."宫颈钳产后",b."窥器产前",b."窥器术中",b."窥器产后",b."刮匙产前",b."刮匙术中",b."刮匙产后",b."艾利斯产前",b."艾利斯术中",b."艾利斯产后",b."产前产前",b."产前术中",b."产前产后",b."纱布产前",b."纱布术中",b."纱布产后",b."卵圆钳产前",b."卵圆钳术中",b."卵圆钳产后"
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
产前产前   Number(2) PATH '$.产前产前',
产前术中   Number(2) PATH '$.产前术中',
产前产后   Number(2) PATH '$.产前产后',
纱布产前   Number(2) PATH '$.纱布产前',
纱布术中   Number(2) PATH '$.纱布术中',
纱布产后   Number(2) PATH '$.纱布产后',
卵圆钳产前   Number(2) PATH '$.卵圆钳产前',
卵圆钳术中   Number(2) PATH '$.卵圆钳术中',
卵圆钳产后   Number(2) PATH '$.卵圆钳产后'
)) as b;

create or replace force view ZLSOL.v_sol_inf_newborns as
Select a.Mid, b.Bid, b.Babyno, b.Addtime, b.Recorder, a.Name, a.Old, a.Bedno, a.Pno, b.Sex, d."隐藏2",d."身长",d."体重",d."头围",d."胸围",d."一般情况反应",d."一般情况面色",d."一般情况皮肤",d."一般情况毳毛",d."头部变形",d."颅骨重叠",d."胎头水肿血肿",d."胎头水肿大小",d."前囟",d."张力",d."眼神",d."口腔",d."心",d."乳结",d."肝",d."脾",d."四肢",d."外展试验",d."肛门",d."生殖器", e."隐藏3",e."心率1分钟",e."心率5分钟",e."心率10分钟",e."呼吸1分钟",e."呼吸5分钟",e."呼吸10分钟",e."喉反射1分钟",e."喉反射5分钟",e."喉反射10分钟",e."肌张力1分钟",e."肌张力5分钟",e."肌张力10分钟",e."肤色1分钟",e."肤色5分钟",e."肤色10分钟",e."总分1分钟",e."总分5分钟",e."总分10分钟", f."隐藏4",f."出孕期产时合并症及用药情况",f."出生前胎儿情况",f."婴儿出生时抢救情况",f."出生缺陷",f."母乳喂养指导",f."诊断"
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
       Decode(Checkoutroom, 1, '√', '') 出房,Decode(Equipment, 1, '√', '') 器械,outtime,pid,tid
From Sol_Inf_Puerpera;

create or replace force view ZLSOL.v_sol_rs_birth as
Select a.mid,b."妊次",b."产次",b."血型",b."既往妊娠史",b."末次月经",b."预产期",b."髂前上棘间径",b."髂嵴间径",b."坐骨结节间径",b."骶耻外径",b."骶骨弧度",b."骶骨关节",b."坐骨切迹",b."坐骨髂",b."并发症",b."产前记录特征",b."检查时间",b."收缩压",b."舒张压",b."体温",b."脉博",b."胎心率",b."胎儿大小",b."宫缩规律性",b."胎位",b."破膜情况",b."先露",b."宫口",b."检查者",b."宫缩开始时间",b."破膜时间",b."入院处理"
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
脉博            Varchar2(10) PATH '$.脉博',
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
Select  a.courseid,a.mid,b."检查时间",b."是否剖宫产",b."胎方位",b."收缩压",b."舒张压",b."体温",b."脉博",b."胎心率",b."宫缩强度",b."宫缩持续",b."宫缩间隔",b."宫颈厚薄",b."宫口",b."破膜情况",b."先露",b."处理",b."检查者"
From SOL_RS_BIRTH_COURSE a,JSON_TABLE(a.CONTENT,'$' columns(
检查时间        Varchar2(20)  PATH '$.检查时间',
是否剖宫产    Varchar2(20)  PATH '$.是否剖宫产',
胎方位 Varchar2(20)  PATH '$.胎方位',
收缩压        Number(3) PATH '$.收缩压',
舒张压        Number(3) PATH '$.舒张压',
体温        Number(4,2)  PATH '$.体温',
脉博        Varchar2(10) PATH '$.脉博',
胎心率        Varchar2(10) PATH '$.胎心率',
宫缩强度    Varchar2(10) PATH '$.宫缩强度',
宫缩持续  Varchar2(10) PATH '$.宫缩持续',
宫缩间隔  Varchar2(10) PATH '$.宫缩间隔',
宫颈厚薄    Varchar2(10) PATH '$.宫颈厚薄',
宫口        Number(4,2) PATH '$.宫口',
破膜情况    Varchar2(10) PATH '$.破膜情况',
先露        Number(2) PATH '$.先露'，
处理        Varchar2(50) PATH '$.处理'，
检查者      Varchar2(50) PATH '$.检查者'
)) as b;

create or replace force view ZLSOL.v_sol_rs_druglabor as
Select Mid, To_Char(日期, 'YYYY-MM-DD HH24:MI') 日期, 引产指征, 引产方法 from Sol_Rs_Druglabor;

create or replace force view ZLSOL.v_sol_rs_druglabor_list as
Select a.Mid, a.Courseid ID, b."记录时间",b."收缩压",b."舒张压",b."脉搏",b."胎心率",b."宫缩强度",b."宫缩持续",b."宫缩间隔",b."宫口",b."先露",b."羊水量",b."羊水性状",b."处理",b."记录人",b."剂量",b."滴速"
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
Select a.mid,a.courseid,b."记录时间",b."胎方位",b."收缩压",b."舒张压",b."宫高",b."腹围",b."胎动计数早",b."胎动计数中",b."胎动计数晚",b."胎心率",b."先露",b."宫口",b."破膜情况",b."羊水性状",b."宫缩强度",b."宫缩持续",b."宫缩间隔",b."处理",b."检查者"
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
胎心率 Number(3) PATH '$.胎心率',
先露     Varchar2(20) PATH '$.先露',
宫口     Varchar2(20) PATH '$.宫口',
破膜情况     Varchar2(20) PATH '$.破膜情况',
羊水性状      Varchar2(20) PATH '$.羊水性状',
宫缩强度     Varchar2(20) PATH '$.宫缩强度',
宫缩持续       Varchar2(20) PATH '$.宫缩持续',
宫缩间隔       Varchar2(20) PATH '$.宫缩间隔',
处理     Varchar2(100) PATH '$.处理',
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

CREATE OR REPLACE Procedure his_病人新生儿登记_revise
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
   /* update  病人新生儿记录@ZLHIS_DBL set
   (序号,婴儿姓名,婴儿性别,分娩次数,分娩方式,胎儿状况,出生时间,身长,体重,血型,备注说明,死亡时间,登记时间,登记人)=
   (select 序号,婴儿姓名,婴儿性别,分娩次数,分娩方式,胎儿状况,出生时间,身长,体重,血型,备注说明,死亡时间,登记时间,登记人
    from   his_病人新生儿记录) ;*/
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
CREATE OR REPLACE Function Getsdate
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
End Getsdate;
/
CREATE OR REPLACE Function Changejsonvalue
(
  Content_In In Varchar2,
  Fild_In    In Varchar2,
  Value_In   In Varchar2
) Return Varchar2 Is
  v_Content Varchar2(4000);
  v_Value   Varchar2(4000);
  n_Locate  Number(5);
Begin
  If Fild_In Is Null Or Content_In Is Null Then
    Return Content_In;
  End If;
  If Instr(Content_In, '"' || Fild_In || '"') = 0 Then
    v_Content := Substr(Content_In, 2);
    v_Content := '{"' || Fild_In || '":' || '"' || Value_In || '",' || v_Content;
    Return v_Content;
  End If;
  n_Locate := Instr(Content_In, '"' || Fild_In || '"');
  v_Value  := Substr(Content_In, n_Locate);
  n_Locate := Instr(v_Value, '"', 1, 4);
  v_Value  := Substr(v_Value, 1, n_Locate);

  v_Content := Replace(Content_In, v_Value, '"' || Fild_In || '":"' || Value_In || '"');
  Return v_Content;
End Changejsonvalue;

CREATE OR REPLACE Function Zl_Sol_Getdata
(
  Mid_In       In Number,
  Condition_In In Varchar2
) Return Varchar2 Is
  v_Return Varchar2(1000);
  n_Pid    Number(18);
  n_Tid    Number(18);
Begin
  Select Pid, Tid Into n_Pid, n_Tid From Sol_Inf_Puerpera Where Mid = Mid_In;
  Select max(a.记录内容)
  Into v_Return
  From (Select 病人id, 主页id, 项目名称, 发生时间, 记录内容
         From 病人护理明细@ZLHIS_DBL a, 病人护理数据@ZLHIS_DBL b, 病人护理文件@ZLHIS_DBL c
         Where a.记录id = b.Id And b.文件id = c.Id And c.病人id = n_Pid And c.主页id = n_Tid And c.婴儿 = 0 And c.文件名称 = '产科体温单' And
               a.项目名称 = Condition_In) a,
       (Select 病人id, 主页id, 项目名称, Min(b.发生时间) 第一次时间
         From 病人护理明细@ZLHIS_DBL a, 病人护理数据@ZLHIS_DBL b, 病人护理文件@ZLHIS_DBL c
         Where a.记录id = b.Id And b.文件id = c.Id And c.病人id = n_Pid And c.主页id = n_Tid And c.婴儿 = 0 And c.文件名称 = '产科体温单' And
               a.项目名称 = Condition_In
         Group By 病人id, 主页id, 项目名称) b
  Where a.项目名称 = b.项目名称 And a.发生时间 = b.第一次时间 And a.病人id = b.病人id And a.主页id = b.主页id;
  Return v_Return;
Exception
  When Others Then
    Null;
End Zl_Sol_Getdata;