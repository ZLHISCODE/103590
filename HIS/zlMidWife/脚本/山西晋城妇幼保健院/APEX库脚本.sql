--在APEX库执行（修改zlsol的密码，ip，实例名[SERVICE_NAME]）
create database link To_His  connect to ZLHIS identified by his  using '(DESCRIPTION =
    (ADDRESS = (PROTOCOL = TCP)(HOST = 192.168.0.60)(PORT = 1521))
    (CONNECT_DATA =
      (SERVICE_NAME = orcl)
    )
  )';
create table ZLSOL_JC.HIS_病人新生儿记录
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
;
create table ZLSOL_JC.SOL_INF_PUERPERA
(
  mid          NUMBER(18) generated always as identity,
  pid          NUMBER(18),
  tid          NUMBER(5),
  name         VARCHAR2(50),
  old          VARCHAR2(20),
  bedno        VARCHAR2(16),
  pno          NUMBER(18),
  diagnosis    VARCHAR2(100),
  status       NUMBER(1),
  outroomtime  DATE,
  outtime      DATE,
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
;
comment on table ZLSOL_JC.SOL_INF_PUERPERA
  is '产妇信息';
create index ZLSOL_JC.SOL_INF_PUERPERA_IX_OUTROOMTIME on ZLSOL_JC.SOL_INF_PUERPERA (OUTROOMTIME);
create index ZLSOL_JC.SOL_INF_PUERPERA_IX_OUTTIME on ZLSOL_JC.SOL_INF_PUERPERA (OUTTIME);
create index ZLSOL_JC.SOL_INF_PUERPERA_IX_STATUS on ZLSOL_JC.SOL_INF_PUERPERA (STATUS);
alter table ZLSOL_JC.SOL_INF_PUERPERA
  add constraint SOL_INF_PUERPERA_PK primary key (MID);

create table ZLSOL_JC.SOL_INF_CHECKINROOM
(
  mid      NUMBER(18) not null,
  content  CLOB,
  recorder VARCHAR2(50),
  addtime  DATE
)
;
comment on table ZLSOL_JC.SOL_INF_CHECKINROOM
  is '入房信息';
alter table ZLSOL_JC.SOL_INF_CHECKINROOM
  add constraint SOL_INF_CHECKINROOM_PK primary key (MID);
alter table ZLSOL_JC.SOL_INF_CHECKINROOM
  add constraint SOL_INF_CHECKINROOM_FK_MID foreign key (MID)
  references ZLSOL_JC.SOL_INF_PUERPERA (MID);
alter table ZLSOL_JC.SOL_INF_CHECKINROOM
  add check (Content IS JSON);

create table ZLSOL_JC.SOL_INF_CHECKOUTROOM
(
  mid      NUMBER(18) not null,
  content  CLOB,
  recorder VARCHAR2(50),
  addtime  DATE
)
;
comment on table ZLSOL_JC.SOL_INF_CHECKOUTROOM
  is '出房信息';
alter table ZLSOL_JC.SOL_INF_CHECKOUTROOM
  add constraint SOL_INF_CHECKOUTROOM_PK primary key (MID);
alter table ZLSOL_JC.SOL_INF_CHECKOUTROOM
  add constraint SOL_INF_CHECKOUTROOM_FK_MID foreign key (MID)
  references ZLSOL_JC.SOL_INF_PUERPERA (MID);
alter table ZLSOL_JC.SOL_INF_CHECKOUTROOM
  add check (Content IS JSON);

create table ZLSOL_JC.SOL_INF_DELIVERY
(
  mid           NUMBER(18) not null,
  deliveryinf   CLOB,
  newborndetail CLOB,
  newbornscore  CLOB,
  otherinf      CLOB
)
;
comment on table ZLSOL_JC.SOL_INF_DELIVERY
  is '分娩信息';
alter table ZLSOL_JC.SOL_INF_DELIVERY
  add constraint SOL_INF_DELIVERY_PK primary key (MID);
alter table ZLSOL_JC.SOL_INF_DELIVERY
  add constraint SOL_INF_DELIVERY_FK_MID foreign key (MID)
  references ZLSOL_JC.SOL_INF_PUERPERA (MID);
alter table ZLSOL_JC.SOL_INF_DELIVERY
  add check (DELIVERYINF IS JSON);
alter table ZLSOL_JC.SOL_INF_DELIVERY
  add check (NEWBORNDETAIL IS JSON);
alter table ZLSOL_JC.SOL_INF_DELIVERY
  add check (NEWBORNSCORE IS JSON);
alter table ZLSOL_JC.SOL_INF_DELIVERY
  add check (OTHERINF IS JSON);

create table ZLSOL_JC.SOL_INF_EQUIPMENT
(
  mid        NUMBER(18) not null,
  content    CLOB,
  recordtime DATE,
  deliver    VARCHAR2(50),
  inspector  VARCHAR2(50),
  aftertime  DATE
)
;
comment on table ZLSOL_JC.SOL_INF_EQUIPMENT
  is '器械清点记录';
alter table ZLSOL_JC.SOL_INF_EQUIPMENT
  add constraint SOL_INF_EQUIPMENT_PK primary key (MID);
alter table ZLSOL_JC.SOL_INF_EQUIPMENT
  add constraint SOL_INF_EQUIPMENT_FK_MID foreign key (MID)
  references ZLSOL_JC.SOL_INF_PUERPERA (MID);

create table ZLSOL_JC.SOL_INF_NEWBORNS
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
;
comment on table ZLSOL_JC.SOL_INF_NEWBORNS
  is '新生儿信息';
alter table ZLSOL_JC.SOL_INF_NEWBORNS
  add constraint SOL_INF_NEWBORNS_PK primary key (BID);
alter table ZLSOL_JC.SOL_INF_NEWBORNS
  add constraint SOL_INF_NEWBORNS_UQ unique (MID, BABYNO);
alter table ZLSOL_JC.SOL_INF_NEWBORNS
  add constraint SOL_INF_NEWBORNS_FK_MID foreign key (MID)
  references ZLSOL_JC.SOL_INF_PUERPERA (MID);
alter table ZLSOL_JC.SOL_INF_NEWBORNS
  add check (NEWBORNINF IS JSON);
alter table ZLSOL_JC.SOL_INF_NEWBORNS
  add check (NEWBORNSCORE IS JSON);
alter table ZLSOL_JC.SOL_INF_NEWBORNS
  add check (OTHERINF IS JSON);

create table ZLSOL_JC.SOL_RS_BIRTH
(
  mid     NUMBER(18) not null,
  content CLOB
)
;
comment on table ZLSOL_JC.SOL_RS_BIRTH
  is '产前检查信息';
alter table ZLSOL_JC.SOL_RS_BIRTH
  add constraint SOL_RS_BIRTH_PK primary key (MID);
alter table ZLSOL_JC.SOL_RS_BIRTH
  add constraint SOL_RS_BIRTH_FK_MID foreign key (MID)
  references ZLSOL_JC.SOL_INF_PUERPERA (MID);
alter table ZLSOL_JC.SOL_RS_BIRTH
  add check (Content IS JSON);

create table ZLSOL_JC.SOL_RS_BIRTH_COURSE
(
  courseid NUMBER(18) generated always as identity,
  mid      NUMBER(18),
  content  CLOB
)
;
comment on table ZLSOL_JC.SOL_RS_BIRTH_COURSE
  is '产程经过';
alter table ZLSOL_JC.SOL_RS_BIRTH_COURSE
  add constraint SOL_RS_BIRTH_COURSE_PK primary key (COURSEID);
alter table ZLSOL_JC.SOL_RS_BIRTH_COURSE
  add constraint SOL_RS_BIRTH_COURSE_FK_MID foreign key (MID)
  references ZLSOL_JC.SOL_INF_PUERPERA (MID);
alter table ZLSOL_JC.SOL_RS_BIRTH_COURSE
  add check (CONTENT IS JSON);

create table ZLSOL_JC.SOL_RS_DRUGLABOR
(
  mid  NUMBER(18) not null,
  日期   DATE,
  引产指征 VARCHAR2(50),
  引产方法 VARCHAR2(50)
)
;
comment on table ZLSOL_JC.SOL_RS_DRUGLABOR
  is '药物引产信息';
alter table ZLSOL_JC.SOL_RS_DRUGLABOR
  add constraint SOL_RS_DRUGLABOR_PK primary key (MID);
alter table ZLSOL_JC.SOL_RS_DRUGLABOR
  add constraint SOL_RS_DRUGLABOR_FK_MID foreign key (MID)
  references ZLSOL_JC.SOL_INF_PUERPERA (MID);

create table ZLSOL_JC.SOL_RS_DRUGLABOR_LIST
(
  courseid NUMBER(18) generated always as identity,
  mid      NUMBER(18),
  content  CLOB
)
;
comment on table ZLSOL_JC.SOL_RS_DRUGLABOR_LIST
  is '药物引产记录';
alter table ZLSOL_JC.SOL_RS_DRUGLABOR_LIST
  add constraint SOL_RS_DRUGLABOR_LIST_PK primary key (COURSEID);
alter table ZLSOL_JC.SOL_RS_DRUGLABOR_LIST
  add constraint SOL_RS_DRUGLABOR_LIST_FK_MID foreign key (MID)
  references ZLSOL_JC.SOL_INF_PUERPERA (MID);
alter table ZLSOL_JC.SOL_RS_DRUGLABOR_LIST
  add check (CONTENT IS JSON);

create table ZLSOL_JC.SOL_RS_EXPECTANT
(
  courseid NUMBER(18) generated always as identity,
  mid      NUMBER(18),
  content  CLOB
)
;
comment on table ZLSOL_JC.SOL_RS_EXPECTANT
  is '待产记录';
alter table ZLSOL_JC.SOL_RS_EXPECTANT
  add constraint SOL_RS_EXPECTANT_PK primary key (COURSEID);
alter table ZLSOL_JC.SOL_RS_EXPECTANT
  add constraint SOL_RS_EXPECTANT_FK_MID foreign key (MID)
  references ZLSOL_JC.SOL_INF_PUERPERA (MID);
alter table ZLSOL_JC.SOL_RS_EXPECTANT
  add check (CONTENT IS JSON);

create table ZLSOL_JC.SOL_RS_POSTPARTUM
(
  mid     NUMBER(18) not null,
  content CLOB
)
;
comment on table ZLSOL_JC.SOL_RS_POSTPARTUM
  is '产后观察信息';
alter table ZLSOL_JC.SOL_RS_POSTPARTUM
  add constraint SOL_RS_POSTPARTUM_PK primary key (MID);
alter table ZLSOL_JC.SOL_RS_POSTPARTUM
  add constraint SOL_RS_POSTPARTUM_FK_MID foreign key (MID)
  references ZLSOL_JC.SOL_INF_PUERPERA (MID);
alter table ZLSOL_JC.SOL_RS_POSTPARTUM
  add check (CONTENT IS JSON);

create table ZLSOL_JC.SOL_RS_POSTPARTUM_LIST
(
  courseid NUMBER(18) generated always as identity,
  mid      NUMBER(18),
  content  CLOB
)
;
comment on table ZLSOL_JC.SOL_RS_POSTPARTUM_LIST
  is '产后观察记录';
alter table ZLSOL_JC.SOL_RS_POSTPARTUM_LIST
  add constraint SOL_RS_POSTPARTUM_LIST_PK primary key (COURSEID);
alter table ZLSOL_JC.SOL_RS_POSTPARTUM_LIST
  add constraint SOL_RS_POSTPARTUM_LIST_FK_MID foreign key (MID)
  references ZLSOL_JC.SOL_INF_PUERPERA (MID);
alter table ZLSOL_JC.SOL_RS_POSTPARTUM_LIST
  add check (CONTENT IS JSON);

create table ZLSOL_JC.SOL_USERLIST
(
  user_code VARCHAR2(20) not null,
  user_name VARCHAR2(50)
)
;
comment on table ZLSOL_JC.SOL_USERLIST
  is '系统用户信息';
alter table ZLSOL_JC.SOL_USERLIST
  add constraint SOL_USERLIST_PK primary key (USER_CODE);

create sequence ZLSOL_JC.ISEQ$$_78144
minvalue 1
maxvalue 9999999999999999999999999999
start with 581
increment by 1
cache 20;

create sequence ZLSOL_JC.ISEQ$$_78154
minvalue 1
maxvalue 9999999999999999999999999999
start with 381
increment by 1
cache 20;

create sequence ZLSOL_JC.ISEQ$$_78167
minvalue 1
maxvalue 9999999999999999999999999999
start with 361
increment by 1
cache 20;

create sequence ZLSOL_JC.ISEQ$$_78174
minvalue 1
maxvalue 9999999999999999999999999999
start with 281
increment by 1
cache 20;

create sequence ZLSOL_JC.ISEQ$$_78189
minvalue 1
maxvalue 9999999999999999999999999999
start with 361
increment by 1
cache 20;

create sequence ZLSOL_JC.ISEQ$$_78203
minvalue 1
maxvalue 9999999999999999999999999999
start with 221
increment by 1
cache 20;

create or replace force view ZLSOL_JC.v_newborn as
Select a.Mid, b.Bid, b.Babyno, b.Addtime, b.Recorder, a.Name, a.Old, a.Bedno, a.Pno, b.Sex, d."隐藏2",d."身长",d."体重",d."头围",d."胸围",d."血型",d."胎儿状况",d."在院情况",d."死亡时间",
d."一般情况反应",d."一般情况面色",d."一般情况皮肤",d."一般情况毳毛",d."头部变形",d."颅骨重叠",d."产瘤大小",d."胎头水肿血肿",d."胎头水肿大小",d."前囟",d."张力",

d."生后即刻",d."生后半小时",d."生后一小时",d."生后二小时",d."生后三小时",d."生后四小时",

 d."足月",d."羊水清",d."正常呼吸或哭声",d."肌张力好",d."正常呼吸或哭声1",d."肌张力好1",d."心率",
        d."初步复苏30秒后心率", d."初步复苏30秒后呼吸",d."初步复苏30秒后肤色",
        d."正压通气30秒后心率",d."正压通气30秒后呼吸",d."正压通气30秒后肤色",
        d."继续正压通气几秒后心率",d."继续正压通气几秒后呼吸",d."继续正压通气几秒后肤色",
        d."正压通气加胸外按压30秒后心率", d."正压通气加胸外按压30秒后呼吸", d."正压通气加胸外按压30秒后肤色",
        d."使用肾上腺素后评估",d."实施其他重要措施后的评估",

       f.初步复苏步骤三十秒,f.初步复苏步骤六十秒,f.初步复苏步骤九十秒,f.初步复苏步骤两分钟,f.初步复苏步骤三分钟,f.初步复苏步骤五分钟,f.初步复苏步骤十分钟,f.初步复苏步骤二十分钟,
       f.常压给氧三十秒,f.常压给氧六十秒,f.常压给氧九十秒,f.常压给氧两分钟,f.常压给氧三分钟,f.常压给氧五分钟,f.常压给氧十分钟,f.常压给氧二十分钟,
       f.气管插管吸引胎粪三十秒,f.气管插管吸引胎粪六十秒,f.气管插管吸引胎粪九十秒,f.气管插管吸引胎粪两分钟,f.气管插管吸引胎粪三分钟,f.气管插管吸引胎粪五分钟,f.气管插管吸引胎粪十分钟,f.气管插管吸引胎粪二十分钟,
       f.正压通气三十秒,f.正压通气六十秒,f.正压通气九十秒,f.正压通气两分钟,f.正压通气三分钟,f.正压通气五分钟,f.正压通气十分钟,f.正压通气二十分钟,
       f.气管插管三十秒,f.气管插管六十秒,f.气管插管九十秒,f.气管插管两分钟,f.气管插管三分钟,f.气管插管五分钟,f.气管插管十分钟,f.气管插管二十分钟,
       f.胸外按压三十秒,f.胸外按压六十秒,f.胸外按压九十秒,f.胸外按压两分钟,f.胸外按压三分钟,f.胸外按压五分钟,f.胸外按压十分钟,f.胸外按压二十分钟,
       f.肾上腺素三十秒,f.肾上腺素六十秒,f.肾上腺素九十秒,f.肾上腺素两分钟,f.肾上腺素三分钟,f.肾上腺素五分钟,f.肾上腺素十分钟,f.肾上腺素二十分钟,
       f.生理盐水三十秒,f.生理盐水六十秒,f.生理盐水九十秒,f.生理盐水两分钟,f.生理盐水三分钟,f.生理盐水五分钟,f.生理盐水十分钟,f.生理盐水二十分钟,
        d."出生时间",d."复苏开始时间",d."复苏结束时间",d."分娩前时间",d."分娩后时间",d."主要复苏人员",d."抢救结局",
e."隐藏3",e."心率1分钟",e."心率5分钟",e."心率10分钟",
e."呼吸1分钟",e."呼吸5分钟",e."呼吸10分钟",e."喉反射1分钟",e."喉反射5分钟",e."喉反射10分钟",e."肌张力1分钟",e."肌张力5分钟",e."肌张力10分钟",e."肤色1分钟",
e."肤色5分钟",e."肤色10分钟",e."总分1分钟",e."总分5分钟",e."总分10分钟", f."隐藏4",f."出孕期产时合并症及用药情况",f."出生前胎儿情况",f."婴儿出生时抢救情况",
f."出生缺陷",f."母乳喂养指导",f."诊断",f."备注",f."接生者签名",f."剖宫产患者家属确认新生儿签名",f."与新生儿关系",

f."出院时体重",f."出院时脐带",f."出院时臀部",f."出院时皮肤",f."哺乳",f."卡介苗接种日期",f."卡介苗未接种原因",f."乙肝疫苗接种日期",f."乙肝疫苗未接种原因",
f."死亡原因",f."主管医师签名"
From Sol_Inf_Puerpera A, Sol_Inf_Newborns B,
     Json_Table(Nvl(b.Newborninf, '{"隐藏2":"2"}'),
                 '$' Columns(隐藏2 Varchar2(50) Path '$.隐藏2', 身长 Varchar2(50) Path '$.身长', 体重 Varchar2(50) Path '$.体重',
                          头围 Varchar2(50) Path '$.头围', 胸围 Varchar2(50) Path '$.胸围',血型 Varchar2(50) Path '$.血型',胎儿状况 Varchar2(50) Path '$.胎儿状况',
                          在院情况 Varchar2(10) PATH '$.在院情况',死亡时间 Varchar2(19) Path '$.死亡时间 ',一般情况反应 Varchar2(50) Path '$.一般情况反应',
                          一般情况面色 Varchar2(50) Path '$.一般情况面色', 一般情况皮肤 Varchar2(50) Path '$.一般情况皮肤',
                          一般情况毳毛 Varchar2(50) Path '$.一般情况毳毛', 头部变形 Varchar2(50) Path '$.头部变形',
                          颅骨重叠 Varchar2(50) Path '$.颅骨重叠', 产瘤大小 Varchar2(50) Path '$.产瘤大小',胎头水肿血肿 Varchar2(50) Path '$.胎头水肿血肿',
                          胎头水肿大小 Varchar2(50) Path '$.胎头水肿大小', 前囟 Varchar2(50) Path '$.前囟', 张力 Varchar2(50) Path '$.张力',
                          生后即刻 Varchar2(50) Path '$.生后即刻',
                          生后半小时 Varchar2(50) Path '$.生后半小时',生后一小时 Varchar2(50) Path '$.生后一小时',生后二小时 Varchar2(50) Path '$.生后二小时',
                          生后三小时 Varchar2(50) Path '$.生后三小时',生后四小时 Varchar2(50) Path '$.生后四小时',

                          足月 Varchar2(50) Path '$.足月', 羊水清 Varchar2(50) Path '$.羊水清',正常呼吸或哭声 Varchar2(50) Path '$.正常呼吸或哭声', 肌张力好 Varchar2(50) Path '$.肌张力好',
                          正常呼吸或哭声1 Varchar2(50) Path '$.正常呼吸或哭声1', 肌张力好1 Varchar2(50) Path '$.肌张力好1', 心率 Varchar2(50) Path '$.心率',

                          初步复苏30秒后心率 Varchar2(50) Path '$.初步复苏30秒后心率',初步复苏30秒后呼吸 Varchar2(50) Path '$.初步复苏30秒后呼吸',初步复苏30秒后肤色 Varchar2(50) Path '$.初步复苏30秒后肤色',
                          正压通气30秒后心率 Varchar2(50) Path '$.正压通气30秒后心率',正压通气30秒后呼吸 Varchar2(50) Path '$.正压通气30秒后呼吸',正压通气30秒后肤色 Varchar2(50) Path '$.正压通气30秒后肤色',
                          继续正压通气几秒后心率 Varchar2(50) Path '$.继续正压通气几秒后心率',继续正压通气几秒后呼吸 Varchar2(50) Path '$.继续正压通气几秒后呼吸',继续正压通气几秒后肤色 Varchar2(50) Path '$.继续正压通气几秒后肤色',
                          正压通气加胸外按压30秒后心率 Varchar2(50) Path '$.正压通气加胸外按压30秒后心率',正压通气加胸外按压30秒后呼吸 Varchar2(50) Path '$.正压通气加胸外按压30秒后呼吸',正压通气加胸外按压30秒后肤色 Varchar2(50) Path '$.正压通气加胸外按压30秒后肤色',
                          使用肾上腺素后评估 Varchar2(50) Path '$.使用肾上腺素后评估',实施其他重要措施后的评估 Varchar2(50) Path '$.实施其他重要措施后的评估',

                          出生时间 Varchar2(19) Path '$.出生时间',复苏开始时间 Varchar2(19) Path '$.复苏开始时间',复苏结束时间 Varchar2(19) Path '$.复苏结束时间',
                          分娩前时间 Varchar2(19) Path '$.分娩前时间', 分娩后时间 Varchar2(19) Path '$.分娩后时间',
                          主要复苏人员 Varchar2(50) Path '$.主要复苏人员',抢救结局 Varchar2(50) Path '$.抢救结局' )) as D,
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
                          诊断 Varchar2(50) Path '$.诊断  ',  备注 Varchar2(50) Path '$.备注  ',接生者签名 Varchar2(50) Path '$.接生者签名 ',
                          剖宫产患者家属确认新生儿签名 Varchar2(50) Path '$.剖宫产患者家属确认新生儿签名', 与新生儿关系 Varchar2(50) Path '$.与新生儿关系',

                          出院时体重 Varchar2(50) Path '$.出院时体重  ',出院时脐带 Varchar2(50) Path '$.出院时脐带  ',出院时臀部 Varchar2(50) Path '$.出院时臀部  ',
                          出院时皮肤 Varchar2(50) Path '$.出院时皮肤  ',哺乳 Varchar2(50) Path '$.哺乳  ',
                          卡介苗接种日期 Varchar2(50) Path '$.卡介苗接种日期  ',卡介苗未接种原因 Varchar2(50) Path '$.卡介苗未接种原因  ',
                          乙肝疫苗接种日期 Varchar2(50) Path '$.乙肝疫苗接种日期  ',乙肝疫苗未接种原因 Varchar2(50) Path '$.乙肝疫苗未接种原因  ',
                          死亡原因 Varchar2(50) Path '$.死亡原因  ',主管医师签名 Varchar2(50) Path '$.主管医师签名  ',

                          初步复苏步骤三十秒 Varchar2(50) Path '$.初步复苏步骤三十秒',初步复苏步骤六十秒 Varchar2(50) Path '$.初步复苏步骤六十秒',初步复苏步骤九十秒 Varchar2(50) Path '$.初步复苏步骤九十秒',
                          初步复苏步骤两分钟 Varchar2(50) Path '$.初步复苏步骤两分钟',初步复苏步骤三分钟 Varchar2(50) Path '$.初步复苏步骤三分钟',初步复苏步骤五分钟 Varchar2(50) Path '$.初步复苏步骤五分钟',
                          初步复苏步骤十分钟 Varchar2(50) Path '$.初步复苏步骤十分钟',初步复苏步骤二十分钟 Varchar2(50) Path '$.初步复苏步骤二十分钟',
                          常压给氧三十秒 Varchar2(50) Path '$.常压给氧三十秒',常压给氧六十秒 Varchar2(50) Path '$.常压给氧六十秒',常压给氧九十秒 Varchar2(50) Path '$.常压给氧九十秒',
                          常压给氧两分钟 Varchar2(50) Path '$.常压给氧两分钟',常压给氧三分钟 Varchar2(50) Path '$.常压给氧三分钟',常压给氧五分钟 Varchar2(50) Path '$.常压给氧五分钟',
                          常压给氧十分钟 Varchar2(50) Path '$.常压给氧十分钟',常压给氧二十分钟 Varchar2(50) Path '$.常压给氧二十分钟',
                          气管插管吸引胎粪三十秒 Varchar2(50) Path '$.气管插管吸引胎粪三十秒',气管插管吸引胎粪六十秒 Varchar2(50) Path '$.气管插管吸引胎粪六十秒',气管插管吸引胎粪九十秒 Varchar2(50) Path '$.气管插管吸引胎粪九十秒',
                          气管插管吸引胎粪两分钟 Varchar2(50) Path '$.气管插管吸引胎粪两分钟',气管插管吸引胎粪三分钟 Varchar2(50) Path '$.气管插管吸引胎粪三分钟',气管插管吸引胎粪五分钟 Varchar2(50) Path '$.气管插管吸引胎粪五分钟',
                          气管插管吸引胎粪十分钟 Varchar2(50) Path '$.气管插管吸引胎粪十分钟',气管插管吸引胎粪二十分钟 Varchar2(50) Path '$.气管插管吸引胎粪二十分钟',
                          正压通气三十秒 Varchar2(50) Path '$.正压通气三十秒',正压通气六十秒 Varchar2(50) Path '$.正压通气六十秒',正压通气九十秒 Varchar2(50) Path '$.正压通气九十秒',
                          正压通气两分钟 Varchar2(50) Path '$.正压通气两分钟',正压通气三分钟 Varchar2(50) Path '$.正压通气三分钟',正压通气五分钟 Varchar2(50) Path '$.正压通气五分钟',
                          正压通气十分钟 Varchar2(50) Path '$.正压通气十分钟',正压通气二十分钟 Varchar2(50) Path '$.正压通气二十分钟',
                          气管插管三十秒 Varchar2(50) Path '$.气管插管三十秒',气管插管六十秒 Varchar2(50) Path '$.气管插管六十秒',气管插管九十秒 Varchar2(50) Path '$.气管插管九十秒',
                          气管插管两分钟 Varchar2(50) Path '$.气管插管两分钟',气管插管三分钟 Varchar2(50) Path '$.气管插管三分钟',气管插管五分钟 Varchar2(50) Path '$.气管插管五分钟',
                          气管插管十分钟 Varchar2(50) Path '$.气管插管十分钟',气管插管二十分钟 Varchar2(50) Path '$.气管插管二十分钟',
                          胸外按压三十秒 Varchar2(50) Path '$.胸外按压三十秒',胸外按压六十秒 Varchar2(50) Path '$.胸外按压六十秒',胸外按压九十秒 Varchar2(50) Path '$.胸外按压九十秒',
                          胸外按压两分钟 Varchar2(50) Path '$.胸外按压两分钟',胸外按压三分钟 Varchar2(50) Path '$.胸外按压三分钟',胸外按压五分钟 Varchar2(50) Path '$.胸外按压五分钟',
                          胸外按压十分钟 Varchar2(50) Path '$.胸外按压十分钟',胸外按压二十分钟 Varchar2(50) Path '$.胸外按压二十分钟',
                          肾上腺素三十秒 Varchar2(50) Path '$.肾上腺素三十秒',肾上腺素六十秒 Varchar2(50) Path '$.肾上腺素六十秒',肾上腺素九十秒 Varchar2(50) Path '$.肾上腺素九十秒',
                          肾上腺素两分钟 Varchar2(50) Path '$.肾上腺素两分钟',肾上腺素三分钟 Varchar2(50) Path '$.肾上腺素三分钟',肾上腺素五分钟 Varchar2(50) Path '$.肾上腺素五分钟',
                          肾上腺素十分钟 Varchar2(50) Path '$.肾上腺素十分钟',肾上腺素二十分钟 Varchar2(50) Path '$.肾上腺素二十分钟',
                          生理盐水三十秒 Varchar2(50) Path '$.生理盐水三十秒',生理盐水六十秒 Varchar2(50) Path '$.生理盐水六十秒',生理盐水九十秒 Varchar2(50) Path '$.生理盐水九十秒',
                          生理盐水两分钟 Varchar2(50) Path '$.生理盐水两分钟',生理盐水三分钟 Varchar2(50) Path '$.生理盐水三分钟',生理盐水五分钟 Varchar2(50) Path '$.生理盐水五分钟',
                          生理盐水十分钟 Varchar2(50) Path '$.生理盐水十分钟',生理盐水二十分钟 Varchar2(50) Path '$.生理盐水二十分钟' )) As F
Where a.Mid = b.Mid
Order By Addtime;

create or replace force view ZLSOL_JC.v_sol_inf_delivery as
Select a.Mid, b."隐藏1", b."产程开始时间", b."宫口全开时间", b."胎儿娩出时间", b."胎盘娩出时间", b."第一产程", b."第二产程", b."第三产程", b."宫缩情况",
       b."结扎", b."破膜方式", b."破膜时间", b."羊水性状", b."羊水量", b."羊水颜色", b."胎盘娩出方式", b."胎盘剥离方式", b."胎盘完整度",
       b."胎盘胎膜残留", b."胎盘体积", b."胎盘形态", b."胎盘大小", b."胎盘重量", b."脐带附着", b."脐带长度", b."脐带真假结", b."脐带",b."绕颈周数",b."绕颈周数1",b."钙化程度",b."娩出方式",
       b."娩出胎方位", b."产瘤大小", b."产瘤部位", b."会阴裂伤程度", b."会阴裂伤切口", b."会阴裂伤缝合", b."会阴裂伤麻醉", b."宫颈裂伤长度", b."宫颈裂伤部位", b."宫颈裂伤状况",
       b."阴道裂伤部位大小",b."阴道裂伤部位长度", b."阴道裂伤血肿部位", b."阴道裂伤血肿大小",

       b."足月",b."羊水清",b."正常呼吸或哭声",b."肌张力好",b."正常呼吸或哭声1",b."肌张力好1",b."心率",
        b."初步复苏30秒后心率", b."初步复苏30秒后呼吸",b."初步复苏30秒后肤色",
        b."正压通气30秒后心率",b."正压通气30秒后呼吸",b."正压通气30秒后肤色",
        b."继续正压通气几秒后心率",b."继续正压通气几秒后呼吸",b."继续正压通气几秒后肤色",
        b."正压通气加胸外按压30秒后心率", b."正压通气加胸外按压30秒后呼吸", b."正压通气加胸外按压30秒后肤色",
        b."使用肾上腺素后评估",b."实施其他重要措施后的评估",

       b.初步复苏步骤三十秒,b.初步复苏步骤六十秒,b.初步复苏步骤九十秒,b.初步复苏步骤两分钟,b.初步复苏步骤三分钟,b.初步复苏步骤五分钟,b.初步复苏步骤十分钟,b.初步复苏步骤二十分钟,
       b.常压给氧三十秒,b.常压给氧六十秒,b.常压给氧九十秒,b.常压给氧两分钟,b.常压给氧三分钟,b.常压给氧五分钟,b.常压给氧十分钟,b.常压给氧二十分钟,
       b.气管插管吸引胎粪三十秒,b.气管插管吸引胎粪六十秒,b.气管插管吸引胎粪九十秒,b.气管插管吸引胎粪两分钟,b.气管插管吸引胎粪三分钟,b.气管插管吸引胎粪五分钟,b.气管插管吸引胎粪十分钟,b.气管插管吸引胎粪二十分钟,
       b.正压通气三十秒,b.正压通气六十秒,b.正压通气九十秒,b.正压通气两分钟,b.正压通气三分钟,b.正压通气五分钟,b.正压通气十分钟,b.正压通气二十分钟,
       b.气管插管三十秒,b.气管插管六十秒,b.气管插管九十秒,b.气管插管两分钟,b.气管插管三分钟,b.气管插管五分钟,b.气管插管十分钟,b.气管插管二十分钟,
       b.胸外按压三十秒,b.胸外按压六十秒,b.胸外按压九十秒,b.胸外按压两分钟,b.胸外按压三分钟,b.胸外按压五分钟,b.胸外按压十分钟,b.胸外按压二十分钟,
       b.肾上腺素三十秒,b.肾上腺素六十秒,b.肾上腺素九十秒,b.肾上腺素两分钟,b.肾上腺素三分钟,b.肾上腺素五分钟,b.肾上腺素十分钟,b.肾上腺素二十分钟,
       b.生理盐水三十秒,b.生理盐水六十秒,b.生理盐水九十秒,b.生理盐水两分钟,b.生理盐水三分钟,b.生理盐水五分钟,b.生理盐水十分钟,b.生理盐水二十分钟,

       b."出生时间",b."复苏开始时间",b."复苏结束时间",b."分娩前时间",b."分娩后时间",b."主要复苏人员",b."抢救结局",

       b."母婴早接触早吸吮开始时间",b."母婴早接触早吸吮结束时间", b."产后血压", b."收缩压", b."舒张压", b."产后流血",b."出血处理", b."产时用药", b."产后用药",
       b."产后诊断",b."产后诊断2",b."产后诊断3",b."产后诊断4",b."手术名称", b."手术名称2",b."手术名称3",b."手术名称4",
       b."特殊情况", b."未吸吮原因",d."隐藏3",d."出产房时间",d."出产房宫高脐下", d."护送人", d."接生人", d."记录人"
From Sol_Inf_Delivery A,
     Json_Table(Nvl(a.Newborndetail, '{"隐藏1":"1"}'),
                 '$'
                  Columns(隐藏1 Varchar2(50) Path '$.隐藏1', 产程开始时间 Varchar2(19) Path '$.产程开始时间',
                          宫口全开时间 Varchar2(19) Path '$.宫口全开时间', 胎儿娩出时间 Varchar2(19) Path '$.胎儿娩出时间',
                          胎盘娩出时间 Varchar2(19) Path '$.胎盘娩出时间', 第一产程 Varchar2(50) Path '$.第一产程',
                          第二产程 Varchar2(50) Path '$.第二产程', 第三产程 Varchar2(50) Path '$.第三产程',结扎 Varchar2(50) Path '$.结扎',
                          破膜方式 Varchar2(50) Path '$.破膜方式', 破膜时间 Varchar2(19) Path '$.破膜时间', 羊水性状 Varchar2(50) Path '$.羊水性状',
                          羊水量 Varchar2(50) Path '$.羊水量', 羊水颜色 Varchar2(50) Path '$.羊水颜色',
                          胎盘娩出方式 Varchar2(50) Path '$.胎盘娩出方式', 胎盘剥离方式 Varchar2(50) Path '$.胎盘剥离方式',
                          胎盘完整度 Varchar2(50) Path '$.胎盘完整度', 胎盘胎膜残留 Varchar2(50) Path '$.胎盘胎膜残留',
                          胎盘体积 Varchar2(50) Path '$.胎盘体积', 胎盘形态 Varchar2(50) Path '$.胎盘形态', 胎盘大小 Varchar2(50) Path '$.胎盘大小',
                          胎盘重量 Varchar2(50) Path '$.胎盘重量', 脐带附着 Varchar2(50) Path '$.脐带附着', 脐带长度 Varchar2(50) Path '$.脐带长度',
                         脐带真假结 Varchar2(50) Path '$.脐带真假结',脐带 Varchar2(50) Path '$.脐带', 绕颈周数 Varchar2(50) Path '$.绕颈周数',绕颈周数1 Varchar2(50) Path '$.绕颈周数1',
                         钙化程度 Varchar2(20) Path '$.钙化程度',
                          娩出方式 Varchar2(50) Path '$.娩出方式',
                          娩出胎方位 Varchar2(50) Path '$.娩出胎方位', 产瘤大小 Varchar2(50) Path '$.产瘤大小',
                          产瘤部位 Varchar2(50) Path '$.产瘤部位', 会阴裂伤程度 Varchar2(50) Path '$.会阴裂伤程度',
                          会阴裂伤切口 Varchar2(50) Path '$.会阴裂伤切口', 会阴裂伤缝合 Varchar2(50) Path '$.会阴裂伤缝合',
                          会阴裂伤麻醉 Varchar2(50) Path '$.会阴裂伤麻醉', 宫颈裂伤长度 Varchar2(50) Path '$.宫颈裂伤长度',
                          宫颈裂伤部位 Varchar2(50) Path '$.宫颈裂伤部位', 宫颈裂伤状况 Varchar2(50) Path '$.宫颈裂伤状况',
                          阴道裂伤部位大小 Varchar2(50) Path '$.阴道裂伤部位大小', 阴道裂伤部位长度 Varchar2(50) Path '$.阴道裂伤部位长度',
                          阴道裂伤血肿部位 Varchar2(50) Path '$.阴道裂伤血肿部位', 阴道裂伤血肿大小 Varchar2(50) Path '$.阴道裂伤血肿大小',

                          足月 Varchar2(50) Path '$.足月', 羊水清 Varchar2(50) Path '$.羊水清',正常呼吸或哭声 Varchar2(50) Path '$.正常呼吸或哭声', 肌张力好 Varchar2(50) Path '$.肌张力好',
                          正常呼吸或哭声1 Varchar2(50) Path '$.正常呼吸或哭声1', 肌张力好1 Varchar2(50) Path '$.肌张力好1', 心率 Varchar2(50) Path '$.心率',

                          初步复苏30秒后心率 Varchar2(50) Path '$.初步复苏30秒后心率',初步复苏30秒后呼吸 Varchar2(50) Path '$.初步复苏30秒后呼吸',初步复苏30秒后肤色 Varchar2(50) Path '$.初步复苏30秒后肤色',
                          正压通气30秒后心率 Varchar2(50) Path '$.正压通气30秒后心率',正压通气30秒后呼吸 Varchar2(50) Path '$.正压通气30秒后心率',正压通气30秒后肤色 Varchar2(50) Path '$.正压通气30秒后心率',
                          继续正压通气几秒后心率 Varchar2(50) Path '$.继续正压通气几秒后心率',继续正压通气几秒后呼吸 Varchar2(50) Path '$.继续正压通气几秒后呼吸',继续正压通气几秒后肤色 Varchar2(50) Path '$.继续正压通气几秒后肤色',
                          正压通气加胸外按压30秒后心率 Varchar2(50) Path '$.继续正压通气几秒后心率',正压通气加胸外按压30秒后呼吸 Varchar2(50) Path '$.继续正压通气几秒后呼吸',正压通气加胸外按压30秒后肤色 Varchar2(50) Path '$.继续正压通气几秒后肤色',
                          使用肾上腺素后评估 Varchar2(50) Path '$.使用肾上腺素后评估',实施其他重要措施后的评估 Varchar2(50) Path '$.实施其他重要措施后的评估',

                          初步复苏步骤三十秒 Varchar2(50) Path '$.初步复苏步骤三十秒',初步复苏步骤六十秒 Varchar2(50) Path '$.初步复苏步骤六十秒',初步复苏步骤九十秒 Varchar2(50) Path '$.初步复苏步骤九十秒',
                          初步复苏步骤两分钟 Varchar2(50) Path '$.初步复苏步骤两分钟',初步复苏步骤三分钟 Varchar2(50) Path '$.初步复苏步骤三分钟',初步复苏步骤五分钟 Varchar2(50) Path '$.初步复苏步骤五分钟',
                          初步复苏步骤十分钟 Varchar2(50) Path '$.初步复苏步骤十分钟',初步复苏步骤二十分钟 Varchar2(50) Path '$.初步复苏步骤二十分钟',
                          常压给氧三十秒 Varchar2(50) Path '$.常压给氧三十秒',常压给氧六十秒 Varchar2(50) Path '$.常压给氧六十秒',常压给氧九十秒 Varchar2(50) Path '$.常压给氧九十秒',
                          常压给氧两分钟 Varchar2(50) Path '$.常压给氧两分钟',常压给氧三分钟 Varchar2(50) Path '$.常压给氧三分钟',常压给氧五分钟 Varchar2(50) Path '$.常压给氧五分钟',
                          常压给氧十分钟 Varchar2(50) Path '$.常压给氧十分钟',常压给氧二十分钟 Varchar2(50) Path '$.常压给氧二十分钟',
                          气管插管吸引胎粪三十秒 Varchar2(50) Path '$.气管插管吸引胎粪三十秒',气管插管吸引胎粪六十秒 Varchar2(50) Path '$.气管插管吸引胎粪六十秒',气管插管吸引胎粪九十秒 Varchar2(50) Path '$.气管插管吸引胎粪九十秒',
                          气管插管吸引胎粪两分钟 Varchar2(50) Path '$.气管插管吸引胎粪两分钟',气管插管吸引胎粪三分钟 Varchar2(50) Path '$.气管插管吸引胎粪三分钟',气管插管吸引胎粪五分钟 Varchar2(50) Path '$.气管插管吸引胎粪五分钟',
                          气管插管吸引胎粪十分钟 Varchar2(50) Path '$.气管插管吸引胎粪十分钟',气管插管吸引胎粪二十分钟 Varchar2(50) Path '$.气管插管吸引胎粪二十分钟',
                          正压通气三十秒 Varchar2(50) Path '$.正压通气三十秒',正压通气六十秒 Varchar2(50) Path '$.正压通气六十秒',正压通气九十秒 Varchar2(50) Path '$.正压通气九十秒',
                          正压通气两分钟 Varchar2(50) Path '$.正压通气两分钟',正压通气三分钟 Varchar2(50) Path '$.正压通气三分钟',正压通气五分钟 Varchar2(50) Path '$.正压通气五分钟',
                          正压通气十分钟 Varchar2(50) Path '$.正压通气十分钟',正压通气二十分钟 Varchar2(50) Path '$.正压通气二十分钟',
                          气管插管三十秒 Varchar2(50) Path '$.气管插管三十秒',气管插管六十秒 Varchar2(50) Path '$.气管插管六十秒',气管插管九十秒 Varchar2(50) Path '$.气管插管九十秒',
                          气管插管两分钟 Varchar2(50) Path '$.气管插管两分钟',气管插管三分钟 Varchar2(50) Path '$.气管插管三分钟',气管插管五分钟 Varchar2(50) Path '$.气管插管五分钟',
                          气管插管十分钟 Varchar2(50) Path '$.气管插管十分钟',气管插管二十分钟 Varchar2(50) Path '$.气管插管二十分钟',
                          胸外按压三十秒 Varchar2(50) Path '$.胸外按压三十秒',胸外按压六十秒 Varchar2(50) Path '$.胸外按压六十秒',胸外按压九十秒 Varchar2(50) Path '$.胸外按压九十秒',
                          胸外按压两分钟 Varchar2(50) Path '$.胸外按压两分钟',胸外按压三分钟 Varchar2(50) Path '$.胸外按压三分钟',胸外按压五分钟 Varchar2(50) Path '$.胸外按压五分钟',
                          胸外按压十分钟 Varchar2(50) Path '$.胸外按压十分钟',胸外按压二十分钟 Varchar2(50) Path '$.胸外按压二十分钟',
                          肾上腺素三十秒 Varchar2(50) Path '$.肾上腺素三十秒',肾上腺素六十秒 Varchar2(50) Path '$.肾上腺素六十秒',肾上腺素九十秒 Varchar2(50) Path '$.肾上腺素九十秒',
                          肾上腺素两分钟 Varchar2(50) Path '$.肾上腺素两分钟',肾上腺素三分钟 Varchar2(50) Path '$.肾上腺素三分钟',肾上腺素五分钟 Varchar2(50) Path '$.肾上腺素五分钟',
                          肾上腺素十分钟 Varchar2(50) Path '$.肾上腺素十分钟',肾上腺素二十分钟 Varchar2(50) Path '$.肾上腺素二十分钟',
                          生理盐水三十秒 Varchar2(50) Path '$.生理盐水三十秒',生理盐水六十秒 Varchar2(50) Path '$.生理盐水六十秒',生理盐水九十秒 Varchar2(50) Path '$.生理盐水九十秒',
                          生理盐水两分钟 Varchar2(50) Path '$.生理盐水两分钟',生理盐水三分钟 Varchar2(50) Path '$.生理盐水三分钟',生理盐水五分钟 Varchar2(50) Path '$.生理盐水五分钟',
                          生理盐水十分钟 Varchar2(50) Path '$.生理盐水十分钟',生理盐水二十分钟 Varchar2(50) Path '$.生理盐水二十分钟',

                          出生时间 Varchar2(19) Path '$.出生时间',复苏开始时间 Varchar2(19) Path '$.复苏开始时间',复苏结束时间 Varchar2(19) Path '$.复苏结束时间',
                          分娩前时间 Varchar2(19) Path '$.分娩前时间', 分娩后时间 Varchar2(19) Path '$.分娩后时间',
                          主要复苏人员 Varchar2(19) Path '$.主要复苏人员',抢救结局 Varchar2(19) Path '$.主要复苏人员',

                          宫缩情况 Varchar2(50) Path '$.宫缩情况', 母婴早接触早吸吮开始时间 Varchar2(50) Path '$.母婴早接触早吸吮开始时间',
                          母婴早接触早吸吮结束时间 Varchar2(50) Path '$.母婴早接触早吸吮结束时间',产后血压 Varchar2(50) Path '$.产后血压',
                          收缩压 Varchar2(50) Path '$.收缩压',
                          舒张压 Varchar2(50) Path '$.舒张压',
                          产后流血 Varchar2(50) Path '$.产后流血', 出血处理 Varchar2(50) Path '$.出血处理', 产时用药 Varchar2(50) Path '$.产时用药', 产后用药 Varchar2(50) Path '$.产后用药',
                          产后诊断 Varchar2(50) Path '$.产后诊断', 产后诊断2 Varchar2(50) Path '$.产后诊断2', 产后诊断3 Varchar2(50) Path '$.产后诊断3', 产后诊断4 Varchar2(50) Path '$.产后诊断4',
                          手术名称 Varchar2(50) Path '$.手术名称',手术名称2 Varchar2(50) Path '$.手术名称2', 手术名称3 Varchar2(50) Path '$.手术名称3', 手术名称4 Varchar2(50) Path '$.手术名称4',
                          特殊情况 Varchar2(50) Path '$.特殊情况',未吸吮原因 Varchar2(50) Path '$.未吸吮原因'
                          )) As B,
     Json_Table(Nvl(a.Deliveryinf, '{"隐藏3":"1"}'),
                 '$' Columns(隐藏3 Varchar2(1) Path '$.隐藏3', 出产房时间 Varchar2(50) Path '$.出产房时间',出产房宫高脐下 Varchar2(50) Path '$.出产房宫高脐下',
                          护送人 Varchar2(50) Path '$.护送人', 接生人 Varchar2(50) Path '$.接生人', 记录人 Varchar2(50) Path '$.记录人')
                          ) As D;

create or replace force view ZLSOL_JC.v_sol_rs_birth as
Select a.mid,b."妊次",b."产次",b."血型",b."既往妊娠史",b."末次月经",b."预产期",b."髂前上棘间径",b."髂嵴间径",b."坐骨结节间径",b."骶耻外径",b."骶骨弧度",b."骶骨关节",b."坐骨切迹",b."坐骨髂",b."并发症",b."产前记录特征",b."检查时间",b."血压",b."收缩压",b."舒张压",b."体温",b."脉博",b."胎心率",b."胎儿大小",b."宫缩规律性",b."胎位",b."衔接",b."破膜情况",b."先露",b."宫口",b."检查者",b."宫缩开始时间",b."破膜时间",b."入院处理"
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
血压            Varchar2(10) PATH '$.血压',
收缩压            Varchar2(10) PATH '$.收缩压',
舒张压            Varchar2(10) PATH '$.舒张压 ',
体温            Number(4,2)  PATH '$.体温',
脉博            Varchar2(10) PATH '$.脉博',
胎心率            Varchar2(10) PATH '$.胎心率',
胎儿大小        Number(5,2) PATH '$.胎儿大小',
宫缩规律性      Varchar2(10) PATH '$.宫缩规律性',
胎位            Varchar2(10) PATH '$.胎位',
衔接            Varchar2(10) PATH '$.衔接',
破膜情况        Varchar2(10) PATH '$.破膜情况',
先露            Varchar2(2) PATH '$.先露',
宫口            Number(4,2) PATH '$.宫口',
检查者          Varchar2(50) PATH '$.检查者',
宫缩开始时间    Varchar2(20) PATH '$.宫缩开始时间',
破膜时间        Varchar2(20) PATH '$.破膜时间',
入院处理        Varchar2(100) PATH '$.入院处理'
)) as b;

create or replace force view ZLSOL_JC.v_his_病人新生儿记录 as
select d.pid 病人id,d.tid 住院次数,b.Babyno 序号,d.name||decode(b.Sex,'男','之子','之女')||t.顺序 as 婴儿姓名,
b.Sex as 婴儿性别,
c.妊次 as 分娩次数,
a.娩出方式,b.胎儿状况,
b.身长,b.体重,b.血型,
a.胎儿娩出时间 as 出生时间,
b.死亡时间,'' as 备注说明,
b.Recorder 登记人,
b.Addtime 登记时间
from V_SOL_INF_DELIVERY a,v_newborn b,v_sol_rs_birth c,sol_inf_puerpera d,
(select mid,sex,babyno,row_number() over(partition by mid,sex order by babyno) 顺序 from SOL_INF_NEWBORNS t  ) t
where a.Mid=b.Mid and a.mid=d.mid and a.mid=t.mid and b.Babyno=t.babyno
and b.Mid=c.mid(+);

create or replace force view ZLSOL_JC.v_sol_inf_checkinroom as
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

create or replace force view ZLSOL_JC.v_sol_inf_checkoutroom as
Select a.mid, b."OUTROOMTIME",b."出房状态",b."医疗病历",b."护理病历",b."静脉通道",b."局部情况",b."会阴裂伤",b."会阴切开术",b."会阴切口缝合",b."会阴水肿",b."会阴血肿",b."阴道出血",b."出血量",
b."阴道填塞纱卷",b."排尿情况",b."留置尿管",b."特殊药物",b."交班者",b."接班者",b."药物",b."备注"
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
阴道出血     Varchar2(10) PATH '$.阴道出血',
出血量       Number(5) PATH '$.出血量',
阴道填塞纱卷   Varchar2(20) PATH '$.阴道填塞纱卷',
排尿情况     Varchar2(20) PATH '$.排尿情况',
留置尿管     Varchar2(20) PATH '$.留置尿管',
特殊药物     Varchar2(50) PATH '$.特殊药物',
交班者       Varchar2(20) PATH '$.交班者',
接班者       Varchar2(20) PATH '$.接班者',
药物         Varchar2(50) PATH '$.药物',
备注         Varchar2(50) PATH '$.备注'
)) as b;

create or replace force view ZLSOL_JC.v_sol_inf_equipment as
Select a.mid, b."侧切剪产前",b."侧切剪术中",b."侧切剪产后",b."脐带剪产前",b."脐带剪术中",b."脐带剪产后",
b."止血钳产前",b."止血钳术中",b."止血钳产后",b."牙镊产前",b."牙镊术中",b."牙镊产后",
b."持针器产前",b."持针器术中",b."持针器产后",b."穿刺针产前",b."穿刺针术中",b."穿刺针产后",b."洗耳球产前",
b."洗耳球术中",b."洗耳球产后",b."胎吸产前",b."胎吸术中",b."胎吸产后",b."缝合针产前",b."缝合针术中",b."缝合针产后",
b."拉钩产前",b."拉钩术中",b."拉钩产后",b."纱布产前",b."纱布术中",b."纱布产后",b."卵圆钳产前",b."卵圆钳术中",b."卵圆钳产后",
b."宫颈钳产前",b."宫颈钳术中",b."宫颈钳产后",
b."窥器产前",b."窥器术中",b."窥器产后",b."刮匙产前",b."刮匙术中",b."刮匙产后",b."艾利斯产前",b."艾利斯术中",b."艾利斯产后",
b."产钳产前",b."产钳术中",b."产钳产后"
From SOL_INF_Equipment a,JSON_TABLE(a.Content,'$' columns(
侧切剪产前   varchar2(2) PATH '$.侧切剪产前',
侧切剪术中   varchar2(2) PATH '$.侧切剪术中',
侧切剪产后   varchar2(2) PATH '$.侧切剪产后',
脐带剪产前   varchar2(2) PATH '$.脐带剪产前',
脐带剪术中   varchar2(2) PATH '$.脐带剪术中',
脐带剪产后   varchar2(2) PATH '$.脐带剪产后',
止血钳产前   varchar2(2) PATH '$.止血钳产前',
止血钳术中   varchar2(2) PATH '$.止血钳术中',
止血钳产后   varchar2(2) PATH '$.止血钳产后',
牙镊产前   varchar2(2) PATH '$.牙镊产前',
牙镊术中   varchar2(2) PATH '$.牙镊术中',
牙镊产后   varchar2(2) PATH '$.牙镊产后',
持针器产前   varchar2(2) PATH '$.持针器产前',
持针器术中   varchar2(2) PATH '$.持针器术中',
持针器产后   varchar2(2) PATH '$.持针器产后',
穿刺针产前   varchar2(2) PATH '$.穿刺针产前',
穿刺针术中   varchar2(2) PATH '$.穿刺针术中',
穿刺针产后   varchar2(2) PATH '$.穿刺针产后',
洗耳球产前   varchar2(2) PATH '$.洗耳球产前',
洗耳球术中   varchar2(2) PATH '$.洗耳球术中',
洗耳球产后   varchar2(2) PATH '$.洗耳球产后',
胎吸产前   varchar2(2) PATH '$.胎吸产前',
胎吸术中   varchar2(2) PATH '$.胎吸术中',
胎吸产后   varchar2(2) PATH '$.胎吸产后',
缝合针产前   varchar2(2) PATH '$.缝合针产前',
缝合针术中   varchar2(2) PATH '$.缝合针术中',
缝合针产后   varchar2(2) PATH '$.缝合针产后',
拉钩产前   varchar2(2) PATH '$.拉钩产前',
拉钩术中   varchar2(2) PATH '$.拉钩术中',
拉钩产后   varchar2(2) PATH '$.拉钩产后',
产钳产前   varchar2(2) PATH '$.产钳产前',
产钳术中   varchar2(2) PATH '$.产钳术中',
产钳产后   varchar2(2) PATH '$.产钳产后',
纱布产前   varchar2(2) PATH '$.纱布产前',
纱布术中   varchar2(2) PATH '$.纱布术中',
纱布产后   varchar2(2) PATH '$.纱布产后',
卵圆钳产前   varchar2(2) PATH '$.卵圆钳产前',
卵圆钳术中   varchar2(2) PATH '$.卵圆钳术中',
卵圆钳产后   varchar2(2) PATH '$.卵圆钳产后',
宫颈钳产前   varchar2(2) PATH '$.宫颈钳产前',
宫颈钳术中   varchar2(2) PATH '$.宫颈钳术中',
宫颈钳产后   varchar2(2) PATH '$.宫颈钳产后',
窥器产前   varchar2(2) PATH '$.窥器产前',
窥器术中   varchar2(2) PATH '$.窥器术中',
窥器产后   varchar2(2) PATH '$.窥器产后',
刮匙产前   varchar2(2) PATH '$.刮匙产前',
刮匙术中   varchar2(2) PATH '$.刮匙术中',
刮匙产后   varchar2(2) PATH '$.刮匙产后',
艾利斯产前   varchar2(2) PATH '$.艾利斯产前',
艾利斯术中   varchar2(2) PATH '$.艾利斯术中',
艾利斯产后   varchar2(2) PATH '$.艾利斯产后'
)) as b;

create or replace force view ZLSOL_JC.v_sol_inf_newborns as
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

create or replace force view ZLSOL_JC.v_sol_inf_puerpera as
Select Name, Mid, Old, LPad(Bedno, 10) Bedno, Pno, Diagnosis, Status, Decode(Expectant, 1, '√', '') 待产,
       Decode(Checkinroom, 1, '√', '') 入房, Decode(Birth, 1, '√', '') 临产, Decode(Druglabor, 1, '√', '') 引产,
       Decode(Delivery, 1, '√', '') 分娩, Decode(Newborns, 1, '√', '') 新生儿, Decode(Postpartum, 1, '√', '') 产后,
       Decode(Checkoutroom, 1, '√', '') 出房,Decode(Equipment, 1, '√', '') 器械,outtime,pid,tid
From Sol_Inf_Puerpera;

create or replace force view ZLSOL_JC.v_sol_rs_birth_course as
Select  a.courseid,a.mid,b."检查时间",b."胎方位",b."血压",b."舒张压",b."收缩压",b."血糖",b."呼吸",b."体温",b."脉博",b."胎心率",b."宫缩强度",b."宫缩持续",b."宫缩间隔",b."宫颈管消失",b."宫口",b."羊水",b."破膜情况",b."分娩情况",b."先露",b."血氧饱和度",b."意识",b."处理",b."检查者"
From SOL_RS_BIRTH_COURSE a,JSON_TABLE(a.CONTENT,'$' columns(
检查时间        Varchar2(20)  PATH '$.检查时间',
胎方位 Varchar2(20)  PATH '$.胎方位',
血压        Varchar2(10) PATH '$.血压',

舒张压        Varchar2(10) PATH '$.舒张压',
收缩压        Varchar2(10) PATH '$.收缩压',
血糖          Varchar2(10) PATH '$.血糖',
呼吸          Varchar2(10) PATH  '$.呼吸',
体温        Number(4,2)  PATH '$.体温',
脉博        Varchar2(10) PATH '$.脉博',
胎心率        Varchar2(10) PATH '$.胎心率',
宫缩强度    Varchar2(10) PATH '$.宫缩强度',
宫缩持续  Varchar2(10) PATH '$.宫缩持续',
宫缩间隔  Varchar2(10) PATH '$.宫缩间隔',
宫颈管消失    Varchar2(20) PATH '$.宫颈管消失',
宫口        Number(4,2) PATH '$.宫口',
先露        Number(2) PATH '$.先露'，
血氧饱和度     Varchar2(10) PATH '$.血氧饱和度'，
意识     Varchar2(10) PATH '$.意识'，
羊水        Varchar2(10) PATH '$.羊水'，
破膜情况    Varchar2(10) PATH '$.破膜情况',
分娩情况    Varchar2(20) PATH '$.分娩情况',
处理        Varchar2(500) PATH '$.处理'，
检查者      Varchar2(50) PATH '$.检查者'
)) as b;

create or replace force view ZLSOL_JC.v_sol_rs_druglabor as
Select Mid, To_Char(日期, 'YYYY-MM-DD HH24:MI') 日期, 引产指征, 引产方法 from Sol_Rs_Druglabor;

create or replace force view ZLSOL_JC.v_sol_rs_druglabor_list as
Select a.Mid, a.Courseid ID, b."记录时间",b."血压",b."收缩压",b."舒张压",b."脉搏",b."胎心率",b."宫缩强度",b."宫缩持续",b."宫缩间隔",b."宫口",b."意识",b."体温",b."呼吸",b."先露",b."血糖",b."血氧饱和度",b."羊水量",b."羊水性状",b."处理",b."记录人",b."剂量",b."滴速",b."宫颈管长度"
From Sol_Rs_Druglabor_List a,
     Json_Table(a.Content,'$' Columns(
     记录时间 Varchar2(20) Path '$.记录时间',
     剂量 Number(3,1) Path '$.剂量',
     滴速 Number(3) Path '$.滴速',
     血压 Varchar2(7) Path '$.血压',
     收缩压 Varchar2(7) Path '$.收缩压',
     舒张压 Varchar2(7) Path '$.舒张压',
     脉搏 Number(3) Path '$.脉搏',
     胎心率 Number(3) Path '$.胎心率',
     宫缩强度 Varchar2(10) Path '$.宫缩强度',
     宫缩持续 Varchar2(20) Path '$.宫缩持续',
     宫缩间隔 Varchar2(20) Path '$.宫缩间隔',
     宫颈管长度 Varchar2(20) Path '$.宫颈管长度',
     宫口 Number(3) Path '$.宫口',
     意识 Varchar2(20) Path  '$.意识',
     体温 Varchar2(20) Path  '$.体温',
     呼吸 Varchar2(20) Path  '$.呼吸',
     先露 Varchar2(10) Path '$.先露',
     血糖 Varchar2(10) Path '$.血糖',
     血氧饱和度 Varchar2(10) Path '$.血氧饱和度',

     羊水量 Number(4) Path '$.羊水量',
     羊水性状 Varchar2(10) Path '$.羊水性状',
     处理 Varchar2(500) Path '$.处理',
     记录人 Varchar2(100) Path '$.记录人')) b;

create or replace force view ZLSOL_JC.v_sol_rs_expectant as
Select a.mid,a.courseid,b."记录时间",b."脉搏",b."血压",b."收缩压",b."舒张压",b."体温",b."呼吸",b."血糖",b."宫高",b."腹围",b."胎动计数早",b."胎动计数中",b."胎动计数晚",b."胎心率",b."先露",b."宫口",b."宫颈管消失",b."破膜情况",b."羊水性状",b."宫缩强度",b."宫缩持续",b."宫缩间隔",b."处理",b."检查者"
From SOL_RS_EXPECTANT a,JSON_TABLE(a.Content,'$' columns(
记录时间    Varchar2(50) PATH '$.记录时间',
脉搏  Varchar2(20) PATH '$.脉搏',
血压     Varchar2(20) PATH '$.血压',
收缩压   Varchar2(20) PATH '$.收缩压',
舒张压   Varchar2(20) PATH '$.舒张压',

体温   Varchar2(20) PATH '$.体温',
呼吸   Varchar2(20) PATH '$.呼吸',

血糖     Varchar2(20) PATH '$.血糖',
宫高     Number(4,2) PATH '$.宫高',
腹围     Varchar2(20) PATH '$.腹围',
胎动计数早     Number(3) PATH '$.胎动计数早',
胎动计数中     Number(3) PATH '$.胎动计数中',
胎动计数晚   Number(3) PATH '$.胎动计数晚',
胎心率 Number(3) PATH '$.胎心率',
先露     Varchar2(20) PATH '$.先露',
宫口     Varchar2(20) PATH '$.宫口',
宫颈管消失     Varchar2(20) PATH '$.宫颈管消失',
破膜情况     Varchar2(20) PATH '$.破膜情况',
羊水性状      Varchar2(20) PATH '$.羊水性状',
宫缩强度     Varchar2(20) PATH '$.宫缩强度',
宫缩持续       Varchar2(20) PATH '$.宫缩持续',
宫缩间隔       Varchar2(20) PATH '$.宫缩间隔',
处理     Varchar2(500) PATH '$.处理',
检查者       Varchar2(20) PATH '$.检查者'
)) as b;

create or replace force view ZLSOL_JC.v_sol_rs_postpartum as
Select a.Mid, 分娩日期, 入产房时间, 分娩方式, 出产房时间, 出产房时bp, 出产房时脉搏, 出产房时宫高脐下, 出产房时阴道流血, 出产房时一般情况, 会阴,  拆线
From Sol_Rs_Postpartum A,
     Json_Table(a.Content,
                 '$' Columns(分娩日期 varchar2(20) Path '$.分娩日期', 入产房时间 varchar2(20) Path '$.入产房时间', 分娩方式 Varchar2(20) Path '$.分娩方式',
                          出产房时间 varchar2(20) Path '$.出产房时间', 出产房时bp varchar2(7) Path '$.出产房时BP', 出产房时脉搏 Number(3) Path '$.出产房时脉搏',
                          出产房时宫高脐下 Number(2) Path '$.出产房时宫高脐下', 出产房时阴道流血 Number(3) Path '$.出产房时阴道流血',
                          出产房时一般情况 Varchar2(10) Path '$.出产房时一般情况', 会阴 Varchar2(20) Path '$.会阴', 拆线 Varchar2(10) Path '$.拆线'));

create or replace force view ZLSOL_JC.v_sol_rs_postpartum_jcfy_list as
Select a.Mid, a.Courseid, 记录时间, 意识, 体温, 脉搏, 呼吸, 收缩压,舒张压, 血氧饱和度, 血糖, 尿量, 阴道出血, 宫底高度,   特殊情况及处理, 签名
From Sol_Rs_Postpartum_List A,
     Json_Table(a.Content,
                 '$' Columns(记录时间 Varchar2(30) Path '$.记录时间', 意识 Varchar2(20) Path '$.意识', 体温 Varchar2(10) Path '$.体温',
                          脉搏 Varchar2(50) Path '$.脉搏', 呼吸 Varchar2(30) Path '$.呼吸', 收缩压 Varchar2(50) Path '$.收缩压',舒张压 Varchar2(50) Path '$.舒张压',
                          血氧饱和度 Varchar2(10) Path '$.血氧饱和度', 血糖 Varchar2(20) Path '$.血糖', 尿量 Varchar2(20) Path '$.尿量',
                          阴道出血 Varchar2(10) Path '$.阴道出血', 宫底高度 Varchar2(30) Path '$.宫底高度', 特殊情况及处理 Varchar2(500) Path '$.特殊情况及处理',
                          签名 Varchar2(100) Path '$.签名'));

create or replace force view ZLSOL_JC.v_sol_rs_postpartum_list as
Select a.Mid, a.Courseid ID, 记录时间, 乳量, 乳房红肿, 乳头, 子宫宫高, 子宫压痛, 恶露量, 恶露颜色, 恶露臭味, 会阴正常, 会阴红肿, 会阴其他, 小便, 大便, 特殊情况, 签名
From Sol_Rs_Postpartum_List A,
     Json_Table(a.Content,
                 '$' Columns(记录时间 Varchar2(20) Path '$.记录时间', 乳量 Number(4) Path '$.乳量', 乳房红肿 Varchar2(10) Path '$.乳房红肿',
                          乳头 Varchar2(50) Path '$.乳头', 子宫宫高 Number(3) Path '$.子宫宫高', 子宫压痛 Varchar2(50) Path '$.子宫压痛',
                          恶露量 Number(4) Path '$.恶露量', 恶露颜色 Varchar2(20) Path '$.恶露颜色', 恶露臭味 Varchar2(20) Path '$.恶露臭味',
                          会阴正常 Varchar2(10) Path '$.会阴正常', 会阴红肿 Varchar2(10) Path '$.会阴红肿', 会阴其他 Varchar2(50) Path '$.会阴其他',
                          小便 Varchar2(50) Path '$.小便', 大便 Varchar2(50) Path '$.大便', 特殊情况 Varchar2(100) Path '$.特殊情况',
                          签名 Varchar2(100) Path '$.签名'));

CREATE OR REPLACE TYPE ZLSOL_JC."T_STRLIST"                                          as Table of Varchar2(4000)
/

CREATE OR REPLACE Function ZLSOL_JC.f_Str2list
(
  Str_In   In Varchar2,
  Split_In In Varchar2 := ','
) Return t_Strlist
  Pipelined As
  v_Str Long;
  P     Number;
  --功能：将由逗号分隔的不带引号的字符序列转换为单列数据表
  --参数：STR_IN,如:G0000123,G0000124,G0000125...,SPLIT_IN,分隔符,缺省为,号
  --说明：
  --1．当SQL语句中涉及“IN(常量1, 常量2,…) ”子句时使用这种方式以便利用绑定变量。
  --2．使用这两个函数时，需要在SQL语句中加入“/*+ cardinality(b 3)*/”提示，因为CBO下临时内存表没有统计数据,。
  --3．两种调用示例
  --SELECT /*+ cardinality(b 3)*/ * FROM 门诊费用记录 WHERE NO IN (SELECT * FROM TABLE(F_STR2LIST('A01,A02,A03')) B);
  --SELECT /*+ cardinality(b 3)*/ A.* FROM 门诊费用记录 A, TABLE(F_STR2LIST('A01,A02,A03')) B WHERE A.NO = B.COLUMN_VALUE;
Begin
  If Str_In Is Null Then
    Return;
  End If;
  v_Str := Str_In || Split_In;
  Loop
    P := Instr(v_Str, Split_In);
    Exit When(Nvl(P, 0) = 0);
    Pipe Row(Substr(v_Str, 1, P - 1));
    v_Str := Substr(v_Str, P + 1);
  End Loop;
  Return;
End;
/

CREATE OR REPLACE Function ZLSOL_JC.Zl_Split
(
  Expression In Varchar2, --需要分割字符串
  Delimiter  In Varchar2, --分割字符串
  Mimit      In Number := -1 --分割位
)
--功能：通过函数实现字符串分割，根据传入分割位实现分割字符。
  --参数：Expression【需要分割字符串】、Delimiter【分割字符】、Mimit【分割后的子串数】
  --返回：返回分割位置的字符串
  --程序：谢荣
  --日期：2010-12-24
  --修改：
  --修改日期
 Return Varchar2 --返回分割字符串
 As
  Intmimit      Number; --Mimit
  Strexpression Varchar2(4000);
  Strdelimiter  Varchar2(4000);
  Strresult     Varchar2(4000);
  Strtmp        Varchar2(4000);
  Inttmp        Number(3);
  v_Error       Varchar2(255);
  Err_Custom Exception;
Begin
  Strexpression := Expression;
  Strdelimiter  := Delimiter;
  Intmimit      := Mimit;
  Strtmp        := Strexpression || Strdelimiter;
  Inttmp        := 0;
  While Inttmp < Intmimit Loop
    If Instr(Strtmp, Strdelimiter) > 0 Then
      Strtmp := Substr(Strtmp, Lengthb(Strdelimiter) + Instr(Strtmp, Strdelimiter));
    Else
      Strtmp  := '';
      v_Error := '下标值越界！';
      --Raise Err_Custom;
    End If;
    Inttmp := Inttmp + 1;
  End Loop;
  Strtmp    := Substr(Strtmp, 1, Instr(Strtmp, Strdelimiter) - 1);
  Strresult := Strtmp;
  If Mimit = -1 Then
    Strresult := Expression;
  End If;
  Return Strresult;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20999, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
End;
/

CREATE OR REPLACE Procedure ZLSOL_JC.his_病人新生儿登记_revise
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
  select count(1) into n_count from 病人新生儿记录@to_his where 病人id=n_病人id and 主页id = n_主页id and 序号=babyno_In;
  --新生儿新增修改
  if  state_in=2 then
   if n_count=0 then  ----新增
      insert into   his_病人新生儿记录 
      (病人id,主页id,序号,婴儿姓名,婴儿性别,分娩次数,分娩方式,胎儿状况,出生时间,身长,体重,血型,备注说明,死亡时间,登记时间,登记人)
      select 病人id,住院次数,序号,婴儿姓名,婴儿性别,分娩次数,娩出方式,胎儿状况,出生时间,身长,体重,血型,备注说明,死亡时间,登记时间,登记人
       from   ( select  d.pid 病人id,d.tid 住院次数,b.Babyno 序号,d.name||decode(b.Sex,'男','之子','之女')||t.顺序 as 婴儿姓名,
              b.Sex as 婴儿性别,c.妊次 as 分娩次数,a.娩出方式,b.胎儿状况,b.身长,b.体重,b.血型,to_date(a.胎儿娩出时间,'yyyy-mm-dd hh24:mi:ss') as 出生时间,
              b.死亡时间,'' as 备注说明,b.Recorder 登记人,b.Addtime 登记时间
              from V_SOL_INF_DELIVERY a,v_newborn b,v_sol_rs_birth c,sol_inf_puerpera d,
              (select mid,sex,babyno,row_number() over(partition by mid,sex order by babyno) 顺序 from SOL_INF_NEWBORNS t  ) t
              where a.Mid=b.Mid and a.mid=d.mid and a.mid=t.mid and b.Babyno=t.babyno
              and b.Mid=c.mid(+) and a.mid=mid_In
                ) where 序号=babyno_In;
                
      select count(*) into n_count from his_病人新生儿记录;
      dbms_output.put_line(n_count);
      
      insert into 病人新生儿记录@to_his value select * from his_病人新生儿记录 ;
      Zl_病区自动标记_Update@To_His(n_病人id, n_主页id); 
      b_Message.Zlhis_Patient_011@To_His(n_病人id, n_主页id, babyno_In);
      delete from his_病人新生儿记录;
    else  ----修改
      insert into   his_病人新生儿记录 
      (病人id,主页id,序号,婴儿姓名,婴儿性别,分娩次数,分娩方式,胎儿状况,出生时间,身长,体重,血型,备注说明,死亡时间,登记时间,登记人)
      select 病人id,住院次数,序号,婴儿姓名,婴儿性别,分娩次数,娩出方式,胎儿状况,出生时间,身长,体重,血型,备注说明,死亡时间,登记时间,登记人
       from   ( select  d.pid 病人id,d.tid 住院次数,b.Babyno 序号,d.name||decode(b.Sex,'男','之子','之女')||t.顺序 as 婴儿姓名,
              b.Sex as 婴儿性别,c.妊次 as 分娩次数,a.娩出方式,b.胎儿状况,b.身长,b.体重,b.血型,to_date(a.胎儿娩出时间,'yyyy-mm-dd hh24:mi:ss') as 出生时间,
              b.死亡时间,'' as 备注说明,b.Recorder 登记人,b.Addtime 登记时间
              from V_SOL_INF_DELIVERY a,v_newborn b,v_sol_rs_birth c,sol_inf_puerpera d,
              (select mid,sex,babyno,row_number() over(partition by mid,sex order by babyno) 顺序 from SOL_INF_NEWBORNS t  ) t
              where a.Mid=b.Mid and a.mid=d.mid and a.mid=t.mid and b.Babyno=t.babyno
              and b.Mid=c.mid(+) and a.mid=mid_In
                ) where 序号=babyno_In;
        delete from 病人新生儿记录@to_his where 病人id=n_病人id and 主页id=n_主页id and 序号=babyno_In;
        insert into 病人新生儿记录@to_his value select * from his_病人新生儿记录 ;
        Zl_病区自动标记_Update@To_His(n_病人id, n_主页id); 
      b_Message.Zlhis_Patient_011@To_His(n_病人id, n_主页id, babyno_In);
   /* update  病人新生儿记录@to_his set  
   (序号,婴儿姓名,婴儿性别,分娩次数,分娩方式,胎儿状况,出生时间,身长,体重,血型,备注说明,死亡时间,登记时间,登记人)=
   (select 序号,婴儿姓名,婴儿性别,分娩次数,分娩方式,胎儿状况,出生时间,身长,体重,血型,备注说明,死亡时间,登记时间,登记人
    from   his_病人新生儿记录) ;*/
     delete from his_病人新生儿记录;          
      end if;
 
      
     --新生儿登记删除
   elsif state_in=3 then
     delete from 病人新生儿记录@to_his where 病人id=n_病人id and 主页id=n_主页id and 序号=babyno_In;
     Zl_病区自动标记_Update@To_His(n_病人id,n_主页id); 
 
     b_Message.ZLHIS_PATIENT_013@To_His(n_病人id,n_主页id,babyno_In); 

  End If;

 
   
End his_病人新生儿登记_revise;
/

