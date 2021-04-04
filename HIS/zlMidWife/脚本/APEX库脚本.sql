--��APEX��ִ�У��޸����룬ip��ʵ����[SERVICE_NAME]��
create database link ZLHIS_DBL  connect to zlhis identified by zlhis_password  using '(DESCRIPTION =
    (ADDRESS = (PROTOCOL = TCP)(HOST = 192.168.0.60)(PORT = 1521))
    (CONNECT_DATA =
      (SERVICE_NAME = dyey)
    )
  )';
create table ZLSOL.HIS_������������¼
(
  ����id NUMBER(18) not null,
  ��ҳid NUMBER(18) not null,
  ���   NUMBER(3) not null,
  Ӥ������ VARCHAR2(100),
  Ӥ���Ա� VARCHAR2(4),
  ������� NUMBER(3),
  ���䷽ʽ VARCHAR2(20),
  ̥��״�� VARCHAR2(20),
  ����ʱ�� DATE,
  ��   NUMBER(16,5),
  ����   NUMBER(16,5),
  Ѫ��   VARCHAR2(10),
  ��ע˵�� VARCHAR2(100),
  ����ʱ�� DATE,
  �Ǽ�ʱ�� DATE,
  �Ǽ���  VARCHAR2(20)
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
  is '������Ϣ';
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
  is 'ϵͳ�û���Ϣ';
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
  is '������Ϣ';
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
  is '������Ϣ';
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
  is '��е����¼';

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
  is '��������Ϣ';
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
  is '��ǰ�����Ϣ';
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
  is '���̾���';
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
  ����   DATE,
  ����ָ�� VARCHAR2(100),
  �������� VARCHAR2(100)
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
  is 'ҩ��������Ϣ';
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
  is 'ҩ��������¼';
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
  is '������¼';
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
  is '����۲���Ϣ';
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
  is '����۲��¼';
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
Select a.Mid, b."����1",b."���̿�ʼʱ��",b."����ȫ��ʱ��",b."̥�����ʱ��",b."̥�����ʱ��",b."��һ����",b."�ڶ�����",b."��������",b."�������",b."�������������¾���",b."����",b."��Ĥ��ʽ",b."��Ĥʱ��",b."��ˮ��״",b."��ˮ��",b."��ˮ��ɫ",b."̥�������ʽ",b."̥�̰��뷽ʽ",b."̥��������",b."̥��̥Ĥ����",b."̥�����",b."̥����̬",b."̥�̴�С",b."̥������",b."�������",b."�������",b."����ƾ�",b."�����ٽ�",b."����Ѵ�",b."�����ʽ",b."���̥��λ",b."������С",b."������λ",b."�������˳̶�",b."���������п�",b."�������˷��",b."������������",b."�������˳���",b."�������˲�λ",b."��������״��",b."�������˲�λ��С",b."��������Ѫ�״�С",b."�������Ա�",b."����������",b."��������",b."��������������",b."����������������",b."������������������״",b."��������������ҩ��",b."���������Ȼ���",b."������������̥",b."��������������",b."����Ѫѹ",b."������Ѫ",b."��ʱ��ҩ",b."������ҩ",b."�������", d."����3",d."������ʱ��",d."������",d."������",d."��¼��"
From Sol_Inf_Delivery A,
     Json_Table(Nvl(a.Newborndetail, '{"����1":"1"}'),'$' Columns(����1 Varchar2(50) Path '$.����1',
                  ���̿�ʼʱ�� Varchar2(19) Path '$.���̿�ʼʱ��',����ȫ��ʱ�� Varchar2(19) Path '$.����ȫ��ʱ��',
                          ̥�����ʱ�� Varchar2(19) Path '$.̥�����ʱ��',̥�����ʱ�� Varchar2(19) Path '$.̥�����ʱ��',
                          ��һ���� Varchar2(50) Path '$.��һ����', �ڶ����� Varchar2(50) Path '$.�ڶ�����',�������� Varchar2(50) Path '$.��������',
                          ������� Varchar2(50) Path '$.�������', �������������¾��� Varchar2(50) Path '$.�������������¾���', ���� Varchar2(50) Path '$.����',
                          ��Ĥ��ʽ Varchar2(50) Path '$.��Ĥ��ʽ', ��Ĥʱ�� Varchar2(19) Path '$.��Ĥʱ��',
                          ��ˮ��״ Varchar2(50) Path '$.��ˮ��״', ��ˮ�� Varchar2(50) Path '$.��ˮ��',��ˮ��ɫ Varchar2(50) Path '$.��ˮ��ɫ',
                          ̥�������ʽ Varchar2(50) Path '$.̥�������ʽ',̥�̰��뷽ʽ Varchar2(50) Path '$.̥�̰��뷽ʽ',
                          ̥�������� Varchar2(50) Path '$.̥��������', ̥��̥Ĥ���� Varchar2(50) Path '$.̥��̥Ĥ����',
                          ̥����� Varchar2(50) Path '$.̥�����', ̥����̬ Varchar2(50) Path '$.̥����̬',
                          ̥�̴�С Varchar2(50) Path '$.̥�̴�С',̥������ Varchar2(50) Path '$.̥������',
                          ������� Varchar2(50) Path '$.�������', ������� Varchar2(50) Path '$.�������',
                          ����ƾ� Varchar2(50) Path '$.����ƾ�', �����ٽ� Varchar2(50) Path '$.�����ٽ�', ����Ѵ� Varchar2(50) Path '$.����Ѵ�',
                          �����ʽ Varchar2(50) Path '$.�����ʽ',���̥��λ Varchar2(50) Path '$.���̥��λ',
                          ������С Varchar2(50) Path '$.������С',������λ Varchar2(50) Path '$.������λ',
                          �������˳̶� Varchar2(50) Path '$.�������˳̶�',
                          ���������п� Varchar2(50) Path '$.���������п�', �������˷�� Varchar2(50) Path '$.�������˷��',
                          ������������ Varchar2(50) Path '$.������������', �������˳��� Varchar2(50) Path '$.�������˳���',
                          �������˲�λ Varchar2(50) Path '$.�������˲�λ', ��������״�� Varchar2(50) Path '$.��������״��',
                          �������˲�λ��С Varchar2(50) Path '$.�������˲�λ��С', ��������Ѫ�״�С Varchar2(50) Path '$.��������Ѫ�״�С',
                          �������Ա� Varchar2(50) Path '$.�������Ա�', ���������� Varchar2(50) Path '$.����������',
                          �������� Varchar2(50) Path '$.��������', �������������� Varchar2(50) Path '$.��������������',
                          ���������������� Varchar2(50) Path '$.����������������', ������������������״ Varchar2(50) Path '$.������������������״',
                          ��������������ҩ�� Varchar2(50) Path '$.��������������ҩ��', ���������Ȼ��� Varchar2(50) Path '$.���������Ȼ���',
                          ������������̥ Varchar2(50) Path '$.������������̥', �������������� Varchar2(50) Path '$.��������������',
                          ����Ѫѹ Varchar2(50) Path '$.����Ѫѹ', ������Ѫ Varchar2(50) Path '$.������Ѫ', ��ʱ��ҩ Varchar2(50) Path '$.��ʱ��ҩ',
                          ������ҩ Varchar2(50) Path '$.������ҩ', ������� Varchar2(50) Path '$.�������')) As B,
      Json_Table(Nvl(a.Deliveryinf, '{"����3":"1"}'),
                 '$' Columns(����3 Varchar2(1) Path '$.����3', ������ʱ�� Varchar2(50) Path '$.������ʱ��',
                          ������ Varchar2(50) Path '$.������', ������ Varchar2(50) Path '$.������', ��¼�� Varchar2(50) Path '$.��¼��')) As D;

create or replace force view ZLSOL.v_newborn as
Select a.Mid, b.Bid, b.Babyno, b.Addtime, b.Recorder, a.Name, a.Old, a.Bedno, a.Pno, b.Sex, d."����2",d."��",d."����",d."ͷΧ",d."��Χ",d."һ�������Ӧ",d."һ�������ɫ",d."һ�����Ƥ��",d."һ������ë",d."ͷ������",d."­���ص�",d."̥ͷˮ��Ѫ��",d."̥ͷˮ�״�С",d."ǰض",d."����",d."����",d."��ǻ",d."��",d."���",d."��",d."Ƣ",d."��֫",d."��չ����",d."����",d."��ֳ��", e."����3",e."����1����",e."����5����",e."����10����",e."����1����",e."����5����",e."����10����",e."����1����",e."����5����",e."����10����",e."������1����",e."������5����",e."������10����",e."��ɫ1����",e."��ɫ5����",e."��ɫ10����",e."�ܷ�1����",e."�ܷ�5����",e."�ܷ�10����", f."����4",f."�����ڲ�ʱ�ϲ�֢����ҩ���",f."����ǰ̥�����",f."Ӥ������ʱ�������",f."����ȱ��",f."ĸ��ι��ָ��",f."���"
From Sol_Inf_Puerpera A, Sol_Inf_Newborns B,
     Json_Table(Nvl(b.Newborninf, '{"����2":"2"}'),
                 '$' Columns(����2 Varchar2(50) Path '$.����2', �� Varchar2(50) Path '$.��', ���� Varchar2(50) Path '$.����',
                          ͷΧ Varchar2(50) Path '$.ͷΧ', ��Χ Varchar2(50) Path '$.��Χ', һ�������Ӧ Varchar2(50) Path '$.һ�������Ӧ',
                          һ�������ɫ Varchar2(50) Path '$.һ�������ɫ', һ�����Ƥ�� Varchar2(50) Path '$.һ�����Ƥ��',
                          һ������ë Varchar2(50) Path '$.һ������ë', ͷ������ Varchar2(50) Path '$.ͷ������',
                          ­���ص� Varchar2(50) Path '$.­���ص�', ̥ͷˮ��Ѫ�� Varchar2(50) Path '$.̥ͷˮ��Ѫ��',
                          ̥ͷˮ�״�С Varchar2(50) Path '$.̥ͷˮ�״�С', ǰض Varchar2(50) Path '$.ǰض', ���� Varchar2(50) Path '$.����',
                          ���� Varchar2(50) Path '$.����', ��ǻ Varchar2(50) Path '$.��ǻ', �� Varchar2(50) Path '$.��',
                          ��� Varchar2(50) Path '$.���', �� Varchar2(50) Path '$.��', Ƣ Varchar2(50) Path '$.Ƣ',
                          ��֫ Varchar2(50) Path '$.��֫', ��չ���� Varchar2(50) Path '$.��չ����', ���� Varchar2(50) Path '$.����',
                          ��ֳ�� Varchar2(50) Path '$.��ֳ��')) As D,
     Json_Table(Nvl(b.Newbornscore, '{"����3":"3"}'),
                 '$' Columns(����3 Varchar2(50) Path '$.����3', ����1���� Varchar2(50) Path '$.����1����',
                          ����5���� Varchar2(50) Path '$.����5����', ����10���� Varchar2(50) Path '$.����10����',
                          ����1���� Varchar2(50) Path '$.����1����', ����5���� Varchar2(50) Path '$.����5����',
                          ����10���� Varchar2(50) Path '$.����10����', ����1���� Varchar2(50) Path '$.����1����',
                          ����5���� Varchar2(50) Path '$.����5����', ����10���� Varchar2(50) Path '$.����10����',
                          ������1���� Varchar2(50) Path '$.������1����', ������5���� Varchar2(50) Path '$.������5����',
                          ������10���� Varchar2(50) Path '$.������10����', ��ɫ1���� Varchar2(50) Path '$.��ɫ1����',
                          ��ɫ5���� Varchar2(50) Path '$.��ɫ5����', ��ɫ10���� Varchar2(50) Path '$.��ɫ10����',
                          �ܷ�1���� Varchar2(50) Path '$.�ܷ�1����', �ܷ�5���� Varchar2(50) Path '$.�ܷ�5����',
                          �ܷ�10���� Varchar2(50) Path '$.�ܷ�10����')) As E,
     Json_Table(Nvl(b.Otherinf, '{"����4":"4"}'),
                 '$' Columns(����4 Varchar2(50) Path '$.����4', �����ڲ�ʱ�ϲ�֢����ҩ��� Varchar2(50) Path '$.�����ڲ�ʱ�ϲ�֢����ҩ���  ',
                          ����ǰ̥����� Varchar2(50) Path '$.����ǰ̥�����  ', Ӥ������ʱ������� Varchar2(50) Path '$.Ӥ������ʱ�������  ',
                          ����ȱ�� Varchar2(50) Path '$.����ȱ��  ', ĸ��ι��ָ�� Varchar2(50) Path '$.ĸ��ι��ָ��  ',
                          ��� Varchar2(50) Path '$.���  ')) As F
Where a.Mid = b.Mid
Order By Addtime;

create or replace force view ZLSOL.v_sol_inf_checkinroom as
Select a.mid, b."�뷿Ŀ��",b."�뷿ʱ��",b."ҽ�Ʋ���",b."������",b."����֪��֪ͨ��",b."����������",b."̥����",b."̥�Ĵ���",b."��Ĥ���",b."�Ƿ��кϲ�֢",b."����",b."��Һ��",b."����ͨ��",b."�ֲ����",b."����ҩ��",b."����",b."������",b."�Ӱ���"
From SOL_INF_CHECKINROOM a,JSON_TABLE(a.Content,'$' columns(
�뷿Ŀ��       Varchar2(50) PATH '$.�뷿Ŀ��',
�뷿ʱ��       Varchar2(50) PATH '$.�뷿ʱ��',
ҽ�Ʋ���       Varchar2(10) PATH '$.ҽ�Ʋ���',
������       Varchar2(10) PATH '$.������',
����֪��֪ͨ��      Varchar2(10) PATH '$.����֪��֪ͨ��',
����������   Varchar2(20) PATH '$.����������',
̥����        Varchar2(10) PATH '$.̥����',
̥�Ĵ���       Varchar2(10) PATH '$.̥�Ĵ���',
��Ĥ���       Varchar2(10) PATH '$.��Ĥ���',
�Ƿ��кϲ�֢       Varchar2(10) PATH '$.�Ƿ��кϲ�֢',
����         Varchar2(50) PATH '$.����',
��Һ��        Varchar2(10) PATH '$.��Һ��',
����ͨ��       Varchar2(10) PATH '$.����ͨ��',
�ֲ����       Varchar2(50) PATH '$.�ֲ����',
����ҩ��       Varchar2(50) PATH '$.����ҩ��',
����         Varchar2(50) PATH '$.����',
������        Varchar2(50) PATH '$.������',
�Ӱ���               Varchar2(50) PATH '$.�Ӱ���'
)) as b;

create or replace force view ZLSOL.v_sol_inf_checkoutroom as
Select a.mid, b."OUTROOMTIME",b."����״̬",b."ҽ�Ʋ���",b."������",b."����ͨ��",b."�ֲ����",b."��������",b."�����п���",b."�����пڷ��",b."����ˮ��",b."����Ѫ��",b."�����Ѫ",b."��Ѫ��",b."����ҩ��",b."������",b."�Ӱ���",b."ҩ��",b."��ע"
From SOL_INF_CheckOutRoom a,JSON_TABLE(a.Content,'$' columns(
OUTROOMTIME     Varchar2(50) PATH '$.OUTROOMTIME',
����״̬     Varchar2(50) PATH '$.����״̬',
ҽ�Ʋ���     Varchar2(10) PATH '$.ҽ�Ʋ���',
������     Varchar2(10) PATH '$.������',
����ͨ��     Varchar2(10) PATH '$.����ͨ��',
�ֲ����     Varchar2(50) PATH '$.�ֲ����',
��������     Varchar2(20) PATH '$.��������',
�����п���   Varchar2(20) PATH '$.�����п���',
�����пڷ�� Varchar2(10) PATH '$.�����пڷ��',
����ˮ��     Varchar2(10) PATH '$.����ˮ��',
����Ѫ��     Varchar2(10) PATH '$.����Ѫ��',
�����Ѫ     Varchar2(10) PATH '$.�����Ѫ',
��Ѫ��       Number(5) PATH '$.��Ѫ��',
����ҩ��     Varchar2(50) PATH '$.����ҩ��',
������       Varchar2(20) PATH '$.������',
�Ӱ���       Varchar2(20) PATH '$.�Ӱ���',
ҩ��         Varchar2(50) PATH '$.ҩ��',
��ע         Varchar2(50) PATH '$.��ע'
)) as b;

create or replace force view ZLSOL.v_sol_inf_delivery as
Select a.Mid, b."����1", b."���̿�ʼʱ��", b."����ȫ��ʱ��", b."̥�����ʱ��", b."̥�����ʱ��", b."��һ����", b."�ڶ�����", b."��������", b."�������",
       b."��������������", b."����", b."��Ĥ��ʽ", b."��Ĥʱ��", b."��ˮ��״", b."��ˮ��", b."��ˮ��ɫ", b."̥�������ʽ", b."̥�̰��뷽ʽ", b."̥��������",
       b."̥��̥Ĥ����", b."̥�����", b."̥����̬", b."̥�̴�С", b."̥������", b."�������", b."�������", b."�����ٽ�", b."���",b."�ƾ�����", b."�����ʽ",
       b."���̥��λ", b."������С", b."������λ", b."�������˳̶�", b."���������п�", b."�������˷��", b."������������", b."�������˳���", b."�������˲�λ", b."��������״��",
       b."�������˲�λ��С", b."��������Ѫ�״�С", b."��������������", b."����������������", b."������������������״",
       b."��������������ҩ��", b."���������Ȼ���", b."������������̥", b."��������������",b."ĸӤ��Ӵ�����˱ʱ��", b."����Ѫѹ", b."������Ѫ", b."��ʱ��ҩ", b."������ҩ",
       b."�������", b."�������", d."����3",d."������ʱ��", d."������", d."������", d."��¼��"
From Sol_Inf_Delivery A,
     Json_Table(Nvl(a.Newborndetail, '{"����1":"1"}'),
                 '$'
                  Columns(����1 Varchar2(50) Path '$.����1', ���̿�ʼʱ�� Varchar2(19) Path '$.���̿�ʼʱ��',
                          ����ȫ��ʱ�� Varchar2(19) Path '$.����ȫ��ʱ��', ̥�����ʱ�� Varchar2(19) Path '$.̥�����ʱ��',
                          ̥�����ʱ�� Varchar2(19) Path '$.̥�����ʱ��', ��һ���� Varchar2(50) Path '$.��һ����',
                          �ڶ����� Varchar2(50) Path '$.�ڶ�����', �������� Varchar2(50) Path '$.��������',
                          �������������� Varchar2(50) Path '$.��������������', ���� Varchar2(50) Path '$.����',
                          ��Ĥ��ʽ Varchar2(50) Path '$.��Ĥ��ʽ', ��Ĥʱ�� Varchar2(19) Path '$.��Ĥʱ��', ��ˮ��״ Varchar2(50) Path '$.��ˮ��״',
                          ��ˮ�� Varchar2(50) Path '$.��ˮ��', ��ˮ��ɫ Varchar2(50) Path '$.��ˮ��ɫ',
                          ̥�������ʽ Varchar2(50) Path '$.̥�������ʽ', ̥�̰��뷽ʽ Varchar2(50) Path '$.̥�̰��뷽ʽ',
                          ̥�������� Varchar2(50) Path '$.̥��������', ̥��̥Ĥ���� Varchar2(50) Path '$.̥��̥Ĥ����',
                          ̥����� Varchar2(50) Path '$.̥�����', ̥����̬ Varchar2(50) Path '$.̥����̬', ̥�̴�С Varchar2(50) Path '$.̥�̴�С',
                          ̥������ Varchar2(50) Path '$.̥������', ������� Varchar2(50) Path '$.�������', ������� Varchar2(50) Path '$.�������',
                         �����ٽ� Varchar2(50) Path '$.�����ٽ�',��� Varchar2(50) Path '$.���', �ƾ����� Varchar2(50) Path '$.�ƾ�����',
                          �����ʽ Varchar2(50) Path '$.�����ʽ',
                          ���̥��λ Varchar2(50) Path '$.���̥��λ', ������С Varchar2(50) Path '$.������С',
                          ������λ Varchar2(50) Path '$.������λ', �������˳̶� Varchar2(50) Path '$.�������˳̶�',
                          ���������п� Varchar2(50) Path '$.���������п�', �������˷�� Varchar2(50) Path '$.�������˷��',
                          ������������ Varchar2(50) Path '$.������������', �������˳��� Varchar2(50) Path '$.�������˳���',
                          �������˲�λ Varchar2(50) Path '$.�������˲�λ', ��������״�� Varchar2(50) Path '$.��������״��',
                          �������˲�λ��С Varchar2(50) Path '$.�������˲�λ��С', ��������Ѫ�״�С Varchar2(50) Path '$.��������Ѫ�״�С',
                          �������������� Varchar2(50) Path '$.��������������', ���������������� Varchar2(50) Path '$.����������������',
                          ������������������״ Varchar2(50) Path '$.������������������״', ��������������ҩ�� Varchar2(50) Path '$.��������������ҩ��',
                          ���������Ȼ��� Varchar2(50) Path '$.���������Ȼ���', ������������̥ Varchar2(50) Path '$.������������̥',
                          �������������� Varchar2(50) Path '$.��������������',������� Varchar2(50) Path '$.�������', ĸӤ��Ӵ�����˱ʱ�� Varchar2(50) Path '$.ĸӤ��Ӵ�����˱ʱ��',
                          ����Ѫѹ Varchar2(50) Path '$.����Ѫѹ',
                          ������Ѫ Varchar2(50) Path '$.������Ѫ', ��ʱ��ҩ Varchar2(50) Path '$.��ʱ��ҩ', ������ҩ Varchar2(50) Path '$.������ҩ',
                          ������� Varchar2(50) Path '$.�������',������� Varchar2(50) Path '$.�������')) As B,
     Json_Table(Nvl(a.Deliveryinf, '{"����3":"1"}'),
                 '$' Columns(����3 Varchar2(1) Path '$.����3', ������ʱ�� Varchar2(50) Path '$.������ʱ��',
                          ������ Varchar2(50) Path '$.������', ������ Varchar2(50) Path '$.������', ��¼�� Varchar2(50) Path '$.��¼��')) As D;

create or replace force view ZLSOL.v_sol_inf_equipment as
Select a.mid, b."���м���ǰ",b."���м�����",b."���м�����",b."�������ǰ",b."���������",b."���������",b."ֹѪǯ��ǰ",b."ֹѪǯ����",b."ֹѪǯ����",b."������ǰ",b."��������",b."��������",b."��������ǰ",b."����������",b."����������",b."�������ǰ",b."����������",b."���������",b."ϴ�����ǰ",b."ϴ��������",b."ϴ�������",b."������ǰ",b."���������",b."��������",b."������ǰ",b."��������",b."��������",b."����ǯ��ǰ",b."����ǯ����",b."����ǯ����",b."������ǰ",b."��������",b."��������",b."�γײ�ǰ",b."�γ�����",b."�γײ���",b."����˹��ǰ",b."����˹����",b."����˹����",b."��ǰ��ǰ",b."��ǰ����",b."��ǰ����",b."ɴ����ǰ",b."ɴ������",b."ɴ������",b."��Բǯ��ǰ",b."��Բǯ����",b."��Բǯ����"
From SOL_INF_Equipment a,JSON_TABLE(a.Content,'$' columns(
���м���ǰ   Number(2) PATH '$.���м���ǰ',
���м�����   Number(2) PATH '$.���м�����',
���м�����   Number(2) PATH '$.���м�����',
�������ǰ   Number(2) PATH '$.�������ǰ',
���������   Number(2) PATH '$.���������',
���������   Number(2) PATH '$.���������',
ֹѪǯ��ǰ   Number(2) PATH '$.ֹѪǯ��ǰ',
ֹѪǯ����   Number(2) PATH '$.ֹѪǯ����',
ֹѪǯ����   Number(2) PATH '$.ֹѪǯ����',
������ǰ   Number(2) PATH '$.������ǰ',
��������   Number(2) PATH '$.��������',
��������   Number(2) PATH '$.��������',
��������ǰ   Number(2) PATH '$.��������ǰ',
����������   Number(2) PATH '$.����������',
����������   Number(2) PATH '$.����������',
�������ǰ   Number(2) PATH '$.�������ǰ',
����������   Number(2) PATH '$.����������',
���������   Number(2) PATH '$.���������',
ϴ�����ǰ   Number(2) PATH '$.ϴ�����ǰ',
ϴ��������   Number(2) PATH '$.ϴ��������',
ϴ�������   Number(2) PATH '$.ϴ�������',
������ǰ   Number(2) PATH '$.������ǰ',
���������   Number(2) PATH '$.���������',
��������   Number(2) PATH '$.��������',
������ǰ   Number(2) PATH '$.������ǰ',
��������   Number(2) PATH '$.��������',
��������   Number(2) PATH '$.��������',
����ǯ��ǰ   Number(2) PATH '$.����ǯ��ǰ',
����ǯ����   Number(2) PATH '$.����ǯ����',
����ǯ����   Number(2) PATH '$.����ǯ����',
������ǰ   Number(2) PATH '$.������ǰ',
��������   Number(2) PATH '$.��������',
��������   Number(2) PATH '$.��������',
�γײ�ǰ   Number(2) PATH '$.�γײ�ǰ',
�γ�����   Number(2) PATH '$.�γ�����',
�γײ���   Number(2) PATH '$.�γײ���',
����˹��ǰ   Number(2) PATH '$.����˹��ǰ',
����˹����   Number(2) PATH '$.����˹����',
����˹����   Number(2) PATH '$.����˹����',
��ǰ��ǰ   Number(2) PATH '$.��ǰ��ǰ',
��ǰ����   Number(2) PATH '$.��ǰ����',
��ǰ����   Number(2) PATH '$.��ǰ����',
ɴ����ǰ   Number(2) PATH '$.ɴ����ǰ',
ɴ������   Number(2) PATH '$.ɴ������',
ɴ������   Number(2) PATH '$.ɴ������',
��Բǯ��ǰ   Number(2) PATH '$.��Բǯ��ǰ',
��Բǯ����   Number(2) PATH '$.��Բǯ����',
��Բǯ����   Number(2) PATH '$.��Բǯ����'
)) as b;

create or replace force view ZLSOL.v_sol_inf_newborns as
Select a.Mid, b.Bid, b.Babyno, b.Addtime, b.Recorder, a.Name, a.Old, a.Bedno, a.Pno, b.Sex, d."����2",d."��",d."����",d."ͷΧ",d."��Χ",d."һ�������Ӧ",d."һ�������ɫ",d."һ�����Ƥ��",d."һ������ë",d."ͷ������",d."­���ص�",d."̥ͷˮ��Ѫ��",d."̥ͷˮ�״�С",d."ǰض",d."����",d."����",d."��ǻ",d."��",d."���",d."��",d."Ƣ",d."��֫",d."��չ����",d."����",d."��ֳ��", e."����3",e."����1����",e."����5����",e."����10����",e."����1����",e."����5����",e."����10����",e."����1����",e."����5����",e."����10����",e."������1����",e."������5����",e."������10����",e."��ɫ1����",e."��ɫ5����",e."��ɫ10����",e."�ܷ�1����",e."�ܷ�5����",e."�ܷ�10����", f."����4",f."�����ڲ�ʱ�ϲ�֢����ҩ���",f."����ǰ̥�����",f."Ӥ������ʱ�������",f."����ȱ��",f."ĸ��ι��ָ��",f."���"
From Sol_Inf_Puerpera A, Sol_Inf_Newborns B,
     Json_Table(Nvl(b.Newborninf, '{"����2":"2"}'),
                 '$' Columns(����2 Varchar2(50) Path '$.����2', �� Varchar2(50) Path '$.��', ���� Varchar2(50) Path '$.����',
                          ͷΧ Varchar2(50) Path '$.ͷΧ', ��Χ Varchar2(50) Path '$.��Χ', һ�������Ӧ Varchar2(50) Path '$.һ�������Ӧ',
                          һ�������ɫ Varchar2(50) Path '$.һ�������ɫ', һ�����Ƥ�� Varchar2(50) Path '$.һ�����Ƥ��',
                          һ������ë Varchar2(50) Path '$.һ������ë', ͷ������ Varchar2(50) Path '$.ͷ������',
                          ­���ص� Varchar2(50) Path '$.­���ص�', ̥ͷˮ��Ѫ�� Varchar2(50) Path '$.̥ͷˮ��Ѫ��',
                          ̥ͷˮ�״�С Varchar2(50) Path '$.̥ͷˮ�״�С', ǰض Varchar2(50) Path '$.ǰض', ���� Varchar2(50) Path '$.����',
                          ���� Varchar2(50) Path '$.����', ��ǻ Varchar2(50) Path '$.��ǻ', �� Varchar2(50) Path '$.��',
                          ��� Varchar2(50) Path '$.���', �� Varchar2(50) Path '$.��', Ƣ Varchar2(50) Path '$.Ƣ',
                          ��֫ Varchar2(50) Path '$.��֫', ��չ���� Varchar2(50) Path '$.��չ����', ���� Varchar2(50) Path '$.����',
                          ��ֳ�� Varchar2(50) Path '$.��ֳ��')) As D,
     Json_Table(Nvl(b.Newbornscore, '{"����3":"3"}'),
                 '$' Columns(����3 Varchar2(50) Path '$.����3', ����1���� Varchar2(50) Path '$.����1����',
                          ����5���� Varchar2(50) Path '$.����5����', ����10���� Varchar2(50) Path '$.����10����',
                          ����1���� Varchar2(50) Path '$.����1����', ����5���� Varchar2(50) Path '$.����5����',
                          ����10���� Varchar2(50) Path '$.����10����', ����1���� Varchar2(50) Path '$.����1����',
                          ����5���� Varchar2(50) Path '$.����5����', ����10���� Varchar2(50) Path '$.����10����',
                          ������1���� Varchar2(50) Path '$.������1����', ������5���� Varchar2(50) Path '$.������5����',
                          ������10���� Varchar2(50) Path '$.������10����', ��ɫ1���� Varchar2(50) Path '$.��ɫ1����',
                          ��ɫ5���� Varchar2(50) Path '$.��ɫ5����', ��ɫ10���� Varchar2(50) Path '$.��ɫ10����',
                          �ܷ�1���� Varchar2(50) Path '$.�ܷ�1����', �ܷ�5���� Varchar2(50) Path '$.�ܷ�5����',
                          �ܷ�10���� Varchar2(50) Path '$.�ܷ�10����')) As E,
     Json_Table(Nvl(b.Otherinf, '{"����4":"4"}'),
                 '$' Columns(����4 Varchar2(50) Path '$.����4', �����ڲ�ʱ�ϲ�֢����ҩ��� Varchar2(50) Path '$.�����ڲ�ʱ�ϲ�֢����ҩ���  ',
                          ����ǰ̥����� Varchar2(50) Path '$.����ǰ̥�����  ', Ӥ������ʱ������� Varchar2(50) Path '$.Ӥ������ʱ�������  ',
                          ����ȱ�� Varchar2(50) Path '$.����ȱ��  ', ĸ��ι��ָ�� Varchar2(50) Path '$.ĸ��ι��ָ��  ',
                          ��� Varchar2(50) Path '$.���  ')) As F
Where a.Mid = b.Mid
Order By Addtime;

create or replace force view ZLSOL.v_sol_inf_puerpera as
Select Name, Mid, Old, LPad(Bedno, 10) Bedno, Pno, Diagnosis, Status, Decode(Expectant, 1, '��', '') ����,
       Decode(Checkinroom, 1, '��', '') �뷿, Decode(Birth, 1, '��', '') �ٲ�, Decode(Druglabor, 1, '��', '') ����,
       Decode(Delivery, 1, '��', '') ����, Decode(Newborns, 1, '��', '') ������, Decode(Postpartum, 1, '��', '') ����,
       Decode(Checkoutroom, 1, '��', '') ����,Decode(Equipment, 1, '��', '') ��е,outtime,pid,tid
From Sol_Inf_Puerpera;

create or replace force view ZLSOL.v_sol_rs_birth as
Select a.mid,b."�Ѵ�",b."����",b."Ѫ��",b."��������ʷ",b."ĩ���¾�",b."Ԥ����",b."��ǰ�ϼ��侶",b."���ռ侶",b."���ǽ�ڼ侶",b."�����⾶",b."���ǻ���",b."���ǹؽ�",b."�����м�",b."������",b."����֢",b."��ǰ��¼����",b."���ʱ��",b."����ѹ",b."����ѹ",b."����",b."����",b."̥����",b."̥����С",b."����������",b."̥λ",b."��Ĥ���",b."��¶",b."����",b."�����",b."������ʼʱ��",b."��Ĥʱ��",b."��Ժ����"
From SOL_RS_BIRTH a,JSON_TABLE(Nvl(a.CONTENT,'{����:1}'),'$' columns(
�Ѵ�            Number(3)    PATH '$.�Ѵ�',
����            Number(3)    PATH '$.����',
Ѫ��            Varchar2(10) PATH '$.Ѫ��',
��������ʷ      Varchar2(50) PATH '$.��������ʷ',
ĩ���¾�        Varchar2(20) PATH '$.ĩ���¾�',
Ԥ����          Varchar2(20) PATH '$.Ԥ����',
��ǰ�ϼ��侶    Number(5) PATH '$.��ǰ�ϼ��侶',
���ռ侶        Number(5) PATH '$.���ռ侶',
���ǽ�ڼ侶    Number(5) PATH '$.���ǽ�ڼ侶',
�����⾶        Number(5) PATH '$.�����⾶',
���ǻ���        Varchar2(10) PATH '$.���ǻ���',
���ǹؽ�        Varchar2(10) PATH '$.���ǹؽ�',
�����м�        Varchar2(10) PATH '$.�����м�',
������          Varchar2(10) PATH '$.������',
����֢          Varchar2(100) PATH '$.����֢',
��ǰ��¼����    Varchar2(100) PATH '$.��ǰ��¼����',
���ʱ��        Varchar2(20) PATH '$.���ʱ��',
����ѹ        Number(3) PATH '$.����ѹ',
����ѹ        Number(3) PATH '$.����ѹ',
����            Number(4,2)  PATH '$.����',
����            Varchar2(10) PATH '$.����',
̥����            Varchar2(10) PATH '$.̥����',
̥����С        Number(5,2) PATH '$.̥����С',
����������      Varchar2(10) PATH '$.����������',
̥λ            Varchar2(10) PATH '$.̥λ',
��Ĥ���        Varchar2(10) PATH '$.��Ĥ���',
��¶            Varchar2(2) PATH '$.��¶',
����            Number(4,2) PATH '$.����',
�����          Varchar2(50) PATH '$.�����',
������ʼʱ��    Varchar2(20) PATH '$.������ʼʱ��',
��Ĥʱ��        Varchar2(20) PATH '$.��Ĥʱ��',
��Ժ����        Varchar2(100) PATH '$.��Ժ����'
)) as b;

create or replace force view ZLSOL.v_sol_rs_birth_course as
Select  a.courseid,a.mid,b."���ʱ��",b."�Ƿ��ʹ���",b."̥��λ",b."����ѹ",b."����ѹ",b."����",b."����",b."̥����",b."����ǿ��",b."��������",b."�������",b."������",b."����",b."��Ĥ���",b."��¶",b."����",b."�����"
From SOL_RS_BIRTH_COURSE a,JSON_TABLE(a.CONTENT,'$' columns(
���ʱ��        Varchar2(20)  PATH '$.���ʱ��',
�Ƿ��ʹ���    Varchar2(20)  PATH '$.�Ƿ��ʹ���',
̥��λ Varchar2(20)  PATH '$.̥��λ',
����ѹ        Number(3) PATH '$.����ѹ',
����ѹ        Number(3) PATH '$.����ѹ',
����        Number(4,2)  PATH '$.����',
����        Varchar2(10) PATH '$.����',
̥����        Varchar2(10) PATH '$.̥����',
����ǿ��    Varchar2(10) PATH '$.����ǿ��',
��������  Varchar2(10) PATH '$.��������',
�������  Varchar2(10) PATH '$.�������',
������    Varchar2(10) PATH '$.������',
����        Number(4,2) PATH '$.����',
��Ĥ���    Varchar2(10) PATH '$.��Ĥ���',
��¶        Number(2) PATH '$.��¶'��
����        Varchar2(50) PATH '$.����'��
�����      Varchar2(50) PATH '$.�����'
)) as b;

create or replace force view ZLSOL.v_sol_rs_druglabor as
Select Mid, To_Char(����, 'YYYY-MM-DD HH24:MI') ����, ����ָ��, �������� from Sol_Rs_Druglabor;

create or replace force view ZLSOL.v_sol_rs_druglabor_list as
Select a.Mid, a.Courseid ID, b."��¼ʱ��",b."����ѹ",b."����ѹ",b."����",b."̥����",b."����ǿ��",b."��������",b."�������",b."����",b."��¶",b."��ˮ��",b."��ˮ��״",b."����",b."��¼��",b."����",b."����"
From ZLSOL.Sol_Rs_Druglabor_List a,
     Json_Table(a.Content,'$' Columns(
     ��¼ʱ�� Varchar2(20) Path '$.��¼ʱ��',
     ���� Number(3,1) Path '$.����',
     ���� Number(3) Path '$.����',
     ����ѹ        Number(3) PATH '$.����ѹ',
     ����ѹ        Number(3) PATH '$.����ѹ',
     ���� Number(3) Path '$.����',
     ̥���� Number(3) Path '$.̥����',
     ����ǿ�� Varchar2(10) Path '$.����ǿ��',
     �������� Number(3) Path '$.��������',
     ������� Number(2) Path '$.�������',
     ���� Number(3) Path '$.����',
     ��¶ Varchar2(10) Path '$.��¶',
     ��ˮ�� Number(4) Path '$.��ˮ��',
     ��ˮ��״ Varchar2(10) Path '$.��ˮ��״',
     ���� Varchar2(100) Path '$.����',
     ��¼�� Varchar2(100) Path '$.��¼��')) b;

create or replace force view ZLSOL.v_sol_rs_expectant as
Select a.mid,a.courseid,b."��¼ʱ��",b."̥��λ",b."����ѹ",b."����ѹ",b."����",b."��Χ",b."̥��������",b."̥��������",b."̥��������",b."̥����",b."��¶",b."����",b."��Ĥ���",b."��ˮ��״",b."����ǿ��",b."��������",b."�������",b."����",b."�����"
From SOL_RS_EXPECTANT a,JSON_TABLE(a.Content,'$' columns(
��¼ʱ��    Varchar2(50) PATH '$.��¼ʱ��',
̥��λ  Varchar2(20) PATH '$.̥��λ',
����ѹ  Number(3)  PATH '$.����ѹ',
����ѹ  Number(3)  PATH '$.����ѹ',
����     Number(4,2) PATH '$.����',
��Χ     Varchar2(20) PATH '$.��Χ',
̥��������     Number(3) PATH '$.̥��������',
̥��������     Number(3) PATH '$.̥��������',
̥��������   Number(3) PATH '$.̥��������',
̥���� Number(3) PATH '$.̥����',
��¶     Varchar2(20) PATH '$.��¶',
����     Varchar2(20) PATH '$.����',
��Ĥ���     Varchar2(20) PATH '$.��Ĥ���',
��ˮ��״      Varchar2(20) PATH '$.��ˮ��״',
����ǿ��     Varchar2(20) PATH '$.����ǿ��',
��������       Varchar2(20) PATH '$.��������',
�������       Varchar2(20) PATH '$.�������',
����     Varchar2(100) PATH '$.����',
�����       Varchar2(20) PATH '$.�����'
)) as b;

create or replace force view ZLSOL.v_sol_rs_postpartum as
Select a.Mid, ��������, �����ʱ��, ���䷽ʽ, ������ʱ��, ������ʱbp, ������ʱ����, ������ʱ��������, ������ʱ������Ѫ, ������ʱһ�����, ����,  ����
From ZLSOL.Sol_Rs_Postpartum A,
     Json_Table(a.Content,
                 '$' Columns(�������� varchar2(20) Path '$.��������', �����ʱ�� varchar2(20) Path '$.�����ʱ��', ���䷽ʽ Varchar2(20) Path '$.���䷽ʽ',
                          ������ʱ�� varchar2(20) Path '$.������ʱ��', ������ʱbp varchar2(7) Path '$.������ʱBP', ������ʱ���� Number(3) Path '$.������ʱ����',
                          ������ʱ�������� Number(2) Path '$.������ʱ��������', ������ʱ������Ѫ Number(3) Path '$.������ʱ������Ѫ',
                          ������ʱһ����� Varchar2(10) Path '$.������ʱһ�����', ���� Varchar2(20) Path '$.����', ���� Varchar2(10) Path '$.����'));

create or replace force view ZLSOL.v_sol_rs_postpartum_list as
Select a.Mid, a.Courseid ID, ��¼ʱ��, ����, �鷿����, ��ͷ, �ӹ�����, �ӹ�ѹʹ, ��¶��, ��¶��ɫ, ��¶��ζ, ��������, ��������, ��������, С��, ���, �������, ǩ��
From ZLSOL.Sol_Rs_Postpartum_List A,
     Json_Table(a.Content,
                 '$' Columns(��¼ʱ�� Varchar2(20) Path '$.��¼ʱ��', ���� Number(4) Path '$.����', �鷿���� Varchar2(10) Path '$.�鷿����',
                          ��ͷ Varchar2(50) Path '$.��ͷ', �ӹ����� Number(3) Path '$.�ӹ�����', �ӹ�ѹʹ Varchar2(50) Path '$.�ӹ�ѹʹ',
                          ��¶�� Number(4) Path '$.��¶��', ��¶��ɫ Varchar2(20) Path '$.��¶��ɫ', ��¶��ζ Varchar2(20) Path '$.��¶��ζ',
                          �������� Varchar2(10) Path '$.��������', �������� Varchar2(10) Path '$.��������', �������� Varchar2(50) Path '$.��������',
                          С�� Varchar2(50) Path '$.С��', ��� Varchar2(50) Path '$.���', ������� Varchar2(100) Path '$.�������',
                          ǩ�� Varchar2(100) Path '$.ǩ��'));

create or replace force view zlsol.v_his_������������¼ as
select d.pid ����id,d.tid סԺ����,b.Babyno ���,d.name||decode(b.Sex,'��','֮��','֮Ů')||t.˳�� as Ӥ������,
b.Sex as Ӥ���Ա�,
c.�Ѵ� as �������,
a.�����ʽ,b.̥��״��,
b.��,b.����,b.Ѫ��,
b.boutt as ����ʱ��,
b.����ʱ��,'' as ��ע˵��,
b.Recorder �Ǽ���,
b.Addtime �Ǽ�ʱ��
from V_SOL_INF_DELIVERY a,v_newborn b,v_sol_rs_birth c,sol_inf_puerpera d,
(select mid,sex,babyno,row_number() over(partition by mid,sex order by babyno) ˳�� from SOL_INF_NEWBORNS t  ) t
where a.Mid=b.Mid and a.mid=d.mid and a.mid=t.mid and b.Babyno=t.babyno
and b.Mid=c.mid(+);

CREATE OR REPLACE Procedure his_�����������Ǽ�_revise
(
  mid_In   SOL_INF_NEWBORNS.mid%Type,
  bid_In   SOL_INF_NEWBORNS.bid%Type,
  state_in number    -----2���ӣ��޸� ��3ɾ��
) As
  n_����id Number(20);
  n_��ҳid   Number(20);
  n_count number(2);
  babyno_In number(2);
Begin

  select pid,tid into n_����id,n_��ҳid from sol_inf_puerpera where mid=mid_in;
  select babyno into babyno_In  from SOL_INF_NEWBORNS where bid=bid_in;
  select count(1) into n_count from ������������¼@ZLHIS_DBL where ����id=n_����id and ��ҳid = n_��ҳid and ���=babyno_In;
  --�����������޸�
  if  state_in=2 then
   if n_count=0 then  ----����
      insert into   his_������������¼
      (����id,��ҳid,���,Ӥ������,Ӥ���Ա�,�������,���䷽ʽ,̥��״��,����ʱ��,��,����,Ѫ��,��ע˵��,����ʱ��,�Ǽ�ʱ��,�Ǽ���)
      select ����id,סԺ����,���,Ӥ������,Ӥ���Ա�,�������,�����ʽ,̥��״��,����ʱ��,��,����,Ѫ��,��ע˵��,����ʱ��,�Ǽ�ʱ��,�Ǽ���
       from   ( select  d.pid ����id,d.tid סԺ����,b.Babyno ���,d.name||decode(b.Sex,'��','֮��','֮Ů')||t.˳�� as Ӥ������,
              b.Sex as Ӥ���Ա�,c.�Ѵ� as �������,a.�����ʽ,b.̥��״��,b.��,b.����,b.Ѫ��,to_date(b.boutt,'yyyy-mm-dd hh24:mi:ss') as ����ʱ��,
              b.����ʱ��,'' as ��ע˵��,b.Recorder �Ǽ���,b.Addtime �Ǽ�ʱ��
              from V_SOL_INF_DELIVERY a,v_newborn b,v_sol_rs_birth c,sol_inf_puerpera d,
              (select mid,sex,babyno,row_number() over(partition by mid,sex order by babyno) ˳�� from SOL_INF_NEWBORNS t  ) t
              where a.Mid=b.Mid and a.mid=d.mid and a.mid=t.mid and b.Babyno=t.babyno
              and b.Mid=c.mid(+) and a.mid=mid_In
                ) where ���=babyno_In;
      select count(*) into n_count from his_������������¼;
      dbms_output.put_line(n_count);
        insert into ������������¼@ZLHIS_DBL(����id,��ҳid,���,Ӥ������,Ӥ���Ա�,�������,���䷽ʽ,̥��״��,����ʱ��,��,����,Ѫ��,��ע˵��,����ʱ��,�Ǽ�ʱ��,�Ǽ���)
         select  ����id,��ҳid,���,Ӥ������,Ӥ���Ա�,�������,���䷽ʽ,̥��״��,����ʱ��,��,����,Ѫ��,��ע˵��,����ʱ��,�Ǽ�ʱ��,�Ǽ��� from his_������������¼ ;
      Zl_�����Զ����_Update@ZLHIS_DBL(n_����id, n_��ҳid);
      b_Message.Zlhis_Patient_011@ZLHIS_DBL(n_����id, n_��ҳid, babyno_In);
      delete from his_������������¼;
    else  ----�޸�
      insert into   his_������������¼
      (����id,��ҳid,���,Ӥ������,Ӥ���Ա�,�������,���䷽ʽ,̥��״��,����ʱ��,��,����,Ѫ��,��ע˵��,����ʱ��,�Ǽ�ʱ��,�Ǽ���)
      select ����id,סԺ����,���,Ӥ������,Ӥ���Ա�,�������,�����ʽ,̥��״��,����ʱ��,��,����,Ѫ��,��ע˵��,����ʱ��,�Ǽ�ʱ��,�Ǽ���
       from   ( select  d.pid ����id,d.tid סԺ����,b.Babyno ���,d.name||decode(b.Sex,'��','֮��','֮Ů')||t.˳�� as Ӥ������,
              b.Sex as Ӥ���Ա�,c.�Ѵ� as �������,a.�����ʽ,b.̥��״��,b.��,b.����,b.Ѫ��,to_date(b.boutt,'yyyy-mm-dd hh24:mi:ss') as ����ʱ��,
              b.����ʱ��,'' as ��ע˵��,b.Recorder �Ǽ���,b.Addtime �Ǽ�ʱ��
              from V_SOL_INF_DELIVERY a,v_newborn b,v_sol_rs_birth c,sol_inf_puerpera d,
              (select mid,sex,babyno,row_number() over(partition by mid,sex order by babyno) ˳�� from SOL_INF_NEWBORNS t  ) t
              where a.Mid=b.Mid and a.mid=d.mid and a.mid=t.mid and b.Babyno=t.babyno
              and b.Mid=c.mid(+) and a.mid=mid_In
                ) where ���=babyno_In;
        delete from ������������¼@ZLHIS_DBL where ����id=n_����id and ��ҳid=n_��ҳid and ���=babyno_In;
        insert into ������������¼@ZLHIS_DBL(����id,��ҳid,���,Ӥ������,Ӥ���Ա�,�������,���䷽ʽ,̥��״��,����ʱ��,��,����,Ѫ��,��ע˵��,����ʱ��,�Ǽ�ʱ��,�Ǽ���)
        select ����id,��ҳid,���,Ӥ������,Ӥ���Ա�,�������,���䷽ʽ,̥��״��,����ʱ��,��,����,Ѫ��,��ע˵��,����ʱ��,�Ǽ�ʱ��,�Ǽ��� from his_������������¼ ;
        Zl_�����Զ����_Update@ZLHIS_DBL(n_����id, n_��ҳid);
      b_Message.Zlhis_Patient_011@ZLHIS_DBL(n_����id, n_��ҳid, babyno_In);
   /* update  ������������¼@ZLHIS_DBL set
   (���,Ӥ������,Ӥ���Ա�,�������,���䷽ʽ,̥��״��,����ʱ��,��,����,Ѫ��,��ע˵��,����ʱ��,�Ǽ�ʱ��,�Ǽ���)=
   (select ���,Ӥ������,Ӥ���Ա�,�������,���䷽ʽ,̥��״��,����ʱ��,��,����,Ѫ��,��ע˵��,����ʱ��,�Ǽ�ʱ��,�Ǽ���
    from   his_������������¼) ;*/
     delete from his_������������¼;
      end if;
     --�������Ǽ�ɾ��
   elsif state_in=3 then
     delete from ������������¼@ZLHIS_DBL where ����id=n_����id and ��ҳid=n_��ҳid and ���=babyno_In;
     Zl_�����Զ����_Update@ZLHIS_DBL(n_����id,n_��ҳid);

     b_Message.ZLHIS_PATIENT_013@ZLHIS_DBL(n_����id,n_��ҳid,babyno_In);
  End If;
End his_�����������Ǽ�_revise;
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
                           Second) || '��' || '|' ||
            Extract(Hour From(To_Date(Dbegin_In, 'YYYY-MM-DD hh24:mi') - To_Date(Dend_In, 'YYYY-MM-DD hh24:mi')) Day To
                    Second) || 'ʱ' || '|' ||
            Extract(Minute From(To_Date(Dbegin_In, 'YYYY-MM-DD hh24:mi') - To_Date(Dend_In, 'YYYY-MM-DD hh24:mi')) Day To
                    Second) || '��'
    Into v_Temp
    From Dual;
    v_Temp := Replace(v_Temp, '|0��', '');
    v_Temp := Replace(v_Temp, '|0ʱ', '');
    v_Temp := Replace(v_Temp, '|0��', '');
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
  Select max(a.��¼����)
  Into v_Return
  From (Select ����id, ��ҳid, ��Ŀ����, ����ʱ��, ��¼����
         From ���˻�����ϸ@ZLHIS_DBL a, ���˻�������@ZLHIS_DBL b, ���˻����ļ�@ZLHIS_DBL c
         Where a.��¼id = b.Id And b.�ļ�id = c.Id And c.����id = n_Pid And c.��ҳid = n_Tid And c.Ӥ�� = 0 And c.�ļ����� = '�������µ�' And
               a.��Ŀ���� = Condition_In) a,
       (Select ����id, ��ҳid, ��Ŀ����, Min(b.����ʱ��) ��һ��ʱ��
         From ���˻�����ϸ@ZLHIS_DBL a, ���˻�������@ZLHIS_DBL b, ���˻����ļ�@ZLHIS_DBL c
         Where a.��¼id = b.Id And b.�ļ�id = c.Id And c.����id = n_Pid And c.��ҳid = n_Tid And c.Ӥ�� = 0 And c.�ļ����� = '�������µ�' And
               a.��Ŀ���� = Condition_In
         Group By ����id, ��ҳid, ��Ŀ����) b
  Where a.��Ŀ���� = b.��Ŀ���� And a.����ʱ�� = b.��һ��ʱ�� And a.����id = b.����id And a.��ҳid = b.��ҳid;
  Return v_Return;
Exception
  When Others Then
    Null;
End Zl_Sol_Getdata;