--��zlsol��¼ִ�����½ű���conn zlsol/his@ORA_SOL
--ZLHIS_DBL��dblink���ӣ�������ͬ����HIS��Ҫ
create database link ZLHIS_DBL connect to ZLHIS identified by &zlhis����HIS������� using '(DESCRIPTION =
    (ADDRESS = (PROTOCOL = TCP)(HOST = &HIS��IP)(PORT = &HIS��˿�))
    (CONNECT_DATA =
      (SERVICE_NAME = &HIS��ʵ����)
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
  �������� VARCHAR2(100),
  ����� VARCHAR2(100)
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

create or replace force view ZLSOL.newborn as
Select a.mid, b.SEX,b.���ھ���,b.����ʱ������ȷ���,b.�������,b.����1����,b.����5����,b.����10����,b.�۾���ҩ,b.һ�����,b.Ƥ��,b.̥֬,
b.����,b.��,b.����,b.ͷ������,b.����,b.ˮ��,b.Ѫ��,b.���,b.��,b.��ǻ,b.�ز�,b.��,b.��,b.���Ѫ,
b.��,b.Ƣ,b.����,b.��֫,b.ָ,b.ֺ,b.��ֳ��,b.����
From Sol_Inf_Newborns a,JSON_TABLE(a.newborninf,'$' columns(
SEX   Varchar2(50) PATH '$.SEX',
���ھ���            Varchar2(50) PATH '$.���ھ���',
����ʱ������ȷ���          Varchar2(50) PATH '$.����ʱ������ȷ���',
�������      Varchar2(50) PATH '$.�������',
����1���� Number(2) PATH '$.����1����',
����5���� Number(2) PATH '$.����5����',
����10���� Number(2) PATH '$.����10����',
�۾���ҩ      Varchar2(50) PATH '$.�۾���ҩ',
һ�����       Varchar2(50) PATH '$.һ�����',
Ƥ��       Varchar2(50) PATH '$.Ƥ��',
̥֬       Varchar2(10) PATH '$.̥֬',
����       Varchar2(10) PATH '$.����',
��      Varchar2(10) PATH '$.��',
����   Varchar2(20) PATH '$.����',
ͷ������        Varchar2(10) PATH '$.ͷ������',
����       Varchar2(10) PATH '$.����',
ˮ��       Varchar2(10) PATH '$.ˮ��',
Ѫ��       Varchar2(10) PATH '$.Ѫ��',
���         Varchar2(50) PATH '$.���',
��        Varchar2(10) PATH '$.��',
��ǻ       Varchar2(10) PATH '$.��ǻ',
�ز�       Varchar2(50) PATH '$.�ز�',
��       Varchar2(50) PATH '$.��',
��         Varchar2(50) PATH '$.��',
���Ѫ        Varchar2(50) PATH '$.���Ѫ',
��               Varchar2(50) PATH '$.��',
Ƣ               Varchar2(50) PATH '$.Ƣ',
����              Varchar2(50) PATH '$.����',
��֫              Varchar2(50) PATH '$.��֫',
ָ               Varchar2(50) PATH '$.ָ',
ֺ               Varchar2(50) PATH '$.ֺ',
��ֳ��             Varchar2(50) PATH '$.��ֳ��',
����              Varchar2(50) PATH '$.����'
)) as b;

create or replace force view ZLSOL.sol_userlist as
Select User_Name User_Code, Last_Name || First_Name  User_Name
  From Apex_200100.Wwv_Flow_Fnd_User
  Where First_Name Is Not Null And Last_Name Is Not Null;

create or replace force view ZLSOL.v_delivery as
Select a.Mid, b.����1,b.���̿�ʼʱ��,b.����ȫ��ʱ��,b.̥�����ʱ��,b.̥�����ʱ��,b.��һ����,b.�ڶ�����,b.��������,b.�������,b.�������������¾���,b.����,b.��Ĥ��ʽ,b.��Ĥʱ��,b.��ˮ��״,b.��ˮ��,b.��ˮ��ɫ,b.̥�������ʽ,b.̥�̰��뷽ʽ,b.̥��������,b.̥��̥Ĥ����,b.̥�����,b.̥����̬,b.̥�̴�С,b.̥������,b.�������,b.�������,b.����ƾ�,b.�����ٽ�,b.����Ѵ�,b.�����ʽ,b.���̥��λ,b.������С,b.������λ,b.�������˳̶�,b.���������п�,b.�������˷��,b.������������,b.�������˳���,b.�������˲�λ,b.��������״��,b.�������˲�λ��С,b.��������Ѫ�״�С,b.�������Ա�,b.����������,b.��������,b.��������������,b.����������������,b.������������������״,b.��������������ҩ��,b.���������Ȼ���,b.������������̥,b.��������������,b.����Ѫѹ,b.������Ѫ,b.��ʱ��ҩ,b.������ҩ,b.�������, d.����3,d.������ʱ��,d.������,d.������,d.��¼��
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
Select a.Mid, b.Bid, b.Babyno, b.Addtime, b.Recorder, a.Name, a.Old, a.Bedno, a.Pno, b.Sex, d.����2,d.BOUTT,d.��,d.����,d.Ѫ��,d.̥��״��,d.ͷΧ,d.��Χ,
d.һ�������Ӧ,d.һ�������ɫ,d.һ�����Ƥ��,d.һ������ë,d.ͷ������,d.­���ص�,d.̥ͷˮ��Ѫ��,d.̥ͷˮ�״�С,d.ǰض,d.����,d.����,d.��ǻ,
d.��,d.���,d.��,d.Ƣ,d.��֫,d.��չ����,d.����,d.��ֳ��,d.�ʺ���������,d.�ʺ���������״,d.���ܲ����������,d.���ܲ����������״,
d.����������״��,d.����ҩ��,d.����˱,d.Ƥ���Ӵ�,d.������ʽ,
e.����3,e.����1����,e.����5����,e.����10����,e.����1����,e.����5����,
e.����10����,e.����1����,e.����5����,e.����10����,e.������1����,e.������5����,e.������10����,e.��ɫ1����,e.��ɫ5����,e.��ɫ10����,
e.�ܷ�1����,e.�ܷ�5����,e.�ܷ�10����, f.����4,f.�����ڲ�ʱ�ϲ�֢����ҩ���,f.����ǰ̥�����,f.Ӥ������ʱ�������,f.����ȱ��,
f.���,f.����ʱ��
From Sol_Inf_Puerpera A, Sol_Inf_Newborns B,
     Json_Table(Nvl(b.Newborninf, '{"����2":"2"}'),
                 '$' Columns(����2 Varchar2(50) Path '$.����2',BOUTT Varchar2(50) Path '$.BOUTT', �� Varchar2(50) Path '$.��', ���� Varchar2(50) Path '$.����',
                          ͷΧ Varchar2(50) Path '$.ͷΧ', ��Χ Varchar2(50) Path '$.��Χ', һ�������Ӧ Varchar2(50) Path '$.һ�������Ӧ',
                          Ѫ�� Varchar2(50) Path '$.Ѫ��', ̥��״�� Varchar2(50) Path '$.̥��״��',
                          һ�������ɫ Varchar2(50) Path '$.һ�������ɫ', һ�����Ƥ�� Varchar2(50) Path '$.һ�����Ƥ��',
                          һ������ë Varchar2(50) Path '$.һ������ë', ͷ������ Varchar2(50) Path '$.ͷ������',
                          ­���ص� Varchar2(50) Path '$.­���ص�', ̥ͷˮ��Ѫ�� Varchar2(50) Path '$.̥ͷˮ��Ѫ��',
                          ̥ͷˮ�״�С Varchar2(50) Path '$.̥ͷˮ�״�С', ǰض Varchar2(50) Path '$.ǰض', ���� Varchar2(50) Path '$.����',
                          ���� Varchar2(50) Path '$.����', ��ǻ Varchar2(50) Path '$.��ǻ', �� Varchar2(50) Path '$.��',
                          ��� Varchar2(50) Path '$.���', �� Varchar2(50) Path '$.��', Ƣ Varchar2(50) Path '$.Ƣ',
                          ��֫ Varchar2(50) Path '$.��֫', ��չ���� Varchar2(50) Path '$.��չ����', ���� Varchar2(50) Path '$.����',
                          ��ֳ�� Varchar2(50) Path '$.��ֳ��',�ʺ��������� Varchar2(50) Path '$.�ʺ���������',�ʺ���������״ Varchar2(50) Path '$.�ʺ���������״',
                          ���ܲ���������� Varchar2(50) Path '$.���ܲ����������',���ܲ����������״ Varchar2(50) Path '$.���ܲ����������״',
                          ������ʽ Varchar2(50) Path '$.������ʽ',����������״�� Varchar2(50) Path '$.����������״��',����ҩ�� Varchar2(50) Path '$.����ҩ��',
                          ����˱ Varchar2(50) Path '$.����˱',Ƥ���Ӵ� Varchar2(50) Path '$.Ƥ���Ӵ�')) As D,
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
                 '$' Columns(����4 Varchar2(50) Path '$.����4', �����ڲ�ʱ�ϲ�֢����ҩ��� Varchar2(50) Path '$.�����ڲ�ʱ�ϲ�֢����ҩ���',
                          ����ǰ̥����� Varchar2(50) Path '$.����ǰ̥�����', Ӥ������ʱ������� Varchar2(50) Path '$.Ӥ������ʱ�������',
                          ����ȱ�� Varchar2(50) Path '$.����ȱ��', ����ʱ�� Varchar2(50) Path '$.����ʱ��',
                          ��� Varchar2(50) Path '$.���')) As F
Where a.Mid = b.Mid
Order By Addtime;

create or replace force view ZLSOL.v_sol_inf_checkinroom as
Select a.mid, b.�뷿Ŀ��,b.�뷿ʱ��,b.ҽ�Ʋ���,b.������,b.����֪��֪ͨ��,b.����������,b.̥����,b.̥�Ĵ���,b.��Ĥ���,b.�Ƿ��кϲ�֢,b.����,b.��Һ��,b.����ͨ��,b.�ֲ����,b.����ҩ��,b.����,b.������,b.�Ӱ���
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
Select a.mid, b.OUTROOMTIME,b.����״̬,b.ҽ�Ʋ���,b.������,b.����ͨ��,b.�ֲ����,b.��������,b.�����п���,b.�����пڷ��,b.����ˮ��,b.����Ѫ��,b.�����Ѫ,b.��Ѫ��,b.����ҩ��,b.������,b.�Ӱ���,b.ҩ��,b.��ע
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
Select a.Mid, b.����1, b.BEGINT, b.ALLT, b.OUTT, b.ALLOUTT, b.��һ����, b.�ڶ�����, b.��������, b.�������,
       b.����, b.��Ĥ��ʽ, b.��Ĥʱ��, b.��ˮ��״, b.��ˮ��, b.��ˮ��ɫ, b.̥Ĥ����ʽ, b.̥�̰��뷽ʽ, b.̥��������,
       b.̥��̥Ĥ����, b.̥�����, b.̥����̬, b.̥�̴�С, b.̥������, b.�������, b.�������, b.�����ٽ�, b.���,b.�ƾ�����, b.�����ʽ,
       b.���̥��λ, b.������С, b.������λ, b.�������˳̶�, b.���������п�, b.�������˷��, b.������������, b.�������˳���, b.�������˲�λ, b.��������״��,
       b.�������˲�λ��С, b.��������Ѫ�״�С, b.���󼴿�����ѹ,b.���󼴿�����ѹ,b.����1Сʱ����ѹ,b.����1Сʱ����ѹ,b.����2Сʱ����ѹ,b.����2Сʱ����ѹ,
       b.DNOW,b.DONE,b.DTWO,b.�����Ѫ����,b.���󼴿�����,b.����1Сʱ����,b.����2Сʱ����, b.��ʱ��ҩ, b.������ҩ,b.��ʹ������ҩ,
       b.�������, b.�������, d.����3,d.������ʱ��,d.��������������, d.������, d.������, d.��¼��
From Sol_Inf_Delivery A,
     Json_Table(Nvl(a.Newborndetail, '{"����1":"1"}'),
                 '$'Columns(����1 Varchar2(50) Path '$.����1', BEGINT Varchar2(19) Path '$.BEGINT',
                          ALLT Varchar2(19) Path '$.ALLT', OUTT Varchar2(19) Path '$.OUTT',
                          ALLOUTT Varchar2(19) Path '$.ALLOUTT', ��һ���� Varchar2(50) Path '$.��һ����',
                          �ڶ����� Varchar2(50) Path '$.�ڶ�����', �������� Varchar2(50) Path '$.��������',���� Varchar2(50) Path '$.����',
                          ��Ĥ��ʽ Varchar2(50) Path '$.��Ĥ��ʽ', ��Ĥʱ�� Varchar2(19) Path '$.��Ĥʱ��', ��ˮ��״ Varchar2(50) Path '$.��ˮ��״',
                          ��ˮ�� Varchar2(50) Path '$.��ˮ��', ��ˮ��ɫ Varchar2(50) Path '$.��ˮ��ɫ',
                          ̥Ĥ����ʽ Varchar2(50) Path '$.̥Ĥ����ʽ', ̥�̰��뷽ʽ Varchar2(50) Path '$.̥�̰��뷽ʽ',
                          ̥�������� Varchar2(50) Path '$.̥��������', ̥��̥Ĥ���� Varchar2(50) Path '$.̥��̥Ĥ����',
                          ̥����� Varchar2(50) Path '$.̥�����', ̥����̬ Varchar2(50) Path '$.̥����̬', ̥�̴�С Varchar2(50) Path '$.̥�̴�С',
                          ̥������ Varchar2(50) Path '$.̥������', ������� Varchar2(50) Path '$.�������', ������� Varchar2(50) Path '$.�������',
                          �����ٽ� Varchar2(50) Path '$.�����ٽ�',��� Varchar2(50) Path '$.���', �ƾ����� Varchar2(50) Path '$.�ƾ�����',
                          �����ʽ Varchar2(50) Path '$.�����ʽ',���̥��λ Varchar2(50) Path '$.���̥��λ', ������С Varchar2(50) Path '$.������С',
                          ������λ Varchar2(50) Path '$.������λ', �������˳̶� Varchar2(50) Path '$.�������˳̶�',
                          ���������п� Varchar2(50) Path '$.���������п�', �������˷�� Varchar2(50) Path '$.�������˷��',
                          ������������ Varchar2(50) Path '$.������������', �������˳��� Varchar2(50) Path '$.�������˳���',
                          �������˲�λ Varchar2(50) Path '$.�������˲�λ', ��������״�� Varchar2(50) Path '$.��������״��',
                          �������˲�λ��С Varchar2(50) Path '$.�������˲�λ��С', ��������Ѫ�״�С Varchar2(50) Path '$.��������Ѫ�״�С',
                          ������� Varchar2(50) Path '$.�������',���󼴿�����ѹ Varchar2(50) Path '$.���󼴿�����ѹ',���󼴿�����ѹ Varchar2(50) Path '$.���󼴿�����ѹ',
                          ����1Сʱ����ѹ Varchar2(50) Path '$.����1Сʱ����ѹ',����1Сʱ����ѹ Varchar2(50) Path '$.����1Сʱ����ѹ',
                          ����2Сʱ����ѹ Varchar2(50) Path '$.����2Сʱ����ѹ',����2Сʱ����ѹ Varchar2(50) Path '$.����2Сʱ����ѹ',
                          ���󼴿����� Varchar2(50) Path '$.���󼴿�����',����1Сʱ���� Varchar2(50) Path '$.����1Сʱ����',����2Сʱ���� Varchar2(50) Path '$.����2Сʱ����',
                          DNOW Varchar2(50) Path '$.DNOW',DONE Varchar2(50) Path '$.DONE',DTWO Varchar2(50) Path '$.DTWO',�����Ѫ���� Varchar2(50) Path '$.�����Ѫ����',
                          ��ʱ��ҩ Varchar2(50) Path '$.��ʱ��ҩ', ������ҩ Varchar2(50) Path '$.������ҩ',��ʹ������ҩ Varchar2(50) Path '$.��ʹ������ҩ',
                          ������� Varchar2(50) Path '$.�������',������� Varchar2(50) Path '$.�������')) As B,
     Json_Table(Nvl(a.Deliveryinf, '{"����3":"1"}'),
                 '$' Columns(����3 Varchar2(1) Path '$.����3', ������ʱ�� Varchar2(50) Path '$.������ʱ��',�������������� Varchar2(50) Path '$.��������������',
                          ������ Varchar2(50) Path '$.������', ������ Varchar2(50) Path '$.������', ��¼�� Varchar2(50) Path '$.��¼��')) As D;

create or replace force view ZLSOL.v_sol_inf_equipment as
Select a.mid, b.���м���ǰ,b.���м�����,b.���м�����,b.�������ǰ,b.���������,b.���������,b.ֹѪǯ��ǰ,b.ֹѪǯ����,b.ֹѪǯ����,b.������ǰ,b.��������,b.��������,b.��������ǰ,b.����������,b.����������,b.�������ǰ,b.����������,b.���������,b.ϴ�����ǰ,b.ϴ��������,b.ϴ�������,b.������ǰ,b.���������,b.��������,b.������ǰ,b.��������,b.��������,b.����ǯ��ǰ,b.����ǯ����,b.����ǯ����,b.������ǰ,b.��������,b.��������,b.�γײ�ǰ,b.�γ�����,b.�γײ���,b.����˹��ǰ,b.����˹����,b.����˹����,b.��ǯ��ǰ,b.��ǯ����,b.��ǯ����,b.ɴ����ǰ,b.ɴ������,b.ɴ������,b.��Բǯ��ǰ,b.��Բǯ����,b.��Բǯ����
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
��ǯ��ǰ   Number(2) PATH '$.��ǯ��ǰ',
��ǯ����   Number(2) PATH '$.��ǯ����',
��ǯ����   Number(2) PATH '$.��ǯ����',
ɴ����ǰ   Number(2) PATH '$.ɴ����ǰ',
ɴ������   Number(2) PATH '$.ɴ������',
ɴ������   Number(2) PATH '$.ɴ������',
��Բǯ��ǰ   Number(2) PATH '$.��Բǯ��ǰ',
��Բǯ����   Number(2) PATH '$.��Բǯ����',
��Բǯ����   Number(2) PATH '$.��Բǯ����'
)) as b;

create or replace force view ZLSOL.v_sol_inf_newborns as
Select a.Mid, b.Bid, b.Babyno, b.Addtime, b.Recorder, a.Name, a.Old, a.Bedno, a.Pno, b.Sex, d.����2,d.��,d.����,d.ͷΧ,d.��Χ,d.һ�������Ӧ,d.һ�������ɫ,d.һ�����Ƥ��,d.һ������ë,d.ͷ������,d.­���ص�,d.̥ͷˮ��Ѫ��,d.̥ͷˮ�״�С,d.ǰض,d.����,d.����,d.��ǻ,d.��,d.���,d.��,d.Ƣ,d.��֫,d.��չ����,d.����,d.��ֳ��, e.����3,e.����1����,e.����5����,e.����10����,e.����1����,e.����5����,e.����10����,e.����1����,e.����5����,e.����10����,e.������1����,e.������5����,e.������10����,e.��ɫ1����,e.��ɫ5����,e.��ɫ10����,e.�ܷ�1����,e.�ܷ�5����,e.�ܷ�10����, f.����4,f.�����ڲ�ʱ�ϲ�֢����ҩ���,f.����ǰ̥�����,f.Ӥ������ʱ�������,f.����ȱ��,f.ĸ��ι��ָ��,f.���
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
       Decode(Checkoutroom, 1, '��', '') ����,Decode(Equipment, 1, '��', '') ��е,
       Outtime,pid,tid
From Sol_Inf_Puerpera;

create or replace force view ZLSOL.v_sol_rs_birth as
Select a.mid,b.�Ѵ�,b.����,b.Ѫ��,b.��������ʷ,b.ĩ���¾�,b.Ԥ����,b.��ǰ�ϼ��侶,b.���ռ侶,b.���ǽ�ڼ侶,b.�����⾶,b.���ǻ���,b.���ǹؽ�,b.�����м�,b.������,b.����֢,b.��ǰ��¼����,b.���ʱ��,b.����ѹ,b.����ѹ,b.����,b.����,b.̥����,b.̥����С,b.����������,b.̥λ,b.��Ĥ���,b.��¶,b.����,b.�����,b.������ʼʱ��,b.��Ĥʱ��,b.��Ժ����
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
Select  a.courseid,a.mid,b.���ʱ��,b.�Ƿ��ʹ���,b.̥��λ,b.����ѹ,b.����ѹ,b.����,b.����,b.̥����,b.����ǿ��,b.��������,b.�������,b.������,b.����,b.��Ĥ���,b.��¶,b.����,b.�����
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
����        Varchar2(200) PATH '$.����'��
�����      Varchar2(50) PATH '$.�����'
)) as b;

create or replace force view ZLSOL.v_sol_rs_druglabor as
Select Mid, To_Char(����, 'YYYY-MM-DD HH24:MI') ����, ����ָ��, ��������,����� from Sol_Rs_Druglabor;

create or replace force view ZLSOL.v_sol_rs_druglabor_list as
Select a.Mid, a.Courseid ID, b.��¼ʱ��,b.����ѹ,b.����ѹ,b.����,b.̥����,b.����ǿ��,b.��������,b.�������,b.����,b.��¶,b.��ˮ��,b.��ˮ��״,b.����,b.��¼��,b.����,b.����
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
Select a.mid,a.courseid,b.��¼ʱ��,b.̥��λ,b.����ѹ,b.����ѹ,b.����,b.��Χ,b.̥��������,b.̥��������,b.̥��������,b.̥����,b.��¶,b.����,b.��Ĥ���,b.��ˮ��״,b.����ǿ��,b.��������,b.�������,b.����,b.�����
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
̥���� varchar2(50) PATH '$.̥����',
��¶     Varchar2(20) PATH '$.��¶',
����     Varchar2(20) PATH '$.����',
��Ĥ���     Varchar2(20) PATH '$.��Ĥ���',
��ˮ��״      Varchar2(20) PATH '$.��ˮ��״',
����ǿ��     Varchar2(20) PATH '$.����ǿ��',
��������       Varchar2(20) PATH '$.��������',
�������       Varchar2(20) PATH '$.�������',
����     Varchar2(500) PATH '$.����',
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

CREATE OR REPLACE Function ZLSOL.Sol_Getsdate
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
End Sol_Getsdate;
/
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

CREATE OR REPLACE Procedure ZLSOL.his_�����������Ǽ�_revise
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
CREATE OR REPLACE Function ZLSOL.f_List2str
( 
  p_Strlist   In t_Strlist, 
  p_Delimiter In Varchar2 Default ',', 
  p_Distinct  In Number Default 1, 
  p_Maxlength In Number Default 4000 
) Return Varchar2 Is 
  l_String Long; 
  l_Add    Number; 
  --���ܣ���һ���б���ת��Ϊһ��ȱʡ�Զ��ŷָ����ַ����� 
  --���� 
  --Select ����, f_List2str(Cast(Collect(��Ա Order By ���) As t_Strlist)) ��Ա�б� 
  --From (Select a.���� As ����, c.���� As ��Ա,c.��� 
  --      From ���ű� A, ������Ա B, ��Ա�� C 
  --      Where a.Id = b.����id And b.��Աid = c.Id 
  --      Order By ����, ��Ա) 
  --Group By ���� 
 
  --�˺�����֧��with��ʽ�������ʱ�ڴ���⽫�ᱨ��ORA-00932: �������Ͳ�һ��: ӦΪ -, ��ȴ��� -�� 
  --���磺With Test As (Select '�ڿ�' As ����,'����' As ��Ա From Dual Union All......) 
  --     Select ����,f_List2str(cast(COLLECT(��Ա) as t_Strlist)) tt From Test Group By ���� 
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
--�û�ͬ������ű�
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
  ----������Ա���޸���Ա��ɾ����Ա��������Ա��ͣ����Ա
Begin
  n_Group_Id := Apex_Util.Find_Security_Group_Id(p_Workspace => 'ZLSOL');
  If Inserting Then
    --������Ա
    Apex_Util.Set_Security_Group_Id(p_Security_Group_Id => n_Group_Id);
    Apex_Util.Create_User(p_User_Name => :New.Code, p_First_Name => Substr(:New.Name, 2),
                          p_Last_Name => Substr(:New.Name, 1, 1), p_Web_Password => '123',
                          p_Change_Password_On_First_Use => 'N');
  Elsif Deleting Then
    --ɾ����Ա
    Apex_Util.Set_Security_Group_Id(p_Security_Group_Id => n_Group_Id);
    Apex_Util.Remove_User(p_User_Name => :Old.Code);
  Elsif Updating Then
    --�޸���Ա����
    If :New.Name <> :Old.Name Then
      Apex_Util.Set_Security_Group_Id(p_Security_Group_Id => n_Group_Id);
      n_User_Id := Apex_Util.Get_User_Id(p_Username => :Old.Code);
      Apex_Util.Set_First_Name(p_Userid => n_User_Id, p_First_Name => Substr(:New.Name, 2));
      Apex_Util.Set_Last_Name(p_Userid => n_User_Id, p_Last_Name => Substr(:New.Name, 1, 1));
    End If;
    --���á�ͣ����Ա��Ӧ���ú������˻�
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
--��������
create table ZLSOL.SOL_STD_FetalPosition--̥��λ
(
code varchar2(10),
name varchar2(50),
Description varchar2(100)
);
create table ZLSOL.SOL_STD_Delivery--���䷽ʽ
(
code varchar2(10),
name varchar2(50),
Description varchar2(100)
);
create table ZLSOL.SOL_STD_PerinealLaceration--�����������
(
code varchar2(10),
name varchar2(50),
Description varchar2(100)
);
create table ZLSOL.SOL_STD_Anesthesia--����ʽ
(
code varchar2(10),
name varchar2(50),
Description varchar2(500)
);
create table ZLSOL.SOL_STD_FetalPresentation--̥��¶
(
code varchar2(10),
name varchar2(50),
Description varchar2(100)
);
create table ZLSOL.SOL_STD_NeonatalAbnormality--�������쳣���
(
code varchar2(10),
name varchar2(50),
Description varchar2(100)
);

--̥��λ
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('01','����ǰ(LOA)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('02','����ǰ(ROA)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('03','�����(LOP)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('04','�����(ROP)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('05','�����(LOT)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('06','�����(ROT)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('07','���ǰ(LMA)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('08','���ǰ(RMA)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('09','����(LMP)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('10','����(RMP)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('11','����(LMT)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('12','����(RMT)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('13','����ǰ(LSA)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('14','����ǰ(RSA)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('15','������(LSP)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('16','������(RSP)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('17','������(LST)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('18','������(RST)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('19','���ǰ(LScA)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('20','�Ҽ�ǰ(RscA)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('21','����(LScP)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('22','�Ҽ��(RScP)','');
Insert Into ZLSOL.SOL_STD_FetalPosition(code,name,Description) Values('99','����','');
--���䷽ʽ
Insert Into ZLSOL.SOL_STD_Delivery(code,name,Description) Values('1','������Ȼ����','');
Insert Into ZLSOL.SOL_STD_Delivery(code,name,Description) Values('11','�����п�','');
Insert Into ZLSOL.SOL_STD_Delivery(code,name,Description) Values('12','����δ��','');
Insert Into ZLSOL.SOL_STD_Delivery(code,name,Description) Values('2','������������','');
Insert Into ZLSOL.SOL_STD_Delivery(code,name,Description) Values('21','��ǯ����','');
Insert Into ZLSOL.SOL_STD_Delivery(code,name,Description) Values('22','��λ����','');
Insert Into ZLSOL.SOL_STD_Delivery(code,name,Description) Values('23','̥ͷ����','');
Insert Into ZLSOL.SOL_STD_Delivery(code,name,Description) Values('3','�ʹ���','');
Insert Into ZLSOL.SOL_STD_Delivery(code,name,Description) Values('31','�ӹ��¶κ��п��ʹ���','');
Insert Into ZLSOL.SOL_STD_Delivery(code,name,Description) Values('32','�ӹ����ʹ���','');
Insert Into ZLSOL.SOL_STD_Delivery(code,name,Description) Values('33','��Ĥ���ʹ���','');
Insert Into ZLSOL.SOL_STD_Delivery(code,name,Description) Values('9','����','');
--�����������
Insert Into ZLSOL.SOL_STD_PerinealLaceration(code,name,Description) Values('1','������','');
Insert Into ZLSOL.SOL_STD_PerinealLaceration(code,name,Description) Values('2','�������','');
Insert Into ZLSOL.SOL_STD_PerinealLaceration(code,name,Description) Values('3','�������','');
Insert Into ZLSOL.SOL_STD_PerinealLaceration(code,name,Description) Values('4','�������','');
Insert Into ZLSOL.SOL_STD_PerinealLaceration(code,name,Description) Values('5','�����п�','');
--����ʽ
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('1','ȫ������','�������ʹȫ��������״̬');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('11','��������','������������ķ���ʹȫ��������״̬');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('12','��������','������ע�������ʹȫ��������״̬');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('13','��������','����ǰ��ʹ������־��ʧ�ķ���');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('2','׵��������','������ҩע��׵���ڴﵽ�ֲ�����Ч���ķ���');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('21','����Ĥ��ǻ��������','������ҩע������Ĥ��ǻ�ﵽ�ֲ�����Ч���ķ���');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('22','Ӳ��Ĥ��ǻ��������','������ҩע��Ӳ��Ĥ��ǻ�����ֲ�����Ч���ķ���');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('3','�ֲ�����','������ҩֱ��ע��ʩ����������֯�ڻ�������λ��Χ��������');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('31','�񾭴���������','���ֲ�����ҩע�����񾭴Ը�����ʹͨ���񾭴Ե��񾭼������ֲ�����������ֲ�����ķ���');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('32','�񾭽���������','���ֲ�����ҩע�����񾭽ڸ�����ʹͨ���񾭽ڵ��񾭼������ֲ�����������ֲ�����ķ���');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('33','����������','������ҩ��ע�����񾭸ɵ���Χ��ʹ���񾭷ֲ�����������������õķ���');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('34','������������','������ҩע��������Ұ���ܣ�ʹͨ������Ұ�Լ�������Ұ��������ĩ�ҽ��ܵ����͵ľֲ�������');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('35','�ֲ���������','������ҩ�������п��߷ֲ�ע����֯�ڣ���������֯�е���ĩ�ҵ�������');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('36','��������','������ҩֱ����ճĤ��Ƥ���Ӵ���ʹ֧��ò���ճĤ��Ƥ���ڵ���ĩ�ұ����͵�������');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('4','��������','��һ������ҩ�����ö�������������ǿ����Ч��');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('41','��������ȫ��','�����������������ͬ���ò�������Ч��');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('42','��ҩ��������','��������ҩ������ͬ���ò�������Ч��');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('43','�񾭴���ӲĤ�����͸�������','�񾭴����������Ӳ��Ĥ��ǻ��������ͬ���ò�������Ч��');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('44','ȫ�鸴��ȫ����','��ȫ�������ͬʱ�������ͻ���Ѫѹ');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('45','ȫ�鸴�Ͽ����Խ�ѹ','��ȫ�������ͬʱ���ͻ��ߵ�����');
Insert Into ZLSOL.SOL_STD_Anesthesia(code,name,Description) Values('9','����������','����δ�ἰ������������');
--̥��¶
Insert Into ZLSOL.SOL_STD_FetalPresentation(code,name,Description) Values('1','ͷ��¶','');
Insert Into ZLSOL.SOL_STD_FetalPresentation(code,name,Description) Values('2','����¶','');
Insert Into ZLSOL.SOL_STD_FetalPresentation(code,name,Description) Values('3','����¶','');
Insert Into ZLSOL.SOL_STD_FetalPresentation(code,name,Description) Values('4','����¶','');
Insert Into ZLSOL.SOL_STD_FetalPresentation(code,name,Description) Values('9','����','');
--�������쳣���
Insert Into ZLSOL.SOL_STD_NeonatalAbnormality(code,name,Description) Values('1','��','');
Insert Into ZLSOL.SOL_STD_NeonatalAbnormality(code,name,Description) Values('2','��������������','');
Insert Into ZLSOL.SOL_STD_NeonatalAbnormality(code,name,Description) Values('3','����','');
Insert Into ZLSOL.SOL_STD_NeonatalAbnormality(code,name,Description) Values('4','���','');
Insert Into ZLSOL.SOL_STD_NeonatalAbnormality(code,name,Description) Values('5','��Ϣ','');
Insert Into ZLSOL.SOL_STD_NeonatalAbnormality(code,name,Description) Values('6','�ͳ�������','');
Insert Into ZLSOL.SOL_STD_NeonatalAbnormality(code,name,Description) Values('9','����','');