--��APEX��ִ�У��޸�zlsol�����룬ip��ʵ����[SERVICE_NAME]��
create database link To_His  connect to ZLHIS identified by his  using '(DESCRIPTION =
    (ADDRESS = (PROTOCOL = TCP)(HOST = 192.168.0.60)(PORT = 1521))
    (CONNECT_DATA =
      (SERVICE_NAME = orcl)
    )
  )';
create table ZLSOL_JC.HIS_������������¼
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
  is '������Ϣ';
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
  is '�뷿��Ϣ';
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
  is '������Ϣ';
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
  is '������Ϣ';
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
  is '��е����¼';
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
  is '��������Ϣ';
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
  is '��ǰ�����Ϣ';
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
  is '���̾���';
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
  ����   DATE,
  ����ָ�� VARCHAR2(50),
  �������� VARCHAR2(50)
)
;
comment on table ZLSOL_JC.SOL_RS_DRUGLABOR
  is 'ҩ��������Ϣ';
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
  is 'ҩ��������¼';
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
  is '������¼';
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
  is '����۲���Ϣ';
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
  is '����۲��¼';
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
  is 'ϵͳ�û���Ϣ';
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
Select a.Mid, b.Bid, b.Babyno, b.Addtime, b.Recorder, a.Name, a.Old, a.Bedno, a.Pno, b.Sex, d."����2",d."��",d."����",d."ͷΧ",d."��Χ",d."Ѫ��",d."̥��״��",d."��Ժ���",d."����ʱ��",
d."һ�������Ӧ",d."һ�������ɫ",d."һ�����Ƥ��",d."һ������ë",d."ͷ������",d."­���ص�",d."������С",d."̥ͷˮ��Ѫ��",d."̥ͷˮ�״�С",d."ǰض",d."����",

d."���󼴿�",d."�����Сʱ",d."����һСʱ",d."�����Сʱ",d."������Сʱ",d."������Сʱ",

 d."����",d."��ˮ��",d."�������������",d."��������",d."�������������1",d."��������1",d."����",
        d."��������30�������", d."��������30������",d."��������30����ɫ",
        d."��ѹͨ��30�������",d."��ѹͨ��30������",d."��ѹͨ��30����ɫ",
        d."������ѹͨ�����������",d."������ѹͨ����������",d."������ѹͨ��������ɫ",
        d."��ѹͨ�������ⰴѹ30�������", d."��ѹͨ�������ⰴѹ30������", d."��ѹͨ�������ⰴѹ30����ɫ",
        d."ʹ���������غ�����",d."ʵʩ������Ҫ��ʩ�������",

       f.�������ղ�����ʮ��,f.�������ղ�����ʮ��,f.�������ղ����ʮ��,f.�������ղ���������,f.�������ղ���������,f.�������ղ��������,f.�������ղ���ʮ����,f.�������ղ����ʮ����,
       f.��ѹ������ʮ��,f.��ѹ������ʮ��,f.��ѹ������ʮ��,f.��ѹ����������,f.��ѹ����������,f.��ѹ���������,f.��ѹ����ʮ����,f.��ѹ������ʮ����,
       f.���ܲ������̥����ʮ��,f.���ܲ������̥����ʮ��,f.���ܲ������̥���ʮ��,f.���ܲ������̥��������,f.���ܲ������̥��������,f.���ܲ������̥�������,f.���ܲ������̥��ʮ����,f.���ܲ������̥���ʮ����,
       f.��ѹͨ����ʮ��,f.��ѹͨ����ʮ��,f.��ѹͨ����ʮ��,f.��ѹͨ��������,f.��ѹͨ��������,f.��ѹͨ�������,f.��ѹͨ��ʮ����,f.��ѹͨ����ʮ����,
       f.���ܲ����ʮ��,f.���ܲ����ʮ��,f.���ܲ�ܾ�ʮ��,f.���ܲ��������,f.���ܲ��������,f.���ܲ�������,f.���ܲ��ʮ����,f.���ܲ�ܶ�ʮ����,
       f.���ⰴѹ��ʮ��,f.���ⰴѹ��ʮ��,f.���ⰴѹ��ʮ��,f.���ⰴѹ������,f.���ⰴѹ������,f.���ⰴѹ�����,f.���ⰴѹʮ����,f.���ⰴѹ��ʮ����,
       f.����������ʮ��,f.����������ʮ��,f.�������ؾ�ʮ��,f.��������������,f.��������������,f.�������������,f.��������ʮ����,f.�������ض�ʮ����,
       f.������ˮ��ʮ��,f.������ˮ��ʮ��,f.������ˮ��ʮ��,f.������ˮ������,f.������ˮ������,f.������ˮ�����,f.������ˮʮ����,f.������ˮ��ʮ����,
        d."����ʱ��",d."���տ�ʼʱ��",d."���ս���ʱ��",d."����ǰʱ��",d."�����ʱ��",d."��Ҫ������Ա",d."���Ƚ��",
e."����3",e."����1����",e."����5����",e."����10����",
e."����1����",e."����5����",e."����10����",e."����1����",e."����5����",e."����10����",e."������1����",e."������5����",e."������10����",e."��ɫ1����",
e."��ɫ5����",e."��ɫ10����",e."�ܷ�1����",e."�ܷ�5����",e."�ܷ�10����", f."����4",f."�����ڲ�ʱ�ϲ�֢����ҩ���",f."����ǰ̥�����",f."Ӥ������ʱ�������",
f."����ȱ��",f."ĸ��ι��ָ��",f."���",f."��ע",f."������ǩ��",f."�ʹ������߼���ȷ��������ǩ��",f."����������ϵ",

f."��Ժʱ����",f."��Ժʱ���",f."��Ժʱ�β�",f."��ԺʱƤ��",f."����",f."�������������",f."������δ����ԭ��",f."�Ҹ������������",f."�Ҹ�����δ����ԭ��",
f."����ԭ��",f."����ҽʦǩ��"
From Sol_Inf_Puerpera A, Sol_Inf_Newborns B,
     Json_Table(Nvl(b.Newborninf, '{"����2":"2"}'),
                 '$' Columns(����2 Varchar2(50) Path '$.����2', �� Varchar2(50) Path '$.��', ���� Varchar2(50) Path '$.����',
                          ͷΧ Varchar2(50) Path '$.ͷΧ', ��Χ Varchar2(50) Path '$.��Χ',Ѫ�� Varchar2(50) Path '$.Ѫ��',̥��״�� Varchar2(50) Path '$.̥��״��',
                          ��Ժ��� Varchar2(10) PATH '$.��Ժ���',����ʱ�� Varchar2(19) Path '$.����ʱ�� ',һ�������Ӧ Varchar2(50) Path '$.һ�������Ӧ',
                          һ�������ɫ Varchar2(50) Path '$.һ�������ɫ', һ�����Ƥ�� Varchar2(50) Path '$.һ�����Ƥ��',
                          һ������ë Varchar2(50) Path '$.һ������ë', ͷ������ Varchar2(50) Path '$.ͷ������',
                          ­���ص� Varchar2(50) Path '$.­���ص�', ������С Varchar2(50) Path '$.������С',̥ͷˮ��Ѫ�� Varchar2(50) Path '$.̥ͷˮ��Ѫ��',
                          ̥ͷˮ�״�С Varchar2(50) Path '$.̥ͷˮ�״�С', ǰض Varchar2(50) Path '$.ǰض', ���� Varchar2(50) Path '$.����',
                          ���󼴿� Varchar2(50) Path '$.���󼴿�',
                          �����Сʱ Varchar2(50) Path '$.�����Сʱ',����һСʱ Varchar2(50) Path '$.����һСʱ',�����Сʱ Varchar2(50) Path '$.�����Сʱ',
                          ������Сʱ Varchar2(50) Path '$.������Сʱ',������Сʱ Varchar2(50) Path '$.������Сʱ',

                          ���� Varchar2(50) Path '$.����', ��ˮ�� Varchar2(50) Path '$.��ˮ��',������������� Varchar2(50) Path '$.�������������', �������� Varchar2(50) Path '$.��������',
                          �������������1 Varchar2(50) Path '$.�������������1', ��������1 Varchar2(50) Path '$.��������1', ���� Varchar2(50) Path '$.����',

                          ��������30������� Varchar2(50) Path '$.��������30�������',��������30������ Varchar2(50) Path '$.��������30������',��������30����ɫ Varchar2(50) Path '$.��������30����ɫ',
                          ��ѹͨ��30������� Varchar2(50) Path '$.��ѹͨ��30�������',��ѹͨ��30������ Varchar2(50) Path '$.��ѹͨ��30������',��ѹͨ��30����ɫ Varchar2(50) Path '$.��ѹͨ��30����ɫ',
                          ������ѹͨ����������� Varchar2(50) Path '$.������ѹͨ�����������',������ѹͨ���������� Varchar2(50) Path '$.������ѹͨ����������',������ѹͨ��������ɫ Varchar2(50) Path '$.������ѹͨ��������ɫ',
                          ��ѹͨ�������ⰴѹ30������� Varchar2(50) Path '$.��ѹͨ�������ⰴѹ30�������',��ѹͨ�������ⰴѹ30������ Varchar2(50) Path '$.��ѹͨ�������ⰴѹ30������',��ѹͨ�������ⰴѹ30����ɫ Varchar2(50) Path '$.��ѹͨ�������ⰴѹ30����ɫ',
                          ʹ���������غ����� Varchar2(50) Path '$.ʹ���������غ�����',ʵʩ������Ҫ��ʩ������� Varchar2(50) Path '$.ʵʩ������Ҫ��ʩ�������',

                          ����ʱ�� Varchar2(19) Path '$.����ʱ��',���տ�ʼʱ�� Varchar2(19) Path '$.���տ�ʼʱ��',���ս���ʱ�� Varchar2(19) Path '$.���ս���ʱ��',
                          ����ǰʱ�� Varchar2(19) Path '$.����ǰʱ��', �����ʱ�� Varchar2(19) Path '$.�����ʱ��',
                          ��Ҫ������Ա Varchar2(50) Path '$.��Ҫ������Ա',���Ƚ�� Varchar2(50) Path '$.���Ƚ��' )) as D,
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
                          ��� Varchar2(50) Path '$.���  ',  ��ע Varchar2(50) Path '$.��ע  ',������ǩ�� Varchar2(50) Path '$.������ǩ�� ',
                          �ʹ������߼���ȷ��������ǩ�� Varchar2(50) Path '$.�ʹ������߼���ȷ��������ǩ��', ����������ϵ Varchar2(50) Path '$.����������ϵ',

                          ��Ժʱ���� Varchar2(50) Path '$.��Ժʱ����  ',��Ժʱ��� Varchar2(50) Path '$.��Ժʱ���  ',��Ժʱ�β� Varchar2(50) Path '$.��Ժʱ�β�  ',
                          ��ԺʱƤ�� Varchar2(50) Path '$.��ԺʱƤ��  ',���� Varchar2(50) Path '$.����  ',
                          ������������� Varchar2(50) Path '$.�������������  ',������δ����ԭ�� Varchar2(50) Path '$.������δ����ԭ��  ',
                          �Ҹ������������ Varchar2(50) Path '$.�Ҹ������������  ',�Ҹ�����δ����ԭ�� Varchar2(50) Path '$.�Ҹ�����δ����ԭ��  ',
                          ����ԭ�� Varchar2(50) Path '$.����ԭ��  ',����ҽʦǩ�� Varchar2(50) Path '$.����ҽʦǩ��  ',

                          �������ղ�����ʮ�� Varchar2(50) Path '$.�������ղ�����ʮ��',�������ղ�����ʮ�� Varchar2(50) Path '$.�������ղ�����ʮ��',�������ղ����ʮ�� Varchar2(50) Path '$.�������ղ����ʮ��',
                          �������ղ��������� Varchar2(50) Path '$.�������ղ���������',�������ղ��������� Varchar2(50) Path '$.�������ղ���������',�������ղ�������� Varchar2(50) Path '$.�������ղ��������',
                          �������ղ���ʮ���� Varchar2(50) Path '$.�������ղ���ʮ����',�������ղ����ʮ���� Varchar2(50) Path '$.�������ղ����ʮ����',
                          ��ѹ������ʮ�� Varchar2(50) Path '$.��ѹ������ʮ��',��ѹ������ʮ�� Varchar2(50) Path '$.��ѹ������ʮ��',��ѹ������ʮ�� Varchar2(50) Path '$.��ѹ������ʮ��',
                          ��ѹ���������� Varchar2(50) Path '$.��ѹ����������',��ѹ���������� Varchar2(50) Path '$.��ѹ����������',��ѹ��������� Varchar2(50) Path '$.��ѹ���������',
                          ��ѹ����ʮ���� Varchar2(50) Path '$.��ѹ����ʮ����',��ѹ������ʮ���� Varchar2(50) Path '$.��ѹ������ʮ����',
                          ���ܲ������̥����ʮ�� Varchar2(50) Path '$.���ܲ������̥����ʮ��',���ܲ������̥����ʮ�� Varchar2(50) Path '$.���ܲ������̥����ʮ��',���ܲ������̥���ʮ�� Varchar2(50) Path '$.���ܲ������̥���ʮ��',
                          ���ܲ������̥�������� Varchar2(50) Path '$.���ܲ������̥��������',���ܲ������̥�������� Varchar2(50) Path '$.���ܲ������̥��������',���ܲ������̥������� Varchar2(50) Path '$.���ܲ������̥�������',
                          ���ܲ������̥��ʮ���� Varchar2(50) Path '$.���ܲ������̥��ʮ����',���ܲ������̥���ʮ���� Varchar2(50) Path '$.���ܲ������̥���ʮ����',
                          ��ѹͨ����ʮ�� Varchar2(50) Path '$.��ѹͨ����ʮ��',��ѹͨ����ʮ�� Varchar2(50) Path '$.��ѹͨ����ʮ��',��ѹͨ����ʮ�� Varchar2(50) Path '$.��ѹͨ����ʮ��',
                          ��ѹͨ�������� Varchar2(50) Path '$.��ѹͨ��������',��ѹͨ�������� Varchar2(50) Path '$.��ѹͨ��������',��ѹͨ������� Varchar2(50) Path '$.��ѹͨ�������',
                          ��ѹͨ��ʮ���� Varchar2(50) Path '$.��ѹͨ��ʮ����',��ѹͨ����ʮ���� Varchar2(50) Path '$.��ѹͨ����ʮ����',
                          ���ܲ����ʮ�� Varchar2(50) Path '$.���ܲ����ʮ��',���ܲ����ʮ�� Varchar2(50) Path '$.���ܲ����ʮ��',���ܲ�ܾ�ʮ�� Varchar2(50) Path '$.���ܲ�ܾ�ʮ��',
                          ���ܲ�������� Varchar2(50) Path '$.���ܲ��������',���ܲ�������� Varchar2(50) Path '$.���ܲ��������',���ܲ������� Varchar2(50) Path '$.���ܲ�������',
                          ���ܲ��ʮ���� Varchar2(50) Path '$.���ܲ��ʮ����',���ܲ�ܶ�ʮ���� Varchar2(50) Path '$.���ܲ�ܶ�ʮ����',
                          ���ⰴѹ��ʮ�� Varchar2(50) Path '$.���ⰴѹ��ʮ��',���ⰴѹ��ʮ�� Varchar2(50) Path '$.���ⰴѹ��ʮ��',���ⰴѹ��ʮ�� Varchar2(50) Path '$.���ⰴѹ��ʮ��',
                          ���ⰴѹ������ Varchar2(50) Path '$.���ⰴѹ������',���ⰴѹ������ Varchar2(50) Path '$.���ⰴѹ������',���ⰴѹ����� Varchar2(50) Path '$.���ⰴѹ�����',
                          ���ⰴѹʮ���� Varchar2(50) Path '$.���ⰴѹʮ����',���ⰴѹ��ʮ���� Varchar2(50) Path '$.���ⰴѹ��ʮ����',
                          ����������ʮ�� Varchar2(50) Path '$.����������ʮ��',����������ʮ�� Varchar2(50) Path '$.����������ʮ��',�������ؾ�ʮ�� Varchar2(50) Path '$.�������ؾ�ʮ��',
                          �������������� Varchar2(50) Path '$.��������������',�������������� Varchar2(50) Path '$.��������������',������������� Varchar2(50) Path '$.�������������',
                          ��������ʮ���� Varchar2(50) Path '$.��������ʮ����',�������ض�ʮ���� Varchar2(50) Path '$.�������ض�ʮ����',
                          ������ˮ��ʮ�� Varchar2(50) Path '$.������ˮ��ʮ��',������ˮ��ʮ�� Varchar2(50) Path '$.������ˮ��ʮ��',������ˮ��ʮ�� Varchar2(50) Path '$.������ˮ��ʮ��',
                          ������ˮ������ Varchar2(50) Path '$.������ˮ������',������ˮ������ Varchar2(50) Path '$.������ˮ������',������ˮ����� Varchar2(50) Path '$.������ˮ�����',
                          ������ˮʮ���� Varchar2(50) Path '$.������ˮʮ����',������ˮ��ʮ���� Varchar2(50) Path '$.������ˮ��ʮ����' )) As F
Where a.Mid = b.Mid
Order By Addtime;

create or replace force view ZLSOL_JC.v_sol_inf_delivery as
Select a.Mid, b."����1", b."���̿�ʼʱ��", b."����ȫ��ʱ��", b."̥�����ʱ��", b."̥�����ʱ��", b."��һ����", b."�ڶ�����", b."��������", b."�������",
       b."����", b."��Ĥ��ʽ", b."��Ĥʱ��", b."��ˮ��״", b."��ˮ��", b."��ˮ��ɫ", b."̥�������ʽ", b."̥�̰��뷽ʽ", b."̥��������",
       b."̥��̥Ĥ����", b."̥�����", b."̥����̬", b."̥�̴�С", b."̥������", b."�������", b."�������", b."�����ٽ�", b."���",b."�ƾ�����",b."�ƾ�����1",b."�ƻ��̶�",b."�����ʽ",
       b."���̥��λ", b."������С", b."������λ", b."�������˳̶�", b."���������п�", b."�������˷��", b."������������", b."�������˳���", b."�������˲�λ", b."��������״��",
       b."�������˲�λ��С",b."�������˲�λ����", b."��������Ѫ�ײ�λ", b."��������Ѫ�״�С",

       b."����",b."��ˮ��",b."�������������",b."��������",b."�������������1",b."��������1",b."����",
        b."��������30�������", b."��������30������",b."��������30����ɫ",
        b."��ѹͨ��30�������",b."��ѹͨ��30������",b."��ѹͨ��30����ɫ",
        b."������ѹͨ�����������",b."������ѹͨ����������",b."������ѹͨ��������ɫ",
        b."��ѹͨ�������ⰴѹ30�������", b."��ѹͨ�������ⰴѹ30������", b."��ѹͨ�������ⰴѹ30����ɫ",
        b."ʹ���������غ�����",b."ʵʩ������Ҫ��ʩ�������",

       b.�������ղ�����ʮ��,b.�������ղ�����ʮ��,b.�������ղ����ʮ��,b.�������ղ���������,b.�������ղ���������,b.�������ղ��������,b.�������ղ���ʮ����,b.�������ղ����ʮ����,
       b.��ѹ������ʮ��,b.��ѹ������ʮ��,b.��ѹ������ʮ��,b.��ѹ����������,b.��ѹ����������,b.��ѹ���������,b.��ѹ����ʮ����,b.��ѹ������ʮ����,
       b.���ܲ������̥����ʮ��,b.���ܲ������̥����ʮ��,b.���ܲ������̥���ʮ��,b.���ܲ������̥��������,b.���ܲ������̥��������,b.���ܲ������̥�������,b.���ܲ������̥��ʮ����,b.���ܲ������̥���ʮ����,
       b.��ѹͨ����ʮ��,b.��ѹͨ����ʮ��,b.��ѹͨ����ʮ��,b.��ѹͨ��������,b.��ѹͨ��������,b.��ѹͨ�������,b.��ѹͨ��ʮ����,b.��ѹͨ����ʮ����,
       b.���ܲ����ʮ��,b.���ܲ����ʮ��,b.���ܲ�ܾ�ʮ��,b.���ܲ��������,b.���ܲ��������,b.���ܲ�������,b.���ܲ��ʮ����,b.���ܲ�ܶ�ʮ����,
       b.���ⰴѹ��ʮ��,b.���ⰴѹ��ʮ��,b.���ⰴѹ��ʮ��,b.���ⰴѹ������,b.���ⰴѹ������,b.���ⰴѹ�����,b.���ⰴѹʮ����,b.���ⰴѹ��ʮ����,
       b.����������ʮ��,b.����������ʮ��,b.�������ؾ�ʮ��,b.��������������,b.��������������,b.�������������,b.��������ʮ����,b.�������ض�ʮ����,
       b.������ˮ��ʮ��,b.������ˮ��ʮ��,b.������ˮ��ʮ��,b.������ˮ������,b.������ˮ������,b.������ˮ�����,b.������ˮʮ����,b.������ˮ��ʮ����,

       b."����ʱ��",b."���տ�ʼʱ��",b."���ս���ʱ��",b."����ǰʱ��",b."�����ʱ��",b."��Ҫ������Ա",b."���Ƚ��",

       b."ĸӤ��Ӵ�����˱��ʼʱ��",b."ĸӤ��Ӵ�����˱����ʱ��", b."����Ѫѹ", b."����ѹ", b."����ѹ", b."������Ѫ",b."��Ѫ����", b."��ʱ��ҩ", b."������ҩ",
       b."�������",b."�������2",b."�������3",b."�������4",b."��������", b."��������2",b."��������3",b."��������4",
       b."�������", b."δ��˱ԭ��",d."����3",d."������ʱ��",d."��������������", d."������", d."������", d."��¼��"
From Sol_Inf_Delivery A,
     Json_Table(Nvl(a.Newborndetail, '{"����1":"1"}'),
                 '$'
                  Columns(����1 Varchar2(50) Path '$.����1', ���̿�ʼʱ�� Varchar2(19) Path '$.���̿�ʼʱ��',
                          ����ȫ��ʱ�� Varchar2(19) Path '$.����ȫ��ʱ��', ̥�����ʱ�� Varchar2(19) Path '$.̥�����ʱ��',
                          ̥�����ʱ�� Varchar2(19) Path '$.̥�����ʱ��', ��һ���� Varchar2(50) Path '$.��һ����',
                          �ڶ����� Varchar2(50) Path '$.�ڶ�����', �������� Varchar2(50) Path '$.��������',���� Varchar2(50) Path '$.����',
                          ��Ĥ��ʽ Varchar2(50) Path '$.��Ĥ��ʽ', ��Ĥʱ�� Varchar2(19) Path '$.��Ĥʱ��', ��ˮ��״ Varchar2(50) Path '$.��ˮ��״',
                          ��ˮ�� Varchar2(50) Path '$.��ˮ��', ��ˮ��ɫ Varchar2(50) Path '$.��ˮ��ɫ',
                          ̥�������ʽ Varchar2(50) Path '$.̥�������ʽ', ̥�̰��뷽ʽ Varchar2(50) Path '$.̥�̰��뷽ʽ',
                          ̥�������� Varchar2(50) Path '$.̥��������', ̥��̥Ĥ���� Varchar2(50) Path '$.̥��̥Ĥ����',
                          ̥����� Varchar2(50) Path '$.̥�����', ̥����̬ Varchar2(50) Path '$.̥����̬', ̥�̴�С Varchar2(50) Path '$.̥�̴�С',
                          ̥������ Varchar2(50) Path '$.̥������', ������� Varchar2(50) Path '$.�������', ������� Varchar2(50) Path '$.�������',
                         �����ٽ� Varchar2(50) Path '$.�����ٽ�',��� Varchar2(50) Path '$.���', �ƾ����� Varchar2(50) Path '$.�ƾ�����',�ƾ�����1 Varchar2(50) Path '$.�ƾ�����1',
                         �ƻ��̶� Varchar2(20) Path '$.�ƻ��̶�',
                          �����ʽ Varchar2(50) Path '$.�����ʽ',
                          ���̥��λ Varchar2(50) Path '$.���̥��λ', ������С Varchar2(50) Path '$.������С',
                          ������λ Varchar2(50) Path '$.������λ', �������˳̶� Varchar2(50) Path '$.�������˳̶�',
                          ���������п� Varchar2(50) Path '$.���������п�', �������˷�� Varchar2(50) Path '$.�������˷��',
                          ������������ Varchar2(50) Path '$.������������', �������˳��� Varchar2(50) Path '$.�������˳���',
                          �������˲�λ Varchar2(50) Path '$.�������˲�λ', ��������״�� Varchar2(50) Path '$.��������״��',
                          �������˲�λ��С Varchar2(50) Path '$.�������˲�λ��С', �������˲�λ���� Varchar2(50) Path '$.�������˲�λ����',
                          ��������Ѫ�ײ�λ Varchar2(50) Path '$.��������Ѫ�ײ�λ', ��������Ѫ�״�С Varchar2(50) Path '$.��������Ѫ�״�С',

                          ���� Varchar2(50) Path '$.����', ��ˮ�� Varchar2(50) Path '$.��ˮ��',������������� Varchar2(50) Path '$.�������������', �������� Varchar2(50) Path '$.��������',
                          �������������1 Varchar2(50) Path '$.�������������1', ��������1 Varchar2(50) Path '$.��������1', ���� Varchar2(50) Path '$.����',

                          ��������30������� Varchar2(50) Path '$.��������30�������',��������30������ Varchar2(50) Path '$.��������30������',��������30����ɫ Varchar2(50) Path '$.��������30����ɫ',
                          ��ѹͨ��30������� Varchar2(50) Path '$.��ѹͨ��30�������',��ѹͨ��30������ Varchar2(50) Path '$.��ѹͨ��30�������',��ѹͨ��30����ɫ Varchar2(50) Path '$.��ѹͨ��30�������',
                          ������ѹͨ����������� Varchar2(50) Path '$.������ѹͨ�����������',������ѹͨ���������� Varchar2(50) Path '$.������ѹͨ����������',������ѹͨ��������ɫ Varchar2(50) Path '$.������ѹͨ��������ɫ',
                          ��ѹͨ�������ⰴѹ30������� Varchar2(50) Path '$.������ѹͨ�����������',��ѹͨ�������ⰴѹ30������ Varchar2(50) Path '$.������ѹͨ����������',��ѹͨ�������ⰴѹ30����ɫ Varchar2(50) Path '$.������ѹͨ��������ɫ',
                          ʹ���������غ����� Varchar2(50) Path '$.ʹ���������غ�����',ʵʩ������Ҫ��ʩ������� Varchar2(50) Path '$.ʵʩ������Ҫ��ʩ�������',

                          �������ղ�����ʮ�� Varchar2(50) Path '$.�������ղ�����ʮ��',�������ղ�����ʮ�� Varchar2(50) Path '$.�������ղ�����ʮ��',�������ղ����ʮ�� Varchar2(50) Path '$.�������ղ����ʮ��',
                          �������ղ��������� Varchar2(50) Path '$.�������ղ���������',�������ղ��������� Varchar2(50) Path '$.�������ղ���������',�������ղ�������� Varchar2(50) Path '$.�������ղ��������',
                          �������ղ���ʮ���� Varchar2(50) Path '$.�������ղ���ʮ����',�������ղ����ʮ���� Varchar2(50) Path '$.�������ղ����ʮ����',
                          ��ѹ������ʮ�� Varchar2(50) Path '$.��ѹ������ʮ��',��ѹ������ʮ�� Varchar2(50) Path '$.��ѹ������ʮ��',��ѹ������ʮ�� Varchar2(50) Path '$.��ѹ������ʮ��',
                          ��ѹ���������� Varchar2(50) Path '$.��ѹ����������',��ѹ���������� Varchar2(50) Path '$.��ѹ����������',��ѹ��������� Varchar2(50) Path '$.��ѹ���������',
                          ��ѹ����ʮ���� Varchar2(50) Path '$.��ѹ����ʮ����',��ѹ������ʮ���� Varchar2(50) Path '$.��ѹ������ʮ����',
                          ���ܲ������̥����ʮ�� Varchar2(50) Path '$.���ܲ������̥����ʮ��',���ܲ������̥����ʮ�� Varchar2(50) Path '$.���ܲ������̥����ʮ��',���ܲ������̥���ʮ�� Varchar2(50) Path '$.���ܲ������̥���ʮ��',
                          ���ܲ������̥�������� Varchar2(50) Path '$.���ܲ������̥��������',���ܲ������̥�������� Varchar2(50) Path '$.���ܲ������̥��������',���ܲ������̥������� Varchar2(50) Path '$.���ܲ������̥�������',
                          ���ܲ������̥��ʮ���� Varchar2(50) Path '$.���ܲ������̥��ʮ����',���ܲ������̥���ʮ���� Varchar2(50) Path '$.���ܲ������̥���ʮ����',
                          ��ѹͨ����ʮ�� Varchar2(50) Path '$.��ѹͨ����ʮ��',��ѹͨ����ʮ�� Varchar2(50) Path '$.��ѹͨ����ʮ��',��ѹͨ����ʮ�� Varchar2(50) Path '$.��ѹͨ����ʮ��',
                          ��ѹͨ�������� Varchar2(50) Path '$.��ѹͨ��������',��ѹͨ�������� Varchar2(50) Path '$.��ѹͨ��������',��ѹͨ������� Varchar2(50) Path '$.��ѹͨ�������',
                          ��ѹͨ��ʮ���� Varchar2(50) Path '$.��ѹͨ��ʮ����',��ѹͨ����ʮ���� Varchar2(50) Path '$.��ѹͨ����ʮ����',
                          ���ܲ����ʮ�� Varchar2(50) Path '$.���ܲ����ʮ��',���ܲ����ʮ�� Varchar2(50) Path '$.���ܲ����ʮ��',���ܲ�ܾ�ʮ�� Varchar2(50) Path '$.���ܲ�ܾ�ʮ��',
                          ���ܲ�������� Varchar2(50) Path '$.���ܲ��������',���ܲ�������� Varchar2(50) Path '$.���ܲ��������',���ܲ������� Varchar2(50) Path '$.���ܲ�������',
                          ���ܲ��ʮ���� Varchar2(50) Path '$.���ܲ��ʮ����',���ܲ�ܶ�ʮ���� Varchar2(50) Path '$.���ܲ�ܶ�ʮ����',
                          ���ⰴѹ��ʮ�� Varchar2(50) Path '$.���ⰴѹ��ʮ��',���ⰴѹ��ʮ�� Varchar2(50) Path '$.���ⰴѹ��ʮ��',���ⰴѹ��ʮ�� Varchar2(50) Path '$.���ⰴѹ��ʮ��',
                          ���ⰴѹ������ Varchar2(50) Path '$.���ⰴѹ������',���ⰴѹ������ Varchar2(50) Path '$.���ⰴѹ������',���ⰴѹ����� Varchar2(50) Path '$.���ⰴѹ�����',
                          ���ⰴѹʮ���� Varchar2(50) Path '$.���ⰴѹʮ����',���ⰴѹ��ʮ���� Varchar2(50) Path '$.���ⰴѹ��ʮ����',
                          ����������ʮ�� Varchar2(50) Path '$.����������ʮ��',����������ʮ�� Varchar2(50) Path '$.����������ʮ��',�������ؾ�ʮ�� Varchar2(50) Path '$.�������ؾ�ʮ��',
                          �������������� Varchar2(50) Path '$.��������������',�������������� Varchar2(50) Path '$.��������������',������������� Varchar2(50) Path '$.�������������',
                          ��������ʮ���� Varchar2(50) Path '$.��������ʮ����',�������ض�ʮ���� Varchar2(50) Path '$.�������ض�ʮ����',
                          ������ˮ��ʮ�� Varchar2(50) Path '$.������ˮ��ʮ��',������ˮ��ʮ�� Varchar2(50) Path '$.������ˮ��ʮ��',������ˮ��ʮ�� Varchar2(50) Path '$.������ˮ��ʮ��',
                          ������ˮ������ Varchar2(50) Path '$.������ˮ������',������ˮ������ Varchar2(50) Path '$.������ˮ������',������ˮ����� Varchar2(50) Path '$.������ˮ�����',
                          ������ˮʮ���� Varchar2(50) Path '$.������ˮʮ����',������ˮ��ʮ���� Varchar2(50) Path '$.������ˮ��ʮ����',

                          ����ʱ�� Varchar2(19) Path '$.����ʱ��',���տ�ʼʱ�� Varchar2(19) Path '$.���տ�ʼʱ��',���ս���ʱ�� Varchar2(19) Path '$.���ս���ʱ��',
                          ����ǰʱ�� Varchar2(19) Path '$.����ǰʱ��', �����ʱ�� Varchar2(19) Path '$.�����ʱ��',
                          ��Ҫ������Ա Varchar2(19) Path '$.��Ҫ������Ա',���Ƚ�� Varchar2(19) Path '$.��Ҫ������Ա',

                          ������� Varchar2(50) Path '$.�������', ĸӤ��Ӵ�����˱��ʼʱ�� Varchar2(50) Path '$.ĸӤ��Ӵ�����˱��ʼʱ��',
                          ĸӤ��Ӵ�����˱����ʱ�� Varchar2(50) Path '$.ĸӤ��Ӵ�����˱����ʱ��',����Ѫѹ Varchar2(50) Path '$.����Ѫѹ',
                          ����ѹ Varchar2(50) Path '$.����ѹ',
                          ����ѹ Varchar2(50) Path '$.����ѹ',
                          ������Ѫ Varchar2(50) Path '$.������Ѫ', ��Ѫ���� Varchar2(50) Path '$.��Ѫ����', ��ʱ��ҩ Varchar2(50) Path '$.��ʱ��ҩ', ������ҩ Varchar2(50) Path '$.������ҩ',
                          ������� Varchar2(50) Path '$.�������', �������2 Varchar2(50) Path '$.�������2', �������3 Varchar2(50) Path '$.�������3', �������4 Varchar2(50) Path '$.�������4',
                          �������� Varchar2(50) Path '$.��������',��������2 Varchar2(50) Path '$.��������2', ��������3 Varchar2(50) Path '$.��������3', ��������4 Varchar2(50) Path '$.��������4',
                          ������� Varchar2(50) Path '$.�������',δ��˱ԭ�� Varchar2(50) Path '$.δ��˱ԭ��'
                          )) As B,
     Json_Table(Nvl(a.Deliveryinf, '{"����3":"1"}'),
                 '$' Columns(����3 Varchar2(1) Path '$.����3', ������ʱ�� Varchar2(50) Path '$.������ʱ��',�������������� Varchar2(50) Path '$.��������������',
                          ������ Varchar2(50) Path '$.������', ������ Varchar2(50) Path '$.������', ��¼�� Varchar2(50) Path '$.��¼��')
                          ) As D;

create or replace force view ZLSOL_JC.v_sol_rs_birth as
Select a.mid,b."�Ѵ�",b."����",b."Ѫ��",b."��������ʷ",b."ĩ���¾�",b."Ԥ����",b."��ǰ�ϼ��侶",b."���ռ侶",b."���ǽ�ڼ侶",b."�����⾶",b."���ǻ���",b."���ǹؽ�",b."�����м�",b."������",b."����֢",b."��ǰ��¼����",b."���ʱ��",b."Ѫѹ",b."����ѹ",b."����ѹ",b."����",b."����",b."̥����",b."̥����С",b."����������",b."̥λ",b."�ν�",b."��Ĥ���",b."��¶",b."����",b."�����",b."������ʼʱ��",b."��Ĥʱ��",b."��Ժ����"
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
Ѫѹ            Varchar2(10) PATH '$.Ѫѹ',
����ѹ            Varchar2(10) PATH '$.����ѹ',
����ѹ            Varchar2(10) PATH '$.����ѹ ',
����            Number(4,2)  PATH '$.����',
����            Varchar2(10) PATH '$.����',
̥����            Varchar2(10) PATH '$.̥����',
̥����С        Number(5,2) PATH '$.̥����С',
����������      Varchar2(10) PATH '$.����������',
̥λ            Varchar2(10) PATH '$.̥λ',
�ν�            Varchar2(10) PATH '$.�ν�',
��Ĥ���        Varchar2(10) PATH '$.��Ĥ���',
��¶            Varchar2(2) PATH '$.��¶',
����            Number(4,2) PATH '$.����',
�����          Varchar2(50) PATH '$.�����',
������ʼʱ��    Varchar2(20) PATH '$.������ʼʱ��',
��Ĥʱ��        Varchar2(20) PATH '$.��Ĥʱ��',
��Ժ����        Varchar2(100) PATH '$.��Ժ����'
)) as b;

create or replace force view ZLSOL_JC.v_his_������������¼ as
select d.pid ����id,d.tid סԺ����,b.Babyno ���,d.name||decode(b.Sex,'��','֮��','֮Ů')||t.˳�� as Ӥ������,
b.Sex as Ӥ���Ա�,
c.�Ѵ� as �������,
a.�����ʽ,b.̥��״��,
b.��,b.����,b.Ѫ��,
a.̥�����ʱ�� as ����ʱ��,
b.����ʱ��,'' as ��ע˵��,
b.Recorder �Ǽ���,
b.Addtime �Ǽ�ʱ��
from V_SOL_INF_DELIVERY a,v_newborn b,v_sol_rs_birth c,sol_inf_puerpera d,
(select mid,sex,babyno,row_number() over(partition by mid,sex order by babyno) ˳�� from SOL_INF_NEWBORNS t  ) t
where a.Mid=b.Mid and a.mid=d.mid and a.mid=t.mid and b.Babyno=t.babyno
and b.Mid=c.mid(+);

create or replace force view ZLSOL_JC.v_sol_inf_checkinroom as
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

create or replace force view ZLSOL_JC.v_sol_inf_checkoutroom as
Select a.mid, b."OUTROOMTIME",b."����״̬",b."ҽ�Ʋ���",b."������",b."����ͨ��",b."�ֲ����",b."��������",b."�����п���",b."�����пڷ��",b."����ˮ��",b."����Ѫ��",b."������Ѫ",b."��Ѫ��",
b."��������ɴ��",b."�������",b."�������",b."����ҩ��",b."������",b."�Ӱ���",b."ҩ��",b."��ע"
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
������Ѫ     Varchar2(10) PATH '$.������Ѫ',
��Ѫ��       Number(5) PATH '$.��Ѫ��',
��������ɴ��   Varchar2(20) PATH '$.��������ɴ��',
�������     Varchar2(20) PATH '$.�������',
�������     Varchar2(20) PATH '$.�������',
����ҩ��     Varchar2(50) PATH '$.����ҩ��',
������       Varchar2(20) PATH '$.������',
�Ӱ���       Varchar2(20) PATH '$.�Ӱ���',
ҩ��         Varchar2(50) PATH '$.ҩ��',
��ע         Varchar2(50) PATH '$.��ע'
)) as b;

create or replace force view ZLSOL_JC.v_sol_inf_equipment as
Select a.mid, b."���м���ǰ",b."���м�����",b."���м�����",b."�������ǰ",b."���������",b."���������",
b."ֹѪǯ��ǰ",b."ֹѪǯ����",b."ֹѪǯ����",b."������ǰ",b."��������",b."��������",
b."��������ǰ",b."����������",b."����������",b."�������ǰ",b."����������",b."���������",b."ϴ�����ǰ",
b."ϴ��������",b."ϴ�������",b."̥����ǰ",b."̥������",b."̥������",b."������ǰ",b."���������",b."��������",
b."������ǰ",b."��������",b."��������",b."ɴ����ǰ",b."ɴ������",b."ɴ������",b."��Բǯ��ǰ",b."��Բǯ����",b."��Բǯ����",
b."����ǯ��ǰ",b."����ǯ����",b."����ǯ����",
b."������ǰ",b."��������",b."��������",b."�γײ�ǰ",b."�γ�����",b."�γײ���",b."����˹��ǰ",b."����˹����",b."����˹����",
b."��ǯ��ǰ",b."��ǯ����",b."��ǯ����"
From SOL_INF_Equipment a,JSON_TABLE(a.Content,'$' columns(
���м���ǰ   varchar2(2) PATH '$.���м���ǰ',
���м�����   varchar2(2) PATH '$.���м�����',
���м�����   varchar2(2) PATH '$.���м�����',
�������ǰ   varchar2(2) PATH '$.�������ǰ',
���������   varchar2(2) PATH '$.���������',
���������   varchar2(2) PATH '$.���������',
ֹѪǯ��ǰ   varchar2(2) PATH '$.ֹѪǯ��ǰ',
ֹѪǯ����   varchar2(2) PATH '$.ֹѪǯ����',
ֹѪǯ����   varchar2(2) PATH '$.ֹѪǯ����',
������ǰ   varchar2(2) PATH '$.������ǰ',
��������   varchar2(2) PATH '$.��������',
��������   varchar2(2) PATH '$.��������',
��������ǰ   varchar2(2) PATH '$.��������ǰ',
����������   varchar2(2) PATH '$.����������',
����������   varchar2(2) PATH '$.����������',
�������ǰ   varchar2(2) PATH '$.�������ǰ',
����������   varchar2(2) PATH '$.����������',
���������   varchar2(2) PATH '$.���������',
ϴ�����ǰ   varchar2(2) PATH '$.ϴ�����ǰ',
ϴ��������   varchar2(2) PATH '$.ϴ��������',
ϴ�������   varchar2(2) PATH '$.ϴ�������',
̥����ǰ   varchar2(2) PATH '$.̥����ǰ',
̥������   varchar2(2) PATH '$.̥������',
̥������   varchar2(2) PATH '$.̥������',
������ǰ   varchar2(2) PATH '$.������ǰ',
���������   varchar2(2) PATH '$.���������',
��������   varchar2(2) PATH '$.��������',
������ǰ   varchar2(2) PATH '$.������ǰ',
��������   varchar2(2) PATH '$.��������',
��������   varchar2(2) PATH '$.��������',
��ǯ��ǰ   varchar2(2) PATH '$.��ǯ��ǰ',
��ǯ����   varchar2(2) PATH '$.��ǯ����',
��ǯ����   varchar2(2) PATH '$.��ǯ����',
ɴ����ǰ   varchar2(2) PATH '$.ɴ����ǰ',
ɴ������   varchar2(2) PATH '$.ɴ������',
ɴ������   varchar2(2) PATH '$.ɴ������',
��Բǯ��ǰ   varchar2(2) PATH '$.��Բǯ��ǰ',
��Բǯ����   varchar2(2) PATH '$.��Բǯ����',
��Բǯ����   varchar2(2) PATH '$.��Բǯ����',
����ǯ��ǰ   varchar2(2) PATH '$.����ǯ��ǰ',
����ǯ����   varchar2(2) PATH '$.����ǯ����',
����ǯ����   varchar2(2) PATH '$.����ǯ����',
������ǰ   varchar2(2) PATH '$.������ǰ',
��������   varchar2(2) PATH '$.��������',
��������   varchar2(2) PATH '$.��������',
�γײ�ǰ   varchar2(2) PATH '$.�γײ�ǰ',
�γ�����   varchar2(2) PATH '$.�γ�����',
�γײ���   varchar2(2) PATH '$.�γײ���',
����˹��ǰ   varchar2(2) PATH '$.����˹��ǰ',
����˹����   varchar2(2) PATH '$.����˹����',
����˹����   varchar2(2) PATH '$.����˹����'
)) as b;

create or replace force view ZLSOL_JC.v_sol_inf_newborns as
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

create or replace force view ZLSOL_JC.v_sol_inf_puerpera as
Select Name, Mid, Old, LPad(Bedno, 10) Bedno, Pno, Diagnosis, Status, Decode(Expectant, 1, '��', '') ����,
       Decode(Checkinroom, 1, '��', '') �뷿, Decode(Birth, 1, '��', '') �ٲ�, Decode(Druglabor, 1, '��', '') ����,
       Decode(Delivery, 1, '��', '') ����, Decode(Newborns, 1, '��', '') ������, Decode(Postpartum, 1, '��', '') ����,
       Decode(Checkoutroom, 1, '��', '') ����,Decode(Equipment, 1, '��', '') ��е,outtime,pid,tid
From Sol_Inf_Puerpera;

create or replace force view ZLSOL_JC.v_sol_rs_birth_course as
Select  a.courseid,a.mid,b."���ʱ��",b."̥��λ",b."Ѫѹ",b."����ѹ",b."����ѹ",b."Ѫ��",b."����",b."����",b."����",b."̥����",b."����ǿ��",b."��������",b."�������",b."��������ʧ",b."����",b."��ˮ",b."��Ĥ���",b."�������",b."��¶",b."Ѫ�����Ͷ�",b."��ʶ",b."����",b."�����"
From SOL_RS_BIRTH_COURSE a,JSON_TABLE(a.CONTENT,'$' columns(
���ʱ��        Varchar2(20)  PATH '$.���ʱ��',
̥��λ Varchar2(20)  PATH '$.̥��λ',
Ѫѹ        Varchar2(10) PATH '$.Ѫѹ',

����ѹ        Varchar2(10) PATH '$.����ѹ',
����ѹ        Varchar2(10) PATH '$.����ѹ',
Ѫ��          Varchar2(10) PATH '$.Ѫ��',
����          Varchar2(10) PATH  '$.����',
����        Number(4,2)  PATH '$.����',
����        Varchar2(10) PATH '$.����',
̥����        Varchar2(10) PATH '$.̥����',
����ǿ��    Varchar2(10) PATH '$.����ǿ��',
��������  Varchar2(10) PATH '$.��������',
�������  Varchar2(10) PATH '$.�������',
��������ʧ    Varchar2(20) PATH '$.��������ʧ',
����        Number(4,2) PATH '$.����',
��¶        Number(2) PATH '$.��¶'��
Ѫ�����Ͷ�     Varchar2(10) PATH '$.Ѫ�����Ͷ�'��
��ʶ     Varchar2(10) PATH '$.��ʶ'��
��ˮ        Varchar2(10) PATH '$.��ˮ'��
��Ĥ���    Varchar2(10) PATH '$.��Ĥ���',
�������    Varchar2(20) PATH '$.�������',
����        Varchar2(500) PATH '$.����'��
�����      Varchar2(50) PATH '$.�����'
)) as b;

create or replace force view ZLSOL_JC.v_sol_rs_druglabor as
Select Mid, To_Char(����, 'YYYY-MM-DD HH24:MI') ����, ����ָ��, �������� from Sol_Rs_Druglabor;

create or replace force view ZLSOL_JC.v_sol_rs_druglabor_list as
Select a.Mid, a.Courseid ID, b."��¼ʱ��",b."Ѫѹ",b."����ѹ",b."����ѹ",b."����",b."̥����",b."����ǿ��",b."��������",b."�������",b."����",b."��ʶ",b."����",b."����",b."��¶",b."Ѫ��",b."Ѫ�����Ͷ�",b."��ˮ��",b."��ˮ��״",b."����",b."��¼��",b."����",b."����",b."�����ܳ���"
From Sol_Rs_Druglabor_List a,
     Json_Table(a.Content,'$' Columns(
     ��¼ʱ�� Varchar2(20) Path '$.��¼ʱ��',
     ���� Number(3,1) Path '$.����',
     ���� Number(3) Path '$.����',
     Ѫѹ Varchar2(7) Path '$.Ѫѹ',
     ����ѹ Varchar2(7) Path '$.����ѹ',
     ����ѹ Varchar2(7) Path '$.����ѹ',
     ���� Number(3) Path '$.����',
     ̥���� Number(3) Path '$.̥����',
     ����ǿ�� Varchar2(10) Path '$.����ǿ��',
     �������� Varchar2(20) Path '$.��������',
     ������� Varchar2(20) Path '$.�������',
     �����ܳ��� Varchar2(20) Path '$.�����ܳ���',
     ���� Number(3) Path '$.����',
     ��ʶ Varchar2(20) Path  '$.��ʶ',
     ���� Varchar2(20) Path  '$.����',
     ���� Varchar2(20) Path  '$.����',
     ��¶ Varchar2(10) Path '$.��¶',
     Ѫ�� Varchar2(10) Path '$.Ѫ��',
     Ѫ�����Ͷ� Varchar2(10) Path '$.Ѫ�����Ͷ�',

     ��ˮ�� Number(4) Path '$.��ˮ��',
     ��ˮ��״ Varchar2(10) Path '$.��ˮ��״',
     ���� Varchar2(500) Path '$.����',
     ��¼�� Varchar2(100) Path '$.��¼��')) b;

create or replace force view ZLSOL_JC.v_sol_rs_expectant as
Select a.mid,a.courseid,b."��¼ʱ��",b."����",b."Ѫѹ",b."����ѹ",b."����ѹ",b."����",b."����",b."Ѫ��",b."����",b."��Χ",b."̥��������",b."̥��������",b."̥��������",b."̥����",b."��¶",b."����",b."��������ʧ",b."��Ĥ���",b."��ˮ��״",b."����ǿ��",b."��������",b."�������",b."����",b."�����"
From SOL_RS_EXPECTANT a,JSON_TABLE(a.Content,'$' columns(
��¼ʱ��    Varchar2(50) PATH '$.��¼ʱ��',
����  Varchar2(20) PATH '$.����',
Ѫѹ     Varchar2(20) PATH '$.Ѫѹ',
����ѹ   Varchar2(20) PATH '$.����ѹ',
����ѹ   Varchar2(20) PATH '$.����ѹ',

����   Varchar2(20) PATH '$.����',
����   Varchar2(20) PATH '$.����',

Ѫ��     Varchar2(20) PATH '$.Ѫ��',
����     Number(4,2) PATH '$.����',
��Χ     Varchar2(20) PATH '$.��Χ',
̥��������     Number(3) PATH '$.̥��������',
̥��������     Number(3) PATH '$.̥��������',
̥��������   Number(3) PATH '$.̥��������',
̥���� Number(3) PATH '$.̥����',
��¶     Varchar2(20) PATH '$.��¶',
����     Varchar2(20) PATH '$.����',
��������ʧ     Varchar2(20) PATH '$.��������ʧ',
��Ĥ���     Varchar2(20) PATH '$.��Ĥ���',
��ˮ��״      Varchar2(20) PATH '$.��ˮ��״',
����ǿ��     Varchar2(20) PATH '$.����ǿ��',
��������       Varchar2(20) PATH '$.��������',
�������       Varchar2(20) PATH '$.�������',
����     Varchar2(500) PATH '$.����',
�����       Varchar2(20) PATH '$.�����'
)) as b;

create or replace force view ZLSOL_JC.v_sol_rs_postpartum as
Select a.Mid, ��������, �����ʱ��, ���䷽ʽ, ������ʱ��, ������ʱbp, ������ʱ����, ������ʱ��������, ������ʱ������Ѫ, ������ʱһ�����, ����,  ����
From Sol_Rs_Postpartum A,
     Json_Table(a.Content,
                 '$' Columns(�������� varchar2(20) Path '$.��������', �����ʱ�� varchar2(20) Path '$.�����ʱ��', ���䷽ʽ Varchar2(20) Path '$.���䷽ʽ',
                          ������ʱ�� varchar2(20) Path '$.������ʱ��', ������ʱbp varchar2(7) Path '$.������ʱBP', ������ʱ���� Number(3) Path '$.������ʱ����',
                          ������ʱ�������� Number(2) Path '$.������ʱ��������', ������ʱ������Ѫ Number(3) Path '$.������ʱ������Ѫ',
                          ������ʱһ����� Varchar2(10) Path '$.������ʱһ�����', ���� Varchar2(20) Path '$.����', ���� Varchar2(10) Path '$.����'));

create or replace force view ZLSOL_JC.v_sol_rs_postpartum_jcfy_list as
Select a.Mid, a.Courseid, ��¼ʱ��, ��ʶ, ����, ����, ����, ����ѹ,����ѹ, Ѫ�����Ͷ�, Ѫ��, ����, ������Ѫ, ���׸߶�,   �������������, ǩ��
From Sol_Rs_Postpartum_List A,
     Json_Table(a.Content,
                 '$' Columns(��¼ʱ�� Varchar2(30) Path '$.��¼ʱ��', ��ʶ Varchar2(20) Path '$.��ʶ', ���� Varchar2(10) Path '$.����',
                          ���� Varchar2(50) Path '$.����', ���� Varchar2(30) Path '$.����', ����ѹ Varchar2(50) Path '$.����ѹ',����ѹ Varchar2(50) Path '$.����ѹ',
                          Ѫ�����Ͷ� Varchar2(10) Path '$.Ѫ�����Ͷ�', Ѫ�� Varchar2(20) Path '$.Ѫ��', ���� Varchar2(20) Path '$.����',
                          ������Ѫ Varchar2(10) Path '$.������Ѫ', ���׸߶� Varchar2(30) Path '$.���׸߶�', ������������� Varchar2(500) Path '$.�������������',
                          ǩ�� Varchar2(100) Path '$.ǩ��'));

create or replace force view ZLSOL_JC.v_sol_rs_postpartum_list as
Select a.Mid, a.Courseid ID, ��¼ʱ��, ����, �鷿����, ��ͷ, �ӹ�����, �ӹ�ѹʹ, ��¶��, ��¶��ɫ, ��¶��ζ, ��������, ��������, ��������, С��, ���, �������, ǩ��
From Sol_Rs_Postpartum_List A,
     Json_Table(a.Content,
                 '$' Columns(��¼ʱ�� Varchar2(20) Path '$.��¼ʱ��', ���� Number(4) Path '$.����', �鷿���� Varchar2(10) Path '$.�鷿����',
                          ��ͷ Varchar2(50) Path '$.��ͷ', �ӹ����� Number(3) Path '$.�ӹ�����', �ӹ�ѹʹ Varchar2(50) Path '$.�ӹ�ѹʹ',
                          ��¶�� Number(4) Path '$.��¶��', ��¶��ɫ Varchar2(20) Path '$.��¶��ɫ', ��¶��ζ Varchar2(20) Path '$.��¶��ζ',
                          �������� Varchar2(10) Path '$.��������', �������� Varchar2(10) Path '$.��������', �������� Varchar2(50) Path '$.��������',
                          С�� Varchar2(50) Path '$.С��', ��� Varchar2(50) Path '$.���', ������� Varchar2(100) Path '$.�������',
                          ǩ�� Varchar2(100) Path '$.ǩ��'));

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
  --���ܣ����ɶ��ŷָ��Ĳ������ŵ��ַ�����ת��Ϊ�������ݱ�
  --������STR_IN,��:G0000123,G0000124,G0000125...,SPLIT_IN,�ָ���,ȱʡΪ,��
  --˵����
  --1����SQL������漰��IN(����1, ����2,��) ���Ӿ�ʱʹ�����ַ�ʽ�Ա����ð󶨱�����
  --2��ʹ������������ʱ����Ҫ��SQL����м��롰/*+ cardinality(b 3)*/����ʾ����ΪCBO����ʱ�ڴ��û��ͳ������,��
  --3�����ֵ���ʾ��
  --SELECT /*+ cardinality(b 3)*/ * FROM ������ü�¼ WHERE NO IN (SELECT * FROM TABLE(F_STR2LIST('A01,A02,A03')) B);
  --SELECT /*+ cardinality(b 3)*/ A.* FROM ������ü�¼ A, TABLE(F_STR2LIST('A01,A02,A03')) B WHERE A.NO = B.COLUMN_VALUE;
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
  Expression In Varchar2, --��Ҫ�ָ��ַ���
  Delimiter  In Varchar2, --�ָ��ַ���
  Mimit      In Number := -1 --�ָ�λ
)
--���ܣ�ͨ������ʵ���ַ����ָ���ݴ���ָ�λʵ�ַָ��ַ���
  --������Expression����Ҫ�ָ��ַ�������Delimiter���ָ��ַ�����Mimit���ָ����Ӵ�����
  --���أ����طָ�λ�õ��ַ���
  --����л��
  --���ڣ�2010-12-24
  --�޸ģ�
  --�޸�����
 Return Varchar2 --���طָ��ַ���
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
      v_Error := '�±�ֵԽ�磡';
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

CREATE OR REPLACE Procedure ZLSOL_JC.his_�����������Ǽ�_revise
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
  select count(1) into n_count from ������������¼@to_his where ����id=n_����id and ��ҳid = n_��ҳid and ���=babyno_In;
  --�����������޸�
  if  state_in=2 then
   if n_count=0 then  ----����
      insert into   his_������������¼ 
      (����id,��ҳid,���,Ӥ������,Ӥ���Ա�,�������,���䷽ʽ,̥��״��,����ʱ��,��,����,Ѫ��,��ע˵��,����ʱ��,�Ǽ�ʱ��,�Ǽ���)
      select ����id,סԺ����,���,Ӥ������,Ӥ���Ա�,�������,�����ʽ,̥��״��,����ʱ��,��,����,Ѫ��,��ע˵��,����ʱ��,�Ǽ�ʱ��,�Ǽ���
       from   ( select  d.pid ����id,d.tid סԺ����,b.Babyno ���,d.name||decode(b.Sex,'��','֮��','֮Ů')||t.˳�� as Ӥ������,
              b.Sex as Ӥ���Ա�,c.�Ѵ� as �������,a.�����ʽ,b.̥��״��,b.��,b.����,b.Ѫ��,to_date(a.̥�����ʱ��,'yyyy-mm-dd hh24:mi:ss') as ����ʱ��,
              b.����ʱ��,'' as ��ע˵��,b.Recorder �Ǽ���,b.Addtime �Ǽ�ʱ��
              from V_SOL_INF_DELIVERY a,v_newborn b,v_sol_rs_birth c,sol_inf_puerpera d,
              (select mid,sex,babyno,row_number() over(partition by mid,sex order by babyno) ˳�� from SOL_INF_NEWBORNS t  ) t
              where a.Mid=b.Mid and a.mid=d.mid and a.mid=t.mid and b.Babyno=t.babyno
              and b.Mid=c.mid(+) and a.mid=mid_In
                ) where ���=babyno_In;
                
      select count(*) into n_count from his_������������¼;
      dbms_output.put_line(n_count);
      
      insert into ������������¼@to_his value select * from his_������������¼ ;
      Zl_�����Զ����_Update@To_His(n_����id, n_��ҳid); 
      b_Message.Zlhis_Patient_011@To_His(n_����id, n_��ҳid, babyno_In);
      delete from his_������������¼;
    else  ----�޸�
      insert into   his_������������¼ 
      (����id,��ҳid,���,Ӥ������,Ӥ���Ա�,�������,���䷽ʽ,̥��״��,����ʱ��,��,����,Ѫ��,��ע˵��,����ʱ��,�Ǽ�ʱ��,�Ǽ���)
      select ����id,סԺ����,���,Ӥ������,Ӥ���Ա�,�������,�����ʽ,̥��״��,����ʱ��,��,����,Ѫ��,��ע˵��,����ʱ��,�Ǽ�ʱ��,�Ǽ���
       from   ( select  d.pid ����id,d.tid סԺ����,b.Babyno ���,d.name||decode(b.Sex,'��','֮��','֮Ů')||t.˳�� as Ӥ������,
              b.Sex as Ӥ���Ա�,c.�Ѵ� as �������,a.�����ʽ,b.̥��״��,b.��,b.����,b.Ѫ��,to_date(a.̥�����ʱ��,'yyyy-mm-dd hh24:mi:ss') as ����ʱ��,
              b.����ʱ��,'' as ��ע˵��,b.Recorder �Ǽ���,b.Addtime �Ǽ�ʱ��
              from V_SOL_INF_DELIVERY a,v_newborn b,v_sol_rs_birth c,sol_inf_puerpera d,
              (select mid,sex,babyno,row_number() over(partition by mid,sex order by babyno) ˳�� from SOL_INF_NEWBORNS t  ) t
              where a.Mid=b.Mid and a.mid=d.mid and a.mid=t.mid and b.Babyno=t.babyno
              and b.Mid=c.mid(+) and a.mid=mid_In
                ) where ���=babyno_In;
        delete from ������������¼@to_his where ����id=n_����id and ��ҳid=n_��ҳid and ���=babyno_In;
        insert into ������������¼@to_his value select * from his_������������¼ ;
        Zl_�����Զ����_Update@To_His(n_����id, n_��ҳid); 
      b_Message.Zlhis_Patient_011@To_His(n_����id, n_��ҳid, babyno_In);
   /* update  ������������¼@to_his set  
   (���,Ӥ������,Ӥ���Ա�,�������,���䷽ʽ,̥��״��,����ʱ��,��,����,Ѫ��,��ע˵��,����ʱ��,�Ǽ�ʱ��,�Ǽ���)=
   (select ���,Ӥ������,Ӥ���Ա�,�������,���䷽ʽ,̥��״��,����ʱ��,��,����,Ѫ��,��ע˵��,����ʱ��,�Ǽ�ʱ��,�Ǽ���
    from   his_������������¼) ;*/
     delete from his_������������¼;          
      end if;
 
      
     --�������Ǽ�ɾ��
   elsif state_in=3 then
     delete from ������������¼@to_his where ����id=n_����id and ��ҳid=n_��ҳid and ���=babyno_In;
     Zl_�����Զ����_Update@To_His(n_����id,n_��ҳid); 
 
     b_Message.ZLHIS_PATIENT_013@To_His(n_����id,n_��ҳid,babyno_In); 

  End If;

 
   
End his_�����������Ǽ�_revise;
/

