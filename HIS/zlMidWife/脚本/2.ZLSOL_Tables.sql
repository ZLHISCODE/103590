--��zlsol�û�����
create table SOL_USERLIST(
USER_CODE  VARCHAR(20),
USER_NAME  VARCHAR(50)
);
comment on table SOL_USERLIST IS 'ϵͳ�û���Ϣ';
alter table SOL_USERLIST add constraint SOL_USERLIST_PK primary key(USER_CODE);

create table SOL_STD_FETALPOSITION
(
  code        VARCHAR2(10),
  name        VARCHAR2(50),
  description VARCHAR2(100)
);
comment on table SOL_STD_FETALPOSITION IS '̥��λ';
alter table SOL_STD_FETALPOSITION add constraint SOL_STD_FETALPOSITION_PK primary key(code);

create table SOL_INF_PUERPERA(
 MID NUMBER(18)  
 generated as identity( start with 1 nocycle noorder),
 PID NUMBER(18)  ,
 TID NUMBER(5) ,
 NAME   VARCHAR(50),
 OLD VARCHAR(20),
 BEDNO  VARCHAR(10),
 PNO NUMBER(18),
 DIAGNOSIS  VARCHAR(100),
 STATUS   NUMBER(1),  --״̬�����Ϊ0���뷿���Ӻ�Ϊ1����������Ϊ2����ԺΪ3
 OUTROOMTIME  DATE,
 OUTTIME  DATE��
 EXPECTANT      number(1)��  --1Ϊ�������
 CHECKINROOM    number(1)��  --�뷿
 BIRTH          number(1)��  --�ٲ�     
 DRUGLABOR      number(1)��  --ҩ������   
 DELIVERY       number(1)��  --����   
 NEWBORNS       number(1)��  --������    
 POSTPARTUM     number(1)��  --���� 
 CHECKOUTROOM   number(1),    --����
 equipment      NUMBER(1)     --��е
);
comment on table SOL_INF_PUERPERA IS '������Ϣ';
alter table SOL_INF_PUERPERA add constraint SOL_INF_Puerpera_PK primary key(MID);
create index SOL_INF_PUERPERA_IX_STATUS on SOL_INF_PUERPERA(STATUS);
create index SOL_INF_PUERPERA_IX_OUTTIME on SOL_INF_PUERPERA(OUTTIME);
create index SOL_INF_PUERPERA_IX_OUTROOMTIME on SOL_INF_PUERPERA(OUTROOMTIME);

create table SOL_INF_EQUIPMENT
(
  mid        NUMBER(18),
  content    CLOB,
  recordtime DATE,
  deliver    VARCHAR2(50),
  inspector  VARCHAR2(50)
);
comment on table SOL_INF_EQUIPMENT IS '��е����¼';
alter table SOL_INF_EQUIPMENT add constraint SOL_INF_EQUIPMENT_PK primary key(MID);
alter table SOL_INF_EQUIPMENT add constraint SOL_INF_EQUIPMENT_FK_MID foreign key(MID)  references SOL_INF_PUERPERA(MID);


create table SOL_RS_EXPECTANT(
 COURSEID NUMBER(18) generated as identity( start with 1 nocycle noorder),
 MID NUMBER(18)  ,
 CONTENT  CLOB CHECK(CONTENT IS JSON)
);
comment on table SOL_RS_EXPECTANT IS '������¼';
alter table SOL_RS_EXPECTANT add constraint SOL_RS_Expectant_PK primary key(COURSEID);
alter table SOL_RS_EXPECTANT add constraint SOL_RS_Expectant_FK_MID foreign key(MID)  references SOL_INF_PUERPERA(MID);


create table SOL_INF_CHECKINROOM(
 MID NUMBER(18)  ,
 Content  CLOB CHECK(Content IS JSON),
 RECORDER VARCHAR(50),
 ADDTIME  DATE
);
comment on table SOL_INF_CHECKINROOM IS '�뷿��Ϣ';
alter table SOL_INF_CHECKINROOM add constraint SOL_INF_CheckInRoom_PK primary key(MID);
alter table SOL_INF_CHECKINROOM add constraint SOL_INF_CheckInRoom_FK_MID foreign key(MID) references SOL_INF_PUERPERA(MID);



--�ٲ���¼
create table SOL_RS_BIRTH(
 MID NUMBER(18)  ,
 Content  CLOB CHECK(Content IS JSON)
);
comment on table SOL_RS_BIRTH IS '��ǰ�����Ϣ';
alter table SOL_RS_BIRTH add constraint SOL_RS_Birth_PK primary key(MID);
alter table SOL_RS_BIRTH add constraint SOL_RS_Birth_FK_MID foreign key(MID) references SOL_INF_PUERPERA(MID);



create table SOL_RS_BIRTH_COURSE(
 COURSEID NUMBER(18) generated as identity( start with 1 nocycle noorder),
 MID NUMBER(18) ,
 CONTENT   CLOB CHECK(CONTENT IS JSON)
);
comment on table SOL_RS_BIRTH_COURSE IS '���̾���';
alter table SOL_RS_BIRTH_COURSE  add constraint SOL_RS_BIRTH_COURSE_PK primary key(COURSEID);
alter table SOL_RS_BIRTH_COURSE  add constraint SOL_RS_BIRTH_COURSE_FK_MID foreign key(MID) references SOL_INF_PUERPERA(MID);


create table SOL_RS_DRUGLABOR(
 MID NUMBER(18)  ,
 ����   Date,
 ����ָ�� Varchar2(50),
 �������� Varchar2(50)
);
comment on table SOL_RS_DRUGLABOR IS 'ҩ��������Ϣ';
alter table SOL_RS_DRUGLABOR  add constraint SOL_RS_DrugLabor_PK primary key(MID);
alter table SOL_RS_DRUGLABOR  add constraint SOL_RS_DrugLabor_FK_MID foreign key(MID) references SOL_INF_PUERPERA(MID);

create table SOL_RS_DRUGLABOR_LIST(
 COURSEID NUMBER(18) generated as identity( start with 1 nocycle noorder),
 MID NUMBER(18) ,
 CONTENT  CLOB CHECK(CONTENT IS JSON)
);
comment on table SOL_RS_DRUGLABOR_LIST IS 'ҩ��������¼';
alter table SOL_RS_DRUGLABOR_LIST  add constraint SOL_RS_DRUGLABOR_LIST_PK primary key(COURSEID);
alter table SOL_RS_DRUGLABOR_LIST  add constraint SOL_RS_DRUGLABOR_LIST_FK_MID foreign key(MID) references SOL_INF_PUERPERA(MID);



create table SOL_INF_DELIVERY(
 MID NUMBER(18)  ,
 DELIVERYINF   CLOB CHECK(DELIVERYINF IS JSON),
 NEWBORNDETAIL CLOB CHECK(NEWBORNDETAIL IS JSON),
 NEWBORNSCORE  CLOB CHECK(NEWBORNSCORE IS JSON),
 OTHERINF CLOB CHECK(OTHERINF IS JSON)
);
comment on table SOL_INF_DELIVERY IS '������Ϣ';
alter table SOL_INF_DELIVERY add constraint SOL_INF_Delivery_PK primary key(MID);
alter table SOL_INF_DELIVERY add constraint SOL_INF_Delivery_FK_MID foreign key(MID) references SOL_INF_PUERPERA(MID);


create table SOL_INF_NEWBORNS(
 BID NUMBER(18)  generated as identity( start with 1 nocycle noorder),
 MID NUMBER(18),
 BABYNO   NUMBER(5) ,
 SEX VARCHAR2(10)   ,
 NEWBORNINF  CLOB CHECK(NEWBORNINF IS JSON),
 NEWBORNSCORE  CLOB CHECK(NEWBORNSCORE IS JSON),
 OTHERINF CLOB CHECK(OTHERINF IS JSON),
 RECORDER VARCHAR2(50),
 ADDTIME  DATE
);
comment on table SOL_INF_NEWBORNS IS '��������Ϣ';
alter table SOL_INF_NEWBORNS add constraint SOL_INF_Newborns_PK primary key(BID);
Alter Table SOL_INF_NEWBORNS Add Constraint SOL_INF_NEWBORNS_UQ Unique(MID,BABYNO);
alter table SOL_INF_NEWBORNS add constraint SOL_INF_Newborns_FK_MID foreign key(MID) references SOL_INF_PUERPERA(MID);



create table SOL_RS_POSTPARTUM(
 MID NUMBER(18)  ,
 CONTENT  CLOB CHECK(CONTENT IS JSON)
);
comment on table SOL_RS_POSTPARTUM IS '����۲���Ϣ';
alter table SOL_RS_POSTPARTUM add constraint SOL_RS_Postpartum_PK primary key(MID);
alter table SOL_RS_POSTPARTUM add constraint SOL_RS_Postpartum_FK_MID foreign key(MID) references SOL_INF_PUERPERA(MID);

create table SOL_RS_POSTPARTUM_LIST(
 COURSEID NUMBER(18) generated as identity( start with 1 nocycle noorder),
 MID NUMBER(18) ,
 CONTENT  CLOB CHECK(CONTENT IS JSON)
);
comment on table SOL_RS_POSTPARTUM_LIST IS '����۲��¼';
alter table SOL_RS_POSTPARTUM_LIST  add CONSTRAINT SOL_RS_POSTPARTUM_LIST_PK primary key(COURSEID);
alter table SOL_RS_POSTPARTUM_LIST  add constraint SOL_RS_POSTPARTUM_LIST_FK_MID foreign key(MID) references SOL_INF_PUERPERA(MID);



create table SOL_INF_CHECKOUTROOM(
 MID NUMBER(18)  ,
 Content  CLOB CHECK(Content IS JSON),
 RECORDER VARCHAR(50),
 ADDTIME  DATE
);
comment on table SOL_INF_CHECKOUTROOM IS '������Ϣ';
alter table SOL_INF_CHECKOUTROOM add constraint SOL_INF_CheckOutRoom_PK primary key(MID);
alter table SOL_INF_CHECKOUTROOM add constraint SOL_INF_CheckOutRoom_FK_MID foreign key(MID) references SOL_INF_PUERPERA(MID);



