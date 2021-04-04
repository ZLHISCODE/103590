--以zlsol用户运行
create table SOL_USERLIST(
USER_CODE  VARCHAR(20),
USER_NAME  VARCHAR(50)
);
comment on table SOL_USERLIST IS '系统用户信息';
alter table SOL_USERLIST add constraint SOL_USERLIST_PK primary key(USER_CODE);

create table SOL_STD_FETALPOSITION
(
  code        VARCHAR2(10),
  name        VARCHAR2(50),
  description VARCHAR2(100)
);
comment on table SOL_STD_FETALPOSITION IS '胎方位';
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
 STATUS   NUMBER(1),  --状态：入科为0，入房交接后为1，出房交接为2，出院为3
 OUTROOMTIME  DATE,
 OUTTIME  DATE，
 EXPECTANT      number(1)，  --1为已填，待产
 CHECKINROOM    number(1)，  --入房
 BIRTH          number(1)，  --临产     
 DRUGLABOR      number(1)，  --药物引产   
 DELIVERY       number(1)，  --分娩   
 NEWBORNS       number(1)，  --新生儿    
 POSTPARTUM     number(1)，  --产后 
 CHECKOUTROOM   number(1),    --出房
 equipment      NUMBER(1)     --器械
);
comment on table SOL_INF_PUERPERA IS '产妇信息';
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
comment on table SOL_INF_EQUIPMENT IS '器械清点记录';
alter table SOL_INF_EQUIPMENT add constraint SOL_INF_EQUIPMENT_PK primary key(MID);
alter table SOL_INF_EQUIPMENT add constraint SOL_INF_EQUIPMENT_FK_MID foreign key(MID)  references SOL_INF_PUERPERA(MID);


create table SOL_RS_EXPECTANT(
 COURSEID NUMBER(18) generated as identity( start with 1 nocycle noorder),
 MID NUMBER(18)  ,
 CONTENT  CLOB CHECK(CONTENT IS JSON)
);
comment on table SOL_RS_EXPECTANT IS '待产记录';
alter table SOL_RS_EXPECTANT add constraint SOL_RS_Expectant_PK primary key(COURSEID);
alter table SOL_RS_EXPECTANT add constraint SOL_RS_Expectant_FK_MID foreign key(MID)  references SOL_INF_PUERPERA(MID);


create table SOL_INF_CHECKINROOM(
 MID NUMBER(18)  ,
 Content  CLOB CHECK(Content IS JSON),
 RECORDER VARCHAR(50),
 ADDTIME  DATE
);
comment on table SOL_INF_CHECKINROOM IS '入房信息';
alter table SOL_INF_CHECKINROOM add constraint SOL_INF_CheckInRoom_PK primary key(MID);
alter table SOL_INF_CHECKINROOM add constraint SOL_INF_CheckInRoom_FK_MID foreign key(MID) references SOL_INF_PUERPERA(MID);



--临产记录
create table SOL_RS_BIRTH(
 MID NUMBER(18)  ,
 Content  CLOB CHECK(Content IS JSON)
);
comment on table SOL_RS_BIRTH IS '产前检查信息';
alter table SOL_RS_BIRTH add constraint SOL_RS_Birth_PK primary key(MID);
alter table SOL_RS_BIRTH add constraint SOL_RS_Birth_FK_MID foreign key(MID) references SOL_INF_PUERPERA(MID);



create table SOL_RS_BIRTH_COURSE(
 COURSEID NUMBER(18) generated as identity( start with 1 nocycle noorder),
 MID NUMBER(18) ,
 CONTENT   CLOB CHECK(CONTENT IS JSON)
);
comment on table SOL_RS_BIRTH_COURSE IS '产程经过';
alter table SOL_RS_BIRTH_COURSE  add constraint SOL_RS_BIRTH_COURSE_PK primary key(COURSEID);
alter table SOL_RS_BIRTH_COURSE  add constraint SOL_RS_BIRTH_COURSE_FK_MID foreign key(MID) references SOL_INF_PUERPERA(MID);


create table SOL_RS_DRUGLABOR(
 MID NUMBER(18)  ,
 日期   Date,
 引产指征 Varchar2(50),
 引产方法 Varchar2(50)
);
comment on table SOL_RS_DRUGLABOR IS '药物引产信息';
alter table SOL_RS_DRUGLABOR  add constraint SOL_RS_DrugLabor_PK primary key(MID);
alter table SOL_RS_DRUGLABOR  add constraint SOL_RS_DrugLabor_FK_MID foreign key(MID) references SOL_INF_PUERPERA(MID);

create table SOL_RS_DRUGLABOR_LIST(
 COURSEID NUMBER(18) generated as identity( start with 1 nocycle noorder),
 MID NUMBER(18) ,
 CONTENT  CLOB CHECK(CONTENT IS JSON)
);
comment on table SOL_RS_DRUGLABOR_LIST IS '药物引产记录';
alter table SOL_RS_DRUGLABOR_LIST  add constraint SOL_RS_DRUGLABOR_LIST_PK primary key(COURSEID);
alter table SOL_RS_DRUGLABOR_LIST  add constraint SOL_RS_DRUGLABOR_LIST_FK_MID foreign key(MID) references SOL_INF_PUERPERA(MID);



create table SOL_INF_DELIVERY(
 MID NUMBER(18)  ,
 DELIVERYINF   CLOB CHECK(DELIVERYINF IS JSON),
 NEWBORNDETAIL CLOB CHECK(NEWBORNDETAIL IS JSON),
 NEWBORNSCORE  CLOB CHECK(NEWBORNSCORE IS JSON),
 OTHERINF CLOB CHECK(OTHERINF IS JSON)
);
comment on table SOL_INF_DELIVERY IS '分娩信息';
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
comment on table SOL_INF_NEWBORNS IS '新生儿信息';
alter table SOL_INF_NEWBORNS add constraint SOL_INF_Newborns_PK primary key(BID);
Alter Table SOL_INF_NEWBORNS Add Constraint SOL_INF_NEWBORNS_UQ Unique(MID,BABYNO);
alter table SOL_INF_NEWBORNS add constraint SOL_INF_Newborns_FK_MID foreign key(MID) references SOL_INF_PUERPERA(MID);



create table SOL_RS_POSTPARTUM(
 MID NUMBER(18)  ,
 CONTENT  CLOB CHECK(CONTENT IS JSON)
);
comment on table SOL_RS_POSTPARTUM IS '产后观察信息';
alter table SOL_RS_POSTPARTUM add constraint SOL_RS_Postpartum_PK primary key(MID);
alter table SOL_RS_POSTPARTUM add constraint SOL_RS_Postpartum_FK_MID foreign key(MID) references SOL_INF_PUERPERA(MID);

create table SOL_RS_POSTPARTUM_LIST(
 COURSEID NUMBER(18) generated as identity( start with 1 nocycle noorder),
 MID NUMBER(18) ,
 CONTENT  CLOB CHECK(CONTENT IS JSON)
);
comment on table SOL_RS_POSTPARTUM_LIST IS '产后观察记录';
alter table SOL_RS_POSTPARTUM_LIST  add CONSTRAINT SOL_RS_POSTPARTUM_LIST_PK primary key(COURSEID);
alter table SOL_RS_POSTPARTUM_LIST  add constraint SOL_RS_POSTPARTUM_LIST_FK_MID foreign key(MID) references SOL_INF_PUERPERA(MID);



create table SOL_INF_CHECKOUTROOM(
 MID NUMBER(18)  ,
 Content  CLOB CHECK(Content IS JSON),
 RECORDER VARCHAR(50),
 ADDTIME  DATE
);
comment on table SOL_INF_CHECKOUTROOM IS '出房信息';
alter table SOL_INF_CHECKOUTROOM add constraint SOL_INF_CheckOutRoom_PK primary key(MID);
alter table SOL_INF_CHECKOUTROOM add constraint SOL_INF_CheckOutRoom_FK_MID foreign key(MID) references SOL_INF_PUERPERA(MID);



