--药房自动化接口虚拟模块
Insert Into zlPrograms
  (序号, 标题, 说明, 系统, 部件)
  Select 1348, '药房自动化接口', 'HIS与药房自动配、发药系统接口', &n_Syttem, 'zlDrugPacker'
  From Dual
  Where Not Exists (Select 1 From zlPrograms Where 序号 = 1348 And 标题 = '药房自动化接口');

--药房自动化接口虚拟模块
Insert Into zlProgFuncs
  (系统, 序号, 功能, 排列, 说明)
  Select &n_System, 1348, '基本', Null, Null
  From Dual
  Where Not Exists (Select 1 From zlProgFuncs Where 系统 = &n_System And 序号 = 1348 And 功能 = '基本');

--药房自动化接口虚拟模块
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) 
  select &n_System,1348,'基本',User,'部门表','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='部门表') union all
  select &n_System,1348,'基本',User,'部门人员','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='部门人员') union all
  select &n_System,1348,'基本',User,'部门性质说明','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='部门性质说明') union all
  select &n_System,1348,'基本',User,'人员表','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='人员表') union all
  select &n_System,1348,'基本',User,'上机人员表','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='上机人员表') union all
  select &n_System,1348,'基本',User,'药品剂型','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='药品剂型') union all
  select &n_System,1348,'基本',User,'药房设备连接','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='药房设备连接') union all
  select &n_System,1348,'基本',User,'药房设备参数','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='药房设备参数') union all
  select &n_System,1348,'基本',User,'药房注册设备','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='药房注册设备') union all
  select &n_System,1348,'基本',User,'ZL_药房设备连接_INSERT','EXECUTE' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='ZL_药房设备连接_INSERT') union all
  select &n_System,1348,'基本',User,'ZL_药房设备连接_UPDATE','EXECUTE' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='ZL_药房设备连接_UPDATE') union all
  select &n_System,1348,'基本',User,'ZL_药房设备连接_DELETE','EXECUTE' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='ZL_药房设备连接_DELETE') union all
  select &n_System,1348,'基本',User,'ZL_药房注册设备_INSERT','EXECUTE' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='ZL_药房注册设备_INSERT') union all
  select &n_System,1348,'基本',User,'ZL_药房注册设备_UPDATE','EXECUTE' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='ZL_药房注册设备_UPDATE') union all
  select &n_System,1348,'基本',User,'ZL_药房注册设备_DELETE','EXECUTE' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='ZL_药房注册设备_DELETE') union all
  select &n_System,1348,'基本',User,'ZL_药房注册设备_SWITCH','EXECUTE' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='ZL_药房注册设备_SWITCH') union all
  select &n_System,1348,'基本',User,'ZL_药房注册设备_SETTING','EXECUTE' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='ZL_药房注册设备_SETTING') ;

--药房自动化接口虚拟模块（1348）
Insert Into zlParameters(ID,系统,模块,私有,本机,授权,固定,参数号,参数名,参数值,缺省值,参数说明)
Select Rownum+B.ID,A.* From (
  Select 系统,模块,私有,本机,授权,固定,参数号,参数名,参数值,缺省值,参数说明 From zlParameters Where ID=0 Union All
  Select &n_Syttem,1348,0,0,0,0,1,'服务对象',NULL,'1','药房设备可执行的医嘱或处方。1-门诊；2-住院' From Dual Union All
  Select &n_Syttem,1348,0,0,0,0,2,'配药对应业务',NULL,NULL,'1-门诊收费；2-处方发药配药功能；3-处方发药发药功能' From Dual Union All
  Select &n_Syttem,1348,0,0,0,0,3,'发送对应业务',NULL,NULL,'1-启用药品处方发药' From Dual Union All
  Select &n_Syttem,1348,0,0,0,0,4,'单据类型',NULL,NULL,'Null表示未选择；全部表示所有三种单据；1-长嘱；2-临嘱;3-记帐单”' From Dual Union All
  Select &n_Syttem,1348,0,0,0,0,5,'药品剂型',NULL,NULL,'Null表示未选择；全部表示所有药品剂型；如果需要指定某些剂型，格式：“粉型|片剂|…' From Dual
  ) A,(Select Nvl(Max(ID),0) AS ID From zlParameters) B;



create table 药房设备连接  (
   ID          NUMBER(10)      not null,
   名称        VARCHAR2(20)    not null,
   连接类型    NUMBER(2)       not null,
   连接内容    VARCHAR2(200)   not null
)
  tablespace ZL9MEDLST;

Alter Table 药房设备连接 Add Constraint 药房设备连接_PK Primary Key (ID) Using Index Tablespace ZL9INDEXHIS;
Alter Table 药房设备连接 Add Constraint 药房设备连接_UQ_名称 Unique (名称) Using Index Tablespace ZL9INDEXHIS;

create sequence 药房设备连接_ID start with 1;




create table 药房注册设备  (
   ID               NUMBER(18)            not null,
   编码             VARCHAR2(20)          not null,
   名称             VARCHAR2(20)          not null,
   型号             VARCHAR2(20),
   制造商           VARCHAR2(100),
   部门ID           NUMBER(18)            not null,
   连接ID           NUMBER(10)            not null,
   启用             NUMBER(1)
)
  tablespace ZL9MEDLST;
 
Alter Table 药房注册设备 Add Constraint 药房注册设备_PK Primary Key (ID) Using Index Tablespace ZL9INDEXHIS;
Alter Table 药房注册设备 Add Constraint 药房注册设备_UQ_编码 Unique (编码) Using Index Tablespace ZL9INDEXHIS;
Alter Table 药房注册设备 Add constraint 药房注册设备_UQ_部门ID unique (部门ID, 名称, 型号) using index  tablespace ZL9INDEXHIS;

alter table 药房注册设备 add constraint 药房注册设备_FK_连接ID foreign key (连接ID) references 药房设备连接 (ID);
alter table 药房注册设备 add constraint 药房注册设备_FK_部门ID foreign key (部门ID) references 部门表 (ID);

create sequence 药房注册设备_ID start with 1;



create table 药房设备参数  (
   参数ID             NUMBER(18)                      not null,
   设备ID             NUMBER(18)                      not null,
   参数值             VARCHAR2(4000)
);

alter table 药房设备参数 add constraint 药房设备参数_PK primary key (参数ID, 设备ID) Using Index Tablespace ZL9INDEXHIS;

alter table 药房设备参数 add constraint 药房设备参数_FK_设备ID foreign key (设备ID) references 药房注册设备 (ID);






CREATE OR REPLACE Procedure Zl_药房设备连接_Insert
(
  名称_In In 药房设备连接.名称%Type,
  类型_In In 药房设备连接.连接类型%Type,
  内容_In In 药房设备连接.连接内容%Type
) Is

Begin

  Insert Into 药房设备连接 
    (ID, 名称, 连接类型, 连接内容) 
    Values 
    (药房设备连接_Id.Nextval, 名称_In, 类型_In, 内容_In);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药房设备连接_Insert;
/

Create Or Replace Procedure Zl_药房设备连接_Delete(Id_In In 药房设备连接.Id%Type) Is
Begin

  Delete 药房设备连接 Where ID = Id_In;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药房设备连接_Delete;
/

CREATE OR REPLACE Procedure Zl_药房设备连接_Update
(
  Id_In   In 药房设备连接.Id%Type,
  名称_In In 药房设备连接.名称%Type,
  类型_In In 药房设备连接.连接类型%Type,
  内容_In In 药房设备连接.连接内容%Type
) Is
Begin

  Update 药房设备连接 Set 名称 = 名称_In, 连接类型 = 类型_In, 连接内容 = 内容_In Where ID = Id_In;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药房设备连接_Update;
/

Create Or Replace Procedure Zl_药房注册设备_Insert
(
  编码_In         In 药房注册设备.编码%Type,
  名称_In         In 药房注册设备.名称%Type,
  型号_In         In 药房注册设备.型号%Type := Null,
  制造商_In       In 药房注册设备.制造商%Type := Null,
  连接id_In       In 药房注册设备.连接id%Type,
  部门id_In       In 药房注册设备.部门id%Type,
  启用_In         In 药房注册设备.启用%Type := Null,
  服务对象_In     In 药房设备参数.参数值%Type,
  配药对应业务_In In 药房设备参数.参数值%Type := Null,
  发送对应业务_In In 药房设备参数.参数值%Type := Null
) Is

  n_设备id Number;

Begin

  Select 药房注册设备_Id.Nextval Into n_设备id From Dual;

  Insert Into 药房注册设备
    (ID, 编码, 名称, 型号, 制造商, 部门id, 连接id, 启用)
  Values
    (n_设备id, 编码_In, 名称_In, 型号_In, 制造商_In, 部门id_In, 连接id_In, 启用_In);

  Insert Into 药房设备参数
    (参数id, 设备id, 参数值)
    Select ID, n_设备id, 服务对象_In
    From Zlparameters
    Where 系统 = &n_Syttem And 模块 = 1348 And 参数号 = 1
    Union All
    Select ID, n_设备id, 配药对应业务_In
    From Zlparameters
    Where 系统 = &n_Syttem And 模块 = 1348 And 参数号 = 2 And 配药对应业务_In Is Not Null
    Union All
    Select ID, n_设备id, 发送对应业务_In
    From Zlparameters
    Where 系统 = &n_Syttem And 模块 = 1348 And 参数号 = 3 And 发送对应业务_In Is Not Null;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药房注册设备_Insert;
/

Create Or Replace Procedure Zl_药房注册设备_Delete(设备id_In In 药房注册设备.Id%Type) Is
Begin

  Delete 药房设备参数 Where 设备id = 设备id_In;

  Delete 药房注册设备 Where ID = 设备id_In;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药房注册设备_Delete;
/

Create Or Replace Procedure Zl_药房注册设备_Switch
(
  设备id_In In 药房注册设备.Id%Type,
  开关_In   In 药房注册设备.启用%Type := Null
) Is
Begin

  Update 药房注册设备 Set 启用 = 开关_In Where ID = 设备id_In;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药房注册设备_Switch;
/

Create Or Replace Procedure Zl_药房注册设备_Setting
(
  设备id_In   In 药房注册设备.Id%Type,
  单据类型_In In 药房设备参数.参数值%Type := Null,
  药品剂型_In In 药房设备参数.参数值%Type := Null
) Is

  n_参数id Number;

Begin

  --单据类型
  Select ID Into n_参数id From Zlparameters Where 系统 = 100 And 模块 = 1348 And 参数号 = 4;
  If Nvl(n_参数id, 0) > 0 Then
--    If 单据类型_In Is Null Then
--      Delete 药房设备参数 Where 参数id = n_参数id And 设备id = 设备id_In;
--    Else
      Update 药房设备参数 Set 参数值 = 单据类型_In Where 参数id = n_参数id And 设备id = 设备id_In;
      If Sql%NotFound Then
        Insert Into 药房设备参数 (参数id, 设备id, 参数值) Values (n_参数id, 设备id_In, 单据类型_In);
      End If;
--    End If;
  End If;

  --药品剂型
  Select ID Into n_参数id From Zlparameters Where 系统 = 100 And 模块 = 1348 And 参数号 = 5;
  If Nvl(n_参数id, 0) > 0 Then
--    If 药品剂型_In Is Null Then
--      Delete 药房设备参数 Where 参数id = n_参数id And 设备id = 设备id_In;
--    Else
      Update 药房设备参数 Set 参数值 = 药品剂型_In Where 参数id = n_参数id And 设备id = 设备id_In;
      If Sql%NotFound Then
        Insert Into 药房设备参数 (参数id, 设备id, 参数值) Values (n_参数id, 设备id_In, 药品剂型_In);
      End If;
--    End If;
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药房注册设备_Setting;
/

Create Or Replace Procedure Zl_药房注册设备_Update
(
  设备id_In       In 药房注册设备.Id%Type,
  编码_In         In 药房注册设备.编码%Type,
  名称_In         In 药房注册设备.名称%Type,
  型号_In         In 药房注册设备.型号%Type := Null,
  制造商_In       In 药房注册设备.制造商%Type := Null,
  连接id_In       In 药房注册设备.连接id%Type,
  部门id_In       In 药房注册设备.部门id%Type,
  启用_In         In 药房注册设备.启用%Type := Null,
  服务对象_In     In 药房设备参数.参数值%Type,
  配药对应业务_In In 药房设备参数.参数值%Type := Null,
  发送对应业务_In In 药房设备参数.参数值%Type := Null
) Is

  n_参数id Number;

Begin

  Update 药房注册设备
  Set 编码 = 编码_In, 名称 = 名称_In, 型号 = 型号_In, 制造商 = 制造商_In, 部门id = 部门id_In, 连接id = 连接id_In, 启用 = 启用_In
  Where ID = 设备id_In;

  --服务对象
  Begin
    Select ID Into n_参数id From Zlparameters Where 系统 = 100 And 模块 = 1348 And 参数号 = 1;
  Exception
    When Others Then
      n_参数id := Null;
  End;
  If n_参数id Is Not Null Then
    Update 药房设备参数 Set 参数值 = 服务对象_In Where 参数id = n_参数id And 设备id = 设备id_In;
    If Sql%NotFound Then
      Insert Into 药房设备参数
        (参数id, 设备id, 参数值)
        Select ID, 设备id_In, 服务对象_In From Zlparameters Where 系统 = 100 And 模块 = 1348 And 参数号 = 1;
    End If;
  End If;

  --配药业务
  Begin
    Select ID Into n_参数id From Zlparameters Where 系统 = 100 And 模块 = 1348 And 参数号 = 2;
  Exception
    When Others Then
      n_参数id := Null;
  End;
  If n_参数id Is Not Null Then
    Update 药房设备参数 Set 参数值 = 配药对应业务_In Where 参数id = n_参数id And 设备id = 设备id_In;
    If Sql%NotFound Then
      Insert Into 药房设备参数
        (参数id, 设备id, 参数值)
        Select ID, 设备id_In, 配药对应业务_In From Zlparameters Where 系统 = 100 And 模块 = 1348 And 参数号 = 2;
    End If;
  End If;

  --发送业务
  Begin
    Select ID Into n_参数id From Zlparameters Where 系统 = 100 And 模块 = 1348 And 参数号 = 3;
  Exception
    When Others Then
      n_参数id := Null;
  End;
  If n_参数id Is Not Null Then
    Update 药房设备参数 Set 参数值 = 发送对应业务_In Where 参数id = n_参数id And 设备id = 设备id_In;
    If Sql%NotFound Then
      Insert Into 药房设备参数
        (参数id, 设备id, 参数值)
        Select ID, 设备id_In, 发送对应业务_In From Zlparameters Where 系统 = 100 And 模块 = 1348 And 参数号 = 3;
    End If;
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药房注册设备_Update;
/
