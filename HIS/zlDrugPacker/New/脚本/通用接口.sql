--药房自动化接口虚拟模块
Insert Into zlPrograms
  (序号, 标题, 说明, 系统, 部件)
  Select 1348, '药房自动化接口', 'HIS与药房自动配、发药系统接口', &n_System, 'zlDrugPacker'
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
  select &n_System,1348,'基本',User,'药房发药设备','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='药房发药设备') union all
  select &n_System,1348,'基本',User,'自动发药参数','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='自动发药参数') union all
  select &n_System,1348,'基本',User,'药房设备参数','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='药房设备参数') union all
  select &n_System,1348,'基本',User,'病人挂号记录','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='病人挂号记录') union all
  select &n_System,1348,'基本',User,'病人信息','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='病人信息') union all
  select &n_System,1348,'基本',User,'病人医嘱发送','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='病人医嘱发送') union all
  select &n_System,1348,'基本',User,'病人医嘱记录','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='病人医嘱记录') union all
  select &n_System,1348,'基本',User,'病人诊断记录','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='病人诊断记录') union all
  select &n_System,1348,'基本',User,'发药窗口','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='发药窗口') union all
  select &n_System,1348,'基本',User,'供应商','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='供应商') union all
  select &n_System,1348,'基本',User,'门诊费用记录','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='门诊费用记录') union all
  select &n_System,1348,'基本',User,'收费价目','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='收费价目') union all
  select &n_System,1348,'基本',User,'收费项目别名','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='收费项目别名') union all
  select &n_System,1348,'基本',User,'收费项目目录','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='收费项目目录') union all
  select &n_System,1348,'基本',User,'未发药品记录','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='未发药品记录') union all
  select &n_System,1348,'基本',User,'药品储备限额','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='药品储备限额') union all
  select &n_System,1348,'基本',User,'药品规格','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='药品规格') union all
  select &n_System,1348,'基本',User,'药品库存','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='药品库存') union all
  select &n_System,1348,'基本',User,'药品生产商','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='药品生产商') union all
  select &n_System,1348,'基本',User,'药品收发记录','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='药品收发记录') union all
  select &n_System,1348,'基本',User,'药品特性','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='药品特性') union all
  select &n_System,1348,'基本',User,'诊疗分类目录','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='诊疗分类目录') union all
  select &n_System,1348,'基本',User,'诊疗项目目录','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='诊疗项目目录') union all
  select &n_System,1348,'基本',User,'住院费用记录','SELECT' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='住院费用记录') union all
  select &n_System,1348,'基本',User,'Zl_药房发药设备_Insert','EXECUTE' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='Zl_药房发药设备_Insert') union all
  select &n_System,1348,'基本',User,'Zl_药房发药设备_Update','EXECUTE' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='Zl_药房发药设备_Update') union all
  select &n_System,1348,'基本',User,'Zl_药房发药设备_Delete','EXECUTE' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='Zl_药房发药设备_Delete') union all
  select &n_System,1348,'基本',User,'Zl_药房发药设备_Switch','EXECUTE' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='Zl_药房发药设备_Switch') union all
  select &n_System,1348,'基本',User,'Zl_药房设备参数_Update','EXECUTE' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='Zl_药房设备参数_Update') union all
  select &n_System,1348,'基本',User,'Zl_未发药品记录_分配发药窗口','EXECUTE' from dual where not exists(select 1 from zlprogprivs where 系统=&n_System and 序号=1348 and 功能='基本' and 对象='Zl_未发药品记录_分配发药窗口');
  
--增加系统参数
Insert Into zlParameters(ID,系统,模块,参数号,参数名,参数值,缺省值,参数说明)
Select zlParameters_ID.Nextval, &n_System,-Null,222, '药房自动化发药接口','0','0','是否启用药房自动化发药接口：0-不启动；1-启动' From Dual;


--数据结构
--Drop Table 药房发药设备;
Create Table 药房发药设备(
   Id NUMBER(4),
   编码 VARCHAR2(20),
   名称 VARCHAR2(20),
   型号 VARCHAR2(20),
   制造商 VARCHAR2(100),
   使用部门ID NUMBER(18),
   连接类型 NUMBER(1),
   连接内容 VARCHAR2(200),
   服务对象 NUMBER(1),
   是否启用 NUMBER(1))
   TABLESPACE ZL9MEDLST;
Alter Table 药房发药设备 Add Constraint 药房发药设备_PK Primary Key (ID) Using Index Tablespace ZL9INDEXHIS;
Alter Table 药房发药设备 Add Constraint 药房发药设备_UQ_编码 Unique (编码) Using Index Tablespace ZL9INDEXHIS;
Alter Table 药房发药设备 Add constraint 药房发药设备_UQ_使用部门ID unique (使用部门ID, 编码, 名称, 型号) using index  tablespace ZL9INDEXHIS;
Alter table 药房发药设备 add constraint 药房发药设备_FK_使用部门ID foreign key (使用部门ID) references 部门表 (ID);

Create Sequence 药房发药设备_ID Start With 1;

Create Table 自动发药参数(
    Id NUMBER(4),
    参数号 NUMBER(4),
    参数名 VARCHAR2(100),
    参数值 VARCHAR2(4000),
    缺省值 VARCHAR2(4000),
    参数说明 VARCHAR2(255))
    TABLESPACE ZL9MEDLST;
Alter Table 自动发药参数 Add Constraint 自动发药参数_PK Primary Key(ID) Using Index PCTFREE 5;
Alter Table 自动发药参数 Add Constraint 自动发药参数_UQ_参数号 Unique(参数号) Using Index PCTFREE 5;
Alter Table 自动发药参数 Add Constraint 自动发药参数_UQ_参数名 Unique(参数名) Using Index PCTFREE 5;

Insert Into 自动发药参数(ID,参数号,参数名,参数值,缺省值,参数说明)
Select 1,1,'药品剂型',NULL,NULL,'Null表示所有药品剂型；如果需要指定某些剂型，格式：“粉型,片剂,…' From Dual;
  
  
Create Table 药房设备参数(
   参数ID NUMBER(4),
   设备ID NUMBER(4),
   参数值 VARCHAR2(4000))
   TABLESPACE ZL9MEDLST;
Alter Table 药房设备参数 Add Constraint 药房设备参数_PK Primary key (参数ID, 设备ID) Using Index Tablespace ZL9INDEXHIS;
Alter Table 药房设备参数 Add Constraint 药房设备参数_FK_参数ID Foreign key (参数ID) references 自动发药参数 (ID);
Alter Table 药房设备参数 Add Constraint 药房设备参数_FK_设备ID Foreign key (设备ID) references 药房发药设备 (ID) On Delete Cascade;



--设备新增
CREATE OR REPLACE Procedure Zl_药房发药设备_Insert
(
  编码_In         In 药房发药设备.编码%Type,
  名称_In         In 药房发药设备.名称%Type,
  型号_In         In 药房发药设备.型号%Type,
  制造商_In       In 药房发药设备.制造商%Type,
  使用部门id_In   In 药房发药设备.使用部门id%Type,
  连接类型_In     In 药房发药设备.连接类型%Type,
  连接内容_In     In 药房发药设备.连接内容%Type,
  是否启用_In     In 药房发药设备.是否启用%Type,
  服务对象_In     In 药房发药设备.服务对象%Type
) Is
  n_设备id Number;
Begin
  Select 药房发药设备_Id.Nextval Into n_设备id From Dual;

  Insert Into 药房发药设备
    (ID, 编码, 名称, 型号, 制造商, 使用部门id, 连接类型, 连接内容, 服务对象, 是否启用)
  Values
    (n_设备id, 编码_In, 名称_In, 型号_In, 制造商_In, 使用部门id_In, 连接类型_In, 连接内容_In, 服务对象_In, 是否启用_In);

  Insert Into 药房设备参数
    (参数id, 设备id, 参数值)
    Select 1, n_设备id, Null From Dual;
    
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药房发药设备_Insert;
/

--设备更新
CREATE OR REPLACE Procedure Zl_药房发药设备_Update
(
  设备id_In     In 药房发药设备.Id%Type,
  编码_In       In 药房发药设备.编码%Type,
  名称_In       In 药房发药设备.名称%Type,
  型号_In       In 药房发药设备.型号%Type,
  制造商_In     In 药房发药设备.制造商%Type,
  使用部门id_In In 药房发药设备.使用部门id%Type,
  连接类型_In   In 药房发药设备.连接类型%Type,
  连接内容_In   In 药房发药设备.连接内容%Type,
  是否启用_In   In 药房发药设备.是否启用%Type,
  服务对象_In   In 药房发药设备.服务对象%Type
) Is
Begin
  Update 药房发药设备
  Set 编码 = 编码_In, 名称 = 名称_In, 型号 = 型号_In, 制造商 = 制造商_In, 使用部门id = 使用部门id_In, 连接类型 = 连接类型_In, 连接内容 = 连接内容_In,
      服务对象 = 服务对象_In, 是否启用 = 是否启用_In
  Where ID = 设备id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药房发药设备_Update;
/

--设备删除
Create Or Replace Procedure Zl_药房发药设备_Delete
(
  设备id_In In 药房发药设备.Id%Type
) Is
Begin
  Delete 药房发药设备 Where ID = 设备id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药房发药设备_Delete;
/

--设备启用/停用设置
Create Or Replace Procedure Zl_药房发药设备_Switch
(
  设备id_In   In 药房发药设备.Id%Type,
  是否启用_In In 药房发药设备.是否启用%Type
) Is
Begin
  Update 药房发药设备 Set 是否启用 = 是否启用_In Where ID = 设备id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药房发药设备_Switch;
/

--设备参数修改
Create Or Replace Procedure Zl_药房设备参数_Update
(
  参数id_In In 自动发药参数.Id%Type,
  设备id_In In 药房发药设备.Id%Type,
  参数值_In In 药房设备参数.参数值%Type
) Is
Begin
  Update 药房设备参数 Set 参数值 = 参数值_In Where 参数id = 参数id_In And 设备id = 设备id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药房设备参数_Update;
/

