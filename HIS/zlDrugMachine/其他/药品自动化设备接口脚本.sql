--数据结构
Create Table 药品设备接口(
  ID Number(18) Not Null, 
  编号 Varchar2(10) Not Null, 
  名称 Varchar2(20), 
  类型 Number(2), 
  启用日期 Date,
  停用日期 Date, 
  连接信息 Varchar2(2000), 
  扩展信息 Xmltype, 
  备注 Varchar2(200)
)
Pctfree 10 Initrans 1 
Tablespace Zl9medlst;

Create Sequence 药品设备接口_Id Start With 1;

Alter Table 药品设备接口 Add Constraint 药品设备接口_Pk Primary Key(ID) Using Index Tablespace Zl9indexhis;
Alter Table 药品设备接口 Add Constraint 药品设备接口_Uq_编号 Unique(编号) Using Index Tablespace Zl9indexhis;
Alter Table 药品设备接口 Add Constraint 药品设备接口_Uq_名称 Unique(名称) Using Index Tablespace Zl9indexhis;

Create Table 药品收发门诊标志(
  处方号 Varchar2(8), 
  单据 Number(2),
  库房ID Number(18),
  业务分类 Number(2), 
  标志 Number(2),
  待转出 Number(3)
) Pctfree 10 Initrans 20
Tablespace Zl9medlst;

Alter Table 药品收发门诊标志 Add Constraint 药品收发门诊标志_Pk Primary Key(处方号, 单据, 库房ID) Using Index Tablespace Zl9indexhis;
Create Index 药品收发门诊标志_IX_待转出 ON 药品收发门诊标志(待转出) Tablespace Zl9indexhis;

Create Table 药品收发住院标志(
  收发ID NUMBER(18), 
  业务分类 Number(2), 
  标志 Number(2),
  待转出 Number(3)
) Pctfree 10 Initrans 20
Tablespace Zl9medlst;

Alter Table 药品收发住院标志 Add Constraint 药品收发住院标志_Pk Primary Key(收发ID, 业务分类) Using Index Tablespace Zl9indexhis;
Alter Table 药品收发住院标志 Add Constraint 药品收发住院标志_FK_收发ID Foreign Key(收发ID) References 药品收发记录(ID);
Create Index 药品收发住院标志_IX_待转出 ON 药品收发住院标志(待转出) Tablespace Zl9indexhis;

--权限控制
Insert Into zlPrograms(序号,标题,说明,系统,部件) Values(9010,'药品自动化设备接口','药品自动化设备接口的虚拟模块',&n_System,'zlDrugMachine');

Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select &n_System,9010,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0 Union All
  Select '基本',-NULL,NULL,1 From Dual Union All
  Select '参数设置',1,'进行参数设定',0 From Dual Union All
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,9010,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All
  Select '部门表','SELECT' From Dual Union All
  Select '药品设备接口','SELECT' From Dual Union All
  Select '药品剂型','SELECT' From Dual Union All
  Select '部门性质说明','SELECT' From Dual Union All
  Select '部门性质分类','SELECT' From Dual Union All
  Select '人员性质分类','SELECT' From Dual Union All
  Select '人员性质说明','SELECT' From Dual Union All
  Select '人员表','SELECT' From Dual Union All
  Select '上机人员表','SELECT' From Dual Union All
  Select '部门人员','SELECT' From Dual Union All
  Select '药品收发记录','SELECT' From Dual Union All
  Select '药品收发门诊标志','SELECT' From Dual Union All
  Select '药品收发住院标志','SELECT' From Dual Union All
  Select '收费项目目录','SELECT' From Dual Union All
  Select '收费项目别名','SELECT' From Dual Union All
  Select '药品规格','SELECT' From Dual Union All
  Select '药品特性','SELECT' From Dual Union All
  Select '诊疗项目目录','SELECT' From Dual Union All
  Select '药品生产商','SELECT' From Dual Union All
  Select '诊疗项目别名','SELECT' From Dual Union All
  Select '药品库存','SELECT' From Dual Union All
  Select '药品储备限额','SELECT' From Dual Union All
  Select '供应商','SELECT' From Dual Union All
  Select '发药窗口','SELECT' From Dual Union All
  Select '门诊费用记录','SELECT' From Dual Union All
  Select '病人信息','SELECT' From Dual Union All
  Select '身份','SELECT' From Dual Union All
  Select '病人医嘱记录','SELECT' From Dual Union All
  Select '病人诊断医嘱','SELECT' From Dual Union All
  Select '病人诊断记录','SELECT' From Dual Union All
  Select '住院费用记录','SELECT' From Dual Union All
  Select '病人医嘱发送','SELECT' From Dual Union All
  Select '医嘱执行时间','SELECT' From Dual Union All
  Select 'ZL_药品设备接口_UPDATE','EXECUTE' From Dual Union All
  Select 'ZL_药品设备接口_STATE','EXECUTE' From Dual Union All
  Select 'ZL_药品设备接口_DELETE','EXECUTE' From Dual Union All
  Select 'ZL_FUN_DRUG_MACHINE','EXECUTE' From Dual Union All
  Select 'ZL_未发药品记录_分配发药窗口','EXECUTE' From Dual Union All
  Select 'ZL_药品收发门诊标志_FLAG','EXECUTE' From Dual Union All
  Select 'ZL_药品收发住院标志_FLAG','EXECUTE' From Dual Union All
  Select 'ZL_DRUG_MAC_WIN','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;




--应用数据
Insert Into zlBakTables(系统,组号,表名,序号,直接转出,停用触发器)
Select &n_System,2,A.* From (
Select 表名,序号,直接转出,停用触发器 From zlBakTables Where 1 = 0 Union All 
  Select '药品收发门诊标志',8,1,-Null From Dual Union All 
  Select '药品收发住院标志',9,1,-Null From Dual Union All 
Select 表名,序号,直接转出,停用触发器 From ZLBAKTABLES Where 1 = 0) A;


Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
  Select Zlparameters_Id.Nextval, &n_System, 9010, 0, 0, 0, 0, -null, -null, 1, '启用药品自动化设备接口', '0', '0',
         '是否启用药品自动化设备接口向第三方接口提供ZLHIS数据', '0-不启用；1-启用', Null, Null, Null
  From Dual
  Where Not Exists (Select 1 From zlParameters Where 参数名 = '启用药品自动化设备接口' And 模块 = 9010 And 系统 = &n_System);
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
  Select Zlparameters_Id.Nextval, &n_System, 9010, 0, 0, 0, 0, -null, -null, 2, '启用信息交互平台', '0|', '0|',
         '启用时，本参数确定接口是否走信息交互平台与第三方接口交互', '竖线左：0-不启用；1-启用。竖线右：信息交互平台的WebService地址', Null, Null, Null
  From Dual
  Where Not Exists (Select 1 From zlParameters Where 参数名 = '启用信息交互平台' And 模块 = 9010 And 系统 = &n_System);
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
  Select Zlparameters_Id.Nextval, &n_System, 9010, 0, 0, 0, 0, -null, -null, 3, '信息交互平台密钥', '', '',
         '信息交互平台密钥', Null, Null, Null, Null
  From Dual
  Where Not Exists (Select 1 From zlParameters Where 参数名 = '信息交互平台密钥' And 模块 = 9010 And 系统 = &n_System);




--过程函数
CREATE OR REPLACE Procedure Zl_药品设备接口_Update
(
  编号_In     In 药品设备接口.编号%Type,
  名称_In     In 药品设备接口.名称%Type,
  类型_In     In 药品设备接口.类型%Type,
  连接信息_In In 药品设备接口.连接信息%Type,
  扩展信息_In In Varchar2,
  Id_In       In 药品设备接口.Id%Type := Null,
  备注_In     In 药品设备接口.备注%Type := Null
) Is

  v_Error Varchar2(255);
  Err_Custom Exception;

Begin

  --功能：药品设备接口表新增、修改记录

  If Id_In Is Null Then
    --新增
    Insert Into 药品设备接口
      (ID, 编号, 名称, 类型, 连接信息, 扩展信息, 备注)
    Values
      (药品设备接口_Id.Nextval, 编号_In, 名称_In, 类型_In, 连接信息_In, 扩展信息_In, 备注_In);
  Else
    --修改 
    Update 药品设备接口
    Set 编号 = 编号_In, 名称 = 名称_In, 类型 = 类型_In, 连接信息 = 连接信息_In, 扩展信息 = 扩展信息_In, 备注 = 备注_In
    Where ID = Id_In;
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品设备接口_Update;
/

CREATE OR REPLACE Procedure Zl_药品设备接口_State
(
  Id_In   In 药品设备接口.Id%Type,
  启用_In In Number
) Is

  v_Error Varchar2(255);
  Err_Custom Exception;

Begin

  --功能：药品设备接口的状态调整

  If Id_In Is Null Then
    v_Error := '药品设备接口ID不正确！';
    Raise Err_Custom;
  End If;

  If 启用_In = 1 Then
    --启用
    Update 药品设备接口 Set 启用日期 = Sysdate, 停用日期 = Null Where ID = Id_In;
  Else
    --停用
    Update 药品设备接口 Set 停用日期 = Sysdate Where ID = Id_In;
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品设备接口_State;
/

CREATE OR REPLACE Procedure Zl_药品设备接口_Delete(Id_In In 药品设备接口.Id%Type) Is

  v_Error Varchar2(255);
  Err_Custom Exception;

Begin

  --功能：药品设备接口表删除记录

  Delete From 药品设备接口 Where ID = Id_In;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品设备接口_Delete;
/

Create Or Replace Function Zl_Fun_Drug_Machine
(
  库房id_In   In 部门表.Id%Type,
  药品剂型_In In 药品剂型.名称%Type,
  收发id_In   In 药品收发记录.Id%Type := Null
) Return 药品设备接口.编号%Type Is

  v_Code 药品设备接口.编号%Type;

Begin

  --功能：计算参数对应的接口编号
  --说明：药品自动化设备接口部件的专用函数。
  --参数：
  --  收发ID_In：扩展参数，标准调用不传入

  Begin
    Select a.编号
    Into v_Code
    From 药品设备接口 A,
         Xmltable('//root/bm' Passing a.扩展信息 Columns 库房id Number(18) Path 'id', 剂型编码 Varchar2(20) Path 'jxbm') B, 药品剂型 C
    Where b.剂型编码 = c.编码 And a.停用日期 Is Null And a.启用日期 Is Not Null And b.库房id = 库房id_In And c.名称 = 药品剂型_In And
          Rownum < 2;
  Exception
    When Others Then
      Begin
        Select a.编号
        Into v_Code
        From 药品设备接口 A,
             Xmltable('//root/bm' Passing a.扩展信息 Columns 库房id Number(18) Path 'id', 剂型编码 Varchar2(20) Path 'jxbm') B
        Where a.停用日期 Is Null And a.启用日期 Is Not Null And (b.剂型编码 = '' Or b.剂型编码 Is Null) And b.库房id = 库房id_In And
              Rownum < 2;
      Exception
        When Others Then
          v_Code := Null;
      End;
  End;

  Return v_Code;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Fun_Drug_Machine;
/

CREATE OR REPLACE Procedure Zl_药品收发门诊标志_Flag
(
  业务分类_In In 药品收发门诊标志.业务分类%Type,
  库房id_In   In 药品收发门诊标志.库房id%Type,
  处方信息_In In Varchar2,
  传送标志_In In Number
) Is
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin

  --参数
  --  处方信息：单据1,处方号1;单据2,处方号2...

  For r_Tmp In (Select b.c2 处方号, b.c1 单据, a.标志
                From 药品收发门诊标志 A, Table(f_Str2list2(处方信息_In, ';', ',')) B
                Where a.处方号(+) = b.C2 And a.单据(+) = b.C1 And a.库房id(+) = 库房id_In And a.业务分类(+) = 业务分类_In) Loop
    If r_Tmp.标志 Is Null Then
      Delete 药品收发门诊标志 Where 处方号 = r_Tmp.处方号 And 单据 = r_Tmp.单据 And 库房id = 库房id_In;
      If 传送标志_In = 1 Then
        Insert Into 药品收发门诊标志
          (处方号, 单据, 库房id, 业务分类, 标志)
        Values
          (r_Tmp.处方号, r_Tmp.单据, 库房id_In, 业务分类_In, 1);
      Else
        Insert Into 药品收发门诊标志
          (处方号, 单据, 库房id, 业务分类, 标志)
        Values
          (r_Tmp.处方号, r_Tmp.单据, 库房id_In, 业务分类_In, 11);
      End If;
    Elsif r_Tmp.标志 Between 11 And 12 Then
      Update 药品收发门诊标志
      Set 标志 = 标志 + 1
      Where 处方号 = r_Tmp.处方号 And 单据 = r_Tmp.单据 And 库房id = 库房id_In And 业务分类 = 业务分类_In;
    End If;
  End Loop;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品收发门诊标志_Flag;
/

CREATE OR REPLACE Procedure Zl_药品收发住院标志_Flag
(
  业务分类_In In 药品收发住院标志.业务分类%Type,
  医嘱信息_In In Varchar2,
  传送标志_In In Number
) Is
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin

  --参数
  --  医嘱信息：医嘱id1;医嘱2...

  For r_Tmp In (Select b.Column_Value 收发id, a.标志
                From 药品收发住院标志 A, Table(f_Str2list(医嘱信息_In, ';')) B
                Where a.收发id(+) = b.Column_Value And a.业务分类(+) = 业务分类_In) Loop
    If r_Tmp.标志 Is Null Then
      Delete 药品收发住院标志 Where 收发id = r_Tmp.收发id;
      If 传送标志_In = 1 Then
        Insert Into 药品收发住院标志 (收发id, 业务分类, 标志) Values (r_Tmp.收发id, 业务分类_In, 1);
      Else
        Insert Into 药品收发住院标志 (收发id, 业务分类, 标志) Values (r_Tmp.收发id, 业务分类_In, 11);
      End If;
    Elsif r_Tmp.标志 Between 11 And 12 Then
      Update 药品收发住院标志 Set 标志 = 标志 + 1 Where 收发id = r_Tmp.收发id And 业务分类 = 业务分类_In;
    End If;
  End Loop;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品收发住院标志_Flag;
/

Create Or Replace Procedure Zl_Drug_Mac_Win
(
  No_In       In Varchar2,
  库房id_In   In 药品收发记录.库房id%Type,
  窗口编码_In In 发药窗口.编码%Type,
  病人id_In   In 病人医嘱记录.病人id%Type := Null
) Is
  v_Error Varchar2(255);
  Err_Custom Exception;
  v_No   药品收发记录.No%Type;
  v_Tmp  Varchar2(50);
  n_Bill 药品收发记录.单据%Type;
Begin

  --功能：第三方通知ZLHIS门诊处方发药窗口调整
  --参数：
  --  NO_In：设备的NO格式：处方号_单据_库房id

  If No_In Is Null Or No_In = '' Then
    v_Error := '处方信息无';
    Raise Err_Custom;
  End If;

  If 窗口编码_In Is Null Or 窗口编码_In = '' Then
    v_Error := '窗口信息无';
    Raise Err_Custom;
  End If;

  v_No := Substr(No_In, 1, 8);

  If Length(No_In) >= 10 Then
    v_Tmp := Substr(No_In, 10);
  End If;

  If v_Tmp Is Null Or v_Tmp = '' Then
    v_Error := '处方信息异常';
    Raise Err_Custom;
  End If;

  Select Column_Value Into n_Bill From Table(f_Num2list(v_Tmp, '_')) Where Rownum < 2;

  Zl_未发药品记录_分配发药窗口(v_No, n_Bill, 库房id_In, 窗口编码_In);

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drug_Mac_Win;
/

Create Or Replace Procedure Zl1_Datamove_Reb
(
  System_In    In Number,
  Speedmode_In In Number,
  Func_In      In Number,
  Enable_In    In Number := 0,
  Parallel_In  In Number := 0,
  Rebscope_In  In Number := 0
) As
  --功能：在历史数据转出之前，禁用触发器、自动作业、约束、索引，转出之后启用这些对象，以及重建待转出索引，收回标记转出所用索引的空间 
  --参数： 
  --System_In:    应用系统编号,100=标准版 
  --speedmode_in：数据转出模式，0-在线模式，1-离线模式（在客户端停用时，转出期间禁用转出表的主键、唯一键、外键约束和索引，以加快已转数据的删除操作） 
  --func_in:      1=触发器，2=自动作业，3=约束，4=索引，5=重建待转出索引，6-收回标记转出所用索引的空间，7-重组表的存储空间（move），并恢复被禁用的约束和索引 ,8-重建标记转出查询所需索引以外的其他索引 
  --Enable_in:    0-禁用，1=启用，对func_in值为1-4有效 
  --rebScope_in:   Func_In=6时，指重建索引的范围(0-经济核算类,1-经济核算类及医嘱类,2-全部)，Func_In=7时指Move表的范围(0-经济核算类，1-全部) 

  v_Sql      Varchar2(4000);
  n_Do       Number(1);
  n_Parallel Number(1);
  v_Tbs      Varchar2(100);

  --转出标记中的SQL查询所需的索引
  v_Indexeswithtag Varchar2(4000) := '门诊费用记录_IX_结帐ID,住院费用记录_IX_结帐ID,费用补充记录_IX_结算ID,费用补充记录_IX_登记时间,病人预交记录_IX_主页ID,病人预交记录_IX_结帐ID,病人预交记录_IX_收款时间,门诊费用记录_IX_登记时间,门诊费用记录_IX_医嘱序号,住院费用记录_IX_登记时间,病人结帐记录_IX_收费时间,病人结帐记录_IX_病人id' ||
                                     ',药品收发记录_IX_费用ID,收发记录补充信息_IX_收发ID,输液配药内容_IX_收发ID,药品留存计划_IX_留存ID,药品签名明细_IX_收发ID' ||
                                     ',人员借款记录_IX_借出时间,人员收缴记录_IX_登记时间,人员暂存记录_IX_收缴ID,人员暂存记录_IX_登记时间,票据领用记录_IX_登记时间,票据使用明细_IX_领用ID,票据打印明细_IX_使用ID' ||
                                     ',病人挂号记录_IX_登记时间,病人医嘱发送_IX_发送时间,病人医嘱记录_IX_挂号单,病人医嘱记录_IX_主页ID,病人医嘱记录_IX_相关ID' ||
                                     ',病案主页_IX_出院日期,住院费用记录_IX_病人ID,病人过敏记录_IX_病人ID,病人诊断记录_IX_病人ID,病人手麻记录_IX_主页ID' ||
                                     ',病人护理记录_IX_主页ID,病人护理内容_IX_记录id,病人护理文件_IX_主页ID,病人护理数据_IX_文件ID,病人护理明细_IX_记录ID,病人护理打印_IX_文件ID' ||
                                     ',电子病历记录_IX_病人ID,病人医嘱报告_IX_病历ID,影像报告驳回_IX_医嘱ID,报告查阅记录_IX_病历ID,病人诊断记录_IX_病历ID' ||
                                     ',病人临床路径_IX_病人ID,病人合并路径_IX_首要路径记录ID,病人路径执行_IX_路径记录ID,病人出径记录_IX_路径记录ID,病人诊断医嘱_IX_医嘱ID' ||
                                     ',影像报告记录_IX_医嘱ID,影像报告操作记录_IX_医嘱ID,影像申请单图像_IX_医嘱ID,影像收藏内容_IX_医嘱ID,检验标本记录_IX_医嘱ID,检验项目分布_IX_标本ID,检验分析记录_IX_标本ID' ||
                                     ',检验操作记录_IX_标本ID,检验图像结果_IX_标本ID,检验拒收记录_IX_医嘱ID,检验普通结果_IX_检验标本ID,处方审查明细_IX_医嘱ID';

  --转出标记中的SQL查询所需的索引(主键及唯一键对应的索引)
  v_Constraintswithtag Varchar2(4000) := '病人预交记录_UQ_NO,病人结帐记录_UQ_NO,病人结帐记录_PK,门诊费用记录_UQ_NO,住院费用记录_UQ_NO,医保结算明细_PK' ||
                                         ',病人卡结算对照_PK,费用补充记录_PK,病人卡结算记录_PK,三方结算交易_PK,三方退款信息_PK,输液配药记录_PK,药品签名记录_PK,票据打印内容_PK,病人挂号记录_PK,病人挂号汇总_UQ_日期,病人转诊记录_UQ_NO' ||
                                         ',病人护理活动项目_UQ_页号,病人护理要素内容_UQ_页号,产程要素内容_PK,电子病历记录_PK,电子病历附件_PK,电子病历格式_PK,电子病历内容_UQ_对象序号,电子病历图形_PK,疾病申报记录_PK,疾病报告反馈_PK' ||
                                         ',病人合并路径评估_PK,病人路径评估_PK,病人路径变异_PK,病人路径指标_UQ_评估指标,病人路径医嘱_PK' ||
                                         ',病人医嘱记录_PK,病人医嘱报告_PK,病人医嘱计价_UQ_收费细目ID,病人医嘱附费_PK,病人医嘱附件_PK,病人医嘱执行_PK,医嘱执行时间_PK,医嘱执行打印_PK,病人医嘱打印_UQ_医嘱ID,输血申请记录_PK,输血检验结果_PK' ||
                                         ',病人诊断记录_PK,病人医嘱状态_PK,医嘱签名记录_PK,病人医嘱发送_PK,诊疗单据打印_PK,医嘱执行计价_PK,执行打印记录_PK' ||
                                         ',影像检查记录_PK,影像检查序列_UQ_序列号,影像检查图象_UQ_图像号,影像危急值记录_UQ_医嘱ID' ||
                                         ',检验申请项目_PK,检验质控记录_PK,检验签名记录_PK,检验试剂记录_PK,检验质控报告_PK,检验药敏结果_PK,人员收缴记录_PK,人员收缴明细_PK,人员收缴票据_PK,人员收缴对照_PK' ||
                                         ',处方审查记录_PK,处方审查结果_UQ_审方ID,费用清单打印_UQ_NO,RIS检查预约_PK,药品收发门诊标志_PK,药品收发住院标志_PK';

  --功能：1.禁用或启用引用转出表主键的他表外键,避免删除主表记录时对子表每行记录执行一次SQL查询或删除 
  --      2.禁用或启用主键或唯一键约束（禁用时会自动删除对应的索引，启用时自动创建），以提高数据删除性能 
  --例如：病人医嘱发送_FK_医嘱ID，如果这些外键所在的表，数据未转出（未在zlbaktables表中定义），执行前会检查并限制转出。 
  Procedure Setconstraintstatus As
    v_Pcol Varchar2(50);
    v_Fcol Varchar2(50);
    v_Del  Varchar2(4000);
  Begin
    --禁用时，先禁用引用转出表主键的他表外键，再禁用转出表的主键 
    If Enable_In = 0 Then
      --1.在线模式转出时，由于有业务产生删除操作，所以，对于级联删除的外键，用触发器来替代对子表数据的删除操作
      If Speedmode_In = 0 Then
        For Rp In (Select Distinct a.Table_Name As Ptable_Name, a.Constraint_Name
                   From User_Constraints A, User_Constraints C, zlBakTables B
                   Where a.Table_Name = b.表名 And b.直接转出 = 1 And b.系统 = System_In And a.Constraint_Type In ('P', 'U') And
                         c.r_Constraint_Name = a.Constraint_Name And c.Constraint_Type = 'R' And
                         c.Delete_Rule = 'CASCADE'
                   Order By a.Table_Name) Loop
        
          Select f_List2str(Cast(Collect(Column_Name Order By Position) As t_Strlist))
          Into v_Pcol
          From User_Cons_Columns
          Where Constraint_Name = Rp.Constraint_Name;
        
    v_Del := '';
          For Rf In (Select b.Table_Name, b.Constraint_Name,
                            f_List2str(Cast(Collect(b.Column_Name Order By b.Position) As t_Strlist)) As r_Col
                     From User_Constraints A, User_Cons_Columns B
                     Where a.r_Constraint_Name = Rp.Constraint_Name And a.Constraint_Name = b.Constraint_Name
                     Group By b.Table_Name, b.Constraint_Name) Loop
            If Instr(v_Pcol, ',') > 0 Then
              v_Del := v_Del || Chr(10) || '        Delete ' || Rf.Table_Name || ' Where (' || Rf.r_Col ||
                       ') in ((:Old.' || Replace(v_Pcol, ',', ',:Old.') || '));';
            Else
              v_Del := v_Del || Chr(10) || '        Delete ' || Rf.Table_Name || ' Where ' || Rf.r_Col || ' = :Old.' ||
                       v_Pcol || ';';
            End If;
          End Loop;
        
          v_Sql := 'Create Or Replace Trigger ' || Rp.Ptable_Name || '_Cascade_Del' || Chr(10) ||
                   '    After Delete On ' || Rp.Ptable_Name || Chr(10) || '    For Each Row' || Chr(10) || 'Begin' ||
                   Chr(10) || '    If :Old.待转出 Is Null Then ' || v_Del || Chr(10) || '    End If; ' || Chr(10) ||
                   'End ' || Rp.Ptable_Name || '_Cascade_Del;';
          Execute Immediate v_Sql;
        End Loop;
      End If;
    
      --2.禁用引用转出表主键的他表外键
      For R In (Select c.Table_Name, c.Constraint_Name, a.Table_Name As Ptable_Name
                From User_Constraints A, User_Constraints C, zlBakTables B
                Where a.Table_Name = b.表名 And b.直接转出 = 1 And b.系统 = System_In And a.Constraint_Type In ('P', 'U') And
                      c.r_Constraint_Name = a.Constraint_Name And c.Constraint_Type = 'R' And c.Status = 'ENABLED'
                Order By a.Table_Name) Loop
        v_Sql := 'Alter Table ' || r.Table_Name || ' Disable Constraint ' || r.Constraint_Name;
        Execute Immediate v_Sql;
      End Loop;
    
      --3.禁用主键或唯一键索引(离线转出时)
      If Speedmode_In = 1 Then
        --必须删除索引，否则即使skip_unusable_indexes为true，也无法删除存在Unusable状态的唯一性索引的表中的记录
        --保留转出标记中的SQL查询所需的索引(主键和唯一键对应的索引) 
        For R In (Select a.Table_Name, a.Constraint_Name
                  From User_Constraints A, zlBakTables T, User_Tables B
                  Where a.Table_Name = t.表名 And t.直接转出 = 1 And t.系统 = System_In And a.Status = 'ENABLED' And
                        a.Constraint_Type In ('P', 'U') And a.Table_Name = b.Table_Name And b.Iot_Type Is Null And
                        a.Constraint_Name Not In
                        (Select Upper(Column_Value) As Constraint_Name From Table(f_Str2list(v_Constraintswithtag)))
                  Order By Constraint_Name) Loop
          v_Sql := 'Alter Table ' || r.Table_Name || ' Disable Constraint ' || r.Constraint_Name ||
                   ' Cascade Drop Index';
          Execute Immediate v_Sql;
        End Loop;
      End If;
    Else
      --启用时
      --1.先启用主键和唯一键，再启用引用转出表主键的他表外键 
      If Speedmode_In = 1 Then
        --先重建索引，再启用约束，以便重建索引时利用并行执行缩短时间，并且启用约束时也可以采用novalidate方式 
        For R In (Select d.Table_Name, d.Constraint_Name,
                         f_List2str(Cast(Collect(d.Column_Name Order By d.Position) As t_Strlist)) Colstr
                  From User_Cons_Columns D,
                       (Select a.Table_Name, a.Constraint_Name
                         From User_Constraints A, zlBakTables T
                         Where a.Table_Name = t.表名 And t.直接转出 = 1 And t.系统 = System_In And a.Status = 'DISABLED' And
                               a.Constraint_Type In ('P', 'U')) A
                  Where a.Constraint_Name = d.Constraint_Name And a.Table_Name = d.Table_Name
                  Group By d.Table_Name, d.Constraint_Name
                  Order By Constraint_Name) Loop
          Update Zldatamovelog
          Set 当前进度 = '正在恢复约束:' || r.Constraint_Name
          Where 系统 = System_In And 批次 = (Select Max(批次) From Zldatamovelog Where 系统 = System_In);
        
          Select Tablespace_Name Into v_Tbs From User_Indexes Where Table_Name = r.Table_Name And Rownum < 2;
        
          --禁用主键或唯一键时，索引是被删除了的，所以这里要用Create 
          v_Sql := 'Create Unique Index ' || r.Constraint_Name || ' On ' || r.Table_Name || '(' || r.Colstr ||
                   ') Tablespace ' || v_Tbs || ' Nologging';
          Begin
            Execute Immediate v_Sql;
          Exception
            When Others Then
              Null; --可能有些主键或唯一键不是本次转出期间被禁用的，之前就存在不唯一数据，创建唯一索引会出错 
          End;
        
          --会自动建立约束与索引的关联 
          v_Sql := 'Alter Table ' || r.Table_Name || ' Enable Novalidate Constraint ' || r.Constraint_Name;
          Execute Immediate v_Sql;
        End Loop;
      End If;
    
      --2.启用引用转出表主键的他表外键 
      For R In (Select c.Table_Name, c.Constraint_Name
                From User_Constraints A, User_Constraints C, zlBakTables B
                Where a.Table_Name = b.表名 And b.直接转出 = 1 And b.系统 = System_In And a.Constraint_Type In ('P', 'U') And
                      c.r_Constraint_Name = a.Constraint_Name And c.Constraint_Type = 'R' And c.Status = 'DISABLED'
                Order By a.Table_Name) Loop
        --为了加快速度，采用novalidate，不验证已有数据 
        --可能引用转出表主键的他表，在zlbaktables中定义了，但没有编写对应的数据转出脚本，未验证的数据可能有违反约束的情况。 
        v_Sql := 'Alter Table ' || r.Table_Name || ' Enable Novalidate Constraint ' || r.Constraint_Name;
        Execute Immediate v_Sql;
      End Loop;
    
      --3.在线模式转出时，删除之前创建的用来替代级联删除外键的触发器
      If Speedmode_In = 0 Then
        For R In (Select a.Trigger_Name
                  From User_Triggers A, zlBakTables B
                  Where a.Table_Name = b.表名 And b.直接转出 = 1 And b.系统 = System_In And
                        Trigger_Name = Table_Name || '_CASCADE_DEL' And Triggering_Event = 'DELETE') Loop
          v_Sql := 'Drop Trigger ' || r.Trigger_Name;
          Execute Immediate v_Sql;
        End Loop;
      End If;
    End If;
  End Setconstraintstatus;

  --功能：高速模式时禁用LOB以外的所有索引，在线模式时仅禁用转出表引用非转出表的外键索引(例如：病人医嘱计价_IX_收费细目ID) 
  --说明：禁用索引是为了提高删除数据的性能 
  Procedure Setindexstatus As
  Begin
    If Speedmode_In = 1 Then
      --保留转出标记中的SQL查询所需的索引 
      For R In (Select /*+ rule*/
                 a.Index_Name
                From User_Indexes A, zlBakTables T
                Where a.Table_Name = t.表名 And t.直接转出 = 1 And t.系统 = System_In And t.直接转出 = 1 And
                      a.Index_Name <> a.Table_Name || '_IX_待转出' And
                      a.Index_Name Not In
                      (Select Upper(Column_Value) As Index_Name From Table(f_Str2list(v_Indexeswithtag))) And
                      a.Status = Decode(Enable_In, 0, 'VALID', 'UNUSABLE') And a.Index_Type = 'NORMAL' And Not Exists
                 (Select 1
                       From User_Constraints C
                       Where c.Index_Name = a.Index_Name And c.Constraint_Type In ('P', 'U'))
                Order By Index_Name) Loop
      
        If Enable_In = 0 Then
          v_Sql := 'Alter Index ' || r.Index_Name || ' Unusable';
          Execute Immediate v_Sql;
        Else
          Update Zldatamovelog
          Set 当前进度 = '正在重建索引:' || r.Index_Name
          Where 系统 = System_In And 批次 = (Select Max(批次) From Zldatamovelog Where 系统 = System_In);
        
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Nologging';
          Begin
            Execute Immediate v_Sql;
            --在线重建比较慢，不在线重建则需要锁表，如果有其他并发事务，则会出错：ORA-00054: 资源正忙, 但指定以 NOWAIT 方式获取资源 
          
          Exception
            When Others Then
              If SQLErrM Like 'ORA-00054%' Then
                v_Sql := Replace(v_Sql, 'Rebuild', 'Rebuild Online');
                Execute Immediate v_Sql;
              End If;
          End;
        End If;
      End Loop;
    Else
      For R In (Select a.Index_Name
                From (Select d.Table_Name, d.Index_Name,
                              f_List2str(Cast(Collect(d.Column_Name Order By d.Column_Position) As t_Strlist)) Colstr
                       From User_Ind_Columns D, zlBakTables T, User_Indexes C
                       Where c.Table_Name = t.表名 And t.直接转出 = 1 And t.系统 = System_In And c.Uniqueness = 'NONUNIQUE' And
                             c.Index_Type = 'NORMAL' And c.Status = Decode(Enable_In, 0, 'VALID', 'UNUSABLE') And
                             c.Index_Name = d.Index_Name And c.Table_Name = d.Table_Name
                       Group By d.Table_Name, d.Index_Name) A,
                     (Select e.Table_Name,
                              f_List2str(Cast(Collect(e.Column_Name Order By e.Position) As t_Strlist)) Colstr
                       From User_Cons_Columns E, User_Constraints F, zlBakTables T, User_Constraints C
                       Where e.Table_Name = t.表名 And t.直接转出 = 1 And t.系统 = System_In And
                             e.Constraint_Name = f.Constraint_Name And f.Constraint_Type = 'R' And
                             c.Constraint_Name = f.r_Constraint_Name And c.Table_Name Not In ('病案主页', '病人信息') And
                             Not Exists
                        (Select 1 From zlBakTables G Where g.表名 = c.Table_Name And g.系统 = System_In)
                       Group By e.Table_Name, e.Constraint_Name) B
                Where a.Table_Name = b.Table_Name And a.Colstr = b.Colstr
                Order By Index_Name) Loop
      
        If Enable_In = 0 Then
          --特殊处理：以下两个索引不禁用，是由于药品目录修改规格，财务缴款需要使用 
          If r.Index_Name Not In ('病人医嘱记录_IX_收费细目ID', '药品收发记录_IX_药品ID', '药品收发记录_IX_价格ID') Then
            v_Sql := 'Alter Index ' || r.Index_Name || ' Unusable';
            Execute Immediate v_Sql;
          End If;
        Else
          Update Zldatamovelog
          Set 当前进度 = '正在重建索引:' || r.Index_Name
          Where 系统 = System_In And 批次 = (Select Max(批次) From Zldatamovelog Where 系统 = System_In);
        
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Online Nologging';
          Execute Immediate v_Sql;
          --在线重建比较慢，不在线重建则需要锁表，如果有其他并发事务，则会出错：ORA-00054: 资源正忙, 但指定以 NOWAIT 方式获取资源 
        End If;
      End Loop;
    End If;
  End Setindexstatus;

  --功能：转出数据期间，停用转出表上的所有触发器，转出后再恢复 
  Procedure Settriggerstatus As
  Begin
    For R In (Select Distinct a.Table_Name, t.停用触发器
              From User_Triggers A, zlBakTables T
              Where a.Status = Decode(Enable_In, 0, 'ENABLED', 'DISABLED') And a.Table_Name = t.表名 And t.直接转出 = 1 And
                    t.系统 = System_In) Loop
      If Enable_In = 0 Then
        v_Sql := 'Alter Table ' || r.Table_Name || ' DISABLE ALL TRIGGERS';
        Update zlBakTables Set 停用触发器 = 1 Where 系统 = System_In And 表名 = r.Table_Name;
      Elsif Nvl(r.停用触发器, 0) = 1 Then
        v_Sql := 'Alter Table ' || r.Table_Name || ' ENABLE ALL TRIGGERS';
        Update zlBakTables Set 停用触发器 = Null Where 系统 = System_In And 表名 = r.Table_Name;
      End If;
      Execute Immediate v_Sql;
    End Loop;
    Commit;
  End Settriggerstatus;

  --功能：转出数据期间，停用当前所有者的所有自动作业，转出后再启用 
  Procedure Setjobstatus As
    v_Jobs Varchar2(4000);
  Begin
    --停用 
    If Enable_In = 0 Then
      For R In (Select Job From User_Jobs Where Broken = 'N') Loop
        Dbms_Job.Broken(r.Job, True);
        v_Jobs := v_Jobs || ',' || r.Job;
      End Loop;
    
      If v_Jobs Is Not Null Then
        v_Jobs := Substr(v_Jobs, 2);
        Update zlDataMove Set 停用作业号 = v_Jobs Where 系统 = System_In And 组号 = 1;
      End If;
    Else
      --启用 
      Select 停用作业号 Into v_Jobs From zlDataMove Where 系统 = System_In And 组号 = 1;
      If v_Jobs Is Not Null Then
        For R In (Select Job
                  From User_Jobs
                  Where Broken = 'Y' And Job In (Select Column_Value From Table(f_Num2list(v_Jobs)))) Loop
          Dbms_Job.Broken(r.Job, False);
        End Loop;
        Update zlDataMove Set 停用作业号 = Null Where 系统 = System_In And 组号 = 1;
      End If;
    End If;
    --作业设置后必须提交事务才生效 
    Commit;
  End Setjobstatus;
Begin
  If Parallel_In < 2 Then
    Execute Immediate 'Alter Session DISABLE PARALLEL DDL';
  Else
    If Func_In In (6, 7, 8) Or Func_In In (3, 4) And Enable_In = 1 Then
      --为重建索引设置并行执行（由于通常受限于IO设备的性能，设置太高的并行度反而会降低性能，如有高性能存储设备，可加大并行度） 
      --执行重建索引后会自动为索引加上并行度属性，如果不取消，会影响相关SQL的执行计划(全表扫描+并行查询，巨慢),在后面取消索引的并行度 
      --恢复在线库的约束和索引时，不管是不是在线模式，都加上并行，否则太慢
      Execute Immediate 'Alter Session FORCE PARALLEL DDL PARALLEL ' || Parallel_In;
      n_Parallel := 1;
    End If;
  End If;

  If Func_In = 1 Then
    --1.设置触发器 
    Settriggerstatus;
  Elsif Func_In = 2 Then
    --2.设置自动作业 
    Setjobstatus;
  Elsif Func_In = 3 Then
    --3.设置约束状态 
    Setconstraintstatus;
  Elsif Func_In = 4 Then
    --4.设置索引状态 
    Setindexstatus;
  Elsif Func_In = 5 Then
    --5.重建"待转出"索引 
    For R In (Select b.Index_Name
              From zlBakTables A, User_Indexes B
              Where a.表名 = b.Table_Name And a.直接转出 = 1 And a.系统 = System_In And b.Index_Name = b.Table_Name || '_IX_待转出'
              Union All
              Select '病案主页_IX_待转出'
              From Dual
              Where System_In = 100) Loop
      Update Zldatamovelog
      Set 当前进度 = '正在重建索引:' || r.Index_Name
      Where 系统 = System_In And 批次 = (Select Max(批次) From Zldatamovelog Where 系统 = System_In);
    
      --耗时太短，无须并行DDL 
      --在线转出时如果重建索引会锁表，如果有其他并发事务，则会出错：ORA-00054: 资源正忙, 但指定以 NOWAIT 方式获取资源 
      --在线重建索引太慢，所以，即使在线转出模式也不用在线重建
      v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Nologging';
      Begin
        Execute Immediate v_Sql;
      Exception
        When Others Then
          If SQLErrM Like 'ORA-00054%' Then
            v_Sql := Replace(v_Sql, 'Rebuild', 'Rebuild Online');
            Execute Immediate v_Sql;
          End If;
      End;
    End Loop;
  
  Elsif Func_In = 6 Then
    --6.重建标记转出查询所用到的索引（测试表明重建后最多可缩短一半的查询时间） 
    --根据业务的启用阶段来决定重建哪些索引，以避免一些不必要的重建耗时 
    For R In (Select b.Index_Name, a.组号
              From User_Indexes B, zlBakTables A
              Where a.系统 = System_In And a.表名 = b.Table_Name And
                    b.Index_Name In (Select Upper(Column_Value)
                                     From Table(f_Str2list(v_Indexeswithtag))
                                     Union
                                     Select Upper(Column_Value)
                                     From Table(f_Str2list(v_Constraintswithtag)))
              Order By Index_Name) Loop
      n_Do := 0;
      If Rebscope_In = 0 Then
        If r.组号 < 5 Then
          n_Do := 1; --仅经济核算类 
        End If;
      Elsif Rebscope_In = 1 Then
        If r.组号 < 5 Or r.组号 = 8 Then
          n_Do := 1; --仅经济核算类、医嘱类 
        End If;
      Else
        n_Do := 1;
      End If;
    
      If n_Do = 1 Then
        Update Zldatamovelog
        Set 当前进度 = '正在重建索引:' || r.Index_Name
        Where 系统 = System_In And 批次 = (Select Max(批次) From Zldatamovelog Where 系统 = System_In);
      
        --v_Sql := 'Alter Index ' || r.Index_Name || ' shrink Space'; 
        --使用shrink方式不能并行执行,试验表明速度比rebuild PARALLEL 8 慢6倍 
        If Speedmode_In = 1 Then
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Nologging';
        Else
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Online Nologging';
        End If;
        Begin
          Execute Immediate v_Sql;
          --在线重建比较慢，不在线重建则需要锁表，如果有其他并发事务，则会出错：ORA-00054: 资源正忙, 但指定以 NOWAIT 方式获取资源 
        
        Exception
          When Others Then
            If SQLErrM Like 'ORA-00054%' Then
              v_Sql := Replace(v_Sql, 'Rebuild', 'Rebuild Online');
              Execute Immediate v_Sql;
            End If;
        End;
      End If;
    End Loop;
  
    --重组表的数据
  Elsif Func_In = 7 Then
    --rebScope_in=0,只重组组号小于5的经济核算类表（费用、药品、票据），否则全部重组 
    For R In (Select a.表名 As Table_Name
              From zlBakTables A
              Where a.直接转出 = 1 And (组号 < Decode(Rebscope_In, 0, 5, 100))
              Order By 组号, 序号) Loop
    
      Update Zldatamovelog
      Set 当前进度 = '正在重组表:' || r.Table_Name
      Where 系统 = System_In And 批次 = (Select Max(批次) From Zldatamovelog Where 系统 = System_In);
    
      --如果有空闲的空间，最好移到其他表空间，只有这样才能绝对移动文件尾部的数据块，以便进行表空间文件的收缩 
      --在前面设置了会话级的强制并行 
      v_Sql := 'Alter Table ' || r.Table_Name || ' Move Nologging';
      Execute Immediate v_Sql;
    
      --单独移动Lob对象 
      For L In (Select Column_Name, Tablespace_Name From User_Lobs Where Table_Name = r.Table_Name) Loop
        v_Sql := 'Alter Table ' || r.Table_Name || ' Move Lob(' || l.Column_Name || ') Store as (Tablespace ' ||
                 l.Tablespace_Name || ') Nologging';
        Execute Immediate v_Sql;
      End Loop;
    
      v_Sql := 'Alter Table ' || r.Table_Name || ' Noparallel';
      Execute Immediate v_Sql;
    
      --move后，表相关的索引会全部失效，需要全部重建 
      For S In (Select Index_Name
                From User_Indexes
                Where Table_Name = r.Table_Name And Status = 'UNUSABLE'
                Order By Index_Name) Loop
        Update Zldatamovelog
        Set 当前进度 = '正在恢复失效索引:' || s.Index_Name
        Where 系统 = System_In And 批次 = (Select Max(批次) From Zldatamovelog Where 系统 = System_In);
      
        --在前面设置了会话级的强制并行 
        v_Sql := 'Alter Index ' || s.Index_Name || ' Rebuild Nologging';
        Execute Immediate v_Sql;
      End Loop;
    End Loop;
    --重建转出表上标记转出以外的其他索引（用于转出完成后收回空闲空间）
    --失效的索引不重建，因为转出完后有单独的重建功能
  Elsif Func_In = 8 Then
    For R In (Select b.Index_Name, a.组号
              From User_Indexes B, zlBakTables A
              Where a.系统 = System_In And a.表名 = b.Table_Name And b.Status = 'VALID' And b.Index_Type = 'NORMAL' And
                    b.Index_Name Not Like 'BIN$%' And
                    b.Index_Name Not In (Select Upper(Column_Value)
                                         From Table(f_Str2list(v_Indexeswithtag))
                                         Union
                                         Select Upper(Column_Value)
                                         From Table(f_Str2list(v_Constraintswithtag)))
              Order By Index_Name) Loop
      Update Zldatamovelog
      Set 当前进度 = '正在重建索引:' || r.Index_Name
      Where 系统 = System_In And 批次 = (Select Max(批次) From Zldatamovelog Where 系统 = System_In);
    
      If Speedmode_In = 1 Then
        v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Nologging';
      Else
        v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Online Nologging';
      End If;
      Begin
        Execute Immediate v_Sql;
        --在线重建比较慢，不在线重建则需要锁表，如果有其他并发事务，则会出错：ORA-00054: 资源正忙, 但指定以 NOWAIT 方式获取资源    
      Exception
        When Others Then
          If SQLErrM Like 'ORA-00054%' Then
            v_Sql := Replace(v_Sql, 'Rebuild', 'Rebuild Online');
            Execute Immediate v_Sql;
          End If;
      End;
    End Loop;
  End If;

  --执行重建索引后会自动为索引加上并行度属性，如果不取消，会影响相关SQL的执行计划(全表扫描+并行查询，巨慢) 
  --------------------------------------------------------------------------------------------------- 
  If n_Parallel = 1 Then
    Execute Immediate 'ALTER Session DISABLE PARALLEL DDL';
  
    For R In (Select Index_Name From User_Indexes Where Degree Not In ('1', '0')) Loop
      v_Sql := 'Alter Index ' || r.Index_Name || ' Noparallel';
      Execute Immediate v_Sql;
    End Loop;
  End If;

  Update Zldatamovelog
  Set 当前进度 = '重建完成'
  Where 系统 = System_In And 批次 = (Select Max(批次) From Zldatamovelog Where 系统 = System_In);
  Commit;
  --本过程不进行错误处理，错误由调用过程处理 
End Zl1_Datamove_Reb;
/

Create Or Replace Procedure Zl1_Datamove_Tag
(
  d_End    In Date,
  n_批次   In Number,
  n_System In Number
) As
  --功能：标记待转出的数据 
  --说明：为避免Undo表空间膨胀过大，分段提交 
Begin
  --1.经济核算（费用,药品,收款和票据等）  
  --新加子查询注意性能优化，把能够将数据过滤到最小的条件放到最后，Exists类条件放前面
  Update /*+ rule*/ 病人预交记录 L
  Set 待转出 = n_批次
  Where 结帐id In
        (Select Distinct a.结帐id --1.门诊收费和挂号的收费结算记录(排除之后退号和退费的,一张单据中只要其中一行退了) 
         From 门诊费用记录 A
         Where (a.记录状态 In (1, 2) Or a.记录状态 = 3 And Not Exists
                (Select 1
                 From 门诊费用记录 B
                 Where a.No = b.No And a.记录性质 = b.记录性质 And b.记录状态 = 2 And b.登记时间 >= d_End))
     And a.待转出 Is Null And a.记录性质 In (1, 4) And a.登记时间 < d_End
         Union All
         Select Distinct a.结算id --2.医保补结算 
         From 费用补充记录 A
         Where (a.记录状态 In (1, 2) Or a.记录状态 = 3 And Not Exists
                (Select 1
                 From 费用补充记录 B
                 Where a.No = b.No And a.记录性质 = b.记录性质 And b.记录状态 In (1, 2) And b.登记时间 >= d_End))
     And a.待转出 Is Null And a.记录性质 = 1 And a.登记时间 < d_End
         Union All
         Select Distinct a.结帐id --3.就诊卡的收费结算记录(排除之后退卡费的,一张单据中只要其中一行退了) 
         From 住院费用记录 A
         Where (a.记录状态 In (1, 2) Or a.记录状态 = 3 And Not Exists
                (Select 1
                 From 住院费用记录 B
                 Where a.No = b.No And a.记录性质 = b.记录性质 And b.记录状态 = 2 And b.登记时间 >= d_End))
    And a.待转出 Is Null And a.记帐费用 = 0 And a.记录性质 = 5 And a.登记时间 < d_End
         Union All --4.门诊(记帐单)和住院的结帐结算记录 
         Select 结帐id
         From (With Settle As (Select Distinct a.Id As 结帐id, a.病人id --3.门诊(记帐单)和住院的结帐结算记录(排除之后结帐作废的) 
                               From 病人结帐记录 A
                               Where (a.记录状态 In (1, 2) Or a.记录状态 = 3 And Not Exists
                                      (Select 1 From 病人结帐记录 B Where a.No = b.No And b.记录状态 = 2 And b.收费时间 >= d_End))
              And a.待转出 Is Null And a.收费时间 < d_End)
                Select 结帐id
                From Settle
                Minus
                --以下结帐ID要整体排除,避免部分费用明细被转出后影响后续的计算是否冲完 
                --1.一张预交款被多笔结帐冲完（结帐ID不同）
                --2.费用单据的结帐ID相关的可能还有其他NO的其他结帐ID(结帐作废后分多次结帐结清，可能部分在转出时间之后)
                --考虑到这情况的复杂性，为简化逻辑，提升查询性能，按病人ID来排除 
                Select Distinct d.Id
                From 病人结帐记录 D,
                     (Select Distinct c.病人id --多次住院可以一起结，以及门诊记帐和住院记帐可以一起结且冲同一笔预交，所以这里不加主页ID 
                       From 住院费用记录 C,
                            (Select Distinct d.No, d.序号, Mod(d.记录性质, 10) As 记录性质
                              From 住院费用记录 D,
                                   (Select s.结帐id From Settle S, 病人结帐记录 E --没有结清且该病人之后没有再结过就成了呆帐，这种就不排除 
                                     Where s.病人id = e.病人id And (e.收费时间 > d_End Or Exists (Select 1 From 在院病人 F Where s.病人id = f.病人id))) S 
                              Where d.结帐id = s.结帐id) D
                       Where c.No = d.No And Mod(c.记录性质, 10) = d.记录性质 And c.序号 = d.序号 --结帐后作废后，再对包含记帐单销帐的结帐ID为空的记录,一起汇总计算是否结清,这种结帐ID为空的数据转出在后面单独转出 
                       Group By c.No, Mod(c.记录性质, 10), c.病人id --一张单据中的一行可部分结帐，以单据为对象来判断，避免一张单据的其中一部分被转出 
                       Having Nvl(Sum(c.实收金额), 0) <> Nvl(Sum(c.结帐金额), 0) Or Exists (Select 1 --排除转出时间之后再次结帐的(作废后再次结帐)，避免原始单据转走后，后续结帐时无法正确判断 
                                                                                   From 住院费用记录 E, 病人结帐记录 S
                                                                                   Where e.No = c.No And Mod(e.记录性质, 10) = Mod(c.记录性质, 10) And
                                                                                         e.记录性质 In (12, 13, 15) And e.结帐id = s.Id  And s.待转出 Is Null And s.收费时间 >= d_End)
                       Union All
                       Select Distinct c.病人id
                       From 门诊费用记录 C,
                            (Select Distinct d.No, d.序号, Mod(d.记录性质, 10) As 记录性质
                              From 门诊费用记录 D, Settle S
                              Where d.结帐id = s.结帐id) D --因为是门诊病人，所以，只要没有结清,该病人的都不转出 
                       Where c.No = d.No And Mod(c.记录性质, 10) = d.记录性质 And c.序号 = d.序号
                       Group By c.No, Mod(c.记录性质, 10), c.病人id
                       Having Nvl(Sum(c.实收金额), 0) <> Nvl(Sum(c.结帐金额), 0) Or Exists (Select 1
                                                                                   From 门诊费用记录 E, 病人结帐记录 S
                                                                                   Where e.No = c.No And Mod(e.记录性质, 10) = Mod(c.记录性质, 10) And
                                                                                         e.记录性质 In (12, 13, 15) And e.结帐id = s.Id And s.待转出 Is Null And s.收费时间 >= d_End)) N
                Where d.病人id = n.病人id)
         );

  --排除预交款未冲完的
  --为了降低逻辑的复杂性，不排除在转出时间之后发药或未发药的费用记录对应的结帐ID，将这种情况的结算数据和费用数据强制转走 
  --因为前面的SQL查出的结帐ID可能不全是冲预交的(门诊收费和住院结帐补费等)，所以，需要单独一个SQL来排除 
  --由于可能存在数据异常(住院费用结帐冲预交类别为1的门诊预交)，所以没有加预交类别条件限定 
  Update /*+ rule*/ 病人预交记录
  Set 待转出 = Null
  Where 待转出 = n_批次 And
        结帐id In (Select Distinct d.结帐id
                 From 病人预交记录 D,
                      --连接D表是为了查冲同一预交单据的其他结帐ID（退预交款，冲预交作废的，再次冲同一预交单据） 
                      --该预交或冲预交单据涉及的所有结帐ID的都不转出，避免部分冲预交的结帐ID被排除后，原始预交单被转走，或者其他结帐ID将费用单据的一部分(原始结帐、结帐作废、再次结一部分、再次结全部)转走 
                      (Select Distinct l.No
                        From 病人预交记录 L, 病人预交记录 P --可能本次结帐冲的只是剩余款，所以需要连接L表，查原始交预交的单据，以及记录性质为11的可能还有转出时间之后其他冲剩余款的结帐ID 
                        Where l.记录性质 = p.记录性质 And l.No = p.No And p.记录性质 In (1, 11) And p.待转出 = n_批次
                        Group By l.No, l.病人id
                        Having Nvl(Sum(l.金额), 0) <> Nvl(Sum(l.冲预交), 0) And (Exists (Select 1
                                                                                  From 病人预交记录 E --没有冲完且之后没有再冲过或结算过就成了呆帐（以及存在用负的结帐补款来表示冲预交当成冲完的清况），这种就不排除
                                                                                  Where l.病人id = e.病人id And e.待转出 Is Null And e.收款时间 > d_End)
                                                                                  Or Exists (Select 1 From 在院病人 E Where l.病人id =e.病人id)
                                                                                  Or Exists (Select 1 From 病人未结费用 E Where l.病人id =e.病人id))  
                        Or Nvl(Sum(l.金额), 0) = Nvl(Sum(l.冲预交), 0) And Exists (Select 1
                                                                                  From 病人预交记录 E --排除转出时间之后的其他结帐ID冲的,10.34.20后，冲预交全部单独增加了一条记录，收费时间就是冲预交时间(以前是在原始交预交款的记录上填冲预交字段，不能直接查到冲预交款的时间)
                                                                                  Where e.No = l.No And e.记录性质 = 11 And e.待转出 Is Null And e.收款时间 >= d_End)) N
                 Where d.No = n.No And d.记录性质 In (1, 11));

  --预交款没有使用就直接退了的记录(结帐ID为空) 
  Update /*+ rule*/ 病人预交记录
  Set 待转出 = n_批次
  Where 记录性质 = 1 And
        NO In (Select a.No
               From 病人预交记录 A
               Where a.结帐id Is Null And a.记录性质 = 1 And a.记录状态 In (2, 3) And a.待转出 Is Null And a.收款时间 < d_End
               Group By a.No
               Having Sum(a.金额) = 0);

  --冲预交款作废的记录（记录性质为2），没有结帐ID 
  Update /*+ rule*/ 病人预交记录
  Set 待转出 = n_批次
  Where 结帐id Is Null And 记录性质 = 2 And NO In (Select a.No From 病人预交记录 A Where a.待转出 = n_批次 And a.记录性质 = 3);

  Update Zldatamovelog
  Set 当前进度 = '(1/10)结算数据标记完成，正在标记费用数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  Update /*+ rule*/ 病人结帐记录
  Set 待转出 = n_批次
  Where ID In (Select 结帐id From 病人预交记录 Where 待转出 = n_批次);

  --结帐无结算的记录(为了提升性能，不判断费用，只要结了帐且无预交记录就当成是零费用结帐) 
  Update /*+ rule*/ 病人结帐记录 L
  Set 待转出 = n_批次
  Where 收费时间 < d_End And 待转出 Is Null And Not Exists (Select 1 From 病人预交记录 P Where l.Id = p.结帐id);

  Update /*+ rule*/ 病人卡结算对照
  Set 待转出 = n_批次
  Where 预交id In (Select a.Id From 病人预交记录 A Where 待转出 = n_批次);

  Update /*+ rule*/ 病人卡结算记录
  Set 待转出 = n_批次
  Where ID In (Select 卡结算id From 病人卡结算对照 Where 待转出 = n_批次);

  Update /*+ rule*/ 三方结算交易
  Set 待转出 = n_批次
  Where 交易id In (Select a.Id From 病人预交记录 A Where 待转出 = n_批次);

  Update /*+ rule*/ 三方退款信息
  Set 待转出 = n_批次
  Where (记录id,结帐ID) In (Select a.Id,A.结帐ID From 病人预交记录 A Where 待转出 = n_批次);

  --1.挂号打折后实收金额为0的(没有对应的预交记录),即使之后有退号费用也不管，因为金额为零不影响计算),而卡费即使为零也有预交记录 
  --结帐ID为空的是异常数据（德阳医院仅有3笔此类数据）
  --根据挂号记录再找门诊费用，比直接按时间查门诊费用要快 
  Update /*+ rule*/ 门诊费用记录
  Set 待转出 = n_批次
  Where NO In (Select NO From 病人挂号记录 Where 待转出 Is Null And 登记时间 < d_End) And 记录性质 = 4 And (实收金额 = 0 Or 结帐id Is Null);

  --2.直接收费的和结帐无结算（预交）记录的，Union不加all去掉重复以减少in的数量 
  Update /*+ rule*/ 门诊费用记录
  Set 待转出 = n_批次
  Where 结帐id In
        (Select 结帐id From 病人预交记录 Where 待转出 = n_批次 Union Select ID From 病人结帐记录 Where 待转出 = n_批次);

  --3.没有结帐id的数据(按登记时间)
  --1)未结帐的门诊记帐费用(赖账)，该病人没有预交记录或冲预交记录，并且该时间之后无门诊费用发生
  --2)未结帐的划价记录
  --3)未收费（也没有冲预交）的零费用
  --加条件"待转出 Is Null"是为了处理连续多次标记转出的情况 
  Update /*+ rule*/ 门诊费用记录 A
  Set 待转出 = n_批次
  Where (Not Exists (Select 1 From 病人预交记录 B Where a.病人id = b.病人id And b.待转出 Is Null And 记录性质 In (1, 11)) And Not Exists
         (Select 1 From 门诊费用记录 B Where a.病人id = b.病人id And b.待转出 Is Null And 登记时间 > d_End) And 记录性质 = 2 Or 记录状态 = 0 Or
         记录性质 = 1 And 实收金额 = 0 And 结帐金额 = 0) And 结帐id Is Null And 待转出 Is Null And 登记时间 < d_End;

  --4.没有结帐id的数据(按发生时间)
  --冲销产生的记帐记录（记录状态为2），登记时间可能在当前指定转出时间之后，而原始记帐记录（记录状态为3），登记时间在指定转出时间之前。前后两者的发生时间是相同的。
  --1)未结帐的零记帐费用或打折后实收金额为零的（结帐模块参数没有勾选对零费用结帐）
  --2)结帐作废后，记帐单销帐的记录（结帐ID为空且记录状态为2的），记录状态为3的且有结帐ID的在最前面已转出. 
  Update /*+ rule*/ 门诊费用记录 A
  Set 待转出 = n_批次
  Where (Exists (Select 1
                 From 门诊费用记录 B
                 Where a.No = b.No And a.记录性质 = b.记录性质 And a.序号 = b.序号 And b.记录状态 = 3 And b.结帐id Is Not Null And
                       b.待转出 + 0 = n_批次) And 记录状态 = 2 Or Exists
         (Select 1
          From 门诊费用记录 B
          Where a.No = b.No And a.记录性质 = b.记录性质 And a.序号 = b.序号 And b.结帐id Is Null
          Group By b.No, b.记录性质, b.序号
          Having Nvl(Sum(b.实收金额), 0) = 0)) And 记录性质 = 2 And 结帐id Is Null And 待转出 Is Null And 发生时间 < d_End;

  --5.有结帐id的零费用(按发生时间)
  --按费别打折后结帐金额为零的收费记录,或者一张单据相同结帐ID的结帐金额之和为0(冲销后为零)
  --即使在转出时间之后发药的，也强制转出（为了减少逻辑复杂性，提高查询性能）
  Update /*+ rule*/ 门诊费用记录 A
  Set 待转出 = n_批次
  Where (结帐金额 = 0 Or Exists
         (Select 1 From 门诊费用记录 C Where a.结帐id = c.结帐id Group By c.结帐id, c.No Having Sum(c.结帐金额) = 0)) And Not Exists
   (Select 1 From 病人预交记录 B Where a.结帐id = b.结帐id And b.待转出 Is Null) And 记录性质 = 1 And 结帐id Is Not Null And
        待转出 Is Null And 发生时间 < d_End;

  Update /*+ rule*/ 医保结算明细
  Set 待转出 = n_批次
  Where 结帐id In (Select 结帐id From 病人预交记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 费用补充记录
  Set 待转出 = n_批次
  Where 结算id In (Select 结帐id From 病人预交记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 凭条打印记录
  Set 待转出 = n_批次
  Where (NO, 记录性质) In (Select NO, 记录性质 From 门诊费用记录 Where 待转出 = n_批次);

   --1.从预交记录读是为了取就诊卡直接收费的（无结帐ID）,再加结帐记录是为了取结帐无结算（预交）记录的 
  Update /*+ rule*/ 住院费用记录
  Set 待转出 = n_批次
  Where 结帐id In
        (Select 结帐id From 病人预交记录 Where 待转出 = n_批次 Union Select ID From 病人结帐记录 Where 待转出 = n_批次);

  --2.没有结帐id的数据(按发生时间)
  --冲销产生的记帐记录（记录状态为2），登记时间可能在当前指定转出时间之后，而原始记帐记录（记录状态为3），登记时间在指定转出时间之前。前后两者的发生时间是相同的。
  --1)转出结帐作废后，记帐单销帐的记录（记帐状态为2且没有结帐ID且(记录状态为3的有结帐ID的)在最前面已转出） 
  --2)未结帐的零费用(已冲销的记帐单或打折后实收金额为零) 
  --3)没有结帐ID的划价记录处理为转出 
  
  Update /*+ rule*/ 住院费用记录 A
  Set 待转出 = n_批次
  Where ((Exists (Select 1
                  From 住院费用记录 B
                  Where a.No = b.No And a.记录性质 = b.记录性质 And a.序号 = b.序号 And b.记录状态 = 3 And b.结帐id Is Not Null And
                        b.待转出 + 0 = n_批次) And 记录状态 = 2 Or Exists
         (Select 1
           From 住院费用记录 B
           Where a.No = b.No And a.记录性质 = b.记录性质 And a.序号 = b.序号 And b.结帐id Is Null
           Group By b.No, b.记录性质, b.序号
           Having Nvl(Sum(b.实收金额), 0) = 0)) And 记录性质 = 2 Or 记录状态 = 0) And 结帐id Is Null And 待转出 Is Null And 发生时间 < d_End;

  --3.离院未结帐的（赖帐病人），因为是很久以前的这些数据，如果预交已冲完，则处理为要转出 
  Update /*+ rule*/ 住院费用记录 A
  Set 待转出 = n_批次
  Where 待转出 Is Null And 结帐id Is Null And
        (病人id, 主页id) In (Select 病人id, 主页id
                         From 病案主页 C
                         Where 出院日期 < d_End And 待转出 Is Null And 数据转出 Is Null And Not Exists
                          (Select 1
                                From 病人预交记录 B
                                Where b.病人id = c.病人id And b.待转出 Is Null And b.预交类别 = 2 And b.记录性质 In (1, 11)
                                Having Nvl(Sum(b.金额), 0) - Nvl(Sum(b.冲预交), 0) <> 0));

  Update /*+ rule*/ 费用清单打印
  Set 待转出 = n_批次
  Where (NO, Mod(记录性质,10),Decode(记录状态,3,1,记录状态),序号) In 
        (Select NO, Mod(记录性质,10) as 记录性质,Decode(记录状态,3,1,记录状态) as 记录状态,序号 From 门诊费用记录 Where 待转出 = n_批次
        Union
        Select NO, Mod(记录性质,10) as 记录性质,Decode(记录状态,3,1,记录状态) as 记录状态,序号 From 住院费用记录 Where 待转出 = n_批次);

  Update Zldatamovelog
  Set 当前进度 = '(2/10)费用数据标记完成，正在标记药品数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  Update /*+ Rule*/ 药品收发记录 A
  Set 待转出 = n_批次
  Where Rowid In (Select m.Rowid
                  From 药品收发记录 M, 门诊费用记录 E
                  Where m.费用id = e.Id And (e.记录性质 = 1 And m.单据 In (8, 24) Or e.记录性质 = 2 And m.单据 In (9, 25)) And
                        e.收费类别 In ('4', '5', '6', '7') And e.待转出 = n_批次
                  Union All
                  Select m.Rowid
                  From 药品收发记录 M, 住院费用记录 E
                  Where m.费用id = e.Id And m.单据 In (9, 10, 25, 26) And e.记录性质 = 2 And e.收费类别 In ('4', '5', '6', '7') And
                        e.待转出 = n_批次);

  Update /*+ rule*/ 收发记录补充信息
  Set 待转出 = n_批次
  Where 收发id In (Select ID From 药品收发记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 输液配药内容
  Set 待转出 = n_批次
  Where 收发id In (Select ID From 药品收发记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 输液配药记录
  Set 待转出 = n_批次
  Where ID In (Select 记录id From 输液配药内容 Where 待转出 = n_批次);

  Update /*+ rule*/ 输液配药附费
  Set 待转出 = n_批次
  Where 配药id In (Select ID From 输液配药记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 输液配药状态
  Set 待转出 = n_批次
  Where 配药id In (Select ID From 输液配药记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 药品留存计划
  Set 待转出 = n_批次
  Where 留存id In (Select ID From 药品收发记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 药品签名明细
  Set 待转出 = n_批次
  Where 收发id In (Select ID From 药品收发记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 药品签名记录
  Set 待转出 = n_批次
  Where ID In (Select 签名id From 药品签名明细 Where 待转出 = n_批次);

  Update /*+ rule*/ 药品收发门诊标志 A
  Set 待转出 = n_批次
  Where Exists(Select 1 From 药品收发记录 B Where b.NO = a.处方号 And b.单据 = a.单据 And b.待转出 = n_批次);

  Update /*+ rule*/ 药品收发住院标志
  Set 待转出 = n_批次
  Where 收发id In (Select ID From 药品收发记录 Where 待转出 = n_批次);

  Update Zldatamovelog
  Set 当前进度 = '(3/10)药品数据标记完成，正在标记缴款与票据数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  Update /*+ rule*/ 人员借款记录 Set 待转出 = n_批次 Where 待转出 Is Null And 借出时间 < d_End;

  Update /*+ rule*/ 人员收缴记录 Set 待转出 = n_批次 Where 待转出 Is Null And 登记时间 < d_End;

  Update /*+ rule*/ 人员收缴对照
  Set 待转出 = n_批次
  Where 收缴id In (Select ID From 人员收缴记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 人员收缴明细
  Set 待转出 = n_批次
  Where 收缴id In (Select ID From 人员收缴记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 人员收缴票据
  Set 待转出 = n_批次
  Where 收缴id In (Select ID From 人员收缴记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 人员暂存记录
  Set 待转出 = n_批次
  Where 收缴id In (Select ID From 人员收缴记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 人员暂存记录 Set 待转出 = n_批次 Where 待转出 Is Null And 记录性质 = 1 And 登记时间 < d_End;

  Update /*+ rule*/ 票据领用记录 A
  Set 待转出 = n_批次
  Where Not Exists
   (Select 1 From 票据使用明细 B Where b.领用id = a.Id And b.使用时间 >= d_End) And 待转出 Is Null And 剩余数量 = 0 And 登记时间 < d_End;

  Update /*+ rule*/ 票据使用明细
  Set 待转出 = n_批次
  Where 领用id In (Select ID From 票据领用记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 票据打印内容
  Set 待转出 = n_批次
  Where ID In (Select 打印id From 票据使用明细 Where 待转出 = n_批次);

  Update /*+ rule*/ 票据打印明细
  Set 待转出 = n_批次
  Where 使用id In (Select ID From 票据使用明细 Where 待转出 = n_批次);

  Update Zldatamovelog
  Set 当前进度 = '(4/10)缴款与票据数据标记完成，正在标记就诊及诊治数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  --2.就诊及诊治数据 
  --不转出的条件：挂号费用未转出的，转出时间之后存在医嘱，医嘱对应的费用未转出的 
  --即使正在就诊(r.执行状态 <> 2 )的也强制转出 
  Update /*+ rule*/ 病人挂号记录 T
  Set 待转出 = n_批次
  Where Rowid In
        (Select Rowid
         From 病人挂号记录 R
         Where Not Exists (Select 1
                From 门诊费用记录 A
                Where r.No = a.No And a.登记时间 < d_End And a.记录性质 = 4 And a.待转出 Is Null) And Not Exists
          (Select 1
                From 病人医嘱记录 A
                Where a.挂号单 = r.No And a.待转出 Is Null And a.病人来源 <> 4 And Nvl(a.停嘱时间, a.开嘱时间) >= d_End) And Not Exists
          (Select 1
                From 门诊费用记录 E, 病人医嘱记录 A
                Where r.No = a.挂号单 And a.Id = e.医嘱序号 And a.病人来源 <> 4 And e.待转出 Is Null) And r.待转出 Is Null And
               r.登记时间 < d_End);

  --由于有一部分挂号数据未转出，所以，汇总表的数据可能与挂号数据不匹配 
  Update 病人挂号汇总 Set 待转出 = n_批次 Where 待转出 Is Null And 日期 < d_End;
  Update /*+ rule*/ 病人转诊记录 Set 待转出 = n_批次 Where NO In (Select NO From 病人挂号记录 Where 待转出 = n_批次);

  --通过"住院费用记录"来查询，而不是"病人结帐记录",因为离院未结的赖帐病人也转出了费用 
  --出院日期条件仍然需要，因为可能某次结帐转出了，但病人当时并未出院(一次住院多次结帐)。 
  --通过指定索引方式进行特殊优化（缺省采用"病案主页IX_出院日期"索引的效率太低） 
  Update /*+ rule*/ 病案主页 P
  Set 待转出 = n_批次
  Where Not Exists (Select 1 From 住院费用记录 A Where a.病人id = p.病人id And a.主页id = p.主页id And a.待转出 Is Null) And 待转出 Is Null And
        数据转出 Is Null And 出院日期 < d_End And
        (病人id, 主页id) In (Select Distinct 病人id, 主页id From 住院费用记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人过敏记录
  Set 待转出 = n_批次
  Where (病人id, 主页id) In (Select 病人id, ID
                         From 病人挂号记录
                         Where 待转出 = n_批次
                         Union All
                         Select 病人id, 主页id
                         From 病案主页
                         Where 待转出 = n_批次);

  Update /*+ rule*/ 病人诊断记录
  Set 待转出 = n_批次
  Where (病人id, 主页id) In (Select 病人id, ID
                         From 病人挂号记录
                         Where 待转出 = n_批次
                         Union All
                         Select 病人id, 主页id
                         From 病案主页
                         Where 待转出 = n_批次);

  Update /*+ rule*/ 病人手麻记录
  Set 待转出 = n_批次
  Where (病人id, 主页id) In (Select 病人id, ID
                         From 病人挂号记录
                         Where 待转出 = n_批次
                         Union All
                         Select 病人id, 主页id
                         From 病案主页
                         Where 待转出 = n_批次);

  Update Zldatamovelog
  Set 当前进度 = '(5/10)就诊及诊治数据标记完成，正在标记护理数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  --3.护理数据 
  Update /*+ rule*/ 病人护理文件
  Set 待转出 = n_批次
  Where (病人id, 主页id) In (Select 病人id, 主页id From 病案主页 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人护理数据
  Set 待转出 = n_批次
  Where 文件id In (Select ID From 病人护理文件 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人护理明细
  Set 待转出 = n_批次
  Where 记录id In (Select ID From 病人护理数据 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人护理打印
  Set 待转出 = n_批次
  Where 文件id In (Select ID From 病人护理文件 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人护理活动项目
  Set 待转出 = n_批次
  Where 文件id In (Select ID From 病人护理文件 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人护理要素内容
  Set 待转出 = n_批次
  Where 文件id In (Select ID From 病人护理文件 Where 待转出 = n_批次);
  Update /*+ rule*/ 产程要素内容
  Set 待转出 = n_批次
  Where 文件id In (Select ID From 病人护理文件 Where 待转出 = n_批次);

  --老版护理系统数据 
  Update /*+ rule*/ 病人护理记录
  Set 待转出 = n_批次
  Where (病人id, 主页id) In (Select 病人id, 主页id From 病案主页 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人护理内容
  Set 待转出 = n_批次
  Where 记录id In (Select ID From 病人护理记录 Where 待转出 = n_批次);

  Update Zldatamovelog
  Set 当前进度 = '(6/10)护理数据标记完成，正在标记病历数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  --4.病历数据 
  Update /*+ rule*/ 电子病历记录
  Set 待转出 = n_批次
  Where 病人来源 <> 4 And (病人id, 主页id) In (Select 病人id, ID
                                       From 病人挂号记录
                                       Where 待转出 = n_批次
                                       Union All
                                       Select 病人id, 主页id
                                       From 病案主页
                                       Where 待转出 = n_批次);

  --自登记类病人(无挂号单号) 
  --病历ID可能重复是因为检验报告之类的，如肝功、肾功共打一张报告，即在病人医嘱报告表中，多个医嘱id对应同一报告ID 
  --为提升性能，不从医嘱发送记录的发送时间查询，不采用精确的时间，因为直接登记的检验医嘱，一般开嘱时间与发送时间相差不大
  Update /*+ rule*/ 电子病历记录
  Set 待转出 = N_批次
  Where ID In (Select C.病历id
             From 病人医嘱记录 B, 病人医嘱报告 C
             Where C.医嘱id = B.Id And Nvl(B.主页id, 0) = 0 And B.挂号单 Is Null And B.相关id Is Null And B.待转出 Is Null And
                   B.开嘱时间 < d_End);

  Update /*+ rule*/ 电子病历附件
  Set 待转出 = n_批次
  Where 病历id In (Select ID From 电子病历记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 电子病历格式
  Set 待转出 = n_批次
  Where 文件id In (Select ID From 电子病历记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 电子病历内容
  Set 待转出 = n_批次
  Where 文件id In (Select ID From 电子病历记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 电子病历图形
  Set 待转出 = n_批次
  Where 对象id In (Select ID From 电子病历内容 Where 待转出 = n_批次 And 对象类型 = 5);

  Update /*+ rule*/ 病人医嘱报告
  Set 待转出 = n_批次
  Where 病历id In (Select ID From 电子病历记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 影像报告驳回
  Set 待转出 = n_批次
  Where (医嘱id, 病历id) In (Select 医嘱id, 病历id From 病人医嘱报告 Where 待转出 = n_批次);

  Update /*+ rule*/ 报告查阅记录
  Set 待转出 = n_批次
  Where 病历id In (Select ID From 电子病历记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 疾病申报记录
  Set 待转出 = n_批次
  Where 文件id In (Select ID From 电子病历记录 Where 待转出 = n_批次);
  
  Update /*+ rule*/ 疾病报告反馈
  Set 待转出 = n_批次
  Where 文件id In (Select ID From 电子病历记录 Where 待转出 = n_批次);
  
  Update /*+ rule*/ 病人诊断记录
  Set 待转出 = n_批次
  Where 病历id In (Select ID From 电子病历记录 Where 待转出 = n_批次);

  Update Zldatamovelog
  Set 当前进度 = '(7/10)病历数据标记完成，正在标记临床路径数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  --5.临床路径 
  Update /*+ rule*/ 病人临床路径
  Set 待转出 = n_批次
  Where (病人id, 主页id) In (Select 病人id, 主页id From 病案主页 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人合并路径
  Set 待转出 = n_批次
  Where 首要路径记录id In (Select ID From 病人临床路径 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人合并路径评估
  Set 待转出 = n_批次
  Where 路径记录id In (Select ID From 病人临床路径 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人出径记录
  Set 待转出 = n_批次
  Where 路径记录id In (Select ID From 病人临床路径 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人路径执行
  Set 待转出 = n_批次
  Where 路径记录id In (Select ID From 病人临床路径 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人路径评估
  Set 待转出 = n_批次
  Where 路径记录id In (Select ID From 病人临床路径 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人路径变异
  Set 待转出 = n_批次
  Where 路径记录id In (Select ID From 病人临床路径 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人路径指标
  Set 待转出 = n_批次
  Where 路径记录id In (Select ID From 病人临床路径 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人路径医嘱
  Set 待转出 = n_批次
  Where 路径执行id In (Select ID From 病人路径执行 Where 待转出 = n_批次);

  Update Zldatamovelog
  Set 当前进度 = '(8/10)临床路径数据标记完成，正在标记医嘱数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  --6.医嘱，检验，检查 
  Update /*+ rule*/ 病人医嘱记录
  Set 待转出 = n_批次
  Where 挂号单 In (Select NO From 病人挂号记录 Where 待转出 = n_批次) And 病人来源 <> 4;
  Update /*+ rule*/ 病人医嘱记录
  Set 待转出 = n_批次
  Where (病人id, 主页id) In (Select 病人id, 主页id From 病案主页 Where 待转出 = n_批次);

  --自登记类病人(无挂号单)，病人医嘱报告在前面转病历时已转出 
  Update /*+ rule*/ 病人医嘱记录
  Set 待转出 = n_批次
  Where Rowid In (Select b.Rowid
                  From 病人医嘱记录 B, 病人医嘱报告 C
                  Where (b.相关id = c.医嘱id Or b.Id = c.医嘱id) And c.待转出 = n_批次);

  Update /*+ rule*/ 病人医嘱计价
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人医嘱附费
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人医嘱附件
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 输血申请记录
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 输血检验结果
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人医嘱执行
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人医嘱打印
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 医嘱执行打印
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人诊断医嘱
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 病人诊断记录
  Set 待转出 = n_批次
  Where ID In (Select 诊断id From 病人诊断医嘱 Where 待转出 = n_批次);

  Update /*+ rule*/ 病人医嘱状态
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 医嘱签名记录
  Set 待转出 = n_批次
  Where ID In (Select 签名id From 病人医嘱状态 Where 待转出 = n_批次 And 签名id Is Not Null);

  Update /*+ rule*/ 病人医嘱发送
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 诊疗单据打印
  Set 待转出 = n_批次
  Where (NO, 记录性质) In (Select NO, 记录性质 From 病人医嘱发送 Where 待转出 = n_批次);

  Update /*+ rule*/ 医嘱执行时间
  Set 待转出 = n_批次
  Where (医嘱id, 发送号) In (Select 医嘱id, 发送号 From 病人医嘱发送 Where 待转出 = n_批次);

  Update /*+ rule*/ 医嘱执行计价
  Set 待转出 = n_批次
  Where (医嘱id, 发送号) In (Select 医嘱id, 发送号 From 病人医嘱发送 Where 待转出 = n_批次);

  Update /*+ rule*/ 执行打印记录
  Set 待转出 = n_批次
  Where (医嘱id, 发送号) In (Select 医嘱id, 发送号 From 病人医嘱发送 Where 待转出 = n_批次);

  Update /*+ rule*/ 处方审查明细
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 处方审查记录
  Set 待转出 = n_批次
  Where ID In (Select 审方id From 处方审查明细 Where 待转出 = n_批次);

  Update /*+ rule*/ 处方审查结果
  Set 待转出 = n_批次
  Where 审方id In (Select ID From 处方审查记录 Where 待转出 = n_批次);
  
  Update /*+ rule*/ RIS检查预约
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  
  Update Zldatamovelog
  Set 当前进度 = '(9/10)医嘱数据标记完成，正在标记检查检验数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  Update /*+ rule*/ 影像检查记录
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 影像报告记录
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 影像报告操作记录
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 影像检查序列
  Set 待转出 = n_批次
  Where 检查uid In (Select 检查uid From 影像检查记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 影像检查图象
  Set 待转出 = n_批次
  Where 序列uid In (Select 序列uid From 影像检查序列 Where 待转出 = n_批次);

  Update /*+ rule*/ 影像申请单图像
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 影像收藏内容
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 影像危急值记录
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update Zldatamovelog
  Set 当前进度 = '(10/10)影像数据标记完成，正在标记检验数据'
  Where 系统 = n_System And 批次 = n_批次;
  Commit;

  Update /*+ rule*/ 检验标本记录
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 检验申请项目
  Set 待转出 = n_批次
  Where 标本id In (Select ID From 检验标本记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 检验项目分布
  Set 待转出 = n_批次
  Where 标本id In (Select ID From 检验标本记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 检验分析记录
  Set 待转出 = n_批次
  Where 标本id In (Select ID From 检验标本记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 检验质控记录
  Set 待转出 = n_批次
  Where 标本id In (Select ID From 检验标本记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 检验操作记录
  Set 待转出 = n_批次
  Where 标本id In (Select ID From 检验标本记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 检验签名记录
  Set 待转出 = n_批次
  Where 检验标本id In (Select ID From 检验标本记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 检验图像结果
  Set 待转出 = n_批次
  Where 标本id In (Select ID From 检验标本记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 检验试剂记录
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 检验拒收记录
  Set 待转出 = n_批次
  Where 医嘱id In (Select ID From 病人医嘱记录 Where 待转出 = n_批次);

  Update /*+ rule*/ 检验普通结果
  Set 待转出 = n_批次
  Where 检验标本id In (Select ID From 检验标本记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 检验质控报告
  Set 待转出 = n_批次
  Where 结果id In (Select ID From 检验普通结果 Where 待转出 = n_批次);
  Update /*+ rule*/ 检验药敏结果
  Set 待转出 = n_批次
  Where 细菌结果id In (Select ID From 检验普通结果 Where 待转出 = n_批次);
  Update /*+ rule*/ 检验流水线标本
  Set 待转出 = n_批次
  Where 标本id In (Select ID From 检验标本记录 Where 待转出 = n_批次);
  Update /*+ rule*/ 检验流水线指标
  Set 待转出 = n_批次
  Where 标本id In (Select ID From 检验标本记录 Where 待转出 = n_批次);

  Commit;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl1_Datamove_Tag;
/
