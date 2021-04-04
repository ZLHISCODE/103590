----------------------------------------------------------------------------------------------------------------
--本脚本支持从ZLHIS+ v10.35.90升级到 v10.35.90
--请以数据空间的所有者登录PLSQL并执行下列脚本
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------
--结构修正部份
------------------------------------------------------------------------------
--137889:殷瑞,2019-03-18,新增静配对自备药的处理
create table 输液自备药清单
(
  序号     number(3) not null,
  药品id   number(18) not null,
  是否检查库存 number(1)
)
tablespace zl9BaseItem;
alter table 输液自备药清单 add constraint 输液自备药清单_PK_序号 primary key (序号);
alter table 输液自备药清单 add constraint 输液自备药清单_UQ_药品id unique (药品ID);


------------------------------------------------------------------------------
--数据修正部份
------------------------------------------------------------------------------
--137889:殷瑞,2019-03-18,新增静配对自备药的处理
Insert into zlTables(系统,表名,表空间,分类) Values(100,'输液自备药清单','ZL9BASEITEM','A2');

-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------
--137889:殷瑞,2019-03-18,新增静配对自备药的处理
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1022,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0
Union All Select 'Zl_输液自备药清单_设置','EXECUTE' From Dual
Union All Select '输液自备药清单','SELECT' From Dual) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1345,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0
Union All Select '输液自备药清单','SELECT' From Dual) A;

--137893:胡俊勇,2019-03-18,输液自备药清单
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1254,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0
Union All Select '输液自备药清单','SELECT' From Dual) A;


-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------

--137889:殷瑞,2019-03-18,新增静配对自备药的处理
Create Or Replace Procedure Zl_输液自备药清单_设置
(
  序号_In         In 输液自备药清单.序号%Type,
  药品id_In       In 输液自备药清单.药品id%Type,
  是否检查库存_In In 输液自备药清单.是否检查库存%Type,
  n_First_In      In Number
) Is
Begin
  --新增前先删除之前的
  If n_First_In = 1 Then
    Delete From 输液自备药清单;
  End If;

  --插入[输液自备药清单]数据
  If 药品id_In <> 0 Then
    Insert Into 输液自备药清单 (序号, 药品id, 是否检查库存) Values (序号_In, 药品id_In, 是否检查库存_In);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_输液自备药清单_设置;
/

--137889:殷瑞,2019-03-18,新增静配对自备药的处理
Create Or Replace Procedure Zl_输液配药记录_销帐审核
(
  配药id_In   In Varchar2, --ID串:ID1,审核标志1,ID2,审核标志2....
  操作人员_In In 输液配药记录.操作人员%Type,
  操作时间_In In 输液配药记录.操作时间%Type
) Is
  v_Tansid     Varchar2(20);
  v_Tansids    Varchar2(4000);
  v_Tmp        Varchar2(4000);
  v_Usercode   Varchar2(100);
  v_发药id     药品收发记录.Id%Type;
  n_Count      Number(1);
  d_审核时间   药品收发记录.审核日期%Type;
  v_No         药品收发记录.No%Type;
  v_上次no     药品收发记录.No%Type;
  n_审核标志   Number(1);
  n_操作状态   Number(2);
  v_收发ids    Varchar2(4000);
  v_退药待发id 药品收发记录.Id%Type;

  v_原始id     药品收发记录.Id%Type;
  v_Error      Varchar2(255);
  n_门诊         Number; --1：门诊单据；2：住院单据
  Err_Custom Exception;

  Cursor c_销帐记录 Is
    Select Distinct a.费用id, b.操作时间
    From 药品收发记录 A, 输液配药记录 B, 输液配药内容 C
    Where a.Id = c.收发id And b.Id = c.记录id And b.Id = v_Tansid And b.操作状态 = 9;

  v_销帐记录 c_销帐记录%RowType;

  Cursor c_退药记录 Is
    Select /*+ rule*/
    Distinct a.Id As 退药id, c.收发id, c.数量, a.药品id, a.批次, c.记录id As 配药id
    From 药品收发记录 A, 药品收发记录 B, 输液配药内容 C, Table(Cast(f_Num2list(v_Tansids) As Zltools.t_Numlist)) D
    Where c.记录id = d.Column_Value And b.Id = c.收发id And a.单据 = b.单据 And a.No = b.No And a.库房id + 0 = b.库房id And
          a.药品id + 0 = b.药品id And a.序号 = b.序号 And (a.记录状态 = 1 Or Mod(a.记录状态, 3) = 0)
    Order By a.药品id, a.批次;

  v_退药记录 c_退药记录%RowType;

  Cursor c_费用销帐 Is
    Select /*+ rule*/
     a.No, a.序号 || ':' || c.数量 || ':' || c.记录id As 费用序号
    From 住院费用记录 A, 药品收发记录 B, 输液配药内容 C, Table(Cast(f_Num2list(v_Tansids) As Zltools.t_Numlist)) D
    Where a.Id = b.费用id And b.Id = c.收发id And Mod(b.记录状态, 3) = 1 And c.记录id = d.Column_Value;

  v_费用销帐 c_费用销帐%RowType;

  Cursor c_自备药记录 Is
    Select Distinct a.Id, b.单次用量, c.剂量系数, c.药品id
    From 输液配药记录 A, 病人医嘱记录 B, 药品规格 C, Table(Cast(f_Num2list(v_Tansids) As Zltools.t_Numlist)) D
    Where a.医嘱id = b.相关id And a.Id = d.Column_Value And b.收费细目id = c.药品id And b.执行性质 = 5 And b.执行标记 = 0 And
          b.收费细目id In (Select d.药品id From 输液自备药清单 D Where d.是否检查库存 = 1);

  v_自备药记录 c_自备药记录%RowType;
Begin
  If 配药id_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := 配药id_In || ',';
  End If;

  v_Usercode := Zl_Identity;
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ';') + 1);
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ',') + 1);
  v_Usercode := Substr(v_Usercode, 1, Instr(v_Usercode, ',') - 1);

  While v_Tmp Is Not Null Loop
    --分解单据ID串
    v_Tansid   := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp      := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
    n_审核标志 := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp      := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
  
    v_收发ids := Null;
  
    --统计审核确认的输液单(n_审核标志 = 1)
    If n_审核标志 = 1 Then
      If v_Tansids Is Null Then
        v_Tansids := v_Tansid;
      Else
        v_Tansids := v_Tansids || ',' || v_Tansid;
      End If;
    End If;
  
    --检查当前输液单的状态是否为待摆药状态
    Begin
      Select 操作状态 Into n_操作状态 From 输液配药记录 Where ID = v_Tansid;
    
      If n_操作状态 <> 9 Then
        v_Error := '该数据已被操作，不能进行销帐审核！';
        Raise Err_Custom;
      End If;
    Exception
      When Others Then
        v_Error := '该数据已被操作！';
        Raise Err_Custom;
    End;
  
    If n_审核标志 = 1 Then
      n_操作状态 := 10;
    Elsif n_审核标志 = 2 Then
      n_操作状态 := 11;
    End If;
  
    --查找输液单对应的收发NO
    Begin
      Select NO
      Into v_No
      From 药品收发记录
      Where ID In (Select 收发id From 输液配药内容 Where 记录id In (Select ID From 输液配药记录 Where ID = v_Tansid)) And Rownum < 2;
    Exception
      When Others Then
        v_No := '';
    End;
  
    --收发NO相同的配药ID，审核时间以此设置为延长1秒
    If v_No = v_上次no Then
      d_审核时间 := d_审核时间 + 1 / 24 / 60 / 60;
    Else
      d_审核时间 := 操作时间_In;
      v_上次no   := v_No;
    End If;
  
    --销帐记录处理
    For v_销帐记录 In c_销帐记录 Loop
      Zl_病人费用销帐_Audit(v_销帐记录.费用id, v_销帐记录.操作时间, 操作人员_In, d_审核时间, n_审核标志);
    End Loop;
  
    Select Count(*) Into n_Count From 输液配药状态 Where 配药id = v_Tansid And 操作时间 = 操作时间_In;
  
    If n_Count <> 1 Then
      Insert Into 输液配药状态
        (配药id, 操作类型, 操作人员, 操作时间)
      Values
        (v_Tansid, n_操作状态, 操作人员_In, 操作时间_In);
    End If;
    Update 输液配药记录 Set 操作人员 = 操作人员_In, 操作时间 = 操作时间_In, 操作状态 = n_操作状态 Where ID = v_Tansid;
  End Loop;

  --先退药
  For v_退药记录 In c_退药记录 Loop
    Zl_药品收发记录_部门退药(v_退药记录.退药id, 操作人员_In, 操作时间_In, Null, Null, Null, v_退药记录.数量, Null, 操作人员_In);
  
    --取退药待发id
    Select a.Id
    Into v_发药id
    From 药品收发记录 A, 药品收发记录 B
    Where b.Id = v_退药记录.退药id And a.单据 = b.单据 And a.No = b.No And a.库房id + 0 = b.库房id And a.药品id + 0 = b.药品id And
          a.序号 = b.序号 And Mod(a.记录状态, 3) = 1 And a.审核日期 Is Null;
  
    --输液配药内容中的收发ID更新为退药待发的收发ID
    Update 输液配药内容 Set 收发id = v_发药id Where 记录id = v_退药记录.配药id And 收发id = v_退药记录.收发id;
  
    If v_收发ids Is Null Then
      v_收发ids := v_发药id;
    Else
      v_收发ids := v_收发ids || ',' || v_发药id;
    End If;
  
    --取原始id
    Select a.Id
    Into v_原始id
    From 药品收发记录 A, 药品收发记录 B
    Where b.Id = v_退药记录.退药id And a.单据 = b.单据 And a.No = b.No And a.库房id + 0 = b.库房id And a.药品id + 0 = b.药品id And
          a.序号 = b.序号 And Mod(a.记录状态, 3) = 0 And a.审核日期 Is Not Null;
  
    Insert Into 输液配药内容
      (记录id, 收发id, 数量)
      Select 记录id, v_原始id, 数量 From 输液配药内容 Where 记录id = v_退药记录.配药id And 收发id = v_发药id;
  
    v_收发ids := v_收发ids || ',' || v_原始id;
  End Loop;

  --费用销帐
  For v_费用销帐 In c_费用销帐 Loop
    Zl_住院记帐记录_Delete(v_费用销帐.No, v_费用销帐.费用序号, v_Usercode, Zl_Username, 2, 1, 1, d_审核时间);
  End Loop;

  --检查表【输液自备药清单】中设置相关药品的已发药数据
  For v_自备药记录 In c_自备药记录 Loop
    --若输液单存在相关自备药,则收集【药品收发记录】中的id
    For v_自备药收发记录 In (Select a.Id, a.批号, a.效期, a.产地, a.实际数量 As 退药数, a.批次, a.费用id
                      From 药品收发记录 A
                      Where a.计划id = v_自备药记录.Id And a.药品id = v_自备药记录.药品id And a.审核人 Is Not Null And a.审核日期 Is Not Null And
                            (a.记录状态 = 1 Or Mod(a.记录状态, 3) = 0)
                      Order By a.批次) Loop
    
      --判断这个单据是门诊还是住院 
      Begin
        Select 1 Into n_门诊 From 门诊费用记录 Where ID = v_自备药收发记录.费用id;
      Exception
        When Others Then
          n_门诊 := 2;
      End;
    
      Zl_药品收发记录_部门退药(v_自备药收发记录.Id, 操作人员_In, 操作时间_In, Null, Null, Null, v_自备药收发记录.退药数, Null, 操作人员_In, 2, n_门诊);
    End Loop;
  End Loop;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_输液配药记录_销帐审核;
/

--137889:殷瑞,2019-03-18,新增静配对自备药的处理
Create Or Replace Procedure Zl_输液配药记录_取消摆药(配药id_In In Varchar2 --ID串:配药ID1,配药ID2....
                                           ) Is
  v_Id       Varchar2(20);
  v_发药id   Varchar2(20);
  v_Tmp      Varchar2(4000);
  v_Date     Date;
  v_操作人员 输液配药记录.操作人员%Type;
  d_操作时间 输液配药记录.操作时间%Type;
  n_操作状态 输液配药记录.操作状态%Type;

  v_停止配药ids    Varchar2(4000); --因自备药退药时数量核对不上，故取消相关输液单的【取消摆药】操作
  n_自备药数量     药品收发记录.实际数量%Type;
  n_自备药汇总数量 药品收发记录.实际数量%Type; --该自备药在药品收发记录中可以被发的总数量
  n_门诊           Number; --1：门诊单据；2：住院单据

  v_Error Varchar2(255);
  Err_Custom Exception;

  Cursor c_配药内容 Is
    Select /*+ rule*/
    Distinct c.记录id, a.Id As 退药id, c.收发id, a.批号, a.效期, a.产地, c.数量 As 退药数, a.药品id, a.批次
    From 药品收发记录 A, 药品收发记录 B, 输液配药内容 C, Table(Cast(f_Num2list(配药id_In) As Zltools.t_Numlist)) D
    Where c.记录id = d.Column_Value And b.Id = c.收发id And a.单据 = b.单据 And a.No = b.No And a.库房id + 0 = b.库房id And
          a.药品id + 0 = b.药品id And a.序号 = b.序号 And (a.记录状态 = 1 Or Mod(a.记录状态, 3) = 0)
    Order By a.药品id, a.批次;

  v_配药内容 c_配药内容%RowType;

  Cursor c_自备药记录 Is
    Select Distinct a.Id, b.单次用量, c.剂量系数, c.药品id
    From 输液配药记录 A, 病人医嘱记录 B, 药品规格 C, Table(Cast(f_Num2list(配药id_In) As Zltools.t_Numlist)) D
    Where a.医嘱id = b.相关id And a.Id = d.Column_Value And b.收费细目id = c.药品id And b.执行性质 = 5 And b.执行标记 = 0 And
          b.收费细目id In (Select d.药品id From 输液自备药清单 D Where d.是否检查库存 = 1);

  v_自备药记录 c_自备药记录%RowType;
Begin
  If 配药id_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := 配药id_In || ',';
  End If;
  While v_Tmp Is Not Null Loop
    --分解单据ID串
    v_Id  := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp := Replace(',' || v_Tmp, ',' || v_Id || ',');
  
    --检查当前输液单的状态是否为待摆药状态
    Begin
      Select 操作状态 Into n_操作状态 From 输液配药记录 Where ID = v_Id;
    
      If n_操作状态 != 2 Then
        v_Error := '该数据已被操作，不能进行取消摆药操作！';
        Raise Err_Custom;
      End If;
    Exception
      When Others Then
        v_Error := '该数据已被操作！';
        Raise Err_Custom;
    End;
  
    Select 操作人员, 操作时间
    Into v_操作人员, d_操作时间
    From 输液配药状态
    Where 配药id = v_Id And 操作类型 = 1 And Rownum = 1;
  
    Update 输液配药记录 Set 操作状态 = 1, 操作人员 = v_操作人员, 操作时间 = d_操作时间 Where ID = v_Id;
  
    --向[输液配药状态]表中记录“取消摆药”的操作
    Insert Into 输液配药状态
      (配药id, 操作类型, 操作人员, 操作时间, 操作说明)
    Values
      (v_Id, 1, v_操作人员, Sysdate, '取消摆药');
  
  End Loop;

  Select Sysdate Into v_Date From Dual;

  --检查表【输液自备药清单】中设置相关药品的已发药数据
  For v_自备药记录 In c_自备药记录 Loop
  
    n_自备药数量 := v_自备药记录.单次用量 / v_自备药记录.剂量系数;
  
    Select Sum(a.实际数量)
    Into n_自备药汇总数量
    From 药品收发记录 A
    Where Mod(a.记录状态, 3) = 1 And a.计划id = v_自备药记录.Id And a.药品id = v_自备药记录.药品id And a.审核人 Is Not Null And
          a.审核日期 Is Not Null;
  
    If n_自备药汇总数量 < n_自备药数量 Then
      --如果数量核对不上，则收集当前配药id,并在下面同步处理该输液单的对应药品
      If v_停止配药ids Is Null Then
        v_停止配药ids := v_自备药记录.Id;
      Else
        v_停止配药ids := v_停止配药ids || ',' || v_自备药记录.Id;
      End If;
    
      Exit;
    
    End If;
  
    --若输液单存在相关自备药,则收集【药品收发记录】中的id
    For v_自备药收发记录 In (Select a.Id, a.批号, a.效期, a.产地, a.实际数量 As 退药数, a.批次, a.费用id
                      From 药品收发记录 A
                      Where a.计划id = v_自备药记录.Id And a.药品id = v_自备药记录.药品id And a.审核人 Is Not Null And a.审核日期 Is Not Null And
                            (a.记录状态 = 1 Or Mod(a.记录状态, 3) = 0)
                      Order By a.批次) Loop
    
      --判断这个单据是门诊还是住院 
      Begin
        Select 1 Into n_门诊 From 门诊费用记录 Where ID = v_自备药收发记录.费用id;
      Exception
        When Others Then
          n_门诊 := 2;
      End;

      Zl_药品收发记录_部门退药(v_自备药收发记录.Id, Zl_Username, v_Date, v_自备药收发记录.批号, v_自备药收发记录.效期, v_自备药收发记录.产地, v_自备药收发记录.退药数, Null,
                     Zl_Username, 2, n_门诊);
    End Loop;
  End Loop;

  For v_配药内容 In c_配药内容 Loop
    --排除被中断的输液单
    If Instr(',' || v_停止配药ids || ',', ',' || v_配药内容.记录id || ',') = 0 Then
      --处理退药
      Zl_药品收发记录_部门退药(v_配药内容.退药id, Zl_Username, v_Date, v_配药内容.批号, v_配药内容.效期, v_配药内容.产地, v_配药内容.退药数, Null, Zl_Username);
    
      Select Max(a.Id)
      Into v_发药id
      From 药品收发记录 A, 药品收发记录 B
      Where b.Id = v_配药内容.退药id And a.单据 = b.单据 And a.No = b.No And a.库房id + 0 = b.库房id And a.药品id + 0 = b.药品id And
            a.序号 = b.序号 And Mod(a.记录状态, 3) = 1 And a.审核日期 Is Null;
    
      --替换输液配药内容中的收发ID
      Update 输液配药内容 Set 收发id = v_发药id Where 记录id = v_配药内容.记录id And 收发id = v_配药内容.收发id;
    End If;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_输液配药记录_取消摆药;
/

--137889:殷瑞,2019-03-19,修正逻辑错误
Create Or Replace Procedure Zl_输液配药记录_摆药
(
  部门id_In   In 输液配药记录.部门id%Type,
  配药id_In   In Varchar2, --ID串:ID1,ID2....
  摆药单号_In In 输液配药记录.摆药单号%Type,
  操作人员_In In 输液配药状态.操作人员%Type := Null,
  操作时间_In In 输液配药状态.操作时间%Type := Null,
  移动操作_In In Number := 0
) Is
  v_Tansid Varchar2(20);
  v_Tmp    Varchar2(4000);

  v_收发ids      Varchar2(4000);
  v_Error        Varchar2(255);
  n_是否打包     输液配药记录.是否打包%Type;
  n_操作状态     输液配药记录.操作状态%Type;
  v_摆药人       Varchar2(20);
  v_配药台       Varchar2(20);
  n_配药台id     Number(4);
  n_部门id       Number(18);
  n_批次         Number(2);
  d_日期         Date;
  n_流通金额小数 Number;
  n_流通单价小数 Number;

  v_自备药收发ids      Varchar2(4000);
  n_自备药数量         药品收发记录.实际数量%Type;
  n_自备药汇总数量     药品收发记录.实际数量%Type; --该自备药在药品收发记录中可以被发的总数量
  n_自备药执行汇总数量 药品收发记录.实际数量%Type; --该自备药退药待发的汇总数量
  n_自备药已收集数量   药品收发记录.实际数量%Type; --根据收发id，统计当前已准备的数量
  n_自备药未收集数量   药品收发记录.实际数量%Type; --根据收发id，统计当前未准备的数量
  n_自备药序号         药品收发记录.序号%Type; --根据收发id，统计当前未准备的数量
  n_收发id             药品收发记录.Id%Type;

  Err_Custom Exception;
  Cursor c_收发记录 Is
    Select /*+ rule*/
     a.Id, Nvl(a.批次, 0) As 批次
    From 药品收发记录 A,
         (Select Distinct 收发id
           From 输液配药内容 A, Table(Cast(f_Num2list(配药id_In) As Zltools.t_Numlist)) B
           Where a.记录id = b.Column_Value) B
    Where a.Id = b.收发id And a.审核人 Is Null
    Order By a.药品id, a.批次;

  v_收发记录 c_收发记录%RowType;
Begin
  If 配药id_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := 配药id_In || ',';
  End If;

  v_自备药收发ids := Null;
  Select 精度 Into n_流通单价小数 From 药品卫材精度 Where 类别 = 1 And 内容 = 2 And 单位 = 1;
  Select 精度 Into n_流通金额小数 From 药品卫材精度 Where 类别 = 1 And 内容 = 4 And 单位 = 5;

  While v_Tmp Is Not Null Loop
    --分解单据ID串
    v_Tansid := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp    := Replace(',' || v_Tmp, ',' || v_Tansid || ',');
  
    --检查当前输液单的状态是否为待摆药状态
    Begin
      Select 操作状态 Into n_操作状态 From 输液配药记录 Where ID = v_Tansid;
    
      If n_操作状态 > 1 Then
        v_Error := '该数据已被操作，不能进行发药！';
        Raise Err_Custom;
      End If;
    Exception
      When Others Then
        v_Error := '该数据已被操作！';
        Raise Err_Custom;
    End;
  
    Begin
      Select 是否打包 Into n_是否打包 From 输液配药记录 Where ID = v_Tansid For Update Nowait;
    Exception
      When Others Then
        v_Error := '已有其他用户在执行发药，不能重复操作！';
        Raise Err_Custom;
    End;
  
    --检查表【输液自备药清单】中设置相关药品的待发药数据
    For v_自备药记录 In (Select a.Id, b.收费细目id As 药品id, b.单次用量, c.剂量系数, a.部门id, b.病人id, b.总给予量, b.标本部位 As 药品品种
                    From 输液配药记录 A, 病人医嘱记录 B, 药品规格 C
                    Where a.医嘱id = b.相关id And b.收费细目id = c.药品id And a.Id = v_Tansid And b.执行性质 = 5 And b.执行标记 = 0 And
                          b.收费细目id In (Select d.药品id From 输液自备药清单 D Where d.是否检查库存 = 1)) Loop
    
      n_自备药数量 := v_自备药记录.单次用量 / v_自备药记录.剂量系数;
    
      --检查是否存在已执行过的待发药单据
      Select Nvl(Sum(b.实际数量), 0)
      Into n_自备药执行汇总数量
      From 未发药品记录 A, 药品收发记录 B
      Where a.单据 = b.单据 And a.No = b.No And a.库房id = b.库房id And Mod(b.记录状态, 3) = 1 And b.记录状态 <> 1 And
            a.病人id = v_自备药记录.病人id And a.库房id = v_自备药记录.部门id And b.审核人 Is Null And b.审核日期 Is Null And
            b.药品id = v_自备药记录.药品id And b.计划id = v_自备药记录.Id;
    
      If n_自备药执行汇总数量 > 0 And n_自备药执行汇总数量 < n_自备药数量 Then
        v_Error := '药品【' || v_自备药记录.药品品种 || '】的数量不足，不能进行发药！';
        Raise Err_Custom;
      Elsif n_自备药执行汇总数量 = n_自备药数量 Then
        --收集已执行过的配药id记录
        For v_自备药已执行记录 In (Select b.Id As 收发id, b.实际数量, b.批次, b.记录状态, b.单据, b.No, b.库房id
                           From 未发药品记录 A, 药品收发记录 B
                           Where a.单据 = b.单据 And a.No = b.No And a.库房id = b.库房id And Mod(b.记录状态, 3) = 1 And b.记录状态 <> 1 And
                                 a.病人id = v_自备药记录.病人id And a.库房id = v_自备药记录.部门id And b.药品id = v_自备药记录.药品id And
                                 b.审核日期 Is Null And b.审核人 Is Null And b.计划id = v_自备药记录.Id
                           Order By b.批次) Loop
          If v_自备药收发ids Is Null Then
            v_自备药收发ids := v_自备药已执行记录.收发id || ',' || v_自备药已执行记录.批次;
          Else
            v_自备药收发ids := v_自备药收发ids || '|' || v_自备药已执行记录.收发id || ',' || v_自备药已执行记录.批次;
          End If;
        End Loop;
      Elsif n_自备药执行汇总数量 = 0 Then
        --检查对应药品各批次的总和是否满足本次发药数量
        Select Nvl(Sum(b.实际数量), 0)
        Into n_自备药汇总数量
        From 未发药品记录 A, 药品收发记录 B
        Where a.单据 = b.单据 And a.No = b.No And a.库房id = b.库房id And Mod(b.记录状态, 3) = 1 And a.病人id = v_自备药记录.病人id And
              a.库房id = v_自备药记录.部门id And b.审核人 Is Null And b.审核日期 Is Null And b.药品id = v_自备药记录.药品id And b.计划id Is Null And
              Exists (Select 1 From 门诊费用记录 C Where c.Id = b.费用id);
      
        If n_自备药汇总数量 < n_自备药数量 Then
          v_Error := '药品【' || v_自备药记录.药品品种 || '】的数量不足，不能进行发药！';
          Raise Err_Custom;
        End If;
      
        --循环拆分并收集可执行收发ID串,格式为:"id1,批次1|id2,批次2|....."
        n_自备药已收集数量 := 0;
        n_自备药未收集数量 := 0;
      
        For v_自备药收发记录 In (Select b.Id As 收发id, b.实际数量, b.批次, b.记录状态, b.单据, b.No, b.库房id
                          From 未发药品记录 A, 药品收发记录 B
                          Where a.单据 = b.单据 And a.No = b.No And a.库房id = b.库房id And Mod(b.记录状态, 3) = 1 And
                                a.病人id = v_自备药记录.病人id And a.库房id = v_自备药记录.部门id And b.药品id = v_自备药记录.药品id And
                                b.审核日期 Is Null And b.审核人 Is Null And b.计划id Is Null And Exists
                           (Select 1 From 门诊费用记录 C Where c.Id = b.费用id)
                          Order By b.批次) Loop
        
          n_自备药未收集数量 := n_自备药数量 - n_自备药已收集数量;
          n_自备药已收集数量 := n_自备药已收集数量 + v_自备药收发记录.实际数量;
        
          If n_自备药已收集数量 < n_自备药数量 Then
            --直接收集当前收发记录
            If v_自备药收发ids Is Null Then
              v_自备药收发ids := v_自备药收发记录.收发id || ',' || v_自备药收发记录.批次;
            Else
              v_自备药收发ids := v_自备药收发ids || '|' || v_自备药收发记录.收发id || ',' || v_自备药收发记录.批次;
            End If;
          
            Update 药品收发记录 Set 计划id = v_Tansid Where ID = v_自备药收发记录.收发id;
          
          Elsif n_自备药已收集数量 > n_自备药数量 Then
            --需要拆分，并收集相关收发记录
          
            Select 药品收发记录_Id.Nextval Into n_收发id From Dual;
          
            Select Max(a.序号)
            Into n_自备药序号
            From 药品收发记录 A
            Where a.单据 = v_自备药收发记录.单据 And a.No = v_自备药收发记录.No And a.库房id = v_自备药收发记录.库房id;
          
            Update 药品收发记录
            Set 填写数量 = 填写数量 - n_自备药未收集数量, 实际数量 = 实际数量 - n_自备药未收集数量, 零售金额 = 零售价 * (实际数量 - n_自备药未收集数量)
            Where ID = v_自备药收发记录.收发id;
          
            Insert Into 药品收发记录
              (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 付数, 填写数量, 实际数量, 成本价, 成本金额, 扣率,
               零售价, 零售金额, 差价, 摘要, 填制人, 填制日期, 配药人, 审核人, 审核日期, 费用id, 单量, 频次, 用法, 发药窗口, 供药单位id, 生产日期, 批准文号, 注册证号, 计划id, 原产地,
               紧急标志)
              Select n_收发id, 记录状态, 单据, NO, n_自备药序号 + 1, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 付数, n_自备药未收集数量,
                     n_自备药未收集数量, 成本价, 成本金额, 扣率, 零售价, Round(零售价 * n_自备药未收集数量, n_流通金额小数), 差价, 摘要, 填制人, 填制日期, 配药人, 审核人,
                     审核日期, 费用id, 单量, 频次, 用法, 发药窗口, 供药单位id, 生产日期, 批准文号, 注册证号, v_Tansid, 原产地, 紧急标志




              
              From 药品收发记录
              Where ID = v_自备药收发记录.收发id;
          
            If v_自备药收发ids Is Null Then
              v_自备药收发ids := n_收发id || ',' || v_自备药收发记录.批次;
            Else
              v_自备药收发ids := v_自备药收发ids || '|' || n_收发id || ',' || v_自备药收发记录.批次;
            End If;
          
            --跳出循环
            Exit;
          
          Else
            --收集完成
            If v_自备药收发ids Is Null Then
              v_自备药收发ids := v_自备药收发记录.收发id || ',' || v_自备药收发记录.批次;
            Else
              v_自备药收发ids := v_自备药收发ids || '|' || v_自备药收发记录.收发id || ',' || v_自备药收发记录.批次;
            End If;
          
            Update 药品收发记录 Set 计划id = v_Tansid Where ID = v_自备药收发记录.收发id;
          
            --跳出循环
            Exit;
          
          End If;
        End Loop;
      End If;
    End Loop;
  
    v_配药台   := '';
    n_配药台id := 0;
    n_部门id   := 0;
    v_摆药人   := '';
    Begin
      Select 名称, ID, 部门id, 配药批次, 执行时间
      Into v_配药台, n_配药台id, n_部门id, n_批次, d_日期
      From (Select f.名称, f.Id, a.部门id, a.配药批次, a.执行时间
             From 输液配药记录 A, 输液配药内容 B, 药品收发记录 C, 配液台药品对照 D, 配液台 F
             Where a.Id = b.记录id And b.收发id = c.Id And c.药品id = d.药品id And d.配药台id = f.Id And c.库房id = d.部门id And
                   a.Id = v_Tansid
             Order By d.配药台id)
      Where Rownum = 1;
    
      Select 摆药人
      Into v_摆药人
      From 配液工作安排
      Where 部门id = n_部门id And 配药台id = n_配药台id And 批次 = n_批次 And
            日期 = To_Date(To_Char(Sysdate, 'yyyy-mm-dd'), 'yyyy-mm-dd');
    Exception
      When Others Then
        Null;
    End;
  
    Update 输液配药记录
    Set 操作状态 = 2, 操作人员 = 操作人员_In, 操作时间 = 操作时间_In, 摆药单号 = 摆药单号_In, 配药台 = v_配药台
    Where ID = v_Tansid;
  
    Insert Into 输液配药状态
      (配药id, 操作类型, 操作人员, 操作时间, 实际工作人员)
    Values
      (v_Tansid, 2, 操作人员_In, 操作时间_In, v_摆药人);
    If n_是否打包 <> 0 And 移动操作_In = 0 Then
      Update 输液配药记录 Set 操作状态 = 4, 操作人员 = 操作人员_In, 操作时间 = 操作时间_In Where ID = v_Tansid;
      Insert Into 输液配药状态 (配药id, 操作类型, 操作人员, 操作时间) Values (v_Tansid, 4, 操作人员_In, 操作时间_In);
    End If;
  End Loop;

  For v_收发记录 In c_收发记录 Loop
    If v_收发ids Is Null Then
      v_收发ids := v_收发记录.Id || ',' || v_收发记录.批次;
    Else
      If Length(v_收发ids || '|' || v_收发记录.Id || ',' || v_收发记录.批次) > 3950 Then
        Zl_药品收发记录_批量发药(v_收发ids, 部门id_In, 操作人员_In, 操作时间_In, 4, 操作人员_In, 摆药单号_In);
        v_收发ids := v_收发记录.Id || ',' || v_收发记录.批次;
      Else
        v_收发ids := v_收发ids || '|' || v_收发记录.Id || ',' || v_收发记录.批次;
      End If;
    End If;
  End Loop;

  If Not v_收发ids Is Null Then
    Zl_药品收发记录_批量发药(v_收发ids, 部门id_In, 操作人员_In, 操作时间_In, 4, 操作人员_In, 摆药单号_In);
  End If;

  --处理自备药
  If Not v_自备药收发ids Is Null Then
    Zl_药品收发记录_批量发药(v_自备药收发ids, 部门id_In, 操作人员_In, 操作时间_In, 4, 操作人员_In, 摆药单号_In);
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_输液配药记录_摆药;
/

--137889:殷瑞,2019-03-18,新增静配对自备药的处理
Create Or Replace Procedure Zl_药品收发记录_部门退药
(
  Billid_In     In 药品收发记录.Id%Type,
  People_In     In 药品收发记录.审核人%Type,
  Date_In       In 药品收发记录.审核日期%Type,
  批号_In       In 药品库存.上次批号%Type := Null,
  效期_In       In 药品库存.效期%Type := Null,
  产地_In       In 药品库存.上次产地%Type := Null,
  退药数量_In   In 药品收发记录.实际数量%Type := Null,
  退药库房_In   In 药品收发记录.库房id%Type := Null,
  退药人_In     In 药品收发记录.领用人%Type := Null,
  Intdigit_In   In Number := 2,
  门诊_In       In Number := 2,
  汇总发药号_In In 药品收发记录.汇总发药号%Type := Null
) Is
  --只读变量
  Int记录状态   药品收发记录.记录状态%Type;
  Int执行状态   住院费用记录.执行状态%Type;
  Bln部分退药   Number;
  Lng入出类别id Number(18);
  Strno         药品收发记录.No%Type;
  Int单据       药品收发记录.单据%Type;
  Lng库房id     药品收发记录.库房id%Type;
  Lng药品id     药品收发记录.药品id%Type;
  Dbl实际数量   药品收发记录.实际数量%Type;
  Dbl实际金额   药品收发记录.零售金额%Type;
  Dbl实际成本   药品收发记录.成本金额%Type;
  Dbl实际差价   药品收发记录.差价%Type;
  Lng费用id     药品收发记录.费用id%Type;
  n_零售价      药品收发记录.零售价%Type;
  n_是否变价    Number;
  n_时价分批    Number;

  --20020731 Modified by zyb
  --处理退药时，分批核算性质改变后的处理
  Lng新批次 药品收发记录.批次%Type;
  Lng分批   药品规格.药房分批%Type;
  Lng批次   药品收发记录.批次%Type; --原批次

  Str批号        药品收发记录.批号%Type; --原批号
  Date效期       药品收发记录.效期%Type; --原效期
  n_上次供应商id 药品库存.上次供应商id%Type;
  n_上次采购价   药品库存.上次采购价%Type;
  v_上次产地     药品库存.上次产地%Type;
  v_原产地       药品库存.原产地%Type;
  d_上次生产日期 药品库存.上次生产日期%Type;
  v_批准文号     药品库存.批准文号%Type;

  n_记录性质   住院费用记录.记录性质%Type;
  v_收费类别   住院费用记录.收费类别%Type;
  n_付数       药品收发记录.付数%Type;
  n_原始数量   药品收发记录.实际数量%Type;
  v_冲销记录id 药品收发记录.Id%Type;
  Err_Custom Exception;
  v_Error    Varchar2(255);
  v_配药确认 药房配药控制.配药确认%Type;
  v_配药     药房配药控制.配药%Type;
  v_排队状态 Number(1);
  v_执行时间 药品收发记录.审核日期%Type;

Begin
  If 退药数量_In Is Not Null Then
    If 退药数量_In = 0 Then
      Return;
    End If;
  End If;

  --获取该收发记录的单据、药品ID、库房ID
  Select a.单据, a.No, a.库房id, a.药品id, a.费用id, a.入出类别id, a.记录状态, Nvl(a.批次, 0), a.批号, a.效期, a.供药单位id, a.产地, a.原产地, a.生产日期,
         a.批准文号, a.成本价, a.付数, Nvl(a.实际数量, 0) * Nvl(a.付数, 1) As 实际数量, a.零售价, Nvl(b.是否变价, 0) 是否变价
  Into Int单据, Strno, Lng库房id, Lng药品id, Lng费用id, Lng入出类别id, Int记录状态, Lng批次, Str批号, Date效期, n_上次供应商id, v_上次产地, v_原产地,
       d_上次生产日期, v_批准文号, n_上次采购价, n_付数, n_原始数量, n_零售价, n_是否变价
  From 药品收发记录 A, 收费项目目录 B
  Where a.药品id = b.Id And a.Id = Billid_In;

  Begin
    Select Nvl(配药确认, 0), Nvl(配药, 0)
    Into v_配药确认, v_配药
    From 药房配药控制
    Where 药房id = Lng库房id And Rownum = 1;
  
  Exception
    When Others Then
      v_配药确认 := 0;
      v_配药     := 0;
      Null;
  End;

  If v_配药确认 = 0 And v_配药 = 0 Then
    v_排队状态 := 2;
  Elsif v_配药确认 = 1 Then
    v_排队状态 := 0;
  Elsif v_配药 = 1 Then
    v_排队状态 := 1;
  End If;

  --获取该笔记录剩余未退数量、金额及差价
  --尽量避免金额及差价未出完的现象
  Select Sum(Nvl(实际数量, 0) * Nvl(付数, 1)), Sum(Nvl(零售金额, 0)), Sum(Nvl(成本金额, 0)), Sum(Nvl(差价, 0))
  Into Dbl实际数量, Dbl实际金额, Dbl实际成本, Dbl实际差价
  From 药品收发记录
  Where 审核人 Is Not Null And NO = Strno And 单据 = Int单据 And 序号 = (Select 序号 From 药品收发记录 Where ID = Billid_In);

  --如果允许退药数为零，表示已退药
  If Dbl实际数量 = 0 Then
    v_Error := '该单据已被其他操作员退药，请刷新后再试！';
    Raise Err_Custom;
  End If;
  If Nvl(退药数量_In, 0) > Dbl实际数量 Then
    v_Error := '该单据已被其他操作员部分退药，请刷新后再试！';
    Raise Err_Custom;
  End If;

  --获取该药品当前是否分批的信息
  Select Nvl(药房分批, 0) Into Lng分批 From 药品规格 Where 药品id = Lng药品id;
  --如果是部分退药，则重新计算零售金额及差价
  Bln部分退药 := 0;
  If Not (退药数量_In Is Null Or Nvl(退药数量_In, 0) = Dbl实际数量) Then
    Bln部分退药 := 1;
  End If;
  If Bln部分退药 = 1 Then
    Dbl实际金额 := Round(Dbl实际金额 * 退药数量_In / Dbl实际数量, Intdigit_In);
    Dbl实际成本 := Round(Dbl实际成本 * 退药数量_In / Dbl实际数量, Intdigit_In);
    Dbl实际差价 := Round(Dbl实际差价 * 退药数量_In / Dbl实际数量, Intdigit_In);
    Dbl实际数量 := 退药数量_In;
  End If;

  If n_原始数量 = 退药数量_In Then
    Dbl实际数量 := 退药数量_In / n_付数;
  Else
    n_付数 := 1;
  End If;

  --lng分批:0-不分批;1-分批;2-原分批，现不分批，按不分批处理;3-原不分批，现分批，产生新批次
  If Lng分批 = 0 And Lng批次 <> 0 Then
    --原分批，现不分批，按不分批处理
    Lng分批 := 2;
  Elsif Lng分批 <> 0 And Lng批次 = 0 Then
    --原不分批,现分批,产生新的批次，并在新产生的发药记录中使用
    Lng分批 := 3;
  Else
    If Lng批次 = 0 Then
      Lng分批 := 0;
    Else
      Lng分批 := 1;
    End If;
  End If;
  --判断是否时价分批
  If (Lng分批 = 1 Or Lng分批 = 3) And n_是否变价 = 1 Then
    n_时价分批 := 1;
  Else
    n_时价分批 := 0;
  End If;

  --记录状态的含义有所变化
  --冲销的记录状态        :iif(int记录状态=1,0,1)+1
  --被冲销的记录状态        :iif(int记录状态=1,0,1)+2
  --等待发药的记录状态    :iif(int记录状态=1,0,1)+3

  --产生冲销记录
  Select 药品收发记录_Id.Nextval Into v_冲销记录id From Dual;
  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 付数, 填写数量, 实际数量, 成本价, 成本金额, 扣率, 零售价, 零售金额,
     差价, 摘要, 填制人, 填制日期, 配药人, 审核人, 审核日期, 费用id, 单量, 频次, 用法, 发药窗口, 外观, 领用人, 供药单位id, 生产日期, 批准文号, 汇总发药号, 发药方式, 注册证号, 计划id,
     原产地)
    Select v_冲销记录id, Int记录状态 + Decode(Int记录状态, 1, 0, 1) + 1, Int单据, Strno, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地,
           批号, 效期, n_付数, -dbl实际数量, -dbl实际数量, 成本价, -dbl实际成本, 扣率, 零售价, -dbl实际金额, -dbl实际差价, 摘要, People_In, Date_In, 配药人,
           People_In, Date_In, 费用id, 单量, 频次, 用法, 发药窗口, 退药库房_In, 退药人_In, 供药单位id, 生产日期, 批准文号, 汇总发药号_In, 发药方式, 注册证号, 计划id,
           原产地
    From 药品收发记录
    Where ID = Billid_In;

  --如果是部分冲销，则付数填为1，实际数量为付数与实际数量的积
  --产生正常记录以供继续发药
  Select 药品收发记录_Id.Nextval Into Lng新批次 From Dual;
  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 付数, 填写数量, 实际数量, 成本价, 成本金额, 扣率, 零售价, 零售金额,
     差价, 摘要, 填制人, 填制日期, 配药人, 审核人, 审核日期, 费用id, 单量, 频次, 用法, 发药窗口, 供药单位id, 生产日期, 批准文号, 注册证号, 计划id, 原产地)
    Select Lng新批次, Int记录状态 + Decode(Int记录状态, 1, 0, 1) + 3, Int单据, Strno, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id,
           Decode(Lng分批, 1, 批次, 3, Lng新批次, 0), Decode(Lng分批, 3, 产地_In, 1, 产地, 产地), Decode(Lng分批, 3, 批号_In, 批号),
           Decode(Lng分批, 3, 效期_In, 效期), n_付数, Dbl实际数量, Dbl实际数量, 成本价, Dbl实际成本, 扣率, 零售价, Dbl实际金额, Dbl实际差价, 摘要, 填制人, 填制日期,
           Null, Null, Null, 费用id, 单量, 频次, 用法, 发药窗口, 供药单位id, 生产日期, 批准文号, 注册证号, 计划id, 原产地
    
    From 药品收发记录
    Where ID = Billid_In;

  Zl_未审药品记录_Insert(Lng新批次);

  --更新费用记录的执行状态(0-未执行;1-完全执行;2-部分执行)
  Select Decode(Sum(Nvl(付数, 1) * 实际数量), Null, 0, 0, 0, 2)
  Into Int执行状态
  From 药品收发记录
  Where 单据 = Int单据 And NO = Strno And 费用id = Lng费用id And 审核人 Is Not Null;

  If 门诊_In = 1 Then
    Select 记录性质, 收费类别 Into n_记录性质, v_收费类别 From 门诊费用记录 Where ID = Lng费用id;
  Else
    Select 记录性质, 收费类别 Into n_记录性质, v_收费类别 From 住院费用记录 Where ID = Lng费用id;
  End If;

  If Int执行状态 = 0 Then
    If 门诊_In = 1 Then
      Update 门诊费用记录
      Set 执行状态 = Int执行状态, 执行人 = Null, 执行时间 = Null
      Where NO = Strno And
            序号 = (Select 序号 From 门诊费用记录 Where ID = (Select 费用id From 药品收发记录 Where ID = Billid_In)) And
            Mod(记录性质, 10) = n_记录性质 And 记录状态 <> 2 And 执行部门id = Lng库房id;
    Else
      Update 住院费用记录 Set 执行状态 = Int执行状态, 执行人 = Null, 执行时间 = Null Where ID = Lng费用id;
    End If;
  Else
    If 门诊_In = 1 Then
      Update 门诊费用记录
      Set 执行状态 = Int执行状态
      Where NO = Strno And
            序号 = (Select 序号 From 门诊费用记录 Where ID = (Select 费用id From 药品收发记录 Where ID = Billid_In)) And
            Mod(记录性质, 10) = n_记录性质 And 记录状态 <> 2 And 执行部门id = Lng库房id;
    Else
      Update 住院费用记录 Set 执行状态 = Int执行状态 Where ID = Lng费用id;
    End If;
  End If;

  --插入未发药品记录
  Begin
    If 门诊_In = 1 Then
      Insert Into 未发药品记录
        (单据, NO, 病人id, 主页id, 姓名, 优先级, 对方部门id, 库房id, 发药窗口, 填制日期, 已收费, 配药人, 打印状态, 未发数, 领药号, 排队状态)
        Select a.单据, a.No, a.病人id, Null, a.姓名, Nvl(b.优先级, 0) 优先级, a.对方部门id, a.库房id, a.发药窗口, a.填制日期, a.已收费, Null, 1, 1,
               a.产品合格证, v_排队状态
        From (Select b.单据, b.No, a.病人id, a.姓名, Decode(a.记录状态, 0, 0, 1) 已收费, b.对方部门id, b.库房id, b.发药窗口, b.填制日期, c.身份,
                      b.产品合格证
               From 门诊费用记录 A, 药品收发记录 B, 病人信息 C
               Where b.Id = Billid_In And a.Id = b.费用id + 0 And a.病人id = c.病人id(+)) A, 身份 B
        Where b.名称(+) = a.身份;
    Else
      Insert Into 未发药品记录
        (单据, NO, 病人id, 主页id, 姓名, 优先级, 对方部门id, 库房id, 发药窗口, 填制日期, 已收费, 配药人, 打印状态, 未发数, 领药号, 排队状态)
        Select a.单据, a.No, a.病人id, a.主页id, a.姓名, Nvl(b.优先级, 0) 优先级, a.对方部门id, a.库房id, a.发药窗口, a.填制日期, a.已收费, Null, 1, 1,
               a.产品合格证, v_排队状态
        From (Select b.单据, b.No, a.病人id, a.主页id, a.姓名, Decode(a.记录状态, 0, 0, 1) 已收费, b.对方部门id, b.库房id, b.发药窗口, b.填制日期,
                      c.身份, b.产品合格证
               From 住院费用记录 A, 药品收发记录 B, 病人信息 C
               Where b.Id = Billid_In And a.Id = b.费用id + 0 And a.病人id = c.病人id(+)) A, 身份 B
        Where b.名称(+) = a.身份;
    End If;
  
    --修改处方类型
    Zl_Prescription_Type_Update(Strno, n_记录性质, Lng药品id, v_收费类别);
  Exception
    When Others Then
      Null;
  End;

  --修改原记录为被冲销记录
  Update 药品收发记录 Set 记录状态 = Int记录状态 + Decode(Int记录状态, 1, 0, 1) + 2 Where ID = Billid_In;

  --修改药品库存(反冲库存)
  If Lng分批 <> 3 Then
    --正常单据需要将库存表实际数量和金额、差价还回去，如果库存表没有则在库存表插入数据
    Zl_药品库存_Update(v_冲销记录id, 3, 0);
  Else
    --原不分批，现在分批，直接在库存表产生新单据
    Insert Into 药品库存
      (库房id, 药品id, 批次, 效期, 性质, 实际数量, 实际金额, 实际差价, 零售价, 上次批号, 上次产地, 上次供应商id, 上次采购价, 上次生产日期, 批准文号, 平均成本价)
    Values
      (Lng库房id, Lng药品id, Lng新批次, 效期_In, 1, Dbl实际数量 * n_付数, Dbl实际金额, Dbl实际差价, Decode(n_时价分批, 1, n_零售价, Null), 批号_In,
       产地_In, n_上次供应商id, n_上次采购价, d_上次生产日期, v_批准文号, n_上次采购价);
  End If;

  Delete 药品库存
  Where 库房id + 0 = Lng库房id And 药品id = Lng药品id And 性质 = 1 And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
        Nvl(实际差价, 0) = 0;

  --处理调价修正
  Zl_药品收发记录_调价修正(v_冲销记录id);

  Begin
    --移动支付宝项目在发药后动态调用生成推送信息的过程
    Execute Immediate 'Begin zl_服务窗消息_发送(:1,:2); End;'
      Using 7, Billid_In || ',' || 退药数量_In || ',' || 门诊_In;
  Exception
    When Others Then
      Null;
  End;

  --消息处理，剩余全部退数量传0
  If Bln部分退药 = 1 Then
    b_Message.Zlhis_Drug_006(v_冲销记录id, Lng新批次, Dbl实际数量 * n_付数, Lng费用id);
  Else
    b_Message.Zlhis_Drug_006(v_冲销记录id, Lng新批次, 0, Lng费用id);
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品收发记录_部门退药;
/

--130781:胡俊勇,2019-03-18,检验危急值记录删除
Create Or Replace Procedure Zl_病人危急值记录_Delete(Id_In In 病人危急值记录.Id%Type) Is
Begin
  Delete 业务消息清单
  Where 类型编码 = 'ZLHIS_LIS_003' And 业务标识 = (Select To_Char(医嘱id) From 病人危急值记录 Where ID = Id_In);
  Delete 病人危急值记录 Where ID = Id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人危急值记录_Delete;
/



------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0054' Where 编号=&n_System;
Commit;
