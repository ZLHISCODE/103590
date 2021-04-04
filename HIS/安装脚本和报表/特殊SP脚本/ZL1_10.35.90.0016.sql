----------------------------------------------------------------------------------------------------------------
--本脚本支持从ZLHIS+ v10.35.90升级到 v10.35.90
--请以数据空间的所有者登录PLSQL并执行下列脚本
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------



------------------------------------------------------------------------------
--结构修正部份
------------------------------------------------------------------------------
--127220:胡俊勇,2018-06-12,区分医嘱执行时的操作来源
Alter Table 病人医嘱执行 Add 执行方式 Number(1);




------------------------------------------------------------------------------
--数据修正部份
------------------------------------------------------------------------------




-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------
--126262:殷瑞,2018-06-13,处方发药和部门发药新增电子病案查阅
Insert Into zlModuleRelas(系统,模块,功能,相关系统,相关模块,相关类型,相关功能,缺省值)
Select 100,1341,A.* From (
Select 功能,相关系统,相关模块,相关类型,相关功能,缺省值 From zlModuleRelas Where 1 = 0
Union All Select NULL,100,1259,1,NULL,1 From Dual) A;

Insert Into zlModuleRelas(系统,模块,功能,相关系统,相关模块,相关类型,相关功能,缺省值)
Select 100,1342,A.* From (
Select 功能,相关系统,相关模块,相关类型,相关功能,缺省值 From zlModuleRelas Where 1 = 0
Union All Select NULL,100,1259,1,NULL,1 From Dual) A;

Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select  &n_System,1341,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0 
Union All Select '电子病案查阅',30,'有该权限时，可以进行电子病案查阅',1 From Dual) A;

Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select  &n_System,1342,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0 
Union All Select '电子病案查阅',23,'有该权限时，可以进行电子病案查阅',1 From Dual) A;

-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------




-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--127220:胡俊勇,2018-06-12,区分医嘱执行时的操作来源
Create Or Replace Procedure Zl_病人医嘱执行_Insert
(
  医嘱id_In       In 病人医嘱执行.医嘱id%Type,
  发送号_In       In 病人医嘱执行.发送号%Type,
  要求时间_In     In 病人医嘱执行.要求时间%Type,
  本次数次_In     In 病人医嘱执行.本次数次%Type,
  执行摘要_In     In 病人医嘱执行.执行摘要%Type,
  执行人_In       In 病人医嘱执行.执行人%Type,
  执行时间_In     In 病人医嘱执行.执行时间%Type,
  单独执行_In     In Number := 0,
  自动完成_In     In Number := 0,
  执行结果_In     In 病人医嘱执行.执行结果%Type := 1,
  未执行原因_In   In 病人医嘱执行.说明%Type := Null,
  操作员编号_In   In 人员表.编号%Type := Null,
  操作员姓名_In   In 人员表.姓名%Type := Null,
  执行部门id_In   In 门诊费用记录.执行部门id%Type := 0,
  配液检查_In     In Number := 0,
  检验项目记帐_In In Number := 0,
  输液通道_In     In 病人医嘱执行.输液通道%Type := Null,
  记录来源_In     In 病人医嘱执行.记录来源%Type := Null,
  执行方式_In     In 病人医嘱执行.执行方式%Type := Null
  --参数：医嘱ID_IN=单独执行的医嘱ID，检验组合为显示的检验项目的ID。
  --      执行结果_In=1- 完成   =0  -未执行
  --      如果是台式机调用 操作员编号_In 操作员姓名_In 这两个参数必须传入
  --执行部门id_In=仅处理指定执行部门的费用，不传或传入0时不限制执行部门
  --配液检查_In=移动工作站调用时，是否检查配液信息。
  --检验项目记帐_In=如果是检验项目时，需要记帐但不完成医嘱发送状态
) Is
  --除了要执行的主记录,还包含了附加手术,检查部位的记录
  --手术麻醉,中药煎法,采集方法单独控制,检验组合只填写在第一个项目上，但执行状态相同
  v_组id       病人医嘱记录.Id%Type;
  v_诊疗类别   病人医嘱记录.诊疗类别%Type;
  v_自动完成   Number;
  v_病人来源   病人医嘱记录.病人来源%Type;
  v_费用性质   病人医嘱发送.记录性质%Type;
  v_操作类型   诊疗项目目录.操作类型%Type;
  n_执行分类   诊疗项目目录.执行分类%Type;
  v_病区id     病案主页.当前病区id%Type;
  v_配液病区   Varchar2(200);
  v_Count      Number;
  v_Temp       Varchar2(255);
  v_人员编号   人员表.编号%Type;
  v_人员姓名   人员表.姓名%Type;
  n_期效       病人医嘱记录.医嘱期效%Type;
  n_诊疗项目id 病人医嘱记录.诊疗项目id%Type;
  v_叮嘱执行   Varchar2(5);
  n_本次数次   病人医嘱执行.本次数次%Type;
  n_病人id     病人医嘱记录.病人id%Type;
  n_主页id     病人医嘱记录.主页id%Type;
  v_挂号单     病人医嘱记录.挂号单%Type;

  n_执行次数   Number;
  n_剩余次数   Number;
  n_执行状态   Number;
  d_终止时间   Date;
  d_开始时间   Date;
  n_发送数次   Number;
  n_登记数次   Number;
  n_单次数次   Number;
  d_要求时间   Date;
  n_执行科室id Number;
  n_用血医嘱   Number(1);

  v_Date  Date;
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --并发查检，防止产生多条执行记录
  Begin
    Select (a.发送数次 - c.登记次数) As 剩余数次, a.发送数次, a.执行部门id, Nvl(d.诊疗项目id, 0), c.登记次数
    Into v_Count, n_发送数次, n_执行科室id, n_诊疗项目id, n_登记数次
    From 病人医嘱发送 a,
         (Select 医嘱id_In As 医嘱id, 发送号_In As 发送号, Nvl(Sum(b.本次数次), 0) As 登记次数
           From 病人医嘱执行 b
           Where b.医嘱id = 医嘱id_In And b.发送号 = 发送号_In) c, 病人医嘱记录 d
    Where a.医嘱id = c.医嘱id And a.发送号 = c.发送号 And a.医嘱id = 医嘱id_In And a.医嘱id = d.Id And a.发送号 = 发送号_In;
  Exception
    When Others Then
      v_Count := 本次数次_In;
  End;
  v_叮嘱执行 := Zl_Getsysparameter(288);
  n_本次数次 := 本次数次_In;
  If 本次数次_In > v_Count And (Not (n_诊疗项目id = 0 And v_叮嘱执行 = 1)) Then
    If Round(n_登记数次 + 本次数次_In) = 1 Then
      --表明是输血执行
      n_本次数次 := 1 - n_登记数次;
    Else
      v_Error := '由于并发操作可能已经被他人登记，请刷新后再试。';
      Raise Err_Custom;
    End If;
  End If;
  --当前操作人员
  If 操作员编号_In Is Not Null And 操作员姓名_In Is Not Null Then
    v_人员编号 := 操作员编号_In;
    v_人员姓名 := 操作员姓名_In;
  Else
    Begin
      Select 姓名, 编号 Into v_人员姓名, v_人员编号 From 人员表 Where 姓名 = 执行人_In;
    Exception
      When Others Then
        v_Temp     := Zl_Identity;
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
        v_人员编号 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
        v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    End;
  End If;
  --对医嘱终止时间进行检查
  Select a.执行终止时间, a.开始执行时间, a.医嘱期效, a.病人id, a.主页id, a.挂号单
  Into d_终止时间, d_开始时间, n_期效, n_病人id, n_主页id, v_挂号单
  From 病人医嘱记录 a
  Where a.Id = 医嘱id_In;
  If Not d_终止时间 Is Null And n_期效 = 0 Then
    If 要求时间_In > d_终止时间 Then
      v_Error := '要求时间超过了医嘱终止时间，请确认医嘱是否提前停止！';
      Raise Err_Custom;
    End If;
  End If;
  If Not d_开始时间 Is Null Then
    If 执行时间_In < d_开始时间 Then
      v_Error := '执行时间必须大于医嘱的开始执行时间''' || To_Char(d_开始时间, 'yyyy-mm-dd HH24:mi:ss') || '''！';
      Raise Err_Custom;
    End If;
  End If;
  Select Sysdate Into v_Date From Dual;
  Select a.病人来源, 执行科室id, Nvl(a.相关id, a.Id), Nvl(a.诊疗类别, '*'), Nvl(b.操作类型, '0') 操作类型, Nvl(b.执行分类, 0) 执行分类
  Into v_病人来源, v_病区id, v_组id, v_诊疗类别, v_操作类型, n_执行分类
  From 病人医嘱记录 a, 诊疗项目目录 b
  Where a.Id = 医嘱id_In And a.诊疗项目id = b.Id(+);

  If v_病人来源 = 2 Then
    Select Decode(记录性质, 1, 1, Decode(门诊记帐, 1, 1, 2))
    Into v_费用性质
    From 病人医嘱发送
    Where 发送号 = 发送号_In And 医嘱id = 医嘱id_In;
  Else
    v_费用性质 := 1;
  End If;

  --移动系统配液检查
  If 配液检查_In = 1 Then
    --检查当前病人所属病区是否进行配液登记管理
    Select Nvl(Zl_Getsysparameter(184), '') Into v_配液病区 From Dual;
  
    If v_配液病区 Is Not Null And 执行结果_In <> 0 Then
      If Instr(',' || v_配液病区 || ',', ',' || v_病区id || ',') > 0 Then
        v_病区id   := 0;
        v_配液病区 := 'Select 1 From 病区配液记录 where 医嘱ID=:YZID AND 发送号=:FSH AND 要求时间=:YQSJ';
        Begin
          Execute Immediate v_配液病区
            Into v_病区id
            Using 医嘱id_In, 发送号_In, 要求时间_In;
        Exception
          When Others Then
            Null;
        End;
        If v_病区id = 0 Then
          v_Error := '当前医嘱还未进行配液，不允许进行执行登记！';
          Raise Err_Custom;
        End If;
      End If;
    End If;
    --检查当前医嘱是否已配液
  End If;

  --病人医嘱执行
  Select Count(1)
  Into v_Count
  From 病人医嘱执行
  Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In And 执行时间 = 执行时间_In;
  If v_Count > 0 Then
    v_Error := '您指定的执行时间，已经执行过本条医嘱，请更改一个执行时间。';
    Raise Err_Custom;
  End If;
  Insert Into 病人医嘱执行
    (医嘱id, 发送号, 要求时间, 本次数次, 执行摘要, 执行人, 执行时间, 登记时间, 登记人, 执行结果, 说明, 输液通道, 执行科室id, 记录来源,执行方式)
  Values
    (医嘱id_In, 发送号_In, 要求时间_In, n_本次数次, 执行摘要_In, 执行人_In, 执行时间_In, v_Date, v_人员姓名, 执行结果_In, 未执行原因_In, 输液通道_In, n_执行科室id,
     记录来源_In,执行方式_In);

  b_Message.Zlhis_Cis_050(n_病人id, n_主页id, v_挂号单, 发送号_In, 医嘱id_In, 要求时间_In, 执行时间_In);

  --费用记录的执行状态进行更新
  Select Decode(a.执行状态, 1, a.发送数次, c.登记次数), Decode(a.执行状态, 1, 0, a.发送数次 - c.登记次数), c.登记次数
  Into n_执行次数, n_剩余次数, n_登记数次
  From 病人医嘱发送 a,
       (Select 医嘱id_In 医嘱id, 发送号_In 发送号, Nvl(Sum(b.本次数次), 0) As 登记次数
         From 病人医嘱执行 b
         Where b.医嘱id = 医嘱id_In And b.发送号 = 发送号_In And Nvl(b.执行结果, 1) = 1) c
  Where a.医嘱id = c.医嘱id And a.发送号 = c.发送号 And a.医嘱id = 医嘱id_In And a.发送号 = 发送号_In;
  --如果全部执行则状态为1，未执行状态为0，部分执行状态为2
  Select Decode(n_剩余次数, 0, 1, Decode(n_执行次数, 0, 0, 2)) Into n_执行状态 From Dual;

  --填写了执行状态后就标记为正在执行
  If Nvl(单独执行_In, 0) = 1 Then
    Update 病人医嘱发送
    Set 执行状态 = Decode(n_执行次数, 0, 0, 3)
    Where 执行状态 In (0, 3) And 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In;
  Else
    Update 病人医嘱发送
    Set 执行状态 = Decode(n_执行次数, 0, 0, 3)
    Where 执行状态 In (0, 3) And 发送号 + 0 = 发送号_In And
          医嘱id In (Select Id
                   From 病人医嘱记录
                   Where Id = v_组id And Nvl(诊疗类别, '*') = v_诊疗类别
                   Union All
                   Select Id
                   From 病人医嘱记录
                   Where 相关id = v_组id And Nvl(诊疗类别, '*') = v_诊疗类别);
  End If;

  --更新对应的费用执行状态为已执行(无正在执行)
  --不应该处理药品和跟踪在用的卫材
  If 执行结果_In = 1 Then
    If v_费用性质 = 2 Then
      If Nvl(单独执行_In, 0) = 1 Then
        Update 住院费用记录 a
        Set 执行状态 = n_执行状态, 执行人 = Decode(n_执行状态, 0, Null, 执行人_In), 执行时间 = Decode(n_执行状态, 0, Null, 执行时间_In)
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And Not Exists
         (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1) And a.记录状态 In (0, 1, 3) And
              (医嘱序号, No, 记录性质) In
              (Select 医嘱id, No, 记录性质
               From 病人医嘱发送
               Where 执行状态 = Decode(n_执行次数, 0, 0, 3) And 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In);
      Else
        Update 住院费用记录 a
        Set 执行状态 = n_执行状态, 执行人 = Decode(n_执行状态, 0, Null, 执行人_In), 执行时间 = Decode(n_执行状态, 0, Null, 执行时间_In)
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And Not Exists
         (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1) And a.记录状态 In (0, 1, 3) And
              (医嘱序号, No, 记录性质) In (Select 医嘱id, No, 记录性质
                                   From 病人医嘱发送
                                   Where 执行状态 = Decode(n_执行次数, 0, 0, 3) And 发送号 + 0 = 发送号_In And
                                         医嘱id In (Select Id
                                                  From 病人医嘱记录
                                                  Where Id = v_组id And 诊疗类别 = v_诊疗类别
                                                  Union All
                                                  Select Id
                                                  From 病人医嘱记录
                                                  Where 相关id = v_组id And 诊疗类别 = v_诊疗类别));
      End If;
    Else
      If Nvl(单独执行_In, 0) = 1 Then
        --对于门诊单据n_执行状态可能为0（登记执行情况，选择执行结果为未执行），因此需判断
        Update 门诊费用记录 a
        Set 执行状态 = n_执行状态, 执行人 = Decode(n_执行状态, 0, Null, 执行人_In), 执行时间 = Decode(n_执行状态, 0, Null, 执行时间_In)
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And Not Exists
         (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1) And a.记录状态 In (0, 1, 3) And
              (医嘱序号, No, 记录性质) In
              (Select 医嘱id, No, 记录性质
               From 病人医嘱发送
               Where 执行状态 = Decode(n_执行次数, 0, 0, 3) And 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In);
      Else
        Update 门诊费用记录 a
        Set 执行状态 = n_执行状态, 执行人 = Decode(n_执行状态, 0, Null, 执行人_In), 执行时间 = Decode(n_执行状态, 0, Null, 执行时间_In)
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And Not Exists
         (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1) And a.记录状态 In (0, 1, 3) And
              (医嘱序号, No, 记录性质) In (Select 医嘱id, No, 记录性质
                                   From 病人医嘱发送
                                   Where 执行状态 = Decode(n_执行次数, 0, 0, 3) And 发送号 + 0 = 发送号_In And
                                         医嘱id In (Select Id
                                                  From 病人医嘱记录
                                                  Where Id = v_组id And 诊疗类别 = v_诊疗类别
                                                  Union All
                                                  Select Id
                                                  From 病人医嘱记录
                                                  Where 相关id = v_组id And 诊疗类别 = v_诊疗类别));
      End If;
    End If;
    --检验自动完成采集
    If v_诊疗类别 = 'E' And v_操作类型 = '6' Then
      Update 病人医嘱发送 a
      Set a.采样人 = 执行人_In, a.采样时间 = 执行时间_In
      Where 医嘱id In
            (Select Id From 病人医嘱记录 Where Id = v_组id Union All Select Id From 病人医嘱记录 Where 相关id = v_组id) And
            发送号 = 发送号_In;
    End If;
  
    --执行数次达到之后自动完成执行(主要用于PDA自动执行)，如果启用了移动临床，则护士站和PDA一致。
    v_自动完成 := 自动完成_In;
    If 自动完成_In = 1 Then
      --医嘱已经是完成状态则不用再调用执行完成过程此处先设为不自动完成
      Select Max(a.执行状态) Into v_Count From 病人医嘱发送 a Where a.医嘱id = 医嘱id_In And a.发送号 = 发送号_In;
      If v_Count = 1 Then
        v_自动完成 := 0;
      End If;
      v_Count := Null;
    End If;
    --移动端执行如果非用血医嘱，且自动完成=0，则不自动完成医嘱执行（有血库接口内部处理）
    n_用血医嘱 := 0;
    If v_诊疗类别 = 'E' And v_操作类型 = '8' And n_执行分类 = 1 Then
      n_用血医嘱 := Zl_To_Number(Nvl(Zl_Getsysparameter(236), '0'));
    End If;
    If Nvl(v_自动完成, 0) = 0 And (v_病人来源 = 2 Or v_病人来源 = 1) And Instr('C,D', v_诊疗类别) = 0 And n_用血医嘱 = 0 Then
      Begin
        Execute Immediate 'Select Count(1) From ZLMBSYSTEMS'
          Into v_Count;
      Exception
        When Others Then
          Null;
      End;
      If v_Count > 0 Then
        v_自动完成 := 1;
      End If;
    End If;
  
    If Nvl(v_自动完成, 0) = 1 Or 检验项目记帐_In = 1 Then
      Begin
        Select Decode(Sign(Nvl(Sum(b.本次数次), 0) - a.发送数次), 1, 1, 0, 1, 0)
        Into v_自动完成
        From 病人医嘱发送 a, 病人医嘱执行 b
        Where a.医嘱id = b.医嘱id And a.发送号 = b.发送号 And a.执行状态 In (0, 3) And a.医嘱id = 医嘱id_In And a.发送号 = 发送号_In
        Group By a.发送数次;
      Exception
        When Others Then
          Null;
      End;
    
      If Nvl(v_自动完成, 0) = 1 Or 检验项目记帐_In = 1 Then
        Zl_病人医嘱执行_Finish(医嘱id_In,
                         发送号_In,
                         Null,
                         单独执行_In,
                         v_人员编号,
                         v_人员姓名,
                         执行部门id_In,
                         检验项目记帐_In);
      End If;
    End If;
    --更新医嘱执行计价.执行状态
    If n_发送数次 > 0 Then
      Select Count(Distinct 要求时间) Into v_Count From 医嘱执行计价 Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In;
      If v_Count > 0 Then
        n_单次数次 := n_发送数次 / v_Count;
        --已执行数量+本次数次 总共能够执行多少个时间点,取最大整数
        v_Count := Ceil((n_登记数次) / n_单次数次);
        --获取执行截至要求时间
        Select 要求时间
        Into d_要求时间
        From (Select 要求时间, Rownum As 次数
               From (Select Distinct 要求时间
                      From 医嘱执行计价
                      Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In
                      Order By 要求时间))
        Where 次数 = v_Count;
      
        If Not d_要求时间 Is Null Then
          --先检查是否已经退费
          Select Max(Nvl(执行状态, 0))
          Into v_Count
          From 医嘱执行计价
          Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In And 要求时间 <= d_要求时间;
          If v_Count = 2 Then
            v_Error := '您指定的执行时间段的医嘱费用已经被退费，不允许再执行。';
            Raise Err_Custom;
          End If;
          --更新截至要求时间之前(含)的记录执行状态；
          Update 医嘱执行计价
          Set 执行状态 = 1
          Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In And 要求时间 <= d_要求时间 And Nvl(执行状态, 0) <> 2;
        End If;
      End If;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病人医嘱执行_Insert;
/
--127220:胡俊勇,2018-06-12,区分医嘱执行时的操作来源
Create Or Replace Procedure Zl_Third_Adviceoperation
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --功能：医嘱执行登记/取消执行登记 /写入
  --1、用于每次医嘱执行时的执行登记，”来源“用于区分是否移动端执行
  --2、对已执行的医嘱取消执行操作
  --入参：xml_in
  --<IN>
  -- <TYPE>1</TYPE>   --操作类型：1、执行登记；0、取消执行登记
  -- <YZID>1162695</YZID>   --医嘱id
  -- <FSH>202704</FSH>   --发送号
  -- <YQSJ>2017-12-05 10:00:00</YQSJ>   --要求时间
  -- <CZY></CZY>   --操作员
  -- <CZSJ>2017-12-05 16:26:54</CZSJ>   --操作时间

  --以下节点取消执行时传空
  -- <ZXSM>PDA执行</ZXSM>   --执行摘要
  -- <ZXCS></ZXCS>   --执行次数
  -- <SYTD>左手背</SYTD>
  -- <JLLY>1</JLLY>    --记录来源，0-PC端登记，1-移动临床登记，移动端固定传1
  -- <ZXFS>2</ZXFS>  --执行方式，0-常规(无意义)，1-手工执行，2-扫码执行

  --</IN>

  --出参：Xml_Out
  --<OUTPUT>
  --   <RESULT>true</RESULT>
  --</OUTPUT>

  --失败：
  --<OUTPUT>
  --   <RESULT>false</RESULT>
  --   <ERROR>
  --     <MSG>详细错误提示</MSG>
  --   </ERROR>
  --</OUTPUT>

  n_Type     Number;
  n_医嘱id   病人医嘱记录.Id%Type;
  n_发送号   病人医嘱发送.发送号%Type;
  d_要求时间 病人医嘱执行.要求时间%Type;
  d_执行时间 病人医嘱执行.执行时间%Type;
  v_执行摘要 病人医嘱执行.执行摘要%Type;
  n_本次数次 病人医嘱执行.本次数次%Type;
  v_输液通道 病人医嘱执行.输液通道%Type;
  n_记录来源 病人医嘱执行.记录来源%Type;
  v_执行人   病人医嘱执行.执行人%Type;
  n_执行方式 病人医嘱执行.执行方式%Type;
Begin
  Select Extractvalue(Value(A), 'IN/TYPE'), Extractvalue(Value(A), 'IN/YZID') As 医嘱id,
         Extractvalue(Value(A), 'IN/FSH') As 发送号,
         To_Date(Extractvalue(Value(A), 'IN/YQSJ'), 'yyyy-mm-dd hh24:mi:ss') As 要求时间,
         Extractvalue(Value(A), 'IN/CZY') As 执行人,
         To_Date(Extractvalue(Value(A), 'IN/CZSJ'), 'yyyy-mm-dd hh24:mi:ss') As 执行时间,
         Extractvalue(Value(A), 'IN/ZXSM') As 执行摘要, Extractvalue(Value(A), 'IN/ZXCS') As 本次数次,
         Extractvalue(Value(A), 'IN/SYTD') As 输液通道, Extractvalue(Value(A), 'IN/JLLY') As 记录来源,
         Extractvalue(Value(A), 'IN/ZXFS') As 执行方式
  Into n_Type, n_医嘱id, n_发送号, d_要求时间, v_执行人, d_执行时间, v_执行摘要, n_本次数次, v_输液通道, n_记录来源, n_执行方式
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If n_Type = 1 Then
    Zl_病人医嘱执行_Insert(n_医嘱id, n_发送号, d_要求时间, n_本次数次, v_执行摘要, v_执行人, d_执行时间, 0, 0, 1, Null, Null, Null, 0, 0, 0, v_输液通道,
                     n_记录来源, n_执行方式);
  Else
    Zl_病人医嘱执行_Delete(n_医嘱id, n_发送号, d_执行时间, 0, 0, 0);
  End If;
  Xml_Out := Xmltype('<OUTPUT><RESULT>true</RESULT></OUTPUT>');
Exception
  When Others Then
    Xml_Out := Xmltype('<OUTPUT><RESULT>false</RESULT><ERROR><MSG>' || SQLCode || '***' || SQLErrM ||
                       '</MSG></ERROR></OUTPUT>');
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Adviceoperation;
/





------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0016' Where 编号=&n_System;
Commit;