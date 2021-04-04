----------------------------------------------------------------------------------------------------------------
--本脚本支持从ZLHIS+ v10.35.90升级到 v10.35.90
--请以数据空间的所有者登录PLSQL并执行下列脚本
Define n_System=100;

----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------



------------------------------------------------------------------------------
--结构修正部份
------------------------------------------------------------------------------
--122812:焦博,2018-06-07,票据入库记录增加字段批次,票据领用记录增加字段入库ID
Alter Table 票据入库记录 Add 批次 varchar2(20);

Alter Table 消费卡入库记录 Add 批次 varchar2(20);

Alter Table 票据领用记录 Add 入库ID number(18);

Alter Table 消费卡领用记录 Add 入库ID number(18);

Create Index 票据入库记录_IX_批次 On 票据入库记录(批次) Tablespace zl9Indexhis;

Create Index 消费卡入库记录_IX_批次 On 消费卡入库记录(批次) Tablespace zl9Indexhis;

Create Index 票据领用记录_IX_入库ID On 票据领用记录(入库ID) Tablespace zl9Indexhis;

Create Index 消费卡领用记录_IX_入库ID On 消费卡领用记录(入库ID) Tablespace zl9Indexhis;

Alter Table 票据领用记录 Add Constraint 票据领用记录_FK_入库ID foreign Key(入库ID) References 票据入库记录(ID) On Delete Cascade;
Alter Table 消费卡领用记录 Add Constraint 消费卡领用记录_FK_入库ID foreign Key(入库ID) References 消费卡入库记录(ID) On Delete Cascade;

------------------------------------------------------------------------------
--数据修正部份
------------------------------------------------------------------------------
--122812:焦博,2018-06-07,票据入库记录增加字段批次,票据领用记录增加字段入库ID
Update 票据入库记录 Set 批次 = ID;

Update 票据领用记录 A
Set a.入库id = a.批次
Where Exists (Select 1 From 票据入库记录 Where ID = a.批次);

Update 消费卡入库记录 Set 批次 = ID;

Update 消费卡领用记录 A
Set a.入库id = a.批次
Where Exists (Select 1 From 消费卡入库记录 Where ID = a.批次);





-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------





-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------




-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--126826:刘鹏飞,2018-06-08,用血医嘱执行登记不自动完成
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
  记录来源_In     In 病人医嘱执行.记录来源%Type := Null
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
    (医嘱id, 发送号, 要求时间, 本次数次, 执行摘要, 执行人, 执行时间, 登记时间, 登记人, 执行结果, 说明, 输液通道, 执行科室id, 记录来源)
  Values
    (医嘱id_In, 发送号_In, 要求时间_In, n_本次数次, 执行摘要_In, 执行人_In, 执行时间_In, v_Date, v_人员姓名, 执行结果_In, 未执行原因_In, 输液通道_In, n_执行科室id,
     记录来源_In);

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

--122812:焦博,2018-06-07,票据入库记录增加字段批次,票据领用记录增加字段入库ID
CREATE OR REPLACE Procedure ZL_票据入库记录_INSERT
(
  Id_In       In 票据入库记录.Id%Type,
  票种_In     In 票据入库记录.票种%Type,
  使用类别_In In 票据入库记录.使用类别%Type,
  前缀文本_In In 票据入库记录.前缀文本%Type,
  开始号码_In In 票据入库记录.开始号码%Type,
  终止号码_In In 票据入库记录.终止号码%Type,
  入库数量_In In 票据入库记录.入库数量%Type,
  剩余数量_In In 票据入库记录.剩余数量%Type,
  备注_In     In 票据入库记录.备注%Type,
  登记人_In   In 票据入库记录.登记人%Type,
  修改标志_In Integer := 0,
  批次_In     In 票据入库记录.批次%Type := Null
) Is
  --修改标志_In:0-增加;1-修改 
  v_Err_Msg Varchar2(100);
  Err_Item Exception;
  n_Max_Len Number(18);
Begin
  If Nvl(修改标志_In, 0) = 0 Then
    Insert Into 票据入库记录
      (ID, 票种, 使用类别, 前缀文本, 开始号码, 终止号码, 入库数量, 剩余数量, 有无票据, 备注, 登记人, 登记时间, 批次)
    Values
      (Id_In, 票种_In, 使用类别_In, 前缀文本_In, 开始号码_In, 终止号码_In, 入库数量_In, 剩余数量_In, Decode(Sign(Nvl(剩余数量_In, 0)), 1, 1, Null),
       备注_In, 登记人_In, Sysdate, Nvl(批次_In, Id_In));
    Return;
  End If;

  Begin
    Select Length(Min(开始号码))
    Into n_Max_Len
    From (Select Min(开始号码) As 开始号码
           From 票据报损记录
           Where 入库id = Id_In
           Union All
           Select Min(开始号码) As 开始号码 From 票据领用记录 Where 批次 = Id_In);
  Exception
    When Others Then
      n_Max_Len := Null;
  End;

  If Not n_Max_Len Is Null Then
    If Length(开始号码_In) <> n_Max_Len Then
      v_Err_Msg := '[ZLSOFT]这张入库的票据已经被使用过, 号码长度不能改变,号码长度应该是' || n_Max_Len || '![ZLSOFT]';
      Raise Err_Item;
    End If;
  End If;

  --修改 
  Update 票据领用记录
  Set 使用类别 = 使用类别_In, 批次 = 批次_In
  Where (票种, 入库id) In (Select 票种, ID From 票据入库记录 Where ID = Id_In) And Nvl(剩余数量, 0) > 0;

  Update 票据入库记录
  Set 前缀文本 = 前缀文本_In, 开始号码 = 开始号码_In, 终止号码 = 终止号码_In, 入库数量 = 入库数量_In, 剩余数量 = 剩余数量_In,
      有无票据 = Decode(Sign(Nvl(剩余数量_In, 0)), -1, Null, 0, Null, 1), 备注 = 备注_In, 登记人 = 登记人_In, 登记时间 = Sysdate,
      使用类别 = 使用类别_In, 批次 = 批次_In
  Where ID = Id_In;
  If Sql%NotFound Then
    v_Err_Msg := '[ZLSOFT]该入库票据未找到,可能已经被他人删除,不能修改![ZLSOFT]';
    Raise Err_Item;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_票据入库记录_Insert;
/

--122812:焦博,2018-06-07,票据入库记录增加字段批次,票据领用记录增加字段入库ID
CREATE OR REPLACE Procedure Zl_消费卡入库记录_Insert
(
  Id_In       In 消费卡入库记录.Id%Type,
  接口编号_In In 消费卡入库记录.接口编号%Type,
  前缀文本_In In 消费卡入库记录.前缀文本%Type,
  开始卡号_In In 消费卡入库记录.开始卡号%Type,
  终止卡号_In In 消费卡入库记录.终止卡号%Type,
  入库数量_In In 消费卡入库记录.入库数量%Type,
  剩余数量_In In 消费卡入库记录.剩余数量%Type,
  备注_In     In 消费卡入库记录.备注%Type,
  登记人_In   In 消费卡入库记录.登记人%Type,
  修改标志_In Integer := 0,
  批次_In     In 消费卡入库记录.批次%Type := Null
) Is
  --修改标志_In:0-增加;1-修改
  v_Err_Msg Varchar2(100);
  Err_Item Exception;

  n_Max_Len Number(18);
Begin
  If Nvl(修改标志_In, 0) = 0 Then
    Insert Into 消费卡入库记录
      (ID, 接口编号, 前缀文本, 开始卡号, 终止卡号, 入库数量, 剩余数量, 是否存在卡, 备注, 登记人, 登记时间, 批次)
    Values
      (Id_In, 接口编号_In, 前缀文本_In, 开始卡号_In, 终止卡号_In, 入库数量_In, 剩余数量_In, Decode(Sign(Nvl(剩余数量_In, 0)), 1, 1, 0), 备注_In,
       登记人_In, Sysdate, Nvl(批次_In, Id_In));
    Return;
  End If;

  Begin
    Select Length(Min(开始卡号))
    Into n_Max_Len
    From (Select Min(开始卡号) As 开始卡号
           From 消费卡报损记录
           Where 入库id = Id_In
           Union All
           Select Min(开始卡号) As 开始卡号 From 消费卡领用记录 Where 批次 = Id_In);
  Exception
    When Others Then
      n_Max_Len := Null;
  End;

  If Not n_Max_Len Is Null Then
    If Length(开始卡号_In) <> n_Max_Len Then
      v_Err_Msg := '这张入库单已经被使用过，卡号长度不能改变，卡号长度应该是' || n_Max_Len || '！';
      Raise Err_Item;
    End If;
  End If;

  --修改
  Update 消费卡领用记录 Set 接口编号 = 接口编号_In, 批次 = 批次_In Where 入库id = Id_In And Nvl(剩余数量, 0) > 0;

  Update 消费卡入库记录
  Set 前缀文本 = 前缀文本_In, 开始卡号 = 开始卡号_In, 终止卡号 = 终止卡号_In, 入库数量 = 入库数量_In, 剩余数量 = 剩余数量_In,
      是否存在卡 = Decode(Sign(Nvl(剩余数量_In, 0)), 1, 1, 0), 备注 = 备注_In, 登记人 = 登记人_In, 登记时间 = Sysdate, 接口编号 = 接口编号_In,
      批次 = 批次_In
  Where ID = Id_In;
  If Sql%NotFound Then
    v_Err_Msg := '该入库单未找到，可能已经被他人删除，不能修改！';
    Raise Err_Item;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_消费卡入库记录_Insert;
/

--122812:焦博,2018-06-07,票据入库记录增加字段批次,票据领用记录增加字段入库ID
CREATE OR REPLACE Procedure Zl_票据领用记录_Insert
(
  Id_In       In 票据领用记录.Id%Type,
  票种_In     In 票据领用记录.票种%Type,
  使用类别_In In 票据领用记录.使用类别%Type,
  领用人_In   In 票据领用记录.领用人%Type,
  前缀文本_In In 票据领用记录.前缀文本%Type,
  开始号码_In In 票据领用记录.开始号码%Type,
  终止号码_In In 票据领用记录.终止号码%Type,
  使用方式_In In 票据领用记录.使用方式%Type,
  登记时间_In In 票据领用记录.登记时间%Type := Null,
  登记人_In   In 票据领用记录.登记人%Type := Null,
  剩余数量_In In 票据领用记录.剩余数量%Type := Null,
  批次_In     In 票据领用记录.批次%Type := Null,
  签字人_In   In 票据领用记录.签字人%Type := Null,
  入库id_In   In 票据领用记录.入库id%Type := Null
) Is
  v_Err_Msg Varchar2(500);
  Err_Item Exception;
  n_Count  Number(18);
  n_剩余数 票据入库记录.剩余数量%Type;
Begin

  For v_入库 In (Select ID, 前缀文本, 使用类别, 开始号码, Nvl(终止号码, 开始号码) As 终止号码
               From 票据入库记录
               Where ID = Nvl(入库id_In, 0) And 票种 = 票种_In) Loop
    --1. 入库检查
    If 开始号码_In < v_入库.开始号码 Or 开始号码_In > v_入库.终止号码 Then
      --不在入库范围,不能领用
      v_Err_Msg := '[ZLSOFT]当前领用的开始号码『' || 开始号码_In || '』不在入库范围' || Chr(10) || Chr(13) || '『' || v_入库.开始号码 || '-' ||
                   v_入库.终止号码 || '』不能领用该票据![ZLSOFT]';
      Raise Err_Item;
    End If;
  
    If 终止号码_In < v_入库.开始号码 Or 终止号码_In > v_入库.终止号码 Then
      --不在入库范围,不能领用
      v_Err_Msg := '[ZLSOFT]当前领用的终止号码『' || 终止号码_In || '』不在入库范围' || Chr(10) || Chr(13) || '『' || v_入库.开始号码 || '-' ||
                   v_入库.终止号码 || '』不能领用该票据![ZLSOFT]';
      Raise Err_Item;
    End If;
    If Nvl(v_入库.使用类别, 'LXH') <> Nvl(使用类别_In, 'LXH') Then
      v_Err_Msg := '[ZLSOFT]入库的使用类别『' || Nvl(v_入库.使用类别, '') || '』与领用的类别『' || Nvl(使用类别_In, '') || '』不一致![ZLSOFT]';
      Raise Err_Item;
    End If;
  
    --2.检查票据是否已经被报损,不能重复报损
    Select Count(*)
    Into n_Count
    From 票据报损记录
    Where 入库id = Nvl(入库id_In, 0) And ((开始号码_In Between 开始号码 And 终止号码) Or (终止号码_In Between 开始号码 And 终止号码) Or
          (开始号码 Between 开始号码_In And 终止号码_In) Or (终止号码 Between 开始号码_In And 终止号码_In));
  
    If n_Count <> 0 Then
      If 开始号码_In = 终止号码_In Then
        v_Err_Msg := '[ZLSOFT]号码:' || 开始号码_In || '在报损记录中已经存在,不能再进行领用![ZLSOFT]';
      Else
        v_Err_Msg := '[ZLSOFT]号码:' || 开始号码_In || '-' || 终止号码_In || '在报损记录中已经存在,不能再进行领用![ZLSOFT]';
      End If;
      Raise Err_Item;
    End If;
  
    --3.检查票据是否已经被领用,领用的不能再进行报损
    Select Count(*)
    Into n_Count
    From 票据领用记录
    Where 批次 = Nvl(批次_In, 0) And ((开始号码_In Between 开始号码 And 终止号码) Or (终止号码_In Between 开始号码 And 终止号码) Or
          (开始号码 Between 开始号码_In And 终止号码_In) Or (终止号码 Between 开始号码_In And 终止号码_In));
    If n_Count <> 0 Then
      If 开始号码_In = 终止号码_In Then
        v_Err_Msg := '[ZLSOFT]号码:' || 开始号码_In || '在票据领用记录中已经存在,不能再进行领用![ZLSOFT]';
      Else
        v_Err_Msg := '[ZLSOFT]号码:' || 开始号码_In || '-' || 终止号码_In || '在票据领用记录中已经存在,不能再进行领用![ZLSOFT]';
      End If;
      Raise Err_Item;
    End If;
    --减少库存
    Update 票据入库记录
    Set 剩余数量 = Nvl(剩余数量, 0) - Nvl(剩余数量_In, 0)
    Where ID = Nvl(入库id_In, 0) And 票种 = 票种_In And Nvl(使用类别, 'LXH') = Nvl(使用类别_In, 'LXH')
    Returning Nvl(剩余数量, 0) Into n_剩余数;
  
    If n_剩余数 < 0 Then
      v_Err_Msg := '[ZLSOFT]入库票据的剩余票据数不足,请检查![ZLSOFT]';
      Raise Err_Item;
    End If;
    If n_剩余数 = 0 Then
      Update 票据入库记录
      Set 有无票据 = Null
      Where ID = Nvl(入库id_In, 0) And 票种 = 票种_In And Nvl(使用类别, 'LXH') = Nvl(使用类别_In, 'LXH');
    End If;
  End Loop;

  Insert Into 票据领用记录
    (ID, 票种, 使用类别, 领用人, 前缀文本, 开始号码, 终止号码, 使用方式, 登记时间, 登记人, 剩余数量, 批次, 签字人, 签字时间, 入库id)
  Values
    (Id_In, 票种_In, 使用类别_In, 领用人_In, 前缀文本_In, 开始号码_In, 终止号码_In, 使用方式_In, 登记时间_In, 登记人_In, 剩余数量_In, 批次_In, 签字人_In,
     Decode(签字人_In, Null, Null + Sysdate, Sysdate), 入库id_In);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_票据领用记录_Insert;
/

--122812:焦博,2018-06-07,票据入库记录增加字段批次,票据领用记录增加字段入库ID
CREATE OR REPLACE Procedure Zl_消费卡领用记录_Insert
(
  Id_In       消费卡领用记录.Id%Type,
  接口编号_In 消费卡领用记录. 接口编号%Type,
  领用人_In   消费卡领用记录.领用人%Type,
  前缀文本_In 消费卡领用记录.前缀文本%Type,
  开始卡号_In 消费卡领用记录.开始卡号%Type,
  终止卡号_In 消费卡领用记录.终止卡号%Type,
  使用方式_In 消费卡领用记录.使用方式%Type,
  登记时间_In 消费卡领用记录.登记时间%Type := Null,
  登记人_In   消费卡领用记录.登记人%Type := Null,
  剩余数量_In 消费卡领用记录.剩余数量%Type := Null,
  批次_In     消费卡领用记录.批次%Type := Null,
  签字人_In   消费卡领用记录.签字人%Type := Null,
  入库ID_In   消费卡领用记录.入库ID%Type := Null
) Is
  v_Err_Msg Varchar2(500);
  Err_Item Exception;

  n_Count  Number(18);
  n_剩余数 消费卡入库记录.剩余数量%Type;
Begin

  For r_入库 In (Select ID, 前缀文本, 接口编号, 开始卡号, Nvl(终止卡号, 开始卡号) As 终止卡号
               From 消费卡入库记录
               Where ID = Nvl(入库ID_In, 0)) Loop
    --1. 入库检查
    If 开始卡号_In < r_入库.开始卡号 Or 开始卡号_In > r_入库.终止卡号 Then
      --不在入库范围,不能领用
      v_Err_Msg := '当前领用的开始卡号『' || 开始卡号_In || '』不在入库范围' || Chr(10) || Chr(13) || '『' || r_入库.开始卡号 || '-' || r_入库.终止卡号 ||
                   '』不能领用该卡片！';
      Raise Err_Item;
    End If;

    If 终止卡号_In < r_入库.开始卡号 Or 终止卡号_In > r_入库.终止卡号 Then
      --不在入库范围,不能领用
      v_Err_Msg := '当前领用的终止卡号『' || 终止卡号_In || '』不在入库范围' || Chr(10) || Chr(13) || '『' || r_入库.开始卡号 || '-' || r_入库.终止卡号 ||
                   '』不能领用该卡片！';
      Raise Err_Item;
    End If;
    If r_入库.接口编号 <> 接口编号_In Then
      v_Err_Msg := '入库的卡类别『' || Nvl(接口编号_In, '') || '』与入库的类别不一致『' || Nvl(r_入库.接口编号, '') || '』!';
      Raise Err_Item;
    End If;

    --2.检查卡号是否已经被报损,不能重复报损
    Select Count(1)
    Into n_Count
    From 消费卡报损记录
    Where 入库id = Nvl(入库ID_In, 0) And ((开始卡号_In Between 开始卡号 And 终止卡号) Or (终止卡号_In Between 开始卡号 And 终止卡号) Or
          (开始卡号 Between 开始卡号_In And 终止卡号_In) Or (终止卡号 Between 开始卡号_In And 终止卡号_In));

    If n_Count <> 0 Then
      If 开始卡号_In = 终止卡号_In Then
        v_Err_Msg := '卡号:' || 开始卡号_In || '在报损记录中已经存在，不能再进行领用！';
      Else
        v_Err_Msg := '卡号:' || 开始卡号_In || '-' || 终止卡号_In || '在报损记录中已经存在，不能再进行领用！';
      End If;
      Raise Err_Item;
    End If;

    --3.检查卡号是否已经被领用,领用的不能再进行报损
    Select Count(1)
    Into n_Count
    From 消费卡领用记录
    Where 批次 = Nvl(批次_In, 0) And ((开始卡号_In Between 开始卡号 And 终止卡号) Or (终止卡号_In Between 开始卡号 And 终止卡号) Or
          (开始卡号 Between 开始卡号_In And 终止卡号_In) Or (终止卡号 Between 开始卡号_In And 终止卡号_In));
    If n_Count <> 0 Then
      If 开始卡号_In = 终止卡号_In Then
        v_Err_Msg := '卡号:' || 开始卡号_In || '在消费卡领用记录中已经存在，不能再进行领用！';
      Else
        v_Err_Msg := '卡号:' || 开始卡号_In || '-' || 终止卡号_In || '在消费卡领用记录中已经存在，不能再进行领用！';
      End If;
      Raise Err_Item;
    End If;
    --减少库存
    Update 消费卡入库记录
    Set 剩余数量 = Nvl(剩余数量, 0) - Nvl(剩余数量_In, 0)
    Where ID = Nvl(入库ID_In, 0) And 接口编号 = 接口编号_In
    Returning Nvl(剩余数量, 0) Into n_剩余数;

    If n_剩余数 < 0 Then
      v_Err_Msg := '入库卡片的剩余票据数不足，请检查！';
      Raise Err_Item;
    End If;
    If n_剩余数 = 0 Then
      Update 消费卡入库记录 Set 是否存在卡 = 0 Where ID = Nvl(入库ID_In, 0) And 接口编号 = 接口编号_In;
    End If;
  End Loop;

  Insert Into 消费卡领用记录
    (ID, 接口编号, 领用人, 前缀文本, 开始卡号, 终止卡号, 使用方式, 登记时间, 登记人, 剩余数量, 批次, 签字人, 签字时间)
  Values
    (Id_In, 接口编号_In, 领用人_In, 前缀文本_In, 开始卡号_In, 终止卡号_In, 使用方式_In, 登记时间_In, 登记人_In, 剩余数量_In, 批次_In, 签字人_In,
     Decode(签字人_In, Null, Null + Sysdate, Sysdate));
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_消费卡领用记录_Insert;
/

--122812:焦博,2018-06-07,票据入库记录增加字段批次,票据领用记录增加字段入库ID
CREATE OR REPLACE Procedure ZL_票据领用记录_UPDATE
(
  Id_In       In 票据领用记录.Id%Type,
  使用类别_In In 票据领用记录.使用类别%Type,
  领用人_In   In 票据领用记录.领用人%Type,
  开始号码_In In 票据领用记录.开始号码%Type,
  终止号码_In In 票据领用记录.终止号码%Type,
  前缀文本_In In 票据领用记录.前缀文本%Type := Null,
  使用方式_In In 票据领用记录.使用方式%Type := 1,
  登记时间_In In 票据领用记录.登记时间%Type := Null,
  登记人_In   In 票据领用记录.登记人%Type := Null,
  批次_In     In 票据领用记录.批次%Type := Null,
  签字人_In   In 票据领用记录.签字人%Type := Null,
  入库id_In   In 票据领用记录.入库id%Type := Null
  
) Is
  Cursor c_领用记录 Is
    Select * From 票据领用记录 Where ID = Id_In For Update;

  c_记录     票据领用记录%RowType;
  n_使用数量 票据领用记录.剩余数量%Type;
  n_剩余数量 票据领用记录.剩余数量%Type;
  n_原领用数 票据领用记录.剩余数量%Type;
  n_现领用数 票据领用记录.剩余数量%Type;

  v_开始号码 票据领用记录.开始号码%Type;
  v_终止号码 票据领用记录.终止号码%Type;
  v_Err_Msg  Varchar2(500);
  Err_Item Exception;
  n_Count  Number(18);
  n_剩余数 票据入库记录.剩余数量%Type;
Begin
  Open c_领用记录;
  Fetch c_领用记录
    Into c_记录;

  If c_领用记录%NotFound Then
    --记录未找到 
    v_Err_Msg := '[ZLSOFT]该条记录已经被删除，不能修改。[ZLSOFT]';
    Raise Err_Item;
  End If;

  Select Min(号码), Max(号码) Into v_开始号码, v_终止号码 From 票据使用明细 Where 领用id = Id_In;

  If 前缀文本_In Is Null Then
    n_剩余数量 := To_Number(终止号码_In) - To_Number(开始号码_In) + 1;
  Else
    n_剩余数量 := To_Number(Substr(终止号码_In, Length(前缀文本_In) + 1)) - To_Number(Substr(开始号码_In, Length(前缀文本_In) + 1)) + 1;
  End If;

  n_现领用数 := n_剩余数量;
  If c_记录.前缀文本 Is Null Then
    n_原领用数 := To_Number(c_记录.终止号码) - To_Number(c_记录.开始号码) + 1;
  Else
    n_原领用数 := To_Number(Substr(c_记录.终止号码, Length(c_记录.前缀文本) + 1)) - To_Number(Substr(c_记录.开始号码, Length(c_记录.前缀文本) + 1)) + 1;
  End If;

  If v_开始号码 Is Not Null Then
    --已经使用，对一些项目进行验证 
    If Nvl(前缀文本_In, ' ') <> Nvl(c_记录.前缀文本, ' ') Then
      v_Err_Msg := '[ZLSOFT]该条记录领用的票据已经使用，不能修改号码的前缀。[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    If Length(开始号码_In) <> Length(c_记录.开始号码) Then
      v_Err_Msg := '[ZLSOFT]该条记录领用的票据已经使用，不能修改号码的长度。[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    If 开始号码_In > v_开始号码 Then
      v_Err_Msg := '[ZLSOFT]该条记录领用的票据已经使用，开始号码最大只能是' || v_开始号码 || '。[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    If 终止号码_In < v_终止号码 Then
      v_Err_Msg := '[ZLSOFT]该条记录领用的票据已经使用，终止号码最小只能是' || v_终止号码 || '。[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    --下面计算数量 
    If 前缀文本_In Is Null Then
      n_使用数量 := To_Number(c_记录.终止号码) - To_Number(c_记录.开始号码) + 1 - c_记录.剩余数量;
    Else
      n_使用数量 := To_Number(Substr(c_记录.终止号码, Length(前缀文本_In) + 1)) - To_Number(Substr(c_记录.开始号码, Length(前缀文本_In) + 1)) + 1 -
                c_记录.剩余数量;
    End If;
  
    n_剩余数量 := n_剩余数量 - n_使用数量;
  End If;

  For v_入库 In (Select ID, 前缀文本, 使用类别, 开始号码, Nvl(终止号码, 开始号码) As 终止号码
               From 票据入库记录
               Where ID = Nvl(入库id_In, 0) And 票种 = c_记录.票种) Loop
  
    If Nvl(使用类别_In, 'LXH') <> Nvl(v_入库.使用类别, 'LXH') Then
      v_Err_Msg := '[ZLSOFT]当前领用的使用类别『' || Nvl(使用类别_In, '') || '』与入库的使用类别不一致『' || Nvl(v_入库.使用类别, '') || '』![ZLSOFT]';
      Raise Err_Item;
    End If;
  
    --1. 入库检查 
    If 开始号码_In < v_入库.开始号码 Or 开始号码_In > v_入库.终止号码 Then
      --不在入库范围,不能领用 
      v_Err_Msg := '[ZLSOFT]当前领用的开始号码『' || 开始号码_In || '』不在入库范围' || Chr(10) || Chr(13) || '『' || v_入库.开始号码 || '-' ||
                   v_入库.终止号码 || '』不能领用该票据![ZLSOFT]';
      Raise Err_Item;
    End If;
    If 终止号码_In < v_入库.开始号码 Or 终止号码_In > v_入库.终止号码 Then
      --不在入库范围,不能领用 
      v_Err_Msg := '[ZLSOFT]当前领用的终止号码『' || 终止号码_In || '』不在入库范围' || Chr(10) || Chr(13) || '『' || v_入库.开始号码 || '-' ||
                   v_入库.终止号码 || '』不能领用该票据![ZLSOFT]';
      Raise Err_Item;
    End If;
    --2.检查票据是否已经被报损,不能重复报损 
    Select Count(*)
    Into n_Count
    From 票据报损记录
    Where 入库id = Nvl(入库id_In, 0) And ((开始号码_In Between 开始号码 And 终止号码) Or (终止号码_In Between 开始号码 And 终止号码) Or
          (开始号码 Between 开始号码_In And 终止号码_In) Or (终止号码 Between 开始号码_In And 终止号码_In));
  
    If n_Count <> 0 Then
      If 开始号码_In = 终止号码_In Then
        v_Err_Msg := '[ZLSOFT]号码:' || 开始号码_In || '在报损记录中已经存在,不能再进行领用![ZLSOFT]';
      Else
        v_Err_Msg := '[ZLSOFT]号码:' || 开始号码_In || '-' || 终止号码_In || '在报损记录中已经存在,不能再进行领用![ZLSOFT]';
      End If;
      Raise Err_Item;
    End If;
  
    --3.检查票据是否已经被领用,领用的不能再进行报损 
    Select Count(*)
    Into n_Count
    From 票据领用记录
    Where 批次 = Nvl(批次_In, 0) And ID <> Id_In And
          ((开始号码_In Between 开始号码 And 终止号码) Or (终止号码_In Between 开始号码 And 终止号码) Or (开始号码 Between 开始号码_In And 终止号码_In) Or
          (终止号码 Between 开始号码_In And 终止号码_In));
    If n_Count <> 0 Then
      If 开始号码_In = 终止号码_In Then
        v_Err_Msg := '[ZLSOFT]号码:' || 开始号码_In || '在票据领用记录中已经存在,不能再进行领用![ZLSOFT]';
      Else
        v_Err_Msg := '[ZLSOFT]号码:' || 开始号码_In || '-' || 终止号码_In || '在票据领用记录中已经存在,不能再进行领用![ZLSOFT]';
      End If;
      Raise Err_Item;
    End If;
  
    --减少库存 
    Update 票据入库记录
    Set 剩余数量 = Nvl(剩余数量, 0) + (Nvl(n_原领用数, 0) - Nvl(n_现领用数, 0))
    Where ID = Nvl(入库id_In, 0) And 票种 = c_记录.票种
    Returning Nvl(剩余数量, 0) Into n_剩余数;
    If n_剩余数 < 0 Then
      v_Err_Msg := '[ZLSOFT]入库票据的剩余票据数不足,请检查![ZLSOFT]';
      Raise Err_Item;
    End If;
    Update 票据入库记录
    Set 有无票据 = Decode(Sign(Nvl(n_剩余数, 0)), 1, 1, Null)
    Where ID = (Select 入库id From 票据领用记录 Where ID = Id_In) And 票种 = c_记录.票种;
  End Loop;

  Update 票据领用记录
  Set 领用人 = 领用人_In, 前缀文本 = 前缀文本_In, 开始号码 = 开始号码_In, 终止号码 = 终止号码_In, 使用方式 = 使用方式_In, 登记时间 = 登记时间_In, 登记人 = 登记人_In,
      剩余数量 = n_剩余数量, 批次 = 批次_In, 使用类别 = 使用类别_In, 签字人 = 签字人_In, 签字时间 = Decode(签字人_In, Null, Null + Sysdate, Sysdate)
  Where ID = Id_In;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_票据领用记录_Update;
/

--122812:焦博,2018-06-07,票据入库记录增加字段批次,票据领用记录增加字段入库ID
CREATE OR REPLACE Procedure Zl_消费卡领用记录_Update
(
  Id_In       消费卡领用记录.Id%Type,
  接口编号_In 消费卡领用记录.接口编号%Type,
  领用人_In   消费卡领用记录.领用人%Type,
  开始卡号_In 消费卡领用记录.开始卡号%Type,
  终止卡号_In 消费卡领用记录.终止卡号%Type,
  前缀文本_In 消费卡领用记录.前缀文本%Type := Null,
  使用方式_In 消费卡领用记录.使用方式%Type := 1,
  登记时间_In 消费卡领用记录.登记时间%Type := Null,
  登记人_In   消费卡领用记录.登记人%Type := Null,
  批次_In     消费卡领用记录.批次%Type := Null,
  签字人_In   消费卡领用记录.签字人%Type := Null,
  入库id_In   消费卡领用记录.入库id%Type := Null
) Is
  Cursor c_领用记录 Is
    Select 接口编号, 前缀文本, 开始卡号, 终止卡号, 剩余数量 From 消费卡领用记录 Where ID = Id_In For Update;

  c_记录     c_领用记录%RowType;
  n_剩余数量 消费卡领用记录.剩余数量%Type;
  n_原领用数 消费卡领用记录.剩余数量%Type;
  n_现领用数 消费卡领用记录.剩余数量%Type;

  v_开始卡号 消费卡领用记录.开始卡号%Type;
  v_终止卡号 消费卡领用记录.终止卡号%Type;
  n_剩余数   消费卡入库记录.剩余数量%Type;
  n_Count    Number(18);

  v_Err_Msg Varchar2(500);
  Err_Item Exception;
Begin
  Open c_领用记录;
  Fetch c_领用记录
    Into c_记录;

  If c_领用记录%NotFound Then
    --记录未找到
    v_Err_Msg := '该条记录已经被删除，不能修改。';
    Raise Err_Item;
  End If;

  Select Min(卡号), Max(卡号) Into v_开始卡号, v_终止卡号 From 消费卡使用记录 Where 领用id = Id_In;

  If 前缀文本_In Is Null Then
    n_剩余数量 := To_Number(终止卡号_In) - To_Number(开始卡号_In) + 1;
  Else
    n_剩余数量 := To_Number(Substr(终止卡号_In, Length(前缀文本_In) + 1)) - To_Number(Substr(开始卡号_In, Length(前缀文本_In) + 1)) + 1;
  End If;

  n_现领用数 := n_剩余数量;
  If c_记录.前缀文本 Is Null Then
    n_原领用数 := To_Number(c_记录.终止卡号) - To_Number(c_记录.开始卡号) + 1;
  Else
    n_原领用数 := To_Number(Substr(c_记录.终止卡号, Length(c_记录.前缀文本) + 1)) - To_Number(Substr(c_记录.开始卡号, Length(c_记录.前缀文本) + 1)) + 1;
  End If;

  If v_开始卡号 Is Not Null Then
    --已经使用，对一些项目进行验证
    If Nvl(前缀文本_In, '-') <> Nvl(c_记录.前缀文本, '-') Then
      v_Err_Msg := '该条记录领用的卡号已经使用，不能修改卡号的前缀。';
      Raise Err_Item;
    End If;
  
    If Length(开始卡号_In) <> Length(c_记录.开始卡号) Then
      v_Err_Msg := '该条记录领用的卡号已经使用，不能修改卡号的长度。';
      Raise Err_Item;
    End If;
  
    If 开始卡号_In > v_开始卡号 Then
      v_Err_Msg := '该条记录领用的卡号已经使用，开始卡号最大只能是' || v_开始卡号 || '。';
      Raise Err_Item;
    End If;
  
    If 终止卡号_In < v_终止卡号 Then
      v_Err_Msg := '该条记录领用的卡号已经使用，终止卡号最小只能是' || v_终止卡号 || '。';
      Raise Err_Item;
    End If;
  
    --下面计算数量
    n_剩余数量 := n_剩余数量 - (n_原领用数 - c_记录.剩余数量);
  End If;

  For v_入库 In (Select ID, 前缀文本, 接口编号, 开始卡号, Nvl(终止卡号, 开始卡号) As 终止卡号
               From 消费卡入库记录
               Where ID = Nvl(入库id_In, 0)) Loop
  
    If 接口编号_In <> v_入库.接口编号 Then
      v_Err_Msg := '当前领用的卡类别『' || Nvl(接口编号_In, '') || '』与入库的类别不一致『' || Nvl(v_入库.接口编号, '') || '』!';
      Raise Err_Item;
    End If;
  
    --1. 入库检查
    If 开始卡号_In < v_入库.开始卡号 Or 开始卡号_In > v_入库.终止卡号 Then
      --不在入库范围,不能领用
      v_Err_Msg := '当前领用的开始卡号『' || 开始卡号_In || '』不在入库范围' || Chr(10) || Chr(13) || '『' || v_入库.开始卡号 || '-' || v_入库.终止卡号 ||
                   '』不能领用该卡片！';
      Raise Err_Item;
    End If;
    If 终止卡号_In < v_入库.开始卡号 Or 终止卡号_In > v_入库.终止卡号 Then
      --不在入库范围,不能领用
      v_Err_Msg := '当前领用的终止卡号『' || 终止卡号_In || '』不在入库范围' || Chr(10) || Chr(13) || '『' || v_入库.开始卡号 || '-' || v_入库.终止卡号 ||
                   '』不能领用该卡片！';
      Raise Err_Item;
    End If;
    --2.检查卡号是否已经被报损,不能重复报损
    Select Count(1)
    Into n_Count
    From 消费卡报损记录
    Where 入库id = Nvl(入库id_In, 0) And ((开始卡号_In Between 开始卡号 And 终止卡号) Or (终止卡号_In Between 开始卡号 And 终止卡号) Or
          (开始卡号 Between 开始卡号_In And 终止卡号_In) Or (终止卡号 Between 开始卡号_In And 终止卡号_In));
  
    If n_Count <> 0 Then
      If 开始卡号_In = 终止卡号_In Then
        v_Err_Msg := '卡号:' || 开始卡号_In || '在报损记录中已经存在，不能再进行领用！';
      Else
        v_Err_Msg := '卡号:' || 开始卡号_In || '-' || 终止卡号_In || '在报损记录中已经存在，不能再进行领用！';
      End If;
      Raise Err_Item;
    End If;
  
    --3.检查卡号是否已经被领用,领用的不能再进行报损
    Select Count(*)
    Into n_Count
    From 消费卡领用记录
    Where 批次 = Nvl(批次_In, 0) And ID <> Id_In And
          ((开始卡号_In Between 开始卡号 And 终止卡号) Or (终止卡号_In Between 开始卡号 And 终止卡号) Or (开始卡号 Between 开始卡号_In And 终止卡号_In) Or
          (终止卡号 Between 开始卡号_In And 终止卡号_In));
    If n_Count <> 0 Then
      If 开始卡号_In = 终止卡号_In Then
        v_Err_Msg := '卡号:' || 开始卡号_In || '在消费卡领用记录中已经存在，不能再进行领用！';
      Else
        v_Err_Msg := '卡号:' || 开始卡号_In || '-' || 终止卡号_In || '在消费卡领用记录中已经存在，不能再进行领用！';
      End If;
      Raise Err_Item;
    End If;
  
    --减少库存
    Update 消费卡入库记录
    Set 剩余数量 = Nvl(剩余数量, 0) + (Nvl(n_原领用数, 0) - Nvl(n_现领用数, 0))
    Where ID = (Select 入库id From 消费卡领用记录 Where ID = Id_In)
    Returning Nvl(剩余数量, 0) Into n_剩余数;
    If n_剩余数 < 0 Then
      v_Err_Msg := '入库卡片的剩余票据数不足，请检查！';
      Raise Err_Item;
    End If;
  
    Update 消费卡入库记录 Set 是否存在卡 = Decode(Sign(Nvl(n_剩余数, 0)), 1, 1, Null) Where ID = Nvl(入库id_In, 0);
  End Loop;

  Update 消费卡领用记录
  Set 领用人 = 领用人_In, 前缀文本 = 前缀文本_In, 开始卡号 = 开始卡号_In, 终止卡号 = 终止卡号_In, 使用方式 = 使用方式_In, 登记时间 = 登记时间_In, 登记人 = 登记人_In,
      剩余数量 = n_剩余数量, 批次 = Nvl(批次_In, 0), 接口编号 = 接口编号_In, 签字人 = 签字人_In,
      签字时间 = Decode(签字人_In, Null, Null + Sysdate, Sysdate)
  Where ID = Id_In;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_消费卡领用记录_Update;
/

CREATE OR REPLACE Procedure Zl_Ris检查预约_Delete(医嘱id_In In Ris检查预约.医嘱id%Type) Is
  v_预约id       Ris检查预约.预约id%Type;
  v_预约日期     Ris检查预约.预约日期%Type;
  v_预约序号     Ris检查预约.序号%Type;
  v_检查设备名称 Ris检查预约.检查设备名称%Type;
Begin
  v_预约id := 0;
  Begin
    Select 预约id, 预约日期, 序号, 检查设备名称
    Into v_预约id, v_预约日期, v_预约序号, v_检查设备名称
    From Ris检查预约
    Where 医嘱id = 医嘱id_In;
  Exception
    When Others Then
      Null;
  End;

  Delete Ris检查预约 Where 医嘱id = 医嘱id_In;

  --发送消息
  If v_预约id <> 0 Then
    b_Message.Zlhis_Pacs_007(医嘱id_In, v_预约id, v_预约日期, v_预约序号, v_检查设备名称);
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Ris检查预约_Delete;
/





------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0015' Where 编号=&n_System;
Commit;