----------------------------------------------------------------------------------------------------------------
--本脚本支持从ZLHIS+ v10.35.90升级到 v10.35.90
--请以数据空间的所有者登录PLSQL并执行下列脚本
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------



------------------------------------------------------------------------------
--结构修正部份
------------------------------------------------------------------------------
--129973:秦龙,2018-09-10,增加属性“是否易至跌倒”
Alter Table 药品规格 Add 是否易至跌倒 number(1);

--130300:秦龙,2018-09-11,增加属性“严格控制用法用量”
Alter Table 药品规格 Add 严格控制用法用量 number(1);
Alter Table 药品特性 Add 严格控制用法用量 number(1);

------------------------------------------------------------------------------
--数据修正部份
------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--128798:蒋廷中,2018-09-14,Zl_病人医嘱执行_Insert过程提交不正确
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
  d_执行时间   Date;

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
    If n_执行状态 != 0 Then
      d_执行时间 := 执行时间_In;
    End If;
    If v_费用性质 = 2 Then
      If Nvl(单独执行_In, 0) = 1 Then
        Update 住院费用记录 A
        Set 执行状态 = n_执行状态, 执行人 = Decode(n_执行状态, 0, Null, 执行人_In), 执行时间 = d_执行时间
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And Not Exists
         (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1) And a.记录状态 In (0, 1, 3) And
              (医嘱序号, NO, 记录性质) In
              (Select 医嘱id, NO, 记录性质
               From 病人医嘱发送
               Where 执行状态 = Decode(n_执行次数, 0, 0, 3) And 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In);
      Else
        Update 住院费用记录 A
        Set 执行状态 = n_执行状态, 执行人 = Decode(n_执行状态, 0, Null, 执行人_In), 执行时间 = d_执行时间
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And Not Exists
         (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1) And a.记录状态 In (0, 1, 3) And
              (医嘱序号, NO, 记录性质) In (Select 医嘱id, NO, 记录性质
                                   From 病人医嘱发送
                                   Where 执行状态 = Decode(n_执行次数, 0, 0, 3) And 发送号 + 0 = 发送号_In And
                                         医嘱id In (Select ID
                                                  From 病人医嘱记录
                                                  Where ID = v_组id And 诊疗类别 = v_诊疗类别
                                                  Union All
                                                  Select ID
                                                  From 病人医嘱记录
                                                  Where 相关id = v_组id And 诊疗类别 = v_诊疗类别));
      End If;
    Else
      If Nvl(单独执行_In, 0) = 1 Then
        --对于门诊单据n_执行状态可能为0（登记执行情况，选择执行结果为未执行），因此需判断
        Update 门诊费用记录 A
        Set 执行状态 = n_执行状态, 执行人 = Decode(n_执行状态, 0, Null, 执行人_In), 执行时间 = d_执行时间
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And Not Exists
         (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1) And a.记录状态 In (0, 1, 3) And
              (医嘱序号, NO, 记录性质) In
              (Select 医嘱id, NO, 记录性质
               From 病人医嘱发送
               Where 执行状态 = Decode(n_执行次数, 0, 0, 3) And 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In);
      Else
        Update 门诊费用记录 A
        Set 执行状态 = n_执行状态, 执行人 = Decode(n_执行状态, 0, Null, 执行人_In), 执行时间 = d_执行时间
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And Not Exists
         (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1) And a.记录状态 In (0, 1, 3) And
              (医嘱序号, NO, 记录性质) In (Select 医嘱id, NO, 记录性质
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


--130541:胡俊勇,2018-09-14,集成平台消息问题
Create Or Replace Package b_Message Is
  Procedure p_Msg_Todo_Insert
  (
    Msg_Code_In  Zlmsg_Todo.Msg_Code%Type,
    Key_Value_In Zlmsg_Todo.Key_Value%Type
  );
  --设置平台调用类型
  Procedure Set_Platform_Call(Platform_Call Number);
  --新增部门
  Procedure Zlhis_Dict_001(Id_In 部门表.Id%Type);
  --修改部门
  Procedure Zlhis_Dict_002(部门id_In 部门表.Id%Type);
  --停用部门
  Procedure Zlhis_Dict_003(部门id_In 部门表.Id%Type);
  --启用部门
  Procedure Zlhis_Dict_004(部门id_In 部门表.Id%Type);
  --新增人员
  Procedure Zlhis_Dict_005(人员id_In 人员表.Id%Type);
  --修改人员
  Procedure Zlhis_Dict_006(人员id_In 人员表.Id%Type);
  --停用人员
  Procedure Zlhis_Dict_007(人员id_In 人员表.Id%Type);
  --启用人员
  Procedure Zlhis_Dict_008(人员id_In 人员表.Id%Type);
  --新增收费项目
  Procedure Zlhis_Dict_009(细目id_In 收费项目目录.Id%Type);
  --修改收费项目
  Procedure Zlhis_Dict_010(细目id_In 收费项目目录.Id%Type);
  --停用收费项目
  Procedure Zlhis_Dict_011(细目id_In 收费项目目录.Id%Type);
  --启用收费项目
  Procedure Zlhis_Dict_012(细目id_In 收费项目目录.Id%Type);
  --新增诊疗项目
  Procedure Zlhis_Dict_013(诊疗id_In 诊疗项目目录.Id%Type);
  --修改诊疗项目
  Procedure Zlhis_Dict_014(诊疗id_In 诊疗项目目录.Id%Type);
  --停用诊疗项目
  Procedure Zlhis_Dict_015(诊疗id_In 诊疗项目目录.Id%Type);
  --启用诊疗项目
  Procedure Zlhis_Dict_016(诊疗id_In 诊疗项目目录.Id%Type);
  --新增检验项目
  Procedure Zlhis_Dict_017(诊疗id_In 诊疗项目目录.Id%Type);
  --修改检验项目
  Procedure Zlhis_Dict_018(诊疗id_In 诊疗项目目录.Id%Type);
  --删除检验项目
  Procedure Zlhis_Dict_019
  (
    诊疗id_In 诊疗项目目录.Id%Type,
    编码_In   诊治所见项目.编码%Type,
    中文名_In 诊治所见项目.中文名%Type,
    英文名_In 诊治所见项目.英文名%Type
  );

  --新增疾病编码目录
  Procedure Zlhis_Dict_021(疾病id_In 疾病编码目录.Id%Type);
  --修改疾病编码目录
  Procedure Zlhis_Dict_022(疾病id_In 疾病编码目录.Id%Type);
  --停用疾病编码目录
  Procedure Zlhis_Dict_023(疾病id_In 疾病编码目录.Id%Type);
  --启用疾病编码目录
  Procedure Zlhis_Dict_024(疾病id_In 疾病编码目录.Id%Type);
  --新增药品分类
  Procedure Zlhis_Dict_025
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  );
  --修改药品分类
  Procedure Zlhis_Dict_026
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  );
  --删除药品分类
  Procedure Zlhis_Dict_027
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type,
    编码_In 诊疗分类目录.编码%Type,
    名称_In 诊疗分类目录.名称%Type
  );
  --停用药品分类
  Procedure Zlhis_Dict_028
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  );
  --启用药品分类
  Procedure Zlhis_Dict_029
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  );
  --新增药品品种
  Procedure Zlhis_Dict_030
  (
    类别_In 诊疗项目目录.类别%Type,
    Id_In   诊疗项目目录.Id%Type
  );
  --修改药品品种
  Procedure Zlhis_Dict_031
  (
    类别_In 诊疗项目目录.类别%Type,
    Id_In   诊疗项目目录.Id%Type
  );
  --删除药品品种
  Procedure Zlhis_Dict_032
  (
    类别_In 诊疗项目目录.类别%Type,
    Id_In   诊疗项目目录.Id%Type,
    编码_In 诊疗项目目录.编码%Type,
    名称_In 诊疗项目目录.名称%Type
  );
  --停用药品品种
  Procedure Zlhis_Dict_033
  (
    类别_In 诊疗项目目录.类别%Type,
    Id_In   诊疗项目目录.Id%Type
  );
  --启用药品品种
  Procedure Zlhis_Dict_034
  (
    类别_In 诊疗项目目录.类别%Type,
    Id_In   诊疗项目目录.Id%Type
  );
  --新增药品规格
  Procedure Zlhis_Dict_035
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  );
  --修改药品规格
  Procedure Zlhis_Dict_036
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  );
  --删除药品规格
  Procedure Zlhis_Dict_037
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type,
    编码_In   收费项目目录.编码%Type,
    名称_In   收费项目目录.名称%Type,
    规格_In   收费项目目录.规格%Type,
    产地_In   收费项目目录.产地%Type
  );
  --停用药品规格
  Procedure Zlhis_Dict_038
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  );
  --启用药品规格
  Procedure Zlhis_Dict_039
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  );
  --设置药品存储库房
  Procedure Zlhis_Dict_040
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  );
  --设置药品储备限额
  Procedure Zlhis_Dict_041
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  );
  --新增卫材品种
  Procedure Zlhis_Dict_042(Id_In 诊疗项目目录.Id%Type);
  --新增卫材规格
  Procedure Zlhis_Dict_043(Id_In 收费项目目录.Id%Type);
  --修改卫材规格
  Procedure Zlhis_Dict_044(Id_In 收费项目目录.Id%Type);
  --删除卫材规格
  Procedure Zlhis_Dict_045
  (
    Id_In   收费项目目录.Id%Type,
    编码_In 收费项目目录.编码%Type,
    名称_In 收费项目目录.名称%Type,
    规格_In 收费项目目录.规格%Type,
    产地_In 收费项目目录.产地%Type
  );
  --停用卫材规格
  Procedure Zlhis_Dict_046(Id_In 收费项目目录.Id%Type);
  --启用卫材规格
  Procedure Zlhis_Dict_047(Id_In 收费项目目录.Id%Type);
  --医保对码
  Procedure Zlhis_Dict_048
  (
    险类_In       In 保险支付项目.险类%Type,
    收费细目id_In In 保险支付项目.收费细目id%Type
  );
  --删除医保对码
  Procedure Zlhis_Dict_049
  (
    险类_In       In 保险支付项目.险类%Type,
    收费细目id_In In 保险支付项目.收费细目id%Type,
    项目编码_In   In 收费项目目录.编码%Type,
    项目名称_In   In 收费项目目录.名称%Type,
    医保编码_In   In 保险支付项目.项目编码%Type,
    医保名称_In   In 保险支付项目.项目名称%Type
  );
  --新增卫材分类
  Procedure Zlhis_Dict_050
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  );
  --修改卫材分类
  Procedure Zlhis_Dict_051
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  );
  --删除卫材分类
  Procedure Zlhis_Dict_052
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type,
    编码_In 诊疗分类目录.编码%Type,
    名称_In 诊疗分类目录.名称%Type
  );
  --收费价目变动
  Procedure Zlhis_Dict_053
  (
    收费项目Id_In       收费项目目录.Id%Type
  );
  --诊疗收费对照变动
  Procedure Zlhis_Dict_054
  (
    诊疗项目Id_In     诊疗分类目录.Id%Type
  );
  --新增诊疗检查类型
  Procedure Zlhis_Dictpacs_001
  (
    编码_In   诊疗检查类型.编码%Type,
    名称_In   诊疗检查类型.名称%Type,
    简码_In   诊疗检查类型.简码%Type,
    建病案_In 诊疗检查类型.建病案%Type
  );
  --修改诊疗检查类型
  Procedure Zlhis_Dictpacs_002
  (
    编码_In   诊疗检查类型.编码%Type,
    名称_In   诊疗检查类型.名称%Type,
    简码_In   诊疗检查类型.简码%Type,
    建病案_In 诊疗检查类型.建病案%Type
  );
  --删除诊疗检查类型
  Procedure Zlhis_Dictpacs_003
  (
    编码_In   诊疗检查类型.编码%Type,
    名称_In   诊疗检查类型.名称%Type,
    简码_In   诊疗检查类型.简码%Type,
    建病案_In 诊疗检查类型.建病案%Type
  );
  --新增诊疗检查部位
  Procedure Zlhis_Dictpacs_004
  (
    类型_In     诊疗检查部位.类型%Type,
    编码_In     诊疗检查部位.编码%Type,
    名称_In     诊疗检查部位.名称%Type,
    分组_In     诊疗检查部位.分组%Type,
    备注_In     诊疗检查部位.备注%Type,
    方法_In     诊疗检查部位.方法%Type,
    适用性别_In 诊疗检查部位.适用性别%Type
  );
  --修改诊疗检查部位
  Procedure Zlhis_Dictpacs_005
  (
    类型_In     诊疗检查部位.类型%Type,
    编码_In     诊疗检查部位.编码%Type,
    名称_In     诊疗检查部位.名称%Type,
    分组_In     诊疗检查部位.分组%Type,
    备注_In     诊疗检查部位.备注%Type,
    方法_In     诊疗检查部位.方法%Type,
    适用性别_In 诊疗检查部位.适用性别%Type
  );
  --删除诊疗检查部位
  Procedure Zlhis_Dictpacs_006
  (
    类型_In     诊疗检查部位.类型%Type,
    编码_In     诊疗检查部位.编码%Type,
    名称_In     诊疗检查部位.名称%Type,
    分组_In     诊疗检查部位.分组%Type,
    备注_In     诊疗检查部位.备注%Type,
    方法_In     诊疗检查部位.方法%Type,
    适用性别_In 诊疗检查部位.适用性别%Type
  );
  --新增诊疗项目部位
  Procedure Zlhis_Dictpacs_007
  (
    Id_In     诊疗项目部位.Id%Type,
    项目id_In 诊疗项目部位.项目id%Type,
    类型_In   诊疗项目部位.类型%Type,
    部位_In   诊疗项目部位.部位%Type,
    方法_In   诊疗项目部位.方法%Type,
    默认_In   诊疗项目部位.默认%Type
  );
  --修改诊疗项目部位
  Procedure Zlhis_Dictpacs_008
  (
    Id_In     诊疗项目部位.Id%Type,
    项目id_In 诊疗项目部位.项目id%Type,
    类型_In   诊疗项目部位.类型%Type,
    部位_In   诊疗项目部位.部位%Type,
    方法_In   诊疗项目部位.方法%Type,
    默认_In   诊疗项目部位.默认%Type
  );
  --删除诊疗项目部位
  Procedure Zlhis_Dictpacs_009
  (
    Id_In     诊疗项目部位.Id%Type,
    项目id_In 诊疗项目部位.项目id%Type,
    类型_In   诊疗项目部位.类型%Type,
    部位_In   诊疗项目部位.部位%Type,
    方法_In   诊疗项目部位.方法%Type,
    默认_In   诊疗项目部位.默认%Type
  );
  --新增诊疗检验标本
  Procedure Zlhis_Dictlis_004
  (
    编码_In     诊疗检验标本.编码%Type,
    名称_In     诊疗检验标本.名称%Type,
    简码_In     诊疗检验标本.简码%Type,
    适用性别_In 诊疗检验标本.适用性别%Type
  );
  --修改诊疗检验标本
  Procedure Zlhis_Dictlis_005
  (
    编码_In     诊疗检验标本.编码%Type,
    名称_In     诊疗检验标本.名称%Type,
    简码_In     诊疗检验标本.简码%Type,
    适用性别_In 诊疗检验标本.适用性别%Type
  );
  --删除诊疗项目部位
  Procedure Zlhis_Dictlis_006
  (
    编码_In     诊疗检验标本.编码%Type,
    名称_In     诊疗检验标本.名称%Type,
    简码_In     诊疗检验标本.简码%Type,
    适用性别_In 诊疗检验标本.适用性别%Type
  );
  --新增采血管类型
  Procedure Zlhis_Dictlis_007
  (
    编码_In   采血管类型.编码%Type,
    名称_In   采血管类型.名称%Type,
    简码_In   采血管类型.简码%Type,
    添加剂_In 采血管类型.添加剂%Type,
    采血量_In 采血管类型.采血量%Type,
    规格_In   采血管类型.规格%Type,
    颜色_In   采血管类型.颜色%Type,
    材料id_In 采血管类型.材料id%Type
  );
  --修改采血管类型
  Procedure Zlhis_Dictlis_008
  (
    编码_In   采血管类型.编码%Type,
    名称_In   采血管类型.名称%Type,
    简码_In   采血管类型.简码%Type,
    添加剂_In 采血管类型.添加剂%Type,
    采血量_In 采血管类型.采血量%Type,
    规格_In   采血管类型.规格%Type,
    颜色_In   采血管类型.颜色%Type,
    材料id_In 采血管类型.材料id%Type
  );
  --删除采血管类型
  Procedure Zlhis_Dictlis_009
  (
    编码_In   采血管类型.编码%Type,
    名称_In   采血管类型.名称%Type,
    简码_In   采血管类型.简码%Type,
    添加剂_In 采血管类型.添加剂%Type,
    采血量_In 采血管类型.采血量%Type,
    规格_In   采血管类型.规格%Type,
    颜色_In   采血管类型.颜色%Type,
    材料id_In 采血管类型.材料id%Type
  );

  --药品备药发送
  Procedure Zlhis_Drug_001(No_In 药品收发记录.No%Type);
  --取消药品备药发送
  Procedure Zlhis_Drug_002(No_In 药品收发记录.No%Type);
  --药品移库单接收
  Procedure Zlhis_Drug_003(No_In 药品收发记录.No%Type);
  --药品移库单冲销
  Procedure Zlhis_Drug_004(No_In 药品收发记录.No%Type);

  --部门发药
  Procedure Zlhis_Drug_005
  (
    库房id_In 药品收发记录.库房id%Type,
    收发id_In 药品收发记录.Id%Type
  );
  --部门退药
  Procedure Zlhis_Drug_006
  (
    冲销收发id_In 药品收发记录.Id%Type,
    待发收发id_In 药品收发记录.Id%Type,
    数量_In       药品收发记录.实际数量%Type,
    费用id_In     门诊费用记录.Id%Type
  );
  --药品调价
  Procedure Zlhis_Drug_007(价格id_In 药品价格记录.Id%Type);
  --静配发送
  Procedure Zlhis_Drug_008(记录ids_In Varchar2);
  --药品调售价
  Procedure Zlhis_Drug_009
  (
    价格id_In 药品价格记录.Id%Type,
    时价_In   Number
  );
  --卫材调成本价
  Procedure Zlhis_Drug_010(价格id_In 成本价调价信息.Id%Type);
  --卫材调售价
  Procedure Zlhis_Drug_011
  (
    价格id_In 收费价目.Id%Type,
    时价_In   Number
  );
  --2.停止患者医嘱，住院
  Procedure Zlhis_Cis_002
  (
    病人id_In  In 病人医嘱记录.病人id%Type,
    主页id_In  In 病人医嘱记录.主页id%Type,
    医嘱id_In  In 病人医嘱记录.Id%Type,
    医嘱ids_In In Varchar2
  );
  --3.作废患者医嘱，门诊/住院
  Procedure Zlhis_Cis_003
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );

  --4.患者术后医嘱，住院
  Procedure Zlhis_Cis_004
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );

  --5.撤消患者术后医嘱，住院
  Procedure Zlhis_Cis_005
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );

  --6.患者护理常规医嘱，住院
  Procedure Zlhis_Cis_006
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );

  --7.撤消患者护理常规医嘱，住院
  Procedure Zlhis_Cis_007
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  );

  --门诊患者接诊
  Procedure Zlhis_Cis_008
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    挂号单_In In 病人挂号记录.No%Type
  );

  --门诊患者取消接诊
  Procedure Zlhis_Cis_009
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    挂号单_In In 病人挂号记录.No%Type
  );

  --10.下达患者诊断，门诊/住院
  Procedure Zlhis_Cis_010
  (
    病人id_In In 病人挂号记录.病人id%Type,
    就诊id_In In 病人挂号记录.Id%Type, --门诊病人 挂号ID，住院病人 主页ID
    诊断id_In In 病人诊断记录.Id%Type
  );
  --11.撤消患者诊断
  Procedure Zlhis_Cis_011
  (
    病人id_In   In 病人挂号记录.病人id%Type,
    就诊id_In   In 病人挂号记录.Id%Type, --门诊病人 挂号ID，住院病人 主页ID
    Id_In       In 病人诊断记录.Id%Type,
    疾病id_In   In 病人诊断记录.疾病id%Type,
    诊断id_In   In 病人诊断记录.诊断id%Type,
    诊断描述_In In 病人诊断记录.诊断描述%Type
  );

  --病区执行医嘱校对
  Procedure Zlhis_Cis_012
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );

  --13.检验危急值阅读通知
  Procedure Zlhis_Cis_014
  (
    病人id_In In 病人挂号记录.病人id%Type,
    就诊id_In In 病人挂号记录.Id%Type, --门诊病人 挂号ID，住院病人 主页ID
    医嘱id_In In 病人医嘱记录.Id%Type,
    消息id_In In 业务消息清单.Id%Type
  );

  --15.患者检验申请，门诊/住院
  Procedure Zlhis_Cis_016
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    病人来源_In In 病人医嘱记录.病人来源%Type
  );
  --16.患者检查申请，门诊/住院
  Procedure Zlhis_Cis_017
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    病人来源_In In 病人医嘱记录.病人来源%Type
  );
  --17.患者手术申请，门诊/住院
  Procedure Zlhis_Cis_018
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );
  --18.患者输血申请，住院
  Procedure Zlhis_Cis_019
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );
  --19.患者会诊申请，住院
  Procedure Zlhis_Cis_020
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );
  --20.患者抢救医嘱，住院
  Procedure Zlhis_Cis_021
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );
  --21.患者死亡医嘱，住院
  Procedure Zlhis_Cis_022
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );
  --22.患者特殊治疗医嘱，住院
  Procedure Zlhis_Cis_023
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );

  --24.检查危急值阅读通知
  Procedure Zlhis_Cis_025
  (
    病人id_In In 病人挂号记录.病人id%Type,
    就诊id_In In 病人挂号记录.Id%Type, --门诊病人 挂号ID，住院病人 主页ID
    医嘱id_In In 病人医嘱记录.Id%Type,
    消息id_In In 业务消息清单.Id%Type
  );

  --病区执行医嘱发送
  Procedure Zlhis_Cis_026
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );

  --撤消患者检验申请
  Procedure Zlhis_Cis_036
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    No_In       In 病人医嘱发送.No%Type,
    病人来源_In In 病人医嘱记录.病人来源%Type
  );

  --撤消患者检查申请
  Procedure Zlhis_Cis_037
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    No_In       In 病人医嘱发送.No%Type,
    病人来源_In In 病人医嘱记录.病人来源%Type
  );

  --撤消患者手术申请
  Procedure Zlhis_Cis_038
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  );

  --撤消患者输血申请
  Procedure Zlhis_Cis_039
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  );

  --撤消患者会诊申请
  Procedure Zlhis_Cis_040
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  );

  --撤消患者抢救医嘱
  Procedure Zlhis_Cis_041
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  );

  --撤消患者死亡医嘱
  Procedure Zlhis_Cis_042
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  );

  --撤消特殊治疗医嘱
  Procedure Zlhis_Cis_043
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  );

  --撤消病区执行医嘱
  Procedure Zlhis_Cis_044
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    No_In       In 病人医嘱发送.No%Type,
    发送数次_In In 病人医嘱发送.发送数次%Type,
    首次时间_In In 病人医嘱发送.首次时间%Type,
    末次时间_In In 病人医嘱发送.末次时间%Type,
    样本条码_In In 病人医嘱发送.样本条码%Type
  );
  --患者医嘱执行登记
  Procedure Zlhis_Cis_050
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    要求时间_In In 病人医嘱执行.要求时间%Type,
    执行时间_In In 病人医嘱执行.执行时间%Type
  );

  --患者医嘱取消执行登记
  Procedure Zlhis_Cis_051
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    要求时间_In In 病人医嘱执行.要求时间%Type,
    执行时间_In In 病人医嘱执行.执行时间%Type,
    本次数次_In In 病人医嘱执行.本次数次%Type,
    执行结果_In In 病人医嘱执行.执行结果%Type,
    执行摘要_In In 病人医嘱执行.执行摘要%Type,
    执行科室_In In 病人医嘱执行.执行科室id%Type,
    执行人_In   In 病人医嘱执行.执行人%Type,
    核对人_In   In 病人医嘱执行.核对人%Type,
    记录来源_In In 病人医嘱执行.记录来源%Type
  );
  --患者医嘱执行完成
  Procedure Zlhis_Cis_052
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );
  --患者医嘱撤消执行完成
  Procedure Zlhis_Cis_053
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );
  --病理申请发送后修改
  Procedure Zlhis_Cis_056
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    病人来源_In In 病人医嘱记录.病人来源%Type
  );

  --门诊患者完成就诊
  Procedure Zlhis_Cis_057
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    挂号单_In In 病人挂号记录.No%Type
  );

  --门诊患者取消完成就诊
  Procedure Zlhis_Cis_058
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    挂号单_In In 病人挂号记录.No%Type
  );

  --确认停止患者医嘱 
  Procedure Zlhis_Cis_059
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );

  --26.检查报告完成，检查完成时
  Procedure Zlhis_Pacs_001
  (
    医嘱id_In   In 影像检查记录.医嘱id%Type,
    报告id_Ins  In Varchar2,
    报告类型_In In Number --1-老版PACS报告，2-老版病历编辑器报告，3-新版编辑器报告
  );
  --27.检查状态同步，检查状态改变后
  Procedure Zlhis_Pacs_002
  (
    医嘱id_In In 影像检查记录.医嘱id%Type,
    原状态_In In 病人医嘱发送.执行过程%Type,
    新状态_In In 病人医嘱发送.执行过程%Type
  );
  --28.检查状态回退，检查状态回退后
  Procedure Zlhis_Pacs_003
  (
    医嘱id_In In 影像检查记录.医嘱id%Type,
    原状态_In In 病人医嘱发送.执行过程%Type,
    新状态_In In 病人医嘱发送.执行过程%Type
  );
  --29.检查报告撤销，撤销检查完成时
  Procedure Zlhis_Pacs_004
  (
    医嘱id_In   In 影像检查记录.医嘱id%Type,
    报告id_Ins  In Varchar2,
    报告类型_In In Number --1-老版PACS报告，2-老版病历编辑器报告，3-新版编辑器报告
  );
  --30.检查危急值通知，检查发生危急值时
  Procedure Zlhis_Pacs_005(医嘱id_In In 影像检查记录.医嘱id%Type);
  -- 检查预约通知，检查预约时
  Procedure Zlhis_Pacs_006
  (
    医嘱id_In In 影像检查记录.医嘱id%Type,
    预约id_In In Ris检查预约.预约id%Type
  );
  -- 取消检查预约，取消预约时
  Procedure Zlhis_Pacs_007
  (
    医嘱id_In       In 影像检查记录.医嘱id%Type,
    预约id_In       In Ris检查预约.预约id%Type,
    预约日期_In     In Ris检查预约.预约日期%Type,
    预约序号_In     In Ris检查预约.序号%Type,
    检查设备名称_In In Ris检查预约.检查设备名称%Type
  );

  --36.患者发卡
  Procedure Zlhis_Patient_018
  (
    变动id_In   In 病人变动记录.Id%Type,
    病人id_In   In 病人信息.病人id%Type,
    卡类别id_In In 医疗卡类别.Id%Type,
    卡号_In     In 病人医疗卡信息.卡号%Type
  );

  --37.患者退卡
  Procedure Zlhis_Patient_019
  (
    变动id_In   In 病人变动记录.Id%Type,
    病人id_In   In 病人信息.病人id%Type,
    卡类别id_In In 医疗卡类别.Id%Type,
    卡号_In     In 病人医疗卡信息.卡号%Type
  );

  --38.患者退卡
  Procedure Zlhis_Patient_020
  (
    变动id_In   In 病人变动记录.Id%Type,
    病人id_In   In 病人信息.病人id%Type,
    卡类别id_In In 医疗卡类别.Id%Type,
    原卡号_In   In 病人医疗卡信息.卡号%Type,
    新卡号_In   In 病人医疗卡信息.卡号%Type
  );

  --39.病人挂号登记（包含预约登记)
  Procedure Zlhis_Regist_001
  (
    挂号id_In In 病人挂号记录.Id%Type,
    No_In     In 病人挂号记录.No%Type
  );

  --40.病人分诊
  Procedure Zlhis_Regist_002
  (
    挂号id_In In 病人挂号记录.Id%Type,
    No_In     In 病人挂号记录.No%Type,
    诊室_In   In 病人挂号记录.诊室%Type
  );

  --41.病人退号
  Procedure Zlhis_Regist_003
  (
    挂号id_In In 病人挂号记录.Id%Type,
    No_In     In 病人挂号记录.No%Type
  );

  --42.临床出诊安排调整
  Procedure Zlhis_Regist_004
  (
    变动原因_In In Integer, --1-停诊;2-替诊;3-诊室变动
    记录id_In   In 临床出诊记录.Id%Type,
    变动id_In   In 临床出诊变动记录.Id%Type
  );

  --43.门诊患者挂号换号操作
  Procedure Zlhis_Regist_005
  (
    No_In         In 病人挂号记录.No%Type,
    变动原因_In   Integer, --1-替诊;2-换号;3-预约日期变动,
    就诊变动id_In 就诊变动记录.Id%Type
  );

  --费用门诊收费及补充结算
  --结算类型_In:1-收费结算，2-补充结算
  Procedure Zlhis_Charge_002
  (
    结算类型_In In Number,
    结帐id_In   In 门诊费用记录.结帐id%Type
  );

  --46.门诊退费单据
  Procedure Zlhis_Charge_004
  (
    退费类型_In In Number,
    结帐id_In   In 门诊费用记录.结帐id%Type
  );

  --47.收预交款
  Procedure Zlhis_Charge_005
  (
    预交id_In In 病人预交记录.Id%Type,
    单据号_In In 病人预交记录.No%Type
  );

  --48.退预交款(包含负数退预交款部分)
  Procedure Zlhis_Charge_006
  (
    退预交id_In In 病人预交记录.Id%Type,
    单据号_In   In 病人预交记录.No%Type
  );

  --住院记帐单据
  Procedure Zlhis_Charge_007
  (
    收费类别_In In 住院费用记录.收费类别%Type,
    费用id_In   In 住院费用记录.Id%Type
  );

  --住院记帐单据销账
  Procedure Zlhis_Charge_008
  (
    收费类别_In In 住院费用记录.收费类别%Type,
    费用id_In   In 住院费用记录.Id%Type,
    收发ids_In  In Varchar2 := Null --可能费用ID对应多个收发id，对应格式：收发id,数量|收发id,数量；非药品不传
  );

  --53.住院患者入院登记
  Procedure Zlhis_Patient_001
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );
  --54.住院患者入院入科
  Procedure Zlhis_Patient_002
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );
  --56.住院患者床位变更
  Procedure Zlhis_Patient_004
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );
  --57.住院患者病情变更
  Procedure Zlhis_Patient_005
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );
  --58.住院患者变更撤消
  Procedure Zlhis_Patient_006
  (
    病人id_In   In 病案主页.病人id%Type,
    主页id_In   In 病案主页.主页id%Type,
    撤销方式_In In Varchar2
  );
  --59.住院患者医护变更
  Procedure Zlhis_Patient_007
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );
  --住院患者护理等级变更
  Procedure Zlhis_Patient_008
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );
  --60.住院患者预出院
  Procedure Zlhis_Patient_009
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );
  --61.住院患者出院
  Procedure Zlhis_Patient_010
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );
  --62.住院患者新生儿登记
  Procedure Zlhis_Patient_011
  (
    病人id_In   In 病案主页.病人id%Type,
    主页id_In   In 病案主页.主页id%Type,
    婴儿序号_In 病人医嘱记录.婴儿%Type
  );
  --63.住院患者转入科室
  Procedure Zlhis_Patient_012
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );
  --64.新生儿登记作废
  Procedure Zlhis_Patient_013
  (
    病人id_In   In 病案主页.病人id%Type,
    主页id_In   In 病案主页.主页id%Type,
    婴儿序号_In 病人医嘱记录.婴儿%Type
  );
  --65.门诊患者登记
  Procedure Zlhis_Patient_015(病人id_In In 病案主页.病人id%Type);
  --66.患者信息修改
  Procedure Zlhis_Patient_016(病人id_In In 病案主页.病人id%Type);

  --67.患者合并
  Procedure Zlhis_Patient_017
  (
    病人id_In   In 病案主页.病人id%Type,
    原病人id_In In 病案主页.病人id%Type
  );

  --69.患者转病区转入
  Procedure Zlhis_Patient_026
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );

  Procedure Zlhis_Patient_028(病人id_In In 病案主页.病人id%Type);

  --血库:科室配血完成
  Procedure Zlhis_Blood_001(医嘱id_In In 病人医嘱记录.Id%Type);
  --血库:科室配血拒绝
  Procedure Zlhis_Blood_002(医嘱id_In In 病人医嘱记录.Id%Type);

  --70.检验标本审核
  Procedure Zlhis_Lis_001(标本id_In In 检验标本记录.Id%Type);
  --71.检验标本审核撤消
  Procedure Zlhis_Lis_002(标本id_In In 检验标本记录.Id%Type);
  --73.检验标本条码打印
  Procedure Zlhis_Lis_004
  (
    样本条码_In In 病人医嘱发送.样本条码%Type,
    医嘱id_In   In 病人医嘱发送.医嘱id%Type,
    医嘱ids_In  In Varchar2
  );
  --74.检验标本条码打印撤销
  Procedure Zlhis_Lis_005
  (
    样本条码_In In 病人医嘱发送.样本条码%Type,
    医嘱id_In   In 病人医嘱发送.医嘱id%Type,
    医嘱ids_In  In Varchar2
  );
  --75.检验标本核收
  Procedure Zlhis_Lis_006(标本id_In In 检验标本记录.Id%Type);
  --76.检验标本核收撤销
  Procedure Zlhis_Lis_007(标本id_In In 检验标本记录.Id%Type);
  --77.检验标本拒收
  Procedure Zlhis_Lis_008(标本id_In In 检验标本记录.Id%Type);
End b_Message;
/
Create Or Replace Package Body b_Message Is
  --是否是平台调用
  Is_Platform_Call Number(1) := 0;
  --消息公共方法
  Message_Creator Zlmsg_Todo.Creator%Type := Null;
  --缓存消息查询结果
  Type Tmap_Msg_Using Is Table Of Number(1) Index By Varchar2(30);
  Zlmsg_Map Tmap_Msg_Using;
  --消息是否启用
  Function p_Msg_Using(Msg_Code_In Zlmsg_Lists.Code%Type) Return Number As
    n_Using Zlmsg_Lists.Using%Type;
    v_Code  Zlmsg_Lists.Code%Type;
  Begin
    If Is_Platform_Call = 1 Then
      Return 0;
    End If;
    v_Code := Upper(Msg_Code_In);
    Begin
      n_Using := Zlmsg_Map(v_Code);
      Return n_Using;
    Exception
      When No_Data_Found Then
        --不采取Max容错处理，错误相当于外键,用户可能没有采取同步修改或自己增加了消息类型但是未注册到Zlmsg_Lists，这两种情况会出现错误。


      
        Select Nvl(Using, 0) Into n_Using From Zlmsg_Lists A Where Code = v_Code;
        Zlmsg_Map(v_Code) := n_Using;
        --查询生成消息的人员，放在这里减少执行次数
        If Message_Creator Is Null Then
          Message_Creator := Zl_Username;
        End If;
        Return n_Using;
    End;
  Exception
    When Others Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || '未在Zlmsg_Lists中找到消息"' || v_Code || '"！请联系管理员进行处理。' || '[ZLSOFT]');
      Return 0;
  End;
  Procedure p_Msg_Todo_Insert
  (
    Msg_Code_In  Zlmsg_Todo.Msg_Code%Type,
    Key_Value_In Zlmsg_Todo.Key_Value%Type
  ) Is
  Begin
    If p_Msg_Using(Msg_Code_In) = 0 Then
      Return;
    End If;
    Insert Into Zlmsg_Todo
      (ID, Msg_Code, Key_Value, State, Create_Time, Creator)
    Values
      (Zlmsg_Todo_Id.Nextval, Upper(Msg_Code_In), Key_Value_In, 0, Sysdate, Message_Creator);
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Msg_Todo_Insert;
  --设置当前会话为平台调用
  Procedure Set_Platform_Call(Platform_Call Number) Is
  Begin
    Is_Platform_Call := Platform_Call;
  End Set_Platform_Call;
  --消息Zlhis_Dict_001
  Procedure Zlhis_Dict_001(Id_In 部门表.Id%Type) Is
    v_Define Xmltype;
    v_Value  Zlmsg_Todo.Key_Value%Type;
  Begin
    If p_Msg_Using('ZLHIS_DICT_001') = 0 Then
      Return;
    End If;
    Begin
      Select Xmltype(Key_Define) Into v_Define From Zlmsg_Lists Where Code = 'ZLHIS_DICT_001';
    Exception
      When Others Then
        v_Define := Xmltype('<root><ID>NULL</ID></root>');
    End;
    Select Updatexml(v_Define, '/root/ID/text()', Id_In).Getstringval() Into v_Value From Dual;
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_001', v_Value);
  End Zlhis_Dict_001;
  --修改部门
  Procedure Zlhis_Dict_002(部门id_In 部门表.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><部门ID>' || 部门id_In || '</部门ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_002', v_Value);
  End Zlhis_Dict_002;
  --停用部门
  Procedure Zlhis_Dict_003(部门id_In 部门表.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><部门ID>' || 部门id_In || '</部门ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_003', v_Value);
  End Zlhis_Dict_003;
  --启用部门
  Procedure Zlhis_Dict_004(部门id_In 部门表.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><部门ID>' || 部门id_In || '</部门ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_004', v_Value);
  End Zlhis_Dict_004;
  --新增人员
  Procedure Zlhis_Dict_005(人员id_In 人员表.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><人员ID>' || 人员id_In || '</人员ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_005', v_Value);
  End Zlhis_Dict_005;
  --修改人员
  Procedure Zlhis_Dict_006(人员id_In 人员表.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><人员ID>' || 人员id_In || '</人员ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_006', v_Value);
  End Zlhis_Dict_006;
  --停用人员
  Procedure Zlhis_Dict_007(人员id_In 人员表.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><人员ID>' || 人员id_In || '</人员ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_007', v_Value);
  End Zlhis_Dict_007;
  --启用人员
  Procedure Zlhis_Dict_008(人员id_In 人员表.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><人员ID>' || 人员id_In || '</人员ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_008', v_Value);
  End Zlhis_Dict_008;
  --新增收费项目
  Procedure Zlhis_Dict_009(细目id_In 收费项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><细目ID>' || 细目id_In || '</细目ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_009', v_Value);
  End Zlhis_Dict_009;
  --修改收费项目
  Procedure Zlhis_Dict_010(细目id_In 收费项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><细目ID>' || 细目id_In || '</细目ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_010', v_Value);
  End Zlhis_Dict_010;
  --停用收费项目
  Procedure Zlhis_Dict_011(细目id_In 收费项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><细目ID>' || 细目id_In || '</细目ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_011', v_Value);
  End Zlhis_Dict_011;
  --启用收费项目
  Procedure Zlhis_Dict_012(细目id_In 收费项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><细目ID>' || 细目id_In || '</细目ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_012', v_Value);
  End Zlhis_Dict_012;
  --新增诊疗项目
  Procedure Zlhis_Dict_013(诊疗id_In 诊疗项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗ID>' || 诊疗id_In || '</诊疗ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_013', v_Value);
  End Zlhis_Dict_013;
  --修改诊疗项目
  Procedure Zlhis_Dict_014(诊疗id_In 诊疗项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗ID>' || 诊疗id_In || '</诊疗ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_014', v_Value);
  End Zlhis_Dict_014;
  --停用诊疗项目
  Procedure Zlhis_Dict_015(诊疗id_In 诊疗项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗ID>' || 诊疗id_In || '</诊疗ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_015', v_Value);
  End Zlhis_Dict_015;
  --启用诊疗项目
  Procedure Zlhis_Dict_016(诊疗id_In 诊疗项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗ID>' || 诊疗id_In || '</诊疗ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_016', v_Value);
  End Zlhis_Dict_016;
  --新增检验项目
  Procedure Zlhis_Dict_017(诊疗id_In 诊疗项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗ID>' || 诊疗id_In || '</诊疗ID><系统>1</系统></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_017', v_Value);
  End Zlhis_Dict_017;
  --修改检验项目
  Procedure Zlhis_Dict_018(诊疗id_In 诊疗项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗ID>' || 诊疗id_In || '</诊疗ID><系统>1</系统></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_018', v_Value);
  End Zlhis_Dict_018;
  --删除检验项目
  Procedure Zlhis_Dict_019
  (
    诊疗id_In 诊疗项目目录.Id%Type,
    编码_In   诊治所见项目.编码%Type,
    中文名_In 诊治所见项目.中文名%Type,
    英文名_In 诊治所见项目.英文名%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗ID>' || 诊疗id_In || '</诊疗ID>' || '<编码>' || 编码_In || '</编码>' || '<中文名>' || 中文名_In || '</中文名>' ||
               '<英文名>' || 英文名_In || '</英文名>' || '<系统>1</系统></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_019', v_Value);
  End Zlhis_Dict_019;
  --新增疾病编码目录
  Procedure Zlhis_Dict_021(疾病id_In 疾病编码目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><疾病ID>' || 疾病id_In || '</疾病ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_021', v_Value);
  End Zlhis_Dict_021;
  --修改疾病编码目录
  Procedure Zlhis_Dict_022(疾病id_In 疾病编码目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><疾病ID>' || 疾病id_In || '</疾病ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_022', v_Value);
  End Zlhis_Dict_022;
  --停用疾病编码目录
  Procedure Zlhis_Dict_023(疾病id_In 疾病编码目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><疾病ID>' || 疾病id_In || '</疾病ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_023', v_Value);
  End Zlhis_Dict_023;
  --启用疾病编码目录
  Procedure Zlhis_Dict_024(疾病id_In 疾病编码目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><疾病ID>' || 疾病id_In || '</疾病ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_024', v_Value);
  End Zlhis_Dict_024;
  --新增药品分类
  Procedure Zlhis_Dict_025
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_025', v_Value);
  End Zlhis_Dict_025;
  --修改药品分类
  Procedure Zlhis_Dict_026
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_026', v_Value);
  End Zlhis_Dict_026;
  --删除药品分类
  Procedure Zlhis_Dict_027
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type,
    编码_In 诊疗分类目录.编码%Type,
    名称_In 诊疗分类目录.名称%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || Id_In || '</ID><编码>' || 编码_In || '</编码><名称>' || 名称_In ||
               '</名称></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_027', v_Value);
  End Zlhis_Dict_027;
  --停用药品分类
  Procedure Zlhis_Dict_028
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_028', v_Value);
  End Zlhis_Dict_028;
  --启用药品分类
  Procedure Zlhis_Dict_029
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_029', v_Value);
  End Zlhis_Dict_029;
  --新增药品品种
  Procedure Zlhis_Dict_030
  (
    类别_In 诊疗项目目录.类别%Type,
    Id_In   诊疗项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_030', v_Value);
  End Zlhis_Dict_030;
  --修改药品品种
  Procedure Zlhis_Dict_031
  (
    类别_In 诊疗项目目录.类别%Type,
    Id_In   诊疗项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_031', v_Value);
  End Zlhis_Dict_031;
  --删除药品品种
  Procedure Zlhis_Dict_032
  (
    类别_In 诊疗项目目录.类别%Type,
    Id_In   诊疗项目目录.Id%Type,
    编码_In 诊疗项目目录.编码%Type,
    名称_In 诊疗项目目录.名称%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><ID>' || Id_In || '</ID><编码>' || 编码_In || '</编码><名称>' || 名称_In ||
               '</名称></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_032', v_Value);
  End Zlhis_Dict_032;
  --停用药品品种
  Procedure Zlhis_Dict_033
  (
    类别_In 诊疗项目目录.类别%Type,
    Id_In   诊疗项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_033', v_Value);
  End Zlhis_Dict_033;
  --启用药品品种
  Procedure Zlhis_Dict_034
  (
    类别_In 诊疗项目目录.类别%Type,
    Id_In   诊疗项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_034', v_Value);
  End Zlhis_Dict_034;
  --新增药品规格
  Procedure Zlhis_Dict_035
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><药品ID>' || 药品id_In || '</药品ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_035', v_Value);
  End Zlhis_Dict_035;
  --修改药品规格
  Procedure Zlhis_Dict_036
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><药品ID>' || 药品id_In || '</药品ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_036', v_Value);
  End Zlhis_Dict_036;
  --删除药品规格
  Procedure Zlhis_Dict_037
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type,
    编码_In   收费项目目录.编码%Type,
    名称_In   收费项目目录.名称%Type,
    规格_In   收费项目目录.规格%Type,
    产地_In   收费项目目录.产地%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><药品ID>' || 药品id_In || '</药品ID><编码>' || 编码_In || '</编码><名称>' || 名称_In ||
               '</名称><规格>' || 规格_In || '</规格><产地>' || 产地_In || '</产地></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_037', v_Value);
  End Zlhis_Dict_037;
  --停用药品规格
  Procedure Zlhis_Dict_038
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><药品ID>' || 药品id_In || '</药品ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_038', v_Value);
  End Zlhis_Dict_038;
  --启用药品规格
  Procedure Zlhis_Dict_039
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><药品ID>' || 药品id_In || '</药品ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_039', v_Value);
  End Zlhis_Dict_039;
  --设置药品存储库房
  Procedure Zlhis_Dict_040
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><药品ID>' || 药品id_In || '</药品ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_040', v_Value);
  End Zlhis_Dict_040;
  --设置药品储备限额
  Procedure Zlhis_Dict_041
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><药品ID>' || 药品id_In || '</药品ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_041', v_Value);
  End Zlhis_Dict_041;
  --新增卫材品种
  Procedure Zlhis_Dict_042(Id_In 诊疗项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><材料ID>' || Id_In || '</材料ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_042', v_Value);
  End Zlhis_Dict_042;
  --新增卫材规格
  Procedure Zlhis_Dict_043(Id_In 收费项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><材料ID>' || Id_In || '</材料ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_043', v_Value);
  End Zlhis_Dict_043;
  --修改卫材规格
  Procedure Zlhis_Dict_044(Id_In 收费项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><材料ID>' || Id_In || '</材料ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_044', v_Value);
  End Zlhis_Dict_044;
  --删除卫材规格
  Procedure Zlhis_Dict_045
  (
    Id_In   收费项目目录.Id%Type,
    编码_In 收费项目目录.编码%Type,
    名称_In 收费项目目录.名称%Type,
    规格_In 收费项目目录.规格%Type,
    产地_In 收费项目目录.产地%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><材料ID>' || Id_In || '</材料ID><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><规格>' || 规格_In ||
               '</规格><产地>' || 产地_In || '</产地></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_045', v_Value);
  End Zlhis_Dict_045;
  --停用卫材规格
  Procedure Zlhis_Dict_046(Id_In 收费项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><材料ID>' || Id_In || '</材料ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_046', v_Value);
  End Zlhis_Dict_046;
  --启用卫材规格
  Procedure Zlhis_Dict_047(Id_In 收费项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><材料ID>' || Id_In || '</材料ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_047', v_Value);
  End Zlhis_Dict_047;
  --医保对码
  Procedure Zlhis_Dict_048
  (
    险类_In       In 保险支付项目.险类%Type,
    收费细目id_In In 保险支付项目.收费细目id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><险类>' || 险类_In || '</险类><收费细目ID>' || 收费细目id_In || '</收费细目ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_048', v_Value);
  End Zlhis_Dict_048;
  --删除医保对码
  Procedure Zlhis_Dict_049
  (
    险类_In       In 保险支付项目.险类%Type,
    收费细目id_In In 保险支付项目.收费细目id%Type,
    项目编码_In   In 收费项目目录.编码%Type,
    项目名称_In   In 收费项目目录.名称%Type,
    医保编码_In   In 保险支付项目.项目编码%Type,
    医保名称_In   In 保险支付项目.项目名称%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><险类>' || 险类_In || '</险类><收费细目ID>' || 收费细目id_In || '</收费细目ID><项目编码>' || 项目编码_In || '</项目编码><项目名称>' ||
               项目名称_In || '</项目名称><医保编码>' || 医保编码_In || '</医保编码><医保名称>' || 医保名称_In || '</医保名称></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_049', v_Value);
  End Zlhis_Dict_049;
  --新增卫材分类
  Procedure Zlhis_Dict_050
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_050', v_Value);
  End Zlhis_Dict_050;
  --修改卫材分类
  Procedure Zlhis_Dict_051
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_051', v_Value);
  End Zlhis_Dict_051;
  --删除卫材分类
  Procedure Zlhis_Dict_052
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type,
    编码_In 诊疗分类目录.编码%Type,
    名称_In 诊疗分类目录.名称%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || Id_In || '</ID><编码>' || 编码_In || '</编码><名称>' || 名称_In ||
               '</名称></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_052', v_Value);
  End Zlhis_Dict_052;
  --收费价目变动
  Procedure Zlhis_Dict_053
  (
    收费项目Id_In       收费项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><收费项目ID>' || 收费项目Id_In || '</收费项目ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_053', v_Value);
  End Zlhis_Dict_053;

  --诊疗收费对照变动
  Procedure Zlhis_Dict_054
  (
    诊疗项目Id_In     诊疗分类目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗项目ID>' || 诊疗项目Id_In || '</诊疗项目ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_054', v_Value);
  End Zlhis_Dict_054;
  --新增诊疗检查类型
  Procedure Zlhis_Dictpacs_001
  (
    编码_In   诊疗检查类型.编码%Type,
    名称_In   诊疗检查类型.名称%Type,
    简码_In   诊疗检查类型.简码%Type,
    建病案_In 诊疗检查类型.建病案%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><建病案>' || 建病案_In ||
               '</建病案></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_001', v_Value);
  End Zlhis_Dictpacs_001;

  --修改诊疗检查类型
  Procedure Zlhis_Dictpacs_002
  (
    编码_In   诊疗检查类型.编码%Type,
    名称_In   诊疗检查类型.名称%Type,
    简码_In   诊疗检查类型.简码%Type,
    建病案_In 诊疗检查类型.建病案%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><建病案>' || 建病案_In ||
               '</建病案></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_002', v_Value);
  End Zlhis_Dictpacs_002;
  --删除诊疗检查类型
  Procedure Zlhis_Dictpacs_003
  (
    编码_In   诊疗检查类型.编码%Type,
    名称_In   诊疗检查类型.名称%Type,
    简码_In   诊疗检查类型.简码%Type,
    建病案_In 诊疗检查类型.建病案%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><建病案>' || 建病案_In ||
               '</建病案></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_003', v_Value);
  End Zlhis_Dictpacs_003;
  --新增诊疗检查部位
  Procedure Zlhis_Dictpacs_004
  (
    类型_In     诊疗检查部位.类型%Type,
    编码_In     诊疗检查部位.编码%Type,
    名称_In     诊疗检查部位.名称%Type,
    分组_In     诊疗检查部位.分组%Type,
    备注_In     诊疗检查部位.备注%Type,
    方法_In     诊疗检查部位.方法%Type,
    适用性别_In 诊疗检查部位.适用性别%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><分组>' || 分组_In ||
               '</分组><备注>' || 备注_In || '</备注><方法>' || 方法_In || '</方法><适用性别>' || 适用性别_In || '</适用性别></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_004', v_Value);
  End Zlhis_Dictpacs_004;
  --修改诊疗检查部位
  Procedure Zlhis_Dictpacs_005
  (
    类型_In     诊疗检查部位.类型%Type,
    编码_In     诊疗检查部位.编码%Type,
    名称_In     诊疗检查部位.名称%Type,
    分组_In     诊疗检查部位.分组%Type,
    备注_In     诊疗检查部位.备注%Type,
    方法_In     诊疗检查部位.方法%Type,
    适用性别_In 诊疗检查部位.适用性别%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><分组>' || 分组_In ||
               '</分组><备注>' || 备注_In || '</备注><方法>' || 方法_In || '</方法><适用性别>' || 适用性别_In || '</适用性别></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_005', v_Value);
  End Zlhis_Dictpacs_005;
  --删除诊疗检查部位
  Procedure Zlhis_Dictpacs_006
  (
    类型_In     诊疗检查部位.类型%Type,
    编码_In     诊疗检查部位.编码%Type,
    名称_In     诊疗检查部位.名称%Type,
    分组_In     诊疗检查部位.分组%Type,
    备注_In     诊疗检查部位.备注%Type,
    方法_In     诊疗检查部位.方法%Type,
    适用性别_In 诊疗检查部位.适用性别%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><分组>' || 分组_In ||
               '</分组><备注>' || 备注_In || '</备注><方法>' || 方法_In || '</方法><适用性别>' || 适用性别_In || '</适用性别></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_006', v_Value);
  End Zlhis_Dictpacs_006;
  --新增诊疗项目部位
  Procedure Zlhis_Dictpacs_007
  (
    Id_In     诊疗项目部位.Id%Type,
    项目id_In 诊疗项目部位.项目id%Type,
    类型_In   诊疗项目部位.类型%Type,
    部位_In   诊疗项目部位.部位%Type,
    方法_In   诊疗项目部位.方法%Type,
    默认_In   诊疗项目部位.默认%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ID>' || Id_In || '</ID><项目ID>' || 项目id_In || '</项目ID><类型>' || 类型_In || '</类型><部位>' || 部位_In ||
               '</部位><方法>' || 方法_In || '</方法><默认>' || 默认_In || '</默认></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_007', v_Value);
  End Zlhis_Dictpacs_007;
  --修改诊疗项目部位
  Procedure Zlhis_Dictpacs_008
  (
    Id_In     诊疗项目部位.Id%Type,
    项目id_In 诊疗项目部位.项目id%Type,
    类型_In   诊疗项目部位.类型%Type,
    部位_In   诊疗项目部位.部位%Type,
    方法_In   诊疗项目部位.方法%Type,
    默认_In   诊疗项目部位.默认%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ID>' || Id_In || '</ID><项目ID>' || 项目id_In || '</项目ID><类型>' || 类型_In || '</类型><部位>' || 部位_In ||
               '</部位><方法>' || 方法_In || '</方法><默认>' || 默认_In || '</默认></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_008', v_Value);
  End Zlhis_Dictpacs_008;
  --删除诊疗项目部位
  Procedure Zlhis_Dictpacs_009
  (
    Id_In     诊疗项目部位.Id%Type,
    项目id_In 诊疗项目部位.项目id%Type,
    类型_In   诊疗项目部位.类型%Type,
    部位_In   诊疗项目部位.部位%Type,
    方法_In   诊疗项目部位.方法%Type,
    默认_In   诊疗项目部位.默认%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ID>' || Id_In || '</ID><项目ID>' || 项目id_In || '</项目ID><类型>' || 类型_In || '</类型><部位>' || 部位_In ||
               '</部位><方法>' || 方法_In || '</方法><默认>' || 默认_In || '</默认></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_009', v_Value);
  End Zlhis_Dictpacs_009;
  --新增诊疗项目部位
  Procedure Zlhis_Dictlis_004
  (
    编码_In     诊疗检验标本.编码%Type,
    名称_In     诊疗检验标本.名称%Type,
    简码_In     诊疗检验标本.简码%Type,
    适用性别_In 诊疗检验标本.适用性别%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><适用性别>' || 适用性别_In ||
               '</适用性别></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_004', v_Value);
  End Zlhis_Dictlis_004;
  --修改诊疗项目部位
  Procedure Zlhis_Dictlis_005
  (
    编码_In     诊疗检验标本.编码%Type,
    名称_In     诊疗检验标本.名称%Type,
    简码_In     诊疗检验标本.简码%Type,
    适用性别_In 诊疗检验标本.适用性别%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><适用性别>' || 适用性别_In ||
               '</适用性别></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_005', v_Value);
  End Zlhis_Dictlis_005;
  --删除诊疗项目部位
  Procedure Zlhis_Dictlis_006
  (
    编码_In     诊疗检验标本.编码%Type,
    名称_In     诊疗检验标本.名称%Type,
    简码_In     诊疗检验标本.简码%Type,
    适用性别_In 诊疗检验标本.适用性别%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><适用性别>' || 适用性别_In ||
               '</适用性别></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_006', v_Value);
  End Zlhis_Dictlis_006;
  --新增采血管类型
  Procedure Zlhis_Dictlis_007
  (
    编码_In   采血管类型.编码%Type,
    名称_In   采血管类型.名称%Type,
    简码_In   采血管类型.简码%Type,
    添加剂_In 采血管类型.添加剂%Type,
    采血量_In 采血管类型.采血量%Type,
    规格_In   采血管类型.规格%Type,
    颜色_In   采血管类型.颜色%Type,
    材料id_In 采血管类型.材料id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><添加剂>' || 添加剂_In ||
               '</添加剂><采血量>' || 采血量_In || '</采血量><颜色>' || 颜色_In || '</颜色><规格>' || 规格_In || '</规格><材料ID_In>' || 材料id_In ||
               '</材料ID_In></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_007', v_Value);
  End Zlhis_Dictlis_007;
  --新增采血管类型
  Procedure Zlhis_Dictlis_008
  (
    编码_In   采血管类型.编码%Type,
    名称_In   采血管类型.名称%Type,
    简码_In   采血管类型.简码%Type,
    添加剂_In 采血管类型.添加剂%Type,
    采血量_In 采血管类型.采血量%Type,
    规格_In   采血管类型.规格%Type,
    颜色_In   采血管类型.颜色%Type,
    材料id_In 采血管类型.材料id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><添加剂>' || 添加剂_In ||
               '</添加剂><采血量>' || 采血量_In || '</采血量><颜色>' || 颜色_In || '</颜色><规格>' || 规格_In || '</规格><材料ID_In>' || 材料id_In ||
               '</材料ID_In></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_008', v_Value);
  End Zlhis_Dictlis_008;
  --新增采血管类型
  Procedure Zlhis_Dictlis_009
  (
    编码_In   采血管类型.编码%Type,
    名称_In   采血管类型.名称%Type,
    简码_In   采血管类型.简码%Type,
    添加剂_In 采血管类型.添加剂%Type,
    采血量_In 采血管类型.采血量%Type,
    规格_In   采血管类型.规格%Type,
    颜色_In   采血管类型.颜色%Type,
    材料id_In 采血管类型.材料id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><添加剂>' || 添加剂_In ||
               '</添加剂><采血量>' || 采血量_In || '</采血量><颜色>' || 颜色_In || '</颜色><规格>' || 规格_In || '</规格><材料ID_In>' || 材料id_In ||
               '</材料ID_In></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_009', v_Value);
  End Zlhis_Dictlis_009;
  --药品备药发送
  Procedure Zlhis_Drug_001(No_In 药品收发记录.No%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><单据号>' || No_In || '</单据号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_001', v_Value);
  End Zlhis_Drug_001;
  --取消药品备药发送
  Procedure Zlhis_Drug_002(No_In 药品收发记录.No%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><单据号>' || No_In || '</单据号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_002', v_Value);
  End Zlhis_Drug_002;
  --药品移库单接收
  Procedure Zlhis_Drug_003(No_In 药品收发记录.No%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><单据号>' || No_In || '</单据号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_003', v_Value);
  End Zlhis_Drug_003;
  --药品移库单冲销
  Procedure Zlhis_Drug_004(No_In 药品收发记录.No%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><单据号>' || No_In || '</单据号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_004', v_Value);
  End Zlhis_Drug_004;
  --部门发药
  Procedure Zlhis_Drug_005
  (
    库房id_In 药品收发记录.库房id%Type,
    收发id_In 药品收发记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><库房ID>' || 库房id_In || '</库房ID><收发ID>' || 收发id_In || '</收发ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_005', v_Value);
  End Zlhis_Drug_005;
  --部门退药
  Procedure Zlhis_Drug_006
  (
    冲销收发id_In 药品收发记录.Id%Type,
    待发收发id_In 药品收发记录.Id%Type,
    数量_In       药品收发记录.实际数量%Type,
    费用id_In     门诊费用记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><冲销记录ID>' || 冲销收发id_In || '</冲销记录ID><待发记录ID>' || 待发收发id_In || '</待发记录ID><数量>' || 数量_In ||
               '</数量><费用ID>' || 费用id_In || '</费用ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_006', v_Value);
  End Zlhis_Drug_006;
  --药品调价
  Procedure Zlhis_Drug_007(价格id_In 药品价格记录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><价格ID>' || 价格id_In || '</价格ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_007', v_Value);
  End Zlhis_Drug_007;
  --静配发送
  Procedure Zlhis_Drug_008(记录ids_In Varchar2) Is
    v_Value  Zlmsg_Todo.Key_Value%Type;
    n_记录id 输液配药记录.Id%Type;
    v_Tmp    Varchar2(4000);
	n_Length Number(18);
  Begin
    If 记录ids_In Is Null Then
      v_Tmp := Null;
    Else
      v_Tmp := 记录ids_In || ',';
    End If;
  
    v_Value := '<root><记录IDS>';
  
    While v_Tmp Is Not Null Loop
      --分解单据ID串
      n_记录id := To_Number(Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1));
      v_Tmp    := Replace(',' || v_Tmp, ',' || n_记录id || ',');
      
      --判断当前长度是否即将超过缓存                                                                        
      Select Lengthb(v_Value || '<记录ID>' || n_记录id || '</记录ID>') Into n_Length From Dual;            
      If n_Length > 950 Then								                   
        v_Value := v_Value || '</记录IDs></root>';                                                         
        b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_008', v_Value);                                            
        v_Value := '<root><记录IDs>';                                                                      
      End If;

      v_Value := v_Value || '<记录ID>' || n_记录id || '</记录ID>';
    End Loop;
  
    v_Value := v_Value || '</记录IDS></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_008', v_Value);
  End Zlhis_Drug_008;
  --药品调售价
  Procedure Zlhis_Drug_009
  (
    价格id_In 药品价格记录.Id%Type,
    时价_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><价格ID>' || 价格id_In || '</价格ID><时价>' || 时价_In || '</时价></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_009', v_Value);
  End Zlhis_Drug_009;
  --卫材调成本价
  Procedure Zlhis_Drug_010(价格id_In 成本价调价信息.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><价格ID>' || 价格id_In || '</价格ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_010', v_Value);
  End Zlhis_Drug_010;
  --卫材调售价
  Procedure Zlhis_Drug_011
  (
    价格id_In 收费价目.Id%Type,
    时价_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><价格ID>' || 价格id_In || '</价格ID><时价>' || 时价_In || '</时价></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_011', v_Value);
  End Zlhis_Drug_011;

  --2.停止患者医嘱，住院
  Procedure Zlhis_Cis_002
  (
    病人id_In  In 病人医嘱记录.病人id%Type,
    主页id_In  In 病人医嘱记录.主页id%Type,
    医嘱id_In  In 病人医嘱记录.Id%Type,
    医嘱ids_In In Varchar2
  ) Is
  Begin
    If p_Msg_Using('ZLHIS_CIS_002') = 0 Then
      Return;
    End If;
    If 医嘱id_In Is Not Null Then
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_002',
                                  '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><ID>' || 医嘱id_In ||
                                   '</ID></root>');
    Else
      For R In (Select '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><ID>' || ID || '</ID></root>' As Xml_Value
                From 病人医嘱记录
                Where ID In (Select Column_Value From Table(f_Num2list(医嘱ids_In))) And 相关id Is Null) Loop
        b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_002', r.Xml_Value);
      End Loop;
    End If;
  End Zlhis_Cis_002;
  --3.作废患者医嘱，门诊/住院
  Procedure Zlhis_Cis_003
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><ID>' ||
               医嘱id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_003', v_Value);
  End Zlhis_Cis_003;

  --4.患者术后医嘱，住院
  Procedure Zlhis_Cis_004
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_004',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><ID>' || 医嘱id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_004;

  --5.撤消患者术后医嘱，住院
  Procedure Zlhis_Cis_005
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_005',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><ID>' || 医嘱id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_005;

  --6.患者护理常规医嘱，住院
  Procedure Zlhis_Cis_006
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_006',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In ||
                                 '</发送号><ID>' || 医嘱id_In || '</ID></root>');
  End Zlhis_Cis_006;

  --7.撤消患者护理常规医嘱，住院
  Procedure Zlhis_Cis_007
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_007', v_Value);
  End Zlhis_Cis_007;

  --门诊患者接诊
  Procedure Zlhis_Cis_008
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    挂号单_In In 病人挂号记录.No%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_008', '<root><病人ID>' || 病人id_In || '</病人ID><NO>' || 挂号单_In || '</NO></root>');
  End Zlhis_Cis_008;

  --门诊患者取消接诊
  Procedure Zlhis_Cis_009
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    挂号单_In In 病人挂号记录.No%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_009', '<root><病人ID>' || 病人id_In || '</病人ID><NO>' || 挂号单_In || '</NO></root>');
  End Zlhis_Cis_009;

  --10.下达患者诊断，门诊/住院
  Procedure Zlhis_Cis_010
  (
    病人id_In In 病人挂号记录.病人id%Type,
    就诊id_In In 病人挂号记录.Id%Type, --门诊病人 挂号ID，住院病人 主页ID
    诊断id_In In 病人诊断记录.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_010',
                                '<root><病人ID>' || 病人id_In || '</病人ID><就诊ID>' || 就诊id_In || '</就诊ID><ID>' || 诊断id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_010;
  --11.撤消患者诊断
  Procedure Zlhis_Cis_011
  (
    病人id_In   In 病人挂号记录.病人id%Type,
    就诊id_In   In 病人挂号记录.Id%Type, --门诊病人 挂号ID，住院病人 主页ID
    Id_In       In 病人诊断记录.Id%Type,
    疾病id_In   In 病人诊断记录.疾病id%Type,
    诊断id_In   In 病人诊断记录.诊断id%Type,
    诊断描述_In In 病人诊断记录.诊断描述%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><就诊ID>' || 就诊id_In || '</就诊ID><ID>' || Id_In || '</ID><疾病ID>' ||
               疾病id_In || '</疾病ID><诊断ID>' || 诊断id_In || '</诊断ID><诊断描述>' || 诊断描述_In || '</诊断描述></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_011', v_Value);
  End Zlhis_Cis_011;

  --病区执行医嘱校对
  Procedure Zlhis_Cis_012
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_012',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><ID>' || 医嘱id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_012;

  --13.检验危急值阅读通知
  Procedure Zlhis_Cis_014
  (
    病人id_In In 病人挂号记录.病人id%Type,
    就诊id_In In 病人挂号记录.Id%Type, --门诊病人 挂号ID，住院病人 主页ID
    医嘱id_In In 病人医嘱记录.Id%Type,
    消息id_In In 业务消息清单.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><就诊ID>' || 就诊id_In || '</就诊ID><医嘱ID>' || 医嘱id_In || '</医嘱ID><ID>' ||
               消息id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_014', v_Value);
  End Zlhis_Cis_014;
  --15.患者检验申请，门诊/住院
  Procedure Zlhis_Cis_016
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    病人来源_In In 病人医嘱记录.病人来源%Type --1-门诊;2-住院;3-外来(今后用于辅诊部门接收外来病人);4-体检病人
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><病人来源>' || 病人来源_In || '</病人来源></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_016', v_Value);
  End Zlhis_Cis_016;
  --16.患者检查申请，门诊/住院
  Procedure Zlhis_Cis_017
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    病人来源_In In 病人医嘱记录.病人来源%Type --1-门诊;2-住院;3-外来(今后用于辅诊部门接收外来病人);4-体检病人
  ) Is
    v_Value    Zlmsg_Todo.Key_Value%Type;
    v_操作类型 诊疗项目目录.操作类型%Type;
  Begin
    Select Max(a.操作类型)
    Into v_操作类型
    From 诊疗项目目录 A, 病人医嘱记录 B
    Where b.诊疗项目id = a.Id And b.Id = 医嘱id_In;
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><病人来源>' || 病人来源_In || '</病人来源></root>';
    If v_操作类型 = '病理' Then
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_054', v_Value);
    Else
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_017', v_Value);
    End If;
  End Zlhis_Cis_017;
  --17.患者手术申请，门诊/住院
  Procedure Zlhis_Cis_018
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_018', v_Value);
  End Zlhis_Cis_018;
  --18.患者输血申请，住院
  Procedure Zlhis_Cis_019
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_019', v_Value);
  End Zlhis_Cis_019;
  --19.患者会诊申请，住院
  Procedure Zlhis_Cis_020
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_020', v_Value);
  End Zlhis_Cis_020;
  --20.患者抢救医嘱，住院
  Procedure Zlhis_Cis_021
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_021', v_Value);
  End Zlhis_Cis_021;
  --21.患者死亡医嘱，住院
  Procedure Zlhis_Cis_022
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_022', v_Value);
  End Zlhis_Cis_022;
  --22.患者特殊治疗医嘱，住院
  Procedure Zlhis_Cis_023
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_023', v_Value);
  End Zlhis_Cis_023;

  --24.检查危急值阅读通知
  Procedure Zlhis_Cis_025
  (
    病人id_In In 病人挂号记录.病人id%Type,
    就诊id_In In 病人挂号记录.Id%Type, --门诊病人 挂号ID，住院病人 主页ID
    医嘱id_In In 病人医嘱记录.Id%Type,
    消息id_In In 业务消息清单.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><就诊ID>' || 就诊id_In || '</就诊ID><医嘱ID>' || 医嘱id_In || '</医嘱ID><ID>' ||
               消息id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_025', v_Value);
  End Zlhis_Cis_025;

  --病区执行医嘱发送
  Procedure Zlhis_Cis_026
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_026',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In ||
                                 '</发送号><ID>' || 医嘱id_In || '</ID></root>');
  End Zlhis_Cis_026;

  --撤消患者检验申请
  Procedure Zlhis_Cis_036
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    No_In       In 病人医嘱发送.No%Type,
    病人来源_In In 病人医嘱记录.病人来源%Type --1-门诊;2-住院;3-外来(今后用于辅诊部门接收外来病人);4-体检病人
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><NO>' || No_In || '</NO><病人来源>' || 病人来源_In ||
               '</病人来源></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_036', v_Value);
  End Zlhis_Cis_036;

  --撤消患者检查申请
  Procedure Zlhis_Cis_037
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    No_In       In 病人医嘱发送.No%Type,
    病人来源_In In 病人医嘱记录.病人来源%Type --1-门诊;2-住院;3-外来(今后用于辅诊部门接收外来病人);4-体检病人
  ) Is
    v_Value    Zlmsg_Todo.Key_Value%Type;
    v_操作类型 诊疗项目目录.操作类型%Type;
  Begin
    Select Max(a.操作类型)
    Into v_操作类型
    From 诊疗项目目录 A, 病人医嘱记录 B
    Where b.诊疗项目id = a.Id And b.Id = 医嘱id_In;
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><NO>' || No_In || '</NO><病人来源>' || 病人来源_In ||
               '</病人来源></root>';
    If v_操作类型 = '病理' Then
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_055', v_Value);
    Else
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_037', v_Value);
    End If;
  End Zlhis_Cis_037;

  --撤消患者手术申请
  Procedure Zlhis_Cis_038
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_038', v_Value);
  End Zlhis_Cis_038;

  --撤消患者输血申请
  Procedure Zlhis_Cis_039
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_039', v_Value);
  End Zlhis_Cis_039;

  --撤消患者会诊申请
  Procedure Zlhis_Cis_040
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_040', v_Value);
  End Zlhis_Cis_040;

  --撤消患者抢救医嘱
  Procedure Zlhis_Cis_041
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_041', v_Value);
  End Zlhis_Cis_041;

  --撤消患者死亡医嘱
  Procedure Zlhis_Cis_042
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_042', v_Value);
  End Zlhis_Cis_042;

  --撤消特殊治疗医嘱
  Procedure Zlhis_Cis_043
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_043', v_Value);
  End Zlhis_Cis_043;

  --撤消病区执行医嘱
  Procedure Zlhis_Cis_044
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    No_In       In 病人医嘱发送.No%Type,
    发送数次_In In 病人医嘱发送.发送数次%Type,
    首次时间_In In 病人医嘱发送.首次时间%Type,
    末次时间_In In 病人医嘱发送.末次时间%Type,
    样本条码_In In 病人医嘱发送.样本条码%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID><NO>' || No_In || '</NO><发送数次>' || 发送数次_In || '</发送数次><首次时间>' ||
               To_Char(首次时间_In, 'yyyy-mm-dd hh24:mi:ss') || '</首次时间><末次时间>' ||
               To_Char(末次时间_In, 'yyyy-mm-dd hh24:mi:ss') || '</末次时间><样本条码>' || 样本条码_In || '</样本条码></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_044', v_Value);
  End Zlhis_Cis_044;

  --患者医嘱执行登记
  Procedure Zlhis_Cis_050
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    要求时间_In In 病人医嘱执行.要求时间%Type,
    执行时间_In In 病人医嘱执行.执行时间%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><要求时间>' || To_Char(要求时间_In, 'yyyy-mm-dd hh24:mi:ss') ||
               '</要求时间><执行时间>' || To_Char(执行时间_In, 'yyyy-mm-dd hh24:mi:ss') || '</执行时间></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_050', v_Value);
  End Zlhis_Cis_050;

  --患者医嘱取消执行登记
  Procedure Zlhis_Cis_051
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    要求时间_In In 病人医嘱执行.要求时间%Type,
    执行时间_In In 病人医嘱执行.执行时间%Type,
    本次数次_In In 病人医嘱执行.本次数次%Type,
    执行结果_In In 病人医嘱执行.执行结果%Type,
    执行摘要_In In 病人医嘱执行.执行摘要%Type,
    执行科室_In In 病人医嘱执行.执行科室id%Type,
    执行人_In   In 病人医嘱执行.执行人%Type,
    核对人_In   In 病人医嘱执行.核对人%Type,
    记录来源_In In 病人医嘱执行.记录来源%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><要求时间>' || To_Char(要求时间_In, 'yyyy-mm-dd hh24:mi:ss') ||
               '</要求时间><执行时间>' || To_Char(执行时间_In, 'yyyy-mm-dd hh24:mi:ss') || '</执行时间><本次数次>' || 本次数次_In ||
               '</本次数次><执行结果>' || 执行结果_In || '</执行结果><执行摘要>' || 执行摘要_In || '</执行摘要><执行科室ID>' || 执行科室_In ||
               '</执行科室ID><执行人>' || 执行人_In || '</执行人><核对人>' || 核对人_In || '</核对人><记录来源>' || 记录来源_In || '</记录来源></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_051', v_Value);
  End Zlhis_Cis_051;
  --患者医嘱执行完成
  Procedure Zlhis_Cis_052
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_052',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In ||
                                 '</挂号单><发送号>' || 发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID></root>');
  End Zlhis_Cis_052;
  --患者医嘱撤消执行完成
  Procedure Zlhis_Cis_053
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_053',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In ||
                                 '</挂号单><发送号>' || 发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID></root>');
  End Zlhis_Cis_053;

  --病理申请发送后修改
  Procedure Zlhis_Cis_056
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    病人来源_In In 病人医嘱记录.病人来源%Type --1-门诊;2-住院;3-外来(今后用于辅诊部门接收外来病人);4-体检病人
  ) Is
    v_Value    Zlmsg_Todo.Key_Value%Type;
    v_操作类型 诊疗项目目录.操作类型%Type;
  Begin
    Select Max(a.操作类型)
    Into v_操作类型
    From 诊疗项目目录 A, 病人医嘱记录 B
    Where b.诊疗项目id = a.Id And b.Id = 医嘱id_In;
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><病人来源>' || 病人来源_In || '</病人来源></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_056', v_Value);
  End Zlhis_Cis_056;

  --门诊患者完成就诊
  Procedure Zlhis_Cis_057
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    挂号单_In In 病人挂号记录.No%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_057', '<root><病人ID>' || 病人id_In || '</病人ID><NO>' || 挂号单_In || '</NO></root>');
  End Zlhis_Cis_057;

  --门诊患者取消完成就诊
  Procedure Zlhis_Cis_058
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    挂号单_In In 病人挂号记录.No%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_058', '<root><病人ID>' || 病人id_In || '</病人ID><NO>' || 挂号单_In || '</NO></root>');
  End Zlhis_Cis_058;

  --确认停止患者医嘱 
  Procedure Zlhis_Cis_059
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_059','<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><ID>' || 医嘱id_In || '</ID></root>');
  End Zlhis_Cis_059;

  --26.检查报告完成，检查完成时
  Procedure Zlhis_Pacs_001
  (
    医嘱id_In   In 影像检查记录.医嘱id%Type,
    报告id_Ins  In Varchar2,
    报告类型_In In Number --1-老版PACS报告，2-老版病历编辑器报告，3-新版编辑器报告
  ) Is
  Begin
    If p_Msg_Using('ZLHIS_PACS_001') = 0 Then
      Return;
    End If;
    For R In (Select '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID><报告ID>' || Column_Value || '</报告ID><报告类型>' || 报告类型_In ||
                      '<报告类型></root>' As Xml_Value
              From Table(f_Str2list(报告id_Ins))) Loop
      b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_001', r.Xml_Value);
    End Loop;
  End Zlhis_Pacs_001;
  --27.检查状态同步，检查状态改变后
  Procedure Zlhis_Pacs_002
  (
    医嘱id_In In 影像检查记录.医嘱id%Type,
    原状态_In In 病人医嘱发送.执行过程%Type,
    新状态_In In 病人医嘱发送.执行过程%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID><原状态>' || 原状态_In || '</原状态><新状态>' || 新状态_In || '</新状态></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_002', v_Value);
  End Zlhis_Pacs_002;
  --28.检查状态回退，检查状态回退后
  Procedure Zlhis_Pacs_003
  (
    医嘱id_In In 影像检查记录.医嘱id%Type,
    原状态_In In 病人医嘱发送.执行过程%Type,
    新状态_In In 病人医嘱发送.执行过程%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID><原状态>' || 原状态_In || '</原状态><新状态>' || 新状态_In || '</新状态></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_003', v_Value);
  End Zlhis_Pacs_003;
  --29.检查报告撤销，撤销检查完成时
  Procedure Zlhis_Pacs_004
  (
    医嘱id_In   In 影像检查记录.医嘱id%Type,
    报告id_Ins  In Varchar2,
    报告类型_In In Number --1-老版PACS报告，2-老版病历编辑器报告，3-新版编辑器报告
  ) Is
  Begin
    If p_Msg_Using('ZLHIS_PACS_004') = 0 Then
      Return;
    End If;
    For R In (Select '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID><报告ID>' || Column_Value || '</报告ID><报告类型>' || 报告类型_In ||
                      '<报告类型></root>' As Xml_Value
              From Table(f_Str2list(报告id_Ins))) Loop
      b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_004', r.Xml_Value);
    End Loop;
  End Zlhis_Pacs_004;
  --30.检查危急值通知，检查发生危急值时
  Procedure Zlhis_Pacs_005(医嘱id_In In 影像检查记录.医嘱id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_005', v_Value);
  End Zlhis_Pacs_005;
  -- 检查预约通知，检查预约时
  Procedure Zlhis_Pacs_006
  (
    医嘱id_In In 影像检查记录.医嘱id%Type,
    预约id_In In Ris检查预约.预约id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID><预约ID>' || 预约id_In || '</预约ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_006', v_Value);
  End Zlhis_Pacs_006;
  -- 取消检查预约，取消预约时
  Procedure Zlhis_Pacs_007
  (
    医嘱id_In       In 影像检查记录.医嘱id%Type,
    预约id_In       In Ris检查预约.预约id%Type,
    预约日期_In     In Ris检查预约.预约日期%Type,
    预约序号_In     In Ris检查预约.序号%Type,
    检查设备名称_In In Ris检查预约.检查设备名称%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID><预约ID>' || 预约id_In || '</预约ID><预约日期>' || 预约日期_In || '</预约日期><预约序号>' ||
               预约序号_In || '</预约序号><检查设备名称>' || 检查设备名称_In || '</检查设备名称></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_007', v_Value);
  End Zlhis_Pacs_007;

  --36.患者发卡或绑定卡
  Procedure Zlhis_Patient_018
  (
    变动id_In   In 病人变动记录.Id%Type,
    病人id_In   In 病人信息.病人id%Type,
    卡类别id_In In 医疗卡类别.Id%Type,
    卡号_In     In 病人医疗卡信息.卡号%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><变动ID>' || 变动id_In || '</变动ID><病人ID>' || 病人id_In || '</病人ID><卡类别ID>' || 卡类别id_In ||
               '</卡类别ID><卡号>' || 卡号_In || '</卡号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_018', v_Value);
  End;

  --37.患者退卡
  Procedure Zlhis_Patient_019
  (
    变动id_In   In 病人变动记录.Id%Type,
    病人id_In   In 病人信息.病人id%Type,
    卡类别id_In In 医疗卡类别.Id%Type,
    卡号_In     In 病人医疗卡信息.卡号%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><变动ID>' || 变动id_In || '</变动ID><病人ID>' || 病人id_In || '</病人ID><卡类别ID>' || 卡类别id_In ||
               '</卡类别ID><卡号>' || 卡号_In || '</卡号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_019', v_Value);
  End;

  --38.患者补卡/换卡
  Procedure Zlhis_Patient_020
  (
    变动id_In   In 病人变动记录.Id%Type,
    病人id_In   In 病人信息.病人id%Type,
    卡类别id_In In 医疗卡类别.Id%Type,
    原卡号_In   In 病人医疗卡信息.卡号%Type,
    新卡号_In   In 病人医疗卡信息.卡号%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><变动ID>' || 变动id_In || '</变动ID><病人ID>' || 病人id_In || '</病人ID><卡类别ID>' || 卡类别id_In ||
               '</卡类别ID><原卡号>' || 原卡号_In || '</原卡号><新卡号>' || 新卡号_In || '</新卡号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_020', v_Value);
  End;

  --39.病人挂号登记（包含预约登记)
  Procedure Zlhis_Regist_001
  (
    挂号id_In In 病人挂号记录.Id%Type,
    No_In     In 病人挂号记录.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><挂号ID>' || 挂号id_In || '</挂号ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_001', v_Value);
  End;

  --40.病人分诊
  Procedure Zlhis_Regist_002
  (
    挂号id_In In 病人挂号记录.Id%Type,
    No_In     In 病人挂号记录.No%Type,
    诊室_In   In 病人挂号记录.诊室%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><挂号ID>' || 挂号id_In || '</挂号ID><NO>' || No_In || '</NO><诊室>' || Nvl(诊室_In, '') || '</诊室></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_002', v_Value);
  End;

  --41.病人退号（含取消预约)
  Procedure Zlhis_Regist_003
  (
    挂号id_In In 病人挂号记录.Id%Type,
    No_In     In 病人挂号记录.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><挂号ID>' || 挂号id_In || '</挂号ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_003', v_Value);
  End;

  --42.临床出诊安排调整
  Procedure Zlhis_Regist_004
  (
    变动原因_In In Integer, --1-停诊;2-替诊;3-诊室变动
    记录id_In   In 临床出诊记录.Id%Type,
    变动id_In   In 临床出诊变动记录.Id%Type
    
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><变动原因>' || 变动原因_In || '</变动原因><记录ID>' || 记录id_In || '</记录ID><变动ID>' || 变动id_In ||
               '</变动ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_004', v_Value);
  End;

  --43.门诊患者挂号换号操作
  Procedure Zlhis_Regist_005
  (
    No_In         In 病人挂号记录.No%Type,
    变动原因_In   Integer, --1-替诊;2-换号;3-预约日期变动,
    就诊变动id_In 就诊变动记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><NO>' || No_In || '</NO><变动原因>' || 变动原因_In || '</变动原因><就诊变动ID>' || 就诊变动id_In ||
               '</就诊变动ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_005', v_Value);
  End;

  --费用门诊收费及补充结算
  Procedure Zlhis_Charge_002
  (
    结算类型_In In Number,
    结帐id_In   In 门诊费用记录.结帐id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    --结算类型_In:1-收费结算，2-补充结算
    v_Value := '<root><结算类型>' || 结算类型_In || '</结算类型><结帐ID>' || 结帐id_In || '</结帐ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_002', v_Value);
  End;

  --46.门诊退费单据
  Procedure Zlhis_Charge_004
  (
    退费类型_In In Number,
    结帐id_In   In 门诊费用记录.结帐id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    --退费类型_In:1-收费结算，2-补充结算
    v_Value := '<root><退费类型>' || 退费类型_In || '</退费类型><结帐ID>' || 结帐id_In || '</结帐ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_004', v_Value);
  End;

  --47.收预交款
  Procedure Zlhis_Charge_005
  (
    预交id_In In 病人预交记录.Id%Type,
    单据号_In In 病人预交记录.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><预交ID>' || 预交id_In || '</预交ID><单据号>' || 单据号_In || '</单据号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_005', v_Value);
  End;

  --48.退预交款(包含负数退预交款部分)
  Procedure Zlhis_Charge_006
  (
    退预交id_In In 病人预交记录.Id%Type,
    单据号_In   In 病人预交记录.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><退预交ID>' || 退预交id_In || '</退预交ID><单据号>' || 单据号_In || '</单据号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_006', v_Value);
  End;

  --住院记帐单据
  Procedure Zlhis_Charge_007
  (
    收费类别_In In 住院费用记录.收费类别%Type,
    费用id_In   In 住院费用记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><收费类别>' || 收费类别_In || '</收费类别><费用ID>' || 费用id_In || '</费用ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_007', v_Value);
  End;

  --住院记帐单据销账
  Procedure Zlhis_Charge_008
  (
    收费类别_In In 住院费用记录.收费类别%Type,
    费用id_In   In 住院费用记录.Id%Type,
    收发ids_In  In Varchar2 := Null --可能费用ID对应多个收发id，对应格式：收发id,数量|收发id,数量；非药品不传
  ) Is
    v_Value   Zlmsg_Todo.Key_Value%Type;
    v_Tmp     Varchar2(4000);
    v_Infotmp Varchar2(4000);
    v_Fields  Varchar2(4000);
    v_收发id  Varchar2(50);
    v_数量    Varchar2(20);
  Begin
    If p_Msg_Using('ZLHIS_CHARGE_008') = 0 Then
      Return;
    End If;
    v_Value := '<root><收费类别>' || 收费类别_In || '</收费类别><费用ID>' || 费用id_In || '</费用ID>';
  
    If 收发ids_In Is Null Then
      v_Infotmp := Null;
      v_Tmp     := '<收发IDS>' || '<收发ID>' || '</收发ID>' || '<数量>' || '</数量>' || '</收发IDS>';
    Else
      v_Infotmp := 收发ids_In || '|';
      While v_Infotmp Is Not Null Loop
        --分解收发ID串
        v_Fields  := Substr(v_Infotmp, 1, Instr(v_Infotmp, '|') - 1);
        v_收发id  := Substr(v_Fields, 1, Instr(v_Fields, ',') - 1);
        v_数量    := Substr(v_Fields, Instr(v_Fields, ',') + 1);
        v_Infotmp := Replace('|' || v_Infotmp, '|' || v_Fields || '|');
      
        v_Tmp := v_Tmp || '<收发IDS>' || '<收发ID>' || v_收发id || '</收发ID>' || '<数量>' || v_数量 || '</数量>' || '</收发IDS>';
      End Loop;
    End If;
  
    v_Value := v_Value || v_Tmp || '</root>';
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_008', v_Value);
  End;

  --53.住院患者入院登记
  Procedure Zlhis_Patient_001
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
    n_变动id 病人变动记录.Id%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_001') = 0 Then
      Return;
    End If;
    Select Max(ID)
    Into n_变动id
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 开始原因 = 1 And Nvl(附加床位, 0) = 0;
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_001',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><变动ID>' || n_变动id ||
                                 '</变动ID></root>');
  End Zlhis_Patient_001;
  --54.住院患者入院入科
  Procedure Zlhis_Patient_002
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
    n_变动id 病人变动记录.Id%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_002') = 0 Then
      Return;
    End If;
    Select Max(ID)
    Into n_变动id
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null And Nvl(附加床位, 0) = 0;
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_002',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><变动ID>' || n_变动id ||
                                 '</变动ID></root>');
  End Zlhis_Patient_002;
  --56.住院患者床位变更
  Procedure Zlhis_Patient_004
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
    v_原床号   Varchar2(255);
    v_新床号   Varchar2(255);
    n_变动id   Number(18);
    n_开始原因 Number(3);
    d_开始时间 Date;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_004') = 0 Then
      Return;
    End If;
    Select ID, 床号, 开始时间, 开始原因
    Into n_变动id, v_新床号, d_开始时间, n_开始原因
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null And Nvl(附加床位, 0) = 0;
  
    Select Max(床号)
    Into v_原床号
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 = d_开始时间 And 终止原因 = n_开始原因 And Nvl(附加床位, 0) = 0;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_004',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID>' || '<原床号>' ||
                                 v_原床号 || '</原床号>' || '<新床号>' || v_新床号 || '</新床号>' || '<变动ID>' || n_变动id || '</变动ID>' ||
                                 '</root>');
  End Zlhis_Patient_004;
  --57.住院患者病情变更
  Procedure Zlhis_Patient_005
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
    n_变动id 病人变动记录.Id%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_005') = 0 Then
      Return;
    End If;
    Select Max(ID)
    Into n_变动id
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null And Nvl(附加床位, 0) = 0;
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_005',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><变动ID>' || n_变动id ||
                                 '</变动ID></root>');
  End Zlhis_Patient_005;
  --58.住院患者变更撤消
  Procedure Zlhis_Patient_006
  (
    病人id_In   In 病案主页.病人id%Type,
    主页id_In   In 病案主页.主页id%Type,
    撤销方式_In In Varchar2
  ) Is
    n_科室id     病人变动记录.科室id%Type;
    n_病区id     病人变动记录.病区id%Type;
    n_护理等级id 病人变动记录.护理等级id%Type;
    n_医疗小组id 病人变动记录.医疗小组id%Type;
    v_床号       病人变动记录.床号%Type;
    v_责任护士   病人变动记录.责任护士%Type;
    v_主任医师   病人变动记录.主任医师%Type;
    v_主治医师   病人变动记录.主治医师%Type;
    v_经治医师   病人变动记录.经治医师%Type;
    v_病情       病人变动记录.病情%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_006') = 0 Then
      Return;
    End If;
    Select Max(科室id), Max(病区id), Max(护理等级id), Max(医疗小组id), Max(床号), Max(责任护士), Max(主任医师), Max(主治医师), Max(经治医师), Max(病情)
    Into n_科室id, n_病区id, n_护理等级id, n_医疗小组id, v_床号, v_责任护士, v_主任医师, v_主治医师, v_经治医师, v_病情
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And (终止时间 Is Null Or 终止原因 = 1) And Nvl(附加床位, 0) = 0;
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_006',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><撤销方式>' || 撤销方式_In ||
                                 '</撤销方式><科室ID>' || n_科室id || '</科室ID>' || '<病区ID>' || n_病区id || '</病区ID>' || '<护理等级ID>' ||
                                 n_护理等级id || '</护理等级ID>' || '<医疗小组ID>' || n_医疗小组id || '</医疗小组ID>' || '<床号>' || v_床号 ||
                                 '</床号>' || '<责任护士>' || v_责任护士 || '</责任护士>' || '<主任医师>' || v_主任医师 || '</主任医师>' ||
                                 '<主治医师>' || v_主治医师 || '</主治医师>' || '<经治医师>' || v_经治医师 || '</经治医师>' || '<病情>' || v_病情 ||
                                 '</病情>' || '</root>');
  End Zlhis_Patient_006;
  --59.住院患者医护变更
  Procedure Zlhis_Patient_007
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
    v_原住院医生 Varchar2(100);
    v_新住院医生 Varchar2(100);
    v_原主治医生 Varchar2(100);
    v_新主治医生 Varchar2(100);
    v_原主任医生 Varchar2(100);
    v_新主任医生 Varchar2(100);
    v_原责任护士 Varchar2(100);
    v_新责任护士 Varchar2(100);
    n_变动id     Number(18);
    n_开始原因   Number(3);
    d_开始时间   Date;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_007') = 0 Then
      Return;
    End If;
    Select ID, 经治医师, 主治医师, 主任医师, 责任护士, 开始时间, 开始原因
    Into n_变动id, v_新住院医生, v_新主治医生, v_新主任医生, v_新责任护士, d_开始时间, n_开始原因
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null And Nvl(附加床位, 0) = 0;
  
    Select Max(经治医师), Max(主治医师), Max(主任医师), Max(责任护士)
    Into v_原住院医生, v_原主治医生, v_原主任医生, v_原责任护士
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 = d_开始时间 And 终止原因 = n_开始原因 And Nvl(附加床位, 0) = 0;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_007',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID>' || '<原住院医生>' ||
                                 v_原住院医生 || '</原住院医生>' || '<新住院医生>' || v_新住院医生 || '</新住院医生>' || '<原主治医生>' || v_原主治医生 ||
                                 '</原主治医生>' || '<新主治医生>' || v_新主治医生 || '</新主治医生>' || '<原主任医生>' || v_原主任医生 || '</原主任医生>' ||
                                 '<新主任医生>' || v_新主任医生 || '</新主任医生>' || '<原责任护士>' || v_原责任护士 || '</原责任护士>' || '<新责任护士>' ||
                                 v_新责任护士 || '</新责任护士>' || '<变动ID>' || n_变动id || '</变动ID>' || '</root>');
  End Zlhis_Patient_007;
  --住院患者护理等级变更
  Procedure Zlhis_Patient_008
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
    v_原护理等级id Number(18);
    v_新护理等级id Number(18);
    n_变动id       Number(18);
    n_开始原因     Number(3);
    d_开始时间     Date;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_008') = 0 Then
      Return;
    End If;
    Select ID, 护理等级id, 开始时间, 开始原因
    Into n_变动id, v_新护理等级id, d_开始时间, n_开始原因
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null And Nvl(附加床位, 0) = 0;
  
    Select Max(护理等级id)
    Into v_原护理等级id
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 = d_开始时间 And 终止原因 = n_开始原因 And Nvl(附加床位, 0) = 0;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_008',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID>' || '<原护理等级ID>' ||
                                 v_原护理等级id || '</原护理等级ID>' || '<新护理等级ID>' || v_新护理等级id || '</新护理等级ID>' || '<变动ID>' ||
                                 n_变动id || '</变动ID>' || '</root>');
  End Zlhis_Patient_008;
  --60.住院患者预出院
  Procedure Zlhis_Patient_009
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
    n_变动id 病人变动记录.Id%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_009') = 0 Then
      Return;
    End If;
    Select Max(ID)
    Into n_变动id
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null And Nvl(附加床位, 0) = 0;
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_009',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><变动ID>' || n_变动id ||
                                 '</变动ID></root>');
  End Zlhis_Patient_009;
  --61.住院患者出院
  Procedure Zlhis_Patient_010
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_010',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID></root>');
  End Zlhis_Patient_010;
  --62.住院患者新生儿登记
  Procedure Zlhis_Patient_011
  (
    病人id_In   In 病案主页.病人id%Type,
    主页id_In   In 病案主页.主页id%Type,
    婴儿序号_In 病人医嘱记录.婴儿%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_011',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><婴儿序号>' || 婴儿序号_In ||
                                 '</婴儿序号></root>');
  End Zlhis_Patient_011;
  --63.住院患者转入科室
  Procedure Zlhis_Patient_012
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
    v_转出科室id Number(18);
    v_转入科室id Number(18);
    n_变动id     Number(18);
    n_开始原因   Number(3);
    d_开始时间   Date;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_012') = 0 Then
      Return;
    End If;
    Select ID, 科室id, 开始时间, 开始原因
    Into n_变动id, v_转入科室id, d_开始时间, n_开始原因
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null And Nvl(附加床位, 0) = 0;
  
    Select Max(科室id)
    Into v_转出科室id
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 = d_开始时间 And 终止原因 = n_开始原因 And Nvl(附加床位, 0) = 0;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_012',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID>' || '<转出科室ID>' ||
                                 v_转出科室id || '</转出科室ID>' || '<转入科室ID>' || v_转入科室id || '</转入科室ID>' || '<变动ID>' || n_变动id ||
                                 '</变动ID>' || '</root>');
  End Zlhis_Patient_012;
  --64.新生儿登记作废
  Procedure Zlhis_Patient_013
  (
    病人id_In   In 病案主页.病人id%Type,
    主页id_In   In 病案主页.主页id%Type,
    婴儿序号_In 病人医嘱记录.婴儿%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_013',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><婴儿序号>' || 婴儿序号_In ||
                                 '</婴儿序号></root>');
  End Zlhis_Patient_013;
  --65.门诊患者登记
  Procedure Zlhis_Patient_015(病人id_In In 病案主页.病人id%Type) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_015', '<root><病人ID>' || 病人id_In || '</病人ID></root>');
  End Zlhis_Patient_015;
  --66.患者信息修改
  Procedure Zlhis_Patient_016(病人id_In In 病案主页.病人id%Type) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_016', '<root><病人ID>' || 病人id_In || '</病人ID></root>');
  End Zlhis_Patient_016;

  --67.患者合并
  Procedure Zlhis_Patient_017
  (
    病人id_In   In 病案主页.病人id%Type,
    原病人id_In In 病案主页.病人id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_017',
                                '<root><病人ID>' || 病人id_In || '</病人ID><原病人ID>' || 原病人id_In || '</原病人ID></root>');
  End Zlhis_Patient_017;

  --69.住院患者转入病区
  Procedure Zlhis_Patient_026
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
    v_转出病区id Number(18);
    v_转入病区id Number(18);
    n_变动id     Number(18);
    n_开始原因   Number(3);
    d_开始时间   Date;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_026') = 0 Then
      Return;
    End If;
    Select ID, 病区id, 开始时间, 开始原因
    Into n_变动id, v_转入病区id, d_开始时间, n_开始原因
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null And Nvl(附加床位, 0) = 0;
  
    Select Max(病区id)
    Into v_转出病区id
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 = d_开始时间 And 终止原因 = n_开始原因 And Nvl(附加床位, 0) = 0;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_026',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID>' || '<转出病区ID>' ||
                                 v_转出病区id || '</转出病区ID>' || '<转入病区ID>' || v_转入病区id || '</转入病区ID>' || '<变动ID>' || n_变动id ||
                                 '</变动ID>' || '</root>');
  End Zlhis_Patient_026;

  Procedure Zlhis_Patient_028(病人id_In In 病案主页.病人id%Type) Is 
    v_姓名     病人信息.姓名%Type; 
    v_性别     病人信息.性别%Type; 
    v_年龄     病人信息.年龄%Type; 
    v_门诊号   病人信息.门诊号%Type; 
    v_身份证号 病人信息.身份证号%Type; 
    v_出生日期 varchar2(50); 
  Begin 
    If p_Msg_Using('ZLHIS_PATIENT_028') = 0 Then 
      Return; 
    End If; 
    Select 姓名, 性别, 年龄, To_Char(出生日期, 'yyyymmdd'), 门诊号, 身份证号 
    Into v_姓名, v_性别, v_年龄, v_出生日期, v_门诊号, v_身份证号 
    From 病人信息 
    Where 病人id = 病人id_In; 
 
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_028', 
                                '<root><病人ID>' || 病人id_In || '</病人ID><姓名>' || v_姓名 || '</姓名>' || '<性别>' || v_性别 || 
                                 '</性别>' || '<年龄>' || v_年龄 || '</年龄>' || '<出生日期>' || v_出生日期 || '</出生日期>' || '<门诊号>' || 
                                 v_门诊号 || '</门诊号>' || '<身份证号>' || v_身份证号 || '</身份证号>' || '</root>'); 
  End Zlhis_Patient_028; 

  --血库:科室配血完成
  Procedure Zlhis_Blood_001(医嘱id_In In 病人医嘱记录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If 医嘱id_In Is Not Null Then
      v_Value := '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_BLOOD_001', v_Value);
    End If;
  End Zlhis_Blood_001;

  --血库:科室拒绝配血
  Procedure Zlhis_Blood_002(医嘱id_In In 病人医嘱记录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If 医嘱id_In Is Not Null Then
      v_Value := '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_BLOOD_002', v_Value);
    End If;
  End Zlhis_Blood_002;

  --70.检验报告审核
  Procedure Zlhis_Lis_001(标本id_In In 检验标本记录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If 标本id_In Is Not Null Then
      v_Value := '<root><标本ID>' || 标本id_In || '</标本ID><系统>1</系统></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_001', v_Value);
    End If;
  End Zlhis_Lis_001;
  --71.检验报告审核撤消
  Procedure Zlhis_Lis_002(标本id_In In 检验标本记录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If 标本id_In Is Not Null Then
      v_Value := '<root><标本ID>' || 标本id_In || '</标本ID><系统>1</系统></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_002', v_Value);
    End If;
  End Zlhis_Lis_002;
  --73.检验标本条码打印
  Procedure Zlhis_Lis_004
  (
    样本条码_In In 病人医嘱发送.样本条码%Type,
    医嘱id_In   In 病人医嘱发送.医嘱id%Type,
    医嘱ids_In  In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If p_Msg_Using('ZLHIS_LIS_004') = 0 Then
      Return;
    End If;
    If 医嘱id_In Is Not Null Then
      v_Value := '<root><样本条码>' || 样本条码_In || '</样本条码><医嘱ID>' || 医嘱id_In || '</医嘱ID><系统>1</系统></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_004', v_Value);
    Else
      For R In (Select '<root><样本条码>' || 样本条码_In || '</样本条码><医嘱ID>' || 医嘱id_In || '</医嘱ID><系统>1</系统></root>' As Xml_Value
                From 病人医嘱发送
                Where 医嘱id In (Select Column_Value From Table(f_Num2list(医嘱ids_In)))) Loop
        b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_004', r.Xml_Value);
      End Loop;
    End If;
  End Zlhis_Lis_004;
  --74.检验标本条码打印撤销
  Procedure Zlhis_Lis_005
  (
    样本条码_In In 病人医嘱发送.样本条码%Type,
    医嘱id_In   In 病人医嘱发送.医嘱id%Type,
    医嘱ids_In  In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If p_Msg_Using('ZLHIS_LIS_005') = 0 Then
      Return;
    End If;
    If 医嘱id_In Is Not Null Then
      v_Value := '<root><样本条码>' || 样本条码_In || '</样本条码><医嘱ID>' || 医嘱id_In || '</医嘱ID><系统>1</系统></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_005', v_Value);
    Else
      For R In (Select '<root><样本条码>' || 样本条码_In || '</样本条码><医嘱ID>' || 医嘱id_In || '</医嘱ID><系统>1</系统></root>' As Xml_Value
                From 病人医嘱发送
                Where 医嘱id In (Select Column_Value From Table(f_Num2list(医嘱ids_In)))) Loop
        b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_005', r.Xml_Value);
      End Loop;
    End If;
  End Zlhis_Lis_005;
  --75.检验标本核收
  Procedure Zlhis_Lis_006(标本id_In In 检验标本记录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If 标本id_In Is Not Null Then
      v_Value := '<root><标本ID>' || 标本id_In || '</标本ID><系统>1</系统></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_006', v_Value);
    End If;
  End Zlhis_Lis_006;
  --76.检验标本核收撤销
  Procedure Zlhis_Lis_007(标本id_In In 检验标本记录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If 标本id_In Is Not Null Then
      v_Value := '<root><标本ID>' || 标本id_In || '</标本ID><系统>1</系统></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_007', v_Value);
    End If;
  End Zlhis_Lis_007;
  --77.检验标本拒收
  Procedure Zlhis_Lis_008(标本id_In In 检验标本记录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If 标本id_In Is Not Null Then
      v_Value := '<root><标本ID>' || 标本id_In || '</标本ID><系统>1</系统></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_008', v_Value);
    End If;
  End Zlhis_Lis_008;

End b_Message;
/

--122812:焦博,2018-09-13,票据入库记录增加字段批次,票据领用记录增加字段入库ID
Create Or Replace Procedure Zl_消费卡领用记录_Insert
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
  入库id_In   消费卡领用记录.入库id%Type := Null
) Is
  v_Err_Msg Varchar2(500);
  Err_Item Exception;

  n_Count  Number(18);
  n_剩余数 消费卡入库记录.剩余数量%Type;
Begin

  For r_入库 In (Select ID, 前缀文本, 接口编号, 开始卡号, Nvl(终止卡号, 开始卡号) As 终止卡号
               From 消费卡入库记录
               Where ID = Nvl(入库id_In, 0)) Loop
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
    Where ID = Nvl(入库id_In, 0) And 接口编号 = 接口编号_In
    Returning Nvl(剩余数量, 0) Into n_剩余数;
  
    If n_剩余数 < 0 Then
      v_Err_Msg := '入库卡片的剩余票据数不足，请检查！';
      Raise Err_Item;
    End If;
    If n_剩余数 = 0 Then
      Update 消费卡入库记录 Set 是否存在卡 = 0 Where ID = Nvl(入库id_In, 0) And 接口编号 = 接口编号_In;
    End If;
  End Loop;

  Insert Into 消费卡领用记录
    (ID, 接口编号, 领用人, 前缀文本, 开始卡号, 终止卡号, 使用方式, 登记时间, 登记人, 剩余数量, 批次, 签字人, 签字时间, 入库id)
  Values
    (Id_In, 接口编号_In, 领用人_In, 前缀文本_In, 开始卡号_In, 终止卡号_In, 使用方式_In, 登记时间_In, 登记人_In, 剩余数量_In, 批次_In, 签字人_In,
     Decode(签字人_In, Null, Null + Sysdate, Sysdate), 入库id_In);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_消费卡领用记录_Insert;
/

Create Or Replace Procedure Zl_消费卡入库记录_Delete(Id_In In 消费卡入库记录.Id%Type) Is
  v_Err_Msg varchar2(100);
  Err_Item Exception;

  n_Count Number(2);
Begin
  --检查是否存在报损记录 
  Select Count(1) Into n_Count From 消费卡报损记录 Where 入库id = Id_In And Rownum < 2;
  If n_Count = 1 Then
    v_Err_Msg := '该入库批次已经存在报损记录，不能再进行删除！';
    Raise Err_Item;
  End If;

  --检查是否存在使用记录,如果存在,不不允许进行删除 
  Select Count(1) Into n_Count From 消费卡领用记录 Where 入库id = Id_In And Rownum < 2;
  If n_Count = 1 Then
    v_Err_Msg := '该入库批次已经存在领用记录，不能再进行删除！';
    Raise Err_Item;
  End If;

  --可能领用的已经转入到历史数据空间,因此检查数量是否相等,不等,肯定已经使用 
  Select Count(1) Into n_Count From 消费卡入库记录 Where ID = Id_In And Nvl(入库数量, 0) > Nvl(剩余数量, 0);
  If n_Count = 1 Then
    v_Err_Msg := '该入库批次已经被使用，不能再进行删除！';
    Raise Err_Item;
  End If;

  Delete From 消费卡入库记录 Where ID = Id_In;
  If Sql%NotFound Then
    v_Err_Msg := '由于并发原因，该入库批次可能已经被他人删除，不能再进行删除！';
    Raise Err_Item;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_消费卡入库记录_Delete;
/

Create Or Replace Procedure Zl_票据入库记录_Delete(Id_In In 票据入库记录.ID%Type) Is
  v_Err_Msg varchar2(100);
  Err_Item Exception;
  n_Exists Number(2);
Begin

  --检查是否存在报损记录
  Begin
    Select 1 Into n_Exists From 票据报损记录 Where 入库id = Id_In And Rownum = 1;
  Exception
    When Others Then
      n_Exists := 0;
  End;
  If n_Exists = 1 Then
    v_Err_Msg := '[ZLSOFT]票据已经存在报损记录,不能再进行删除![ZLSOFT]';
    Raise Err_Item;
  End If;

  --检查是否存在使用记录,如果存在,不不允许进行删除
  Begin
    Select 1 Into n_Exists From 票据领用记录 Where 入库id = Id_In And Rownum = 1;
  Exception
    When Others Then
      n_Exists := 0;
  End;
  If n_Exists = 1 Then
    v_Err_Msg := '[ZLSOFT]票据已经存在领用记录,不能再进行删除![ZLSOFT]';
    Raise Err_Item;
  End If;

  --可能领用的已经转入到历史数据空间,因此检查数量是否相等,不等,肯定已经使用
  Begin
    Select 1 Into n_Exists From 票据入库记录 Where ID = Id_In And Nvl(入库数量, 0) > Nvl(剩余数量, 0);
  Exception
    When Others Then
      n_Exists := 0;
  End;

  If n_Exists = 1 Then
    v_Err_Msg := '[ZLSOFT]票据已经被使用,不能再进行删除![ZLSOFT]';
    Raise Err_Item;
  End If;

  Delete From 票据入库记录 Where ID = Id_In;
  If Sql%NotFound Then
    v_Err_Msg := '[ZLSOFT]由于并发原因,该入库批次可能已经被他人删除,不能再进行删除![ZLSOFT]';
    Raise Err_Item;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_票据入库记录_Delete;
/

--130300:秦龙,2018-09-11,区分规格和品种设置的“严格控制用法用量”
Create Or Replace Procedure Zl_用法用量_Update
(
  药名id_In           In 诊疗用法用量.项目id%Type,
  过敏试验id_In       In Varchar2, --以"|"分隔的过敏实验的内容 
  处方限量_In         In 药品特性.处方限量%Type,
  疗程_In             In 诊疗用法用量.疗程%Type,
  用法用量_In         In Varchar2, --以"|"分隔的用法用量内容，每条记录按"用法ID^频次^成人剂量^小儿剂量^医生嘱托"组织 
  方式_In             In Number := 0, --0-诊疗项目本身,1-当前类别;2-特定分类项目 
  类别_In             In Varchar2 := '0',
  分类id_In           In 诊疗项目目录.分类id%Type := 0,
  药品id_In           In 药品规格.药品id%Type := 0,
  严格控制用法用量_In In 药品规格.严格控制用法用量%Type := 0
) Is
  --药品id_in :不等于空说明是规格，否则为品种或者分类 
  v_Records      Varchar2(4000);
  v_Currrec      Varchar2(1000);
  v_Fields       Varchar2(1000);
  v_用法id       诊疗用法用量.用法id%Type;
  v_频次         诊疗用法用量.频次%Type;
  v_成人剂量     诊疗用法用量.成人剂量%Type;
  v_小儿剂量     诊疗用法用量.小儿剂量%Type;
  v_医生嘱托     诊疗用法用量.医生嘱托%Type;
  v_Ddd值        诊疗用法用量.Ddd值%Type;
  v_性质         诊疗用法用量.性质%Type;
  v_是否皮试     药品特性.是否皮试%Type;
  n_药名id       药品规格.药名id%Type;
  n_药品用法用量 Number; --0-药品用法用量中不存在数据，1-药品用法用量中存在数据 

  Cursor c_Item Is
    Select i.Id
    From 诊疗项目目录 I, 药品特性 T, (Select 药品剂型 From 药品特性 Where 药名id = 药名id_In) C
    Where i.Id = t.药名id And t.药品剂型 = c.药品剂型 And i.分类id = 分类id_In And i.Id <> 药名id_In;
Begin
  --品种和分类 
  If 药品id_In = 0 Then
    For r_Item In (Select ID
                   From 诊疗项目目录
                   Where (方式_In = 0 And ID = 药名id_In) Or (方式_In = 1 And 类别 = 类别_In) Or
                         (分类id In (Select ID From 诊疗分类目录 Start With ID = 分类id_In Connect By Prior ID = 上级id))) Loop
      --过敏试验
      If 过敏试验id_In Is Not Null Then
        v_是否皮试 := 1;
      Else
        v_是否皮试 := 0;
      End If;
    
      Update 药品特性
      Set 处方限量 = 处方限量_In, 是否皮试 = v_是否皮试, 严格控制用法用量 = 严格控制用法用量_In
      Where 药名id = r_Item.Id;
    
      For r_Spec In (Select b.药品id From 药品规格 B Where b.药名id = r_Item.Id) Loop
        Delete From 药品用法用量 Where 药品id = r_Spec.药品id And 性质 = 0;
      End Loop;
    
      Delete From 诊疗用法用量 Where 项目id = r_Item.Id And 性质 = 0;
    
      v_Records := 过敏试验id_In;
    
      While v_Records Is Not Null Loop
        v_Currrec := Substr(v_Records, 1, Instr(v_Records, '|') - 1);
        v_Fields  := v_Currrec;
        v_用法id  := To_Number(v_Fields);
        Insert Into 诊疗用法用量 (项目id, 性质, 用法id) Values (r_Item.Id, 0, v_用法id);
        v_Records := Replace('|' || v_Records, '|' || v_Currrec || '|');
      
        --判断在药品用法用量中是否存在对应该品种的规格的数据 
        Begin
          n_药品用法用量 := 0;
          Select 1
          Into n_药品用法用量
          From 药品规格 A, 药品用法用量 B
          Where a.药品id = b.药品id And a.药名id = r_Item.Id And b.用法id = v_用法id And Rownum <= 1;
        Exception
          When Others Then
            n_药品用法用量 := 0;
        End;
      
        If n_药品用法用量 = 0 Then
          For r_药品id In (Select 药品id From 药品规格 Where 药名id = r_Item.Id) Loop
            Insert Into 药品用法用量 (药品id, 用法id, 性质) Values (r_药品id.药品id, v_用法id, 0);
          End Loop;
        End If;
      End Loop;
    
      --用法用量   Select 药品id From 药品规格 Where 药名id = r_Item.Id
      For r_Spec In (Select b.药品id From 药品规格 B Where b.药名id = r_Item.Id) Loop
        Delete From 药品用法用量 Where 药品id = r_Spec.药品id And 性质 > 0;
      End Loop;
      Delete From 诊疗用法用量 Where 项目id = r_Item.Id And 性质 > 0;
    
      If 用法用量_In Is Null Then
        v_Records := Null;
      Else
        v_Records := 用法用量_In || '|';
      End If;
      v_性质 := 0;
      While v_Records Is Not Null Loop
        v_Currrec  := Substr(v_Records, 1, Instr(v_Records, '|') - 1);
        v_Fields   := v_Currrec;
        v_用法id   := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
        v_Fields   := Substr(v_Fields, Instr(v_Fields, '^') + 1);
        v_频次     := Substr(v_Fields, 1, Instr(v_Fields, '^') - 1);
        v_Fields   := Substr(v_Fields, Instr(v_Fields, '^') + 1);
        v_成人剂量 := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
        v_Fields   := Substr(v_Fields, Instr(v_Fields, '^') + 1);
        v_小儿剂量 := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
        v_Fields   := Substr(v_Fields, Instr(v_Fields, '^') + 1);
        v_医生嘱托 := Substr(v_Fields, 1, Instr(v_Fields, '^') - 1);
        v_Fields   := Substr(v_Fields, Instr(v_Fields, '^') + 1);
        v_Ddd值    := To_Number(v_Fields);
        v_性质     := v_性质 + 1;
        Insert Into 诊疗用法用量
          (项目id, 性质, 用法id, 频次, 成人剂量, 小儿剂量, 医生嘱托, 疗程, Ddd值)
        Values
          (r_Item.Id, v_性质, v_用法id, v_频次, v_成人剂量, v_小儿剂量, v_医生嘱托, 疗程_In, v_Ddd值);
        If 分类id_In <> 0 Then
          For t_Item In c_Item Loop
            Delete From 诊疗用法用量 Where 项目id = t_Item.Id And 用法id = v_用法id And 性质 > 0;
            Insert Into 诊疗用法用量 (项目id, 性质, 用法id, 频次) Values (t_Item.Id, v_性质, v_用法id, v_频次);
          End Loop;
        End If;
      
        --判断在药品用法用量中是否存在对应该品种的规格的数据 
        Begin
          n_药品用法用量 := 0;
          Select 1
          Into n_药品用法用量
          From 药品规格 A, 药品用法用量 B
          Where a.药品id = b.药品id And a.药名id = r_Item.Id And b.用法id = v_用法id And Rownum <= 1;
        Exception
          When Others Then
            n_药品用法用量 := 0;
        End;
      
        If n_药品用法用量 = 0 Then
          For r_药品id In (Select 药品id From 药品规格 Where 药名id = r_Item.Id) Loop
            Insert Into 药品用法用量
              (药品id, 用法id, 频次, 成人剂量, 小儿剂量, 医生嘱托, 疗程, Ddd值, 性质)
            Values
              (r_药品id.药品id, v_用法id, v_频次, v_成人剂量, v_小儿剂量, v_医生嘱托, 疗程_In, v_Ddd值, 1);
          End Loop;
        End If;
      
        v_Records := Replace('|' || v_Records, '|' || v_Currrec || '|');
      End Loop;
    End Loop;
  Else
    --规格 
    --过敏试验
    If 过敏试验id_In Is Not Null Then
      v_是否皮试 := 1;
    Else
      v_是否皮试 := 0;
    End If;
    Select 药名id Into n_药名id From 药品规格 Where 药品id = 药品id_In;
    Update 药品特性 Set 处方限量 = 处方限量_In, 是否皮试 = v_是否皮试 Where 药名id = n_药名id;
    Update 药品规格 Set 严格控制用法用量 = 严格控制用法用量_In Where 药品id = 药品id_In;
  
    Delete From 药品用法用量 Where 药品id = 药品id_In And 性质 = 0;
    v_Records := 过敏试验id_In;
  
    While v_Records Is Not Null Loop
      v_Currrec := Substr(v_Records, 1, Instr(v_Records, '|') - 1);
      v_Fields  := v_Currrec;
      v_用法id  := To_Number(v_Fields);
      Insert Into 药品用法用量 (药品id, 用法id, 性质) Values (药品id_In, v_用法id, 0);
      v_Records := Replace('|' || v_Records, '|' || v_Currrec || '|');
    End Loop;
  
    --用法用量
    Delete From 药品用法用量 Where 药品id = 药品id_In And 性质 > 0;
    If 用法用量_In Is Null Then
      v_Records := Null;
    Else
      v_Records := 用法用量_In || '|';
    End If;
    v_性质 := 0;
    While v_Records Is Not Null Loop
      v_Currrec  := Substr(v_Records, 1, Instr(v_Records, '|') - 1);
      v_Fields   := v_Currrec;
      v_用法id   := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
      v_Fields   := Substr(v_Fields, Instr(v_Fields, '^') + 1);
      v_频次     := Substr(v_Fields, 1, Instr(v_Fields, '^') - 1);
      v_Fields   := Substr(v_Fields, Instr(v_Fields, '^') + 1);
      v_成人剂量 := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
      v_Fields   := Substr(v_Fields, Instr(v_Fields, '^') + 1);
      v_小儿剂量 := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
      v_Fields   := Substr(v_Fields, Instr(v_Fields, '^') + 1);
      v_医生嘱托 := Substr(v_Fields, 1, Instr(v_Fields, '^') - 1);
      v_Fields   := Substr(v_Fields, Instr(v_Fields, '^') + 1);
      v_Ddd值    := To_Number(v_Fields);
      v_性质     := v_性质 + 1;
    
      Insert Into 药品用法用量
        (药品id, 用法id, 频次, 成人剂量, 小儿剂量, 医生嘱托, 疗程, Ddd值, 性质)
      Values
        (药品id_In, v_用法id, v_频次, v_成人剂量, v_小儿剂量, v_医生嘱托, 疗程_In, v_Ddd值, 1);
    
      v_Records := Replace('|' || v_Records, '|' || v_Currrec || '|');
    End Loop;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_用法用量_Update;
/

--129973:秦龙,2018-09-10,增加传参“是否易至跌倒”
Create Or Replace Procedure Zl_成药规格_Update
(
  药品id_In         In 药品规格.药品id%Type,
  编码_In           In 收费项目目录.编码%Type,
  规格_In           In 收费项目目录.规格%Type,
  产地_In           In 收费项目目录.产地%Type := Null,
  品名_In           In 收费项目别名.名称%Type := Null,
  拼音_In           In 收费项目别名.简码%Type := Null,
  五笔_In           In 收费项目别名.简码%Type := Null,
  数字码_In         In 收费项目别名.简码%Type := Null,
  标识码_In         In 药品规格.标识码%Type := Null,
  药品来源_In       In 药品规格.药品来源%Type := Null,
  批准文号_In       In 药品规格.批准文号%Type := Null,
  注册商标_In       In 药品规格.注册商标%Type := Null,
  售价单位_In       In 收费项目目录.计算单位%Type := Null,
  剂量系数_In       In 药品规格.剂量系数%Type := Null,
  门诊单位_In       In 药品规格.门诊单位%Type := Null,
  门诊包装_In       In 药品规格.门诊包装%Type := Null,
  住院单位_In       In 药品规格.住院单位%Type := Null,
  住院包装_In       In 药品规格.住院包装%Type := Null,
  药库单位_In       In 药品规格.药库单位%Type := Null,
  药库包装_In       In 药品规格.药库包装%Type := Null,
  申领单位_In       In 药品规格.申领单位%Type := 1,
  申领阀值_In       In 药品规格.申领阀值%Type := Null,
  是否变价_In       In 收费项目目录.是否变价%Type := Null,
  指导批发价_In     In 药品规格.指导批发价%Type := Null,
  扣率_In           In 药品规格.扣率%Type := 95,
  指导零售价_In     In 药品规格.指导零售价%Type := Null,
  加成率_In         In 药品规格.加成率%Type := Null,
  管理费比例_In     In 药品规格.管理费比例%Type := Null,
  药价级别_In       In 药品规格.药价级别%Type := Null,
  费用类型_In       In 收费项目目录.费用类型%Type := Null,
  服务对象_In       In 收费项目目录.服务对象%Type := Null,
  Gmp认证_In        In 药品规格.Gmp认证%Type := 0,
  招标药品_In       In 药品规格.招标药品%Type := 0,
  屏蔽费别_In       In 收费项目目录.屏蔽费别%Type := 0,
  住院可否分零_In   In 药品规格.住院可否分零%Type := 0,
  药库分批_In       In 药品规格.药库分批%Type := Null,
  药房分批_In       In 药品规格.药房分批%Type := Null,
  最大效期_In       In 药品规格.最大效期%Type := Null,
  差价让利比_In     In 药品规格.差价让利比%Type := 0,
  成本价_In         In 药品规格.成本价%Type := 0,
  当前售价_In       In 收费价目.现价%Type := 0,
  收入id_In         In 收费价目.收入项目id%Type := Null,
  合同单位id_In     In 药品规格.合同单位id%Type := Null,
  说明_In           In 收费项目目录.说明%Type := Null,
  动态分零_In       In 药品规格.动态分零%Type := 0,
  发药类型_In       In 药品规格.发药类型%Type := Null,
  备选码_In         In 收费项目目录.备选码%Type := Null,
  增值税率_In       In 药品规格.增值税率%Type := Null,
  基本药物_In       In 药品规格.基本药物%Type := Null,
  站点_In           In 收费项目目录.站点%Type := Null,
  是否常备_In       In 药品规格.是否常备%Type := Null,
  存储温度_In       In 输液药品属性.存储温度%Type := Null,
  存储条件_In       In 输液药品属性.存储条件%Type := Null,
  配药类型_In       In 输液药品属性.配药类型%Type := Null,
  是否不予配置_In   In 输液药品属性.是否不予配置%Type := Null,
  容量_In           In 药品规格.容量%Type := Null,
  病案费目_In       In 收费项目目录.病案费目%Type := Null,
  门诊可否分零_In   In 药品规格.门诊可否分零%Type := 0,
  Ddd值_In          In 药品规格.Ddd值%Type := 0,
  高危药品_In       药品规格.高危药品%Type := Null,
  送货单位_In       In 药品规格.送货单位%Type := Null,
  送货包装_In       In 药品规格.送货包装%Type := Null,
  输液注意事项_In   In 输液药品属性.输液注意事项%Type := Null,
  是否摆药_In       In 药品规格.是否摆药%Type := Null,
  是否零差价管理_In In 药品规格.是否零差价管理%Type := Null,
  本位码_In         In 药品规格.本位码%Type := Null,
  是否易至跌倒_In   In 药品规格.是否易至跌倒%Type := Null
) Is
  v_药名id   诊疗项目目录.Id%Type;
  v_名称     诊疗项目目录.名称%Type;
  v_是否变价 收费项目目录.是否变价%Type; --允许定价药品随时改为时价，时价药品只能在未发生的情况下修改为定价，其它情况不允许修改定价属性 
  v_发生     Number(2);
  Err_Notfind Exception;
  v_No           收费价目.No%Type;
  v_Temp         收费项目目录.病案费目%Type;
  v_病案费目     收费项目目录.病案费目%Type;
  n_指导差价率   药品规格.指导差价率%Type;
  n_药品上次售价 药品规格.上次售价%Type;
  n_零售金额     药品库存.实际金额%Type;
  n_收发id       药品收发记录.Id%Type;
  n_流通金额小数 Number;
  n_序号         Number(8);
  Classid        Number(18); --入出类别
  v_Billno       药品收发记录.No%Type; --调价单号
  n_价格id       收费价目.Id%Type;
  n_收费价目现价 收费价目.现价%Type;
  n_原价         药品价格记录.原价%Type;
  n_药品价格记录 Number(1);
  v_类别         收费项目目录.类别%Type;
  --定价->时价后更新药品价格记录的值

  Cursor c_Priceadjust Is
    Select s.药品id, s.库房id, Nvl(s.批次, 0) As 批次, s.上次供应商id As 供应商id, s.上次批号 As 批号, s.效期, s.上次产地 As 产地,
           Nvl(s.实际数量, 0) As 实际数量, s.上次扣率 As 扣率, Nvl(s.实际金额, 0) As 实际金额, Nvl(s.实际差价, 0) As 实际差价, Nvl(s.零售价, 0) As 零售价,
           s.平均成本价, s.灭菌效期, s.批准文号, s.上次生产日期 As 生产日期
    From 药品库存 S
    Where s.药品id = 药品id_In And s.性质 = 1 
    Order By s.药品id, s.批次, s.库房id;

  r_Priceadjust c_Priceadjust%RowType;

Begin
  v_病案费目 := 病案费目_In;
  --判断病案费目 
  If v_病案费目 Is Null Then
    If 收入id_In Is Not Null Then
      Begin
        Select 病案费目 Into v_Temp From 收入项目 Where ID = 收入id_In;
      Exception
        When Others Then
          v_Temp := Null;
      End;
      If v_Temp Is Not Null Then
        v_病案费目 := v_Temp;
      End If;
    End If;
  End If;
  --通用名称 
  Select ID, 名称
  Into v_药名id, v_名称
  From 诊疗项目目录
  Where ID = (Select 药名id From 药品规格 Where 药品id = 药品id_In);
  --取原始的定价属性 
  Select 是否变价 Into v_是否变价 From 收费项目目录 Where ID = 药品id_In;
  --规格信息 
  Update 收费项目目录
  Set 编码 = 编码_In, 名称 = v_名称, 规格 = 规格_In, 产地 = 产地_In, 计算单位 = 售价单位_In, 费用类型 = 费用类型_In, 服务对象 = 服务对象_In, 屏蔽费别 = 屏蔽费别_In,
      病案费目 = v_病案费目, 说明 = 说明_In, 备选码 = 备选码_In, 站点 = 站点_In
  Where ID = 药品id_In
  Returning 类别 Into v_类别;
  If Sql%RowCount = 0 Then
    Raise Err_Notfind;
  End If;
  n_指导差价率 := (1 - 1 / (1 + 加成率_In / 100)) * 100;
  Update 药品规格
  Set 标识码 = 标识码_In, 药品来源 = 药品来源_In, 批准文号 = 批准文号_In, 注册商标 = 注册商标_In, 剂量系数 = 剂量系数_In, 门诊单位 = 门诊单位_In, 门诊包装 = 门诊包装_In,
      住院单位 = 住院单位_In, 住院包装 = 住院包装_In, 药库单位 = 药库单位_In, 药库包装 = 药库包装_In, 申领单位 = 申领单位_In, 申领阀值 = 申领阀值_In, 指导批发价 = 指导批发价_In,
      扣率 = 扣率_In, 指导零售价 = 指导零售价_In, 指导差价率 = n_指导差价率, 管理费比例 = 管理费比例_In, 药价级别 = 药价级别_In, 住院可否分零 = 住院可否分零_In,
      药库分批 = 药库分批_In, 药房分批 = 药房分批_In, 最大效期 = 最大效期_In, 招标药品 = 招标药品_In, Gmp认证 = Gmp认证_In, 差价让利比 = 差价让利比_In,
      合同单位id = 合同单位id_In, 动态分零 = 动态分零_In, 发药类型 = 发药类型_In, 增值税率 = 增值税率_In, 基本药物 = 基本药物_In, 是否常备 = 是否常备_In, 容量 = 容量_In,
      门诊可否分零 = 门诊可否分零_In, Ddd值 = Ddd值_In, 高危药品 = 高危药品_In, 送货单位 = 送货单位_In, 送货包装 = 送货包装_In, 加成率 = 加成率_In, 是否摆药 = 是否摆药_In,
      是否零差价管理 = 是否零差价管理_In, 本位码 = 本位码_In, 是否易至跌倒 = 是否易至跌倒_In
  Where 药品id = 药品id_In;

  --朱玉宝修改：建立药品（西成药、中成药）时，缺省服务对象为门诊和住院，因此修改规格药品时，不再根据规格药品的服务对象更新药品的服务对象 
  --诊疗项目服务对象的更改 
  --select nvl(sum(distinct I.服务对象),0) into v_对象 
  --from 收费项目目录 I,药品规格 S 
  --where I.ID=S.药品ID and S.药名ID=v_药名ID; 
  --update 诊疗项目目录 
  --set 服务对象=decode(v_对象,0,0,1,1,2,2,3) 
  --where ID=v_药名ID; 

  --别名的处理 
  If 数字码_In Is Null Then
    Delete 收费项目别名 Where 收费细目id = 药品id_In And 性质 = 1 And 码类 = 3;
  Else
    Update 收费项目别名 Set 名称 = v_名称, 简码 = 数字码_In Where 收费细目id = 药品id_In And 性质 = 1 And 码类 = 3;
    If Sql%RowCount = 0 Then
      Insert Into 收费项目别名 (收费细目id, 名称, 性质, 简码, 码类) Values (药品id_In, v_名称, 1, 数字码_In, 3);
    End If;
  End If;
  If 品名_In Is Null Then
    Delete 收费项目别名 Where 收费细目id = 药品id_In And 性质 = 3;
  Else
    If 拼音_In Is Null Then
      Delete 收费项目别名 Where 收费细目id = 药品id_In And 性质 = 3 And 码类 = 1;
    Else
      Update 收费项目别名 Set 名称 = 品名_In, 简码 = 拼音_In Where 收费细目id = 药品id_In And 性质 = 3 And 码类 = 1;
      If Sql%RowCount = 0 Then
        Insert Into 收费项目别名 (收费细目id, 名称, 性质, 简码, 码类) Values (药品id_In, 品名_In, 3, 拼音_In, 1);
      End If;
    End If;
    If 五笔_In Is Null Then
      Delete 收费项目别名 Where 收费细目id = 药品id_In And 性质 = 3 And 码类 = 2;
    Else
      Update 收费项目别名 Set 名称 = 品名_In, 简码 = 五笔_In Where 收费细目id = 药品id_In And 性质 = 3 And 码类 = 2;
      If Sql%RowCount = 0 Then
        Insert Into 收费项目别名 (收费细目id, 名称, 性质, 简码, 码类) Values (药品id_In, 品名_In, 3, 五笔_In, 2);
      End If;
    End If;
  End If;

  --定价信息：如果已经有发生，则不允许直接更改这些信息 
  Select Nvl(Count(*), 0) Into v_发生 From 药品收发记录 Where 药品id = 药品id_In And Rownum < 2;
  If v_发生 = 0 Then
    Update 药品规格 Set 成本价 = 成本价_In Where 药品id = 药品id_In;
    If 收入id_In Is Not Null Then
      Update 收费价目
      Set 现价 = 当前售价_In, 收入项目id = 收入id_In, 变动原因 = 1, 调价说明 = '修改定价', 调价人 = User
      Where 收费细目id = 药品id_In And (Sysdate Between 执行日期 And 终止日期 Or Sysdate >= 执行日期 And 终止日期 Is Null) And 变动原因 = 1;
      If Sql%RowCount = 0 Then
        v_No := Nextno(9);
        Insert Into 收费价目
          (ID, 原价id, 收费细目id, 原价, 现价, 收入项目id, 变动原因, 调价说明, 调价人, 执行日期, 终止日期, NO, 序号)
        Values
          (收费价目_Id.Nextval, Null, 药品id_In, 0, 当前售价_In, 收入id_In, 1, '新增定价', User, Sysdate,
           To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
      End If;
    End If;
  Else
    --发生业务数据时，不能修改价格但是可以修改收入项目 
    Update 收费价目
    Set 收入项目id = 收入id_In
    Where 收费细目id = 药品id_In And (Sysdate Between 执行日期 And 终止日期 Or Sysdate >= 执行日期 And 终止日期 Is Null) And 变动原因 = 1;
  End If;

  --时价->定价
  If v_是否变价 = 1 And 是否变价_In = 0 Then
    Update 收费项目目录 Set 是否变价 = 是否变价_In Where ID = 药品id_In;
  
    Begin
      Select 现价, ID As 价格id
      Into n_收费价目现价, n_价格id
      From 收费价目
      Where 收费细目id = 药品id_In And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And 变动原因 = 1;
    Exception
      When Others Then
        n_收费价目现价 := Null;
        n_价格id       := Null;
    End;
  
    Begin
      Select 上次售价 Into n_药品上次售价 From 药品规格 Where 药品id = 药品id_In;
    Exception
      When Others Then
        n_药品上次售价 := Null;
    End;
  
    If n_药品上次售价 Is Null Then
      n_药品上次售价 := n_收费价目现价;
    End If;
  
    Zl_收费价目_Stop(药品id_In, Sysdate - 1 / 24 / 60 / 60);
    v_No := Nextno(9);
    Insert Into 收费价目
      (ID, 原价id, 收费细目id, 原价, 现价, 收入项目id, 变动原因, 调价说明, 调价人, 执行日期, 终止日期, NO, 序号)
    Values
      (收费价目_Id.Nextval, n_价格id, 药品id_In, n_收费价目现价, n_药品上次售价, 收入id_In, 1, '时价转定价', Zl_Username, Sysdate,
       To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
  
    Select 精度 Into n_流通金额小数 From 药品卫材精度 Where 类别 = 1 And 内容 = 4 And 单位 = 5;
  
    --取入出类别ID
    Select 类别id Into Classid From 药品单据性质 Where 单据 = 13;
  
    n_序号   := 0;
    v_Billno := Null;
  
    For r_Priceadjust In c_Priceadjust Loop
      If n_药品上次售价 <> r_Priceadjust.零售价 Then
        If v_Billno Is Null Then
          Select Nextno(147) Into v_Billno From Dual;
        End If;
        n_序号 := n_序号 + 1;
        Select 药品收发记录_Id.Nextval Into n_收发id From Dual;
        n_零售金额 := Round(n_药品上次售价 * r_Priceadjust.实际数量, n_流通金额小数) -
                  Round(r_Priceadjust.零售价 * r_Priceadjust.实际数量, n_流通金额小数);
        --产生调价影响记录
        Insert Into 药品收发记录
          (ID, 记录状态, 单据, NO, 序号, 入出类别id, 药品id, 批次, 批号, 效期, 产地, 付数, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 扣率, 零售金额, 差价, 摘要, 填制人,
           填制日期, 库房id, 入出系数, 价格id, 审核人, 审核日期, 单量, 频次, 供药单位id)
        Values
          (n_收发id, 1, 13, v_Billno, n_序号, Classid, r_Priceadjust.药品id, r_Priceadjust.批次, r_Priceadjust.批号,
           r_Priceadjust.效期, r_Priceadjust.产地, 1, r_Priceadjust.实际数量, 0, r_Priceadjust.零售价, 0, n_药品上次售价,
           r_Priceadjust.扣率, n_零售金额, n_零售金额, '时价转定价', Zl_Username, Sysdate, r_Priceadjust.库房id, 1, n_价格id, Zl_Username,
           Sysdate, r_Priceadjust.实际金额, r_Priceadjust.实际差价, r_Priceadjust.供应商id);
      
        Zl_药品库存_Update(n_收发id, 2, 0);
      End If;
    End Loop;
  
    --定价->时价
  Elsif v_是否变价 = 0 And 是否变价_In = 1 Then
    Update 收费项目目录 Set 是否变价 = 是否变价_In Where ID = 药品id_In;
    Begin
      Select 现价, ID As 价格id
      Into n_收费价目现价, n_价格id
      From 收费价目
      Where 收费细目id = 药品id_In And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And 变动原因 = 1;
    Exception
      When Others Then
        n_收费价目现价 := Null;
        n_价格id       := Null;
    End;
  
    Zl_收费价目_Stop(药品id_In, Sysdate - 1 / 24 / 60 / 60);
    v_No := Nextno(9);
    Insert Into 收费价目
      (ID, 原价id, 收费细目id, 原价, 现价, 收入项目id, 变动原因, 调价说明, 调价人, 执行日期, 终止日期, NO, 序号)
    Values
      (收费价目_Id.Nextval, n_价格id, 药品id_In, n_收费价目现价, n_收费价目现价, 收入id_In, 1, '定价转时价', Zl_Username, Sysdate,
       To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
  
    For r_Priceadjust In c_Priceadjust Loop
      n_药品价格记录 := 0;
      Begin
        Select 1, 现价
        Into n_药品价格记录, n_原价
        From 药品价格记录
        Where 药品id = r_Priceadjust.药品id And 库房id = r_Priceadjust.库房id And Nvl(批次, 0) = r_Priceadjust.批次 And 记录状态 = 1 And
              价格类型 = 1;
      Exception
        When Others Then
          n_药品价格记录 := 0;
          n_原价         := n_收费价目现价;
      End;
    
      If n_药品价格记录 = 1 Then
        Zl_药品价格记录_Stop(1, r_Priceadjust.库房id, r_Priceadjust.药品id, r_Priceadjust.批次, Sysdate - 1 / 24 / 60 / 60, 2);
      End If;
      Zl_药品价格记录_Insert(0, 1, r_Priceadjust.库房id, r_Priceadjust.药品id, r_Priceadjust.批次, n_原价, n_收费价目现价, Sysdate, '定价转时价',
                       Zl_Username, Null, r_Priceadjust.供应商id, r_Priceadjust.批号, r_Priceadjust.效期, r_Priceadjust.产地,
                       r_Priceadjust.灭菌效期, Null, Null, Null, Null, 1);
    
      Update 药品库存
      Set 零售价 = n_收费价目现价
      Where 性质 = 1 And 库房id = r_Priceadjust.库房id And 药品id = r_Priceadjust.药品id And Nvl(批次, 0) = r_Priceadjust.批次;
    
    End Loop;
  End If;

  --药品生产商比较增加 
  If 产地_In Is Not Null Then
    Update 药品生产商 Set 名称 = 产地_In Where 名称 = 产地_In;
    If Sql%RowCount = 0 Then
      Insert Into 药品生产商
        (编码, 名称, 简码)
        Select Nvl(Max(To_Number(编码)), 0) + 1, 产地_In, zlSpellCode(产地_In, 10) From 药品生产商;
    End If;
  End If;

  --修改输液药品属性 
  Update 输液药品属性
  Set 存储温度 = 存储温度_In, 存储条件 = 存储条件_In, 配药类型 = 配药类型_In, 是否不予配置 = 是否不予配置_In, 输液注意事项 = 输液注意事项_In
  Where 药品id = 药品id_In;

  If Sql%NotFound Then
    Insert Into 输液药品属性
      (药品id, 存储温度, 存储条件, 配药类型, 是否不予配置, 输液注意事项)
    Values
      (药品id_In, 存储温度_In, 存储条件_In, 配药类型_In, 是否不予配置_In, 输液注意事项_In);
  End If;

  --药品精度调整(零差价模式时)
  Zl_药品卫材精度_零差价调整;

  b_Message.Zlhis_Dict_036(v_类别, 药品id_In);
Exception
  When Err_Notfind Then
    Raise_Application_Error(-20101, '[ZLSOFT]该规格不存在，可能已被其他用户删除！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_成药规格_Update;
/

Create Or Replace Procedure Zl_成药规格_Insert
(
  药名id_In         In 药品规格.药名id%Type,
  药品id_In         In 药品规格.药品id%Type,
  编码_In           In 收费项目目录.编码%Type,
  规格_In           In 收费项目目录.规格%Type,
  产地_In           In 收费项目目录.产地%Type := Null,
  品名_In           In 收费项目别名.名称%Type := Null,
  拼音_In           In 收费项目别名.简码%Type := Null,
  五笔_In           In 收费项目别名.简码%Type := Null,
  数字码_In         In 收费项目别名.简码%Type := Null,
  标识码_In         In 药品规格.标识码%Type := Null,
  药品来源_In       In 药品规格.药品来源%Type := Null,
  批准文号_In       In 药品规格.批准文号%Type := Null,
  注册商标_In       In 药品规格.注册商标%Type := Null,
  售价单位_In       In 收费项目目录.计算单位%Type := Null,
  剂量系数_In       In 药品规格.剂量系数%Type := Null,
  门诊单位_In       In 药品规格.门诊单位%Type := Null,
  门诊包装_In       In 药品规格.门诊包装%Type := Null,
  住院单位_In       In 药品规格.住院单位%Type := Null,
  住院包装_In       In 药品规格.住院包装%Type := Null,
  药库单位_In       In 药品规格.药库单位%Type := Null,
  药库包装_In       In 药品规格.药库包装%Type := Null,
  申领单位_In       In 药品规格.申领单位%Type := 1,
  申领阀值_In       In 药品规格.申领阀值%Type := Null,
  是否变价_In       In 收费项目目录.是否变价%Type := Null,
  指导批发价_In     In 药品规格.指导批发价%Type := Null,
  扣率_In           In 药品规格.扣率%Type := 95,
  指导零售价_In     In 药品规格.指导零售价%Type := Null,
  加成率_In         In 药品规格.加成率%Type := Null,
  管理费比例_In     In 药品规格.管理费比例%Type := Null,
  药价级别_In       In 药品规格.药价级别%Type := Null,
  费用类型_In       In 收费项目目录.费用类型%Type := Null,
  服务对象_In       In 收费项目目录.服务对象%Type := Null,
  Gmp认证_In        In 药品规格.Gmp认证%Type := 0,
  招标药品_In       In 药品规格.招标药品%Type := 0,
  屏蔽费别_In       In 收费项目目录.屏蔽费别%Type := 0,
  住院可否分零_In   In 药品规格.住院可否分零%Type := 0,
  药库分批_In       In 药品规格.药库分批%Type := Null,
  药房分批_In       In 药品规格.药房分批%Type := Null,
  最大效期_In       In 药品规格.最大效期%Type := Null,
  差价让利比_In     In 药品规格.差价让利比%Type := 0,
  成本价_In         In 药品规格.成本价%Type := 0,
  当前售价_In       In 收费价目.现价%Type := 0,
  收入id_In         In 收费价目.收入项目id%Type := Null,
  合同单位id_In     In 药品规格.合同单位id%Type := Null,
  说明_In           In 收费项目目录.说明%Type := Null,
  动态分零_In       In 药品规格.动态分零%Type := 0,
  发药类型_In       In 药品规格.发药类型%Type := Null,
  备选码_In         In 收费项目目录.备选码%Type := Null,
  增值税率_In       In 药品规格.增值税率%Type := Null,
  基本药物_In       In 药品规格.基本药物%Type := Null,
  站点_In           In 收费项目目录.站点%Type := Null,
  是否常备_In       In 药品规格.是否常备%Type := Null,
  存储温度_In       In 输液药品属性.存储温度%Type := Null,
  存储条件_In       In 输液药品属性.存储条件%Type := Null,
  配药类型_In       In 输液药品属性.配药类型%Type := Null,
  是否不予配置_In   In 输液药品属性.是否不予配置%Type := Null,
  容量_In           In 药品规格.容量%Type := Null,
  病案费目_In       In 收费项目目录.病案费目%Type := Null,
  门诊可否分零_In   In 药品规格.门诊可否分零%Type := 0,
  Ddd值_In          In 药品规格.Ddd值%Type := 0,
  高危药品_In       In 药品规格.高危药品%Type := Null,
  送货单位_In       In 药品规格.送货单位%Type := Null,
  送货包装_In       In 药品规格.送货包装%Type := Null,
  输液注意事项_In   In 输液药品属性.输液注意事项%Type := Null,
  是否摆药_In       In 药品规格.是否摆药%Type := Null,
  是否零差价管理_In In 药品规格.是否零差价管理%Type := Null,
  本位码_In         In 药品规格.本位码%Type := Null,
  是否易至跌倒_In   In 药品规格.是否易至跌倒%Type := Null
) Is

  v_类别       诊疗项目目录.类别%Type;
  v_名称       诊疗项目目录.名称%Type;
  v_Kind       Varchar2(20);
  v_No         收费价目.No%Type;
  v_Temp       收费项目目录.病案费目%Type;
  v_病案费目   收费项目目录.病案费目%Type;
  n_指导差价率 药品规格.指导差价率%Type;

  --盘点库房的工作性质 
  Cursor c_Storageid Is
    Select Distinct 部门id From 部门性质说明 Where 工作性质 Like v_Kind Or 工作性质 = '制剂室';
  r_Storageid c_Storageid%RowType;
Begin
  v_病案费目 := 病案费目_In;
  --判断病案费目 
  If v_病案费目 Is Null Then
    If 收入id_In Is Not Null Then
      Begin
        Select 病案费目 Into v_Temp From 收入项目 Where ID = 收入id_In;
      Exception
        When Others Then
          v_Temp := Null;
      End;
      If v_Temp Is Not Null Then
        v_病案费目 := v_Temp;
      End If;
    End If;
  End If;
  --类别和名称 
  Select 类别, 名称 Into v_类别, v_名称 From 诊疗项目目录 Where ID = 药名id_In;
  n_指导差价率 := (1 - 1 / (1 + 加成率_In / 100)) * 100;
  --规格信息 
  Insert Into 收费项目目录
    (类别, ID, 编码, 名称, 规格, 产地, 计算单位, 费用类型, 服务对象, 屏蔽费别, 是否变价, 建档时间, 撤档时间, 说明, 备选码, 站点, 病案费目)
  Values
    (v_类别, 药品id_In, 编码_In, v_名称, 规格_In, 产地_In, 售价单位_In, 费用类型_In, 服务对象_In, 屏蔽费别_In, 是否变价_In, Sysdate,
     To_Date('3000-01-01', 'YYYY-MM-DD'), 说明_In, 备选码_In, 站点_In, v_病案费目);
  Insert Into 药品规格
    (药名id, 药品id, 标识码, 药品来源, 批准文号, 注册商标, 剂量系数, 门诊单位, 门诊包装, 住院单位, 住院包装, 药库单位, 药库包装, 申领单位, 申领阀值, 指导批发价, 扣率, 指导零售价, 指导差价率,
     管理费比例, 药价级别, 成本价, Gmp认证, 招标药品, 差价让利比, 住院可否分零, 药库分批, 药房分批, 最大效期, 合同单位id, 动态分零, 发药类型, 增值税率, 基本药物, 是否常备, 容量, 门诊可否分零,
     Ddd值, 高危药品, 送货单位, 送货包装, 加成率, 是否摆药, 是否零差价管理, 本位码, 是否易至跌倒)
  Values
    (药名id_In, 药品id_In, 标识码_In, 药品来源_In, 批准文号_In, 注册商标_In, 剂量系数_In, 门诊单位_In, 门诊包装_In, 住院单位_In, 住院包装_In, 药库单位_In, 药库包装_In,
     申领单位_In, 申领阀值_In, 指导批发价_In, 扣率_In, 指导零售价_In, n_指导差价率, 管理费比例_In, 药价级别_In, 成本价_In, Gmp认证_In, 招标药品_In, 差价让利比_In,
     住院可否分零_In, 药库分批_In, 药房分批_In, 最大效期_In, 合同单位id_In, 动态分零_In, 发药类型_In, 增值税率_In, 基本药物_In, 是否常备_In, 容量_In, 门诊可否分零_In,
     Ddd值_In, 高危药品_In, 送货单位_In, 送货包装_In, 加成率_In, 是否摆药_In, 是否零差价管理_In, 本位码_In, 是否易至跌倒_In);

  --朱玉宝修改：建立药品（西成药、中成药）时，缺省服务对象为门诊和住院，因此建立规格药品时，不再根据规格药品的服务对象更新药品的服务对象 
  --诊疗项目服务对象的更改 
  --select nvl(sum(distinct I.服务对象),0) into v_对象 
  --from 收费项目目录 I,药品规格 S 
  --where I.ID=S.药品ID and S.药名ID=药名ID_IN; 
  --update 诊疗项目目录 
  --set 服务对象=decode(v_对象,0,0,1,1,2,2,3) 
  --where ID=药名ID_IN; 

  --别名的处理 
  Insert Into 收费项目别名
    (收费细目id, 名称, 性质, 简码, 码类)
    Select 药品id_In, 名称, 性质, 简码, 码类 From 诊疗项目别名 Where 诊疗项目id = 药名id_In;
  If 数字码_In Is Not Null Then
    Insert Into 收费项目别名 (收费细目id, 名称, 性质, 简码, 码类) Values (药品id_In, v_名称, 1, 数字码_In, 3);
  End If;
  If (品名_In Is Not Null) And (拼音_In Is Not Null) Then
    Insert Into 收费项目别名 (收费细目id, 名称, 性质, 简码, 码类) Values (药品id_In, 品名_In, 3, 拼音_In, 1);
  End If;
  If (品名_In Is Not Null) And (五笔_In Is Not Null) Then
    Insert Into 收费项目别名 (收费细目id, 名称, 性质, 简码, 码类) Values (药品id_In, 品名_In, 3, 五笔_In, 2);
  End If;

  --定价信息 
  If 收入id_In Is Not Null Then
    v_No := Nextno(9);
    Insert Into 收费价目
      (ID, 原价id, 收费细目id, 原价, 现价, 收入项目id, 变动原因, 调价说明, 调价人, 执行日期, 终止日期, NO, 序号)
    Values
      (收费价目_Id.Nextval, Null, 药品id_In, 0, 当前售价_In, 收入id_In, 1, '新增定价', User, Sysdate,
       To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
  End If;

  --药品生产商比较增加 
  If 产地_In Is Not Null Then
    Update 药品生产商 Set 名称 = 产地_In Where 名称 = 产地_In;
    If Sql%RowCount = 0 Then
      Insert Into 药品生产商
        (编码, 名称, 简码)
        Select Nvl(Max(To_Number(编码)), 0) + 1, 产地_In, zlSpellCode(产地_In, 10) From 药品生产商;
    End If;
  End If;

  --插入该规格的服务科室 
  Insert Into 收费执行科室
    (收费细目id, 病人来源, 开单科室id, 执行科室id)
    Select 药品id_In, 病人来源, 开单科室id, 执行科室id From 诊疗执行科室 Where 诊疗项目id = 药名id_In;

  --插入盘点属性 

  If v_类别 = 5 Then
    v_Kind := '西药%';
  Else
    v_Kind := '成药%';
  End If;

  For r_Storageid In c_Storageid Loop
    Insert Into 药品储备限额
      (库房id, 药品id, 上限, 下限, 盘点属性, 库房货位)
    Values
      (r_Storageid.部门id, 药品id_In, 0, 0, '1111', Null);
  End Loop;

  --插入输液药品属性 
  Insert Into 输液药品属性
    (药品id, 存储温度, 存储条件, 配药类型, 是否不予配置, 输液注意事项)
  Values
    (药品id_In, 存储温度_In, 存储条件_In, 配药类型_In, 是否不予配置_In, 输液注意事项_In);

  --药品精度调整(零差价模式时)
  Zl_药品卫材精度_零差价调整;

  b_Message.Zlhis_Dict_035(v_类别, 药品id_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_成药规格_Insert;
/



------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0029' Where 编号=&n_System;
Commit;