----------------------------------------------------------------------------------------------------------------
--本脚本支持从ZLHIS+ v10.35.90升级到 v10.35.90
--请以数据空间的所有者登录PLSQL并执行下列脚本
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------
--结构修正部份
------------------------------------------------------------------------------



------------------------------------------------------------------------------
--数据修正部份
------------------------------------------------------------------------------
--138877:蒋廷中,2019-04-23,增加留观病人转住院病人锚点定义
Insert Into Zlmsg_Lists (Bz_Type, Code, Name, Key_Define, Note, Using)
Select '病人', 'ZLHIS_PATIENT_029', '留观病人转住院病人', '<root><病人ID></病人ID><主页ID></主页ID><变动ID></变动ID></root>', '留观病人转住院病人时', 1 From Dual;

-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--139063:冉俊明,2019-04-25,门诊留观病人按门诊流程就诊
Create Or Replace Procedure Zl_门诊记帐记录_Delete
(
  No_In           门诊费用记录.No%Type,
  序号_In         Varchar2,
  操作员编号_In   门诊费用记录.操作员编号%Type,
  操作员姓名_In   门诊费用记录.操作员姓名%Type,
  输液配药检查_In Number := 1,
  登记时间_In     住院费用记录.登记时间%Type := Sysdate
) As
  --功能：冲销一张门诊记帐单据中指定序号行 
  --序号：格式如"1,3,5,7,8",或"1:2:33456,3:2,5:2,7:2,8:2",冒号前面的数字表示行号,中间的数字表示退的数量,后面的数字表示配药记录的ID,目前仅在销帐审核时才传入 
  --      为空表示冲销所有可冲销行 

  --该游标为要退费单据的所有原始记录 
  Cursor c_Bill(n_标志 Number) Is
    Select a.Id, a.价格父号, a.序号, a.执行状态, a.收费类别, a.医嘱序号, a.病人id, a.主页id, a.收入项目id, a.开单部门id, a.执行部门id, a.病人病区id, a.病人科室id,
           a.实收金额, Decode(a.记录状态, 0, 1, 0) As 划价, j.诊疗类别, m.跟踪在用
    From 门诊费用记录 A, 病人医嘱记录 J, 材料特性 M
    Where a.医嘱序号 = j.Id(+) And a.收费细目id + 0 = m.材料id(+) And a.No = No_In And a.记录性质 = 2 And a.记录状态 In (0, 1, 3) And
          a.门诊标志 = n_标志
    Order By a.收费细目id, a.序号;

  --该游标用于处理费用记录序号 
  Cursor c_Serial Is
    Select 序号, 价格父号 From 门诊费用记录 Where NO = No_In And 记录性质 = 2 And 记录状态 In (0, 1, 3) Order By 序号;
  l_划价 t_Numlist := t_Numlist();

  v_医嘱ids  Varchar2(4000);
  n_父号     门诊费用记录.价格父号%Type;
  n_门诊标志 门诊费用记录.门诊标志%Type;

  --部分退费计算变量 
  n_剩余数量 Number;
  n_剩余应收 Number;
  n_剩余实收 Number;
  n_剩余统筹 Number;

  n_准退数量 Number;
  n_退费次数 Number;
  n_退费数量 Number;
  n_部分销帐 Number;

  n_应收金额 Number;
  n_实收金额 Number;
  n_统筹金额 Number;

  v_序号   Varchar2(4000);
  v_配药id Varchar2(4000);
  v_Tmp    Varchar2(4000);

  n_未执行数量 药品收发记录.实际数量%Type;
  n_已执行数量 药品收发记录.实际数量%Type;

  n_Dec Number;

  n_Count   Number;
  d_Curdate Date;
  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin
  --销帐审核时,非药品会传入行号的销帐数量 
  If Not 序号_In Is Null Then
    If Instr(序号_In, ':') > 0 Then
      --格式：1:2:33456,3:2,5:2,7:2,8:2 
      For c_序号 In (Select C1, C2 From Table(f_Str2list2(序号_In, ',', ':'))) Loop
        v_序号 := v_序号 || ',' || c_序号.C1;
        If Instr(c_序号.C2, ':') > 0 Then
          v_配药id := v_配药id || ',' || Substr(c_序号.C2, Instr(c_序号.C2, ':') + 1);
        End If;
      End Loop;
      v_序号   := Substr(v_序号, 2);
      v_配药id := Substr(v_配药id, 2);
    Else
      v_序号 := 序号_In;
    End If;
  End If;

  --是否已经全部完全执行(只是整张单据的检查) 
  Select Nvl(Count(1), 0), Max(Nvl(门诊标志, 1))
  Into n_Count, n_门诊标志
  From 门诊费用记录
  Where NO = No_In And 记录性质 = 2 And 记录状态 In (0, 1, 3) And Nvl(执行状态, 0) <> 1;
  If n_Count = 0 Then
    v_Err_Msg := '该单据中的项目已经全部完全执行！';
    Raise Err_Item;
  End If;

  If Nvl(n_门诊标志, 0) = 0 Then
    n_门诊标志 := 1;
  End If;

  --未完全执行的项目是否有剩余数量(只是整张单据的检查) 
  Select Nvl(Count(1), 0)
  Into n_Count
  From (Select 序号, Sum(数量) As 剩余数量
         From (Select 记录状态, Nvl(价格父号, 序号) As 序号, Avg(Nvl(付数, 1) * 数次) As 数量
                From 门诊费用记录
                Where NO = No_In And 记录性质 = 2 And 门诊标志 = n_门诊标志 And
                      Nvl(价格父号, 序号) In
                      (Select Nvl(价格父号, 序号)
                       From 门诊费用记录
                       Where NO = No_In And 记录性质 = 2 And 门诊标志 = n_门诊标志 And 记录状态 In (0, 1, 3) And Nvl(执行状态, 0) <> 1)
                Group By 记录状态, Nvl(价格父号, 序号))
         Group By 序号
         Having Sum(数量) <> 0);
  If n_Count = 0 Then
    v_Err_Msg := '该单据中未完全执行部分项目剩余数量为零,没有可以销帐的费用！';
    Raise Err_Item;
  End If;

  --------------------------------------------------------------------------------- 
  --公用变量 
  Select Nvl(登记时间_In, Sysdate) Into d_Curdate From Dual;

  --金额小数位数 
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  --循环处理每行费用(收入项目行) 
  For r_Bill In c_Bill(n_门诊标志) Loop
    If Instr(',' || v_序号 || ',', ',' || Nvl(r_Bill.价格父号, r_Bill.序号) || ',') > 0 Or v_序号 Is Null Then
      If Nvl(r_Bill.执行状态, 0) <> 1 Then
        --求剩余数量,剩余应收,剩余实收 
        Select Sum(Nvl(付数, 1) * 数次), Sum(应收金额), Sum(实收金额), Sum(统筹金额)
        Into n_剩余数量, n_剩余应收, n_剩余实收, n_剩余统筹
        From 门诊费用记录
        Where NO = No_In And 记录性质 = 2 And 序号 = r_Bill.序号;
      
        n_部分销帐 := 0;
        n_退费数量 := 0;
        If n_剩余数量 = 0 Then
          If v_序号 Is Not Null Then
            v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已经全部销帐！';
            Raise Err_Item;
          End If;
          --情况：未限定行号,原始单据中的该笔已经全部销帐(执行状态=0的一种可能) 
        Else
          If Instr(序号_In, ':') > 0 Then
            Select Max(a.C2) Into v_Tmp From Table(f_Str2list2(序号_In, ',', ':')) A Where a.C1 = r_Bill.序号;
            If Instr(v_Tmp, ':') > 0 Then
              n_退费数量 := Substr(v_Tmp, 1, Instr(v_Tmp, ':') - 1);
            Else
              n_退费数量 := To_Number(v_Tmp);
            End If;
            n_部分销帐 := 1;
          End If;
        
          --准销数量(非药品项目为剩余数量,原始数量) 
          If Instr(',4,5,6,7,', r_Bill.收费类别) = 0 Or (r_Bill.收费类别 = '4' And Nvl(r_Bill.跟踪在用, 0) = 0) Then
            --@@@ 
            --非药品部分(以具体医嘱执行为准进行检查) 
            --: 1.存在医嘱发送的,则以医嘱执行为准(但不能包含:检查;检验;手术;麻醉及输血) 
            --: 2.对于病人医吃计价中的收费方式为:0-正常收取 的,才支持部分退;如果是其他的,则只能全退 
            --: 3.不存在医嘱的,则以剩余数量为准 
            n_Count := 0;
            If Instr(',C,D,F,G,K,', ',' || r_Bill.诊疗类别 || ',') = 0 And r_Bill.诊疗类别 Is Not Null Then
              Select Nvl(Sum(数量), 0), Count(*)
              Into n_准退数量, n_Count
              From (Select j.医嘱序号 As 医嘱id, j.收费细目id, Nvl(j.付数, 1) * Nvl(j.数次, 1) As 数量
                     From 门诊费用记录 J, 病人医嘱记录 M
                     Where j.医嘱序号 = m.Id And j.No = No_In And j.记录性质 = 2 And j.序号 = r_Bill.序号 And j.记录状态 In (1, 3) And
                           Exists
                      (Select 1
                            From 病人医嘱发送 A
                            Where a.医嘱id = j.医嘱序号 And Nvl(a.执行状态, 0) <> 1 And a.No || '' = No_In) And Exists
                      (Select 1
                            From 病人医嘱计价 A
                            Where a.医嘱id = j.医嘱序号 And a.收费细目id = j.收费细目id And Nvl(a.收费方式, 0) = 0) And j.价格父号 Is Null And
                           Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And
                           (j.记录状态 In (1, 3) And Not Exists
                            (Select 1
                             From 药品收发记录
                             Where 费用id = j.Id And Instr(',8,9,10,21,24,25,26,', ',' || 单据 || ',') > 0) Or
                            j.记录状态 = 2 And Not Exists
                            (Select 1 From 药品收发记录 Where NO = No_In And 单据 In (8, 24) And 药品id = j.收费细目id))
                     Union All
                     Select a.医嘱id, a.收费细目id, -1 * Nvl(a.数量, 1) * Nvl(c.本次数次, 1) As 数量
                     From 病人医嘱计价 A, 病人医嘱发送 B, 病人医嘱执行 C, 门诊费用记录 J, 病人医嘱记录 M
                     Where a.医嘱id = b.医嘱id And b.医嘱id = c.医嘱id And Nvl(a.收费方式, 0) = 0 And b.发送号 = c.发送号 And a.医嘱id = m.Id And
                           Nvl(c.执行结果, 1) = 1 And Nvl(b.执行状态, 0) <> 1 And a.医嘱id = j.医嘱序号 And a.收费细目id = j.收费细目id And
                           j.No = No_In And j.记录性质 = 2 And j.序号 = r_Bill.序号 And j.记录状态 In (1, 3) And j.价格父号 Is Null And
                           Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And Not Exists
                      (Select 1
                            From 药品收发记录
                            Where 费用id = j.Id And Instr(',8,9,10,21,24,25,26,', ',' || 单据 || ',') > 0) And Not Exists
                      (Select 1 From 材料特性 Where 材料id = j.收费细目id And Nvl(跟踪在用, 0) = 1)
                     Union All
                     Select a.医嘱id, a.收费细目id, 0 As 数量
                     From 病人医嘱计价 A, 门诊费用记录 J, 病人医嘱记录 M
                     Where a.医嘱id = m.Id And a.医嘱id = j.医嘱序号 And a.收费细目id = j.收费细目id And Nvl(a.收费方式, 0) <> 0 And
                           j.No = No_In And j.记录性质 = 2 And Nvl(j.执行状态, 0) = 2 And Not Exists
                      (Select 1 From 材料特性 Where 材料id = j.收费细目id And Nvl(跟踪在用, 0) = 1) And
                           Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0);
            End If;
          
            If Nvl(n_Count, 0) = 0 Then
              n_准退数量 := n_剩余数量;
            End If;
          Else
            Select Sum(Nvl(付数, 1) * 实际数量)
            Into n_准退数量
            From 药品收发记录
            Where NO = No_In And 单据 In (9, 25) And Mod(记录状态, 3) = 1 And 审核人 Is Null And 费用id = r_Bill.Id;
          
            --不跟踪在用的卫生材料 
            If r_Bill.收费类别 = '4' And Nvl(n_准退数量, 0) = 0 Then
              n_准退数量 := n_剩余数量;
            End If;
          End If;
        
          If Nvl(n_退费数量, 0) = 0 Then
            n_退费数量 := n_准退数量;
          Else
            If n_准退数量 < n_退费数量 Then
              v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用准退数量不足本次销帐数量！';
              Raise Err_Item;
            End If;
          End If;
        
          --金额=剩余金额*(准退数/剩余数) 
          n_应收金额 := Round(n_剩余应收 * (n_退费数量 / n_剩余数量), n_Dec);
          n_实收金额 := Round(n_剩余实收 * (n_退费数量 / n_剩余数量), n_Dec);
          n_统筹金额 := Round(n_剩余统筹 * (n_退费数量 / n_剩余数量), n_Dec);
        
          If Nvl(r_Bill.划价, 0) = 0 Then
            --该笔项目第几次销帐 
            Select Nvl(Max(Abs(执行状态)), 0) + 1
            Into n_退费次数
            From 门诊费用记录
            Where NO = No_In And 记录性质 = 2 And 记录状态 = 2 And 序号 = r_Bill.序号;
          
            --插入退费记录 
            Insert Into 门诊费用记录
              (ID, NO, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 婴儿费, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id, 收费类别,
               收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 划价人,
               执行人, 执行状态, 执行时间, 操作员编号, 操作员姓名, 发生时间, 登记时间, 保险项目否, 保险大类id, 统筹金额, 记帐单id, 摘要, 保险编码, 是否急诊, 结论, 挂号id, 主页id,
               病人病区id)
              Select 病人费用记录_Id.Nextval, NO, 记录性质, 2, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 婴儿费, 姓名, 性别, 年龄, 标识号, 付款方式, 费别,
                     病人科室id, 收费类别, 收费细目id, 计算单位, Decode(Sign(n_退费数量 - Nvl(付数, 1) * 数次), 0, 付数, 1), 发药窗口,
                     Decode(Sign(n_退费数量 - Nvl(付数, 1) * 数次), 0, -1 * 数次, -1 * n_退费数量), 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用,
                     标准单价, -1 * n_应收金额, -1 * n_实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人, -1 * n_退费次数, 执行时间, 操作员编号_In,
                     操作员姓名_In, 发生时间, d_Curdate, 保险项目否, 保险大类id, -1 * n_统筹金额, 记帐单id, 摘要, 保险编码, 是否急诊, 结论, 挂号id, 主页id,
                     病人病区id
              From 门诊费用记录
              Where ID = r_Bill.Id;
          
            --病人余额 
            If n_门诊标志 <> 4 Then
              Update 病人余额
              Set 费用余额 = Nvl(费用余额, 0) - n_实收金额
              Where 病人id = r_Bill.病人id And 性质 = 1 And 类型 = 1;
              If Sql%RowCount = 0 Then
                Insert Into 病人余额
                  (病人id, 性质, 类型, 费用余额, 预交余额)
                Values
                  (r_Bill.病人id, 1, 1, -1 * n_实收金额, 0);
              End If;
            End If;
          
            --病人未结费用 
            Update 病人未结费用
            Set 金额 = Nvl(金额, 0) - n_实收金额
            Where 病人id = r_Bill.病人id And Nvl(主页id, 0) = Nvl(r_Bill.主页id, 0) And Nvl(病人病区id, 0) = Nvl(r_Bill.病人病区id, 0) And
                  Nvl(病人科室id, 0) = Nvl(r_Bill.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(r_Bill.开单部门id, 0) And
                  Nvl(执行部门id, 0) = Nvl(r_Bill.执行部门id, 0) And 收入项目id + 0 = r_Bill.收入项目id And 来源途径 + 0 = n_门诊标志;
            If Sql%RowCount = 0 Then
              Insert Into 病人未结费用
                (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
              Values
                (r_Bill.病人id, r_Bill.主页id, r_Bill.病人病区id, r_Bill.病人科室id, r_Bill.开单部门id, r_Bill.执行部门id, r_Bill.收入项目id,
                 n_门诊标志, -1 * n_实收金额);
            End If;
          
            --标记原费用记录 
            --执行状态:全部退完(准退数=剩余数)标记为0,否则标记为1 
            If Instr(',4,5,6,7,', r_Bill.收费类别) = 0 Then
              --一般情况非药品和卫材的项目,不存在部分销帐的情况,只有销帐申请和销帐审核时,才会出现部分销帐,所以 
              --执行状态只有两种:0.未执行;1已执行; 
              --由于在销帐审核过程中将已执行强制改为了2部分执行,因此需要在此处改为1已执行.未执行的不变. 
              Update 门诊费用记录
              Set 记录状态 = 3, 执行状态 = Decode(Sign(n_退费数量 - n_剩余数量), 0, 0, Decode(执行状态, 2, 1, 执行状态))
              Where ID = r_Bill.Id;
            Else
              Select Nvl(Sum(Decode(审核人, Null, 1, 0) * Nvl(付数, 1) * 实际数量), 0),
                     Nvl(Sum(Decode(审核人, Null, 0, 1) * Nvl(付数, 1) * 实际数量), 0)
              Into n_未执行数量, n_已执行数量
              From 药品收发记录
              Where NO = No_In And 单据 In (9, 10, 25, 26) And 费用id = r_Bill.Id;
            
              Update 门诊费用记录
              Set 记录状态 = 3,
                  执行状态 = Decode(Sign(n_退费数量 - n_剩余数量), 0, 0,
                                 Decode(Sign(n_未执行数量 - n_退费数量), 1, Decode(n_已执行数量, 0, 0, 2), 1))
              Where ID = r_Bill.Id;
            End If;
          Else
            --划价记账单 
            If Nvl(n_部分销帐, 0) = 0 Then
              l_划价.Extend;
              l_划价(l_划价.Count) := r_Bill.Id;
            Else
              --更新数量 
              --划价的,先将相关的数据处理在内部表集中 
              Update 住院费用记录
              Set 付数 = 1, 数次 = Nvl(付数, 1) * 数次 - n_退费数量, 应收金额 = Nvl(应收金额, 0) - n_应收金额, 实收金额 = Nvl(实收金额, 0) - n_实收金额,
                  登记时间 = d_Curdate, 统筹金额 = Nvl(统筹金额, 0) - n_统筹金额
              Where ID = r_Bill.Id
              Returning 数次 Into n_剩余数量;
              If Nvl(n_剩余数量, 0) <= 0 Then
                l_划价.Extend;
                l_划价(l_划价.Count) := r_Bill.Id;
              End If;
            End If;
          
            If r_Bill.医嘱序号 Is Not Null Then
              If Instr(',' || Nvl(v_医嘱ids, '') || ',', ',' || r_Bill.医嘱序号 || ',') = 0 Then
                v_医嘱ids := Nvl(v_医嘱ids, '') || ',' || r_Bill.医嘱序号;
              End If;
            End If;
          End If;
        End If;
      Else
        If v_序号 Is Not Null Then
          v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已经完全执行,不能销帐！';
          Raise Err_Item;
        End If;
        --情况:没限定行号,原始单据中包括已经完全执行的 
      End If;
    End If;
  End Loop;

  --不存在配药ID,检查该药品是否在输液配药中心 
  If v_配药id Is Null And 输液配药检查_In = 1 Then
    For v_费用 In (Select ID
                 From 门诊费用记录
                 Where NO = No_In And 记录性质 = 2 And 记录状态 In (0, 1, 3) And 收费类别 In ('4', '5', '6', '7') And 门诊标志 = n_门诊标志 And
                       (Instr(',' || v_序号 || ',', ',' || 序号 || ',') > 0 Or v_序号 Is Null)) Loop
      Begin
        Select Count(1)
        Into n_Count
        From 输液配药内容 A, 药品收发记录 B
        Where a.收发id = b.Id And b.费用id = v_费用.Id And Instr(',8,9,10,21,24,25,26,', ',' || b.单据 || ',') > 0;
      Exception
        When Others Then
          n_Count := 0;
      End;
      If n_Count <> 0 Then
        v_Err_Msg := '存在已经进入输液配药中心的待销帐药品，无法完成销帐！';
        Raise Err_Item;
      End If;
    End Loop;
  End If;

  --------------------------------------------------------------------------------- 
  --药品相关处理:主要是对销帐审核有效.(可以是部分) 
  --必须按照“收费细目id”升序排序，防止并发锁“药品库存”表 
  For v_费用 In (Select ID, 序号
               From 门诊费用记录
               Where NO = No_In And 记录性质 = 2 And 记录状态 In (0, 1, 3) And 收费类别 In ('4', '5', '6', '7') And 门诊标志 = n_门诊标志 And
                     (Instr(',' || v_序号 || ',', ',' || 序号 || ',') > 0 Or v_序号 Is Null)
               Order By 收费细目id) Loop
    --根据费用ID来进行相关的处理 
    n_退费数量 := 0;
    If Instr(序号_In, ':') > 0 Then
      Select Max(a.C2) Into v_Tmp From Table(f_Str2list2(序号_In, ',', ':')) A Where a.C1 = v_费用.序号;
      If Instr(v_Tmp, ':') > 0 Then
        n_退费数量 := Substr(v_Tmp, 1, Instr(v_Tmp, ':') - 1);
      Else
        n_退费数量 := To_Number(v_Tmp);
      End If;
    End If;
    Zl_药品收发记录_销售退费(v_费用.Id, n_退费数量, v_配药id);
  End Loop;

  --删除划价记录 
  n_Count := l_划价.Count;
  Forall I In 1 .. l_划价.Count
    Delete From 门诊费用记录 Where ID = l_划价(I);

  --删除之后再统一调整序号 
  If n_Count > 0 Then
    n_Count := 1;
    For r_Serial In c_Serial Loop
      If r_Serial.价格父号 Is Null Then
        n_父号 := n_Count;
      End If;
    
      Update 门诊费用记录
      Set 序号 = n_Count, 价格父号 = Decode(价格父号, Null, Null, n_父号)
      Where NO = No_In And 记录性质 = 2 And 序号 = r_Serial.序号;
    
      Update 门诊费用记录 Set 从属父号 = n_Count Where NO = No_In And 记录性质 = 2 And 从属父号 = r_Serial.序号;
    
      n_Count := n_Count + 1;
    End Loop;
  End If;

  --整张单据全部冲完时，删除病人医嘱附费 
  For c_医嘱 In (Select Distinct 医嘱序号
               From 门诊费用记录
               Where NO = No_In And 记录性质 = 2 And 记录状态 = 3 And 医嘱序号 Is Not Null) Loop
    Select Nvl(Count(*), 0)
    Into n_Count
    From (Select 序号, Sum(数量) As 剩余数量
           From (Select 记录状态, Nvl(价格父号, 序号) As 序号, Avg(Nvl(付数, 1) * 数次) As 数量
                  From 门诊费用记录
                  Where 记录性质 = 2 And 医嘱序号 + 0 = c_医嘱.医嘱序号 And NO = No_In
                  Group By 记录状态, Nvl(价格父号, 序号))
           Group By 序号
           Having Sum(数量) <> 0);
  
    If n_Count = 0 Then
      Delete From 病人医嘱附费 Where 医嘱id = c_医嘱.医嘱序号 And 记录性质 = 2 And NO = No_In;
    End If;
  End Loop;

  If v_医嘱ids Is Not Null Then
    --医嘱处理 
    --场合_In    Integer:=0, --0:门诊;1-住院 
    --性质_In    Integer:=1, --1-收费单;2-记帐单 
    --操作_In    Integer:=0, --0:删除划价单;1-收费或记帐;2-退费或销帐 
    --No_In      门诊费用记录.No%Type, 
    --医嘱ids_In Varchar2 := Null 
    v_医嘱ids := Substr(v_医嘱ids, 2);
    Zl_医嘱发送_计费状态_Update(0, 2, 0, No_In, v_医嘱ids);
  Else
    Zl_医嘱发送_计费状态_Update(0, 2, 2, No_In, v_医嘱ids);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊记帐记录_Delete;
/

--140358:陈刘,2019-04-24,Zl_Getpatient获取病人婴儿手术信息时,出生日期没有索引导致全表扫描
Create Or Replace Function Zl_Getpatient
(
  Begintime_In In Date,
  Endtime_In   In Date,
  病区id_In    In 部门表.Id%Type,
  格式id_In    In 病人护理文件.格式id%Type,
  Type_In      In Varchar2,
  Typeall_In   In Number := 1,
  Split_In     In Varchar2 := ';'
) Return t_Numlist2
  Pipelined As
  n_Index  Number(1);
  v_Str    Varchar2(50);
  n_病人id Number;
  n_主页id Number;
  P        Number;
  Out_Rec  t_Numobj2 := t_Numobj2(Null, Null);

  --提取对应病区科室所有在院病人 
  Cursor c_List_All Is
    Select b.病人id, b.主页id
    From 病人信息 A, 病案主页 B, 在院病人 C
    Where a.病人id = b.病人id And a.主页id = b.主页id And Nvl(b.主页id, 0) <> 0 And Nvl(b.病案状态, 0) <> 5 And b.封存时间 Is Null And
          a.病人id = c.病人id And c.病区id = 病区id_In;

  --入院三天内的病人 
  Cursor c_List_Ry Is
    Select b.病人id, b.主页id
    From 病人信息 A, 病案主页 B, 在院病人 C
    Where a.病人id = b.病人id And a.主页id = b.主页id And Nvl(b.主页id, 0) <> 0 And Nvl(b.病案状态, 0) <> 5 And b.封存时间 Is Null And
          b.入院日期 Between Begintime_In And Endtime_In And a.病人id = c.病人id And c.病区id = 病区id_In;

  --手术三天内的病人 
  Cursor c_List_Ss Is
    Select b.病人id, b.主页id
    From 病人信息 A, 病案主页 B, 在院病人 F, 病人护理文件 C, 病人护理数据 D, 病人护理明细 E
    Where a.病人id = b.病人id And a.主页id = b.主页id And a.病人id = f.病人id And Nvl(b.主页id, 0) <> 0 And f.病区id = 病区id_In And
          Nvl(b.病案状态, 0) <> 5 And b.封存时间 Is Null And b.病人id = c.病人id And b.主页id = c.主页id And c.格式id = 格式id_In And
          c.Id = d.文件id And d.Id = e.记录id And e.记录类型 = 4 And e.项目名称 <> '分娩' And Nvl(e.复试合格, 0) <> 1 And e.终止版本 Is Null And
          d.发生时间 Between Begintime_In And Endtime_In
    
    Union
    --从医嘱中提取病人手术信息 
    Select b.病人id, b.主页id
    From 病人信息 A, 病案主页 B, 在院病人 F,
         (Select d.病人id, d.主页id
           From (Select Distinct a.病人id, a.主页id
                  From 病人医嘱记录 A, 诊疗项目目录 B
                  Where a.诊疗项目id = b.Id And a.诊疗类别 = 'F' And a.相关id Is Null And a.医嘱状态 In (3, 8) And
                        a.开始执行时间 Between Begintime_In And Endtime_In
                  Union
                  Select Distinct a.病人id, a.主页id
                  From 病人新生儿记录 A, 在院病人 F
                  Where a.病人id = f.病人id And a.主页id = f.主页id And f.病区id = 病区id_In And a.出生时间 Between Begintime_In And
                        Endtime_In) D
           Group By d.病人id, d.主页id) C
    Where a.病人id = b.病人id And a.主页id = b.主页id And a.病人id = f.病人id And Nvl(b.主页id, 0) <> 0 And f.病区id = 病区id_In And
          Nvl(b.病案状态, 0) <> 5 And b.封存时间 Is Null And b.病人id = c.病人id And b.主页id = c.主页id;

  --三天内体温存在超过37.5度的病人 
  Cursor c_List_Tw Is
    Select b.病人id, b.主页id
    From 病人信息 A, 病案主页 B, 在院病人 F, 病人护理文件 C, 病人护理数据 D, 病人护理明细 E
    Where a.病人id = b.病人id And a.主页id = b.主页id And a.病人id = f.病人id And Nvl(b.主页id, 0) <> 0 And f.病区id = 病区id_In And
          Nvl(b.病案状态, 0) <> 5 And b.封存时间 Is Null And b.病人id = c.病人id And b.主页id = c.主页id And c.格式id = 格式id_In And
          c.Id = d.文件id And d.Id = e.记录id And e.记录类型 = 1 And e.项目序号 = 1 And
          Length(Translate(e.记录内容, '-.0123456789' || e.记录内容, '-.0123456789')) = Length(e.记录内容) And
          Zl_To_Number(e.记录内容) >= 37.5 And e.终止版本 Is Null And d.发生时间 Between Begintime_In And Endtime_In;

  --危/重病人 
  Cursor c_List_Wz Is
    Select b.病人id, b.主页id
    From 病人信息 A, 病案主页 B, 在院病人 F
    Where a.病人id = b.病人id And a.主页id = b.主页id And a.病人id = f.病人id And Nvl(b.主页id, 0) <> 0 And f.病区id = 病区id_In And
          Nvl(b.病案状态, 0) <> 5 And b.封存时间 Is Null And Instr(',' || '危,重' || ',', ',' || b.当前病况 || ',') > 0;

  --转入三天内的病人 
  Cursor c_List_Zr Is
    Select b.病人id, b.主页id
    From 病人信息 A, 病案主页 B, 病人变动记录 C, 在院病人 F
    Where a.病人id = b.病人id And a.主页id = b.主页id And Nvl(b.主页id, 0) <> 0 And b.病人id = c.病人id And b.主页id = c.主页id And
          a.病人id = f.病人id And f.病区id = 病区id_In And Nvl(b.病案状态, 0) <> 5 And b.封存时间 Is Null And Nvl(c.附加床位, 0) = 0 And
          c.病区id + 0 = f.病区id And c.开始原因 In (3, 15) And c.开始时间 Is Not Null And b.状态 = 0 And c.开始时间 Between Begintime_In And
          Endtime_In;

  -- 一级及以上护理等级的病人 
  Cursor c_List_Yj Is
    Select b.病人id, b.主页id
    From 病人信息 A, 病案主页 B, 在院病人 F
    Where a.病人id = b.病人id And a.主页id = b.主页id And a.病人id = f.病人id And Nvl(b.主页id, 0) <> 0 And f.病区id = 病区id_In And
          Nvl(b.病案状态, 0) <> 5 And b.封存时间 Is Null And Zl_Patittendgrade(b.病人id, b.主页id) <= 1;

  --分娩后三天内的病人 
  Cursor c_List_Fm Is
    Select b.病人id, b.主页id
    From 病人信息 A, 病案主页 B, 在院病人 F, 病人护理文件 C, 病人护理数据 D, 病人护理明细 E
    Where a.病人id = b.病人id And a.主页id = b.主页id And a.病人id = f.病人id And
          Nvl(b.主页id,
              
              0) <> 0 And f.病区id = 病区id_In And Nvl(b.病案状态, 0) <> 5 And b.封存时间 Is Null And b.病人id = c.病人id And
          b.主页id = c.主页id And c.格式id = 格式id_In And c.Id = d.文件id And d.Id = e.记录id And e.记录类型 = 4 And e.项目名称 = '分娩' And
          e.终止版本 Is Null And d.发生时间 Between Begintime_In And Endtime_In;

  Type v_病人id_Type Is Table Of 病案主页.病人id%Type;
  v_病人id v_病人id_Type;
  Type v_主页id_Type Is Table Of 病案主页.主页id%Type;
  v_主页id v_主页id_Type;
Begin

  v_Str := Type_In || Split_In;

  If Typeall_In = 1 Then
    Open c_List_All;
    Fetch c_List_All Bulk Collect
      Into v_病人id, v_主页id;
    Close c_List_All;
    For I In 1 .. v_病人id.Count Loop
      Out_Rec.C1 := v_病人id(I);
      Out_Rec.C2 := v_主页id(I);
      Pipe Row(Out_Rec);
    End Loop;
  Else
    Loop
      P := Instr(v_Str, Split_In);
      Exit When(Nvl(P, 0) = 0);
      n_Index := Trim(Substr(v_Str, 1, P - 1));
      If n_Index Is Not Null Then
        If n_Index = 0 Then
          Open c_List_Ry;
          Fetch c_List_Ry Bulk Collect
            Into v_病人id, v_主页id;
          Close c_List_Ry;
        Elsif n_Index = 1 Then
          Open c_List_Ss;
          Fetch c_List_Ss Bulk Collect
            Into v_病人id, v_主页id;
          Close c_List_Ss;
        Elsif n_Index = 2 Then
          Open c_List_Tw;
          Fetch c_List_Tw Bulk Collect
            Into v_病人id, v_主页id;
          Close c_List_Tw;
        Elsif n_Index = 3 Then
          Open c_List_Wz;
          Fetch c_List_Wz Bulk Collect
            Into v_病人id, v_主页id;
          Close c_List_Wz;
        Elsif n_Index = 4 Then
          Open c_List_Zr;
          Fetch c_List_Zr Bulk Collect
            Into v_病人id, v_主页id;
          Close c_List_Zr;
        Elsif n_Index = 5 Then
          Open c_List_Yj;
          Fetch c_List_Yj Bulk Collect
            Into v_病人id, v_主页id;
          Close c_List_Yj;
        Elsif n_Index = 6 Then
          Open c_List_Fm;
          Fetch c_List_Fm Bulk Collect
            Into v_病人id, v_主页id;
          Close c_List_Fm;
        End If;
      
        For I In 1 .. v_病人id.Count Loop
          Out_Rec.C1 := v_病人id(I);
          Out_Rec.C2 := v_主页id(I);
          Pipe Row(Out_Rec);
        End Loop;
      End If;
      v_Str := Substr(v_Str, P + 1);
    End Loop;
  End If;
  Return;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Getpatient;
/
--139097:胡俊勇,2019-04-23,门诊留观病人处理
Create Or Replace Procedure Zl_病人医嘱记录_回退
(
  医嘱id_In     In 病人医嘱记录.Id%Type,
  Flag_In       In Number := 0,
  医嘱内容_In   In 病人医嘱记录.医嘱内容%Type := Null,
  操作类型_In   In 病人医嘱状态.操作类型%Type := Null,
  操作员编号_In In 人员表.编号%Type := Null,
  操作员姓名_In In 人员表.姓名%Type := Null
  --功能：回退住院医嘱的状态操作或发送操作(回退重整操作通过调用Zl_病人医嘱记录_批量回退来进行) 
  --参数：医嘱ID_IN=一组医嘱ID 
  --      FLAG_IN=附加数据。回退停止：0=清除执行终止时间,1=保留现有的执行终止时间。 
  --      医嘱内容_IN=该过程被批量回退调用时才用，用于错误提示。 
  --      操作类型_IN=该过程被批量回退调用时才用，用于核对回退数据。0-回退发送,n=回退具体医嘱操作 
) Is
  --包含指定医嘱的操作记录,第一条为要回退的内容(状态操作优先) 
  --临嘱不回退发送后的自动停止,在回退发送时自动回退停止操作 
  Cursor c_Rolladvice Is
    Select b.操作人员, b.操作时间, 0 As 发送号, a.序号, Null As NO, b.操作类型, 0 As 执行状态, Sysdate + Null As 首次时间, Sysdate + Null As 末次时间,
           a.上次执行时间, a.医嘱期效, a.诊疗类别 As 类别, a.诊疗项目id, Null As 类型, a.病人id, a.主页id, a.婴儿, 0 As 记录性质, 0 As 门诊记帐, 0 As 开嘱科室id,
           a.审核标记, a.开嘱医生, a.执行科室id, Nvl(a.相关id, a.Id) As 组id, a.相关id, a.Id As 医嘱id, -null As 发送数次, Null As 样本条码
    From 病人医嘱记录 A, 病人医嘱状态 B
    Where a.Id = b.医嘱id And (a.Id = 医嘱id_In Or a.相关id = 医嘱id_In) And
          (Nvl(a.医嘱期效, 0) = 0 And b.操作类型 Not In (1, 2, 3) Or Nvl(a.医嘱期效, 0) = 1 And b.操作类型 Not In (1, 2, 3, 8))
    Union
    Select b.发送人 As 操作人员, b.发送时间 As 操作时间, b.发送号, a.序号, b.No, -null As 操作类型, b.执行状态, b.首次时间, b.末次时间, a.上次执行时间, a.医嘱期效,
           c.类别, a.诊疗项目id, c.操作类型 As 类型, a.病人id, a.主页id, a.婴儿, b.记录性质, b.门诊记帐, a.开嘱科室id, a.审核标记, a.开嘱医生, a.执行科室id,
           Nvl(a.相关id, a.Id) As 组id, a.相关id, a.Id As 医嘱id, b.发送数次, b.样本条码
    From 病人医嘱记录 A, 病人医嘱发送 B, 诊疗项目目录 C
    Where a.Id = b.医嘱id And a.诊疗项目id = c.Id And (a.Id = 医嘱id_In Or a.相关id = 医嘱id_In)
    Order By 操作时间 Desc, 发送号, 序号;
  r_Rolladvice c_Rolladvice%RowType;

  --方式同c_Rolladvice，只取发送部份用了自动回退处理 
  Cursor c_Rollsend(v_发送号 病人医嘱发送.发送号%Type) Is
    Select Distinct b.医嘱id, b.发送时间 As 操作时间, b.发送号, b.执行状态, a.诊疗类别 As 类别, c.当前病区id As 病人病区id, a.病人科室id,
                    b.执行部门id As 执行科室id
    From 病人医嘱记录 A, 病人医嘱发送 B, 病案主页 C
    Where a.Id = b.医嘱id And b.发送号 = v_发送号 And a.病人id = c.病人id And a.主页id = c.主页id And
          (a.Id = 医嘱id_In Or a.相关id = 医嘱id_In)
    Order By b.发送时间 Desc, b.发送号;

  --根据医嘱及发送NO求出本次回退要销帐的费用记录 
  --一组医嘱并不是都填写了发送记录,且可能NO不同(药品有,用法煎法不一定有) 
  --不管发送记录的计费状态(可能无需计费),有费用记录自然关联出来 
  --费用只求价格父号为空的,以便取序号销帐 
  --只管记录状态为1的费用,对于已销帐或部份销帐的记录,不再处理；其中"记录状态=3"的读取出来仅用于判断，不处理。 
  Cursor c_Rollmoneyout
  (
    v_发送号    病人医嘱发送.发送号%Type,
    v_医嘱id    病人医嘱记录.Id%Type,
    t_Adviceids t_Numlist
  ) Is
    Select /*+ Rule*/
     a.Id, a.记录状态, a.No, a.序号, a.收费类别, a.执行状态, d.跟踪在用, a.执行部门id, a.记录性质
    From 门诊费用记录 A, Table(t_Adviceids) B, 病人医嘱发送 C, 材料特性 D
    Where c.医嘱id = b.Column_Value And c.发送号 = v_发送号 And a.医嘱序号 = b.Column_Value And
          (a.医嘱序号 = v_医嘱id Or Nvl(v_医嘱id, 0) = 0) And a.记录状态 In (0, 1, 3) And a.No = c.No And a.记录性质 = c.记录性质 And
          a.价格父号 Is Null And a.收费细目id = d.材料id(+)
    Order By a.No, a.序号;

  Cursor c_Rollmoneyin
  (
    v_发送号    病人医嘱发送.发送号%Type,
    v_医嘱id    病人医嘱记录.Id%Type,
    t_Adviceids t_Numlist
  ) Is
    Select /*+ Rule*/
     a.Id, a.记录状态, a.No, a.序号, a.收费类别, a.执行状态, d.跟踪在用, a.执行部门id, a.记录性质
    From 住院费用记录 A, Table(t_Adviceids) B, 病人医嘱发送 C, 材料特性 D
    Where c.医嘱id = b.Column_Value And c.发送号 = v_发送号 And a.医嘱序号 = b.Column_Value And
          (a.医嘱序号 = v_医嘱id Or Nvl(v_医嘱id, 0) = 0) And a.记录状态 In (0, 1, 3) And a.No = c.No And a.记录性质 = c.记录性质 And
          a.价格父号 Is Null And a.收费细目id = d.材料id(+)
    Order By a.No, a.序号;

  --取发送住院记帐时自动发放的卫材(还没有退料的) 
  Cursor c_Stuff_Drug(v_费用id 药品收发记录.费用id%Type) Is
    Select ID
    From 药品收发记录
    Where 费用id = v_费用id And (记录状态 = 1 Or Mod(记录状态, 3) = 0) And 审核人 Is Not Null
    Order By 药品id;

  --用于处理特殊医嘱的回退 
  Cursor c_Patilog
  (
    v_病人id 病人变动记录.病人id%Type,
    v_主页id 病人变动记录.主页id%Type
  ) Is
    Select *
    From 病人变动记录
    Where 病人id = v_病人id And 主页id = v_主页id And 终止时间 Is Null
    Order By 开始时间 Desc;
  r_Patilog c_Patilog%RowType;

  Cursor c_Adviceids Is
    Select ID From 病人医嘱记录 Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
  t_Adviceids t_Numlist;

  v_医嘱状态     病人医嘱记录.医嘱状态%Type;
  v_医嘱期效     病人医嘱记录.医嘱期效%Type;
  v_费用no       病人医嘱发送.No%Type;
  v_费用序号     Varchar2(255);
  v_末次时间     病人医嘱发送.末次时间%Type;
  v_重整时间     病人医嘱状态.操作时间%Type;
  v_操作类型     诊疗项目目录.操作类型%Type;
  v_执行频率     诊疗项目目录.执行频率%Type;
  v_上次时间     病人医嘱记录.上次执行时间%Type;
  v_执行时间     病人医嘱记录.执行时间方案%Type;
  v_开始执行时间 病人医嘱记录.开始执行时间%Type;
  v_上次打印时间 病人医嘱记录.上次打印时间%Type;
  v_频率间隔     病人医嘱记录.频率间隔%Type;
  v_间隔单位     病人医嘱记录.间隔单位%Type;
  v_发送号       病人医嘱发送.发送号%Type;
  n_护理等级id   病人变动记录.护理等级id%Type;
  d_开始时间     病人变动记录.开始时间%Type;
  d_操作时间     病人医嘱状态.操作时间%Type;
  v_Tmp发送号    病人医嘱发送.发送号%Type;
  n_执行         Number;

  Intdigit   Number(3);
  v_Update   Number(1);
  v_Count    Number(5);
  v_Temp     Varchar2(2000);
  v_人员编号 人员表.编号%Type;
  v_人员姓名 人员表.姓名%Type;
  v_Time     Varchar2(4000);
  n_Blndo    Number;

  v_Error Varchar2(2000);
  Err_Custom Exception;

  Function Checkmoneyundo
  (
    v_No       住院费用记录.No%Type,
    v_记录性质 住院费用记录.记录性质%Type,
    v_序号     住院费用记录.序号%Type,
    n_场合     Number := 0 --0住院，1门诊 
  ) Return Number Is
    n_Num      Number;
    n_执行状态 Number;
  Begin
    n_Num := 0;
    If n_场合 = 0 Then
      Select Nvl(Sum(Nvl(付数, 1) * 数次), 0) As 数量
      Into n_Num
      From 住院费用记录
      Where NO = v_No And 记录性质 = v_记录性质 And 序号 = v_序号 And 记录状态 In (2, 3);
      Select Nvl(执行状态, 0)
      Into n_执行状态
      From 住院费用记录
      Where NO = v_No And 记录性质 = v_记录性质 And 序号 = v_序号 And 记录状态 = 3;
    Else
      Select Nvl(Sum(Nvl(付数, 1) * 数次), 0) As 数量
      Into n_Num
      From 门诊费用记录
      Where NO = v_No And 记录性质 = v_记录性质 And 序号 = v_序号 And 记录状态 In (2, 3);
      Select Nvl(执行状态, 0)
      Into n_执行状态
      From 门诊费用记录
      Where NO = v_No And 记录性质 = v_记录性质 And 序号 = v_序号 And 记录状态 = 3;
    End If;
    If n_Num <> 0 Then
      n_Num := 1;
    End If;
    --如果主记录是已执行（部分执行的）则不自动退。 
    If n_执行状态 <> 0 Then
      n_Num := 0;
    End If;
    Return(n_Num);
  End;
Begin
  v_Tmp发送号 := -1;
  Open c_Rolladvice;
  Loop
    Fetch c_Rolladvice
      Into r_Rolladvice;
    If c_Rolladvice%RowCount = 0 Then
      Close c_Rolladvice;
      v_Error := Nvl(医嘱内容_In, '该医嘱') || '当前没有可以回退的内容。';
      Raise Err_Custom;
    End If;
    Exit When c_Rolladvice%NotFound;
    Exit When d_操作时间 <> r_Rolladvice.操作时间 And d_操作时间 Is Not Null;
    d_操作时间 := r_Rolladvice.操作时间;
  
    --批量回退调用时判断 
    If 医嘱内容_In Is Not Null Then
      If Nvl(r_Rolladvice.操作类型, 0) <> Nvl(操作类型_In, 0) Then
        v_Error := Nvl(医嘱内容_In, '该医嘱') || '不能与当前医嘱一起回退，可能该医嘱已经执行了其他操作。';
        Raise Err_Custom;
      End If;
    End If;
  
    --一组发送号只执行一次 
    If v_Tmp发送号 <> r_Rolladvice.发送号 Then
      v_Tmp发送号 := r_Rolladvice.发送号;
      n_执行      := 1;
    Else
      n_执行 := 0;
    End If;
  
    If n_执行 = 1 Then
      Open c_Adviceids;
      Fetch c_Adviceids Bulk Collect
        Into t_Adviceids;
      Close c_Adviceids;
    
      If r_Rolladvice.发送号 = 0 Then
        --回退医嘱状态操作(以时间关键字) 
        --4-作废；5-重整；6-暂停；7-启用；8-停止；9-确认停止；10-皮试结果;13-停嘱申请 
        ------------------------------------------------------------------ 
        --最多只能退回到校对状态 
        If r_Rolladvice.操作类型 = 3 Then
          v_Error := Nvl(医嘱内容_In, '该医嘱') || '当前处于通过校对状态，不能再回退。';
          Raise Err_Custom;
        Elsif r_Rolladvice.操作类型 = 4 And Nvl(r_Rolladvice.婴儿, 0) = 0 Then
          If r_Rolladvice.类别 = 'H' Then
            Select 操作类型, 执行频率 Into v_操作类型, v_执行频率 From 诊疗项目目录 Where ID = r_Rolladvice.诊疗项目id;
            If v_操作类型 = '1' And v_执行频率 = '2' Then
              v_Error := '护理等级作废后不能再回退。';
              Raise Err_Custom;
            End If;
          End If;
        End If;
      
        --检查是否回退最近次重整之前的操作 
        If r_Rolladvice.操作类型 <> 5 Then
          --取最后重整时间 
          Select Nvl(医嘱重整时间, To_Date('1900-01-01', 'YYYY-MM-DD'))
          Into v_重整时间
          From 病案主页
          Where 病人id = r_Rolladvice.病人id And 主页id = r_Rolladvice.主页id;
        
          If r_Rolladvice.操作时间 < v_重整时间 Then
            v_Error := '该病人最近次重整之前的操作不能再回退。';
            Raise Err_Custom;
          End If;
        End If;
      
        --删除(该组医嘱)最近的状态操作记录 
        Delete /*+ Rule*/
        From 病人医嘱状态
        Where 医嘱id In (Select Column_Value From Table(t_Adviceids)) And 操作时间 = r_Rolladvice.操作时间;
      
        --取删除后应恢复的医嘱状态 
        Select 操作类型
        Into v_医嘱状态
        From 病人医嘱状态
        Where 操作时间 = (Select Max(操作时间) From 病人医嘱状态 Where 医嘱id = 医嘱id_In) And 医嘱id = 医嘱id_In;
      
        --恢复(该组医嘱)回退后的状态 
        Update 病人医嘱记录 Set 医嘱状态 = v_医嘱状态 Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
      
        --其它额外的处理 
        If r_Rolladvice.操作类型 = 8 Then
          --被超期发送收回过的医嘱 ，如果是销帐申请模式，则判断对应的“病人费用销帐”申请是否取消，是则允许回退，否则不允许， 
          --                       如果是产生负数费用模式，则不允许再回退。 
          --可能超期发送收回时被全部收回(无上次执行时间) 
          Select /*+ Rule*/
           Nvl(Count(*), 0)
          Into v_Count
          From 病人医嘱记录 A, 病人医嘱发送 B
          Where b.医嘱id = a.Id And (a.Id = 医嘱id_In Or a.相关id = 医嘱id_In) And
                b.发送号 =
                (Select Max(发送号) From 病人医嘱发送 Where 医嘱id In (Select Column_Value From Table(t_Adviceids))) And
                a.执行终止时间 Is Not Null And ((a.上次执行时间 < b.末次时间) Or (a.上次执行时间 Is Null And b.末次时间 Is Not Null));
          If v_Count > 0 Then
            If zl_GetSysParameter('超期收回产生负数费用', 1254) = '1' Then
              v_Error := Nvl(医嘱内容_In, '该医嘱') || '已被超期发送收回，不能再撤消停止操作。';
              Raise Err_Custom;
            Else
              --如果已经取消销帐申请，则允许回退. 
              Select Count(1)
              Into v_Count
              From 病人费用销帐 A, 住院费用记录 B, 病人医嘱记录 C
              Where a.费用id = b.Id And c.Id = b.医嘱序号 And (c.Id = 医嘱id_In Or c.相关id = 医嘱id_In);
              If v_Count > 0 Then
                v_Error := Nvl(医嘱内容_In, '该医嘱') || '已被超期发送收回，不能再撤消停止操作。';
                Raise Err_Custom;
              Else
                --得到上次执行时间等信息 
                Select 上次执行时间, 执行时间方案, 开始执行时间, 上次打印时间, 频率间隔, 间隔单位
                Into v_上次时间, v_执行时间, v_开始执行时间, v_上次打印时间, v_频率间隔, v_间隔单位
                From 病人医嘱记录
                Where ID = 医嘱id_In;
                v_上次时间 := To_Date(To_Char(v_上次时间 + 1 / 24 / 60 / 60, 'yyyy-MM-dd hh24:mi:ss'), 'yyyy-MM-dd hh24:mi:ss');
              
                --修改上次执行时间为收回后的末次执行时间。 
                v_末次时间 := Null;
                Begin
                  --一组医嘱的发送首末时间相同,一并给药是取最小的 
                  --取相关ID为NULL的医嘱的发送记录的时间 
                  --但给药途径或中药用法可能未填写发送记录 
                  Select /*+ Rule*/
                   末次时间, 发送号
                  Into v_末次时间, v_发送号
                  From 病人医嘱发送
                  Where 医嘱id In (Select Column_Value From Table(t_Adviceids)) And
                        发送号 = (Select Max(发送号)
                               From 病人医嘱发送
                               Where 医嘱id In (Select Column_Value From Table(t_Adviceids))) And Rownum = 1;
                Exception
                  When Others Then
                    Null;
                End;
                Update 病人医嘱记录 Set 上次执行时间 = v_末次时间 Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
              
                Select Count(1) Into v_Count From 病人医嘱发送 Where 医嘱id = 医嘱id_In And 发送号 = v_发送号;
                If v_Count > 0 Then
                  --还原医嘱执行时间 
                  Select Zl_Adviceexetimes(医嘱id_In, v_上次时间, v_末次时间, v_执行时间, v_开始执行时间, v_上次打印时间, v_频率间隔, v_间隔单位, 0)
                  Into v_Time
                  From Dual;
                  Insert Into 医嘱执行时间
                    (要求时间, 医嘱id, 发送号)
                    Select To_Date(Column_Value, 'yyyy-mm-dd hh24:mi:ss'), 医嘱id_In, v_发送号
                    From Table(f_Str2list(v_Time));
                End If;
              End If;
            End If;
          End If;
        
          --护理等级变动，后续有其他变动时，不允许回退 
          If r_Rolladvice.类别 = 'H' And Nvl(r_Rolladvice.婴儿, 0) = 0 Then
            Select 操作类型, 执行频率 Into v_操作类型, v_执行频率 From 诊疗项目目录 Where ID = r_Rolladvice.诊疗项目id;
            If v_操作类型 = '1' And v_执行频率 = '2' Then
              Select Count(*), Max(a.护理等级id), Max(a.开始时间)
              Into v_Count, n_护理等级id, d_开始时间
              From 病人变动记录 A
              Where a.病人id = r_Rolladvice.病人id And a.主页id = r_Rolladvice.主页id And a.开始原因 = 6 And a.终止时间 Is Null And
                    a.附加床位 = 0;
              --如果没有找到最后一条是护理等级变动则禁止 
              If v_Count = 0 Then
                --医嘱护理等级和入住时候的护理等级一致时要单独判断 
                Select Count(*)
                Into v_Count
                From 病人变动记录 A
                Where a.病人id = r_Rolladvice.病人id And a.主页id = r_Rolladvice.主页id And a.开始原因 = 6;
                If v_Count > 0 Then
                  v_Error := '由于护理等级医嘱停止后该病人已经产生了其他变动记录,不能回退该医嘱的停止操作。';
                  Raise Err_Custom;
                End If;
              Else
                --如果n_护理等级ID为Null，则检查是否是当前回退的医嘱对应的变动记录,目的是有多个护理等级医嘱时要求按顺序回退。 
                --如果n_护理等级ID不为Null，则有可能是校对下一条护理等级时，自动停止的，未产生变动记录， 
                --     则需要检查当前最后一条变动的护理等级ID是否是当前医嘱的护理等级ID,目的是有多个护理等级医嘱时要求按顺序回退，如果是则不需要再撤销最后一次变动，直接回退医嘱即可。 
                If n_护理等级id Is Null Then
                  Select Count(*)
                  Into v_Count
                  From 病人变动记录 B, 病人医嘱计价 C
                  Where b.病人id = r_Rolladvice.病人id And b.主页id = r_Rolladvice.主页id And c.医嘱id = 医嘱id_In And
                        c.收费细目id = b.护理等级id And b.终止时间 = d_开始时间 And b.终止原因 = 6 And b.附加床位 = 0;
                Else
                  --开始时间只取分钟对比，校对的时候护理等级的开始时间是医嘱开始时间+当前时间的秒钟 
                  Select Count(*)
                  Into v_Count
                  From 病人医嘱计价 C, 病人医嘱记录 A
                  Where a.Id = c.医嘱id And a.Id = 医嘱id_In And c.收费细目id = n_护理等级id And
                        a.开始执行时间 = To_Date(To_Char(d_开始时间, 'yyyy-mm-dd hh24:mi'), 'yyyy-mm-dd hh24:mi');
                End If;
                If v_Count = 0 Then
                  v_Error := '您回退的医嘱不是最后一条护理等级医嘱，请将后面的护理等级医嘱作废后再回退本条医嘱。';
                  Raise Err_Custom;
                End If;
              
                If n_护理等级id Is Null Then
                  --当前操作人员 
                  If 操作员编号_In Is Not Null And 操作员姓名_In Is Not Null Then
                    v_人员编号 := 操作员编号_In;
                    v_人员姓名 := 操作员姓名_In;
                  Else
                    v_Temp     := Zl_Identity;
                    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
                    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
                    v_人员编号 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
                    v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
                  End If;
                
                  Zl_病人变动记录_Undo(r_Rolladvice.病人id, r_Rolladvice.主页id, v_人员编号, v_人员姓名, '1', Null, Null, '护理等级变动');
                End If;
              End If;
            End If;
          End If;
        
          If r_Rolladvice.类别 = 'Z' And Instr(',9,10,', ',' || r_Rolladvice.类型 || ',') > 0 And
             Nvl(r_Rolladvice.婴儿, 0) = 0 Then
            --回退病况医嘱时，调用变动记录回退 
            Zl_病人变动记录_Undo(r_Rolladvice.病人id, r_Rolladvice.主页id, v_人员编号, v_人员姓名, Null, Null, Null, '病况变动');
          End If;
        
          --回退医嘱停止时,清空停嘱医生和时间,如果是实习医师申请后审核的，则恢复待审核状态 
          Update 病人医嘱记录
          Set 执行终止时间 = Decode(Flag_In, 1, 执行终止时间, Null), 停嘱医生 = Null, 停嘱时间 = Null,
              审核标记 = Decode(r_Rolladvice.审核标记, 3, 2, r_Rolladvice.审核标记)
          Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
        Elsif r_Rolladvice.操作类型 = 9 Then
          --回退医嘱确认停止时,检查是否已打印停嘱时间 
          Select /*+ Rule*/
           Count(*)
          Into v_Count
          From 病人医嘱打印
          Where 打印标记 = 1 And 医嘱id In (Select Column_Value From Table(t_Adviceids));
          If v_Count > 0 Then
            v_Error := Nvl(医嘱内容_In, '该医嘱') || '的停嘱时间已经打印，不能再撤消确认停止操作。';
            Raise Err_Custom;
          End If;
        
          --回退医嘱确认停止时,清空停嘱医生和时间 
          Update 病人医嘱记录 Set 确认停嘱时间 = Null, 确认停嘱护士 = Null Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
        Elsif r_Rolladvice.操作类型 = 10 Then
          --回退标注皮试结果,同时删除过敏登记(+)或(-),根据记录时间 
          --过敏的记录与医嘱操作无观，不需要处理 
          Delete From 病人过敏记录
          Where 病人id = r_Rolladvice.病人id And Nvl(主页id, 0) = Nvl(r_Rolladvice.主页id, 0) And 记录时间 = r_Rolladvice.操作时间 And
                Nvl(结果, 0) = 0;
        
          Update 病人医嘱记录 Set 皮试结果 = Null Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
        Elsif r_Rolladvice.操作类型 = 13 Then
          If Instr(r_Rolladvice.开嘱医生, '/') > 0 Then
            Update 病人医嘱记录 Set 审核标记 = 1 Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
          Else
            Update 病人医嘱记录 Set 审核标记 = Null Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
          End If;
        End If;
        --回退术后医嘱作废操作 
        --术后医嘱 
        If r_Rolladvice.操作类型 = 4 And r_Rolladvice.类别 = 'Z' Then
          Select Count(1) Into v_Count From 诊疗项目目录 Where ID = r_Rolladvice.诊疗项目id And 操作类型 = '4';
          If v_Count = 1 Then
            b_Message.Zlhis_Cis_004(r_Rolladvice.病人id, r_Rolladvice.主页id, 医嘱id_In);
          End If;
        End If;
      Else
        --回退医嘱发送(以发送号关键字) 
        ------------------------------------------------------------------ 
        --当前操作人员 
        v_Temp     := Zl_Identity;
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
        v_人员编号 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
        v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
      
        --检查是否是输液配液记录，并是否已经锁定，如果查询有数据说明是配液记录 
        Begin
          Select Decode(Max(是否锁定), 1, 1, 0)
          Into v_Count
          From 输液配药记录
          Where 医嘱id = 医嘱id_In And 发送号 = r_Rolladvice.发送号;
        Exception
          When Others Then
            v_Count := -1;
        End;
      
        If v_Count = 1 Then
          v_Error := '医嘱"' || 医嘱内容_In || '"是输液药品，已经被输液配置中心锁定，不能回退发送。';
          Raise Err_Custom;
        Elsif v_Count = 0 Then
          Zl_输液配药记录_医嘱回退(医嘱id_In, r_Rolladvice.发送号, v_人员姓名, Sysdate);
        End If;
      
        --检查是否存在未审核的销帐申请 
        Select Count(*)
        Into v_Count
        From 病人医嘱记录 A, 病人医嘱发送 B, 住院费用记录 C, 病人费用销帐 D
        Where (a.Id = 医嘱id_In Or a.相关id = 医嘱id_In) And a.Id = b.医嘱id And b.医嘱id = c.医嘱序号 And c.Id = d.费用id And
              c.记录状态 In (0, 1, 3) And d.状态 = 0;
      
        If v_Count > 0 Then
          v_Error := '医嘱"' || 医嘱内容_In || '"存在未审核的销帐申请，请取消或审核销帐申请后再回退发送。';
          Raise Err_Custom;
        End If;
      
        --检查医嘱是否存在有效的医嘱附费 
        Select Count(*)
        Into v_Count
        From 病人医嘱附费 A, 住院费用记录 B
        Where a.医嘱id = b.医嘱序号 And a.No = b.No And b.记录状态 = 1 And b.实收金额 <> 0 And a.发送号 = r_Rolladvice.发送号 And
              a.医嘱id In (Select Column_Value From Table(t_Adviceids));
        If v_Count > 0 Then
          v_Error := '该医嘱下还存在附费项目，请先冲销。';
          Raise Err_Custom;
        End If;
      
        --本科发送自动执行时，回退也自动回退执行(仅护士站有此功能) 
        --非跟踪在用的卫材医嘱，同普通医嘱执行处理 
        Select 医嘱期效 Into v_医嘱期效 From 病人医嘱记录 Where ID = 医嘱id_In;
        If Substr(zl_GetSysParameter('本科执行自动完成', 1254), v_医嘱期效 + 1, 1) = '1' Then
          Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into Intdigit From Dual;
        
          For r_Rollsend In c_Rollsend(r_Rolladvice.发送号) Loop
            If Nvl(r_Rollsend.执行状态, 0) = 1 And
               (Nvl(r_Rollsend.执行科室id, 0) = Nvl(r_Rollsend.病人病区id, 0) Or
                Nvl(r_Rollsend.执行科室id, 0) = Nvl(r_Rollsend.病人科室id, 0)) Then
            
              --医嘱的执行状态 
              Update 病人医嘱发送 Set 执行状态 = 0 Where 发送号 = r_Rollsend.发送号 And 医嘱id = r_Rollsend.医嘱id;
              v_Update := 1;
            
              If r_Rolladvice.记录性质 = 2 And Nvl(r_Rolladvice.门诊记帐, 0) = 0 Then
                --费用的执行状态 
                For r_Rollmoney In c_Rollmoneyin(r_Rollsend.发送号, r_Rollsend.医嘱id, t_Adviceids) Loop
                  n_Blndo := 0;
                  If r_Rollmoney.记录状态 <> 3 Then
                    n_Blndo := 1;
                  Else
                    n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.记录性质, r_Rollmoney.序号);
                  End If;
                  If n_Blndo > 0 Then
                    If Not (r_Rollmoney.收费类别 = '4' And Nvl(r_Rollmoney.跟踪在用, 0) = 1) And
                       Not r_Rollmoney.收费类别 In ('5', '6', '7') Then
                      --普通费用直接取消执行状态，不含药品和跟踪在用的卫材 
                      Update 住院费用记录
                      Set 执行状态 = 0, 执行时间 = Null, 执行人 = Null
                      Where NO = r_Rollmoney.No And 记录性质 = r_Rolladvice.记录性质 And 记录状态 = r_Rollmoney.记录状态 And
                            Nvl(价格父号, 序号) = r_Rollmoney.序号 And 医嘱序号 = r_Rollsend.医嘱id;
                    Elsif r_Rollmoney.收费类别 = '4' And Nvl(r_Rollmoney.跟踪在用, 0) = 1 Then
                      --跟踪在用的卫材，才自动退料 
                      For r_Stuff In c_Stuff_Drug(r_Rollmoney.Id) Loop
                        Zl_材料收发记录_部门退料(r_Stuff.Id, v_人员姓名, Sysdate, Null, Null, Null, Null, 0, v_人员姓名);
                      End Loop;
                    Elsif r_Rollmoney.收费类别 In ('5', '6', '7') Then
                      --住院科室发药的药品自动退药 
                      If r_Rollmoney.执行部门id = r_Rollsend.病人病区id Or r_Rollmoney.执行部门id = r_Rollsend.病人科室id Then
                        For r_Drug In c_Stuff_Drug(r_Rollmoney.Id) Loop
                          Zl_药品收发记录_部门退药(r_Drug.Id, v_人员姓名, Sysdate, Null, Null, Null, Null, Null, Null, Intdigit, 2);
                        End Loop;
                      End If;
                    End If;
                  End If;
                End Loop;
              Else
                --住院病人费用发送到门诊的情况，病人来源都是住院的 
                For r_Rollmoney In c_Rollmoneyout(r_Rollsend.发送号, r_Rollsend.医嘱id, t_Adviceids) Loop
                  n_Blndo := 0;
                  If r_Rollmoney.记录状态 <> 3 Then
                    n_Blndo := 1;
                  Else
                    n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.记录性质, r_Rollmoney.序号, 1);
                  End If;
                  If n_Blndo > 0 Then
                    If Not (r_Rollmoney.收费类别 = '4' And Nvl(r_Rollmoney.跟踪在用, 0) = 1) And
                       Not r_Rollmoney.收费类别 In ('5', '6', '7') Then
                      --普通费用直接取消执行状态，不含药品和跟踪在用的卫材 
                      Update 门诊费用记录
                      Set 执行状态 = 0, 执行时间 = Null, 执行人 = Null
                      Where NO = r_Rollmoney.No And 记录性质 = r_Rolladvice.记录性质 And 记录状态 = r_Rollmoney.记录状态 And
                            Nvl(价格父号, 序号) = r_Rollmoney.序号 And 医嘱序号 = r_Rollsend.医嘱id;
                    Elsif r_Rollmoney.收费类别 = '4' And Nvl(r_Rollmoney.跟踪在用, 0) = 1 Then
                      --跟踪在用的卫材，才自动退料 
                      For r_Stuff In c_Stuff_Drug(r_Rollmoney.Id) Loop
                        Zl_材料收发记录_部门退料(r_Stuff.Id, v_人员姓名, Sysdate, Null, Null, Null, Null, 0, v_人员姓名);
                      End Loop;
                    Elsif r_Rollmoney.收费类别 In ('5', '6', '7') Then
                      --本科室发药的药品自动退药 
                      If r_Rollmoney.执行部门id = r_Rollsend.病人病区id Or r_Rollmoney.执行部门id = r_Rollsend.病人科室id Then
                        For r_Drug In c_Stuff_Drug(r_Rollmoney.Id) Loop
                          Zl_药品收发记录_部门退药(r_Drug.Id, v_人员姓名, Sysdate, Null, Null, Null, Null, Null, Null, Intdigit, 1);
                        End Loop;
                      End If;
                    End If;
                  End If;
                End Loop;
              End If;
            End If;
          End Loop;
        End If;
        ------------------------------------------------------------------ 
        --被超期收回的长期药品医嘱不允许回退(再退费用就多退了) 
        If Nvl(r_Rolladvice.医嘱期效, 0) = 0 Then
          If r_Rolladvice.上次执行时间 Is Not Null And r_Rolladvice.末次时间 Is Not Null Then
            If r_Rolladvice.上次执行时间 < r_Rolladvice.末次时间 Then
              v_Error := Nvl(医嘱内容_In, '该医嘱') || '最近超期发送的内容已被收回，不能再回退。';
              Raise Err_Custom;
            End If;
          Elsif r_Rolladvice.上次执行时间 Is Null And r_Rolladvice.末次时间 Is Not Null Then
            --长嘱可能被全部超期收回 
            v_Error := Nvl(医嘱内容_In, '该医嘱') || '未被发送，或发送的内容已被全部超期收回，不能再回退。';
            Raise Err_Custom;
          End If;
        End If;
      
        If Nvl(r_Rolladvice.执行状态, 0) In (1, 3) And v_Update <> 1 Then
          --1-完全执行;3-正在执行 
          v_Error := Nvl(医嘱内容_In, '该医嘱') || '最近发送的内容已经执行或正在执行，不能回退。';
          Raise Err_Custom;
        Else
          --如果相关医嘱已执行，则也要限制回退（例如：检验的采集方式） 
          Select /*+ Rule*/
           Count(1)
          Into v_Count
          From 病人医嘱发送
          Where 医嘱id In (Select Column_Value From Table(t_Adviceids)) And 执行状态 In (1, 3) And
                发送号 =
                (Select Max(发送号) From 病人医嘱发送 Where 医嘱id In (Select Column_Value From Table(t_Adviceids)));
          If v_Count > 0 Then
            v_Error := Nvl(医嘱内容_In, '该医嘱') || '最近发送的内容已经执行或正在执行，不能回退。';
            Raise Err_Custom;
          End If;
        End If;
      
        ------------------------------------------------------------------ 
        --将该组医嘱的费用销帐(按一组医嘱可能有不同NO处理) 
        --如果原始费用已被销帐(或部分销帐),调用过程中有判断 
        v_费用no   := Null;
        v_费用序号 := Null;
        If r_Rolladvice.记录性质 = 2 And Nvl(r_Rolladvice.门诊记帐, 0) = 0 Then
          For r_Rollmoney In c_Rollmoneyin(r_Rolladvice.发送号, Null, t_Adviceids) Loop
            --对应的费用已执行 
            If Nvl(r_Rollmoney.执行状态, 0) <> 0 Then
              v_Error := Nvl(医嘱内容_In, '该医嘱') || '发送的费用单据"' || r_Rollmoney.No || '"中的内容已被部分或完全执行，不能回退。';
              Raise Err_Custom;
            End If;
            n_Blndo := 0;
            If r_Rollmoney.记录状态 <> 3 Then
              n_Blndo := 1;
            Else
              n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.记录性质, r_Rollmoney.序号);
            End If;
            If n_Blndo > 0 Then
              --这种仅用于判断部分退药 
              If v_费用no <> r_Rollmoney.No And v_费用序号 Is Not Null Then
                Zl_住院记帐记录_Delete(v_费用no, Substr(v_费用序号, 2), v_人员编号, v_人员姓名, 2, 0, 0);
                v_费用序号 := Null;
              End If;
              v_费用no   := r_Rollmoney.No;
              v_费用序号 := v_费用序号 || ',' || r_Rollmoney.序号;
            End If;
          End Loop;
        Else
          For r_Rollmoney In c_Rollmoneyout(r_Rolladvice.发送号, Null, t_Adviceids) Loop
            --对应的费用已执行 
            If Nvl(r_Rollmoney.执行状态, 0) <> 0 And Not (Nvl(r_Rollmoney.执行状态, 0) = -1 And Nvl(r_Rollmoney.记录状态, 0) = 0) Then
              v_Error := Nvl(医嘱内容_In, '该医嘱') || '发送的费用单据"' || r_Rollmoney.No || '"中的内容已被部分或完全执行，不能回退。';
              Raise Err_Custom;
            End If;
            --收费单据已收费 
            If r_Rollmoney.记录状态 = 1 And Not (r_Rolladvice.记录性质 = 2 And Nvl(r_Rolladvice.门诊记帐, 0) = 1) Then
              v_Error := Nvl(医嘱内容_In, '该医嘱') || '发送的门诊单据"' || r_Rollmoney.No || '"已收费，不能回退。';
              Raise Err_Custom;
            End If;
            n_Blndo := 0;
            If r_Rollmoney.记录状态 <> 3 Then
              n_Blndo := 1;
            Else
              n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.记录性质, r_Rollmoney.序号, 1);
            End If;
            If n_Blndo > 0 Then
              --这种仅用于判断部分退药 
              If v_费用no <> r_Rollmoney.No And v_费用序号 Is Not Null Then
                If r_Rolladvice.记录性质 = 2 And Nvl(r_Rolladvice.门诊记帐, 0) = 1 Then
                  --住院发送为门诊记帐(如果是门诊医生发送为门诊记帐，门诊医嘱没有回退功能) 
                  Zl_门诊记帐记录_Delete(v_费用no, v_费用序号, v_人员编号, v_人员姓名, 0);
                Else
                  Zl_门诊划价记录_Delete(v_费用no, Substr(v_费用序号, 2));
                End If;
                v_费用序号 := Null;
              End If;
              v_费用no   := r_Rollmoney.No;
              v_费用序号 := v_费用序号 || ',' || r_Rollmoney.序号;
            End If;
          End Loop;
        End If;
        If v_费用序号 Is Not Null And v_费用no Is Not Null Then
          v_费用序号 := Substr(v_费用序号, 2);
          If r_Rolladvice.记录性质 = 2 And Nvl(r_Rolladvice.门诊记帐, 0) = 0 Then
            Zl_住院记帐记录_Delete(v_费用no, v_费用序号, v_人员编号, v_人员姓名, 2, 0, 0);
          Elsif r_Rolladvice.记录性质 = 2 And Nvl(r_Rolladvice.门诊记帐, 0) = 1 Then
            --住院发送为门诊记帐(如果是门诊医生发送为门诊记帐，门诊医嘱没有回退功能) 
            Zl_门诊记帐记录_Delete(v_费用no, v_费用序号, v_人员编号, v_人员姓名, 0);
          Else
            Zl_门诊划价记录_Delete(v_费用no, v_费用序号);
          End If;
        End If;
      
        --回退发送操作，通过发送记录产生消息，对于组合类医嘱要提前      
        For R In (Select a.病人id, a.主页id, b.No, b.发送号, b.发送数次, b.首次时间, b.末次时间, b.样本条码, a.Id, a.相关id,
                         Nvl(a.相关id, a.Id) As 组id, c.类别, c.操作类型, a.执行科室id
                  From 病人医嘱记录 A, 病人医嘱发送 B, 诊疗项目目录 C
                  Where a.Id = b.医嘱id And a.诊疗项目id = c.Id And b.发送号 = r_Rolladvice.发送号 And
                        b.医嘱id In (Select Column_Value From Table(t_Adviceids))
                  Order By a.序号) Loop
        
          --此处添加消息触发
          If r.类别 = 'D' And r.相关id Is Null Then
            --检查 
            b_Message.Zlhis_Cis_037(r.病人id, r.主页id, Null, r.发送号, r.组id, r.No, 2);
          Elsif r.类别 = 'F' And r.相关id Is Null Then
            --手术 
            b_Message.Zlhis_Cis_038(r.病人id, r.主页id, Null, r.发送号, r.组id, r.No);
          Elsif r.类别 = 'K' And r.相关id Is Null Then
            --输血 
            b_Message.Zlhis_Cis_039(r.病人id, r.主页id, Null, r.发送号, r.组id, r.No);
          Elsif r.类别 = 'E' And r.操作类型 = '6' Then
            --检验
            b_Message.Zlhis_Cis_036(r.病人id, r.主页id, Null, r.发送号, r.组id, r.No, 2);
          End If;
        
          Select Count(1) Into v_Count From 部门性质说明 A Where a.部门id = r.执行科室id And a.工作性质 = '护理';
          If v_Count > 0 Then
            --病区执行医嘱回退发送
            b_Message.Zlhis_Cis_044(r.病人id, r.主页id, r.发送号, r.Id, r.No, r.发送数次, r.首次时间, r.末次时间, r.样本条码);
          End If;
        End Loop;
      
        --输血医嘱先删除病人医嘱附费 
        Delete From 病人医嘱附费 Where 发送号 = r_Rolladvice.发送号 And 医嘱id = 医嘱id_In;
      
        --删除医嘱执行时间 (仅主医嘱ID才产生了记录) 
        Delete From 医嘱执行时间 Where 发送号 = r_Rolladvice.发送号 And 医嘱id = 医嘱id_In;
      
        --删除发送记录(该组医嘱的) 
        Delete /*+ Rule*/
        From 病人医嘱发送
        Where 发送号 = r_Rolladvice.发送号 And 医嘱id In (Select Column_Value From Table(t_Adviceids));
      
        --标记(该组医嘱)上次执行时间(以上次发送的末次执行时间) 
        --所有长嘱(包括持续性长嘱)发送时都填写了末次时间 
        --临嘱可能没有，且只可能发送了一次。 
        v_末次时间 := Null;
        Begin
          --一组医嘱的发送首末时间相同,一并给药是取最小的 
          --取相关ID为NULL的医嘱的发送记录的时间 
          --但给药途径或中药用法可能未填写发送记录 
          Select /*+ Rule*/
           末次时间
          Into v_末次时间
          From 病人医嘱发送
          Where 医嘱id In (Select Column_Value From Table(t_Adviceids)) And
                发送号 =
                (Select Max(发送号) From 病人医嘱发送 Where 医嘱id In (Select Column_Value From Table(t_Adviceids))) And
                Rownum = 1;
        Exception
          When Others Then
            Null;
        End;
        Update 病人医嘱记录 Set 上次执行时间 = v_末次时间 Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
      
        --回退临嘱发送时，同时自动回退停止 
        If Nvl(r_Rolladvice.医嘱期效, 0) = 1 Then
          --删除(该组临嘱)最近的停止状态操作记录 
          Delete /*+ Rule*/
          From 病人医嘱状态
          Where 医嘱id In (Select Column_Value From Table(t_Adviceids)) And
                操作时间 = (Select Max(操作时间) From 病人医嘱状态 Where 医嘱id = 医嘱id_In) And 操作类型 = 8;
          --r_RollAdvice.操作时间:因发送时间可能不与自动停止时间相同。 
        
          --取删除后应恢复的医嘱状态 
          Select 操作类型
          Into v_医嘱状态
          From 病人医嘱状态
          Where 操作时间 = (Select Max(操作时间) From 病人医嘱状态 Where 医嘱id = 医嘱id_In) And 医嘱id = 医嘱id_In;
        
          --恢复(该组医嘱)回退后的状态 
          Update 病人医嘱记录
          Set 医嘱状态 = v_医嘱状态, 执行终止时间 = Null, 停嘱医生 = Null, 停嘱时间 = Null
          Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
        End If;
      
        --住院特殊医嘱发送后的回退(3-转科;5-出院;6-转院,11-死亡) 
        If r_Rolladvice.类别 = 'Z' And Instr(',3,5,6,11,', ',' || r_Rolladvice.类型 || ',') > 0 And
           Nvl(r_Rolladvice.婴儿, 0) = 0 Then
          Open c_Patilog(r_Rolladvice.病人id, r_Rolladvice.主页id);
          Fetch c_Patilog
            Into r_Patilog;
          If c_Patilog%Found Then
            If r_Rolladvice.类型 = '3' And r_Patilog.开始原因 = 3 Then
              --取消病人转科状态 
              If r_Patilog.开始时间 Is Null Then
                --转科医嘱的特殊处理，当一个病人有两条转科医嘱时，只能回退最近的一条,70443 
                Select Count(1)
                Into v_Count
                From 病人医嘱记录 A, 诊疗项目目录 B
                Where a.诊疗项目id = b.Id And a.病人id = r_Rolladvice.病人id And a.主页id = r_Rolladvice.主页id And a.诊疗类别 = 'Z' And
                      b.操作类型 = '3' And a.医嘱状态 = 8 And
                      a.开始执行时间 > (Select 开始执行时间 From 病人医嘱记录 Where ID = 医嘱id_In);
                If v_Count = 0 Then
                  Zl_病人变动记录_Undo(r_Rolladvice.病人id, r_Rolladvice.主页id, v_人员编号, v_人员姓名, Null, Null, Null, '转科');
                Else
                  v_Error := '病人转科已经入科，不能再回退。';
                  Raise Err_Custom;
                End If;
              Else
                v_Error := '病人转科已经入科，不能再回退。';
                Raise Err_Custom;
              End If;
            Elsif r_Rolladvice.类型 In ('5', '6', '11') And r_Patilog.开始原因 = 10 Then
              --取消病人预出院状态 
              Zl_病人变动记录_Undo(r_Rolladvice.病人id, r_Rolladvice.主页id, v_人员编号, v_人员姓名, Null, Null, Null, '预出院');
            End If;
          End If;
          Close c_Patilog;
        End If;
      
        --回退病历时机 
        --1.特殊事件(只有一条医嘱记录)：手术，7-会诊,8-抢救,11-死亡 
        If r_Rolladvice.类别 = 'F' Or r_Rolladvice.类别 = 'Z' And Instr(',7,8,11,', ',' || r_Rolladvice.类型 || ',') > 0 Then
          Zl_电子病历时机_Delete(r_Rolladvice.病人id, r_Rolladvice.主页id, '医嘱', r_Rolladvice.开嘱科室id, 医嘱id_In);
        End If;
      
        --2.额外处理：知情同意书(手术相关的知情同意需再次调用，因为附加手术或麻醉项目可能有关联的知情同意书) 
        If Instr('C,D,E,F,G,K,L', r_Rolladvice.类别) > 0 Then
          For R In (Select a.Id, a.诊疗类别 From 病人医嘱记录 A Where a.Id = 医嘱id_In Or a.相关id = 医嘱id_In) Loop
            --相关id的一组医嘱不一定是这个类别的，所以要再判断一次类别 
            If Instr('C,D,E,F,G,K,L', r.诊疗类别) > 0 Then
              Zl_电子病历时机_Delete(r_Rolladvice.病人id, r_Rolladvice.主页id, '医嘱', r_Rolladvice.开嘱科室id, r.Id);
            End If;
          End Loop;
        End If;
      
        If r_Rolladvice.类别 = 'Z' And r_Rolladvice.操作类型 = '6' Then
          --会诊 
          b_Message.Zlhis_Cis_040(r_Rolladvice.病人id, r_Rolladvice.主页id, r_Rolladvice.发送号, r_Rolladvice.组id,
                                  r_Rolladvice.No);
        Elsif r_Rolladvice.类别 = 'Z' And r_Rolladvice.操作类型 = '8' Then
          --抢救 
          b_Message.Zlhis_Cis_041(r_Rolladvice.病人id, r_Rolladvice.主页id, r_Rolladvice.发送号, r_Rolladvice.组id,
                                  r_Rolladvice.No);
        Elsif r_Rolladvice.类别 = 'Z' And r_Rolladvice.操作类型 = '11' Then
          --死亡 
          b_Message.Zlhis_Cis_042(r_Rolladvice.病人id, r_Rolladvice.主页id, r_Rolladvice.发送号, r_Rolladvice.组id,
                                  r_Rolladvice.No);
        Elsif r_Rolladvice.类别 = 'E' And r_Rolladvice.操作类型 = '5' Then
          --特殊治疗 
          b_Message.Zlhis_Cis_043(r_Rolladvice.病人id, r_Rolladvice.主页id, r_Rolladvice.发送号, r_Rolladvice.组id,
                                  r_Rolladvice.No);
        Elsif r_Rolladvice.类别 = 'H' And Nvl(r_Rolladvice.操作类型, '0') = '0' Then
          --护理常规 
          b_Message.Zlhis_Cis_007(r_Rolladvice.病人id, r_Rolladvice.主页id, r_Rolladvice.发送号, r_Rolladvice.组id,
                                  r_Rolladvice.No);
        End If;
      End If;
    End If;
    Exit When r_Rolladvice.发送号 = 0;
  End Loop;
  Close c_Rolladvice;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人医嘱记录_回退;
/
--138877:蒋廷中,2019-04-23,新增留观病人转住院病人消息
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
    原病人id_In In 病案主页.病人id%Type,
    变化ids_In  In Varchar2 
  ); 

  --69.患者转病区转入
  Procedure Zlhis_Patient_026
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );

  Procedure Zlhis_Patient_028(病人id_In In 病案主页.病人id%Type);


  --79.留观病人转住院病人
  Procedure Zlhis_Patient_029
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );


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
    原病人id_In In 病案主页.病人id%Type,
    变化ids_In  In Varchar2
  ) Is 
  --参数： 1病人id,1主页id:1原病人id,1原主页id; 2病人id,2主页id:2原病人id,2原主页id;….
  Begin 
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_017', 
                                '<root><病人ID>' || 病人id_In || '</病人ID><原病人ID>' || 原病人id_In || '</原病人ID><CINFO>'||变化ids_In||'</CINFO></root>'); 
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

  --79.留观病人转住院病人
  Procedure Zlhis_Patient_029
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
    n_变动id Number(18);
  Begin
    Select max(ID)
    Into n_变动id
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null And Nvl(附加床位, 0) = 0 And 开始原因 = 9;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_029',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><变动ID>' || n_变动id ||
                                 '</变动ID></root>');
  
  End Zlhis_Patient_029;

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



--138877:蒋廷中,2019-04-23,新增留观病人转住院病人消息
Create Or Replace Procedure Zl_病人变动记录_转住院
(
  病人id_In 病案主页.病人id%Type,
  主页id_In 病案主页.主页id%Type,
  住院号_In 病人信息.住院号%Type
) Is
  --功能：将住院留观病人转为住院病人
  v_Count      Number;
  v_出院科室id Number;
  v_Date       Date;
  v_Temp       Varchar2(255);
  v_人员编号   住院费用记录.操作员编号%Type;
  v_人员姓名   住院费用记录.操作员姓名%Type;
  Cursor c_Oldinfo Is
    Select b.*
    From (Select c.*
           From 病人变动记录 C
           Where c.病人id = 病人id_In And c.主页id = 主页id_In And
                 c.开始时间 = (Select Min(开始时间)
                           From 病人变动记录
                           Where 病人id = 病人id_In And 主页id = 主页id_In And 开始时间 > v_Date)) A, 病人变动记录 B
    
    Where b.病人id = 病人id_In And b.主页id = 主页id_In And a.开始时间 = b.终止时间 And a.开始原因 = b.终止原因 And a.附加床位 = b.附加床位
    Union
    Select *
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null And 开始时间 <= v_Date;

  Cursor c_Endinfo Is
    Select * From 病人变动记录 Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null;
  r_Oldinfo c_Oldinfo%RowType;
  r_Endinfo c_Endinfo%RowType;

  v_终止原因 病人变动记录.终止原因%Type;
  v_终止时间 病人变动记录.终止时间%Type;
  v_终止人员 病人变动记录.终止人员%Type;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --并发操作检查
  Select Nvl(状态, 0), 出院科室id
  Into v_Count, v_出院科室id
  From 病案主页
  Where 病人id = 病人id_In And 主页id = 主页id_In And 病人性质 = 2;
  If v_Count = 1 Then
    v_Error := '病人当前尚未入科,不能转为住院病人。请先将病人入科后再试。';
    Raise Err_Custom;
  Elsif v_Count = 2 Then
    v_Error := '病人当前正在转科,不能转为住院病人。请先将病人转科或取消转科后再试。';
    Raise Err_Custom;
  End If;

  Select Zl_住院日报_Count(v_出院科室id, Trunc(Sysdate)) Into v_Count From Dual;
  If v_Count > 0 Then
    v_Error := '已产生业务时间内的住院日报,不能办理该业务!';
    Raise Err_Custom;
  End If;

  Select Sysdate Into v_Date From Dual;
  v_Temp     := Zl_Identity;
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_人员编号 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);

  Open c_Oldinfo; --必须先打开
  Fetch c_Oldinfo
    Into r_Oldinfo;
  Open c_Endinfo;
  Fetch c_Endinfo
    Into r_Endinfo;
  If c_Endinfo%RowCount = 0 Then
    Close c_Endinfo;
    v_Error := '未发现该病人当前有效的变动记录！';
    Raise Err_Custom;
  End If;
  Select Count(*)
  Into v_Count
  From 病人变动记录
  Where 病人id = 病人id_In And 主页id = 主页id_In And 开始时间 Is Null And 终止时间 Is Null;
  If v_Count > 0 Then
    v_Error := '该病人正在转科或转病区，不能进行其他变动！';
    Raise Err_Custom;
  End If;

  --取消上次变动
  If r_Oldinfo.终止时间 Is Not Null Then
    v_终止时间 := r_Oldinfo.终止时间;
    v_终止原因 := r_Oldinfo.终止原因;
    v_终止人员 := r_Oldinfo.终止人员;
    --取消上次变动
    Update 病人变动记录
    Set 终止时间 = v_Date, 终止原因 = 9, 终止人员 = v_人员姓名, 上次计算时间 = Null
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 = v_终止时间 And 终止原因 = v_终止原因;
    --更新将来的记录如果有停止到将来的则删除上次计算时间
    Update 病人变动记录 Set 上次计算时间 = Null Where 病人id = 病人id_In And 主页id = 主页id_In And 开始时间 > v_Date;
  Else
    Update 病人变动记录
    Set 终止时间 = v_Date, 终止原因 = 9, 终止人员 = v_人员姓名
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null;
  End If;

  --产生变动记录
  While c_Oldinfo%Found Loop
    Insert Into 病人变动记录
      (ID, 病人id, 主页id, 开始时间, 开始原因, 附加床位, 病区id, 科室id, 医疗小组id, 护理等级id, 床位等级id, 床号, 责任护士, 经治医师, 主治医师, 主任医师, 病情, 操作员编号,
       操作员姓名, 终止时间, 终止原因, 终止人员)
    Values
      (病人变动记录_Id.Nextval, 病人id_In, 主页id_In, v_Date, 9, r_Oldinfo.附加床位, r_Oldinfo.病区id, r_Oldinfo.科室id, r_Oldinfo.医疗小组id,
       r_Oldinfo.护理等级id, r_Oldinfo.床位等级id, r_Oldinfo.床号, r_Oldinfo.责任护士, r_Oldinfo.经治医师, r_Oldinfo.主治医师, r_Oldinfo.主任医师,
       r_Oldinfo.病情, v_人员编号, v_人员姓名, v_终止时间, v_终止原因, v_终止人员);
    If Nvl(r_Oldinfo.附加床位, 0) = 0 Then
      --产生病历书写时机
      Zl_电子病历时机_Insert(病人id_In, 主页id_In, 2, '入院', r_Oldinfo.科室id, r_Oldinfo.经治医师, v_Date, v_Date);
    End If;
    Fetch c_Oldinfo
      Into r_Oldinfo;
  End Loop;

  Close c_Oldinfo;
  Close c_Endinfo;

  Update 病案主页 Set 病人性质 = 0, 住院号 = 住院号_In Where 病人id = 病人id_In And 主页id = 主页id_In;
  Update 病人信息 Set 住院号 = 住院号_In, 住院次数 = Nvl(住院次数, 0) + 1 Where 病人id = 病人id_In;

  --并发操作检查
  Select Count(*)
  Into v_Count
  From 病人变动记录
  Where 病人id = 病人id_In And 主页id = 主页id_In And Nvl(附加床位, 0) = 0 And 开始时间 Is Not Null And 终止时间 Is Null;

  If v_Count > 1 Then
    v_Error := '发现病人存在非法的变动记录,当前操作不能继续！' || Chr(13) || Chr(10) || '这可能是由于网络并发操作引起的,请刷新病人状态后再试！';
    Raise Err_Custom;
  End If;

 --留观转住院消息提醒
  b_Message.Zlhis_Patient_029(病人id_In, 主页id_In);
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人变动记录_转住院;
/

--139595:李业庆,2019-04-15,静配支持门诊费用数据
Create Or Replace Procedure Zl_输液配药记录_发送
(
  配药id_In   In Varchar2, --ID串:ID1,ID2....
  操作人员_In In 输液配药记录.操作人员%Type := Null,
  操作时间_In In 输液配药记录.操作时间%Type := Null,
  操作说明_In In 输液配药状态.操作说明%Type := Null
) Is
  v_Tansid     Varchar2(20);
  v_Tmp        Varchar2(4000);
  v_Error      Varchar2(255);
  n_操作状态   输液配药记录.操作状态%Type;
  n_People     Number(2);
  v_No         Varchar2(20);
  n_项目id     Number(18);
  v_收费项目id Varchar2(200);
  n_Row        Number(2);
  n_Out        Number(10);
  n_Outnum     Number(10);
  n_Count      Number(18);
  n_Packet     Number(2);
  v_Usercode   Varchar2(100);
  v_复核人     Varchar2(20);
  Err_Custom Exception;

  Cursor c_Bill Is
    Select a.病人id, a.主页id, a.标识号, a.姓名, a.性别, a.年龄, a.床号, a.费别, a.病人病区id, a.病人科室id, a.婴儿费, e.药品id, b.库房id, f.配药类型,
           c.执行时间, g.序号, 1 As 门诊标志
    From 住院费用记录 A, 药品收发记录 B, 输液配药记录 C, 输液配药内容 D, 药品规格 E, 输液药品属性 F, 配置收费方案 G
    Where a.Id = b.费用id And b.Id = d.收发id And d.记录id = c.Id And b.药品id = e.药品id And b.药品id = f.药品id And
          Substr(f.配药类型, Instr(f.配药类型, '-') + 1) = g.配药类型(+) And Nvl(c.是否打包, 0) <> 0 And c.Id = v_Tansid
    Union All
    Select a.病人id, a.主页id, a.标识号, a.姓名, a.性别, a.年龄, '' As 床号, a.费别, a.病人病区id, a.病人科室id, a.婴儿费, e.药品id, b.库房id, f.配药类型,
           c.执行时间, g.序号, 0 As 门诊标志
    From 门诊费用记录 A, 药品收发记录 B, 输液配药记录 C, 输液配药内容 D, 药品规格 E, 输液药品属性 F, 配置收费方案 G
    Where a.Id = b.费用id And b.Id = d.收发id And d.记录id = c.Id And b.药品id = e.药品id And b.药品id = f.药品id And
          Substr(f.配药类型, Instr(f.配药类型, '-') + 1) = g.配药类型(+) And Nvl(c.是否打包, 0) <> 0 And c.Id = v_Tansid
    Order By 序号;
Begin
  v_Usercode := Zl_Identity;
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ';') + 1);
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ',') + 1);
  v_Usercode := Substr(v_Usercode, 1, Instr(v_Usercode, ',') - 1);
  n_People   := Nvl(zl_GetSysParameter('配置费按病人收取', 1345), 0);
  n_Out      := Nvl(zl_GetSysParameter('出院病人不收配置费', 1345), 0);
  n_Packet   := Nvl(zl_GetSysParameter('打包药品在发送环节收取配置费', 1345), 0);

  If 配药id_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := 配药id_In || ',';
  End If;
  While v_Tmp Is Not Null Loop
    --分解单据ID串
    v_Tansid := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp    := Replace(',' || v_Tmp, ',' || v_Tansid || ',');
  
    Begin
      Select 操作状态 Into n_操作状态 From 输液配药记录 Where ID = v_Tansid;
    
      If n_操作状态 > 4 Then
        v_Error := '该数据已被操作，不能进行发送操作！';
        Raise Err_Custom;
      End If;
    Exception
      When Others Then
        v_Error := '该数据已被操作！';
        Raise Err_Custom;
    End;
  
    v_复核人 := '';
    Begin
      Select 复核人
      Into v_复核人
      From (Select e.复核人
             From 输液配药记录 A, 输液配药内容 B, 药品收发记录 C, 配液台药品对照 D, 配液工作安排 E
             Where a.Id = b.记录id And b.收发id = c.Id And c.药品id = d.药品id And c.库房id = d.部门id And d.配药台id = e.配药台id And
                   a.配药批次 = e.批次 And e.日期 = To_Date(To_Char(Sysdate, 'yyyy-mm-dd'), 'yyyy-mm-dd') And a.Id = v_Tansid And
                   Rownum = 1
             Order By d.配药台id)
      Where Rownum = 1;
    Exception
      When Others Then
        Null;
    End;
  
    Insert Into 输液配药状态
      (配药id, 操作类型, 操作人员, 操作时间, 操作说明, 实际工作人员)
    Values
      (v_Tansid, 5, 操作人员_In, 操作时间_In, 操作说明_In, v_复核人);
    Update 输液配药记录 Set 操作状态 = 5, 操作人员 = 操作人员_In, 操作时间 = 操作时间_In Where ID = v_Tansid;
  
    --打包药品收费
    If n_Packet = 1 Then
      n_Count := 0;
      Select Nextno(14) Into v_No From Dual;
    
      For r_Bill In c_Bill Loop
        Select Count(病人id)
        Into n_Outnum
        From 病案主页
        Where 主页id = r_Bill.主页id And 病人id = r_Bill.病人id And (Nvl(状态, 0) = 3 Or 出院日期 Is Not Null);
      
        --先查询是否有按给药途径收取的配置费方案
        Select Nvl(Max(项目id), 0)
        Into n_项目id
        From 输液配药记录 A, 病人医嘱记录 B, 配置收费方案 C
        Where a.Id = v_Tansid And a.医嘱id = b.Id And b.诊疗项目id = c.诊疗id;
        If n_项目id = 0 Then
          --若无对应给药途径的配置费收取方案，则再查询是否有按配药类型收取的配置费方案
          Select Nvl(Max(项目id), 0)
          Into n_项目id
          From 配置收费方案
          Where 配药类型 = Substr(r_Bill.配药类型, Instr(r_Bill.配药类型, '-', 1, 1) + 1);
        End If;
      
        If n_项目id <> 0 Then
          n_Row := 0;
        
          If n_People = 1 Then
            Select Count(配药id)
            Into n_Row
            From (Select 配药id
                   From 输液配药附费 A, 住院费用记录 B, 输液配药记录 C
                   Where a.No = b.No And a.配药id = c.Id And b.病人id = r_Bill.病人id And b.记录状态 = 1 And b.收费细目id = n_项目id And
                         r_Bill.执行时间 Between Trunc(c.执行时间) And Trunc(c.执行时间 + 1) - 1 / 24 / 60 / 60
                   Union All
                   Select 配药id
                   From 输液配药附费 A, 门诊费用记录 B, 输液配药记录 C
                   Where a.No = b.No And a.配药id = c.Id And b.病人id = r_Bill.病人id And b.记录状态 = 1 And b.收费细目id = n_项目id And
                         r_Bill.执行时间 Between Trunc(c.执行时间) And Trunc(c.执行时间 + 1) - 1 / 24 / 60 / 60);
          End If;
        Else
          n_Row := 1;
        End If;
      
        If n_Row = 0 And (n_Outnum = 0 Or n_Out = 0) Then
          For r_Item In (Select a.Id 收费细目id, a.类别 收费类别, a.计算单位, a.加班加价 加班标志, d.Id 收入项目id, d.收据费目, b.现价
                         From 收费项目目录 A, 收费价目 B, 收入项目 D
                         Where a.Id = b.收费细目id And b.收入项目id = d.Id And a.Id = n_项目id And b.执行日期 <= Sysdate And
                               (b.终止日期 >= Sysdate Or b.终止日期 Is Null)) Loop
            If n_Count = 0 Then
              Insert Into 输液配药附费 (配药id, NO, 病人id) Values (v_Tansid, v_No, r_Bill.病人id);
            End If;
          
            n_Count := n_Count + 1;
          
            If r_Bill.门诊标志 = 1 Then
              Zl_住院记帐记录_Insert(v_No, n_Count, r_Bill.病人id, r_Bill.主页id, r_Bill.标识号, r_Bill.姓名, r_Bill.性别, r_Bill.年龄,
                               r_Bill.床号, r_Bill.费别, r_Bill.病人病区id, r_Bill.病人科室id, r_Item.加班标志, r_Bill.婴儿费, r_Bill.库房id,
                               操作人员_In, Null, r_Item.收费细目id, r_Item.收费类别, r_Item.计算单位, Null, Null, Null, 1, 1, Null,
                               r_Bill.库房id, Null, r_Item.收入项目id, r_Item.收据费目, r_Item.现价, r_Item.现价, r_Item.现价, Null,
                               Sysdate, Sysdate, Null, Null, v_Usercode, 操作人员_In);
            Else
              Zl_门诊记帐记录_Insert(v_No, n_Count, r_Bill.病人id, r_Bill.标识号, r_Bill.姓名, r_Bill.性别, r_Bill.年龄, r_Bill.费别,
                               r_Item.加班标志, r_Bill.婴儿费, r_Bill.病人科室id, r_Bill.库房id, 操作人员_In, Null, r_Item.收费细目id,
                               r_Item.收费类别, r_Item.计算单位, 1, 1, Null, r_Bill.库房id, Null, r_Item.收入项目id, r_Item.收据费目,
                               r_Item.现价, r_Item.现价, r_Item.现价, Sysdate, Sysdate, Null, Null, v_Usercode, 操作人员_In);
            End If;
          End Loop;
        End If;
      
        If n_Row = 0 Then
          Exit;
        End If;
      End Loop;
    End If;
  End Loop;

  --消息处理
  b_Message.Zlhis_Drug_008(配药id_In);
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_输液配药记录_发送;
/

--139595:李业庆,2019-04-15,静配支持门诊费用数据
Create Or Replace Procedure Zl_输液配药记录_核查
(
  部门id_In   In 输液配药记录.部门id%Type,
  医嘱id_In   In Varchar2, --输液医嘱给药途径对应的医嘱ID:医嘱ID1,医嘱ID2...
  发送号_In   In 病人医嘱发送.发送号%Type,
  核查人_In   In 输液配药状态.操作人员%Type,
  核查时间_In In 输液配药状态.操作时间%Type
) Is
  v_Count    Number;
  v_序号     Number;
  v_执行时间 Date;

  v_相关id      Number;
  v_New相关id   Number;
  v_Old相关id   Number;
  v_发送号      Number;
  v_Tmp         Varchar2(200);
  I             Number;
  v_配药id      Number;
  v_批次        Number;
  v_Maxno       Varchar2(4000);
  v_Lableno     Varchar2(200);
  v_Maxbatch    Number;
  v_Curdose     Number;
  v_Sumdose     Number;
  v_Drugcount   Number;
  v_Currdate    Date;
  n_Needcheck   Number;
  n_Lngid       药品收发记录.Id%Type;
  n_Count       Number(3);
  n_单据        药品收发记录.单据%Type;
  v_No          药品收发记录.No%Type;
  n_发送次数    Number(5);
  n_病人id      病人信息.病人id%Type := 0;
  b_Change      Boolean := True;
  n_Sum         Number;
  n_调整批次    Number(1);
  n_Cur         Number(5);
  v_上次发送号  病人医嘱发送.发送号%Type;
  v_医嘱ids     Varchar2(4000);
  v_Tansid      Varchar2(12);
  v_当前病人    Varchar2(20);
  n_Num         Number(8);
  d_Old执行时间 Date;
  n_是否打包    Number(1);
  n_打包        Number(1);
  n_摆药单      Number(2);
  --控制参数
  v_医嘱类型       Number;
  v_输液总量       Number;
  v_大输液剂型     Varchar2(2000);
  v_大输液给药途径 Varchar2(2000);
  v_来源科室       Varchar2(4000);
  v_Continue       Number := 1;
  v_Nodosage       Number := 0;
  v_保持上次批次   Number := 0;
  d_手工打包时间   Date;
  n_Tpn处置方式    Number := 0;
  v_药品类型       Varchar2(20);
  n_打包药品批次   Number(1);
  n_特殊药品批次   Number(1);
  n_优先级         Number := 999;
  n_自动排批       Number := 0;
  n_科室id         Number := 0;
  n_Row            Number(2);
  n_备用批次       Number := 0;
  n_剩余数量       Number := 0;
  n_单次数量       Number := 0;
  n_累计数量       Number := 0;
  n_医嘱id         Number := 0;
  n_填写数量       Number := 0;
  v_配药类型       Varchar2(20);
  v_时间串         Varchar2(100);
  v_时间值         Date;
  v_Fields         Varchar2(100);
  v_是否改变       Varchar2(20);
  v_时间串1        Varchar2(100);
  Err_Item Exception;
  n_流通金额小数 Number;

  Cursor c_医嘱记录 Is
    Select /*+rule */
    Distinct e.医嘱id As 相关id, e.发送号, b.频率间隔, b.间隔单位, b.执行时间方案, a.姓名, a.性别, a.年龄, a.住院号, a.当前床号 As 床号, a.当前病区id As 病人病区id,
             a.当前科室id As 病人科室id, e.首次时间, e.末次时间, b.开始执行时间, Nvl(e.发送数次, 0) As 次数, e.发送时间, Decode(b.医嘱期效, 0, 1, 2) As 医嘱类型,
             b.诊疗项目id As 给药途径, b.病人id, Nvl(c.执行标记, 0) As 是否tpn
    From 病人医嘱发送 E, 病人医嘱记录 B, 病人信息 A, 诊疗项目目录 C, Table(f_Num2list(医嘱id_In)) D
    Where e.医嘱id = b.Id And b.病人id = a.病人id And c.类别 = 'E' And c.操作类型 = '2' And c.执行分类 = 1 And b.诊疗项目id = c.Id And
          e.医嘱id = d.Column_Value And e.发送号 = 发送号_In
    Order By b.病人id, e.医嘱id, e.发送号;

  Cursor c_单个医嘱记录 Is
    Select /*+rule */
    Distinct e.医嘱id, e.发送号, b.频率间隔, b.间隔单位, b.执行时间方案, a.姓名, a.性别, a.年龄, a.住院号, a.当前床号 As 床号, a.当前病区id As 病人病区id,
             a.当前科室id As 病人科室id, e.首次时间, e.末次时间, b.开始执行时间, Nvl(e.发送数次, 0) As 次数, e.发送时间, Decode(b.医嘱期效, 0, 1, 2) As 医嘱类型,
             b.诊疗项目id As 给药途径, b.病人id
    From 病人医嘱发送 E, 病人医嘱记录 B, 病人信息 A, 诊疗项目目录 C
    Where e.医嘱id = b.Id And b.病人id = a.病人id And b.诊疗项目id = c.Id And b.相关id = v_相关id And e.发送号 = 发送号_In
    Order By e.医嘱id, e.发送号;

  Cursor c_收发记录 Is
    Select Distinct c.Id As 收发id, c.序号, c.实际数量 As 数量, Nvl(e.是否不予配置, 0) As 是否不予配置, c.单据, c.No
    From 病人医嘱记录 A, 病人医嘱发送 B, 药品收发记录 C, 住院费用记录 D, 输液药品属性 E
    Where a.Id = b.医嘱id And c.费用id = d.Id And a.Id = d.医嘱序号 And b.No = c.No And c.药品id = e.药品id(+) And
          b.执行部门id + 0 = 部门id_In And c.单据 = 9 And c.审核日期 Is Null And a.Id = n_医嘱id And b.发送号 = 发送号_In And c.序号 < 1000
    Union All
    Select Distinct c.Id As 收发id, c.序号, c.实际数量 As 数量, Nvl(e.是否不予配置, 0) As 是否不予配置, c.单据, c.No
    From 病人医嘱记录 A, 病人医嘱发送 B, 药品收发记录 C, 门诊费用记录 D, 输液药品属性 E
    Where a.Id = b.医嘱id And c.费用id = d.Id And a.Id = d.医嘱序号 And b.No = c.No And c.药品id = e.药品id(+) And
          b.执行部门id + 0 = 部门id_In And c.单据 = 9 And c.审核日期 Is Null And a.Id = n_医嘱id And b.发送号 = 发送号_In And c.序号 < 1000
    Order By NO, 序号;

  Cursor c_原始收发记录 Is
    Select Distinct c.Id As 收发id, c.序号, c.实际数量 As 数量, Nvl(e.是否不予配置, 0) As 是否不予配置, c.单据, c.No
    From 病人医嘱记录 A, 病人医嘱发送 B, 药品收发记录 C, 住院费用记录 D, 输液药品属性 E
    Where a.Id = b.医嘱id And c.费用id = d.Id And a.Id = d.医嘱序号 And b.No = c.No And c.药品id = e.药品id(+) And
          b.执行部门id + 0 = 部门id_In And c.单据 = 9 And c.审核日期 Is Null And a.相关id = v_相关id And b.发送号 = 发送号_In And c.序号 < 1000
    Union All
    Select Distinct c.Id As 收发id, c.序号, c.实际数量 As 数量, Nvl(e.是否不予配置, 0) As 是否不予配置, c.单据, c.No
    From 病人医嘱记录 A, 病人医嘱发送 B, 药品收发记录 C, 门诊费用记录 D, 输液药品属性 E
    Where a.Id = b.医嘱id And c.费用id = d.Id And a.Id = d.医嘱序号 And b.No = c.No And c.药品id = e.药品id(+) And
          b.执行部门id + 0 = 部门id_In And c.单据 = 9 And c.审核日期 Is Null And a.相关id = v_相关id And b.发送号 = 发送号_In And c.序号 < 1000
    Order By NO, 序号;

  Cursor c_输液单记录 Is
    Select a.Id, a.执行时间, a.配药批次, a.医嘱id, d.发送时间
    From 输液配药记录 A, 病人医嘱记录 B, 配药工作批次 C, 病人医嘱发送 D
    Where a.医嘱id = b.Id And a.配药批次 = c.批次 And d.医嘱id = a.医嘱id And a.发送号 = d.发送号 And c.批次 <> 0 And c.药品类型 Is Null And
          b.病人id = n_病人id And a.操作状态 < 2 And a.执行时间 Between Trunc(v_时间值) And Trunc(v_时间值 + 1) - 1 / 24 / 60 / 60;

  v_输液单记录   c_输液单记录%RowType;
  v_医嘱记录     c_医嘱记录%RowType;
  v_收发记录     c_收发记录%RowType;
  v_单个医嘱记录 c_单个医嘱记录%RowType;

  Function Zl_Getpivaworkbatch
  (
    执行时间_In In Date,
    发送时间_In In Date,
    药品类型_In In Varchar2 := Null
  ) Return Number As
  
    v_Exetime   Date;
    v_Starttime Date;
    v_Endtime   Date;
    v_Maxbatch  Number(2);
    v_Batch     Number;
    Cursor c_配药批次 Is
      Select 批次, 配药时间, 给药时间, 打包, 药品类型
      From 配药工作批次
      Where 启用 = 1 And 配置中心id = 部门id_In
      Order By 药品类型, 批次;
  
    v_配药批次 c_配药批次%RowType;
  Begin
    v_Exetime := To_Date(Substr(To_Char(执行时间_In, 'yyyy-mm-dd hh24:mi'), 12), 'hh24:mi');
  
    Select Max(Nvl(批次, 0)) + 1 Into v_Maxbatch From 配药工作批次 Where 启用 = 1 And 配置中心id = 部门id_In;
  
    For v_配药批次 In c_配药批次 Loop
      v_Batch := 0;
    
      --当天发送的医嘱发送到备用批次
      If (Trunc(执行时间_In) >= Trunc(v_Currdate) And Trunc(发送时间_In) < Trunc(执行时间_In)) Or n_备用批次 = 0 Then
        If v_配药批次.批次 <> '0' And
           ((Nvl(v_配药批次.药品类型, '0') <> '0' And v_配药批次.药品类型 = 药品类型_In) Or Nvl(v_配药批次.药品类型, '0') = '0') Then
          v_Starttime := To_Date(Substr(v_配药批次.给药时间, 1, Instr(v_配药批次.给药时间, '-') - 1), 'hh24:mi');
          v_Endtime   := To_Date(Substr(v_配药批次.给药时间, Instr(v_配药批次.给药时间, '-') + 1), 'hh24:mi');
        
          If v_Exetime >= v_Starttime And v_Exetime <= v_Endtime Then
            v_Batch := v_配药批次.批次;
            n_打包  := v_配药批次.打包;
            Exit When v_Batch > 0;
          End If;
        End If;
      End If;
    End Loop;
  
    If v_Batch = 0 And (n_打包药品批次 <> 1 Or n_备用批次 = 1) Then
      v_Batch := v_Maxbatch;
    End If;
    Return(v_Batch);
  End;

  Function Zl_Getfirst
  (
    配药id_In In Number,
    科室id_In In Number
  ) Return Number As
    n_First  Number;
    n_科室id Number;
    Cursor c_优先级 Is
      Select 科室id, 配药类型, 优先级, 频次
      From 输液药品优先级
      Where (科室id = 科室id_In Or 科室id = 0)
      Order By 科室id, 优先级 Desc;
  
    r_优先级 c_优先级%RowType;
  Begin
    n_First := 0;
    For r_优先级 In c_优先级 Loop
      If n_科室id <> 0 And r_优先级.科室id = 0 Then
        Exit;
      End If;
      n_科室id := r_优先级.科室id;
    
      For r_配药记录 In (Select Distinct d.配药类型, e.执行频次
                     From 输液配药记录 A, 输液配药内容 B, 药品收发记录 C, 输液药品属性 D, 病人医嘱记录 E
                     Where a.医嘱id = e.Id And a.Id = b.记录id And b.收发id = c.Id And c.药品id = d.药品id And a.Id = 配药id_In) Loop
        If Instr(r_配药记录.配药类型, r_优先级.配药类型, 1) > 0 And (Instr(r_优先级.频次, r_配药记录.执行频次, 1) > 0 Or r_优先级.频次 = '所有频次') Then
          n_First := r_优先级.优先级;
          Exit;
        End If;
      End Loop;
    End Loop;
  
    If n_First = 0 Then
      n_First := 999;
    End If;
    Return(n_First);
  End;
Begin
  n_Count          := 0;
  v_医嘱类型       := Zl_To_Number(Nvl(zl_GetSysParameter('医嘱类型', 1345), 1));
  v_输液总量       := Zl_To_Number(Nvl(zl_GetSysParameter('同批次输液总量', 1345), 0));
  v_大输液剂型     := Nvl(zl_GetSysParameter('大输液药品剂型', 1345), '');
  v_大输液给药途径 := Nvl(zl_GetSysParameter('输液给药途径', 1345), '');
  v_来源科室       := Nvl(zl_GetSysParameter('来源科室', 1345), '');
  v_保持上次批次   := Zl_To_Number(Nvl(zl_GetSysParameter('保持上次批次', 1345), 0));
  n_Tpn处置方式    := Zl_To_Number(Nvl(zl_GetSysParameter('静脉营养药物处置方式', 1345), 0));
  n_打包药品批次   := Zl_To_Number(Nvl(zl_GetSysParameter('单个药品，不予配置药品及根据给药时间没有配药批次的输液单默认为0批次并打包', 1345), 0));
  n_特殊药品批次   := Zl_To_Number(Nvl(zl_GetSysParameter('特殊药品按药品类型指定批次', 1345), 0));
  n_自动排批       := Zl_To_Number(Nvl(zl_GetSysParameter('启动自动排批', 1345), 0));
  n_备用批次       := Zl_To_Number(Nvl(zl_GetSysParameter('当天发送的医嘱产生的输液单全部到备用批次', 1345), 0));
  v_医嘱ids        := 医嘱id_In;
  v_当前病人       := '';
  n_发送次数       := 0;

  --取流通业务精度位数
  --类别:1-药品 2-卫材
  --内容：2-零售价 4-金额
  --单位：药品:1-售价 5-金额单位
  Select 精度 Into n_流通金额小数 From 药品卫材精度 Where 类别 = 1 And 内容 = 4 And 单位 = 5;

  Select Trunc(Sysdate) Into v_Currdate From Dual;

  Select Max(Nvl(批次, 0)) + 1 Into v_Maxbatch From 配药工作批次;

  --检查当前病人的医嘱是否有今天需要执行的输液单是锁定状态的
  If Instr(v_医嘱ids, ',') = 0 Then
    v_Tansid := v_医嘱ids;
  Else
    v_Tansid := Substr(v_Tmp, 1, Instr(v_医嘱ids, ',') - 1);
  End If;

  Select Count(ID)
  Into n_Num
  From 输液配药记录
  Where 是否锁定 = 1 And 执行时间 Between Trunc(v_Currdate) And Trunc(v_Currdate + 1) - 1 / 24 / 60 / 60 And
        医嘱id In
        (Select 相关id
         From 病人医嘱记录
         Where 病人id = (Select 病人id From 病人医嘱记录 Where 相关id = v_Tansid And Rownum < 2) And (诊疗类别 = '5' Or 诊疗类别 = '6')) And
        Rownum < 2;

  If n_Num > 0 Then
    Select 姓名
    Into v_当前病人
    From 输液配药记录
    Where 是否锁定 = 1 And 执行时间 Between Trunc(v_Currdate) And Trunc(v_Currdate + 1) - 1 / 24 / 60 / 60 And
          医嘱id In
          (Select 相关id
           From 病人医嘱记录
           Where 病人id = (Select 病人id From 病人医嘱记录 Where 相关id = v_Tansid And Rownum < 2) And (诊疗类别 = '5' Or 诊疗类别 = '6')) And
          Rownum < 2;
    Raise Err_Item;
  End If;

  --先将原收发记录的序号增大，新的收发记录产生后再删除
  --Update 药品收发记录
  --Set 序号 = 序号 + 10000
  --Where ID In (Select \*+rule *\
  --             Distinct c.Id
  --             From 病人医嘱记录 A, 病人医嘱发送 B, 药品收发记录 C, 住院费用记录 D, Table(f_Num2list(医嘱id_In)) F
  --             Where a.Id = b.医嘱id And c.费用id = d.Id And a.Id = d.医嘱序号 And b.No = c.No And b.执行部门id + 0 = 部门id_In And
  --                   c.单据 = 9 And c.审核日期 Is Null And a.相关id = f.Column_Value And b.发送号 = 发送号_In And c.序号 < 10000);

  For v_医嘱记录 In c_医嘱记录 Loop
    v_Continue := 1;
    n_病人id   := v_医嘱记录.病人id;
    n_科室id   := v_医嘱记录.病人科室id;
  
    Select Count(1)
    Into v_Continue
    From (Select 1
           From 病人医嘱记录 A, 输液不配置药品 B, 住院费用记录 C
           Where c.收费细目id = b.药品id And c.医嘱序号 = a.Id And a.相关id = v_医嘱记录.相关id
           Union All
           Select 1
           From 病人医嘱记录 A, 输液不配置药品 B, 门诊费用记录 C
           Where c.收费细目id = b.药品id And c.医嘱序号 = a.Id And a.相关id = v_医嘱记录.相关id);
    If v_Continue = 0 Then
      v_Continue := 1;
    Else
      v_Continue := 0;
    End If;
  
    --参数控制产生输液单
    If (v_医嘱类型 = 1 And v_医嘱记录.医嘱类型 <> 1) Or (v_医嘱类型 = 2 And v_医嘱记录.医嘱类型 <> 2) Then
      v_Continue := 0;
    End If;
  
    If Not v_大输液给药途径 Is Null Then
      If Instr(',' || v_大输液给药途径 || ',', ',' || v_医嘱记录.给药途径 || ',') = 0 Then
        v_Continue := 0;
      End If;
    End If;
  
    If Not v_来源科室 Is Null Then
      If Instr(',' || v_来源科室 || ',', ',' || v_医嘱记录.病人科室id || ',') = 0 Then
        v_Continue := 0;
      End If;
    End If;
  
    v_药品类型 := Null;
    For r_药品类型 In (Select Decode(Nvl(d.抗生素, 0), 0, Decode(Nvl(d.是否肿瘤药, 0), 0, '', '肿瘤药'), '抗生素') 药品类型
                   From 病人医嘱记录 A, 药品规格 B, 住院费用记录 C, 药品特性 D
                   Where c.收费细目id = b.药品id And b.药名id = d.药名id And c.医嘱序号 = a.Id And a.相关id = v_医嘱记录.相关id
                   Union All
                   Select Decode(Nvl(d.抗生素, 0), 0, Decode(Nvl(d.是否肿瘤药, 0), 0, '', '肿瘤药'), '抗生素') 药品类型
                   From 病人医嘱记录 A, 药品规格 B, 门诊费用记录 C, 药品特性 D
                   Where c.收费细目id = b.药品id And b.药名id = d.药名id And c.医嘱序号 = a.Id And a.相关id = v_医嘱记录.相关id) Loop
      If r_药品类型.药品类型 Is Not Null Then
        v_药品类型 := r_药品类型.药品类型;
      End If;
    End Loop;
  
    If v_药品类型 Is Null Then
      If v_医嘱记录.是否tpn = 2 Then
        v_药品类型 := '营养药';
      End If;
    End If;
  
    If v_Continue = 1 Then
      v_Old相关id := v_New相关id;
      v_相关id    := v_医嘱记录.相关id;
      v_New相关id := v_相关id;
      v_发送号    := v_医嘱记录.发送号;
      v_序号      := 0;
    
      If v_Continue = 1 Then
        --v_Count := Zl_Gettransexenumber(v_医嘱记录.开始执行时间, v_医嘱记录.首次时间, v_医嘱记录.末次时间, v_医嘱记录.频率间隔, v_医嘱记录.间隔单位, v_医嘱记录.执行时间方案);
        Select Count(医嘱id)
        Into v_Count
        From 医嘱执行时间
        Where 医嘱id = v_医嘱记录.相关id And 发送号 = v_医嘱记录.发送号;
      
        v_Nodosage := 0;
      
        For I In 1 .. v_Count Loop
          Select 输液配药记录_Id.Nextval Into v_配药id From Dual;
          v_序号 := v_序号 + 1;
        
          If I > 1 Then
            --从医嘱执行时间表中取医嘱的执行时间
            Select 要求时间
            Into v_执行时间
            From 医嘱执行时间
            Where 医嘱id = v_医嘱记录.相关id And 发送号 = v_医嘱记录.发送号 And 要求时间 > v_执行时间 And Rownum = 1
            Order By 要求时间;
          Else
            Select Min(要求时间)
            Into v_执行时间
            From 医嘱执行时间
            Where 医嘱id = v_医嘱记录.相关id And 发送号 = v_医嘱记录.发送号 And Rownum = 1
            Order By 要求时间;
          End If;
        
          v_批次 := 0;
        
          If d_Old执行时间 <> Trunc(v_执行时间) Or d_Old执行时间 Is Null Then
            b_Change := True;
          End If;
        
          If b_Change = True Then
            If d_Old执行时间 <> Trunc(v_执行时间) Or d_Old执行时间 Is Null Then
              d_Old执行时间 := v_执行时间;
            
              Select Count(Distinct a.摆药单号)
              Into n_摆药单
              From 输液配药记录 A
              Where a.医嘱id In (Select ID From 病人医嘱记录 Where 病人id = v_医嘱记录.病人id And 相关id Is Null) And
                    a.执行时间 Between Trunc(v_执行时间 - 1) And Trunc(v_执行时间) - 1 / 24 / 60 / 60 And 操作状态 >= 2 And 操作状态 < 9;
            
              If n_摆药单 > 1 Then
                Update 输液配药记录
                Set 是否调整批次 = 1
                Where 医嘱id In (Select ID From 病人医嘱记录 Where 病人id = n_病人id) And
                     
                      执行时间 Between Trunc(v_执行时间) And Trunc(v_执行时间 + 1) - 1 / 24 / 60 / 60 And 操作状态 < 2;
                b_Change := False;
              
              End If;
            End If;
          End If;
        
          If b_Change = True Then
            n_病人id := v_医嘱记录.病人id;
            Select Count(ID)
            
            Into n_Sum
            From 输液配药记录
            Where 医嘱id = v_医嘱记录.相关id And 执行时间 Between Trunc(v_执行时间 - 1) And Trunc(v_执行时间) - 1 / 24 / 60 / 60;
            If n_Sum = 0 Then
              Update 输液配药记录
              Set 是否调整批次 = 1
              Where 医嘱id In (Select ID From 病人医嘱记录 Where 病人id = n_病人id) And
                   
                    执行时间 Between Trunc(v_执行时间) And Trunc(v_执行时间 + 1) - 1 / 24 / 60 / 60 And 操作状态 < 2;
              b_Change := False;
            
            End If;
          
            If b_Change = True Then
              --检查输液单是否调整到打包状态
              Select Count(a.Id)
              Into n_Sum
              From 输液配药记录 A, 输液配药内容 B, 药品收发记录 C
              Where a.Id = b.记录id And b.收发id = c.Id And
                    a.医嘱id In (Select ID
                               From 病人医嘱记录
                               Where 病人id = (Select 病人id From 病人医嘱记录 Where ID = v_医嘱记录.相关id And Rownum < 2)) And
                    a.执行时间 Between Trunc(v_执行时间 - 1) And Trunc(v_执行时间) - 1 / 24 / 60 / 60 And a.打包时间 Is Not Null;
              If n_Sum <> 0 Then
                Update 输液配药记录
                Set 是否调整批次 = 1
                Where 医嘱id In (Select ID From 病人医嘱记录 Where 病人id = n_病人id) And 执行时间 Between Trunc(v_执行时间) And
                      Trunc(v_执行时间 + 1) - 1 / 24 / 60 / 60 And 操作状态 < 2;
                b_Change := False;
              End If;
            
              Select Count(医嘱id)
              Into n_Cur
              From 医嘱执行时间
              Where 医嘱id = v_医嘱记录.相关id And 要求时间 Between Trunc(v_执行时间) And Trunc(v_执行时间 + 1) - 1 / 24 / 60 / 60;
            
              Select Count(医嘱id)
              Into n_Sum
              From 医嘱执行时间
              Where 医嘱id = v_医嘱记录.相关id And 要求时间 Between Trunc(v_执行时间 - 1) And Trunc(v_执行时间) - 1 / 24 / 60 / 60;
            
              If n_Sum <> n_Cur Then
                Update 输液配药记录
                Set 是否调整批次 = 1
                Where 医嘱id In (Select ID From 病人医嘱记录 Where 病人id = n_病人id) And 执行时间 Between Trunc(v_执行时间) And
                      Trunc(v_执行时间 + 1) - 1 / 24 / 60 / 60 And 操作状态 < 2;
                b_Change := False;
              End If;
            End If;
          End If;
        
          If v_时间串 <> Trunc(Sysdate) || ';false\' Or v_时间串 Is Null Then
            If Trunc(v_执行时间) = Trunc(Sysdate) Then
              If b_Change = False Then
                v_时间串 := Trunc(v_执行时间) || ';false\';
              Else
                v_时间串 := Trunc(v_执行时间) || ';true\';
              End If;
            End If;
          End If;
        
          If v_时间串1 <> Trunc(Sysdate + 1) || ';false\' Or v_时间串1 Is Null Then
            If Trunc(v_执行时间) = Trunc(Sysdate + 1) Then
              If b_Change = False Then
                v_时间串1 := Trunc(v_执行时间) || ';false\';
              Else
                v_时间串1 := Trunc(v_执行时间) || ';true\';
              End If;
            End If;
          End If;
        
          If v_药品类型 Is Null Or n_特殊药品批次 = 0 Then
            v_批次 := Zl_Getpivaworkbatch(v_执行时间, Sysdate);
          Else
            --药品类型不为空，直接根据药品类型匹配批次
            v_批次 := Zl_Getpivaworkbatch(v_执行时间, Sysdate, v_药品类型);
          End If;
        
          Select Count(医嘱id)
          Into n_发送次数
          From 医嘱执行时间
          Where 医嘱id = v_医嘱记录.相关id And 要求时间 <= v_执行时间
          Order By 要求时间;
        
          If n_发送次数 > 99 Then
            n_发送次数 := Mod(n_发送次数, 99);
          End If;
        
          If Length(v_医嘱记录.相关id) > 9 Then
            If n_发送次数 < 10 Then
              Select '91' || Substr(To_Char(v_医嘱记录.相关id), Length(v_医嘱记录.相关id) - 8) || To_Char(v_医嘱记录.相关id) || '0' ||
                      To_Char(n_发送次数)
              Into v_Maxno
              From Dual;
            Else
              Select '91' || Substr(To_Char(v_医嘱记录.相关id), Length(v_医嘱记录.相关id) - 8) || To_Char(v_医嘱记录.相关id) ||
                      To_Char(n_发送次数)
              Into v_Maxno
              From Dual;
            End If;
          Else
            If n_发送次数 < 10 Then
              Select '91' || Substr('000000000', Length(v_医嘱记录.相关id) + 1) || To_Char(v_医嘱记录.相关id) || '0' ||
                      To_Char(n_发送次数)
              Into v_Maxno
              From Dual;
            Else
              Select '91' || Substr('000000000', Length(v_医嘱记录.相关id) + 1) || To_Char(v_医嘱记录.相关id) || To_Char(n_发送次数)
              Into v_Maxno
              From Dual;
            End If;
          End If;
          n_调整批次 := 0;
          If b_Change = False Then
            n_调整批次 := 1;
          End If;
        
          If v_批次 <> 0 Then
            Select Nvl(Max(打包), 0), Max(药品类型)
            Into n_打包, v_配药类型
            From 配药工作批次
            Where 批次 = v_批次 And 配置中心id = 部门id_In;
          End If;
        
          If (Trunc(v_执行时间) <= v_Currdate Or n_打包 <> 0) And v_配药类型 Is Null Then
            n_是否打包     := 1;
            d_手工打包时间 := Null;
          Else
            n_是否打包     := 0;
            d_手工打包时间 := Null;
          End If;
        
          --如果是TPN不管其他条件如何都设置为配置
          If v_医嘱记录.是否tpn = 2 Then
            n_是否打包     := 0;
            d_手工打包时间 := Null;
          End If;
        
          If v_批次 = 0 Then
            n_是否打包 := 1;
          End If;
          --产生配药记录
          Insert Into 输液配药记录
            (ID, 部门id, 序号, 姓名, 性别, 年龄, 住院号, 床号, 病人病区id, 病人科室id, 执行时间, 医嘱id, 发送号, 配药批次, 瓶签号, 是否调整批次, 是否打包, 打包时间, 操作状态,
             操作人员, 操作时间)
          Values
            (v_配药id, 部门id_In, v_序号, v_医嘱记录.姓名, v_医嘱记录.性别, v_医嘱记录.年龄, v_医嘱记录.住院号, v_医嘱记录.床号, v_医嘱记录.病人病区id,
             v_医嘱记录.病人科室id, v_执行时间, v_医嘱记录.相关id, v_医嘱记录.发送号, v_批次, v_Maxno, n_调整批次, n_是否打包, d_手工打包时间, 1, 核查人_In, 核查时间_In);
        
          Insert Into 输液配药状态 (配药id, 操作类型, 操作人员, 操作时间) Values (v_配药id, 1, 核查人_In, 核查时间_In);
        
          For v_单个医嘱记录 In c_单个医嘱记录 Loop
            n_医嘱id   := v_单个医嘱记录.医嘱id;
            n_累计数量 := 0;
            n_剩余数量 := 0;
          
            Select Nvl(Sum(实际数量), 0)
            Into n_Sum
            From (Select c.实际数量
                   From 病人医嘱记录 A, 病人医嘱发送 B, 药品收发记录 C, 住院费用记录 D
                   Where a.Id = b.医嘱id And c.费用id = d.Id And a.Id = d.医嘱序号 And b.No = c.No And b.执行部门id + 0 = 部门id_In And
                         c.单据 = 9 And c.审核日期 Is Null And a.Id = n_医嘱id And b.发送号 = v_医嘱记录.发送号 And c.序号 < 1000
                   Union All
                   Select c.实际数量
                   From 病人医嘱记录 A, 病人医嘱发送 B, 药品收发记录 C, 门诊费用记录 D
                   Where a.Id = b.医嘱id And c.费用id = d.Id And a.Id = d.医嘱序号 And b.No = c.No And b.执行部门id + 0 = 部门id_In And
                         c.单据 = 9 And c.审核日期 Is Null And a.Id = n_医嘱id And b.发送号 = v_医嘱记录.发送号 And c.序号 < 1000);
          
            --产生配药记录对应的药品记录
            For v_收发记录 In c_收发记录 Loop
              If v_收发记录.是否不予配置 = 1 Then
                v_Nodosage := 1;
              End If;
            
              Select 药品收发记录_Id.Nextval Into n_Lngid From Dual;
              n_累计数量 := n_累计数量 + v_收发记录.数量;
            
              If n_剩余数量 = 0 Then
                n_剩余数量 := n_Sum / v_Count;
              End If;
              n_单次数量 := n_Sum / v_Count;
            
              If n_累计数量 >= n_Sum / v_Count * I Then
                n_Count := n_Count + 1;
                Insert Into 药品收发记录
                  (ID, 记录状态, 单据, NO, 序号, 库房id, 供药单位id, 入出类别id, 对方部门id, 入出系数, 药品id, 批次, 产地, 批号, 生产日期, 效期, 付数, 填写数量, 实际数量,
                   成本价, 成本金额, 扣率, 零售价, 零售金额, 差价, 摘要, 填制人, 填制日期, 配药人, 配药日期, 审核人, 审核日期, 价格id, 费用id, 单量, 频次, 用法, 外观, 灭菌日期,
                   灭菌效期, 产品合格证, 发药方式, 发药窗口, 领用人, 批准文号, 汇总发药号, 注册证号, 库房货位, 商品条码, 内部条码, 核查人, 核查日期, 签到确认人, 签到时间)
                  Select n_Lngid, 记录状态, 单据, NO, n_Count + 1000, 库房id, 供药单位id, 入出类别id, 对方部门id, 入出系数, 药品id, 批次, 产地, 批号,
                         生产日期, 效期, 付数, n_剩余数量, n_剩余数量, 成本价, Round(成本价 * n_剩余数量, n_流通金额小数), 扣率, 零售价,
                         Round(零售价 * n_剩余数量, n_流通金额小数), Round(差价 * (实际数量 / n_剩余数量), n_流通金额小数), '复制', 填制人, 填制日期, 配药人,
                         配药日期, 审核人, 审核日期, 价格id, 费用id, 单量, 频次, 用法, 外观, 灭菌日期, 灭菌效期, 产品合格证, 发药方式, 发药窗口, 领用人, 批准文号, 汇总发药号,
                         注册证号, 库房货位, 商品条码, 内部条码, 核查人, 核查日期, 签到确认人, 签到时间
                  From 药品收发记录
                  Where ID = v_收发记录.收发id;
              
                Zl_未审药品记录_Insert(n_Lngid);
              
                Insert Into 输液配药内容 (记录id, 收发id, 数量) Values (v_配药id, n_Lngid, n_剩余数量);
              
                n_剩余数量 := 0;
                Exit;
              Elsif n_累计数量 > (n_Sum / v_Count * (I - 1)) Then
                n_Count    := n_Count + 1;
                n_填写数量 := n_累计数量 - (n_Sum / v_Count * (I - 1)) - (n_单次数量 - n_剩余数量);
                Insert Into 药品收发记录
                  (ID, 记录状态, 单据, NO, 序号, 库房id, 供药单位id, 入出类别id, 对方部门id, 入出系数, 药品id, 批次, 产地, 批号, 生产日期, 效期, 付数, 填写数量, 实际数量,
                   成本价, 成本金额, 扣率, 零售价, 零售金额, 差价, 摘要, 填制人, 填制日期, 配药人, 配药日期, 审核人, 审核日期, 价格id, 费用id, 单量, 频次, 用法, 外观, 灭菌日期,
                   灭菌效期, 产品合格证, 发药方式, 发药窗口, 领用人, 批准文号, 汇总发药号, 注册证号, 库房货位, 商品条码, 内部条码, 核查人, 核查日期, 签到确认人, 签到时间)
                  Select n_Lngid, 记录状态, 单据, NO, n_Count + 1000, 库房id, 供药单位id, 入出类别id, 对方部门id, 入出系数, 药品id, 批次, 产地, 批号,
                         生产日期, 效期, 付数, n_填写数量, n_填写数量, 成本价, Round(成本价 * n_填写数量, n_流通金额小数), 扣率, 零售价,
                         Round(零售价 * n_填写数量, n_流通金额小数), Round(差价 * (实际数量 / n_填写数量), n_流通金额小数), '复制', 填制人, 填制日期, 配药人,
                         配药日期, 审核人, 审核日期, 价格id, 费用id, 单量, 频次, 用法, 外观, 灭菌日期, 灭菌效期, 产品合格证, 发药方式, 发药窗口, 领用人, 批准文号, 汇总发药号,
                         注册证号, 库房货位, 商品条码, 内部条码, 核查人, 核查日期, 签到确认人, 签到时间
                  From 药品收发记录
                  Where ID = v_收发记录.收发id;
              
                Zl_未审药品记录_Insert(n_Lngid);
              
                Insert Into 输液配药内容 (记录id, 收发id, 数量) Values (v_配药id, n_Lngid, n_填写数量);
              
                n_剩余数量 := n_剩余数量 - n_填写数量;
              End If;
            End Loop;
          End Loop;
          n_优先级 := Zl_Getfirst(v_配药id, v_医嘱记录.病人科室id);
          Update 输液配药记录 Set 优先级 = n_优先级 Where ID = v_配药id;
        
        End Loop;
      
        For v_收发记录 In c_原始收发记录 Loop
          n_单据 := v_收发记录.单据;
        
          v_No := v_收发记录.No;
          Delete From 药品收发记录 Where ID = v_收发记录.收发id;
        End Loop;
      
        --单个药品或者不予配置的药品默认为0批次
        Select Count(收发id) Into n_Row From 输液配药内容 Where 记录id = v_配药id;
        If (v_Nodosage = 1 Or n_Row = 1) And n_打包药品批次 = 1 Then
          Update 输液配药记录
          Set 配药批次 = 0, 是否打包 = 1
          Where 医嘱id = v_医嘱记录.相关id And 发送号 = v_医嘱记录.发送号 And 操作状态 < 2;
        End If;
        --如果存在“不予配置”属性的药品，也设置为打包
        If v_Nodosage = 1 Then
          Update 输液配药记录
          Set 是否打包 = 1
          Where 医嘱id = v_医嘱记录.相关id And 发送号 = v_医嘱记录.发送号 And 操作状态 < 2;
        End If;
      End If;
    End If;
  End Loop;

  For v_收发记录 In (Select ID From 药品收发记录 Where 序号 < 1000 And 单据 = n_单据 And NO = v_No) Loop
    n_Count := n_Count + 1;
    Update 药品收发记录 Set 序号 = n_Count + 1000, 摘要 = '复制' Where ID = v_收发记录.Id;
  End Loop;

  Update 药品收发记录
  Set 序号 = 序号 - 1000, 摘要 = '医嘱发送'
  Where 摘要 = '复制' And 序号 > 1000 And 单据 = n_单据 And NO = v_No;

  If n_备用批次 = 1 Then
  
    Select Count(a.Id)
    Into n_Sum
    From 输液配药记录 A, 病人医嘱发送 B
    Where a.医嘱id = b.医嘱id And a.发送号 = b.发送号 And
          a.医嘱id In (Select ID From 病人医嘱记录 Where 病人id = n_病人id And 相关id Is Null) And b.发送时间 Between Trunc(Sysdate) And
          Trunc(Sysdate + 1) - 1 / 24 / 60 / 60 And a.执行时间 Between Trunc(Sysdate) And
          Trunc(Sysdate + 1) - 1 / 24 / 60 / 60 And 操作状态 < 9;
    If n_Sum <> 0 Then
      b_Change  := False;
      v_时间串1 := Trunc(Sysdate + 1) || ';false\';
    
      Update 输液配药记录
      Set 是否调整批次 = 1
      Where 医嘱id In (Select ID From 病人医嘱记录 Where 病人id = n_病人id) And 执行时间 Between Trunc(Sysdate + 1) And
            Trunc(Sysdate + 2) - 1 / 24 / 60 / 60 And 操作状态 < 2;
    End If;
  End If;
  If v_时间串 Is Null Then
    v_时间串 := v_时间串1;
  Else
    v_时间串 := v_时间串 || v_时间串1;
  End If;

  While v_时间串 Is Not Null Loop
    --分解单据ID串
    v_Fields   := Substr(v_时间串, 1, Instr(v_时间串, '\') - 1);
    v_时间值   := Substr(v_Fields, 1, Instr(v_Fields, ';') - 1);
    v_是否改变 := Substr(v_Fields, Instr(v_Fields, ';') + 1);
  
    v_时间串 := Replace('\' || v_时间串, '\' || v_Fields || '\');
  
    If v_是否改变 = 'true' Then
      b_Change := True;
    Else
      b_Change := False;
    End If;

    If b_Change = True Then
      Select Count(医嘱id)
      Into n_Cur
      From (Select Distinct a.要求时间, a.医嘱id
             From 医嘱执行时间 A, 输液配药记录 B
             Where a.要求时间 + 0 = b.执行时间 And a.医嘱id = b.医嘱id And b.执行时间 + 0 Between Trunc(v_时间值) And
                   Trunc(v_时间值 + 1) - 1 / 24 / 60 / 60 And
                   b.医嘱id In (Select ID From 病人医嘱记录 Where 病人id = n_病人id And 相关id Is Null));
    
      Select Count(医嘱id)
      Into n_Sum
      From (Select Distinct a.要求时间, a.医嘱id
             From 医嘱执行时间 A, 输液配药记录 B
             Where a.要求时间 + 0 = b.执行时间 And a.医嘱id = b.医嘱id And b.执行时间 + 0 Between Trunc(v_时间值 - 1) And
                   Trunc(v_时间值) - 1 / 24 / 60 / 60 And
                   b.医嘱id In (Select ID From 病人医嘱记录 Where 病人id = n_病人id And 相关id Is Null));
    
      If n_Cur <> n_Sum Then
        Update 输液配药记录
        Set 是否调整批次 = 1
        Where 医嘱id In (Select ID From 病人医嘱记录 Where 病人id = n_病人id) And 执行时间 Between Trunc(v_时间值) And
              Trunc(v_时间值 + 1) - 1 / 24 / 60 / 60 And 操作状态 < 2;
        b_Change := False;
      End If;
    End If;
  
    If v_保持上次批次 = 1 And b_Change = True Then
      For v_输液单记录 In c_输液单记录 Loop
        Begin
          Select Distinct 配药批次
          Into v_批次
          From 输液配药记录 A, 输液配药内容 B, 药品收发记录 C
          Where a.Id = b.记录id And b.收发id = c.Id And a.医嘱id = v_输液单记录.医嘱id And
                To_Char(a.执行时间, 'hh24:mi:ss') = To_Char(v_输液单记录.执行时间, 'hh24:mi:ss') And
                a.执行时间 Between Trunc(v_输液单记录.执行时间 - 1) And Trunc(v_输液单记录.执行时间) - 1 / 24 / 60 / 60 And Rownum = 1;
        Exception
          When Others Then
            Begin
              Select Distinct 配药批次
              Into v_批次
              From 输液配药记录 A
              Where a.医嘱id = v_输液单记录.医嘱id And To_Char(a.执行时间, 'hh24:mi:ss') = To_Char(v_输液单记录.执行时间, 'hh24:mi:ss') And
                    a.操作状态 <> 12 And a.执行时间 Between Trunc(v_输液单记录.执行时间 - 1) And Trunc(v_输液单记录.执行时间) - 1 / 24 / 60 / 60 And
                    Rownum = 1;
            Exception
              When Others Then
                v_批次 := v_输液单记录.配药批次;
            End;
        End;
      
        Update 输液配药记录
        Set 是否确认调整 = 0, 是否调整批次 = 0
        Where 医嘱id In (Select ID From 病人医嘱记录 Where 病人id = n_病人id) And 执行时间 Between Trunc(v_输液单记录.执行时间) And
              Trunc(v_输液单记录.执行时间 + 1) - 1 / 24 / 60 / 60 And 操作状态 < 2;
      
        If v_输液单记录.配药批次 <> v_批次 Then
          Update 输液配药记录 Set 配药批次 = v_批次 Where ID = v_输液单记录.Id;
          Select Nvl(Max(打包), 0) Into n_打包 From 配药工作批次 Where 批次 = v_批次 And 配置中心id = 部门id_In;
          If n_打包 <> 0 Then
            Update 输液配药记录 Set 是否打包 = n_打包 Where ID = v_输液单记录.Id;
          Else
            Select Nvl(Max(打包), 0)
            Into n_打包
            From 配药工作批次
            Where 批次 = v_输液单记录.配药批次 And 配置中心id = 部门id_In;
          
            If n_打包 <> 0 Then
              Update 输液配药记录 Set 是否打包 = 0 Where ID = v_输液单记录.Id;
            End If;
          End If;
        End If;
      End Loop;
    End If;
  
    If n_自动排批 = 1 And (b_Change = False Or v_保持上次批次 = 0) Then
      For v_输液单记录 In c_输液单记录 Loop
        v_批次 := Zl_Getpivaworkbatch(v_输液单记录.执行时间, v_输液单记录.发送时间);
        Update 输液配药记录 Set 配药批次 = v_批次 Where ID = v_输液单记录.Id;
      End Loop;
      Zl_输液配药记录_自动排批(n_病人id, n_科室id, 部门id_In, v_时间值);
    End If;
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]病人' || v_当前病人 || '在输液配置中心有被锁定的输液单，发送失败！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_输液配药记录_核查;
/

--139595:李业庆,2019-04-15,静配支持门诊费用数据
Create Or Replace Procedure Zl_输液配药记录_配药
(
  配药id_In   In Varchar2, --ID串:ID1,ID2.... 
  操作人员_In In 输液配药记录.操作人员%Type := Null,
  操作时间_In In 输液配药记录.操作时间%Type := Null,
  操作说明_In In 输液配药状态.操作说明%Type := Null,
  移动操作_In In Number := 0
) Is
  v_Tansid     Varchar2(20);
  v_Tmp        Varchar2(4000);
  v_No         Varchar2(20);
  v_Usercode   Varchar2(100);
  n_操作状态   输液配药记录.操作状态%Type;
  v_Error      Varchar2(255);
  n_People     Number(1);
  n_Row        Number(2);
  d_执行时间   Date;
  v_配药类型   Varchar2(50);
  n_项目id     Number(18);
  v_收费项目id Varchar2(200);
  v_Info       Varchar2(200);
  v_Id         Varchar2(20);
  n_数次       Number(2);
  n_Count      Number(18);
  n_Out        Number(10);
  n_Outnum     Number(10);
  n_打包状态   Number(1);
  v_核对人     Varchar2(20);
  v_配液人     Varchar2(20);
  Err_Custom Exception;

  Cursor c_Bill Is
    Select a.病人id, a.主页id, a.标识号, a.姓名, a.性别, a.年龄, a.床号, a.费别, a.病人病区id, a.病人科室id, a.婴儿费, e.药品id, b.库房id, f.配药类型, g.序号,
           1 As 门诊标志
    From 住院费用记录 A, 药品收发记录 B, 输液配药记录 C, 输液配药内容 D, 药品规格 E, 输液药品属性 F, 配置收费方案 G
    Where a.Id = b.费用id And b.Id = d.收发id And d.记录id = c.Id And b.药品id = e.药品id And b.药品id = f.药品id And
          Substr(f.配药类型, Instr(f.配药类型, '-') + 1) = g.配药类型(+) And Nvl(c.是否打包, 0) <> 1 And c.Id = v_Tansid
    Union All
    Select a.病人id, a.主页id, a.标识号, a.姓名, a.性别, a.年龄, '' As 床号, a.费别, a.病人病区id, a.病人科室id, a.婴儿费, e.药品id, b.库房id, f.配药类型,
           g.序号, 0 As 门诊标志
    From 门诊费用记录 A, 药品收发记录 B, 输液配药记录 C, 输液配药内容 D, 药品规格 E, 输液药品属性 F, 配置收费方案 G
    Where a.Id = b.费用id And b.Id = d.收发id And d.记录id = c.Id And b.药品id = e.药品id And b.药品id = f.药品id And
          Substr(f.配药类型, Instr(f.配药类型, '-') + 1) = g.配药类型(+) And Nvl(c.是否打包, 0) <> 1 And c.Id = v_Tansid
    Order By 序号;

Begin
  v_Usercode := Zl_Identity;
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ';') + 1);
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ',') + 1);
  v_Usercode := Substr(v_Usercode, 1, Instr(v_Usercode, ',') - 1);
  n_People   := Nvl(zl_GetSysParameter('配置费按病人收取', 1345), 0);
  n_Out      := Nvl(zl_GetSysParameter('出院病人不收配置费', 1345), 0);

  If 配药id_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := 配药id_In || ',';
  End If;

  While v_Tmp Is Not Null Loop
    --分解单据ID串 
    v_Tansid := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp    := Replace(',' || v_Tmp, ',' || v_Tansid || ',');
  
    --检查当前输液单的状态是否为待摆药状态 
    Begin
      Select 操作状态, 执行时间, Nvl(是否打包, 0)
      Into n_操作状态, d_执行时间, n_打包状态
      From 输液配药记录
      Where ID = v_Tansid;
    
      If n_操作状态 > 3 Then
        v_Error := '该数据已被操作，不能进行发药！';
        Raise Err_Custom;
      End If;
    Exception
      When Others Then
        v_Error := '该数据已被操作！';
        Raise Err_Custom;
    End;
  
    v_核对人 := '';
    v_配液人 := '';
    Begin
      Select 核对人, 配液人
      Into v_核对人, v_配液人
      From (Select e.核对人, e.配液人
             From 输液配药记录 A, 输液配药内容 B, 药品收发记录 C, 配液台药品对照 D, 配液工作安排 E
             Where a.Id = b.记录id And b.收发id = c.Id And c.药品id = d.药品id And c.库房id = d.部门id And d.配药台id = e.配药台id And
                   a.配药批次 = e.批次 And e.日期 = To_Date(To_Char(Sysdate, 'yyyy-mm-dd'), 'yyyy-mm-dd') And a.Id = v_Tansid And
                   Rownum = 1
             Order By d.配药台id)
      Where Rownum = 1;
    Exception
      When Others Then
        Null;
    End;
  
    Update 输液配药记录 Set 操作状态 = 4, 操作人员 = 操作人员_In, 操作时间 = 操作时间_In Where ID = v_Tansid;
    If 移动操作_In = 0 Then
      Insert Into 输液配药状态
        (配药id, 操作类型, 操作人员, 操作时间, 操作说明, 实际工作人员)
      Values
        (v_Tansid, 3, 操作人员_In, 操作时间_In, 操作说明_In, v_核对人);
    End If;
    Insert Into 输液配药状态
      (配药id, 操作类型, 操作人员, 操作时间, 操作说明, 实际工作人员)
    Values
      (v_Tansid, 4, 操作人员_In, 操作时间_In, 操作说明_In, v_配液人);
  
    If n_打包状态 = 0 Then
      n_Count := 0;
      Select Nextno(14) Into v_No From Dual;
      For r_Bill In c_Bill Loop
        Select Count(病人id)
        Into n_Outnum
        From 病案主页
        Where 主页id = r_Bill.主页id And 病人id = r_Bill.病人id And (Nvl(状态, 0) = 3 Or 出院日期 Is Not Null);
        If n_Count = 0 And (n_Outnum = 0 Or n_Out = 0) Then
          --收取材料费 
          --v_收费项目id:='6970,2;6971,1;'; 
          Select Zl_Fun_Pivacustom(v_Tansid) Into v_收费项目id From Dual;
          While v_收费项目id Is Not Null Loop
            v_Info       := Substr(v_收费项目id, 1, Instr(v_收费项目id, ';') - 1);
            v_收费项目id := Replace(';' || v_收费项目id, ';' || v_Info || ';');
          
            v_Id   := Substr(v_Info, 1, Instr(v_Info, ',') - 1);
            v_Info := Replace(',' || v_Info, ',' || v_Id || ',');
          
            For r_Item In (Select a.Id 收费细目id, a.类别 收费类别, a.计算单位, a.加班加价 加班标志, d.Id 收入项目id, d.收据费目, b.现价
                           From 收费项目目录 A, 收费价目 B, 收入项目 D
                           Where a.Id = b.收费细目id And b.收入项目id = d.Id And a.Id = v_Id And b.执行日期 <= Sysdate And
                                 (b.终止日期 >= Sysdate Or b.终止日期 Is Null)) Loop
              If n_Count = 0 Then
                Insert Into 输液配药附费 (配药id, NO, 病人id) Values (v_Tansid, v_No, r_Bill.病人id);
              End If;
            
              n_Count := n_Count + 1;
            
              If r_Bill.门诊标志 = 1 Then
                Zl_住院记帐记录_Insert(v_No, n_Count, r_Bill.病人id, r_Bill.主页id, r_Bill.标识号, r_Bill.姓名, r_Bill.性别, r_Bill.年龄,
                                 r_Bill.床号, r_Bill.费别, r_Bill.病人病区id, r_Bill.病人科室id, r_Item.加班标志, r_Bill.婴儿费,
                                 r_Bill.库房id, 操作人员_In, Null, r_Item.收费细目id, r_Item.收费类别, r_Item.计算单位, Null, Null, Null,
                                 1, v_Info, Null, r_Bill.库房id, Null, r_Item.收入项目id, r_Item.收据费目, r_Item.现价,
                                 r_Item.现价 * v_Info, r_Item.现价 * v_Info, Null, Sysdate, Sysdate, Null, Null, v_Usercode,
                                 操作人员_In);
              Else
                Zl_门诊记帐记录_Insert(v_No, n_Count, r_Bill.病人id, r_Bill.标识号, r_Bill.姓名, r_Bill.性别, r_Bill.年龄, r_Bill.费别,
                                 r_Item.加班标志, r_Bill.婴儿费, r_Bill.病人科室id, r_Bill.库房id, 操作人员_In, Null, r_Item.收费细目id,
                                 r_Item.收费类别, r_Item.计算单位, 1, 1, Null, r_Bill.库房id, Null, r_Item.收入项目id, r_Item.收据费目,
                                 r_Item.现价, r_Item.现价, r_Item.现价, Sysdate, Sysdate, Null, Null, v_Usercode, 操作人员_In);
              End If;
            End Loop;
          End Loop;
        End If;
      
        --先查询是否有按给药途径收取的配置费方案
        Select Nvl(Max(项目id), 0)
        Into n_项目id
        From 输液配药记录 A, 病人医嘱记录 B, 配置收费方案 C
        Where a.Id = v_Tansid And a.医嘱id = b.Id And b.诊疗项目id = c.诊疗id;
        If n_项目id = 0 Then
          --若无对应给药途径的配置费收取方案，则再查询是否有按配药类型收取的配置费方案
          Select Nvl(Max(项目id), 0)
          Into n_项目id
          From 配置收费方案
          Where 配药类型 = Substr(r_Bill.配药类型, Instr(r_Bill.配药类型, '-', 1, 1) + 1);
        End If;
      
        If n_项目id <> 0 Then
          n_Row := 0;
        
          If n_People = 1 Then
            Select Count(配药id)
            Into n_Row
            From (Select 配药id
                   From 输液配药附费 A, 住院费用记录 B, 输液配药记录 C
                   Where a.No = b.No And a.配药id = c.Id And b.病人id = r_Bill.病人id And b.记录状态 = 1 And b.收费细目id = n_项目id And
                         d_执行时间 Between Trunc(c.执行时间) And Trunc(c.执行时间 + 1) - 1 / 24 / 60 / 60
                   Union All
                   Select 配药id
                   From 输液配药附费 A, 门诊费用记录 B, 输液配药记录 C
                   Where a.No = b.No And a.配药id = c.Id And b.病人id = r_Bill.病人id And b.记录状态 = 1 And b.收费细目id = n_项目id And
                         d_执行时间 Between Trunc(c.执行时间) And Trunc(c.执行时间 + 1) - 1 / 24 / 60 / 60);
          End If;
        Else
          n_Row := 1;
        End If;
      
        If n_Row = 0 And (n_Outnum = 0 Or n_Out = 0) Then
          For r_Item In (Select a.Id 收费细目id, a.类别 收费类别, a.计算单位, a.加班加价 加班标志, d.Id 收入项目id, d.收据费目, b.现价
                         From 收费项目目录 A, 收费价目 B, 收入项目 D
                         Where a.Id = b.收费细目id And b.收入项目id = d.Id And a.Id = n_项目id And b.执行日期 <= Sysdate And
                               (b.终止日期 >= Sysdate Or b.终止日期 Is Null)) Loop
            If n_Count = 0 Then
              Insert Into 输液配药附费 (配药id, NO, 病人id) Values (v_Tansid, v_No, r_Bill.病人id);
            End If;
          
            n_Count := n_Count + 1;
          
            If r_Bill.门诊标志 = 1 Then
              Zl_住院记帐记录_Insert(v_No, n_Count, r_Bill.病人id, r_Bill.主页id, r_Bill.标识号, r_Bill.姓名, r_Bill.性别, r_Bill.年龄,
                               r_Bill.床号, r_Bill.费别, r_Bill.病人病区id, r_Bill.病人科室id, r_Item.加班标志, r_Bill.婴儿费, r_Bill.库房id,
                               操作人员_In, Null, r_Item.收费细目id, r_Item.收费类别, r_Item.计算单位, Null, Null, Null, 1, 1, Null,
                               r_Bill.库房id, Null, r_Item.收入项目id, r_Item.收据费目, r_Item.现价, r_Item.现价, r_Item.现价, Null,
                               Sysdate, Sysdate, Null, Null, v_Usercode, 操作人员_In);
            Else
              Zl_门诊记帐记录_Insert(v_No, n_Count, r_Bill.病人id, r_Bill.标识号, r_Bill.姓名, r_Bill.性别, r_Bill.年龄, r_Bill.费别,
                               r_Item.加班标志, r_Bill.婴儿费, r_Bill.病人科室id, r_Bill.库房id, 操作人员_In, Null, r_Item.收费细目id,
                               r_Item.收费类别, r_Item.计算单位, 1, 1, Null, r_Bill.库房id, Null, r_Item.收入项目id, r_Item.收据费目,
                               r_Item.现价, r_Item.现价, r_Item.现价, Sysdate, Sysdate, Null, Null, v_Usercode, 操作人员_In);
            End If;
          End Loop;
        End If;
      
        If n_Row = 0 Then
          Exit;
        End If;
      End Loop;
    End If;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_输液配药记录_配药;
/

--139595:李业庆,2019-04-15,静配支持门诊费用数据
Create Or Replace Procedure Zl_住院记帐记录_发药审核
(
  Billid_In     In Varchar2, --药品收发记录ID串,格式:"id1,id2,id3....."
  操作员编号_In 住院费用记录.操作员编号%Type,
  操作员姓名_In 住院费用记录.操作员姓名%Type,
  审核时间_In   住院费用记录.登记时间%Type := Null
) As
  --功能：审核一张住院记帐划价单
  --139595修改支持审核门诊费用
  --参数：
  --    审核时间_IN：用于部份需要统一控制或返回时间的地方

  n_Billid Number;
  n_类别   Number;
  Cursor c_Bill Is
    Select ID, NO, 记录性质, 记录状态, 病人id, 主页id, 门诊标志, 实收金额, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 病人病区id, 1 As 住院费用
    From 住院费用记录
    Where ID = (Select 费用id From 药品收发记录 Where ID = n_Billid) And 记录状态 = 0
    Union All
    Select ID, NO, 记录性质, 记录状态, 病人id, 主页id, 门诊标志, 实收金额, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 病人病区id, 0 As 住院费用
    From 门诊费用记录
    Where ID = (Select 费用id From 药品收发记录 Where ID = n_Billid) And 记录状态 = 0;

  v_Infotmp Varchar2(4000);
  v_Date    Date;
Begin
  If 审核时间_In Is Null Then
    Select Sysdate Into v_Date From Dual;
  Else
    v_Date := 审核时间_In;
  End If;

  v_Infotmp := Billid_In || ',';
  While v_Infotmp Is Not Null Loop
    --分解单据ID串
    n_Billid  := Substr(v_Infotmp, 1, Instr(v_Infotmp, ',') - 1);
    v_Infotmp := Replace(',' || v_Infotmp, ',' || n_Billid || ',');
  
    For r_Bill In c_Bill Loop
      If r_Bill.住院费用 = 1 Then
        Update 住院费用记录
        Set 记录状态 = 1, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In, 登记时间 = v_Date --已产生的药品记录的时间不变
        Where ID = r_Bill.Id;
      Else
        Update 门诊费用记录
        Set 记录状态 = 1, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In, 登记时间 = v_Date --已产生的药品记录的时间不变
        Where ID = r_Bill.Id;
      End If;
    
      --药品收发记录.填制日期
      Update 药品收发记录
      Set 填制日期 = Decode(Sign(Nvl(审核日期, v_Date) - v_Date), -1, 填制日期, v_Date)
      Where ID = n_Billid;
    
      If Nvl(r_Bill.门诊标志, 0) = 1 Or Nvl(r_Bill.门诊标志, 0) = 2 Then
        n_类别 := r_Bill.门诊标志;
      Elsif Nvl(r_Bill.主页id, 0) = 0 Or Nvl(r_Bill.门诊标志, 0) = 4 Then
        n_类别 := 1;
      Else
        n_类别 := 2;
      End If;
    
      --病人余额
      Update 病人余额
      Set 费用余额 = Nvl(费用余额, 0) + Nvl(r_Bill.实收金额, 0)
      Where 病人id = r_Bill.病人id And 性质 = 1 And 类型 = n_类别;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 性质, 类型, 费用余额, 预交余额)
        Values
          (r_Bill.病人id, 1, n_类别, r_Bill.实收金额, 0);
      End If;
    
      --病人未结费用
      Update 病人未结费用
      Set 金额 = Nvl(金额, 0) + Nvl(r_Bill.实收金额, 0)
      Where 病人id = r_Bill.病人id And Nvl(主页id, 0) = Nvl(r_Bill.主页id, 0) And Nvl(病人病区id, 0) = Nvl(r_Bill.病人病区id, 0) And
            Nvl(病人科室id, 0) = Nvl(r_Bill.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(r_Bill.开单部门id, 0) And
            Nvl(执行部门id, 0) = Nvl(r_Bill.执行部门id, 0) And 收入项目id + 0 = r_Bill.收入项目id And 来源途径 + 0 = r_Bill.门诊标志;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人未结费用
          (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
        Values
          (r_Bill.病人id, r_Bill.主页id, r_Bill.病人病区id, r_Bill.病人科室id, r_Bill.开单部门id, r_Bill.执行部门id, r_Bill.收入项目id,
           r_Bill.门诊标志, Nvl(r_Bill.实收金额, 0));
      End If;
    
      --库房中的药品已全部审核则标为已收费
      If r_Bill.住院费用 = 1 Then
        Update 未发药品记录
        Set 已收费 = 1, 填制日期 = v_Date
        Where NO = r_Bill.No And 单据 In (9, 10) And Nvl(已收费, 0) = 0 And
              Nvl(库房id, 0) Not In
              (Select Distinct Nvl(执行部门id, 0)
               From 住院费用记录
               Where 记录性质 = 2 And NO = r_Bill.No And 收费类别 In ('5', '6', '7') And 记录状态 = 0);
      Else
        Update 未发药品记录
        Set 已收费 = 1, 填制日期 = v_Date
        Where NO = r_Bill.No And 单据 In (9, 10) And Nvl(已收费, 0) = 0 And
              Nvl(库房id, 0) Not In
              (Select Distinct Nvl(执行部门id, 0)
               From 门诊费用记录
               Where 记录性质 = 2 And NO = r_Bill.No And 收费类别 In ('5', '6', '7') And 记录状态 = 0);
      End If;
    End Loop;
  End Loop;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_住院记帐记录_发药审核;
/

--139595:李业庆,2019-04-15,静配支持门诊费用数据
Create Or Replace Procedure Zl_输液配药记录_取消发送(配药id_In In Varchar2 --ID串:ID1,ID2....
                                           ) Is
  v_Tansid   Varchar2(20);
  v_Tmp      Varchar2(4000);
  v_操作人员 输液配药记录.操作人员%Type;
  d_操作时间 输液配药记录.操作时间%Type;
  n_打包     Number(2);

  v_Error    Varchar2(255);
  n_操作状态 输液配药记录.操作状态%Type;
  v_Usercode Varchar2(100);
  Err_Custom Exception;
Begin
  v_Usercode := Zl_Identity;
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ';') + 1);
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ',') + 1);
  v_Usercode := Substr(v_Usercode, 1, Instr(v_Usercode, ',') - 1);

  If 配药id_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := 配药id_In || ',';
  End If;
  While v_Tmp Is Not Null Loop
    --分解单据ID串
    v_Tansid := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp    := Replace(',' || v_Tmp, ',' || v_Tansid || ',');
  
    --检查当前输液单的状态是否为已发送状态
    Begin
      Select 操作状态 Into n_操作状态 From 输液配药记录 Where ID = v_Tansid;
    
      If n_操作状态 != 5 Then
        v_Error := '该数据已被操作，不能进行取消发送操作！';
        Raise Err_Custom;
      End If;
    Exception
      When Others Then
        v_Error := '该数据已被操作！';
        Raise Err_Custom;
    End;
  
    Select 操作人员, 操作时间
    Into v_操作人员, d_操作时间
    From (Select 操作人员, 操作时间 From 输液配药状态 Where 配药id = v_Tansid And 操作类型 = 4 Order By 操作时间 Desc)
    Where Rownum = 1;
  
    Update 输液配药记录 Set 操作状态 = 4, 操作人员 = v_操作人员, 操作时间 = d_操作时间 Where ID = v_Tansid;
  
    --向[输液配药状态]表中记录“取消发送”的操作
    Insert Into 输液配药状态
      (配药id, 操作类型, 操作人员, 操作时间, 操作说明)
    Values
      (v_Tansid, 4, v_操作人员, Sysdate, '取消发送');
  
    Select 是否打包 Into n_打包 From 输液配药记录 Where ID = v_Tansid;
    If n_打包 <> 0 Then
      For r_Item In (Select a.No, b.序号, 1 As 住院费用
                     From 输液配药附费 A, 住院费用记录 B
                     Where a.病人id = b.病人id And a.No = b.No And b.记录状态 = 1 And a.配药id = v_Tansid
                     Union All
                     Select a.No, b.序号, 0 As 住院费用
                     From 输液配药附费 A, 门诊费用记录 B
                     Where a.病人id = b.病人id And a.No = b.No And b.记录状态 = 1 And a.配药id = v_Tansid) Loop
        If r_Item.No Is Not Null Then
          If r_Item.住院费用 = 1 Then
            Zl_住院记帐记录_Delete(r_Item.No, r_Item.序号, v_Usercode, Zl_Username);
          Else
            Zl_门诊记帐记录_Delete(r_Item.No, r_Item.序号, v_Usercode, Zl_Username);
          End If;
        End If;
      End Loop;
    End If;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_输液配药记录_取消发送;
/

--139595:李业庆,2019-04-15,静配支持门诊费用数据
Create Or Replace Procedure Zl_输液配药记录_取消配药(配药id_In In varchar2 --ID串:ID1,ID2....
                                           ) Is
  v_Tansid   varchar2(20);
  v_Tmp      varchar2(4000);
  v_No       varchar2(20);
  v_Usercode varchar2(100);
  n_打包     输液配药记录.是否打包%Type := 0;
  v_操作人员 输液配药记录.操作人员%Type;
  d_操作时间 输液配药记录.操作时间%Type;

  v_Error    varchar2(255);
  n_操作状态 输液配药记录.操作状态%Type;
  Err_Custom Exception;
  n_Row Number(10);
  n_Out Number(1);

Begin
  v_Usercode := Zl_Identity;
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ';') + 1);
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ',') + 1);
  v_Usercode := Substr(v_Usercode, 1, Instr(v_Usercode, ',') - 1);
  n_Out      := Nvl(zl_GetSysParameter('出院病人不收配置费', 1345), 0);

  If 配药id_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := 配药id_In || ',';
  End If;

  While v_Tmp Is Not Null Loop
    --分解单据ID串 
    v_Tansid := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp    := Replace(',' || v_Tmp, ',' || v_Tansid || ',');
  
    --检查当前输液单的状态是否为待摆药状态 
    Begin
      Select 操作状态 Into n_操作状态 From 输液配药记录 Where ID = v_Tansid;
    
      If n_操作状态 != 4 Then
        v_Error := '该数据当前不是配药状态，不能进行取消配药！';
        Raise Err_Custom;
      End If;
    Exception
      When Others Then
        v_Error := '该数据已被操作！';
        Raise Err_Custom;
    End;
  
    Select 操作人员, 操作时间
    Into v_操作人员, d_操作时间
    From (Select 操作人员, 操作时间 From 输液配药状态 Where 配药id = v_Tansid And 操作类型 = 2 Order By 操作时间 Desc)
    Where Rownum = 1;
  
    Update 输液配药记录 Set 操作状态 = 2, 操作人员 = v_操作人员, 操作时间 = d_操作时间 Where ID = v_Tansid;
  
    --向[输液配药状态]表中记录“取消配药”的操作
    Insert Into 输液配药状态
      (配药id, 操作类型, 操作人员, 操作时间, 操作说明)
    Values
      (v_Tansid, 2, v_操作人员, Sysdate, '取消配药');
  
    Select 是否打包 Into n_打包 From 输液配药记录 Where ID = v_Tansid;
    If n_打包 <> 1 Then
      For r_Item In (Select a.No, b.序号, 1 As 住院费用
                     From 输液配药附费 A, 住院费用记录 B
                     Where a.病人id = b.病人id And a.No = b.No And b.记录状态 = 1 And a.配药id = v_Tansid
                     Union All
                     Select a.No, b.序号, 0 As 住院费用
                     From 输液配药附费 A, 门诊费用记录 B
                     Where a.病人id = b.病人id And a.No = b.No And b.记录状态 = 1 And a.配药id = v_Tansid) Loop
        If r_Item.No Is Not Null Then
          If r_Item.住院费用 = 1 Then
             Zl_住院记帐记录_Delete(r_Item.No, r_Item.序号, v_Usercode, Zl_Username);
          Else
            Zl_门诊记帐记录_Delete(r_Item.No, r_Item.序号, v_Usercode, Zl_Username);
          End If;
        End If;
      End Loop;
    Else
      Zl_输液配药记录_取消摆药(v_Tansid);
    End If;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_输液配药记录_取消配药;
/

--139595:李业庆,2019-04-15,静配支持门诊费用数据
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
    Distinct c.记录id, a.Id As 退药id, c.收发id, a.批号, a.效期, a.产地, c.数量 As 退药数, a.药品id, a.批次, b.费用id
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
      Begin
        Select 1 Into n_门诊 From 门诊费用记录 Where ID = v_配药内容.费用id;
      Exception
        When Others Then
          n_门诊 := 2;
      End;
    
      --处理退药
      Zl_药品收发记录_部门退药(v_配药内容.退药id, Zl_Username, v_Date, v_配药内容.批号, v_配药内容.效期, v_配药内容.产地, v_配药内容.退药数, Null, Zl_Username,
                     2, n_门诊);
    
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

--139595:李业庆,2019-04-15,静配支持门诊费用数据
Create Or Replace Procedure Zl_输液配药记录_确认拒绝
(
  配药id_In   In 输液配药记录.Id%Type,
  操作人员_In In 输液配药记录.操作人员%Type
) Is
  n_配药id   输液配药记录.Id%Type;
  n_发药id   药品收发记录.Id%Type;
  n_退药id   药品收发记录.Id%Type;
  n_收发id   药品收发记录.Id%Type;
  n_操作状态 输液配药记录.操作状态%Type;
  v_操作人员 输液配药记录.操作人员%Type;
  d_操作时间 输液配药记录.操作时间%Type;
  v_Error    Varchar2(255);
  Err_Custom Exception;
  n_门诊 Number; --1：门诊单据；2：住院单据

  Cursor c_配药内容 Is
    Select /*+ rule*/
    Distinct c.记录id, a.Id As 退药id, c.收发id, a.费用id
    From 药品收发记录 A, 药品收发记录 B, 输液配药内容 C
    Where c.记录id = 配药id_In And b.Id = c.收发id And a.单据 = b.单据 And a.No = b.No And a.库房id + 0 = b.库房id And
          a.药品id + 0 = b.药品id And a.序号 = b.序号 And (a.记录状态 = 1 Or Mod(a.记录状态, 3) = 0);

  r_配药内容 c_配药内容%RowType;

  Cursor c_退药记录 Is
    Select a.Id, a.批号, a.效期, a.产地, b.数量 As 退药数
    From 药品收发记录 A, 输液配药内容 B
    Where a.Id = n_退药id And a.审核人 Is Not Null And b.收发id = n_收发id And b.记录id = 配药id_In;

  r_退药记录 c_退药记录%RowType;

Begin
  Select Nvl(操作状态, 0) Into n_操作状态 From 输液配药记录 Where ID = 配药id_In;

  If n_操作状态 = 7 Then
    Insert Into 输液配药状态 (配药id, 操作类型, 操作人员, 操作时间) Values (配药id_In, 8, 操作人员_In, Sysdate);
  
    Select 操作人员, 操作时间
    Into v_操作人员, d_操作时间
    From 输液配药状态
    Where 配药id = 配药id_In And 操作类型 = 1
    Order By 操作时间 Desc;
  
    Update 输液配药记录 Set 操作状态 = 1, 操作人员 = v_操作人员, 操作时间 = d_操作时间 Where ID = 配药id_In;
  
    For r_配药内容 In c_配药内容 Loop
      n_退药id := r_配药内容.退药id;
      n_收发id := r_配药内容.收发id;
      For r_退药记录 In c_退药记录 Loop
        Begin
          Select 1 Into n_门诊 From 门诊费用记录 Where ID = r_配药内容.费用id;
        Exception
          When Others Then
            n_门诊 := 2;
        End;
      
        --处理退药
        Zl_药品收发记录_部门退药(r_退药记录.Id, Zl_Username, Sysdate, r_退药记录.批号, r_退药记录.效期, r_退药记录.产地, r_退药记录.退药数, Null, Zl_Username);
      
        Select Max(a.Id)
        Into n_发药id
        From 药品收发记录 A, 药品收发记录 B
        Where b.Id = n_退药id And a.单据 = b.单据 And a.No = b.No And a.库房id + 0 = b.库房id And a.药品id + 0 = b.药品id And
              a.序号 = b.序号 And Mod(a.记录状态, 3) = 1 And a.审核日期 Is Null;
      
        --替换输液配药内容中的收发ID
        Update 输液配药内容 Set 收发id = n_发药id Where 记录id = 配药id_In And 收发id = r_配药内容.收发id;
      End Loop;
    End Loop;
  Else
    v_Error := '已有其他用户将该输液单操作，不能重复操作！';
    Raise Err_Custom;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_输液配药记录_确认拒绝;
/

--139595:李业庆,2019-04-15,静配支持门诊费用数据
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

  v_原始id 药品收发记录.Id%Type;
  v_Error  Varchar2(255);
  n_门诊   Number; --1：门诊单据；2：住院单据
  Err_Custom Exception;

  Cursor c_销帐记录 Is
    Select Distinct a.费用id, b.操作时间
    From 药品收发记录 A, 输液配药记录 B, 输液配药内容 C
    Where a.Id = c.收发id And b.Id = c.记录id And b.Id = v_Tansid And b.操作状态 = 9;

  v_销帐记录 c_销帐记录%RowType;

  Cursor c_退药记录 Is
    Select /*+ rule*/
    Distinct a.Id As 退药id, c.收发id, c.数量, a.药品id, a.批次, c.记录id As 配药id, a.费用id
    From 药品收发记录 A, 药品收发记录 B, 输液配药内容 C, Table(Cast(f_Num2list(v_Tansids) As Zltools.t_Numlist)) D
    Where c.记录id = d.Column_Value And b.Id = c.收发id And a.单据 = b.单据 And a.No = b.No And a.库房id + 0 = b.库房id And
          a.药品id + 0 = b.药品id And a.序号 = b.序号 And (a.记录状态 = 1 Or Mod(a.记录状态, 3) = 0)
    Order By a.药品id, a.批次;

  v_退药记录 c_退药记录%RowType;

  Cursor c_费用销帐 Is
    Select /*+ rule*/
     a.No, a.序号 || ':' || c.数量 || ':' || c.记录id As 费用序号, 1 As 住院费用
    From 住院费用记录 A, 药品收发记录 B, 输液配药内容 C, Table(Cast(f_Num2list(v_Tansids) As Zltools.t_Numlist)) D
    Where a.Id = b.费用id And b.Id = c.收发id And Mod(b.记录状态, 3) = 1 And c.记录id = d.Column_Value
    Union All
    Select /*+ rule*/
     a.No, a.序号 || ':' || c.数量 || ':' || c.记录id As 费用序号, 0 As 住院费用
    From 门诊费用记录 A, 药品收发记录 B, 输液配药内容 C, Table(Cast(f_Num2list(v_Tansids) As Zltools.t_Numlist)) D
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
    --判断这个单据是门诊还是住院 
    Begin
      Select 1 Into n_门诊 From 门诊费用记录 Where ID = v_退药记录.费用id;
    Exception
      When Others Then
        n_门诊 := 2;
    End;
  
    Zl_药品收发记录_部门退药(v_退药记录.退药id, 操作人员_In, 操作时间_In, Null, Null, Null, v_退药记录.数量, Null, 操作人员_In, 2, n_门诊);
  
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
    If v_费用销帐.住院费用 = 1 Then
      Zl_住院记帐记录_Delete(v_费用销帐.No, v_费用销帐.费用序号, v_Usercode, Zl_Username, 2, 1, 1, d_审核时间);
    Else
      Zl_门诊记帐记录_Delete(v_费用销帐.No, v_费用销帐.费用序号, v_Usercode, Zl_Username);
    End If;
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

--139595:李业庆,2019-04-15,静配支持门诊费用数据
Create Or Replace Procedure Zl_药品收发记录_批量发药
(
  Billinfo_In   In Varchar2, --格式:"id1,批次1|id2,批次2|....."
  Partid_In     In 药品收发记录.库房id%Type,
  People_In     In 药品收发记录.审核人%Type,
  Date_In       In 药品收发记录.审核日期%Type,
  发药方式_In   In 药品收发记录.发药方式%Type := 3,
  领药人_In     In 药品收发记录.领用人%Type := Null,
  汇总发药号_In In 药品收发记录.汇总发药号%Type := Null,
  Intdigit_In   In Number := 2,
  配药人_In     In 药品收发记录.配药人%Type := Null,
  核查人_In     In 药品收发记录.核查人%Type := Null
) Is
  --只读变量
  v_Infotmp     Varchar2(4000);
  v_Fields      Varchar2(4000);
  n_Billid      药品收发记录.Id%Type;
  n_批次        药品收发记录.批次%Type;
  Lng入出类别id Number(18);
  Int入出系数   Number;
  Int执行状态   Number;
  Int单据       药品收发记录.单据%Type;
  Strno         药品收发记录.No%Type;
  Lng库房id     药品收发记录.库房id%Type;
  Lng药品id     药品收发记录.药品id%Type;
  Lng费用id     药品收发记录.费用id%Type;
  Dbl差价率     Number;
  v_零售价      药品收发记录.零售价%Type;
  Int未发数     未发药品记录.未发数%Type;
  v_核查日期    药品收发记录.核查日期%Type;
  --可写变量
  Dbl实际数量 药品收发记录.实际数量%Type;
  Dbl实际金额 药品收发记录.零售金额%Type;
  Dbl成本金额 药品收发记录.成本金额%Type;
  Dbl实际差价 药品收发记录.差价%Type;
  --2002-07-31朱玉宝
  --LNGLAST批次 发药前确定的批次(已减可用数量)
  Str药名           Varchar2(200);
  Dbl可用数量       药品收发记录.填写数量%Type;
  Lnglast批次       药品收发记录.批次%Type;
  Lngcur批次        药品收发记录.批次%Type;
  Str批号           药品收发记录.批号%Type;
  Str效期           药品收发记录.效期%Type;
  n_上次供应商id    药品库存.上次供应商id%Type;
  n_上次采购价      药品库存.上次采购价%Type;
  v_上次产地        药品库存.上次产地%Type;
  d_上次生产日期    药品库存.上次生产日期%Type;
  v_批准文号        药品库存.批准文号%Type;
  n_记录状态        药品收发记录.记录状态%Type;
  n_平均成本价      药品库存.平均成本价%Type;
  n_发药方式        药品收发记录.发药方式%Type;
  v_摘要            药品收发记录.摘要%Type;
  Bln收费与发药分离 Number(1);
  v_Error           Varchar2(255);
  Err_Custom Exception;
  n_流通售价小数   Number;
  n_流通金额小数   Number;
  n_零差价管理模式 Number(1);
  n_药品零差价管理 Number(1);
  n_处方类型       未发药品记录.处方类型%Type;
  n_住院费用       Number(1);
Begin
  Select Sysdate Into v_核查日期 From Dual;
  If Billinfo_In Is Null Then
    v_Infotmp := Null;
  Else
    v_Infotmp := Billinfo_In || '|';
  End If;
  While v_Infotmp Is Not Null Loop
    --分解单据ID串
    v_Fields  := Substr(v_Infotmp, 1, Instr(v_Infotmp, '|') - 1);
    n_Billid  := Substr(v_Fields, 1, Instr(v_Fields, ',') - 1);
    n_批次    := Substr(v_Fields, Instr(v_Fields, ',') + 1);
    v_Infotmp := Replace('|' || v_Infotmp, '|' || v_Fields || '|');
  
    --获取该收发记录的单据、药品ID、库房ID,零售金额及实际数量、入出类别ID
    Begin
      Select a.单据, a.No, a.药品id, a.库房id, a.费用id, Nvl(a.零售价, 0), Nvl(a.零售金额, 0), Nvl(a.实际数量, 0) * Nvl(a.付数, 1), a.入出类别id,
             a.入出系数, Nvl(a.批次, 0), a.批号, a.效期, a.供药单位id, a.产地, a.生产日期, a.批准文号, Nvl(a.发药方式, 0), a.摘要, 记录状态
      Into Int单据, Strno, Lng药品id, Lng库房id, Lng费用id, v_零售价, Dbl实际金额, Dbl实际数量, Lng入出类别id, Int入出系数, Lnglast批次, Str批号, Str效期,
           n_上次供应商id, v_上次产地, d_上次生产日期, v_批准文号, n_发药方式, v_摘要, n_记录状态
      From 药品收发记录 A
      Where a.Id = n_Billid And a.审核日期 Is Null
      For Update Nowait;
    
      Select '[' || c.编码 || ']' || c.名称 Into Str药名 From 收费项目目录 C Where c.Id = Lng药品id;
    Exception
      When Others Then
        Int单据 := 0;
        v_Error := '已有其他用户在执行发药，不能重复操作！';
        Raise Err_Custom;
    End;
  
    --取流通业务精度位数
    --类别:1-药品 2-卫材
    --内容：2-零售价 4-金额
    --单位：药品:1-售价 5-金额单位
    Begin
      Select 精度 Into n_流通金额小数 From 药品卫材精度 Where 类别 = 1 And 内容 = 4 And 单位 = 5;
    Exception
      When Others Then
        n_流通金额小数 := 2;
    End;
  
    Begin
      Select 精度 Into n_流通售价小数 From 药品卫材精度 Where 类别 = 1 And 内容 = 2 And 单位 = 1;
    Exception
      When Others Then
        n_流通售价小数 := 2;
    End;
  
    Select Zl_To_Number(Nvl(zl_GetSysParameter(275), '0')) Into n_零差价管理模式 From Dual;
  
    If n_发药方式 = -1 Or v_摘要 = '拒发' Then
      Int单据 := 0;
    End If;
  
    If Int单据 > 0 Then
      If Nvl(n_批次, 0) = 0 Then
        Lngcur批次 := Lnglast批次;
      Else
        Lngcur批次 := Nvl(n_批次, 0);
      End If;
    
      --检查是否已经填写库房
      Bln收费与发药分离 := 0;
      If Lng库房id Is Null Then
        Bln收费与发药分离 := 1;
      End If;
      Lng库房id := Partid_In;
    
      --取该批药品的批号
      Begin
        Select 上次批号, 效期, Nvl(可用数量, 0), 上次供应商id, 上次产地, 上次生产日期, 批准文号, 上次采购价
        Into Str批号, Str效期, Dbl可用数量, n_上次供应商id, v_上次产地, d_上次生产日期, v_批准文号, n_上次采购价
        From 药品库存
        Where 库房id + 0 = Lng库房id And 药品id = Lng药品id And 性质 = 1 And Nvl(批次, 0) = Lngcur批次;
      Exception
        When Others Then
          n_上次采购价 := 0;
          Dbl可用数量  := 0;
      End;
    
      --可用数量不足则退出
      If Lngcur批次 <> Nvl(Lnglast批次, 0) Then
        If Dbl可用数量 < Dbl实际数量 And Lngcur批次 <> 0 Then
          v_Error := Str药名 || '的可用数量不足，操作中止！';
          Raise Err_Custom;
        End If;
      End If;
    
      If n_零差价管理模式 <> 0 Then
        Select Nvl(是否零差价管理, 0) Into n_药品零差价管理 From 药品规格 Where 药品id = Lng药品id;
      End If;
    
      If n_记录状态 = 1 Then
        --原始发药记录，取最新价格
        n_平均成本价 := Zl_Fun_Getoutcost(Lng药品id, Lngcur批次, Lng库房id);
      
        If n_零差价管理模式 <> 0 And n_药品零差价管理 = 1 And (v_零售价 = n_平均成本价 Or Round(v_零售价, n_流通售价小数) = Round(n_平均成本价, n_流通售价小数)) Then
          Dbl成本金额 := Dbl实际金额;
        Else
          Dbl成本金额 := Round(n_平均成本价 * Nvl(Dbl实际数量, 0), n_流通金额小数);
        End If;
      Else
        --退药再发记录，取原始单据价格
        Select a.成本价
        Into n_平均成本价
        From 药品收发记录 A, 药品收发记录 B
        Where b.Id = n_Billid And a.单据 = b.单据 And a.No = b.No And a.库房id = b.库房id And a.药品id + 0 = b.药品id And
              a.序号 = b.序号 And Nvl(a.批次, 0) = Nvl(b.批次, 0) And (a.记录状态 = 1 Or Mod(a.记录状态, 3) = 0);
      
        Dbl成本金额 := Round(n_平均成本价 * Nvl(Dbl实际数量, 0), n_流通金额小数);
      End If;
    
      Dbl实际差价 := Round(Dbl实际金额 - Dbl成本金额, n_流通金额小数);
    
      --更新药品收发记录的零售金额、成本金额及差价
      Update 药品收发记录
      Set 库房id = Lng库房id, 成本价 = n_平均成本价, 成本金额 = Dbl成本金额, 差价 = Dbl实际差价, 批次 = Lngcur批次, 批号 = Str批号, 效期 = Str效期,
          配药人 = 配药人_In, 核查人 = 核查人_In, 核查日期 = v_核查日期, 审核人 = People_In, 审核日期 = Date_In, 发药方式 = 发药方式_In, 领用人 = 领药人_In,
          汇总发药号 = 汇总发药号_In, 供药单位id = n_上次供应商id, 产地 = v_上次产地, 生产日期 = d_上次生产日期, 批准文号 = v_批准文号
      Where ID = n_Billid;
      --并发操作检查
      If Sql%RowCount = 0 Then
        v_Error := '要发药的药品记录"' || Str药名 || '"不存在，操作中止！';
        Raise Err_Custom;
      End If;
    
      --更新住院费用记录的执行状态(已执行)
      Select Decode(Sum(Nvl(付数, 1) * 实际数量), Null, 1, 0, 1, 2)
      Into Int执行状态
      From 药品收发记录
      Where 单据 = Int单据 And NO = Strno And 费用id = Lng费用id And 审核人 Is Null;
    
      --判断这个单据是门诊费用还是住院费用
      Begin
        Select 1 Into n_住院费用 From 门诊费用记录 Where ID = Lng费用id;
      Exception
        When Others Then
          n_住院费用 := 2;
      End;
    
      If n_住院费用 = 2 Then
        Update 住院费用记录
        Set 执行状态 = Int执行状态, 执行人 = Decode(People_In, Null, Zl_Username, People_In), 执行时间 = Date_In, 执行部门id = Partid_In
        Where ID = Lng费用id;
      Else
        Update 门诊费用记录
        Set 执行状态 = Int执行状态, 执行人 = Decode(People_In, Null, Zl_Username, People_In), 执行时间 = Date_In, 执行部门id = Partid_In
        Where ID = Lng费用id;
      End If;
    
      --更新未发药品记录(如果未发数为零则删除)
      Select Count(*)
      Into Int未发数
      From 药品收发记录
      Where 单据 = Int单据 And (库房id + 0 = Lng库房id Or 库房id Is Null) And NO = Strno And 审核人 Is Null And
            Nvl(LTrim(RTrim(摘要)), '小宝') <> '拒发';
    
      If Int未发数 = 0 Then
        Delete 未发药品记录
        Where NO = Strno And 单据 = Int单据 And (库房id + 0 = Lng库房id Or 库房id Is Null)
        Returning 处方类型 Into n_处方类型;
      
        --更新处方类型，按整张处方更新
        Update 药品收发记录 Set 注册证号 = n_处方类型 Where 单据 = Int单据 And NO = Strno And 库房id = Lng库房id;
      End If;
    
      --更新库存
      Zl_药品库存_Update(n_Billid, 2, 1);
    
      Zl_未审药品记录_Delete(n_Billid);
    
      --处理调价修正
      Zl_药品收发记录_调价修正(n_Billid);
    
      b_Message.Zlhis_Drug_005(Lng库房id, n_Billid);
    End If;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品收发记录_批量发药;
/



------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0059' Where 编号=&n_System;
Commit;
