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


-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--148868:李业庆,2020-05-13,高值卫材重复销账处理
Create Or Replace Procedure Zl_材料收发记录_部门退料
(
  收发id_In   In 药品收发记录.Id%Type,
  审核人_In   In 药品收发记录.审核人%Type,
  审核日期_In In 药品收发记录.审核日期%Type,
  批号_In     In 药品库存.上次批号%Type := Null,
  效期_In     In 药品库存.效期%Type := Null,
  产地_In     In 药品库存.上次产地%Type := Null,
  退料数量_In In 药品收发记录.实际数量%Type := Null,
  自动销帐_In In Integer := 0,
  退料人_In   In 药品收发记录.领用人%Type := Null,
  是否销帐_In In Integer := 0
) Is
  Err_Item Exception;
  v_Err_Msg Varchar2(100);
  v_No      药品收发记录.No%Type;

  n_记录状态   药品收发记录.记录状态%Type;
  n_执行状态   住院费用记录.执行状态%Type;
  n_部分退料   Number;
  n_入出类别id Number(18);
  n_单据       药品收发记录.单据%Type;
  n_库房id     药品收发记录.库房id%Type;
  n_药品id     药品收发记录.药品id%Type;
  n_实际数量   药品收发记录.实际数量%Type;
  n_实际金额   药品收发记录.零售金额%Type;
  n_实际成本   药品收发记录.成本金额%Type;
  n_实际差价   药品收发记录.差价%Type;
  n_费用id     药品收发记录.费用id%Type;
  n_零售价     药品收发记录.零售价%Type;
  n_实价卫材   收费项目目录.是否变价%Type;

  --处理退料时，分批核算性质改变后的处理
  n_新批次       药品收发记录.批次%Type;
  n_批次         药品收发记录.批次%Type;
  n_分批         材料特性.在用分批%Type;
  n_小数         Number(2);
  n_上次供应商id 药品库存.上次供应商id%Type;
  n_成本价       药品收发记录.成本价%Type;
  d_上次生产日期 药品库存.上次生产日期%Type;
  d_灭菌效期     药品库存.灭菌效期%Type;
  v_批准文号     药品库存.批准文号%Type;
  v_产地         药品收发记录.产地%Type;
  v_费用no       住院费用记录.No%Type;
  v_Temp         Varchar2(255);
  v_人员编号     人员表.编号%Type;
  v_人员姓名     人员表.姓名%Type;
  n_主页id       住院费用记录.主页id%Type;
  n_序号         住院费用记录.序号%Type;

  v_备货id     药品收发记录.Id%Type;
  v_入库no     药品收发记录.No%Type;
  v_入库序号   Number(5) := 0;
  n_冲销记录id 药品收发记录.Id%Type;
  n_移库       Number(1) := 0;
  v_商品条码   药品库存.商品条码%Type;
  v_内部条码   药品库存.内部条码%Type;
  v_批号       药品库存.上次批号%Type;
  d_效期       药品库存.效期%Type;
Begin
  v_Temp     := Zl_Identity;
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_人员编号 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);

  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_小数 From Dual;

  If 退料数量_In Is Not Null Then
    If 退料数量_In = 0 Then
      Return;
    End If;
  End If;

  --1、判断当前数据是否是备货卫材
  Begin
    Select 汇总发药号
    Into v_备货id
    From 药品收发记录
    Where 单据 = 21 And 审核日期 Is Not Null And
          汇总发药号 = (Select Max(a.Id)
                   From 药品收发记录 A, 药品收发记录 B
                   Where a.单据 = b.单据 And a.No = b.No And a.序号 = b.序号 And a.审核人 Is Not Null And b.Id = 收发id_In And
                         (Mod(a.记录状态, 3) = 1 Or a.记录状态 = 1)) And Rownum = 1;
  Exception
    When Others Then
      v_备货id := 0;
  End;

  Begin
    If v_备货id = 0 Then
      Select 汇总发药号
      Into v_备货id
      From 药品收发记录
      Where 单据 = 21 And 审核日期 Is Not Null And
            汇总发药号 = (Select Max(a.Id)
                     From 药品收发记录 A, 药品收发记录 B
                     Where a.单据 = b.单据 And a.No = b.No And a.序号 = b.序号 And a.审核人 Is Not Null And b.Id = 收发id_In And
                           (Mod(a.记录状态, 3) = 0)) And Rownum = 1;
    End If;
  Exception
    When Others Then
      v_备货id := 0;
  End;

  --获取该收发记录的单据、药品ID、库房ID
  Select 单据, NO, 库房id, 药品id, 费用id, 入出类别id, 记录状态, Nvl(批次, 0), 生产日期, 灭菌效期, 批准文号, 供药单位id, 成本价, 产地, 零售价, 商品条码, 内部条码, 效期, 批号
  Into n_单据, v_No, n_库房id, n_药品id, n_费用id, n_入出类别id, n_记录状态, n_批次, d_上次生产日期, d_灭菌效期, v_批准文号, n_上次供应商id, n_成本价, v_产地,
       n_零售价, v_商品条码, v_内部条码, d_效期, v_批号
  From 药品收发记录
  Where ID = 收发id_In;

  --获取该笔记录剩余未退数量、金额及差价
  --尽量避免金额及差价未出完的现象
  Select Sum(Nvl(实际数量, 0) * Nvl(付数, 1)), Sum(Nvl(零售金额, 0)), Sum(Nvl(成本金额, 0)), Sum(Nvl(差价, 0))
  Into n_实际数量, n_实际金额, n_实际成本, n_实际差价
  From 药品收发记录
  Where 审核人 Is Not Null And NO = v_No And 单据 = n_单据 And 序号 = (Select 序号 From 药品收发记录 Where ID = 收发id_In);

  --如果允许退药数为零，表示已退药
  If n_实际数量 = 0 Then
    v_Err_Msg := '该单据已被其他操作员退料，请刷新后再试！';
    Raise Err_Item;
  End If;

  If Nvl(退料数量_In, 0) > n_实际数量 Then
    v_Err_Msg := '该单据已被其他操作员部分退料，请刷新后再试！';
    Raise Err_Item;
  End If;

  --获取该材料当前是否分批的信息
  Select Nvl(在用分批, 0) Into n_分批 From 材料特性 Where 材料id = n_药品id;

  --如果是部分退料，则重新计算零售金额及差价
  n_部分退料 := 0;
  If Not (退料数量_In Is Null Or Nvl(退料数量_In, 0) = n_实际数量) Then
    n_部分退料 := 1;
  End If;

  If n_部分退料 = 1 Then
    n_实际金额 := Round(n_实际金额 * 退料数量_In / n_实际数量, n_小数);
    n_实际成本 := Round(n_实际成本 * 退料数量_In / n_实际数量, n_小数);
    n_实际差价 := Round(n_实际差价 * 退料数量_In / n_实际数量, n_小数);
    n_实际数量 := 退料数量_In;
  End If;

  --n_分批:0-不分批;1-分批;2-原分批，现不分批，按不分批处理;3-原不分批，现分批，产生新批次
  If n_分批 = 0 And n_批次 <> 0 Then
    --原分批，现不分批，按不分批处理
    n_分批 := 2;
  Elsif n_分批 <> 0 And n_批次 = 0 Then
    --原不分批,现分批,产生新的批次，并在新产生的发药记录中使用
    n_分批 := 3;
  Else
    If n_批次 = 0 Then
      n_分批 := 0;
    Else
      n_分批 := 1;
    End If;
  End If;

  If 产地_In Is Not Null Then
    v_产地 := 产地_In;
  End If;

  If 批号_In Is Not Null Then
    v_批号 := 批号_In;
  End If;
  If 效期_In Is Not Null Then
    d_效期 := 效期_In;
  End If;

  --记录状态的含义有所变化
  --冲销的记录状态        :iif(n_记录状态=1,0,1)+1
  --被冲销的记录状态        :iif(n_记录状态=1,0,1)+2
  --等待发料的记录状态    :iif(n_记录状态=1,0,1)+3
  Select 药品收发记录_Id.Nextval Into n_冲销记录id From Dual;
  --产生冲销记录
  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 灭菌效期, 付数, 填写数量, 实际数量, 成本价, 成本金额, 扣率, 零售价,
     零售金额, 差价, 摘要, 填制人, 填制日期, 配药人, 审核人, 审核日期, 费用id, 单量, 频次, 用法, 发药窗口, 领用人, 供药单位id, 生产日期, 批准文号, 商品条码, 内部条码)
    Select n_冲销记录id, n_记录状态 + Decode(n_记录状态, 1, 0, 1) + 1, n_单据, v_No, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号,
           效期, 灭菌效期, 1, -n_实际数量, -n_实际数量, 成本价, -n_实际成本, 扣率, 零售价, -n_实际金额, -n_实际差价, 摘要, 审核人_In, 审核日期_In, 配药人, 审核人_In,
           审核日期_In, 费用id, 单量, 频次, 用法, 发药窗口, 退料人_In, 供药单位id, 生产日期, 批准文号, 商品条码, 内部条码
    From 药品收发记录
    Where ID = 收发id_In;

  --如果是部分冲销，则付数填为1，实际数量为付数与实际数量的积
  --产生正常记录以供继续发料
  Select 药品收发记录_Id.Nextval Into n_新批次 From Dual;

  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 灭菌效期, 付数, 填写数量, 实际数量, 成本价, 成本金额, 扣率, 零售价,
     零售金额, 差价, 摘要, 填制人, 填制日期, 配药人, 审核人, 审核日期, 费用id, 单量, 频次, 用法, 发药窗口, 供药单位id, 生产日期, 批准文号, 商品条码, 内部条码)
    Select n_新批次, n_记录状态 + Decode(n_记录状态, 1, 0, 1) + 3, n_单据, v_No, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id,
           Decode(n_分批, 1, 批次, 3, n_新批次, Null), Decode(n_分批, 3, 产地_In, 1, 产地, Null), Decode(n_分批, 3, v_批号, 1, 批号, Null),
           Decode(n_分批, 3, d_效期, 1, 效期, Null), 灭菌效期, 1, n_实际数量, n_实际数量, 成本价, n_实际成本, 扣率, 零售价, n_实际金额, n_实际差价, 摘要, 填制人,
           填制日期, Null, Null, Null, 费用id, 单量, 频次, 用法, 发药窗口, 供药单位id, 生产日期, 批准文号, 商品条码, 内部条码
    From 药品收发记录
    Where ID = 收发id_In;

  Zl_未审药品记录_Insert(n_新批次);

  --更新病人费用记录的执行状态(0-未执行;1-完全执行;2-部分执行)
  Select Decode(Sum(Nvl(付数, 1) * 实际数量), Null, 0, 0, 0, 2)
  Into n_执行状态
  From 药品收发记录
  Where 单据 = n_单据 And NO = v_No And 费用id = n_费用id And 审核人 Is Not Null;

  If n_执行状态 = 0 Then
    Update 住院费用记录 Set 执行状态 = n_执行状态, 执行人 = Null, 执行时间 = Null Where ID = n_费用id;
    Update 门诊费用记录
    Set 执行状态 = n_执行状态, 执行人 = Null, 执行时间 = Null
    Where NO = v_No And
          序号 = (Select 序号 From 门诊费用记录 Where ID = (Select 费用id From 药品收发记录 Where ID = 收发id_In)) And
          (Mod(记录性质, 10) = 1 Or Mod(记录性质, 10) = 2) And 记录状态 <> 2 And 执行部门id = n_库房id;
  Else
    Update 住院费用记录 Set 执行状态 = n_执行状态 Where ID = n_费用id;
    Update 门诊费用记录
    Set 执行状态 = n_执行状态
    Where NO = v_No And
          序号 = (Select 序号 From 门诊费用记录 Where ID = (Select 费用id From 药品收发记录 Where ID = 收发id_In)) And
          (Mod(记录性质, 10) = 1 Or Mod(记录性质, 10) = 2) And 记录状态 <> 2 And 执行部门id = n_库房id;
  End If;

  --插入未发药品记录
  Begin
    Insert Into 未发药品记录
      (单据, NO, 病人id, 主页id, 姓名, 优先级, 对方部门id, 库房id, 发药窗口, 填制日期, 已收费, 配药人, 打印状态, 未发数)
      Select a.单据, a.No, a.病人id, a.主页id, a.姓名, Nvl(b.优先级, 0) 优先级, a.对方部门id, a.库房id, a.发药窗口, a.填制日期, a.已收费, Null, 1, 1
      From (Select b.单据, b.No, a.病人id, a.主页id, a.姓名, Decode(a.记录性质, 1, Decode(a.操作员姓名, Null, 0, 1), 1) 已收费, b.对方部门id,
                    b.库房id, b.发药窗口, b.填制日期, c.身份
             From 住院费用记录 A, 药品收发记录 B, 病人信息 C
             Where b.Id = 收发id_In And a.Id = b.费用id + 0 And a.病人id = c.病人id(+)
             Union All
             Select b.单据, b.No, a.病人id, Null As 主页id, a.姓名, Decode(a.记录性质, 1, Decode(a.操作员姓名, Null, 0, 1), 1) 已收费,
                    b.对方部门id, b.库房id, b.发药窗口, b.填制日期, c.身份
             From 门诊费用记录 A, 药品收发记录 B, 病人信息 C
             Where b.Id = 收发id_In And a.Id = b.费用id + 0 And a.病人id = c.病人id(+)) A, 身份 B
      Where b.名称(+) = a.身份;
  Exception
    When Others Then
      Null;
  End;

  --修改原记录为被冲销记录
  Update 药品收发记录 Set 记录状态 = n_记录状态 + Decode(n_记录状态, 1, 0, 1) + 2 Where ID = 收发id_In;

  --修改药品库存(反冲库存)
  Select 是否变价 Into n_实价卫材 From 收费项目目录 Where ID = n_药品id;

  If n_分批 <> 3 Then
  
    Update 药品库存
    Set 实际数量 = Nvl(实际数量, 0) + n_实际数量, 实际金额 = Nvl(实际金额, 0) + n_实际金额, 实际差价 = Nvl(实际差价, 0) + n_实际差价,
        零售价 = Decode(n_实价卫材, 1, Decode(Nvl(n_批次, 0), 0, Null, Decode(Nvl(零售价, 0), 0, n_零售价, 零售价)), Null)
    Where 库房id + 0 = n_库房id And 药品id = n_药品id And 性质 = 1 And Nvl(批次, 0) = n_批次;
  
    If Sql%RowCount = 0 Then
      Insert Into 药品库存
        (库房id, 药品id, 批次, 性质, 实际数量, 实际金额, 实际差价, 效期, 灭菌效期, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 批准文号, 零售价, 平均成本价, 商品条码,
         内部条码)
      Values
        (n_库房id, n_药品id, Decode(n_分批, 2, Null, n_批次), 1, n_实际数量, n_实际金额, n_实际差价, Decode(n_分批, 1, d_效期, Null), d_灭菌效期,
         n_上次供应商id, n_成本价, Decode(n_分批, 1, v_批号, Null), d_上次生产日期, v_产地, v_批准文号,
         Decode(n_实价卫材, 1, Decode(Nvl(n_批次, 0), 0, Null, n_零售价), Null), n_成本价, v_商品条码, v_内部条码);
    End If;
  Else
    Insert Into 药品库存
      (库房id, 药品id, 批次, 性质, 实际数量, 实际金额, 实际差价, 效期, 灭菌效期, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 批准文号, 零售价, 平均成本价, 商品条码, 内部条码)
    Values
      (n_库房id, n_药品id, n_新批次, 1, n_实际数量, n_实际金额, n_实际差价, d_效期, d_灭菌效期, n_上次供应商id, n_成本价, v_批号, d_上次生产日期, v_产地, v_批准文号,
       Decode(n_实价卫材, 1, Decode(Nvl(n_新批次, 0), 0, Null, n_零售价), Null), n_成本价, v_商品条码, v_内部条码);
  End If;

  Delete 药品库存
  Where 库房id + 0 = n_库房id And 药品id = n_药品id And 性质 = 1 And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
        Nvl(实际差价, 0) = 0;

  If 自动销帐_In = 1 And n_单据 <> 24 Then
    Begin
      Select 主页id, NO, 序号 Into n_主页id, v_费用no, n_序号 From 住院费用记录 Where ID = n_费用id;
    Exception
      When Others Then
        Begin
          Select Null, NO, 序号 Into n_主页id, v_费用no, n_序号 From 门诊费用记录 Where ID = n_费用id;
        Exception
          When Others Then
            n_主页id := Null;
        End;
    End;
    If n_主页id Is Null Then
      Zl_门诊记帐记录_Delete(v_费用no, n_序号, v_人员编号, v_人员姓名);
    Else
      Zl_住院记帐记录_Delete(v_费用no, n_序号, v_人员编号, v_人员姓名);
    End If;
  End If;

  If Not (退料数量_In Is Null) Then
    --备货卫材处理
    If v_备货id > 0 Then
      --2、自动冲销已审核的其他出库单据
      Begin
        Select 1
        Into n_移库
        From 药品收发记录
        Where 单据 = 15 And 审核日期 Is Null And
              费用id In (Select Distinct 费用id From 药品收发记录 Where NO = v_No And 药品id = n_药品id And 批次 = n_批次);
      Exception
        When Others Then
          n_移库 := 0;
      End;
    
      If n_移库 <> 0 Then
        For v_出库冲销 In (Select 1 行次, 记录状态, NO, 序号, 药品id
                       From 药品收发记录
                       Where 单据 = 21 And 审核日期 Is Not Null And 汇总发药号 = v_备货id) Loop
          Zl_材料其他出库_Strike(v_出库冲销.行次, v_出库冲销.记录状态, v_出库冲销.No, v_出库冲销.序号, v_出库冲销.药品id, 退料数量_In, 审核人_In, 审核日期_In);
        End Loop;
      
        --3、删除未审核的外购入库单据（已审核则不管）
        If n_部分退料 = 1 Then
          Update 药品收发记录
          Set 填写数量 = 填写数量 - 退料数量_In, 实际数量 = 实际数量 - 退料数量_In, 零售金额 = 零售金额 - n_实际金额, 成本金额 = 成本金额 - n_实际成本, 差价 = 差价 - n_实际差价
          Where 单据 = 15 And 药品id = n_药品id And Nvl(批次, 0) = n_批次 And 费用id = n_费用id And 审核日期 Is Null;
        Else
          Delete 药品收发记录
          Where 单据 = 15 And 药品id = n_药品id And Nvl(批次, 0) = n_批次 And 费用id = n_费用id And 审核日期 Is Null;
        End If;
      End If;
    End If;
  Else
    --备货卫材处理
    If v_备货id > 0 Then
      --2、自动冲销已审核的其他出库单据
      Begin
        Select 1
        Into n_移库
        From 药品收发记录
        Where 单据 = 15 And 审核日期 Is Null And
              费用id In (Select Distinct 费用id From 药品收发记录 Where NO = v_No And 药品id = n_药品id And 批次 = n_批次);
      Exception
        When Others Then
          n_移库 := 0;
      End;
    
      If n_移库 <> 0 Then
        For v_出库冲销 In (Select 1 行次, 记录状态, NO, 序号, 药品id
                       From 药品收发记录
                       Where 单据 = 21 And 审核日期 Is Not Null And 汇总发药号 = v_备货id) Loop
          Zl_材料其他出库_Strike(v_出库冲销.行次, v_出库冲销.记录状态, v_出库冲销.No, v_出库冲销.序号, v_出库冲销.药品id, 退料数量_In, 审核人_In, 审核日期_In, 1);
        End Loop;
      
        --3、产生新的其他出库单据
        If v_入库no Is Null Then
          v_入库no := Nextno(74, n_库房id);
        End If;
        v_入库序号 := v_入库序号 + 1;
      
        For v_入库 In (Select 入出类别id, 库房id, 药品id, 批次, 填写数量, 成本价, 成本金额, 零售价, 零售金额, 差价, 产地, 批号, 效期, 灭菌效期, 摘要, 单量, 发药窗口
                     From 药品收发记录
                     Where 单据 = 21 And 审核日期 Is Not Null And 汇总发药号 = v_备货id) Loop
        
          Zl_材料其他出库_Insert(v_入库.入出类别id, v_入库no, v_入库序号, v_入库.库房id, v_入库.药品id, v_入库.批次, v_入库.填写数量, v_入库.成本价, v_入库.成本金额,
                           v_入库.零售价, v_入库.零售金额, v_入库.差价, 审核人_In, 审核日期_In, v_入库.产地, v_入库.批号, v_入库.效期, v_入库.灭菌效期, v_入库.摘要,
                           v_入库.单量, v_入库.发药窗口);
        
          Update 药品收发记录
          Set 费用id = n_费用id, 汇总发药号 = n_新批次
          Where 单据 = 21 And NO = v_入库no And 序号 = v_入库序号;
        End Loop;
      
        --4、删除未审核的外购入库单据（已审核则不管）
        Delete 药品收发记录
        Where 单据 = 15 And 药品id = n_药品id And Nvl(批次, 0) = n_批次 And 费用id = n_费用id And 审核日期 Is Null;
      End If;
    End If;
  End If;
  --处理调价修正单据
  Zl_材料收发记录_调价修正(n_冲销记录id);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料收发记录_部门退料;
/


-------------------------------------------------------------------------------
---标准部件信息
EXECUTE Zlfiles_Autoupdate('zl9Stuff.dll','A5034D978107F88821EB8AD9C20384C0','10.35.90.0093',to_date('2020-05-14 16:04:48','YYYY-MM-DD HH24:MI:SS'),SYSDATE,'1','[APPSOFT]\APPLY','zl9Stuff','1','卫材管理部件','1','0','');
-------------------------------------------------------------------------------


------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0093' Where 编号=&n_System;
Commit;
