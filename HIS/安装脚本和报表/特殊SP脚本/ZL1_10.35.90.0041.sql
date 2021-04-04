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
--135585:胡俊勇,2018-12-12,卫材跨批次超期收回
CREATE OR REPLACE Procedure Zl_病人医嘱记录_收回
(
  --功能：将指定医嘱超期发送部分收回。如果上次发送没有产生费用，则仅收回医嘱的上次执行时间。
  --参数：
  --      收回量_IN=对西药、中成药为按住院单位的收回量,对中药为收回付数,对其它医嘱为收回总量或次数。
  --      医嘱ID_IN=每条要收回的医嘱记录的ID(明细存储的ID),对成药或配方,不一定包含给药途径或用法煎法(可能为叮嘱而未读取)
  --      上次时间_IN=医嘱超期发送部分收回后应该还原的上次执行时间(严格按频率计算得来),为空时表示被全部收回了。
  --      NO_IN=当收回要产生负数费用记录时，为新生成记录的单据号(供费用及药品使用),当前处理的只是新NO的一部份。
  --            因为药品可能分批,所以序号在处理时取。
  --            如果全是划价单（传入值为：调整划价单），则不产生负数单据，直接修改或删除划价单
  收回量_In     In 病人医嘱发送.发送数次%Type,
  医嘱id_In     In 病人医嘱记录.Id%Type,
  上次时间_In   In 病人医嘱记录.上次执行时间%Type,
  收回时间_In   In 病人医嘱记录.上次执行时间%Type,
  No_In         In 住院费用记录.No%Type := Null,
  操作员编号_In In 人员表.编号%Type := Null,
  操作员姓名_In In 人员表.姓名%Type := Null
) Is
  --收回医嘱对应的发送费用明细的剩余数量,按后产生的费用先收回
  --剩余数量没有排开已申请的数量部份，在产生新申请时覆盖原来的申请
  --对药品和卫材，对一个数量，可能存在未执行和已执行部分，需分别填写申请记录，且以未执行优先
  --执行标志=0-未执行,1-已执行；药品的有部分执行，以收发记录中的明细量区分为准；非药品的只优先处理未执行的
  Cursor c_Detail Is
    Select a.费用id, a.No, a.序号, a.收费细目id, a.病人病区id, a.收费类别, a.跟踪在用, a.诊疗类别, a.医嘱内容, a.单次用量, a.剩余数量, a.已执行量, a.未执行量,
           a.执行标志, a.记录状态, a.登记时间, a.收费方式
    From (With 医嘱费用记录 As (Select Max(Decode(b.记录状态, 2, 0, b.Id)) As 费用id, b.No, Nvl(b.价格父号, b.序号) As 序号, b.收费细目id,
                                 b.病人病区id, Sum(Nvl(b.付数, 1) * b.数次) As 剩余数量, b.收费类别, Max(Nvl(b.执行状态, 0)) As 执行状态, d.跟踪在用,
                                 c.诊疗类别, c.医嘱内容, c.单次用量, Max(b.记录状态) As 记录状态, Max(b.登记时间) As 登记时间, Nvl(e.收费方式, 0) As 收费方式
                          From 病人医嘱发送 A, 住院费用记录 B, 病人医嘱记录 C, 材料特性 D, 病人医嘱计价 E
                          Where a.医嘱id = 医嘱id_In And a.No = b.No And a.记录性质 = b.记录性质 And a.医嘱id = b.医嘱序号 And
                                b.价格父号 Is Null And b.收费细目id = d.材料id(+) And c.Id = 医嘱id_In And e.医嘱id(+) = b.医嘱序号 And
                                e.收费细目id(+) = b.收费细目id And Not Exists
                           (Select 1 From 输液配药记录 F Where f.医嘱id = c.相关id And a.发送号 = f.发送号)
                          Group By b.No, b.记录性质, Nvl(b.价格父号, b.序号), b.收费细目id, b.病人病区id, b.收费类别, d.跟踪在用, c.诊疗类别, c.医嘱内容,
                                   c.单次用量, e.收费方式
                          Having Sum(Nvl(b.付数, 1) * b.数次) > 0)
           Select 费用id, NO, 序号, 收费细目id, 病人病区id, 收费类别, 跟踪在用, 诊疗类别, 医嘱内容, 单次用量, 剩余数量, Null As 已执行量, Null As 未执行量,
                  执行状态 As 执行标志, 记录状态, 登记时间, 收费方式
           From 医嘱费用记录
           Where 收费类别 Not In ('5', '6', '7') And Not (收费类别 = '4' And Nvl(跟踪在用, 0) = 1)
           Union All
           Select a.费用id, a.No, a.序号, a.收费细目id, a.病人病区id, a.收费类别, a.跟踪在用, a.诊疗类别, a.医嘱内容, a.单次用量, a.剩余数量, 0 As 已执行量,
                  Sum(Nvl(b.付数, 1) * b.实际数量) As 未执行量, 0 As 执行标志, a.记录状态, Max(a.登记时间) As 登记时间, a.收费方式
           From 医嘱费用记录 A, 药品收发记录 B
           Where (a.收费类别 In ('5', '6', '7') Or (a.收费类别 = '4' And Nvl(a.跟踪在用, 0) = 1)) And a.费用id = b.费用id And
                 a.No = b.No And b.单据 In (9, 10, 25, 26) And Mod(b.记录状态, 3) = 1 And b.审核人 Is Null
           Group By a.费用id, a.No, a.序号, a.记录状态, a.收费细目id, a.病人病区id, a.收费类别, a.跟踪在用, a.诊疗类别, a.医嘱内容, a.单次用量, a.剩余数量,
                    a.收费方式
           Having Sum(Nvl(b.付数, 1) * b.实际数量) > 0
           Union All
           Select a.费用id, a.No, a.序号, a.收费细目id, a.病人病区id, a.收费类别, a.跟踪在用, a.诊疗类别, a.医嘱内容, a.单次用量, a.剩余数量,
                  Sum(Nvl(b.付数, 1) * b.实际数量) As 已执行量, 0 As 未执行量, 1 As 执行标志, a.记录状态, Max(a.登记时间) As 登记时间, a.收费方式
           From 医嘱费用记录 A, 药品收发记录 B
           Where (a.收费类别 In ('5', '6', '7') Or (a.收费类别 = '4' And Nvl(a.跟踪在用, 0) = 1)) And a.费用id = b.费用id And
                 a.No = b.No And b.单据 In (9, 10, 25, 26) And Not (Mod(b.记录状态, 3) = 1 And b.审核人 Is Null)
           Group By a.费用id, a.No, a.序号, a.记录状态, a.收费细目id, a.病人病区id, a.收费类别, a.跟踪在用, a.诊疗类别, a.医嘱内容, a.单次用量, a.剩余数量,
                    a.收费方式
           Having Sum(Nvl(b.付数, 1) * b.实际数量) > 0) A
           Order By Decode(a.诊疗类别, '5', 0, '6', 0, '7', 0, a.收费细目id), a.执行标志, a.登记时间 Desc;


  Cursor c_Applay(v_费用ids Varchar2) Is
    Select a.费用id, b.No, b.序号, a.数量, a.申请时间, a.申请类别
    From 病人费用销帐 A, 住院费用记录 B
    Where a.费用id = b.Id And a.申请部门id = a.审核部门id And a.申请时间 = 收回时间_In And
          a.费用id In (Select * From Table(Cast(f_Num2list(v_费用ids) As Zltools.t_Numlist)))
    Order By NO, 序号;

  --包含指定药品长嘱发送时产生的相关费用及药品/卫材记录信息(因多次发送有多条记录,分批的已在界面禁止)
  --药品医嘱填写了"病人医嘱发送"记录,对应的给药途径不一定填写了的(可能为叮嘱),且NO不同。
  --因为要收回的次数可能包含了多次发送的内容,所以要将多次发送的收发记录都取出来，多次发送时，划价的先收回（修改或删除）
  Cursor c_Drug Is
    Select a.病人id, a.主页id, d.姓名, Nvl(Nvl(x.剂量系数, y.换算系数), 1) As 剂量系数, Nvl(x.住院包装, 1) As 住院包装,
           Nvl(x.最大效期, y.最大效期) As 最大效期, Nvl(b.付数, 1) * b.实际数量 As 数量, b.Id As 收发id, b.单据, b.药品id, b.对方部门id, b.库房id,
           b.费用id, Nvl(Nvl(x.药房分批, y.在用分批), 0) As 分批, b.批次, b.批号, b.效期, a.记录状态, a.No, a.序号, a.收费细目id, a.执行状态 As 执行标志
    From 住院费用记录 A, 药品收发记录 B, 病人医嘱发送 C, 病人信息 D, 药品规格 X, 材料特性 Y
    Where c.医嘱id = 医嘱id_In And a.No = c.No And a.记录性质 = c.记录性质 And a.记录状态 In (0, 1, 3) And a.医嘱序号 + 0 = 医嘱id_In And
          a.No = b.No And a.Id = b.费用id + 0 And b.单据 In (9, 10, 25, 26) And (b.记录状态 = 1 Or Mod(b.记录状态, 3) = 0) And
          a.病人id = d.病人id And b.药品id = x.药品id(+) And b.药品id = y.材料id(+)
    Order By a.记录状态, b.No Desc, b.Id Desc;

  --包含非药长嘱(含给药途径)发送时所产生的费用(因多个收入而有多条记录)
  --对非药医嘱,直接收回指定量,不管多次发送(如果多次发送价格不同,则收回的价格是以最后次的；不然就要根据多个收入依次减收回量)。
  --卫材本身是售价单位，无需住院单位转换
  --非药长嘱都填写了发送记录(除开了叮嘱及护理等级)
  --一天只收一次或一次发送只收一次的项目暂时不支持负数申请
  Cursor c_Other(n_发送号 病人医嘱发送.发送号%Type) Is
    With 医嘱费用记录 As
     (Select a.No, a.序号, a.记录状态, a.收费细目id, a.Id As 费用id, a.数次 As 剩余数量, Nvl(a.执行状态, 0) As 执行状态, a.医嘱序号, b.发送号,
             c.数量 As 对照数量, Nvl(c.收费方式, 0) As 收费方式, a.收费类别
      From 住院费用记录 A, 病人医嘱发送 B, 病人医嘱计价 C
      Where a.No = b.No And a.记录性质 = b.记录性质 And a.医嘱序号 + 0 = b.医嘱id And b.医嘱id = 医嘱id_In And a.医嘱序号 = c.医嘱id(+) And
            a.收费细目id = c.收费细目id(+))
    Select a.No, a.序号, a.费用id, a.剩余数量, a.收费细目id, a.记录状态, a.执行状态, a.对照数量, a.收费方式, a.收费类别
    From (Select a.No, a.序号, a.记录状态, a.收费细目id, a.费用id, a.剩余数量, a.对照数量, a.执行状态, a.医嘱序号, a.收费方式, a.收费类别
           From 医嘱费用记录 A
           Where a.记录状态 In (1, 3) And a.发送号 = n_发送号
           Union All
           Select a.No, a.序号, a.记录状态, a.收费细目id, a.费用id, a.剩余数量, a.对照数量, a.执行状态, a.医嘱序号, a.收费方式, a.收费类别
           From 医嘱费用记录 A
           Where a.记录状态 = 0) A
    Order By a.收费细目id, a.序号, a.记录状态;

  --按序号排序是为了产生新记录时,填写同一收费细目的不同收入项目的价格父号

  --该游标用于处理费用相关汇总表
  Cursor c_Money
  (
    v_Start 住院费用记录.序号%Type,
    v_End   住院费用记录.序号%Type
  ) Is
    Select 病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, Sum(Nvl(应收金额, 0)) As 应收金额, Sum(Nvl(实收金额, 0)) As 实收金额
    From 住院费用记录
    Where 记录性质 = 2 And 记录状态 = 1 And NO = No_In And 序号 Between v_Start And v_End
    Group By 病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id;

  --系统参数指定执行后需要自动审核的划价费用：用于非药医嘱，包含对应的药品及卫材费用
  Cursor c_Verify
  (
    v_Start 住院费用记录.序号%Type,
    v_End   住院费用记录.序号%Type
  ) Is
    Select NO, 序号
    From 住院费用记录
    Where 记录性质 = 2 And 记录状态 = 0 And NO = No_In And 价格父号 Is Null And 序号 Between v_Start And v_End;

  Cursor c_Compound
  (
    相关id_In       病人医嘱记录.相关id%Type,
    执行终止时间_In 病人医嘱记录.执行终止时间%Type,
    配药id_In       输液配药记录.Id%Type,
    医嘱序号_In     病人医嘱记录.Id%Type
  ) Is
    Select b.费用id, b.药品id As 收费细目id, Sum(a.数量) As 数量, c.住院包装, c.住院单位, d.名称, e.病人病区id, e.操作状态, e.Id As 配药id, f.No,
           Nvl(f.价格父号, f.序号) As 序号, f.记录状态 As 记录状态, f.执行状态 As 执行标志
    From 输液配药内容 A, 药品收发记录 B, 药品规格 C, 收费项目目录 D, 输液配药记录 E, 住院费用记录 F
    Where a.收发id = b.Id And b.药品id = c.药品id And c.药品id = d.Id And e.Id = a.记录id And f.No = b.No And f.Id = b.费用id And
          e.医嘱id = 相关id_In And e.执行时间 > 执行终止时间_In And e.Id = 配药id_In And f.医嘱序号 + 0 = 医嘱序号_In
    Group By b.费用id, b.药品id, c.住院包装, c.住院单位, d.名称, e.病人病区id, e.操作状态, e.Id, f.No, f.价格父号, f.序号, f.记录状态, f.执行状态;

  --审核中包含跟踪在用的未发卫料时，根据参数设置是否自动发料
  Cursor c_Stuff Is
    Select m.Id, m.库房id, Decode(b.在用分批, 1, m.批次, 0) As 批次
    From 药品收发记录 M, 住院费用记录 A, 材料特性 B
    Where m.No = No_In And m.单据 In (25, 26) And m.库房id Is Not Null And m.记录状态 = 1 And m.审核人 Is Null And m.No = a.No And
          a.Id = m.费用id + 0 And a.记录性质 = 2 And a.记录状态 = 1 And a.收费细目id = b.材料id And b.跟踪在用 = 1
    Order By m.库房id, m.药品id;

  v_Dec      Number;
  v_First    Number;
  v_划价类别 Varchar2(255);

  v_诊疗类别 病人医嘱记录.诊疗类别%Type;
  v_单次用量 病人医嘱记录.单次用量%Type;
  v_跟踪在用 材料特性.跟踪在用%Type;

  v_费用序号 住院费用记录.序号%Type;
  v_收发序号 药品收发记录.序号%Type;
  v_费用id   住院费用记录.Id%Type;
  v_实收金额 住院费用记录.实收金额%Type;

  v_开始序号 住院费用记录.序号%Type;
  v_结束序号 住院费用记录.序号%Type;

  v_医嘱执行 病人医嘱发送.执行状态%Type;

  v_剂量系数 药品规格.剂量系数%Type;
  v_住院包装 药品规格.住院包装%Type;
  v_医嘱内容 病人医嘱记录.医嘱内容%Type;

  v_结帐参数       Zlparameters.参数值%Type;
  v_配液药销帐申请 Zlparameters.参数值%Type;
  v_结帐金额       住院费用记录.结帐金额%Type;

  v_收费细目id   住院费用记录.收费细目id%Type;
  v_剩余数量     住院费用记录.数次%Type;
  v_收回数量     住院费用记录.数次%Type;
  v_当前数量     住院费用记录.数次%Type;
  v_当前付数     住院费用记录.付数%Type;
  v_费用ids      Varchar2(4000);
  v_组id         病人医嘱记录.Id%Type;
  v_对照数量     病人医嘱计价.数量%Type;
  v_收回量       住院费用记录.数次%Type;
  v_收回剩余     住院费用记录.数次%Type;
  v_输液收回剩余 住院费用记录.数次%Type;
  n_数量         药品收发记录.填写数量%Type;

  v_Delno    Varchar2(4000);
  v_Temp     Varchar2(4000);
  v_收费内容 Varchar2(4000);
  v_No       住院费用记录.No%Type;
  v_人员编号 住院费用记录.操作员编号%Type;
  v_人员姓名 住院费用记录.操作员姓名%Type;

  n_相关id       病人医嘱记录.相关id%Type;
  d_执行终止时间 病人医嘱记录.执行终止时间%Type;
  b_输液配药记录 Boolean;
  d_收回时间     病人医嘱记录.执行终止时间%Type;
  n_申请类别     病人费用销帐.申请类别%Type;
  v_销帐原因     病人费用销帐.销帐原因%Type;
  n_Count        Number;
  v_Lngid        药品收发记录.Id%Type; --收发ID
  n_Tmp序号      病人医嘱记录.序号%Type;
  n_配液更新     Number; ----是否更新输液配药记录的状态
  n_发料号       药品收发记录.汇总发药号%Type;
  n_库房id       药品收发记录.库房id%Type;
  v_收发ids      Varchar2(4000);
  v_Error        Varchar2(255);
  Err_Custom Exception;

  Procedure 负数收发记录_Insert
  (
    费用id_In     Number,
    批次_In       药品收发记录.批次%Type,
    分批_In       药品规格.药房分批%Type,
    批号_In       药品收发记录.批号%Type,
    效期_In       药品收发记录.效期%Type,
    最大效期_In   药品规格.最大效期%Type,
    收发id_In     药品收发记录.Id%Type,
    病人id_In     住院费用记录.病人id%Type,
    主页id_In     住院费用记录.主页id%Type,
    库房id_In     药品收发记录.库房id%Type,
    单据_In       药品收发记录.单据%Type,
    姓名_In       病人信息.姓名%Type,
    对方部门id_In 药品收发记录.对方部门id%Type,
    收费类别_In   住院费用记录.收费类别%Type,
    划价类别_In   Varchar,
    P付数         药品收发记录.付数%Type,
    P数量         药品收发记录.填写数量%Type
  ) Is
    v_批次   药品收发记录.批次%Type;
    v_效期   药品收发记录.效期%Type;
    v_批号   药品收发记录.批号%Type;
    v_优先级 身份.优先级%Type;
  Begin
    --确定批次
    If Nvl(批次_In, 0) <> 0 And 分批_In = 0 Then
      --原分批,现不分批
      v_批次 := Null;
      v_批号 := 批号_In;
      v_效期 := 效期_In;
    Elsif Nvl(批次_In, 0) = 0 And 分批_In = 1 Then
      --原不分批,现分批
      Select 药品收发记录_Id.Nextval Into v_批次 From Dual;
      Select To_Char(Sysdate, 'YYYYMMDD') Into v_批号 From Dual;
      If 最大效期_In Is Not Null Then
        v_效期 := Trunc(Sysdate + 最大效期_In * 30);
      Else
        v_效期 := Null;
      End If;
    Else
      v_批次 := 批次_In;
      v_批号 := 批号_In;
      v_效期 := 效期_In;
    End If;
  
    Select 药品收发记录_Id.Nextval Into v_Lngid From Dual;
    Insert Into 药品收发记录
      (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 付数, 填写数量, 实际数量, 零售价, 零售金额, 摘要, 填制人, 填制日期,
       费用id, 单量, 频次, 用法, 供药单位id, 生产日期, 批准文号, 灭菌效期)
      Select v_Lngid, 1, 单据, No_In, v_收发序号, 库房id, 对方部门id, 入出类别id, -1, 药品id, Nvl(v_批次, 0), 产地, v_批号, v_效期, P付数, -1 * P数量,
             -1 * P数量, 零售价, Round(-1 * P付数 * P数量 * 零售价, v_Dec), '超期发送收回', v_人员姓名, 收回时间_In, 费用id_In, 单量, 频次, 用法, 供药单位id,
             生产日期, 批准文号, 灭菌效期
      From 药品收发记录
      Where ID = 收发id_In;
  
    Zl_未审药品记录_Insert(v_Lngid);
  
    Zl_药品库存_Update(v_Lngid, 0, 1);
  
    --未发药品记录
    Update 未发药品记录
    Set 病人id = 病人id_In, 主页id = 主页id_In, 姓名 = 姓名_In
    Where 单据 = 单据_In And NO = No_In And 库房id + 0 = 库房id_In;
  
    If Sql%RowCount = 0 Then
      --取身份优先级
      Begin
        Select b.优先级 Into v_优先级 From 病人信息 A, 身份 B Where a.身份 = b.名称(+) And a.病人id = 病人id_In;
      Exception
        When Others Then
          Null;
      End;
    
      Insert Into 未发药品记录
        (单据, NO, 病人id, 主页id, 姓名, 优先级, 对方部门id, 库房id, 填制日期, 已收费, 打印状态)
      Values
        (单据_In, No_In, 病人id_In, 主页id_In, 姓名_In, v_优先级, 对方部门id_In, 库房id_In, 收回时间_In,
         Decode(Nvl(Instr(划价类别_In, Decode(收费类别_In, '4', '4', '5')), 0), 0, 1, 0), 0);
    End If;
  
    v_收发序号 := v_收发序号 + 1;
  End;
Begin
  --取操作员信息(部门ID,部门名称;人员ID,人员编号,人员姓名)
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
  --检查是否是输液配液记录，并是否已经锁定
  Select 医嘱内容, Nvl(相关id, ID) Into v_医嘱内容, n_相关id From 病人医嘱记录 Where ID = 医嘱id_In;
  Select Count(1)
  Into n_Count
  From 输液配药记录 A, 病人医嘱记录 B
  Where a.医嘱id = b.Id And 医嘱id = 医嘱id_In And a.执行时间 > b.执行终止时间 And a.是否锁定 = 1;

  If n_Count > 0 Then
    v_Error := '医嘱"' || v_医嘱内容 || '"是输液药品，已经被输液配置中心锁定，不能超期收回。';
    Raise Err_Custom;
  End If;
  Select Max(操作说明) Into v_销帐原因 From 病人医嘱状态 Where 医嘱id = 医嘱id_In And 操作类型 = 8;
  If Nvl(收回量_In, 0) > 0 Then
    --判断是否是输液配药药品(输液配制中心药品统一走销帐申请)
    b_输液配药记录 := False;
    v_输液收回剩余 := 收回量_In;
  
    Select Max(a.执行终止时间)
    Into d_执行终止时间
    From 输液配药记录 E, 病人医嘱记录 A
    Where a.Id = n_相关id And e.医嘱id = a.Id And e.执行时间 > a.执行终止时间;
  
    If d_执行终止时间 Is Not Null Then
      d_收回时间       := 收回时间_In;
      v_配液药销帐申请 := zl_GetSysParameter('配液输液单配药后允许销帐申请', 1345);
      b_输液配药记录   := True;
    
      If n_相关id = 医嘱id_In Then
        --给药途径行，更新状态，但不改变数量
        n_配液更新 := 1;
      Else
        n_配液更新 := 0;
        n_Tmp序号  := 医嘱id_In;
      End If;
    
      For X In (Select e.Id As 配药id, e.操作状态, e.是否打包
                From 输液配药记录 E
                Where e.医嘱id = n_相关id And e.执行时间 > d_执行终止时间 And Nvl(e.操作状态, 0) In (1, 2, 3, 4, 5, 6, 7, 8)) Loop
        If Not (x.操作状态 In (4, 5, 6, 7, 8) And Nvl(x.是否打包, 0) = 0 And Nvl(v_配液药销帐申请, '0') = '0') Then
          If n_配液更新 = 0 Then
            --产生药品行明细销帐申请
            For r_Compound In c_Compound(n_相关id, d_执行终止时间, x.配药id, n_Tmp序号) Loop
            
              v_输液收回剩余 := v_输液收回剩余 - r_Compound.数量;
              If x.操作状态 = 1 Then
                n_申请类别 := 0;
              Else
                n_申请类别 := 1;
              End If;
              Zl_病人费用销帐_Insert(r_Compound.费用id, r_Compound.收费细目id, r_Compound.病人病区id, r_Compound.数量, v_人员姓名, d_收回时间,
                               n_申请类别, Null, r_Compound.配药id, v_销帐原因, 0);
              If x.操作状态 = 1 Then
                --未发药的，自动审核。
                Zl_病人费用销帐_Audit(r_Compound.费用id, d_收回时间, v_人员姓名, d_收回时间, 1, 1, n_申请类别);
                Zl_住院记帐记录_Delete(r_Compound.No, r_Compound.序号 || ':' || r_Compound.数量 || ':' || r_Compound.配药id, v_人员编号,
                                 v_人员姓名, 2, Null, Null, d_收回时间);
              End If;
            End Loop;
          End If;
        
          --更新状态
          If n_配液更新 = 1 Then
            Select Count(1)
            Into n_Count
            From 输液配药状态
            Where 配药id = x.配药id And 操作类型 = 9 And 操作时间 = d_收回时间;
            If n_Count = 0 Then
              Insert Into 输液配药状态
                (配药id, 操作类型, 操作人员, 操作时间, 操作说明)
              Values
                (x.配药id, 9, v_人员姓名, d_收回时间, v_销帐原因);
            End If;
            Update 输液配药记录 Set 操作人员 = v_人员姓名, 操作时间 = d_收回时间, 操作状态 = 9 Where ID = x.配药id;
          
            If x.操作状态 = 1 Then
              Insert Into 输液配药状态
                (配药id, 操作类型, 操作人员, 操作时间)
              Values
                (x.配药id, 10, v_人员姓名, d_收回时间);
              Update 输液配药记录 Set 操作人员 = v_人员姓名, 操作时间 = d_收回时间, 操作状态 = 10 Where ID = x.配药id;
            End If;
          End If;
        
          --由于不同批次（执行时间）申请时，申请时间和费用ID有唯一约束，所以同时销帐多个批次时，依次加一秒
          d_收回时间 := d_收回时间 + 1 / 24 / 60 / 60;
        End If;
      End Loop;
    End If;
  
    --a.销帐申请收回模式
    --输液配药记录单独进行销帐
    If b_输液配药记录 = False Or v_输液收回剩余 > 0 Then
      If No_In Is Null Then
        v_结帐参数 := zl_GetSysParameter(23);
        --根据收回数量对照原始费用进行分摊申请
        For r_Detail In c_Detail Loop
          --确定该收费细目ID的收回总数量
          If Nvl(v_收费细目id, 0) <> r_Detail.收费细目id And (r_Detail.诊疗类别 Not In ('5', '6', '7') Or Nvl(v_收费细目id, 0) = 0) Then
            --数量未分摊完成
            If v_收费细目id Is Not Null And v_收回数量 > 0 Then
              v_Error := '医嘱"' || v_医嘱内容 || '"对应的费用剩余数量不足收回数量，可能存在手工销帐或已结帐的费用，或相应的划价单已被删除。';
              Raise Err_Custom;
            End If;
            --药品收回总量是以最后发送规格为准计算的，以此计算出收回售价数量
            Begin
              Select 剂量系数, 住院包装 Into v_剂量系数, v_住院包装 From 药品规格 Where 药品id = r_Detail.收费细目id;
            Exception
              When Others Then
                Null;
            End;
            --计算收回数量，如果收费方式不是0正常收取，则使用特殊的方法进行计算
            If r_Detail.收费方式 = 0 Then
              If r_Detail.诊疗类别 = '7' Then
                --中药配方药品：付数*单量
                v_收回数量 := Round(v_输液收回剩余 * r_Detail.单次用量 / Nvl(v_剂量系数, 1), 5);
              Else
                If r_Detail.诊疗类别 Not In ('5', '6') Then
                  Select Nvl(Max(数量), 1)
                  Into v_对照数量
                  From 病人医嘱计价
                  Where 医嘱id = 医嘱id_In And 收费细目id = r_Detail.收费细目id;
                Else
                  v_对照数量 := 1;
                End If;
                v_收回数量 := Round(v_输液收回剩余 * Nvl(v_住院包装, 1), 5) * v_对照数量;
              End If;
            Else
              Select Nvl(Sum(数量), 0)
              Into v_收回数量
              From 医嘱执行计价
              Where 医嘱id = 医嘱id_In And 收费细目id = r_Detail.收费细目id And
                    要求时间 > Nvl(上次时间_In, To_Date('1900-01-01', 'yyyy-MM-dd'));
            
              v_收回数量 := Round(v_收回数量, 5);
            
            End If;
            v_医嘱内容 := r_Detail.医嘱内容;
          End If;
        
          --该收费细目的每个费用明细分摊收回
          If v_收回数量 > 0 Then
            --检查对应费用是否已结帐，当禁止时
            v_结帐金额 := 0;
            If v_结帐参数 = '2' And r_Detail.记录状态 <> 0 Then
              Select Sum(结帐金额)
              Into v_结帐金额
              From 住院费用记录
              Where NO = r_Detail.No And 记录性质 In (2, 12) And Nvl(价格父号, 序号) = r_Detail.序号;
            End If;
          
            If Nvl(v_结帐金额, 0) = 0 Then
              If r_Detail.收费类别 In ('5', '6', '7') Or r_Detail.收费类别 = '4' And r_Detail.跟踪在用 = 1 Then
                --药品和跟踪在用的卫材
                If r_Detail.执行标志 = 0 Then
                  v_剩余数量 := r_Detail.未执行量;
                Elsif r_Detail.执行标志 = 1 Then
                  v_剩余数量 := r_Detail.已执行量;
                End If;
              Else
                --普通费用
                v_剩余数量 := r_Detail.剩余数量;
              End If;
              If v_收回数量 > v_剩余数量 Then
                v_当前数量 := v_剩余数量;
              Else
                v_当前数量 := v_收回数量;
              End If;
              v_收回数量 := v_收回数量 - v_当前数量;
              --系统参数决定执行后是否审核划价单，所以，已执行的仍然可能是划价单
              If r_Detail.执行标志 = 0 And r_Detail.记录状态 = 0 Then
                v_Delno := v_Delno || '|' || r_Detail.No || ',' || r_Detail.序号 || ':' || v_当前数量;
              Else
                If Not (r_Detail.收费类别 = '7' And r_Detail.执行标志 <> 0) Then
                  Zl_病人费用销帐_Insert(r_Detail.费用id, r_Detail.收费细目id, r_Detail.病人病区id, v_当前数量, v_人员姓名, 收回时间_In,
                                   r_Detail.执行标志, Null, Null, v_销帐原因);
                End If;
              End If;
              v_费用ids := v_费用ids || ',' || r_Detail.费用id;
            End If;
          End If;
          v_收费细目id := r_Detail.收费细目id;
        End Loop;
      
        --数量未分摊完成
        If v_收回数量 > 0 Then
          v_Error := '医嘱"' || v_医嘱内容 || '"对应的费用剩余数量不足收回数量，可能存在手工销帐或已结帐的费用，或相应的划价单已被删除。';
          Raise Err_Custom;
        End If;
        --本科的销帐申请自动审核
        If zl_GetSysParameter('超期收回费用本科自动审核', 1254) = '1' And v_费用ids Is Not Null Then
          For r_Applay In c_Applay(Substr(v_费用ids, 2)) Loop
            Zl_病人费用销帐_Audit(r_Applay.费用id, r_Applay.申请时间, v_人员姓名, 收回时间_In, 1, 1, r_Applay.申请类别);
            v_Delno := v_Delno || '|' || r_Applay.No || ',' || r_Applay.序号 || ':' || r_Applay.数量;
          End Loop;
        End If;
      Else
        ---b.负数收回模式-------------------------------------------------------------------------------------------------------
        --如果全是划价单，就不用产生负数冲销单据
        If No_In = '调整划价单' Then
          --未审核的划价单，先进行修改或删除，可能多次发送为不同的NO,为了计算每次的收回量，需要按收费细目ID排序
          For r_Price In (Select c.诊疗类别, b.No, b.序号, b.收费细目id, Nvl(b.付数, 1) * b.数次 As 剩余数量, c.单次用量, d.剂量系数, d.住院包装,
                                 c.医嘱内容, Nvl(e.收费方式, 0) As 收费方式
                          From 病人医嘱发送 A, 住院费用记录 B, 病人医嘱记录 C, 药品规格 D, 病人医嘱计价 E
                          Where a.医嘱id = 医嘱id_In And a.No = b.No And a.记录性质 = b.记录性质 And a.医嘱id = b.医嘱序号 And
                                b.价格父号 Is Null And b.收费细目id = d.药品id(+) And b.记录状态 = 0 And c.Id = a.医嘱id And
                                b.医嘱序号 = e.医嘱id(+) And b.收费细目id = e.收费细目id(+)
                          Order By 收费细目id, NO Desc) Loop
            If Nvl(v_收费细目id, 0) <> r_Price.收费细目id Then
              --数量未分摊完成
              If v_收费细目id Is Not Null And v_收回数量 > 0 Then
                v_Error := '医嘱"' || v_医嘱内容 || '"对应的费用剩余数量不足收回数量，可能相关划价单已被删除或审核。';
                Raise Err_Custom;
              End If;
              --计算收回数量，如果收费方式不是0正常收取，则使用特殊的方法进行计算
              If r_Price.收费方式 = 0 Then
                If r_Price.诊疗类别 = '7' Then
                  --中药配方药品：付数*单量
                  v_收回数量 := Round(v_输液收回剩余 * r_Price.单次用量 / Nvl(r_Price.剂量系数, 1), 5);
                Else
                  If r_Price.诊疗类别 Not In ('5', '6') Then
                    Select Nvl(Max(数量), 1)
                    Into v_对照数量
                    From 病人医嘱计价
                    Where 医嘱id = 医嘱id_In And 收费细目id = r_Price.收费细目id;
                  Else
                    v_对照数量 := 1;
                  End If;
                  v_收回数量 := Round(v_输液收回剩余 * Nvl(r_Price.住院包装, 1), 5) * v_对照数量;
                End If;
              Else
                Select Nvl(Sum(数量), 0)
                Into v_收回数量
                From 医嘱执行计价
                Where 医嘱id = 医嘱id_In And 收费细目id = r_Price.收费细目id And
                      要求时间 > Nvl(上次时间_In, To_Date('1900-01-01', 'yyyy-MM-dd'));
              
                v_收回数量 := Round(v_收回数量, 5);
              End If;
              v_医嘱内容 := r_Price.医嘱内容;
            End If;
            If v_收回数量 > 0 Then
              If v_收回数量 > r_Price.剩余数量 Then
                v_当前数量 := r_Price.剩余数量;
              Else
                v_当前数量 := v_收回数量;
              End If;
              v_收回数量 := v_收回数量 - v_当前数量;
              v_Delno    := v_Delno || '|' || r_Price.No || ',' || r_Price.序号 || ':' || v_当前数量;
            End If;
            v_收费细目id := r_Price.收费细目id;
          End Loop;
          --数量未分摊完成
          If v_收回数量 > 0 Then
            v_Error := '医嘱"' || v_医嘱内容 || '"对应的费用剩余数量不足收回数量，可能相关划价单已被删除或审核。';
            Raise Err_Custom;
          End If;
        Else
          --负数冲销，可能存在划价单与记帐单混合的情况
          --金额小数位数
          Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into v_Dec From Dual;
          --生成划价单系统参数
          Select zl_GetSysParameter(80) Into v_划价类别 From Dual;
          v_开始序号 := Null;
          v_结束序号 := Null;
        
          Select a.诊疗类别, a.单次用量, b.跟踪在用
          Into v_诊疗类别, v_单次用量, v_跟踪在用
          From 病人医嘱记录 A, 材料特性 B
          Where ID = 医嘱id_In And a.收费细目id = b.材料id(+);
        
          If v_诊疗类别 In ('5', '6', '7') Or (v_诊疗类别 = '4' And Nvl(v_跟踪在用, 0) = 1) Then
            --药品、卫材
            -----------------------------------------------------------------------------------------------------
            v_收回数量 := Null;
            Select Nvl(Max(序号), 0) + 1
            Into v_收发序号
            From 药品收发记录
            Where 单据 In (9, 10, 25, 26) And 记录状态 = 1 And NO = No_In;
            Select Nvl(Max(序号), 0) + 1
            Into v_费用序号
            From 住院费用记录
            Where 记录性质 = 2 And 记录状态 In (0, 1) And NO = No_In;
          
            --一条医嘱的药品只有一行，这里的循环是为了处理多次发送的情况，分批药品在界面已禁用负数收回
            For r_Drug In c_Drug Loop
              --初始化要收回的总数量(零售数量)
              v_First := 0;
              If v_收回数量 Is Null Then
                If v_诊疗类别 = '7' Then
                  v_收回数量 := Round(v_输液收回剩余 * v_单次用量 / r_Drug.剂量系数, 5);
                Else
                  If v_诊疗类别 Not In ('5', '6') Then
                    Select Nvl(Max(数量), 1)
                    Into v_对照数量
                    From 病人医嘱计价
                    Where 医嘱id = 医嘱id_In And 收费细目id = r_Drug.收费细目id;
                  Else
                    v_对照数量 := 1;
                  End If;
                  v_收回数量 := Round(v_输液收回剩余 * r_Drug.住院包装, 5) * v_对照数量;
                End If;
                v_First := 1;
              End If;
            
              --如果第一次数量就足够，则按付数处理，否则付数不好处理
              If v_收回数量 > r_Drug.数量 Then
                v_当前付数 := 1;
                v_当前数量 := r_Drug.数量;
                v_收回数量 := v_收回数量 - r_Drug.数量;
              Else
                If v_First = 1 And v_诊疗类别 = '7' Then
                  v_当前付数 := v_输液收回剩余;
                  v_当前数量 := Round(v_单次用量 / r_Drug.剂量系数, 5);
                Else
                  v_当前付数 := 1;
                  v_当前数量 := v_收回数量;
                End If;
                v_收回数量 := 0;
              End If;
            
              If r_Drug.记录状态 = 0 Then
                v_Delno := v_Delno || '|' || r_Drug.No || ',' || r_Drug.序号 || ':' || v_当前数量 * v_当前付数;
              Else
                If Not (v_诊疗类别 = '7' And r_Drug.执行标志 <> 0) Then
                
                  Select 病人费用记录_Id.Nextval Into v_费用id From Dual;
                  负数收发记录_Insert(v_费用id, r_Drug.批次, r_Drug.分批, r_Drug.批号, r_Drug.效期, r_Drug.最大效期, r_Drug.收发id,
                                r_Drug.病人id, r_Drug.主页id, r_Drug.库房id, r_Drug.单据, r_Drug.姓名, r_Drug.对方部门id, v_诊疗类别,
                                v_划价类别, v_当前付数, v_当前数量);
                
                  --住院费用记录
                  -------------------------------------------------------------------------------------
                  --记录序号范围以处理汇总表
                  If v_开始序号 Is Null Then
                    v_开始序号 := v_费用序号;
                  End If;
                  v_结束序号 := v_费用序号;
                
                  Insert Into 住院费用记录
                    (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 门诊标志, 病人id, 主页id, 标识号, 姓名, 性别, 年龄, 床号, 病人病区id, 病人科室id,
                     费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id, 付数, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额,
                     统筹金额, 记帐费用, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 医嘱序号, 划价人, 操作员编号, 操作员姓名)
                    Select v_费用id, 2, No_In, Decode(Nvl(Instr(v_划价类别, Decode(v_诊疗类别, '4', '4', '5')), 0), 0, 1, 0),
                           v_费用序号, Null, Null, 多病人单, 2, 病人id, 主页id, 标识号, 姓名, 性别, 年龄, 床号, 病人病区id, 病人科室id, 费别, 收费类别,
                           收费细目id, 计算单位, 保险项目否, 保险大类id, v_当前付数, -1 * v_当前数量, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价,
                           Round(-1 * v_当前付数 * v_当前数量 * 标准单价, v_Dec), Round(-1 * v_当前付数 * v_当前数量 * 标准单价, v_Dec), Null, 1,
                           开单部门id, 开单人, 收回时间_In, 收回时间_In, 执行部门id, 0, 医嘱序号, v_人员姓名,
                           Decode(Nvl(Instr(v_划价类别, Decode(v_诊疗类别, '4', '4', '5')), 0), 0, v_人员编号, Null),
                           Decode(Nvl(Instr(v_划价类别, Decode(v_诊疗类别, '4', '4', '5')), 0), 0, v_人员姓名, Null)
                    From 住院费用记录
                    Where ID = r_Drug.费用id;
                
                  Select Zl_Actualmoney(费别, 收费细目id, 收入项目id, 应收金额, 数次, 执行部门id)
                  Into v_Temp
                  From 住院费用记录
                  Where ID = v_费用id;
                  v_实收金额 := Round(Substr(v_Temp, Instr(v_Temp, ':') + 1), v_Dec);
                  Update 住院费用记录 A Set 实收金额 = v_实收金额 Where ID = v_费用id;
                
                  v_费用序号 := v_费用序号 + 1;
                End If;
                If v_收回数量 <= 0 Then
                  Exit;
                End If;
              End If;
            End Loop;
          
            If v_收回数量 <> 0 Then
              --没有收回所有数量,收发记录本身有问题(如记录不全或数量为负)
              Null;
            End If;
          Else
            --其它非药医嘱(包括给药途径，及绑定的卫材等)
            -----------------------------------------------------------------------------------------------------
            Select Nvl(Max(序号), 0) + 1
            Into v_收发序号
            From 药品收发记录
            Where 单据 In (9, 10, 25, 26) And 记录状态 = 1 And NO = No_In;
            --取费用序号
            Select Nvl(Max(序号), 0) + 1
            Into v_费用序号
            From 住院费用记录
            Where 记录性质 = 2 And 记录状态 In (0, 1) And NO = No_In;
          
            v_收回剩余 := v_输液收回剩余;
          
            For r_Othersend In (Select 发送号, 发送数次
                                From 病人医嘱发送
                                Where 医嘱id = 医嘱id_In And 末次时间 > Nvl(上次时间_In, To_Date('1900-01-01', 'yyyy-MM-dd'))
                                Order By 发送号 Desc) Loop
              If r_Othersend.发送数次 < v_收回剩余 Then
                --一次收回多次发送，但是每次发送费用有所变动（计价）
                v_收回剩余 := v_收回剩余 - r_Othersend.发送数次;
                v_收回量   := r_Othersend.发送数次;
              Else
                --一次发送中收回剩余；
                v_收回量   := v_收回剩余;
                v_收回剩余 := 0;
              End If;
              v_收费内容 := '';
              For r_Other In c_Other(r_Othersend.发送号) Loop
                If Nvl(v_收费内容, '0') <> r_Other.收费细目id || ',' || r_Other.序号 Then
                  --根据最近一次发送的费用记录，按需要收回的数量全部收回
                  --计算收回数量，如果收费方式不是0正常收取，则使用特殊的方法进行计算
                  If r_Other.收费方式 = 0 Then
                    v_收回数量 := v_收回量 * Nvl(r_Other.对照数量, 1);
                  Else
                    Select Nvl(Sum(数量), 0)
                    Into v_收回数量
                    From 医嘱执行计价
                    Where 医嘱id = 医嘱id_In And 收费细目id = r_Other.收费细目id And
                          要求时间 > Nvl(上次时间_In, To_Date('1900-01-01', 'yyyy-MM-dd'));
                  End If;
                End If;
              
                If v_收回数量 > 0 Then
                  If r_Other.记录状态 = 0 Then
                    If v_收回数量 > r_Other.剩余数量 Then
                      v_当前数量 := r_Other.剩余数量;
                    Else
                      v_当前数量 := v_收回数量;
                    End If;
                  Else
                    v_当前数量 := v_收回数量;
                  End If;
                  v_收回数量 := v_收回数量 - v_当前数量;
                  v_当前付数 := 1;
                
                  If r_Other.记录状态 = 0 Then
                    v_Delno := v_Delno || '|' || r_Other.No || ',' || r_Other.序号 || ':' || v_当前数量;
                  Else
                    --记录序号范围以处理汇总表
                    If v_开始序号 Is Null Then
                      v_开始序号 := v_费用序号;
                    End If;
                    v_结束序号 := v_费用序号;
                  
                    --住院费用记录:按理如果收回量大于了上次发送量,则不正确，跨批次的卫材可能有多条收发记录
                    Select 病人费用记录_Id.Nextval Into v_费用id From Dual;
                    If r_Other.收费类别 In ('4', '5', '6', '7') Then
                      n_数量 := v_当前数量;
                      For r_Otherdrug In (Select a.病人id, a.主页id, d.姓名, Nvl(Nvl(x.剂量系数, y.换算系数), 1) As 剂量系数,
                                                 Nvl(x.住院包装, 1) As 住院包装, Nvl(x.最大效期, y.最大效期) As 最大效期,
                                                 Nvl(b.付数, 1) * b.实际数量 As 数量, b.Id As 收发id, b.单据, b.药品id, b.对方部门id,
                                                 b.库房id, b.费用id, Nvl(Nvl(x.药房分批, y.在用分批), 0) As 分批, b.批次, b.批号, b.效期,
                                                 a.记录状态, a.No, a.序号, a.收费细目id
                                          From 住院费用记录 A, 药品收发记录 B, 病人信息 D, 药品规格 X, 材料特性 Y
                                          Where a.Id = r_Other.费用id And a.记录状态 In (0, 1, 3) And a.No = b.No And
                                                a.Id = b.费用id + 0 And b.单据 In (9, 10, 25, 26) And
                                                (b.记录状态 = 1 Or Mod(b.记录状态, 3) = 0) And a.病人id = d.病人id And
                                                b.药品id = x.药品id(+) And b.药品id = y.材料id(+)
                                          Order By a.记录状态, b.No Desc, b.Id Desc) Loop
                        If n_数量 > 0 Then
                          n_Count := r_Otherdrug.数量;
                          If n_数量 < n_Count Then
                            n_Count := n_数量;
                          End If;
                          负数收发记录_Insert(v_费用id, r_Otherdrug.批次, r_Otherdrug.分批, r_Otherdrug.批号, r_Otherdrug.效期,
                                        r_Otherdrug.最大效期, r_Otherdrug.收发id, r_Otherdrug.病人id, r_Otherdrug.主页id,
                                        r_Otherdrug.库房id, r_Otherdrug.单据, r_Otherdrug.姓名, r_Otherdrug.对方部门id,
                                        r_Other.收费类别, v_划价类别, 1, n_Count);
                          n_数量 := n_数量 - r_Otherdrug.数量;
                        End If;
                      End Loop;
                    End If;
                    --医嘱已执行，收回的费用也填为已执行：不包含药品和跟踪在用的卫材，因为实际发放表示执行
                    Insert Into 住院费用记录
                      (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 门诊标志, 病人id, 主页id, 标识号, 姓名, 性别, 年龄, 床号, 病人病区id, 病人科室id,
                       费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id, 付数, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额,
                       统筹金额, 记帐费用, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 执行时间, 执行人, 医嘱序号, 划价人, 操作员编号, 操作员姓名)
                      Select v_费用id, 2, No_In, Decode(Nvl(Instr(v_划价类别, r_Other.收费类别), 0), 0, 1, 0), v_费用序号, Null,
                             Decode(a.价格父号, Null, Null, v_费用序号 + a.价格父号 - a.序号), a.多病人单, 2, a.病人id, a.主页id, a.标识号, a.姓名,
                             a.性别, a.年龄, a.床号, a.病人病区id, a.病人科室id, a.费别, a.收费类别, a.收费细目id, a.计算单位, a.保险项目否, a.保险大类id, 1,
                             -1 * v_当前数量, a.加班标志, a.附加标志, a.婴儿费, a.收入项目id, a.收据费目, a.标准单价,
                             Round(-1 * v_当前数量 * a.标准单价, v_Dec), Round(-1 * v_当前数量 * a.标准单价, v_Dec), Null, 1, a.开单部门id,
                             a.开单人, 收回时间_In, 收回时间_In, a.执行部门id,
                             Decode(r_Other.执行状态, 1,
                                     Decode(a.收费类别, '4', Decode(b.跟踪在用, 1, 0, 1),
                                             Decode(Instr(',5,6,7,', a.收费类别), 0, 1, 0)), 0),
                             Decode(r_Other.执行状态, 1,
                                     Decode(a.收费类别, '4', Decode(b.跟踪在用, 1, Null, 收回时间_In),
                                             Decode(Instr(',5,6,7,', a.收费类别), 0, 收回时间_In, Null)), Null),
                             Decode(r_Other.执行状态, 1,
                                     Decode(a.收费类别, '4', Decode(b.跟踪在用, 1, Null, v_人员姓名),
                                             Decode(Instr(',5,6,7,', a.收费类别), 0, v_人员姓名, Null)), Null), a.医嘱序号, v_人员姓名,
                             Decode(Nvl(Instr(v_划价类别, r_Other.收费类别), 0), 0, v_人员编号, Null),
                             Decode(Nvl(Instr(v_划价类别, r_Other.收费类别), 0), 0, v_人员姓名, Null)
                      From 住院费用记录 A, 材料特性 B
                      Where a.Id = r_Other.费用id And a.收费细目id = b.材料id(+);
                  
                    Select Zl_Actualmoney(费别, 收费细目id, 收入项目id, 应收金额, 数次, 执行部门id)
                    Into v_Temp
                    From 住院费用记录
                    Where ID = v_费用id;
                    v_实收金额 := Round(Substr(v_Temp, Instr(v_Temp, ':') + 1), v_Dec);
                    Update 住院费用记录 A Set 实收金额 = v_实收金额 Where ID = v_费用id;
                  
                    v_费用序号 := v_费用序号 + 1;
                    v_医嘱执行 := r_Other.执行状态; --多个收费项目的执行状态是一样的
                  End If;
                
                  v_收费内容 := r_Other.收费细目id || ',' || r_Other.序号;
                End If;
              End Loop;
              If v_收回剩余 = 0 Then
                Exit;
              End If;
            End Loop;
          
            --如果医嘱已执行，则按系统参数执行后自动审核费用：包含已执行医嘱对应的药品和卫材费用。
            -----------------------------------------------------------------------------------------------------
            If Nvl(v_医嘱执行, 0) = 1 And v_开始序号 Is Not Null And v_结束序号 Is Not Null Then
              For r_Verify In c_Verify(v_开始序号, v_结束序号) Loop
                Zl_住院记帐记录_Verify(r_Verify.No, v_人员编号, v_人员姓名, r_Verify.序号, Null, 收回时间_In);
              End Loop;
            End If;
          End If;
        
          --处理费用汇总表
          -----------------------------------------------------------------------------------------------------
          If v_开始序号 Is Not Null And v_结束序号 Is Not Null Then
            --最后统一处理费用相关汇总表
            For r_Money In c_Money(v_开始序号, v_结束序号) Loop
              --病人余额
              Update 病人余额
              Set 费用余额 = Nvl(费用余额, 0) + r_Money.实收金额
              Where 病人id = r_Money.病人id And 性质 = 1 And 类型 = 2;
            
              If Sql%RowCount = 0 Then
                Insert Into 病人余额
                  (病人id, 性质, 类型, 费用余额, 预交余额)
                Values
                  (r_Money.病人id, 1, 2, r_Money.实收金额, 0);
              End If;
            
              --病人未结费用
              Update 病人未结费用
              Set 金额 = Nvl(金额, 0) + r_Money.实收金额
              Where 病人id = r_Money.病人id And 主页id = r_Money.主页id And Nvl(病人病区id, 0) = Nvl(r_Money.病人病区id, 0) And
                    Nvl(病人科室id, 0) = Nvl(r_Money.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(r_Money.开单部门id, 0) And
                    Nvl(执行部门id, 0) = Nvl(r_Money.执行部门id, 0) And 收入项目id + 0 = r_Money.收入项目id And 来源途径 + 0 = 2;
            
              If Sql%RowCount = 0 Then
                Insert Into 病人未结费用
                  (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
                Values
                  (r_Money.病人id, r_Money.主页id, r_Money.病人病区id, r_Money.病人科室id, r_Money.开单部门id, r_Money.执行部门id,
                   r_Money.收入项目id, 2, r_Money.实收金额);
              End If;
            End Loop;
          End If;
        End If;
      End If;
    End If;
  End If;

  --过程Zl_住院记帐记录_Delete，不支持每次删除一行的循环处理（序号重整），必须把一个单据要删除的序号一次性传入
  If Not v_Delno Is Null Then
    v_Temp := '';
    v_No   := '';
    For r_Price In (Select /*+ rule*/
                     C1 As NO, C2 As 序号数量
                    From Table(f_Str2list2(Substr(v_Delno, 2), '|', ','))
                    Order By NO) Loop
      If v_No Is Not Null And v_No <> r_Price.No Then
        Zl_住院记帐记录_Delete(v_No, v_Temp, v_人员编号, v_人员姓名, 2);
        v_No := '';
      End If;
      If v_No Is Null Then
        v_No   := r_Price.No;
        v_Temp := r_Price.序号数量;
      Else
        v_Temp := v_Temp || ',' || r_Price.序号数量;
      End If;
    End Loop;
    If Not v_No Is Null Then
      Zl_住院记帐记录_Delete(v_No, v_Temp, v_人员编号, v_人员姓名, 2);
    End If;
  End If;

  --处理医嘱的上次执行时间:给药途径等可能因为未发送而没调用收回过程。
  -----------------------------------------------------------------------------------------------------
  Select Nvl(相关id, ID) Into v_组id From 病人医嘱记录 Where ID = 医嘱id_In;
  Update 病人医嘱记录 Set 上次执行时间 = 上次时间_In Where ID = v_组id Or 相关id = v_组id;

  --删除医嘱执行时间
  If 上次时间_In Is Null Then
    --全部收回
    Delete From 医嘱执行时间 Where 医嘱id = v_组id;
    Delete From 医嘱执行计价 Where 医嘱id = 医嘱id_In;
  Else
    --可能收回多次发送的数据
    Delete From 医嘱执行时间 Where 医嘱id = v_组id And 要求时间 > 上次时间_In;
    Delete From 医嘱执行计价 Where 医嘱id = 医嘱id_In And 要求时间 > 上次时间_In;
  End If;
  --处理输液配液记录的批次问题，每个医嘱都进行调用，在过程里面只处理了输液配液的医嘱
  Zl_输液配药记录_批次调整(医嘱id_In);

  If zl_GetSysParameter(63) = '1' And Nvl(No_In, '调整划价单') <> '调整划价单' Then
    --处理跟踪在用卫料自动发料
    For r_Stuff In c_Stuff Loop
      If n_发料号 Is Null Then
        n_发料号 := Nextno(20);
      End If;
      If r_Stuff.库房id <> Nvl(n_库房id, 0) Then
        If Nvl(n_库房id, 0) <> 0 And v_收发ids Is Not Null Then
          v_收发ids := Substr(v_收发ids, 2);
          Zl_药品收发记录_批量发料(v_收发ids, n_库房id, v_人员姓名, Sysdate, 1, v_人员姓名, n_发料号, v_人员姓名);
        End If;
      
        n_库房id  := r_Stuff.库房id;
        v_收发ids := Null;
      End If;
    
      v_收发ids := v_收发ids || '|' || r_Stuff.Id || ',' || r_Stuff.批次;
    End Loop;
    If Nvl(n_库房id, 0) <> 0 And v_收发ids Is Not Null Then
      v_收发ids := Substr(v_收发ids, 2);
      Zl_药品收发记录_批量发料(v_收发ids, n_库房id, v_人员姓名, Sysdate, 1, v_人员姓名, n_发料号, v_人员姓名);
    End If;
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人医嘱记录_收回;
/

--124609:董露露,2018-12-12,诊断同步录入的问题
--124609:董露露,2018-12-13,诊断同步录入
CREATE OR REPLACE Procedure Zl_病人诊断记录_Insert
(
  病人id_In   病人诊断记录.病人id%Type,
  主页id_In   病人诊断记录.主页id%Type,
  记录来源_In 病人诊断记录.记录来源%Type,
  病历id_In   病人诊断记录.病历id%Type,
  诊断类型_In 病人诊断记录.诊断类型%Type,
  疾病id_In   病人诊断记录.疾病id%Type,
  诊断id_In   病人诊断记录.诊断id%Type,
  证候id_In   病人诊断记录.证候id%Type,
  诊断描述_In 病人诊断记录.诊断描述%Type,
  出院情况_In 病人诊断记录.出院情况%Type,
  是否未治_In 病人诊断记录.是否未治%Type,
  是否疑诊_In 病人诊断记录.是否疑诊%Type,
  记录日期_In 病人诊断记录.记录日期%Type,
  医嘱id_In   varchar2 := Null,
  诊断次序_In 病人诊断记录.诊断次序%Type := 1,
  备注_In     病人诊断记录.备注%Type := Null,
  入院病情_In 病人诊断记录.入院病情%Type := Null,
  发病时间_In 病人诊断记录.发病时间%Type := Null,
  记录人_In   病人诊断记录.记录人%Type := Null,
  Id_In       病人诊断记录.Id%Type := Null,
  附码id_In   病人诊断记录.疾病id%Type := Null
) Is
  --功能：插入病人诊断记录
  --医嘱id_In=与当前诊断相关联的，用","间隔的医嘱ID串
  v_诊断id 病人诊断记录.Id%Type;
  v_医嘱id 病人医嘱记录.Id%Type;

  v_病人科室id 病人信息.当前科室id%Type;
  v_经治医师   人员表.姓名%Type;
  v_编码       疾病编码目录.编码%Type;
  n_Count      Number;
  n_Mz         Number;
  v_发病时间   病人诊断记录.发病时间%Type;

  v_Temp     varchar2(255);
  v_人员姓名 人员表.姓名%Type;
  v_Error Varchar2(255); 
  Err_Custom Exception; 
Begin
  --当前操作人员
  If 记录人_In Is Null Then
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  Else
    v_人员姓名 := 记录人_In;
  End If;

  If Id_In Is Null Then
    Select 病人诊断记录_Id.Nextval Into v_诊断id From Dual;
  Else
    v_诊断id := Id_In;
  End If;

  v_医嘱id := Zl_To_Number(医嘱id_In);
  If v_医嘱id = 0 Then
    v_医嘱id := Null;
  End If;

  Select count(*) into n_Count from 病人诊断记录 where 病人id=病人id_In And 主页id=主页id_In And nvl(诊断id,0)=nvl(诊断id_In,0) And nvl(疾病id,0)=nvl(疾病id_In,0) And nvl(证候id,0)=nvl(证候id_In,0) And 诊断类型=诊断类型_In And 记录来源=记录来源_In And 诊断描述=诊断描述_In And 诊断次序=诊断次序_In; 
  Select Count(1) Into n_Mz From 病人挂号记录 Where 病人id = 病人id_In And ID = 主页id_In; 
  If n_Count=0 Or (n_Mz > 0 And Instr(',1,11,', 诊断类型_In) > 0) then 
	  Insert Into 病人诊断记录
	    (ID, 病人id, 主页id, 记录来源, 病历id, 诊断类型, 诊断次序, 疾病id, 诊断id, 证候id, 诊断描述, 入院病情, 出院情况, 是否未治, 是否疑诊, 
	         记录日期, 记录人, 医嘱id, 备注, 发病时间)
	  Values
	    (v_诊断id, 病人id_In, 主页id_In, 记录来源_In, 病历id_In, 诊断类型_In, 诊断次序_In, 疾病id_In, 诊断id_In, 证候id_In, 诊断描述_In, 入院病情_In, 出院情况_In,
	     是否未治_In, 是否疑诊_In, 记录日期_In, v_人员姓名, v_医嘱id, 备注_In, 发病时间_In);
  Else 
	  v_Error:='该病人已经存在相同的诊断内容，不能保存该诊断！'; 
	  Raise Err_Custom; 
  End If; 

 If 附码id_In Is Not Null Then
    Insert Into 病人诊断记录
      (Id, 病人id, 主页id, 记录来源, 病历id, 诊断类型, 诊断次序,疾病id, 诊断id, 证候id, 诊断描述, 入院病情, 出院情况, 是否未治, 是否疑诊, 
          记录日期, 记录人, 医嘱id, 备注, 发病时间,编码序号)
      Select 病人诊断记录_Id.Nextval, 病人id_In, 主页id_In, 记录来源_In, 病历id_In, 诊断类型_In, 诊断次序_In, 附码id_In, 诊断id_In, 证候id_In, 诊断描述_In,
           入院病情_In, 出院情况_In, 是否未治_In, 是否疑诊_In, 记录日期_In, v_人员姓名, v_医嘱id, 备注_In, 发病时间_In,2
      From Dual;
  End If;

  --如果是门诊第一诊断则更新病人挂号记录.发病时间
  v_发病时间 := 发病时间_In;
  If 诊断类型_In = 1 And 诊断次序_In = 1 Then
    If 发病时间_In Is Null Then
      --检查中医的发病时间，有则取中医的，否则清空
      Select Max(发病时间)
      Into v_发病时间
      From 病人诊断记录
      Where 病人id = 病人id_In And 主页id = 主页id_In And 诊断类型 = 11 And 诊断次序 = 1;
    End If;
    If v_发病时间 Is Null Then
      --如果都为NULL，则取挂号记录中的
      Select Max(发病时间) Into v_发病时间 From 病人挂号记录 Where 病人id = 病人id_In And ID = 主页id_In;
    End If;
    Update 病人挂号记录 Set 发病时间 = v_发病时间 Where 病人id = 病人id_In And ID = 主页id_In;
  End If;
  If 诊断类型_In = 11 And 诊断次序_In = 1 Then
    --如果是中医，则判断是否填写了西医的发病时间，没有填写，则修改，否则以西医发病时间为准
    Select Count(*)
    Into n_Count
    From 病人诊断记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 诊断类型 = 1 And 诊断次序 = 1 And 发病时间 Is Not Null;
    If n_Count = 0 Then
      If v_发病时间 Is Null Then
        --如果都为NULL，则取挂号记录中的
        Select Max(发病时间) Into v_发病时间 From 病人挂号记录 Where 病人id = 病人id_In And ID = 主页id_In;
      End If;
      Update 病人挂号记录 Set 发病时间 = v_发病时间 Where 病人id = 病人id_In And ID = 主页id_In;
    End If;
  End If;

  If 医嘱id_In Is Not Null Then
    For r_Advice In (Select Column_Value As 医嘱id
                    From Table(Cast(f_Num2list(医嘱id_In) As Zltools.t_Numlist)) A, 病人医嘱记录 B
                    Where a.Column_Value = b.Id) Loop 
      Insert Into 病人诊断医嘱 (诊断id, 医嘱id) Values (v_诊断id, r_Advice.医嘱id);
    End Loop;
  End If;

  --如果是入院第一诊断，则判断是否是单病种
  If 诊断类型_In = 2 And 诊断次序_In = 1 And 记录来源_In = 3 Then
    If 疾病id_In Is Not Null Then
      Select 编码 Into v_编码 From 疾病编码目录 Where ID = 疾病id_In;
      Select Max(Upper(编码))
      Into v_编码
      From 单病种目录
      Where Instr('/' || Replace(Upper(Icd编码), ' ', '') || '/', '/' || Upper(v_编码) || '/') > 0 And Rownum < 2;
    Else
      v_编码 := '';
    End If;
    Update 病案主页 Set 单病种 = v_编码 Where 病人id = 病人id_In And 主页id = 主页id_In;
  End If;

  --根据传入的主页id_In查询挂号记录来区分是门诊首页还是住院首页调用
  Begin
    Select 执行人, 执行部门id Into v_病人科室id, v_经治医师 From 病人挂号记录 Where ID = 主页id_In;
    Zl_电子病历时机_Insert(病人id_In, 主页id_In, 1, '诊断', v_病人科室id, v_经治医师, 记录日期_In, 记录日期_In);
  Exception
    When Others Then
      Null;
  End;
  If v_病人科室id Is Null And (诊断类型_In <> 1 Or 诊断类型_In <> 11) Then
    Begin
      Select 出院科室id, 住院医师
      Into v_病人科室id, v_经治医师
      From 病案主页
      Where 病人id = 病人id_In And 主页id = 主页id_In;
      Zl_电子病历时机_Insert(病人id_In, 主页id_In, 2, '诊断', v_病人科室id, v_经治医师, 记录日期_In, 记录日期_In);
    Exception
      When Others Then
        Null;
    End;
  End If;
  b_Message.Zlhis_Cis_010(病人id_In, 主页id_In, v_诊断id);
Exception
  When Err_Custom Then 
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人诊断记录_Insert;
/





------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0041' Where 编号=&n_System;
Commit;
