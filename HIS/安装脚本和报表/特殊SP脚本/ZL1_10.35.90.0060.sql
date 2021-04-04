----------------------------------------------------------------------------------------------------------------
--本脚本支持从ZLHIS+ v10.35.90升级到 v10.35.90
--请以数据空间的所有者登录PLSQL并执行下列脚本
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------
--结构修正部份
------------------------------------------------------------------------------
--140325:刘涛,2019-05-06,应付记录增加读库单据号的索引
Create Index 应付记录_IX_入库单据号 On 应付记录(入库单据号) Tablespace zl9Indexhis;




------------------------------------------------------------------------------
--数据修正部份
------------------------------------------------------------------------------
--140678:刘硕,2019-05-07,上机人员变动消息通知
Insert Into Zlmsg_Lists(Bz_Type, Code, Name, Key_Define, Note)
Select '管理工具', 'ZLTOOLS_USERS_001', '上机人员删除', '<root><用户名></用户名><人员ID></人员ID></root>', '用户授权管理:修改用户、删除用户时'  From Dual Union All 
Select '管理工具', 'ZLTOOLS_USERS_002', '上机人员新增', '<root><用户名></用户名><人员ID></人员ID></root>', '用户授权管理:新增用户、修改用户时;批量创建用户'  From Dual;


-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--139097:胡俊勇,2019-04-28,门诊留观病人处理
Create Or Replace Procedure Zl_病人医嘱记录_收回
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
  Select Count(1)
  Into n_Count
  From 输液配药记录 A, 病人医嘱记录 B
  Where a.医嘱id = b.Id And 医嘱id = 医嘱id_In And a.执行时间 > b.执行终止时间 And a.是否锁定 = 1;

  If n_Count > 0 Then
    v_Error := '医嘱"' || v_医嘱内容 || '"是输液药品，已经被输液配置中心锁定，不能超期收回。';
    Raise Err_Custom;
  End If;

  Select a.医嘱内容, Nvl(a.相关id, a.Id), b.病人性质
  Into v_医嘱内容, n_相关id, n_Count
  From 病人医嘱记录 A, 病案主页 B
  Where a.病人id = b.病人id And a.主页id = b.主页id And a.Id = 医嘱id_In;
  --判断门诊留观病人执行单独的过程维护时请同步修改
  If n_Count = 1 Then
    Zl_病人医嘱记录_收回_门诊留观(收回量_In, 医嘱id_In, 上次时间_In, 收回时间_In, No_In, 操作员编号_In, 操作员姓名_In);
    Return;
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

--139097:胡俊勇,2019-04-28,门诊留观病人处理
CREATE OR REPLACE Procedure Zl_病人医嘱记录_收回_门诊留观
(
  --功能：门诊留观病人超期收回内部逻辑和 Zl_病人医嘱记录_收回 相同,相关说明见这个过程
  --说明:Zl_病人医嘱记录_收回/Zl_病人医嘱记录_收回_门诊留观 查询的表不一样,门诊留观 统一都查询  门诊费用记录
  
  收回量_In     In 病人医嘱发送.发送数次%Type,
  医嘱id_In     In 病人医嘱记录.Id%Type,
  上次时间_In   In 病人医嘱记录.上次执行时间%Type,
  收回时间_In   In 病人医嘱记录.上次执行时间%Type,
  No_In         In 门诊费用记录.No%Type := Null,
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
                          From 病人医嘱发送 A, 门诊费用记录 B, 病人医嘱记录 C, 材料特性 D, 病人医嘱计价 E
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
    From 病人费用销帐 A, 门诊费用记录 B
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
    From 门诊费用记录 A, 药品收发记录 B, 病人医嘱发送 C, 病人信息 D, 药品规格 X, 材料特性 Y
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
      From 门诊费用记录 A, 病人医嘱发送 B, 病人医嘱计价 C
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
    v_Start 门诊费用记录.序号%Type,
    v_End   门诊费用记录.序号%Type
  ) Is
    Select 病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, Sum(Nvl(应收金额, 0)) As 应收金额, Sum(Nvl(实收金额, 0)) As 实收金额
    From 门诊费用记录
    Where 记录性质 = 2 And 记录状态 = 1 And NO = No_In And 序号 Between v_Start And v_End
    Group By 病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id;

  --系统参数指定执行后需要自动审核的划价费用：用于非药医嘱，包含对应的药品及卫材费用
  Cursor c_Verify
  (
    v_Start 门诊费用记录.序号%Type,
    v_End   门诊费用记录.序号%Type
  ) Is
    Select NO, 序号
    From 门诊费用记录
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
    From 输液配药内容 A, 药品收发记录 B, 药品规格 C, 收费项目目录 D, 输液配药记录 E, 门诊费用记录 F
    Where a.收发id = b.Id And b.药品id = c.药品id And c.药品id = d.Id And e.Id = a.记录id And f.No = b.No And f.Id = b.费用id And
          e.医嘱id = 相关id_In And e.执行时间 > 执行终止时间_In And e.Id = 配药id_In And f.医嘱序号 + 0 = 医嘱序号_In
    Group By b.费用id, b.药品id, c.住院包装, c.住院单位, d.名称, e.病人病区id, e.操作状态, e.Id, f.No, f.价格父号, f.序号, f.记录状态, f.执行状态;

  --审核中包含跟踪在用的未发卫料时，根据参数设置是否自动发料
  Cursor c_Stuff Is
    Select m.Id, m.库房id, Decode(b.在用分批, 1, m.批次, 0) As 批次
    From 药品收发记录 M, 门诊费用记录 A, 材料特性 B
    Where m.No = No_In And m.单据 In (25, 26) And m.库房id Is Not Null And m.记录状态 = 1 And m.审核人 Is Null And m.No = a.No And
          a.Id = m.费用id + 0 And a.记录性质 = 2 And a.记录状态 = 1 And a.收费细目id = b.材料id And b.跟踪在用 = 1
    Order By m.库房id, m.药品id;

  v_Dec      Number;
  v_First    Number;
  v_划价类别 Varchar2(255);

  v_诊疗类别 病人医嘱记录.诊疗类别%Type;
  v_单次用量 病人医嘱记录.单次用量%Type;
  v_跟踪在用 材料特性.跟踪在用%Type;

  v_费用序号 门诊费用记录.序号%Type;
  v_收发序号 药品收发记录.序号%Type;
  v_费用id   门诊费用记录.Id%Type;
  v_实收金额 门诊费用记录.实收金额%Type;

  v_开始序号 门诊费用记录.序号%Type;
  v_结束序号 门诊费用记录.序号%Type;

  v_医嘱执行 病人医嘱发送.执行状态%Type;

  v_剂量系数 药品规格.剂量系数%Type;
  v_住院包装 药品规格.住院包装%Type;
  v_医嘱内容 病人医嘱记录.医嘱内容%Type;

  v_结帐参数       Zlparameters.参数值%Type;
  v_配液药销帐申请 Zlparameters.参数值%Type;
  v_结帐金额       门诊费用记录.结帐金额%Type;

  v_收费细目id   门诊费用记录.收费细目id%Type;
  v_剩余数量     门诊费用记录.数次%Type;
  v_收回数量     门诊费用记录.数次%Type;
  v_当前数量     门诊费用记录.数次%Type;
  v_当前付数     门诊费用记录.付数%Type;
  v_费用ids      Varchar2(4000);
  v_组id         病人医嘱记录.Id%Type;
  v_对照数量     病人医嘱计价.数量%Type;
  v_收回量       门诊费用记录.数次%Type;
  v_收回剩余     门诊费用记录.数次%Type;
  v_输液收回剩余 门诊费用记录.数次%Type;
  n_数量         药品收发记录.填写数量%Type;

  v_Delno    Varchar2(4000);
  v_Temp     Varchar2(4000);
  v_收费内容 Varchar2(4000);
  v_No       门诊费用记录.No%Type;
  v_人员编号 门诊费用记录.操作员编号%Type;
  v_人员姓名 门诊费用记录.操作员姓名%Type;

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
    病人id_In     门诊费用记录.病人id%Type,
    主页id_In     门诊费用记录.主页id%Type,
    库房id_In     药品收发记录.库房id%Type,
    单据_In       药品收发记录.单据%Type,
    姓名_In       病人信息.姓名%Type,
    对方部门id_In 药品收发记录.对方部门id%Type,
    收费类别_In   门诊费用记录.收费类别%Type,
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
  Select a.医嘱内容, Nvl(a.相关id, a.Id) Into v_医嘱内容, n_相关id From 病人医嘱记录 A Where a.Id = 医嘱id_In;
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
                Zl_门诊记帐记录_Delete(r_Compound.No, r_Compound.序号 || ':' || r_Compound.数量 || ':' || r_Compound.配药id, v_人员编号,
                                 v_人员姓名, 0, d_收回时间);
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
              From 门诊费用记录
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
                          From 病人医嘱发送 A, 门诊费用记录 B, 病人医嘱记录 C, 药品规格 D, 病人医嘱计价 E
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
            From 门诊费用记录
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
                
                  --门诊费用记录
                  -------------------------------------------------------------------------------------
                  --记录序号范围以处理汇总表
                  If v_开始序号 Is Null Then
                    v_开始序号 := v_费用序号;
                  End If;
                  v_结束序号 := v_费用序号;
                
                  Insert Into 门诊费用记录
                    (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 门诊标志, 病人id, 主页id, 标识号, 姓名, 性别, 年龄, 病人病区id, 病人科室id, 费别, 收费类别,
                     收费细目id, 计算单位, 保险项目否, 保险大类id, 付数, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额, 记帐费用,
                     开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 医嘱序号, 划价人, 操作员编号, 操作员姓名)
                    Select v_费用id, 2, No_In, Decode(Nvl(Instr(v_划价类别, Decode(v_诊疗类别, '4', '4', '5')), 0), 0, 1, 0),
                           v_费用序号, Null, Null, 1, 病人id, 主页id, 标识号, 姓名, 性别, 年龄, 病人病区id, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位,
                           保险项目否, 保险大类id, v_当前付数, -1 * v_当前数量, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价,
                           Round(-1 * v_当前付数 * v_当前数量 * 标准单价, v_Dec), Round(-1 * v_当前付数 * v_当前数量 * 标准单价, v_Dec), Null, 1,
                           开单部门id, 开单人, 收回时间_In, 收回时间_In, 执行部门id, 0, 医嘱序号, v_人员姓名,
                           Decode(Nvl(Instr(v_划价类别, Decode(v_诊疗类别, '4', '4', '5')), 0), 0, v_人员编号, Null),
                           Decode(Nvl(Instr(v_划价类别, Decode(v_诊疗类别, '4', '4', '5')), 0), 0, v_人员姓名, Null)
                    From 门诊费用记录
                    Where ID = r_Drug.费用id;
                
                  Select Zl_Actualmoney(费别, 收费细目id, 收入项目id, 应收金额, 数次, 执行部门id)
                  Into v_Temp
                  From 门诊费用记录
                  Where ID = v_费用id;
                  v_实收金额 := Round(Substr(v_Temp, Instr(v_Temp, ':') + 1), v_Dec);
                  Update 门诊费用记录 A Set 实收金额 = v_实收金额 Where ID = v_费用id;
                
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
            From 门诊费用记录
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
                  
                    --门诊费用记录:按理如果收回量大于了上次发送量,则不正确，跨批次的卫材可能有多条收发记录
                    Select 病人费用记录_Id.Nextval Into v_费用id From Dual;
                    If r_Other.收费类别 In ('4', '5', '6', '7') Then
                      n_数量 := v_当前数量;
                      For r_Otherdrug In (Select a.病人id, a.主页id, d.姓名, Nvl(Nvl(x.剂量系数, y.换算系数), 1) As 剂量系数,
                                                 Nvl(x.住院包装, 1) As 住院包装, Nvl(x.最大效期, y.最大效期) As 最大效期,
                                                 Nvl(b.付数, 1) * b.实际数量 As 数量, b.Id As 收发id, b.单据, b.药品id, b.对方部门id,
                                                 b.库房id, b.费用id, Nvl(Nvl(x.药房分批, y.在用分批), 0) As 分批, b.批次, b.批号, b.效期,
                                                 a.记录状态, a.No, a.序号, a.收费细目id
                                          From 门诊费用记录 A, 药品收发记录 B, 病人信息 D, 药品规格 X, 材料特性 Y
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
                    Insert Into 门诊费用记录
                      (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 门诊标志, 病人id, 主页id, 标识号, 姓名, 性别, 年龄, 病人病区id, 病人科室id, 费别, 收费类别,
                       收费细目id, 计算单位, 保险项目否, 保险大类id, 付数, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额, 记帐费用,
                       开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 执行时间, 执行人, 医嘱序号, 划价人, 操作员编号, 操作员姓名)
                      Select v_费用id, 2, No_In, Decode(Nvl(Instr(v_划价类别, r_Other.收费类别), 0), 0, 1, 0), v_费用序号, Null,
                             Decode(a.价格父号, Null, Null, v_费用序号 + a.价格父号 - a.序号), 1, a.病人id, a.主页id, a.标识号, a.姓名, a.性别,
                             a.年龄, a.病人病区id, a.病人科室id, a.费别, a.收费类别, a.收费细目id, a.计算单位, a.保险项目否, a.保险大类id, 1, -1 * v_当前数量,
                             a.加班标志, a.附加标志, a.婴儿费, a.收入项目id, a.收据费目, a.标准单价, Round(-1 * v_当前数量 * a.标准单价, v_Dec),
                             Round(-1 * v_当前数量 * a.标准单价, v_Dec), Null, 1, a.开单部门id, a.开单人, 收回时间_In, 收回时间_In, a.执行部门id,
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
                      From 门诊费用记录 A, 材料特性 B
                      Where a.Id = r_Other.费用id And a.收费细目id = b.材料id(+);
                  
                    Select Zl_Actualmoney(费别, 收费细目id, 收入项目id, 应收金额, 数次, 执行部门id)
                    Into v_Temp
                    From 门诊费用记录
                    Where ID = v_费用id;
                    v_实收金额 := Round(Substr(v_Temp, Instr(v_Temp, ':') + 1), v_Dec);
                    Update 门诊费用记录 A Set 实收金额 = v_实收金额 Where ID = v_费用id;
                  
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
            --If Nvl(v_医嘱执行, 0) = 1 And v_开始序号 Is Not Null And v_结束序号 Is Not Null Then
            -- For r_Verify In c_Verify(v_开始序号, v_结束序号) Loop
            --   Zl_住院记帐记录_Verify(r_Verify.No, v_人员编号, v_人员姓名, r_Verify.序号, Null, 收回时间_In);
            -- End Loop;
            --End If;
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

  --过程Zl_门诊记帐记录_Delete，不支持每次删除一行的循环处理（序号重整），必须把一个单据要删除的序号一次性传入
  If Not v_Delno Is Null Then
    v_Temp := '';
    v_No   := '';
    For r_Price In (Select /*+ rule*/
                     C1 As NO, C2 As 序号数量
                    From Table(f_Str2list2(Substr(v_Delno, 2), '|', ','))
                    Order By NO) Loop
      If v_No Is Not Null And v_No <> r_Price.No Then
        Zl_门诊记帐记录_Delete(v_No, v_Temp, v_人员编号, v_人员姓名, 0);
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
      Zl_门诊记帐记录_Delete(v_No, v_Temp, v_人员编号, v_人员姓名, 0);
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
End Zl_病人医嘱记录_收回_门诊留观;
/
--140678:刘硕,2019-05-07,上机人员变动消息通知
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
  --管理工具上机人员变动消息
  Procedure Zltools_Users_001
  (
    用户名_In In 上机人员表.用户名%Type,
    人员id_In In 上机人员表.人员id%Type
  );
  Procedure Zltools_Users_002
  (
    用户名_In In 上机人员表.用户名%Type,
    人员id_In In 上机人员表.人员id%Type
  );
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
  --管理工具上机人员变动消息
  Procedure Zltools_Users_001
  (
    用户名_In In 上机人员表.用户名%Type,
    人员id_In In 上机人员表.人员id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><用户名>' || 用户名_In || '</用户名><人员ID>' || 人员id_In || '</人员ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLTOOLS_USERS_001', v_Value);
  End Zltools_Users_001;
  Procedure Zltools_Users_002
  (
    用户名_In In 上机人员表.用户名%Type,
    人员id_In In 上机人员表.人员id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><用户名>' || 用户名_In || '</用户名><人员ID>' || 人员id_In || '</人员ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLTOOLS_USERS_002', v_Value);
  End Zltools_Users_002;
End b_Message;
/
------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0060' Where 编号=&n_System;
Commit;
