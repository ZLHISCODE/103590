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
--133839:焦博,2018-11-26,增加模块公共参数医生站挂号排序控制,用于控制医生站挂号时显示挂号排班的顺序
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
  Select Zlparameters_Id.Nextval, &n_System, 9000, 0, 0, 0, 0, 0, 0, 18, '医生站挂号排序控制', Null, '医生,1|执行时间,1|科室,1|号别,1|项目,1',
         '主要是针对医生站挂号时号源的排列顺序，该排列顺序，不针对“显示所有号别”。', '排序字段1 ，排序方式(0-DESC，1-ASC)|排序字段2 ，排序方式(0-DESC，1-ASC)|...', '',
         '适用于需要根据自身情况显示挂号排班顺序的情况', Null
  From Dual;

-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--133584:李南春,2018-11-30,利用失效时间查询挂号安排计划
Create Or Replace Procedure Zl_挂号安排计划_Verify
(
  Id_In         In 挂号安排计划.Id%Type,
  立即生效_In   Number := 0,
  上次计划ID_In In 挂号安排计划.上次计划Id%Type := Null
) Is
  Err_Item     Exception;
  v_Err_Msg    Varchar2(100);
  v_User_Name  人员表.姓名%Type;
  n_Valied     Number(1);
  d_生效时间   挂号安排计划.生效时间%Type;
  n_上次计划ID 挂号安排计划.ID%Type;
Begin
  
  Select Nvl(Max(p.姓名),'') Into v_User_Name From 上机人员表 o, 人员表 p Where o.人员id = p.Id And 用户名 = User;
  If v_User_Name Is Null Then
    v_Err_Msg := '[ZLSOFT]当前用户未设置对应的人员信息,请与' || Chr(10) || Chr(13) ||
                 '系统管理员联系,先到用户授权管理中设置！[ZLSOFT]';
    Raise Err_Item;
  End If;

  Select Count(1) Into n_Valied From 挂号安排计划 a Where Nvl(生效时间, Sysdate) < Sysdate And a.Id = Id_In And Rownum < 2;
  If n_Valied = 1 And Nvl(立即生效_In, 0) = 0 Then
    v_Err_Msg := '[ZLSOFT]该计划安排的生效时间已经到期，不能进行审核！[ZLSOFT]';
    Raise Err_Item;
  End If;
  
  if Nvl(上次计划ID_In, 0) = 0 Then
    Select Max(Id)
    Into n_上次计划id
    From (Select Max(Id) As Id, Max(失效时间) As 失效时间, Count(1) As Count
           From 挂号安排计划
           Where 安排id = (Select Max(安排id) From 挂号安排计划 Where Id = Id_In) And 审核时间 Is Not Null)
    Where Count = 1 And 失效时间 = To_Date('3000-01-01', 'yyyy-mm-dd');
  Else
    n_上次计划ID := 上次计划ID_In;
  End if;
  
  Update 挂号安排计划
  Set 审核人 = v_User_Name, 审核时间 = Sysdate, 上次计划ID = n_上次计划ID,
      生效时间 = Case Nvl(立即生效_In, 0) When 0 Then 生效时间 Else Sysdate - 1 / 24 / 60 / 60 End
  Where Id = Id_In And 审核时间 Is Null
  Return 生效时间 Into d_生效时间;
  If Sql%Notfound Then
    v_Err_Msg := '[ZLSOFT]该计划安排已经被他人审核或删除,不能再审核![ZLSOFT]';
    Raise Err_Item;
  End If;
  IF Nvl(n_上次计划ID, 0) <> 0 Then
    Update 挂号安排计划 Set 失效时间 = d_生效时间 Where ID = n_上次计划ID;
  End IF;
  If Nvl(立即生效_In, 0) = 1 Then
    Begin
      Zl_挂号安排_Autoupdate();
    Exception
      When Others Then
        Null;
    End;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_挂号安排计划_Verify;
/

--128110:胡俊勇,2018-11-30,材卫负数超期收回自动执行
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
    药品id_In     药品收发记录.药品id%Type,
    库房id_In     药品收发记录.库房id%Type,
    单据_In       药品收发记录.单据%Type,
    姓名_In       病人信息.姓名%Type,
    对方部门id_In 药品收发记录.对方部门id%Type,
    收费类别_In   住院费用记录.收费类别%Type,
    划价类别_In   Varchar
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
      Select v_Lngid, 1, 单据, No_In, v_收发序号, 库房id, 对方部门id, 入出类别id, -1, 药品id, Nvl(v_批次, 0), 产地, v_批号, v_效期, v_当前付数,
             -1 * v_当前数量, -1 * v_当前数量, 零售价, Round(-1 * v_当前付数 * v_当前数量 * 零售价, v_Dec), '超期发送收回', v_人员姓名, 收回时间_In, 费用id_In,
             单量, 频次, 用法, 供药单位id, 生产日期, 批准文号, 灭菌效期
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
                                r_Drug.病人id, r_Drug.主页id, r_Drug.药品id, r_Drug.库房id, r_Drug.单据, r_Drug.姓名, r_Drug.对方部门id,
                                v_诊疗类别, v_划价类别);
                
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
                  
                    --住院费用记录:按理如果收回量大于了上次发送量,则不正确
                    Select 病人费用记录_Id.Nextval Into v_费用id From Dual;
                    If r_Other.收费类别 In ('4', '5', '6', '7') Then
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
                        负数收发记录_Insert(v_费用id, r_Otherdrug.批次, r_Otherdrug.分批, r_Otherdrug.批号, r_Otherdrug.效期,
                                      r_Otherdrug.最大效期, r_Otherdrug.收发id, r_Otherdrug.病人id, r_Otherdrug.主页id,
                                      r_Otherdrug.药品id, r_Otherdrug.库房id, r_Otherdrug.单据, r_Otherdrug.姓名,
                                      r_Otherdrug.对方部门id, r_Other.收费类别, v_划价类别);
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
--000000:胡俊勇,2018-11-29,修正锚点消息问题
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
                  Zl_门诊记帐记录_Delete(v_费用no, v_费用序号, v_人员编号, v_人员姓名);
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
            Zl_门诊记帐记录_Delete(v_费用no, v_费用序号, v_人员编号, v_人员姓名);
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
--126851:李小东,2018-11-28,标本拒收后改变费用执行状态为0
Create Or Replace Procedure Zl_检验申请拒收_Update
(
  医嘱ids_In  In Varchar2, --多个医嘱ID用逗号分隔
  执行说明_In In 病人医嘱发送.执行说明%Type
) Is
  v_费用性质 病人医嘱发送.记录性质%Type;

  Cursor c_Samplequest Is
    Select Distinct ID As 医嘱id, 病人来源
    From 病人医嘱记录 A
    Where a.Id In (Select * From Table(Cast(f_Num2list(医嘱ids_In) As Zltools.t_Numlist)));
Begin
  --处理医嘱执行状态
  Update 病人医嘱发送
  Set 执行状态 = 2, 执行说明 = 执行说明_In, 采样人 = Null, 采样时间 = Null, 送检人 = Null, 标本送出时间 = Null, 标本发送批号 = Null, 接收人 = Null,
      接收时间 = Null
  Where 医嘱id In (Select * From Table(Cast(f_Num2list(医嘱ids_In) As Zltools.t_Numlist)));

  --处理费用执行状态
  For r_Samplequest In c_Samplequest Loop
    If r_Samplequest.病人来源 = 2 Then
      Select Decode(记录性质, 1, 1, Decode(门诊记帐, 1, 1, 2))
      Into v_费用性质
      From 病人医嘱发送
      Where 医嘱id = r_Samplequest.医嘱id;
    Else
      v_费用性质 := 1;
    End If;
  
    If v_费用性质 = 2 Then
      Update 住院费用记录
      Set 执行状态 = 0, 执行时间 = Null, 执行人 = Null
      Where 收费类别 Not In ('5', '6', '7') And
            (医嘱序号, 记录性质, NO) In
            (Select 医嘱id, 记录性质, NO
             From 病人医嘱附费
             Where 医嘱id = r_Samplequest.医嘱id
             Union All
             Select 医嘱id, 记录性质, NO
             From 病人医嘱发送
             Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = r_Samplequest.医嘱id And 相关id Is Not Null));
    Else
      Update 门诊费用记录
      Set 执行状态 = 0, 执行时间 = Null, 执行人 = Null
      Where 收费类别 Not In ('5', '6', '7') And
            (医嘱序号, 记录性质, NO) In
            (Select 医嘱id, 记录性质, NO
             From 病人医嘱附费
             Where 医嘱id = r_Samplequest.医嘱id
             Union All
             Select 医嘱id, 记录性质, NO
             From 病人医嘱发送
             Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = r_Samplequest.医嘱id And 相关id Is Not Null));
    End If;
  
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_检验申请拒收_Update;
/



------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0039' Where 编号=&n_System;
Commit;