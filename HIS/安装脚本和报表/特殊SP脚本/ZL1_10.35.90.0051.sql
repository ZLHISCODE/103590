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
--138241:刘兴洪,2019-02-28,条码卫材识别控制
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
  Select Zlparameters_Id.Nextval, &n_System, NULL, 0, 0, 0, 0, 0, 0, 320, '条码卫材识别控制', '0', '0',
          ' 如果参数设置条码卫材必须扫码录入，则在记帐、补费等模块中录入具体条码的卫生材料时，只能通过条码来录入;否则允许通过条码、简码、编码等进行录入。',
         '1-必须通过扫码录入或录入条码来认别卫生材料;0-不控制，可以通过简码、编码、条码等录入方式来识别卫生材料', 
		 '本参数设置为1-必须通过扫码录入或录入条码来认别卫生材料时，需要启用卫生材料条码管理才有效。',
         '适用于需要在入出通过条码严格管理的用户。', Null
  From Dual;

--123946:董露露,2019-02-25,解决将病人新生儿记录提取到分娩新界面上的问题
Insert into 分娩方式 values (3, '剖宫产', 'PGC', 0);

Insert into 分娩方式 values (4, '早产', 'ZC', 0);

Insert into 分娩方式 values (5, '产钳', 'CQ', 0);

Insert into 分娩方式 values (6, '臀抽', 'TC', 0);

Insert into 分娩方式 values (7, '臀助', 'TZ', 0);


-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------
--137272:李南春,2019-02-25,预约提前接收锁定序号
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1111, '基本', User, 'Zl_挂号序号状态_Lock', 'EXECUTE' From Dual;

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 9000, '基本', User, 'Zl_挂号序号状态_Lock', 'EXECUTE' From Dual;




-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--138040:冉俊明,2019-02-28,医保费用多单据一次结算部分退费，全退和重收的“门诊费用记录.是否上传”填写错误
Create Or Replace Procedure Zl_门诊收费记录_重收
(
  原结帐id_In     门诊费用记录.结帐id%Type,
  冲销id_In       门诊费用记录.结帐id%Type,
  重收结帐id_In   门诊费用记录.结帐id%Type,
  排开医保结算_In Varchar2 := Null
) As
  --排开医保结算_IN:多个用逗号分离(只某些医保结算,允许退现金)
  Cursor c_Fee_Data Is
    Select ID
    From 门诊费用记录 A
    Where 结帐id = 原结帐id_In And Not Exists
     (Select 1
           From 门诊费用记录 B
           Where Mod(b.记录性质, 10) = 1 And a.No = b.No And a.序号 = b.序号 And 结帐id = 冲销id_In)
    Order By ID;

  v_操作员编号 门诊费用记录.操作员编号%Type;
  v_操作员姓名 门诊费用记录.操作员姓名%Type;
  d_登记时间   门诊费用记录.登记时间%Type;
  n_缴款组id   门诊费用记录.缴款组id%Type;
  n_病人id     门诊费用记录.病人id%Type;
  Err_Item Exception;
  v_Err_Msg    Varchar2(255);
  n_Array_Size Number := 200;
  t_费用id     t_Numlist;
  n_结算金额   门诊费用记录.实收金额%Type;
  n_冲销金额   病人预交记录.冲预交%Type;
  n_Count      Number(18);
Begin
  Begin
    Select 操作员编号, 操作员姓名, 登记时间, 缴款组id, 病人id
    Into v_操作员编号, v_操作员姓名, d_登记时间, n_缴款组id, n_病人id
    From 门诊费用记录
    Where 结帐id = 冲销id_In And Rownum < 2;
  Exception
    When Others Then
      v_Err_Msg := 'NO';
  End;

  If Nvl(v_Err_Msg, '-') = 'NO' Then
    v_Err_Msg := '由于并发操作,该单据可能已经初他人退费或删除,不能再进行退费操作！';
    Raise Err_Item;
  End If;

  --1.处理界面选择的且是部分退或部分执行的
  Insert Into 门诊费用记录
    (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id, 收费类别, 收费细目id, 计算单位,
     付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人, 执行状态, 费用状态, 执行时间,
     操作员编号, 操作员姓名, 发生时间, 登记时间, 结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 结论, 缴款组id, 挂号id, 主页id)
    Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id,
           收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 划价人,
           执行人, 执行状态, 费用状态, 执行时间, 操作员编号, 操作员姓名, 发生时间, 登记时间, 结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 结论,
           缴款组id, 挂号id, 主页id
    From (Select NO, Max(实际票号) As 实际票号, 11 As 记录性质, 1 As 记录状态, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式,
                  费别, 病人科室id, 收费类别, 收费细目id, 计算单位, 1 As 付数, Max(发药窗口) As 发药窗口, Sum(Nvl(付数, 1) * Nvl(数次, 0)) As 数次,
                  Max(加班标志) As 加班标志, Max(附加标志) As 附加标志, 收入项目id, 收据费目, 记帐费用, Avg(标准单价) As 标准单价, Sum(应收金额) As 应收金额,
                  Sum(实收金额) As 实收金额, 开单部门id, 开单人, 执行部门id, Max(划价人) As 划价人, Max(执行人) 执行人, Max(执行状态) As 执行状态, 1 As 费用状态,
                  Max(执行时间) 执行时间, v_操作员编号 As 操作员编号, v_操作员姓名 As 操作员姓名, 发生时间, d_登记时间 As 登记时间, 重收结帐id_In As 结帐id,
                  Sum(结帐金额) As 结帐金额, Max(保险项目否) As 保险项目否, 保险大类id, Sum(统筹金额) As 统筹金额,
                  Max(Decode(记录性质, 1, 摘要, 11, 摘要, Null)) As 摘要, 0 As 是否上传, Max(保险编码) As 保险编码, Max(费用类型) As 费用类型,
                  Max(Decode(记录性质, 1, 结论, 11, 结论, Null)) As 结论, n_缴款组id As 缴款组id, Max(挂号id) As 挂号id, Max(主页id) As 主页id
           From 门诊费用记录
           Where Mod(记录性质, 10) = 1 And (NO, 序号) In (Select NO, 序号 From 门诊费用记录 Where 结帐id = 冲销id_In)
           Group By NO, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id, 收费类别, 收费细目id, 计算单位, 收入项目id,
                    收据费目, 记帐费用, 开单部门id, 开单人, 执行部门id, 发生时间, 保险大类id
           Having Sum(Nvl(付数, 1) * Nvl(数次, 0)) <> 0);

  For c_冲销 In (Select NO, 序号, 从属父号, 价格父号, 收入项目id, -1 * Sum(Nvl(付数, 1) * Nvl(数次, 0)) As 数次, Sum(标准单价) As 标准单价,
                      -1 * Sum(应收金额) As 应收金额, -1 * Sum(实收金额) As 实收金额, -1 * Sum(统筹金额) As 统筹金额, -1 * Sum(结帐金额) As 结帐金额
               From 门诊费用记录
               Where 记录性质 = 11 And 结帐id = 重收结帐id_In
               Group By NO, 序号, 从属父号, 价格父号, 收入项目id) Loop
    Update 门诊费用记录
    Set 数次 = Nvl(数次, 0) + Nvl(c_冲销.数次, 0), 实收金额 = Nvl(实收金额, 0) + Nvl(c_冲销.实收金额, 0),
        应收金额 = Nvl(应收金额, 0) + Nvl(c_冲销.应收金额, 0), 结帐金额 = Nvl(结帐金额, 0) + Nvl(c_冲销.结帐金额, 0),
        统筹金额 = Nvl(统筹金额, 0) + Nvl(c_冲销.统筹金额, 0)
    Where NO = c_冲销.No And 序号 = c_冲销.序号 And Nvl(从属父号, -1) = Nvl(c_冲销.从属父号, '-1') And
          Nvl(价格父号, -1) = Nvl(c_冲销.价格父号, '-1') And 收入项目id = c_冲销.收入项目id And 结帐id = 冲销id_In;
  End Loop;

  --2.处理界面未选退费部分,需要全退且产生11的重收记录
  Open c_Fee_Data;
  Loop
    Fetch c_Fee_Data Bulk Collect
      Into t_费用id Limit n_Array_Size;
    Exit When t_费用id.Count = 0;
  
    --退费记录
    Forall I In 1 .. t_费用id.Count
      Insert Into 门诊费用记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id, 收费类别, 收费细目id,
         计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人, 执行状态, 费用状态,
         执行时间, 操作员编号, 操作员姓名, 发生时间, 登记时间, 结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 结论, 缴款组id, 挂号id, 主页id)
        Select 病人费用记录_Id.Nextval, a.No, a.实际票号, 1, 2, a.序号, a.从属父号, a.价格父号, a.病人id, a.医嘱序号, a.门诊标志, a.姓名, a.性别, a.年龄,
               a.标识号, a.付款方式, a.费别, a.病人科室id, a.收费类别, a.收费细目id, a.计算单位, a.付数, a.发药窗口, -1 * a.数次, a.加班标志, a.附加标志,
               a.收入项目id, a.收据费目, a.记帐费用, a.标准单价, -1 * a.应收金额, -1 * a.实收金额, a.开单部门id, a.开单人, a.执行部门id, a.划价人, 执行人,
               Nvl(q.执行状态, -1) As 执行状态, 1, a.执行时间, v_操作员编号, v_操作员姓名, a.发生时间, d_登记时间, 冲销id_In, -1 * a.结帐金额, a.保险项目否,
               a.保险大类id, -1 * a.统筹金额, a.摘要, 0 As 是否上传, a.保险编码, a.费用类型, a.结论, n_缴款组id As 缴款组id, 挂号id, 主页id
        From 门诊费用记录 A,
             (Select j.No, j.序号, Nvl(Max(j.执行状态), 0) - 1 As 执行状态
               From 门诊费用记录 M, 门诊费用记录 J
               Where m.Id = t_费用id(I) And m.No = j.No And m.序号 = j.序号 And Mod(j.记录性质, 10) = 1 And j.记录状态 = 2
               Group By j.No, j.序号) Q
        Where ID = t_费用id(I) And a.No = q.No(+) And a.序号 = q.序号(+);
  
    --将原记录状态由1变为3
    Forall I In 1 .. t_费用id.Count
      Update 门诊费用记录 Set 记录状态 = 3 Where ID = t_费用id(I) And 记录状态 = 1;
  
    --重新收费记录
    If Nvl(重收结帐id_In, 0) <> 0 Then
      Forall I In 1 .. t_费用id.Count
        Insert Into 门诊费用记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id, 收费类别, 收费细目id,
           计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人, 执行状态,
           费用状态, 执行时间, 操作员编号, 操作员姓名, 发生时间, 登记时间, 结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 结论, 缴款组id, 挂号id,
           主页id)
          Select 病人费用记录_Id.Nextval, NO, 实际票号, 11, 1, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id,
                 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id,
                 划价人, 执行人, 执行状态, 1, 执行时间, v_操作员编号, v_操作员姓名, 发生时间, d_登记时间, 重收结帐id_In, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要,
                 0 As 是否上传, 保险编码, 费用类型, 结论, n_缴款组id As 缴款组id, 挂号id, 主页id
          From 门诊费用记录
          Where ID = t_费用id(I);
    End If;
  End Loop;
  Close c_Fee_Data;

  Select Count(1) Into n_Count From 病人预交记录 Where 结帐id = 冲销id_In And 结算方式 Is Null;
  If n_Count = 0 Then
    --退费结算方式
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 结算性质)
      Select 病人预交记录_Id.Nextval, 3, Null, 2, n_病人id, 结算方式, d_登记时间, v_操作员编号, v_操作员姓名, -1 * 冲预交, 冲销id_In, n_缴款组id,
             -1 * 冲销id_In, 2, 3
      From 病人预交记录
      Where 结帐id = 原结帐id_In And 结算方式 In (Select 名称 From 结算方式 Where 性质 In (3, 4)) And
            Instr(',' || 排开医保结算_In || ',', ',' || 结算方式 || ',') = 0 And Mod(记录性质, 10) <> 1;
    --将原误差费全部退了
    --Insert Into 病人预交记录
    --  (ID, 记录性质, NO, 记录状态, 病人id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志)
    --  Select 病人预交记录_Id.Nextval, 3, Null, 2, n_病人id, 结算方式, d_登记时间, v_操作员编号, v_操作员姓名, -1 * 冲预交, 冲销id_In, n_缴款组id,
    --         -1 * 冲销id_In, 2
    --  From 病人预交记录
    --  Where 结帐id = 原结帐id_In And 结算方式 = v_误差费 And Mod(记录性质, 10) <> 1;
  
    Select Sum(冲预交) Into n_冲销金额 From 病人预交记录 Where 结帐id = 冲销id_In;
    Select Sum(结帐金额) Into n_结算金额 From 门诊费用记录 Where 结帐id = 冲销id_In;
  
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 结算性质)
    Values
      (病人预交记录_Id.Nextval, 3, Null, 2, n_病人id, Null, d_登记时间, v_操作员编号, v_操作员姓名, -1 * (Nvl(n_冲销金额, 0) - Nvl(n_结算金额, 0)),
       冲销id_In, n_缴款组id, -1 * 冲销id_In, 1, 3);
  
  End If;
  If Nvl(重收结帐id_In, 0) <> 0 Then
    Select Count(1) Into n_Count From 病人预交记录 Where 结帐id = 重收结帐id_In And 结算方式 Is Null;
    If n_Count = 0 Then
      Select Sum(结帐金额) Into n_结算金额 From 门诊费用记录 Where 结帐id = 重收结帐id_In;
    
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 结算性质)
      Values
        (病人预交记录_Id.Nextval, 3, Null, 1, n_病人id, Null, d_登记时间, v_操作员编号, v_操作员姓名, n_结算金额, 重收结帐id_In, n_缴款组id,
         -1 * 冲销id_In, 1, 3);
    End If;
  
    --场合_In    Integer:=0, --0:门诊;1-住院
    --性质_In    Integer:=1, --1-收费单;2-记帐单
    --操作_In    Integer:=0, --0:删除划价单;1-收费或记帐;2-退费或销帐
    --No_In      门诊费用记录.No%Type,
    --医嘱ids_In Varchar2 := Null
    For c_No In (Select Distinct NO From 门诊费用记录 Where 记录性质 = 11 And 结帐id = 重收结帐id_In) Loop
      Zl_医嘱发送_计费状态_Update(0, 1, 2, c_No.No);
    End Loop;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊收费记录_重收;
/

--127817:董露露,2019-02-25,处理首页提取籍贯时数据提取错误的问题
Create Or Replace Function Zl_Adderss_Structure(v_Addressinfo Varchar2,n_Type Number := Null) Return Varchar2 Is
  --返回结构：省,省编码,是否虚拟,是否不显示,是否只有虚拟级|市,市编码,是否虚拟,是否不显示,是否只有虚拟级 
  --          |区县,区县编码,是否虚拟,是否不显示,是否只有虚拟级|乡镇,乡镇编码,是否虚拟,是否不显示,是否只有虚拟级 
  --          |街道,街道编码,是否虚拟,是否不显示,是否只有虚拟级 
  v_省       Varchar2(100);
  v_Code省   Varchar2(15);
  v_Info省   Varchar2(150);
  v_市       Varchar2(100);
  v_Code市   Varchar2(15);
  v_Info市   Varchar2(150);
  v_区县     Varchar2(100);
  v_Code区县 Varchar2(15);
  v_Info区县 Varchar2(150);
  v_乡镇     Varchar2(100);
  v_Code乡镇 Varchar2(15);
  v_Info乡镇 Varchar2(150);
  v_街道     Varchar2(500);
  v_Code街道 Varchar2(15);
  v_Info街道 Varchar2(550);
  v_Tmp      Varchar2(100);
  v_Adrstmp  Varchar2(500);
  n_Pos      Number(5);
  n_虚拟     Number(1);
  n_不显示   Number(1);
  n_Count    Number(3);
  v_Return   Varchar2(700);
Begin
  --传入结构化的地址，不用进行地址标准化分割解析 
  v_Adrstmp := v_Addressinfo;
  If v_Addressinfo Like '%,%,%,%,%' Then
    n_Pos     := Instr(v_Adrstmp, ',');
    v_省      := Substr(v_Adrstmp, 1, n_Pos - 1);
    v_Adrstmp := Substr(v_Adrstmp, n_Pos + 1);
    n_Pos     := Instr(v_Adrstmp, ',');
    v_市      := Substr(v_Adrstmp, 1, n_Pos - 1);
    v_Adrstmp := Substr(v_Adrstmp, n_Pos + 1);
    n_Pos     := Instr(v_Adrstmp, ',');
    v_区县    := Substr(v_Adrstmp, 1, n_Pos - 1);
    v_Adrstmp := Substr(v_Adrstmp, n_Pos + 1);
    n_Pos     := Instr(v_Adrstmp, ',');
    v_乡镇    := Substr(v_Adrstmp, 1, n_Pos - 1);
    v_街道    := Substr(v_Adrstmp, n_Pos + 1);
    Select Max(编码) Into v_Code省 From 区域 Where 名称 = v_省 And Nvl(级数, 0) = 0;
    --省级地址都没有，就不做处理 
    If v_Code省 Is Not Null Then
      Select Max(编码), Max(是否虚拟), Max(是否不显示)
      Into v_Code市, n_虚拟, n_不显示
      From 区域
      Where 名称 = v_市 And Nvl(级数, 0) = 1 And 上级编码 = v_Code省;
      If v_Code市 Is Not Null Then
        v_Info市 := v_市 || ',' || v_Code市 || ',' || n_虚拟 || ',' || n_不显示;
        Select Max(编码), Max(是否虚拟), Max(是否不显示)
        Into v_Code区县, n_虚拟, n_不显示
        From 区域
        Where 名称 = v_区县 And Nvl(级数, 0) = 2 And 上级编码 = v_Code市;
        --可能是虚拟地址 
      Else
        Select Max(编码), Max(上级编码)
        Into v_Code区县, v_Code市
        From 区域
        Where 名称 = v_区县 And Nvl(级数, 0) = 2 And 上级编码 In (Select 编码 From 区域 Where 上级编码 = v_Code省);
        If v_Code市 Is Not Null Then
          Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示)
          Into v_市, v_Code市, n_虚拟, n_不显示
          From 区域
          Where 编码 = v_Code市;
        End If;
        v_Info市 := v_市 || ',' || v_Code市 || ',' || n_虚拟 || ',' || n_不显示;
        Select Max(编码), Max(是否虚拟), Max(是否不显示)
        Into v_Code区县, n_虚拟, n_不显示
        From 区域
        Where 编码 = v_Code区县;
      End If;
      v_Info区县 := v_区县 || ',' || v_Code区县 || ',' || n_虚拟 || ',' || n_不显示;
      If v_Code区县 Is Not Null Then
        --可能乡镇在详细地址中，关联参数乡镇地址结构化录入 

        Select Max(编码), Max(是否虚拟), Max(是否不显示)
        Into v_Code乡镇, n_虚拟, n_不显示
        From 区域
        Where 名称 = v_乡镇 And Nvl(级数, 0) = 3 And 上级编码 = v_Code区县;
        --可能是虚拟地址 
        If v_Code乡镇 Is Null Then
          Select Max(编码), Max(上级编码)
          Into v_Code街道, v_Code乡镇
          From 区域
          Where 名称 = v_街道 And Nvl(级数, 0) = 4 And 上级编码 In (Select 编码 From 区域 Where 上级编码 = v_Code区县);
          If v_Code乡镇 Is Not Null Then
            Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示)
            Into v_乡镇, v_Code乡镇, n_虚拟, n_不显示
            From 区域
            Where 编码 = v_Code乡镇;
          End If;
        End If;
        v_Info乡镇 := v_乡镇 || ',' || v_Code乡镇 || ',' || n_虚拟 || ',' || n_不显示;
        If v_Code乡镇 Is Not Null Then
          Select Max(编码), Max(是否虚拟), Max(是否不显示)
          Into v_Code街道, n_虚拟, n_不显示
          From 区域
          Where 名称 = v_街道 And Nvl(级数, 0) = 4 And 上级编码 = v_Code乡镇;
          v_Info街道 := v_街道 || ',' || v_Code街道 || ',' || n_虚拟 || ',' || n_不显示;
        End If;
      End If;
    End If;
    --非标准地址，是完整地址，需要分割省，市，县, 
  Else
    v_Adrstmp := v_Addressinfo;
    v_Tmp     := Substr(v_Adrstmp, 1, 2);
    Select Max(名称), Max(编码) Into v_省, v_Code省 From 区域 Where 名称 Like v_Tmp || '%' And Nvl(级数, 0) = 0;
    --有省级地址，说明可以结构化 
    If v_Code省 Is Not Null Then
      --省级地址是标准的 
      If Substr(v_Adrstmp, 1, Length(v_省)) = v_省 Then
        v_Adrstmp := Substr(v_Adrstmp, Length(v_省) + 1);
        --省级地址不标准,可能新疆省略自治区等,此时，市级地址可能是标准化的。 
      Else
        --先判断二级地址是否存在虚拟地址与不显示的地址 
        If v_Tmp = '内蒙' Then
          v_Tmp := '内蒙古';
        Elsif v_Tmp = '黑龙' Then
          v_Tmp := '黑龙江';
        End If;
        v_Adrstmp := Substr(v_Adrstmp, Length(v_Tmp) + 1);
      End If;
      --先截取市级的两个字做关键字，来匹配 
      v_Tmp := Substr(v_Adrstmp, 1, 2);
      If Nvl(n_Type, 0) <> 2 Then
        Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示), Count(1)
        Into v_市, v_Code市, n_虚拟, n_不显示, n_Count
        From 区域
        Where 名称 Like v_Tmp || '%' And Nvl(级数, 0) = 1 And 上级编码 = v_Code省;
	  End If;
      --存在多行匹配，则继续增加长度，暂时从2个字增加到3个字匹配 
      If n_Count > 1 Then
        v_Tmp := Substr(v_Adrstmp, 1, 3);
        Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示)
        Into v_市, v_Code市, n_虚拟, n_不显示
        From 区域
        Where 名称 Like v_Tmp || '%' And Nvl(级数, 0) = 1 And 上级编码 = v_Code省;
      End If;
      --判断是否存在虚拟地址或不显示的地址导致的,如果存在，则根据第三级地址来确定虚拟地址 
      --可能是没有第二级，因此需要第三级判断
      If v_Code市 Is Null Then
        Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示), Count(1), Max(上级编码)
        Into v_区县, v_Code区县, n_虚拟, n_不显示, n_Count, v_Code市
        From 区域
        Where 名称 Like v_Tmp || '%' And Nvl(级数, 0) = 2 And 上级编码 In (Select 编码 From 区域 Where 上级编码 = v_Code省);
        --存在多行匹配，则继续增加长度，暂时从2个字增加到3个字匹配 
        If n_Count > 1 Then
          v_Tmp := Substr(v_Adrstmp, 1, 3);
          Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示), Max(上级编码)
          Into v_区县, v_Code区县, n_虚拟, n_不显示, v_Code市
          From 区域
          Where 名称 Like v_Tmp || '%' And Nvl(级数, 0) = 2 And 上级编码 In (Select 编码 From 区域 Where 上级编码 = v_Code省);
        End If;
        If v_Code市 Is Not Null Then
          v_Info区县 := v_区县 || ',' || v_Code区县 || ',' || n_虚拟 || ',' || n_不显示;
          v_Adrstmp  := Substr(v_Adrstmp, Length(v_区县) + 1);
          Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示)
          Into v_市, v_Code市, n_虚拟, n_不显示
          From 区域
          Where 编码 = v_Code市;
          v_Info市 := v_市 || ',' || v_Code市 || ',' || n_虚拟 || ',' || n_不显示;
        End If;
      Else
        v_Info市  := v_市 || ',' || v_Code市 || ',' || n_虚拟 || ',' || n_不显示;
        v_Adrstmp := Substr(v_Adrstmp, Length(v_市) + 1);
      End If;
      --没有区县，则解析区县 
      If Not v_Code市 Is Null And v_Code区县 Is Null Then
        --先截取县级的两个字做关键字，来匹配 
        v_Tmp := Substr(v_Adrstmp, 1, 2);
        Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示), Count(1)
        Into v_区县, v_Code区县, n_虚拟, n_不显示, n_Count
        From 区域
        Where 名称 Like v_Tmp || '%' And Nvl(级数, 0) = 2 And 上级编码 = v_Code市;
        --存在多行匹配，则继续增加长度，暂时从2个字增加到3个字匹配 
        If n_Count > 1 Then
          v_Tmp := Substr(v_Adrstmp, 1, 3);
          Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示)
          Into v_区县, v_Code区县, n_虚拟, n_不显示
          From 区域
          Where 名称 Like v_Tmp || '%' And Nvl(级数, 0) = 2 And 上级编码 = v_Code市;
        End If;
        If v_Code区县 Is Null Then
          Select Max(是否虚拟), Max(是否不显示) Into n_虚拟, n_不显示 From 区域 Where 上级编码 = v_Code市;
          If Nvl(n_虚拟, 0) = 1 Or Nvl(n_不显示, 0) = 1 Then
            Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示), Count(1), Max(上级编码)
            Into v_乡镇, v_Code乡镇, n_虚拟, n_不显示, n_Count, v_Code区县
            From 区域
            Where 名称 Like v_Tmp || '%' And Nvl(级数, 0) = 3 And 上级编码 In (Select 编码 From 区域 Where 上级编码 = v_Code市);
            --存在多行匹配，则继续增加长度，暂时从2个字增加到3个字匹配 
            If n_Count > 1 Then
              v_Tmp := Substr(v_Adrstmp, 1, 3);
              Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示), Max(上级编码)
              Into v_乡镇, v_Code乡镇, n_虚拟, n_不显示, v_Code区县
              From 区域
              Where 名称 Like v_Tmp || '%' And Nvl(级数, 0) = 3 And 上级编码 In (Select 编码 From 区域 Where 上级编码 = v_Code市);
            End If;
          
            If v_Code乡镇 Is Not Null Then
              v_Info乡镇 := v_乡镇 || ',' || v_Code乡镇 || ',' || n_虚拟 || ',' || n_不显示;
              v_Adrstmp  := Substr(v_Adrstmp, Length(v_乡镇) + 1);
              Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示)
              Into v_区县, v_Code区县, n_虚拟, n_不显示
              From 区域
              Where 编码 = v_Code区县;
              v_Info区县 := v_区县 || ',' || v_Code区县 || ',' || n_虚拟 || ',' || n_不显示;
            End If;
          End If;
        Else
          v_Info区县 := v_区县 || ',' || v_Code区县 || ',' || n_虚拟 || ',' || n_不显示;
          v_Adrstmp  := Substr(v_Adrstmp, Length(v_区县) + 1);
        End If;
      End If;
      If v_Code区县 Is Not Null And v_Code乡镇 Is Null Then
        --先截取乡镇级的两个字做关键字，来匹配 
        v_Tmp := Substr(v_Adrstmp, 1, 2);
        Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示), Count(1)
        Into v_乡镇, v_Code乡镇, n_虚拟, n_不显示, n_Count
        From 区域
        Where 名称 Like v_Tmp || '%' And Nvl(级数, 0) = 3 And 上级编码 = v_Code区县;
        --存在多行匹配，则继续增加长度，暂时从2个字增加到3个字匹配 
        If n_Count > 1 Then
          v_Tmp := Substr(v_Adrstmp, 1, 3);
          Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示), Count(1)
          Into v_乡镇, v_Code乡镇, n_虚拟, n_不显示, n_Count
          From 区域
          Where 名称 Like v_Tmp || '%' And Nvl(级数, 0) = 3 And 上级编码 = v_Code区县;
        End If;
        If v_Code乡镇 Is Null Then
          Select Max(是否虚拟), Max(是否不显示) Into n_虚拟, n_不显示 From 区域 Where 上级编码 = v_Code区县;
          If Nvl(n_虚拟, 0) = 1 Or Nvl(n_不显示, 0) = 1 Then
            Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示), Count(1), Max(上级编码)
            Into v_街道, v_Code街道, n_虚拟, n_不显示, n_Count, v_Code乡镇
            From 区域
            Where 名称 = v_Adrstmp And 上级编码 In (Select 编码 From 区域 Where 上级编码 = v_Code区县);
          End If;
          If v_Code乡镇 Is Not Null Then
            v_Info街道 := v_街道 || ',' || v_Code街道 || ',' || n_虚拟 || ',' || n_不显示;
            Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示)
            Into v_乡镇, v_Code乡镇, n_虚拟, n_不显示
            From 区域
            Where 编码 = v_Code乡镇;
            v_Info乡镇 := v_乡镇 || ',' || v_Code乡镇 || ',' || n_虚拟 || ',' || n_不显示;
          End If;
        Else
          v_Info乡镇 := v_乡镇 || ',' || v_Code乡镇 || ',' || n_虚拟 || ',' || n_不显示;
          v_Adrstmp  := Substr(v_Adrstmp, Length(v_乡镇) + 1);
        End If;
        If v_Code乡镇 Is Not Null And v_Code街道 Is Null Then
          Select Max(名称), Max(编码), Max(是否虚拟), Max(是否不显示)
          Into v_街道, v_Code街道, n_虚拟, n_不显示
          From 区域
          Where 名称 = v_Adrstmp And Nvl(级数, 0) = 4 And 上级编码 = v_Code乡镇;
          If v_Code街道 Is Not Null Then
            v_Info街道 := v_街道 || ',' || v_Code街道 || ',' || n_虚拟 || ',' || n_不显示;
          End If;
        End If;
      End If;
    End If;
    If v_街道 Is Null Then
      v_街道 := v_Adrstmp;
    End If;
  End If;
  v_Info省 := v_省 || ',' || v_Code省 || ',,,';
  If v_Info市 Is Null Then
    v_Info市 := v_市 || ',,,';
  End If;
  --只有省没有市，判断市是否只有虚拟级 
  If Not v_Code省 Is Null And v_市 Is Null Then
    Select Count(1)
    Into n_Count
    From 区域
    Where 上级编码 = v_Code省 And Nvl(是否虚拟, 0) = 0 And Nvl(是否不显示, 0) = 0 And Rownum < 2;
    If n_Count = 0 Then
      Select Count(1) Into n_Count From 区域 Where 上级编码 = v_Code省 And Rownum < 2;
      If n_Count = 0 Then
        v_Info市 := v_Info市 || ',';
      Else
        v_Info市 := v_Info市 || ',1';
      End If;
    Else
      v_Info市 := v_Info市 || ',';
    End If;
  Else
    v_Info市 := v_Info市 || ',';
  End If;
  If v_Info区县 Is Null Then
    v_Info区县 := v_区县 || ',,,';
  End If;
  --只有市没有区县，判断区县只有虚拟级 
  If Not v_Code市 Is Null And v_区县 Is Null Then
    Select Count(1)
    Into n_Count
    From 区域
    Where 上级编码 = v_Code市 And Nvl(是否虚拟, 0) = 0 And Nvl(是否不显示, 0) = 0 And Rownum < 2;
    If n_Count = 0 Then
      Select Count(1) Into n_Count From 区域 Where 上级编码 = v_Code市 And Rownum < 2;
      If n_Count = 0 Then
        v_Info区县 := v_Info区县 || ',';
      Else
        v_Info区县 := v_Info区县 || ',1';
      End If;
    Else
      v_Info区县 := v_Info区县 || ',';
    End If;
  Else
    v_Info区县 := v_Info区县 || ',';
  End If;
  If v_Info乡镇 Is Null Then
    v_Info乡镇 := v_乡镇 || ',,,';
  End If;
  --只有区县没有乡镇，判断乡镇是否只有虚拟的下级 
  If Not v_Code区县 Is Null And v_乡镇 Is Null Then
    Select Count(1)
    Into n_Count
    From 区域
    Where 上级编码 = v_Code区县 And Nvl(是否虚拟, 0) = 0 And Nvl(是否不显示, 0) = 0 And Rownum < 2;
    If n_Count = 0 Then
      Select Count(1) Into n_Count From 区域 Where 上级编码 = v_Code区县 And Rownum < 2;
      If n_Count = 0 Then
        v_Info乡镇 := v_Info乡镇 || ',';
      Else
        v_Info乡镇 := v_Info乡镇 || ',1';
      End If;
    Else
      v_Info乡镇 := v_Info乡镇 || ',';
    End If;
  Else
    v_Info乡镇 := v_Info乡镇 || ',';
  End If;
  If v_Info街道 Is Null Then
    v_Info街道 := v_街道 || ',,,,';
  Else
    v_Info街道 := v_Info街道 || ',';
  End If;
  v_Return := v_Info省 || '|' || v_Info市 || '|' || v_Info区县 || '|' || v_Info乡镇 || '|' || v_Info街道;
  Return(v_Return);
End;
/


------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0051' Where 编号=&n_System;
Commit;
