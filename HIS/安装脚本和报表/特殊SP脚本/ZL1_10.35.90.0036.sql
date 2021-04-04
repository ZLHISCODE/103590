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
--133498:冉俊明,2018-11-06,支持补充结算后的门诊费用转为住院费用
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_system,1131,'办理登记',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0
Union All Select 'Zl_门诊转住院_补结算转出','EXECUTE' From Dual) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_system,1137,'门诊费用转住院',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0
Union All Select 'Zl_门诊转住院_补结算转出','EXECUTE' From Dual) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_system,1604,'补充入院',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0
Union All Select 'Zl_门诊转住院_补结算转出','EXECUTE' From Dual) A;




-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--133498:冉俊明,2018-11-07,支持补充结算后的门诊费用转为住院费用
Create Or Replace Procedure Zl_门诊费用转住院_Insert
(
  No_In         住院费用记录.No%Type,
  住院号_In     住院费用记录.标识号%Type, --医保入院补充登记时才传入
  主页id_In     住院费用记录.主页id%Type, --医保入院补充登记时才传入
  入院时间_In   住院费用记录.发生时间%Type,
  入院科室id_In 病人预交记录.科室id%Type,
  退费时间_In   住院费用记录.登记时间%Type, --多张单据退费时,每张单据的退费时间相同,都是系统当前时间
  操作员编号_In 住院费用记录.操作员编号%Type,
  操作员姓名_In 住院费用记录.操作员姓名%Type,
  入院病区id_In 住院费用记录.病人病区id%Type := Null,
  单据_In       Number := 1,
  结帐id_In     住院费用记录.结帐id%Type := Null,
  原结帐id_In   住院费用记录.结帐id%Type := Null,
  立即销帐_In   Number := 1
) As
  --单据_In:1-门诊收费单;2-记帐单
  v_Billno   住院费用记录.No%Type;
  n_实收合计 住院费用记录.实收金额%Type;
  n_返回值   病人余额.预交余额%Type;

  n_病区id     住院费用记录.病人病区id%Type;
  v_床号       住院费用记录.床号%Type;
  n_医疗小组id 住院费用记录.医疗小组id%Type;

  n_开单部门id     部门表.Id%Type;
  n_操作员编号     门诊费用记录.操作员编号%Type;
  v_操作员姓名     门诊费用记录.操作员姓名%Type;
  v_开单人         人员表.姓名%Type;
  n_病人id         病人信息.病人id%Type;
  n_审核标志       病案主页.审核标志%Type;
  n_住院状态       病案主页.状态%Type;
  n_病人审核方式   Number(2);
  n_未入科禁止记账 Number(2);

  n_结帐id  门诊费用记录.结帐id%Type;
  v_Err_Msg Varchar2(255);
  n_组id    财务缴款分组.Id%Type;
  Err_Item Exception;
  n_Count Number(18);
Begin
  n_组id := Zl_Get组id(操作员姓名_In);
  If Nvl(立即销帐_In, 0) = 1 Then
    If 结帐id_In Is Null Then
      Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
    Else
      n_结帐id := 结帐id_In;
    End If;
  End If;

  If Nvl(主页id_In, 0) <> 0 Then
    n_病人审核方式   := Nvl(zl_GetSysParameter(185), 0);
    n_未入科禁止记账 := Nvl(zl_GetSysParameter(215), 0);
    If n_病人审核方式 = 1 Or n_未入科禁止记账 = 1 Then
      Begin
        Select 病人id
        Into n_病人id
        From 门诊费用记录
        Where NO = No_In And Mod(记录性质, 10) = Mod(单据_In, 10) And Rownum = 1;
      Exception
        When Others Then
          n_病人id := 0;
      End;
      Begin
        Select 审核标志, 状态
        Into n_审核标志, n_住院状态
        From 病案主页
        Where 病人id = Nvl(n_病人id, 0) And 主页id = Nvl(主页id_In, 0);
      Exception
        When Others Then
          n_审核标志 := 0;
          n_住院状态 := 0;
      End;
      If n_未入科禁止记账 = 1 And n_住院状态 = 1 Then
        v_Err_Msg := '病人未入科,禁止对病人相关费用的操作!';
        Raise Err_Item;
      End If;
    
      If n_病人审核方式 = 1 Then
        If Nvl(n_审核标志, 0) = 1 Then
          v_Err_Msg := '该病人目前正在审核费用,不能进行费用相关调整!';
          Raise Err_Item;
        End If;
        If Nvl(n_审核标志, 0) = 2 Then
          v_Err_Msg := '该病人目前已经完成了费用审核,不能进行费用相关调整!';
          Raise Err_Item;
        End If;
      End If;
    End If;
  End If;

  If 原结帐id_In Is Null Then
    If Nvl(立即销帐_In, 0) = 1 Then
      If Mod(单据_In, 10) = 1 Then
        --转收费单
        --No_In;操作员编号_In,操作员姓名_In,退费时间_In,门诊退费_In(0-门诊转住院立即销帐;1-门诊退费模式),入院科室id_In,主页id_In
        Zl_门诊转住院_收费转出(No_In, 操作员编号_In, 操作员姓名_In, 退费时间_In, 0, 入院科室id_In, 主页id_In, Null, n_结帐id);
      Else
        --转记帐单
        --No_In;操作员编号_In,操作员姓名_In,退费时间_In
        Zl_门诊转住院_记帐转出(No_In, 操作员编号_In, 操作员姓名_In, 退费时间_In);
      End If;
    End If;
    --规则
    -- 1.入院病区ID_IN<>NULL 就为:入院病区ID_IN
    -- 2.主页id_In<>0 :
    If Nvl(入院病区id_In, 0) <> 0 Then
      n_病区id := 入院病区id_In;
    Elsif Nvl(主页id_In, 0) <> 0 Then
      Begin
        Select Nvl(b.当前病区id, a.当前病区id), Nvl(b.出院病床, a.当前床号)
        Into n_病区id, v_床号
        From 病人信息 A, 病案主页 B
        Where a.病人id = b.病人id(+) And a.病人id = n_病人id And b.主页id(+) = 主页id_In;
      Exception
        When Others Then
          n_病区id := Null;
      End;
    End If;
  
    If Nvl(n_病区id, 0) = 0 Then
      --以入院科室为准
      n_病区id := Nvl(入院科室id_In, 0);
    End If;
  
    --入院登记之前,主页ID还没有产生,病人预交记录,病人未结费用,未发药品记录,门诊费用记录
    --建有病人ID,主页ID的外键,只有入院登记后再调用Zl_门诊费用转住院_Update填写
    Select 病人id, 开单部门id, 开单人
    Into n_病人id, n_开单部门id, v_开单人
    From 门诊费用记录
    Where NO = No_In And Mod(记录性质, 10) = Mod(单据_In, 10) And 记录状态 In (1, 3) And Rownum = 1;
  
    --5.产生记帐单
    --需要检查是否已经转出
    Select Count(*)
    Into n_Count
    From 门诊费用记录 A
    Where NO = No_In And Mod(记录性质, 10) = Mod(单据_In, 10) And 记录状态 In (1, 3) And Exists
     (Select 1 From 费用审核记录 Where a.Id = 费用id And 转出id Is Not Null) And Rownum <= 1;
    If n_Count >= 1 Then
      v_Err_Msg := '可能因并发原因,该费用已经被他人转出,不能继续操作!';
      Raise Err_Item;
    End If;
    If Mod(单据_In, 10) = 1 Then
      --收费按照结算序号查出包含NO号进行处理
      n_医疗小组id := Zl_医疗小组_Get(n_开单部门id, v_开单人, n_病人id, 主页id_In, 入院时间_In);
      v_Billno     := Nextno(14);
    
      Insert Into 住院费用记录
        (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 门诊标志, 病人id, 主页id, 标识号, 姓名, 性别, 年龄, 床号, 病人病区id, 病人科室id, 费别, 收费类别,
         收费细目id, 计算单位, 保险项目否, 保险大类id, 保险编码, 费用类型, 发药窗口, 付数, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额,
         记帐费用, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 执行时间, 划价人, 操作员编号, 操作员姓名, 记帐单id, 摘要, 是否急诊, 医嘱序号, 缴款组id, 医疗小组id)
        Select 病人费用记录_Id.Nextval, 2, v_Billno, 1, 序号, 从属父号, 价格父号, 0, 2, 病人id, 主页id_In, 住院号_In, 姓名, 性别, 年龄, v_床号,
               Decode(n_病区id, Null, Null, 0, Null, n_病区id), 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id, 保险编码, 费用类型,
               发药窗口, 付数, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额, 1, 开单部门id, 开单人, 入院时间_In, 退费时间_In,
               执行部门id, 执行人, 执行状态, 执行时间, 划价人, 操作员编号, 操作员姓名, 记帐单id, '门诊费用转入', 是否急诊, 医嘱序号, n_组id, n_医疗小组id
        From 门诊费用记录
        Where NO = No_In And Mod(记录性质, 10) = Mod(单据_In, 10) And 记录状态 = 1 And Nvl(附加标志, 0) Not In (8, 9);
    
      If Nvl(立即销帐_In, 0) = 1 Then
        Update 门诊费用记录 Set 记录状态 = 3 Where NO = No_In And Mod(记录性质, 10) In (1, 2) And 记录状态 = 1;
      End If;
    
      For r_Clinic In (Select Min(记录性质) As 记录性质, 序号, 从属父号, 价格父号, 病人id, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否,
                              保险大类id, 保险编码, 费用类型, 发药窗口, 付数, Sum(数次) As 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价,
                              Sum(应收金额) As 应收金额, Sum(实收金额) As 实收金额, Sum(统筹金额) As 统筹金额, 开单部门id, 开单人, 执行部门id,
                              Max(执行人) As 执行人, 划价人, Max(记帐单id) As 记帐单id, Max(是否急诊) As 是否急诊, 发生时间, Min(实际票号) As 实际票号,
                              Max(执行状态) As 执行状态, Max(执行时间) As 执行时间
                       From 门诊费用记录
                       Where NO = No_In And Mod(记录性质, 10) = Mod(单据_In, 10) And 记录状态 In (2, 3) And
                             Nvl(附加标志, 0) Not In (8, 9)
                       Group By 序号, 从属父号, 价格父号, 病人id, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id, 保险编码,
                                费用类型, 发药窗口, 付数, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, 开单部门id, 开单人, 执行部门id, 划价人, 发生时间
                       Having Sum(数次) <> 0) Loop
        Select 操作员编号, 操作员姓名
        Into n_操作员编号, v_操作员姓名
        From 门诊费用记录
        Where NO = No_In And Mod(记录性质, 10) = Mod(单据_In, 10) And 记录状态 = 3 And Rownum < 2;
        Insert Into 住院费用记录
          (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 门诊标志, 病人id, 主页id, 标识号, 姓名, 性别, 年龄, 床号, 病人病区id, 病人科室id, 费别, 收费类别,
           收费细目id, 计算单位, 保险项目否, 保险大类id, 保险编码, 费用类型, 发药窗口, 付数, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额, 记帐费用,
           开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人, 划价人, 操作员编号, 操作员姓名, 记帐单id, 摘要, 是否急诊, 缴款组id, 医疗小组id, 执行状态, 执行时间)
        Values
          (病人费用记录_Id.Nextval, 2, v_Billno, 1, r_Clinic.序号, r_Clinic.从属父号, r_Clinic.价格父号, 0, 2, r_Clinic.病人id, 主页id_In,
           住院号_In, r_Clinic.姓名, r_Clinic.性别, r_Clinic.年龄, v_床号, Decode(n_病区id, Null, Null, 0, Null, n_病区id),
           r_Clinic.病人科室id, r_Clinic.费别, r_Clinic.收费类别, r_Clinic.收费细目id, r_Clinic.计算单位, r_Clinic.保险项目否, r_Clinic.保险大类id,
           r_Clinic.保险编码, r_Clinic.费用类型, r_Clinic.发药窗口, r_Clinic.付数, r_Clinic.数次, r_Clinic.加班标志, r_Clinic.附加标志,
           r_Clinic.收入项目id, r_Clinic.收据费目, r_Clinic.标准单价, r_Clinic.应收金额, r_Clinic.实收金额, r_Clinic.统筹金额, 1,
           r_Clinic.开单部门id, r_Clinic.开单人, 入院时间_In, 退费时间_In, r_Clinic.执行部门id, r_Clinic.执行人, r_Clinic.划价人, n_操作员编号,
           v_操作员姓名, r_Clinic.记帐单id, '门诊费用转入', r_Clinic.是否急诊, n_组id, n_医疗小组id, r_Clinic.执行状态, r_Clinic.执行时间);
      
        If Nvl(立即销帐_In, 0) = 1 Then
          Insert Into 门诊费用记录
            (ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 门诊标志, 病人id, 标识号, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位,
             保险项目否, 保险大类id, 保险编码, 费用类型, 发药窗口, 付数, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额, 记帐费用, 开单部门id,
             开单人, 发生时间, 登记时间, 执行部门id, 划价人, 操作员编号, 操作员姓名, 记帐单id, 摘要, 是否急诊, 缴款组id, 结帐id, 结帐金额, 执行状态, 费用状态, 执行时间)
          Values
            (病人费用记录_Id.Nextval, r_Clinic.记录性质, No_In, r_Clinic.实际票号, 2, r_Clinic.序号, r_Clinic.从属父号, r_Clinic.价格父号, 1,
             r_Clinic.病人id, 住院号_In, r_Clinic.姓名, r_Clinic.性别, r_Clinic.年龄, r_Clinic.病人科室id, r_Clinic.费别, r_Clinic.收费类别,
             r_Clinic.收费细目id, r_Clinic.计算单位, r_Clinic.保险项目否, r_Clinic.保险大类id, r_Clinic.保险编码, r_Clinic.费用类型,
             r_Clinic.发药窗口, r_Clinic.付数, -1 * r_Clinic.数次, r_Clinic.加班标志, r_Clinic.附加标志, r_Clinic.收入项目id, r_Clinic.收据费目,
             r_Clinic.标准单价, -1 * r_Clinic.应收金额, -1 * r_Clinic.实收金额, -1 * r_Clinic.统筹金额, 0, r_Clinic.开单部门id, r_Clinic.开单人,
             r_Clinic.发生时间, 退费时间_In, r_Clinic.执行部门id, r_Clinic.划价人, 操作员编号_In, 操作员姓名_In, r_Clinic.记帐单id, '',
             r_Clinic.是否急诊, n_组id, n_结帐id, -1 * r_Clinic.实收金额, -1, 1, r_Clinic.执行时间);
        End If;
      End Loop;
    
      --8-工本费，9-误差费
      --病人余额
      Select Nvl(Sum(实收金额), 0) Into n_实收合计 From 住院费用记录 Where NO = v_Billno And 记录性质 = 2;
      Update 病人余额
      Set 费用余额 = Nvl(费用余额, 0) + n_实收合计
      Where 病人id = n_病人id And 性质 = 1 And 类型 = 2
      Returning 费用余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额 (病人id, 性质, 类型, 费用余额, 预交余额) Values (n_病人id, 1, 2, n_实收合计, 0);
        n_返回值 := n_实收合计;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete 病人余额 Where 性质 = 1 And 病人id = n_病人id And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
      End If;
    
      --病人未结费用
      For r_Fee In (Select 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, Nvl(Sum(实收金额), 0) 实收合计
                    From 住院费用记录
                    Where NO = v_Billno And 记录性质 = 2
                    Group By 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id) Loop
        Update 病人未结费用
        Set 金额 = Nvl(金额, 0) + r_Fee.实收合计
        Where 病人id = n_病人id And Nvl(主页id, 0) = Nvl(主页id_In, 0) And Nvl(病人病区id, 0) = Nvl(r_Fee.病人病区id, 0) And
              Nvl(病人科室id, 0) = Nvl(r_Fee.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(r_Fee.开单部门id, 0) And
              Nvl(执行部门id, 0) = Nvl(r_Fee.执行部门id, 0) And 收入项目id + 0 = r_Fee.收入项目id And 来源途径 + 0 = 2;
        If Sql%RowCount = 0 Then
          Insert Into 病人未结费用
            (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
          Values
            (n_病人id, 主页id_In, r_Fee.病人病区id, r_Fee.病人科室id, r_Fee.开单部门id, r_Fee.执行部门id, r_Fee.收入项目id, 2, r_Fee.实收合计);
        End If;
      End Loop;
    
      --6.药品相关数据处理
      For r_Fee In (Select a.Id, a.序号, b.材料id, b.跟踪在用
                    From 住院费用记录 A, 材料特性 B
                    Where a.收费细目id = b.材料id(+) And a.No = v_Billno And a.记录性质 = 2 And a.记录状态 In (1, 3)) Loop
        Update 药品收发记录
        Set 单据 = Decode(单据, 8, 9, 24, 25, 单据), 费用id = r_Fee.Id, NO = v_Billno
        Where NO = No_In And 单据 In (8, 9, 24, 25) And
              费用id In
              (Select ID
               From 门诊费用记录
               Where NO = No_In And Mod(记录性质, 10) = Mod(单据_In, 10) And 记录状态 In (1, 3) And 序号 = r_Fee.序号);
        If Nvl(r_Fee.材料id, 0) <> 0 And Nvl(r_Fee.跟踪在用, 0) = 1 Then
          --更新备货材料
          Update 药品收发记录
          Set 费用id = r_Fee.Id
          Where 单据 = 21 And
                费用id In
                (Select ID
                 From 门诊费用记录
                 Where NO = No_In And Mod(记录性质, 10) = Mod(单据_In, 10) And 记录状态 In (1, 3) And 序号 = r_Fee.序号);
        End If;
        --更新费用审核记录
        Update 费用审核记录
        Set 转出id = r_Fee.Id, 记录状态 = Decode(立即销帐_In, 1, 2, 1), 主页id = 主页id_In, 转出人 = 操作员姓名_In, 转出时间 = 退费时间_In
        Where 费用id In
              (Select ID
               From 门诊费用记录
               Where NO = No_In And Mod(记录性质, 10) = Mod(单据_In, 10) And 记录状态 In (1, 3) And 序号 = r_Fee.序号) And 性质 = 1;
        If Sql%NotFound Then
          --未找到数据时，要强制进行对应.
          Insert Into 费用审核记录
            (性质, 费用id, 病人id, 主页id, 审核人, 审核日期, 记录状态, 转出id, 转出人, 转出时间)
            Select 1, ID, n_病人id, 主页id_In, 操作员姓名_In, 退费时间_In, Decode(立即销帐_In, 1, 2, 1), r_Fee.Id, 操作员姓名_In, 退费时间_In
            From 门诊费用记录
            Where NO = No_In And Mod(记录性质, 10) = Mod(单据_In, 10) And 记录状态 In (1, 3) And 序号 = r_Fee.序号;
        End If;
      End Loop;
      Update 未发药品记录
      Set 单据 = Decode(单据, 8, 9, 24, 25, 单据), 主页id = 主页id_In, NO = v_Billno
      Where NO = No_In And 单据 In (8, 24) And 病人id = n_病人id;
    Else
      --记账按照单据NO进行处理
      v_Billno     := Nextno(14);
      n_医疗小组id := Zl_医疗小组_Get(n_开单部门id, v_开单人, n_病人id, 主页id_In, 入院时间_In);
    
      Insert Into 住院费用记录
        (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 门诊标志, 病人id, 主页id, 标识号, 姓名, 性别, 年龄, 床号, 病人病区id, 病人科室id, 费别, 收费类别,
         收费细目id, 计算单位, 保险项目否, 保险大类id, 保险编码, 费用类型, 发药窗口, 付数, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额,
         记帐费用, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 执行时间, 划价人, 操作员编号, 操作员姓名, 记帐单id, 摘要, 是否急诊, 医嘱序号, 缴款组id, 医疗小组id)
        Select 病人费用记录_Id.Nextval, 2, v_Billno, 1, 序号, 从属父号, 价格父号, 0, 2, 病人id, 主页id_In, 住院号_In, 姓名, 性别, 年龄, v_床号,
               Decode(n_病区id, Null, Null, 0, Null, n_病区id), 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id, 保险编码, 费用类型,
               发药窗口, 付数, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额, 1, 开单部门id, 开单人, 入院时间_In, 退费时间_In,
               执行部门id, 执行人, 执行状态, 执行时间, 划价人, 操作员编号, 操作员姓名, 记帐单id, '门诊费用转入', 是否急诊, 医嘱序号, n_组id, n_医疗小组id
        From 门诊费用记录
        Where NO = No_In And 记录性质 = 单据_In And 记录状态 = 1 And Nvl(附加标志, 0) Not In (8, 9);
    
      If Nvl(立即销帐_In, 0) = 1 Then
        Update 门诊费用记录 Set 记录状态 = 3 Where NO = No_In And 记录性质 In (1, 2) And 记录状态 = 1;
      End If;
    
      For r_Clinic In (Select 记录性质, 序号, 从属父号, 价格父号, 病人id, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id, 保险编码,
                              费用类型, 发药窗口, 付数, Sum(数次) As 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, Sum(应收金额) As 应收金额,
                              Sum(实收金额) As 实收金额, Sum(统筹金额) As 统筹金额, 开单部门id, 开单人, 执行部门id, Max(执行人) As 执行人, 划价人,
                              Max(记帐单id) As 记帐单id, 发生时间, 实际票号, Max(执行状态) As 执行状态, Max(执行时间) As 执行时间

                       
                       From 门诊费用记录
                       Where NO = No_In And 记录性质 = 单据_In And 记录状态 In (2, 3) And Nvl(附加标志, 0) Not In (8, 9)
                       Group By 记录性质, 序号, 从属父号, 价格父号, 病人id, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id,
                                保险编码, 费用类型, 发药窗口, 付数, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, 开单部门id, 开单人, 执行部门id, 划价人, 发生时间,
                                实际票号
                       Having Sum(数次) <> 0) Loop
        Select 操作员编号, 操作员姓名
        Into n_操作员编号, v_操作员姓名
        From 门诊费用记录
        Where NO = No_In And 记录性质 = 单据_In And 记录状态 = 3 And Rownum < 2;
        Insert Into 住院费用记录
          (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 门诊标志, 病人id, 主页id, 标识号, 姓名, 性别, 年龄, 床号, 病人病区id, 病人科室id, 费别, 收费类别,
           收费细目id, 计算单位, 保险项目否, 保险大类id, 保险编码, 费用类型, 发药窗口, 付数, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额, 记帐费用,
           开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人, 划价人, 操作员编号, 操作员姓名, 记帐单id, 摘要, 缴款组id, 医疗小组id, 执行状态, 执行时间)
        Values
          (病人费用记录_Id.Nextval, 2, v_Billno, 1, r_Clinic.序号, r_Clinic.从属父号, r_Clinic.价格父号, 0, 2, r_Clinic.病人id, 主页id_In,
           住院号_In, r_Clinic.姓名, r_Clinic.性别, r_Clinic.年龄, v_床号, Decode(n_病区id, Null, Null, 0, Null, n_病区id),
           r_Clinic.病人科室id, r_Clinic.费别, r_Clinic.收费类别, r_Clinic.收费细目id, r_Clinic.计算单位, r_Clinic.保险项目否, r_Clinic.保险大类id,
           r_Clinic.保险编码, r_Clinic.费用类型, r_Clinic.发药窗口, r_Clinic.付数, r_Clinic.数次, r_Clinic.加班标志, r_Clinic.附加标志,
           r_Clinic.收入项目id, r_Clinic.收据费目, r_Clinic.标准单价, r_Clinic.应收金额, r_Clinic.实收金额, r_Clinic.统筹金额, 1,
           r_Clinic.开单部门id, r_Clinic.开单人, 入院时间_In, 退费时间_In, r_Clinic.执行部门id, r_Clinic.执行人, r_Clinic.划价人, n_操作员编号,
           v_操作员姓名, r_Clinic.记帐单id, '门诊费用转入', n_组id, n_医疗小组id, r_Clinic.执行状态, r_Clinic.执行时间);
        If Nvl(立即销帐_In, 0) = 1 Then
          Insert Into 门诊费用记录
            (ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 门诊标志, 病人id, 标识号, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位,
             保险项目否, 保险大类id, 保险编码, 费用类型, 发药窗口, 付数, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额, 记帐费用, 开单部门id,
             开单人, 发生时间, 登记时间, 执行部门id, 划价人, 操作员编号, 操作员姓名, 记帐单id, 摘要, 缴款组id, 结帐id, 结帐金额, 费用状态, 执行时间)
          Values
            (病人费用记录_Id.Nextval, r_Clinic.记录性质, No_In, r_Clinic.实际票号, 2, r_Clinic.序号, r_Clinic.从属父号, r_Clinic.价格父号, 1,
             r_Clinic.病人id, 住院号_In, r_Clinic.姓名, r_Clinic.性别, r_Clinic.年龄, r_Clinic.病人科室id, r_Clinic.费别, r_Clinic.收费类别,
             r_Clinic.收费细目id, r_Clinic.计算单位, r_Clinic.保险项目否, r_Clinic.保险大类id, r_Clinic.保险编码, r_Clinic.费用类型,
             r_Clinic.发药窗口, r_Clinic.付数, -1 * r_Clinic.数次, r_Clinic.加班标志, r_Clinic.附加标志, r_Clinic.收入项目id, r_Clinic.收据费目,
             r_Clinic.标准单价, -1 * r_Clinic.应收金额, -1 * r_Clinic.实收金额, -1 * r_Clinic.统筹金额, 1, r_Clinic.开单部门id, r_Clinic.开单人,
             r_Clinic.发生时间, 退费时间_In, r_Clinic.执行部门id, r_Clinic.划价人, 操作员编号_In, 操作员姓名_In, r_Clinic.记帐单id, '', n_组id,
             n_结帐id, -1 * r_Clinic.实收金额, 1, r_Clinic.执行时间);
        End If;
      End Loop;
    
      --8-工本费，9-误差费
      --病人余额
      Select Nvl(Sum(实收金额), 0) Into n_实收合计 From 住院费用记录 Where NO = v_Billno And 记录性质 = 2;
      Update 病人余额
      Set 费用余额 = Nvl(费用余额, 0) + n_实收合计
      Where 病人id = n_病人id And 性质 = 1 And 类型 = 2
      Returning 费用余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额 (病人id, 性质, 类型, 费用余额, 预交余额) Values (n_病人id, 1, 2, n_实收合计, 0);
        n_返回值 := n_实收合计;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete 病人余额 Where 性质 = 1 And 病人id = n_病人id And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
      End If;
    
      --病人未结费用
      For r_Fee In (Select 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, Nvl(Sum(实收金额), 0) 实收合计
                    From 住院费用记录
                    Where NO = v_Billno And 记录性质 = 2
                    Group By 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id) Loop
        Update 病人未结费用
        Set 金额 = Nvl(金额, 0) + r_Fee.实收合计
        Where 病人id = n_病人id And Nvl(主页id, 0) = Nvl(主页id_In, 0) And Nvl(病人病区id, 0) = Nvl(r_Fee.病人病区id, 0) And
              Nvl(病人科室id, 0) = Nvl(r_Fee.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(r_Fee.开单部门id, 0) And
              Nvl(执行部门id, 0) = Nvl(r_Fee.执行部门id, 0) And 收入项目id + 0 = r_Fee.收入项目id And 来源途径 + 0 = 2;
        If Sql%RowCount = 0 Then
          Insert Into 病人未结费用
            (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
          Values
            (n_病人id, 主页id_In, r_Fee.病人病区id, r_Fee.病人科室id, r_Fee.开单部门id, r_Fee.执行部门id, r_Fee.收入项目id, 2, r_Fee.实收合计);
        End If;
      End Loop;
    
      --6.药品相关数据处理
      For r_Fee In (Select a.Id, a.序号, b.材料id, b.跟踪在用
                    From 住院费用记录 A, 材料特性 B
                    Where a.收费细目id = b.材料id(+) And a.No = v_Billno And a.记录性质 = 2 And a.记录状态 In (1, 3)) Loop
        Update 药品收发记录
        Set 单据 = Decode(单据, 8, 9, 24, 25, 单据), 费用id = r_Fee.Id, NO = v_Billno
        Where NO = No_In And 单据 In (8, 9, 24, 25) And
              费用id In (Select ID
                       From 门诊费用记录
                       Where NO = No_In And 记录性质 = 单据_In And 记录状态 In (1, 3) And 序号 = r_Fee.序号);
        If Nvl(r_Fee.材料id, 0) <> 0 And Nvl(r_Fee.跟踪在用, 0) = 1 Then
          --更新备货材料
          Update 药品收发记录
          Set 费用id = r_Fee.Id
          Where 单据 = 21 And
                费用id In (Select ID
                         From 门诊费用记录
                         Where NO = No_In And 记录性质 = 单据_In And 记录状态 In (1, 3) And 序号 = r_Fee.序号);
        End If;
        --更新费用审核记录
        Update 费用审核记录
        Set 转出id = r_Fee.Id, 记录状态 = Decode(立即销帐_In, 1, 2, 1), 主页id = 主页id_In, 转出人 = 操作员姓名_In, 转出时间 = 退费时间_In
        Where 费用id In (Select ID
                       From 门诊费用记录
                       Where NO = No_In And 记录性质 = 单据_In And 记录状态 In (1, 3) And 序号 = r_Fee.序号) And 性质 = 1;
        If Sql%NotFound Then
          --未找到数据时，要强制进行对应.
          Insert Into 费用审核记录
            (性质, 费用id, 病人id, 主页id, 审核人, 审核日期, 记录状态, 转出id, 转出人, 转出时间)
            Select 1, ID, n_病人id, 主页id_In, 操作员姓名_In, 退费时间_In, Decode(立即销帐_In, 1, 2, 1), r_Fee.Id, 操作员姓名_In, 退费时间_In
            From 门诊费用记录
            Where NO = No_In And 记录性质 = 单据_In And 记录状态 In (1, 3) And 序号 = r_Fee.序号;
        End If;
      End Loop;
      Update 未发药品记录
      Set 主页id = 主页id_In, NO = v_Billno
      Where NO = No_In And 单据 In (9, 25) And 病人id = n_病人id;
    End If;
  Else
    If Nvl(立即销帐_In, 0) = 1 Then
      If Mod(单据_In, 10) = 1 Then
        Zl_门诊转住院_收费转出(No_In, 操作员编号_In, 操作员姓名_In, 退费时间_In, 0, 入院科室id_In, 主页id_In, Null, n_结帐id, 原结帐id_In);
      Else
        --转记帐单
        --No_In;操作员编号_In,操作员姓名_In,退费时间_In
        Zl_门诊转住院_记帐转出(No_In, 操作员编号_In, 操作员姓名_In, 退费时间_In);
      End If;
    End If;
    --规则
    -- 1.入院病区ID_IN<>NULL 就为:入院病区ID_IN
    -- 2.主页id_In<>0 :
    If Nvl(入院病区id_In, 0) <> 0 Then
      n_病区id := 入院病区id_In;
    Elsif Nvl(主页id_In, 0) <> 0 Then
      Begin
        Select Nvl(b.当前病区id, a.当前病区id), Nvl(b.出院病床, a.当前床号)
        Into n_病区id, v_床号
        From 病人信息 A, 病案主页 B
        Where a.病人id = b.病人id(+) And a.病人id = n_病人id And b.主页id(+) = 主页id_In;
      Exception
        When Others Then
          n_病区id := Null;
      End;
    End If;
  
    If Nvl(n_病区id, 0) = 0 Then
      --以入院科室为准
      n_病区id := Nvl(入院科室id_In, 0);
    End If;
  
    --入院登记之前,主页ID还没有产生,病人预交记录,病人未结费用,未发药品记录,门诊费用记录
    --建有病人ID,主页ID的外键,只有入院登记后再调用Zl_门诊费用转住院_Update填写
    Select 病人id, 开单部门id, 开单人
    Into n_病人id, n_开单部门id, v_开单人
    From 门诊费用记录
    Where NO = No_In And Mod(记录性质, 10) = Mod(单据_In, 10) And 记录状态 In (1, 3) And Rownum = 1;
  
    --5.产生记帐单
    --需要检查是否已经转出
    Select Count(*)
    Into n_Count
    From 门诊费用记录 A
    Where NO = No_In And Mod(记录性质, 10) = Mod(单据_In, 10) And 记录状态 In (1, 3) And Exists
     (Select 1 From 费用审核记录 Where a.Id = 费用id And 转出id Is Not Null) And Rownum <= 1;
    If n_Count >= 1 Then
      v_Err_Msg := '可能因并发原因,该费用已经被他人转出,不能继续操作!';
      Raise Err_Item;
    End If;
    If Mod(单据_In, 10) = 1 Then
      --收费按照结算序号查出包含NO号进行处理
      n_医疗小组id := Zl_医疗小组_Get(n_开单部门id, v_开单人, n_病人id, 主页id_In, 入院时间_In);
      For r_Nos In (Select Distinct c.No
                    From 门诊费用记录 A, 门诊费用记录 C
                    Where Mod(a.记录性质, 10) = 1 And a.记录状态 In (1, 3) And a.结帐id = 原结帐id_In And c.No = a.No And
                          Mod(c.记录性质, 10) = 1 And c.记录状态 In (1, 3)) Loop
        v_Billno := Nextno(14);
      
        Insert Into 住院费用记录
          (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 门诊标志, 病人id, 主页id, 标识号, 姓名, 性别, 年龄, 床号, 病人病区id, 病人科室id, 费别, 收费类别,
           收费细目id, 计算单位, 保险项目否, 保险大类id, 保险编码, 费用类型, 发药窗口, 付数, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额,
           记帐费用, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 执行时间, 划价人, 操作员编号, 操作员姓名, 记帐单id, 摘要, 是否急诊, 医嘱序号, 缴款组id,
           医疗小组id)
          Select 病人费用记录_Id.Nextval, 2, v_Billno, 1, 序号, 从属父号, 价格父号, 0, 2, 病人id, 主页id_In, 住院号_In, 姓名, 性别, 年龄, v_床号,
                 Decode(n_病区id, Null, Null, 0, Null, n_病区id), 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id, 保险编码, 费用类型,
                 发药窗口, 付数, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额, 1, 开单部门id, 开单人, 入院时间_In, 退费时间_In,
                 执行部门id, 执行人, 执行状态, 执行时间, 划价人, 操作员编号, 操作员姓名, 记帐单id, '门诊费用转入', 是否急诊, 医嘱序号, n_组id, n_医疗小组id
          From 门诊费用记录
          Where NO = r_Nos.No And Mod(记录性质, 10) = Mod(单据_In, 10) And 记录状态 = 1 And Nvl(附加标志, 0) Not In (8, 9);
      
        If Nvl(立即销帐_In, 0) = 1 Then
          Update 门诊费用记录 Set 记录状态 = 3 Where NO = r_Nos.No And Mod(记录性质, 10) In (1, 2) And 记录状态 = 1;
        End If;
      
        For r_Clinic In (Select Min(记录性质) As 记录性质, 序号, 从属父号, 价格父号, 病人id, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位,
                                保险项目否, 保险大类id, 保险编码, 费用类型, 发药窗口, 付数, Sum(数次) As 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价,
                                Sum(应收金额) As 应收金额, Sum(实收金额) As 实收金额, Sum(统筹金额) As 统筹金额, 开单部门id, 开单人, 执行部门id,
                                Max(执行人) As 执行人, 划价人, Max(记帐单id) As 记帐单id, Max(是否急诊) As 是否急诊, 发生时间, Min(实际票号) As 实际票号,
                                Max(执行状态) As 执行状态, Max(执行时间) As 执行时间
                         From 门诊费用记录
                         Where NO = r_Nos.No And Mod(记录性质, 10) = Mod(单据_In, 10) And 记录状态 In (2, 3) And
                               Nvl(附加标志, 0) Not In (8, 9)
                         Group By 序号, 从属父号, 价格父号, 病人id, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id, 保险编码,
                                  费用类型, 发药窗口, 付数, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, 开单部门id, 开单人, 执行部门id, 划价人, 发生时间
                         Having Sum(数次) <> 0) Loop
          Select 操作员编号, 操作员姓名
          Into n_操作员编号, v_操作员姓名
          From 门诊费用记录
          Where NO = r_Nos.No And Mod(记录性质, 10) = Mod(单据_In, 10) And 记录状态 = 3 And Rownum < 2;
          Insert Into 住院费用记录
            (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 门诊标志, 病人id, 主页id, 标识号, 姓名, 性别, 年龄, 床号, 病人病区id, 病人科室id, 费别, 收费类别,
             收费细目id, 计算单位, 保险项目否, 保险大类id, 保险编码, 费用类型, 发药窗口, 付数, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额,
             记帐费用, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人, 划价人, 操作员编号, 操作员姓名, 记帐单id, 摘要, 是否急诊, 缴款组id, 医疗小组id, 执行状态, 执行时间)
          Values
            (病人费用记录_Id.Nextval, 2, v_Billno, 1, r_Clinic.序号, r_Clinic.从属父号, r_Clinic.价格父号, 0, 2, r_Clinic.病人id, 主页id_In,
             住院号_In, r_Clinic.姓名, r_Clinic.性别, r_Clinic.年龄, v_床号, Decode(n_病区id, Null, Null, 0, Null, n_病区id),
             r_Clinic.病人科室id, r_Clinic.费别, r_Clinic.收费类别, r_Clinic.收费细目id, r_Clinic.计算单位, r_Clinic.保险项目否,
             r_Clinic.保险大类id, r_Clinic.保险编码, r_Clinic.费用类型, r_Clinic.发药窗口, r_Clinic.付数, r_Clinic.数次, r_Clinic.加班标志,
             r_Clinic.附加标志, r_Clinic.收入项目id, r_Clinic.收据费目, r_Clinic.标准单价, r_Clinic.应收金额, r_Clinic.实收金额, r_Clinic.统筹金额,
             1, r_Clinic.开单部门id, r_Clinic.开单人, 入院时间_In, 退费时间_In, r_Clinic.执行部门id, r_Clinic.执行人, r_Clinic.划价人, n_操作员编号,
             v_操作员姓名, r_Clinic.记帐单id, '门诊费用转入', r_Clinic.是否急诊, n_组id, n_医疗小组id, r_Clinic.执行状态, r_Clinic.执行时间);
        
          If Nvl(立即销帐_In, 0) = 1 Then
            Insert Into 门诊费用记录
              (ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 门诊标志, 病人id, 标识号, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位,
               保险项目否, 保险大类id, 保险编码, 费用类型, 发药窗口, 付数, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额, 记帐费用, 开单部门id,
               开单人, 发生时间, 登记时间, 执行部门id, 划价人, 操作员编号, 操作员姓名, 记帐单id, 摘要, 是否急诊, 缴款组id, 结帐id, 结帐金额, 执行状态, 费用状态, 执行时间)
            Values
              (病人费用记录_Id.Nextval, r_Clinic.记录性质, r_Nos.No, r_Clinic.实际票号, 2, r_Clinic.序号, r_Clinic.从属父号, r_Clinic.价格父号,
               1, r_Clinic.病人id, 住院号_In, r_Clinic.姓名, r_Clinic.性别, r_Clinic.年龄, r_Clinic.病人科室id, r_Clinic.费别,
               r_Clinic.收费类别, r_Clinic.收费细目id, r_Clinic.计算单位, r_Clinic.保险项目否, r_Clinic.保险大类id, r_Clinic.保险编码,
               r_Clinic.费用类型, r_Clinic.发药窗口, r_Clinic.付数, -1 * r_Clinic.数次, r_Clinic.加班标志, r_Clinic.附加标志,
               r_Clinic.收入项目id, r_Clinic.收据费目, r_Clinic.标准单价, -1 * r_Clinic.应收金额, -1 * r_Clinic.实收金额, -1 * r_Clinic.统筹金额,
               0, r_Clinic.开单部门id, r_Clinic.开单人, r_Clinic.发生时间, 退费时间_In, r_Clinic.执行部门id, r_Clinic.划价人, 操作员编号_In,
               操作员姓名_In, r_Clinic.记帐单id, '', r_Clinic.是否急诊, n_组id, n_结帐id, -1 * r_Clinic.实收金额, -1, 1, r_Clinic.执行时间);
          End If;
        End Loop;
      
        --8-工本费，9-误差费
        --病人余额
        Select Nvl(Sum(实收金额), 0) Into n_实收合计 From 住院费用记录 Where NO = v_Billno And 记录性质 = 2;
        Update 病人余额
        Set 费用余额 = Nvl(费用余额, 0) + n_实收合计
        Where 病人id = n_病人id And 性质 = 1 And 类型 = 2
        Returning 费用余额 Into n_返回值;
        If Sql%RowCount = 0 Then
          Insert Into 病人余额 (病人id, 性质, 类型, 费用余额, 预交余额) Values (n_病人id, 1, 2, n_实收合计, 0);
          n_返回值 := n_实收合计;
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete 病人余额 Where 性质 = 1 And 病人id = n_病人id And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
        End If;
      
        --病人未结费用
        For r_Fee In (Select 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, Nvl(Sum(实收金额), 0) 实收合计
                      From 住院费用记录
                      Where NO = v_Billno And 记录性质 = 2
                      Group By 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id) Loop
          Update 病人未结费用
          Set 金额 = Nvl(金额, 0) + r_Fee.实收合计
          Where 病人id = n_病人id And Nvl(主页id, 0) = Nvl(主页id_In, 0) And Nvl(病人病区id, 0) = Nvl(r_Fee.病人病区id, 0) And
                Nvl(病人科室id, 0) = Nvl(r_Fee.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(r_Fee.开单部门id, 0) And
                Nvl(执行部门id, 0) = Nvl(r_Fee.执行部门id, 0) And 收入项目id + 0 = r_Fee.收入项目id And 来源途径 + 0 = 2;
          If Sql%RowCount = 0 Then
            Insert Into 病人未结费用
              (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
            Values
              (n_病人id, 主页id_In, r_Fee.病人病区id, r_Fee.病人科室id, r_Fee.开单部门id, r_Fee.执行部门id, r_Fee.收入项目id, 2, r_Fee.实收合计);
          End If;
        End Loop;
      
        --6.药品相关数据处理
        For r_Fee In (Select a.Id, a.序号, b.材料id, b.跟踪在用
                      From 住院费用记录 A, 材料特性 B
                      Where a.收费细目id = b.材料id(+) And a.No = v_Billno And a.记录性质 = 2 And a.记录状态 In (1, 3)) Loop
          Update 药品收发记录
          Set 单据 = Decode(单据, 8, 9, 24, 25, 单据), 费用id = r_Fee.Id, NO = v_Billno
          Where NO = r_Nos.No And 单据 In (8, 9, 24, 25) And
                费用id In
                (Select ID
                 From 门诊费用记录
                 Where NO = r_Nos.No And Mod(记录性质, 10) = Mod(单据_In, 10) And 记录状态 In (1, 3) And 序号 = r_Fee.序号);
          If Nvl(r_Fee.材料id, 0) <> 0 And Nvl(r_Fee.跟踪在用, 0) = 1 Then
            --更新备货材料
            Update 药品收发记录
            Set 费用id = r_Fee.Id
            Where 单据 = 21 And 费用id In (Select ID
                                       From 门诊费用记录
                                       Where NO = r_Nos.No And Mod(记录性质, 10) = Mod(单据_In, 10) And 记录状态 In (1, 3) And
                                             序号 = r_Fee.序号);
          End If;
          --更新费用审核记录
          Update 费用审核记录
          Set 转出id = r_Fee.Id, 记录状态 = Decode(立即销帐_In, 1, 2, 1), 主页id = 主页id_In, 转出人 = 操作员姓名_In, 转出时间 = 退费时间_In
          Where 费用id In
                (Select ID
                 From 门诊费用记录
                 Where NO = r_Nos.No And Mod(记录性质, 10) = Mod(单据_In, 10) And 记录状态 In (1, 3) And 序号 = r_Fee.序号) And
                性质 = 1;
          If Sql%NotFound Then
            --未找到数据时，要强制进行对应.
            Insert Into 费用审核记录
              (性质, 费用id, 病人id, 主页id, 审核人, 审核日期, 记录状态, 转出id, 转出人, 转出时间)
              Select 1, ID, n_病人id, 主页id_In, 操作员姓名_In, 退费时间_In, Decode(立即销帐_In, 1, 2, 1), r_Fee.Id, 操作员姓名_In, 退费时间_In
              From 门诊费用记录
              Where NO = r_Nos.No And Mod(记录性质, 10) = Mod(单据_In, 10) And 记录状态 In (1, 3) And 序号 = r_Fee.序号;
          End If;
        End Loop;
        Update 未发药品记录
        Set 单据 = Decode(单据, 8, 9, 24, 25, 单据), 主页id = 主页id_In, NO = v_Billno
        Where NO = No_In And 单据 In (8, 24) And 病人id = n_病人id;
      End Loop;
    Else
      --记账按照单据NO进行处理
      v_Billno     := Nextno(14);
      n_医疗小组id := Zl_医疗小组_Get(n_开单部门id, v_开单人, n_病人id, 主页id_In, 入院时间_In);
    
      Insert Into 住院费用记录
        (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 门诊标志, 病人id, 主页id, 标识号, 姓名, 性别, 年龄, 床号, 病人病区id, 病人科室id, 费别, 收费类别,
         收费细目id, 计算单位, 保险项目否, 保险大类id, 保险编码, 费用类型, 发药窗口, 付数, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额,
         记帐费用, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 执行时间, 划价人, 操作员编号, 操作员姓名, 记帐单id, 摘要, 是否急诊, 医嘱序号, 缴款组id, 医疗小组id)
        Select 病人费用记录_Id.Nextval, 2, v_Billno, 1, 序号, 从属父号, 价格父号, 0, 2, 病人id, 主页id_In, 住院号_In, 姓名, 性别, 年龄, v_床号,
               Decode(n_病区id, Null, Null, 0, Null, n_病区id), 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id, 保险编码, 费用类型,
               发药窗口, 付数, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额, 1, 开单部门id, 开单人, 入院时间_In, 退费时间_In,
               执行部门id, 执行人, 执行状态, 执行时间, 划价人, 操作员编号, 操作员姓名, 记帐单id, '门诊费用转入', 是否急诊, 医嘱序号, n_组id, n_医疗小组id
        From 门诊费用记录
        Where NO = No_In And 记录性质 = 单据_In And 记录状态 = 1 And Nvl(附加标志, 0) Not In (8, 9);
    
      If Nvl(立即销帐_In, 0) = 1 Then
        Update 门诊费用记录 Set 记录状态 = 3 Where NO = No_In And 记录性质 In (1, 2) And 记录状态 = 1;
      End If;
    
      For r_Clinic In (Select 记录性质, 序号, 从属父号, 价格父号, 病人id, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id, 保险编码,
                              费用类型, 发药窗口, 付数, Sum(数次) As 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, Sum(应收金额) As 应收金额,
                              Sum(实收金额) As 实收金额, Sum(统筹金额) As 统筹金额, 开单部门id, 开单人, 执行部门id, Max(执行人) As 执行人, 划价人,
                              Max(记帐单id) As 记帐单id, 发生时间, 实际票号, Max(执行状态) As 执行状态, Max(执行时间) As 执行时间

                       
                       From 门诊费用记录
                       Where NO = No_In And 记录性质 = 单据_In And 记录状态 In (2, 3) And Nvl(附加标志, 0) Not In (8, 9)
                       Group By 记录性质, 序号, 从属父号, 价格父号, 病人id, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id,
                                保险编码, 费用类型, 发药窗口, 付数, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, 开单部门id, 开单人, 执行部门id, 划价人, 发生时间,
                                实际票号
                       Having Sum(数次) <> 0) Loop
        Select 操作员编号, 操作员姓名
        Into n_操作员编号, v_操作员姓名
        From 门诊费用记录
        Where NO = No_In And 记录性质 = 单据_In And 记录状态 = 3 And Rownum < 2;
        Insert Into 住院费用记录
          (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 门诊标志, 病人id, 主页id, 标识号, 姓名, 性别, 年龄, 床号, 病人病区id, 病人科室id, 费别, 收费类别,
           收费细目id, 计算单位, 保险项目否, 保险大类id, 保险编码, 费用类型, 发药窗口, 付数, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额, 记帐费用,
           开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人, 划价人, 操作员编号, 操作员姓名, 记帐单id, 摘要, 缴款组id, 医疗小组id, 执行状态, 执行时间)
        Values
          (病人费用记录_Id.Nextval, 2, v_Billno, 1, r_Clinic.序号, r_Clinic.从属父号, r_Clinic.价格父号, 0, 2, r_Clinic.病人id, 主页id_In,
           住院号_In, r_Clinic.姓名, r_Clinic.性别, r_Clinic.年龄, v_床号, Decode(n_病区id, Null, Null, 0, Null, n_病区id),
           r_Clinic.病人科室id, r_Clinic.费别, r_Clinic.收费类别, r_Clinic.收费细目id, r_Clinic.计算单位, r_Clinic.保险项目否, r_Clinic.保险大类id,
           r_Clinic.保险编码, r_Clinic.费用类型, r_Clinic.发药窗口, r_Clinic.付数, r_Clinic.数次, r_Clinic.加班标志, r_Clinic.附加标志,
           r_Clinic.收入项目id, r_Clinic.收据费目, r_Clinic.标准单价, r_Clinic.应收金额, r_Clinic.实收金额, r_Clinic.统筹金额, 1,
           r_Clinic.开单部门id, r_Clinic.开单人, 入院时间_In, 退费时间_In, r_Clinic.执行部门id, r_Clinic.执行人, r_Clinic.划价人, n_操作员编号,
           v_操作员姓名, r_Clinic.记帐单id, '门诊费用转入', n_组id, n_医疗小组id, r_Clinic.执行状态, r_Clinic.执行时间);
        If Nvl(立即销帐_In, 0) = 1 Then
          Insert Into 门诊费用记录
            (ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 门诊标志, 病人id, 标识号, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位,
             保险项目否, 保险大类id, 保险编码, 费用类型, 发药窗口, 付数, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额, 记帐费用, 开单部门id,
             开单人, 发生时间, 登记时间, 执行部门id, 划价人, 操作员编号, 操作员姓名, 记帐单id, 摘要, 缴款组id, 结帐id, 结帐金额, 费用状态, 执行时间)
          Values
            (病人费用记录_Id.Nextval, r_Clinic.记录性质, No_In, r_Clinic.实际票号, 2, r_Clinic.序号, r_Clinic.从属父号, r_Clinic.价格父号, 1,
             r_Clinic.病人id, 住院号_In, r_Clinic.姓名, r_Clinic.性别, r_Clinic.年龄, r_Clinic.病人科室id, r_Clinic.费别, r_Clinic.收费类别,
             r_Clinic.收费细目id, r_Clinic.计算单位, r_Clinic.保险项目否, r_Clinic.保险大类id, r_Clinic.保险编码, r_Clinic.费用类型,
             r_Clinic.发药窗口, r_Clinic.付数, -1 * r_Clinic.数次, r_Clinic.加班标志, r_Clinic.附加标志, r_Clinic.收入项目id, r_Clinic.收据费目,
             r_Clinic.标准单价, -1 * r_Clinic.应收金额, -1 * r_Clinic.实收金额, -1 * r_Clinic.统筹金额, 1, r_Clinic.开单部门id, r_Clinic.开单人,
             r_Clinic.发生时间, 退费时间_In, r_Clinic.执行部门id, r_Clinic.划价人, 操作员编号_In, 操作员姓名_In, r_Clinic.记帐单id, '', n_组id,
             n_结帐id, -1 * r_Clinic.实收金额, 1, r_Clinic.执行时间);
        End If;
      End Loop;
    
      --8-工本费，9-误差费
      --病人余额
      Select Nvl(Sum(实收金额), 0) Into n_实收合计 From 住院费用记录 Where NO = v_Billno And 记录性质 = 2;
      Update 病人余额
      Set 费用余额 = Nvl(费用余额, 0) + n_实收合计
      Where 病人id = n_病人id And 性质 = 1 And 类型 = 2
      Returning 费用余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额 (病人id, 性质, 类型, 费用余额, 预交余额) Values (n_病人id, 1, 2, n_实收合计, 0);
        n_返回值 := n_实收合计;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete 病人余额 Where 性质 = 1 And 病人id = n_病人id And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
      End If;
    
      --病人未结费用
      For r_Fee In (Select 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, Nvl(Sum(实收金额), 0) 实收合计
                    From 住院费用记录
                    Where NO = v_Billno And 记录性质 = 2
                    Group By 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id) Loop
        Update 病人未结费用
        Set 金额 = Nvl(金额, 0) + r_Fee.实收合计
        Where 病人id = n_病人id And Nvl(主页id, 0) = Nvl(主页id_In, 0) And Nvl(病人病区id, 0) = Nvl(r_Fee.病人病区id, 0) And
              Nvl(病人科室id, 0) = Nvl(r_Fee.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(r_Fee.开单部门id, 0) And
              Nvl(执行部门id, 0) = Nvl(r_Fee.执行部门id, 0) And 收入项目id + 0 = r_Fee.收入项目id And 来源途径 + 0 = 2;
        If Sql%RowCount = 0 Then
          Insert Into 病人未结费用
            (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
          Values
            (n_病人id, 主页id_In, r_Fee.病人病区id, r_Fee.病人科室id, r_Fee.开单部门id, r_Fee.执行部门id, r_Fee.收入项目id, 2, r_Fee.实收合计);
        End If;
      End Loop;
    
      --6.药品相关数据处理
      For r_Fee In (Select a.Id, a.序号, b.材料id, b.跟踪在用
                    From 住院费用记录 A, 材料特性 B
                    Where a.收费细目id = b.材料id(+) And a.No = v_Billno And a.记录性质 = 2 And a.记录状态 In (1, 3)) Loop
        Update 药品收发记录
        Set 单据 = Decode(单据, 8, 9, 24, 25, 单据), 费用id = r_Fee.Id, NO = v_Billno
        Where NO = No_In And 单据 In (8, 9, 24, 25) And
              费用id In (Select ID
                       From 门诊费用记录
                       Where NO = No_In And 记录性质 = 单据_In And 记录状态 In (1, 3) And 序号 = r_Fee.序号);
        If Nvl(r_Fee.材料id, 0) <> 0 And Nvl(r_Fee.跟踪在用, 0) = 1 Then
          --更新备货材料
          Update 药品收发记录
          Set 费用id = r_Fee.Id
          Where 单据 = 21 And
                费用id In (Select ID
                         From 门诊费用记录
                         Where NO = No_In And 记录性质 = 单据_In And 记录状态 In (1, 3) And 序号 = r_Fee.序号);
        End If;
        --更新费用审核记录
        Update 费用审核记录
        Set 转出id = r_Fee.Id, 记录状态 = Decode(立即销帐_In, 1, 2, 1), 主页id = 主页id_In, 转出人 = 操作员姓名_In, 转出时间 = 退费时间_In
        Where 费用id In (Select ID
                       From 门诊费用记录
                       Where NO = No_In And 记录性质 = 单据_In And 记录状态 In (1, 3) And 序号 = r_Fee.序号) And 性质 = 1;
        If Sql%NotFound Then
          --未找到数据时，要强制进行对应.
          Insert Into 费用审核记录
            (性质, 费用id, 病人id, 主页id, 审核人, 审核日期, 记录状态, 转出id, 转出人, 转出时间)
            Select 1, ID, n_病人id, 主页id_In, 操作员姓名_In, 退费时间_In, Decode(立即销帐_In, 1, 2, 1), r_Fee.Id, 操作员姓名_In, 退费时间_In
            From 门诊费用记录
            Where NO = No_In And 记录性质 = 单据_In And 记录状态 In (1, 3) And 序号 = r_Fee.序号;
        End If;
      End Loop;
      Update 未发药品记录
      Set 主页id = 主页id_In, NO = v_Billno
      Where NO = No_In And 单据 In (9, 25) And 病人id = n_病人id;
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊费用转住院_Insert;
/

--133498:冉俊明,2018-11-07,支持补充结算后的门诊费用转为住院费用
Create Or Replace Procedure Zl_门诊转住院_补结算转出
(
  No_In         费用补充记录.No%Type,
  费用冲销id_In 病人预交记录.结帐id%Type,
  结算冲销id_In 病人预交记录.结帐id%Type,
  结算序号_In   病人预交记录.结算序号%Type,
  退费时间_In   住院费用记录.发生时间%Type,
  操作员编号_In 住院费用记录.操作员编号%Type,
  操作员姓名_In 住院费用记录.操作员姓名%Type,
  主页id_In     病人预交记录.主页id%Type,
  入院科室id_In 病人预交记录.科室id%Type,
  结算方式_In   病人预交记录.结算方式%Type := Null,
  误差费_In     病人预交记录.冲预交%Type := Null
) As
  --功能：对费用补充结算的门诊费用进行转住院费用处理
  --入参：
  --  结算方式_In 不为空，表示所有除预交款的非医保金额全部退为指定的结算方式；
  --              为空，表示所有除预交款的非医保金额全部转为住院预交款
  Err_Item Exception;
  v_Err_Msg Varchar2(200);
  n_返回值  病人预交记录.冲预交%Type;

  n_组id   财务缴款分组.Id%Type;
  v_误差费 结算方式.名称%Type;
  n_误差费 病人预交记录.冲预交%Type;
  n_Dec    Number; --金额小数位数 

  v_Nos    Varchar2(4000);
  n_病人id 病人预交记录.病人id%Type;

  n_已退金额 病人预交记录.冲预交%Type;
  n_未退金额 病人预交记录.冲预交%Type;
  n_冲预交   病人预交记录.冲预交%Type;
  v_结算方式 Varchar2(4000);
  v_预交no   病人预交记录.No%Type;

  --保存预交款单据
  Procedure 病人预交记录_Insert
  (
    病人id_In     病人预交记录.病人id%Type,
    金额_In       病人预交记录.金额%Type,
    结算方式_In   病人预交记录.结算方式%Type,
    收款时间_In   病人预交记录.收款时间%Type,
    结算号码_In   病人预交记录.结算号码%Type,
    卡类别id_In   病人预交记录.卡类别id%Type := Null,
    卡号_In       病人预交记录.卡号%Type := Null,
    交易流水号_In 病人预交记录.交易流水号%Type := Null,
    交易说明_In   病人预交记录.交易说明%Type := Null
  ) As
    v_预交no 病人预交记录.No%Type;
    n_返回值 病人预交记录.金额%Type;
  Begin
    If Nvl(金额_In, 0) = 0 Or 结算方式_In Is Null Then
      Return;
    End If;
  
    --一卡通，每一笔都生成一条预交款记录
    --其它，同一种结算方式只生成一条预交款记录
    Update 病人预交记录
    Set 金额 = Nvl(金额, 0) + 金额_In
    Where 记录性质 = 1 And 记录状态 = 1 And 收款时间 = 收款时间_In And 病人id + 0 = 病人id_In And 结算方式 = 结算方式_In And Nvl(卡类别id, 0) = 0;
    If Sql%RowCount = 0 Or Nvl(卡类别id_In, 0) <> 0 Then
      v_预交no := Nextno(11);
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要, 缴款组id, 预交类别,
         卡类别id, 卡号, 交易说明, 交易流水号, 结算号码)
      Values
        (病人预交记录_Id.Nextval, v_预交no, Null, 1, 1, 病人id_In, 主页id_In, 入院科室id_In, 金额_In, 结算方式_In, 收款时间_In, Null, Null, Null,
         操作员编号_In, 操作员姓名_In, '门诊转住院预交', n_组id, 2, 卡类别id_In, 卡号_In, 交易说明_In, 交易流水号_In, 结算号码_In);
    End If;
  
    Update 病人余额
    Set 预交余额 = Nvl(预交余额, 0) + 金额_In
    Where 性质 = 1 And 病人id = 病人id_In And 类型 = 2
    Returning 预交余额 Into n_返回值;
    If Sql%RowCount = 0 Then
      Insert Into 病人余额 (病人id, 性质, 类型, 预交余额, 费用余额) Values (病人id_In, 1, 2, 金额_In, 0);
      n_返回值 := 金额_In;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 病人余额 Where 病人id = 病人id_In And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
    End If;
  End;
Begin
  n_组id := Zl_Get组id(操作员姓名_In);
  --误差费
  Begin
    Select 名称 Into v_误差费 From 结算方式 Where 性质 = 9 And Rownum < 2;
  Exception
    When Others Then
      v_Err_Msg := '没有发现误差结算方式，请检查是否正确设置！';
      Raise Err_Item;
  End;
  n_误差费 := Nvl(误差费_In, 0);

  --金额小数位数 
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  Select f_List2str(Cast(Collect(a.No) As t_Strlist), ',', 1), Max(a.病人id)
  Into v_Nos, n_病人id
  From 门诊费用记录 A, 费用补充记录 B
  Where a.结帐id = b.收费结帐id And b.记录性质 = 1 And b.附加标志 = 0 And b.No = No_In;
  If v_Nos Is Null Then
    v_Err_Msg := '未找到原医保补结算数据，费用转出失败!';
    Raise Err_Item;
  End If;

  --1.更新费用审核记录 
  Update 费用审核记录
  Set 记录状态 = 2
  Where 性质 = 1 And 费用id In (Select /*+cardinality(b,10)*/
                             a.Id
                            From 门诊费用记录 A, (Select Column_Value As NO From Table(f_Str2list(v_Nos))) B
                            Where a.No = b.No And Mod(a.记录性质, 10) = 1 And a.记录状态 In (1, 3));

  --2.作废门诊费用记录 
  Update 门诊费用记录
  Set 记录状态 = 3
  Where Mod(记录性质, 10) = 1 And 记录状态 = 1 And NO In (Select Column_Value As NO From Table(f_Str2list(v_Nos)));

  For c_费用 In (Select /*+cardinality(b,10)*/
                a.No, a.序号, a.从属父号, a.价格父号, a.病人id, a.医嘱序号, a.门诊标志, a.姓名, a.性别, a.年龄, a.标识号, a.付款方式, a.病人科室id, a.费别,
                a.收费类别, a.收费细目id, a.计算单位, a.发药窗口, Sum(Nvl(a.付数, 1) * a.数次) As 数次, a.加班标志, a.附加标志, a.婴儿费, a.收入项目id, a.收据费目,
                a.标准单价, Sum(a.应收金额) As 应收金额, Sum(a.实收金额) As 实收金额, a.划价人, a.开单部门id, a.开单人, a.发生时间, a.执行部门id, a.执行人,
                Min(Decode(a.记录状态, 2, a.执行状态, 0)) - 1 As 执行状态, a.结论, Sum(a.结帐金额) As 结帐金额, Max(保险大类id) As 保险大类id,
                Max(保险项目否) As 保险项目否, Max(保险编码) As 保险编码, Max(费用类型) As 费用类型, Sum(a.统筹金额) As 统筹金额, Max(是否上传) As 是否上传, 是否急诊,
                a.挂号id, a.主页id
               From 门诊费用记录 A, (Select Column_Value As NO From Table(f_Str2list(v_Nos))) B
               Where a.No = b.No And a.记录性质 In (1, 11)
               Group By a.No, a.序号, a.从属父号, a.价格父号, a.病人id, a.医嘱序号, a.门诊标志, a.姓名, a.性别, a.年龄, a.标识号, a.付款方式, a.病人科室id,
                        a.费别, a.收费类别, a.收费细目id, a.计算单位, a.发药窗口, a.加班标志, a.附加标志, a.婴儿费, a.收入项目id, a.收据费目, a.标准单价, a.划价人,
                        a.开单部门id, a.开单人, a.发生时间, a.执行部门id, a.执行人, a.结论, 是否急诊, a.挂号id, a.主页id
               Having Nvl(Sum(Nvl(a.付数, 1) * a.数次), 0) <> 0) Loop
  
    Insert Into 门诊费用记录
      (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数,
       发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 执行时间,
       结论, 操作员编号, 操作员姓名, 结帐id, 结帐金额, 保险大类id, 保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要, 是否急诊, 缴款组id, 费用状态, 挂号id, 主页id)
    Values
      (病人费用记录_Id.Nextval, 1, c_费用.No, 2, c_费用.序号, c_费用.从属父号, c_费用.价格父号, c_费用.病人id, c_费用.医嘱序号, c_费用.门诊标志, c_费用.姓名,
       c_费用.性别, c_费用.年龄, c_费用.标识号, c_费用.付款方式, c_费用.病人科室id, c_费用.费别, c_费用.收费类别, c_费用.收费细目id, c_费用.计算单位, 1, c_费用.发药窗口,
       -1 * c_费用.数次, c_费用.加班标志, c_费用.附加标志, c_费用.婴儿费, c_费用.收入项目id, c_费用.收据费目, c_费用.标准单价, -1 * c_费用.应收金额, -1 * c_费用.实收金额,
       c_费用.划价人, c_费用.开单部门id, c_费用.开单人, c_费用.发生时间, 退费时间_In, c_费用.执行部门id, c_费用.执行人, c_费用.执行状态, Null, c_费用.结论, 操作员编号_In,
       操作员姓名_In, 费用冲销id_In, -1 * c_费用.结帐金额, c_费用.保险大类id, c_费用.保险项目否, c_费用.保险编码, c_费用.费用类型, -1 * c_费用.统筹金额, c_费用.是否上传, '',
       c_费用.是否急诊, n_组id, 0, c_费用.挂号id, c_费用.主页id);
  End Loop;
  Zl_门诊退费结算_Modify(1, n_病人id, 费用冲销id_In, Null);

  --3.作废补充结算记录（同时已进行了票据回收和医保原样退）
  Zl_费用补充记录_Delete(No_In, 结算冲销id_In, Null, 结算序号_In, 费用冲销id_In, 操作员编号_In, 操作员姓名_In, 退费时间_In);
  Update 费用补充记录 Set 费用状态 = 0 Where 结算序号 = 结算序号_In;
  --处理为医保接口已调用成功
  Update 病人预交记录
  Set 校对标志 = 2
  Where 记录性质 = 6 And 结帐id = 结算冲销id_In And 结算方式 In (Select 名称 From 结算方式 Where 性质 In (3, 4));

  --4.结算数据处理
  Select -1 * Nvl(Sum(a.冲预交), 0)
  Into n_未退金额
  From 病人预交记录 A
  Where a.结算序号 = 结算序号_In And a.结算方式 Is Null;
  If Nvl(n_误差费, 0) = 0 Then
    n_误差费 := Round(n_未退金额, n_Dec) - n_未退金额;
  End If;
  n_未退金额 := n_未退金额 - n_误差费;

  For r_预交 In (Select Case
                        When Mod(a.记录性质, 10) = 1 Then
                         1
                        When Nvl(a.卡类别id, 0) <> 0 Then
                         2
                        Else
                         0
                      End As 类型, a.结帐id, Nvl(a.冲预交, 0) As 冲预交, a.No, a.病人id, a.结算方式, a.卡类别id, a.卡号, a.交易流水号, a.交易说明,
                      a.结算号码
               From 病人预交记录 A, 结算方式 B
               Where a.结算方式 = b.名称 And a.记录状态 In (1, 3) And b.性质 Not In (3, 4, 9) And
                     a.结帐id In (Select 收费结帐id From 费用补充记录 Where 记录性质 = 1 And 附加标志 = 0 And NO = No_In)) Loop
  
    --都是单种结算方式
    If r_预交.类型 = 1 Then
      --预交款
      Zl_费用补充结算_完成退费(结算冲销id_In, Null, Null, Null, Null, Null, n_误差费, 0, 0, -1 * n_未退金额);
      Exit;
    Elsif r_预交.类型 = 2 Then
      --一卡通
      Select Nvl(Sum(金额), 0) Into n_已退金额 From 三方退款信息 Where 记录id = r_预交.结帐id;
      If r_预交.冲预交 - n_已退金额 > 0 Then
        If r_预交.冲预交 - n_已退金额 > n_未退金额 Then
          n_冲预交 := n_未退金额;
        Else
          n_冲预交 := r_预交.冲预交 - n_已退金额;
        End If;
      
        v_结算方式 := r_预交.结算方式 || '|' || -1 * n_冲预交 || '| | ';
        Zl_费用补充结算_完成退费(结算冲销id_In, v_结算方式, r_预交.卡类别id, r_预交.卡号, r_预交.交易流水号, r_预交.交易说明, n_误差费, 0, 1);
        Zl_三方退款信息_Insert(结算序号_In, r_预交.结帐id, n_冲预交, r_预交.卡号, r_预交.交易流水号, r_预交.交易说明);
      
        --转为住院预交款
        病人预交记录_Insert(r_预交.病人id, n_冲预交, r_预交.结算方式, 退费时间_In, r_预交.结算号码, r_预交.卡类别id, r_预交.卡号, r_预交.交易流水号, r_预交.交易说明);
      
        n_未退金额 := n_未退金额 - n_冲预交;
        n_误差费   := 0;
      End If;
      If n_未退金额 = 0 Then
        Exit;
      End If;
    Else
      --其它非医保结算方式
      --结算方式|结算金额|结算号码|结算摘要
      v_结算方式 := r_预交.结算方式 || '|' || n_未退金额 || '| | ';
      Zl_费用补充结算_完成退费(结算冲销id_In, v_结算方式, Null, Null, Null, Null, n_误差费, 0);
    
      --转为住院预交款
      病人预交记录_Insert(r_预交.病人id, n_未退金额, r_预交.结算方式, 退费时间_In, r_预交.结算号码);
      Exit;
    End If;
  End Loop;

  --5.转出完成处理   
  Delete From 病人预交记录 Where 结帐id = 结算冲销id_In And 结算方式 Is Null And Nvl(冲预交, 0) = 0;
  If Sql%NotFound Then
    v_Err_Msg := '还存在未缴款的数据，不能完成结算！';
    Raise Err_Item;
  End If;
  Delete From 病人预交记录 Where 结帐id = 费用冲销id_In And 结算方式 Is Null And Nvl(冲预交, 0) = 0;
  If Sql%NotFound Then
    v_Err_Msg := '还存在未缴款的数据，不能完成结算！';
    Raise Err_Item;
  End If;
  Update 病人预交记录 Set 校对标志 = 0, 会话号 = Null Where 结算序号 = 结算序号_In;

  --人员缴款余额（主要是医保）
  For c_预交 In (Select a.结算方式, a.操作员姓名, Nvl(Sum(a.冲预交), 0) As 冲预交
               From 病人预交记录 A, 结算方式 B
               Where a.结算方式 = b.名称 And b.性质 In (3, 4) And a.结算序号 = 结算序号_In
               Group By a.结算方式, a.操作员姓名) Loop
  
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) + c_预交.冲预交
    Where 收款员 = c_预交.操作员姓名 And 性质 = 1 And 结算方式 = c_预交.结算方式
    Returning 余额 Into n_返回值;
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额
        (收款员, 结算方式, 性质, 余额)
      Values
        (c_预交.操作员姓名, c_预交.结算方式, 1, c_预交.冲预交);
      n_返回值 := c_预交.冲预交;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 人员缴款余额
      Where 收款员 = c_预交.操作员姓名 And 性质 = 1 And 结算方式 = c_预交.结算方式 And Nvl(余额, 0) = 0;
    End If;
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊转住院_补结算转出;
/





------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0036' Where 编号=&n_System;
Commit;
