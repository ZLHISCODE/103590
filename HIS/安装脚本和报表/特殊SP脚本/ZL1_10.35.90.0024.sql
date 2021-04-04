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
--129571:焦博,2018-08-08,修改Oracle过程Zl_三方机构挂号_Delete,读取正确的项目ID等数据
Create Or Replace Procedure Zl_三方机构挂号_Delete
(
  单据号_In     门诊费用记录.No%Type,
  交易流水号_In 病人预交记录.交易流水号%Type,
  交易说明_In   病人预交记录.交易说明%Type,
  退号时间_In   门诊费用记录.登记时间%Type := Null,
  预交id_In     病人预交记录.Id%Type := Null
) As
  v_Error Varchar2(255);
  Err_Custom Exception;

  --该游标用于判断是否单独收病历费,及挂号汇总表处理
  Cursor c_Registinfo
  (
    v_状态     病人挂号记录.记录状态%Type,
    v_性质     病人挂号记录.记录性质%Type,
    v_无效单据 Number := 0
  ) Is
  
    Select a.发生时间, a.登记时间, b.收费细目id As 项目id, a.执行部门id As 科室id, a.执行人 As 医生姓名, c.Id As 医生id, a.号别 As 号码
    From 病人挂号记录 A, 门诊费用记录 B, 人员表 C
    Where a.记录性质 = Decode(v_无效单据, 0, v_性质, a.记录性质) And b.记录性质 = 4 And a.记录状态 = v_状态 And a.No = 单据号_In And a.No = b.No And
          a.执行人 = c.姓名(+) And b.序号 = 1 And Rownum = 1;

  r_Registrow c_Registinfo%RowType;

  --该光标用于处理人员缴款余额中退的不同结算方式的金额
  Cursor c_Opermoney Is
    Select Distinct b.结算方式, b.冲预交
    From 门诊费用记录 A, 病人预交记录 B
    Where a.结帐id = b.结帐id And a.No = 单据号_In And a.记录性质 = 4 And a.记录状态 = 3 And b.记录性质 = 4 And b.记录状态 = 3 And
          Nvl(b.冲预交, 0) <> 0;

  n_执行状态       病人挂号记录.执行状态%Type;
  n_打印id         票据打印内容.Id%Type;
  n_结帐id         门诊费用记录.结帐id%Type;
  n_原结帐id       病人预交记录.结帐id%Type;
  n_病人id         病人信息.病人id%Type;
  n_返回值         病人余额.预交余额%Type;
  n_分诊台签到排队 Number;
  n_预约挂号       Number;
  n_无效单据       Number; --无效单据没有产生费用单据
  n_挂号生成队列   Number;
  n_Count          Number;
  n_组id           财务缴款分组.Id%Type;
  d_退号时间       Date;
  v_操作员编号     人员表.编号%Type;
  v_操作员姓名     人员表.姓名%Type;
  v_合作单位       合作单位挂号汇总.合作单位%Type;
  n_预约状态       病人挂号记录.预约%Type;
  v_Temp           Varchar2(100);
  d_登记时间       病人挂号记录.登记时间%Type;
  v_号别           病人挂号记录.号别%Type;
  n_号序           病人挂号记录.号序%Type;
  d_预约时间       病人挂号记录.预约时间%Type;
  n_合作单位限制   Number(18);
  n_预约生成队列   Number;
  n_记录性质       Number;
  n_状态           Number;
  n_退号重用       Number(3);
  n_挂号排班模式   Number;
  n_记帐           门诊费用记录.记帐费用%Type;
  n_挂号id         病人挂号记录.Id%Type;
  n_已结帐         Number;
  n_返回额         病人余额.费用余额%Type;
  n_预交支付       Number(3);
  n_正常支付       Number(3);
  n_安排id         挂号安排.Id%Type;
  n_计划id         挂号安排计划.Id%Type;
  v_时间段         时间段.时间段%Type;
  d_检查开始时间   Date;
  d_检查结束时间   Date;
  d_启用时间       Date;
  n_出诊记录id     Number(18);
  Function Zl_操作员
  (
    Type_In     Integer,
    Splitstr_In Varchar2
  ) Return Varchar2 As
    n_Step Number(18);
    v_Sub  Varchar2(1000);
    --Type_In:0-获取缺省部门ID;1-获取操作员编号;2-获取操作员姓名
    -- SplitStr:格式为:部门ID,部门名称;人员ID,人员编号,人员姓名(用Zl_Identity获取的)
  Begin
    If Type_In = 0 Then
      --缺省部门
      n_Step := Instr(Splitstr_In, ',');
      v_Sub  := Substr(Splitstr_In, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 1 Then
      --操作员编码
      n_Step := Instr(Splitstr_In, ';');
      v_Sub  := Substr(Splitstr_In, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 2 Then
      --操作员姓名
      n_Step := Instr(Splitstr_In, ';');
      v_Sub  := Substr(Splitstr_In, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      Return v_Sub;
    End If;
  End;

  Procedure Zl_三方机构挂号_出诊_Delete
  (
    单据号_In     门诊费用记录.No%Type,
    交易流水号_In 病人预交记录.交易流水号%Type,
    交易说明_In   病人预交记录.交易说明%Type,
    退号时间_In   门诊费用记录.登记时间%Type := Null,
    预交id_In     病人预交记录.Id%Type := Null
  ) As
    v_Error Varchar2(255);
    Err_Custom Exception;
  
    --该游标用于判断是否单独收病历费,及挂号汇总表处理
    Cursor c_Registinfo
    (
      v_状态     病人挂号记录.记录状态%Type,
      v_性质     病人挂号记录.记录性质%Type,
      v_无效单据 Number := 0
    ) Is
      Select a.发生时间, a.登记时间, b.项目id, b.科室id, b.医生姓名, b.医生id, b.Id As 记录id, a.号别 As 号码
      From 病人挂号记录 A, 临床出诊记录 B
      Where a.记录性质 = Decode(v_无效单据, 0, v_性质, a.记录性质) And a.记录状态 = v_状态 And a.No = 单据号_In And a.出诊记录id = b.Id And
            Rownum < 2;
  
    r_Registrow c_Registinfo%RowType;
  
    --该光标用于处理人员缴款余额中退的不同结算方式的金额
    Cursor c_Opermoney Is
      Select Distinct b.结算方式, b.冲预交
      From 门诊费用记录 A, 病人预交记录 B
      Where a.结帐id = b.结帐id And a.No = 单据号_In And a.记录性质 = 4 And a.记录状态 = 3 And b.记录性质 = 4 And b.记录状态 = 3 And
            Nvl(b.冲预交, 0) <> 0;
  
    n_执行状态       病人挂号记录.执行状态%Type;
    n_打印id         票据打印内容.Id%Type;
    n_结帐id         门诊费用记录.结帐id%Type;
    n_原结帐id       病人预交记录.结帐id%Type;
    n_病人id         病人信息.病人id%Type;
    n_返回值         病人余额.预交余额%Type;
    n_分诊台签到排队 Number;
    n_预约挂号       Number;
    n_无效单据       Number; --无效单据没有产生费用单据
    n_挂号生成队列   Number;
    n_Count          Number;
    n_组id           财务缴款分组.Id%Type;
    d_退号时间       Date;
    v_操作员编号     人员表.编号%Type;
    v_操作员姓名     人员表.姓名%Type;
    v_合作单位       合作单位挂号汇总.合作单位%Type;
    n_预约状态       病人挂号记录.预约%Type;
    v_Temp           Varchar2(100);
    d_登记时间       病人挂号记录.登记时间%Type;
    v_号别           病人挂号记录.号别%Type;
    n_号序           病人挂号记录.号序%Type;
    d_预约时间       病人挂号记录.预约时间%Type;
    n_合作单位限制   Number(18);
    n_预约生成队列   Number;
    n_记录性质       Number;
    n_状态           Number;
    n_退号重用       Number(3);
    n_挂号id         病人挂号记录.Id%Type;
    n_记帐           门诊费用记录.记帐费用%Type;
    n_记录id         临床出诊记录.Id%Type;
  
    n_已结帐   Number;
    n_返回额   病人余额.费用余额%Type;
    n_预交支付 Number(3);
    n_正常支付 Number(3);
    Function Zl_操作员
    (
      Type_In     Integer,
      Splitstr_In Varchar2
    ) Return Varchar2 As
      n_Step Number(18);
      v_Sub  Varchar2(1000);
      --Type_In:0-获取缺省部门ID;1-获取操作员编号;2-获取操作员姓名
      -- SplitStr:格式为:部门ID,部门名称;人员ID,人员编号,人员姓名(用Zl_Identity获取的)
    Begin
      If Type_In = 0 Then
        --缺省部门
        n_Step := Instr(Splitstr_In, ',');
        v_Sub  := Substr(Splitstr_In, 1, n_Step - 1);
        Return v_Sub;
      End If;
      If Type_In = 1 Then
        --操作员编码
        n_Step := Instr(Splitstr_In, ';');
        v_Sub  := Substr(Splitstr_In, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, 1, n_Step - 1);
        Return v_Sub;
      End If;
      If Type_In = 2 Then
        --操作员姓名
        n_Step := Instr(Splitstr_In, ';');
        v_Sub  := Substr(Splitstr_In, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        Return v_Sub;
      End If;
    End;
  Begin
    --部门ID,部门名称;人员ID,人员编号,人员姓名
    v_Temp := Zl_Identity(0);
    If Nvl(v_Temp, ' ') = ' ' Then
      v_Error := '当前操作人员未设置对应的人员关系,不能继续。';
      Raise Err_Custom;
    End If;
    v_操作员编号 := Zl_操作员(1, v_Temp);
    v_操作员姓名 := Zl_操作员(2, v_Temp);
  
    n_组id := Zl_Get组id(v_操作员姓名);
  
    d_退号时间 := 退号时间_In;
    If d_退号时间 Is Null Then
      d_退号时间 := Sysdate;
    End If;
  
    --首先判断要退号/取消预约的记录是否存在
    Begin
      Select Decode(记录性质, 2, 1, 0), 记录性质, 登记时间, 号别, 号序, Nvl(预约时间, 发生时间), 合作单位, Nvl(预约, 0), Decode(记录状态, 0, 1, 0), 出诊记录id
      Into n_预约挂号, n_记录性质, d_登记时间, v_号别, n_号序, d_预约时间, v_合作单位, n_预约状态, n_无效单据, n_记录id
      From 病人挂号记录
      Where NO = 单据号_In And 记录状态 In (0, 1) And Rownum < 2;
    Exception
      When Others Then
        n_预约挂号 := -1;
    End;
  
    If n_预约挂号 = -1 Then
      v_Error := '单据可能已经被退号或单据输入错误!';
      Raise Err_Custom;
    End If;
  
    Begin
      Select Nvl(记帐费用, 0), Decode(Sign(Nvl(结帐id, 0)), 0, 0, 1)
      Into n_记帐, n_已结帐
      From 门诊费用记录
      Where 记录性质 = 4 And NO = 单据号_In And 记录状态 In (1, 3) And Rownum < 2;
    Exception
      When Others Then
        n_记帐   := 0;
        n_已结帐 := 0;
    End;
  
    --预约检查是否添加合作单位控制
    --如果设置了合作单位控制 则
    Select Count(0) Into n_合作单位限制 From 临床出诊挂号控制记录 Where 类型 = 1 And 性质 = 1 And Rownum < 2;
    --更新挂号序号状态
    n_退号重用 := Zl_To_Number(zl_GetSysParameter('已退序号允许挂号', 1111));
    If n_退号重用 = 0 Then
      Update 临床出诊序号控制 Set 挂号状态 = 4 Where 记录id = n_记录id And (序号 = n_号序 Or 备注 = To_Char(n_号序));
    Else
      Update 临床出诊序号控制
      Set 挂号状态 = 0, 类型 = Null, 名称 = Null, 操作员姓名 = Null, 工作站名称 = Null
      Where 记录id = n_记录id And (序号 = n_号序 Or 备注 = To_Char(n_号序));
    End If;
    If Nvl(n_预约挂号, 0) = 1 Or Nvl(n_无效单据, 0) = 1 Then
      If Nvl(n_无效单据, 0) = 0 Then
        --N天内不能取消预约号
        n_Count := Zl_To_Number(zl_GetSysParameter('N天内不能取消预约号', 1111));
        If n_Count <> 0 Then
          If Trunc(Sysdate - n_Count) < d_登记时间 Then
            v_Error := '不能退掉预约在' || To_Char(Trunc(Sysdate - n_Count), 'yyyy-mm-dd') || '以前的预约单!';
            Raise Err_Custom;
          End If;
        End If;
      End If;
    
      n_状态 := Case n_无效单据
                When 1 Then
                 0
                Else
                 1
              End;
      --减少已约数
      Open c_Registinfo(n_状态, 2, n_无效单据);
      Fetch c_Registinfo
        Into r_Registrow;
    
      Update 病人挂号汇总
      Set 已约数 = Nvl(已约数, 0) - n_预约状态, 已挂数 = Nvl(已挂数, 0) - Decode(n_预约状态, 0, 1, 0)
      Where 日期 = Trunc(r_Registrow.发生时间) And 科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And
            Nvl(医生姓名, '医生') = Nvl(r_Registrow.医生姓名, '医生') And Nvl(医生id, 0) = Nvl(r_Registrow.医生id, 0) And
            (号码 = r_Registrow.号码 Or 号码 Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into 病人挂号汇总
          (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已约数, 已挂数)
        Values
          (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
           Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, -n_预约状态, Decode(n_预约状态, 0, 1, 0));
      End If;
    
      Update 临床出诊记录
      Set 已约数 = Nvl(已约数, 0) - n_预约状态, 已挂数 = Nvl(已挂数, 0) - Decode(n_预约状态, 0, 1, 0)
      Where ID = n_记录id;
      Close c_Registinfo;
    
      If Nvl(n_无效单据, 0) = 0 Then
        --删除门诊费用记录
        Delete From 门诊费用记录 Where NO = 单据号_In And 记录性质 = 4 And 记录状态 = 0;
        --如果预约生成队列时需要清除队列
        n_挂号生成队列 := Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113));
        If Nvl(n_挂号生成队列, 0) = 1 Then
          n_预约生成队列 := Zl_To_Number(zl_GetSysParameter('预约生成队列', 1113));
          If Nvl(n_预约生成队列, 0) = 1 Then
            --要删除队列
            For v_挂号 In (Select ID, 诊室, 执行部门id, 执行人 From 病人挂号记录 Where NO = 单据号_In) Loop
              Zl_排队叫号队列_Delete(v_挂号.执行部门id, v_挂号.Id);
            End Loop;
          End If;
        End If;
      End If;
    Else
      Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
    
      --更新挂号序号状态
    
      --病人就诊状态
      Select 病人id
      Into n_病人id
      From 门诊费用记录
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And 序号 = 1;
    
      If n_病人id Is Not Null Then
        Update 病人信息 Set 就诊状态 = 0, 就诊诊室 = Null Where 病人id = n_病人id;
        --删除门诊号相关处理,只有当只有一条挂号记录并且病人建档日期与挂号日期近似时才会处理
      End If;
    
      --门诊费用记录
      Insert Into 门诊费用记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位,
         付数, 数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 执行人, 操作员编号, 操作员姓名, 发生时间,
         登记时间, 结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 缴款组id)
        Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别,
               收费细目id, 计算单位, 付数, -数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, -应收金额, -实收金额, 开单部门id, 开单人, 执行部门id, 执行人,
               v_操作员编号, v_操作员姓名, 发生时间, d_退号时间, n_结帐id,
               Decode(n_记帐, 1, Decode(Nvl(n_已结帐, 0), 0, -1 * 实收金额, Null), -1 * 结帐金额), 保险项目否, 保险大类id, -1 * 统筹金额, 摘要,
               Decode(Nvl(附加标志, 0), 9, 1, 0), 保险编码, 费用类型, n_组id
        From 门诊费用记录
        Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In;
    
      --原始记录
      If n_记帐 = 1 And Nvl(n_已结帐, 0) = 0 Then
        Update 门诊费用记录
        Set 记录状态 = 3, 结帐id = n_结帐id, 结帐金额 = 实收金额
        Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In;
      Else
        Update 门诊费用记录 Set 记录状态 = 3 Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In;
      End If;
    
      n_原结帐id := 0;
      If n_记帐 = 0 Then
        --获取结帐ID
        Select Nvl(结帐id, 0)
        Into n_原结帐id
        From 门诊费用记录
        Where 记录性质 = 4 And 记录状态 = 3 And NO = 单据号_In And Rownum < 2;
      End If;
    
      If n_记帐 = 1 Then
        --记帐
        For c_费用 In (Select 实收金额, 病人科室id, 开单部门id, 执行部门id, 收入项目id
                     From 门诊费用记录
                     Where 记录性质 = 4 And 记录状态 = 3 And NO = 单据号_In And Nvl(记帐费用, 0) = 1) Loop
          --病人余额
          Update 病人余额
          Set 费用余额 = Nvl(费用余额, 0) - Nvl(c_费用.实收金额, 0)
          Where 病人id = Nvl(n_病人id, 0) And 性质 = 1 And 类型 = 1
          Returning 费用余额 Into n_返回额;
        
          If Sql%RowCount = 0 Then
            Insert Into 病人余额
              (病人id, 性质, 类型, 费用余额, 预交余额)
            Values
              (n_病人id, 1, 1, -1 * Nvl(c_费用.实收金额, 0), 0);
            n_返回额 := Nvl(c_费用.实收金额, 0);
          End If;
          If Nvl(n_返回额, 0) = 0 Then
            Delete 病人余额
            Where 病人id = Nvl(n_病人id, 0) And 性质 = 1 And 类型 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
          End If;
          --病人未结费用
          Update 病人未结费用
          Set 金额 = Nvl(金额, 0) - Nvl(c_费用.实收金额, 0)
          Where 病人id = n_病人id And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(病人科室id, 0) = Nvl(c_费用.病人科室id, 0) And
                Nvl(开单部门id, 0) = Nvl(c_费用.开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(c_费用.执行部门id, 0) And
                收入项目id + 0 = c_费用.收入项目id And 来源途径 + 0 = 1;
        
          If Sql%RowCount = 0 Then
            Insert Into 病人未结费用
              (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
            Values
              (n_病人id, Null, Null, c_费用.病人科室id, c_费用.开单部门id, c_费用.执行部门id, c_费用.收入项目id, 1, -1 * Nvl(c_费用.实收金额, 0));
          End If;
        End Loop;
        Delete 病人未结费用
        Where 病人id = n_病人id And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(金额, 0) = 0 And 来源途径 + 0 = 1;
      End If;
    
      If n_记帐 = 0 Then
        Begin
          Select 1
          Into n_预交支付
          From 病人预交记录
          Where Mod(记录性质, 10) = 1 And 记录状态 = 1 And 结帐id = n_原结帐id And Rownum < 2;
        Exception
          When Others Then
            n_预交支付 := 0;
        End;
        Begin
          Select 1
          Into n_正常支付
          From 病人预交记录
          Where Mod(记录性质, 10) = 4 And 记录状态 = 1 And 结帐id = n_原结帐id And Rownum < 2;
        Exception
          When Others Then
            n_正常支付 := 0;
        End;
        If n_预交支付 = 1 And n_正常支付 = 1 Then
          v_Error := '不能处理多种结算方式,请检查传入的退号单据是否正确!';
          Raise Err_Custom;
        End If;
        If n_预交支付 = 1 Then
          --原样退回预交
          If Nvl(预交id_In, 0) = 0 Then
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,
               冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
              Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号,
                     d_退号时间, v_操作员姓名, v_操作员编号, -1 * 冲预交, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
              From 病人预交记录
              Where 记录性质 In (1, 11) And 结帐id = n_原结帐id And Nvl(冲预交, 0) <> 0;
          Else
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,
               冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
              Select 预交id_In, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, d_退号时间,
                     v_操作员姓名, v_操作员编号, -1 * 冲预交, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
              From 病人预交记录
              Where 记录性质 In (1, 11) And 结帐id = n_原结帐id And Nvl(冲预交, 0) <> 0;
          End If;
          --处理病人预交余额
          For c_预交 In (Select 病人id, 预交类别, -1 * Sum(Nvl(冲预交, 0)) As 冲预交
                       From 病人预交记录
                       Where 记录性质 In (1, 11) And 结帐id = n_结帐id
                       Group By 病人id, 预交类别) Loop
            Update 病人余额
            Set 预交余额 = Nvl(预交余额, 0) + Nvl(c_预交.冲预交, 0)
            Where 病人id = c_预交.病人id And 类型 = Nvl(c_预交.预交类别, 2) And 性质 = 1
            Returning 预交余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 病人余额
                (病人id, 预交余额, 性质, 类型)
              Values
                (c_预交.病人id, Nvl(c_预交.冲预交, 0), 1, Nvl(c_预交.预交类别, 2));
              n_返回值 := Nvl(c_预交.冲预交, 0);
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 病人余额
              Where 病人id = c_预交.病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
            End If;
          End Loop;
        Else
          If Nvl(预交id_In, 0) = 0 Then
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易流水号, 交易说明,
               合作单位, 结算序号, 卡类别id, 结算性质)
              Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_退号时间, v_操作员编号, v_操作员姓名, -冲预交,
                     n_结帐id, n_组id, 交易流水号_In, 交易说明_In, 合作单位, n_结帐id, 卡类别id, 结算性质
              From 病人预交记录
              Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_原结帐id;
          Else
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易流水号, 交易说明,
               合作单位, 结算序号, 卡类别id, 结算性质)
              Select 预交id_In, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_退号时间, v_操作员编号, v_操作员姓名, -冲预交, n_结帐id,
                     n_组id, 交易流水号_In, 交易说明_In, 合作单位, n_结帐id, 卡类别id, 结算性质
              From 病人预交记录
              Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_原结帐id;
          End If;
          Update 病人预交记录 Set 记录状态 = 3 Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_原结帐id;
        End If;
        --退卡收回票据(可能上次挂号使用票据,不能收回)
        --从最后一次的打印内容中取
        Select Max(ID)
        Into n_打印id
        From (Select b.Id
               From 票据使用明细 A, 票据打印内容 B
               Where a.打印id = b.Id And a.性质 = 1 And b.数据性质 = 4 And b.No = 单据号_In
               Order By a.使用时间 Desc)
        Where Rownum < 2;
        If n_打印id Is Not Null Then
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
            Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_退号时间, v_操作员姓名, 票据金额
            From 票据使用明细
            Where 打印id = n_打印id And 性质 = 1;
        End If;
      End If;
    
      --相关汇总表的处理
    
      --病人挂号汇总
      Open c_Registinfo(1, 1);
      Fetch c_Registinfo
        Into r_Registrow;
    
      If c_Registinfo%RowCount = 0 Then
        --只收病历费时无号别,不处理
        Close c_Registinfo;
      Else
      
        --需要确定是否预约挂号
        --1.如果是退预约挂号产生的挂号记录,则需要减已挂数和其中已接数
        --2.如果是正常挂号,则只减已挂数
        Begin
          Select Decode(预约, Null, 0, 0, 0, 1), 执行状态
          Into n_预约挂号, n_执行状态
          From 病人挂号记录
          Where NO = 单据号_In And 记录状态 = 1 And Rownum = 1;
        Exception
          When Others Then
            n_预约挂号 := 0;
        End;
        --0-等待接诊,1-完成就诊,2-正在就诊,-1标记为不就诊
        If n_执行状态 > 0 Then
          If n_执行状态 = 1 Then
            v_Error := '该病人已经完成就诊,不能再退号!';
          Else
            v_Error := '该病人正在就诊, 不能退号!';
          End If;
          Raise Err_Custom;
        End If;
      
        Update 病人挂号汇总
        Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - n_预约状态, 已约数 = Nvl(已约数, 0) - n_预约状态
        Where 日期 = Trunc(r_Registrow.发生时间) And 科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And
              Nvl(医生姓名, '医生') = Nvl(r_Registrow.医生姓名, '医生') And Nvl(医生id, 0) = Nvl(r_Registrow.医生id, 0) And
              (号码 = r_Registrow.号码 Or 号码 Is Null);
      
        If Sql%RowCount = 0 Then
          Insert Into 病人挂号汇总
            (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 其中已接收)
          Values
            (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
             Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, -1, -1 * n_预约挂号);
        End If;
      
        Update 临床出诊记录
        Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - n_预约状态, 已约数 = Nvl(已约数, 0) - n_预约状态
        Where ID = n_记录id;
        Close c_Registinfo;
      End If;
    
      --人员缴款余额(包括个人帐户等的结算金额,不含退冲预交款)
      If n_记帐 = 0 Then
        For r_Opermoney In c_Opermoney Loop
          Update 人员缴款余额
          Set 余额 = Nvl(余额, 0) + (-1 * r_Opermoney.冲预交)
          Where 收款员 = v_操作员姓名 And 结算方式 = r_Opermoney.结算方式 And 性质 = 1
          Returning 余额 Into n_返回值;
          If Sql%RowCount = 0 Then
            Insert Into 人员缴款余额
              (收款员, 结算方式, 性质, 余额)
            Values
              (v_操作员姓名, r_Opermoney.结算方式, 1, -1 * r_Opermoney.冲预交);
            n_返回值 := r_Opermoney.冲预交;
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 人员缴款余额
            Where 收款员 = v_操作员姓名 And 结算方式 = r_Opermoney.结算方式 And 性质 = 1 And Nvl(余额, 0) = 0;
          End If;
        End Loop;
      End If;
    
      n_挂号生成队列 := Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113));
      If n_挂号生成队列 <> 0 Then
        --要删除队列
        For v_挂号 In (Select ID, 诊室, 执行部门id, 执行人 From 病人挂号记录 Where NO = 单据号_In) Loop
          n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113, Nvl(v_挂号.执行部门id, 0)));
          If Nvl(n_分诊台签到排队, 0) = 0 Then
            Zl_排队叫号队列_Delete(v_挂号.执行部门id, v_挂号.Id);
          End If;
        End Loop;
      End If;
    
      --医保产生的就诊登记记录
      Delete From 就诊登记记录
      Where (病人id, 主页id, 就诊时间) In (Select 病人id, 主页id, 发生时间 From 病人挂号记录 Where NO = 单据号_In);
    End If;
  
    If Nvl(n_无效单据, 0) = 0 Then
      Update 病人挂号记录 Set 记录状态 = 3 Where NO = 单据号_In And 记录状态 = 1;
      If Sql%NotFound Then
        v_Error := '未找到挂号单据,请检查!';
        Raise Err_Custom;
      End If;
      Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
      Insert Into 病人挂号记录
        (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 操作员编号, 操作员姓名,
         复诊, 号序, 预约, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 出诊记录id)
        Select n_挂号id, NO, 记录性质, 2, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, d_退号时间, 发生时间,
               v_操作员编号, v_操作员姓名, 复诊, 号序, 预约, 接收人, 接收时间, 交易流水号_In, 交易说明_In, 合作单位, 出诊记录id
        From 病人挂号记录
        Where NO = 单据号_In And 记录状态 = 3;
    End If;
    --消息推送
    Begin
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 2, 单据号_In;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_003(n_挂号id, 单据号_In);
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Error || '+[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End;

Begin
  n_挂号排班模式 := To_Number(Substr(Nvl(zl_GetSysParameter('挂号排班模式'), 0), 1, 1));
  If n_挂号排班模式 = 1 Then
    --出诊表排班模式
    Zl_三方机构挂号_出诊_Delete(单据号_In, 交易流水号_In, 交易说明_In, 退号时间_In, 预交id_In);
  Else
    v_Temp := zl_GetSysParameter(256);
    If v_Temp Is Null Or Substr(v_Temp, 1, 1) = '0' Then
      Null;
    Else
      Begin
        d_启用时间 := To_Date(Substr(v_Temp, 3), 'YYYY-MM-DD hh24:mi:ss');
      Exception
        When Others Then
          d_启用时间 := Null;
      End;
    End If;
    --部门ID,部门名称;人员ID,人员编号,人员姓名
    v_Temp := Zl_Identity(0);
    If Nvl(v_Temp, ' ') = ' ' Then
      v_Error := '当前操作人员未设置对应的人员关系,不能继续。';
      Raise Err_Custom;
    End If;
    v_操作员编号 := Zl_操作员(1, v_Temp);
    v_操作员姓名 := Zl_操作员(2, v_Temp);
  
    n_组id := Zl_Get组id(v_操作员姓名);
  
    d_退号时间 := 退号时间_In;
    If d_退号时间 Is Null Then
      d_退号时间 := Sysdate;
    End If;
  
    --首先判断要退号/取消预约的记录是否存在
    Begin
      Select Decode(记录性质, 2, 1, 0), 记录性质, 登记时间, 号别, 号序, Nvl(预约时间, 发生时间), 合作单位, Nvl(预约, 0), Decode(记录状态, 0, 1, 0)
      Into n_预约挂号, n_记录性质, d_登记时间, v_号别, n_号序, d_预约时间, v_合作单位, n_预约状态, n_无效单据
      From 病人挂号记录
      Where NO = 单据号_In And 记录状态 In (0, 1) And Rownum <= 1;
    Exception
      When Others Then
        n_预约挂号 := -1;
    End;
  
    If n_预约挂号 = -1 Then
      v_Error := '单据可能已经被退号或单据输入错误!';
      Raise Err_Custom;
    End If;
  
    Begin
      Select Nvl(记帐费用, 0), Decode(Sign(Nvl(结帐id, 0)), 0, 0, 1)
      Into n_记帐, n_已结帐
      From 门诊费用记录
      Where 记录性质 = 4 And NO = 单据号_In And 记录状态 In (1, 3) And Rownum < 2;
    Exception
      When Others Then
        n_记帐   := 0;
        n_已结帐 := 0;
    End;
  
    Begin
      Select a.Id Into n_安排id From 挂号安排 A Where a.号码 = v_号别;
    Exception
      When Others Then
        n_安排id := -1;
    End;
  
    Begin
      Select ID
      Into n_计划id
      From 挂号安排计划
      Where 安排id = n_安排id And 审核时间 Is Not Null And
            Nvl(生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) =
            (Select Max(a.生效时间) As 生效
             From 挂号安排计划 A
             Where a.审核时间 Is Not Null And d_预约时间 Between Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                   Nvl(a.失效时间, To_Date('3000-01-01', 'yyyy-mm-dd')) And a.安排id = n_安排id) And
            d_预约时间 Between Nvl(生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
            Nvl(失效时间, To_Date('3000-01-01', 'yyyy-mm-dd'));
    Exception
      When Others Then
        n_计划id := 0;
    End;
    If Nvl(n_计划id, 0) = 0 Then
      Select Decode(To_Char(d_预约时间, 'D'), '1', 周日, '2', 周一, '3', 周二, '4', 周三, '5', 周四, '6', 周五, '7', 周六, Null)
      Into v_时间段
      From 挂号安排
      Where ID = n_安排id;
    Else
      Select Decode(To_Char(d_预约时间, 'D'), '1', 周日, '2', 周一, '3', 周二, '4', 周三, '5', 周四, '6', 周五, '7', 周六, Null)
      Into v_时间段
      From 挂号安排计划
      Where ID = n_计划id;
    End If;
  
    If v_时间段 Is Not Null And d_启用时间 Is Not Null Then
      --检查是否跨模式挂号安排
      Select To_Date(To_Char(d_预约时间, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
             To_Date(To_Char(d_预约时间, 'yyyy-mm-dd') || ' ' || To_Char(终止时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
      Into d_检查开始时间, d_检查结束时间
      From 时间段
      Where 时间段 = v_时间段 And 站点 Is Null And 号类 Is Null;
      If d_检查开始时间 > d_检查结束时间 Then
        d_检查结束时间 := d_检查结束时间 + 1;
      End If;
      If d_检查结束时间 > d_启用时间 Then
        --获取出诊记录id
        Select Max(a.Id)
        Into n_出诊记录id
        From 临床出诊记录 A, 临床出诊号源 B
        Where a.号源id = b.Id And b.号码 = v_号别 And 上班时段 = v_时间段 And d_预约时间 Between 开始时间 And 终止时间;
      End If;
    End If;
  
    --预约检查是否添加合作单位控制
    --如果设置了合作单位控制 则
    Select Count(0) Into n_合作单位限制 From 合作单位安排控制 Where Rownum = 1;
    --更新挂号序号状态
    n_退号重用 := Zl_To_Number(zl_GetSysParameter('已退序号允许挂号', 1111));
    If n_退号重用 = 0 Then
      Update 挂号序号状态
      Set 状态 = 4
      Where 号码 = v_号别 And 序号 = n_号序 And 日期 Between Trunc(d_预约时间) And Trunc(d_预约时间 + 1) - 1 / 24 / 60 / 60;
    Else
      Delete 挂号序号状态
      Where 号码 = v_号别 And 序号 = n_号序 And 日期 Between Trunc(d_预约时间) And Trunc(d_预约时间 + 1) - 1 / 24 / 60 / 60;
    End If;
    If Nvl(n_预约挂号, 0) = 1 Or Nvl(n_无效单据, 0) = 1 Then
      If Nvl(n_无效单据, 0) = 0 Then
        --N天内不能取消预约号
        n_Count := Zl_To_Number(zl_GetSysParameter('N天内不能取消预约号', 1111));
        If n_Count <> 0 Then
          If Trunc(Sysdate - n_Count) < d_登记时间 Then
            v_Error := '不能退掉预约在' || To_Char(Trunc(Sysdate - n_Count), 'yyyy-mm-dd') || '以前的预约单!';
            Raise Err_Custom;
          End If;
        End If;
      End If;
    
      n_状态 := Case n_无效单据
                When 1 Then
                 0
                Else
                 1
              End;
      --减少已约数
      Open c_Registinfo(n_状态, 2, n_无效单据);
      Fetch c_Registinfo
        Into r_Registrow;
    
      Update 病人挂号汇总
      Set 已约数 = Nvl(已约数, 0) - n_预约状态, 已挂数 = Nvl(已挂数, 0) - Decode(n_预约状态, 0, 1, 0)
      Where 日期 = Trunc(r_Registrow.发生时间) And 科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And
            Nvl(医生姓名, '医生') = Nvl(r_Registrow.医生姓名, '医生') And Nvl(医生id, 0) = Nvl(r_Registrow.医生id, 0) And
            (号码 = r_Registrow.号码 Or 号码 Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into 病人挂号汇总
          (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已约数, 已挂数)
        Values
          (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
           Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, -n_预约状态, Decode(n_预约状态, 0, 1, 0));
      End If;
    
      If n_出诊记录id Is Not Null Then
        Update 临床出诊记录
        Set 已约数 = Nvl(已约数, 0) - n_预约状态, 已挂数 = Nvl(已挂数, 0) - Decode(n_预约状态, 0, 1, 0)
        Where ID = n_出诊记录id And Nvl(已约数, 0) > 0;
        Update 临床出诊序号控制 Set 挂号状态 = Null, 操作员姓名 = Null Where 记录id = n_出诊记录id And 序号 = n_号序;
      End If;
      If Nvl(n_合作单位限制, 0) <> 0 And v_合作单位 Is Not Null And Nvl(n_预约状态, 0) <> 0 Then
        Update 合作单位挂号汇总
        Set 已约数 = Nvl(已约数, 0) - n_预约状态, 已接数 = Nvl(已接数, 0) - Decode(n_预约状态, 0, 1, 0)
        Where 日期 = Trunc(r_Registrow.发生时间) And (号码 = r_Registrow.号码 Or 号码 Is Null) And 合作单位 = v_合作单位 And
              序号 = Nvl(n_号序, 0);
        If Sql%RowCount = 0 Then
          Insert Into 合作单位挂号汇总
            (日期, 号码, 已约数, 合作单位, 序号, 已接数)
          Values
            (Trunc(r_Registrow.发生时间), r_Registrow.号码, -n_预约状态, v_合作单位, Nvl(n_号序, 0), -decode(n_预约状态, 0, 1, 0));
        End If;
      End If;
      Close c_Registinfo;
    
      If Nvl(n_无效单据, 0) = 0 Then
        --删除门诊费用记录
        Delete From 门诊费用记录 Where NO = 单据号_In And 记录性质 = 4 And 记录状态 = 0;
        --如果预约生成队列时需要清除队列
        n_挂号生成队列 := Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113));
        If Nvl(n_挂号生成队列, 0) = 1 Then
          n_预约生成队列 := Zl_To_Number(zl_GetSysParameter('预约生成队列', 1113));
          If Nvl(n_预约生成队列, 0) = 1 Then
            --要删除队列
            For v_挂号 In (Select ID, 诊室, 执行部门id, 执行人 From 病人挂号记录 Where NO = 单据号_In) Loop
              Zl_排队叫号队列_Delete(v_挂号.执行部门id, v_挂号.Id);
            End Loop;
          End If;
        End If;
      End If;
    Else
      Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
    
      --更新挂号序号状态
    
      --病人就诊状态
      Select 病人id
      Into n_病人id
      From 门诊费用记录
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And 序号 = 1;
    
      If n_病人id Is Not Null Then
        Update 病人信息 Set 就诊状态 = 0, 就诊诊室 = Null Where 病人id = n_病人id;
        --删除门诊号相关处理,只有当只有一条挂号记录并且病人建档日期与挂号日期近似时才会处理
      End If;
    
      --门诊费用记录
      Insert Into 门诊费用记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位,
         付数, 数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 执行人, 操作员编号, 操作员姓名, 发生时间,
         登记时间, 结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 缴款组id)
        Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别,
               收费细目id, 计算单位, 付数, -数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, -应收金额, -实收金额, 开单部门id, 开单人, 执行部门id, 执行人,
               v_操作员编号, v_操作员姓名, 发生时间, d_退号时间, n_结帐id,
               Decode(n_记帐, 1, Decode(Nvl(n_已结帐, 0), 0, -1 * 实收金额, Null), -1 * 结帐金额), 保险项目否, 保险大类id, -1 * 统筹金额, 摘要,
               Decode(Nvl(附加标志, 0), 9, 1, 0), 保险编码, 费用类型, n_组id
        From 门诊费用记录
        Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In;
    
      --原始记录
      If n_记帐 = 1 And Nvl(n_已结帐, 0) = 0 Then
        Update 门诊费用记录
        Set 记录状态 = 3, 结帐id = n_结帐id, 结帐金额 = 实收金额
        Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In;
      Else
        Update 门诊费用记录 Set 记录状态 = 3 Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In;
      End If;
    
      n_原结帐id := 0;
      If n_记帐 = 0 Then
        --获取结帐ID
        Select Nvl(结帐id, 0)
        Into n_原结帐id
        From 门诊费用记录
        Where 记录性质 = 4 And 记录状态 = 3 And NO = 单据号_In And Rownum < 2;
      End If;
    
      If n_记帐 = 1 Then
        --记帐
        For c_费用 In (Select 实收金额, 病人科室id, 开单部门id, 执行部门id, 收入项目id
                     From 门诊费用记录
                     Where 记录性质 = 4 And 记录状态 = 3 And NO = 单据号_In And Nvl(记帐费用, 0) = 1) Loop
          --病人余额
          Update 病人余额
          Set 费用余额 = Nvl(费用余额, 0) - Nvl(c_费用.实收金额, 0)
          Where 病人id = Nvl(n_病人id, 0) And 性质 = 1 And 类型 = 1
          Returning 费用余额 Into n_返回额;
        
          If Sql%RowCount = 0 Then
            Insert Into 病人余额
              (病人id, 性质, 类型, 费用余额, 预交余额)
            Values
              (n_病人id, 1, 1, -1 * Nvl(c_费用.实收金额, 0), 0);
            n_返回额 := Nvl(c_费用.实收金额, 0);
          End If;
          If Nvl(n_返回额, 0) = 0 Then
            Delete 病人余额
            Where 病人id = Nvl(n_病人id, 0) And 性质 = 1 And 类型 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
          End If;
          --病人未结费用
          Update 病人未结费用
          Set 金额 = Nvl(金额, 0) - Nvl(c_费用.实收金额, 0)
          Where 病人id = n_病人id And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(病人科室id, 0) = Nvl(c_费用.病人科室id, 0) And
                Nvl(开单部门id, 0) = Nvl(c_费用.开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(c_费用.执行部门id, 0) And
                收入项目id + 0 = c_费用.收入项目id And 来源途径 + 0 = 1;
        
          If Sql%RowCount = 0 Then
            Insert Into 病人未结费用
              (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
            Values
              (n_病人id, Null, Null, c_费用.病人科室id, c_费用.开单部门id, c_费用.执行部门id, c_费用.收入项目id, 1, -1 * Nvl(c_费用.实收金额, 0));
          End If;
        End Loop;
        Delete 病人未结费用
        Where 病人id = n_病人id And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(金额, 0) = 0 And 来源途径 + 0 = 1;
      End If;
    
      If n_记帐 = 0 Then
        Begin
          Select 1
          Into n_预交支付
          From 病人预交记录
          Where Mod(记录性质, 10) = 1 And 记录状态 = 1 And 结帐id = n_原结帐id And Rownum < 2;
        Exception
          When Others Then
            n_预交支付 := 0;
        End;
      
        Begin
          Select 1
          Into n_正常支付
          From 病人预交记录
          Where Mod(记录性质, 10) = 4 And 记录状态 = 1 And 结帐id = n_原结帐id And Rownum < 2;
        Exception
          When Others Then
            n_正常支付 := 0;
        End;
      
        If n_预交支付 = 1 And n_正常支付 = 1 Then
          v_Error := '不能处理多种结算方式,请检查传入的退号单据是否正确!';
          Raise Err_Custom;
        End If;
      
        If n_预交支付 = 1 Then
          --原样退回预交
          If Nvl(预交id_In, 0) = 0 Then
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,
               冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
              Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号,
                     d_退号时间, v_操作员姓名, v_操作员编号, -1 * 冲预交, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
              From 病人预交记录
              Where 记录性质 In (1, 11) And 结帐id = n_原结帐id And Nvl(冲预交, 0) <> 0;
          Else
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,
               冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
              Select 预交id_In, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, d_退号时间,
                     v_操作员姓名, v_操作员编号, -1 * 冲预交, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
              From 病人预交记录
              Where 记录性质 In (1, 11) And 结帐id = n_原结帐id And Nvl(冲预交, 0) <> 0;
          End If;
          --处理病人预交余额
          For c_预交 In (Select 病人id, 预交类别, -1 * Sum(Nvl(冲预交, 0)) As 冲预交
                       From 病人预交记录
                       Where 记录性质 In (1, 11) And 结帐id = n_结帐id
                       Group By 病人id, 预交类别) Loop
            Update 病人余额
            Set 预交余额 = Nvl(预交余额, 0) + Nvl(c_预交.冲预交, 0)
            Where 病人id = c_预交.病人id And 类型 = Nvl(c_预交.预交类别, 2) And 性质 = 1
            Returning 预交余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 病人余额
                (病人id, 预交余额, 性质, 类型)
              Values
                (c_预交.病人id, Nvl(c_预交.冲预交, 0), 1, Nvl(c_预交.预交类别, 2));
              n_返回值 := Nvl(c_预交.冲预交, 0);
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 病人余额
              Where 病人id = c_预交.病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
            End If;
          End Loop;
        Else
          If Nvl(预交id_In, 0) = 0 Then
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易流水号, 交易说明,
               合作单位, 结算序号, 卡类别id, 结算性质)
              Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_退号时间, v_操作员编号, v_操作员姓名, -冲预交,
                     n_结帐id, n_组id, 交易流水号_In, 交易说明_In, 合作单位, n_结帐id, 卡类别id, 结算性质
              From 病人预交记录
              Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_原结帐id;
          Else
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易流水号, 交易说明,
               合作单位, 结算序号, 卡类别id, 结算性质)
              Select 预交id_In, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_退号时间, v_操作员编号, v_操作员姓名, -冲预交, n_结帐id,
                     n_组id, 交易流水号_In, 交易说明_In, 合作单位, n_结帐id, 卡类别id, 结算性质
              From 病人预交记录
              Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_原结帐id;
          End If;
        
          Update 病人预交记录 Set 记录状态 = 3 Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_原结帐id;
        End If;
        --退卡收回票据(可能上次挂号使用票据,不能收回)
        --从最后一次的打印内容中取
        Select Max(ID)
        Into n_打印id
        From (Select b.Id
               From 票据使用明细 A, 票据打印内容 B
               Where a.打印id = b.Id And a.性质 = 1 And b.数据性质 = 4 And b.No = 单据号_In
               Order By a.使用时间 Desc)
        Where Rownum < 2;
        If n_打印id Is Not Null Then
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
            Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_退号时间, v_操作员姓名, 票据金额
            From 票据使用明细
            Where 打印id = n_打印id And 性质 = 1;
        End If;
      End If;
    
      --相关汇总表的处理
    
      --病人挂号汇总
      Open c_Registinfo(1, 1);
      Fetch c_Registinfo
        Into r_Registrow;
    
      If c_Registinfo%RowCount = 0 Then
        --只收病历费时无号别,不处理
        Close c_Registinfo;
      Else
      
        --需要确定是否预约挂号
        --1.如果是退预约挂号产生的挂号记录,则需要减已挂数和其中已接数
        --2.如果是正常挂号,则只减已挂数
        Begin
          Select Decode(预约, Null, 0, 0, 0, 1), 执行状态
          Into n_预约挂号, n_执行状态
          From 病人挂号记录
          Where NO = 单据号_In And 记录状态 = 1 And Rownum = 1;
        Exception
          When Others Then
            n_预约挂号 := 0;
        End;
        --0-等待接诊,1-完成就诊,2-正在就诊,-1标记为不就诊
        If n_执行状态 > 0 Then
          If n_执行状态 = 1 Then
            v_Error := '该病人已经完成就诊,不能再退号!';
          Else
            v_Error := '该病人正在就诊, 不能退号!';
          End If;
          Raise Err_Custom;
        End If;
      
        Update 病人挂号汇总
        Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - n_预约状态, 已约数 = Nvl(已约数, 0) - n_预约状态
        Where 日期 = Trunc(r_Registrow.发生时间) And 科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And
              Nvl(医生姓名, '医生') = Nvl(r_Registrow.医生姓名, '医生') And Nvl(医生id, 0) = Nvl(r_Registrow.医生id, 0) And
              (号码 = r_Registrow.号码 Or 号码 Is Null);
      
        If Sql%RowCount = 0 Then
          Insert Into 病人挂号汇总
            (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 其中已接收)
          Values
            (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
             Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, -1, -1 * n_预约挂号);
        End If;
        If n_出诊记录id Is Not Null Then
          Update 临床出诊记录
          Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - n_预约状态, 已约数 = Nvl(已约数, 0) - n_预约状态
          Where ID = n_出诊记录id And Nvl(已约数, 0) > 0;
          Update 临床出诊序号控制 Set 挂号状态 = Null, 操作员姓名 = Null Where 记录id = n_出诊记录id And 序号 = n_号序;
        End If;
        If Nvl(n_合作单位限制, 0) <> 0 And v_合作单位 Is Not Null And Nvl(n_预约状态, 0) <> 0 Then
          Update 合作单位挂号汇总
          Set 已接数 = Nvl(已接数, 0) - 1, 已约数 = Nvl(已约数, 0) - n_预约挂号
          Where 日期 = Trunc(r_Registrow.发生时间) And (号码 = r_Registrow.号码 Or 号码 Is Null) And 合作单位 = v_合作单位 And
                序号 = Nvl(n_号序, 0);
          If Sql%RowCount = 0 Then
            Insert Into 合作单位挂号汇总
              (日期, 号码, 已约数, 合作单位, 已接数, 序号)
            Values
              (Trunc(r_Registrow.发生时间), r_Registrow.号码, -1, v_合作单位, -1 * n_预约挂号, Nvl(n_号序, 0));
          End If;
        End If;
        Close c_Registinfo;
      End If;
    
      --人员缴款余额(包括个人帐户等的结算金额,不含退冲预交款)
      If n_记帐 = 0 Then
        For r_Opermoney In c_Opermoney Loop
          Update 人员缴款余额
          Set 余额 = Nvl(余额, 0) + (-1 * r_Opermoney.冲预交)
          Where 收款员 = v_操作员姓名 And 结算方式 = r_Opermoney.结算方式 And 性质 = 1
          Returning 余额 Into n_返回值;
          If Sql%RowCount = 0 Then
            Insert Into 人员缴款余额
              (收款员, 结算方式, 性质, 余额)
            Values
              (v_操作员姓名, r_Opermoney.结算方式, 1, -1 * r_Opermoney.冲预交);
            n_返回值 := r_Opermoney.冲预交;
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 人员缴款余额
            Where 收款员 = v_操作员姓名 And 结算方式 = r_Opermoney.结算方式 And 性质 = 1 And Nvl(余额, 0) = 0;
          End If;
        End Loop;
      End If;
    
      n_挂号生成队列 := Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113));
      If n_挂号生成队列 <> 0 Then
        --要删除队列
        For v_挂号 In (Select ID, 诊室, 执行部门id, 执行人 From 病人挂号记录 Where NO = 单据号_In) Loop
          n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113, 1, Nvl(v_挂号.执行部门id, 0)));
          If Nvl(n_分诊台签到排队, 0) = 0 Then
            Zl_排队叫号队列_Delete(v_挂号.执行部门id, v_挂号.Id);
          End If;
        End Loop;
      End If;
    
      --医保产生的就诊登记记录
      Delete From 就诊登记记录
      Where (病人id, 主页id, 就诊时间) In (Select 病人id, 主页id, 发生时间 From 病人挂号记录 Where NO = 单据号_In);
    End If;
  
    If Nvl(n_无效单据, 0) = 0 Then
      Update 病人挂号记录 Set 记录状态 = 3 Where NO = 单据号_In And 记录状态 = 1;
      If Sql%NotFound Then
        v_Error := '未找到挂号单据,请检查!';
        Raise Err_Custom;
      End If;
      Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
      Insert Into 病人挂号记录
        (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 操作员编号, 操作员姓名,
         复诊, 号序, 预约, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位)
        Select n_挂号id, NO, 记录性质, 2, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, d_退号时间, 发生时间,
               v_操作员编号, v_操作员姓名, 复诊, 号序, 预约, 接收人, 接收时间, 交易流水号_In, 交易说明_In, 合作单位
        From 病人挂号记录
        Where NO = 单据号_In And 记录状态 = 3;
    End If;
    --消息推送
    Begin
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 2, 单据号_In;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_003(n_挂号id, 单据号_In);
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Error || '+[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_三方机构挂号_Delete;
/

--128435:焦博,2018-08-08,修改Oracle过程Zl_Third_Saveexes,新增入参节点医嘱ID和出参节点单据号
Create Or Replace Procedure Zl_Third_Saveexes
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:保存费用记录
  --入参:Xml_In:
  --<IN>
  --  <PATIID></PATIID>  //病人ID
  --  <PAGEID></PAGEID>  //主页ID
  --  <MZBZ></MZBZ>   //门诊标志，1-门诊，2-住院
  --  <JZBZ></JZBZ>   //记帐标志，0-收费，1-记帐
  --  <CZY></CZY>   //操作员
  --  <CZSJ></CZSJ>   //操作时间
  --  <KDR></KDR>  //开单人
  --  <KDKSID></KDKSID>  //开单科室ID
  --  <YQBH></YQBH>  //院区编号
  --  <MXLIST>
  --    <MX>
  --      <YZID><YZID> //医嘱ID
  --      <SFXMID></SFXMID>  //收费细目ID
  --      <SL></SL>   //数次
  --      <ZXR></ZXR>  //执行人,表示完全执行
  --      <ZXKSID></ZXKSID>  //执行科室ID
  --    </MX>
  --    ...
  --  </MXLIST>
  --</IN>
  --出参:Xml_Out
  --<OUTPUT>
  --   <RESULT></RESULT> //true-表示保存成功;false-表示保存失败
  --   <DJH></DJH>  //单据号
  --   <ERROR>      //失败时返回
  --     <MSG></MSG>   //详细错误提示
  --   </ERROR>
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  n_病人id     病人信息.病人id%Type;
  n_主页id     病案主页.主页id%Type;
  n_门诊标志   门诊费用记录.门诊标志%Type; --1-门诊，2-住院
  n_记帐标志   Number(2); --0-收费，1-记帐
  v_操作员姓名 门诊费用记录.操作员姓名%Type;
  d_操作时间   门诊费用记录.登记时间%Type;
  v_开单人     门诊费用记录.开单人%Type;
  n_开单部门id 门诊费用记录.开单部门id%Type;
  v_站点       部门表.站点%Type;
  Xml_明细列表 Xmltype;

  v_No           门诊费用记录.No%Type;
  n_标识号       门诊费用记录.标识号%Type;
  v_姓名         门诊费用记录.姓名%Type;
  v_性别         门诊费用记录.性别%Type;
  v_年龄         门诊费用记录.年龄%Type;
  v_费别         门诊费用记录.费别%Type;
  v_付款方式编码 门诊费用记录.付款方式%Type;
  v_付款方式名称 病人信息.医疗付款方式%Type;
  n_病区id       住院费用记录.病人病区id%Type;
  n_科室id       门诊费用记录.病人科室id%Type;
  v_床号         住院费用记录.床号%Type;
  v_操作员编号   门诊费用记录.操作员编号%Type;
  d_出院日期     病案主页.出院日期%Type;
  d_入院日期     病案主页.入院日期%Type;

  Type Ty_Rec_Bill Is Record(
    序号       门诊费用记录.序号%Type,
    价格父号   门诊费用记录.价格父号%Type,
    收费细目id 门诊费用记录.收费细目id%Type,
    收费类别   门诊费用记录.收费类别%Type,
    计算单位   门诊费用记录.计算单位%Type,
    收入项目id 门诊费用记录.收入项目id%Type,
    收据费目   门诊费用记录.收据费目%Type,
    数次       门诊费用记录.数次%Type,
    标准单价   门诊费用记录.标准单价%Type,
    应收金额   门诊费用记录.应收金额%Type,
    实收金额   门诊费用记录.实收金额%Type,
    执行人     门诊费用记录.执行人%Type,
    执行部门id 门诊费用记录.执行部门id%Type,
    费用摘要   门诊费用记录.摘要%Type,
    医嘱id     门诊费用记录.医嘱序号%Type);

  Type Ty_Tb_Bill Is Table Of Ty_Rec_Bill;
  c_Bill Ty_Tb_Bill := Ty_Tb_Bill();

  n_单价小数 Number;
  n_金额小数 Number;

  v_Temp     Varchar2(4000);
  v_价格等级 收费价目.价格等级%Type;
  v_普通等级 收费价目.价格等级%Type;
  v_药品等级 收费价目.价格等级%Type;
  v_卫材等级 收费价目.价格等级%Type;

  n_序号         门诊费用记录.序号%Type;
  n_价格父号     门诊费用记录.价格父号%Type;
  n_当前价格父号 门诊费用记录.价格父号%Type;
  n_价格         门诊费用记录.标准单价%Type;
  n_剩余数       门诊费用记录.数次%Type;
  n_实收金额     门诊费用记录.实收金额%Type;
  d_登记时间     门诊费用记录.登记时间%Type;

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  Procedure Zl_Third_Checkdata
  (
    病人来源_In   In Number,
    病人id_In     门诊费用记录.病人id%Type,
    主页id_In     病案主页.主页id%Type,
    病区id_In     病案主页.当前病区id%Type,
    收费细目id_In 门诊费用记录.收费细目id%Type,
    数次_In       门诊费用记录.数次%Type,
    实收金额_In   门诊费用记录.实收金额%Type,
    执行部门id_In 门诊费用记录.执行部门id%Type,
    是否记账_In   Number := 0
  ) Is
  
    --入参：
    --        病人来源_In  1-门诊/2-住院
    --        是否记账_In 是否记账费用:0-收费/1-记帐
    n_跟踪在用 材料特性.跟踪在用%Type;
    n_在用分批 材料特性.在用分批%Type;
    n_是否变价 收费项目目录.是否变价%Type;
    n_库存     药品库存.可用数量%Type;
    n_项目名称 收费项目目录.名称%Type;
  
    n_检查方式  材料出库检查.检查方式%Type;
    v_收费类别  门诊费用记录.收费类别%Type;
    n_报警方法  记帐报警线.报警方法%Type;
    n_报警值    记帐报警线.报警值%Type;
    v_报警标志2 记帐报警线.报警标志2%Type;
    v_报警标志3 记帐报警线.报警标志3%Type;
    n_标志      Number;
    v_类别名称  收费项目类别.名称%Type;
    n_担保金额  门诊费用记录.实收金额%Type;
    v_担保      Varchar2(100);
    n_剩余款额  门诊费用记录.实收金额%Type;
    n_当日金额  门诊费用记录.实收金额%Type;
  
    v_Err_Msg Varchar2(255);
    Err_Item Exception;
  Begin
  
    --1.药品/卫材库存检查，要分开
    Begin
      Select b.是否变价, b.名称, b.类别, a.跟踪在用, Decode(b.类别, '4', a.在用分批, c.药房分批)
      Into n_是否变价, n_项目名称, v_收费类别, n_跟踪在用, n_在用分批
      From 材料特性 A, 收费项目目录 B, 药品规格 C
      Where b.Id = a.材料id(+) And b.Id = c.药品id(+) And b.Id = 收费细目id_In;
    Exception
      When Others Then
        v_Err_Msg := '未发现收费项目！';
        Raise Err_Item;
    End;
  
    If Instr('5,6,7', v_收费类别) > 0 Or v_收费类别 = '4' And Nvl(n_跟踪在用, 0) = 1 Then
      Select Nvl(Sum(a.可用数量), 0)
      Into n_库存
      From 药品库存 A
      Where a.性质 = 1 And a.库房id = 执行部门id_In And (Nvl(a.批次, 0) = 0 Or a.效期 Is Null Or a.效期 > Trunc(Sysdate)) And
            a.药品id = 收费细目id_In;
      If n_库存 < 数次_In Then
        If Nvl(n_在用分批, 0) = 1 Or Nvl(n_是否变价, 0) = 1 Then
          v_Err_Msg := '[' || n_项目名称 || ']的当前可用库存不足输入数量！';
          Raise Err_Item;
        Else
          Begin
            If Instr('5,6,7', v_收费类别) > 0 Then
              Select a.检查方式 Into n_检查方式 From 药品出库检查 A Where a.库房id = 执行部门id_In;
            Else
              Select a.检查方式 Into n_检查方式 From 材料出库检查 A Where a.库房id = 执行部门id_In;
            End If;
          Exception
            When Others Then
              n_检查方式 := 0;
          End;
          If Nvl(n_检查方式, 0) = 2 Then
            v_Err_Msg := '[' || n_项目名称 || ']的当前可用库存不足输入数量！';
            Raise Err_Item;
          End If;
        End If;
      End If;
    End If;
  
    --2.住院记账和门诊记账要进行记账分类报警
    If Nvl(是否记账_In, 0) = 1 Then
      --记帐分类报警
      Begin
        Select Nvl(报警方法, 1) As 报警方法, 报警值, 报警标志2, 报警标志3
        Into n_报警方法, n_报警值, v_报警标志2, v_报警标志3
        From 记帐报警线
        Where 适用病人 = Zl_Patiwarnscheme(病人id_In, 主页id_In) And 病区id = 病区id_In;
      Exception
        When Others Then
          n_报警方法 := 0;
      End;
    
      If n_报警值 Is Not Null And Nvl(n_报警方法, 0) > 0 Then
        If v_报警标志2 Is Not Null Then
          If v_报警标志2 = '-' Or Instr(v_报警标志2, v_收费类别) > 0 Then
            n_标志 := 2;
          End If;
          If v_报警标志2 = '-' Then
            v_类别名称 := ''; --所有类别时,不必提示具体的类别
          End If;
        End If;
        If Nvl(n_标志, 0) = 0 And v_报警标志3 Is Not Null Then
          If v_报警标志3 = '-' Or Instr(v_报警标志3, v_收费类别) > 0 Then
            n_标志 := 3;
          End If;
          If v_报警标志3 = '-' Then
            v_类别名称 := ''; --所有类别时,不必提示具体的类别
          End If;
        End If;
      
        If n_报警方法 = 1 Then
          --累计费用报警(低于)\
          n_担保金额 := Zl_Patientsurety(病人id_In, 主页id_In);
          If n_担保金额 > 0 Then
            v_担保 := '(含担保额：' || n_担保金额 || ')';
          End If;
        
          Select Nvl(Sum(预交余额 - 费用余额), 0)
          Into n_剩余款额
          From 病人余额
          Where 性质 = 1 And 类型 = Decode(病人来源_In, 1, 2, 1) And 病人id = 病人id_In;
        
          n_剩余款额 := n_剩余款额 + n_担保金额 - Nvl(实收金额_In, 0);
          If n_标志 = 2 Then
            --预交款耗尽时禁止记帐
            If n_剩余款额 < 0 Then
              v_Err_Msg := '剩余款' || v_担保 || '已经耗尽，' || v_类别名称 || '禁止记帐。';
              Raise Err_Item;
            End If;
          Elsif n_标志 = 3 Then
            --低于报警值禁止记帐
            If n_剩余款额 < n_报警值 Then
              v_Err_Msg := '剩余款' || v_担保 || '低于' || v_类别名称 || '报警值：' || n_报警值 || '，禁止记帐。';
              Raise Err_Item;
            End If;
          End If;
        Elsif n_报警方法 = 2 Then
          --每日费用报警(高于)
          If n_标志 = 3 Then
            --高于报警值禁止记帐
            n_当日金额 := Zl_Patidaycharge(病人id_In);
            n_当日金额 := n_当日金额 + Nvl(实收金额_In, 0);
            If n_当日金额 > n_报警值 Then
              v_Err_Msg := '当日费用：' || n_当日金额 || '，高于' || v_类别名称 || '报警值：' || n_报警值 || '，禁止记帐。';
              Raise Err_Item;
            End If;
          End If;
        End If;
      End If;
    End If;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Zl_Third_Checkdata;

Begin
  --获取入参
  Select Extractvalue(Value(A), 'IN/PATIID'),
         Decode(Extractvalue(Value(A), 'IN/PAGEID'), 0, Null, Extractvalue(Value(A), 'IN/PAGEID')),
         Nvl(Extractvalue(Value(A), 'IN/MZBZ'), 0), Nvl(Extractvalue(Value(A), 'IN/JZBZ'), 0),
         Extractvalue(Value(A), 'IN/CZY'), To_Date(Extractvalue(Value(A), 'IN/CZSJ'), 'yyyy-mm-dd hh24:mi:ss'),
         Extractvalue(Value(A), 'IN/KDR'),
         Decode(Extractvalue(Value(A), 'IN/KDKSID'), 0, Null, Extractvalue(Value(A), 'IN/KDKSID')),
         Extractvalue(Value(A), 'IN/YQBH'), Extract(Value(A), 'IN/MXLIST')
  Into n_病人id, n_主页id, n_门诊标志, n_记帐标志, v_操作员姓名, d_操作时间, v_开单人, n_开单部门id, v_站点, Xml_明细列表
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  Begin
    Select a.编号 Into v_操作员编号 From 人员表 A Where a.姓名 = v_操作员姓名;
  Exception
    When Others Then
      v_Err_Msg := '未查找到操作员信息，单据保存失败！';
      Raise Err_Item;
  End;

  Begin
    If n_门诊标志 = 2 Then
      Select a.姓名, a.性别, a.年龄, a.费别, a.住院号, a.当前病区id, Nvl(a.出院科室id, n_开单部门id), a.出院病床, c.编码, c.名称, a.出院日期, a.入院日期,
             a.当前病区id
      Into v_姓名, v_性别, v_年龄, v_费别, n_标识号, n_病区id, n_科室id, v_床号, v_付款方式编码, v_付款方式名称, d_出院日期, d_入院日期, n_病区id
      From 病案主页 A, 病人信息 B, 医疗付款方式 C
      Where a.病人id = b.病人id And b.病人id = n_病人id And a.主页id = n_主页id And a.医疗付款方式 = c.名称(+);
    Else
      Select a.姓名, a.性别, a.年龄, a.费别, a.门诊号, n_开单部门id, b.编码, b.名称
      Into v_姓名, v_性别, v_年龄, v_费别, n_标识号, n_科室id, v_付款方式编码, v_付款方式名称
      From 病人信息 A, 医疗付款方式 B
      Where a.病人id = n_病人id And a.医疗付款方式 = b.名称(+);
    End If;
  Exception
    When Others Then
      v_Err_Msg := '未查找到病人信息，单据保存失败！';
      Raise Err_Item;
  End;

  --住院病人发生时间的检查
  If Nvl(n_门诊标志, 0) = 2 Then
    If d_出院日期 Is Not Null Then
      If d_操作时间 > d_出院日期 Then
        v_Err_Msg := '强制对出院病人记帐时，费用时间不能大于病人出院时间:' || To_Char(d_出院日期, 'yyyy-mm-dd hh24:mi:ss') || '！';
        Raise Err_Item;
      End If;
    End If;
    If d_入院日期 Is Not Null Then
      If d_操作时间 < d_入院日期 Then
        v_Err_Msg := '费用的发生时间不能小于病人的入院时间:' || To_Char(d_入院日期, 'yyyy-mm-dd hh24:mi:ss') || '！';
        Raise Err_Item;
      End If;
    End If;
  End If;

  If v_费别 Is Null Then
    Select Max(名称) Into v_费别 From 费别 Where 缺省标志 = 1 And Rownum < 2;
    If v_费别 Is Null Then
      v_Err_Msg := '无法确定病人费别，请检查缺省费别是否正确设置！';
      Raise Err_Item;
    End If;
  End If;

  --金额及单价小数位数
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')), Zl_To_Number(Nvl(zl_GetSysParameter(157), '5'))
  Into n_金额小数, n_单价小数
  From Dual;

  --价格等级
  v_Temp := Zl_Get_Pricegrade(v_站点, n_病人id, n_主页id, v_付款方式名称);
  For c_价格等级 In (Select Rownum As 序号, Column_Value As 价格等级 From Table(f_Str2list(v_Temp, '|'))) Loop
    If c_价格等级.序号 = 1 Then
      v_普通等级 := c_价格等级.价格等级;
    End If;
    If c_价格等级.序号 = 2 Then
      v_药品等级 := c_价格等级.价格等级;
    End If;
    If c_价格等级.序号 = 3 Then
      v_卫材等级 := c_价格等级.价格等级;
    End If;
  End Loop;

  n_序号 := 1;
  For c_明细 In (Select a.收费细目id, a.数次, a.执行人, a.执行科室id, a.摘要, a.医嘱id, b.类别, b.名称, b.计算单位, b.是否变价, b.屏蔽费别, b.撤档时间
               From (Select Extractvalue(Value(J), '/MX/SFXMID') As 收费细目id, Extractvalue(Value(J), '/MX/SL') As 数次,
                             Extractvalue(Value(J), '/MX/ZXR') As 执行人, Extractvalue(Value(J), '/MX/ZXKSID') As 执行科室id,
                             Extractvalue(Value(J), '/MX/FYZY') As 摘要, Extractvalue(Value(J), '/MX/YZID') As 医嘱id
                      From Table(Xmlsequence(Extract(Xml_明细列表, '/MXLIST/MX'))) J) A, 收费项目目录 B
               Where a.收费细目id = b.Id) Loop
  
    If Nvl(c_明细.撤档时间, Sysdate + 1) < Sysdate Then
      v_Err_Msg := '“' || c_明细.名称 || '”已停用，单据保存失败！';
      Raise Err_Item;
    End If;
  
    n_价格父号 := n_序号;
    If c_明细.类别 = '4' Then
      v_价格等级 := v_卫材等级;
    Elsif Instr(',5,6,7,', ',' || c_明细.类别 || ',') > 0 Then
      v_价格等级 := v_药品等级;
    Else
      v_价格等级 := v_普通等级;
    End If;
    For c_收费价目 In (Select a.收入项目id, b.收据费目, a.现价, a.缺省价格
                   From 收费价目 A, 收入项目 B
                   Where a.收入项目id = b.Id And Sysdate Between a.执行日期 And Nvl(a.终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')) And
                         a.收费细目id = c_明细.收费细目id And
                         (a.价格等级 = v_价格等级 Or
                         (a.价格等级 Is Null And Not Exists
                          (Select 1
                            From 收费价目
                            Where a.收费细目id = 收费细目id And 价格等级 = v_价格等级 And Sysdate Between 执行日期 And
                                  Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')))))) Loop
    
      If n_价格父号 = n_序号 Then
        n_当前价格父号 := Null;
      Else
        n_当前价格父号 := n_价格父号;
      End If;
    
      If Instr(',4,5,6,7,', ',' || c_明细.类别 || ',') = 0 Then
        --普通收费项目
        If Nvl(c_明细.是否变价, 0) = 0 Then
          n_价格 := Nvl(c_收费价目.现价, 0);
        Else
          n_价格 := Nvl(c_收费价目.缺省价格, 0);
        End If;
      Else
        --药品卫材
        v_Temp   := Zl_Get_Retailprice(c_明细.收费细目id, v_价格等级, c_明细.执行科室id, c_明细.数次) || '||';
        n_价格   := Nvl(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1), 0);
        v_Temp   := Substr(v_Temp, Instr(v_Temp, '|') + 1);
        n_剩余数 := Substr(v_Temp, 1, Instr(v_Temp, '|') - 1);
      
        If Nvl(n_剩余数, 0) <> 0 And Nvl(c_明细.是否变价, 0) = 1 Then
          --数量未分解完毕
          If Instr(',5,6,7,', ',' || c_明细.类别 || ',') > 0 Then
            v_Err_Msg := '时价药品"' || c_明细.名称 || '"库存不足，无法计算价格！';
          Else
            v_Err_Msg := '时价卫生材料"' || c_明细.名称 || '"库存不足，无法计算价格！';
          End If;
          Raise Err_Item;
        End If;
      End If;
    
      c_Bill.Extend;
      c_Bill(c_Bill.Count).序号 := n_序号;
      c_Bill(c_Bill.Count).价格父号 := n_当前价格父号;
      c_Bill(c_Bill.Count).收费细目id := c_明细.收费细目id;
      c_Bill(c_Bill.Count).收费类别 := c_明细.类别;
      c_Bill(c_Bill.Count).计算单位 := c_明细.计算单位;
      c_Bill(c_Bill.Count).收入项目id := c_收费价目.收入项目id;
      c_Bill(c_Bill.Count).收据费目 := c_收费价目.收据费目;
      c_Bill(c_Bill.Count).数次 := c_明细.数次;
      c_Bill(c_Bill.Count).标准单价 := Round(n_价格, n_单价小数);
      c_Bill(c_Bill.Count).应收金额 := Round(c_Bill(c_Bill.Count).标准单价 * c_Bill(c_Bill.Count).数次, n_金额小数);
      If Nvl(c_明细.屏蔽费别, 0) = 1 Or c_Bill(c_Bill.Count).应收金额 = 0 Then
        c_Bill(c_Bill.Count).实收金额 := c_Bill(c_Bill.Count).应收金额;
      Else
        v_Temp := Zl_Actualmoney(v_费别, c_明细.收费细目id, c_收费价目.收入项目id, c_Bill(c_Bill.Count).应收金额, c_明细.数次, c_明细.执行科室id) || '::';
        v_Temp := Substr(v_Temp, Instr(v_Temp, ':') + 1);
        n_实收金额 := Substr(v_Temp, 1, Instr(v_Temp, ':') - 1);
        c_Bill(c_Bill.Count).实收金额 := Round(Nvl(n_实收金额, 0), n_金额小数);
      End If;
      c_Bill(c_Bill.Count).执行人 := c_明细.执行人;
      c_Bill(c_Bill.Count).执行部门id := c_明细.执行科室id;
      c_Bill(c_Bill.Count).费用摘要 := c_明细.摘要;
      c_Bill(c_Bill.Count).医嘱id := c_明细.医嘱id;
    
      n_序号 := n_序号 + 1;
      Zl_Third_Checkdata(n_门诊标志, n_病人id, n_主页id, n_病区id, c_明细.收费细目id, c_明细.数次, c_Bill(c_Bill.Count).实收金额, c_明细.执行科室id,
                         n_记帐标志);
    End Loop;
  End Loop;

  --单据号
  If (n_门诊标志 = 1 And n_记帐标志 = 1) Or n_门诊标志 = 2 Then
    v_No := Nextno(14);
  Else
    v_No := Nextno(13);
  End If;

  --保存单据
  d_登记时间 := Sysdate;
  For I In 1 .. c_Bill.Count Loop
    If n_门诊标志 = 1 Then
      If n_记帐标志 = 0 Then
        --门诊划价
        Zl_门诊划价记录_Insert(v_No, c_Bill(I).序号, n_病人id, n_主页id, n_标识号, v_付款方式编码, v_姓名, v_性别, v_年龄, v_费别, 0, n_科室id,
                         n_开单部门id, v_开单人, Null, c_Bill(I).收费细目id, c_Bill(I).收费类别, c_Bill(I).计算单位, Null, 1, c_Bill(I).数次,
                         0, c_Bill(I).执行部门id, c_Bill(I).价格父号, c_Bill(I).收入项目id, c_Bill(I).收据费目, c_Bill(I).标准单价,
                         c_Bill(I).应收金额, c_Bill(I).实收金额, d_操作时间, d_登记时间, Null, v_操作员姓名, c_Bill(I).费用摘要, c_Bill(I).医嘱id);
      
        If c_Bill(I).执行人 Is Not Null Then
          --标记为完全执行
          Update 门诊费用记录
          Set 执行状态 = 1, 执行人 = c_Bill(I).执行人, 执行时间 = Sysdate
          Where 记录性质 = 1 And NO = v_No And 序号 = c_Bill(I).序号;
        End If;
      Else
        --门诊记帐
        Zl_门诊记帐记录_Insert(v_No, c_Bill(I).序号, n_病人id, n_标识号, v_姓名, v_性别, v_年龄, v_费别, 0, 0, n_科室id, n_开单部门id, v_开单人, Null,
                         c_Bill(I).收费细目id, c_Bill(I).收费类别, c_Bill(I).计算单位, 1, c_Bill(I).数次, 0, c_Bill(I).执行部门id,
                         c_Bill(I).价格父号, c_Bill(I).收入项目id, c_Bill(I).收据费目, c_Bill(I).标准单价, c_Bill(I).应收金额,
                         c_Bill(I).实收金额, d_操作时间, d_登记时间, Null, 0, v_操作员编号, v_操作员姓名, Null, c_Bill(I).费用摘要, c_Bill(I).医嘱id);
      
        If c_Bill(I).执行人 Is Not Null Then
          --标记为完全执行
          Update 门诊费用记录
          Set 执行状态 = 1, 执行人 = c_Bill(I).执行人, 执行时间 = Sysdate
          Where 记录性质 = 2 And NO = v_No And 序号 = c_Bill(I).序号;
        End If;
      End If;
    Elsif n_门诊标志 = 2 Then
      --住院记帐
      Zl_住院记帐记录_Insert(v_No, c_Bill(I).序号, n_病人id, n_主页id, n_标识号, v_姓名, v_性别, v_年龄, v_床号, v_费别, n_病区id, n_科室id, 0, 0,
                       n_开单部门id, v_开单人, Null, c_Bill(I).收费细目id, c_Bill(I).收费类别, c_Bill(I).计算单位, 0, Null, Null, 1,
                       c_Bill(I).数次, 0, c_Bill(I).执行部门id, c_Bill(I).价格父号, c_Bill(I).收入项目id, c_Bill(I).收据费目,
                       c_Bill(I).标准单价, c_Bill(I).应收金额, c_Bill(I).实收金额, Null, d_操作时间, d_登记时间, Null, 0, v_操作员编号, v_操作员姓名,
                       0, Null, Null, c_Bill(I).费用摘要, 0, c_Bill(I).医嘱id);
    
      If c_Bill(I).执行人 Is Not Null Then
        --标记为完全执行
        Update 住院费用记录
        Set 执行状态 = 1, 执行人 = c_Bill(I).执行人, 执行时间 = Sysdate
        Where 记录性质 = 2 And NO = v_No And 序号 = c_Bill(I).序号;
      End If;
    End If;
  End Loop;

  Xml_Out := Xmltype('<OUTPUT><RESULT>true</RESULT><DJH>' || v_No || '</DJH></OUTPUT>');
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Saveexes;
/

--129714:蒋廷中,2018-08-06,处理执行时间在开始时间之前
Create Or Replace Procedure Zl_病人医嘱执行_Update
(
  原执行时间_In 病人医嘱执行.执行时间%Type,
  医嘱id_In     病人医嘱执行.医嘱id%Type,
  发送号_In     病人医嘱执行.发送号%Type,
  要求时间_In   病人医嘱执行.要求时间%Type,
  本次数次_In   病人医嘱执行.本次数次%Type,
  执行摘要_In   病人医嘱执行.执行摘要%Type,
  执行人_In     病人医嘱执行.执行人%Type,
  执行时间_In   病人医嘱执行.执行时间%Type,
  执行结果_In   病人医嘱执行.执行结果%Type := 1,
  未执行原因_In 病人医嘱执行.说明%Type := Null,
  单独执行_In   Number := 0,
  操作员编号_In 人员表.编号%Type := Null,
  操作员姓名_In 人员表.姓名%Type := Null,
  执行部门id_In 门诊费用记录.执行部门id%Type := 0
  --参数：医嘱ID_IN=单独执行的医嘱ID，检验组合为显示的检验项目的ID。
  --执行部门id_In=仅处理指定执行部门的费用，不传或传入0时不限制执行部门
) Is
  --除了要执行的主记录,还包含了附加手术,检查部位的记录
  --手术麻醉,中药煎法,采集方法单独控制,检验组合只填写在第一个项目上，但执行状态相同
  v_Temp     Varchar2(255); 
  v_人员姓名 人员表.姓名%Type;

  v_组id        病人医嘱记录.Id%Type;
  v_诊疗类别    病人医嘱记录.诊疗类别%Type;
  v_执行结果old 病人医嘱执行.执行结果%Type;
  n_本次数次old 病人医嘱执行.本次数次%Type;

  v_病人来源 病人医嘱记录.病人来源%Type;
  v_费用性质 病人医嘱发送.记录性质%Type;

  n_执行次数 Number;
  n_剩余次数 Number;
  n_执行状态 Number;
  n_发送数次 Number;
  n_单次数次 Number;
  v_Count    Number;
  n_登记数次 Number;
  d_要求时间 Date;
  d_执行时间 Date;
  d_开始时间 Date;

  d_登记时间   病人医嘱执行.登记时间%Type;
  n_取消执行   Number;
  n_Diffday    Number(18, 3);
  n_执行科室id Number;

  v_Date  Date;
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --当前操作人员
  If 操作员编号_In Is Not Null And 操作员姓名_In Is Not Null Then
    v_人员姓名 := 操作员姓名_In;
  Else
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  End If;

  Select Sysdate Into v_Date From Dual;
  Select Nvl(执行结果, 1), Nvl(本次数次, 0), 登记时间
  Into v_执行结果old, n_本次数次old, d_登记时间
  From 病人医嘱执行
  Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In And 执行时间 = 原执行时间_In;
  -----取消执行有效天数限制
  Select Zl_To_Number(Nvl(zl_GetSysParameter(220), '999')) Into n_取消执行 From Dual;
  Select v_Date - d_登记时间 Into n_Diffday From Dual;
  --登记时间超过取消执行天数的记录，不允许修改医嘱执行情况
  If n_Diffday > n_取消执行 Then
    v_Error := '医嘱执行登记时间超过了取消执行有效天数，不能修改医嘱执行情况！';
    Raise Err_Custom;
  End If;

  If 本次数次_In = 1 Then
    --对医嘱开始时间进行检查 
    Select a.开始执行时间 Into d_开始时间 From 病人医嘱记录 A Where a.Id = 医嘱id_In;
    If Not d_开始时间 Is Null Then
      If 执行时间_In < d_开始时间 Then
        v_Error := '执行时间必须大于医嘱的开始执行时间''' || To_Char(d_开始时间, 'yyyy-mm-dd HH24:mi:ss') || '''！';
        Raise Err_Custom;
      End If;
    End If;
  End If;

  Select 执行部门id Into n_执行科室id From 病人医嘱发送 Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In;
  --病人医嘱执行
  Update 病人医嘱执行
  Set 要求时间 = 要求时间_In, 本次数次 = 本次数次_In, 执行摘要 = 执行摘要_In, 执行人 = 执行人_In, 执行时间 = 执行时间_In, 登记时间 = v_Date, 登记人 = v_人员姓名,
      执行结果 = 执行结果_In, 说明 = 未执行原因_In, 执行科室id = n_执行科室id
  Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In And 执行时间 = 原执行时间_In;
  --本次执行次数或这执行结果修改后需要更新单据的执行状态
  If v_执行结果old <> 执行结果_In Or n_本次数次old <> 本次数次_In Then
    Select 病人来源, Nvl(相关id, ID), 诊疗类别
    Into v_病人来源, v_组id, v_诊疗类别
    From 病人医嘱记录
    Where ID = 医嘱id_In;
  
    If v_病人来源 = 2 Then
      Select Decode(记录性质, 1, 1, Decode(门诊记帐, 1, 1, 2))
      Into v_费用性质
      From 病人医嘱发送
      Where 发送号 = 发送号_In And 医嘱id = 医嘱id_In;
    Else
      v_费用性质 := 1;
    End If;
  
    Select Decode(a.执行状态, 1, a.发送数次, c.登记次数), Decode(a.执行状态, 1, 0, a.发送数次 - c.登记次数), a.发送数次, c.登记次数



    Into n_执行次数, n_剩余次数, n_发送数次, n_登记数次
    From 病人医嘱发送 A,
         (Select 医嘱id_In 医嘱id, 发送号_In 发送号, Nvl(Sum(b.本次数次), 0) As 登记次数
           From 病人医嘱执行 B
           Where b.医嘱id = 医嘱id_In And b.发送号 = 发送号_In And Nvl(b.执行结果, 1) = 1) C
    Where a.医嘱id = c.医嘱id And a.发送号 = c.发送号 And a.医嘱id = 医嘱id_In And a.发送号 = 发送号_In;
  
    --如果全部执行则状态为1，未执行状态为0，部分执行状态为2
    Select Decode(n_剩余次数, 0, 1, Decode(n_执行次数, 0, 0, 2)) Into n_执行状态 From Dual;
  
    --更新医嘱执行计价.执行状态
    If n_发送数次 > 0 Then
      Select Count(Distinct 要求时间) Into v_Count From 医嘱执行计价 Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In;
      If v_Count > 0 Then
        n_单次数次 := n_发送数次 / v_Count;
        --已执行数量+本次数次 总共能够执行多少个时间点,取最大整数
        v_Count := Ceil((n_登记数次) / n_单次数次);
        If n_登记数次 = 0 Then
          Update 医嘱执行计价
          Set 执行状态 = 0
          Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In And Nvl(执行状态, 0) <> 2;
        Else
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
            Update 医嘱执行计价
            Set 执行状态 = 0
            Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In And 要求时间 > d_要求时间 And Nvl(执行状态, 0) <> 2;
          End If;
        End If;
      End If;
    End If;
  
    --执行次数不为0就标记为正在执行
    If Nvl(单独执行_In, 0) = 1 Then
      Update 病人医嘱发送
      Set 执行状态 = Decode(n_执行次数, 0, 0, 3), 完成人 = Null, 完成时间 = Null
      Where 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In;
    Else
      Update 病人医嘱发送
      Set 执行状态 = Decode(n_执行次数, 0, 0, 3), 完成人 = Null, 完成时间 = Null
      Where 执行状态 In (0, 3) And 发送号 + 0 = 发送号_In And
            医嘱id In (Select ID From 病人医嘱记录 Where (ID = v_组id Or 相关id = v_组id) And 诊疗类别 = v_诊疗类别);
    End If;
  
    If v_费用性质 = 2 Then
      If Nvl(单独执行_In, 0) = 1 Then
        Update 住院费用记录 A
        Set 执行状态 = n_执行状态, 执行人 = Decode(n_执行状态, 0, Null, 执行人_In), 执行时间 = Decode(n_执行状态, 0, d_执行时间, 执行时间_In)
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And Not Exists
         (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1) And a.记录状态 In (0, 1, 3) And
              (医嘱序号, NO, 记录性质) In
              (Select 医嘱id, NO, 记录性质
               From 病人医嘱发送
               Where 执行状态 = Decode(n_执行次数, 0, 0, 3) And 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In);
      Else
        Update 住院费用记录 A
        Set 执行状态 = n_执行状态, 执行人 = Decode(n_执行状态, 0, Null, 执行人_In), 执行时间 = Decode(n_执行状态, 0, d_执行时间, 执行时间_In)
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And Not Exists
         (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1) And a.记录状态 In (0, 1, 3) And
              (医嘱序号, NO, 记录性质) In
              (Select 医嘱id, NO, 记录性质
               From 病人医嘱发送
               Where 执行状态 = Decode(n_执行次数, 0, 0, 3) And 发送号 + 0 = 发送号_In And
                     医嘱id In
                     (Select ID From 病人医嘱记录 Where (ID = v_组id Or 相关id = v_组id) And 诊疗类别 = v_诊疗类别));
      End If;
    Else
      If Nvl(单独执行_In, 0) = 1 Then
        Update 门诊费用记录 A
        Set 执行状态 = n_执行状态, 执行人 = Decode(n_执行状态, 0, Null, 执行人_In), 执行时间 = Decode(n_执行状态, 0, d_执行时间, 执行时间_In)
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And Not Exists
         (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1) And a.记录状态 In (0, 1, 3) And
              (医嘱序号, NO, 记录性质) In
              (Select 医嘱id, NO, 记录性质
               From 病人医嘱发送
               Where 执行状态 = Decode(n_执行次数, 0, 0, 3) And 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In);
      Else
        Update 门诊费用记录 A
        Set 执行状态 = n_执行状态, 执行人 = Decode(n_执行状态, 0, Null, 执行人_In), 执行时间 = Decode(n_执行状态, 0, d_执行时间, 执行时间_In)
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And Not Exists
         (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1) And a.记录状态 In (0, 1, 3) And
              (医嘱序号, NO, 记录性质) In
              (Select 医嘱id, NO, 记录性质
               From 病人医嘱发送
               Where 执行状态 = Decode(n_执行次数, 0, 0, 3) And 发送号 + 0 = 发送号_In And
                     医嘱id In
                     (Select ID From 病人医嘱记录 Where (ID = v_组id Or 相关id = v_组id) And 诊疗类别 = v_诊疗类别));
      End If;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人医嘱执行_Update;
/





------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0024' Where 编号=&n_System;
Commit;
