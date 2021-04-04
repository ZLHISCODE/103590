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
--129954:蒋廷中,2018-09-07,新增参数控制医生工作站危急值弹窗提醒
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
  Select Zlparameters_Id.Nextval, &n_System, 1260, 1, 1, 0, 0, 0, 0, 39, '门诊危急值弹窗提醒', '1', '1', '控制门诊危急值提醒是否弹窗',
         '0-控制门诊危急值提醒不弹窗；1-控制门诊危急值弹窗提醒', '', '适用于用户想要控制门诊危急值弹窗提醒', Null
  From Dual;

Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
  Select Zlparameters_Id.Nextval, &n_System, 1261, 1, 1, 0, 0, 0, 0, 55, '住院危急值弹窗提醒', '1', '1', '控制住院危急值提醒是否弹窗',
         '0-控制住院危急值提醒不弹窗；1-控制住院危急值弹窗提醒', '', '适用于用户想要控制住院危急值弹窗提醒', Null
  From Dual;


--126863:焦博,2018-09-06,门诊记帐模块新增参数病人来源
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
  Select Zlparameters_Id.Nextval, &n_System, 1122, 1, 0, 0, 0, 0, 0, 84, '病人来源', Null, '1',
         ' 控制当前记帐是缺省按门诊病人或住院病人来查找病人信息和控制开单科室及执行科室的:' || Chr(13) || '1) 主要有两个地方设置参数:一是参数设置中,二是在记帐窗口的状态栏上.' || Chr(13) ||
          '2)在姓名通过输入"1. 病人ID,2.住院号,3. 就诊卡号,4.门诊号,5.医保号,6.身份证号,7.IC卡号"时,将会自动切换到该病人来的来源,比如:当前病人是在院病人,如果当前设置的是门诊病人,将会自动切换到住院病人状态,在收完此住院病人后,将自动切换到此参数设置的状态' ||
          Chr(13) || '3)根据病人来源来确定"开单科室"及具体收费项目的执行科室,即病人来源为门诊的,则开单科室或执行科室只能是服务于门诊或即能服务于门诊和住院的科室',
         '' || Chr(13) || '1-门诊病人,2-住院病人', Null, '适用于需要在门诊记帐时根据病人来源来进行控制的用户.', Null
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
--126863:焦博,2018-09-06,门诊记帐模块新增参数病人来源
Create Or Replace Procedure Zl_门诊记帐记录_Insert
(
  No_In         门诊费用记录.No%Type,
  序号_In       门诊费用记录.序号%Type,
  病人id_In     门诊费用记录.病人id%Type,
  标识号_In     门诊费用记录.标识号%Type,
  姓名_In       门诊费用记录.姓名%Type,
  性别_In       门诊费用记录.性别%Type,
  年龄_In       门诊费用记录.年龄%Type,
  费别_In       门诊费用记录.费别%Type,
  加班标志_In   门诊费用记录.加班标志%Type,
  婴儿费_In     门诊费用记录.婴儿费%Type,
  病人科室id_In 门诊费用记录.病人科室id%Type,
  开单部门id_In 门诊费用记录.开单部门id%Type,
  开单人_In     门诊费用记录.开单人%Type,
  从属父号_In   门诊费用记录.从属父号%Type,
  收费细目id_In 门诊费用记录.收费细目id%Type,
  收费类别_In   门诊费用记录.收费类别%Type,
  计算单位_In   门诊费用记录.计算单位%Type,
  付数_In       门诊费用记录.付数%Type,
  数次_In       门诊费用记录.数次%Type,
  附加标志_In   门诊费用记录.附加标志%Type,
  执行部门id_In 门诊费用记录.执行部门id%Type,
  价格父号_In   门诊费用记录.价格父号%Type,
  收入项目id_In 门诊费用记录.收入项目id%Type,
  收据费目_In   门诊费用记录.收据费目%Type,
  标准单价_In   门诊费用记录.标准单价%Type,
  应收金额_In   门诊费用记录.应收金额%Type,
  实收金额_In   门诊费用记录.实收金额%Type,
  发生时间_In   门诊费用记录.发生时间%Type,
  登记时间_In   门诊费用记录.登记时间%Type,
  药品摘要_In   药品收发记录.摘要%Type,
  划价_In       Number,
  操作员编号_In 门诊费用记录.操作员编号%Type,
  操作员姓名_In 门诊费用记录.操作员姓名%Type,
  记帐单id_In   门诊费用记录.记帐单id%Type := Null,
  费用摘要_In   门诊费用记录.摘要%Type := Null,
  医嘱序号_In   门诊费用记录.医嘱序号%Type := Null,
  频次_In       药品收发记录.频次%Type := Null,
  单量_In       药品收发记录.单量%Type := Null,
  用法_In       药品收发记录.用法%Type := Null, --用法[|煎法]
  期效_In       药品收发记录.扣率%Type := Null,
  计价特性_In   药品收发记录.扣率%Type := Null,
  门诊标志_In   门诊费用记录.门诊标志%Type := 1,
  中药形态_In   门诊费用记录.结论%Type := Null,
  备货材料_In   Number := 0,
  批次_In       药品收发记录.批次%Type := Null
) As
  --功能：新收一张门诊记帐单据
  --参数：
  --   药品摘要_IN:修改保存新单据时用。目前仅用于存放于药品收发记录的摘要中。
  --         原单据(记录状态=2)记录修改产生的新单据号。
  --         新单据(记录状态=1)记录所修改的原单据号。
  v_费用id 门诊费用记录.Id%Type;
  n_急诊   病人挂号记录.急诊%Type;

  --临时变量
  v_用法     药品收发记录.用法%Type;
  v_煎法     药品收发记录.外观%Type;
  n_单价小数 Number;
  n_挂号id   病人挂号记录.Id%Type;
  n_留观次数 病案主页.主页id%Type;

  n_Dec     Number;
  n_Count   Number;
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_发药窗口 药品收发记录.发药窗口%Type;
  n_跟踪在用 材料特性.跟踪在用%Type;

Begin
  n_跟踪在用 := 0;
  If 收费类别_In = '4' Then
    --跟踪在用的卫材才处理
    Select Nvl(跟踪在用, 0) Into n_跟踪在用 From 材料特性 Where 材料id = 收费细目id_In;
  End If;

  --金额小数位数
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')), Zl_To_Number(Nvl(zl_GetSysParameter(157), '5'))
  Into n_Dec, n_单价小数
  From Dual;

  Select Max(主页id) Into n_留观次数 From 病案主页 Where 病人id = 病人id_In And 病人性质 = 1 And 出院日期 Is Null;

  If (收费类别_In In ('5', '6', '7') Or 收费类别_In = '4' And n_跟踪在用 = 1) And Nvl(划价_In, 0) = 0 Then
    --同一张单据,满足同一药房同一窗口
    Begin
      Select 发药窗口
      Into v_发药窗口
      From 门诊费用记录
      Where 收费类别 In ('5', '6', '7', '4') And NO = No_In And 记录性质 = 2 And 执行部门id = 执行部门id_In And 发药窗口 Is Not Null And
            Rownum <= 1;
    Exception
      When Others Then
        v_发药窗口 := Null;
    End;
    If v_发药窗口 Is Null Then
      --同一病人在普通号挂号有效挂号天数内且未发药的且上班的,以最近一次记账窗口为准
      n_Count := To_Number(Substr(Nvl(zl_GetSysParameter(21), '11') || '11', 1, 1));
      If n_Count = 0 Then
        n_Count := 1;
      End If;
    
      Begin
        Select 发药窗口
        Into v_发药窗口
        From (Select 登记时间, 发药窗口
               From 门诊费用记录 A
               Where 收费类别 In ('5', '6', '7', '4') And 病人id = 病人id_In And 登记时间 Between Sysdate - n_Count And Sysdate And
                     记录性质 = 2 And 执行部门id = 执行部门id_In And 发药窗口 Is Not Null And Exists
                (Select 1
                      From 未发药品记录
                      Where a.No = NO And 单据 In (9, 26) And 库房id + 0 = 执行部门id_In And 病人id + 0 = 病人id_In) And Exists
                (Select 1
                      From 发药窗口
                      Where Nvl(上班否, 0) = 1 And 名称 = a.发药窗口 And Nvl(专家, 0) = 0 And 药房id = 执行部门id_In)
               Order By 登记时间 Desc)
        Where Rownum <= 1;
      
      Exception
        When Others Then
          v_发药窗口 := Null;
      End;
      If v_发药窗口 Is Null Then
        v_发药窗口 := Zl_Get发药窗口(执行部门id_In);
      End If;
    End If;
  End If;
  --门诊费用记录
  Select 病人费用记录_Id.Nextval Into v_费用id From Dual;

  --是否是急诊挂号单
  If Nvl(医嘱序号_In, 0) <> 0 Then
    Begin
      Select Nvl(Max(急诊), 0), Max(ID)
      Into n_急诊, n_挂号id
      From 病人挂号记录
      Where NO In (Select 挂号单 From 病人医嘱记录 Where ID = Nvl(医嘱序号_In, 0)) And 病人id = 病人id_In;
    Exception
      When Others Then
        n_急诊   := Null;
        n_挂号id := Null;
    End;
  End If;

  Insert Into 门诊费用记录
    (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 门诊标志, 病人id, 标识号, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 数次, 加班标志,
     附加标志, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 记帐费用, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 操作员编号, 操作员姓名, 婴儿费, 记帐单id,
     摘要, 医嘱序号, 结论, 发药窗口, 是否急诊, 主页id, 挂号id)
  Values
    (v_费用id, 2, No_In, Decode(划价_In, 1, 0, 1), 序号_In, Decode(从属父号_In, 0, Null, 从属父号_In),
     Decode(价格父号_In, 0, Null, 价格父号_In), 门诊标志_In, 病人id_In, Decode(标识号_In, 0, Null, 标识号_In), 姓名_In, 性别_In, 年龄_In,
     病人科室id_In, 费别_In, 收费类别_In, 收费细目id_In, 计算单位_In, 付数_In, 数次_In, 加班标志_In, 附加标志_In, 收入项目id_In, 收据费目_In, 标准单价_In, 应收金额_In,
     实收金额_In, 1, 操作员姓名_In, 开单部门id_In, 开单人_In, 发生时间_In, 登记时间_In, 执行部门id_In, 0, Decode(划价_In, 1, Null, 操作员编号_In),
     Decode(划价_In, 1, Null, 操作员姓名_In), 婴儿费_In, 记帐单id_In, 费用摘要_In, 医嘱序号_In, 中药形态_In, v_发药窗口, Nvl(n_急诊, 0), n_留观次数, n_挂号id);

  --相关汇总表的处理
  If Nvl(划价_In, 0) = 0 Then
    --病人余额
    If Nvl(门诊标志_In, 0) <> 4 Then
      Update 病人余额
      Set 费用余额 = Nvl(费用余额, 0) + 实收金额_In
      Where 病人id = 病人id_In And 性质 = 1 And 类型 = Decode(门诊标志_In, 2, 2, 1);
    
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 性质, 类型, 费用余额, 预交余额)
        Values
          (病人id_In, 1, Decode(门诊标志_In, 2, 2, 1), 实收金额_In, 0);
      End If;
    End If;
  
    --病人未结费用
    Update 病人未结费用
    Set 金额 = Nvl(金额, 0) + 实收金额_In
    Where 病人id = 病人id_In And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(病人科室id, 0) = Nvl(病人科室id_In, 0) And
          Nvl(开单部门id, 0) = Nvl(开单部门id_In, 0) And Nvl(执行部门id, 0) = Nvl(执行部门id_In, 0) And 收入项目id + 0 = 收入项目id_In And
          来源途径 + 0 = 门诊标志_In;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人未结费用
        (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
      Values
        (病人id_In, Null, Null, 病人科室id_In, 开单部门id_In, 执行部门id_In, 收入项目id_In, 门诊标志_In, 实收金额_In);
    End If;
  
  End If;

  --药品和卫生材料部分
  If 收费类别_In In ('4', '5', '6', '7') Then
    --药品用法煎法分解
    If 用法_In Is Not Null Then
      If Instr(用法_In, '|') > 0 Then
        v_用法 := Substr(用法_In, 1, Instr(用法_In, '|') - 1);
        v_煎法 := Substr(用法_In, Instr(用法_In, '|') + 1);
      Else
        v_用法 := 用法_In;
      End If;
    End If;
    Zl_药品收发记录_销售出库(v_费用id, 药品摘要_In, 频次_In, 单量_In, v_用法, v_煎法, 期效_In, 计价特性_In, Null, 备货材料_In, 批次_In);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊记帐记录_Insert;
/

--130471:余伟节,2018-09-04,病人预约

Create Or Replace Procedure Zl_Third_Outpatireg
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --功能：用于产生预入院记录/取消预入院    数据写入
  --入参：xml_in
  --<IN>
  -- <TYPE>1</TYPE>   --操作类型：1-产生预入院记录；0-取消预入院
  -- <GHID>1162695</GHID>       --挂号id
  -- <RYKSID>202704</RYKSID>    --入院科室ID
  -- <RYBQID>202704</RYBQID>    --入院病区ID
  -- <CH>5</CH>   --床号
  -- <YZID>3</YZID> --医嘱id
  -- <CZYBH></CZYBH> --操作员编号
  -- <CZYXM></CZYXM> --操作员姓名
  --</IN>

  --出参：Xml_Out
  --<OUTPUT>
  --   <RESULT>true</RESULT>
  --</OUTPUT>

  --失败：
  --<OUTPUT>
  --   <RESULT>false</RESULT>
  --   <ERROR>
  --     <MSG>详细错误提示</MSG>
  --   </ERROR>
  --</OUTPUT>

  n_医嘱id 病人医嘱记录.Id%Type;
  Cursor c_Advice Is
    Select Nvl(a.相关id, a.Id) As 组id, a.相关id, a.序号, a.病人id, a.挂号单, a.婴儿, a.姓名, c.操作类型, a.诊疗类别, a.医嘱状态, a.医嘱内容, a.开嘱医生,
           a.开始执行时间, a.执行时间方案, a.频率次数, a.频率间隔, a.间隔单位, Nvl(a.紧急标志, 0) As 紧急标志, a.诊疗项目id, a.收费细目id
    From 病人医嘱记录 A, 诊疗项目目录 C
    Where a.诊疗项目id = c.Id And a.诊疗类别 = 'Z' And c.操作类型 = '2' And a.Id = n_医嘱id;
  r_Advice c_Advice%RowType;

  Cursor c_Pati(v_病人id 病人信息.病人id%Type) Is
    Select a.病人id, a.住院号, a.姓名, a.性别, a.年龄, a.费别, a.出生日期, a.国籍, a.民族, a.学历, a.婚姻状况, a.职业, a.身份, a.身份证号, a.出生地点, a.家庭地址,
           a.家庭地址邮编, a.家庭电话, a.户口地址, a.户口地址邮编, a.联系人姓名, a.联系人关系, a.联系人地址, a.联系人电话, a.工作单位, a.合同单位id, a.单位电话, a.单位邮编,
           a.单位开户行, a.单位帐号, a.担保人, a.担保额, a.担保性质, a.籍贯, a.区域, a.医疗付款方式, a.险类
    From 病人信息 A
    Where a.病人id = v_病人id;
  r_Pati c_Pati%RowType;

  n_Type   Number;
  n_挂号id 病人医嘱记录.Id%Type;
  n_科室id 病人医嘱记录.Id%Type;
  n_病区id 病人医嘱记录.Id%Type;
  v_床号   病案主页.入院病床%Type;

  n_病人id 病案主页.病人id%Type;
  v_No     病人挂号记录.No%Type;
  n_Count  Number;

  v_入院方式 病案主页.入院方式%Type;
  v_人员编号 人员表.编号%Type;
  v_人员姓名 人员表.姓名%Type;
  v_Temp     Varchar2(4000);
  v_Error    Varchar2(2000);

  Err_Custom Exception;
Begin
  Select Extractvalue(Value(A), 'IN/TYPE'), Extractvalue(Value(A), 'IN/GHID') As 挂号id,
         Extractvalue(Value(A), 'IN/RYKSID') As 入院科室id, Extractvalue(Value(A), 'IN/RYBQID') As 入院病区id,
         Extractvalue(Value(A), 'IN/CH') As 床号, Extractvalue(Value(A), 'IN/CZYBH') As 编号,
         Extractvalue(Value(A), 'IN/CZYXM') As 姓名, Extractvalue(Value(A), 'IN/YZID') As 医嘱id
  Into n_Type, n_挂号id, n_科室id, n_病区id, v_床号, v_人员编号, v_人员姓名, n_医嘱id
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If n_Type = 1 Then
    --住院预约登记
    Select a.病人id, a.No, Decode(a.急诊, 1, '急诊', Null)
    Into n_病人id, v_No, v_入院方式
    From 病人挂号记录 A
    Where a.Id = n_挂号id;
  
    Open c_Advice;
    Fetch c_Advice
      Into r_Advice;
  
    If r_Advice.紧急标志 = 1 Then
      v_入院方式 := '急诊';
    End If;
  
    Open c_Pati(n_病人id);
    Fetch c_Pati
      Into r_Pati;
  
    --当前操作人员
    If v_人员编号 Is Null Or v_人员姓名 Is Null Then
      v_Temp     := Zl_Identity;
      v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
      v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
      v_人员编号 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
      v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    End If;
  
    --删除留观记录和住院预约记录不能并存
    Begin
      Select Count(1) Into n_Count From 病案主页 Where 病人id = r_Advice.病人id And Nvl(主页id, 0) = 0;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If Nvl(n_Count, 0) > 0 Then
      Zl_入院病案主页_Delete(r_Advice.病人id, 0, 0, 0);
      n_Count := 0;
    End If;
  
    If n_Count = 0 Then
      Select Count(1) Into n_Count From 病案主页 Where 病人id = r_Advice.病人id And 出院日期 Is Null;
    End If;
    If n_Count = 0 Then
      Select Count(1)
      Into n_Count
      From 病案主页
      Where 病人id = r_Advice.病人id And (入院日期 >= r_Advice.开始执行时间 Or 出院日期 >= r_Advice.开始执行时间);
    End If;
  
    If n_Count = 0 Then
      Zl_入院病案主页_Insert(1, 0, r_Pati.病人id, r_Pati.住院号, Null, r_Pati.姓名, r_Pati.性别, r_Pati.年龄, r_Pati.费别, r_Pati.出生日期,
                       r_Pati.国籍, r_Pati.民族, r_Pati.学历, r_Pati.婚姻状况, r_Pati.职业, r_Pati.身份, r_Pati.身份证号, r_Pati.出生地点,
                       r_Pati.家庭地址, r_Pati.家庭地址邮编, r_Pati.家庭电话, r_Pati.户口地址, r_Pati.户口地址邮编, r_Pati.联系人姓名, r_Pati.联系人关系,
                       r_Pati.联系人地址, r_Pati.联系人电话, r_Pati.工作单位, r_Pati.合同单位id, r_Pati.单位电话, r_Pati.单位邮编, r_Pati.单位开户行,
                       r_Pati.单位帐号, r_Pati.担保人, r_Pati.担保额, r_Pati.担保性质, n_科室id, Null, Null, v_入院方式, Null, Null,
                       r_Advice.开嘱医生, r_Pati.籍贯, r_Pati.区域, r_Advice.开始执行时间, Null, Null, r_Pati.医疗付款方式, Null, Null, Null,
                       Null, Null, Null, r_Pati.险类, v_人员编号, v_人员姓名, 0, Null, n_病区id, 0, Null, Null, Null, Null, Null,
                       Null, Null, n_挂号id);
    End If;
  Else
    --取消登记
    Select Count(1) Into n_Count From 病案主页 B Where b.挂号id = n_挂号id;
    If n_Count > 0 Then
      Select b.病人id Into n_病人id From 病案主页 B Where b.挂号id = n_挂号id;
      Zl_入院病案主页_Delete(n_病人id, 0);
    End If;
  End If;
  Xml_Out := Xmltype('<OUTPUT><RESULT>true</RESULT></OUTPUT>');
Exception
  When Err_Custom Then
    Xml_Out := Xmltype('<OUTPUT><RESULT>false</RESULT><ERROR><MSG>' || v_Error || '</MSG></ERROR></OUTPUT>');
  When Others Then
    Xml_Out := Xmltype('<OUTPUT><RESULT>false</RESULT><ERROR><MSG>' || SQLCode || '***' || SQLErrM ||
                       '</MSG></ERROR></OUTPUT>');
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Outpatireg;
/



------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0028' Where 编号=&n_System;
Commit;