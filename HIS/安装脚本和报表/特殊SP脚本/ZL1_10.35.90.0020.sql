----------------------------------------------------------------------------------------------------------------
--本脚本支持从ZLHIS+ v10.35.90升级到 v10.35.90
--请以数据空间的所有者登录PLSQL并执行下列脚本
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------



------------------------------------------------------------------------------
--结构修正部份
------------------------------------------------------------------------------
--126898:秦龙,2018-07-09,增加属性“批准文号”
Alter Table 药品计划内容 Add 批准文号 varchar2(40);




------------------------------------------------------------------------------
--数据修正部份
------------------------------------------------------------------------------
--127075:焦博,2018-07-12,修改分诊模块参数分诊台签到排队
Update zlParameters
Set 部门 = 1, 关联说明 = '1.启用“排队叫号模式”本参数才有效;' || Chr(13) || '2.启用分诊台签道排队的科室时，如果后期新增号源的科室，默认为分诊台签道排队',
    警告说明 = '调整本参数，将会影响排队队列的产生时机。',
    影响控制说明 = '1. 启用了此参数后,则在分诊台将启用"签到"功能,在病人签到后,才进入排队队列,否则挂号后就进行队列.' || Chr(13) ||
              '2. 也可在分诊台签道排队的模式下，设置哪些科室允许进行签道排队,分四种情况;' || Chr(13) || '1）如果不设置启用的科室，则为所有科室分诊台签道排队;' || Chr(13) ||
              '2)如果设置了启用的科室,针对未启用的科室，则不按分诊台签道排队的规则处理,即挂号或预约相关参数进行排队;' || Chr(13) ||
              '3)如果设置了启用的科室,针对后期新增号源的科室，则按分诊台签道排队规则处理;' || Chr(13) || '4）针对启用的科室，则只能在分诊台签道时才能排队。'
Where 系统 = &n_System And 模块 = 1113 And 参数名 = '分诊台签到排队';

--127487:刘涛,2018-07-11,分批卫材入库是否检查批号产地
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
  Select Zlparameters_Id.Nextval, &n_System, Null, 0, 0, 0, 0, 0, 0, 305, '分批卫材批号产地控制',  '1', '1',
         '启用该参数后,在涉及到卫材入库的地方检查分批卫材是否录入了批号和产地', '0-不检查分批卫材是否录入了批号和产地；1-检查分批卫材是否录入了批号和产地。', Null,
         '适用于用户想要检查分批卫材入库时是否录入了批号和产地', Null
  From Dual;

--119905:冉俊明,2018-07-09,病人收费管理按病人补打票据含挂号费
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
  Select Zlparameters_Id.Nextval, &n_System, 1121, 1, 0, 0, 0, 0, 0, 115, '按病人补打票据含挂号费', Null, '0', '按病人补打票据时，是否缺省提取挂号费',
         '0-不缺省提取挂号费，1-缺省提取挂号费', Null, '用于个性化设置', Null
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
--125421:刘鹏飞,2018-07-12,作废时支持外来病人费用回退
Create Or Replace Procedure Zl_病人医嘱记录_作废
(
  Id_In         In 病人医嘱记录.Id%Type,
  操作员编号_In In 人员表.编号%Type := Null,
  操作员姓名_In In 人员表.姓名%Type := Null,
  护理医嘱id_In In 病人医嘱记录.Id%Type := Null,
  作废时间_In   In 病人医嘱状态.操作时间%Type := Null
) Is
  --功能：作废指定的医嘱(未发送的长嘱或临嘱)
  --说明：一并给药的只能调用一次(界面显示有多行)
  --参数：ID_IN=组医嘱ID
  --      护理医嘱id_In 取除开本次作废的护理等级医嘱外的最近的自动停止的护理等级医嘱id
  v_发送号       病人医嘱发送.发送号%Type;
  v_费用no       门诊费用记录.No%Type;
  v_记录性质     门诊费用记录.记录性质%Type;
  v_费用序号     Varchar2(255);
  n_自动取消执行 Number(1) := 0;
  n_先作废后退药 Number(1) := 0;

  v_Date     Date;
  v_Count    Number;
  v_Temp     Varchar2(255);
  v_人员编号 人员表.编号%Type;
  v_人员姓名 人员表.姓名%Type;
  v_No       病人医嘱发送.No%Type;

  --包含医嘱相关信息
  Cursor c_Advice Is
    Select Nvl(a.相关id, a.Id) As 组id, a.病人id, a.挂号单, a.主页id, a.婴儿, a.医嘱状态, a.上次执行时间, a.医嘱内容, a.诊疗类别, b.操作类型, a.病人来源,
           a.执行科室id, b.执行频率, a.诊疗项目id, a.开始执行时间
    From 病人医嘱记录 a, 诊疗项目目录 b
    Where a.诊疗项目id = b.Id And a.Id = Id_In;
  r_Advice c_Advice%Rowtype;

  --门诊医嘱作废时，取对应的费用销帐或作废(收费划价单)：
  --根据医嘱及发送NO求出本次回退要销帐或退费的记录
  --一组医嘱并不是都填写了发送记录,也不一定都计费了,且可能NO不同
  --只管记录状态为1的记录,如果已经销帐或部份销帐的记录,不再处理
  --费用只求价格父号为空的,以便取序号销帐
  --如果【门诊药嘱先作废后退药】,则不对相应费用(包括给药途径的)进行检查和处理,除非是还没有执行的记帐单,或未执行、收费的划价单，可以先删了。

  Cursor c_Rollmoney(v_发送号 病人医嘱发送.发送号%Type) Is
    Select Decode(a.记录性质, 11, 1, a.记录性质) As 记录性质, a.记录状态, a.No, a.序号, a.执行状态 As 费用执行, c.执行状态 As 医嘱执行, c.执行部门id, b.病人科室id,
           b.诊疗类别, i.操作类型
    From 门诊费用记录 a, 病人医嘱记录 b, 病人医嘱发送 c, 诊疗项目目录 i
    Where c.医嘱id = b.Id And c.发送号 = v_发送号 And (b.Id = Id_In Or b.相关id = Id_In) And a.医嘱序号 = b.Id And a.记录状态 In (0, 1) And
          a.No = c.No And (a.记录性质 = c.记录性质 Or a.记录性质 = 11 And c.记录性质 = 1) And b.诊疗项目id = i.Id And a.价格父号 Is Null And
          (n_先作废后退药 = 0 Or
          n_先作废后退药 = 1 And
          Not (Exists (Select 1
                        From 门诊费用记录 d
                        Where d.医嘱序号 = b.Id And d.记录状态 In (0, 1) And d.No = c.No And
                              (d.记录性质 = c.记录性质 Or d.记录性质 = 11 And c.记录性质 = 1) And d.收费类别 In ('5', '6', '7'))) Or
          Nvl(a.执行状态, 0) = 0 And Not (a.记录性质 = 1 And a.记录状态 <> 0))
    Order By a.记录性质, a.No, a.序号, a.收费细目id;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --检查医嘱状态是否正确:并发操作
  Open c_Advice;
  Fetch c_Advice
    Into r_Advice;

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

  --检查是否已经出了报告单，已经出报告单的医嘱不能够作废
  Select Count(1) Into v_Count From 病人医嘱报告 Where 医嘱id = Id_In;
  If v_Count > 0 Then
    If Not (r_Advice.操作类型 = '7' And r_Advice.诊疗类别 = 'Z') Then
      v_Error := '医嘱"' || r_Advice.医嘱内容 || '"已出报告，不能作废。';
      Raise Err_Custom;
    End If;
  End If;

  --检查是否是输液配液记录，并是否已经锁定
  Select Count(1) Into v_Count From 输液配药记录 Where 是否锁定 = 1 And 医嘱id = Id_In;
  If v_Count > 0 Then
    v_Error := '医嘱"' || r_Advice.医嘱内容 || '"是输液药品，已经被输液配置中心锁定，不能作废。';
    Raise Err_Custom;
  End If;

  If r_Advice.挂号单 Is Null And r_Advice.病人来源 <> 3 Then
    If r_Advice.医嘱状态 In (4, 8, 9) Then
      v_Error := '医嘱"' || r_Advice.医嘱内容 || '"已经被作废或停止，不能再作废。';
      Raise Err_Custom;
    Elsif r_Advice.上次执行时间 Is Not Null Then
      v_Error := '医嘱"' || r_Advice.医嘱内容 || '"已经发送，不能被作废。';
      Raise Err_Custom;
    End If;
  
    --持续性护理等级无须发送，校对后就可能已自动计费，作废及回退作废都应按停止流程处理。
    If r_Advice.诊疗类别 = 'H' And r_Advice.操作类型 = '1' And r_Advice.执行频率 = '2' And Nvl(r_Advice.婴儿, 0) = 0 Then
      --(已取消，由于存在无费退院的情况，问题号：45977)a.开始时间是当天之前的，说明已生效（自动费用计算），不允许作废。
      --医嘱的时间只精确到了分钟，所以变动记录的开始时间要去掉秒来比较。
      v_Count := 0;
      Begin
        Select b.终止时间
        Into v_Date
        From 病人变动记录 b, 病人医嘱计价 c
        Where b.病人id = r_Advice.病人id And b.主页id = r_Advice.主页id And c.医嘱id = Id_In And c.收费细目id = b.护理等级id And
              b.开始原因 = 6 And b.附加床位 = 0 And
              To_Char(b.开始时间, 'yyyy-mm-dd hh24:mi') = To_Char(r_Advice.开始执行时间, 'yyyy-mm-dd hh24:mi');
      Exception
        When Others Then
          v_Count := 1;
      End;
      If v_Count = 0 Then
        --d.后续有其他变动发生
        If v_Date Is Not Null Then
          v_Error := '由于护理等级医嘱生效后已经产生了其他变动记录,不能作废该医嘱。';
          Raise Err_Custom;
        Else
          --本次有要自动启用的护理等级，如果和原来护理等级相同则不用撤消护理变动记录
          If Nvl(护理医嘱id_In, 0) <> 0 Then
            Delete 病人医嘱状态 Where 医嘱id = 护理医嘱id_In And 操作类型 In (8, 9);
            Select 操作类型
            Into v_Count
            From (Select 操作类型 From 病人医嘱状态 Where 医嘱id = 护理医嘱id_In Order By 操作时间 Desc)
            Where Rownum < 2;
            Update 病人医嘱记录
            Set 医嘱状态 = v_Count, 执行终止时间 = Null, 停嘱医生 = Null, 停嘱时间 = Null, 确认停嘱时间 = Null, 确认停嘱护士 = Null
            Where Id = 护理医嘱id_In;
            --排除过于频繁的操作
            Select Count(a.Id)
            Into v_Count
            From 病人医嘱记录 a, 诊疗收费关系 b, 病案主页 c
            Where a.诊疗项目id = b.诊疗项目id And c.护理等级id = b.收费项目id And c.病人id = a.病人id And c.主页id = a.主页id And
                  a.Id = 护理医嘱id_In;
          End If;
          If v_Count = 0 Then
            --c.护理等级是最后一条变动
            Zl_病人变动记录_Undo(r_Advice.病人id, r_Advice.主页id, v_人员编号, v_人员姓名, '1', Null, Null, '护理等级变动');
          End If;
        End If;
      Else
        --恢复最近一次被自动停止的护理等级
        If Nvl(护理医嘱id_In, 0) <> 0 Then
          Delete 病人医嘱状态 Where 医嘱id = 护理医嘱id_In And 操作类型 In (8, 9);
          Select 操作类型
          Into v_Count
          From (Select 操作类型 From 病人医嘱状态 Where 医嘱id = 护理医嘱id_In Order By 操作时间 Desc)
          Where Rownum < 2;
          Update 病人医嘱记录
          Set 医嘱状态 = v_Count, 执行终止时间 = Null, 停嘱医生 = Null, 停嘱时间 = Null, 确认停嘱时间 = Null, 确认停嘱护士 = Null
          Where Id = 护理医嘱id_In;
        Else
          --病人入院时指定的护理级产生的变动记录和医嘱新开产生的变动记录不同，这里要先判断
          Select Count(a.Id)
          Into v_Count
          From 病人变动记录 a
          Where a.病人id = r_Advice.病人id And a.主页id = r_Advice.主页id And a.开始原因 = 6;
          If v_Count <> 0 Then
            --b.如果与以前的护理等级相同，则校对时没有产生护理等级变动,产生护理等级停止变动
            Zl_病人变动记录_Nurse(r_Advice.病人id, r_Advice.主页id, Null, Sysdate, v_人员编号, v_人员姓名);
          End If;
        End If;
      End If;
    End If;
  Else
    If r_Advice.医嘱状态 <> 8 Then
      v_Error := '医嘱"' || r_Advice.医嘱内容 || '"尚未发送或已经作废。';
      Raise Err_Custom;
    End If;
    --医嘱附费判断
    Select Count(1)
    Into v_Count
    From 病人医嘱附费 a, 病人医嘱记录 b
    Where a.医嘱id = b.Id And (b.Id = Id_In Or b.相关id = Id_In);
    If v_Count <> 0 Then
      v_Error := '医嘱"' || r_Advice.医嘱内容 || '"存在附加费用，不能作废。';
      Raise Err_Custom;
    End If;
  
    Begin
      --医嘱ID为传入值的这条医嘱不一定发送了的,甚至无发送。
      Select Distinct 发送号
      Into v_发送号
      From 病人医嘱发送
      Where 医嘱id In (Select Id From 病人医嘱记录 Where Id = Id_In Or 相关id = Id_In);
    Exception
      When Others Then
        v_发送号 := Null;
    End;
  
    Select Zl_To_Number(Nvl(Zl_Getsysparameter(68), 0)) Into n_先作废后退药 From Dual;
    Select Zl_To_Number(Nvl(Zl_Getsysparameter('门诊本科自动执行', '1252'), 0)) Into n_自动取消执行 From Dual;
    If n_自动取消执行 = 1 And v_发送号 Is Not Null Then
      --先更新医嘱和费用的执行状态，因为后续的判断，以及过程Zl_门诊记帐记录_Delete中有检查
      For Rc In (Select a.医嘱id, a.执行部门id
                 From 病人医嘱发送 a, 病人医嘱记录 b
                 Where a.医嘱id = b.Id And (b.Id = Id_In Or b.相关id = Id_In) And a.执行部门id = b.病人科室id) Loop
        Zl_病人医嘱执行_Cancel(Rc.医嘱id, v_发送号, Null, 1, Rc.执行部门id);
      End Loop;
    End If;
  
    --门诊医嘱只可能发送一次
    --后面退费时还有检查，因为可能医嘱没有费用，所以要检查一次执行状态
    Select Count(*)
    Into v_Count
    From 病人医嘱发送 a, 病人医嘱记录 b, 诊疗项目目录 i
    Where a.医嘱id = b.Id And b.诊疗项目id = i.Id And a.执行状态 In (1, 3) And (b.Id = Id_In Or b.相关id = Id_In) And
          (n_先作废后退药 = 0 Or
          n_先作废后退药 = 1 And Not (b.诊疗类别 In ('5', '6', '7') Or b.诊疗类别 = 'E' And i.操作类型 In ('2', '3', '4')));
    If v_Count > 0 Then
      v_Error := '该医嘱已经执行或正在执行，不能作废。';
      Raise Err_Custom;
    End If;
  End If;

  If 作废时间_In Is Null Then
    Select Sysdate Into v_Date From Dual;
  Else
    v_Date := 作废时间_In;
  End If;

  Update 病人医嘱记录 Set 医嘱状态 = 4 Where Id = Id_In Or 相关id = Id_In;

  Insert Into 病人医嘱状态
    (医嘱id, 操作类型, 操作人员, 操作时间)
    Select Id, 4, v_人员姓名, v_Date From 病人医嘱记录 Where Id = Id_In Or 相关id = Id_In;

  --住院医嘱作废时,未打印的情况下,缺省设置为屏蔽打印
  If r_Advice.挂号单 Is Null And r_Advice.病人来源 <> 3 Then
    Select Count(*)
    Into v_Count
    From 病人医嘱打印
    Where 医嘱id In (Select Id From 病人医嘱记录 Where Id = Id_In Or 相关id = Id_In);
    If Nvl(v_Count, 0) = 0 Then
      Zl_病人医嘱记录_屏蔽打印(Id_In, 1);
    End If;
    If Nvl(r_Advice.婴儿, 0) > 0 And r_Advice.操作类型 = '11' Then
      Update 病人新生儿记录
      Set 死亡时间 = Null
      Where 病人id = r_Advice.病人id And Nvl(主页id, 0) = Nvl(r_Advice.主页id, 0) And 序号 = Nvl(r_Advice.婴儿, 0);
    End If;
  Else
    --门诊医嘱(临嘱)作废时还需要回退相关内容:只有一次发送
    --回退划价或记帐费用
    If v_发送号 Is Not Null Then
      --将该组医嘱的费用删除或销帐(按一组医嘱可能有不同NO处理)
      --门诊记帐：如果原始费用已被销帐(或部分销帐),调用过程中有判断
      --门诊划价：如果已收费，则不允许删除
      v_费用no   := Null;
      v_费用序号 := Null;
      For r_Rollmoney In c_Rollmoney(v_发送号) Loop
        If Nvl(r_Rollmoney.医嘱执行, 0) In (1, 3) Then
          --1-完全执行;3-正在执行
          v_Error := '医嘱"' || r_Advice.医嘱内容 || '"已经执行或正在执行，不能作废。';
          Raise Err_Custom;
        End If;
        If Nvl(r_Rollmoney.费用执行, 0) In (1, 2) Then
          --1-完全执行;2-部份执行
          v_Error := '医嘱费用单据"' || r_Rollmoney.No || '"中的内容已经全部或部分执行，不能作废。';
          Raise Err_Custom;
        End If;
        If r_Rollmoney.费用执行 = 9 Then
          v_Error := '医嘱费用单据"' || r_Rollmoney.No || '"中的收费结算产生异常，不能作废。';
          Raise Err_Custom;
        End If;
        v_Count := 1;
        If r_Rollmoney.记录性质 = 1 And r_Rollmoney.记录状态 <> 0 Then
          If 1 = n_先作废后退药 And r_Rollmoney.诊疗类别 = 'E' And r_Rollmoney.操作类型 In ('2', '3', '4') Then
            v_Count := 0;
          Else
            v_Error := '医嘱费用单据"' || r_Rollmoney.No || '"已经收费，不能作废。';
            Raise Err_Custom;
          End If;
        End If;
        If 1 = v_Count Then
          If Nvl(v_费用no, '空') <> r_Rollmoney.No Then
            If v_费用序号 Is Not Null And v_费用no Is Not Null Then
              v_费用序号 := Substr(v_费用序号, 2);
              If v_记录性质 = 1 Then
                Zl_门诊划价记录_Delete(v_费用no, v_费用序号);
              Elsif v_记录性质 = 2 Then
                Zl_门诊记帐记录_Delete(v_费用no, v_费用序号, v_人员编号, v_人员姓名);
              End If;
            End If;
            v_费用序号 := Null;
          End If;
          v_记录性质 := r_Rollmoney.记录性质;
          v_费用no   := r_Rollmoney.No;
          v_费用序号 := v_费用序号 || ',' || r_Rollmoney.序号;
        End If;
      End Loop;
      If v_费用序号 Is Not Null And v_费用no Is Not Null Then
        v_费用序号 := Substr(v_费用序号, 2);
        If v_记录性质 = 1 Then
          Zl_门诊划价记录_Delete(v_费用no, v_费用序号);
        Elsif v_记录性质 = 2 Then
          Zl_门诊记帐记录_Delete(v_费用no, v_费用序号, v_人员编号, v_人员姓名);
        End If;
      End If;
    
      --如果"门诊药嘱先作废后退药"，则对应的给药途径费用设置为未执行，以便退费
      If n_先作废后退药 = 1 Then
        Update 门诊费用记录
        Set 执行状态 = 0
        Where 执行状态 = 1 And 医嘱序号 = Id_In And Exists
         (Select 1
               From 病人医嘱记录 a, 诊疗项目目录 b
               Where a.诊疗项目id = b.Id And b.类别 = 'E' And b.操作类型 In ('2', '3', '4') And a.Id = Id_In);
      End If;
    
      --回退医嘱发送记录(及执行记录)
      Delete From 病人医嘱执行 Where 医嘱id In (Select Id From 病人医嘱记录 Where Id = Id_In Or 相关id = Id_In);
      Delete From 病人医嘱发送 Where 医嘱id In (Select Id From 病人医嘱记录 Where Id = Id_In Or 相关id = Id_In);
    
      --回退特殊医嘱的处理
      If r_Advice.诊疗类别 = 'Z' And Nvl(r_Advice.操作类型, '0') <> '0' And Nvl(r_Advice.婴儿, 0) = 0 Then
      
        If r_Advice.操作类型 = '1' And r_Advice.执行科室id Is Not Null Then
          --留观医嘱
          Select Count(*)
          Into v_Count
          From 病案主页
          Where 病人id = r_Advice.病人id And Nvl(主页id, 0) = 0 And 入院科室id = r_Advice.执行科室id And 病人性质 In (1, 2);
          If v_Count = 1 Then
            Zl_入院病案主页_Delete(r_Advice.病人id, 0);
          End If;
        Elsif r_Advice.操作类型 = '2' And r_Advice.执行科室id Is Not Null Then
          --住院医嘱
          Select Count(*)
          Into v_Count
          From 病案主页
          Where 病人id = r_Advice.病人id And Nvl(主页id, 0) = 0 And 入院科室id = r_Advice.执行科室id And Nvl(病人性质, 0) = 0;
          If v_Count = 1 Then
            Zl_入院病案主页_Delete(r_Advice.病人id, 0);
          End If;
        End If;
      End If;
    End If;
  End If;

  --删除过敏登记记录
  If r_Advice.诊疗类别 = 'E' And r_Advice.操作类型 = '1' Then
    --Update 病人医嘱记录 Set 皮试结果=Null Where ID=ID_IN; --保留最后的皮试结果
    --删除不过敏的记录，过敏记录保留，因为不管医嘱是否作废，病人对该药过敏
    For r_Test In (Select 操作时间 From 病人医嘱状态 Where 医嘱id = Id_In And 操作类型 = 10) Loop
      Delete From 病人过敏记录
      Where 病人id = r_Advice.病人id And 记录来源 = 2 And Nvl(主页id, 0) = Nvl(r_Advice.主页id, 0) And 记录时间 = r_Test.操作时间 And
            Nvl(结果, 0) = 0;
    End Loop;
  End If;
  If r_Advice.诊疗项目id Is Not Null Then
    --非自由录入医嘱，住院病区执行医嘱作废，Zlhis_Cis_003
    If r_Advice.主页id Is Not Null Then
      For r In (Select a.Id
                From 病人医嘱记录 a
                Where (a.Id = Id_In Or a.相关id = Id_In) And Exists
                 (Select 1 From 部门性质说明 b Where b.部门id = a.执行科室id And b.工作性质 = '护理')) Loop
        b_Message.Zlhis_Cis_003(r_Advice.病人id, r_Advice.主页id, Null, r.Id);
      End Loop;
    
      If r_Advice.诊疗类别 = 'Z' And r_Advice.操作类型 = '4' Then
        --术后医嘱作废
        b_Message.Zlhis_Cis_005(r_Advice.病人id, r_Advice.主页id, r_Advice.组id);
      End If;
    End If;
  
    If r_Advice.挂号单 Is Not Null Then
      If r_Advice.诊疗类别 = 'E' And r_Advice.操作类型 = '6' Or Instr(',D,F,K,', r_Advice.诊疗类别) > 0 Then
        Select Max(a.No) Into v_No From 病人医嘱发送 a Where a.医嘱id = r_Advice.组id;
        If r_Advice.诊疗类别 = 'E' And r_Advice.操作类型 = '6' Then
          --检验
          b_Message.Zlhis_Cis_036(r_Advice.病人id, Null, r_Advice.挂号单, v_发送号, r_Advice.组id, v_No, 1);
        Elsif r_Advice.诊疗类别 = 'D' Then
          --检查
          b_Message.Zlhis_Cis_037(r_Advice.病人id, Null, r_Advice.挂号单, v_发送号, r_Advice.组id, v_No, 1);
        Elsif r_Advice.诊疗类别 = 'F' Then
          --手术
          b_Message.Zlhis_Cis_038(r_Advice.病人id, Null, r_Advice.挂号单, v_发送号, r_Advice.组id, v_No);
        Elsif r_Advice.诊疗类别 = 'K' Then
          --输血
          b_Message.Zlhis_Cis_039(r_Advice.病人id, Null, r_Advice.挂号单, v_发送号, r_Advice.组id, v_No);
        End If;
      End If;
    End If;
  End If;
  Close c_Advice;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病人医嘱记录_作废;
/
--125595:李南春,2018-07-11,挂号序号错误问题
Create Or Replace Procedure Zl_病人预约挂号记录_Update
(
  单据号_In     门诊费用记录.No%Type,
  序号_In       门诊费用记录.序号%Type,
  价格父号_In   门诊费用记录.价格父号%Type,
  从属父号_In   门诊费用记录.从属父号%Type,
  收费类别_In   门诊费用记录.收费类别%Type,
  收费细目id_In 门诊费用记录.收费细目id%Type,
  数次_In       门诊费用记录.数次%Type,
  标准单价_In   门诊费用记录.标准单价%Type,
  收入项目id_In 门诊费用记录.收入项目id%Type,
  收据费目_In   门诊费用记录.收据费目%Type,
  应收金额_In   门诊费用记录.应收金额%Type,
  实收金额_In   门诊费用记录.实收金额%Type,
  病历费_In     Number, --该条记录是否病历工本费
  保险大类id_In 门诊费用记录.保险大类id%Type,
  保险项目否_In 门诊费用记录.保险项目否%Type,
  统筹金额_In   门诊费用记录.统筹金额%Type,
  保险编码_In   门诊费用记录.保险编码%Type,
  病人科室id_In 门诊费用记录.病人科室id%Type,
  执行部门id_In 门诊费用记录.执行部门id%Type,
  摘要_In       门诊费用记录.摘要%Type := Null,
  是否挂号项_In Number := 0
) As
  v_费用id 门诊费用记录.Id%Type;
  v_Error  Varchar2(255);
  Err_Custom Exception;
  Cursor c_费用 Is
    Select ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 记帐单id, 病人id, 医嘱序号, 门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号, 付款方式, 病人科室id, 费别,
           收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人, 开单部门id, 开单人, 发生时间,
           登记时间, 执行部门id, 执行人, 执行状态, 执行时间, 结论, 操作员编号, 操作员姓名, 结帐id, 结帐金额, 保险大类id, 保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要, 是否急诊
    From 门诊费用记录
    Where NO = 单据号_In And 记录性质 = 4 And 序号 = 1 And 记录状态 = 0;
Begin

  If Nvl(序号_In, 1) = 1 Then
    --第一条记录,只更新数据
    Update 门诊费用记录
    Set 价格父号 = Decode(价格父号_In, 0, Null, 价格父号_In), 从属父号 = Decode(从属父号_In, 0, Null, 从属父号_In), 附加标志 = 病历费_In,
        收费类别 = 收费类别_In, 收费细目id = 收费细目id_In, 收入项目id = 收入项目id_In, 收据费目 = 收据费目_In, 付数 = 1, 数次 = 数次_In, 标准单价 = 标准单价_In,
        应收金额 = 应收金额_In, 实收金额 = 实收金额_In, 保险大类id = 保险大类id_In, 保险项目否 = 保险项目否_In, 保险编码 = 保险编码_In, 统筹金额 = 统筹金额_In,
        病人科室id =  Decode(是否挂号项_In, 1, 病人科室id, 病人科室id_In), 执行部门id = Decode(是否挂号项_In, 1, 执行部门id, 执行部门id_In), 摘要 = Nvl(摘要_In, 摘要)
    Where NO = 单据号_In And 序号 = 1 And 记录状态 = 0 And 记录性质 = 4;
    --删除序号大于1的数据;
    Delete 门诊费用记录 Where NO = 单据号_In And 序号 > 1 And 记录性质 = 4;
  Else
    --插入数据
    If 病历费_In <> 3 Then
      Select 病人费用记录_Id.Nextval Into v_费用id From Dual; --应该通过程序得到
      For r_费用 In c_费用 Loop
        Insert Into 门诊费用记录
          (ID, 记录性质, 记录状态, 序号, 价格父号, 从属父号, NO, 实际票号, 门诊标志, 加班标志, 附加标志, 发药窗口, 病人id, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 病人科室id,
           收费类别, 计算单位, 收费细目id, 收入项目id, 收据费目, 付数, 数次, 标准单价, 应收金额, 实收金额, 结帐金额, 结帐id, 记帐费用, 开单部门id, 开单人, 划价人, 执行部门id, 执行人,
           操作员编号, 操作员姓名, 发生时间, 登记时间, 保险大类id, 保险项目否, 保险编码, 统筹金额, 摘要, 结论)
        Values
          (v_费用id, 4, 0, 序号_In, Decode(价格父号_In, 0, Null, 价格父号_In), 从属父号_In, 单据号_In, r_费用.实际票号, 1, r_费用.加班标志, 病历费_In,
           r_费用.发药窗口, r_费用.病人id, r_费用.标识号, r_费用.付款方式, r_费用.姓名, r_费用.性别, r_费用.年龄, r_费用.费别, 病人科室id_In, 收费类别_In, r_费用.计算单位,
           收费细目id_In, 收入项目id_In, 收据费目_In, 1, 数次_In, 标准单价_In, 应收金额_In, 实收金额_In, Null, Null, 0, r_费用.开单部门id, r_费用.操作员姓名,
           r_费用.操作员姓名, 执行部门id_In, r_费用.执行人, r_费用.操作员编号, r_费用.操作员姓名, r_费用.发生时间, r_费用.登记时间, 保险大类id_In, 保险项目否_In, 保险编码_In,
           统筹金额_In, Nvl(摘要_In, r_费用.摘要), r_费用.结论);
      End Loop;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人预约挂号记录_Update;
/

--125075:焦博,2018-07-10,调整使用到了部门参数分诊台签到排队的Oracle函数
--127075:焦博,2018-07-10,调整使用到了部门参数分诊台签到排队的Oracle函数
Create Or Replace Procedure Zl_病人挂号记录_签到
(
  Id_In       病人挂号记录.Id%Type,
  操作类型_In Integer := 0,
  预约方式_In 预约方式.名称%Type := Null,
  诊室_In     病人挂号记录.诊室%Type := Null,
  医生_In     病人挂号记录.执行人%Type := Null
  
) As
  --:操作类型_In:0-需要产生相关的排队记录或者重新签到,修改排队信息;1-不再处理排队记录,只作签到的标识填写； 
  Err_Item Exception;
  v_Err_Msg          Varchar2(200);
  n_挂号生成队列     Number;
  n_分诊台签到排队   Number;
  n_再次签到重新排队 Number;
  n_已排队           Number;
  v_原队列名称       排队叫号队列.队列名称%Type;
  v_现队列名称       排队叫号队列.队列名称%Type;
  v_病人姓名         病人挂号记录.姓名%Type;
  n_挂号id           病人挂号记录.Id%Type;
  n_执行部门id       病人挂号记录.执行部门id%Type;
  n_排队部门id       病人挂号记录.Id%Type;
  v_原始号码         排队叫号队列.排队号码%Type;
  v_排队号码         排队叫号队列.排队号码%Type;
  d_排队时间         排队叫号队列.排队时间%Type;
  v_诊室             病人挂号记录.诊室%Type;
  v_医生             病人挂号记录.执行人%Type;
  n_病人id           病人挂号记录.病人id%Type;
  v_号别             病人挂号记录.号别%Type;
  n_号序             病人挂号记录.号序%Type;
  v_排队序号         排队叫号队列.排队序号%Type;
  n_记录性质         病人挂号记录.记录性质%Type;
  d_发生时间         病人挂号记录.发生时间%Type;
  n_转诊             Number;
  n_换号             Number;
  v_挂号单           病人挂号记录.No%Type;
  v_预约方式         病人挂号记录.预约方式%Type;
  n_免挂号模式       Number;
  v_No               病人挂号记录.No%Type;
Begin
  --:记录标志:0表示就诊病人,1-表示签到的病人,2-表示需要回诊就诊的病人; 3-表示已回诊但还未接收的病人; 
  Update 病人挂号记录
  Set 记录标志 = 1, 诊室 = Nvl(诊室_In, 诊室), 执行人 = Nvl(医生_In, 执行人)
  Where ID = Id_In
  Returning 预约方式, NO Into v_预约方式, v_No;

  If Sql%NotFound Then
    v_Err_Msg := '挂号项目未找到,不能继续操作!';
    Raise Err_Item;
  End If;
  Update 门诊费用记录
  Set 执行人 = Nvl(医生_In, 执行人), 发药窗口 = Nvl(诊室_In, 发药窗口)
  Where NO = v_No And 记录性质 = 4 And 记录状态 In (0, 1, 3);

  If Nvl(操作类型_In, 0) = 1 Then
    Return;
  End If;

  If 预约方式_In Is Not Null Then
    v_预约方式 := 预约方式_In;
  End If;

  Begin
    Select ID, Nvl(转诊科室id, 执行部门id), 姓名, Decode(转诊科室id, Null, 诊室, 转诊诊室), Decode(转诊科室id, Null, 执行人, 转诊医生), 病人id,
           Nvl(转诊号别, 号别), 号序, Nvl(记录性质, 0), Decode(转诊科室id, Null, 0, 1), NO, 发生时间
    Into n_挂号id, n_执行部门id, v_病人姓名, v_诊室, v_医生, n_病人id, v_号别, n_号序, n_记录性质, n_转诊, v_挂号单, d_发生时间
    From 病人挂号记录
    Where ID = Id_In And Rownum = 1;
  Exception
    When Others Then
      n_挂号id := -1;
  End;

  n_挂号生成队列     := Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113));
  n_分诊台签到排队   := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113, 1, Nvl(n_执行部门id, 0)));
  n_再次签到重新排队 := Zl_To_Number(zl_GetSysParameter('再次签到需重新排队', 1113));

  n_免挂号模式 := Zl_To_Number(zl_GetSysParameter('免挂号模式', 0));
  If Nvl(n_免挂号模式, 0) = 1 And Nvl(n_分诊台签到排队, 0) = 0 Then
    n_分诊台签到排队 := 1;
  End If;

  If n_挂号生成队列 = 0 Then
    Return;
  End If;

  Begin
    Select 1, 科室id, 排队号码, 队列名称
    Into n_已排队, n_排队部门id, v_原始号码, v_原队列名称
    From 排队叫号队列
    Where 业务id = Id_In And 业务类型 = 0;
  Exception
    When Others Then
      n_已排队 := 0;
  End;
  --如果有排队记录，说明是重新签到，不应该退出 
  If Nvl(n_分诊台签到排队, 0) = 0 And Nvl(n_已排队, 0) = 0 Then
    Return;
  End If;

  Begin
    Select 1 Into n_换号 From 就诊变动记录 Where 挂号单 = v_挂号单 And Rownum = 1;
  Exception
    When Others Then
      n_换号 := 0;
  End;

  v_现队列名称 := n_执行部门id;
  If n_挂号id > 0 Then
    If n_已排队 > 0 And (Nvl(n_再次签到重新排队, 0) = 0 Or Nvl(n_记录性质, 0) = 2) Then
      v_排队号码 := Zl_Get_Requeue(0, n_挂号id, n_执行部门id, v_医生, v_诊室);
      --重新获取排队时间 
      d_排队时间 := Zl_Get_Requeuedate(0, n_挂号id, n_执行部门id, v_医生, v_诊室);
      If v_原始号码 <> v_排队号码 Then
        --重新获取序号 
        v_排队序号 := Zlgetsequencenum(0, n_挂号id, 1);
        --新队列名称_IN, 业务类型_In, 业务id_In , 科室id_In , 患者姓名_In , 诊室_In , 医生姓名_In 
        Zl_排队叫号队列_Update(v_现队列名称, 0, n_挂号id, n_执行部门id, v_病人姓名, v_诊室, v_医生, v_排队号码, v_排队序号, d_排队时间);
      Else
        --新队列名称_IN, 业务类型_In, 业务id_In , 科室id_In , 患者姓名_In , 诊室_In , 医生姓名_In 
        Zl_排队叫号队列_Update(v_现队列名称, 0, n_挂号id, n_执行部门id, v_病人姓名, v_诊室, v_医生);
      End If;
    Elsif n_已排队 > 0 Then
      --重新排队 
      v_排队号码 := Zlgetnextqueue(n_执行部门id, n_挂号id, v_号别 || '|' || n_号序);
      v_排队序号 := Zlgetsequencenum(0, n_挂号id, 0);
      --重新获取排队时间 
      d_排队时间 := Zl_Get_Requeuedate(0, n_挂号id, n_执行部门id, v_医生, v_诊室);
      --新队列名称_IN, 业务类型_In, 业务id_In , 科室id_In , 患者姓名_In , 诊室_In , 医生姓名_In 
      Zl_排队叫号队列_Update(v_现队列名称, 0, n_挂号id, n_执行部门id, v_病人姓名, v_诊室, v_医生, v_排队号码, v_排队序号, d_排队时间);
    Else
      --正常排队 
      For v_队列 In (Select 队列名称, 业务id
                   From 排队叫号队列
                   Where 病人id = n_病人id And 业务类型 = 0 And Trunc(排队时间) < Sysdate) Loop
        --删除原来排队记录重新排队：队列名称_IN，业务ID_IN 
        Zl_排队叫号队列_Delete(v_队列.队列名称, v_队列.业务id);
      End Loop;
    
      v_排队号码 := Zlgetnextqueue(n_执行部门id, n_挂号id, v_号别 || '|' || n_号序);
      v_排队序号 := Zlgetsequencenum(0, n_挂号id, 0);
      --重新获取排队时间 
      If n_转诊 = 1 Then
        d_排队时间 := Zl_Get_Requeuedate(2, n_挂号id, n_执行部门id, v_医生, v_诊室);
      Elsif n_换号 = 1 Then
        d_排队时间 := Zl_Get_Requeuedate(3, n_挂号id, n_执行部门id, v_医生, v_诊室);
      Else
        d_排队时间 := Zl_Get_Requeuedate(-1, n_挂号id, n_执行部门id, v_医生, v_诊室);
      End If;
      --  队列名称_In , 业务类型_In, 业务id_In,科室id_In,排队号码_In,排队标记_In,患者姓名_In,病人ID_IN, 诊室_In, 医生姓名_In, 排队时间_In 
      Zl_排队叫号队列_Insert(v_现队列名称, 0, n_挂号id, n_执行部门id, v_排队号码, Null, v_病人姓名, n_病人id, v_诊室, v_医生, d_排队时间, v_预约方式, Null,
                       v_排队序号);
    End If;
    Update 排队叫号队列 Set 排队状态 = 0 Where 业务id = n_挂号id And 业务类型 = 0;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人挂号记录_签到;
/

--126898:秦龙,2018-07-09,增加传参“批准文号”
Create Or Replace Procedure Zl_药品计划管理次表_Insert
(
  计划id_In     In 药品计划内容.计划id%Type,
  药品id_In     In 药品计划内容.药品id%Type,
  序号_In       In 药品计划内容.序号%Type,
  计划数量_In   In 药品计划内容.计划数量%Type,
  单价_In       In 药品计划内容.单价%Type := Null,
  金额_In       In 药品计划内容.金额%Type := Null,
  前期数量_In   In 药品计划内容.前期数量%Type := Null,
  上期数量_In   In 药品计划内容.上期数量%Type := Null,
  库存数量_In   In 药品计划内容.库存数量%Type := Null,
  上次供应商_In In 药品计划内容.上次供应商%Type := Null,
  上次生产商_In In 药品计划内容.上次生产商%Type := Null,
  说明_In       In 药品计划内容.说明%Type := Null,
  售价_In       In 药品计划内容.售价%Type := Null,
  售价金额_In   In 药品计划内容.售价金额%Type := Null,
  上期销量_In   In 药品计划内容.上期销量%Type := Null,
  本期销量_In   In 药品计划内容.本期销量%Type := Null,
  送货数量_In   In 药品计划内容.送货数量%Type := Null,
  批准文号_In   In 药品计划内容.批准文号%Type := Null
) Is
Begin
  Insert Into 药品计划内容
    (计划id, 药品id, 序号, 前期数量, 上期数量, 库存数量, 计划数量, 单价, 金额, 上次供应商, 上次生产商, 说明, 售价, 售价金额, 上期销量, 本期销量, 送货数量, 批准文号)
  Values
    (计划id_In, 药品id_In, 序号_In, 前期数量_In, 上期数量_In, 库存数量_In, 计划数量_In, 单价_In, 金额_In, 上次供应商_In, 上次生产商_In, 说明_In, 售价_In,
     售价金额_In, 上期销量_In, 本期销量_In, 送货数量_In, 批准文号_In);
End Zl_药品计划管理次表_Insert;
/

--125595:李南春,2018-07-11,挂号序号错误问题
--127075:焦博,2018-07-06,调整使用到了部门参数分诊台签到排队的Oracle函数
Create Or Replace Procedure Zl_病人挂号记录_出诊_Insert
(
  出诊记录id_In    临床出诊记录.Id%Type,
  病人id_In        门诊费用记录.病人id%Type,
  门诊号_In        门诊费用记录.标识号%Type,
  姓名_In          门诊费用记录.姓名%Type,
  性别_In          门诊费用记录.性别%Type,
  年龄_In          门诊费用记录.年龄%Type,
  付款方式_In      门诊费用记录.付款方式%Type, --用于存放病人的医疗付款方式编号
  费别_In          门诊费用记录.费别%Type,
  单据号_In        门诊费用记录.No%Type,
  票据号_In        门诊费用记录.实际票号%Type,
  序号_In          门诊费用记录.序号%Type,
  价格父号_In      门诊费用记录.价格父号%Type,
  从属父号_In      门诊费用记录.从属父号%Type,
  收费类别_In      门诊费用记录.收费类别%Type,
  收费细目id_In    门诊费用记录.收费细目id%Type,
  数次_In          门诊费用记录.数次%Type,
  标准单价_In      门诊费用记录.标准单价%Type,
  收入项目id_In    门诊费用记录.收入项目id%Type,
  收据费目_In      门诊费用记录.收据费目%Type,
  结算方式_In      Varchar2,
  应收金额_In      门诊费用记录.应收金额%Type,
  实收金额_In      门诊费用记录.实收金额%Type,
  病人科室id_In    门诊费用记录.病人科室id%Type,
  开单部门id_In    门诊费用记录.开单部门id%Type,
  执行部门id_In    门诊费用记录.执行部门id%Type,
  操作员编号_In    门诊费用记录.操作员编号%Type,
  操作员姓名_In    门诊费用记录.操作员姓名%Type,
  发生时间_In      门诊费用记录.发生时间%Type,
  登记时间_In      门诊费用记录.登记时间%Type,
  医生姓名_In      挂号安排.医生姓名%Type,
  医生id_In        挂号安排.医生id%Type,
  病历费_In        Number, --该条记录是否病历工本费
  急诊_In          Number,
  号别_In          挂号安排.号码%Type,
  诊室_In          门诊费用记录.发药窗口%Type,
  结帐id_In        门诊费用记录.结帐id%Type,
  领用id_In        票据使用明细.领用id%Type,
  预交支付_In      病人预交记录.冲预交%Type, --刷卡挂号时使用的预交金额,序号为1传入.
  现金支付_In      病人预交记录.冲预交%Type, --挂号时现金支付部份金额,序号为1传入.
  个帐支付_In      病人预交记录.冲预交%Type, --挂号时个人帐户支付金额,,序号为1传入.
  保险大类id_In    门诊费用记录.保险大类id%Type,
  保险项目否_In    门诊费用记录.保险项目否%Type,
  统筹金额_In      门诊费用记录.统筹金额%Type,
  摘要_In          门诊费用记录.摘要%Type, --预约挂号摘要信息
  预约挂号_In      Number := 0, --预约挂号时用(记录状态=0,发生时间为预约时间),此时不需要传入结算相关参数
  收费票据_In      Number := 0, --挂号是否使用收费票据
  保险编码_In      门诊费用记录.保险编码%Type,
  复诊_In          病人挂号记录.复诊%Type := 0,
  号序_In          挂号序号状态.序号%Type := Null, --预约时填入费用记录的发药窗口字段,挂号时填入挂号记录
  社区_In          病人挂号记录.社区%Type := Null,
  预约接收_In      Number := 0,
  预约方式_In      预约方式.名称%Type := Null,
  生成队列_In      Number := 0,
  卡类别id_In      病人预交记录.卡类别id%Type := Null,
  结算卡序号_In    病人预交记录.结算卡序号%Type := Null,
  卡号_In          病人预交记录.卡号%Type := Null,
  交易流水号_In    病人预交记录.交易流水号%Type := Null,
  交易说明_In      病人预交记录.交易说明%Type := Null,
  合作单位_In      病人预交记录.合作单位%Type := Null,
  操作类型_In      Number := 0,
  险类_In          病人挂号记录.险类%Type := Null,
  结算模式_In      Number := 0,
  记帐费用_In      Number := 0,
  退号重用_In      Number := 1,
  冲预交病人ids_In Varchar2 := Null,
  修正病人费别_In  Number := 0,
  预约顺序号_In    临床出诊序号控制.预约顺序号%Type := Null,
  修正病人年龄_In  Number := 0,
  收费单_In        病人挂号记录.收费单%Type := Null,
  更新交款余额_In  Number := 1 --是否更新人员交款余额，主要是处理统一操作员登录多台自助机的情况
) As
  ---------------------------------------------------------------------------
  --
  --参数:
  --     操作类型_in:0-正常挂号或者预约 1-操作员拥有加号权限加号
  --     修正病人费别_In:0-不修改病人费别 1-修改病人费别
  ----------------------------------------------------------------------------
  --该游标用于收费冲预交的可用预交列表
  --以ID排序，优先冲上次未冲完的。
  Cursor c_Deposit
  (
    v_病人id        病人信息.病人id%Type,
    v_冲预交病人ids Varchar2
  ) Is
    Select 病人id, NO, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额, Min(记录状态) As 记录状态, Nvl(Max(结帐id), 0) As 结帐id,
           Max(Decode(记录性质, 1, ID, 0) * Decode(记录状态, 2, 0, 1)) As 原预交id, Min(Decode(记录性质, 1, 收款时间, Null)) As 收款时间
    From 病人预交记录
    Where 记录性质 In (1, 11) And 病人id In (Select Column_Value From Table(f_Num2list(v_冲预交病人ids))) And Nvl(预交类别, 2) = 1
     Having Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
    Group By NO, 病人id
    Order By Decode(病人id, Nvl(v_病人id, 0), 0, 1), 收款时间;
  --功能：新增一行病人挂号费用，并将结算汇总到病人预交记录
  --       同时汇总相关的汇总表(病人挂号汇总、费用汇总)
  --       第一行费用处理票据使用情况(领用ID_IN>0)

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_排队号码 排队叫号队列.排队号码%Type;
  v_现金     结算方式.名称%Type;
  v_个人帐户 结算方式.名称%Type;
  v_队列名称 排队叫号队列.队列名称%Type;

  n_分时段       Number;
  n_原始分时段   Number;
  n_时段限号     Number;
  n_时段限约     Number;
  d_时段时间     Date;
  d_最大序号时间 Date;
  n_追加号       Number := 0; --处理时段过期 追加挂号的情况
  n_已约数       病人挂号汇总.已约数%Type;
  n_预约有效时间 Number;
  n_失效数       Number;
  n_失约挂号     Number := 0;
  n_已用数量     Number;
  n_锁定         Number := 0;

  n_费用id        门诊费用记录.Id%Type;
  n_预交金额      病人预交记录.金额%Type;
  n_当前金额      病人预交记录.金额%Type;
  n_返回值        病人预交记录.金额%Type;
  n_预交id        病人预交记录.Id%Type;
  n_挂号id        病人挂号记录.Id%Type;
  v_冲预交病人ids Varchar2(4000);

  n_组id           财务缴款分组.Id%Type;
  n_门诊号         病人信息.门诊号%Type;
  n_序号           挂号序号状态.序号%Type;
  n_已用序号       挂号序号状态.序号%Type;
  n_序号控制       挂号安排.序号控制%Type;
  n_分诊台签到排队 Number;
  n_Count          Number;
  n_限号数         Number(18);
  d_排队时间       Date;
  v_结算方式记录   Varchar2(1000);
  d_序号时间       Date;
  v_费别           费别.名称%Type;
  n_时段序号       Number := -1;
  n_预约生成队列   Number;
  v_结算方式       结算方式.名称%Type;
  v_结算内容       Varchar2(1000);
  v_当前结算       Varchar2(200);
  v_结算号码       病人预交记录.结算号码%Type;
  n_结算金额       病人预交记录.冲预交%Type;
  n_三方卡标志     Number(2);
  n_预约顺序号     临床出诊序号控制.预约顺序号%Type;
  n_限约数         Number(18);
  n_已挂数         Number(4) := 0;
  n_Exists         Number;
  n_挂出的最大序号 Number(4) := 0;
  n_分时点显示     Number(3);
  n_结算模式       病人信息.结算模式%Type;
  v_排队序号       排队叫号队列.排队序号%Type;
  v_机器名         挂号序号状态.机器名%Type;
  v_序号操作员     挂号序号状态.操作员姓名%Type;
  v_序号机器名     挂号序号状态.机器名%Type;
  v_付款方式       病人挂号记录.医疗付款方式%Type;
  n_状态           临床出诊序号控制.挂号状态%Type;
Begin
  --记录锁定判断
  If 出诊记录id_In Is Not Null Then
    Begin
      Select 1
      Into n_Exists
      From 临床出诊记录
      Where ID = 出诊记录id_In And Nvl(是否发布, 0) = 1 And Nvl(是否锁定, 0) = 0;
    Exception
      When Others Then
        v_Err_Msg := '无法确定出诊记录，请检查出诊记录是否存在或被锁定！';
        Raise Err_Item;
    End;
  End If;

  --获取当前机器名称
  Select Terminal Into v_机器名 From V$session Where Audsid = Userenv('sessionid');
  n_组id          := Zl_Get组id(操作员姓名_In);
  v_冲预交病人ids := Nvl(冲预交病人ids_In, 病人id_In);

  If 费别_In Is Null Then
    Begin
      Select 名称 Into v_费别 From 费别 Where 缺省标志 = 1 And Rownum < 2;
      Update 病人信息 Set 费别 = v_费别 Where 病人id = 病人id_In;
    Exception
      When Others Then
        v_Err_Msg := '无法确定病人费别，请检查缺省费别是否正确设置！';
        Raise Err_Item;
    End;
  Else
    v_费别 := 费别_In;
    If Nvl(修正病人费别_In, 0) = 1 Then
      Begin
        Update 病人信息 Set 费别 = v_费别 Where 病人id = 病人id_In;
      Exception
        When Others Then
          v_Err_Msg := '没有找到对应的病人！';
          Raise Err_Item;
      End;
    End If;
  End If;

  If Nvl(修正病人年龄_In, 0) = 1 Then
    Begin
      Update 病人信息 Set 年龄 = 年龄_In Where 病人id = 病人id_In;
    Exception
      When Others Then
        v_Err_Msg := '没有找到对应的病人！';
        Raise Err_Item;
    End;
  End If;

  If 门诊号_In Is Not Null Then
    Begin
      Select Nvl(门诊号, 0) Into n_门诊号 From 病人信息 Where 病人id = 病人id_In;
    Exception
      When Others Then
        n_门诊号 := 0;
    End;
    If n_门诊号 = 0 Then
      Update 病人信息 Set 门诊号 = 门诊号_In Where 病人id = 病人id_In;
    End If;
  End If;

  Begin
    Update 临床出诊序号控制
    Set 挂号状态 = 0
    Where 记录id = 出诊记录id_In And 序号 = 号序_In And Nvl(挂号状态, 0) = 3 And 操作员姓名 = 操作员姓名_In;
  Exception
    When Others Then
      Null;
  End;

  If Nvl(预约挂号_In, 0) = 0 Then
    --挂号或者预约接收
    --因为现在有按日编单据号规则,日挂号量不能超过10000次,所以要检查唯一约束。
    Select Count(*)
    Into n_Count
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 In (1, 3) And 序号 = 序号_In And NO = 单据号_In;
    If n_Count <> 0 Then
      v_Err_Msg := '挂号单据号重复,不能保存！' || Chr(13) || '如果使用了按日顺序编号,当日挂号量不能超过10000人次。';
      Raise Err_Item;
    End If;
  
    --获取结算方式名称
    Begin
      Select 名称 Into v_现金 From 结算方式 Where 性质 = 1;
    Exception
      When Others Then
        v_现金 := '现金';
    End;
    Begin
      Select 名称 Into v_个人帐户 From 结算方式 Where 性质 = 3;
    Exception
      When Others Then
        v_个人帐户 := '个人帐户';
    End;
  End If;

  n_序号 := 号序_In;

  --获取是否分时段
  Begin
    Select Nvl(是否分时段, 0), Nvl(是否序号控制, 0), 限号数, 限约数
    Into n_分时段, n_序号控制, n_限号数, n_限约数
    From 临床出诊记录
    Where ID = 出诊记录id_In;
    n_原始分时段 := n_分时段;
  Exception
    When Others Then
      n_分时段     := 0;
      n_原始分时段 := n_分时段;
      n_序号控制   := 0;
      n_限号数     := Null;
      n_限约数     := Null;
  End;

  If n_序号 Is Null And n_分时段 = 1 And n_序号控制 = 0 Then
    Begin
      Select 序号
      Into n_序号
      From 临床出诊序号控制
      Where 记录id = 出诊记录id_In And 开始时间 = 发生时间_In And Rownum < 2;
    Exception
      When Others Then
        n_序号 := Null;
    End;
  End If;

  --分时段挂号时判断是否是过期挂号 也就是追加号的情况 目前只针对专家号分时段进行处理
  If Nvl(预约挂号_In, 0) = 0 And n_分时段 > 0 And Nvl(n_序号控制, 0) = 1 And 号序_In Is Null And Nvl(操作类型_In, 0) = 0 Then
    Begin
      Select Max(To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'))
      Into d_最大序号时间
      From 临床出诊序号控制
      Where 记录id = 出诊记录id_In And Nvl(数量, 0) <> 0;
    
      n_追加号 := Case Sign(发生时间_In - d_最大序号时间)
                 When -1 Then
                  0
                 Else
                  1
               End;
    Exception
      When Others Then
        n_追加号 := 0;
    End;
  End If;
  d_时段时间 := 发生时间_In;

  If 序号_In = 1 And n_分时段 > 0 Then
    If Nvl(n_序号控制, 0) = 1 Then
      Begin
        Select Nvl(序号, 0),
               To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
               数量, Decode(Nvl(是否预约, 0), 0, 0, 数量)
        Into n_时段序号, d_时段时间, n_时段限号, n_时段限约
        From 临床出诊序号控制
        Where 记录id = 出诊记录id_In And 序号 = n_序号;
      Exception
        When Others Then
          n_时段序号 := -1;
          n_分时段   := 0;
          d_时段时间 := 发生时间_In;
          n_时段限号 := 0;
          n_时段限约 := 0;
      End;
    Else
      --挂号时检查 是否分了时段,分了时段,把时段的限制条件给取出来
      Begin
        Select Nvl(序号, 0),
               To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
               数量, Decode(Nvl(是否预约, 0), 0, 0, 数量)
        Into n_时段序号, d_时段时间, n_时段限号, n_时段限约
        From 临床出诊序号控制
        Where 记录id = 出诊记录id_In And 序号 = n_序号 And 预约顺序号 Is Null;
      Exception
        When Others Then
          n_时段序号 := -1;
          n_分时段   := 0;
          d_时段时间 := 发生时间_In;
          n_时段限号 := 0;
          n_时段限约 := 0;
      End;
    End If;
  End If;

  If 序号_In = 1 Then
    --获取当前未使用的序号
    If Nvl(预约挂号_In, 0) = 0 Then
      n_预约有效时间 := Zl_To_Number(zl_GetSysParameter('预约有效时间', 1111));
      n_失约挂号     := Zl_To_Number(zl_GetSysParameter('失约用于挂号', 1111));
    End If;
    If Nvl(n_序号控制, 0) = 1 And n_分时段 = 0 Then
      --<序号控制 未设置时段 获取可用的最大序号,以及已经使用的数量>
      Begin
        --最大序号
        Select Count(1) Into n_已用数量 From 病人挂号记录 Where 出诊记录id = 出诊记录id_In And 记录状态 = 1;
        Select Max(序号) Into n_已用序号 From 临床出诊序号控制 Where 记录id = 出诊记录id_In;
      Exception
        When Others Then
          n_已用序号 := 0;
          n_已用数量 := 0;
      End;
      Begin
        --最大序号
        Select Sum(Nvl(数量, 0))
        
        Into n_已约数
        From 临床出诊序号控制
        Where 记录id = 出诊记录id_In And Nvl(挂号状态, 0) = 2;
      Exception
        When Others Then
          n_已约数 := 0;
      End;
    
      n_失效数 := 0;
      If Nvl(预约挂号_In, 0) = 0 Then
        If Nvl(n_预约有效时间, 0) <> 0 And Nvl(n_失约挂号, 0) > 0 Then
          Begin
            Select Sum(Decode(Sign((Sysdate - n_预约有效时间 / 24 / 60) - 预约时间), 1, 1, 0))
            Into n_失效数
            From 病人挂号记录
            Where 出诊记录id = 出诊记录id_In And 记录状态 = 1 And 记录性质 = 2;
          Exception
            When Others Then
              n_失效数 := 0;
          End;
        End If;
      End If;
    
      If n_原始分时段 = 0 Then
        Select Min(序号) Into n_已用序号 From 临床出诊序号控制 Where 记录id = 出诊记录id_In And Nvl(挂号状态, 0) = 0;
        If n_序号 Is Null Then
          n_序号 := Nvl(n_已用序号, 0);
        End If;
        IF nvl(n_序号,0)=0 THEN 
          Select Nvl(Max(序号), 0) + 1 Into n_序号 From 临床出诊序号控制 Where 记录id = 出诊记录id_In;
	      END IF;
      Else
        Select Max(序号) Into n_已用序号 From 临床出诊序号控制 Where 记录id = 出诊记录id_In;
        If n_序号 Is Null Then
          n_序号 := Nvl(n_已用序号, 0) + 1;
        End If;
      End If;
      --<序号控制 未设置时段 获取可用的最大序号 以及已经使用的数量 --end>
    
      --非加号的情况需要检查是否超过了限制
      If 操作类型_In = 0 Then
        If Nvl(预约挂号_In, 0) = 0 Then
          --挂号时检查
          --启用序号控制未分时段 达到了限制
          If n_限号数 <= n_已用数量 And n_限号数 > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(Trunc(发生时间_In), 'yyyy-mm-dd ') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        
        Else
          --预约时检查
          If n_限约数 = 0 Then
            n_限约数 := n_限号数;
          End If;
        
          If n_限约数 <= n_已约数 And n_限约数 > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(Trunc(发生时间_In), 'yyyy-mm-dd') || '已达到最大限约数！';
            Raise Err_Item;
          End If;
        End If;
      
      Else
        Null;
        --启用序号控制,未分时段 加号情况   不处理,如果以后有限制条件以后补充
      End If;
    
    Elsif Nvl(n_分时段, 0) > 0 And Nvl(n_序号控制, 0) = 0 Then
      --<--普通号分时段 这里只有预约一种情况-->
      If 操作类型_In = 0 Then
        --<正常预约挂号-->
        Begin
          Select Count(0) As 已挂数, Nvl(Sum(Decode(Nvl(Sign(a.开始时间 - d_时段时间), 0), 0, 1, 0)), 0) As 已约数
          Into n_已挂数, n_已约数
          From 临床出诊序号控制 A
          Where 记录id = 出诊记录id_In And 挂号状态 Not In (0, 4, 5);
        Exception
          When Others Then
            n_已挂数 := 0;
            n_已约数 := 0;
        End;
      
        n_时段限约 := n_时段限号; --普通号分时段的情况,29中n_时段限约始终是0 这里特殊处理
        --检查限制数量
        If n_限约数 = 0 Then
          n_限约数 := n_限号数;
        End If;
        If n_时段限约 <= n_已约数 Or n_限约数 <= n_已约数 Then
          v_Err_Msg := '号别' || 号别_In || '在时段' || To_Char(d_时段时间, 'hh24:mi') || '已达到最大限约数！';
          Raise Err_Item;
        End If;
        If n_限号数 <= n_已挂数 Then
          v_Err_Msg := '号别' || 号别_In || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
          Raise Err_Item;
        End If;
      End If;
    
      --没有达到时段的限号数 号码在当前时段往后追加
    
      --获取当天挂出的最大号序
      Select Nvl(Max(序号), 0)
      Into n_挂出的最大序号
      From 临床出诊序号控制 A
      Where 记录id = 出诊记录id_In And 预约顺序号 Is Null And 挂号状态 Not In (0, 5);
      If 预约顺序号_In Is Not Null Then
        n_预约顺序号 := 预约顺序号_In;
      Else
        Begin
          Select Nvl(Max(预约顺序号), 0) + 1
          Into n_预约顺序号
          From 临床出诊序号控制
          Where 记录id = 出诊记录id_In And 序号 = n_时段序号 And 预约顺序号 Is Not Null;
        Exception
          When Others Then
            n_预约顺序号 := Null;
        End;
      End If;
      --设置序号
      n_序号 := RPad(Nvl(n_时段序号, 0), Length(n_限号数) + Length(Nvl(n_时段序号, 0)), 0) + n_预约顺序号;
      If n_预约顺序号 Is Null Then
        n_序号 := Nvl(n_挂出的最大序号, 0) + 1;
      End If;
    
      --<--普通号分时段--End>
    Elsif Nvl(n_分时段, 0) > 0 And Nvl(n_序号控制, 0) = 1 Then
      --<启用序号控制 设置时段
      --专家号分时段
      Begin
        Select Max(序号), Sum(Decode(序号, Nvl(号序_In, 0), 0, 1)), Sum(Decode(Sign(开始时间 - d_时段时间), 0, 1, 0))
        Into n_已用序号, n_已挂数, n_已用数量
        From 临床出诊序号控制
        Where 记录id = 出诊记录id_In And 挂号状态 Not In (0, 4, 5);
      Exception
        When Others Then
          n_已用序号 := 0;
          n_已用数量 := 0;
          n_已挂数   := 0;
      End;
    
      n_失效数 := 0;
      If Nvl(预约挂号_In, 0) = 0 Then
        If Nvl(n_预约有效时间, 0) <> 0 And Nvl(n_失约挂号, 0) > 0 Then
          Begin
            Select Sum(Decode(Sign((Sysdate - n_预约有效时间 / 24 / 60) - 开始时间), 1, 1, 0))
            Into n_失效数
            From 临床出诊序号控制
            Where 记录id = 出诊记录id_In And 开始时间 Between Trunc(Sysdate) And Sysdate And Nvl(挂号状态, 0) = 2;
          Exception
            When Others Then
              n_失效数 := 0;
          End;
        End If;
      End If;
    
      If 操作类型_In = 0 Then
        If Nvl(预约挂号_In, 0) = 0 Then
        
          --挂号 追加号码时不检查时段限号数
          If n_时段限号 <= n_已用数量 And Nvl(n_追加号, 0) = 0 Then
            v_Err_Msg := '号别' || 号别_In || '在时段' || To_Char(d_时段时间, 'hh24:mi') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
          If n_限号数 <= n_已挂数 - n_失效数 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        
        Else
          --挂号
          If n_限约数 = 0 Then
            n_限约数 := n_限号数;
          End If;
          If n_限约数 <= n_已用数量 Then
            v_Err_Msg := '号别' || 号别_In || '在时段' || To_Char(d_时段时间, 'hh24:mi') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
          If n_限号数 <= n_已挂数 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        End If;
      End If;
      If n_序号 Is Null Then
        --设置序号
        If Nvl(n_已用序号, 0) < Nvl(n_时段序号, 0) Then
          n_已用序号 := Nvl(n_时段序号, 0);
        End If;
        n_序号 := Nvl(n_已用序号, 0) + 1;
      End If;
    Elsif Nvl(n_分时段, 0) = 0 And Nvl(n_序号控制, 0) = 0 And Nvl(病历费_In, 0) = 0 And Nvl(号别_In, 0) > 0 Then
      ---<--普通号  -->
      Begin
        Select 已挂数, 已约数 Into n_已用数量, n_已约数 From 临床出诊记录 Where ID = 出诊记录id_In;
      Exception
        When Others Then
          n_已用数量 := 0;
          n_已约数   := 0;
      End;
      If Nvl(操作类型_In, 0) = 0 Then
        If Nvl(预约挂号_In, 0) = 0 Then
          --挂号
          If Nvl(n_已用数量, 0) >= Nvl(n_限号数, 0) And Nvl(n_限号数, 0) > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        Else
          --预约
          If (Nvl(n_限约数, 0) > 0 And Nvl(n_已约数, 0) >= Nvl(n_限约数, 0)) Or
             (Nvl(n_已用数量, 0) >= Nvl(n_限号数, 0) And Nvl(n_限号数, 0) > 0) Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        End If;
      End If;
      ---<--普通号  -->
    End If;
  End If;

  --更新挂号序号状态
  If 序号_In = 1 And Not n_序号 Is Null Then
    If n_分时段 = 1 Then
      d_序号时间 := 发生时间_In;
    Else
      d_序号时间 := Trunc(发生时间_In);
    End If;
    --锁定序号的处理
    Begin
      If n_预约顺序号 Is Null Then
        Select 操作员姓名, 工作站名称
        Into v_序号操作员, v_序号机器名
        From 临床出诊序号控制
        Where Nvl(挂号状态, 0) = 5 And 记录id = 出诊记录id_In And 序号 = n_序号;
      Else
        Select 操作员姓名, 工作站名称
        Into v_序号操作员, v_序号机器名
        From 临床出诊序号控制
        Where Nvl(挂号状态, 0) = 5 And 记录id = 出诊记录id_In And 序号 = n_时段序号 And 预约顺序号 = n_预约顺序号;
      End If;
      n_锁定 := 1;
    Exception
      When Others Then
        v_序号操作员 := Null;
        v_序号机器名 := Null;
        n_锁定       := 0;
    End;
    If n_锁定 = 0 Then
      If n_预约顺序号 Is Null Then
        Update 临床出诊序号控制
        Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1)
        Where 记录id = 出诊记录id_In And 序号 = n_序号 And Nvl(挂号状态, 0) = 3 And 操作员姓名 = 操作员姓名_In;
      Else
        Update 临床出诊序号控制
        Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1)
        Where 记录id = 出诊记录id_In And 序号 = n_序号 And 预约顺序号 = n_预约顺序号 And Nvl(挂号状态, 0) = 3 And 操作员姓名 = 操作员姓名_In;
      End If;
      If Sql%RowCount = 0 Then
        Begin
          If Nvl(n_分时段, 0) > 0 Then
            If Nvl(n_序号控制, 0) = 1 Then
              --分时段后专家号 失约的预约号允许挂号
              Update 临床出诊序号控制
              Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1), 操作员姓名 = 操作员姓名_In
              Where 记录id = 出诊记录id_In And 序号 = n_序号 And Nvl(挂号状态, 0) In (0, 2);
              If Sql%NotFound Then
                Begin
                  Select 挂号状态 Into n_状态 From 临床出诊序号控制 Where 记录id = 出诊记录id_In And 序号 = n_序号;
                Exception
                  When Others Then
                    n_状态 := -1;
                End;
                If n_状态 = -1 Then
                  Insert Into 临床出诊序号控制
                    (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 锁号时间, 类型, 名称, 操作员姓名, 备注)
                    Select 出诊记录id_In, n_序号, d_序号时间, d_序号时间, 1, Decode(预约挂号_In, 1, 1, 0), Decode(预约挂号_In, 1, 2, 1), Null,
                           Null, Null, 操作员姓名_In, '追加号'
                    From Dual;
                Else
                  v_Err_Msg := '序号' || n_序号 || '已被使用,请重新选择一个序号.';
                  Raise Err_Item;
                End If;
              End If;
            Else
              If Nvl(预约接收_In, 0) = 1 Then
                Insert Into 临床出诊序号控制
                  (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 锁号时间, 类型, 名称, 操作员姓名, 备注, 预约顺序号)
                  Select 记录id, 序号, 开始时间, 终止时间, 1, 1, Decode(预约挂号_In, 1, 2, 1), Null, Null, Null, 操作员姓名_In, n_序号, n_预约顺序号
                  From 临床出诊序号控制
                  Where 记录id = 出诊记录id_In And 序号 = n_时段序号 And 预约顺序号 Is Null;
              End If;
            End If;
          Else
            If Nvl(n_序号控制, 0) = 1 Then
              Update 临床出诊序号控制
              Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1), 操作员姓名 = 操作员姓名_In
              Where 记录id = 出诊记录id_In And 序号 = n_序号 And Nvl(挂号状态, 0) = 0;
            
              If Sql%RowCount = 0 Then
                Begin
                  Select 挂号状态 Into n_状态 From 临床出诊序号控制 Where 记录id = 出诊记录id_In And 序号 = n_序号;
                Exception
                  When Others Then
                    n_状态 := -1;
                End;
                If n_状态 = -1 Then
                  Insert Into 临床出诊序号控制
                    (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 锁号时间, 类型, 名称, 操作员姓名, 备注)
                    Select 出诊记录id_In, n_序号, 发生时间_In, 发生时间_In, 1, Decode(预约挂号_In, 1, 1, 0), Decode(预约挂号_In, 1, 2, 1),
                           Null, Null, Null, 操作员姓名_In, '追加号'
                    From Dual;
                Else
                  v_Err_Msg := '序号' || n_序号 || '已被使用,请重新选择一个序号.';
                  Raise Err_Item;
                End If;
              End If;
            End If;
          End If;
        Exception
          When Others Then
            v_Err_Msg := '序号' || n_序号 || '已被使用,请重新选择一个序号.';
            Raise Err_Item;
        End;
      End If;
    Else
      If 操作员姓名_In <> v_序号操作员 Or v_机器名 <> v_序号机器名 Then
        v_Err_Msg := '序号' || n_序号 || '已被自助机' || v_机器名 || '锁定,请重新选择一个序号.';
        Raise Err_Item;
      Else
        If n_预约顺序号 Is Null Then
          Update 临床出诊序号控制
          Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1)
          Where 记录id = 出诊记录id_In And 序号 = n_序号 And Nvl(挂号状态, 0) = 5 And 操作员姓名 = 操作员姓名_In And 工作站名称 = v_机器名;
        Else
          Update 临床出诊序号控制
          Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1)
          Where 记录id = 出诊记录id_In And 序号 = n_时段序号 And 预约顺序号 = n_预约顺序号 And Nvl(挂号状态, 0) = 5 And 操作员姓名 = 操作员姓名_In And
                工作站名称 = v_机器名;
        End If;
      End If;
    End If;
  End If;

  --产生病人挂号费用(可能单独是或包括病历费用)
  Select 病人费用记录_Id.Nextval Into n_费用id From Dual; --应该通过程序得到

  Insert Into 门诊费用记录
    (ID, 记录性质, 记录状态, 序号, 价格父号, 从属父号, NO, 实际票号, 门诊标志, 加班标志, 附加标志, 发药窗口, 病人id, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 病人科室id, 收费类别,
     计算单位, 收费细目id, 收入项目id, 收据费目, 付数, 数次, 标准单价, 应收金额, 实收金额, 结帐金额, 结帐id, 记帐费用, 开单部门id, 开单人, 划价人, 执行部门id, 执行人, 操作员编号, 操作员姓名,
     发生时间, 登记时间, 保险大类id, 保险项目否, 保险编码, 统筹金额, 摘要, 结论, 缴款组id)
  Values
    (n_费用id, 4, Decode(预约挂号_In, 1, 0, 1), 序号_In, Decode(价格父号_In, 0, Null, 价格父号_In), 从属父号_In, 单据号_In, 票据号_In, 1, 急诊_In,
     病历费_In, Decode(预约挂号_In, 1, To_Char(n_序号), 诊室_In), Decode(病人id_In, 0, Null, 病人id_In),
     Decode(门诊号_In, 0, Null, 门诊号_In), 付款方式_In, 姓名_In, Decode(姓名_In, Null, Null, 性别_In), Decode(姓名_In, Null, Null, 年龄_In),
     v_费别, 病人科室id_In, 收费类别_In, 号别_In, 收费细目id_In, 收入项目id_In, 收据费目_In, 1, 数次_In, 标准单价_In, 应收金额_In, 实收金额_In,
     Decode(预约挂号_In, 1, Null, Decode(Nvl(记帐费用_In, 0), 1, Null, 实收金额_In)),
     Decode(预约挂号_In, 1, Null, Decode(Nvl(记帐费用_In, 0), 1, Null, 结帐id_In)), Decode(Nvl(记帐费用_In, 0), 1, 1, 0), 开单部门id_In,
     操作员姓名_In, Decode(预约挂号_In, 1, 操作员姓名_In, Null), 执行部门id_In, 医生姓名_In, 操作员编号_In, 操作员姓名_In, 发生时间_In, 登记时间_In, 保险大类id_In,
     保险项目否_In, 保险编码_In, 统筹金额_In, Decode(收费单_In, Null, 摘要_In, '划价:' || 收费单_In), 预约方式_In, Decode(预约挂号_In, 1, Null, n_组id));

  --汇总结算到病人预交记录
  If Nvl(预约挂号_In, 0) = 0 And Nvl(记帐费用_In, 0) = 0 Then
    If Nvl(现金支付_In, 0) = 0 And Nvl(个帐支付_In, 0) = 0 And Nvl(预交支付_In, 0) = 0 And 序号_In = 1 Then
      Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      Insert Into 病人预交记录
        (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
         结算性质)
      Values
        (n_预交id, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), v_现金, 0, 登记时间_In, 操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费',
         n_组id, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In, 4);
    End If;
    If Nvl(现金支付_In, 0) <> 0 And 序号_In = 1 Then
      v_结算内容     := 结算方式_In || '|'; --以空格分开以|结尾,没有结算号码的
      v_结算方式记录 := '';
      While v_结算内容 Is Not Null Loop
        v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '|') - 1);
        v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
      
        v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1));
      
        v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        v_结算号码 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
      
        v_当前结算   := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        n_三方卡标志 := To_Number(v_当前结算);
      
        If Instr('|' || v_结算方式记录 || '|', '|' || Nvl(v_结算方式, v_现金) || '|') <> 0 Then
          v_Err_Msg := '使用了重复的结算方式,请检查!';
          Raise Err_Item;
        Else
          v_结算方式记录 := v_结算方式记录 || '|' || Nvl(v_结算方式, v_现金);
        End If;
      
        If n_三方卡标志 = 0 Then
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明,
             合作单位, 结算性质, 结算号码)
          Values
            (n_预交id, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(v_结算方式, v_现金), Nvl(n_结算金额, 0), 登记时间_In,
             操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费', n_组id, Null, Null, Null, Null, Null, 合作单位_In, 4, v_结算号码);
        Else
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明,
             合作单位, 结算性质, 结算号码)
          Values
            (n_预交id, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(v_结算方式, v_现金), Nvl(n_结算金额, 0), 登记时间_In,
             操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费', n_组id, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In, 4,
             v_结算号码);
        
          If Nvl(结算卡序号_In, 0) <> 0 And Nvl(现金支付_In, 0) <> 0 Then
            Zl_病人卡结算记录_支付(结算卡序号_In, 卡号_In, 0, Nvl(n_结算金额, 0), n_预交id, 操作员编号_In, 操作员姓名_In, 登记时间_In);
          End If;
        End If;
      
        If Nvl(更新交款余额_In, 1) = 1 Then
          Update 人员缴款余额
          Set 余额 = Nvl(余额, 0) + n_结算金额
          Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = Nvl(v_结算方式, v_现金)
          Returning 余额 Into n_返回值;
        
          If Sql%RowCount = 0 Then
            Insert Into 人员缴款余额
              (收款员, 结算方式, 性质, 余额)
            Values
              (操作员姓名_In, Nvl(v_结算方式, v_现金), 1, n_结算金额);
            n_返回值 := n_结算金额;
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 人员缴款余额
            Where 收款员 = 操作员姓名_In And 结算方式 = Nvl(v_结算方式, v_现金) And 性质 = 1 And Nvl(余额, 0) = 0;
          End If;
        End If;
      
        v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '|') + 1);
      End Loop;
    End If;
  
    --对于医保挂号
    If Nvl(个帐支付_In, 0) <> 0 And 序号_In = 1 Then
      Insert Into 病人预交记录
        (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 结算性质)
      Values
        (病人预交记录_Id.Nextval, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), v_个人帐户, 个帐支付_In, 登记时间_In, 操作员编号_In,
         操作员姓名_In, 结帐id_In, '医保挂号', n_组id, 4);
    End If;
  
    --对于就诊卡通过预交金挂号
    If Nvl(预交支付_In, 0) <> 0 And 序号_In = 1 Then
      n_预交金额 := 预交支付_In;
      For r_Deposit In c_Deposit(病人id_In, v_冲预交病人ids) Loop
        n_当前金额 := Case
                    When r_Deposit.金额 - n_预交金额 < 0 Then
                     r_Deposit.金额
                    Else
                     n_预交金额
                  End;
      
        If r_Deposit.结帐id = 0 Then
          --第一次冲预交(填上结帐ID,金额为0)
          Update 病人预交记录 Set 冲预交 = 0, 结帐id = 结帐id_In, 结算性质 = 4 Where ID = r_Deposit.原预交id;
        
        End If;
        --冲上次剩余额
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 预交类别, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,
           冲预交, 结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
          Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 预交类别, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号,
                 登记时间_In, 操作员姓名_In, 操作员编号_In, n_当前金额, 结帐id_In, n_组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
          From 病人预交记录
          Where NO = r_Deposit.No And 记录状态 = r_Deposit.记录状态 And 记录性质 In (1, 11) And Rownum = 1;
      
        --更新病人预交余额
        Update 病人余额
        Set 预交余额 = Nvl(预交余额, 0) - n_当前金额
        Where 病人id = r_Deposit.病人id And 性质 = 1 And 类型 = Nvl(1, 2);
        --检查是否已经处理完
        If r_Deposit.金额 < n_预交金额 Then
          n_预交金额 := n_预交金额 - r_Deposit.金额;
        Else
          n_预交金额 := 0;
        End If;
      
        If n_预交金额 = 0 Then
          Exit;
        End If;
      End Loop;
      If n_预交金额 > 0 Then
        v_Err_Msg := '预交余不够支付本次支付金额,不能继续操作！';
        Raise Err_Item;
      
      End If;
      Delete From 病人余额 Where 病人id = 病人id_In And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
    End If;
  
    --相关汇总表的处理
    --人员缴款余额
    If 序号_In = 1 And Nvl(更新交款余额_In, 1) = 1 Then
      If Nvl(个帐支付_In, 0) <> 0 Then
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) + 个帐支付_In
        Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = v_个人帐户
        Returning 余额 Into n_返回值;
      
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (操作员姓名_In, v_个人帐户, 1, 个帐支付_In);
          n_返回值 := 个帐支付_In;
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 人员缴款余额
          Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_个人帐户 And Nvl(余额, 0) = 0;
        End If;
      End If;
    End If;
  End If;

  --病人挂号汇总(只处理一次,且单独收取病历费不处理)
  If Nvl(预约挂号_In, 0) = 0 Then
    --病人本次就诊(以发生时间为准)
    If Nvl(病人id_In, 0) <> 0 And 序号_In = 1 Then
      Update 病人信息 Set 就诊时间 = 发生时间_In, 就诊状态 = 1, 就诊诊室 = 诊室_In Where 病人id = 病人id_In;
    End If;
  End If;

  If Nvl(预约挂号_In, 0) = 0 And Nvl(记帐费用_In, 0) = 1 Then
    --记帐
    If Nvl(病人id_In, 0) = 0 Then
      v_Err_Msg := '要针对病人的挂号费进行记帐，必须是建档病人才能记帐挂号。';
      Raise Err_Item;
    End If;
  
    --病人余额
    Update 病人余额
    Set 费用余额 = Nvl(费用余额, 0) + Nvl(实收金额_In, 0)
    Where 病人id = Nvl(病人id_In, 0) And 性质 = 1 And 类型 = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人余额 (病人id, 性质, 类型, 费用余额, 预交余额) Values (病人id_In, 1, 1, Nvl(实收金额_In, 0), 0);
    End If;
  
    --病人未结费用
    Update 病人未结费用
    Set 金额 = Nvl(金额, 0) + Nvl(实收金额_In, 0)
    Where 病人id = 病人id_In And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(病人科室id, 0) = Nvl(病人科室id_In, 0) And
          Nvl(开单部门id, 0) = Nvl(开单部门id_In, 0) And Nvl(执行部门id, 0) = Nvl(执行部门id_In, 0) And 收入项目id + 0 = 收入项目id_In And
          来源途径 + 0 = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人未结费用
        (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
      Values
        (病人id_In, Null, Null, 病人科室id_In, 开单部门id_In, 执行部门id_In, 收入项目id_In, 1, Nvl(实收金额_In, 0));
    End If;
  End If;

  --病人挂号记录
  If 号别_In Is Not Null And 序号_In = 1 Then
    --And Nvl(预约挂号_In, 0) = 0
    Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
    Begin
      Select 名称 Into v_付款方式 From 医疗付款方式 Where 编码 = 付款方式_In And Rownum < 2;
    Exception
      When Others Then
        v_付款方式 := Null;
    End;
    Insert Into 病人挂号记录
      (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 预约时间, 操作员编号,
       操作员姓名, 复诊, 号序, 社区, 预约, 预约方式, 摘要, 交易流水号, 交易说明, 合作单位, 接收时间, 接收人, 预约操作员, 预约操作员编号, 险类, 医疗付款方式, 出诊记录id, 收费单)
    Values
      (n_挂号id, 单据号_In, Decode(Nvl(预约挂号_In, 0), 1, 2, 1), 1, 病人id_In, 门诊号_In, 姓名_In, 性别_In, 年龄_In, 号别_In, 急诊_In, 诊室_In,
       Null, 执行部门id_In, 医生姓名_In, 0, Null, 登记时间_In, 发生时间_In, Decode(Nvl(预约挂号_In, 0), 1, 发生时间_In, Null), 操作员编号_In,
       操作员姓名_In, 复诊_In, n_序号, 社区_In, Decode(预约接收_In, 1, 1, 0), 预约方式_In, 摘要_In, 交易流水号_In, 交易说明_In, 合作单位_In,
       Decode(Nvl(预约挂号_In, 0), 0, 登记时间_In, Null), Decode(Nvl(预约挂号_In, 0), 0, 操作员姓名_In, Null),
       Decode(Nvl(预约挂号_In, 0), 1, 操作员姓名_In, Null), Decode(Nvl(预约挂号_In, 0), 1, 操作员编号_In, Null),
       Decode(Nvl(险类_In, 0), 0, Null, 险类_In), v_付款方式, 出诊记录id_In, 收费单_In);
  
    If Nvl(预约挂号_In, 0) = 0 And 预约方式_In Is Not Null Then
      Update 病人挂号记录
      Set 预约 = 1, 预约时间 = 发生时间_In, 预约操作员 = 操作员姓名_In, 预约操作员编号 = 操作员编号_In
      Where ID = n_挂号id;
    End If;
  
    n_预约生成队列 := 0;
    If Nvl(预约挂号_In, 0) = 1 Then
      n_预约生成队列 := Zl_To_Number(zl_GetSysParameter('预约生成队列', 1113));
    End If;
  
    --0-不产生队列;1-按医生或分诊台排队;2-先分诊,后医生站
    If Nvl(生成队列_In, 0) <> 0 And Nvl(预约挂号_In, 0) = 0 Or n_预约生成队列 = 1 Then
      n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113, 1, Nvl(执行部门id_In, 0)));
      If Nvl(n_分诊台签到排队, 0) = 0 Or n_预约生成队列 = 1 Then
        n_分时点显示 := Nvl(Zl_To_Number(zl_GetSysParameter(270)), 0);
        If Nvl(预约挂号_In, 0) = 1 And n_分时点显示 = 1 And n_分时段 = 1 Then
          n_分时点显示 := 1;
        Else
          n_分时点显示 := Null;
        End If;
        --产生队列
        --.按”执行部门” 的方式生成队列
        v_队列名称 := 执行部门id_In;
        v_排队号码 := Zlgetnextqueue(执行部门id_In, n_挂号id, 号别_In || '|' || n_序号);
        v_排队序号 := Zlgetsequencenum(0, n_挂号id, 0);
        d_排队时间 := Zl_Get_Queuedate(n_挂号id, 号别_In, n_序号, 发生时间_In);
        --  队列名称_In , 业务类型_In, 业务id_In,科室id_In,排队号码_In,排队标记_In,患者姓名_In,病人ID_IN, 诊室_In, 医生姓名_In, 排队时间_In
        Zl_排队叫号队列_Insert(v_队列名称, 0, n_挂号id, 执行部门id_In, v_排队号码, Null, 姓名_In, 病人id_In, 诊室_In, 医生姓名_In, d_排队时间, 预约方式_In,
                         n_分时点显示, v_排队序号);
      
        --挂号立即排队
        If Nvl(n_分诊台签到排队, 0) = 0 Then
          Update 病人挂号记录 Set 记录标志 = 1 Where ID = n_挂号id;
        End If;
      End If;
    End If;
  End If;
  --病人担保信息
  If 病人id_In Is Not Null And 序号_In = 1 Then
    --取参数:
    If Nvl(n_结算模式, 0) <> Nvl(结算模式_In, 0) Then
      --结算模式的确定
    
      v_Err_Msg := Null;
      Begin
        Select Nvl(结算模式, 0) Into n_结算模式 From 病人信息 Where 病人id = 病人id_In;
      Exception
        When Others Then
          v_Err_Msg := '未找到指定的病人信息,不允许挂号';
      End;
    
      If v_Err_Msg Is Not Null Then
        Raise Err_Item;
      End If;
      If n_结算模式 = 1 And Nvl(结算模式_In, 0) = 0 Then
        --病人已经是"先诊疗后结算的",本次是"先结算后诊疗的",则检查是否存在未结数据
        Select Count(1)
        Into n_Count
        From 病人未结费用
        Where 病人id = 病人id_In And (来源途径 In (1, 4) Or 来源途径 = 3 And Nvl(主页id, 0) = 0) And Nvl(金额, 0) <> 0 And Rownum < 2;
        If Nvl(n_Count, 0) <> 0 Then
          --存在未结算数据，必须先结算后才允许执行
          v_Err_Msg := '当前病人的就诊模式为先诊疗后结算且存在未结费用，不允许调整该病人的就诊模式,你可以先对未结费用结帐后再挂号或不调整病人的就诊模式!';
          Raise Err_Item;
        End If;
        --检查
        --未发生医嘱业务的（即当时就挂号的,需要保证同一次的就诊模式是一至的(程序已经检查，不用再处理)
      End If;
      Update 病人信息 Set 结算模式 = 结算模式_In Where 病人id = 病人id_In;
    End If;
  
    Update 病人信息
    Set 担保人 = Null, 担保额 = Null, 担保性质 = Null
    Where 病人id = 病人id_In And Exists (Select 1
           From 病人担保记录
           Where 病人id = 病人id_In And 主页id Is Not Null And
                 登记时间 = (Select Max(登记时间) From 病人担保记录 Where 病人id = 病人id_In));
  
    If Sql%RowCount > 0 Then
      Update 病人担保记录
      Set 到期时间 = Sysdate
      Where 病人id = 病人id_In And 主页id Is Not Null And Nvl(到期时间, Sysdate) > Sysdate;
    End If;
  End If;
  If 序号_In = 1 Then
    --消息推送
    Begin
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 1, n_挂号id;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_001(n_挂号id, 单据号_In);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人挂号记录_出诊_Insert;
/

--127075:焦博,2018-07-06,调整使用到了部门参数分诊台签到排队的Oracle函数
Create Or Replace Procedure Zl_病人挂号记录_出诊_Delete
(
  单据号_In       门诊费用记录.No%Type,
  操作员编号_In   门诊费用记录.操作员编号%Type,
  操作员姓名_In   门诊费用记录.操作员姓名%Type,
  摘要_In         门诊费用记录.摘要%Type := Null, --预约取消时 填写 存放预约取消原因
  删除门诊号_In   Number := 0,
  非原样退结算_In Varchar2 := Null,
  退费类型_In     In Number := 0, --0-全退 1-退挂号费 2-退病历费
  退指定结算_In   病人预交记录.结算方式%Type := Null,
  退号重用_In     Number := 1,
  结算方式_In     Varchar2 := Null,
  退预交_In       病人预交记录.冲预交%Type := Null,
  收回票据号_In   Varchar2 := Null,
  交易说明_In     病人预交记录.交易说明%Type := Null
) As
  --退费类型_In,在一下几种情况下不准进行部分退费
  --    2.三方接口,暂时不支持
  -- 挂号费病历费分开退,规则
  --    普通结算方式:原结算方式退部分费用
  --    预交款:预交款,退部分
  --    预交款与普通结算方式混合:退款按照普通结算方式部分退
  --    消费卡:原样将费用部分退入消费卡
  --非原样退结算_In:指不能退还给原样结算方式(如医保的个人账户,三方账户的退现等),多个用逗分离
  --退指定结算_IN:指非原样退结算部分,应该退给哪种结算方式,为空时缺省退给现金,否则退给指定的结算方式

  --该游标用于判断是否单独收病历费,及挂号汇总表处理
  Cursor c_Registinfo(v_状态 门诊费用记录.记录状态%Type) Is
    Select a.发生时间, a.登记时间, c.接收时间, a.收费细目id As 项目id, c.执行部门id As 科室id, c.执行人 As 医生姓名, d.Id As 医生id, c.号别 As 号码
    From 门诊费用记录 A, 病人挂号记录 C, 人员表 D
    Where a.记录性质 = 4 And a.No = 单据号_In And a.No = c.No And a.记录状态 = v_状态 And c.执行人 = d.姓名(+) And Rownum < 2;
  r_Registrow c_Registinfo%RowType;

  --该游标用于判断记录是否存在,及费用汇总表处理
  Cursor c_Moneyinfo(v_状态 门诊费用记录.记录状态%Type) Is
    Select 病人科室id, 开单部门id, 执行部门id, 收入项目id, Nvl(Sum(应收金额), 0) As 应收, Nvl(Sum(实收金额), 0) As 实收, Nvl(Sum(结帐金额), 0) As 结帐
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 = v_状态 And NO = 单据号_In
    Group By 病人科室id, 开单部门id, 执行部门id, 收入项目id;
  r_Moneyrow c_Moneyinfo%RowType;

  --该光标用于处理人员缴款余额中退的不同结算方式的金额
  Cursor c_Opermoney(n_Id 病人预交记录.结帐id%Type) Is
    Select Distinct b.结算方式, -1 * Nvl(b.冲预交, 0) As 冲预交
    From 病人预交记录 B
    Where b.结帐id = n_Id And b.记录性质 = 4 And b.记录状态 = 2 And Nvl(b.冲预交, 0) <> 0;

  v_Err_Msg Varchar(255);
  Err_Item Exception;

  n_结帐id 病人预交记录.结帐id%Type;
  n_销帐id 门诊费用记录.结帐id%Type;

  v_退指定结算方式 病人预交记录.结算方式%Type;
  n_退款金额       病人预交记录.冲预交%Type;
  n_打印id         票据打印内容.Id%Type;
  n_病人id         病人信息.病人id%Type;
  n_退费金额       病人预交记录.冲预交%Type;
  n_预交金额       病人预交记录.冲预交%Type; --原记录 预交缴款金额
  n_返回值         病人余额.预交余额%Type;
  n_挂号id         病人挂号记录.Id%Type;
  n_组id           财务缴款分组.Id%Type;

  n_二次退费       Number; --记录是否是此单据的第二次退费
  n_分诊台签到排队 Number;
  n_预约生成队列   Number;
  n_预约挂号       Number;
  n_挂号生成队列   Number;
  d_Date           Date;
  n_记帐           门诊费用记录.记帐费用%Type;
  n_病人id1        病人信息.病人id%Type;
  n_返回额         门诊费用记录.实收金额%Type;
  n_已结帐         Number;
  n_序号           病人挂号记录.号序%Type;
  n_就诊病人id     病人信息.病人id%Type;
  d_就诊时间       就诊登记记录.就诊时间%Type;
  n_出诊记录id     临床出诊记录.Id%Type;
  v_结算内容       Varchar2(5000);
  v_当前结算       Varchar2(1000);
  v_附加ids        Varchar2(500);
  v_Temp           Varchar2(500);
  v_结算方式       病人预交记录.结算方式%Type;
  n_三方卡标志     Number;
  n_结算金额       病人预交记录.冲预交%Type;
  n_检查数         Number;
Begin
  n_组id           := Zl_Get组id(操作员姓名_In);
  v_退指定结算方式 := 退指定结算_In;

  Select 出诊记录id, 号序 Into n_出诊记录id, n_序号 From 病人挂号记录 Where NO = 单据号_In And Rownum < 2;

  --首先判断要退号/取消预约的记录是否存在
  Open c_Moneyinfo(1);
  Fetch c_Moneyinfo
    Into r_Moneyrow;
  If c_Moneyinfo%RowCount = 0 Then
    Close c_Moneyinfo;
    Open c_Moneyinfo(0);
    Fetch c_Moneyinfo
      Into r_Moneyrow;
    If c_Moneyinfo%RowCount = 0 Then
      v_Err_Msg := '要处理的单据不存在。';
      Raise Err_Item;
    End If;
    n_预约挂号 := 1;
  End If;
  Close c_Moneyinfo;

  Begin
    Select Zl_Fun_Regcustomname Into v_Temp From Dual;
    If v_Temp Is Not Null Then
      v_附加ids := Substr(v_Temp, Instr(v_Temp, '|') + 1);
    End If;
  Exception
    When Others Then
      v_附加ids := Null;
  End;

  --1.预约处理
  If Nvl(n_预约挂号, 0) = 1 Then
    --减少已约数
    Open c_Registinfo(0);
    Fetch c_Registinfo
      Into r_Registrow;
  
    n_检查数 := Null;
    Update 临床出诊记录 Set 已约数 = Nvl(已约数, 0) - 1 Where ID = n_出诊记录id Returning 已约数 Into n_检查数;
    If Nvl(n_检查数, 0) < 0 Then
      Update 临床出诊记录 Set 已约数 = 0 Where ID = n_出诊记录id;
    End If;
  
    Update 病人挂号汇总
    Set 已约数 = Nvl(已约数, 0) - 1
    Where 日期 = Trunc(r_Registrow.发生时间) And Nvl(医生id, 0) = Nvl(r_Registrow.医生id, 0) And
          Nvl(医生姓名, '-') = Nvl(r_Registrow.医生姓名, '-') And 科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And
          (号码 = r_Registrow.号码 Or 号码 Is Null);
  
    If Sql%RowCount = 0 Then
      Insert Into 病人挂号汇总
        (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已约数)
      Values
        (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
         Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, -1);
    End If;
  
    Close c_Registinfo;
  
    --更新挂号序号状态
    Update 临床出诊序号控制
    Set 挂号状态 = 0, 操作员姓名 = Null
    Where 挂号状态 = 2 And 记录id = n_出诊记录id And 序号 = n_序号;
  
    Update 临床出诊序号控制
    Set 挂号状态 = 4, 操作员姓名 = Null
    Where 挂号状态 = 2 And 记录id = n_出诊记录id And 备注 = To_Char(n_序号);
  
    --添加病人挂号记录的 冲销记录
    Select 病人挂号记录_Id.Nextval, Sysdate Into n_挂号id, d_Date From Dual;
  
    Update 病人挂号记录 Set 记录状态 = 3 Where NO = 单据号_In And 记录状态 = 1 And 记录性质 = 2;
    If Sql%NotFound Then
      v_Err_Msg := '预约单【' || 单据号_In || '】不存在或由于并发原因已经被取消预约';
      Raise Err_Item;
    End If;
  
    Insert Into 病人挂号记录
      (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 操作员编号, 操作员姓名,
       复诊, 号序, 社区, 预约, 摘要, 预约方式, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 医疗付款方式, 出诊记录id, 预约操作员, 预约操作员编号)
      Select n_挂号id, NO, 记录性质, 2, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, d_Date, 发生时间,
             操作员编号_In, 操作员姓名_In, 复诊, 号序, 社区, 预约, Nvl(摘要_In, 摘要) As 摘要, 预约方式, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 医疗付款方式,
             n_出诊记录id, 预约操作员, 预约操作员编号
      From 病人挂号记录
      Where NO = 单据号_In;
  
    Update 门诊费用记录 Set 记录状态 = 3 Where NO = 单据号_In And 记录性质 = 4 And 记录状态 = 0;
    Insert Into 门诊费用记录
      (ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 记帐单id, 病人id, 医嘱序号, 门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号, 付款方式, 病人科室id, 费别, 收费类别,
       收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id,
       执行人, 执行状态, 执行时间, 结论, 操作员编号, 操作员姓名, 结帐id, 结帐金额, 保险大类id, 保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要, 是否急诊, 缴款组id, 费用状态, 待转出,
       挂号id, 主页id)
      Select 病人费用记录_Id.Nextval, 记录性质, NO, 实际票号, 2, 序号, 从属父号, 价格父号, 记帐单id, 病人id, 医嘱序号, 门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号, 付款方式,
             病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, -1 * 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, -1 * 应收金额,
             -1 * 实收金额, 划价人, 开单部门id, 开单人, 发生时间, d_Date, 执行部门id, 执行人, -1, 执行时间, 结论, 操作员编号_In, 操作员姓名_In, Null, Null,
             保险大类id, 保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要, 是否急诊, 缴款组id, 费用状态, 待转出, 挂号id, 主页id
      From 门诊费用记录
      Where NO = 单据号_In And 记录性质 = 4 And 记录状态 = 3;
  
    --如果预约生成队列时需要清除队列
  
    n_预约生成队列 := Zl_To_Number(zl_GetSysParameter('预约生成队列', 1113));
    If Nvl(n_预约生成队列, 0) = 1 Then
      --要删除队列
      For v_挂号 In (Select ID, 诊室, 执行部门id, 执行人 From 病人挂号记录 Where NO = 单据号_In) Loop
        Zl_排队叫号队列_Delete(v_挂号.执行部门id, v_挂号.Id);
      End Loop;
    End If;
    Return;
  End If;

  Select Nvl(记帐费用, 0), 病人id, Decode(Sign(Nvl(结帐id, 0)), 0, 0, 1)
  Into n_记帐, n_病人id, n_已结帐
  From 门诊费用记录
  Where 记录性质 = 4 And NO = 单据号_In And 记录状态 In (1, 3) And Rownum < 2;

  --2.挂号处理
  n_已结帐 := Nvl(n_已结帐, 0);

  If n_已结帐 = 1 And n_记帐 = 1 Then
    Select Sysdate, Null Into d_Date, n_销帐id From Dual;
  Else
    Select Sysdate, 病人结帐记录_Id.Nextval Into d_Date, n_销帐id From Dual;
  End If;

  ----0-全退 1-退挂号费 2-退病历费
  If Nvl(退费类型_In, 0) <> 2 Then
    --不是光退病历费时处理
    --更新挂号序号状态
    If 退号重用_In = 1 Then
      Update 临床出诊序号控制
      Set 挂号状态 = 0, 操作员姓名 = Null
      Where 挂号状态 = 1 And 记录id = n_出诊记录id And 序号 = n_序号;
    
      Update 临床出诊序号控制
      Set 挂号状态 = 4, 操作员姓名 = Null
      Where 挂号状态 = 1 And 记录id = n_出诊记录id And 备注 = To_Char(n_序号);
    Else
      Update 临床出诊序号控制
      Set 挂号状态 = 4, 操作员姓名 = 操作员姓名_In
      Where 挂号状态 = 1 And 记录id = n_出诊记录id And (序号 = n_序号 Or 备注 = To_Char(n_序号));
    End If;
  
    --病人就诊状态
    If n_病人id Is Not Null Then
      Update 病人信息 Set 就诊状态 = 0, 就诊诊室 = Null Where 病人id = n_病人id;
    
      --删除门诊号相关处理,只有当只有一条挂号记录并且病人建档日期与挂号日期近似时才会处理
      If 删除门诊号_In = 1 Then
        Delete 门诊病案记录 Where 病人id = n_病人id;
        Update 病人信息 Set 门诊号 = Null Where 病人id = n_病人id;
        --费用记录包括挂号及病案、就诊卡费用,以及病人交费后退费或销帐的费用,挂号记录在最后处理
        Update 门诊费用记录 Set 标识号 = Null Where 门诊标志 = 1 And 病人id = n_病人id;
      End If;
    End If;
  
    --如果挂时收了就诊卡费,退费时清除就诊卡号,在非光退病历费时
    n_病人id1 := Null;
    Begin
      Select 病人id
      Into n_病人id1
      From 门诊费用记录
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And 附加标志 = 2 And Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
  
    If n_病人id1 Is Not Null And Nvl(退费类型_In, 0) <> 2 Then
      Update 病人信息
      Set 就诊卡号 = Null, 卡验证码 = Null, Ic卡号 = Decode(Ic卡号, 就诊卡号, Null, Ic卡号)
      Where 病人id = n_病人id1;
    End If;
  
  End If;

  --检查前面是否已经部分退过费用
  Begin
    Select 1 Into n_二次退费 From 门诊费用记录 Where 记录性质 = 4 And NO = 单据号_In And 记录状态 = 3 And Rownum < 2;
  Exception
    When Others Then
      n_二次退费 := 0;
  End;

  --门诊费用记录
  --冲销记录
  If Nvl(退费类型_In, 0) = 0 Or Nvl(退费类型_In, 0) = 2 Then
    --全退,退病历费
    --门诊费用记录，冲销记录
    Insert Into 门诊费用记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位, 付数,
       数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 执行人, 操作员编号, 操作员姓名, 发生时间, 登记时间,
       结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 缴款组id)
      Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别,
             收费细目id, 计算单位, 付数, -数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, -应收金额, -实收金额, 开单部门id, 开单人, 执行部门id, 执行人,
             操作员编号_In, 操作员姓名_In, 发生时间, d_Date, n_销帐id,
             Decode(n_记帐, 1, Decode(Nvl(n_已结帐, 0), 0, -1 * 实收金额, Null), -1 * 结帐金额), 保险项目否, 保险大类id, -1 * 统筹金额,
             Nvl(摘要_In, 摘要) As 摘要, Decode(Nvl(附加标志, 0), 9, 1, 0), 保险编码, 费用类型, n_组id
      From 门诊费用记录
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
  
    --原始记录
    If n_记帐 = 1 And Nvl(n_已结帐, 0) = 0 Then
      Update 门诊费用记录
      Set 记录状态 = 3, 结帐id = n_销帐id, 结帐金额 = 实收金额
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
    Else
      Update 门诊费用记录
      Set 记录状态 = 3
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
    End If;
  Elsif Nvl(退费类型_In, 0) = 1 Or Nvl(退费类型_In, 0) = 4 Then
    --退挂号费,退挂号与病历费
    --门诊费用记录，冲销记录
    Insert Into 门诊费用记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位, 付数,
       数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 执行人, 操作员编号, 操作员姓名, 发生时间, 登记时间,
       结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 缴款组id)
      Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别,
             收费细目id, 计算单位, 付数, -数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, -应收金额, -实收金额, 开单部门id, 开单人, 执行部门id, 执行人,
             操作员编号_In, 操作员姓名_In, 发生时间, d_Date, n_销帐id,
             Decode(n_记帐, 1, Decode(Nvl(n_已结帐, 0), 0, -1 * 实收金额, Null), -1 * 结帐金额), 保险项目否, 保险大类id, -1 * 统筹金额,
             Nvl(摘要_In, 摘要) As 摘要, Decode(Nvl(附加标志, 0), 9, 1, 0), 保险编码, 费用类型, n_组id
      From 门诊费用记录
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Instr(',' || v_附加ids || ',', ',' || 收费细目id || ',') = 0 And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
  
    --原始记录
    If n_记帐 = 1 And Nvl(n_已结帐, 0) = 0 Then
      Update 门诊费用记录
      Set 记录状态 = 3, 结帐id = n_销帐id, 结帐金额 = 实收金额
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Instr(',' || v_附加ids || ',', ',' || 收费细目id || ',') = 0 And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
    Else
      Update 门诊费用记录
      Set 记录状态 = 3
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Instr(',' || v_附加ids || ',', ',' || 收费细目id || ',') = 0 And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
    End If;
  Elsif Nvl(退费类型_In, 0) = 3 Then
    --退附加费
    --门诊费用记录，冲销记录
    Insert Into 门诊费用记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位, 付数,
       数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 执行人, 操作员编号, 操作员姓名, 发生时间, 登记时间,
       结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 缴款组id)
      Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别,
             收费细目id, 计算单位, 付数, -数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, -应收金额, -实收金额, 开单部门id, 开单人, 执行部门id, 执行人,
             操作员编号_In, 操作员姓名_In, 发生时间, d_Date, n_销帐id,
             Decode(n_记帐, 1, Decode(Nvl(n_已结帐, 0), 0, -1 * 实收金额, Null), -1 * 结帐金额), 保险项目否, 保险大类id, -1 * 统筹金额,
             Nvl(摘要_In, 摘要) As 摘要, Decode(Nvl(附加标志, 0), 9, 1, 0), 保险编码, 费用类型, n_组id
      From 门诊费用记录
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Instr(',' || v_附加ids || ',', ',' || 收费细目id || ',') > 0 And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
  
    --原始记录
    If n_记帐 = 1 And Nvl(n_已结帐, 0) = 0 Then
      Update 门诊费用记录
      Set 记录状态 = 3, 结帐id = n_销帐id, 结帐金额 = 实收金额
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Instr(',' || v_附加ids || ',', ',' || 收费细目id || ',') > 0 And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
    Else
      Update 门诊费用记录
      Set 记录状态 = 3
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Instr(',' || v_附加ids || ',', ',' || 收费细目id || ',') > 0 And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
    End If;
  Elsif Nvl(退费类型_In, 0) = 5 Then
    --退挂号与附加费
    --门诊费用记录，冲销记录
    Insert Into 门诊费用记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位, 付数,
       数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 执行人, 操作员编号, 操作员姓名, 发生时间, 登记时间,
       结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 缴款组id)
      Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别,
             收费细目id, 计算单位, 付数, -数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, -应收金额, -实收金额, 开单部门id, 开单人, 执行部门id, 执行人,
             操作员编号_In, 操作员姓名_In, 发生时间, d_Date, n_销帐id,
             Decode(n_记帐, 1, Decode(Nvl(n_已结帐, 0), 0, -1 * 实收金额, Null), -1 * 结帐金额), 保险项目否, 保险大类id, -1 * 统筹金额,
             Nvl(摘要_In, 摘要) As 摘要, Decode(Nvl(附加标志, 0), 9, 1, 0), 保险编码, 费用类型, n_组id
      From 门诊费用记录
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Nvl(附加标志, 0) <> 1;
  
    --原始记录
    If n_记帐 = 1 And Nvl(n_已结帐, 0) = 0 Then
      Update 门诊费用记录
      Set 记录状态 = 3, 结帐id = n_销帐id, 结帐金额 = 实收金额
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Nvl(附加标志, 0) <> 1;
    Else
      Update 门诊费用记录
      Set 记录状态 = 3
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Nvl(附加标志, 0) <> 1;
    End If;
  End If;

  n_结帐id := 0;
  If n_记帐 = 0 Then
    --获取结帐ID
    Select Nvl(结帐id, 0)
    Into n_结帐id
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 = 3 And NO = 单据号_In And
          Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0)) And
          Rownum = 1;
  End If;

  If n_记帐 = 1 Then
    --记帐
    For c_费用 In (Select 实收金额, 病人科室id, 开单部门id, 执行部门id, 收入项目id
                 From 门诊费用记录
                 Where 记录性质 = 4 And 记录状态 = 3 And NO = 单据号_In And
                       Nvl(附加标志, 0) =
                       Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0)) And
                       Nvl(记帐费用, 0) = 1) Loop
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
            Nvl(开单部门id, 0) = Nvl(c_费用.开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(c_费用.执行部门id, 0) And 收入项目id + 0 = c_费用.收入项目id And
            来源途径 + 0 = 1;
    
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
    --1.退费
    --病人挂号结算:现金和个人帐户部份
    If 结算方式_In Is Null Then
      If 非原样退结算_In Is Not Null Then
        --退款金额获取
        If Nvl(退费类型_In, 0) <> 0 Or Nvl(n_二次退费, 0) <> 0 Then
          --如果是单独退病历费,或者只退挂号费,先获取退费金额
          Begin
            --获取本次退款金额
            Select Sum(Nvl(实收金额, 0)) As 收款金额
            Into n_退款金额
            From 门诊费用记录
            Where NO = 单据号_In And 记录性质 = 4 And 记录状态 = 3 And
                  Nvl(附加标志, 0) =
                  Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
          
          Exception
            When Others Then
              v_Err_Msg := '单据【' || 单据号_In || '】的' || Case Nvl(退费类型_In, 0)
                             When 1 Then
                              '挂号费用'
                             When 2 Then
                              '病历费'
                           End || '可能由于并发原因已经进行了退费或者单据不存在!';
              Raise Err_Item;
          End;
          Begin
            Select 冲预交
            Into n_退费金额
            From 病人预交记录
            Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id And Instr(',' || 非原样退结算_In || ',', ',' || 结算方式 || ',') = 0;
          Exception
            When Others Then
              n_退费金额 := 0;
          End;
        
          --a.允许的结算方式
        
          Insert Into 病人预交记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
             结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
            Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In, -n_退款金额,
                   n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
            From 病人预交记录
            Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id And Instr(',' || 非原样退结算_In || ',', ',' || 结算方式 || ',') = 0;
        
          If n_退费金额 = 0 Then
            --b.不允许的退现金
            If n_退款金额 <> 0 Then
              If v_退指定结算方式 Is Null Then
                --退给现金
                Begin
                  Select 名称 Into v_退指定结算方式 From 结算方式 Where 性质 = 1;
                Exception
                  When Others Then
                    v_退指定结算方式 := '现金';
                End;
              End If;
              Update 病人预交记录
              Set 冲预交 = 冲预交 + (-1 * n_退款金额), 交易说明 = Nvl(交易说明_In, 交易说明)
              Where 记录性质 = 4 And 记录状态 = 2 And 结帐id = n_销帐id And 结算方式 = v_退指定结算方式;
              If Sql%RowCount = 0 Then
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别,
                   卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
                  Select 病人预交记录_Id.Nextval, a.No, a.实际票号, a.记录性质, 2, a.病人id, a.主页id, a.科室id, a.摘要, v_退指定结算方式, d_Date,
                         操作员编号_In, 操作员姓名_In, -1 * n_退款金额, n_销帐id, n_组id, 预交类别, Decode(交易说明_In, Null, 卡类别id, Null), 结算卡序号,
                         Decode(交易说明_In, Null, 卡号, Null), Decode(交易说明_In, Null, 交易流水号, Null), Nvl(交易说明_In, 交易说明), 合作单位,
                         4
                  From 病人预交记录 A
                  Where a.记录性质 = 4 And a.记录状态 = 1 And a.结帐id = n_结帐id And Rownum < 2;
              End If;
            End If;
          End If;
        Else
          --a.允许的结算方式原样退
          Insert Into 病人预交记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
             结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
            Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In, -冲预交,
                   n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
            From 病人预交记录
            Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id And Instr(',' || 非原样退结算_In || ',', ',' || 结算方式 || ',') = 0;
        
          --b.不允许的退现金
          Begin
            Select Sum(冲预交)
            Into n_退费金额
            From 病人预交记录
            Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id And Instr(',' || 非原样退结算_In || ',', ',' || 结算方式 || ',') > 0;
          Exception
            When Others Then
              n_退费金额 := 0;
          End;
          If n_退费金额 <> 0 Then
            If v_退指定结算方式 Is Null Then
              --退给现金
              Begin
                Select 结算方式
                Into v_退指定结算方式
                From 病人预交记录
                Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id And
                      Instr(',' || 非原样退结算_In || ',', ',' || 结算方式 || ',') = 0;
              
              Exception
                When Others Then
                  Begin
                    Select 名称 Into v_退指定结算方式 From 结算方式 Where 性质 = 1;
                  Exception
                    When Others Then
                      v_退指定结算方式 := '现金';
                  End;
              End;
            End If;
            Update 病人预交记录
            Set 冲预交 = 冲预交 + (-1 * n_退费金额), 交易说明 = Nvl(交易说明_In, 交易说明)
            Where 记录性质 = 4 And 记录状态 = 2 And 结帐id = n_销帐id And 结算方式 = v_退指定结算方式;
            If Sql%RowCount = 0 Then
              Insert Into 病人预交记录
                (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别,
                 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
                Select 病人预交记录_Id.Nextval, a.No, a.实际票号, a.记录性质, 2, a.病人id, a.主页id, a.科室id, a.摘要, v_退指定结算方式, d_Date,
                       操作员编号_In, 操作员姓名_In, -1 * n_退费金额, n_销帐id, n_组id, 预交类别, Decode(交易说明_In, Null, 卡类别id, Null), 结算卡序号,
                       Decode(交易说明_In, Null, 卡号, Null), Decode(交易说明_In, Null, 交易流水号, Null), Nvl(交易说明_In, 交易说明), 合作单位, 4
                From 病人预交记录 A
                Where a.记录性质 = 4 And a.记录状态 = 1 And a.结帐id = n_结帐id And Rownum < 2;
            End If;
          End If;
        End If;
      Else
        --退款金额获取
        If Nvl(退费类型_In, 0) <> 0 Or Nvl(n_二次退费, 0) <> 0 Then
          --如果是单独退病历费,或者只退挂号费,先获取退费金额
          Begin
            --获取本次退款金额
            Select Sum(Nvl(实收金额, 0)) As 收款金额
            Into n_退款金额
            From 门诊费用记录
            Where NO = 单据号_In And 记录性质 = 4 And 记录状态 = 3 And
                  Nvl(附加标志, 0) =
                  Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
          Exception
            When Others Then
              v_Err_Msg := '单据【' || 单据号_In || '】的' || Case Nvl(退费类型_In, 0)
                             When 1 Then
                              '挂号费用'
                             When 2 Then
                              '病历费'
                           End || '可能由于并发原因已经进行了退费或者单据不存在!';
              Raise Err_Item;
          End;
        End If;
        If Nvl(n_二次退费, 0) = 0 And Nvl(退费类型_In, 0) = 0 Then
          --首次全退
          Insert Into 病人预交记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
             结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
            Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In,
                   -1 * 冲预交, n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
            From 病人预交记录
            Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id;
        Else
          --二次退费,或者本次单退一部分
          --二次退费时,记录状态=3 ,首次部分退,记录状态为1
          Insert Into 病人预交记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
             结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
            Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In,
                   -1 * n_退款金额, n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
            From 病人预交记录
            Where 记录性质 = 4 And 记录状态 = Decode(Nvl(n_二次退费, 0), 0, 1, 3) And 结帐id = n_结帐id And 摘要 = '医保挂号' And
                  冲预交 = n_退款金额 And Rownum < 2;
          If Sql%RowCount = 0 Then
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
               结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
              Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In,
                     -1 * n_退款金额, n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
              From 病人预交记录
              Where 记录性质 = 4 And 记录状态 = Decode(Nvl(n_二次退费, 0), 0, 1, 3) And 结帐id = n_结帐id And 冲预交 = n_退款金额 And
                    Rownum < 2;
          End If;
          If Sql%RowCount = 0 Then
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
               结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
              Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In,
                     -1 * n_退款金额, n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
              From 病人预交记录
              Where 记录性质 = 4 And 记录状态 = Decode(Nvl(n_二次退费, 0), 0, 1, 3) And 结帐id = n_结帐id And Rownum < 2;
            If Sql%RowCount = 0 Then
              --部分退费,并且全部使用预交款缴费时才存在此种情况
              n_预交金额 := n_退款金额;
            End If;
          End If;
        
        End If;
      End If;
    Else
      --按结算方式退
      v_结算内容 := 结算方式_In || '|'; --以空格分开以|结尾,没有结算号码的
      While v_结算内容 Is Not Null Loop
        v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '|') - 1);
        v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
      
        v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1));
      
        v_当前结算   := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        n_三方卡标志 := To_Number(v_当前结算);
      
        If n_三方卡标志 = 0 Then
          Insert Into 病人预交记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
             结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质, 结算号码)
            Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, v_结算方式, d_Date, 操作员编号_In, 操作员姓名_In,
                   -1 * n_结算金额, n_销帐id, n_组id, 预交类别, Null, Null, Null, Null, 交易说明_In, 合作单位, 4, 结算号码
            From 病人预交记录
            Where 记录性质 = 4 And 记录状态 = Decode(Nvl(n_二次退费, 0), 0, 1, 3) And 结帐id = n_结帐id And Rownum < 2;
        Else
          Insert Into 病人预交记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
             结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质, 结算号码)
            Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, v_结算方式, d_Date, 操作员编号_In, 操作员姓名_In,
                   -1 * n_结算金额, n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, Nvl(交易说明_In, 交易说明), 合作单位, 4, 结算号码
            
            From 病人预交记录
            Where 记录性质 = 4 And 记录状态 = Decode(Nvl(n_二次退费, 0), 0, 1, 3) And 结帐id = n_结帐id And
                  (卡类别id Is Not Null Or 结算卡序号 Is Not Null) And Rownum < 2;
        End If;
        v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '|') + 1);
      End Loop;
      n_预交金额 := Nvl(退预交_In, 0);
    End If;
    --首次退费时,记录状态便调整为了3
    Update 病人预交记录
    Set 记录状态 = 3
    Where 记录性质 = 4 And 记录状态 = Decode(Nvl(n_二次退费, 0), 0, 1, 3) And 结帐id = n_结帐id;
  
    --冲预交 1-全退 2-部分退,部分退时当全部使用预交进行缴款
    If Nvl(退费类型_In, 0) = 0 Or (Nvl(退费类型_In, 0) <> 0 And n_预交金额 <> 0) Then
      --病人挂号结算:冲预交款部份
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
         结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
        Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, d_Date,
               操作员姓名_In, 操作员编号_In, -1 * Decode(Nvl(退费类型_In, 0) + Nvl(n_二次退费, 0), 0, 冲预交, n_预交金额), n_销帐id, n_组id, 预交类别,
               卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
        From 病人预交记录
        Where 记录性质 In (1, 11) And 结帐id = n_结帐id And Nvl(冲预交, 0) <> 0 And
              Rownum = Decode(Nvl(退费类型_In, 0) + Nvl(n_二次退费, 0), 0, Rownum, 1);
    End If;
  
    --处理病人预交余额
    For c_预交 In (Select 病人id, 预交类别, -1 * Sum(Nvl(冲预交, 0)) As 冲预交
                 From 病人预交记录
                 Where 记录性质 In (1, 11) And 结帐id = n_销帐id
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
  
    If 收回票据号_In Is Not Null Then
      --光退挂号费,不回收票据
      --退卡收回票据(可能上次挂号使用票据,不能收回)
      Begin
        --从最后一次打印的内容中取
        Select ID
        Into n_打印id
        From (Select b.Id
               From 票据使用明细 A, 票据打印内容 B
               Where a.打印id = b.Id And a.性质 = 1 And a.原因 In (1, 3) And b.数据性质 = 4 And b.No = 单据号_In
               Order By a.使用时间 Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          n_打印id := Null;
      End;
    
      --先收回原票据
      If n_打印id Is Not Null Then
        Begin
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
            Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_Date, 操作员姓名_In, 票据金额
            From 票据使用明细
            Where 打印id = n_打印id And 性质 = 1 And Instr(',' || 收回票据号_In || ',', ',' || 号码 || ',') > 0;
        Exception
          When Others Then
            Delete From 票据使用明细 Where 打印id = n_打印id And 性质 = 2 And 原因 = 2;
            Insert Into 票据使用明细
              (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
              Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_Date, 操作员姓名_In, 票据金额
              From 票据使用明细
              Where 打印id = n_打印id And 性质 = 1 And Instr(',' || 收回票据号_In || ',', ',' || 号码 || ',') > 0;
        End;
      End If;
    End If;
  End If;

  --单独退病历费用,不处理汇总记录
  --相关汇总表的处理

  --病人挂号汇总
  Open c_Registinfo(3);
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
      Select Decode(预约, Null, 0, 0, 0, 1) Into n_预约挂号 From 病人挂号记录 Where NO = 单据号_In And Rownum = 1;
    Exception
      When Others Then
        n_预约挂号 := 0;
    End;
    n_检查数 := Null;
    Update 临床出诊记录
    Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - n_预约挂号, 已约数 = Nvl(已约数, 0) - n_预约挂号
    Where ID = n_出诊记录id
    Returning 已挂数 Into n_检查数;
  
    If Nvl(n_检查数, 0) < 0 Then
      Update 临床出诊记录 Set 已挂数 = 0 Where ID = n_出诊记录id;
    End If;
  
    Update 病人挂号汇总
    Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - n_预约挂号, 已约数 = Nvl(已约数, 0) - n_预约挂号
    Where 日期 = Trunc(r_Registrow.发生时间) And Nvl(医生id, 0) = Nvl(r_Registrow.医生id, 0) And
          Nvl(医生姓名, '-') = Nvl(r_Registrow.医生姓名, '-') And 科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And
          (号码 = r_Registrow.号码 Or 号码 Is Null);
  
    If Sql%RowCount = 0 Then
      Insert Into 病人挂号汇总
        (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 其中已接收, 已约数)
      Values
        (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
         Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, -1, -1 * n_预约挂号, -1 * n_预约挂号);
    End If;
  
    Close c_Registinfo;
  End If;

  If n_记帐 = 0 Then
    --人员缴款余额(包括个人帐户等的结算金额,不含退冲预交款)
    For r_Opermoney In c_Opermoney(n_销帐id) Loop
      Update 人员缴款余额
      Set 余额 = Nvl(余额, 0) + (-1 * r_Opermoney.冲预交)
      Where 收款员 = 操作员姓名_In And 结算方式 = r_Opermoney.结算方式 And 性质 = 1
      Returning 余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 人员缴款余额
          (收款员, 结算方式, 性质, 余额)
        Values
          (操作员姓名_In, r_Opermoney.结算方式, 1, -1 * r_Opermoney.冲预交);
        n_返回值 := r_Opermoney.冲预交;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 人员缴款余额
        Where 收款员 = 操作员姓名_In And 结算方式 = r_Opermoney.结算方式 And 性质 = 1 And Nvl(余额, 0) = 0;
      End If;
    End Loop;
  End If;
  If Nvl(退费类型_In, 0) Not In (2, 3) Then
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
    Begin
      Select 病人id, 发生时间 Into n_就诊病人id, d_就诊时间 From 病人挂号记录 Where NO = 单据号_In;
      Delete From 就诊登记记录 Where 病人id = n_就诊病人id And 就诊时间 = d_就诊时间 And 主页id Is Null;
    Exception
      When Others Then
        Null;
    End;
  
    --病人挂号记录
    Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
  
    Update 病人挂号记录 Set 记录状态 = 3 Where NO = 单据号_In And 记录性质 = 1 And 记录状态 = 1;
    If Sql%NotFound Then
      v_Err_Msg := '挂号单【' || 单据号_In || '】不存在或由于并发原因已经被退号';
      Raise Err_Item;
    End If;
    Insert Into 病人挂号记录
      (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 操作员编号, 操作员姓名,
       复诊, 号序, 社区, 预约, 摘要, 预约方式, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 险类, 医疗付款方式, 出诊记录id)
      Select n_挂号id, NO, 记录性质, 2, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, d_Date, 发生时间,
             操作员编号_In, 操作员姓名_In, 复诊, 号序, 社区, 预约, Nvl(摘要_In, 摘要) As 摘要, 预约方式, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 险类, 医疗付款方式,
             n_出诊记录id
      From 病人挂号记录
      Where NO = 单据号_In;
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
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人挂号记录_出诊_Delete;
/

--127075:焦博,2018-07-06,调整使用到了部门参数分诊台签到排队的Oracle函数
Create Or Replace Procedure Zl_预约挂号接收_出诊_Insert
(
  No_In            门诊费用记录.No%Type,
  票据号_In        门诊费用记录.实际票号%Type,
  领用id_In        票据使用明细.领用id%Type,

  结帐id_In        门诊费用记录.结帐id%Type,
  诊室_In          门诊费用记录.发药窗口%Type,
  病人id_In        门诊费用记录.病人id%Type,
  门诊号_In        门诊费用记录.标识号%Type,
  姓名_In          门诊费用记录.姓名%Type,
  性别_In          门诊费用记录.性别%Type,
  年龄_In          门诊费用记录.年龄%Type,
  付款方式_In      门诊费用记录.付款方式%Type, --用于存放病人的医疗付款方式编号
  费别_In          门诊费用记录.费别%Type,
  结算方式_In      Varchar2, --现金的结算名称
  现金支付_In      病人预交记录.冲预交%Type, --挂号时现金支付部份金额
  预交支付_In      病人预交记录.冲预交%Type, --挂号时使用的预交金额
  个帐支付_In      病人预交记录.冲预交%Type, --挂号时个人帐户支付金额
  发生时间_In      门诊费用记录.发生时间%Type,
  号序_In          挂号序号状态.序号%Type,
  操作员编号_In    门诊费用记录.操作员编号%Type,
  操作员姓名_In    门诊费用记录.操作员姓名%Type,
  生成队列_In      Number := 0,
  登记时间_In      门诊费用记录.登记时间%Type := Null,
  卡类别id_In      病人预交记录.卡类别id%Type := Null,
  结算卡序号_In    病人预交记录.结算卡序号%Type := Null,
  卡号_In          病人预交记录.卡号%Type := Null,
  交易流水号_In    病人预交记录.交易流水号%Type := Null,
  交易说明_In      病人预交记录.交易说明%Type := Null,
  险类_In          病人挂号记录.险类%Type := Null,
  结算模式_In      Number := 0,
  记帐费用_In      Number := 0,
  冲预交病人ids_In Varchar2 := Null,
  三方调用_In      Number := 0,
  更新交款余额_In  Number := 1, --是否更新人员交款余额，主要是处理统一操作员登录多台自助机的情况
  摘要_In          病人挂号记录.摘要%Type := Null,
  收费单_In        病人挂号记录.收费单%Type := Null
) As
  --该游标用于收费冲预交的可用预交列表
  --以ID排序，优先冲上次未冲完的。
  Cursor c_Deposit
  (
    v_病人id        病人信息.病人id%Type,
    v_冲预交病人ids Varchar2
  ) Is
    Select 病人id, NO, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额, Min(记录状态) As 记录状态, Nvl(Max(结帐id), 0) As 结帐id,
           Max(Decode(记录性质, 1, ID, 0) * Decode(记录状态, 2, 0, 1)) As 原预交id
    From 病人预交记录
    Where 记录性质 In (1, 11) And 病人id In (Select Column_Value From Table(f_Num2list(v_冲预交病人ids))) And Nvl(预交类别, 2) = 1
     Having Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
    Group By NO, 病人id
    Order By Decode(病人id, Nvl(v_病人id, 0), 0, 1), 结帐id, NO;

  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  Err_Special Exception;
  v_操作员姓名 病人挂号记录.接收人%Type;
  v_现金       结算方式.名称%Type;
  v_个人帐户   结算方式.名称%Type;
  v_队列名称   排队叫号队列.队列名称%Type;
  v_号别       门诊费用记录.计算单位%Type;
  v_号序       门诊费用记录.发药窗口%Type;
  v_排队号码   排队叫号队列.排队号码 %Type;
  v_预约方式   病人挂号记录.预约方式 %Type;

  n_预交金额      病人预交记录.金额%Type;
  n_返回值        病人预交记录.金额%Type;
  v_冲预交病人ids Varchar2(4000);

  n_挂号id         病人挂号记录.Id%Type;
  n_分诊台签到排队 Number;
  n_组id           财务缴款分组.Id%Type;
  n_Count          Number(18);
  n_排队           Number;
  n_当天排队       Number;
  n_当前金额       病人预交记录.金额%Type;
  n_预交id         病人预交记录.Id%Type;

  d_Date         Date;
  d_预约时间     门诊费用记录.发生时间%Type;
  d_发生时间     Date;
  d_排队时间     Date;
  n_时段         Number := 0;
  n_存在         Number := 0;
  v_结算内容     Varchar2(2000);
  v_当前结算     Varchar2(500);
  n_结算金额     病人预交记录.冲预交%Type;
  v_结算号码     病人预交记录.结算号码%Type;
  v_结算方式     病人预交记录.结算方式%Type;
  n_三方卡标志   Number(3);
  v_排队序号     排队叫号队列.排队序号%Type;
  n_结算模式     病人信息.结算模式%Type;
  v_付款方式     病人挂号记录.医疗付款方式%Type;
  n_接收模式     Number := 0;
  n_出诊记录id   病人挂号记录.出诊记录id%Type;
  n_新出诊记录id 病人挂号记录.出诊记录id%Type;
  n_号源id       临床出诊记录.号源id%Type;
  n_预约顺序号   临床出诊序号控制.预约顺序号%Type;
  n_旧分时段     临床出诊记录.是否分时段%Type;
  n_旧序号控制   临床出诊记录.是否序号控制%Type;
  n_旧科室id     临床出诊记录.科室id%Type;
  n_旧项目id     临床出诊记录.项目id%Type;
  n_旧医生id     临床出诊记录.医生id%Type;
  n_挂号模式     Number(3);
  d_启用时间     Date;
  v_Paratemp     Varchar2(500);
  v_Registtemp   Varchar2(500);
  n_检查         Number(3);
  n_序号控制     临床出诊记录.是否序号控制%Type;
  v_旧上班时段   临床出诊记录.上班时段%Type;
Begin
  n_组id          := Zl_Get组id(操作员姓名_In);
  v_冲预交病人ids := Nvl(冲预交病人ids_In, 病人id_In);
  v_Paratemp      := Nvl(zl_GetSysParameter('挂号排班模式'), 0);
  n_接收模式      := Nvl(zl_GetSysParameter('预约接收模式', 1111), 0);
  n_挂号模式      := To_Number(Substr(v_Paratemp, 1, 1));
  If n_挂号模式 = 1 Then
    Begin
      d_启用时间 := To_Date(Substr(v_Paratemp, 3), 'yyyy-mm-dd hh24:mi:ss');
    Exception
      When Others Then
        d_启用时间 := Null;
    End;
  End If;

  --获取结算方式名称
  Begin
    Select 名称 Into v_现金 From 结算方式 Where 性质 = 1;
  Exception
    When Others Then
      v_现金 := '现金';
  End;
  Begin
    Select 名称 Into v_个人帐户 From 结算方式 Where 性质 = 3;
  Exception
    When Others Then
      v_个人帐户 := '个人帐户';
  End;
  If 登记时间_In Is Null Then
    Select Sysdate Into d_Date From Dual;
  Else
    d_Date := 登记时间_In;
  End If;

  --更新挂号序号状态
  Begin
    Select 号别, 号序, Trunc(发生时间), 发生时间, 预约方式, 出诊记录id
    Into v_号别, v_号序, d_预约时间, d_发生时间, v_预约方式, n_出诊记录id
    From 病人挂号记录
    Where 记录性质 = 2 And 记录状态 = 1 And Rownum = 1 And NO = No_In;
  Exception
    When Others Then
      Select Max(接收人) Into v_操作员姓名 From 病人挂号记录 Where 记录性质 = 2 And 记录状态 In (1, 3) And NO = No_In;
      If v_操作员姓名 Is Null Then
        v_Err_Msg := '当前预约挂号单已被取消';
        Raise Err_Item;
      Else
        If v_操作员姓名 = 操作员姓名_In Then
          v_Err_Msg := '当前预约挂号单已被接收';
          Raise Err_Special;
        Else
          v_Err_Msg := '当前预约挂号单已被其它人接收';
          Raise Err_Item;
        End If;
      End If;
  End;

  --判断是否分时段
  Select Nvl(是否分时段, 0), 号源id, Nvl(是否序号控制, 0)
  Into n_时段, n_号源id, n_序号控制
  From 临床出诊记录
  Where ID = n_出诊记录id;

  If n_时段 = 1 And 三方调用_In = 0 And n_接收模式 = 0 Then
    If Trunc(发生时间_In) <> Trunc(Sysdate) Then
      v_Err_Msg := '分时段的预约挂号单只能当天接收！';
      Raise Err_Item;
    End If;
  End If;

  If n_时段 = 0 And 三方调用_In = 0 Then
    If n_接收模式 = 0 Then
      If Trunc(发生时间_In) = Trunc(Sysdate) Then
        d_发生时间 := 发生时间_In;
      Else
        d_发生时间 := Sysdate;
      End If;
    Else
      d_发生时间 := 发生时间_In;
    End If;
  Else
    If Not 发生时间_In Is Null Then
      d_发生时间 := 发生时间_In;
    End If;
  End If;

  If d_启用时间 Is Not Null Then
    If d_发生时间 < d_启用时间 Then
      v_Err_Msg := '当前预约挂号单属于出诊表排班模式安排，不能在' || To_Char(d_启用时间, 'yyyy-mm-dd hh24:mi:ss') || '之前接收!';
      Raise Err_Item;
    End If;
  End If;

  If Not v_号序 Is Null Then
    If 号序_In Is Null Then
      Update 临床出诊序号控制 Set 挂号状态 = 0 Where (序号 = v_号序 Or 备注 = v_号序) And 记录id = n_出诊记录id;
    Else
      If Trunc(d_预约时间) <> Trunc(Sysdate) And n_接收模式 = 0 Then
        If n_时段 = 0 And 三方调用_In = 0 Then
          --提前接收或延迟接收
          Update 临床出诊序号控制 Set 挂号状态 = 0 Where 序号 = v_号序 And 记录id = n_出诊记录id;
        
          Select 是否分时段, 是否序号控制, 科室id, 医生id, 项目id, 上班时段
          Into n_旧分时段, n_旧序号控制, n_旧科室id, n_旧医生id, n_旧项目id, v_旧上班时段
          From 临床出诊记录
          Where ID = n_出诊记录id;
          Begin
            Select ID
            Into n_新出诊记录id
            From 临床出诊记录
            Where 号源id = n_号源id And 是否分时段 = n_旧分时段 And 是否序号控制 = n_旧序号控制 And 科室id = n_旧科室id And
                  Nvl(医生id, 0) = Nvl(n_旧医生id, 0) And 上班时段 = v_旧上班时段 And Nvl(是否发布, 0) = 1 And 出诊日期 = Trunc(Sysdate) And
                  Rownum < 2;
          Exception
            When Others Then
              v_Err_Msg := '接收当天没有对应的出诊安排,无法接收!';
              Raise Err_Item;
          End;
        
          Begin
            Select 1
            Into n_存在
            From 临床出诊序号控制
            Where 记录id = n_新出诊记录id And 序号 = v_号序 And Nvl(挂号状态, 0) = 0;
          Exception
            When Others Then
              n_存在 := 0;
          End;
        
          If n_存在 = 1 Then
            Update 临床出诊序号控制
            Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In
            Where 记录id = n_新出诊记录id And 序号 = v_号序 And Nvl(挂号状态, 0) = 0;
          Else
            --号码已被使用的情况
            Select Min(序号) Into v_号序 From 临床出诊序号控制 Where 记录id = n_新出诊记录id And Nvl(挂号状态, 0) = 0;
            If v_号序 Is Null Then
              v_Err_Msg := '接收当天没有可用序号,无法接收!';
              Raise Err_Item;
            End If;
            Update 临床出诊序号控制
            Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In
            Where 记录id = n_新出诊记录id And 序号 = v_号序 And Nvl(挂号状态, 0) = 0;
          End If;
        Else
          Select 是否分时段, 是否序号控制, 科室id, 医生id, 项目id, 上班时段
          Into n_旧分时段, n_旧序号控制, n_旧科室id, n_旧医生id, n_旧项目id, v_旧上班时段
          From 临床出诊记录
          Where ID = n_出诊记录id;
          Begin
            Select ID
            Into n_新出诊记录id
            From 临床出诊记录
            Where 号源id = n_号源id And 是否分时段 = n_旧分时段 And 是否序号控制 = n_旧序号控制 And 科室id = n_旧科室id And
                  Nvl(医生id, 0) = Nvl(n_旧医生id, 0) And 上班时段 = v_旧上班时段 And Nvl(是否发布, 0) = 1 And 出诊日期 = Trunc(Sysdate) And
                  Rownum < 2;
          Exception
            When Others Then
              v_Err_Msg := '接收当天没有对应的出诊安排,无法接收!';
              Raise Err_Item;
          End;
          Update 临床出诊序号控制
          Set 挂号状态 = 0, 操作员姓名 = 操作员姓名_In
          Where (序号 = v_号序 Or 备注 = v_号序) And 记录id = n_出诊记录id And Nvl(挂号状态, 0) = 2
          Returning 预约顺序号 Into n_预约顺序号;
        
          Update 临床出诊序号控制
          Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In, 预约顺序号 = n_预约顺序号
          Where 序号 = v_号序 And 记录id = n_新出诊记录id And Nvl(挂号状态, 0) = 0;
          If Sql% RowCount = 0 Then
            v_Err_Msg := '接收当天序号' || v_号序 || '已被其它人使用,无法接收.';
            Raise Err_Item;
          End If;
        End If;
      Else
        Update 临床出诊序号控制
        Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In
        Where (序号 = v_号序 Or 备注 = v_号序) And 记录id = n_出诊记录id;
        If Sql%RowCount = 0 Then
          v_Err_Msg := '序号' || v_号序 || '已被其它人使用,请重新选择一个序号.';
          Raise Err_Item;
        End If;
      End If;
    End If;
  Else
    If Not 号序_In Is Null Then
      If Trunc(d_预约时间) <> Trunc(Sysdate) And n_接收模式 = 0 Then
        Select 是否分时段, 是否序号控制, 科室id, 医生id, 项目id, 上班时段
        Into n_旧分时段, n_旧序号控制, n_旧科室id, n_旧医生id, n_旧项目id, v_旧上班时段
        From 临床出诊记录
        Where ID = n_出诊记录id;
        Begin
          Select ID
          Into n_新出诊记录id
          From 临床出诊记录
          Where 号源id = n_号源id And 是否分时段 = n_旧分时段 And 是否序号控制 = n_旧序号控制 And 科室id = n_旧科室id And
                Nvl(医生id, 0) = Nvl(n_旧医生id, 0) And 上班时段 = v_旧上班时段 And Nvl(是否发布, 0) = 1 And 出诊日期 = Trunc(Sysdate) And
                Rownum < 2;
        Exception
          When Others Then
            v_Err_Msg := '接收当天没有对应的出诊安排,无法接收!';
            Raise Err_Item;
        End;
        Update 临床出诊序号控制
        Set 挂号状态 = 0, 操作员姓名 = 操作员姓名_In
        Where (序号 = 号序_In Or 备注 = 号序_In) And 记录id = n_出诊记录id And Nvl(挂号状态, 0) = 2
        Returning 预约顺序号 Into n_预约顺序号;
        Update 临床出诊序号控制
        Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In, 预约顺序号 = n_预约顺序号
        Where 序号 = 号序_In And 记录id = n_新出诊记录id And Nvl(挂号状态, 0) = 0;
        If Sql%RowCount = 0 Then
          v_Err_Msg := '接收当天序号' || 号序_In || '已被其它人使用,无法接收.';
          Raise Err_Item;
        End If;
      Else
        Update 临床出诊序号控制
        Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In
        Where (序号 = 号序_In Or 备注 = 号序_In) And 记录id = n_出诊记录id;
      
      End If;
      v_号序 := 号序_In;
    Else
      v_号序 := Null;
    End If;
  End If;

  --更新门诊费用记录
  Update 门诊费用记录
  Set 记录状态 = 1, 实际票号 = Decode(Nvl(记帐费用_In, 0), 1, Null, 票据号_In), 结帐id = Decode(Nvl(记帐费用_In, 0), 1, Null, 结帐id_In),
      结帐金额 = Decode(Nvl(记帐费用_In, 0), 1, Null, 实收金额), 发药窗口 = 诊室_In, 病人id = 病人id_In, 标识号 = 门诊号_In, 姓名 = 姓名_In, 年龄 = 年龄_In,
      性别 = 性别_In, 付款方式 = 付款方式_In, 费别 = 费别_In, 发生时间 = d_发生时间, 登记时间 = d_Date, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In,
      缴款组id = n_组id, 记帐费用 = Decode(Nvl(记帐费用_In, 0), 1, 1, 0), 摘要 = Decode(收费单_In, Null, Nvl(摘要_In, 摘要), '划价:' || 收费单_In)
  Where 记录性质 = 4 And 记录状态 = 0 And NO = No_In;

  v_Registtemp := zl_GetSysParameter('挂号排班模式');
  If Substr(v_Registtemp, 1, 1) = 1 Then
    Begin
      If To_Date(Substr(v_Registtemp, 3), 'yyyy-mm-dd hh24:mi:ss') > d_发生时间 Then
        v_Err_Msg := '接收时间' || To_Char(d_发生时间, 'yyyy-mm-dd hh24:mi:ss') || '未启用出诊表排班模式,目前无法接收!';
        Raise Err_Item;
      End If;
    Exception
      When Others Then
        Null;
    End;
    Begin
      Select 1
      Into n_检查
      From 临床出诊记录
      Where ID = Nvl(n_新出诊记录id, n_出诊记录id) And d_发生时间 Between 停诊开始时间 And 停诊终止时间;
    Exception
      When Others Then
        n_检查 := 0;
    End;
    If n_检查 = 1 And Not (n_时段 = 1 And n_序号控制 = 1) Then
      v_Err_Msg := '接收时间' || To_Char(d_发生时间, 'yyyy-mm-dd hh24:mi:ss') || '的安排已经被停诊,无法接收!';
      Raise Err_Item;
    End If;
  End If;

  --病人挂号记录
  Update 病人挂号记录
  Set 接收人 = 操作员姓名_In, 接收时间 = d_Date, 记录性质 = 1, 病人id = 病人id_In, 门诊号 = 门诊号_In, 发生时间 = d_发生时间, 姓名 = 姓名_In, 性别 = 性别_In,
      年龄 = 年龄_In, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In, 险类 = Decode(Nvl(险类_In, 0), 0, Null, 险类_In), 号序 = v_号序, 诊室 = 诊室_In,
      出诊记录id = Nvl(n_新出诊记录id, n_出诊记录id), 摘要 = Nvl(摘要_In, 摘要), 收费单 = 收费单_In
  Where 记录状态 = 1 And NO = No_In And 记录性质 = 2
  Returning ID Into n_挂号id;
  If Sql%NotFound Then
    Begin
      Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
      Begin
        Select 名称 Into v_付款方式 From 医疗付款方式 Where 编码 = 付款方式_In And Rownum < 2;
      Exception
        When Others Then
          v_付款方式 := Null;
      End;
      Insert Into 病人挂号记录
        (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 操作员编号, 操作员姓名,
         摘要, 号序, 预约, 预约方式, 接收人, 接收时间, 预约时间, 险类, 医疗付款方式, 出诊记录id, 收费单)
        Select n_挂号id, No_In, 1, 1, 病人id_In, 门诊号_In, 姓名_In, 性别_In, 年龄_In, 计算单位, 加班标志, 诊室_In, Null, 执行部门id, 执行人, 0, Null,
               登记时间, 发生时间, 操作员编号, 操作员姓名, Nvl(摘要_In, 摘要), v_号序, 1, Substr(结论, 1, 10) As 预约方式, 操作员姓名_In,
               Nvl(登记时间_In, Sysdate), 发生时间, Decode(Nvl(险类_In, 0), 0, Null, 险类_In), v_付款方式, Nvl(n_新出诊记录id, n_出诊记录id),
               收费单_In
        From 门诊费用记录
        Where 记录性质 = 4 And 记录状态 = 1 And Rownum = 1 And NO = No_In;
    Exception
      When Others Then
        v_Err_Msg := '由于并发原因,单据号为【' || No_In || '】的病人' || 姓名_In || '已经被接收';
        Raise Err_Item;
    End;
  End If;

  --0-不产生队列;1-按医生或分诊台排队;2-先分诊,后医生站
  If Nvl(生成队列_In, 0) <> 0 Then
    For v_挂号 In (Select ID, 姓名, 诊室, 执行人, 执行部门id, 发生时间, 号别, 号序 From 病人挂号记录 Where NO = No_In) Loop
      n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113, 1, Nvl(v_挂号.执行部门id, 0)));
      If Nvl(n_分诊台签到排队, 0) = 0 Then
        Begin
          Select 1,
                 Case
                   When 排队时间 < Trunc(Sysdate) Then
                    1
                   Else
                    0
                 End
          Into n_排队, n_当天排队
          From 排队叫号队列
          Where 业务类型 = 0 And 业务id = v_挂号.Id And Rownum <= 1;
        Exception
          When Others Then
            n_排队 := 0;
        End;
        If n_排队 = 0 Then
          --产生队列
          --按”执行部门”产生队列
          n_挂号id   := v_挂号.Id;
          v_队列名称 := v_挂号.执行部门id;
          v_排队号码 := Zlgetnextqueue(v_挂号.执行部门id, n_挂号id, v_挂号.号别 || '|' || v_挂号.号序);
          v_排队序号 := Zlgetsequencenum(0, n_挂号id, 0);
        
          --挂号id_In,号码_In,号序_In,缺省日期_In,扩展_In(暂无用)
          d_排队时间 := Zl_Get_Queuedate(n_挂号id, v_挂号.号别, v_挂号.号序, v_挂号.发生时间);
          --   队列名称_In , 业务类型_In, 业务id_In,科室id_In,排队号码_In,排队标记_In,患者姓名_In,病人ID_IN, 诊室_In, 医生姓名_In,
          Zl_排队叫号队列_Insert(v_队列名称, 0, n_挂号id, v_挂号.执行部门id, v_排队号码, Null, 姓名_In, 病人id_In, v_挂号.诊室, v_挂号.执行人, d_排队时间,
                           v_预约方式, Null, v_排队序号);
        Elsif Nvl(n_当天排队, 0) = 1 Then
          --更新队列号
          v_排队号码 := Zlgetnextqueue(v_挂号.执行部门id, v_挂号.Id, v_挂号.号别 || '|' || Nvl(v_挂号.号序, 0));
          v_排队序号 := Zlgetsequencenum(0, v_挂号.Id, 1);
          --新队列名称_IN, 业务类型_In, 业务id_In , 科室id_In , 患者姓名_In , 诊室_In, 医生姓名_In ,排队号码_In
          Zl_排队叫号队列_Update(v_挂号.执行部门id, 0, v_挂号.Id, v_挂号.执行部门id, v_挂号.姓名, v_挂号.诊室, v_挂号.执行人, v_排队号码, v_排队序号);
        
        Else
          --新队列名称_IN, 业务类型_In, 业务id_In , 科室id_In , 患者姓名_In , 诊室_In, 医生姓名_In ,排队号码_In
          Zl_排队叫号队列_Update(v_挂号.执行部门id, 0, v_挂号.Id, v_挂号.执行部门id, v_挂号.姓名, v_挂号.诊室, v_挂号.执行人);
        End If;
      End If;
    End Loop;
  End If;

  --汇总结算到病人预交记录
  If Nvl(记帐费用_In, 0) = 0 Then
    If Nvl(现金支付_In, 0) = 0 And Nvl(个帐支付_In, 0) = 0 And Nvl(预交支付_In, 0) = 0 Then
      Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      Insert Into 病人预交记录
        (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
         结算性质)
      Values
        (n_预交id, 4, 1, No_In, Decode(病人id_In, 0, Null, 病人id_In), v_现金, 0, 登记时间_In, 操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费',
         n_组id, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, Null, 4);
    End If;
    If Nvl(现金支付_In, 0) <> 0 Then
      v_结算内容 := 结算方式_In || '|'; --以空格分开以|结尾,没有结算号码的
      While v_结算内容 Is Not Null Loop
        v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '|') - 1);
        v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
      
        v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1));
      
        v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        v_结算号码 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
      
        v_当前结算   := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        n_三方卡标志 := To_Number(v_当前结算);
      
        If n_三方卡标志 = 0 Then
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明,
             合作单位, 结算性质, 结算号码)
          Values
            (n_预交id, 4, 1, No_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(v_结算方式, v_现金), Nvl(n_结算金额, 0), 登记时间_In,
             操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费', n_组id, Null, Null, Null, Null, Null, Null, 4, v_结算号码);
        Else
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明,
             合作单位, 结算性质, 结算号码)
          Values
            (n_预交id, 4, 1, No_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(v_结算方式, v_现金), Nvl(n_结算金额, 0), 登记时间_In,
             操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费', n_组id, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, Null, 4, v_结算号码);
        
          If Nvl(结算卡序号_In, 0) <> 0 And Nvl(n_结算金额, 0) <> 0 Then
            Zl_病人卡结算记录_支付(结算卡序号_In, 卡号_In, 0, n_结算金额, n_预交id, 操作员编号_In, 操作员姓名_In, 登记时间_In);
          End If;
        End If;
      
        If Nvl(更新交款余额_In, 1) = 1 Then
          Update 人员缴款余额
          Set 余额 = Nvl(余额, 0) + n_结算金额
          Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = Nvl(v_结算方式, v_现金)
          Returning 余额 Into n_返回值;
        
          If Sql%RowCount = 0 Then
            Insert Into 人员缴款余额
              (收款员, 结算方式, 性质, 余额)
            Values
              (操作员姓名_In, Nvl(v_结算方式, v_现金), 1, n_结算金额);
            n_返回值 := n_结算金额;
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 人员缴款余额
            Where 收款员 = 操作员姓名_In And 结算方式 = Nvl(v_结算方式, v_现金) And 性质 = 1 And Nvl(余额, 0) = 0;
          End If;
        End If;
      
        v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '|') + 1);
      End Loop;
    End If;
  End If;

  --对于就诊卡通过预交金挂号
  If Nvl(预交支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 Then
    n_预交金额 := 预交支付_In;
    For r_Deposit In c_Deposit(病人id_In, v_冲预交病人ids) Loop
      n_当前金额 := Case
                  When r_Deposit.金额 - n_预交金额 < 0 Then
                   r_Deposit.金额
                  Else
                   n_预交金额
                End;
      If r_Deposit.结帐id = 0 Then
        --第一次冲预交(填上结帐ID,金额为0)
        Update 病人预交记录 Set 冲预交 = 0, 结帐id = 结帐id_In, 结算性质 = 4 Where ID = r_Deposit.原预交id;
      End If;
      --冲上次剩余额
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
         结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 预交类别, 结算序号, 结算性质)
        Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 登记时间_In,
               操作员姓名_In, 操作员编号_In, n_当前金额, 结帐id_In, n_组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 预交类别, 结帐id_In, 4
        From 病人预交记录
        Where NO = r_Deposit.No And 记录状态 = r_Deposit.记录状态 And 记录性质 In (1, 11) And Rownum = 1;
    
      --更新病人预交余额
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) - n_当前金额
      Where 病人id = r_Deposit.病人id And 性质 = 1 And 类型 = Nvl(1, 2)
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (r_Deposit.病人id, Nvl(1, 2), -1 * n_当前金额, 1);
        n_返回值 := -1 * n_当前金额;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 病人余额
        Where 病人id = r_Deposit.病人id And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
      End If;
    
      --检查是否已经处理完
      If r_Deposit.金额 < n_预交金额 Then
        n_预交金额 := n_预交金额 - r_Deposit.金额;
      Else
        n_预交金额 := 0;
      End If;
    
      If n_预交金额 = 0 Then
        Exit;
      End If;
    End Loop;
  End If;

  --对于医保挂号
  If Nvl(个帐支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 Then
    Insert Into 病人预交记录
      (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
       预交类别, 结算序号, 结算性质)
    Values
      (病人预交记录_Id.Nextval, 4, 1, No_In, 病人id_In, v_个人帐户, 个帐支付_In, d_Date, 操作员编号_In, 操作员姓名_In, 结帐id_In, '医保挂号', n_组id,
       Null, Null, Null, Null, Null, Null, Null, 结帐id_In, 4);
  End If;

  --相关汇总表的处理
  --人员缴款余额
  If Nvl(个帐支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 And Nvl(更新交款余额_In, 1) = 1 Then
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) + 个帐支付_In
    Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = v_个人帐户
    Returning 余额 Into n_返回值;
  
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (操作员姓名_In, v_个人帐户, 1, 个帐支付_In);
      n_返回值 := 个帐支付_In;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 人员缴款余额 Where 收款员 = 操作员姓名_In And 性质 = 1 And Nvl(余额, 0) = 0;
    End If;
  End If;

  If Nvl(记帐费用_In, 0) = 1 Then
    --记帐
    If Nvl(病人id_In, 0) = 0 Then
      v_Err_Msg := '要针对病人的挂号费进行记帐，必须是建档病人才能记帐挂号。';
      Raise Err_Item;
    End If;
    For c_费用 In (Select 实收金额, 病人科室id, 开单部门id, 执行部门id, 收入项目id
                 From 门诊费用记录
                 Where 记录性质 = 4 And 记录状态 = 1 And NO = No_In And Nvl(记帐费用, 0) = 1) Loop
      --病人余额
      Update 病人余额
      Set 费用余额 = Nvl(费用余额, 0) + Nvl(c_费用.实收金额, 0)
      Where 病人id = Nvl(病人id_In, 0) And 性质 = 1 And 类型 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 性质, 类型, 费用余额, 预交余额)
        Values
          (病人id_In, 1, 1, Nvl(c_费用.实收金额, 0), 0);
      End If;
    
      --病人未结费用
      Update 病人未结费用
      Set 金额 = Nvl(金额, 0) + Nvl(c_费用.实收金额, 0)
      Where 病人id = 病人id_In And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(病人科室id, 0) = Nvl(c_费用.病人科室id, 0) And
            Nvl(开单部门id, 0) = Nvl(c_费用.开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(c_费用.执行部门id, 0) And 收入项目id + 0 = c_费用.收入项目id And
            来源途径 + 0 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人未结费用
          (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
        Values
          (病人id_In, Null, Null, c_费用.病人科室id, c_费用.开单部门id, c_费用.执行部门id, c_费用.收入项目id, 1, Nvl(c_费用.实收金额, 0));
      End If;
    End Loop;
  End If;
  If Nvl(病人id_In, 0) <> 0 Then
    n_结算模式 := 0;
    Update 病人信息
    Set 就诊时间 = d_发生时间, 就诊状态 = 1, 就诊诊室 = 诊室_In
    Where 病人id = 病人id_In
    Returning Nvl(结算模式, 0) Into n_结算模式;
    --取参数:
    If Nvl(n_结算模式, 0) <> Nvl(结算模式_In, 0) Then
      --结算模式的确定
      If n_结算模式 = 1 And Nvl(结算模式_In, 0) = 0 Then
        --病人已经是"先诊疗后结算的",本次是"先结算后诊疗的",则检查是否存在未结数据
        Select Count(1)
        Into n_Count
        From 病人未结费用
        Where 病人id = 病人id_In And (来源途径 In (1, 4) Or 来源途径 = 3 And Nvl(主页id, 0) = 0) And Nvl(金额, 0) <> 0 And Rownum < 2;
        If Nvl(n_Count, 0) <> 0 Then
          --存在未结算数据，必须先结算后才允许执行
          v_Err_Msg := '当前病人的就诊模式为先诊疗后结算且存在未结费用，不允许调整该病人的就诊模式,你可以先对未结费用结帐后再挂号或不调整病人的就诊模式!';
          Raise Err_Item;
        End If;
        --检查
        --未发生医嘱业务的（即当时就挂号的,需要保证同一次的就诊模式是一至的(程序已经检查，不用再处理)
      End If;
      Update 病人信息 Set 结算模式 = 结算模式_In Where 病人id = 病人id_In;
    End If;
  End If;

  --病人担保信息
  If 病人id_In Is Not Null Then
    Update 病人信息
    Set 担保人 = Null, 担保额 = Null, 担保性质 = Null
    Where 病人id = 病人id_In And Nvl(在院, 0) = 0 And Exists
     (Select 1
           From 病人担保记录
           Where 病人id = 病人id_In And 主页id Is Not Null And
                 登记时间 = (Select Max(登记时间) From 病人担保记录 Where 病人id = 病人id_In));
    If Sql%RowCount > 0 Then
      Update 病人担保记录
      Set 到期时间 = d_Date
      Where 病人id = 病人id_In And 主页id Is Not Null And Nvl(到期时间, d_Date) >= d_Date;
    End If;
  End If;
  --消息推送
  Begin
    Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
      Using 1, n_挂号id;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Err_Special Then
    Raise_Application_Error(-20105, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_预约挂号接收_出诊_Insert;
/

--127075:焦博,2018-07-06,调整使用到了部门参数分诊台签到排队的Oracle函数
Create Or Replace Procedure Zl_病人挂号补结算_Delete
(
  单据号_In     门诊费用记录.No%Type,
  操作员编号_In 门诊费用记录.操作员编号%Type,
  操作员姓名_In 门诊费用记录.操作员姓名%Type,
  冲销id_In     门诊费用记录.结帐id%Type := Null,
  结算序号_In   病人预交记录.结算序号%Type := Null,
  退号时间_In   门诊费用记录.登记时间%Type := Null,
  删除门诊号_In Number := 0
) As
  --该游标用于判断是否单独收病历费,及挂号汇总表处理

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_结帐id   病人预交记录.结帐id%Type;
  n_冲销id   门诊费用记录.结帐id%Type;
  n_结算序号 病人预交记录.结算序号%Type;

  n_病人id   病人信息.病人id%Type;
  n_退费金额 病人预交记录.冲预交%Type;
  n_挂号id   病人挂号记录.Id%Type;
  n_组id     财务缴款分组.Id%Type;

  n_分诊台签到排队 Number;
  n_挂号生成队列   Number;
  d_Date           Date;
  n_病人id1        病人信息.病人id%Type;
  d_Temp           Date;
  n_就诊病人id     病人信息.病人id%Type;
  d_就诊时间       就诊登记记录.就诊时间%Type;
Begin
  n_组id := Zl_Get组id(操作员姓名_In);
  --首先判断要退号/取消预约的记录是否存在

  Begin
    Select a.结帐id, a.病人id
    Into n_结帐id, n_病人id
    From 门诊费用记录 A
    Where a.记录性质 = 4 And a.No = 单据号_In And a.记录状态 = 1 And Rownum < 2;
  Exception
  
    When Others Then
      n_病人id := -1;
  End;
  If Nvl(n_病人id, 0) = -1 Then
    v_Err_Msg := '未找到指定的挂号单:' || 单据号_In || ',可能已经被人退号,不允许再次退号。';
    Raise Err_Item;
  End If;

  --2.挂号处理

  d_Date     := 退号时间_In;
  n_冲销id   := 冲销id_In;
  n_结算序号 := 结算序号_In;

  If d_Date Is Null Then
    d_Date := Sysdate;
  End If;
  If n_冲销id Is Null Then
    Select 病人结帐记录_Id.Nextval Into n_冲销id From Dual;
  End If;

  If n_结算序号 Is Null Then
    n_结算序号 := -1 * n_冲销id;
  End If;
  --更新挂号序号状态
  If Zl_To_Number(zl_GetSysParameter('已退序号允许挂号', 1111)) = 1 Then
    Delete 挂号序号状态
    Where 状态 = 1 And
          (号码, 序号, 日期) = (Select 号别, 号序, Trunc(发生时间) From 病人挂号记录 Where NO = 单据号_In And Rownum = 1) Or
          (号码, 序号, 日期) = (Select 号别, 号序, 发生时间 From 病人挂号记录 Where NO = 单据号_In And Rownum = 1);
  Else
    Update 挂号序号状态
    Set 状态 = 4
    Where 状态 = 1 And
          (号码, 序号, 日期) = (Select 号别, 号序, Trunc(发生时间) From 病人挂号记录 Where NO = 单据号_In And Rownum = 1) Or
          (号码, 序号, 日期) = (Select 号别, 号序, 发生时间 From 病人挂号记录 Where NO = 单据号_In And Rownum = 1);
  End If;

  --病人就诊状态
  If n_病人id Is Not Null Then
    Update 病人信息 Set 就诊状态 = 0, 就诊诊室 = Null Where 病人id = n_病人id;
    --删除门诊号相关处理,只有当只有一条挂号记录并且病人建档日期与挂号日期近似时才会处理(界面要根据提示来删除)
    If 删除门诊号_In = 1 Then
      Delete 门诊病案记录 Where 病人id = n_病人id;
      Update 病人信息 Set 门诊号 = Null Where 病人id = n_病人id;
      --费用记录包括挂号及病案、就诊卡费用,以及病人交费后退费或销帐的费用,挂号记录在最后处理
      Update 门诊费用记录 Set 标识号 = Null Where 门诊标志 = 1 And 病人id = n_病人id;
    End If;
  End If;

  --如果挂时收了就诊卡费,退费时清除就诊卡号,在非光退病历费时
  n_病人id1 := Null;
  Begin
    Select 病人id
    Into n_病人id1
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And 附加标志 = 2 And Rownum < 2;
  Exception
    When Others Then
      Null;
  End;
  If n_病人id1 Is Not Null Then
    Update 病人信息
    Set 就诊卡号 = Null, 卡验证码 = Null, Ic卡号 = Decode(Ic卡号, 就诊卡号, Null, Ic卡号)
    Where 病人id = n_病人id1;
  End If;

  --门诊费用记录
  --冲销记录
  Insert Into 门诊费用记录
    (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位, 付数,
     数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 执行人, 操作员编号, 操作员姓名, 发生时间, 登记时间,
     结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 缴款组id, 费用状态, 执行状态)
    Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别,
           收费细目id, 计算单位, 付数, -数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, -应收金额, -实收金额, 开单部门id, 开单人, 执行部门id, 执行人,
           操作员编号_In, 操作员姓名_In, 发生时间, d_Date, n_冲销id, -1 * 结帐金额, 保险项目否, 保险大类id, -1 * 统筹金额, 摘要 As 摘要, 附加标志, 保险编码, 费用类型,
           n_组id, 1, 1
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In;

  Update 门诊费用记录 Set 记录状态 = 3 Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In;
  Select Sum(实收金额) Into n_退费金额 From 门诊费用记录 Where 结帐id = n_冲销id;
  Insert Into 病人预交记录
    (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 结算序号, 缴款组id, 预交类别, 卡类别id,
     结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算性质)
    Select 病人预交记录_Id.Nextval, a.No, a.实际票号, 4, 2, a.病人id, a.主页id, a.科室id, a.摘要, Null, d_Date, 操作员编号_In, 操作员姓名_In, n_退费金额,
           n_冲销id, n_结算序号, n_组id, 预交类别, Null, Null, Null, Null, Null, Null, 1, 4
    From 病人预交记录 A
    Where a.结帐id = n_结帐id And Rownum < 2;
  If Sql%NotFound Then
    v_Err_Msg := '未找到挂号单为【' || 单据号_In || '】的原始挂号记录!';
    Raise Err_Item;
  End If;
  Update 病人预交记录 Set 记录状态 = 3 Where Mod(记录性质, 10) <> 1 And 结帐id = n_结帐id;

  --病人挂号汇总
  For c_挂号 In (Select a.收费细目id, a.发生时间, a.登记时间, c.接收时间, c.执行部门id, c.执行人, m.Id As 医生id, Nvl(c.号别, b.号码) As 号码, a.结帐id,
                      a.病人id, Decode(c.预约, Null, 0, 0, 0, 1) As 预约
               From 门诊费用记录 A, 病人挂号记录 C, 挂号安排 B, 人员表 M
               Where a.记录性质 = 4 And a.结帐id = n_冲销id And a.从属父号 Is Null And c.执行人 = m.姓名(+) And a.No = c.No And
                     Nvl(c.号别, Nvl(a.计算单位, '-')) = b.号码 And Nvl(a.附加标志, 0) = 0 And Rownum < 2) Loop
    --退非挂号费用,则不处理汇总表数据
  
    If Nvl(c_挂号.预约, 0) <> 0 Then
      d_Temp := Trunc(c_挂号.接收时间);
    Else
      d_Temp := Trunc(c_挂号.发生时间);
    End If;
  
    Update 病人挂号汇总
    Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - Nvl(c_挂号.预约, 0), 已约数 = Nvl(已约数, 0) - Nvl(c_挂号.预约, 0)
    Where 日期 = d_Temp And 科室id = c_挂号.执行部门id And 项目id = c_挂号.收费细目id And Nvl(医生姓名, '医生') = Nvl(c_挂号.执行人, '医生') And
          Nvl(医生id, 0) = Nvl(c_挂号.医生id, 0) And (号码 = c_挂号.号码 Or 号码 Is Null);
  
    If Sql%RowCount = 0 Then
      Insert Into 病人挂号汇总
        (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 其中已接收, 已约数)
      Values
        (d_Temp, c_挂号.执行部门id, c_挂号.收费细目id, c_挂号.执行人, Decode(c_挂号.医生id, 0, Null, c_挂号.医生id), c_挂号.号码, -1,
         -1 * Nvl(c_挂号.预约, 0), -1 * Nvl(c_挂号.预约, 0));
    End If;
  End Loop;

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
  Begin
    Select 病人id, 发生时间 Into n_就诊病人id, d_就诊时间 From 病人挂号记录 Where NO = 单据号_In;
    Delete From 就诊登记记录 Where 病人id = n_就诊病人id And 就诊时间 = d_就诊时间 And 主页id Is Null;
  Exception
    When Others Then
      Null;
  End;
  --病人挂号记录
  Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;

  Update 病人挂号记录 Set 记录状态 = 3 Where NO = 单据号_In And 记录性质 = 1 And 记录状态 = 1;
  If Sql%NotFound Then
    v_Err_Msg := '挂号单【' || 单据号_In || '】不存在或由于并发原因已经被退号';
    Raise Err_Item;
  End If;

  Insert Into 病人挂号记录
    (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 操作员编号, 操作员姓名, 复诊,
     号序, 社区, 预约, 摘要, 预约方式, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 险类, 医疗付款方式)
    Select n_挂号id, NO, 记录性质, 2, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, d_Date, 发生时间, 操作员编号_In,
           操作员姓名_In, 复诊, 号序, 社区, 预约, 摘要 As 摘要, 预约方式, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 险类, 医疗付款方式
    From 病人挂号记录
    Where NO = 单据号_In And 记录状态 = 3;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人挂号补结算_Delete;
/

--127075:焦博,2018-07-06,调整使用到了部门参数分诊台签到排队的Oracle函数
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
    Select a.发生时间, a.登记时间, b.项目id, b.科室id, b.医生姓名, b.医生id, b.号码
    From 病人挂号记录 A, 挂号安排 B
    Where a.记录性质 = Decode(v_无效单据, 0, v_性质, a.记录性质) And a.记录状态 = v_状态 And a.No = 单据号_In And a.号别 = b.号码 And Rownum = 1;

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

--127075:焦博,2018-07-06,调整使用到了部门参数分诊台签到排队的Oracle函数
Create Or Replace Procedure Zl_病人挂号记录_Insert
(
  病人id_In        门诊费用记录.病人id%Type,
  门诊号_In        门诊费用记录.标识号%Type,
  姓名_In          门诊费用记录.姓名%Type,
  性别_In          门诊费用记录.性别%Type,
  年龄_In          门诊费用记录.年龄%Type,
  付款方式_In      门诊费用记录.付款方式%Type, --用于存放病人的医疗付款方式编号
  费别_In          门诊费用记录.费别%Type,
  单据号_In        门诊费用记录.No%Type,
  票据号_In        门诊费用记录.实际票号%Type,
  序号_In          门诊费用记录.序号%Type,
  价格父号_In      门诊费用记录.价格父号%Type,
  从属父号_In      门诊费用记录.从属父号%Type,
  收费类别_In      门诊费用记录.收费类别%Type,
  收费细目id_In    门诊费用记录.收费细目id%Type,
  数次_In          门诊费用记录.数次%Type,
  标准单价_In      门诊费用记录.标准单价%Type,
  收入项目id_In    门诊费用记录.收入项目id%Type,
  收据费目_In      门诊费用记录.收据费目%Type,
  结算方式_In      病人预交记录.结算方式%Type, --现金的结算名称
  应收金额_In      门诊费用记录.应收金额%Type,
  实收金额_In      门诊费用记录.实收金额%Type,
  病人科室id_In    门诊费用记录.病人科室id%Type,
  开单部门id_In    门诊费用记录.开单部门id%Type,
  执行部门id_In    门诊费用记录.执行部门id%Type,
  操作员编号_In    门诊费用记录.操作员编号%Type,
  操作员姓名_In    门诊费用记录.操作员姓名%Type,
  发生时间_In      门诊费用记录.发生时间%Type,
  登记时间_In      门诊费用记录.登记时间%Type,
  医生姓名_In      挂号安排.医生姓名%Type,
  医生id_In        挂号安排.医生id%Type,
  病历费_In        Number, --该条记录是否病历工本费
  急诊_In          Number,
  号别_In          挂号安排.号码%Type,
  诊室_In          门诊费用记录.发药窗口%Type,
  结帐id_In        门诊费用记录.结帐id%Type,
  领用id_In        票据使用明细.领用id%Type,
  预交支付_In      病人预交记录.冲预交%Type, --刷卡挂号时使用的预交金额,序号为1传入.
  现金支付_In      病人预交记录.冲预交%Type, --挂号时现金支付部份金额,序号为1传入.
  个帐支付_In      病人预交记录.冲预交%Type, --挂号时个人帐户支付金额,,序号为1传入.
  保险大类id_In    门诊费用记录.保险大类id%Type,
  保险项目否_In    门诊费用记录.保险项目否%Type,
  统筹金额_In      门诊费用记录.统筹金额%Type,
  摘要_In          门诊费用记录.摘要%Type, --预约挂号摘要信息
  预约挂号_In      Number := 0, --预约挂号时用(记录状态=0,发生时间为预约时间),此时不需要传入结算相关参数
  收费票据_In      Number := 0, --挂号是否使用收费票据
  保险编码_In      门诊费用记录.保险编码%Type,
  复诊_In          病人挂号记录.复诊%Type := 0,
  号序_In          挂号序号状态.序号%Type := Null, --预约时填入费用记录的发药窗口字段,挂号时填入挂号记录
  社区_In          病人挂号记录.社区%Type := Null,
  预约接收_In      Number := 0,
  预约方式_In      预约方式.名称%Type := Null,
  生成队列_In      Number := 0,
  卡类别id_In      病人预交记录.卡类别id%Type := Null,
  结算卡序号_In    病人预交记录.结算卡序号%Type := Null,
  卡号_In          病人预交记录.卡号%Type := Null,
  交易流水号_In    病人预交记录.交易流水号%Type := Null,
  交易说明_In      病人预交记录.交易说明%Type := Null,
  合作单位_In      病人预交记录.合作单位%Type := Null,
  操作类型_In      Number := 0,
  险类_In          病人挂号记录.险类%Type := Null,
  结算模式_In      Number := 0,
  记帐费用_In      Number := 0,
  退号重用_In      Number := 1,
  冲预交病人ids_In Varchar2 := Null,
  修正病人费别_In  Number := 0,
  修正病人年龄_In  Number := 0,
  收费单_In        病人挂号记录.收费单%Type := Null,
  更新交款余额_In  Number := 1
) As
  ---------------------------------------------------------------------------
  --
  --参数:
  --     操作类型_in:0-正常挂号或者预约 1-操作员拥有加号权限加号
  --     修正病人费别_In:0-不修改病人费别 1-修改病人费别
  --     更新交款余额_In:0-在zl_人员缴款余额_Update 中更新 1-在本过程中更新
  ----------------------------------------------------------------------------
  --该游标用于收费冲预交的可用预交列表
  --以ID排序，优先冲上次未冲完的。
  Cursor c_Deposit
  (
    v_病人id        病人信息.病人id%Type,
    v_冲预交病人ids Varchar2
  ) Is
    Select 病人id, NO, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额, Min(记录状态) As 记录状态, Nvl(Max(结帐id), 0) As 结帐id,
           Max(Decode(记录性质, 1, ID, 0) * Decode(记录状态, 2, 0, 1)) As 原预交id, Min(Decode(记录性质, 1, 收款时间, Null)) As 收款时间
    From 病人预交记录
    Where 记录性质 In (1, 11) And 病人id In (Select Column_Value From Table(f_Num2list(v_冲预交病人ids))) And Nvl(预交类别, 2) = 1
     Having Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
    Group By NO, 病人id
    Order By Decode(病人id, Nvl(v_病人id, 0), 0, 1), 收款时间;
  --功能：新增一行病人挂号费用，并将结算汇总到病人预交记录
  --       同时汇总相关的汇总表(病人挂号汇总、费用汇总)
  --       第一行费用处理票据使用情况(领用ID_IN>0)

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_排队号码 排队叫号队列.排队号码%Type;
  v_现金     结算方式.名称%Type;
  v_个人帐户 结算方式.名称%Type;
  v_队列名称 排队叫号队列.队列名称%Type;

  n_分时段       Number;
  n_时段限号     Number;
  n_时段限约     Number;
  d_时段时间     Date;
  d_最大序号时间 Date;
  n_追加号       Number := 0; --处理时段过期 追加挂号的情况
  n_已约数       病人挂号汇总.已约数%Type;
  n_预约有效时间 Number;
  n_失效数       Number;
  n_失约挂号     Number := 0;
  n_已用数量     Number;
  n_锁定         Number := 0;

  n_费用id        门诊费用记录.Id%Type;
  n_预交金额      病人预交记录.金额%Type;
  n_当前金额      病人预交记录.金额%Type;
  n_返回值        病人预交记录.金额%Type;
  n_预交id        病人预交记录.Id%Type;
  n_挂号id        病人挂号记录.Id%Type;
  v_冲预交病人ids Varchar2(4000);

  n_组id           财务缴款分组.Id%Type;
  n_门诊号         病人信息.门诊号%Type;
  n_序号           挂号序号状态.序号%Type;
  n_已用序号       挂号序号状态.序号%Type;
  n_序号控制       挂号安排.序号控制%Type;
  n_分诊台签到排队 Number;
  n_Count          Number;
  n_限号数         Number(18);
  d_排队时间       Date;
  d_序号时间       Date;
  v_费别           费别.名称%Type;
  n_时段序号       Number := -1;
  n_预约生成队列   Number;
  n_安排id         挂号安排.Id%Type;
  n_计划id         挂号安排计划.Id%Type := 0;
  v_星期           挂号安排限制.限制项目%Type;
  n_限约数         Number(18);
  n_已挂数         Number(4) := 0;

  n_挂出的最大序号 Number(4) := 0;
  n_结算模式       病人信息.结算模式%Type;
  v_排队序号       排队叫号队列.排队序号%Type;
  v_机器名         挂号序号状态.机器名%Type;
  v_序号操作员     挂号序号状态.操作员姓名%Type;
  v_序号机器名     挂号序号状态.机器名%Type;
  v_付款方式       病人挂号记录.医疗付款方式%Type;
  v_Temp           Varchar2(3000);
  v_时间段         时间段.时间段%Type;
  d_检查开始时间   时间段.开始时间%Type;
  d_检查结束时间   时间段.终止时间%Type;
  n_出诊记录id     临床出诊记录.Id%Type;
  n_分时点显示     Number(3);
  d_启用时间       Date;
Begin
  --获取当前机器名称
  Select Terminal Into v_机器名 From V$session Where Audsid = Userenv('sessionid');
  n_组id          := Zl_Get组id(操作员姓名_In);
  v_冲预交病人ids := Nvl(冲预交病人ids_In, 病人id_In);

  If 费别_In Is Null Then
    Begin
      Select 名称 Into v_费别 From 费别 Where 缺省标志 = 1 And Rownum < 2;
      Update 病人信息 Set 费别 = v_费别 Where 病人id = 病人id_In;
    Exception
      When Others Then
        v_Err_Msg := '无法确定病人费别，请检查缺省费别是否正确设置！';
        Raise Err_Item;
    End;
  Else
    v_费别 := 费别_In;
    If Nvl(修正病人费别_In, 0) = 1 Then
      Begin
        Update 病人信息 Set 费别 = v_费别 Where 病人id = 病人id_In;
      Exception
        When Others Then
          v_Err_Msg := '没有找到对应的病人！';
          Raise Err_Item;
      End;
    End If;
  End If;

  If Nvl(修正病人年龄_In, 0) = 1 Then
    Begin
      Update 病人信息 Set 年龄 = 年龄_In Where 病人id = 病人id_In;
    Exception
      When Others Then
        v_Err_Msg := '没有找到对应的病人！';
        Raise Err_Item;
    End;
  End If;

  If 门诊号_In Is Not Null Then
    Begin
      Select Nvl(门诊号, 0) Into n_门诊号 From 病人信息 Where 病人id = 病人id_In;
    Exception
      When Others Then
        n_门诊号 := 0;
    End;
    If n_门诊号 = 0 Then
      Update 病人信息 Set 门诊号 = 门诊号_In Where 病人id = 病人id_In;
    End If;
  End If;

  Begin
    Delete From 挂号序号状态
    Where 号码 = 号别_In And 日期 = 发生时间_In And 序号 = 号序_In And 状态 = 3 And 操作员姓名 = 操作员姓名_In;
  Exception
    When Others Then
      Null;
  End;
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
    If d_启用时间 Is Not Null Then
      If 发生时间_In > d_启用时间 Then
        v_Err_Msg := '当前挂号的发生时间' || To_Char(发生时间_In, 'yyyy-mm-dd hh24:mi:ss') || '已经启用了出诊表排班模式,不能再使用计划排班模式挂号!';
        Raise Err_Item;
      End If;
    End If;
  End If;

  If Nvl(预约挂号_In, 0) = 0 Then
    --挂号或者预约接收
    --因为现在有按日编单据号规则,日挂号量不能超过10000次,所以要检查唯一约束。
    Select Count(*)
    Into n_Count
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 In (1, 3) And 序号 = 序号_In And NO = 单据号_In;
    If n_Count <> 0 Then
      v_Err_Msg := '挂号单据号重复,不能保存！' || Chr(13) || '如果使用了按日顺序编号,当日挂号量不能超过10000人次。';
      Raise Err_Item;
    End If;
  
    --获取结算方式名称
    Begin
      Select 名称 Into v_现金 From 结算方式 Where 性质 = 1;
    Exception
      When Others Then
        v_现金 := '现金';
    End;
    Begin
      Select 名称 Into v_个人帐户 From 结算方式 Where 性质 = 3;
    Exception
      When Others Then
        v_个人帐户 := '个人帐户';
    End;
  End If;

  n_序号 := 号序_In;
  Select Decode(To_Char(发生时间_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null)
  Into v_星期
  From Dual;

  --挂号获取安排
  Begin
    Select a.Id, a.序号控制, Nvl(b.限号数, 0), Nvl(b.限约数, 0)
    Into n_安排id, n_序号控制, n_限号数, n_限约数
    From 挂号安排 A, 挂号安排限制 B
    Where a.Id = b.安排id(+) And b.限制项目(+) = v_星期 And a.号码 = 号别_In;
  
  Exception
    When Others Then
      n_安排id := -1;
  End;

  --如果是病历费或者号别为空时不检查
  If Nvl(病历费_In, 0) = 0 Or 号别_In Is Not Null Then
    If n_安排id = -1 Then
      v_Err_Msg := '不存相应的挂号安排数据,请检查';
      Raise Err_Item;
    End If;
  End If;

  If Nvl(预约挂号_In, 0) = 1 Then
    --首先获取计划
    Begin
      Select ID
      Into n_计划id
      From 挂号安排计划
      Where 安排id = n_安排id And 审核时间 Is Not Null And
            Nvl(生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) =
            (Select Max(a.生效时间) As 生效
             From 挂号安排计划 A
             Where a.审核时间 Is Not Null And 发生时间_In Between Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                   Nvl(a.失效时间, To_Date('3000-01-01', 'yyyy-mm-dd')) And a.安排id = n_安排id) And
            发生时间_In Between Nvl(生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
            Nvl(失效时间, To_Date('3000-01-01', 'yyyy-mm-dd'));
    
    Exception
      When Others Then
        n_计划id := 0;
    End;
    If Nvl(n_计划id, 0) <> 0 Then
      Begin
        --获取计划的限制
        Select a.Id, a.序号控制, Nvl(b.限号数, 0) As 限号数, Nvl(b.限约数, 0) As 限约数
        Into n_计划id, n_序号控制, n_限号数, n_限约数
        From 挂号安排计划 A, 挂号计划限制 B
        Where a.号码 = 号别_In And a.Id = n_计划id And a.审核时间 Is Not Null And a.Id = b.计划id(+) And b.限制项目(+) = v_星期;
      Exception
        When Others Then
          v_Err_Msg := '不存相应的挂号安排或计划数据,请检查';
          Raise Err_Item;
      End;
    End If;
  End If;

  --获取是否分时段
  Begin
    If Nvl(n_计划id, 0) = 0 Then
      Select Count(Rownum) Into n_分时段 From 挂号安排时段 Where 星期 = v_星期 And 安排id = n_安排id And Rownum <= 1;
      Select Decode(To_Char(发生时间_In, 'D'), '1', 周日, '2', 周一, '3', 周二, '4', 周三, '5', 周四, '6', 周五, '7', 周六, Null)
      Into v_时间段
      From 挂号安排
      Where ID = n_安排id;
    Else
      Select Count(Rownum) Into n_分时段 From 挂号计划时段 Where 星期 = v_星期 And 计划id = n_计划id And Rownum <= 1;
      Select Decode(To_Char(发生时间_In, 'D'), '1', 周日, '2', 周一, '3', 周二, '4', 周三, '5', 周四, '6', 周五, '7', 周六, Null)
      Into v_时间段
      From 挂号安排计划
      Where ID = n_计划id;
    End If;
  Exception
    When Others Then
      v_时间段 := Null;
  End;

  If v_时间段 Is Not Null And d_启用时间 Is Not Null And 序号_In = 1 Then
    --检查是否跨模式挂号安排
    Select To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
           To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(终止时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
    Into d_检查开始时间, d_检查结束时间
    From 时间段
    Where 时间段 = v_时间段 And 站点 Is Null And 号类 Is Null;
    If d_检查开始时间 > d_检查结束时间 Then
      d_检查结束时间 := d_检查结束时间 + 1;
    End If;
    If d_检查开始时间 < d_启用时间 And d_检查结束时间 > d_启用时间 Then
      --获取出诊记录id
      Begin
        Select a.Id
        Into n_出诊记录id
        From 临床出诊记录 A, 临床出诊号源 B
        Where a.号源id = b.Id And b.号码 = 号别_In And 上班时段 = v_时间段 And 发生时间_In Between 开始时间 And 终止时间;
      Exception
        When Others Then
          n_出诊记录id := Null;
      End;
    End If;
  End If;

  --分时段挂号时判断是否是过期挂号 也就是追加号的情况 目前只针对专家号分时段进行处理
  If Nvl(预约挂号_In, 0) = 0 And n_分时段 > 0 And Nvl(n_序号控制, 0) = 1 And 号序_In Is Null And Nvl(操作类型_In, 0) = 0 Then
    --发生时间_in>Sysdate 发生时间>最大的时段时间--号序_in is null
    Begin
      Select Max(To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'))
      Into d_最大序号时间
      From 挂号安排时段
      Where 安排id = n_安排id And 星期 = v_星期 And Nvl(限制数量, 0) <> 0;
      n_追加号 := Case Sign(发生时间_In - d_最大序号时间)
                 When -1 Then
                  0
                 Else
                  1
               End;
    Exception
      When Others Then
        n_追加号 := 0;
    End;
  End If;
  d_时段时间 := 发生时间_In;

  If 序号_In = 1 And Nvl(预约挂号_In, 0) = 0 And n_分时段 > 0 Then
    --挂号时检查 是否分了时段,分了时段,把时段的限制条件给取出来
    Begin
      Select Nvl(序号, 0),
             To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
             限制数量, Decode(Nvl(是否预约, 0), 0, 0, 限制数量)
      Into n_时段序号, d_时段时间, n_时段限号, n_时段限约
      From 挂号安排时段
      Where 安排id = n_安排id And 星期 = v_星期 And
            (序号, 安排id, 星期) In (Select Nvl(Max(序号), -1), 安排id, 星期
                               From 挂号安排时段
                               Where 安排id = n_安排id And 星期 = v_星期 And
                                     Decode(操作类型_In + n_追加号, 0, To_Char(发生时间_In, 'hh24:mi'),
                                            To_Char(Nvl(开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')) =
                                     To_Char(Nvl(开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
                               Group By 安排id, 星期);
    Exception
      When Others Then
        n_时段序号 := -1;
        n_分时段   := 0;
        d_时段时间 := 发生时间_In;
        n_时段限号 := 0;
        n_时段限约 := 0;
    End;
  End If;

  If 序号_In = 1 And Nvl(预约挂号_In, 0) = 1 And n_分时段 > 0 Then
    --预约号,取计划
    Begin
      If Nvl(n_计划id, 0) = 0 Then
        --没计划生效,取安排的数据
        Select Nvl(序号, 0),
               To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(c.开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 时段时间,
               限制数量, Decode(Nvl(是否预约, 0), 0, 0, 限制数量)
        Into n_时段序号, d_时段时间, n_时段限号, n_时段限约
        From 挂号安排时段 C
        Where 安排id = n_安排id And 星期 = v_星期 And
              (序号, 安排id, 星期) In
              (Select Nvl(Max(c.序号), -1), 安排id, 星期
               From 挂号安排时段 C
               Where 安排id = n_安排id And c.星期 = v_星期 And
                     Decode(操作类型_In, 1, To_Char(Nvl(c.开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi'),
                            To_Char(发生时间_In, 'hh24:mi')) =
                     To_Char(Nvl(c.开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
               Group By 安排id, 星期);
      Else
        --有计划生效取计划
        --没生效，代表是从挂号计划时段查询
        Select Nvl(序号, -1),
               To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(c.开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 时段时间,
               限制数量, Decode(Nvl(是否预约, 0), 0, 0, 限制数量)
        Into n_时段序号, d_时段时间, n_时段限号, n_时段限约
        From 挂号计划时段 C
        Where 计划id = n_计划id And 星期 = v_星期 And
              (序号, 计划id, 星期) In
              (Select Nvl(Max(c.序号), -1), 计划id, 星期
               From 挂号计划时段 C
               Where 计划id = n_计划id And c.星期 = v_星期 And
                     Decode(操作类型_In, 1, To_Char(Nvl(c.开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi'),
                            To_Char(发生时间_In, 'hh24:mi')) =
                     To_Char(Nvl(c.开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
               Group By 计划id, 星期);
      End If;
    Exception
      When Others Then
        n_时段序号 := -1;
        n_分时段   := 0;
        d_时段时间 := 发生时间_In;
        n_时段限号 := 0;
        n_时段限约 := 0;
    End;
  End If;

  If 序号_In = 1 Then
  
    --获取当前未使用的序号
    If Nvl(n_序号控制, 0) = 1 And n_分时段 = 0 Then
      --<序号控制 未设置时段 获取可用的最大序号,以及已经使用的数量>
      Begin
        --最大序号
        If 退号重用_In = 1 Then
          Select Max(Nvl(序号, 0)), Sum(Decode(序号, Nvl(号序_In, 0), 0, 1)), Sum(Nvl(预约, 0))
          Into n_已用序号, n_已用数量, n_已约数
          From 挂号序号状态
          Where 号码 = 号别_In And 日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 状态 Not In (4, 5);
        Else
          Select Max(Nvl(序号, 0)), Sum(Decode(序号, Nvl(号序_In, 0), 0, 1)), Sum(Nvl(预约, 0))
          Into n_已用序号, n_已用数量, n_已约数
          From 挂号序号状态
          Where 号码 = 号别_In And 日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 状态 <> 5;
        End If;
      Exception
        When Others Then
          n_已用序号 := 0;
          n_已用数量 := 0;
      End;
      If n_序号 Is Null Then
        n_序号 := Nvl(n_已用序号, 0) + 1;
      End If;
      --<序号控制 未设置时段 获取可用的最大序号 以及已经使用的数量 --end>
    
      --非加号的情况需要检查是否超过了限制
      If 操作类型_In = 0 Then
        If Nvl(预约挂号_In, 0) = 0 Then
          --挂号时检查
          --启用序号控制未分时段 达到了限制
          If n_限号数 <= n_已用数量 And n_限号数 > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(Trunc(发生时间_In), 'yyyy-mm-dd ') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        Else
          --预约时检查
          If n_限约数 = 0 Then
            n_限约数 := n_限号数;
          End If;
        
          If n_限约数 <= n_已约数 And n_限约数 > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(Trunc(发生时间_In), 'yyyy-mm-dd') || '已达到最大限约数！';
            Raise Err_Item;
          End If;
        End If;
      
      Else
        Null;
        --启用序号控制,未分时段 加号情况   不处理,如果以后有限制条件以后补充
      End If;
    
    Elsif Nvl(n_分时段, 0) > 0 And Nvl(n_序号控制, 0) = 0 Then
      --<--普通号分时段 这里只有预约一种情况-->
      If 操作类型_In = 0 Then
        --<正常预约挂号-->
        Begin
          Select Count(0) As 已挂数, Nvl(Sum(Decode(Nvl(Sign(a.日期 - d_时段时间), 0), 0, 1, 0)), 0) As 已约数
          Into n_已挂数, n_已约数
          From 挂号序号状态 A
          Where a.号码 = 号别_In And 日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And
                状态 Not In (4, 5);
        Exception
          When Others Then
            n_已挂数 := 0;
            n_已约数 := 0;
        End;
      
        n_时段限约 := n_时段限号; --普通号分时段的情况,29中n_时段限约始终是0 这里特殊处理
        --检查限制数量
        If n_限约数 = 0 Then
          n_限约数 := n_限号数;
        End If;
        If n_时段限约 <= n_已约数 Or n_限约数 <= n_已约数 Then
          v_Err_Msg := '号别' || 号别_In || '在时段' || To_Char(d_时段时间, 'hh24:mi') || '已达到最大限约数！';
          Raise Err_Item;
        End If;
        If n_限号数 <= n_已挂数 Then
          v_Err_Msg := '号别' || 号别_In || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
          Raise Err_Item;
        End If;
      End If;
    
      --没有达到时段的限号数 号码在当前时段往后追加
    
      --获取当天挂出的最大号序
      If 退号重用_In = 1 Then
        Select Nvl(Max(序号), 0)
        Into n_挂出的最大序号
        From 挂号序号状态 A
        Where a.日期 = d_时段时间 And 号码 = 号别_In And 状态 Not In (4, 5);
      Else
        Select Nvl(Max(序号), 0)
        Into n_挂出的最大序号
        From 挂号序号状态 A
        Where a.日期 = d_时段时间 And 号码 = 号别_In And 状态 <> 5;
      End If;
    
      --设置序号
      n_序号 := RPad(Nvl(n_时段序号, 0), Length(n_限号数) + Length(Nvl(n_时段序号, 0)), 0) + n_已约数 + 1;
      If n_序号 <= Nvl(n_挂出的最大序号, 0) Then
        n_序号 := Nvl(n_挂出的最大序号, 0) + 1;
      End If;
    
      --<--普通号分时段--End>
    Elsif Nvl(n_分时段, 0) > 0 And Nvl(n_序号控制, 0) = 1 Then
      --<启用序号控制 设置时段
      --专家号分时段
      Begin
        If 退号重用_In = 1 Then
          Select Max(序号), Sum(Decode(序号, Nvl(号序_In, 0), 0, 1)), Sum(Decode(Sign(日期 - d_时段时间), 0, 1, 0))
          Into n_已用序号, n_已挂数, n_已用数量
          From 挂号序号状态
          Where 号码 = 号别_In And 日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 状态 Not In (4, 5);
        Else
          Select Max(序号), Sum(Decode(序号, Nvl(号序_In, 0), 0, 1)), Sum(Decode(Sign(日期 - d_时段时间), 0, 1, 0))
          Into n_已用序号, n_已挂数, n_已用数量
          From 挂号序号状态
          Where 号码 = 号别_In And 日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 状态 <> 5;
        End If;
      Exception
        When Others Then
          n_已用序号 := 0;
          n_已用数量 := 0;
          n_已挂数   := 0;
      End;
    
      n_失效数 := 0;
      If Nvl(预约挂号_In, 0) = 0 Then
        n_预约有效时间 := Zl_To_Number(zl_GetSysParameter('预约有效时间', 1111));
        n_失约挂号     := Zl_To_Number(zl_GetSysParameter('失约用于挂号', 1111));
        If Nvl(n_预约有效时间, 0) <> 0 And Nvl(n_失约挂号, 0) > 0 Then
          Begin
            Select Sum(Decode(Sign((Sysdate - n_预约有效时间 / 24 / 60) - 日期), 1, 1, 0))
            Into n_失效数
            From 挂号序号状态
            Where 号码 = 号别_In And 日期 Between Trunc(Sysdate) And Sysdate And Nvl(预约, 0) = 1 And 状态 = 2;
          Exception
            When Others Then
              n_失效数 := 0;
          End;
        End If;
      End If;
    
      If 操作类型_In = 0 Then
        If Nvl(预约挂号_In, 0) = 0 Then
        
          --挂号 追加号码时不检查时段限号数
          If n_时段限号 <= n_已用数量 And Nvl(n_追加号, 0) = 0 Then
            v_Err_Msg := '号别' || 号别_In || '在时段' || To_Char(d_时段时间, 'hh24:mi') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
          If n_限号数 <= n_已挂数 - n_失效数 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        
        Else
          --挂号
          If n_限约数 = 0 Then
            n_限约数 := n_限号数;
          End If;
          If n_限约数 <= n_已用数量 Then
            v_Err_Msg := '号别' || 号别_In || '在时段' || To_Char(d_时段时间, 'hh24:mi') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
          If n_限号数 <= n_已挂数 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        End If;
      End If;
      If n_序号 Is Null Then
        --设置序号
        If Nvl(n_已用序号, 0) < Nvl(n_时段序号, 0) Then
          n_已用序号 := Nvl(n_时段序号, 0);
        End If;
        n_序号 := Nvl(n_已用序号, 0) + 1;
      End If;
    Elsif Nvl(n_分时段, 0) = 0 And Nvl(n_序号控制, 0) = 0 And Nvl(病历费_In, 0) = 0 And Nvl(号别_In, 0) > 0 Then
      ---<--普通号  -->
      Begin
        Select 已挂数, 已约数
        Into n_已用数量, n_已约数
        From 病人挂号汇总
        Where 日期 = Trunc(发生时间_In) And 号码 = 号别_In;
      Exception
        When Others Then
          n_已用数量 := 0;
          n_已约数   := 0;
      End;
      If Nvl(操作类型_In, 0) = 0 Then
        If Nvl(预约挂号_In, 0) = 0 Then
          --挂号
          If Nvl(n_已用数量, 0) >= Nvl(n_限号数, 0) And Nvl(n_限号数, 0) > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        Else
          --预约
          If (Nvl(n_限约数, 0) > 0 And Nvl(n_已约数, 0) >= Nvl(n_限约数, 0)) Or
             (Nvl(n_已用数量, 0) >= Nvl(n_限号数, 0) And Nvl(n_限号数, 0) > 0) Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        End If;
      End If;
      ---<--普通号  -->
    End If;
  End If;

  --更新挂号序号状态
  If 序号_In = 1 And Not n_序号 Is Null Then
    If n_分时段 = 1 Then
      d_序号时间 := 发生时间_In;
    Else
      d_序号时间 := Trunc(发生时间_In);
    End If;
    --锁定序号的处理
    Begin
      Select 操作员姓名, 机器名
      Into v_序号操作员, v_序号机器名
      From 挂号序号状态
      Where 状态 = 5 And 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号;
      n_锁定 := 1;
    Exception
      When Others Then
        v_序号操作员 := Null;
        v_序号机器名 := Null;
        n_锁定       := 0;
    End;
    If n_锁定 = 0 Then
      Update 挂号序号状态
      Set 状态 = Decode(预约挂号_In, 1, 2, 1), 预约 = Decode(预约接收_In, 1, 1, 0), 登记时间 = Sysdate
      Where 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号 And 状态 = 3 And 操作员姓名 = 操作员姓名_In;
      If Sql%RowCount = 0 Then
        Begin
          If Nvl(n_分时段, 0) = 0 Or Nvl(预约挂号_In, 0) = 1 Or (Nvl(n_序号控制, 0) = 0 And Nvl(号序_In, 0) = 0) Then
            Insert Into 挂号序号状态
              (号码, 日期, 序号, 状态, 操作员姓名, 预约, 登记时间)
            Values
              (号别_In, d_序号时间, n_序号, Decode(预约挂号_In, 1, 2, 1), 操作员姓名_In, Decode(预约接收_In, 1, 1, 0), Sysdate);
            If Nvl(预约挂号_In, 0) = 0 And 预约方式_In Is Not Null Then
              Update 挂号序号状态 Set 预约 = 1 Where 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号;
            End If;
          Elsif Nvl(n_分时段, 0) > 0 Then
            --分时段后专家号 失约的预约号允许挂号
            Update 挂号序号状态
            Set 状态 = Decode(预约挂号_In, 1, 2, 1), 操作员姓名 = 操作员姓名_In, 预约 = Decode(预约接收_In, 1, 1, 0), 登记时间 = Sysdate
            Where 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号 And 状态 = 2;
            If Sql%NotFound Then
              Insert Into 挂号序号状态
                (号码, 日期, 序号, 状态, 操作员姓名, 预约, 登记时间)
              Values
                (号别_In, d_序号时间, n_序号, Decode(预约挂号_In, 1, 2, 1), 操作员姓名_In, Decode(预约接收_In, 1, 1, 0), Sysdate);
            End If;
            If Nvl(预约挂号_In, 0) = 0 And 预约方式_In Is Not Null Then
              Update 挂号序号状态 Set 预约 = 1 Where 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号;
            End If;
          End If;
        Exception
          When Others Then
            v_Err_Msg := '序号' || n_序号 || '已被使用,请重新选择一个序号.';
            Raise Err_Item;
        End;
      End If;
    Else
      If 操作员姓名_In <> v_序号操作员 Or v_机器名 <> v_序号机器名 Then
        v_Err_Msg := '序号' || n_序号 || '已被自助机' || v_机器名 || '锁定,请重新选择一个序号.';
        Raise Err_Item;
      Else
        Update 挂号序号状态
        Set 状态 = Decode(预约挂号_In, 1, 2, 1), 预约 = Decode(预约接收_In, 1, 1, 0), 登记时间 = Sysdate
        Where 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号 And 状态 = 5 And 操作员姓名 = 操作员姓名_In And 机器名 = v_机器名;
        If Nvl(预约挂号_In, 0) = 0 And 预约方式_In Is Not Null Then
          Update 挂号序号状态 Set 预约 = 1 Where 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号;
        End If;
      End If;
    End If;
  End If;

  If n_出诊记录id Is Not Null Then
    Update 临床出诊序号控制
    Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1), 操作员姓名 = 操作员姓名_In
    Where 记录id = n_出诊记录id And 序号 = n_序号;
    If 预约挂号_In = 1 Then
      Update 临床出诊记录 Set 已约数 = 已约数 + 1 Where ID = n_出诊记录id;
    Else
      If 预约接收_In = 1 Then
        Update 临床出诊记录
        Set 已约数 = 已约数 + 1, 已挂数 = 已挂数 + 1, 其中已接收 = 其中已接收 + 1
        Where ID = n_出诊记录id;
      Else
        Update 临床出诊记录 Set 已挂数 = 已挂数 + 1 Where ID = n_出诊记录id;
      End If;
    End If;
  End If;

  --产生病人挂号费用(可能单独是或包括病历费用)
  Select 病人费用记录_Id.Nextval Into n_费用id From Dual; --应该通过程序得到

  Insert Into 门诊费用记录
    (ID, 记录性质, 记录状态, 序号, 价格父号, 从属父号, NO, 实际票号, 门诊标志, 加班标志, 附加标志, 发药窗口, 病人id, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 病人科室id, 收费类别,
     计算单位, 收费细目id, 收入项目id, 收据费目, 付数, 数次, 标准单价, 应收金额, 实收金额, 结帐金额, 结帐id, 记帐费用, 开单部门id, 开单人, 划价人, 执行部门id, 执行人, 操作员编号, 操作员姓名,
     发生时间, 登记时间, 保险大类id, 保险项目否, 保险编码, 统筹金额, 摘要, 结论, 缴款组id)
  Values
    (n_费用id, 4, Decode(预约挂号_In, 1, 0, 1), 序号_In, Decode(价格父号_In, 0, Null, 价格父号_In), 从属父号_In, 单据号_In, 票据号_In, 1, 急诊_In,
     病历费_In, Decode(预约挂号_In, 1, To_Char(n_序号), 诊室_In), Decode(病人id_In, 0, Null, 病人id_In),
     Decode(门诊号_In, 0, Null, 门诊号_In), 付款方式_In, 姓名_In, Decode(姓名_In, Null, Null, 性别_In), Decode(姓名_In, Null, Null, 年龄_In),
     v_费别, 病人科室id_In, 收费类别_In, 号别_In, 收费细目id_In, 收入项目id_In, 收据费目_In, 1, 数次_In, 标准单价_In, 应收金额_In, 实收金额_In,
     Decode(预约挂号_In, 1, Null, Decode(Nvl(记帐费用_In, 0), 1, Null, 实收金额_In)),
     Decode(预约挂号_In, 1, Null, Decode(Nvl(记帐费用_In, 0), 1, Null, 结帐id_In)), Decode(Nvl(记帐费用_In, 0), 1, 1, 0), 开单部门id_In,
     操作员姓名_In, Decode(预约挂号_In, 1, 操作员姓名_In, Null), 执行部门id_In, 医生姓名_In, 操作员编号_In, 操作员姓名_In, 发生时间_In, 登记时间_In, 保险大类id_In,
     保险项目否_In, 保险编码_In, 统筹金额_In, Decode(收费单_In, Null, 摘要_In, '划价:' || 收费单_In), 预约方式_In, Decode(预约挂号_In, 1, Null, n_组id));

  --汇总结算到病人预交记录
  If Nvl(预约挂号_In, 0) = 0 And Nvl(记帐费用_In, 0) = 0 Then
  
    If (Nvl(现金支付_In, 0) <> 0 Or (Nvl(现金支付_In, 0) = 0 And Nvl(个帐支付_In, 0) = 0 And Nvl(预交支付_In, 0) = 0)) And 序号_In = 1 Then
      Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      Insert Into 病人预交记录
        (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
         结算性质)
      Values
        (n_预交id, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(结算方式_In, v_现金), Nvl(现金支付_In, 0), 登记时间_In,
         操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费', n_组id, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In, 4);
    
      If Nvl(结算卡序号_In, 0) <> 0 And Nvl(现金支付_In, 0) <> 0 Then
        Zl_病人卡结算记录_支付(结算卡序号_In, 卡号_In, 0, 现金支付_In, n_预交id, 操作员编号_In, 操作员姓名_In, 登记时间_In);
      End If;
    End If;
  
    --对于医保挂号
    If Nvl(个帐支付_In, 0) <> 0 And 序号_In = 1 Then
      Insert Into 病人预交记录
        (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 结算性质)
      Values
        (病人预交记录_Id.Nextval, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), v_个人帐户, 个帐支付_In, 登记时间_In, 操作员编号_In,
         操作员姓名_In, 结帐id_In, '医保挂号', n_组id, 4);
    End If;
  
    --对于就诊卡通过预交金挂号
    If Nvl(预交支付_In, 0) <> 0 And 序号_In = 1 Then
      n_预交金额 := 预交支付_In;
      For r_Deposit In c_Deposit(病人id_In, v_冲预交病人ids) Loop
        n_当前金额 := Case
                    When r_Deposit.金额 - n_预交金额 < 0 Then
                     r_Deposit.金额
                    Else
                     n_预交金额
                  End;
      
        If r_Deposit.结帐id = 0 Then
          --第一次冲预交(填上结帐ID,金额为0)
          Update 病人预交记录 Set 冲预交 = 0, 结帐id = 结帐id_In, 结算性质 = 4 Where ID = r_Deposit.原预交id;
        
        End If;
        --冲上次剩余额
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 预交类别, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,
           冲预交, 结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
          Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 预交类别, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号,
                 登记时间_In, 操作员姓名_In, 操作员编号_In, n_当前金额, 结帐id_In, n_组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
          From 病人预交记录
          Where NO = r_Deposit.No And 记录状态 = r_Deposit.记录状态 And 记录性质 In (1, 11) And Rownum = 1;
      
        --更新病人预交余额
        Update 病人余额
        Set 预交余额 = Nvl(预交余额, 0) - n_当前金额
        Where 病人id = r_Deposit.病人id And 性质 = 1 And 类型 = Nvl(1, 2);
        --检查是否已经处理完
        If r_Deposit.金额 < n_预交金额 Then
          n_预交金额 := n_预交金额 - r_Deposit.金额;
        Else
          n_预交金额 := 0;
        End If;
      
        If n_预交金额 = 0 Then
          Exit;
        End If;
      End Loop;
      If n_预交金额 > 0 Then
        v_Err_Msg := '预交余不够支付本次支付金额,不能继续操作！';
        Raise Err_Item;
      
      End If;
      Delete From 病人余额 Where 病人id = 病人id_In And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
    End If;
  
    --相关汇总表的处理
    --人员缴款余额
    If 序号_In = 1 And Nvl(更新交款余额_In, 1) = 1 Then
      If Nvl(现金支付_In, 0) <> 0 Then
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) + 现金支付_In
        Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = Nvl(结算方式_In, v_现金)
        Returning 余额 Into n_返回值;
      
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额
            (收款员, 结算方式, 性质, 余额)
          Values
            (操作员姓名_In, Nvl(结算方式_In, v_现金), 1, 现金支付_In);
          n_返回值 := 现金支付_In;
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 人员缴款余额
          Where 收款员 = 操作员姓名_In And 结算方式 = Nvl(结算方式_In, v_现金) And 性质 = 1 And Nvl(余额, 0) = 0;
        End If;
      End If;
    
      If Nvl(个帐支付_In, 0) <> 0 Then
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) + 个帐支付_In
        Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = v_个人帐户
        Returning 余额 Into n_返回值;
      
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (操作员姓名_In, v_个人帐户, 1, 个帐支付_In);
          n_返回值 := 个帐支付_In;
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 人员缴款余额
          Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_个人帐户 And Nvl(余额, 0) = 0;
        End If;
      End If;
    End If;
  End If;

  --病人挂号汇总(只处理一次,且单独收取病历费不处理)
  If Nvl(预约挂号_In, 0) = 0 Then
    --病人本次就诊(以发生时间为准)
    If Nvl(病人id_In, 0) <> 0 And 序号_In = 1 Then
      Update 病人信息 Set 就诊时间 = 发生时间_In, 就诊状态 = 1, 就诊诊室 = 诊室_In Where 病人id = 病人id_In;
    End If;
  End If;

  If Nvl(预约挂号_In, 0) = 0 And Nvl(记帐费用_In, 0) = 1 Then
    --记帐
    If Nvl(病人id_In, 0) = 0 Then
      v_Err_Msg := '要针对病人的挂号费进行记帐，必须是建档病人才能记帐挂号。';
      Raise Err_Item;
    End If;
  
    --病人余额
    Update 病人余额
    Set 费用余额 = Nvl(费用余额, 0) + Nvl(实收金额_In, 0)
    Where 病人id = Nvl(病人id_In, 0) And 性质 = 1 And 类型 = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人余额 (病人id, 性质, 类型, 费用余额, 预交余额) Values (病人id_In, 1, 1, Nvl(实收金额_In, 0), 0);
    End If;
  
    --病人未结费用
    Update 病人未结费用
    Set 金额 = Nvl(金额, 0) + Nvl(实收金额_In, 0)
    Where 病人id = 病人id_In And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(病人科室id, 0) = Nvl(病人科室id_In, 0) And
          Nvl(开单部门id, 0) = Nvl(开单部门id_In, 0) And Nvl(执行部门id, 0) = Nvl(执行部门id_In, 0) And 收入项目id + 0 = 收入项目id_In And
          来源途径 + 0 = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人未结费用
        (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
      Values
        (病人id_In, Null, Null, 病人科室id_In, 开单部门id_In, 执行部门id_In, 收入项目id_In, 1, Nvl(实收金额_In, 0));
    End If;
  End If;

  --病人挂号记录
  If 号别_In Is Not Null And 序号_In = 1 Then
    --And Nvl(预约挂号_In, 0) = 0
    Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
    Begin
      Select 名称 Into v_付款方式 From 医疗付款方式 Where 编码 = 付款方式_In And Rownum < 2;
    Exception
      When Others Then
        v_付款方式 := Null;
    End;
    Insert Into 病人挂号记录
      (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 预约时间, 操作员编号,
       操作员姓名, 复诊, 号序, 社区, 预约, 预约方式, 摘要, 交易流水号, 交易说明, 合作单位, 接收时间, 接收人, 预约操作员, 预约操作员编号, 险类, 医疗付款方式, 收费单)
    Values
      (n_挂号id, 单据号_In, Decode(Nvl(预约挂号_In, 0), 1, 2, 1), 1, 病人id_In, 门诊号_In, 姓名_In, 性别_In, 年龄_In, 号别_In, 急诊_In, 诊室_In,
       Null, 执行部门id_In, 医生姓名_In, 0, Null, 登记时间_In, 发生时间_In, Decode(Nvl(预约挂号_In, 0), 1, 发生时间_In, Null), 操作员编号_In,
       操作员姓名_In, 复诊_In, n_序号, 社区_In, Decode(预约接收_In, 1, 1, 0), 预约方式_In, 摘要_In, 交易流水号_In, 交易说明_In, 合作单位_In,
       Decode(Nvl(预约挂号_In, 0), 0, 登记时间_In, Null), Decode(Nvl(预约挂号_In, 0), 0, 操作员姓名_In, Null),
       Decode(Nvl(预约挂号_In, 0), 1, 操作员姓名_In, Null), Decode(Nvl(预约挂号_In, 0), 1, 操作员编号_In, Null),
       Decode(Nvl(险类_In, 0), 0, Null, 险类_In), v_付款方式, 收费单_In);
    If Nvl(预约挂号_In, 0) = 0 And 预约方式_In Is Not Null Then
      Update 病人挂号记录
      Set 预约 = 1, 预约时间 = 发生时间_In, 预约操作员 = 操作员姓名_In, 预约操作员编号 = 操作员编号_In
      Where ID = n_挂号id;
    End If;
    n_预约生成队列 := 0;
    If Nvl(预约挂号_In, 0) = 1 Then
      n_预约生成队列 := Zl_To_Number(zl_GetSysParameter('预约生成队列', 1113));
    End If;
  
    --0-不产生队列;1-按医生或分诊台排队;2-先分诊,后医生站
    If Nvl(生成队列_In, 0) <> 0 And Nvl(预约挂号_In, 0) = 0 Or n_预约生成队列 = 1 Then
      n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113, 1, Nvl(执行部门id_In, 0)));
      If Nvl(n_分诊台签到排队, 0) = 0 Or n_预约生成队列 = 1 Then
        n_分时点显示 := Nvl(Zl_To_Number(zl_GetSysParameter(270)), 0);
        If Nvl(预约挂号_In, 0) = 1 And n_分时点显示 = 1 And n_分时段 = 1 Then
          n_分时点显示 := 1;
        Else
          n_分时点显示 := Null;
        End If;
      
        --产生队列
        --.按”执行部门” 的方式生成队列
        v_队列名称 := 执行部门id_In;
        v_排队号码 := Zlgetnextqueue(执行部门id_In, n_挂号id, 号别_In || '|' || n_序号);
        v_排队序号 := Zlgetsequencenum(0, n_挂号id, 0);
        d_排队时间 := Zl_Get_Queuedate(n_挂号id, 号别_In, n_序号, 发生时间_In);
        --  队列名称_In , 业务类型_In, 业务id_In,科室id_In,排队号码_In,排队标记_In,患者姓名_In,病人ID_IN, 诊室_In, 医生姓名_In, 排队时间_In
        Zl_排队叫号队列_Insert(v_队列名称, 0, n_挂号id, 执行部门id_In, v_排队号码, Null, 姓名_In, 病人id_In, 诊室_In, 医生姓名_In, d_排队时间, 预约方式_In,
                         n_分时点显示, v_排队序号);
      
        --挂号立即排队
        If Nvl(n_分诊台签到排队, 0) = 0 Then
          Update 病人挂号记录 Set 记录标志 = 1 Where ID = n_挂号id;
        End If;
      End If;
    End If;
  End If;
  --病人担保信息
  If 病人id_In Is Not Null And 序号_In = 1 Then
    --取参数:
    If Nvl(n_结算模式, 0) <> Nvl(结算模式_In, 0) Then
      --结算模式的确定
    
      v_Err_Msg := Null;
      Begin
        Select Nvl(结算模式, 0) Into n_结算模式 From 病人信息 Where 病人id = 病人id_In;
      Exception
        When Others Then
          v_Err_Msg := '未找到指定的病人信息,不允许挂号';
      End;
    
      If v_Err_Msg Is Not Null Then
        Raise Err_Item;
      End If;
      If n_结算模式 = 1 And Nvl(结算模式_In, 0) = 0 Then
        --病人已经是"先诊疗后结算的",本次是"先结算后诊疗的",则检查是否存在未结数据
        Select Count(1)
        Into n_Count
        From 病人未结费用
        Where 病人id = 病人id_In And (来源途径 In (1, 4) Or 来源途径 = 3 And Nvl(主页id, 0) = 0) And Nvl(金额, 0) <> 0 And Rownum < 2;
        If Nvl(n_Count, 0) <> 0 Then
          --存在未结算数据，必须先结算后才允许执行
          v_Err_Msg := '当前病人的就诊模式为先诊疗后结算且存在未结费用，不允许调整该病人的就诊模式,你可以先对未结费用结帐后再挂号或不调整病人的就诊模式!';
          Raise Err_Item;
        End If;
        --检查
        --未发生医嘱业务的（即当时就挂号的,需要保证同一次的就诊模式是一至的(程序已经检查，不用再处理)
      End If;
      Update 病人信息 Set 结算模式 = 结算模式_In Where 病人id = 病人id_In;
    End If;
  
    Update 病人信息
    Set 担保人 = Null, 担保额 = Null, 担保性质 = Null
    Where 病人id = 病人id_In And Nvl(在院, 0) = 0 And Exists
     (Select 1
           From 病人担保记录
           Where 病人id = 病人id_In And 主页id Is Not Null And
                 登记时间 = (Select Max(登记时间) From 病人担保记录 Where 病人id = 病人id_In));
  
    If Sql%RowCount > 0 Then
      Update 病人担保记录
      Set 到期时间 = Sysdate
      Where 病人id = 病人id_In And 主页id Is Not Null And Nvl(到期时间, Sysdate) >= Sysdate;
    End If;
  End If;
  If 序号_In = 1 Then
    --消息推送
    Begin
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 1, n_挂号id;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_001(n_挂号id, 单据号_In);
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人挂号记录_Insert;
/

--127075:焦博,2018-07-06,调整使用到了部门参数分诊台签到排队的Oracle函数
Create Or Replace Procedure Zl_病人挂号记录_Delete
(
  单据号_In       门诊费用记录.No%Type,
  操作员编号_In   门诊费用记录.操作员编号%Type,
  操作员姓名_In   门诊费用记录.操作员姓名%Type,
  摘要_In         门诊费用记录.摘要%Type := Null, --预约取消时 填写 存放预约取消原因
  删除门诊号_In   Number := 0,
  非原样退结算_In Varchar2 := Null,
  退费类型_In     In Number := 0, --0-全退 1-退挂号费 2-退病历费 3-退附加费 4-退挂号与病历 5-退挂号与附加
  退指定结算_In   病人预交记录.结算方式%Type := Null,
  退号重用_In     Number := 1,
  收回票据号_In   Varchar2 := Null,
  交易说明_In     病人预交记录.交易说明%Type := Null
) As
  --退费类型_In,在一下几种情况下不准进行部分退费
  --    2.三方接口,暂时不支持
  -- 挂号费病历费分开退,规则
  --    普通结算方式:原结算方式退部分费用
  --    预交款:预交款,退部分
  --    预交款与普通结算方式混合:退款按照普通结算方式部分退
  --    消费卡:原样将费用部分退入消费卡
  --非原样退结算_In:指不能退还给原样结算方式(如医保的个人账户,三方账户的退现等),多个用逗分离
  --退指定结算_IN:指非原样退结算部分,应该退给哪种结算方式,为空时缺省退给现金,否则退给指定的结算方式

  --该游标用于判断是否单独收病历费,及挂号汇总表处理
  Cursor c_Registinfo(v_状态 门诊费用记录.记录状态%Type) Is
    Select a.发生时间, a.登记时间, c.接收时间, a.收费细目id As 项目id, c.执行部门id As 科室id, c.执行人 As 医生姓名, d.Id As 医生id, c.号别 As 号码
    From 门诊费用记录 A, 挂号安排 B, 病人挂号记录 C, 人员表 D
    Where a.记录性质 = 4 And a.序号 = 1 And a.记录状态 = v_状态 And c.No = a.No And c.执行人 = d.姓名(+) And a.No = 单据号_In And
          Nvl(a.计算单位, '号别') = c.号别 And Rownum < 2;
  r_Registrow c_Registinfo%RowType;

  --该游标用于判断记录是否存在,及费用汇总表处理
  Cursor c_Moneyinfo(v_状态 门诊费用记录.记录状态%Type) Is
    Select 病人科室id, 开单部门id, 执行部门id, 收入项目id, Nvl(Sum(应收金额), 0) As 应收, Nvl(Sum(实收金额), 0) As 实收, Nvl(Sum(结帐金额), 0) As 结帐
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 = v_状态 And NO = 单据号_In
    Group By 病人科室id, 开单部门id, 执行部门id, 收入项目id;
  r_Moneyrow c_Moneyinfo%RowType;

  --该光标用于处理人员缴款余额中退的不同结算方式的金额
  Cursor c_Opermoney(n_Id 病人预交记录.结帐id%Type) Is
    Select Distinct b.结算方式, -1 * Nvl(b.冲预交, 0) As 冲预交
    From 病人预交记录 B
    Where b.结帐id = n_Id And b.记录性质 = 4 And b.记录状态 = 2 And Nvl(b.冲预交, 0) <> 0;

  v_Err_Msg Varchar(255);
  Err_Item Exception;

  n_结帐id 病人预交记录.结帐id%Type;
  n_销帐id 门诊费用记录.结帐id%Type;

  v_退指定结算方式 病人预交记录.结算方式%Type;
  n_退款金额       病人预交记录.冲预交%Type;
  n_打印id         票据打印内容.Id%Type;
  n_病人id         病人信息.病人id%Type;
  n_退费金额       病人预交记录.冲预交%Type;
  n_预交金额       病人预交记录.冲预交%Type; --原记录 预交缴款金额
  n_返回值         病人余额.预交余额%Type;
  n_挂号id         病人挂号记录.Id%Type;
  n_组id           财务缴款分组.Id%Type;

  n_二次退费       Number; --记录是否是此单据的第二次退费
  n_分诊台签到排队 Number;
  n_预约生成队列   Number;
  n_预约挂号       Number;
  n_挂号生成队列   Number;
  d_Date           Date;
  n_记帐           门诊费用记录.记帐费用%Type;
  n_病人id1        病人信息.病人id%Type;
  n_返回额         门诊费用记录.实收金额%Type;
  n_已结帐         Number;
  n_安排id         挂号安排.Id%Type;
  n_计划id         挂号安排计划.Id%Type;
  d_启用时间       Date;
  d_发生时间       病人挂号记录.发生时间%Type;
  n_出诊记录id     临床出诊记录.Id%Type;
  v_号码           挂号安排.号码%Type;
  n_序号           病人挂号记录.号序%Type;
  v_时间段         时间段.时间段%Type;
  d_检查开始时间   Date;
  d_检查结束时间   Date;
  v_Temp           Varchar2(500);
  v_附加ids        Varchar2(500);
  n_就诊病人id     病人信息.病人id%Type;
  d_就诊时间       就诊登记记录.就诊时间%Type;
Begin
  n_组id           := Zl_Get组id(操作员姓名_In);
  v_退指定结算方式 := 退指定结算_In;

  --首先判断要退号/取消预约的记录是否存在
  Open c_Moneyinfo(1);
  Fetch c_Moneyinfo
    Into r_Moneyrow;
  If c_Moneyinfo%RowCount = 0 Then
    Close c_Moneyinfo;
    Open c_Moneyinfo(0);
    Fetch c_Moneyinfo
      Into r_Moneyrow;
    If c_Moneyinfo%RowCount = 0 Then
      v_Err_Msg := '要处理的单据不存在。';
      Raise Err_Item;
    End If;
    n_预约挂号 := 1;
  End If;
  Close c_Moneyinfo;

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

  Select 号别, 号序, 发生时间 Into v_号码, n_序号, d_发生时间 From 病人挂号记录 Where NO = 单据号_In And Rownum < 2;

  Begin
    Select a.Id Into n_安排id From 挂号安排 A Where a.号码 = v_号码;
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
           Where a.审核时间 Is Not Null And d_发生时间 Between Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                 Nvl(a.失效时间, To_Date('3000-01-01', 'yyyy-mm-dd')) And a.安排id = n_安排id) And
          d_发生时间 Between Nvl(生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
          Nvl(失效时间, To_Date('3000-01-01', 'yyyy-mm-dd'));
  Exception
    When Others Then
      n_计划id := 0;
  End;

  Begin
    If Nvl(n_计划id, 0) = 0 Then
      Select Decode(To_Char(d_发生时间, 'D'), '1', 周日, '2', 周一, '3', 周二, '4', 周三, '5', 周四, '6', 周五, '7', 周六, Null)
      Into v_时间段
      From 挂号安排
      Where ID = n_安排id;
    Else
      Select Decode(To_Char(d_发生时间, 'D'), '1', 周日, '2', 周一, '3', 周二, '4', 周三, '5', 周四, '6', 周五, '7', 周六, Null)
      Into v_时间段
      From 挂号安排计划
      Where ID = n_计划id;
    End If;
  Exception
    When Others Then
      v_时间段 := Null;
  End;

  If v_时间段 Is Not Null And d_启用时间 Is Not Null Then
    --检查是否跨模式挂号安排
    Select To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
           To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || ' ' || To_Char(终止时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
    Into d_检查开始时间, d_检查结束时间
    From 时间段
    Where 时间段 = v_时间段 And 站点 Is Null And 号类 Is Null;
    If d_检查开始时间 > d_检查结束时间 Then
      d_检查结束时间 := d_检查结束时间 + 1;
    End If;
    If d_检查结束时间 > d_启用时间 Then
      --获取出诊记录id
      Begin
        Select a.Id
        Into n_出诊记录id
        From 临床出诊记录 A, 临床出诊号源 B
        Where a.号源id = b.Id And b.号码 = v_号码 And 上班时段 = v_时间段 And d_发生时间 Between 开始时间 And 终止时间;
      Exception
        When Others Then
          n_出诊记录id := Null;
      End;
    End If;
  End If;

  Begin
    Select Zl_Fun_Regcustomname Into v_Temp From Dual;
    If v_Temp Is Not Null Then
      v_附加ids := Substr(v_Temp, Instr(v_Temp, '|') + 1);
    End If;
  Exception
    When Others Then
      v_附加ids := Null;
  End;

  --1.预约处理
  If Nvl(n_预约挂号, 0) = 1 Then
    --减少已约数
    Open c_Registinfo(0);
    Fetch c_Registinfo
      Into r_Registrow;
  
    Update 病人挂号汇总
    Set 已约数 = Nvl(已约数, 0) - 1
    Where 日期 = Trunc(r_Registrow.发生时间) And 科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And
          (号码 = r_Registrow.号码 Or 号码 Is Null);
  
    If Sql%RowCount = 0 Then
      Insert Into 病人挂号汇总
        (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已约数)
      Values
        (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
         Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, -1);
    End If;
    Close c_Registinfo;
  
    --更新挂号序号状态
    Delete 挂号序号状态
    Where 状态 = 2 And
          (号码, 序号, 日期) = (Select 计算单位, 发药窗口, Trunc(发生时间)
                          From 门诊费用记录
                          Where 记录性质 = 4 And 记录状态 = 0 And 序号 = 1 And Rownum = 1 And NO = 单据号_In) Or
          (号码, 序号, 日期) = (Select 计算单位, 发药窗口, 发生时间
                          From 门诊费用记录
                          Where 记录性质 = 4 And 记录状态 = 0 And 序号 = 1 And Rownum = 1 And NO = 单据号_In);
  
    --添加病人挂号记录的 冲销记录
    Select 病人挂号记录_Id.Nextval, Sysdate Into n_挂号id, d_Date From Dual;
  
    Update 病人挂号记录 Set 记录状态 = 3 Where NO = 单据号_In And 记录状态 = 1 And 记录性质 = 2;
    If Sql%NotFound Then
      v_Err_Msg := '预约单【' || 单据号_In || '】不存在或由于并发原因已经被取消预约';
      Raise Err_Item;
    End If;
  
    Insert Into 病人挂号记录
      (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 操作员编号, 操作员姓名,
       复诊, 号序, 社区, 预约, 摘要, 预约方式, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 医疗付款方式)
      Select n_挂号id, NO, 记录性质, 2, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, d_Date, 发生时间,
             操作员编号_In, 操作员姓名_In, 复诊, 号序, 社区, 预约, Nvl(摘要_In, 摘要) As 摘要, 预约方式, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 医疗付款方式
      From 病人挂号记录
      Where NO = 单据号_In;
  
    If n_出诊记录id Is Not Null Then
      Update 临床出诊记录 Set 已约数 = 已约数 - 1 Where ID = n_出诊记录id And Nvl(已约数, 0) > 0;
      Update 临床出诊序号控制 Set 挂号状态 = Null, 操作员姓名 = Null Where 记录id = n_出诊记录id And 序号 = n_序号;
    End If;
  
    --Update 病人挂号记录 set 摘要=nvl(摘要_IN,摘要) where NO=单据号_IN;
    --删除门诊费用记录
    Delete From 门诊费用记录 Where NO = 单据号_In And 记录性质 = 4 And 记录状态 = 0;
    --如果预约生成队列时需要清除队列
  
    n_预约生成队列 := Zl_To_Number(zl_GetSysParameter('预约生成队列', 1113));
    If Nvl(n_预约生成队列, 0) = 1 Then
      --要删除队列
      For v_挂号 In (Select ID, 诊室, 执行部门id, 执行人 From 病人挂号记录 Where NO = 单据号_In) Loop
        Zl_排队叫号队列_Delete(v_挂号.执行部门id, v_挂号.Id);
      End Loop;
    End If;
    Return;
  End If;

  Select Nvl(记帐费用, 0), 病人id, Decode(Sign(Nvl(结帐id, 0)), 0, 0, 1)
  Into n_记帐, n_病人id, n_已结帐
  From 门诊费用记录
  Where 记录性质 = 4 And NO = 单据号_In And 记录状态 In (1, 3) And Rownum < 2;

  --2.挂号处理
  n_已结帐 := Nvl(n_已结帐, 0);

  If n_已结帐 = 1 And n_记帐 = 1 Then
    Select Sysdate, Null Into d_Date, n_销帐id From Dual;
  Else
    Select Sysdate, 病人结帐记录_Id.Nextval Into d_Date, n_销帐id From Dual;
  End If;

  ----0-全退 1-退挂号费 2-退病历费
  If Nvl(退费类型_In, 0) Not In (2, 3) Then
    --不是光退病历费时处理
    --更新挂号序号状态
    If 退号重用_In = 1 Then
      Delete 挂号序号状态
      Where 状态 = 1 And
            (号码, 序号, 日期) = (Select 号别, 号序, Trunc(发生时间) From 病人挂号记录 Where NO = 单据号_In And Rownum = 1) Or
            (号码, 序号, 日期) = (Select 号别, 号序, 发生时间 From 病人挂号记录 Where NO = 单据号_In And Rownum = 1);
    Else
      Update 挂号序号状态
      Set 状态 = 4
      Where 状态 = 1 And
            (号码, 序号, 日期) = (Select 号别, 号序, Trunc(发生时间) From 病人挂号记录 Where NO = 单据号_In And Rownum = 1) Or
            (号码, 序号, 日期) = (Select 号别, 号序, 发生时间 From 病人挂号记录 Where NO = 单据号_In And Rownum = 1);
    End If;
  
    --病人就诊状态
    If n_病人id Is Not Null Then
      Update 病人信息 Set 就诊状态 = 0, 就诊诊室 = Null Where 病人id = n_病人id;
    
      --删除门诊号相关处理,只有当只有一条挂号记录并且病人建档日期与挂号日期近似时才会处理
      If 删除门诊号_In = 1 Then
        Delete 门诊病案记录 Where 病人id = n_病人id;
        Update 病人信息 Set 门诊号 = Null Where 病人id = n_病人id;
        --费用记录包括挂号及病案、就诊卡费用,以及病人交费后退费或销帐的费用,挂号记录在最后处理
        Update 门诊费用记录 Set 标识号 = Null Where 门诊标志 = 1 And 病人id = n_病人id;
      End If;
    End If;
  
    --如果挂时收了就诊卡费,退费时清除就诊卡号,在非光退病历费时
    n_病人id1 := Null;
    Begin
      Select 病人id
      Into n_病人id1
      From 门诊费用记录
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And 附加标志 = 2 And Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
  
    If n_病人id1 Is Not Null And Nvl(退费类型_In, 0) Not In (2, 3) Then
      Update 病人信息
      Set 就诊卡号 = Null, 卡验证码 = Null, Ic卡号 = Decode(Ic卡号, 就诊卡号, Null, Ic卡号)
      Where 病人id = n_病人id1;
    End If;
  
  End If;

  --检查前面是否已经部分退过费用
  Begin
    Select 1 Into n_二次退费 From 门诊费用记录 Where 记录性质 = 4 And NO = 单据号_In And 记录状态 = 3 And Rownum < 2;
  Exception
    When Others Then
      n_二次退费 := 0;
  End;

  If Nvl(退费类型_In, 0) = 0 Or Nvl(退费类型_In, 0) = 2 Then
    --全退,退病历费
    --门诊费用记录，冲销记录
    Insert Into 门诊费用记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位, 付数,
       数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 执行人, 操作员编号, 操作员姓名, 发生时间, 登记时间,
       结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 缴款组id)
      Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别,
             收费细目id, 计算单位, 付数, -数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, -应收金额, -实收金额, 开单部门id, 开单人, 执行部门id, 执行人,
             操作员编号_In, 操作员姓名_In, 发生时间, d_Date, n_销帐id,
             Decode(n_记帐, 1, Decode(Nvl(n_已结帐, 0), 0, -1 * 实收金额, Null), -1 * 结帐金额), 保险项目否, 保险大类id, -1 * 统筹金额,
             Nvl(摘要_In, 摘要) As 摘要, Decode(Nvl(附加标志, 0), 9, 1, 0), 保险编码, 费用类型, n_组id
      From 门诊费用记录
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
  
    --原始记录
    If n_记帐 = 1 And Nvl(n_已结帐, 0) = 0 Then
      Update 门诊费用记录
      Set 记录状态 = 3, 结帐id = n_销帐id, 结帐金额 = 实收金额
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
    Else
      Update 门诊费用记录
      Set 记录状态 = 3
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
    End If;
  Elsif Nvl(退费类型_In, 0) = 1 Or Nvl(退费类型_In, 0) = 4 Then
    --退挂号费,退挂号与病历费
    --门诊费用记录，冲销记录
    Insert Into 门诊费用记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位, 付数,
       数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 执行人, 操作员编号, 操作员姓名, 发生时间, 登记时间,
       结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 缴款组id)
      Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别,
             收费细目id, 计算单位, 付数, -数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, -应收金额, -实收金额, 开单部门id, 开单人, 执行部门id, 执行人,
             操作员编号_In, 操作员姓名_In, 发生时间, d_Date, n_销帐id,
             Decode(n_记帐, 1, Decode(Nvl(n_已结帐, 0), 0, -1 * 实收金额, Null), -1 * 结帐金额), 保险项目否, 保险大类id, -1 * 统筹金额,
             Nvl(摘要_In, 摘要) As 摘要, Decode(Nvl(附加标志, 0), 9, 1, 0), 保险编码, 费用类型, n_组id
      From 门诊费用记录
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Instr(',' || v_附加ids || ',', ',' || 收费细目id || ',') = 0 And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
  
    --原始记录
    If n_记帐 = 1 And Nvl(n_已结帐, 0) = 0 Then
      Update 门诊费用记录
      Set 记录状态 = 3, 结帐id = n_销帐id, 结帐金额 = 实收金额
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Instr(',' || v_附加ids || ',', ',' || 收费细目id || ',') = 0 And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
    Else
      Update 门诊费用记录
      Set 记录状态 = 3
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Instr(',' || v_附加ids || ',', ',' || 收费细目id || ',') = 0 And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
    End If;
  Elsif Nvl(退费类型_In, 0) = 3 Then
    --退附加费
    --门诊费用记录，冲销记录
    Insert Into 门诊费用记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位, 付数,
       数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 执行人, 操作员编号, 操作员姓名, 发生时间, 登记时间,
       结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 缴款组id)
      Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别,
             收费细目id, 计算单位, 付数, -数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, -应收金额, -实收金额, 开单部门id, 开单人, 执行部门id, 执行人,
             操作员编号_In, 操作员姓名_In, 发生时间, d_Date, n_销帐id,
             Decode(n_记帐, 1, Decode(Nvl(n_已结帐, 0), 0, -1 * 实收金额, Null), -1 * 结帐金额), 保险项目否, 保险大类id, -1 * 统筹金额,
             Nvl(摘要_In, 摘要) As 摘要, Decode(Nvl(附加标志, 0), 9, 1, 0), 保险编码, 费用类型, n_组id
      From 门诊费用记录
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Instr(',' || v_附加ids || ',', ',' || 收费细目id || ',') > 0 And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
  
    --原始记录
    If n_记帐 = 1 And Nvl(n_已结帐, 0) = 0 Then
      Update 门诊费用记录
      Set 记录状态 = 3, 结帐id = n_销帐id, 结帐金额 = 实收金额
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Instr(',' || v_附加ids || ',', ',' || 收费细目id || ',') > 0 And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
    Else
      Update 门诊费用记录
      Set 记录状态 = 3
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Instr(',' || v_附加ids || ',', ',' || 收费细目id || ',') > 0 And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
    End If;
  Elsif Nvl(退费类型_In, 0) = 5 Then
    --退挂号与附加费
    --门诊费用记录，冲销记录
    Insert Into 门诊费用记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位, 付数,
       数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 执行人, 操作员编号, 操作员姓名, 发生时间, 登记时间,
       结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 缴款组id)
      Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别,
             收费细目id, 计算单位, 付数, -数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, -应收金额, -实收金额, 开单部门id, 开单人, 执行部门id, 执行人,
             操作员编号_In, 操作员姓名_In, 发生时间, d_Date, n_销帐id,
             Decode(n_记帐, 1, Decode(Nvl(n_已结帐, 0), 0, -1 * 实收金额, Null), -1 * 结帐金额), 保险项目否, 保险大类id, -1 * 统筹金额,
             Nvl(摘要_In, 摘要) As 摘要, Decode(Nvl(附加标志, 0), 9, 1, 0), 保险编码, 费用类型, n_组id
      From 门诊费用记录
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Nvl(附加标志, 0) <> 1;
  
    --原始记录
    If n_记帐 = 1 And Nvl(n_已结帐, 0) = 0 Then
      Update 门诊费用记录
      Set 记录状态 = 3, 结帐id = n_销帐id, 结帐金额 = 实收金额
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Nvl(附加标志, 0) <> 1;
    Else
      Update 门诊费用记录
      Set 记录状态 = 3
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Nvl(附加标志, 0) <> 1;
    End If;
  End If;

  n_结帐id := 0;
  If n_记帐 = 0 Then
    --获取结帐ID
    Select Nvl(结帐id, 0)
    Into n_结帐id
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 = 3 And NO = 单据号_In And Rownum < 2;
  End If;

  If n_记帐 = 1 Then
    --记帐
    For c_费用 In (Select 实收金额, 病人科室id, 开单部门id, 执行部门id, 收入项目id
                 From 门诊费用记录
                 Where 记录性质 = 4 And 记录状态 = 3 And NO = 单据号_In And
                       Nvl(附加标志, 0) =
                       Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0)) And
                       Nvl(记帐费用, 0) = 1) Loop
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
            Nvl(开单部门id, 0) = Nvl(c_费用.开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(c_费用.执行部门id, 0) And 收入项目id + 0 = c_费用.收入项目id And
            来源途径 + 0 = 1;
    
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
    --1.退费
    --病人挂号结算:现金和个人帐户部份
    If 非原样退结算_In Is Not Null Then
      --退款金额获取
      If Nvl(退费类型_In, 0) <> 0 Or Nvl(n_二次退费, 0) <> 0 Then
        --如果是单独退病历费,或者只退挂号费,先获取退费金额
        Begin
          --获取本次退款金额
          Select -1 * Sum(Nvl(实收金额, 0)) As 收款金额
          Into n_退款金额
          From 门诊费用记录
          Where NO = 单据号_In And 记录性质 = 4 And 记录状态 = 2 And 结帐id = n_销帐id;
        
        Exception
          When Others Then
            v_Err_Msg := '单据【' || 单据号_In || '】的' || Case Nvl(退费类型_In, 0)
                           When 1 Then
                            '挂号费用'
                           When 2 Then
                            '病历费'
                         End || '可能由于并发原因已经进行了退费或者单据不存在!';
            Raise Err_Item;
        End;
        Begin
          Select 冲预交
          Into n_退费金额
          From 病人预交记录
          Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id And Instr(',' || 非原样退结算_In || ',', ',' || 结算方式 || ',') = 0;
        Exception
          When Others Then
            n_退费金额 := 0;
        End;
      
        --a.允许的结算方式
      
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
           结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
          Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In, -n_退款金额,
                 n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
          From 病人预交记录
          Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id And Instr(',' || 非原样退结算_In || ',', ',' || 结算方式 || ',') = 0;
      
        If n_退费金额 = 0 Then
          --b.不允许的退现金
          If n_退款金额 <> 0 Then
            If v_退指定结算方式 Is Null Then
              --退给现金
              Begin
                Select 名称 Into v_退指定结算方式 From 结算方式 Where 性质 = 1;
              Exception
                When Others Then
                  v_退指定结算方式 := '现金';
              End;
            End If;
            Update 病人预交记录
            Set 冲预交 = 冲预交 + (-1 * n_退款金额), 交易说明 = Nvl(交易说明_In, 交易说明)
            Where 记录性质 = 4 And 记录状态 = 2 And 结帐id = n_销帐id And 结算方式 = v_退指定结算方式;
            If Sql%RowCount = 0 Then
              Insert Into 病人预交记录
                (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别,
                 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
                Select 病人预交记录_Id.Nextval, a.No, a.实际票号, a.记录性质, 2, a.病人id, a.主页id, a.科室id, a.摘要, v_退指定结算方式, d_Date,
                       操作员编号_In, 操作员姓名_In, -1 * n_退款金额, n_销帐id, n_组id, 预交类别, Decode(交易说明_In, Null, 卡类别id, Null), 结算卡序号,
                       Decode(交易说明_In, Null, 卡号, Null), Decode(交易说明_In, Null, 交易流水号, Null), Nvl(交易说明_In, 交易说明), 合作单位, 4
                From 病人预交记录 A
                Where a.记录性质 = 4 And a.记录状态 = 1 And a.结帐id = n_结帐id And Rownum < 2;
            End If;
          End If;
        End If;
      Else
        --a.允许的结算方式原样退
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
           结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
          Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In, -冲预交,
                 n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
          From 病人预交记录
          Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id And Instr(',' || 非原样退结算_In || ',', ',' || 结算方式 || ',') = 0;
      
        --b.不允许的退现金
        Begin
          Select Sum(冲预交)
          Into n_退费金额
          From 病人预交记录
          Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id And Instr(',' || 非原样退结算_In || ',', ',' || 结算方式 || ',') > 0;
        Exception
          When Others Then
            n_退费金额 := 0;
        End;
        If n_退费金额 <> 0 Then
          If v_退指定结算方式 Is Null Then
            --退给现金
            Begin
              Select 结算方式
              Into v_退指定结算方式
              From 病人预交记录
              Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id And Instr(',' || 非原样退结算_In || ',', ',' || 结算方式 || ',') = 0;
            
            Exception
              When Others Then
                Begin
                  Select 名称 Into v_退指定结算方式 From 结算方式 Where 性质 = 1;
                Exception
                  When Others Then
                    v_退指定结算方式 := '现金';
                End;
            End;
          End If;
          Update 病人预交记录
          Set 冲预交 = 冲预交 + (-1 * n_退费金额), 交易说明 = Nvl(交易说明_In, 交易说明)
          Where 记录性质 = 4 And 记录状态 = 2 And 结帐id = n_销帐id And 结算方式 = v_退指定结算方式;
          If Sql%RowCount = 0 Then
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
               结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
              Select 病人预交记录_Id.Nextval, a.No, a.实际票号, a.记录性质, 2, a.病人id, a.主页id, a.科室id, a.摘要, v_退指定结算方式, d_Date,
                     操作员编号_In, 操作员姓名_In, -1 * n_退费金额, n_销帐id, n_组id, 预交类别, Decode(交易说明_In, Null, 卡类别id, Null), 结算卡序号,
                     Decode(交易说明_In, Null, 卡号, Null), Decode(交易说明_In, Null, 交易流水号, Null), Nvl(交易说明_In, 交易说明), 合作单位, 4
              From 病人预交记录 A
              Where a.记录性质 = 4 And a.记录状态 = 1 And a.结帐id = n_结帐id And Rownum < 2;
          End If;
        End If;
      End If;
    Else
      --退款金额获取
      If Nvl(退费类型_In, 0) <> 0 Or Nvl(n_二次退费, 0) <> 0 Then
        --如果是单独退病历费,或者只退挂号费,先获取退费金额
        Begin
          --获取本次退款金额
          Select -1 * Sum(Nvl(实收金额, 0)) As 收款金额
          Into n_退款金额
          From 门诊费用记录
          Where NO = 单据号_In And 记录性质 = 4 And 记录状态 = 2 And 结帐id = n_销帐id;
        Exception
          When Others Then
            v_Err_Msg := '单据【' || 单据号_In || '】的' || Case Nvl(退费类型_In, 0)
                           When 1 Then
                            '挂号费用'
                           When 2 Then
                            '病历费'
                         End || '可能由于并发原因已经进行了退费或者单据不存在!';
            Raise Err_Item;
        End;
      End If;
      If Nvl(n_二次退费, 0) = 0 And Nvl(退费类型_In, 0) = 0 Then
        --首次全退
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
           结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
          Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In, -1 * 冲预交,
                 n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
          From 病人预交记录
          Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id;
      Else
        --二次退费,或者本次单退一部分
        --二次退费时,记录状态=3 ,首次部分退,记录状态为1
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
           结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
          Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In,
                 -1 * n_退款金额, n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
          From 病人预交记录
          Where 记录性质 = 4 And 记录状态 = Decode(Nvl(n_二次退费, 0), 0, 1, 3) And 结帐id = n_结帐id And 摘要 = '医保挂号' And 冲预交 = n_退款金额 And
                Rownum < 2;
        If Sql%RowCount = 0 Then
          Insert Into 病人预交记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
             结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
            Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In,
                   -1 * n_退款金额, n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
            From 病人预交记录
            Where 记录性质 = 4 And 记录状态 = Decode(Nvl(n_二次退费, 0), 0, 1, 3) And 结帐id = n_结帐id And 冲预交 = n_退款金额 And Rownum < 2;
        End If;
        If Sql%RowCount = 0 Then
          Insert Into 病人预交记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
             结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
            Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In,
                   -1 * n_退款金额, n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
            From 病人预交记录
            Where 记录性质 = 4 And 记录状态 = Decode(Nvl(n_二次退费, 0), 0, 1, 3) And 结帐id = n_结帐id And Rownum < 2;
          If Sql%RowCount = 0 Then
            --部分退费,并且全部使用预交款缴费时才存在此种情况
            n_预交金额 := n_退款金额;
          End If;
        End If;
      
      End If;
    End If;
    --首次退费时,记录状态便调整为了3
    Update 病人预交记录
    Set 记录状态 = 3
    Where 记录性质 = 4 And 记录状态 = Decode(Nvl(n_二次退费, 0), 0, 1, 3) And 结帐id = n_结帐id;
  
    --冲预交 1-全退 2-部分退,部分退时当全部使用预交进行缴款
    If Nvl(退费类型_In, 0) = 0 Or (Nvl(退费类型_In, 0) <> 0 And n_预交金额 <> 0) Then
      --病人挂号结算:冲预交款部份
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
         结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
        Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, d_Date,
               操作员姓名_In, 操作员编号_In, -1 * Decode(Nvl(退费类型_In, 0) + Nvl(n_二次退费, 0), 0, 冲预交, n_预交金额), n_销帐id, n_组id, 预交类别,
               卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
        From 病人预交记录
        Where 记录性质 In (1, 11) And 结帐id = n_结帐id And Nvl(冲预交, 0) <> 0 And
              Rownum = Decode(Nvl(退费类型_In, 0) + Nvl(n_二次退费, 0), 0, Rownum, 1);
    End If;
  
    --处理病人预交余额
    For c_预交 In (Select 病人id, 预交类别, -1 * Sum(Nvl(冲预交, 0)) As 冲预交
                 From 病人预交记录
                 Where 记录性质 In (1, 11) And 结帐id = n_销帐id
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
  
    If 收回票据号_In Is Not Null Then
      --光退挂号费,不回收票据
      --退卡收回票据(可能上次挂号使用票据,不能收回)
      Begin
        --从最后一次打印的内容中取
        Select ID
        Into n_打印id
        From (Select b.Id
               From 票据使用明细 A, 票据打印内容 B
               Where a.打印id = b.Id And a.性质 = 1 And a.原因 In (1, 3) And b.数据性质 = 4 And b.No = 单据号_In
               Order By a.使用时间 Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          n_打印id := Null;
      End;
    
      --先收回原票据
      If n_打印id Is Not Null Then
        Begin
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
            Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_Date, 操作员姓名_In, 票据金额
            From 票据使用明细
            Where 打印id = n_打印id And 性质 = 1 And Instr(',' || 收回票据号_In || ',', ',' || 号码 || ',') > 0;
        Exception
          When Others Then
            Delete From 票据使用明细 Where 打印id = n_打印id And 性质 = 2 And 原因 = 2;
            Insert Into 票据使用明细
              (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
              Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_Date, 操作员姓名_In, 票据金额
              From 票据使用明细
              Where 打印id = n_打印id And 性质 = 1 And Instr(',' || 收回票据号_In || ',', ',' || 号码 || ',') > 0;
        End;
      End If;
    End If;
  End If;

  --单独退病历费用,不处理汇总记录
  --相关汇总表的处理

  --病人挂号汇总
  If Nvl(退费类型_In, 0) Not In (2, 3) Then
    Open c_Registinfo(3);
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
        Select Decode(预约, Null, 0, 0, 0, 1) Into n_预约挂号 From 病人挂号记录 Where NO = 单据号_In And Rownum = 1;
      Exception
        When Others Then
          n_预约挂号 := 0;
      End;
    
      Update 病人挂号汇总
      Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - n_预约挂号, 已约数 = Nvl(已约数, 0) - n_预约挂号
      Where 日期 = Trunc(r_Registrow.发生时间) And 科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And
            (号码 = r_Registrow.号码 Or 号码 Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into 病人挂号汇总
          (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 其中已接收, 已约数)
        Values
          (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
           Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, -1, -1 * n_预约挂号, -1 * n_预约挂号);
      End If;
    
      If n_出诊记录id Is Not Null Then
        Update 临床出诊记录
        Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - n_预约挂号, 已约数 = Nvl(已约数, 0) - n_预约挂号
        Where ID = n_出诊记录id And Nvl(已约数, 0) > 0;
        Update 临床出诊序号控制 Set 挂号状态 = Null, 操作员姓名 = Null Where 记录id = n_出诊记录id And 序号 = n_序号;
      End If;
    
      Close c_Registinfo;
    End If;
  End If;

  If n_记帐 = 0 Then
    --人员缴款余额(包括个人帐户等的结算金额,不含退冲预交款)
    For r_Opermoney In c_Opermoney(n_销帐id) Loop
      Update 人员缴款余额
      Set 余额 = Nvl(余额, 0) + (-1 * r_Opermoney.冲预交)
      Where 收款员 = 操作员姓名_In And 结算方式 = r_Opermoney.结算方式 And 性质 = 1
      Returning 余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 人员缴款余额
          (收款员, 结算方式, 性质, 余额)
        Values
          (操作员姓名_In, r_Opermoney.结算方式, 1, -1 * r_Opermoney.冲预交);
        n_返回值 := r_Opermoney.冲预交;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 人员缴款余额
        Where 收款员 = 操作员姓名_In And 结算方式 = r_Opermoney.结算方式 And 性质 = 1 And Nvl(余额, 0) = 0;
      End If;
    End Loop;
  End If;
  If Nvl(退费类型_In, 0) Not In (2, 3) Then
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
    Begin
      Select 病人id, 发生时间 Into n_就诊病人id, d_就诊时间 From 病人挂号记录 Where NO = 单据号_In;
      Delete From 就诊登记记录 Where 病人id = n_就诊病人id And 就诊时间 = d_就诊时间 And 主页id Is Null;
    Exception
      When Others Then
        Null;
    End;
  
    --病人挂号记录
    Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
  
    Update 病人挂号记录 Set 记录状态 = 3 Where NO = 单据号_In And 记录性质 = 1 And 记录状态 = 1;
    If Sql%NotFound Then
      v_Err_Msg := '挂号单【' || 单据号_In || '】不存在或由于并发原因已经被退号';
      Raise Err_Item;
    End If;
    Insert Into 病人挂号记录
      (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 操作员编号, 操作员姓名,
       复诊, 号序, 社区, 预约, 摘要, 预约方式, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 险类, 医疗付款方式)
      Select n_挂号id, NO, 记录性质, 2, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, d_Date, 发生时间,
             操作员编号_In, 操作员姓名_In, 复诊, 号序, 社区, 预约, Nvl(摘要_In, 摘要) As 摘要, 预约方式, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 险类, 医疗付款方式
      From 病人挂号记录
      Where NO = 单据号_In;
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
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人挂号记录_Delete;
/

--127075:焦博,2018-07-06,调整使用到了部门参数分诊台签到排队的Oracle函数
Create Or Replace Procedure Zl_预约挂号接收_Insert
(
  No_In            门诊费用记录.No%Type,
  票据号_In        门诊费用记录.实际票号%Type,
  领用id_In        票据使用明细.领用id%Type,
  结帐id_In        门诊费用记录.结帐id%Type,
  诊室_In          门诊费用记录.发药窗口%Type,
  病人id_In        门诊费用记录.病人id%Type,
  门诊号_In        门诊费用记录.标识号%Type,
  姓名_In          门诊费用记录.姓名%Type,
  性别_In          门诊费用记录.性别%Type,
  年龄_In          门诊费用记录.年龄%Type,
  付款方式_In      门诊费用记录.付款方式%Type, --用于存放病人的医疗付款方式编号
  费别_In          门诊费用记录.费别%Type,
  结算方式_In      病人预交记录.结算方式%Type, --现金的结算名称
  现金支付_In      病人预交记录.冲预交%Type, --挂号时现金支付部份金额
  预交支付_In      病人预交记录.冲预交%Type, --挂号时使用的预交金额
  个帐支付_In      病人预交记录.冲预交%Type, --挂号时个人帐户支付金额
  发生时间_In      门诊费用记录.发生时间%Type,
  号序_In          挂号序号状态.序号%Type,
  操作员编号_In    门诊费用记录.操作员编号%Type,
  操作员姓名_In    门诊费用记录.操作员姓名%Type,
  生成队列_In      Number := 0,
  登记时间_In      门诊费用记录.登记时间%Type := Null,
  卡类别id_In      病人预交记录.卡类别id%Type := Null,
  结算卡序号_In    病人预交记录.结算卡序号%Type := Null,
  卡号_In          病人预交记录.卡号%Type := Null,
  交易流水号_In    病人预交记录.交易流水号%Type := Null,
  交易说明_In      病人预交记录.交易说明%Type := Null,
  险类_In          病人挂号记录.险类%Type := Null,
  结算模式_In      Number := 0,
  记帐费用_In      Number := 0,
  冲预交病人ids_In Varchar2 := Null,
  三方调用_In      Number := 0,
  更新交款余额_In  Number := 1,
  摘要_In          病人挂号记录.摘要%Type := Null,
  收费单_In        病人挂号记录.收费单%Type := Null
) As
  --该游标用于收费冲预交的可用预交列表
  --以ID排序，优先冲上次未冲完的。
  --更新交款余额_In:0-在zl_人员缴款余额_Update 中更新 1-在本过程中更新
  Cursor c_Deposit
  (
    v_病人id        病人信息.病人id%Type,
    v_冲预交病人ids Varchar2
  ) Is
    Select 病人id, NO, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额, Min(记录状态) As 记录状态, Nvl(Max(结帐id), 0) As 结帐id,
           Max(Decode(记录性质, 1, ID, 0) * Decode(记录状态, 2, 0, 1)) As 原预交id
    From 病人预交记录
    Where 记录性质 In (1, 11) And 病人id In (Select Column_Value From Table(f_Num2list(v_冲预交病人ids))) And Nvl(预交类别, 2) = 1
     Having Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
    Group By NO, 病人id
    Order By Decode(病人id, Nvl(v_病人id, 0), 0, 1), 结帐id, NO;

  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  Err_Special Exception;

  v_现金     结算方式.名称%Type;
  v_个人帐户 结算方式.名称%Type;
  v_队列名称 排队叫号队列.队列名称%Type;
  v_号别     门诊费用记录.计算单位%Type;
  v_号序     门诊费用记录.发药窗口%Type;
  v_排队号码 排队叫号队列.排队号码 %Type;
  v_预约方式 病人挂号记录.预约方式 %Type;

  n_预交金额      病人预交记录.金额%Type;
  n_返回值        病人预交记录.金额%Type;
  v_冲预交病人ids Varchar2(4000);

  n_挂号id         病人挂号记录.Id%Type;
  n_分诊台签到排队 Number;
  n_组id           财务缴款分组.Id%Type;
  n_Count          Number(18);
  n_排队           Number;
  n_当天排队       Number;
  n_当前金额       病人预交记录.金额%Type;
  n_预交id         病人预交记录.Id%Type;

  d_Date     Date;
  d_预约时间 门诊费用记录.发生时间%Type;
  d_发生时间 Date;
  d_排队时间 Date;
  n_时段     Number := 0;
  n_存在     Number := 0;
  v_排队序号 排队叫号队列.排队序号%Type;
  n_结算模式 病人信息.结算模式%Type;

  v_付款方式   病人挂号记录.医疗付款方式%Type;
  v_操作员姓名 病人挂号记录.接收人%Type;
  n_接收模式   Number := 0;
Begin
  n_组id          := Zl_Get组id(操作员姓名_In);
  v_冲预交病人ids := Nvl(冲预交病人ids_In, 病人id_In);
  n_接收模式      := Nvl(zl_GetSysParameter('预约接收模式', 1111), 0);

  --获取结算方式名称
  Begin
    Select 名称 Into v_现金 From 结算方式 Where 性质 = 1;
  Exception
    When Others Then
      v_现金 := '现金';
  End;
  Begin
    Select 名称 Into v_个人帐户 From 结算方式 Where 性质 = 3;
  Exception
    When Others Then
      v_个人帐户 := '个人帐户';
  End;
  If 登记时间_In Is Null Then
    Select Sysdate Into d_Date From Dual;
  Else
    d_Date := 登记时间_In;
  End If;

  --更新挂号序号状态
  Begin
    Select 号别, 号序, Trunc(发生时间), 发生时间, 预约方式
    Into v_号别, v_号序, d_预约时间, d_发生时间, v_预约方式
    From 病人挂号记录
    Where 记录性质 = 2 And 记录状态 = 1 And Rownum = 1 And NO = No_In;
  Exception
    When Others Then
      Select Max(接收人) Into v_操作员姓名 From 病人挂号记录 Where 记录性质 = 2 And 记录状态 In (1, 3) And NO = No_In;
      If v_操作员姓名 Is Null Then
        v_Err_Msg := '当前预约挂号单已被取消';
        Raise Err_Item;
      Else
        If v_操作员姓名 = 操作员姓名_In Then
          v_Err_Msg := '当前预约挂号单已被接收';
          Raise Err_Special;
        Else
          v_Err_Msg := '当前预约挂号单已被其它人接收';
          Raise Err_Item;
        End If;
      End If;
  End;

  --判断是否分时段
  Begin
    Select 1
    Into n_时段
    From Dual
    Where Exists (Select 1
           From 挂号安排时段 A, 挂号安排 B
           Where a.安排id = b.Id And b.号码 = v_号别 And Rownum < 2
           Union All
           Select 1
           From 挂号计划时段 C, 挂号安排计划 D 　
           Where c.计划id = d.Id And d.号码 = v_号别 And d.生效时间 > Sysdate And Rownum < 2);
  Exception
    When Others Then
      n_时段 := 0;
  End;
  --分时段的号别，只能当天接收
  If n_时段 = 1 And 三方调用_In = 0 And n_接收模式 = 0 Then
    If Trunc(发生时间_In) <> Trunc(Sysdate) Then
      v_Err_Msg := '分时段的预约挂号单只能当天接收！';
      Raise Err_Item;
    End If;
  End If;

  If n_时段 = 0 And 三方调用_In = 0 Then
    If n_接收模式 = 0 Then
      If Trunc(发生时间_In) = Trunc(Sysdate) Then
        d_发生时间 := 发生时间_In;
      Else
        d_发生时间 := Sysdate;
      End If;
    Else
      d_发生时间 := 发生时间_In;
    End If;
  Else
    If Not 发生时间_In Is Null Then
      d_发生时间 := 发生时间_In;
    End If;
  End If;
  If Not v_号序 Is Null Then
    If 号序_In Is Null Then
      Delete 挂号序号状态 Where 号码 = v_号别 And Trunc(日期) = Trunc(d_预约时间) And 序号 = v_号序;
    Else
      If Trunc(d_预约时间) <> Trunc(Sysdate) And n_接收模式 = 0 Then
      
        If n_时段 = 0 And 三方调用_In = 0 Then
          --提前接收或延迟接收
          Delete 挂号序号状态 Where 号码 = v_号别 And Trunc(日期) = Trunc(d_预约时间) And 序号 = v_号序;
          Begin
            Select 1 Into n_存在 From 挂号序号状态 Where 号码 = v_号别 And 日期 = Trunc(Sysdate) And 序号 = v_号序;
          Exception
            When Others Then
              n_存在 := 0;
          End;
          If n_存在 = 0 Then
            Insert Into 挂号序号状态
              (号码, 日期, 序号, 状态, 操作员姓名, 登记时间)
            Values
              (v_号别, Trunc(Sysdate), v_号序, 1, 操作员姓名_In, Sysdate);
          Else
            --号码已被使用的情况
            Begin
              v_号序 := 1;
              Insert Into 挂号序号状态
                (号码, 日期, 序号, 状态, 操作员姓名, 登记时间)
              Values
                (v_号别, Trunc(Sysdate), v_号序, 1, 操作员姓名_In, Sysdate);
            Exception
              When Others Then
                Select Min(序号 + 1)
                Into v_号序
                From 挂号序号状态 A
                Where 号码 = v_号别 And 日期 = Trunc(Sysdate) And Not Exists
                 (Select 1 From 挂号序号状态 Where 号码 = a.号码 And 日期 = a.日期 And 序号 = a.序号 + 1);
                Insert Into 挂号序号状态
                  (号码, 日期, 序号, 状态, 操作员姓名, 登记时间)
                Values
                  (v_号别, Trunc(Sysdate), v_号序, 1, 操作员姓名_In, Sysdate);
            End;
          End If;
        Else
          Update 挂号序号状态
          Set 状态 = 1, 登记时间 = Sysdate
          Where Trunc(日期) = Trunc(d_预约时间) And 序号 = v_号序 And 号码 = v_号别 And 状态 = 2;
          If Sql% NotFound Then
            Begin
              Insert Into 挂号序号状态
                (号码, 日期, 序号, 状态, 操作员姓名, 登记时间)
              Values
                (v_号别, Trunc(Sysdate), v_号序, 1, 操作员姓名_In, Sysdate);
            Exception
              When Others Then
                v_Err_Msg := '序号' || v_号序 || '已被其它人使用,请重新选择一个序号.';
                Raise Err_Item;
            End;
          End If;
        
        End If;
      
      Else
        Update 挂号序号状态
        Set 序号 = 号序_In, 状态 = 1, 登记时间 = Sysdate
        Where 号码 = v_号别 And Trunc(日期) = Trunc(d_预约时间) And 序号 = v_号序;
        If Sql%RowCount = 0 Then
          Begin
            Insert Into 挂号序号状态
              (号码, 日期, 序号, 状态, 操作员姓名, 登记时间)
            Values
              (v_号别, Trunc(d_发生时间), v_号序, 1, 操作员姓名_In, Sysdate);
          Exception
            When Others Then
              v_Err_Msg := '序号' || v_号序 || '已被其它人使用,请重新选择一个序号.';
              Raise Err_Item;
          End;
        End If;
      End If;
    End If;
  Else
    If Not 号序_In Is Null Then
      Begin
        Insert Into 挂号序号状态
          (号码, 日期, 序号, 状态, 操作员姓名, 登记时间)
        Values
          (v_号别, Trunc(Sysdate), 号序_In, 1, 操作员姓名_In, Sysdate);
      Exception
        When Others Then
          v_Err_Msg := '序号' || 号序_In || '已被其它人使用,请重新选择一个序号.';
          Raise Err_Item;
      End;
      v_号序 := 号序_In;
    Else
      v_号序 := Null;
    End If;
  End If;

  --更新门诊费用记录
  Update 门诊费用记录
  Set 记录状态 = 1, 实际票号 = Decode(Nvl(记帐费用_In, 0), 1, Null, 票据号_In), 结帐id = Decode(Nvl(记帐费用_In, 0), 1, Null, 结帐id_In),
      结帐金额 = Decode(Nvl(记帐费用_In, 0), 1, Null, 实收金额), 发药窗口 = 诊室_In, 病人id = 病人id_In, 标识号 = 门诊号_In, 姓名 = 姓名_In, 年龄 = 年龄_In,
      性别 = 性别_In, 付款方式 = 付款方式_In, 费别 = 费别_In, 发生时间 = d_发生时间, 登记时间 = d_Date, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In,
      缴款组id = n_组id, 记帐费用 = Decode(Nvl(记帐费用_In, 0), 1, 1, 0), 摘要 = Decode(收费单_In, Null, Nvl(摘要_In, 摘要), '划价:' || 收费单_In)
  Where 记录性质 = 4 And 记录状态 = 0 And NO = No_In;

  --病人挂号记录
  Update 病人挂号记录
  Set 接收人 = 操作员姓名_In, 接收时间 = d_Date, 记录性质 = 1, 病人id = 病人id_In, 门诊号 = 门诊号_In, 发生时间 = d_发生时间, 姓名 = 姓名_In, 性别 = 性别_In,
      年龄 = 年龄_In, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In, 险类 = Decode(Nvl(险类_In, 0), 0, Null, 险类_In), 号序 = v_号序, 诊室 = 诊室_In,
      摘要 = Nvl(摘要_In, 摘要), 收费单 = 收费单_In
  Where 记录状态 = 1 And NO = No_In And 记录性质 = 2
  Returning ID Into n_挂号id;
  If Sql%NotFound Then
    Begin
      Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
      Begin
        Select 名称 Into v_付款方式 From 医疗付款方式 Where 编码 = 付款方式_In And Rownum < 2;
      Exception
        When Others Then
          v_付款方式 := Null;
      End;
      Insert Into 病人挂号记录
        (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 操作员编号, 操作员姓名,
         摘要, 号序, 预约, 预约方式, 接收人, 接收时间, 预约时间, 险类, 医疗付款方式, 收费单)
        Select n_挂号id, No_In, 1, 1, 病人id_In, 门诊号_In, 姓名_In, 性别_In, 年龄_In, 计算单位, 加班标志, 诊室_In, Null, 执行部门id, 执行人, 0, Null,
               登记时间, 发生时间, 操作员编号, 操作员姓名, Nvl(摘要_In, 摘要), v_号序, 1, Substr(结论, 1, 10) As 预约方式, 操作员姓名_In,
               Nvl(登记时间_In, Sysdate), 发生时间, Decode(Nvl(险类_In, 0), 0, Null, 险类_In), v_付款方式, 收费单_In
        From 门诊费用记录
        Where 记录性质 = 4 And 记录状态 = 1 And Rownum = 1 And NO = No_In;
    Exception
      When Others Then
        v_Err_Msg := '由于并发原因,单据号为【' || No_In || '】的病人' || 姓名_In || '已经被接收';
        Raise Err_Item;
    End;
  End If;

  --0-不产生队列;1-按医生或分诊台排队;2-先分诊,后医生站
  If Nvl(生成队列_In, 0) <> 0 Then
    For v_挂号 In (Select ID, 姓名, 诊室, 执行人, 执行部门id, 发生时间, 号别, 号序 From 病人挂号记录 Where NO = No_In) Loop
      n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113, 1, Nvl(v_挂号.执行部门id, 0)));
      If Nvl(n_分诊台签到排队, 0) = 0 Then
        Begin
          Select 1,
                 Case
                   When 排队时间 < Trunc(Sysdate) Then
                    1
                   Else
                    0
                 End
          Into n_排队, n_当天排队
          From 排队叫号队列
          Where 业务类型 = 0 And 业务id = v_挂号.Id And Rownum <= 1;
        Exception
          When Others Then
            n_排队 := 0;
        End;
        If n_排队 = 0 Then
          --产生队列
          --按”执行部门”产生队列
          n_挂号id   := v_挂号.Id;
          v_队列名称 := v_挂号.执行部门id;
          v_排队号码 := Zlgetnextqueue(v_挂号.执行部门id, n_挂号id, v_挂号.号别 || '|' || v_挂号.号序);
          v_排队序号 := Zlgetsequencenum(0, n_挂号id, 0);
        
          --挂号id_In,号码_In,号序_In,缺省日期_In,扩展_In(暂无用)
          d_排队时间 := Zl_Get_Queuedate(n_挂号id, v_挂号.号别, v_挂号.号序, v_挂号.发生时间);
          --   队列名称_In , 业务类型_In, 业务id_In,科室id_In,排队号码_In,排队标记_In,患者姓名_In,病人ID_IN, 诊室_In, 医生姓名_In,
          Zl_排队叫号队列_Insert(v_队列名称, 0, n_挂号id, v_挂号.执行部门id, v_排队号码, Null, 姓名_In, 病人id_In, v_挂号.诊室, v_挂号.执行人, d_排队时间,
                           v_预约方式, Null, v_排队序号);
        Elsif Nvl(n_当天排队, 0) = 1 Then
          --更新队列号
          v_排队号码 := Zlgetnextqueue(v_挂号.执行部门id, v_挂号.Id, v_挂号.号别 || '|' || Nvl(v_挂号.号序, 0));
          v_排队序号 := Zlgetsequencenum(0, v_挂号.Id, 1);
          --新队列名称_IN, 业务类型_In, 业务id_In , 科室id_In , 患者姓名_In , 诊室_In, 医生姓名_In ,排队号码_In
          Zl_排队叫号队列_Update(v_挂号.执行部门id, 0, v_挂号.Id, v_挂号.执行部门id, v_挂号.姓名, v_挂号.诊室, v_挂号.执行人, v_排队号码, v_排队序号);
        
        Else
          --新队列名称_IN, 业务类型_In, 业务id_In , 科室id_In , 患者姓名_In , 诊室_In, 医生姓名_In ,排队号码_In
          Zl_排队叫号队列_Update(v_挂号.执行部门id, 0, v_挂号.Id, v_挂号.执行部门id, v_挂号.姓名, v_挂号.诊室, v_挂号.执行人);
        End If;
        --预约接收时，改变记录标志
        Update 病人挂号记录 Set 记录标志 = 1 Where ID = n_挂号id;
      End If;
    End Loop;
  End If;

  --汇总结算到病人预交记录
  If (Nvl(现金支付_In, 0) <> 0 Or (Nvl(现金支付_In, 0) = 0 And Nvl(个帐支付_In, 0) = 0 And Nvl(预交支付_In, 0) = 0)) And
     Nvl(记帐费用_In, 0) = 0 Then
    Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
    Insert Into 病人预交记录
      (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 结算序号,
       结算性质)
    Values
      (n_预交id, 4, 1, No_In, 病人id_In, Nvl(结算方式_In, v_现金), Nvl(现金支付_In, 0), d_Date, 操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费',
       n_组id, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, 结帐id_In, 4);
  
    If Nvl(结算卡序号_In, 0) <> 0 And Nvl(现金支付_In, 0) <> 0 Then
      Zl_病人卡结算记录_支付(结算卡序号_In, 卡号_In, 0, 现金支付_In, n_预交id, 操作员编号_In, 操作员姓名_In, d_Date);
    End If;
  
  End If;

  --对于就诊卡通过预交金挂号
  If Nvl(预交支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 Then
    n_预交金额 := 预交支付_In;
    For r_Deposit In c_Deposit(病人id_In, v_冲预交病人ids) Loop
      n_当前金额 := Case
                  When r_Deposit.金额 - n_预交金额 < 0 Then
                   r_Deposit.金额
                  Else
                   n_预交金额
                End;
      If r_Deposit.结帐id = 0 Then
        --第一次冲预交(填上结帐ID,金额为0)
        Update 病人预交记录 Set 冲预交 = 0, 结帐id = 结帐id_In, 结算性质 = 4 Where ID = r_Deposit.原预交id;
      End If;
      --冲上次剩余额
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
         结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 预交类别, 结算序号, 结算性质)
        Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 登记时间_In,
               操作员姓名_In, 操作员编号_In, n_当前金额, 结帐id_In, n_组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 预交类别, 结帐id_In, 4
        From 病人预交记录
        Where NO = r_Deposit.No And 记录状态 = r_Deposit.记录状态 And 记录性质 In (1, 11) And Rownum = 1;
    
      --更新病人预交余额
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) - n_当前金额
      Where 病人id = r_Deposit.病人id And 性质 = 1 And 类型 = Nvl(1, 2)
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (r_Deposit.病人id, Nvl(1, 2), -1 * n_当前金额, 1);
        n_返回值 := -1 * n_当前金额;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 病人余额
        Where 病人id = r_Deposit.病人id And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
      End If;
    
      --检查是否已经处理完
      If r_Deposit.金额 < n_预交金额 Then
        n_预交金额 := n_预交金额 - r_Deposit.金额;
      Else
        n_预交金额 := 0;
      End If;
    
      If n_预交金额 = 0 Then
        Exit;
      End If;
    End Loop;
  End If;

  --对于医保挂号
  If Nvl(个帐支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 Then
    Insert Into 病人预交记录
      (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
       预交类别, 结算序号, 结算性质)
    Values
      (病人预交记录_Id.Nextval, 4, 1, No_In, 病人id_In, v_个人帐户, 个帐支付_In, d_Date, 操作员编号_In, 操作员姓名_In, 结帐id_In, '医保挂号', n_组id,
       Null, Null, Null, Null, Null, Null, Null, 结帐id_In, 4);
  End If;

  --相关汇总表的处理
  --人员缴款余额
  If Nvl(现金支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 And Nvl(更新交款余额_In, 1) = 1 Then
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) + 现金支付_In
    Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = Nvl(结算方式_In, v_现金)
    Returning 余额 Into n_返回值;
  
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额
        (收款员, 结算方式, 性质, 余额)
      Values
        (操作员姓名_In, Nvl(结算方式_In, v_现金), 1, 现金支付_In);
      n_返回值 := 现金支付_In;
    
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 人员缴款余额
      Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = Nvl(结算方式_In, v_现金) And Nvl(余额, 0) = 0;
    End If;
  End If;

  If Nvl(个帐支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 And Nvl(更新交款余额_In, 1) = 1 Then
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) + 个帐支付_In
    Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = v_个人帐户
    Returning 余额 Into n_返回值;
  
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (操作员姓名_In, v_个人帐户, 1, 个帐支付_In);
      n_返回值 := 个帐支付_In;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 人员缴款余额 Where 收款员 = 操作员姓名_In And 性质 = 1 And Nvl(余额, 0) = 0;
    End If;
  End If;

  If Nvl(记帐费用_In, 0) = 1 Then
    --记帐
    If Nvl(病人id_In, 0) = 0 Then
      v_Err_Msg := '要针对病人的挂号费进行记帐，必须是建档病人才能记帐挂号。';
      Raise Err_Item;
    End If;
    For c_费用 In (Select 实收金额, 病人科室id, 开单部门id, 执行部门id, 收入项目id
                 From 门诊费用记录
                 Where 记录性质 = 4 And 记录状态 = 1 And NO = No_In And Nvl(记帐费用, 0) = 1) Loop
      --病人余额
      Update 病人余额
      Set 费用余额 = Nvl(费用余额, 0) + Nvl(c_费用.实收金额, 0)
      Where 病人id = Nvl(病人id_In, 0) And 性质 = 1 And 类型 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 性质, 类型, 费用余额, 预交余额)
        Values
          (病人id_In, 1, 1, Nvl(c_费用.实收金额, 0), 0);
      End If;
    
      --病人未结费用
      Update 病人未结费用
      Set 金额 = Nvl(金额, 0) + Nvl(c_费用.实收金额, 0)
      Where 病人id = 病人id_In And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(病人科室id, 0) = Nvl(c_费用.病人科室id, 0) And
            Nvl(开单部门id, 0) = Nvl(c_费用.开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(c_费用.执行部门id, 0) And 收入项目id + 0 = c_费用.收入项目id And
            来源途径 + 0 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人未结费用
          (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
        Values
          (病人id_In, Null, Null, c_费用.病人科室id, c_费用.开单部门id, c_费用.执行部门id, c_费用.收入项目id, 1, Nvl(c_费用.实收金额, 0));
      End If;
    End Loop;
  End If;
  If Nvl(病人id_In, 0) <> 0 Then
    n_结算模式 := 0;
    Update 病人信息
    Set 就诊时间 = d_发生时间, 就诊状态 = 1, 就诊诊室 = 诊室_In
    Where 病人id = 病人id_In
    Returning Nvl(结算模式, 0) Into n_结算模式;
    --取参数:
    If Nvl(n_结算模式, 0) <> Nvl(结算模式_In, 0) Then
      --结算模式的确定
      If n_结算模式 = 1 And Nvl(结算模式_In, 0) = 0 Then
        --病人已经是"先诊疗后结算的",本次是"先结算后诊疗的",则检查是否存在未结数据
        Select Count(1)
        Into n_Count
        From 病人未结费用
        Where 病人id = 病人id_In And (来源途径 In (1, 4) Or 来源途径 = 3 And Nvl(主页id, 0) = 0) And Nvl(金额, 0) <> 0 And Rownum < 2;
        If Nvl(n_Count, 0) <> 0 Then
          --存在未结算数据，必须先结算后才允许执行
          v_Err_Msg := '当前病人的就诊模式为先诊疗后结算且存在未结费用，不允许调整该病人的就诊模式,你可以先对未结费用结帐后再挂号或不调整病人的就诊模式!';
          Raise Err_Item;
        End If;
        --检查
        --未发生医嘱业务的（即当时就挂号的,需要保证同一次的就诊模式是一至的(程序已经检查，不用再处理)
      End If;
      Update 病人信息 Set 结算模式 = 结算模式_In Where 病人id = 病人id_In;
    End If;
  End If;

  --病人担保信息
  If 病人id_In Is Not Null Then
    Update 病人信息
    Set 担保人 = Null, 担保额 = Null, 担保性质 = Null
    Where 病人id = 病人id_In And Nvl(在院, 0) = 0 And Exists
     (Select 1
           From 病人担保记录
           Where 病人id = 病人id_In And 主页id Is Not Null And
                 登记时间 = (Select Max(登记时间) From 病人担保记录 Where 病人id = 病人id_In));
    If Sql%RowCount > 0 Then
      Update 病人担保记录
      Set 到期时间 = d_Date
      Where 病人id = 病人id_In And 主页id Is Not Null And Nvl(到期时间, d_Date) >= d_Date;
    End If;
  End If;
  --消息推送
  Begin
    Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
      Using 1, n_挂号id;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Err_Special Then
    Raise_Application_Error(-20105, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_预约挂号接收_Insert;
/

--127075:焦博,2018-07-06,调整使用到了部门参数分诊台签到排队的Oracle函数
Create Or Replace Procedure Zl_病人挂号记录_取消签到
(
  Id_In           病人挂号记录.Id%Type,
  仅改签到标志_In Integer := 0
) As
  --:仅改签到标志_In:0-需要产生相关的排队记录;1-不再处理排队记录,只作签到的标识填写 
  Err_Item Exception;
  v_Err_Msg      Varchar2(200);
  n_挂号生成队列 Number;
Begin
  --:记录标志:0表示就诊病人,1-表示签到的病人,2-表示需要回诊就诊的病人; 3-表示已回诊但还未接收的病人; 
  Update 病人挂号记录 Set 记录标志 = 0 Where ID = Id_In;
  If Sql%NotFound Then
    v_Err_Msg := '挂号项目未找到,不能继续操作!';
    Raise Err_Item;
  End If;
  If Nvl(仅改签到标志_In, 0) = 1 Then
    Return;
  End If;
  n_挂号生成队列 := Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113));
  --挂号立即签到也应该支持取消签到 
  If n_挂号生成队列 = 0 Then
    Return;
  End If;
  For v_挂号 In (Select 执行部门id, ID From 病人挂号记录 Where ID = Id_In) Loop
    Zl_排队叫号队列_Delete(v_挂号.执行部门id, v_挂号.Id);
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人挂号记录_取消签到;
/





------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0020' Where 编号=&n_System;
Commit;
