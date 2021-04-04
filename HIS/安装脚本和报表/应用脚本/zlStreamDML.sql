------------------------------------------------------------------------------------------------------------------------------------------

Create Or Replace Procedure 病人费用汇总_Dml(Any_In In Anydata) Is
  l_Lcr_App  Sys.Lcr$_Row_Record;
  i_Result   Pls_Integer;
  v_Cmd_Type Varchar2(30);

  r_End 病人费用汇总%Rowtype; --记录的最后形态：新增记录状态，修改后的新状态，删除前的状态
  r_Old 病人费用汇总%Rowtype; --记录的最后形态：记录的原状态，仅用于修改操作

  ---------------------
  --应用处理子过程
  Procedure Apply_To
  (
    r_Dat  In 病人费用汇总%Rowtype,
    n_Sign In Number := 1 --方向：1,增加；-1,减少
  ) Is
  Begin
    Update 病人费用汇总
    Set 应收金额 = Nvl(应收金额, 0) + Nvl(r_Dat.应收金额, 0) * n_Sign,
        实收金额 = Nvl(实收金额, 0) + Nvl(r_Dat.实收金额, 0) * n_Sign,
        结帐金额 = Nvl(结帐金额, 0) + Nvl(r_Dat.结帐金额, 0) * n_Sign
    Where 日期 = r_Dat.日期 And Nvl(病人病区id, 0) = Nvl(r_Dat.病人病区id, 0) And
          Nvl(病人科室id, 0) = Nvl(r_Dat.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(r_Dat.开单部门id, 0) And
          Nvl(执行部门id, 0) = Nvl(r_Dat.执行部门id, 0) And Nvl(收入项目id, 0) = Nvl(r_Dat.收入项目id, 0) And
          来源途径 = r_Dat.来源途径 And 记帐费用 = r_Dat.记帐费用;
    If Sql%Rowcount = 0 Then
      Insert Into 病人费用汇总
        (日期, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 记帐费用, 应收金额, 实收金额,
         结帐金额)
      Values
        (r_Dat.日期, r_Dat.病人病区id, r_Dat.病人科室id, r_Dat.开单部门id, r_Dat.执行部门id, r_Dat.收入项目id,
         r_Dat.来源途径, r_Dat.记帐费用, Nvl(r_Dat.应收金额, 0) * n_Sign, Nvl(r_Dat.实收金额, 0) * n_Sign,
         Nvl(r_Dat.结帐金额, 0) * n_Sign);
    End If;
    Delete 病人费用汇总
    Where 日期 = r_Dat.日期 And Nvl(病人病区id, 0) = Nvl(r_Dat.病人病区id, 0) And
          Nvl(病人科室id, 0) = Nvl(r_Dat.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(r_Dat.开单部门id, 0) And
          Nvl(执行部门id, 0) = Nvl(r_Dat.执行部门id, 0) And Nvl(收入项目id, 0) = Nvl(r_Dat.收入项目id, 0) And
          来源途径 = r_Dat.来源途径 And 记帐费用 = r_Dat.记帐费用 And Nvl(应收金额, 0) = 0 And Nvl(实收金额, 0) = 0 And
          Nvl(结帐金额, 0) = 0;
  End Apply_To;

  ---------------------
  --主过程
  ---------------------
Begin
  i_Result   := Any_In.Getobject(l_Lcr_App);
  v_Cmd_Type := l_Lcr_App.Get_Command_Type();

  if v_Cmd_Type = 'INSERT' or v_Cmd_Type = 'UPDATE' then
    i_Result := l_Lcr_App.Get_Value('new', '日期', 'Y').Getdate(r_End.日期);
    i_Result := l_Lcr_App.Get_Value('new', '病人病区ID', 'Y').Getnumber(r_End.病人病区id);
    i_Result := l_Lcr_App.Get_Value('new', '病人科室ID', 'Y').Getnumber(r_End.病人科室id);
    i_Result := l_Lcr_App.Get_Value('new', '开单部门ID', 'Y').Getnumber(r_End.开单部门id);
    i_Result := l_Lcr_App.Get_Value('new', '执行部门ID', 'Y').Getnumber(r_End.执行部门id);
    i_Result := l_Lcr_App.Get_Value('new', '收入项目ID', 'Y').Getnumber(r_End.收入项目id);
    i_Result := l_Lcr_App.Get_Value('new', '来源途径', 'Y').Getnumber(r_End.来源途径);
    i_Result := l_Lcr_App.Get_Value('new', '记帐费用', 'Y').Getnumber(r_End.记帐费用);
    i_Result := l_Lcr_App.Get_Value('new', '应收金额', 'Y').Getnumber(r_End.应收金额);
    i_Result := l_Lcr_App.Get_Value('new', '实收金额', 'Y').Getnumber(r_End.实收金额);
    i_Result := l_Lcr_App.Get_Value('new', '结帐金额', 'Y').Getnumber(r_End.结帐金额);
  end if;

  If v_Cmd_Type = 'UPDATE' Or v_Cmd_Type = 'DELETE' Then
    i_Result := l_Lcr_App.Get_Value('old', '日期', 'N').Getdate(r_Old.日期);
    i_Result := l_Lcr_App.Get_Value('old', '病人病区ID', 'N').Getnumber(r_Old.病人病区id);
    i_Result := l_Lcr_App.Get_Value('old', '病人科室ID', 'N').Getnumber(r_Old.病人科室id);
    i_Result := l_Lcr_App.Get_Value('old', '开单部门ID', 'N').Getnumber(r_Old.开单部门id);
    i_Result := l_Lcr_App.Get_Value('old', '执行部门ID', 'N').Getnumber(r_Old.执行部门id);
    i_Result := l_Lcr_App.Get_Value('old', '收入项目ID', 'N').Getnumber(r_Old.收入项目id);
    i_Result := l_Lcr_App.Get_Value('old', '来源途径', 'N').Getnumber(r_Old.来源途径);
    i_Result := l_Lcr_App.Get_Value('old', '记帐费用', 'N').Getnumber(r_Old.记帐费用);
    i_Result := l_Lcr_App.Get_Value('old', '应收金额', 'N').Getnumber(r_Old.应收金额);
    i_Result := l_Lcr_App.Get_Value('old', '实收金额', 'N').Getnumber(r_Old.实收金额);
    i_Result := l_Lcr_App.Get_Value('old', '结帐金额', 'N').Getnumber(r_Old.结帐金额);
  end if;

  If v_Cmd_Type = 'INSERT' Then
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'UPDATE' Then
    Apply_To(r_Old, -1);
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'DELETE' Then
    Apply_To(r_Old, -1);
  End If;
End 病人费用汇总_Dml;
/

Create Or Replace Procedure 病人挂号汇总_Dml(Any_In In Anydata) Is
  l_Lcr_App  Sys.Lcr$_Row_Record;
  i_Result   Pls_Integer;
  v_Cmd_Type Varchar2(30);

  r_End 病人挂号汇总%Rowtype; --记录的最后形态：新增记录状态，修改后的新状态，删除前的状态
  r_Old 病人挂号汇总%Rowtype; --记录的最后形态：记录的原状态，仅用于修改操作

  ---------------------
  --应用处理子过程
  Procedure Apply_To
  (
    r_Dat  In 病人挂号汇总%Rowtype,
    n_Sign In Number := 1 --方向：1,增加；-1,减少
  ) Is
  Begin
    Update 病人挂号汇总
    Set 已挂数 = Nvl(已挂数, 0) + Nvl(r_Dat.已挂数, 0) * n_Sign, 已约数 = Nvl(已约数, 0) + Nvl(r_Dat.已约数, 0) * n_Sign
    Where 日期 = r_Dat.日期 And 科室id = Nvl(r_Dat.科室id, 0) And 项目id = Nvl(r_Dat.项目id, 0) And
          Nvl(医生姓名, '-') = Nvl(r_Dat.医生姓名, '-') And Nvl(医生id, 0) = Nvl(r_Dat.医生id, 0);
    If Sql%Rowcount = 0 Then
      Insert Into 病人挂号汇总
        (日期, 科室id, 项目id, 医生姓名, 医生id, 已挂数, 已约数)
      Values
        (r_Dat.日期, r_Dat.科室id, r_Dat.项目id, r_Dat.医生姓名, r_Dat.医生id, Nvl(r_Dat.已挂数, 0) * n_Sign,
         Nvl(r_Dat.已约数, 0) * n_Sign);
    End If;
    Delete 病人挂号汇总
    Where 日期 = r_Dat.日期 And 科室id = Nvl(r_Dat.科室id, 0) And 项目id = Nvl(r_Dat.项目id, 0) And
          Nvl(医生姓名, '-') = Nvl(r_Dat.医生姓名, '-') And Nvl(医生id, 0) = Nvl(r_Dat.医生id, 0) And
          Nvl(已挂数, 0) = 0 And Nvl(已约数, 0) = 0;
  End Apply_To;

  ---------------------
  --主过程
  ---------------------
Begin
  i_Result   := Any_In.Getobject(l_Lcr_App);
  v_Cmd_Type := l_Lcr_App.Get_Command_Type();

  if v_Cmd_Type = 'INSERT' or v_Cmd_Type = 'UPDATE' then
   i_Result := l_Lcr_App.Get_Value('new', '日期', 'Y').Getdate(r_End.日期);
   i_Result := l_Lcr_App.Get_Value('new', '科室ID', 'Y').Getnumber(r_End.科室id);
   i_Result := l_Lcr_App.Get_Value('new', '项目ID', 'Y').Getnumber(r_End.项目id);
   i_Result := l_Lcr_App.Get_Value('new', '医生姓名', 'Y').Getvarchar2(r_End.医生姓名);
   i_Result := l_Lcr_App.Get_Value('new', '医生ID', 'Y').Getnumber(r_End.医生id);
   i_Result := l_Lcr_App.Get_Value('new', '已挂数', 'Y').Getnumber(r_End.已挂数);
   i_Result := l_Lcr_App.Get_Value('new', '已约数', 'Y').Getnumber(r_End.已约数);
  end if;

  If v_Cmd_Type = 'UPDATE' Or v_Cmd_Type = 'DELETE' Then
    i_Result := l_Lcr_App.Get_Value('old', '日期', 'N').Getdate(r_Old.日期);
    i_Result := l_Lcr_App.Get_Value('old', '科室ID', 'N').Getnumber(r_Old.科室id);
    i_Result := l_Lcr_App.Get_Value('old', '项目ID', 'N').Getnumber(r_Old.项目id);
    i_Result := l_Lcr_App.Get_Value('old', '医生姓名', 'N').Getvarchar2(r_Old.医生姓名);
    i_Result := l_Lcr_App.Get_Value('old', '医生ID', 'N').Getnumber(r_Old.医生id);
    i_Result := l_Lcr_App.Get_Value('old', '已挂数', 'N').Getnumber(r_Old.已挂数);
    i_Result := l_Lcr_App.Get_Value('old', '已约数', 'N').Getnumber(r_Old.已约数);
  end if;

  If v_Cmd_Type = 'INSERT' Then
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'UPDATE' Then
    Apply_To(r_Old, -1);
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'DELETE' Then
    Apply_To(r_Old, -1);
  End If;
End 病人挂号汇总_Dml;
/

Create Or Replace Procedure 病人未结费用_Dml(Any_In In Anydata) Is
  l_Lcr_App  Sys.Lcr$_Row_Record;
  i_Result   Pls_Integer;
  v_Cmd_Type Varchar2(30);

  r_End 病人未结费用%Rowtype; --记录的最后形态：新增记录状态，修改后的新状态，删除前的状态
  r_Old 病人未结费用%Rowtype; --记录的最后形态：记录的原状态，仅用于修改操作

  ---------------------
  --应用处理子过程
  Procedure Apply_To
  (
    r_Dat  In 病人未结费用%Rowtype,
    n_Sign In Number := 1 --方向：1,增加；-1,减少
  ) Is
  Begin
    Update 病人未结费用
    Set 金额 = 金额 + Nvl(r_Dat.金额, 0) * n_Sign
    Where 病人id = r_Dat.病人id And Nvl(主页id, 0) = Nvl(r_Dat.主页id, 0) And
          Nvl(病人病区id, 0) = Nvl(r_Dat.病人病区id, 0) And Nvl(病人科室id, 0) = Nvl(r_Dat.病人科室id, 0) And
          Nvl(开单部门id, 0) = Nvl(r_Dat.开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(r_Dat.执行部门id, 0) And
          Nvl(收入项目id, 0) = Nvl(r_Dat.收入项目id, 0) And 来源途径 = r_Dat.来源途径;
    If Sql%Rowcount = 0 Then
      Insert Into 病人未结费用
        (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
      Values
        (r_Dat.病人id, r_Dat.主页id, r_Dat.病人病区id, r_Dat.病人科室id, r_Dat.开单部门id, r_Dat.执行部门id,
         r_Dat.收入项目id, r_Dat.来源途径, Nvl(r_Dat.金额, 0) * n_Sign);
    End If;
    Delete 病人未结费用
    Where 病人id = r_Dat.病人id And Nvl(主页id, 0) = Nvl(r_Dat.主页id, 0) And
          Nvl(病人病区id, 0) = Nvl(r_Dat.病人病区id, 0) And Nvl(病人科室id, 0) = Nvl(r_Dat.病人科室id, 0) And
          Nvl(开单部门id, 0) = Nvl(r_Dat.开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(r_Dat.执行部门id, 0) And
          Nvl(收入项目id, 0) = Nvl(r_Dat.收入项目id, 0) And 来源途径 = r_Dat.来源途径 And Nvl(金额, 0) = 0;
  End Apply_To;

  ---------------------
  --主过程
  ---------------------
Begin
  i_Result   := Any_In.Getobject(l_Lcr_App);
  v_Cmd_Type := l_Lcr_App.Get_Command_Type();

  if v_Cmd_Type = 'INSERT' or v_Cmd_Type = 'UPDATE' then
   i_Result := l_Lcr_App.Get_Value('new', '病人ID', 'Y').Getnumber(r_End.病人id);
   i_Result := l_Lcr_App.Get_Value('new', '主页ID', 'Y').Getnumber(r_End.主页id);
   i_Result := l_Lcr_App.Get_Value('new', '病人病区ID', 'Y').Getnumber(r_End.病人病区id);
   i_Result := l_Lcr_App.Get_Value('new', '病人科室ID', 'Y').Getnumber(r_End.病人科室id);
   i_Result := l_Lcr_App.Get_Value('new', '开单部门ID', 'Y').Getnumber(r_End.开单部门id);
   i_Result := l_Lcr_App.Get_Value('new', '执行部门ID', 'Y').Getnumber(r_End.执行部门id);
   i_Result := l_Lcr_App.Get_Value('new', '收入项目ID', 'Y').Getnumber(r_End.收入项目id);
   i_Result := l_Lcr_App.Get_Value('new', '来源途径', 'Y').Getnumber(r_End.来源途径);
   i_Result := l_Lcr_App.Get_Value('new', '金额', 'Y').Getnumber(r_End.金额);
  end if;

  If v_Cmd_Type = 'UPDATE' Or v_Cmd_Type = 'DELETE' Then
    i_Result := l_Lcr_App.Get_Value('old', '病人ID', 'N').Getnumber(r_Old.病人id);
    i_Result := l_Lcr_App.Get_Value('old', '主页ID', 'N').Getnumber(r_Old.主页id);
    i_Result := l_Lcr_App.Get_Value('old', '病人病区ID', 'N').Getnumber(r_Old.病人病区id);
    i_Result := l_Lcr_App.Get_Value('old', '病人科室ID', 'N').Getnumber(r_Old.病人科室id);
    i_Result := l_Lcr_App.Get_Value('old', '开单部门ID', 'N').Getnumber(r_Old.开单部门id);
    i_Result := l_Lcr_App.Get_Value('old', '执行部门ID', 'N').Getnumber(r_Old.执行部门id);
    i_Result := l_Lcr_App.Get_Value('old', '收入项目ID', 'N').Getnumber(r_Old.收入项目id);
    i_Result := l_Lcr_App.Get_Value('old', '来源途径', 'N').Getnumber(r_Old.来源途径);
    i_Result := l_Lcr_App.Get_Value('old', '金额', 'N').Getnumber(r_Old.金额);
  end if;

  If v_Cmd_Type = 'INSERT' Then
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'UPDATE' Then
    Apply_To(r_Old, -1);
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'DELETE' Then
    Apply_To(r_Old, -1);
  End If;
End 病人未结费用_Dml;
/

Create Or Replace Procedure 病人余额_Dml(Any_In In Anydata) Is
  l_Lcr_App  Sys.Lcr$_Row_Record;
  i_Result   Pls_Integer;
  v_Cmd_Type Varchar2(30);

  r_End 病人余额%Rowtype; --记录的最后形态：新增记录状态，修改后的新状态，删除前的状态
  r_Old 病人余额%Rowtype; --记录的最后形态：记录的原状态，仅用于修改操作

  ---------------------
  --应用处理子过程
  Procedure Apply_To
  (
    r_Dat  In 病人余额%Rowtype,
    n_Sign In Number := 1 --方向：1,增加；-1,减少
  ) Is
  Begin
    Update 病人余额
    Set 预交余额 = Nvl(预交余额, 0) + Nvl(r_Dat.预交余额, 0) * n_Sign,
        费用余额 = Nvl(费用余额, 0) + Nvl(r_Dat.费用余额, 0) * n_Sign
    Where 病人id = r_Dat.病人id And 性质 = r_Dat.性质;
    If Sql%Rowcount = 0 Then
      Insert Into 病人余额
        (病人id, 性质, 预交余额, 费用余额)
      Values
        (r_Dat.病人id, r_Dat.性质, Nvl(r_Dat.预交余额, 0) * n_Sign, Nvl(r_Dat.费用余额, 0) * n_Sign);
    End If;
    Delete 病人余额
    Where 病人id = r_Dat.病人id And 性质 = r_Dat.性质 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
  End Apply_To;

  ---------------------
  --主过程
  ---------------------
Begin
  i_Result   := Any_In.Getobject(l_Lcr_App);
  v_Cmd_Type := l_Lcr_App.Get_Command_Type();


  IF v_Cmd_Type = 'INSERT' OR v_Cmd_Type = 'UPDATE' THEN
   i_Result := l_Lcr_App.Get_Value('new', '病人ID', 'Y').Getnumber(r_End.病人id);
   i_Result := l_Lcr_App.Get_Value('new', '性质', 'Y').Getnumber(r_End.性质);
   i_Result := l_Lcr_App.Get_Value('new', '预交余额', 'Y').Getnumber(r_End.预交余额);
   i_Result := l_Lcr_App.Get_Value('new', '费用余额', 'Y').Getnumber(r_End.费用余额);	
  end if;
  
  If v_Cmd_Type = 'UPDATE' Or v_Cmd_Type = 'DELETE' Then
    i_Result := l_Lcr_App.Get_Value('old', '病人ID', 'N').Getnumber(r_Old.病人id);
    i_Result := l_Lcr_App.Get_Value('old', '性质', 'N').Getnumber(r_Old.性质);
    i_Result := l_Lcr_App.Get_Value('old', '预交余额', 'N').Getnumber(r_Old.预交余额);
    i_Result := l_Lcr_App.Get_Value('old', '费用余额', 'N').Getnumber(r_Old.费用余额);
  END IF;

  If v_Cmd_Type = 'INSERT' Then
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'UPDATE' Then
    Apply_To(r_Old, -1);
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'DELETE' Then     
    Apply_To(r_Old, -1);
  End If;
End 病人余额_Dml;
/

Create Or Replace Procedure 人员缴款余额_Dml(Any_In In Anydata) Is
  l_Lcr_App  Sys.Lcr$_Row_Record;
  i_Result   Pls_Integer;
  v_Cmd_Type Varchar2(30);

  r_End 人员缴款余额%Rowtype; --记录的最后形态：新增记录状态，修改后的新状态，删除前的状态
  r_Old 人员缴款余额%Rowtype; --记录的最后形态：记录的原状态，仅用于修改操作

  ---------------------
  --应用处理子过程
  Procedure Apply_To
  (
    r_Dat  In 人员缴款余额%Rowtype,
    n_Sign In Number := 1 --方向：1,增加；-1,减少
  ) Is
  Begin
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) + Nvl(r_Dat.余额, 0) * n_Sign
    Where 收款员 = r_Dat.收款员 And 结算方式 = r_Dat.结算方式 And 性质 = r_Dat.性质;
    If Sql%Rowcount = 0 Then
      Insert Into 人员缴款余额
        (收款员, 结算方式, 性质, 余额)
      Values
        (r_Dat.收款员, r_Dat.结算方式, r_Dat.性质, Nvl(r_Dat.余额, 0) * n_Sign);
    End If;
    Delete 人员缴款余额
    Where 收款员 = r_Dat.收款员 And 结算方式 = r_Dat.结算方式 And 性质 = r_Dat.性质 And Nvl(余额, 0) = 0;
  End Apply_To;

  ---------------------
  --主过程
  ---------------------
Begin
  i_Result   := Any_In.Getobject(l_Lcr_App);
  v_Cmd_Type := l_Lcr_App.Get_Command_Type();

  IF v_Cmd_Type = 'INSERT' OR v_Cmd_Type = 'UPDATE' THEN
   i_Result := l_Lcr_App.Get_Value('new', '收款员', 'Y').Getvarchar2(r_End.收款员);
   i_Result := l_Lcr_App.Get_Value('new', '结算方式', 'Y').Getvarchar2(r_End.结算方式);
   i_Result := l_Lcr_App.Get_Value('new', '性质', 'Y').Getnumber(r_End.性质);
   i_Result := l_Lcr_App.Get_Value('new', '余额', 'Y').Getnumber(r_End.余额);
  end if;

  If v_Cmd_Type = 'UPDATE' Or v_Cmd_Type = 'DELETE' Then
    i_Result := l_Lcr_App.Get_Value('old', '收款员', 'N').Getvarchar2(r_Old.收款员);
    i_Result := l_Lcr_App.Get_Value('old', '结算方式', 'N').Getvarchar2(r_Old.结算方式);
    i_Result := l_Lcr_App.Get_Value('old', '性质', 'N').Getnumber(r_Old.性质);
    i_Result := l_Lcr_App.Get_Value('old', '余额', 'N').Getnumber(r_Old.余额);
  END IF;

  If v_Cmd_Type = 'INSERT' Then
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'UPDATE' Then
    Apply_To(r_Old, -1);
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'DELETE' Then
    Apply_To(r_Old, -1);
  End If;
End 人员缴款余额_Dml;
/

--2007/07/05 刘兴宏:处理药品库存\药品收发汇总\应付余额的DML过程

--相关的DML过程
Create Or Replace Procedure 药品库存_Dml(Any_In In Anydata) Is
  l_Lcr_App  Sys.Lcr$_Row_Record;
  i_Result   Pls_Integer;
  v_Cmd_Type Varchar2(30);

  r_End 药品库存%RowType; --记录的最后形态：新增记录状态，修改后的新状态，删除前的状态
  r_Old 药品库存%RowType; --记录的最后形态：记录的原状态，仅用于修改操作

  ---------------------
  --应用处理子过程
  Procedure Apply_To
  (
    r_Dat  In 药品库存%RowType,
    n_Sign In Number := 1 --方向：1,增加；-1,减少
  ) Is
  Begin
    Update 药品库存
    Set 可用数量 = Nvl(可用数量, 0) + Nvl(r_Dat.可用数量, 0) * n_Sign, 实际数量 = Nvl(实际数量, 0) + Nvl(r_Dat.实际数量, 0) * n_Sign,
        实际金额 = Nvl(实际金额, 0) + Nvl(r_Dat.实际金额, 0) * n_Sign, 实际差价 = Nvl(实际差价, 0) + Nvl(r_Dat.实际差价, 0) * n_Sign,
        上次采购价 = Nvl(r_Dat.上次采购价, 上次采购价), 上次产地 = Nvl(r_Dat.上次产地, 上次产地), 上次批号 = Nvl(r_Dat.上次批号, 上次批号),
        上次生产日期 = Nvl(r_Dat.上次生产日期, 上次生产日期), 零售价 = Nvl(r_Dat.零售价, 零售价), 上次扣率 = Nvl(r_Dat.上次扣率, 上次扣率)
    Where 库房id = r_Dat.库房id And 药品id = r_Dat.药品id And Nvl(批次, 0) = Nvl(r_Dat.批次, 0) And 性质 = r_Dat.性质;
    If Sql%RowCount = 0 Then
      Insert Into 药品库存
        (库房id, 药品id, 批次, 效期, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 灭菌效期, 批准文号, 零售价, 上次扣率)
      Values
        (r_Dat.库房id, r_Dat.药品id, r_Dat.批次, r_Dat.效期, r_Dat.性质, Nvl(r_Dat.可用数量, 0) * n_Sign, Nvl(r_Dat.实际数量, 0) * n_Sign,
         Nvl(r_Dat.实际金额, 0) * n_Sign, Nvl(r_Dat.实际差价, 0) * n_Sign, r_Dat.上次供应商id, r_Dat.上次采购价, r_Dat.上次批号, r_Dat.上次生产日期,
         r_Dat.上次产地, r_Dat.灭菌效期, r_Dat.批准文号, r_Dat.零售价, r_Dat.上次扣率);
    End If;
    Delete 药品库存
    Where 库房id = r_Dat.库房id And 药品id = r_Dat.药品id And Nvl(批次, 0) = Nvl(r_Dat.批次, 0) And 性质 = r_Dat.性质 And
          Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And Nvl(实际差价, 0) = 0;
  End Apply_To;

  ---------------------
  --主过程
  ---------------------
Begin
  i_Result   := Any_In.Getobject(l_Lcr_App);
  v_Cmd_Type := l_Lcr_App.Get_Command_Type();

  If v_Cmd_Type = 'INSERT' Or v_Cmd_Type = 'UPDATE' Then
    i_Result := l_Lcr_App.Get_Value('new', '库房ID', 'Y').Getnumber(r_End.库房id);
    i_Result := l_Lcr_App.Get_Value('new', '药品ID', 'Y').Getnumber(r_End.药品id);
    i_Result := l_Lcr_App.Get_Value('new', '批次', 'Y').Getnumber(r_End.批次);
    i_Result := l_Lcr_App.Get_Value('new', '效期', 'Y').Getdate(r_End.效期);
    i_Result := l_Lcr_App.Get_Value('new', '性质', 'Y').Getnumber(r_End.性质);
    i_Result := l_Lcr_App.Get_Value('new', '可用数量', 'Y').Getnumber(r_End.可用数量);
    i_Result := l_Lcr_App.Get_Value('new', '实际数量', 'Y').Getnumber(r_End.实际数量);
    i_Result := l_Lcr_App.Get_Value('new', '实际金额', 'Y').Getnumber(r_End.实际金额);
    i_Result := l_Lcr_App.Get_Value('new', '实际差价', 'Y').Getnumber(r_End.实际差价);
    i_Result := l_Lcr_App.Get_Value('new', '上次供应商ID', 'Y').Getnumber(r_End.上次供应商id);
    i_Result := l_Lcr_App.Get_Value('new', '上次采购价', 'Y').Getnumber(r_End.上次采购价);
    i_Result := l_Lcr_App.Get_Value('new', '上次批号', 'Y').Getvarchar2(r_End.上次批号);
    i_Result := l_Lcr_App.Get_Value('new', '上次生产日期', 'Y').Getdate(r_End.上次生产日期);
    i_Result := l_Lcr_App.Get_Value('new', '上次产地', 'Y').Getvarchar2(r_End.上次产地);
    i_Result := l_Lcr_App.Get_Value('new', '灭菌效期', 'Y').Getdate(r_End.灭菌效期);
    i_Result := l_Lcr_App.Get_Value('new', '批准文号', 'Y').Getvarchar2(r_End.批准文号);
    i_Result := l_Lcr_App.Get_Value('new', '零售价', 'Y').Getnumber(r_End.零售价);
    i_Result := l_Lcr_App.Get_Value('new', '上次扣率', 'Y').Getnumber(r_End.上次扣率);
  End If;

  If v_Cmd_Type = 'UPDATE' Or v_Cmd_Type = 'DELETE' Then
    i_Result := l_Lcr_App.Get_Value('old', '库房ID', 'N').Getnumber(r_Old.库房id);
    i_Result := l_Lcr_App.Get_Value('old', '药品ID', 'N').Getnumber(r_Old.药品id);
    i_Result := l_Lcr_App.Get_Value('old', '批次', 'N').Getnumber(r_Old.批次);
    i_Result := l_Lcr_App.Get_Value('old', '效期', 'N').Getdate(r_Old.效期);
    i_Result := l_Lcr_App.Get_Value('old', '性质', 'N').Getnumber(r_Old.性质);
    i_Result := l_Lcr_App.Get_Value('old', '可用数量', 'N').Getnumber(r_Old.可用数量);
    i_Result := l_Lcr_App.Get_Value('old', '实际数量', 'N').Getnumber(r_Old.实际数量);
    i_Result := l_Lcr_App.Get_Value('old', '实际金额', 'N').Getnumber(r_Old.实际金额);
    i_Result := l_Lcr_App.Get_Value('old', '实际差价', 'N').Getnumber(r_Old.实际差价);
    i_Result := l_Lcr_App.Get_Value('old', '上次供应商ID', 'N').Getnumber(r_Old.上次供应商id);
    i_Result := l_Lcr_App.Get_Value('old', '上次采购价', 'N').Getnumber(r_Old.上次采购价);
    i_Result := l_Lcr_App.Get_Value('old', '上次批号', 'N').Getvarchar2(r_Old.上次批号);
    i_Result := l_Lcr_App.Get_Value('old', '上次生产日期', 'N').Getdate(r_Old.上次生产日期);
    i_Result := l_Lcr_App.Get_Value('old', '上次产地', 'N').Getvarchar2(r_Old.上次产地);
    i_Result := l_Lcr_App.Get_Value('old', '灭菌效期', 'N').Getdate(r_Old.灭菌效期);
    i_Result := l_Lcr_App.Get_Value('old', '批准文号', 'N').Getvarchar2(r_Old.批准文号);
    i_Result := l_Lcr_App.Get_Value('old', '零售价', 'N').Getnumber(r_Old.零售价);
    i_Result := l_Lcr_App.Get_Value('old', '上次扣率', 'N').Getnumber(r_Old.上次扣率);
  End If;
  If v_Cmd_Type = 'INSERT' Then
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'UPDATE' Then
    Apply_To(r_Old, -1);
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'DELETE' Then
    Apply_To(r_Old, -1);
  End If;
End 药品库存_Dml;
/


Create Or Replace Procedure 药品收发汇总_Dml(Any_In In Anydata) Is
  l_Lcr_App  Sys.Lcr$_Row_Record;
  i_Result   Pls_Integer;
  v_Cmd_Type Varchar2(30);

  r_End 药品收发汇总%Rowtype; --记录的最后形态：新增记录状态，修改后的新状态，删除前的状态
  r_Old 药品收发汇总%Rowtype; --记录的最后形态：记录的原状态，仅用于修改操作

  ---------------------
  --应用处理子过程
  Procedure Apply_To
  (
    r_Dat  In 药品收发汇总%Rowtype,
    n_Sign In Number := 1 --方向：1,增加；-1,减少
  ) Is
  Begin
    Update 药品收发汇总
    Set 数量 = Nvl(数量, 0) + Nvl(r_Dat.数量, 0) * n_Sign, 金额 = Nvl(金额, 0) + Nvl(r_Dat.金额, 0) * n_Sign,
        差价 = Nvl(差价, 0) + Nvl(r_Dat.差价, 0) * n_Sign
    Where 日期 = r_Dat.日期 And Nvl(库房id, 0) = Nvl(r_Dat.库房id, 0) And Nvl(药品id, 0) = Nvl(r_Dat.药品id, 0) And
          Nvl(类别id, 0) = Nvl(r_Dat.类别id, 0) And 单据 = r_Dat.单据;
    If Sql%Rowcount = 0 Then
      Insert Into 药品收发汇总
        (日期, 库房id, 药品id, 类别id, 单据, 数量, 金额, 差价)
      Values
        (r_Dat.日期, r_Dat.库房id, r_Dat.药品id, r_Dat.类别id, r_Dat.单据, Nvl(r_Dat.数量, 0) * n_Sign,
         Nvl(r_Dat.金额, 0) * n_Sign, Nvl(r_Dat.差价, 0) * n_Sign);
    End If;
    Delete 药品收发汇总
    Where 日期 = r_Dat.日期 And Nvl(库房id, 0) = Nvl(r_Dat.库房id, 0) And Nvl(药品id, 0) = Nvl(r_Dat.药品id, 0) And
          Nvl(类别id, 0) = Nvl(r_Dat.类别id, 0) And 单据 = r_Dat.单据 And Nvl(数量, 0) = 0 And Nvl(金额, 0) = 0 And
          Nvl(差价, 0) = 0;
  End Apply_To;

  ---------------------
  --主过程
  ---------------------
Begin
  i_Result   := Any_In.Getobject(l_Lcr_App);
  v_Cmd_Type := l_Lcr_App.Get_Command_Type();

  if v_Cmd_Type = 'INSERT' or v_Cmd_Type = 'UPDATE' then
   i_Result := l_Lcr_App.Get_Value('new', '日期', 'Y').Getdate(r_End.日期);
   i_Result := l_Lcr_App.Get_Value('new', '库房ID', 'Y').Getnumber(r_End.库房id);
   i_Result := l_Lcr_App.Get_Value('new', '药品ID', 'Y').Getnumber(r_End.药品id);
   i_Result := l_Lcr_App.Get_Value('new', '类别ID', 'Y').Getnumber(r_End.类别id);
   i_Result := l_Lcr_App.Get_Value('new', '单据', 'Y').Getnumber(r_End.单据);
   i_Result := l_Lcr_App.Get_Value('new', '数量', 'Y').Getnumber(r_End.数量);
   i_Result := l_Lcr_App.Get_Value('new', '金额', 'Y').Getnumber(r_End.金额);
   i_Result := l_Lcr_App.Get_Value('new', '差价', 'Y').Getnumber(r_End.差价);
  end if;

  If v_Cmd_Type = 'UPDATE' Or v_Cmd_Type = 'DELETE' Then
    i_Result := l_Lcr_App.Get_Value('old', '日期', 'N').Getdate(r_Old.日期);
    i_Result := l_Lcr_App.Get_Value('old', '库房ID', 'N').Getnumber(r_Old.库房id);
    i_Result := l_Lcr_App.Get_Value('old', '药品ID', 'N').Getnumber(r_Old.药品id);
    i_Result := l_Lcr_App.Get_Value('old', '类别ID', 'N').Getnumber(r_Old.类别id);
    i_Result := l_Lcr_App.Get_Value('old', '单据', 'N').Getnumber(r_Old.单据);
    i_Result := l_Lcr_App.Get_Value('old', '数量', 'N').Getnumber(r_Old.数量);
    i_Result := l_Lcr_App.Get_Value('old', '金额', 'N').Getnumber(r_Old.金额);
    i_Result := l_Lcr_App.Get_Value('old', '差价', 'N').Getnumber(r_Old.差价);
  end if;

  If v_Cmd_Type = 'INSERT' Then
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'UPDATE' Then
    Apply_To(r_Old, -1);
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'DELETE' Then
    Apply_To(r_Old, -1);
  End If;
End 药品收发汇总_Dml;
/


Create Or Replace Procedure 应付余额_Dml(Any_In In Anydata) Is
  l_Lcr_App  Sys.Lcr$_Row_Record;
  i_Result   Pls_Integer;
  v_Cmd_Type Varchar2(30);

  r_End 应付余额%Rowtype; --记录的最后形态：新增记录状态，修改后的新状态，删除前的状态
  r_Old 应付余额%Rowtype; --记录的最后形态：记录的原状态，仅用于修改操作

  ---------------------
  --应用处理子过程
  Procedure Apply_To
  (
    r_Dat  In 应付余额%Rowtype,
    n_Sign In Number := 1 --方向：1,增加；-1,减少
  ) Is
  Begin
    Update 应付余额
    Set 金额 = Nvl(金额, 0) + Nvl(r_Dat.金额, 0) * n_Sign
    Where 单位id = r_Dat.单位id And 性质 = r_Dat.性质;
    If Sql%Rowcount = 0 Then
      Insert Into 应付余额 (单位id, 性质, 金额) Values (r_Dat.单位id, r_Dat.性质, Nvl(r_Dat.金额, 0) * n_Sign);
    End If;
    Delete 应付余额 Where 单位id = r_Dat.单位id And 性质 = r_Dat.性质 And Nvl(金额, 0) = 0;
  End Apply_To;

  ---------------------
  --主过程
  ---------------------
Begin
  i_Result   := Any_In.Getobject(l_Lcr_App);
  v_Cmd_Type := l_Lcr_App.Get_Command_Type();

  IF v_Cmd_Type = 'INSERT' OR v_Cmd_Type = 'UPDATE' THEN
   i_Result := l_Lcr_App.Get_Value('new', '单位ID', 'Y').Getnumber(r_End.单位id);
   i_Result := l_Lcr_App.Get_Value('new', '性质', 'Y').Getnumber(r_End.性质);
   i_Result := l_Lcr_App.Get_Value('new', '金额', 'Y').Getnumber(r_End.金额);
  end if;

  If v_Cmd_Type = 'UPDATE' Or v_Cmd_Type = 'DELETE' Then 
    i_Result := l_Lcr_App.Get_Value('old', '单位ID', 'Y').Getnumber(r_Old.单位id);
    i_Result := l_Lcr_App.Get_Value('old', '性质', 'N').Getnumber(r_Old.性质);
    i_Result := l_Lcr_App.Get_Value('old', '金额', 'N').Getnumber(r_Old.金额);
  END IF;  

  If v_Cmd_Type = 'INSERT' Then
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'UPDATE' Then
    Apply_To(r_Old, -1);
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'DELETE' Then
    Apply_To(r_Old, -1);
  End If;
End 应付余额_Dml;
/

------------------------------------------------------------------------------------------------------------------------------------------

--李业庆：药品留存DML
Create Or Replace Procedure 药品留存_Dml(Any_In In Anydata) Is
  l_Lcr_App  Sys.Lcr$_Row_Record;
  i_Result   Pls_Integer;
  v_Cmd_Type Varchar2(30);

  r_End 药品留存%Rowtype; --记录的最后形态：新增记录状态，修改后的新状态，删除前的状态
  r_Old 药品留存%Rowtype; --记录的最后形态：记录的原状态，仅用于修改操作

  ---------------------
  --应用处理子过程
  Procedure Apply_To
  (
    r_Dat  In 药品留存%Rowtype,
    n_Sign In Number := 1 --方向：1,增加；-1,减少
  ) Is
  Begin
    Update 药品留存
    Set 可用数量 = Nvl(可用数量, 0) + Nvl(r_Dat.可用数量, 0) * n_Sign,
        实际数量 = Nvl(实际数量, 0) + Nvl(r_Dat.实际数量, 0) * n_Sign,
        实际金额 = Nvl(实际金额, 0) + Nvl(r_Dat.实际金额, 0) * n_Sign
    Where 期间 = r_Dat.期间 And 科室id = r_Dat.科室id And 库房id = r_Dat.库房id And 药品id = r_Dat.药品id;
    If Sql%Rowcount = 0 Then
      Insert Into 药品留存
        (期间, 科室id, 库房id, 药品id, 可用数量, 实际数量, 实际金额)
      Values
        (r_Dat.期间, r_Dat.科室id, r_Dat.库房id, r_Dat.药品id, Nvl(r_Dat.可用数量, 0) * n_Sign,
         Nvl(r_Dat.实际数量, 0) * n_Sign, Nvl(r_Dat.实际金额, 0) * n_Sign);
    End If;
    Delete 药品留存
    Where 期间 = r_Dat.期间 And 科室id = r_Dat.科室id And 库房id = r_Dat.库房id And 药品id = r_Dat.药品id And
          Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0;
  End Apply_To;

  ---------------------
  --主过程
  ---------------------
Begin
  i_Result   := Any_In.Getobject(l_Lcr_App);
  v_Cmd_Type := l_Lcr_App.Get_Command_Type();

  IF v_Cmd_Type = 'INSERT' OR v_Cmd_Type = 'UPDATE' THEN
   i_Result := l_Lcr_App.Get_Value('new', '期间', 'Y').Getvarchar2(r_End.期间);
   i_Result := l_Lcr_App.Get_Value('new', '科室ID', 'Y').Getnumber(r_End.科室id);
   i_Result := l_Lcr_App.Get_Value('new', '库房ID', 'Y').Getnumber(r_End.库房id);
   i_Result := l_Lcr_App.Get_Value('new', '药品ID', 'Y').Getnumber(r_End.药品id);
   i_Result := l_Lcr_App.Get_Value('new', '可用数量', 'Y').Getnumber(r_End.可用数量);
   i_Result := l_Lcr_App.Get_Value('new', '实际数量', 'Y').Getnumber(r_End.实际数量);
   i_Result := l_Lcr_App.Get_Value('new', '实际金额', 'Y').Getnumber(r_End.实际金额);
  end if;

  If v_Cmd_Type = 'UPDATE' Or v_Cmd_Type = 'DELETE' Then
    i_Result := l_Lcr_App.Get_Value('old', '期间', 'Y').Getvarchar2(r_Old.期间);
    i_Result := l_Lcr_App.Get_Value('old', '科室ID', 'Y').Getnumber(r_Old.科室id);
    i_Result := l_Lcr_App.Get_Value('old', '库房ID', 'N').Getnumber(r_Old.库房id);
    i_Result := l_Lcr_App.Get_Value('old', '药品ID', 'N').Getnumber(r_Old.药品id);
    i_Result := l_Lcr_App.Get_Value('old', '可用数量', 'N').Getnumber(r_Old.可用数量);
    i_Result := l_Lcr_App.Get_Value('old', '实际数量', 'N').Getnumber(r_Old.实际数量);
    i_Result := l_Lcr_App.Get_Value('old', '实际金额', 'N').Getnumber(r_Old.实际金额);
  END IF;

  If v_Cmd_Type = 'INSERT' Then
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'UPDATE' Then
    Apply_To(r_Old, -1);
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'DELETE' Then
    Apply_To(r_Old, -1);
  End If;
End 药品留存_Dml;
/

