Create Or Replace Procedure Zl_座位状况记录_Update
(
  科室id_In     In 座位状况记录.科室id%Type,
  编号_In       In 座位状况记录.编号%Type,
  收费细目id_In In 座位状况记录.收费细目id%Type,
  状态_In       In 座位状况记录.状态%Type,
  备注_In       In 座位状况记录.备注%Type
) Is
Begin
  Update 座位状况记录
  Set 收费细目id = 收费细目id_In, 状态 = 状态_In, 备注 = 备注_In
  Where 科室id = 科室id_In And 编号 = 编号_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_座位状况记录_Update;
/

CREATE OR REPLACE Procedure Zl_座位状况记录_Insert
(
  科室id_In     In 座位状况记录.科室id%Type,
  编号_In       In 座位状况记录.编号%Type,
  类别_In       In 座位状况记录.类别%Type,
  状态_In       In 座位状况记录.状态%Type,
  收费细目id_In In 座位状况记录.收费细目id%Type,
  备注_In       In 座位状况记录.备注%Type
) Is
Begin
  Insert Into 座位状况记录
    (科室id, 编号, 类别, 状态, 收费细目id, 备注)
  Values
    (科室id_In, 编号_In, 类别_In, 状态_In, 收费细目id_In, 备注_In);
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_座位状况记录_Insert;
/

Create Or Replace Procedure Zl_座位状况记录_Delete
(
  科室id_In In 座位状况记录.科室id%Type,
  编号_In   In 座位状况记录.编号%Type
) Is
Begin
  Delete 座位状况记录 Where nvl(病人id,0) = 0 And 状态 <> 1 And 科室id = 科室id_In And 编号 = 编号_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_座位状况记录_Delete;
/

Create Or Replace Procedure Zl_座位状况记录_Setseating
(
  科室id_In In 座位状况记录.科室id%Type,
  类别_In   In 座位状况记录.类别%Type,
  编号_In   In 座位状况记录.编号%Type,
  病人id_In In 座位状况记录.病人id%Type,
  NO_In   In 座位状况记录.NO%Type
) Is
Begin
  If 病人id_In <> 0 Then
    -- 占用
    Update 座位状况记录
    Set 病人id = 病人id_In, 状态 = 1, NO = NO_In
    Where 科室id = 科室id_In And 类别 = 类别_In And 编号 = 编号_In And Nvl(状态, 0) = 0 And Nvl(病人id, 0) = 0;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_座位状况记录_Setseating;
/

Create Or Replace Procedure Zl_座位状况记录_Clear
(
  科室id_In In 座位状况记录.科室id%Type,
  编号_In   In 座位状况记录.编号%Type
) Is
Begin
  Update 座位状况记录
  Set 病人id = Null, 状态 = 0, NO = Null
  Where Nvl(病人id, 0) <> 0 And 状态 = 1 And 科室id = 科室id_In And 编号 = 编号_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_座位状况记录_Clear;
/

Create Or Replace Procedure Zl_病人医嘱执行_Transfusion
(
  医嘱id_In   In 病人医嘱执行.医嘱id%Type,
  发送号_In   In 病人医嘱执行.发送号%Type,
  执行时间_In In 病人医嘱执行.执行时间%Type,
  流水号_In   In 病人医嘱执行.流水号%Type,
  配药人_In   In 病人医嘱执行.配药人%Type,
  组数_In     In 病人医嘱执行.组数%Type,
  组次_In     In 病人医嘱执行.组次%Type,
  滴速_In     In 病人医嘱执行.滴速%Type,
  滴系数_In   In 病人医嘱执行.滴系数%Type,
  液体量_In   In 病人医嘱执行.液体量%Type,
  说明_In     In 病人医嘱执行.说明%Type,
  接单人_In   In 病人医嘱执行.接单人%Type,
  耗时_In     In 病人医嘱执行.耗时%Type,
  提醒_In     In 病人医嘱执行.提醒%Type
) Is

Begin

  Update 病人医嘱执行
  Set 流水号 = 流水号_In, 配药人 = 配药人_In, 组数 = 组数_In, 组次 = 组次_In, 滴速 = 滴速_In, 滴系数 = 滴系数_In,
      液体量 = 液体量_In, 说明 = 说明_In, 接单人 = 接单人_In, 耗时 = 耗时_In, 提醒 = 提醒_In
  Where 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In And 执行时间 = 执行时间_In;

Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病人医嘱执行_Transfusion;
/

Create Or Replace Procedure Zl_病人医嘱执行_Modify
(
  流水号_In In 病人医嘱执行.流水号%Type,
  医嘱id_In In 病人医嘱执行.医嘱id%Type,
  发送号_In In 病人医嘱执行.发送号%Type,
  滴速_In   In 病人医嘱执行.滴速%Type,
  液体量_In In 病人医嘱执行.液体量%Type,
  滴系数_In In 病人医嘱执行.滴系数%Type,
  耗时_In   In 病人医嘱执行.耗时%Type,
  说明_In   In 病人医嘱执行.说明%Type
) Is
  --除了要执行的主记录,还包含了附加手术,检查部位的记录 
  --手术麻醉,中药煎法,采集方法单独控制,检验组合只填写在第一个项目上，但执行状态相同 
  Cursor c_Advice Is
    Select A.医嘱id, B.相关id, B.诊疗类别
    From 病人医嘱发送 A, 病人医嘱记录 B
    Where (B.ID = 医嘱id_In Or (B.相关id = 医嘱id_In And B.诊疗类别 In ('F', 'D'))) And A.医嘱id = B.ID And
          A.发送号 + 0 = 发送号_In;

  v_Temp     Varchar2(255);
  v_人员姓名 病人费用记录.操作员姓名%Type;
  v_Date  Date;
Begin
  v_Temp     := Zl_Identity;
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);

  Select Sysdate Into v_Date From Dual;

  For r_Advice In c_Advice Loop
    Update 病人医嘱执行
    Set 滴速 = 滴速_In, 液体量 = 液体量_In, 滴系数 = 滴系数_In, 耗时 = 耗时_In, 说明 = 说明_In, 登记时间 = v_Date,
        登记人 = v_人员姓名
    Where 医嘱id = r_Advice.医嘱id And 发送号 + 0 = 发送号_In And 流水号 = 流水号_In;
  End Loop;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病人医嘱执行_Modify;
/

Create Or Replace Procedure Zl_排队记录_Addqueue
(
	病人id_In In 排队记录.病人id%Type,
	科室id_In In 排队记录.科室id%Type,
	顺序号_In In 排队记录.顺序号%Type
) Is

Begin
	-- 一个病人在一个科室只能有一条排队记录 ,所以,先删除该科室原来的排队记录,再写入新记录.
	Delete 排队记录 Where 病人id = 病人id_In And 科室id = 科室id_In;
	Insert Into 排队记录
		(病人id, 科室id, 顺序号, 加权号, 状态, 备注, 日期)
	Values
		(病人id_In, 科室id_In, 顺序号_In, 0, 1, '', Sysdate);
Exception
	When Others Then
		Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_排队记录_Addqueue;
/

CREATE OR REPLACE Procedure Zl_排队记录_Update
(
	病人id_In In 排队记录.病人id%Type,
	科室id_In In 排队记录.科室id%Type,
	顺序号_In In 排队记录.顺序号%Type,
	加权号_In In 排队记录.加权号%Type,
	状态_In   In 排队记录.状态%Type
) Is

Begin
	Update 排队记录
	Set 加权号 = 加权号_In, 状态 = 状态_In
	Where 病人id = 病人id_In And 科室id = 科室id_In And 顺序号 = 顺序号_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);

End Zl_排队记录_Update;
/

Create Or Replace Procedure Zl_暂存药品记录_Insert
(
	No_In       In 暂存药品记录.No%Type,
	序号_In     In 暂存药品记录.序号%Type,
	病人id_In   In 暂存药品记录.病人id %Type,
	医嘱id_In   In 暂存药品记录.医嘱id%Type,
	发送号_In   In 暂存药品记录.发送号%Type,
	药品id_In   In 暂存药品记录.药品id %Type,
	药品名称_In In 暂存药品记录.药品名称%Type,
	规格_In     In 暂存药品记录.规格%Type,
	执行分类_In In 暂存药品记录.执行分类%Type,
	使用状态_In In 暂存药品记录.使用状态%Type,
	摘要_In     In 暂存药品记录.摘要%Type,
	入出系数_In In 暂存药品记录.入出系数%Type,
	单位_In     In 暂存药品记录.单位%Type,
	容量_In     In 暂存药品记录.容量%Type,
	数量_In     In 暂存药品记录.数量%Type,
	单价_In     In 暂存药品记录.单价%Type,
	金额_In     In 暂存药品记录.金额%Type,
	操作员_In   In 暂存药品记录.操作员%Type,
	科室id_In   In 暂存药品记录.科室id%Type,
	登记时间_In In 暂存药品记录.登记时间%Type
) Is
Begin
	Insert Into 暂存药品记录
		(No, 序号, 病人id, 医嘱id, 发送号, 药品id, 药品名称, 规格, 执行分类, 使用状态, 摘要, 入出系数, 单位, 容量, 数量,
		 单价, 金额, 操作员, 登记时间, 科室id)
	Values
		(No_In, 序号_In, 病人id_In, 医嘱id_In, 发送号_In, 药品id_In, 药品名称_In, 规格_In, 执行分类_In, 使用状态_In,
		 摘要_In, 入出系数_In, 单位_In, 容量_In, 数量_In, 单价_In, 金额_In, 操作员_In, 登记时间_In, 科室id_In);
	-- 修改 使用状态
	If 入出系数_In = -1 Then
		Update 暂存药品记录
		Set 使用状态 = 1
		Where No = No_In And 序号 = 序号_In And 病人id = 病人id_In And 医嘱id = 医嘱id_In And 发送号 = 发送号_In And
					药品id = 药品id_In;
	End If;
Exception
	When Others Then
		Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_暂存药品记录_Insert;
/

Create Or Replace Procedure Zl_暂存药品记录_Delete(No_In In 暂存药品记录.NO%Type) Is
Begin
  Delete 暂存药品记录 Where NO = No_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_暂存药品记录_Delete;
/

Create Or Replace Procedure Zl_暂存药品记录_Undouse
(
  No_In       In 暂存药品记录.NO%Type,
  序号_In     In 暂存药品记录.序号%Type,
  入出系数_In In 暂存药品记录.入出系数%Type,
  登记时间_In In 暂存药品记录.登记时间%Type
) Is
  n_Use 暂存药品记录.数量%Type;
Begin
  Delete 暂存药品记录 Where NO = No_In And 序号 = 序号_In And 入出系数 = 入出系数_In And 登记时间 = 登记时间_In;
  Select Sum(Nvl(数量, 0)) Into n_Use From 暂存药品记录 Where NO = No_In And 序号 = 序号_In And 入出系数 = 入出系数_In;
  If Nvl(n_Use, 0) = 0 Then
    Update 暂存药品记录 Set 使用状态 = 0 Where NO = No_In And 序号 = 序号_In And 入出系数 = 1;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_暂存药品记录_Undouse;
/

Create Or Replace Procedure Zl_暂存药品记录_Adviceused
(
	No_In       In 暂存药品记录.No%Type,
	序号_In     In 暂存药品记录.序号%Type,
	医嘱id_In   In 暂存药品记录.医嘱id%Type,
	发送号_In   In 暂存药品记录.发送号%Type,
	药品id_In   In 暂存药品记录.药品id %Type,
	数量_In     In 暂存药品记录.数量%Type,
	操作员_In   In 暂存药品记录.操作员%Type,
	登记时间_In In 暂存药品记录.登记时间%Type
) Is
Begin
	Insert Into 暂存药品记录
		(No, 序号, 病人id, 医嘱id, 发送号, 药品id, 药品名称, 规格, 执行分类, 使用状态, 摘要, 入出系数, 单位, 容量, 数量,
		 单价, 金额, 操作员, 登记时间, 科室id)
		Select b.No, b.序号, b.病人id, b.医嘱id, b.发送号, b.药品id, b.药品名称, b.规格, b.执行分类, 1, b.摘要, -1, b.单位,
					 b.容量, 数量_In, b.单价, 数量_In * b.单价, 操作员_In, 登记时间_In, b.科室id
		From 暂存药品记录 b
		Where b.入出系数 = 1 And Nvl(b.使用状态, 0) = 0 And b.药品id = 药品id_In And b.医嘱id = 医嘱id_In And
					b.发送号 = 发送号_In;

	Update 暂存药品记录
	Set 使用状态 = 1
	Where No = No_In And 序号 = 序号_In And 医嘱id = 医嘱id_In And 发送号 = 发送号_In And 药品id = 药品id_In;
Exception
	When Others Then
		Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_暂存药品记录_Adviceused;
/
-------------------------------------------------------------------------------------------------------------------
-- 以下是修改原有的过程
-------------------------------------------------------------------------------------------------------------------
Create Or Replace Procedure Zl_病人医嘱执行_Insert
(
  医嘱id_In   病人医嘱执行.医嘱id%Type,
  发送号_In   病人医嘱执行.发送号%Type,
  要求时间_In 病人医嘱执行.要求时间%Type,
  本次数次_In 病人医嘱执行.本次数次%Type,
  执行摘要_In 病人医嘱执行.执行摘要%Type,
  执行人_In   病人医嘱执行.执行人%Type,
  执行时间_In 病人医嘱执行.执行时间%Type
) Is
  --除了要执行的主记录,还包含了附加手术,检查部位的记录 
  --手术麻醉,中药煎法,采集方法单独控制,检验组合只填写在第一个项目上，但执行状态相同 
  Cursor c_Advice Is
    Select A.医嘱id, B.相关id, B.诊疗类别
    From 病人医嘱发送 A, 病人医嘱记录 B
    Where (B.ID = 医嘱id_In Or (B.相关id = 医嘱id_In And B.诊疗类别 In ('F', 'D'))) And A.医嘱id = B.ID And
          A.发送号 + 0 = 发送号_In;

  v_Temp Varchar2(255);
  --v_人员编号 病人费用记录.操作员编号%Type;
  v_人员姓名 病人费用记录.操作员姓名%Type;

  v_Date  Date;
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  v_Temp := Zl_Identity;
  v_Temp := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  --v_人员编号 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);

  Select Sysdate Into v_Date From Dual;

  For r_Advice In c_Advice Loop
    Insert Into 病人医嘱执行
      (医嘱id, 发送号, 要求时间, 本次数次, 执行摘要, 执行人, 执行时间, 登记时间, 登记人)
    Values
      (r_Advice.医嘱id, 发送号_In, 要求时间_In, 本次数次_In, 执行摘要_In, 执行人_In, 执行时间_In, v_Date, v_人员姓名);
  
    --填写了执行状态后就标记为正在执行 
    If r_Advice.诊疗类别 = 'C' And r_Advice.相关id Is Not Null Then
      Update 病人医嘱发送
      Set 执行状态 = 3
      Where 发送号 + 0 = 发送号_In And 医嘱id In (Select ID From 病人医嘱记录 Where 相关id = r_Advice.相关id);
    Else
      Update 病人医嘱发送 Set 执行状态 = 3 Where 医嘱id = r_Advice.医嘱id And 发送号 + 0 = 发送号_In;
    End If;
    --Beging 2007-01-04 删除时，标记费用记录中的执行状态，不允许退费
    Update 病人费用记录
    Set 执行状态 = 2, 执行时间 = 执行时间_In, 执行人 = v_人员姓名
    Where Nvl(执行状态, 0) = 0 And 收费类别 Not In ('5', '6', '7') And
          医嘱序号 + 0 In
          (Select ID
           From 病人医嘱记录
           Where ID = Nvl(r_Advice.医嘱id, r_Advice.相关id) Or 相关id = Nvl(r_Advice.医嘱id, r_Advice.相关id)) And
          (记录性质, NO) In
          (Select 记录性质, NO
           From 病人医嘱附费
           Where 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In
           Union All
           Select 记录性质, NO From 病人医嘱发送 Where 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In);
    --End 2007-01-04 删除时，标记费用记录中的执行状态，不允许退费  
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病人医嘱执行_Insert;
/

Create Or Replace Procedure Zl_病人医嘱执行_Delete
(
  医嘱id_In   病人医嘱执行.医嘱id%Type,
  发送号_In   病人医嘱执行.发送号%Type,
  执行时间_In 病人医嘱执行.执行时间%Type
) Is
  --除了要执行的主记录,还包含了附加手术,检查部位的记录 
  --手术麻醉,中药煎法,采集方法单独控制,检验组合只填写在第一个项目上，但执行状态相同 
  Cursor c_Advice Is
    Select A.医嘱id, B.相关id, B.诊疗类别
    From 病人医嘱发送 A, 病人医嘱记录 B
    Where (B.ID = 医嘱id_In Or (B.相关id = 医嘱id_In And B.诊疗类别 In ('F', 'D'))) And A.医嘱id = B.ID And
          A.发送号 + 0 = 发送号_In;

  v_Count Number;
Begin
  For r_Advice In c_Advice Loop
    Delete From 病人医嘱执行 Where 医嘱id = r_Advice.医嘱id And 发送号 + 0 = 发送号_In And 执行时间 = 执行时间_In;
    --Beging 2007-01-04 删除时，清除费用记录中的执行状态，可以退费
    Update 病人费用记录
    Set 执行状态 = 0, 执行时间 = Null, 执行人 = Null
    Where Nvl(执行状态, 0) = 2 And 收费类别 Not In ('5', '6', '7') And
          医嘱序号 + 0 In
          (Select ID
           From 病人医嘱记录
           Where ID = Nvl(r_Advice.医嘱id, r_Advice.相关id) Or 相关id = Nvl(r_Advice.医嘱id, r_Advice.相关id)) And
          (记录性质, NO) In
          (Select 记录性质, NO
           From 病人医嘱附费
           Where 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In
           Union All
           Select 记录性质, NO From 病人医嘱发送 Where 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In);
    --End 2007-01-04 删除时，清除费用记录中的执行状态，可以退费  
  End Loop;

  --如果执行情况删完了就标记执行状态为未执行 
  Select Count(*) Into v_Count From 病人医嘱执行 Where 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In;
  If Nvl(v_Count, 0) = 0 Then
    For r_Advice In c_Advice Loop
      If r_Advice.诊疗类别 = 'C' And r_Advice.相关id Is Not Null Then
        Update 病人医嘱发送
        Set 执行状态 = 0
        Where 发送号 + 0 = 发送号_In And 医嘱id In (Select ID From 病人医嘱记录 Where 相关id = r_Advice.相关id);
      Else
        Update 病人医嘱发送 Set 执行状态 = 0 Where 医嘱id = r_Advice.医嘱id And 发送号 + 0 = 发送号_In;
      End If;
    End Loop;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病人医嘱执行_Delete;
/

