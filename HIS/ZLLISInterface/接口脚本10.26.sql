--86342:王振涛,2015-07-08,添加虚拟模块权限
Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select &n_System,1215,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0 Union All
Select '所有科室',1,'可不可以操作其他科室的检验单的权限。',1 From Dual Union All
Select '直接申请',2,'可不经过医生或护士站直接在检验站开出检验申请。',1 From Dual Union All
Select '生成主费',3,'产生由检验申请单直接引起的费用。',1 From Dual Union All
Select '附费处理',4,'添加检验附加的费用单据。',1 From Dual Union All
Select '核收标本',5,'核收检验申请单,并确定检验人及检验时间。',1 From Dual Union All
Select '核收撤消',6,'是否可以撤消已经核收的标本。',1 From Dual Union All
Select '审核标本',7,'对已经检验的标本进行审核确认。',1 From Dual Union All
Select '未收费审核',8,'能够审核未收取检验相关费用的检验单。',1 From Dual Union All
Select '审核取消',9,'对已经审核了的标本进行撤消处理。',1 From Dual Union All
Select '审核限制',10,'审核时不允许报告人和审核人同一人',0 From Dual Union All
Select '强制审核规则',10,'有权限时，可以对违反了不能跳过的规则进行审核；无权限时，如果是不能跳过的规则，不能审核。',1 From Dual Union All
Select '修改标本号',11,'对正在检验的标本进行标本号调整。',1 From Dual Union All
Select '修改他人结果',12,'能够填写或修改非本人检验的报告结果。',1 From Dual Union All
Select '修改往日结果',13,'能够填写或修改非本日检验的报告结果。',1 From Dual Union All
Select '无主处理',14,'处理没有申请信息的检验标本（质控，重做，删除）。',1 From Dual Union All
Select '报告打印',15,'是否可以打印审核后的检验报告。',1 From Dual Union All
Select '批量打印',16,'是否可以打印审核后的成批检验报告。',1 From Dual Union All
Select '无主打印',17,'打印没有申请信息的检验报告。',1 From Dual Union All
Select '综合查询',18,'以多种条件组合查询正在检验或已经检验的检验单及相关情况。',1 From Dual Union All
Select '参数设置',19,'进行公共模块参数设置的权限',1 From Dual Union All
Select '修改质控结果',20,'有此权限才能对质控结果进行修改',1 From Dual Union All
Select '修改比对结果',21,'有此权限才能对比对结果进行修改',1 From Dual Union All
Select '24小时审核取消',22,'允许操作人员取消24小时内的审核报告单',1 From Dual Union All
Select '通讯参数设置',23,'是否允许操作人员修改仪器的通讯参数',1 From Dual Union All
Select '报告填写',24,'是否允许填写和修改检验结果',1 From Dual Union All
Select '未审核打印',25,'允许打印未审核的报告单',1 From Dual Union All
Select '查看其他科室报告',26,'有此权限，则可以查看其他科室的报告。',1 From Dual Union All
Select '已审已打印可回滚',26,'有此权限，则可以回滚已审核并且已打印的报告。',1 From Dual Union All
Select '未收费核收',27,'能够核收未收取检验相关费用的检验单。',1 From Dual Union All
Select '删除无主标本',28,'有些权限时才能删除无主标本。',1 From Dual Union All
Select '批量强制审核规则',29,'批量时有权限时，可以对违反了不能跳过的规则进行审核；无权限时，如果是不能跳过的规则，不能审核。',1 From Dual Union All
Select '修改病人信息',30,'允许修改病人信息（姓名、性别、年龄）。',1 From Dual Union All
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0) A;

Create Or Replace Procedure Zl_检验报告单_Insert
(
	Id_In   In 病人医嘱记录.Id%Type,
	Type_In In Number -- 0=新增 1=删除
) Is
	--HIS和其他LIS接口使用
	v_主页id     病人医嘱记录.主页id%Type;
	v_医嘱id     病人医嘱记录.Id%Type;
	v_开嘱科室id 病人医嘱记录.开嘱科室id%Type;
	v_病人来源   检验标本记录.病人来源%Type;
	v_病人id     检验标本记录.病人id%Type;
	v_婴儿       检验标本记录.婴儿%Type;
	v_病历文件id 病历单据应用.病历文件id%Type;
	v_病历文件名 病历文件列表.名称%Type;
	v_文件id     电子病历内容.文件id%Type;
	v_Temp       Varchar2(255);
	v_人员部门id 部门人员.部门id%Type;
	v_人员编号   人员表.编号%Type;
	v_人员姓名   人员表.姓名%Type;
	v_执行       Number;
	v_No         病人医嘱发送.No%Type;
	v_性质       病人医嘱发送.记录性质%Type;
	v_序号       Varchar2(1000);
	v_查阅       Number;
	v_Error      Varchar2(255);
	Err_Custom Exception;
	--查找当前标本的相关申请
	Cursor c_Samplequest Is
		Select Distinct Id As 医嘱id From 病人医嘱记录 Where Id_In In (Id, 相关id);

	--未审核的费用行(不包含药品)
	Cursor c_Verify(v_医嘱id In Number) Is
		Select Distinct 记录性质, No, 序号
		From 住院费用记录
		Where 收费类别 Not In ('5', '6', '7') And
					医嘱序号 + 0 In (Select Id From 病人医嘱记录 Where v_医嘱id In (Id, 相关id)) And 记帐费用 = 1 And
					记录状态 = 0 And 价格父号 Is Null And
					(记录性质, No) In
					(Select 记录性质, No
					 From 病人医嘱附费
					 Where 医嘱id = v_医嘱id
					 Union All
					 Select 记录性质, No
					 From 病人医嘱发送
					 Where 医嘱id In (Select Id From 病人医嘱记录 Where v_医嘱id In (Id, 相关id)))
		Order By 记录性质, No, 序号;

Begin
	--操作员信息:部门ID,部门名称;人员ID,人员编号,人员姓名
	v_Temp       := Zl_Identity;
	v_人员部门id := To_Number(Substr(v_Temp, 1, Instr(v_Temp, ',') - 1));
	v_Temp       := Substr(v_Temp, Instr(v_Temp, ';') + 1);
	v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
	v_人员编号   := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
	v_人员姓名   := Substr(v_Temp, Instr(v_Temp, ',') + 1);

	Select Distinct Nvl(b.主页id, 0), Nvl(b.相关id, 0), Decode(b.病人来源, 2, 2, 4, 4, 1), Nvl(b.病人id, 0),
									Nvl(b.开嘱科室id, 0), Nvl(b.婴儿, 0)
	Into v_主页id, v_医嘱id, v_病人来源, v_病人id, v_开嘱科室id, v_婴儿
	From 病人医嘱记录 b
	Where b.相关id = Id_In;

	Begin
		Select 病历文件id, c.名称
		Into v_病历文件id, v_病历文件名
		From 病人医嘱记录 a, 病历单据应用 b, 病历文件列表 c
		Where a.诊疗项目id = b.诊疗项目id And b.病历文件id = c.Id And a.相关id = v_医嘱id And b.应用场合 = v_病人来源 And
					Rownum <= 1;
	Exception
		When Others Then
			Return;
	End;

	If Type_In = 0 Then
		--新增
		--删除以前的报告记录
		Begin
			Select 病历id Into v_文件id From 病人医嘱报告 Where 医嘱id = v_医嘱id And Rownum <= 1;
			Delete 电子病历记录 Where Id = v_文件id;
			Delete 电子病历内容 Where 文件id = v_文件id;
		Exception
			When Others Then
				Select 电子病历记录_Id.Nextval Into v_文件id From Dual;
				--Insert Into 病人医嘱报告 (医嘱id, 病历id) Values (v_医嘱id, v_文件id);
		End;
	
		Insert Into 电子病历记录
			(Id, 病人来源, 病人id, 主页id, 婴儿, 科室id, 病历种类, 文件id, 病历名称, 创建人, 创建时间, 保存人, 保存时间,
			 最后版本, 签名级别)
		Values
			(v_文件id, v_病人来源, v_病人id, v_主页id, v_婴儿, v_开嘱科室id, 7, v_病历文件id, v_病历文件名, Null, Sysdate,
			 Null, Sysdate, 1, 0);
	
		Insert Into 病人医嘱报告 (医嘱id, 病历id) Values (v_医嘱id, v_文件id);
	
		Insert Into 电子病历内容
			(Id, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行)
		Values
			(电子病历内容_Id.Nextval, v_文件id, 1, 1, Null, 1, 2, Null, Null, 0, 0, 0, 0);
	
		Update 病人医嘱发送 Set 执行状态 = 1 Where 医嘱id In (Select Id From 病人医嘱记录 Where v_医嘱id In (Id, 相关id));
	
		--执行后自动审核对应的记帐划价单(不包含药品)
		Select Zl_To_Number(Nvl(Zl_Getsysparameter(81), '0')) Into v_执行 From Dual;
		--2.检查当前标本相关的申请的相关标本是否完成审核
		For r_Samplequest In c_Samplequest Loop
		
			--r_SampleQuest.医嘱id申请已经完成,处理后续环节
		
			--2.费用执行处理
			IF If v_性质 = 1 Then
			Update 门诊费用记录
			Set 执行状态 = 1, 执行时间 = Sysdate, 执行人 = v_人员姓名
			Where 收费类别 Not In ('5', '6', '7') And
						(医嘱序号, 记录性质, No) In
						(Select 医嘱id, 记录性质, No
						 From 病人医嘱附费
						 Where 医嘱id = r_Samplequest.医嘱id
						 Union All
						 Select 医嘱id, 记录性质, No
						 From 病人医嘱发送
						 Where 医嘱id In (Select Id From 病人医嘱记录 Where r_Samplequest.医嘱id In (Id, 相关id)));
			 ELSE 
			Update 住院费用记录
			Set 执行状态 = 1, 执行时间 = Sysdate, 执行人 = v_人员姓名
			Where 收费类别 Not In ('5', '6', '7') And
						(医嘱序号, 记录性质, No) In
						(Select 医嘱id, 记录性质, No
						 From 病人医嘱附费
						 Where 医嘱id = r_Samplequest.医嘱id
						 Union All
						 Select 医嘱id, 记录性质, No
						 From 病人医嘱发送
						 Where 医嘱id In (Select Id From 病人医嘱记录 Where r_Samplequest.医嘱id In (Id, 相关id)));
		         END if;
			--3.自动审核记帐
			If Nvl(v_执行, 0) = 1 Then
				For r_Verify In c_Verify(r_Samplequest.医嘱id) Loop
					If r_Verify.No || ',' || r_Verify.记录性质 <> v_No || ',' || v_性质 Then
						If v_序号 Is Not Null Then
							If v_性质 = 1 Then
								Zl_门诊记帐记录_Verify(v_No, v_人员编号, v_人员姓名, Substr(v_序号, 2));
							Elsif v_性质 = 2 Then
								Zl_住院记帐记录_Verify(v_No, v_人员编号, v_人员姓名, Substr(v_序号, 2));
							End If;
						End If;
						v_序号 := Null;
					End If;
					v_No   := r_Verify.No;
					v_性质 := r_Verify.记录性质;
					v_序号 := v_序号 || ',' || r_Verify.序号;
				End Loop;
				If v_序号 Is Not Null Then
					If v_性质 = 1 Then
						Zl_门诊记帐记录_Verify(v_No, v_人员编号, v_人员姓名, Substr(v_序号, 2));
					Elsif v_性质 = 2 Then
						Zl_住院记帐记录_Verify(v_No, v_人员编号, v_人员姓名, Substr(v_序号, 2));
					End If;
				End If;
			End If;
		
		End Loop;
	Else
		--删除
	
		v_查阅 := 0;
		Select Nvl(查阅状态, 0) Into v_查阅 From 病人医嘱报告 Where 医嘱id = v_医嘱id;
		If v_查阅 = 0 Then
			Select 病历id Into v_文件id From 病人医嘱报告 Where 医嘱id = v_医嘱id And Rownum <= 1;
			Delete 病人医嘱报告 Where 医嘱id = v_医嘱id;
			Delete 电子病历记录 Where Id = v_文件id;
			Delete 电子病历内容 Where 文件id = v_文件id;
			Update 病人医嘱发送
			Set 执行状态 = 0
			Where 医嘱id In (Select Id From 病人医嘱记录 Where v_医嘱id In (Id, 相关id));
			For r_Samplequest In c_Samplequest Loop
				--2.费用执行处理
				If v_性质 = 1 Then
				Update 门诊费用记录
				Set 执行状态 = 0, 执行时间 = Null, 执行人 = Null
				Where 收费类别 Not In ('5', '6', '7') And
							(医嘱序号, 记录性质, No) In
							(Select 医嘱id, 记录性质, No
							 From 病人医嘱附费
							 Where 医嘱id = r_Samplequest.医嘱id
							 Union All
							 Select 医嘱id, 记录性质, No
							 From 病人医嘱发送
							 Where 医嘱id In (Select Id From 病人医嘱记录 Where r_Samplequest.医嘱id In (Id, 相关id)));
				ELSE 
				Update 住院费用记录
				Set 执行状态 = 0, 执行时间 = Null, 执行人 = Null
				Where 收费类别 Not In ('5', '6', '7') And
							(医嘱序号, 记录性质, No) In
							(Select 医嘱id, 记录性质, No
							 From 病人医嘱附费
							 Where 医嘱id = r_Samplequest.医嘱id
							 Union All
							 Select 医嘱id, 记录性质, No
							 From 病人医嘱发送
							 Where 医嘱id In (Select Id From 病人医嘱记录 Where r_Samplequest.医嘱id In (Id, 相关id)));
				END if;
			End Loop;
		Else
			v_Error := '该报告已经被医生查阅，不能取消，请联系医生。';
			Raise Err_Custom;
		End If;
	End If;
Exception
	When Err_Custom Then
		Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
	When Others Then
		Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_检验报告单_Insert;
/
Create Or Replace Procedure Zl_电子病历格式_Insert
(
  Id_In   In 电子病历格式.文件id%Type,
  Txt_In  In Varchar2,
  开始_In In Number -- 1=开始
) Is
  l_Blob Blob;
Begin

  If 开始_In = 1 Then
    Delete 电子病历格式 Where 文件id = Id_In;
  End If;
  If 开始_In = 1 Then
    Update 电子病历格式 Set 内容 = Empty_Blob() Where 文件id = Id_In;
    If Sql%Rowcount = 0 Then
      Insert Into 电子病历格式 (文件id, 内容) Values (Id_In, Empty_Blob());
    End If;
  End If;
  Select 内容 Into l_Blob From 电子病历格式 Where 文件id = Id_In For Update;
  Dbms_Lob.Writeappend(l_Blob, Length(Hextoraw(Txt_In)) / 2, Hextoraw(Txt_In));
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_电子病历格式_Insert;
/
Create Or Replace Procedure Zl_检验医嘱标记_Edit
(
  Id_In   In 病人医嘱记录.Id%Type,
  Type_In In Number -- 1=核收 0=取消核收
) Is
Begin
  Update 病人医嘱发送 Set 执行状态 = Type_In Where 医嘱id In (Select ID From 病人医嘱记录 Where Id_In In (ID, 相关id));
  Update 门诊费用记录
  Set 执行状态 = Type_In, 执行时间 = Null, 执行人 = Null
  Where 医嘱序号 In (Select ID From 病人医嘱记录 Where 病人来源<>2 AND Id_In In (ID, 相关id));
Update 住院费用记录
  Set 执行状态 = Type_In, 执行时间 = Null, 执行人 = Null
  Where 医嘱序号 In (Select ID From 病人医嘱记录 Where  病人来源=2 AND Id_In In (ID, 相关id));
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_检验医嘱标记_Edit;
/

--  2009-09-21 增加体检指标保存过程
Create Or Replace Procedure Zl_体检指标_Externaledit
(
	任务id_In     In 体检任务结果.任务id%Type,
	病人id_In     In 体检任务结果.病人id%Type,
	清单id_In     In 体检任务结果.清单id%Type,
	体检指标id_In In 体检任务结果.体检指标id%Type,
	检验人_In     In 体检任务结果.检查人%Type,
	检验时间_In   In 体检任务结果.检查时间%Type,
	结果_In       In 体检任务结果.结果%Type,
	单位_In       In 体检任务结果.单位%Type,
	参考_In       In 体检任务结果.参考%Type,
	报警_In       In 体检任务结果.报警%Type
) Is
Begin

	Update 体检任务结果
	Set 结果 = 结果_In, 报警 = 报警_In, 单位 = 单位_In, 参考 = 参考_In, 检查人 = 检验人_In, 检查时间 = 检验时间_In
	Where 任务id = 任务id_In And 病人id = 病人id_In And 清单id = 清单id_In And 体检指标id = 体检指标id_In;

	Update 体检任务发送
	Set 报告人 = 检验人_In, 报告时间 = 检验时间_In, 执行状态 = 1
	Where 任务id = 任务id_In And 病人id = 病人id_In And 清单id = 清单id_In;
Exception
	When Others Then
		Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_体检指标_Externaledit;
/