Alter Table 诊疗项目目录 drop Column 执行分类;
Alter Table 药品规格 drop column 容量;
Alter Table 病人医嘱执行 drop column 流水号;
Alter Table 病人医嘱执行 drop column 接单人;
Alter Table 病人医嘱执行 drop column 配药人;
Alter Table 病人医嘱执行 drop column 组数;
Alter Table 病人医嘱执行 drop column 组次;
Alter Table 病人医嘱执行 drop column 滴速;
Alter Table 病人医嘱执行 drop column 滴系数;
Alter Table 病人医嘱执行 drop column 液体量;
Alter Table 病人医嘱执行 drop column 耗时;
Alter Table 病人医嘱执行 drop column 提醒;
Alter Table 病人医嘱执行 drop column 说明;

drop table 执行打印记录;
drop table 暂存药品记录;
drop table 座位状况记录;
drop table 排队记录;

drop sequence 病人医嘱执行_流水号;

drop procedure Zl_座位状况记录_Update;
drop procedure Zl_座位状况记录_Insert;
drop procedure Zl_座位状况记录_Delete;
drop procedure Zl_座位状况记录_Setseating;
drop procedure Zl_座位状况记录_Clear;
drop procedure Zl_病人医嘱执行_Transfusion;
drop procedure Zl_病人医嘱执行_Modify;
drop procedure Zl_排队记录_Addqueue;
drop procedure Zl_排队记录_Update;
drop procedure Zl_暂存药品记录_Insert;
drop procedure Zl_暂存药品记录_Delete;
drop procedure Zl_暂存药品记录_Undouse;
drop procedure Zl_暂存药品记录_Adviceused;

--
delete zlComponent where 部件='zl9Transfusion';
delete zlPrograms where 序号=1264;
delete zlProgFuncs where 序号=1264;
delete zlProgPrivs where 序号=1264;
delete zlMenus where 标题='门诊输液注射管理';
delete 号码控制表 where 项目序号=19;

Delete zlNotices where 提醒内容='[姓名][名称]时间已到，请查看结果。' And 系统=100;

-- 皮试提醒

--------------------------
-- 还原过程(10.16.0)
CREATE OR REPLACE Procedure ZL_病人医嘱执行_Insert(
	医嘱ID_IN		病人医嘱执行.医嘱ID%Type,
	发送号_IN		病人医嘱执行.发送号%Type,
	要求时间_IN		病人医嘱执行.要求时间%Type,
	本次数次_IN		病人医嘱执行.本次数次%Type,
	执行摘要_IN		病人医嘱执行.执行摘要%Type,
	执行人_IN		病人医嘱执行.执行人%Type,
	执行时间_IN		病人医嘱执行.执行时间%Type
--参数：医嘱ID_IN=单独执行的医嘱ID，检验组合为显示的检验项目的ID。
) IS
	--除了要执行的主记录,还包含了附加手术,检查部位的记录
	--手术麻醉,中药煎法,采集方法单独控制,检验组合只填写在第一个项目上，但执行状态相同
	Cursor c_Advice IS
		Select A.医嘱ID,B.相关ID,B.诊疗类别
		From 病人医嘱发送 A,病人医嘱记录 B
		Where (B.ID=医嘱ID_IN Or (B.相关ID=医嘱ID_IN And B.诊疗类别 IN('F','D')))
			And A.医嘱ID=B.ID And A.发送号+0=发送号_IN;

    v_Temp			Varchar2(255);
    v_人员编号		病人费用记录.操作员编号%Type;
    v_人员姓名		病人费用记录.操作员姓名%Type;

	v_Date			Date;
    v_Error			Varchar2(255);
    Err_Custom		Exception;
Begin
    --当前操作人员
    v_Temp:=zl_Identity;
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_人员编号:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_人员姓名:=Substr(v_Temp,Instr(v_Temp,',')+1);

	Select Sysdate Into v_Date From Dual;

    --病人医嘱执行
	For r_Advice In c_Advice Loop
		Insert Into 病人医嘱执行(
			医嘱ID,发送号,要求时间,本次数次,执行摘要,执行人,执行时间,登记时间,登记人)
		Values(
			r_Advice.医嘱ID,发送号_IN,要求时间_IN,本次数次_IN,执行摘要_IN,执行人_IN,执行时间_IN,v_Date,v_人员姓名);

		--填写了执行状态后就标记为正在执行
		If r_Advice.诊疗类别='C' And r_Advice.相关ID IS Not NULL Then
			Update 病人医嘱发送 
				Set 执行状态=3 
			Where 发送号+0=发送号_IN And 医嘱ID IN(
				Select ID From 病人医嘱记录 Where 相关ID=r_Advice.相关ID);
		Else
			Update 病人医嘱发送 Set 执行状态=3 Where 医嘱ID=r_Advice.医嘱ID And 发送号+0=发送号_IN;
		End IF;
	End Loop;
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then Zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_病人医嘱执行_Insert;
/

CREATE OR REPLACE Procedure ZL_病人医嘱执行_Delete(
	医嘱ID_IN		病人医嘱执行.医嘱ID%Type,
	发送号_IN		病人医嘱执行.发送号%Type,
	执行时间_IN		病人医嘱执行.执行时间%Type
--参数：医嘱ID_IN=单独执行的医嘱ID，检验组合为显示的检验项目的ID。
) IS
	--除了要执行的主记录,还包含了附加手术,检查部位的记录
	--手术麻醉,中药煎法,采集方法单独控制,检验组合只填写在第一个项目上，但执行状态相同
	Cursor c_Advice IS
		Select A.医嘱ID,B.相关ID,B.诊疗类别
		From 病人医嘱发送 A,病人医嘱记录 B
		Where (B.ID=医嘱ID_IN Or (B.相关ID=医嘱ID_IN And B.诊疗类别 IN('F','D')))
			And A.医嘱ID=B.ID And A.发送号+0=发送号_IN;

	v_Count			Number;
Begin
    --病人医嘱执行
	For r_Advice In c_Advice Loop
		Delete From 病人医嘱执行 Where 医嘱ID=r_Advice.医嘱ID And 发送号+0=发送号_IN And 执行时间=执行时间_IN;
	End Loop;

	--如果执行情况删完了就标记执行状态为未执行
	Select Count(*) Into v_Count From 病人医嘱执行 Where 医嘱ID=医嘱ID_IN And 发送号+0=发送号_IN;
	If Nvl(v_Count,0)=0 Then
		For r_Advice In c_Advice Loop
			If r_Advice.诊疗类别='C' And r_Advice.相关ID IS Not NULL Then
				Update 病人医嘱发送 
					Set 执行状态=0
				Where 发送号+0=发送号_IN And 医嘱ID IN(
					Select ID From 病人医嘱记录 Where 相关ID=r_Advice.相关ID);
			Else
				Update 病人医嘱发送 Set 执行状态=0 Where 医嘱ID=r_Advice.医嘱ID And 发送号+0=发送号_IN;
			End IF;
		End Loop;
	End IF;
Exception
    When OTHERS Then Zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_病人医嘱执行_Delete;
/

commit;