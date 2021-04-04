*****************************************************************************************************************************
--脚本名称：	ZLHIS+与其他PACS系统的标准接口脚本
--
--脚本内容：	1.接口视图
--		2.接口过程
--		3.接口权限
*****************************************************************************************************************************

Connect ZLHIS/his;

Create User ZLPACS Identified by PACS Default Tablespace zl9baseitem Temporary Tablespace zltoolstmp;

--1.接口视图
----------------------------------------------------------------------------------------------------------------------------
--功能：查询检查项目信息
--返回：ID、编码、名称、适用性别、单位、适用场合、检查类型、检查部位、检查方法
--说明: 在HIS没有设定检查部位及方法时无部位方法，设定后可能出现相同项目但部位方法不同的记录
----------------------------------------------------------------------------------------------------------------------------
Create Or Replace View ZLPACS_检查项目 As 
          Select /*+ RULE */
          Distinct a.Id, a.编码, a.名称, Decode(a.适用性别, 1, '男', 2, '女', '通用') 适用性别, a.计算单位 As 单位,
                   Decode(a.服务对象, 1, '门诊', 2, '住院', '通用') 适用场合, a.操作类型 检查类型, b.部位 检查部位,
                   b.方法 检查方法
          From 诊疗项目目录 a, 诊疗项目部位 b
          Where Nvl(To_Char(a.撤档时间, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01' And a.类别 = 'D' And a.Id = b.项目id(+)
/


----------------------------------------------------------------------------------------------------------------------------
--功能：查询病人信息资料
--返回：病人id、姓名、门诊号、住院号、性别、年龄、出生日期、婚姻状况、身份证号、职业、民族、学历、当前科室、当前病区、当前床号
----------------------------------------------------------------------------------------------------------------------------
Create Or Replace View ZLPACS_病人信息 as
	Select	 /*+ rule */ A.病人id,
		A.姓名,
		A.门诊号,
		A.住院号,
		A.性别,
		A.年龄,
		A.出生日期,
		A.婚姻状况,
		A.身份证号,
		A.职业,
		A.民族,
		A.学历,
		B.名称 As 当前科室,
		C.名称 As 当前病区
	From 病人信息 A,部门表 B,部门表 C
	Where	A.当前科室ID=B.ID(+) And A.当前病区ID=C.ID(+)
/

----------------------------------------------------------------------------------------------------------------------------
--功能：查询检查申请单信息
--返回：申请id、申请科室、申请人、申请时间、申请内容、项目ID、执行部门、病人id、姓名、门诊号、住院号、性别、年龄、病人来源
--说明: 出于性能考虑，只检索当前时间-30天之内的检查申请;申请内容 格式(名称,执行标记:部位1(方法1,方法2),部位2(方法1,方法2)---)
----------------------------------------------------------------------------------------------------------------------------
Create Or Replace View ZLPACS_检查申请 As
      Select /*+ RULE */
      Distinct a.Id 申请id, e.名称 申请科室, a.开嘱医生 申请人, a.开嘱时间 As 申请时间, a.医嘱内容, a.诊疗项目id As 项目id,
               g.名称 执行部门, a.病人id, d.姓名, d.门诊号, d.住院号, d.性别, d.年龄,
               Decode(a.病人来源, 1, '门诊', 2, '住院', 3, '外来', 4, '体验') 病人来源
      From 病人医嘱记录 a, 病人医嘱发送 b, 病人信息 d, 部门表 e,
           (Select 部门id From 部门性质说明 Where 工作性质 = '检查') f, 部门表 g
      Where a.相关id Is Null And a.开嘱时间 > Sysdate - 30 And b.执行部门id = f.部门id And
            a.医嘱期效 = 1 And a.诊疗类别 = 'D' And a.Id = b.医嘱id And b.执行状态 = 0 And Nvl(b.执行过程, 0)=0 And
            a.病人id = d.病人id　and a.开嘱科室id = e.Id And b.执行部门id = g.Id
/

--2.接口过程
----------------------------------------------------------------------------------------------------------------------------
--功能：接收检查申请单
--参数：申请id_IN		数字型(18),检查申请id
--	执行间_IN		字符型(20)
--	检查号_IN		数值型(18)
--	检查设备_IN		字符型(30)
--	身高_IN			数值型(5)
--	体重_IN			数值型(5)
--	检查技师_IN		字符型(20)
--说明：当检查号不传时以HIS记录提取检查号
----------------------------------------------------------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZLPACS_接收申请(
       申请id_IN 病人医嘱发送.医嘱ID%TYPE,
       执行间_IN 病人医嘱发送.执行间%TYPE:=Null,
       检查号_IN 影像检查记录.检查号%TYPE:=NULL,
       检查设备_IN 影像检查记录.检查设备%TYPE:=Null,
       身高_IN 影像检查记录.身高%TYPE:=Null,
       体重_IN 影像检查记录.体重%TYPE:=Null,
       检查技师_IN 影像检查记录.检查技师%TYPE:=Null)
Is
  Cursor c_AdviceInfo Is 
  Select B.Id,B.相关ID,B.病人来源,A.发送号,C.姓名,zlspellcode(C.姓名) 英文名,D.操作类型
  From 病人医嘱发送 A,病人医嘱记录 B,病人信息 C,诊疗项目目录 D
  Where A.医嘱ID=B.id And (B.Id=申请id_IN Or B.相关ID=申请id_IN) And B.病人ID=C.病人ID And B.诊疗项目ID=D.Id
  Order By B.Id;
  
  v_原检查号 影像检查记录.检查号%Type;
  v_新检查号 影像检查记录.检查号%Type;
  v_姓名   影像检查记录.姓名%Type;
  v_英文名 影像检查记录.英文名%Type;
  v_影像类别 影像检查记录.影像类别%Type;
  v_发送号 影像检查记录.发送号%Type;
  v_病人来源 病人医嘱记录.病人来源%Type;
Begin
      
	For r_Advice In c_AdviceInfo Loop 
		Update 病人医嘱发送 
		Set 首次时间 = Sysdate, 末次时间=Sysdate, 执行状态 = 3, 执行过程 = 2, 执行间 = 执行间_In 
		Where 医嘱id = r_Advice.Id And 发送号 = r_Advice.发送号;
		
		If r_Advice.相关ID Is Null Then--主医嘱记录
			v_姓名:=r_Advice.姓名;
			v_英文名:=r_Advice.英文名;
			v_影像类别:=r_Advice.操作类型;
			v_发送号:=r_Advice.发送号;
			v_病人来源:=r_Advice.病人来源;
		End IF;
	End Loop;

	if (v_病人来源 <> 2)then
		Update 门诊费用记录 Set 执行状态=1,执行时间=Sysdate,执行人=检查技师_IN Where 医嘱序号=申请id_IN;
	Else
		Update 住院费用记录 Set 执行状态=1,执行时间=Sysdate,执行人=检查技师_IN Where 医嘱序号=申请id_IN;
       End if;
      
	If 检查号_IN Is Null Then --没传检查号则由HIS跟据类别生成新检查号
	     begin
		      Select /*+ rule */ 检查号 Into v_原检查号 From 影像检查记录 Where 医嘱id = 申请id_IN;
	     Exception
	          When Others Then
			Select 最大号码+1 Into v_新检查号 From 影像检查类别 Where 编码=v_影像类别;
	     End;
	End If;
           
	Update /*+ RULE */ 影像检查记录
	Set 影像类别 = v_影像类别, 检查号 = NVL(Nvl(检查号_In, v_原检查号),v_新检查号), 姓名 = v_姓名, 英文名 = v_英文名, 身高 = 身高_In,
	  体重 = 体重_In, 检查设备 = 检查设备_In, 检查技师 = 检查技师_In
	Where 医嘱id = 申请id_IN;

	If Sql%Rowcount = 0 Then 
		Insert Into 影像检查记录
		(医嘱id, 发送号, 影像类别, 检查号, 姓名, 英文名, 身高, 体重, 检查设备, 检查技师)
		Values
		(申请id_In, v_发送号, v_影像类别, NVL(Nvl(检查号_In, v_原检查号),v_新检查号), v_姓名, v_英文名, 身高_In, 体重_In, 检查设备_In,
		检查技师_In);
	End If; 

	If v_新检查号 Is NOT Null Then
	       Update 影像检查类别 Set 最大号码 = v_新检查号 Where 编码 = v_影像类别;
	End If;
Exception 
When Others Then 
Zl_Errorcenter(Sqlcode, Sqlerrm); 
End ZLPACS_接收申请;
/
----------------------------------------------------------------------------------------------------------------------------
--功能：取消检查申请单
--参数：申请id_IN		数字型(18)，检查申请id
--说明：
----------------------------------------------------------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZLPACS_取消申请(
	申请id_IN 病人医嘱发送.医嘱ID%TYPE)
IS
BEGIN
	--先清除报告
	ZLPACS_清除报告(申请id_in);

	--置申请的执行标志为未执行
	Update 病人医嘱发送
	Set 首次时间 = Null, 末次时间 = Null, 执行状态 = 0, 执行过程 = 0, 执行间 = Null
	Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = 申请id_in Or 相关id = 申请id_in);

	---置费用记录中的执行标志为未执行
	--Update 病人费用记录 Set 执行状态=0,执行时间=Null,执行人=Null Where 医嘱序号=申请id_IN;

EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZLPACS_取消申请;
/


----------------------------------------------------------------------------------------------------------------------------
--功能：完成检查，包括状态修改、报告填写
--参数：申请id_IN		检查申请单id
--	报告所见_IN		字符型(4000)
--	报告建议_IN		字符型(4000)
--	报告人_IN		填写报告人员姓名,必须是ZLHIS中的人员表中检查医生并且聘任职位设置不低于 4组理/师级
--说明：仅填写一条返回结果记录
----------------------------------------------------------------------------------------------------------------------------
CREATE OR REPLACE Procedure ZLPACS_检查完成(
	申请id_IN       病人医嘱发送.医嘱ID%TYPE,
  	报告所见_IN     电子病历内容.内容文本%TYPE,
	报告建议_IN     电子病历内容.内容文本%TYPE,
	报告人_IN       电子病历记录.创建人%TYPE)
IS

	--提取病人医嘱及报告的相关信息
    CURSOR c_Advice IS
        Select E.Id,E.病人来源,E.病人ID,E.主页ID,E.婴儿,E.病人科室ID,E.文件id, E.病历种类,E.病历名称,F.病历ID
        From (Select C.ID,C.病人来源,C.病人ID,C.主页ID,C.婴儿,C.病人科室ID,C.文件id, D.种类 病历种类, D.名称 病历名称
        	From (Select A.ID,A.病人来源,A.病人ID,A.主页ID,A.婴儿,A.病人科室ID, B.病历文件id 文件id
             				From 病人医嘱记录 A, 病历单据应用 B
             				Where A.Id=申请id_IN And A.诊疗项目id = B.诊疗项目id(+) And A.病人来源 = B.应用场合(+)) C,病历文件列表 D
        	Where C.文件id = D.Id(+)) E,病人医嘱报告 F
        Where E.Id=F.医嘱ID(+);

    --查找文件的组成元素
    CURSOR c_File(v_File number) IS
        Select A.Id, A.文件id, A.父id, A.对象序号, A.对象类型, A.对象标记, A.保留对象, A.对象属性, A.内容行次,
               A.内容文本, A.是否换行, A.预制提纲id, A.复用提纲, A.使用时机, A.诊治要素id, A.替换域, A.要素名称,
               A.要素类型, A.要素长度, A.要素小数, A.要素单位, A.要素表示, A.输入形态, A.要素值域
        From 病历文件结构 A
        Where A.文件id = v_File
        Order By A.对象序号;

    Cursor c_Report(v_电子病历记录ID Number) Is
        Select /*+ rule */ B.Id, A.内容文本
               From 电子病历内容 A, 电子病历内容 B
               Where A.文件id = v_电子病历记录ID And Nvl(A.定义提纲id, 0) <> 0 And
                     (A.内容文本 Like '%所见%' Or A.内容文本 Like '%描述%' Or A.内容文本 Like '%建议%' Or A.内容文本 Like '%意见%') And
                     B.父id = A.Id And B.是否换行 = 1;

    r_Advice	c_Advice%Rowtype;
    v_病历id      电子病历内容.文件ID%Type;
    v_病历内容id	电子病历内容.Id%Type;
    v_对象序号    电子病历内容.对象序号%Type;
    v_父ID        电子病历内容.父ID%Type;
    v_内容文本    电子病历内容.内容文本%Type;
    v_定义提纲ID  电子病历内容.定义提纲ID%Type;
    --v_格式内容    电子病历格式.内容%Type;
    v_Error Varchar2(255);
    Err_Custom Exception;
Begin

    Open c_Advice;
    Fetch c_Advice Into r_Advice;

    If Nvl(r_Advice.文件ID,0)=0 Then
        v_Error:='本次检查项目没有对应相关的检查报告，请与管理员联系！';
        Raise Err_Custom;
    Else
        If Nvl(r_Advice.病历id,0)>0 Then  ----产生过报告
            --找出检查已填写的报告提纲中含有'%所见%','%描述%,'%建议%','%意见%',并用传入的参数更新
            For r_Report In c_Report(r_Advice.病历id) Loop
                If r_Report.内容文本 Like '%所见%' Or r_Report.内容文本 Like '%描述%' Then
                    Update 电子病历内容 Set 内容文本=报告所见_IN Where ID=r_Report.Id;
                Elsif r_Report.内容文本 Like '%建议%' Or r_Report.内容文本 Like '%意见%' Then
                    Update 电子病历内容 Set 内容文本=报告建议_IN Where ID=r_Report.Id;
                End If;
            End Loop;
            --更新保存时间
            Update 电子病历记录 Set 完成时间=Sysdate,保存人=报告人_IN,保存时间=Sysdate Where ID=r_Advice.病历id;
        Else
            --产生电子病历记录
            Select 电子病历记录_ID.Nextval Into v_病历id From Dual;
            Insert Into 电子病历记录
              (Id, 病人来源, 病人id, 主页id, 婴儿, 科室id, 病历种类, 文件id, 病历名称, 创建人, 创建时间, 完成时间,
               保存人, 保存时间, 最后版本, 签名级别)
            Values
              (v_病历id, r_Advice.病人来源, r_Advice.病人id, r_Advice.主页id, r_Advice.婴儿, r_Advice.病人科室id,
               r_Advice.病历种类, r_Advice.文件id, r_Advice.病历名称, 报告人_In, Sysdate, Sysdate, 报告人_In, Sysdate, 1, 2);

            --产生医嘱报告记录
            Insert Into 病人医嘱报告 (医嘱ID,病历ID) Values(申请ID_IN,v_病历ID);

            --新产生报告内容
            For r_File In c_File(r_Advice.文件ID) Loop
                Select 电子病历内容_ID.Nextval Into v_病历内容id From Dual;
                If nvl(v_对象序号,0)=0 Then
                   v_对象序号:=r_File.对象序号;
                Else
                   v_对象序号:=v_对象序号+1;
                End If;

                If NVL(r_File.父ID,0)<>0 And (r_File.内容文本 Like '%所见%' Or r_File.内容文本 Like '%描述%') Then--所见定义行(非提纲)
                     v_内容文本:=报告所见_IN || Chr(13) || Chr(13);
                     v_定义提纲ID:=0;
                Elsif NVL(r_File.父ID,0)<>0 And (r_File.内容文本 Like '%建议%' Or r_File.内容文本 Like '%意见%') Then--建议定义行(非提纲)
                     v_内容文本:=报告建议_IN || Chr(13) || Chr(13);
                     v_定义提纲ID:=0;
                Elsif nvl(r_File.对象类型,0)=1 And NVL(r_File.父ID,0)=0 Then--提纲定义行
                     v_父ID:=v_病历内容id;
                     v_内容文本:=r_File.内容文本;
                     v_定义提纲ID:=r_File.id;
                Elsif nvl(r_File.对象类型,0)=4 And r_File.要素名称 Is Not Null Then  --自动替换要素
                     v_内容文本:=zl_replace_element_value(r_File.要素名称,r_Advice.病人ID,r_Advice.主页ID,r_Advice.病人来源,r_Advice.Id);
                     v_定义提纲ID:=0;
                Else
                    v_内容文本:=r_File.内容文本;
                    v_定义提纲ID:=0;
                End If;
                
                --报告内容单独写一行
                If NVL(r_File.父ID,0)<>0 And (r_File.内容文本 Like '%所见%' Or r_File.内容文本 Like '%描述%' Or r_File.内容文本 Like '%建议%' Or r_File.内容文本 Like '%意见%') Then--先写提纲显示名称，再写内容，同时对象序号发生变化
                   Select 电子病历内容_ID.Nextval Into v_病历内容id From Dual;
                   v_对象序号 := v_对象序号 + 1;
                    Insert Into 电子病历内容
                      (Id, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次,
                       内容文本, 是否换行, 预制提纲id, 复用提纲, 使用时机, 诊治要素id, 替换域, 要素名称, 要素类型,
                       要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域, 定义提纲id)
                    Values
                      (v_病历内容id, v_病历id, 0, 0, Decode(v_定义提纲id, 0, v_父id, Null), v_对象序号, r_File.对象类型,
                       r_File.对象标记, r_File.保留对象, 0, Null, v_内容文本, r_File.是否换行,
                       r_File.预制提纲id, r_File.复用提纲, r_File.使用时机, r_File.诊治要素id, r_File.替换域,
                       r_File.要素名称, r_File.要素类型, r_File.要素长度, r_File.要素小数, r_File.要素单位,
                       r_File.要素表示, r_File.输入形态, r_File.要素值域, Decode(v_定义提纲id, 0, Null, v_定义提纲id));
                    v_对象序号 := v_对象序号 - 1;
                    v_病历内容id:= v_病历内容id-1;
                    v_内容文本:=r_File.内容文本;
                End If;

                Insert Into 电子病历内容
                  (Id, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次,
                   内容文本, 是否换行, 预制提纲id, 复用提纲, 使用时机, 诊治要素id, 替换域, 要素名称, 要素类型, 要素长度,
                   要素小数, 要素单位, 要素表示, 输入形态, 要素值域, 定义提纲id)
                Values
                  (v_病历内容id, v_病历id, 1, 0, Decode(v_定义提纲id, 0, v_父id, Null), v_对象序号, r_File.对象类型,
                   r_File.对象标记, r_File.保留对象, r_File.对象属性, Null, v_内容文本, r_File.是否换行, r_File.预制提纲id,
                   r_File.复用提纲, r_File.使用时机, r_File.诊治要素id, r_File.替换域, r_File.要素名称, r_File.要素类型,
                   r_File.要素长度, r_File.要素小数, r_File.要素单位, r_File.要素表示, r_File.输入形态, r_File.要素值域,
                   Decode(v_定义提纲id, 0, Null, v_定义提纲id));
             End Loop;
        /* 因电子病历格式中含了内容文字格式，此种方法导入之后内容文字将不可见
        Select 内容 Into v_格式内容 From 病历文件格式 Where 文件ID=r_Advice.文件ID;
	      Insert Into 电子病历格式 (文件ID,内容) Values (v_病历id,v_格式内容);
        */
        End If;
        Update 病人医嘱发送 Set 执行状态=1,执行过程=6 Where 医嘱id=申请id_IN;--更新为 完成状态，审核过程
    End If;
    Close c_Advice;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_ErrorCenter(Sqlcode, Sqlerrm);
End ZLPACS_检查完成;
/
-----------------------------------------------------------------------------------------------------------------------------------------------------
--功能：清除指定申请单的检查报告
--参数：申请id_IN		检查申请单id
--说明：
-----------------------------------------------------------------------------------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZLPACS_清除报告(
	申请id_IN 病人医嘱发送.医嘱ID%TYPE)
IS
BEGIN
	--其它相关表有级联删除功能,会随电子病历记录一并删除
	Delete 电子病历记录 Where Id In (Select 病历ID From 病人医嘱报告 Where 医嘱ID=申请id_IN);
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZLPACS_清除报告;
/
----------------------------------------------------------------------------------------------------------------------------

--3.接口权限

--授权
Grant Connect,Resource,Create Synonym to ZLPACS;

Grant Select On ZLPACS_检查项目 To ZLPACS;
Grant Select On ZLPACS_病人信息 To ZLPACS;
Grant Select On ZLPACS_检查申请 To ZLPACS;

Grant Execute On ZLPACS_接收申请 To ZLPACS;
Grant Execute On ZLPACS_取消申请 To ZLPACS;
Grant Execute On ZLPACS_清除报告 To ZLPACS;
Grant Execute On ZLPACS_检查完成 To ZLPACS;

--创建私人同义词
Connect ZLPACS/PACS;

Create Synonym 检查项目 For ZLHIS.ZLPACS_检查项目;
Create Synonym 病人信息 For ZLHIS.ZLPACS_病人信息;
Create Synonym 检查申请 For ZLHIS.ZLPACS_检查申请;

Create Synonym ZLPACS_接收申请 For ZLHIS.ZLPACS_接收申请;
Create Synonym ZLPACS_取消申请 For ZLHIS.ZLPACS_取消申请;
Create Synonym ZLPACS_清除报告 For ZLHIS.ZLPACS_清除报告;
Create Synonym ZLPACS_检查完成 For ZLHIS.ZLPACS_检查完成;

