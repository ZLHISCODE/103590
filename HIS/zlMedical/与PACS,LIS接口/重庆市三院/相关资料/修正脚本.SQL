
--创建数据表

Create Table 诊疗项目目录_LIS(
	诊疗项目id	NUMBER(18),
	LIS编码		Varchar2(50),
	LIS组合编码	Varchar2(50))	
	PCTFREE 15 PCTUSED 70 STORAGE (NEXT 8K PCTINCREASE 0 MAXEXTENTS UNLIMITED);

--创建约束

ALTER TABLE 诊疗项目目录_LIS ADD CONSTRAINT 诊疗项目目录_LIS_PK PRIMARY KEY (诊疗项目id) USING INDEX PCTFREE 10;

ALTER TABLE 诊疗项目目录_LIS ADD CONSTRAINT 诊疗项目目录_LIS_UQ_LIS编码 UNIQUE (LIS编码) USING INDEX PCTFREE 10;

--权限处理

--模块
insert into zlPrograms(序号,标题,说明,系统,部件) values(6001,'三院检验接口平台','完成检验项目对码及接受检验科的检验结果。',100,NULL);

--基本
Insert Into zlProgFuncs(系统,序号,功能) Values(100,6001,'基本');


--数据对码
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,6001,'数据对码','完成体检软件基础项目同接口软件基础项目间的对应关系。');

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6001,'数据对码',user,'诊疗分类目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6001,'数据对码',user,'诊疗项目目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6001,'数据对码',user,'诊疗项目别名','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6001,'数据对码',user,'诊疗项目目录_LIS','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6001,'数据对码',user,'诊疗项目目录_LIS','DELETE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6001,'数据对码',user,'诊疗项目目录_LIS','INSERT');

--接受结果
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,6001,'接受结果','接受接口体检软件产生的体检任务（包括受检人员、体检项目等）。');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6001,'接受结果',user,'体检登记记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6001,'接受结果',user,'合约单位','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6001,'接受结果',user,'体检人员档案','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6001,'接受结果',user,'病人信息','SELECT');

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6001,'接受结果',user,'诊疗项目目录_LIS','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6001,'接受结果',user,'检验报告项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6001,'接受结果',user,'体检项目清单','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6001,'接受结果',user,'体检项目医嘱','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6001,'接受结果',user,'诊治所见项目','SELECT');

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6001,'接受结果',user,'ZL_ZLLIS_填写结果','EXECUTE');

-----------------------------------------------------------------------------------------------------------------------------------------------------
--功能：取消已检验完成申请单为未完成状态
--参数：申请id_IN		检验申请单id
--说明：需要取消完成时调用
-----------------------------------------------------------------------------------------------------------------------------------------------------
CREATE OR REPLACE Procedure ZL_ZLLIS_清除结果(
	体检单号_IN	IN	体检登记记录.体检号%TYPE,
	病人id_IN	IN	病人医嘱记录.病人ID%TYPE) 
IS	
	CURSOR c_Advice IS 
		SELECT ID FROM 病人医嘱记录 WHERE 诊疗类别='C' AND 病人来源=4 And 挂号单=体检单号_IN And 病人id=病人id_IN
		UNION 
		SELECT 相关ID As ID FROM 病人医嘱记录 WHERE 诊疗类别='C' AND 病人来源=4 And 挂号单=体检单号_IN And 病人id=病人id_IN;

	v_报告id Number(18);
Begin
	
	
	FOR r_Advice In c_Advice LOOP
		v_报告id:=0;
		begin
			Select nvl(报告id,0) into v_报告id from 病人医嘱发送 Where 医嘱id=r_Advice.ID AND ROWNUM<2;
		exception
			When Others then v_报告id:=0;
		end;
		
		 Update 病人医嘱发送 Set 报告id=Null Where 报告id=v_报告id;

		Delete From 病人病历记录 Where ID=v_报告id;
	End Loop;

Exception
	When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_ZLLIS_清除结果;
/

CREATE OR REPLACE PROCEDURE ZL_ZLLIS_填写结果(
	申请id_IN	IN	病人医嘱记录.ID%TYPE,
	检验项目_IN	IN 	Varchar2,
	检验项目id_IN	In	Number,
	检验结果_IN	IN	Varchar2,
	结果类型_IN	IN 	Number,
  	结果单位_IN	In 	Varchar2,
	结果标志_IN	IN	Varchar2,
	结果参考_IN	IN	Varchar2,
	填写人员_IN	IN 	Varchar2) 
IS	
	
	CURSOR c_Advice2(v_医嘱id In Number) IS SELECT * FROM 病人医嘱记录 WHERE ID=v_医嘱id;
	r_Advice2 c_Advice2%RowType;
	
	CURSOR c_File(v_File number) IS	SELECT 类型,编码,文本转储,标题文本,标题显示,标题字体,标题位置,内容字体,内容位置,嵌入方式 
					FROM 病历文件组成 A,病历元素目录 B 
					where A.病历元素id=B.ID 
					      AND A.病历文件id=v_File 
					order by A.排列序号;

	v_病历内容id		number(18);
	v_病历id		number(18);
	v_报告id		number(18);
	v_文件ID		number(18);
	v_Index			number(18);
	v_病历种类		病历文件目录.种类%Type;
	v_病历名称 		病历文件目录.名称%Type;

	v_MaxIndex		Number(18);
Begin

	v_报告id:=0;
	v_病历id:=0;
	begin
		Select nvl(报告id,0) into v_报告id 
		from 病人医嘱发送 
		Where Rownum<2 AND 医嘱id in (SELECT ID FROM 病人医嘱记录 WHERE 申请id_IN In (ID,相关id));
	exception
		When Others then v_报告id:=0;
	end;

	If Nvl(v_报告id,0)=0 then
			
		--产生病人病历记录
		Open c_Advice2(申请id_IN);
		Fetch c_Advice2 Into r_Advice2;

		--检查要填写的报告格式中是否含有检验专用纸,如果没有则返回
		v_文件ID:=0;
		begin
			Select U.ID,U.种类,U.名称 
			Into v_文件ID,v_病历种类,v_病历名称 
			From 病历文件组成 X,病历元素目录 Y,病历文件目录 U
			Where X.病历文件id in (select A.病历文件id 
						from 诊疗单据应用 A,病人医嘱记录 B 
						where A.诊疗项目id=B.诊疗项目id 
							and B.相关ID=申请id_IN
							and A.应用场合=4)
				AND X.病历元素id=Y.ID 
				AND U.ID=X.病历文件id
				AND Y.类型=4 and Y.编码='000009';
		exception
			when others then v_文件ID:=0;
		end;
			
		--有LIS专用纸,要填写
		If nvl(v_文件ID,0)>0 then 
		
			--新产生报告id
			v_病历id:=0;
			Select 病人病历记录_ID.Nextval Into v_报告id From Dual;
			
			ZL_病人病历_INSERT(v_报告id,r_Advice2.病人id,r_Advice2.主页id,r_Advice2.挂号单,r_Advice2.婴儿,r_Advice2.病人科室id,v_病历种类,v_文件ID,v_病历名称,填写人员_IN,申请id_IN);

			--按病历组成依次产生病历组成元素记录
			v_Index:=0;
			FOR r_File In c_File(v_文件ID) LOOP
				v_Index:=v_Index+1;
					
				Select 病人病历内容_ID.Nextval Into v_病历内容id From Dual;

				if r_File.类型=4 and r_File.编码='000009' then
					v_病历id:=v_病历内容id;
				end if;

				ZL_病人病历内容_INSERT(v_病历内容id,NULL,v_报告id,v_Index,r_File.类型,r_File.编码,r_File.文本转储,r_File.标题文本,r_File.标题显示,r_File.标题字体,r_File.标题位置,0,r_File.内容字体,r_File.内容位置,0,r_File.嵌入方式);
			END LOOP;

		End if;
		Close c_Advice2;
	Else
		Update 病人病历记录 Set 书写人=Decode(书写人,Null,填写人员_IN,书写人),
					书写日期=Decode(书写日期,Null,SYSDATE,书写日期)
		Where ID=v_报告id;

		--检查要已填写的报告格式中是否含有检验专用纸,并找出检验专用纸在报告中的位置,如果没有则返回
		v_病历id:=0;
		begin
			select nvl(id,0) into v_病历id from 病人病历内容 where 元素类型=4 and 元素编码='000009' and 病历记录id=v_报告id;
		exception
			when others then v_病历id:=0;
		End;	
	End If;
		
	--有Lis专用纸病历,则填写检验结果到此专用纸中
	If v_病历id>0 And v_报告id>0 Then				
		
		Delete From 病人病历所见单 Where 病历ID=v_病历id And 所见项ID=检验项目id_IN;

		v_MaxIndex:=1;
		Begin
			Select Nvl(Max(A.控件号),0)+1 Into v_MaxIndex From 病人病历所见单 A,病人病历内容 B Where A.病历id=B.ID AND B.元素类型=4 and B.元素编码='000009' and B.病历记录id=v_报告id;
		Exception
			When Others Then v_MaxIndex:=1;
		End;
		
		Insert Into 病人病历所见单(病历ID,控件号,控件类,标题,所见项ID,数值类型,计量单位,所见内容) 
		Select ID,v_MaxIndex,2,检验项目_IN,检验项目id_IN,结果类型_IN,结果单位_IN,检验结果_IN||''''||结果标志_IN||''''||结果参考_IN
		From 病人病历内容 Where 元素类型=4 and 元素编码='000009' and 病历记录id=v_报告id;			
	End If;

	--修改病人医嘱发送记录的报告id列
	If v_报告id>0 Then
		Update 病人医嘱发送 SET 报告id=v_报告id WHERE 医嘱id in (SELECT ID FROM 病人医嘱记录 WHERE 申请id_IN In (ID,相关id));	
	End If;

Exception
	When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_ZLLIS_填写结果;
/



