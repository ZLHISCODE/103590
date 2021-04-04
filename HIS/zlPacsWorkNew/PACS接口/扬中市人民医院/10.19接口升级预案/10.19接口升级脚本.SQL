-------------------------------------------------------------------------------
--本脚本支持扬中接口从10.17升级到10.19
--表结构部份
-------------------------------------------------------------------------------
-------------------------------

-------------------------------------------------------------------------------
--存储过程部份
-------------------------------------------------------------------------------
----------------------------
create or replace procedure ZLPACS_取消申请
(
  医嘱ID_IN	  影像检查记录.医嘱ID%Type
) is
  N_RowCount      Number;
  N_ExecState     Number;
  N_ExecProcess   Number;
  Err_Custom      Exception;
  v_Error         Varchar2(255);
begin
    --只有符合以下条件的申请可以被取消
    --1.正在进行的检查(医嘱执行状态=3，执行过程=2)
    --2.没有关联图象（影像检查UID.检查UID为空）  
    BEGIN
        select 执行状态,执行过程 into N_ExecState, N_ExecProcess 
            from 病人医嘱发送 where 医嘱ID = 医嘱ID_IN;
    EXCEPTION
        WHEN No_Data_Found THEN 
    		  	v_Error:='没有符合条件可以取消的医嘱记录';
    		  	Raise Err_Custom;
    END;
    IF (N_ExecState =3 AND N_ExecProcess = 2) THEN 
        select Count(*) into N_RowCount from 影像检查记录 
            where 医嘱ID = 医嘱ID_IN and 检查UID is null;
        IF N_RowCount=1 THEN
            update 病人医嘱发送 set 执行状态 = 2 where 医嘱ID = 医嘱ID_IN;
        ELSE
            v_Error:='检查已经关联图像，无法取消，请先取消医嘱关联的图像';
    		    Raise Err_Custom;
        END IF;
    ELSE
        v_Error:='检查已经完成或者还没有开始，无法取消，请先将医嘱回退到正在进行状态';
    		Raise Err_Custom;
    END IF;
    EXCEPTION
        WHEN Err_Custom THEN
   	       Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
        WHEN OTHERS THEN
   	       Zl_Errorcenter(SQLCODE, SQLERRM);
end ZLPACS_取消申请;


/
----------------------------
create or replace procedure ZLPACS_恢复申请
(
    医嘱ID_IN	  影像检查记录.医嘱ID%Type
)  is
    N_RowCount      Number;
    N_ExecState     Number;
    N_ExecProcess   Number;
    Err_Custom      Exception;
    v_Error         Varchar2(255);
begin
    --只有符合以下条件的申请可以被恢复
    --1.原来正在进行，但是被取消(拒绝)的检查(医嘱执行状态=2，执行过程=2)
    --2.有相关联的影像检查记录
    BEGIN
        select 执行状态,执行过程 into N_ExecState, N_ExecProcess 
            from 病人医嘱发送 where 医嘱ID = 医嘱ID_IN;
    EXCEPTION
        WHEN No_Data_Found THEN 
    		  	v_Error:='没有符合条件可以恢复的医嘱记录';
    		  	Raise Err_Custom;
    END;
    IF N_ExecState =2 AND N_ExecProcess = 2 THEN 
        select Count(*) into N_RowCount from 影像检查记录 
            where 医嘱ID = 医嘱ID_IN;
        IF N_RowCount=1 THEN
            update 病人医嘱发送 set 执行状态 = 3 where 医嘱ID = 医嘱ID_IN;
        ELSE
            v_Error:='没有找到对应的影像检查记录，无法恢复';
    		    Raise Err_Custom;
        END IF;
    ELSE
        v_Error:='医嘱执行过程和执行状态不正确，无法恢复';
    		Raise Err_Custom;
    END IF;
    EXCEPTION
        WHEN Err_Custom THEN
   	       Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
        WHEN OTHERS THEN
   	       Zl_Errorcenter(SQLCODE, SQLERRM);
end ZLPACS_恢复申请;
/
----------------------------

-------------------------------------------------------------------------------
--触发器部份
-------------------------------------------------------------------------------

--删除原有触发器
Drop Trigger TBI_ZLPACS_病人病历记录_INSERT;
Drop Trigger TBI_ZLPACS_病人病历记录_UPDATE;
Drop Trigger Zltg_产生驳回记录;

--新增加的触发器
create or replace trigger TBI_ZLPACS_病人医嘱发送_UPDATE
  after update on 病人医嘱发送
  for each row
declare
      -- local variables here
      N_RowCount Number;
      N_WriteDoctorNo 人员表.编号%Type;
      N_CheckDoctorNo 人员表.编号%Type;
      N_WriteDoctor   电子病历记录.创建人%Type;
      N_CheckDortor   电子病历记录.保存人%Type;
      N_WriteTime     电子病历记录.保存时间%Type;
      N_SignClass     电子病历记录.签名级别%Type;
      N_ID Number;
begin
      --判断是否将执行过程修改成4－报告填写；5-报告审核；6-报告完成
      IF :NEW.执行过程 =4 OR :NEW.执行过程 =5 OR :NEW.执行过程 =6  THEN
      	   Select c.创建人 As 书写人,c.保存人 As 审核人,c.创建时间 As 报告时间,c.签名级别
	   			 into N_WriteDoctor,N_CheckDortor,N_WriteTime,N_SignClass
	   			 From 病人医嘱记录 a ,病人医嘱报告 b,电子病历记录 c
	   			 Where a.Id=b.医嘱ID And b.病历Id =c.Id And a.id=:NEW.医嘱ID
	   			 order by c.最后版本 Desc;
	   			 if N_RowCount = 1 then
      	   	--有报告，才查找和记录报告人信息
      	   	--创建记录ID
      			 			Select PACS_TMP病人病历记录_ID.Nextval Into N_ID From Dual;

	   							--查找书写医生编号
	      					select count(*) into N_RowCount from 人员表 A where a.姓名=N_WriteDoctor;
	      					if N_RowCount =1 then
	         					 select 编号 into N_WriteDoctorNo from 人员表 A where a.姓名=N_WriteDoctor;
        	      	else
        	         	N_WriteDoctorNo:='9999';
        	      	end if;

									if N_SignClass >=2 then
      	      		--查找审阅医生编号
      		      	select count(*) into N_RowCount from 人员表 A where a.姓名=N_CheckDortor;
      		      	if N_RowCount =1 then
      		         	select 编号 into N_CheckDoctorNo from 人员表 A where a.姓名=N_CheckDortor;
      		      	else
      		         	N_CheckDoctorNo:='9999';
      		      	end if;
      		      	--插入临时病人病历记录表
      	      		insert into PACS_tmp病人病历记录(id,报告id,科室ID,书写人ID,书写人,
      	             		书写日期,审阅人ID,审阅人,记录类型)
      	             		values(N_ID,:NEW.医嘱id,:NEW.执行部门ID,N_WriteDoctorNo,
      	             		N_WriteDoctor,N_WriteTime,N_CheckDoctorNo,N_CheckDortor,2);
            	   	else
            	   		--插入临时病人病历记录表
            	      		insert into PACS_tmp病人病历记录(id,报告id,科室ID,书写人ID,书写人,
            	             		书写日期,记录类型)
            	             		values(N_ID,:NEW.医嘱id,:NEW.执行部门ID,N_WriteDoctorNo,
            	             		N_WriteDoctor,N_WriteTime,1);
            	   	End If;
	   					end if;
      END IF;
exception
       when others then
            null;
end TBI_ZLPACS_病人医嘱发送_UPDATE;

/

-------------------------------------------------------------------------------
--权限部份
-------------------------------------------------------------------------------

Commit;