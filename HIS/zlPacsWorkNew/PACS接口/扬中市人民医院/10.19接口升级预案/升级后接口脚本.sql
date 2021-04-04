-------------------------------------------------------------------------------
--表结构部份
-------------------------------------------------------------------------------
-- Create table
create table PACS_TMP病人病历记录
(
  ID       NUMBER(18) not null,
  报告ID   NUMBER(18) not null,
  病人ID   NUMBER(18),
  科室ID   NUMBER(18),
  病历名称 VARCHAR2(20),
  书写人ID NUMBER(18),
  书写人   VARCHAR2(50),
  书写日期 DATE,
  审阅人ID NUMBER(18),
  审阅人   VARCHAR2(50),
  审阅日期 DATE,
  记录类型 NUMBER
)
tablespace ZL9CISREC
  pctfree 15;
-- Add comments to the columns 
comment on column PACS_TMP病人病历记录.记录类型
  is '1-由INSERT传入，回传报告人，报告时间；2-由UPDATE传入，回传审阅人，报告状态为完成；3-由"驳回"传入，修改报告状态为未完成';
-- Create/Recreate primary, unique and foreign key constraints 
alter table PACS_TMP病人病历记录
  add constraint PACS_TMP病人病历记录_PK primary key (ID)
  using index 
  tablespace ZL9CISREC
  pctfree 5;


-- 创建序列
create sequence PACS_TMP病人病历记录_ID
minvalue 1
maxvalue 999999999999999999999999999
start with 1
increment by 1
cache 20;
-------------------------------

-- Create table
create table PACS_ERR
(
  ID       NUMBER not null,
  错误号   NUMBER,
  错误描述 VARCHAR2(100),
  错误时间 DATE
)
tablespace ZL9CISREC
  pctfree 10;
-- Create/Recreate primary, unique and foreign key constraints 
alter table PACS_ERR
  add constraint PACS_ERR_PK primary key (ID)
  using index 
  tablespace ZL9BASEITEM
  pctfree 10;

-------------------------------



-------------------------------------------------------------------------------
--存储过程部份
-------------------------------------------------------------------------------
CREATE OR REPLACE Procedure Zlpacs_申请
(
  医嘱ID_IN       病人医嘱记录.ID%TYPE,
  标识号_IN       病人信息.门诊号%Type,
  姓名_In         病人信息.姓名%Type,
  性别_In         病人信息.性别%Type,
  年龄_In         病人信息.年龄%Type,
  出生日期_IN	  病人信息.出生日期%TYPE,
  国籍_In         病人信息.国籍%TYPE,
  民族_In         病人信息.民族%TYPE,
  婚姻状况_In     病人信息.婚姻状况%TYPE,
  职业_In         病人信息.职业%TYPE,
  身份证号_In     病人信息.身份证号%Type,
  工作单位_In     病人信息.工作单位%Type,
  单位邮编_In     病人信息.单位邮编%Type,
  家庭地址_In     病人信息.家庭地址%Type,
  家庭电话_In     病人信息.家庭电话%Type,
  检查项目编码_In 诊疗项目目录.编码%Type,
  标本部位_In     病人医嘱记录.标本部位%Type,
  开嘱科室id_In   病人医嘱记录.开嘱科室id%Type,
  开嘱医生_In     病人医嘱记录.开嘱医生%Type,
  开嘱时间_In     病人医嘱记录.开嘱时间%Type,
  病人来源_IN	  病人医嘱记录.病人来源%TYPE,
  病人科室ID      病人医嘱记录.病人科室ID%TYPE,
  床号_In         病案主页.入院病床%TYPE,
  记录性质_IN     病人医嘱发送.记录性质%Type,
  计费状态_IN     病人医嘱发送.计费状态%Type,
  修改_IN         Number:=0
) Is
  --病人来源_IN ：1-门诊和体检；2-住院
  --记录性质_IN：1-收费记录；2-记帐记录。
  --计费状态_IN：-1-无须计费(通常无执行和院外执行的都无须计费);0-未计费;1-已计费。
  --标识号_IN：根据病人来源确定，门诊病人用门诊号，住院病人用住院号
  Nclinicid   Number;
  Scliniccont Varchar2(40);
  Nexedeptid  Number;
  Npatientid  Number;
  Npatientid1 Number;
  Nsendno     Number;
  Scheckno    Varchar2(40);
  N_RowCount  Number;
  N_Add       Number;
  N_ClinicState Number;
  Err_Custom Exception;
  v_Error Varchar2(255);
Begin
  --判断病人信息是否已经存在，如果已经存在，则只修改病人信息,不用挂号或如入院，如果不存在则挂号或入院
  --修改成通过“检查号”提取出病人ID，然后修改基本信息
    if 修改_IN=1 then 
       --修改基本信息
       IF 病人来源_IN = 1 THEN 
          select count(*) into N_RowCount from 病人信息 a where a.门诊号=标识号_IN;
          IF N_RowCount =1  THEN
              select a.病人id into Npatientid from 病人信息 a where a.门诊号=标识号_IN;
              Zl_病人信息_Update(Npatientid,标识号_IN,'','','', 姓名_In, 
      	          性别_In, 年龄_In, 出生日期_In,'', 身份证号_In,'', 职业_In, 
                  民族_In, 国籍_In,'', 婚姻状况_In, 家庭地址_In, 家庭电话_In,
                  '','','','','',Null, 工作单位_In, 单位邮编_In,'','','',Null,Null,0);
          END IF;
       ELSE 
          select count(*) into N_RowCount from 病人信息 a where a.住院号=标识号_IN;
          IF N_RowCount =1  THEN
              select a.病人id into Npatientid from 病人信息 a where a.住院号=标识号_IN;
              Zl_病人信息_Update(Npatientid, '',标识号_IN, '','', 姓名_In, 
      	          性别_In, 年龄_In, 出生日期_In,'', 身份证号_In,'', 职业_In, 
                  民族_In, 国籍_In,'', 婚姻状况_In, 家庭地址_In, 家庭电话_In,
                  '','','','','',Null, 工作单位_In, 单位邮编_In,'','','',Null,Null,0);
   		        update 病案主页 set 入院病床=床号_In where 病人id=Npatientid;
        	END IF;
       END IF;
       
       --修改医嘱记录，查找检查项目
       BEGIN
    		    Select A.ID, A.名称, B.执行科室id
    		   	Into Nclinicid, Scliniccont, Nexedeptid
    		  	From 诊疗项目目录 A, 诊疗执行科室 B
    		  	Where A.编码 = 检查项目编码_In And A.ID = B.诊疗项目id And B.病人来源 =病人来源_IN;
    		EXCEPTION 
    		    WHEN No_Data_Found THEN 
    		  	v_Error:='检查项目编码无对应的执行科室';
    		  	Raise Err_Custom;
    		END;
        
  		  --修改PACS医嘱
        BEGIN
            select 医嘱状态 into n_ClinicState from 病人医嘱记录 where id = 医嘱ID_IN;
        EXCEPTION
            WHEN No_Data_Found THEN 
            v_Error:='未找到对应的医嘱记录';
    		  	Raise Err_Custom;
        END;
        
        update 病人医嘱记录 set 医嘱状态 = 1 where id = 医嘱ID_IN;
        ZL_病人医嘱记录_UPDATE(医嘱ID_IN,Null, 1,1,1,Nclinicid, Null,Null, 1,Scliniccont || '(' || 标本部位_In || ')',
                               '', 标本部位_In,'一次性', Null,Null, '',Null, 0,Nexedeptid, 4, 0,
                               开嘱时间_In, Null, 病人科室ID,开嘱科室id_In, 开嘱医生_In,开嘱时间_In);                                     
        update 病人医嘱记录 set 医嘱状态 = n_ClinicState where id = 医嘱ID_IN;
    ELSE 
        N_Add:=1;
        IF 病人来源_IN = 1 THEN
            select count(*) into N_RowCount from 病人信息 a where a.门诊号=标识号_IN;
            IF N_RowCount =1  THEN
                select a.病人id into Npatientid from 病人信息 a where a.门诊号=标识号_IN;
                N_Add:=3;
            ELSE
                Select 最大号码 + 1 Into Npatientid From 号码控制表 Where 项目序号 = 1;
    	  	      Select Nvl(Max(病人id), 0) + 1 Into Npatientid1 From 病人信息 Where 病人id >= Npatientid;
    	  	      If Npatientid1 > Npatientid Then
    	    	        Npatientid := Npatientid1;
    	  	      End If;
    	  	      Update 号码控制表 Set 最大号码 = Npatientid Where 项目序号 = 1;          
            END IF;    
            Zl_挂号病人病案_Insert(N_Add, Npatientid, 标识号_IN, '', '', 姓名_In, 性别_In, 年龄_In, 
                '', '', 国籍_In,民族_In, 婚姻状况_In, 职业_In, 身份证号_In,工作单位_In, Null, 
                '', 单位邮编_In, 家庭地址_In, 家庭电话_In, '', 开嘱时间_In, Null,Null, 出生日期_IN);
        ELSE
            select count(*) into N_RowCount from 病人信息 a where a.住院号=标识号_IN;
            IF N_RowCount =1  THEN
                select a.病人id into Npatientid from 病人信息 a where a.住院号=标识号_IN;
                N_Add:=0;
            ELSE
                Select 最大号码 + 1 Into Npatientid From 号码控制表 Where 项目序号 = 1;
    	  	      Select Nvl(Max(病人id), 0) + 1 Into Npatientid1 From 病人信息 Where 病人id >= Npatientid;
    	  	      If Npatientid1 > Npatientid Then
    	    	        Npatientid := Npatientid1;
    	  	      End If;
    	  	      Update 号码控制表 Set 最大号码 = Npatientid Where 项目序号 = 1;  
            END IF;
            Zl_入院病案主页_Insert(0,0, Npatientid,标识号_IN,Null, 姓名_In, 性别_In, 年龄_In, 
              	'', 出生日期_IN, 国籍_In, 民族_In, '', 婚姻状况_In, 职业_In, '', 身份证号_In, 
               	'', 家庭地址_In, '', 家庭电话_In, '', '', '', '', 工作单位_In, Null, '', 
               	单位邮编_In, '', '', '', Null, Null, Null, Null, '', '', '', '', '', '',
               	Null,Null, '', '',Null,Null, '',Null,Null, '',Null, '', '',
                N_Add,'',Null,Null);
            update 病案主页 set 入院病床=床号_In,出院日期=sysdate where 病人id=Npatientid;    
        END IF;    
        
        Select 最大号码 + 1 Into Nsendno From 号码控制表 Where 项目序号 = 10;
  		  Update 号码控制表 Set 最大号码 = Nsendno Where 项目序号 = 10;
  		  Scheckno := Nextno(13);
        
    		BEGIN
    		    Select A.ID, A.名称, B.执行科室id
    		   	Into Nclinicid, Scliniccont, Nexedeptid
    		  	From 诊疗项目目录 A, 诊疗执行科室 B
    		  	Where A.编码 = 检查项目编码_In And A.ID = B.诊疗项目id And B.病人来源 =病人来源_IN;
    		EXCEPTION 
    		    WHEN No_Data_Found THEN 
    		  	v_Error:='检查项目编码无对应的执行科室';
    		  	Raise Err_Custom;
    		END;
  		  --PACS医嘱
  	  
  	  	Zl_病人医嘱记录_Insert(医嘱ID_IN, Null, 1, 病人来源_IN, Npatientid, 1, 0, 1, 1, 'D', Nclinicid, Null, Null, Null, 1,
  	                         Scliniccont || '(' || 标本部位_In || ')', '', 标本部位_In, '一次性', Null, Null, '', Null, 0,
  	                         Nexedeptid, 4, 0, Sysdate + 1 / (24 * 3600), Null, 病人科室ID, 开嘱科室id_In, 开嘱医生_In,
  	                         开嘱时间_In,1,医嘱ID_IN);
  	  	    
  	  	Zl_病人医嘱发送_Insert(医嘱ID_IN, Nsendno, 记录性质_IN, Scheckno, 1, 1, Null, Null, Sysdate + 1 / (24 * 3600), 0, Nexedeptid, 计费状态_IN, 1);
    END IF;     
EXCEPTION
   WHEN Err_Custom THEN
    	Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
   WHEN OTHERS THEN
    	Zl_Errorcenter(SQLCODE, SQLERRM);
End Zlpacs_申请;
/

----------------------------

CREATE OR REPLACE Procedure Zlpacs_开始检查
(
  执行间_IN   病人医嘱发送.执行间%Type,
  检查号_IN	  影像检查记录.检查号%Type,
  医嘱ID_IN	  影像检查记录.医嘱ID%Type,
  标识号_IN   病人信息.门诊号%Type,
  影像类别_IN 影像检查记录.影像类别%Type,
  姓名_IN     影像检查记录.姓名%Type,
  英文名_IN   影像检查记录.英文名%Type,
  性别_IN     影像检查记录.性别%Type,
  年龄_IN     影像检查记录.年龄%Type,
  出生日期_IN 影像检查记录.出生日期%Type,
  身高_IN     影像检查记录.身高%Type,
  体重_IN     影像检查记录.体重%Type,
  检查设备_IN 影像检查记录.检查设备%Type,
  电话_IN     影像检查记录.联系电话%Type:=Null,
  匹配方式_IN Number:=1,
  修改_IN     Number:=0
) Is
  --修改_IN: 0-开始检查；1-修改开始检查信息
	--匹配方式_IN：1-检查号匹配；2-门诊/住院号匹配；3-检查标识（医嘱ID）匹配
	--内部参数
	
  N_标识号 		Number;
  V_检查UID  	影像检查记录.检查UID%Type;
  N_RowCount  Number;
  Nsendno     Number;
  Err_Custom Exception;
  v_Error Varchar2(255);
BEGIN
     BEGIN	
          select D.发送号 into Nsendno from 病人医嘱发送 D where D.医嘱ID = 医嘱ID_IN;
     EXCEPTION
          WHEN No_Data_Found THEN
          	  v_Error:='医嘱ID不正确，未找到对应医嘱。';
          	  Raise Err_Custom;
     END;
     --开始影像检查
     ZL_影像检查_BEGIN(执行间_IN,检查号_IN,医嘱ID_IN,Nsendno,影像类别_IN,姓名_IN,
                    英文名_IN,性别_IN, 年龄_IN, 出生日期_IN, 身高_IN, 体重_IN,1,1, 
                    检查设备_IN, 修改_IN, 电话_IN);

  	 --查找提前进行的检查 '将图像和检查自动匹配
  	 --查找根据匹配方式，查找图像的检查UID
  	 IF 匹配方式_IN=1 THEN
  	     N_标识号:= 检查号_IN;
  	 ELSE 
         IF 匹配方式_IN=2 THEN
		         N_标识号:=标识号_IN;
	       ELSE
		         N_标识号:= 医嘱ID_IN;
	       END IF;
  	 END IF;
     
     select count(*) into N_RowCount from 影像临时记录 a  
          Where a.检查号= N_标识号 And a.影像类别=影像类别_IN;
     IF n_Rowcount =1 THEN 
          Select A.检查UID into V_检查UID From 影像临时记录 a  
               Where a.检查号= N_标识号 And a.影像类别=影像类别_IN;
  	      ZL_影像检查_SET(医嘱ID_IN, Nsendno, V_检查UID);
     END IF;
     EXCEPTION
         WHEN Err_Custom THEN
    	       Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
         WHEN OTHERS THEN
    	       Zl_Errorcenter(SQLCODE, SQLERRM);
END Zlpacs_开始检查;


/
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