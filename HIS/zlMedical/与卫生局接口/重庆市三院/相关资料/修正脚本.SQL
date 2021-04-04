
--创建数据表

Drop Table 体检人员档案_干保;
Drop Table 体检登记记录_干保;
--Drop Table 体检诊断建议_干保;

Drop Table 诊治所见项目_干保;
Drop Table 诊疗项目目录_干保;
Drop Table 体检组别_干保;

--Create Table 体检检查项目(
--	诊疗项目id	NUMBER(18),
--	诊治项目id	NUMBER(18),
--	排列顺序	Number(5))	
--	PCTFREE 15 PCTUSED 70;

--Create Table 部门表_干保(
--	部门id		NUMBER(18),
--	干保编码	Varchar2(30),
--	干保名称	Varchar2(50))	
--	PCTFREE 15 PCTUSED 70;

Create Table 诊疗项目目录_干保(
	诊疗项目id	Number(18),
	干保编码	Varchar2(30),
	干保名称	Varchar2(50),
	组合科室	Varchar2(50))	
	PCTFREE 15 PCTUSED 70;

Create Table 诊治所见项目_干保(
	诊治项目id	NUMBER(18),
	干保编码	Varchar2(30),
	干保名称	Varchar2(50),
	项目分支	Varchar2(50),
	项目方法	Varchar2(50))	
	PCTFREE 15 PCTUSED 70;

Create Table 体检诊断建议_干保(
	结论id		NUMBER(18),
	干保编码	Varchar2(30),
	干保名称	Varchar2(50),
	疾病编码	Varchar2(50))	
	PCTFREE 15 PCTUSED 70;

Create Table 体检登记记录_干保(
	登记id		NUMBER(18),
	任务包号	Varchar2(20),
	任务包名	Varchar2(50),
	发送状态	Number(3))	
	PCTFREE 15 PCTUSED 70;

--Create Table 病人信息_干保(
--	病人id		NUMBER(18),
--	保健号		varchar2(30))	
--	PCTFREE 15 PCTUSED 70;

Create Table 体检组别_干保(
	登记id		NUMBER(18),
	组别名称	Varchar2(50),
	套餐编码	Varchar2(50),
	套餐序号	Varchar2(50),
	套餐名称	Varchar2(50))	
	PCTFREE 15 PCTUSED 70;

Create Table 体检人员档案_干保(
	登记id		NUMBER(18),
	病人id		NUMBER(18),
	任务包号	Varchar2(20),
	人员序号	Varchar2(20),
	单位编码	Varchar2(50),
	单位名称	Varchar2(50),
	在职情况	Varchar2(50),
	任职级别	Varchar2(50))	
	PCTFREE 15 PCTUSED 70;

Create Table 体检项目排列(
	诊疗项目id	Number(18),
	排列性质	Number(3),
	排列顺序	Number(18))	
	PCTFREE 15 PCTUSED 70;

--创建约束

--ALTER TABLE 部门表_干保 ADD CONSTRAINT 部门表_干保_PK PRIMARY KEY (部门id) USING INDEX PCTFREE 10;
--ALTER TABLE 部门表_干保 ADD CONSTRAINT 部门表_干保_UQ_干保编码 UNIQUE (干保编码) USING INDEX PCTFREE 10;

Alter Table 体检组别
ALTER TABLE 诊疗项目目录_干保 ADD CONSTRAINT 诊疗项目目录_干保_PK PRIMARY KEY (诊疗项目id) USING INDEX PCTFREE 10;
ALTER TABLE 诊疗项目目录_干保 ADD CONSTRAINT 诊疗项目目录_干保_UQ_干保编码 UNIQUE (干保编码) USING INDEX PCTFREE 10;

ALTER TABLE 诊治所见项目_干保 ADD CONSTRAINT 诊治所见项目_干保_PK PRIMARY KEY (诊治项目id) USING INDEX PCTFREE 10;
ALTER TABLE 诊治所见项目_干保 ADD CONSTRAINT 诊治所见项目_干保_UQ_干保编码 UNIQUE (干保编码) USING INDEX PCTFREE 10;

ALTER TABLE 体检诊断建议_干保 ADD CONSTRAINT 体检诊断建议_干保_PK PRIMARY KEY (结论id) USING INDEX PCTFREE 10;
ALTER TABLE 体检诊断建议_干保 ADD CONSTRAINT 体检诊断建议_干保_UQ_干保编码 UNIQUE (干保编码) USING INDEX PCTFREE 10;

--ALTER TABLE 病人信息_干保 ADD CONSTRAINT 病人信息_干保_FK_病人id FOREIGN KEY (病人id) REFERENCES 病人信息(病人id) ON DELETE CASCADE;

ALTER TABLE 体检登记记录_干保 ADD CONSTRAINT 体检登记记录_干保_FK_登记id FOREIGN KEY (登记id) REFERENCES 体检登记记录(ID) ON DELETE CASCADE;
ALTER TABLE 体检登记记录_干保 ADD CONSTRAINT 体检登记记录_干保_PK PRIMARY KEY (任务包号) USING INDEX PCTFREE 10;

ALTER TABLE 体检人员档案_干保 ADD CONSTRAINT 体检人员档案_干保_FK_登记id FOREIGN KEY (登记id) REFERENCES 体检登记记录(ID) ON DELETE CASCADE;
ALTER TABLE 体检人员档案_干保 ADD CONSTRAINT 体检人员档案_干保_FK_任务包号 FOREIGN KEY (任务包号) REFERENCES 体检登记记录_干保(任务包号) ON DELETE CASCADE;

--ALTER TABLE 体检项目清单_干保 ADD CONSTRAINT 体检项目清单_干保_FK_清单id FOREIGN KEY (清单id) REFERENCES 体检项目清单(ID) ON DELETE CASCADE;

CREATE INDEX 体检登记记录_干保_IX_登记id on 体检登记记录_干保(登记id) PCTFREE 10 TABLESPACE zl9CisRec;

CREATE INDEX 体检人员档案_干保_IX_任务包号 on 体检人员档案_干保(任务包号) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX 体检人员档案_干保_IX_登记id on 体检人员档案_干保(登记id) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX 体检人员档案_干保_IX_病人id on 体检人员档案_干保(病人id) PCTFREE 10 TABLESPACE zl9CisRec;


--权限处理

--模块
insert into zlPrograms(序号,标题,说明,系统,部件) values(1852,'体检项目管理','完成体检项目定义及包括的子项目关系。',100,'ZL9MEDICAL');
Insert Into zlProgFuncs(系统,序号,功能) Values(100,1852,'基本');

insert into zlPrograms(序号,标题,说明,系统,部件) values(6000,'干保体检接口平台','完成和接口体检软件的基础项目对码、接受接口软件的任务，并发回体检结果给接口软件。',100,NULL);

--基本
Insert Into zlProgFuncs(系统,序号,功能) Values(100,6000,'基本');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'诊疗分类目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'诊疗项目类别','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'诊疗项目别名','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'诊治所见分类','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'床位状况记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'诊疗执行科室','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'诊疗用法用量','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'检验项目参考','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'检验报告项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'诊疗检验标本','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'体检登记记录_ID','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'部门表','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'部门表_干保','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'体检诊断建议','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'体检诊断建议_干保','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'疾病编码目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'诊疗项目目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'诊疗项目目录_干保','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'诊治所见项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'诊治所见项目_干保','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'体检登记记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'体检登记记录_干保','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'病人信息','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'病人信息_干保','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'体检人员档案','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'体检人员档案_干保','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'体检项目清单','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'体检项目清单_干保','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'病人医嘱记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'病人医嘱发送','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'体检项目医嘱','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'病人病历记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'病人病历内容','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'病人病历所见单','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'体检人员结论','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'基本',user,'病人病历文本段','SELECT');

--数据对码
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,6000,'数据对码','完成体检软件基础项目同接口软件基础项目间的对应关系。');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'数据对码',user,'部门表_干保','Insert');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'数据对码',user,'部门表_干保','Delete');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'数据对码',user,'体检诊断建议_干保','Insert');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'数据对码',user,'体检诊断建议_干保','Delete');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'数据对码',user,'诊疗项目目录_干保','Insert');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'数据对码',user,'诊疗项目目录_干保','Delete');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'数据对码',user,'诊治所见项目_干保','Insert');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'数据对码',user,'诊治所见项目_干保','Delete');

--接受任务
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,6000,'接受任务','接受接口体检软件产生的体检任务（包括受检人员、体检项目等）。');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'接受任务',user,'ZL_体检登记记录_DELETE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'接受任务',user,'ZL_体检登记记录_STATE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'接受任务',user,'ZL_体检登记记录_INSERT','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'接受任务',user,'ZL_体检组别_INSERT','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'接受任务',user,'ZL_体检项目清单_INSERT','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'接受任务',user,'ZL_体检人员档案_INSERT','EXECUTE');

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'接受任务',user,'体检人员档案_干保','INSERT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'接受任务',user,'体检人员档案_干保','DELETE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'接受任务',user,'体检登记记录_干保','INSERT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'接受任务',user,'体检登记记录_干保','DELETE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'接受任务',user,'体检项目清单_干保','INSERT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'接受任务',user,'体检项目清单_干保','DELETE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'接受任务',user,'病人信息_干保','INSERT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'接受任务',user,'病人信息_干保','DELETE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'接受任务',user,'病人信息','UPDATE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'接受任务',user,'号码控制表','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'接受任务',user,'号码控制表','UPDATE');

--发送结果
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,6000,'发送结果','完成将体检结果发送给接口体检软件工作。');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,6000,'发送结果',user,'体检登记记录_干保','UPDATE');


Insert Into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) 
Select 组别,zlMenus_id.nextval,ID,'体检项目管理','体检项目','B',99,'。',100,1852 From zlMenus Where 标题='体检管理系统';

----病历元素目录
Insert Into 病历元素目录 (ID,编码,名称,说明,类型,转文本,字体,部件,适用,外编辑)
Select 病历元素目录_ID.nextval,'000067','体检检查报告','对体检人员已经进行的检查结果的记录报告',4,1,'宋体,9,0,0','ZL9CISCORE.USRMEDICALCHECK','00001',0 From dual;

----------------------------------------------------------------------------
---  INSERT   for   体检项目清单
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_体检项目清单_INSERT(
	登记id_IN IN 体检项目清单.登记id%TYPE,
	组别名称_IN IN 体检项目清单.组别名称%TYPE,
	诊疗项目id_IN IN 体检项目清单.诊疗项目id%TYPE,
	体检类型_IN IN 体检项目清单.体检类型%TYPE,
	基本价格_IN IN 体检项目清单.基本价格%TYPE,
	体检价格_IN IN 体检项目清单.体检价格%TYPE,
	执行科室id_IN IN 体检项目清单.执行科室id%TYPE:=NULL,
	采集方式id_IN IN 体检项目清单.采集方式id%TYPE:=NULL,
	采集科室id_IN IN 体检项目清单.采集科室id%TYPE:=NULL,
	检验标本_IN IN 体检项目清单.检验标本%TYPE:=NULL,
	检查部位_IN IN 体检项目清单.检查部位%TYPE:=NULL,
	检查部位id_IN IN 体检项目清单.检查部位id%TYPE:=NULL,
	病人id_IN IN 体检项目清单.病人id%TYPE:=0,
	结算途径_IN IN 体检项目清单.结算途径%TYPE:=1,
	计费明细_IN IN VARCHAR2:=Null,
	清单id_IN In Number:=0
)
IS
	
	v_清单id		NUMBER(18);
	strTmp			Varchar2(4000);
	str计费项		Varchar2(4000);
	intPos			Number(18);
	v_收费细目id		体检项目计价.收费细目id%TYPE;
	v_数次			体检项目计价.数次%TYPE;
	v_单价			体检项目计价.单价%TYPE;
	v_标准单价		体检项目计价.标准单价%TYPE;

	v_执行科室id		体检项目计价.执行科室id%TYPE;
	v_计价性质		体检项目计价.计价性质%TYPE;
BEGIN
	If 清单id_IN=0 Then
		SELECT 体检项目清单_ID.NEXTVAL INTO v_清单id FROM DUAL;
	Else
		v_清单id:=清单id_IN;
	End If;

	Insert Into 体检项目清单(ID,登记id,组别名称,诊疗项目id,执行科室id,采集方式id,采集科室id,检验标本,检查部位,检查部位id,病人id,体检类型,结算途径,基本价格,体检价格)
	VALUES(v_清单id,登记id_IN,组别名称_IN,诊疗项目id_IN,DECODE(执行科室id_IN,0,NULL,执行科室id_IN),DECODE(采集方式id_IN,0,NULL,采集方式id_IN),DECODE(采集科室id_IN,0,NULL,采集科室id_IN),检验标本_IN,检查部位_IN,检查部位id_IN,DECODE(病人id_IN,0,NULL,病人id_IN),体检类型_IN,结算途径_IN,基本价格_IN,体检价格_IN);
	
	Delete from 体检项目计价 where 清单id=v_清单id;
	
	strTmp := 计费明细_IN||';';

	WHILE strTmp IS NOT NULL LOOP
		
		intPos := INSTR (strTmp, ';');

		IF intPos >0 Then
			
			str计费项 := SUBSTR (strTmp, 1, intPos - 1)||':';
			strTmp := SUBSTR (strTmp, intPos + 1);
			
			If str计费项 Is Not Null Then
				intPos := INSTR (str计费项, ':');
				v_收费细目id := To_Number(SUBSTR (str计费项, 1, intPos - 1));				
				str计费项 := SUBSTR (str计费项, intPos + 1);

				intPos := INSTR (str计费项, ':');
				v_数次 := To_Number(SUBSTR (str计费项, 1, intPos - 1));
				str计费项 := SUBSTR (str计费项,intPos + 1);

				intPos := INSTR (str计费项, ':');
				v_标准单价 := To_Number(SUBSTR (str计费项, 1, intPos - 1));
				str计费项 := SUBSTR (str计费项, intPos + 1);

				intPos := INSTR (str计费项, ':');
				v_单价 := To_Number(SUBSTR (str计费项, 1, intPos - 1));
				str计费项 := SUBSTR (str计费项, intPos + 1);
				
				intPos := INSTR (str计费项, ':');
				v_执行科室id := To_Number(SUBSTR (str计费项, 1, intPos - 1));
				str计费项 := SUBSTR (str计费项, intPos + 1);

				intPos := INSTR (str计费项, ':');
				v_计价性质 := To_Number(SUBSTR (str计费项, 1, intPos - 1));
				str计费项 := SUBSTR (str计费项, intPos + 1);

				Insert Into 体检项目计价(清单id,收费细目id,数次,标准单价,单价,执行科室id,计价性质) values (v_清单id,v_收费细目id,v_数次,v_标准单价,v_单价,v_执行科室id,v_计价性质);

			End If;
		End If;       
	END LOOP;


	--向该组的人员添加此项目
	If 组别名称_IN Is Not Null Then

		Insert Into 体检项目医嘱(病人id,清单id,医嘱id)
		Select B.病人id,A.ID,NULL From 体检项目清单 A,体检人员档案 B Where A.登记id=B.登记id And B.登记id=登记id_IN And B.组别名称=A.组别名称 And A.ID=v_清单id;
	Else
		Insert Into 体检项目医嘱(病人id,清单id,医嘱id)
		Select B.病人id,A.ID,NULL From 体检项目清单 A,体检人员档案 B Where A.登记id=B.登记id And B.登记id=登记id_IN And A.病人id=B.病人id And A.ID=v_清单id And B.病人id=病人id_IN;
	End If;

EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_体检项目清单_INSERT;
/

----------------------------------------------------------------------------
---  DELETE   for   病历所见单
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_体检检查组合_DELETE(
	ID_IN IN 诊疗项目目录.ID%TYPE
)
IS
BEGIN
	DELETE FROM 病历所见单 WHERE 行=ID_IN And 元素id In (Select ID From 病历元素目录 Where 类型=-1);
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_体检检查组合_DELETE;
/

----------------------------------------------------------------------------
---  DELETE   for   体检项目排列
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_体检项目排列_DELETE
IS
BEGIN
	DELETE FROM 体检项目排列;
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_体检项目排列_DELETE;
/

----------------------------------------------------------------------------
---  INSERT   for   体检项目排列
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_体检项目排列_INSERT(
	诊疗项目id_IN	IN	体检项目排列.诊疗项目id%type,
	排列顺序_IN	IN	体检项目排列.排列顺序%type
)
IS
BEGIN
	INSERT INTO 体检项目排列(诊疗项目id,排列顺序) Values (诊疗项目id_IN,排列顺序_IN);
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_体检项目排列_INSERT;
/