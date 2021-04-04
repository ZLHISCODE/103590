
Create Table zlNotices(
	序号		NUMBER(5) NOT NULL,
	系统		NUMBER(5),		
	提醒条件	VARCHAR2(4000),
	提醒内容	VARCHAR2(250),
	提醒报表	VARCHAR2(50),
	提醒声音	NUMBER(5),
	提醒窗口	NUMBER(1),
	提醒顺序	VARCHAR2(200) DEFAULT '',
	检查周期	NUMBER(5),
	提醒周期	NUMBER(5),	
	开始时间	DATE NOT NULL,
	终止时间	DATE)			
	PCTFREE 5
	PCTUSED 90
	STORAGE (INITIAL 512 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/

Alter Table zlNotices ADD CONSTRAINT 
	zlNotices_PK PRIMARY KEY (序号)
	USING INDEX PCTFREE 5
	STORAGE(INITIAL 256 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/
Create Table zlNoticeUsr(
	提醒序号	NUMBER(5) NOT NULL,
	提醒对象	NUMBER(1) DEFAULT 0,	--0-所有;1-指定人员;2-指定部门;3-指定工作站
	对象名称	VARCHAR2(50))			
	PCTFREE 5
	PCTUSED 90
	STORAGE (INITIAL 512 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/
Create Table zlNoticeRec(
	提醒序号	NUMBER(5) NOT NULL,	
	用户名		VARCHAR2(30),
	检查时间	DATE,
	检查结果	NUMBER(1) DEFAULT 0,	--1表示有要提醒的内容;0表示无要提醒的内容
	提醒标志	NUMBER(1) DEFAULT 0,	--1表示要提醒;0表示不提醒
	提醒时间	DATE,
	提醒内容	VARCHAR2(250))			
	PCTFREE 5
	PCTUSED 90
	STORAGE (INITIAL 512 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/

Alter Table zlNoticeUsr ADD CONSTRAINT 
	zlNoticeUsr_PK PRIMARY KEY (提醒序号,提醒对象,对象名称)
	USING INDEX PCTFREE 5
	STORAGE(INITIAL 256 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/
Alter Table zlNoticeUsr ADD CONSTRAINT 
	zlNoticeUsr_FK_提醒序号 FOREIGN KEY(提醒序号) 
	REFERENCES zlNotices(序号) ON DELETE CASCADE
/
Alter Table zlNoticeRec ADD CONSTRAINT 
	zlNoticeRec_PK PRIMARY KEY (提醒序号,用户名)
	USING INDEX PCTFREE 5
	STORAGE(INITIAL 256 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/
Alter Table zlNoticeRec ADD CONSTRAINT 
	zlNoticeRec_FK_提醒序号 FOREIGN KEY(提醒序号) 
	REFERENCES zlNotices(序号) ON DELETE CASCADE
/

--设置菜单位置
insert into zlSvrTools(编号,上级,标题,快键,说明) values ('0504','05','自动提醒','H',Null)
/

----------------------------------------------------------------------------
---  INSERT   for   ZLNOTICES
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_ZLNOTICES_INSERT(
	序号_IN IN ZLNOTICES.序号%TYPE,
	系统_IN IN ZLNOTICES.系统%TYPE,
	提醒条件_IN IN ZLNOTICES.提醒条件%TYPE,
	提醒内容_IN IN ZLNOTICES.提醒内容%TYPE,
	提醒报表_IN IN ZLNOTICES.提醒报表%TYPE,
	提醒声音_IN IN ZLNOTICES.提醒声音%TYPE,
	提醒窗口_IN IN ZLNOTICES.提醒窗口%TYPE,
	检查周期_IN IN ZLNOTICES.检查周期%TYPE,
	提醒周期_IN IN ZLNOTICES.提醒周期%TYPE,
	开始时间_IN IN ZLNOTICES.开始时间%TYPE,
	终止时间_IN IN ZLNOTICES.终止时间%TYPE,
	提醒顺序_IN IN ZLNOTICES.提醒顺序%TYPE
)
IS
BEGIN
	Insert Into ZLNOTICES
		(序号,系统,提醒条件,提醒内容,提醒报表,提醒声音,提醒窗口,检查周期,提醒周期,开始时间,终止时间,提醒顺序)
		VALUES
		(序号_IN,系统_IN,提醒条件_IN,提醒内容_IN,提醒报表_IN,提醒声音_IN,提醒窗口_IN,检查周期_IN,提醒周期_IN,开始时间_IN,终止时间_IN,提醒顺序_IN);
END ZL_ZLNOTICES_INSERT;
/

----------------------------------------------------------------------------
---  UPDATE   for   ZLNOTICES
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_ZLNOTICES_UPDATE(
	序号_IN IN ZLNOTICES.序号%TYPE,
	系统_IN IN ZLNOTICES.系统%TYPE,
	提醒条件_IN IN ZLNOTICES.提醒条件%TYPE,
	提醒内容_IN IN ZLNOTICES.提醒内容%TYPE,
	提醒报表_IN IN ZLNOTICES.提醒报表%TYPE,
	提醒声音_IN IN ZLNOTICES.提醒声音%TYPE,
	提醒窗口_IN IN ZLNOTICES.提醒窗口%TYPE,
	检查周期_IN IN ZLNOTICES.检查周期%TYPE,
	提醒周期_IN IN ZLNOTICES.提醒周期%TYPE,
	开始时间_IN IN ZLNOTICES.开始时间%TYPE,
	终止时间_IN IN ZLNOTICES.终止时间%TYPE,
	提醒顺序_IN IN ZLNOTICES.提醒顺序%TYPE
)
IS
BEGIN
	Update ZLNOTICES
		Set 序号=序号_IN,
		    系统=系统_IN,
		    提醒条件=提醒条件_IN,
		    提醒内容=提醒内容_IN,
		    提醒报表=提醒报表_IN,
		    提醒声音=提醒声音_IN,
		    提醒窗口=提醒窗口_IN,
		    检查周期=检查周期_IN,
		    提醒周期=提醒周期_IN,
		    开始时间=开始时间_IN,
		    终止时间=终止时间_IN,
		    提醒顺序=提醒顺序_IN
		Where  序号=序号_IN;
END ZL_ZLNOTICES_UPDATE;
/

----------------------------------------------------------------------------
---  DELETE   for   ZLNOTICES
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_ZLNOTICES_DELETE(
	序号_IN IN ZLNOTICES.序号%TYPE
)
IS
BEGIN
	Delete From ZLNOTICEUSR Where  提醒序号=序号_IN;
	Delete From ZLNOTICES Where  序号=序号_IN;
END ZL_ZLNOTICES_DELETE;
/

----------------------------------------------------------------------------
---  INSERT   for   ZLNOTICEUSR
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_ZLNOTICEUSR_INSERT(
	提醒序号_IN IN ZLNOTICEUSR.提醒序号%TYPE,
	提醒对象_IN IN ZLNOTICEUSR.提醒对象%TYPE,
	对象名称_IN IN ZLNOTICEUSR.对象名称%TYPE
)
IS
BEGIN
	Insert Into ZLNOTICEUSR
		(提醒序号,提醒对象,对象名称)
		VALUES
		(提醒序号_IN,提醒对象_IN,对象名称_IN);
END ZL_ZLNOTICEUSR_INSERT;
/
----------------------------------------------------------------------------
---  DELETE   for   ZLNOTICEUSR
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_ZLNOTICEUSR_DELETE(
	提醒序号_IN IN ZLNOTICEUSR.提醒序号%TYPE
)
IS
BEGIN
	Delete From ZLNOTICEUSR
		Where  提醒序号=提醒序号_IN;
END ZL_ZLNOTICEUSR_DELETE;
/

CREATE OR REPLACE PROCEDURE ZLTEST
IS
	v_test varchar2(10);
BEGIN
	SELECT sysdate INTO v_test from dual;

END;
/
----------------------------------------------------------------------------
---  检查更新
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_ZLNOTICEREC_CHECKNOTICE(
	用户名_IN	IN	zlNoticeUsr.对象名称%type,
	部门名_IN	IN	zlNoticeUsr.对象名称%type:='',
	工作站_IN	IN	zlNoticeUsr.对象名称%type:='',
	启动检查_IN	IN	NUMBER:=0)
IS
	Cursor c_Notices IS
		SELECT A.*,B.检查时间 FROM zlNotices A,
			(SELECT 提醒序号,检查时间 FROM zlNoticeRec WHERE 用户名=用户名_IN) B 
		WHERE A.序号=B.提醒序号(+)
			AND ((检查周期 IS NULL AND 1=启动检查_IN) OR (检查周期 IS NOT NULL AND 0=启动检查_IN))
			AND A.开始时间<=SYSDATE AND (A.终止时间>=SYSDATE OR A.终止时间 IS NULL)
			AND (A.序号 IN (SELECT 提醒序号 FROM zlNoticeUsr
					WHERE (提醒对象 = 1 AND 对象名称 = 用户名_IN) 
						OR (提醒对象 = 2 AND 对象名称 = 部门名_IN) 
						OR (提醒对象 = 3 AND 对象名称 = 工作站_IN))
			OR A.序号 NOT IN (SELECT 提醒序号 FROM zlNoticeUsr));

	r_Notice c_Notices%RowType;
	
	v_检查结果 number(1);
	v_提醒内容 varchar2(500);
	
	v_检查否 number(1);
	v_CursorID INTEGER;
	v_return INTEGER;
  
	v_Result varchar2(250);
	v_SQL varchar2(4000);

	v_Tmp varchar2(1000);
	v_TmpField varchar2(100);
	v_Pos number(18);
	v_FieldPos number(18);  
	v_FieldType varchar2(50);
	v_Field varchar2(50);
BEGIN
	
	FOR r_Notice In c_Notices Loop
		
		v_检查否:=0;
				
		--通过检查上次检查时间是否为空来检测是否为第一次检查
		if r_Notice.检查时间 is null then		
			--第一次检查,新增记录
			insert into zlNoticeRec(提醒序号,用户名,检查时间,提醒时间,提醒内容) values (r_Notice.序号,用户名_IN,SYSDATE,NULL,NULL);
			v_检查否:=1;
		else
			--第 2、3、... 次检查
			--当前时间是否大于上次检查时间加上一个检查周期,如果大于了，则更新检查时间
			if r_Notice.检查周期 is null then
				if 启动检查_IN=1 then
					update zlNoticeRec set 检查时间=SYSDATE	where 提醒序号=r_Notice.序号 and 用户名=用户名_IN;
					v_检查否:=1;
				end if;
			else
				if SYSDATE>=(r_Notice.检查时间+r_Notice.检查周期/(24*60)) then
					update zlNoticeRec set 检查时间=SYSDATE	where 提醒序号=r_Notice.序号 and 用户名=用户名_IN;
					v_检查否:=1;
				end if;
			end if;
		end if;	
		
		if v_检查否=1 then
			v_检查结果:=0;
			v_提醒内容:='';

			--检查提醒		
			if not (r_Notice.提醒条件 is null) then					
				v_提醒内容:=r_Notice.提醒内容;

				--strTmp格式:如'[姓名];varchar2|[性别];date'
				v_Tmp:=r_Notice.提醒顺序||'|';
				WHILE not (v_Tmp is null) LOOP

					v_Pos := instr(v_Tmp, '|');								
					v_TmpField:=substr(v_Tmp,1,v_Pos - 1);	
					
					v_FieldPos:=instr(v_TmpField,';');
					v_Field:=substr(v_TmpField,1,v_FieldPos - 1);
					v_FieldType:=trim(Upper(substr(v_TmpField,v_FieldPos+1,100)));

					v_Tmp:=trim(substr(v_Tmp,v_Pos + 1,1000));

					v_Pos:=instr(v_提醒内容,v_Field);

					if v_Pos>0 then
						
						v_Result:=trim(substr(v_Field,2,1000));
						v_Result:=substr(v_Result,1,LENGTH(v_Result)-1);

						if v_FieldType='NUMBER' then
							v_Result:='to_char('||v_Result||')';
						Elsif v_FieldType='DATE' then
							v_Result:='to_char('||v_Result||',''yyyy-mm-dd'')';
						End if;

						v_提醒内容:=trim(substr(v_提醒内容,1,v_Pos - 1)||'''||'||v_Result||'||'''||substr(v_提醒内容,v_Pos + length(v_Field),1000));

					end if;

				END LOOP;
				v_Pos:=instr(Upper(r_Notice.提醒条件),' FROM ');
				
				if v_Pos>0 then
					v_SQL:=TRIM('SELECT '''||v_提醒内容||''''||substr(r_Notice.提醒条件,v_Pos,4000));

					v_CursorID:=sys.DBMS_SQL.OPEN_CURSOR;
					sys.DBMS_SQL.PARSE(v_CursorID,v_SQL,sys.DBMS_SQL.NATIVE);
					
					dbms_sql.define_column(v_CursorID,1,v_Result,1000);

					v_return :=DBMS_SQL.execute(v_CursorID); 
					
					if DBMS_SQL.FETCH_ROWS(v_CursorID)>0 then
						--检查后有新的情况发生

						v_检查结果:=1;	
						dbms_sql.column_value(v_CursorID,1,v_Result);
						v_提醒内容:=trim(v_Result);						
					end if;
				end if;
				
			else
				v_检查结果:=1;
				v_提醒内容:=r_Notice.提醒内容;
			end if;
			
			update zlNoticeRec set 检查结果=v_检查结果,提醒内容=v_提醒内容 where 提醒序号=r_Notice.序号 and 用户名=用户名_IN;
		end if;
	END Loop;

END ZL_ZLNOTICEREC_CHECKNOTICE;
/
----------------------------------------------------------------------------
---  提醒更新
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_ZLNOTICEREC_NOTICE(
	用户名_IN	IN	zlNoticerec.用户名%type,
	启动检查_IN	IN	NUMBER:=0)
IS
	Cursor c_Notices IS
		SELECT B.*,A.提醒周期,A.检查周期 FROM zlNotices A,
			(SELECT * FROM zlNoticeRec WHERE 用户名=用户名_IN) B 
		WHERE A.序号=B.提醒序号
			AND B.检查结果=1 
			AND ((启动检查_IN=1 AND 检查周期 IS NULL) OR (检查周期 IS NOT NULL AND 启动检查_IN=0))
			AND A.开始时间<=SYSDATE AND (A.终止时间>=SYSDATE OR A.终止时间 IS NULL);

	r_Notice c_Notices%RowType;
  
	v_提醒否 number(1);
  
BEGIN
	
	Update zlNoticeRec Set	提醒标志=0 where 用户名=用户名_IN;
		
	FOR r_Notice In c_Notices Loop
		
		v_提醒否:=0;

		if r_Notice.提醒时间 is null then
			
			--第一次提醒
			v_提醒否:=1;
		else
			--第 2、3、... 次提醒
			if r_Notice.检查周期 is null then
				if 启动检查_IN=1 then
					v_提醒否:=1;
				end if;
			else
				--当前时间是否大于上次提醒时间加上一个提醒周期
				if SYSDATE>=(r_Notice.提醒时间+r_Notice.提醒周期/(24*60)) then
					v_提醒否:=1;
				end if;
			end if;
		end if;		
		
		if v_提醒否=1 then
			Update zlNoticeRec Set	提醒时间=SYSDATE,
						提醒标志=v_提醒否
			where 提醒序号=r_Notice.提醒序号 
				and 用户名=用户名_IN;
		end if;

	END Loop;

END ZL_ZLNOTICEREC_NOTICE;
/

--公共同义词
Create Public Synonym zlNotices for zlNotices
/
Create Public Synonym zlNoticeUsr for zlNoticeUsr
/
Create Public Synonym zlNoticeRec for zlNoticeRec
/
Create Public Synonym ZL_ZLNOTICES_INSERT for ZL_ZLNOTICES_INSERT
/
Create Public Synonym ZL_ZLNOTICES_UPDATE for ZL_ZLNOTICES_UPDATE
/
Create Public Synonym ZL_ZLNOTICES_DELETE for ZL_ZLNOTICES_DELETE
/
Create Public Synonym ZL_ZLNOTICEUSR_INSERT for ZL_ZLNOTICEUSR_INSERT
/
Create Public Synonym ZL_ZLNOTICEUSR_DELETE for ZL_ZLNOTICEUSR_DELETE
/
Create Public Synonym ZL_ZLNOTICEREC_CHECKNOTICE for ZL_ZLNOTICEREC_CHECKNOTICE
/
Create Public Synonym ZL_ZLNOTICEREC_NOTICE for ZL_ZLNOTICEREC_NOTICE
/
--权限
Grant select on zlNotices to PUBLIC
/
Grant select on zlNoticeUsr to PUBLIC
/
Grant select on zlNoticeRec to PUBLIC
/
Grant execute on ZL_ZLNOTICES_INSERT to PUBLIC
/
Grant execute on ZL_ZLNOTICES_UPDATE to PUBLIC
/
Grant execute on ZL_ZLNOTICES_DELETE to PUBLIC
/
Grant execute on ZL_ZLNOTICEUSR_INSERT to PUBLIC
/
Grant execute on ZL_ZLNOTICEUSR_DELETE to PUBLIC
/
Grant execute on ZL_ZLNOTICEREC_CHECKNOTICE to PUBLIC
/
Grant execute on ZL_ZLNOTICEREC_NOTICE to PUBLIC
/