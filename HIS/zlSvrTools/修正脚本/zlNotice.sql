
Create Table zlNotices(
	���		NUMBER(5) NOT NULL,
	ϵͳ		NUMBER(5),		
	��������	VARCHAR2(4000),
	��������	VARCHAR2(250),
	���ѱ���	VARCHAR2(50),
	��������	NUMBER(5),
	���Ѵ���	NUMBER(1),
	����˳��	VARCHAR2(200) DEFAULT '',
	�������	NUMBER(5),
	��������	NUMBER(5),	
	��ʼʱ��	DATE NOT NULL,
	��ֹʱ��	DATE)			
	PCTFREE 5
	PCTUSED 90
	STORAGE (INITIAL 512 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/

Alter Table zlNotices ADD CONSTRAINT 
	zlNotices_PK PRIMARY KEY (���)
	USING INDEX PCTFREE 5
	STORAGE(INITIAL 256 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/
Create Table zlNoticeUsr(
	�������	NUMBER(5) NOT NULL,
	���Ѷ���	NUMBER(1) DEFAULT 0,	--0-����;1-ָ����Ա;2-ָ������;3-ָ������վ
	��������	VARCHAR2(50))			
	PCTFREE 5
	PCTUSED 90
	STORAGE (INITIAL 512 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/
Create Table zlNoticeRec(
	�������	NUMBER(5) NOT NULL,	
	�û���		VARCHAR2(30),
	���ʱ��	DATE,
	�����	NUMBER(1) DEFAULT 0,	--1��ʾ��Ҫ���ѵ�����;0��ʾ��Ҫ���ѵ�����
	���ѱ�־	NUMBER(1) DEFAULT 0,	--1��ʾҪ����;0��ʾ������
	����ʱ��	DATE,
	��������	VARCHAR2(250))			
	PCTFREE 5
	PCTUSED 90
	STORAGE (INITIAL 512 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/

Alter Table zlNoticeUsr ADD CONSTRAINT 
	zlNoticeUsr_PK PRIMARY KEY (�������,���Ѷ���,��������)
	USING INDEX PCTFREE 5
	STORAGE(INITIAL 256 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/
Alter Table zlNoticeUsr ADD CONSTRAINT 
	zlNoticeUsr_FK_������� FOREIGN KEY(�������) 
	REFERENCES zlNotices(���) ON DELETE CASCADE
/
Alter Table zlNoticeRec ADD CONSTRAINT 
	zlNoticeRec_PK PRIMARY KEY (�������,�û���)
	USING INDEX PCTFREE 5
	STORAGE(INITIAL 256 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/
Alter Table zlNoticeRec ADD CONSTRAINT 
	zlNoticeRec_FK_������� FOREIGN KEY(�������) 
	REFERENCES zlNotices(���) ON DELETE CASCADE
/

--���ò˵�λ��
insert into zlSvrTools(���,�ϼ�,����,���,˵��) values ('0504','05','�Զ�����','H',Null)
/

----------------------------------------------------------------------------
---  INSERT   for   ZLNOTICES
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_ZLNOTICES_INSERT(
	���_IN IN ZLNOTICES.���%TYPE,
	ϵͳ_IN IN ZLNOTICES.ϵͳ%TYPE,
	��������_IN IN ZLNOTICES.��������%TYPE,
	��������_IN IN ZLNOTICES.��������%TYPE,
	���ѱ���_IN IN ZLNOTICES.���ѱ���%TYPE,
	��������_IN IN ZLNOTICES.��������%TYPE,
	���Ѵ���_IN IN ZLNOTICES.���Ѵ���%TYPE,
	�������_IN IN ZLNOTICES.�������%TYPE,
	��������_IN IN ZLNOTICES.��������%TYPE,
	��ʼʱ��_IN IN ZLNOTICES.��ʼʱ��%TYPE,
	��ֹʱ��_IN IN ZLNOTICES.��ֹʱ��%TYPE,
	����˳��_IN IN ZLNOTICES.����˳��%TYPE
)
IS
BEGIN
	Insert Into ZLNOTICES
		(���,ϵͳ,��������,��������,���ѱ���,��������,���Ѵ���,�������,��������,��ʼʱ��,��ֹʱ��,����˳��)
		VALUES
		(���_IN,ϵͳ_IN,��������_IN,��������_IN,���ѱ���_IN,��������_IN,���Ѵ���_IN,�������_IN,��������_IN,��ʼʱ��_IN,��ֹʱ��_IN,����˳��_IN);
END ZL_ZLNOTICES_INSERT;
/

----------------------------------------------------------------------------
---  UPDATE   for   ZLNOTICES
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_ZLNOTICES_UPDATE(
	���_IN IN ZLNOTICES.���%TYPE,
	ϵͳ_IN IN ZLNOTICES.ϵͳ%TYPE,
	��������_IN IN ZLNOTICES.��������%TYPE,
	��������_IN IN ZLNOTICES.��������%TYPE,
	���ѱ���_IN IN ZLNOTICES.���ѱ���%TYPE,
	��������_IN IN ZLNOTICES.��������%TYPE,
	���Ѵ���_IN IN ZLNOTICES.���Ѵ���%TYPE,
	�������_IN IN ZLNOTICES.�������%TYPE,
	��������_IN IN ZLNOTICES.��������%TYPE,
	��ʼʱ��_IN IN ZLNOTICES.��ʼʱ��%TYPE,
	��ֹʱ��_IN IN ZLNOTICES.��ֹʱ��%TYPE,
	����˳��_IN IN ZLNOTICES.����˳��%TYPE
)
IS
BEGIN
	Update ZLNOTICES
		Set ���=���_IN,
		    ϵͳ=ϵͳ_IN,
		    ��������=��������_IN,
		    ��������=��������_IN,
		    ���ѱ���=���ѱ���_IN,
		    ��������=��������_IN,
		    ���Ѵ���=���Ѵ���_IN,
		    �������=�������_IN,
		    ��������=��������_IN,
		    ��ʼʱ��=��ʼʱ��_IN,
		    ��ֹʱ��=��ֹʱ��_IN,
		    ����˳��=����˳��_IN
		Where  ���=���_IN;
END ZL_ZLNOTICES_UPDATE;
/

----------------------------------------------------------------------------
---  DELETE   for   ZLNOTICES
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_ZLNOTICES_DELETE(
	���_IN IN ZLNOTICES.���%TYPE
)
IS
BEGIN
	Delete From ZLNOTICEUSR Where  �������=���_IN;
	Delete From ZLNOTICES Where  ���=���_IN;
END ZL_ZLNOTICES_DELETE;
/

----------------------------------------------------------------------------
---  INSERT   for   ZLNOTICEUSR
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_ZLNOTICEUSR_INSERT(
	�������_IN IN ZLNOTICEUSR.�������%TYPE,
	���Ѷ���_IN IN ZLNOTICEUSR.���Ѷ���%TYPE,
	��������_IN IN ZLNOTICEUSR.��������%TYPE
)
IS
BEGIN
	Insert Into ZLNOTICEUSR
		(�������,���Ѷ���,��������)
		VALUES
		(�������_IN,���Ѷ���_IN,��������_IN);
END ZL_ZLNOTICEUSR_INSERT;
/
----------------------------------------------------------------------------
---  DELETE   for   ZLNOTICEUSR
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_ZLNOTICEUSR_DELETE(
	�������_IN IN ZLNOTICEUSR.�������%TYPE
)
IS
BEGIN
	Delete From ZLNOTICEUSR
		Where  �������=�������_IN;
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
---  ������
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_ZLNOTICEREC_CHECKNOTICE(
	�û���_IN	IN	zlNoticeUsr.��������%type,
	������_IN	IN	zlNoticeUsr.��������%type:='',
	����վ_IN	IN	zlNoticeUsr.��������%type:='',
	�������_IN	IN	NUMBER:=0)
IS
	Cursor c_Notices IS
		SELECT A.*,B.���ʱ�� FROM zlNotices A,
			(SELECT �������,���ʱ�� FROM zlNoticeRec WHERE �û���=�û���_IN) B 
		WHERE A.���=B.�������(+)
			AND ((������� IS NULL AND 1=�������_IN) OR (������� IS NOT NULL AND 0=�������_IN))
			AND A.��ʼʱ��<=SYSDATE AND (A.��ֹʱ��>=SYSDATE OR A.��ֹʱ�� IS NULL)
			AND (A.��� IN (SELECT ������� FROM zlNoticeUsr
					WHERE (���Ѷ��� = 1 AND �������� = �û���_IN) 
						OR (���Ѷ��� = 2 AND �������� = ������_IN) 
						OR (���Ѷ��� = 3 AND �������� = ����վ_IN))
			OR A.��� NOT IN (SELECT ������� FROM zlNoticeUsr));

	r_Notice c_Notices%RowType;
	
	v_����� number(1);
	v_�������� varchar2(500);
	
	v_���� number(1);
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
		
		v_����:=0;
				
		--ͨ������ϴμ��ʱ���Ƿ�Ϊ��������Ƿ�Ϊ��һ�μ��
		if r_Notice.���ʱ�� is null then		
			--��һ�μ��,������¼
			insert into zlNoticeRec(�������,�û���,���ʱ��,����ʱ��,��������) values (r_Notice.���,�û���_IN,SYSDATE,NULL,NULL);
			v_����:=1;
		else
			--�� 2��3��... �μ��
			--��ǰʱ���Ƿ�����ϴμ��ʱ�����һ���������,��������ˣ�����¼��ʱ��
			if r_Notice.������� is null then
				if �������_IN=1 then
					update zlNoticeRec set ���ʱ��=SYSDATE	where �������=r_Notice.��� and �û���=�û���_IN;
					v_����:=1;
				end if;
			else
				if SYSDATE>=(r_Notice.���ʱ��+r_Notice.�������/(24*60)) then
					update zlNoticeRec set ���ʱ��=SYSDATE	where �������=r_Notice.��� and �û���=�û���_IN;
					v_����:=1;
				end if;
			end if;
		end if;	
		
		if v_����=1 then
			v_�����:=0;
			v_��������:='';

			--�������		
			if not (r_Notice.�������� is null) then					
				v_��������:=r_Notice.��������;

				--strTmp��ʽ:��'[����];varchar2|[�Ա�];date'
				v_Tmp:=r_Notice.����˳��||'|';
				WHILE not (v_Tmp is null) LOOP

					v_Pos := instr(v_Tmp, '|');								
					v_TmpField:=substr(v_Tmp,1,v_Pos - 1);	
					
					v_FieldPos:=instr(v_TmpField,';');
					v_Field:=substr(v_TmpField,1,v_FieldPos - 1);
					v_FieldType:=trim(Upper(substr(v_TmpField,v_FieldPos+1,100)));

					v_Tmp:=trim(substr(v_Tmp,v_Pos + 1,1000));

					v_Pos:=instr(v_��������,v_Field);

					if v_Pos>0 then
						
						v_Result:=trim(substr(v_Field,2,1000));
						v_Result:=substr(v_Result,1,LENGTH(v_Result)-1);

						if v_FieldType='NUMBER' then
							v_Result:='to_char('||v_Result||')';
						Elsif v_FieldType='DATE' then
							v_Result:='to_char('||v_Result||',''yyyy-mm-dd'')';
						End if;

						v_��������:=trim(substr(v_��������,1,v_Pos - 1)||'''||'||v_Result||'||'''||substr(v_��������,v_Pos + length(v_Field),1000));

					end if;

				END LOOP;
				v_Pos:=instr(Upper(r_Notice.��������),' FROM ');
				
				if v_Pos>0 then
					v_SQL:=TRIM('SELECT '''||v_��������||''''||substr(r_Notice.��������,v_Pos,4000));

					v_CursorID:=sys.DBMS_SQL.OPEN_CURSOR;
					sys.DBMS_SQL.PARSE(v_CursorID,v_SQL,sys.DBMS_SQL.NATIVE);
					
					dbms_sql.define_column(v_CursorID,1,v_Result,1000);

					v_return :=DBMS_SQL.execute(v_CursorID); 
					
					if DBMS_SQL.FETCH_ROWS(v_CursorID)>0 then
						--�������µ��������

						v_�����:=1;	
						dbms_sql.column_value(v_CursorID,1,v_Result);
						v_��������:=trim(v_Result);						
					end if;
				end if;
				
			else
				v_�����:=1;
				v_��������:=r_Notice.��������;
			end if;
			
			update zlNoticeRec set �����=v_�����,��������=v_�������� where �������=r_Notice.��� and �û���=�û���_IN;
		end if;
	END Loop;

END ZL_ZLNOTICEREC_CHECKNOTICE;
/
----------------------------------------------------------------------------
---  ���Ѹ���
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_ZLNOTICEREC_NOTICE(
	�û���_IN	IN	zlNoticerec.�û���%type,
	�������_IN	IN	NUMBER:=0)
IS
	Cursor c_Notices IS
		SELECT B.*,A.��������,A.������� FROM zlNotices A,
			(SELECT * FROM zlNoticeRec WHERE �û���=�û���_IN) B 
		WHERE A.���=B.�������
			AND B.�����=1 
			AND ((�������_IN=1 AND ������� IS NULL) OR (������� IS NOT NULL AND �������_IN=0))
			AND A.��ʼʱ��<=SYSDATE AND (A.��ֹʱ��>=SYSDATE OR A.��ֹʱ�� IS NULL);

	r_Notice c_Notices%RowType;
  
	v_���ѷ� number(1);
  
BEGIN
	
	Update zlNoticeRec Set	���ѱ�־=0 where �û���=�û���_IN;
		
	FOR r_Notice In c_Notices Loop
		
		v_���ѷ�:=0;

		if r_Notice.����ʱ�� is null then
			
			--��һ������
			v_���ѷ�:=1;
		else
			--�� 2��3��... ������
			if r_Notice.������� is null then
				if �������_IN=1 then
					v_���ѷ�:=1;
				end if;
			else
				--��ǰʱ���Ƿ�����ϴ�����ʱ�����һ����������
				if SYSDATE>=(r_Notice.����ʱ��+r_Notice.��������/(24*60)) then
					v_���ѷ�:=1;
				end if;
			end if;
		end if;		
		
		if v_���ѷ�=1 then
			Update zlNoticeRec Set	����ʱ��=SYSDATE,
						���ѱ�־=v_���ѷ�
			where �������=r_Notice.������� 
				and �û���=�û���_IN;
		end if;

	END Loop;

END ZL_ZLNOTICEREC_NOTICE;
/

--����ͬ���
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
--Ȩ��
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