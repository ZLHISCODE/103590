-------------------------------------------------------------------------------
--��ṹ����
-------------------------------------------------------------------------------
-- Create table
create table PACS_TMP���˲�����¼
(
  ID       NUMBER(18) not null,
  ����ID   NUMBER(18) not null,
  ����ID   NUMBER(18),
  ����ID   NUMBER(18),
  �������� VARCHAR2(20),
  ��д��ID NUMBER(18),
  ��д��   VARCHAR2(50),
  ��д���� DATE,
  ������ID NUMBER(18),
  ������   VARCHAR2(50),
  �������� DATE,
  ��¼���� NUMBER
)
tablespace ZL9CISREC
  pctfree 15;
-- Add comments to the columns 
comment on column PACS_TMP���˲�����¼.��¼����
  is '1-��INSERT���룬�ش������ˣ�����ʱ�䣻2-��UPDATE���룬�ش������ˣ�����״̬Ϊ��ɣ�3-��"����"���룬�޸ı���״̬Ϊδ���';
-- Create/Recreate primary, unique and foreign key constraints 
alter table PACS_TMP���˲�����¼
  add constraint PACS_TMP���˲�����¼_PK primary key (ID)
  using index 
  tablespace ZL9CISREC
  pctfree 5;


-- ��������
create sequence PACS_TMP���˲�����¼_ID
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
  �����   NUMBER,
  �������� VARCHAR2(100),
  ����ʱ�� DATE
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
--�洢���̲���
-------------------------------------------------------------------------------
CREATE OR REPLACE Procedure Zlpacs_����
(
  ҽ��ID_IN       ����ҽ����¼.ID%TYPE,
  ��ʶ��_IN       ������Ϣ.�����%Type,
  ����_In         ������Ϣ.����%Type,
  �Ա�_In         ������Ϣ.�Ա�%Type,
  ����_In         ������Ϣ.����%Type,
  ��������_IN	  ������Ϣ.��������%TYPE,
  ����_In         ������Ϣ.����%TYPE,
  ����_In         ������Ϣ.����%TYPE,
  ����״��_In     ������Ϣ.����״��%TYPE,
  ְҵ_In         ������Ϣ.ְҵ%TYPE,
  ���֤��_In     ������Ϣ.���֤��%Type,
  ������λ_In     ������Ϣ.������λ%Type,
  ��λ�ʱ�_In     ������Ϣ.��λ�ʱ�%Type,
  ��ͥ��ַ_In     ������Ϣ.��ͥ��ַ%Type,
  ��ͥ�绰_In     ������Ϣ.��ͥ�绰%Type,
  �����Ŀ����_In ������ĿĿ¼.����%Type,
  �걾��λ_In     ����ҽ����¼.�걾��λ%Type,
  ��������id_In   ����ҽ����¼.��������id%Type,
  ����ҽ��_In     ����ҽ����¼.����ҽ��%Type,
  ����ʱ��_In     ����ҽ����¼.����ʱ��%Type,
  ������Դ_IN	  ����ҽ����¼.������Դ%TYPE,
  ���˿���ID      ����ҽ����¼.���˿���ID%TYPE,
  ����_In         ������ҳ.��Ժ����%TYPE,
  ��¼����_IN     ����ҽ������.��¼����%Type,
  �Ʒ�״̬_IN     ����ҽ������.�Ʒ�״̬%Type,
  �޸�_IN         Number:=0
) Is
  --������Դ_IN ��1-�������죻2-סԺ
  --��¼����_IN��1-�շѼ�¼��2-���ʼ�¼��
  --�Ʒ�״̬_IN��-1-����Ʒ�(ͨ����ִ�к�Ժ��ִ�еĶ�����Ʒ�);0-δ�Ʒ�;1-�ѼƷѡ�
  --��ʶ��_IN�����ݲ�����Դȷ�������ﲡ��������ţ�סԺ������סԺ��
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
  --�жϲ�����Ϣ�Ƿ��Ѿ����ڣ�����Ѿ����ڣ���ֻ�޸Ĳ�����Ϣ,���ùҺŻ�����Ժ�������������ҺŻ���Ժ
  --�޸ĳ�ͨ�������š���ȡ������ID��Ȼ���޸Ļ�����Ϣ
    if �޸�_IN=1 then 
       --�޸Ļ�����Ϣ
       IF ������Դ_IN = 1 THEN 
          select count(*) into N_RowCount from ������Ϣ a where a.�����=��ʶ��_IN;
          IF N_RowCount =1  THEN
              select a.����id into Npatientid from ������Ϣ a where a.�����=��ʶ��_IN;
              Zl_������Ϣ_Update(Npatientid,��ʶ��_IN,'','','', ����_In, 
      	          �Ա�_In, ����_In, ��������_In,'', ���֤��_In,'', ְҵ_In, 
                  ����_In, ����_In,'', ����״��_In, ��ͥ��ַ_In, ��ͥ�绰_In,
                  '','','','','',Null, ������λ_In, ��λ�ʱ�_In,'','','',Null,Null,0);
          END IF;
       ELSE 
          select count(*) into N_RowCount from ������Ϣ a where a.סԺ��=��ʶ��_IN;
          IF N_RowCount =1  THEN
              select a.����id into Npatientid from ������Ϣ a where a.סԺ��=��ʶ��_IN;
              Zl_������Ϣ_Update(Npatientid, '',��ʶ��_IN, '','', ����_In, 
      	          �Ա�_In, ����_In, ��������_In,'', ���֤��_In,'', ְҵ_In, 
                  ����_In, ����_In,'', ����״��_In, ��ͥ��ַ_In, ��ͥ�绰_In,
                  '','','','','',Null, ������λ_In, ��λ�ʱ�_In,'','','',Null,Null,0);
   		        update ������ҳ set ��Ժ����=����_In where ����id=Npatientid;
        	END IF;
       END IF;
       
       --�޸�ҽ����¼�����Ҽ����Ŀ
       BEGIN
    		    Select A.ID, A.����, B.ִ�п���id
    		   	Into Nclinicid, Scliniccont, Nexedeptid
    		  	From ������ĿĿ¼ A, ����ִ�п��� B
    		  	Where A.���� = �����Ŀ����_In And A.ID = B.������Ŀid And B.������Դ =������Դ_IN;
    		EXCEPTION 
    		    WHEN No_Data_Found THEN 
    		  	v_Error:='�����Ŀ�����޶�Ӧ��ִ�п���';
    		  	Raise Err_Custom;
    		END;
        
  		  --�޸�PACSҽ��
        BEGIN
            select ҽ��״̬ into n_ClinicState from ����ҽ����¼ where id = ҽ��ID_IN;
        EXCEPTION
            WHEN No_Data_Found THEN 
            v_Error:='δ�ҵ���Ӧ��ҽ����¼';
    		  	Raise Err_Custom;
        END;
        
        update ����ҽ����¼ set ҽ��״̬ = 1 where id = ҽ��ID_IN;
        ZL_����ҽ����¼_UPDATE(ҽ��ID_IN,Null, 1,1,1,Nclinicid, Null,Null, 1,Scliniccont || '(' || �걾��λ_In || ')',
                               '', �걾��λ_In,'һ����', Null,Null, '',Null, 0,Nexedeptid, 4, 0,
                               ����ʱ��_In, Null, ���˿���ID,��������id_In, ����ҽ��_In,����ʱ��_In);                                     
        update ����ҽ����¼ set ҽ��״̬ = n_ClinicState where id = ҽ��ID_IN;
    ELSE 
        N_Add:=1;
        IF ������Դ_IN = 1 THEN
            select count(*) into N_RowCount from ������Ϣ a where a.�����=��ʶ��_IN;
            IF N_RowCount =1  THEN
                select a.����id into Npatientid from ������Ϣ a where a.�����=��ʶ��_IN;
                N_Add:=3;
            ELSE
                Select ������ + 1 Into Npatientid From ������Ʊ� Where ��Ŀ��� = 1;
    	  	      Select Nvl(Max(����id), 0) + 1 Into Npatientid1 From ������Ϣ Where ����id >= Npatientid;
    	  	      If Npatientid1 > Npatientid Then
    	    	        Npatientid := Npatientid1;
    	  	      End If;
    	  	      Update ������Ʊ� Set ������ = Npatientid Where ��Ŀ��� = 1;          
            END IF;    
            Zl_�ҺŲ��˲���_Insert(N_Add, Npatientid, ��ʶ��_IN, '', '', ����_In, �Ա�_In, ����_In, 
                '', '', ����_In,����_In, ����״��_In, ְҵ_In, ���֤��_In,������λ_In, Null, 
                '', ��λ�ʱ�_In, ��ͥ��ַ_In, ��ͥ�绰_In, '', ����ʱ��_In, Null,Null, ��������_IN);
        ELSE
            select count(*) into N_RowCount from ������Ϣ a where a.סԺ��=��ʶ��_IN;
            IF N_RowCount =1  THEN
                select a.����id into Npatientid from ������Ϣ a where a.סԺ��=��ʶ��_IN;
                N_Add:=0;
            ELSE
                Select ������ + 1 Into Npatientid From ������Ʊ� Where ��Ŀ��� = 1;
    	  	      Select Nvl(Max(����id), 0) + 1 Into Npatientid1 From ������Ϣ Where ����id >= Npatientid;
    	  	      If Npatientid1 > Npatientid Then
    	    	        Npatientid := Npatientid1;
    	  	      End If;
    	  	      Update ������Ʊ� Set ������ = Npatientid Where ��Ŀ��� = 1;  
            END IF;
            Zl_��Ժ������ҳ_Insert(0,0, Npatientid,��ʶ��_IN,Null, ����_In, �Ա�_In, ����_In, 
              	'', ��������_IN, ����_In, ����_In, '', ����״��_In, ְҵ_In, '', ���֤��_In, 
               	'', ��ͥ��ַ_In, '', ��ͥ�绰_In, '', '', '', '', ������λ_In, Null, '', 
               	��λ�ʱ�_In, '', '', '', Null, Null, Null, Null, '', '', '', '', '', '',
               	Null,Null, '', '',Null,Null, '',Null,Null, '',Null, '', '',
                N_Add,'',Null,Null);
            update ������ҳ set ��Ժ����=����_In,��Ժ����=sysdate where ����id=Npatientid;    
        END IF;    
        
        Select ������ + 1 Into Nsendno From ������Ʊ� Where ��Ŀ��� = 10;
  		  Update ������Ʊ� Set ������ = Nsendno Where ��Ŀ��� = 10;
  		  Scheckno := Nextno(13);
        
    		BEGIN
    		    Select A.ID, A.����, B.ִ�п���id
    		   	Into Nclinicid, Scliniccont, Nexedeptid
    		  	From ������ĿĿ¼ A, ����ִ�п��� B
    		  	Where A.���� = �����Ŀ����_In And A.ID = B.������Ŀid And B.������Դ =������Դ_IN;
    		EXCEPTION 
    		    WHEN No_Data_Found THEN 
    		  	v_Error:='�����Ŀ�����޶�Ӧ��ִ�п���';
    		  	Raise Err_Custom;
    		END;
  		  --PACSҽ��
  	  
  	  	Zl_����ҽ����¼_Insert(ҽ��ID_IN, Null, 1, ������Դ_IN, Npatientid, 1, 0, 1, 1, 'D', Nclinicid, Null, Null, Null, 1,
  	                         Scliniccont || '(' || �걾��λ_In || ')', '', �걾��λ_In, 'һ����', Null, Null, '', Null, 0,
  	                         Nexedeptid, 4, 0, Sysdate + 1 / (24 * 3600), Null, ���˿���ID, ��������id_In, ����ҽ��_In,
  	                         ����ʱ��_In,1,ҽ��ID_IN);
  	  	    
  	  	Zl_����ҽ������_Insert(ҽ��ID_IN, Nsendno, ��¼����_IN, Scheckno, 1, 1, Null, Null, Sysdate + 1 / (24 * 3600), 0, Nexedeptid, �Ʒ�״̬_IN, 1);
    END IF;     
EXCEPTION
   WHEN Err_Custom THEN
    	Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
   WHEN OTHERS THEN
    	Zl_Errorcenter(SQLCODE, SQLERRM);
End Zlpacs_����;
/

----------------------------

CREATE OR REPLACE Procedure Zlpacs_��ʼ���
(
  ִ�м�_IN   ����ҽ������.ִ�м�%Type,
  ����_IN	  Ӱ�����¼.����%Type,
  ҽ��ID_IN	  Ӱ�����¼.ҽ��ID%Type,
  ��ʶ��_IN   ������Ϣ.�����%Type,
  Ӱ�����_IN Ӱ�����¼.Ӱ�����%Type,
  ����_IN     Ӱ�����¼.����%Type,
  Ӣ����_IN   Ӱ�����¼.Ӣ����%Type,
  �Ա�_IN     Ӱ�����¼.�Ա�%Type,
  ����_IN     Ӱ�����¼.����%Type,
  ��������_IN Ӱ�����¼.��������%Type,
  ���_IN     Ӱ�����¼.���%Type,
  ����_IN     Ӱ�����¼.����%Type,
  ����豸_IN Ӱ�����¼.����豸%Type,
  �绰_IN     Ӱ�����¼.��ϵ�绰%Type:=Null,
  ƥ�䷽ʽ_IN Number:=1,
  �޸�_IN     Number:=0
) Is
  --�޸�_IN: 0-��ʼ��飻1-�޸Ŀ�ʼ�����Ϣ
	--ƥ�䷽ʽ_IN��1-����ƥ�䣻2-����/סԺ��ƥ�䣻3-����ʶ��ҽ��ID��ƥ��
	--�ڲ�����
	
  N_��ʶ�� 		Number;
  V_���UID  	Ӱ�����¼.���UID%Type;
  N_RowCount  Number;
  Nsendno     Number;
  Err_Custom Exception;
  v_Error Varchar2(255);
BEGIN
     BEGIN	
          select D.���ͺ� into Nsendno from ����ҽ������ D where D.ҽ��ID = ҽ��ID_IN;
     EXCEPTION
          WHEN No_Data_Found THEN
          	  v_Error:='ҽ��ID����ȷ��δ�ҵ���Ӧҽ����';
          	  Raise Err_Custom;
     END;
     --��ʼӰ����
     ZL_Ӱ����_BEGIN(ִ�м�_IN,����_IN,ҽ��ID_IN,Nsendno,Ӱ�����_IN,����_IN,
                    Ӣ����_IN,�Ա�_IN, ����_IN, ��������_IN, ���_IN, ����_IN,1,1, 
                    ����豸_IN, �޸�_IN, �绰_IN);

  	 --������ǰ���еļ�� '��ͼ��ͼ���Զ�ƥ��
  	 --���Ҹ���ƥ�䷽ʽ������ͼ��ļ��UID
  	 IF ƥ�䷽ʽ_IN=1 THEN
  	     N_��ʶ��:= ����_IN;
  	 ELSE 
         IF ƥ�䷽ʽ_IN=2 THEN
		         N_��ʶ��:=��ʶ��_IN;
	       ELSE
		         N_��ʶ��:= ҽ��ID_IN;
	       END IF;
  	 END IF;
     
     select count(*) into N_RowCount from Ӱ����ʱ��¼ a  
          Where a.����= N_��ʶ�� And a.Ӱ�����=Ӱ�����_IN;
     IF n_Rowcount =1 THEN 
          Select A.���UID into V_���UID From Ӱ����ʱ��¼ a  
               Where a.����= N_��ʶ�� And a.Ӱ�����=Ӱ�����_IN;
  	      ZL_Ӱ����_SET(ҽ��ID_IN, Nsendno, V_���UID);
     END IF;
     EXCEPTION
         WHEN Err_Custom THEN
    	       Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
         WHEN OTHERS THEN
    	       Zl_Errorcenter(SQLCODE, SQLERRM);
END Zlpacs_��ʼ���;


/
----------------------------
create or replace procedure ZLPACS_ȡ������
(
  ҽ��ID_IN	  Ӱ�����¼.ҽ��ID%Type
) is
  N_RowCount      Number;
  N_ExecState     Number;
  N_ExecProcess   Number;
  Err_Custom      Exception;
  v_Error         Varchar2(255);
begin
    --ֻ�з�������������������Ա�ȡ��
    --1.���ڽ��еļ��(ҽ��ִ��״̬=3��ִ�й���=2)
    --2.û�й���ͼ��Ӱ����UID.���UIDΪ�գ�  
    BEGIN
        select ִ��״̬,ִ�й��� into N_ExecState, N_ExecProcess 
            from ����ҽ������ where ҽ��ID = ҽ��ID_IN;
    EXCEPTION
        WHEN No_Data_Found THEN 
    		  	v_Error:='û�з�����������ȡ����ҽ����¼';
    		  	Raise Err_Custom;
    END;
    IF (N_ExecState =3 AND N_ExecProcess = 2) THEN 
        select Count(*) into N_RowCount from Ӱ�����¼ 
            where ҽ��ID = ҽ��ID_IN and ���UID is null;
        IF N_RowCount=1 THEN
            update ����ҽ������ set ִ��״̬ = 2 where ҽ��ID = ҽ��ID_IN;
        ELSE
            v_Error:='����Ѿ�����ͼ���޷�ȡ��������ȡ��ҽ��������ͼ��';
    		    Raise Err_Custom;
        END IF;
    ELSE
        v_Error:='����Ѿ���ɻ��߻�û�п�ʼ���޷�ȡ�������Ƚ�ҽ�����˵����ڽ���״̬';
    		Raise Err_Custom;
    END IF;
    EXCEPTION
        WHEN Err_Custom THEN
   	       Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
        WHEN OTHERS THEN
   	       Zl_Errorcenter(SQLCODE, SQLERRM);
end ZLPACS_ȡ������;


/
----------------------------
create or replace procedure ZLPACS_�ָ�����
(
    ҽ��ID_IN	  Ӱ�����¼.ҽ��ID%Type
)  is
    N_RowCount      Number;
    N_ExecState     Number;
    N_ExecProcess   Number;
    Err_Custom      Exception;
    v_Error         Varchar2(255);
begin
    --ֻ�з�������������������Ա��ָ�
    --1.ԭ�����ڽ��У����Ǳ�ȡ��(�ܾ�)�ļ��(ҽ��ִ��״̬=2��ִ�й���=2)
    --2.���������Ӱ�����¼
    BEGIN
        select ִ��״̬,ִ�й��� into N_ExecState, N_ExecProcess 
            from ����ҽ������ where ҽ��ID = ҽ��ID_IN;
    EXCEPTION
        WHEN No_Data_Found THEN 
    		  	v_Error:='û�з����������Իָ���ҽ����¼';
    		  	Raise Err_Custom;
    END;
    IF N_ExecState =2 AND N_ExecProcess = 2 THEN 
        select Count(*) into N_RowCount from Ӱ�����¼ 
            where ҽ��ID = ҽ��ID_IN;
        IF N_RowCount=1 THEN
            update ����ҽ������ set ִ��״̬ = 3 where ҽ��ID = ҽ��ID_IN;
        ELSE
            v_Error:='û���ҵ���Ӧ��Ӱ�����¼���޷��ָ�';
    		    Raise Err_Custom;
        END IF;
    ELSE
        v_Error:='ҽ��ִ�й��̺�ִ��״̬����ȷ���޷��ָ�';
    		Raise Err_Custom;
    END IF;
    EXCEPTION
        WHEN Err_Custom THEN
   	       Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
        WHEN OTHERS THEN
   	       Zl_Errorcenter(SQLCODE, SQLERRM);
end ZLPACS_�ָ�����;
/
----------------------------

-------------------------------------------------------------------------------
--����������
-------------------------------------------------------------------------------
create or replace trigger TBI_ZLPACS_����ҽ������_UPDATE
  after update on ����ҽ������
  for each row
declare
      -- local variables here
      N_RowCount Number;
      N_WriteDoctorNo ��Ա��.���%Type;
      N_CheckDoctorNo ��Ա��.���%Type;
      N_WriteDoctor   ���Ӳ�����¼.������%Type;
      N_CheckDortor   ���Ӳ�����¼.������%Type;
      N_WriteTime     ���Ӳ�����¼.����ʱ��%Type;
      N_SignClass     ���Ӳ�����¼.ǩ������%Type;
      N_ID Number;
begin
      --�ж��Ƿ�ִ�й����޸ĳ�4��������д��5-������ˣ�6-�������
      IF :NEW.ִ�й��� =4 OR :NEW.ִ�й��� =5 OR :NEW.ִ�й��� =6  THEN
      	   Select c.������ As ��д��,c.������ As �����,c.����ʱ�� As ����ʱ��,c.ǩ������
	   			 into N_WriteDoctor,N_CheckDortor,N_WriteTime,N_SignClass
	   			 From ����ҽ����¼ a ,����ҽ������ b,���Ӳ�����¼ c
	   			 Where a.Id=b.ҽ��ID And b.����Id =c.Id And a.id=:NEW.ҽ��ID
	   			 order by c.���汾 Desc;
	   			 if N_RowCount = 1 then
      	   	--�б��棬�Ų��Һͼ�¼��������Ϣ
      	   	--������¼ID
      			 			Select PACS_TMP���˲�����¼_ID.Nextval Into N_ID From Dual;

	   							--������дҽ�����
	      					select count(*) into N_RowCount from ��Ա�� A where a.����=N_WriteDoctor;
	      					if N_RowCount =1 then
	         					 select ��� into N_WriteDoctorNo from ��Ա�� A where a.����=N_WriteDoctor;
        	      	else
        	         	N_WriteDoctorNo:='9999';
        	      	end if;

									if N_SignClass >=2 then
      	      		--��������ҽ�����
      		      	select count(*) into N_RowCount from ��Ա�� A where a.����=N_CheckDortor;
      		      	if N_RowCount =1 then
      		         	select ��� into N_CheckDoctorNo from ��Ա�� A where a.����=N_CheckDortor;
      		      	else
      		         	N_CheckDoctorNo:='9999';
      		      	end if;
      		      	--������ʱ���˲�����¼��
      	      		insert into PACS_tmp���˲�����¼(id,����id,����ID,��д��ID,��д��,
      	             		��д����,������ID,������,��¼����)
      	             		values(N_ID,:NEW.ҽ��id,:NEW.ִ�в���ID,N_WriteDoctorNo,
      	             		N_WriteDoctor,N_WriteTime,N_CheckDoctorNo,N_CheckDortor,2);
            	   	else
            	   		--������ʱ���˲�����¼��
            	      		insert into PACS_tmp���˲�����¼(id,����id,����ID,��д��ID,��д��,
            	             		��д����,��¼����)
            	             		values(N_ID,:NEW.ҽ��id,:NEW.ִ�в���ID,N_WriteDoctorNo,
            	             		N_WriteDoctor,N_WriteTime,1);
            	   	End If;
	   					end if;
      END IF;
exception
       when others then
            null;
end TBI_ZLPACS_����ҽ������_UPDATE;

/

-------------------------------------------------------------------------------
--Ȩ�޲���
-------------------------------------------------------------------------------

Commit;