--���ݽṹ
Create Table ҩƷ�豸�ӿ�(
  ID Number(18) Not Null, 
  ��� Varchar2(10) Not Null, 
  ���� Varchar2(20), 
  ���� Number(2), 
  �������� Date,
  ͣ������ Date, 
  ������Ϣ Varchar2(2000), 
  ��չ��Ϣ Xmltype, 
  ��ע Varchar2(200)
)
Pctfree 10 Initrans 1 
Tablespace Zl9medlst;

Create Sequence ҩƷ�豸�ӿ�_Id Start With 1;

Alter Table ҩƷ�豸�ӿ� Add Constraint ҩƷ�豸�ӿ�_Pk Primary Key(ID) Using Index Tablespace Zl9indexhis;
Alter Table ҩƷ�豸�ӿ� Add Constraint ҩƷ�豸�ӿ�_Uq_��� Unique(���) Using Index Tablespace Zl9indexhis;
Alter Table ҩƷ�豸�ӿ� Add Constraint ҩƷ�豸�ӿ�_Uq_���� Unique(����) Using Index Tablespace Zl9indexhis;

Create Table ҩƷ�շ������־(
  ������ Varchar2(8), 
  ���� Number(2),
  �ⷿID Number(18),
  ҵ����� Number(2), 
  ��־ Number(2),
  ��ת�� Number(3)
) Pctfree 10 Initrans 20
Tablespace Zl9medlst;

Alter Table ҩƷ�շ������־ Add Constraint ҩƷ�շ������־_Pk Primary Key(������, ����, �ⷿID) Using Index Tablespace Zl9indexhis;
Create Index ҩƷ�շ������־_IX_��ת�� ON ҩƷ�շ������־(��ת��) Tablespace Zl9indexhis;

Create Table ҩƷ�շ�סԺ��־(
  �շ�ID NUMBER(18), 
  ҵ����� Number(2), 
  ��־ Number(2),
  ��ת�� Number(3)
) Pctfree 10 Initrans 20
Tablespace Zl9medlst;

Alter Table ҩƷ�շ�סԺ��־ Add Constraint ҩƷ�շ�סԺ��־_Pk Primary Key(�շ�ID, ҵ�����) Using Index Tablespace Zl9indexhis;
Alter Table ҩƷ�շ�סԺ��־ Add Constraint ҩƷ�շ�סԺ��־_FK_�շ�ID Foreign Key(�շ�ID) References ҩƷ�շ���¼(ID);
Create Index ҩƷ�շ�סԺ��־_IX_��ת�� ON ҩƷ�շ�סԺ��־(��ת��) Tablespace Zl9indexhis;

--Ȩ�޿���
Insert Into zlPrograms(���,����,˵��,ϵͳ,����) Values(9010,'ҩƷ�Զ����豸�ӿ�','ҩƷ�Զ����豸�ӿڵ�����ģ��',&n_System,'zlDrugMachine');

Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��,ȱʡֵ)
Select &n_System,9010,A.* From (
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0 Union All
  Select '����',-NULL,NULL,1 From Dual Union All
  Select '��������',1,'���в����趨',0 From Dual Union All
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,9010,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All
  Select '���ű�','SELECT' From Dual Union All
  Select 'ҩƷ�豸�ӿ�','SELECT' From Dual Union All
  Select 'ҩƷ����','SELECT' From Dual Union All
  Select '��������˵��','SELECT' From Dual Union All
  Select '�������ʷ���','SELECT' From Dual Union All
  Select '��Ա���ʷ���','SELECT' From Dual Union All
  Select '��Ա����˵��','SELECT' From Dual Union All
  Select '��Ա��','SELECT' From Dual Union All
  Select '�ϻ���Ա��','SELECT' From Dual Union All
  Select '������Ա','SELECT' From Dual Union All
  Select 'ҩƷ�շ���¼','SELECT' From Dual Union All
  Select 'ҩƷ�շ������־','SELECT' From Dual Union All
  Select 'ҩƷ�շ�סԺ��־','SELECT' From Dual Union All
  Select '�շ���ĿĿ¼','SELECT' From Dual Union All
  Select '�շ���Ŀ����','SELECT' From Dual Union All
  Select 'ҩƷ���','SELECT' From Dual Union All
  Select 'ҩƷ����','SELECT' From Dual Union All
  Select '������ĿĿ¼','SELECT' From Dual Union All
  Select 'ҩƷ������','SELECT' From Dual Union All
  Select '������Ŀ����','SELECT' From Dual Union All
  Select 'ҩƷ���','SELECT' From Dual Union All
  Select 'ҩƷ�����޶�','SELECT' From Dual Union All
  Select '��Ӧ��','SELECT' From Dual Union All
  Select '��ҩ����','SELECT' From Dual Union All
  Select '������ü�¼','SELECT' From Dual Union All
  Select '������Ϣ','SELECT' From Dual Union All
  Select '���','SELECT' From Dual Union All
  Select '����ҽ����¼','SELECT' From Dual Union All
  Select '�������ҽ��','SELECT' From Dual Union All
  Select '������ϼ�¼','SELECT' From Dual Union All
  Select 'סԺ���ü�¼','SELECT' From Dual Union All
  Select '����ҽ������','SELECT' From Dual Union All
  Select 'ҽ��ִ��ʱ��','SELECT' From Dual Union All
  Select 'ZL_ҩƷ�豸�ӿ�_UPDATE','EXECUTE' From Dual Union All
  Select 'ZL_ҩƷ�豸�ӿ�_STATE','EXECUTE' From Dual Union All
  Select 'ZL_ҩƷ�豸�ӿ�_DELETE','EXECUTE' From Dual Union All
  Select 'ZL_FUN_DRUG_MACHINE','EXECUTE' From Dual Union All
  Select 'ZL_δ��ҩƷ��¼_���䷢ҩ����','EXECUTE' From Dual Union All
  Select 'ZL_ҩƷ�շ������־_FLAG','EXECUTE' From Dual Union All
  Select 'ZL_ҩƷ�շ�סԺ��־_FLAG','EXECUTE' From Dual Union All
  Select 'ZL_DRUG_MAC_WIN','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;




--Ӧ������
Insert Into zlBakTables(ϵͳ,���,����,���,ֱ��ת��,ͣ�ô�����)
Select &n_System,2,A.* From (
Select ����,���,ֱ��ת��,ͣ�ô����� From zlBakTables Where 1 = 0 Union All 
  Select 'ҩƷ�շ������־',8,1,-Null From Dual Union All 
  Select 'ҩƷ�շ�סԺ��־',9,1,-Null From Dual Union All 
Select ����,���,ֱ��ת��,ͣ�ô����� From ZLBAKTABLES Where 1 = 0) A;


Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 9010, 0, 0, 0, 0, -null, -null, 1, '����ҩƷ�Զ����豸�ӿ�', '0', '0',
         '�Ƿ�����ҩƷ�Զ����豸�ӿ���������ӿ��ṩZLHIS����', '0-�����ã�1-����', Null, Null, Null
  From Dual
  Where Not Exists (Select 1 From zlParameters Where ������ = '����ҩƷ�Զ����豸�ӿ�' And ģ�� = 9010 And ϵͳ = &n_System);
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 9010, 0, 0, 0, 0, -null, -null, 2, '������Ϣ����ƽ̨', '0|', '0|',
         '����ʱ��������ȷ���ӿ��Ƿ�����Ϣ����ƽ̨��������ӿڽ���', '������0-�����ã�1-���á������ң���Ϣ����ƽ̨��WebService��ַ', Null, Null, Null
  From Dual
  Where Not Exists (Select 1 From zlParameters Where ������ = '������Ϣ����ƽ̨' And ģ�� = 9010 And ϵͳ = &n_System);
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 9010, 0, 0, 0, 0, -null, -null, 3, '��Ϣ����ƽ̨��Կ', '', '',
         '��Ϣ����ƽ̨��Կ', Null, Null, Null, Null
  From Dual
  Where Not Exists (Select 1 From zlParameters Where ������ = '��Ϣ����ƽ̨��Կ' And ģ�� = 9010 And ϵͳ = &n_System);




--���̺���
CREATE OR REPLACE Procedure Zl_ҩƷ�豸�ӿ�_Update
(
  ���_In     In ҩƷ�豸�ӿ�.���%Type,
  ����_In     In ҩƷ�豸�ӿ�.����%Type,
  ����_In     In ҩƷ�豸�ӿ�.����%Type,
  ������Ϣ_In In ҩƷ�豸�ӿ�.������Ϣ%Type,
  ��չ��Ϣ_In In Varchar2,
  Id_In       In ҩƷ�豸�ӿ�.Id%Type := Null,
  ��ע_In     In ҩƷ�豸�ӿ�.��ע%Type := Null
) Is

  v_Error Varchar2(255);
  Err_Custom Exception;

Begin

  --���ܣ�ҩƷ�豸�ӿڱ��������޸ļ�¼

  If Id_In Is Null Then
    --����
    Insert Into ҩƷ�豸�ӿ�
      (ID, ���, ����, ����, ������Ϣ, ��չ��Ϣ, ��ע)
    Values
      (ҩƷ�豸�ӿ�_Id.Nextval, ���_In, ����_In, ����_In, ������Ϣ_In, ��չ��Ϣ_In, ��ע_In);
  Else
    --�޸� 
    Update ҩƷ�豸�ӿ�
    Set ��� = ���_In, ���� = ����_In, ���� = ����_In, ������Ϣ = ������Ϣ_In, ��չ��Ϣ = ��չ��Ϣ_In, ��ע = ��ע_In
    Where ID = Id_In;
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�豸�ӿ�_Update;
/

CREATE OR REPLACE Procedure Zl_ҩƷ�豸�ӿ�_State
(
  Id_In   In ҩƷ�豸�ӿ�.Id%Type,
  ����_In In Number
) Is

  v_Error Varchar2(255);
  Err_Custom Exception;

Begin

  --���ܣ�ҩƷ�豸�ӿڵ�״̬����

  If Id_In Is Null Then
    v_Error := 'ҩƷ�豸�ӿ�ID����ȷ��';
    Raise Err_Custom;
  End If;

  If ����_In = 1 Then
    --����
    Update ҩƷ�豸�ӿ� Set �������� = Sysdate, ͣ������ = Null Where ID = Id_In;
  Else
    --ͣ��
    Update ҩƷ�豸�ӿ� Set ͣ������ = Sysdate Where ID = Id_In;
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�豸�ӿ�_State;
/

CREATE OR REPLACE Procedure Zl_ҩƷ�豸�ӿ�_Delete(Id_In In ҩƷ�豸�ӿ�.Id%Type) Is

  v_Error Varchar2(255);
  Err_Custom Exception;

Begin

  --���ܣ�ҩƷ�豸�ӿڱ�ɾ����¼

  Delete From ҩƷ�豸�ӿ� Where ID = Id_In;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�豸�ӿ�_Delete;
/

Create Or Replace Function Zl_Fun_Drug_Machine
(
  �ⷿid_In   In ���ű�.Id%Type,
  ҩƷ����_In In ҩƷ����.����%Type,
  �շ�id_In   In ҩƷ�շ���¼.Id%Type := Null
) Return ҩƷ�豸�ӿ�.���%Type Is

  v_Code ҩƷ�豸�ӿ�.���%Type;

Begin

  --���ܣ����������Ӧ�Ľӿڱ��
  --˵����ҩƷ�Զ����豸�ӿڲ�����ר�ú�����
  --������
  --  �շ�ID_In����չ��������׼���ò�����

  Begin
    Select a.���
    Into v_Code
    From ҩƷ�豸�ӿ� A,
         Xmltable('//root/bm' Passing a.��չ��Ϣ Columns �ⷿid Number(18) Path 'id', ���ͱ��� Varchar2(20) Path 'jxbm') B, ҩƷ���� C
    Where b.���ͱ��� = c.���� And a.ͣ������ Is Null And a.�������� Is Not Null And b.�ⷿid = �ⷿid_In And c.���� = ҩƷ����_In And
          Rownum < 2;
  Exception
    When Others Then
      Begin
        Select a.���
        Into v_Code
        From ҩƷ�豸�ӿ� A,
             Xmltable('//root/bm' Passing a.��չ��Ϣ Columns �ⷿid Number(18) Path 'id', ���ͱ��� Varchar2(20) Path 'jxbm') B
        Where a.ͣ������ Is Null And a.�������� Is Not Null And (b.���ͱ��� = '' Or b.���ͱ��� Is Null) And b.�ⷿid = �ⷿid_In And
              Rownum < 2;
      Exception
        When Others Then
          v_Code := Null;
      End;
  End;

  Return v_Code;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Fun_Drug_Machine;
/

CREATE OR REPLACE Procedure Zl_ҩƷ�շ������־_Flag
(
  ҵ�����_In In ҩƷ�շ������־.ҵ�����%Type,
  �ⷿid_In   In ҩƷ�շ������־.�ⷿid%Type,
  ������Ϣ_In In Varchar2,
  ���ͱ�־_In In Number
) Is
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin

  --����
  --  ������Ϣ������1,������1;����2,������2...

  For r_Tmp In (Select b.c2 ������, b.c1 ����, a.��־
                From ҩƷ�շ������־ A, Table(f_Str2list2(������Ϣ_In, ';', ',')) B
                Where a.������(+) = b.C2 And a.����(+) = b.C1 And a.�ⷿid(+) = �ⷿid_In And a.ҵ�����(+) = ҵ�����_In) Loop
    If r_Tmp.��־ Is Null Then
      Delete ҩƷ�շ������־ Where ������ = r_Tmp.������ And ���� = r_Tmp.���� And �ⷿid = �ⷿid_In;
      If ���ͱ�־_In = 1 Then
        Insert Into ҩƷ�շ������־
          (������, ����, �ⷿid, ҵ�����, ��־)
        Values
          (r_Tmp.������, r_Tmp.����, �ⷿid_In, ҵ�����_In, 1);
      Else
        Insert Into ҩƷ�շ������־
          (������, ����, �ⷿid, ҵ�����, ��־)
        Values
          (r_Tmp.������, r_Tmp.����, �ⷿid_In, ҵ�����_In, 11);
      End If;
    Elsif r_Tmp.��־ Between 11 And 12 Then
      Update ҩƷ�շ������־
      Set ��־ = ��־ + 1
      Where ������ = r_Tmp.������ And ���� = r_Tmp.���� And �ⷿid = �ⷿid_In And ҵ����� = ҵ�����_In;
    End If;
  End Loop;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�շ������־_Flag;
/

CREATE OR REPLACE Procedure Zl_ҩƷ�շ�סԺ��־_Flag
(
  ҵ�����_In In ҩƷ�շ�סԺ��־.ҵ�����%Type,
  ҽ����Ϣ_In In Varchar2,
  ���ͱ�־_In In Number
) Is
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin

  --����
  --  ҽ����Ϣ��ҽ��id1;ҽ��2...

  For r_Tmp In (Select b.Column_Value �շ�id, a.��־
                From ҩƷ�շ�סԺ��־ A, Table(f_Str2list(ҽ����Ϣ_In, ';')) B
                Where a.�շ�id(+) = b.Column_Value And a.ҵ�����(+) = ҵ�����_In) Loop
    If r_Tmp.��־ Is Null Then
      Delete ҩƷ�շ�סԺ��־ Where �շ�id = r_Tmp.�շ�id;
      If ���ͱ�־_In = 1 Then
        Insert Into ҩƷ�շ�סԺ��־ (�շ�id, ҵ�����, ��־) Values (r_Tmp.�շ�id, ҵ�����_In, 1);
      Else
        Insert Into ҩƷ�շ�סԺ��־ (�շ�id, ҵ�����, ��־) Values (r_Tmp.�շ�id, ҵ�����_In, 11);
      End If;
    Elsif r_Tmp.��־ Between 11 And 12 Then
      Update ҩƷ�շ�סԺ��־ Set ��־ = ��־ + 1 Where �շ�id = r_Tmp.�շ�id And ҵ����� = ҵ�����_In;
    End If;
  End Loop;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�շ�סԺ��־_Flag;
/

Create Or Replace Procedure Zl_Drug_Mac_Win
(
  No_In       In Varchar2,
  �ⷿid_In   In ҩƷ�շ���¼.�ⷿid%Type,
  ���ڱ���_In In ��ҩ����.����%Type,
  ����id_In   In ����ҽ����¼.����id%Type := Null
) Is
  v_Error Varchar2(255);
  Err_Custom Exception;
  v_No   ҩƷ�շ���¼.No%Type;
  v_Tmp  Varchar2(50);
  n_Bill ҩƷ�շ���¼.����%Type;
Begin

  --���ܣ�������֪ͨZLHIS���ﴦ����ҩ���ڵ���
  --������
  --  NO_In���豸��NO��ʽ��������_����_�ⷿid

  If No_In Is Null Or No_In = '' Then
    v_Error := '������Ϣ��';
    Raise Err_Custom;
  End If;

  If ���ڱ���_In Is Null Or ���ڱ���_In = '' Then
    v_Error := '������Ϣ��';
    Raise Err_Custom;
  End If;

  v_No := Substr(No_In, 1, 8);

  If Length(No_In) >= 10 Then
    v_Tmp := Substr(No_In, 10);
  End If;

  If v_Tmp Is Null Or v_Tmp = '' Then
    v_Error := '������Ϣ�쳣';
    Raise Err_Custom;
  End If;

  Select Column_Value Into n_Bill From Table(f_Num2list(v_Tmp, '_')) Where Rownum < 2;

  Zl_δ��ҩƷ��¼_���䷢ҩ����(v_No, n_Bill, �ⷿid_In, ���ڱ���_In);

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drug_Mac_Win;
/

Create Or Replace Procedure Zl1_Datamove_Reb
(
  System_In    In Number,
  Speedmode_In In Number,
  Func_In      In Number,
  Enable_In    In Number := 0,
  Parallel_In  In Number := 0,
  Rebscope_In  In Number := 0
) As
  --���ܣ�����ʷ����ת��֮ǰ�����ô��������Զ���ҵ��Լ����������ת��֮��������Щ�����Լ��ؽ���ת���������ջر��ת�����������Ŀռ� 
  --������ 
  --System_In:    Ӧ��ϵͳ���,100=��׼�� 
  --speedmode_in������ת��ģʽ��0-����ģʽ��1-����ģʽ���ڿͻ���ͣ��ʱ��ת���ڼ����ת�����������Ψһ�������Լ�����������Լӿ���ת���ݵ�ɾ�������� 
  --func_in:      1=��������2=�Զ���ҵ��3=Լ����4=������5=�ؽ���ת��������6-�ջر��ת�����������Ŀռ䣬7-�����Ĵ洢�ռ䣨move�������ָ������õ�Լ�������� ,8-�ؽ����ת����ѯ��������������������� 
  --Enable_in:    0-���ã�1=���ã���func_inֵΪ1-4��Ч 
  --rebScope_in:   Func_In=6ʱ��ָ�ؽ������ķ�Χ(0-���ú�����,1-���ú����༰ҽ����,2-ȫ��)��Func_In=7ʱָMove��ķ�Χ(0-���ú����࣬1-ȫ��) 

  v_Sql      Varchar2(4000);
  n_Do       Number(1);
  n_Parallel Number(1);
  v_Tbs      Varchar2(100);

  --ת������е�SQL��ѯ���������
  v_Indexeswithtag Varchar2(4000) := '������ü�¼_IX_����ID,סԺ���ü�¼_IX_����ID,���ò����¼_IX_����ID,���ò����¼_IX_�Ǽ�ʱ��,����Ԥ����¼_IX_��ҳID,����Ԥ����¼_IX_����ID,����Ԥ����¼_IX_�տ�ʱ��,������ü�¼_IX_�Ǽ�ʱ��,������ü�¼_IX_ҽ�����,סԺ���ü�¼_IX_�Ǽ�ʱ��,���˽��ʼ�¼_IX_�շ�ʱ��,���˽��ʼ�¼_IX_����id' ||
                                     ',ҩƷ�շ���¼_IX_����ID,�շ���¼������Ϣ_IX_�շ�ID,��Һ��ҩ����_IX_�շ�ID,ҩƷ����ƻ�_IX_����ID,ҩƷǩ����ϸ_IX_�շ�ID' ||
                                     ',��Ա����¼_IX_���ʱ��,��Ա�սɼ�¼_IX_�Ǽ�ʱ��,��Ա�ݴ��¼_IX_�ս�ID,��Ա�ݴ��¼_IX_�Ǽ�ʱ��,Ʊ�����ü�¼_IX_�Ǽ�ʱ��,Ʊ��ʹ����ϸ_IX_����ID,Ʊ�ݴ�ӡ��ϸ_IX_ʹ��ID' ||
                                     ',���˹Һż�¼_IX_�Ǽ�ʱ��,����ҽ������_IX_����ʱ��,����ҽ����¼_IX_�Һŵ�,����ҽ����¼_IX_��ҳID,����ҽ����¼_IX_���ID' ||
                                     ',������ҳ_IX_��Ժ����,סԺ���ü�¼_IX_����ID,���˹�����¼_IX_����ID,������ϼ�¼_IX_����ID,���������¼_IX_��ҳID' ||
                                     ',���˻����¼_IX_��ҳID,���˻�������_IX_��¼id,���˻����ļ�_IX_��ҳID,���˻�������_IX_�ļ�ID,���˻�����ϸ_IX_��¼ID,���˻����ӡ_IX_�ļ�ID' ||
                                     ',���Ӳ�����¼_IX_����ID,����ҽ������_IX_����ID,Ӱ�񱨸沵��_IX_ҽ��ID,������ļ�¼_IX_����ID,������ϼ�¼_IX_����ID' ||
                                     ',�����ٴ�·��_IX_����ID,���˺ϲ�·��_IX_��Ҫ·����¼ID,����·��ִ��_IX_·����¼ID,���˳�����¼_IX_·����¼ID,�������ҽ��_IX_ҽ��ID' ||
                                     ',Ӱ�񱨸��¼_IX_ҽ��ID,Ӱ�񱨸������¼_IX_ҽ��ID,Ӱ�����뵥ͼ��_IX_ҽ��ID,Ӱ���ղ�����_IX_ҽ��ID,����걾��¼_IX_ҽ��ID,������Ŀ�ֲ�_IX_�걾ID,���������¼_IX_�걾ID' ||
                                     ',���������¼_IX_�걾ID,����ͼ����_IX_�걾ID,������ռ�¼_IX_ҽ��ID,������ͨ���_IX_����걾ID,���������ϸ_IX_ҽ��ID';

  --ת������е�SQL��ѯ���������(������Ψһ����Ӧ������)
  v_Constraintswithtag Varchar2(4000) := '����Ԥ����¼_UQ_NO,���˽��ʼ�¼_UQ_NO,���˽��ʼ�¼_PK,������ü�¼_UQ_NO,סԺ���ü�¼_UQ_NO,ҽ��������ϸ_PK' ||
                                         ',���˿��������_PK,���ò����¼_PK,���˿������¼_PK,�������㽻��_PK,�����˿���Ϣ_PK,��Һ��ҩ��¼_PK,ҩƷǩ����¼_PK,Ʊ�ݴ�ӡ����_PK,���˹Һż�¼_PK,���˹ҺŻ���_UQ_����,����ת���¼_UQ_NO' ||
                                         ',���˻�����Ŀ_UQ_ҳ��,���˻���Ҫ������_UQ_ҳ��,����Ҫ������_PK,���Ӳ�����¼_PK,���Ӳ�������_PK,���Ӳ�����ʽ_PK,���Ӳ�������_UQ_�������,���Ӳ���ͼ��_PK,�����걨��¼_PK,�������淴��_PK' ||
                                         ',���˺ϲ�·������_PK,����·������_PK,����·������_PK,����·��ָ��_UQ_����ָ��,����·��ҽ��_PK' ||
                                         ',����ҽ����¼_PK,����ҽ������_PK,����ҽ���Ƽ�_UQ_�շ�ϸĿID,����ҽ������_PK,����ҽ������_PK,����ҽ��ִ��_PK,ҽ��ִ��ʱ��_PK,ҽ��ִ�д�ӡ_PK,����ҽ����ӡ_UQ_ҽ��ID,��Ѫ�����¼_PK,��Ѫ������_PK' ||
                                         ',������ϼ�¼_PK,����ҽ��״̬_PK,ҽ��ǩ����¼_PK,����ҽ������_PK,���Ƶ��ݴ�ӡ_PK,ҽ��ִ�мƼ�_PK,ִ�д�ӡ��¼_PK' ||
                                         ',Ӱ�����¼_PK,Ӱ��������_UQ_���к�,Ӱ����ͼ��_UQ_ͼ���,Ӱ��Σ��ֵ��¼_UQ_ҽ��ID' ||
                                         ',����������Ŀ_PK,�����ʿؼ�¼_PK,����ǩ����¼_PK,�����Լ���¼_PK,�����ʿر���_PK,����ҩ�����_PK,��Ա�սɼ�¼_PK,��Ա�ս���ϸ_PK,��Ա�ս�Ʊ��_PK,��Ա�սɶ���_PK' ||
                                         ',��������¼_PK,���������_UQ_��ID,�����嵥��ӡ_UQ_NO,RIS���ԤԼ_PK,ҩƷ�շ������־_PK,ҩƷ�շ�סԺ��־_PK';

  --���ܣ�1.���û���������ת�����������������,����ɾ�������¼ʱ���ӱ�ÿ�м�¼ִ��һ��SQL��ѯ��ɾ�� 
  --      2.���û�����������Ψһ��Լ��������ʱ���Զ�ɾ����Ӧ������������ʱ�Զ������������������ɾ������ 
  --���磺����ҽ������_FK_ҽ��ID�������Щ������ڵı�����δת����δ��zlbaktables���ж��壩��ִ��ǰ���鲢����ת���� 
  Procedure Setconstraintstatus As
    v_Pcol Varchar2(50);
    v_Fcol Varchar2(50);
    v_Del  Varchar2(4000);
  Begin
    --����ʱ���Ƚ�������ת��������������������ٽ���ת��������� 
    If Enable_In = 0 Then
      --1.����ģʽת��ʱ��������ҵ�����ɾ�����������ԣ����ڼ���ɾ����������ô�������������ӱ����ݵ�ɾ������
      If Speedmode_In = 0 Then
        For Rp In (Select Distinct a.Table_Name As Ptable_Name, a.Constraint_Name
                   From User_Constraints A, User_Constraints C, zlBakTables B
                   Where a.Table_Name = b.���� And b.ֱ��ת�� = 1 And b.ϵͳ = System_In And a.Constraint_Type In ('P', 'U') And
                         c.r_Constraint_Name = a.Constraint_Name And c.Constraint_Type = 'R' And
                         c.Delete_Rule = 'CASCADE'
                   Order By a.Table_Name) Loop
        
          Select f_List2str(Cast(Collect(Column_Name Order By Position) As t_Strlist))
          Into v_Pcol
          From User_Cons_Columns
          Where Constraint_Name = Rp.Constraint_Name;
        
    v_Del := '';
          For Rf In (Select b.Table_Name, b.Constraint_Name,
                            f_List2str(Cast(Collect(b.Column_Name Order By b.Position) As t_Strlist)) As r_Col
                     From User_Constraints A, User_Cons_Columns B
                     Where a.r_Constraint_Name = Rp.Constraint_Name And a.Constraint_Name = b.Constraint_Name
                     Group By b.Table_Name, b.Constraint_Name) Loop
            If Instr(v_Pcol, ',') > 0 Then
              v_Del := v_Del || Chr(10) || '        Delete ' || Rf.Table_Name || ' Where (' || Rf.r_Col ||
                       ') in ((:Old.' || Replace(v_Pcol, ',', ',:Old.') || '));';
            Else
              v_Del := v_Del || Chr(10) || '        Delete ' || Rf.Table_Name || ' Where ' || Rf.r_Col || ' = :Old.' ||
                       v_Pcol || ';';
            End If;
          End Loop;
        
          v_Sql := 'Create Or Replace Trigger ' || Rp.Ptable_Name || '_Cascade_Del' || Chr(10) ||
                   '    After Delete On ' || Rp.Ptable_Name || Chr(10) || '    For Each Row' || Chr(10) || 'Begin' ||
                   Chr(10) || '    If :Old.��ת�� Is Null Then ' || v_Del || Chr(10) || '    End If; ' || Chr(10) ||
                   'End ' || Rp.Ptable_Name || '_Cascade_Del;';
          Execute Immediate v_Sql;
        End Loop;
      End If;
    
      --2.��������ת�����������������
      For R In (Select c.Table_Name, c.Constraint_Name, a.Table_Name As Ptable_Name
                From User_Constraints A, User_Constraints C, zlBakTables B
                Where a.Table_Name = b.���� And b.ֱ��ת�� = 1 And b.ϵͳ = System_In And a.Constraint_Type In ('P', 'U') And
                      c.r_Constraint_Name = a.Constraint_Name And c.Constraint_Type = 'R' And c.Status = 'ENABLED'
                Order By a.Table_Name) Loop
        v_Sql := 'Alter Table ' || r.Table_Name || ' Disable Constraint ' || r.Constraint_Name;
        Execute Immediate v_Sql;
      End Loop;
    
      --3.����������Ψһ������(����ת��ʱ)
      If Speedmode_In = 1 Then
        --����ɾ������������ʹskip_unusable_indexesΪtrue��Ҳ�޷�ɾ������Unusable״̬��Ψһ�������ı��еļ�¼
        --����ת������е�SQL��ѯ���������(������Ψһ����Ӧ������) 
        For R In (Select a.Table_Name, a.Constraint_Name
                  From User_Constraints A, zlBakTables T, User_Tables B
                  Where a.Table_Name = t.���� And t.ֱ��ת�� = 1 And t.ϵͳ = System_In And a.Status = 'ENABLED' And
                        a.Constraint_Type In ('P', 'U') And a.Table_Name = b.Table_Name And b.Iot_Type Is Null And
                        a.Constraint_Name Not In
                        (Select Upper(Column_Value) As Constraint_Name From Table(f_Str2list(v_Constraintswithtag)))
                  Order By Constraint_Name) Loop
          v_Sql := 'Alter Table ' || r.Table_Name || ' Disable Constraint ' || r.Constraint_Name ||
                   ' Cascade Drop Index';
          Execute Immediate v_Sql;
        End Loop;
      End If;
    Else
      --����ʱ
      --1.������������Ψһ��������������ת����������������� 
      If Speedmode_In = 1 Then
        --���ؽ�������������Լ�����Ա��ؽ�����ʱ���ò���ִ������ʱ�䣬��������Լ��ʱҲ���Բ���novalidate��ʽ 
        For R In (Select d.Table_Name, d.Constraint_Name,
                         f_List2str(Cast(Collect(d.Column_Name Order By d.Position) As t_Strlist)) Colstr
                  From User_Cons_Columns D,
                       (Select a.Table_Name, a.Constraint_Name
                         From User_Constraints A, zlBakTables T
                         Where a.Table_Name = t.���� And t.ֱ��ת�� = 1 And t.ϵͳ = System_In And a.Status = 'DISABLED' And
                               a.Constraint_Type In ('P', 'U')) A
                  Where a.Constraint_Name = d.Constraint_Name And a.Table_Name = d.Table_Name
                  Group By d.Table_Name, d.Constraint_Name
                  Order By Constraint_Name) Loop
          Update Zldatamovelog
          Set ��ǰ���� = '���ڻָ�Լ��:' || r.Constraint_Name
          Where ϵͳ = System_In And ���� = (Select Max(����) From Zldatamovelog Where ϵͳ = System_In);
        
          Select Tablespace_Name Into v_Tbs From User_Indexes Where Table_Name = r.Table_Name And Rownum < 2;
        
          --����������Ψһ��ʱ�������Ǳ�ɾ���˵ģ���������Ҫ��Create 
          v_Sql := 'Create Unique Index ' || r.Constraint_Name || ' On ' || r.Table_Name || '(' || r.Colstr ||
                   ') Tablespace ' || v_Tbs || ' Nologging';
          Begin
            Execute Immediate v_Sql;
          Exception
            When Others Then
              Null; --������Щ������Ψһ�����Ǳ���ת���ڼ䱻���õģ�֮ǰ�ʹ��ڲ�Ψһ���ݣ�����Ψһ��������� 
          End;
        
          --���Զ�����Լ���������Ĺ��� 
          v_Sql := 'Alter Table ' || r.Table_Name || ' Enable Novalidate Constraint ' || r.Constraint_Name;
          Execute Immediate v_Sql;
        End Loop;
      End If;
    
      --2.��������ת����������������� 
      For R In (Select c.Table_Name, c.Constraint_Name
                From User_Constraints A, User_Constraints C, zlBakTables B
                Where a.Table_Name = b.���� And b.ֱ��ת�� = 1 And b.ϵͳ = System_In And a.Constraint_Type In ('P', 'U') And
                      c.r_Constraint_Name = a.Constraint_Name And c.Constraint_Type = 'R' And c.Status = 'DISABLED'
                Order By a.Table_Name) Loop
        --Ϊ�˼ӿ��ٶȣ�����novalidate������֤�������� 
        --��������ת����������������zlbaktables�ж����ˣ���û�б�д��Ӧ������ת���ű���δ��֤�����ݿ�����Υ��Լ��������� 
        v_Sql := 'Alter Table ' || r.Table_Name || ' Enable Novalidate Constraint ' || r.Constraint_Name;
        Execute Immediate v_Sql;
      End Loop;
    
      --3.����ģʽת��ʱ��ɾ��֮ǰ�����������������ɾ������Ĵ�����
      If Speedmode_In = 0 Then
        For R In (Select a.Trigger_Name
                  From User_Triggers A, zlBakTables B
                  Where a.Table_Name = b.���� And b.ֱ��ת�� = 1 And b.ϵͳ = System_In And
                        Trigger_Name = Table_Name || '_CASCADE_DEL' And Triggering_Event = 'DELETE') Loop
          v_Sql := 'Drop Trigger ' || r.Trigger_Name;
          Execute Immediate v_Sql;
        End Loop;
      End If;
    End If;
  End Setconstraintstatus;

  --���ܣ�����ģʽʱ����LOB�������������������ģʽʱ������ת�������÷�ת������������(���磺����ҽ���Ƽ�_IX_�շ�ϸĿID) 
  --˵��������������Ϊ�����ɾ�����ݵ����� 
  Procedure Setindexstatus As
  Begin
    If Speedmode_In = 1 Then
      --����ת������е�SQL��ѯ��������� 
      For R In (Select /*+ rule*/
                 a.Index_Name
                From User_Indexes A, zlBakTables T
                Where a.Table_Name = t.���� And t.ֱ��ת�� = 1 And t.ϵͳ = System_In And t.ֱ��ת�� = 1 And
                      a.Index_Name <> a.Table_Name || '_IX_��ת��' And
                      a.Index_Name Not In
                      (Select Upper(Column_Value) As Index_Name From Table(f_Str2list(v_Indexeswithtag))) And
                      a.Status = Decode(Enable_In, 0, 'VALID', 'UNUSABLE') And a.Index_Type = 'NORMAL' And Not Exists
                 (Select 1
                       From User_Constraints C
                       Where c.Index_Name = a.Index_Name And c.Constraint_Type In ('P', 'U'))
                Order By Index_Name) Loop
      
        If Enable_In = 0 Then
          v_Sql := 'Alter Index ' || r.Index_Name || ' Unusable';
          Execute Immediate v_Sql;
        Else
          Update Zldatamovelog
          Set ��ǰ���� = '�����ؽ�����:' || r.Index_Name
          Where ϵͳ = System_In And ���� = (Select Max(����) From Zldatamovelog Where ϵͳ = System_In);
        
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Nologging';
          Begin
            Execute Immediate v_Sql;
            --�����ؽ��Ƚ������������ؽ�����Ҫ���������������������������ORA-00054: ��Դ��æ, ��ָ���� NOWAIT ��ʽ��ȡ��Դ 
          
          Exception
            When Others Then
              If SQLErrM Like 'ORA-00054%' Then
                v_Sql := Replace(v_Sql, 'Rebuild', 'Rebuild Online');
                Execute Immediate v_Sql;
              End If;
          End;
        End If;
      End Loop;
    Else
      For R In (Select a.Index_Name
                From (Select d.Table_Name, d.Index_Name,
                              f_List2str(Cast(Collect(d.Column_Name Order By d.Column_Position) As t_Strlist)) Colstr
                       From User_Ind_Columns D, zlBakTables T, User_Indexes C
                       Where c.Table_Name = t.���� And t.ֱ��ת�� = 1 And t.ϵͳ = System_In And c.Uniqueness = 'NONUNIQUE' And
                             c.Index_Type = 'NORMAL' And c.Status = Decode(Enable_In, 0, 'VALID', 'UNUSABLE') And
                             c.Index_Name = d.Index_Name And c.Table_Name = d.Table_Name
                       Group By d.Table_Name, d.Index_Name) A,
                     (Select e.Table_Name,
                              f_List2str(Cast(Collect(e.Column_Name Order By e.Position) As t_Strlist)) Colstr
                       From User_Cons_Columns E, User_Constraints F, zlBakTables T, User_Constraints C
                       Where e.Table_Name = t.���� And t.ֱ��ת�� = 1 And t.ϵͳ = System_In And
                             e.Constraint_Name = f.Constraint_Name And f.Constraint_Type = 'R' And
                             c.Constraint_Name = f.r_Constraint_Name And c.Table_Name Not In ('������ҳ', '������Ϣ') And
                             Not Exists
                        (Select 1 From zlBakTables G Where g.���� = c.Table_Name And g.ϵͳ = System_In)
                       Group By e.Table_Name, e.Constraint_Name) B
                Where a.Table_Name = b.Table_Name And a.Colstr = b.Colstr
                Order By Index_Name) Loop
      
        If Enable_In = 0 Then
          --���⴦�������������������ã�������ҩƷĿ¼�޸Ĺ�񣬲���ɿ���Ҫʹ�� 
          If r.Index_Name Not In ('����ҽ����¼_IX_�շ�ϸĿID', 'ҩƷ�շ���¼_IX_ҩƷID', 'ҩƷ�շ���¼_IX_�۸�ID') Then
            v_Sql := 'Alter Index ' || r.Index_Name || ' Unusable';
            Execute Immediate v_Sql;
          End If;
        Else
          Update Zldatamovelog
          Set ��ǰ���� = '�����ؽ�����:' || r.Index_Name
          Where ϵͳ = System_In And ���� = (Select Max(����) From Zldatamovelog Where ϵͳ = System_In);
        
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Online Nologging';
          Execute Immediate v_Sql;
          --�����ؽ��Ƚ������������ؽ�����Ҫ���������������������������ORA-00054: ��Դ��æ, ��ָ���� NOWAIT ��ʽ��ȡ��Դ 
        End If;
      End Loop;
    End If;
  End Setindexstatus;

  --���ܣ�ת�������ڼ䣬ͣ��ת�����ϵ����д�������ת�����ٻָ� 
  Procedure Settriggerstatus As
  Begin
    For R In (Select Distinct a.Table_Name, t.ͣ�ô�����
              From User_Triggers A, zlBakTables T
              Where a.Status = Decode(Enable_In, 0, 'ENABLED', 'DISABLED') And a.Table_Name = t.���� And t.ֱ��ת�� = 1 And
                    t.ϵͳ = System_In) Loop
      If Enable_In = 0 Then
        v_Sql := 'Alter Table ' || r.Table_Name || ' DISABLE ALL TRIGGERS';
        Update zlBakTables Set ͣ�ô����� = 1 Where ϵͳ = System_In And ���� = r.Table_Name;
      Elsif Nvl(r.ͣ�ô�����, 0) = 1 Then
        v_Sql := 'Alter Table ' || r.Table_Name || ' ENABLE ALL TRIGGERS';
        Update zlBakTables Set ͣ�ô����� = Null Where ϵͳ = System_In And ���� = r.Table_Name;
      End If;
      Execute Immediate v_Sql;
    End Loop;
    Commit;
  End Settriggerstatus;

  --���ܣ�ת�������ڼ䣬ͣ�õ�ǰ�����ߵ������Զ���ҵ��ת���������� 
  Procedure Setjobstatus As
    v_Jobs Varchar2(4000);
  Begin
    --ͣ�� 
    If Enable_In = 0 Then
      For R In (Select Job From User_Jobs Where Broken = 'N') Loop
        Dbms_Job.Broken(r.Job, True);
        v_Jobs := v_Jobs || ',' || r.Job;
      End Loop;
    
      If v_Jobs Is Not Null Then
        v_Jobs := Substr(v_Jobs, 2);
        Update zlDataMove Set ͣ����ҵ�� = v_Jobs Where ϵͳ = System_In And ��� = 1;
      End If;
    Else
      --���� 
      Select ͣ����ҵ�� Into v_Jobs From zlDataMove Where ϵͳ = System_In And ��� = 1;
      If v_Jobs Is Not Null Then
        For R In (Select Job
                  From User_Jobs
                  Where Broken = 'Y' And Job In (Select Column_Value From Table(f_Num2list(v_Jobs)))) Loop
          Dbms_Job.Broken(r.Job, False);
        End Loop;
        Update zlDataMove Set ͣ����ҵ�� = Null Where ϵͳ = System_In And ��� = 1;
      End If;
    End If;
    --��ҵ���ú�����ύ�������Ч 
    Commit;
  End Setjobstatus;
Begin
  If Parallel_In < 2 Then
    Execute Immediate 'Alter Session DISABLE PARALLEL DDL';
  Else
    If Func_In In (6, 7, 8) Or Func_In In (3, 4) And Enable_In = 1 Then
      --Ϊ�ؽ��������ò���ִ�У�����ͨ��������IO�豸�����ܣ�����̫�ߵĲ��жȷ����ή�����ܣ����и����ܴ洢�豸���ɼӴ��жȣ� 
      --ִ���ؽ���������Զ�Ϊ�������ϲ��ж����ԣ������ȡ������Ӱ�����SQL��ִ�мƻ�(ȫ��ɨ��+���в�ѯ������),�ں���ȡ�������Ĳ��ж� 
      --�ָ����߿��Լ��������ʱ�������ǲ�������ģʽ�������ϲ��У�����̫��
      Execute Immediate 'Alter Session FORCE PARALLEL DDL PARALLEL ' || Parallel_In;
      n_Parallel := 1;
    End If;
  End If;

  If Func_In = 1 Then
    --1.���ô����� 
    Settriggerstatus;
  Elsif Func_In = 2 Then
    --2.�����Զ���ҵ 
    Setjobstatus;
  Elsif Func_In = 3 Then
    --3.����Լ��״̬ 
    Setconstraintstatus;
  Elsif Func_In = 4 Then
    --4.��������״̬ 
    Setindexstatus;
  Elsif Func_In = 5 Then
    --5.�ؽ�"��ת��"���� 
    For R In (Select b.Index_Name
              From zlBakTables A, User_Indexes B
              Where a.���� = b.Table_Name And a.ֱ��ת�� = 1 And a.ϵͳ = System_In And b.Index_Name = b.Table_Name || '_IX_��ת��'
              Union All
              Select '������ҳ_IX_��ת��'
              From Dual
              Where System_In = 100) Loop
      Update Zldatamovelog
      Set ��ǰ���� = '�����ؽ�����:' || r.Index_Name
      Where ϵͳ = System_In And ���� = (Select Max(����) From Zldatamovelog Where ϵͳ = System_In);
    
      --��ʱ̫�̣����벢��DDL 
      --����ת��ʱ����ؽ����������������������������������ORA-00054: ��Դ��æ, ��ָ���� NOWAIT ��ʽ��ȡ��Դ 
      --�����ؽ�����̫�������ԣ���ʹ����ת��ģʽҲ���������ؽ�
      v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Nologging';
      Begin
        Execute Immediate v_Sql;
      Exception
        When Others Then
          If SQLErrM Like 'ORA-00054%' Then
            v_Sql := Replace(v_Sql, 'Rebuild', 'Rebuild Online');
            Execute Immediate v_Sql;
          End If;
      End;
    End Loop;
  
  Elsif Func_In = 6 Then
    --6.�ؽ����ת����ѯ���õ������������Ա����ؽ�����������һ��Ĳ�ѯʱ�䣩 
    --����ҵ������ý׶��������ؽ���Щ�������Ա���һЩ����Ҫ���ؽ���ʱ 
    For R In (Select b.Index_Name, a.���
              From User_Indexes B, zlBakTables A
              Where a.ϵͳ = System_In And a.���� = b.Table_Name And
                    b.Index_Name In (Select Upper(Column_Value)
                                     From Table(f_Str2list(v_Indexeswithtag))
                                     Union
                                     Select Upper(Column_Value)
                                     From Table(f_Str2list(v_Constraintswithtag)))
              Order By Index_Name) Loop
      n_Do := 0;
      If Rebscope_In = 0 Then
        If r.��� < 5 Then
          n_Do := 1; --�����ú����� 
        End If;
      Elsif Rebscope_In = 1 Then
        If r.��� < 5 Or r.��� = 8 Then
          n_Do := 1; --�����ú����ࡢҽ���� 
        End If;
      Else
        n_Do := 1;
      End If;
    
      If n_Do = 1 Then
        Update Zldatamovelog
        Set ��ǰ���� = '�����ؽ�����:' || r.Index_Name
        Where ϵͳ = System_In And ���� = (Select Max(����) From Zldatamovelog Where ϵͳ = System_In);
      
        --v_Sql := 'Alter Index ' || r.Index_Name || ' shrink Space'; 
        --ʹ��shrink��ʽ���ܲ���ִ��,��������ٶȱ�rebuild PARALLEL 8 ��6�� 
        If Speedmode_In = 1 Then
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Nologging';
        Else
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Online Nologging';
        End If;
        Begin
          Execute Immediate v_Sql;
          --�����ؽ��Ƚ������������ؽ�����Ҫ���������������������������ORA-00054: ��Դ��æ, ��ָ���� NOWAIT ��ʽ��ȡ��Դ 
        
        Exception
          When Others Then
            If SQLErrM Like 'ORA-00054%' Then
              v_Sql := Replace(v_Sql, 'Rebuild', 'Rebuild Online');
              Execute Immediate v_Sql;
            End If;
        End;
      End If;
    End Loop;
  
    --����������
  Elsif Func_In = 7 Then
    --rebScope_in=0,ֻ�������С��5�ľ��ú���������á�ҩƷ��Ʊ�ݣ�������ȫ������ 
    For R In (Select a.���� As Table_Name
              From zlBakTables A
              Where a.ֱ��ת�� = 1 And (��� < Decode(Rebscope_In, 0, 5, 100))
              Order By ���, ���) Loop
    
      Update Zldatamovelog
      Set ��ǰ���� = '���������:' || r.Table_Name
      Where ϵͳ = System_In And ���� = (Select Max(����) From Zldatamovelog Where ϵͳ = System_In);
    
      --����п��еĿռ䣬����Ƶ�������ռ䣬ֻ���������ܾ����ƶ��ļ�β�������ݿ飬�Ա���б�ռ��ļ������� 
      --��ǰ�������˻Ự����ǿ�Ʋ��� 
      v_Sql := 'Alter Table ' || r.Table_Name || ' Move Nologging';
      Execute Immediate v_Sql;
    
      --�����ƶ�Lob���� 
      For L In (Select Column_Name, Tablespace_Name From User_Lobs Where Table_Name = r.Table_Name) Loop
        v_Sql := 'Alter Table ' || r.Table_Name || ' Move Lob(' || l.Column_Name || ') Store as (Tablespace ' ||
                 l.Tablespace_Name || ') Nologging';
        Execute Immediate v_Sql;
      End Loop;
    
      v_Sql := 'Alter Table ' || r.Table_Name || ' Noparallel';
      Execute Immediate v_Sql;
    
      --move�󣬱���ص�������ȫ��ʧЧ����Ҫȫ���ؽ� 
      For S In (Select Index_Name
                From User_Indexes
                Where Table_Name = r.Table_Name And Status = 'UNUSABLE'
                Order By Index_Name) Loop
        Update Zldatamovelog
        Set ��ǰ���� = '���ڻָ�ʧЧ����:' || s.Index_Name
        Where ϵͳ = System_In And ���� = (Select Max(����) From Zldatamovelog Where ϵͳ = System_In);
      
        --��ǰ�������˻Ự����ǿ�Ʋ��� 
        v_Sql := 'Alter Index ' || s.Index_Name || ' Rebuild Nologging';
        Execute Immediate v_Sql;
      End Loop;
    End Loop;
    --�ؽ�ת�����ϱ��ת���������������������ת����ɺ��ջؿ��пռ䣩
    --ʧЧ���������ؽ�����Ϊת������е������ؽ�����
  Elsif Func_In = 8 Then
    For R In (Select b.Index_Name, a.���
              From User_Indexes B, zlBakTables A
              Where a.ϵͳ = System_In And a.���� = b.Table_Name And b.Status = 'VALID' And b.Index_Type = 'NORMAL' And
                    b.Index_Name Not Like 'BIN$%' And
                    b.Index_Name Not In (Select Upper(Column_Value)
                                         From Table(f_Str2list(v_Indexeswithtag))
                                         Union
                                         Select Upper(Column_Value)
                                         From Table(f_Str2list(v_Constraintswithtag)))
              Order By Index_Name) Loop
      Update Zldatamovelog
      Set ��ǰ���� = '�����ؽ�����:' || r.Index_Name
      Where ϵͳ = System_In And ���� = (Select Max(����) From Zldatamovelog Where ϵͳ = System_In);
    
      If Speedmode_In = 1 Then
        v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Nologging';
      Else
        v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Online Nologging';
      End If;
      Begin
        Execute Immediate v_Sql;
        --�����ؽ��Ƚ������������ؽ�����Ҫ���������������������������ORA-00054: ��Դ��æ, ��ָ���� NOWAIT ��ʽ��ȡ��Դ    
      Exception
        When Others Then
          If SQLErrM Like 'ORA-00054%' Then
            v_Sql := Replace(v_Sql, 'Rebuild', 'Rebuild Online');
            Execute Immediate v_Sql;
          End If;
      End;
    End Loop;
  End If;

  --ִ���ؽ���������Զ�Ϊ�������ϲ��ж����ԣ������ȡ������Ӱ�����SQL��ִ�мƻ�(ȫ��ɨ��+���в�ѯ������) 
  --------------------------------------------------------------------------------------------------- 
  If n_Parallel = 1 Then
    Execute Immediate 'ALTER Session DISABLE PARALLEL DDL';
  
    For R In (Select Index_Name From User_Indexes Where Degree Not In ('1', '0')) Loop
      v_Sql := 'Alter Index ' || r.Index_Name || ' Noparallel';
      Execute Immediate v_Sql;
    End Loop;
  End If;

  Update Zldatamovelog
  Set ��ǰ���� = '�ؽ����'
  Where ϵͳ = System_In And ���� = (Select Max(����) From Zldatamovelog Where ϵͳ = System_In);
  Commit;
  --�����̲����д����������ɵ��ù��̴��� 
End Zl1_Datamove_Reb;
/

Create Or Replace Procedure Zl1_Datamove_Tag
(
  d_End    In Date,
  n_����   In Number,
  n_System In Number
) As
  --���ܣ���Ǵ�ת�������� 
  --˵����Ϊ����Undo��ռ����͹��󣬷ֶ��ύ 
Begin
  --1.���ú��㣨����,ҩƷ,�տ��Ʊ�ݵȣ�  
  --�¼��Ӳ�ѯע�������Ż������ܹ������ݹ��˵���С�������ŵ����Exists��������ǰ��
  Update /*+ rule*/ ����Ԥ����¼ L
  Set ��ת�� = n_����
  Where ����id In
        (Select Distinct a.����id --1.�����շѺ͹Һŵ��շѽ����¼(�ų�֮���˺ź��˷ѵ�,һ�ŵ�����ֻҪ����һ������) 
         From ������ü�¼ A
         Where (a.��¼״̬ In (1, 2) Or a.��¼״̬ = 3 And Not Exists
                (Select 1
                 From ������ü�¼ B
                 Where a.No = b.No And a.��¼���� = b.��¼���� And b.��¼״̬ = 2 And b.�Ǽ�ʱ�� >= d_End))
     And a.��ת�� Is Null And a.��¼���� In (1, 4) And a.�Ǽ�ʱ�� < d_End
         Union All
         Select Distinct a.����id --2.ҽ�������� 
         From ���ò����¼ A
         Where (a.��¼״̬ In (1, 2) Or a.��¼״̬ = 3 And Not Exists
                (Select 1
                 From ���ò����¼ B
                 Where a.No = b.No And a.��¼���� = b.��¼���� And b.��¼״̬ In (1, 2) And b.�Ǽ�ʱ�� >= d_End))
     And a.��ת�� Is Null And a.��¼���� = 1 And a.�Ǽ�ʱ�� < d_End
         Union All
         Select Distinct a.����id --3.���￨���շѽ����¼(�ų�֮���˿��ѵ�,һ�ŵ�����ֻҪ����һ������) 
         From סԺ���ü�¼ A
         Where (a.��¼״̬ In (1, 2) Or a.��¼״̬ = 3 And Not Exists
                (Select 1
                 From סԺ���ü�¼ B
                 Where a.No = b.No And a.��¼���� = b.��¼���� And b.��¼״̬ = 2 And b.�Ǽ�ʱ�� >= d_End))
    And a.��ת�� Is Null And a.���ʷ��� = 0 And a.��¼���� = 5 And a.�Ǽ�ʱ�� < d_End
         Union All --4.����(���ʵ�)��סԺ�Ľ��ʽ����¼ 
         Select ����id
         From (With Settle As (Select Distinct a.Id As ����id, a.����id --3.����(���ʵ�)��סԺ�Ľ��ʽ����¼(�ų�֮��������ϵ�) 
                               From ���˽��ʼ�¼ A
                               Where (a.��¼״̬ In (1, 2) Or a.��¼״̬ = 3 And Not Exists
                                      (Select 1 From ���˽��ʼ�¼ B Where a.No = b.No And b.��¼״̬ = 2 And b.�շ�ʱ�� >= d_End))
              And a.��ת�� Is Null And a.�շ�ʱ�� < d_End)
                Select ����id
                From Settle
                Minus
                --���½���IDҪ�����ų�,���ⲿ�ַ�����ϸ��ת����Ӱ������ļ����Ƿ���� 
                --1.һ��Ԥ�����ʽ��ʳ��꣨����ID��ͬ��
                --2.���õ��ݵĽ���ID��صĿ��ܻ�������NO����������ID(�������Ϻ�ֶ�ν��ʽ��壬���ܲ�����ת��ʱ��֮��)
                --���ǵ�������ĸ����ԣ�Ϊ���߼���������ѯ���ܣ�������ID���ų� 
                Select Distinct d.Id
                From ���˽��ʼ�¼ D,
                     (Select Distinct c.����id --���סԺ����һ��ᣬ�Լ�������ʺ�סԺ���ʿ���һ����ҳ�ͬһ��Ԥ�����������ﲻ����ҳID 
                       From סԺ���ü�¼ C,
                            (Select Distinct d.No, d.���, Mod(d.��¼����, 10) As ��¼����
                              From סԺ���ü�¼ D,
                                   (Select s.����id From Settle S, ���˽��ʼ�¼ E --û�н����Ҹò���֮��û���ٽ���ͳ��˴��ʣ����־Ͳ��ų� 
                                     Where s.����id = e.����id And (e.�շ�ʱ�� > d_End Or Exists (Select 1 From ��Ժ���� F Where s.����id = f.����id))) S 
                              Where d.����id = s.����id) D
                       Where c.No = d.No And Mod(c.��¼����, 10) = d.��¼���� And c.��� = d.��� --���ʺ����Ϻ��ٶ԰������ʵ����ʵĽ���IDΪ�յļ�¼,һ����ܼ����Ƿ����,���ֽ���IDΪ�յ�����ת���ں��浥��ת�� 
                       Group By c.No, Mod(c.��¼����, 10), c.����id --һ�ŵ����е�һ�пɲ��ֽ��ʣ��Ե���Ϊ�������жϣ�����һ�ŵ��ݵ�����һ���ֱ�ת�� 
                       Having Nvl(Sum(c.ʵ�ս��), 0) <> Nvl(Sum(c.���ʽ��), 0) Or Exists (Select 1 --�ų�ת��ʱ��֮���ٴν��ʵ�(���Ϻ��ٴν���)������ԭʼ����ת�ߺ󣬺�������ʱ�޷���ȷ�ж� 
                                                                                   From סԺ���ü�¼ E, ���˽��ʼ�¼ S
                                                                                   Where e.No = c.No And Mod(e.��¼����, 10) = Mod(c.��¼����, 10) And
                                                                                         e.��¼���� In (12, 13, 15) And e.����id = s.Id  And s.��ת�� Is Null And s.�շ�ʱ�� >= d_End)
                       Union All
                       Select Distinct c.����id
                       From ������ü�¼ C,
                            (Select Distinct d.No, d.���, Mod(d.��¼����, 10) As ��¼����
                              From ������ü�¼ D, Settle S
                              Where d.����id = s.����id) D --��Ϊ�����ﲡ�ˣ����ԣ�ֻҪû�н���,�ò��˵Ķ���ת�� 
                       Where c.No = d.No And Mod(c.��¼����, 10) = d.��¼���� And c.��� = d.���
                       Group By c.No, Mod(c.��¼����, 10), c.����id
                       Having Nvl(Sum(c.ʵ�ս��), 0) <> Nvl(Sum(c.���ʽ��), 0) Or Exists (Select 1
                                                                                   From ������ü�¼ E, ���˽��ʼ�¼ S
                                                                                   Where e.No = c.No And Mod(e.��¼����, 10) = Mod(c.��¼����, 10) And
                                                                                         e.��¼���� In (12, 13, 15) And e.����id = s.Id And s.��ת�� Is Null And s.�շ�ʱ�� >= d_End)) N
                Where d.����id = n.����id)
         );

  --�ų�Ԥ����δ�����
  --Ϊ�˽����߼��ĸ����ԣ����ų���ת��ʱ��֮��ҩ��δ��ҩ�ķ��ü�¼��Ӧ�Ľ���ID������������Ľ������ݺͷ�������ǿ��ת�� 
  --��Ϊǰ���SQL����Ľ���ID���ܲ�ȫ�ǳ�Ԥ����(�����շѺ�סԺ���ʲ��ѵ�)�����ԣ���Ҫ����һ��SQL���ų� 
  --���ڿ��ܴ��������쳣(סԺ���ý��ʳ�Ԥ�����Ϊ1������Ԥ��)������û�м�Ԥ����������޶� 
  Update /*+ rule*/ ����Ԥ����¼
  Set ��ת�� = Null
  Where ��ת�� = n_���� And
        ����id In (Select Distinct d.����id
                 From ����Ԥ����¼ D,
                      --����D����Ϊ�˲��ͬһԤ�����ݵ���������ID����Ԥ�����Ԥ�����ϵģ��ٴγ�ͬһԤ�����ݣ� 
                      --��Ԥ�����Ԥ�������漰�����н���ID�Ķ���ת�������ⲿ�ֳ�Ԥ���Ľ���ID���ų���ԭʼԤ������ת�ߣ�������������ID�����õ��ݵ�һ����(ԭʼ���ʡ��������ϡ��ٴν�һ���֡��ٴν�ȫ��)ת�� 
                      (Select Distinct l.No
                        From ����Ԥ����¼ L, ����Ԥ����¼ P --���ܱ��ν��ʳ��ֻ��ʣ��������Ҫ����L����ԭʼ��Ԥ���ĵ��ݣ��Լ���¼����Ϊ11�Ŀ��ܻ���ת��ʱ��֮��������ʣ���Ľ���ID 
                        Where l.��¼���� = p.��¼���� And l.No = p.No And p.��¼���� In (1, 11) And p.��ת�� = n_����
                        Group By l.No, l.����id
                        Having Nvl(Sum(l.���), 0) <> Nvl(Sum(l.��Ԥ��), 0) And (Exists (Select 1
                                                                                  From ����Ԥ����¼ E --û�г�����֮��û���ٳ���������ͳ��˴��ʣ��Լ������ø��Ľ��ʲ�������ʾ��Ԥ�����ɳ��������������־Ͳ��ų�
                                                                                  Where l.����id = e.����id And e.��ת�� Is Null And e.�տ�ʱ�� > d_End)
                                                                                  Or Exists (Select 1 From ��Ժ���� E Where l.����id =e.����id)
                                                                                  Or Exists (Select 1 From ����δ����� E Where l.����id =e.����id))  
                        Or Nvl(Sum(l.���), 0) = Nvl(Sum(l.��Ԥ��), 0) And Exists (Select 1
                                                                                  From ����Ԥ����¼ E --�ų�ת��ʱ��֮�����������ID���,10.34.20�󣬳�Ԥ��ȫ������������һ����¼���շ�ʱ����ǳ�Ԥ��ʱ��(��ǰ����ԭʼ��Ԥ����ļ�¼�����Ԥ���ֶΣ�����ֱ�Ӳ鵽��Ԥ�����ʱ��)
                                                                                  Where e.No = l.No And e.��¼���� = 11 And e.��ת�� Is Null And e.�տ�ʱ�� >= d_End)) N
                 Where d.No = n.No And d.��¼���� In (1, 11));

  --Ԥ����û��ʹ�þ�ֱ�����˵ļ�¼(����IDΪ��) 
  Update /*+ rule*/ ����Ԥ����¼
  Set ��ת�� = n_����
  Where ��¼���� = 1 And
        NO In (Select a.No
               From ����Ԥ����¼ A
               Where a.����id Is Null And a.��¼���� = 1 And a.��¼״̬ In (2, 3) And a.��ת�� Is Null And a.�տ�ʱ�� < d_End
               Group By a.No
               Having Sum(a.���) = 0);

  --��Ԥ�������ϵļ�¼����¼����Ϊ2����û�н���ID 
  Update /*+ rule*/ ����Ԥ����¼
  Set ��ת�� = n_����
  Where ����id Is Null And ��¼���� = 2 And NO In (Select a.No From ����Ԥ����¼ A Where a.��ת�� = n_���� And a.��¼���� = 3);

  Update Zldatamovelog
  Set ��ǰ���� = '(1/10)�������ݱ����ɣ����ڱ�Ƿ�������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  Update /*+ rule*/ ���˽��ʼ�¼
  Set ��ת�� = n_����
  Where ID In (Select ����id From ����Ԥ����¼ Where ��ת�� = n_����);

  --�����޽���ļ�¼(Ϊ���������ܣ����жϷ��ã�ֻҪ����������Ԥ����¼�͵���������ý���) 
  Update /*+ rule*/ ���˽��ʼ�¼ L
  Set ��ת�� = n_����
  Where �շ�ʱ�� < d_End And ��ת�� Is Null And Not Exists (Select 1 From ����Ԥ����¼ P Where l.Id = p.����id);

  Update /*+ rule*/ ���˿��������
  Set ��ת�� = n_����
  Where Ԥ��id In (Select a.Id From ����Ԥ����¼ A Where ��ת�� = n_����);

  Update /*+ rule*/ ���˿������¼
  Set ��ת�� = n_����
  Where ID In (Select ������id From ���˿�������� Where ��ת�� = n_����);

  Update /*+ rule*/ �������㽻��
  Set ��ת�� = n_����
  Where ����id In (Select a.Id From ����Ԥ����¼ A Where ��ת�� = n_����);

  Update /*+ rule*/ �����˿���Ϣ
  Set ��ת�� = n_����
  Where (��¼id,����ID) In (Select a.Id,A.����ID From ����Ԥ����¼ A Where ��ת�� = n_����);

  --1.�ҺŴ��ۺ�ʵ�ս��Ϊ0��(û�ж�Ӧ��Ԥ����¼),��ʹ֮�����˺ŷ���Ҳ���ܣ���Ϊ���Ϊ�㲻Ӱ�����),�����Ѽ�ʹΪ��Ҳ��Ԥ����¼ 
  --����IDΪ�յ����쳣���ݣ�����ҽԺ����3�ʴ������ݣ�
  --���ݹҺż�¼����������ã���ֱ�Ӱ�ʱ����������Ҫ�� 
  Update /*+ rule*/ ������ü�¼
  Set ��ת�� = n_����
  Where NO In (Select NO From ���˹Һż�¼ Where ��ת�� Is Null And �Ǽ�ʱ�� < d_End) And ��¼���� = 4 And (ʵ�ս�� = 0 Or ����id Is Null);

  --2.ֱ���շѵĺͽ����޽��㣨Ԥ������¼�ģ�Union����allȥ���ظ��Լ���in������ 
  Update /*+ rule*/ ������ü�¼
  Set ��ת�� = n_����
  Where ����id In
        (Select ����id From ����Ԥ����¼ Where ��ת�� = n_���� Union Select ID From ���˽��ʼ�¼ Where ��ת�� = n_����);

  --3.û�н���id������(���Ǽ�ʱ��)
  --1)δ���ʵ�������ʷ���(����)���ò���û��Ԥ����¼���Ԥ����¼�����Ҹ�ʱ��֮����������÷���
  --2)δ���ʵĻ��ۼ�¼
  --3)δ�շѣ�Ҳû�г�Ԥ�����������
  --������"��ת�� Is Null"��Ϊ�˴���������α��ת������� 
  Update /*+ rule*/ ������ü�¼ A
  Set ��ת�� = n_����
  Where (Not Exists (Select 1 From ����Ԥ����¼ B Where a.����id = b.����id And b.��ת�� Is Null And ��¼���� In (1, 11)) And Not Exists
         (Select 1 From ������ü�¼ B Where a.����id = b.����id And b.��ת�� Is Null And �Ǽ�ʱ�� > d_End) And ��¼���� = 2 Or ��¼״̬ = 0 Or
         ��¼���� = 1 And ʵ�ս�� = 0 And ���ʽ�� = 0) And ����id Is Null And ��ת�� Is Null And �Ǽ�ʱ�� < d_End;

  --4.û�н���id������(������ʱ��)
  --���������ļ��ʼ�¼����¼״̬Ϊ2�����Ǽ�ʱ������ڵ�ǰָ��ת��ʱ��֮�󣬶�ԭʼ���ʼ�¼����¼״̬Ϊ3�����Ǽ�ʱ����ָ��ת��ʱ��֮ǰ��ǰ�����ߵķ���ʱ������ͬ�ġ�
  --1)δ���ʵ�����ʷ��û���ۺ�ʵ�ս��Ϊ��ģ�����ģ�����û�й�ѡ������ý��ʣ�
  --2)�������Ϻ󣬼��ʵ����ʵļ�¼������IDΪ���Ҽ�¼״̬Ϊ2�ģ�����¼״̬Ϊ3�����н���ID������ǰ����ת��. 
  Update /*+ rule*/ ������ü�¼ A
  Set ��ת�� = n_����
  Where (Exists (Select 1
                 From ������ü�¼ B
                 Where a.No = b.No And a.��¼���� = b.��¼���� And a.��� = b.��� And b.��¼״̬ = 3 And b.����id Is Not Null And
                       b.��ת�� + 0 = n_����) And ��¼״̬ = 2 Or Exists
         (Select 1
          From ������ü�¼ B
          Where a.No = b.No And a.��¼���� = b.��¼���� And a.��� = b.��� And b.����id Is Null
          Group By b.No, b.��¼����, b.���
          Having Nvl(Sum(b.ʵ�ս��), 0) = 0)) And ��¼���� = 2 And ����id Is Null And ��ת�� Is Null And ����ʱ�� < d_End;

  --5.�н���id�������(������ʱ��)
  --���ѱ���ۺ���ʽ��Ϊ����շѼ�¼,����һ�ŵ�����ͬ����ID�Ľ��ʽ��֮��Ϊ0(������Ϊ��)
  --��ʹ��ת��ʱ��֮��ҩ�ģ�Ҳǿ��ת����Ϊ�˼����߼������ԣ���߲�ѯ���ܣ�
  Update /*+ rule*/ ������ü�¼ A
  Set ��ת�� = n_����
  Where (���ʽ�� = 0 Or Exists
         (Select 1 From ������ü�¼ C Where a.����id = c.����id Group By c.����id, c.No Having Sum(c.���ʽ��) = 0)) And Not Exists
   (Select 1 From ����Ԥ����¼ B Where a.����id = b.����id And b.��ת�� Is Null) And ��¼���� = 1 And ����id Is Not Null And
        ��ת�� Is Null And ����ʱ�� < d_End;

  Update /*+ rule*/ ҽ��������ϸ
  Set ��ת�� = n_����
  Where ����id In (Select ����id From ����Ԥ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ���ò����¼
  Set ��ת�� = n_����
  Where ����id In (Select ����id From ����Ԥ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ƾ����ӡ��¼
  Set ��ת�� = n_����
  Where (NO, ��¼����) In (Select NO, ��¼���� From ������ü�¼ Where ��ת�� = n_����);

   --1.��Ԥ����¼����Ϊ��ȡ���￨ֱ���շѵģ��޽���ID��,�ټӽ��ʼ�¼��Ϊ��ȡ�����޽��㣨Ԥ������¼�� 
  Update /*+ rule*/ סԺ���ü�¼
  Set ��ת�� = n_����
  Where ����id In
        (Select ����id From ����Ԥ����¼ Where ��ת�� = n_���� Union Select ID From ���˽��ʼ�¼ Where ��ת�� = n_����);

  --2.û�н���id������(������ʱ��)
  --���������ļ��ʼ�¼����¼״̬Ϊ2�����Ǽ�ʱ������ڵ�ǰָ��ת��ʱ��֮�󣬶�ԭʼ���ʼ�¼����¼״̬Ϊ3�����Ǽ�ʱ����ָ��ת��ʱ��֮ǰ��ǰ�����ߵķ���ʱ������ͬ�ġ�
  --1)ת���������Ϻ󣬼��ʵ����ʵļ�¼������״̬Ϊ2��û�н���ID��(��¼״̬Ϊ3���н���ID��)����ǰ����ת���� 
  --2)δ���ʵ������(�ѳ����ļ��ʵ�����ۺ�ʵ�ս��Ϊ��) 
  --3)û�н���ID�Ļ��ۼ�¼����Ϊת�� 
  
  Update /*+ rule*/ סԺ���ü�¼ A
  Set ��ת�� = n_����
  Where ((Exists (Select 1
                  From סԺ���ü�¼ B
                  Where a.No = b.No And a.��¼���� = b.��¼���� And a.��� = b.��� And b.��¼״̬ = 3 And b.����id Is Not Null And
                        b.��ת�� + 0 = n_����) And ��¼״̬ = 2 Or Exists
         (Select 1
           From סԺ���ü�¼ B
           Where a.No = b.No And a.��¼���� = b.��¼���� And a.��� = b.��� And b.����id Is Null
           Group By b.No, b.��¼����, b.���
           Having Nvl(Sum(b.ʵ�ս��), 0) = 0)) And ��¼���� = 2 Or ��¼״̬ = 0) And ����id Is Null And ��ת�� Is Null And ����ʱ�� < d_End;

  --3.��Ժδ���ʵģ����ʲ��ˣ�����Ϊ�Ǻܾ���ǰ����Щ���ݣ����Ԥ���ѳ��꣬����ΪҪת�� 
  Update /*+ rule*/ סԺ���ü�¼ A
  Set ��ת�� = n_����
  Where ��ת�� Is Null And ����id Is Null And
        (����id, ��ҳid) In (Select ����id, ��ҳid
                         From ������ҳ C
                         Where ��Ժ���� < d_End And ��ת�� Is Null And ����ת�� Is Null And Not Exists
                          (Select 1
                                From ����Ԥ����¼ B
                                Where b.����id = c.����id And b.��ת�� Is Null And b.Ԥ����� = 2 And b.��¼���� In (1, 11)
                                Having Nvl(Sum(b.���), 0) - Nvl(Sum(b.��Ԥ��), 0) <> 0));

  Update /*+ rule*/ �����嵥��ӡ
  Set ��ת�� = n_����
  Where (NO, Mod(��¼����,10),Decode(��¼״̬,3,1,��¼״̬),���) In 
        (Select NO, Mod(��¼����,10) as ��¼����,Decode(��¼״̬,3,1,��¼״̬) as ��¼״̬,��� From ������ü�¼ Where ��ת�� = n_����
        Union
        Select NO, Mod(��¼����,10) as ��¼����,Decode(��¼״̬,3,1,��¼״̬) as ��¼״̬,��� From סԺ���ü�¼ Where ��ת�� = n_����);

  Update Zldatamovelog
  Set ��ǰ���� = '(2/10)�������ݱ����ɣ����ڱ��ҩƷ����'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  Update /*+ Rule*/ ҩƷ�շ���¼ A
  Set ��ת�� = n_����
  Where Rowid In (Select m.Rowid
                  From ҩƷ�շ���¼ M, ������ü�¼ E
                  Where m.����id = e.Id And (e.��¼���� = 1 And m.���� In (8, 24) Or e.��¼���� = 2 And m.���� In (9, 25)) And
                        e.�շ���� In ('4', '5', '6', '7') And e.��ת�� = n_����
                  Union All
                  Select m.Rowid
                  From ҩƷ�շ���¼ M, סԺ���ü�¼ E
                  Where m.����id = e.Id And m.���� In (9, 10, 25, 26) And e.��¼���� = 2 And e.�շ���� In ('4', '5', '6', '7') And
                        e.��ת�� = n_����);

  Update /*+ rule*/ �շ���¼������Ϣ
  Set ��ת�� = n_����
  Where �շ�id In (Select ID From ҩƷ�շ���¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ��Һ��ҩ����
  Set ��ת�� = n_����
  Where �շ�id In (Select ID From ҩƷ�շ���¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ��Һ��ҩ��¼
  Set ��ת�� = n_����
  Where ID In (Select ��¼id From ��Һ��ҩ���� Where ��ת�� = n_����);

  Update /*+ rule*/ ��Һ��ҩ����
  Set ��ת�� = n_����
  Where ��ҩid In (Select ID From ��Һ��ҩ��¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ��Һ��ҩ״̬
  Set ��ת�� = n_����
  Where ��ҩid In (Select ID From ��Һ��ҩ��¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ҩƷ����ƻ�
  Set ��ת�� = n_����
  Where ����id In (Select ID From ҩƷ�շ���¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ҩƷǩ����ϸ
  Set ��ת�� = n_����
  Where �շ�id In (Select ID From ҩƷ�շ���¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ҩƷǩ����¼
  Set ��ת�� = n_����
  Where ID In (Select ǩ��id From ҩƷǩ����ϸ Where ��ת�� = n_����);

  Update /*+ rule*/ ҩƷ�շ������־ A
  Set ��ת�� = n_����
  Where Exists(Select 1 From ҩƷ�շ���¼ B Where b.NO = a.������ And b.���� = a.���� And b.��ת�� = n_����);

  Update /*+ rule*/ ҩƷ�շ�סԺ��־
  Set ��ת�� = n_����
  Where �շ�id In (Select ID From ҩƷ�շ���¼ Where ��ת�� = n_����);

  Update Zldatamovelog
  Set ��ǰ���� = '(3/10)ҩƷ���ݱ����ɣ����ڱ�ǽɿ���Ʊ������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  Update /*+ rule*/ ��Ա����¼ Set ��ת�� = n_���� Where ��ת�� Is Null And ���ʱ�� < d_End;

  Update /*+ rule*/ ��Ա�սɼ�¼ Set ��ת�� = n_���� Where ��ת�� Is Null And �Ǽ�ʱ�� < d_End;

  Update /*+ rule*/ ��Ա�սɶ���
  Set ��ת�� = n_����
  Where �ս�id In (Select ID From ��Ա�սɼ�¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ��Ա�ս���ϸ
  Set ��ת�� = n_����
  Where �ս�id In (Select ID From ��Ա�սɼ�¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ��Ա�ս�Ʊ��
  Set ��ת�� = n_����
  Where �ս�id In (Select ID From ��Ա�սɼ�¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ��Ա�ݴ��¼
  Set ��ת�� = n_����
  Where �ս�id In (Select ID From ��Ա�սɼ�¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ��Ա�ݴ��¼ Set ��ת�� = n_���� Where ��ת�� Is Null And ��¼���� = 1 And �Ǽ�ʱ�� < d_End;

  Update /*+ rule*/ Ʊ�����ü�¼ A
  Set ��ת�� = n_����
  Where Not Exists
   (Select 1 From Ʊ��ʹ����ϸ B Where b.����id = a.Id And b.ʹ��ʱ�� >= d_End) And ��ת�� Is Null And ʣ������ = 0 And �Ǽ�ʱ�� < d_End;

  Update /*+ rule*/ Ʊ��ʹ����ϸ
  Set ��ת�� = n_����
  Where ����id In (Select ID From Ʊ�����ü�¼ Where ��ת�� = n_����);
  Update /*+ rule*/ Ʊ�ݴ�ӡ����
  Set ��ת�� = n_����
  Where ID In (Select ��ӡid From Ʊ��ʹ����ϸ Where ��ת�� = n_����);

  Update /*+ rule*/ Ʊ�ݴ�ӡ��ϸ
  Set ��ת�� = n_����
  Where ʹ��id In (Select ID From Ʊ��ʹ����ϸ Where ��ת�� = n_����);

  Update Zldatamovelog
  Set ��ǰ���� = '(4/10)�ɿ���Ʊ�����ݱ����ɣ����ڱ�Ǿ��Ｐ��������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  --2.���Ｐ�������� 
  --��ת�����������Һŷ���δת���ģ�ת��ʱ��֮�����ҽ����ҽ����Ӧ�ķ���δת���� 
  --��ʹ���ھ���(r.ִ��״̬ <> 2 )��Ҳǿ��ת�� 
  Update /*+ rule*/ ���˹Һż�¼ T
  Set ��ת�� = n_����
  Where Rowid In
        (Select Rowid
         From ���˹Һż�¼ R
         Where Not Exists (Select 1
                From ������ü�¼ A
                Where r.No = a.No And a.�Ǽ�ʱ�� < d_End And a.��¼���� = 4 And a.��ת�� Is Null) And Not Exists
          (Select 1
                From ����ҽ����¼ A
                Where a.�Һŵ� = r.No And a.��ת�� Is Null And a.������Դ <> 4 And Nvl(a.ͣ��ʱ��, a.����ʱ��) >= d_End) And Not Exists
          (Select 1
                From ������ü�¼ E, ����ҽ����¼ A
                Where r.No = a.�Һŵ� And a.Id = e.ҽ����� And a.������Դ <> 4 And e.��ת�� Is Null) And r.��ת�� Is Null And
               r.�Ǽ�ʱ�� < d_End);

  --������һ���ֹҺ�����δת�������ԣ����ܱ�����ݿ�����Һ����ݲ�ƥ�� 
  Update ���˹ҺŻ��� Set ��ת�� = n_���� Where ��ת�� Is Null And ���� < d_End;
  Update /*+ rule*/ ����ת���¼ Set ��ת�� = n_���� Where NO In (Select NO From ���˹Һż�¼ Where ��ת�� = n_����);

  --ͨ��"סԺ���ü�¼"����ѯ��������"���˽��ʼ�¼",��Ϊ��Ժδ������ʲ���Ҳת���˷��� 
  --��Ժ����������Ȼ��Ҫ����Ϊ����ĳ�ν���ת���ˣ������˵�ʱ��δ��Ժ(һ��סԺ��ν���)�� 
  --ͨ��ָ��������ʽ���������Ż���ȱʡ����"������ҳIX_��Ժ����"������Ч��̫�ͣ� 
  Update /*+ rule*/ ������ҳ P
  Set ��ת�� = n_����
  Where Not Exists (Select 1 From סԺ���ü�¼ A Where a.����id = p.����id And a.��ҳid = p.��ҳid And a.��ת�� Is Null) And ��ת�� Is Null And
        ����ת�� Is Null And ��Ժ���� < d_End And
        (����id, ��ҳid) In (Select Distinct ����id, ��ҳid From סԺ���ü�¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ���˹�����¼
  Set ��ת�� = n_����
  Where (����id, ��ҳid) In (Select ����id, ID
                         From ���˹Һż�¼
                         Where ��ת�� = n_����
                         Union All
                         Select ����id, ��ҳid
                         From ������ҳ
                         Where ��ת�� = n_����);

  Update /*+ rule*/ ������ϼ�¼
  Set ��ת�� = n_����
  Where (����id, ��ҳid) In (Select ����id, ID
                         From ���˹Һż�¼
                         Where ��ת�� = n_����
                         Union All
                         Select ����id, ��ҳid
                         From ������ҳ
                         Where ��ת�� = n_����);

  Update /*+ rule*/ ���������¼
  Set ��ת�� = n_����
  Where (����id, ��ҳid) In (Select ����id, ID
                         From ���˹Һż�¼
                         Where ��ת�� = n_����
                         Union All
                         Select ����id, ��ҳid
                         From ������ҳ
                         Where ��ת�� = n_����);

  Update Zldatamovelog
  Set ��ǰ���� = '(5/10)���Ｐ�������ݱ����ɣ����ڱ�ǻ�������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  --3.�������� 
  Update /*+ rule*/ ���˻����ļ�
  Set ��ת�� = n_����
  Where (����id, ��ҳid) In (Select ����id, ��ҳid From ������ҳ Where ��ת�� = n_����);
  Update /*+ rule*/ ���˻�������
  Set ��ת�� = n_����
  Where �ļ�id In (Select ID From ���˻����ļ� Where ��ת�� = n_����);
  Update /*+ rule*/ ���˻�����ϸ
  Set ��ת�� = n_����
  Where ��¼id In (Select ID From ���˻������� Where ��ת�� = n_����);

  Update /*+ rule*/ ���˻����ӡ
  Set ��ת�� = n_����
  Where �ļ�id In (Select ID From ���˻����ļ� Where ��ת�� = n_����);
  Update /*+ rule*/ ���˻�����Ŀ
  Set ��ת�� = n_����
  Where �ļ�id In (Select ID From ���˻����ļ� Where ��ת�� = n_����);
  Update /*+ rule*/ ���˻���Ҫ������
  Set ��ת�� = n_����
  Where �ļ�id In (Select ID From ���˻����ļ� Where ��ת�� = n_����);
  Update /*+ rule*/ ����Ҫ������
  Set ��ת�� = n_����
  Where �ļ�id In (Select ID From ���˻����ļ� Where ��ת�� = n_����);

  --�ϰ滤��ϵͳ���� 
  Update /*+ rule*/ ���˻����¼
  Set ��ת�� = n_����
  Where (����id, ��ҳid) In (Select ����id, ��ҳid From ������ҳ Where ��ת�� = n_����);
  Update /*+ rule*/ ���˻�������
  Set ��ת�� = n_����
  Where ��¼id In (Select ID From ���˻����¼ Where ��ת�� = n_����);

  Update Zldatamovelog
  Set ��ǰ���� = '(6/10)�������ݱ����ɣ����ڱ�ǲ�������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  --4.�������� 
  Update /*+ rule*/ ���Ӳ�����¼
  Set ��ת�� = n_����
  Where ������Դ <> 4 And (����id, ��ҳid) In (Select ����id, ID
                                       From ���˹Һż�¼
                                       Where ��ת�� = n_����
                                       Union All
                                       Select ����id, ��ҳid
                                       From ������ҳ
                                       Where ��ת�� = n_����);

  --�ԵǼ��ಡ��(�޹Һŵ���) 
  --����ID�����ظ�����Ϊ���鱨��֮��ģ���ι�����������һ�ű��棬���ڲ���ҽ��������У����ҽ��id��Ӧͬһ����ID 
  --Ϊ�������ܣ�����ҽ�����ͼ�¼�ķ���ʱ���ѯ�������þ�ȷ��ʱ�䣬��Ϊֱ�ӵǼǵļ���ҽ����һ�㿪��ʱ���뷢��ʱ������
  Update /*+ rule*/ ���Ӳ�����¼
  Set ��ת�� = N_����
  Where ID In (Select C.����id
             From ����ҽ����¼ B, ����ҽ������ C
             Where C.ҽ��id = B.Id And Nvl(B.��ҳid, 0) = 0 And B.�Һŵ� Is Null And B.���id Is Null And B.��ת�� Is Null And
                   B.����ʱ�� < d_End);

  Update /*+ rule*/ ���Ӳ�������
  Set ��ת�� = n_����
  Where ����id In (Select ID From ���Ӳ�����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ���Ӳ�����ʽ
  Set ��ת�� = n_����
  Where �ļ�id In (Select ID From ���Ӳ�����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ���Ӳ�������
  Set ��ת�� = n_����
  Where �ļ�id In (Select ID From ���Ӳ�����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ���Ӳ���ͼ��
  Set ��ת�� = n_����
  Where ����id In (Select ID From ���Ӳ������� Where ��ת�� = n_���� And �������� = 5);

  Update /*+ rule*/ ����ҽ������
  Set ��ת�� = n_����
  Where ����id In (Select ID From ���Ӳ�����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ Ӱ�񱨸沵��
  Set ��ת�� = n_����
  Where (ҽ��id, ����id) In (Select ҽ��id, ����id From ����ҽ������ Where ��ת�� = n_����);

  Update /*+ rule*/ ������ļ�¼
  Set ��ת�� = n_����
  Where ����id In (Select ID From ���Ӳ�����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ �����걨��¼
  Set ��ת�� = n_����
  Where �ļ�id In (Select ID From ���Ӳ�����¼ Where ��ת�� = n_����);
  
  Update /*+ rule*/ �������淴��
  Set ��ת�� = n_����
  Where �ļ�id In (Select ID From ���Ӳ�����¼ Where ��ת�� = n_����);
  
  Update /*+ rule*/ ������ϼ�¼
  Set ��ת�� = n_����
  Where ����id In (Select ID From ���Ӳ�����¼ Where ��ת�� = n_����);

  Update Zldatamovelog
  Set ��ǰ���� = '(7/10)�������ݱ����ɣ����ڱ���ٴ�·������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  --5.�ٴ�·�� 
  Update /*+ rule*/ �����ٴ�·��
  Set ��ת�� = n_����
  Where (����id, ��ҳid) In (Select ����id, ��ҳid From ������ҳ Where ��ת�� = n_����);
  Update /*+ rule*/ ���˺ϲ�·��
  Set ��ת�� = n_����
  Where ��Ҫ·����¼id In (Select ID From �����ٴ�·�� Where ��ת�� = n_����);
  Update /*+ rule*/ ���˺ϲ�·������
  Set ��ת�� = n_����
  Where ·����¼id In (Select ID From �����ٴ�·�� Where ��ת�� = n_����);

  Update /*+ rule*/ ���˳�����¼
  Set ��ת�� = n_����
  Where ·����¼id In (Select ID From �����ٴ�·�� Where ��ת�� = n_����);

  Update /*+ rule*/ ����·��ִ��
  Set ��ת�� = n_����
  Where ·����¼id In (Select ID From �����ٴ�·�� Where ��ת�� = n_����);
  Update /*+ rule*/ ����·������
  Set ��ת�� = n_����
  Where ·����¼id In (Select ID From �����ٴ�·�� Where ��ת�� = n_����);
  Update /*+ rule*/ ����·������
  Set ��ת�� = n_����
  Where ·����¼id In (Select ID From �����ٴ�·�� Where ��ת�� = n_����);
  Update /*+ rule*/ ����·��ָ��
  Set ��ת�� = n_����
  Where ·����¼id In (Select ID From �����ٴ�·�� Where ��ת�� = n_����);
  Update /*+ rule*/ ����·��ҽ��
  Set ��ת�� = n_����
  Where ·��ִ��id In (Select ID From ����·��ִ�� Where ��ת�� = n_����);

  Update Zldatamovelog
  Set ��ǰ���� = '(8/10)�ٴ�·�����ݱ����ɣ����ڱ��ҽ������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  --6.ҽ�������飬��� 
  Update /*+ rule*/ ����ҽ����¼
  Set ��ת�� = n_����
  Where �Һŵ� In (Select NO From ���˹Һż�¼ Where ��ת�� = n_����) And ������Դ <> 4;
  Update /*+ rule*/ ����ҽ����¼
  Set ��ת�� = n_����
  Where (����id, ��ҳid) In (Select ����id, ��ҳid From ������ҳ Where ��ת�� = n_����);

  --�ԵǼ��ಡ��(�޹Һŵ�)������ҽ��������ǰ��ת����ʱ��ת�� 
  Update /*+ rule*/ ����ҽ����¼
  Set ��ת�� = n_����
  Where Rowid In (Select b.Rowid
                  From ����ҽ����¼ B, ����ҽ������ C
                  Where (b.���id = c.ҽ��id Or b.Id = c.ҽ��id) And c.��ת�� = n_����);

  Update /*+ rule*/ ����ҽ���Ƽ�
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ����ҽ������
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ����ҽ������
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ��Ѫ�����¼
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ��Ѫ������
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ����ҽ��ִ��
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ����ҽ����ӡ
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ҽ��ִ�д�ӡ
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ �������ҽ��
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ������ϼ�¼
  Set ��ת�� = n_����
  Where ID In (Select ���id From �������ҽ�� Where ��ת�� = n_����);

  Update /*+ rule*/ ����ҽ��״̬
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ҽ��ǩ����¼
  Set ��ת�� = n_����
  Where ID In (Select ǩ��id From ����ҽ��״̬ Where ��ת�� = n_���� And ǩ��id Is Not Null);

  Update /*+ rule*/ ����ҽ������
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ���Ƶ��ݴ�ӡ
  Set ��ת�� = n_����
  Where (NO, ��¼����) In (Select NO, ��¼���� From ����ҽ������ Where ��ת�� = n_����);

  Update /*+ rule*/ ҽ��ִ��ʱ��
  Set ��ת�� = n_����
  Where (ҽ��id, ���ͺ�) In (Select ҽ��id, ���ͺ� From ����ҽ������ Where ��ת�� = n_����);

  Update /*+ rule*/ ҽ��ִ�мƼ�
  Set ��ת�� = n_����
  Where (ҽ��id, ���ͺ�) In (Select ҽ��id, ���ͺ� From ����ҽ������ Where ��ת�� = n_����);

  Update /*+ rule*/ ִ�д�ӡ��¼
  Set ��ת�� = n_����
  Where (ҽ��id, ���ͺ�) In (Select ҽ��id, ���ͺ� From ����ҽ������ Where ��ת�� = n_����);

  Update /*+ rule*/ ���������ϸ
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ��������¼
  Set ��ת�� = n_����
  Where ID In (Select ��id From ���������ϸ Where ��ת�� = n_����);

  Update /*+ rule*/ ���������
  Set ��ת�� = n_����
  Where ��id In (Select ID From ��������¼ Where ��ת�� = n_����);
  
  Update /*+ rule*/ RIS���ԤԼ
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  
  Update Zldatamovelog
  Set ��ǰ���� = '(9/10)ҽ�����ݱ����ɣ����ڱ�Ǽ���������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  Update /*+ rule*/ Ӱ�����¼
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ Ӱ�񱨸��¼
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ Ӱ�񱨸������¼
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ Ӱ��������
  Set ��ת�� = n_����
  Where ���uid In (Select ���uid From Ӱ�����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ Ӱ����ͼ��
  Set ��ת�� = n_����
  Where ����uid In (Select ����uid From Ӱ�������� Where ��ת�� = n_����);

  Update /*+ rule*/ Ӱ�����뵥ͼ��
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ Ӱ���ղ�����
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ Ӱ��Σ��ֵ��¼
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update Zldatamovelog
  Set ��ǰ���� = '(10/10)Ӱ�����ݱ����ɣ����ڱ�Ǽ�������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  Update /*+ rule*/ ����걾��¼
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ����������Ŀ
  Set ��ת�� = n_����
  Where �걾id In (Select ID From ����걾��¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ������Ŀ�ֲ�
  Set ��ת�� = n_����
  Where �걾id In (Select ID From ����걾��¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ���������¼
  Set ��ת�� = n_����
  Where �걾id In (Select ID From ����걾��¼ Where ��ת�� = n_����);
  Update /*+ rule*/ �����ʿؼ�¼
  Set ��ת�� = n_����
  Where �걾id In (Select ID From ����걾��¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ���������¼
  Set ��ת�� = n_����
  Where �걾id In (Select ID From ����걾��¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ����ǩ����¼
  Set ��ת�� = n_����
  Where ����걾id In (Select ID From ����걾��¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ����ͼ����
  Set ��ת�� = n_����
  Where �걾id In (Select ID From ����걾��¼ Where ��ת�� = n_����);

  Update /*+ rule*/ �����Լ���¼
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ������ռ�¼
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ������ͨ���
  Set ��ת�� = n_����
  Where ����걾id In (Select ID From ����걾��¼ Where ��ת�� = n_����);
  Update /*+ rule*/ �����ʿر���
  Set ��ת�� = n_����
  Where ���id In (Select ID From ������ͨ��� Where ��ת�� = n_����);
  Update /*+ rule*/ ����ҩ�����
  Set ��ת�� = n_����
  Where ϸ�����id In (Select ID From ������ͨ��� Where ��ת�� = n_����);
  Update /*+ rule*/ ������ˮ�߱걾
  Set ��ת�� = n_����
  Where �걾id In (Select ID From ����걾��¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ������ˮ��ָ��
  Set ��ת�� = n_����
  Where �걾id In (Select ID From ����걾��¼ Where ��ת�� = n_����);

  Commit;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl1_Datamove_Tag;
/
