----------------------------------------------------------------------------------------------------------------
--���ű�֧�ִ�ZLHIS+ v10.35.90������ v10.35.90
--�������ݿռ�������ߵ�¼PLSQL��ִ�����нű�
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------
--�ṹ��������
------------------------------------------------------------------------------
--137889:����,2019-03-18,����������Ա�ҩ�Ĵ���
create table ��Һ�Ա�ҩ�嵥
(
  ���     number(3) not null,
  ҩƷid   number(18) not null,
  �Ƿ����� number(1)
)
tablespace zl9BaseItem;
alter table ��Һ�Ա�ҩ�嵥 add constraint ��Һ�Ա�ҩ�嵥_PK_��� primary key (���);
alter table ��Һ�Ա�ҩ�嵥 add constraint ��Һ�Ա�ҩ�嵥_UQ_ҩƷid unique (ҩƷID);


------------------------------------------------------------------------------
--������������
------------------------------------------------------------------------------
--137889:����,2019-03-18,����������Ա�ҩ�Ĵ���
Insert into zlTables(ϵͳ,����,��ռ�,����) Values(100,'��Һ�Ա�ҩ�嵥','ZL9BASEITEM','A2');

-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------
--137889:����,2019-03-18,����������Ա�ҩ�Ĵ���
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1022,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0
Union All Select 'Zl_��Һ�Ա�ҩ�嵥_����','EXECUTE' From Dual
Union All Select '��Һ�Ա�ҩ�嵥','SELECT' From Dual) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1345,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0
Union All Select '��Һ�Ա�ҩ�嵥','SELECT' From Dual) A;

--137893:������,2019-03-18,��Һ�Ա�ҩ�嵥
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1254,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0
Union All Select '��Һ�Ա�ҩ�嵥','SELECT' From Dual) A;


-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------

--137889:����,2019-03-18,����������Ա�ҩ�Ĵ���
Create Or Replace Procedure Zl_��Һ�Ա�ҩ�嵥_����
(
  ���_In         In ��Һ�Ա�ҩ�嵥.���%Type,
  ҩƷid_In       In ��Һ�Ա�ҩ�嵥.ҩƷid%Type,
  �Ƿ�����_In In ��Һ�Ա�ҩ�嵥.�Ƿ�����%Type,
  n_First_In      In Number
) Is
Begin
  --����ǰ��ɾ��֮ǰ��
  If n_First_In = 1 Then
    Delete From ��Һ�Ա�ҩ�嵥;
  End If;

  --����[��Һ�Ա�ҩ�嵥]����
  If ҩƷid_In <> 0 Then
    Insert Into ��Һ�Ա�ҩ�嵥 (���, ҩƷid, �Ƿ�����) Values (���_In, ҩƷid_In, �Ƿ�����_In);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Һ�Ա�ҩ�嵥_����;
/

--137889:����,2019-03-18,����������Ա�ҩ�Ĵ���
Create Or Replace Procedure Zl_��Һ��ҩ��¼_�������
(
  ��ҩid_In   In Varchar2, --ID��:ID1,��˱�־1,ID2,��˱�־2....
  ������Ա_In In ��Һ��ҩ��¼.������Ա%Type,
  ����ʱ��_In In ��Һ��ҩ��¼.����ʱ��%Type
) Is
  v_Tansid     Varchar2(20);
  v_Tansids    Varchar2(4000);
  v_Tmp        Varchar2(4000);
  v_Usercode   Varchar2(100);
  v_��ҩid     ҩƷ�շ���¼.Id%Type;
  n_Count      Number(1);
  d_���ʱ��   ҩƷ�շ���¼.�������%Type;
  v_No         ҩƷ�շ���¼.No%Type;
  v_�ϴ�no     ҩƷ�շ���¼.No%Type;
  n_��˱�־   Number(1);
  n_����״̬   Number(2);
  v_�շ�ids    Varchar2(4000);
  v_��ҩ����id ҩƷ�շ���¼.Id%Type;

  v_ԭʼid     ҩƷ�շ���¼.Id%Type;
  v_Error      Varchar2(255);
  n_����         Number; --1�����ﵥ�ݣ�2��סԺ����
  Err_Custom Exception;

  Cursor c_���ʼ�¼ Is
    Select Distinct a.����id, b.����ʱ��
    From ҩƷ�շ���¼ A, ��Һ��ҩ��¼ B, ��Һ��ҩ���� C
    Where a.Id = c.�շ�id And b.Id = c.��¼id And b.Id = v_Tansid And b.����״̬ = 9;

  v_���ʼ�¼ c_���ʼ�¼%RowType;

  Cursor c_��ҩ��¼ Is
    Select /*+ rule*/
    Distinct a.Id As ��ҩid, c.�շ�id, c.����, a.ҩƷid, a.����, c.��¼id As ��ҩid
    From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B, ��Һ��ҩ���� C, Table(Cast(f_Num2list(v_Tansids) As Zltools.t_Numlist)) D
    Where c.��¼id = d.Column_Value And b.Id = c.�շ�id And a.���� = b.���� And a.No = b.No And a.�ⷿid + 0 = b.�ⷿid And
          a.ҩƷid + 0 = b.ҩƷid And a.��� = b.��� And (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0)
    Order By a.ҩƷid, a.����;

  v_��ҩ��¼ c_��ҩ��¼%RowType;

  Cursor c_�������� Is
    Select /*+ rule*/
     a.No, a.��� || ':' || c.���� || ':' || c.��¼id As �������
    From סԺ���ü�¼ A, ҩƷ�շ���¼ B, ��Һ��ҩ���� C, Table(Cast(f_Num2list(v_Tansids) As Zltools.t_Numlist)) D
    Where a.Id = b.����id And b.Id = c.�շ�id And Mod(b.��¼״̬, 3) = 1 And c.��¼id = d.Column_Value;

  v_�������� c_��������%RowType;

  Cursor c_�Ա�ҩ��¼ Is
    Select Distinct a.Id, b.��������, c.����ϵ��, c.ҩƷid
    From ��Һ��ҩ��¼ A, ����ҽ����¼ B, ҩƷ��� C, Table(Cast(f_Num2list(v_Tansids) As Zltools.t_Numlist)) D
    Where a.ҽ��id = b.���id And a.Id = d.Column_Value And b.�շ�ϸĿid = c.ҩƷid And b.ִ������ = 5 And b.ִ�б�� = 0 And
          b.�շ�ϸĿid In (Select d.ҩƷid From ��Һ�Ա�ҩ�嵥 D Where d.�Ƿ����� = 1);

  v_�Ա�ҩ��¼ c_�Ա�ҩ��¼%RowType;
Begin
  If ��ҩid_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := ��ҩid_In || ',';
  End If;

  v_Usercode := Zl_Identity;
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ';') + 1);
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ',') + 1);
  v_Usercode := Substr(v_Usercode, 1, Instr(v_Usercode, ',') - 1);

  While v_Tmp Is Not Null Loop
    --�ֽⵥ��ID��
    v_Tansid   := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp      := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
    n_��˱�־ := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp      := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
  
    v_�շ�ids := Null;
  
    --ͳ�����ȷ�ϵ���Һ��(n_��˱�־ = 1)
    If n_��˱�־ = 1 Then
      If v_Tansids Is Null Then
        v_Tansids := v_Tansid;
      Else
        v_Tansids := v_Tansids || ',' || v_Tansid;
      End If;
    End If;
  
    --��鵱ǰ��Һ����״̬�Ƿ�Ϊ����ҩ״̬
    Begin
      Select ����״̬ Into n_����״̬ From ��Һ��ҩ��¼ Where ID = v_Tansid;
    
      If n_����״̬ <> 9 Then
        v_Error := '�������ѱ����������ܽ���������ˣ�';
        Raise Err_Custom;
      End If;
    Exception
      When Others Then
        v_Error := '�������ѱ�������';
        Raise Err_Custom;
    End;
  
    If n_��˱�־ = 1 Then
      n_����״̬ := 10;
    Elsif n_��˱�־ = 2 Then
      n_����״̬ := 11;
    End If;
  
    --������Һ����Ӧ���շ�NO
    Begin
      Select NO
      Into v_No
      From ҩƷ�շ���¼
      Where ID In (Select �շ�id From ��Һ��ҩ���� Where ��¼id In (Select ID From ��Һ��ҩ��¼ Where ID = v_Tansid)) And Rownum < 2;
    Exception
      When Others Then
        v_No := '';
    End;
  
    --�շ�NO��ͬ����ҩID�����ʱ���Դ�����Ϊ�ӳ�1��
    If v_No = v_�ϴ�no Then
      d_���ʱ�� := d_���ʱ�� + 1 / 24 / 60 / 60;
    Else
      d_���ʱ�� := ����ʱ��_In;
      v_�ϴ�no   := v_No;
    End If;
  
    --���ʼ�¼����
    For v_���ʼ�¼ In c_���ʼ�¼ Loop
      Zl_���˷�������_Audit(v_���ʼ�¼.����id, v_���ʼ�¼.����ʱ��, ������Ա_In, d_���ʱ��, n_��˱�־);
    End Loop;
  
    Select Count(*) Into n_Count From ��Һ��ҩ״̬ Where ��ҩid = v_Tansid And ����ʱ�� = ����ʱ��_In;
  
    If n_Count <> 1 Then
      Insert Into ��Һ��ҩ״̬
        (��ҩid, ��������, ������Ա, ����ʱ��)
      Values
        (v_Tansid, n_����״̬, ������Ա_In, ����ʱ��_In);
    End If;
    Update ��Һ��ҩ��¼ Set ������Ա = ������Ա_In, ����ʱ�� = ����ʱ��_In, ����״̬ = n_����״̬ Where ID = v_Tansid;
  End Loop;

  --����ҩ
  For v_��ҩ��¼ In c_��ҩ��¼ Loop
    Zl_ҩƷ�շ���¼_������ҩ(v_��ҩ��¼.��ҩid, ������Ա_In, ����ʱ��_In, Null, Null, Null, v_��ҩ��¼.����, Null, ������Ա_In);
  
    --ȡ��ҩ����id
    Select a.Id
    Into v_��ҩid
    From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B
    Where b.Id = v_��ҩ��¼.��ҩid And a.���� = b.���� And a.No = b.No And a.�ⷿid + 0 = b.�ⷿid And a.ҩƷid + 0 = b.ҩƷid And
          a.��� = b.��� And Mod(a.��¼״̬, 3) = 1 And a.������� Is Null;
  
    --��Һ��ҩ�����е��շ�ID����Ϊ��ҩ�������շ�ID
    Update ��Һ��ҩ���� Set �շ�id = v_��ҩid Where ��¼id = v_��ҩ��¼.��ҩid And �շ�id = v_��ҩ��¼.�շ�id;
  
    If v_�շ�ids Is Null Then
      v_�շ�ids := v_��ҩid;
    Else
      v_�շ�ids := v_�շ�ids || ',' || v_��ҩid;
    End If;
  
    --ȡԭʼid
    Select a.Id
    Into v_ԭʼid
    From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B
    Where b.Id = v_��ҩ��¼.��ҩid And a.���� = b.���� And a.No = b.No And a.�ⷿid + 0 = b.�ⷿid And a.ҩƷid + 0 = b.ҩƷid And
          a.��� = b.��� And Mod(a.��¼״̬, 3) = 0 And a.������� Is Not Null;
  
    Insert Into ��Һ��ҩ����
      (��¼id, �շ�id, ����)
      Select ��¼id, v_ԭʼid, ���� From ��Һ��ҩ���� Where ��¼id = v_��ҩ��¼.��ҩid And �շ�id = v_��ҩid;
  
    v_�շ�ids := v_�շ�ids || ',' || v_ԭʼid;
  End Loop;

  --��������
  For v_�������� In c_�������� Loop
    Zl_סԺ���ʼ�¼_Delete(v_��������.No, v_��������.�������, v_Usercode, Zl_Username, 2, 1, 1, d_���ʱ��);
  End Loop;

  --������Һ�Ա�ҩ�嵥�����������ҩƷ���ѷ�ҩ����
  For v_�Ա�ҩ��¼ In c_�Ա�ҩ��¼ Loop
    --����Һ����������Ա�ҩ,���ռ���ҩƷ�շ���¼���е�id
    For v_�Ա�ҩ�շ���¼ In (Select a.Id, a.����, a.Ч��, a.����, a.ʵ������ As ��ҩ��, a.����, a.����id
                      From ҩƷ�շ���¼ A
                      Where a.�ƻ�id = v_�Ա�ҩ��¼.Id And a.ҩƷid = v_�Ա�ҩ��¼.ҩƷid And a.����� Is Not Null And a.������� Is Not Null And
                            (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0)
                      Order By a.����) Loop
    
      --�ж�������������ﻹ��סԺ 
      Begin
        Select 1 Into n_���� From ������ü�¼ Where ID = v_�Ա�ҩ�շ���¼.����id;
      Exception
        When Others Then
          n_���� := 2;
      End;
    
      Zl_ҩƷ�շ���¼_������ҩ(v_�Ա�ҩ�շ���¼.Id, ������Ա_In, ����ʱ��_In, Null, Null, Null, v_�Ա�ҩ�շ���¼.��ҩ��, Null, ������Ա_In, 2, n_����);
    End Loop;
  End Loop;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Һ��ҩ��¼_�������;
/

--137889:����,2019-03-18,����������Ա�ҩ�Ĵ���
Create Or Replace Procedure Zl_��Һ��ҩ��¼_ȡ����ҩ(��ҩid_In In Varchar2 --ID��:��ҩID1,��ҩID2....
                                           ) Is
  v_Id       Varchar2(20);
  v_��ҩid   Varchar2(20);
  v_Tmp      Varchar2(4000);
  v_Date     Date;
  v_������Ա ��Һ��ҩ��¼.������Ա%Type;
  d_����ʱ�� ��Һ��ҩ��¼.����ʱ��%Type;
  n_����״̬ ��Һ��ҩ��¼.����״̬%Type;

  v_ֹͣ��ҩids    Varchar2(4000); --���Ա�ҩ��ҩʱ�����˶Բ��ϣ���ȡ�������Һ���ġ�ȡ����ҩ������
  n_�Ա�ҩ����     ҩƷ�շ���¼.ʵ������%Type;
  n_�Ա�ҩ�������� ҩƷ�շ���¼.ʵ������%Type; --���Ա�ҩ��ҩƷ�շ���¼�п��Ա�����������
  n_����           Number; --1�����ﵥ�ݣ�2��סԺ����

  v_Error Varchar2(255);
  Err_Custom Exception;

  Cursor c_��ҩ���� Is
    Select /*+ rule*/
    Distinct c.��¼id, a.Id As ��ҩid, c.�շ�id, a.����, a.Ч��, a.����, c.���� As ��ҩ��, a.ҩƷid, a.����
    From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B, ��Һ��ҩ���� C, Table(Cast(f_Num2list(��ҩid_In) As Zltools.t_Numlist)) D
    Where c.��¼id = d.Column_Value And b.Id = c.�շ�id And a.���� = b.���� And a.No = b.No And a.�ⷿid + 0 = b.�ⷿid And
          a.ҩƷid + 0 = b.ҩƷid And a.��� = b.��� And (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0)
    Order By a.ҩƷid, a.����;

  v_��ҩ���� c_��ҩ����%RowType;

  Cursor c_�Ա�ҩ��¼ Is
    Select Distinct a.Id, b.��������, c.����ϵ��, c.ҩƷid
    From ��Һ��ҩ��¼ A, ����ҽ����¼ B, ҩƷ��� C, Table(Cast(f_Num2list(��ҩid_In) As Zltools.t_Numlist)) D
    Where a.ҽ��id = b.���id And a.Id = d.Column_Value And b.�շ�ϸĿid = c.ҩƷid And b.ִ������ = 5 And b.ִ�б�� = 0 And
          b.�շ�ϸĿid In (Select d.ҩƷid From ��Һ�Ա�ҩ�嵥 D Where d.�Ƿ����� = 1);

  v_�Ա�ҩ��¼ c_�Ա�ҩ��¼%RowType;
Begin
  If ��ҩid_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := ��ҩid_In || ',';
  End If;
  While v_Tmp Is Not Null Loop
    --�ֽⵥ��ID��
    v_Id  := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp := Replace(',' || v_Tmp, ',' || v_Id || ',');
  
    --��鵱ǰ��Һ����״̬�Ƿ�Ϊ����ҩ״̬
    Begin
      Select ����״̬ Into n_����״̬ From ��Һ��ҩ��¼ Where ID = v_Id;
    
      If n_����״̬ != 2 Then
        v_Error := '�������ѱ����������ܽ���ȡ����ҩ������';
        Raise Err_Custom;
      End If;
    Exception
      When Others Then
        v_Error := '�������ѱ�������';
        Raise Err_Custom;
    End;
  
    Select ������Ա, ����ʱ��
    Into v_������Ա, d_����ʱ��
    From ��Һ��ҩ״̬
    Where ��ҩid = v_Id And �������� = 1 And Rownum = 1;
  
    Update ��Һ��ҩ��¼ Set ����״̬ = 1, ������Ա = v_������Ա, ����ʱ�� = d_����ʱ�� Where ID = v_Id;
  
    --��[��Һ��ҩ״̬]���м�¼��ȡ����ҩ���Ĳ���
    Insert Into ��Һ��ҩ״̬
      (��ҩid, ��������, ������Ա, ����ʱ��, ����˵��)
    Values
      (v_Id, 1, v_������Ա, Sysdate, 'ȡ����ҩ');
  
  End Loop;

  Select Sysdate Into v_Date From Dual;

  --������Һ�Ա�ҩ�嵥�����������ҩƷ���ѷ�ҩ����
  For v_�Ա�ҩ��¼ In c_�Ա�ҩ��¼ Loop
  
    n_�Ա�ҩ���� := v_�Ա�ҩ��¼.�������� / v_�Ա�ҩ��¼.����ϵ��;
  
    Select Sum(a.ʵ������)
    Into n_�Ա�ҩ��������
    From ҩƷ�շ���¼ A
    Where Mod(a.��¼״̬, 3) = 1 And a.�ƻ�id = v_�Ա�ҩ��¼.Id And a.ҩƷid = v_�Ա�ҩ��¼.ҩƷid And a.����� Is Not Null And
          a.������� Is Not Null;
  
    If n_�Ա�ҩ�������� < n_�Ա�ҩ���� Then
      --��������˶Բ��ϣ����ռ���ǰ��ҩid,��������ͬ���������Һ���Ķ�ӦҩƷ
      If v_ֹͣ��ҩids Is Null Then
        v_ֹͣ��ҩids := v_�Ա�ҩ��¼.Id;
      Else
        v_ֹͣ��ҩids := v_ֹͣ��ҩids || ',' || v_�Ա�ҩ��¼.Id;
      End If;
    
      Exit;
    
    End If;
  
    --����Һ����������Ա�ҩ,���ռ���ҩƷ�շ���¼���е�id
    For v_�Ա�ҩ�շ���¼ In (Select a.Id, a.����, a.Ч��, a.����, a.ʵ������ As ��ҩ��, a.����, a.����id
                      From ҩƷ�շ���¼ A
                      Where a.�ƻ�id = v_�Ա�ҩ��¼.Id And a.ҩƷid = v_�Ա�ҩ��¼.ҩƷid And a.����� Is Not Null And a.������� Is Not Null And
                            (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0)
                      Order By a.����) Loop
    
      --�ж�������������ﻹ��סԺ 
      Begin
        Select 1 Into n_���� From ������ü�¼ Where ID = v_�Ա�ҩ�շ���¼.����id;
      Exception
        When Others Then
          n_���� := 2;
      End;

      Zl_ҩƷ�շ���¼_������ҩ(v_�Ա�ҩ�շ���¼.Id, Zl_Username, v_Date, v_�Ա�ҩ�շ���¼.����, v_�Ա�ҩ�շ���¼.Ч��, v_�Ա�ҩ�շ���¼.����, v_�Ա�ҩ�շ���¼.��ҩ��, Null,
                     Zl_Username, 2, n_����);
    End Loop;
  End Loop;

  For v_��ҩ���� In c_��ҩ���� Loop
    --�ų����жϵ���Һ��
    If Instr(',' || v_ֹͣ��ҩids || ',', ',' || v_��ҩ����.��¼id || ',') = 0 Then
      --������ҩ
      Zl_ҩƷ�շ���¼_������ҩ(v_��ҩ����.��ҩid, Zl_Username, v_Date, v_��ҩ����.����, v_��ҩ����.Ч��, v_��ҩ����.����, v_��ҩ����.��ҩ��, Null, Zl_Username);
    
      Select Max(a.Id)
      Into v_��ҩid
      From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B
      Where b.Id = v_��ҩ����.��ҩid And a.���� = b.���� And a.No = b.No And a.�ⷿid + 0 = b.�ⷿid And a.ҩƷid + 0 = b.ҩƷid And
            a.��� = b.��� And Mod(a.��¼״̬, 3) = 1 And a.������� Is Null;
    
      --�滻��Һ��ҩ�����е��շ�ID
      Update ��Һ��ҩ���� Set �շ�id = v_��ҩid Where ��¼id = v_��ҩ����.��¼id And �շ�id = v_��ҩ����.�շ�id;
    End If;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Һ��ҩ��¼_ȡ����ҩ;
/

--137889:����,2019-03-19,�����߼�����
Create Or Replace Procedure Zl_��Һ��ҩ��¼_��ҩ
(
  ����id_In   In ��Һ��ҩ��¼.����id%Type,
  ��ҩid_In   In Varchar2, --ID��:ID1,ID2....
  ��ҩ����_In In ��Һ��ҩ��¼.��ҩ����%Type,
  ������Ա_In In ��Һ��ҩ״̬.������Ա%Type := Null,
  ����ʱ��_In In ��Һ��ҩ״̬.����ʱ��%Type := Null,
  �ƶ�����_In In Number := 0
) Is
  v_Tansid Varchar2(20);
  v_Tmp    Varchar2(4000);

  v_�շ�ids      Varchar2(4000);
  v_Error        Varchar2(255);
  n_�Ƿ���     ��Һ��ҩ��¼.�Ƿ���%Type;
  n_����״̬     ��Һ��ҩ��¼.����״̬%Type;
  v_��ҩ��       Varchar2(20);
  v_��ҩ̨       Varchar2(20);
  n_��ҩ̨id     Number(4);
  n_����id       Number(18);
  n_����         Number(2);
  d_����         Date;
  n_��ͨ���С�� Number;
  n_��ͨ����С�� Number;

  v_�Ա�ҩ�շ�ids      Varchar2(4000);
  n_�Ա�ҩ����         ҩƷ�շ���¼.ʵ������%Type;
  n_�Ա�ҩ��������     ҩƷ�շ���¼.ʵ������%Type; --���Ա�ҩ��ҩƷ�շ���¼�п��Ա�����������
  n_�Ա�ҩִ�л������� ҩƷ�շ���¼.ʵ������%Type; --���Ա�ҩ��ҩ�����Ļ�������
  n_�Ա�ҩ���ռ�����   ҩƷ�շ���¼.ʵ������%Type; --�����շ�id��ͳ�Ƶ�ǰ��׼��������
  n_�Ա�ҩδ�ռ�����   ҩƷ�շ���¼.ʵ������%Type; --�����շ�id��ͳ�Ƶ�ǰδ׼��������
  n_�Ա�ҩ���         ҩƷ�շ���¼.���%Type; --�����շ�id��ͳ�Ƶ�ǰδ׼��������
  n_�շ�id             ҩƷ�շ���¼.Id%Type;

  Err_Custom Exception;
  Cursor c_�շ���¼ Is
    Select /*+ rule*/
     a.Id, Nvl(a.����, 0) As ����
    From ҩƷ�շ���¼ A,
         (Select Distinct �շ�id
           From ��Һ��ҩ���� A, Table(Cast(f_Num2list(��ҩid_In) As Zltools.t_Numlist)) B
           Where a.��¼id = b.Column_Value) B
    Where a.Id = b.�շ�id And a.����� Is Null
    Order By a.ҩƷid, a.����;

  v_�շ���¼ c_�շ���¼%RowType;
Begin
  If ��ҩid_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := ��ҩid_In || ',';
  End If;

  v_�Ա�ҩ�շ�ids := Null;
  Select ���� Into n_��ͨ����С�� From ҩƷ���ľ��� Where ��� = 1 And ���� = 2 And ��λ = 1;
  Select ���� Into n_��ͨ���С�� From ҩƷ���ľ��� Where ��� = 1 And ���� = 4 And ��λ = 5;

  While v_Tmp Is Not Null Loop
    --�ֽⵥ��ID��
    v_Tansid := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp    := Replace(',' || v_Tmp, ',' || v_Tansid || ',');
  
    --��鵱ǰ��Һ����״̬�Ƿ�Ϊ����ҩ״̬
    Begin
      Select ����״̬ Into n_����״̬ From ��Һ��ҩ��¼ Where ID = v_Tansid;
    
      If n_����״̬ > 1 Then
        v_Error := '�������ѱ����������ܽ��з�ҩ��';
        Raise Err_Custom;
      End If;
    Exception
      When Others Then
        v_Error := '�������ѱ�������';
        Raise Err_Custom;
    End;
  
    Begin
      Select �Ƿ��� Into n_�Ƿ��� From ��Һ��ҩ��¼ Where ID = v_Tansid For Update Nowait;
    Exception
      When Others Then
        v_Error := '���������û���ִ�з�ҩ�������ظ�������';
        Raise Err_Custom;
    End;
  
    --������Һ�Ա�ҩ�嵥�����������ҩƷ�Ĵ���ҩ����
    For v_�Ա�ҩ��¼ In (Select a.Id, b.�շ�ϸĿid As ҩƷid, b.��������, c.����ϵ��, a.����id, b.����id, b.�ܸ�����, b.�걾��λ As ҩƷƷ��
                    From ��Һ��ҩ��¼ A, ����ҽ����¼ B, ҩƷ��� C
                    Where a.ҽ��id = b.���id And b.�շ�ϸĿid = c.ҩƷid And a.Id = v_Tansid And b.ִ������ = 5 And b.ִ�б�� = 0 And
                          b.�շ�ϸĿid In (Select d.ҩƷid From ��Һ�Ա�ҩ�嵥 D Where d.�Ƿ����� = 1)) Loop
    
      n_�Ա�ҩ���� := v_�Ա�ҩ��¼.�������� / v_�Ա�ҩ��¼.����ϵ��;
    
      --����Ƿ������ִ�й��Ĵ���ҩ����
      Select Nvl(Sum(b.ʵ������), 0)
      Into n_�Ա�ҩִ�л�������
      From δ��ҩƷ��¼ A, ҩƷ�շ���¼ B
      Where a.���� = b.���� And a.No = b.No And a.�ⷿid = b.�ⷿid And Mod(b.��¼״̬, 3) = 1 And b.��¼״̬ <> 1 And
            a.����id = v_�Ա�ҩ��¼.����id And a.�ⷿid = v_�Ա�ҩ��¼.����id And b.����� Is Null And b.������� Is Null And
            b.ҩƷid = v_�Ա�ҩ��¼.ҩƷid And b.�ƻ�id = v_�Ա�ҩ��¼.Id;
    
      If n_�Ա�ҩִ�л������� > 0 And n_�Ա�ҩִ�л������� < n_�Ա�ҩ���� Then
        v_Error := 'ҩƷ��' || v_�Ա�ҩ��¼.ҩƷƷ�� || '�����������㣬���ܽ��з�ҩ��';
        Raise Err_Custom;
      Elsif n_�Ա�ҩִ�л������� = n_�Ա�ҩ���� Then
        --�ռ���ִ�й�����ҩid��¼
        For v_�Ա�ҩ��ִ�м�¼ In (Select b.Id As �շ�id, b.ʵ������, b.����, b.��¼״̬, b.����, b.No, b.�ⷿid
                           From δ��ҩƷ��¼ A, ҩƷ�շ���¼ B
                           Where a.���� = b.���� And a.No = b.No And a.�ⷿid = b.�ⷿid And Mod(b.��¼״̬, 3) = 1 And b.��¼״̬ <> 1 And
                                 a.����id = v_�Ա�ҩ��¼.����id And a.�ⷿid = v_�Ա�ҩ��¼.����id And b.ҩƷid = v_�Ա�ҩ��¼.ҩƷid And
                                 b.������� Is Null And b.����� Is Null And b.�ƻ�id = v_�Ա�ҩ��¼.Id
                           Order By b.����) Loop
          If v_�Ա�ҩ�շ�ids Is Null Then
            v_�Ա�ҩ�շ�ids := v_�Ա�ҩ��ִ�м�¼.�շ�id || ',' || v_�Ա�ҩ��ִ�м�¼.����;
          Else
            v_�Ա�ҩ�շ�ids := v_�Ա�ҩ�շ�ids || '|' || v_�Ա�ҩ��ִ�м�¼.�շ�id || ',' || v_�Ա�ҩ��ִ�м�¼.����;
          End If;
        End Loop;
      Elsif n_�Ա�ҩִ�л������� = 0 Then
        --����ӦҩƷ�����ε��ܺ��Ƿ����㱾�η�ҩ����
        Select Nvl(Sum(b.ʵ������), 0)
        Into n_�Ա�ҩ��������
        From δ��ҩƷ��¼ A, ҩƷ�շ���¼ B
        Where a.���� = b.���� And a.No = b.No And a.�ⷿid = b.�ⷿid And Mod(b.��¼״̬, 3) = 1 And a.����id = v_�Ա�ҩ��¼.����id And
              a.�ⷿid = v_�Ա�ҩ��¼.����id And b.����� Is Null And b.������� Is Null And b.ҩƷid = v_�Ա�ҩ��¼.ҩƷid And b.�ƻ�id Is Null And
              Exists (Select 1 From ������ü�¼ C Where c.Id = b.����id);
      
        If n_�Ա�ҩ�������� < n_�Ա�ҩ���� Then
          v_Error := 'ҩƷ��' || v_�Ա�ҩ��¼.ҩƷƷ�� || '�����������㣬���ܽ��з�ҩ��';
          Raise Err_Custom;
        End If;
      
        --ѭ����ֲ��ռ���ִ���շ�ID��,��ʽΪ:"id1,����1|id2,����2|....."
        n_�Ա�ҩ���ռ����� := 0;
        n_�Ա�ҩδ�ռ����� := 0;
      
        For v_�Ա�ҩ�շ���¼ In (Select b.Id As �շ�id, b.ʵ������, b.����, b.��¼״̬, b.����, b.No, b.�ⷿid
                          From δ��ҩƷ��¼ A, ҩƷ�շ���¼ B
                          Where a.���� = b.���� And a.No = b.No And a.�ⷿid = b.�ⷿid And Mod(b.��¼״̬, 3) = 1 And
                                a.����id = v_�Ա�ҩ��¼.����id And a.�ⷿid = v_�Ա�ҩ��¼.����id And b.ҩƷid = v_�Ա�ҩ��¼.ҩƷid And
                                b.������� Is Null And b.����� Is Null And b.�ƻ�id Is Null And Exists
                           (Select 1 From ������ü�¼ C Where c.Id = b.����id)
                          Order By b.����) Loop
        
          n_�Ա�ҩδ�ռ����� := n_�Ա�ҩ���� - n_�Ա�ҩ���ռ�����;
          n_�Ա�ҩ���ռ����� := n_�Ա�ҩ���ռ����� + v_�Ա�ҩ�շ���¼.ʵ������;
        
          If n_�Ա�ҩ���ռ����� < n_�Ա�ҩ���� Then
            --ֱ���ռ���ǰ�շ���¼
            If v_�Ա�ҩ�շ�ids Is Null Then
              v_�Ա�ҩ�շ�ids := v_�Ա�ҩ�շ���¼.�շ�id || ',' || v_�Ա�ҩ�շ���¼.����;
            Else
              v_�Ա�ҩ�շ�ids := v_�Ա�ҩ�շ�ids || '|' || v_�Ա�ҩ�շ���¼.�շ�id || ',' || v_�Ա�ҩ�շ���¼.����;
            End If;
          
            Update ҩƷ�շ���¼ Set �ƻ�id = v_Tansid Where ID = v_�Ա�ҩ�շ���¼.�շ�id;
          
          Elsif n_�Ա�ҩ���ռ����� > n_�Ա�ҩ���� Then
            --��Ҫ��֣����ռ�����շ���¼
          
            Select ҩƷ�շ���¼_Id.Nextval Into n_�շ�id From Dual;
          
            Select Max(a.���)
            Into n_�Ա�ҩ���
            From ҩƷ�շ���¼ A
            Where a.���� = v_�Ա�ҩ�շ���¼.���� And a.No = v_�Ա�ҩ�շ���¼.No And a.�ⷿid = v_�Ա�ҩ�շ���¼.�ⷿid;
          
            Update ҩƷ�շ���¼
            Set ��д���� = ��д���� - n_�Ա�ҩδ�ռ�����, ʵ������ = ʵ������ - n_�Ա�ҩδ�ռ�����, ���۽�� = ���ۼ� * (ʵ������ - n_�Ա�ҩδ�ռ�����)
            Where ID = v_�Ա�ҩ�շ���¼.�շ�id;
          
            Insert Into ҩƷ�շ���¼
              (ID, ��¼״̬, ����, NO, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ����,
               ���ۼ�, ���۽��, ���, ժҪ, ������, ��������, ��ҩ��, �����, �������, ����id, ����, Ƶ��, �÷�, ��ҩ����, ��ҩ��λid, ��������, ��׼�ĺ�, ע��֤��, �ƻ�id, ԭ����,
               ������־)
              Select n_�շ�id, ��¼״̬, ����, NO, n_�Ա�ҩ��� + 1, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ����, n_�Ա�ҩδ�ռ�����,
                     n_�Ա�ҩδ�ռ�����, �ɱ���, �ɱ����, ����, ���ۼ�, Round(���ۼ� * n_�Ա�ҩδ�ռ�����, n_��ͨ���С��), ���, ժҪ, ������, ��������, ��ҩ��, �����,
                     �������, ����id, ����, Ƶ��, �÷�, ��ҩ����, ��ҩ��λid, ��������, ��׼�ĺ�, ע��֤��, v_Tansid, ԭ����, ������־




              
              From ҩƷ�շ���¼
              Where ID = v_�Ա�ҩ�շ���¼.�շ�id;
          
            If v_�Ա�ҩ�շ�ids Is Null Then
              v_�Ա�ҩ�շ�ids := n_�շ�id || ',' || v_�Ա�ҩ�շ���¼.����;
            Else
              v_�Ա�ҩ�շ�ids := v_�Ա�ҩ�շ�ids || '|' || n_�շ�id || ',' || v_�Ա�ҩ�շ���¼.����;
            End If;
          
            --����ѭ��
            Exit;
          
          Else
            --�ռ����
            If v_�Ա�ҩ�շ�ids Is Null Then
              v_�Ա�ҩ�շ�ids := v_�Ա�ҩ�շ���¼.�շ�id || ',' || v_�Ա�ҩ�շ���¼.����;
            Else
              v_�Ա�ҩ�շ�ids := v_�Ա�ҩ�շ�ids || '|' || v_�Ա�ҩ�շ���¼.�շ�id || ',' || v_�Ա�ҩ�շ���¼.����;
            End If;
          
            Update ҩƷ�շ���¼ Set �ƻ�id = v_Tansid Where ID = v_�Ա�ҩ�շ���¼.�շ�id;
          
            --����ѭ��
            Exit;
          
          End If;
        End Loop;
      End If;
    End Loop;
  
    v_��ҩ̨   := '';
    n_��ҩ̨id := 0;
    n_����id   := 0;
    v_��ҩ��   := '';
    Begin
      Select ����, ID, ����id, ��ҩ����, ִ��ʱ��
      Into v_��ҩ̨, n_��ҩ̨id, n_����id, n_����, d_����
      From (Select f.����, f.Id, a.����id, a.��ҩ����, a.ִ��ʱ��
             From ��Һ��ҩ��¼ A, ��Һ��ҩ���� B, ҩƷ�շ���¼ C, ��Һ̨ҩƷ���� D, ��Һ̨ F
             Where a.Id = b.��¼id And b.�շ�id = c.Id And c.ҩƷid = d.ҩƷid And d.��ҩ̨id = f.Id And c.�ⷿid = d.����id And
                   a.Id = v_Tansid
             Order By d.��ҩ̨id)
      Where Rownum = 1;
    
      Select ��ҩ��
      Into v_��ҩ��
      From ��Һ��������
      Where ����id = n_����id And ��ҩ̨id = n_��ҩ̨id And ���� = n_���� And
            ���� = To_Date(To_Char(Sysdate, 'yyyy-mm-dd'), 'yyyy-mm-dd');
    Exception
      When Others Then
        Null;
    End;
  
    Update ��Һ��ҩ��¼
    Set ����״̬ = 2, ������Ա = ������Ա_In, ����ʱ�� = ����ʱ��_In, ��ҩ���� = ��ҩ����_In, ��ҩ̨ = v_��ҩ̨
    Where ID = v_Tansid;
  
    Insert Into ��Һ��ҩ״̬
      (��ҩid, ��������, ������Ա, ����ʱ��, ʵ�ʹ�����Ա)
    Values
      (v_Tansid, 2, ������Ա_In, ����ʱ��_In, v_��ҩ��);
    If n_�Ƿ��� <> 0 And �ƶ�����_In = 0 Then
      Update ��Һ��ҩ��¼ Set ����״̬ = 4, ������Ա = ������Ա_In, ����ʱ�� = ����ʱ��_In Where ID = v_Tansid;
      Insert Into ��Һ��ҩ״̬ (��ҩid, ��������, ������Ա, ����ʱ��) Values (v_Tansid, 4, ������Ա_In, ����ʱ��_In);
    End If;
  End Loop;

  For v_�շ���¼ In c_�շ���¼ Loop
    If v_�շ�ids Is Null Then
      v_�շ�ids := v_�շ���¼.Id || ',' || v_�շ���¼.����;
    Else
      If Length(v_�շ�ids || '|' || v_�շ���¼.Id || ',' || v_�շ���¼.����) > 3950 Then
        Zl_ҩƷ�շ���¼_������ҩ(v_�շ�ids, ����id_In, ������Ա_In, ����ʱ��_In, 4, ������Ա_In, ��ҩ����_In);
        v_�շ�ids := v_�շ���¼.Id || ',' || v_�շ���¼.����;
      Else
        v_�շ�ids := v_�շ�ids || '|' || v_�շ���¼.Id || ',' || v_�շ���¼.����;
      End If;
    End If;
  End Loop;

  If Not v_�շ�ids Is Null Then
    Zl_ҩƷ�շ���¼_������ҩ(v_�շ�ids, ����id_In, ������Ա_In, ����ʱ��_In, 4, ������Ա_In, ��ҩ����_In);
  End If;

  --�����Ա�ҩ
  If Not v_�Ա�ҩ�շ�ids Is Null Then
    Zl_ҩƷ�շ���¼_������ҩ(v_�Ա�ҩ�շ�ids, ����id_In, ������Ա_In, ����ʱ��_In, 4, ������Ա_In, ��ҩ����_In);
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Һ��ҩ��¼_��ҩ;
/

--137889:����,2019-03-18,����������Ա�ҩ�Ĵ���
Create Or Replace Procedure Zl_ҩƷ�շ���¼_������ҩ
(
  Billid_In     In ҩƷ�շ���¼.Id%Type,
  People_In     In ҩƷ�շ���¼.�����%Type,
  Date_In       In ҩƷ�շ���¼.�������%Type,
  ����_In       In ҩƷ���.�ϴ�����%Type := Null,
  Ч��_In       In ҩƷ���.Ч��%Type := Null,
  ����_In       In ҩƷ���.�ϴβ���%Type := Null,
  ��ҩ����_In   In ҩƷ�շ���¼.ʵ������%Type := Null,
  ��ҩ�ⷿ_In   In ҩƷ�շ���¼.�ⷿid%Type := Null,
  ��ҩ��_In     In ҩƷ�շ���¼.������%Type := Null,
  Intdigit_In   In Number := 2,
  ����_In       In Number := 2,
  ���ܷ�ҩ��_In In ҩƷ�շ���¼.���ܷ�ҩ��%Type := Null
) Is
  --ֻ������
  Int��¼״̬   ҩƷ�շ���¼.��¼״̬%Type;
  Intִ��״̬   סԺ���ü�¼.ִ��״̬%Type;
  Bln������ҩ   Number;
  Lng������id Number(18);
  Strno         ҩƷ�շ���¼.No%Type;
  Int����       ҩƷ�շ���¼.����%Type;
  Lng�ⷿid     ҩƷ�շ���¼.�ⷿid%Type;
  LngҩƷid     ҩƷ�շ���¼.ҩƷid%Type;
  Dblʵ������   ҩƷ�շ���¼.ʵ������%Type;
  Dblʵ�ʽ��   ҩƷ�շ���¼.���۽��%Type;
  Dblʵ�ʳɱ�   ҩƷ�շ���¼.�ɱ����%Type;
  Dblʵ�ʲ��   ҩƷ�շ���¼.���%Type;
  Lng����id     ҩƷ�շ���¼.����id%Type;
  n_���ۼ�      ҩƷ�շ���¼.���ۼ�%Type;
  n_�Ƿ���    Number;
  n_ʱ�۷���    Number;

  --20020731 Modified by zyb
  --������ҩʱ�������������ʸı��Ĵ���
  Lng������ ҩƷ�շ���¼.����%Type;
  Lng����   ҩƷ���.ҩ������%Type;
  Lng����   ҩƷ�շ���¼.����%Type; --ԭ����

  Str����        ҩƷ�շ���¼.����%Type; --ԭ����
  DateЧ��       ҩƷ�շ���¼.Ч��%Type; --ԭЧ��
  n_�ϴι�Ӧ��id ҩƷ���.�ϴι�Ӧ��id%Type;
  n_�ϴβɹ���   ҩƷ���.�ϴβɹ���%Type;
  v_�ϴβ���     ҩƷ���.�ϴβ���%Type;
  v_ԭ����       ҩƷ���.ԭ����%Type;
  d_�ϴ��������� ҩƷ���.�ϴ���������%Type;
  v_��׼�ĺ�     ҩƷ���.��׼�ĺ�%Type;

  n_��¼����   סԺ���ü�¼.��¼����%Type;
  v_�շ����   סԺ���ü�¼.�շ����%Type;
  n_����       ҩƷ�շ���¼.����%Type;
  n_ԭʼ����   ҩƷ�շ���¼.ʵ������%Type;
  v_������¼id ҩƷ�շ���¼.Id%Type;
  Err_Custom Exception;
  v_Error    Varchar2(255);
  v_��ҩȷ�� ҩ����ҩ����.��ҩȷ��%Type;
  v_��ҩ     ҩ����ҩ����.��ҩ%Type;
  v_�Ŷ�״̬ Number(1);
  v_ִ��ʱ�� ҩƷ�շ���¼.�������%Type;

Begin
  If ��ҩ����_In Is Not Null Then
    If ��ҩ����_In = 0 Then
      Return;
    End If;
  End If;

  --��ȡ���շ���¼�ĵ��ݡ�ҩƷID���ⷿID
  Select a.����, a.No, a.�ⷿid, a.ҩƷid, a.����id, a.������id, a.��¼״̬, Nvl(a.����, 0), a.����, a.Ч��, a.��ҩ��λid, a.����, a.ԭ����, a.��������,
         a.��׼�ĺ�, a.�ɱ���, a.����, Nvl(a.ʵ������, 0) * Nvl(a.����, 1) As ʵ������, a.���ۼ�, Nvl(b.�Ƿ���, 0) �Ƿ���
  Into Int����, Strno, Lng�ⷿid, LngҩƷid, Lng����id, Lng������id, Int��¼״̬, Lng����, Str����, DateЧ��, n_�ϴι�Ӧ��id, v_�ϴβ���, v_ԭ����,
       d_�ϴ���������, v_��׼�ĺ�, n_�ϴβɹ���, n_����, n_ԭʼ����, n_���ۼ�, n_�Ƿ���
  From ҩƷ�շ���¼ A, �շ���ĿĿ¼ B
  Where a.ҩƷid = b.Id And a.Id = Billid_In;

  Begin
    Select Nvl(��ҩȷ��, 0), Nvl(��ҩ, 0)
    Into v_��ҩȷ��, v_��ҩ
    From ҩ����ҩ����
    Where ҩ��id = Lng�ⷿid And Rownum = 1;
  
  Exception
    When Others Then
      v_��ҩȷ�� := 0;
      v_��ҩ     := 0;
      Null;
  End;

  If v_��ҩȷ�� = 0 And v_��ҩ = 0 Then
    v_�Ŷ�״̬ := 2;
  Elsif v_��ҩȷ�� = 1 Then
    v_�Ŷ�״̬ := 0;
  Elsif v_��ҩ = 1 Then
    v_�Ŷ�״̬ := 1;
  End If;

  --��ȡ�ñʼ�¼ʣ��δ�������������
  --������������δ���������
  Select Sum(Nvl(ʵ������, 0) * Nvl(����, 1)), Sum(Nvl(���۽��, 0)), Sum(Nvl(�ɱ����, 0)), Sum(Nvl(���, 0))
  Into Dblʵ������, Dblʵ�ʽ��, Dblʵ�ʳɱ�, Dblʵ�ʲ��
  From ҩƷ�շ���¼
  Where ����� Is Not Null And NO = Strno And ���� = Int���� And ��� = (Select ��� From ҩƷ�շ���¼ Where ID = Billid_In);

  --���������ҩ��Ϊ�㣬��ʾ����ҩ
  If Dblʵ������ = 0 Then
    v_Error := '�õ����ѱ���������Ա��ҩ����ˢ�º����ԣ�';
    Raise Err_Custom;
  End If;
  If Nvl(��ҩ����_In, 0) > Dblʵ������ Then
    v_Error := '�õ����ѱ���������Ա������ҩ����ˢ�º����ԣ�';
    Raise Err_Custom;
  End If;

  --��ȡ��ҩƷ��ǰ�Ƿ��������Ϣ
  Select Nvl(ҩ������, 0) Into Lng���� From ҩƷ��� Where ҩƷid = LngҩƷid;
  --����ǲ�����ҩ�������¼������۽����
  Bln������ҩ := 0;
  If Not (��ҩ����_In Is Null Or Nvl(��ҩ����_In, 0) = Dblʵ������) Then
    Bln������ҩ := 1;
  End If;
  If Bln������ҩ = 1 Then
    Dblʵ�ʽ�� := Round(Dblʵ�ʽ�� * ��ҩ����_In / Dblʵ������, Intdigit_In);
    Dblʵ�ʳɱ� := Round(Dblʵ�ʳɱ� * ��ҩ����_In / Dblʵ������, Intdigit_In);
    Dblʵ�ʲ�� := Round(Dblʵ�ʲ�� * ��ҩ����_In / Dblʵ������, Intdigit_In);
    Dblʵ������ := ��ҩ����_In;
  End If;

  If n_ԭʼ���� = ��ҩ����_In Then
    Dblʵ������ := ��ҩ����_In / n_����;
  Else
    n_���� := 1;
  End If;

  --lng����:0-������;1-����;2-ԭ�������ֲ�������������������;3-ԭ���������ַ���������������
  If Lng���� = 0 And Lng���� <> 0 Then
    --ԭ�������ֲ�������������������
    Lng���� := 2;
  Elsif Lng���� <> 0 And Lng���� = 0 Then
    --ԭ������,�ַ���,�����µ����Σ������²����ķ�ҩ��¼��ʹ��
    Lng���� := 3;
  Else
    If Lng���� = 0 Then
      Lng���� := 0;
    Else
      Lng���� := 1;
    End If;
  End If;
  --�ж��Ƿ�ʱ�۷���
  If (Lng���� = 1 Or Lng���� = 3) And n_�Ƿ��� = 1 Then
    n_ʱ�۷��� := 1;
  Else
    n_ʱ�۷��� := 0;
  End If;

  --��¼״̬�ĺ��������仯
  --�����ļ�¼״̬        :iif(int��¼״̬=1,0,1)+1
  --�������ļ�¼״̬        :iif(int��¼״̬=1,0,1)+2
  --�ȴ���ҩ�ļ�¼״̬    :iif(int��¼״̬=1,0,1)+3

  --����������¼
  Select ҩƷ�շ���¼_Id.Nextval Into v_������¼id From Dual;
  Insert Into ҩƷ�շ���¼
    (ID, ��¼״̬, ����, NO, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ����, ���ۼ�, ���۽��,
     ���, ժҪ, ������, ��������, ��ҩ��, �����, �������, ����id, ����, Ƶ��, �÷�, ��ҩ����, ���, ������, ��ҩ��λid, ��������, ��׼�ĺ�, ���ܷ�ҩ��, ��ҩ��ʽ, ע��֤��, �ƻ�id,
     ԭ����)
    Select v_������¼id, Int��¼״̬ + Decode(Int��¼״̬, 1, 0, 1) + 1, Int����, Strno, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����,
           ����, Ч��, n_����, -dblʵ������, -dblʵ������, �ɱ���, -dblʵ�ʳɱ�, ����, ���ۼ�, -dblʵ�ʽ��, -dblʵ�ʲ��, ժҪ, People_In, Date_In, ��ҩ��,
           People_In, Date_In, ����id, ����, Ƶ��, �÷�, ��ҩ����, ��ҩ�ⷿ_In, ��ҩ��_In, ��ҩ��λid, ��������, ��׼�ĺ�, ���ܷ�ҩ��_In, ��ҩ��ʽ, ע��֤��, �ƻ�id,
           ԭ����
    From ҩƷ�շ���¼
    Where ID = Billid_In;

  --����ǲ��ֳ�����������Ϊ1��ʵ������Ϊ������ʵ�������Ļ�
  --����������¼�Թ�������ҩ
  Select ҩƷ�շ���¼_Id.Nextval Into Lng������ From Dual;
  Insert Into ҩƷ�շ���¼
    (ID, ��¼״̬, ����, NO, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ����, ���ۼ�, ���۽��,
     ���, ժҪ, ������, ��������, ��ҩ��, �����, �������, ����id, ����, Ƶ��, �÷�, ��ҩ����, ��ҩ��λid, ��������, ��׼�ĺ�, ע��֤��, �ƻ�id, ԭ����)
    Select Lng������, Int��¼״̬ + Decode(Int��¼״̬, 1, 0, 1) + 3, Int����, Strno, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid,
           Decode(Lng����, 1, ����, 3, Lng������, 0), Decode(Lng����, 3, ����_In, 1, ����, ����), Decode(Lng����, 3, ����_In, ����),
           Decode(Lng����, 3, Ч��_In, Ч��), n_����, Dblʵ������, Dblʵ������, �ɱ���, Dblʵ�ʳɱ�, ����, ���ۼ�, Dblʵ�ʽ��, Dblʵ�ʲ��, ժҪ, ������, ��������,
           Null, Null, Null, ����id, ����, Ƶ��, �÷�, ��ҩ����, ��ҩ��λid, ��������, ��׼�ĺ�, ע��֤��, �ƻ�id, ԭ����
    
    From ҩƷ�շ���¼
    Where ID = Billid_In;

  Zl_δ��ҩƷ��¼_Insert(Lng������);

  --���·��ü�¼��ִ��״̬(0-δִ��;1-��ȫִ��;2-����ִ��)
  Select Decode(Sum(Nvl(����, 1) * ʵ������), Null, 0, 0, 0, 2)
  Into Intִ��״̬
  From ҩƷ�շ���¼
  Where ���� = Int���� And NO = Strno And ����id = Lng����id And ����� Is Not Null;

  If ����_In = 1 Then
    Select ��¼����, �շ���� Into n_��¼����, v_�շ���� From ������ü�¼ Where ID = Lng����id;
  Else
    Select ��¼����, �շ���� Into n_��¼����, v_�շ���� From סԺ���ü�¼ Where ID = Lng����id;
  End If;

  If Intִ��״̬ = 0 Then
    If ����_In = 1 Then
      Update ������ü�¼
      Set ִ��״̬ = Intִ��״̬, ִ���� = Null, ִ��ʱ�� = Null
      Where NO = Strno And
            ��� = (Select ��� From ������ü�¼ Where ID = (Select ����id From ҩƷ�շ���¼ Where ID = Billid_In)) And
            Mod(��¼����, 10) = n_��¼���� And ��¼״̬ <> 2 And ִ�в���id = Lng�ⷿid;
    Else
      Update סԺ���ü�¼ Set ִ��״̬ = Intִ��״̬, ִ���� = Null, ִ��ʱ�� = Null Where ID = Lng����id;
    End If;
  Else
    If ����_In = 1 Then
      Update ������ü�¼
      Set ִ��״̬ = Intִ��״̬
      Where NO = Strno And
            ��� = (Select ��� From ������ü�¼ Where ID = (Select ����id From ҩƷ�շ���¼ Where ID = Billid_In)) And
            Mod(��¼����, 10) = n_��¼���� And ��¼״̬ <> 2 And ִ�в���id = Lng�ⷿid;
    Else
      Update סԺ���ü�¼ Set ִ��״̬ = Intִ��״̬ Where ID = Lng����id;
    End If;
  End If;

  --����δ��ҩƷ��¼
  Begin
    If ����_In = 1 Then
      Insert Into δ��ҩƷ��¼
        (����, NO, ����id, ��ҳid, ����, ���ȼ�, �Է�����id, �ⷿid, ��ҩ����, ��������, ���շ�, ��ҩ��, ��ӡ״̬, δ����, ��ҩ��, �Ŷ�״̬)
        Select a.����, a.No, a.����id, Null, a.����, Nvl(b.���ȼ�, 0) ���ȼ�, a.�Է�����id, a.�ⷿid, a.��ҩ����, a.��������, a.���շ�, Null, 1, 1,
               a.��Ʒ�ϸ�֤, v_�Ŷ�״̬
        From (Select b.����, b.No, a.����id, a.����, Decode(a.��¼״̬, 0, 0, 1) ���շ�, b.�Է�����id, b.�ⷿid, b.��ҩ����, b.��������, c.���,
                      b.��Ʒ�ϸ�֤
               From ������ü�¼ A, ҩƷ�շ���¼ B, ������Ϣ C
               Where b.Id = Billid_In And a.Id = b.����id + 0 And a.����id = c.����id(+)) A, ��� B
        Where b.����(+) = a.���;
    Else
      Insert Into δ��ҩƷ��¼
        (����, NO, ����id, ��ҳid, ����, ���ȼ�, �Է�����id, �ⷿid, ��ҩ����, ��������, ���շ�, ��ҩ��, ��ӡ״̬, δ����, ��ҩ��, �Ŷ�״̬)
        Select a.����, a.No, a.����id, a.��ҳid, a.����, Nvl(b.���ȼ�, 0) ���ȼ�, a.�Է�����id, a.�ⷿid, a.��ҩ����, a.��������, a.���շ�, Null, 1, 1,
               a.��Ʒ�ϸ�֤, v_�Ŷ�״̬
        From (Select b.����, b.No, a.����id, a.��ҳid, a.����, Decode(a.��¼״̬, 0, 0, 1) ���շ�, b.�Է�����id, b.�ⷿid, b.��ҩ����, b.��������,
                      c.���, b.��Ʒ�ϸ�֤
               From סԺ���ü�¼ A, ҩƷ�շ���¼ B, ������Ϣ C
               Where b.Id = Billid_In And a.Id = b.����id + 0 And a.����id = c.����id(+)) A, ��� B
        Where b.����(+) = a.���;
    End If;
  
    --�޸Ĵ�������
    Zl_Prescription_Type_Update(Strno, n_��¼����, LngҩƷid, v_�շ����);
  Exception
    When Others Then
      Null;
  End;

  --�޸�ԭ��¼Ϊ��������¼
  Update ҩƷ�շ���¼ Set ��¼״̬ = Int��¼״̬ + Decode(Int��¼״̬, 1, 0, 1) + 2 Where ID = Billid_In;

  --�޸�ҩƷ���(������)
  If Lng���� <> 3 Then
    --����������Ҫ������ʵ�������ͽ���ۻ���ȥ���������û�����ڿ����������
    Zl_ҩƷ���_Update(v_������¼id, 3, 0);
  Else
    --ԭ�����������ڷ�����ֱ���ڿ�������µ���
    Insert Into ҩƷ���
      (�ⷿid, ҩƷid, ����, Ч��, ����, ʵ������, ʵ�ʽ��, ʵ�ʲ��, ���ۼ�, �ϴ�����, �ϴβ���, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ���������, ��׼�ĺ�, ƽ���ɱ���)
    Values
      (Lng�ⷿid, LngҩƷid, Lng������, Ч��_In, 1, Dblʵ������ * n_����, Dblʵ�ʽ��, Dblʵ�ʲ��, Decode(n_ʱ�۷���, 1, n_���ۼ�, Null), ����_In,
       ����_In, n_�ϴι�Ӧ��id, n_�ϴβɹ���, d_�ϴ���������, v_��׼�ĺ�, n_�ϴβɹ���);
  End If;

  Delete ҩƷ���
  Where �ⷿid + 0 = Lng�ⷿid And ҩƷid = LngҩƷid And ���� = 1 And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And
        Nvl(ʵ�ʲ��, 0) = 0;

  --�����������
  Zl_ҩƷ�շ���¼_��������(v_������¼id);

  Begin
    --�ƶ�֧������Ŀ�ڷ�ҩ��̬��������������Ϣ�Ĺ���
    Execute Immediate 'Begin zl_������Ϣ_����(:1,:2); End;'
      Using 7, Billid_In || ',' || ��ҩ����_In || ',' || ����_In;
  Exception
    When Others Then
      Null;
  End;

  --��Ϣ����ʣ��ȫ����������0
  If Bln������ҩ = 1 Then
    b_Message.Zlhis_Drug_006(v_������¼id, Lng������, Dblʵ������ * n_����, Lng����id);
  Else
    b_Message.Zlhis_Drug_006(v_������¼id, Lng������, 0, Lng����id);
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�շ���¼_������ҩ;
/

--130781:������,2019-03-18,����Σ��ֵ��¼ɾ��
Create Or Replace Procedure Zl_����Σ��ֵ��¼_Delete(Id_In In ����Σ��ֵ��¼.Id%Type) Is
Begin
  Delete ҵ����Ϣ�嵥
  Where ���ͱ��� = 'ZLHIS_LIS_003' And ҵ���ʶ = (Select To_Char(ҽ��id) From ����Σ��ֵ��¼ Where ID = Id_In);
  Delete ����Σ��ֵ��¼ Where ID = Id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Σ��ֵ��¼_Delete;
/



------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0054' Where ���=&n_System;
Commit;
