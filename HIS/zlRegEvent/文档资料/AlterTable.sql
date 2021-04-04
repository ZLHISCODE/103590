UPDATE �ٴ������Դ set  �Ƿ���ſ��� =NULL,�Ƿ��ʱ��=NULL,ԤԼ����=NULL ,���﷽ʽ=NULL ,����ID=NULL  ;
ALTER TABLE �ٴ������Դ DROP column �Ƿ���ſ���;
ALTER TABLE �ٴ������Դ DROP column �Ƿ��ʱ��;
ALTER TABLE �ٴ������Դ DROP column ԤԼ����;
ALTER TABLE �ٴ������Դ DROP column ���﷽ʽ;
ALTER TABLE �ٴ������Դ DROP column ����ID;

DROP TABLE �ٴ������Դ����;

Create Sequence �ٴ������Դ����_ID start with 1;
Create Table �ٴ������Դ����(
   ID number(18) not null,
   ��ԴID number(18),
   �ϰ�ʱ�� varchar2(10),
   �޺��� number(10),
   ��Լ�� number(10),
   �Ƿ���ſ��� number(2) default 0,
   �Ƿ��ʱ��  NUMBER(2),
   ԤԼ���� number(2),
   �Ƿ��ռ number(2) default 0,   
   ���﷽ʽ number(3),
   ����ID number(18))
TABLESPACE zl9BaseItem ;

Alter Table �ٴ������Դ����  Add Constraint �ٴ������Դ����_PK  Primary Key (ID) Using Index Tablespace zl9Indexhis;
Alter Table �ٴ������Դ���� Add Constraint �ٴ������Դ����_FK_��ԴID Foreign Key (��ԴID) References �ٴ������Դ( ID) ;
Alter Table �ٴ������Դ����  Add Constraint �ٴ������Դ����_UQ_��ԴID  Unique (��ԴID,�ϰ�ʱ��) Using Index Tablespace zl9Indexhis; 
Alter Table �ٴ������Դ���� Add Constraint �ٴ������Դ����_FK_����ID Foreign Key (����ID) References ��������( ID) ;
create Index �ٴ������Դ����_IX_����ID on �ٴ������Դ����(����ID);



Create Table �ٴ������Դ����(
   ����ID number(18),
   ����ID number(18))
TABLESPACE zl9BaseItem ;

Alter Table �ٴ������Դ����  Add Constraint �ٴ������Դ����_PK  Primary Key (����ID,����ID) Using Index Tablespace zl9Indexhis;
Alter Table �ٴ������Դ���� Add Constraint �ٴ������Դ����_FK_����ID Foreign Key (����ID) References �ٴ������Դ����( ID) ;
Alter Table �ٴ������Դ���� Add Constraint �ٴ������Դ����_FK_����ID Foreign Key (����ID) References ��������( ID) ;
create Index �ٴ������Դ����_IX_����ID on �ٴ������Դ����(����ID);

Create Table �ٴ������Դʱ��(
   ����ID number(18),
   ��� number(18),
   ��ʼʱ�� Date,
   ��ֹʱ�� Date,
   �������� number(10),
   �Ƿ�ԤԼ number(2))
TABLESPACE zl9BaseItem;

Alter Table �ٴ������Դʱ��  Add Constraint �ٴ������Դʱ��_PK  Primary Key (����ID,���) Using Index Tablespace zl9Indexhis;
Alter Table �ٴ������Դʱ�� Add Constraint �ٴ������Դʱ��_FK_����ID Foreign Key (����ID) References �ٴ������Դ����( ID) ;



Create Table �ٴ������Դ����(
   ����ID number(18),
   ���� number(2),
   ���� number(2),
   ���� varchar2(50),
   ��� number(18),
   ���Ʒ�ʽ number(2),
   ���� number(16,5))
TABLESPACE zl9BaseItem ;

Alter Table �ٴ������Դ����  Add Constraint �ٴ������Դ����_PK  Primary Key (����ID,����,����,����,���) Using Index Tablespace zl9Indexhis;
Alter Table �ٴ������Դ���� Add Constraint �ٴ������Դ����_FK_����ID Foreign Key (����ID) References �ٴ������Դ����(ID);



Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1114,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All
Select '�ٴ������Դ����_ID','SELECT' From Dual Union All
Select '�ٴ������Դ����','SELECT' From Dual Union All
Select '�ٴ������Դʱ��','SELECT' From Dual Union All
Select '�ٴ������Դ����','SELECT' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;




Insert Into zlParameters(ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
Select Zlparameters_Id.Nextval, &n_System, 1114, A.* From (
  Select ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵�� From zlParameters Where 1 = 0 Union All 
  Select 1, -null, -null, 1, -null, -null, 7, '��ʾȱʡ������Ϣ', '', '1', '�ں�Դ�����п�����ѡ���Դʱ���Ƿ����·���ʾ���Ƶ������Ϣ�����磺ȱʡ�������Ϣ��������Ϣ������ԤԼ������Ϣ�ȣ���', '0-����ʾ��1-��ʾ��', Null, Null, Null From Dual Union All
  Select ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵�� From zlParameters Where 1 = 0) A;





Create Or Replace Procedure Zl_�ٴ������Դ_Delete(Id_In �ٴ������Դ.Id%Type) As
  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  n_Count  Number;
  l_����id t_Numlist := t_Numlist();
Begin
  Select Count(1) Into n_Count From �ٴ����ﰲ�� Where ��Դid = Id_In;

  If n_Count = 0 Then
  
    Select ID Bulk Collect Into l_����id From �ٴ������Դ���� Where ��Դid = Id_In;
  
    Forall I In 1 .. l_����id.Count
      Delete �ٴ������Դʱ�� Where ����id = l_����id(I);
  
    Forall I In 1 .. l_����id.Count
      Delete �ٴ������Դ���� Where ����id = l_����id(I);
  
    Forall I In 1 .. l_����id.Count
      Delete �ٴ������Դ���� Where ����id = l_����id(I);
  
    Delete �ٴ������Դ���� Where ��Դid = Id_In;
    --��ɾ��
  
    Delete From �ٴ������Դ Where ID = Id_In;
    If Sql%NotFound Then
      v_Err_Msg := '��ǰ��Դ�����ѱ�����ɾ����������ɾ��!';
      Raise Err_Item;
    End If;
    Return;
  End If;
  Update �ٴ������Դ Set �Ƿ�ɾ�� = 1, ����ʱ�� = Sysdate Where ID = Id_In And Nvl(�Ƿ�ɾ��, 0) = 0;
  If Sql%NotFound Then
    v_Err_Msg := '��ǰ��Դ�����ѱ�����ɾ����������ɾ��!';
    Raise Err_Item;
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ������Դ_Delete;
/



Create Or Replace Procedure Zl_�ٴ������Դ_Modify
(
  ��������_In     Number,
  Id_In           �ٴ������Դ.Id%Type,
  ����_In         �ٴ������Դ.����%Type := Null,
  ����_In         �ٴ������Դ.����%Type := Null,
  ����id_In       �ٴ������Դ.����id%Type := 0,
  ��Ŀid_In       �ٴ������Դ.��Ŀid%Type := 0,
  ҽ��id_In       �ٴ������Դ.ҽ��id%Type := Null,
  ҽ������_In     �ٴ������Դ.ҽ������%Type := Null,
  �Ƿ񽨲���_In   �ٴ������Դ.�Ƿ񽨲���%Type := 0,
  ԤԼ����_In     �ٴ������Դ.ԤԼ����%Type := 0,
  ����Ƶ��_In     �ٴ������Դ.����Ƶ��%Type := 0,
  ���տ���״̬_In �ٴ������Դ.���տ���״̬%Type := 0,
  �Ƿ���ջ���_In �ٴ������Դ.�Ƿ���ջ���%Type := 0,
  �Ƿ��ٴ��Ű�_In �ٴ������Դ.�Ƿ��ٴ��Ű�%Type := 0,
  �Ű෽ʽ_In     �ٴ������Դ.�Ű෽ʽ%Type := 0
) As
  --��������_In 0-������1-�޸ģ�2-ɾ��
  --��������_In ����ID����ʽ������ID1;����ID2;����ID13;...
  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  n_��Դid �ٴ������Դ.Id%Type;
  n_Count  Number;
Begin

  If ��������_In = 0 Then
    --���Ӻ�Դ
    n_��Դid := Id_In;
  
    If Nvl(n_��Դid, 0) = 0 Then
      Select �ٴ������Դ_Id.Nextval Into n_��Դid From Dual;
    End If;
    Insert Into �ٴ������Դ
      (ID, ����, ����, ����id, ��Ŀid, ҽ��id, ҽ������, �Ƿ񽨲���, ԤԼ����, ����Ƶ��, ���տ���״̬, �Ƿ���ջ���, �Ƿ��ٴ��Ű�, �Ű෽ʽ, �Ƿ�ɾ��, ����ʱ��, ����ʱ��)
    Values
      (n_��Դid, ����_In, ����_In, ����id_In, ��Ŀid_In, ҽ��id_In, ҽ������_In, �Ƿ񽨲���_In, ԤԼ����_In, ����Ƶ��_In, ���տ���״̬_In, �Ƿ���ջ���_In,
       �Ƿ��ٴ��Ű�_In, �Ű෽ʽ_In, 0, Sysdate, To_Date('3000-01-01', 'yyyy-mm-dd'));
  
    Return;
  End If;

  --�޸ĺ�Դ
  Update �ٴ������Դ
  Set ���� = ����_In, ���� = ����_In, ����id = ����id_In, ��Ŀid = ��Ŀid_In, ҽ��id = ҽ��id_In, ҽ������ = ҽ������_In, �Ƿ񽨲��� = �Ƿ񽨲���_In,
      ԤԼ���� = ԤԼ����_In, ����Ƶ�� = ����Ƶ��_In, ���տ���״̬ = ���տ���״̬_In, �Ƿ���ջ��� = �Ƿ���ջ���_In, �Ƿ��ٴ��Ű� = �Ƿ��ٴ��Ű�_In, �Ű෽ʽ = �Ű෽ʽ_In
  Where ID = Id_In And Nvl(�Ƿ�ɾ��, 0) = 0 And Nvl(����ʱ��, Sysdate) >= Sysdate;
  If Sql%NotFound Then
    v_Err_Msg := '��ǰ��Դ�����ѱ�����ɾ����ͣ�ã����ܶԸú�Դ��Ϣ�����޸�!';
    Raise Err_Item;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ������Դ_Modify;
/


Create Or Replace Procedure Zl_�ٴ������Դ����_Modify
(
  Id_In           �ٴ������Դ����.Id%Type,
  ��Դid_In       �ٴ������Դ����.��Դid%Type,
  �ϰ�ʱ��_In     �ٴ������Դ����.�ϰ�ʱ��%Type,
  �޺���_In       �ٴ������Դ����.�޺���%Type,
  ��Լ��_In       �ٴ������Դ����.��Լ��%Type,
  �Ƿ���ſ���_In �ٴ������Դ����.�Ƿ���ſ���%Type,
  �Ƿ��ʱ��_In   �ٴ������Դ����.�Ƿ��ʱ��%Type,
  ԤԼ����_In     �ٴ������Դ����.ԤԼ����%Type,
  �Ƿ��ռ_In     �ٴ������Դ����.�Ƿ��ռ%Type,
  ���﷽ʽ_In     �ٴ������Դ����.���﷽ʽ%Type,
  ����id_In       �ٴ������Դ����.����id%Type,
  ��Դ����_In     Varchar2 := Null,
  ��Դʱ��_In     Varchar2 := Null,
  ��Դ����_In     Varchar2 := Null,
  ɾ����Դ����_In Integer := 0
  
) As
  --��Դʱ��_IN:���,��ʼʱ��(HH:MM:SS),��ֹʱ(HH:MM:SS)��,����,�Ƿ�ԤԼ|...
  --��Դ����_IN:����id1,����id2,....
  --��Դ����_IN:����,����,����,���Ʒ�ʽ,���,����|
  --ɾ����Դ����_in:1-��������ǰ����ɾ����Դ����,0-��ɾ�����ݣ�ֱ�Ӳ���

  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  l_����id   t_Numlist := t_Numlist();
  n_Count    Number;
  v_��ʼʱ�� Varchar2(20);
  v_��ֹʱ�� Varchar2(20);

  n_���     �ٴ������Դʱ��.���%Type;
  d_��ʼʱ�� �ٴ������Դʱ��.��ʼʱ��%Type;
  d_��ֹʱ�� �ٴ������Դʱ��.��ֹʱ��%Type;
  n_����     �ٴ������Դʱ��.��������%Type;
  n_�Ƿ�ԤԼ �ٴ������Դʱ��.�Ƿ�ԤԼ%Type;
  n_����     �ٴ������Դ����.����%Type;
  n_����     �ٴ������Դ����.����%Type;
  v_����     �ٴ������Դ����.����%Type;
  n_���Ʒ�ʽ �ٴ������Դ����.���Ʒ�ʽ%Type;

Begin
  If Nvl(ɾ����Դ����_In, 0) = 1 Then
    Select ID Bulk Collect Into l_����id From �ٴ������Դ���� Where ��Դid = ��Դid_In;
    Forall I In 1 .. l_����id.Count
      Delete �ٴ������Դʱ�� Where ����id = l_����id(I);
  
    Forall I In 1 .. l_����id.Count
      Delete �ٴ������Դ���� Where ����id = l_����id(I);
  
    Forall I In 1 .. l_����id.Count
      Delete �ٴ������Դ���� Where ����id = l_����id(I);
  
    Delete �ٴ������Դ���� Where ��Դid = ��Դid_In;
    Delete From �ٴ������Դ���� Where ��Դid = ��Դid_In;
  
  End If;

  Select Count(1) Into n_Count From �ٴ������Դ���� Where ID = Id_In;
  If n_Count = 0 Then
    Insert Into �ٴ������Դ����
      (ID, ��Դid, �ϰ�ʱ��, �޺���, ��Լ��, �Ƿ���ſ���, �Ƿ��ʱ��, ԤԼ����, �Ƿ��ռ, ���﷽ʽ, ����id)
    Values
      (Id_In, ��Դid_In, �ϰ�ʱ��_In, �޺���_In, ��Լ��_In, �Ƿ���ſ���_In, �Ƿ��ʱ��_In, ԤԼ����_In, �Ƿ��ռ_In, ���﷽ʽ_In, ����id_In);
  
  End If;

  If ��Դʱ��_In Is Not Null Then
    --�����Դȱʡʱ���
    For c_ʱ��μ� In (Select Rownum As ���, Column_Value As ֵ From Table(f_Str2list(��Դʱ��_In, '|'))) Loop
      n_���     := Null;
      v_��ʼʱ�� := Null;
      v_��ֹʱ�� := Null;
      n_����     := Null;
      n_�Ƿ�ԤԼ := Null;
      For c_ʱ��� In (Select Rownum As ���, Column_Value As ֵ From Table(f_Str2list(c_ʱ��μ�.ֵ)) Order By ���) Loop
        If c_ʱ���.��� = 1 Then
          n_��� := To_Number(c_ʱ���.ֵ);
        End If;
      
        If c_ʱ���.��� = 2 Then
          v_��ʼʱ�� := c_ʱ���.ֵ;
        End If;
      
        If c_ʱ���.��� = 3 Then
          v_��ֹʱ�� := c_ʱ���.ֵ;
        End If;
      
        If c_ʱ���.��� = 4 Then
          n_���� := To_Number(c_ʱ���.ֵ);
        End If;
      
        If c_ʱ���.��� = 5 Then
          n_�Ƿ�ԤԼ := To_Number(c_ʱ���.ֵ);
        End If;
      
      End Loop;
      d_��ʼʱ�� := To_Date('3000-01-01 ' || Nvl(v_��ʼʱ��, ''), 'yyyy-mm-dd hh24:mi:ss');
      d_��ֹʱ�� := To_Date('3000-01-01 ' || Nvl(v_��ֹʱ��, ''), 'yyyy-mm-dd hh24:mi:ss');
    
      If d_��ʼʱ�� >= d_��ʼʱ�� Then
        d_��ֹʱ�� := d_��ֹʱ�� + 1;
      End If;
    
      If Nvl(n_���, 0) <> 0 Then
        Insert Into �ٴ������Դʱ��
          (����id, ���, ��ʼʱ��, ��ֹʱ��, ��������, �Ƿ�ԤԼ)
        Values
          (Id_In, n_���, d_��ʼʱ��, d_��ֹʱ��, n_����, n_�Ƿ�ԤԼ);
      End If;
    End Loop;
  
  End If;

  --�����Դ��ȱʡ����
  --��Դ����_IN:����,����,����,���Ʒ�ʽ,���,����|
  If ��Դ����_In Is Not Null Then
    For c_ʱ��μ� In (Select Rownum As ���, Column_Value As ֵ From Table(f_Str2list(��Դ����_In, '|'))) Loop
      n_����     := Null;
      n_����     := Null;
      v_����     := Null;
      n_���     := Null;
      n_���Ʒ�ʽ := Null;
      n_����     := Null;
    
      --����,����,����,���Ʒ�ʽ,���,����|
      For c_ʱ��� In (Select Rownum As ���, Column_Value As ֵ From Table(f_Str2list(c_ʱ��μ�.ֵ)) Order By ���) Loop
        If c_ʱ���.��� = 1 Then
          n_���� := To_Number(c_ʱ���.ֵ);
        End If;
      
        If c_ʱ���.��� = 2 Then
          n_���� := To_Number(c_ʱ���.ֵ);
        End If;
      
        If c_ʱ���.��� = 3 Then
          v_���� := c_ʱ���.ֵ;
        End If;
      
        If c_ʱ���.��� = 4 Then
          n_���Ʒ�ʽ := To_Number(c_ʱ���.ֵ);
        End If;
      
        If c_ʱ���.��� = 5 Then
          n_��� := To_Number(c_ʱ���.ֵ);
        End If;
      
        If c_ʱ���.��� = 6 Then
          n_���� := To_Number(c_ʱ���.ֵ);
        End If;
      
      End Loop;
    
      If v_���� Is Not Null Then
        Insert Into �ٴ������Դ����
          (����id, ����, ����, ����, ���, ���Ʒ�ʽ, ����)
        Values
          (Id_In, n_����, n_����, v_����, n_���, n_���Ʒ�ʽ, n_����);
      
      End If;
    End Loop;
  End If;
  --�����Դ����
  If ��Դ����_In Is Not Null Then
    Insert Into �ٴ������Դ����
      (����id, ����id)
      Select Id_In As ����id, Column_Value As ����id From Table(f_Num2list(��Դ����_In));
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ������Դ����_Modify;
/
