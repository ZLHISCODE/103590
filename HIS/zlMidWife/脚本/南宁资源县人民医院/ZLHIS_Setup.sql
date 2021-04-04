--��zlhis�û���¼����ʿ�����ݿ⣬
--����dblink����
create database link ZLSOL_DBL  connect to ZLSOL identified by &zlsol��������ʿ�������  using '(DESCRIPTION =
    (ADDRESS = (PROTOCOL = TCP)(HOST = &����ʿ��IP)(PORT = &����ʿ��˿�))
    (CONNECT_DATA =
      (SERVICE_NAME = &����ʿ��ʵ����)
    )
  )';
--������Ժ�Ĳ��Ʋ���������ʿ����վ��
Insert into Sol_Inf_Puerpera@ZLSOL_DBL( Pid, Tid, Name, Old, Bedno, Pno, Diagnosis, Status, Outtime) 
Select b.����id, b.��ҳid, a.����, a.����, a.��ǰ����, a.סԺ��, c.�������, 0 As Status,To_Date('3000-01-01', 'yyyy-mm-dd')
From ������Ϣ A, ��Ժ���� B, ������ϼ�¼ C
Where a.����id = b.����id And a.��ҳid = b.��ҳid And b.����id = &����id And b.����id = c.����id(+) And b.��ҳid = c.��ҳid(+)��
--����ͬ��������������
CREATE OR REPLACE Trigger t_Apex_����״̬ͬ��
  After Insert Or Delete Or Update On ���˱䶯��¼
  For Each Row
Declare
  n_����id  ���˱䶯��¼.����id%Type;
  v_Err_Msg Varchar2(255);
  Err_Item Exception; --1.��ơ���Ժ��ơ�ת����� 2.������� 3.��Ժ 4.������Ժ 5������ 6.��������
Begin
  If Inserting Then
    --��Ժ���
    If :New.��ʼԭ�� = 1 And Nvl(:New.����, 0) <> 0 Then
      Zl_Apex_����״̬ͬ��(:New.����id, :New.��ҳid, :New.����id, 1, :New.����);
    End If ;
    --��ƣ�ת�����
    If :New.��ʼԭ�� In (2, 3) And :New.���Ӵ�λ = 0 Then
      Zl_Apex_����״̬ͬ��(:New.����id, :New.��ҳid, :New.����id, 1, :New.����);
    End If;
    --����
    If :New.��ʼԭ�� = 4 Then
      Zl_Apex_����״̬ͬ��(:New.����id, :New.��ҳid, :New.����id, 5, :New.����);
    End If;
  Elsif Deleting Then
    --�������
    If :Old.��ʼԭ�� In (2, 3) Then
      Zl_Apex_����״̬ͬ��(:Old.����id, :Old.��ҳid, :Old.����id, 2);
    End If;
    --��������
    If :Old.��ʼԭ�� = 4 Then
      Zl_Apex_����״̬ͬ��(:Old.����id, :Old.��ҳid, :Old.����id, 6, :New.����);
    End If;
  Elsif Updating Then
    --��Ժ
    If :New.��ֹԭ�� = 1 And :Old.��ֹʱ�� Is Null And :Old.���Ӵ�λ = 0 Then
      Zl_Apex_����״̬ͬ��(:New.����id, :New.��ҳid, :New.����id, 3, :New.����, :New.��ֹʱ��);
    End If;
    --������Ժ
    If :Old.��ֹԭ�� = 1 And :New.��ֹʱ�� Is Null Then
      Zl_Apex_����״̬ͬ��(:New.����id, :New.��ҳid, :New.����id, 4);
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End t_Apex_����״̬ͬ��;

/


CREATE OR REPLACE Procedure Zl_Apex_����״̬ͬ��
(
  ����id_In   ���˱䶯��¼.����id%Type,
  ��ҳid_In   ���˱䶯��¼.��ҳid%Type,
  ����id_In   ���˱䶯��¼.����id%Type,
  ��������_In Number, --1.��ơ���Ժ��ơ�ת����� 2.������� 3.��Ժ 4.������Ժ 5������ 6.��������
  ����_In     ���˱䶯��¼.����%Type:= Null,
  ��Ժʱ��_In ���˱䶯��¼.��ֹʱ��%Type := Null
) As
  n_Count   Number(5);
  v_Bedno   Varchar(10);
  v_Err_Msg Varchar2(255);
  Err_Item Exception;
Begin
  If ����id_In In (205) Then
    --1.��ơ�ת����ס����Ժ���
    If ��������_In = 1 Then
      Select Count(1) Into n_Count From Sol_Inf_Puerpera@Zlsol_Dbl Where Pid = ����id_In And Tid = ��ҳid_In;

      If n_Count = 0 Then
        Insert Into Sol_Inf_Puerpera@Zlsol_Dbl
          (Pid, Tid, Name, Old, Bedno, Pno, Diagnosis, Status, Outtime)
          Select a.����id, a.��ҳid, a.����, a.����, a.��ǰ����, a.סԺ��, c.�������, 0, To_Date('3000-01-01', 'yyyy-mm-dd')
          From ������Ϣ a, ������ϼ�¼ c
          Where a.����id = ����id_In And a.��ҳid = ��ҳid_In And a.����id = c.����id(+) And a.��ҳid = c.��ҳid(+);
      Else
        --ת��ʱ�Ѵ��ڲ���
        Update Sol_Inf_Puerpera@Zlsol_Dbl a Set Status = 0 Where a.Pid = ����id_In And a.Tid = ��ҳid_In;
      End If;
    --2.�������
    Elsif ��������_In = 2 Then
      Update Sol_Inf_Puerpera@Zlsol_Dbl a Set Status = 4 Where a.Pid = ����id_In And a.Tid = ��ҳid_In;
    --3.��Ժ
    Elsif ��������_In = 3 Then
      Update Sol_Inf_Puerpera@Zlsol_Dbl a
      Set Status = 3, Outtime = ��Ժʱ��_In
      Where a.Pid = ����id_In And a.Tid = ��ҳid_In;
    --4.������Ժ
    Elsif ��������_In = 4 Then
      Update Sol_Inf_Puerpera@Zlsol_Dbl a
      Set Status = 2, Outtime = To_Date('3000-01-01', 'yyyy-mm-dd')
      Where a.Pid = ����id_In And a.Tid = ��ҳid_In And a.Status = 3;
    --5.����
    Elsif ��������_In = 5 Then
      Update Sol_Inf_Puerpera@Zlsol_Dbl a
      Set Bedno = Nvl(����_In, '��ͥ����')
      Where a.Pid = ����id_In And a.Tid = ��ҳid_In;
    --6.��������
    Elsif ��������_In = 6 Then
      Select ��ǰ���� Into v_Bedno From ������Ϣ Where ����id = ����id_In And ��ҳid = ��ҳid_In;
      Update Sol_Inf_Puerpera@Zlsol_Dbl a Set Bedno = v_Bedno Where a.Pid = ����id_In And a.Tid = ��ҳid_In;
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Apex_����״̬ͬ��;




/
--�������ͬ��
CREATE OR REPLACE Trigger t_Apex_����״̬ͬ��_���
  After Insert Or Delete Or Update On ������ϼ�¼
  For Each Row
Declare

  v_Err_Msg Varchar2(255);
  Err_Item Exception;
Begin
  --�������޸�
  If Inserting Or Updating Then
    If :New.��¼��Դ In (2, 3) And :New.������� = 2 And :New.��ϴ��� = 1 Then
      Zl_Apex_����״̬ͬ��_���(:New.����id, :New.��ҳid, 1, :New.�������);
    End If;
    --ɾ��
  Elsif Deleting Then
    If :Old.��¼��Դ In (2, 3) And :New.������� = 2 And :New.��ϴ��� = 1 Then
      Zl_Apex_����״̬ͬ��_���(:Old.����id, :Old.��ҳid, 2);
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End t_Apex_����״̬ͬ��_���;

/





CREATE OR REPLACE Procedure Zl_Apex_����״̬ͬ��_���
(
  ����id_In   ������ϼ�¼.����id%Type,
  ��ҳid_In   ������ϼ�¼.��ҳid%Type,
  ��������_In Number, --1.�޸���ϣ�2.ɾ�����
  �������_In ������ϼ�¼.�������%Type := Null
) As
  n_Count Number(5);
Begin
  Select Count(1)
  Into n_Count
  From ��Ժ���� B, ��������˵�� C
  Where b.����id = c.����id And c.�������� = '����' And b.����id = ����id_In And b.��ҳid = ��ҳid_In;
  If n_Count = 1 Then
    If ��������_In = 1 Then
      Update Sol_Inf_Puerpera@Zlsol_Dbl Set Diagnosis = �������_In Where Pid = ����id_In And Tid = ��ҳid_In;
    Elsif ��������_In = 2 Then
      update Sol_Inf_Puerpera@Zlsol_Dbl Set Diagnosis = Null Where Pid = ����id_In And Tid = ��ҳid_In;
    End If;
  End If;
End Zl_Apex_����״̬ͬ��_���;
/





--����ʿ�û�ͬ��������
--ע�⣺zl_��Ա��_delete ��������Ҫ��delete from...�����һ�� Delete From Sol_User@Zlsol_Dbl Where Code = v_User;
CREATE OR REPLACE Trigger t_Apex_��Ա�䶯
  After Insert Or Delete Or Update On �ϻ���Ա��
  For Each Row
Declare
  v_Name    Varchar2(20);
  n_Count   Number(5);
  v_Err_Msg Varchar2(255);
  Err_Item Exception; ----1.������Ա��2.ɾ����Ա��3.�޸���Ա��ʵ������ɾ��������
Begin
  If Inserting Then
    --������Ա
    Select Max(Distinct b.����) ����
    Into v_Name
    From ������Ա A, ��Ա�� B
    Where a.��Աid = b.Id And a.��Աid = :New.��Աid And a.����id = &�贴������ʿ�û��Ŀ���ID;
    If v_Name Is Not Null Then
      Insert Into Sol_User@Zlsol_Dbl (Code, Name, State) Values (:New.�û���, v_Name, 1);
    End If;
    --ɾ����Ա
  Elsif Deleting Then
    Select Count(*) Into n_Count From ������Ա A Where a.��Աid = :Old.��Աid And a.����id = &�贴������ʿ�û��Ŀ���ID;
    If n_Count > 0 Then
      Delete From Sol_User@Zlsol_Dbl Where Code = :Old.�û���;
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End t_Apex_��Ա�䶯;
/
CREATE OR REPLACE Trigger t_Apex_��Ա����
  After Insert Or Delete Or Update On ��Ա��
  For Each Row
Declare
  v_Code    Varchar2(20);
Begin
  If Updating Then    
    Select max(Distinct b.�û���) �û���
    Into v_Code
    From ������Ա a, �ϻ���Ա�� b
    Where a.��Աid = b.��Աid And a.��Աid = :Old.Id And a.����id = &�贴������ʿ�û��Ŀ���ID;
    If v_Code Is Not Null Then
      --�޸�����
      If :New.���� <> :Old.���� Then
        Update Sol_User@Zlsol_Dbl Set Name = :New.���� Where Code = v_Code;
      End If;
      --���á�ͣ���û�
      If :New.����ʱ�� = To_Date('3000-1-1', 'yyyy-mm-dd') Then
        Update Sol_User@Zlsol_Dbl Set State = 1 Where Code = v_Code;
      Else
        Update Sol_User@Zlsol_Dbl Set State = 0 Where Code = v_Code;
      End If;
    End If;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End t_Apex_��Ա����;
/