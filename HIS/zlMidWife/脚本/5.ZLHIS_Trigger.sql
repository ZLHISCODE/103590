--��ZLHIS��ִ�У��޸�zlsol�����룬ip��ʵ����[SERVICE_NAME]��
create database link ZLSOL_DBL  connect to ZLSOL identified by ZLSOL_PASSWORD  using '(DESCRIPTION =
    (ADDRESS = (PROTOCOL = TCP)(HOST = 192.168.0.60)(PORT = 1521))
    (CONNECT_DATA =
      (SERVICE_NAME = orcl12)
    )
  )';


--������Ժ��������(�޸Ĳ���ID����7748)
insert into Sol_Inf_Puerpera@ZLSOL_DBL( Pid, Tid, Name, Old, Bedno, Pno, Diagnosis, Status, Outtime) 
Select b.����id, b.��ҳid, a.����, a.����, a.��ǰ����, a.סԺ��, c.�������, 1 As Status,To_Date('3000-01-01', 'yyyy-mm-dd')
From ������Ϣ A, ��Ժ���� B, ������ϼ�¼ C
Where a.����id = b.����id And a.��ҳid = b.��ҳid And b.����id = 7748 And b.����id = c.����id(+) And b.��ҳid = c.��ҳid(+) And c.��¼��Դ(+) = 3 And
    c.�������(+) = 2 And c.��ϴ���(+) = 1

CREATE OR REPLACE Procedure Zl_Apex_����״̬ͬ��
(
  ����id_In   ���˱䶯��¼.����id%Type,
  ��ҳid_In   ���˱䶯��¼.��ҳid%Type,
  ����id_In   ���˱䶯��¼.����id%Type,
  ��������_In Number, --1.��ƣ�2.������ƣ�3.��Ժ,4-������Ժ,5���� 6.��������,7.��Ժ��ϸ���,8,ɾ����Ժ���
  ����_in     ���˱䶯��¼.����%Type,
  ��Ժʱ��_In ���˱䶯��¼.��ֹʱ��%Type := Null
) As
  n_Count Number(5);
  n_Mid   Number(18);
  v_bedno varchar(10);
Begin
  Select Count(1) Into n_Count From ��������˵�� Where ����id = ����id_In And �������� = '����';

  If n_Count = 1 Then
    --1.���
    If ��������_In = 1 Then
      Select Count(1) Into n_Count From Sol_Inf_Puerpera@ZLSOL_DBL Where Pid = ����id_In And Tid = ��ҳid_In;

      If n_Count = 0 Then
        Insert Into Sol_Inf_Puerpera@ZLSOL_DBL
          (Pid, Tid, Name, Old, Bedno, Pno, Diagnosis,status,outtime)
          Select a.����id, a.��ҳid, a.����, a.����, a.��ǰ����, a.סԺ��, c.�������,0,To_Date('3000-01-01', 'yyyy-mm-dd')
          From ������Ϣ A, ������ϼ�¼ C
          Where a.����id = ����id_In And a.��ҳid = ��ҳid_In And a.����id = c.����id(+) And a.��ҳid = c.��ҳid(+) And c.��¼��Դ(+) = 3 And
                c.�������(+) = 2 And c.��ϴ���(+) = 1;
      End If;
    Elsif ��������_In = 2 Then
      Select Nvl(Max(Mid), 0)
      Into n_Mid
      From Sol_Inf_Puerpera@ZLSOL_DBL A
      Where a.Pid = ����id_In And a.Tid = ��ҳid_In And a.Status = 0;

      --δ�뷿������δ��д������¼���ٲ���¼����ɾ��,������¼�����Ѵ��ڣ��ɲ�����д
      If n_Mid > 0 Then
        /*Select Count(1) Into n_Count From Sol_Rs_Expectant@ZLSOL_DBL Where Mid = n_Mid;
        If n_Count = 0 Then
          Select Count(1) Into n_Count From Sol_Rs_Birth@ZLSOL_DBL Where Mid = n_Mid;
          If n_Count = 0 Then*/
            Delete Sol_Inf_Puerpera@ZLSOL_DBL Where Mid = n_Mid;
         /* End If;
        End If;*/
      End If;

    Elsif ��������_In = 3 Then
      --��Ժ
      Update Sol_Inf_Puerpera@ZLSOL_DBL A
      Set Status = 3, Outtime = ��Ժʱ��_In
      Where a.Pid = ����id_In And a.Tid = ��ҳid_In;

    Elsif ��������_In = 4 Then
      --������Ժ
      Update Sol_Inf_Puerpera@ZLSOL_DBL A
      Set Status = 2, Outtime = To_Date('3000-01-01', 'yyyy-mm-dd')
      Where a.Pid = ����id_In And a.Tid = ��ҳid_In And a.Status = 3;
    Elsif ��������_In = 5 Then
      --����
      Update Sol_Inf_Puerpera@ZLSOL_DBL A
      Set  bedno=nvl(����_in,'��ͥ����')
      Where a.Pid = ����id_In And a.Tid = ��ҳid_In;
    Elsif ��������_In = 6 Then
      --��������
      select ��ǰ���� into v_bedno from ������Ϣ where ����id = ����id_In And ��ҳid = ��ҳid_In;
      Update Sol_Inf_Puerpera@ZLSOL_DBL A
      Set bedno=v_bedno
      Where a.Pid = ����id_In And a.Tid = ��ҳid_In ;
    End If;
  End If;
End Zl_Apex_����״̬ͬ��;
/
CREATE OR REPLACE Trigger t_Apex_����״̬ͬ��
  After Insert Or Delete Or Update On ���˱䶯��¼
  For Each Row
Declare

  v_Err_Msg Varchar2(255);
  Err_Item Exception;  ----1.��� 2��������� 3.��Ժ 4.������Ժ 5������ 6.��������
Begin

  If Inserting Then
    --��Ժ���
    If :New.��ʼԭ�� =1 And Nvl(:new.����,0) <>0Then
      Zl_Apex_����״̬ͬ��(:New.����id, :New.��ҳid, :New.����id, 1,:new.����);
    End If;    
    --���
    If :New.��ʼԭ�� In (2, 3) And :New.���Ӵ�λ = 0 Then
      Zl_Apex_����״̬ͬ��(:New.����id, :New.��ҳid, :New.����id, 1,:new.����);
    End If;
    -----����
    If :New.��ʼԭ�� =4 Then
      Zl_Apex_����״̬ͬ��(:New.����id, :New.��ҳid, :New.����id, 5,:new.����);
    End If;
  Elsif Deleting Then
    --�������
    If :Old.��ʼԭ��  In (2, 3) Then
      Zl_Apex_����״̬ͬ��(:Old.����id, :Old.��ҳid, :Old.����id, 2,:new.����);
    End If;
    --��������
    If :Old.��ʼԭ�� = 4 Then
      Zl_Apex_����״̬ͬ��(:Old.����id, :Old.��ҳid, :Old.����id, 6,:new.����);
    End If;
    --��Ժ
  Elsif Updating Then
    If :New.��ֹԭ�� = 1 And :Old.��ֹʱ�� Is Null And :Old.���Ӵ�λ = 0 Then
      Zl_Apex_����״̬ͬ��(:New.����id, :New.��ҳid, :New.����id, 3,:new.����, :New.��ֹʱ��);
    End If;
    --������Ժ
    If :Old.��ֹԭ�� = 1 And :New.��ֹʱ�� Is Null Then
      Zl_Apex_����״̬ͬ��(:New.����id, :New.��ҳid, :New.����id, 4,:new.����);
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End t_Apex_����״̬ͬ��;
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
  From ������ϼ�¼ a, ��Ժ���� b, ��������˵�� c
  Where a.����id = b.����id And a.��ҳid = b.��ҳid And b.����id = c.����id And c.�������� = '����' And a.����id = ����id_In And
        a.��ҳid = ��ҳid_In;
  If n_Count = 1 Then
    If ��������_In = 1 Then
      Update Sol_Inf_Puerpera@ZLSOL_DBL Set Diagnosis = �������_In Where Pid = ����id_In And Tid = ��ҳid_In;
    Elsif ��������_In = 2 Then
      Update Sol_Inf_Puerpera@ZLSOL_DBL Set Diagnosis = Null Where Pid = ����id_In And Tid = ��ҳid_In;
    End If;
  End If;
End Zl_Apex_����״̬ͬ��_���;
/
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