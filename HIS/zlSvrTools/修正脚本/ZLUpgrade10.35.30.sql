----10.35.20---��10.35.30
--100680:��˶,2016-10-13,������������쳣
Create Or Replace Procedure ZLTOOLS.Zl_zluserparas_Clear
(
  �û���_In In zluserparas.�û���%Type,
  ������_In In zluserparas.������%Type
) Is
Begin
  Delete From zluserparas Where �û���=�û���_In or ������=������_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_zluserparas_Clear;
/
--100258:������,2016-09-05,���zlReports��������
Alter Table zlTools.zlReports Add (ִ����Ա Varchar2(20), ���ִ��ʱ�� Date);
Alter Table zlTools.Zlrptrunhistory Add (ִ����Ա Varchar2(20));

Begin  
  For r_Alter In (Select 'alter index ' || Index_Name || ' initrans 20' Line_
                  From All_Indexes
                  Where Table_Name = 'ZLREPORTS' And ini_trans < 20) Loop
    Execute Immediate r_Alter.Line_;
  End Loop;
End;
/

Alter Table zlTools.zlReports Move;
Begin  
  For r_Alter In (Select 'alter index ' || Index_Name || ' rebuild' Line_
                  From All_Indexes
                  Where Table_Name = 'ZLREPORTS') Loop
    Execute Immediate r_Alter.Line_;
  End Loop;
End;
/

--100258:������,2016-09-05,���zlReports��������
CREATE OR REPLACE Procedure zlTools.Zl_Rptrun_Update
(
  Id_In       In Zlreports.Id%Type,
  ִ����Ա_In In Zlreports.ִ����Ա%Type
) Is
  Pragma Autonomous_Transaction;
  n_Count Number(1);
Begin
  Select Nvl(Count(1), 0)
  Into n_Count
  From zlReports
  Where ID = Id_In And (���ִ��ʱ�� < Sysdate - 5 / 24 / 60 Or ���ִ��ʱ�� Is Null);

  If n_Count > 0 Then
    Update zlReports Set ִ����Ա = ִ����Ա_In, ���ִ��ʱ�� = Sysdate Where ID = Id_In;
    Commit;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Rptrun_Update;
/

--100258:������,2016-09-05,���zlReports��������
Create Or Replace Procedure zlTools.Zl_Rptrunhistory_Update
(
  ����id_In       In Zlreports.Id%Type,
  ִ����Ա_In     In Zlrptrunhistory.ִ����Ա%Type,
  ִ�п�ʼʱ��_In In Zlrptrunhistory.ִ�п�ʼʱ��%Type,
  ִ�н���ʱ��_In In Zlrptrunhistory.ִ�н���ʱ��%Type
) Is
Begin
  Insert Into Zlrptrunhistory
    (ID, ����id, ִ����Ա, ִ�п�ʼʱ��, ִ�н���ʱ��)
  Values
    (Zlrptrunhistory_Id.Nextval, ����id_In, ִ����Ա_In, ִ�п�ʼʱ��_In, ִ�н���ʱ��_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Rptrunhistory_Update;
/

--100994:������,2016-09-30,��ϪҽԺ��ʷ����ת�����Լ����޸�
Alter Table zltools.Zldatamovelog modify ��ǰ���� varchar2(200);

--100994:������,2016-09-30,��ϪҽԺ��ʷ����ת�����Լ����޸�
Create Table zltools.zlBakTableIndex(
	ϵͳ NUMBER(5) ,
	���� Varchar2(30),
	������ VARCHAR2(30));
Alter Table zltools.zlBakTableIndex Add Constraint zlBakTableIndex_PK Primary Key (ϵͳ,����,������) USING INDEX PCTFREE 5;
Alter Table zltools.zlBakTableIndex Add Constraint zlBakTableIndex_FK_ϵͳ Foreign Key (ϵͳ,����) References zlBakTables(ϵͳ,����) On Delete Cascade;

--100412:��˶,2016-10-26,�Զ������Ľ�
Alter Table zltools.zlFilesUpgrade Drop Constraint zlFilesUpgrade_CK_�Զ�ע��;
Alter Table zltools.ZLCLIENTS Drop Constraint ZLCLIENTS_CK_Ԥ�����;
Alter Table zltools.zlFilesUpgrade Modify �ļ��� Varchar2(100);
--100412:��˶,2016-10-26,�Զ������Ľ�
create table ZLTOOLS.ZLKillProcess
(
���     number(5),
����     varchar2(50),
����     number(1),--0-���̣�1-����
����     varchar2(200)
);
alter table  ZLTOOLS.ZLKillProcess add constraint ZLKillProcess_UQ_���� unique(����) using index;


--101981:��˶,2016-10-27,���Ԥ����ʱ�㱨��
Create Or Replace Procedure zltools.Zl_Zlclients_Control
(
  n_Mode_In       Number,
  v_����վ_In     Zlclients.����վ%Type := Null,
  v_Ip_In         Zlclients.Ip%Type := Null,
  n_������־_In   Zlclients.������־%Type := Null,
  n_����������_In Zlclients.����������%Type := Null,
  d_Ԥ��ʱ��_In   Zlclients.Ԥ��ʱ��%Type := Null,
  n_Ԥ�����_In   Zlclients.Ԥ�����%Type := Null,
  n_Ftp������_In  Zlclients.Ftp������%Type := Null,
  n_�ռ���־_In   Zlclients.�ռ���־%Type := Null,
  n_��ֹʹ��_In   Zlclients.��ֹʹ��%Type := Null,
  v_˵��_In       Zlclients.˵��%Type := Null
  --�Կͻ��˽��п���
  --N_Mode_In��0-���û����ÿͻ���(IP��Ϊ��Ҫ������,1-Ԥ��������,2 -������Ϣ����(IP��Ϊ��Ҫ������
  --3-ȡ��Ԥ������־,4-������վ������Ϊ����,5-�����Ѽ��������Ѽ���־��,6-��������״̬
) Is
  v_Timeset Varchar2(300);
  v_Err     Varchar2(500);
  Err_Custom Exception;
Begin
  --0-���û����ÿͻ���(IP��Ϊ��Ҫ������
  If n_Mode_In = 0 Then
    If v_����վ_In Is Not Null Then
      Update zlClients Set ��ֹʹ�� = n_��ֹʹ��_In Where Ip = v_Ip_In;
    End If;
    --1-Ԥ��������,����Ҫ����������
  Elsif n_Mode_In = 1 Then
    Select Max(����) Into v_Timeset From zlRegInfo Where ��Ŀ = '�ͻ���Ԥ����ʱ���';
    If v_Timeset Is Not Null Then
      For r_Ip In (Select To_Date(Today || ' ' || Date_d, 'yyyy-mm-dd HH24:mi:ss') Ԥ��ʱ��, ����վ, Ip
                   From (Select ����վ, Ip, Rownum Rn_c From zlClients) A,
                        (Select To_Char(Sysdate, 'yyyy-mm-dd') Today, Column_Value Date_d, Rownum Rn_d, Count(1) Over() Sn
                          From Table(f_Str2list(v_Timeset, ','))) B
                   Where Mod(a.Rn_c, Sn) + 1 = Rn_d) Loop
      
        Update zlClients Set Ԥ��ʱ�� = r_Ip.Ԥ��ʱ�� Where ����վ = r_Ip.����վ And Ip = r_Ip.Ip;
      End Loop;
    Else
      Update zlClients Set Ԥ��ʱ�� = NuLL;
    End If;
    --2 -������Ϣ����(IP��Ϊ��Ҫ������
  Elsif n_Mode_In = 2 Then
    If n_Ftp������_In Is Null Then
      Update zlClients
      Set ������־ = n_������־_In, ���������� = n_����������_In, Ԥ��ʱ�� = d_Ԥ��ʱ��_In, Ԥ����� = n_Ԥ�����_In
      Where Ip = v_Ip_In;
    
    Else
      Update zlClients
      Set ������־ = n_������־_In, Ftp������ = n_Ftp������_In, Ԥ��ʱ�� = d_Ԥ��ʱ��_In, Ԥ����� = n_Ԥ�����_In
      Where Ip = v_Ip_In;
    End If;
    --3-ȡ��Ԥ������־
  Elsif n_Mode_In = 3 Then
    Update zlClients Set Ԥ����� = n_Ԥ�����_In;
    --4-������վ������Ϊ����
  Elsif n_Mode_In = 4 Then
    Update zlClients Set ������־ = n_������־_In;
    --5-�����Ѽ��������Ѽ���־��
  Elsif n_Mode_In = 5 Then
    If v_����վ_In Is Null Then
      Update zlClients Set �ռ���־ = n_�ռ���־_In;
    Else
      Update zlClients Set �ռ���־ = n_�ռ���־_In Where ����վ = v_����վ_In;
    End If;
  Elsif n_Mode_In = 6 Then
    Update zlClients Set ������� = 0 Where ����վ = v_����վ_In;
  Elsif n_Mode_In = 7 Then
    --7δ����
    Update zlClients Set ������� = 1 Where ����վ = v_����վ_In;
  Elsif n_Mode_In = 8 Then
    --8������
    Update zlClients Set ������� = 2 Where ����վ = v_����վ_In;
  Elsif n_Mode_In = 9 Then
    --9�޸�˵��
    Update zlClients Set ˵�� = v_˵��_In Where ����վ = v_����վ_In;
  Elsif n_Mode_In = 10 Then
    --10�޸�˵�����ռ���־
    Update zlClients Set ˵�� = v_˵��_In, �ռ���־ = 0 Where ����վ = v_����վ_In;
  Elsif n_Mode_In = 11 Then
    --11�޸�˵����������־
    Update zlClients Set ˵�� = v_˵��_In, ������־ = 0 Where ����վ = v_����վ_In;
  Elsif n_Mode_In = 12 Then
    --12����վ���Ԥ��������״̬
    Update zlClients Set ˵�� = v_˵��_In, Ԥ����� = 0 Where ����վ = v_����վ_In;
  Elsif n_Mode_In = 13 Then
    --13����վ���Ԥ�������״̬
    Update zlClients Set Ԥ����� = 1 Where ����վ = v_����վ_In;
  Elsif n_Mode_In = 14 Then
    --14��ʱ����
    Update zlClients Set Ԥ��ʱ�� = Null, Ԥ����� = Null Where ����վ = v_����վ_In;
  Elsif n_Mode_In = 15 Then
    --15�����ɹ�
    Update zlClients Set ������� = 1 Where ����վ = v_����վ_In;
  Elsif n_Mode_In = 16 Then
    --16����ʧ��
    Update zlClients Set ������� = 2 Where ����վ = v_����վ_In;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlclients_Control;
/
