----10.35.70---��10.35.80
--00000:��˶,2017-09-25,Ψһ�����������Ϲ淶����
alter table Zltools.zlProgFuncs rename constraint zlProgFuncs_PK to zlProgFuncs_UQ_����;
alter index Zltools.zlProgFuncs_PK rename to zlProgFuncs_UQ_����;
Alter Table Zltools.zlProgFuncs Modify ���  constraint zlProgFuncs_NN_���   not  null;
alter table Zltools.zlProgPrivs rename constraint zlProgPrivs_PK to zlProgPrivs_UQ_Ȩ��;
alter index Zltools.zlProgPrivs_PK rename to zlProgPrivs_UQ_Ȩ��;
Alter Table Zltools.zlProgPrivs Modify ���  constraint zlProgPrivs_NN_���   not  null;
alter table Zltools.zlPrograms rename constraint zlPrograms_PK to zlPrograms_UQ_���;
alter index Zltools.zlPrograms_PK rename to zlPrograms_UQ_���;
Alter Table Zltools.zlPrograms Modify ���  constraint zlPrograms_NN_���   not  null;

--111526:����,2017-9-27,������־
Alter Table Zltools.Zllogconfig Modify(���� Varchar2(50));
Alter Table Zltools.Zllogconfig Add(ϵͳ Number(5));
Alter Table Zltools.Zllogconfig Drop Primary Key;
Alter Table Zltools.Zllogconfig Add Constraint Zllogconfig_Pk Primary Key(����) Using Index;
Alter Table Zltools.Zllogconfig Add Constraint Zllogconfig_Uq_��� Unique(ϵͳ, ���) Using Index;
Alter Table Zltools.Zllogconfig Add Constraint Zllogconfig_Fk_ϵͳ Foreign Key(ϵͳ) References Zlsystems(���) On Delete Cascade;

--114503:����,2017-9-25,Ժ����������޸�
Update Zltools.Zlnodelist Set ���� = Decode(����,'վ��0','��Ժ','վ��1','һ��Ժ','վ��2','����Ժ','վ��3','����Ժ','վ��4','�ķ�Ժ','վ��5','���Ժ','վ��6','����Ժ','վ��7','�߷�Ժ','վ��8','�˷�Ժ','վ��9','�ŷ�Ժ',����);

--116691:����һ,2017-11-29,Զ�̿����ʺ����뱣��
Alter Table Zltools.zlClients Add(����Ա�û� Varchar2(20));
Alter Table Zltools.zlClients Add(����Ա���� Varchar2(20));

--113406:����,2017-9-5,ɾ�����ù���վ
Alter Table Zltools.Zlclients Add �����½ʱ�� Date;
Update Zltools.Zlsvrtools Set ���� = '�ͻ������п���' Where ��� = '0308';

--113406:����,2017-9-5,ɾ�����ù���վ
--116691:����һ,2017-11-29,Զ�̿����ʺ����뱣��
Create Or Replace Procedure Zltools.Zl_Zlclients_Set
(
  n_Mode_In       Number,
  n_Rowid_In      Varchar2 := Null,
  v_����վ_In     Zlclients.����վ%Type := Null,
  v_Ip_In         Zlclients.Ip%Type := Null,
  v_Cpu_In        Zlclients.Cpu%Type := Null,
  v_�ڴ�_In       Zlclients.�ڴ�%Type := Null,
  v_Ӳ��_In       Zlclients.Ӳ��%Type := Null,
  v_����ϵͳ_In   Zlclients.����ϵͳ%Type := Null,
  v_����_In       Zlclients.����%Type := Null,
  v_��;_In       Zlclients.��;%Type := Null,
  v_˵��_In       Zlclients.˵��%Type := Null,
  n_����������_In Zlclients.����������%Type := Null,
  n_������־_In   Zlclients.������־%Type := 0,
  n_������_In     Zlclients.������%Type := 0,
  v_վ��_In       Zlclients.վ��%Type := Null,
  n_Apply_In      Number := 0,
  v_Ipbegin_In    Varchar2 := Null,
  v_Ipend_In      Varchar2 := Null,
  n_������ƵԴ    Zlclients.������ƵԴ%Type := Null,
  v_����Ա�û�_In Zlclients.����Ա�û�%Type := Null,
  v_����Ա����_In Zlclients.����Ա����%Type := Null
  --���ܣ������ͻ��˻�վ�� ���߸��¿ͻ�������
  --Ӧ�ã�1�������ߣ��������޸�վ�� ���޸�ʱ��IP��ͻ������ж����������贫��N_Rowid_In��
  --      2��Ӧ��ϵͳ����¼ʱ���ݵ�ǰ��¼�Ŀͻ������ж��Ƿ�
  --                   ����վ����޸�վ�����������ʱN_Rowid_In�贫�룩
  --վ������:0-����վ�㣬1-����վ��
  --N_Apply_In,վ�����Ӧ�÷�Χ��0-��վ�㣬1�������ţ�2������վ�㣬3���̶�IP��
  --V_Ipbegin_In,V_Ipend_In:�ڹ̶�IP��Ӧ��ʱ����,������һ��IP���ϣ���ǰ�沿����ͬ
) Is
  n_Pos         Number(3);
  n_Ipbegin_Num Number;
  n_Ipend_Num   Number;
  n_Ip_Num      Number;
  n_Count       Number;

  v_Err Varchar2(500);
  Err_Custom Exception;

  Function Get_Ipnum(v_Ip_Input Varchar2) Return Number Is
    v_Ip_Num  Varchar2(20);
    n_Pos_Tmp Number;
    v_Ip_Tmp  Varchar2(20);
  Begin
    n_Pos_Tmp := Length(v_Ip_Input);
    n_Pos_Tmp := n_Pos_Tmp - Length(Replace(v_Ip_Input, '.', ''));
    If n_Pos_Tmp <> 3 Then
      Return Null;
    Else
      v_Ip_Tmp := v_Ip_Input;
      Loop
        n_Pos_Tmp := Instr(v_Ip_Tmp, '.');
        Exit When(Nvl(n_Pos_Tmp, 0) = 0);
        --��ÿһ������ת��Ϊ3λ��
        v_Ip_Num := v_Ip_Num || Trim(To_Char(Substr(v_Ip_Tmp, 1, n_Pos_Tmp - 1), '099'));
        v_Ip_Tmp := Substr(v_Ip_Tmp, n_Pos_Tmp + 1);
      End Loop;
      v_Ip_Num := v_Ip_Num || Trim(To_Char(v_Ip_Tmp, '099'));
      n_Ip_Num := To_Number(Trim(v_Ip_Num));
      Return n_Ip_Num;
    End If;
  End;
Begin
  If n_Mode_In = 0 Then
  
    Select Count(1) Into n_Count From zlClients Where ����վ = v_����վ_In;
    If n_Count = 0 Then
      Insert Into zlClients
        (Ip, ����վ, Cpu, �ڴ�, Ӳ��, ����ϵͳ, ����, ��;, ˵��, ����������, ������־, ������, վ��, ������ƵԴ, �����½ʱ��, ����Ա�û�, ����Ա����)
      Values
        (v_Ip_In, v_����վ_In, v_Cpu_In, v_�ڴ�_In, v_Ӳ��_In, v_����ϵͳ_In, v_����_In, v_��;_In, v_˵��_In, n_����������_In, n_������־_In,
         n_������_In, v_վ��_In, n_������ƵԴ, Sysdate, v_����Ա�û�_In, v_����Ա����_In);
    Else
      v_Err := '�Ѿ���������ͬIP��ַ����վ,��������!';
      Raise Err_Custom;
    End If;
  Else
    If n_Rowid_In Is Null Then
      Update zlClients
      Set Cpu = v_Cpu_In, �ڴ� = v_�ڴ�_In, Ӳ�� = v_Ӳ��_In, ����ϵͳ = v_����ϵͳ_In, ���� = v_����_In, ��; = v_��;_In, ˵�� = v_˵��_In,
          ������ = n_������_In, վ�� = v_վ��_In, ������ƵԴ = n_������ƵԴ, ���������� = n_����������_In, ������־ = n_������־_In, �����½ʱ�� = Sysdate,
          ����Ա�û� = Decode(v_����Ա�û�_In, '�տ�', ����Ա�û�, Nvl(v_����Ա�û�_In, ����Ա�û�)),
          ����Ա���� = Decode(v_����Ա����_In, '�տ�', ����Ա����, Nvl(v_����Ա����_In, ����Ա����))
      Where ����վ = v_����վ_In And Ip = v_Ip_In;
    Else
      Update zlClients
      Set ����վ = v_����վ_In, Ip = v_Ip_In, Cpu = Decode(v_Cpu_In, Null, Cpu, v_Cpu_In),
          �ڴ� = Decode(v_�ڴ�_In, Null, �ڴ�, v_�ڴ�_In), Ӳ�� = Decode(v_Ӳ��_In, Null, Ӳ��, v_Ӳ��_In),
          ����ϵͳ = Decode(v_����ϵͳ_In, Null, ����ϵͳ, v_����ϵͳ_In), ���� = v_����_In, վ�� = v_վ��_In, ������ƵԴ = n_������ƵԴ, �����½ʱ�� = Sysdate,
          ����Ա�û� = Decode(v_����Ա�û�_In, '�տ�', ����Ա�û�, Nvl(v_����Ա�û�_In, ����Ա�û�)),
          ����Ա���� = Decode(v_����Ա����_In, '�տ�', ����Ա����, Nvl(v_����Ա����_In, ����Ա����))
      Where Rowid = n_Rowid_In;
    End If;
  End If;
  --������
  If n_Apply_In = 1 Then
    Update zlClients
    Set ������ = n_������_In, վ�� = v_վ��_In
    Where Nvl(����, 'NONE') = Nvl(v_����_In, 'NONE') And Ip <> v_Ip_In;
  Elsif n_Apply_In = 2 Then
    Update zlClients Set ������ = n_������_In, վ�� = v_վ��_In Where Ip <> v_Ip_In;
  Elsif n_Apply_In = 3 Then
    n_Pos := Length(v_Ipbegin_In);
    n_Pos := n_Pos - Length(Replace(v_Ipbegin_In, '.', ''));
    If n_Pos <> 3 Then
      v_Err := '��ʼIP��ʽ����';
      Raise Err_Custom;
    End If;
    n_Pos := Length(v_Ipend_In);
    n_Pos := n_Pos - Length(Replace(v_Ipend_In, '.', ''));
    If n_Pos <> 3 Then
      v_Err := '����IP��ʽ����';
      Raise Err_Custom;
    End If;
  
    n_Ipbegin_Num := Get_Ipnum(v_Ipbegin_In);
    n_Ipend_Num   := Get_Ipnum(v_Ipend_In);
    For r_Ip In (Select ����վ, Ip From zlClients) Loop
      n_Ip_Num := Get_Ipnum(r_Ip.Ip);
      If n_Ip_Num >= n_Ipbegin_Num And n_Ip_Num <= n_Ipend_Num Then
        Update zlClients Set ������ = n_������_In, վ�� = v_վ��_In Where ����վ = r_Ip.����վ And Ip = r_Ip.Ip;
      End If;
    End Loop;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlclients_Set;
/

--113406:����,2017-9-5,ɾ�����ù���վ
Create Or Replace Procedure Zltools.Zl_Zlclients_Deletebatch Is
  d_��½ʱ�� Zlclients.�����½ʱ��%Type;
Begin
  Select Min(�����½ʱ��) Into d_��½ʱ�� From Zlclients;
  Delete Zlclients Where Add_Months(Nvl(�����½ʱ��, d_��½ʱ��), 3) < Sysdate;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zlclients_Deletebatch;
/

--113406:����,2017-9-26,���������δ��¼�ͻ�����������޸�
--116691:����һ,2017-11-29,Զ�̿����ʺ����뱣��
Create Or Replace Package Zltools.b_Runmana Is

  Type t_Refcur Is Ref Cursor;

  Procedure Get_Parameters
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In Number := 0
  );

  Procedure Get_Parameter
  (
    Cursor_Out Out t_Refcur,
    ����id_In  In Zlparameters.Id%Type
  );

  Procedure Get_Parachangedlog
  (
    Cursor_Out Out t_Refcur,
    ����id_In  In Zlparachangedlog.����id%Type
  );

  Procedure Get_Job_Number
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In Number
  );

  Procedure Get_Depict
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In Zldatamove.ϵͳ%Type,
    ���_In    In Zldatamove.���%Type
  );

  Procedure Get_Client_Maxip(Cur_Out Out t_Refcur);

  Procedure Get_Client
  (
    Cur_Out   Out t_Refcur,
    ����վ_In In Zlclients.����վ%Type := Null
  );

  Procedure Get_Client_Station(Cur_Out Out t_Refcur);

  Procedure Get_Project_No(Cur_Out Out t_Refcur);

  Procedure Get_Client_Scheme(Cur_Out Out t_Refcur);

  Procedure Get_Resile
  (
    Cur_Out   Out t_Refcur,
    ������_In In Zlclientparaset.������%Type,
    ����_In   In Number := 0
  );

  Procedure Get_Zldatamove
  (
    Cur_Out Out t_Refcur,
    ϵͳ_In In Zldatamove.ϵͳ%Type
  );

  Procedure Get_Log
  (
    Cur_Out     Out t_Refcur,
    ��־����_In In Varchar2,
    Where_In    In Varchar2
  );

  Procedure Get_Log_Count
  (
    Cur_Out     Out t_Refcur,
    ��־����_In In Varchar2
  );

  Procedure Get_Zlfilesupgrade(Cur_Out Out t_Refcur);

  Procedure Get_Not_Regist(Cur_Out Out t_Refcur);

  Procedure Get_Zloption
  (
    Cur_Out   Out t_Refcur,
    ������_In In Zloptions.������%Type
  );

End b_Runmana;
/
--113406:����,2017-9-26,���������δ��¼�ͻ�����������޸�
Create Or Replace Package Body Zltools.b_Runmana Is

  --���ܣ�ȡ������Ϣ
  --frmParameters
  Procedure Get_Parameters
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In Number := 0
  ) Is
  Begin
    If Nvl(ϵͳ_In, 0) = 0 Then
      Open Cursor_Out For
        Select a.Id, Nvl(a.ϵͳ, 0) ϵͳ, Nvl(a.ģ��, 0) ģ��, Nvl(a.˽��, 0) ˽��, a.������, a.������, a.����ֵ, a.ȱʡֵ, Nvl(a.����, 0) ����,
               a.Ӱ�����˵��, a.����ֵ����, a.����˵��, a.����˵��, a.����˵��, Nvl(a.����, 0) ����, Nvl(a.��Ȩ, 0) ��Ȩ, Nvl(a.�̶�, 0) �̶�,
               Nvl(a.����, 0) ����, b.���� As ģ������, zlSpellCode(b.����) As ģ�����
        From zlParameters A, zlPrograms B
        Where Nvl(a.ϵͳ, 0) = 0 And Nvl(a.ϵͳ, 0) = b.ϵͳ(+) And Nvl(a.ģ��, 0) = b.���(+);
    Else
      Open Cursor_Out For
        Select a.Id, Nvl(a.ϵͳ, 0) ϵͳ, Nvl(a.ģ��, 0) ģ��, Nvl(a.˽��, 0) ˽��, a.������, a.������, a.����ֵ, a.ȱʡֵ, Nvl(a.����, 0) ����,
               a.Ӱ�����˵��, a.����ֵ����, a.����˵��, a.����˵��, a.����˵��, Nvl(a.����, 0) ����, Nvl(a.��Ȩ, 0) ��Ȩ, Nvl(a.�̶�, 0) �̶�,
               Nvl(a.����, 0) ����, b.���� As ģ������, zlSpellCode(b.����) As ģ�����
        From zlParameters A, zlPrograms B,
             --����Ȩ�޲��֣�ֻ����Ȩ�Ĳ�����ʾ
             (Select Distinct f.���
               From zlProgFuncs F, zlRegFunc R
               Where Trunc(f.ϵͳ / 100) = r.ϵͳ(+) And f.��� = r.���(+) And f.���� = r.����(+) And
                     (r.���� Is Not Null Or r.���� Is Null And (f.��� Between 10000 And 19999)) And f.ϵͳ = ϵͳ_In And
                     1 = (Select 1 From Zlregaudit A Where a.��Ŀ = '��Ȩ֤��')
               Union All
               Select 0 As ��� From Dual) M
        Where a.ϵͳ = Nvl(ϵͳ_In, 0) And Nvl(a.ϵͳ, 0) = b.ϵͳ(+) And Nvl(a.ģ��, 0) = b.���(+) And Nvl(a.ģ��, 0) = m.���;
    End If;
  End Get_Parameters;

  --���ܣ�����ָ���Ĳ���IDȡ������Ϣ
  --�����б�frmParameters;frmParaChangeSet
  Procedure Get_Parameter
  (
    Cursor_Out Out t_Refcur,
    ����id_In  In Zlparameters.Id%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select a.Id, Nvl(a.ϵͳ, 0) ϵͳ, Nvl(a.ģ��, 0) ģ��, Nvl(a.˽��, 0) ˽��, a.������, a.������, a.����ֵ, a.ȱʡֵ, Nvl(a.����, 0) ����,
             a.Ӱ�����˵��, a.����ֵ����, a.����˵��, a.����˵��, a.����˵��, Nvl(a.����, 0) ����, Nvl(a.��Ȩ, 0) ��Ȩ, Nvl(a.�̶�, 0) �̶�,
             Nvl(a.����, 0) ����, b.���� As ģ������, zlSpellCode(b.����) As ģ�����
      From zlParameters A, zlPrograms B
      Where a.Id = Nvl(����id_In, 0) And Nvl(a.ϵͳ, 0) = b.ϵͳ(+) And Nvl(a.ģ��, 0) = b.���(+);
  End Get_Parameter;
  --���ܣ�ȡ�����޸���Ϣ
  --�����б�frmParameters
  Procedure Get_Parachangedlog
  (
    Cursor_Out Out t_Refcur,
    ����id_In  In Zlparachangedlog.����id%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select ����id, ���, �䶯˵��, �䶯����, �䶯��, �䶯ʱ��, �䶯ԭ��
      From Zlparachangedlog
      Where ����id = Nvl(����id_In, 0);
  
  End;
  --���ܣ�ȡZlAutoJob���к�
  Procedure Get_Job_Number
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In Number
  ) Is
  Begin
    Open Cursor_Out For
      Select ��� + 1 As ���
      From zlAutoJobs
      Where Nvl(ϵͳ, 0) = ϵͳ_In And ���� = 3 And
            ��� + 1 Not In (Select ��� From zlAutoJobs Where Nvl(ϵͳ, 0) = ϵͳ_In And ���� = 3);
  End Get_Job_Number;

  --���ܣ�ȡZlDataMove����
  Procedure Get_Depict
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In Zldatamove.ϵͳ%Type,
    ���_In    In Zldatamove.���%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select ת������ From zlDataMove Where Nvl(ϵͳ, 0) = ϵͳ_In And ��� = ���_In;
  End Get_Depict;

  --���ܣ�ȡzlClients��MAX IP
  Procedure Get_Client_Maxip(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select Max(Ip) As Ip From zlClients;
  End Get_Client_Maxip;

  --���ܣ�ȡzlClients�ļ�¼
  Procedure Get_Client
  (
    Cur_Out   Out t_Refcur,
    ����վ_In In Zlclients.����վ%Type := Null
  ) Is
    v_Sql Varchar2(1000);
  Begin
    If Nvl(����վ_In, '��') = '��' Then
      v_Sql := 'Select a.Ip, a.����վ, a.Cpu, a.�ڴ�, a.Ӳ��, a.����ϵͳ, a.����, a.��;, a.˵��, a.������־, a.��ֹʹ��,
                             a.������, Decode(b.Terminal, Null, 0, 1) As ״̬, a.�ռ���־,a.����������,a.վ��,c.���� Ժ��,a.������ƵԴ,a.�����½ʱ��,a.����Ա�û�,a.����Ա����
                From Zlclients a, (Select Distinct Terminal From GV$session) b, zlnodelist c
                Where a.����վ = b.Terminal(+) and a.վ�� = c.���(+)
                Order By a.վ��, a.Ip';
      Open Cur_Out For v_Sql;
    Else
      Open Cur_Out For
        Select Ip, ����վ, Cpu, �ڴ�, Ӳ��, ����ϵͳ, ����, ��;, ˵��, ������־, �ռ���־, ��ֹʹ��, ������, ����������, վ��, ������ƵԴ, ����Ա�û�, ����Ա����
        From zlClients
        Where ����վ = ����վ_In;
    End If;
  End Get_Client;

  --���ܣ�ȡzlClients��վ��
  Procedure Get_Client_Station(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select Distinct Upper(����վ) || '[' || Ip || ']' As վ��, Upper(����վ) ����վ From zlClients;
  End Get_Client_Station;

  --���ܣ�ȡ������
  Procedure Get_Project_No(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select ������ From Zlclientparaset Where Rownum = 1;
  End Get_Project_No;

  --���ܣ�ȡ����
  Procedure Get_Client_Scheme(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select ������, ������ || '-' || �������� As ��������, ��������, ����վ, �û��� From Zlclientscheme;
  End Get_Client_Scheme;

  --���ܣ�ȡ�ָ���Ϣ
  Procedure Get_Resile
  (
    Cur_Out   Out t_Refcur,
    ������_In In Zlclientparaset.������%Type,
    ����_In   In Number := 0
  ) Is
  Begin
    If ����_In = 0 Then
      Open Cur_Out For
        Select Distinct a.����վ || Decode(m.����վ, Null, ' ', '[' || m.Ip || ']') As ����վ, a.�û���, a.�ָ���־,
                        '[' || b.������ || ']' || b.�������� As ��������
        From Zlclientparaset A, Zlclientscheme B, zlClients M
        Where a.������ = b.������ And a.����վ = m.����վ(+) And a.������ = ������_In;
    End If;
  
    If ����_In = 1 Then
      Open Cur_Out For
        Select Distinct Upper(����վ) ����վ, Min(�ָ���־) �ָ���־
        From Zlclientparaset A
        Where a.������ = ������_In
        Group By ����վ;
    End If;
  
    If ����_In = 2 Then
      Open Cur_Out For
        Select Distinct Upper(�û���) �û���, Max(����վ) ����վ, Min(Decode(�ָ���־, 2, 0, �ָ���־)) �ָ���־
        From Zlclientparaset A
        Where a.������ = ������_In
        Group By �û���
        Order By �û���;
    End If;
  
  End Get_Resile;

  --���ܣ�ȡzldataMove����
  Procedure Get_Zldatamove
  (
    Cur_Out Out t_Refcur,
    ϵͳ_In In Zldatamove.ϵͳ%Type
  ) Is
  Begin
    Open Cur_Out For
      Select ���, ����, ˵��, �����ֶ�, ת������, �ϴ����� From zlDataMove Where ϵͳ = ϵͳ_In Order By ���;
  End Get_Zldatamove;

  --���ܣ�ȡ��־����
  Procedure Get_Log
  (
    Cur_Out     Out t_Refcur,
    ��־����_In In Varchar2,
    Where_In    In Varchar2
  ) Is
    v_Sql Varchar2(1000);
  Begin
    If ��־����_In = '������־' Then
      v_Sql := 'Select �Ự��,����վ,�û���,�������,������Ϣ,To_char(ʱ��,''yyyy-MM-dd hh24:mi:ss'') ʱ��
                     ,Decode(����,1,''�洢���̴���'',2,''������������'',3,''Ӧ�ó�������'',''�ͻ�����������'') ��������
                        From ZlErrorLog Where ' || Where_In;
      Open Cur_Out For v_Sql;
    End If;
    If ��־����_In = '������־' Then
      v_Sql := 'Select �Ự��,����վ,�û���,������,��������,To_char(����ʱ��,''yyyy-MM-dd hh24:mi:ss'') ����ʱ��
                                 ,To_char(�˳�ʱ��,''yyyy-MM-dd hh24:mi:ss'') �˳�ʱ��,Decode(�˳�ԭ��,1,''�����˳�'',''�쳣�˳�'') �˳�ԭ��
                                    From ZlDiaryLog Where ' || Where_In;
      Open Cur_Out For v_Sql;
    End If;
  End Get_Log;

  --���ܣ�ȡ��־��¼��
  Procedure Get_Log_Count
  (
    Cur_Out     Out t_Refcur,
    ��־����_In In Varchar2
  ) Is
  Begin
    If ��־����_In = '������־' Then
      Open Cur_Out For
        Select Count(*) ����
        From zlErrorLog
        Union All
        Select Nvl(To_Number(����ֵ), 0) From zlOptions Where ������ = 4;
    End If;
    If ��־����_In = '������־' Then
      Open Cur_Out For
        Select Count(*) ����
        From zlDiaryLog
        Union All
        Select Nvl(To_Number(����ֵ), 0) From zlOptions Where ������ = 2;
    
    End If;
  End Get_Log_Count;

  --���ܣ�ȡzlfilesupgradeg����
  Procedure Get_Zlfilesupgrade(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select ���, �ļ���, �汾��, �޸�����, �ļ�˵�� As ˵��,
             Decode(�ļ�����, 0, '��������', 1, 'Ӧ�ò���', 2, '�����ļ�', 3, '�����ļ�', 4, '��������', 5, 'ϵͳ�ļ�', '') As ����, ��װ·�� As ��װ·��,
             Md5 As Md5, ��������
      From zlFilesUpgrade
      Order By ���;
  End Get_Zlfilesupgrade;

  --���ܣ�ȡ��ע����Ŀ
  Procedure Get_Not_Regist(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select ��Ŀ, ����
      From zlRegInfo
      Where ��Ŀ Not In ('������', '�汾��', '������Ŀ¼', '�����û�', '��������', '�ռ�Ŀ¼', '�ռ�����', 'ע����', '��Ȩ֤��', '��Ȩ����', '��Ȩ�ʴ�');
  End Get_Not_Regist;

  --���ܣ�ȡ����ֵ
  Procedure Get_Zloption
  (
    Cur_Out   Out t_Refcur,
    ������_In In Zloptions.������%Type
  ) Is
  Begin
    Open Cur_Out For
      Select Nvl(����ֵ, ȱʡֵ) Option_Value From zlOptions Where ������ = ������_In;
  End Get_Zloption;

End b_Runmana;
/

--109444:����һ,2017-09-28,��Ϣ״̬�ṹ����
drop table zltools.ZLPROCEDURENOTE;
drop INDEX zltools.zlMsgState_IX_��ϢID;
Alter Table Zltools.ZLMSGSTATE Add Constraint ZLMSGSTATE_PK PRIMARY KEY(��ϢID,����,�û�,���) Using Index;


--109444:����һ,2017-09-28,zlFilesExpired��������
Delete From Zltools.zlFilesExpired Where (�ļ���, ��װ·��) In (Select �ļ���, ��װ·�� From Zltools.zlFilesExpired Group By �ļ���, ��װ·�� Having Count(1) > 1) And
Rowid Not In (Select Min(Rowid) From Zltools.zlFilesExpired Group By �ļ���, ��װ·��);

Alter Table zltools.zlFilesExpired Add Constraint zlFilesExpired_PK PRIMARY KEY(�ļ���,��װ·��) Using Index;



Begin
--�ͻ���������־�ṹ������zlClientUpdateLog
  If Zl_Checkobject(1, 'ZLCLIENTUPDATELOG_bak') = 0 Then
    Execute Immediate 'Alter Table zltools.Zlclientupdatelog Rename To Zlclientupdatelog_Bak';
    Execute Immediate 'Create Table zltools.Zlclientupdatelog As Select * From zltools.Zlclientupdatelog_Bak Where 1 = 0';
  End If;
End;
/
Alter Table zltools.Zlclientupdatelog Add (˳��� Number(5));
Alter Table zltools.Zlclientupdatelog Add Constraint Zlclientupdatelog_PK PRIMARY KEY (��������,����վ,˳���) Using Index;


Drop Index zltools.ZLDIARYLOG_IX_�Ự����;
Drop Index zltools.ZLDIARYLOG_IX_����ʱ��;
Begin
--������־�ṹ����:zlDiaryLog
  If Zl_Checkobject(1, 'Zldiarylog_bak') = 0 Then
    Execute Immediate 'Alter table zltools.Zldiarylog rename to Zldiarylog_bak';
    Execute Immediate 'Create table zltools.zldiarylog as select * from zltools.zldiarylog_bak where 1=0';
  End If;
End;
/
ALTER TABLE zltools.Zldiarylog ADD CONSTRAINT Zldiarylog_PK PRIMARY KEY (����ʱ��,�Ự��,������) USING INDEX;

Drop Index zltools.ZLERRORLOG_IX_ʱ��;
Begin
  If Zl_Checkobject(1, 'zlErrorLog_bak') = 0 Then
    Execute Immediate 'Alter table zltools.zlErrorLog rename to zlErrorLog_bak';
    Execute Immediate 'Create table zltools.zlErrorLog as select * from zltools.zlErrorLog_bak where 1=0';
  End If;
End;
/
Alter table zltools.zlErrorLog Add (˳��� Number(5));
Alter Table zltools.zlErrorLog Add Constraint zlErrorLog_PK PRIMARY KEY(ʱ��,�Ự��,�������,˳���) Using Index;

Update zltools.zlOptions Set ����ֵ = decode(sign(����ֵ-1000),1,1000,����ֵ), ȱʡֵ = 1000, ������ = '��־�����������', ����˵�� = '��־����ܱ��������������ʱϵͳ�����Զ�ɾ����' Where ������ = 2;
Update zltools.zlOptions Set ����ֵ = decode(sign(����ֵ-1000),1,1000,����ֵ), ȱʡֵ = 1000, ������ = '���󱣴��������', ����˵�� = '��������ܱ��������������ʱϵͳ�����Զ�ɾ����' Where ������ = 4;

Create Or Replace Procedure Zltools.Zl_Autologprocess As
  --���ܣ�
  --   �Զ����������־�ʹ�����־�������
  v_Limit Number;
Begin
  --ɾ�������������־
  Select Nvl(Max(To_Number(����ֵ)), 0) Into v_Limit From zlOptions Where ������ = 2;
  Delete From zlDiaryLog Where ����ʱ�� < Sysdate - v_Limit;

  --ɾ������Ĵ�����־

  Select Nvl(Max(To_Number(����ֵ)), 0) Into v_Limit From zlOptions Where ������ = 4;
  Delete From zlErrorLog Where ʱ�� < Sysdate - v_Limit;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Autologprocess;
/

Create Or Replace Procedure Zltools.Zl_Zlerrorlog_Insert
(
  ����վ_In    Zlerrorlog.����վ%Type,
  ����_In      Zlerrorlog.����%Type,
  �������_In  Zlerrorlog.�������%Type,
  ������Ϣ_In  Zlerrorlog.������Ϣ%Type,
  Sessionid_In Number := Null
) Is
  n_Audsid Number;
  --���ܣ�
  --   ������־����
Begin
  If Sessionid_In Is Null Then
    Select Userenv('SessionID') Into n_Audsid From Dual;
  Else
    n_Audsid := Sessionid_In;
  End If;
  Insert Into zlErrorLog
    (�Ự��, �û���, ����վ, ʱ��, ����, �������, ������Ϣ, ˳���)
    Select n_Audsid, User, ����վ_In, Sysdate, ����_In, �������_In, ������Ϣ_In, Count(1) + 1
    From zlErrorLog
    Where �Ự�� = n_Audsid And ʱ�� = Sysdate And ������� = �������_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlerrorlog_Insert;
/

--   ������־����
Create Or Replace Procedure Zltools.Zl_Zldiarylog_Insert
(
  ����վ_In    Zldiarylog.����վ%Type,
  ������_In    Zldiarylog.������%Type,
  ������_In    Zldiarylog.������%Type,
  ��������_In  Zldiarylog.��������%Type,
  Sessionid_In Number := Null
) Is
  n_Audsid Number;
  --���ܣ�
  --   ������־����
Begin
  If Sessionid_In Is Null Then
    Select Userenv('SessionID') Into n_Audsid From dual;
  Else
    n_Audsid := Sessionid_In;
  End If;
  Insert Into zlDiaryLog
    (�Ự��, �û���, ����վ, ������, ������, ��������, ����ʱ��)
    Select n_Audsid, User, ����վ_In, ������_In, ������_In, ��������_In, Sysdate From Dual;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zldiarylog_Insert;
/

--   ������־����
Create Or Replace Procedure Zltools.Zl_Zldiarylog_Update
(
  ����վ_In    Zldiarylog.����վ%Type,
  ������_In    Zldiarylog.������%Type,
  ������_In    Zldiarylog.������%Type,
  �˳�ԭ��_In  Zldiarylog.�˳�ԭ��%Type,
  Sessionid_In Number := Null
) Is
  n_�Ự�� Zldiarylog.�Ự��%Type;
  --���ܣ�
  --   ������־����
Begin
  If Sessionid_In Is Null Then
    Select Userenv('SessionID') Into n_�Ự�� From Dual;
  Else
    n_�Ự�� := Sessionid_In;
  End If;
  Update zlDiaryLog
  Set �˳�ԭ�� = �˳�ԭ��_In, �˳�ʱ�� = Sysdate
  Where �˳�ԭ�� Is Null And �û��� = User And ����վ = ����վ_In And �Ự�� = n_�Ự�� And ������ = ������_In And ������ = ������_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zldiarylog_Update;
/

--   �ͻ���������־����
Create Or Replace Procedure Zltools.Zl_Zlclientupdatelog_Insert
(
  ����_In   Zlclientupdatelog.����%Type,
  ����վ_In Zlclientupdatelog.����վ%Type
) Is
  v_����վ Zlclientupdatelog.����վ%Type;
  --���ܣ�
  --   �ͻ���������־����
Begin

  If ����վ_In Is Null Then
    Select Userenv('Terminal') Into v_����վ From Dual;
  Else
    v_����վ := ����վ_In;
  End If;
  Insert Into Zlclientupdatelog
    (����վ, ��������, ����, ˳���)
    Select v_����վ, Sysdate, ����_In, Count(1) + 1
    From Zlclientupdatelog
    Where ����վ = v_����վ And �������� = Sysdate;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlclientupdatelog_Insert;
/

  --   �Զ���������
 Create Or Replace Procedure Zltools.Zl_Zlfilesexpired_Insert
 (
   �ļ���_In   Zlfilesexpired.�ļ���%Type,
   ��װ·��_In Zlfilesexpired.��װ·��%Type,
   ϵͳ���_In Zlfilesexpired.ϵͳ���%Type,
   ϵͳ�汾_In Zlfilesexpired.ϵͳ�汾%Type,
   ˵��_In     Zlfilesexpired.˵��%Type
 ) Is
  --���ܣ�
  --   �Զ���������
 Begin
   Insert Into Zlfilesexpired
     (�ļ���, ��װ·��, ϵͳ���, ϵͳ�汾, ˵��)
     Select �ļ���_In, ��װ·��_In, ϵͳ���_In, ϵͳ�汾_In, ˵��_In From Dual;
 Exception
   When Others Then
     zl_ErrorCenter(SQLCode, SQLErrM);
 End Zl_Zlfilesexpired_Insert;
/

--����һ,������ͷ
Create Or Replace Package zltools.b_Comfunc Is
--��Ҫ���ڹ��������Ĺ���
  Type t_Refcur Is Ref Cursor;
--���ܣ����������־
--�����б�clsComLib.SaveErrLog
  Procedure Save_Error_Log
  (
    ����_In     In zlErrorLog.����%Type,
    �������_In In zlErrorLog.�������%Type,
    ������Ϣ_In In zlErrorLog.������Ϣ%Type
  );
--���ܣ�ȡ���ù���
--�����б�clsComLib.ShowAbout
  Procedure Get_Usable_Function
  (
    Cursor_Out Out t_Refcur,
    ����_In    In zlPrograms.����%Type
  );
--���ܣ�ȡ��д���
--�����б�clsCommFun.UppeMoney
  Procedure Get_Uppmoney
  (
    Cursor_Out Out t_Refcur,
    ���_In    In Number
  );
--���ܣ�����ָ�������ڡ���š�ϵͳ�ж�ָ�����ڵ������Ƿ���ת���������ݱ���
--�����б�clsDatabase.DateMoved
  Procedure Get_Datamoved
  (
    Cursor_Out  Out t_Refcur,
    ���_In     In zlDataMove.���%Type,
    ϵͳ_In     In zlDataMove.ϵͳ%Type,
    �ϴ�����_In In zlDataMove.�ϴ�����%Type
  );
--���ܣ�ȡϵͳ������
--�����б�clsDatabase.GetOwner
  Procedure Get_Owner
  (
    Cursor_Out Out t_Refcur,
    ���_In    In zlSystems.���%Type
  );
--���ܣ�ȡ����
--�����б�clsCommFun.SpellCode
  Procedure Get_Spell_Code
  (
    Cursor_Out Out t_Refcur,
    �ַ���_In  In Varchar2,
    ��ʽ_In    In Number := 0
  );
--���ܣ�����������־
--�����б�clsComLib.RestoreWinState
  Procedure Save_Diary_Log
  (
    ������_In   In zlDiaryLog.������%Type,
    ������_In   In zlDiaryLog.������%Type,
    ��������_In In zlDiaryLog.��������%Type
  );
--���ܣ�����������־
--�����б�clsComLib.SaveWinState
  Procedure Update_Diary_Log
  (
    ������_In In zlDiaryLog.������%Type,
    ������_In In zlDiaryLog.������%Type
  );
--���ܣ�ȡ�̶�����������û���������
--�����б�clsDatabase.ShowReportMenu
  Procedure Get_Report_Menu
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In zlPrograms.ϵͳ%Type,
    ���_In    In zlPrograms.���%Type,
    ����_In    In zlReports.����%Type,
    ���_In    In zlReports.���%Type
  );
--���ܣ�ȡ�û�������Ϣ
--�����б�zlApptools.frmAlert
  Procedure Get_Zlnoticerec
  (
    Cursor_Out Out t_Refcur,
    �û���_In  In zlNoticeRec.�û���%Type
  );
--���ܣ�ȡ�ʼ�����
--�����б�zlApptools.frmMessageEdit.LoadMessage
  Procedure Get_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In zlMessages.ID%Type,
    ����_In    In zlMsgState.����%Type,
    �û�_In    In zlMsgState.�û�%Type
  );
--���ܣ�ȡ�ʼ�����
--�����б�zlApptools.frmMessageManager.FillText��zlApptools.frmMessageRelate.FillText
  Procedure Get_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In zlMessages.ID%Type
  );
--���ܣ�ȡ�ʵݵ�ַ
--�����б�zlApptools.frmMessageEdit.LoadMessage
  Procedure Get_Zlmsgstate
  (
    Cursor_Out Out t_Refcur,
    ��Ϣid_In  In zlMsgState.��Ϣid%Type
  );
--���ܣ�ɾ����Ϣ
--�����б�zlApptools.frmMessageManager.mnuEditDelete_Click
  Procedure Delete_Zlmsgstate
  (
    ɾ��_In   In zlMsgState.ɾ��%Type,
    ��Ϣid_In In zlMsgState.��Ϣid%Type,
    ����_In   In zlMsgState.����%Type,
    �û�_In   In zlMsgState.�û�%Type
  );
--���ܣ�ɾ��������Ϣ
--�����б�zlApptools.frmMessageManager.DeleteMessage
  Procedure Delete_Zlmessage;
--���ܣ�ȡ�ʼ��б�
--�����б�zlApptools.frmMessageManager.FillList��zlApptools.frmMessageRelate.FillList
  Procedure Get_Mail_List
  (
    Cursor_Out  Out t_Refcur,
    ��Ϣ����_In In Varchar2,
    �û�_In     In zlMsgState.�û�%Type,
    ��ʾ�Ѷ�_In In Number,
    �Ựid_In   In zlMessages.�Ựid%Type
  );
--���ܣ���ԭɾ������Ϣ
--�����б�zlApptools.frmMessageManager.mnuEditRestore_Click
  Procedure Restore_Zlmsgstate
  (
    ��Ϣid_In In zlMsgState.��Ϣid%Type,
    ����_In   In zlMsgState.����%Type,
    �û�_In   In zlMsgState.�û�%Type
  );
--���ܣ�����������Ϣ
--�����б�zlApptools.frmMessageEdit.SaveMessage
  Procedure Save_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In zlMessages.ID%Type,
    �Ựid_In  In zlMessages.�Ựid%Type,
    ������_In  In zlMessages.������%Type,
    �ռ���_In  In zlMessages.�ռ���%Type,
    ����_In    In zlMessages.����%Type,
    ����_In    In zlMessages.����%Type,
    ����ɫ_In  In zlMessages.����ɫ%Type
  );
--���ܣ�����zlMsgstate
--�����б�zlApptools.frmMessageEdit.SaveMessage
  Procedure Insert_Zlmsgstate
  (
    ��Ϣid_In In zlMsgState.��Ϣid%Type,
    ����_In   In zlMsgState.����%Type,
    �û�_In   In zlMsgState.�û�%Type,
    ���_In   In zlMsgState.���%Type,
    ɾ��_In   In zlMsgState.ɾ��%Type,
    ״̬_In   In zlMsgState.״̬%Type
  );
--���ܣ�Ϊԭ�����ϴ𸴻�ת����־
--�����б�zlApptools.frmMessageEdit.SaveMessage
  Procedure Update_Zlmsgstate_State
  (
    ģʽ_In   In Number,
    ��Ϣid_In In zlMsgState.��Ϣid%Type,
    ����_In   In zlMsgState.����%Type,
    �û�_In   In zlMsgState.�û�%Type
  );
--���ܣ�Ϊԭ�����ϴ𸴻�ת����־
--�����б�zlApptools.frmMessageEdit.LoadMessage
  Procedure Update_Zlmsgstate_Idtntify
  (
    ���_In   In zlMsgState.���%Type,
    ��Ϣid_In In zlMsgState.��Ϣid%Type,
    ����_In   In zlMsgState.����%Type,
    �û�_In   In zlMsgState.�û�%Type
  );

End b_Comfunc;
/

Create Or Replace Package Body Zltools.b_Comfunc Is
  --���ܣ����������־
  Procedure Save_Error_Log
  (
    ����_In     In Zlerrorlog.����%Type,
    �������_In In Zlerrorlog.�������%Type,
    ������Ϣ_In In Zlerrorlog.������Ϣ%Type
  ) Is
    n_�Ự�� Number;
    v_����վ Zlerrorlog.����վ%Type;
  Begin
    Select Userenv('SessionID'), Userenv('Terminal') Into n_�Ự��, v_����վ From Dual;
    Insert Into zlErrorLog
      (�Ự��, �û���, ����վ, ʱ��, ����, �������, ������Ϣ, ˳���)
      Select n_�Ự��, User, v_����վ, Sysdate, ����_In, �������_In, ������Ϣ_In, Count(1) + 1
      From zlErrorLog
      Where �Ự�� = n_�Ự�� And ʱ�� = Sysdate And ������� = �������_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Save_Error_Log;

  --���ܣ�ȡ���ù���
  Procedure Get_Usable_Function
  (
    Cursor_Out Out t_Refcur,
    ����_In    In Zlprograms.����%Type
  ) Is
  Begin
    If Nvl(����_In, '�տ�') = '�տ�' Then
      Open Cursor_Out For
        Select Distinct a.���, a.����, a.˵��
        From zlPrograms A, zlProgFuncs B, zlRegFunc C
        Where a.ϵͳ = b.ϵͳ And a.��� = b.��� And Trunc(b.ϵͳ / 100) = c.ϵͳ(+) And b.��� = c.���(+) And b.���� = c.����(+) And
              (c.���� Is Not Null Or c.���� Is Null And (a.��� Between 10000 And 19999))
        Order By a.���;
    Else
      Open Cursor_Out For
        Select Distinct a.���, a.����, a.˵��
        From zlPrograms A, zlProgFuncs B, zlRegFunc C
        Where a.ϵͳ = b.ϵͳ And a.��� = b.��� And Upper(a.����) = Upper(����_In) And Trunc(b.ϵͳ / 100) = c.ϵͳ(+) And
              b.��� = c.���(+) And b.���� = c.����(+) And
              (c.���� Is Not Null Or c.���� Is Null And (a.��� Between 10000 And 19999))
        Order By a.���;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Usable_Function;

  --���ܣ�ȡ��д���
  Procedure Get_Uppmoney
  (
    Cursor_Out Out t_Refcur,
    ���_In    In Number
  ) Is
  Begin
    Open Cursor_Out For
      Select zlUppMoney(Nvl(���_In, 0)) As Num From Dual;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Uppmoney;

  --���ܣ�����ָ�������ڡ���š�ϵͳ�ж�ָ�����ڵ������Ƿ���ת���������ݱ���
  Procedure Get_Datamoved
  (
    Cursor_Out  Out t_Refcur,
    ���_In     In Zldatamove.���%Type,
    ϵͳ_In     In Zldatamove.ϵͳ%Type,
    �ϴ�����_In In Zldatamove.�ϴ�����%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select ϵͳ, ���
      From zlDataMove
      Where ��� = ���_In And ϵͳ = ϵͳ_In And �ϴ����� > �ϴ�����_In And �ϴ����� Is Not Null;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Datamoved;

  --���ܣ�ȡϵͳ������
  Procedure Get_Owner
  (
    Cursor_Out Out t_Refcur,
    ���_In    In Zlsystems.���%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select ������ From zlSystems Where ��� = ���_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Owner;

  --���ܣ�ȡ����
  Procedure Get_Spell_Code
  (
    Cursor_Out Out t_Refcur,
    �ַ���_In  In Varchar2,
    ��ʽ_In    In Number := 0
  ) Is
  Begin
    If Nvl(��ʽ_In, 0) = 0 Then
      Open Cursor_Out For
        Select zlSpellCode(�ַ���_In) As ���� From Dual;
    Else
      Open Cursor_Out For
        Select zlWbCode(�ַ���_In) As ���� From Dual;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Spell_Code;

  --���ܣ�����������־
  Procedure Save_Diary_Log
  (
    ������_In   In Zldiarylog.������%Type,
    ������_In   In Zldiarylog.������%Type,
    ��������_In In Zldiarylog.��������%Type
  ) Is
  Begin
    Insert Into zlDiaryLog
      (�Ự��, �û���, ����վ, ������, ������, ��������, ����ʱ��)
      Select Userenv('SessionID'), User, RTrim(LTrim(Replace(Userenv('Terminal'), Chr(0), ''))), ������_In, ������_In, ��������_In, Sysdate
      From dual;
    Commit;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Save_Diary_Log;

  --���ܣ�����������־
  --�����б�clsComLib.SaveWinState
  Procedure Update_Diary_Log
  (
    ������_In In Zldiarylog.������%Type,
    ������_In In Zldiarylog.������%Type
  ) Is
    Cursor c_Session Is
      Select Userenv('SessionID') As �Ự��, User As �û���, RTrim(LTrim(Replace(Userenv('Terminal'), Chr(0), ''))) As ����վ
      From dual;
  Begin
    For r_Tmp In c_Session Loop
      Update zlDiaryLog
      Set �˳�ԭ�� = 1, �˳�ʱ�� = Sysdate
      Where �˳�ԭ�� Is Null And �û��� = r_Tmp.�û��� And ����վ = r_Tmp.����վ And �Ự�� = r_Tmp.�Ự�� And ������ = ������_In And ������ = ������_In;
    End Loop;
    Commit;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Update_Diary_Log;

  --���ܣ�ȡ�̶�����������û���������
  Procedure Get_Report_Menu
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In Zlprograms.ϵͳ%Type,
    ���_In    In Zlprograms.���%Type,
    ����_In    In Zlreports.����%Type,
    ���_In    In Zlreports.���%Type
  ) Is
  Begin
    If Nvl(���_In, '�տ�') <> '�տ�' Then
      Open Cursor_Out For
        Select ��־, ϵͳ, ���, ����
        From (Select 1 As ��־, a.ϵͳ, a.���, a.����
               From zlReports A, zlPrograms B
               Where a.ϵͳ = b.ϵͳ And a.����id = b.��� And Not Upper(a.���) Like '%BILL%' And
                     Upper(b.����) <> Upper('zl9Report') And b.ϵͳ = ϵͳ_In And b.��� = ���_In And
                     Instr(����_In, ';' || a.���� || ';') > 0
               Union All
               Select Decode(a.ϵͳ, Null, 2, 1) As ��־, a.ϵͳ, a.���, a.����
               From zlReports A, zlRPTPuts B, zlPrograms C
               Where a.Id = b.����id And b.ϵͳ = c.ϵͳ And b.����id = c.��� And (Not Upper(a.���) Like '%BILL%' Or a.ϵͳ Is Null) And
                     Instr(����_In, ';' || b.���� || ';') > 0 And c.ϵͳ = ϵͳ_In And c.��� = ���_In)
        Where Instr(���_In, ',' || ��� || ',') = 0
        Order By ��־, ���;
    Else
      Open Cursor_Out For
        Select ��־, ϵͳ, ���, ����
        From (Select 1 As ��־, a.ϵͳ, a.���, a.����
               From zlReports A, zlPrograms B
               Where a.ϵͳ = b.ϵͳ And a.����id = b.��� And Not Upper(a.���) Like '%BILL%' And
                     Upper(b.����) <> Upper('zl9Report') And b.ϵͳ = ϵͳ_In And b.��� = ���_In And
                     Instr(����_In, ';' || a.���� || ';') > 0
               Union All
               Select Decode(a.ϵͳ, Null, 2, 1) As ��־, a.ϵͳ, a.���, a.����
               From zlReports A, zlRPTPuts B, zlPrograms C
               Where a.Id = b.����id And b.ϵͳ = c.ϵͳ And b.����id = c.��� And (Not Upper(a.���) Like '%BILL%' Or a.ϵͳ Is Null) And
                     Instr(����_In, ';' || b.���� || ';') > 0 And c.ϵͳ = ϵͳ_In And c.��� = ���_In)
        Order By ��־, ���;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Report_Menu;

  --���ܣ�ȡ�û�������Ϣ
  Procedure Get_Zlnoticerec
  (
    Cursor_Out Out t_Refcur,
    �û���_In  In Zlnoticerec.�û���%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select a.���, a.ϵͳ, c.����id As ģ��, c.ϵͳ As ����ϵͳ, b.�������� As �������, c.���� As ���ѱ���, a.��������, b.���ʱ��, b.�Ѷ���־
      From zlNotices A, zlNoticeRec B, (Select * From zlReports Where ����ʱ�� Is Not Null) C
      Where b.�û��� = �û���_In And b.���ѱ�־ > 0 And c.���(+) = a.���ѱ��� And a.��� = b.������� And b.�������� Is Not Null;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Zlnoticerec;

  --���ܣ�ȡ�ʼ�����
  Procedure Get_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In Zlmessages.Id%Type,
    ����_In    In Zlmsgstate.����%Type,
    �û�_In    In Zlmsgstate.�û�%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select a.*, b.ɾ��, b.״̬
      From zlMessages A, zlMsgState B
      Where a.Id = b.��Ϣid And b.��Ϣid = Id_In And b.���� = ����_In And b.�û� = �û�_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Zlmessage;

  --���ܣ�ȡ�ʼ�����
  Procedure Get_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In Zlmessages.Id%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select ����, ����ɫ From zlMessages Where ID = Id_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Zlmessage;

  --���ܣ�ȡ�ʵݵ�ַ
  Procedure Get_Zlmsgstate
  (
    Cursor_Out Out t_Refcur,
    ��Ϣid_In  In Zlmsgstate.��Ϣid%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select ����, �û�, ��� From zlMsgState Where ��Ϣid = ��Ϣid_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Zlmsgstate;

  --���ܣ�ɾ����Ϣ
  Procedure Delete_Zlmsgstate
  (
    ɾ��_In   In Zlmsgstate.ɾ��%Type,
    ��Ϣid_In In Zlmsgstate.��Ϣid%Type,
    ����_In   In Zlmsgstate.����%Type,
    �û�_In   In Zlmsgstate.�û�%Type
  ) Is
    n_���� Number(10);
    n_���� Number(10);
  Begin
    If Nvl(ɾ��_In, 0) = 1 Then
      Update zlMsgState Set ɾ�� = 1 Where ��Ϣid = ��Ϣid_In And ���� = ����_In And �û� = �û�_In;
    Else
      If ����_In = 0 Then
        --���ڲݸ壬���ռ��˵�Ҳһ��ɾ��
        Update zlMsgState Set ɾ�� = 2 Where ��Ϣid = ��Ϣid_In And �û� = �û�_In;
      Else
        Update zlMsgState Set ɾ�� = 2 Where ��Ϣid = ��Ϣid_In And ���� = ����_In And �û� = �û�_In;
      End If;
      -- ɾ��ָ��ID����Ϣ  mnuEditDelete_Click ����
      Select Count(*) As ����, Sum(Decode(ɾ��, 2, 1, 0)) As ����
      Into n_����, n_����
      From zlMsgState
      Where ��Ϣid = ��Ϣid_In;

      If n_���� = n_���� Then
        Delete From zlMessages Where ID = ��Ϣid_In;
      End If;
    End If;
    Commit;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Delete_Zlmsgstate;

  --���ܣ�ɾ��������Ϣ
  Procedure Delete_Zlmessage Is
    n_Days Number;
  Begin
    Select Nvl(����ֵ, ȱʡֵ) Into n_Days From zlOptions Where ������ = 5;
    If Nvl(n_Days, 0) > 0 Then
      Delete From zlMessages Where ʱ�� < Sysdate - n_Days;
      Commit;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Delete_Zlmessage;

  --���ܣ�ȡ�ʼ��б�
  Procedure Get_Mail_List
  (
    Cursor_Out  Out t_Refcur,
    ��Ϣ����_In In Varchar2,
    �û�_In     In Zlmsgstate.�û�%Type,
    ��ʾ�Ѷ�_In In Number,
    �Ựid_In   In Zlmessages.�Ựid%Type
  ) Is
    v_Sql  Varchar2(1000);
    v_�Ѷ� Varchar2(100);
    v_���� Varchar2(100);
  Begin

    If Nvl(��ʾ�Ѷ�_In, 0) = 1 Then
      v_�Ѷ� := ' and substr(S.״̬,1,1)=''0''';
    Else
      v_�Ѷ� := '';
    End If;

    If Instr(';�ݸ�;�ռ���;�ѷ�����Ϣ;��ɾ����Ϣ;�����Ϣ;', ';' || ��Ϣ����_In || ';') <= 0 Then
      v_���� := '�ݸ�';
    Else
      v_���� := ��Ϣ����_In;
    End If;

    If v_���� = '�ݸ�' Then
      v_Sql := 'Select M.ID, M.�Ựid, M.������, M.�ռ���, M.����, To_Char(M.ʱ��, ''YYYY-MM-DD HH24:MI:SS'') As ʱ��, S.����, S.״̬
              From zlMessages M, zlMsgState S
              Where M.ID = S.��Ϣid  and S.ɾ��=0 and S.�û�= ''' || �û�_In || ''' And S.����=0 ' || v_�Ѷ�;
    End If;

    If v_���� = '�ռ���' Then
      v_Sql := 'Select M.ID, M.�Ựid, M.������, M.�ռ���, M.����, To_Char(M.ʱ��, ''YYYY-MM-DD HH24:MI:SS'') As ʱ��, S.����, S.״̬
              From zlMessages M, zlMsgState S
              Where M.ID = S.��Ϣid  and S.ɾ��=0 and S.�û�= ''' || �û�_In || ''' And S.����=2 ' || v_�Ѷ�;
    End If;

    If v_���� = '�ѷ�����Ϣ' Then
      v_Sql := 'Select M.ID, M.�Ựid, M.������, M.�ռ���, M.����, To_Char(M.ʱ��, ''YYYY-MM-DD HH24:MI:SS'') As ʱ��, S.����, S.״̬
              From zlMessages M, zlMsgState S
              Where M.ID = S.��Ϣid  and S.ɾ��=0 and S.�û�= ''' || �û�_In || ''' And S.����=1 ' || v_�Ѷ�;
    End If;

    If v_���� = '��ɾ����Ϣ' Then
      v_Sql := 'Select M.ID, M.�Ựid, M.������, M.�ռ���, M.����, To_Char(M.ʱ��, ''YYYY-MM-DD HH24:MI:SS'') As ʱ��, S.����, S.״̬
              From zlMessages M, zlMsgState S
              Where M.ID = S.��Ϣid  and S.�û�= ''' || �û�_In || ''' And S.ɾ��=1 ' || v_�Ѷ�;
    End If;

    If v_���� = '�����Ϣ' Then
      v_Sql := 'select M.ID,M.�ỰID,M.������,M.�ռ���,M.����,to_char(M.ʱ��,''YYYY-MM-DD HH24:MI:SS'') as ʱ��,S.����,S.״̬
         from zlMessages M,zlMsgState S where M.ID=S.��ϢID and S.ɾ��<>2 and S.�û�= ''' || �û�_In ||
               '''  and M.�ỰID=' || �Ựid_In;
    End If;

    If Nvl(v_Sql, '�տ�') <> '�տ�' Then
      Open Cursor_Out For v_Sql;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Mail_List;

  --���ܣ���ԭɾ������Ϣ
  Procedure Restore_Zlmsgstate
  (
    ��Ϣid_In In Zlmsgstate.��Ϣid%Type,
    ����_In   In Zlmsgstate.����%Type,
    �û�_In   In Zlmsgstate.�û�%Type
  ) Is
  Begin
    Update zlMsgState Set ɾ�� = 0 Where ��Ϣid = ��Ϣid_In And ���� = ����_In And �û� = �û�_In;
    Commit;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Restore_Zlmsgstate;

  --���ܣ�������Ϣ
  --�����б�zlApptools.frmMessageEdit.SaveMessage
  Procedure Save_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In Zlmessages.Id%Type,
    �Ựid_In  In Zlmessages.�Ựid%Type,
    ������_In  In Zlmessages.������%Type,
    �ռ���_In  In Zlmessages.�ռ���%Type,
    ����_In    In Zlmessages.����%Type,
    ����_In    In Zlmessages.����%Type,
    ����ɫ_In  In Zlmessages.����ɫ%Type
  ) Is
    n_Id     Zlmessages.Id%Type;
    n_�Ựid Zlmessages.�Ựid%Type;
  Begin
    If Nvl(Id_In, 0) = 0 Then
      Select Zlmessages_Id.Nextval Into n_Id From Dual;
      n_Id := Nvl(n_Id, 0);
      If Nvl(�Ựid_In, 0) = 0 Then
        n_�Ựid := n_Id;
      Else
        n_�Ựid := �Ựid_In;
      End If;
      Insert Into zlMessages
        (ID, �Ựid, ������, ʱ��, �ռ���, ����, ����, ����ɫ)
      Values
        (n_Id, n_�Ựid, ������_In, Sysdate, �ռ���_In, ����_In, ����_In, ����ɫ_In);
      Open Cursor_Out For
        Select n_Id As ID, n_�Ựid As �Ựid From Dual;
    Else
      Update zlMessages
      Set ������ = ������_In, ʱ�� = Sysdate, �ռ��� = �ռ���_In, ���� = ����_In, ���� = ����_In, ����ɫ = ����ɫ_In
      Where ID = Id_In;
      Open Cursor_Out For
        Select Id_In As ID, �Ựid_In As �Ựid From Dual;
    End If;
    Commit;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Save_Zlmessage;

  --���ܣ�����zlMsgstate
  --�����б�zlApptools.frmMessageEdit.SaveMessage
  Procedure Insert_Zlmsgstate
  (
    ��Ϣid_In In Zlmsgstate.��Ϣid%Type,
    ����_In   In Zlmsgstate.����%Type,
    �û�_In   In Zlmsgstate.�û�%Type,
    ���_In   In Zlmsgstate.���%Type,
    ɾ��_In   In Zlmsgstate.ɾ��%Type,
    ״̬_In   In Zlmsgstate.״̬%Type
  ) Is
  Begin

    If ����_In < 2 Then
      Delete From zlMsgState Where ��Ϣid = ��Ϣid_In;
    End If;
    Insert Into zlMsgState
      (��Ϣid, ����, �û�, ���, ɾ��, ״̬)
    Values
      (��Ϣid_In, ����_In, �û�_In, ���_In, ɾ��_In, ״̬_In);
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Insert_Zlmsgstate;

  --���ܣ�Ϊԭ�����ϴ𸴻�ת����־
  Procedure Update_Zlmsgstate_State
  (
    ģʽ_In   In Number,
    ��Ϣid_In In Zlmsgstate.��Ϣid%Type,
    ����_In   In Zlmsgstate.����%Type,
    �û�_In   In Zlmsgstate.�û�%Type
  ) Is
  Begin
    If Nvl(ģʽ_In, 0) = 1 Or Nvl(ģʽ_In, 0) = 2 Then
      Update zlMsgState
      Set ״̬ = Substr(״̬, 1, 1) || '1' || Substr(״̬, 3, 2)
      Where ��Ϣid = ��Ϣid_In And ���� = ����_In And �û� = �û�_In;
      Commit;
    End If;
    If Nvl(ģʽ_In, 0) = 3 Then
      Update zlMsgState
      Set ״̬ = Substr(״̬, 1, 1) || '1' || Substr(״̬, 4, 1)
      Where ��Ϣid = ��Ϣid_In And ���� = ����_In And �û� = �û�_In;
      Commit;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Update_Zlmsgstate_State;

  --���ܣ�����״̬�����
  Procedure Update_Zlmsgstate_Idtntify
  (
    ���_In   In Zlmsgstate.���%Type,
    ��Ϣid_In In Zlmsgstate.��Ϣid%Type,
    ����_In   In Zlmsgstate.����%Type,
    �û�_In   In Zlmsgstate.�û�%Type
  ) Is
  Begin
    Update zlMsgState
    Set ״̬ = '1' || Substr(״̬, 2), ��� = ���_In
    Where ��Ϣid = ��Ϣid_In And ���� = ����_In And �û� = �û�_In;
    Commit;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Update_Zlmsgstate_Idtntify;

End b_Comfunc;
/
--113538:������,2017-10-30,�����Ա����Ӷ����ֶ�
Alter Table Zltools.Zlrptcolproterty Add ���� Number(1);

--115010:����,2017-11-7,��ϵͳ��¼���ⴴ����ɫ��
Create Table zltools.Zlroles(
       ����   Varchar2(50),
       ϵͳ   Number(5));
ALTER TABLE zltools.Zlroles ADD CONSTRAINT ZLRoles_PK PRIMARY KEY (����) USING INDEX;
Insert Into zlTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLROLES','ZLTOOLSTBS','A2');
--115010:����,2017-11-7,��ϵͳ��¼�������ӽ�ɫ��Ȩ��Լ��
Alter Table zltools.zlRoleGroups Modify ���� Constraint zlRoleGroups_NN_���� Not Null;
--115010:����,2017-11-7,��ϵͳ��¼�����ʼ����ɫ����
Truncate Table Zltools.Zlroles;
Insert Into zltools.Zlroles(����) Select Role From Sys.Dba_Roles Where Upper(Role) Like 'ZL_%';
Delete zltools.Zlrolegroups Where ���� = 'δ����';
--115010:����,2017-11-7,��ϵͳ��¼����
CREATE OR REPLACE Procedure zltools.Zl_Zlroles_Edit
(
  ����_In In Number, --1-���ӣ�2-�޸ģ�3��ɾ��
  ����_In In Zlroles.����%Type,
  ϵͳ_In In Zlroles.ϵͳ%Type := Null
) Is
Begin
  If ����_In = 1 Then
    Insert Into Zlroles Values (����_In, ϵͳ_In);
  Elsif ����_In = 2 Then
    Update Zlroles Set ϵͳ = ϵͳ_In Where ���� = ����_In;
  Else
    Delete Zlroles Where ���� = ����_In;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zlroles_Edit;
/
--115010:����,2017-11-7,��ϵͳ��¼����
Create Or Replace Procedure Zltools.Zl_Checkrolesdiff As
  --���ܣ����Dba_Roles���е����ݺ�Zlroles�е��Ƿ�һ�£�����һ�£���Dba_Roles��
  --     ����������ݲ��뵽zlroles��
Begin
  --��Dba_Roles�д��������ɫ����Zlroles���治���ڣ��򽫸ý�ɫ��ӵ�Zlroles��ȥ
  Insert Into Zlroles
    (����)
    Select a.Role
    From Dba_Roles a
    Where Not Exists (Select 1 From Zlroles b Where a.Role = b.����) And a.Role Like 'ZL_%';
  --��Dba_Roles�в����ڸý�ɫ����Zlroles�д��ڣ��򽫸ý�ɫ��Zlroles��ɾ��
  Delete Zlroles a Where Not Exists (Select 1 From Dba_Roles b Where a.���� = b.Role);
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Checkrolesdiff;
/

--116688:����һ,2017-11-08,�����ߵ�¼����
Insert Into zlTools.zlSvrTools(���,�ϼ�,����,���,˵��,����) Values('0607','06','�û���IP����','I',Null,7);
Insert Into zlTools.zlSvrTools(���,�ϼ�,����,���,˵��,����) Values('0608','06','Ӧ�ó�����Ȩ','A',Null,8);
Insert Into zlTools.zlSvrTools(���,�ϼ�,����,���,˵��,����) Values('0609','06','�û���¼��־','L',Null,9);


--116688:����һ,2017-11-08,�û���¼����
Create Table zlTools.zlLoginLimit(
    ID  Number(18),
    �û��� Varchar2(50),
    ��ʼIP Varchar2(20),
    ����IP Varchar2(20),
    ��ʼʱ�� Date,
    ����ʱ�� Date,
    ״̬  Number(1),
    ˵��  Varchar2(200));
CREATE Sequence zlTools.zlLoginLimit_ID start with 1;
ALTER TABLE zlTools.zlLoginLimit ADD CONSTRAINT zlLoginLimit_PK PRIMARY KEY (ID) USING INDEX;
Alter Table zlTools.zlLoginLimit Add Constraint zlLoginLimit_Uq_�û��� Unique(�û���,��ʼIP,����IP,��ʼʱ��,����ʱ��) Using Index;
Insert into zlTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLLOGINLIMIT','ZLTOOLSTBS','C1');

--116688:����һ,2017-11-08,���ƹ�������/�޸�
Create Or Replace Procedure Zltools.Zl_Zlloginlimit_Edit
(
  ��������_In In Number,
  Id_In       In Zlloginlimit.Id%Type,
  ��ʼʱ��_In In Zlloginlimit.��ʼʱ��%Type,
  ����ʱ��_In In Zlloginlimit.����ʱ��%Type,
  ��ʼip_In   In Zlloginlimit.��ʼip%Type,
  ����ip_In   In Zlloginlimit.����ip%Type,
  ״̬_In     In Zlloginlimit.״̬%Type,
  ˵��_In     In Zlloginlimit.˵��%Type,
  �û���_In   In Zlloginlimit.�û���%Type
) Is
  v_Err_Msg Varchar2(500);
  Err_Item Exception;
  -------------------------------------------------------------------------------------
  --����:��¼���ƹ������/�޸�
  --����:
  --��������_In=1:��� ,������ID_In , ��������_In-����ֵ=�޸�,�贫��ID_In
  --�û���_In-һ��������ԶԶ���û���Ч,ÿ���û���֮����","����
  -------------------------------------------------------------------------------------
Begin
  If ��������_In = 1 Or Id_In Is Null Then
    If �û���_In Is Null Then
      Insert Into Zlloginlimit
        (ID, ��ʼʱ��, ����ʱ��, ��ʼip, ����ip, ״̬, ˵��, �û���)
      Values
        (Zlloginlimit_Id.Nextval, ��ʼʱ��_In, ����ʱ��_In, ��ʼip_In, ����ip_In, ״̬_In, ˵��_In, �û���_In);
    Else
      --��������Ϊ1,����Id_Ϊ��,�������
      Insert Into Zlloginlimit
        (ID, ��ʼʱ��, ����ʱ��, ��ʼip, ����ip, ״̬, ˵��, �û���)
        Select Zlloginlimit_Id.Nextval, ��ʼʱ��_In, ����ʱ��_In, ��ʼip_In, ����ip_In, ״̬_In, ˵��_In, Column_Value
        From Table(f_Str2list(�û���_In)) A;
    End If;
  Else
    --�޸�����
    Update Zlloginlimit
    Set ��ʼʱ�� = ��ʼʱ��_In, ����ʱ�� = ����ʱ��_In, ��ʼip = ��ʼip_In, ����ip = ����ip_In, ״̬ = ״̬_In, ˵�� = ˵��_In, �û��� = �û���_In
    Where ID = Id_In;
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlloginlimit_Edit;
/

--116688:����һ,2017-11-08,��¼���ƹ���ɾ��
Create Or Replace Procedure Zltools.Zl_Zlloginlimit_Delete(Ids_In In Varchar2) Is
  v_Err_Msg Varchar2(500);
  Err_Item Exception;
  -------------------------------------------------------------------------------------
  --����:��¼���ƹ���ɾ��
  --����:Ids_In-�����ַ�������,��������ɾ��,ÿ��ID֮����","��Ϊ���
  -------------------------------------------------------------------------------------
Begin
  Delete Zlloginlimit Where ID In (Select Column_Value From Table(f_Str2list(Ids_In)) A);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlloginlimit_Delete;
/

--116688:����һ,2017-11-08,Ӧ�ó�����Ȩ

Create Table zlTools.zlAppPermission(
    Ӧ�ó��� Varchar2(100),
    �û��� Varchar2(50),
    ��ʼIP Varchar2(20),
    ����IP Varchar2(20),
    ״̬ Number(1),
    ˵�� Varchar2(200));
ALTER TABLE zlTools.zlAppPermission ADD CONSTRAINT zlAppPermission_PK PRIMARY KEY (Ӧ�ó���,�û���) USING INDEX;
Insert into zlTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLAPPPERMISSION','ZLTOOLSTBS','C1');

--116688:����һ,2017-11-08,Ӧ�ó�����Ȩ����/�޸�
Create Or Replace Procedure Zltools.Zl_Zlapppermission_Edit
(
  ��������_In    In Number,
  Ӧ�ó���_In    In Zlapppermission.Ӧ�ó���%Type,
  �û���_In      In Varchar,
  ��ʼip_In      In Zlapppermission.��ʼip%Type,
  ����ip_In      In Zlapppermission.����ip%Type,
  ״̬_In        In Zlapppermission.״̬%Type,
  ˵��_In        In Zlapppermission.˵��%Type,
  Ӧ�ó���new_In In Zlapppermission.Ӧ�ó���%Type,
  �û���new_In   In Varchar
) Is
  v_Err_Msg Varchar2(500);
  Err_Item Exception;
  -------------------------------------------------------------------------------------
  --����:��¼���ƹ������/�޸�
  --����:
  --��������_In=1:��� ,������ID_In , ��������_In-����ֵ=�޸�,�贫��ID_In
  --�û���_In-һ��������ԶԶ���û���Ч,ÿ���û���֮����","����
  -------------------------------------------------------------------------------------
Begin
  If ��������_In = 1 Then
    Insert Into Zlapppermission
      (Ӧ�ó���, �û���, ��ʼip, ����ip, ״̬, ˵��)
      Select Ӧ�ó���_In, Column_Value, ��ʼip_In, ����ip_In, ״̬_In, ˵��_In From Table(f_Str2list(�û���_In)) A;
  Else
    Update Zlapppermission
    Set ��ʼip = ��ʼip_In, ����ip = ����ip_In, ״̬ = ״̬_In, ˵�� = ˵��_In, Ӧ�ó��� = Ӧ�ó���new_In, �û��� = �û���new_In
    Where Ӧ�ó��� = Ӧ�ó���_In And �û��� = �û���_In;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlapppermission_Edit;
/

--116688:����һ,2017-11-08,Ӧ�ó�����Ȩɾ��
Create Or Replace Procedure Zltools.Zl_Zlapppermission_Delete(Ids_In In Varchar2) Is
  v_Err_Msg Varchar2(500);
  Err_Item Exception;
  -------------------------------------------------------------------------------------
  --����:��¼���ƹ���ɾ��
  --����:Ids_In-�����ַ�������,��������ɾ��,��ʽΪ Ӧ�ó���1:�û���1,Ӧ�ó���2:�û���2
  -- �� SQLPLUS:ZLHIS,plsql DEV:ZLHIS
  -------------------------------------------------------------------------------------
Begin
  Delete From Zlapppermission Where (Ӧ�ó���, �û���) In (Select C1, C2 From Table(f_Str2list2(Ids_In)) A);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlapppermission_Delete;
/

--116688:����һ,2017-11-08,�û���¼������
Create Or Replace Trigger zlTools.zl_Trigger_LoginLimit
        After logon On Database
Declare
  --IP_1:ip��ǰ��λ, ip_2:IP��ַ�ĺ�һλ
  v_Ip     Varchar2(25);
  v_Ip_1   Varchar2(20);
  v_Ip_2   Varchar2(5);
  v_User   Varchar2(40);
  v_Date   Date;
  v_Module Varchar2(100);
  n_Count  Number(5);
Begin

  --���û�����ƹ���,�Ͳ�ִ�в���,��ֹ�����ʱ
  Select Count(1)
  Into n_Count
  From (Select 1
         From Zlapppermission
         Where Rownum = 1
         Union All
         Select 1 From Zlloginlimit Where Rownum = 1);

  If n_Count <> 0 Then
    Select Sys_Context('userenv', 'ip_address'), User, Sysdate, Module
    Into v_Ip, v_User, v_Date, v_Module
    From V$session
    Where Audsid = Userenv('sessionid') And Rownum = 1;
    v_Ip_1 := Substr(v_Ip, 1, Instr(v_Ip, '.', 1, 3) - 1);
    v_Ip_2 := Substr(v_Ip, Instr(v_Ip, '.', 1, 3) + 1);

    --����¼����
    Select Count(1)
    Into n_Count
    From Zlloginlimit
    Where ״̬ = 1 And
          ((�û��� = User And Substr(��ʼip, 1, Instr(��ʼip, '.', 1, 3) - 1) = v_Ip_1 And
          v_Ip_2 Between Substr(��ʼip, Instr(��ʼip, '.', 1, 3) + 1) And Substr(����ip, Instr(����ip, '.', 1, 3) + 1) And
          v_Date Between ��ʼʱ�� And ����ʱ��) Or (�û��� = User And ��ʼip Is Null And ��ʼʱ�� Is Null) Or
          (Substr(��ʼip, 1, Instr(��ʼip, '.', 1, 3) - 1) = v_Ip_1 And
          v_Ip_2 Between Substr(��ʼip, Instr(��ʼip, '.', 1, 3) + 1) And Substr(����ip, Instr(����ip, '.', 1, 3) + 1) And
          �û��� Is Null And ��ʼʱ�� Is Null) Or (v_Date Between ��ʼʱ�� And ����ʱ�� And �û��� Is Null And ��ʼip Is Null) Or
          (�û��� = User And Substr(��ʼip, 1, Instr(��ʼip, '.', 1, 3) - 1) = v_Ip_1 And
          v_Ip_2 Between Substr(��ʼip, Instr(��ʼip, '.', 1, 3) + 1) And Substr(����ip, Instr(����ip, '.', 1, 3) + 1) And
          ��ʼʱ�� Is Null) Or (�û��� = User And v_Date Between ��ʼʱ�� And ����ʱ�� And ��ʼip Is Null) Or
          (Substr(��ʼip, 1, Instr(��ʼip, '.', 1, 3) - 1) = v_Ip_1 And
          v_Ip_2 Between Substr(��ʼip, Instr(��ʼip, '.', 1, 3) + 1) And Substr(����ip, Instr(����ip, '.', 1, 3) + 1) And
          v_Date Between ��ʼʱ�� And ����ʱ�� And �û��� Is Null));

    If n_Count > 0 Then
      Raise_Application_Error(-20001, '��ǰ�û�����ֹ��¼���ݿ⣬����ϵ����Ա��');
    End If;

    --���Ӧ����Ȩ
    Select Count(1) Into n_Count From Zlapppermission Where Ӧ�ó��� = v_Module And ״̬ = 1;

    If n_Count > 0 Then
      Select Count(1)
      Into n_Count
      From Zlapppermission
      Where ״̬ = 1 And
            ((Ӧ�ó��� = v_Module And �û��� = v_User And ��ʼip Is Null) Or
            (Ӧ�ó��� = v_Module And �û��� = v_User And Substr(��ʼip, 1, Instr(��ʼip, '.', 1, 3) - 1) = v_Ip_1 And
            v_Ip_2 Between Substr(��ʼip, Instr(��ʼip, '.', 1, 3) + 1) And Substr(����ip, Instr(����ip, '.', 1, 3) + 1)));
      If n_Count = 0 Then
        Raise_Application_Error(-20002, '��ǰ�û�����ʹ�ø�Ӧ�õ�¼���ݿ⣬����ϵ����Ա��');
      End If;
    End If;
  End If;
End Zl_Trigger_Loginlimit;
/
--116691:����һ,2017-11-15,Զ�̿��Ʋ���
Insert Into zlParameters(ID,ϵͳ,ģ��,˽��,����,��Ȩ,�̶�,����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵��)
    Select zlParameters_ID.Nextval,-Null,-null,-Null,1,-Null,-Null,A.* From (
    Select ����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵�� From zlParameters Where 1 = 0 Union All
    Select 0,0,31,'����Զ�̿���',Null,'1001','���뵼��̨��,�Ƿ���������ԱԶ�̿���','Զ�̿���',Null,'�����������Աͨ��Զ�̿��������������֤�������ʾ',Null From dual Union All
    Select ����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵�� From ZLPARAMETERS Where 1 = 0) A;

--113395:����,2017-11-15,�����߰�������Ȩ
Create Table Zltools.ZLSVRFuncs(
       ��� varchar2(6),
       ���� varchar2(30),
       ���� number(3),
       ˵�� varchar2(250),
       ȱʡ number(1));
Alter Table Zltools.ZLSVRFuncs Add Constraint ZLSVRFuncs_UQ_���� Unique(���, ����) Using Index;
Alter Table Zltools.ZLSVRFuncs Modify ��� Constraint ZLSVRFuncs_NN_��� Not Null;
Alter Table Zltools.ZLSVRFuncs Add Constraint ZLSVRFuncs_FK_��� Foreign Key(���) References Zlsvrtools(���) On Delete Cascade;
Alter Table Zltools.ZLMgrGrant Modify(���� Varchar2(4000));
Insert Into zlTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLSVRFUNCS','ZLTOOLSTBS','A2');
Insert Into Zltools.Zlsvrfuncs(���, ����, ����, ˵��, ȱʡ)
  Select ���, '����', '0', Null, 1 From Zlsvrtools Where �ϼ� Is Not Null;
Insert Into Zltools.Zlsvrfuncs(���, ����, ����, ˵��, ȱʡ) Values ('0307', '�ļ�����������', 1, '���ڶԿͻ�����������ķ������������ã���������������ͣ�����������͵�', 1);
Insert Into Zltools.Zlsvrfuncs(���, ����, ����, ˵��, ȱʡ) Values ('0307', '�ļ���������', 2, '���ڶԿͻ����������õ������ļ����й�����ҪΪ�����ļ�����ɾ��', 1);
Insert Into Zltools.Zlsvrfuncs(���, ����, ����, ˵��, ȱʡ) Values ('0307', '�ͻ�����������', 3, '�������ÿͻ���������Ϣ�������Ƿ�������������ʱ������', 1);

--000000:����һ,2017-11-30,����������Ӷ����еļ��
Create Or Replace Function Zltools.Zl_Checkobject
(
  n_Type        In Number, --1=��,2=�ֶ�,3=Լ��,4=���� ,5=����
  v_Object_Name In Varchar2,
  v_Table_Name  In Varchar2 := Null --����n_Type=2ʱ����Ҫ���� 
) Return Number Authid Current_User As
  --���ܣ���ִ���ߵ���ݼ��ָ�����ָ�������Ƿ���� 
  --����ֵ��>0��ʾ���ڣ�0��ʾ������ 
  n_Count Number(5);
Begin
  If n_Type = 1 Then
    If v_Table_Name Is Null Then
      Select Count(1) Into n_Count From User_Tables Where Table_Name = Upper(v_Object_Name);
    Else
      Select Count(1) Into n_Count From User_Tables Where Table_Name = Upper(v_Table_Name);
    End If;
  Elsif n_Type = 2 Then
    Select Count(1)
    Into n_Count
    From User_Tab_Columns
    Where Table_Name = Upper(v_Table_Name) And Column_Name = Upper(v_Object_Name);
  
  Elsif n_Type = 3 Then
    Select Count(1) Into n_Count From User_Constraints Where Constraint_Name = Upper(v_Object_Name);
  
  Elsif n_Type = 4 Then
    Select Count(1) Into n_Count From User_Indexes Where Index_Name = Upper(v_Object_Name);
  
  Elsif n_Type = 5 Then
    Select Count(1) Into n_Count From User_Sequences Where Sequence_Name = Upper(v_Object_Name);
  End If;

  Return n_Count;
End Zl_Checkobject;
/