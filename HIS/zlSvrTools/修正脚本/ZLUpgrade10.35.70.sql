----10.35.60---��10.35.70
--000000:��˶,2017-9-11,�淶���
ALTER TABLE ZLTOOLS.ZLUPGRADESERVER Rename CONSTRAINT ZLUPGRADESERVER_PK_��� to ZLUPGRADESERVER_PK;
alter index  ZLTOOLS.ZLUPGRADESERVER_PK_��� rename to ZLUPGRADESERVER_PK;
--000000:����,2017-8-07,�޸�����
Delete from ZLTools.zlTables Where ����='ZLTools.zlTables';
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLTABLES','ZLTOOLSTBS','A1');
--111526:����,2017-7-5,��־��ͣ����
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLLOGCONFIG','ZLTOOLSTBS','A2');
--112138:����,2017-8-31,�ر���������̨
Insert Into ZLTools.zlOptions(������,������,����ֵ,ȱʡֵ,����˵��) Values(24, '����ر������ĵ���̨', '0','0', '���Ƶ�����̨��������ʱ���ܷ�رյ���̨��');
--104763:����,2017-9-7,��ʾͣ�ñ���
Insert Into ZLTools.zlParameters(ID,ϵͳ,ģ��,˽��,����,��Ȩ,�̶�,����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵��)
Select zlParameters_ID.Nextval,-Null,-Null,1,-Null,-Null,-Null,A.* From (
Select ����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵�� From zlParameters Where 1 = 0 Union All
Select 0,0,29,'��ʾͣ�ñ���',Null,'0','�Ƿ���ʾ�Ѿ���ͣ�õı���','0��NUll������ʾͣ�ñ���1����ʾͣ�ñ���',NULL,NULL,NULL From Dual Union All
Select ����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵�� From ZLPARAMETERS Where 1 = 0) A;
--000000:��˶,2017-8-07,��Ч��ɾ��
Drop table ZLTools.ZLPROCEDURENOTE;
Drop public synonym ZLPROCEDURENOTE;
Drop Procedure zlTools.Zl_zlProcedureNote_Delete;
Drop public synonym Zl_zlProcedureNote_Delete;
Drop Procedure zlTools.Zl_zlProcedureNote_Update;
Drop public synonym Zl_zlProcedureNote_Update;
Create Or Replace Procedure zlTools.Zl_Zlprocedure_Delete
(
  Id_In           In Zlprocedure.Id%Type
) Is
Begin
  Delete zlProcedureText Where ����ID=Id_In;
  Delete zlProcedure Where ID = Id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlprocedure_Delete;
/
--111911:��˶,2017-7-28,7z���������������ֿͻ�����������
Drop Procedure Zltools.Zl_Zlupgradeserver_Delete;
Drop public synonym Zl_Zlupgradeserver_Delete;
Drop Procedure Zltools.Zl_Zlupgradeserver_Insert;
Drop public synonym Zl_Zlupgradeserver_Insert;
Drop Procedure Zltools.Zlreginfo_Defaultserver;
Drop public synonym Zlreginfo_Defaultserver;
Alter table ZLTOOLS.ZLUPGRADESERVER drop constraint ZLUPGRADESERVER_UQ_λ��;
Create Or Replace Procedure Zltools.Zl_Zlupgradeserver_Update
(
  ģʽ_In     In Number,
  ���_In     In Zlupgradeserver.���%Type,
  ����_In     In Zlupgradeserver.����%Type := Null,
  λ��_In     In Zlupgradeserver.λ��%Type := Null,
  �û���_In   In Zlupgradeserver.�û���%Type := Null,
  ����_In     In Zlupgradeserver.����%Type := Null,
  �˿�_In     In Zlupgradeserver.�˿�%Type := Null,
  �Ƿ�����_In In Zlupgradeserver.�Ƿ�����%Type := Null,
  �Ƿ�ȱʡ_In In Zlupgradeserver.�Ƿ�ȱʡ%Type := Null,
  �Ƿ��ռ�_In In Zlupgradeserver.�Ƿ��ռ�%Type := Null,
  �ռ�����_In In Zlupgradeserver.�ռ�����%Type := Null,
  ����ex_In   In Zlupgradeserver.����%Type := Null
) Is
  --ģʽ_IN=0-������1-�޸�,11-ֻ�޸��Ƿ��������Ƿ�ȱʡ���Ƿ��ռ����ռ������ֶ� ,2-ɾ��
  v_�ռ����� Zlupgradeserver.�ռ�����%Type;
  n_�Ƿ����� Zlupgradeserver.�Ƿ�����%Type;
  n_���     Zlupgradeserver.���%Type;
Begin
  --�������µ�ȱʡ���������ǰ��ȱʡ
  If �Ƿ�ȱʡ_In = 1 Then
    Update Zlupgradeserver Set �Ƿ�ȱʡ = 0 Where Nvl(�Ƿ�ȱʡ, 0) = 1;
  End If;
  If �Ƿ��ռ�_In = 1 Then
    Select Max(v_�ռ�����) Into v_�ռ����� From Zlupgradeserver Where Nvl(�Ƿ��ռ�, 0) = 1;
    Update Zlupgradeserver Set �Ƿ��ռ� = 0, �ռ����� = Null Where Nvl(�Ƿ��ռ�, 0) = 1;
  End If;
  If Nvl(ģʽ_In, 0) = 0 Or Nvl(���_In, 0) = 0 Then
    Select Nvl(Max(���), 0) + 1 Into n_��� From Zlupgradeserver;
    Insert Into Zlupgradeserver
      (���, ����, λ��, �û���, ����, �˿�, �Ƿ�����, �Ƿ�ȱʡ, �Ƿ��ռ�, �ռ�����, ����)
    Values
      (n_���, ����_In, λ��_In, �û���_In, ����_In, �˿�_In, �Ƿ�����_In, �Ƿ�ȱʡ_In, �Ƿ��ռ�_In, �ռ�����_In, 0);
  Elsif Nvl(ģʽ_In, 0) = 2 Then
    Delete From Zlupgradeserver Where ��� = ���_In;
    Update Zlclients Set �����ļ������� = Null Where �����ļ������� = ���_In;
  Else
    Select Max(�Ƿ�����) Into n_�Ƿ����� From Zlupgradeserver Where ��� = ���_In;
    If Nvl(ģʽ_In, 0) = 1 Then
      Update Zlupgradeserver
      Set ���� = ����_In, λ�� = λ��_In, �û��� = �û���_In, ���� = ����_In, �˿� = �˿�_In, �Ƿ����� = �Ƿ�����_In, �Ƿ�ȱʡ = �Ƿ�ȱʡ_In, �Ƿ��ռ� = �Ƿ��ռ�_In,
          �ռ����� = �ռ�����_In
      Where ��� = ���_In;
    Else
      Update Zlupgradeserver
      Set �Ƿ����� = �Ƿ�����_In, �Ƿ�ȱʡ = �Ƿ�ȱʡ_In, �Ƿ��ռ� = �Ƿ��ռ�_In, �ռ����� = v_�ռ�����
      Where ��� = ���_In;
    End If;
    --����������������Ϊ�������������������������������
    If Nvl(n_�Ƿ�����, 0) = 1 And Nvl(�Ƿ�����_In, 0) = 0 Then
      Update Zlclients Set �����ļ������� = Null Where �����ļ������� = ���_In;
    End If;
  End If;
  --�Զ�����ZLRegINFO�����ݣ���֤����ǰ����
  If Nvl(�Ƿ�ȱʡ_In, 0) = 1 Then
    Insert Into Zltools.Zlreginfo
      (��Ŀ, ����)
      Select '��������', Null From Dual Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = '��������');
    If Nvl(����_In, 0) = 0 Then
      Insert Into Zltools.Zlreginfo
        (��Ŀ, ����)
        Select '������Ŀ¼0', Null
        From Dual
        Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = '������Ŀ¼0')
        Union All
        Select '�����û�0', Null
        From Dual
        Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = '�����û�0')
        Union All
        Select '��������0', Null
        From Dual
        Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = '��������0')
        Union All
        Select '������Ŀ¼', Null
        From Dual
        Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = '������Ŀ¼')
        Union All
        Select '�����û�', Null
        From Dual
        Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = '�����û�')
        Union All
        Select '��������', '' From Dual Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = '��������');
      Update Zltools.Zlreginfo Set ���� = '0' Where ��Ŀ = '��������';
      Update Zltools.Zlreginfo Set ���� = λ��_In Where ��Ŀ = '������Ŀ¼0';
      Update Zltools.Zlreginfo Set ���� = �û���_In Where ��Ŀ = '�����û�0';
      Update Zltools.Zlreginfo Set ���� = ����ex_In Where ��Ŀ = '��������0';
      Update Zltools.Zlreginfo Set ���� = λ��_In Where ��Ŀ = '������Ŀ¼';
      Update Zltools.Zlreginfo Set ���� = �û���_In Where ��Ŀ = '�����û�';
      Update Zltools.Zlreginfo Set ���� = ����ex_In Where ��Ŀ = '��������';
    Else
      Insert Into Zltools.Zlreginfo
        (��Ŀ, ����)
        Select 'FTP������0', Null
        From Dual
        Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = 'FTP������0')
        Union All
        Select 'FTP�û�0', Null
        From Dual
        Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = 'FTP�û�0')
        Union All
        Select 'FTP����0', Null
        From Dual
        Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = 'FTP����0')
        Union All
        Select 'FTP�˿�0', Null
        From Dual
        Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = 'FTP�˿�0')
        Union All
        Select 'FTP������', Null
        From Dual
        Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = 'FTP������')
        Union All
        Select 'FTP�û�', Null
        From Dual
        Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = 'FTP�û�')
        Union All
        Select 'FTP����', Null
        From Dual
        Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = 'FTP����')
        Union All
        Select 'FTP�˿�', Null From Dual Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = 'FTP�˿�');
      Update Zltools.Zlreginfo Set ���� = '1' Where ��Ŀ = '��������';
      Update Zltools.Zlreginfo Set ���� = λ��_In Where ��Ŀ = 'FTP������0';
      Update Zltools.Zlreginfo Set ���� = �û���_In Where ��Ŀ = 'FTP�û�0';
      Update Zltools.Zlreginfo Set ���� = ����ex_In Where ��Ŀ = 'FTP����0';
      Update Zltools.Zlreginfo Set ���� = �˿�_In Where ��Ŀ = 'FTP�˿�0';
      Update Zltools.Zlreginfo Set ���� = λ��_In Where ��Ŀ = 'FTP������';
      Update Zltools.Zlreginfo Set ���� = �û���_In Where ��Ŀ = 'FTP�û�';
      Update Zltools.Zlreginfo Set ���� = ����ex_In Where ��Ŀ = 'FTP����';
      Update Zltools.Zlreginfo Set ���� = �˿�_In Where ��Ŀ = 'FTP�˿�';
    End If;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zlupgradeserver_Update;
/
--102323:����,2017-7-13,Ϊ��̨��ҵ�����������ʱ�䵥λ.��.��.����
Alter Table Zltools.Zlautojobs Add ʱ�䵥λ Varchar2(5);

--102323:����,2017-7-13,Ϊ��̨��ҵ�����������ʱ�䵥λ.��.��.����
CREATE OR REPLACE Procedure zltools.Zl_Jobsubmit
(
  Job_System In Integer,
  Job_Kind   In Integer,
  Job_Odd    In Integer
) Authid Current_User Is
  v_Content   Varchar2(200);
  v_Parameter Varchar2(200);
  v_Paraitem  Varchar2(200);
  v_Starttime Date;
  v_Cyclekeep Integer;
  v_Jobnum    Number := 0;
  v_What      Varchar2(1000);
  v_Nextdate  Date;
  v_Interval  Varchar2(1000);
  v_Timeunit  Varchar2(5);
  v_Week      Varchar2(10);
  v_Day       Varchar2(1);
Begin
  Select ����, ����, ִ��ʱ��, ���ʱ��, ʱ�䵥λ
  Into v_Content, v_Parameter, v_Starttime, v_Cyclekeep, v_Timeunit
  From Zlautojobs
  Where Nvl(ϵͳ, 0) = Nvl(Job_System, 0) And ���� = Job_Kind And ��� = Job_Odd;
  v_What := '';

  If Length(v_Parameter) > 0 Then
    Loop
      If Instr(v_Parameter, ';') > 0 Then
        v_Paraitem  := Substr(v_Parameter, 1, Instr(v_Parameter, ';') - 1);
        v_Parameter := Substr(v_Parameter, Instr(v_Parameter, ';') + 1);
      Else
        v_Paraitem := v_Parameter;
      End If;
    
      v_What := v_What || ',' || Substr(v_Paraitem, Instr(v_Paraitem, ',') + 1);
      Exit When Instr(v_Parameter, ';') = 0;
    End Loop;
  End If;

  If Length(v_What) <> 0 Then
    v_What := v_Content || '(' || Substr(v_What, 2) || ');';
  Else
    v_What := v_Content || ';';
  End If;

  If v_Timeunit = '��' Then
    If To_Char(Sysdate, 'HH24:MI:SS') >= To_Char(v_Starttime, 'HH24:MI:SS') Then
      v_Nextdate := v_Starttime + 1;
    Else
      v_Nextdate := v_Starttime;
    End If;
    v_Interval := 'trunc(Sysdate)+' || v_Cyclekeep || '+' || To_Char(v_Starttime, 'HH24') || '/24' || '+' ||
                  To_Char(v_Starttime, 'MI') || '/(24*60)' || '+' || To_Char(v_Starttime, 'SS') || '/(24*60*60)';
  Elsif v_Timeunit = '��' Then
    If To_Char(Sysdate, 'yyyy-MM-dd HH24:MI:SS') >= To_Char(v_Starttime, 'yyyy-MM-dd HH24:MI:SS') Then
      v_Nextdate := v_Starttime + 7;
    Else
      v_Nextdate := v_Starttime;
    End If;
    Select To_Char(v_Starttime, 'DY') Into v_Week From Dual;
    v_Interval := 'TRUNC(next_day(sysdate,''' || v_Week || '''))+7*(' || v_Cyclekeep || '-1)' || '+' ||
                  To_Char(v_Starttime, 'HH24') || '/24' || '+' || To_Char(v_Starttime, 'MI') || '/(24*60)' || '+' ||
                  To_Char(v_Starttime, 'SS') || '/(24*60*60)';
  Elsif v_Timeunit = '��' Then
    If To_Char(Sysdate, 'yyyy-MM-dd HH24:MI:SS') >= To_Char(v_Starttime, 'yyyy-MM-dd HH24:MI:SS') Then
      v_Nextdate := Add_Months(v_Starttime, 1);
    Else
      v_Nextdate := v_Starttime;
    End If;
    If To_Char(v_Starttime, 'dd') <= 28 Then
      v_Interval := 'TRUNC(ADD_MONTHS(sysdate,' || v_Cyclekeep || '))+' || To_Char(v_Starttime, 'HH24') || '/24' || '+' ||
                    To_Char(v_Starttime, 'MI') || '/(24*60)' || '+' || To_Char(v_Starttime, 'SS') || '/(24*60*60)';
    Else
      Select To_Char(Last_Day(Sysdate), 'dd') - To_Char(v_Starttime, 'dd') Into v_Day From Dual;
      v_Interval := 'TRUNC(last_day(ADD_MONTHS(sysdate,' || v_Cyclekeep || '))-' || v_Day || ')+' ||
                    To_Char(v_Starttime, 'HH24') || '/24' || '+' || To_Char(v_Starttime, 'MI') || '/(24*60)' || '+' ||
                    To_Char(v_Starttime, 'SS') || '/(24*60*60)';
    End If;
  Else
    If To_Char(Sysdate, 'yyyy-MM-dd HH24:MI:SS') >= To_Char(v_Starttime, 'yyyy-MM-dd HH24:MI:SS') Then
      v_Nextdate := Add_Months(v_Starttime, 3);
    Else
      v_Nextdate := v_Starttime;
    End If;
    If To_Char(v_Starttime, 'dd') <= 28 Then
      v_Interval := 'TRUNC(ADD_MONTHS(SYSDATE,3*' || v_Cyclekeep || '))+' || To_Char(v_Starttime, 'HH24') || '/24' || '+' ||
                    To_Char(v_Starttime, 'MI') || '/(24*60)' || '+' || To_Char(v_Starttime, 'SS') || '/(24*60*60)';
    Else
      Select To_Char(Last_Day(Sysdate), 'dd') - To_Char(v_Starttime, 'dd') Into v_Day From Dual;
      v_Interval := 'TRUNC(last_day(ADD_MONTHS(SYSDATE,3*' || v_Cyclekeep || '))-' || v_Day || ')+' ||
                    To_Char(v_Starttime, 'HH24') || '/24' || '+' || To_Char(v_Starttime, 'MI') || '/(24*60)' || '+' ||
                    To_Char(v_Starttime, 'SS') || '/(24*60*60)';
    End If;
  End If;

  --�ύ��ҵ 
  Dbms_Job.Submit(v_Jobnum, v_What, v_Nextdate, v_Interval);

  Update Zlautojobs
  Set ��ҵ�� = v_Jobnum
  Where Nvl(ϵͳ, 0) = Nvl(Job_System, 0) And ���� = Job_Kind And ��� = Job_Odd;
End Zl_Jobsubmit;
/

--102323:����,2017-7-13,Ϊ��̨��ҵ�����������ʱ�䵥λ.��.��.����
Create Or Replace Procedure Zltools.Zl_Jobremove
(
  Job_System In Integer,
  Job_Kind   In Integer,
  Job_Odd    In Integer
) Authid Current_User Is
  v_Jobnum Number := 0;
Begin
  Select ��ҵ��
  Into v_Jobnum
  From Zlautojobs
  Where Nvl(ϵͳ, 0) = Nvl(Job_System, 0) And ���� = Job_Kind And ��� = Job_Odd;
  --ɾ����ҵ 
  Dbms_Job.Remove(v_Jobnum);

  Update Zlautojobs Set ��ҵ�� = Null Where Nvl(ϵͳ, 0) = Nvl(Job_System, 0) And ���� = Job_Kind And ��� = Job_Odd;
End Zl_Jobremove;
/

--102323:����,2017-7-13,Ϊ��̨��ҵ�����������ʱ�䵥λ.��.��.����
Create Or Replace Procedure Zltools.Zl_Jobrun
(
  Job_System In Integer,
  Job_Kind   In Integer,
  Job_Odd    In Integer
) Authid Current_User Is
  v_Jobnum Number := 0;
Begin
  Select ��ҵ��
  Into v_Jobnum
  From Zlautojobs
  Where Nvl(ϵͳ, 0) = Nvl(Job_System, 0) And ���� = Job_Kind And ��� = Job_Odd;
  --ִ����ҵ 
  Dbms_Job.Run(v_Jobnum);
End Zl_Jobrun;
/

--102323:����,2017-7-13,Ϊ��̨��ҵ�����������ʱ�䵥λ.��.��.����
CREATE OR REPLACE Procedure zltools.Zl_Jobchange
(
  Job_System In Integer,
  Job_Kind   In Integer,
  Job_Odd    In Integer
) Authid Current_User Is
  v_Content   Varchar2(200);
  v_Parameter Varchar2(200);
  v_Paraitem  Varchar2(200);
  v_Starttime Date;
  v_Cyclekeep Integer;
  v_Jobnum    Number := 0;
  v_What      Varchar2(1000);
  v_Nextdate  Date;
  v_Interval  Varchar2(1000);
  v_Timeunit  Varchar2(5);
  v_Week      Varchar2(10);
  v_Day       Varchar2(1);
Begin
  Select ����, ����, ִ��ʱ��, ���ʱ��, ��ҵ��, ʱ�䵥λ
  Into v_Content, v_Parameter, v_Starttime, v_Cyclekeep, v_Jobnum, v_Timeunit
  From Zlautojobs
  Where Nvl(ϵͳ, 0) = Nvl(Job_System, 0) And ���� = Job_Kind And ��� = Job_Odd;
  v_What := '';

  If Length(v_Parameter) > 0 Then
    Loop
      If Instr(v_Parameter, ';') > 0 Then
        v_Paraitem  := Substr(v_Parameter, 1, Instr(v_Parameter, ';') - 1);
        v_Parameter := Substr(v_Parameter, Instr(v_Parameter, ';') + 1);
      Else
        v_Paraitem := v_Parameter;
      End If;
    
      v_What := v_What || ',' || Substr(v_Paraitem, Instr(v_Paraitem, ',') + 1);
      Exit When Instr(v_Parameter, ';') = 0;
    End Loop;
  End If;

  If Length(v_What) <> 0 Then
    v_What := v_Content || '(' || Substr(v_What, 2) || ');';
  Else
    v_What := v_Content || ';';
  End If;

  If v_Timeunit = '��' Then
    If To_Char(Sysdate, 'HH24:MI:SS') >= To_Char(v_Starttime, 'HH24:MI:SS') Then
      v_Nextdate := v_Starttime + 1;
    Else
      v_Nextdate := v_Starttime;
    End If;
    v_Interval := 'trunc(Sysdate)+' || v_Cyclekeep || '+' || To_Char(v_Starttime, 'HH24') || '/24' || '+' ||
                  To_Char(v_Starttime, 'MI') || '/(24*60)' || '+' || To_Char(v_Starttime, 'SS') || '/(24*60*60)';
  Elsif v_Timeunit = '��' Then
    If To_Char(Sysdate, 'yyyy-MM-dd HH24:MI:SS') >= To_Char(v_Starttime, 'yyyy-MM-dd HH24:MI:SS') Then
      v_Nextdate := v_Starttime + 7;
    Else
      v_Nextdate := v_Starttime;
    End If;
    Select To_Char(v_Starttime, 'DY') Into v_Week From Dual;
    v_Interval := 'TRUNC(next_day(sysdate,''' || v_Week || '''))+7*(' || v_Cyclekeep || '-1)' || '+' ||
                  To_Char(v_Starttime, 'HH24') || '/24' || '+' || To_Char(v_Starttime, 'MI') || '/(24*60)' || '+' ||
                  To_Char(v_Starttime, 'SS') || '/(24*60*60)';
  Elsif v_Timeunit = '��' Then
    If To_Char(Sysdate, 'yyyy-MM-dd HH24:MI:SS') >= To_Char(v_Starttime, 'yyyy-MM-dd HH24:MI:SS') Then
      v_Nextdate := Add_Months(v_Starttime, 1);
    Else
      v_Nextdate := v_Starttime;
    End If;
    If To_Char(v_Starttime, 'dd') <= 28 Then
      v_Interval := 'TRUNC(ADD_MONTHS(sysdate,' || v_Cyclekeep || '))+' || To_Char(v_Starttime, 'HH24') || '/24' || '+' ||
                    To_Char(v_Starttime, 'MI') || '/(24*60)' || '+' || To_Char(v_Starttime, 'SS') || '/(24*60*60)';
    Else
      Select To_Char(Last_Day(Sysdate), 'dd') - To_Char(v_Starttime, 'dd') Into v_Day From Dual;
      v_Interval := 'TRUNC(last_day(ADD_MONTHS(sysdate,' || v_Cyclekeep || '))-' || v_Day || ')+' ||
                    To_Char(v_Starttime, 'HH24') || '/24' || '+' || To_Char(v_Starttime, 'MI') || '/(24*60)' || '+' ||
                    To_Char(v_Starttime, 'SS') || '/(24*60*60)';
    End If;
  Else
    If To_Char(Sysdate, 'yyyy-MM-dd HH24:MI:SS') >= To_Char(v_Starttime, 'yyyy-MM-dd HH24:MI:SS') Then
      v_Nextdate := Add_Months(v_Starttime, 3);
    Else
      v_Nextdate := v_Starttime;
    End If;
    If To_Char(v_Starttime, 'dd') <= 28 Then
      v_Interval := 'TRUNC(ADD_MONTHS(SYSDATE,3*' || v_Cyclekeep || '))+' || To_Char(v_Starttime, 'HH24') || '/24' || '+' ||
                    To_Char(v_Starttime, 'MI') || '/(24*60)' || '+' || To_Char(v_Starttime, 'SS') || '/(24*60*60)';
    Else
      Select To_Char(Last_Day(Sysdate), 'dd') - To_Char(v_Starttime, 'dd') Into v_Day From Dual;
      v_Interval := 'TRUNC(last_day(ADD_MONTHS(SYSDATE,3*' || v_Cyclekeep || '))-' || v_Day || ')+' ||
                    To_Char(v_Starttime, 'HH24') || '/24' || '+' || To_Char(v_Starttime, 'MI') || '/(24*60)' || '+' ||
                    To_Char(v_Starttime, 'SS') || '/(24*60*60)';
    End If;
  End If;

  --�޸���ҵ 
  Dbms_Job.Change(v_Jobnum, v_What, v_Nextdate, v_Interval);
End Zl_Jobchange;
/

--111523:����,2017-7-20,����¼������־��SQL����Ϊ����,����Ӳ����
Create Or Replace Procedure Zltools.Zl_Zlerrorlog_Insert
(
  ����վ_In    Zlerrorlog.����վ%Type,
  ����_In      Zlerrorlog.����%Type,
  �������_In  Zlerrorlog.�������%Type,
  ������Ϣ_In  Zlerrorlog.������Ϣ%Type,
  Sessionid_In Number := Null
) Is
Begin
  If Sessionid_In Is Null Then
    Insert Into Zlerrorlog
      (�Ự��, �û���, ����վ, ʱ��, ����, �������, ������Ϣ)
      Select Sid, User, ����վ_In, Sysdate, ����_In, �������_In, ������Ϣ_In
      From V$session
      Where Audsid = Userenv('SessionID');
  Else
    Insert Into Zlerrorlog
      (�Ự��, �û���, ����վ, ʱ��, ����, �������, ������Ϣ)
      Select Sid, User, ����վ_In, Sysdate, ����_In, �������_In, ������Ϣ_In
      From GV$session
      Where Audsid = Sessionid_In;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zlerrorlog_Insert;
/

--111523:����,2017-7-20,����¼������־SQL����Ϊ����,����Ӳ����
Create Or Replace Procedure Zltools.Zl_Zldiarylog_Insert
(
  ����վ_In    Zldiarylog.����վ%Type,
  ������_In    Zldiarylog.������%Type,
  ������_In    Zldiarylog.������%Type,
  ��������_In  Zldiarylog.��������%Type,
  Sessionid_In Number := Null
) Is
Begin
  If Sessionid_In Is Null Then
    Insert Into Zldiarylog
      (�Ự��, �û���, ����վ, ������, ������, ��������, ����ʱ��)
      Select Sid + Serial#, User, ����վ_In, ������_In, ������_In, ��������_In, Sysdate
      From V$session
      Where Audsid = Userenv('SessionID') And Machine Is Not Null;
  Else
    Insert Into Zldiarylog
      (�Ự��, �û���, ����վ, ������, ������, ��������, ����ʱ��)
      Select Sid + Serial#, User, ����վ_In, ������_In, ������_In, ��������_In, Sysdate
      From GV$session
      Where Audsid = Sessionid_In And Machine Is Not Null;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zldiarylog_Insert;
/

--111523:����,2017-7-20,������������־SQL����Ϊ����,����Ӳ����
Create Or Replace Procedure Zltools.Zl_Zldiarylog_Update
(
  ����վ_In    Zldiarylog.����վ%Type,
  ������_In    Zldiarylog.������%Type,
  ������_In    Zldiarylog.������%Type,
  �˳�ԭ��_In  Zldiarylog.�˳�ԭ��%Type,
  Sessionid_In Number := Null
) Is
  n_�Ự�� zldiarylog.�Ự��%type;
  v_�û��� zldiarylog.�û���%type;
Begin
  If Sessionid_In Is Null Then
    Select Sid + Serial#, User
    Into n_�Ự��, v_�û���
    From V$session
    Where Audsid = Userenv('SessionID') And Machine Is Not Null;
  Else
    Select Sid + Serial#, User
    Into n_�Ự��, v_�û���
    From GV$session
    Where Audsid = Sessionid_In And Machine Is Not Null;
  End If;
  Update Zldiarylog
  Set �˳�ԭ�� = �˳�ԭ��_In, �˳�ʱ�� = Sysdate
  Where �˳�ԭ�� Is Null And �û��� = v_�û��� And ����վ = ����վ_In And �Ự�� = n_�Ự�� And ������ = ������_In And ������ = ������_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zldiarylog_Update;
/

--100682:����,2017-8-29,��Ҫ������־��¼
Create Table Zltools.zlAuditLog(
�û��� Varchar2(30),
����վ Varchar2(50),
����ʱ�� Date,
�������� Number(2),
����ģ���� Varchar2(18),
�������� Varchar2(1024),
����˵�� Varchar2(256));
Alter Table zltools.zlAuditLog Add Constraint zlAuditLog_PK Primary Key (�û���,����վ,����ʱ��,����ģ����) Using Index;
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLAUDITLOG','ZLTOOLSTBS','B3');

--100682:����,2017-8-29,��Ҫ������־��¼
Create Or Replace Procedure Zltools.Zl_Zlauditlog_Insert
(
  �û���_In       Zlauditlog.�û���%Type,
  ����վ_In       Zlauditlog.����վ%Type,
  ��������_In     Zlauditlog.��������%Type, --1-������2-�޸ģ�3-ɾ��
  ����ģ����_In Zlauditlog.����ģ����%Type,
  ��������_In     Zlauditlog.��������%Type,
  ����˵��_In     Zlauditlog.����˵��%Type --������¼�����ṩ������Ա����ı�ע��Ϣ
) Is
Begin
  Insert Into Zlauditlog
    (�û���, ����վ, ����ʱ��, ��������, ����ģ����, ��������, ����˵��)
  Values
    (�û���_In, ����վ_In, Sysdate, ��������_In, ����ģ����_In, ��������_In, ����˵��_In);
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zlauditlog_Insert;
/

--111526:����,2017-7-5,��־��ͣ����
Create Table Zltools.ZLLogConfig(
    ��� number(4),
    ���� varchar(20),
    ˵�� varchar2(500));
Alter Table Zltools.ZLLogConfig Add Constraint ZLLogConfig_Pk Primary Key(���) Using Index;

--111538:����,2017-7-25,���RAC������Zl_Autologprocess�е�v$Session��ΪGv$Session
Create Or Replace Procedure Zltools.Zl_Autologprocess As
  --���ܣ� 
  --   1.�Զ����������־�ʹ�����־������� 
  --   2.���쳣��������־���б�� 
  v_Count Number;
  v_Limit Number;
Begin
  --ɾ�������������־ 
  Select Count(*) Into v_Count From Zldiarylog;
  Begin
    Select Nvl(To_Number(����ֵ), 0) Into v_Limit From Zloptions Where ������ = 2;
  Exception
    When Others Then
      v_Limit := 10000;
  End;
  If v_Count > v_Limit Then
    Delete From Zldiarylog
    Where Rowid In (Select Id
                    From (Select Rowid As Id From Zldiarylog Group By ����ʱ��, Rowid)
                    Where Rownum < v_Count - v_Limit + 1);
  End If;

  --���쳣�˳���������־��¼���д��� 
  Update Zldiarylog
  Set �˳�ԭ�� = 2, �˳�ʱ�� = Sysdate
  Where �˳�ԭ�� Is Null And �Ự�� Not In (Select Sid + Serial# From Gv$session Where User# <> 0);

  --ɾ������Ĵ�����־ 
  Select Count(*) Into v_Count From Zlerrorlog;
  Begin
    Select Nvl(To_Number(����ֵ), 0) Into v_Limit From Zloptions Where ������ = 4;
  Exception
    When Others Then
      v_Limit := 10000;
  End;
  If v_Count > v_Limit Then
    Delete From Zlerrorlog
    Where Rowid In
          (Select Id From (Select Rowid As Id From Zlerrorlog Group By ʱ��, Rowid) Where Rownum < v_Count - v_Limit + 1);
  End If;
End Zl_Autologprocess;
/

--111538:����,2017-7-25,���RAC������getClient�е�v$Session��ΪGv$Session
CREATE OR REPLACE Package Zltools.b_Runmana Is
 
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

--111538:����,2017-7-25,���RAC������getClient�е�v$Session��ΪGv$Session
CREATE OR REPLACE Package Body Zltools.b_Runmana Is
 
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
               Select 0 As ��� 
               From Dual) M 
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
                             a.������, Decode(b.Terminal, Null, 0, 1) As ״̬, a.�ռ���־,a.����������,a.վ��,a.������ƵԴ
                From Zlclients a, (Select Distinct Terminal From GV$session) b
                Where Upper(a.����վ) = Upper(b.Terminal(+))
                Order By a.Ip'; 
      Open Cur_Out For v_Sql; 
    Else 
      Open Cur_Out For 
        Select Ip, ����վ, Cpu, �ڴ�, Ӳ��, ����ϵͳ, ����, ��;, ˵��, ������־, �ռ���־, ��ֹʹ��, ������, ����������, վ��, ������ƵԴ 
        From zlClients 
        Where Upper(����վ) = ����վ_In; 
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
        Select Nvl(To_Number(����ֵ), 0) 
        From zlOptions 
        Where ������ = 4; 
    End If; 
    If ��־����_In = '������־' Then 
      Open Cur_Out For 
        Select Count(*) ���� 
        From zlDiaryLog 
        Union All 
        Select Nvl(To_Number(����ֵ), 0) 
        From zlOptions 
        Where ������ = 2; 
 
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

--110944:������,2017-07-25,�������Ӧ�и�
Alter Table zlTools.zlRPTItems Add ����Ӧ�и� Number(1);

--104724:������,2017-08-08,������ͣ����
Alter Table zlTools.zlReports Add �Ƿ�ͣ�� Number(1);
Alter Table zlTools.zlRPTGroups Add �Ƿ�ͣ�� Number(1);

--113763:����һ,2017-08-31,�����������DBA������
Insert Into zlTools.zlSvrTools(���,�ϼ�,����,���,˵��,����) Values('06',Null,'DBA����','D',Null,Null);
Insert Into zlTools.zlSvrTools(���,�ϼ�,����,���,˵��,����) Values('0601','06','���ݿ�����','M',Null,1);
Insert Into zlTools.zlSvrTools(���,�ϼ�,����,���,˵��,����) Values('0602','06','SQL����','T',Null,2);
Insert Into zlTools.zlSvrTools(���,�ϼ�,����,���,˵��,����) Values('0603','06','SQL����','S',Null,3);
Insert Into zlTools.zlSvrTools(���,�ϼ�,����,���,˵��,����) Values('0604','06','�Ự����','B',Null,4);
Insert Into zlTools.zlSvrTools(���,�ϼ�,����,���,˵��,����) Values('0605','06','�������','F',Null,5);
Insert Into zlTools.zlSvrTools(���,�ϼ�,����,���,˵��,����) Values('0606','06','�ռ����','R',Null,6);