----10.35.30---��10.35.40
--000000:��˶,2017-02-20,����Ŀ��汾
alter table zlTools.zlUpgradeLog modify Ŀ��汾 VARCHAR2(20);

--104473:����ԭ,2017-02-15,�ͻ�����������,�汾�����ݽṹ����
Update zltools.zlFilesUpgrade Set �汾�� = Null;
Alter Table Zltools.Zlfilesupgrade Modify(�汾�� Varchar2(20));

--102830:��˶,2016-11-21,֧���ļ����ص����Ŀ¼
alter table zltools.Zlfilesupgrade  add ���Ӱ�װ·�� varchar2(500);

--102814:��˶,2016-11-21,ǿ����������
alter table zltools.zlclients  add �Ƿ��������� number(1);

--104473:����ԭ,2016-12-26,�ͻ�����������,ɾ��վ���ļ��ռ���վ�㲿�������������ļ�����˵���Ŀɾ��
delete from zltools.zlsvrtools where  ��� in('0311','0307','0309');

--104473:����ԭ,2016-12-26,�ͻ�����������,�����ͻ��˹����������߲˵���Ŀ
insert into zltools.zlsvrtools(���,�ϼ�,����,���,˵��,����) select '0307','03','�ͻ�����������','A','',22 from dual;

--104473:����ԭ,2016-12-26,�ͻ�����������,�ͻ��˱������ֶ�
alter table zltools.zlclients add  �Ƿ�Ԥ���� NUMBER(1) default 0;
alter table zltools.zlclients add  ����˵�� VARCHAR2(2000);
alter table zltools.zlclients add  �ռ�˵�� VARCHAR2(2000);
alter table zltools.zlclients add  �޸�˵�� VARCHAR2(2000);
alter table zltools.zlclients add  Ԥ����˵�� VARCHAR2(2000);
alter table zltools.zlclients add  �����ļ������� NUMBER(3);
alter table zltools.zlclients add  �޸�״̬ NUMBER(1);
alter table zltools.zlclients add  �ռ�״̬ NUMBER(1);
alter table zltools.zlclients add  ���� NUMBER(5);

--104473:����ԭ,2016-12-26,�ͻ�����������,����������װ·��
alter table zltools.zlfiles add  ���Ӱ�װ·�� varchar2(500);

--104473:����ԭ,2016-12-26,�ͻ�����������,�������������ñ�
Create Table ZLTOOLS.ZLUpgradeServer(
  ���     Number(3),
  ����     Number(1),
  λ��     Varchar2(100),
  �û���   Varchar2(20),
  ����     Varchar2(40),
  �˿�     Number(5),
  �Ƿ����� Number(1),
  �Ƿ�ȱʡ Number(1),
  �Ƿ��ռ� Number(1),
  �ռ����� Varchar2(100),
  ����        NUMBER(5))
PCTFREE 5;
ALTER TABLE Zltools.ZLUpgradeServer ADD CONSTRAINT ZLUpgradeServer_UQ_λ�� Unique (λ��) USING INDEX PCTFREE 5;
ALTER TABLE Zltools.ZLUpgradeServer ADD CONSTRAINT ZLUpgradeServer_PK_��� PRIMARY KEY (���) USING INDEX;

--104473:����ԭ,2016-12-26,�ͻ�����������,��������������
Create Or Replace Procedure Zltools.Zl_Zlupgradeserver_Insert
(
  ���_In     In Zlupgradeserver.���%Type,
  ����_In     In Zlupgradeserver.����%Type,
  λ��_In     In Zlupgradeserver.λ��%Type,
  �û���_In   In Zlupgradeserver.�û���%Type,
  ����_In     In Zlupgradeserver.����%Type,
  �˿�_In     In Zlupgradeserver.�˿�%Type,
  �Ƿ�����_In In Zlupgradeserver.�Ƿ�����%Type,
  �Ƿ�ȱʡ_In In Zlupgradeserver.�Ƿ�ȱʡ%Type,
  �Ƿ��ռ�_In In Zlupgradeserver.�Ƿ��ռ�%Type,
  �ռ�����_In In Zlupgradeserver.�ռ�����%Type
) Is
Begin
  --�ж�����ȱʡ�������Լ��ռ�ȱʡ������
  If �Ƿ�ȱʡ_In = 1 Then
    Update Zlupgradeserver Set �Ƿ�ȱʡ = 0 Where Nvl(�Ƿ�ȱʡ, 0) = 1;
  End If;
  If �Ƿ��ռ�_In = 1 Then
    Update Zlupgradeserver Set �Ƿ��ռ� = 0 Where Nvl(�Ƿ��ռ�, 0) = 1;
  End If;
  --�����¼ 
  Insert Into Zlupgradeserver
    (���, ����, λ��, �û���, ����, �˿�, �Ƿ�����, �Ƿ�ȱʡ, �Ƿ��ռ�, �ռ�����, ����)
  Values
    (���_In, ����_In, λ��_In, �û���_In, ����_In, �˿�_In, �Ƿ�����_In, �Ƿ�ȱʡ_In, �Ƿ��ռ�_In, �ռ�����_In, 0);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
  
End Zl_Zlupgradeserver_Insert;
/


--104473:����ԭ,2016-12-26,�ͻ�����������,ɾ������������
Create Or Replace Procedure Zltools.Zl_Zlupgradeserver_Delete(���_In In Zlupgradeserver.���%Type) Is
Begin
  --ɾ������
  Delete From Zlupgradeserver Where ��� = ���_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
  
End Zl_Zlupgradeserver_Delete;
/


--104473:����ԭ,2016-12-26,�ͻ�����������,�޸ķ���������
Create Or Replace Procedure Zltools.Zl_Zlupgradeserver_Update
(
  ���_In     In Zlupgradeserver.���%Type,
  ����_In     In Zlupgradeserver.����%Type,
  λ��_In     In Zlupgradeserver.λ��%Type,
  �û���_In   In Zlupgradeserver.�û���%Type,
  ����_In     In Zlupgradeserver.����%Type,
  �˿�_In     In Zlupgradeserver.�˿�%Type,
  �Ƿ�����_In In Zlupgradeserver.�Ƿ�����%Type,
  �Ƿ�ȱʡ_In In Zlupgradeserver.�Ƿ�ȱʡ%Type,
  �Ƿ��ռ�_In In Zlupgradeserver.�Ƿ��ռ�%Type,
  �ռ�����_In In Zlupgradeserver.�ռ�����%Type,
  Intedittype Pls_Integer
) Is
  --���� Zlupgradeserver.�Ƿ�����%Type;
  --ȱʡ Zlupgradeserver.�Ƿ�ȱʡ%Type;
  --�ռ� Zlupgradeserver.�Ƿ��ռ�%Type;
Begin
  If �Ƿ�ȱʡ_In = 1 Then
    Update Zlupgradeserver Set �Ƿ�ȱʡ = 0 Where Nvl(�Ƿ�ȱʡ, 0) = 1;
  End If;
  If �Ƿ��ռ�_In = 1 Then
    Update Zlupgradeserver Set �Ƿ��ռ� = 0 Where Nvl(�Ƿ��ռ�, 0) = 1;
  End If;
  --�޸�����Ϊ0 �����޸������ֶ�����
  If Intedittype = 0 Then
    Update Zlupgradeserver
    Set ���� = ����_In, λ�� = λ��_In, �û��� = �û���_In, ���� = ����_In, �˿� = �˿�_In, �Ƿ����� = �Ƿ�����_In, �Ƿ�ȱʡ = �Ƿ�ȱʡ_In, �Ƿ��ռ� = �Ƿ��ռ�_In,
        �ռ����� = �ռ�����_In
    Where ��� = ���_In;
  End If;

  --�޸�����Ϊ1 �޸ķ�����ȱʡ����
  If Intedittype = 1 Then
    Update Zlupgradeserver
    Set �Ƿ����� = �Ƿ�����_In, �Ƿ�ȱʡ = �Ƿ�ȱʡ_In, �Ƿ��ռ� = �Ƿ��ռ�_In, �ռ����� = �ռ�����_In
    Where ��� = ���_In;
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlupgradeserver_Update;
/

--104473:����ԭ,2017-1-16,�ͻ�����������,�����ļ��嵥�޸�
Create Or Replace Procedure Zltools.Zlfilesupgrade_Repair Is
Begin
  Delete From zlFilesUpgrade C
  Where c.�ļ����� <> 4 Or
        c.�ļ��� In
        (Select ���� From zlFilesUpgrade A, Zlfiles B Where a.�ļ��� = b.���� And (a.�ļ����� = 4 Or b.�ļ����� = 4));

  Insert Into zlFilesUpgrade
    (�ļ���, Md5, �汾��, �޸�����, ��������, �ļ�����, ��װ·��, ҵ�񲿼�, ����ϵͳ, �ļ�˵��, �Զ�ע��, ǿ�Ƹ���)
    Select ����, ��׼md5, �汾��, �޸�����, ��������, �ļ�����, ��װ·��, ҵ�񲿼�, ����ϵͳ, �ļ�˵��, �Զ�ע��, ǿ�Ƹ���
    From Zlfiles B
    Where Not Exists (Select 1 From zlFilesUpgrade R Where r.�ļ��� = b.����);

  Update zlFilesUpgrade Set Md5 = Null;

End Zlfilesupgrade_Repair;
/

--104473:����ԭ,2016-12-26,�ͻ�����������,Ĭ�Ϸ���������
Create Or Replace Procedure Zltools.Zlreginfo_Defaultserver
(
  ����_In   In Zlreginfo.����%Type,
  λ��_In   In Zlreginfo.����%Type,
  �û���_In In Zlreginfo.����%Type,
  ����_In   In Zlreginfo.����%Type,
  �˿�_In   In Zlreginfo.����%Type
) Is
Begin

  Insert Into Zltools.Zlreginfo
    (��Ŀ, ����)
    Select '��������', '' From Dual Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = '��������');
  Insert Into Zltools.Zlreginfo
    (��Ŀ, ����)
    Select 'FTP������0', '' From Dual Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = 'FTP������0');
  Insert Into Zltools.Zlreginfo
    (��Ŀ, ����)
    Select 'FTP�û�0', '' From Dual Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = 'FTP�û�0');
  Insert Into Zltools.Zlreginfo
    (��Ŀ, ����)
    Select 'FTP����0', '' From Dual Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = 'FTP����0');
  Insert Into Zltools.Zlreginfo
    (��Ŀ, ����)
    Select 'FTP�˿�0', '' From Dual Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = 'FTP�˿�0');
  Insert Into Zltools.Zlreginfo
    (��Ŀ, ����)
    Select '������Ŀ¼0', '' From Dual Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = '������Ŀ¼0');
  Insert Into Zltools.Zlreginfo
    (��Ŀ, ����)
    Select '�����û�0', '' From Dual Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = '�����û�0');
  Insert Into Zltools.Zlreginfo
    (��Ŀ, ����)
    Select '��������0', '' From Dual Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = '��������0');

  If ����_In = 0 Then
    Update Zltools.Zlreginfo Set ���� = '0' Where ��Ŀ = '��������';
    Update Zltools.Zlreginfo Set ���� = λ��_In Where ��Ŀ = '������Ŀ¼0';
    Update Zltools.Zlreginfo Set ���� = �û���_In Where ��Ŀ = '�����û�0';
    Update Zltools.Zlreginfo Set ���� = ����_In Where ��Ŀ = '��������0';
    Update Zltools.Zlclients Set Ftp������ = '0';
  Else
    Update Zltools.Zlreginfo Set ���� = '1' Where ��Ŀ = '��������';
    Update Zltools.Zlreginfo Set ���� = λ��_In Where ��Ŀ = 'FTP������0';
    Update Zltools.Zlreginfo Set ���� = �û���_In Where ��Ŀ = 'FTP�û�0';
    Update Zltools.Zlreginfo Set ���� = ����_In Where ��Ŀ = 'FTP����0';
    Update Zltools.Zlreginfo Set ���� = �˿�_In Where ��Ŀ = 'FTP�˿�0';
    Update Zltools.Zlclients Set ���������� = '0';
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zlreginfo_Defaultserver;
/


--104473:����ԭ,2016-12-26,�ͻ�����������,�ͻ�������Ԥ������־����
Create Or Replace Procedure Zltools.Zl_Zlclients_Update
(
  ����վ_In   In Varchar2,
  �����ֶ�_In In Number,
  ����ֵ_In   In Number
) Is
  --����վ_In ����վ�ַ������Զ���������
  --�����ֶ�_In 0��������־ 1���Ƿ�Ԥ���� 2���ռ���־
  --����ֵ_In 0������ѡ 1����ѡ
  v_����վ Varchar2(4000);
  n_����ֵ Number;
Begin
  v_����վ := ����վ_In;
  n_����ֵ := ����ֵ_In;

  If �����ֶ�_In = 0 Then
    Update zlClients Set ������־ = n_����ֵ Where ����վ In (Select Column_Value From Table(f_Str2list(v_����վ)));
    If n_����ֵ = 1 Then
      Update zlClients Set ������� = 0 Where ����վ In (Select Column_Value From Table(f_Str2list(v_����վ)));
      Update zlClients Set �޸�״̬ = 0 Where ����վ In (Select Column_Value From Table(f_Str2list(v_����վ)));
      Update zlClients Set ����˵�� = Null Where ����վ In (Select Column_Value From Table(f_Str2list(v_����վ)));
      Update zlClients Set �޸�˵�� = Null Where ����վ In (Select Column_Value From Table(f_Str2list(v_����վ)));
    End If;
  End If;

  If �����ֶ�_In = 1 Then
    Update zlClients Set �Ƿ�Ԥ���� = n_����ֵ Where ����վ In (Select Column_Value From Table(f_Str2list(v_����վ)));
    If n_����ֵ = 1 Then
      Update zlClients Set Ԥ����� = 0 Where ����վ In (Select Column_Value From Table(f_Str2list(v_����վ)));
      Update zlClients Set Ԥ����˵�� = Null Where ����վ In (Select Column_Value From Table(f_Str2list(v_����վ)));
    End If;
  End If;

  If �����ֶ�_In = 2 Then
    Update zlClients Set �ռ���־ = n_����ֵ Where ����վ In (Select Column_Value From Table(f_Str2list(v_����վ)));
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlclients_Update;
/

CREATE OR REPLACE Procedure ZLTOOLS.Zlfiles_Autoupdate
(
  ����_In         In Zlfiles.����%Type,
  ��׼md5_In      In Zlfiles.��׼md5%Type,
  �汾��_In       In Zlfiles.�汾��%Type,
  �޸�����_In     In Zlfiles.�޸�����%Type,
  ��������_In     In Zlfiles.��������%Type,
  �ļ�����_In     In Zlfiles.�ļ�����%Type,
  ��װ·��_In     In Zlfiles.��װ·��%Type,
  ҵ�񲿼�_In     In Zlfiles.ҵ�񲿼�%Type,
  ����ϵͳ_In     In Zlfiles.����ϵͳ%Type,
  �ļ�˵��_In     In Zlfiles.�ļ�˵��%Type,
  �Զ�ע��_In     In Zlfiles.�Զ�ע��%Type,
  ǿ�Ƹ���_In     In Zlfiles.ǿ�Ƹ���%Type,
  ���Ӱ�װ·��_In In Zlfiles.���Ӱ�װ·��%Type
) Is
  n_Count Number(3);
Begin
  n_Count := 0;
  --��������
  For Rs In (Select Rowid From Zlfiles A Where Upper(a.����) = Upper(����_In)) Loop
    n_Count := n_Count + 1;
    Update Zlfiles
    Set ���� = ����_In, ��׼md5 = ��׼md5_In, �汾�� = �汾��_In, �޸����� = �޸�����_In, �������� = ��������_In, �ļ����� = �ļ�����_In, ��װ·�� = ��װ·��_In,
        ҵ�񲿼� = ҵ�񲿼�_In, ����ϵͳ = ����ϵͳ_In, �ļ�˵�� = �ļ�˵��_In, �Զ�ע�� = �Զ�ע��_In, ǿ�Ƹ��� = ǿ�Ƹ���_In, ���Ӱ�װ·�� = ���Ӱ�װ·��_In
    Where Rowid = Rs.Rowid;
  End Loop;
  --��������
  If n_Count = 0 Then
    Insert Into Zlfiles
      (����, ��׼md5, �汾��, �޸�����, ��������, �ļ�����, ��װ·��, ҵ�񲿼�, ����ϵͳ, �ļ�˵��, �Զ�ע��, ǿ�Ƹ���, ���Ӱ�װ·��)
    Values
      (����_In, ��׼md5_In, �汾��_In, �޸�����_In, ��������_In, �ļ�����_In, ��װ·��_In, ҵ�񲿼�_In, ����ϵͳ_In, �ļ�˵��_In, �Զ�ע��_In, ǿ�Ƹ���_In, ���Ӱ�װ·��_In);
  End If;
  n_Count := 0;
  --��������
  For Rs In (Select Rowid From zlFilesUpgrade A Where Upper(a.�ļ���) = Upper(����_In)) Loop
    n_Count := n_Count + 1;
    Update zlFilesUpgrade
    Set �ļ��� = ����_In, �ļ����� = �ļ�����_In, ��װ·�� = ��װ·��_In, ҵ�񲿼� = ҵ�񲿼�_In, ����ϵͳ = ����ϵͳ_In, �ļ�˵�� = �ļ�˵��_In, �Զ�ע�� = �Զ�ע��_In,
        ǿ�Ƹ��� = ǿ�Ƹ���_In, ���Ӱ�װ·�� = ���Ӱ�װ·��_In
    Where Rowid = Rs.Rowid;
  End Loop;
  --��������
  If n_Count = 0 Then
    Insert Into zlFilesUpgrade
      (�ļ���, �ļ�����, ��װ·��, ҵ�񲿼�, ����ϵͳ, �ļ�˵��, �Զ�ע��, ǿ�Ƹ���, ���Ӱ�װ·��)
    Values
      (����_In, �ļ�����_In, ��װ·��_In, ҵ�񲿼�_In, ����ϵͳ_In, �ļ�˵��_In, �Զ�ע��_In, ǿ�Ƹ���_In, ���Ӱ�װ·��_In);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zlfiles_Autoupdate;
/

--105087:��˶,2017-01-16,֧�ֲ��Ų������͵ı䶯
Create Or Replace Procedure Zltools.Zl_Parameters_Change
(
  ����id_In   Zlparameters.Id%Type,
  ˽��_In     Zlparameters.˽��%Type,
  ����_In     Zlparameters.����%Type,
  ��Ȩ_In     Zlparameters.��Ȩ%Type,
  �䶯��_In   Zlparachangedlog.�䶯��%Type,
  �䶯ԭ��_In Zlparachangedlog.�䶯ԭ��%Type,
  ����_In     Zlparameters.����%Type := 0
) Is
  v_Temp     Varchar2(200);
  n_ģ��     Zlparameters.ģ��%Type;
  n_˽��     Zlparameters.˽��%Type;
  n_����     Zlparameters.����%Type;
  n_����     Zlparameters.����%Type;
  n_��Ȩ     Zlparameters.��Ȩ%Type;
  n_���     Zlparachangedlog.���%Type;
  v_�䶯˵�� Zlparachangedlog.�䶯˵��%Type;
  v_�䶯���� Zlparachangedlog.�䶯����%Type;

  Function Gettype
  (
    ģ��_In Zlparameters.˽��%Type,
    ˽��_In Zlparameters.˽��%Type,
    ����_In Zlparameters.����%Type,
    ����_In Zlparameters.����%Type
  ) Return Varchar2 Is
  Begin
    If Nvl(����_In, 0) = 1 Then
      Return '���Ų���';
    End If;
    If Nvl(ģ��_In, 0) = 0 Then
      --����ģ��,֤��ֻ����������:����ȫ�ֺ�˽��ȫ�� 
      If Nvl(˽��_In, 0) = 0 Then
        Return '����ȫ��';
      End If;
      Return '˽��ȫ��';
    End If;
  
    --��ģ��Ĵ��� 
    If ����_In = 0 Then
      --���Ǳ��������,ֻ����������:����ģ���˽��ģ�� 
      If Nvl(˽��_In, 0) = 0 Then
        Return '����ģ��';
      End If;
      Return '˽��ģ��';
    End If;
    --�Ա�����ģ����д���Ҳ���������: 
    If Nvl(˽��_In, 0) = 0 Then
      Return '��������ģ��';
    End If;
    Return '����˽��ģ��';
  Exception
    When Others Then
      Return Null;
  End Gettype;
Begin

  Select Nvl(ģ��, 0), Nvl(˽��, 0), Nvl(����, 0), Nvl(��Ȩ, 0), Nvl(����, 0)
  Into n_ģ��, n_˽��, n_����, n_��Ȩ, n_����
  From Zlparameters
  Where Id = ����id_In;
  Select Nvl(Max(���), 0) + 1 Into n_��� From Zlparachangedlog Where ����id = ����id_In;
  --�������� 
  --˵���䶯˵��:����:˽��ģ���Ϊ����ģ�顣 
  -- �䶯����:˵���䶯�ֶεı仯���:����:˽��:1-->0,����:1-->0 
  v_�䶯˵�� := Null;
  v_�䶯���� := Null;
  --���ͷ����˸ı� 
  If n_˽�� <> Nvl(˽��_In, 0) Or n_���� <> Nvl(����_In, 0) Or n_���� <> Nvl(����_In, 0) Then
    v_Temp     := '��' || Gettype(n_ģ��, n_˽��, n_����, n_����);
    v_Temp     := v_Temp || '��Ϊ' || Gettype(n_ģ��, Nvl(˽��_In, 0), Nvl(����_In, 0), Nvl(����_In, 0));
    v_�䶯˵�� := v_Temp;
    v_Temp     := '';
    If n_���� <> Nvl(����_In, 0) Then
      v_Temp := v_Temp || ',����:' || n_���� || '-->' || Nvl(����_In, 0);
    End If;
    If n_˽�� <> Nvl(˽��_In, 0) Then
      v_Temp := v_Temp || ',˽��:' || n_˽�� || '-->' || Nvl(˽��_In, 0);
    End If;
    If n_˽�� <> Nvl(˽��_In, 0) Then
      v_Temp := v_Temp || ',����:' || n_���� || '-->' || Nvl(����_In, 0);
    End If;
    v_�䶯���� := Substr(v_Temp, 2);
  End If;
  --�����Ȩ�����ı�û�� 
  If n_��Ȩ <> Nvl(��Ȩ_In, 0) Then
    If Not v_�䶯˵�� Is Null Then
      v_�䶯˵�� := v_�䶯˵�� || ',';
    End If;
    If n_��Ȩ = 0 Then
      v_Temp := '����Ҫ��Ȩ';
    Else
      v_Temp := '��Ҫ��Ȩ';
    End If;
    v_�䶯˵�� := Nvl(v_�䶯˵��, '') || '��' || v_Temp || '��Ϊ';
    If ��Ȩ_In = 0 Then
      v_Temp := '����Ҫ��Ȩ';
    Else
      v_Temp := '��Ҫ��Ȩ';
    End If;
    v_�䶯˵�� := Nvl(v_�䶯˵��, '') || v_Temp;
  
    If Not v_�䶯���� Is Null Then
      v_�䶯���� := v_�䶯���� || ',';
    End If;
    v_�䶯���� := Nvl(v_�䶯����, '') || '��Ȩ:' || n_��Ȩ || '-->' || Nvl(��Ȩ_In, 0);
  End If;

  Insert Into Zlparachangedlog
    (����id, ���, �䶯˵��, �䶯����, �䶯��, �䶯ʱ��, �䶯ԭ��)
  Values
    (����id_In, n_���, v_�䶯˵��, v_�䶯����, �䶯��_In, Sysdate, �䶯ԭ��_In);

  Update Zlparameters
  Set ˽�� = Nvl(˽��_In, 0), ���� = Nvl(����_In, 0), ��Ȩ = Nvl(��Ȩ_In, 0), ���� = Nvl(����_In, 0)
  Where Id = ����id_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Parameters_Change;
/

--104473:����ԭ,2016-1-18,�ͻ�����������,Ԥ����ʱ�����ô洢���̣������������Ӹ��ǲ����������ǻ򲻸���
Create Or Replace Procedure Zltools.Zl_Zlclients_Setpretime
(
  n_Mode_In  Number,
  n_Cover_In Number := 1,
  v_����վ_In   Zlclients.����վ%Type := Null,
  d_Ԥ��ʱ��_In Zlclients.Ԥ��ʱ��%Type := Null
) Is
  v_Timeset Varchar2(300);
  v_Err     Varchar2(500);
  Err_Custom Exception;
  --n_Cover_In 0-������ 1-���� 
Begin
  --0-�����Կͻ���Ԥ����ʱ������
  If n_Mode_In = 0 Then
    If v_����վ_In Is Not Null Then
      Update zlClients Set Ԥ��ʱ�� = d_Ԥ��ʱ��_In Where ����վ = v_����վ_In;
    End If;
  --1-Ԥ����ʱ���Զ�����n_Cover_In�����������
  Elsif n_Mode_In = 1 Then
    Select Max(����) Into v_Timeset From zlRegInfo Where ��Ŀ = '�ͻ���Ԥ����ʱ���';
    If v_Timeset Is Not Null Then
      For r_Ip In (Select To_Date(Today || ' ' || Date_d, 'yyyy-mm-dd HH24:mi:ss') Ԥ��ʱ��, ����վ, Ip
                   From (Select ����վ, Ip, Rownum Rn_c From zlClients) A,
                        (Select To_Char(Sysdate, 'yyyy-mm-dd') Today, Column_Value Date_d, Rownum Rn_d, Count(1) Over() Sn
                          From Table(f_Str2list(v_Timeset, ','))) B
                   Where Mod(a.Rn_c, Sn) + 1 = Rn_d) Loop
        If n_Cover_In = 1 Then
          Update zlClients Set Ԥ��ʱ�� = r_Ip.Ԥ��ʱ�� Where ����վ = r_Ip.����վ And Ip = r_Ip.Ip;
        Elsif n_Cover_In = 0 Then
          Update zlClients
          Set Ԥ��ʱ�� = r_Ip.Ԥ��ʱ��
          Where ����վ = r_Ip.����վ And Ip = r_Ip.Ip And Ԥ��ʱ�� Is Null;
        End If;
      End Loop;
    Else
      Update zlClients Set Ԥ��ʱ�� = Null;
    End If;
  Elsif n_Mode_In = 3 Then
    Update zlClients Set Ԥ����� = 0;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlclients_Setpretime;
/
--00000:��˶,2017-03-03,����޸Ĺ��̲���
CREATE OR REPLACE Procedure ZLTOOLS.Zl_Zlclients_Updateprocess
(
  v_����վ_In  Zlclients.����վ%Type := Null,
  n_Operate_In Number,
  n_State_In   Zlclients.�������%Type := 0,
  v_˵��_In    Zlclients.˵��%Type := Null,
  n_����_In    Zlclients.����%Type := Null
  --���ܣ��ͻ��˲���״̬����
  --Ӧ�ã�N_Operate_In=0-�����޸�����ʱ�൱����ʽ����������Ԥ������Ϣ��������־��д�������޸���Ϣ
  --                  =1-Ԥ���������Ԥ������־��д��Ԥ������Ϣ
  --                  =2-��ʽ����������Ԥ������Ϣ��������־��д��������Ϣ��
  --                  =3-�ռ�������������־��д���ռ���Ϣ��
  --      n_State_In=0-���������1-�����ɹ���2-����ʧ�ܡ�3-����ִ����
) Is
  v_Err Varchar2(500);
  Err_Custom Exception;
Begin
  If n_Operate_In = 0 Then
    If n_State_In = 1 Then
      Update Zltools.Zlclients
      Set ���� = n_����_In, �޸�״̬ = n_State_In, �޸�˵�� = v_˵��_In, ������־ = 0, ������� = 0, �Ƿ�Ԥ���� = 0, Ԥ����� = 0, Ԥ����˵�� = Null,
          ����˵�� = Null, �Ƿ��������� = 0
      Where ����վ = v_����վ_In;
    Else
      Update Zltools.Zlclients Set �޸�״̬ = n_State_In, �޸�˵�� = v_˵��_In Where ����վ = v_����վ_In;
    End If;
  Elsif n_Operate_In = 1 Then
    If n_State_In = 1 Then
      Update Zltools.Zlclients
      Set Ԥ����� = n_State_In, Ԥ����˵�� = v_˵��_In, �Ƿ�Ԥ���� = 0
      Where ����վ = v_����վ_In;
    Else
      Update Zltools.Zlclients Set Ԥ����� = n_State_In, Ԥ����˵�� = v_˵��_In Where ����վ = v_����վ_In;
    End If;
  Elsif n_Operate_In = 2 Then
    If n_State_In = 1 Then
      Update Zltools.Zlclients
      Set ���� = n_����_In, ������� = n_State_In, ����˵�� = v_˵��_In, ������־ = 0, �Ƿ�Ԥ���� = 0, Ԥ����� = 0, �޸�״̬ = 0, Ԥ����˵�� = Null,
          �޸�˵�� = Null, �Ƿ��������� = 0
      Where ����վ = v_����վ_In;
    Else
      Update Zltools.Zlclients Set ������� = n_State_In, ����˵�� = v_˵��_In Where ����վ = v_����վ_In;
    End If;
  Elsif n_Operate_In = 3 Then
    If n_State_In = 1 Then
      Update Zltools.Zlclients
      Set �ռ�״̬ = n_State_In, �ռ�˵�� = v_˵��_In, �ռ���־ = 0
      Where ����վ = v_����վ_In;
    Else
      Update Zltools.Zlclients Set �ռ�״̬ = n_State_In, �ռ�˵�� = v_˵��_In Where ����վ = v_����վ_In;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zlclients_Updateprocess;
/
