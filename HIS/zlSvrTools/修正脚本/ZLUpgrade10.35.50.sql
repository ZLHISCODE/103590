----10.35.40---��10.35.50
--107086:��˶,2017-03-15,�汾�Ų�����
update Zltools.Zlfilesupgrade set �汾��=Null;
--108005:��˶,2017-04-06,����ʽ�汾�Ų����ݵ����޷�����
Update Zltools.Zlfilesupgrade
Set �汾�� = '1000350040'
Where Upper(�ļ���) In ('ZLHISCRUST.EXE', '7Z.EXE;7Z.DLL', 'AAMD532.DLL', 'ZLRUNAS.EXE', 'REGCOM.DLL', 'GACUTIL.EXE','GACUTIL.EXE.CONFIG');
--107086:��˶,2017-03-15,�汾�Ų�����
Alter Table Zltools.Zlfilesupgrade add �ļ��汾�� Varchar2(20);
Create Or Replace Procedure Zltools.Zlfilesupgrade_Repair Is
Begin
  Delete From Zlfilesupgrade a Where Exists (Select 1 From Zlfiles b Where Upper(b.����) = Upper(a.�ļ���));
  Insert Into Zlfilesupgrade
    (�ļ���, �ļ�����, ��װ·��, ���Ӱ�װ·��, ҵ�񲿼�, ����ϵͳ, �ļ�˵��, �Զ�ע��, ǿ�Ƹ���)
    Select ����, �ļ�����, ��װ·��, ���Ӱ�װ·��, ҵ�񲿼�, ����ϵͳ, �ļ�˵��, �Զ�ע��, ǿ�Ƹ��� From Zlfiles;
  Update Zlfilesupgrade Set Md5 = Null;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zlfilesupgrade_Repair;
/
--107146:��˶,2017-03-16,���ֿͻ����޷�����
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
  If ����_In = 0 Then
    Insert Into Zltools.Zlreginfo
      (��Ŀ, ����)
      Select '������Ŀ¼0', ''
      From Dual
      Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = '������Ŀ¼0')
      Union All
      Select '�����û�0', ''
      From Dual
      Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = '�����û�0')
      Union All
      Select '��������0', ''
      From Dual
      Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = '��������0')
      Union All
      Select '������Ŀ¼', ''
      From Dual
      Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = '������Ŀ¼')
      Union All
      Select '�����û�', ''
      From Dual
      Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = '�����û�')
      Union All
      Select '��������', ''
      From Dual
      Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = '��������');
    Update Zltools.Zlreginfo Set ���� = '0' Where ��Ŀ = '��������';
    Update Zltools.Zlreginfo Set ���� = λ��_In Where ��Ŀ = '������Ŀ¼0';
    Update Zltools.Zlreginfo Set ���� = �û���_In Where ��Ŀ = '�����û�0';
    Update Zltools.Zlreginfo Set ���� = ����_In Where ��Ŀ = '��������0';
    Update Zltools.Zlreginfo Set ���� = λ��_In Where ��Ŀ = '������Ŀ¼';
    Update Zltools.Zlreginfo Set ���� = �û���_In Where ��Ŀ = '�����û�';
    Update Zltools.Zlreginfo Set ���� = ����_In Where ��Ŀ = '��������';
    Update Zltools.Zlclients Set ���������� = 0;
  Else
    Insert Into Zltools.Zlreginfo
      (��Ŀ, ����)
      Select 'FTP������0', ''
      From Dual
      Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = 'FTP������0')
      Union All
      Select 'FTP�û�0', ''
      From Dual
      Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = 'FTP�û�0')
      Union All
      Select 'FTP����0', ''
      From Dual
      Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = 'FTP����0')
      Union All
      Select 'FTP�˿�0', ''
      From Dual
      Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = 'FTP�˿�0')
      Union All
      Select 'FTP������', ''
      From Dual
      Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = 'FTP������')
      Union All
      Select 'FTP�û�', ''
      From Dual
      Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = 'FTP�û�')
      Union All
      Select 'FTP����', ''
      From Dual
      Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = 'FTP����')
      Union All
      Select 'FTP�˿�', ''
      From Dual
      Where Not Exists (Select 1 From Zltools.Zlreginfo Where ��Ŀ = 'FTP�˿�');
    Update Zltools.Zlreginfo Set ���� = '1' Where ��Ŀ = '��������';
    Update Zltools.Zlreginfo Set ���� = λ��_In Where ��Ŀ = 'FTP������0';
    Update Zltools.Zlreginfo Set ���� = �û���_In Where ��Ŀ = 'FTP�û�0';
    Update Zltools.Zlreginfo Set ���� = ����_In Where ��Ŀ = 'FTP����0';
    Update Zltools.Zlreginfo Set ���� = �˿�_In Where ��Ŀ = 'FTP�˿�0';
    Update Zltools.Zlreginfo Set ���� = λ��_In Where ��Ŀ = 'FTP������';
    Update Zltools.Zlreginfo Set ���� = �û���_In Where ��Ŀ = 'FTP�û�';
    Update Zltools.Zlreginfo Set ���� = ����_In Where ��Ŀ = 'FTP����';
    Update Zltools.Zlreginfo Set ���� = �˿�_In Where ��Ŀ = 'FTP�˿�';
    Update Zltools.Zlclients Set Ftp������ = 0;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zlreginfo_Defaultserver;
/

--108032:����,2017-04-07,���"��������"����
Insert Into Zltools.Zlsvrtools (���, �ϼ�, ����, ���, ˵��, ����) Values ('0207', '02', '��������', 'S', Null, 19);

--108032:����,2017-04-07,���"��������"����
Create Table Zltools.zlConnections(
    ��� number(4),
    ���� varchar2(20),
    �û��� varchar2(30),
    ���� varchar2(30),
    IP varchar2(30),
    �˿� number(5),
    ʵ���� varchar2(50),
    ˵�� varchar2(500));
    
Alter Table Zltools.zlConnections Add Constraint zlConnections_PK Primary Key (���) Using Index;
Alter Table Zltools.zlConnections Add Constraint zlConnections_UQ_���� Unique (����) Using Index;

--107979:������,2017-04-17,�Զ��屨����֧�ֶ���������
Alter Table Zltools.zlRPTDatas Add �������ӱ�� Number(4);
Alter Table Zltools.zlRPTDatas Add Constraint ZLRPTDATAS_UQ_�������ӱ�� Unique(�������ӱ��, ID) Using Index;
Alter Table Zltools.zlRPTDatas Add Constraint ZLRPTDATAS_FK_�������ӱ�� Foreign Key(�������ӱ��) References Zltools.ZlConnections(���) Enable Novalidate;

--108032:����,2017-04-07,���"��������"����
Create Or Replace Procedure Zltools.Zl_Zlconnections_Edit
(
  ����_In   Number, --0-����,1-�޸�,2-ɾ��
  ���_In   Zlconnections.���%Type,
  ����_In   Zlconnections.����%Type := Null,
  �û���_In Zlconnections.�û���%Type := Null,
  ����_In   Zlconnections.����%Type := Null,
  Ip_In     Zlconnections.Ip%Type := Null,
  �˿�_In   Zlconnections.�˿�%Type := Null,
  ʵ����_In Zlconnections.ʵ����%Type := Null,
  ˵��_In   Zlconnections.˵��%Type := Null
) Is
  n_���    Zlconnections.���%Type;
  n_Count   Number(1);
  v_Err_Msg Varchar2(255);
  Err_Item Exception;
Begin
  If ����_In = 0 Then
    Select Count(1) Into n_Count From Zlconnections Where ���� = ����_In;
    If n_Count = 1 Then
      v_Err_Msg := '�����������Ѵ��ڣ�';
      Raise Err_Item;
    End If;
    Select Nvl(Max(���), 0) Into n_��� From Zlconnections;
    Insert Into Zlconnections
      (���, ����, �û���, ����, Ip, �˿�, ʵ����, ˵��)
    Values
      (n_��� + 1, ����_In, �û���_In, ����_In, Ip_In, �˿�_In, ʵ����_In, ˵��_In);
  Elsif ����_In = 1 Then
    Update Zlconnections
    Set �û��� = �û���_In, ���� = ����_In, ���� = ����_In, Ip = Ip_In, �˿� = �˿�_In, ʵ���� = ʵ����_In, ˵�� = ˵��_In
    Where ��� = ���_In;
  Else
    Delete Zlconnections Where ��� = ���_In;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zlconnections_Edit;
/