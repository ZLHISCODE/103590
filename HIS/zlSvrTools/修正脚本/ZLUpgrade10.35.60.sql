----10.35.50---��10.35.60
--111205:��˶,2017-07-06,�����嵥��������
alter table ZLTOOLS.zlFilesUpgrade modify ҵ�񲿼� varchar2(500);
--110050:��˶,2017-6-19,�Ż��Ự��֤
Alter Table ZLTOOLS.zlRegFunc  Drop Constraint zlRegFunc_UQ cascade drop index;
Alter Table ZLTOOLS.zlRegFunc  Add Constraint zlRegFunc_PK PRIMARY KEY(ϵͳ, ���, ����) USING INDEX PCTFREE 5;
--97672:����,2017-4-11,�ϵ���Ϣ����(zlapptools),�޸�"�ʼ���Ϣ�������"
Update Zltools.Zlparameters Set ȱʡֵ = '60' Where Nvl(ϵͳ, 0) = 0 And Nvl(ģ��, 0) = 0 And ������ = '�ʼ���Ϣ�������';

--00000:������,2017-07-04,���Ӳ������Ʒ����ַ�����󳤶ȣ����ⳬ���ֶ����޶�����
Create Or Replace Function Zltools.f_List2str
(
  p_Strlist   In t_Strlist,
  p_Delimiter In Varchar2 Default ',',
  p_Distinct  In Number Default 1,
  p_Maxlength In Number Default 4000
) Return Varchar2 Is
  l_String Long;
  l_Add    Number;
  --���ܣ���һ���б���ת��Ϊһ��ȱʡ�Զ��ŷָ����ַ�����
  --����
  --Select ����, f_List2str(Cast(Collect(��Ա Order By ���) As t_Strlist)) ��Ա�б�
  --From (Select a.���� As ����, c.���� As ��Ա,c.���
  --      From ���ű� A, ������Ա B, ��Ա�� C
  --      Where a.Id = b.����id And b.��Աid = c.Id
  --      Order By ����, ��Ա)
  --Group By ����

  --�˺�����֧��with��ʽ�������ʱ�ڴ���⽫�ᱨ��ORA-00932: �������Ͳ�һ��: ӦΪ -, ��ȴ��� -��
  --���磺With Test As (Select '�ڿ�' As ����,'����' As ��Ա From Dual Union All......)
  --     Select ����,f_List2str(cast(COLLECT(��Ա) as t_Strlist)) tt From Test Group By ����
Begin
  If p_Strlist.Count > 0 Then
    For I In p_Strlist.First .. p_Strlist.Last Loop
      l_Add := 0;
      If p_Distinct = 1 Then
        If Instr(',' || l_String || ',', ',' || p_Strlist(I) || ',') = 0 Then
          l_Add := 1;
        End If;
      Else
        l_Add := 1;
      End If;
      If l_Add = 1 Then
        If I != p_Strlist.First Then
          l_String := l_String || p_Delimiter;
        End If;
        l_String := l_String || p_Strlist(I);
        If Lengthb(l_String) > p_Maxlength Then
          l_String := Substr(l_String, 1, p_Maxlength);
          Return l_String;
        End If;
      End If;
    End Loop;
  End If;
  Return l_String;
End f_List2str;
/

--97672:����,2017-4-11,�ϵ���Ϣ����(zlapptools),��SQL����޸�Ϊ����
CREATE OR REPLACE Procedure zltools.Zl_Zlmsgstate_Edit
(
  ����_In     Number, --0-����,1-�޸�,2-ɾ��
  ��Ϣid_In   Zlmsgstate.��Ϣid%Type,
  ����_In     Zlmsgstate.����%Type := Null,
  �û�_In     Zlmsgstate.�û�%Type := Null,
  ���_In     Zlmsgstate.���%Type := Null,
  ɾ��_In     Zlmsgstate.ɾ��%Type := Null,
  ״̬_In     Zlmsgstate.״̬%Type := Null,
  ��������_In Number := Null
) Is
  n_���� Number;
  n_���� Number;
Begin
  If ����_In = 0 Then
    Insert Into Zlmsgstate
      (��Ϣid, ����, �û�, ���, ɾ��, ״̬)
    Values
      (��Ϣid_In, ����_In, �û�_In, ���_In, ɾ��_In, ״̬_In);
  Elsif ����_In = 1 Then
    If ״̬_In Is Not Null Then
      If ���_In Is Null Then
        Update Zlmsgstate Set ״̬ = ״̬_In Where ��Ϣid = ��Ϣid_In And ���� = ����_In And �û� = �û�_In;
      Else
        Update Zlmsgstate
        Set ״̬ = ״̬_In, ��� = ���_In
        Where ��Ϣid = ��Ϣid_In And ���� = ����_In And �û� = �û�_In;
      End If;
    End If;
  
    If ɾ��_In Is Not Null Then
      If ����_In Is Not Null Then
        Update Zlmsgstate Set ɾ�� = ɾ��_In Where ��Ϣid = ��Ϣid_In And ���� = ����_In And �û� = �û�_In;
      Else
        Update Zlmsgstate Set ɾ�� = ɾ��_In Where ��Ϣid = ��Ϣid_In And �û� = �û�_In;
      End If;
      Select Count(*), Sum(Decode(ɾ��, 2, 1, 0)) Into n_����, n_���� From Zlmsgstate Where ��Ϣid = ��Ϣid_In;
      If n_���� = n_���� Then
        Delete From Zlmessages Where Id = ��Ϣid_In;
      End If;
    End If;
  Elsif ����_In = 2 Then
    Delete From Zlmessages Where ʱ�� < Sysdate - ��������_In;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zlmsgstate_Edit;
/

--97672:����,2017-4-11,�ϵ���Ϣ����(zlapptools),��SQL����޸�Ϊ����
CREATE OR REPLACE Procedure zltools.Zl_Zlmsgstate_Addaddressee
(
  ��Ϣid_In   Zlmsgstate.��Ϣid%Type,
  ����_In     Zlmsgstate.����%Type,
  ״̬_In     Zlmsgstate.״̬%Type,
  �û����_In Varchar2 --��ʽ���û���1,���1#�û���2,���2#�û���3,���3
) Is
Begin
  Insert Into Zlmsgstate
    (��Ϣid, ����, �û�, ���, ɾ��, ״̬)
    Select ��Ϣid_In, ����_In, C1, C2, 0, ״̬_In From Table(f_Str2list2(�û����_In, '#', ','));
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zlmsgstate_Addaddressee;
/

--97672:����,2017-4-11,�ϵ���Ϣ����(zlapptools),��SQL����޸�Ϊ����
CREATE OR REPLACE Procedure zltools.Zl_Zlmessages_New
(
  ��Ϣid_In   Zlmessages.Id%Type,
  �Ựid_In   Zlmessages.�Ựid%Type,
  �ռ���_In   Zlmessages.�ռ���%Type,
  ����_In     Zlmessages.����%Type,
  ����_In     Zlmessages.����%Type,
  ����ɫ_In   Zlmessages.����ɫ%Type,
  ����_In     Zlmsgstate.����%Type,
  �û�_In     Zlmsgstate.�û�%Type,
  ���_In     Zlmsgstate.���%Type,
  ״̬_In     Zlmsgstate.״̬%Type,
  ��������_In Number, --�������͡�1-�𸴣�2-ȫ���𸴣�3-ת����0-�½��ʼ�
  �޸�id_In   Zlmsgstate.��Ϣid%Type,
  �޸�����_In Zlmsgstate.����%Type
) Is
  n_Count Number;
Begin
  --������޸���Ϣ��¼
  Select Count(1) Into n_Count From Zlmessages Where Id = ��Ϣid_In;
  If n_Count = 0 Then
    Insert Into Zlmessages
      (Id, �Ựid, ������, ʱ��, �ռ���, ����, ����, ����ɫ)
    Values
      (��Ϣid_In, �Ựid_In, ���_In, Sysdate, �ռ���_In, ����_In, ����_In, ����ɫ_In);
  Else
    Update Zlmessages
    Set ʱ�� = Sysdate, �ռ��� = �ռ���_In, ���� = ����_In, ���� = ����_In, ����ɫ = ����ɫ_In
    Where Id = ��Ϣid_In;
  End If;

  --ɾ�����м�¼
  Delete Zlmsgstate Where ��Ϣid = ��Ϣid_In;
  --���ӷ����˼�¼
  Insert Into Zlmsgstate
    (��Ϣid, ����, �û�, ���, ɾ��, ״̬)
  Values
    (��Ϣid_In, ����_In, �û�_In, ���_In, 0, ״̬_In);
  --Ϊԭ�����ϴ𸴻�ת����־
  If ��������_In = 1 Or ��������_In = 2 Then
    Update Zlmsgstate
    Set ״̬ = Substr(״̬, 1, 1) || '1' || Substr(״̬, 3, 2)
    Where ��Ϣid = �޸�id_In And ���� = �޸�����_In And �û� = �û�_In;
  Elsif ��������_In = 3 Then
    Update Zlmsgstate
    Set ״̬ = Substr(״̬, 1, 1) || '11' || Substr(״̬, 4, 1)
    Where ��Ϣid = �޸�id_In And ���� = �޸�����_In And �û� = �û�_In;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zlmessages_New;
/

--97672:����,2017-4-17,zlSvrNotice�Զ����ѷ���,��SQL����޸�Ϊ����
Create Or Replace Procedure Zltools.Zl_Zlnoticerec_Edit
(
  ����_In     Number, --0:���ӣ�1:�޸ģ�2:ɾ��
  �������_In Zlnoticerec.�������%type,
  �û���_In   Zlnoticerec.�û���%type,
  ���ʱ��_In Zlnoticerec.���ʱ��%type,
  �����_In Zlnoticerec.�����%type,
  ���ѱ�־_In Zlnoticerec.���ѱ�־%type,
  �Ѷ���־_In Zlnoticerec.�Ѷ���־%type,
  ��������_In Zlnoticerec.��������%type
) Is
Begin
  If ����_In = 0 Then
    Insert Into Zlnoticerec
      (�������, �û���, ���ʱ��, �����, ��������, ���ѱ�־, �Ѷ���־)
    Values
      (�������_In, �û���_In, ���ʱ��_In, �����_In, ��������_In, ���ѱ�־_In, �Ѷ���־_In);
  Elsif ����_In = 1 Then
    If �Ѷ���־_In Is Null Then
      Update Zlnoticerec
      Set ����ʱ�� = Sysdate, ���ѱ�־ = ���ѱ�־_In
      Where ������� = �������_In And �û��� = �û���_In;
    Else
      Update Zlnoticerec Set �Ѷ���־ = �Ѷ���־_In Where ������� = �������_In And �û��� = �û���_In;
    End If;
  Else
    Delete From Zlnoticerec Where ������� = �������_In And �û��� = �û���_In;
  End If;
End Zl_Zlnoticerec_Edit;
/

--105511:������,2017-06-21,�Զ��屨�����ӷ���
Create Table Zltools.Zlrptclasses(
  ID Number(18), 
  �ϼ�id Number(18), 
  ���� Varchar2(30), 
  ˵�� Varchar2(100)
) PCTFREE 5;

--105511:������,2017-06-21,�Զ��屨�����ӷ���
alter table Zltools.zlReports add ����ID Number(18);
alter table Zltools.zlRPTGroups add ����ID Number(18);
Create Sequence Zltools.Zlrptclasses_Id Start With 1;
Alter Table Zltools.Zlrptclasses Add Constraint Zlrptclasses_Pk Primary Key(ID) Using Index;
Alter Table Zltools.Zlrptclasses Add Constraint Zlrptclasses_Uq_���� Unique(����) Using Index;
Alter Table Zltools.Zlrptclasses Add Constraint Zlrptclasses_Fk_�ϼ�id Foreign Key(�ϼ�id) References Zlrptclasses(ID) On Delete Cascade;
Alter Table Zltools.zlReports Add Constraint Zlreports_Uq_����id Unique(����id, ID) Using Index;
alter table Zltools.zlReports add constraint ZLREPORTS_FK_����ID foreign key (����ID) references zlRPTClasses (ID);
alter table Zltools.zlRPTGroups add constraint ZLRPTGROUPS_UQ_����ID unique (����ID, ID) using index;
alter table Zltools.zlRPTGroups add constraint ZLRPTGROUPS_FK_����ID foreign key (����ID) references zlRPTClasses (ID);
Create Index Zltools.Zlrptclasses_Ix_�ϼ�id On Zlrptclasses(�ϼ�id);

--00000:������,2017-07-06,���ݱ����
Create Table ZLTools.zlTables(
    ϵͳ    Number(5),
    ����    Varchar2(30),
    ��ռ�  Varchar2(30),
    ����    Varchar2(3)
);
Alter Table ZLTools.zlTables Add Constraint zlTables_PK Primary Key (����,ϵͳ) USING INDEX PCTFREE 5;

--A1:��̬��������
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLBAKTABLEINDEX','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLBAKTABLES','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLBASECODE','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLBIGTABLES','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLCOMPONENT','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLDATAMOVE','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLFILES','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLFILESEXPIRED','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLFILESUPGRADE','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLKILLPROCESS','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLMENUS','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLMODULERELAS','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLOPTIONS','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLPARAMETERS','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLPINYIN','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLPROGFUNCS','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLPROGPRIVS','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLPROGRAMS','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLPROGRELAS','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLREGFUNC','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLREGINFO','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLSVRTOOLS','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLSYSTEMS','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLTools.zlTables','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLUSUALFUNC','ZLTOOLSTBS','A1');

--A2:��̬��������
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLAUTOJOBS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLBAKSPACES','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLCLIENTPARALIST','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLCLIENTPARASET','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLCLIENTSCHEME','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLCONNECTIONS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLDEPTPARAS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLFUNCPARS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLFUNCTIONS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLINSUREBASE','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLINSURECOMPONENTS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLINSUREFUNCS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLINSUREMODULS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLINSUREOPERATION','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLINSUREPRIVS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLMGRGRANT','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLNODELIST','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLNOTICES','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLNOTICEUSR','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLPERIODS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLREPORTS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLROLEGRANT','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLROLEGROUPS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLRPTCLASSES','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLRPTCOLPROTERTY','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLRPTCONDS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLRPTDATAS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLRPTFMTS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLRPTGRAPHS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLRPTGROUPS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLRPTITEMS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLRPTPARS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLRPTPUTS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLRPTRELATION','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLRPTSQLS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLRPTSUBS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLSYSFILES','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLUNITINFOIMAGE','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLUNITINFOITEM','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLUPGRADESERVER','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLUSERPARAS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLUSERROLES','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLXLSDIRECTORY','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLXLSVERIFY','ZLTOOLSTBS','A2');

--A3:֪ʶ����

--B1:ҵ������
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLMESSAGES','ZLTOOLSTBS','B1');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLMSGSTATE','ZLTOOLSTBS','B1');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLPROCEDURE','ZLTOOLSTBS','B1');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLPROCEDURETEXT','ZLTOOLSTBS','B1');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLUPGRADE','ZLTOOLSTBS','B1');

--B2:��ʱ����
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLREGAUDIT','','B2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLREGFILE','','B2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLTRIGGERS','ZLTOOLSTBS','B2');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLUPGRADECONFIG','ZLTOOLSTBS','B2');

--B3:��־����
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLCLIENTUPDATELOG','ZLTOOLSTBS','B3');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLDATAMOVELOG','ZLTOOLSTBS','B3');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLDIARYLOG','ZLTOOLSTBS','B3');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLERRORLOG','ZLTOOLSTBS','B3');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLNOTICEREC','ZLTOOLSTBS','B3');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLPARACHANGEDLOG','ZLTOOLSTBS','B3');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLRPTRUNHISTORY','ZLTOOLSTBS','B3');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLRPTSQLSHISTORY','ZLTOOLSTBS','B3');
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLUPGRADELOG','ZLTOOLSTBS','B3');

--C1:״̬����
Insert into ZLTools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLCLIENTS','ZLTOOLSTBS','C1');

--C2:��������

--C3:�������