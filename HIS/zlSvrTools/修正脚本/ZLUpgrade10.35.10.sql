----10.35.0---��10.35.10
-----------------------------------------------------------------------------------------------------
-----------------------------------------------------------------------------------------------------
--93723:����,2016-03-15,�޸Ŀͻ����Զ��������򣬶Բ���Ҫ���ļ�����ɾ��
Create Table ZLTOOLS.ZLFILESEXPIRED
(
  �ļ���  VARCHAR2(50),
  ��װ·�� VARCHAR2(250),
  ϵͳ��� NUMBER(5),
  ϵͳ�汾 VARCHAR2(10),
  ˵��   VARCHAR2(250))
  PCTFREE 5 initrans 20;
INSERT INTO zltools.zlFilesExpired(�ļ���,��װ·��,ϵͳ���,ϵͳ�汾,˵��) VALUES('zlLogin.dll','[APPSOFT]\zlQueueShow',Null,'10.35.10','����ͳһ��¼����������ԭ���������Ŷӽкŵĵ�¼����');

--88843:��˶,2016-01-11,�Զ���洢���̹���
alter table zltools.ZLPROCEDURETEXT drop constraint ZLPROCEDURETEXT_FK_����ID;
alter table zltools.ZLPROCEDURETEXT add constraint ZLPROCEDURETEXT_FK_����ID foreign key (����ID) references ZLPROCEDURE (ID)  on delete cascade;
alter  table zltools.ZLPROCEDURETEXT add constraint ZLPROCEDURETEXT_PK primary key(����ID,����,���) using index;
Alter Table zltools.Zlprocedure Add Constraint Zlprocedure_Uq_���� Unique(����)Using Index Pctfree 5;
--91515:����,2015-12-10,�Զ��屨��SQL�༭���������ú���ѡ��
CREATE TABLE Zltools.zlUsualFunc(
    ϵͳ��� NUMBER(5),
    ���� VARCHAR2(50),
    ˵�� VARCHAR2(500))
    PCTFREE 5;
ALTER TABLE Zltools.zlUsualFunc ADD CONSTRAINT zlUsualFunc_PK UNIQUE (ϵͳ���,����)  USING INDEX;

CREATE Sequence Zltools.zlRPTRunHistory_ID start with 1;
CREATE TABLE Zltools.zlRPTRunHistory(
    ID       NUMBER(18),
    ����ID NUMBER(18),
    ִ����ԱID Number(18),
    ִ�п�ʼʱ�� Date,
    ִ�н���ʱ�� Date)
    PCTFREE 5 initrans 20;
ALTER TABLE Zltools.zlRPTRunHistory ADD CONSTRAINT zlRPTRunHistory_PK PRIMARY KEY (ID) USING INDEX;
ALTER TABLE Zltools.zlRPTRunHistory ADD CONSTRAINT zlRPTRunHistory_FK_����ID FOREIGN KEY(����ID) REFERENCES zlReports(ID) ON DELETE CASCADE;
CREATE INDEX zlRPTRunHistory_IX_����ID   ON Zltools.zlRPTRunHistory(����ID) PCTFREE 5;

Insert Into Zltools.zlUsualFunc(ϵͳ���,����,˵��)
Select 0,'zlSpellCode','��ȡ�����ַ����е�ÿ���ֵ�ƴ������ĸ,ȱʡ������10λ' From Dual Union All
Select 0,'zlWBCode','��ȡ�����ַ����е�ÿ���ֵ��������ĸ,ȱʡ������10λ' From Dual Union All
Select 0,'Zlpinyincode','��ȡ�����ַ��е�ÿ���ֵ�ƴ������ĸ��ȫƴ,֧�ֺ��ֶ����֡�ȫƴ����ĸ��д�ͷָ�����ȱʡ������10λ' From Dual Union All
Select 0,'zlUppMoney','�����ֽ��ת��Ϊ���ִ�д����ַ���' From Dual Union All
Select 0,'zlUppNumber','�����ִ�д����ַ���ת��Ϊ���ֽ��' From Dual Union All
Select 0,'Zl_To_Number','���ַ������ֻ�ϵ��ַ���������ת��Ϊȥ���ַ��������������ݣ���2����������1��' From Dual Union All
Select 0,'f_str2list','���ɶ��ŷָ��Ĳ������ŵ��ַ�����ת��Ϊ�������ݱ�����ΪColumn_Value,�����Table������ʹ�ã��������У�����f_Str2list2�� ' From Dual Union All
Select 0,'f_num2list','���ɶ��ŷָ��Ĳ������ŵ���������ת��Ϊ�������ݱ�����ΪColumn_Value,�����Table������ʹ�ã��������У�����f_Num2list2�� ' From Dual Union All
Select 0,'f_list2str','�����ж��е��ַ��б�ƴ��Ϊһ��ȱʡ�Զ��ŷָ����ַ����������Collect������ʹ�ã����������WM_CONCAT��sys_connect_by_path����' From Dual Union All
Select 0,'zl_GetSysParameter','��ȡϵͳ��������ģ����������ż������Ĳ���ֵ��֧�ְ������Ż��������ȡ' From Dual Union All
Select 100,'Zl_Identity','��ȡ��ǰ�û��Ĳ�����Ա��Ϣ' From Dual Union All
Select 100,'Zl_Username','��ȡ��ǰ�û�������' From Dual Union All
Select 100,'zl_IncStr','��ȡ�ַ�����Ascii��������ַ���' From Dual Union All
Select 100,'Zl_Incstr_Pre','��ȡ�ַ�����Ascii�ݼ�����ַ���' From Dual Union All
Select 100,'Zl_Age_Calc','���ݳ������ڼ��㲡�����䣬���ذ������䵥λ���ַ���' From Dual Union All
Select 100,'ZL_AgeToDays','���ݰ������䵥λ���ַ���������������' From Dual Union All
Select 100,'Zl_Cent_Money','���ݷֱҴ�������ϵͳ���������ؾ����봦���Ľ������' From Dual;


--91515:����,2016-04-08,�Զ��屨���������֧�ֹ������ű���
alter table Zltools.Zlrptrelation Add Ĭ�� Number(1);

--94568:��˶,2016-03-26,������Ϣά��
CREATE TABLE zlTools.zlUnitInfoItem(
    ���� VARCHAR2(3),
    ���� VARCHAR2(20),
    �Ƿ�ͼƬ NUMBER(1))
    PCTFREE 5;
Alter Table zlTools.zlUnitInfoItem Add Constraint zlUnitInfoItem_PK Primary Key (����) USING INDEX PCTFREE 5;
Alter Table zlTools.zlUnitInfoItem Add Constraint zlUnitInfoItem_UQ_���� Unique (����) USING INDEX PCTFREE 5;                 

CREATE TABLE zlTools.zlUnitInfoImage(   
    ��Ŀ VARCHAR2(20),
    ͼƬ Blob)
    PCTFREE 5
    Cache Storage(Buffer_Pool Keep);
Alter Table zlTools.zlUnitInfoImage Add Constraint zlUnitInfoImage_PK Primary Key (��Ŀ) USING INDEX PCTFREE 5;
ALTER TABLE zlTools.zlUnitInfoImage ADD CONSTRAINT  zlUnitInfoImage_FK_��Ŀ FOREIGN KEY (��Ŀ) REFERENCES zlTools.zlUnitInfoItem(����) ON DELETE CASCADE;
alter Table ZLTOOLS.Zlsvrtools add ���� number(3);
--94568:��˶,2016-03-26,������Ϣά��
Update Zlsvrtools
Set ���� = 1 + (To_Number(Substr(���, 3, 2)) - 1) * 3
Where �ϼ� Is Not Null And (�ϼ� = '03' And ��� = '0301' Or �ϼ� <> '03');
Update Zlsvrtools
Set ���� = 1 + (To_Number(Substr(���, 3, 2))) * 3
Where �ϼ� Is Not Null And �ϼ� = '03' And ��� <> '0301';
Update Zlsvrtools Set ���� = 4 Where ��� = '0312';
Insert Into zlSvrTools(���,�ϼ�,����,���,˵��,����) Values('0312','03','ҽԺ��Ϣά��','H',Null,4);
insert into zlTools.zlUnitInfoItem(����,����,�Ƿ�ͼƬ)values('001','ҽԺ����',0);
insert into zlTools.zlUnitInfoItem(����,����,�Ƿ�ͼƬ)values('002','��ͨ',0);
insert into zlTools.zlUnitInfoItem(����,����,�Ƿ�ͼƬ)values('003','��ַ',0);
insert into zlTools.zlUnitInfoItem(����,����,�Ƿ�ͼƬ)values('004','�绰',0);
insert into zlTools.zlUnitInfoItem(����,����,�Ƿ�ͼƬ)values('005','��ϵ��',0);
insert into zlTools.zlUnitInfoItem(����,����,�Ƿ�ͼƬ)values('006','������',0);
insert into zlTools.zlUnitInfoItem(����,����,�Ƿ�ͼƬ)values('007','��������',0);
insert into zlTools.zlUnitInfoItem(����,����,�Ƿ�ͼƬ)values('008','�����ʺ�',0);
insert into zlTools.zlUnitInfoItem(����,����,�Ƿ�ͼƬ)values('009','˰���',0);
insert into zlTools.zlUnitInfoItem(����,����,�Ƿ�ͼƬ)values('010','Ժ��',0);
insert into zlTools.zlUnitInfoItem(����,����,�Ƿ�ͼƬ)values('011','ҽԺ�ȼ�',0);
insert into zlTools.zlUnitInfoItem(����,����,�Ƿ�ͼƬ)values('012','�����ʼ�',0);
insert into zlTools.zlUnitInfoItem(����,����,�Ƿ�ͼƬ)values('013','ҽԺ����',0);
insert into zlTools.zlUnitInfoItem(����,����,�Ƿ�ͼƬ)values('014','��ҳ',0);
--88843:��˶,2016-04-14,�Զ���洢���̹���
Create Or Replace Procedure Zltools.Zl_Zlproceduretext_Move Is
Begin
  Delete From Zlproceduretext Where ���� In (1, 2);
  Update Zlprocedure Set ״̬ = 0;
  Insert Into Zlproceduretext
    (����id, ����, ���, ����)
    Select ����id, Decode(����, 3, 1, 4, 2), ���, ���� From Zlproceduretext Where ���� In (3, 4);
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zlproceduretext_Move;
/

Create Or Replace Procedure Zltools.Zl_Zlprocedure_Confirm(Id_In Zlprocedure.Id%Type) Is
Begin
  Update Zlprocedure Set ״̬ = 3 Where Id = Id_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zlprocedure_Confirm;
/
Create Or Replace Procedure Zltools.Zl_Zlprocedure_Manage(Nstep_In Number := 0) Is
  --Nstep_In 0-�ռ�ǰ���ã���ʱ�����й��̱��Ϊ�����״̬��
  --         1-������ã���ʱ��״̬�Ծɴ��ڴ������û����̵���Ϊ���������Դ����Ķ��������䶯���̱��Ϊ�ޱ仯
Begin
  If Nvl(Nstep_In, 0) = 0 Then
    Update Zlprocedure Set ״̬ = 0;
  Else
    Update Zlprocedure Set ״̬ = 1 Where ���� = 3 And ״̬ = 0;
    Update Zlprocedure Set ״̬ = 4 Where ���� In (1, 2) And ״̬ = 0;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zlprocedure_Manage;
/
--94568:��˶,2016-03-26,������Ϣά��
Create Or Replace Package Zltools.b_Public Is
  --��������
  Type t_Refcur Is Ref Cursor;
  --���ܣ�ȡϵͳ����
  --�����б�mdlMain.CurrentDate��clsDatabase.CurrentDate
  Procedure Get_Current_Date(Cursor_Out Out t_Refcur);
  --���ܣ�ɾ��������־��������־
  --�����б�mdlMain.DeleteAllLog
  Procedure Delete_All_Log(Runtimelog_In In Number := 0);
  --���ܣ�ɾ����ǰ������־
  --�����б�mdlMain.DeleteCurLog
  Procedure Delete_Diarylog
  (
    �Ự��_In   Number,
    �û���_In   Varchar2,
    ����վ_In   Varchar2,
    ������_In   Varchar2,
    ��������_In Varchar2,
    ����ʱ��_In Date
  );
  --���ܣ�ɾ����ǰ������־
  --�����б�mdlMain.DeleteCurLog
  Procedure Delete_Errorlog
  (
    �Ự��_In   Number,
    �û���_In   Varchar2,
    ����վ_In   Varchar2,
    ����_In     Number,
    �������_In Number,
    ʱ��_In     Date
  );
  --���ܣ�ȡע����
  --�����б�mdlMain.Getע����
  Procedure Get_Regcode(Cursor_Out Out t_Refcur);
  --���ܣ�ȡ�汾��
  --�����б�mdlMain.UpgradeManager
  Procedure Get_Ver(Cursor_Out Out t_Refcur);
  --���ܣ����°汾��
  --�����б�mdlMain.UpgradeManager
  Procedure Update_Ver(Verstring_In In Varchar2);
  --���ܣ�ȡ��ϵͳ����������
  --�����б�
  --frmStatus.cmbsystem_Click��mdlMain.GetOwnerName��mdlMain.cmbSystem_Click
  --frmAutoJobs.cmbSystem_Click��frmDataMove.cmbSystem_Click ��frmNoticeTools.cboSystem_Click
  --frmProgPriv.ProgPriv��frmAppScript.cmbSystem_Click
  Procedure Get_Owner_Name
  (
    Cursor_Out Out t_Refcur,
    ���_In    In Zlsystems.���%Type := 0
  );

  --���ܣ�ȡע�������Ϣ
  --�����б�
  --frmAbout.GetUnitInfo��frmAutoJobs.From_load��frmClientsUpgrade.InitInfor
  --frmFilesSet.ShowEdit��frmRegist.From_load��frmAppScript.From_Load
  --frmFilesSendToServer.InitInfo
  Procedure Get_Reginfo
  (
    Cursor_Out Out t_Refcur,
    ��Ŀ_In    In Zlreginfo.��Ŀ%Type := Null
  );
  --���ܣ�ȡzlGetSvrToolsg����
  --�����б�frmMDIMain.MDIForm_Load
  Procedure Get_Zlsvrtools(Cursor_Out Out t_Refcur);
  --���ܣ�ȡ�Ѱ�װϵͳ�嵥
  --�����б�
  --frmAppCheck.Form_Load��frmClearData.Form_Load��frmDataMove.Form_Load
  --frmImp.FillSystem��frmLoadIn.FillSystem��frmLoadOut.FillSystem
  --frmMDIMain.mnuFileRemove_Click��frmNoticeTools.Form_Activate��frmRoleGrant.FillSystem
  --frmAppUpgrade.Form_Load��frmAppScript.Form_Load��frmExp.FillSystem
  --frmInputTools.from_activate��fromRole.FillSystem��frmAutoJobs.From_load
  --frmAppstart.sysCreated
  Procedure Get_Zlsystems
  (
    Cursor_Out Out t_Refcur,
    ������_In  In Zlsystems.������%Type := Null
  );
  --���ܣ��洢BLObͼƬ
  --�����б�frmUnitInfoEdit
  --����˵����
  --Tab_In������LOB�����ݱ�
  --        0-zlUnitInfoImage
  --Key_In�����ݼ�¼�Ĺؼ���
  --Txt_In��16���Ƶ��ļ�Ƭ�λ�����Ƭ��
  --Cls_In���Ƿ����ԭ�������ݣ���һƬ�δ���ʱΪ1���Ժ�Ϊ0
  Procedure Zllobappend
  (
    Tab_In In Number,
    Key_In In Varchar2,
    Txt_In In Varchar2,
    Cls_In In Number := 0
  );
  --���ܣ���ȡBLObͼƬ
  --�����б�frmUnitInfoEdit
  --����˵����
  --Tab_In������LOB�����ݱ�
  --        0-zlUnitInfoImage
  --Key_In�����ݼ�¼�Ĺؼ���
  --Pos_In����0��ʼ���϶�ȡ��ֱ������Ϊ��
  Function Zllobread
  (
    Tab_In In Number,
    Key_In In Varchar2,
    Pos_In In Number
  ) Return Varchar2;

  --���ܣ�����ɾ�������ZLRegInfoͼƬ
  --�����б�frmUnitInfoEdit
  --����˵����
  --��Ŀ_In����Ŀ����
  --�к�_In:Ϊ�ջ���Ϊ1ʱ��ɾ������Ŀ��Ȼ�����
  --����_In:Ϊ���򲻲���
  Procedure Zlreginfoupdate
  (
    ��Ŀ_In In Zlreginfo.��Ŀ%Type,
    �к�_In In Zlreginfo.�к�%Type,
    ����_In In Zlreginfo.����%Type
  );
  --���ܣ�����ɾ�������Zlunitinfoitem
  --�����б�frmUnitInfoEdit,frmUnitItemEdit
  --����˵����
  --Type_n:0-������1-�޸�,2-ɾ��
  --����_In����Ŀ����
  --����_In:��Ŀ����
  --ͼƬ_In:��Ŀ�Ƿ���ͼƬ����
  Procedure Zlunitinfoitemchange
  (
    Type_n  In Number,
    ����_In In Zlunitinfoitem.����%Type,
    ����_In In Zlunitinfoitem.����%Type := Null,
    ͼƬ_In In Zlunitinfoitem.�Ƿ�ͼƬ%Type := Null
  );
End b_Public;
/

--94568:��˶,2016-03-26,������Ϣά��
Create Or Replace Package Body Zltools.b_Public Is
  --���ܣ�ȡϵͳ����
  Procedure Get_Current_Date(Cursor_Out Out t_Refcur) Is
  Begin
    Open Cursor_Out For
      Select Sysdate As ���� From Dual;
  End Get_Current_Date;

  --���ܣ�ɾ��������־��������־
  Procedure Delete_All_Log(Runtimelog_In In Number := 0) Is
    n_Count Number;
    n_Loop  Number;
  Begin
    If Runtimelog_In = 1 Then
      Select Count(����ʱ��) Into n_Count From Zldiarylog;
      If n_Count > 1000 Then
        For n_Loop In 1 .. Ceil(n_Count - 1000) Loop
          Delete Zldiarylog Where Rownum < 10001;
          Commit;
        End Loop;
      Else
        If n_Count > 0 Then
          Delete Zldiarylog;
          Commit;
        End If;
      End If;
    Else
      Select Count(ʱ��) Into n_Count From Zlerrorlog;
      If n_Count > 1000 Then
        For n_Loop In 1 .. Ceil(n_Count - 1000) Loop
          Delete Zlerrorlog Where Rownum < 10001;
          Commit;
        End Loop;
      Else
        If n_Count > 0 Then
          Delete Zlerrorlog;
          Commit;
        End If;
      End If;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Delete_All_Log;

  --���ܣ�ɾ����ǰ������־
  Procedure Delete_Diarylog
  (
    �Ự��_In   Number,
    �û���_In   Varchar2,
    ����վ_In   Varchar2,
    ������_In   Varchar2,
    ��������_In Varchar2,
    ����ʱ��_In Date
  ) Is
  Begin
    Delete Zldiarylog
    Where �Ự�� = �Ự��_In And �û��� = �û���_In And ����վ = ����վ_In And ������ = ������_In And �������� = ��������_In And ����ʱ�� = ����ʱ��_In;
    Commit;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Delete_Diarylog;

  --���ܣ�ɾ����ǰ������־
  Procedure Delete_Errorlog
  (
    �Ự��_In   Number,
    �û���_In   Varchar2,
    ����վ_In   Varchar2,
    ����_In     Number,
    �������_In Number,
    ʱ��_In     Date
  ) Is
  Begin
    Delete Zlerrorlog
    Where �Ự�� = �Ự��_In And �û��� = �û���_In And ����վ = ����վ_In And ���� = ����_In And ������� = �������_In And ʱ�� = ʱ��_In;
    Commit;
  
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Delete_Errorlog;

  --���ܣ�ȡע����
  Procedure Get_Regcode(Cursor_Out Out t_Refcur) Is
  Begin
    Open Cursor_Out For
      Select ���� From Zlreginfo Where ��Ŀ = 'ע����' Or ��Ŀ = '��Ȩ֤��' Order By �к�;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Regcode;

  --���ܣ�ȡ�汾��
  Procedure Get_Ver(Cursor_Out Out t_Refcur) Is
  Begin
    Open Cursor_Out For
      Select ���� From Zlreginfo Where ��Ŀ = '�汾��';
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Ver;

  --���ܣ����°汾��
  Procedure Update_Ver(Verstring_In In Varchar2) Is
  Begin
    Update Zlreginfo Set ���� = Verstring_In Where ��Ŀ = '�汾��';
    If Sql%Notfound Then
      Insert Into Zlreginfo (��Ŀ, �к�, ����) Values ('�汾��', 1, Verstring_In);
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Update_Ver;

  --���ܣ�ȡ��ϵͳ����������
  Procedure Get_Owner_Name
  (
    Cursor_Out Out t_Refcur,
    ���_In    In Zlsystems.���%Type := 0
  ) Is
  Begin
    Open Cursor_Out For
      Select Upper(������) As ������ From Zlsystems Where ��� = ���_In;
  End Get_Owner_Name;

  --���ܣ�ȡע�������Ϣ
  Procedure Get_Reginfo
  (
    Cursor_Out Out t_Refcur,
    ��Ŀ_In    In Zlreginfo.��Ŀ%Type := Null
  ) Is
  Begin
    If Trim(Nvl(��Ŀ_In, '��')) = '��' Then
      Open Cursor_Out For
        Select * From Zlreginfo;
    Else
      Open Cursor_Out For
        Select ���� From Zlreginfo Where ��Ŀ = ��Ŀ_In Order By �к�;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Reginfo;

  --���ܣ�ȡzlGetSvrToolsg����
  Procedure Get_Zlsvrtools(Cursor_Out Out t_Refcur) Is
  Begin
    Open Cursor_Out For
      Select * From Zlsvrtools Start With �ϼ� Is Null Connect By Prior ��� = �ϼ� Order By Level, ���;
  End Get_Zlsvrtools;

  --���ܣ�ȡ�Ѱ�װϵͳ�嵥
  Procedure Get_Zlsystems
  (
    Cursor_Out Out t_Refcur,
    ������_In  In Zlsystems.������%Type := Null
  ) Is
  Begin
    If Nvl(������_In, '��') = '��' Then
      Open Cursor_Out For
        Select ���, ����, �����, Upper(������) ������, ��װ����, ������װ, �汾�� From Zlsystems Order By ���;
    Else
      Open Cursor_Out For
        Select ���, ����, �����, Upper(������) ������, ��װ����, ������װ, �汾��
        From Zlsystems
        Where Upper(������) = Upper(������_In)
        Order By ���;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Zlsystems;
  --���ܣ��洢BLObͼƬ
  Procedure Zllobappend
  (
    Tab_In In Number,
    Key_In In Varchar2,
    Txt_In In Varchar2,
    Cls_In In Number := 0
    --����˵���� 
    --Tab_In������LOB�����ݱ� 
    --        0-zlUnitInfoImage
    --Key_In�����ݼ�¼�Ĺؼ��� 
    --Txt_In��16���Ƶ��ļ�Ƭ�λ�����Ƭ�� 
    --Cls_In���Ƿ����ԭ�������ݣ���һƬ�δ���ʱΪ1���Ժ�Ϊ0 
  ) Is
    l_Blob Blob;
  Begin
    If Tab_In = 0 Then
      If Txt_In Is Null And Cls_In = 1 Then
        Delete Zltools.Zlunitinfoimage Where ��Ŀ = Key_In;
      Else
        If Cls_In = 1 Then
          Update Zltools.Zlunitinfoimage Set ͼƬ = Empty_Blob() Where ��Ŀ = Key_In;
          If Sql%Rowcount = 0 Then
            Insert Into Zltools.Zlunitinfoimage (��Ŀ, ͼƬ) Values (Key_In, Empty_Blob());
          End If;
        End If;
        Select ͼƬ Into l_Blob From Zltools.Zlunitinfoimage Where ��Ŀ = Key_In For Update;
      End If;
    End If;
    If Tab_In = 0 And Txt_In Is Null And Cls_In = 1 Then
      Null;
    Else
      Dbms_Lob.Writeappend(l_Blob, Length(Hextoraw(Txt_In)) / 2, Hextoraw(Txt_In));
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Zllobappend;
  --���ܣ���ȡBLObͼƬ
  Function Zllobread
  (
    Tab_In In Number,
    Key_In In Varchar2,
    Pos_In In Number --����˵���� 
    --Tab_In������LOB�����ݱ� 
    --        0-zlUnitInfoImage
    --Key_In�����ݼ�¼�Ĺؼ��� 
    --Pos_In����0��ʼ���϶�ȡ��ֱ������Ϊ�� 
  ) Return Varchar2 Is
    l_Blob   Blob;
    v_Buffer Varchar2(32767);
    n_Amount Number := 2000;
    n_Offset Number := 1;
  Begin
    If Tab_In = 0 Then
      Select ͼƬ Into l_Blob From Zltools.Zlunitinfoimage Where ��Ŀ = Key_In;
    End If;
    n_Offset := n_Offset + Pos_In * n_Amount;
    If l_Blob Is Null Then
      v_Buffer := Null;
    Else
      Dbms_Lob.Read(l_Blob, n_Amount, n_Offset, v_Buffer);
    End If;
    Return v_Buffer;
  Exception
    When No_Data_Found Then
      Return Null;
  End Zllobread;

  Procedure Zlreginfoupdate
  (
    ��Ŀ_In In Zlreginfo.��Ŀ%Type,
    �к�_In In Zlreginfo.�к�%Type,
    ����_In In Zlreginfo.����%Type
  ) Is
    --��Ŀ_In����Ŀ����
    --�к�_In:Ϊ�ջ���Ϊ1ʱ��ɾ������Ŀ��Ȼ�����
    --����_In:Ϊ���򲻲���
  Begin
    If Nvl(�к�_In, 0) < 2 Then
      Delete Zlreginfo Where ��Ŀ = ��Ŀ_In;
    End If;
    If Not ����_In Is Null Then
      Insert Into Zlreginfo (��Ŀ, �к�, ����) Values (��Ŀ_In, �к�_In, ����_In);
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Zlreginfoupdate;
  Procedure Zlunitinfoitemchange
  (
    Type_n  In Number,
    ����_In In Zlunitinfoitem.����%Type,
    ����_In In Zlunitinfoitem.����%Type := Null,
    ͼƬ_In In Zlunitinfoitem.�Ƿ�ͼƬ%Type := Null
  ) Is
    --Type_n:0-������1-�޸�,2-ɾ��
    --����_In����Ŀ����
    --����_In:��Ŀ����
    --ͼƬ_In:��Ŀ�Ƿ���ͼƬ����
    v_���� Zlunitinfoitem.����%Type;
    n_ͼƬ Zlunitinfoitem.�Ƿ�ͼƬ%Type;
  Begin
    If Type_n = 0 Then
      Insert Into Zlunitinfoitem (����, ����, �Ƿ�ͼƬ) Values (����_In, ����_In, ͼƬ_In);
    Elsif Type_n = 1 Then
      Select Nvl(����, 0), Nvl(�Ƿ�ͼƬ, 0) Into v_����, n_ͼƬ From Zlunitinfoitem Where ���� = ����_In;
      --���ڸ���Ŀ
      If Not n_ͼƬ Is Null Then
        --���ͱ����ɾ����������
        If n_ͼƬ <> Nvl(ͼƬ_In, 0) Then
          If n_ͼƬ = 0 Then
            Delete Zlreginfo Where ��Ŀ = v_����;
          Else
            Delete Zlunitinfoimage Where ��Ŀ = v_����;
          End If;
          --���Ʊ��
        Elsif v_���� <> Nvl(����_In, '�տ�') Then
          If n_ͼƬ = 0 Then
            Update Zlreginfo Set ��Ŀ = ����_In Where ��Ŀ = v_����;
          Else
            Update Zlunitinfoimage Set ��Ŀ = ����_In Where ��Ŀ = v_����;
          End If;
        End If;
        Update Zlunitinfoitem Set ���� = ����_In, �Ƿ�ͼƬ = ͼƬ_In Where ���� = ����_In;
      End If;
    Else
      Select Nvl(����, 0), Nvl(�Ƿ�ͼƬ, 0) Into v_����, n_ͼƬ From Zlunitinfoitem Where ���� = ����_In;
      --���ڸ���Ŀ
      If Not n_ͼƬ Is Null Then
        If n_ͼƬ = 0 Then
          Delete Zlreginfo Where ��Ŀ = v_����;
        Else
          Delete Zlunitinfoimage Where ��Ŀ = v_����;
        End If;
        Delete Zlunitinfoitem Where ���� = ����_In;
      End If;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Zlunitinfoitemchange;
End b_Public;
/
--92592:��˶,2015-01-18,��¼�����û���ɫ
Delete Zltools.Zluserroles;
Insert Into Zltools.Zluserroles
(�û�, ��ɫ, ����)
Select Grantee, Granted_Role, Decode(Admin_Option, 'YES', 1, 0)
From Dba_Role_Privs
Where Granted_Role Like 'ZL_%'
Order By Grantee;
--92592:��˶,2015-01-18,��¼�����û���ɫ
Create Or Replace Procedure Zltools.Zl_Zluserroles_Add
(
  User_In In Zluserroles.�û�%Type := Null,
  Role_In In Zluserroles.��ɫ%Type := Null,
  ����_In In Zluserroles.����%Type := 0
) Is
  --���û���ɫ��Ϊ��ʱ����¼�����û���ɫ����
Begin
  If Not User_In Is Null And Not Role_In Is Null Then
    Insert Into Zluserroles (�û�, ��ɫ, ����) Values (User_In, Role_In, ����_In);
    --�û���ɫ������ʱ������������ݣ�����������
  Elsif User_In Is Null And Role_In Is Null Then
    Delete Zltools.Zluserroles;
    Insert Into Zltools.Zluserroles
      (�û�, ��ɫ, ����)
      Select Grantee, Granted_Role, Decode(Admin_Option, 'YES', 1, 0)
      From Dba_Role_Privs
      Where Granted_Role Like 'ZL_%'
      Order By Grantee;
    --�û�����ʱ����ո��û�����������������
  Elsif Not User_In Is Null Then
    Delete Zltools.Zluserroles Where �û� = User_In;
    Insert Into Zltools.Zluserroles
      (�û�, ��ɫ, ����)
      Select Grantee, Granted_Role, Decode(Admin_Option, 'YES', 1, 0)
      From Dba_Role_Privs
      Where Granted_Role Like 'ZL_%' And Grantee = User_In;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zluserroles_Add;
/
--92592:��˶,2015-01-18,��¼�����û���ɫ
Create Or Replace Procedure Zltools.Zl_Zluserroles_Del
(
  User_In In Zluserroles.�û�%Type,
  Role_In In Zluserroles.��ɫ%Type
) Is
Begin
  If Not User_In Is Null And Not Role_In Is Null Then
    Delete Zluserroles Where �û� = User_In And ��ɫ = Role_In;
  Elsif Not User_In Is Null Then
    Delete Zluserroles Where �û� = User_In;
  Elsif Not Role_In Is Null Then
    Delete Zluserroles Where ��ɫ = Role_In;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zluserroles_Del;
/

--91515:����,2015-12-18,��¼����ִ��ִ��ʱ��
alter table Zltools.Zlreports Add ִ����ԱID Number(18);
alter table Zltools.Zlreports Add ִ�п�ʼʱ�� Date;
alter table Zltools.Zlreports Add ִ�н���ʱ�� Date;
alter table Zltools.zlReports initrans 20;
alter table Zltools.ZlRptitems Add ����߼Ӵ� Number(1);
ALTER TABLE zltools.zlRPTITems ADD CONSTRAINT zlRPTItems_CK_����߼Ӵ� Check(����߼Ӵ� IN(1,0));

Insert Into zlParameters(ID,ϵͳ,ģ��,˽��,����,��Ȩ,�̶�,����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵��)
Select zlParameters_ID.Nextval,-Null,-Null,-Null,-Null,-Null,-Null,A.* From (
Select ����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵�� From zlParameters Where 1 = 0 Union All
select 0,0,26,'��������������־','0','0','��ر���ִ���ˣ�ִ�п�ʼʱ���ִ�н���ʱ��','0-�رգ�1-����',NULL,NULL,NULL From Dual Union All
select 0,0,27,'������ͱ�','','3000,1000000','������ͱ��¼����Χ','Ϊ���򲻼�飬��Ϊ������ݷ�Χ������ͱ�',NULL,NULL,NULL From Dual Union All
select 0,0,28,'��¼����ʹ�úۼ�','1','1','�ڴ򿪱���ˢ�¡���ѡ��ʽ�����õȲ�����ɺ��¼���һ�ε�ִ���ˡ�ִ��ʱ��','0-�رգ�1-����',NULL,NULL,NULL From Dual Union All
Select ����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵�� From ZLPARAMETERS Where 1 = 0) A;

Create Or Replace Procedure Zltools.Zl_Rptrun_Update
(
  Id_In           In Zlreports.Id%Type,
  ִ����Աid_In   In Zlreports.ִ����Աid%Type,
  ִ�п�ʼʱ��_In In Zlreports.ִ�п�ʼʱ��%Type,
  ִ�н���ʱ��_In In Zlreports.ִ�н���ʱ��%Type
) Is
Begin
  Update Zlreports
  Set ִ����Աid = ִ����Աid_In, ִ�п�ʼʱ�� = ִ�п�ʼʱ��_In, ִ�н���ʱ�� = ִ�н���ʱ��_In
  Where Id = Id_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Rptrun_Update;
/

Create Or Replace Procedure Zltools.Zl_Rptrunhistory_Update
(
  Id_In           In Zlrptrunhistory.Id%Type,
  ����id_In       In Zlreports.Id%Type,
  ִ����Աid_In   In Zlrptrunhistory.ִ����Աid%Type,
  ִ�п�ʼʱ��_In In Zlrptrunhistory.ִ�п�ʼʱ��%Type,
  ִ�н���ʱ��_In In Zlrptrunhistory.ִ�н���ʱ��%Type
) Is
Begin
  Insert Into Zlrptrunhistory
    (Id, ����id, ִ����Աid, ִ�п�ʼʱ��, ִ�н���ʱ��)
  Values
    (Id_In, ����id_In, ִ����Աid_In, ִ�п�ʼʱ��_In, ִ�н���ʱ��_In);
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Rptrunhistory_Update;
/

--00000:��˶,2015-12-25,���ʱ��֧�ֱ�������ֻ��������
Create Or Replace Function Zltools.Zl_Checkobject
(
  n_Type        In Number, --1=��,2=�ֶ�,3=Լ��,4=����
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
  End If;

  Return n_Count;
End Zl_Checkobject;
/
--88843:��˶,2016-01-11,�Զ���洢���̹���
Create Or Replace Procedure Zltools.Zl_Zlproceduretext_Update
(
  ����id_In In Zlproceduretext.����id%Type,
  ����_In   In Zlproceduretext.����%Type,
  ���_In   In Zlproceduretext.���%Type,
  ����_In   In Zlproceduretext.����%Type
) Is
Begin
  --���ڹ��̴��ڷֶδ洢�����˱�����ɾ�������
  If Nvl(���_In, 1) = 1 Then
    Delete Zlproceduretext Where ����id = ����id_In And ���� = ����_In;
  End If;
  If Not ����_In Is Null Then
    Insert Into Zlproceduretext (����id, ����, ���, ����) Values (����id_In, ����_In, ���_In, ����_In);
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zlproceduretext_Update;
/
--92669:������,2016-01-13,ע���밲ȫ���Ƽ���ѯ��������
Create Or Replace Function zltools.f_Str2list
(
  Str_In   In Varchar2,
  Split_In In Varchar2 := ','
) Return t_Strlist
  Pipelined As
  v_Str Long;
  P     Number;
  --���ܣ����ɶ��ŷָ��Ĳ������ŵ��ַ�����ת��Ϊ�������ݱ� 
  --������STR_IN,��:G0000123,G0000124,G0000125...,SPLIT_IN,�ָ���,ȱʡΪ,�� 
  --˵���� 
  --1����SQL������漰��IN(����1, ����2,��) ���Ӿ�ʱʹ�����ַ�ʽ�Ա����ð󶨱����� 
  --2��ʹ������������ʱ����Ҫ��SQL����м��롰/*+ cardinality(b 3)*/����ʾ����ΪCBO����ʱ�ڴ��û��ͳ������,�� 
  --3�����ֵ���ʾ�� 
  --SELECT /*+ cardinality(b 3)*/ * FROM ������ü�¼ WHERE NO IN (SELECT * FROM TABLE(F_STR2LIST('A01,A02,A03')) B); 
  --SELECT /*+ cardinality(b 3)*/ A.* FROM ������ü�¼ A, TABLE(F_STR2LIST('A01,A02,A03')) B WHERE A.NO = B.COLUMN_VALUE; 
Begin
  If Str_In Is Null Then
    Return;
  End If;
  v_Str := Str_In || Split_In;
  Loop
    P := Instr(v_Str, Split_In);
    Exit When(Nvl(P, 0) = 0);
    Pipe Row(Substr(v_Str, 1, P - 1));
    v_Str := Substr(v_Str, P + 1);
  End Loop;
  Return;
End;
/

--92669:������,2016-01-13,ע���밲ȫ���Ƽ���ѯ��������
Create Or Replace Function zltools.f_Str2list2
(
  Str_In      In Varchar2,
  Split_In    In Varchar2 := ',',
  Subsplit_In In Varchar2 := ':'
) Return t_Strlist2
  Pipelined As
  v_Str   Long;
  P       Number;
  v_Tmp   Varchar2(4000);
  Out_Rec t_Strobj2 := t_Strobj2(Null, Null);
Begin
  If Str_In Is Null Then
    Return;
  End If;
  v_Str := Str_In || Split_In;
  Loop
    P := Instr(v_Str, Split_In);
    Exit When(Nvl(P, 0) = 0);
    v_Tmp      := Substr(v_Str, 1, P - 1);
    Out_Rec.C1 := Substr(v_Tmp, 1, Instr(v_Tmp, Subsplit_In) - 1);
    Out_Rec.C2 := Substr(v_Tmp, Instr(v_Tmp, Subsplit_In) + 1);
    Pipe Row(Out_Rec);
    v_Str := Substr(v_Str, P + 1);
  End Loop;
  Return;
End;
/


--92669:������,2016-01-13,ע���밲ȫ���Ƽ���ѯ��������
Create Or Replace Function zltools.f_Reg_Menu
(
  Menu_Group_In  In Zlmenus.���%Type := Null, --����ѡ��Ĳ˵����
  System_List_In In Varchar2, --���λỰ�漰��Ӧ��ϵͳ
  Part_List_In   In Varchar2 --�Զ��ŷָ��ı�����ִ�в����б�
) Return t_Menu_Rowset Is
  t_Return t_Menu_Rowset := t_Menu_Rowset();
  t_Middle t_Menu_Rowset := t_Menu_Rowset();

  v_Parts   Varchar2(32767);
  t_Parts   t_Reg_Rowset := t_Reg_Rowset();
  v_Systems Varchar2(32767);
  t_Systems t_Reg_Rowset := t_Reg_Rowset();

Begin
  --���������γ����������
  v_Parts := Upper(Part_List_In) || ',';
  While v_Parts Is Not Null Loop
    t_Parts.Extend;
    t_Parts(t_Parts.Count) := t_Reg_Record(Null, Null, Substr(v_Parts, 1, Instr(v_Parts, ',') - 1));
    v_Parts := Trim(Substr(v_Parts, Instr(v_Parts, ',') + 1));
  End Loop;
  t_Parts.Extend;
  t_Parts(t_Parts.Count) := t_Reg_Record(Null, Null, 'ZL9REPORT');
  v_Systems := System_List_In || ',';
  While v_Systems Is Not Null Loop
    t_Systems.Extend;
    t_Systems(t_Systems.Count) := t_Reg_Record(Null, To_Number(Substr(v_Systems, 1, Instr(v_Systems, ',') - 1)), Null);
    v_Systems := Trim(Substr(v_Systems, Instr(v_Systems, ',') + 1));
  End Loop;
  t_Systems.Extend;
  t_Systems(t_Systems.Count) := t_Reg_Record(Null, 0, Null);

  --�˵����ݻ�ȡ��
  Select t_Menu_Record(m.���, m.Id, m.�ϼ�id, m.����, m.�̱���, m.���, m.˵��, m.ģ��, m.ϵͳ, m.ͼ��, p.����, 0)
  Bulk Collect
  Into t_Middle
  From (Select Level As ���, ID, �ϼ�id, ����, �̱���, ���, ˵��, ģ��, ϵͳ, ͼ��
         From zlMenus
         Where ��� = Menu_Group_In
         Start With �ϼ�id Is Null
         Connect By Prior ID = �ϼ�id) M,
       (Select /*+ cardinality(C 20) cardinality(S 2)*/
         Distinct p.ϵͳ, p.���, p.����
         From zlPrograms P, zlProgFuncs F, zlRegFunc R, zlRPTGroups X, Table(Cast(t_Parts As t_Reg_Rowset)) C,
              Table(Cast(t_Systems As t_Reg_Rowset)) S,
              (Select Decode(Count(*), 0, 0, 1) As ���
                From zlSystems
                Where ������ = User
                Union All
                Select ���
                From zlSystems
                Where ������ = User) O,
              (Select Distinct g.ϵͳ, g.��� From zlRoleGrant G, zlUserRoles R Where g.��ɫ = r.��ɫ And r.�û� = User) G
         Where Nvl(f.ϵͳ, 0) = Nvl(p.ϵͳ, 0) And f.��� = p.��� And Trunc(f.ϵͳ / 100) = r.ϵͳ(+) And f.��� = r.���(+) And
               f.���� = r.����(+) And
               (r.���� Is Null And f.ϵͳ Is Null Or r.���� Is Not Null And r.���� = '����' Or
                r.���� Is Not Null And x.����id Is Not Null Or r.���� Is Null And (p.��� Between 10000 And 19999)) And
               p.ϵͳ = x.ϵͳ(+) And p.��� = x.����id(+) And Upper(p.����) = c.Text And Nvl(p.ϵͳ, 0) = s.Prog And
               Nvl(p.ϵͳ, 1) = o.���(+) And Nvl(p.ϵͳ, 0) = Nvl(g.ϵͳ(+), 0) And p.��� = g.���(+) And
               (o.��� Is Not Null Or g.��� Is Not Null)) P
  Where Nvl(m.ϵͳ, 0) = Nvl(p.ϵͳ(+), 0) And m.ģ�� = p.���(+) And (m.ģ�� Is Null Or m.ģ�� Is Not Null And p.��� Is Not Null)
  Order By m.��� Desc;

  --�������¼���ִ�еĲ˵���Ŀ
  For n_Child In 1 .. t_Middle.Count Loop
    If t_Middle(n_Child).���� Is Not Null Or t_Middle(n_Child).��� = 1 Then
      t_Return.Extend;
      t_Return(t_Return.Count) := t_Middle(n_Child);
      If t_Middle(n_Child).�ϼ�id Is Not Null Then
        For n_Parent In n_Child + 1 .. t_Middle.Count Loop
          If t_Middle(n_Parent).��� = 0 And t_Middle(n_Parent).Id = t_Middle(n_Child).�ϼ�id Then
            t_Middle(n_Parent).��� := 1;
            Exit;
          End If;
        End Loop;
      End If;
    End If;
  End Loop;

  Return t_Return;
End f_Reg_Menu;
/

--93131:������,2016-01-29,�Զ����������Ż�
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
      v_Err := '����δ���пͻ���Ԥ����ʱ������ã�';
      Raise Err_Custom;
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

Drop Procedure zltools.Zl_Zlclients_Upgrade;