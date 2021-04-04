----10.35.10---��10.35.20
-----------------------------------------------------------------------------------------------------
-----------------------------------------------------------------------------------------------------
--98732:��˶,2016-08-13,����SP�汾֧��
alter table zltools.ZLFILESEXPIRED Add �Ƿ�̶� Number(1); 
alter table zltools.ZLSYSTEMS modify �汾�� varchar2(20);
alter table zltools.ZLFILESEXPIRED modify ϵͳ�汾 varchar2(20);
alter table zltools.ZLUPGRADE modify Ŀ��汾 varchar2(20);
alter table zltools.ZLUPGRADE modify ����汾 varchar2(20);
alter table zltools.ZLUPGRADE modify ԭʼ�汾 varchar2(20);
alter table zltools.ZLCOMPONENT modify ע���Ʒ�汾 varchar2(20);
alter table zltools.ZLUPGRADE modify Ŀ��汾 varchar2(20);

create table zltools.ZLFiles(
  ����  Varchar2(50),
  ��׼MD5  Varchar2(32),
  �汾��  Varchar2(20),
  �޸�����  Date,  
  ��������  Date,
  �ļ�����  Number (1),
  ��װ·��  Varchar2(250),
  ҵ�񲿼�  Varchar2(2000),
  ����ϵͳ  Varchar2(250),
  �ļ�˵��  Varchar2(2000),
  �Զ�ע��  Number (1),
  ǿ�Ƹ���	Number (1))
PCTFREE 5;
Alter Table zlTools.ZLFiles Add Constraint ZLFiles_PK Primary Key (����) USING INDEX PCTFREE 5;

Create Or Replace Procedure Zltools.Zlfiles_Autoupdate
(
  ����_In     In Zlfiles.����%Type,
  ��׼md5_In  In Zlfiles.��׼md5%Type,
  �汾��_In   In Zlfiles.�汾��%Type,
  �޸�����_In In Zlfiles.�޸�����%Type,
  ��������_In In Zlfiles.��������%Type,
  �ļ�����_In In Zlfiles.�ļ�����%Type,
  ��װ·��_In In Zlfiles.��װ·��%Type,
  ҵ�񲿼�_In In Zlfiles.ҵ�񲿼�%Type,
  ����ϵͳ_In In Zlfiles.����ϵͳ%Type,
  �ļ�˵��_In In Zlfiles.�ļ�˵��%Type,
  �Զ�ע��_In In Zlfiles.�Զ�ע��%Type,
  ǿ�Ƹ���_In In Zlfiles.ǿ�Ƹ���%Type
) Is
  n_Count Number(3);
Begin
  n_Count := 0;
  --��������
  For Rs In (Select Rowid From Zlfiles a Where Upper(a.����) = Upper(����_In)) Loop
    n_Count := n_Count + 1;
    Update Zlfiles
    Set ���� = ����_In, ��׼md5 = ��׼md5_In, �汾�� = �汾��_In, �޸����� = �޸�����_In, �������� = ��������_In, �ļ����� = �ļ�����_In, ��װ·�� = ��װ·��_In,
        ҵ�񲿼� = ҵ�񲿼�_In, ����ϵͳ = ����ϵͳ_In, �ļ�˵�� = �ļ�˵��_In, �Զ�ע�� = �Զ�ע��_In, ǿ�Ƹ��� = ǿ�Ƹ���_In
    Where Rowid = Rs.Rowid;
  End Loop;
  --��������
  If n_Count = 0 Then
    Insert Into Zlfiles
      (����, ��׼md5, �汾��, �޸�����, ��������, �ļ�����, ��װ·��, ҵ�񲿼�, ����ϵͳ, �ļ�˵��, �Զ�ע��, ǿ�Ƹ���)
    Values
      (����_In, ��׼md5_In, �汾��_In, �޸�����_In, ��������_In, �ļ�����_In, ��װ·��_In, ҵ�񲿼�_In, ����ϵͳ_In, �ļ�˵��_In, �Զ�ע��_In, ǿ�Ƹ���_In);
  End If;
  n_Count := 0;
  --��������
  For Rs In (Select Rowid From Zlfilesupgrade a Where Upper(a.�ļ���) = Upper(����_In)) Loop
    n_Count := n_Count + 1;
    Update Zlfilesupgrade
    Set �ļ��� = ����_In, �ļ����� = �ļ�����_In, ��װ·�� = ��װ·��_In, ҵ�񲿼� = ҵ�񲿼�_In, ����ϵͳ = ����ϵͳ_In, �ļ�˵�� = �ļ�˵��_In, �Զ�ע�� = �Զ�ע��_In,
        ǿ�Ƹ��� = ǿ�Ƹ���_In
    Where Rowid = Rs.Rowid;
  End Loop;
  --��������
  If n_Count = 0 Then
    Insert Into Zlfilesupgrade
      (�ļ���, �ļ�����, ��װ·��, ҵ�񲿼�, ����ϵͳ, �ļ�˵��, �Զ�ע��, ǿ�Ƹ���)
    Values
      (����_In, �ļ�����_In, ��װ·��_In, ҵ�񲿼�_In, ����ϵͳ_In, �ļ�˵��_In, �Զ�ע��_In, ǿ�Ƹ���_In);
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zlfiles_Autoupdate;
/
--00000:��˶,2016-06-30,������ǰ�İ汾�ȽϺ���
Create Or Replace Function Zltools.Zlverdiff
(
  Verpre_In  Varchar2,
  Vernext_In Varchar2
) Return Number
--���أ�1��ǰһ�汾��-1��ǰһ�汾С��0�������汾��ͬ
 As
  n_Pos   Number(2);
  v_Pre   Varchar2(20);
  v_Next  Varchar2(20);
  v_Temp1 Varchar2(20);
  v_Temp2 Varchar2(20);
Begin
  v_Pre  := Verpre_In;
  v_Next := Vernext_In;

  While v_Pre Is Not Null Loop
    n_Pos := Instr(v_Pre, '.');
    If n_Pos = 0 Then
      --û�ҵ���㣬�Ͱ������ַ�����Ϊ�汾
      v_Temp1 := v_Pre;
      v_Pre   := '';
    Else
      v_Temp1 := Substr(v_Pre, 1, n_Pos - 1);
      v_Pre   := Substr(v_Pre, n_Pos + 1);
    End If;
    n_Pos := Instr(v_Next, '.');
    If n_Pos = 0 Then
      v_Temp2 := v_Next;
      v_Next  := '';
    Else
      v_Temp2 := Substr(v_Next, 1, n_Pos - 1);
      v_Next  := Substr(v_Next, n_Pos + 1);
    End If;
  
    If To_Number(v_Temp1) > To_Number(v_Temp2) Then
      Return 1;
    Elsif To_Number(v_Temp1) < To_Number(v_Temp2) Then
      Return - 1;
    End If;
    --��ǰһ��ʣ���Ϊ�գ���һ��ʣ��β�Ϊ�գ���ǰһ����ֵΪ0�������Ƚ�
    If v_Pre Is Null And Not v_Next Is Null Then
      v_Pre := '0';
    End If;
  End Loop;
  Return 0;
End;
/
--00000:��˶,2016-06-30,���Ӱ汾�ȽϺ�����������
Create Or Replace Function Zltools.Zlcheck_Version_Upon
--���ܣ��жϴ���İ汾�Ƿ����û���ǰʹ�ð汾֮��
  --˵������ҪӦ�������������ű���
  --Sysno_In��ϵͳ���
  --Version_In����Ӱ�����Ͱ汾��
  --���أ�1=�û���ǰʹ�ð汾���ڻ����ָ���汾֮�ϡ�0-�û���ǰʹ�ð汾С��ָ���汾��
(
  Sysno_In   Zlsystems.���%Type,
  Version_In Zlsystems.�汾��%Type
) Return Number Is
  v_Startversion Varchar2(20);
  n_Count        Number(5);
Begin
  --��ȡ�û���ǰʹ�ð汾
  --1��û����Ǩ��¼�����ȡzlsysTem��¼�汾��
  --2������Ǩ��¼�������������֮�ڵ���Ǩ��¼�����ȡ��С����ʼ�汾��
  --3������Ǩ��¼���������������֮�ڵ���Ǩ��¼�����ȡ��һ����Ǩ��¼�Ľ���汾��
  Select Count(1) Into n_Count From Zlupgrade Where Nvl(ϵͳ, 0) = Sysno_In;
  --��ϵͳû����Ǩ��¼
  If n_Count = 0 Then
    Select �汾�� Into v_Startversion From Zlsystems Where ��� = Sysno_In;
  Else
    Select Max(ԭʼ�汾), Count(1)
    Into v_Startversion, n_Count
    From (Select ԭʼ�汾, ��Ǩʱ��
           From Zlupgrade
           Where Nvl(ϵͳ, 0) = Sysno_In And ��Ǩʱ�� Between Sysdate - 2 And Sysdate
           Order By ��Ǩʱ��)
    Where Rownum < 2;
    If n_Count = 0 Then
      Select Max(����汾), Count(1)
      Into v_Startversion, n_Count
      From (Select ����汾, ��Ǩʱ�� From Zlupgrade Where Nvl(ϵͳ, 0) = Sysno_In Order By ��Ǩʱ�� Desc)
      Where Rownum < 2;
    End If;
  End If;
  If Zlverdiff(v_Startversion, Version_In) >= 0 Then
    n_Count := 1;
  Else
    n_Count := 0;
  End If;
  Return n_Count;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End;
/
--94568:��˶,2016-05-14,������Ϣά��
ALTER TABLE  zltools.zlRegInfo ADD վ�� Varchar2(1);
ALTER TABLE zltools.zlRegInfo Drop CONSTRAINT zlRegInfo_UQ_��Ŀ Cascade Drop Index;
ALTER TABLE zltools.zlRegInfo ADD CONSTRAINT zlRegInfo_UQ_��Ŀ UNIQUE (��Ŀ,�к�,վ��) USING INDEX PCTFREE 5;

ALTER TABLE  zltools.Zlunitinfoimage ADD վ�� Varchar2(1);
ALTER TABLE zltools.zlUnitInfoImage Drop CONSTRAINT zlUnitInfoImage_PK Cascade Drop Index;
Alter Table zlTools.zlUnitInfoImage Add Constraint zlUnitInfoImage_UQ_��Ŀ UNIQUE (��Ŀ,վ��) USING INDEX PCTFREE 5;
Alter Table ZLTOOLS.zlUnitInfoImage Modify ��Ŀ  constraint zlUnitInfoImage_NN_��Ŀ   not  null;

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
    ����_In In Zlreginfo.����%Type,
    վ��_In In Zlreginfo.վ��%Type
  );
  --���ܣ�����ɾ�������Zlunitinfoitem
  --�����б�frmUnitInfoEdit,frmUnitItemEdit
  --����˵����
  --Type_n:0-������1-�޸�,2-ɾ��
  --����_In����Ŀ����
  --����_In:��Ŀ����,վ��
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
        Delete Zltools.Zlunitinfoimage
        Where ��Ŀ = Substr(Key_In, 1, Instr(Key_In, ',') - 1) And
              Nvl(վ��, '�տ�') = Nvl(Substr(Key_In, Instr(Key_In, ',') + 1), '�տ�');
      Else
        If Cls_In = 1 Then
          Update Zltools.Zlunitinfoimage
          Set ͼƬ = Empty_Blob()
          Where ��Ŀ = Substr(Key_In, 1, Instr(Key_In, ',') - 1) And
                Nvl(վ��, '�տ�') = Nvl(Substr(Key_In, Instr(Key_In, ',') + 1), '�տ�');
          If Sql%Rowcount = 0 Then
            Insert Into Zltools.Zlunitinfoimage
              (��Ŀ, ͼƬ, վ��)
            Values
              (Substr(Key_In, 1, Instr(Key_In, ',') - 1), Empty_Blob(), Substr(Key_In, Instr(Key_In, ',') + 1));
          End If;
        End If;
        Select ͼƬ
        Into l_Blob
        From Zltools.Zlunitinfoimage
        Where ��Ŀ = Substr(Key_In, 1, Instr(Key_In, ',') - 1) And
              Nvl(վ��, '�տ�') = Nvl(Substr(Key_In, Instr(Key_In, ',') + 1), '�տ�')
        For Update;
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
      Select ͼƬ
      Into l_Blob
      From Zltools.Zlunitinfoimage
      Where ��Ŀ = Substr(Key_In, 1, Instr(Key_In, ',') - 1) And
            Nvl(վ��, '�տ�') = Nvl(Substr(Key_In, Instr(Key_In, ',') + 1), '�տ�');
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
    ����_In In Zlreginfo.����%Type,
    վ��_In In Zlreginfo.վ��%Type
  ) Is
    --��Ŀ_In����Ŀ����
    --�к�_In:Ϊ�ջ���Ϊ1ʱ��ɾ������Ŀ��Ȼ�����
    --����_In:Ϊ���򲻲���
  Begin
    If Nvl(�к�_In, 0) < 2 Then
      Delete Zlreginfo Where ��Ŀ = ��Ŀ_In And Nvl(վ��, '�տ�') = Nvl(վ��_In, '�տ�');
    End If;
    If Not ����_In Is Null Then
      Insert Into Zlreginfo (��Ŀ, �к�, ����, վ��) Values (��Ŀ_In, �к�_In, ����_In, վ��_In);
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


--91515:������,2016-08-16,�������������������䴦��֧�ָ߲���ҵ��
alter index ZLTOOLS.ZLREPORTS_PK initrans 20;
alter index ZLTOOLS.ZLREPORTS_UQ_��� initrans 20;
alter index ZLTOOLS.ZLREPORTS_IX_����ID initrans 20;


--98644:������,2016-07-15,�����ս�ֹʱ�����ж�ÿ���ο�ת��������
alter table zltools.zlDataMove add ������������ date;
