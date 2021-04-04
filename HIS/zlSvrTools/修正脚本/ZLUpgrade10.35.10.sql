----10.35.0---》10.35.10
-----------------------------------------------------------------------------------------------------
-----------------------------------------------------------------------------------------------------
--93723:吴涛,2016-03-15,修改客户端自动升级程序，对不需要的文件进行删除
Create Table ZLTOOLS.ZLFILESEXPIRED
(
  文件名  VARCHAR2(50),
  安装路径 VARCHAR2(250),
  系统编号 NUMBER(5),
  系统版本 VARCHAR2(10),
  说明   VARCHAR2(250))
  PCTFREE 5 initrans 20;
INSERT INTO zltools.zlFilesExpired(文件名,安装路径,系统编号,系统版本,说明) VALUES('zlLogin.dll','[APPSOFT]\zlQueueShow',Null,'10.35.10','新增统一登录部件，弃用原来仅用于排队叫号的登录部件');

--88843:刘硕,2016-01-11,自定义存储过程管理
alter table zltools.ZLPROCEDURETEXT drop constraint ZLPROCEDURETEXT_FK_过程ID;
alter table zltools.ZLPROCEDURETEXT add constraint ZLPROCEDURETEXT_FK_过程ID foreign key (过程ID) references ZLPROCEDURE (ID)  on delete cascade;
alter  table zltools.ZLPROCEDURETEXT add constraint ZLPROCEDURETEXT_PK primary key(过程ID,性质,序号) using index;
Alter Table zltools.Zlprocedure Add Constraint Zlprocedure_Uq_名称 Unique(名称)Using Index Pctfree 5;
--91515:吴涛,2015-12-10,自定义报表SQL编辑器新增常用函数选择
CREATE TABLE Zltools.zlUsualFunc(
    系统编号 NUMBER(5),
    名称 VARCHAR2(50),
    说明 VARCHAR2(500))
    PCTFREE 5;
ALTER TABLE Zltools.zlUsualFunc ADD CONSTRAINT zlUsualFunc_PK UNIQUE (系统编号,名称)  USING INDEX;

CREATE Sequence Zltools.zlRPTRunHistory_ID start with 1;
CREATE TABLE Zltools.zlRPTRunHistory(
    ID       NUMBER(18),
    报表ID NUMBER(18),
    执行人员ID Number(18),
    执行开始时间 Date,
    执行结束时间 Date)
    PCTFREE 5 initrans 20;
ALTER TABLE Zltools.zlRPTRunHistory ADD CONSTRAINT zlRPTRunHistory_PK PRIMARY KEY (ID) USING INDEX;
ALTER TABLE Zltools.zlRPTRunHistory ADD CONSTRAINT zlRPTRunHistory_FK_报表ID FOREIGN KEY(报表ID) REFERENCES zlReports(ID) ON DELETE CASCADE;
CREATE INDEX zlRPTRunHistory_IX_报表ID   ON Zltools.zlRPTRunHistory(报表ID) PCTFREE 5;

Insert Into Zltools.zlUsualFunc(系统编号,名称,说明)
Select 0,'zlSpellCode','获取汉字字符串中的每个字的拼音首字母,缺省不超过10位' From Dual Union All
Select 0,'zlWBCode','获取汉字字符串中的每个字的五笔首字母,缺省不超过10位' From Dual Union All
Select 0,'Zlpinyincode','获取汉字字符中的每个字的拼音首字母或全拼,支持汉字多音字、全拼首字母大写和分隔符，缺省不超过10位' From Dual Union All
Select 0,'zlUppMoney','将数字金额转换为汉字大写金额字符串' From Dual Union All
Select 0,'zlUppNumber','将汉字大写金额字符串转换为数字金额' From Dual Union All
Select 0,'Zl_To_Number','将字符与数字混合的字符类型数据转换为去掉字符的数字类型数据（第2个参数传入1）' From Dual Union All
Select 0,'f_str2list','将由逗号分隔的不带引号的字符序列转换为单列数据表（列名为Column_Value,需配合Table函数来使用；如需两列，可用f_Str2list2） ' From Dual Union All
Select 0,'f_num2list','将由逗号分隔的不带引号的数字序列转换为单列数据表（列名为Column_Value,需配合Table函数来使用；如需两列，可用f_Num2list2） ' From Dual Union All
Select 0,'f_list2str','将单列多行的字符列表拼接为一个缺省以逗号分隔的字符串（需配合Collect函数来使用，可用来替代WM_CONCAT和sys_connect_by_path）。' From Dual Union All
Select 0,'zl_GetSysParameter','获取系统参数，或模块参数，或部门级参数的参数值，支持按参数号或参数名获取' From Dual Union All
Select 100,'Zl_Identity','获取当前用户的部门人员信息' From Dual Union All
Select 100,'Zl_Username','获取当前用户的姓名' From Dual Union All
Select 100,'zl_IncStr','获取字符串按Ascii递增后的字符串' From Dual Union All
Select 100,'Zl_Incstr_Pre','获取字符串按Ascii递减后的字符串' From Dual Union All
Select 100,'Zl_Age_Calc','根据出生日期计算病人年龄，返回包含年龄单位的字符串' From Dual Union All
Select 100,'ZL_AgeToDays','根据包含年龄单位的字符串，返回天数。' From Dual Union All
Select 100,'Zl_Cent_Money','根据分币处理规则的系统参数，返回经舍入处理后的金额数字' From Dual;


--91515:吴涛,2016-04-08,自定义报表关联报表支持关联多张报表
alter table Zltools.Zlrptrelation Add 默认 Number(1);

--94568:刘硕,2016-03-26,服务窗信息维护
CREATE TABLE zlTools.zlUnitInfoItem(
    编码 VARCHAR2(3),
    名称 VARCHAR2(20),
    是否图片 NUMBER(1))
    PCTFREE 5;
Alter Table zlTools.zlUnitInfoItem Add Constraint zlUnitInfoItem_PK Primary Key (编码) USING INDEX PCTFREE 5;
Alter Table zlTools.zlUnitInfoItem Add Constraint zlUnitInfoItem_UQ_名称 Unique (名称) USING INDEX PCTFREE 5;                 

CREATE TABLE zlTools.zlUnitInfoImage(   
    项目 VARCHAR2(20),
    图片 Blob)
    PCTFREE 5
    Cache Storage(Buffer_Pool Keep);
Alter Table zlTools.zlUnitInfoImage Add Constraint zlUnitInfoImage_PK Primary Key (项目) USING INDEX PCTFREE 5;
ALTER TABLE zlTools.zlUnitInfoImage ADD CONSTRAINT  zlUnitInfoImage_FK_项目 FOREIGN KEY (项目) REFERENCES zlTools.zlUnitInfoItem(名称) ON DELETE CASCADE;
alter Table ZLTOOLS.Zlsvrtools add 次序 number(3);
--94568:刘硕,2016-03-26,服务窗信息维护
Update Zlsvrtools
Set 次序 = 1 + (To_Number(Substr(编号, 3, 2)) - 1) * 3
Where 上级 Is Not Null And (上级 = '03' And 编号 = '0301' Or 上级 <> '03');
Update Zlsvrtools
Set 次序 = 1 + (To_Number(Substr(编号, 3, 2))) * 3
Where 上级 Is Not Null And 上级 = '03' And 编号 <> '0301';
Update Zlsvrtools Set 次序 = 4 Where 编号 = '0312';
Insert Into zlSvrTools(编号,上级,标题,快键,说明,次序) Values('0312','03','医院信息维护','H',Null,4);
insert into zlTools.zlUnitInfoItem(编码,名称,是否图片)values('001','医院介绍',0);
insert into zlTools.zlUnitInfoItem(编码,名称,是否图片)values('002','交通',0);
insert into zlTools.zlUnitInfoItem(编码,名称,是否图片)values('003','地址',0);
insert into zlTools.zlUnitInfoItem(编码,名称,是否图片)values('004','电话',0);
insert into zlTools.zlUnitInfoItem(编码,名称,是否图片)values('005','联系人',0);
insert into zlTools.zlUnitInfoItem(编码,名称,是否图片)values('006','负责人',0);
insert into zlTools.zlUnitInfoItem(编码,名称,是否图片)values('007','开户银行',0);
insert into zlTools.zlUnitInfoItem(编码,名称,是否图片)values('008','银行帐号',0);
insert into zlTools.zlUnitInfoItem(编码,名称,是否图片)values('009','税务号',0);
insert into zlTools.zlUnitInfoItem(编码,名称,是否图片)values('010','院长',0);
insert into zlTools.zlUnitInfoItem(编码,名称,是否图片)values('011','医院等级',0);
insert into zlTools.zlUnitInfoItem(编码,名称,是否图片)values('012','电子邮件',0);
insert into zlTools.zlUnitInfoItem(编码,名称,是否图片)values('013','医院代码',0);
insert into zlTools.zlUnitInfoItem(编码,名称,是否图片)values('014','主页',0);
--88843:刘硕,2016-04-14,自定义存储过程管理
Create Or Replace Procedure Zltools.Zl_Zlproceduretext_Move Is
Begin
  Delete From Zlproceduretext Where 性质 In (1, 2);
  Update Zlprocedure Set 状态 = 0;
  Insert Into Zlproceduretext
    (过程id, 性质, 序号, 内容)
    Select 过程id, Decode(性质, 3, 1, 4, 2), 序号, 内容 From Zlproceduretext Where 性质 In (3, 4);
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zlproceduretext_Move;
/

Create Or Replace Procedure Zltools.Zl_Zlprocedure_Confirm(Id_In Zlprocedure.Id%Type) Is
Begin
  Update Zlprocedure Set 状态 = 3 Where Id = Id_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zlprocedure_Confirm;
/
Create Or Replace Procedure Zltools.Zl_Zlprocedure_Manage(Nstep_In Number := 0) Is
  --Nstep_In 0-收集前调用，此时将所有过程变更为待检查状态。
  --         1-检查后待用，此时将状态仍旧处于待检查的用户过程调整为待调整，自待检查的定义过程与变动过程变更为无变化
Begin
  If Nvl(Nstep_In, 0) = 0 Then
    Update Zlprocedure Set 状态 = 0;
  Else
    Update Zlprocedure Set 状态 = 1 Where 类型 = 3 And 状态 = 0;
    Update Zlprocedure Set 状态 = 4 Where 类型 In (1, 2) And 状态 = 0;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zlprocedure_Manage;
/
--94568:刘硕,2016-03-26,服务窗信息维护
Create Or Replace Package Zltools.b_Public Is
  --公共过程
  Type t_Refcur Is Ref Cursor;
  --功能：取系统日期
  --调用列表：mdlMain.CurrentDate，clsDatabase.CurrentDate
  Procedure Get_Current_Date(Cursor_Out Out t_Refcur);
  --功能：删除错误日志或运行日志
  --调用列表：mdlMain.DeleteAllLog
  Procedure Delete_All_Log(Runtimelog_In In Number := 0);
  --功能：删除当前运行日志
  --调用列表：mdlMain.DeleteCurLog
  Procedure Delete_Diarylog
  (
    会话号_In   Number,
    用户名_In   Varchar2,
    工作站_In   Varchar2,
    部件名_In   Varchar2,
    工作内容_In Varchar2,
    进入时间_In Date
  );
  --功能：删除当前错误日志
  --调用列表：mdlMain.DeleteCurLog
  Procedure Delete_Errorlog
  (
    会话号_In   Number,
    用户名_In   Varchar2,
    工作站_In   Varchar2,
    类型_In     Number,
    错误序号_In Number,
    时间_In     Date
  );
  --功能：取注册码
  --调用列表：mdlMain.Get注册码
  Procedure Get_Regcode(Cursor_Out Out t_Refcur);
  --功能：取版本号
  --调用列表：mdlMain.UpgradeManager
  Procedure Get_Ver(Cursor_Out Out t_Refcur);
  --功能：更新版本号
  --调用列表：mdlMain.UpgradeManager
  Procedure Update_Ver(Verstring_In In Varchar2);
  --功能：取得系统所有者名称
  --调用列表：
  --frmStatus.cmbsystem_Click、mdlMain.GetOwnerName、mdlMain.cmbSystem_Click
  --frmAutoJobs.cmbSystem_Click、frmDataMove.cmbSystem_Click 、frmNoticeTools.cboSystem_Click
  --frmProgPriv.ProgPriv、frmAppScript.cmbSystem_Click
  Procedure Get_Owner_Name
  (
    Cursor_Out Out t_Refcur,
    编号_In    In Zlsystems.编号%Type := 0
  );

  --功能：取注册表中信息
  --调用列表：
  --frmAbout.GetUnitInfo、frmAutoJobs.From_load、frmClientsUpgrade.InitInfor
  --frmFilesSet.ShowEdit、frmRegist.From_load、frmAppScript.From_Load
  --frmFilesSendToServer.InitInfo
  Procedure Get_Reginfo
  (
    Cursor_Out Out t_Refcur,
    项目_In    In Zlreginfo.项目%Type := Null
  );
  --功能：取zlGetSvrToolsg数据
  --调用列表：frmMDIMain.MDIForm_Load
  Procedure Get_Zlsvrtools(Cursor_Out Out t_Refcur);
  --功能：取已安装系统清单
  --调用列表：
  --frmAppCheck.Form_Load、frmClearData.Form_Load、frmDataMove.Form_Load
  --frmImp.FillSystem、frmLoadIn.FillSystem、frmLoadOut.FillSystem
  --frmMDIMain.mnuFileRemove_Click、frmNoticeTools.Form_Activate、frmRoleGrant.FillSystem
  --frmAppUpgrade.Form_Load、frmAppScript.Form_Load、frmExp.FillSystem
  --frmInputTools.from_activate、fromRole.FillSystem、frmAutoJobs.From_load
  --frmAppstart.sysCreated
  Procedure Get_Zlsystems
  (
    Cursor_Out Out t_Refcur,
    所有者_In  In Zlsystems.所有者%Type := Null
  );
  --功能：存储BLOb图片
  --调用列表：frmUnitInfoEdit
  --参数说明：
  --Tab_In：包含LOB的数据表
  --        0-zlUnitInfoImage
  --Key_In：数据记录的关键字
  --Txt_In：16进制的文件片段或文字片段
  --Cls_In：是否清除原来的内容，第一片段传递时为1，以后为0
  Procedure Zllobappend
  (
    Tab_In In Number,
    Key_In In Varchar2,
    Txt_In In Varchar2,
    Cls_In In Number := 0
  );
  --功能：读取BLOb图片
  --调用列表：frmUnitInfoEdit
  --参数说明：
  --Tab_In：包含LOB的数据表
  --        0-zlUnitInfoImage
  --Key_In：数据记录的关键字
  --Pos_In：从0开始不断读取，直到返回为空
  Function Zllobread
  (
    Tab_In In Number,
    Key_In In Varchar2,
    Pos_In In Number
  ) Return Varchar2;

  --功能：插入删除或更新ZLRegInfo图片
  --调用列表：frmUnitInfoEdit
  --参数说明：
  --项目_In：项目名称
  --行号_In:为空或者为1时先删除该项目，然后插入
  --内容_In:为空则不插入
  Procedure Zlreginfoupdate
  (
    项目_In In Zlreginfo.项目%Type,
    行号_In In Zlreginfo.行号%Type,
    内容_In In Zlreginfo.内容%Type
  );
  --功能：插入删除或更新Zlunitinfoitem
  --调用列表：frmUnitInfoEdit,frmUnitItemEdit
  --参数说明：
  --Type_n:0-新增，1-修改,2-删除
  --编码_In：项目编码
  --名称_In:项目名称
  --图片_In:项目是否是图片类型
  Procedure Zlunitinfoitemchange
  (
    Type_n  In Number,
    编码_In In Zlunitinfoitem.编码%Type,
    名称_In In Zlunitinfoitem.名称%Type := Null,
    图片_In In Zlunitinfoitem.是否图片%Type := Null
  );
End b_Public;
/

--94568:刘硕,2016-03-26,服务窗信息维护
Create Or Replace Package Body Zltools.b_Public Is
  --功能：取系统日期
  Procedure Get_Current_Date(Cursor_Out Out t_Refcur) Is
  Begin
    Open Cursor_Out For
      Select Sysdate As 日期 From Dual;
  End Get_Current_Date;

  --功能：删除错误日志或运行日志
  Procedure Delete_All_Log(Runtimelog_In In Number := 0) Is
    n_Count Number;
    n_Loop  Number;
  Begin
    If Runtimelog_In = 1 Then
      Select Count(进入时间) Into n_Count From Zldiarylog;
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
      Select Count(时间) Into n_Count From Zlerrorlog;
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

  --功能：删除当前运行日志
  Procedure Delete_Diarylog
  (
    会话号_In   Number,
    用户名_In   Varchar2,
    工作站_In   Varchar2,
    部件名_In   Varchar2,
    工作内容_In Varchar2,
    进入时间_In Date
  ) Is
  Begin
    Delete Zldiarylog
    Where 会话号 = 会话号_In And 用户名 = 用户名_In And 工作站 = 工作站_In And 部件名 = 部件名_In And 工作内容 = 工作内容_In And 进入时间 = 进入时间_In;
    Commit;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Delete_Diarylog;

  --功能：删除当前错误日志
  Procedure Delete_Errorlog
  (
    会话号_In   Number,
    用户名_In   Varchar2,
    工作站_In   Varchar2,
    类型_In     Number,
    错误序号_In Number,
    时间_In     Date
  ) Is
  Begin
    Delete Zlerrorlog
    Where 会话号 = 会话号_In And 用户名 = 用户名_In And 工作站 = 工作站_In And 类型 = 类型_In And 错误序号 = 错误序号_In And 时间 = 时间_In;
    Commit;
  
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Delete_Errorlog;

  --功能：取注册码
  Procedure Get_Regcode(Cursor_Out Out t_Refcur) Is
  Begin
    Open Cursor_Out For
      Select 内容 From Zlreginfo Where 项目 = '注册码' Or 项目 = '授权证章' Order By 行号;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Regcode;

  --功能：取版本号
  Procedure Get_Ver(Cursor_Out Out t_Refcur) Is
  Begin
    Open Cursor_Out For
      Select 内容 From Zlreginfo Where 项目 = '版本号';
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Ver;

  --功能：更新版本号
  Procedure Update_Ver(Verstring_In In Varchar2) Is
  Begin
    Update Zlreginfo Set 内容 = Verstring_In Where 项目 = '版本号';
    If Sql%Notfound Then
      Insert Into Zlreginfo (项目, 行号, 内容) Values ('版本号', 1, Verstring_In);
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Update_Ver;

  --功能：取得系统所有者名称
  Procedure Get_Owner_Name
  (
    Cursor_Out Out t_Refcur,
    编号_In    In Zlsystems.编号%Type := 0
  ) Is
  Begin
    Open Cursor_Out For
      Select Upper(所有者) As 所有者 From Zlsystems Where 编号 = 编号_In;
  End Get_Owner_Name;

  --功能：取注册表中信息
  Procedure Get_Reginfo
  (
    Cursor_Out Out t_Refcur,
    项目_In    In Zlreginfo.项目%Type := Null
  ) Is
  Begin
    If Trim(Nvl(项目_In, '空')) = '空' Then
      Open Cursor_Out For
        Select * From Zlreginfo;
    Else
      Open Cursor_Out For
        Select 内容 From Zlreginfo Where 项目 = 项目_In Order By 行号;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Reginfo;

  --功能：取zlGetSvrToolsg数据
  Procedure Get_Zlsvrtools(Cursor_Out Out t_Refcur) Is
  Begin
    Open Cursor_Out For
      Select * From Zlsvrtools Start With 上级 Is Null Connect By Prior 编号 = 上级 Order By Level, 编号;
  End Get_Zlsvrtools;

  --功能：取已安装系统清单
  Procedure Get_Zlsystems
  (
    Cursor_Out Out t_Refcur,
    所有者_In  In Zlsystems.所有者%Type := Null
  ) Is
  Begin
    If Nvl(所有者_In, '空') = '空' Then
      Open Cursor_Out For
        Select 编号, 名称, 共享号, Upper(所有者) 所有者, 安装日期, 正常安装, 版本号 From Zlsystems Order By 编号;
    Else
      Open Cursor_Out For
        Select 编号, 名称, 共享号, Upper(所有者) 所有者, 安装日期, 正常安装, 版本号
        From Zlsystems
        Where Upper(所有者) = Upper(所有者_In)
        Order By 编号;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Zlsystems;
  --功能：存储BLOb图片
  Procedure Zllobappend
  (
    Tab_In In Number,
    Key_In In Varchar2,
    Txt_In In Varchar2,
    Cls_In In Number := 0
    --参数说明： 
    --Tab_In：包含LOB的数据表 
    --        0-zlUnitInfoImage
    --Key_In：数据记录的关键字 
    --Txt_In：16进制的文件片段或文字片段 
    --Cls_In：是否清除原来的内容，第一片段传递时为1，以后为0 
  ) Is
    l_Blob Blob;
  Begin
    If Tab_In = 0 Then
      If Txt_In Is Null And Cls_In = 1 Then
        Delete Zltools.Zlunitinfoimage Where 项目 = Key_In;
      Else
        If Cls_In = 1 Then
          Update Zltools.Zlunitinfoimage Set 图片 = Empty_Blob() Where 项目 = Key_In;
          If Sql%Rowcount = 0 Then
            Insert Into Zltools.Zlunitinfoimage (项目, 图片) Values (Key_In, Empty_Blob());
          End If;
        End If;
        Select 图片 Into l_Blob From Zltools.Zlunitinfoimage Where 项目 = Key_In For Update;
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
  --功能：读取BLOb图片
  Function Zllobread
  (
    Tab_In In Number,
    Key_In In Varchar2,
    Pos_In In Number --参数说明： 
    --Tab_In：包含LOB的数据表 
    --        0-zlUnitInfoImage
    --Key_In：数据记录的关键字 
    --Pos_In：从0开始不断读取，直到返回为空 
  ) Return Varchar2 Is
    l_Blob   Blob;
    v_Buffer Varchar2(32767);
    n_Amount Number := 2000;
    n_Offset Number := 1;
  Begin
    If Tab_In = 0 Then
      Select 图片 Into l_Blob From Zltools.Zlunitinfoimage Where 项目 = Key_In;
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
    项目_In In Zlreginfo.项目%Type,
    行号_In In Zlreginfo.行号%Type,
    内容_In In Zlreginfo.内容%Type
  ) Is
    --项目_In：项目名称
    --行号_In:为空或者为1时先删除该项目，然后插入
    --内容_In:为空则不插入
  Begin
    If Nvl(行号_In, 0) < 2 Then
      Delete Zlreginfo Where 项目 = 项目_In;
    End If;
    If Not 内容_In Is Null Then
      Insert Into Zlreginfo (项目, 行号, 内容) Values (项目_In, 行号_In, 内容_In);
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Zlreginfoupdate;
  Procedure Zlunitinfoitemchange
  (
    Type_n  In Number,
    编码_In In Zlunitinfoitem.编码%Type,
    名称_In In Zlunitinfoitem.名称%Type := Null,
    图片_In In Zlunitinfoitem.是否图片%Type := Null
  ) Is
    --Type_n:0-新增，1-修改,2-删除
    --编码_In：项目编码
    --名称_In:项目名称
    --图片_In:项目是否是图片类型
    v_名称 Zlunitinfoitem.名称%Type;
    n_图片 Zlunitinfoitem.是否图片%Type;
  Begin
    If Type_n = 0 Then
      Insert Into Zlunitinfoitem (编码, 名称, 是否图片) Values (编码_In, 名称_In, 图片_In);
    Elsif Type_n = 1 Then
      Select Nvl(名称, 0), Nvl(是否图片, 0) Into v_名称, n_图片 From Zlunitinfoitem Where 编码 = 编码_In;
      --存在该项目
      If Not n_图片 Is Null Then
        --类型变更，删除所有数据
        If n_图片 <> Nvl(图片_In, 0) Then
          If n_图片 = 0 Then
            Delete Zlreginfo Where 项目 = v_名称;
          Else
            Delete Zlunitinfoimage Where 项目 = v_名称;
          End If;
          --名称变更
        Elsif v_名称 <> Nvl(名称_In, '空空') Then
          If n_图片 = 0 Then
            Update Zlreginfo Set 项目 = 名称_In Where 项目 = v_名称;
          Else
            Update Zlunitinfoimage Set 项目 = 名称_In Where 项目 = v_名称;
          End If;
        End If;
        Update Zlunitinfoitem Set 名称 = 名称_In, 是否图片 = 图片_In Where 编码 = 编码_In;
      End If;
    Else
      Select Nvl(名称, 0), Nvl(是否图片, 0) Into v_名称, n_图片 From Zlunitinfoitem Where 编码 = 编码_In;
      --存在该项目
      If Not n_图片 Is Null Then
        If n_图片 = 0 Then
          Delete Zlreginfo Where 项目 = v_名称;
        Else
          Delete Zlunitinfoimage Where 项目 = v_名称;
        End If;
        Delete Zlunitinfoitem Where 编码 = 编码_In;
      End If;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Zlunitinfoitemchange;
End b_Public;
/
--92592:刘硕,2015-01-18,记录所有用户角色
Delete Zltools.Zluserroles;
Insert Into Zltools.Zluserroles
(用户, 角色, 管理)
Select Grantee, Granted_Role, Decode(Admin_Option, 'YES', 1, 0)
From Dba_Role_Privs
Where Granted_Role Like 'ZL_%'
Order By Grantee;
--92592:刘硕,2015-01-18,记录所有用户角色
Create Or Replace Procedure Zltools.Zl_Zluserroles_Add
(
  User_In In Zluserroles.用户%Type := Null,
  Role_In In Zluserroles.角色%Type := Null,
  管理_In In Zluserroles.管理%Type := 0
) Is
  --当用户角色均为空时，记录所有用户角色数据
Begin
  If Not User_In Is Null And Not Role_In Is Null Then
    Insert Into Zluserroles (用户, 角色, 管理) Values (User_In, Role_In, 管理_In);
    --用户角色都不传时，清空所有数据，并重新生成
  Elsif User_In Is Null And Role_In Is Null Then
    Delete Zltools.Zluserroles;
    Insert Into Zltools.Zluserroles
      (用户, 角色, 管理)
      Select Grantee, Granted_Role, Decode(Admin_Option, 'YES', 1, 0)
      From Dba_Role_Privs
      Where Granted_Role Like 'ZL_%'
      Order By Grantee;
    --用户传入时，清空该用户，并重新生成数据
  Elsif Not User_In Is Null Then
    Delete Zltools.Zluserroles Where 用户 = User_In;
    Insert Into Zltools.Zluserroles
      (用户, 角色, 管理)
      Select Grantee, Granted_Role, Decode(Admin_Option, 'YES', 1, 0)
      From Dba_Role_Privs
      Where Granted_Role Like 'ZL_%' And Grantee = User_In;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zluserroles_Add;
/
--92592:刘硕,2015-01-18,记录所有用户角色
Create Or Replace Procedure Zltools.Zl_Zluserroles_Del
(
  User_In In Zluserroles.用户%Type,
  Role_In In Zluserroles.角色%Type
) Is
Begin
  If Not User_In Is Null And Not Role_In Is Null Then
    Delete Zluserroles Where 用户 = User_In And 角色 = Role_In;
  Elsif Not User_In Is Null Then
    Delete Zluserroles Where 用户 = User_In;
  Elsif Not Role_In Is Null Then
    Delete Zluserroles Where 角色 = Role_In;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zluserroles_Del;
/

--91515:吴涛,2015-12-18,记录报表执行执行时间
alter table Zltools.Zlreports Add 执行人员ID Number(18);
alter table Zltools.Zlreports Add 执行开始时间 Date;
alter table Zltools.Zlreports Add 执行结束时间 Date;
alter table Zltools.zlReports initrans 20;
alter table Zltools.ZlRptitems Add 表格线加粗 Number(1);
ALTER TABLE zltools.zlRPTITems ADD CONSTRAINT zlRPTItems_CK_表格线加粗 Check(表格线加粗 IN(1,0));

Insert Into zlParameters(ID,系统,模块,私有,本机,授权,固定,部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明)
Select zlParameters_ID.Nextval,-Null,-Null,-Null,-Null,-Null,-Null,A.* From (
Select 部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明 From zlParameters Where 1 = 0 Union All
select 0,0,26,'开启报表运行日志','0','0','监控报表执行人，执行开始时间和执行结束时间','0-关闭，1-开启',NULL,NULL,NULL From Dual Union All
select 0,0,27,'检查中型表','','3000,1000000','检查中型表记录数范围','为空则不检查，不为空则根据范围检查中型表',NULL,NULL,NULL From Dual Union All
select 0,0,28,'记录报表使用痕迹','1','1','在打开报表、刷新、重选格式、重置等操作完成后记录最后一次的执行人、执行时间','0-关闭，1-开启',NULL,NULL,NULL From Dual Union All
Select 部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明 From ZLPARAMETERS Where 1 = 0) A;

Create Or Replace Procedure Zltools.Zl_Rptrun_Update
(
  Id_In           In Zlreports.Id%Type,
  执行人员id_In   In Zlreports.执行人员id%Type,
  执行开始时间_In In Zlreports.执行开始时间%Type,
  执行结束时间_In In Zlreports.执行结束时间%Type
) Is
Begin
  Update Zlreports
  Set 执行人员id = 执行人员id_In, 执行开始时间 = 执行开始时间_In, 执行结束时间 = 执行结束时间_In
  Where Id = Id_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Rptrun_Update;
/

Create Or Replace Procedure Zltools.Zl_Rptrunhistory_Update
(
  Id_In           In Zlrptrunhistory.Id%Type,
  报表id_In       In Zlreports.Id%Type,
  执行人员id_In   In Zlrptrunhistory.执行人员id%Type,
  执行开始时间_In In Zlrptrunhistory.执行开始时间%Type,
  执行结束时间_In In Zlrptrunhistory.执行结束时间%Type
) Is
Begin
  Insert Into Zlrptrunhistory
    (Id, 报表id, 执行人员id, 执行开始时间, 执行结束时间)
  Values
    (Id_In, 报表id_In, 执行人员id_In, 执行开始时间_In, 执行结束时间_In);
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Rptrunhistory_Update;
/

--00000:刘硕,2015-12-25,表的时候支持表名不传只传对象名
Create Or Replace Function Zltools.Zl_Checkobject
(
  n_Type        In Number, --1=表,2=字段,3=约束,4=索引
  v_Object_Name In Varchar2,
  v_Table_Name  In Varchar2 := Null --仅当n_Type=2时才需要传入
) Return Number Authid Current_User As
  --功能：以执行者的身份检查指定表的指定对象是否存在
  --返回值：>0表示存在，0表示不存在
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
--88843:刘硕,2016-01-11,自定义存储过程管理
Create Or Replace Procedure Zltools.Zl_Zlproceduretext_Update
(
  过程id_In In Zlproceduretext.过程id%Type,
  性质_In   In Zlproceduretext.性质%Type,
  序号_In   In Zlproceduretext.序号%Type,
  内容_In   In Zlproceduretext.内容%Type
) Is
Begin
  --由于过程存在分段存储情况因此必须先删除后插入
  If Nvl(序号_In, 1) = 1 Then
    Delete Zlproceduretext Where 过程id = 过程id_In And 性质 = 性质_In;
  End If;
  If Not 内容_In Is Null Then
    Insert Into Zlproceduretext (过程id, 性质, 序号, 内容) Values (过程id_In, 性质_In, 序号_In, 内容_In);
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zlproceduretext_Update;
/
--92669:张永康,2016-01-13,注册码安全控制及查询性能提升
Create Or Replace Function zltools.f_Str2list
(
  Str_In   In Varchar2,
  Split_In In Varchar2 := ','
) Return t_Strlist
  Pipelined As
  v_Str Long;
  P     Number;
  --功能：将由逗号分隔的不带引号的字符序列转换为单列数据表 
  --参数：STR_IN,如:G0000123,G0000124,G0000125...,SPLIT_IN,分隔符,缺省为,号 
  --说明： 
  --1．当SQL语句中涉及“IN(常量1, 常量2,…) ”子句时使用这种方式以便利用绑定变量。 
  --2．使用这两个函数时，需要在SQL语句中加入“/*+ cardinality(b 3)*/”提示，因为CBO下临时内存表没有统计数据,。 
  --3．两种调用示例 
  --SELECT /*+ cardinality(b 3)*/ * FROM 门诊费用记录 WHERE NO IN (SELECT * FROM TABLE(F_STR2LIST('A01,A02,A03')) B); 
  --SELECT /*+ cardinality(b 3)*/ A.* FROM 门诊费用记录 A, TABLE(F_STR2LIST('A01,A02,A03')) B WHERE A.NO = B.COLUMN_VALUE; 
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

--92669:张永康,2016-01-13,注册码安全控制及查询性能提升
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


--92669:张永康,2016-01-13,注册码安全控制及查询性能提升
Create Or Replace Function zltools.f_Reg_Menu
(
  Menu_Group_In  In Zlmenus.组别%Type := Null, --本机选择的菜单组别
  System_List_In In Varchar2, --本次会话涉及的应用系统
  Part_List_In   In Varchar2 --以逗号分隔的本机可执行部件列表
) Return t_Menu_Rowset Is
  t_Return t_Menu_Rowset := t_Menu_Rowset();
  t_Middle t_Menu_Rowset := t_Menu_Rowset();

  v_Parts   Varchar2(32767);
  t_Parts   t_Reg_Rowset := t_Reg_Rowset();
  v_Systems Varchar2(32767);
  t_Systems t_Reg_Rowset := t_Reg_Rowset();

Begin
  --变量解析形成类型数组表
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

  --菜单数据获取：
  Select t_Menu_Record(m.层次, m.Id, m.上级id, m.标题, m.短标题, m.快键, m.说明, m.模块, m.系统, m.图标, p.部件, 0)
  Bulk Collect
  Into t_Middle
  From (Select Level As 层次, ID, 上级id, 标题, 短标题, 快键, 说明, 模块, 系统, 图标
         From zlMenus
         Where 组别 = Menu_Group_In
         Start With 上级id Is Null
         Connect By Prior ID = 上级id) M,
       (Select /*+ cardinality(C 20) cardinality(S 2)*/
         Distinct p.系统, p.序号, p.部件
         From zlPrograms P, zlProgFuncs F, zlRegFunc R, zlRPTGroups X, Table(Cast(t_Parts As t_Reg_Rowset)) C,
              Table(Cast(t_Systems As t_Reg_Rowset)) S,
              (Select Decode(Count(*), 0, 0, 1) As 编号
                From zlSystems
                Where 所有者 = User
                Union All
                Select 编号
                From zlSystems
                Where 所有者 = User) O,
              (Select Distinct g.系统, g.序号 From zlRoleGrant G, zlUserRoles R Where g.角色 = r.角色 And r.用户 = User) G
         Where Nvl(f.系统, 0) = Nvl(p.系统, 0) And f.序号 = p.序号 And Trunc(f.系统 / 100) = r.系统(+) And f.序号 = r.序号(+) And
               f.功能 = r.功能(+) And
               (r.功能 Is Null And f.系统 Is Null Or r.功能 Is Not Null And r.功能 = '基本' Or
                r.功能 Is Not Null And x.程序id Is Not Null Or r.功能 Is Null And (p.序号 Between 10000 And 19999)) And
               p.系统 = x.系统(+) And p.序号 = x.程序id(+) And Upper(p.部件) = c.Text And Nvl(p.系统, 0) = s.Prog And
               Nvl(p.系统, 1) = o.编号(+) And Nvl(p.系统, 0) = Nvl(g.系统(+), 0) And p.序号 = g.序号(+) And
               (o.编号 Is Not Null Or g.序号 Is Not Null)) P
  Where Nvl(m.系统, 0) = Nvl(p.系统(+), 0) And m.模块 = p.序号(+) And (m.模块 Is Null Or m.模块 Is Not Null And p.序号 Is Not Null)
  Order By m.层次 Desc;

  --清理无下级可执行的菜单项目
  For n_Child In 1 .. t_Middle.Count Loop
    If t_Middle(n_Child).部件 Is Not Null Or t_Middle(n_Child).标记 = 1 Then
      t_Return.Extend;
      t_Return(t_Return.Count) := t_Middle(n_Child);
      If t_Middle(n_Child).上级id Is Not Null Then
        For n_Parent In n_Child + 1 .. t_Middle.Count Loop
          If t_Middle(n_Parent).标记 = 0 And t_Middle(n_Parent).Id = t_Middle(n_Child).上级id Then
            t_Middle(n_Parent).标记 := 1;
            Exit;
          End If;
        End Loop;
      End If;
    End If;
  End Loop;

  Return t_Return;
End f_Reg_Menu;
/

--93131:张永康,2016-01-29,自动升级性能优化
Create Or Replace Procedure zltools.Zl_Zlclients_Control
(
  n_Mode_In       Number,
  v_工作站_In     Zlclients.工作站%Type := Null,
  v_Ip_In         Zlclients.Ip%Type := Null,
  n_升级标志_In   Zlclients.升级标志%Type := Null,
  n_升级服务器_In Zlclients.升级服务器%Type := Null,
  d_预升时点_In   Zlclients.预升时点%Type := Null,
  n_预升完成_In   Zlclients.预升完成%Type := Null,
  n_Ftp服务器_In  Zlclients.Ftp服务器%Type := Null,
  n_收集标志_In   Zlclients.收集标志%Type := Null,
  n_禁止使用_In   Zlclients.禁止使用%Type := Null,
  v_说明_In       Zlclients.说明%Type := Null
  --对客户端进行控制
  --N_Mode_In：0-禁用或启用客户端(IP做为主要条件）,1-预升级设置,2 -升级信息保存(IP做为主要条件）
  --3-取消预升级标志,4-将所有站点设置为升级,5-部件搜集（设置搜集标志）,6-重置升级状态
) Is
  v_Timeset Varchar2(300);
  v_Err     Varchar2(500);
  Err_Custom Exception;
Begin
  --0-禁用或启用客户端(IP做为主要条件）
  If n_Mode_In = 0 Then
    If v_工作站_In Is Not Null Then
      Update zlClients Set 禁止使用 = n_禁止使用_In Where Ip = v_Ip_In;
    End If;
    --1-预升级设置,不需要传其他参数
  Elsif n_Mode_In = 1 Then
    Select Max(内容) Into v_Timeset From zlRegInfo Where 项目 = '客户端预升级时间点';
    If v_Timeset Is Not Null Then
      For r_Ip In (Select To_Date(Today || ' ' || Date_d, 'yyyy-mm-dd HH24:mi:ss') 预升时点, 工作站, Ip
                   From (Select 工作站, Ip, Rownum Rn_c From zlClients) A,
                        (Select To_Char(Sysdate, 'yyyy-mm-dd') Today, Column_Value Date_d, Rownum Rn_d, Count(1) Over() Sn
                          From Table(f_Str2list(v_Timeset, ','))) B
                   Where Mod(a.Rn_c, Sn) + 1 = Rn_d) Loop
      
        Update zlClients Set 预升时点 = r_Ip.预升时点 Where 工作站 = r_Ip.工作站 And Ip = r_Ip.Ip;
      End Loop;
    Else
      v_Err := '你尚未进行客户端预升级时间点设置！';
      Raise Err_Custom;
    End If;
    --2 -升级信息保存(IP做为主要条件）
  Elsif n_Mode_In = 2 Then
    If n_Ftp服务器_In Is Null Then
      Update zlClients
      Set 升级标志 = n_升级标志_In, 升级服务器 = n_升级服务器_In, 预升时点 = d_预升时点_In, 预升完成 = n_预升完成_In
      Where Ip = v_Ip_In;
    
    Else
      Update zlClients
      Set 升级标志 = n_升级标志_In, Ftp服务器 = n_Ftp服务器_In, 预升时点 = d_预升时点_In, 预升完成 = n_预升完成_In
      Where Ip = v_Ip_In;
    End If;
    --3-取消预升级标志
  Elsif n_Mode_In = 3 Then
    Update zlClients Set 预升完成 = n_预升完成_In;
    --4-将所有站点设置为升级
  Elsif n_Mode_In = 4 Then
    Update zlClients Set 升级标志 = n_升级标志_In;
    --5-部件搜集（设置搜集标志）
  Elsif n_Mode_In = 5 Then
    If v_工作站_In Is Null Then
      Update zlClients Set 收集标志 = n_收集标志_In;
    Else
      Update zlClients Set 收集标志 = n_收集标志_In Where 工作站 = v_工作站_In;
    End If;
  Elsif n_Mode_In = 6 Then
    Update zlClients Set 升级情况 = 0 Where 工作站 = v_工作站_In;
  Elsif n_Mode_In = 7 Then
    --7未升级
    Update zlClients Set 升级情况 = 1 Where 工作站 = v_工作站_In;
  Elsif n_Mode_In = 8 Then
    --8已升级
    Update zlClients Set 升级情况 = 2 Where 工作站 = v_工作站_In;
  Elsif n_Mode_In = 9 Then
    --9修改说明
    Update zlClients Set 说明 = v_说明_In Where 工作站 = v_工作站_In;
  Elsif n_Mode_In = 10 Then
    --10修改说明和收集标志
    Update zlClients Set 说明 = v_说明_In, 收集标志 = 0 Where 工作站 = v_工作站_In;
  Elsif n_Mode_In = 11 Then
    --11修改说明和升级标志
    Update zlClients Set 说明 = v_说明_In, 升级标志 = 0 Where 工作站 = v_工作站_In;
  Elsif n_Mode_In = 12 Then
    --12更改站点的预升级出错状态
    Update zlClients Set 说明 = v_说明_In, 预升完成 = 0 Where 工作站 = v_工作站_In;
  Elsif n_Mode_In = 13 Then
    --13更改站点的预升级完成状态
    Update zlClients Set 预升完成 = 1 Where 工作站 = v_工作站_In;
  Elsif n_Mode_In = 14 Then
    --14定时升级
    Update zlClients Set 预升时点 = Null, 预升完成 = Null Where 工作站 = v_工作站_In;
  Elsif n_Mode_In = 15 Then
    --15升级成功
    Update zlClients Set 升级情况 = 1 Where 工作站 = v_工作站_In;
  Elsif n_Mode_In = 16 Then
    --16升级失败
    Update zlClients Set 升级情况 = 2 Where 工作站 = v_工作站_In;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlclients_Control;
/

Drop Procedure zltools.Zl_Zlclients_Upgrade;