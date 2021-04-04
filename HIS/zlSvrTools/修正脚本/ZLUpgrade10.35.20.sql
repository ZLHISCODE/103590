----10.35.10---》10.35.20
-----------------------------------------------------------------------------------------------------
-----------------------------------------------------------------------------------------------------
--98732:刘硕,2016-08-13,特殊SP版本支持
alter table zltools.ZLFILESEXPIRED Add 是否固定 Number(1); 
alter table zltools.ZLSYSTEMS modify 版本号 varchar2(20);
alter table zltools.ZLFILESEXPIRED modify 系统版本 varchar2(20);
alter table zltools.ZLUPGRADE modify 目标版本 varchar2(20);
alter table zltools.ZLUPGRADE modify 结果版本 varchar2(20);
alter table zltools.ZLUPGRADE modify 原始版本 varchar2(20);
alter table zltools.ZLCOMPONENT modify 注册产品版本 varchar2(20);
alter table zltools.ZLUPGRADE modify 目标版本 varchar2(20);

create table zltools.ZLFiles(
  名称  Varchar2(50),
  标准MD5  Varchar2(32),
  版本号  Varchar2(20),
  修改日期  Date,  
  加入日期  Date,
  文件类型  Number (1),
  安装路径  Varchar2(250),
  业务部件  Varchar2(2000),
  所属系统  Varchar2(250),
  文件说明  Varchar2(2000),
  自动注册  Number (1),
  强制覆盖	Number (1))
PCTFREE 5;
Alter Table zlTools.ZLFiles Add Constraint ZLFiles_PK Primary Key (名称) USING INDEX PCTFREE 5;

Create Or Replace Procedure Zltools.Zlfiles_Autoupdate
(
  名称_In     In Zlfiles.名称%Type,
  标准md5_In  In Zlfiles.标准md5%Type,
  版本号_In   In Zlfiles.版本号%Type,
  修改日期_In In Zlfiles.修改日期%Type,
  加入日期_In In Zlfiles.加入日期%Type,
  文件类型_In In Zlfiles.文件类型%Type,
  安装路径_In In Zlfiles.安装路径%Type,
  业务部件_In In Zlfiles.业务部件%Type,
  所属系统_In In Zlfiles.所属系统%Type,
  文件说明_In In Zlfiles.文件说明%Type,
  自动注册_In In Zlfiles.自动注册%Type,
  强制覆盖_In In Zlfiles.强制覆盖%Type
) Is
  n_Count Number(3);
Begin
  n_Count := 0;
  --部件更新
  For Rs In (Select Rowid From Zlfiles a Where Upper(a.名称) = Upper(名称_In)) Loop
    n_Count := n_Count + 1;
    Update Zlfiles
    Set 名称 = 名称_In, 标准md5 = 标准md5_In, 版本号 = 版本号_In, 修改日期 = 修改日期_In, 加入日期 = 加入日期_In, 文件类型 = 文件类型_In, 安装路径 = 安装路径_In,
        业务部件 = 业务部件_In, 所属系统 = 所属系统_In, 文件说明 = 文件说明_In, 自动注册 = 自动注册_In, 强制覆盖 = 强制覆盖_In
    Where Rowid = Rs.Rowid;
  End Loop;
  --新增部件
  If n_Count = 0 Then
    Insert Into Zlfiles
      (名称, 标准md5, 版本号, 修改日期, 加入日期, 文件类型, 安装路径, 业务部件, 所属系统, 文件说明, 自动注册, 强制覆盖)
    Values
      (名称_In, 标准md5_In, 版本号_In, 修改日期_In, 加入日期_In, 文件类型_In, 安装路径_In, 业务部件_In, 所属系统_In, 文件说明_In, 自动注册_In, 强制覆盖_In);
  End If;
  n_Count := 0;
  --部件更新
  For Rs In (Select Rowid From Zlfilesupgrade a Where Upper(a.文件名) = Upper(名称_In)) Loop
    n_Count := n_Count + 1;
    Update Zlfilesupgrade
    Set 文件名 = 名称_In, 文件类型 = 文件类型_In, 安装路径 = 安装路径_In, 业务部件 = 业务部件_In, 所属系统 = 所属系统_In, 文件说明 = 文件说明_In, 自动注册 = 自动注册_In,
        强制覆盖 = 强制覆盖_In
    Where Rowid = Rs.Rowid;
  End Loop;
  --新增部件
  If n_Count = 0 Then
    Insert Into Zlfilesupgrade
      (文件名, 文件类型, 安装路径, 业务部件, 所属系统, 文件说明, 自动注册, 强制覆盖)
    Values
      (名称_In, 文件类型_In, 安装路径_In, 业务部件_In, 所属系统_In, 文件说明_In, 自动注册_In, 强制覆盖_In);
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zlfiles_Autoupdate;
/
--00000:刘硕,2016-06-30,修正以前的版本比较函数
Create Or Replace Function Zltools.Zlverdiff
(
  Verpre_In  Varchar2,
  Vernext_In Varchar2
) Return Number
--返回：1、前一版本大；-1、前一版本小；0、两个版本相同
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
      --没找到句点，就把整个字符串作为版本
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
    --当前一个剩余段为空，后一个剩余段不为空，则将前一个赋值为0，继续比较
    If v_Pre Is Null And Not v_Next Is Null Then
      v_Pre := '0';
    End If;
  End Loop;
  Return 0;
End;
/
--00000:刘硕,2016-06-30,增加版本比较函数公共函数
Create Or Replace Function Zltools.Zlcheck_Version_Upon
--功能：判断传入的版本是否在用户当前使用版本之上
  --说明：主要应用于数据修正脚本。
  --Sysno_In：系统编号
  --Version_In：受影响的最低版本号
  --返回：1=用户当前使用版本大于或等于指定版本之上。0-用户当前使用版本小于指定版本。
(
  Sysno_In   Zlsystems.编号%Type,
  Version_In Zlsystems.版本号%Type
) Return Number Is
  v_Startversion Varchar2(20);
  n_Count        Number(5);
Begin
  --获取用户当前使用版本
  --1、没有升迁记录，则获取zlsysTem记录版本号
  --2、有升迁记录，存在最近两天之内的升迁记录，则获取最小的起始版本。
  --3、有升迁记录，不存在最近两天之内的升迁记录，则获取后一条升迁记录的结果版本。
  Select Count(1) Into n_Count From Zlupgrade Where Nvl(系统, 0) = Sysno_In;
  --该系统没有升迁记录
  If n_Count = 0 Then
    Select 版本号 Into v_Startversion From Zlsystems Where 编号 = Sysno_In;
  Else
    Select Max(原始版本), Count(1)
    Into v_Startversion, n_Count
    From (Select 原始版本, 升迁时间
           From Zlupgrade
           Where Nvl(系统, 0) = Sysno_In And 升迁时间 Between Sysdate - 2 And Sysdate
           Order By 升迁时间)
    Where Rownum < 2;
    If n_Count = 0 Then
      Select Max(结果版本), Count(1)
      Into v_Startversion, n_Count
      From (Select 结果版本, 升迁时间 From Zlupgrade Where Nvl(系统, 0) = Sysno_In Order By 升迁时间 Desc)
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
--94568:刘硕,2016-05-14,服务窗信息维护
ALTER TABLE  zltools.zlRegInfo ADD 站点 Varchar2(1);
ALTER TABLE zltools.zlRegInfo Drop CONSTRAINT zlRegInfo_UQ_项目 Cascade Drop Index;
ALTER TABLE zltools.zlRegInfo ADD CONSTRAINT zlRegInfo_UQ_项目 UNIQUE (项目,行号,站点) USING INDEX PCTFREE 5;

ALTER TABLE  zltools.Zlunitinfoimage ADD 站点 Varchar2(1);
ALTER TABLE zltools.zlUnitInfoImage Drop CONSTRAINT zlUnitInfoImage_PK Cascade Drop Index;
Alter Table zlTools.zlUnitInfoImage Add Constraint zlUnitInfoImage_UQ_项目 UNIQUE (项目,站点) USING INDEX PCTFREE 5;
Alter Table ZLTOOLS.zlUnitInfoImage Modify 项目  constraint zlUnitInfoImage_NN_项目   not  null;

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
    内容_In In Zlreginfo.内容%Type,
    站点_In In Zlreginfo.站点%Type
  );
  --功能：插入删除或更新Zlunitinfoitem
  --调用列表：frmUnitInfoEdit,frmUnitItemEdit
  --参数说明：
  --Type_n:0-新增，1-修改,2-删除
  --编码_In：项目编码
  --名称_In:项目名称,站点
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
        Delete Zltools.Zlunitinfoimage
        Where 项目 = Substr(Key_In, 1, Instr(Key_In, ',') - 1) And
              Nvl(站点, '空空') = Nvl(Substr(Key_In, Instr(Key_In, ',') + 1), '空空');
      Else
        If Cls_In = 1 Then
          Update Zltools.Zlunitinfoimage
          Set 图片 = Empty_Blob()
          Where 项目 = Substr(Key_In, 1, Instr(Key_In, ',') - 1) And
                Nvl(站点, '空空') = Nvl(Substr(Key_In, Instr(Key_In, ',') + 1), '空空');
          If Sql%Rowcount = 0 Then
            Insert Into Zltools.Zlunitinfoimage
              (项目, 图片, 站点)
            Values
              (Substr(Key_In, 1, Instr(Key_In, ',') - 1), Empty_Blob(), Substr(Key_In, Instr(Key_In, ',') + 1));
          End If;
        End If;
        Select 图片
        Into l_Blob
        From Zltools.Zlunitinfoimage
        Where 项目 = Substr(Key_In, 1, Instr(Key_In, ',') - 1) And
              Nvl(站点, '空空') = Nvl(Substr(Key_In, Instr(Key_In, ',') + 1), '空空')
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
      Select 图片
      Into l_Blob
      From Zltools.Zlunitinfoimage
      Where 项目 = Substr(Key_In, 1, Instr(Key_In, ',') - 1) And
            Nvl(站点, '空空') = Nvl(Substr(Key_In, Instr(Key_In, ',') + 1), '空空');
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
    内容_In In Zlreginfo.内容%Type,
    站点_In In Zlreginfo.站点%Type
  ) Is
    --项目_In：项目名称
    --行号_In:为空或者为1时先删除该项目，然后插入
    --内容_In:为空则不插入
  Begin
    If Nvl(行号_In, 0) < 2 Then
      Delete Zlreginfo Where 项目 = 项目_In And Nvl(站点, '空空') = Nvl(站点_In, '空空');
    End If;
    If Not 内容_In Is Null Then
      Insert Into Zlreginfo (项目, 行号, 内容, 站点) Values (项目_In, 行号_In, 内容_In, 站点_In);
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


--91515:张永康,2016-08-16,根据西渠道反馈，补充处理支持高并发业务
alter index ZLTOOLS.ZLREPORTS_PK initrans 20;
alter index ZLTOOLS.ZLREPORTS_UQ_编号 initrans 20;
alter index ZLTOOLS.ZLREPORTS_IX_程序ID initrans 20;


--98644:张永康,2016-07-15,以最终截止时间来判断每批次可转出的数据
alter table zltools.zlDataMove add 本次最终日期 date;
