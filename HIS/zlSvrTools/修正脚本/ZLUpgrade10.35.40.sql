----10.35.30---》10.35.40
--000000:刘硕,2017-02-20,调整目标版本
alter table zlTools.zlUpgradeLog modify 目标版本 VARCHAR2(20);

--104473:陈振原,2017-02-15,客户端升级管理,版本号数据结构修正
Update zltools.zlFilesUpgrade Set 版本号 = Null;
Alter Table Zltools.Zlfilesupgrade Modify(版本号 Varchar2(20));

--102830:刘硕,2016-11-21,支持文件下载到多个目录
alter table zltools.Zlfilesupgrade  add 附加安装路径 varchar2(500);

--102814:刘硕,2016-11-21,强制立即升级
alter table zltools.zlclients  add 是否立即升级 number(1);

--104473:陈振原,2016-12-26,客户端升级管理,删除站点文件收集、站点部件升级、升级文件管理菜单项目删除
delete from zltools.zlsvrtools where  编号 in('0311','0307','0309');

--104473:陈振原,2016-12-26,客户端升级管理,新增客户端管理升级工具菜单项目
insert into zltools.zlsvrtools(编号,上级,标题,快键,说明,次序) select '0307','03','客户端升级管理','A','',22 from dual;

--104473:陈振原,2016-12-26,客户端升级管理,客户端表新增字段
alter table zltools.zlclients add  是否预升级 NUMBER(1) default 0;
alter table zltools.zlclients add  升级说明 VARCHAR2(2000);
alter table zltools.zlclients add  收集说明 VARCHAR2(2000);
alter table zltools.zlclients add  修复说明 VARCHAR2(2000);
alter table zltools.zlclients add  预升级说明 VARCHAR2(2000);
alter table zltools.zlclients add  升级文件服务器 NUMBER(3);
alter table zltools.zlclients add  修复状态 NUMBER(1);
alter table zltools.zlclients add  收集状态 NUMBER(1);
alter table zltools.zlclients add  批次 NUMBER(5);

--104473:陈振原,2016-12-26,客户端升级管理,新增附件安装路径
alter table zltools.zlfiles add  附加安装路径 varchar2(500);

--104473:陈振原,2016-12-26,客户端升级管理,新增服务器配置表
Create Table ZLTOOLS.ZLUpgradeServer(
  编号     Number(3),
  类型     Number(1),
  位置     Varchar2(100),
  用户名   Varchar2(20),
  密码     Varchar2(40),
  端口     Number(5),
  是否升级 Number(1),
  是否缺省 Number(1),
  是否收集 Number(1),
  收集类型 Varchar2(100),
  批次        NUMBER(5))
PCTFREE 5;
ALTER TABLE Zltools.ZLUpgradeServer ADD CONSTRAINT ZLUpgradeServer_UQ_位置 Unique (位置) USING INDEX PCTFREE 5;
ALTER TABLE Zltools.ZLUpgradeServer ADD CONSTRAINT ZLUpgradeServer_PK_编号 PRIMARY KEY (编号) USING INDEX;

--104473:陈振原,2016-12-26,客户端升级管理,新增服务器配置
Create Or Replace Procedure Zltools.Zl_Zlupgradeserver_Insert
(
  编号_In     In Zlupgradeserver.编号%Type,
  类型_In     In Zlupgradeserver.类型%Type,
  位置_In     In Zlupgradeserver.位置%Type,
  用户名_In   In Zlupgradeserver.用户名%Type,
  密码_In     In Zlupgradeserver.密码%Type,
  端口_In     In Zlupgradeserver.端口%Type,
  是否升级_In In Zlupgradeserver.是否升级%Type,
  是否缺省_In In Zlupgradeserver.是否缺省%Type,
  是否收集_In In Zlupgradeserver.是否收集%Type,
  收集类型_In In Zlupgradeserver.收集类型%Type
) Is
Begin
  --判断升级缺省服务器以及收集缺省服务器
  If 是否缺省_In = 1 Then
    Update Zlupgradeserver Set 是否缺省 = 0 Where Nvl(是否缺省, 0) = 1;
  End If;
  If 是否收集_In = 1 Then
    Update Zlupgradeserver Set 是否收集 = 0 Where Nvl(是否收集, 0) = 1;
  End If;
  --插入记录 
  Insert Into Zlupgradeserver
    (编号, 类型, 位置, 用户名, 密码, 端口, 是否升级, 是否缺省, 是否收集, 收集类型, 批次)
  Values
    (编号_In, 类型_In, 位置_In, 用户名_In, 密码_In, 端口_In, 是否升级_In, 是否缺省_In, 是否收集_In, 收集类型_In, 0);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
  
End Zl_Zlupgradeserver_Insert;
/


--104473:陈振原,2016-12-26,客户端升级管理,删除服务器配置
Create Or Replace Procedure Zltools.Zl_Zlupgradeserver_Delete(编号_In In Zlupgradeserver.编号%Type) Is
Begin
  --删除数据
  Delete From Zlupgradeserver Where 编号 = 编号_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
  
End Zl_Zlupgradeserver_Delete;
/


--104473:陈振原,2016-12-26,客户端升级管理,修改服务器配置
Create Or Replace Procedure Zltools.Zl_Zlupgradeserver_Update
(
  编号_In     In Zlupgradeserver.编号%Type,
  类型_In     In Zlupgradeserver.类型%Type,
  位置_In     In Zlupgradeserver.位置%Type,
  用户名_In   In Zlupgradeserver.用户名%Type,
  密码_In     In Zlupgradeserver.密码%Type,
  端口_In     In Zlupgradeserver.端口%Type,
  是否升级_In In Zlupgradeserver.是否升级%Type,
  是否缺省_In In Zlupgradeserver.是否缺省%Type,
  是否收集_In In Zlupgradeserver.是否收集%Type,
  收集类型_In In Zlupgradeserver.收集类型%Type,
  Intedittype Pls_Integer
) Is
  --升级 Zlupgradeserver.是否升级%Type;
  --缺省 Zlupgradeserver.是否缺省%Type;
  --收集 Zlupgradeserver.是否收集%Type;
Begin
  If 是否缺省_In = 1 Then
    Update Zlupgradeserver Set 是否缺省 = 0 Where Nvl(是否缺省, 0) = 1;
  End If;
  If 是否收集_In = 1 Then
    Update Zlupgradeserver Set 是否收集 = 0 Where Nvl(是否收集, 0) = 1;
  End If;
  --修改类型为0 可以修改所有字段数据
  If Intedittype = 0 Then
    Update Zlupgradeserver
    Set 类型 = 类型_In, 位置 = 位置_In, 用户名 = 用户名_In, 密码 = 密码_In, 端口 = 端口_In, 是否升级 = 是否升级_In, 是否缺省 = 是否缺省_In, 是否收集 = 是否收集_In,
        收集类型 = 收集类型_In
    Where 编号 = 编号_In;
  End If;

  --修改类型为1 修改服务器缺省设置
  If Intedittype = 1 Then
    Update Zlupgradeserver
    Set 是否升级 = 是否升级_In, 是否缺省 = 是否缺省_In, 是否收集 = 是否收集_In, 收集类型 = 收集类型_In
    Where 编号 = 编号_In;
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlupgradeserver_Update;
/

--104473:陈振原,2017-1-16,客户端升级管理,在用文件清单修复
Create Or Replace Procedure Zltools.Zlfilesupgrade_Repair Is
Begin
  Delete From zlFilesUpgrade C
  Where c.文件类型 <> 4 Or
        c.文件名 In
        (Select 名称 From zlFilesUpgrade A, Zlfiles B Where a.文件名 = b.名称 And (a.文件类型 = 4 Or b.文件类型 = 4));

  Insert Into zlFilesUpgrade
    (文件名, Md5, 版本号, 修改日期, 加入日期, 文件类型, 安装路径, 业务部件, 所属系统, 文件说明, 自动注册, 强制覆盖)
    Select 名称, 标准md5, 版本号, 修改日期, 加入日期, 文件类型, 安装路径, 业务部件, 所属系统, 文件说明, 自动注册, 强制覆盖
    From Zlfiles B
    Where Not Exists (Select 1 From zlFilesUpgrade R Where r.文件名 = b.名称);

  Update zlFilesUpgrade Set Md5 = Null;

End Zlfilesupgrade_Repair;
/

--104473:陈振原,2016-12-26,客户端升级管理,默认服务器设置
Create Or Replace Procedure Zltools.Zlreginfo_Defaultserver
(
  类型_In   In Zlreginfo.内容%Type,
  位置_In   In Zlreginfo.内容%Type,
  用户名_In In Zlreginfo.内容%Type,
  密码_In   In Zlreginfo.内容%Type,
  端口_In   In Zlreginfo.内容%Type
) Is
Begin

  Insert Into Zltools.Zlreginfo
    (项目, 内容)
    Select '升级类型', '' From Dual Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = '升级类型');
  Insert Into Zltools.Zlreginfo
    (项目, 内容)
    Select 'FTP服务器0', '' From Dual Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = 'FTP服务器0');
  Insert Into Zltools.Zlreginfo
    (项目, 内容)
    Select 'FTP用户0', '' From Dual Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = 'FTP用户0');
  Insert Into Zltools.Zlreginfo
    (项目, 内容)
    Select 'FTP密码0', '' From Dual Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = 'FTP密码0');
  Insert Into Zltools.Zlreginfo
    (项目, 内容)
    Select 'FTP端口0', '' From Dual Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = 'FTP端口0');
  Insert Into Zltools.Zlreginfo
    (项目, 内容)
    Select '服务器目录0', '' From Dual Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = '服务器目录0');
  Insert Into Zltools.Zlreginfo
    (项目, 内容)
    Select '访问用户0', '' From Dual Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = '访问用户0');
  Insert Into Zltools.Zlreginfo
    (项目, 内容)
    Select '访问密码0', '' From Dual Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = '访问密码0');

  If 类型_In = 0 Then
    Update Zltools.Zlreginfo Set 内容 = '0' Where 项目 = '升级类型';
    Update Zltools.Zlreginfo Set 内容 = 位置_In Where 项目 = '服务器目录0';
    Update Zltools.Zlreginfo Set 内容 = 用户名_In Where 项目 = '访问用户0';
    Update Zltools.Zlreginfo Set 内容 = 密码_In Where 项目 = '访问密码0';
    Update Zltools.Zlclients Set Ftp服务器 = '0';
  Else
    Update Zltools.Zlreginfo Set 内容 = '1' Where 项目 = '升级类型';
    Update Zltools.Zlreginfo Set 内容 = 位置_In Where 项目 = 'FTP服务器0';
    Update Zltools.Zlreginfo Set 内容 = 用户名_In Where 项目 = 'FTP用户0';
    Update Zltools.Zlreginfo Set 内容 = 密码_In Where 项目 = 'FTP密码0';
    Update Zltools.Zlreginfo Set 内容 = 端口_In Where 项目 = 'FTP端口0';
    Update Zltools.Zlclients Set 升级服务器 = '0';
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zlreginfo_Defaultserver;
/


--104473:陈振原,2016-12-26,客户端升级管理,客户端升级预升级标志更新
Create Or Replace Procedure Zltools.Zl_Zlclients_Update
(
  工作站_In   In Varchar2,
  更新字段_In In Number,
  更新值_In   In Number
) Is
  --工作站_In 工作站字符串，以逗号相连接
  --更新字段_In 0：升级标志 1：是否预升级 2：收集标志
  --更新值_In 0：不勾选 1：勾选
  v_工作站 Varchar2(4000);
  n_更新值 Number;
Begin
  v_工作站 := 工作站_In;
  n_更新值 := 更新值_In;

  If 更新字段_In = 0 Then
    Update zlClients Set 升级标志 = n_更新值 Where 工作站 In (Select Column_Value From Table(f_Str2list(v_工作站)));
    If n_更新值 = 1 Then
      Update zlClients Set 升级情况 = 0 Where 工作站 In (Select Column_Value From Table(f_Str2list(v_工作站)));
      Update zlClients Set 修复状态 = 0 Where 工作站 In (Select Column_Value From Table(f_Str2list(v_工作站)));
      Update zlClients Set 升级说明 = Null Where 工作站 In (Select Column_Value From Table(f_Str2list(v_工作站)));
      Update zlClients Set 修复说明 = Null Where 工作站 In (Select Column_Value From Table(f_Str2list(v_工作站)));
    End If;
  End If;

  If 更新字段_In = 1 Then
    Update zlClients Set 是否预升级 = n_更新值 Where 工作站 In (Select Column_Value From Table(f_Str2list(v_工作站)));
    If n_更新值 = 1 Then
      Update zlClients Set 预升完成 = 0 Where 工作站 In (Select Column_Value From Table(f_Str2list(v_工作站)));
      Update zlClients Set 预升级说明 = Null Where 工作站 In (Select Column_Value From Table(f_Str2list(v_工作站)));
    End If;
  End If;

  If 更新字段_In = 2 Then
    Update zlClients Set 收集标志 = n_更新值 Where 工作站 In (Select Column_Value From Table(f_Str2list(v_工作站)));
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlclients_Update;
/

CREATE OR REPLACE Procedure ZLTOOLS.Zlfiles_Autoupdate
(
  名称_In         In Zlfiles.名称%Type,
  标准md5_In      In Zlfiles.标准md5%Type,
  版本号_In       In Zlfiles.版本号%Type,
  修改日期_In     In Zlfiles.修改日期%Type,
  加入日期_In     In Zlfiles.加入日期%Type,
  文件类型_In     In Zlfiles.文件类型%Type,
  安装路径_In     In Zlfiles.安装路径%Type,
  业务部件_In     In Zlfiles.业务部件%Type,
  所属系统_In     In Zlfiles.所属系统%Type,
  文件说明_In     In Zlfiles.文件说明%Type,
  自动注册_In     In Zlfiles.自动注册%Type,
  强制覆盖_In     In Zlfiles.强制覆盖%Type,
  附加安装路径_In In Zlfiles.附加安装路径%Type
) Is
  n_Count Number(3);
Begin
  n_Count := 0;
  --部件更新
  For Rs In (Select Rowid From Zlfiles A Where Upper(a.名称) = Upper(名称_In)) Loop
    n_Count := n_Count + 1;
    Update Zlfiles
    Set 名称 = 名称_In, 标准md5 = 标准md5_In, 版本号 = 版本号_In, 修改日期 = 修改日期_In, 加入日期 = 加入日期_In, 文件类型 = 文件类型_In, 安装路径 = 安装路径_In,
        业务部件 = 业务部件_In, 所属系统 = 所属系统_In, 文件说明 = 文件说明_In, 自动注册 = 自动注册_In, 强制覆盖 = 强制覆盖_In, 附加安装路径 = 附加安装路径_In
    Where Rowid = Rs.Rowid;
  End Loop;
  --新增部件
  If n_Count = 0 Then
    Insert Into Zlfiles
      (名称, 标准md5, 版本号, 修改日期, 加入日期, 文件类型, 安装路径, 业务部件, 所属系统, 文件说明, 自动注册, 强制覆盖, 附加安装路径)
    Values
      (名称_In, 标准md5_In, 版本号_In, 修改日期_In, 加入日期_In, 文件类型_In, 安装路径_In, 业务部件_In, 所属系统_In, 文件说明_In, 自动注册_In, 强制覆盖_In, 附加安装路径_In);
  End If;
  n_Count := 0;
  --部件更新
  For Rs In (Select Rowid From zlFilesUpgrade A Where Upper(a.文件名) = Upper(名称_In)) Loop
    n_Count := n_Count + 1;
    Update zlFilesUpgrade
    Set 文件名 = 名称_In, 文件类型 = 文件类型_In, 安装路径 = 安装路径_In, 业务部件 = 业务部件_In, 所属系统 = 所属系统_In, 文件说明 = 文件说明_In, 自动注册 = 自动注册_In,
        强制覆盖 = 强制覆盖_In, 附加安装路径 = 附加安装路径_In
    Where Rowid = Rs.Rowid;
  End Loop;
  --新增部件
  If n_Count = 0 Then
    Insert Into zlFilesUpgrade
      (文件名, 文件类型, 安装路径, 业务部件, 所属系统, 文件说明, 自动注册, 强制覆盖, 附加安装路径)
    Values
      (名称_In, 文件类型_In, 安装路径_In, 业务部件_In, 所属系统_In, 文件说明_In, 自动注册_In, 强制覆盖_In, 附加安装路径_In);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zlfiles_Autoupdate;
/

--105087:刘硕,2017-01-16,支持部门参数类型的变动
Create Or Replace Procedure Zltools.Zl_Parameters_Change
(
  参数id_In   Zlparameters.Id%Type,
  私有_In     Zlparameters.私有%Type,
  本机_In     Zlparameters.本机%Type,
  授权_In     Zlparameters.授权%Type,
  变动人_In   Zlparachangedlog.变动人%Type,
  变动原因_In Zlparachangedlog.变动原因%Type,
  部门_In     Zlparameters.部门%Type := 0
) Is
  v_Temp     Varchar2(200);
  n_模块     Zlparameters.模块%Type;
  n_私有     Zlparameters.私有%Type;
  n_本机     Zlparameters.本机%Type;
  n_部门     Zlparameters.部门%Type;
  n_授权     Zlparameters.授权%Type;
  n_序号     Zlparachangedlog.序号%Type;
  v_变动说明 Zlparachangedlog.变动说明%Type;
  v_变动内容 Zlparachangedlog.变动内容%Type;

  Function Gettype
  (
    模块_In Zlparameters.私有%Type,
    私有_In Zlparameters.私有%Type,
    本机_In Zlparameters.本机%Type,
    部门_In Zlparameters.部门%Type
  ) Return Varchar2 Is
  Begin
    If Nvl(部门_In, 0) = 1 Then
      Return '部门参数';
    End If;
    If Nvl(模块_In, 0) = 0 Then
      --不存模块,证明只有两种类型:公共全局和私有全局 
      If Nvl(私有_In, 0) = 0 Then
        Return '公共全局';
      End If;
      Return '私有全局';
    End If;
  
    --对模块的处理 
    If 本机_In = 0 Then
      --不是本机的情况,只有两种类型:公共模块和私有模块 
      If Nvl(私有_In, 0) = 0 Then
        Return '公共模块';
      End If;
      Return '私有模块';
    End If;
    --对本机的模块进行处理也有两种情况: 
    If Nvl(私有_In, 0) = 0 Then
      Return '本机公共模块';
    End If;
    Return '本机私有模块';
  Exception
    When Others Then
      Return Null;
  End Gettype;
Begin

  Select Nvl(模块, 0), Nvl(私有, 0), Nvl(本机, 0), Nvl(授权, 0), Nvl(部门, 0)
  Into n_模块, n_私有, n_本机, n_授权, n_部门
  From Zlparameters
  Where Id = 参数id_In;
  Select Nvl(Max(序号), 0) + 1 Into n_序号 From Zlparachangedlog Where 参数id = 参数id_In;
  --插入数据 
  --说明变动说明:比如:私有模块变为公用模块。 
  -- 变动内容:说明变动字段的变化情况:比如:私有:1-->0,本机:1-->0 
  v_变动说明 := Null;
  v_变动内容 := Null;
  --类型发生了改变 
  If n_私有 <> Nvl(私有_In, 0) Or n_本机 <> Nvl(本机_In, 0) Or n_部门 <> Nvl(部门_In, 0) Then
    v_Temp     := '从' || Gettype(n_模块, n_私有, n_本机, n_部门);
    v_Temp     := v_Temp || '变为' || Gettype(n_模块, Nvl(私有_In, 0), Nvl(本机_In, 0), Nvl(部门_In, 0));
    v_变动说明 := v_Temp;
    v_Temp     := '';
    If n_部门 <> Nvl(部门_In, 0) Then
      v_Temp := v_Temp || ',部门:' || n_部门 || '-->' || Nvl(部门_In, 0);
    End If;
    If n_私有 <> Nvl(私有_In, 0) Then
      v_Temp := v_Temp || ',私有:' || n_私有 || '-->' || Nvl(私有_In, 0);
    End If;
    If n_私有 <> Nvl(私有_In, 0) Then
      v_Temp := v_Temp || ',本机:' || n_本机 || '-->' || Nvl(本机_In, 0);
    End If;
    v_变动内容 := Substr(v_Temp, 2);
  End If;
  --检查授权发生改变没有 
  If n_授权 <> Nvl(授权_In, 0) Then
    If Not v_变动说明 Is Null Then
      v_变动说明 := v_变动说明 || ',';
    End If;
    If n_授权 = 0 Then
      v_Temp := '不需要授权';
    Else
      v_Temp := '需要授权';
    End If;
    v_变动说明 := Nvl(v_变动说明, '') || '从' || v_Temp || '改为';
    If 授权_In = 0 Then
      v_Temp := '不需要授权';
    Else
      v_Temp := '需要授权';
    End If;
    v_变动说明 := Nvl(v_变动说明, '') || v_Temp;
  
    If Not v_变动内容 Is Null Then
      v_变动内容 := v_变动内容 || ',';
    End If;
    v_变动内容 := Nvl(v_变动内容, '') || '授权:' || n_授权 || '-->' || Nvl(授权_In, 0);
  End If;

  Insert Into Zlparachangedlog
    (参数id, 序号, 变动说明, 变动内容, 变动人, 变动时间, 变动原因)
  Values
    (参数id_In, n_序号, v_变动说明, v_变动内容, 变动人_In, Sysdate, 变动原因_In);

  Update Zlparameters
  Set 私有 = Nvl(私有_In, 0), 本机 = Nvl(本机_In, 0), 授权 = Nvl(授权_In, 0), 部门 = Nvl(部门_In, 0)
  Where Id = 参数id_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Parameters_Change;
/

--104473:陈振原,2016-1-18,客户端升级管理,预升级时间设置存储过程，批量设置增加覆盖参数，允许覆盖或不覆盖
Create Or Replace Procedure Zltools.Zl_Zlclients_Setpretime
(
  n_Mode_In  Number,
  n_Cover_In Number := 1,
  v_工作站_In   Zlclients.工作站%Type := Null,
  d_预升时点_In Zlclients.预升时点%Type := Null
) Is
  v_Timeset Varchar2(300);
  v_Err     Varchar2(500);
  Err_Custom Exception;
  --n_Cover_In 0-不覆盖 1-覆盖 
Begin
  --0-单独对客户端预升级时间设置
  If n_Mode_In = 0 Then
    If v_工作站_In Is Not Null Then
      Update zlClients Set 预升时点 = d_预升时点_In Where 工作站 = v_工作站_In;
    End If;
  --1-预升级时间自动设置n_Cover_In决定覆盖与否
  Elsif n_Mode_In = 1 Then
    Select Max(内容) Into v_Timeset From zlRegInfo Where 项目 = '客户端预升级时间点';
    If v_Timeset Is Not Null Then
      For r_Ip In (Select To_Date(Today || ' ' || Date_d, 'yyyy-mm-dd HH24:mi:ss') 预升时点, 工作站, Ip
                   From (Select 工作站, Ip, Rownum Rn_c From zlClients) A,
                        (Select To_Char(Sysdate, 'yyyy-mm-dd') Today, Column_Value Date_d, Rownum Rn_d, Count(1) Over() Sn
                          From Table(f_Str2list(v_Timeset, ','))) B
                   Where Mod(a.Rn_c, Sn) + 1 = Rn_d) Loop
        If n_Cover_In = 1 Then
          Update zlClients Set 预升时点 = r_Ip.预升时点 Where 工作站 = r_Ip.工作站 And Ip = r_Ip.Ip;
        Elsif n_Cover_In = 0 Then
          Update zlClients
          Set 预升时点 = r_Ip.预升时点
          Where 工作站 = r_Ip.工作站 And Ip = r_Ip.Ip And 预升时点 Is Null;
        End If;
      End Loop;
    Else
      Update zlClients Set 预升时点 = Null;
    End If;
  Elsif n_Mode_In = 3 Then
    Update zlClients Set 预升完成 = 0;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlclients_Setpretime;
/
--00000:刘硕,2017-03-03,配合修改过程补充
CREATE OR REPLACE Procedure ZLTOOLS.Zl_Zlclients_Updateprocess
(
  v_工作站_In  Zlclients.工作站%Type := Null,
  n_Operate_In Number,
  n_State_In   Zlclients.升级情况%Type := 0,
  v_说明_In    Zlclients.说明%Type := Null,
  n_批次_In    Zlclients.批次%Type := Null
  --功能：客户端操作状态处理
  --应用：N_Operate_In=0-主动修复，此时相当于正式升级。清理预升级信息，升级标志，写入主动修复信息
  --                  =1-预升级，清除预升级标志，写入预升级信息
  --                  =2-正式升级，清理预升级信息，升级标志。写入升级信息。
  --                  =3-收集，清理升级标志，写入收集信息。
  --      n_State_In=0-无需操作。1-操作成功。2-操作失败。3-操作执行中
) Is
  v_Err Varchar2(500);
  Err_Custom Exception;
Begin
  If n_Operate_In = 0 Then
    If n_State_In = 1 Then
      Update Zltools.Zlclients
      Set 批次 = n_批次_In, 修复状态 = n_State_In, 修复说明 = v_说明_In, 升级标志 = 0, 升级情况 = 0, 是否预升级 = 0, 预升完成 = 0, 预升级说明 = Null,
          升级说明 = Null, 是否立即升级 = 0
      Where 工作站 = v_工作站_In;
    Else
      Update Zltools.Zlclients Set 修复状态 = n_State_In, 修复说明 = v_说明_In Where 工作站 = v_工作站_In;
    End If;
  Elsif n_Operate_In = 1 Then
    If n_State_In = 1 Then
      Update Zltools.Zlclients
      Set 预升完成 = n_State_In, 预升级说明 = v_说明_In, 是否预升级 = 0
      Where 工作站 = v_工作站_In;
    Else
      Update Zltools.Zlclients Set 预升完成 = n_State_In, 预升级说明 = v_说明_In Where 工作站 = v_工作站_In;
    End If;
  Elsif n_Operate_In = 2 Then
    If n_State_In = 1 Then
      Update Zltools.Zlclients
      Set 批次 = n_批次_In, 升级情况 = n_State_In, 升级说明 = v_说明_In, 升级标志 = 0, 是否预升级 = 0, 预升完成 = 0, 修复状态 = 0, 预升级说明 = Null,
          修复说明 = Null, 是否立即升级 = 0
      Where 工作站 = v_工作站_In;
    Else
      Update Zltools.Zlclients Set 升级情况 = n_State_In, 升级说明 = v_说明_In Where 工作站 = v_工作站_In;
    End If;
  Elsif n_Operate_In = 3 Then
    If n_State_In = 1 Then
      Update Zltools.Zlclients
      Set 收集状态 = n_State_In, 收集说明 = v_说明_In, 收集标志 = 0
      Where 工作站 = v_工作站_In;
    Else
      Update Zltools.Zlclients Set 收集状态 = n_State_In, 收集说明 = v_说明_In Where 工作站 = v_工作站_In;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zlclients_Updateprocess;
/
