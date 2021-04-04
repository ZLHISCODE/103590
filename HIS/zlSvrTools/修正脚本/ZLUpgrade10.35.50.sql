----10.35.40---》10.35.50
--107086:刘硕,2017-03-15,版本号不兼容
update Zltools.Zlfilesupgrade set 版本号=Null;
--108005:刘硕,2017-04-06,共享方式版本号不兼容导致无法升级
Update Zltools.Zlfilesupgrade
Set 版本号 = '1000350040'
Where Upper(文件名) In ('ZLHISCRUST.EXE', '7Z.EXE;7Z.DLL', 'AAMD532.DLL', 'ZLRUNAS.EXE', 'REGCOM.DLL', 'GACUTIL.EXE','GACUTIL.EXE.CONFIG');
--107086:刘硕,2017-03-15,版本号不兼容
Alter Table Zltools.Zlfilesupgrade add 文件版本号 Varchar2(20);
Create Or Replace Procedure Zltools.Zlfilesupgrade_Repair Is
Begin
  Delete From Zlfilesupgrade a Where Exists (Select 1 From Zlfiles b Where Upper(b.名称) = Upper(a.文件名));
  Insert Into Zlfilesupgrade
    (文件名, 文件类型, 安装路径, 附加安装路径, 业务部件, 所属系统, 文件说明, 自动注册, 强制覆盖)
    Select 名称, 文件类型, 安装路径, 附加安装路径, 业务部件, 所属系统, 文件说明, 自动注册, 强制覆盖 From Zlfiles;
  Update Zlfilesupgrade Set Md5 = Null;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zlfilesupgrade_Repair;
/
--107146:刘硕,2017-03-16,部分客户端无法升级
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
  If 类型_In = 0 Then
    Insert Into Zltools.Zlreginfo
      (项目, 内容)
      Select '服务器目录0', ''
      From Dual
      Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = '服务器目录0')
      Union All
      Select '访问用户0', ''
      From Dual
      Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = '访问用户0')
      Union All
      Select '访问密码0', ''
      From Dual
      Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = '访问密码0')
      Union All
      Select '服务器目录', ''
      From Dual
      Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = '服务器目录')
      Union All
      Select '访问用户', ''
      From Dual
      Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = '访问用户')
      Union All
      Select '访问密码', ''
      From Dual
      Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = '访问密码');
    Update Zltools.Zlreginfo Set 内容 = '0' Where 项目 = '升级类型';
    Update Zltools.Zlreginfo Set 内容 = 位置_In Where 项目 = '服务器目录0';
    Update Zltools.Zlreginfo Set 内容 = 用户名_In Where 项目 = '访问用户0';
    Update Zltools.Zlreginfo Set 内容 = 密码_In Where 项目 = '访问密码0';
    Update Zltools.Zlreginfo Set 内容 = 位置_In Where 项目 = '服务器目录';
    Update Zltools.Zlreginfo Set 内容 = 用户名_In Where 项目 = '访问用户';
    Update Zltools.Zlreginfo Set 内容 = 密码_In Where 项目 = '访问密码';
    Update Zltools.Zlclients Set 升级服务器 = 0;
  Else
    Insert Into Zltools.Zlreginfo
      (项目, 内容)
      Select 'FTP服务器0', ''
      From Dual
      Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = 'FTP服务器0')
      Union All
      Select 'FTP用户0', ''
      From Dual
      Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = 'FTP用户0')
      Union All
      Select 'FTP密码0', ''
      From Dual
      Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = 'FTP密码0')
      Union All
      Select 'FTP端口0', ''
      From Dual
      Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = 'FTP端口0')
      Union All
      Select 'FTP服务器', ''
      From Dual
      Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = 'FTP服务器')
      Union All
      Select 'FTP用户', ''
      From Dual
      Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = 'FTP用户')
      Union All
      Select 'FTP密码', ''
      From Dual
      Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = 'FTP密码')
      Union All
      Select 'FTP端口', ''
      From Dual
      Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = 'FTP端口');
    Update Zltools.Zlreginfo Set 内容 = '1' Where 项目 = '升级类型';
    Update Zltools.Zlreginfo Set 内容 = 位置_In Where 项目 = 'FTP服务器0';
    Update Zltools.Zlreginfo Set 内容 = 用户名_In Where 项目 = 'FTP用户0';
    Update Zltools.Zlreginfo Set 内容 = 密码_In Where 项目 = 'FTP密码0';
    Update Zltools.Zlreginfo Set 内容 = 端口_In Where 项目 = 'FTP端口0';
    Update Zltools.Zlreginfo Set 内容 = 位置_In Where 项目 = 'FTP服务器';
    Update Zltools.Zlreginfo Set 内容 = 用户名_In Where 项目 = 'FTP用户';
    Update Zltools.Zlreginfo Set 内容 = 密码_In Where 项目 = 'FTP密码';
    Update Zltools.Zlreginfo Set 内容 = 端口_In Where 项目 = 'FTP端口';
    Update Zltools.Zlclients Set Ftp服务器 = 0;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zlreginfo_Defaultserver;
/

--108032:高腾,2017-04-07,添加"数据连接"功能
Insert Into Zltools.Zlsvrtools (编号, 上级, 标题, 快键, 说明, 次序) Values ('0207', '02', '数据连接', 'S', Null, 19);

--108032:高腾,2017-04-07,添加"数据连接"功能
Create Table Zltools.zlConnections(
    编号 number(4),
    名称 varchar2(20),
    用户名 varchar2(30),
    密码 varchar2(30),
    IP varchar2(30),
    端口 number(5),
    实例名 varchar2(50),
    说明 varchar2(500));
    
Alter Table Zltools.zlConnections Add Constraint zlConnections_PK Primary Key (编号) Using Index;
Alter Table Zltools.zlConnections Add Constraint zlConnections_UQ_名称 Unique (名称) Using Index;

--107979:余智勇,2017-04-17,自定义报表工具支持多数据连接
Alter Table Zltools.zlRPTDatas Add 数据连接编号 Number(4);
Alter Table Zltools.zlRPTDatas Add Constraint ZLRPTDATAS_UQ_数据连接编号 Unique(数据连接编号, ID) Using Index;
Alter Table Zltools.zlRPTDatas Add Constraint ZLRPTDATAS_FK_数据连接编号 Foreign Key(数据连接编号) References Zltools.ZlConnections(编号) Enable Novalidate;

--108032:高腾,2017-04-07,添加"数据连接"功能
Create Or Replace Procedure Zltools.Zl_Zlconnections_Edit
(
  操作_In   Number, --0-新增,1-修改,2-删除
  编号_In   Zlconnections.编号%Type,
  名称_In   Zlconnections.名称%Type := Null,
  用户名_In Zlconnections.用户名%Type := Null,
  密码_In   Zlconnections.密码%Type := Null,
  Ip_In     Zlconnections.Ip%Type := Null,
  端口_In   Zlconnections.端口%Type := Null,
  实例名_In Zlconnections.实例名%Type := Null,
  说明_In   Zlconnections.说明%Type := Null
) Is
  n_编号    Zlconnections.编号%Type;
  n_Count   Number(1);
  v_Err_Msg Varchar2(255);
  Err_Item Exception;
Begin
  If 操作_In = 0 Then
    Select Count(1) Into n_Count From Zlconnections Where 名称 = 名称_In;
    If n_Count = 1 Then
      v_Err_Msg := '该连接名称已存在！';
      Raise Err_Item;
    End If;
    Select Nvl(Max(编号), 0) Into n_编号 From Zlconnections;
    Insert Into Zlconnections
      (编号, 名称, 用户名, 密码, Ip, 端口, 实例名, 说明)
    Values
      (n_编号 + 1, 名称_In, 用户名_In, 密码_In, Ip_In, 端口_In, 实例名_In, 说明_In);
  Elsif 操作_In = 1 Then
    Update Zlconnections
    Set 用户名 = 用户名_In, 名称 = 名称_In, 密码 = 密码_In, Ip = Ip_In, 端口 = 端口_In, 实例名 = 实例名_In, 说明 = 说明_In
    Where 编号 = 编号_In;
  Else
    Delete Zlconnections Where 编号 = 编号_In;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zlconnections_Edit;
/