--读取图标资源
CREATE OR REPLACE Function Zl_影像查询_读取图标
(
  名称_In  In 影像查询资源.资源名称%Type,
  Pos_In In Number
) Return Varchar2 Is
  v_Buffer Varchar2(32767);
  l_Blob   Blob;
  n_Amount Number := 2000;
  n_Offset Number := 1;
Begin
  Select 图标 Into l_Blob From 影像查询资源 Where 资源名称 = 名称_In;
  
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
End Zl_影像查询_读取图标;
/

--清除用户关联设置
Create Or Replace Procedure zl_影像查询_清除关联
(
   用户ID_In      影像查询关联.用户ID%Type
) Is
Begin
    Delete From 影像查询关联 Where 用户ID=用户ID_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End;
/

--配置用户关联查询方案
Create Or Replace Procedure zl_影像查询_更新关联
(
   用户ID_In      影像查询关联.用户ID%Type,
   方案ID_In      影像查询关联.查询方案ID%Type,
   是否默认_In    影像查询关联.是否默认%Type,
   是否常用_In    影像查询关联.是否常用%Type,
   所属站点_In    影像查询关联.所属站点%Type
) Is
Begin
    Insert Into 影像查询关联(ID,用户ID,查询方案ID,是否默认,是否常用,所属站点)
    Values(影像查询关联_ID.NEXTVAL,用户ID_In,方案ID_In,是否默认_In,是否常用_In,所属站点_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End zl_影像查询_更新关联;
/ 

--配置用户过滤窗口录入条件
Create Or Replace Procedure zl_影像查询_条件配置
(
   用户ID_In      影像查询特性.用户ID%Type,
   查询方案ID_In  影像查询特性.查询方案ID%Type,
   条件配置_In    影像查询特性.条件配置%Type
) Is
Begin
    Update 影像查询特性 
    Set 条件配置=条件配置_In
    Where 用户ID=用户ID_In And 查询方案ID=查询方案ID_In;
   
    If Sql%RowCount <=0 Then
        Insert Into 
               影像查询特性(ID, 用户ID, 查询方案ID, 条件配置)
        Values
               	(影像查询特性_ID.NEXTVAL, 用户ID_In, 查询方案ID_In, 条件配置_In);
    End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End zl_影像查询_条件配置;
/
--********************************************************************************

Create Or Replace Procedure Zl_影像查询_编辑方案内容
(
  Id_In   In 影像查询方案.Id%Type,
  Text_In In 影像查询方案.方案内容%Type
) Is
  l_Clob Clob;
Begin
  Update 影像查询方案 Set 方案内容 = Empty_Clob() Where Id = Id_In;
  Select 方案内容 Into l_Clob From 影像查询方案 Where Id = Id_In For Update;
  Dbms_Lob.Writeappend(l_Clob, Length(Text_In), Text_In);
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_影像查询_编辑方案内容;
/
--------------------------------------------------------------------------------------------------------------

Create Or Replace Procedure Zl_影像查询_删除图标(资源名称_In In 影像查询资源.资源名称 %Type) Is
Begin
  Delete From 影像查询资源 Where 资源名称 = 资源名称_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_影像查询_删除图标;

/

Create Or Replace Procedure Zl_影像查询_新增图标
(
  资源名称_In In 影像查询资源.资源名称%Type,
  资源类型_In In 影像查询资源.资源类型%Type
) Is
  n_Id Number(18);
Begin
  Select 影像查询资源_Id.Nextval Into n_Id From Dual;
  Insert Into 影像查询资源 (Id, 资源名称, 资源类型) Values (n_Id, 资源名称_In, 资源类型_In);
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_影像查询_新增图标;
/

Create Or Replace Procedure Zl_影像查询_保存图标
(
  资源名称_In In 影像查询资源.资源名称%Type,
  图标_In     In Varchar2, --16进制的文件片段或文字片段 
  Cls_In      In Number := 0 --是否清除原来的内容，第一片段传递时为1，以后为0 
) Is
  l_Blob Blob;
Begin

  If Cls_In = 1 Then
    Update 影像查询资源 Set 图标 = Empty_Blob() Where 资源名称 = 资源名称_In;
  End If;
  Select 图标 Into l_Blob From 影像查询资源 Where 资源名称 = 资源名称_In For Update;
  Dbms_Lob.Writeappend(l_Blob, Length(Hextoraw(图标_In)) / 2, Hextoraw(图标_In));
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_影像查询_保存图标;
/

Create Or Replace Procedure Zl_影像查询_新增方案
(
  方案名称_In In 影像查询方案.方案名称%Type,
  方案说明_In In 影像查询方案.方案说明%Type,
  是否默认_In In 影像查询方案.是否默认%Type,
  使用状态_In In 影像查询方案.方案名称%Type,
  是否常用_In In 影像查询方案.是否常用%Type,
  所属模块_In In 影像查询方案.所属模块%Type,
  方案内容_In In 影像查询方案.方案内容%Type
) Is
  n_Id       Number(18);
  n_方案序号 Number(18);
Begin
  Select 影像查询方案_Id.Nextval Into n_Id From Dual;
  Select Nvl(Max(方案序号), 0) + 1 Into n_方案序号 From 影像查询方案 Where 所属模块 = 所属模块_In;

  Insert Into 影像查询方案
    (Id, 方案名称, 方案说明, 是否默认, 使用状态, 方案序号, 是否常用, 所属模块, 版本)
  Values
    (n_Id, 方案名称_In, 方案说明_In, 是否默认_In, 使用状态_In, n_方案序号, 是否常用_In, 所属模块_In, '1');

  Zl_影像查询_编辑方案内容(n_Id, 方案内容_In);

Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_影像查询_新增方案;
/

Create Or Replace Procedure Zl_影像查询_更新方案
(
  Id_In       In 影像查询方案.Id%Type,
  方案名称_In In 影像查询方案.方案名称%Type,
  方案说明_In In 影像查询方案.方案说明%Type,
  是否默认_In In 影像查询方案.是否默认%Type,
  使用状态_In In 影像查询方案.方案名称%Type,
  是否常用_In In 影像查询方案.是否常用%Type,
  所属模块_In In 影像查询方案.所属模块%Type,
  方案内容_In In 影像查询方案.方案内容%Type
) Is
Begin

  Update 影像查询方案
  Set 方案名称 = 方案名称_In, 方案说明 = 方案说明_In, 是否默认 = 是否默认_In, 使用状态 = 使用状态_In, 是否常用 = 是否常用_In, 所属模块 = 所属模块_In, 版本 = 版本 + 1
  Where Id = Id_In;

  Zl_影像查询_编辑方案内容(Id_In, 方案内容_In);
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_影像查询_更新方案;
/

Create Or Replace Procedure Zl_影像查询_删除方案(Id_In In 影像查询方案.Id%Type) Is
Begin
  Delete From 影像查询方案 Where Id = Id_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_影像查询_删除方案;
/

Create Or Replace Procedure Zl_影像查询_移动方案
(
  方案id_In   In 影像查询方案.Id%Type,
  新序号_In   In 影像查询方案.方案序号%Type,
  所属模块_In In 影像查询方案.所属模块%Type
) Is
  n_Order Number;
Begin
  Begin
    Select 方案序号 Into n_Order From 影像查询方案 Where Id = 方案id_In And 所属模块 = 所属模块_In;
  Exception
    When Others Then
      Return;
  End;

  If 新序号_In < n_Order Then
    Update 影像查询方案
    Set 方案序号 = 方案序号 + 1
    Where 方案序号 >= 新序号_In And 方案序号 < n_Order And 所属模块 = 所属模块_In;
  
    Update 影像查询方案 Set 方案序号 = 新序号_In Where Id = 方案id_In And 所属模块 = 所属模块_In;
  Else
    Update 影像查询方案
    Set 方案序号 = 方案序号 - 1
    Where 方案序号 > n_Order And 方案序号 <= 新序号_In And 所属模块 = 所属模块_In;
  
    Update 影像查询方案 Set 方案序号 = 新序号_In Where Id = 方案id_In And 所属模块 = 所属模块_In;
  End If;

Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_影像查询_移动方案;
/

Create Or Replace Procedure Zl_影像查询_默认方案
(
  方案id_In   In 影像查询方案.Id%Type,
  是否默认_In In 影像查询方案.是否默认%Type,
  所属模块_In In 影像查询方案.所属模块%Type
) Is
Begin
  Update 影像查询方案 Set 是否默认 = 0 Where 是否默认 = 1 And 所属模块 = 所属模块_In;
  Update 影像查询方案 Set 是否默认 = 是否默认_In Where Id = 方案id_In And 所属模块 = 所属模块_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_影像查询_默认方案;
/

Create Or Replace Procedure Zl_影像查询_常用方案
(
  方案id_In   影像查询方案.Id%Type,
  是否常用_In 影像查询方案.是否系统查询%Type
) Is
Begin
  Update 影像查询方案 Set 是否常用 = 是否常用_In Where Id = 方案id_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_影像查询_常用方案;
/

Create Or Replace Procedure Zl_影像查询_启用方案
(
  方案id_In   影像查询方案.Id%Type,
  使用状态_In 影像查询方案.使用状态%Type
) Is
Begin
  Update 影像查询方案 Set 使用状态 = 使用状态_In Where Id = 方案id_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_影像查询_启用方案;
/

Create Or Replace Procedure Zl_影像查询_个性化配置
(
  用户id_In     影像查询特性.用户id%Type,
  查询方案id_In 影像查询特性.查询方案id%Type,
  过滤配置_In   影像查询特性.过滤配置%Type,
  列表配置_In   影像查询特性.列表配置%Type
) Is
Begin
  Update 影像查询特性
  Set 过滤配置 = 过滤配置_In, 列表配置 = 列表配置_In
  Where 用户id = 用户id_In And 查询方案id = 查询方案id_In;

  If Sql%RowCount <= 0 Then
    Insert Into 影像查询特性
      (ID, 用户id, 查询方案id, 过滤配置, 列表配置)
    Values
      (影像查询特性_Id.Nextval, 用户id_In, 查询方案id_In, 过滤配置_In, 列表配置_In);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_影像查询_个性化配置;
/


