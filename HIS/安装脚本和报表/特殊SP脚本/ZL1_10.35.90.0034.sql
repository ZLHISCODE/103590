----------------------------------------------------------------------------------------------------------------
--本脚本支持从ZLHIS+ v10.35.90升级到 v10.35.90
--请以数据空间的所有者登录PLSQL并执行下列脚本
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------
--结构修正部份
------------------------------------------------------------------------------



------------------------------------------------------------------------------
--数据修正部份
------------------------------------------------------------------------------


-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--129301:王振涛,2018-10-22,修改采血管材料
--129301:王振涛,2018-08-15,采血管更新
Create Or Replace Procedure Zl_采血管类型_Update
(
  编码_In   In 采血管类型.编码%Type,
  名称_In   In 采血管类型.名称%Type,
  规格_In   In 采血管类型.规格%Type,
  添加剂_In In 采血管类型.添加剂%Type,
  采血量_In In 采血管类型.采血量%Type,
  颜色_In   In 采血管类型.颜色%Type,
  材料id_In In 采血管类型.材料id%Type := Null
) Is
  v_材料id  Number;
  v_名称    采血管类型.名称%Type;
  v_规格    采血管类型.规格%Type;
  v_添加剂  采血管类型.添加剂%Type;
  v_采血量  采血管类型.采血量%Type;
  v_颜色    采血管类型.颜色%Type;
  v_材料id1 采血管类型.材料id%Type;
Begin
  If Nvl(材料id_In, 0) <> 0 Then
    v_材料id := 材料id_In;
  Else
    v_材料id := Null;
  End If;
  Begin
    Select 名称, 规格, 添加剂, 采血量, 颜色, 材料id
    Into v_名称, v_规格, v_添加剂, v_采血量, v_颜色, v_材料id1
    From 采血管类型
    Where 编码 = 编码_In;
  Exception
    When Others Then
      v_名称 := '空';
  End;
  If Nvl(v_名称, 0) <> Nvl(名称_In, 0) Or Nvl(v_规格, 0) <> Nvl(规格_In, 0) Or Nvl(v_添加剂, 0) <> Nvl(添加剂_In, 0) Or
     Nvl(v_采血量, 0) <> Nvl(采血量_In, 0) Or Nvl(v_颜色, 0) <> Nvl(颜色_In, 0) Or Nvl(v_材料id, 0) <> Nvl(v_材料id1, 0) Then
    If v_名称 <> '空' Then
      Update 采血管类型
      Set 名称 = 名称_In, 规格 = 规格_In, 添加剂 = 添加剂_In, 采血量 = 采血量_In, 颜色 = 颜色_In, 材料id = v_材料id
      Where 编码 = 编码_In;
      b_Message.Zlhis_Dictlis_008(编码_In, 名称_In, Null, 规格_In, 添加剂_In, 采血量_In, 颜色_In, v_材料id);
    End If;
  End If;
  If Sql%NotFound Then
    Insert Into 采血管类型
      (编码, 名称, 规格, 添加剂, 采血量, 颜色, 材料id)
    Values
      (编码_In, 名称_In, 规格_In, 添加剂_In, 采血量_In, 颜色_In, v_材料id);
    b_Message.Zlhis_Dictlis_007(编码_In, 名称_In, Null, 规格_In, 添加剂_In, 采血量_In, 颜色_In, v_材料id);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_采血管类型_Update;
/


------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0034' Where 编号=&n_System;
Commit;