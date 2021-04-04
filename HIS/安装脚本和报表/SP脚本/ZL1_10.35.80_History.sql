--[连续升级]1
--[管理工具版本号]10.35.80
--本脚本支持从ZLHIS+ v10.35.70 升级到 v10.35.80
--请以系统所有者登录PLSQL并执行下列脚本
--脚本执行后，请手工升级导出报表
Define n_System=100;
-------------------------------------------------------------------------------
--结构修正部份
-------------------------------------------------------------------------------
--92802:冉俊明,2017-11-14,消费卡管理改进(含问题号：92802、102917、110651)
Alter Table 病人卡结算记录 Rename Column 结算金额 To 实收金额;
Alter Table 病人卡结算记录 Add 记录性质 Number(3);
Alter Table 病人卡结算记录 Add 结算id Number(18);
Alter Table 病人卡结算记录 Add 应收金额 Number(16, 5);
Alter Table 病人卡结算记录 Add 扣率 Number(16, 5);
Alter Table 病人卡结算记录 Add 操作员编号 Varchar2(6);
Alter Table 病人卡结算记录 Add 操作员姓名 Varchar2(20);
Alter Table 病人卡结算记录 Add 登记时间 Date;
Alter Table 病人卡结算记录 Add 结算序号 Number(18);
Alter Table 病人卡结算记录 Add 交易序号 Number(18);


-------------------------------------------------------------------------------
--数据修正部份
-------------------------------------------------------------------------------
--92802:冉俊明,2017-11-14,消费卡管理改进(含问题号：92802、102917、110651)
--备份数据
Declare
  n_Count Number(5);
Begin
  Select Count(1) Into n_Count From User_Tables Where Table_Name = Upper('病人卡结算记录_20171114_bak');
  If n_Count = 0 Then
    Execute Immediate 'Create Table 病人卡结算记录_20171114_bak As Select * From 病人卡结算记录';
  End If;
  Select Count(1) Into n_Count From User_Tables Where Table_Name = Upper('病人卡结算对照_20171114_bak');
  If n_Count = 0 Then
    Execute Immediate 'Create Table 病人卡结算对照_20171114_bak As Select * From 病人卡结算对照';
  End If;
End;
/

--升级病人卡结算记录
Declare
  Cursor c_结算数据 Is
    Select c.Id, a.预交id, b.操作员编号, b.操作员姓名, b.收款时间
    From 病人卡结算记录 C, 病人卡结算对照 A, 病人预交记录 B
    Where c.Id = a.卡结算id And a.预交id = b.Id And c.记录性质 Is Null;

  Type t_Id Is Table Of 病人卡结算记录.Id%Type;
  Type t_结算id Is Table Of 病人卡结算对照.预交id%Type;
  Type t_操作员编号 Is Table Of 病人预交记录.操作员编号%Type;
  Type t_操作员姓名 Is Table Of 病人预交记录.操作员姓名%Type;
  Type t_收款时间 Is Table Of 病人预交记录.收款时间%Type;
  c_Id         t_Id := t_Id();
  c_结算id     t_结算id := t_结算id();
  c_操作员编号 t_操作员编号 := t_操作员编号();
  c_操作员姓名 t_操作员姓名 := t_操作员姓名();
  c_收款时间   t_收款时间 := t_收款时间();

  n_Array_Size Number := 10000; --每批一万,多了可能PGA不够
  I            Number(8) := 0; --每修正10万条记录提交一次,多了可能Undo不够,少了提交过于频繁
Begin
  Open c_结算数据();
  Loop
    Fetch c_结算数据 Bulk Collect
      Into c_Id, c_结算id, c_操作员编号, c_操作员姓名, c_收款时间 Limit n_Array_Size;
    Exit When c_Id.Count = 0;
  
    Forall K In 1 .. c_Id.Count
      Update 病人卡结算记录
      Set 记录性质 = 4, 应收金额 = 实收金额, 扣率 = 100, 结算id = c_结算id(K), 操作员编号 = c_操作员编号(K), 操作员姓名 = c_操作员姓名(K), 登记时间 = c_收款时间(K),
          结算序号 = ID, 交易序号 = -1
      Where ID = c_Id(K);
  
    If I = 10 Then
      Commit;
      I := 0;
    Else
      I := I + 1;
    End If;
  End Loop;
  Close c_结算数据;

  Commit;
End;
/




-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------





-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------






-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------





---------------------------------------------------------------------------------------------------
--更改系统及部件的版本号
-------------------------------------------------------------------------------------------------------
--系统版本号
--部件版本号
Commit;