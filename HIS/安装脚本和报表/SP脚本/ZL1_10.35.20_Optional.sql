--[连续升级]1
--[管理工具版本号]10.35.0
--本脚本支持从ZLHIS+ v10.35.10 升级到 v10.35.20
--请以系统所有者登录PLSQL并执行下列脚本
--脚本执行后，请手工升级导出报表
Define n_System=100;
-------------------------------------------------------------------------------
-------------------------------------------------------------------------------
--96519:李业庆,2016-07-27,可用数量处理
--109023:李业庆,2017-05-15,96519问题可用数量修正脚本药品库存表批次未处理空值修改
--基于10.35修正可用数量，35.0中对可用数量进行了调整，需要重新计算
--这次调整默认流通部分填单下可用数量，修正可用数量时流通部分所有未审核的出库单都计算可用数量
--需要在新增/删除单据时减少/增加可用数量的单据类型
----1.发药/发料单据(收费处方，记账单，记账表)，排除已标记为停止发药的
----2.领用、其他出库、自制/协定入库中原料出库那笔、盘点单中盘亏那笔
----3.移库中出的那笔(单据 = 6 And 入出系数 = -1 And 记录状态 = 1)
----4.移库申请冲销单（原入库那笔的冲销记录，单据 = 6 And Mod(记录状态, 3) = 2 And 入出系数 = 1)
----5.入库业务中的退库单（单据 = 1 And 发药方式 = 1）
Create Or Replace Procedure Zl_Optional_可用数量修正 Is
  Cursor c_Data Is
  Select 库房id, 药品id, 批次, Sum(Nvl(实际数量, 0)) As 实际数量
  From (Select a.库房id, a.药品id, Nvl(a.批次, 0) As 批次,
                Case
                  When a.单据 In (8, 9, 10) And Nvl(a.发药方式, -999) <> -1 Then
                  a.实际数量 * Nvl(a.付数, 1)
                  When a.单据 In (2, 3, 7, 11, 12) And a.入出系数 = -1 Then
                  a.实际数量
                  When a.单据 = 6 And a.入出系数 = -1 And a.记录状态 = 1 Then
                  a.实际数量
                  When a.单据 = 6 And a.入出系数 = 1 And Mod(a.记录状态, 3) = 2 Then
                  -1 * a.实际数量
                  When a.单据 = 1 And a.发药方式 = 1 Then
                  -1 * a.实际数量
                End As 实际数量
        From 药品收发记录 A
        Where a.库房ID is not null and a.审核日期 Is Null And ((a.单据 In (8, 9, 10) And Nvl(a.发药方式, -999) <> -1) Or
              (a.单据 In (2, 3, 7, 11, 12) And a.入出系数 = -1) Or (a.单据 = 6 And a.入出系数 = -1 And a.记录状态 = 1) Or
              (a.单据 = 6 And a.入出系数 = 1 And Mod(a.记录状态, 3) = 2) Or (a.单据 = 1 And a.发药方式 = 1)) And Exists
              (Select 1 From 药品规格 B Where a.药品id = b.药品id))
  Group By 库房id, 药品id, 批次
  Order By 库房id, 药品id, 批次;
Begin
  --先更新可用数量=实际数量
  Update 药品库存 A
  Set a.可用数量 = a.实际数量
  Where a.性质 = 1 And a.库房id In (Select Distinct 部门id
                                From 部门性质说明
                                Where 工作性质 In ('西药库', '成药库', '中药库', '西药房', '成药房', '中药房', '制剂室')) And Exists
   (Select 1 From 药品规格 B Where a.药品id = b.药品id);

  --再根据未发数据更新可用数量
  For r_Data In c_Data Loop
    Update 药品库存
    Set 可用数量 = 实际数量 - r_Data.实际数量
    Where 性质 = 1 And 库房id = r_Data.库房id And 药品id = r_Data.药品id And Nvl(批次, 0) = r_Data.批次;
  
    If Sql%RowCount = 0 Then
      Insert Into 药品库存
        (库房id, 药品id, 批次, 性质, 可用数量, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 批准文号, 零售价, 平均成本价)
        Select r_Data.库房id, r_Data.药品id, r_Data.批次, 1, -1 * r_Data.实际数量, 上次供应商id, 成本价, 上次批号, 上次生产日期, 上次产地, 上次批号, 上次售价,
               成本价
        From 药品规格
        Where 药品id = r_Data.药品id;
    End If;
  End Loop;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Optional_可用数量修正;
/





---------------------------------------------------------------------------------------------------
--更改系统及部件的版本号
-------------------------------------------------------------------------------------------------------
--系统版本号
--部件版本号
Commit;