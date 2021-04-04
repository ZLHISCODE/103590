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
--139591:李业庆,2019-04-18,药品卫材库存和价格批次为空和0的问题处理
Declare
  n_可用数量 药品库存.可用数量%Type;
  n_数量     药品库存.实际数量%Type;
  n_金额     药品库存.实际金额%Type;
  n_差价     药品库存.实际差价%Type;
  n_时价售价 药品库存.零售价%Type;
  n_成本价   药品库存.平均成本价%Type;
  n_Count    Number(18) := 0;
Begin
  --1.库房不分批，只有一个批次，且批次=null
  --修正库存记录
  --写更新日志
  Update Zlupgradeconfig Set 内容 = Null Where 项目 = User || '_药品库存批次修正_20190312_1';
  If Sql%NotFound Then
    Insert Into Zlupgradeconfig (项目, 内容) Values (User || '_药品库存批次修正_20190312_1', Null);
  End If;

  Update 药品库存
  Set 批次 = 0
  Where 性质 = 1 And 批次 Is Null And
        (库房id, 药品id) In (Select b.库房id, b.药品id
                         From 药品库存 B,
                              (Select a.药品id, a.库房id
                                From 药品库存 A
                                Where a.性质 = 1 And Zl_Fun_Getbatchpro(a.库房id, a.药品id) = 0
                                Group By a.库房id, a.药品id
                                Having Count(Nvl(a.批次, 0)) = 1) A
                         Where b.性质 = 1 And b.库房id = a.库房id And b.药品id = a.药品id And b.批次 Is Null);

  Commit;

  Update Zlupgradeconfig Set 内容 = '已处理批次为null库存' Where 项目 = User || '_药品库存批次修正_20190312_1';
  Commit;

  --2.库房不分批，至少有2个批次，可能既有批次为null的，也有批次=0的  
  --修正药品库存  
  --写更新日志
  Update Zlupgradeconfig Set 内容 = Null Where 项目 = User || '_药品库存批次修正_20190312_2';
  If Sql%NotFound Then
    Insert Into Zlupgradeconfig (项目, 内容) Values (User || '_药品库存批次修正_20190312_2', Null);
  End If;
  For r_库存批次调整 In (Select Distinct b.库房id, b.药品id, Nvl(c.是否变价, 0) As 是否时价, Nvl(d.上次售价, e.现价) As 时价售价, d.成本价
                   From 药品库存 B, 收费项目目录 C, 药品规格 D, 收费价目 E,
                        (Select a.药品id, a.库房id
                          From 药品库存 A, 药品规格 B
                          Where a.性质 = 1 And a.药品id = b.药品id And Zl_Fun_Getbatchpro(a.库房id, b.药品id) = 0
                          Group By a.库房id, a.药品id
                          Having Count(Nvl(a.批次, 0)) > 1) A
                   Where b.性质 = 1 And b.库房id = a.库房id And b.药品id = a.药品id And c.Id = a.药品id And d.药品id = c.Id And
                         e.收费细目id = c.Id And Sysdate Between e.执行日期 And Nvl(e.终止日期, Sysdate)
                   Union All
                   Select Distinct b.库房id, b.药品id, Nvl(c.是否变价, 0) As 是否时价, Nvl(d.上次售价, e.现价) As 时价售价, d.成本价
                   From 药品库存 B, 收费项目目录 C, 材料特性 D, 收费价目 E,
                        (Select a.药品id, a.库房id
                          From 药品库存 A, 材料特性 B
                          Where a.性质 = 1 And a.药品id = b.材料id And Zl_Fun_Getbatchpro(a.库房id, b.材料id) = 0
                          Group By a.库房id, a.药品id
                          Having Count(Nvl(a.批次, 0)) > 1) A
                   Where b.性质 = 1 And b.库房id = a.库房id And b.药品id = a.药品id And c.Id = a.药品id And d.材料id = c.Id And
                         e.收费细目id = c.Id And Sysdate Between e.执行日期 And Nvl(e.终止日期, Sysdate)
                   Order By 库房id, 药品id) Loop
  
    --库存中合并批次的数量，金额，差价，并重算零售价（时价）和平均成本价，合并后批次=0
    Select Sum(Nvl(可用数量, 0)), Sum(Nvl(实际数量, 0)), Sum(Nvl(实际金额, 0)), Sum(Nvl(实际差价, 0))
    Into n_可用数量, n_数量, n_金额, n_差价
    From 药品库存
    Where 性质 = 1 And 库房id = r_库存批次调整.库房id And 药品id = r_库存批次调整.药品id And Nvl(批次, 0) = 0;
  
    --计算时价售价        
    If r_库存批次调整.是否时价 = 1 Then
      If n_数量 <> 0 Then
        n_时价售价 := n_金额 / n_数量;
      End If;
    
      If n_数量 = 0 Or Nvl(n_时价售价, 0) <= 0 Then
        n_时价售价 := r_库存批次调整.时价售价;
      End If;
    End If;
  
    --计算成本价
    If n_数量 <> 0 Then
      n_成本价 := (n_金额 - n_差价) / n_数量;
    End If;
  
    If n_数量 = 0 Or Nvl(n_成本价, 0) <= 0 Then
      n_成本价 := r_库存批次调整.成本价;
    End If;
  
    --更新批次=0的记录
    Update 药品库存
    Set 可用数量 = n_可用数量, 实际数量 = n_数量, 实际金额 = n_金额, 实际差价 = n_差价, 零售价 = Decode(r_库存批次调整.是否时价, 1, n_时价售价, Null),
        平均成本价 = n_成本价
    Where 性质 = 1 And 库房id = r_库存批次调整.库房id And 药品id = r_库存批次调整.药品id And 批次 = 0;
  End Loop;

  --删除批次=null的记录
  Delete From 药品库存 A Where a.性质 = 1 And a.批次 Is Null And Zl_Fun_Getbatchpro(a.库房id, a.药品id) = 0;

  Commit;

  Update Zlupgradeconfig Set 内容 = '已处理批次为0和null库存' Where 项目 = User || '_药品库存批次修正_20190312_2';
  Commit;

  --写更新日志
  Update Zlupgradeconfig Set 内容 = Null Where 项目 = User || '_药品库存批次修正_20190312_3';
  If Sql%NotFound Then
    Insert Into Zlupgradeconfig (项目, 内容) Values (User || '_药品库存批次修正_20190312_3', Null);
  End If;

  --3.删除价格表中批次为空的记录
  Delete From 药品价格记录 Where 批次 Is Null;
  Commit;

  --4.根据库存记录批次=0的数据检查对应的价格表
  For r_价格调整 In (Select a.库房id, a.药品id, a.批次, Nvl(c.是否变价, 0) As 时价, Nvl(a.零售价, 0) As 零售价, a.平均成本价
                 From 药品库存 A, 收费项目目录 C
                 Where a.药品id = c.Id And a.性质 = 1 And a.批次 = 0
                 Order By a.库房id, a.药品id) Loop
  
    --处理时价售价
    If r_价格调整.时价 = 1 Then
      Begin
        Select Count(ID)
        Into n_Count
        From 药品价格记录
        Where 价格类型 = 1 And 记录状态 = 1 And 库房id = r_价格调整.库房id And 药品id = r_价格调整.药品id And 批次 = 0;
      Exception
        When Others Then
          n_Count := 0;
      End;
    
      --如果批次=0的生效的价格存在2条或以上，则删除只保留1条
      If n_Count > 1 Then
        Delete From 药品价格记录
        Where 价格类型 = 1 And 记录状态 = 1 And 库房id = r_价格调整.库房id And 药品id = r_价格调整.药品id And 批次 = 0 And
              ID < (Select Max(ID)
                    From 药品价格记录
                    Where 价格类型 = 1 And 记录状态 = 1 And 库房id = r_价格调整.库房id And 药品id = r_价格调整.药品id And 批次 = 0);
      End If;
    
      --停止价格表中的价格，并新产生批次=0的价格
      Update 药品价格记录
      Set 终止日期 = Sysdate - 1 / 24 / 60 / 60, 记录状态 = 2
      Where 库房id = r_价格调整.库房id And 药品id = r_价格调整.药品id And 批次 = 0 And 记录状态 = 1 And 价格类型 = 1;
    
      --产生时价售价价格
      Insert Into 药品价格记录
        (ID, 原价id, 价格类型, 药品id, 库房id, 批次, 原价, 现价, 供药单位id, 批号, 效期, 产地, 灭菌效期, 发票号, 发票日期, 发票金额, 应付款变动, 执行日期, 终止日期, 记录状态,
         调价类型, 调价说明, 调价人, 调价汇总号, 收发id)
        Select 药品价格记录_Id.Nextval, Null, 1, 药品id, 库房id, 批次, 0, r_价格调整.零售价, 上次供应商id, 上次批号, 效期, 上次产地, 灭菌效期, Null, Null,
               Null, Null, Sysdate, To_Date('3000-01-01', 'yyyy-mm-dd'), 1, 0, '批次合并', 'ZLHIS', Null, Null
        From 药品库存
        Where 性质 = 1 And 库房id = r_价格调整.库房id And 药品id = r_价格调整.药品id And 批次 = 0;
    End If;
  
    --处理成本价
    Begin
      Select Count(ID)
      Into n_Count
      From 药品价格记录
      Where 价格类型 = 2 And 记录状态 = 1 And 库房id = r_价格调整.库房id And 药品id = r_价格调整.药品id And 批次 = 0;
    Exception
      When Others Then
        n_Count := 0;
    End;
  
    --如果批次=0的生效的价格存在2条或以上，则删除只保留1条
    If n_Count > 1 Then
      Delete From 药品价格记录
      Where 价格类型 = 2 And 记录状态 = 1 And 库房id = r_价格调整.库房id And 药品id = r_价格调整.药品id And 批次 = 0 And
            ID < (Select Max(ID)
                  From 药品价格记录
                  Where 价格类型 = 2 And 记录状态 = 1 And 库房id = r_价格调整.库房id And 药品id = r_价格调整.药品id And 批次 = 0);
    End If;
  
    --停止价格表中的价格，并新产生批次=0的价格
    Update 药品价格记录
    Set 终止日期 = Sysdate - 1 / 24 / 60 / 60, 记录状态 = 2
    Where 库房id = r_价格调整.库房id And 药品id = r_价格调整.药品id And 批次 = 0 And 记录状态 = 1 And 价格类型 = 2;
  
    --产生时价售价价格
    Insert Into 药品价格记录
      (ID, 原价id, 价格类型, 药品id, 库房id, 批次, 原价, 现价, 供药单位id, 批号, 效期, 产地, 灭菌效期, 发票号, 发票日期, 发票金额, 应付款变动, 执行日期, 终止日期, 记录状态, 调价类型,
       调价说明, 调价人, 调价汇总号, 收发id)
      Select 药品价格记录_Id.Nextval, Null, 2, 药品id, 库房id, 批次, 0, r_价格调整.平均成本价, 上次供应商id, 上次批号, 效期, 上次产地, 灭菌效期, Null, Null,
             Null, Null, Sysdate, To_Date('3000-01-01', 'yyyy-mm-dd'), 1, 0, '批次合并', 'ZLHIS', Null, Null
      From 药品库存
      Where 性质 = 1 And 库房id = r_价格调整.库房id And 药品id = r_价格调整.药品id And 批次 = 0;
  End Loop;
  Commit;

  Update Zlupgradeconfig Set 内容 = '已处理批次为0价格' Where 项目 = User || '_药品库存批次修正_20190312_3';
  Commit;

  --5.处理价格表中可能有多个现生效的批次=0但没有库存记录的价格
  --写更新日志
  Update Zlupgradeconfig Set 内容 = Null Where 项目 = User || '_药品库存批次修正_20190312_4';
  If Sql%NotFound Then
    Insert Into Zlupgradeconfig (项目, 内容) Values (User || '_药品库存批次修正_20190312_4', Null);
  End If;
  For r_无库存价格 In (Select a.库房id, a.药品id, a.价格类型
                  From 药品价格记录 A
                  Where a.记录状态 = 1 And a.批次 = 0 And Not Exists
                   (Select 1
                         From 药品库存 B
                         Where b.性质 = 1 And b.库房id = a.库房id And b.药品id = a.药品id And b.批次 = a.批次)
                  Group By a.库房id, a.药品id, a.价格类型
                  Having Count(a.批次) > 1
                  Order By a.价格类型, a.库房id, a.药品id) Loop
  
    --删除多余的价格，只保留1个
    Delete From 药品价格记录
    Where 价格类型 = r_无库存价格.价格类型 And 记录状态 = 1 And 库房id = r_无库存价格.库房id And 药品id = r_无库存价格.药品id And 批次 = 0 And
          ID < (Select Max(ID)
                From 药品价格记录
                Where 价格类型 = r_无库存价格.价格类型 And 记录状态 = 1 And 库房id = r_无库存价格.库房id And 药品id = r_无库存价格.药品id And 批次 = 0);
  End Loop;
  Commit;

  Update Zlupgradeconfig Set 内容 = '已处理多个批次为0价格' Where 项目 = User || '_药品库存批次修正_20190312_4';
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
--139585:秦龙,2019-04-18,处理批次为空的情况
Create Or Replace Procedure Zl_药品收发记录_Adjust
(
  药品id_In   In Number, --药品ID,为0时检查所有预调价
  调价方式_In In Number := 0 --0-检查售价和成本价预调价,1-只检查售价预调价,2-只检查成本价预调价
) As
  Classid          Number(18); --入出类别
  v_Billno         药品收发记录.No%Type; --调价单号
  Adjustdate       Date; --调价时间
  n_批次           Number(18);
  n_现价           收费价目.现价%Type;
  n_原价           收费价目.原价%Type;
  n_序号           Number(8);
  n_原价id         收费价目.原价id%Type;
  n_零售金额       药品库存.实际金额%Type;
  n_收发id         药品收发记录.Id%Type;
  n_流通金额小数   Number;
  n_Stockid        药品收发记录.库房id%Type;
  n_入出类别id     药品收发记录.入出类别id%Type;
  n_入出系数       药品收发记录.入出系数%Type;
  n_价格id         收费价目.Id%Type;
  n_无库存调价模式 Number(1) := 0;
  n_分批属性       Number(1) := 0;
  n_消息调用       Number(1) := 0;
  --定价售价，时价售价预调价记录
  --价格类型：0-定价售价,1-时价售价
  Cursor c_Priceadjust Is
    Select 0 As 价格类型, p.Id As 价格id, p.原价id, p.执行日期, p.原价, p.现价, i.药品id, s.库房id As 库房id, Nvl(s.批次, 0) As 批次, s.上次批号 As 批号,
           s.效期, s.上次产地 As 产地, s.上次供应商id As 供应商id, Nvl(s.实际数量, 0) As 实际数量, s.上次扣率 As 扣率, Nvl(s.实际金额, 0) As 实际金额,
           Nvl(s.实际差价, 0) As 实际差价, Nvl(s.零售价, 0) As 零售价, s.平均成本价, s.Rowid As 库存记录, s.灭菌效期, s.批准文号, s.上次生产日期 As 生产日期,
           p.调价汇总号
    From 收费价目 P, 药品规格 I, 收费项目目录 A, 药品库存 S
    Where i.药品id = a.Id And p.收费细目id = i.药品id And i.药品id = s.药品id(+) And s.性质(+) = 1 And Nvl(a.是否变价, 0) = 0 And
          Sysdate Between p.执行日期 And Nvl(p.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And Nvl(p.变动原因, 0) = 0 And
          p.收费细目id = Decode(药品id_In, 0, p.收费细目id, 药品id_In) And 调价方式_In In (0, 1)
    Union All
    Select 价格类型, p.Id As 价格id, p.原价id, p.执行日期, p.原价, p.现价, i.药品id, p.库房id As 库房id, Nvl(p.批次, 0) As 批次, p.批号 As 批号, p.效期,
           p.产地 As 产地, s.上次供应商id As 供应商id, Nvl(s.实际数量, 0) As 实际数量, s.上次扣率 As 扣率, Nvl(s.实际金额, 0) As 实际金额,
           Nvl(s.实际差价, 0) As 实际差价, Nvl(s.零售价, 0) As 零售价, s.平均成本价, s.Rowid As 库存记录, s.灭菌效期, s.批准文号, s.上次生产日期 As 生产日期,
           p.调价汇总号
    From 药品价格记录 P, 药品规格 I, 收费项目目录 A, 药品库存 S
    Where i.药品id = a.Id And p.药品id = i.药品id And p.库房id = s.库房id(+) And p.药品id = s.药品id(+) And
          Nvl(p.批次, 0) = Nvl(s.批次(+), 0) And s.性质(+) = 1 And Sysdate Between p.执行日期 And
          Nvl(p.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And Nvl(p.记录状态, 0) = 0 And
          p.药品id = Decode(药品id_In, 0, p.药品id, 药品id_In) And 价格类型 = 1 And 调价方式_In In (0, 1)
    Order By 价格类型, 药品id, 批次, 库房id;

  r_Priceadjust c_Priceadjust%RowType;

  --成本价预调价记录
  Cursor c_Costadjust Is
    Select 价格类型, p.Id As 价格id, p.原价id, p.执行日期, p.原价, p.现价, i.药品id, p.库房id As 库房id, Nvl(p.批次, 0) As 批次, p.批号 As 批号, p.效期,
           p.产地 As 产地, s.上次供应商id As 供应商id, Nvl(s.实际数量, 0) As 实际数量, s.上次扣率 As 扣率, Nvl(s.实际金额, 0) As 实际金额,
           Nvl(s.实际差价, 0) As 实际差价, Nvl(s.零售价, 0) As 零售价, s.平均成本价, s.Rowid As 库存记录, s.灭菌效期, s.批准文号, s.上次生产日期 As 生产日期,
           p.调价汇总号
    From 药品价格记录 P, 药品规格 I, 收费项目目录 A, 药品库存 S
    Where i.药品id = a.Id And p.药品id = i.药品id And p.库房id = s.库房id(+) And p.药品id = s.药品id(+) And
          Nvl(p.批次, 0) = Nvl(s.批次(+), 0) And s.性质(+) = 1 And Sysdate Between p.执行日期 And
          Nvl(p.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And Nvl(p.记录状态, 0) = 0 And
          p.药品id = Decode(药品id_In, 0, p.药品id, 药品id_In) And 价格类型 = 2 And 调价方式_In In (0, 2)
    Order By 药品id, 批次, 库房id;

  r_Costadjust c_Costadjust%RowType;

  --当前生效的价格，用于无库存调价
  Cursor c_Nostockadjust
  (
    Drugid_In 药品价格记录.药品id%Type,
    Type_In   药品价格记录.价格类型%Type
  ) Is
    Select a.价格类型, a.Id As 价格id, a.原价, a.现价, a.药品id, a.库房id, nvl(a.批次,0) as 批次, a.供药单位id, a.批号, a.效期, a.产地
    From 药品价格记录 A
    Where Sysdate Between a.执行日期 And Nvl(a.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And a.记录状态 = 1 And
          a.药品id = Drugid_In And a.价格类型 = Type_In And a.库房id Is Not Null
    Order By a.库房id, a.药品id, a.批次;

  r_Nostockadjust c_Nostockadjust%RowType;
Begin
  --取流通业务精度位数
  --类别:1-药品 2-卫材
  --内容：2-零售价 4-金额
  --单位：药品:1-售价 5-金额单位
  Select 精度 Into n_流通金额小数 From 药品卫材精度 Where 类别 = 1 And 内容 = 4 And 单位 = 5;

  --取入出类别ID
  Select 类别id Into Classid From 药品单据性质 Where 单据 = 13;

  Adjustdate := Sysdate;

  --售价调价处理
  If 调价方式_In = 0 Or 调价方式_In = 1 Then
  
    n_序号 := 0;
  
    --取调价NO取
    Select Nextno(147) Into v_Billno From Dual;
  
    For r_Priceadjust In c_Priceadjust Loop
      If r_Priceadjust.库房id Is Not Null Then
        --有库房id正常调价
      
        --取分批属性
        n_分批属性 := Zl_Fun_Getbatchpro(r_Priceadjust.库房id, r_Priceadjust.药品id);
      
        --产生调价盈亏记录的条件：1.要有库存记录，2.分批属性和库存批次一致
        If r_Priceadjust.库存记录 Is Not Null And ((n_分批属性 = 1 And r_Priceadjust.批次 > 0) Or
           (n_分批属性 = 0 And r_Priceadjust.批次 = 0)) Then
          n_序号 := n_序号 + 1;
        
          Select 药品收发记录_Id.Nextval Into n_收发id From Dual;
        
          n_原价 := r_Priceadjust.原价;
        
          --时价调价，如果原价和当前库存不一致，则以当前库存为准
          If r_Priceadjust.价格类型 = 1 And r_Priceadjust.原价 <> r_Priceadjust.零售价 And r_Priceadjust.库存记录 Is Not Null Then
            n_原价 := r_Priceadjust.零售价;
          End If;
        
          n_零售金额 := Round(r_Priceadjust.现价 * r_Priceadjust.实际数量, n_流通金额小数) - Round(n_原价 * r_Priceadjust.实际数量, n_流通金额小数);
        
          n_价格id := r_Priceadjust.价格id;
          If r_Priceadjust.价格类型 = 1 Then
            Select ID
            Into n_价格id
            From 收费价目
            Where Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And 收费细目id = r_Priceadjust.药品id;
          End If;
        
          --产生调价影响记录
          Insert Into 药品收发记录
            (ID, 记录状态, 单据, NO, 序号, 入出类别id, 药品id, 批次, 批号, 效期, 产地, 付数, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 扣率, 零售金额, 差价, 摘要, 填制人,
             填制日期, 库房id, 入出系数, 价格id, 审核人, 审核日期, 单量, 频次, 供药单位id)
          Values
            (n_收发id, 1, 13, v_Billno, n_序号, Classid, r_Priceadjust.药品id, r_Priceadjust.批次, r_Priceadjust.批号,
             r_Priceadjust.效期, r_Priceadjust.产地, 1, r_Priceadjust.实际数量, 0, n_原价, 0, r_Priceadjust.现价, r_Priceadjust.扣率,
             n_零售金额, n_零售金额, '药品调价', Zl_Username, Adjustdate, r_Priceadjust.库房id, 1, n_价格id, Zl_Username, Adjustdate,
             r_Priceadjust.实际金额, r_Priceadjust.实际差价, r_Priceadjust.供应商id);
        
          Zl_未审药品记录_Insert(n_收发id);
        
          --更新药品库存，无库存不执行
          If r_Priceadjust.库存记录 Is Not Null Then
            Zl_药品库存_Update(n_收发id, 2, 0);
          End If;
        End If;
      
        --更新原价格信息
        If r_Priceadjust.价格类型 = 1 Then
          Update 药品价格记录 Set 记录状态 = 2 Where ID = r_Priceadjust.原价id;
        End If;
      
        --时价调价更新价格表中的信息
        If r_Priceadjust.价格类型 = 1 Then
          --更新当前价格信息
          If r_Priceadjust.库存记录 Is Not Null Then
            Update 药品价格记录
            Set 批号 = r_Priceadjust.批号, 效期 = r_Priceadjust.效期, 产地 = r_Priceadjust.产地, 灭菌效期 = r_Priceadjust.灭菌效期,
                供药单位id = r_Priceadjust.供应商id, 原价 = n_原价, 收发id = n_收发id, 记录状态 = 1
            Where ID = r_Priceadjust.价格id;
          Else
            --无库存时只更新记录状态，收发id
            Update 药品价格记录 Set 收发id = n_收发id, 记录状态 = 1 Where ID = r_Priceadjust.价格id;
          End If;
        End If;
      
        --更新批号对照表售价
        If r_Priceadjust.价格类型 = 1 Then
          --如果是时价，则更新该药品批次对应的价格
          Update 药品批号对照
          Set 售价 = r_Priceadjust.现价
          Where 药品id = r_Priceadjust.药品id And Nvl(批次, 0) = r_Priceadjust.批次 And 售价 <> r_Priceadjust.现价;
        End If;
      
        --消息处理
        --定价只调用一次消息，时价可多次调用
        If (r_Priceadjust.价格类型 = 0 And n_消息调用 = 0) Or r_Priceadjust.价格类型 = 1 Then
          n_消息调用 := 1;
          b_Message.Zlhis_Drug_009(r_Priceadjust.价格id, r_Priceadjust.价格类型);
        End If;
      Else
        --无库存调价模式，价格表中该药品所有生效的价格都要按无库存调价时的价格调整
      
        If r_Priceadjust.价格类型 = 1 Then
          --更新原价格信息
          Update 药品价格记录 Set 记录状态 = 2 Where ID = r_Priceadjust.原价id;
        
          --更新现价格状态
          Update 药品价格记录 Set 记录状态 = 1 Where ID = r_Priceadjust.价格id;
        End If;
      
        n_无库存调价模式 := 0;
        For r_Nostockadjust In c_Nostockadjust(r_Priceadjust.药品id, 1) Loop
          If r_Priceadjust.现价 <> r_Nostockadjust.现价 Then
            Zl_药品价格记录_Stop(1, r_Nostockadjust.库房id, r_Nostockadjust.药品id, r_Nostockadjust.批次,
                           Adjustdate - 2 / 24 / 60 / 60);
            Zl_药品价格记录_Insert(1, 1, r_Nostockadjust.库房id, r_Nostockadjust.药品id, r_Nostockadjust.批次, Null,
                             r_Priceadjust.现价, Adjustdate - 1 / 24 / 60 / 60, '药品调价', Zl_Username, r_Priceadjust.调价汇总号,
                             r_Nostockadjust.供药单位id, r_Nostockadjust.批号, r_Nostockadjust.效期, r_Nostockadjust.产地);
            n_无库存调价模式 := 1;
          End If;
        End Loop;
        If n_无库存调价模式 = 1 Then
          Zl_药品收发记录_Adjust(r_Priceadjust.药品id, 1);
        End If;
      End If;
    
      --更新规格价格
      If r_Priceadjust.现价 <> r_Priceadjust.原价 Then
        Update 药品规格
        Set 上次售价 = r_Priceadjust.现价
        Where 药品id = r_Priceadjust.药品id And 上次售价 <> r_Priceadjust.现价;
      End If;
    
      If r_Priceadjust.价格类型 = 0 Then
        n_价格id := r_Priceadjust.价格id;
      Else
        Begin
          Select ID
          Into n_价格id
          From 收费价目
          Where Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And Nvl(变动原因, 0) = 0 And
                收费细目id = r_Priceadjust.药品id;
        Exception
          When Others Then
            n_价格id := 0;
        End;
      End If;
    
      If n_价格id > 0 Then
        Update 收费价目 Set 变动原因 = 1 Where Nvl(变动原因, 0) = 0 And ID = n_价格id;
      End If;
    
      --更新批号对照表售价
      If r_Priceadjust.价格类型 = 0 Then
        --如果是定价，则更新该药品对应的所有批次的售价
        Update 药品批号对照
        Set 售价 = r_Priceadjust.现价
        Where 药品id = r_Priceadjust.药品id And 售价 <> r_Priceadjust.现价;
      End If;
    End Loop;
  End If;

  --成本价调价处理
  If 调价方式_In = 0 Or 调价方式_In = 2 Then
  
    n_序号    := 0;
    n_Stockid := 0;
  
    Select b.Id, b.系数
    Into n_入出类别id, n_入出系数
    From 药品单据性质 A, 药品入出类别 B
    Where a.类别id = b.Id And a.单据 = 5 And Rownum < 2;
  
    v_Billno := Nextno(25, n_Stockid);
  
    For r_Costadjust In c_Costadjust Loop
      If r_Costadjust.库房id Is Not Null Then
        --有库房id正常调价
      
        --取分批属性
        n_分批属性 := Zl_Fun_Getbatchpro(r_Costadjust.库房id, r_Costadjust.药品id);
      
        --产生调价盈亏记录的条件：1.要有库存记录，2.分批属性和库存批次一致
        If r_Costadjust.库存记录 Is Not Null And ((n_分批属性 = 1 And r_Costadjust.批次 > 0) Or
           (n_分批属性 = 0 And r_Costadjust.批次 = 0)) Then
          n_序号 := n_序号 + 1;
        
          --产生库存差价调整单
          Select 药品收发记录_Id.Nextval Into n_收发id From Dual;
        
          --如果原价和当前库存不一致，则以当前库存为准
          n_原价 := r_Costadjust.原价;
          If r_Costadjust.原价 <> r_Costadjust.平均成本价 And r_Costadjust.库存记录 Is Not Null Then
            n_原价 := r_Costadjust.平均成本价;
          End If;
        
          n_零售金额 := Round(n_原价 * r_Costadjust.实际数量, n_流通金额小数) - Round(r_Costadjust.现价 * r_Costadjust.实际数量, n_流通金额小数);
        
          Insert Into 药品收发记录
            (ID, 记录状态, 单据, NO, 序号, 库房id, 入出类别id, 供药单位id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 填写数量, 实际数量, 零售价, 零售金额, 成本价, 成本金额,
             差价, 摘要, 填制人, 填制日期, 审核人, 审核日期, 生产日期, 批准文号, 单量, 发药方式, 扣率, 灭菌效期)
          Values
            (n_收发id, 1, 5, v_Billno, n_序号, r_Costadjust.库房id, n_入出类别id, r_Costadjust.供应商id, n_入出系数, r_Costadjust.药品id,
             r_Costadjust.批次, r_Costadjust.产地, r_Costadjust.批号, r_Costadjust.效期, r_Costadjust.实际数量, 0, r_Costadjust.实际金额,
             0, r_Costadjust.实际差价, 0, n_零售金额, '成本价调价', Zl_Username, Adjustdate, Zl_Username, Adjustdate,
             r_Costadjust.生产日期, r_Costadjust.批准文号, r_Costadjust.现价, 1, n_原价, r_Costadjust.灭菌效期);
        
          Zl_未审药品记录_Insert(n_收发id);
        
          Zl_药品库存_Update(n_收发id, 2, 0);
        
          --更新当前价格信息
          Update 药品价格记录
          Set 批号 = r_Costadjust.批号, 效期 = r_Costadjust.效期, 产地 = r_Costadjust.产地, 灭菌效期 = r_Costadjust.灭菌效期,
              供药单位id = r_Costadjust.供应商id, 原价 = n_原价, 收发id = n_收发id, 记录状态 = 1
          Where ID = r_Costadjust.价格id;
        Else
          --无库存时只更新记录状态
          Update 药品价格记录 Set 记录状态 = 1 Where ID = r_Costadjust.价格id;
        End If;
      
        --更新原价格信息
        Update 药品价格记录 Set 记录状态 = 2 Where ID = r_Costadjust.原价id;
      
        --更新批号对照表成本价
        Update 药品批号对照
        Set 成本价 = r_Costadjust.现价
        Where 药品id = r_Costadjust.药品id And Nvl(批次, 0) = r_Costadjust.批次 And 成本价 <> r_Costadjust.现价;
      Else
        --无库存调价模式，价格表中该药品所有生效的价格都要按无库存调价时的价格调整
      
        --更新原价格信息
        Update 药品价格记录 Set 记录状态 = 2 Where ID = r_Costadjust.原价id;
      
        --更新现价格状态
        Update 药品价格记录 Set 记录状态 = 1 Where ID = r_Costadjust.价格id;
      
        n_无库存调价模式 := 0;
        For r_Nostockadjust In c_Nostockadjust(r_Costadjust.药品id, 2) Loop
          If r_Costadjust.现价 <> r_Nostockadjust.现价 Then
            Zl_药品价格记录_Stop(2, r_Nostockadjust.库房id, r_Nostockadjust.药品id, r_Nostockadjust.批次,
                           Adjustdate - 2 / 24 / 60 / 60);
            Zl_药品价格记录_Insert(1, 2, r_Nostockadjust.库房id, r_Nostockadjust.药品id, r_Nostockadjust.批次, Null,
                             r_Costadjust.现价, Adjustdate - 1 / 24 / 60 / 60, '成本价调价', Zl_Username, r_Costadjust.调价汇总号,
                             r_Nostockadjust.供药单位id, r_Nostockadjust.批号, r_Nostockadjust.效期, r_Nostockadjust.产地);
            n_无库存调价模式 := 1;
          End If;
        End Loop;
        If n_无库存调价模式 = 1 Then
          Zl_药品收发记录_Adjust(r_Costadjust.药品id, 2);
        End If;
      End If;
    
      --更新规格价格
      If r_Costadjust.原价 <> r_Costadjust.现价 Then
        Update 药品规格
        Set 成本价 = r_Costadjust.现价
        Where 药品id = r_Costadjust.药品id And 成本价 <> r_Costadjust.现价;
      End If;
    
      --消息处理
      b_Message.Zlhis_Drug_007(r_Costadjust.价格id);
    End Loop;
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品收发记录_Adjust;
/

--139589:刘涛,2019-04-18,防止批次为空特殊处理
Create Or Replace Procedure Zl_药品库存_Update
(
  Id_In       In 药品收发记录.Id%Type,
  业务类型_In In Number := 0,
  入出类型_In In Number := 0,
  操作类型_In In Number := 0
) Is
  --功能：
  --      根据业务类型处理库存表，处理业务药品所有业务和卫材发料业务，卫材内部流通业务不处理
  --id_in  需要处理收发记录单
  --业务类型_in  业务类型，0-新增、1-删除、2-审核、3-冲销
  --入出类型_in  0-入库，1-出库
  --操作类型：根据业务情况确定
  ----外购入库中，表示财务审核：  0-不是财务审核，1-财务审核，目前只有外购入库有财务审核
  ----申领，移库中冲销流程：0-正常冲销流程，1-申请，审核冲销流程

  n_可用数量 药品库存.实际数量%Type;
  n_实际数量 药品库存.实际数量%Type;
  n_零售金额 药品库存.实际金额%Type;
  n_差价     药品库存.实际差价%Type;
  n_时价     Number(1);
  n_成本价   药品收发记录.成本价%Type;
  n_零售价   药品库存.零售价%Type;

  n_库存数量     药品库存.实际数量%Type;
  n_库存平均价   药品库存.平均成本价%Type;
  n_库存售价     药品库存.零售价%Type;
  n_总数量       药品收发记录.实际数量%Type;
  n_总成本价     药品收发记录.成本价%Type;
  n_总售价       药品收发记录.零售价%Type;
  n_库房分批     药品规格.药库分批%Type;
  n_申请冲销     Number(1);
  n_更新库存     Number(1) := 0;
  v_现价         药品收发记录.零售价%Type;
  v_执行新价格   Number(1) := 0;
  n_有库存       Number(1) := 0;
  v_审核日期     药品收发记录.审核日期%Type;
  n_价格更新     Number(1) := 0;
  n_新增时价售价 Number(1) := 0;
  n_新增成本价   Number(1) := 0;
  n_分批属性     Number(1) := 0; --0-分批属性不符合，1-分批属性符合
  n_新增库存     Number(1) := 0; --0-更新库存，1-新增库存
  n_重算价格     Number(1) := 0; --0-不重算价格,1-重算价格
  --业务明细数据，把库存数据更新需要的数据都列出来
  Cursor c_Detail Is
    Select a.Id, a.记录状态, a.单据, a.No, a.序号, a.库房id, a.供药单位id, a.入出类别id, a.对方部门id, a.入出系数, Nvl(a.发药方式, 0) As 发药方式, a.药品id,
           Nvl(a.批次, 0) 批次, a.产地, a.原产地, a.批号, a.生产日期, a.效期, a.付数, Nvl(a.填写数量, 0) As 填写数量, a.实际数量, a.成本价, a.成本金额, a.扣率,
           a.零售价, Nvl(a.零售金额, 0) As 零售金额, Nvl(a.差价, 0) As 差价, a.配药人, a.配药日期, a.审核人, a.审核日期, a.灭菌日期, a.灭菌效期, a.批准文号,
           a.商品条码, a.内部条码, Nvl(b.是否变价, 0) As 是否变价, a.单量, a.频次, a.摘要, Nvl(a.费用id, 0) As 费用id,
           Decode(a.批次, Null, 1, 0) 空批次
    From 药品收发记录 A, 收费项目目录 B
    Where a.药品id = b.Id And a.Id = Id_In;

  r_Detail c_Detail%RowType;
Begin
  For r_Detail In c_Detail Loop
  	If r_Detail.空批次 = 1 Then
      Update 药品收发记录 Set 批次 = 0 Where NO = r_Detail.No And 单据 = r_Detail.单据 And 序号 = r_Detail.序号;
    End If;

    If Zl_Fun_Getbatchpro(r_Detail.库房id, r_Detail.药品id) = 1 Then
      If r_Detail.批次 > 0 Then
        n_分批属性 := 1;
      Else
        n_分批属性 := 0;
      End If;
    Else
      If r_Detail.批次 = 0 Then
        n_分批属性 := 1;
      Else
        n_分批属性 := 0;
      End If;
    End If;
  
    n_实际数量 := r_Detail.入出系数 * r_Detail.实际数量 * Nvl(r_Detail.付数, 1);
    If n_实际数量 Is Null Then
      n_实际数量 := 0;
    End If;
    n_可用数量 := 0;
    n_零售价   := r_Detail.零售价;
    If r_Detail.单据 = 12 Then
      n_成本价 := r_Detail.单量;
    Else
      n_成本价 := r_Detail.成本价;
    End If;
    n_零售金额 := r_Detail.入出系数 * r_Detail.零售金额;
    n_差价     := r_Detail.入出系数 * r_Detail.差价;
  
    --先取库存和单据的数量和成本价
    Begin
      Select Nvl(实际数量, 0)
      Into n_库存数量
      From 药品库存
      Where 性质 = 1 And 库房id = r_Detail.库房id And 药品id = r_Detail.药品id And Nvl(批次, 0) = r_Detail.批次;
    Exception
      When Others Then
        n_库存数量 := 0;
    End;
  
    n_库存平均价 := Zl_Fun_Getoutcost(r_Detail.药品id, r_Detail.批次, r_Detail.库房id);
    n_库存售价   := Zl_Fun_Getoutprice(r_Detail.药品id, r_Detail.批次, r_Detail.库房id);
  
    --时价药品都需要更新库存表零售价字段
    If r_Detail.是否变价 = 1 Then
      n_时价 := 1;
    Else
      n_时价 := 0;
    End If;
  
    --特殊业务处理库存
    --包含业务--5-差价调整；13-调价变动
    --单据5，13都是业务类型_in，2-审核、入出类型_in  0-入库
    If r_Detail.单据 = 5 Or r_Detail.单据 = 13 Then
      --这种类型的单据收发记录成本价字段不是保存的真正成本价而是存储的其他数据
      If r_Detail.单据 = 5 Then
        If r_Detail.填写数量 <> 0 Then
          n_零售价 := Nvl(r_Detail.零售价, 0) / r_Detail.填写数量;
        Else
          n_零售价 := 0;
        End If;
        --审核
        If r_Detail.记录状态 = 1 Then
          --差价调整发药方式=0；主动调价、退货、发药产生的调价修正发药方式=1
          n_成本价 := r_Detail.单量;
        Else
          --冲销 还原原始成本价
          Begin
            --成本价=(金额-差价)/数量
            n_成本价 := (Nvl(r_Detail.零售价, 0) - Nvl(r_Detail.成本价, 0)) / r_Detail.填写数量;
          Exception
            When Others Then
              Select 成本价 Into n_成本价 From 药品规格 Where 药品id = r_Detail.药品id;
          End;
        End If;
      Else
        n_成本价 := Nvl(r_Detail.单量, 0) - Nvl(r_Detail.频次, 0);
      End If;
    
      If r_Detail.单据 = 5 Then
        --单据=5 的成本价修正记录 平均成本价不需要重算，因为保存了最新价格的
        If r_Detail.摘要 = '外购退库差价误差自动修正' Or r_Detail.摘要 = '财务审核价格变动修正' Then
          --这一步肯定是外购退库，外购退库只更新成本价,且肯定有库存
          Update 药品库存
          Set 实际差价 = Nvl(实际差价, 0) + n_差价
          Where 性质 = 1 And 库房id = r_Detail.库房id And 药品id = r_Detail.药品id And Nvl(批次, 0) = r_Detail.批次;
        Else
          Update 药品库存
          Set 平均成本价 = n_成本价, 上次采购价 = n_成本价, 实际金额 = Nvl(实际金额, 0) + n_零售金额, 实际差价 = Nvl(实际差价, 0) + n_差价
          Where 性质 = 1 And 库房id = r_Detail.库房id And 药品id = r_Detail.药品id And Nvl(批次, 0) = r_Detail.批次;
          If Sql%NotFound Then
            Insert Into 药品库存
              (库房id, 药品id, 批次, 性质, 实际差价, 上次批号, 效期, 上次产地, 原产地, 上次供应商id, 上次生产日期, 批准文号, 实际金额, 上次采购价, 平均成本价)
            Values
              (r_Detail.库房id, r_Detail.药品id, r_Detail.批次, 1, n_差价, r_Detail.批号, r_Detail.效期, r_Detail.产地, r_Detail.原产地,
               r_Detail.供药单位id, r_Detail.生产日期, r_Detail.批准文号, n_零售金额, n_成本价, n_成本价);
          
            Insert Into 药品入库信息
              (药品id, 库房id, 批次, 入库日期)
              Select r_Detail.药品id, r_Detail.库房id, r_Detail.批次, r_Detail.审核日期
              From Dual
              Where Not Exists (Select 1
                     From 药品入库信息
                     Where 药品id = r_Detail.药品id And 库房id = r_Detail.库房id And 批次 = r_Detail.批次);
          End If;
        End If;
      Elsif r_Detail.单据 = 13 Then
        --单据=13 的售价修正记录 同步更新的金额和差价，所以不需要重算平均成本价
        If r_Detail.费用id = 0 Then
          Update 药品库存
          Set 零售价 = Decode(n_时价, 1, n_零售价, Null), 实际金额 = Nvl(实际金额, 0) + n_零售金额, 实际差价 = Nvl(实际差价, 0) + n_差价
          Where 性质 = 1 And 库房id = r_Detail.库房id And 药品id = r_Detail.药品id And Nvl(批次, 0) = r_Detail.批次;
        Else
          Update 药品库存
          Set 实际金额 = Nvl(实际金额, 0) + n_零售金额, 实际差价 = Nvl(实际差价, 0) + n_差价
          Where 性质 = 1 And 库房id = r_Detail.库房id And 药品id = r_Detail.药品id And Nvl(批次, 0) = r_Detail.批次;
        End If;
      
        If Sql%RowCount = 0 Then
          Insert Into 药品库存
            (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 零售价)
          Values
            (r_Detail.库房id, r_Detail.药品id, r_Detail.批次, 1, 0, 0, n_零售金额, n_零售金额, Decode(n_时价, 1, n_零售价, Null));
        
          Insert Into 药品入库信息
            (药品id, 库房id, 批次, 入库日期)
            Select r_Detail.药品id, r_Detail.库房id, r_Detail.批次, r_Detail.审核日期
            From Dual
            Where Not Exists (Select 1
                   From 药品入库信息
                   Where 药品id = r_Detail.药品id And 库房id = r_Detail.库房id And 批次 = r_Detail.批次);
        End If;
      End If;
    Else
      --一般业务处理库存
      --包含业务--1-外购入库；2-自制入库；3-协药入库；4-其他入库；6-库房移出；
      --7-部门领用；8-收费处方发药；9-记帐单处方发药；10-记帐表处方发药；11-其他出库；
      --12-盘点；14-药品盘点记录单
      --21-材料其他出库；24-收费处方发料；25-记帐单处方发料；26-记帐表处方发料
      If 业务类型_In = 0 Or 业务类型_In = 1 Then
        --新增，删除
        If r_Detail.单据 = 8 Or r_Detail.单据 = 9 Or r_Detail.单据 = 10 Or r_Detail.单据 = 21 Or r_Detail.单据 = 24 Or
           r_Detail.单据 = 25 Or r_Detail.单据 = 26 Or r_Detail.单据 = 7 Or r_Detail.单据 = 11 Or
           ((r_Detail.单据 = 2 Or r_Detail.单据 = 3 Or r_Detail.单据 = 12) And r_Detail.入出系数 = -1) Or
           (r_Detail.单据 = 1 And r_Detail.发药方式 = 1) Or (r_Detail.单据 = 6 And r_Detail.入出系数 = -1 And r_Detail.记录状态 = 1) Or
           (r_Detail.单据 = 6 And Mod(r_Detail.记录状态, 3) = 2 And r_Detail.入出系数 = 1) Then
          --需要在新增/删除单据时减少/增加可用数量的单据类型
          ----1.发药/发料单据(收费处方，记账单，记账表)
          ----2.普通出库（领用、其他出库、移库中出的那笔(r_Detail.单据 = 6 And r_Detail.入出系数 = -1 And r_Detail.记录状态 = 1)、盘点单中盘亏那笔）
          ----3.退库单（r_Detail.单据 = 1 And r_Detail.发药方式 = 1）
          ----4.移库申请冲销单（原入库那笔的冲销记录，r_Detail.单据 = 6 And Mod(r_Detail.记录状态, 3) = 2 And r_Detail.入出系数 = 1）
        
          --新增，删除单据时，因为没有审核所以只处理数量不处理金额和差价
        
          If 业务类型_In = 0 Then
            --新增时正常处理可用数量
            n_可用数量 := n_实际数量;
          Else
            --删除时按相反数计算可用数量
            n_可用数量 := -1 * n_实际数量;
          End If;
        
          n_实际数量 := 0;
          n_零售金额 := 0;
          n_差价     := 0;
        
          --处理库存
          Update 药品库存
          Set 可用数量 = 可用数量 + n_可用数量
          Where 药品id = r_Detail.药品id And 库房id = r_Detail.库房id And Nvl(批次, 0) = r_Detail.批次 And 性质 = 1;
        
          n_更新库存 := 1;
        End If;
      Elsif 业务类型_In = 2 Then
        --审核
        --10.35开始，理论上所有的出库类单据在审核时都不再处理可用数量
        If r_Detail.单据 = 8 Or r_Detail.单据 = 9 Or r_Detail.单据 = 10 Or r_Detail.单据 = 21 Or r_Detail.单据 = 24 Or
           r_Detail.单据 = 25 Or r_Detail.单据 = 26 Or r_Detail.单据 = 7 Or r_Detail.单据 = 11 Or
           ((r_Detail.单据 = 2 Or r_Detail.单据 = 3 Or r_Detail.单据 = 12) And r_Detail.入出系数 = -1) Or
           (r_Detail.单据 = 1 And r_Detail.发药方式 = 1) Or (r_Detail.单据 = 6 And r_Detail.入出系数 = -1 And r_Detail.记录状态 = 1) Or
           (r_Detail.单据 = 6 And Mod(r_Detail.记录状态, 3) = 2 And r_Detail.入出系数 = 1) Then
          n_可用数量 := 0;
        Else
          n_可用数量 := n_实际数量;
        End If;
      
        --处理库存
        If 入出类型_In = 0 Then
          --入库审核
          Update 药品库存
          Set 可用数量 = Nvl(可用数量, 0) + n_可用数量, 实际数量 = Nvl(实际数量, 0) + n_实际数量, 实际金额 = Nvl(实际金额, 0) + n_零售金额,
              实际差价 = Nvl(实际差价, 0) + n_差价, 上次供应商id = r_Detail.供药单位id,
              上次采购价 = Decode(r_Detail.单据, 1, Decode(r_Detail.发药方式, 1, 上次采购价, n_成本价), n_成本价),
              上次批号 = Nvl(r_Detail.批号, 上次批号), 上次生产日期 = Nvl(r_Detail.生产日期, 上次生产日期), 上次产地 = Nvl(r_Detail.产地, 上次产地),
              原产地 = Nvl(r_Detail.原产地, 原产地), 灭菌效期 = Nvl(r_Detail.灭菌效期, 灭菌效期), 效期 = Nvl(r_Detail.效期, 效期),
              批准文号 = Nvl(r_Detail.批准文号, 批准文号), 上次扣率 = Decode(r_Detail.单据, 12, 上次扣率, r_Detail.扣率),
              商品条码 = Nvl(r_Detail.商品条码, 商品条码), 内部条码 = Nvl(r_Detail.内部条码, 内部条码)
          Where 库房id = r_Detail.库房id And 药品id = r_Detail.药品id And Nvl(批次, 0) = r_Detail.批次 And 性质 = 1;
        Else
          --出库审核，只需要下数量和金额
          Update 药品库存
          Set 可用数量 = Nvl(可用数量, 0) + n_可用数量, 实际数量 = Nvl(实际数量, 0) + n_实际数量, 实际金额 = Nvl(实际金额, 0) + n_零售金额,
              实际差价 = Nvl(实际差价, 0) + n_差价, 平均成本价 = Decode(平均成本价, Null, n_成本价, 平均成本价),
              上次采购价 = Decode(上次采购价, Null, n_成本价, 上次采购价)
          Where 库房id = r_Detail.库房id And 药品id = r_Detail.药品id And Nvl(批次, 0) = r_Detail.批次 And 性质 = 1;
        End If;
      
        n_更新库存 := 1;
      Elsif 业务类型_In = 3 Then
        --冲销
        If r_Detail.单据 = 8 Or r_Detail.单据 = 9 Or r_Detail.单据 = 10 Or r_Detail.单据 = 21 Or r_Detail.单据 = 24 Or
           r_Detail.单据 = 25 Or r_Detail.单据 = 26 Then
          --发药/发料单退药/退料时同时又产生了未发单据，所以就不处理可用数量
          n_可用数量 := 0;
        Elsif r_Detail.单据 = 6 And Mod(r_Detail.记录状态, 3) = 2 And r_Detail.入出系数 = 1 Then
          --药库单的冲销单据，要判断是否需要申请
          n_申请冲销 := 操作类型_In;
          If n_申请冲销 = 0 Then
            --不需要申请的在冲销时处理可用数量
            n_可用数量 := n_实际数量;
          Else
            --需要申请的，已经在申请时处理了可用数量
            n_可用数量 := 0;
          End If;
        Else
          n_可用数量 := n_实际数量;
        End If;
      
        --处理库存
        If 入出类型_In = 0 Then
          --出库单据冲销需要将入库库房数据都更新
          Update 药品库存
          Set 可用数量 = Nvl(可用数量, 0) + n_可用数量, 实际数量 = Nvl(实际数量, 0) + n_实际数量, 实际金额 = Nvl(实际金额, 0) + n_零售金额,
              实际差价 = Nvl(实际差价, 0) + n_差价, 上次供应商id = r_Detail.供药单位id,
              上次采购价 = Decode(r_Detail.单据, 1, Decode(r_Detail.发药方式, 1, 上次采购价, n_成本价), 上次采购价),
              上次批号 = Nvl(r_Detail.批号, 上次批号), 上次生产日期 = Nvl(r_Detail.生产日期, 上次生产日期), 上次产地 = Nvl(r_Detail.产地, 上次产地),
              原产地 = Nvl(r_Detail.原产地, 原产地), 灭菌效期 = Nvl(r_Detail.灭菌效期, 灭菌效期), 效期 = Nvl(r_Detail.效期, 效期),
              批准文号 = Nvl(r_Detail.批准文号, 批准文号), 上次扣率 = Decode(r_Detail.单据, 12, 上次扣率, r_Detail.扣率),
              商品条码 = Nvl(r_Detail.商品条码, 商品条码), 内部条码 = Nvl(r_Detail.内部条码, 内部条码)
          Where 库房id = r_Detail.库房id And 药品id = r_Detail.药品id And Nvl(批次, 0) = r_Detail.批次 And 性质 = 1;
        Else
          --入库单据冲销只需要下数量和金额
          Update 药品库存
          Set 可用数量 = Nvl(可用数量, 0) + n_可用数量, 实际数量 = Nvl(实际数量, 0) + n_实际数量, 实际金额 = Nvl(实际金额, 0) + n_零售金额,
              实际差价 = Nvl(实际差价, 0) + n_差价
          Where 库房id = r_Detail.库房id And 药品id = r_Detail.药品id And Nvl(批次, 0) = r_Detail.批次 And 性质 = 1;
        End If;
      
        n_更新库存 := 1;
      End If;
    
      --新增/删除/审核/冲销业务时，库存表未找到数据则需要产生库存表所有信息
      If Sql%RowCount = 0 And n_更新库存 = 1 Then
        --入库业务取界面价格（之前已经取过了），出库业务冲销业务取最新价格
        If 业务类型_In = 3 Or 入出类型_In = 1 Then
          --取最新成本价
          v_现价 := Zl_Fun_Getoutcost(r_Detail.药品id, r_Detail.批次, r_Detail.库房id, n_成本价);
          If v_现价 Is Not Null Then
            n_成本价 := v_现价;
          End If;
        
          --时价售价取最新价格
          If r_Detail.是否变价 = 1 Then
            v_现价 := Zl_Fun_Getoutprice(r_Detail.药品id, r_Detail.批次, r_Detail.库房id, n_零售价);
            If v_现价 Is Not Null Then
              n_零售价 := v_现价;
            End If;
          End If;
        End If;
      
        --新增库存
        n_新增库存 := 1;
        Insert Into 药品库存
          (库房id, 药品id, 批次, 效期, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 原产地, 灭菌效期, 批准文号, 零售价,
           上次扣率, 商品条码, 内部条码, 平均成本价)
        Values
          (r_Detail.库房id, r_Detail.药品id, r_Detail.批次, r_Detail.效期, 1, n_可用数量, n_实际数量, n_零售金额, n_差价, r_Detail.供药单位id,
           n_成本价, r_Detail.批号, r_Detail.生产日期, r_Detail.产地, r_Detail.原产地, r_Detail.灭菌效期, r_Detail.批准文号,
           Decode(n_时价, 1, n_零售价, Null), r_Detail.扣率, r_Detail.商品条码, r_Detail.内部条码, n_成本价);
      
        Insert Into 药品入库信息
          (药品id, 库房id, 批次, 入库日期)
          Select r_Detail.药品id, r_Detail.库房id, r_Detail.批次, r_Detail.审核日期
          From Dual
          Where Not Exists (Select 1
                 From 药品入库信息
                 Where 药品id = r_Detail.药品id And 库房id = r_Detail.库房id And 批次 = r_Detail.批次);
      End If;
    
      --重算平均成本价，入库审核需要重算平均成本价和零售价，注意只限于不分批药品，分批药品不用重算（确保和之前库存的数据一致）
      --只有更新库存状态需要重新计算价格，新增库存状态不用重算
      If 入出类型_In = 0 And 业务类型_In = 2 And r_Detail.批次 = 0 And n_新增库存 <> 1 Then
        --按总金额/总数量方式计算平均成本价而不用（金额-差价）/数量是为了数据的准确性
        n_重算价格 := 1;
        n_总数量   := (n_库存数量 + n_实际数量);
        If n_总数量 <> 0 Then
          n_总成本价 := (n_库存数量 * n_库存平均价 + n_实际数量 * n_成本价) / n_总数量;
        
          If n_总成本价 < 0 Then
            n_总成本价 := n_成本价;
          End If;
        
          Update 药品库存
          Set 平均成本价 = n_总成本价
          Where 性质 = 1 And 库房id = r_Detail.库房id And 药品id = r_Detail.药品id And Nvl(批次, 0) = r_Detail.批次;
        
          --更新时价零售价
          If n_时价 = 1 Then
            n_总售价 := (n_库存数量 * n_库存售价 + n_实际数量 * n_零售价) / n_总数量;
            Update 药品库存
            Set 零售价 = n_总售价
            Where 性质 = 1 And 库房id = r_Detail.库房id And 药品id = r_Detail.药品id And Nvl(批次, 0) = r_Detail.批次 And
                  Nvl(实际数量, 0) <> 0;
            If Sql%NotFound Then
              n_总售价 := n_零售价;
              Update 药品库存
              Set 零售价 = n_总售价
              Where 性质 = 1 And 库房id = r_Detail.库房id And 药品id = r_Detail.药品id And Nvl(批次, 0) = r_Detail.批次;
            End If;
          End If;
        End If;
      End If;
    
      --价格处理
      --分批属性正确时才进行价格处理
      If n_分批属性 = 1 Then
        --新增库存，或者更新库存并且重算了价格的情况下才处理价格
        If n_新增库存 = 1 Or (n_新增库存 = 0 And n_重算价格 = 1) Then
          --时价价格
          If r_Detail.是否变价 = 1 Then
            Zl_药品价格记录_Addnew(n_新增库存, 0, 1, r_Detail.库房id, r_Detail.药品id, r_Detail.批次, 0, Nvl(n_总售价, n_零售价), Sysdate,
                             '入库新增批次价格', Zl_Username, Null, r_Detail.供药单位id, r_Detail.批号, r_Detail.效期, r_Detail.产地, Null,
                             Null, Null, Null, Null, 1);
          End If;
        
          --成本价价格
          Zl_药品价格记录_Addnew(n_新增库存, 0, 2, r_Detail.库房id, r_Detail.药品id, r_Detail.批次, 0, Nvl(n_总成本价, n_成本价), Sysdate,
                           '入库新增批次价格', Zl_Username, Null, r_Detail.供药单位id, r_Detail.批号, r_Detail.效期, r_Detail.产地, Null,
                           Null, Null, Null, Null, 1);
        End If;
      End If;
    End If;
  
    --删除多余的库存数据，外购入库财务审核为了确保库存不变产生修正数据必须保证不删除库存
    If Not (r_Detail.单据 = 1 And 操作类型_In = 1) Then
      Delete From 药品库存
      Where 性质 = 1 And 库房id = r_Detail.库房id And 药品id = r_Detail.药品id And Nvl(批次, 0) = r_Detail.批次 And Nvl(可用数量, 0) = 0 And
            Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And Nvl(实际差价, 0) = 0;
    End If;
  
    Zl_药品库存_可用数量异常处理(r_Detail.库房id, r_Detail.药品id, r_Detail.批次);
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品库存_Update;
/

--139589:刘涛,2019-04-18,防止批次为空特殊处理
Create Or Replace Procedure Zl_药品其他入库_Insert
(
  No_In         In 药品收发记录.No%Type,
  序号_In       In 药品收发记录.序号%Type,
  库房id_In     In 药品收发记录.库房id%Type,
  入出类别id_In In 药品收发记录.入出类别id%Type,
  药品id_In     In 药品收发记录.药品id%Type,
  实际数量_In   In 药品收发记录.实际数量%Type,
  成本价_In     In 药品收发记录.成本价%Type,
  成本金额_In   In 药品收发记录.成本金额%Type,
  零售价_In     In 药品收发记录.零售价%Type,
  零售金额_In   In 药品收发记录.零售金额%Type,
  差价_In       In 药品收发记录.差价%Type,
  填制人_In     In 药品收发记录.填制人%Type,
  填制日期_In   In 药品收发记录.填制日期%Type,
  摘要_In       In 药品收发记录.摘要%Type := Null,
  产地_In       In 药品收发记录.产地%Type := Null,
  批号_In       In 药品收发记录.批号%Type := Null,
  效期_In       In 药品收发记录.效期%Type := Null,
  生产日期_In   In 药品收发记录.生产日期%Type := Null,
  批准文号_In   In 药品收发记录.批准文号%Type := Null,
  外观_In       In 药品收发记录.外观%Type := Null,
  金额差_In     In 药品收发记录.零售金额%Type := Null,
  原产地_In       In 药品收发记录.原产地%Type := Null,
  修改人_In     In 药品收发记录.修改人%Type,
  修改日期_In   In 药品收发记录.修改日期%Type
) Is
  v_Lngid    药品收发记录.Id%Type; --收发ID 
  v_入出系数 药品收发记录.入出系数%Type;
  v_批次     药品收发记录.批次%Type := 0; --批次 
  v_药库分批 Integer; --是否药库分批    1:分批;0：不分批 
  v_药房分批 Integer; --是否药库分批    1:分批;0：不分批 
  v_时价分批 Number(1);

Begin

  If Not 批准文号_In Is Null And Not 产地_In Is Null Then
    Update 药品生产商对照 Set 批准文号 = 批准文号_In Where 药品id = 药品id_In And 厂家名称 = 产地_In;
  End If;
  If Sql%RowCount = 0 And Not 产地_In Is Null And Not 批准文号_In Is Null Then
    Insert Into 药品生产商对照 (药品id, 厂家名称, 批准文号) Values (药品id_In, 产地_In, 批准文号_In);
  End If;

  v_入出系数 := 1;
  Select 药品收发记录_Id.Nextval Into v_Lngid From Dual;
  Select Nvl(药库分批, 0), Nvl(药房分批, 0) Into v_药库分批, v_药房分批 From 药品规格 Where 药品id = 药品id_In;

  If v_药房分批 = 0 Then
    If v_药库分批 = 1 Then
      Begin
        Select Distinct 0
        Into v_药库分批
        From 部门性质说明
        Where ((工作性质 Like '%药房') Or (工作性质 Like '制剂室')) And 部门id = 库房id_In;
      Exception
        When Others Then
          v_药库分批 := 1;
      End;
    
      If v_药库分批 = 1 Then
        v_批次 := Zl_Fun_Getbatchnum(药品id_In, 产地_In, 批号_In, 成本价_In, 零售价_In, v_Lngid, Null);
      End If;
    End If;
  Else
    v_批次 := Zl_Fun_Getbatchnum(药品id_In, 产地_In, 批号_In, 成本价_In, 零售价_In, v_Lngid, Null);
  End If;

  Select Nvl(是否变价, 0) Into v_时价分批 From 收费项目目录 Where ID = 药品id_In;

  If v_时价分批 = 1 And v_批次 > 0 Then
    v_时价分批 := 1;
  Else
    v_时价分批 := 0;
  End If;

  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 入出类别id, 入出系数, 药品id, 批次, 产地, 原产地,批号, 效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额, 差价, 摘要, 填制人,
     填制日期, 生产日期, 批准文号, 外观, 用法, 修改人, 修改日期)
  Values
    (v_Lngid, 1, 4, No_In, 序号_In, 库房id_In, 入出类别id_In, v_入出系数, 药品id_In, v_批次, 产地_In, 原产地_In,批号_In, 效期_In, 实际数量_In, 实际数量_In,
     成本价_In, 成本金额_In, 零售价_In, 零售金额_In, 差价_In, 摘要_In, 填制人_In, 填制日期_In, 生产日期_In, 批准文号_In, 外观_In,
     Decode(v_时价分批, 1, 金额差_In, Null), 修改人_In, 修改日期_In);
  
  Zl_未审药品记录_Insert(v_Lngid);
  
  --更新库存
  Zl_药品库存_Update(v_Lngid, 0);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品其他入库_Insert;
/

--139589:刘涛,2019-04-18,防止批次为空特殊处理
Create Or Replace Procedure Zl_协定入库_Strike
(
  No_In     In 药品收发记录.No%Type,
  审核人_In In 药品收发记录.审核人%Type
) Is
  Err_Isstriked Exception;
  

  Cursor c_药品收发记录 Is
    Select ID, 库房id, 入出类别id, 入出系数, 药品id, 填写数量, 批次, 实际数量, 成本价, 零售金额, 差价, 产地, 批号, 效期, 供药单位id, 生产日期, 批准文号
    From 药品收发记录 A
    Where NO = No_In And 单据 = 3 And 记录状态 = 2
	Order By 药品id,批次;
Begin
  Update 药品收发记录 Set 记录状态 = 3 Where NO = No_In And 单据 = 3 And 记录状态 = 1;

  If Sql%RowCount = 0 Then
    Raise Err_Isstriked;
  End If;
  
  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 入出类别id, 对方部门id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额, 差价, 摘要,
     填制人, 填制日期, 审核人, 审核日期, 费用id, 扣率, 供药单位id, 生产日期, 批准文号)
    Select 药品收发记录_Id.Nextval, 2, 单据, No_In, 序号, 库房id, 入出类别id, 对方部门id, 入出系数, 药品id, nvl(批次,0), 产地, 批号, 效期, -填写数量, -实际数量, 成本价,
           -成本金额, 零售价, -零售金额, -差价, 摘要, 审核人_In, Sysdate, 审核人_In, Sysdate, 费用id, 扣率, 供药单位id, 生产日期, 批准文号
    From 药品收发记录
    Where NO = No_In And 单据 = 3 And 记录状态 = 3;
  
  
  For v_药品收发记录 In c_药品收发记录 Loop
    --更改药品库存表的相应数据
    If v_药品收发记录.入出系数 = 1 Then
      Zl_药品库存_Update(v_药品收发记录.Id, 3, 1);
    Else
      Zl_药品库存_Update(v_药品收发记录.Id, 3, 0);
    End If;
  
    --处理调价后冲销
    Zl_药品收发记录_调价修正(v_药品收发记录.Id);
  End Loop;
Exception
  When Err_Isstriked Then
    Raise_Application_Error(-20101, '[ZLSOFT]该单据已经被他人冲销！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_协定入库_Strike;
/

--139589:刘涛,2019-04-18,防止批次为空特殊处理
Create Or Replace Procedure Zl_药品收发记录_部门退药
(
  Billid_In     In 药品收发记录.Id%Type,
  People_In     In 药品收发记录.审核人%Type,
  Date_In       In 药品收发记录.审核日期%Type,
  批号_In       In 药品库存.上次批号%Type := Null,
  效期_In       In 药品库存.效期%Type := Null,
  产地_In       In 药品库存.上次产地%Type := Null,
  退药数量_In   In 药品收发记录.实际数量%Type := Null,
  退药库房_In   In 药品收发记录.库房id%Type := Null,
  退药人_In     In 药品收发记录.领用人%Type := Null,
  Intdigit_In   In Number := 2,
  门诊_In       In Number := 2,
  汇总发药号_In In 药品收发记录.汇总发药号%Type := Null
) Is
  --只读变量
  Int记录状态   药品收发记录.记录状态%Type;
  Int执行状态   住院费用记录.执行状态%Type;
  Bln部分退药   Number;
  Lng入出类别id Number(18);
  Strno         药品收发记录.No%Type;
  Int单据       药品收发记录.单据%Type;
  Lng库房id     药品收发记录.库房id%Type;
  Lng药品id     药品收发记录.药品id%Type;
  Dbl实际数量   药品收发记录.实际数量%Type;
  Dbl实际金额   药品收发记录.零售金额%Type;
  Dbl实际成本   药品收发记录.成本金额%Type;
  Dbl实际差价   药品收发记录.差价%Type;
  Lng费用id     药品收发记录.费用id%Type;
  n_零售价      药品收发记录.零售价%Type;
  n_是否变价    Number;
  n_时价分批    Number;

  --20020731 Modified by zyb
  --处理退药时，分批核算性质改变后的处理
  Lng新批次 药品收发记录.批次%Type;
  Lng分批   药品规格.药房分批%Type;
  Lng批次   药品收发记录.批次%Type; --原批次

  Str批号        药品收发记录.批号%Type; --原批号
  Date效期       药品收发记录.效期%Type; --原效期
  n_上次供应商id 药品库存.上次供应商id%Type;
  n_上次采购价   药品库存.上次采购价%Type;
  v_上次产地     药品库存.上次产地%Type;
  v_原产地       药品库存.原产地%Type;
  d_上次生产日期 药品库存.上次生产日期%Type;
  v_批准文号     药品库存.批准文号%Type;

  n_记录性质   住院费用记录.记录性质%Type;
  v_收费类别   住院费用记录.收费类别%Type;
  n_付数       药品收发记录.付数%Type;
  n_原始数量   药品收发记录.实际数量%Type;
  v_冲销记录id 药品收发记录.Id%Type;
  Err_Custom Exception;
  v_Error    Varchar2(255);
  v_配药确认 药房配药控制.配药确认%Type;
  v_配药     药房配药控制.配药%Type;
  v_排队状态 Number(1);
  v_执行时间 药品收发记录.审核日期%Type;

Begin
  If 退药数量_In Is Not Null Then
    If 退药数量_In = 0 Then
      Return;
    End If;
  End If;

  --获取该收发记录的单据、药品ID、库房ID
  Select a.单据, a.No, a.库房id, a.药品id, a.费用id, a.入出类别id, a.记录状态, Nvl(a.批次, 0), a.批号, a.效期, a.供药单位id, a.产地, a.原产地, a.生产日期,
         a.批准文号, a.成本价, a.付数, Nvl(a.实际数量, 0) * Nvl(a.付数, 1) As 实际数量, a.零售价, Nvl(b.是否变价, 0) 是否变价
  Into Int单据, Strno, Lng库房id, Lng药品id, Lng费用id, Lng入出类别id, Int记录状态, Lng批次, Str批号, Date效期, n_上次供应商id, v_上次产地, v_原产地,
       d_上次生产日期, v_批准文号, n_上次采购价, n_付数, n_原始数量, n_零售价, n_是否变价
  From 药品收发记录 A, 收费项目目录 B
  Where a.药品id = b.Id And a.Id = Billid_In;

  Begin
    Select Nvl(配药确认, 0), Nvl(配药, 0)
    Into v_配药确认, v_配药
    From 药房配药控制
    Where 药房id = Lng库房id And Rownum = 1;
  
  Exception
    When Others Then
      v_配药确认 := 0;
      v_配药     := 0;
      Null;
  End;

  If v_配药确认 = 0 And v_配药 = 0 Then
    v_排队状态 := 2;
  Elsif v_配药确认 = 1 Then
    v_排队状态 := 0;
  Elsif v_配药 = 1 Then
    v_排队状态 := 1;
  End If;

  --获取该笔记录剩余未退数量、金额及差价
  --尽量避免金额及差价未出完的现象
  Select Sum(Nvl(实际数量, 0) * Nvl(付数, 1)), Sum(Nvl(零售金额, 0)), Sum(Nvl(成本金额, 0)), Sum(Nvl(差价, 0))
  Into Dbl实际数量, Dbl实际金额, Dbl实际成本, Dbl实际差价
  From 药品收发记录
  Where 审核人 Is Not Null And NO = Strno And 单据 = Int单据 And 序号 = (Select 序号 From 药品收发记录 Where ID = Billid_In);

  --如果允许退药数为零，表示已退药
  If Dbl实际数量 = 0 Then
    v_Error := '该单据已被其他操作员退药，请刷新后再试！';
    Raise Err_Custom;
  End If;
  If Nvl(退药数量_In, 0) > Dbl实际数量 Then
    v_Error := '该单据已被其他操作员部分退药，请刷新后再试！';
    Raise Err_Custom;
  End If;

  --获取该药品当前是否分批的信息
  Select Nvl(药房分批, 0) Into Lng分批 From 药品规格 Where 药品id = Lng药品id;
  --如果是部分退药，则重新计算零售金额及差价
  Bln部分退药 := 0;
  If Not (退药数量_In Is Null Or Nvl(退药数量_In, 0) = Dbl实际数量) Then
    Bln部分退药 := 1;
  End If;
  If Bln部分退药 = 1 Then
    Dbl实际金额 := Round(Dbl实际金额 * 退药数量_In / Dbl实际数量, Intdigit_In);
    Dbl实际成本 := Round(Dbl实际成本 * 退药数量_In / Dbl实际数量, Intdigit_In);
    Dbl实际差价 := Round(Dbl实际差价 * 退药数量_In / Dbl实际数量, Intdigit_In);
    Dbl实际数量 := 退药数量_In;
  End If;

  If n_原始数量 = 退药数量_In Then
    Dbl实际数量 := 退药数量_In / n_付数;
  Else
    n_付数 := 1;
  End If;

  --lng分批:0-不分批;1-分批;2-原分批，现不分批，按不分批处理;3-原不分批，现分批，产生新批次
  If Lng分批 = 0 And Lng批次 <> 0 Then
    --原分批，现不分批，按不分批处理
    Lng分批 := 2;
  Elsif Lng分批 <> 0 And Lng批次 = 0 Then
    --原不分批,现分批,产生新的批次，并在新产生的发药记录中使用
    Lng分批 := 3;
  Else
    If Lng批次 = 0 Then
      Lng分批 := 0;
    Else
      Lng分批 := 1;
    End If;
  End If;
  --判断是否时价分批
  If (Lng分批 = 1 Or Lng分批 = 3) And n_是否变价 = 1 Then
    n_时价分批 := 1;
  Else
    n_时价分批 := 0;
  End If;

  --记录状态的含义有所变化
  --冲销的记录状态        :iif(int记录状态=1,0,1)+1
  --被冲销的记录状态        :iif(int记录状态=1,0,1)+2
  --等待发药的记录状态    :iif(int记录状态=1,0,1)+3

  --产生冲销记录
  Select 药品收发记录_Id.Nextval Into v_冲销记录id From Dual;
  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 付数, 填写数量, 实际数量, 成本价, 成本金额, 扣率, 零售价, 零售金额,
     差价, 摘要, 填制人, 填制日期, 配药人, 审核人, 审核日期, 费用id, 单量, 频次, 用法, 发药窗口, 外观, 领用人, 供药单位id, 生产日期, 批准文号, 汇总发药号, 发药方式, 注册证号, 计划id,
     原产地)
    Select v_冲销记录id, Int记录状态 + Decode(Int记录状态, 1, 0, 1) + 1, Int单据, Strno, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, nvl(批次,0), 产地,
           批号, 效期, n_付数, -dbl实际数量, -dbl实际数量, 成本价, -dbl实际成本, 扣率, 零售价, -dbl实际金额, -dbl实际差价, 摘要, People_In, Date_In, 配药人,
           People_In, Date_In, 费用id, 单量, 频次, 用法, 发药窗口, 退药库房_In, 退药人_In, 供药单位id, 生产日期, 批准文号, 汇总发药号_In, 发药方式, 注册证号, 计划id,
           原产地
    From 药品收发记录
    Where ID = Billid_In;

  --如果是部分冲销，则付数填为1，实际数量为付数与实际数量的积
  --产生正常记录以供继续发药
  Select 药品收发记录_Id.Nextval Into Lng新批次 From Dual;
  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 付数, 填写数量, 实际数量, 成本价, 成本金额, 扣率, 零售价, 零售金额,
     差价, 摘要, 填制人, 填制日期, 配药人, 审核人, 审核日期, 费用id, 单量, 频次, 用法, 发药窗口, 供药单位id, 生产日期, 批准文号, 注册证号, 计划id, 原产地)
    Select Lng新批次, Int记录状态 + Decode(Int记录状态, 1, 0, 1) + 3, Int单据, Strno, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id,
           Decode(Lng分批, 1, nvl(批次,0), 3, Lng新批次, 0), Decode(Lng分批, 3, 产地_In, 1, 产地, 产地), Decode(Lng分批, 3, 批号_In, 批号),
           Decode(Lng分批, 3, 效期_In, 效期), n_付数, Dbl实际数量, Dbl实际数量, 成本价, Dbl实际成本, 扣率, 零售价, Dbl实际金额, Dbl实际差价, 摘要, 填制人, 填制日期,
           Null, Null, Null, 费用id, 单量, 频次, 用法, 发药窗口, 供药单位id, 生产日期, 批准文号, 注册证号, 计划id, 原产地
    
    From 药品收发记录
    Where ID = Billid_In;

  Zl_未审药品记录_Insert(Lng新批次);

  --更新费用记录的执行状态(0-未执行;1-完全执行;2-部分执行)
  Select Decode(Sum(Nvl(付数, 1) * 实际数量), Null, 0, 0, 0, 2)
  Into Int执行状态
  From 药品收发记录
  Where 单据 = Int单据 And NO = Strno And 费用id = Lng费用id And 审核人 Is Not Null;

  If 门诊_In = 1 Then
    Select 记录性质, 收费类别 Into n_记录性质, v_收费类别 From 门诊费用记录 Where ID = Lng费用id;
  Else
    Select 记录性质, 收费类别 Into n_记录性质, v_收费类别 From 住院费用记录 Where ID = Lng费用id;
  End If;

  If Int执行状态 = 0 Then
    If 门诊_In = 1 Then
      Update 门诊费用记录
      Set 执行状态 = Int执行状态, 执行人 = Null, 执行时间 = Null
      Where NO = Strno And
            序号 = (Select 序号 From 门诊费用记录 Where ID = (Select 费用id From 药品收发记录 Where ID = Billid_In)) And
            Mod(记录性质, 10) = n_记录性质 And 记录状态 <> 2 And 执行部门id = Lng库房id;
    Else
      Update 住院费用记录 Set 执行状态 = Int执行状态, 执行人 = Null, 执行时间 = Null Where ID = Lng费用id;
    End If;
  Else
    If 门诊_In = 1 Then
      Update 门诊费用记录
      Set 执行状态 = Int执行状态
      Where NO = Strno And
            序号 = (Select 序号 From 门诊费用记录 Where ID = (Select 费用id From 药品收发记录 Where ID = Billid_In)) And
            Mod(记录性质, 10) = n_记录性质 And 记录状态 <> 2 And 执行部门id = Lng库房id;
    Else
      Update 住院费用记录 Set 执行状态 = Int执行状态 Where ID = Lng费用id;
    End If;
  End If;

  --插入未发药品记录
  Begin
    If 门诊_In = 1 Then
      Insert Into 未发药品记录
        (单据, NO, 病人id, 主页id, 姓名, 优先级, 对方部门id, 库房id, 发药窗口, 填制日期, 已收费, 配药人, 打印状态, 未发数, 领药号, 排队状态)
        Select a.单据, a.No, a.病人id, Null, a.姓名, Nvl(b.优先级, 0) 优先级, a.对方部门id, a.库房id, a.发药窗口, a.填制日期, a.已收费, Null, 1, 1,
               a.产品合格证, v_排队状态
        From (Select b.单据, b.No, a.病人id, a.姓名, Decode(a.记录状态, 0, 0, 1) 已收费, b.对方部门id, b.库房id, b.发药窗口, b.填制日期, c.身份,
                      b.产品合格证
               From 门诊费用记录 A, 药品收发记录 B, 病人信息 C
               Where b.Id = Billid_In And a.Id = b.费用id + 0 And a.病人id = c.病人id(+)) A, 身份 B
        Where b.名称(+) = a.身份;
    Else
      Insert Into 未发药品记录
        (单据, NO, 病人id, 主页id, 姓名, 优先级, 对方部门id, 库房id, 发药窗口, 填制日期, 已收费, 配药人, 打印状态, 未发数, 领药号, 排队状态)
        Select a.单据, a.No, a.病人id, a.主页id, a.姓名, Nvl(b.优先级, 0) 优先级, a.对方部门id, a.库房id, a.发药窗口, a.填制日期, a.已收费, Null, 1, 1,
               a.产品合格证, v_排队状态
        From (Select b.单据, b.No, a.病人id, a.主页id, a.姓名, Decode(a.记录状态, 0, 0, 1) 已收费, b.对方部门id, b.库房id, b.发药窗口, b.填制日期,
                      c.身份, b.产品合格证
               From 住院费用记录 A, 药品收发记录 B, 病人信息 C
               Where b.Id = Billid_In And a.Id = b.费用id + 0 And a.病人id = c.病人id(+)) A, 身份 B
        Where b.名称(+) = a.身份;
    End If;
  
    --修改处方类型
    Zl_Prescription_Type_Update(Strno, n_记录性质, Lng药品id, v_收费类别);
  Exception
    When Others Then
      Null;
  End;

  --修改原记录为被冲销记录
  Update 药品收发记录 Set 记录状态 = Int记录状态 + Decode(Int记录状态, 1, 0, 1) + 2 Where ID = Billid_In;

  --修改药品库存(反冲库存)
  If Lng分批 <> 3 Then
    --正常单据需要将库存表实际数量和金额、差价还回去，如果库存表没有则在库存表插入数据
    Zl_药品库存_Update(v_冲销记录id, 3, 0);
  Else
    --原不分批，现在分批，直接在库存表产生新单据
    Insert Into 药品库存
      (库房id, 药品id, 批次, 效期, 性质, 实际数量, 实际金额, 实际差价, 零售价, 上次批号, 上次产地, 上次供应商id, 上次采购价, 上次生产日期, 批准文号, 平均成本价)
    Values
      (Lng库房id, Lng药品id, Lng新批次, 效期_In, 1, Dbl实际数量 * n_付数, Dbl实际金额, Dbl实际差价, Decode(n_时价分批, 1, n_零售价, Null), 批号_In,
       产地_In, n_上次供应商id, n_上次采购价, d_上次生产日期, v_批准文号, n_上次采购价);
  End If;

  Delete 药品库存
  Where 库房id + 0 = Lng库房id And 药品id = Lng药品id And 性质 = 1 And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
        Nvl(实际差价, 0) = 0;

  --处理调价修正
  Zl_药品收发记录_调价修正(v_冲销记录id);

  Begin
    --移动支付宝项目在发药后动态调用生成推送信息的过程
    Execute Immediate 'Begin zl_服务窗消息_发送(:1,:2); End;'
      Using 7, Billid_In || ',' || 退药数量_In || ',' || 门诊_In;
  Exception
    When Others Then
      Null;
  End;

  --消息处理，剩余全部退数量传0
  If Bln部分退药 = 1 Then
    b_Message.Zlhis_Drug_006(v_冲销记录id, Lng新批次, Dbl实际数量 * n_付数, Lng费用id);
  Else
    b_Message.Zlhis_Drug_006(v_冲销记录id, Lng新批次, 0, Lng费用id);
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品收发记录_部门退药;
/

--139589:刘涛,2019-04-18,防止批次为空特殊处理
Create Or Replace Procedure Zl_药品盘点_Strike
(
  No_In     In 药品收发记录.No%Type,
  审核人_In In 药品收发记录.审核人%Type
) Is
  Err_Isstriked Exception;
  Err_Isbatch Exception;
  v_Err_Msg     Varchar2(255);
  n_Batch_Count Number;
  n_药品id      药品收发记录.药品id%Type;

  Cursor c_药品收发记录 Is
    Select a.Id, a.实际数量, a.零售金额, a.差价, a.零售价, Nvl(b.是否变价, 0) As 是否变价, a.库房id, a.药品id, a.批次, a.批号, a.效期, a.产地, a.原产地, a.入出类别id,
           a.入出系数, a.单量, a.批准文号, a.供药单位id, a.生产日期
    From 药品收发记录 A, 收费项目目录 B
    Where a.药品id = b.Id And NO = No_In And 单据 = 12 And 记录状态 = 2
    Order By 药品id,批次;
Begin
  Update 药品收发记录 Set 记录状态 = 3 Where NO = No_In And 单据 = 12 And 记录状态 = 1;
  If Sql%RowCount = 0 Then
    Raise Err_Isstriked;
  End If;

  --主要针对原不分批现在分批的材料，不能对其审核 
  Select Count(*), Max(a.药品id)
  Into n_Batch_Count, n_药品id
  From 药品收发记录 A, 药品规格 B
  Where a.药品id = b.药品id And a.No = No_In And a.单据 = 12 And a.记录状态 = 3 And Nvl(a.批次, 0) = 0 And
        ((Nvl(b.药房分批, 0) = 1 And
        a.库房id Not In (Select 部门id From 部门性质说明 Where (工作性质 Like '%药房') Or (工作性质 Like '制剂室'))) Or Nvl(b.药房分批, 0) = 1);

  If n_Batch_Count > 0 Then
    Begin
      Select 编码 || '-' || 名称 Into v_Err_Msg From 收费项目目录 Where ID = n_药品id;
    Exception
      When Others Then
        Null;
    End;
    v_Err_Msg := '该单据中为:' || v_Err_Msg || Chr(10) || Chr(13) || '的药品,原来不分批,而现在分批，因此不能审核！';
    Raise Err_Isbatch;
  End If;
  
  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 入出类别id, 入出系数, 药品id, 批次, 产地, 原产地, 批号, 效期, 填写数量, 扣率, 实际数量, 成本价, 成本金额, 零售价, 零售金额, 差价, 摘要, 填制人,
     填制日期, 审核人, 审核日期, 频次, 单量, 批准文号, 供药单位id, 生产日期, 库房货位)
    Select 药品收发记录_Id.Nextval, 2, 单据, NO, 序号, 库房id, 入出类别id, 入出系数, a.药品id,
           Decode(Nvl(a.批次, 0), 0, 0, (Decode(Nvl(b.药库分批, 0), 0, 0, a.批次))), a.产地, a.原产地, 批号, 效期, 填写数量, a.扣率, -实际数量,
           a.成本价, 成本金额, 零售价, -零售金额, -差价, 摘要, 审核人_In, Sysdate, 审核人_In, Sysdate, 频次, 单量, a.批准文号, a.供药单位id, a.生产日期, 
		   a.库房货位
    From (Select * From 药品收发记录 Where NO = No_In And 单据 = 12 And 记录状态 = 3 Order By 药品id) A, 药品规格 B
    Where a.药品id = b.药品id;
  
  For v_药品收发记录 In c_药品收发记录 Loop
    --处理库存
    If v_药品收发记录.入出系数 = 1 Then
      Zl_药品库存_Update(v_药品收发记录.Id, 3, 1);
    Else
      Zl_药品库存_Update(v_药品收发记录.Id, 3, 0);
    End If;
  
    --处理调价后冲销
    Zl_药品收发记录_调价修正(v_药品收发记录.Id);
  End Loop;
Exception
  When Err_Isbatch Then
    Raise_Application_Error(-20102, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Err_Isstriked Then
    Raise_Application_Error(-20101, '[ZLSOFT]该单据已经被他人冲销！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品盘点_Strike;
/

--139589:刘涛,2019-04-18,防止批次为空特殊处理
Create Or Replace Procedure Zl_药品其他出库_Strike
(
  行次_In       In Integer,
  原记录状态_In In 药品收发记录.记录状态%Type,
  No_In         In 药品收发记录.No%Type,
  序号_In       In 药品收发记录.序号%Type,
  药品id_In     In 药品收发记录.药品id%Type,
  冲销数量_In   In 药品收发记录.实际数量%Type,
  填制人_In     In 药品收发记录.填制人%Type,
  填制日期_In   In 药品收发记录.填制日期%Type
) Is
  Err_Isstriked Exception;
  Err_Isoutstock Exception;
  Err_Isnonum Exception;
  Err_Isbatch Exception;
  v_Druginf Varchar2(50); --原不分批现在分批的药品信息

  v_库房id       药品收发记录.库房id%Type;
  v_入出类别id   药品收发记录.入出类别id%Type;
  v_产地         药品收发记录.产地%Type;
  v_原产地       药品收发记录.原产地%Type;
  v_批次         药品收发记录.批次%Type;
  v_批号         药品收发记录.批号%Type;
  v_效期         药品收发记录.效期%Type;
  v_成本价       药品收发记录.成本价%Type;
  v_成本金额     药品收发记录.成本金额%Type;
  v_扣率         药品收发记录.扣率%Type;
  v_零售价       药品收发记录.零售价%Type;
  v_零售金额     药品收发记录.零售金额%Type;
  v_差价         药品收发记录.差价%Type;
  v_摘要         药品收发记录.摘要%Type;
  v_剩余数量     药品收发记录.实际数量%Type;
  v_剩余成本金额 药品收发记录.成本金额%Type;
  v_剩余零售金额 药品收发记录.零售金额%Type;
  v_入出系数     药品收发记录.入出系数%Type;
  v_外调价       药品收发记录.单量%Type;
  v_外调单位     药品收发记录.发药窗口%Type;
  v_批准文号     药品收发记录.批准文号%Type;
  v_增值税率     药品收发记录.频次%Type;

  v_收发id 药品收发记录.Id%Type;
  Intdigit Number;

  n_上次供应商id 药品库存.上次供应商id%Type;
  d_上次生产日期 药品库存.上次生产日期%Type;
Begin
  --获取金额小数位数
  Select Nvl(精度, 2) Into Intdigit From 药品卫材精度 Where 性质 = 0 And 类别 = 1 And 内容 = 4 And 单位 = 5;

  If 行次_In = 1 Then
    Update 药品收发记录
    Set 记录状态 = Decode(原记录状态_In, 1, 3, 原记录状态_In + 3)
    Where NO = No_In And 单据 = 11 And 记录状态 = 原记录状态_In;
    If Sql%RowCount = 0 Then
      Raise Err_Isstriked;
    End If;
  End If;

  --主要针对原不分批现在分批的药品，不能对其冲销
  Begin
    Select Distinct '(' || i.编码 || ')' || Nvl(n.名称, i.名称) As 药品信息
    Into v_Druginf
    From 药品收发记录 A, 药品规格 B, 收费项目目录 I, 收费项目别名 N
    Where a.药品id = b.药品id And a.药品id = i.Id And a.药品id = n.收费细目id(+) And n.性质(+) = 3 And a.No = No_In And a.单据 = 11 And
          a.药品id + 0 = 药品id_In And Mod(a.记录状态, 3) = 0 And Nvl(a.批次, 0) = 0 And
          ((Nvl(b.药库分批, 0) = 1 And
          a.库房id Not In (Select 部门id From 部门性质说明 Where (工作性质 Like '%药房') Or (工作性质 Like '制剂室'))) Or Nvl(b.药房分批, 0) = 1) And
          Rownum = 1;
  Exception
    When Others Then
      v_Druginf := '';
  End;

  If v_Druginf Is Not Null Then
    Raise Err_Isbatch;
  End If;

  Select Sum(实际数量) As 剩余数量, Sum(成本金额) As 剩余成本金额, Sum(零售金额) As 剩余零售金额, 库房id, 入出类别id, 入出系数, Nvl(批次, 0) As 批次, 产地, 原产地, 批号, 效期, 成本价, 扣率, 零售价,
         摘要, 单量, 发药窗口, 批准文号, 供药单位id, 生产日期, To_Number(Trim(To_Char(Nvl(频次, '0'), '999999999999.0000'))) As 增值税率
  Into v_剩余数量, v_剩余成本金额, v_剩余零售金额, v_库房id, v_入出类别id, v_入出系数, v_批次, v_产地, v_原产地, v_批号, v_效期, v_成本价, v_扣率, v_零售价, v_摘要, v_外调价,
       v_外调单位, v_批准文号, n_上次供应商id, d_上次生产日期, v_增值税率
  From 药品收发记录
  Where NO = No_In And 单据 = 11 And 药品id = 药品id_In And 序号 = 序号_In
  Group By 库房id, 入出类别id, 入出系数, Nvl(批次, 0), 产地, 原产地, 批号, 效期, 成本价, 扣率, 零售价, 摘要, 单量, 发药窗口, 批准文号, 供药单位id, 生产日期,
           To_Number(Trim(To_Char(Nvl(频次, '0'), '999999999999.0000')));

  --冲销数量大于剩余数量，不允许
  If Abs(v_剩余数量) < Abs(冲销数量_In) Then
    Raise Err_Isnonum;
  End If;

  v_成本金额 := Round(冲销数量_In / v_剩余数量 * v_剩余成本金额, Intdigit);
  v_零售金额 := Round(冲销数量_In / v_剩余数量 * v_剩余零售金额, Intdigit);
  v_差价     := v_零售金额 - v_成本金额;

  Select 药品收发记录_Id.Nextval Into v_收发id From Dual;
  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 入出类别id, 入出系数, 药品id, 批次, 产地, 原产地, 批号, 效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额, 差价, 单量, 发药窗口, 摘要,
     填制人, 填制日期, 审核人, 审核日期, 批准文号, 供药单位id, 生产日期, 扣率, 频次)
  Values
    (v_收发id, Decode(原记录状态_In, 1, 2, 原记录状态_In + 2), 11, No_In, 序号_In, v_库房id, v_入出类别id, v_入出系数, 药品id_In, v_批次, v_产地,v_原产地, 
     v_批号, v_效期, -冲销数量_In, -冲销数量_In, v_成本价, -v_成本金额, v_零售价, -v_零售金额, -v_差价, v_外调价, v_外调单位, v_摘要, 填制人_In, 填制日期_In, 填制人_In,
     填制日期_In, v_批准文号, n_上次供应商id, d_上次生产日期, v_扣率, v_增值税率);
  
  --更新库存，出库冲销是入库
  Zl_药品库存_Update(v_收发id, 3, 0);

  --处理调价后冲销
  Zl_药品收发记录_调价修正(v_收发id);
Exception
  When Err_Isstriked Then
    Raise_Application_Error(-20101, '[ZLSOFT]该单据已经被他人冲销！[ZLSOFT]');
  When Err_Isbatch Then
    Raise_Application_Error(-20102, '[ZLSOFT]该单据中包含有一条原来不分批，现在分批的药品[' || v_Druginf || ']，不能冲销！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品其他出库_Strike;
/

--139589:刘涛,2019-04-18,防止批次为空特殊处理
Create Or Replace Procedure Zl_药品领用_Strike
(
  行次_In       In Integer,
  原记录状态_In In 药品收发记录.记录状态%Type,
  No_In         In 药品收发记录.No%Type,
  序号_In       In 药品收发记录.序号%Type,
  药品id_In     In 药品收发记录.药品id%Type,
  冲销数量_In   In 药品收发记录.实际数量%Type,
  填制人_In     In 药品收发记录.填制人%Type,
  填制日期_In   In 药品收发记录.填制日期%Type,
  冲销方式_In   In Integer := 0, --0－正常冲销方式；1－产生冲销申请单据；2－审核已产生的冲销申请单据  
  冲销原因_In   In 药品收发记录.冲销原因%Type := Null
) Is
  Err_Isstriked Exception;
  Err_Isoutstock Exception;
  Err_Isnonum Exception;
  Err_Isbatch Exception;
  v_Druginf Varchar2(50); --原不分批现在分批的药品信息 

  v_库房id       药品收发记录.库房id%Type;
  v_对方部门id   药品收发记录.对方部门id%Type;
  v_入出类别id   药品收发记录.入出类别id%Type;
  v_产地         药品收发记录.产地%Type;
  v_原产地       药品收发记录.原产地%Type;
  v_批次         药品收发记录.批次%Type;
  v_批号         药品收发记录.批号%Type;
  v_效期         药品收发记录.效期%Type;
  v_成本价       药品收发记录.成本价%Type;
  v_成本金额     药品收发记录.成本金额%Type;
  v_扣率         药品收发记录.扣率%Type;
  v_零售价       药品收发记录.零售价%Type;
  v_零售金额     药品收发记录.零售金额%Type;
  v_差价         药品收发记录.差价%Type;
  v_摘要         药品收发记录.摘要%Type;
  v_剩余数量     药品收发记录.实际数量%Type;
  v_剩余成本金额 药品收发记录.成本金额%Type;
  v_剩余零售金额 药品收发记录.零售金额%Type;
  v_入出系数     药品收发记录.入出系数%Type;

  v_收发id   药品收发记录.Id%Type;
  v_领用人   药品收发记录.领用人%Type;
  v_批准文号 药品收发记录.批准文号%Type;
  v_发药方式 药品收发记录.发药方式%Type;

  v_是否变价     收费项目目录.是否变价%Type;
  Intdigit       Number;
  n_上次供应商id 药品库存.上次供应商id%Type;
  d_上次生产日期 药品库存.上次生产日期%Type;
  v_按月留存领用 Varchar2(4000);
Begin
  --获取金额小数位数 
  Select Nvl(精度, 2) Into Intdigit From 药品卫材精度 Where 性质 = 0 And 类别 = 1 And 内容 = 4 And 单位 = 5;
  Select Nvl(是否变价, 0) Into v_是否变价 From 收费项目目录 Where Id = 药品id_In;
  Select Zl_Getsysparameter('按月留存领用', 1305) Into v_按月留存领用 From Dual;

  If 冲销方式_In = 1 Then
    --产生冲销申请单据，不填写审核人、审核日期，不更新库存记录 
    If 行次_In = 1 Then
      Update 药品收发记录
      Set 记录状态 = Decode(原记录状态_In, 1, 3, 原记录状态_In + 3)
      Where No = No_In And 单据 = 7 And 记录状态 = 原记录状态_In;
      If Sql%Rowcount = 0 Then
        Raise Err_Isstriked;
      End If;
    End If;
  
    --主要针对原不分批现在分批的药品，不能对其冲销
    Begin
      Select Distinct '(' || i.编码 || ')' || Nvl(n.名称, i.名称) As 药品信息
      Into v_Druginf
      From 药品收发记录 a, 药品规格 b, 收费项目目录 i, 收费项目别名 n
      Where a.药品id = b.药品id And a.药品id = i.Id And a.药品id = n.收费细目id(+) And n.性质(+) = 3 And a.No = No_In And a.单据 = 7 And
            Mod(a.记录状态, 3) = 0 And Nvl(a.批次, 0) = 0 And a.药品id + 0 = 药品id_In And
            ((Nvl(b.药库分批, 0) = 1 And
            a.库房id Not In (Select 部门id From 部门性质说明 Where (工作性质 Like '%药房') Or (工作性质 Like '制剂室'))) Or
            Nvl(b.药房分批, 0) = 1) And Rownum = 1;
    Exception
      When Others Then
        v_Druginf := '';
    End;
  
    If v_Druginf Is Not Null Then
      Raise Err_Isbatch;
    End If;
  
    Select Sum(实际数量) As 剩余数量, Sum(成本金额) As 剩余成本金额, Sum(零售金额) As 剩余零售金额, 库房id, 对方部门id, 入出类别id, 入出系数, Nvl(批次, 0) As 批次, 产地, 原产地, 批号, 效期, 成本价,
           扣率, 零售价, 摘要, 领用人, 批准文号, 发药方式, 供药单位id, 生产日期
    Into v_剩余数量, v_剩余成本金额, v_剩余零售金额, v_库房id, v_对方部门id, v_入出类别id, v_入出系数, v_批次, v_产地, v_原产地, v_批号, v_效期, v_成本价, v_扣率, v_零售价,
         v_摘要, v_领用人, v_批准文号, v_发药方式, n_上次供应商id, d_上次生产日期
    From 药品收发记录
    Where No = No_In And 单据 = 7 And 药品id = 药品id_In And 序号 = 序号_In
    Group By 库房id, 对方部门id, 入出类别id, 入出系数, Nvl(批次, 0), 产地, 原产地, 批号, 效期, 成本价, 扣率, 零售价, 摘要, 领用人, 批准文号, 发药方式, 供药单位id, 生产日期;
  
    --冲销数量大于剩余数量，不允许 
    If v_剩余数量 < 冲销数量_In Then
      Raise Err_Isnonum;
    End If;
  
    v_成本金额 := Round(冲销数量_In / v_剩余数量 * v_剩余成本金额, Intdigit);
    v_零售金额 := Round(冲销数量_In / v_剩余数量 * v_剩余零售金额, Intdigit);
    v_差价     := v_零售金额 - v_成本金额;
  
    Select 药品收发记录_Id.Nextval Into v_收发id From Dual;
  
    Insert Into 药品收发记录
      (Id, 记录状态, 单据, No, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 原产地, 批号, 效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额, 差价, 摘要,
       填制人, 填制日期, 审核人, 审核日期, 领用人, 批准文号, 发药方式, 供药单位id, 生产日期, 扣率, 冲销原因)
    Values
      (v_收发id, Decode(原记录状态_In, 1, 2, 原记录状态_In + 2), 7, No_In, 序号_In, v_库房id, v_对方部门id, v_入出类别id, v_入出系数, 药品id_In, v_批次,
       v_产地, v_原产地, v_批号, v_效期, -冲销数量_In, -冲销数量_In, v_成本价, -v_成本金额, v_零售价, -v_零售金额, -v_差价, v_摘要, 填制人_In, 填制日期_In, Null, Null,
       v_领用人, v_批准文号, v_发药方式, n_上次供应商id, d_上次生产日期, v_扣率, 冲销原因_In);
    
    Zl_未审药品记录_Insert(v_收发id);
    
  Elsif 冲销方式_In = 2 Then
    --审核已产生的冲销申请单据，填写审核人、审核日期，更新库存记录 
  
    --填写审核人、审核日期 
    Update 药品收发记录
    Set 审核人 = 填制人_In, 审核日期 = 填制日期_In
    Where 单据 = 7 And No = No_In And 序号 = 序号_In And 记录状态 = 原记录状态_In;
  
    --查询当前行记录的对应ID
    Select Id
    Into v_收发id
    From 药品收发记录
    Where 单据 = 7 And No = No_In And 序号 = 序号_In And 记录状态 = 原记录状态_In;
  
    --更新库存信息 领用冲销相当于入库 
    Zl_药品库存_Update(v_收发id, 3, 0);
    
    Zl_未审药品记录_Delete(v_收发id);

    --科室药品留存处理 
    If v_发药方式 = 1 Then
      Update 药品留存
      Set 可用数量 = Nvl(可用数量, 0) + 冲销数量_In, 实际数量 = Nvl(实际数量, 0) + 冲销数量_In, 实际金额 = Nvl(实际金额, 0) + v_零售金额
      Where 期间 = To_Char(Sysdate, Decode(v_按月留存领用, '1', 'yyyymm', 'yyyy')) And 科室id = v_对方部门id And 库房id = v_库房id And
            药品id = 药品id_In;
      --将金额和数量等于0的记录删除掉 
      Delete From 药品留存 Where Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0;
    End If;
  
    --处理调价后冲销 
    Zl_药品收发记录_调价修正(v_收发id);
  Else
    --正常冲销方式，产生冲销记录，填写审核人、审核日期，更新库存记录      
    If 行次_In = 1 Then
      Update 药品收发记录
      Set 记录状态 = Decode(原记录状态_In, 1, 3, 原记录状态_In + 3)
      Where No = No_In And 单据 = 7 And 记录状态 = 原记录状态_In;
      If Sql%Rowcount = 0 Then
        Raise Err_Isstriked;
      End If;
    End If;
  
    --主要针对原不分批现在分批的药品，不能对其冲销
    Begin
      Select Distinct '(' || i.编码 || ')' || Nvl(n.名称, i.名称) As 药品信息
      Into v_Druginf
      From 药品收发记录 a, 药品规格 b, 收费项目目录 i, 收费项目别名 n
      Where a.药品id = b.药品id And a.药品id = i.Id And a.药品id = n.收费细目id(+) And n.性质(+) = 3 And a.No = No_In And a.单据 = 7 And
            Mod(a.记录状态, 3) = 0 And Nvl(a.批次, 0) = 0 And a.药品id + 0 = 药品id_In And
            ((Nvl(b.药库分批, 0) = 1 And
            a.库房id Not In (Select 部门id From 部门性质说明 Where (工作性质 Like '%药房') Or (工作性质 Like '制剂室'))) Or
            Nvl(b.药房分批, 0) = 1) And Rownum = 1;
    Exception
      When Others Then
        v_Druginf := '';
    End;
  
    If v_Druginf Is Not Null Then
      Raise Err_Isbatch;
    End If;
  
    Select Sum(实际数量) As 剩余数量, Sum(成本金额) As 剩余成本金额, Sum(零售金额) As 剩余零售金额, 库房id, 对方部门id, 入出类别id, 入出系数, Nvl(批次, 0) As 批次, 产地, 原产地, 批号, 效期, 成本价,
           扣率, 零售价, 摘要, 领用人, 批准文号, 发药方式, 供药单位id, 生产日期
    Into v_剩余数量, v_剩余成本金额, v_剩余零售金额, v_库房id, v_对方部门id, v_入出类别id, v_入出系数, v_批次, v_产地, v_原产地, v_批号, v_效期, v_成本价, v_扣率, v_零售价,
         v_摘要, v_领用人, v_批准文号, v_发药方式, n_上次供应商id, d_上次生产日期
    From 药品收发记录
    Where No = No_In And 单据 = 7 And 药品id = 药品id_In And 序号 = 序号_In
    Group By 库房id, 对方部门id, 入出类别id, 入出系数, Nvl(批次, 0), 产地, 原产地, 批号, 效期, 成本价, 扣率, 零售价, 摘要, 领用人, 批准文号, 发药方式, 供药单位id, 生产日期;
  
    --冲销数量大于剩余数量，不允许 
    If v_剩余数量 < 冲销数量_In Then
      Raise Err_Isnonum;
    End If;
  
    v_成本金额 := Round(冲销数量_In / v_剩余数量 * v_剩余成本金额, Intdigit);
    v_零售金额 := Round(冲销数量_In / v_剩余数量 * v_剩余零售金额, Intdigit);
    v_差价     := v_零售金额 - v_成本金额;
  
    Select 药品收发记录_Id.Nextval Into v_收发id From Dual;
    Insert Into 药品收发记录
      (Id, 记录状态, 单据, No, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 原产地, 批号, 效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额, 差价, 摘要,
       填制人, 填制日期, 审核人, 审核日期, 领用人, 批准文号, 发药方式, 供药单位id, 生产日期, 扣率, 冲销原因)
    Values
      (v_收发id, Decode(原记录状态_In, 1, 2, 原记录状态_In + 2), 7, No_In, 序号_In, v_库房id, v_对方部门id, v_入出类别id, v_入出系数, 药品id_In, v_批次,
       v_产地, v_原产地, v_批号, v_效期, -冲销数量_In, -冲销数量_In, v_成本价, -v_成本金额, v_零售价, -v_零售金额, -v_差价, v_摘要, 填制人_In, 填制日期_In, 填制人_In,
       填制日期_In, v_领用人, v_批准文号, v_发药方式, n_上次供应商id, d_上次生产日期, v_扣率, 冲销原因_In);
    
    --更新库存信息 领用冲销相当于入库 
    Zl_药品库存_Update(v_收发id, 3, 0);
    
    --科室药品留存处理 
    If v_发药方式 = 1 Then
      Update 药品留存
      Set 可用数量 = Nvl(可用数量, 0) + 冲销数量_In, 实际数量 = Nvl(实际数量, 0) + 冲销数量_In, 实际金额 = Nvl(实际金额, 0) + v_零售金额
      Where 期间 = To_Char(Sysdate, Decode(v_按月留存领用, '1', 'yyyymm', 'yyyy')) And 科室id = v_对方部门id And 库房id = v_库房id And
            药品id = 药品id_In;
      --将金额和数量等于0的记录删除掉 
      Delete From 药品留存 Where Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0;
    End If;
  
    --处理调价后冲销 
    Zl_药品收发记录_调价修正(v_收发id);
  End If;
Exception
  When Err_Isstriked Then
    Raise_Application_Error(-20101, '[ZLSOFT]该单据已经被他人冲销！[ZLSOFT]');
  When Err_Isbatch Then
    Raise_Application_Error(-20102,
                            '[ZLSOFT]该单据中包含有一条原来不分批，现在分批的药品[' || v_Druginf || ']，不能冲销！[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_药品领用_Strike;
/

--139589:刘涛,2019-04-18,防止批次为空特殊处理
Create Or Replace Procedure Zl_药品移库_Strike
(
  行次_In       In Integer,
  原记录状态_In In 药品收发记录.记录状态%Type,
  No_In         In 药品收发记录.No%Type,
  序号_In       In 药品收发记录.序号%Type,
  药品id_In     In 药品收发记录.药品id%Type,
  冲销数量_In   In 药品收发记录.实际数量%Type,
  填制人_In     In 药品收发记录.填制人%Type,
  填制日期_In   In 药品收发记录.填制日期%Type,
  摘要_In       In 药品收发记录.摘要%Type := Null,
  冲销方式_In   In Integer := 0 --0－正常冲销方式；1－产生冲销申请单据；2－审核已产生的冲销申请单据
) Is
  Err_Isstriked Exception;
  Err_Isoutstock Exception;
  Err_Isnonum Exception;
  Err_Isbatch Exception;
  v_Druginf      Varchar2(50); --原不分批现在分批的药品信息
  v_库房id       药品收发记录.库房id%Type;
  v_批次         药品收发记录.批次%Type;
  v_成本价       药品收发记录.成本价%Type;
  v_成本金额     药品收发记录.成本金额%Type;
  v_零售价       药品收发记录.零售价%Type;
  v_零售金额     药品收发记录.零售金额%Type;
  v_差价         药品收发记录.差价%Type;
  v_剩余数量     药品收发记录.实际数量%Type;
  v_剩余成本金额 药品收发记录.成本金额%Type;
  v_剩余零售金额 药品收发记录.零售金额%Type;
  v_收发id       药品收发记录.Id%Type;
  v_批准文号     药品收发记录.批准文号%Type;

  v_药库分批 Integer;
  v_药房分批 Integer;
  Intdigit   Number;
  n_操作类型 Number;

  Cursor c_药品收发记录 Is
    Select a.Id, a.序号, a.库房id, a.对方部门id, a.入出类别id, a.入出系数, a.药品id, nvl(a.批次,0) 批次, a.产地, a.原产地, a.批号, a.效期, a.配药人, a.配药日期, a.摘要,
           a.供药单位id, a.批准文号, a.生产日期, a.成本价, a.零售价, Nvl(b.是否变价, 0) As 时价, a.扣率, a.单量, a.频次
    From 药品收发记录 A, 收费项目目录 B
    Where a.药品id = b.Id And a.No = No_In And a.单据 = 6 And (a.序号 >= 序号_In And a.序号 <= 序号_In + 1) And
          (a.记录状态 = 1 Or Mod(a.记录状态, 3) = 0)
    Order By a.药品id,a.批次;

  Cursor c_冲销申请记录 Is
    Select a.Id, a.序号, a.库房id, a.对方部门id, a.入出类别id, a.入出系数, a.药品id, nvl(a.批次,0) 批次, a.产地, a.原产地, a.批号, a.效期, a.配药人, a.配药日期, a.摘要,
           a.供药单位id, a.批准文号, a.生产日期, a.成本价, a.实际数量, a.零售金额, a.差价, a.零售价, Nvl(b.是否变价, 0) As 时价, a.扣率, a.单量, a.频次
    From 药品收发记录 A, 收费项目目录 B
    Where a.药品id = b.Id And a.No = No_In And a.单据 = 6 And (a.序号 >= 序号_In And a.序号 <= 序号_In + 1) And
          (a.记录状态 = 原记录状态_In And Mod(a.记录状态, 3) = 2) And a.审核日期 Is Null
    Order By a.药品id,a.批次;
Begin
  --获取金额小数位数
  Select Nvl(精度, 2) Into Intdigit From 药品卫材精度 Where 性质 = 0 And 类别 = 1 And 内容 = 4 And 单位 = 5;

  If 冲销方式_In = 0 Then
    n_操作类型 := 0;
  Else
    n_操作类型 := 1;
  End If;

  If 冲销方式_In = 1 Then
    --产生冲销申请单据，不填写审核人、审核日期，不更新库存记录
    If 行次_In = 1 Then
      Update 药品收发记录
      Set 记录状态 = Decode(原记录状态_In, 1, 3, 原记录状态_In + 3)
      Where NO = No_In And 单据 = 6 And 记录状态 = 原记录状态_In;
      If Sql%RowCount = 0 Then
        Raise Err_Isstriked;
      End If;
    End If;
  
    --主要针对原不分批现在分批的药品，不能对其冲销
    Begin
      Select Distinct '(' || i.编码 || ')' || Nvl(n.名称, i.名称) As 药品信息
      Into v_Druginf
      From 药品收发记录 A, 药品规格 B, 收费项目目录 I, 收费项目别名 N
      Where a.药品id = b.药品id And a.药品id = i.Id And a.药品id = n.收费细目id(+) And n.性质(+) = 3 And a.No = No_In And a.单据 = 6 And
            a.药品id + 0 = 药品id_In And Mod(a.记录状态, 3) = 0 And Nvl(a.批次, 0) = 0 And a.序号 = 序号_In And
            ((Nvl(b.药库分批, 0) = 1 And
            a.库房id Not In (Select 部门id From 部门性质说明 Where (工作性质 Like '%药房') Or (工作性质 Like '制剂室'))) Or
            Nvl(b.药房分批, 0) = 1) And Rownum = 1;
    Exception
      When Others Then
        v_Druginf := '';
    End;
  
    If v_Druginf Is Not Null Then
      Raise Err_Isbatch;
    End If;
  
    Select Sum(a.实际数量) As 剩余数量, Sum(a.成本金额) As 剩余成本金额, Sum(a.零售金额) As 剩余零售金额, a.成本价, a.零售价, a.对方部门id, Nvl(a.批次, 0),
           b.药库分批, b.药房分批, a.批准文号
    Into v_剩余数量, v_剩余成本金额, v_剩余零售金额, v_成本价, v_零售价, v_库房id, v_批次, v_药库分批, v_药房分批, v_批准文号
    From 药品收发记录 A, 药品规格 B
    Where a.No = No_In And a.药品id = b.药品id And a.单据 = 6 And a.药品id = 药品id_In And a.序号 = 序号_In
    Group By a.成本价, a.零售价, a.对方部门id, Nvl(a.批次, 0), b.药库分批, b.药房分批, a.批准文号;
  
    --V_分批:(原分批,现分批为批次;否则为零)原不分批,现分批,本过程不考虑
    --因为对于对方库房，相当于出库，而对于当前库房，相当于入库，所以当前库房不予检查，仅检查退库那条记录
    Select Nvl(a.批次, 0)
    Into v_批次
    From 药品收发记录 A
    Where a.No = No_In And a.单据 = 6 And a.药品id = 药品id_In And a.序号 = 序号_In + 1 And Mod(a.记录状态, 3) = 0;
  
    --冲销数量大于剩余数量，不允许
    If v_剩余数量 < 冲销数量_In Then
      Raise Err_Isnonum;
    End If;
  
    v_成本金额 := Round(冲销数量_In / v_剩余数量 * v_剩余成本金额, Intdigit);
    v_零售金额 := Round(冲销数量_In / v_剩余数量 * v_剩余零售金额, Intdigit);
    v_差价     := v_零售金额 - v_成本金额;
  
    For v_药品收发记录 In c_药品收发记录 Loop
    
      Select 药品收发记录_Id.Nextval Into v_收发id From Dual;
      Insert Into 药品收发记录
        (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 原产地, 批号, 效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额,
         差价, 摘要, 填制人, 填制日期, 审核人, 审核日期, 配药人, 配药日期, 供药单位id, 批准文号, 生产日期, 扣率, 单量, 频次)
      Values
        (v_收发id, Decode(原记录状态_In, 1, 2, 原记录状态_In + 2), 6, No_In, v_药品收发记录.序号, v_药品收发记录.库房id, v_药品收发记录.对方部门id,
         v_药品收发记录.入出类别id, v_药品收发记录.入出系数, 药品id_In, v_药品收发记录.批次, v_药品收发记录.产地, v_药品收发记录.原产地, v_药品收发记录.批号, v_药品收发记录.效期,
         -冲销数量_In, -冲销数量_In, v_成本价, -v_成本金额, v_零售价, -v_零售金额, -v_差价, 摘要_In, 填制人_In, 填制日期_In, Null, Null, v_药品收发记录.配药人,
         v_药品收发记录.配药日期, v_药品收发记录.供药单位id, v_药品收发记录.批准文号, v_药品收发记录.生产日期, v_药品收发记录.扣率, v_药品收发记录.单量, v_药品收发记录.频次);
    
      Zl_未审药品记录_Insert(v_收发id);
    
      --处理库存，原入的那笔相当于出库
      If v_药品收发记录.入出系数 = 1 Then
        Zl_药品库存_Update(v_收发id, 0, 1);
      End If;
    
      If v_药品收发记录.入出系数 = -1 Then
        v_库房id := v_药品收发记录.库房id;
      End If;
    End Loop;
  
  Elsif 冲销方式_In = 2 Then
    --审核已产生的冲销申请单据，填写审核人、审核日期，更新库存记录
    For v_药品收发记录 In c_冲销申请记录 Loop
      --填写审核人、审核日期
      Update 药品收发记录
      Set 审核人 = 填制人_In, 审核日期 = 填制日期_In
      Where NO = No_In And 单据 = 6 And ID = v_药品收发记录.Id;
    
      --更改药品库存表的相应数据，注意这时传入的数量等是负数
      --参数为1表示申请冲销时下可用数量，仅对原移入库房，下了可用数量就不用再更新可用数量了
      If v_药品收发记录.入出系数 = 1 Then
        Zl_药品库存_Update(v_药品收发记录.Id, 3, 1, n_操作类型);
      Else
        Zl_药品库存_Update(v_药品收发记录.Id, 3, 0);
      End If;
    
      Zl_未审药品记录_Delete(v_药品收发记录.Id);
    
      If v_药品收发记录.入出系数 = -1 Then
        v_库房id := v_药品收发记录.库房id;
      
      End If;
    
      --处理调价后冲销
      Zl_药品收发记录_调价修正(v_药品收发记录.Id);
    End Loop;
  
    b_Message.Zlhis_Drug_004(No_In);
  Else
    --正常冲销方式，产生冲销记录，填写审核人、审核日期，更新库存记录
    If 行次_In = 1 Then
      Update 药品收发记录
      Set 记录状态 = Decode(原记录状态_In, 1, 3, 原记录状态_In + 3)
      Where NO = No_In And 单据 = 6 And 记录状态 = 原记录状态_In;
      If Sql%RowCount = 0 Then
        Raise Err_Isstriked;
      End If;
    End If;
  
    --主要针对原不分批现在分批的药品，不能对其冲销
    Begin
      Select Distinct '(' || i.编码 || ')' || Nvl(n.名称, i.名称) As 药品信息
      Into v_Druginf
      From 药品收发记录 A, 药品规格 B, 收费项目目录 I, 收费项目别名 N
      Where a.药品id = b.药品id And a.药品id = i.Id And a.药品id = n.收费细目id(+) And n.性质(+) = 3 And a.No = No_In And a.单据 = 6 And
            a.药品id + 0 = 药品id_In And Mod(a.记录状态, 3) = 0 And Nvl(a.批次, 0) = 0 And a.序号 = 序号_In And
            ((Nvl(b.药库分批, 0) = 1 And
            a.库房id Not In (Select 部门id From 部门性质说明 Where (工作性质 Like '%药房') Or (工作性质 Like '制剂室'))) Or
            Nvl(b.药房分批, 0) = 1) And Rownum = 1;
    Exception
      When Others Then
        v_Druginf := '';
    End;
  
    If v_Druginf Is Not Null Then
      Raise Err_Isbatch;
    End If;
  
    Select Sum(a.实际数量) As 剩余数量, Sum(a.成本金额) As 剩余成本金额, Sum(a.零售金额) As 剩余零售金额, a.成本价, a.零售价, a.对方部门id, Nvl(a.批次, 0),
           b.药库分批, b.药房分批, a.批准文号
    Into v_剩余数量, v_剩余成本金额, v_剩余零售金额, v_成本价, v_零售价, v_库房id, v_批次, v_药库分批, v_药房分批, v_批准文号
    From 药品收发记录 A, 药品规格 B
    Where a.No = No_In And a.药品id = b.药品id And a.单据 = 6 And a.药品id = 药品id_In And a.序号 = 序号_In
    Group By a.成本价, a.零售价, a.对方部门id, Nvl(a.批次, 0), b.药库分批, b.药房分批, a.批准文号;
  
    --V_分批:(原分批,现分批为批次;否则为零)原不分批,现分批,本过程不考虑
    --因为对于对方库房，相当于出库，而对于当前库房，相当于入库，所以当前库房不予检查，仅检查退库那条记录
    Select Nvl(a.批次, 0)
    Into v_批次
    From 药品收发记录 A
    Where a.No = No_In And a.单据 = 6 And a.药品id = 药品id_In And a.序号 = 序号_In + 1 And Mod(a.记录状态, 3) = 0;
  
    --冲销数量大于剩余数量，不允许
    If v_剩余数量 < 冲销数量_In Then
      Raise Err_Isnonum;
    End If;
  
    v_成本金额 := Round(冲销数量_In / v_剩余数量 * v_剩余成本金额, Intdigit);
    v_零售金额 := Round(冲销数量_In / v_剩余数量 * v_剩余零售金额, Intdigit);
    v_差价     := v_零售金额 - v_成本金额;
  
    For v_药品收发记录 In c_药品收发记录 Loop
      Select 药品收发记录_Id.Nextval Into v_收发id From Dual;
      Insert Into 药品收发记录
        (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 原产地, 批号, 效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额,
         差价, 摘要, 填制人, 填制日期, 审核人, 审核日期, 配药人, 配药日期, 供药单位id, 批准文号, 生产日期, 扣率, 单量, 频次)
      Values
        (v_收发id, Decode(原记录状态_In, 1, 2, 原记录状态_In + 2), 6, No_In, v_药品收发记录.序号, v_药品收发记录.库房id, v_药品收发记录.对方部门id,
         v_药品收发记录.入出类别id, v_药品收发记录.入出系数, 药品id_In, v_药品收发记录.批次, v_药品收发记录.产地, v_药品收发记录.原产地, v_药品收发记录.批号, v_药品收发记录.效期,
         -冲销数量_In, -冲销数量_In, v_成本价, -v_成本金额, v_零售价, -v_零售金额, -v_差价, 摘要_In, 填制人_In, 填制日期_In, 填制人_In, 填制日期_In,
         v_药品收发记录.配药人, v_药品收发记录.配药日期, v_药品收发记录.供药单位id, v_药品收发记录.批准文号, v_药品收发记录.生产日期, v_药品收发记录.扣率, v_药品收发记录.单量,
         v_药品收发记录.频次);
    
      --更改药品库存表的相应数据
      If v_药品收发记录.入出系数 = 1 Then
        Zl_药品库存_Update(v_收发id, 3, 1, n_操作类型);
      Else
        Zl_药品库存_Update(v_收发id, 3, 0);
      End If;
    
      --处理调价后冲销
      Zl_药品收发记录_调价修正(v_收发id);
    End Loop;
  
    b_Message.Zlhis_Drug_004(No_In);
  End If;
Exception
  When Err_Isstriked Then
    Raise_Application_Error(-20101, '[ZLSOFT]该单据已经被他人冲销！[ZLSOFT]');
  When Err_Isbatch Then
    Raise_Application_Error(-20102, '[ZLSOFT]该单据中包含有一条原来不分批，现在分批的药品[' || v_Druginf || ']，不能冲销！[ZLSOFT]');
  When Err_Isnonum Then
    Raise_Application_Error(-20103, '[ZLSOFT]该单据中第' || Ceil(序号_In / 2) || '行的药品冲销的数量大于了剩余的数据，不能冲销！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品移库_Strike;
/

--139589:刘涛,2019-04-18,防止批次为空特殊处理
Create Or Replace Procedure Zl_自制入库_Strike
(
  No_In     In 药品收发记录.No%Type,
  审核人_In In 药品收发记录.审核人%Type
) Is
  Err_Isstriked Exception;

  v_入出类别id 药品收发记录.入出类别id%Type;

  Cursor c_药品收发记录 Is
    Select ID, 库房id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 填写数量, 实际数量, 成本价, 零售金额, 差价, 供药单位id, 生产日期, 批准文号, 单量
    From 药品收发记录 A
    Where NO = No_In And 单据 = 2 And 记录状态 = 2
    Order By 药品id,批次;
Begin
  Update 药品收发记录 Set 记录状态 = 3 Where NO = No_In And 单据 = 2 And 记录状态 = 1;

  If Sql%RowCount = 0 Then
    Raise Err_Isstriked;
  End If;
  
  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 入出类别id, 对方部门id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额, 差价, 摘要,
     填制人, 填制日期, 审核人, 审核日期, 费用id, 扣率, 供药单位id, 生产日期, 批准文号, 单量)
    Select 药品收发记录_Id.Nextval, 2, 2, No_In, 序号, 库房id, 入出类别id, 对方部门id, 入出系数, 药品id, nvl(批次,0), 产地, 批号, 效期, -填写数量, -实际数量, 成本价,
           -成本金额, 零售价, -零售金额, -差价, 摘要, 审核人_In, Sysdate, 审核人_In, Sysdate, 费用id, 扣率, 供药单位id, 生产日期, 批准文号, 单量
    From 药品收发记录
    Where NO = No_In And 单据 = 2 And 记录状态 = 3;
  

  For v_药品收发记录 In c_药品收发记录 Loop
    --更改药品库存表的相应数据
    If v_药品收发记录.入出系数 = 1 Then
      Zl_药品库存_Update(v_药品收发记录.Id, 3, 1);
    Else
      Zl_药品库存_Update(v_药品收发记录.Id, 3, 0);
    End If;
  
    --处理调价后冲销
    Zl_药品收发记录_调价修正(v_药品收发记录.Id);
  End Loop;
Exception
  When Err_Isstriked Then
    Raise_Application_Error(-20101, '[ZLSOFT]该单据已经被他人冲销！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_自制入库_Strike;
/

--139589:刘涛,2019-04-18,防止批次为空特殊处理
Create Or Replace Function Zl_Fun_Getbatchnum
(
  药品id_In   药品批号对照.药品id%Type,
  生产厂家_In 药品批号对照.生产厂家%Type,
  批号_In     药品批号对照.批号%Type,
  成本价_In   药品批号对照.成本价%Type,
  售价_In     药品批号对照.售价%Type,
  新批次_In   药品批号对照.批次%Type,
  供应商ID_In   药品批号对照.供应商ID%Type
) Return Number Is
  --功能：药品入库产生入库记录时根据传递过来的参数找对应的批次
  --返回值：查询到的批次，如果批次>0则说明找到了批次,如果批次=0则说明没有找到
  --参数：
  --     生产厂家_in：入库传递过来的生产商
  --     批号_in：入库时录入的批号
  --     成本价_in 入库时的成本价
  --     售价_in  入库时的售价
  --     
  n_批次     药品批号对照.批次%Type;
  n_药库包装 药品规格.药库包装%Type;
  n_是否变价 收费项目目录.是否变价%Type;
  n_Count    Number(1);
Begin
  --只处理生产厂家和批号不为空的情况
  If 生产厂家_In Is Not Null And 批号_In Is Not Null Then
    Begin
      Select nvl(批次,0)
      Into n_批次
      From 药品批号对照
      Where 药品id = 药品id_In And Nvl(生产厂家, 'a') = Nvl(生产厂家_In, 'a') And Nvl(批号, 'b') = Nvl(批号_In, 'b') And 成本价 = 成本价_In And
            售价 = 售价_In And Nvl(供应商id, 0) = Nvl(供应商id_In, 0);
    Exception
      When Others Then
        n_批次 := 新批次_In;
      
        If n_批次 > 0 Then
          --检查有无重复记录
          Begin
            Select 1
            Into n_Count
            From 药品批号对照
            Where 药品id = 药品id_In And Nvl(生产厂家, 'a') = Nvl(生产厂家_In, 'a') And Nvl(批号, 'b') = Nvl(批号_In, 'b') And
                  nvl(批次,0) = n_批次;
          Exception
            When Others Then
              n_Count := 0;
          End;
          
          --没有重复记录才能插入
          If n_Count = 0 Then
            Insert Into 药品批号对照
              (药品id, 生产厂家, 批号, 批次, 成本价, 售价,供应商ID)
            Values
              (药品id_In, 生产厂家_In, 批号_In, 新批次_In, 成本价_In, 售价_In,供应商ID_In);
          End If;
        End If;
    End;
  End If;

  Return(n_批次);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Fun_Getbatchnum;
/






------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0058' Where 编号=&n_System;
Commit;
