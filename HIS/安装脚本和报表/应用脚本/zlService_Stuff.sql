Create Or Replace Procedure Zl_StuffSvr_ExecutePrice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --执行调价
  ---------------------------------------------------------------------------
  --input      检查卫材售价，成本价是否存在已生效但未执行的价格，如果存在则执行调价
  --  stuff_id      N    材料id
  --output      
  --  code  C  1  应答码：0-失败；1-成功
  --  message  C  1  应答消息：
  ---------------------------------------------------------------------------
  j_Jsonin Pljson;
  j_Json   Pljson;

  n_材料id 材料特性.材料id%Type;

  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_材料id := j_Json.Get_Number('stuff_id');

  If n_材料id = 0 Then
    v_Err_Msg := '未传入卫材ID信息！';
    Raise Err_Item;
  End If;

  For r_调价 In (Select Distinct i.Id As 材料id
               From 收费项目目录 I, 收费价目 N, 材料特性 P
               Where i.Id = n.收费细目id And i.Id = p.材料id And
                     (i.撤档时间 Is Null Or i.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) And n.变动原因 = 0 And p.材料id = n_材料id And
                     Sysdate > n.执行日期
               Union
               Select Distinct a.药品id As 材料id
               From 药品价格记录 A
               Where a.记录状态 = 0 And a.药品id = n_材料id And a.执行日期 <= Sysdate
               Order By 材料id) Loop
  
    n_材料id := r_调价.材料id;
    Exit;
  End Loop;

  If n_材料id > 0 Then
    Zl_材料收发记录_Adjust(n_材料id);
  End If;

  Json_Out := '{"output":{"code": 1,"message":"成功"}}';
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_StuffSvr_ExecutePrice;
/

Create Or Replace Procedure Zl_StuffSvr_CheckHCostExistRec
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --判断高值卫材是否存在使用记录
  ---------------------------------------------------------------------------
  --input      判断高值卫材是否存在使用记录
  --  stuff_id      N    材料id
  --output      
  --  code  C  1  应答码：0-失败；1-成功
  --  message  C  1  应答消息：
  --  isexist  N 1 是否存在: 1-存在;0-不存在
  ---------------------------------------------------------------------------
  j_Jsonin Pljson;
  j_Json   Pljson;
  n_材料id 材料特性.材料id%Type;
  n_Exist  Number(1);
Begin
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_材料id := j_Json.Get_Number('stuff_id');

  Select Count(1)
  Into n_Exist
  From 药品收发记录 A, 收发记录补充信息 B
  Where a.药品id = n_材料id And a.Id = b.收发id And Rownum < 2;

  Json_Out := '{"output":{"code": 1,"message":"成功","isexist":' || n_Exist || '}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_StuffSvr_CheckHCostExistRec;
/

Create Or Replace Procedure Zl_StuffSvr_GetStockShow
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：获取指定库房的库存数据，用于显示
  --入参：Json_In:格式
  --  input
  --    warehouse_ids        C   1   库房ID串
  --出参: Json_Out,格式如下
  --  output
  --    code                N 1 应答吗：0-失败；1-成功
  --    message             C 1 应答消息：失败时返回具体的错误信息
  --    item_list
  --      stuff_id              N   1   材料ID
  --      warehouse_id          N   1   库房ID
  --      stock                N   1   可用数量
  --      real_stock          N  1 实际库存
  --      avg_price           N  1 平均售价
  --      avg_cost            N  1 平均成本价
  ---------------------------------------------------------------------------

  v_库房ids Varchar2(32767);

  j_Jsonin Pljson;
  j_Json   Pljson;

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;
Begin
  --解析入参
  j_Jsonin  := Pljson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  v_库房ids := j_Json.Get_String('warehouse_ids');

  If Nvl(v_库房ids, 0) = 0 Then
    Json_Out := Zljsonout('未传入相关库房信息');
    Return;
  End If;

  For c_库存 In (Select a.库房id, a.药品id, Sum(Nvl(a.可用数量, 0)) As 可用数量, Sum(Nvl(a.实际数量, 0)) As 实际数量,
                      Decode(Sum(Nvl(a.实际数量, 0)), 0, Max(a.零售价), Sum(Nvl(a.实际金额, 0)) / Sum(Nvl(a.实际数量, 0))) As 平均售价,
                      Decode(Sum(Nvl(a.实际数量, 0)), 0, Max(a.平均成本价),
                              (Sum(Nvl(a.实际金额, 0)) - Sum(Nvl(a.实际差价, 0))) / Sum(Nvl(a.实际数量, 0))) As 平均成本价
               From 药品库存 A
               Where a.性质 = 1 And Instr(',' || v_库房ids || ',', ',' || a.库房id || ',') > 0
               Group By a.库房id, a.药品id) Loop
  
    v_Jtmp := v_Jtmp || ',';
    Zljsonputvalue(v_Jtmp, 'stuff_id', c_库存.药品id, 1, 1);
    Zljsonputvalue(v_Jtmp, 'warehouse_id', c_库存.库房id, 1);
    Zljsonputvalue(v_Jtmp, 'stock', c_库存.可用数量, 1);
    Zljsonputvalue(v_Jtmp, 'real_stock', c_库存.实际数量, 1);
    Zljsonputvalue(v_Jtmp, 'avg_price', c_库存.平均售价, 1);
    Zljsonputvalue(v_Jtmp, 'avg_cost', c_库存.平均成本价, 1, 2);
  
    If Length(v_Jtmp) > 30000 Then
      If c_Jtmp Is Null Then
        c_Jtmp := Substr(v_Jtmp, 2);
      Else
        c_Jtmp := c_Jtmp || v_Jtmp;
      End If;
      v_Jtmp := Null;
    End If;
  End Loop;

  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || c_Jtmp || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_StuffSvr_GetStockShow;
/

Create Or Replace Procedure Zl_StuffSvr_GetCostPriceAdjust
(
  Json_In  Varchar2,
  Json_Out Out Clob
) Is
  ---------------------------------------------------------------------------
  --功能：获取卫材成本价调价记录
  --input      
  --  stuff_id      N   1 材料id
  --  show_unit    N   1   显示单位:0-散装单位;1-库房单位
  --output      
  --  code  C  1  应答码：0-失败；1-成功
  --  message  C  1  应答消息：  
  --  price_list[]  卫材成本价调价记录
  --     stuff_id   N 1  材料ID
  --     stuff_name   C 1  材料信息
  --     stock_name   C 1  库房
  --     batch_number   C 1  批号
  --     effective_time   C 1  效期
  --     place_name   C 1  产地
  --     unit_name   C 1  单位
  --     cost_old   N 1  原成本价
  --     cost_new    N 1  现成本价
  --     adjust_time   C 1  调价时间
  --     adjust_reson   C 1  调价说明
  --     adjust_no   C 1  调价单据号
  --     stuff_revoke_time  C 1 撤档时间
  --     node_no      C    0  站点编码   
  --     is_stock    N   1 是否有库存数据  0-否，1-是
  ---------------------------------------------------------------------------
  j_Jsonin Pljson;
  j_Json   Pljson;
  n_材料id 材料特性.材料ID%Type;
  n_单位   Number(1);

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;
Begin
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_材料id := j_Json.Get_Number('stuff_id');
  n_单位   := j_Json.Get_Number('show_unit');

  v_Jtmp := Null;
  For r_Costprice In (Select distinct b.No, i.Id As 材料id, '[' || i.编码 || ']' || i.名称 || ' ' || i.规格 || ' ' || i.产地 As 材料,
                             p.名称 As 库房, a.批号, a.效期, a.产地, Decode(n_单位, 0, i.计算单位, s.包装单位) As 单位,
                             Decode(n_单位, 0, a.原价, a.原价 * Nvl(s.换算系数, 1)) As 原成本价,
                             Decode(n_单位, 0, a.现价, a.现价 * Nvl(s.换算系数, 1)) As 成本价, a.执行日期, a.调价说明, i.撤档时间, i.站点,
                             Decode(k.库房id, Null, 0, 1) As 库存
                      From 药品收发记录 B, 收费项目目录 I, 材料特性 S, 部门表 P, 药品价格记录 A, 药品库存 K
                      Where a.价格类型 = 2 And a.收发id = b.Id(+) And a.药品id = i.Id And i.Id = s.材料id And a.库房id = p.Id(+) And
                            s.诊疗id = n_材料id And k.性质(+) = 1 And k.库房id(+) = a.库房id And k.药品id(+) = a.药品id And
                            k.批次(+) = a.批次
                      Order By '[' || i.编码 || ']' || i.名称 || ' ' || i.规格 || ' ' || i.产地, p.名称, a.批号, a.执行日期 Desc, b.No Desc) Loop
  
    v_Jtmp := v_Jtmp || ',';
    Zljsonputvalue(v_Jtmp, 'stuff_id', r_Costprice.材料id, 1, 1);
    Zljsonputvalue(v_Jtmp, 'stuff_name', r_Costprice.材料, 0);
    Zljsonputvalue(v_Jtmp, 'stock_name', r_Costprice.库房, 0);
    Zljsonputvalue(v_Jtmp, 'batch_number', r_Costprice.批号, 0);
    Zljsonputvalue(v_Jtmp, 'effective_time', r_Costprice.效期, 0);
  
    Zljsonputvalue(v_Jtmp, 'place_name', r_Costprice.产地, 0);
    Zljsonputvalue(v_Jtmp, 'unit_name', r_Costprice.单位, 0);
    Zljsonputvalue(v_Jtmp, 'cost_old', r_Costprice.原成本价, 1);
    Zljsonputvalue(v_Jtmp, 'cost_new', r_Costprice.成本价, 1);
    Zljsonputvalue(v_Jtmp, 'adjust_time', r_Costprice.执行日期, 0);
  
    Zljsonputvalue(v_Jtmp, 'adjust_reson', r_Costprice.调价说明, 0);
    Zljsonputvalue(v_Jtmp, 'adjust_no', r_Costprice.No, 0);
    Zljsonputvalue(v_Jtmp, 'stuff_revoke_time', r_Costprice.撤档时间, 0);
    Zljsonputvalue(v_Jtmp, 'node_no', r_Costprice.站点, 0);
    Zljsonputvalue(v_Jtmp, 'is_stock', r_Costprice.库存, 1, 2);
  
    If Length(v_Jtmp) > 30000 Then
      If c_Jtmp Is Null Then
        c_Jtmp := Substr(v_Jtmp, 2);
      Else
        c_Jtmp := c_Jtmp || v_Jtmp;
      End If;
      v_Jtmp := Null;
    End If;
  
  End Loop;

  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"成功","price_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","price_list":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_StuffSvr_GetCostPriceAdjust;
/

Create Or Replace Procedure Zl_StuffSvr_AdjustPriceType
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --执行调价
  ---------------------------------------------------------------------------
  --input      卫材价格属性调整时产生的调价盈亏和库存变化数据处理
  --    item_list[]         材料列表
  --       stuff_id      N    材料id
  --       price_type_old    N    原价格类型：0-定价；1-时价
  --       price_type_new    N    新价格类型：0-定价；1-时价
  --output      
  --  code  C  1  应答码：0-失败；1-成功
  --  message  C  1  应答消息：
  ---------------------------------------------------------------------------
  j_Jsonin   Pljson;
  j_Json     Pljson;
  j_Jsonlist Pljson_List;
  n_Count    Number;

  n_材料id     材料特性.材料id%Type;
  n_收入id     收费价目.收入项目id%Type;
  n_原价格类型 Number(1); --0-定价；1-时价
  n_新价格类型 Number(1); --0-定价；1-时价

  n_零售金额     药品库存.实际金额%Type;
  n_收发id       药品收发记录.Id%Type;
  n_流通金额小数 Number;
  n_序号         Number(8);
  n_入出类别id   Number(18); --入出类别
  v_Billno       药品收发记录.No%Type; --调价单号
  n_价格id       收费价目.Id%Type;
  n_收费价目现价 收费价目.现价%Type;
  n_收费价目原价 收费价目.原价%Type;
  n_原价         药品价格记录.原价%Type;
  n_价格记录     Number(1);
  v_类别         收费项目目录.类别%Type;

  --定价->时价后更新价格记录的值
  Cursor c_Priceadjust Is
    Select s.药品id, s.库房id, Nvl(s.批次, 0) As 批次, s.上次供应商id As 供应商id, s.上次批号 As 批号, s.效期, s.上次产地 As 产地,
           Nvl(s.实际数量, 0) As 实际数量, s.上次扣率 As 扣率, Nvl(s.实际金额, 0) As 实际金额, Nvl(s.实际差价, 0) As 实际差价, Nvl(s.零售价, 0) As 零售价,
           s.平均成本价, s.灭菌效期, s.批准文号, s.上次生产日期 As 生产日期
    From 药品库存 S
    Where s.药品id = n_材料id And s.性质 = 1
    Order By s.药品id, s.批次, s.库房id;

  r_Priceadjust c_Priceadjust%RowType;
Begin
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  j_Jsonlist := j_Json.Get_Pljson_List('item_list');

  n_Count := j_Jsonlist.Count;
  If j_Jsonlist.Count = 0 Then
    Json_Out := Zljsonout('未传入卫材信息！');
    Return;
  End If;

  For I In 1 .. j_Jsonlist.Count Loop
    j_Json := Pljson();
    j_Json := Pljson(j_Jsonlist.Get(I));
  
    n_材料id     := j_Json.Get_Number('stuff_id');
    n_原价格类型 := j_Json.Get_Number('price_type_old');
    n_新价格类型 := j_Json.Get_Number('price_type_new');
  
    If n_原价格类型 <> n_新价格类型 Then
      --取原价和原价id(调用该过程前已经产生了新价格)
      Begin
        Select 原价, 现价, 原价id As 价格id
        Into n_收费价目原价, n_收费价目现价, n_价格id
        From 收费价目
        Where 收费细目id = n_材料id And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And 变动原因 = 1;
      Exception
        When Others Then
          n_收费价目原价 := Null;
          n_收费价目现价 := Null;
          n_价格id       := Null;
      End;
      
      --时价->定价
      If n_原价格类型 = 1 And n_新价格类型 = 0 Then
        Select 精度 Into n_流通金额小数 From 药品卫材精度 Where 类别 = 2 And 内容 = 4 And 单位 = 5;
      
        --取入出类别ID
        Select 类别id Into n_入出类别id From 药品单据性质 Where 单据 = 38;        
              
        n_序号   := 0;
        v_Billno := Null;
      
        For r_Priceadjust In c_Priceadjust Loop
          If n_收费价目现价 <> r_Priceadjust.零售价 Then
            If v_Billno Is Null Then
              Select Nextno(147) Into v_Billno From Dual;
            End If;
            n_序号 := n_序号 + 1;
            Select 药品收发记录_Id.Nextval Into n_收发id From Dual;
            n_零售金额 := Round(n_收费价目现价 * r_Priceadjust.实际数量, n_流通金额小数) -
                      Round(r_Priceadjust.零售价 * r_Priceadjust.实际数量, n_流通金额小数);
            --产生调价影响记录
            Insert Into 药品收发记录
              (ID, 记录状态, 单据, NO, 序号, 入出类别id, 药品id, 批次, 批号, 效期, 产地, 付数, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 扣率, 零售金额, 差价, 摘要,
               填制人, 填制日期, 库房id, 入出系数, 价格id, 审核人, 审核日期, 单量, 频次, 供药单位id)
            Values
              (n_收发id, 1, 13, v_Billno, n_序号, n_入出类别id, r_Priceadjust.药品id, r_Priceadjust.批次, r_Priceadjust.批号,
               r_Priceadjust.效期, r_Priceadjust.产地, 1, r_Priceadjust.实际数量, 0, r_Priceadjust.零售价, 0, n_收费价目现价,
               r_Priceadjust.扣率, n_零售金额, n_零售金额, '卫材时价转定价', zl_UserName, Sysdate, r_Priceadjust.库房id, 1, n_价格id,
               zl_UserName, Sysdate, r_Priceadjust.实际金额, r_Priceadjust.实际差价, r_Priceadjust.供应商id);
          
            Zl_药品库存_Update(n_收发id, 2, 0);
          End If;
        End Loop;
      
        --定价->时价
      Elsif n_原价格类型 = 0 And n_新价格类型 = 1 Then
        For r_Priceadjust In c_Priceadjust Loop
          n_价格记录 := 0;
          Begin
            Select 1, 现价
            Into n_价格记录, n_原价
            From 药品价格记录
            Where 药品id = r_Priceadjust.药品id And 库房id = r_Priceadjust.库房id And Nvl(批次, 0) = r_Priceadjust.批次 And
                  记录状态 = 1 And 价格类型 = 1;
          Exception
            When Others Then
              n_价格记录 := 0;
              n_原价     := n_收费价目原价;
          End;
        
          If n_价格记录 = 1 Then
            Zl_药品价格记录_Stop(1, r_Priceadjust.库房id, r_Priceadjust.药品id, r_Priceadjust.批次, Sysdate - 1 / 24 / 60 / 60, 2);
          End If;
          Zl_药品价格记录_Insert(0, 1, r_Priceadjust.库房id, r_Priceadjust.药品id, r_Priceadjust.批次, n_原价, n_收费价目现价, Sysdate,
                           '卫材定价转时价', zl_UserName, Null, r_Priceadjust.供应商id, r_Priceadjust.批号, r_Priceadjust.效期,
                           r_Priceadjust.产地, r_Priceadjust.灭菌效期, Null, Null, Null, Null, 1);
        
          Update 药品库存
          Set 零售价 = n_收费价目现价
          Where 性质 = 1 And 库房id = r_Priceadjust.库房id And 药品id = r_Priceadjust.药品id And Nvl(批次, 0) = r_Priceadjust.批次;
        
        End Loop;
      End If;
    End If;
  End Loop;

  Json_Out := '{"output":{"code": 1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_StuffSvr_AdjustPriceType;
/


Create Or Replace Procedure Zl_StuffSvr_CheckExistStock
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  ---------------------------------------------------------------------------
  --input      判断卫材是否存在库数据
  --  stuff_id      N  1  卫材id
  --  is_item      N  1  是否按品种查询：0-按规格查询，1-按品种查询
  --output      
  --  code  C  1  应答码：0-失败；1-成功
  --  message  C  1  应答消息：
  --  isexist  N 1 是否存在: 1-存在;0-不存在
  ---------------------------------------------------------------------------
  j_Jsonin Pljson;
  j_Json   Pljson;
  n_材料id 材料特性.材料ID%Type;
  n_品种   Number(1);
  n_Exist  Number(1);
Begin
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_材料id := j_Json.Get_Number('stuff_id');
  n_品种   := j_Json.Get_Number('is_item');

  If n_品种 = 0 Then
    Select Count(1) Into n_Exist From 药品库存 Where 药品id = n_材料id And Rownum < 2;
  Else
    Select Count(1)
    Into n_Exist
    From 材料特性 A, 药品库存 B
    Where a.材料id = b.药品id And a.诊疗id = n_材料id And Rownum < 2;
  End If;

  Json_Out := '{"output":{"code": 1,"message":"成功","isexist":' || n_Exist || '}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_StuffSvr_CheckExistStock;
/

Create Or Replace Procedure Zl_StuffSvr_CheckStuffExistRec
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --判断卫材是否存在收发记录
  ---------------------------------------------------------------------------
  --input      判断卫材是否存在收发记录
  --  stuff_id      N    材料id
  --output      
  --  code  C  1  应答码：0-失败；1-成功
  --  message  C  1  应答消息：
  --  isexist  N 1 是否存在: 1-存在;0-不存在
  ---------------------------------------------------------------------------
  j_Jsonin Pljson;
  j_Json   Pljson;
  n_材料id 材料特性.材料ID%Type;
  n_Exist  Number(1);
Begin
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_材料id := j_Json.Get_Number('stuff_id');

  Select Count(1) Into n_Exist From 药品收发记录 Where 药品id = n_材料id And Rownum < 2;

  Json_Out := '{"output":{"code": 1,"message":"成功","isexist":' || n_Exist || '}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_StuffSvr_CheckStuffExistRec;
/

Create Or Replace Procedure Zl_Stuffsvr_Patiinfoupdate
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：住院病人药品收发记录基本信息修改
  --入参：Json_In:格式
  --  input
  --   pati_id        N   1   病人id
  --   pati_name      C   1   患者姓名
  --   pati_sex       C   1   患者性别
  --   pati_age       C   1   患者年龄
  --   visit_id       N   1   就诊id
  --出参: Json_Out,格式如下
  --  output
  --    code                 N   1 应答吗：0-失败；1-成功
  --    message              C   1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  n_病人id Number;
  n_就诊id Number;
  v_姓名   Varchar2(100);
  v_性别   Varchar2(100);
  v_年龄   Varchar2(100);
Begin
  --解析入参
  j_Jsonin := PLJson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');
  n_就诊id := j_Json.Get_Number('visit_id');
  v_姓名   := j_Json.Get_String('pati_name');
  v_性别   := j_Json.Get_String('pati_sex');
  v_年龄   := j_Json.Get_String('pati_age');

  Update 药品收发记录
  Set 姓名 = Nvl(v_姓名, 姓名), 性别 = Nvl(v_性别, 性别), 年龄 = Nvl(v_年龄, 年龄)
  Where 病人id = n_病人id And 主页id = n_就诊id;
  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Stuffsvr_Patiinfoupdate;
/

Create Or Replace Procedure Zl_Stuffsvr_Newbill_Check
(
  Json_In  Clob,
  Json_Out Out Varchar2
) Is
  --调用Zl_药品销售出库_Check完成功能
  --检查类过程，不需要事务控制
  --检查的内容：
  --1.高值备货卫材虚拟库房设置检查
  --2.药品卫材售价检查，收费界面和保存时可能发生变化（时价分批）
  --3.库存检查，根据参数及库存实际情况
  --4.分批属性检查，分批属性变化
  -------------------------------------------------------------------------------------------------
  --入参      json
  --input     根据条件对要产生的处方进行检查
  --  fee_list      收费明细信息，支持多个，[数组]
  --    stuff_id  N 1 卫材id
  --    send_num  N 1 发药数量
  --    warehouse_id  N 1 库房id
  --    price           N       1       售价
  --    is_bakstuff N   是否备货卫材:有高值卫材才需要传入，非0表示是高值卫材模式(如扫码时使用)
  --    bakstuff_batch  N   备货材料批次
  --出参      json
  --output      
  --  code  C 1 应答码：0-失败；1-成功
  --  message C 1 应答消息：
  -------------------------------------------------------------------------------------------------
Begin

  Zl_卫材销售出库_Check(Json_In, Json_Out);

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Newbill_Check;
/

Create Or Replace Procedure Zl_Stuffsvr_Send_Check
(
  Json_In  Clob,
  Json_Out Out Varchar2
) Is
  -------------------------------------------------------------------------------------------------
  --功能：在批量发料前检查发料数据是否存在，是否已经发料，是否已经拒发，是否不允许未审核/未收费的记录发料等
  --入参：json格式
  --Input
  --  stuff_rec_id C   药品收发记录ID串,支持多个id，用“,”分隔
  --出参：json格式
  --Json_Out
  --  code  C  1  应答码：0-失败；1-成功
  --  message  C  1  "应答消息：
  --  成功时返回成功信息
  --  失败时返回具体的错误信息"
  -------------------------------------------------------------------------------------------------

Begin

  Zl_卫材批量发料_Check(Json_In, Json_Out);

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Send_Check;
/

Create Or Replace Procedure Zl_Stuffsvr_Checkpatiexecute
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：根据病人信息或获取未发料数据
  --入参：Json_In:格式
  --  input
  --     check_type         N 1 检查方式:0-按如下参数值进行检查；1-仅按病人ID和发料库房进行检查
  --     pati_id            N 1 病人ID
  --     pati_pageid        N 1 主页ID
  --     baby_num           N 1 婴儿序号:-1表示不区分;0-母亲的;>0具体婴儿费用
  --     fee_source         N 1 费用来源:1-门诊;2-住院;4-体检
  --     stuff_nos             处方单据号，数组如：["A0001","A0002"]
  --     warehouse_id       N 0 库房ID
  --出参: Json_Out,格式如下
  --  output
  --    code                    N 1 应答吗：0-失败；1-成功
  --    message                 C 1 应答消息：失败时返回具体的错误信息
  --    isexist                 N 1 是否存在: 1-存在;0-不存在
  --    stuff_notsend_infor     C 1 未发料信息,isexist=1时返回
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  n_检查方式 Number(1);
  v_Stuff    Varchar2(32767);
  n_病人id   药品收发记录.病人id%Type;
  n_主页id   药品收发记录.主页id%Type;
  n_婴儿序号 药品收发记录.婴儿序号%Type;
  v_No       药品收发记录.No%Type;
  n_费用来源 药品收发记录.费用来源%Type;
  j_Jsonlist Pljson_List := Pljson_List();
  n_Count    Number(18);
  l_Nos      t_StrList := t_StrList();
  v_项目     收费项目目录.名称%Type;
  v_部门     部门表.名称%Type;
  n_库房id   药品收发记录.库房id%Type;

  Type t_未发料 Is Ref Cursor;
  c_未发料 t_未发料;

Begin
  --解析入参
  j_Jsonin   := PLJson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_检查方式 := j_Json.Get_Number('check_type');
  n_病人id   := j_Json.Get_Number('pati_id');
  n_主页id   := j_Json.Get_Number('pati_pageid');
  n_婴儿序号 := j_Json.Get_Number('baby_num');
  n_费用来源 := j_Json.Get_Number('fee_source');
  j_Jsonlist := j_Json.Get_Pljson_List('stuff_nos');
  n_库房id   := j_Json.Get_Number('warehouse_id');

  If Nvl(n_检查方式, 0) = 0 Then
    n_Count := 0;
    If j_Jsonlist Is Not Null Then
      n_Count := j_Jsonlist.Count;
    End If;
    If n_Count <> 0 Then
      For I In 1 .. n_Count Loop
        v_No := j_Jsonlist.Get_String(I);
        l_Nos.Extend();
        l_Nos(l_Nos.Count) := v_No;
      End Loop;
    
      Open c_未发料 For
        Select Distinct b.No, d.名称 项目, c.名称 As 部门
        From 药品收发记录 B, 部门表 C, 收费项目目录 D
        Where b.药品id = d.Id And b.库房id + 0 = c.Id(+) And b.单据 In (25, 26) And Mod(b.记录状态, 3) = 1 And b.审核人 Is Null And
              b.病人id = n_病人id And (Nvl(b.主页id, 0) = Nvl(n_主页id, 0) Or n_费用来源 <> 2) And
              (Nvl(b.婴儿序号, 0) = Nvl(n_婴儿序号, 0) Or Nvl(n_婴儿序号, 0) = -1) And Nvl(b.摘要, '大医') <> '拒发' And
              b.No In (Select Column_Value From Table(l_Nos)) And Nvl(b.病人来源, 1) = n_费用来源;
    
    Else
      Open c_未发料 For
        Select Distinct b.No, d.名称 项目, c.名称 As 部门
        From 药品收发记录 B, 部门表 C, 收费项目目录 D
        Where b.药品id = d.Id And b.库房id + 0 = c.Id(+) And b.单据 In (25, 26) And Mod(b.记录状态, 3) = 1 And b.审核人 Is Null And
              b.病人id = n_病人id And (Nvl(b.主页id, 0) = Nvl(n_主页id, 0) Or n_费用来源 <> 2) And
              (Nvl(b.婴儿序号, 0) = Nvl(n_婴儿序号, 0) Or n_婴儿序号 = -1) And Nvl(b.摘要, '大医') <> '拒发' And b.病人来源 = n_费用来源;
    
    End If;
  
  Elsif n_检查方式 = 1 Then
    Open c_未发料 For
      Select Distinct b.No, d.名称 项目, c.名称 As 部门
      From 药品收发记录 B, 部门表 C, 收费项目目录 D
      Where b.药品id = d.Id And b.库房id + 0 = c.Id(+) And b.单据 In (24, 25) And Mod(b.记录状态, 3) = 1 And b.审核人 Is Null And
            b.病人id = n_病人id And b.库房id = n_库房id;
  End If;

  Loop
    Fetch c_未发料
      Into v_No, v_项目, v_部门;
    Exit When c_未发料%NotFound;
  
    If v_Stuff Is Not Null Then
      If Instr(Chr(13) || Chr(10) || v_Stuff || Chr(13) || Chr(10),
               Chr(13) || Chr(10) || '单据[' || Nvl(v_No, '') || ']中的' || Nvl(v_项目, '') || '：在' || Nvl(v_部门, '[未定部门]') ||
                '未发放' || Chr(13) || Chr(10), 1) = 0 Then
        If Lengthb(v_Stuff || Chr(13) || Chr(10) || '单据[' || Nvl(v_No, '') || ']中的' || Nvl(v_项目, '') || '：在' ||
                   Nvl(v_部门, '[未定部门]') || '未发放') <= 1000 Then
          v_Stuff := v_Stuff || Chr(13) || Chr(10) || '单据[' || Nvl(v_No, '') || ']中的' || Nvl(v_项目, '') || '：在' ||
                     Nvl(v_部门, '[未定部门]') || '未发放';
        Else
          v_Stuff := v_Stuff || Chr(13) || Chr(10) || '... ...';
        End If;
      End If;
    Else
      v_Stuff := '单据[' || Nvl(v_No, '') || ']中的' || Nvl(v_项目, '') || '：在' || Nvl(v_部门, '[未定部门]') || '未发放';
    End If;
  End Loop;

  n_Count := 0;
  If v_Stuff Is Not Null Then
    v_Stuff := '存在未发放卫材料：' || Chr(13) || Chr(10) || Chr(13) || Chr(10) || v_Stuff;
    n_Count := 1;
  End If;

  Json_Out := '{"output":{"code":1,"message":"成功","isexist":' || n_Count || ',"stuff_notsend_infor":"' ||
              Zltools.Zljsonstr(v_Stuff) || '"}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Checkpatiexecute;
/

Create Or Replace Procedure Zl_Stuffsvr_Getrefusesendlist
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：获取拒发药清单
  --入参：Json_In:格式
  --  input
  --     pati_id  N 1 病人Id
  --     pati_pageids C 1 主页IDs:多次住院 ，用逗号分离
  --出参: Json_Out,格式如下
  --  output
  --    code                    N   1   应答吗：0-失败；1-成功
  --    message                 C   1   应答消息：失败时返回具体的错误信息
  --    item_list []
  --        stuff_no          C   1   费用单据号
  --        stuffdtl_id          C   1   费用ID
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  v_主页ids Varchar2(4000);
  n_病人id  药品收发记录.病人id%Type;

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;
Begin
  --解析入参
  j_Jsonin  := PLJson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  n_病人id  := j_Json.Get_Number('pati_id');
  v_主页ids := j_Json.Get_String('pati_pageids');

  If Nvl(n_病人id, 0) = 0 Then
    Json_Out := zlJsonOut('未传入相关病人id信息');
    Return;
  End If;

  v_Jtmp := Null;
  For r_Info In (Select NO As Stuff_No, 费用id As Stuffdtl_Id
                 From 药品收发记录
                 Where 病人id = n_病人id And
                       (Instr(',' || Nvl(v_主页ids, '-') || ',', ',' || Nvl(主页id, 0) || ',') > 0 Or v_主页ids Is Null) And
                       Mod(记录状态, 3) = 1 And Nvl(摘要, '大一') = '拒发' And Instr(',21,24,25,26,', ',' || 单据 || ',') > 0
                 Order By NO, 费用id) Loop
  
    v_Jtmp := v_Jtmp || ',{"stuff_no":"' || r_Info.Stuff_No || '"';
    v_Jtmp := v_Jtmp || ',"stuffdtl_id":' || r_Info.Stuffdtl_Id;
    v_Jtmp := v_Jtmp || '}';
  
    If Length(v_Jtmp) > 30000 Then
      If c_Jtmp Is Null Then
        c_Jtmp := Substr(v_Jtmp, 2);
      Else
        c_Jtmp := c_Jtmp || v_Jtmp;
      End If;
      v_Jtmp := Null;
    End If;
  End Loop;

  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Getrefusesendlist;
/

Create Or Replace Procedure Zl_Stuffsvr_Getexecutednum
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：获取卫生材料已发放数量
  --入参：Json_In:格式
  --  input
  --     billtype               N 1 单据类型:1-收费处方发料；2-记帐单处方发料；3-记帐表处方发料
  --     stuff_nos              C 1 单据号:可以传入多张单据
  --     notcontain_zero        N 1 是否不包含已发数量为0的：1-不包含，0-包含
  --     stuffdtl_ids           C 0 卫材明细ids，多个用英文的逗号分隔,未传入时按单据号查找,传入时按明细id进行查找
  --     order_ids              C 0 医嘱id串，本次处理的一批医嘱id逗号分割
  --出参: Json_Out,格式如下
  --  output
  --    code                    N   1   应答吗：0-失败；1-成功
  --    message                 C   1   应答消息：失败时返回具体的错误信息
  --    item_list[]
  --       stuff_no             C   1   NO
  --       stuffdtl_id          N   1   费用ID
  --       order_id             N   0   医嘱id
  --       stuff_id             N   1   卫材ID
  --       sended_num           N   1   已发数量
  --       barcode_goods        C       商品条码
  --       barcode_inside       C       内部条码
  ---------------------------------------------------------------------------
  j_Jsonin Pljson;
  j_Json   Pljson;

  n_不包含    Number(1);
  n_单据      药品收发记录.单据%Type;
  v_Nos       Varchar2(4000);
  c_Order_Ids Clob;

  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_明细ids Collection_Type;
  I           Number;
  v_明细ids   Clob;

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;
Begin
  --解析入参
  j_Jsonin    := Pljson(Json_In);
  j_Json      := j_Jsonin.Get_Pljson('input');
  n_单据      := j_Json.Get_Number('billtype');
  v_Nos       := j_Json.Get_String('stuff_nos');
  n_不包含    := j_Json.Get_Number('notcontain_zero');
  c_Order_Ids := j_Json.Get_Clob('order_ids');

  If j_Json.Exist('stuffdtl_ids') Then
    v_明细ids := j_Json.Get_Clob('stuffdtl_ids');
  End If;

  If n_单据 = 1 Then
    n_单据 := 24;
  Elsif n_单据 = 2 Then
    n_单据 := 25;
  Elsif n_单据 = 3 Then
    n_单据 := 26;
  Elsif c_Order_Ids Is Null Then
    If v_明细ids Is Null Then
      Json_Out := Zljsonout('传入节点【billtype】错误，请检查！');
      Return;
    End If;
  End If;

  v_Jtmp := Null;
  If v_明细ids Is Null Then
    For r_卫材 In (Select /*+cardinality(j,10)*/
                  a.No, a.费用id, a.医嘱id, a.药品id, Sum(Nvl(a.付数, 1) * Decode(a.审核人, Null, 0, a.实际数量)) As 已发数量,
                  Max(商品条码) As 商品条码, Max(内部条码) As 内部条码
                 From 药品收发记录 A, Table(f_Str2list(v_Nos)) J
                 Where a.No = j.Column_Value And
                       (a.单据 = 24 And n_单据 = 24 Or n_单据 <> 24 And Instr(',25,26,', ',' || a.单据 || ',') > 0 Or
                       n_单据 Is Null) And
                       (c_Order_Ids Is Null Or Instr(',' || c_Order_Ids || ',', ',' || a.医嘱id || ',') > 0)
                 Group By a.No, a.费用id, a.药品id, a.医嘱id) Loop
    
      If Not (Nvl(n_不包含, 0) = 1 And Nvl(r_卫材.已发数量, 0) = 0) Then
        v_Jtmp := v_Jtmp || ',';
        Zljsonputvalue(v_Jtmp, 'stuff_no', r_卫材.No, 0, 1);
        Zljsonputvalue(v_Jtmp, 'stuffdtl_id', r_卫材.费用id, 1);
        Zljsonputvalue(v_Jtmp, 'order_id', r_卫材.医嘱id, 1);
        Zljsonputvalue(v_Jtmp, 'stuff_id', r_卫材.药品id, 1);
        Zljsonputvalue(v_Jtmp, 'sended_num', Nvl(r_卫材.已发数量, 0), 1);
        Zljsonputvalue(v_Jtmp, 'barcode_goods', r_卫材.商品条码, 0);
        Zljsonputvalue(v_Jtmp, 'barcode_inside', r_卫材.内部条码, 0, 2);
      
        If Length(v_Jtmp) > 30000 Then
          If c_Jtmp Is Null Then
            c_Jtmp := Substr(v_Jtmp, 2);
          Else
            c_Jtmp := c_Jtmp || v_Jtmp;
          End If;
          v_Jtmp := Null;
        End If;
      End If;
    End Loop;
  Else
    I := 0;
    While v_明细ids Is Not Null Loop
      If Length(v_明细ids) <= 4000 Then
        Col_明细ids(I) := v_明细ids;
        v_明细ids := Null;
      Else
        Col_明细ids(I) := Substr(v_明细ids, 1, Instr(v_明细ids, ',', 3980) - 1);
        v_明细ids := Substr(v_明细ids, Instr(v_明细ids, ',', 3980) + 1);
      End If;
      I := I + 1;
    End Loop;
  
    I := 0;
    For I In 0 .. Col_明细ids.Count - 1 Loop
      For r_卫材 In (Select /*+cardinality(j,10)*/
                    a.No, a.费用id, a.药品id, Sum(Nvl(a.付数, 1) * Decode(a.审核人, Null, 0, a.实际数量)) As 已发数量, Max(商品条码) As 商品条码,
                    Max(内部条码) As 内部条码
                   From 药品收发记录 A, Table(f_Num2list(Col_明细ids(I))) J
                   Where a.费用id = j.Column_Value
                   Group By a.No, a.费用id, a.药品id) Loop
      
        If Not (Nvl(n_不包含, 0) = 1 And Nvl(r_卫材.已发数量, 0) = 0) Then
          v_Jtmp := v_Jtmp || ',';
          Zljsonputvalue(v_Jtmp, 'stuff_no', r_卫材.No, 0, 1);
          Zljsonputvalue(v_Jtmp, 'stuffdtl_id', r_卫材.费用id, 1);
          Zljsonputvalue(v_Jtmp, 'stuff_id', r_卫材.药品id, 1);
          Zljsonputvalue(v_Jtmp, 'sended_num', Nvl(r_卫材.已发数量, 0), 1);
          Zljsonputvalue(v_Jtmp, 'barcode_goods', r_卫材.商品条码, 0);
          Zljsonputvalue(v_Jtmp, 'barcode_inside', r_卫材.内部条码, 0, 2);
        
          If Length(v_Jtmp) > 30000 Then
            If c_Jtmp Is Null Then
              c_Jtmp := Substr(v_Jtmp, 2);
            Else
              c_Jtmp := c_Jtmp || v_Jtmp;
            End If;
            v_Jtmp := Null;
          End If;
        End If;
      End Loop;
    End Loop;
  End If;

  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Getexecutednum;
/

Create Or Replace Procedure Zl_Stuffsvr_Getstockcheck
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：获取卫材库存检查方式
  --入参：Json_In:格式
  --  input
  --出参: Json_Out,格式如下
  --  output
  --    code                N 1 应答吗：0-失败；1-成功
  --    message             C 1 应答消息：失败时返回具体的错误信息
  --    item_list             [数组]
  --        warehouse_id    N   1   库房ID
  --        check_type      N   1   检查方式：0-不检查，1-检查提示名，2-检查禁止
  ---------------------------------------------------------------------------

  v_Output Varchar2(32767);
Begin
  --解析入参

  For r_Data In (Select Distinct b.部门id, Nvl(c.检查方式, 0) As 检查方式
                 From 部门性质说明 B, 材料出库检查 C
                 Where b.部门id = c.库房id(+) And b.服务对象 In (1, 2, 3) And b.工作性质 = '发料部门') Loop
  
    zlJsonPutValue(v_Output, 'warehouse_id', r_Data.部门id, 1, 1);
    zlJsonPutValue(v_Output, 'check_type', r_Data.检查方式, 1, 2);
  
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || v_Output || ']}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Getstockcheck;
/

Create Or Replace Procedure Zl_Stuffsvr_Getstock
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：获取指定卫材在指定库房的可用库存数
  --入参：Json_In:格式
  --  input
  --    stuff_id        N   1   卫材ID
  --    warehouse_ids   C   1   库房ids,多个用逗号分离
  --    batch           N       批次：<=0-不区分批次，>0只查某批次
  --出参: Json_Out,格式如下
  --  output
  --    code            N   1 应答吗：0-失败；1-成功
  --    message         C   1 应答消息：失败时返回具体的错误信息
  --    stock           N   1  可用库存
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  n_批次     药品收发记录.批次%Type;
  n_卫材id   药品收发记录.药品id%Type;
  n_库存数量 药品库存.可用数量 %Type;
  v_库房ids  Varchar2(4000);
Begin
  --解析入参
  j_Jsonin  := PLJson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  n_卫材id  := j_Json.Get_Number('stuff_id');
  v_库房ids := j_Json.Get_String('warehouse_ids');
  n_批次    := j_Json.Get_Number('batch');

  If Nvl(n_卫材id, 0) = 0 Then
    Json_Out := zlJsonOut('未传入相关卫生材料信息');
    Return;
  End If;

  --获取库存(不分批或分批),药房不分批(批次=0,这里为药房)不管效期
  If Nvl(n_批次, 0) <= 0 Then
    Select Nvl(Sum(a.可用数量), 0)
    Into n_库存数量
    From 药品库存 A
    Where (Nvl(a.批次, 0) = 0 Or a.效期 Is Null Or a.效期 > Trunc(Sysdate)) And a.性质 = 1 And a.药品id = n_卫材id And
          Instr(',' || v_库房ids || ',', ',' || a.库房id || ',') > 0;
  Else
    Select Nvl(Sum(a.可用数量), 0)
    Into n_库存数量
    From 药品库存 A
    Where Nvl(a.批次, 0) = n_批次 And (a.效期 Is Null Or a.效期 > Trunc(Sysdate)) And a.性质 = 1 And a.药品id = n_卫材id And
          (Instr(',' || v_库房ids || ',', ',' || a.库房id || ',') > 0 Or
           a.库房id In (Select 虚拟库房id
                      From 虚拟库房对照
                      Where Instr(',' || v_库房ids || ',', ',' || 科室id || ',') > 0 And Rownum < 2));
  End If;

  Json_Out := '{"output":{"code":1,"message":"成功","stock":' || n_库存数量 || '}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Getstock;
/

Create Or Replace Procedure Zl_Stuffsvr_Getstockbatch
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：批量获取多个卫材库存及价格信息:在项目选择器中展示库存及价格信息
  --入参：Json_In:格式
  --  input
  --   stuff_ids            C   1   卫材ID，多个用英文的逗号分隔
  --   warehouse_ids        C   0   库房IDs，库房IDs为NULL时为所有库房
  --   return_price         N   1   是否返回售价：1-返回价格信息(售价);0-不返回
  --   type                 N   1   0-不返回库房ID;1-返回库房ID
  --出参: Json_Out,格式如下
  --  output
  --    code                N 1 应答吗：0-失败；1-成功
  --    message             C 1 应答消息：失败时返回具体的错误信息
  --    item_list
  --      stuff_id            N   1   卫材ID
  --      warehouse_id        N   1  库房ID
  --      stock               N   1   可用数量
  --      price               N   1   零售价(返回价格时才有此项)
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  c_卫材ids  Clob;
  v_库房ids  Varchar2(32767);
  n_返回价格 Number(2);
  n_Type     Number(2);
  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_卫材ids Collection_Type;
  I           Number;

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;
Begin
  --解析入参
  j_Jsonin   := PLJson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  c_卫材ids  := j_Json.Get_Clob('stuff_ids');
  v_库房ids  := j_Json.Get_String('warehouse_ids');
  n_返回价格 := Nvl(j_Json.Get_Number('return_price'), 0);
  n_Type     := Nvl(j_Json.Get_Number('type'), 0);

  I := 0;
  While c_卫材ids Is Not Null Loop
    If Length(c_卫材ids) <= 4000 Then
      Col_卫材ids(I) := c_卫材ids;
      c_卫材ids := Null;
    Else
      Col_卫材ids(I) := Substr(c_卫材ids, 1, Instr(c_卫材ids, ',', 3980) - 1);
      c_卫材ids := Substr(c_卫材ids, Instr(c_卫材ids, ',', 3980) + 1);
    End If;
    I := I + 1;
  End Loop;

  v_Jtmp := Null;
  If n_返回价格 = 0 Then
    If n_Type = 0 Then
      For I In 0 .. Col_卫材ids.Count - 1 Loop
        For c_库存 In (Select /*+cardinality(b,10)*/
                      a.药品id, Sum(Nvl(a.可用数量, 0)) As 库存
                     From 药品库存 A, Table(f_Num2List(Col_卫材ids(I))) B
                     Where a.药品id = b.Column_Value And a.性质 = 1 And
                           (Instr(',' || v_库房ids || ',', ',' || a.库房id || ',') > 0 Or v_库房ids Is Null) And
                           (Nvl(a.批次, 0) = 0 Or a.效期 Is Null Or a.效期 > Trunc(Sysdate))
                     Group By a.药品id
                     Having Sum(Nvl(a.可用数量, 0)) <> 0) Loop
        
          v_Jtmp := v_Jtmp || ',{"stuff_id":' || c_库存.药品id;
          v_Jtmp := v_Jtmp || ',"stock":' || zlJsonStr(c_库存.库存, 1);
          v_Jtmp := v_Jtmp || '}';
        
          If Length(v_Jtmp) > 30000 Then
            If c_Jtmp Is Null Then
              c_Jtmp := Substr(v_Jtmp, 2);
            Else
              c_Jtmp := c_Jtmp || v_Jtmp;
            End If;
            v_Jtmp := Null;
          End If;
        End Loop;
      End Loop;
    
    Elsif n_Type = 1 Then
      For I In 0 .. Col_卫材ids.Count - 1 Loop
        For c_库存 In (Select a.药品id, a.库房id, a.库存
                     From (Select /*+cardinality(b,10)*/
                             a.药品id, a.库房id, Sum(Nvl(a.可用数量, 0)) As 库存
                            From 药品库存 A, Table(f_Num2List(Col_卫材ids(I))) B
                            Where a.药品id = b.Column_Value And a.性质 = 1 And
                                  (Instr(',' || v_库房ids || ',', ',' || a.库房id || ',') > 0 Or v_库房ids Is Null) And
                                  (Nvl(a.批次, 0) = 0 Or a.效期 Is Null Or a.效期 > Trunc(Sysdate))
                            Group By a.药品id, a.库房id
                            Having Sum(Nvl(a.可用数量, 0)) <> 0) A) Loop
        
          v_Jtmp := v_Jtmp || ',{"stuff_id":' || c_库存.药品id;
          v_Jtmp := v_Jtmp || ',"warehouse_id":' || c_库存.库房id;
          v_Jtmp := v_Jtmp || ',"stock":' || zlJsonStr(c_库存.库存, 1);
          v_Jtmp := v_Jtmp || '}';
        
          If Length(v_Jtmp) > 30000 Then
            If c_Jtmp Is Null Then
              c_Jtmp := Substr(v_Jtmp, 2);
            Else
              c_Jtmp := c_Jtmp || v_Jtmp;
            End If;
            v_Jtmp := Null;
          End If;
        End Loop;
      End Loop;
    End If;
  
    If c_Jtmp Is Null Then
      Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
    Else
      c_Jtmp   := c_Jtmp || v_Jtmp;
      Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || c_Jtmp || ']}}';
    End If;
    Return;
  End If;

  --包含价格
  v_Jtmp := Null;
  For I In 0 .. Col_卫材ids.Count - 1 Loop
    For c_库存 In (Select a.药品id, Nvl(a.库存, 0) As 库存, Decode(Nvl(b.是否变价, 0), 1, 0, Nvl(c.现价, 0)) As 价格
                 From (Select /*+cardinality(b,10)*/
                         a.药品id, Sum(Nvl(a.可用数量, 0)) As 库存
                        From 药品库存 A, Table(f_Num2List(Col_卫材ids(I))) B
                        Where a.药品id = b.Column_Value And a.性质 = 1 And
                              (Instr(',' || v_库房ids || ',', ',' || a.库房id || ',') > 0 Or v_库房ids Is Null) And
                              (Nvl(a.批次, 0) = 0 Or a.效期 Is Null Or a.效期 > Trunc(Sysdate))
                        Group By a.药品id
                        Having Sum(Nvl(a.可用数量, 0)) <> 0) A, 收费项目目录 B, 收费价目 C
                 Where a.药品id = c.收费细目id And a.药品id = b.Id And c.价格等级 Is Null And Sysdate Between c.执行日期 And
                       Nvl(c.终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))) Loop
    
      v_Jtmp := v_Jtmp || ',{"stuff_id":' || c_库存.药品id;
      v_Jtmp := v_Jtmp || ',"stock":' || zlJsonStr(c_库存.库存, 1);
      v_Jtmp := v_Jtmp || ',"price":' || zlJsonStr(c_库存.价格, 1);
      v_Jtmp := v_Jtmp || '}';
    
      If Length(v_Jtmp) > 30000 Then
        If c_Jtmp Is Null Then
          c_Jtmp := Substr(v_Jtmp, 2);
        Else
          c_Jtmp := c_Jtmp || v_Jtmp;
        End If;
        v_Jtmp := Null;
      End If;
    End Loop;
  End Loop;

  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Getstockbatch;
/

Create Or Replace Procedure Zl_Stuffsvr_Batchgetprice
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：批量获取卫材售价(自助机使用)
  --入参：Json_In:格式
  --  input
  --   stuff_ids    C   1   卫材IDs，多个用英文的逗号分隔
  --出参: Json_Out,格式如下
  --  output
  --    code          N 1 应答吗：0-失败；1-成功
  --    message         C 1 应答消息：失败时返回具体的错误信息
  --    item_list
  --      stuff_id N   1   卫材ID
  --      price   N   1   零售价(返回价格时才有此项)
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_卫材ids Collection_Type;
  I           Integer;
  c_卫材ids   Clob;

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;
Begin
  --解析入参
  j_Jsonin  := PLJson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  c_卫材ids := j_Json.Get_Clob('stuff_ids');

  If c_卫材ids Is Null Then
    Json_Out := zlJsonOut('未传入有效的卫材id,请检查!');
  End If;

  I := 0;
  While c_卫材ids Is Not Null Loop
    If Length(c_卫材ids) <= 4000 Then
      Col_卫材ids(I) := c_卫材ids;
      c_卫材ids := Null;
    Else
      Col_卫材ids(I) := Substr(c_卫材ids, 1, Instr(c_卫材ids, ',', 3980) - 1);
      c_卫材ids := Substr(c_卫材ids, Instr(c_卫材ids, ',', 3980) + 1);
    End If;
    I := I + 1;
  End Loop;

  v_Jtmp := Null;
  For I In 0 .. Col_卫材ids.Count - 1 Loop
    --包含价格
    For c_库存 In (With c_药品信息 As
                    (Select /*+cardinality(D,10)*/
                     d.Column_Value As 药品id
                    From Table(f_Num2List(Col_卫材ids(I))) D)
                   Select a.药品id, Sum(a.实际金额) / Sum(a.实际数量) As 价格
                   From 药品库存 A, c_药品信息 B
                   Where a.药品id = b.药品id And a.性质 = 1 And (Nvl(a.批次, 0) = 0 Or a.效期 Is Null Or a.效期 > Trunc(Sysdate))
                   Group By a.药品id
                   Having Sum (Nvl(a.实际数量, 0)) <> 0) Loop
    
      v_Jtmp := v_Jtmp || ',{"stuff_id":' || c_库存.药品id;
      v_Jtmp := v_Jtmp || ',"price":' || zlJsonStr(c_库存.价格, 1);
      v_Jtmp := v_Jtmp || '}';
    
      If Length(v_Jtmp) > 30000 Then
        If c_Jtmp Is Null Then
          c_Jtmp := Substr(v_Jtmp, 2);
        Else
          c_Jtmp := c_Jtmp || v_Jtmp;
        End If;
        v_Jtmp := Null;
      End If;
    End Loop;
  End Loop;

  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Batchgetprice;
/

Create Or Replace Procedure Zl_Stuffsvr_Checkexpirydate
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：检查卫生卫材的灭菌效期是否过期
  --入参：Json_In:格式
  --    input
  --        stuff_id        N   1   卫材ID
  --        warehouse_id    N   1   库房ID
  --        quantity        N   1   数量
  --        batch           N   1   批次 
  --出参: Json_Out,格式如下
  --  output
  --    code                N   1   应答码：0-失败；1-成功
  --    message             C   1   应答消息：失败时返回具体的错误信息
  --    exist_expiried      N   1   是否存在过期:0-不存在已过期项目，1-存在已过期项目
  --    min_expirydate      C       最小灭菌效期，exist_expiried=1时返回，格式：yyyy-mm-dd
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  n_库房id  药品收发记录.库房id%Type;
  n_卫材id  药品收发记录.药品id%Type;
  n_数量    药品库存.可用数量 %Type;
  n_批次    药品库存.批次 %Type;
  d_Mindate Date;
  d_Sysdate Date;
  n_Find    Number(2);
  v_Tmp     Varchar2(20);
Begin
  --解析入参
  j_Jsonin := PLJson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_卫材id := j_Json.Get_Number('stuff_id');
  n_库房id := j_Json.Get_Number('warehouse_id');
  n_数量   := j_Json.Get_Number('quantity');
  n_批次   := j_Json.Get_Number('batch');

  If Nvl(n_卫材id, 0) = 0 Then
    Json_Out := zlJsonOut('未传入相关卫生材料信息');
    Return;
  End If;

  d_Sysdate := Sysdate;
  d_Mindate := To_Date('3000-01-01', 'yyyy-mm-dd');
  n_Find    := 0;
  --仅一次性材料才判断
  -- 因为可能各批次灭菌效期不同, 检查要用到的批次中最小的效期
  For c_卫材 In (Select c.名称, Nvl(b.批次, 0) As 批次, b.可用数量 As 库存, b.灭菌效期
               From 材料特性 A, 药品库存 B, 收费项目目录 C
               Where a.材料id = b.药品id And a.材料id = c.Id And a.一次性材料 = 1 And b.性质 = 1 And Nvl(b.可用数量, 0) > 0 And
                     a.灭菌效期 Is Not Null And a.材料id = n_卫材id And b.库房id = n_库房id And
                     Decode(n_批次, Null, -1, b.批次) = Nvl(n_批次, -1)
               Order By Nvl(b.批次, 0)) Loop
    If Nvl(n_Find, 0) = 0 Then
      n_Find := 1;
    End If;
    If c_卫材.灭菌效期 < d_Mindate Then
      d_Mindate := c_卫材.灭菌效期;
    End If;
  
    If Nvl(c_卫材.库存, 0) < n_数量 Then
      n_数量 := n_数量 - Nvl(c_卫材.库存, 0);
    Else
      n_数量 := 0;
    End If;
    If n_数量 = 0 Then
      Exit;
    End If;
  
  End Loop;

  If d_Sysdate > d_Mindate And Nvl(n_Find, 0) = 1 Then
    v_Tmp  := To_Char(d_Mindate, 'yyyy -mm-dd');
    n_Find := 1;
  Else
    n_Find := 0;
  End If;

  Json_Out := '{"output":{"code":1,"message":"成功","exist_expiried":' || n_Find || ',"min_expirydate":"' || v_Tmp ||
              '"}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Checkexpirydate;
/

CREATE OR REPLACE Procedure Zl_Stuffsvr_Getprice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：获取指定卫材的售价、成本价
  --入参：Json_In:格式
  --  input
  --    stuff_id          N 1 卫材ID
  --    warehouse_id      N 1 药房ID
  --    quantity          N 1 数量
  --    batch             N   批次：0-不区分批次，>0只查某批次
  --    price_grade       C   价格等级名称
  --    item_list[]列表
  --            stuff_id          N 1 卫材ID
  --             warehouse_id      N 1 药房ID
  --             quantity          N 1 数量
  --             batch             N   批次：0-不区分批次，>0只查某批次
  --             price_grade       C   价格等级名称
  --出参: Json_Out,格式如下
  --  output
  --    code              N 1   应答吗：0-失败；1-成功
  --    message           C 1   应答消息：失败时返回具体的错误信息
  --    price             N 1 售价
  --    price_cost        N 1 成本价
  --    quantity_remain   N 1 剩余数量，等于0表示数量足够，大于0则表示数量不够
  --    item_list[]列表
  --            price             N 1 售价
  --            price_cost        N 1 成本价
  --            quantity_remain   N 1 剩余数量，等于0表示数量足够，大于0则表示数量不够
  ---------------------------------------------------------------------------
  j_Jsonin Pljson;
  j_Json   Pljson;
  j_List   Pljson_List := Pljson_List();
  j_Tmpout Varchar2(32767);

  n_卫材id   药品库存.药品id%Type;
  n_库房id   药品库存.库房id%Type;
  n_数量     药品库存.实际数量%Type;
  n_批次     药品库存.批次%Type;
  v_Temp     Varchar2(4000);
  n_单价     药品规格.成本价%Type;
  n_成本价   药品规格.成本价%Type;
  n_剩余数量 药品库存.实际数量%Type;
Begin
  --解析入参
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  
  If j_Json.Exist('item_list') Then
    j_List := j_Json.Get_Pljson_List('item_list');
    If j_List Is Not Null Then
      For I In 1 .. j_List.Count Loop
        j_Json := Pljson();
        j_Json := Pljson(j_List.Get(I));
      
        n_卫材id := j_Json.Get_Number('stuff_id');
        n_库房id := j_Json.Get_Number('warehouse_id');
        n_数量   := j_Json.Get_Number('quantity');
        n_批次   := j_Json.Get_Number('batch');
      
        v_Temp := Zl_Fun_Getprice(n_卫材id, n_库房id, n_数量, 0, n_批次);
        --分解
        n_单价     := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1));
        v_Temp     := Substr(v_Temp, Instr(v_Temp, '|') + 1);
        n_成本价   := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1));
        v_Temp     := Substr(v_Temp, Instr(v_Temp, '|') + 1);
        n_剩余数量 := To_Number(v_Temp);
      
        j_Tmpout := j_Tmpout || ',{"price":' || Zljsonstr(n_单价, 1);
        j_Tmpout := j_Tmpout || ',"price_cost": ' || Zljsonstr(n_成本价, 1);
        j_Tmpout := j_Tmpout || ',"quantity_remain":' || Zljsonstr(n_剩余数量, 1);
        j_Tmpout := j_Tmpout || '}';
      
      End Loop;
    End If;
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || Substr(j_Tmpout, 2) || ']}}';
  Else
  
    n_卫材id := j_Json.Get_Number('stuff_id');
    n_库房id := j_Json.Get_Number('warehouse_id');
    n_数量   := j_Json.Get_Number('quantity');
    n_批次   := j_Json.Get_Number('batch');
  
    v_Temp := Zl_Fun_Getprice(n_卫材id, n_库房id, n_数量, 0, n_批次);
    --分解
    n_单价     := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1));
    v_Temp     := Substr(v_Temp, Instr(v_Temp, '|') + 1);
    n_成本价   := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1));
    v_Temp     := Substr(v_Temp, Instr(v_Temp, '|') + 1);
    n_剩余数量 := To_Number(v_Temp);
  
    Json_Out := '{"output":{"code":1,"message":"成功"';
    Json_Out := Json_Out || ',"price":' || Zljsonstr(n_单价, 1);
    Json_Out := Json_Out || ',"price_cost": ' || Zljsonstr(n_成本价, 1);
    Json_Out := Json_Out || ',"quantity_remain":' || Zljsonstr(n_剩余数量, 1) || '}}';
  End If;
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Getprice;
/

Create Or Replace Procedure Zl_Stuffsvr_Getidbybarcode
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：根据商品条码或内部条码获取卫材ID及批次
  --入参：Json_In:格式
  --  input
  --    barcode             C 1 输入的条码串
  --    type                N 0 1-不检查库存及期效,0-检查库存及期效
  --    only_barcode_inside N 0 1-仅对内部条码进行查找,0-对商品条码及内部条码进行查找
  --出参: Json_Out,格式如下
  --  output
  --    code              N 1   应答吗：0-失败；1-成功
  --    message           C 1   应答消息：失败时返回具体的错误信息
  --    item_list
  --      stuff_id        N 1   材料ID
  --      batch           N     批次：0-不区分批次，>0只查某批次
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  v_条码   药品库存.商品条码%Type;
  n_Type   Number(1);
  n_Inside Number(1);

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;
Begin
  --解析入参
  j_Jsonin := PLJson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  v_条码   := j_Json.Get_String('barcode');
  n_Type   := Nvl(j_Json.Get_Number('type'), 0);
  n_Inside := Nvl(j_Json.Get_Number('only_barcode_inside'), 0);

  v_Jtmp := Null;
  If n_Type = 0 Then
    If n_Inside = 1 Then
      For c_条码 In (Select Distinct 药品id, 批次
                   From 药品库存
                   Where (内部条码 = v_条码) And 可用数量 > 0 And (效期 Is Null Or 效期 > Trunc(Sysdate))) Loop
      
        v_Jtmp := v_Jtmp || ',{"stuff_id":' || c_条码.药品id;
        v_Jtmp := v_Jtmp || ',"batch":' || Nvl(c_条码.批次, 0);
        v_Jtmp := v_Jtmp || '}';
      
        If Length(v_Jtmp) > 30000 Then
          If c_Jtmp Is Null Then
            c_Jtmp := Substr(v_Jtmp, 2);
          Else
            c_Jtmp := c_Jtmp || v_Jtmp;
          End If;
          v_Jtmp := Null;
        End If;
      End Loop;
    Else
      For c_条码 In (Select Distinct 药品id, 批次
                   From 药品库存
                   Where (商品条码 = v_条码 Or 内部条码 = v_条码) And 可用数量 > 0 And (效期 Is Null Or 效期 > Trunc(Sysdate))) Loop
      
        v_Jtmp := v_Jtmp || ',{"stuff_id":' || c_条码.药品id;
        v_Jtmp := v_Jtmp || ',"batch":' || Nvl(c_条码.批次, 0);
        v_Jtmp := v_Jtmp || '}';
      
        If Length(v_Jtmp) > 30000 Then
          If c_Jtmp Is Null Then
            c_Jtmp := Substr(v_Jtmp, 2);
          Else
            c_Jtmp := c_Jtmp || v_Jtmp;
          End If;
          v_Jtmp := Null;
        End If;
      End Loop;
    End If;
  Else
    If n_Inside = 1 Then
      For c_条码 In (Select Distinct 药品id, 批次
                   From 药品库存
                   Where 商品条码 Like v_条码 || '%' Or 内部条码 = v_条码 || '%') Loop
      
        v_Jtmp := v_Jtmp || ',{"stuff_id":' || c_条码.药品id;
        v_Jtmp := v_Jtmp || ',"batch":' || Nvl(c_条码.批次, 0);
        v_Jtmp := v_Jtmp || '}';
      
        If Length(v_Jtmp) > 30000 Then
          If c_Jtmp Is Null Then
            c_Jtmp := Substr(v_Jtmp, 2);
          Else
            c_Jtmp := c_Jtmp || v_Jtmp;
          End If;
          v_Jtmp := Null;
        End If;
      End Loop;
    Else
      For c_条码 In (Select Distinct 药品id, 批次 From 药品库存 Where 内部条码 = v_条码 || '%') Loop
      
        v_Jtmp := v_Jtmp || ',{"stuff_id":' || c_条码.药品id;
        v_Jtmp := v_Jtmp || ',"batch":' || Nvl(c_条码.批次, 0);
        v_Jtmp := v_Jtmp || '}';
      
        If Length(v_Jtmp) > 30000 Then
          If c_Jtmp Is Null Then
            c_Jtmp := Substr(v_Jtmp, 2);
          Else
            c_Jtmp := c_Jtmp || v_Jtmp;
          End If;
          v_Jtmp := Null;
        End If;
      End Loop;
    End If;
  End If;

  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Getidbybarcode;
/

Create Or Replace Procedure Zl_Stuffsvr_Checkstorelimit
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：检查指定卫材在指定库房的库存是否低于储备下限
  --入参：Json_In:格式
  --  input
  --    stuff_id            N   1   卫材ID
  --    warehouse_id        N   1   库房ID
  --    stock               N   1   库存数量
  --出参: Json_Out,格式如下
  --  output
  --    code                N 1 应答吗：0-失败；1-成功
  --    message             C 1 应答消息：失败时返回具体的错误信息
  --    below_limit_lower   N 1 1-低于储备下限，0-大于等于储备下限

  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  n_卫材id 药品库存.药品id%Type;
  n_库房id 药品库存.库房id%Type;
  n_库存   药品库存.实际数量%Type;
  n_Count  Number(1);
Begin
  --解析入参
  j_Jsonin := PLJson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_卫材id := j_Json.Get_Number('stuff_id');
  n_库房id := j_Json.Get_Number('warehouse_id');
  n_库存   := j_Json.Get_Number('stock');

  --读取药品储备限额
  Select Count(1)
  Into n_Count
  From 材料储备限额
  Where 材料id = n_卫材id And 库房id = n_库房id And Nvl(下限, 0) <> 0 And 下限 > n_库存 And Rownum < 2;

  Json_Out := '{"output":{"code":1,"message":"成功","below_limit_lower":' || n_Count || '}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Checkstorelimit;
/


Create Or Replace Procedure Zl_Stuffsvr_Adjustdata
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：门诊费用转住院时调整卫材关联数据
  --入参：Json_In:格式
  --  input
  --    pati_id          N  1 病人ID
  --    pati_pageid      N  1 主页ID
  --    billtype         N  1 单据类型：1-收费单;2-记帐单
  --    item_list
  --      stuff_no_old       C  1 原单据号
  --      stuffdtl_id_old    N  1 原处方明细ID
  --      stuff_no_new       C  1 新单据号
  --      stuffdtl_id_new    N  1 新处方明细ID
  --出参: Json_Out,格式如下
  --  output
  --    code              N   1   应答码：0-失败；1-成功
  --    message           C   1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Jsonin   PLJson;
  j_Json     PLJson;
  j_Jsonlist Pljson_List;

  n_病人id 药品收发记录.病人id%Type;
  n_主页id 药品收发记录.主页id%Type;
  n_单据   药品收发记录.单据%Type;

  v_原单据号 未发药品记录.No%Type;
  n_原明细id 药品收发记录.费用id%Type;
  v_新单据号 药品收发记录.No%Type;
  n_新明细id 药品收发记录.费用id%Type;

  n_跟踪在用 材料特性.跟踪在用%Type;

  n_Count Number;
  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin
  --解析入参
  j_Jsonin   := PLJson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_病人id   := j_Json.Get_Number('pati_id');
  n_主页id   := j_Json.Get_Number('pati_pageid');
  n_单据     := j_Json.Get_Number('billtype');
  j_Jsonlist := j_Json.Get_Pljson_List('item_list');

  If j_Jsonlist Is Not Null Then
    n_Count := j_Jsonlist.Count;
  End If;
  If n_Count = 0 Then
    v_Err_Msg := '未传入药品单据信息!';
    Raise Err_Item;
  End If;

  For I In 1 .. j_Jsonlist.Count Loop
    j_Json     := PLJson();
    j_Json     := PLJson(j_Jsonlist.Get(I));
    v_原单据号 := j_Json.Get_String('stuff_no_old');
    n_原明细id := j_Json.Get_Number('stuffdtl_id_old');
    v_新单据号 := j_Json.Get_String('stuff_no_new');
    n_新明细id := j_Json.Get_Number('stuffdtl_id_new');
  
    Update 未发药品记录
    Set 单据 = Decode(单据, 24, 25, 单据), 主页id = n_主页id, NO = v_新单据号
    Where NO = v_原单据号 And 单据 = Decode(n_单据, 1, 24, 2, 25, 26) And 病人id = n_病人id;
  
    Update 药品收发记录
    Set 单据 = Decode(单据, 24, 25, 单据), 费用id = n_新明细id, NO = v_新单据号, 主页id = n_主页id, 费用来源 = 2, 病人来源 = 2
    Where NO = v_原单据号 And 单据 = Decode(n_单据, 1, 24, 2, 25, 26) And 费用id = n_原明细id;
  
    Select Max(跟踪在用)
    Into n_跟踪在用
    From 材料特性
    Where 材料id In (Select 药品id From 药品收发记录 Where 单据 = 25 And NO = v_新单据号 And 费用id = n_新明细id);
    If Nvl(n_跟踪在用, 0) = 1 Then
      --更新备货材料
      Update 药品收发记录
      Set 费用id = n_新明细id, 主页id = n_主页id, 费用来源 = 2, 病人来源 = 2
      Where 单据 = 21 And 费用id = n_原明细id;
    End If;
  End Loop;

  Json_Out := zlJsonOut('成功', 1);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Stuffsvr_Adjustdata;
/

Create Or Replace Procedure Zl_Stuffsvr_Getbatchnumber
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：获取卫材批号信息
  --入参：Json_In:格式
  --  input
  --    stuff_no            C  1 单据号
  --    stuffdtl_id         N  1 处方明细ID
  --    billtype          N  1 单据类型：1-收费处方,2-记帐单处方,3-记帐表处方
  --出参: Json_Out,格式如下
  --  output
  --    code              N   1   应答吗：0-失败；1-成功
  --    message           C   1   应答消息：失败时返回具体的错误信息
  --    batch_number      C   1   批号
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  v_单据号 药品收发记录.No%Type;
  n_明细id 药品收发记录.费用id%Type;
  n_单据   药品收发记录.单据%Type;
  v_批号   药品收发记录.批号%Type;
Begin

  --解析入参
  j_Jsonin := PLJson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  v_单据号 := j_Json.Get_String('stuff_no');
  n_明细id := j_Json.Get_Number('stuffdtl_id');
  n_单据   := j_Json.Get_Number('billtype');

  Select Max(批号)
  Into v_批号
  From 药品收发记录
  Where 单据 = Decode(n_单据, 1, 8, 2, 9, 10) And 费用id = n_明细id And NO = v_单据号;

  Json_Out := '{"output":{"code":1,"message":"成功","batch_number":"' || v_批号 || '"}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Getbatchnumber;
/

Create Or Replace Procedure Zl_Stuffsvr_Checkcontainstuff
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：检查单据是否含备货卫材
  --入参：Json_In:格式
  --  input
  --    stuffdtl_ids      C   0  单据明细id串,目前传入的费用ID串，用逗号分隔 ：1,2,3,4
  --出参: Json_Out,格式如下
  --  output
  --    code              N   1   应答吗：0-失败；1-成功
  --    message           C   1   应答消息：失败时返回具体的错误信息
  --    contain_stuff     N   1   是否含备货卫材:1-含有备货材料,0-不含备货材料
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  v_Ids   Varchar2(32767);
  n_Count Number(1);
Begin
  --解析入参
  j_Jsonin := PLJson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  v_Ids    := j_Json.Get_String('stuffdtl_ids');

  Select /*+cardinality(j,10)*/
   Count(1)
  Into n_Count
  From 药品收发记录 A, Table(f_Num2List(v_Ids)) J
  Where a.费用id = j.Column_Value And a.单据 = 21 And Rownum < 2;

  Json_Out := '{"output":{"code":1,"message":"成功","contain_stuff":' || Nvl(n_Count, 0) || '}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Checkcontainstuff;
/

Create Or Replace Procedure Zl_Stuffsvr_Getstuffputinbills
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：获取备货卫材的单据信息
  --入参：Json_In:格式
  --  input
  --   query_type         N   1 查询方式：
  --                                     查询入库单据：0-仅按入库单号查询,1-查询时包含入库时间,2-最后一次备货且有库存的入库单
  --                                     查询卫材收费、记账单：3-查询收费、记账单，需要配合billType使用
  --   warehouse_id       N   1 库房ID，虚拟库房
  --   stuff_no             C   1 入库单号
  --   begin_time         C   1 入库单查询开始时间
  --   end_time           C   1 入库单查询终止时间
  --   billType           N     单据类型 1-收费单;2-记账单
  --出参: Json_Out,格式如下
  --  output
  --    code              N   1   应答吗：0-失败；1-成功
  --    message           C   1   应答消息：失败时返回具体的错误信息
  --    item_list                入库记录
  --      bill_id         N   1 卫材入库单id
  --      bill_no         C   1 卫材入库单号
  --      serial_number   N     序号
  --      origin          C     产地
  --      provider        C     供应商
  --      batch_number    C     批号
  --      production_date C     生产日期
  --      expiry_date     C     效期
  --      sterilization_expiry_date C     灭菌效期
  --      putin_quantity  N     入库数量
  --      putin_price     N     入库零售价
  --      putin_money     N     入库零售金额
  --      audit_date      C     审核日期
  --      stuff_id        N   1 卫材ID
  --      batch           C     批次
  --      barcode_goods   C     商品条码
  --      barcode_inside  C     内部条码
  --      stock           N     可用库存
  --      stuffdtl_id     N     处方明细id
  ---------------------------------------------------------------------------
  Cursor c_入库信息 Is
    Select a.Id, a.No, a.序号, a.产地, c.名称 As 供应商, a.批号, a.生产日期, a.效期, a.灭菌效期, a.实际数量 As 入库数量, a.零售价 As 入库零售价,
           a.零售金额 As 入库零售金额, a.审核日期, a.药品id, a.批次, b.商品条码, b.内部条码, b.可用数量 As 库存
    From 药品收发记录 A, 药品库存 B, 供应商 C
    Where a.单据 = 15 And a.No Is Null And a.库房id Is Null And a.库房id = b.库房id And Nvl(a.批次, 0) = Nvl(b.批次, 0) And
          a.供药单位id = c.Id And Rownum < 1;
  r_入库信息 c_入库信息%RowType;

  Type Ty_库存信息 Is Ref Cursor;
  c_库存信息 Ty_库存信息; --动态游标变量

  j_Jsonin PLJson;
  j_Json   PLJson;

  n_查询方式 Number(1);
  d_开始时间 Date;
  d_终止时间 Date;
  v_入库单号 药品收发记录.No%Type;
  n_库房id   药品库存.库房id%Type;
  n_Billtype 药品收发记录.单据%Type;

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;
Begin
  --解析入参
  j_Jsonin   := PLJson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_查询方式 := j_Json.Get_Number('query_type');
  n_库房id   := j_Json.Get_Number('warehouse_id');
  v_入库单号 := j_Json.Get_String('stuff_no');
  d_开始时间 := To_Date(j_Json.Get_String('begin_time'), 'yyyy-mm-dd hh24:mi:ss');
  d_终止时间 := To_Date(j_Json.Get_String('end_time'), 'yyyy-mm-dd hh24:mi:ss');
  n_Billtype := j_Json.Get_Number('billType');
  If n_Billtype = 1 Then
    n_Billtype := 25;
  Else
    n_Billtype := 26;
  End If;

  --读取卫材入库信息
  v_Jtmp := Null;
  If Nvl(n_查询方式, 0) = 3 Then
    For c_单据信息 In (Select c.费用id, c.批次, c.商品条码, c.内部条码, Sum(b.可用数量) As 可用数量
                   From (Select a.费用id, Max(a.药品id) As 药品id, Max(a.批次) As 批次, Max(a.商品条码) As 商品条码, Max(a.内部条码) As 内部条码
                          From 药品收发记录 A
                          Where a.No = v_入库单号 And a.单据 = n_Billtype And Mod(a.记录状态, 3) In (0, 1)
                          Group By a.费用id) C, 药品库存 B
                   Where c.药品id = b.药品id(+) And b.库房id(+) = n_库房id
                   Group By c.费用id, c.批次, c.商品条码, c.内部条码) Loop
    
      v_Jtmp := v_Jtmp || ',';
      zlJsonPutValue(v_Jtmp, 'stuffdtl_id', c_单据信息.费用id, 1, 1);
      zlJsonPutValue(v_Jtmp, 'batch', c_单据信息.批次);
      zlJsonPutValue(v_Jtmp, 'barcode_goods', c_单据信息.商品条码);
      zlJsonPutValue(v_Jtmp, 'barcode_inside', c_单据信息.内部条码);
      zlJsonPutValue(v_Jtmp, 'stock', c_单据信息.可用数量, 1, 2);
    
      If Length(v_Jtmp) > 30000 Then
        If c_Jtmp Is Null Then
          c_Jtmp := Substr(v_Jtmp, 2);
        Else
          c_Jtmp := c_Jtmp || v_Jtmp;
        End If;
        v_Jtmp := Null;
      End If;
    End Loop;
  Else
    If Nvl(n_查询方式, 0) = 0 Then
      --仅按入库单号查询
      Open c_库存信息 For
        Select a.Id, a.No, a.序号, a.产地, c.名称 As 供应商, a.批号, To_Char(a.生产日期, 'yyyy-mm-dd') As 生产日期,
               To_Char(a.效期, 'yyyy-mm-dd') As 效期, To_Char(a.灭菌效期, 'yyyy-mm-dd') As 灭菌效期, a.实际数量 As 入库数量,
               LTrim(To_Char(a.零售价, '9999990.00000')) As 入库零售价, a.零售金额 As 入库零售金额, a.审核日期, a.药品id, a.批次, b.商品条码, b.内部条码,
               To_Char(b.可用数量, '9999990.00000') As 库存
        From 药品收发记录 A, 药品库存 B, 供应商 C
        Where a.单据 = 15 And a.No = v_入库单号 And a.库房id = n_库房id And a.库房id = b.库房id And Nvl(a.批次, 0) = Nvl(b.批次, 0) And
              a.供药单位id = c.Id;
    Elsif Nvl(n_查询方式, 0) = 1 Then
      --查询时包含入库时间
      Open c_库存信息 For
        Select a.Id, a.No, a.序号, a.产地, c.名称 As 供应商, a.批号, To_Char(a.生产日期, 'yyyy-mm-dd') As 生产日期,
               To_Char(a.效期, 'yyyy-mm-dd') As 效期, To_Char(a.灭菌效期, 'yyyy-mm-dd') As 灭菌效期, a.实际数量 As 入库数量,
               LTrim(To_Char(a.零售价, '9999990.00000')) As 入库零售价, a.零售金额 As 入库零售金额, a.审核日期, a.药品id, a.批次, b.商品条码, b.内部条码,
               To_Char(b.可用数量, '9999990.00000') As 库存
        From 药品收发记录 A, 药品库存 B, 供应商 C
        Where a.单据 = 15 And Decode(v_入库单号, Null, '-', a.No) = Decode(v_入库单号, Null, '-', v_入库单号) And a.库房id = n_库房id And
              (a.审核日期 Between d_开始时间 And d_终止时间) And a.库房id = b.库房id And Nvl(a.批次, 0) = Nvl(b.批次, 0) And
              a.供药单位id = c.Id;
    Else
      --最后一次备货且有库存的入库单
      Open c_库存信息 For
        Select a.Id, a.No, a.序号, a.产地, c.名称 As 供应商, a.批号, To_Char(a.生产日期, 'yyyy-mm-dd') As 生产日期,
               To_Char(a.效期, 'yyyy-mm-dd') As 效期, To_Char(a.灭菌效期, 'yyyy-mm-dd') As 灭菌效期, a.实际数量 As 入库数量,
               LTrim(To_Char(a.零售价, '9999990.00000')) As 入库零售价, a.零售金额 As 入库零售金额, a.审核日期, a.药品id, a.批次, b.商品条码, b.内部条码,
               To_Char(b.可用数量, '9999990.00000') As 库存
        From 药品收发记录 A, 药品库存 B, 供应商 C
        Where a.单据 = 15 And a.库房id = n_库房id And a.库房id = b.库房id And Nvl(a.批次, 0) = Nvl(b.批次, 0) And a.供药单位id = c.Id And
              a.No = (Select Max(NO) As NO
                      From 药品收发记录 A1, 药品库存 B1
                      Where A1.审核日期 Between Sysdate - 7 And Sysdate And A1.药品id = B1.药品id And A1.库房id = B1.库房id And
                            Nvl(A1.批次, 0) = Nvl(B1.批次, 0) And Nvl(B1.可用数量, 0) > 0 And A1.库房id = n_库房id);
    End If;
    Loop
      Fetch c_库存信息
        Into r_入库信息;
      Exit When c_库存信息%NotFound;
    
      v_Jtmp := v_Jtmp || ',';
      zlJsonPutValue(v_Jtmp, 'bill_id', r_入库信息.Id, 1, 1);
      zlJsonPutValue(v_Jtmp, 'bill_no', r_入库信息.No);
      zlJsonPutValue(v_Jtmp, 'serial_number', r_入库信息.序号, 1);
      zlJsonPutValue(v_Jtmp, 'origin', r_入库信息.产地);
      zlJsonPutValue(v_Jtmp, 'provider', r_入库信息.供应商);
      zlJsonPutValue(v_Jtmp, 'batch_number', r_入库信息.批号);
      zlJsonPutValue(v_Jtmp, 'production_date', r_入库信息.生产日期);
      zlJsonPutValue(v_Jtmp, 'expiry_date', r_入库信息.效期);
      zlJsonPutValue(v_Jtmp, 'sterilization_expiry_date', r_入库信息.灭菌效期);
      zlJsonPutValue(v_Jtmp, 'putin_quantity', r_入库信息.入库数量, 1);
      zlJsonPutValue(v_Jtmp, 'putin_price', r_入库信息.入库零售价, 1);
      zlJsonPutValue(v_Jtmp, 'putin_money', r_入库信息.入库零售金额, 1);
      zlJsonPutValue(v_Jtmp, 'audit_date', r_入库信息.审核日期);
      zlJsonPutValue(v_Jtmp, 'stuff_id', r_入库信息.药品id, 1);
      zlJsonPutValue(v_Jtmp, 'batch', r_入库信息.批次);
      zlJsonPutValue(v_Jtmp, 'barcode_goods', r_入库信息.商品条码);
      zlJsonPutValue(v_Jtmp, 'barcode_inside', r_入库信息.内部条码);
      zlJsonPutValue(v_Jtmp, 'stock', r_入库信息.库存, 1, 2);
    
      If Length(v_Jtmp) > 30000 Then
        If c_Jtmp Is Null Then
          c_Jtmp := Substr(v_Jtmp, 2);
        Else
          c_Jtmp := c_Jtmp || v_Jtmp;
        End If;
        v_Jtmp := Null;
      End If;
    End Loop;
  End If;

  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Getstuffputinbills;
/

Create Or Replace Procedure Zl_Stuffsvr_Billaffirm
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：卫材处方记帐确认或卫材处方收费确认
  --入参：Json_In:格式
  --  input
  --    pati_id                   N  0 病人id：针对收费时有效
  --    pati_name                 C  0 姓名：针对收费时有效
  --    pati_sex                  C  0 性别：针对收费时有效
  --    pati_age                  C  0 年龄：针对收费时有效
  --    pati_outpno               C  0 门诊号：针对收费时有效
  --    auditor                   C  1 审核人
  --    auditor_code              C  1 审核人编号
  --    audit_time                C  1 审核时间：yyyy-mm-dd hh24:mi:ss
  --    item_list[]                   更新数据列表[数组]
  --      billtype                N  1 单据类型:1 -收费处方发药  ;2- 记帐单处方发药;3- 记帐表处方发药
  --      stuff_no                C  1 单据号
  --      stuffdtl_ids            C  0 费用ID,可以传入多个,用逗号分离
  --      stuff_auto_send         N  0  卫材自动发料;0-不自动发料;1-自动发料
  --      auto_send_ids           C  0  自动发料的明细ids,多个用逗号分隔
  --出参: Json_Out,格式如下
  --  output
  --    code                      N   1   应答吗：0-失败；1-成功
  --    message                   C   1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Jsonin      PLJson;
  j_Json        PLJson;
  j_Jsonlist_In Pljson_List;

  n_病人id 药品收发记录.病人id%Type;
  v_姓名   药品收发记录.姓名%Type;
  v_性别   药品收发记录.性别%Type;
  v_年龄   药品收发记录.年龄%Type;
  v_门诊号 Number(18);
  --卫材其它出库单摘要，格式：病人姓名:XXX    性别:XXX    年龄XXX    门诊号:XXX
  v_摘要     药品收发记录.摘要%Type;
  v_单据号   药品收发记录.No%Type;
  v_明细ids  Varchar2(4000);
  n_单据     药品收发记录.单据%Type;
  n_单据_In  药品收发记录.单据%Type;
  d_审核时间 Date;
  n_自动发料 Number(1);
  v_Err      Varchar2(255);

  v_Nos         Varchar2(32767);
  v_发料明细id  Varchar2(400);
  v_发料明细ids Varchar2(4000);
  v_审核人      人员表.姓名%Type;
  v_审核人编号  人员表.编号%Type;
  Err_Custom Exception;
Begin
  --解析入参
  j_Jsonin      := PLJson(Json_In);
  j_Json        := j_Jsonin.Get_Pljson('input');
  n_病人id      := j_Json.Get_Number('pati_id');
  v_姓名        := j_Json.Get_String('pati_name');
  v_性别        := j_Json.Get_String('pati_sex');
  v_年龄        := j_Json.Get_String('pati_age');
  v_门诊号      := j_Json.Get_String('pati_outpno');
  v_审核人      := j_Json.Get_String('auditor');
  v_审核人编号  := j_Json.Get_String('auditor_code');
  d_审核时间    := To_Date(j_Json.Get_String('audit_time'), 'yyyy-mm-dd hh24:mi:ss');
  j_Jsonlist_In := j_Json.Get_Pljson_List('item_list');

  If Nvl(n_病人id, 0) <> 0 Then
    v_摘要 := '病人姓名:' || v_姓名 || '    ' || '性别:' || v_性别 || '    ' || '年龄' || v_年龄 || '    ' || '门诊号:' || v_门诊号;
  End If;

  If d_审核时间 Is Null Then
    d_审核时间 := Sysdate;
  End If;

  If j_Jsonlist_In.Count = 0 Then
    v_Err := '未传入卫材单据信息！';
    Raise Err_Custom;
  End If;

  For I In 1 .. j_Jsonlist_In.Count Loop
    j_Json       := PLJson();
    j_Json       := PLJson(j_Jsonlist_In.Get(I));
    n_单据       := j_Json.Get_Number('billtype');
    v_单据号     := j_Json.Get_String('stuff_no');
    v_明细ids    := j_Json.Get_String('stuffdtl_ids');
    n_自动发料   := j_Json.Get_Number('stuff_auto_send');
    v_发料明细id := j_Json.Get_String('auto_send_ids');
  
    n_单据_In := n_单据;
    If n_单据 = 1 Then
      n_单据 := 24;
    Elsif n_单据 = 2 Then
      n_单据 := 25;
    Elsif n_单据 = 3 Then
      n_单据 := 26;
    Else
      v_Err := '传入单据类型无效，请检查！';
      Raise Err_Custom;
    End If;
  
    If Nvl(v_单据号, '-') = '-' Then
      v_Err := '未传入处方单号，请检查！';
      Raise Err_Custom;
    End If;
  
    If Nvl(n_自动发料, 0) = 1 Then
      If v_发料明细id Is Null Then
        v_Nos := v_Nos || ',' || v_单据号;
      Else
        v_发料明细ids := v_发料明细ids || ',' || v_发料明细id;
      End If;
    End If;
  
    If Nvl(v_明细ids, '-') = '-' Then
      v_Err := '未传入处方明细ID，请检查！';
      Raise Err_Custom;
    End If;
  
    Zl_卫材收发记录_费用审核(n_单据, v_单据号, v_明细ids, d_审核时间, n_病人id, v_姓名, v_性别, v_年龄, v_摘要);
  End Loop;

  --按处方号发料
  If Nvl(v_Nos, '-') <> '-' Then
    v_Nos := Substr(v_Nos, 2);
    Zl_卫材收发记录_自动发料_s(n_单据, v_审核人, v_审核人编号, v_Nos, 0);
  End If;

  --按费用ID发料
  If Nvl(v_发料明细ids, '-') <> '-' Then
    v_发料明细ids := Substr(v_发料明细ids, 2);
    Zl_卫材收发记录_自动发料_s(n_单据, v_审核人, v_审核人编号, v_发料明细ids, 1);
  End If;

  Json_Out := zlJsonOut('成功', 1);
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Stuffsvr_Billaffirm;
/

Create Or Replace Procedure Zl_Stuffsvr_Autosendstuff
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：卫生自动发料（按NO或NO明细）
  --入参：Json_In:格式
  --  input
  --    billtype             N 1 单据类型: 1-收费处方发料；2-记帐单处方发料；3-记帐表处方发料
  --    operator_name        C 1 操作员姓名
  --    operator_code        C 1 操作员编号
  --    stuff_nos            C 1 单据号串：NO1,NO2...
  --    stuffdtl_ids         C 1 单据明细id串,目前传入的费用ID串，用逗号分隔 ：1,2,3,4
  --    send_type            N 1 发药类型,0-按 传入的no单据类型发药,1-只按 处理方明细id串发药
  --出参: Json_Out,格式如下
  --  output
  --    code                 N 1 应答吗：0-失败；1-成功
  --    message              C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  n_单据       药品收发记录.单据%Type;
  v_操作员姓名 人员表.姓名%Type := Null;
  v_操作员编号 人员表.编号%Type := Null;
  v_Err        Varchar2(255);
  v_Nos        Varchar2(32767);
  Err_Custom Exception;
  v_Ids       Varchar2(32767);
  n_Send_Type Number;
Begin
  --解析入参 
  j_Jsonin    := PLJson(Json_In);
  j_Json      := j_Jsonin.Get_Pljson('input');
  n_Send_Type := j_Json.Get_Number('send_type');
  If Nvl(n_Send_Type, 0) = 0 Then
    n_单据 := j_Json.Get_Number('billtype');
    If n_单据 = 1 Then
      n_单据 := 24;
    Elsif n_单据 = 2 Then
      n_单据 := 25;
    Elsif n_单据 = 3 Then
      n_单据 := 26;
    Else
      v_Err := '传入节点【billtype】错误，请检查！';
      Raise Err_Custom;
    End If;
    v_Nos := j_Json.Get_String('stuff_nos');
    If j_Json.Exist('stuffdtl_ids') Then
      v_Ids := j_Json.Get_String('stuffdtl_ids');
    End If;
    If v_Ids Is Null And v_Nos Is Null Then
      v_Err := '未传入卫材单据【rcp_nos】节点或明细信息【stuffdtl_ids】节点';
      Raise Err_Custom;
      Return;
    End If;
  End If;

  If Nvl(n_Send_Type, 0) = 1 Then
    If j_Json.Exist('stuffdtl_ids') Then
      v_Ids := j_Json.Get_String('stuffdtl_ids');
    End If;
    If v_Ids Is Null Then
      v_Err := '未传入卫材明细信息【stuffdtl_ids】节点';
      Raise Err_Custom;
      Return;
    End If;
  End If;

  v_操作员姓名 := j_Json.Get_String('operator_name');
  v_操作员编号 := j_Json.Get_String('operator_code');

  --按单据号发料
  If v_Nos Is Not Null Then
    Zl_卫材收发记录_自动发料_s(n_单据, v_操作员姓名, v_操作员编号, v_Nos, 0);
  End If;

  --按单据ID发料
  If v_Ids Is Not Null Then
    Zl_卫材收发记录_自动发料_s(n_单据, v_操作员姓名, v_操作员编号, v_Ids, 1);
  End If;

  Json_Out := '{"output":{"code":1,"message": "成功"}}';
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Stuffsvr_Autosendstuff;
/


Create Or Replace Procedure Zl_Stuffsvr_Autoreturnstuff
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：自动退料（按处方明细即费用ID退料，默认是全退）
  --入参：Json_In:格式
  --  input
  --    audit_operator        C 1 审核人
  --    stuffdtl_ids           单据明细id,目前传入的费用ID,包括退料数量(冒号加逗号组合)，数量为空表示全部退料 ：费用id1:数量1,费用id2:数量2...
  --出参: Json_Out,格式如下
  --  output
  --    code                 N 1   应答吗：0-失败；1-成功
  --    message              C 1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  d_操作时间   药品收发记录.审核日期%Type;
  v_操作员姓名 人员表.姓名%Type := Null;
  v_Err        Varchar2(255);
  Err_Custom Exception;
  v_Ids      Clob;
  n_费用id   Number;
  n_数量     药品收发记录.实际数量%Type;
  n_退药数量 药品收发记录.实际数量%Type;
  v_Tmp      Clob;
  v_Field    Varchar2(32767);
Begin
  --解析入参  
  j_Jsonin     := PLJson(Json_In);
  j_Json       := j_Jsonin.Get_Pljson('input');
  v_操作员姓名 := j_Json.Get_String('audit_operator');

  If j_Json.Exist('stuffdtl_ids') Then
    v_Ids := j_Json.Get_Clob('stuffdtl_ids');
  End If;

  If v_Ids Is Null Then
    v_Err := '未处方明细信息【stuffdtl_ids】节点';
    Raise Err_Custom;
  End If;

  d_操作时间 := Sysdate;

  v_Tmp := v_Ids || ',';
  While Length(v_Tmp) <> 0 Loop
    v_Field  := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    n_费用id := To_Number(Substr(v_Field, 1, Instr(v_Field, ':') - 1));
    n_数量   := Substr(v_Field, Instr(v_Field, ':') + 1);
    v_Tmp    := Replace(',' || v_Tmp, ',' || v_Field || ',');
  
    If n_数量 Is Not Null Then
      n_退药数量 := n_数量;
    End If;
  
    --分解退药数量
    For r_处方明细 In (Select a.库房id, a.Id, Nvl(a.付数, 1) * a.实际数量 As 数量
                   From 药品收发记录 A
                   Where a.单据 In (24, 25, 26) And (a.记录状态 = 1 Or Mod(a.记录状态, 3) = 0) And 审核人 Is Not Null And
                         a.费用id = n_费用id
                   Order By a.库房id, a.药品id, a.批次) Loop
    
      If n_数量 Is Null Then
        --传入的数量为空表示全退
      
        --部门退料（处方明细）
        Zl_材料收发记录_部门退料_s(r_处方明细.Id, v_操作员姓名, d_操作时间, Null, Null, Null, Null, 0, v_操作员姓名);
      Else
        If n_退药数量 > 0 Then
          If n_退药数量 > r_处方明细.数量 Then
            --部门退料（处方明细）
            Zl_材料收发记录_部门退料_s(r_处方明细.Id, v_操作员姓名, d_操作时间, Null, Null, Null, r_处方明细.数量, 0, v_操作员姓名);
          
            n_退药数量 := n_退药数量 - r_处方明细.数量;
          Else
            --部门退料（处方明细）
            Zl_材料收发记录_部门退料_s(r_处方明细.Id, v_操作员姓名, d_操作时间, Null, Null, Null, n_退药数量, 0, v_操作员姓名);
          
            n_退药数量 := 0;
          End If;
        End If;
      End If;
    End Loop;
  End Loop;

  Json_Out := '{"output":{"code":1,"message": "成功"}}';

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Stuffsvr_Autoreturnstuff;
/

Create Or Replace Procedure Zl_Stuffsvr_Newbill
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：主要是在记帐（含划价）， 收费(含划价)后产生新的处方或药嘱记录
  --入参：Json_In:格式
  --  input
  --     billtype             N   1 单据类型: 1 -收费处方  ;2- 记帐单处方;3- 记帐表处方
  --     pati_source          N   1 病人来源:1-门诊;2-住院;4-体检

  ---------------------------billtype = 3,记帐表处方（药品收发记录.单据=26）时，无以下节点--------------------------------------
  --     pati_id                    N   1 病人ID
  --     pati_pageid                N   1 主页ID
  --     pati_name                  C   1 病人姓名
  --     pati_sex_code              C   1 性别编号（新门诊)
  --     pati_sex                   C   1 性别
  --     pati_age                   C   1 年龄
  --     pati_identity              C     身份
  --     pati_birthdate             C     出生日期:yyyy-mm-dd hh:mi:ss
  --     pati_idcard                C     身份证号
  --     pati_deptid                N   1 病人科室ID
  --     pati_wardarea_id           N     病人病区ID
  ---------------------------billtype = 3,记帐表处方（药品收发记录.单据=26）时，无以上节点-----------------------------------------

  --     bill_list[]                      更新数据列表[数组]
  --        stuff_no                  C  1 NO
  --        charge_tag                N  1 收费标志:0-未收费或记帐划价;1-已收费或记帐
  --        fee_acnter                C    划价人
  --        plcdept_id                C    开单科室id（新门诊)
  --        plcdept                   C    开单科室名称（新门诊)
  --        placer_id                 C    开单医师id（新门诊)
  --        placer                    C    开单医师（新门诊)  增加
  --        apply_fee_category_code   C    申请单费别编码(医疗付款方式编码)(新门诊) 增加；
  --        apply_fee_category_name   C    申请单费别名称（医疗付款方式名称）(新门诊) 增加；
  --        operator_name             C  1 操作员姓名
  --        operator_code             C  1 操作员编号
  --        create_time               C  1 登记时间:yyyy-mm-dd hh:mi:ss
  --        item_list[]                    更新数据列表[数组]

  ---------------------------billtype = 3,记帐表处方（药品收发记录.单据=26）时，有以下节点----------------------------------------
  --           pati_id                 N  1 病人ID
  --           pati_pageid             N    主页ID
  --           pati_name               C  1 病人姓名
  --           pati_sex                C  1 性别
  --           pati_age                C  1 年龄
  --           pati_identity           C    身份
  --           pati_birthdate          C    出生日期:yyyy-mm-dd hh:mi:ss
  --           pati_idcard             C    身份证号
  --           pati_wardarea_id        N    病人病区ID
  --           pati_deptid             N  1 病人科室ID
  ---------------------------billtype = 3,记帐表处方（药品收发记录.单据=26）时，有以上节点-----------------------------------------
  --           stuffdtl_id             N  1 处方明细ID
  --           serial_num              N  1 序号
  --           warehouse_id            N  1 库房ID
  --           is_bakstuff             N  1 是否备货卫材:有高值卫材才需要传入，非0表示是高值卫材模式(如扫码时使用)
  --           bakstuff_batch             1 备货材料批次
  --           stuff_id                N  1 材料ID
  --           baby_num                N    婴儿序号

  ---------------------------以下节点为可选参数，医嘱记录产生-----------------------------------------------
  --           advice_id               N  0 医嘱ID
  --           emergency_tag           N    医嘱记录中的紧急标志(0-普通;1-紧急;2-补录(对门诊无效))
  --           effectivetime           N  0 医嘱期效
  --           freq_name               C  0 频次名称
  --           single                  N  0 单量
  ---------------------------以上节点为可选参数，医嘱记录产生-----------------------------------------------

  --           packages_num            N  1 付数
  --           outbound_num            N  1 出库数量
  --           price                   N    售价
  --           warehouse_window        C  0 发料窗口
  --           memo                    C  0 摘要
  --           fee_source              N  0 费用来源
  --           stuff_auto_send         N  0 卫材自动发料;0-不自动发料;1-自动发料

  --出参: Json_Out,格式如下
  --  output
  --    code                 N   1   应答吗：0-失败；1-成功
  --    message              C   1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
Begin
  --直接调用卫材业务过程（入参格式一致）
  Zl_药品收发记录_Newstuffbill(Json_In, Json_Out);

  Json_Out := Zljsonout('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Stuffsvr_Newbill;
/

Create Or Replace Procedure Zl_Stuffsvr_Checkexistsbarcode
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：检查指定卫材在药品库存中是否存在商品条码或内部条码
  --入参：Json_In:格式
  --  input
  --    stuff_ids           C 1 输入的卫材ids,多个用逗号分隔
  --    barcode             C 0 当前查询的条码串
  --    only_barcode_inside N 0 1-仅对内部条码进行查找,0-对商品条码及内部条码进行查找
  --出参: Json_Out,格式如下
  --  output
  --    code                N 1   应答吗：0-失败；1-成功
  --    message             C 1   应答消息：失败时返回具体的错误信息
  --    stuff_ids           C 1  返回的药品库存中存在商品条码或内部条码的卫材ids,多个用逗号分隔
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  v_卫材ids Varchar2(4000);
  v_条码    药品库存.商品条码%Type;
  n_Inside  Number(1);
  v_Temp    Varchar2(4000);
Begin
  --解析入参
  j_Jsonin  := PLJson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  v_卫材ids := j_Json.Get_String('stuff_ids');
  v_条码    := j_Json.Get_String('barcode');
  n_Inside  := Nvl(j_Json.Get_Number('only_barcode_inside'), 0);

  v_Temp := Null;
  If v_条码 Is Null Then
    For c_条码 In (Select /*+cardinality(b,10) */
                 Distinct 药品id
                 From 药品库存 A, Table(f_Str2List(v_卫材ids)) B
                 Where (Nvl(n_Inside, 0) = 0 And a.商品条码 Is Not Null Or a.内部条码 Is Not Null) And a.可用数量 > 0 And
                       (a.效期 Is Null Or a.效期 > Trunc(Sysdate)) And a.药品id = b.Column_Value And Not Exists
                  (Select 1
                        From 药品库存
                        Where 药品id = a.药品id And (Nvl(n_Inside, 0) = 1 Or 商品条码 Is Null) And 内部条码 Is Null)) Loop
      v_Temp := v_Temp || ',' || c_条码.药品id;
    End Loop;
  Else
    For c_条码 In (Select /*+cardinality(b,10) */
                 Distinct 药品id
                 From 药品库存 A, Table(f_Str2List(v_卫材ids)) B
                 Where (Nvl(n_Inside, 0) = 0 And a.商品条码 Is Not Null Or a.内部条码 Is Not Null) And a.可用数量 > 0 And
                       (a.效期 Is Null Or a.效期 > Trunc(Sysdate)) And a.药品id = b.Column_Value And Not Exists
                  (Select 1
                        From 药品库存
                        Where 药品id = a.药品id And (Nvl(n_Inside, 0) = 0 And Nvl(商品条码, '-') = v_条码 Or Nvl(内部条码, '-') = v_条码)) And
                       Not Exists
                  (Select 1
                        From 药品库存
                        Where 药品id = a.药品id And (Nvl(n_Inside, 0) = 1 Or 商品条码 Is Null) And 内部条码 Is Null)) Loop
      v_Temp := v_Temp || ',' || c_条码.药品id;
    End Loop;
  End If;

  Json_Out := '{"output":{"code":1,"message":"成功","stuff_ids":"' || Substr(v_Temp, 2) || '"}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Checkexistsbarcode;
/

Create Or Replace Procedure Zl_Stuffsvr_Delbill
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：卫材处方销帐(或退费)
  --入参：Json_In:格式
  -- input
  --     billtype                 N   1   单据类型:1 -收费处方发料  ;2- 记帐单处方发料
  --     stuff_no                 C   1   单据号,有该节点是按整个NO进行销帐（暂时只用由于新门诊系统接口）
  --     item_list[]                更新数据列表[数组]
  --          stuffdtl_id         N 1 处方明细id,目前传入的费用ID
  --          return_num          N 1 退料数量

  --     return_list[]自动退药列表
  --           audit_operator        C 1 审核人
  --           operator_time         C 1 操作时间
  --           stuffdtl_ids          C 1 单据明细id,目前传入的费用ID,包括退料数量(冒号加逗号组合)，数量为空表示全部退料 ：费用id1:数量1,费用id2:数量2...

  --出参: Json_Out,格式如下
  -- output
  --    code                 N 1 应答吗：0-失败；1-成功
  --    message              C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Json     Pljson;
  o_Json     Pljson;
  j_Jsonlist Pljson_List;
  j_Item     Pljson;

  n_退料数量   药品收发记录.实际数量%Type;
  n_单据明细id 药品收发记录.费用id%Type;
  d_操作时间   药品收发记录.审核日期%Type;
  v_操作员姓名 人员表.姓名%Type;
  v_Ids        Clob;
  n_费用id     Number;
  n_数量       药品收发记录.实际数量%Type;
  n_性质       Number(1);
  v_No         药品收发记录.No%Type;
  n_销帐数量   药品收发记录.实际数量%Type;

  v_Tmp   Clob;
  v_Field Varchar2(32767);
  v_Err   Varchar2(255);
  Err_Custom Exception;
Begin
  --解析入参
  j_Json := Pljson(Json_In);
  o_Json := j_Json.Get_Pljson('input');

  --按NO退料和销账，目前用于新门诊接口
  v_No := o_Json.Get_String('stuff_no');
  If v_No Is Not Null Then
    n_性质 := o_Json.Get_Number('billtype');
  
    For r_卫材明细 In (Select 费用id, Sum(Nvl(付数, 1) * 实际数量) As 退料数量
                   From 药品收发记录
                   Where 单据 = Decode(n_性质, 1, 24, 2, 25, 26) And NO = v_No And 审核日期 Is Null
                   Group By 费用id
                   Order By 费用id) Loop
      --按费用id进行销账
      Zl_卫材收发记录_销售退费_s(r_卫材明细.费用id, r_卫材明细.退料数量, 1);
    End Loop;
  End If;

  --自动退料
  j_Jsonlist := Pljson_List();
  j_Jsonlist := o_Json.Get_Pljson_List('return_list');
  If j_Jsonlist Is Not Null Then
    For I In 1 .. j_Jsonlist.Count Loop
      j_Item       := Pljson();
      j_Item       := Pljson(j_Jsonlist.Get(I));
      v_操作员姓名 := j_Item.Get_String('audit_operator');
      d_操作时间   := To_Date(j_Item.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');
      v_Ids        := j_Item.Get_Clob('stuffdtl_ids');
      v_Tmp        := v_Ids || ',';
      While Length(v_Tmp) <> 0 Loop
        v_Field  := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
        n_费用id := To_Number(Substr(v_Field, 1, Instr(v_Field, ':') - 1));
        n_数量   := Substr(v_Field, Instr(v_Field, ':') + 1);
        v_Tmp    := Replace(',' || v_Tmp, ',' || v_Field || ',');
      
        If n_数量 Is Not Null Then
          n_退料数量 := n_数量;
        End If;
      
        --分解退料数量
        For r_卫材明细 In (Select a.库房id, a.Id, Nvl(a.付数, 1) * a.实际数量 As 数量
                       From 药品收发记录 A
                       Where a.单据 In (24, 25, 26) And (a.记录状态 = 1 Or Mod(a.记录状态, 3) = 0) And 审核人 Is Not Null And
                             a.费用id = n_费用id
                       Order By a.库房id, a.药品id, a.批次) Loop
        
          If n_数量 Is Null Then
            --传入的数量为空表示全退
            --部门退料（处方明细）
            Zl_材料收发记录_部门退料_s(r_卫材明细.Id, v_操作员姓名, d_操作时间, Null, Null, Null, Null, 0, v_操作员姓名);
          Else
            If n_退料数量 > 0 Then
              If n_退料数量 > r_卫材明细.数量 Then
                --部门退料（处方明细）
                Zl_材料收发记录_部门退料_s(r_卫材明细.Id, v_操作员姓名, d_操作时间, Null, Null, Null, r_卫材明细.数量, 0, v_操作员姓名);
              
                n_退料数量 := n_退料数量 - r_卫材明细.数量;
              Else
                --部门退料（处方明细）
                Zl_材料收发记录_部门退料_s(r_卫材明细.Id, v_操作员姓名, d_操作时间, Null, Null, Null, n_退料数量, 0, v_操作员姓名);
              
                n_退料数量 := 0;
              End If;
            End If;
          End If;
        End Loop;
      End Loop;
    End Loop;
  End If;

  --删卫材单据
  n_单据明细id := Null;
  j_Jsonlist   := Pljson_List();
  j_Jsonlist   := o_Json.Get_Pljson_List('item_list');
  If j_Jsonlist Is Not Null Then
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json       := Pljson();
      o_Json       := Pljson(j_Jsonlist.Get(I));
      n_退料数量   := o_Json.Get_Number('return_num');
      n_单据明细id := o_Json.Get_Number('stuffdtl_id');
    
      If n_单据明细id Is Null Then
        v_Err := '传入节点【stuffdtl_id】错误，请检查！';
        Raise Err_Custom;
      End If;
    
      If n_退料数量 Is Null Then
        v_Err := '传入节点【return_num】错误，请检查！';
        Raise Err_Custom;
      End If;
      Zl_卫材收发记录_销售退费_s(n_单据明细id, n_退料数量, 1);
    End Loop;
  End If;
  Json_Out := Zljsonout('成功', 1);
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Stuffsvr_Delbill;
/


Create Or Replace Procedure Zl_Stuffsvr_Getbakstockbatch
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：批量获取多个卫材库存及价格信息:在项目选择器中展示库存及价格信息
  --入参：Json_In:格式
  --  input
  --   stuff_ids            C   1   卫材ID，多个用英文的逗号分隔
  --   warehouse_id         C   0   库房ID
  --出参: Json_Out,格式如下
  --  output
  --    code                N 1 应答吗：0-失败；1-成功
  --    message             C 1 应答消息：失败时返回具体的错误信息
  --    item_list
  --      stuff_id            N   1   卫材ID
  --      stock               N   1   可用数量
  --      batch               N   0   批次
  --      batch_number        C   0   批号
  --      barcode_goods       C   0   商品条码
  --      barcode_inside      C   0   内部条码
  --      provider            C   0   供应商
  ---------------------------------------------------------------------------
  j_Jsonin  PLJson;
  j_Json    PLJson;
  c_卫材ids Clob;
  n_库房id  药品库存.库房id%Type;
  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_卫材ids Collection_Type;
  I           Number;

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;
Begin
  --解析入参
  j_Jsonin  := PLJson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  c_卫材ids := j_Json.Get_Clob('stuff_ids');
  n_库房id  := j_Json.Get_String('warehouse_id');

  I := 0;
  While c_卫材ids Is Not Null Loop
    If Length(c_卫材ids) <= 4000 Then
      Col_卫材ids(I) := c_卫材ids;
      c_卫材ids := Null;
    Else
      Col_卫材ids(I) := Substr(c_卫材ids, 1, Instr(c_卫材ids, ',', 3980) - 1);
      c_卫材ids := Substr(c_卫材ids, Instr(c_卫材ids, ',', 3980) + 1);
    End If;
    I := I + 1;
  End Loop;

  For I In 0 .. Col_卫材ids.Count - 1 Loop
    For c_库存 In (Select /*+cardinality(b,10)*/
                  a.药品id, Sum(Nvl(a.可用数量, 0)) As 库存, a.批次, a.上次批号 As 批号, a.商品条码, a.内部条码, Max(c.名称) As 供应商
                 From 药品库存 A, Table(f_Num2List(Col_卫材ids(I))) B, 供应商 C
                 Where a.药品id = b.Column_Value And a.性质 = 1 And a.上次供应商id = c.Id(+) And a.库房id = n_库房id And
                       (a.效期 Is Null Or a.效期 > Trunc(Sysdate))
                 Group By a.药品id, a.库房id, a.批次, a.上次批号, a.商品条码, a.内部条码
                 Having Sum(Nvl(a.可用数量, 0)) > 0
                 Order By a.药品id, a.批次) Loop
    
      v_Jtmp := v_Jtmp || ',';
      zlJsonPutValue(v_Jtmp, 'stuff_id', c_库存.药品id, 1, 1);
      zlJsonPutValue(v_Jtmp, 'stock', c_库存.库存, 1);
      zlJsonPutValue(v_Jtmp, 'batch', c_库存.批次, 1);
      zlJsonPutValue(v_Jtmp, 'batch_number', c_库存.批号);
      zlJsonPutValue(v_Jtmp, 'barcode_goods', c_库存.商品条码);
      zlJsonPutValue(v_Jtmp, 'barcode_inside', c_库存.内部条码);
      zlJsonPutValue(v_Jtmp, 'provider', c_库存.供应商, 0, 2);
    
      If Length(v_Jtmp) > 30000 Then
        If c_Jtmp Is Null Then
          c_Jtmp := Substr(v_Jtmp, 2);
        Else
          c_Jtmp := c_Jtmp || v_Jtmp;
        End If;
        v_Jtmp := Null;
      End If;
    End Loop;
  End Loop;

  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Getbakstockbatch;
/

Create Or Replace Procedure Zl_Stuffsvr_Getbill
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：执行医嘱超期收回时所涉及的检查和查询通过费用id获相关信息
  --入参      json
  --input     
  --  fee_ids                                    C 1 费用id拼串，逗号分割 
  --出参      json
  --output      
  --  code                                       C 1 应答码：0-失败；1-成功
  --  message                                    C 1 应答消息：成功时返回成功信息失败时返回具体的错误信息
  --  item_list                                     时价明细信息，支持多个，[数组]
  --    fee_id                                   N 1 费用id
  --    order_total_qunt                         N 1 数量
  --    drug_reocrd_id                           N 1 药品收发id
  --    billtype                                 N 1 单据
  --    drug_id                                  N 1 药品id
  --    obj_dept_id                              N 1 对方部门id
  --    warehouse_id                             N 1 库房id
  --    batch                                    N 1 批次
  --    batch_number                             C 1 批号
  --    expiry_date                              C 1 期效
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  v_费用ids Varchar(4000);

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;
Begin
  --解析入参
  j_Jsonin  := PLJson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  v_费用ids := j_Json.Get_String('fee_ids');

  v_Jtmp := Null;
  For R In (Select Nvl(b.付数, 1) * b.实际数量 As 数量, b.Id As 收发id, b.单据, b.药品id, b.对方部门id, b.库房id, b.费用id, b.批次, b.批号,
                   To_Char(b.效期, 'yyyy-mm-dd hh24:mi:ss') As 效期
            From 药品收发记录 B
            Where b.费用id In (Select /*+cardinality(x,10)*/
                              x.Column_Value
                             From Table(Cast(f_Num2List(v_费用ids) As Zltools.t_Numlist)) X) And
                  (b.记录状态 = 1 Or Mod(b.记录状态, 3) = 0)
            Order By b.费用id) Loop
  
    v_Jtmp := v_Jtmp || ',';
    zlJsonPutValue(v_Jtmp, 'fee_id', r.费用id, 1, 1);
    zlJsonPutValue(v_Jtmp, 'order_total_qunt', r.数量, 1);
    zlJsonPutValue(v_Jtmp, 'drug_reocrd_id', r.收发id, 1);
    zlJsonPutValue(v_Jtmp, 'billtype', r.单据, 1);
    zlJsonPutValue(v_Jtmp, 'drug_id', r.药品id, 1);
  
    zlJsonPutValue(v_Jtmp, 'obj_dept_id', r.对方部门id, 1);
    zlJsonPutValue(v_Jtmp, 'warehouse_id', r.库房id, 1);
    zlJsonPutValue(v_Jtmp, 'batch', r.批次, 1);
    zlJsonPutValue(v_Jtmp, 'batch_number', r.批号, 0);
    zlJsonPutValue(v_Jtmp, 'expiry_date', r.效期, 1, 2);
  
    If Length(v_Jtmp) > 30000 Then
      If c_Jtmp Is Null Then
        c_Jtmp := Substr(v_Jtmp, 2);
      Else
        c_Jtmp := c_Jtmp || v_Jtmp;
      End If;
      v_Jtmp := Null;
    End If;
  End Loop;

  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Getbill;
/

Create Or Replace Procedure Zl_Stuffsvr_Getvirwarehouse
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：获取虚拟库房
  --入参：Json_In:格式
  --  input
  --出参: Json_Out,格式如下
  --  output
  --    code                        N   1   应答吗：0-失败；1-成功
  --    message                     C   1   应答消息：失败时返回具体的错误信息
  --    item_list     [数组]
  --        dept_id                 N   1   科室ID
  --        warehouse_id            N   1   库房ID
  --        vir_warehouse_id    N   1   虚拟库房ID
  ---------------------------------------------------------------------------

  v_List Varchar2(32767);
Begin
  --解析入参
  For r_Data In (Select Distinct 科室id, 库房id, 虚拟库房id From 虚拟库房对照) Loop
    zlJsonPutValue(v_List, 'dept_id', r_Data.科室id, 1, 1);
    zlJsonPutValue(v_List, 'warehouse_id', r_Data.库房id, 1);
    zlJsonPutValue(v_List, 'vir_warehouse_id', r_Data.虚拟库房id, 1, 2);
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || v_List || ']}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Getvirwarehouse;
/

Create Or Replace Procedure Zl_Stuffsvr_Getnotsendrec
(
  Json_In  Varchar,
  Json_Out Out Clob
) Is
  ---------------------------------------------------------------------------
  --功能：获取未发料记录
  --入参：JSON格式
  --input
  --  billtypes             C  1 单据类型，多个用英文逗号分隔:  1-收费发料单;2-记帐发料单;3-记帐表发料单
  --  charge_tag            N  1 收费标志:0-未收费或记帐划价;1-已收费或记帐
  --  fee_source            C  1 费用来源，多个用英文逗号分隔:1-门诊,2-住院,4-体检
  --  start_time            C  0 开始时间:yyyy-mm-dd hh:mi:ss
  --  end_time              C  0 结束时间:yyyy-mm-dd hh:mi:ss
  --出参：JSON格式
  --output
  --  code  N  1  应答吗：0-失败；1-成功
  --  message C 1 应答消息：失败时返回具体的错误信息
  --  item_list[]
  --    billtype              N  1 单据类型: 1-收费发料单;2-记帐发料单;3-记帐表发料单
  --    stuff_no              C  1 发料单号
  --    warehouse_id          N  1 库房ID
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  v_单据类型 Varchar2(100);
  n_收费标志 Number(1);
  v_费用来源 Varchar2(100);
  d_开始时间 Date;
  d_结束时间 Date;

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;
Begin
  --解析入参
  j_Jsonin   := PLJson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  v_单据类型 := j_Json.Get_String('billtypes');
  n_收费标志 := j_Json.Get_Number('charge_tag');
  v_费用来源 := j_Json.Get_String('fee_source');
  d_开始时间 := To_Date(j_Json.Get_String('start_time'), 'yyyy-mm-dd hh24:mi:ss');
  d_结束时间 := To_Date(j_Json.Get_String('end_time'), 'yyyy-mm-dd hh24:mi:ss');

  v_单据类型 := ',' || v_单据类型 || ',';
  v_单据类型 := Replace(v_单据类型, ',1,', ',24,');
  v_单据类型 := Replace(v_单据类型, ',2,', ',25,');
  v_单据类型 := Replace(v_单据类型, ',3,', ',26,');

  v_Jtmp := Null;
  For r_卫材 In (Select Decode(b.单据, 24, 1, 25, 2, 26, 3) As 单据类型, b.No, b.库房id
               From 未发药品记录 B
               Where Instr(v_单据类型, ',' || b.单据 || ',') > 0 And Nvl(b.已收费, 0) = n_收费标志 And
                     b.填制日期 Between Nvl(d_开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And Nvl(d_结束时间, Sysdate) And Exists
                (Select 1
                      From 药品收发记录
                      Where 单据 = b.单据 And NO = b.No And
                            (Instr(',' || v_费用来源 || ',', ',' || 费用来源 || ',') > 0 Or 费用来源 Is Null))) Loop
  
    v_Jtmp := v_Jtmp || ',{"billtype":' || r_卫材.单据类型;
    v_Jtmp := v_Jtmp || ',"stuff_no":"' || r_卫材.No || '"';
    v_Jtmp := v_Jtmp || ',"warehouse_id":' || Nvl(r_卫材.库房id, 0);
    v_Jtmp := v_Jtmp || '}';
  
    If Length(v_Jtmp) > 30000 Then
      If c_Jtmp Is Null Then
        c_Jtmp := Substr(v_Jtmp, 2);
      Else
        c_Jtmp := c_Jtmp || v_Jtmp;
      End If;
      v_Jtmp := Null;
    End If;
  End Loop;

  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Getnotsendrec;
/
CREATE OR REPLACE Procedure Zl_Stuffsvr_Odr_Check
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：超期发送收回卫材相关取数和检查
  --入参：Json_In:格式
  --  input
  --     item_list[]列表
  --               order_id                        N 1 医嘱ID
  --               stuff_nos                         C 1 单据号拼串
  --出参: Json_Out,格式如下
  --  output
  --    code                          N 1 应答吗：0-失败；1-成功
  --    message                       C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Input    Pljson;
  j_Item     Pljson;
  j_List     Pljson_List := Pljson_List();
  n_医嘱id   Number(18);
  n_Count    Number;
  v_Nos      Varchar2(30000);
  v_List     Varchar2(32767);
  v_Json_Out Varchar2(32767);
  v_Error    Varchar2(255);
  Err_Custom Exception;
Begin
  j_Item  := Pljson(Json_In);
  j_Input := j_Item.Get_Pljson('input');
  j_List  := j_Input.Get_Pljson_List('item_list');
  If j_List Is Not Null Then
    For I In 1 .. j_List.Count Loop
      j_Item   := Pljson();
      j_Item   := Pljson(j_List.Get(I));
      n_医嘱id := j_Item.Get_Number('order_id');
      v_Nos    := j_Item.Get_String('stuff_nos');
    
      For r_卫材 In (Select /*+cardinality(j,10)*/
                    a.No, a.费用id, a.药品id, Sum(Nvl(a.付数, 1) * Decode(a.审核人, Null, 0, a.实际数量)) As 已发数量
                   From 药品收发记录 A, Table(f_Str2list(v_Nos)) J
                   Where a.No = j.Column_Value And a.医嘱id = n_医嘱id
                   Group By a.No, a.费用id, a.药品id) Loop
      
        v_List := v_List || ',{"stuffdtl_id":' || r_卫材.费用id;
        v_List := v_List || ',"sended_num":' || Zljsonstr(r_卫材.已发数量, 1);
        v_List := v_List || ',"order_id":' || n_医嘱id;
        v_List := v_List || ',"stuff_id":' || r_卫材.药品id;
        v_List := v_List || '}';
      
      End Loop;
    
    End Loop;
  
    If v_List Is Not Null Then
      v_List := ',"item_list":[' || Substr(v_List, 2) || ']';
    End If;
  
    v_Json_Out := '{"code":1,"message":"成功"';
    v_Json_Out := v_Json_Out || v_List;
    v_Json_Out := v_Json_Out || '}';
    Json_Out   := '{"output":' || v_Json_Out || '}';
  
  End If;
Exception
  When Err_Custom Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(v_Error) || '"}}';
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Stuffsvr_Odr_Check;
/
Create Or Replace Procedure Zl_Stuffsvr_Overdue_Recovery
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：医嘱超期发送收回卫材相关处理
  --入参：Json_In:格式
  --  input
  --     operator_name                      C 1 操作员姓名
  --     operator_time                      C 1 操作时间
  --     item_list[]卫材删除列表
  --                  stuffdtl_id            N 1 处方明细id,目前传入的费用ID
  --                  return_num             N 1 销帐数量
  --     roll_list[]负数收回列表
  --                  clinic_type            C 1 医嘱诊疗类别
  --                  stuff_no               C 1 负数记帐的单据号
  --                  stuffdtl_id            N 1 卫材明细id,费用记录id
  --                  stuffdtl_id_old        N 1 原始卫材明细id,费用记录id
  --                  packages_num           N 1 付数
  --                  outbound_num           N 1 数量
  --                  is_stuff_order         N 1 区分是否是绑定的卫材费用0-非卫材医嘱,1-卫材医嘱

  --出参: Json_Out,格式如下
  --  output
  --    code                          N 1 应答吗：0-失败；1-成功
  --    message                       C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Input Pljson;
  j_Item  Pljson;
  j_List  Pljson_List := Pljson_List();

  n_卫材明细id 药品收发记录.Id%Type;
  n_退料数量   药品收发记录.填写数量%Type;
  v_人员姓名   Varchar2(3000);
  收回时间_In  Date;
  v_Dec        Number;
  -- v_划价类别   Varchar2(3000);
  v_收发序号    Number;
  No_In         Varchar2(3000);
  v_费用id      Number;
  Old_费用id    Number;
  v_当前付数    Number;
  v_当前数量    Number;
  v_诊疗类别    Varchar2(3000);
  n_卫材医嘱    Number;
  n_数量        Number;
  n_Count       Number;
  n_收费标志    Number;
  n_自动发料    Number;
  v_单据明细ids Varchar2(32767);

  Cursor c_Stuff Is
    Select b.批次, Nvl(x.在用分批, 0) As 分批, b.批号, b.效期, x.最大效期, b.Id As 收发id, b.病人id, b.主页id, b.库房id, b.单据, b.姓名, b.对方部门id,
           b.身份
    From 药品收发记录 B, 材料特性 X
    Where b.费用id = Old_费用id And b.药品id = x.材料id;

  Procedure 负数收发记录_Insert
  (
    费用id_In     Number,
    批次_In       药品收发记录.批次%Type,
    分批_In       药品规格.药房分批%Type,
    批号_In       药品收发记录.批号%Type,
    效期_In       药品收发记录.效期%Type,
    最大效期_In   药品规格.最大效期%Type,
    收发id_In     药品收发记录.Id%Type,
    病人id_In     药品收发记录.对方部门id%Type,
    主页id_In     药品收发记录.对方部门id%Type,
    库房id_In     药品收发记录.库房id%Type,
    单据_In       药品收发记录.单据%Type,
    姓名_In       Varchar2,
    对方部门id_In 药品收发记录.对方部门id%Type,
    
    P付数 药品收发记录.付数%Type,
    P数量 药品收发记录.填写数量%Type,
    P身份 药品收发记录.身份%Type
  ) Is
    v_批次   药品收发记录.批次%Type;
    v_效期   药品收发记录.效期%Type;
    v_批号   药品收发记录.批号%Type;
    v_Lngid  Number;
    v_优先级 身份.优先级%Type; --取身份优先级,需外部传入
  Begin
    --确定批次
    If Nvl(批次_In, 0) <> 0 And 分批_In = 0 Then
      --原分批,现不分批
      v_批次 := Null;
      v_批号 := 批号_In;
      v_效期 := 效期_In;
    Elsif Nvl(批次_In, 0) = 0 And 分批_In = 1 Then
      --原不分批,现分批
      Select 药品收发记录_Id.Nextval Into v_批次 From Dual;
      Select To_Char(Sysdate, 'YYYYMMDD') Into v_批号 From Dual;
      If 最大效期_In Is Not Null Then
        v_效期 := Trunc(Sysdate + 最大效期_In * 30);
      Else
        v_效期 := Null;
      End If;
    Else
      v_批次 := 批次_In;
      v_批号 := 批号_In;
      v_效期 := 效期_In;
    End If;
  
    Select 药品收发记录_Id.Nextval Into v_Lngid From Dual;
    Insert Into 药品收发记录
      (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 付数, 填写数量, 实际数量, 零售价, 零售金额, 摘要, 填制人, 填制日期,
       费用id, 单量, 频次, 用法, 供药单位id, 生产日期, 批准文号, 灭菌效期, 姓名, 性别, 年龄, 出生日期, 身份证号, 病人id, 主页id, 病人科室id, 病人病区id, 婴儿序号, 病人来源, 医嘱id,
       身份, 处方类型, 皮试结果, 诊断描述, 已收费, 费用来源)
      Select v_Lngid, 1, 单据, No_In, v_收发序号, 库房id, 对方部门id, 入出类别id, -1, 药品id, Nvl(v_批次, 0), 产地, v_批号, v_效期, P付数, -1 * P数量,
             -1 * P数量, 零售价, Round(-1 * P付数 * P数量 * 零售价, v_Dec), '超期发送收回', v_人员姓名, 收回时间_In, 费用id_In, 单量, 频次, 用法, 供药单位id,
             生产日期, 批准文号, 灭菌效期, 姓名, 性别, 年龄, 出生日期, 身份证号, 病人id, 主页id, 病人科室id, 病人病区id, 婴儿序号, 病人来源, 医嘱id, 身份, 处方类型, 皮试结果,
             诊断描述, n_收费标志, 费用来源
      From 药品收发记录
      Where ID = 收发id_In;
  
    Zl_未审药品记录_Insert(v_Lngid);
  
    Zl_药品库存_Update(v_Lngid, 0, 1);
  
    --未发药品记录
    Update 未发药品记录
    Set 病人id = 病人id_In, 主页id = 主页id_In, 姓名 = 姓名_In
    Where 单据 = 单据_In And NO = No_In And 库房id + 0 = 库房id_In;
    If P身份 Is Not Null Then
      Select Max(b.优先级) Into v_优先级 From 身份 B Where b.名称 = P身份;
    End If;
    If Sql%RowCount = 0 Then
      Insert Into 未发药品记录
        (单据, NO, 病人id, 主页id, 姓名, 优先级, 对方部门id, 库房id, 填制日期, 已收费, 打印状态)
      Values
        (单据_In, No_In, 病人id_In, 主页id_In, 姓名_In, v_优先级, 对方部门id_In, 库房id_In, 收回时间_In, n_收费标志, 0);
    End If;
  End;

Begin
  j_Item  := Pljson(Json_In);
  j_Input := j_Item.Get_Pljson('input');

  --药品删除列表
  j_List := j_Input.Get_Pljson_List('item_list');
  If j_List Is Not Null Then
    For I In 1 .. j_List.Count Loop
      j_Item       := Pljson();
      j_Item       := Pljson(j_List.Get(I));
      n_卫材明细id := j_Item.Get_Number('stuffdtl_id');
      n_退料数量   := j_Item.Get_Number('return_num');
      Zl_卫材收发记录_销售退费_s(n_卫材明细id, n_退料数量, 1);
    End Loop;
  End If;

  j_List := Pljson_List();
  j_List := j_Input.Get_Pljson_List('roll_list');

  If j_List Is Not Null Then
  
    v_人员姓名  := j_Input.Get_String('operator_name');
    收回时间_In := To_Date(j_Input.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');
  
    --金额小数位数
    Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into v_Dec From Dual;
  
    For I In 1 .. j_List.Count Loop
      j_Item := Pljson();
      j_Item := Pljson(j_List.Get(I));
    
      Select Nvl(Max(序号), 0) + 1 Into v_收发序号 From 药品收发记录 Where 单据 = 25 And 记录状态 = 1 And NO = No_In;
    
      v_诊疗类别 := j_Item.Get_String('clinic_type');
      No_In      := j_Item.Get_String('stuff_no');
      v_费用id   := j_Item.Get_Number('stuffdtl_id');
      Old_费用id := j_Item.Get_Number('stuffdtl_id_old');
      v_当前付数 := j_Item.Get_Number('packages_num');
      v_当前数量 := j_Item.Get_Number('outbound_num');
      n_卫材医嘱 := j_Item.Get_Number('is_stuff_order'); --区分是否是绑定的药品费用0-非卫材医嘱,1-卫材医嘱
      n_收费标志 := j_Item.Get_Number('charge_tag');
      n_自动发料 := j_Item.Get_Number('stuff_auto_send');
      If n_自动发料 = 1 Then
        v_单据明细ids := v_单据明细ids || ',' || v_费用id;
      End If;
      For r_Stuff In c_Stuff Loop
        If n_卫材医嘱 = 1 Then
          负数收发记录_Insert(v_费用id, r_Stuff.批次, r_Stuff.分批, r_Stuff.批号, r_Stuff.效期, r_Stuff.最大效期, r_Stuff.收发id,
                        r_Stuff.病人id, r_Stuff.主页id, r_Stuff.库房id, r_Stuff.单据, r_Stuff.姓名, r_Stuff.对方部门id, v_当前付数, v_当前数量,
                        r_Stuff.身份);
        Else
          n_数量 := v_当前数量;
          For r_Otherstuff In (Select b.批次, Nvl(x.在用分批, 0) As 分批, Nvl(b.付数, 1) * b.实际数量 As 数量, b.批号, b.效期, x.最大效期,
                                      b.Id As 收发id, b.病人id, b.主页id, b.库房id, b.单据, b.姓名, b.对方部门id, b.身份
                               From 药品收发记录 B, 材料特性 X
                               Where b.费用id = Old_费用id And (b.记录状态 = 1 Or Mod(b.记录状态, 3) = 0) And b.药品id = x.材料id
                               Order By b.Id Desc) Loop
            If n_数量 > 0 Then
              n_Count := r_Otherstuff.数量;
              If n_数量 < n_Count Then
                n_Count := n_数量;
              End If;
              负数收发记录_Insert(v_费用id, r_Otherstuff.批次, r_Otherstuff.分批, r_Otherstuff.批号, r_Otherstuff.效期,
                            r_Otherstuff.最大效期, r_Otherstuff.收发id, r_Otherstuff.病人id, r_Otherstuff.主页id,
                            r_Otherstuff.库房id, r_Otherstuff.单据, r_Otherstuff.姓名, r_Otherstuff.对方部门id, 1, n_Count,
                            r_Otherstuff.身份);
              n_数量 := n_数量 - r_Otherstuff.数量;
            End If;
          End Loop;
        End If;
      End Loop;
    End Loop;
  End If;

  If v_单据明细ids Is Not Null Then
    v_单据明细ids := Substr(v_单据明细ids, 2);
    Zl_卫材收发记录_自动发料_s(Null, v_人员姓名, Null, v_单据明细ids, 1);
  End If;

  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Stuffsvr_Overdue_Recovery;
/