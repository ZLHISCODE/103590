Create Or Replace Procedure Zl_Drugsvr_Getstockshow
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：获取指定库房的库存数据，用于显示
  --入参：Json_In:格式
  --  input
  --    pharmacy_ids        C   1   库房ID
  --出参: Json_Out,格式如下
  --  output
  --    code                N 1 应答吗：0-失败；1-成功
  --    message             C 1 应答消息：失败时返回具体的错误信息
  --    item_list
  --      drug_id              N   1   药品ID
  --      pharmacy_id          N   1   库房ID
  --      stock                N   1   可用数量
  --      real_stock          N  1 实际库存
  --      avg_price           N  1 平均售价
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
  v_库房ids := j_Json.Get_String('pharmacy_ids');

  If Nvl(v_库房ids, 0) = 0 Then
    Json_Out := Zljsonout('未传入相关库房信息');
    Return;
  End If;

  For c_库存 In (Select a.库房id, a.药品id, Nvl(Sum(a.可用数量), 0) As 可用数量, Nvl(Sum(a.实际数量), 0) As 实际数量,
                      Decode(Nvl(Sum(a.实际数量), 0), 0, Max(a.零售价), Nvl(Sum(a.实际金额), 0) / Nvl(Sum(a.实际数量), 0)) As 平均售价
               From 药品库存 A
               Where a.性质 = 1 And Instr(',' || v_库房ids || ',', ',' || a.库房id || ',') > 0
               Group By a.库房id, a.药品id) Loop
  
    v_Jtmp := v_Jtmp || ',';
    Zljsonputvalue(v_Jtmp, 'drug_id', c_库存.药品id, 1, 1);
    Zljsonputvalue(v_Jtmp, 'pharmacy_id', c_库存.库房id, 1);
    Zljsonputvalue(v_Jtmp, 'stock', c_库存.可用数量, 1);
    Zljsonputvalue(v_Jtmp, 'real_stock', c_库存.实际数量, 1);
    Zljsonputvalue(v_Jtmp, 'avg_price', c_库存.平均售价, 1, 2);
  
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
End Zl_Drugsvr_Getstockshow;
/

Create Or Replace Procedure Zl_Drugsvr_CheckDrugExistStock
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  ---------------------------------------------------------------------------
  --input      判断药品是否存在库数据
  --  drug_id      N  1  药品id
  --  is_item      N  1  是否按品种查询：0-按规格查询，1-按品种查询
  --output      
  --  code  C  1  应答码：0-失败；1-成功
  --  message  C  1  应答消息：
  --  isexist  N 1 是否存在: 1-存在;0-不存在
  ---------------------------------------------------------------------------
  j_Jsonin Pljson;
  j_Json   Pljson;
  n_药品id 药品规格.药品id%Type;
  n_品种   Number(1);
  n_Exist  Number(1);
Begin
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_药品id := j_Json.Get_Number('drug_id');
  n_品种   := j_Json.Get_Number('is_item');

  If n_品种 = 0 Then
    Select Count(1) Into n_Exist From 药品库存 Where 药品id = n_药品id And Rownum < 2;
  Else
    Select Count(1)
    Into n_Exist
    From 药品规格 A, 药品库存 B
    Where a.药品id = b.药品id And a.药名id = n_药品id And Rownum < 2;
  End If;

  Json_Out := '{"output":{"code": 1,"message":"成功","isexist":' || n_Exist || '}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugsvr_CheckDrugExistStock;
/

Create Or Replace Procedure Zl_DrugSvr_GetCostPriceAdjust
(
  Json_In  Varchar2,
  Json_Out Out Clob
) Is
  ---------------------------------------------------------------------------
  --功能：获取药品成本价调价记录
  --input      
  --  drug_id      N   1 药品id
  --  show_unit    N   1   显示单位:0-售价单位;3-药库单位
  --output      
  --  code  C  1  应答码：0-失败；1-成功
  --  message  C  1  应答消息：  
  --  price_list[]  药品成本价调价记录
  --     drug_id   N 1  药品ID
  --     drug_name   C 1  药品信息
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
  --     drug_revoke_time  C 1 撤档时间
  --     node_no      C    0  站点编码   
  --     is_stock    N   1 是否有库存数据  0-否，1-是
  ---------------------------------------------------------------------------
  j_Jsonin Pljson;
  j_Json   Pljson;
  n_药品id 药品规格.药品id%Type;
  n_单位   Number(1);

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;
Begin
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_药品id := j_Json.Get_Number('drug_id');
  n_单位   := j_Json.Get_Number('show_unit');

  v_Jtmp := Null;
  For r_Costprice In (Select Distinct b.No, i.Id As 药品id, '[' || i.编码 || ']' || i.名称 || ' ' || i.规格 || ' ' || i.产地 As 药品,
                                      p.名称 As 库房, a.批号, a.效期, a.产地, Decode(n_单位, 0, i.计算单位, s.药库单位) As 单位,
                                      Decode(n_单位, 0, a.原价, a.原价 * Nvl(s.药库包装, 1)) As 原成本价,
                                      Decode(n_单位, 0, a.现价, a.现价 * Nvl(s.药库包装, 1)) As 成本价, a.执行日期, a.调价说明, i.撤档时间, i.站点,
                                      Decode(k.库房id, Null, 0, 1) As 库存
                      From 药品收发记录 B, 收费项目目录 I, 药品规格 S, 部门表 P, 药品价格记录 A, 药品库存 K
                      Where a.价格类型 = 2 And a.收发id = b.Id(+) And a.药品id = i.Id And i.Id = s.药品id And a.库房id = p.Id(+) And
                            s.药名id = n_药品id And k.性质(+) = 1 And k.库房id(+) = a.库房id And k.药品id(+) = a.药品id And
                            k.批次(+) = a.批次
                      Order By '[' || i.编码 || ']' || i.名称 || ' ' || i.规格 || ' ' || i.产地, p.名称, a.批号, a.执行日期 Desc, NO Desc) Loop
  
    v_Jtmp := v_Jtmp || ',';
    Zljsonputvalue(v_Jtmp, 'drug_id', r_Costprice.药品id, 1, 1);
    Zljsonputvalue(v_Jtmp, 'drug_name', r_Costprice.药品, 0);
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
    Zljsonputvalue(v_Jtmp, 'drug_revoke_time', r_Costprice.撤档时间, 0);
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
End Zl_DrugSvr_GetCostPriceAdjust;
/

Create Or Replace Procedure Zl_DrugSvr_AdjustPriceType
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --执行调价
  ---------------------------------------------------------------------------
  --input      药品价格属性调整时产生的调价盈亏和库存变化数据处理
  --    drug_list[]
  --       drug_id      N    药品id
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

  n_药品id     药品规格.药品id%Type;
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
  n_药品价格记录 Number(1);
  v_类别         收费项目目录.类别%Type;

  --定价->时价后更新药品价格记录的值
  Cursor c_Priceadjust Is
    Select s.药品id, s.库房id, Nvl(s.批次, 0) As 批次, s.上次供应商id As 供应商id, s.上次批号 As 批号, s.效期, s.上次产地 As 产地,
           Nvl(s.实际数量, 0) As 实际数量, s.上次扣率 As 扣率, Nvl(s.实际金额, 0) As 实际金额, Nvl(s.实际差价, 0) As 实际差价, Nvl(s.零售价, 0) As 零售价,
           s.平均成本价, s.灭菌效期, s.批准文号, s.上次生产日期 As 生产日期
    From 药品库存 S
    Where s.药品id = n_药品id And s.性质 = 1
    Order By s.药品id, s.批次, s.库房id;

  r_Priceadjust c_Priceadjust%RowType;
Begin
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  j_Jsonlist := j_Json.Get_Pljson_List('drug_list');

  n_Count := j_Jsonlist.Count;
  If j_Jsonlist.Count = 0 Then
    Json_Out := Zljsonout('未传入药品信息！');
    Return;
  End If;

  For I In 1 .. j_Jsonlist.Count Loop
    j_Json := Pljson();
    j_Json := Pljson(j_Jsonlist.Get(I));
  
    n_药品id     := j_Json.Get_Number('drug_id');
    n_原价格类型 := j_Json.Get_Number('price_type_old');
    n_新价格类型 := j_Json.Get_Number('price_type_new');
  
    If n_原价格类型 <> n_新价格类型 Then
      --取原价和原价id(调用该过程前已经产生了新价格)
      Begin
        Select 原价, 现价, 原价id As 价格id
        Into n_收费价目原价, n_收费价目现价, n_价格id
        From 收费价目
        Where 收费细目id = n_药品id And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And 变动原因 = 1;
      Exception
        When Others Then
          n_收费价目原价 := Null;
          n_收费价目现价 := Null;
          n_价格id       := Null;
      End;
    
      --时价->定价
      If n_原价格类型 = 1 And n_新价格类型 = 0 Then
        Select 精度 Into n_流通金额小数 From 药品卫材精度 Where 类别 = 1 And 内容 = 4 And 单位 = 5;
      
        --取入出类别ID
        Select 类别id Into n_入出类别id From 药品单据性质 Where 单据 = 13;
      
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
               r_Priceadjust.扣率, n_零售金额, n_零售金额, '时价转定价', zl_UserName, Sysdate, r_Priceadjust.库房id, 1, n_价格id,
               zl_UserName, Sysdate, r_Priceadjust.实际金额, r_Priceadjust.实际差价, r_Priceadjust.供应商id);
          
            Zl_药品库存_Update(n_收发id, 2, 0);
          End If;
        End Loop;
      
        --定价->时价
      Elsif n_原价格类型 = 0 And n_新价格类型 = 1 Then
        For r_Priceadjust In c_Priceadjust Loop
          n_药品价格记录 := 0;
          Begin
            Select 1, 现价
            Into n_药品价格记录, n_原价
            From 药品价格记录
            Where 药品id = r_Priceadjust.药品id And 库房id = r_Priceadjust.库房id And Nvl(批次, 0) = r_Priceadjust.批次 And
                  记录状态 = 1 And 价格类型 = 1;
          Exception
            When Others Then
              n_药品价格记录 := 0;
              n_原价         := n_收费价目原价;
          End;
        
          If n_药品价格记录 = 1 Then
            Zl_药品价格记录_Stop(1, r_Priceadjust.库房id, r_Priceadjust.药品id, r_Priceadjust.批次, Sysdate - 1 / 24 / 60 / 60, 2);
          End If;
          Zl_药品价格记录_Insert(0, 1, r_Priceadjust.库房id, r_Priceadjust.药品id, r_Priceadjust.批次, n_原价, n_收费价目现价, Sysdate,
                           '定价转时价', zl_UserName, Null, r_Priceadjust.供应商id, r_Priceadjust.批号, r_Priceadjust.效期,
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
End Zl_DrugSvr_AdjustPriceType;
/


Create Or Replace Procedure Zl_DrugSvr_CheckDrugExistRec
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --执行调价
  ---------------------------------------------------------------------------
  --input      判断药品是否存在收发记录
  --  drug_id      N    药品id
  --output      
  --  code  C  1  应答码：0-失败；1-成功
  --  message  C  1  应答消息：
  --  isexist  N 1 是否存在: 1-存在;0-不存在
  ---------------------------------------------------------------------------
  j_Jsonin Pljson;
  j_Json   Pljson;
  n_药品id 药品规格.药品id%Type;
  n_Exist  Number(1);
Begin
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_药品id := j_Json.Get_Number('drug_id');

  Select Count(1) Into n_Exist From 药品收发记录 Where 药品id = n_药品id And Rownum < 2;

  Json_Out := '{"output":{"code": 1,"message":"成功","isexist":' || n_Exist || '}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_DrugSvr_CheckDrugExistRec;
/


Create Or Replace Procedure Zl_Drugsvr_Checkpriceadjust
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --零差价管理模式时，判断价格是否满足零差价管理要（成本价和售价一致）
  ---------------------------------------------------------------------------
  --input      检查药品零差价控制
  --  pharmacy_drug_ids     C    库房药品id串：库房id,药品id;...
  --  is_ignore    N    忽略零差价管理 0-不忽略，1-忽略
  --output      
  --  code  C  1  应答码：0-失败；1-成功
  --  message  C  1  应答消息：
  --  price_list[]  药品成本价调价记录
  --     drug_id   N 1  药品ID
  --     drug_name   C 1  药品信息
  --     pharmacy_id   id 1  库房id
  --     pharmacy_name   C 1  库房名称
  --     isstock   N 1 提示类型：0-没有库存，1-有库存
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  n_药品id         药品库存.药品id%Type;
  n_库房id         药品库存.库房id%Type;
  n_零差价管理     Number(1);
  n_成本价         药品规格.成本价%Type;
  n_售价           药品规格.上次售价%Type;
  v_通用名         Varchar2(32767);
  n_库存           Number(18);
  n_忽略零差价管理 Number(1);
  v_药品串         Varchar2(4000);
  v_Fields         Varchar2(4000);

  v_Jtmp       Varchar2(32767);
  n_Checkvalue Number(1);
Begin
  Select zl_To_Number(Nvl(zl_GetSysParameter(275), '0')) Into n_零差价管理 From Dual;

  If n_零差价管理 = 0 Then
    --没启用零差价管理时退出
    Json_Out := zlJsonOut('成功', 1);
    Return;
  End If;

  j_Jsonin         := PLJson(Json_In);
  j_Json           := j_Jsonin.Get_Pljson('input');
  v_药品串         := j_Json.Get_String('pharmacy_drug_ids');
  n_忽略零差价管理 := Nvl(j_Json.Get_Number('is_ignore'), 0);

  If v_药品串 Is Null Then
    Json_Out := zlJsonOut('未传入药品信息，请检查!');
    Return;
  End If;

  v_药品串 := v_药品串 || ';';

  While v_药品串 Is Not Null Loop
    v_Fields := Substr(v_药品串, 1, Instr(v_药品串, ';') - 1);
    n_库房id := To_Number(Substr(v_Fields, 1, Instr(v_Fields, ',') - 1));
    n_药品id := To_Number(Substr(v_Fields, Instr(v_Fields, ',') + 1));
    v_药品串 := Replace(';' || v_药品串, ';' || v_Fields || ';');
  
    If n_药品id = 0 Then
      Json_Out := zlJsonOut('未传入药品ID信息，请检查');
      Return;
    End If;
  
    n_成本价 := Null;
    n_售价   := Null;
    v_通用名 := Null;
  
    Select a.成本价, b.现价 As 售价, '[' || c.编码 || ']' || c.名称 || Decode(c.产地, Null, Null, '(' || c.产地 || ')') || c.规格 As 通用名,
           Nvl(a.是否零差价管理, 0) As 零差价管理
    Into n_成本价, n_售价, v_通用名, n_零差价管理
    From 药品规格 A, 收费价目 B, 收费项目目录 C
    Where a.药品id = b.收费细目id And a.药品id = c.Id And (Sysdate Between b.执行日期 And b.终止日期) And a.药品id = n_药品id;
  
    --检查有无库存
    If n_库房id > 0 Then
      Select Count(*)
      Into n_库存
      From 药品库存
      Where 性质 = 1 And 药品id = n_药品id And 库房id = n_库房id And
            Not (批次 = 0 And 可用数量 < 0 And 实际数量 = 0 And 实际金额 = 0 And 实际差价 = 0);
    Else
      Select Count(*)
      Into n_库存
      From 药品库存
      Where 性质 = 1 And 药品id = n_药品id And Not (批次 = 0 And 可用数量 < 0 And 实际数量 = 0 And 实际金额 = 0 And 实际差价 = 0);
    End If;
  
    If n_库存 = 0 Then
      --无库存时，从收费价目取售价，从药品规格取成本价，并比较价格
      If n_零差价管理 = 0 And n_忽略零差价管理 = 0 Then
        --不启用零差价管理
        n_Checkvalue := 0;
      Else
        If n_成本价 = n_售价 Then
          --售价和成本价一致时
          n_Checkvalue := 0;
        Else
          --售价和成本价不一致时
          n_Checkvalue := 1;
        End If;
      End If;
    
      If n_Checkvalue = 1 Then
        v_Jtmp := v_Jtmp || ',';
        zlJsonPutValue(v_Jtmp, 'drug_id', n_药品id, 1, 1);
        zlJsonPutValue(v_Jtmp, 'drug_name', v_通用名, 0);
        zlJsonPutValue(v_Jtmp, 'pharmacy_id', n_库房id, 1);
        zlJsonPutValue(v_Jtmp, 'pharmacy_name', '', 0);
        zlJsonPutValue(v_Jtmp, 'isstock', 0, 1, 2);
      End If;
    Else
      --有库存数据时
      n_Checkvalue := 0;
      For r_价格 In (Select 药品id, 通用名, 规格, 库房id, 库房, 生产商, '' As 批号, 批次, 单位, 药库包装, 售价, Sum(成本价 * 实际数量) / Sum(实际数量) As 成本价,
                          是否时价
                   From (Select a.药品id,
                                 '[' || c.编码 || ']' || c.名称 || Decode(c.产地, Null, Null, '(' || c.产地 || ')') || c.规格 As 通用名,
                                 c.规格, c.产地 As 生产商, Null As 批次, a.药库单位 As 单位, a.药库包装, b.现价 As 售价, d.平均成本价 As 成本价, 0 As 是否时价,
                                 d.实际数量, d.库房id, e.名称 As 库房
                          From 药品规格 A, 收费价目 B, 收费项目目录 C, 药品库存 D, 部门表 E
                          Where a.药品id = b.收费细目id And a.药品id = c.Id And a.药品id = d.药品id And d.性质 = 1 And
                                (Sysdate Between b.执行日期 And b.终止日期) And
                                (c.撤档时间 Is Null Or c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And c.是否变价 = 0 And
                                Nvl(a.是否零差价管理, 0) = 1 And b.现价 <> d.平均成本价 And
                                d.库房id In (Select Distinct 部门id From 部门性质说明 Where 工作性质 Like '%药房') And
                                a.药品id = Decode(n_药品id, 0, a.药品id, n_药品id) And d.库房id = Decode(n_库房id, 0, d.库房id, n_库房id) And
                                Not (d.批次 = 0 And d.可用数量 < 0 And d.实际数量 = 0 And d.实际金额 = 0 And d.实际差价 = 0) And
                                d.库房id = e.Id)
                   Group By 药品id, 通用名, 规格, 库房id, 库房, 生产商, 批次, 单位, 药库包装, 售价, 是否时价
                   Having Sum(实际数量) <> 0
                   Union All
                   Select a.药品id,
                          '[' || c.编码 || ']' || c.名称 || Decode(c.产地, Null, Null, '(' || c.产地 || ')') || c.规格 As 通用名, c.规格,
                          d.库房id, e.名称 As 库房, d.上次产地 As 生产商, d.上次批号 As 批号, d.批次, a.药库单位 As 单位, a.药库包装,
                          Nvl(d.零售价, 0) As 售价, d.平均成本价 As 成本价, 1 As 是否时价
                   From 药品规格 A, 收费项目目录 C, 药品库存 D, 部门表 E
                   Where a.药品id = c.Id And a.药品id = d.药品id And d.库房id = e.Id And d.性质 = 1 And c.是否变价 = 1 And
                         (c.撤档时间 Is Null Or c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And Nvl(a.是否零差价管理, 0) = 1 And
                         Nvl(d.零售价, 0) <> d.平均成本价 And
                         d.库房id In (Select Distinct 部门id From 部门性质说明 Where 工作性质 Like '%药房') And
                         a.药品id = Decode(n_药品id, 0, a.药品id, n_药品id) And d.库房id = Decode(n_库房id, 0, d.库房id, n_库房id) And
                         Not (d.批次 = 0 And d.可用数量 < 0 And d.实际数量 = 0 And d.实际金额 = 0 And d.实际差价 = 0)
                   Order By 通用名, 库房id, 批号) Loop
      
        --找到数据时
        n_Checkvalue := 1;
        v_Jtmp       := v_Jtmp || ',';
        zlJsonPutValue(v_Jtmp, 'drug_id', r_价格.药品id, 1, 1);
        zlJsonPutValue(v_Jtmp, 'drug_name', r_价格.通用名, 0);
        zlJsonPutValue(v_Jtmp, 'pharmacy_id', r_价格.库房id, 1);
        zlJsonPutValue(v_Jtmp, 'pharmacy_name', r_价格.库房, 0);
        zlJsonPutValue(v_Jtmp, 'isstock', 1, 1, 2);
      End Loop;
    End If;
  End Loop;

  If v_Jtmp Is Not Null Then
    Json_Out := '{"output":{"code":1,"message":"成功","price_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    Json_Out := zlJsonOut('成功', 1);
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Checkpriceadjust;
/




Create Or Replace Procedure Zl_Drugsvr_Executeprice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --执行调价
  ---------------------------------------------------------------------------
  --input      检查药品售价，成本价是否存在已生效但未执行的价格，如果存在则执行调价
  --  drug_id      N    药品id
  --output      
  --  code  C  1  应答码：0-失败；1-成功
  --  message  C  1  应答消息：
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  n_药品id 药品规格.药品id%Type;

  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin
  j_Jsonin := PLJson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_药品id := j_Json.Get_Number('drug_id');

  If n_药品id = 0 Then
    v_Err_Msg := '未传入药品ID信息！';
    Raise Err_Item;
  End If;

  For r_调价 In (Select Distinct b.药品id As 药品id
               From 收费项目目录 I, 收费价目 N, 药品规格 B
               Where i.Id = n.收费细目id And i.Id = b.药品id And
                     (i.撤档时间 Is Null Or i.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) And n.变动原因 = 0 And
                     Sysdate > n.执行日期 And n.价格等级 Is Null And b.药品id = n_药品id
               Union
               Select Distinct a.药品id
               From 药品价格记录 A, 药品规格 B
               Where a.药品id = b.药品id And a.记录状态 = 0 And a.执行日期 <= Sysdate And b.药品id = n_药品id) Loop
  
    n_药品id := r_调价.药品id;
    Exit;
  End Loop;

  If n_药品id > 0 Then
    Zl_药品收发记录_Adjust(n_药品id);
  End If;

  Json_Out := '{"output":{"code": 1,"message":"成功"}}';
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugsvr_Executeprice;
/
Create Or Replace Procedure Zl_Drugsvr_Patiinfoupdate
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：住院病人药品收发记录基本信息修改
  --入参：Json_In:格式
  --  input
  --   pati_name      C   1   患者姓名
  --   pati_sex       C   1   患者性别
  --   pati_age       C   1   患者年龄
  --   visit_id       N   1   就诊id
  --   pati_id        N   1   病人id

  --出参: Json_Out,格式如下
  --  output
  --    code                 N   1 应答吗：0-失败；1-成功
  --    message              C   1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  v_姓名   Varchar2(100);
  v_性别   Varchar2(100);
  v_年龄   Varchar2(100);
  n_病人id Number;
  n_就诊id Number;
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
End Zl_Drugsvr_Patiinfoupdate;
/

Create Or Replace Procedure Zl_Drugsvr_Getretained
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  --------------------------------------------------------------------
  --功能：住院医嘱医嘱回退发送时药品医嘱留存数量检查
  --入参：JOSN格式
  --input
  --     order_ids                     C 1 医嘱ID拼串
  --出参：JSON
  --output
  --     code                          N 1 应答码：0-失败；1-成功
  --     message                       C 1 应答消息：失败时返回具体的错误信息
  --     data[]列表
  --         warehouse                 C 1 库房名称
  --         drug_name                 C 1 药品名称，诊疗项目目录.名称
  --         inp_unit                  C 1 住院单位
  --         re_quantity               N 1 回退数量
  --         quantity                  N 1 留存数量
  --------------------------------------------------------------------
  l_Vals    t_Strlist;
  P         Number;
  c_医嘱ids Clob;

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;

  j_Jsonin Pljson;
  j_Json   Pljson;
Begin
  --解析入参
  j_Jsonin  := Pljson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  c_医嘱ids := j_Json.Get_Clob('order_ids');

  l_Vals    := t_Strlist();
  c_医嘱ids := c_医嘱ids || ',';
  Loop
    P := Instr(c_医嘱ids, ',');
    Exit When(Nvl(P, 0) = 0);
    l_Vals.Extend;
    l_Vals(l_Vals.Count) := (Substr(c_医嘱ids, 1, P - 1));
    c_医嘱ids := Substr(c_医嘱ids, P + 1);
  End Loop;

  v_Jtmp := Null;
  For r_Data In (Select d.名称 As 库房, e.名称 As 药品, (a.回退 / Nvl(b.住院包装, 1)) As 回退数量, (a.留存 / Nvl(b.住院包装, 1)) As 留存数量, b.住院单位
                 From (Select a.库房id, a.药品id, a.回退数量 As 回退, b.留存数量 As 留存
                        From (Select a.库房id, a.病人病区id As 病区id, a.药品id, Sum(a.实际数量) As 回退数量
                               From 药品收发记录 A
                               Where a.医嘱id In (Select /*+cardinality(f,10)*/
                                                 To_Number(f.Column_Value) As 医嘱id
                                                From Table(l_Vals) F)
                               Group By a.库房id, a.病人病区id, a.药品id) A, 药品留存计划 B
                        Where a.病区id = b.部门id(+) And a.库房id = b.库房id(+) And a.药品id = b.药品id(+) And b.状态(+) = 0) A, 药品规格 B,
                      部门表 D, 诊疗项目目录 E
                 Where a.库房id = d.Id And a.药品id = b.药品id And b.药名id = e.Id And a.留存 <> 0 And a.回退 > a.留存) Loop
  
    v_Jtmp := v_Jtmp || ',';
    Zljsonputvalue(v_Jtmp, 'warehouse', r_Data.库房, 0, 1);
    Zljsonputvalue(v_Jtmp, 'drug_name', r_Data.药品, 0);
    Zljsonputvalue(v_Jtmp, 'inp_unit', r_Data.住院单位, 0);
    Zljsonputvalue(v_Jtmp, 're_quantity', r_Data.回退数量, 1);
    Zljsonputvalue(v_Jtmp, 'quantity', r_Data.留存数量, 1, 2);
  
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
    Json_Out := '{"output":{"code":1,"message":"成功","data":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","data":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getretained;
/

Create Or Replace Procedure Zl_Drugsvr_Newrecbill_Check
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
  --    drug_id   N 1 药品id
  --    send_num  N 1 发药数量
  --    pharmacy_id N 1 药房id
  --    price           N       1       售价
  --出参      json
  --output      
  --  code  C 1 应答码：0-失败；1-成功
  --  message C 1 应答消息：
  -------------------------------------------------------------------------------------------------
Begin

  Zl_药品销售出库_Check(Json_In, Json_Out);

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Newrecbill_Check;
/
Create Or Replace Procedure Zl_Drugsvr_Skintest_Update
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  ---------------------------------------------------------------------------
  --功能：更新处方病人皮试结果
  ---------------------------------------------------------------------------
  --入参：Json_In:格式
  --input  
  --   order_ids              C 1 医嘱ids，逗号拼串
  --   skintest_info          C 1 (+)阴性,(-)阳性,空
  --出参: Json_Out,格式如下
  --output
  --    code                   C 1 应答码：0-失败；1-成功
  --    message                C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  v_皮试结果  Varchar2(60);
  j_Jsonin    Pljson;
  j_Json      Pljson;
  v_Order_Ids Varchar2(32767);
Begin
  --解析入参
  j_Jsonin    := Pljson(Json_In);
  j_Json      := j_Jsonin.Get_Pljson('input');
  v_Order_Ids := j_Json.Get_String('order_ids');
  v_皮试结果  := j_Json.Get_String('skintest_info');

  Update 药品收发记录
  Set 皮试结果 = v_皮试结果
  Where 单据 In (8, 9) And 审核日期 Is Not Null And
        医嘱id In (Select /*+cardinality(x,10)*/
                  x.Column_Value
                 From Table(Cast(f_Num2list(v_Order_Ids) As Zltools.t_Numlist)) X);

  Json_Out := '{"output":{"code": 1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugsvr_Skintest_Update;
/

Create Or Replace Procedure Zl_Drugsvr_Patibase_Upate
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --更新处方中的病人基本信息
  ---------------------------------------------------------------------------
  --入参      json
  --input     更新处方中的病人基本信息
  --  pati_id N 1 病人id
  --  register_id N   挂号id
  --  pati_pageid N   主页id
  --  pati_name C   姓名
  --  pati_sex  C   性别
  --  pati_age  C   年龄
  --  pati_birthdate  D   出生日期
  --  pati_Idcard C   身份证号
  --出参      json
  --output      
  --code  C 1 应答码：0-失败；1-成功
  --message C 1 应答消息：
  ---------------------------------------------------------------------------
  n_病人id   Number(18);
  v_姓名     Varchar2(100);
  v_性别     Varchar2(4);
  v_年龄     Varchar2(20);
  d_出生日期 Date;
  v_身份证号 Varchar2(18);

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --解析入参
  j_Jsonin   := PLJson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_病人id   := j_Json.Get_Number('pati_id');
  v_姓名     := j_Json.Get_String('pati_name');
  v_性别     := j_Json.Get_String('pati_sex');
  v_年龄     := j_Json.Get_String('pati_age');
  d_出生日期 := To_Date(j_Json.Get_String('pati_birthdate'), 'yyyy-mm-dd');
  v_身份证号 := j_Json.Get_String('pati_Idcard');

  --修改未发药品记录
  Update 未发药品记录 Set 姓名 = v_姓名 Where 病人id = n_病人id;

  --修改未审核的药品收发记录
  Update 药品收发记录
  Set 姓名 = v_姓名, 性别 = v_性别, 年龄 = v_年龄, 出生日期 = d_出生日期, 身份证号 = v_身份证号
  Where 审核日期 Is Null And 病人id = n_病人id;

  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugsvr_Patibase_Upate;
/
Create Or Replace Procedure Zl_Drugsvr_Autosenddrug
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：收费或记帐后，自动发药（按NO或处方明细）
  --入参：Json_In:格式
  --  input
  --    billtype             N 1 单据类型:1 -收费处方发药  ;2- 记帐单处方发药;3- 记帐表处方发药;
  --    operator_code        C 1 操作员编号
  --    operator_name        C 1 操作员姓名
  --    rcp_nos              C 1 处方NO串：NO1,NO2...
  --    rcpdtl_ids           C 1 处方明细id串,目前传入的费用ID串，用逗号分隔 ：1,2,3,4
  --    send_type            N 0 发药类型,0-按 传入的no单据类型发药,1-只按 处理方明细id串发药
  --出参: Json_Out,格式如下
  --  output
  --    code                 N 1   应答吗：0-失败；1-成功
  --    message              C 1   应答消息：失败时返回具体的错误信息
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
      n_单据 := 8;
    Elsif n_单据 = 2 Then
      n_单据 := 9;
    Elsif n_单据 = 3 Then
      n_单据 := 10;
    Else
      v_Err := '传入节点【billtype】错误，请检查！';
      Raise Err_Custom;
    End If;
  
    v_Nos := j_Json.Get_String('rcp_nos');
    v_Ids := j_Json.Get_String('rcpdtl_ids');
  
    If v_Ids Is Null And v_Nos Is Null Then
      v_Err := '未传入药品单据【rcp_nos】节点或明细信息【rcpdtl_ids】节点';
      Raise Err_Custom;
      Return;
    End If;
  End If;

  If Nvl(n_Send_Type, 0) = 1 Then
    v_Ids := j_Json.Get_String('rcpdtl_ids');
    If v_Ids Is Null Then
      v_Err := '未传入药品明细信息【rcpdtl_ids】节点';
      Raise Err_Custom;
      Return;
    End If;
  End If;

  v_操作员姓名 := j_Json.Get_String('operator_name');
  v_操作员编号 := j_Json.Get_String('operator_code');

  --按处方号发药
  If v_Nos Is Not Null Then
    Zl_药品收发记录_自动发药_s(n_单据, v_操作员姓名, v_操作员编号, v_Nos, 0);
  End If;

  --按单据ID发药
  If v_Ids Is Not Null Then
    Zl_药品收发记录_自动发药_s(n_单据, v_操作员姓名, v_操作员编号, v_Ids, 1);
  End If;

  Json_Out := '{"output":{"code":1,"message": "成功"}}';
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugsvr_Autosenddrug;
/

Create Or Replace Procedure Zl_Drugsvr_Autoreturndrug
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：自动退药（按处方明细即费用ID退药，默认是全退）
  --入参：Json_In:格式
  --  input
  --    audit_operator        C 1 审核人
  --    rcpdtl_list[]   退药列表信息
  --         rcpdtl_id        N 1 费用id
  --         re_quantity      N 0 退药数量（销售单，收费项目目录.计算单位），可不传此结点或者传空，表示全退，否则按指定数量退药
  --出参: Json_Out,格式如下
  --  output
  --    code                 N 1   应答吗：0-失败；1-成功
  --    message              C 1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Jsonin     Pljson;
  j_Json       Pljson;
  j_Jsonlist   Pljson_List;
  d_操作时间   药品收发记录.审核日期%Type;
  v_操作员姓名 人员表.姓名%Type;
  n_数量       药品收发记录.实际数量%Type;
  n_退药数量   药品收发记录.实际数量%Type;
  n_小数       Number(6);
  n_费用id     Number(18);
Begin
  --解析入参
  j_Jsonin     := Pljson(Json_In);
  j_Json       := j_Jsonin.Get_Pljson('input');
  v_操作员姓名 := j_Json.Get_String('audit_operator');
  j_Jsonlist   := j_Json.Get_Pljson_List('rcpdtl_list');

  --取流通业务精度位数
  --类别:1-药品 2-卫材
  --内容：2-零售价 4-金额
  --单位：药品:1-售价 5-金额单位
  Select 精度 Into n_小数 From 药品卫材精度 Where 类别 = 1 And 内容 = 4 And 单位 = 5;
  Select Sysdate Into d_操作时间 From Dual;

  j_Json := Pljson();
  For I In 1 .. j_Jsonlist.Count Loop
    j_Json   := Pljson(j_Jsonlist.Get(I));
    n_费用id := j_Json.Get_Number('rcpdtl_id');
    n_数量   := j_Json.Get_Number('re_quantity');
    j_Json   := Pljson();
    If n_数量 Is Not Null Then
      n_退药数量 := n_数量;
    End If;
    --分解退药数量
    For r_处方明细 In (Select a.库房id, a.Id, Nvl(a.付数, 1) * a.实际数量 As 数量
                   From 药品收发记录 A
                   Where a.单据 In (8, 9, 10) And (a.记录状态 = 1 Or Mod(a.记录状态, 3) = 0) And 审核人 Is Not Null And
                         a.费用id = n_费用id
                   Order By a.库房id, a.药品id, a.批次) Loop
      If n_数量 Is Null Then
        --传入的数量为空表示全退      
        --部门退药（处方明细）
        Zl_药品收发记录_部门退药_s(r_处方明细.Id, v_操作员姓名, d_操作时间, Null, Null, Null, Null, Null, Null, n_小数);
      Else
        If n_退药数量 > 0 Then
          If n_退药数量 > r_处方明细.数量 Then
            --部门退药（处方明细）
            Zl_药品收发记录_部门退药_s(r_处方明细.Id, v_操作员姓名, d_操作时间, Null, Null, Null, r_处方明细.数量, Null, Null, n_小数);
          
            n_退药数量 := n_退药数量 - r_处方明细.数量;
          Else
            --部门退药（处方明细）
            Zl_药品收发记录_部门退药_s(r_处方明细.Id, v_操作员姓名, d_操作时间, Null, Null, Null, n_退药数量, Null, Null, n_小数);
            n_退药数量 := 0;
          End If;
        End If;
      End If;
    End Loop;
  End Loop;
  Json_Out := '{"output":{"code":1,"message": "成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugsvr_Autoreturndrug;
/
Create Or Replace Procedure Zl_Drugsvr_Checkpatiexecute
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：根据病人信息获取未发药品数据
  --入参：Json_In:格式
  --  input
  --     pati_id                N 1 病人ID
  --     pati_pageid            N 1 主页ID
  --     baby_num               N 0 婴儿序号:-1表示不区分;0-母亲的;>0具体婴儿费用
  --     check_excutenature     N 0 检查院带药：1-需要进行离院带药检查;0-不需要离院带药检查
  --     fee_source             N 1 费用来源:1-门诊;2-住院;4-体检
  --     rcp_nos            处方单据号，数组如：["A0001","A0002"]
  --出参: Json_Out,格式如下
  --  output
  --    code                    N 1 应答吗：0-失败；1-成功
  --    message                 C 1 应答消息：失败时返回具体的错误信息
  --    data{}
  --       isexist              N 1 是否存在: 1-存在;0-不存在
  --       drug_notsend_infor   C 1 未发药信息,isexist=1时返回
  ---------------------------------------------------------------------------
  v_Drug     Varchar2(4000);
  n_病人id   药品收发记录.病人id%Type;
  n_主页id   药品收发记录.主页id%Type;
  n_婴儿序号 药品收发记录.婴儿序号%Type;
  v_No       药品收发记录.No%Type;
  n_费用来源 药品收发记录.费用来源%Type;
  n_院外带药 Number(2);
  n_Add      Number(2);
  j_Jsonlist Pljson_List := Pljson_List();
  n_Count    Number(18);
  l_Nos      t_Strlist := t_Strlist();

  v_项目 收费项目目录.名称%Type;
  v_部门 部门表.名称%Type;
  v_扣率 Varchar2(100);

  Type t_未发药品 Is Ref Cursor;
  c_未发药品 t_未发药品;

  j_Jsonin Pljson;
  j_Json   Pljson;
Begin
  --解析入参
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_病人id   := j_Json.Get_Number('pati_id');
  n_主页id   := j_Json.Get_Number('pati_pageid');
  n_婴儿序号 := j_Json.Get_Number('baby_num');
  n_费用来源 := j_Json.Get_Number('fee_source');
  n_院外带药 := j_Json.Get_Number('check_excutenature');

  j_Jsonlist := j_Json.Get_Pljson_List('rcp_nos');
  n_Count    := 0;
  If j_Jsonlist Is Not Null Then
    n_Count := j_Jsonlist.Count;
  End If;
  If n_Count <> 0 Then
    For I In 1 .. n_Count Loop
      v_No := j_Jsonlist.Get_String(I);
      l_Nos.Extend();
      l_Nos(l_Nos.Count) := v_No;
    End Loop;
  
    Open c_未发药品 For
      Select Distinct b.No, d.名称 项目, c.名称 As 部门, To_Char(b.扣率) As 扣率
      From 药品收发记录 B, 部门表 C, 收费项目目录 D
      Where b.药品id = d.Id And b.库房id + 0 = c.Id(+) And b.单据 In (9, 10) And Mod(b.记录状态, 3) = 1 And b.审核人 Is Null And
            b.病人id = n_病人id And (Nvl(b.主页id, 0) = n_主页id And n_费用来源 = 2 Or n_费用来源 <> 2) And
            (Nvl(b.婴儿序号, 0) = Nvl(n_婴儿序号, 0) Or Nvl(n_婴儿序号, 0) = -1) And Nvl(b.摘要, '大医') <> '拒发' And
            b.No In (Select Column_Value From Table(l_Nos)) And Nvl(b.病人来源, 1) = n_费用来源;
  Else
  
    Open c_未发药品 For
      Select Distinct b.No, d.名称 项目, c.名称 As 部门, To_Char(b.扣率) As 扣率
      From 药品收发记录 B, 部门表 C, 收费项目目录 D
      Where b.药品id = d.Id And b.库房id + 0 = c.Id(+) And b.单据 In (9, 10) And Mod(b.记录状态, 3) = 1 And b.审核人 Is Null And
            b.病人id = n_病人id And (Nvl(b.主页id, 0) = n_主页id And n_费用来源 = 2 Or n_费用来源 <> 2) And
            (Nvl(b.婴儿序号, 0) = Nvl(n_婴儿序号, 0) Or n_婴儿序号 = -1) And Nvl(b.摘要, '大医') <> '拒发' And b.病人来源 = n_费用来源;
  
  End If;

  Loop
    Fetch c_未发药品
      Into v_No, v_项目, v_部门, v_扣率;
    Exit When c_未发药品%NotFound;
  
    n_Add := 1;
    If Substr(v_扣率, 2) = '3' Then
      n_Add := Nvl(n_院外带药, 0);
    End If;
  
    If Nvl(n_Add, 0) = 1 Then
    
      If v_Drug Is Not Null Then
        If Instr(Chr(13) || Chr(10) || v_Drug || Chr(13) || Chr(10),
                 Chr(13) || Chr(10) || '单据[' || Nvl(v_No, '') || ']中的' || Nvl(v_项目, '') || '：在' || Nvl(v_部门, '[未定药房]') ||
                  '未发药' || Chr(13) || Chr(10), 1) = 0 Then
          If Lengthb(v_Drug || Chr(13) || Chr(10) || '单据[' || Nvl(v_No, '') || ']中的' || Nvl(v_项目, '') || '：在' ||
                     Nvl(v_部门, '[未定药房]') || '未发药') <= 1000 Then
            v_Drug := v_Drug || Chr(13) || Chr(10) || '单据[' || Nvl(v_No, '') || ']中的' || Nvl(v_项目, '') || '：在' ||
                      Nvl(v_部门, '[未定药房]') || '未发药';
          Else
            v_Drug := v_Drug || Chr(13) || Chr(10) || '... ...';
          End If;
        End If;
      Else
        v_Drug := '单据[' || Nvl(v_No, '') || ']中的' || Nvl(v_项目, '') || '：在' || Nvl(v_部门, '[未定药房]') || '未发药';
      End If;
    End If;
  End Loop;

  n_Count := 0;
  If v_Drug Is Not Null Then
    v_Drug  := '存在未发药品：' || Chr(13) || Chr(10) || Chr(13) || Chr(10) || v_Drug;
    n_Count := 1;
  End If;

  Json_Out := '{"output":{"code":1,"message":"成功","data":{"isexist":' || n_Count || ',"drug_notsend_infor":"' ||
              Zljsonstr(v_Drug) || '"}}}';
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Checkpatiexecute;
/

Create Or Replace Procedure Zl_Drugsvr_Getrefusesendlist
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：获取拒发药清单
  --入参：Json_In:格式
  --  input
  --     pati_id            N 1 病人Id
  --     pati_pageids       C 1 主页IDs:多次住院 ，用逗号分离 
  --出参: Json_Out,格式如下
  --  output
  --    code                N   1   应答吗：0-失败；1-成功
  --    message             C   1   应答消息：失败时返回具体的错误信息
  --    data []
  --        rcp_no          C   1   费用单据号
  --        rcpdtl_id       C   1   处方明细ID,目前传入的是费用ID
  ---------------------------------------------------------------------------
  v_主页ids Varchar2(4000);
  n_病人id  药品收发记录.病人id%Type;

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;

  j_Jsonin PLJson;
  j_Json   PLJson;
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

  For r_Info In (Select NO As Rcp_No, 费用id As Rcpdtl_Id
                 From 药品收发记录
                 Where 病人id = n_病人id And
                       (Instr(',' || Nvl(v_主页ids, '-') || ',', ',' || Nvl(主页id, 0) || ',') > 0 Or v_主页ids Is Null) And
                       Mod(记录状态, 3) = 1 And Nvl(摘要, '大一') = '拒发' And Instr(',8,9,10,', ',' || 单据 || ',') > 0
                 Order By NO, 费用id) Loop
  
    v_Jtmp := v_Jtmp || ',';
    zlJsonPutValue(v_Jtmp, 'rcp_no', r_Info.Rcp_No, 0, 1);
    zlJsonPutValue(v_Jtmp, 'rcpdtl_id', r_Info.Rcpdtl_Id, 1, 2);
  
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
    Json_Out := '{"output":{"code":1,"message":"成功","data":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","data":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getrefusesendlist;
/

Create Or Replace Procedure Zl_Drugsvr_Getexecutednum
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：获取药品已发放数量
  --入参：Json_In:格式
  --  input
  --    billtype            N   1 单据类型:1-收费处方发药;2-记帐单处方发药;3-记帐表处方发药
  --    rcp_nos             C   1 单据号:可以传入多张单据
  --    notcontain_zero     N   1 是否不包含已发数量为0的：1-不包含，0-包含
  --    rcpdtl_ids          C   0 处方明细ids，多个用英文的逗号分隔,未传入时按单据号查找,传入时按明细id进行查找
  --    order_ids           C   0 医嘱id串，本次处理的一批医嘱id逗号分割
  --出参: Json_Out,格式如下
  --  output
  --    code                 N   1   应答码：0-失败；1-成功
  --    message              C   1   应答消息：失败时返回具体的错误信息
  --    data[]
  --       rcp_no            C   1   处方单号(费用单据号)
  --       rcpdtl_id         N   1   处方明细ID(费用ID)
  --       order_id          N   0   医嘱id
  --       drug_id           N   1   药品ID
  --       sended_num        N   1   已发数量
  ---------------------------------------------------------------------------
  n_不包含  Number(1);
  n_单据    药品收发记录.单据%Type;
  v_Nos     Varchar2(4000);
  v_明细ids Clob;

  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_明细ids Collection_Type;
  I           Number;

  v_Jtmp      Varchar2(32767);
  c_Order_Ids Clob;
  c_Jtmp      Clob;

  j_Jsonin Pljson;
  j_Json   Pljson;
Begin
  --解析入参
  j_Jsonin    := Pljson(Json_In);
  j_Json      := j_Jsonin.Get_Pljson('input');
  n_单据      := j_Json.Get_Number('billtype');
  v_Nos       := j_Json.Get_String('rcp_nos');
  n_不包含    := j_Json.Get_Number('notcontain_zero');
  v_明细ids   := j_Json.Get_Clob('rcpdtl_ids');
  c_Order_Ids := j_Json.Get_Clob('order_ids');

  If n_单据 = 1 Then
    n_单据 := 8;
  Elsif n_单据 = 2 Then
    n_单据 := 9;
  Elsif n_单据 = 3 Then
    n_单据 := 10;
  Elsif c_Order_Ids Is Null Then
    If v_明细ids Is Null Then
      Json_Out := Zljsonout('传入节点【billtype】错误，请检查！');
      Return;
    End If;
  End If;

  If v_明细ids Is Null Then
    For c_药品 In (Select /*+cardinality(j,10)*/
                  a.No, a.费用id, a.医嘱id, a.药品id, Sum(Nvl(a.付数, 1) * Decode(a.审核人, Null, 0, a.实际数量)) As 已发数量
                 From 药品收发记录 A, Table(f_Str2list(v_Nos)) J
                 Where a.No = j.Column_Value And
                       (a.单据 = 8 And n_单据 = 8 Or n_单据 <> 8 And Instr(',9,10,', ',' || a.单据 || ',') > 0 Or n_单据 Is Null) And
                       (c_Order_Ids Is Null Or Instr(',' || c_Order_Ids || ',', ',' || a.医嘱id || ',') > 0)
                 Group By a.No, a.费用id, a.药品id, a.医嘱id) Loop
    
      If Not (Nvl(n_不包含, 0) = 1 And Nvl(c_药品.已发数量, 0) = 0) Then
        v_Jtmp := v_Jtmp || ',';
        Zljsonputvalue(v_Jtmp, 'rcp_no', c_药品.No, 0, 1);
        Zljsonputvalue(v_Jtmp, 'rcpdtl_id', c_药品.费用id, 1);
        Zljsonputvalue(v_Jtmp, 'order_id', c_药品.医嘱id, 1);
        Zljsonputvalue(v_Jtmp, 'drug_id', c_药品.药品id, 1);
        Zljsonputvalue(v_Jtmp, 'sended_num', Nvl(c_药品.已发数量, 0), 1, 2);
      
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
      For c_药品 In (Select /*+cardinality(j,10)*/
                    a.No, a.费用id, a.药品id, Sum(Nvl(a.付数, 1) * Decode(a.审核人, Null, 0, a.实际数量)) As 已发数量
                   From 药品收发记录 A, Table(f_Num2list(Col_明细ids(I))) J
                   Where a.费用id = j.Column_Value
                   Group By a.No, a.费用id, a.药品id) Loop
      
        If Not (Nvl(n_不包含, 0) = 1 And Nvl(c_药品.已发数量, 0) = 0) Then
          v_Jtmp := v_Jtmp || ',';
          Zljsonputvalue(v_Jtmp, 'rcp_no', c_药品.No, 0, 1);
          Zljsonputvalue(v_Jtmp, 'rcpdtl_id', c_药品.费用id, 1);
          Zljsonputvalue(v_Jtmp, 'drug_id', c_药品.药品id, 1);
          Zljsonputvalue(v_Jtmp, 'sended_num', Nvl(c_药品.已发数量, 0), 1, 2);
        
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
    Json_Out := '{"output":{"code":1,"message":"成功","data":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","data":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getexecutednum;
/

Create Or Replace Procedure Zl_Drugsvr_Getsendwindows
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：获取药品的发药窗口
  --入参：Json_In:格式
  --  input
  --    item_list[]
  --        billtype        N   1   单据类型: 1 -收费处方发药  ;2- 记帐单处方发药;3- 记帐表处方发药
  --        pharmacy_id     N   1   药房id
  --        pati_id         N   1   病人id
  --        valid_days      N       未发药品记录查询范围的有效天数
  --        defaultwindow   C       根据各业务模块参数设置中传入的缺省窗口
  --出参: Json_Out,格式如下
  --  output
  --    code                N   1 应答吗：0-失败；1-成功
  --    message             C   1 应答消息：失败时返回具体的错误信息
  --    item_list                   [数组]每个库房ID对应的发药窗口
  --        pharmacy_id     N   1   库房ID
  --        pharmacy_window C   1   库房ID对应的发药窗口
  --窗口获取规则:
  --  判断指定病人在指定药房的未发药品记录中是否存在正在上班的发药窗口
  --  a.发药窗口存在，返回填制日期最近的发药窗口
  --  b.发药窗口不存在：
  --    i:如果存在缺省的发药窗口，且正在上班，则返回缺省的发药窗口，如果窗口未上班则返回null
  --    ii:如果不存在缺省的发药窗口，则根据动态分配规则（0-闲忙;1-平均）获取非专家的发药窗口
  ---------------------------------------------------------------------------
  j_Jsonlist Pljson_List;

  n_单据     药品收发记录.单据%Type;
  n_库房id   药品收发记录.库房id%Type;
  v_缺省窗口 未发药品记录.发药窗口%Type;
  v_发药窗口 未发药品记录.发药窗口%Type;
  n_病人id   Number(18);
  n_有效天数 Number(10);

  Type t_Record Is Record(
    药房id   Number(18),
    发药窗口 Varchar2(10));

  Type t_发药窗口 Is Table Of t_Record;
  c_发药窗口 t_发药窗口 := t_发药窗口();

  v_List  Varchar2(32767);
  n_Count Number;

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --解析入参
  j_Jsonin   := PLJson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  j_Jsonlist := j_Json.Get_Pljson_List('item_list');

  n_Count := j_Jsonlist.Count;
  If j_Jsonlist.Count = 0 Then
    Json_Out := zlJsonOut('未传入库房信息！');
    Return;
  End If;

  For I In 1 .. j_Jsonlist.Count Loop
    j_Json     := PLJson();
    j_Json     := PLJson(j_Jsonlist.Get(I));
    n_单据     := j_Json.Get_Number('billtype');
    n_库房id   := j_Json.Get_Number('pharmacy_id');
    n_病人id   := j_Json.Get_Number('pati_id');
    n_有效天数 := j_Json.Get_Number('valid_days');
    v_缺省窗口 := j_Json.Get_String('defaultwindow');
    If Nvl(n_有效天数, 0) = 0 Then
      n_有效天数 := 7;
    End If;
  
    --0)判断该药房是否已分配了发药窗口
    v_发药窗口 := Null;
    For I In 1 .. c_发药窗口.Count Loop
      If c_发药窗口(I).药房id = Nvl(n_库房id, 0) Then
        v_发药窗口 := c_发药窗口(I).发药窗口;
        Exit;
      End If;
    End Loop;
  
    If v_发药窗口 Is Null Then
      --1)查找指定病人在指定药房中分配了发药窗口且该窗口处于上班的最后一次未发药品的窗口，存在则取该发药窗口
      Select Max(发药窗口)
      Into v_发药窗口
      From (Select a.发药窗口
             From 未发药品记录 A
             Where a.单据 = Decode(n_单据, 1, 8, 2, 9, 10) And a.病人id = n_病人id And a.填制日期 Between Trunc(Sysdate) - n_有效天数 - 1 And
                   Sysdate And a.库房id = n_库房id And a.发药窗口 Is Not Null And Exists
              (Select 1 From 发药窗口 Where Nvl(上班否, 0) = 1 And 名称 = a.发药窗口 And 药房id = a.库房id)
             Order By a.填制日期 Desc)
      Where Rownum < 2;
    
      If v_发药窗口 Is Null Then
        --2)检查缺省窗口是否处于上班的，是上班的，取该发药窗口
        If v_缺省窗口 Is Not Null Then
          Select Count(1)
          Into n_Count
          From 发药窗口
          Where Nvl(上班否, 0) = 1 And 名称 = v_缺省窗口 And 药房id = n_库房id;
          If n_Count <> 0 Then
            v_发药窗口 := v_缺省窗口;
          End If;
        Else
          --3)根据药房窗口的分配规则（忙闲/平均 ），取发药窗口
          v_发药窗口 := Zl_Get发药窗口(n_库房id);
        End If;
      End If;
    
      If v_发药窗口 Is Not Null Then
        c_发药窗口.Extend;
        c_发药窗口(c_发药窗口.Count).药房id := n_库房id;
        c_发药窗口(c_发药窗口.Count).发药窗口 := v_发药窗口;
      End If;
    End If;
  End Loop;

  For I In 1 .. c_发药窗口.Count Loop
    zlJsonPutValue(v_List, 'pharmacy_id', c_发药窗口(I).药房id, 1, 1);
    zlJsonPutValue(v_List, 'pharmacy_window', c_发药窗口(I).发药窗口, 0, 2);
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || v_List || ']}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getsendwindows;
/

Create Or Replace Procedure Zl_Drugsvr_Getpharmacywindows
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：获取药房所涉及的发药窗口
  --入参：Json_In:格式
  --  input
  --    pharmacy_ids            C   1  药房ID1,药房ID2…
  --出参: Json_Out,格式如下
  --  output
  --    code                    N   1   应答码：0-失败；1-成功
  --    message                 C   1   每个药房id对应的发药窗口[数组]
  --    window_list[]    更新数据列表[数组]
  --        pharmacy_id             N 1 药房ID
  --        pharmacy_window         C 1 发药窗口
  --        expert_window           N 1 是否专家窗口：1-是，0-不是
  ---------------------------------------------------------------------------
  v_药房ids Varchar2(32767);
  v_Temp    Varchar2(32767);

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --解析入参
  j_Jsonin  := PLJson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  v_药房ids := j_Json.Get_String('pharmacy_ids');

  If v_药房ids Is Null Then
    Json_Out := zlJsonOut('未传入药房信息');
    Return;
  End If;

  For c_药品 In (Select /*+cardinality(b,10)*/
                a.药房id, a.名称, Nvl(a.专家, 0) As 专家
               From 发药窗口 A, Table(f_Num2List(v_药房ids)) B
               Where a.药房id = b.Column_Value
               Order By a.药房id, a.编码) Loop
  
    zlJsonPutValue(v_Temp, 'pharmacy_id', c_药品.药房id, 1, 1);
    zlJsonPutValue(v_Temp, 'pharmacy_window', c_药品.名称, 0);
    zlJsonPutValue(v_Temp, 'expert_window', c_药品.专家, 1, 2);
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"成功","window_list":[' || v_Temp || ']}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getpharmacywindows;
/

Create Or Replace Procedure Zl_Drugsvr_Getstockcheck
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：获取药品库存检查方式
  --入参：Json_In:格式
  --  input
  --出参: Json_Out,格式如下
  --  output
  --    code               N   1 应答吗：0-失败；1-成功
  --    message            C   1 应答消息：失败时返回具体的错误信息
  --    item_list                   [数组]
  --       pharmacy_id     N   1   药房ID
  --       check_type      N   1   检查方式：0-不检查，1-检查提示名，2-检查禁止
  --------------------------------------------------------------------------- 
  v_Output Varchar2(32767);
Begin

  For r_Data In (Select Distinct b.部门id, Nvl(c.检查方式, 0) As 检查方式
                 From 部门性质说明 B, 药品出库检查 C
                 Where b.部门id = c.库房id(+) And b.服务对象 In (1, 2, 3) And b.工作性质 In ('中药房', '西药房', '成药房')) Loop
  
    zlJsonPutValue(v_Output, 'pharmacy_id', r_Data.部门id, 1, 1);
    zlJsonPutValue(v_Output, 'check_type', r_Data.检查方式, 1, 2);
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || v_Output || ']}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getstockcheck;
/

Create Or Replace Procedure Zl_Drugsvr_Getstock
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：获取指定药品在指定库房的可用库存数
  --入参：Json_In:格式
  --  input
  --    drug_id             N   1   药品ID
  --    pharmacy_ids        C   1   库房ID
  --    batch               N       批次：<=0-不区分批次，>0只查某批次
  --出参: Json_Out,格式如下
  --  output
  --    code                N 1 应答吗：0-失败；1-成功
  --    message             C 1 应答消息：失败时返回具体的错误信息
  --    data                N  1  可用库存
  ---------------------------------------------------------------------------
  n_药品id   药品收发记录.药品id%Type;
  n_库存数量 药品库存.可用数量 %Type;
  v_库房ids  Varchar2(32767);

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --解析入参
  j_Jsonin  := PLJson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  n_药品id  := j_Json.Get_Number('drug_id');
  v_库房ids := j_Json.Get_String('pharmacy_ids');

  If Nvl(n_药品id, 0) = 0 Then
    Json_Out := zlJsonOut('未传入相关药品信息');
    Return;
  End If;

  --获取库存(不分批或分批),药房不分批(批次=0,这里为药房)不管效期
  Select Nvl(Sum(a.可用数量), 0)
  Into n_库存数量
  From 药品库存 A
  Where (Nvl(a.批次, 0) = 0 Or a.效期 Is Null Or a.效期 > Trunc(Sysdate)) And a.性质 = 1 And a.药品id = n_药品id And
        Instr(',' || v_库房ids || ',', ',' || a.库房id || ',') > 0;

  Json_Out := '{"output":{"code":1,"message":"成功","data":' || zlJsonStr(n_库存数量, 1) || '}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getstock;
/

Create Or Replace Procedure Zl_Drugsvr_Getstockbydept
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：根据指定药品及库房性质获取各库房的库存信息
  --入参：Json_In:格式
  --  input
  --    drug_id             N   1   药品ID
  --    pharmacy_nature     C   1  库房性质：中药房，西药房，成药房…
  --出参: Json_Out,格式如下
  --  output
  --    code                N   1 应答吗：0-失败；1-成功
  --    message             C   1 应答消息：失败时返回具体的错误信息
  --    data
  --        pharmacy_id     N   1   药房ID
  --        pharmacy_code   C   1   库房编码
  --        pharmacy_name   C   1   库房名称
  --        stock           N   1   可用数量
  ---------------------------------------------------------------------------
  n_药品id   药品收发记录.药品id%Type;
  v_库房性质 Varchar2(50);
  v_Output   Varchar2(32767);

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --解析入参
  j_Jsonin   := PLJson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_药品id   := j_Json.Get_Number('drug_id');
  v_库房性质 := j_Json.Get_String('pharmacy_nature');

  If Nvl(n_药品id, 0) = 0 Or Nvl(v_库房性质, '-') = '-' Then
    Json_Out := zlJsonOut('未传入相关药品信息');
    Return;
  End If;

  For c_库存 In (Select b.编码, b.名称, a.库房id, Nvl(Sum(a.可用数量), 0) As 库存
               From 药品库存 A,
                    (Select Distinct a.Id, a.编码, a.名称
                      From 部门表 A, 部门性质说明 B
                      Where a.Id = b.部门id And Instr(',' || v_库房性质 || ',', ',' || b.工作性质 || ',') > 0) B
               Where a.库房id = b.Id And ((a.效期 Is Null Or 效期 > Trunc(Sysdate)) Or Nvl(a.批次, 0) = 0) And a.性质 = 1 And
                     a.药品id = n_药品id
               Group By b.编码, b.名称, a.库房id
               Having Sum(Nvl(a.可用数量, 0)) <> 0
               Order By b.编码) Loop
  
    zlJsonPutValue(v_Output, 'pharmacy_id', c_库存.库房id, 1, 1);
    zlJsonPutValue(v_Output, 'pharmacy_code', c_库存.编码);
    zlJsonPutValue(v_Output, 'pharmacy_name', c_库存.名称);
    zlJsonPutValue(v_Output, 'stock', c_库存.库存, 1, 2);
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"成功","data":[' || v_Output || ']}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getstockbydept;
/

CREATE OR REPLACE Procedure Zl_Drugsvr_Getstockbatch
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：批量获取多个药品库存及价格信息:在项目选择器中展示库存及价格信息
  --入参：Json_In:格式
  --  input
  --   drug_ids                   C 1 药品ID，多个用英文的逗号分隔
  --   pharmacy_ids               C 0 库房ID，多个用英文的逗号分隔;空字符串,查询所有库房
  --   return_price               N 0 是否返回售价：1-返回价格信息(售价);0-不返回
  --   return_dept                N 0 按科室返回库存：1-按科室返回库存;0-按药品返回库存;2-返回科室所有药品的库存
  --   query_type                 N 1 查询类型:如：0-查询库存不等于0,1-查询库存小于等于0
  --出参: Json_Out,格式如下
  --  output
  --    code                      N 1 应答吗：0-失败；1-成功
  --    message                   C 1 应答消息：失败时返回具体的错误信息
  --    data[]
  --      drug_id                 N 1 药品ID
  --      pharmacy_id             N 1 库房ID(按科室返回库存才有此项)
  --      stock                   N 1 可用数量
  --      price                   N 1 零售价(返回价格时才有此项)
  ---------------------------------------------------------------------------
  v_药品ids  Clob;
  v_库房ids  Varchar2(32767);
  n_返回价格 Number(2);
  n_科室返回 Number(2);
  n_查询类型 Number(2);

  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_药品ids Collection_Type;
  I           Number;

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --解析入参
  j_Jsonin   := PLJson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  v_药品ids  := j_Json.Get_Clob('drug_ids');
  v_库房ids  := j_Json.Get_String('pharmacy_ids');
  n_返回价格 := Nvl(j_Json.Get_Number('return_price'), 0);
  n_科室返回 := j_Json.Get_Number('return_dept');
  n_查询类型 := Nvl(j_Json.Get_Number('query_type'), 0);

  I := 0;
  While v_药品ids Is Not Null Loop
    If Length(v_药品ids) <= 4000 Then
      Col_药品ids(I) := v_药品ids;
      v_药品ids := Null;
    Else
      Col_药品ids(I) := Substr(v_药品ids, 1, Instr(v_药品ids, ',', 3980) - 1);
      v_药品ids := Substr(v_药品ids, Instr(v_药品ids, ',', 3980) + 1);
    End If;
    I := I + 1;
  End Loop;

  I := 0;
  If Nvl(n_返回价格, 0) = 0 Then
    If Nvl(n_科室返回, 0) = 0 Then
      For I In 0 .. Col_药品ids.Count - 1 Loop
        If n_查询类型 = 0 Then
          For c_库存 In (With c_药品信息 As
                          (Select Column_Value As 药品id From Table(f_Num2List(Col_药品ids(I))))
                         Select /*+cardinality(b,10)*/
                          a.药品id, Sum(Nvl(a.可用数量, 0)) As 库存
                         From 药品库存 A, c_药品信息 B
                         Where a.药品id = b.药品id And a.性质 = 1 And
                               (Instr(',' || v_库房ids || ',', ',' || a.库房id || ',') > 0 Or v_库房ids Is Null) And
                               (Nvl(a.批次, 0) = 0 Or a.效期 Is Null Or a.效期 > Trunc(Sysdate))
                         Group By a.药品id
                         Having Sum (Nvl(a.可用数量, 0)) <> 0) Loop

            v_Jtmp := v_Jtmp || ',';
            zlJsonPutValue(v_Jtmp, 'drug_id', c_库存.药品id, 1, 1);
            zlJsonPutValue(v_Jtmp, 'stock', c_库存.库存, 1, 2);

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
          For c_库存 In (With c_药品信息 As
                          (Select Column_Value As 药品id From Table(f_Num2List(Col_药品ids(I))))
                         Select /*+cardinality(b,10)*/
                          a.药品id, Sum(Nvl(a.可用数量, 0)) As 库存
                         From 药品库存 A, c_药品信息 B
                         Where a.药品id = b.药品id And a.性质 = 1 And
                               (Instr(',' || v_库房ids || ',', ',' || a.库房id || ',') > 0 Or v_库房ids Is Null) And
                               (Nvl(a.批次, 0) = 0 Or a.效期 Is Null Or a.效期 > Trunc(Sysdate))
                         Group By a.药品id
                         Having Sum (Nvl(a.可用数量, 0)) <= 0) Loop

            v_Jtmp := v_Jtmp || ',';
            zlJsonPutValue(v_Jtmp, 'drug_id', c_库存.药品id, 1, 1);
            zlJsonPutValue(v_Jtmp, 'stock', c_库存.库存, 1, 2);

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
      End Loop;
    Elsif Nvl(n_科室返回, 0) = 1 Then
      For I In 0 .. Col_药品ids.Count - 1 Loop
        For c_库存 In (Select /*+cardinality(b,10)*/
                      a.药品id, Sum(Nvl(a.可用数量, 0)) As 库存, a.库房id
                     From 药品库存 A, Table(f_Num2List(Col_药品ids(I))) B
                     Where a.药品id = b.Column_Value And a.性质 = 1 And
                           (Instr(',' || v_库房ids || ',', ',' || a.库房id || ',') > 0 Or v_库房ids Is Null) And
                           (Nvl(a.批次, 0) = 0 Or a.效期 Is Null Or a.效期 > Trunc(Sysdate))
                     Group By a.药品id, a.库房id
                     Having Sum(Nvl(a.可用数量, 0)) <> 0) Loop

          v_Jtmp := v_Jtmp || ',';
          zlJsonPutValue(v_Jtmp, 'drug_id', c_库存.药品id, 1, 1);
          zlJsonPutValue(v_Jtmp, 'pharmacy_id', c_库存.库房id, 1);
          zlJsonPutValue(v_Jtmp, 'stock', c_库存.库存, 1, 2);

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
    Elsif Nvl(n_科室返回, 0) = 2 Then
      For c_库存 In (Select a.药品id, Sum(Nvl(a.可用数量, 0)) As 库存, a.库房id
                   From 药品库存 A,
                        (Select /*+cardinality(c,10)*/
                           Column_Value As 库房id
                          From Table(f_Num2List(Nvl(v_库房ids, 0)))) C
                   Where (Nvl(批次, 0) = 0 Or 效期 Is Null Or 效期 > Trunc(Sysdate)) And 性质 = 1 And a.库房id = c.库房id
                   Group By a.药品id, a.库房id
                   Having Sum(Nvl(a.可用数量, 0)) <> 0) Loop

        v_Jtmp := v_Jtmp || ',';
        zlJsonPutValue(v_Jtmp, 'drug_id', c_库存.药品id, 1, 1);
        zlJsonPutValue(v_Jtmp, 'pharmacy_id', c_库存.库房id, 1);
        zlJsonPutValue(v_Jtmp, 'stock', c_库存.库存, 1, 2);

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
      Json_Out := '{"output":{"code":1,"message":"成功","data":[' || Substr(v_Jtmp, 2) || ']}}';
    Else
      c_Jtmp   := c_Jtmp || v_Jtmp;
      Json_Out := '{"output":{"code":1,"message":"成功","data":[' || c_Jtmp || ']}}';
    End If;
    Return;
  End If;

  --包含价格
  If Nvl(n_科室返回, 0) = 0 Then
    For I In 0 .. Col_药品ids.Count - 1 Loop
      For c_库存 In (Select a.药品id, Nvl(a.库存, 0) As 库存, Decode(Nvl(b.是否变价, 0), 1, 0, Nvl(c.现价, 0)) As 价格
                   From (With c_药品信息 As (Select Column_Value As 药品id From Table(f_Num2List(Col_药品ids(I))))
                          Select /*+cardinality(b,10)*/
                           a.药品id, Sum(Nvl(a.可用数量, 0)) As 库存
                          From 药品库存 A, c_药品信息 B
                          Where a.药品id = b.药品id And a.性质 = 1 And
                                (Instr(',' || v_库房ids || ',', ',' || a.库房id || ',') > 0 Or v_库房ids Is Null) And
                                (Nvl(a.批次, 0) = 0 Or a.效期 Is Null Or a.效期 > Trunc(Sysdate))
                          Group By a.药品id
                          Having Sum(Nvl(a.可用数量, 0)) <> 0) A, 收费项目目录 B, 收费价目 C
                          Where a.药品id = c.收费细目id And a.药品id = b.Id And c.价格等级 Is Null And Sysdate Between c.执行日期 And
                                Nvl(c.终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))
                   ) Loop

        v_Jtmp := v_Jtmp || ',';
        zlJsonPutValue(v_Jtmp, 'drug_id', c_库存.药品id, 1, 1);
        zlJsonPutValue(v_Jtmp, 'pharmacy_id', 0, 1);
        zlJsonPutValue(v_Jtmp, 'stock', c_库存.库存, 1);
        zlJsonPutValue(v_Jtmp, 'price', c_库存.价格, 1, 2);

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
  Elsif n_科室返回 = 1 Then

    For I In 0 .. Col_药品ids.Count - 1 Loop
      For c_库存 In (Select a.药品id, Nvl(a.库存, 0) As 库存, Decode(Nvl(b.是否变价, 0), 1, 0, Nvl(c.现价, 0)) As 价格, a.库房id
                   From (With c_药品信息 As (Select Column_Value As 药品id From Table(f_Num2List(Col_药品ids(I))))
                          Select /*+cardinality(b,10)*/
                           a.药品id, Sum(Nvl(a.可用数量, 0)) As 库存, a.库房id
                          From 药品库存 A, c_药品信息 B
                          Where a.药品id = b.药品id And a.性质 = 1 And
                                (Instr(',' || v_库房ids || ',', ',' || a.库房id || ',') > 0 Or v_库房ids Is Null) And
                                (Nvl(a.批次, 0) = 0 Or a.效期 Is Null Or a.效期 > Trunc(Sysdate))
                          Group By a.药品id, a.库房id
                          Having Sum(Nvl(a.可用数量, 0)) <> 0) A, 收费项目目录 B, 收费价目 C
                          Where a.药品id = c.收费细目id And a.药品id = b.Id And c.价格等级 Is Null And Sysdate Between c.执行日期 And
                                Nvl(c.终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))
                   ) Loop

        v_Jtmp := v_Jtmp || ',';
        zlJsonPutValue(v_Jtmp, 'drug_id', c_库存.药品id, 1, 1);
        zlJsonPutValue(v_Jtmp, 'pharmacy_id', c_库存.库房id, 1);
        zlJsonPutValue(v_Jtmp, 'stock', c_库存.库存, 1);
        zlJsonPutValue(v_Jtmp, 'price', c_库存.价格, 1, 2);

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

  Elsif n_科室返回 = 2 Then
    For c_库存 In (Select a.药品id, Nvl(a.库存, 0) As 库存, Decode(Nvl(b.是否变价, 0), 1, 0, Nvl(c.现价, 0)) As 价格, a.库房id
                 From (Select a.药品id, Sum(Nvl(a.可用数量, 0)) As 库存, a.库房id
                        From 药品库存 A,
                             (Select /*+cardinality(c,10)*/
                                Column_Value As 库房id
                               From Table(f_Num2List(v_库房ids))) C
                        Where (Nvl(批次, 0) = 0 Or 效期 Is Null Or 效期 > Trunc(Sysdate)) And 性质 = 1 And a.库房id = c.库房id
                        Group By a.药品id, a.库房id
                        Having Sum(Nvl(a.可用数量, 0)) <> 0) A, 收费项目目录 B, 收费价目 C
                 Where a.药品id = c.收费细目id And a.药品id = b.Id And c.价格等级 Is Null And Sysdate Between c.执行日期 And
                       Nvl(c.终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))) Loop

      v_Jtmp := v_Jtmp || ',';
      zlJsonPutValue(v_Jtmp, 'drug_id', c_库存.药品id, 1, 1);
      zlJsonPutValue(v_Jtmp, 'pharmacy_id', c_库存.库房id, 1);
      zlJsonPutValue(v_Jtmp, 'stock', c_库存.库存, 1);
      zlJsonPutValue(v_Jtmp, 'price', c_库存.价格, 1, 2);

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
    Json_Out := '{"output":{"code":1,"message":"成功","data":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","data":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getstockbatch;
/

Create Or Replace Procedure Zl_Drugsvr_Getstockbydrugname
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：根据药名（品种）来获取各规格的可用库存
  --入参：Json_In:格式
  --  input 
  --    clinicdrug_id   N   1   药名ID
  --    pharmacy_id     N   1   药房ID；药房ID=0，表示所有的
  --    occasion        N   1   场合：1-门诊 ，2-住院
  --    show_unit       N   1   显示单位:0-售价单位;1-住院单位;2-门诊单位
  --    site_no         C   1   站点号
  --出参: Json_Out,格式如下
  --  output
  --    code            N   1   应答吗：0-失败；1-成功
  --    message         C   1   应答消息：失败时返回具体的错误信息
  --    data[]
  --      drug_code               药品编码
  --      drug_name               药品名称
  --      drug_spec               规格
  --      unit                    单位
  --      drug_id         N   1   药品ID
  --      pharmacy_name           药房名称
  --      pharmacy_id     N   1   药房ID
  --      stock           N   1   可用数量
  ---------------------------------------------------------------------------
  n_药名id   药品库存.药品id%Type;
  n_库房id   药品库存.库房id%Type;
  n_场合     Number(2);
  n_显示单位 Number(2);
  v_站点号   Varchar2(6);

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --解析入参
  j_Jsonin   := PLJson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_药名id   := j_Json.Get_Number('clinicdrug_id');
  n_库房id   := Nvl(j_Json.Get_Number('pharmacy_id'), 0);
  n_场合     := j_Json.Get_Number('occasion');
  n_显示单位 := j_Json.Get_Number('show_unit');
  v_站点号   := j_Json.Get_String('site_no');

  If Nvl(n_场合, 0) = 0 Then
    n_场合 := 1;
  End If;
  If n_显示单位 Is Null Then
  
    n_显示单位 := 0;
  End If;

  v_Jtmp := Null;
  For c_库存 In (Select d.编码, d.规格, d.名称, e.名称 As 药房, Max(Decode(n_显示单位, 0, d.计算单位, 1, a.门诊单位, a.住院单位)) As 单位, 1 As 住院包装,
                      Sum(Nvl(m.可用数量, 0) / Decode(n_显示单位, 0, 1, 1, a.门诊包装, a.住院包装)) As 可用数量, m.库房id, m.药品id
               From 药品库存 M, 药品规格 A, 收费项目目录 D, 部门表 E
               Where m.药品id = d.Id And m.药品id = a.药品id And m.库房id = e.Id And (m.库房id = n_库房id Or n_库房id = 0) And
                     (Nvl(m.批次, 0) = 0 Or m.效期 Is Null Or m.效期 > Trunc(Sysdate)) And a.药名id = n_药名id And
                     (d.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or d.撤档时间 Is Null) And d.服务对象 In (n_场合, 3) And
                     (d.站点 = v_站点号 Or d.站点 Is Null)
               Group By e.名称, d.编码, d.规格, d.名称, d.计算单位, m.库房id, m.药品id
               Having Sum(Nvl(m.可用数量, 0)) > 0
               Order By d.编码) Loop
  
    v_Jtmp := v_Jtmp || ',';
    zlJsonPutValue(v_Jtmp, 'drug_code', c_库存.编码, 0, 1);
    zlJsonPutValue(v_Jtmp, 'drug_name', c_库存.名称);
    zlJsonPutValue(v_Jtmp, 'drug_spec', c_库存.规格);
    zlJsonPutValue(v_Jtmp, 'unit', c_库存.单位);
    zlJsonPutValue(v_Jtmp, 'drug_ide', c_库存.药品id, 1);
    zlJsonPutValue(v_Jtmp, 'pharmacy_name', c_库存.药房);
    zlJsonPutValue(v_Jtmp, 'pharmacy_id', c_库存.库房id, 1);
    zlJsonPutValue(v_Jtmp, 'stock', Round(c_库存.可用数量, 5), 1, 2);
  
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
    Json_Out := '{"output":{"code":1,"message":"成功","data":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","data":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getstockbydrugname;
/

Create Or Replace Procedure Zl_Drugsvr_Batchgetprice
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：批量获取药品售价(自助机使用)
  --入参：Json_In:格式
  --  input
  --   drug_ids    C   1   药品ID，多个用英文的逗号分隔
  --出参: Json_Out,格式如下
  --  output
  --    code          N 1 应答吗：0-失败；1-成功
  --    message         C 1 应答消息：失败时返回具体的错误信息
  --    data
  --      drug_id N   1   药品ID
  --      price   N   1   零售价(返回价格时才有此项)
  ---------------------------------------------------------------------------
  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_药品ids Collection_Type;
  I           Integer;
  c_药品ids   Clob;

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --解析入参
  j_Jsonin  := PLJson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  c_药品ids := j_Json.Get_Clob('drug_ids');
  If c_药品ids Is Null Then
    Json_Out := zlJsonOut('未传入有效的药品id,请检查!');
  End If;

  I := 0;
  While c_药品ids Is Not Null Loop
    If Length(c_药品ids) <= 4000 Then
      Col_药品ids(I) := c_药品ids;
      c_药品ids := Null;
    Else
      Col_药品ids(I) := Substr(c_药品ids, 1, Instr(c_药品ids, ',', 3980) - 1);
      c_药品ids := Substr(c_药品ids, Instr(c_药品ids, ',', 3980) + 1);
    End If;
    I := I + 1;
  End Loop;

  v_Jtmp := Null;
  For I In 0 .. Col_药品ids.Count - 1 Loop
    --包含价格
    For c_库存 In (With c_药品信息 As
                    (Select /*+cardinality(D,10)*/
                     d.Column_Value As 药品id
                    From Table(f_Num2List(Col_药品ids(I))) D)
                   Select a.药品id, Sum(a.实际金额) / Sum(a.实际数量) As 价格
                   From 药品库存 A, c_药品信息 B
                   Where a.药品id = b.药品id And a.性质 = 1 And (Nvl(a.批次, 0) = 0 Or a.效期 Is Null Or a.效期 > Trunc(Sysdate))
                   Group By a.药品id
                   Having Sum (Nvl(a.实际数量, 0)) <> 0) Loop
    
      v_Jtmp := v_Jtmp || ',';
      zlJsonPutValue(v_Jtmp, 'drug_id', c_库存.药品id, 1, 1);
      zlJsonPutValue(v_Jtmp, 'price', c_库存.价格, 1, 2);
    
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
    Json_Out := '{"output":{"code":1,"message":"成功","data":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","data":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Batchgetprice;
/

Create Or Replace Procedure Zl_Drugsvr_Checkmedicinesended
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：检查指定医嘱的最近一次发送是否已发药
  --入参：Json_In:格式
  --  input
  --    item_list[]     发送的医嘱列表
  --       order_id       N 1 医嘱ID
  --       rcpno          C 1 单据号
  --出参: Json_Out,格式如下
  --  output
  --    code              N 1 应答吗：0-失败；1-成功
  --    message           C 1 应答消息：失败时返回具体的错误信息
  --    data              N 1 是否已发药 0-未发药，1-已发药
  ---------------------------------------------------------------------------
  j_Json_Tmp Pljson;
  j_Jsonlist Pljson_List;
  n_Tmp      Number;
  v_明细     Varchar2(32767);

  j_Jsonin Pljson;
  j_Json   Pljson;
Begin
  --解析入参
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  j_Jsonlist := j_Json.Get_Pljson_List('item_list');

  For I In 1 .. j_Jsonlist.Count Loop
    j_Json_Tmp := Pljson(j_Jsonlist.Get(I));
    v_明细     := v_明细 || ',' || j_Json_Tmp.Get_Number('order_id');
    v_明细     := v_明细 || ':' || j_Json_Tmp.Get_String('rcpno');
  End Loop;

  If v_明细 Is Not Null Then
    Select Count(1)
    Into n_Tmp
    From 药品收发记录 A,
         (Select /*+cardinality(b,10)*/
            To_Number(C1) As 医嘱id, C2 As NO
           From Table(Cast(f_Str2list2(Substr(v_明细, 2)) As t_Strlist2)) B) B
    Where a.No = b.No And a.医嘱id = b.医嘱id And a.单据 In (9, 10) And Mod(a.记录状态, 3) = 1 And a.审核人 Is Null And Rownum < 2;
  End If;

  If Nvl(n_Tmp, 0) > 0 Then
    n_Tmp := 0;
  Else
    n_Tmp := 1;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功","data":' || n_Tmp || '}}';
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Checkmedicinesended;
/

Create Or Replace Procedure Zl_Drugsvr_Autosplitspeci
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：针对中草药根据药名（品种）来自动分配药品(自动分解多个规格)
  --入参：Json_In:格式
  --  input
  --    clinic_drug_id    N 1 药名ID
  --    form              N 1 形态：0-散装;1-中药饮片;2-免煎剂
  --    quantity          N 1 数量，按剂量单位传入
  --    packages_num      N 1 付数
  --    pharmacy_id       N 1 药房ID
  --    occasion          N 1 场合：1-门诊 ，2-住院
  --    drug_ids          C 0 指定药品的分配
  --出参: Json_Out,格式如下
  --  output
  --    code              N   1   应答吗：0-失败；1-成功
  --    message           C   1   应答消息：失败时返回具体的错误信息
  --    data{}
  --       quantity_remain   N   0   剩余数量
  --       item_list[] 不能分配时该节点返回空
  --           drug_id           N   1   药品ID
  --           quantity          N   1   数量
  ---------------------------------------------------------------------------
  n_药名id   药品规格.药名id%Type;
  n_形态     药品规格.中药形态%Type;
  n_数量     药品库存.可用数量%Type;
  n_付数     药品库存.可用数量%Type;
  n_药房id   药品库存.库房id%Type;
  n_场合     Number(1);
  n_药品id   药品规格.药品id%Type;
  v_药品ids  Varchar2(4000);
  n_剩余数量 药品库存.可用数量%Type;
  v_Result   Varchar2(4000);
  v_药品信息 Varchar2(200);
  v_Json     Varchar2(32767);

  j_Jsonin Pljson;
  j_Json   Pljson;
Begin
  --解析入参
  j_Jsonin  := Pljson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  n_药名id  := j_Json.Get_Number('clinic_drug_id');
  n_形态    := j_Json.Get_Number('form');
  n_数量    := j_Json.Get_Number('quantity');
  n_付数    := j_Json.Get_Number('packages_num');
  n_药房id  := j_Json.Get_Number('pharmacy_id');
  n_场合    := j_Json.Get_Number('occasion');
  v_药品ids := j_Json.Get_String('drug_ids');

  v_Result := Zl_Dispensechspecs(n_药名id, n_形态, n_数量, n_付数, n_药房id, 0, n_场合, v_药品ids);
  If v_Result Is Null Then
    Json_Out := '{"output":{"code":1,"message":"成功","data":{"quantity_remain":0,"item_list":[]}}}';
  Else
    If Instr(v_Result, '|') > 0 Then
      n_剩余数量 := To_Number(Substr(v_Result, Instr(v_Result, '|') + 1, Length(v_Result)));
      v_Result   := Substr(v_Result, 1, Instr(v_Result, '|') - 1);
    End If;
    For r_Row In (Select /*+cardinality(x,10)*/
                   x.Column_Value As 药品信息
                  From Table(Cast(f_Str2list(v_Result, ';') As t_Strlist)) X) Loop
      v_药品信息 := r_Row.药品信息;
      --分解
      n_药品id   := To_Number(Substr(v_药品信息, 1, Instr(v_药品信息, ',') - 1));
      v_药品信息 := Substr(v_药品信息, Instr(v_药品信息, ',') + 1);
      n_数量     := To_Number(v_药品信息);
      v_Json     := v_Json || ',{"drug_id":' || n_药品id;
      v_Json     := v_Json || ',"quantity":' || Nvl(n_数量, 0);
      v_Json     := v_Json || '}';
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"成功","data":{';
    Json_Out := Json_Out || '"quantity_remain":' || Nvl(n_剩余数量, 0);
    Json_Out := Json_Out || ',"item_list":[' || Substr(v_Json, 2) || ']}}}';
  End If;
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Autosplitspeci;
/

Create Or Replace Procedure Zl_Drugsvr_Getprice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：获取指定药品的售价、成本价
  --入参：Json_In:格式
  --  input
  --    item_list[]列表
  --          drug_id          N 1 药品ID
  --          pharmacy_id      N 1 药房ID
  --          quantity         N 1 数量
  --出参: Json_Out,格式如下
  --  output
  --    code              N 1   应答吗：0-失败；1-成功
  --    message           C 1   应答消息：失败时返回具体的错误信息
  --    data[]列表
  --          price             N 1 售价
  --          price_cost        N 1 成本价
  --          quantity_remain   N 1 剩余数量，等于0表示数量足够，大于0则表示数量不够
  ---------------------------------------------------------------------------
  n_药品id   药品库存.药品id%Type;
  n_药房id   药品库存.库房id%Type;
  n_数量     药品库存.实际数量%Type;
  v_Temp     Varchar2(4000);
  n_单价     药品规格.成本价%Type;
  n_成本价   药品规格.成本价%Type;
  n_剩余数量 药品库存.实际数量%Type;
  j_List     Pljson_List := Pljson_List();
  j_Tmpout   Varchar2(32767);
  j_Jsonin   Pljson;
  j_Json     Pljson;

Begin
  --解析入参
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  j_List   := j_Json.Get_Pljson_List('item_list');
  j_Json   := Pljson();
  If j_List Is Not Null Then
    For I In 1 .. j_List.Count Loop
      j_Json   := Pljson(j_List.Get(I));
      n_药品id := j_Json.Get_Number('drug_id');
      n_药房id := j_Json.Get_Number('pharmacy_id');
      n_数量   := j_Json.Get_Number('quantity');
    
      v_Temp := Zl_Fun_Getprice(n_药品id, n_药房id, n_数量, 0);
      --分解
      n_单价     := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1));
      v_Temp     := Substr(v_Temp, Instr(v_Temp, '|') + 1);
      n_成本价   := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1));
      v_Temp     := Substr(v_Temp, Instr(v_Temp, '|') + 1);
      n_剩余数量 := To_Number(v_Temp);
      j_Tmpout   := j_Tmpout || ',{"price":' || Zljsonstr(Nvl(n_单价, 0), 1);
      j_Tmpout   := j_Tmpout || ',"price_cost": ' || Zljsonstr(Nvl(n_成本价, 0), 1);
      j_Tmpout   := j_Tmpout || ',"quantity_remain":' || Zljsonstr(Nvl(n_剩余数量, 0), 1);
      j_Tmpout   := j_Tmpout || '}';
    
      j_Json := Pljson();
    End Loop;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功","data":[' || Substr(j_Tmpout, 2) || ']}}';

Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getprice;
/

Create Or Replace Procedure Zl_Drugsvr_Getcargospace
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：获取指定药品的指定库房的货位信息
  --入参：Json_In:格式
  --  input
  --   drug_id            N   1   药品ID
  --   pharmacy_id        N   1   库房ID
  --出参: Json_Out,格式如下
  --  output
  --    code              N   1   应答吗：0-失败；1-成功
  --    message           C   1   应答消息：失败时返回具体的错误信息
  --    data              C   1   库房货位
  ---------------------------------------------------------------------------
  n_药品id 药品库存.药品id%Type;
  n_库房id 药品库存.库房id%Type;
  v_货位   药品储备限额.库房货位%Type;

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --解析入参
  j_Jsonin := PLJson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_药品id := j_Json.Get_Number('drug_id');
  n_库房id := j_Json.Get_Number('pharmacy_id');

  --读取库房货位
  Select Max(库房货位) Into v_货位 From 药品储备限额 Where 药品id = n_药品id And 库房id = n_库房id;

  Json_Out := '{"output":{"code":1,"message":"成功","data":"' || v_货位 || '"}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getcargospace;
/

Create Or Replace Procedure Zl_Drugsvr_Checkstorelimit
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：检查指定药品在指定库房的库存是否低于储备下限
  --入参：Json_In:格式
  --  input
  --   drug_id            N   1   药品ID
  --   pharmacy_id        N   1   库房ID
  --   stock              N   1   库存数量
  --出参: Json_Out,格式如下
  --  output
  --    code              N 1 应答吗：0-失败；1-成功
  --    message           C 1 应答消息：失败时返回具体的错误信息
  --    below_limit_lower N 1 1-低于储备下限，0-大于等于储备下限
  ---------------------------------------------------------------------------
  n_药品id 药品库存.药品id%Type;
  n_库房id 药品库存.库房id%Type;
  n_库存   药品库存.实际数量%Type;
  n_Count  Number(1);

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --解析入参
  j_Jsonin := PLJson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_药品id := j_Json.Get_Number('drug_id');
  n_库房id := j_Json.Get_Number('pharmacy_id');
  n_库存   := j_Json.Get_Number('stock');

  --读取药品储备限额
  Select Count(1)
  Into n_Count
  From 药品储备限额
  Where 药品id = n_药品id And 库房id = n_库房id And Nvl(下限, 0) <> 0 And 下限 > n_库存 And Rownum < 2;

  Json_Out := '{"output":{"code":1,"message":"成功","below_limit_lower":' || n_Count || '}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Checkstorelimit;
/

Create Or Replace Procedure Zl_Drugsvr_Checkisputdrug
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：检查药品是否已进行配药
  --入参：Json_In:格式
  --  input
  --   rcp_nos            C  1  药品收发记录.no，多个用英文逗号分隔
  --   billtype           N  1  1-收费处方；2-记帐单处方；3-记帐表处方
  --出参: Json_Out,格式如下
  --  output
  --    code              N   1   应答吗：0-失败；1-成功
  --    message           C   1   应答消息：失败时返回具体的错误信息
  --    isputdrug         N   1   是否已配药：1-已配药,0-未配药
  ---------------------------------------------------------------------------
  v_Rcp_Nos  Varchar2(4000);
  n_Billtype Number(1);
  n_Count    Number(1);

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --解析入参
  j_Jsonin   := PLJson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  v_Rcp_Nos  := j_Json.Get_String('rcp_nos');
  n_Billtype := j_Json.Get_Number('billtype');

  Select /*+cardinality(b,10) */
   Count(1)
  Into n_Count
  From 未发药品记录 A, Table(f_Str2List(v_Rcp_Nos)) J
  Where a.No = j.Column_Value And a.单据 = Decode(n_Billtype, 1, 8, 2, 9, 10) And a.配药人 Is Not Null And Rownum <= 1;

  Json_Out := '{"output":{"code":1,"message":"成功","isputdrug":' || n_Count || '}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Checkisputdrug;
/

Create Or Replace Procedure Zl_Drugsvr_Adjustdata
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：门诊费用转住院时调整药品关联数据
  --入参：Json_In:格式
  --  input
  --    pati_id          N  1 病人ID
  --    pati_pageid      N  1 主页ID
  --    billtype         N  1 单据类型：1-收费单;2-记帐单
  --    item_list
  --      rcp_no_old       C  1 原单据号
  --      rcpdtl_id_old    N  1 原处方明细ID(目前传入的是费用id)
  --      rcp_no_new       C  1 新单据号
  --      rcpdtl_id_new    N  1 新处方明细ID(目前传入的是费用id)
  --出参: Json_Out,格式如下
  --  output
  --    code              N   1   应答码：0-失败；1-成功
  --    message           C   1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Jsonlist Pljson_List;
  n_病人id   药品收发记录.病人id%Type;
  n_主页id   药品收发记录.主页id%Type;
  n_单据     药品收发记录.单据%Type;

  v_原单据号 未发药品记录.No%Type;
  n_原明细id 药品收发记录.费用id%Type;
  v_新单据号 药品收发记录.No%Type;
  n_新明细id 药品收发记录.费用id%Type;

  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --解析入参
  j_Jsonin   := PLJson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_病人id   := j_Json.Get_Number('pati_id');
  n_主页id   := j_Json.Get_Number('pati_pageid');
  n_单据     := j_Json.Get_Number('billtype');
  j_Jsonlist := j_Json.Get_Pljson_List('item_list');

  If j_Jsonlist.Count = 0 Then
    Json_Out := zlJsonOut('未传入药品单据信息');
    Return;
  End If;

  For I In 1 .. j_Jsonlist.Count Loop
    j_Json := PLJson();
  
    j_Json     := PLJson(j_Jsonlist.Get(I));
    v_原单据号 := j_Json.Get_String('rcp_no_old');
    n_原明细id := j_Json.Get_Number('rcpdtl_id_old');
    v_新单据号 := j_Json.Get_String('rcp_no_new');
    n_新明细id := j_Json.Get_Number('rcpdtl_id_new');
  
    Update 未发药品记录
    Set 单据 = Decode(单据, 8, 9, 单据), 主页id = n_主页id, NO = v_新单据号
    Where NO = v_原单据号 And 单据 = Decode(n_单据, 1, 8, 2, 9, 10) And 病人id = n_病人id;
  
    Update 药品收发记录
    Set 单据 = Decode(单据, 8, 9, 单据), 费用id = n_新明细id, NO = v_新单据号, 主页id = n_主页id, 费用来源 = 2, 病人来源 = 2
    Where NO = v_原单据号 And 单据 = Decode(n_单据, 1, 8, 2, 9, 10) And 费用id = n_原明细id;
  End Loop;

  Json_Out := zlJsonOut('成功', 1);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugsvr_Adjustdata;
/

Create Or Replace Procedure Zl_Drugsvr_Recipeaffirm
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：药品处方记帐确认或药品处方收费确认
  --入参：Json_In:格式
  --  input
  --    pati_id                  N  0 病人id：针对收费时有效
  --    pati_name                C  0 姓名：针对收费时有效
  --    pati_sex                 C  0 性别：针对收费时有效
  --    pati_age                 C  0 年龄：针对收费时有效
  --    auditor                  C  1 审核人
  --    auditor_code             C  1 审核人编号
  --    audit_time               C  1 审核时间：yyyy-mm-dd hh24:mi:ss
  --    item_list[]                   更新数据列表[数组]
  --      billtype               N  1 单据类型:1 -收费处方发药  ;2- 记帐单处方发药;3- 记帐表处方发药
  --      rcp_no                 C  1 单据号
  --      rcpdtl_ids             C  0 费用ID,可以传入多个,用逗号分离
  --      pharmacy_window        C  0 发药窗口:发药窗口1:药房ID1| …|发药窗口n:药房Idn
  --      drug_auto_send         N  0 是否自动发放药品:0-不自动发药,1-自动发药
  --      auto_send_ids          C  0 自动发药的明细id数组,多个用逗号分隔
  --出参: Json_Out,格式如下
  --  output
  --    code              N   1   应答吗：0-失败；1-成功
  --    message           C   1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Jsonlist    Pljson_List;
  n_病人id      药品收发记录.病人id%Type;
  v_姓名        药品收发记录.姓名%Type;
  v_性别        药品收发记录.性别%Type;
  v_年龄        药品收发记录.年龄%Type;
  v_单据号      药品收发记录.No%Type;
  v_明细ids     Varchar2(4000);
  v_发药窗口s   Varchar2(4000);
  n_单据        药品收发记录.单据%Type;
  n_单据_In     药品收发记录.单据%Type;
  d_审核时间    Date;
  n_自动发药    Number(1);
  v_Err         Varchar2(255);
  v_发药明细id  Varchar2(400);
  v_发药明细ids Varchar2(4000);
  v_Nos         Varchar2(32767);
  v_审核人      人员表.姓名%Type;
  v_审核人编号  人员表.编号%Type;
  Err_Custom Exception;

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --解析入参
  j_Jsonin     := PLJson(Json_In);
  j_Json       := j_Jsonin.Get_Pljson('input');
  n_病人id     := j_Json.Get_Number('pati_id');
  v_姓名       := j_Json.Get_String('pati_name');
  v_性别       := j_Json.Get_String('pati_sex');
  v_年龄       := j_Json.Get_String('pati_age');
  v_审核人     := j_Json.Get_String('auditor');
  v_审核人编号 := j_Json.Get_String('auditor_code');
  d_审核时间   := To_Date(j_Json.Get_String('audit_time'), 'yyyy-mm-dd hh24:mi:ss');
  j_Jsonlist   := j_Json.Get_Pljson_List('item_list');

  If d_审核时间 Is Null Then
    d_审核时间 := Sysdate;
  End If;

  If j_Jsonlist.Count = 0 Then
    v_Err := '未传入药品单据信息！';
    Raise Err_Custom;
  End If;

  For I In 1 .. j_Jsonlist.Count Loop
    j_Json       := PLJson();
    j_Json       := PLJson(j_Jsonlist.Get(I));
    n_单据       := j_Json.Get_Number('billtype');
    v_单据号     := j_Json.Get_String('rcp_no');
    v_明细ids    := j_Json.Get_String('rcpdtl_ids');
    v_发药窗口s  := j_Json.Get_String('pharmacy_window');
    n_自动发药   := j_Json.Get_Number('drug_auto_send');
    v_发药明细id := j_Json.Get_String('auto_send_ids');
  
    n_单据_In := n_单据;
    If n_单据 = 1 Then
      n_单据 := 8;
    Elsif n_单据 = 2 Then
      n_单据 := 9;
    Elsif n_单据 = 3 Then
      n_单据 := 10;
    Else
      v_Err := '传入单据类型无效，请检查！';
      Raise Err_Custom;
    End If;
  
    If Nvl(v_单据号, '-') = '-' Then
      v_Err := '未传入处方单号，请检查！';
      Raise Err_Custom;
    End If;
  
    If Nvl(n_自动发药, 0) = 1 Then
      If v_发药明细id Is Null Then
        v_Nos := v_Nos || ',' || v_单据号;
      Else
        v_发药明细ids := v_发药明细ids || ',' || v_发药明细id;
      End If;
    End If;
  
    If Nvl(v_明细ids, '-') = '-' Then
      v_Err := '未传入处方明细ID，请检查！';
      Raise Err_Custom;
    End If;
  
    Zl_药品收发记录_费用审核(n_单据, v_单据号, v_明细ids, d_审核时间, v_发药窗口s, n_病人id, v_姓名, v_性别, v_年龄);
  End Loop;

  --按处方号发药
  If Nvl(v_Nos, '-') <> '-' Then
    v_Nos := Substr(v_Nos, 2);
    Zl_药品收发记录_自动发药_s(n_单据, v_审核人, v_审核人编号, v_Nos, 0);
  End If;

  --按费用ID发药
  If Nvl(v_发药明细ids, '-') <> '-' Then
    v_发药明细ids := Substr(v_发药明细ids, 2);
    Zl_药品收发记录_自动发药_s(n_单据, v_审核人, v_审核人编号, v_发药明细ids, 1);
  End If;

  Json_Out := zlJsonOut('成功', 1);
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugsvr_Recipeaffirm;
/

Create Or Replace Procedure Zl_Drugsvr_Gettakeno
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  --功能：获取已经生成的领药号
  --入参：Json_In
  --input
  --       dept_id  N 1 部门ID或病区ID
  --出参：Json_Out
  --output
  --       code     N 1 应答码：0-失败；1-成功
  --       message  C 1 应答消息：失败时返回具体的错误信息
  --       data     C 1 领药号
  n_部门id  未发药品记录.对方部门id%Type;
  v_Takenos Varchar2(5000);

  j_Jsonin Pljson;
  j_Json   Pljson;
Begin
  --解析入参
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_部门id := j_Json.Get_Number('dept_id');

  For R In (Select Distinct a.领药号
            From 未发药品记录 A
            Where a.填制日期 >= Trunc(Sysdate) And a.单据 = 9 And a.对方部门id = n_部门id And a.领药号 Is Not Null
            Order By a.领药号 Desc) Loop
    v_Takenos := v_Takenos || ',' || r.领药号;
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"成功","data":"' || Substr(v_Takenos, 2) || '"}}';
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Gettakeno;
/

Create Or Replace Procedure Zl_Drugsvr_Newrecipebill
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

  ---------------------------billtype = 3,记帐表处方（药品收发记录.单据=10）时，无以下节点--------------------------------------
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
  --     pati_type					        C	  1	病人类型：普通病人,医保病人,城乡居民,城镇职工...
  ---------------------------billtype = 3,记帐表处方（药品收发记录.单据=10）时，无以上节点-----------------------------------------

  --     si_type_id                 N     保险类型id(新门诊):ZLHIS是险类(序号)
  --     si_type_name               C     保险类型名称(新门诊)
  --     rgst_id                    N   1 挂号单id（新门诊)
  --     recipe_proxy_name          C     代办人姓名（新门诊)
  --     recipe_proxy_idno          C     代办人身份证号（新门诊)
  --     recipe_pat_bodywt          C     患者体重（新门诊)
  --     recipe_pat_bodywt_unit     C     患者体重单位（新门诊)  
  --     diag_list[]                      病人临床诊断列表[数组]，（新部门发药）  
  --        diag_rec_id             N     诊断记录id ，（新部门发药） 
  --        diag_type               N     诊断类型 1-西医门诊诊断;2-西医入院诊断;3-西医出院诊断;5-院内感染;6-病理诊断;7-损伤中毒码,8-术前诊断;9-术后诊断;10-并发症;11-中医门诊诊断;12-中医入院诊断;13-中医出院诊断;21-病原学诊断;22-影像学诊断，（新部门发药）  
  --        diag_name               C     临床诊断名称，（新部门发药）  
  
  --     pivas_info                 C   0 静配中心数据生成入参，可以不传 结点为一个json格式，明细格式同：Zl_Pivassvr_Newbill 服务的入参

  --     bill_list[]                      更新数据列表[数组]
  --        recipe_id                 N  1 处方id(新门诊):ZLHIS无，暂用NO转数字(字母用Asci代替)+序号替代
  --        rcp_no                    C  1 NO
  --        recipe_type               N  0 处方类型:0和空-普通,1-儿科,2-急诊,3-精二,4-精一,5-麻醉
  --        charge_tag                N  1 收费标志:0-未收费或记帐划价;1-已收费或记帐
  --        fee_acnter                C    划价人
  --        recipe_plcdept_id         C    开单科室id（新门诊)
  --        recipe_plcdept            C    开单科室名称（新门诊)
  --        recipe_placer_id          C    开单医师id（新门诊)
  --        recipe_placer             C    开单医师（新门诊) 增加
  --        apply_fee_category_code   C    申请单费别编码(医疗付款方式编码)(新门诊)  增加；
  --        apply_fee_category_name   C    申请单费别名称（医疗付款方式名称）(新门诊)  增加；
  --        operator_name             C  1 操作员姓名
  --        operator_code             C  1 操作员编号
  --        create_time               C  1 登记时间:yyyy-mm-dd hh:mi:ss
  --        take_no                   C    领药号 领药号，未发药品记录.领药号，药品收发记录.产品合格证，医嘱发送时生成
  --        item_list[]                    更新数据列表[数组]

  ---------------------------billtype = 3,记帐表处方（药品收发记录.单据=10）时，有以下节点----------------------------------------
  --           pati_id                 N  1 病人ID
  --           pati_pageid             N    主页ID
  --           pati_name               C  1 病人姓名
  --           pati_sex_code           C  1 性别编号（新门诊)
  --           pati_sex                C  1 性别
  --           pati_age                C  1 年龄
  --           pati_identity           C    身份
  --           pati_birthdate          C    出生日期:yyyy-mm-dd hh:mi:ss
  --           pati_idcard             C    身份证号
  --           pati_wardarea_id        N    病人病区ID
  --           pati_deptid             N  1 病人科室ID
  ---------------------------billtype = 3,记帐表处方（药品收发记录.单据=10）时，有以上节点-----------------------------------------

  --           rcpdtl_id               N  1 处方明细ID
  --           serial_num              N  1 序号:(变更(包括存储)：序号和组号，1、2、3、3、3、4…)
  --           group_sno               N    组内序号 (包括存储)：1、2、3
  --           pharmacy_id             N  1 药房ID
  --           pharmacy_name           C  1 药房名称(新门诊)
  --           takedept_id             N  1 领药部门ID:针对住院才传入
  --           cadn_id                 N  1 药品通用名称id(药名ID)(新门诊)
  --           drug_id                 N  1 药品ID
  --           drug_type			         N	1	药品类型：5-西药，6-成药，7-中药，（新部门发药）
  --           baby_num                N    婴儿序号

  --           advice_id               N  0 医嘱ID
  --           drug_method_id          N  1 给药途径id(新门诊):诊疗项目ID
  --           drug_method_name        C  1 给药途径名称
  --           drug_method_class_code  C  1 给药途径分类(新门诊):执行分类编号
  --           drug_freq_id            N  1 给药频次id(新门诊):诊疗频率编码
  --           drug_freq_name          C  1 给药频次名称d(新门诊):

  ---------------------------以下节点为可选参数，医嘱记录产生-----------------------------------------------
  --           emergency_tag           N    医嘱记录中的紧急标志(0-普通;1-紧急;2-补录(对门诊无效))
  --           effectivetime           N  0 医嘱期效
  --           fee_mode                N  0 计价特性：0-正常计价；1-不计价；2-手工计价
  --           use_mode                N  0 取药特性：0-正常方式，1-离院带药，2-自取药
  --           frequency               N  0 频次
  --           single                  N  0 单量
  --           usage                   C  0 用法
  --           rcpdtl_st_result        N    皮试结果(新门诊)1-阴性，2-阳性，3-免试，4-连续用药 处方生成时已确定或已有皮试结果。ZLHIS目前支持不全
  --           rcpdtL_excs_desc        C    超量说明(新门诊)
  --           rcpdtL_drask            C    使用嘱托(新门诊)
  --           disps_mode_code         C  1 发药方式(新门诊)1-正常发放；2-科室贮药；3-自备药；4-代购药 ZLHIS目前支持不全(2,4)
  --           drug_content            N    药品含量（剂量系数）(新门诊)：
  --           rcpdtl_outp_drugdays    N    本院门诊执行天数(新门诊)：ZLHIS是给药执行次数，要转换为天数传
  --           decoction_method        C  0 煎法
  --           advice_purpose			     C		用药目的，（新部门发药）
  ---------------------------以上节点为可选参数，医嘱记录产生-----------------------------------------------

  --           packages_num            N  1 发药付数
  --           send_num                N  1 发药数量
  --           send_unit               C  1 发药单位：zlhis零售单位
  --           price                   N    售价
  --           money                   N    零售金额(新门诊)
  --           pharmacy_window         C  0 发药窗口
  --           memo                    C  0 摘要
  --           fee_source              N  0 费用来源
  --           drug_auto_send          N  0 是否自动发放药品:0-不自动发药,1-自动发药
  --           diag_name               C  0 诊断名称（新门诊)仅门诊传入，诊断描述

  --出参: Json_Out,格式如下
  --  output
  --    code                 N   1   应答吗：0-失败；1-成功
  --    message              C   1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------  
Begin
  --直接调用药品业务过程（入参格式一致）
  Zl_药品收发记录_Newdrugbill(Json_In, Json_Out);

  Json_Out := Zljsonout('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugsvr_Newrecipebill;
/

Create Or Replace Procedure Zl_Drugsvr_Surplus
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  --功能：药品留存登记功能读取药品未发汇总
  --入参: Json格式
  --input
  --    dept_id                 N 1 药房id
  --    regbegin_time           C 1 填制日期起始
  --    regend_time             C 1 填制日期结束
  --    ward_id                 N 1 病区ID
  --    drugname_show_type      N 1 药品名称显示方式
  --    drug_method_ids         C 1 给药途径
  --    site_no                 C 1 站点
  --出参 Json 格式
  --output
  --    code                    N 1 应答码：0-失败；1-成功
  --    message                 C 1 应答消息：失败时返回具体的错误信息
  --    data[]列表信息
  --       drug_id              N 1 药品id
  --       rcpdtl_code          C 1 编码
  --       drug_name            C 1 药品名称
  --       unit                 C 1 单位
  --       spec                 C 1 规格
  --       place_name           C 1 产地
  --       quantity             N 1 填写数量
  --       category             C 1 类别
  --       inpack               N 1 住院包装
  n_Deptid       Number(18);
  n_Ward_Id      Number(18);
  n_Show_Type    Number(1);
  v_Method_Ids   Varchar2(4000);
  v_Site_No      Varchar2(50);
  d_Begindate    Date;
  d_Enddate      Date;
  v_Method_Names Varchar2(4000);

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;

  j_Jsonin Pljson;
  j_Json   Pljson;
Begin
  --解析入参
  j_Jsonin     := Pljson(Json_In);
  j_Json       := j_Jsonin.Get_Pljson('input');
  n_Deptid     := j_Json.Get_Number('dept_id');
  d_Begindate  := To_Date(j_Json.Get_String('regbegin_time'), 'YYYY-MM-DD HH24:MI:SS');
  d_Enddate    := To_Date(j_Json.Get_String('regend_time'), 'YYYY-MM-DD HH24:MI:SS');
  n_Ward_Id    := j_Json.Get_Number('ward_id');
  n_Show_Type  := j_Json.Get_Number('drugname_show_type');
  v_Method_Ids := j_Json.Get_String('drug_method_ids');
  v_Site_No    := j_Json.Get_String('site_no');

  For I In (Select 名称
            From 诊疗项目目录
            Where 类别 = 'E' And 操作类型 = '2' And 服务对象 In (2, 3) And (站点 = v_Site_No Or 站点 Is Null) And
                  ID In (Select Column_Value From Table(f_Str2list(v_Method_Ids)))
            Order By 编码) Loop
    v_Method_Names := v_Method_Names || ',' || i.名称;
  End Loop;
  If Not v_Method_Names Is Null Then
    v_Method_Names := Substr(v_Method_Names, 2);
  End If;

  v_Jtmp := Null;
  For r_Data In (Select /*+ Rule*/
                  a.药品id, d.编码, Nvl(d.名称, e.名称) As 名称, c.住院单位 As 单位, d.规格, d.产地,
                  Decode(d.类别, '7', Sum(a.填写数量 / Nvl(c.住院包装, 1) * Nvl(a.付数, 1)), Sum(a.填写数量 / Nvl(c.住院包装, 1))) As 数量,
                  d.类别, c.住院包装
                 From 药品收发记录 A, 药品规格 C, 收费项目目录 D, 收费项目别名 E
                 Where a.单据 = 9 And a.药品id = c.药品id And c.药品id = d.Id And Mod(a.记录状态, 3) = 1 And a.审核人 Is Null And
                       a.Id = e.收费细目id(+) And e.码类(+) = 1 And e.性质(+) = n_Show_Type And
                       Decode(v_Method_Ids, '', '', Nvl(a.用法, 'Null')) Not In
                       (Select Column_Value From Table(f_Str2list(v_Method_Names))) And a.填制日期 Between d_Begindate And
                       d_Enddate And a.库房id = n_Deptid And a.病人病区id = n_Ward_Id
                 Group By a.药品id, d.编码, Nvl(d.名称, e.名称), c.住院单位, d.规格, d.产地, d.类别, c.住院包装
                 Order By 编码) Loop
  
    v_Jtmp := v_Jtmp || ',';
    Zljsonputvalue(v_Jtmp, 'drug_id', r_Data.药品id, 1, 1);
    Zljsonputvalue(v_Jtmp, 'rcpdtl_code', r_Data.编码, 0);
    Zljsonputvalue(v_Jtmp, 'drug_name', r_Data.名称, 0);
    Zljsonputvalue(v_Jtmp, 'unit', r_Data.单位, 0);
    Zljsonputvalue(v_Jtmp, 'spec', r_Data.规格, 0);
    Zljsonputvalue(v_Jtmp, 'place_name', r_Data.产地, 0);
    Zljsonputvalue(v_Jtmp, 'quantity', r_Data.数量, 1);
    Zljsonputvalue(v_Jtmp, 'category', r_Data.类别, 0);
    Zljsonputvalue(v_Jtmp, 'inpack', r_Data.住院包装, 1, 2);
  
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
    Json_Out := '{"output":{"code":1,"message":"成功","data":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","data":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Surplus;
/

Create Or Replace Procedure Zl_Drugsvr_Getrecipeaudit
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：查看处方审查结果
  --入参：Json_In:格式
  --  input
  --   pati_id              N   1   病人ID
  --   pvid                 N   1   患者就诊id     门诊：挂号ID   住院：主页ID
  --   pat_source           N   1   患者来源       病人来源:1-门诊;2-住院
  --出参: Json_Out,格式如下
  --  output
  --    code                 N   1 应答吗：0-失败；1-成功
  --    message              C   1 应答消息：失败时返回具体的错误信息
  --    item_list
  --     advice_id            N   1   医嘱ID
  --     recipe_audit_status  N   1   处方审查状态      0-待审；1-已审；2-超时免审(门诊药师一直未操作审方)； 11-已审被撤销；
  --     recipe_audit_result  N   1   处方审查结果      1-合格；2-不合格
  ---------------------------------------------------------------------------
  n_病人id Number(18);
  n_就诊id Number(18);
  n_来源   Number(2);
  v_Jtmp   Varchar2(32767);

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --解析入参
  j_Jsonin := PLJson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');
  n_就诊id := j_Json.Get_Number('pvid');
  n_来源   := j_Json.Get_Number('pat_source');

  If Nvl(n_病人id, 0) <> 0 Then
    If Nvl(n_来源, 0) = 1 Then
      For c_审查 In (Select i.医嘱id, j.状态 As 处方审查状态, j.审查结果 As 处方审查结果
                   From 处方审查记录 J, 处方审查明细 I
                   Where j.病人id = n_病人id And j.挂号id = n_就诊id And j.Id = i.审方id(+) And (i.最后提交 = 1 Or i.审方id Is Null)) Loop
      
        v_Jtmp := v_Jtmp || ',{"advice_id":' || c_审查.医嘱id;
        v_Jtmp := v_Jtmp || ',"recipe_audit_status":' || Nvl(c_审查.处方审查状态, 0);
        v_Jtmp := v_Jtmp || ',"recipe_audit_result":' || Nvl(c_审查.处方审查结果, 0);
        v_Jtmp := v_Jtmp || '}';
      End Loop;
    Else
      For c_审查 In (Select i.医嘱id, j.状态 As 处方审查状态, j.审查结果 As 处方审查结果
                   From 处方审查记录 J, 处方审查明细 I
                   Where j.病人id = n_病人id And j.主页id = n_就诊id And j.Id = i.审方id(+) And (i.最后提交 = 1 Or i.审方id Is Null)) Loop
      
        v_Jtmp := v_Jtmp || ',{"advice_id":' || c_审查.医嘱id;
        v_Jtmp := v_Jtmp || ',"recipe_audit_status":' || Nvl(c_审查.处方审查状态, 0);
        v_Jtmp := v_Jtmp || ',"recipe_audit_result":' || Nvl(c_审查.处方审查结果, 0);
        v_Jtmp := v_Jtmp || '}';
      
      End Loop;
    End If;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getrecipeaudit;
/

Create Or Replace Procedure Zl_Drugsvr_Retainedrecords
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  --功能：获取已经填写的药品留存记录
  --入参：JOSN格式
  --input
  --  ward_id                  N 1 部门ID
  --  dept_id                  N 1 库房ID
  --  drugname_show_type       N 1 药品名称显示方式 
  --出参：JSON
  --output
  --  code                     N 应答码：0-失败；1-成功
  --  message                  C 应答消息：失败时返回具体的错误信息
  --  data[]列表
  --     drug_id               N 1 药品id
  --     rcpdtl_code           C 1 编码
  --     drug_name             C 1 药品名称
  --     unit                  C 1 单位
  --     spec                  C 1 规格
  --     place_name            C 1 产地
  --     quantity              N 1 留存数量
  --     category              C 1 类别
  --     inpack                N 1 住院包装 
  n_Dept_Id Number(18);
  n_Ward_Id Number(18);
  n_Type    Number(3);

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;

  j_Jsonin Pljson;
  j_Json   Pljson;
Begin
  --解析入参
  j_Jsonin  := Pljson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  n_Ward_Id := j_Json.Get_Number('ward_id');
  n_Dept_Id := j_Json.Get_Number('dept_id');
  n_Type    := j_Json.Get_Number('drugname_show_type');

  If Nvl(n_Ward_Id, 0) <> 0 Then
    For r_Data In (Select a.药品id, c.编码, Nvl(d.名称, c.名称) As 名称, c.规格, c.产地, b.住院单位 As 单位, a.留存数量 / Nvl(b.住院包装, 1) As 留存数量,
                          c.类别, b.住院包装
                   From 药品留存计划 A, 药品规格 B, 收费项目目录 C, 收费项目别名 D
                   Where a.药品id = b.药品id And a.药品id = c.Id And a.部门id = n_Ward_Id And a.库房id = n_Dept_Id And a.状态 = 0 And
                         c.Id = d.收费细目id(+) And d.码类(+) = 1 And d.性质(+) = n_Type
                   Order By c.编码) Loop
    
      v_Jtmp := v_Jtmp || ',';
      Zljsonputvalue(v_Jtmp, 'drug_id', r_Data.药品id, 1, 1);
      Zljsonputvalue(v_Jtmp, 'rcpdtl_code', r_Data.编码, 0);
      Zljsonputvalue(v_Jtmp, 'drug_name', r_Data.名称, 0);
      Zljsonputvalue(v_Jtmp, 'unit', r_Data.单位, 0);
      Zljsonputvalue(v_Jtmp, 'spec', r_Data.规格, 0);
      Zljsonputvalue(v_Jtmp, 'place_name', r_Data.产地, 0);
      Zljsonputvalue(v_Jtmp, 'quantity', r_Data.留存数量, 1);
      Zljsonputvalue(v_Jtmp, 'category', r_Data.类别, 0);
      Zljsonputvalue(v_Jtmp, 'inpack', r_Data.住院包装, 1, 2);
    
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
    Json_Out := '{"output":{"code":1,"message":"成功","data":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","data":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Retainedrecords;
/

Create Or Replace Procedure Zl_Drugsvr_Getretainbyorder
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  --功能：获取已经填写的药品留存记录根据医嘱id进行查询
  --入参：JOSN格式
  --input
  --  order_ids                C 1 医嘱IDs，根据医嘱id获取药品明细，逗号拼串
  --出参：JSON
  --output
  --  code                     N 1 应答码：0-失败；1-成功
  --  message                  C 1 应答消息：失败时返回具体的错误信息 
  --  data[]列表
  --     order_id              N 1 医嘱id
  --     drug_id               N 1 药品id
  --     drug_content          N 1 剂量系数
  --     si_drug_packg_qunt    N 1 住院包装
  --     is_part               N 1 住院可否分零

  v_Order_Ids Varchar2(32767);
  v_Jtmp      Varchar2(32767);
  c_Jtmp      Clob;
  j_Jsonin    Pljson;
  j_Json      Pljson;
Begin
  --解析入参
  j_Jsonin    := Pljson(Json_In);
  j_Json      := j_Jsonin.Get_Pljson('input');
  v_Order_Ids := j_Json.Get_String('order_ids');
  For r_Data In (Select a.医嘱id, a.药品id, b.剂量系数, b.住院包装, b.住院可否分零
                 From 药品收发记录 A, 药品规格 B
                 Where a.药品id = b.药品id And
                       a.医嘱id In (Select /*+cardinality(x,10)*/
                                   x.Column_Value
                                  From Table(Cast(f_Num2list(v_Order_Ids) As Zltools.t_Numlist)) X)
                 Order By a.医嘱id) Loop
  
    v_Jtmp := v_Jtmp || ',';
    Zljsonputvalue(v_Jtmp, 'order_id', r_Data.医嘱id, 1, 1);
    Zljsonputvalue(v_Jtmp, 'drug_id', r_Data.药品id, 1);
    Zljsonputvalue(v_Jtmp, 'drug_content', r_Data.剂量系数, 1);
    Zljsonputvalue(v_Jtmp, 'si_drug_packg_qunt', r_Data.住院包装, 1);
    Zljsonputvalue(v_Jtmp, 'is_part', r_Data.住院可否分零, 1, 2);
  
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
    Json_Out := '{"output":{"code":1,"message":"成功","data":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","data":[' || c_Jtmp || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getretainbyorder;
/

CREATE OR REPLACE Procedure Zl_Drugsvr_Retainplan_Insert
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  --功能： 插入药品留存计划
  --入参:JSON格式
  --input
  --       ward_id              N 1 病区id
  --       dept_id              N 1 库房id
  --       operator_name        C 1 操作员姓名
  --       drug_list[]药品数量列表
  --           drug_id          N 1 药品id
  --           quantity         N 1 数量
  --出参:JSON格式
  --output
  --       code                 N 1 应答码：0-失败 1-成功
  --       message              C 1 成功和失败后返回的消息
  j_Jsonlist      Pljson_List;
  d_操作时间      Date;
  n_Dept_Id       Number(18); --库房id
  n_Warda_Id      Number(18); --病区id
  v_Operator_Name Varchar2(50);
  j_Jsonin        Pljson;
  j_Json          Pljson;
Begin
  --解析入参
  j_Jsonin        := Pljson(Json_In);
  j_Json          := j_Jsonin.Get_Pljson('input');
  n_Dept_Id       := j_Json.Get_Number('dept_id');
  n_Warda_Id      := j_Json.Get_Number('ward_id');
  v_Operator_Name := j_Json.Get_String('operator_name');
  j_Jsonlist      := j_Json.Get_Pljson_List('drug_list');
  Select Sysdate Into d_操作时间 From Dual;
  j_Json := Pljson();
  For I In 1 .. j_Jsonlist.Count Loop
    j_Json := Pljson(j_Jsonlist.Get(I));
    Insert Into 药品留存计划
      (部门id, 库房id, 药品id, 留存数量, 状态, 登记人, 登记时间)
    Values
      (n_Warda_Id, n_Dept_Id, j_Json.Get_Number('drug_id'), j_Json.Get_Number('quantity'), 0, v_Operator_Name, d_操作时间);
    j_Json := Pljson();
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugsvr_Retainplan_Insert;
/

Create Or Replace Procedure Zl_Drugsvr_Retainplan_Delete
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  --功能： 插入药品留存计划
  --入参:JSON格式
  --input
  --     ward_id            N 1 部门id
  --     dept_id            N 1 库房id
  --     drug_id            N 0 药品id
  --出参:JSON格式
  --output
  --     code               N 1 应答码：0-失败 1-成功
  --     message            C 1 成功和失败后返回的消息
  n_Dept_Id  Number;
  n_Warda_Id Number;
  n_Drug_Id  Number;

  j_Jsonin Pljson;
  j_Json   Pljson;
Begin
  --解析入参
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_Dept_Id  := j_Json.Get_Number('dept_id');
  n_Warda_Id := j_Json.Get_Number('ward_id');
  n_Drug_Id  := j_Json.Get_Number('drug_id');

  If n_Drug_Id Is Null Then
    Delete 药品留存计划 Where 部门id = n_Warda_Id And 库房id = n_Dept_Id And 状态 = 0 And 留存id Is Null;
  Else
    Delete 药品留存计划
    Where 部门id = n_Warda_Id And 库房id = n_Dept_Id And 药品id = n_Drug_Id And 状态 = 0 And 留存id Is Null;
  End If;

  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugsvr_Retainplan_Delete;
/

Create Or Replace Procedure Zl_Drugsvr_Merage
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  --功能：病人信息合并时，药品收发记录合并
  --入参：JSON格式
  --input
  ----retain_id 保留id
  ----merge_id 合并id
  ----pati_name C 姓名
  ----pati_sex C 性别
  ----pati_age C 年龄
  ----pati_borth_time C 出生日期
  ----pati_identity C 身份证号;
  ----item_list
  ------page_id_new N 新主页id 
  ------pati_id_befor N  原病人id 
  ------page_id_befor N  原主页id 
  --出参：JSON格式
  --output
  ----code N 应答码：0-失败；1-成功  
  ----message C 应答消息：失败时返回具体的错误信息 
  j_Json_List  Pljson_List;
  n_Retain_Id  Number;
  n_Merge_Id   Number;
  v_Name       Varchar2(100);
  v_Sex        Varchar2(10);
  v_Age        Varchar2(20);
  d_Borth_Time Date;
  v_Identity   Varchar2(20);
  n_Page_New   Number;
  n_Pati_Befor Number;
  n_Page_Befor Number;
  I            Number;

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --解析入参
  j_Jsonin     := PLJson(Json_In);
  j_Json       := j_Jsonin.Get_Pljson('input');
  n_Retain_Id  := j_Json.Get_Number('retain_id');
  n_Merge_Id   := j_Json.Get_Number('merge_id');
  v_Name       := j_Json.Get_String('pati_name');
  v_Sex        := j_Json.Get_String('pati_sex');
  v_Age        := j_Json.Get_String('pati_age');
  d_Borth_Time := To_Date(j_Json.Get_String('pati_borth_time'), 'YYYY-MM-DD HH24:MI:SS');
  v_Identity   := j_Json.Get_String('pati_identity');
  j_Json_List  := j_Json.Get_Pljson_List('item_list');

  For I In 1 .. j_Json_List.Count Loop
    j_Json       := PLJson();
    j_Json       := PLJson(j_Json_List.Get(I));
    n_Page_New   := j_Json.Get_Number('page_id_new');
    n_Pati_Befor := j_Json.Get_Number('pati_id_befor');
    n_Page_Befor := j_Json.Get_Number('page_id_befor');
  
    Update 未发药品记录
    Set 病人id = n_Retain_Id, 主页id = n_Page_New, 姓名 = v_Name
    Where 病人id = n_Pati_Befor And 主页id = n_Page_Befor;
  
    Update 药品收发记录
    Set 病人id = n_Retain_Id, 主页id = n_Page_New, 姓名 = v_Name, 性别 = v_Sex, 年龄 = v_Age, 出生日期 = d_Borth_Time,
        身份证号 = v_Identity
    Where 病人id = n_Pati_Befor And 主页id = n_Page_Befor;
  End Loop;

  Update 未发药品记录 Set 病人id = n_Retain_Id, 姓名 = v_Name Where 病人id = n_Merge_Id And 主页id Is Null;

  Update 药品收发记录
  Set 病人id = n_Retain_Id, 姓名 = v_Name, 性别 = v_Sex, 年龄 = v_Age, 出生日期 = d_Borth_Time, 身份证号 = v_Identity
  Where 病人id = n_Merge_Id And 主页id Is Null;

  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugsvr_Merage;
/

Create Or Replace Procedure Zl_Drugsvr_Delrecipebill
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：药品处方销帐(或退费)
  --入参：Json_In:格式
  -- input
  --     billtype                 N   1   单据类型:1 -收费处方发药  ;2- 记帐单处方发药
  --     rcp_no                   C   1   单据号,有该节点是按整个NO进行销帐（暂时只用由于新门诊系统接口）
  --     item_list[]更新数据列表[数组]销帐列表
  --          rcpdtl_id           N 1 处方明细id,目前传入的费用ID
  --          chargeoffs_num      N 1 销帐数量
  --          dispensing_ids      N 1 配药IDs :销帐审核时输液配置中心需要传递的记录id，以字符串传递，用逗号分割，如：1001,1002,1003
  --     pivas_list[]对于静配销帐需要自动审核的情况，在删除药品的同时处理静配的销帐数据,普通处方删除不用传这个列表
  --          pivas_ids           C 1 配药IDs，需要销帐并同时审核的配液ids串
  --          operator_name       C 1 操作员，
  --          operator_time       C 1 操作时间
  --          reason              C 1 销帐原因
  --      return_list[]自动退药列表
  --           audit_operator     C 1 审核人
  --           operator_time      C 1 操作时间
  --           rcpdtl_list[]退药列表信息
  --               rcpdtl_id      N 1 费用id
  --               re_quantity    N 0 退药数量（销售单，收费项目目录.计算单位），可不传此结点或者传空，表示全退，否则按指定数量退药

  --出参: Json_Out,格式如下
  -- output
  --    code                      N 1 应答吗：0-失败；1-成功
  --    message                   C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  n_处方明细id 药品收发记录.费用id%Type;
  n_销帐数量   药品收发记录.实际数量%Type;
  v_操作员姓名 Varchar2(4000);
  d_操作时间   Date;
  v_操作说明   Varchar2(4000);
  v_配液ids    Varchar2(32767);
  n_小数       Number(6);
  n_费用id     Number(18);
  n_数量       药品收发记录.实际数量%Type;
  n_退药数量   药品收发记录.实际数量%Type;
  n_操作状态   Number(3);
  n_性质       Number(1);
  v_No         药品收发记录.No%Type;

  v_Err Varchar2(255);
  Err_Custom Exception;

  j_Jsonin   Pljson;
  j_Json     Pljson;
  j_Jsonlist Pljson_List;
  Jl_Re      Pljson_List;
  j_Re       Pljson;
  j_Item     Pljson;
Begin
  --解析入参
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');

  --按NO销账，目前用于新门诊接口
  v_No := j_Json.Get_String('rcp_no');
  If v_No Is Not Null Then
    n_性质 := j_Json.Get_Number('billtype');
  
    For r_处方明细 In (Select 费用id, Sum(Nvl(付数, 1) * 实际数量) As 退药数量
                   From 药品收发记录
                   Where 单据 = Decode(n_性质, 1, 8, 2, 9, 10) And NO = v_No And 审核日期 Is Null
                   Group By 费用id
                   Order By 费用id) Loop
    
      --按费用id进行销账
      Zl_药品收发记录_销售退费_s(r_处方明细.费用id, r_处方明细.退药数量, Null, 1);
    End Loop;
  End If;

  --自动退药 
  j_Jsonlist := j_Json.Get_Pljson_List('return_list');
  If j_Jsonlist Is Not Null Then
    Select 精度 Into n_小数 From 药品卫材精度 Where 类别 = 1 And 内容 = 4 And 单位 = 5;
    For I In 1 .. j_Jsonlist.Count Loop
      j_Item       := Pljson(); --清空中间变量 
      j_Item       := Pljson(j_Jsonlist.Get(I));
      v_操作员姓名 := j_Item.Get_String('audit_operator');
      d_操作时间   := To_Date(j_Item.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');
      Jl_Re        := j_Item.Get_Pljson_List('rcpdtl_list');
      For K In 1 .. Jl_Re.Count Loop
        j_Re     := Pljson(Jl_Re.Get(K));
        n_费用id := j_Re.Get_Number('rcpdtl_id');
        n_数量   := j_Re.Get_Number('re_quantity');
        j_Re     := Pljson();
        If n_数量 Is Not Null Then
          n_退药数量 := n_数量;
        End If;
        --分解退药数量
        For r_处方明细 In (Select a.库房id, a.Id, Nvl(a.付数, 1) * a.实际数量 As 数量
                       From 药品收发记录 A
                       Where a.单据 In (8, 9, 10) And (a.记录状态 = 1 Or Mod(a.记录状态, 3) = 0) And 审核人 Is Not Null And
                             a.费用id = n_费用id
                       Order By a.库房id, a.药品id, a.批次) Loop
          If n_数量 Is Null Then
            --传入的数量为空表示全退
            --部门退药（处方明细）
            Zl_药品收发记录_部门退药_s(r_处方明细.Id, v_操作员姓名, d_操作时间, Null, Null, Null, Null, Null, Null, n_小数);
          Else
            If n_退药数量 > 0 Then
              If n_退药数量 > r_处方明细.数量 Then
                --部门退药（处方明细）
                Zl_药品收发记录_部门退药_s(r_处方明细.Id, v_操作员姓名, d_操作时间, Null, Null, Null, r_处方明细.数量, Null, Null, n_小数);
                n_退药数量 := n_退药数量 - r_处方明细.数量;
              Else
                --部门退药（处方明细）
                Zl_药品收发记录_部门退药_s(r_处方明细.Id, v_操作员姓名, d_操作时间, Null, Null, Null, n_退药数量, Null, Null, n_小数);
                n_退药数量 := 0;
              End If;
            End If;
          End If;
        End Loop;
      End Loop;
      Jl_Re := Pljson_List();
    End Loop;
  End If;

  --删除单据
  j_Jsonlist := Pljson_List();
  j_Jsonlist := j_Json.Get_Pljson_List('item_list');
  If j_Jsonlist Is Not Null Then
    For I In 1 .. j_Jsonlist.Count Loop
      j_Item       := Pljson();
      j_Item       := Pljson(j_Jsonlist.Get(I));
      n_处方明细id := j_Item.Get_Number('rcpdtl_id');
      n_销帐数量   := j_Item.Get_Number('chargeoffs_num');
      v_配液ids    := j_Item.Get_String('dispensing_ids');
      If n_处方明细id Is Null Then
        v_Err := '传入节点【rcpdtl_id】错误，请检查！';
        Raise Err_Custom;
      End If;
      If n_销帐数量 Is Null Then
        v_Err := '传入节点【chargeoffs_num】错误，请检查！';
        Raise Err_Custom;
      End If;
      Zl_药品收发记录_销售退费_s(n_处方明细id, n_销帐数量, v_配液ids, 1);
    End Loop;
  End If;

  --静配相关处理
  j_Jsonlist := Pljson_List();
  j_Jsonlist := j_Json.Get_Pljson_List('pivas_list');
  If j_Jsonlist Is Not Null Then
    For I In 1 .. j_Jsonlist.Count Loop
      j_Item       := Pljson();
      j_Item       := Pljson(j_Jsonlist.Get(I));
      v_操作员姓名 := j_Item.Get_String('operator_name');
      d_操作时间   := To_Date(j_Item.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');
      v_操作说明   := j_Item.Get_String('reason');
      v_配液ids    := j_Item.Get_String('pivas_ids');
      n_操作状态   := j_Item.Get_Number('operator_status');
      Zl_输液配药记录_销帐更新状态_s(v_操作员姓名, d_操作时间, v_操作说明, v_配液ids, n_操作状态);
    End Loop;
  End If;
  Json_Out := Zljsonout('成功', 1);
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugsvr_Delrecipebill;
/

Create Or Replace Procedure Zl_Drugsvr_Sendquery
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  -----------------------------------------------------------------------------------------------------------------------
  --功能：药疗收发查询
  --入参：JSON格式
  --  input
  --      type_query              N 0 查询类型:0-发药明细清单 1-发药汇总清单 2-退药明细清单 3-退药汇总清单[由费用域服务直接获取] 4-退发汇总清单 
  --      dept_id                 N 1 库房id
  --      otherdept_id            N 0 对方部门id,领药部门id
  --      verify_begin_time       C 0 审核起始日期
  --      verify_end_time         C 1 审核起始日期 
  --      rcp_no                  C 1 NO，药品收发记录.NO
  --      group_no                N 1 汇总发药号
  --      record_state            C 1 记录状态,两位的数字拼串，01-查询已发药，10-只查询未发药，其它-不区分同时查询已发药和未发药
  --      effective_time          N 1 期效 0-长嘱，1-临嘱，2-长嘱和临嘱
  --      ward_id                 N 1 病区ID
  --      usages                  C 1 用法，给药途径 名称拼串，逗号分割
  --      drugname_show_type      N 0 药品名称显示方式，可不传缺省为1;对应于【收费项目别名.性质】：1-正名(对应项目中的名称);2-英文名;3-商品名;9-其他别名；其中西药存在英文名、商品名，卫材存在商品名，其他类别只有正名和别名
  --      pati_ids                C 0 病人ID逗号拼串，如果不区分病人，传空
  --      rcpdtl_ids              C 0 退药查询时根把  汇总发药号 获取费用明细id逗号拼串
  --      rcp_nos                 C 0 单据号，逗号拼串，当界面按发送时间查询时传入
  --      return_list[]退药信息列表，说明：type_query=2查询时传入（quantity+rcpdtl_id+serial_num）；type_query=(3、4)查询时传入（drug_id+quantity+re_money）
  --         drug_id              N 1 药品id
  --         quantity             N 1 退药数量
  --         re_money             N 1 金额
  --         rcpdtl_id            N 0 费用id,处方明细 
  --         serial_num           N 0 序号，唯一标识一次申请

  --出参
  --output
  --  code                        N 1 应答码：0-失败 1-成功
  --  message                     C 1 失败后需要返回的错误信息
  --  data[]  列表信息
  --      record_state            N 1 记录状态
  --      rcpdtl_id               N 1 费用id
  --      rcp_no                  C 1 NO，药品收发记录.no（病人医嘱发送.no）
  --      order_number            N 1 序号
  --      drugstore_name          C 1 药房名
  --      advice_dept_name        C 1 开嘱科室
  --      rcp_info                C 1 药品信息
  --      in_unit                 C 1 住院单位
  --      dosage_unit             C 1 剂量单位
  --      quantity                N 1 数量
  --      uint_price              N 1 单价
  --      effective               N 1 效期
  --      singular_quantity       N 1 单量
  --      money                   N 1 金额
  --      category                N 1 类别
  --      frequency               N 1 频次
  --      usage                   C 1 用法
  --      payment                 N 1 付数
  --      advice_id               N 1 医嘱ID
  --      pati_name               C 1 姓名
  --      pati_id                 N 1 病人id
  --      page_id                 N 1 主页id
  --      drug_code               C 1 药品编码   n_type=1、n_type=3、n_type=4时才有这个值 
  --      back_number             N 1 退药数 n_type=4时才有这个值
  --      reality_number          N 1 实发数 n_type=4时才有这个值
  -----------------------------------------------------------------------------------------------------------------------

  n_Type           Number(6);
  n_库房id         Number(18);
  n_对方部门id     Number(18);
  n_病区id         Number(18);
  d_审核起始日期   Date;
  d_审核结束日期   Date;
  v_No             Varchar2(30);
  n_汇总发药号     Number(18);
  v_记录状态       Varchar2(30);
  n_效期           Number(3);
  v_用法s          Varchar2(3980);
  v_病人ids        Varchar2(3980);
  n_Showtype       Number(3);
  d_退药申请时间起 Date;
  d_退药申请时间止 Date;
  l_退药信息       t_Strlist2 := t_Strlist2(); --有两种格式，一种是明细：费用id+退药数量+序号 ，一种是汇总：药品id+退药数量+金额
  n_费用id         Number(18);
  v_Nos            Varchar2(32767);
  c_No             t_Strlist := t_Strlist();

  Cursor c_List_Type Is
    Select a.Id 状态, a.No, a.序号, a.摘要 药房, a.摘要 开嘱科室, a.姓名, a.摘要 药品信息, a.单量 数量, a.摘要 住院单位, a.单量 单价, a.单量 金额, a.摘要 期效, a.单量,
           a.摘要 剂量单位, a.频次, a.用法, a.付数, a.摘要 类别, a.病人id, a.主页id, a.病人id 退药序号
    From 药品收发记录 A
    Where 0 = 1;
  r_Detail c_List_Type%RowType;

  Cursor c_Group_Type Is
    Select a.摘要 药品编码, a.摘要 药品信息, a.摘要 住院单位, a.单量 数量, a.单量 退药数, a.单量 实发数, a.单量 金额
    From 药品收发记录 A
    Where 0 = 1;
  r_Grp c_Group_Type%RowType;

  v_Jtmp     Varchar2(32767); --不要随便使用此变量
  c_Jtmp     Clob; --不要随便使用此变量
  j_Jsonlist Pljson_List;
  j_Tmp      Pljson;
  j_Jsonin   Pljson;
  j_Json     Pljson;

  Procedure Get出参拼串明细 As
  Begin
    v_Jtmp := v_Jtmp || ',';
    Zljsonputvalue(v_Jtmp, 'record_state', r_Detail.状态, 1, 1);
    Zljsonputvalue(v_Jtmp, 'rcp_no', r_Detail.No, 0);
    Zljsonputvalue(v_Jtmp, 'order_number', r_Detail.序号, 1);
    Zljsonputvalue(v_Jtmp, 'drugstore_name', r_Detail.药房, 0);
    Zljsonputvalue(v_Jtmp, 'advice_dept_name', r_Detail.开嘱科室, 0);
    Zljsonputvalue(v_Jtmp, 'rcp_info', r_Detail.药品信息, 0);
    Zljsonputvalue(v_Jtmp, 'in_unit', r_Detail.住院单位, 0);
    Zljsonputvalue(v_Jtmp, 'dosage_unit', r_Detail.剂量单位, 0);
    Zljsonputvalue(v_Jtmp, 'uint_price', r_Detail.单价, 1);
    Zljsonputvalue(v_Jtmp, 'money', r_Detail.金额, 1);
    Zljsonputvalue(v_Jtmp, 'effective', r_Detail.期效, 0);
    Zljsonputvalue(v_Jtmp, 'singular_quantity', r_Detail.单量, 1);
    Zljsonputvalue(v_Jtmp, 'quantity', r_Detail.数量, 1);
    Zljsonputvalue(v_Jtmp, 'category', r_Detail.类别, 0);
    Zljsonputvalue(v_Jtmp, 'frequency', r_Detail.频次, 0);
    Zljsonputvalue(v_Jtmp, 'usage', r_Detail.用法, 0);
    Zljsonputvalue(v_Jtmp, 'payment', r_Detail.付数, 1);
    Zljsonputvalue(v_Jtmp, 'pati_name', r_Detail.姓名, 0);
    Zljsonputvalue(v_Jtmp, 'pati_id', r_Detail.病人id, 1);
    Zljsonputvalue(v_Jtmp, 'rcpdtl_id', r_Detail.退药序号, 1);
    Zljsonputvalue(v_Jtmp, 'page_id', r_Detail.主页id, 1, 2);
  
    If Length(v_Jtmp) > 30000 Then
      If c_Jtmp Is Null Then
        c_Jtmp := Substr(v_Jtmp, 2);
      Else
        c_Jtmp := c_Jtmp || v_Jtmp;
      End If;
      v_Jtmp := Null;
    End If;
  End;

  Procedure Get出参拼串汇总 As
  Begin
    v_Jtmp := v_Jtmp || ',';
    Zljsonputvalue(v_Jtmp, 'rcp_info', r_Grp.药品信息, 0, 1);
    Zljsonputvalue(v_Jtmp, 'drug_code', r_Grp.药品编码, 0);
    Zljsonputvalue(v_Jtmp, 'in_unit', r_Grp.住院单位, 0);
    Zljsonputvalue(v_Jtmp, 'quantity', r_Grp.数量, 1);
    Zljsonputvalue(v_Jtmp, 'money', r_Grp.金额, 1);
    Zljsonputvalue(v_Jtmp, 'back_number', r_Grp.退药数, 1);
    Zljsonputvalue(v_Jtmp, 'reality_number', r_Grp.实发数, 1, 2);
  
    If Length(v_Jtmp) > 30000 Then
      If c_Jtmp Is Null Then
        c_Jtmp := Substr(v_Jtmp, 2);
      Else
        c_Jtmp := c_Jtmp || v_Jtmp;
      End If;
      v_Jtmp := Null;
    End If;
  End;

Begin
  --解析入参
  j_Jsonin     := Pljson(Json_In);
  j_Json       := j_Jsonin.Get_Pljson('input');
  n_Type       := Nvl(j_Json.Get_Number('type_query'), 0);
  n_库房id     := j_Json.Get_Number('dept_id');
  n_对方部门id := j_Json.Get_Number('otherdept_id');
  n_病区id     := j_Json.Get_Number('ward_id');
  Select zl_GetSysParameter('药品名称显示', Null, 100) Into n_Showtype From Dual;
  If Nvl(n_Showtype, 0) = 0 Then
    n_Showtype := 1;
  End If;
  d_审核起始日期 := To_Date(j_Json.Get_String('verify_begin_time'), 'YYYY-MM-DD HH24:MI:SS');
  d_审核结束日期 := To_Date(j_Json.Get_String('verify_end_time'), 'YYYY-MM-DD HH24:MI:SS');
  v_No           := j_Json.Get_String('rcp_no');
  n_汇总发药号   := j_Json.Get_String('group_no');
  v_记录状态     := j_Json.Get_String('record_state');
  If Not (v_记录状态 = '01' Or v_记录状态 = '10') Then
    v_记录状态 := Null;
  End If;
  n_效期    := j_Json.Get_Number('effective_time');
  v_用法s   := j_Json.Get_String('usages');
  v_病人ids := j_Json.Get_String('pati_ids');
  v_Nos     := j_Json.Get_String('rcp_nos');

  If v_Nos Is Not Null Then
    v_Nos := v_Nos || ',';
    While v_Nos Is Not Null Loop
      c_No.Extend;
      c_No(c_No.Count) := Substr(v_Nos, 1, Instr(v_Nos, ',') - 1);
      v_Nos := Substr(v_Nos, Instr(v_Nos, ',') + 1);
    End Loop;
  End If;

  j_Jsonlist := j_Json.Get_Pljson_List('return_list');
  If j_Jsonlist Is Not Null Then
    v_Jtmp := Null;
    For K In 1 .. j_Jsonlist.Count Loop
      j_Tmp    := Pljson(j_Jsonlist.Get(K));
      n_费用id := j_Tmp.Get_Number('rcpdtl_id');
      l_退药信息.Extend();
      If Nvl(n_费用id, 0) = 0 Then
        l_退药信息(l_退药信息.Count) := t_Strobj2(j_Tmp.Get_Number('drug_id'),
                                          j_Tmp.Get_Number('quantity') || '_' || j_Tmp.Get_Number('re_money'));
      Else
        l_退药信息(l_退药信息.Count) := t_Strobj2(n_费用id, j_Tmp.Get_Number('quantity') || '_' || j_Tmp.Get_Number('serial_num'));
      End If;
      j_Tmp := Pljson();
    End Loop;
  End If;

  -----------------------------------------------
  --以发药时间为准的查询,用 审核日期 为主索引
  If n_Type = 0 Then
    --发药明细清单 tbcQuery.Selected.Index = 0
    If d_审核起始日期 Is Not Null Then
      For R In (Select a.状态, a.No, a.序号, i.名称 As 药房, h.名称 As 开嘱科室, a.姓名,
                       Nvl(x.名称, f.名称) || Decode(f.产地, Null, Null, '(' || f.产地 || ')') ||
                        Decode(f.规格, Null, Null, ' ' || f.规格) As 药品信息, a.数量 / Nvl(e.住院包装, 1) As 数量, e.住院单位,
                       a.单价 * Nvl(e.住院包装, 1) As 单价, a.金额, a.期效, a.单量, g.计算单位 As 剂量单位, a.频次, a.用法, a.付数, g.类别, a.病人id,
                       a.主页id, 0 退药id
                From (Select a.状态, a.No, a.序号, a.库房id, a.对方部门id, a.姓名, a.住院号, a.床号, a.药品id, a.单价, a.期效, a.单量, a.频次, a.用法,
                              a.时间, a.人员, a.付数, a.病人id, a.主页id, Sum(a.数量) As 数量, Sum(a.金额) As 金额
                       From (Select b.状态, c.No, c.序号, c.库房id, c.对方部门id, c.姓名, '需补充' 住院号, '需补充' 床号, c.药品id, b.数量, c.零售价 As 单价,
                                     b.金额, Decode(Nvl(Substr(c.扣率, 1, 1), 0), 0, '长嘱', '临嘱') As 期效, c.单量, c.频次, c.用法,
                                     'A.发送时间' As 时间, 'A.发送人' As 人员, c.付数, c.病人id, c.主页id
                              From (Select Decode(a.审核人, Null, 0, 1) As 状态, a.No, a.序号, Sum(a.填写数量 * a.付数) As 数量,
                                            Sum(a.零售金额) As 金额
                                     From 药品收发记录 A
                                     Where a.审核日期 Between d_审核起始日期 And d_审核结束日期 And a.单据 + 0 = 9 And a.医嘱id Is Not Null And
                                           (a.对方部门id = n_对方部门id Or Nvl(n_对方部门id, 0) = 0) And
                                           (a.库房id = n_库房id Or Nvl(n_库房id, 0) = 0) And
                                           (a.No = v_No Or Nvl(v_No, 'NONE') = 'NONE') And
                                           (a.汇总发药号 = n_汇总发药号 Or Nvl(n_汇总发药号, 0) = 0) And
                                           (Nvl(Substr(a.扣率, 1, 1), 0) = n_效期 Or Nvl(n_效期, 2) = 2) And
                                           (v_记录状态 = '01' And a.审核人 Is Not Null Or
                                           v_记录状态 = '10' And Mod(a.记录状态, 3) = 1 And a.审核人 Is Null Or
                                           Nvl(v_记录状态, 'NONE') = 'NONE') And
                                           (a.用法 In (Select /*+cardinality(x,10)*/
                                                      x.Column_Value
                                                     From Table(f_Str2list(v_用法s)) X) Or Nvl(v_用法s, 'NONE') = 'NONE') And
                                           (a.病人id + 0 In (Select /*+cardinality(x,10)*/
                                                            x.Column_Value
                                                           From Table(f_Str2list(v_病人ids)) X) Or
                                           Nvl(v_病人ids, 'NONE') = 'NONE')
                                     Group By Decode(a.审核人, Null, 0, 1), a.No, a.序号
                                     Having Nvl(Sum(a.填写数量), 0) <> 0 Or Nvl(Sum(a.零售金额), 0) <> 0) B, 药品收发记录 C
                              Where b.No = c.No And b.序号 = c.序号 And (c.记录状态 = 1 Or Mod(c.记录状态, 3) = 0) And
                                    (c.病人病区id = n_病区id Or Nvl(n_病区id, 0) = 0)) A
                       Group By a.状态, a.No, a.序号, a.库房id, a.对方部门id, a.姓名, a.住院号, a.床号, a.药品id, a.单价, a.期效, a.单量, a.频次, a.用法,
                                a.时间, a.人员, a.付数, a.病人id, a.主页id) A, 药品规格 E, 收费项目目录 F, 诊疗项目目录 G, 部门表 H, 部门表 I, 收费项目别名 X
                Where a.药品id = e.药品id And a.药品id = f.Id And e.药名id = g.Id And a.对方部门id = h.Id And a.库房id = i.Id And
                      f.Id = x.收费细目id(+) And x.码类(+) = 1 And x.性质(+) = n_Showtype
                Order By a.No, a.序号) Loop
        --程序外部还需要对 床号 进行排序，外部追加伪列表
        r_Detail := R;
        Get出参拼串明细;
      End Loop;
    Else
      For R In (Select a.状态, a.No, a.序号, i.名称 As 药房, h.名称 As 开嘱科室, a.姓名,
                       Nvl(x.名称, f.名称) || Decode(f.产地, Null, Null, '(' || f.产地 || ')') ||
                        Decode(f.规格, Null, Null, ' ' || f.规格) As 药品信息, a.数量 / Nvl(e.住院包装, 1) As 数量, e.住院单位,
                       a.单价 * Nvl(e.住院包装, 1) As 单价, a.金额, a.期效, a.单量, g.计算单位 As 剂量单位, a.频次, a.用法, a.付数, g.类别, a.病人id,
                       a.主页id, 0 退药id
                From (Select a.状态, a.No, a.序号, a.库房id, a.对方部门id, a.姓名, a.住院号, a.床号, a.药品id, a.单价, a.期效, a.单量, a.频次, a.用法,
                              a.时间, a.人员, a.付数, a.病人id, a.主页id, Sum(a.数量) As 数量, Sum(a.金额) As 金额
                       From (Select b.状态, c.No, c.序号, c.库房id, c.对方部门id, c.姓名, '需补充' 住院号, '需补充' 床号, c.药品id, b.数量, c.零售价 As 单价,
                                     b.金额, Decode(Nvl(Substr(c.扣率, 1, 1), 0), 0, '长嘱', '临嘱') As 期效, c.单量, c.频次, c.用法,
                                     'A.发送时间' As 时间, 'A.发送人' As 人员, c.付数, c.病人id, c.主页id
                              From (Select Decode(a.审核人, Null, 0, 1) As 状态, a.No, a.序号, Sum(a.填写数量 * a.付数) As 数量,
                                            Sum(a.零售金额) As 金额
                                     From 药品收发记录 A
                                     Where a.No In (Select /*+cardinality(x,10)*/
                                                     Column_Value
                                                    From Table(c_No) X) And a.单据 + 0 = 9 And a.医嘱id Is Not Null And
                                           (a.对方部门id = n_对方部门id Or Nvl(n_对方部门id, 0) = 0) And
                                           (a.库房id = n_库房id Or Nvl(n_库房id, 0) = 0) And
                                           (a.No = v_No Or Nvl(v_No, 'NONE') = 'NONE') And
                                           (a.汇总发药号 = n_汇总发药号 Or Nvl(n_汇总发药号, 0) = 0) And
                                           (Nvl(Substr(a.扣率, 1, 1), 0) = n_效期 Or Nvl(n_效期, 2) = 2) And
                                           (v_记录状态 = '01' And a.审核人 Is Not Null Or
                                           v_记录状态 = '10' And Mod(a.记录状态, 3) = 1 And a.审核人 Is Null Or
                                           Nvl(v_记录状态, 'NONE') = 'NONE') And
                                           (a.用法 In (Select /*+cardinality(x,10)*/
                                                      x.Column_Value
                                                     From Table(f_Str2list(v_用法s)) X) Or Nvl(v_用法s, 'NONE') = 'NONE') And
                                           (a.病人id + 0 In (Select /*+cardinality(x,10)*/
                                                            x.Column_Value
                                                           From Table(f_Str2list(v_病人ids)) X) Or
                                           Nvl(v_病人ids, 'NONE') = 'NONE')
                                     Group By Decode(a.审核人, Null, 0, 1), a.No, a.序号
                                     Having Nvl(Sum(a.填写数量), 0) <> 0 Or Nvl(Sum(a.零售金额), 0) <> 0) B, 药品收发记录 C
                              Where b.No = c.No And b.序号 = c.序号 And (c.记录状态 = 1 Or Mod(c.记录状态, 3) = 0) And
                                    (c.病人病区id = n_病区id Or Nvl(n_病区id, 0) = 0)) A
                       Group By a.状态, a.No, a.序号, a.库房id, a.对方部门id, a.姓名, a.住院号, a.床号, a.药品id, a.单价, a.期效, a.单量, a.频次, a.用法,
                                a.时间, a.人员, a.付数, a.病人id, a.主页id) A, 药品规格 E, 收费项目目录 F, 诊疗项目目录 G, 部门表 H, 部门表 I, 收费项目别名 X
                Where a.药品id = e.药品id And a.药品id = f.Id And e.药名id = g.Id And a.对方部门id = h.Id And a.库房id = i.Id And
                      f.Id = x.收费细目id(+) And x.码类(+) = 1 And x.性质(+) = n_Showtype
                Order By a.No, a.序号) Loop
        --程序外部还需要对 床号 进行排序，外部追加伪列表
        r_Detail := R;
        Get出参拼串明细;
      End Loop;
    End If;
  
  Elsif n_Type = 1 Then
    --发药汇总清单 tbcQuery.Selected.Index = 1    
    If d_审核起始日期 Is Not Null Then
      For R In (Select a.药品编码,
                       Nvl(b.名称, a.名称) || Decode(a.产地, Null, Null, '(' || a.产地 || ')') ||
                        Decode(a.规格, Null, Null, ' ' || a.规格) As 药品信息, a.住院单位, a.数量, 0 退药数, 0 实发数, a.金额
                From (Select b.药品id, c.编码 As 药品编码, c.名称, c.产地, c.规格, b.住院单位, Sum(a.数量 / Nvl(b.住院包装, 1)) As 数量,
                              Sum(a.金额) As 金额
                       From (Select b.状态, c.No, c.序号, c.库房id, c.对方部门id, c.姓名, '需补充' 住院号, '需补充' 床号, c.药品id, b.数量, c.零售价 As 单价,
                                     b.金额, Decode(Nvl(Substr(c.扣率, 1, 1), 0), 0, '长嘱', '临嘱') As 期效, c.单量, c.频次, c.用法,
                                     '发送时间需补充' As 时间, '发送人需补充' As 人员, c.付数, c.病人id, c.主页id
                              From (Select Decode(a.审核人, Null, 0, 1) As 状态, a.No, a.序号, Sum(a.填写数量 * a.付数) As 数量,
                                            Sum(a.零售金额) As 金额
                                     From 药品收发记录 A
                                     Where a.审核日期 Between d_审核起始日期 And d_审核结束日期 And a.单据 + 0 = 9 And a.医嘱id Is Not Null And
                                           (a.对方部门id = n_对方部门id Or Nvl(n_对方部门id, 0) = 0) And
                                           (a.库房id = n_库房id Or Nvl(n_库房id, 0) = 0) And
                                           (a.No = v_No Or Nvl(v_No, 'NONE') = 'NONE') And
                                           (a.汇总发药号 = n_汇总发药号 Or Nvl(n_汇总发药号, 0) = 0) And
                                           (Nvl(Substr(a.扣率, 1, 1), 0) = n_效期 Or Nvl(n_效期, 2) = 2) And
                                           (v_记录状态 = '01' And a.审核人 Is Not Null Or
                                           v_记录状态 = '10' And Mod(a.记录状态, 3) = 1 And a.审核人 Is Null Or
                                           Nvl(v_记录状态, 'NONE') = 'NONE') And
                                           (a.用法 In (Select /*+cardinality(x,10)*/
                                                      x.Column_Value
                                                     From Table(f_Str2list(v_用法s)) X) Or Nvl(v_用法s, 'NONE') = 'NONE') And
                                           (a.病人id + 0 In (Select /*+cardinality(x,10)*/
                                                            x.Column_Value
                                                           From Table(f_Str2list(v_病人ids)) X) Or
                                           Nvl(v_病人ids, 'NONE') = 'NONE')
                                     Group By Decode(a.审核人, Null, 0, 1), a.No, a.序号
                                     Having Nvl(Sum(a.填写数量), 0) <> 0 Or Nvl(Sum(a.零售金额), 0) <> 0) B, 药品收发记录 C
                              Where b.No = c.No And b.序号 = c.序号 And (c.记录状态 = 1 Or Mod(c.记录状态, 3) = 0) And
                                    (c.病人病区id = n_病区id Or Nvl(n_病区id, 0) = 0)) A, 药品规格 B, 收费项目目录 C
                       Where a.药品id = b.药品id And a.药品id = c.Id
                       Group By b.药品id, c.编码, c.名称, c.产地, c.规格, b.住院单位
                       Having Sum(a.数量 / Nvl(b.住院包装, 1)) <> 0 Or Sum(a.金额) <> 0) A, 收费项目别名 B
                Where a.药品id = b.收费细目id(+) And b.码类(+) = 1 And b.性质(+) = n_Showtype
                Order By a.药品编码) Loop
        r_Grp := R;
        Get出参拼串汇总;
      End Loop;
    Else
      For R In (Select a.药品编码,
                       Nvl(b.名称, a.名称) || Decode(a.产地, Null, Null, '(' || a.产地 || ')') ||
                        Decode(a.规格, Null, Null, ' ' || a.规格) As 药品信息, a.住院单位, a.数量, 0 退药数, 0 实发数, a.金额
                From (Select b.药品id, c.编码 As 药品编码, c.名称, c.产地, c.规格, b.住院单位, Sum(a.数量 / Nvl(b.住院包装, 1)) As 数量,
                              Sum(a.金额) As 金额
                       From (Select b.状态, c.No, c.序号, c.库房id, c.对方部门id, c.姓名, '需补充' 住院号, '需补充' 床号, c.药品id, b.数量, c.零售价 As 单价,
                                     b.金额, Decode(Nvl(Substr(c.扣率, 1, 1), 0), 0, '长嘱', '临嘱') As 期效, c.单量, c.频次, c.用法,
                                     '发送时间需补充' As 时间, '发送人需补充' As 人员, c.付数, c.病人id, c.主页id
                              From (Select Decode(a.审核人, Null, 0, 1) As 状态, a.No, a.序号, Sum(a.填写数量 * a.付数) As 数量,
                                            Sum(a.零售金额) As 金额
                                     From 药品收发记录 A
                                     Where a.No In (Select /*+cardinality(x,10)*/
                                                     Column_Value
                                                    From Table(c_No) X) And a.单据 + 0 = 9 And a.医嘱id Is Not Null And
                                           (a.对方部门id = n_对方部门id Or Nvl(n_对方部门id, 0) = 0) And
                                           (a.库房id = n_库房id Or Nvl(n_库房id, 0) = 0) And
                                           (a.No = v_No Or Nvl(v_No, 'NONE') = 'NONE') And
                                           (a.汇总发药号 = n_汇总发药号 Or Nvl(n_汇总发药号, 0) = 0) And
                                           (Nvl(Substr(a.扣率, 1, 1), 0) = n_效期 Or Nvl(n_效期, 2) = 2) And
                                           (v_记录状态 = '01' And a.审核人 Is Not Null Or
                                           v_记录状态 = '10' And Mod(a.记录状态, 3) = 1 And a.审核人 Is Null Or
                                           Nvl(v_记录状态, 'NONE') = 'NONE') And
                                           (a.用法 In (Select /*+cardinality(x,10)*/
                                                      x.Column_Value
                                                     From Table(f_Str2list(v_用法s)) X) Or Nvl(v_用法s, 'NONE') = 'NONE') And
                                           (a.病人id + 0 In (Select /*+cardinality(x,10)*/
                                                            x.Column_Value
                                                           From Table(f_Str2list(v_病人ids)) X) Or
                                           Nvl(v_病人ids, 'NONE') = 'NONE')
                                     Group By Decode(a.审核人, Null, 0, 1), a.No, a.序号
                                     Having Nvl(Sum(a.填写数量), 0) <> 0 Or Nvl(Sum(a.零售金额), 0) <> 0) B, 药品收发记录 C
                              Where b.No = c.No And b.序号 = c.序号 And (c.记录状态 = 1 Or Mod(c.记录状态, 3) = 0) And
                                    (c.病人病区id = n_病区id Or Nvl(n_病区id, 0) = 0)) A, 药品规格 B, 收费项目目录 C
                       Where a.药品id = b.药品id And a.药品id = c.Id
                       Group By b.药品id, c.编码, c.名称, c.产地, c.规格, b.住院单位
                       Having Sum(a.数量 / Nvl(b.住院包装, 1)) <> 0 Or Sum(a.金额) <> 0) A, 收费项目别名 B
                Where a.药品id = b.收费细目id(+) And b.码类(+) = 1 And b.性质(+) = n_Showtype
                Order By a.药品编码) Loop
        r_Grp := R;
        Get出参拼串汇总;
      End Loop;
    End If;
  Elsif n_Type = 2 Then
    --tbcQuery.Selected.Index = 2 退药明细 
    For R In (Select a.状态, a.No, a.序号, i.名称 As 药房, h.名称 As 开嘱科室, a.姓名,
                     Nvl(x.名称, f.名称) || Decode(f.产地, Null, Null, '(' || f.产地 || ')') ||
                      Decode(f.规格, Null, Null, ' ' || f.规格) As 药品信息, a.数量 / Nvl(e.住院包装, 1) As 数量, e.住院单位,
                     a.单价 * Nvl(e.住院包装, 1) As 单价, a.金额, a.期效, a.单量, g.计算单位 As 剂量单位, a.频次, a.用法, a.付数, g.类别, a.病人id,
                     a.主页id, a.退费id
              From (Select Distinct -1 As 状态, d.No, d.序号, d.库房id, d.对方部门id, d.姓名, '需补充' 住院号, '需补充' 床号, d.药品id, a.数量,
                                     d.零售价 As 单价, a.数量 * d.零售价 金额, Decode(Nvl(Substr(d.扣率, 1, 1), 0), 0, '长嘱', '临嘱') 期效,
                                     d.单量, d.频次, d.用法, a.退药申请序号 As 时间, a.退药申请序号 As 人员, d.付数, d.病人id, d.主页id, a.退药申请序号 退费id
                     From (Select /*+cardinality(x,10)*/
                             To_Number(x.C1) 费用id, To_Number(Substr(x.C2, 1, Instr(x.C2, '_') - 1)) 数量,
                             To_Number(Substr(x.C2, Instr(x.C2, '_') + 1)) 退药申请序号
                            From Table(l_退药信息) X) A, 药品收发记录 D
                     Where a.费用id = d.费用id) A, 药品规格 E, 收费项目目录 F, 诊疗项目目录 G, 部门表 H, 部门表 I, 收费项目别名 X
              Where a.药品id = e.药品id And a.药品id = f.Id And e.药名id = g.Id And a.对方部门id = h.Id And a.库房id = i.Id And
                    f.Id = x.收费细目id(+) And x.码类(+) = 1 And x.性质(+) = n_Showtype
              Order By a.No, a.序号) Loop
      r_Detail := R;
      Get出参拼串明细;
    End Loop;
  Elsif n_Type = 3 Then
    ---tbcQuery.Selected.Index = 3 退药汇总 
    For R In (Select a.药品编码,
                     Nvl(b.名称, a.名称) || Decode(a.产地, Null, Null, '(' || a.产地 || ')') ||
                      Decode(a.规格, Null, Null, ' ' || a.规格) As 药品信息, a.住院单位, a.数量, 0 退药数, 0 实发数, a.金额
              From (Select b.药品id, c.编码 As 药品编码, c.名称, c.产地, c.规格, b.住院单位, Sum(a.数量 / Nvl(b.住院包装, 1)) As 数量,
                            Sum(a.金额) As 金额
                     From (Select a.费用id, a.数量, a.数量 * b.标准单价 金额, b.收费细目id 药品id
                            From 病人费用销帐 A, 住院费用记录 B
                            Where a.费用id = b.Id And a.申请时间 Between d_退药申请时间起 And d_退药申请时间止 And b.医嘱序号 Is Not Null And
                                  a.审核部门id = n_库房id And Nvl(a.状态, 0) = 0 And b.收费类别 In ('5', '6', '7') And
                                  (b.领药部门id = n_对方部门id Or Nvl(n_对方部门id, 0) = 0) And a.申请部门id = n_病区id And
                                  (b.病人id + 0 In (Select /*+cardinality(x,10)*/
                                                   x.Column_Value
                                                  From Table(f_Str2list(v_病人ids)) X) Or Nvl(v_病人ids, 'NONE') = 'NONE') And
                                  (b.医嘱期效 = n_效期 Or Nvl(n_效期, 2) = 2)) A, 药品规格 B, 收费项目目录 C
                     Where a.药品id = b.药品id And a.药品id = c.Id
                     Group By b.药品id, c.编码, c.名称, c.产地, c.规格, b.住院单位
                     Having Sum(a.数量 / Nvl(b.住院包装, 1)) <> 0 Or Sum(a.金额) <> 0) A, 收费项目别名 B
              Where a.药品id = b.收费细目id(+) And b.码类(+) = 1 And b.性质(+) = n_Showtype
              Order By a.药品编码
              -- 汇总发药号，转换成费用ID拼串再传到费用那边去，
              -- 退药查询的时候让 给药途径过滤条件失效，在不冗余的情况下只有这个样子了
              ) Loop
      r_Grp := R;
      Get出参拼串汇总;
    End Loop;
  Elsif n_Type = 4 Then
    --tbcQuery.Selected.Index = 4 发退药汇总页卡
    For R In (Select a.药品编码,
                     Nvl(b.名称, a.名称) || Decode(a.产地, Null, Null, '(' || a.产地 || ')') ||
                      Decode(a.规格, Null, Null, ' ' || a.规格) As 药品信息, a.住院单位, a.应发数, a.退药数, a.实发数, a.金额
              From (Select b.药品id, c.编码 As 药品编码, c.名称, c.产地, c.规格, b.住院单位, Sum(a.应发数 / Nvl(b.住院包装, 1)) As 应发数,
                            Sum(a.退药数 / Nvl(b.住院包装, 1)) As 退药数, (Sum(a.应发数) - Sum(a.退药数)) / Nvl(b.住院包装, 1) As 实发数,
                            Sum(a.金额) As 金额
                     From (Select a.药品id, a.数量 As 应发数, 0 As 退药数, a.金额
                            From (Select a.状态, a.No, a.序号, a.库房id, a.对方部门id, a.姓名, a.住院号, a.床号, a.药品id, a.单价, a.期效, a.单量, a.频次,
                                          a.用法, a.时间, a.人员, a.付数, a.病人id, a.主页id, Sum(a.数量) As 数量, Sum(a.金额) As 金额
                                   From (Select b.状态, c.No, c.序号, c.库房id, c.对方部门id, c.姓名, '需补充' 住院号, '需补充' 床号, c.药品id, b.数量,
                                                 c.零售价 As 单价, b.金额, Decode(Nvl(Substr(c.扣率, 1, 1), 0), 0, '长嘱', '临嘱') As 期效, c.单量,
                                                 c.频次, c.用法, 'A.发送时间' As 时间, 'A.发送人' As 人员, c.付数, c.病人id, c.主页id
                                          From (Select Decode(a.审核人, Null, 0, 1) As 状态, a.No, a.序号, Sum(a.填写数量 * a.付数) As 数量,
                                                        Sum(a.零售金额) As 金额
                                                 From 药品收发记录 A
                                                 Where a.审核日期 Between d_审核起始日期 And d_审核结束日期 And a.单据 + 0 = 9 And
                                                       a.医嘱id Is Not Null And (a.对方部门id = n_对方部门id Or Nvl(n_对方部门id, 0) = 0) And
                                                       (a.库房id = n_库房id Or Nvl(n_库房id, 0) = 0) And
                                                       (a.No = v_No Or Nvl(v_No, 'NONE') = 'NONE') And
                                                       (a.汇总发药号 = n_汇总发药号 Or Nvl(n_汇总发药号, 0) = 0) And
                                                       (Nvl(Substr(a.扣率, 1, 1), 0) = n_效期 Or Nvl(n_效期, 2) = 2) And
                                                       (v_记录状态 = '01' And a.审核人 Is Not Null Or
                                                       v_记录状态 = '10' And Mod(a.记录状态, 3) = 1 And a.审核人 Is Null Or
                                                       Nvl(v_记录状态, 'NONE') = 'NONE') And
                                                       (a.用法 In (Select /*+cardinality(x,10)*/
                                                                  x.Column_Value
                                                                 From Table(f_Str2list(v_用法s)) X) Or Nvl(v_用法s, 'NONE') = 'NONE') And
                                                       (a.病人id + 0 In (Select /*+cardinality(x,10)*/
                                                                        x.Column_Value
                                                                       From Table(f_Str2list(v_病人ids)) X) Or
                                                       Nvl(v_病人ids, 'NONE') = 'NONE')
                                                 Group By Decode(a.审核人, Null, 0, 1), a.No, a.序号
                                                 Having Nvl(Sum(a.填写数量), 0) <> 0 Or Nvl(Sum(a.零售金额), 0) <> 0) B, 药品收发记录 C
                                          Where b.No = c.No And b.序号 = c.序号 And (c.记录状态 = 1 Or Mod(c.记录状态, 3) = 0) And
                                                (c.病人病区id = n_病区id Or Nvl(n_病区id, 0) = 0)) A
                                   Group By a.状态, a.No, a.序号, a.库房id, a.对方部门id, a.姓名, a.住院号, a.床号, a.药品id, a.单价, a.期效, a.单量,
                                            a.频次, a.用法, a.时间, a.人员, a.付数, a.病人id, a.主页id) A
                            Union All
                            --退药明细 l_退药信息 ，药品:数量_金额串  123123:13_33                             
                            Select /*+cardinality(x,10)*/
                             To_Number(x.C1) 药品id, 0 应发数, To_Number(Substr(x.C2, 1, Instr(x.C2, '_') - 1)) 退药数,
                             -1 * To_Number(Substr(x.C2, Instr(x.C2, '_') + 1)) 金额
                            From Table(l_退药信息) X) A, 药品规格 B, 收费项目目录 C
                     Where a.药品id = b.药品id And a.药品id = c.Id
                     Group By b.药品id, c.编码, c.名称, c.产地, c.规格, b.住院单位, Nvl(b.住院包装, 1)
                     Having Sum(a.应发数 / Nvl(b.住院包装, 1)) <> 0 Or Sum(a.退药数 / Nvl(b.住院包装, 1)) <> 0 Or Sum(a.金额) <> 0) A,
                   收费项目别名 B
              Where a.药品id = b.收费细目id(+) And b.码类(+) = 1 And b.性质(+) = n_Showtype
              Order By a.药品编码) Loop
      --程序外部还需要对 床号 进行排序，外部追加伪列表    
      r_Grp := R;
      Get出参拼串汇总;
    End Loop;
    --汇总清单结点
    --药品编码，药品信息，住院单位，数量，退药数，实发数，金额
    --明细清单结点
    --状态，NO，序号，药房，开嘱科室，姓名，住院号，床号，住院单位，数量，单价，金额，期效，单量，剂量单位，频次，用法，时间，人员，付数，类别，退费id      
  End If;

  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"成功","data":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","data":[' || c_Jtmp || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Sendquery;
/


Create Or Replace Procedure Zl_Drugsvr_Getadditional_Infor
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：获取药品的一些扩展或附加的信息，包含：用法，剂量，频次，剂型等
  --入参：Json_In:格式
  --    input
  --        billtype                    N   1   单据类型:1 -收费处方发药  ;2- 记帐单处方发药
  --        rcp_no                  C   1   单据号
  --        rcpdtl_ids                  C       处方明细ids,目前传入的费用ID
  --    出参 json
  --    output
  --        code                    N   1   应答吗：0-失败；1-成功
  --        message                 C   1   应答消息：失败时返回具体的错误信息
  --        data[]                         更新数据列表[数组]
  --            rcp_no              C   1   NO
  --            rcpdtl_id               N   1   处方明细id,目前传入的费用ID
  --            frequency               C   1   频次
  --            usage               C   1   用法
  --            si_drug_form                C   1   剂型
  --            loitem_detail_measunit              C   1   剂量单位
  --            advice_exe_properties               N   1   执行性质:0~2-计价特性,3-离院带药,4-自取药
  ---------------------------------------------------------------------------
  v_No       药品收发记录.No%Type;
  n_单据类型 Number(2);
  v_单据性质 Varchar2(10);
  n_执行性质 Number(2);
  v_费用ids  Varchar2(32767);

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --解析入参
  j_Jsonin   := PLJson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_单据类型 := j_Json.Get_Number('billtype');
  v_No       := j_Json.Get_String('rcp_no');
  v_费用ids  := ',' || Nvl(j_Json.Get_String('rcpdtl_ids'), '') || ',';

  If v_No Is Null Then
    Json_Out := zlJsonOut('未明确单据');
    Return;
  End If;

  If Nvl(n_单据类型, 0) = 2 Then
    v_单据性质 := ',9,10,';
  Else
    v_单据性质 := ',8,';
  End If;

  For c_药品 In (Select a.费用id, a.No, a.扣率, a.频次, a.用法, d.名称 As 剂型, c.剂量单位
               From (Select a.费用id, a.药品id, a.No, Max(a.频次) As 频次, Max(a.用法) As 用法, Max(扣率) As 扣率
                      From 药品收发记录 A
                      Where Instr(v_单据性质, ',' || 单据 || ',') > 0 And NO = v_No
                      Group By a.No, a.费用id, a.药品id) A, 药品目录 B, 药品信息 C, 药品剂型 D
               Where Instr(v_费用ids, ',' || a.费用id || ',') > 0 And a.药品id = b.药品id And b.药名id = c.药名id And c.剂型 = d.编码) Loop
  
    n_执行性质 := To_Number(Substr(Nvl(c_药品.扣率, '00'), 2, 1));
  
    v_Jtmp := v_Jtmp || ',';
    zlJsonPutValue(v_Jtmp, 'rcp_no', c_药品.No, 0, 1);
    zlJsonPutValue(v_Jtmp, 'rcpdtl_id', c_药品.费用id, 1);
    zlJsonPutValue(v_Jtmp, 'frequency', c_药品.频次);
    zlJsonPutValue(v_Jtmp, 'usage', c_药品.用法);
    zlJsonPutValue(v_Jtmp, 'si_drug_form', c_药品.剂型);
    zlJsonPutValue(v_Jtmp, 'loitem_detail_measunit', c_药品.剂量单位);
    zlJsonPutValue(v_Jtmp, 'advice_exe_properties', n_执行性质, 1, 2);
  
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
    Json_Out := '{"output":{"code":1,"message":"成功","data":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","data":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getadditional_Infor;
/


Create Or Replace Procedure Zl_Drugsvr_Drugcalcprice
(
  Json_In  Clob,
  Json_Out Out Clob
) Is
  --批量获取变价药品的时价 
  --入参      json
  --input              根据条件对要产生的处方进行检查
  --   price_ddigits  N   1  金额小数位数  
  --  drug_list       药品明细信息，支持多个，[数组]
  --    drug_id       N   1  药品id
  --    pharmacy_id   N   1  药房id
  --    send_num      N   1  本次需要发送或者是新开药品的数量  
  --出参      json
  --output
  --  code     C 1 应答码：0-失败；1-成功
  --  message  C 1 应答消息：成功时返回成功信息失败时返回具体的错误信息
  --  fee_list       时价明细信息，支持多个，[数组]
  --    drug_id      N   1  药品id
  --    pharmacy_id  N   1  药房id
  --    send_num     N   1  本次需要发送或者是新开药品的数量 
  --    price        N   1  时价
  j_List Pljson_List;
  j_Temp PLJson;

  n_Medioutmode Number; --分批药品出库方式
  n_Decprice    Number; --金额小数位数  

  n_药品id Number;
  n_药房id Number;
  n_数量   Number(16, 5);
  n_时价   Number(16, 5);

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;

  j_Jsonin PLJson;
  j_Json   PLJson;

  Function Calcdrugprice
  (
    药房id_In   药品库存.库房id%Type,
    药品id_In   药品库存.药品id%Type,
    数量_In     药品库存.可用数量%Type,
    出库算法_In Number,
    Decprice_In Number,
    Dec_In      Number
  ) Return Number Is
    --功能:返回变价药品的时价 
    --参数: 
    --     数量_In 本次需要发送或者是新开药品的数量 
    --     出库算法_In 0-按批次先进先出，1-按效期最近先出 
    --     Decprice_In 金额小数位数 
    --     Dec_In 价格小数位 
    n_时价     Number(16, 5);
    n_首批时价 Number(16, 5);
    n_总金额   Number;
    n_总数量   Number;
    n_当前数量 Number;
    n_Cnt      Number;
  Begin
    If 数量_In <= 0 Then
      Return 0;
    End If;
  
    n_总金额 := 0;
    n_总数量 := 数量_In;
    For Rs In (Select Nvl(批次, 0) As 批次, Nvl(可用数量, 0) As 库存,
                      Nvl(零售价, Nvl(Decode(Nvl(实际数量, 0), 0, 0, 实际金额 / 实际数量), 0)) As 时价
               From 药品库存
               Where 库房id = 药房id_In And 药品id = 药品id_In And Nvl(可用数量, 0) > 0 And 性质 = 1 And
                     (Nvl(批次, 0) = 0 Or 效期 Is Null Or 效期 > Trunc(Sysdate))
               Order By Decode(出库算法_In, 1, 效期, To_Date('2008-08-08', 'yyyy-mm-dd')), Decode(出库算法_In, 2, 上次批号, Null),
                        Nvl(批次, 0)) Loop
      --第一个批次的时价 
      n_Cnt := n_Cnt + 1;
      If n_Cnt = 1 Then
        n_首批时价 := Round(Rs.时价, Decprice_In);
      End If;
    
      If n_总数量 = 0 Then
        Exit;
      End If;
    
      If n_总数量 <= Rs.库存 Then
        n_当前数量 := n_总数量;
      Else
        n_当前数量 := Rs.库存;
      End If;
    
      n_总金额 := n_总金额 + Round(n_当前数量 * Round(Rs.时价, Decprice_In), Dec_In);
      n_总数量 := n_总数量 - n_当前数量;
    
      If n_总数量 = 0 Then
        Exit;
      End If;
    End Loop;
  
    If n_总数量 <> 0 Then
      -- 库存不够,只涉及一个批次时以首批时价为准，否则以第一批或者平均价都不合适 
      If n_Cnt = 1 Then
        n_时价 := n_首批时价;
      Else
        n_时价 := 0;
      End If;
    Else
      If n_Cnt = 1 Then
        n_时价 := n_首批时价;
      Else
        n_时价 := Round(n_总金额 / 数量_In, Decprice_In);
      End If;
    End If;
  
    Return n_时价;
  End Calcdrugprice;
Begin
  --解析入参
  j_Jsonin   := PLJson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_Decprice := j_Json.Get_Number('price_ddigits');
  j_List     := j_Json.Get_Pljson_List('drug_list');

  --获取参数
  n_Medioutmode := Nvl(zl_GetSysParameter(150), 0);

  --循环读取时价
  v_Jtmp := Null;
  For I In 1 .. j_List.Count Loop
    j_Temp   := PLJson(j_List.Get(I));
    n_药品id := j_Temp.Get_Number('drug_id');
    n_数量   := Nvl(j_Temp.Get_Number('send_num'), 0);
    n_药房id := j_Temp.Get_Number('pharmacy_id');
  
    n_时价 := Nvl(Calcdrugprice(n_药房id, n_药品id, n_数量, n_Medioutmode, n_Decprice, n_Decprice), 0);
  
    v_Jtmp := v_Jtmp || ',';
    zlJsonPutValue(v_Jtmp, 'drug_id', n_药品id, 1, 1);
    zlJsonPutValue(v_Jtmp, 'pharmacy_id', n_药房id, 1);
    zlJsonPutValue(v_Jtmp, 'send_num', n_数量, 1);
    zlJsonPutValue(v_Jtmp, 'price', n_时价, 1, 2);
  
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
    Json_Out := '{"output":{"code":1,"message":"成功","fee_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","fee_list":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Drugcalcprice;
/
Create Or Replace Procedure Zl_Drugsvr_Rcplock
(
  Json_In  Clob,
  Json_Out Out Clob
) Is
  --功能：处方审查撤销 
  --入参：json格式
  --input
  ----rcp_ids C 1 处方id
  --出参：json格式
  --output
  ----code 0-失败 1-成功
  ----message 成功和失败后返回的信息
  ----lockadvice_ids 锁定的医嘱id
  c_Rcpids Clob;
  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_Rcpid Collection_Type;
  I         Number;

  n_Count Number;
  v_Error Varchar2(255);
  Err_Custom Exception;

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;

  j_Jsonin PLJson;
  j_Json   PLJson;
Begin
  --解析入参
  j_Jsonin := PLJson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  c_Rcpids := j_Json.Get_Clob('rcp_ids');

  I := 0;
  While c_Rcpids Is Not Null Loop
    If Length(c_Rcpids) <= 4000 Then
      Col_Rcpid(I) := c_Rcpids;
      c_Rcpids := Null;
    Else
      Col_Rcpid(I) := Substr(c_Rcpids, 1, Instr(c_Rcpids, ',', 3980) - 1);
      c_Rcpids := Substr(c_Rcpids, Instr(c_Rcpids, ',', 3980) + 1);
    End If;
    I := I + 1;
  End Loop;

  For I In 0 .. Col_Rcpid.Count - 1 Loop
    For r_Info In (Select /* +RULE*/
                   Distinct b.Id, b.状态
                   From 处方审查明细 A, 处方审查记录 B, Table(f_Num2List(Col_Rcpid(I), ',')) C
                   Where a.审方id = b.Id And a.医嘱id = c.Column_Value And a.最后提交 = 1 And
                         (b.状态 Between 0 And 1 Or b.状态 Is Null) And b.审查结果 Is Null) Loop
    
      Select Count(1) Into n_Count From 处方审查记录 Where ID = r_Info.Id And 锁定用户 Is Not Null;
    
      If n_Count = 0 Then
        --未锁定 
        If Nvl(r_Info.状态, 0) = 0 Then
          --未审查，直接删除记录 
          Delete 处方审查记录 Where ID = r_Info.Id And (状态 = 0 Or 状态 Is Null);
        Elsif r_Info.状态 = 1 Then
          --已审查，调整状态 
          Update 处方审查记录 Set 状态 = 状态 + 10 Where ID = r_Info.Id And 状态 = 1;
        End If;
      
      Else
        --被锁定  
        v_Jtmp := v_Jtmp || ',' || r_Info.Id;
      
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

  If c_Jtmp Is Null Then
    c_Jtmp := Substr(v_Jtmp, 2);
  Else
    c_Jtmp := c_Jtmp || v_Jtmp;
  End If;

  If c_Jtmp Is Not Null Then
    v_Error := '操作的医嘱中存在已锁定的医嘱，正在进行处方审查，不能再删除。';
    Raise Err_Custom;
  End If;

  Json_Out := '{"output":{"code":1,"message":"成功","lockadvice_ids":"' || c_Jtmp || '"}}';
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugsvr_Rcplock;
/
Create Or Replace Procedure Zl_Drugsvr_Getnotsendrec
(
  Json_In  Varchar,
  Json_Out Out Varchar
) Is
  --------------------------------------------------------------------------- 
  --功能：获取未发药品记录
  --入参：JSON格式
  --input
  --  billtypes             C  1 单据类型，多个用英文逗号分隔: 1-收费处方;2-记帐单处方;3-记帐表处方
  --  charge_tag            N  1 收费标志:0-未收费或记帐划价;1-已收费或记帐
  --  fee_source            C  1 费用来源，多个用英文逗号分隔:1-门诊,2-住院,4-体检
  --  start_time            C  0 开始时间:yyyy-mm-dd hh:mi:ss
  --  end_time              C  0 结束时间:yyyy-mm-dd hh:mi:ss
  --出参：JSON格式
  --output
  --  code                  N 1 应答吗：0-失败；1-成功
  --  message               C 1 应答消息：失败时返回具体的错误信息
  --  data[]
  --    billtype            N  1 单据类型:1-收费处方,2-记帐单处方,3-记帐表处方
  --    rcp_no              C  1 处方单号
  --    pharmacy_id         N  1 药房ID
  ---------------------------------------------------------------------------
  v_单据类型 Varchar2(100);
  n_收费标志 Number(1);
  v_费用来源 Varchar2(100);
  d_开始时间 Date;
  d_结束时间 Date;

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;

  j_Jsonin PLJson;
  j_Json   PLJson;
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
  v_单据类型 := Replace(v_单据类型, ',1,', ',8,');
  v_单据类型 := Replace(v_单据类型, ',2,', ',9,');
  v_单据类型 := Replace(v_单据类型, ',3,', ',10,');

  For r_药品 In (Select Decode(b.单据, 8, 1, 9, 2, 10, 3) As 单据类型, b.No, b.库房id
               From 未发药品记录 B
               Where Instr(v_单据类型, ',' || b.单据 || ',') > 0 And Nvl(b.已收费, 0) = n_收费标志 And
                     b.填制日期 Between Nvl(d_开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And Nvl(d_结束时间, Sysdate) And Exists
                (Select 1
                      From 药品收发记录
                      Where 单据 = b.单据 And NO = b.No And
                            (Instr(',' || v_费用来源 || ',', ',' || 费用来源 || ',') > 0 Or 费用来源 Is Null))) Loop
  
    v_Jtmp := v_Jtmp || ',{"billtype":' || r_药品.单据类型;
    v_Jtmp := v_Jtmp || ',"rcp_no":"' || r_药品.No || '"';
    v_Jtmp := v_Jtmp || ',"pharmacy_id":' || Nvl(r_药品.库房id, 0);
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
    Json_Out := '{"output":{"code":1,"message":"成功","data":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","data":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Getnotsendrec;
/

Create Or Replace Procedure Zl_Drugsvr_Odr_Check
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：超期发送收回药品相关取数和检查
  --入参：Json_In:格式
  --     chk_type                    N 1 检查方式，0-获取列表，1-判断中药是否已经发药
  --     item_list[]列表
  --               order_id          N 1 医嘱ID
  --               rcp_nos           C 1 单据号拼串
  --出参: Json_Out,格式如下
  --  output
  --    code                          N 1 应答吗：0-失败；1-成功
  --    message                       C 1 应答消息：失败时返回具体的错误信息
  --    data{}
  --         isexist                        N 1 是否存已发药的中药
  --         item_list[]
  --             rcpdtl_id                  N 1 处方明细id
  --             sended_num                 N 1 已发药品数量
  --             order_id                   N 1 医嘱id
  --             drug_id                    N 1 药品id
  ---------------------------------------------------------------------------
  j_Input        Pljson;
  j_Item         Pljson;
  j_List         Pljson_List := Pljson_List();
  n_医嘱id       Number(18);
  n_检查方式     Number;
  v_中药存在发药 Varchar2(3000);
  n_Count        Number;
  v_Nos          Varchar2(30000);
  v_List         Varchar2(32767);
  v_Json_Out     Varchar2(32767);
  v_Data_Out     Varchar2(32767);
  v_Error        Varchar2(255);
  Err_Custom Exception;
Begin
  j_Item     := Pljson(Json_In);
  j_Input    := j_Item.Get_Pljson('input');
  n_检查方式 := j_Input.Get_Number('chk_type');
  j_List     := j_Input.Get_Pljson_List('item_list');
  If j_List Is Not Null Then
    j_Item := Pljson();
    For I In 1 .. j_List.Count Loop
      j_Item   := Pljson(j_List.Get(I));
      n_医嘱id := j_Item.Get_Number('order_id');
      v_Nos    := j_Item.Get_String('rcp_nos');
      j_Item   := Pljson();
      If n_检查方式 = 1 Then
        Select Count(1)
        Into n_Count
        From (Select /*+cardinality(j,10)*/
                a.No, a.费用id, a.药品id, Sum(Nvl(a.付数, 1) * Decode(a.审核人, Null, 0, a.实际数量)) As 已发数量
               From 药品收发记录 A, Table(f_Str2list(v_Nos)) J
               Where a.No = j.Column_Value And a.医嘱id = n_医嘱id
               Group By a.No, a.费用id, a.药品id) A
        Where a.已发数量 > 0;
        If n_Count > 0 Then
          v_中药存在发药 := ',"isexist":1';
          Exit;
        End If;
      Else
        For c_药品 In (Select /*+cardinality(j,10)*/
                      a.No, a.费用id, a.药品id, Sum(Nvl(a.付数, 1) * Decode(a.审核人, Null, 0, a.实际数量)) As 已发数量
                     From 药品收发记录 A, Table(f_Str2list(v_Nos)) J
                     Where a.No = j.Column_Value And a.医嘱id = n_医嘱id
                     Group By a.No, a.费用id, a.药品id) Loop
        
          v_List := v_List || ',{"rcpdtl_id":' || c_药品.费用id;
          v_List := v_List || ',"sended_num":' || Zljsonstr(c_药品.已发数量, 1);
          v_List := v_List || ',"order_id":' || n_医嘱id;
          v_List := v_List || ',"drug_id":' || c_药品.药品id;
          v_List := v_List || '}';
        
        End Loop;
      End If;
    End Loop;
  
    If v_List Is Not Null Then
      v_List := ',"item_list":[' || Substr(v_List, 2) || ']';
    End If;
  
    v_Data_Out := v_中药存在发药 || v_List;
    v_Json_Out := '{"code":1,"message":"成功","data":{';
    v_Json_Out := v_Json_Out || Substr(v_Data_Out, 2);
    v_Json_Out := v_Json_Out || '}}';
    Json_Out   := '{"output":' || v_Json_Out || '}';
  
  End If;
Exception
  When Err_Custom Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(v_Error) || '"}}';
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Drugsvr_Odr_Check;
/
Create Or Replace Procedure Zl_Drugsvr_Overdue_Recovery
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：超期发送收回药品相关处理
  --入参：Json_In:格式
  --  input
  --     operator_name                      C 1 操作员姓名
  --     operator_time                      C 1 操作时间
  --     item_list[]药品删除列表
  --                  rcpdtl_id              N 1 处方明细id,目前传入的费用ID
  --                  chargeoffs_num         N 1 销帐数量

  --     roll_list[]负数收回列表
  --                  clinic_type            C 1 医嘱诊疗类别
  --                  rcp_no                 C 1 处方号,费用单号
  --                  rcpdtl_id              N 1 处方明细ID
  --                  rcpdtl_id_old          N 1 处方明细ID,原始明细id
  --                  packages_num           N 1 发药付数
  --                  send_num               N 1 发药数量
  --                  item_type              C 1 收费项目类别

  --出参: Json_Out,格式如下
  --  output
  --    code                          N 1 应答吗：0-失败；1-成功
  --    message                       C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Input      Pljson;
  j_Item       Pljson;
  j_List       Pljson_List := Pljson_List();
  n_配液id     药品收发记录.Id%Type;
  n_处方明细id 药品收发记录.Id%Type;
  n_销帐数量   药品收发记录.填写数量%Type;

  v_人员姓名  Varchar2(3000);
  收回时间_In Date;
  v_Dec       Number;
  v_收发序号  Number;
  No_In       Varchar2(3000);
  v_费用id    Number;
  n_收费标志  Number;
  Old_费用id  Number;
  v_当前付数  Number;
  v_当前数量  Number;
  v_诊疗类别  Varchar2(3000);
  n_数量      Number;
  n_Count     Number;
  Cursor c_Drug Is
    Select b.批次, Nvl(x.药房分批, 0) As 分批, b.批号, b.效期, x.最大效期, b.Id As 收发id, b.病人id, b.主页id, b.库房id, b.单据, b.姓名, b.对方部门id,
           b.身份
    From 药品收发记录 B, 药品规格 X
    Where b.费用id = Old_费用id And b.药品id = x.药品id;

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
    P付数         药品收发记录.付数%Type,
    P数量         药品收发记录.填写数量%Type,
    P身份         药品收发记录.身份%Type
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
  
    --姓名,性别,年龄,出生日期,身份证号,病人ID,主页ID,
    --病人科室id,病人病区id,婴儿序号
    --,病人来源,医嘱id,身份,处方类型,皮试结果,诊断描述,已收费,费用来源
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
  
    If Sql%RowCount = 0 Then
    
      If P身份 Is Not Null Then
        Select Max(b.优先级) Into v_优先级 From 身份 B Where b.名称 = P身份;
      End If;
    
      Insert Into 未发药品记录
        (单据, NO, 病人id, 主页id, 姓名, 优先级, 对方部门id, 库房id, 填制日期, 已收费, 打印状态)
      Values
        (单据_In, No_In, 病人id_In, 主页id_In, 姓名_In, v_优先级, 对方部门id_In, 库房id_In, 收回时间_In, n_收费标志, 0);
    End If;
  
  End 负数收发记录_Insert;

Begin
  j_Item  := Pljson(Json_In);
  j_Input := j_Item.Get_Pljson('input');

  --药品删除列表
  j_List := j_Input.Get_Pljson_List('item_list');
  If j_List Is Not Null Then
    For I In 1 .. j_List.Count Loop
      j_Item       := Pljson();
      j_Item       := Pljson(j_List.Get(I));
      n_处方明细id := j_Item.Get_Number('rcpdtl_id');
      n_销帐数量   := j_Item.Get_Number('chargeoffs_num');
      n_配液id     := j_Item.Get_Number('pivas_id');
      Zl_药品收发记录_销售退费_s(n_处方明细id, n_销帐数量, n_配液id, 1);
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
    
      Select Nvl(Max(序号), 0) + 1 Into v_收发序号 From 药品收发记录 Where 单据 = 9 And 记录状态 = 1 And NO = No_In;
    
      v_费用id   := j_Item.Get_Number('rcpdtl_id');
      No_In      := j_Item.Get_String('rcp_no');
      v_当前付数 := j_Item.Get_Number('packages_num');
      v_当前数量 := j_Item.Get_Number('send_num');
      v_诊疗类别 := j_Item.Get_String('clinic_type'); --医嘱记录中的诊疗类别
      Old_费用id := j_Item.Get_Number('rcpdtl_id_old');
      n_收费标志 := j_Item.Get_Number('charge_tag');
    
      For r_Drug In c_Drug Loop
      
        If v_诊疗类别 In ('5', '6', '7') Then
          负数收发记录_Insert(v_费用id, r_Drug.批次, r_Drug.分批, r_Drug.批号, r_Drug.效期, r_Drug.最大效期, r_Drug.收发id, r_Drug.病人id,
                        r_Drug.主页id, r_Drug.库房id, r_Drug.单据, r_Drug.姓名, r_Drug.对方部门id, v_当前付数, v_当前数量, r_Drug.身份);
        Else
          n_数量 := v_当前数量;
          For r_Otherdrug In (Select b.批次, Nvl(x.药房分批, 0) As 分批, Nvl(b.付数, 1) * b.实际数量 As 数量, b.批号, b.效期, x.最大效期,
                                     b.Id As 收发id, b.病人id, b.主页id, b.库房id, b.单据, b.姓名, b.对方部门id, b.身份
                              From 药品收发记录 B, 药品规格 X
                              Where b.费用id = Old_费用id And (b.记录状态 = 1 Or Mod(b.记录状态, 3) = 0) And b.药品id = x.药品id
                              Order By b.Id Desc) Loop
            If n_数量 > 0 Then
              n_Count := r_Otherdrug.数量;
              If n_数量 < n_Count Then
                n_Count := n_数量;
              End If;
              负数收发记录_Insert(v_费用id, r_Otherdrug.批次, r_Otherdrug.分批, r_Otherdrug.批号, r_Otherdrug.效期, r_Otherdrug.最大效期,
                            r_Otherdrug.收发id, r_Otherdrug.病人id, r_Otherdrug.主页id, r_Otherdrug.库房id, r_Otherdrug.单据,
                            r_Otherdrug.姓名, r_Otherdrug.对方部门id, 1, n_Count, r_Otherdrug.身份);
              n_数量 := n_数量 - r_Otherdrug.数量;
            End If;
          End Loop;
        End If;
      End Loop;
    End Loop;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugsvr_Overdue_Recovery;
/