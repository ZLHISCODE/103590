Create Or Replace Procedure Zl_Cissvr_Addadviceannex
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：增加医嘱附费信息
  --入参：Json_In:格式
  --  input
  --    bill_no                 C 1 附费:No
  --    bill_prop               N 1 附费:记录性质
  --    advice_id               N 1 医嘱ID
  --    send_no                 N 1 发送号

  --出参: Json_Out,格式如下
  --  output
  --    code                      N 1 应答吗：0-失败；1-成功
  --    message                   C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Json     Pljson;
  j_In       Pljson;
  v_No       病人医嘱附费.No%Type;
  n_记录性质 病人医嘱附费.记录性质%Type;
  n_发送号   病人医嘱附费.发送号%Type;
  n_医嘱id   病人医嘱附费.医嘱id%Type;

Begin
  --解析入参
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  n_医嘱id   := j_Json.Get_Number('advice_id');
  n_发送号   := j_Json.Get_Number('send_no');
  v_No       := j_Json.Get_String('bill_no');
  n_记录性质 := j_Json.Get_Number('bill_prop');

  If Nvl(n_医嘱id, 0) = 0 Or Nvl(n_发送号, 0) = 0 Then
    Json_Out := Zljsonout('未传入医嘱id或发送号，请检查！');
    Return;
  End If;

  If Nvl(n_记录性质, 0) = 0 Then
    Json_Out := Zljsonout('未传记录性质，请检查！');
    Return;
  End If;

  If Nvl(v_No, '-') = '-' Then
    Json_Out := Zljsonout('未传入NO，请检查！');
    Return;
  End If;

  Zl_病人医嘱附费_Insert(n_医嘱id, n_发送号, n_记录性质, v_No);

  Json_Out := Zljsonout('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Addadviceannex;
/
Create Or Replace Procedure Zl_Cissvr_Adviceannex_Add
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:增加医嘱附费信息
  --入参：Json_In:格式
  --    input
  --      advice_send_no  C  1  附费:No
  --      advice_send_properties  N  1  附费:记录性质
  --      advice_id  N  1  医嘱ID
  --      advice_send_number  N  1  发送号
  --出参: Json_Out,格式如下
  --  output
  --    code              N   1   应答码：0-失败；1-成功
  --    message           C   1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------

  j_Json Pljson;
  j_In   Pljson;
  v_No   病人医嘱发送.No%Type;

  n_记录性质 病人医嘱发送.记录性质%Type;
  n_医嘱id   病人医嘱发送.医嘱id%Type;

  n_发送号 病人医嘱发送.发送号%Type;

Begin
  --解析入参
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');

  v_No       := j_Json.Get_String('advice_send_no');
  n_记录性质 := j_Json.Get_Number('advice_send_properties');
  n_医嘱id   := j_Json.Get_Number('advice_id');
  n_发送号   := j_Json.Get_Number('advice_send_number');

  If v_No Is Null Or Nvl(n_记录性质, 0) = 0 Or Nvl(n_医嘱id, 0) = 0 Or Nvl(n_发送号, 0) = 0 Then
    Json_Out := Zljsonout('传入附费信息错误，请检查');
    Return;
  End If;

  Zl_病人医嘱附费_Insert(n_医嘱id, n_发送号, n_记录性质, v_No);

  Json_Out := Zljsonout('成功', 1);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Adviceannex_Add;
/
Create Or Replace Procedure Zl_Cissvr_Adviceexecuting
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：检查指定的医嘱是否正在执行
  --入参：Json_In:格式
  --input
  --    item_list
  --      advice_id               C  1  医嘱ID
  --      advice_send_no          C  1  发送单号
  --      advice_send_properties  N  1  记录性质
  --出参: Json_Out,格式如下
  --    output
  --        code                  N   1   应答吗：0-失败；1-成功
  --        message               C   1   应答消息：失败时返回具体的错误信息
  --        isexist               N   0   是否存在，1-存在;0-不存在
  ---------------------------------------------------------------------------
  j_Json        Pljson;
  j_Jsonlist_In Pljson_List;

  j_Json_Tmp Pljson;

  v_No       病人医嘱发送.No%Type;
  n_记录性质 病人医嘱发送.记录性质%Type;
  n_医嘱id   病人医嘱发送.医嘱id%Type;
  n_存在     Number(2);

  n_Count Number(18);
Begin
  --解析入参
  j_Json_Tmp    := Pljson(Json_In);
  j_Json        := j_Json_Tmp.Get_Pljson('input');
  j_Jsonlist_In := j_Json.Get_Pljson_List('item_list');
  n_存在        := 0;
  If Not j_Jsonlist_In Is Null Then
    n_Count := j_Jsonlist_In.Count;
    For I In 1 .. n_Count Loop
      j_Json_Tmp := Pljson();
      j_Json_Tmp := Pljson(j_Jsonlist_In.Get(I));
      n_医嘱id   := j_Json_Tmp.Get_Number('advice_id');
      v_No       := j_Json_Tmp.Get_String('advice_send_no');
      n_记录性质 := j_Json_Tmp.Get_Number('advice_send_properties');
    
      --走销帐申请流程的，不检查医保执行状态
      Select Nvl(Count(1), 0)
      Into n_Count
      From 病人医嘱发送
      Where 执行状态 = 3 And NO = v_No And 记录性质 = Nvl(n_记录性质, 0) And 医嘱id = Nvl(n_医嘱id, 0);
      If n_Count > 0 Then
        n_存在 := 1;
        Exit;
      End If;
    End Loop;
  End If;

  Json_Out := '{"output":{"code":1,"message":"成功","isexist":' || n_存在 || '}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Adviceexecuting;
/
Create Or Replace Procedure Zl_Cissvr_Adviceexistitem
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  ---------------------------------------------------------------------------
  --功能：获取病人医嘱信息中是否存在指定的收费项目
  --入参：Json_In:格式
  --input   
  --       advice_item_id        N  1  收费项目id
  --出参: Json_Out,格式如下
  --output
  --       code                  C  1  应答码：0-失败；1-成功
  --       message               C  1  应答消息：
  --       item_exits            N  1  是否存在：0-不存在，1-存在
  ---------------------------------------------------------------------------
  j_Input      Pljson;
  j_Json       Pljson;
  n_收费项目id Number(18);
  n_Count      Number(18);
Begin
  j_Input      := Pljson(Json_In);
  j_Json       := j_Input.Get_Pljson('input');
  n_收费项目id := j_Json.Get_Number('advice_item_id');
  Select Count(1) Into n_Count From 病人医嘱记录 Where 收费细目id = n_收费项目id And Rownum < 2;
  Json_Out := '{"output":{"code":1,"message":"成功","item_exits":' || Nvl(n_Count, 0) || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Adviceexistitem;
/
Create Or Replace Procedure Zl_Cissvr_Adviceishistory
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:检查医嘱是否已经进入历史备份表中
  --入参：Json_In:格式
  --    input
  --       advice_list[]     数组
  --             advice_id           N 1 医嘱ID
  --             send_no             N 1 发送号

  --出参: Json_Out,格式如下
  --  output
  --    code                  N   1   应答码：0-失败；1-成功
  --    message               C   1   应答消息：失败时返回具体的错误信息
  --    is_history            N   1   是否存在:1-存在;0-不存在
  ---------------------------------------------------------------------------

  j_Json    Pljson;
  o_Json    Pljson;
  j_List    Pljson_List := Pljson_List();
  n_医嘱id  病人医嘱发送.医嘱id%Type;
  n_发送号  病人医嘱发送.发送号%Type;
  n_Count   Number(18);
  n_Exist   Number(1);
  n_Exist_h Number(1);

Begin
  --解析入参
  o_Json    := Pljson(Json_In);
  j_Json    := o_Json.Get_Pljson('input');
  j_List    := j_Json.Get_Pljson_List('advice_list');
  n_Exist   := 0;
  n_Exist_h := 0;
  If Not j_List Is Null Then
    n_Count := j_List.Count;
    For I In 1 .. n_Count Loop
      o_Json   := Pljson();
      o_Json   := Pljson(j_List.Get(I));
      n_医嘱id := o_Json.Get_Number('advice_id');
      n_发送号 := o_Json.Get_Number('send_no');
      If Nvl(n_医嘱id, 0) = 0 Or Nvl(n_发送号, 0) = 0 Then
        Json_Out := Zljsonout('未传入医嘱id或发送号，请检查！');
        Return;
      End If;
      Select Max(1) Into n_Exist From 病人医嘱发送 Where 医嘱id = n_医嘱id And 发送号 = n_发送号;
      If Nvl(n_Exist, 0) = 0 Then
        Select Max(1) Into n_Exist_h From H病人医嘱发送 Where 医嘱id = n_医嘱id And 发送号 = n_发送号;
        If Nvl(n_Exist_h, 0) = 1 Then
          Exit;
        End If;
      End If;
    End Loop;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功","is_history":' || Nvl(n_Exist_h, 0) || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Adviceishistory;
/
CREATE OR REPLACE Procedure Zl_Cissvr_Adviceisinvalid
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：根据医嘱ID查询医嘱状态
  --入参：Json_In:格式
  --input 
  --   advice_ids           C  1  多个医嘱ID，用,分隔
  --出参: Json_Out,格式如下
  --output
  --    code                 N  1  应答码：0-失败；1-成功
  --    message              C  1  应答消息
  --    advice_ids           C  1  医嘱ID（已作废的）
  ---------------------------------------------------------------------------
  v_医嘱id Clob; --记录医嘱id
  v_Tmp    Varchar2(32767); --作为中间变量
  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_医嘱id Collection_Type;
  I          Number;
  j_In       Pljson;
  j_Json     Pljson;
Begin
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  v_医嘱id := j_Json.Get_Clob('advice_ids');
  --将 v_医嘱id 串组装成不超过4000 的集合串，防止使用 f_Num2list 参数超长
  I := 0;
  While v_医嘱id Is Not Null Loop
    If Length(v_医嘱id) <= 4000 Then
      Col_医嘱id(I) := v_医嘱id;
      v_医嘱id := Null;
    Else
      Col_医嘱id(I) := Substr(v_医嘱id, 1, Instr(v_医嘱id, ',', 3980) - 1);
      v_医嘱id := Substr(v_医嘱id, Instr(v_医嘱id, ',', 3980) + 1);
    End If;
    I := I + 1;
  End Loop;
  I     := 0;
  v_Tmp := Null;
  For I In 0 .. Col_医嘱id.Count - 1 Loop
    For v_医嘱作废 In (Select /*+cardinality(b,10)*/
                   Distinct ID
                   From 病人医嘱记录 A, Table(f_Num2list(Col_医嘱id(I))) B
                   Where Nvl(医嘱状态, 0) = 4 And ID = Column_Value) Loop
    
      v_Tmp := v_Tmp || ',' || v_医嘱作废.Id;
    End Loop;
  End Loop;
  Json_Out := '{"output":{"code":1,"message": "成功","advice_ids":"' || Substr(v_Tmp, 2) || '"}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Adviceisinvalid;
/
Create Or Replace Procedure Zl_Cissvr_Auditadvicecharge
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:对指定医嘱进行费用审核完成
  --入参：Json_In:格式
  -- input
  --   advice_id            N 1 医嘱id
  --   verfy_statu          N 1 审核状态:1-审核;0-取消审核
  --出参: Json_Out,格式如下
  --  output
  --    code               N  1 应答码：0-失败；1-成功
  --    message            C  1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Json     Pljson;
  j_In       Pljson;
  n_医嘱id   病人医嘱记录.Id%Type;
  n_费用审核 病人医嘱记录.是否费用审核%Type;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --解析入参
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  n_医嘱id   := j_Json.Get_Number('advice_id');
  n_费用审核 := Nvl(j_Json.Get_Number('verfy_statu'), 0);

  If Nvl(n_医嘱id, 0) = 0 Then
    v_Error := '必须传入医嘱id！';
    Raise Err_Custom;
  End If;

  Zl_病人医嘱记录_费用审核(n_医嘱id, n_费用审核);
  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Auditadvicecharge;
/
Create Or Replace Procedure Zl_Cissvr_AuditDrugOrder
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:静配医嘱审核
  --入参：Json_In:格式 
  --input     静配医嘱审核
  --  auditor        C  1  审核人
  --  audit_content  C  1  医嘱审核内容，格式化串：ID1,操作1,说明1||ID2,操作2,说明2…

  --出参: Json_Out,格式如下
  --output
  --  code           N 1 应答码：0-失败；1-成功
  --  message        C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Json     Pljson;
  j_In       Pljson;
  v_审核人   Varchar2(20);
  v_审核内容 Varchar2(32767);
  v_Field    Varchar2(32767);
  v_Tmp      Varchar2(32767);
  n_医嘱id   Number(18);
  n_操作     Number(2);
  v_审核原因 Varchar2(100);
  Err_Custom Exception;
  v_Err Varchar2(255);
Begin
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');

  v_审核人   := j_Json.Get_String('auditor');
  v_审核内容 := j_Json.Get_String('audit_content');

  If v_审核内容 Is Null Then
    v_Err := '未处方明细信息【audit_content】节点';
    Raise Err_Custom;
    Return;
  End If;

  v_Tmp := v_审核内容 || '||';
  While v_Tmp Is Not Null Loop
    v_Field := Substr(v_Tmp, 1, Instr(v_Tmp, '||') - 1);
    v_Tmp   := Replace('||' || v_Tmp, '||' || v_Field || '||');
  
    n_医嘱id := To_Number(Substr(v_Field, 1, Instr(v_Field, ',') - 1));
    v_Field  := Substr(v_Field, Instr(v_Field, ',') + 1);
  
    n_操作     := To_Number(Substr(v_Field, 1, Instr(v_Field, ',') - 1));
    v_审核原因 := Substr(v_Field, Instr(v_Field, ',') + 1);
  
    If v_审核人 Is Null And n_操作 <> 0 Then
      Update 病人医嘱记录
      Set 药师审核标志 = n_操作, 药师审核时间 = Sysdate, 药师审核原因 = v_审核原因
      Where ID = n_医嘱id And Nvl(药师审核标志, 0) = 0;
    Else
      If n_操作 = 0 Then
        Update 病人医嘱记录
        Set 药师审核标志 = n_操作, 药师审核时间 = Null, 审核药师 = Null, 药师审核原因 = v_审核原因
        Where ID = n_医嘱id;
      Elsif n_操作 = 3 Then
        --部门发药审核医嘱，已审核过的不进行审核
        Update 病人医嘱记录
        Set 药师审核标志 = n_操作, 药师审核时间 = Sysdate, 审核药师 = v_审核人, 药师审核原因 = v_审核原因
        Where Nvl(药师审核标志, 0) = 0 And ID = n_医嘱id;
      Else
        Update 病人医嘱记录
        Set 药师审核标志 = n_操作, 药师审核时间 = Sysdate, 审核药师 = v_审核人, 药师审核原因 = v_审核原因
        Where ID = n_医嘱id;
      End If;
    End If;
  End Loop;

  Json_Out := '{"output":{"code": 1,"message":"成功"}}';
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_AuditDrugOrder;
/

Create Or Replace Procedure Zl_Cissvr_Buildadviceexecharge
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:重新生成医嘱执行计价数据
  --入参：Json_In:格式
  --    input
  --    pati_id             N 1 病人id
  --    pati_pageid         N 1 主页id
  --    bill_no             C 0 单据号
  --    bill_prop           N 0 记录性质
  --出参: Json_Out,格式如下
  --  output
  --    code                N 1 应答码：0-失败；1-成功
  --    message             C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------

  j_In       Pljson;
  j_Json     Pljson;
  n_病人id   病人医嘱记录.病人id%Type;
  n_主页id   病人医嘱记录.主页id%Type;
  v_No       病人医嘱发送.No%Type;
  n_记录性质 病人医嘱发送.记录性质%Type;

Begin
  --解析入参
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');
  n_主页id := j_Json.Get_Number('pati_pageid');

  v_No       := j_Json.Get_String('bill_no');
  n_记录性质 := j_Json.Get_Number('bill_prop');

  If v_No Is Null And Nvl(n_病人id, 0) <> 0 Then
    Json_Out := Zljsonout('未传入需要更新的单据信息或病人信息！');
    Return;
  End If;

  If Nvl(n_病人id, 0) <> 0 Then
    --按病人ID进行处理
  
    For c_记录 In (Select Distinct b.医嘱id, b.No, b.记录性质
                 From 病人医嘱记录 A, 病人医嘱发送 B, 医嘱执行计价 C
                 Where a.病人id = n_病人id And (a.主页id = Nvl(n_主页id, 0) Or n_主页id Is Null) And a.Id = b.医嘱id And
                       b.医嘱id = c.医嘱id And b.发送号 = c.发送号 And
                       ((b.No = v_No And b.记录性质 = Nvl(n_记录性质, 0) Or n_记录性质 Is Null Or v_No Is Null)) And c.执行状态 Is Null
                 Order By b.No) Loop
    
      Zl_医嘱执行计价_修正(c_记录.医嘱id, c_记录.No, c_记录.记录性质);
    End Loop;
  End If;

  For c_记录 In (Select Distinct b.医嘱id, b.No
               From 病人医嘱发送 B, 医嘱执行计价 C
               Where b.No = v_No And b.记录性质 = n_记录性质 And c.执行状态 Is Null
               Order By b.No) Loop
  
    Zl_医嘱执行计价_修正(c_记录.医嘱id, c_记录.No, n_记录性质);
  End Loop;

  Json_Out := Zljsonout('成功', 1);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Buildadviceexecharge;
/
Create Or Replace Procedure Zl_Cissvr_Checkbabylimit
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：检查转科婴儿婴儿是否允许补录操作
  --入参：Json_In:格式
  --  input
  --    pati_id             N 1 病人id
  --    pati_pageid         N 1 主页id
  --    baby_num            N 1 婴儿序号
  --出参: Json_Out,格式如下
  --  output
  --    code                N 1 应答吗：0-失败；1-成功
  --    message             C 1 应答消息：失败时返回具体的错误信息
  --    in_time             C 1 入院时间
  --    baby_wardarea_id    N 1 婴儿病区id
  --    have_data           N 1 是否有记录
  ---------------------------------------------------------------------------
  j_Json       Pljson;
  j_In         Pljson;
  v_入院日期   Varchar2(200);
  n_婴儿病区id Number;
  n_Havedata   Number := 0;
Begin
  --解析入参
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');
  For R In (Select b.入院日期, Nvl(a.婴儿病人id, 0) As 婴儿病人id
            From 病人新生儿记录 A, 病案主页 B
            Where a.婴儿病人id = b.病人id(+) And a.婴儿主页id = b.主页id(+) And a.病人id = j_Json.Get_Number('pati_id') And
                  a.主页id = j_Json.Get_Number('pati_pageid') And a.序号 = j_Json.Get_Number('baby_num')) Loop
    n_Havedata   := 1;
    v_入院日期   := To_Char(r.入院日期, 'yyyy-mm-dd hh24:mi:ss');
    n_婴儿病区id := r.婴儿病人id;
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"成功","in_time":"' || v_入院日期 || '","baby_wardarea_id":' || Nvl(n_婴儿病区id, 0) ||
              ',"have_data":' || n_Havedata || '}}';

Exception
  When Others Then 
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Checkbabylimit;
/
Create Or Replace Procedure Zl_Cissvr_Checkdepositerrorno
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：根据单据号获取存在病人结算异常记录中的NO
  --入参：Json_In:格式
  --  input
  --   pati_id            N 1 病人id
  --   bill_nos           C 1 病人预交记录.NO,多个用逗号分隔
  --出参: Json_Out,格式如下
  --  output
  --    code              N 1 应答码：0-失败；1-成功
  --    message           C 1 应答消息：失败时返回具体的错误信息
  --    bill_nos          C 1 有效的Nos,多个用逗号分隔
  ---------------------------------------------------------------------------
  n_病人id  病人结算异常记录.病人id%Type;
  v_Nos     Varchar2(32767);
  j_Json    Pljson;
  j_In      Pljson;
  v_Nos_Out Varchar2(32767);

Begin
  --解析入参
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_病人id := Nvl(j_Json.Get_Number('pati_id'), 0);
  v_Nos    := j_Json.Get_String('bill_nos');

  If Nvl(v_Nos, '-') = '-' Then
    Json_Out := Zljsonout('未传入NO，请检查！');
    Return;
  End If;

  Select /*+cardinality(B,10)*/
   f_List2str(Cast(Collect(b.Column_Value) As t_Strlist))
  Into v_Nos_Out
  From 病人结算异常记录 A, Table(f_Str2list(v_Nos)) B
  Where a.操作场景 = 3 And (a.预交单号 = b.Column_Value Or a.医疗卡单号 = b.Column_Value) And Decode(n_病人id, 0, 0, a.病人id) = n_病人id;

  Json_Out := '{"output":{"code":1,"message":"成功","bill_nos":"' || v_Nos_Out || '"}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Checkdepositerrorno;
/
Create Or Replace Procedure Zl_Cissvr_Checkpatexist
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：检查是否存在病人
  --入参：JSON格式
  --input
  --   pati_id       N 1 病人id
  --   visit_id   N 1 就诊id,门诊病人为挂号ID;住院病人为主页ID;为0说明是外来或非挂号就诊的病人
  --   occasion      N 1 场合,1-门诊;2-住院
  --   pati_name     C 1 姓名
  --   pati_sex      C 1 性别
  --   pati_age      C 1 年龄
  --   pati_birthdate C 1 出生日期
  --出参：JSON格式
  --output
  --  code            N 1 应答码：0-失败；1-成功
  --  message        C 1 应答消息：失败时返回具体的错误信息
  --   pati_name     C 1 姓名
  --   pati_sex      C 1 性别
  --   pati_age      C 1 年龄
  --   pati_birthdate C 1 出生日期
  ---------------------------------------------------------------------------
  n_病人id   病案主页.病人id%Type;
  n_就诊id   Number;
  n_场合     Number;
  v_姓名     病案主页.姓名%Type;
  v_性别     病案主页.性别%Type;
  v_年龄     病案主页.年龄%Type;
  d_出生日期 Date;
  j_In       Pljson;
  v_Error    Varchar2(2000);
  j_Json     Pljson;
Begin
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  n_病人id   := j_Json.Get_Number('pati_id');
  n_就诊id   := j_Json.Get_Number('visit_id');
  n_场合     := j_Json.Get_Number('occasion');
  v_姓名     := j_Json.Get_String('pati_name');
  v_性别     := j_Json.Get_String('pati_sex');
  v_年龄     := j_Json.Get_String('pati_age');
  d_出生日期 := To_Date(j_Json.Get_String('pati_birthdate'), 'yyyy-mm-dd hh24:mi:ss');
  Case n_场合
    When 1 Then
      Begin
        Select b.姓名, b.性别, b.年龄, a.出生日期
        Into v_姓名, v_性别, v_年龄, d_出生日期
        From (Select n_病人id As 病人id, v_姓名 As 姓名, v_性别 As 性别, v_年龄 As 年龄, d_出生日期 As 出生日期
               From Dual) A, 病人挂号记录 B
        Where a.病人id = b.病人id And a.病人id = n_病人id And b.Id = n_就诊id;
      Exception
        When Others Then
          v_Error := '病人ID[' || n_病人id || ']、挂号ID[' || n_就诊id || ']在病人挂号记录中不存在,不能继续进行病人信息变更操作!';
      End;
    When 2 Then
      Begin
        Select Nvl(b.姓名, a.姓名), Nvl(b.性别, a.性别), b.年龄, a.出生日期
        Into v_姓名, v_性别, v_年龄, d_出生日期
        From (Select n_病人id As 病人id, v_姓名 As 姓名, v_性别 As 性别, v_年龄 As 年龄, d_出生日期 As 出生日期
               From Dual) A, 病案主页 B
        Where a.病人id = b.病人id And a.病人id = n_病人id And b.主页id = n_就诊id;
      Exception
        When Others Then
          v_Error := '病人ID[' || n_病人id || ']、主页ID[' || n_就诊id || ']在病案主页中不存在,不能继续进行病人信息变更操作!';
      End;
    Else
      v_Error := '过程参数[场合]只能为1或2,不能继续进行病人信息变更操作!';
  End Case;
  If v_Error Is Not Null Then
    Json_Out := Zljsonout(v_Error, 0);
  Else
    Json_Out := '{"output":{"code":1,"message":"成功"';
    Json_Out := Json_Out || ',"pati_name":"' || Zljsonstr(v_姓名, 0) || '"';
    Json_Out := Json_Out || ',"pati_age":"' || Zljsonstr(v_年龄, 0) || '"';
    Json_Out := Json_Out || ',"pati_sex":"' || Zljsonstr(v_性别, 0) || '"';
    Json_Out := Json_Out || ',"pati_birthdate":"' || Zljsonstr(To_Char(d_出生日期, 'yyyy-mm-dd hh24:mi:ss'), 0) || '"' || '}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Checkpatexist;
/
Create Or Replace Procedure Zl_Cissvr_Checkpaticatalogue
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --功能：判断病人是否编目
  --入参：Json_In:格式
  --input
  --  pati_id           N    1 病人id
  --  pati_pageid       N    1主页id
  --出参：JSON格式
  --output
  --    code                N   1 应答吗：0-失败；1-成功
  --    message             C   1 应答消息：失败时返回具体的错误信息
  --    isexist             N  1 是否编目
  -------------------------------------------
  j_Json      Pljson;
  j_In        Pljson;
  n_病人id    Number(18);
  n_主页id    Number;
  d_Catalogue Date;
  n_Isexist   Number;
Begin
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');
  n_主页id := j_Json.Get_Number('pati_pageid');

  Begin
    Select 编目日期 Into d_Catalogue From 病案主页 Where 病人id = n_病人id And 主页id = n_主页id;
  Exception
    When Others Then
      Null;
  End;
  If d_Catalogue Is Null Then
    n_Isexist := 0;
  Else
    n_Isexist := 1;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功","isexist":' || n_Isexist || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Checkpaticatalogue;
/
Create Or Replace Procedure Zl_Cissvr_Checkpatiexecute
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：根据病人信息获取医技未执行的项目
  --入参：Json_In:格式
  --  input
  --     pati_id              N 1 病人ID
  --     pati_pageid          N 1 主页ID
  --     baby_num             N 0 婴儿序号:-1表示不区分;0-母亲的;>0具体婴儿费用
  --     fee_source           N 1 费用来源:1-门诊;2-住院;4-体检
  --     check_type           N 0 检查类型，null/0-表示检查检查其他执行项目，1-表示检查未发血
  --出参: Json_Out,格式如下
  --  output
  --    code                  N 1 应答吗：0-失败；1-成功
  --    message               C 1 应答消息：失败时返回具体的错误信息
  --    isexist               N 1 是否存在: 1-存在;0-不存在
  --    notexecute_infor      C 1 未执行的项目信息,isexist=1时返回
  ---------------------------------------------------------------------------
  j_Json     Pljson;
  j_In       Pljson;
  n_病人id   Number;
  n_主页id   Number;
  n_婴儿序号 Number;
  n_费用来源 Number;
  n_检查方式 Number;
  n_共享     Number;
  Type t_Bool Is Ref Cursor;
  c_Bool t_Bool;
  v_项目 Varchar2(32767);
  v_部门 Varchar2(32767);
  v_扣率 Varchar2(100);

  v_Sql   Varchar2(32767);
  v_Pars  Varchar2(4000);
  v_Text  Varchar2(4000);
  n_Count Number(18);
Begin
  --解析入参
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  n_病人id   := j_Json.Get_Number('pati_id');
  n_主页id   := j_Json.Get_Number('pati_pageid');
  n_婴儿序号 := j_Json.Get_Number('baby_num');
  n_费用来源 := j_Json.Get_Number('fee_source');
  n_检查方式 := j_Json.Get_Number('check_type');
  If Nvl(n_病人id, 0) = 0 Then
    Json_Out := '{"output":{"code":0,"message":"未传入病人信息！"}}';
    Return;
  End If;

  n_Count := 0;
  If Nvl(n_检查方式, 0) = 0 Then
    --1.医技科室执行的项目,临床会诊
    --2.其他类特殊项目不需管执行
    --3.PACS已报到的(执行过程为">=2-检查中"不作为未执行完成的项目
    If Nvl(n_费用来源, 0) = 2 Then
      Select zl_GetSysParameter(234) Into v_Pars From Dual;
      v_Pars := Replace(v_Pars, '|', ',');
      For r_Info In (Select Distinct b.No, c.名称 As 项目, d.名称 As 科室, Decode(Nvl(b.执行状态, 0), 0, '未执行', 3, '正在执行') As 执行状态
                     From 病人医嘱记录 A, 病人医嘱发送 B, 诊疗项目目录 C, 部门表 D
                     Where a.病人id = n_病人id And Nvl(a.主页id, 0) = n_主页id And (Nvl(a.婴儿, 0) = n_婴儿序号 Or n_婴儿序号 = -1) And
                           a.Id = b.医嘱id And b.执行状态 In (0, 3) And a.诊疗项目id = c.Id And b.执行部门id + 0 = d.Id And
                           a.诊疗类别 Not In ('4', '5', '6', '7') And Not (a.诊疗类别 In ('F', 'D') And a.相关id Is Not Null) And
                           (d.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or d.撤档时间 Is Null) And
                           Not (a.诊疗类别 = 'D' And Nvl(b.执行过程, 0) >= 2) And
                           (Not (a.诊疗类别 = 'Z' And Nvl(c.操作类型, '0') <> '0') Or a.诊疗类别 = 'Z' And c.操作类型 = '7') And
                           c.Id Not In (Select /*+cardinality(j,10) */
                                         Column_Value
                                        From Table(Cast(f_Num2list(v_Pars) As Zltools.t_Numlist)) J)) Loop
        If Lengthb(v_Text || Chr(13) || Chr(10) || '单据[' || Nvl(r_Info.No, '') || ']中的' || Nvl(r_Info.项目, '') || '：在' ||
                   Nvl(r_Info.科室, '[未定科室]') || r_Info.执行状态) > 1000 Then
          v_Text := v_Text || Chr(13) || Chr(10) || '... ...';
          Exit;
        Else
          v_Text := v_Text || Chr(13) || Chr(10) || '单据[' || Nvl(r_Info.No, '') || ']中的' || Nvl(r_Info.项目, '') || '：在' ||
                    Nvl(r_Info.科室, '[未定科室]') || r_Info.执行状态;
        End If;
      
        n_Count := n_Count + 1;
      End Loop;
    
      v_Text := Substr(v_Text, 3);
    Else
      For r_Info In (Select Distinct b.No, c.名称 As 项目, d.名称 As 科室, Decode(Nvl(b.执行状态, 0), 0, '未执行', 3, '正在执行') As 执行状态
                     From 病人医嘱记录 A, 病人医嘱发送 B, 诊疗项目目录 C, 部门表 D
                     Where a.病人id = n_病人id And a.主页id Is Null And (Nvl(a.婴儿, 0) = n_婴儿序号 Or n_婴儿序号 = -1) And
                           a.Id = b.医嘱id And b.执行状态 In (0, 3) And a.诊疗项目id = c.Id And b.执行部门id + 0 = d.Id And
                           a.诊疗类别 Not In ('4', '5', '6', '7') And Not (a.诊疗类别 In ('F', 'D') And a.相关id Is Not Null) And
                           (d.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or d.撤档时间 Is Null) And
                           Not (a.诊疗类别 = 'D' And Nvl(b.执行过程, 0) >= 2) And
                           (Not (a.诊疗类别 = 'Z' And Nvl(c.操作类型, '0') <> '0') Or a.诊疗类别 = 'Z' And c.操作类型 = '7')) Loop
        If Lengthb(v_Text || Chr(13) || Chr(10) || '单据[' || Nvl(r_Info.No, '') || ']中的' || Nvl(r_Info.项目, '') || '：在' ||
                   Nvl(r_Info.科室, '[未定科室]') || r_Info.执行状态) > 1000 Then
          v_Text := v_Text || Chr(13) || Chr(10) || '... ...';
          Exit;
        Else
          v_Text := v_Text || Chr(13) || Chr(10) || '单据[' || Nvl(r_Info.No, '') || ']中的' || Nvl(r_Info.项目, '') || '：在' ||
                    Nvl(r_Info.科室, '[未定科室]') || r_Info.执行状态;
        End If;
      
        n_Count := n_Count + 1;
      End Loop;
    
      v_Text := Substr(v_Text, 3);
    End If;
  
    If n_Count > 0 Then
      n_Count := 1;
    End If;
  Elsif n_检查方式 = 1 Then
    v_Text := '';
    --检查是否安装了血库
    Begin
      Select 1
      Into n_共享
      From zlSystems
      Where Trunc(编号 / 100) = 22 And 所有者 = Sys_Context('USERENV', 'CURRENT_SCHEMA');
    Exception
      When Others Then
        n_共享 := 0;
    End;
    v_Sql := 'select e.名称 项目, c.名称 As 部门, To_Char(a.扣率) As 扣率 ';
    v_Sql := v_Sql || ' from  部门表 c,血液收发记录 a, 收费项目目录 e, 血液配血记录 b, 病人医嘱记录 d ';
    v_Sql := v_Sql || ' where b.申请id = d.id  And d.诊疗类别 = :1 and d.医嘱状态<>4 and c.Id = a.库房id and Nvl(a.填写数量, 0) <> 0 And a.单据 = 6 And Mod(a.记录状态, 3) = 1 ';
    v_Sql := v_Sql || ' And a.发血状态 = 1 And a.审核人 Is Null ';
    v_Sql := v_Sql || ' and b.病人id = :2 and b.主页id = :3 And a.血液id = e.Id And a.配发id = b.Id and b.记录性质 + 0 = 1 And b.记录状态 in (1,2)';
    --共享安装：进行未发血液的检查
    If n_共享 = 1 Then
      Open c_Bool For v_Sql
        Using 'K',n_病人id, n_主页id;
      Loop
        Fetch c_Bool
          Into  v_项目, v_部门, v_扣率;
        Exit When c_Bool%NotFound;

        If v_Text Is Not Null Then
          If Instr(Chr(13) || Chr(10) || v_Text || Chr(13) || Chr(10),
                   Chr(13) || Chr(10) || Nvl(v_项目, '') || '：在' || Nvl(v_部门, '[未定血库]') ||
                    '未发血' || Chr(13) || Chr(10), 1) = 0 Then
            If Lengthb(v_Text || Chr(13) || Chr(10) ||  Nvl(v_项目, '') || '：在' ||
                       Nvl(v_部门, '[未定血库]') || '未发血') <= 1000 Then
              v_Text := v_Text || Chr(13) || Chr(10) ||  Nvl(v_项目, '') || '：在' ||
                        Nvl(v_部门, '[未定血库]') || '未发血';
            Else
              v_Text := v_Text || Chr(13) || Chr(10) || '... ...';
            End If;
          End If;
        Else
          v_Text :=  Nvl(v_项目, '') || '：在' || Nvl(v_部门, '[未定血库]') || '未发血';
        End If;
        n_Count := n_Count + 1;
      End Loop;
      Close c_Bool;
      If v_Text Is Not Null Then
        v_Text := '存在未发放的血液：' || Chr(13) || Chr(10) || Chr(13) || Chr(10) || v_Text;
      End If;
      v_Text := v_Text;
    End If;
    If n_Count > 0 Then
      n_Count := 1;
    End If;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功","isexist":' || n_Count || ',"notexecute_infor":"' ||
              Zljsonstr(v_Text) || '"}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Checkpatiexecute;
/
Create Or Replace Procedure Zl_Cissvr_Checkpativisitorin
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：根据病人ID和身份证号检查同一身份证只能对应一个建档病人
  --入参：Json_In:格式
  --  input
  --    pati_id              N   1  病人ID
  --    pati_pageid          N   1  主页id
  --出参: Json_Out,格式如下
  --  output
  --    code                 N   1   应答吗：0-失败；1-成功
  --    message              C   1   应答消息：失败时返回具体的错误信息
  --    isexist              N   1   是否存在（0-不存在 1-存在）
  ---------------------------------------------------------------------------
  n_病人id   病案主页.病人id%Type;
  n_主页id   病案主页.主页id%Type;
  d_出院日期 病案主页.出院日期%Type;
  n_Count    Number;
  j_Json     Pljson;
  j_In       Pljson;
  n_Isexist  Number(1);
Begin
  --解析入参
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');
  n_主页id := j_Json.Get_Number('pati_pageid');
  If Not n_主页id Is Null Then
    Select 出院日期 Into d_出院日期 From 病案主页 Where 病人id = n_病人id And 主页id = n_主页id;
    --存在出院时间，则判断该出院后是否存在就诊或住院数据
    If Not d_出院日期 Is Null Then
      --先判断住院
      Select Count(1) Into n_Count From 病案主页 Where 病人id = n_病人id And 入院日期 >= d_出院日期;
      If n_Count = 0 Then
        Begin
          --该过程病案、标准版均有。病案系统若单独安装没有病人挂号记录
          Execute Immediate 'Select Count(1) From 病人挂号记录 Where 病人id =:1  And 登记时间 >=:2 '
            Into n_Count
            Using n_病人id, d_出院日期;
        Exception
          When Others Then
            Null;
        End;
      End If;
    End If;
  End If;
  If Nvl(n_Count, 0) > 0 Then
    n_Isexist := 1;
  Else
    n_Isexist := 0;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功","isexist":' || n_Isexist || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Checkpativisitorin;
/
Create Or Replace Procedure Zl_Cissvr_Checkskinresult
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：判断医嘱是否标记了皮试结果
  --入参：Json_In:格式
  --  input
  --     advice_id          N 1 医嘱id
  --出参: Json_Out,格式如下
  --  output
  --    code                N 1 应答吗：0-失败；1-成功
  --    message             C 1 应答消息：失败时返回具体的错误信息
  --    skintest_info       N 1 皮试结果 -1表示无需皮试或免试；0表示还未标记皮试结果或未下达皮试医嘱；1表示阴性；2表示阳性
  ---------------------------------------------------------------------------
  j_Json     Pljson;
  j_In       Pljson;
  n_Skintest Number;
  v_Test     Varchar2(3000);
  v_挂号单   Varchar2(30);
  n_病人id   Number;
  n_主页id   Number;
  v_项目ids  Varchar2(32767);
  n_有效天数 Number;
  n_医嘱id   Number;
  v_Err_Msg  Varchar(2000);
  Err_Item Exception;
Begin
  --解析入参
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_医嘱id := j_Json.Get_Number('advice_id');

  n_有效天数 := zl_GetSysParameter(70);
  n_Skintest := -1;

  For R In (Select b.用法id, a.病人id, a.主页id, a.挂号单
            From 病人医嘱记录 A, 诊疗用法用量 B
            Where a.诊疗项目id = b.项目id And b.性质 = 0 And a.Id = n_医嘱id And a.诊疗类别 In ('5', '6')) Loop
    v_项目ids := v_项目ids || ',' || r.用法id;
    n_病人id  := r.病人id;
    n_主页id  := r.主页id;
    v_挂号单  := r.挂号单;
  End Loop;

  If v_项目ids Is Not Null Then
    v_项目ids := v_项目ids || ',';
    If v_挂号单 Is Null Then
      For X In (Select a.皮试结果, Nvl(b.标本部位, '阳性(+);阴性(-)') As 标本部位
                From 病人医嘱记录 A, 诊疗项目目录 B
                Where a.诊疗项目id = b.Id And a.病人id = n_病人id And a.主页id = n_主页id And
                      Instr(v_项目ids, ',' || a.诊疗项目id || ',') > 0 And a.开始执行时间 >= Trunc(Sysdate) - n_有效天数
                Order By a.开始执行时间 Desc) Loop
        --只循环一次
        v_Test := x.皮试结果;
        Exit;
      End Loop;
    Else
      For X In (Select a.皮试结果, Nvl(b.标本部位, '阳性(+);阴性(-)') As 标本部位
                From 病人医嘱记录 A, 诊疗项目目录 B
                Where a.诊疗项目id = b.Id And a.挂号单 = v_挂号单 And Instr(v_项目ids, ',' || a.诊疗项目id || ',') > 0 And
                      a.开始执行时间 >= Trunc(Sysdate) - n_有效天数
                Order By a.开始执行时间 Desc) Loop
        --只循环一次
        v_Test := x.皮试结果;
        Exit;
      End Loop;
    End If;
    If v_Test Is Null Then
      n_Skintest := 0;
    Elsif v_Test = '免试' Then
      n_Skintest := -1;
    Elsif Instr(v_Test, '-') > 0 Then
      n_Skintest := 1;
    Elsif Instr(v_Test, '+') > 0 Then
      n_Skintest := 2;
    End If;
  End If;

  Json_Out := '{"output":{"code":1,"message":"成功","skintest_info":' || Nvl(n_Skintest, 0) || '}}';

Exception
  When Err_Item Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr('-20101:[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]') || '"}}';
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Checkskinresult;
/
Create Or Replace Procedure Zl_Cissvr_Chkoutpatiexistorder
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:挂号时检查病人在挂号有效天数内是否存在医嘱
  --入参：Json_In:格式
  --    input
  --      pati_id                 N 1 病人id
  --      rgst_expidate           N 1 挂号有效天数
  --出参: Json_Out,格式如下
  --  output
  --    code              N   1   应答码：0-失败；1-成功
  --    message           C   1   应答消息：失败时返回具体的错误信息
  --    exist_order       N   1   是否存在医嘱
  ---------------------------------------------------------------------------

  j_Json Pljson;
  j_In   Pljson;

  n_病人id   病人医嘱记录.病人id%Type;
  n_有效天数 Number(10);
  n_Count    Number(1);

Begin
  --解析入参
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  n_病人id   := j_Json.Get_Number('pati_id');
  n_有效天数 := j_Json.Get_Number('rgst_expidate');

  If Nvl(n_病人id, 0) = 0 Then
    Json_Out := Zljsonout('未传入病人id，请检查！');
    Return;
  End If;

  Select Count(1)
  Into n_Count
  From 病人挂号记录 A, 病人医嘱记录 B
  Where a.病人id + 0 = b.病人id And a.No || '' = b.挂号单 And a.记录状态 = 1 And a.记录性质 = 1 And
        a.登记时间 - 0 >= Trunc(Sysdate) - n_有效天数 And a.病人id = n_病人id And Rownum < 2;

  Json_Out := '{"output":{"code":1,"message":"成功","exist_order":' || n_Count || '}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Chkoutpatiexistorder;
/
Create Or Replace Procedure Zl_Cissvr_Delerrbillinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------------------------------------------------------------
  --功能：删除医生站预约接收产生的异常记录，改为窗口处理
  --入参：json格式
  --Input
  --   rgst_no               C  1 挂号单
  --出参：json格式
  --Json_Out
  --   code                  N  1  应答码：0-失败；1-成功
  --   message               C  1  应答消息：失败时返回具体的错误信息
  -------------------------------------------------------------------------------------------------
  n_异常id 病人结算异常记录.Id%Type;
  v_挂号单 病人结算异常记录.预交单号%Type;
  j_Json   Pljson;
  j_In     Pljson;
Begin
  --解析入参
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  v_挂号单 := j_Json.Get_String('rgst_no');

  Select ID Into n_异常id From 病人结算异常记录 Where 预交单号 = v_挂号单 And 操作场景 = 4 And Rownum < 2;
  Zl_病人结算异常记录_Modify(n_异常id, 2, Null, Null, Null);
  Json_Out := '{"output":{"code":1,"message":"成功"}}';

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Delerrbillinfo;
/
Create Or Replace Procedure Zl_Cissvr_Deloutpativisitrec
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能: 退号成功后作废临床就诊登记记录,目前临床的就诊登记记录就是病人挂号记录，所以不用处理。
  --入参：Json_In:格式
  --input
  --  rgst_no     C  1  挂号单号,批量取消预约时会传入多个,如：U0000001,U0000002
  --出参: Json_Out,格式如下
  --output
  --  code        N 1 应答码：0-失败；1-成功
  --  message     C 1 应答消息：成功时返回成功信息,失败时返回具体的错误信息
  ----------------------------------------------------------------------------
  j_Json     Pljson;
  v_挂号单号 varchar2(4000);
  j_In       Pljson;
Begin
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  v_挂号单号 := j_Json.Get_String('rgst_no');

  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Deloutpativisitrec;
/
Create Or Replace Procedure Zl_Cissvr_Existadvice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:判断是否存在医嘱数据或判断指定挂号单是否已经开医嘱
  --入参：Json_In:格式
  -- input
  --   pati_id              N 1 病人ID
  --   pati_pageid          N   主页Id
  --   rgst_no              C 1 挂号单，多个用逗号分隔
  --   only_valid           N   只检查没有作废的医嘱
  --出参: Json_Out,格式如下
  --  output
  --    code               C 1 应答码：0-失败；1-成功
  --    message            C 1 应答消息：失败时返回具体的错误信息
  --    exist              N 1 是否存在，1-存在;0-不存在
  ---------------------------------------------------------------------------
  j_Json       PLJson;
  j_Json_Tmp   PLJson;
  n_病人id     Number(18);
  n_主页id     Number(18);
  v_挂号单     Varchar2(3000);
  n_Exist      Number(2);
  n_不检查作废 Number(1);

Begin
  --解析入参
  j_Json     := PLJson(Json_In);
  j_Json_Tmp := j_Json.Get_Pljson('input');

  n_病人id     := j_Json_Tmp.Get_Number('pati_id');
  n_主页id     := j_Json_Tmp.Get_Number('pati_pageid');
  v_挂号单     := j_Json_Tmp.Get_String('rgst_no');
  n_不检查作废 := Nvl(j_Json_Tmp.Get_Number('only_valid'), 0);

  If v_挂号单 Is Not Null Then
  
    Select Max(1)
    Into n_Exist
    From 病人医嘱记录 A
    Where (病人id + 0 = Nvl(n_病人id, 0) Or Nvl(n_病人id, 0) = 0) And
          挂号单 In (Select Column_Value As 挂号单 From Table(f_Str2List(v_挂号单))) And
          (n_不检查作废 = 0 Or n_不检查作废 = 1 And 医嘱状态 <> 4);
  
  Elsif j_Json_Tmp.Exist('pati_pageid') Then
    Select Max(存在)
    Into n_Exist
    From (Select 1 As 存在
           From 病人医嘱记录 A
           Where 病人id = n_病人id And Nvl(主页id, 0) = Nvl(n_主页id, 0) And (n_不检查作废 = 0 Or n_不检查作废 = 1 And 医嘱状态 <> 4) And
                 Rownum < 2);
  Else
    Select Max(存在)
    Into n_Exist
    From (Select 1 As 存在
           From 病人医嘱记录 A
           Where 病人id = n_病人id And (n_不检查作废 = 0 Or n_不检查作废 = 1 And 医嘱状态 <> 4) And Rownum < 2);
  End If;

  Json_Out := '{"output":{"code":1,"message":"' || '成功' || '","exist":' || Nvl(n_Exist, 0) || '}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || zlJsonStr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Existadvice;
/
Create Or Replace Procedure Zl_Cissvr_Existadvicesend
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：根据费用单据号及性质，判断是否存在医嘱发送数据
  --入参：Json_In:格式
  --  input
  --    fee_no              C 1 单据号
  --    send_no             N 1 发送号

  --出参: Json_Out,格式如下
  --  output
  --    code                N 1 应答吗：0-失败；1-成功
  --    message             C 1 应答消息：失败时返回具体的错误信息
  --    exsit               N 1 是否存在:1-存在;0-不存在
  ---------------------------------------------------------------------------
  j_Json   Pljson;
  j_In     Pljson;
  v_单据号 病人医嘱发送.No%Type;
  n_发送号 病人医嘱发送.发送号%Type;
  n_Exist  Number(1);

Begin
  --解析入参
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_发送号 := j_Json.Get_Number('send_no');
  v_单据号 := j_Json.Get_String('fee_no');

  Select Count(1) Into n_Exist From 病人医嘱发送 A Where a.No = v_单据号 And a.发送号 = n_发送号;

  Json_Out := '{"output":{"code":1,"message":"成功","exsit":' || n_Exist || '}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Existadvicesend;
/
Create Or Replace Procedure Zl_Cissvr_Existoutadvice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:医生是否下达了出院医嘱
  --入参：Json_In:格式
  -- input
  --   pati_id              N 1 病人ID
  --   pati_pageid          N 1 主页Id
  --出参: Json_Out,格式如下
  --  output
  --    code               N  1 应答码：0-失败；1-成功
  --    message            C  1 应答消息：失败时返回具体的错误信息
  --    is_out             N  1 是否已开出院医嘱 ：1-已开出院;0-未开出院
  --    out_advice_id      N  1 已经开了出院医嘱的，返回医嘱的ID
  ---------------------------------------------------------------------------
  j_In     Pljson;
  j_Json   Pljson;
  n_病人id 病人医嘱记录.病人id%Type;
  n_主页id 病人医嘱记录.主页id%Type;

  n_Tmp    Number(1);
  n_医嘱id 病人医嘱记录.Id%Type;

Begin
  --解析入参
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');
  n_主页id := j_Json.Get_Number('pati_pageid');

  If Nvl(n_病人id, 0) = 0 Or Nvl(n_主页id, 0) = 0 Then
    Json_Out := Zljsonout('必须传入医人id和主页id！');
    Return;
  End If;

  Select Max(a.Id), Count(1)
  Into n_医嘱id, n_Tmp
  From 病人医嘱记录 A, 病人变动记录 B, 病案主页 C, 诊疗项目目录 D
  Where a.病人id = n_病人id And a.主页id = n_主页id And a.医嘱状态 = 8 And a.病人id = b.病人id And a.主页id = b.主页id And
        a.开始执行时间 = b.开始时间 + 0 And b.开始原因 = 10 And b.病人id = c.病人id And b.主页id = c.主页id And c.状态 = 3 And d.类别 = 'Z' And
        d.操作类型 In ('5', '6', '11') And a.诊疗项目id = d.Id;

  Json_Out := '{"output":{"code":1,"message":"成功","is_out":' || n_Tmp || ',"out_advice_id":' || Nvl(n_医嘱id, 0) || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Existoutadvice;
/
Create Or Replace Procedure Zl_Cissvr_Getadviceannexinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：获取医嘱附费信息
  --入参：Json_In:格式
  --  input
  --    advice_ids                C 1 医嘱ID,多个用','分隔
  --    send_no                   N 0 发送号
  --    bill_no                   C 0 NO
  --    bill_prop                 N 1 记录性质
  --出参: Json_Out,格式如下
  --  output
  --    code                      N 1 应答吗：0-失败；1-成功
  --    message                   C 1 应答消息：失败时返回具体的错误信息
  --    advice_annex_list             [数组]每个医嘱id对应的附费单据信息
  --      advice_id               N 1 医嘱ID
  --      send_no                 N 1 发送号
  --      bill_no                 C   No
  --      bill_prop               N   记录性质
  ---------------------------------------------------------------------------
  j_Json Pljson;
  j_In   Pljson;

  v_No       病人医嘱附费.No%Type;
  n_记录性质 病人医嘱附费.记录性质%Type;
  n_发送号   病人医嘱附费.发送号%Type;

  v_医嘱ids Clob; --记录医嘱id
  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_医嘱id Collection_Type;
  I          Number;
  v_Jtmp     Varchar2(32767);
  c_Jtmp     Clob;
Begin
  --解析入参
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  v_医嘱ids  := j_Json.Get_String('advice_ids');
  n_发送号   := j_Json.Get_Number('send_no');
  v_No       := j_Json.Get_String('bill_no');
  n_记录性质 := j_Json.Get_Number('bill_prop');

  I := 0;
  While v_医嘱ids Is Not Null Loop
    If Length(v_医嘱ids) <= 4000 Then
      Col_医嘱id(I) := v_医嘱ids;
      v_医嘱ids := Null;
    Else
      Col_医嘱id(I) := Substr(v_医嘱ids, 1, Instr(v_医嘱ids, ',', 3980) - 1);
      v_医嘱ids := Substr(v_医嘱ids, Instr(v_医嘱ids, ',', 3980) + 1);
    End If;
    I := I + 1;
  End Loop;
  I := 0;
  For I In 0 .. Col_医嘱id.Count - 1 Loop
    For r_医嘱 In (Select a.记录性质, a.医嘱id, a.发送号, a.No
                 From 病人医嘱附费 A
                 Where a.医嘱id In (Select /*+cardinality(B,10) */
                                   Column_Value As 医嘱id
                                  From Table(f_Num2list(Col_医嘱id(I))) B) And (a.发送号 = n_发送号 Or n_发送号 Is Null) And
                       a.记录性质 = n_记录性质 And (a.No = v_No Or v_No Is Null)) Loop
    
      v_Jtmp := v_Jtmp || ',{"advice_id":' || r_医嘱.医嘱id;
      v_Jtmp := v_Jtmp || ',"send_no":' || r_医嘱.发送号;
      v_Jtmp := v_Jtmp || ',"bill_no":"' || r_医嘱.No || '"';
      v_Jtmp := v_Jtmp || ',"bill_prop":' || r_医嘱.记录性质;
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
    Json_Out := '{"output":{"code":1,"message":"成功","advice_annex_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","advice_annex_list":[' || c_Jtmp || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getadviceannexinfo;
/
Create Or Replace Procedure Zl_Cissvr_Getadviceannexnote
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：根据医嘱ID及元素名称获取医嘱附件内容
  --入参：Json_In:格式
  --  input
  --    advice_id           N 1 医嘱ID
  --    chn_name            C 1 诊治所见项目.中文名,比如：主刀医生科室
  --出参: Json_Out,格式如下
  --  output
  --    code                      N 1 应答吗：0-失败；1-成功
  --    message                   C 1 应答消息：失败时返回具体的错误信息
  --    note_list[]数据组
  --       annex_note               C 1 医嘱附件内容:病人医嘱附件.内容
  ---------------------------------------------------------------------------
  j_In       Pljson;
  j_Json     Pljson;
 
  n_医嘱id   病人医嘱附件.医嘱id%Type;
  v_所见项目 诊治所见项目.中文名%Type;
  v_Jtmp     Varchar2(32767);
 
Begin
  --解析入参
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');

  n_医嘱id   := j_Json.Get_Number('advice_id');
  v_所见项目 := j_Json.Get_String('chn_name');

  If Nvl(n_医嘱id, 0) = 0 Then
    Json_Out := Zljsonout('未传入医嘱id，请检查！');
    Return;
  End If;

  If Nvl(v_所见项目, '-') = '-' Then
    Json_Out := Zljsonout('未传入诊疗所见项目名称，请检查！');
    Return;
  End If;

  For r_内容 In (Select a.内容
               From 病人医嘱附件 A, 诊治所见项目 B
               Where a.要素id = b.Id And a.医嘱id = n_医嘱id And b.中文名 = v_所见项目) Loop
  
    v_Jtmp := v_Jtmp || ',{"annex_note":"' || Zljsonstr(r_内容.内容) || '"}';
  
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"成功","note_list":[' || Substr(v_Jtmp, 2) || ']}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getadviceannexnote;
/
Create Or Replace Procedure Zl_Cissvr_Getadvicedefinedinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：获取医嘱内容定义的相关信息
  --入参：Json_In:格式
  --无
  --出参: Json_Out,格式如下
  --  output
  --    code                      N 1 应答吗：0-失败；1-成功
  --    message                   C 1 应答消息：失败时返回具体的错误信息
  --    item_list[]
  --      clinic_type             C 1 诊疗类别
  --      advice_note             C 1 医嘱内容
  ---------------------------------------------------------------------------

  v_Jtmp Varchar2(32767);

Begin
  For r_医嘱内容 In (Select 诊疗类别, 医嘱内容 From 医嘱内容定义 Order By 诊疗类别) Loop
  
    v_Jtmp := v_Jtmp || ',{"clinic_type":"' || r_医嘱内容.诊疗类别 || '"';
    v_Jtmp := v_Jtmp || ',"advice_note":"' || Zljsonstr(r_医嘱内容.医嘱内容) || '"';
    v_Jtmp := v_Jtmp || '}';
  
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getadvicedefinedinfo;
/
Create Or Replace Procedure Zl_Cissvr_Getadviceexcutnums
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：获取医嘱发送的已执行数量，根据医嘱ID查询医嘱发送信息
  --入参      json
  --input     
  --  item_list                 数据组
  --    advice_id               N 1 医嘱ID
  --    bill_no                 C 1 NO
  --    bill_prop               N 1 记录性质
  --出参      json
  --output
  --  code                      C 1 应答码：0-失败；1-成功
  --  message                   C 1 应答消息：失败时返回具体的错误信息
  --  item_list[]数据组
  --    advice_id               N 1 医嘱ID
  --    bill_no                 C 1 No
  --    fee_item_id             N 1 收费细目ID
  --    execute_num             N 1 已执行数
  --说明：注意，即使已执行数为0也要返回
  ---------------------------------------------------------------------------
  j_Json     Pljson;
  j_List     Pljson_List;
  j_Json_Tmp Pljson;

  n_医嘱id   病人医嘱发送.医嘱id%Type;
  v_No       病人医嘱发送.No%Type;
  n_记录性质 病人医嘱发送.记录性质%Type;
  v_Jtmp     Varchar2(32767);
  c_Jtmp     Clob;
Begin
  j_Json_Tmp := Pljson(Json_In);
  j_Json     := j_Json_Tmp.Get_Pljson('input');
  j_List     := j_Json.Get_Pljson_List('item_list');
  
  For I In 1 .. j_List.Count Loop
    j_Json_Tmp := Pljson();
    j_Json_Tmp := Pljson(j_List.Get(I));
    n_医嘱id   := j_Json_Tmp.Get_Number('advice_id');
    v_No       := j_Json_Tmp.Get_String('bill_no');
    n_记录性质 := j_Json_Tmp.Get_Number('bill_prop');
  
    --Zl_医嘱执行计价_修正(
    --  医嘱id_In   病人医嘱执行.医嘱id%Type,
    --  No_In       病人医嘱发送.No%Type,
    --  记录性质_In 病人医嘱发送.记录性质%Type
    Zl_医嘱执行计价_修正(n_医嘱id, v_No, n_记录性质);
  
    --1.存在医嘱执行计价的,则以医嘱执行计价为准(但不能包含:检查;检验;手术;麻醉及输血)
    --2.病人医嘱发送.执行状态=1（完成执行）时，准退数为0，不再根据医嘱执行计价来统计
    For c_医嘱 In (Select 医嘱id, NO, 收费细目id, Sum(已执行数) As 已执行数
                 From (Select b.医嘱id, b.No, c.收费细目id, Decode(b.执行状态, 1, 1, Decode(c.执行状态, 1, 1, 0)) * c.数量 As 已执行数
                        From 病人医嘱发送 B, 医嘱执行计价 C, 病人医嘱记录 M
                        Where b.医嘱id = c.医嘱id And b.发送号 = c.发送号 And b.医嘱id = m.Id And
                              Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And b.医嘱id = n_医嘱id And b.No = v_No And
                              b.记录性质 = n_记录性质)
                 Group By 医嘱id, NO, 收费细目id) Loop
    
      v_Jtmp := v_Jtmp || ',{"advice_id":' || c_医嘱.医嘱id;
      v_Jtmp := v_Jtmp || ',"bill_no":"' || c_医嘱.No || '"';
      v_Jtmp := v_Jtmp || ',"fee_item_id":' || c_医嘱.收费细目id;
      v_Jtmp := v_Jtmp || ',"execute_num":' || Zljsonstr(c_医嘱.已执行数, 1);
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
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getadviceexcutnums;
/
Create Or Replace Procedure Zl_Cissvr_Getadviceexestatus
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：获取医嘱发送的执行状态
  --入参：Json_In:格式
  --  input
  --    advice_id               N 1 医嘱ID
  --    send_no                 N 1 发送号

  --出参: Json_Out,格式如下
  --  output
  --    code                      N 1 应答吗：0-失败；1-成功
  --    message                   C 1 应答消息：失败时返回具体的错误信息
  --    exe_status                N 1 执行状态
  ---------------------------------------------------------------------------
  j_In       Pljson;
  j_Json     Pljson;
  n_发送号   病人医嘱发送.发送号%Type;
  n_医嘱id   病人医嘱发送.医嘱id%Type;
  n_执行状态 病人医嘱发送.执行状态%Type;

Begin
  --解析入参
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_医嘱id := j_Json.Get_Number('advice_id');
  n_发送号 := j_Json.Get_Number('send_no');

  If Nvl(n_医嘱id, 0) = 0 Or Nvl(n_发送号, 0) = 0 Then
    Json_Out := Zljsonout('未传入医嘱id或发送号，请检查！');
    Return;
  End If;

  Select Max(执行状态)
  Into n_执行状态
  From 病人医嘱发送
  Where 发送号 = n_发送号 And
        医嘱id In (Select ID From 病人医嘱记录 Where (ID = n_医嘱id Or 相关id = n_医嘱id) And 诊疗类别 In ('C', 'D'));
  Json_Out := '{"output":{"code":1,"message":"成功","exe_status":' || Nvl(n_执行状态, 0) || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getadviceexestatus;
/
Create Or Replace Procedure Zl_Cissvr_Getadvicefeestatus
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：获取医嘱发送的计费状态
  --入参：Json_In:格式
  --  input
  --   advice_id            N 1 医嘱ID
  --   send_no              N 1 发送号
  --   isalone_exe          N 1 是否独立执行:1-独立执行;0-非独立执行

  --出参: Json_Out,格式如下
  --  output
  --    code                      N 1 应答吗：0-失败；1-成功
  --    message                   C 1 应答消息：失败时返回具体的错误信息
  --    charge_status             C 1 计费状态:多个用逗号分离其中-1=无需计费,1=已计费,0=未计费,对于门诊单据，2=部分收费,3=全部收费
  ---------------------------------------------------------------------------
  j_In   Pljson;
  j_Json Pljson;

  n_医嘱id   病人医嘱记录.Id%Type;
  n_相关id   病人医嘱记录.相关id%Type;
  n_独立执行 Number(1);
  n_发送号   病人医嘱发送.发送号%Type;
  v_记费状态 Varchar2(100);
  v_诊疗类别 病人医嘱记录.诊疗类别%Type;

Begin
  --解析入参
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');

  n_医嘱id   := j_Json.Get_Number('advice_id');
  n_发送号   := j_Json.Get_Number('send_no');
  n_独立执行 := Nvl(j_Json.Get_Number('isalone_exe'), 0);

  If Nvl(n_医嘱id, 0) = 0 Or Nvl(n_发送号, 0) = 0 Then
    Json_Out := Zljsonout('未传入医嘱id或发送号，请检查！');
    Return;
  End If;

  Select Decode(a.诊疗类别, 'D', Nvl(a.相关id, a.Id), a.相关id), 诊疗类别
  Into n_相关id, v_诊疗类别
  From 病人医嘱记录 A
  Where a.Id = n_医嘱id;

  If n_独立执行 = 1 Then
    Select Distinct 计费状态 Into v_记费状态 From 病人医嘱发送 Where 医嘱id = n_医嘱id And 发送号 = n_发送号;
  Else
  
    Select Distinct f_List2str(Cast(Collect(To_Char(计费状态)) As t_Strlist))
    Into v_记费状态
    From 病人医嘱发送
    Where 医嘱id In (Select ID From 病人医嘱记录 Where (ID = n_相关id Or 相关id = n_相关id) And 诊疗类别 = v_诊疗类别) And 发送号 = n_发送号;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功","charge_status":"' || v_记费状态 || '"}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getadvicefeestatus;
/
Create Or Replace Procedure Zl_Cissvr_Getadviceids
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:获取病人的医嘱IDs
  --入参：Json_In:格式
  -- input
  --   advice_register_no   C 1 挂号单号:挂号单或病人必传其中一个条件
  --   pati_id              N 1 病人ID:挂号单或病人必传其中一个条件
  --   advice_starttime     C 0 开始的开嘱时间,格式:yyyy-mm-dd hh24:mi:ss
  --   advice_endtime       C 0 结束的开嘱时间,格式:yyyy-mm-dd hh24:mi:ss
  --   isgetlast_adviceid   N 0 是否获取最后一次医嘱id,1-获取最后一次医嘱id;0-全部获取

  --出参: Json_Out,格式如下
  --  output
  --    code               C 1 应答码：0-失败；1-成功
  --    message            C 1 应答消息：失败时返回具体的错误信息
  --    advice_ids         C 1 入参:isgetlast_adviceid=1时，返回最后一次医嘱id,否则返回满足条件的所有医嘱id
  ---------------------------------------------------------------------------
  j_In       Pljson;
  j_Json     Pljson;
  n_病人id   Number(18);
  v_挂号单   Varchar2(100);
  d_开始时间 Date;
  d_结束时间 Date;
  n_最后一次 Number(1);
  v_医嘱ids  Varchar2(32767);
Begin
  --解析入参
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  n_病人id   := j_Json.Get_Number('pati_id');
  v_挂号单   := j_Json.Get_String('advice_register_no');
  d_开始时间 := To_Date(j_Json.Get_String('advice_starttime'), 'yyyy-mm-dd hh24:mi:ss');
  d_结束时间 := To_Date(j_Json.Get_String('advice_endtime'), 'yyyy-mm-dd hh24:mi:ss');
  n_最后一次 := Nvl(j_Json.Get_Number('isgetlast_adviceid'), 0);

  If Nvl(n_病人id, 0) = 0 And Nvl(v_挂号单, '-') = '-' Then
    Json_Out := Zljsonout('未传入挂号单或病人id！');
    Return;
  End If;

  If Nvl(n_病人id, 0) <> 0 Then
    --根据病人id查询医嘱id
    If n_最后一次 = 1 Then
      Select Max(ID)
      Into v_医嘱ids
      From 病人医嘱记录
      Where 病人id = n_病人id And 开嘱时间 Between Nvl(d_开始时间, 开嘱时间) And Nvl(d_结束时间, 开嘱时间);
    Else
      For r_医嘱 In (Select ID
                   From 病人医嘱记录
                   Where 病人id = n_病人id And 开嘱时间 Between Nvl(d_开始时间, 开嘱时间) And Nvl(d_结束时间, 开嘱时间)) Loop
        v_医嘱ids := v_医嘱ids || ',' || r_医嘱.Id;
      End Loop;
    End If;
  Else
    --根据挂号单查询医嘱id
    If n_最后一次 = 1 Then
      Select Max(ID)
      Into v_医嘱ids
      From 病人医嘱记录 M
      Where 挂号单 = v_挂号单 And 开嘱时间 Between Nvl(d_开始时间, 开嘱时间) And Nvl(d_结束时间, 开嘱时间);
    Else
      For r_医嘱 In (Select ID
                   From 病人医嘱记录
                   Where 挂号单 = v_挂号单 And 开嘱时间 Between Nvl(d_开始时间, 开嘱时间) And Nvl(d_结束时间, 开嘱时间)) Loop
        v_医嘱ids := v_医嘱ids || ',' || r_医嘱.Id;
      End Loop;
    End If;
  End If;
  v_医嘱ids := Substr(v_医嘱ids, 2);
  Json_Out  := '{"output":{"code":1,"message":"成功","advice_ids":"' || v_医嘱ids || '"}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getadviceids;
/
Create Or Replace Procedure Zl_Cissvr_Getadviceidsfromdiag
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：根据诊断ID获取对应的医嘱Ids
  --入参：Json_In:格式
  --  input
  --    diag_ids                  C 1 诊断id,多个用逗号分离
  --出参: Json_Out,格式如下
  --  output
  --    code                      N 1 应答吗：0-失败；1-成功
  --    message                   C 1 应答消息：失败时返回具体的错误信息
  --    advice_ids                C 1 医嘱ids,多个用逗号分离
  ---------------------------------------------------------------------------
  j_In      Pljson;
  j_Json    Pljson;
  I         Number;
  c_诊断ids Clob;
  c_医嘱ids Varchar2(32767);
  v_医嘱ids Varchar2(32767);
  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_诊断id Collection_Type;

Begin
  --解析入参
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');
  Begin
    c_诊断ids := j_Json.Get_Clob('diag_ids');
  Exception
    When Others Then
      Json_Out := Zljsonout('未传入诊断id，请检查！');
      Return;
  End;

  If Nvl(c_诊断ids, '-') = '-' Then
    Json_Out := Zljsonout('未传入诊断id，请检查！');
    Return;
  End If;

  I := 0;
  While c_诊断ids Is Not Null Loop
    If Length(c_诊断ids) <= 4000 Then
      Col_诊断id(I) := c_诊断ids;
      c_诊断ids := Null;
    Else
      Col_诊断id(I) := Substr(c_诊断ids, 1, Instr(c_诊断ids, ',', 3980) - 1);
      c_诊断ids := Substr(c_诊断ids, Instr(c_诊断ids, ',', 3980) + 1);
    End If;
    I := I + 1;
  End Loop;
  I := 0;
  For I In 0 .. Col_诊断id.Count - 1 Loop
    Select f_List2str(Cast(Collect(To_Char(医嘱id)) As t_Strlist))
    Into v_医嘱ids
    From (With 医嘱数据 As (Select /*+cardinality(b,10)*/
                         a.医嘱id
                        From 病人诊断医嘱 A, Table(f_Num2list(Col_诊断id(I))) B
                        Where a.诊断id = b.Column_Value)
           Select a.Id As 医嘱id
           From 病人医嘱记录 A, 医嘱数据 B
           Where a.Id = b.医嘱id
           Union
           Select a.Id
           From 病人医嘱记录 A, 医嘱数据 B
           Where a.相关id = b.医嘱id);
  
  
    c_医嘱ids := c_医嘱ids || ',' || v_医嘱ids;
  End Loop;
  c_医嘱ids := Substr(c_医嘱ids, 2);

  Json_Out := '{"output":{"code":1,"message":"成功","advice_ids":"' || c_医嘱ids || '"}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getadviceidsfromdiag;
/
Create Or Replace Procedure Zl_Cissvr_Getadviceinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:获取医嘱信息
  --入参：Json_In:格式
  -- input
  --   query_type                   N 0 查询类型：0:查询基本信息；1:查询基本信息+扩展信息
  --   advice_ids                   C 0 多个医嘱ID，可能是药嘱，也可能是主医嘱（给药途径）,用“,”分隔
  --   rgst_no                      C 0 挂号单号:挂号单或病人ID或医嘱ID必传其中一个条件
  --   pati_id                      N 0 病人ID:挂号单或病人ID或医嘱ID必传其中一个条件
  --   pati_pageid                  N 0 主页Id
  --出参: Json_Out,格式如下
  --  output
  --    code                        C 1 应答码：0-失败；1-成功
  --    message                     C 1 应答消息：失败时返回具体的错误信息
  --    advice_list      [数组]每个医嘱信息
  --      advice_id                 N    id
  --      advice_related_id         N    相关id
  --      pati_id                   N    病人id
  --      pati_pageid               N    主页id
  --      pati_source               N    病人来源
  --      advice_statu              N    医嘱状态:-1-未生效的暂存医嘱；1-新开；2-校对疑问；3-已校对；4-已作废；5-已重整；6-已暂停；7-已启用；8-已停止；9-已确认停止
  --      serial_num                N    序号
  --      advice_day                N    天数
  --      advice_dosage             N    单次用量
  --      oper_type                 C    操作类型:诊疗项目目录.操作类型
  --      clinic_type               C    诊疗类别
  --      advice_exe_properties     N    执行性质
  --      advice_exe_sign           N    执行标记
  --      effective_time            N    医嘱期效
  --      advice_record_time        D    开嘱时间
  --      advice_doctor             C    开嘱医生
  --      advice_purpose            C    用药目的
  --      advice_reason             C    用药理由
  --      advice_taboonote          C    禁忌药品说明
  --      advice_doctor_note        C    医生嘱托
  --      rcpdtl_excs_desc          C    超量说明
  --      advice_audit_result       N    审查结果
  --      advice_audit_sign         N    药师审核标志
  --      advice_audit_time         D    药师审核时间
  --      advice_interval_unit      C    间隔单位
  --      advice_frequency          C    执行频次
  --      advice_frequency_times    N    频率次数
  --      advice_frequency_interval N    频率间隔
  --      advice_exetime_plane      C    执行时间方案
  --      advice_begintime          D    开始执行时间
  --      advice_endtime            D    执行终止时间
  --      rgst_no                   C   挂号单号
  --      advice_receipt_name       C   配方名称
  --      advice_receipt_issecret   N   是否保密
  --      advice_note               C   医嘱内容
  --      advice_cisitem_id         N   诊疗项目id
  --      advice_item_id            N   收费细目id
  --      pati_deptid               N   病人科室id
  --      pati_name                 C   姓名
  --      pati_sex                  C   性别
  --      pati_age                  C   年龄
  --      advice_audit_reason       C   药师审核原因
  --      skintest_info             C   皮试结果

  --      total_qunt                N    总给予量
  --      Total_qunt_unit           C    总量:代单位的总量(总给予量+单位)
  --      single                    C    单量:单次用量+单位
  --      toxicity_type             C    毒理分类:药品有效，药品特性.毒理分类
  --      advice_dept_id            N    开嘱科室ID
  --      advice_stop_doctor        C    停嘱医生
  --      advice_stop_nurse         C    停嘱护士
  --      advice_stoptime           C    停嘱时间,格式：yyyy-mm-dd hh24:mi:ss
  --      advice_stoptime_confirm   C    确认停嘱时间:格式:yyyy-mm-dd hh24:mi:ss
  --      order_chk_nurse           C    校对护士
  --      order_chk_time            C    校对时间:yyyy-mm-dd hh24:mi;ss
  --      lastexe_time              D    上次执行时间
  --      usage                     C    用法
  --      emergency_tag             N    紧急标志:0-普通;1-紧急;2-补录(对门诊无效)
  --      is_charge_verfy           N   是否费用审核:1-审核;0-未审核
  --      baby_num                  N   婴儿序号
  --      valuation_nature          N   计价特性:0-正常计价；1-不计价；2-手工计价
  --      advice_exedept_id         N   执行科室ID:
  --      advice_exedept_name       C   执行科室名称
  --      testtube_code             C   试管编码:诊疗项目目录.试管编码
  --      hide_print                N   屏蔽打印
  --      prerequisite_id           N   前提ID
  --      is_staff_sig              N   是否签名:1-签名;0-未签名
  --      rpt_id                    N   病历报告id
  --      consult_statu             N   查阅状态:0-未查阅,1-已查阅
  --      advice_verfy_statu        N   审核状态:Null-无需审核，1-待审核，2-审核通过，3-审核未通过，4－血库待接收，5－血库配血中，6-血库停止配血，7-输血初审核待签发
  --      apply_num                 N   申请序号
  --      cisitem_unit              C   诊疗项目计算单位
  --      pati_wardarea_id          N   当前病区ID
  --      decoction_method          C   煎法
  ---------------------------------------------------------------------------
  j_In     Pljson;
  j_Json   Pljson;
  v_医嘱id Clob; --记录医嘱id
  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_医嘱id Collection_Type;
  I          Number;

  v_挂号单    Varchar2(50);
  n_病人id    Number;
  n_主页id    Number;
  n_皮试获取  Number; --null 未提取游表，1-已提取，2-已提取且有数据
  v_皮试结果  Varchar2(200);
  v_Pre病人   Varchar2(600);
  n_查询类型  Number(1);
  v_Adviceout Varchar2(32767);
  v_Temp      Varchar2(32767);
  c_Jtmp      Clob;
  n_相关id    病人医嘱记录.Id%Type;
  v_煎法      Varchar2(50);

  Cursor Ctest煎法(相关id_In 病人医嘱记录.Id%Type) Is
    Select b.名称
    From 病人医嘱记录 A, 诊疗项目目录 B
    Where a.诊疗项目id = b.Id And a.相关id = 相关id_In And a.诊疗类别 = 'E';
  r_煎法 Ctest煎法%RowType;

  Cursor Ctest(挂号单_In 病人医嘱记录.挂号单%Type) Is
    Select a.皮试结果, a.药名id, a.药品id
    From (Select a.皮试结果, c.项目id As 药名id, 0 As 药品id, a.开始执行时间
           From 病人医嘱记录 A, 诊疗用法用量 C
           Where a.诊疗项目id = c.用法id And Nvl(c.性质, 0) = 0 And Nvl(a.医嘱状态, 0) = 8 And a.皮试结果 Is Not Null And a.挂号单 = 挂号单_In
           Union All
           Select a.皮试结果, b.药名id, b.药品id, a.开始执行时间
           From 病人医嘱记录 A, 药品规格 B, 药品用法用量 C
           Where a.诊疗项目id = c.用法id And b.药品id = c.药品id And Nvl(c.性质, 0) = 0 And Nvl(a.医嘱状态, 0) = 8 And a.皮试结果 <> '免试' And
                 a.挂号单 = 挂号单_In
           Union All
           Select a.皮试结果, c.项目id As 药名id, 0 As 药品id, a.开始执行时间
           From 病人医嘱记录 A, 诊疗用法用量 C
           Where a.诊疗项目id = c.用法id And Nvl(c.性质, 0) = 0 And a.皮试结果 = '免试' And a.挂号单 = 挂号单_In) A
    Order By a.开始执行时间 Desc;
  Type t_Test Is Table Of Ctest%RowType;
  Rs_Test t_Test;

  Cursor Ctest住院
  (
    病人id_In 病人医嘱记录.病人id%Type,
    主页id_In 病人医嘱记录.主页id%Type
  ) Is
    Select a.皮试结果, a.药名id, a.药品id
    From (Select a.皮试结果, c.项目id As 药名id, 0 As 药品id, a.开始执行时间
           From 病人医嘱记录 A, 诊疗用法用量 C
           Where a.诊疗项目id = c.用法id And Nvl(c.性质, 0) = 0 And Nvl(a.医嘱状态, 0) = 8 And a.皮试结果 Is Not Null And
                 a.病人id = 病人id_In And a.主页id = 主页id_In
           Union All
           Select a.皮试结果, b.药名id, b.药品id, a.开始执行时间
           From 病人医嘱记录 A, 药品规格 B, 药品用法用量 C
           Where a.诊疗项目id = c.用法id And b.药品id = c.药品id And Nvl(c.性质, 0) = 0 And Nvl(a.医嘱状态, 0) = 8 And a.皮试结果 <> '免试' And
                 a.病人id = 病人id_In And a.主页id = 主页id_In
           Union All
           Select a.皮试结果, c.项目id As 药名id, 0 As 药品id, a.开始执行时间
           From 病人医嘱记录 A, 诊疗用法用量 C
           Where a.诊疗项目id = c.用法id And Nvl(c.性质, 0) = 0 And a.皮试结果 = '免试' And a.病人id = 病人id_In And a.主页id = 主页id_In) A
    Order By a.开始执行时间 Desc;

  Procedure Pp_Test
  (
    N药名id    Number,
    v_皮试结果 Out Varchar2
  ) Is
  Begin
    v_皮试结果 := Null;
    For I In 1 .. Rs_Test.Count Loop
      If Rs_Test(I).药名id = N药名id Then
        v_皮试结果 := Rs_Test(I).皮试结果;
      End If;
    End Loop;
  End Pp_Test;
Begin
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');
  If j_Json.Exist('advice_ids') Then
    v_医嘱id := j_Json.Get_Clob('advice_ids');
  End If;
  n_病人id   := j_Json.Get_Number('pati_id');
  n_主页id   := Nvl(j_Json.Get_Number('pati_pageid'), 0);
  v_挂号单   := j_Json.Get_String('rgst_no');
  n_查询类型 := Nvl(j_Json.Get_Number('query_type'), 0);

  If v_医嘱id Is Null Then
    If Nvl(n_病人id, 0) = 0 And Nvl(v_挂号单, '-') = '-' Then
      Json_Out := '{"output":{"code":0,"message":"未传入任何查询条件，请检查！"}}';
      Return;
    End If;
  
    If Nvl(n_病人id, 0) <> 0 Then
      Select f_List2str(Cast(Collect(To_Char(ID)) As t_Strlist))
      Into v_医嘱id
      From 病人医嘱记录
      Where 病人id = n_病人id And Decode(n_主页id, 0, 0, 主页id) = Decode(n_主页id, 0, 0, n_主页id);
    Else
      Select f_List2str(Cast(Collect(To_Char(ID)) As t_Strlist))
      Into v_医嘱id
      From 病人医嘱记录
      Where 挂号单 = v_挂号单;
    End If;
  End If;

  --将 v_医嘱id 串组装成不超过4000 的集合串，防止使用 f_Num2list 参数超长
  I := 0;
  While v_医嘱id Is Not Null Loop
    If Length(v_医嘱id) <= 4000 Then
      Col_医嘱id(I) := v_医嘱id;
      v_医嘱id := Null;
    Else
      Col_医嘱id(I) := Substr(v_医嘱id, 1, Instr(v_医嘱id, ',', 3980) - 1);
      v_医嘱id := Substr(v_医嘱id, Instr(v_医嘱id, ',', 3980) + 1);
    End If;
    I := I + 1;
  End Loop;

  v_Temp := Null;

  For I In 0 .. Col_医嘱id.Count - 1 Loop
    For r_医嘱 In (Select /*+cardinality(j,10)*/
                 Distinct a.Id, a.相关id, a.序号, g.序号 As 药品序号, a.病人id, a.主页id, a.病人来源, a.天数, a.单次用量, a.执行性质, a.执行标记, a.医嘱期效,
                          a.开嘱时间, a.开嘱医生, a.用药目的, a.用药理由, a.禁忌药品说明, a.医生嘱托, a.超量说明, a.审查结果, a.药师审核标志, a.药师审核时间, a.间隔单位,
                          a.执行频次, a.频率次数, a.频率间隔, a.执行时间方案, a.开始执行时间, a.执行终止时间, a.挂号单, a.医嘱内容, a.诊疗项目id, a.收费细目id,
                          a.病人科室id, a.姓名, a.性别, a.年龄, a.药师审核原因, a.皮试结果,
                          Decode(c.类别, '7',
                                  Decode(Nvl(g.配方id, 0), 0,
                                          Decode(Nvl(v.诊断id, 0), 0, '', Substr(g.医嘱内容, 1, Instr(g.医嘱内容, ':') - 1)),
                                          Nvl(n.名称, '')), '') As 配方名称,
                          Decode(c.类别, '7',
                                  Decode(Nvl(g.配方id, 0), 0,
                                          Decode(Nvl(v.诊断id, 0), 0, 0, Decode(Instr(g.医嘱内容, '(保密配方)'), 0, 0, 1)),
                                          Nvl(n.是否保密, 0)), 0) As 是否保密, a.总给予量,
                          Decode(a.总给予量, Null, Null,
                                  Decode(a.诊疗类别, 'E', Decode(p.操作类型, '4', a.总给予量 || '付', a.总给予量 || p.计算单位), '4',
                                          a.总给予量 || c.计算单位, '5', Round(a.总给予量 / d.住院包装, 5) || d.住院单位, '6',
                                          Round(a.总给予量 / d.住院包装, 5) || d.住院单位, a.总给予量 || p.计算单位)) As 总量,
                          Decode(a.单次用量, Null, Null, a.单次用量 || Decode(a.诊疗类别, '4', c.计算单位, p.计算单位)) As 单量, a.诊疗类别, p.操作类型,
                          e.毒理分类, a.开嘱科室id, a.停嘱医生, a.确认停嘱护士, a.停嘱时间, a.确认停嘱时间, a.校对护士, a.校对时间, a.上次执行时间, a.紧急标志,
                          a.是否费用审核, a.婴儿 As 婴儿序号, a.计价特性, a.执行科室id, b.名称 As 执行科室名称, p.试管编码, a.屏蔽打印, a.前提id,
                          Decode(f.签名id, Null, 0, 1) As 是否签名, h.查阅状态, a.审核状态, a.申请序号, p.计算单位 As 诊疗项目计算单位, m.当前病区id,
                          h.病历id As 报告id, a.医嘱状态, Nvl(P1.名称, p.名称) As 用法
                 From 病人医嘱记录 A, 部门表 B, 收费项目目录 C, 药品规格 D, 药品特性 E, 病人医嘱状态 F, 病人医嘱记录 G, 病人医嘱报告 H, 病人中医诊断记录 V, 诊疗项目目录 N,
                      Table(f_Num2list(Col_医嘱id(I))) J, 病案主页 M, 诊疗项目目录 P, 病人医嘱记录 A1, 诊疗项目目录 P1
                 Where Nvl(a.相关id, a.Id) = g.Id And a.收费细目id = c.Id(+) And g.配方id = n.Id(+) And a.执行科室id = b.Id(+) And
                       a.诊疗项目id = e.药名id(+) And a.收费细目id = d.药品id(+) And a.Id = h.医嘱id(+) And a.Id = f.医嘱id(+) And
                       g.Id = v.His医嘱id(+) And a.病人id = m.病人id(+) And a.主页id = m.主页id(+) And a.诊疗项目id = p.Id(+) And
                       (a.Id = j.Column_Value Or a.相关id = j.Column_Value) And f.操作类型 = 1 And a.开始执行时间 Is Not Null And
                       Nvl(a.医嘱状态, 0) <> -1 And a.相关id = A1.Id(+) And A1.诊疗项目id = P1.Id(+)
                 Order By a.相关id, a.Id) Loop
    
      --不重复的医嘱ID才加入到出参
      If v_Adviceout Is Null Or Instr(',' || v_Adviceout || ',', ',' || r_医嘱.Id || ',') = 0 Then
        If v_Adviceout Is Null Then
          v_Adviceout := r_医嘱.Id;
        Else
          v_Adviceout := v_Adviceout || ',' || r_医嘱.Id;
        End If;
        --获取病人皮试信息，可能是多病人
        If Nvl(n_皮试获取, 0) = 0 Or
           (Nvl(v_Pre病人, '*') <> r_医嘱.病人id || '_' || r_医嘱.主页id || '_' || r_医嘱.挂号单 And v_Pre病人 Is Not Null) Then
        
          n_皮试获取 := 1;
          If r_医嘱.挂号单 Is Not Null Then
            Open Ctest(r_医嘱.挂号单);
            Fetch Ctest Bulk Collect
              Into Rs_Test;
            Close Ctest;
          
            If Rs_Test.Count > 0 Then
              n_皮试获取 := 2;
            End If;
          Elsif r_医嘱.主页id Is Not Null Then
            Open Ctest住院(r_医嘱.病人id, r_医嘱.主页id);
            Fetch Ctest住院 Bulk Collect
              Into Rs_Test;
            Close Ctest住院;
          
            If Rs_Test.Count > 0 Then
              n_皮试获取 := 2;
            End If;
          End If;
        
          v_Pre病人 := r_医嘱.病人id || '_' || r_医嘱.主页id || '_' || '_' || r_医嘱.挂号单;
        End If;
      
        v_Temp := v_Temp || ',{"advice_id":' || r_医嘱.Id;
        v_Temp := v_Temp || ',"advice_related_id":' || Nvl(r_医嘱.相关id, 0);
        v_Temp := v_Temp || ',"pati_id":' || Nvl(r_医嘱.病人id, 0);
        v_Temp := v_Temp || ',"pati_pageid":' || Nvl(r_医嘱.主页id, 0);
        v_Temp := v_Temp || ',"pati_source":' || Nvl(r_医嘱.病人来源, 0);
        v_Temp := v_Temp || ',"advice_statu":' || Nvl(r_医嘱.医嘱状态, 0);
      
        v_Temp := v_Temp || ',"serial_num":' || Nvl(r_医嘱.序号, 0);
        v_Temp := v_Temp || ',"advice_day":' || Nvl(r_医嘱.天数, 0);
        v_Temp := v_Temp || ',"advice_dosage":' || Zljsonstr(r_医嘱.单次用量, 1);
        v_Temp := v_Temp || ',"oper_type":"' || Zljsonstr(r_医嘱.操作类型) || '"';
        v_Temp := v_Temp || ',"clinic_type":"' || Zljsonstr(r_医嘱.诊疗类别) || '"';
      
        v_Temp := v_Temp || ',"advice_exe_properties":' || Nvl(r_医嘱.执行性质, 0);
        v_Temp := v_Temp || ',"advice_exe_sign":' || Nvl(r_医嘱.执行标记, 0);
        v_Temp := v_Temp || ',"effective_time":' || Nvl(r_医嘱.医嘱期效, 0);
        v_Temp := v_Temp || ',"advice_record_time":"' || Zljsonstr(To_Char(r_医嘱.开嘱时间, 'yyyy-mm-dd hh24:mi:ss')) || '"';
        v_Temp := v_Temp || ',"advice_doctor":"' || Zljsonstr(r_医嘱.开嘱医生) || '"';
      
        v_Temp := v_Temp || ',"advice_purpose":"' || Zljsonstr(r_医嘱.用药目的) || '"';
        v_Temp := v_Temp || ',"advice_reason":"' || Zljsonstr(r_医嘱.用药理由) || '"';
        v_Temp := v_Temp || ',"advice_taboonote":"' || Zljsonstr(r_医嘱.禁忌药品说明) || '"';
        v_Temp := v_Temp || ',"advice_doctor_note":"' || Zljsonstr(r_医嘱.医生嘱托) || '"';
        v_Temp := v_Temp || ',"rcpdtl_excs_desc":"' || Zljsonstr(r_医嘱.超量说明) || '"';
      
        v_Temp := v_Temp || ',"advice_audit_result":' || Nvl(r_医嘱.审查结果, 0);
        v_Temp := v_Temp || ',"advice_audit_sign":' || Nvl(r_医嘱.药师审核标志, 0);
        v_Temp := v_Temp || ',"advice_audit_time":"' || Zljsonstr(To_Char(r_医嘱.药师审核时间, 'yyyy-mm-dd hh24:mi:ss')) || '"';
        v_Temp := v_Temp || ',"advice_interval_unit":"' || Zljsonstr(r_医嘱.间隔单位) || '"';
        v_Temp := v_Temp || ',"advice_frequency":"' || Zljsonstr(r_医嘱.执行频次) || '"';
      
        v_Temp := v_Temp || ',"advice_frequency_times":' || Zljsonstr(r_医嘱.频率次数, 1);
        v_Temp := v_Temp || ',"advice_frequency_interval":' || Zljsonstr(r_医嘱.频率间隔, 1);
        v_Temp := v_Temp || ',"advice_exetime_plane":"' || Zljsonstr(r_医嘱.执行时间方案) || '"';
        v_Temp := v_Temp || ',"advice_begintime":"' || Zljsonstr(To_Char(r_医嘱.开始执行时间, 'yyyy-mm-dd hh24:mi:ss')) || '"';
        v_Temp := v_Temp || ',"advice_endtime":"' || Zljsonstr(To_Char(r_医嘱.执行终止时间, 'yyyy-mm-dd hh24:mi:ss')) || '"';
      
        v_Temp := v_Temp || ',"rgst_no":"' || Zljsonstr(r_医嘱.挂号单) || '"';
        v_Temp := v_Temp || ',"advice_receipt_name":"' || Zljsonstr(r_医嘱.配方名称) || '"';
        v_Temp := v_Temp || ',"advice_receipt_issecret":' || Nvl(r_医嘱.是否保密, 0);
        v_Temp := v_Temp || ',"advice_note":"' || Zljsonstr(r_医嘱.医嘱内容) || '"';
        v_Temp := v_Temp || ',"advice_cisitem_id":' || Nvl(r_医嘱.诊疗项目id, 0);
      
        v_Temp := v_Temp || ',"advice_item_id":' || Nvl(r_医嘱.收费细目id, 0);
        v_Temp := v_Temp || ',"pati_deptid":' || Nvl(r_医嘱.病人科室id, 0);
        v_Temp := v_Temp || ',"pati_name":"' || Zljsonstr(r_医嘱.姓名) || '"';
        v_Temp := v_Temp || ',"pati_sex":"' || Zljsonstr(r_医嘱.性别) || '"';
        v_Temp := v_Temp || ',"pati_age":"' || Zljsonstr(r_医嘱.年龄) || '"';
      
        v_Temp := v_Temp || ',"advice_audit_reason":"' || Zljsonstr(r_医嘱.药师审核原因) || '"';
        If n_皮试获取 = 2 And r_医嘱.皮试结果 Is Null Then
          Pp_Test(r_医嘱.诊疗项目id, v_皮试结果);
        Else
          v_皮试结果 := r_医嘱.皮试结果;
        End If;
        v_Temp := v_Temp || ',"skintest_info":"' || Zljsonstr(v_皮试结果) || '"';
      
        If n_查询类型 = 1 Then
          v_Temp := v_Temp || ',"total_qunt":' || Zljsonstr(r_医嘱.总给予量, 1);
          v_Temp := v_Temp || ',"Total_qunt_unit":"' || Zljsonstr(r_医嘱.总量) || '"';
          v_Temp := v_Temp || ',"single":"' || Zljsonstr(r_医嘱.单量) || '"';
          v_Temp := v_Temp || ',"toxicity_type":"' || Zljsonstr(r_医嘱.毒理分类) || '"';
          v_Temp := v_Temp || ',"advice_dept_id":' || Nvl(r_医嘱.开嘱科室id, 0);
        
          v_Temp := v_Temp || ',"advice_stop_doctor":"' || Zljsonstr(r_医嘱.停嘱医生) || '"';
          v_Temp := v_Temp || ',"advice_stop_nurse":"' || Zljsonstr(r_医嘱.确认停嘱护士) || '"';
          v_Temp := v_Temp || ',"advice_stoptime":"' || Zljsonstr(To_Char(r_医嘱.停嘱时间, 'yyyy-mm-dd hh24:mi:ss')) || '"';
          v_Temp := v_Temp || ',"advice_stoptime_confirm":"' ||
                    Zljsonstr(To_Char(r_医嘱.确认停嘱时间, 'yyyy-mm-dd hh24:mi:ss')) || '"';
          v_Temp := v_Temp || ',"order_chk_nurse":"' || Zljsonstr(r_医嘱.校对护士) || '"';
        
          v_Temp := v_Temp || ',"order_chk_time":"' || Zljsonstr(To_Char(r_医嘱.校对时间, 'yyyy-mm-dd hh24:mi:ss')) || '"';
          v_Temp := v_Temp || ',"lastexe_time":"' || Zljsonstr(To_Char(r_医嘱.上次执行时间, 'yyyy-mm-dd hh24:mi:ss')) || '"';
          v_Temp := v_Temp || ',"usage":"' || Zljsonstr(r_医嘱.用法) || '"';
          v_Temp := v_Temp || ',"emergency_tag":' || Nvl(r_医嘱.紧急标志, 0);
          v_Temp := v_Temp || ',"is_charge_verfy":' || Nvl(r_医嘱.是否费用审核, 0);
        
          v_Temp := v_Temp || ',"baby_num":' || Nvl(r_医嘱.婴儿序号, 0);
          v_Temp := v_Temp || ',"valuation_nature":' || Nvl(r_医嘱.计价特性, 0);
          v_Temp := v_Temp || ',"advice_exedept_id":' || Nvl(r_医嘱.执行科室id, 0);
          v_Temp := v_Temp || ',"advice_exedept_name":"' || Zljsonstr(r_医嘱.执行科室名称) || '"';
          v_Temp := v_Temp || ',"testtube_code":"' || Zljsonstr(r_医嘱.试管编码) || '"';
        
          v_Temp := v_Temp || ',"hide_print":' || Nvl(r_医嘱.屏蔽打印, 0);
          v_Temp := v_Temp || ',"Prerequisite_id":' || Nvl(r_医嘱.前提id, 0);
          v_Temp := v_Temp || ',"is_staff_sig":' || Nvl(r_医嘱.是否签名, 0);
          v_Temp := v_Temp || ',"rpt_id":' || Nvl(r_医嘱.报告id, 0);
          v_Temp := v_Temp || ',"consult_statu":' || Nvl(r_医嘱.查阅状态, 0);
        
          v_Temp := v_Temp || ',"advice_verfy_statu":' || Nvl(r_医嘱.审核状态, 0);
          v_Temp := v_Temp || ',"apply_num":' || Nvl(r_医嘱.申请序号, 0);
          v_Temp := v_Temp || ',"cisitem_unit":"' || Zljsonstr(r_医嘱.诊疗项目计算单位) || '"';
          v_Temp := v_Temp || ',"pati_wardarea_id":' || Nvl(r_医嘱.当前病区id, 0);
        
          If Nvl(r_医嘱.相关id, 0) <> 0 And (n_相关id Is Null Or n_相关id <> Nvl(r_医嘱.相关id, 0)) Then
            v_煎法 := Null;
            For r_煎法 In Ctest煎法(r_医嘱.相关id) Loop
              v_煎法 := r_煎法.名称;
            End Loop;
          
            n_相关id := r_医嘱.相关id;
          End If;
          v_Temp := v_Temp || ',"decoction_method":"' || Zljsonstr(v_煎法) || '"';
        End If;
        v_Temp := v_Temp || '}';
      
        If Length(v_Temp) > 25000 Then
          If c_Jtmp Is Null Then
            c_Jtmp := Substr(v_Temp, 2);
          Else
            c_Jtmp := c_Jtmp || v_Temp;
          End If;
          v_Temp := Null;
        End If;
      End If;
    End Loop;
  End Loop;

  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"成功","advice_list":[' || Substr(v_Temp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Temp;
    Json_Out := '{"output":{"code":1,"message":"成功","advice_list":[' || c_Jtmp || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getadviceinfo;
/
Create Or Replace Procedure Zl_Cissvr_Getadviceoperinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:获取医嘱相关的操作说明
  --入参：Json_In:格式
  -- input
  --   advice_id            C 1 医嘱id,多个用逗号分隔
  --   oper_type            N 1 操作类型:1-新开；2-校对疑问；3-校对通过；4-作废；5-重整；
  --                                     6-暂停；7-启用；8-停止；9-确认停止；10-皮试结果；
  --                                     11-审核通过；12-审核未通过；13-实习医师停嘱后待审核；14-血库接收；15-血库审核通过；
  --                                     16-血库配血拒绝；17-血库停止配血；18-输血初审通过待签发；9-输血初审回退；20-输血医嘱标记未用



  --   oper_last            N 1 是否取最后一次操作：1-取最后一次,0-取所有
  --出参: Json_Out,格式如下
  --  output
  --    code               N  1 应答码：0-失败；1-成功
  --    message            C  1 应答消息：失败时返回具体的错误信息
  --    oper_list[]数组
  --      advice_id        N  1 医嘱ID
  --      oper_time        C  1 操作时间:yyyy-mm-dd hh24:mi:ss
  --      oper_note        C  1 操作说明
  ---------------------------------------------------------------------------
  j_Json   Pljson;
  j_In     Pljson;
  c_医嘱id Clob;
  l_医嘱id t_Strlist := t_Strlist();

  n_操作类型 病人医嘱状态.操作类型%Type;
  n_Last     Number(2);
  v_Temp     Varchar2(32767);
  c_Jtmp     Clob;
Begin
  --解析入参
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');
  If j_Json.Exist('advice_id') Then
    c_医嘱id := j_Json.Get_Clob('advice_id');
  End If;
  n_操作类型 := j_Json.Get_Number('oper_type');
  n_Last     := j_Json.Get_Number('oper_last');

  While c_医嘱id Is Not Null Loop
    If Length(c_医嘱id) <= 4000 Then
      l_医嘱id.Extend;
      l_医嘱id(l_医嘱id.Count) := c_医嘱id;
      c_医嘱id := Null;
    Else
      l_医嘱id.Extend;
      l_医嘱id(l_医嘱id.Count) := Substr(c_医嘱id, 1, Instr(c_医嘱id, ',', 3950) - 1);
      c_医嘱id := Substr(c_医嘱id, Instr(c_医嘱id, ',', 3950) + 1);
    End If;
  End Loop;

  v_Temp := Null;
  For I In 1 .. l_医嘱id.Count Loop
    For r_医嘱 In (Select /*+cardinality(b,10)*/
                  a.医嘱id, Nvl(a.操作说明, '无') As 操作说明, To_Char(a.操作时间, 'yyyy-mm-dd hh24:mi:ss') As 操作时间
                 From 病人医嘱状态 A, Table(f_Num2list(l_医嘱id(I))) B
                 Where a.医嘱id = b.Column_Value And a.操作类型 = n_操作类型
                 Order By 操作时间 Desc) Loop
    
      v_Temp := v_Temp || ',{"advice_id":' || r_医嘱.医嘱id;
      v_Temp := v_Temp || ',"oper_time":"' || r_医嘱.操作时间 || '"';
      v_Temp := v_Temp || ',"oper_note":"' || Zljsonstr(r_医嘱.操作说明) || '"';
      v_Temp := v_Temp || '}';
    
      If Nvl(n_Last, 0) = 1 Then
        Exit; --只取最后一次
      End If;
    
      If Length(v_Temp) > 30000 Then
        If c_Jtmp Is Null Then
          c_Jtmp := Substr(v_Temp, 2);
        Else
          c_Jtmp := c_Jtmp || v_Temp;
        End If;
        v_Temp := Null;
      End If;
    End Loop;
    If Nvl(n_Last, 0) = 1 Then
      Exit; --只取最后一次
    End If;
  End Loop;
  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || Substr(v_Temp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Temp;
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getadviceoperinfo;
/
Create Or Replace Procedure Zl_Cissvr_Getadvicesendinfo
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：获取医嘱发送信息
  --入参      json
  --input
  --  query_type                N 1 查询类型：0-按医嘱id串（advice_ids）查询;1-按医嘱ID+医嘱发送号查询;2-按医嘱ID+记录性质+NO查询;3-仅按医嘱发送号查询
  --  return_type               N 1 返回类型：0-返回基本信息；1-返回基本信息+扩展信息
  --  contain_related           N 1 是否包含相关医嘱记录
  --  advice_ids                C 1 医嘱ID串，支持多个医嘱ID，用“，”分隔
  --  send_nos                  C 1 医嘱发送号串，支持多个，用“，”分隔,仅查询类型为3时有效
  --  item_list[]
  --    advice_id               N 1 医嘱ID
  --    send_no                 N 1 医嘱发送号
  --    bill_no                 C 1 NO
  --    bill_prop               N 1 记录性质
  --出参      json
  --output
  --  code                      C 1 应答码：0-失败；1-成功
  --  message                   C 1 应答消息：失败时返回具体的错误信息
  --  advice_send_list[]
  --    advice_id               N   医嘱id
  --    send_no                 N   发送号
  --    register_no             C   挂号单号
  --    pati_id                 N   病人id
  --    pati_pageid             N   主页id
  --    pati_deptid             N   病人科室id
  --    advice_dept_id          N   开嘱科室ID
  --    pati_source             N   病人来源
  --    clinic_type             C   诊疗类别
  --    valuation_nature        N   计价性质:0-正常计价；1-不计价；2-手工计价
  --    advice_related_id       N   相关id
  --    outpati_account         N   是否门诊记帐
  --    advice_note             C   医嘱内容
  --    sample_barcode          C   样本条码
  --    bill_no                 C   No
  --    bill_prop               N   记录性质
  --    advice_send_firsttime   D   首次时间
  --    advice_send_endtime     D   末次时间
  --    advice_send_exestatus   N   执行状态
  --    advice_send_sendtime    D   发送时间
  --    effective_time          N   医嘱期效
  ---------------------------------------------------------------------------
  v_医嘱id Clob; --记录医嘱id
  v_发送号 Clob;
  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_医嘱id Collection_Type;
  Col_发送号 Collection_Type;

  I            Number;
  n_查询类型   Number(1);
  n_返回类型   Number(1);
  n_含相关医嘱 Number(1);
  j_In         Pljson;
  j_Json       Pljson;
  j_List       Pljson_List := Pljson_List();
  j_Jsonlist   Pljson_List := Pljson_List();
  j_Json_Tmp   Pljson;

  v_No       病人医嘱发送.No%Type;
  n_记录性质 病人医嘱发送.记录性质%Type;
  n_医嘱id   病人医嘱发送.医嘱id%Type;
  n_发送号   病人医嘱发送.发送号%Type;
  v_Tmp      Varchar2(32767);
  v_Tmp2     Varchar2(32767);
  c_Temp     Clob; --保存json节点

  Cursor c_医嘱发送信息 Is
    Select a.医嘱id, a.发送号, a.门诊记帐, a.No, a.记录性质, a.首次时间, a.末次时间, a.执行状态, a.发送时间, a.样本条码, b.挂号单, b.病人id, b.主页id, b.病人科室id,
           b.开嘱科室id, b.病人来源, b.诊疗类别, b.相关id, b.医嘱内容, c.计价性质, b.医嘱期效
    From 病人医嘱发送 A, 病人医嘱记录 B, 诊疗项目目录 C
    Where a.医嘱id = b.Id And b.诊疗项目id = c.Id And Rownum < 1;
  r_Advice_Send c_医嘱发送信息%RowType;

  Type Ty_Advice_Send Is Ref Cursor;
  c_Advice_Send Ty_Advice_Send; --动态游标变量

  --组装失败时返回的数据
  Function Get_Err_Message(Message_In Varchar2) Return Clob Is
  Begin
    Return '{"output":{"code":0,"message":"' || Zltools.Zljsonstr(Message_In) || '"}}';
  End Get_Err_Message;
Begin
  j_In         := Pljson(Json_In);
  j_Json       := j_In.Get_Pljson('input');
  n_查询类型   := Nvl(j_Json.Get_Number('query_type'), 0);
  n_返回类型   := Nvl(j_Json.Get_Number('return_type'), 0);
  n_含相关医嘱 := Nvl(j_Json.Get_Number('contain_related'), 0);

  If n_查询类型 = 0 Then
    Begin
      v_医嘱id := j_Json.Get_Clob('advice_ids');
    Exception
      When Others Then
        Json_Out := Get_Err_Message('未传入医嘱id，请检查！');
        Return;
    End;
  Elsif n_查询类型 = 3 Then
    Begin
      v_发送号 := j_Json.Get_Clob('send_nos');
    Exception
      When Others Then
        Json_Out := Get_Err_Message('未传入医嘱发送号，请检查！');
        Return;
    End;
  Else
    j_List := j_Json.Get_Pljson_List('item_list');
    If j_List Is Null Then
      Json_Out := Get_Err_Message('未传入医嘱id等信息，请检查！');
      Return;
    End If;
  End If;

  If n_查询类型 = 0 Then
    --将 v_医嘱id 串组装成不超过4000 的集合串，防止使用 f_Num2list 参数超长
    I := 0;
    While v_医嘱id Is Not Null Loop
      If Length(v_医嘱id) <= 4000 Then
        Col_医嘱id(I) := v_医嘱id;
        v_医嘱id := Null;
      Else
        Col_医嘱id(I) := Substr(v_医嘱id, 1, Instr(v_医嘱id, ',', 3980) - 1);
        v_医嘱id := Substr(v_医嘱id, Instr(v_医嘱id, ',', 3980) + 1);
      End If;
      I := I + 1;
    End Loop;
    I := 0;
  
    For I In 0 .. Col_医嘱id.Count - 1 Loop
      Open c_Advice_Send For
        Select /*+cardinality(j,10)*/
         a.医嘱id, a.发送号, a.门诊记帐, a.No, a.记录性质, a.首次时间, a.末次时间, a.执行状态, a.发送时间, a.样本条码, b.挂号单, b.病人id, b.主页id, b.病人科室id,
         b.开嘱科室id, b.病人来源, b.诊疗类别, b.相关id, b.医嘱内容, c.计价性质, b.医嘱期效
        From 病人医嘱发送 A, 病人医嘱记录 B, 诊疗项目目录 C, Table(f_Num2list(Col_医嘱id(I))) J
        Where a.医嘱id = b.Id And b.诊疗项目id = c.Id And b.Id = j.Column_Value
        Union All
        Select /*+cardinality(j,10)*/
         a.医嘱id, a.发送号, a.门诊记帐, a.No, a.记录性质, a.首次时间, a.末次时间, a.执行状态, a.发送时间, a.样本条码, b.挂号单, b.病人id, b.主页id, b.病人科室id,
         b.开嘱科室id, b.病人来源, b.诊疗类别, b.相关id, b.医嘱内容, c.计价性质, b.医嘱期效
        From 病人医嘱发送 A, 病人医嘱记录 B, 诊疗项目目录 C, Table(f_Num2list(Col_医嘱id(I))) J
        Where a.医嘱id = b.Id And b.诊疗项目id = c.Id And n_含相关医嘱 = 1 And b.相关id = j.Column_Value;
    
      Loop
        Fetch c_Advice_Send
          Into r_Advice_Send;
        Exit When c_Advice_Send %NotFound;
      
        --基本信息
        v_Tmp := '';
        Zljsonputvalue(v_Tmp, 'advice_id', r_Advice_Send.医嘱id, 1, 1);
        Zljsonputvalue(v_Tmp, 'send_no', r_Advice_Send.发送号, 1);
        Zljsonputvalue(v_Tmp, 'bill_no', r_Advice_Send.No);
        Zljsonputvalue(v_Tmp, 'bill_prop', r_Advice_Send.记录性质, 1);
        Zljsonputvalue(v_Tmp, 'advice_send_firsttime', r_Advice_Send.首次时间);
        Zljsonputvalue(v_Tmp, 'advice_send_endtime', r_Advice_Send.末次时间);
        Zljsonputvalue(v_Tmp, 'advice_send_exestatus', r_Advice_Send.执行状态, 1);
      
        If n_返回类型 = 1 Then
          Zljsonputvalue(v_Tmp, 'advice_send_sendtime', r_Advice_Send.发送时间);
        
          --扩展信息
          Zljsonputvalue(v_Tmp, 'register_no', r_Advice_Send.挂号单);
          Zljsonputvalue(v_Tmp, 'pati_id', r_Advice_Send.病人id, 1);
          Zljsonputvalue(v_Tmp, 'pati_pageid', r_Advice_Send.主页id, 1);
          Zljsonputvalue(v_Tmp, 'pati_deptid', r_Advice_Send.病人科室id, 1);
          Zljsonputvalue(v_Tmp, 'advice_dept_id', r_Advice_Send.开嘱科室id, 1);
        
          Zljsonputvalue(v_Tmp, 'pati_source', r_Advice_Send.病人来源, 1);
          Zljsonputvalue(v_Tmp, 'clinic_type', r_Advice_Send.诊疗类别);
          Zljsonputvalue(v_Tmp, 'valuation_nature', r_Advice_Send.计价性质, 1);
          Zljsonputvalue(v_Tmp, 'advice_related_id', r_Advice_Send.相关id, 1);
          Zljsonputvalue(v_Tmp, 'outpati_account', r_Advice_Send.门诊记帐, 1);
        
          Zljsonputvalue(v_Tmp, 'advice_note', Nvl(r_Advice_Send.医嘱内容, ''));
          Zljsonputvalue(v_Tmp, 'sample_barcode', Nvl(r_Advice_Send.样本条码, ''));
          Zljsonputvalue(v_Tmp, 'effective_time', r_Advice_Send.医嘱期效, 1, 2);
        Else
          Zljsonputvalue(v_Tmp, 'advice_send_sendtime', r_Advice_Send.发送时间, 0, 2);
        End If;
      
        If (Length(v_Tmp) + Length(v_Tmp2)) > 32000 Then
          c_Temp := c_Temp || v_Tmp2 || ',' || v_Tmp;
          v_Tmp2 := Null;
        Else
          v_Tmp2 := v_Tmp2 || ',' || v_Tmp;
        End If;
      End Loop;
    End Loop;
  
  Elsif n_查询类型 = 3 Then
    --将 v_发送号 串组装成不超过4000 的集合串，防止使用 f_Num2list 参数超长
    I := 0;
    While v_发送号 Is Not Null Loop
      If Length(v_发送号) <= 4000 Then
        Col_发送号(I) := v_发送号;
        v_发送号 := Null;
      Else
        Col_发送号(I) := Substr(v_发送号, 1, Instr(v_发送号, ',', 3980) - 1);
        v_发送号 := Substr(v_发送号, Instr(v_发送号, ',', 3980) + 1);
      End If;
      I := I + 1;
    End Loop;
    I := 0;
  
    For I In 0 .. Col_发送号.Count - 1 Loop
      Open c_Advice_Send For
        Select /*+cardinality(j,10)*/
         a.医嘱id, a.发送号, a.门诊记帐, a.No, a.记录性质, a.首次时间, a.末次时间, a.执行状态, a.发送时间, a.样本条码, b.挂号单, b.病人id, b.主页id, b.病人科室id,
         b.开嘱科室id, b.病人来源, b.诊疗类别, b.相关id, b.医嘱内容, c.计价性质, b.医嘱期效
        From 病人医嘱发送 A, 病人医嘱记录 B, 诊疗项目目录 C, Table(f_Num2list(Col_发送号(I))) J
        Where a.医嘱id = b.Id And b.诊疗项目id = c.Id And a.发送号 = j.Column_Value;
    
      Loop
        Fetch c_Advice_Send
          Into r_Advice_Send;
        Exit When c_Advice_Send %NotFound;
      
        --基本信息
        v_Tmp := '';
        Zljsonputvalue(v_Tmp, 'advice_id', r_Advice_Send.医嘱id, 1, 1);
        Zljsonputvalue(v_Tmp, 'send_no', r_Advice_Send.发送号, 1);
        Zljsonputvalue(v_Tmp, 'bill_no', r_Advice_Send.No);
        Zljsonputvalue(v_Tmp, 'bill_prop', r_Advice_Send.记录性质, 1);
        Zljsonputvalue(v_Tmp, 'advice_send_firsttime', r_Advice_Send.首次时间);
        Zljsonputvalue(v_Tmp, 'advice_send_endtime', r_Advice_Send.末次时间);
        Zljsonputvalue(v_Tmp, 'advice_send_exestatus', r_Advice_Send.执行状态, 1);
      
        If n_返回类型 = 1 Then
          Zljsonputvalue(v_Tmp, 'advice_send_sendtime', r_Advice_Send.发送时间);
        
          --扩展信息
          Zljsonputvalue(v_Tmp, 'register_no', r_Advice_Send.挂号单);
          Zljsonputvalue(v_Tmp, 'pati_id', r_Advice_Send.病人id, 1);
          Zljsonputvalue(v_Tmp, 'pati_pageid', r_Advice_Send.主页id, 1);
          Zljsonputvalue(v_Tmp, 'pati_deptid', r_Advice_Send.病人科室id, 1);
          Zljsonputvalue(v_Tmp, 'advice_dept_id', r_Advice_Send.开嘱科室id, 1);
        
          Zljsonputvalue(v_Tmp, 'pati_source', r_Advice_Send.病人来源, 1);
          Zljsonputvalue(v_Tmp, 'clinic_type', r_Advice_Send.诊疗类别);
          Zljsonputvalue(v_Tmp, 'valuation_nature', r_Advice_Send.计价性质, 1);
          Zljsonputvalue(v_Tmp, 'advice_related_id', r_Advice_Send.相关id, 1);
          Zljsonputvalue(v_Tmp, 'outpati_account', r_Advice_Send.门诊记帐, 1);
        
          Zljsonputvalue(v_Tmp, 'advice_note', Nvl(r_Advice_Send.医嘱内容, ''));
          Zljsonputvalue(v_Tmp, 'sample_barcode', Nvl(r_Advice_Send.样本条码, ''));
          Zljsonputvalue(v_Tmp, 'effective_time', r_Advice_Send.医嘱期效, 1, 2);
        Else
          Zljsonputvalue(v_Tmp, 'advice_send_sendtime', r_Advice_Send.发送时间, 0, 2);
        End If;
      
        If (Length(v_Tmp) + Length(v_Tmp2)) > 32000 Then
          c_Temp := c_Temp || v_Tmp2 || ',' || v_Tmp;
          v_Tmp2 := Null;
        Else
          v_Tmp2 := v_Tmp2 || ',' || v_Tmp;
        End If;
      End Loop;
    End Loop;
  Else
    For I In 1 .. j_List.Count Loop
      j_Json_Tmp := Pljson();
      j_Json_Tmp := Pljson(j_List.Get(I));
    
      If n_查询类型 = 1 Then
        n_医嘱id   := j_Json_Tmp.Get_Number('advice_id');
        n_发送号   := j_Json_Tmp.Get_Number('send_no');
        n_记录性质 := Nvl(j_Json_Tmp.Get_Number('bill_prop'), 0);
      
        Open c_Advice_Send For
          Select a.医嘱id, a.发送号, a.门诊记帐, a.No, a.记录性质, a.首次时间, a.末次时间, a.执行状态, a.发送时间, a.样本条码, b.挂号单, b.病人id, b.主页id,
                 b.病人科室id, b.开嘱科室id, b.病人来源, b.诊疗类别, b.相关id, b.医嘱内容, c.计价性质, b.医嘱期效
          From 病人医嘱发送 A, 病人医嘱记录 B, 诊疗项目目录 C
          Where a.医嘱id = b.Id And b.诊疗项目id = c.Id And b.Id = n_医嘱id And a.发送号 = n_发送号
          Union All
          Select a.医嘱id, a.发送号, a.门诊记帐, a.No, a.记录性质, a.首次时间, a.末次时间, a.执行状态, a.发送时间, a.样本条码, b.挂号单, b.病人id, b.主页id,
                 b.病人科室id, b.开嘱科室id, b.病人来源, b.诊疗类别, b.相关id, b.医嘱内容, c.计价性质, b.医嘱期效
          From 病人医嘱发送 A, 病人医嘱记录 B, 诊疗项目目录 C
          Where a.医嘱id = b.Id And b.诊疗项目id = c.Id And n_含相关医嘱 = 1 And b.相关id = n_医嘱id And a.发送号 = n_发送号;
      End If;
    
      If n_查询类型 = 2 Then
        n_医嘱id   := j_Json_Tmp.Get_Number('advice_id');
        v_No       := j_Json_Tmp.Get_String('bill_no');
        n_记录性质 := j_Json_Tmp.Get_Number('bill_prop');
      
        Open c_Advice_Send For
          Select a.医嘱id, a.发送号, a.门诊记帐, a.No, a.记录性质, a.首次时间, a.末次时间, a.执行状态, a.发送时间, a.样本条码, b.挂号单, b.病人id, b.主页id,
                 b.病人科室id, b.开嘱科室id, b.病人来源, b.诊疗类别, b.相关id, b.医嘱内容, c.计价性质, b.医嘱期效
          From 病人医嘱发送 A, 病人医嘱记录 B, 诊疗项目目录 C
          Where a.医嘱id = b.Id And b.诊疗项目id = c.Id And b.Id = n_医嘱id And a.No = v_No And a.记录性质 = n_记录性质
          Union All
          Select a.医嘱id, a.发送号, a.门诊记帐, a.No, a.记录性质, a.首次时间, a.末次时间, a.执行状态, a.发送时间, a.样本条码, b.挂号单, b.病人id, b.主页id,
                 b.病人科室id, b.开嘱科室id, b.病人来源, b.诊疗类别, b.相关id, b.医嘱内容, c.计价性质, b.医嘱期效
          From 病人医嘱发送 A, 病人医嘱记录 B, 诊疗项目目录 C
          Where a.医嘱id = b.Id And b.诊疗项目id = c.Id And n_含相关医嘱 = 1 And b.相关id = n_医嘱id And a.No = v_No And
                a.记录性质 = n_记录性质;
      End If;
    
      Loop
        Fetch c_Advice_Send
          Into r_Advice_Send;
        Exit When c_Advice_Send %NotFound;
      
        --基本信息
        v_Tmp := '';
        Zljsonputvalue(v_Tmp, 'advice_id', r_Advice_Send.医嘱id, 1, 1);
        Zljsonputvalue(v_Tmp, 'send_no', r_Advice_Send.发送号, 1);
        Zljsonputvalue(v_Tmp, 'bill_no', r_Advice_Send.No);
        Zljsonputvalue(v_Tmp, 'bill_prop', r_Advice_Send.记录性质, 1);
        Zljsonputvalue(v_Tmp, 'advice_send_firsttime', r_Advice_Send.首次时间);
        Zljsonputvalue(v_Tmp, 'advice_send_endtime', r_Advice_Send.末次时间);
        Zljsonputvalue(v_Tmp, 'advice_send_exestatus', r_Advice_Send.执行状态, 1);
      
        If n_返回类型 = 1 Then
          Zljsonputvalue(v_Tmp, 'advice_send_sendtime', r_Advice_Send.发送时间);
        
          --扩展信息
          Zljsonputvalue(v_Tmp, 'register_no', r_Advice_Send.挂号单);
          Zljsonputvalue(v_Tmp, 'pati_id', r_Advice_Send.病人id, 1);
          Zljsonputvalue(v_Tmp, 'pati_pageid', r_Advice_Send.主页id, 1);
          Zljsonputvalue(v_Tmp, 'pati_deptid', r_Advice_Send.病人科室id, 1);
          Zljsonputvalue(v_Tmp, 'advice_dept_id', r_Advice_Send.开嘱科室id, 1);
        
          Zljsonputvalue(v_Tmp, 'pati_source', r_Advice_Send.病人来源, 1);
          Zljsonputvalue(v_Tmp, 'clinic_type', r_Advice_Send.诊疗类别);
          Zljsonputvalue(v_Tmp, 'valuation_nature', r_Advice_Send.计价性质, 1);
          Zljsonputvalue(v_Tmp, 'advice_related_id', r_Advice_Send.相关id, 1);
          Zljsonputvalue(v_Tmp, 'outpati_account', r_Advice_Send.门诊记帐, 1);
        
          Zljsonputvalue(v_Tmp, 'advice_note', Nvl(r_Advice_Send.医嘱内容, ''));
          Zljsonputvalue(v_Tmp, 'sample_barcode', Nvl(r_Advice_Send.样本条码, ''));
          Zljsonputvalue(v_Tmp, 'effective_time', r_Advice_Send.医嘱期效, 1, 2);
        Else
          Zljsonputvalue(v_Tmp, 'advice_send_sendtime', r_Advice_Send.发送时间, 0, 2);
        End If;
      
        If (Length(v_Tmp) + Length(v_Tmp2)) > 32000 Then
          c_Temp := c_Temp || v_Tmp2 || ',' || v_Tmp;
          v_Tmp2 := Null;
        Else
          v_Tmp2 := v_Tmp2 || ',' || v_Tmp;
        End If;
      End Loop;
    End Loop;
  End If;

  If v_Tmp2 Is Not Null Then
    c_Temp := c_Temp || v_Tmp2;
  End If;

  Json_Out := '{"output":{"code":1,"message": "成功","advice_send_list":[' || Substr(c_Temp, 2) || ']}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getadvicesendinfo;
/
Create Or Replace Procedure Zl_Cissvr_Getadvicesendnums
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:获取医嘱发送的数次信息
  --入参：Json_In:格式
  -- input
  --   advice_id            N 1 医嘱id
  --   advice_sendno        N 1 发送号
  --   isalone_exe          N 1 是否独立执行:1-独立执行;0-非独立执行

  --出参: Json_Out,格式如下
  --  output
  --    code                  N  1 应答码：0-失败；1-成功
  --    message               C  1 应答消息：失败时返回具体的错误信息
  --    sendlist              C   数据组
  --      advice_num          N 1 序号:病人医嘱记录.序号
  --      advice_id           N 1 医嘱id
  --      advice_related_id   N 1 相关id
  --      clinic_type         C 1 诊疗类别
  --      advice_cisitem_id   N 1 诊疗项目id
  --      advice_dept_id      N 1 开嘱科室ID
  --      exe_deptid          N 1 执行部门ID
  --      nums                N 1 数次
  --      citem_spcm_parts    C 1 标本部位
  --      citem_exam_method   C 1 检查方法
  --      advice_exe_sign     N 1 执行标记
  --      pati_id             N 1 病人id
  --      pati_pageid         N 1 主页id
  ---------------------------------------------------------------------------
  j_In       Pljson;
  j_Json     Pljson;
  n_医嘱id   病人医嘱记录.Id%Type;
  n_相关id   病人医嘱记录.相关id%Type;
  n_独立执行 Number(1);
  n_发送号   病人医嘱发送.发送号%Type;

  v_Jtmp Varchar2(32767);

Begin
  --解析入参
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  n_医嘱id   := j_Json.Get_Number('advice_id');
  n_独立执行 := Nvl(j_Json.Get_Number('isalone_exe'), 0);
  n_发送号   := j_Json.Get_Number('advice_sendno');

  If Nvl(n_医嘱id, 0) = 0 And Nvl(n_发送号, 0) = 0 Then
    Json_Out := Zljsonout('未传入医嘱id或发送号，请检查！');
    Return;
  End If;

  Select Decode(a.诊疗类别, 'D', Nvl(a.相关id, a.Id), a.相关id)
  Into n_相关id
  From 病人医嘱记录 A
  Where a.Id = n_医嘱id;

  For r_医嘱发送 In (Select b.序号, a.医嘱id, b.相关id, b.诊疗类别, b.诊疗项目id, b.开嘱科室id, a.执行部门id,
                        Nvl(a.发送数次, Sum(Nvl(c.本次数次, 0))) As 数量, b.标本部位, b.检查方法, b.执行标记, b.病人id, b.主页id
                 From 病人医嘱发送 A, 病人医嘱记录 B, 病人医嘱执行 C
                 Where Nvl(a.计费状态, 0) = 0 And b.Id = n_医嘱id And
                       (b.Id = n_医嘱id Or
                        ((b.相关id = n_医嘱id And b.诊疗类别 In ('F', 'D')) Or (b.相关id = n_相关id And b.诊疗类别 = 'C') And n_独立执行 = 0)) And
                       a.发送号 = n_发送号 And c.医嘱id(+) = a.医嘱id And c.发送号(+) = a.发送号
                 Group By b.序号, a.医嘱id, b.相关id, b.诊疗类别, b.诊疗项目id, b.开嘱科室id, a.执行部门id, a.发送数次, b.标本部位, b.检查方法, b.执行标记,
                          b.病人id, b.主页id
                 Having Nvl(a.发送数次, Sum(Nvl(c.本次数次, 0))) <> 0
                 Order By 序号) Loop
  
    --      advice_num          N 1 序号:病人医嘱记录.序号
    --      advice_id           N 1 医嘱id
    --      advice_related_id   N 1 相关id
    --      clinic_type         C 1 诊疗类别
    --      advice_cisitem_id   N 1 诊疗项目id
    --      advice_dept_id      N 1 开嘱科室ID
    --      exe_deptid          N 1 执行部门ID
    v_Jtmp := v_Jtmp || ',{"advice_num":' || r_医嘱发送.序号;
    v_Jtmp := v_Jtmp || ',"advice_id":' || r_医嘱发送.医嘱id;
    v_Jtmp := v_Jtmp || ',"advice_related_id":' || Nvl(r_医嘱发送.相关id || '', 'null');
    v_Jtmp := v_Jtmp || ',"clinic_type":"' || r_医嘱发送.诊疗类别 || '"';
    v_Jtmp := v_Jtmp || ',"advice_cisitem_id":' || Nvl(r_医嘱发送.诊疗项目id || '', 'null');
    v_Jtmp := v_Jtmp || ',"advice_dept_id":' || Nvl(r_医嘱发送.开嘱科室id || '', 'null');
    v_Jtmp := v_Jtmp || ',"exe_deptid":' || Nvl(r_医嘱发送.执行部门id || '', 'null');
  
    --      nums                N 1 数次
    --      citem_spcm_parts    C 1 标本部位
    --      citem_exam_method   C 1 检查方法
    --      advice_exe_sign     N 1 执行标记
    --      pati_id             N 1 病人id
    --      pati_pageid         N 1 主页id
    v_Jtmp := v_Jtmp || ',"nums":' || Zljsonstr(r_医嘱发送.数量, 1);
    v_Jtmp := v_Jtmp || ',"citem_spcm_parts":"' || r_医嘱发送.标本部位 || '"';
    v_Jtmp := v_Jtmp || ',"citem_exam_method":"' || Zljsonstr(r_医嘱发送.检查方法) || '"';
    v_Jtmp := v_Jtmp || ',"advice_exe_sign":' || Nvl(r_医嘱发送.执行标记, 0);
    v_Jtmp := v_Jtmp || ',"pati_id":' || r_医嘱发送.病人id;
    v_Jtmp := v_Jtmp || ',"pati_pageid":' || Nvl(r_医嘱发送.主页id || '', 'null');
  
    v_Jtmp := v_Jtmp || '}';
  
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"成功","sendlist":[' || Substr(v_Jtmp, 2) || ']}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getadvicesendnums;
/
Create Or Replace Procedure Zl_Cissvr_Getaffirmerrordata
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：临床医嘱执行完成时自动审核费用，异常后未对药品和卫材进行收费确认，针对此类异常数据获取服务
  --入参：Json_In:格式
  --  input
  --      pati_list[]病人关键信息，用于获取医嘱
  --           pati_id                    N 1 病人id
  --           pati_pageid                N 1 主页id，住院病人传入，门诊传0
  --           rgst_id                    N 1 挂号id，门诊病人传入，住院病人传空
  --           rgst_no                    C 1 挂号单号
  --出参: Json_Out,格式如下
  --   output:
  --    code                  N 1 应答吗：0-失败；1-成功
  --    message               C 1 应答消息：失败时返回具体的错误信息
  --     pati_bill_list[]
  --         pati_id                      N 1 病人id
  --         pati_pageid                  N 0 主页id，住院病人传入，门诊传0
  --         rgst_id                      N 0 挂号id，门诊病人传入，住院病人传空
  --         rgst_no                      C 0 挂号单号
  --         order_ids                    C 1 有异常的所有医嘱id拼串
  --         fee_nos                      C 1 有异常的所有单据号拼串
  --         order_list[]医嘱发送信息列表
  --             send_no                  N 1 发送号
  --             advice_id                N 1 医嘱id
  --             fee_no                   C 1 单据号
  --             bill_prop                N 1 记录性质
  --             outpati_account          N 1 是否门诊记帐 0-不是门诊记帐，1-是门诊记帐
  --             pati_source              N 1 病人来源 1-门诊医嘱，2-住院医嘱
  ----------------------------------------------------------------------------------

  j_Json        Pljson;
  v_Rgs_No      Varchar2(2000);
  Jl_All_In     Pljson_List := Pljson_List();
  j_Item_a      Pljson;
  n_Pati_Id     Number; --N   1病人ID
  n_Pati_Pageid Number; --N   1主页ID
  n_Rgst_Id     Number; --N   1挂号单id
  v_医嘱ids     Varchar2(32767);
  v_Nos         Varchar2(32767);
  v_Item_a      Varchar2(32767);
  c_Item_a      Clob;
  v_Item        Varchar2(32767);
  c_Item        Clob;
  Cursor c_Out Is
    Select b.医嘱id, b.No, b.发送号, b.记录性质, b.门诊记帐, 1 As 来源
    From 病人医嘱记录 A, 病人医嘱发送 B, 病人医嘱异常记录 C
    Where a.Id = b.医嘱id And b.医嘱id = c.医嘱id And b.发送号 = c.发送号 And a.挂号单 = v_Rgs_No And c.产生环节 In (4, 5)
    Order By b.发送号;

  Type t_Order Is Table Of c_Out%RowType;
  r_Odr t_Order;

  Cursor c_In Is
    Select b.医嘱id, b.No, b.发送号, b.记录性质, b.门诊记帐, 2 As 来源
    From 病人医嘱记录 A, 病人医嘱发送 B, 病人医嘱异常记录 C
    Where a.Id = b.医嘱id And b.医嘱id = c.医嘱id And b.发送号 = c.发送号 And a.病人id = n_Pati_Id And a.主页id = n_Pati_Pageid And
          c.产生环节 In (4, 5)
    Order By b.发送号;

  Procedure p_Getbaseinfo As
  Begin
    If Nvl(n_Rgst_Id, 0) = 0 Then
      Open c_In;
      Fetch c_In Bulk Collect
        Into r_Odr;
      Close c_In;
    Else
      Open c_Out;
      Fetch c_Out Bulk Collect
        Into r_Odr;
      Close c_Out;
    End If;
  End p_Getbaseinfo;

Begin
  --解析入参
  j_Item_a  := Pljson(Json_In);
  j_Json    := j_Item_a.Get_Pljson('input');
  Jl_All_In := j_Json.Get_Pljson_List('pati_list');
  For Pi In 1 .. Jl_All_In.Count Loop
    j_Item_a      := Pljson();
    j_Item_a      := Pljson(Jl_All_In.Get(Pi));
    n_Pati_Id     := j_Item_a.Get_Number('pati_id');
    n_Pati_Pageid := j_Item_a.Get_Number('pati_pageid');
    n_Rgst_Id     := j_Item_a.Get_Number('rgst_id');
    v_Rgs_No      := j_Item_a.Get_String('rgst_no');
    p_Getbaseinfo;
    If r_Odr.Count > 0 Then
      For Ol In 1 .. r_Odr.Count Loop
      
        If Instr(',' || v_医嘱ids || ',', ',' || r_Odr(Ol).医嘱id || ',') = 0 Then
          v_医嘱ids := v_医嘱ids || ',' || r_Odr(Ol).医嘱id;
        End If;
        If Instr(',' || v_Nos || ',', ',' || r_Odr(Ol).No || ',') = 0 Then
          v_Nos := v_Nos || ',' || r_Odr(Ol).No;
        End If;
      
        v_Item := v_Item || ',{"send_no":' || r_Odr(Ol).发送号;
        v_Item := v_Item || ',"advice_id":' || r_Odr(Ol).医嘱id;
        v_Item := v_Item || ',"fee_no":"' || r_Odr(Ol).No || '"';
        v_Item := v_Item || ',"bill_prop":' || r_Odr(Ol).记录性质;
        v_Item := v_Item || ',"outpati_account":' || Nvl(r_Odr(Ol).门诊记帐 || '', 'null');
        v_Item := v_Item || ',"pati_source":' || r_Odr(Ol).来源;
        v_Item := v_Item || '}';
      
        If Length(v_Item) > 30000 Then
          If c_Item Is Null Then
            c_Item := Substr(v_Item, 2);
          Else
            c_Item := c_Item || v_Item;
          End If;
          v_Item := Null;
        End If;
      End Loop;
    
      v_Item_a := v_Item_a || ',{"pati_id":' || n_Pati_Id;
      v_Item_a := v_Item_a || ',"pati_pageid":' || Nvl(n_Pati_Pageid || '', 'null');
      v_Item_a := v_Item_a || ',"rgst_id":' || Nvl(n_Rgst_Id || '', 'null');
      v_Item_a := v_Item_a || ',"rgst_no":"' || v_Rgs_No || '"';
      v_Item_a := v_Item_a || ',"order_ids":"' || Substr(v_医嘱ids, 2) || '"';
      v_Item_a := v_Item_a || ',"fee_nos":"' || Substr(v_Nos, 2) || '"';
    
      If c_Item_a Is Null Then
        If c_Item Is Null Then
          v_Item_a := v_Item_a || ',"order_list":[' || Substr(v_Item, 2) || ']';
          v_Item_a := v_Item_a || '}';
          c_Item_a := Substr(v_Item_a, 2);
        Else
          c_Item   := c_Item || v_Item;
          c_Item_a := Substr(v_Item_a, 2) || ',"order_list":[' || c_Item || ']';
          c_Item_a := c_Item_a || '}';
        End If;
      Else
        If c_Item Is Null Then
          v_Item_a := v_Item_a || ',"order_list":[' || Substr(v_Item, 2) || ']';
          v_Item_a := v_Item_a || '}';
          c_Item_a := c_Item_a || v_Item_a;
        Else
          c_Item   := c_Item || v_Item;
          c_Item_a := c_Item_a || v_Item_a || ',"order_list":[' || c_Item || ']';
          c_Item_a := c_Item_a || '}';
        End If;
      End If;
    
      v_Item    := Null;
      c_Item    := Null;
      v_Item_a  := Null;
      v_医嘱ids := Null;
      v_Nos     := Null;
    End If;
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"成功","pati_bill_list":[' || c_Item_a || ']}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getaffirmerrordata;
/
Create Or Replace Procedure Zl_Cissvr_Getallgroupadviceids
(
  Json_In  Varchar,
  Json_Out Out Varchar
) Is
  ---------------------------------------------------------------------------
  --根据医嘱ID查询一组医嘱的所有医嘱信息
  --入参      json
  --input     根据医嘱ID查询医嘱信息
  --  advice_id                     C 1  多个医嘱ID，用逗号分隔
  --出参      json
  --output
  --  code                          C  1  应答码：0-失败；1-成功
  --  message                       C  1  应答消息：
  --  advice_list[]                 每个医嘱信息
  --    advice_id                   N    id
  --    advice_related_id           N    相关id
  --    clinic_type                 C    诊疗类别
  ---------------------------------------------------------------------------
  j_Json    Pljson;
  j_In      Pljson;
  v_医嘱ids Varchar2(32767);
  v_Temp Varchar2(32767); 
Begin
  j_In      := Pljson(Json_In);
  j_Json    := j_In.Get_Pljson('input');
  v_医嘱ids := j_Json.Get_String('advice_id');

  For r_医嘱 In (Select /*+cardinality(j,10)*/
                a.Id, a.挂号单, a.医嘱内容, a.相关id, a.诊疗类别
               From 病人医嘱记录 A, Table(f_Num2list(v_医嘱ids)) J
               Where a.Id = j.Column_Value Or a.相关id = j.Column_Value
               Union All
               Select /*+cardinality(j,10)*/
                a.Id, a.挂号单, a.医嘱内容, a.相关id, a.诊疗类别
               From 病人医嘱记录 A, 病人医嘱记录 B, Table(f_Num2list(v_医嘱ids)) J
               Where a.Id = b.相关id And b.Id = j.Column_Value) Loop
  
    v_Temp := v_Temp || ',{"advice_id":' || r_医嘱.Id;
    v_Temp := v_Temp || ',"advice_related_id":' || Nvl(r_医嘱.相关id || '', 'null');
    v_Temp := v_Temp || ',"clinic_Type":"' || r_医嘱.诊疗类别 || '"';
    v_Temp := v_Temp || '}';
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"成功","advice_list":[' || Substr(v_Temp, 2) || ']}}';
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Cissvr_Getallgroupadviceids;
/
Create Or Replace Procedure Zl_Cissvr_Getbabydata
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:根据病人id和主页Id获取新生儿数据
  --入参：Json_In:格式
  --    input
  --      pati_id           N 1 病人id
  --      pati_pageid       N 1 主页id

  --出参: Json_Out,格式如下
  --  output
  --    code                  N 1 应答码：0-失败；1-成功
  --    message               C 1 应答消息：失败时返回具体的错误信息
  --    baby_list[]
  --      baby_num            N 1 婴儿序号
  --      baby_name           C 1 婴儿姓名
  ---------------------------------------------------------------------------
  j_In     Pljson;
  j_Json   Pljson;
  v_List   Varchar2(32767);
  n_病人id 病人新生儿记录.病人id%Type;
  n_主页id 病人新生儿记录.主页id%Type;

Begin
  --解析入参
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');
  n_主页id := j_Json.Get_Number('pati_pageid');

  If Nvl(n_病人id, 0) = 0 Then
    Json_Out := Zljsonout('未传入病人id，请检查！');
    Return;
  End If;

  If Nvl(n_主页id, 0) = 0 Then
    Json_Out := Zljsonout('未传入主页id，请检查！');
    Return;
  End If;

  For r_新生儿 In (Select 序号, 婴儿姓名 From 病人新生儿记录 Where 病人id = n_病人id And 主页id = n_主页id) Loop
  
    v_List := v_List || ',{"baby_num":' || r_新生儿.序号;
    v_List := v_List || ',"baby_name":"' || Zljsonstr(r_新生儿.婴儿姓名) || '"';
    v_List := v_List || '}';
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"成功","baby_list":[' || Substr(v_List, 2) || ']}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getbabydata;
/
Create Or Replace Procedure Zl_Cissvr_Getcriticalinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  -------------------------------------------------------------
  --功能：获取病人的危急值信息

  --入参      json
  --input
  --    use_type                            N  1    调用类型 ：1-通过id字符串查询,2-通过挂号单查询,3-通过病人id加主页id查询,4-过滤危急值列表(报告时间为索引  不能为空)
  --    cvalue_ids                          N  1    危急值ids
  --    rgst_no                             C  0    挂号单
  --    pati_id                             N  0    病人id
  --    pati_pageid                         N  0    主页id

  --    cvalue_time_begin                   C  0   报告时间范围开始时间
  --    cvalue_time_end                     C  0   报告时间范围结束时间
  --    pati_type                           N  0   病人类型 0-全部 1-门诊 2-住院 3-外来
  --    rpt_deptid                          N  0   报告科室ID 为空时  不过滤
  --    cnfm_deptid                         N  0   确认科室ID 为空时  不过滤
  --    cvalue_rec_status                   N  0   确认状态 0-全部 1-未确认 2-确认为非危急值 3-确认为是危急值

  --出参      json
  --output
  --    code                N   1 应答吗：0-失败；1-成功
  --    message             C   1 应答消息：失败时返回具体的错误信息
  --    cvalue_list[]         危急值列表，支持多个，[数组]
  --       cvalue_id               N   1  危急值id
  --       advice_id               N   1  医嘱id
  --       pat_name                C   1  病人姓名
  --       pat_sex                 C   1  病人性别
  --       pat_age                 C   1  病人年龄
  --       cvalue_rec_create_time  C   1  报告时间
  --       cvalue_rec_status       N   1  危急值状态
  --       cvitem_result           N   1  是否危急值
  --       cvalue_rec_desc         C   1  危急值说明
  --       cvitem_source           C   1  数据来源
  --       pati_id                 N   1  病人id
  --       pati_pageid             N   1  主页id
  --       rgst_no                 C   1  挂号单
  --       baby_num                N   1  婴儿
  --       lspcm_id                N   1  标本id
  --       rpt_deptid   N   1  报告科室id
  --       rec_rptor               C   1  报告人
  --       proc_note          C   1  处理情况
  --       cvalue_cnfmtime       C   1  确认时间
  --       cvalue_cnfmer         C   1  确认人
  --       cnfm_deptid         N   1  确认科室id
  --       cvalue_dept           C   1  确认科室
  -------------------------------------------------------------
  j_Json      Pljson;
  j_In        Pljson;
  n_Type      Number; --调用类型 ：1-通过id查询,2-通过挂号单查询,3-通过病人id加主页id查询
  v_危急值ids Varchar2(4000);
  v_挂号单    Varchar2(20);
  n_病人id    Number(18);
  n_主页id    Number(18);

  v_List Varchar2(32767);

  d_Begin      Date;
  d_End        Date;
  n_病人类型   Number;
  n_报告科室id Number;
  n_确认科室id Number;
  n_确认状态   Number;

Begin
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');
  n_Type := j_Json.Get_Number('use_type');

  If n_Type = 1 Then
    v_危急值ids := j_Json.Get_String('cvalue_ids');
  
    If v_危急值ids Is Null Then
      Json_Out := Zljsonout('未传入危急值id');
      Return;
    End If;
  
    Json_Out := '{"output":{"code":1,"message":"成功","cvalue_list":[';
  
    For c_危急值 In (Select a.Id, a.数据来源, a.病人id, a.主页id, a.挂号单, a.婴儿, a.姓名, a.性别, a.年龄, a.医嘱id, a.标本id, a.危急值描述,
                         To_Char(a.报告时间, 'YYYY-MM-DD HH24:MI') As 报告时间, a.报告科室id, a.报告人, a.处理情况,
                         To_Char(a.确认时间, 'YYYY-MM-DD HH24:MI') As 确认时间, a.确认人, a.确认科室id, a.状态, a.是否危急值, c.名称 As 确认科室
                  From 病人危急值记录 A,
                       (Select /*+cardinality(b,10)*/
                          Column_Value
                         From Table(f_Str2list(v_危急值ids))) B, 部门表 C
                  Where a.Id = b.Column_Value And a.确认科室id = c.Id(+)) Loop
      Zljsonputvalue(v_List, 'cvalue_id', c_危急值.Id, 1, 1);
      Zljsonputvalue(v_List, 'advice_id', c_危急值.医嘱id, 1);
      Zljsonputvalue(v_List, 'pat_name', c_危急值.姓名, 0);
      Zljsonputvalue(v_List, 'pat_sex', c_危急值.性别, 0);
      Zljsonputvalue(v_List, 'pat_age', c_危急值.年龄, 0);
      Zljsonputvalue(v_List, 'cvalue_rec_create_time', c_危急值.报告时间, 0);
      Zljsonputvalue(v_List, 'cvalue_rec_status', c_危急值.状态, 1);
      Zljsonputvalue(v_List, 'cvitem_result', c_危急值.是否危急值, 1);
      Zljsonputvalue(v_List, 'cvalue_rec_desc', c_危急值.危急值描述, 0);
      Zljsonputvalue(v_List, 'cvitem_source', c_危急值.数据来源, 0);
      Zljsonputvalue(v_List, 'pati_id', c_危急值.病人id, 1);
      Zljsonputvalue(v_List, 'pati_pageid', c_危急值.主页id, 1);
      Zljsonputvalue(v_List, 'rgst_no', c_危急值.挂号单, 0);
      Zljsonputvalue(v_List, 'baby_num', c_危急值.婴儿, 1);
      Zljsonputvalue(v_List, 'lspcm_id', c_危急值.标本id, 1);
      Zljsonputvalue(v_List, 'rpt_deptid', c_危急值.报告科室id, 1);
      Zljsonputvalue(v_List, 'rec_rptor', c_危急值.报告人, 0);
      Zljsonputvalue(v_List, 'proc_note', c_危急值.处理情况, 0);
      Zljsonputvalue(v_List, 'cvalue_cnfmtime', c_危急值.确认时间, 0);
      Zljsonputvalue(v_List, 'cvalue_cnfmer', c_危急值.确认人, 0);
      Zljsonputvalue(v_List, 'cnfm_deptid', c_危急值.确认科室id, 1);
      Zljsonputvalue(v_List, 'cvalue_dept', c_危急值.确认科室, 0, 2);
      If Length(v_List) > 20000 Then
        Json_Out := Json_Out || v_List;
        v_List   := ',';
      End If;
    End Loop;
  Elsif n_Type = 2 Then
    v_挂号单 := j_Json.Get_String('rgst_no');
  
    If v_挂号单 Is Null Then
      Json_Out := Zljsonout('未传入挂号单');
      Return;
    End If;
  
    Json_Out := '{"output":{"code":1,"message":"成功","cvalue_list":[';
  
    For c_危急值 In (Select a.Id, a.数据来源, a.病人id, a.主页id, a.挂号单, a.婴儿, a.姓名, a.性别, a.年龄, a.医嘱id, a.标本id, a.危急值描述,
                         To_Char(a.报告时间, 'YYYY-MM-DD HH24:MI') As 报告时间, a.报告科室id, a.报告人, a.处理情况,
                         To_Char(a.确认时间, 'YYYY-MM-DD HH24:MI') As 确认时间, a.确认人, a.确认科室id, a.状态, a.是否危急值
                  From 病人危急值记录 A
                  Where 挂号单 = v_挂号单) Loop
      Zljsonputvalue(v_List, 'cvalue_id', c_危急值.Id, 1, 1);
      Zljsonputvalue(v_List, 'advice_id', c_危急值.医嘱id, 1);
      Zljsonputvalue(v_List, 'pat_name', c_危急值.姓名, 0);
      Zljsonputvalue(v_List, 'pat_sex', c_危急值.性别, 0);
      Zljsonputvalue(v_List, 'pat_age', c_危急值.年龄, 0);
      Zljsonputvalue(v_List, 'cvalue_rec_create_time', c_危急值.报告时间, 0);
      Zljsonputvalue(v_List, 'cvalue_rec_status', c_危急值.状态, 1);
      Zljsonputvalue(v_List, 'cvitem_result', c_危急值.是否危急值, 1);
      Zljsonputvalue(v_List, 'cvalue_rec_desc', c_危急值.危急值描述, 0, 2);
      If Length(v_List) > 20000 Then
        Json_Out := Json_Out || v_List;
        v_List   := ',';
      End If;
    End Loop;
  
  Elsif n_Type = 3 Then
    n_病人id := j_Json.Get_Number('pati_id');
    n_主页id := j_Json.Get_Number('pati_pageid');
  
    If n_病人id Is Null Then
      Json_Out := Zljsonout('未传入病人id');
      Return;
    End If;
  
    Json_Out := '{"output":{"code":1,"message":"成功","cvalue_list":[';
  
    For c_危急值 In (Select a.Id, a.数据来源, a.病人id, a.主页id, a.挂号单, a.婴儿, a.姓名, a.性别, a.年龄, a.医嘱id, a.标本id, a.危急值描述,
                         To_Char(a.报告时间, 'YYYY-MM-DD HH24:MI') As 报告时间, a.报告科室id, a.报告人, a.处理情况,
                         To_Char(a.确认时间, 'YYYY-MM-DD HH24:MI') As 确认时间, a.确认人, a.确认科室id, a.状态, a.是否危急值
                  From 病人危急值记录 A
                  Where 病人id = n_病人id And 主页id = n_主页id) Loop
      Zljsonputvalue(v_List, 'cvalue_id', c_危急值.Id, 1, 1);
      Zljsonputvalue(v_List, 'advice_id', c_危急值.医嘱id, 1);
      Zljsonputvalue(v_List, 'pat_name', c_危急值.姓名, 0);
      Zljsonputvalue(v_List, 'pat_sex', c_危急值.性别, 0);
      Zljsonputvalue(v_List, 'pat_age', c_危急值.年龄, 0);
      Zljsonputvalue(v_List, 'cvalue_rec_create_time', c_危急值.报告时间, 0);
      Zljsonputvalue(v_List, 'cvalue_rec_status', c_危急值.状态, 1);
      Zljsonputvalue(v_List, 'cvitem_result', c_危急值.是否危急值, 1);
      Zljsonputvalue(v_List, 'cvalue_rec_desc', c_危急值.危急值描述, 0, 2);
      If Length(v_List) > 20000 Then
        Json_Out := Json_Out || v_List;
        v_List   := ',';
      End If;
    End Loop;
  Elsif n_Type = 4 Then
    d_Begin      := To_Date(j_Json.Get_String('cvalue_time_begin'), 'YYYY-MM-DD HH24:MI:SS');
    d_End        := To_Date(j_Json.Get_String('cvalue_time_end'), 'YYYY-MM-DD HH24:MI:SS');
    n_病人类型   := Nvl(j_Json.Get_Number('pati_type'), 0);
    n_报告科室id := Nvl(j_Json.Get_Number('rpt_deptid'), 0);
    n_确认科室id := Nvl(j_Json.Get_Number('cnfm_deptid'), 0);
    n_确认状态   := Nvl(j_Json.Get_Number('cvalue_rec_status'), 0);
  
    Json_Out := '{"output":{"code":1,"message":"成功","cvalue_list":[';
  
    For c_危急值 In (Select a.Id, a.数据来源, a.病人id, a.主页id, a.挂号单, a.婴儿, a.姓名, a.性别, a.年龄, a.医嘱id, a.标本id, a.危急值描述,
                         To_Char(a.报告时间, 'YYYY-MM-DD HH24:MI') As 报告时间, a.报告科室id, a.报告人, a.处理情况,
                         To_Char(a.确认时间, 'YYYY-MM-DD HH24:MI') As 确认时间, a.确认人, a.确认科室id, a.状态, a.是否危急值, c.名称 As 确认科室
                  From 病人危急值记录 A, 部门表 C
                  Where a.报告时间 Between d_Begin And d_End And
                        ((Nvl(n_报告科室id, 0) = a.报告科室id Or Nvl(n_报告科室id, 0) = 0) And
                        (Nvl(n_病人类型, 0) = 0 Or (Nvl(n_病人类型, 0) = 1 And a.挂号单 Is Not Null) Or
                        (Nvl(n_病人类型, 0) = 2 And Nvl(a.主页id, 0) > 0) Or
                        (Nvl(n_病人类型, 0) = 3 And Nvl(a.主页id, 0) = 0 And a.挂号单 Is Null)) And
                        (Nvl(n_确认科室id, 0) = a.确认科室id Or Nvl(n_确认科室id, 0) = 0) And
                        (Nvl(n_确认状态, 0) = 0 Or (Nvl(n_确认状态, 0) = 1 And a.状态 = 1) Or
                        (Nvl(n_确认状态, 0) = 2 And a.状态 = 2 And Nvl(a.是否危急值, 0) = 0) Or
                        (Nvl(n_确认状态, 0) = 3 And a.状态 = 2 And Nvl(a.是否危急值, 0) = 1))) And a.确认科室id = c.Id(+)
                  Order By a.报告时间 Desc) Loop
    
      Zljsonputvalue(v_List, 'cvalue_id', c_危急值.Id, 1, 1);
      Zljsonputvalue(v_List, 'advice_id', c_危急值.医嘱id, 1);
      Zljsonputvalue(v_List, 'pat_name', c_危急值.姓名, 0);
      Zljsonputvalue(v_List, 'pat_sex', c_危急值.性别, 0);
      Zljsonputvalue(v_List, 'pat_age', c_危急值.年龄, 0);
      Zljsonputvalue(v_List, 'cvalue_rec_create_time', c_危急值.报告时间, 0);
      Zljsonputvalue(v_List, 'cvalue_rec_status', c_危急值.状态, 1);
      Zljsonputvalue(v_List, 'cvitem_result', c_危急值.是否危急值, 1);
      Zljsonputvalue(v_List, 'cvalue_rec_desc', c_危急值.危急值描述, 0);
      Zljsonputvalue(v_List, 'cvitem_source', c_危急值.数据来源, 0);
      Zljsonputvalue(v_List, 'pati_id', c_危急值.病人id, 1);
      Zljsonputvalue(v_List, 'pati_pageid', c_危急值.主页id, 1);
      Zljsonputvalue(v_List, 'rgst_no', c_危急值.挂号单, 0);
      Zljsonputvalue(v_List, 'baby_num', c_危急值.婴儿, 1);
      Zljsonputvalue(v_List, 'lspcm_id', c_危急值.标本id, 1);
      Zljsonputvalue(v_List, 'rpt_deptid', c_危急值.报告科室id, 1);
      Zljsonputvalue(v_List, 'rec_rptor', c_危急值.报告人, 0);
      Zljsonputvalue(v_List, 'proc_note', c_危急值.处理情况, 0);
      Zljsonputvalue(v_List, 'cvalue_cnfmtime', c_危急值.确认时间, 0);
      Zljsonputvalue(v_List, 'cvalue_cnfmer', c_危急值.确认人, 0);
      Zljsonputvalue(v_List, 'cnfm_deptid', c_危急值.确认科室id, 1);
      Zljsonputvalue(v_List, 'cvalue_dept', c_危急值.确认科室, 0, 2);
      If Length(v_List) > 20000 Then
        Json_Out := Json_Out || v_List;
        v_List   := ',';
      End If;
    End Loop;
  End If;

  If v_List <> ',' Then
    Json_Out := Json_Out || v_List || ']}}';
  Else
    Json_Out := Json_Out || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getcriticalinfo;
/
Create Or Replace Procedure Zl_Cissvr_Getdiagfitsituation
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:获取诊断符合情况
  --入参：Json_In:格式
  -- input
  --   pati_id              N 1 病人id
  --   pati_pageid          N 1 主页id
  --出参: Json_Out,格式如下
  --  output
  --    code               C  1 应答码：0-失败；1-成功
  --    message            C  1 应答消息：失败时返回具体的错误信息
  --    fit_list[]      [数组]
  --      fit_type         N  1 1.门诊与出院、2.入院与出院、3.放射与病理、4.临床与病理、5.临床与尸检、6.术前与术后、11.中医门诊与出院、12.中医入院与出院、13.中医辨证、14.中医治法、15.中医方药
  --      diag_cnst        N  1 诊断符合情况: 0.未做、1.符合、2.不符合、3.不肯定；对于中医准确度：1.准确、2.基本准确、3.重大缺陷、4.错误

  ---------------------------------------------------------------------------
  j_In     Pljson;
  j_Json   Pljson;
  n_病人id 诊断符合情况.病人id%Type;
  n_主页id 诊断符合情况.主页id%Type;
  v_List   Varchar2(32767);

Begin
  --解析入参
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');
  n_主页id := j_Json.Get_Number('pati_pageid');

  If Nvl(n_病人id, 0) = 0 Or Nvl(n_主页id, 0) = 0 Then
    Json_Out := Zljsonout('必须传入病人id和主页id！');
    Return;
  End If;

  For r_诊断 In (Select 符合类型, Nvl(符合情况, 0) As 符合情况
               From 诊断符合情况
               Where 病人id = n_病人id And 主页id = n_主页id) Loop
    v_List := v_List || ',{"fit_type":' || Nvl(r_诊断.符合类型 || '', 'null');
    v_List := v_List || ',"diag_cnst":' || r_诊断.符合情况;
    v_List := v_List || '}';
  
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"成功","fit_list":[' || Substr(v_List, 2) || ']}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getdiagfitsituation;
/
Create Or Replace Procedure Zl_Cissvr_Getdiaginfo
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:获取病人诊断信息
  --入参：Json_In:格式
  -- input
  --   advice_ids           C 1  医嘱ids,医嘱id拼串
  --   query_type           N 1 查询方式1-按指定条件查询,2-仅按病人id,主页id查询诊断
  --   pati_info            C 0  病人id等其他信息
  --     pati_id            N 1 病人id
  --     pati_pageid        N 1 主页id
  --     diag_types         C 1  诊断类型:0-所有类型,1-西医门诊诊断;2-西医入院诊断;3-西医出院诊断;5-院内感染;6-病理诊断;7-损伤中毒码,8-术前诊断;9-术后诊断;10-并发症;11-中医门诊诊断;12-中医入院诊断;13-中医出院诊断;21-病原学诊断;22-影像学诊断
  --                            可以为多个诊断类型，用逗号分离,如:2,12
  --     rec_source         N 1 记录来源:1-病历；2-入院登记；3-首页整理;4-病案;NULL-不作限制
  --     diag_num           N 1 诊断次序:NULL表示不作限制
  --     code_type          C 1  编码类别:ICD-11的编码编码类别为'E',空时表示读取ICD-10等
  --     input_num          C 1  录入次序:启用了ICD-11编码录入后，诊断的录入次序
  --     rec_sources        C 1 记录来源拼串

  --出参: Json_Out,格式如下
  --  output
  --    code                N  1 应答码：0-失败；1-成功
  --    message             C  1 应答消息：失败时返回具体的错误信息
  --    diag_list     [数组]
  --      diag_type         N 1 诊断类型:1-西医门诊诊断;2-西医入院诊断;3-西医出院诊断;5-院内感染;6-病理诊断;7-损伤中毒码,8-术前诊断;9-术后诊断;10-并发症;11-中医门诊诊断;
  --                                     12-中医入院诊断;13-中医出院诊断;21-病原学诊断;22-影像学诊断
  --      diag_num          N 1 诊断序号
  --      code_num          N 1 编码序号
  --      dz_id             N 1 疾病ID
  --      dz_code           C 1 疾病编码
  --      diag_note         C 1 诊断描述
  --      recoder           C 1 记录人
  --      rec_time          C 1 记录时间:yyyy-mm-dd hh24:mi:ss
  --      adtd_rsn          C 1 出院情况:治愈、好转、未愈、死亡、其他
  --      diag_id           N 1 诊断id
  --      diag_rec_id       N 1 诊断记录ID:病人诊断记录.ID
  --      diag_doubt        N 1 是否疑诊
  --      advice_id         N   医嘱id(根据医嘱ids查询时才返回)
  --      advice_main_id    N   组医嘱id(根据医嘱ids查询时才返回)
  --      advice_related_id N   相关id(根据医嘱ids查询时才返回)
  --      rec_source        N   记录来源:1-病历；2-入院登记；3-首页整理;4-病案;NULL-不作限制
  ---------------------------------------------------------------------------
  j_In       Pljson;
  j_Json     Pljson;
  j_Patiinfo Pljson;
  c_医嘱ids  Clob;
  v_医嘱ids  Varchar2(32767);
  n_病人id   病人诊断记录.病人id%Type;
  n_主页id   病人诊断记录.主页id%Type;
  n_记录来源 病人诊断记录.记录来源%Type;
  v_诊断类型 Varchar2(3000);
  n_诊断次序 病人诊断记录.诊断次序%Type;
  v_编码类别 病人诊断记录.编码类别%Type;
  v_录入次序 病人诊断记录.录入次序%Type;
  I          Number := 0;
  n_查询类型 Number := 0;
  v_诊断来源 Varchar2(100);
  v_List     Varchar2(32767);

  --组装失败时返回的数据
  Function Get_Err_Message(Message_In Varchar2) Return Varchar2 Is
  Begin
    Return '{"output":{"code":0,"message":"' || Zljsonstr(Message_In) || '"}}';
  End Get_Err_Message;
Begin
  --解析入参
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');

  Json_Out := '{"output":{"code":1,"message":"成功","diag_list":[';

  If j_Json.Get_Number('query_type') = 2 Then
    j_Patiinfo := j_Json.Get_Pljson('pati_info');
    If Not j_Patiinfo Is Null Then
      n_病人id := j_Patiinfo.Get_Number('pati_id');
      n_主页id := j_Patiinfo.Get_Number('pati_pageid');
    End If;
    For r_诊断 In (Select 诊断类型, 诊断次序, 诊断描述, 出院情况, 记录来源
                 From 病人诊断记录
                 Where 记录来源 In (2, 3) And Nvl(编码序号, 1) = 1 And 病人id = n_病人id And 主页id = n_主页id And
                       Nvl(录入次序, '01') = '01'
                 Order By 记录来源 Desc) Loop
    
      Zljsonputvalue(v_List, 'diag_type', r_诊断.诊断类型, 1, 1);
      Zljsonputvalue(v_List, 'diag_num', r_诊断.诊断次序, 1);
      Zljsonputvalue(v_List, 'rec_source', r_诊断.记录来源, 1);
      Zljsonputvalue(v_List, 'diag_note', r_诊断.诊断描述);
      Zljsonputvalue(v_List, 'adtd_rsn', r_诊断.出院情况, 0, 2);
      If Length(v_List) > 20000 Then
        Json_Out := Json_Out || v_List;
        v_List   := ',';
      End If;
    End Loop;
    If v_List <> ',' Then
      Json_Out := Json_Out || v_List || ']}}';
    Else
      Json_Out := Json_Out || ']}}';
    End If;
    Return;
  End If;
  Begin
    c_医嘱ids := j_Json.Get_Clob('advice_ids');
  Exception
    When Others Then
      c_医嘱ids := Null;
  End;
  If c_医嘱ids Is Null Then
    j_Patiinfo := Pljson();

    j_Patiinfo := j_Json.Get_Pljson('pati_info');
    If Not j_Patiinfo Is Null Then
      n_病人id   := j_Patiinfo.Get_Number('pati_id');
      n_主页id   := j_Patiinfo.Get_Number('pati_pageid');
      v_诊断类型 := j_Patiinfo.Get_String('diag_types');
      n_记录来源 := j_Patiinfo.Get_Number('rec_source');
      n_诊断次序 := j_Patiinfo.Get_Number('diag_num');
      v_编码类别 := j_Patiinfo.Get_String('code_type');
      v_录入次序 := j_Patiinfo.Get_String('input_num');
      v_诊断来源 := j_Patiinfo.Get_String('rec_sources');
      If Nvl(n_病人id, 0) = 0 Or Nvl(n_主页id, 0) = 0 Then
        Json_Out := Get_Err_Message('未传入病人id或主页id，请检查！');
        Return;
      End If;
    End If;
  End If;

  If c_医嘱ids Is Null Then
    If n_查询类型 = 0 Then
      For r_诊断 In (Select a.诊断类型, a.诊断次序, a.编码序号, b.Id As 疾病id, b.编码 As 疾病编码, a.诊断描述, a.记录人,
                          To_Char(Nvl(a.记录日期, Sysdate), 'yyyy-mm-dd hh24:mi:ss') As 记录日期, a.出院情况, a.Id As 诊断记录id, a.诊断id,
                          a.是否疑诊
                   From 病人诊断记录 A, 疾病编码目录 B
                   Where a.疾病id = b.Id And (a.记录来源 = n_记录来源 Or n_记录来源 Is Null) And (a.诊断次序 = n_诊断次序 Or n_诊断次序 Is Null) And
                         (a.诊断类型 In (Select Column_Value From Table(f_Str2list(v_诊断类型)) C) Or Nvl(v_诊断类型, 0) = 0) And
                         a.病人id = n_病人id And a.主页id = n_主页id And Nvl(a.录入次序, '01') = Nvl(v_录入次序, '01') And
                         Nvl(a.编码类别, 'E') = Nvl(v_编码类别, 'E')) Loop
        --      diag_type         N 1 诊断类型:1-西医门诊诊断;2-西医入院诊断;3-西医出院诊断;5-院内感染;6-病理诊断;7-损伤中毒码,8-术前诊断;9-术后诊断;10-并发症;11-中医门诊诊断;
        --                                     12-中医入院诊断;13-中医出院诊断;21-病原学诊断;22-影像学诊断
        --      diag_num          N 1 诊断序号
        --      code_num          N 1 编码序号
        --      dz_id             N 1 疾病ID
        --      dz_code           C 1 疾病编码
        --      diag_note         C 1 诊断描述
        --      recoder           C 1 记录人
        --      rec_time          C 1 记录时间:yyyy-mm-dd hh24:mi:ss
        --      adtd_rsn          C 1 出院情况:治愈、好转、未愈、死亡、其他
        --      diag_id           N 1 诊断id
        --      diag_rec_id       N 1 诊断记录ID:病人诊断记录.ID
        --      diag_doubt        N 1 是否疑诊
        Zljsonputvalue(v_List, 'diag_type', r_诊断.诊断类型, 1, 1);
        Zljsonputvalue(v_List, 'diag_num', r_诊断.诊断次序, 1);
        Zljsonputvalue(v_List, 'code_num', r_诊断.编码序号);
        Zljsonputvalue(v_List, 'dz_id', r_诊断.疾病id, 1);
        Zljsonputvalue(v_List, 'dz_code', r_诊断.疾病编码);
        Zljsonputvalue(v_List, 'diag_note', r_诊断.诊断描述);
        Zljsonputvalue(v_List, 'recoder', r_诊断.记录人);
        Zljsonputvalue(v_List, 'rec_time', r_诊断.记录日期);
        Zljsonputvalue(v_List, 'adtd_rsn', r_诊断.出院情况);
        Zljsonputvalue(v_List, 'diag_id', r_诊断.诊断id, 1);
        Zljsonputvalue(v_List, 'diag_rec_id', r_诊断.诊断记录id, 1);
        Zljsonputvalue(v_List, 'diag_doubt', r_诊断.是否疑诊, 1, 2);
      
        If Length(v_List) > 20000 Then
          Json_Out := Json_Out || v_List;
          v_List   := ',';
        End If;
      End Loop;
      If v_List <> ',' Then
        Json_Out := Json_Out || v_List || ']}}';
      Else
        Json_Out := Json_Out || ']}}';
      End If;
    Else
      For r_诊断 In (Select 诊断类型, 疾病id, 诊断描述, 出院情况
                   From 病人诊断记录
                   Where ((诊断次序 = 1 And n_诊断次序 = 0) Or (诊断次序 > 1 And n_诊断次序 = 1) Or Nvl(n_诊断次序, 0) = 0) And
                         记录来源 In (v_诊断来源) And Nvl(编码序号, 1) = 1 And 病人id = n_病人id And 主页id = n_主页id And
                         (诊断类型 In (v_诊断类型) Or Nvl(v_诊断类型, '-') = '-') And Nvl(录入次序, '01') = '01'
                   Order By 记录来源 Desc) Loop
        Zljsonputvalue(v_List, 'diag_type', r_诊断.诊断类型, 1, 1);
        Zljsonputvalue(v_List, 'dz_id', r_诊断.疾病id, 1);
        Zljsonputvalue(v_List, 'diag_note', r_诊断.诊断描述);
        Zljsonputvalue(v_List, 'adtd_rsn', r_诊断.出院情况, 0, 2);
      
        If Length(v_List) > 20000 Then
          Json_Out := Json_Out || v_List;
          v_List   := ',';
        End If;
      End Loop;
      If v_List <> ',' Then
        Json_Out := Json_Out || v_List || ']}}';
      Else
        Json_Out := Json_Out || ']}}';
      End If;
    End If;
  
  Else
    While c_医嘱ids Is Not Null Loop
      If Length(c_医嘱ids) <= 4000 Then
        v_医嘱ids := c_医嘱ids;
        c_医嘱ids := Null;
      Else
        v_医嘱ids := Substr(c_医嘱ids, 1, Instr(c_医嘱ids, ',', 3980) - 1);
        c_医嘱ids := Substr(c_医嘱ids, Instr(c_医嘱ids, ',', 3980) + 1);
      End If;
      I := I + 1;
    
      For r_诊断 In (Select a.组id, a.医嘱id, a.相关id, c.诊断类型, c.诊断次序, c.编码序号, c.诊断描述, c.记录人,
                          To_Char(Nvl(c.记录日期, Sysdate), 'yyyy-mm-dd hh24:mi:ss') As 记录日期, c.出院情况, c.Id As 诊断记录id, c.诊断id,
                          c.是否疑诊
                   From (Select Nvl(a.相关id, a.Id) As 组id, a.Id As 医嘱id, a.相关id
                          From 病人医嘱记录 A
                          Where a.Id In (Select /*+cardinality(x,10)*/
                                          x.Column_Value As 医嘱id
                                         From Table(Cast(f_Num2list(v_医嘱ids) As Zltools.t_Numlist)) X)) A, 病人诊断医嘱 B,
                        病人诊断记录 C
                   Where a.组id = b.医嘱id And b.诊断id = c.Id
                   Order By c.Id) Loop
      
        --      diag_type         N 1 诊断类型:1-西医门诊诊断;2-西医入院诊断;3-西医出院诊断;5-院内感染;6-病理诊断;7-损伤中毒码,8-术前诊断;9-术后诊断;10-并发症;11-中医门诊诊断;
        --                                     12-中医入院诊断;13-中医出院诊断;21-病原学诊断;22-影像学诊断
        --      diag_num          N 1 诊断序号
        --      code_num          N 1 编码序号
        --      diag_note         C 1 诊断描述
        --      recoder           C 1 记录人
        --      rec_time          C 1 记录时间:yyyy-mm-dd hh24:mi:ss
        --      adtd_rsn          C 1 出院情况:治愈、好转、未愈、死亡、其他
        --      diag_id           N 1 诊断id
        --      diag_rec_id       N 1 诊断记录ID:病人诊断记录.ID
        --      diag_doubt        N 1 是否疑诊
        --      advice_id         N   医嘱id(根据医嘱ids查询时才返回)
        --      advice_main_id    N   组医嘱id(根据医嘱ids查询时才返回)
        --      advice_related_id N   相关id(根据医嘱ids查询时才返回)
        Zljsonputvalue(v_List, 'advice_id', r_诊断.医嘱id, 1, 1);
        Zljsonputvalue(v_List, 'advice_main_id', r_诊断.组id, 1);
        Zljsonputvalue(v_List, 'advice_related_id', r_诊断.相关id, 1);
        Zljsonputvalue(v_List, 'diag_type', r_诊断.诊断类型, 1);
        Zljsonputvalue(v_List, 'diag_num', r_诊断.诊断次序, 1);
        Zljsonputvalue(v_List, 'code_num', r_诊断.编码序号);
        Zljsonputvalue(v_List, 'diag_note', r_诊断.诊断描述);
        Zljsonputvalue(v_List, 'recoder', r_诊断.记录人);
        Zljsonputvalue(v_List, 'rec_time', r_诊断.记录日期);
        Zljsonputvalue(v_List, 'adtd_rsn', r_诊断.出院情况);
        Zljsonputvalue(v_List, 'diag_id', r_诊断.诊断id, 1);
        Zljsonputvalue(v_List, 'diag_rec_id', r_诊断.诊断记录id, 1);
        Zljsonputvalue(v_List, 'diag_doubt', r_诊断.是否疑诊, 1, 2);
      
        If Length(v_List) > 20000 Then
          Json_Out := Json_Out || v_List;
          v_List   := ',';
        End If;
      
      End Loop;
    End Loop;
  
    If v_List <> ',' Then
      Json_Out := Json_Out || v_List || ']}}';
    Else
      Json_Out := Json_Out || ']}}';
    End If;
  
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getdiaginfo;
/
Create Or Replace Procedure Zl_Cissvr_Getdrugerrdata
(
  Json_In  In Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：临床医嘱发送生成处方,卫才,静配,数据同步
  --入参：Json_In:格式
  --  input
  --      pati_ids                        C 1 病人ids逗号拼串  
  --出参: Json_Out,格式如下
  --   output
  --     code                             N 1 应答吗：0-失败；1-成功
  --     message                          C 1 应答消息：失败时返回具体的错误信息
  --     pati_bill_list[]                 病人医嘱费用单据信息
  --         pati_id                      N 1 病人id
  --         pati_pageid                  N 0 主页id，住院病人传入，门诊传0
  --         rgst_id                      N 0 挂号id，门诊病人传入，住院病人传空
  --         rgst_no                      C 0 挂号单号
  --         send_no                      N 1 发送号
  --         operator_name                C 1 发送人(操作员姓名)
  --         operator_time                C 1 发送时间
  --         pati_type                    C 0 病人类型，当为住院病人时可以获取，如果是门诊病人需外部单独获取
  --         diag_list[]                  临床诊断信息和
  --             diag_rec_id              N 1 诊断记录id
  --             diag_type                N 1 诊断类型 1-西医门诊诊断;2-西医入院诊断;3-西医出院诊断;5-院内感染;6-病理诊断;7-损伤中毒码,8-术前诊断;9-术后诊断;10-并发症;11-中医门诊诊断;12-中医入院诊断;13-中医出院诊断;21-病原学诊断;22-影像学诊断
  --             diag_name                C 1 诊断名称
  --         bill_list[]医嘱信息列表
  --             advice_id                N 1 医嘱id
  --             group_sno                N 0 组内序号 (包括存储)：1、2、3
  --             effectivetime            N  0 医嘱期效
  --             drug_method_id           N 1 给药途径id(新门诊):诊疗项目ID: 204,
  --             drug_method_name         C 1 给药途径名称: 静脉滴入,
  --             drug_method_class_code   C 1 给药途径分类(新门诊):执行分类编号: 001,
  --             drug_freq_id             N 1 给药频次id(新门诊):诊疗频率编码: 1,
  --             drug_freq_name           C 1 给药频次名称(新门诊):: 每天二次,
  --             emergency_tag            N 1 医嘱记录中的紧急标志(0-普通;1-紧急;2-补录(对门诊无效))
  --             fee_mode                 N 1 计价特性：0-正常计价；1-不计价；2-手工计价
  --             use_mode                 N 1 取药特性：0-正常方式，1-离院带药，2-自取药
  --             frequency                N 0 频次: 2,
  --             single                   N 0 单量: 1,
  --             usage                    C 0 用法: 静脉滴入,
  --             rcpdtl_st_result         N 0 皮试结果(新门诊)1-阴性，2-阳性，3-免试，4-连续用药 处方生成时已确定或已有皮试结果。ZLHIS目前支持不全: ,
  --             rcpdtL_excs_desc         C 0 超量说明(新门诊): ,
  --             rcpdtL_drask             C 0 使用嘱托(新门诊): ,
  --             memo                     C 0 摘要: 医嘱发送,
  --             diag_name                C 0 诊断名称（新门诊)仅门诊传入，诊断描述:
  --             take_no                  C 0 领药号
  --             advice_purpose           C 0 用药目的
  --             fee_source               N 0 费用来源：1-门诊；2-住院
  --             fee_billtype             N 0 费用单据类型：1-收费处方；2-记帐单处方
  --             fee_no                   C 0 费用单据号
  --         pivas_list[] 静配信息列表，只有一个元素
  --             pati_id                  N 1 病人id
  --             page_id                  N 1 主页id
  --             pati_name                C 1 姓名
  --             pati_sex                 C 1 性别
  --             pati_age                 C 1 年龄
  --             inpatient_num            C 1 住院号
  --             pati_bed                 C 1 床号
  --             pati_wardarea_id         N 1 病人病区id
  --             pati_deptid              N 1 病人科室id
  --             advice_list[]主医嘱，数组
  --                 pivas_deptid         N 1 静配中心id
  --                 advice_id            N 1 主医嘱ID(给药途径)
  --                 effective_time       N 1 医嘱期效
  --                 drug_method_id       N 1 给药途径id
  --                 is_tpn               N 1 是否tpn
  --                 advice_frequency     C 1 执行频次
  --                 advice_drug_list[]给药途径对应的药嘱，数组
  --                     advice_id        N 1 药嘱id
  --                     advice_rcpno     C 1 药嘱发送产生的费用no
  --                 advice_exetime_list[]医嘱执行时间，给药途径医嘱的执行时间，暂时提供该医嘱所有发送的时间，包括本次发送的执行时间。按发送号倒序组织数组数据
  --                     advice_send_no   N 1 给药途径医嘱的发送号
  --                     advice_require_time  C 1 要求时间: 2019-11-30 23:00:00
  ------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  b_Pivasout Clob;
  v_操作员   Varchar2(300);
  v_操作时间 Varchar2(40);

  c_Outtmp   Clob;
  v_Jtmp     Varchar2(32767);
  v_Tmp      Varchar2(32767);
  n_Cnt      Number;
  v_Err      Varchar2(32767);
  v_P医嘱ids Varchar2(32767);
  n_组id     Number;
  j_Json     Pljson;
  j_Tmp      Pljson;
  n_行号     Number;
  n_Send_No  Number;
  v_Vals     Clob;
  l_Vals     t_Strlist;
  Err_Item Exception;

  n_Rgst_Id                Number; --N   1挂号单id（新门诊)
  v_Take_No                Varchar2(32767); --C 0 领药号 领药号，未发药品记录.领药号，药品收发记录.产品合格证，医嘱发送时生成
  n_Group_Sno              Number; --N   组内序号(包括存储)：1、2、3
  n_Cadn_Id                Number; --N   1药品通用名称id(药名ID)(新门诊)
  n_Advice_Id              Number; --N   0医嘱ID
  n_Drug_Method_Id         Number; --N   1给药途径id(新门诊):诊疗项目ID
  v_Drug_Method_Name       Varchar2(32767); --C   1给药途径名称
  v_Drug_Method_Class_Code Varchar2(32767); --C   1给药途径分类(新门诊):执行分类编号
  n_Drug_Freq_Id           Number; --N   1给药频次id(新门诊):诊疗频率编码
  v_Drug_Freq_Name         Varchar2(32767); --C   1给药频次名称d(新门诊):
  n_Emergency_Tag          Number; --N   医嘱记录中的紧急标志(0-普通;1-紧急;2-补录(对门诊无效))
  n_Effectivetime          Number; --N   0医嘱期效
  n_Denominated            Number; --N   0计价特性：0-2-计价特性,3-离院带药,4-自取药计价特性(0-正常计价；1-不计价；2-手工计价))
  n_Frequency              Number; --N   0频次
  n_Single                 Number; --N   0单量
  v_Usage                  Varchar2(32767); --C   0用法
  v_Rcpdtl_St_Result       Varchar2(32767); --N   皮试结果(新门诊)1-阴性，2-阳性，3-免试，4-连续用药处方生成时已确定或已有皮试结果。ZLHIS目前支持不全
  v_Rcpdtl_Drask           Varchar2(32767); --C   使用嘱托(新门诊)
  v_Diag_Name              Varchar2(32767); --C  0 诊断名称（新门诊)仅门诊传入，诊断描述
  n_Use_Mode               Number;
  n_费用来源               Number;
  n_配液标记               Number(3);

  Vj_Iitem  Varchar2(32767);
  Cjl_Iitem Clob;
  --门诊医嘱数据
  Cursor c_Out
  (
    挂号单_In 病人医嘱记录.挂号单%Type,
    发送号_In 病人医嘱发送.发送号%Type
  ) Is
    Select a.主页id, a.Id, a.相关id, b.No, a.紧急标志, a.计价特性, a.单次用量, d.名称 As 用法, b.记录性质, b.门诊记帐, a.皮试结果, '医嘱发送' As 摘要, a.用药目的,
           a.超量说明, a.医生嘱托, a.医嘱期效, a.频率次数, a.诊疗类别, a.诊疗项目id, c.诊疗项目id As 给药途径id, d.名称 As 给药途径名称, d.类别 给药类别, Null 给药操作类型,
           d.执行分类 As 给药途径分类, e.编码 As 诊疗频率编码, c.执行频次 As 给药频次名称, b.领药号, b.发送号, a.执行标记, a.执行性质 As 药行执行性质, c.执行性质 As 给药执行性质
    From 病人医嘱记录 A, 病人医嘱发送 B, 病人医嘱记录 C, 诊疗项目目录 D, 诊疗频率项目 E
    Where a.相关id = c.Id(+) And c.诊疗项目id = d.Id(+) And c.执行频次 = e.名称(+) And a.Id = b.医嘱id And a.挂号单 = 挂号单_In And
          b.发送号 = 发送号_In
    Order By a.序号, b.发送号;

  Type t_Order Is Table Of c_Out%RowType;
  r_Odr t_Order;

  Cursor c_In
  (
    病人id_In 病人医嘱记录.病人id%Type,
    主页id_In 病人医嘱记录.主页id%Type,
    发送号_In 病人医嘱发送.发送号%Type
  ) Is
    Select a.主页id, a.Id, a.相关id, b.No, a.紧急标志, a.计价特性, a.单次用量, d.名称 As 用法, b.记录性质, b.门诊记帐, a.皮试结果, '医嘱发送' As 摘要, a.用药目的,
           a.超量说明, a.医生嘱托, a.医嘱期效, a.频率次数, a.诊疗类别, a.诊疗项目id, c.诊疗项目id As 给药途径id, d.名称 As 给药途径名称, d.类别 给药类别,
           d.操作类型 给药操作类型, d.执行分类 As 给药途径分类, e.编码 As 诊疗频率编码, c.执行频次 As 给药频次名称, b.领药号, b.发送号, a.执行标记, a.执行性质 As 药行执行性质,
           c.执行性质 As 给药执行性质
    From 病人医嘱记录 A, 病人医嘱发送 B, 病人医嘱记录 C, 诊疗项目目录 D, 诊疗频率项目 E
    Where a.相关id = c.Id(+) And c.诊疗项目id = d.Id(+) And c.执行频次 = e.名称(+) And a.Id = b.医嘱id And a.病人id = 病人id_In And
          a.主页id = 主页id_In And b.发送号 = 发送号_In
    Order By a.序号, b.发送号;
  -- 诊疗类别
  Cursor c_Adv(P组id Number) Is
    Select a.Id From 病人医嘱记录 A Where a.Id = P组id Or a.相关id = P组id Order By a.序号;

  --门诊医嘱诊断数据
  Cursor c_Diag
  (
    病人id_In 病人医嘱记录.病人id%Type,
    挂号id_In 病人诊断记录.主页id%Type
  ) Is
    Select a.医嘱id, b.诊断描述 As 描述
    From 病人诊断医嘱 A, 病人诊断记录 B
    Where a.诊断id = b.Id And b.病人id = 病人id_In And b.主页id = 挂号id_In And Nvl(b.录入次序, '01') = '01'
    Order By a.医嘱id;

  Type t_Diag Is Table Of c_Diag%RowType;
  Rs_Diag t_Diag;

  Cursor Ctest(挂号单_In 病人医嘱记录.挂号单%Type) Is
    Select a.皮试结果, a.药名id, a.药品id
    From (Select a.皮试结果, c.项目id As 药名id, 0 As 药品id, a.开始执行时间
           From 病人医嘱记录 A, 诊疗用法用量 C
           Where a.诊疗项目id = c.用法id And Nvl(c.性质, 0) = 0 And Nvl(a.医嘱状态, 0) = 8 And a.皮试结果 Is Not Null And a.挂号单 = 挂号单_In
           Union All
           Select a.皮试结果, b.药名id, b.药品id, a.开始执行时间
           From 病人医嘱记录 A, 药品规格 B, 药品用法用量 C
           Where a.诊疗项目id = c.用法id And b.药品id = c.药品id And Nvl(c.性质, 0) = 0 And Nvl(a.医嘱状态, 0) = 8 And a.皮试结果 <> '免试' And
                 a.挂号单 = 挂号单_In
           Union All
           Select a.皮试结果, c.项目id As 药名id, 0 As 药品id, a.开始执行时间
           From 病人医嘱记录 A, 诊疗用法用量 C
           Where a.诊疗项目id = c.用法id And Nvl(c.性质, 0) = 0 And a.皮试结果 = '免试' And a.挂号单 = 挂号单_In) A
    Order By a.开始执行时间 Desc;
  Type t_Test Is Table Of Ctest%RowType;
  Rs_Test t_Test;

  Cursor Ctestin
  (
    病人id_In 病人医嘱记录.病人id%Type,
    主页id_In 病人医嘱记录.主页id%Type
  ) Is
    Select a.皮试结果, a.药名id, a.药品id
    From (Select a.皮试结果, c.项目id As 药名id, 0 As 药品id, a.开始执行时间
           From 病人医嘱记录 A, 诊疗用法用量 C
           Where a.诊疗项目id = c.用法id And Nvl(c.性质, 0) = 0 And Nvl(a.医嘱状态, 0) = 8 And a.皮试结果 Is Not Null And
                 a.病人id = 病人id_In And a.主页id = 主页id_In
           Union All
           Select a.皮试结果, b.药名id, b.药品id, a.开始执行时间
           From 病人医嘱记录 A, 药品规格 B, 药品用法用量 C
           Where a.诊疗项目id = c.用法id And b.药品id = c.药品id And Nvl(c.性质, 0) = 0 And Nvl(a.医嘱状态, 0) = 8 And a.皮试结果 <> '免试' And
                 a.病人id = 病人id_In And a.主页id = 主页id_In
           Union All
           Select a.皮试结果, c.项目id As 药名id, 0 As 药品id, a.开始执行时间
           From 病人医嘱记录 A, 诊疗用法用量 C
           Where a.诊疗项目id = c.用法id And Nvl(c.性质, 0) = 0 And a.皮试结果 = '免试' And a.病人id = 病人id_In And a.主页id = 主页id_In) A
    Order By a.开始执行时间 Desc;

  --获取皮试结果
  Procedure Pp_Test
  (
    药名id_In    Number,
    皮试结果_Out Out Varchar2
  ) Is
  Begin
    皮试结果_Out := Null;
    For I In 1 .. Rs_Test.Count Loop
      If Rs_Test(I).药名id = 药名id_In Then
        皮试结果_Out := Rs_Test(I).皮试结果;
		exit;
      End If;
    End Loop;
  End;

  --获取诊断
  Procedure Pp_Diag
  (
    医嘱id_In Number,
    诊断_Out  Out Varchar2
  ) Is
  Begin
    诊断_Out := Null;
    For I In 1 .. Rs_Diag.Count Loop
      If Rs_Diag(I).医嘱id = 医嘱id_In Then
        诊断_Out := 诊断_Out || ',' || Rs_Diag(I).描述;
      End If;
    End Loop;
    诊断_Out := Substr(诊断_Out, 2);
  End;

  Procedure p_Getbaseinfo
  (
    Rgstid_In Number,
    挂号单_In 病人医嘱记录.挂号单%Type,
    病人id_In 病人医嘱记录.病人id%Type,
    主页id_In 病人医嘱记录.主页id%Type,
    发送号_In 病人医嘱发送.发送号%Type
  ) As
  Begin
    If Nvl(Rgstid_In, 0) = 0 Then
      Open c_In(病人id_In, 主页id_In, 发送号_In);
      Fetch c_In Bulk Collect
        Into r_Odr;
      Close c_In;
      --皮试结果
      Open Ctestin(病人id_In, 主页id_In);
      Fetch Ctestin Bulk Collect
        Into Rs_Test;
      Close Ctestin;
    Else
      Open c_Out(挂号单_In, 发送号_In);
      Fetch c_Out Bulk Collect
        Into r_Odr;
      Close c_Out;
      --皮试结果
      Open Ctest(挂号单_In);
      Fetch Ctest Bulk Collect
        Into Rs_Test;
      Close Ctest;
      --诊断
      Open c_Diag(病人id_In, Rgstid_In);
      Fetch c_Diag Bulk Collect
        Into Rs_Diag;
      Close c_Diag;
    End If;
  End;

  Procedure p_Pivasbill_Get
  (
    病人id_In  Number,
    主页id_In  Number,
    发送号_In  Number,
    医嘱ids_In Varchar2,
    j_Out      Out Clob
  ) As
    -----------------------------------------------------------
    --功能:从HIS库中获取可以发送到静配的医嘱信息
    -- 入参
    -- input
    --   pati_id                   N 1 病人id
    --   pati_pageid               N 1 主页id
    --   send_num                  N 1 发送号
    --   order_ids                 C 1 病人主医嘱id 拼串
    -- 出参
    --    pati_id                  N  1  病人id
    --    page_id                  N  1  主页ID
    --    pati_name                C  1  姓名
    --    pati_sex                 C  1  性别
    --    pati_age                 C  1  年龄
    --    inpatient_num            N  1  住院号
    --    pati_bed                 C  1  床号
    --    pati_wardarea_id         N  1  病人病区id
    --    pati_deptid              N  1  病人科室id
    --    advice_list[]            主医嘱列表
    --      advice_id              N  1  医嘱id --主医嘱ID(给药途径)
    --      advice_send_no         N  1  发送号 351,  --发送号
    --      effective_time         N  1  医嘱期效
    --      drug_method_id         N  1  给药途径id
    --      is_tpn                 N  1  是否tpn
    --      advice_frequency       C  1  执行频次，一天两次
    --      acvice_drug_list[]     药嘱信息
    --         advice_id           N  1  嘱id
    --         advice_rcpno        C  1  药嘱发送产生的费用no
    --      advice_exetime_list[]  医嘱执行时间，3天内,本次的发送信息+历史的发送信息, 3天内
    --         advice_id           N  1  给药途径医嘱
    --         advice_send_no      N  1  发送号
    --         advice_require_time C  1  2019-07-02 16:30:00    --要求时间
    -----------------------------------------------------------
    v_Exetimes   Varchar2(32767);
    n_Pre医嘱id  Number(18) := 0;
    n_静配科室id Number(18);
  
    Vj_Pati      Varchar2(32767);
    Vj_Last      Varchar2(32767);
    Vj_Jsonlist1 Varchar2(32767);
    Cj_Ad        Clob;
    Vj_Ad        Varchar2(32767);
  
    Cursor c_Ad Is
      Select Nvl(a.相关id, a.Id) As 主医嘱id, b.医嘱id, a.病人id, a.主页id, c.姓名, c.性别, c.年龄, c.住院号, c.出院病床 As 床号, a.病人科室id,
             c.当前病区id As 病人病区id, Nvl(b.执行部门id, 0) 执行部门id, e.Id As 给药途径id, d.执行频次, Decode(e.执行标记, 2, 1, 0) As Tpn, d.医嘱期效,
             b.发送号, b.No, b.发送人, To_Char(b.发送时间, 'YYYY-MM-DD HH24:MI:SS') As 发送时间
      From 病人医嘱记录 A, 病人医嘱发送 B, 病案主页 C, 病人医嘱记录 D, 诊疗项目目录 E
      Where a.Id = b.医嘱id And a.病人id = c.病人id And a.主页id = c.主页id And a.相关id = d.Id And d.诊疗项目id = e.Id And
            a.病人id = 病人id_In And a.主页id = 主页id_In And b.发送号 = 发送号_In And a.诊疗类别 In ('5', '6') And
            Instr(',' || 医嘱ids_In || ',', ',' || Nvl(a.相关id, a.Id) || ',') > 0
      Order By Nvl(a.相关id, a.Id), a.序号;
  
    Cursor c_Ext(医嘱id_In 医嘱执行时间.医嘱id%Type) Is
      Select /*+cardinality(j,10)*/
       a.发送号, To_Char(a.要求时间, 'YYYY-MM-DD HH24:MI:SS') As 要求时间
      From 医嘱执行时间 A
      Where a.要求时间 Between Sysdate - 3 And Sysdate + 3 And a.医嘱id = 医嘱id_In;
  Begin
  
    Vj_Ad := Null;
    For R In c_Ad Loop
      If r.执行部门id <> 0 Then
        n_静配科室id := r.执行部门id;
      End If;
    
      If n_Pre医嘱id = 0 Then
        Vj_Pati := Vj_Pati || '{"pati_id":' || r.病人id;
        Vj_Pati := Vj_Pati || ',"page_id":' || r.主页id;
        Vj_Pati := Vj_Pati || ',"pati_name":"' || Zljsonstr(r.姓名) || '"';
        Vj_Pati := Vj_Pati || ',"pati_sex":"' || Zljsonstr(r.性别) || '"';
        Vj_Pati := Vj_Pati || ',"pati_age":"' || Zljsonstr(r.年龄) || '"';
        Vj_Pati := Vj_Pati || ',"inpatient_num":"' || r.住院号 || '"';
        Vj_Pati := Vj_Pati || ',"pati_bed":"' || Zljsonstr(r.床号) || '"';
        Vj_Pati := Vj_Pati || ',"pati_wardarea_id":' || Nvl(r.病人病区id || '', 'null');
        Vj_Pati := Vj_Pati || ',"pati_deptid":' || Nvl(r.病人科室id || '', 'null');
      End If;
    
      If n_Pre医嘱id <> 0 And n_Pre医嘱id <> r.主医嘱id Then
        Vj_Ad := Vj_Ad || Vj_Last || ',"advice_drug_list":[' || Substr(Vj_Jsonlist1, 2) || ']';
      
        --本次的发送时间和历次的发送的执行时间点信息
        v_Exetimes := Null;
        For r_Ext In c_Ext(n_Pre医嘱id) Loop
          --某条医嘱的执行时间点
          v_Exetimes := v_Exetimes || ',{"advice_send_no":' || r_Ext.发送号;
          v_Exetimes := v_Exetimes || ',"advice_require_time":"' || r_Ext.要求时间 || '"';
          v_Exetimes := v_Exetimes || '}';
        End Loop;
      
        Vj_Ad := Vj_Ad || ',"advice_exetime_list":[' || Substr(v_Exetimes, 2) || ']';
        Vj_Ad := Vj_Ad || '}';
      
        If Length(Vj_Ad) > 20000 Then
          Cj_Ad := Cj_Ad || Vj_Ad;
          Vj_Ad := Null;
        End If;
        Vj_Jsonlist1 := Null;
      End If;
    
      n_Pre医嘱id := r.主医嘱id;
    
      --药品行医嘱信息
      Vj_Jsonlist1 := Vj_Jsonlist1 || ',{"advice_id":' || r.医嘱id;
      Vj_Jsonlist1 := Vj_Jsonlist1 || ',"advice_rcpno":"' || r.No || '"';
      Vj_Jsonlist1 := Vj_Jsonlist1 || '}';
    
      --用于最后一次的json拼接
      Vj_Last := ',{"pivas_deptid":' || n_静配科室id;
      Vj_Last := Vj_Last || ',"advice_id":' || r.主医嘱id;
      Vj_Last := Vj_Last || ',"advice_send_no":' || r.发送号;
      Vj_Last := Vj_Last || ',"effective_time":' || r.医嘱期效;
      Vj_Last := Vj_Last || ',"drug_method_id":' || r.给药途径id;
      Vj_Last := Vj_Last || ',"is_tpn":' || Nvl(r.Tpn || '', 'null');
      Vj_Last := Vj_Last || ',"advice_frequency":"' || Zljsonstr(r.执行频次) || '"';
    End Loop;
  
    If n_Pre医嘱id <> 0 Then
      Vj_Ad := Vj_Ad || Vj_Last || ',"advice_drug_list":[' || Substr(Vj_Jsonlist1, 2) || ']';
    
      --本次的发送时间和历次的发送的执行时间点信息
      v_Exetimes := Null;
      For r_Ext In c_Ext(n_Pre医嘱id) Loop
        --某条医嘱的执行时间点
        v_Exetimes := v_Exetimes || ',{"advice_send_no":' || r_Ext.发送号;
        v_Exetimes := v_Exetimes || ',"advice_require_time":"' || r_Ext.要求时间 || '"';
        v_Exetimes := v_Exetimes || '}';
      End Loop;
    
      Vj_Ad := Vj_Ad || ',"advice_exetime_list":[' || Substr(v_Exetimes, 2) || ']';
      Vj_Ad := Vj_Ad || '}';
    End If;
  
    Cj_Ad := Cj_Ad || Vj_Ad;
    If Cj_Ad Is Not Null Then
      j_Out := Vj_Pati || ',"advice_list":[' || Substr(Cj_Ad, 2) || ']}';
    End If;
  End;
Begin
  --解析入参
  If Json_In Is Null Then
    Select f_List2str(Cast(Collect(a.病人id || '') As t_Strlist), ',')
    Into v_Vals
    From (Select a.病人id From 病人医嘱异常记录 A Where a.产生环节 = 1 Group By a.病人id) A;
  Else
    j_Tmp  := Pljson(Json_In);
    j_Json := j_Tmp.Get_Pljson('input');
    v_Vals := j_Json.Get_Clob('pati_ids');
    If v_Vals Is Null Then
      Select f_List2str(Cast(Collect(a.病人id || '') As t_Strlist), ',')
      Into v_Vals
      From (Select a.病人id From 病人医嘱异常记录 A Where a.产生环节 = 1 Group By a.病人id) A;
    End If;
  End If;

  l_Vals := t_Strlist();
  While v_Vals Is Not Null Loop
    If Length(v_Vals) <= 4000 Then
      l_Vals.Extend;
      l_Vals(l_Vals.Count) := v_Vals;
      v_Vals := Null;
    Else
      l_Vals.Extend;
      l_Vals(l_Vals.Count) := Substr(v_Vals, 1, Instr(v_Vals, ',', 3980) - 1);
      v_Vals := Substr(v_Vals, Instr(v_Vals, ',', 3980) + 1);
    End If;
  End Loop;

  n_行号 := 0;
  For Lp In 1 .. l_Vals.Count Loop
    For Cp In (Select /*+Cardinality(j,10)*/
                a.病人id, a.主页id, a.挂号单, b.发送号
               From 病人医嘱记录 A, 病人医嘱异常记录 B, Table(f_Num2list(l_Vals(Lp))) J
               Where a.Id = b.医嘱id And a.病人id = j.Column_Value And b.产生环节 = 1
               Group By a.病人id, a.主页id, a.挂号单, b.发送号) Loop
    
      n_Rgst_Id := 0;
      If Cp.挂号单 Is Not Null Then
        Select a.Id Into n_Rgst_Id From 病人挂号记录 A Where a.No = Cp.挂号单 And a.记录性质 = 1 And a.记录状态 = 1;
      End If;
    
      v_Jtmp := Null;
      v_Jtmp := v_Jtmp || ',{"pati_id":' || Cp.病人id;
      v_Jtmp := v_Jtmp || ',"pati_pageid":' || Nvl(Cp.主页id || '', 'null');
      v_Jtmp := v_Jtmp || ',"rgst_id":' || Nvl(n_Rgst_Id || '', 'null');
      v_Jtmp := v_Jtmp || ',"rgst_no":"' || Cp.挂号单 || '"';
      v_Jtmp := v_Jtmp || ',"send_no":' || Cp.发送号;
    
      If v_操作员 Is Null Then
        Select Max(a.发送人), To_Char(Max(a.发送时间), 'yyyy-mm-dd hh24:mi:ss')
        Into v_操作员, v_操作时间
        From 病人医嘱发送 A
        Where a.发送号 = Cp.发送号;
      End If;
      v_Jtmp := v_Jtmp || ',"operator_name":"' || Zljsonstr(v_操作员) || '"';
      v_Jtmp := v_Jtmp || ',"operator_time":"' || v_操作时间 || '"';
    
      --病人类型
      If Cp.主页id Is Not Null Then
        Select a.病人类型 Into v_Tmp From 病案主页 A Where a.病人id = Cp.病人id And a.主页id = Cp.主页id;
        v_Jtmp := v_Jtmp || ',"pati_type":"' || v_Tmp || '"';
      End If;
    
      --S获取病人诊断信息 
      v_Tmp := Null;
      For K In (Select ID, 诊断描述, 诊断类型
                From 病人诊断记录
                Where 病人id = Cp.病人id And 主页id = Nvl(Cp.主页id, n_Rgst_Id) And 诊断描述 Is Not Null
                Order By 诊断类型, 记录来源, 记录日期 Desc, 诊断次序 Asc, Nvl(录入次序, '01'), Nvl(编码序号, 1)) Loop
        v_Tmp := v_Tmp || ',{"diag_rec_id":' || k.Id; -- : N 诊断记录id
        v_Tmp := v_Tmp || ',"diag_type":' || k.诊断类型; -- ：N 诊断类型 
        v_Tmp := v_Tmp || ',"diag_name":"' || Zljsonstr(k.诊断描述) || '"'; -- ：C 诊断名称（诊断描述）
        v_Tmp := v_Tmp || '}';
      End Loop;
      If v_Tmp Is Not Null Then
        v_Jtmp := v_Jtmp || ',"diag_list":[' || Substr(v_Tmp, 2) || ']';
      End If;
      --E获取病人诊断信息 
    
      n_行号 := n_行号 + 1;
      If n_行号 = 1 Then
        c_Outtmp := Substr(v_Jtmp, 2);
      Else
        c_Outtmp := c_Outtmp || v_Jtmp;
      End If;
    
      Cjl_Iitem := Null;
      Vj_Iitem  := Null;
      --获取医嘱明细相关信息缓存到 r_Odr 中
      p_Getbaseinfo(n_Rgst_Id, Cp.挂号单, Cp.病人id, Cp.主页id, Cp.发送号);
      For Ol In 1 .. r_Odr.Count Loop
        n_Send_No                := r_Odr(Ol).发送号;
        n_Advice_Id              := r_Odr(Ol).Id;
        n_Drug_Method_Id         := r_Odr(Ol).给药途径id;
        v_Drug_Method_Name       := r_Odr(Ol).给药途径名称;
        v_Drug_Method_Class_Code := r_Odr(Ol).诊疗频率编码;
        n_Drug_Freq_Id           := r_Odr(Ol).诊疗频率编码;
        v_Drug_Freq_Name         := r_Odr(Ol).给药频次名称;
        n_Emergency_Tag          := r_Odr(Ol).紧急标志;
        n_Denominated            := r_Odr(Ol).计价特性;
        n_Frequency              := r_Odr(Ol).频率次数;
        n_Single                 := r_Odr(Ol).单次用量;
        v_Usage                  := r_Odr(Ol).用法;
        v_Rcpdtl_Drask           := r_Odr(Ol).医生嘱托;
        n_Effectivetime          := r_Odr(Ol).医嘱期效;
        n_Cadn_Id                := r_Odr(Ol).诊疗项目id;
        n_组id                   := Nvl(r_Odr(Ol).相关id, r_Odr(Ol).Id);
      
        n_Group_Sno := Null;
        n_Cnt       := 0;
        For Ir In c_Adv(n_组id) Loop
          n_Cnt := n_Cnt + 1;
          If Ir.Id = r_Odr(Ol).Id Then
            n_Group_Sno := n_Cnt;
            Exit;
          End If;
        End Loop;
      
        v_Take_No   := r_Odr(Ol).领药号;
        v_Diag_Name := Null;
        --门诊病人特有
        If Cp.挂号单 Is Not Null Then
          Pp_Diag(n_组id, v_Diag_Name);
        End If;
      
        If r_Odr(Ol).诊疗类别 In ('5', '6') Then
          Pp_Test(n_Cadn_Id, v_Rcpdtl_St_Result);
        Else
          v_Rcpdtl_St_Result := Null;
        End If;
      
        --n_use_mode--
        n_Use_Mode := 0;
        If r_Odr(Ol).药行执行性质 = 4 And r_Odr(Ol).给药执行性质 = 5 Then
          n_Use_Mode := 1;
        Elsif r_Odr(Ol).执行标记 = 1 And r_Odr(Ol).药行执行性质 = 4 And r_Odr(Ol).给药执行性质 = 2 Then
          n_Use_Mode := 2;
        End If;
      
        If r_Odr(Ol).记录性质 = 1 Or r_Odr(Ol).记录性质 = 2 And r_Odr(Ol).门诊记帐 = 1 Then
          n_费用来源 := 1;
        Else
          n_费用来源 := 2;
        End If;
      
        Vj_Iitem := Vj_Iitem || ',{"advice_id":' || n_Advice_Id;
        Vj_Iitem := Vj_Iitem || ',"group_sno":' || Nvl(n_Group_Sno || '', 'null');
        Vj_Iitem := Vj_Iitem || ',"effectivetime":' || r_Odr(Ol).医嘱期效;
        Vj_Iitem := Vj_Iitem || ',"drug_method_id":' || Nvl(n_Drug_Method_Id || '', 'null');
        Vj_Iitem := Vj_Iitem || ',"drug_method_name":"' || Zljsonstr(v_Drug_Method_Name) || '"';
        Vj_Iitem := Vj_Iitem || ',"drug_method_class_code":"' || Zljsonstr(v_Drug_Method_Class_Code) || '"';
        Vj_Iitem := Vj_Iitem || ',"drug_freq_id":' || Nvl(n_Drug_Freq_Id || '', 'null');
        Vj_Iitem := Vj_Iitem || ',"drug_freq_name":"' || Zljsonstr(v_Drug_Freq_Name) || '"';
        Vj_Iitem := Vj_Iitem || ',"emergency_tag":' || Nvl(n_Emergency_Tag || '', 'null');
        Vj_Iitem := Vj_Iitem || ',"fee_mode":0';
        Vj_Iitem := Vj_Iitem || ',"use_mode":' || Zljsonstr(n_Use_Mode, 1);
        Vj_Iitem := Vj_Iitem || ',"frequency":' || Zljsonstr(n_Frequency, 1);
        Vj_Iitem := Vj_Iitem || ',"single":' || Zljsonstr(n_Single, 1);
        Vj_Iitem := Vj_Iitem || ',"usage":"' || Zljsonstr(v_Usage) || '"';
        Vj_Iitem := Vj_Iitem || ',"rcpdtl_st_result":"' || Zljsonstr(v_Rcpdtl_St_Result) || '"';
        Vj_Iitem := Vj_Iitem || ',"rcpdtL_excs_desc":"' || Zljsonstr(r_Odr(Ol).超量说明) || '"';
        Vj_Iitem := Vj_Iitem || ',"rcpdtL_drask":"' || Zljsonstr(v_Drug_Method_Name) || '"';
        Vj_Iitem := Vj_Iitem || ',"memo":"医嘱发送"';
        Vj_Iitem := Vj_Iitem || ',"diag_name":"' || Zljsonstr(v_Diag_Name) || '"';
        Vj_Iitem := Vj_Iitem || ',"take_no":"' || Zljsonstr(r_Odr(Ol).领药号) || '"';
        Vj_Iitem := Vj_Iitem || ',"advice_purpose":"' || Zljsonstr(r_Odr(Ol).用药目的) || '"';
        Vj_Iitem := Vj_Iitem || ',"fee_source":' || n_费用来源;
        Vj_Iitem := Vj_Iitem || ',"fee_billtype":' || r_Odr(Ol).记录性质;
        Vj_Iitem := Vj_Iitem || ',"fee_no":"' || r_Odr(Ol).No || '"';
        Vj_Iitem := Vj_Iitem || '}';
      
        If Length(Vj_Iitem) > 30000 Then
          If Cjl_Iitem Is Null Then
            Cjl_Iitem := Substr(Vj_Iitem, 2);
          Else
            Cjl_Iitem := Cjl_Iitem || Vj_Iitem;
          End If;
          Vj_Iitem := Null;
        End If;
      
        --a静配相关-----------------------------
        If r_Odr(Ol).给药操作类型 = '2' And r_Odr(Ol).给药类别 = 'E' And r_Odr(Ol).给药途径分类 = '1' Then
          Select Count(1)
          Into n_配液标记
          From 病人医嘱异常记录 A
          Where a.医嘱id = n_组id And a.发送号 = r_Odr(Ol).发送号 And a.产生环节 = 3;
          If n_配液标记 > 0 Then
            --收集静配数据同步异常的主医嘱
            If Instr(',' || v_P医嘱ids || ',', ',' || n_组id || ',') = 0 Then
              v_P医嘱ids := v_P医嘱ids || ',' || n_组id;
            End If;
          End If;
        End If;
        --e静配相关-----------------------------
      End Loop;
    
      If Cjl_Iitem Is Null Then
        c_Outtmp := c_Outtmp || ',"order_list":[' || Substr(Vj_Iitem, 2) || ']';
      Else
        c_Outtmp := c_Outtmp || ',"order_list":[' || Cjl_Iitem || Vj_Iitem || ']';
      End If;
    
      --a静配相关-----------------------------
      If v_P医嘱ids Is Not Null Then
        p_Pivasbill_Get(Cp.病人id, Cp.主页id, Cp.发送号, Substr(v_P医嘱ids, 2), b_Pivasout);
        If b_Pivasout Is Not Null Then
          c_Outtmp := c_Outtmp || ',"pivas_list":[' || b_Pivasout || ']';
        End If;
      End If;
      --e静配相关-----------------------------
    
      c_Outtmp := c_Outtmp || '}';
    End Loop;
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"成功","pati_bill_list":[' || c_Outtmp || ']}}';
Exception
  When Err_Item Then
    Json_Out := Zljsonout(v_Err);
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Cissvr_Getdrugerrdata;
/
CREATE OR REPLACE Procedure Zl_Cissvr_Geterrbillinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) Is
  -------------------------------------------------------------------------------------------------
  --功能：获取医生站预约接收异常数据
  --入参：json格式
  --Input
  --   rgst_no               C  1 挂号单
  --出参：json格式
  --Json_Out
  --   code                  C  1  应答码：0-失败；1-成功
  --   message               C  1  应答消息：失败时返回具体的错误信息
  --   snyc_status           N  1  同步状态：-1-病人ID为同步到业务数据;1-三方支付未校验;2-校验完成未支付成功
  --   pati_id               N  1  病人ID
  --   outpno                C  1  门诊号
  --   rgst_balance          C  1  交易信息:同步状态为2时返回三方交易信息
  -------------------------------------------------------------------------------------------------

  v_挂号单   病人结算异常记录.预交单号%Type;
  n_病人id   病人结算异常记录.病人id%Type;
  n_门诊号   病人结算异常记录.门诊号%Type;
  n_同步状态 病人结算异常记录.同步状态%Type;
  v_交易信息 病人结算异常记录.交易信息%Type;
  j_Json     Pljson;
  j_Jsontmp  Pljson;
  v_Temp     Varchar2(100);
Begin
  --解析入参
  j_Jsontmp := Pljson(Json_In);
  j_Json    := j_Jsontmp.Get_Pljson('input');
  v_挂号单  := j_Json.Get_String('rgst_no');

  Begin
    Select 病人id, 门诊号, 同步状态, 交易信息
    Into n_病人id, n_门诊号, n_同步状态, v_交易信息
    From 病人结算异常记录
    Where 预交单号 = v_挂号单 And 操作场景 = 4 And Rownum < 2;
  Exception
    When Others Then
      Json_Out := '{"output":{"code":1,"message":"成功","snyc_status":0}}';
      Return;
  End;
  --去掉'input'
  If n_同步状态 = 2 Then
    j_Jsontmp := Pljson();
    j_Json    := Pljson(v_交易信息);
    j_Jsontmp := j_Json.Get_Pljson('input');

    v_交易信息 := Empty_Clob();
    Dbms_Lob.Createtemporary(v_交易信息, True);
    j_Jsontmp.To_Clob(v_交易信息);
  Else
    v_交易信息 := '""';
  End If;

  v_Temp := '{"output":{"code":1,"message":"成功"';
  v_Temp := v_Temp || ',"pati_id":' || Nvl(n_病人id, 0);
  v_Temp := v_Temp || ',"outpno":"' || n_门诊号 || '"';
  v_Temp := v_Temp || ',"snyc_status":' || Nvl(n_同步状态, 0);
  v_Temp := v_Temp || ',"rgst_balance":';

  Json_Out := v_Temp || v_交易信息 || '}}';
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Cissvr_Geterrbillinfo;
/
Create Or Replace Procedure Zl_Cissvr_Getexecadvicerecord
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：根据指定执行的医嘱ID，返回执行医嘱信息记录集
  --入参：Json_In:格式
  --  input
  --    advice_send_ids             C 1 医嘱ID和发送号字符串，医嘱ID1:发送号1,医嘱ID2:发送号2
  --出参: Json_Out,格式如下
  --  output
  --    code                        N 1 应答吗：0-失败；1-成功
  --    message                     C 1 应答消息：失败时返回具体的错误信息
  --    item_list
  --         advice_id              N 1 医嘱id
  --         advice_related_id      N 1 相关id
  --         send_no                N 1 发送号
  --         pati_id                N 1 病人id
  --         pati_ageid             N 1 主页id
  --         advice_begintime       C 1 开始执行时间
  --         advice_note            C 1 医嘱内容
  --         nums                   C 1 数次
  --         advice_doctor_note     C 1 医生嘱托
  --         advice_doctor          C 1 开嘱医生
  --         advice_record_time     C 1 开嘱时间
  ---------------------------------------------------------------------------
  j_Tmp            Pljson;
  j_In             Pljson;
  v_List           Varchar2(32767);
  v_医嘱内容       Varchar2(4000);
  v_Order_Send_Ids Varchar2(32767);

  Cursor c_Ad Is
    Select b.Id As 医嘱id, b.相关id, a.发送号, b.病人id, b.主页id, To_Char(b.开始执行时间, 'YYYY-MM-DD HH24:MI:SS') As 开始时间,
           b.医嘱内容 As 医嘱内容, a.发送数次 || Nvl(d.计算单位, c.计算单位) As 数次, b.医生嘱托, b.开嘱医生,
           To_Char(b.开嘱时间, 'YYYY-MM-DD HH24:MI:SS') As 开嘱时间, a.发送时间, a.发送人, Nvl(d.计算单位, c.计算单位) As 计算单位, a.No As 单据号,
           a.记录性质, a.执行部门id, a.完成人, a.完成时间, b.序号, b.病人来源, b.挂号单, b.婴儿, b.姓名, b.医嘱期效, b.诊疗类别, b.诊疗项目id, b.标本部位, b.检查方法,
           b.天数, b.单次用量, b.总给予量, b.执行频次, b.紧急标志, b.收费细目id, b.首次用量, b.执行时间方案, b.皮试结果, b.医嘱内容 As 医嘱内容tmp, 0 As 数据转出
    From 病人医嘱发送 A, 病人医嘱记录 B, 诊疗项目目录 C, 收费项目目录 D
    Where a.医嘱id = b.Id And b.诊疗项目id = c.Id And b.收费细目id = d.Id(+) And
          (a.医嘱id, a.发送号) In (Select /*+cardinality(x,10)*/
                               x.C1 As 医嘱id, x.C2 As 发送号
                              From Table(f_Num2list2(v_Order_Send_Ids)) X)
    Order By b.序号;

  --获取医嘱内容
  Function Get执行内容
  (
    P医嘱id Number,
    P相关id Number,
    P类别   Varchar2,
    P内容   Varchar2
  ) Return Varchar2 Is
  
    v_内容      Varchar2(32767);
    v_Tmp       Varchar2(32767);
    Str皮试结果 Varchar2(32767);
    P行数       Number := 0;
    P记录数     Number := 0;
    Bln给药途径 Boolean := False;
    Cursor c_Pad Is
      Select a.Id, a.相关id, a.诊疗类别, a.医嘱内容, a.皮试结果, a.单次用量, b.计算单位, b.操作类型, a.执行频次, a.执行时间方案, b.名称
      From 病人医嘱记录 A, 诊疗项目目录 B
      Where Not (a.诊疗类别 = 'E' And 相关id Is Not Null) And a.诊疗项目id = b.Id And (a.相关id = P医嘱id Or a.Id = P医嘱id)
      Order By a.序号;
    Type t_Pad Is Table Of c_Pad%RowType;
    Rstmp t_Pad;
  
  Begin
  
    If (P类别 = 'C' And Nvl(P相关id, 0) <> 0) Or P类别 = 'D' Then
      v_内容 := P内容;
    Elsif P类别 <> 'E' Or Nvl(P相关id, 0) <> 0 Then
      v_内容 := P内容;
      If P类别 = 'E' Then
        Select a.医嘱内容 Into v_内容 From 病人医嘱记录 A Where a.Id = P相关id;
      End If;
    Else
      --类别为E,且相关ID=0
      Open c_Pad;
      Fetch c_Pad Bulk Collect
        Into Rstmp;
      Close c_Pad;
      P记录数 := Rstmp.Count;
      For I In 1 .. P记录数 Loop
        If Nvl(Rstmp(I).相关id, 0) = P医嘱id Then
          P行数 := P行数 + 1;
          If Rstmp(I).诊疗类别 In ('5', '6') Then
            Bln给药途径 := True;
          End If;
        End If;
      End Loop;
      If Not Bln给药途径 Then
        v_内容 := P内容;
        If Rstmp(1).诊疗类别 = 'E' And Rstmp(1).操作类型 = '1' Then
          Str皮试结果 := '，皮试结果：' || Rstmp(1).皮试结果;
          For R In (Select b.过敏反应, b.过敏时间
                    From 病人医嘱记录 A, 病人过敏记录 B, 诊疗项目目录 C, 诊疗用法用量 D
                    Where a.病人id = b.病人id And a.诊疗项目id = d.用法id And d.项目id = c.Id And c.类别 In ('5', '6') And
                          d.项目id = b.药物id And Nvl(d.性质, 0) = 0 And
                          b.记录时间 = (Select Max(操作时间) From 病人医嘱状态 Where 医嘱id = a.Id And 操作类型 = 10) And a.Id = P医嘱id And
                          Rownum < 2) Loop
          
            Str皮试结果 := Str皮试结果 || ',过敏时间：' || To_Char(r.过敏时间, 'yyyy-mm-dd');
            If r.过敏反应 Is Not Null Then
              Str皮试结果 := Str皮试结果 || ',过敏反应：' || r.过敏反应;
            End If;
          End Loop;
        End If;
      Else
        --给药途径
        v_内容 := Null;
        For I In 1 .. P行数 Loop
          If I = P行数 Then
            v_内容 := v_内容 || Chr(13) || '┗';
          Else
            v_内容 := v_内容 || Chr(13) || '┣';
          End If;
          v_内容 := v_内容 || Rstmp(I).医嘱内容;
          If Rstmp(I).单次用量 Is Not Null Then
            v_内容 := v_内容 || ' ' || Round(Rstmp(I).单次用量, 5) || Rstmp(I).计算单位;
          End If;
        End Loop;
        v_Tmp := Rstmp(P记录数).名称 || ',' || Rstmp(P记录数).执行频次;
        If Rstmp(P记录数).执行时间方案 Is Not Null Then
          v_Tmp := v_Tmp || '(' || Rstmp(P记录数).执行时间方案 || ')';
        End If;
        v_Tmp  := v_Tmp || ':每' || Rstmp(P记录数).计算单位;
        v_内容 := v_Tmp || ' ' || v_内容;
      End If;
    End If;
    Return v_内容 || Str皮试结果;
  End Get执行内容;

Begin
  --解析入参
  j_In             := Pljson(Json_In);
  j_Tmp            := j_In.Get_Pljson('input');
  v_Order_Send_Ids := j_Tmp.Get_String('advice_send_ids');

  Json_Out := '{"output":{"code":1,"message":"成功","item_list":[';

  For R In c_Ad Loop
  
    v_医嘱内容 := Get执行内容(r.医嘱id, r.相关id, r.诊疗类别, r.医嘱内容);
  
    v_List := v_List || ',{';
    v_List := v_List || '"advice_id":' || r.医嘱id;
    v_List := v_List || ',"advice_related_id":' || Nvl(r.相关id || '', 'null');
    v_List := v_List || ',"send_no":' || r.发送号;
    v_List := v_List || ',"pati_id":' || r.病人id;
    v_List := v_List || ',"pati_ageid":' || Nvl(r.主页id || '', 'null');
    v_List := v_List || ',"advice_begintime":"' || r.开始时间 || '"';
    v_List := v_List || ',"advice_note":"' || Zljsonstr(v_医嘱内容) || '"';
    v_List := v_List || ',"nums":"' || Zljsonstr(r.数次) || '"';
    v_List := v_List || ',"advice_doctor_note":"' || Zljsonstr(r.医生嘱托) || '"';
    v_List := v_List || ',"advice_doctor":"' || Zljsonstr(r.开嘱医生) || '"';
    v_List := v_List || ',"advice_record_time":"' || r.开嘱时间 || '"';
    v_List := v_List || '}';
  
  End Loop;

  Json_Out := Json_Out || Substr(v_List, 2) || ']}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getexecadvicerecord;
/
Create Or Replace Procedure Zl_Cissvr_Getgroupadviceinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:获取一并给药的医嘱内容
  --入参：Json_In:格式
  -- input
  --   advice_ids           C   1 医嘱ID，多个用英文的逗号分隔
  --出参: Json_Out,格式如下
  --  output
  --    code                C  1 应答码：0-失败；1-成功
  --    message             C  1 应答消息：失败时返回具体的错误信息
  --    advice_list[]              [数组]每个医嘱信息
  --      advice_id         N   医嘱id，传入医嘱ID
  --      advice_note       C   医嘱内容
  ---------------------------------------------------------------------------
  j_In        Pljson;
  j_Json      Pljson;
  c_医嘱ids   Clob;
  l_医嘱id    t_Strlist := t_Strlist();
  n_Firstitem Number(1);
  v_Temp      Varchar2(32767);
Begin
  --解析入参
  j_In      := Pljson(Json_In);
  j_Json    := j_In.Get_Pljson('input');
  c_医嘱ids := j_Json.Get_Clob('advice_ids');

  While c_医嘱ids Is Not Null Loop
    If Length(c_医嘱ids) <= 4000 Then
      l_医嘱id.Extend;
      l_医嘱id(l_医嘱id.Count) := c_医嘱ids;
      c_医嘱ids := Null;
    Else
      l_医嘱id.Extend;
      l_医嘱id(l_医嘱id.Count) := Substr(c_医嘱ids, 1, Instr(c_医嘱ids, ',', 3950) - 1);
      c_医嘱ids := Substr(c_医嘱ids, Instr(c_医嘱ids, ',', 3950) + 1);
    End If;
  End Loop;

  If l_医嘱id.Count = 0 Then
    Json_Out := '{"output":{"code":0,"message":"未传入医嘱id，请检查！"}}';
    Return;
  End If;

  Json_Out := '{"output":{"code":1,"message":"成功","advice_list":[';

  v_Temp      := '';
  n_Firstitem := 1;
  For I In 1 .. l_医嘱id.Count Loop
    For r_医嘱 In (Select /*+cardinality(j,10)*/
                  a.Id, b.医嘱内容
                 From 病人医嘱记录 A, 病人医嘱记录 B, Table(f_Num2list(l_医嘱id(I))) J
                 Where a.相关id = b.相关id And a.Id <> b.Id And a.Id = j.Column_Value) Loop
    
      If Nvl(n_Firstitem, 0) = 0 Then
        v_Temp := v_Temp || ',';
      Else
        n_Firstitem := 0;
      End If;
      v_Temp := v_Temp || '{';
      v_Temp := v_Temp || '"advice_id":' || Nvl(r_医嘱.Id, 0);
      v_Temp := v_Temp || ',"advice_note":"' || Zljsonstr(r_医嘱.医嘱内容) || '"';
      v_Temp := v_Temp || '}';
    
      If Length(v_Temp) > 20000 Then
        Json_Out := Json_Out || v_Temp;
        v_Temp   := '';
      End If;
    End Loop;
  End Loop;

  Json_Out := Json_Out || v_Temp || ']}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getgroupadviceinfo;
/
Create Or Replace Procedure Zl_Cissvr_Getinfectinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：查询病人的阳性结果反馈单信息
  --入参：Json_In:格式
  --input
  --    query_type         N    1  调用类型 ：1-通过病人id+主页id查询   2-通过病人id+挂号单查询  3-通过登记科室查询   4-通过疾病id串查询
  --    pati_id            N    1  病人id
  --    create_dept_id     N    1  登记科室ID
  --    pati_pageid        N    0  主页id
  --    reg_no             C    0  挂号单
  --    create_time_begin  C    0  登记开始时间
  --    create_time_end    C    0  登记开始时间
  --    rec_ids            C    0  疾病id串

  --出参：Json_Out:格式
  --output
  --    code                N   1 应答吗：0-失败；1-成功
  --    message             C   1 应答消息：失败时返回具体的错误信息
  --    disease_list         阳性结果反馈单列表，支持多个，[数组]
  --       rec_id                    N   1  疾病id
  --       pati_source              C   1  来源
  --       pati_id                  C   1  病人id
  --       pati_name                C   1  病人姓名
  --       pati_sex                 C   1  病人性别
  --       pati_age                 C   1  病人年龄
  --       pati_dept_name           C   1  病人科室
  --       inpatient_num            C   1  住院号
  --       outpatient_num           C   1  门诊号
  --       spcm_send_time           C   1  标本送检时间
  --       spcm_send_dr             C   1  送检医生
  --       spcm_send_dept           C   1  送检科室
  --       spcm_send_deptid         N   1  送检科室ID

  --       spcm_rec_status          N   1  记录状态
  --       create_dept_name         C   1  登记科室
  --       spcm_name                C   1  标本名称
  --       send_content             C   1  反馈内容
  --       infctdz_name             C   1  疑似疾病
  --       create_dr                C   1  登记医生
  --       create_time              C   1  登记时间
  --       spcm_procor              C   1  检验处理人
  --       spcm_proctime            C   1  检验处理时间
  --       spcm_procdesc            C   1  检验处理说明

  --       pati_pageid              N   1  主页id
  --       reg_no                   C   1  挂号单
  --       advice_id                N   1  医嘱ID
  --       eqpmtn_exetime           C   1  检查时间
  --       create_dept_id           N   1  登记科室ID
  --       clinic_type              C   1  诊疗类别
  --       advice_doctor            C   1  开嘱医生
  ---------------------------------------------------------------------------
  j_Json   Pljson;
  j_In     Pljson;
  n_Type   Number; --调用类型 ：1-通过id查询,2-通过挂号单查询,3-通过病人id加主页id查询
  v_挂号单 Varchar2(20);
  n_病人id Number(18);
  n_主页id Number(18);
  v_List   Varchar2(32765);

  d_登记开始时间 Date;
  d_登记结束时间 Date;

  n_登记科室id Number;
  v_疾病ids    Varchar2(4000);
Begin
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');
  n_Type := Nvl(j_Json.Get_Number('query_type'), 0);

  Json_Out := '{"output":{"code":1,"message":"成功","disease_list":[';

  If n_Type = 1 Then
    n_病人id     := Nvl(j_Json.Get_Number('pati_id'), 0);
    n_登记科室id := Nvl(j_Json.Get_Number('create_dept_id'), 0);
    n_主页id     := Nvl(j_Json.Get_Number('pati_pageid'), 0);
  
    For c_传染病 In (Select a.Id, '住院' As 来源, c.病人id, c.姓名, c.性别, c.年龄, e.名称 As 科室, c.住院号 As 标识号, a.送检时间, a.送检医生, a.送检科室id,
                         g.名称 As 送检科室, a.记录状态, f.名称 As 登记科室, a.标本名称, a.反馈结果, a.传染病名称 As 疑似疾病, a.登记人, a.登记时间, a.处理人,
                         a.处理时间, a.处理情况说明, a.主页id, a.挂号单, a.医嘱id, a.检查时间, a.登记科室id, m.诊疗类别, m.开嘱医生


                  
                  From 疾病阳性记录 A, 病人医嘱记录 M, 病案主页 C, 部门表 E, 部门表 F, 部门表 G
                  Where a.病人id = c.病人id And a.主页id = c.主页id And c.病人id = n_病人id And c.主页id = n_主页id And
                        a.登记科室id = f.Id(+) And c.出院科室id = e.Id(+) And a.送检科室id = g.Id(+) And a.医嘱id = m.Id(+) And
                        (n_登记科室id = 0 Or (a.登记科室id = n_登记科室id))
                  Order By a.登记时间 Desc) Loop
      Zljsonputvalue(v_List, 'rec_id', c_传染病.Id, 1, 1);
      Zljsonputvalue(v_List, 'pati_source', c_传染病.来源, 0);
      Zljsonputvalue(v_List, 'pati_id', c_传染病.病人id, 1);
      Zljsonputvalue(v_List, 'pati_name', c_传染病.姓名, 0);
      Zljsonputvalue(v_List, 'pati_sex', c_传染病.性别, 0);
      Zljsonputvalue(v_List, 'pati_age', c_传染病.年龄, 0);
      Zljsonputvalue(v_List, 'pati_dept_name', c_传染病.科室, 0);
      Zljsonputvalue(v_List, 'inpatient_num', c_传染病.标识号, 0);
      Zljsonputvalue(v_List, 'spcm_send_time', To_Char(c_传染病.送检时间, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'spcm_send_dr', c_传染病.送检医生, 0);
      Zljsonputvalue(v_List, 'spcm_rec_status', c_传染病.记录状态, 1);
      Zljsonputvalue(v_List, 'create_dept_name', c_传染病.登记科室, 0);
      Zljsonputvalue(v_List, 'spcm_name', c_传染病.标本名称, 0);
      Zljsonputvalue(v_List, 'send_content', c_传染病.反馈结果, 0);
      Zljsonputvalue(v_List, 'infctdz_name', c_传染病.疑似疾病, 0);
      Zljsonputvalue(v_List, 'create_dr', c_传染病.登记人, 0);
      Zljsonputvalue(v_List, 'create_time', To_Char(c_传染病.登记时间, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'spcm_procor', c_传染病.处理人, 0);
      Zljsonputvalue(v_List, 'spcm_proctime', To_Char(c_传染病.处理时间, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'spcm_procdesc', c_传染病.处理情况说明, 0);
    
      Zljsonputvalue(v_List, 'pati_pageid', c_传染病.主页id, 1);
      Zljsonputvalue(v_List, 'reg_no', c_传染病.挂号单, 0);
      Zljsonputvalue(v_List, 'advice_id', c_传染病.医嘱id, 1);
      Zljsonputvalue(v_List, 'eqpmtn_exetime', To_Char(c_传染病.检查时间, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'create_dept_id', c_传染病.登记科室id, 1);
      Zljsonputvalue(v_List, 'clinic_type', c_传染病.诊疗类别, 0);
      Zljsonputvalue(v_List, 'advice_doctor', c_传染病.开嘱医生, 0);
    
      Zljsonputvalue(v_List, 'spcm_send_dept', c_传染病.送检科室, 0);
      Zljsonputvalue(v_List, 'spcm_send_deptid', c_传染病.送检科室id, 1, 2);
      If Length(v_List) > 20000 Then
        Json_Out := Json_Out || v_List;
        v_List   := ',';
      End If;
    
    End Loop;
  Elsif n_Type = 2 Then
    n_病人id     := Nvl(j_Json.Get_Number('pati_id'), 0);
    n_登记科室id := Nvl(j_Json.Get_Number('create_dept_id'), 0);
  
    v_挂号单 := j_Json.Get_String('reg_no');
  
    For c_传染病 In (Select a.Id, '门诊' As 来源, b.病人id, b.姓名, b.性别, b.年龄, e.名称 As 科室, b.门诊号 As 标识号, a.送检时间, a.送检医生, a.送检科室id,
                         g.名称 As 送检科室, a.记录状态, f.名称 As 登记科室, a.标本名称, a.反馈结果, a.传染病名称 As 疑似疾病, a.登记人, a.登记时间, a.处理人,
                         a.处理时间, a.处理情况说明, a.主页id, a.挂号单, a.医嘱id, a.检查时间, a.登记科室id, m.诊疗类别, m.开嘱医生


                  
                  From 疾病阳性记录 A, 病人挂号记录 B, 部门表 E, 部门表 F, 病人医嘱记录 M, 部门表 G
                  Where a.病人id = b.病人id And a.挂号单 = b.No And b.病人id = n_病人id And b.No = v_挂号单 And a.登记科室id = f.Id(+) And
                        a.送检科室id = g.Id(+) And b.执行部门id = e.Id(+) And a.医嘱id = m.Id(+) And
                        (n_登记科室id = 0 Or (a.登记科室id = n_登记科室id))
                  Order By a.登记时间 Desc) Loop
      Zljsonputvalue(v_List, 'rec_id', c_传染病.Id, 1, 1);
      Zljsonputvalue(v_List, 'pati_source', c_传染病.来源, 0);
      Zljsonputvalue(v_List, 'pati_id', c_传染病.病人id, 1);
      Zljsonputvalue(v_List, 'pati_name', c_传染病.姓名, 0);
      Zljsonputvalue(v_List, 'pati_sex', c_传染病.性别, 0);
      Zljsonputvalue(v_List, 'pati_age', c_传染病.年龄, 0);
      Zljsonputvalue(v_List, 'pati_dept_name', c_传染病.科室, 0);
      Zljsonputvalue(v_List, 'outpatient_num', c_传染病.标识号, 0);
      Zljsonputvalue(v_List, 'spcm_send_time', To_Char(c_传染病.送检时间, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'spcm_send_dr', c_传染病.送检医生, 0);
      Zljsonputvalue(v_List, 'spcm_rec_status', c_传染病.记录状态, 1);
      Zljsonputvalue(v_List, 'create_dept_name', c_传染病.登记科室, 0);
      Zljsonputvalue(v_List, 'spcm_name', c_传染病.标本名称, 0);
      Zljsonputvalue(v_List, 'send_content', c_传染病.反馈结果, 0);
      Zljsonputvalue(v_List, 'infctdz_name', c_传染病.疑似疾病, 0);
      Zljsonputvalue(v_List, 'create_dr', c_传染病.登记人, 0);
      Zljsonputvalue(v_List, 'create_time', To_Char(c_传染病.登记时间, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'spcm_procor', c_传染病.处理人, 0);
      Zljsonputvalue(v_List, 'spcm_proctime', To_Char(c_传染病.处理时间, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'spcm_procdesc', c_传染病.处理情况说明, 0);
    
      Zljsonputvalue(v_List, 'pati_pageid', c_传染病.主页id, 1);
      Zljsonputvalue(v_List, 'reg_no', c_传染病.挂号单, 0);
      Zljsonputvalue(v_List, 'advice_id', c_传染病.医嘱id, 1);
      Zljsonputvalue(v_List, 'eqpmtn_exetime', To_Char(c_传染病.检查时间, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'create_dept_id', c_传染病.登记科室id, 1);
      Zljsonputvalue(v_List, 'clinic_type', c_传染病.诊疗类别, 0);
      Zljsonputvalue(v_List, 'advice_doctor', c_传染病.开嘱医生, 0);
    
      Zljsonputvalue(v_List, 'spcm_send_dept', c_传染病.送检科室, 0);
      Zljsonputvalue(v_List, 'spcm_send_deptid', c_传染病.送检科室id, 1, 2);
      If Length(v_List) > 20000 Then
        Json_Out := Json_Out || v_List;
        v_List   := ',';
      End If;
    End Loop;
  Elsif n_Type = 3 Then
    n_登记科室id   := Nvl(j_Json.Get_Number('create_dept_id'), 0);
    d_登记开始时间 := To_Date(j_Json.Get_String('create_time_begin'), 'yyyy-mm-dd hh24:mi:ss');
    d_登记结束时间 := To_Date(j_Json.Get_String('create_time_end'), 'yyyy-mm-dd hh24:mi:ss');
    For c_传染病 In (Select a.Id, a.来源, a.病人id, a.姓名, a.性别, a.年龄, e.名称 As 科室, a.标识号, a.送检时间, a.送检医生, a.送检科室id, a.记录状态,
                         f.名称 As 送检科室, a.标本名称, a.反馈结果, a.传染病名称 As 疑似疾病, a.登记人, a.登记时间, a.处理人, a.处理时间, a.处理情况说明, a.主页id,
                         a.挂号单, a.医嘱id, a.检查时间, a.登记科室id, g.名称 As 登记科室, a.诊疗类别, a.开嘱医生
                  From (Select a.Id, '门诊' As 来源, a.病人id, b.姓名, b.性别, b.年龄, b.门诊号 As 标识号, a.送检时间, a.送检医生, a.送检科室id, a.记录状态,
                                a.标本名称, a.反馈结果, a.传染病名称, a.登记人, a.登记时间, a.处理人, a.处理时间, a.处理情况说明, b.执行部门id As 科室id, a.主页id,
                                a.挂号单, a.医嘱id, a.检查时间, a.登记科室id, m.诊疗类别, m.开嘱医生
                         From 疾病阳性记录 A, 病人挂号记录 B, 病人医嘱记录 M
                         Where a.病人id = b.病人id And a.挂号单 = b.No And a.医嘱id = m.Id(+) And a.登记科室id = n_登记科室id And
                               a.登记时间 Between d_登记开始时间 And d_登记结束时间
                         Union All
                         Select a.Id, '住院' As 来源, a.病人id, c.姓名, c.性别, c.年龄, c.住院号 As 标识号, a.送检时间, a.送检医生, a.送检科室id, a.记录状态,
                                a.标本名称, a.反馈结果, a.传染病名称, a.登记人, a.登记时间, a.处理人, a.处理时间, a.处理情况说明, c.出院科室id As 科室id, a.主页id,
                                a.挂号单, a.医嘱id, a.检查时间, a.登记科室id, m.诊疗类别, m.开嘱医生
                         From 疾病阳性记录 A, 病案主页 C, 病人医嘱记录 M
                         Where a.病人id = c.病人id And a.主页id = c.主页id And a.医嘱id = m.Id(+) And a.登记科室id = n_登记科室id And
                               a.登记时间 Between d_登记开始时间 And d_登记结束时间) A, 部门表 E, 部门表 F, 部门表 G
                  Where a.送检科室id = f.Id(+) And a.科室id = e.Id(+) And a.登记科室id = g.Id(+)
                  Order By a.登记时间 Desc) Loop
      Zljsonputvalue(v_List, 'rec_id', c_传染病.Id, 1, 1);
      Zljsonputvalue(v_List, 'pati_source', c_传染病.来源, 0);
      Zljsonputvalue(v_List, 'pati_id', c_传染病.病人id, 1);
      Zljsonputvalue(v_List, 'pati_name', c_传染病.姓名, 0);
      Zljsonputvalue(v_List, 'pati_sex', c_传染病.性别, 0);
      Zljsonputvalue(v_List, 'pati_age', c_传染病.年龄, 0);
      Zljsonputvalue(v_List, 'pati_dept_name', c_传染病.科室, 0);
      If c_传染病.来源 = '门诊' Then
        Zljsonputvalue(v_List, 'outpatient_num', c_传染病.标识号, 0);
      Else
        Zljsonputvalue(v_List, 'inpatient_num', c_传染病.标识号, 0);
      End If;
      Zljsonputvalue(v_List, 'spcm_send_time', To_Char(c_传染病.送检时间, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'spcm_send_dr', c_传染病.送检医生, 0);
      Zljsonputvalue(v_List, 'spcm_rec_status', c_传染病.记录状态, 1);
      Zljsonputvalue(v_List, 'create_dept_name', c_传染病.登记科室, 0);
      Zljsonputvalue(v_List, 'spcm_name', c_传染病.标本名称, 0);
      Zljsonputvalue(v_List, 'send_content', c_传染病.反馈结果, 0);
      Zljsonputvalue(v_List, 'infctdz_name', c_传染病.疑似疾病, 0);
      Zljsonputvalue(v_List, 'create_dr', c_传染病.登记人, 0);
      Zljsonputvalue(v_List, 'create_time', To_Char(c_传染病.登记时间, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'spcm_procor', c_传染病.处理人, 0);
      Zljsonputvalue(v_List, 'spcm_proctime', To_Char(c_传染病.处理时间, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'spcm_procdesc', c_传染病.处理情况说明, 0);
    
      Zljsonputvalue(v_List, 'pati_pageid', c_传染病.主页id, 1);
      Zljsonputvalue(v_List, 'reg_no', c_传染病.挂号单, 0);
      Zljsonputvalue(v_List, 'advice_id', c_传染病.医嘱id, 1);
      Zljsonputvalue(v_List, 'eqpmtn_exetime', To_Char(c_传染病.检查时间, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'create_dept_id', c_传染病.登记科室id, 1);
      Zljsonputvalue(v_List, 'clinic_type', c_传染病.诊疗类别, 0);
      Zljsonputvalue(v_List, 'advice_doctor', c_传染病.开嘱医生, 0);
    
      Zljsonputvalue(v_List, 'spcm_send_dept', c_传染病.送检科室, 0);
      Zljsonputvalue(v_List, 'spcm_send_deptid', c_传染病.送检科室id, 1, 2);
    
      If Length(v_List) > 20000 Then
        Json_Out := Json_Out || v_List;
        v_List   := ',';
      End If;
    
    End Loop;
  
  Elsif n_Type = 4 Then
    v_疾病ids := j_Json.Get_String('rec_ids');
    For c_传染病 In (Select a.Id, a.来源, a.病人id, a.姓名, a.性别, a.年龄, e.名称 As 科室, a.标识号, a.送检时间, a.送检医生, a.送检科室id, a.记录状态,
                         f.名称 As 送检科室, a.标本名称, a.反馈结果, a.传染病名称 As 疑似疾病, a.登记人, a.登记时间, a.处理人, a.处理时间, a.处理情况说明, a.主页id,
                         a.挂号单, a.医嘱id, a.检查时间, a.登记科室id, g.名称 As 登记科室, a.诊疗类别, a.开嘱医生
                  From (Select a.Id, '门诊' As 来源, a.病人id, b.姓名, b.性别, b.年龄, b.门诊号 As 标识号, a.送检时间, a.送检医生, a.送检科室id, a.记录状态,
                                a.标本名称, a.反馈结果, a.传染病名称, a.登记人, a.登记时间, a.处理人, a.处理时间, a.处理情况说明, b.执行部门id As 科室id, a.主页id,
                                a.挂号单, a.医嘱id, a.检查时间, a.登记科室id, m.诊疗类别, m.开嘱医生
                         From 疾病阳性记录 A, 病人挂号记录 B, 病人医嘱记录 M
                         
                         Where a.病人id = b.病人id And a.挂号单 = b.No And a.医嘱id = m.Id(+) And
                               a.Id In
                               (Select Column_Value As 病人id From Table(Cast(f_Str2list(v_疾病ids) As Zltools.t_Strlist)))
                         Union All
                         Select a.Id, '住院' As 来源, a.病人id, c.姓名, c.性别, c.年龄, c.住院号 As 标识号, a.送检时间, a.送检医生, a.送检科室id, a.记录状态,
                                a.标本名称, a.反馈结果, a.传染病名称, a.登记人, a.登记时间, a.处理人, a.处理时间, a.处理情况说明, c.出院科室id As 科室id, a.主页id,
                                a.挂号单, a.医嘱id, a.检查时间, a.登记科室id, m.诊疗类别, m.开嘱医生
                         From 疾病阳性记录 A, 病案主页 C, 病人医嘱记录 M
                         
                         Where a.病人id = c.病人id And a.主页id = c.主页id And a.医嘱id = m.Id(+) And
                               a.Id In
                               (Select Column_Value As 病人id From Table(Cast(f_Str2list(v_疾病ids) As Zltools.t_Strlist)))) A,
                       部门表 E, 部门表 F, 部门表 G
                  Where a.送检科室id = f.Id(+) And a.科室id = e.Id(+) And a.登记科室id = g.Id(+)
                  Order By a.登记时间 Desc) Loop
    
      Zljsonputvalue(v_List, 'rec_id', c_传染病.Id, 1, 1);
      Zljsonputvalue(v_List, 'pati_source', c_传染病.来源, 0);
      Zljsonputvalue(v_List, 'pati_id', c_传染病.病人id, 1);
      Zljsonputvalue(v_List, 'pati_name', c_传染病.姓名, 0);
      Zljsonputvalue(v_List, 'pati_sex', c_传染病.性别, 0);
      Zljsonputvalue(v_List, 'pati_age', c_传染病.年龄, 0);
      Zljsonputvalue(v_List, 'pati_dept_name', c_传染病.科室, 0);
      If c_传染病.来源 = '门诊' Then
        Zljsonputvalue(v_List, 'outpatient_num', c_传染病.标识号, 0);
      Else
        Zljsonputvalue(v_List, 'inpatient_num', c_传染病.标识号, 0);
      End If;
      Zljsonputvalue(v_List, 'spcm_send_time', To_Char(c_传染病.送检时间, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'spcm_send_dr', c_传染病.送检医生, 0);
      Zljsonputvalue(v_List, 'spcm_rec_status', c_传染病.记录状态, 1);
      Zljsonputvalue(v_List, 'create_dept_name', c_传染病.登记科室, 0);
      Zljsonputvalue(v_List, 'spcm_name', c_传染病.标本名称, 0);
      Zljsonputvalue(v_List, 'send_content', c_传染病.反馈结果, 0);
      Zljsonputvalue(v_List, 'infctdz_name', c_传染病.疑似疾病, 0);
      Zljsonputvalue(v_List, 'create_dr', c_传染病.登记人, 0);
      Zljsonputvalue(v_List, 'create_time', To_Char(c_传染病.登记时间, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'spcm_procor', c_传染病.处理人, 0);
      Zljsonputvalue(v_List, 'spcm_proctime', To_Char(c_传染病.处理时间, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'spcm_procdesc', c_传染病.处理情况说明, 0);
    
      Zljsonputvalue(v_List, 'pati_pageid', c_传染病.主页id, 1);
      Zljsonputvalue(v_List, 'reg_no', c_传染病.挂号单, 0);
      Zljsonputvalue(v_List, 'advice_id', c_传染病.医嘱id, 1);
      Zljsonputvalue(v_List, 'eqpmtn_exetime', To_Char(c_传染病.检查时间, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_List, 'create_dept_id', c_传染病.登记科室id, 1);
      Zljsonputvalue(v_List, 'clinic_type', c_传染病.诊疗类别, 0);
      Zljsonputvalue(v_List, 'advice_doctor', c_传染病.开嘱医生, 0);
    
      Zljsonputvalue(v_List, 'spcm_send_dept', c_传染病.送检科室, 0);
      Zljsonputvalue(v_List, 'spcm_send_deptid', c_传染病.送检科室id, 1, 2);
    
      If Length(v_List) > 20000 Then
        Json_Out := Json_Out || v_List;
        v_List   := ',';
      End If;
    
    End Loop;
  End If;

  If v_List <> ',' Then
    Json_Out := Json_Out || v_List || ']}}';
  Else
    Json_Out := Json_Out || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getinfectinfo;
/
Create Or Replace Procedure Zl_Cissvr_Getinfectreport
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ------------------------------------------------
  --功能：查询阳性反馈单关联的疾病报告

  --入参      json
  --input
  --    pati_id                     N    1  病人id
  --    infctdz_name                C    1  疑似疾病

  --出参      json
  --output
  --    code                        N   1 应答吗：0-失败；1-成功
  --    message                     C   1 应答消息：失败时返回具体的错误信息
  --    report_list[]        疾病报告列表，支持多个，[数组]
  --       report_id                N   1  报告id
  --       rec_id                   N   1  反馈单ID
  --       create_time              C   1  创建时间
  --       report_name              C   1  病历名称
  --       infctdz_name             C   1  传染病名称
  ------------------------------------------------
  j_Json Pljson;
  j_In   Pljson;
  v_List Varchar2(32767);

  v_疑似疾病 Varchar2(500);
  n_病人id   Number(18);
Begin
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');

  n_病人id   := Nvl(j_Json.Get_Number('pati_id'), 0);
  v_疑似疾病 := j_Json.Get_String('infctdz_name');

  Json_Out := '{"output":{"code":1,"message":"成功","report_list":[';
  For c_报告 In (Select a.Id, b.Id As 反馈单id, a.创建时间, a.病历名称, b.传染病名称
               From 电子病历记录 A, 疾病阳性记录 B
               Where a.Id = b.文件id And a.病人id = b.病人id And b.病人id = n_病人id And b.传染病名称 = v_疑似疾病) Loop
    Zljsonputvalue(v_List, 'report_id', c_报告.Id, 1, 1);
    Zljsonputvalue(v_List, 'rec_id', c_报告.反馈单id, 1);
    Zljsonputvalue(v_List, 'create_time', To_Char(c_报告.创建时间, 'yyyy-mm-dd hh24:mi:ss'), 0);
    Zljsonputvalue(v_List, 'report_name', c_报告.病历名称, 0);
    Zljsonputvalue(v_List, 'infctdz_name', c_报告.传染病名称, 0, 2);
  
    If Length(v_List) > 20000 Then
      Json_Out := Json_Out || v_List;
      v_List   := ',';
    End If;
  End Loop;

  If v_List <> ',' Then
    Json_Out := Json_Out || v_List || ']}}';
  Else
    Json_Out := Json_Out || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getinfectreport;
/
Create Or Replace Procedure Zl_Cissvr_Getinpatistate
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：获取病人状态
  --入参：Json_In:格式
  --  input
  --   pati_id                  N  1  病人id
  --   pati_pageid              N  1  主页id
  --   pati_type                N  1  病人性质 0-普通住院病人 1-门诊留观病人 2-住院留观病人
  --出参: Json_Out,格式如下
  --  output
  --    code                    N 1  应答吗：0-失败；1-成功
  --    message                 C 1  应答消息：失败时返回具体的错误信息
  --    pati_state              N 1  病人状态
  --    out_time                C 1 出院日期
  --    out_type                C 1 出院方式
  ---------------------------------------------------------------------------
  n_病人id   病案主页.病人id%Type;
  n_主页id   病案主页.主页id%Type;
  n_病人性质 病案主页.病人性质%Type;
  d_出院日期 病案主页.出院日期%Type;
  v_出院方式 病案主页.出院方式%Type;
  n_住院性质 病案主页.病人性质%Type;
  n_State    Number;
  j_Json     Pljson;
  j_In       Pljson;
Begin

  --解析入参
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  n_主页id   := j_Json.Get_Number('pati_pageid');
  n_病人id   := j_Json.Get_Number('pati_id');
  n_病人性质 := j_Json.Get_Number('pati_type');
  Begin
    Select Nvl(状态, 0) 状态, 出院日期, 出院方式, 病人性质
    Into n_State, d_出院日期, v_出院方式,n_住院性质
    From 病案主页
    Where 病人id = n_病人id And 主页id = n_主页id And (病人性质 = n_病人性质 Or Nvl(n_病人性质, 0) = 0);
  Exception
    When Others Then
      Null;
  End;
  If Sql%RowCount > 0 Then
    Json_Out := '{"output":{"code":1,"message":"成功","pati_state":' || Zljsonstr(n_State, 1) || ',"pati_type":' || Zljsonstr(n_住院性质, 1) || ',"out_time":"' ||
                Zljsonstr(To_Char(d_出院日期, 'YYYY-MM-DD HH24:MI:SS'), 0) || '","out_type":"' || Zljsonstr(v_出院方式, 0) ||
                '"}}';
  Else
    Json_Out := '{"output":{"code":1,"message":"成功"}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getinpatistate;
/
Create Or Replace Procedure Zl_Cissvr_Getinsdeptinfor
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:获取所有有病人的住院科室
  --入参：Json_In:格式
  -- input
  --   opr_fun             N  1 执行功能 0-获取所有有病人的住院科室 1-通过科室id/病区id查找所有病人的入院科室或者病区 2-加载站点
  --   pati_source         N  1 病人来源：1-门诊；2-住院
  --   nodeno              C  1 站点编号
  --   wararea_ids         C  1 病人病区ids
  --   find_type           N  1 查找方式 0-按科室查找 1-按病区查找
  --   all_wararea         N  1 是否所有病区
  --   pati_in             N  1 是否在院
  --   dept_srvtype        C  1 服务对象,多个逗号分隔,如:1,2,3
  --出参: Json_Out,格式如下
  --  output
  --    code               C  1 应答码：0-失败；1-成功
  --    message            C  1 应答消息：失败时返回具体的错误信息
  --    dept_list[]             临床科室列表
  --      dept_id          N  1 科室id
  --      dept_code        C  1 科室编码
  --      dept_name        C  1 科室名称
  --      dept_spell       C  1 科室简码
  --      nodeno           C  1 站点
  ---------------------------------------------------------------------------
  j_Json     Pljson;
  j_In       Pljson;
  n_病人来源 Number(1);
  v_站点     部门表.站点%Type;
  v_病区ids  Varchar2(32767);

  n_执行功能 Number;
  n_查找方式 Number;
  n_所有病区 Number;
  n_在院     Number;
  v_服务对象 Varchar(100);
  v_Jtmp     Varchar2(32767);
  c_Jtmp     Clob;

Begin
  --解析入参
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  n_病人来源 := j_Json.Get_Number('pati_source');
  v_站点     := j_Json.Get_String('nodeno');
  v_病区ids  := j_Json.Get_String('wararea_ids');
  n_执行功能 := j_Json.Get_Number('opr_fun');
  n_查找方式 := j_Json.Get_Number('find_type');
  n_所有病区 := j_Json.Get_Number('all_wararea');
  n_在院     := j_Json.Get_Number('pati_in');
  v_服务对象 := Nvl(j_Json.Get_String('dept_srvtype'), '1,2,3');

  If Nvl(n_执行功能, 0) = 0 Then
    If Nvl(n_病人来源, 0) = 0 Then
      Json_Out := '{"output":{"code":0,"message":"未传入病人来源，请检查!"}}';
      Return;
    End If;
  
    If n_病人来源 = 1 Then
      For r_科室 In (Select Distinct a.Id, a.编码, a.名称, a.简码
                   From 部门表 A, 部门性质说明 B
                   Where a.Id = b.部门id And b.工作性质 = '临床' And Instr(',' || v_服务对象 || ',', ',' || b.服务对象 || ',') > 0 And
                         Exists (Select 1 From 床位状况记录 Where 病人id Is Not Null And 科室id = a.Id) And
                         (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And
                         (a.站点 = v_站点 Or a.站点 Is Null)) Loop
      
        v_Jtmp := v_Jtmp || ',{"dept_id":' || r_科室.Id;
        v_Jtmp := v_Jtmp || ',"dept_code":"' || Zljsonstr(r_科室.编码) || '"';
        v_Jtmp := v_Jtmp || ',"dept_name":"' || Zljsonstr(r_科室.名称) || '"';
        v_Jtmp := v_Jtmp || ',"dept_spell":"' || Zljsonstr(r_科室.简码) || '"';
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
      For r_科室 In (Select Distinct a.Id, a.名称, a.简码, a.编码
                   From 部门表 A, 在院病人 B, 部门性质说明 C
                   Where a.Id = b.病区id And a.Id = c.部门id And Instr(',' || v_服务对象 || ',', ',' || c.服务对象 || ',') > 0 And
                         (a.Id In (Select Column_Value From Table(f_Str2list(v_病区ids))) Or v_病区ids Is Null)) Loop
      
        v_Jtmp := v_Jtmp || ',{"dept_id":' || r_科室.Id;
        v_Jtmp := v_Jtmp || ',"dept_code":"' || Zljsonstr(r_科室.编码) || '"';
        v_Jtmp := v_Jtmp || ',"dept_name":"' || Zljsonstr(r_科室.名称) || '"';
        v_Jtmp := v_Jtmp || ',"dept_spell":"' || Zljsonstr(r_科室.简码) || '"';
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
  Elsif Nvl(n_执行功能, 0) = 1 Then
    If Nvl(n_查找方式, 0) = 0 Then
      --0-按科室查找
      For r_科室 In (Select a.Id, a.编码, a.名称, a.简码
                   From 部门表 A, 部门性质说明 B
                   Where a.Id = b.部门id And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And
                         b.工作性质 = '临床' And Instr(',' || v_服务对象 || ',', ',' || b.服务对象 || ',') > 0 And Exists
                    (Select 1
                          From 床位状况记录
                          Where 科室id = a.Id And
                                (Nvl(n_所有病区, 0) = 1 Or 病区id In (Select Column_Value From Table(f_Str2list(v_病区ids)))) And
                                (Nvl(n_在院, 0) = 0 Or 病人id Is Not Null)) And (a.站点 = v_站点 Or a.站点 Is Null)) Loop
        v_Jtmp := v_Jtmp || ',{"dept_id":' || r_科室.Id;
        v_Jtmp := v_Jtmp || ',"dept_code":"' || Zljsonstr(r_科室.编码) || '"';
        v_Jtmp := v_Jtmp || ',"dept_name":"' || Zljsonstr(r_科室.名称) || '"';
        v_Jtmp := v_Jtmp || ',"dept_spell":"' || Zljsonstr(r_科室.简码) || '"';
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
      --1-按病区查找
      For r_病区 In (Select a.Id, a.编码, a.名称, a.简码
                   From 部门表 A, 部门性质说明 B
                   Where a.Id = b.部门id And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And
                         b.工作性质 = '护理' And Instr(',' || v_服务对象 || ',', ',' || b.服务对象 || ',') > 0 And Exists
                    (Select 1
                          From 床位状况记录
                          Where 病区id = a.Id And
                                (Nvl(n_所有病区, 0) = 1 Or 病区id In (Select Column_Value From Table(f_Str2list(v_病区ids)))) And
                                (Nvl(n_在院, 0) = 0 Or 病人id Is Not Null)) And (a.站点 = v_站点 Or a.站点 Is Null)
                   Order By a.编码) Loop
        v_Jtmp := v_Jtmp || ',{"dept_id":' || r_病区.Id;
        v_Jtmp := v_Jtmp || ',"dept_code":"' || Zljsonstr(r_病区.编码) || '"';
        v_Jtmp := v_Jtmp || ',"dept_name":"' || Zljsonstr(r_病区.名称) || '"';
        v_Jtmp := v_Jtmp || ',"dept_spell":"' || Zljsonstr(r_病区.简码) || '"';
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
  Elsif Nvl(n_执行功能, 0) = 2 Then
    For r_科室 In (Select Distinct a.站点, c.名称
                 From 部门表 A, 部门性质说明 B, Zlnodelist C
                 Where a.Id = b.部门id And a.站点 = c.编号 And
                       (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And
                       ((b.工作性质 = '临床' And Nvl(n_查找方式, 0) = 0) Or (b.工作性质 = '护理' And Nvl(n_查找方式, 0) = 1)) And
                       Instr(',' || v_服务对象 || ',', ',' || b.服务对象 || ',') > 0 And
                       ((ID In (Select Distinct Decode(n_查找方式, 0, 科室id, 1, 病区id)
                                From 床位状况记录
                                Where 病人id Is Not Null) And Nvl(n_在院, 0) = 1) Or Nvl(n_在院, 0) = 0) And
                       ((a.Id In (Select Column_Value From Table(f_Str2list(v_病区ids))) And Nvl(n_所有病区, 0) = 1) Or
                       Nvl(n_所有病区, 0) = 0)
                 Order By a.站点) Loop
    
      v_Jtmp := v_Jtmp || ',{"dept_name":"' || Zljsonstr(r_科室.名称) || '"';
      v_Jtmp := v_Jtmp || ',"nodeno":"' || Zljsonstr(r_科室.站点) || '"';
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

  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"成功","dept_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","dept_list":[' || c_Jtmp || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getinsdeptinfor;
/
Create Or Replace Procedure Zl_Cissvr_Getmaxbedlen
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:获取病区床号的最大编号长度
  --入参：Json_In:格式
  -- input            病区id和科室ID节点二者只能传一个
  --   wardarea_id    N  0  病区id
  --   dept_id        N  0 科室ID
  --出参: Json_Out,格式如下
  --  output
  --    code               C  1 应答码：0-失败；1-成功
  --    message            C  1 应答消息：失败时返回具体的错误信息
  --    max_len            N  1 最大长度
  ---------------------------------------------------------------------------
  j_Json   Pljson;
  j_In     Pljson;
  n_病区id 床位状况记录.病区id%Type;
  n_科室id 床位状况记录.科室id%Type;
  n_长度   Number(20);

  n_传入病区id Number(1);
  n_传入科室id Number(1);

Begin
  --解析入参
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');
  If j_Json.Exist('wardarea_id') Then
    n_病区id     := j_Json.Get_Number('wardarea_id');
    n_传入病区id := 1;
  End If;

  If j_Json.Exist('dept_id') Then
    n_科室id     := j_Json.Get_Number('dept_id');
    n_传入科室id := 1;
  End If;

  If Nvl(n_传入病区id, 0) = 0 And Nvl(n_传入科室id, 0) = 0 Then
    Json_Out := Zljsonout('病区id和科室id必须传入一个，请检查');
    Return;
  End If;

  If Nvl(n_传入病区id, 0) = 1 Then
    If Nvl(n_病区id, 0) = 0 Then
      Select Max(Length(床号)) Into n_长度 From 床位状况记录 Where 状态 = '占用' And 病区id Is Not Null;
    Else
      Select Max(Length(床号)) Into n_长度 From 床位状况记录 Where 状态 = '占用' And 病区id = n_病区id;
    End If;
  Else
    If Nvl(n_科室id, 0) = 0 Then
      Select Max(Length(床号)) Into n_长度 From 床位状况记录 Where 状态 = '占用' And 科室id Is Not Null;
    Else
      Select Max(Length(床号)) Into n_长度 From 床位状况记录 Where 状态 = '占用' And 科室id = n_科室id;
    End If;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功","max_len":' || Nvl(n_长度, 0) || '}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getmaxbedlen;
/
Create Or Replace Procedure Zl_Cissvr_Getmedicalgroupid
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:根据病人变动记录来获取对应的医疗小组ID
  --入参：Json_In:格式
  --    input
  --      pati_id           N 1 病人id
  --      pati_pageid       N 1 主页id
  --      plcdept_id        N 1 开单科室ID
  --      placer            C 1 开单人
  --      occur_time        C 1 发生时间:yyyy-mm-dd hh24:mi:ss
  --出参: Json_Out,格式如下
  --  output
  --    code                  N 1 应答码：0-失败；1-成功
  --    message               C 1 应答消息：失败时返回具体的错误信息
  --    group_id              N 0 医疗小组id
  ---------------------------------------------------------------------------
  j_Json   Pljson;
  j_In     Pljson;
  n_病人id 病人变动记录.病人id%Type;
  n_主页id 病人变动记录.主页id%Type;

  n_开单科室id 病人变动记录.科室id%Type;
  v_开单人     人员表.姓名%Type;
  d_发生时间   病人变动记录.开始时间%Type;

  n_组id 病人变动记录.医疗小组id%Type;
Begin
  --解析入参
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');
  n_主页id := j_Json.Get_Number('pati_pageid');

  n_开单科室id := j_Json.Get_Number('plcdept_id');
  v_开单人     := j_Json.Get_String('placer');
  d_发生时间   := To_Date(j_Json.Get_String('occur_time'), 'yyyy-mm-dd hh24:mi:ss');

  n_组id := Zl_医疗小组_Get(n_开单科室id, v_开单人, n_病人id, n_主页id, d_发生时间);

  Json_Out := '{"output":{"code":1,"message":"成功","group_id":' || Nvl(n_组id, 0) || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getmedicalgroupid;
/
Create Or Replace Procedure Zl_Cissvr_Getmedrecscoreresult
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:获取病案评分结果
  --入参：Json_In:格式
  -- input
  --   pati_id              N 1 病人id
  --   pati_pageid          N 1 主页id
  --出参: Json_Out,格式如下
  --  output
  --    code                N  1 应答码：0-失败；1-成功
  --    message             C  1 应答消息：失败时返回具体的错误信息
  --    result              C  1  评份结果或等级:“甲”/“乙”/“丙”/“否”（即不合格）,多条时，取第一条。
  ---------------------------------------------------------------------------
  j_Json   Pljson;
  j_In     Pljson;
  n_病人id 病案评分结果.病人id%Type;
  n_主页id 病案评分结果.主页id%Type;
  v_等级   病案评分结果.等级%Type;

Begin
  --解析入参
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');
  n_主页id := j_Json.Get_Number('pati_pageid');

  If Nvl(n_病人id, 0) = 0 Or Nvl(n_主页id, 0) = 0 Then
    Json_Out := Zljsonout('未传入病人id或主页id，请检查！');
    Return;
  End If;

  Select Max(等级) Into v_等级 From 病案评分结果 Where 病人id = n_病人id And 主页id = n_主页id;
  Json_Out := '{"output":{"code":1,"message":"成功","result":"' || v_等级 || '"}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getmedrecscoreresult;
/
Create Or Replace Procedure Zl_Cissvr_Getnextid
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --功能：读取指定表名对应的序列(按规范，其序列名称为“表名称_id”)的下一数值
  --入参：Json_In:格式
  --input
  --  table_name    C  1 表名
  --  col_name      C  1 字段名  序列名称不一定是ID，例如记录ID
  -- 出参:
  --  output
  --  next_id      N   1  序列
  -------------------------------------------

  v_Table Varchar2(500);
  v_Col   Varchar2(500);

  n_Nextid Number(18);
  j_Json   Pljson;
  j_In     Pljson;
Begin
  --解析入参
  j_In    := Pljson(Json_In);
  j_Json  := j_In.Get_Pljson('input');
  v_Table := j_Json.Get_String('table_name');
  v_Col   := Nvl(j_Json.Get_String('col_name'), 'ID');

  Execute Immediate 'select ' || v_Table || '_' || Nvl(v_Col, 'ID') || '.nextval from dual'
    Into n_Nextid;

  Json_Out := '{"output":{"code":1,"message":"成功","next_id":' || Nvl(n_Nextid, 0) || '}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Getnextid;
/
Create Or Replace Procedure Zl_Cissvr_Getpatiallergyinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：获取病人过敏信息
  --input      获取病人过敏信息
  --  pati_id               N  1  病人id
  --  visit_id              N  1  标识号：挂号id（门诊），主页id（住院）
  --output
  --  code                  C  1  应答码：0-失败；1-成功
  --  message               C  1  应答消息：失败时返回具体的错误信息
  --  allergy_list[]    过敏信息，[数组]
  --     drug_name          C  1  药物名称
  --     allergy_time       C  1  过敏时间
  --     allergy_info       C  1  过敏反应
  ---------------------------------------------------------------------------
  j_Input  Pljson;
  j_In     Pljson;
  n_病人id Number(18);
  n_标识id Number(18); --门诊：挂号id ，住院：主页id

  v_Jtmp Varchar2(32767);

Begin
  j_In     := Pljson(Json_In);
  j_Input  := j_In.Get_Pljson('input');
  n_病人id := j_Input.Get_Number('pati_id');
  n_标识id := j_Input.Get_Number('visit_id');

  For v_过敏记录 In (Select Distinct a.药物名, Nvl(a.过敏时间, a.记录时间) As 过敏时间, a.过敏反应
                 From 病人过敏记录 A, 病人挂号记录 B, 病案主页 C, 部门表 D, 部门表 E
                 Where a.病人id = b.病人id(+) And a.主页id = b.Id(+) And b.记录性质(+) = 1 And b.记录状态(+) = 1 And
                       a.病人id = c.病人id(+) And a.主页id = c.主页id(+) And b.执行部门id = d.Id(+) And c.出院科室id = e.Id(+) And
                       a.结果 = 1 And 药物名 Is Not Null And a.病人id = n_病人id And a.主页id = n_标识id And Not Exists
                  (Select 药物id
                        From 病人过敏记录
                        Where (Nvl(药物id, 0) = Nvl(a.药物id, 0) Or Nvl(药物名, 'Null') = Nvl(a.药物名, 'Null')) And Nvl(结果, 0) = 0 And
                              记录时间 > a.记录时间 And 病人id = n_病人id And 主页id = n_标识id)
                 Order By Nvl(a.过敏时间, a.记录时间) Desc) Loop
  
    v_Jtmp := v_Jtmp || ',{"drug_name":"' || Zljsonstr(v_过敏记录.药物名) || '"';
    v_Jtmp := v_Jtmp || ',"allergy_time":"' || To_Char(v_过敏记录.过敏时间, 'YYYY-MM-DD HH24:MI') || '"';
  
    v_Jtmp := v_Jtmp || ',"allergy_info":"' || Zljsonstr(v_过敏记录.过敏反应) || '"';
    v_Jtmp := v_Jtmp || '}';
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"成功","allergy_list":[' || Substr(v_Jtmp, 2) || ']}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zltools.Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getpatiallergyinfo;
/

Create Or Replace Procedure Zl_Cissvr_Getpatibaseinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：获取病人基本信息获取(临床)
  --入参：Json_In:格式
  --  input
  --    query_type        N 1 查询方式-- 1-通过病人ID+主页ID查询病人信息,2-通过医嘱ID获取病人基本信息 ,3-通过挂号单获取病人基本信息


  --    pati_id           N 1 病人id--
  --    page_id           N 1 主页id--
  --    advice_id         N 1 医嘱ID--
  --    pati_type         N 1 病人类型，0-住院病人 1-门诊病人

  --出参: Json_Out,格式如下
  --  output
  --    code                 N 1 应答吗：0-失败；1-成功
  --    message              C 1 应答消息：失败时返回具体的错误信息
  --    page_list[]        1  数据组
  --       pati_id           N 1 病人id
  --       page_id           N 1 主页id
  --       pati_name         C 1 病人姓名
  --       pati_sex          C 1 病人性别
  --       pati_age          C 1 病人年龄
  --       dept_name         C 1 科室名称
  --       inpatient_num     C 1 住院号
  --       pati_bed          C 1 当前床号
  --       dept_id           N 1 科室id
  --       regist_no         C 1 挂号单
  --       registration_time C 1 就诊时间
  --       adtd_time         C 1 出院时间

  --       pati_content      C 1 当前病况
  --       insurance_type    N 1 险类
  --       pati_wardarea_id  N 1 当前病区ID

  --       reg_id            N 1 挂号id
  --       outpatient_num    C 1 门诊号
  --       return_visit      N 1 复诊标志
  --       outp_room_name    C 1 门诊诊室名称
  ---------------------------------------------------------------------------
  j_In       Pljson;
  j_Json     Pljson;
  n_查询方式 Number;
  n_病人id   Number;
  n_主页id   Number;
  n_医嘱id   Number;
  n_病人类型 Number;
  v_List     Varchar2(32767);

  v_挂号单 Varchar(50);

  v_Err_Msg Varchar(2000);
  Err_Item Exception;
Begin
  --解析入参
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  n_查询方式 := j_Json.Get_Number('query_type');
  n_病人类型 := j_Json.Get_Number('pati_type');
  Json_Out   := '{"output":{"code":1,"message":"成功","page_list":[';
  If n_查询方式 = 1 Then
    n_病人id := j_Json.Get_Number('pati_id');
    n_主页id := j_Json.Get_Number('page_id');
    If Nvl(n_病人类型, 0) = 0 Then
      For R In (Select a.姓名, a.性别, a.年龄, b.名称 As 科室, a.住院号, a.出院病床 As 床号,
                       To_Char(a.入院日期, 'yyyy-MM-dd HH24:MI:SS') As 就诊时间, To_Char(a.出院日期, 'yyyy-MM-dd HH24:MI:SS') As 出院时间,
                       b.Id As 科室id, a.当前病况, a.险类, a.当前病区id
                From 病案主页 A, 部门表 B
                Where a.出院科室id = b.Id And a.病人id = n_病人id And a.主页id = Decode(n_主页id, Null, a.主页id, n_主页id)
                Order By Nvl(a.主页id, 0)) Loop
        Zljsonputvalue(v_List, 'pati_id', n_病人id, 1, 1);
        Zljsonputvalue(v_List, 'page_id', n_主页id, 1);
        Zljsonputvalue(v_List, 'pati_name', r.姓名);
        Zljsonputvalue(v_List, 'pati_sex', r.性别);
        Zljsonputvalue(v_List, 'pati_age', r.年龄);
        Zljsonputvalue(v_List, 'dept_name', r.科室);
        Zljsonputvalue(v_List, 'inpatient_num', r.住院号, 0);
        Zljsonputvalue(v_List, 'pati_bed', r.床号);
        Zljsonputvalue(v_List, 'registration_time', r.就诊时间);
        Zljsonputvalue(v_List, 'adtd_time', r.出院时间);
        Zljsonputvalue(v_List, 'dept_id', r.科室id, 1);
      
        Zljsonputvalue(v_List, 'pati_content', r.当前病况);
        Zljsonputvalue(v_List, 'insurance_type', r.险类, 1);
        Zljsonputvalue(v_List, 'pati_wardarea_id', r.当前病区id, 1, 2);
      
        If Length(v_List) > 20000 Then
          Json_Out := Json_Out || v_List;
          v_List   := ',';
        End If;
      End Loop;
      If v_List <> ',' Then
        Json_Out := Json_Out || v_List || ']}}';
      Else
        Json_Out := Json_Out || ']}}';
      End If;
    Else
      For R In (Select a.No, a.病人id, a.姓名, a.性别, a.年龄
                From 病人挂号记录 A
                Where a.病人id = n_病人id And a.Id = n_主页id) Loop
        Zljsonputvalue(v_List, 'pati_id', n_病人id, 1, 1);
        Zljsonputvalue(v_List, 'pati_name', r.姓名);
        Zljsonputvalue(v_List, 'pati_sex', r.性别);
        Zljsonputvalue(v_List, 'pati_age', r.年龄);
        Zljsonputvalue(v_List, 'regist_no', r.No, 0, 2);
        If Length(v_List) > 20000 Then
          Json_Out := Json_Out || v_List;
          v_List   := ',';
        End If;
      End Loop;
      If v_List <> ',' Then
        Json_Out := Json_Out || v_List || ']}}';
      Else
        Json_Out := Json_Out || ']}}';
      End If;
    End If;
  Elsif n_查询方式 = 2 Then
    n_医嘱id := j_Json.Get_Number('advice_id');
    For R In (Select a.病人id, a.主页id, Nvl(q.婴儿姓名, a.姓名) 姓名, Nvl(q.婴儿性别, a.性别) 性别,
                     Decode(q.序号, Null, a.年龄, Round(Decode(q.死亡时间, Null, Sysdate, q.死亡时间) - q.出生时间) || '天') 年龄,
                     b.名称 As 科室, Null As 住院号, Null As 床号, To_Char(a.开始执行时间, 'yyyy-MM-dd HH24:MI:SS') As 就诊时间,
                     b.Id As 科室id, Null As 当前病况, Null As 险类, Null As 当前病区id
              From 病人医嘱记录 A, 部门表 B, 病人新生儿记录 Q
              Where a.病人科室id = b.Id And a.Id = n_医嘱id And a.病人id = q.病人id(+) And a.主页id = q.主页id(+) And a.婴儿 = q.序号(+)) Loop
      Zljsonputvalue(v_List, 'pati_id', r.病人id, 1, 1);
      Zljsonputvalue(v_List, 'page_id', r.主页id, 1);
      Zljsonputvalue(v_List, 'pati_name', r.姓名);
      Zljsonputvalue(v_List, 'pati_sex', r.性别);
      Zljsonputvalue(v_List, 'pati_age', r.年龄);
      Zljsonputvalue(v_List, 'dept_name', r.科室);
      Zljsonputvalue(v_List, 'inpatient_num', r.住院号, 0);
      Zljsonputvalue(v_List, 'pati_bed', r.床号);
      Zljsonputvalue(v_List, 'registration_time', r.就诊时间);
      Zljsonputvalue(v_List, 'dept_id', r.科室id, 1);
    
      Zljsonputvalue(v_List, 'pati_content', r.当前病况);
      Zljsonputvalue(v_List, 'insurance_type', r.险类, 1);
      Zljsonputvalue(v_List, 'pati_wardarea_id', r.当前病区id, 1, 2);
    
      If Length(v_List) > 20000 Then
        Json_Out := Json_Out || v_List;
        v_List   := ',';
      End If;
    End Loop;
    If v_List <> ',' Then
      Json_Out := Json_Out || v_List || ']}}';
    Else
      Json_Out := Json_Out || ']}}';
    End If;
  Elsif n_查询方式 = 3 Then
    v_挂号单 := j_Json.Get_String('reg_no');
    n_病人id := Nvl(j_Json.Get_Number('pati_id'), 0);
  
    For R In (Select a.姓名, a.性别, a.年龄, b.名称 As 科室, a.门诊号, Null As 床号, a.发生时间 As 就诊时间, b.Id As 科室id, a.复诊, a.病人id, a.Id,
                     Null As 当前病况, a.险类, Null As 当前病区id, a.诊室
              From 病人挂号记录 A, 部门表 B
              Where a.执行部门id = b.Id And a.记录性质 = 1 And a.记录状态 = 1 And a.No = v_挂号单 And (n_病人id = 0 Or (n_病人id = a.病人id))) Loop
      Zljsonputvalue(v_List, 'pati_id', r.病人id, 1, 1);
      Zljsonputvalue(v_List, 'reg_id', r.Id, 1);
      Zljsonputvalue(v_List, 'pati_name', r.姓名);
      Zljsonputvalue(v_List, 'pati_sex', r.性别);
      Zljsonputvalue(v_List, 'pati_age', r.年龄);
      Zljsonputvalue(v_List, 'dept_name', r.科室);
      Zljsonputvalue(v_List, 'outpatient_num', r.门诊号, 0);
      Zljsonputvalue(v_List, 'pati_bed', r.床号);
      Zljsonputvalue(v_List, 'registration_time', To_Date(r.就诊时间, 'yyyy-MM-dd HH24:MI:SS'));
      Zljsonputvalue(v_List, 'dept_id', r.科室id, 1);
      Zljsonputvalue(v_List, 'return_visit', r.复诊, 1);
      Zljsonputvalue(v_List, 'outp_room_name', r.诊室);
    
      Zljsonputvalue(v_List, 'pati_content', r.当前病况);
      Zljsonputvalue(v_List, 'insurance_type', r.险类, 1);
      Zljsonputvalue(v_List, 'pati_wardarea_id', r.当前病区id, 1, 2);
      If Length(v_List) > 20000 Then
        Json_Out := Json_Out || v_List;
        v_List   := ',';
      End If;
    End Loop;
    If v_List <> ',' Then
      Json_Out := Json_Out || v_List || ']}}';
    Else
      Json_Out := Json_Out || ']}}';
    End If;
  End If;
Exception
  When Err_Item Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr('-20101:[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]') || '"}}';
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getpatibaseinfo;
/

Create Or Replace Procedure Zl_Cissvr_Getpatichangerec
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:根据条件获取病人变动信息
  --入参：Json_In:格式
  -- input
  --   query_type           N 1 查询方式：0-根据指定条件查询病人变动信息，1-仅查询病人某次住院的转科信息
  --   pati_id              N 1 病人id
  --   pati_pageid          N 1 主页id
  --   pati_wararea_id      N 0 病区id
  --   pati_dept_id         N 0 科室id
  --   start_reasons        C 0 开始原因s:多个用逗号分离,如:3,15,10,1
  --   stop_reasons         C 0 终止原因s:多个用逗号分离,如:3,15,10,1
  --出参: Json_Out,格式如下
  --  output
  --    code               C  1 应答码：0-失败；1-成功
  --    message            C  1 应答消息：失败时返回具体的错误信息
  --    change_list[]
  --        start_time         C 1 开始时间:yyyy-mm-dd hh24:mi:ss
  --        start_reason       N 1 开始原因
  --        dept_name          C 1 部门名称
  --        stop_time          C 1 终止时间:yyyy-mm-dd hh24:mi:ss
  --        stop_reason        N 1 终止原因
  ---------------------------------------------------------------------------
  j_Json      Pljson;
  j_In        Pljson;
  n_病人id    病人变动记录.病人id%Type;
  n_主页id    病人变动记录.主页id%Type;
  n_病区id    病人变动记录.病区id%Type;
  n_科室id    病人变动记录.科室id%Type;
  v_终止原因  Varchar2(3000);
  v_开始原因  Varchar2(200);
  n_查询方式  Number(2);
  c_Jtmp      Clob; 
  v_Temp      Varchar2(32767);
Begin
  --解析入参
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  n_查询方式 := j_Json.Get_Number('query_type');
  n_病人id   := j_Json.Get_Number('pati_id');
  n_主页id   := j_Json.Get_Number('pati_pageid');

  n_病区id   := j_Json.Get_Number('pati_wararea_id');
  n_科室id   := j_Json.Get_Number('pati_dept_id');
  v_开始原因 := j_Json.Get_String('start_reasons');
  v_终止原因 := j_Json.Get_String('stop_reasons');

  If Nvl(n_病人id, 0) = 0 Or Nvl(n_主页id, 0) = 0 Then
    Json_Out := '{"output":{"code":0,"message":"必须传入病人id、主页id！"}}';
    Return;
  End If; 
  
  If n_查询方式 = 1 Then
    v_Temp := Null;
    For R In (Select Distinct 1 As 开始原因, To_Date('1900-01-01', 'yyyy-mm-dd') As 开始时间, b.名称
              From 病人变动记录 A, 部门表 B
              Where a.科室id = b.Id And a.开始时间 Is Not Null And a.开始原因 In (1, 2) And a.病人id = n_病人id And 主页id = n_主页id
              Union All
              Select a.开始原因, a.开始时间, b.名称
              From 病人变动记录 A, 部门表 B
              Where a.科室id = b.Id And a.开始时间 Is Not Null And a.开始原因 = 3 And a.病人id = n_病人id And 主页id = n_主页id
              Order By 开始时间) Loop
      v_Temp := v_Temp || ',{"start_reason":' || r.开始原因 || ',"dept_name":"' || Zljsonstr(r.名称) || '"}';
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"成功","change_list":[' || Substr(v_Temp, 2) || ']}}';
    Return;
  End If;

  v_Temp := Null;
  For r_变动 In (Select a.开始原因, To_Char(a.开始时间, 'yyyy-mm-dd hh24:mi:ss') As 开始时间, b.名称,
                      To_Char(a.终止时间, 'yyyy-mm-dd hh24:mi:ss') As 终止时间, a.终止原因
               From 病人变动记录 A, 部门表 B
               Where a.科室id = b.Id And ((a.开始时间 Is Not Null And Nvl(v_开始原因, '-') <> '-') Or Nvl(v_开始原因, '-') = '-') And
                     (Instr(',' || v_开始原因 || ',', ',' || a.开始原因 || ',') > 0 Or Nvl(v_开始原因, '-') = '-') And
                     a.病人id = n_病人id And a.主页id = n_主页id And (a.病区id = n_病区id Or Nvl(n_病区id, 0) = 0) And
                     (a.科室id = n_科室id Or Nvl(n_科室id, 0) = 0) And
                     (Instr(',' || v_终止原因 || ',', ',' || a.终止原因 || ',') > 0 Or Nvl(v_终止原因, '-') = '-')
               Order By 开始时间, 终止时间) Loop
  
    v_Temp := v_Temp || ',{"start_time":"' || Zljsonstr(r_变动.开始时间) || '"';
    v_Temp := v_Temp || ',"start_reason":' || Nvl(r_变动.开始原因, 0);
    v_Temp := v_Temp || ',"dept_name":"' || Zljsonstr(r_变动.名称) || '"';
    v_Temp := v_Temp || ',"stop_time":"' || Zljsonstr(r_变动.终止时间) || '"';
    v_Temp := v_Temp || ',"stop_reason":' || Nvl(r_变动.终止原因, 0);
    v_Temp := v_Temp || '}';
  
    If Length(v_Temp) > 30000 Then
      If c_Jtmp Is Null Then
        c_Jtmp := Substr(v_Temp, 2);
      Else
        c_Jtmp := c_Jtmp || v_Temp;
      End If;
      v_Temp := Null;
    End If;
  
  End Loop;
  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"成功","change_list":[' || Substr(v_Temp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Temp;
    Json_Out := '{"output":{"code":1,"message":"成功","change_list":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getpatichangerec;
/
Create Or Replace Procedure Zl_Cissvr_Getpatidiagnose
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：获取病人获取病人诊断信息
  --input      获取病人诊断信息
  --  pati_id               N  1  病人ID
  --  visit_id              N  1  就诊ID
  --output
  --  code                  C  1  应答码：0-失败；1-成功
  --  message               C  1  应答消息：失败时返回具体的错误信息
  --  diagnose_list[]      诊断内容，[数组]
  --     diag_origin        N  1  记录来源
  --     diag_type          N  1  诊断类型
  --     diag_order         N  1 诊断次序
  --     diag_description   C  1  诊断描述
  --     diag_distrustful   N  1  是否疑诊
  --     diag_record_time   C  1  记录日期
  ---------------------------------------------------------------------------
  j_Input  Pljson;
  j_In     Pljson;
  n_病人id Number(18);
  n_就诊id Number(18);
  v_Output Varchar2(32767);
Begin
  j_In     := Pljson(Json_In);
  j_Input  := j_In.Get_Pljson('input');
  n_病人id := j_Input.Get_Number('pati_id');
  n_就诊id := j_Input.Get_Number('visit_id');

  For r_Diagnose In (Select 记录来源, 诊断类型, 诊断次序, 诊断描述, 是否疑诊, 记录日期
                     From 病人诊断记录
                     Where 病人id = n_病人id And 主页id = n_就诊id And Nvl(录入次序, '01') = '01' And Nvl(编码类别, 'E') = 'E'
                     Order By 记录日期 Desc, 诊断类型 Desc) Loop
  
    Zljsonputvalue(v_Output, 'diag_origin', Nvl(r_Diagnose.记录来源, 0), 1, 1);
    Zljsonputvalue(v_Output, 'diag_type', Nvl(r_Diagnose.诊断类型, 0), 1);
    Zljsonputvalue(v_Output, 'diag_order', r_Diagnose.诊断次序, 1);
    Zljsonputvalue(v_Output, 'diag_description', Nvl(r_Diagnose.诊断描述, ''));
    Zljsonputvalue(v_Output, 'diag_distrustful', Nvl(r_Diagnose.是否疑诊, 0), 1);
    Zljsonputvalue(v_Output, 'diag_record_time', Nvl(r_Diagnose.记录日期, ''), 0, 2);
  End Loop;

  Json_Out := '{"output":{"code":1,"message": "成功","diagnose_list":[' || v_Output || ']}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getpatidiagnose;
/
Create Or Replace Procedure Zl_Cissvr_Getpatiid
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：根据床号、住院呈获取病人ID及主页ID
  --input
  --   wardarea_id          N 1 当前病区id
  --   pati_bed             C 1 当前床号
  --   inpatient_num        C 1 住院号
  --   obsv_no              C 1 留观号
  --output
  --    code                N 1 应答码：0-失败；1-成功
  --    message             C 1 应答消息： 失败时返回具体的错误信息
  --    pati_id             N 1 病人ID:未找到时也成功，返回0
  --    pati_pageid         N   主页ID
  ---------------------------------------------------------------------------
  j_In     PLJson;
  j_Input  PLJson;
  n_主页id 病案主页.主页id%Type;
  n_病人id 病案主页.病人id%Type;

  n_当前病区id 病案主页.当前病区id%Type;
  v_当前床号   病案主页.入院病床%Type;

  n_住院号 病案主页.住院号%Type;
  n_留观号 病案主页.留观号%Type;
Begin
  j_In         := PLJson(Json_In);
  j_Input      := j_In.Get_Pljson('input');
  n_当前病区id := j_Input.Get_Number('wardarea_id');
  v_当前床号   := j_Input.Get_String('pati_bed');
  n_住院号     := To_Number(j_Input.Get_String('inpatient_num'));
  n_留观号     := To_Number(j_Input.Get_String('obsv_no'));

  If Nvl(n_留观号, 0) <> 0 Then
    Select Max(病人id), Max(主页id) Into n_病人id, n_主页id From 病案主页 Where 留观号 = n_留观号;
  Elsif Nvl(n_住院号, 0) <> 0 Then
    Select Max(病人id), Max(主页id) Into n_病人id, n_主页id From 病案主页 Where 住院号 = n_住院号;
  Else
    Select Max(a.病人id), Max(a.主页id)
    Into n_病人id, n_主页id
    From 在院病人 A, 床位状况记录 B
    Where a.病人id = b.病人id And b.病区id = Nvl(n_当前病区id, 0) And b.床号 = v_当前床号;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功","pati_id":' || Nvl(n_病人id, 0) || ',"pati_pageid":' || Nvl(n_主页id, 0) || '}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || zlJsonStr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getpatiid;
/
Create Or Replace Procedure Zl_Cissvr_Getpatiidbyinpno
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --功能：获取病人最后一次挂号的挂号记录
  --入参：Json_In:格式
  --input
  --  inpatient_num         C   1 住院号
  --出参      json
  --output
  --    code                N   1 应答吗：0-失败；1-成功
  --    message             C   1 应答消息：失败时返回具体的错误信息
  --    pati_id             N   1 病人id
  -------------------------------------------
  j_Json   Pljson;
  j_In     Pljson;
  n_病人id Number(18);
  n_住院号 Number;
Begin
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_住院号 := j_Json.Get_String('inpatient_num');
  If Nvl(n_住院号, 0) <> 0 Then
    Begin
      Select Nvl(Max(病人id), 0) As 病人id Into n_病人id From 病案主页 Where 住院号 = n_住院号;
    Exception
      When Others Then
        Null;
    End;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功","pati_id":' || Nvl(n_病人id, 0) || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getpatiidbyinpno;
/
Create Or Replace Procedure Zl_Cissvr_Getpatimaxpageid
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:获取指定条件的最大主页ID
  --入参：Json_In:格式
  --    input
  --      pati_id           N 1 病人id
  --      pati_wararea_id   N 1 当前病区iD
  --出参: Json_Out,格式如下
  --  output
  --    code                  N 1 应答码：0-失败；1-成功
  --    message               C 1 应答消息：失败时返回具体的错误信息
  --    pati_pageid           N 1 主页id
  ---------------------------------------------------------------------------

  j_Json Pljson;
  j_In   Pljson;

  n_病区id 病案主页.当前病区id%Type;
  n_病人id 病案主页.病人id%Type;
  n_主页id 病案主页.主页id%Type;
Begin
  --解析入参
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');
  n_病区id := j_Json.Get_Number('pati_wararea_id');

  Select Max(主页id)
  Into n_主页id
  From 病案主页
  Where 病人id = n_病人id And Decode(n_病区id, Null, 0, 当前病区id) = Decode(n_病区id, Null, 0, n_病区id);

  Json_Out := '{"output":{"code":1,"message":"成功","pati_pageid":' || Nvl(n_主页id, 0) || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getpatimaxpageid;
/

Create Or Replace Procedure Zl_Cissvr_Getpatipageextinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:根据病人id及主页id获取病案主页从表信息
  --入参：Json_In:格式
  --    input
  --      pati_id           N 1 病人id
  --      pati_pageid       N 1 主页id
  --      info_names        C 1 信息名：多个用逗号

  --出参: Json_Out,格式如下
  --  output
  --    code                  N 1 应答码：0-失败；1-成功
  --    message               C 1 应答消息：失败时返回具体的错误信息
  --    pati_Name             C 1 姓名
  --    ext_list[]             病案从表信息列表
  --      info_name           C 1 信息名
  --      info_value          C 1 信息值
  ---------------------------------------------------------------------------
  j_In     Pljson;
  j_Json   Pljson;
  n_病人id 病案主页.病人id%Type;
  n_主页id 病案主页.主页id%Type;
  v_信息名 Varchar2(32767);
  v_List   Varchar2(32767);
Begin
  --解析入参
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');
  n_主页id := j_Json.Get_Number('pati_pageid');
  v_信息名 := j_Json.Get_String('info_names');

  If Nvl(n_病人id, 0) = 0 Then
    Json_Out := Zljsonout('未传入病人id，请检查！');
    Return;
  End If;

  If Nvl(n_主页id, 0) = 0 Then
    Json_Out := Zljsonout('未传入主页id，请检查！');
    Return;
  End If;

  If Nvl(v_信息名, '-') = '-' Then
    Json_Out := Zljsonout('未传入信息名，请检查！');
    Return;
  End If;

  For r_信息 In (Select a.信息名, a.信息值
               From 病案主页从表 A,
                    (Select /*+cardinality(B,10) */
                       Column_Value As 信息名
                      From Table(f_Str2list(v_信息名))) B
               Where a.信息名 = b.信息名 And a.病人id = n_病人id And a.主页id = n_主页id) Loop
    v_List := v_List || ',{"info_name":"' || Zljsonstr(r_信息.信息名) || '"';
    v_List := v_List || ',"info_value":"' || Zljsonstr(r_信息.信息值) || '"';
    v_List := v_List || '}';
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"成功","ext_list":[' || Substr(v_List, 2) || ']}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getpatipageextinfo;
/
Create Or Replace Procedure Zl_Cissvr_Getpatipageinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:获取病案主页相关信息
  --入参：Json_In:格式
  --    input
  --      query_type          C 1 查询类型:0-基本信息;1-基本信息的展;2-仅取主页
  --      pati_pageids        C 1 病人信息,格式两种:一种是:病人id:主页ID,…;一种：病人id,…
  --      is_babyinfo         N 1 是否包含婴儿信息:1-包含;0-不包含
  --      is_transdeptinfo    N 1 是否包含转科信息:1-包含;0-不包含
  --      is_lastpage         N 1 是否取最后一次住院
  --      pati_natures        C 1 病人性质:0-普通住院病人,1-门诊留观病人,2-住院留观病人；多个逗号分隔，不传为所有
  --      rgst_id             N 1 挂号ID,根据挂号ID查询
  --      is_badinfo          N 0 是否包含床位信息:1-包含;0-不包含
  --出参: Json_Out,格式如下
  --  output
  --    code                  N 1 应答码：0-失败；1-成功
  --    message               C 1 应答消息：失败时返回具体的错误信息
  --    pati_count            N 1 查询的病人信息条数
  --    page_list[]             1 数据组
  --      pati_id             N 1 病人id
  --      pati_pageid         N 1 主页id
  --      pati_name           C 1 姓名
  --      pati_sex            C 1 性别
  --      pati_age            C 1 年龄
  --      inpatient_num       C 1 住院号
  --      fee_category        C 1 费别
  --      mdlpay_mode_name    C 1 医疗付款方式名称
  --      mdlpay_mode_code    C 1 医疗付款方式编码
  --      pati_bed            C 1 当前床号
  --      pati_type           C 1 病人类型(普通，医保，留观)
  --      pati_show_color     N 1 病人显示颜色
  --      pati_education      C 1 学历
  --      ocpt_name           C 1 职业
  --      country_name        C 1 国籍
  --      pati_marital_cstatus  C 1 婚姻状况
  --      pati_nature         N 1 病人性质:0-普通住院病人,1-门诊留观病人,2-住院留观病人
  --      audit_sign          N 1 审核标志:病案主页.审核标志
  --      si_inp_status       N 1 住院状态:病案主页.状态(0-正常住院；1-尚未入科；2-正在转科或正在转病区；3-已预出院)
  --      pati_wardarea_id    N 1 当前病区id
  --      pati_wardarea_name  C 1 当前病区名称
  --      pati_dept_id        N 1 当前科室id
  --      pati_dept_name      C 1 当前科室名称
  --      adta_time           C 1 入院时间:yyyy-mm-dd hh24:mi:ss
  --      adtd_time           C 1 出院时间:yyyy-mm-dd hh24:mi:ss
  --      insurance_type      N 1 险类
  --      rgst_id             N 1 挂号id
  --      catalog_date        C 1 编目日期:yyyy-mm-dd hh24:mi:ss
  --      in_objective        C 1 住院目的
  --      reg_name            C 1 登记人
  --      reg_date            C 1 住院登记时间
  --      pat_rsdpscn         C 1 住院医师
  --      pati_desc           C 1 病人备注
  --      insurance_num       C 1 医保号
  --      outpatient_doctor   C 1 门诊医师
  --      responsible_nurse   C 1 责任护士
  --      hospital_admissions C 1 入院病况
  --      current_conditions  C 1 当前病况
  --      hospital_days       N 1 住院天数
  --      hospital_dept       C 1 入院科室
  --      level_of_care       C 1 护理等级
  --      level_of_bed        C 1 床位等级
  --      in_dept             N 1 是否已入科
  --      pati_home_addr      C 1 家庭地址
  --      pati_house_addr     C 1 户口地址
  --      pati_contact_addr   C 1 联系人地址
  --      baby_list[]           1 婴儿信息，[数组]
  --        pati_id           N 1 病人id
  --        pati_pageid       N 1 主页id
  --        baby_num          N 1 婴儿序号
  --        baby_name         C 1 婴儿姓名
  --        baby_sex          C 1 婴儿性别
  --        baby_date         C 1 出生时间
  --      trans_list[]        C   转科列表信息
  --        start_reason      C 1 开始原因
  --        start_time        C 1 开始时间:yyyy-mm-dd hh24:mi:ss
  --        dept_name         C 1 科室名称
  --      badinfo_list[]      床位信息，[数组]
  --        wardarea_id       N 1 病区id
  --        wardarea_name     C 1 病区名称
  --        bed_no            C 1 床号
  --        bed_class_code    C 1 分类编码
  --        bed_class_name    C 1 分类名称
  ---------------------------------------------------------------------------
  j_In           Pljson;
  j_Json         Pljson;
  n_类型         Number(1);
  v_病人信息     Varchar2(32767);
  n_包含婴儿信息 Number(1);
  n_包含转科信息 Number(1);
  n_包含床位信息 Number(1);
  n_最后一次住院 Number(1);
  n_查询方式     Number(1); --0：通过病人id:主页id的形式读取；1：通过病人id读取
  n_已入科       Number(1);

  v_病人性质 Varchar2(32767);
  n_挂号id   病案主页.挂号id%Type;
  I          Number(6);
  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_病人信息 Collection_Type;

  Json_Out_Tmp Clob; --全局变量，Setreturnjson 子方法中赋值 ，忽任意使用
  v_List       Varchar2(32767); --全局变量，Setreturnjson 子方法中赋值 ，忽任意使用

  --所有的游标都是用的一个固定结构
  --该游标只为定义RowType结构，而不是要查询数据。
  Cursor c_病案信息 Is
    Select a.病人id, a.主页id, a.姓名, a.性别, a.年龄, a.住院号, a.费别, a.病人性质, a.审核标志, a.状态 As 住院状态,
           To_Char(a.入院日期, 'yyyy-mm-dd hh24:mi:ss') As 入院时间, To_Char(a.出院日期, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.住院目的,
           a.登记人, To_Char(a.登记时间, 'yyyy-mm-dd') As 登记时间, a.住院医师, c.名称 As 医疗付款方式名称, c.编码 医疗付款方式编码, a.出院病床 As 当前床号, a.病人类型,
           a.学历, a.职业, a.国籍, a.婚姻状况, a.当前病区id, D1.名称 As 当前病区名称, a.出院科室id As 当前科室id, D2.名称 As 当前科室名称, a.险类, a.挂号id,
           To_Char(a.编目日期, 'yyyy-mm-dd hh24:mi:ss') As 编目日期, a.备注, a.门诊医师, a.责任护士, a.入院病况, a.当前病况, a.住院天数, D3.名称 As 入院科室,
           e.名称 As 护理等级, a.家庭地址, a.户口地址, a.联系人地址
    From 病案主页 A, 医疗付款方式 C, 部门表 D1, 部门表 D2, 部门表 D3, 收费项目目录 E
    Where a.医疗付款方式 = c.名称 And a.当前病区id = D1.Id And a.出院科室id = D2.Id And a.入院科室id = D3.Id And a.护理等级id = e.Id(+) And
          a.病人id = 0 And a.主页id = 0 And Rownum < 1;
  Type Ty_病人信息 Is Ref Cursor;
  c_Pati Ty_病人信息; --动态游标变量

  r_病案信息 c_病案信息%RowType; --全局变量，Setreturnjson 子方法中赋值 ，忽任意使用

  -- c_Pati_0
  --仅按病人id查询
  Cursor c_Pati_0(P病人信息 Varchar2) Is
    Select /*+cardinality(B,10) */
     a.病人id, a.主页id, a.姓名, Null As 性别, Null As 年龄, Null As 住院号, Null As 费别, Null As 病人性质, Null As 审核标志, Null As 住院状态,
     Null As 入院时间, Null As 出院时间, Null As 住院目的, Null As 登记人, Null As 登记时间, Null As 住院医师, Null As 医疗付款方式名称,
     Null As 医疗付款方式编码, Null As 当前床号, Null As 病人类型, Null As 学历, Null As 职业, Null As 国籍, Null As 婚姻状况, Null As 当前病区id,
     Null As 当前病区名称, Null As 当前科室id, Null As 当前科室名称, Null As 险类, Null As 挂号id, Null As 编目日期, Null As 备注, Null As 门诊医师,
     Null As 责任护士, Null As 入院病况, Null As 当前病况, Null As 住院天数, Null As 入院科室, Null As 护理等级, a.家庭地址, a.户口地址, a.联系人地址
    From 病案主页 A, (Select Column_Value As 病人id From Table(f_Str2list(P病人信息))) B
    Where a.病人id = b.病人id And (n_最后一次住院 = 0 Or 主页id = (Select Max(主页id) From 病案主页 Where 病人id = b.病人id)) And
          (v_病人性质 Is Null Or Instr(',' || v_病人性质 || ',', ',' || a.病人性质 || ',') > 0)
    Order By a.主页id Desc;

  -- c_Pati_1  c_Pati_1(P病人信息
  --按病人id和主页id查询
  Cursor c_Pati_1(P病人信息 Varchar2) Is
    Select /*+cardinality(B,10) */
     a.病人id, a.主页id, a.姓名, Null As 性别, Null As 年龄, Null As 住院号, Null As 费别, Null As 病人性质, Null As 审核标志, Null As 住院状态,
     Null As 入院时间, Null As 出院时间, Null As 住院目的, Null As 登记人, Null As 登记时间, Null As 住院医师, Null As 医疗付款方式名称,
     Null As 医疗付款方式编码, Null As 当前床号, Null As 病人类型, Null As 学历, Null As 职业, Null As 国籍, Null As 婚姻状况, Null As 当前病区id,
     Null As 当前病区名称, Null As 当前科室id, Null As 当前科室名称, Null As 险类, Null As 挂号id, Null As 编目日期, Null As 备注, Null As 门诊医师,
     Null As 责任护士, Null As 入院病况, Null As 当前病况, Null As 住院天数, Null As 入院科室, Null As 护理等级, Null As 家庭地址, Null As 户口地址,
     Null As 户口地址
    From 病案主页 A, (Select C1 As 病人id, C2 As 主页id From Table(f_Str2list2(P病人信息, ',', ':'))) B
    Where a.病人id = b.病人id And a.主页id = b.主页id;

  -- c_Pati_0_1  c_Pati_0_1(P病人信息
  --仅按病人id查询
  Cursor c_Pati_0_1(P病人信息 Varchar2) Is
    Select /*+cardinality(B,10) */
     a.病人id, a.主页id, a.姓名, a.性别, a.年龄, a.住院号, a.费别, a.病人性质, a.审核标志, a.状态 As 住院状态,
     To_Char(a.入院日期, 'yyyy-mm-dd hh24:mi:ss') As 入院时间, To_Char(a.出院日期, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.住院目的, a.登记人,
     To_Char(a.登记时间, 'yyyy-mm-dd') As 登记时间, a.住院医师, c.名称 As 医疗付款方式名称, c.编码 医疗付款方式编码, a.出院病床 As 当前床号, a.病人类型, a.学历, a.职业,
     a.国籍, a.婚姻状况, a.当前病区id, D1.名称 As 当前病区名称, a.出院科室id As 当前科室id, D2.名称 As 当前科室名称, a.险类, a.挂号id,
     To_Char(a.编目日期, 'yyyy-mm-dd hh24:mi:ss') As 编目日期, a.备注, Null As 门诊医师, Null As 责任护士, Null As 入院病况, Null As 当前病况,
     Null As 住院天数, Null As 入院科室, Null As 护理等级, a.家庭地址, a.户口地址, a.联系人地址
    From 病案主页 A, (Select Column_Value As 病人id From Table(f_Str2list(P病人信息))) B, 医疗付款方式 C, 部门表 D1, 部门表 D2
    Where a.病人id = b.病人id And (n_最后一次住院 = 0 Or 主页id = (Select Max(主页id) From 病案主页 Where 病人id = b.病人id)) And
          a.医疗付款方式 = c.名称 And a.当前病区id = D1.Id(+) And a.出院科室id = D2.Id(+)
    Order By a.主页id Desc;

  -- c_Pati_1_1  c_Pati_1_1(P病人信息
  --按病人id和主页id查询
  Cursor c_Pati_1_1(P病人信息 Varchar2) Is
    Select /*+cardinality(B,10) */
     a.病人id, a.主页id, a.姓名, a.性别, a.年龄, a.住院号, a.费别, a.病人性质, a.审核标志, a.状态 As 住院状态,
     To_Char(a.入院日期, 'yyyy-mm-dd hh24:mi:ss') As 入院时间, To_Char(a.出院日期, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.住院目的, a.登记人,
     To_Char(a.登记时间, 'yyyy-mm-dd') As 登记时间, a.住院医师, c.名称 As 医疗付款方式名称, c.编码 医疗付款方式编码, a.出院病床 As 当前床号, a.病人类型, a.学历, a.职业,
     a.国籍, a.婚姻状况, a.当前病区id, D1.名称 As 当前病区名称, a.出院科室id As 当前科室id, D2.名称 As 当前科室名称, a.险类, a.挂号id,
     To_Char(a.编目日期, 'yyyy-mm-dd hh24:mi:ss') As 编目日期, a.备注, a.门诊医师, a.责任护士, a.入院病况, a.当前病况, a.住院天数, D3.名称 As 入院科室,
     e.名称 As 护理等级, a.家庭地址, a.户口地址, a.联系人地址
    From 病案主页 A, (Select C1 As 病人id, C2 As 主页id From Table(f_Str2list2(P病人信息, ',', ':'))) B, 医疗付款方式 C, 部门表 D1, 部门表 D2,
         部门表 D3, 收费项目目录 E
    Where a.病人id = b.病人id And a.主页id = b.主页id And a.医疗付款方式 = c.名称 And a.当前病区id = D1.Id(+) And a.出院科室id = D2.Id(+) And
          a.入院科室id = D3.Id(+) And a.护理等级id = e.Id(+);

  --c_Pati_0_0  c_Pati_0_0(P病人信息
  --仅按病人id查询
  Cursor c_Pati_0_0(P病人信息 Varchar2) Is
    Select /*+cardinality(B,10) */
     a.病人id, a.主页id, a.姓名, a.性别, a.年龄, a.住院号, a.费别, a.病人性质, a.审核标志, a.状态 As 住院状态,
     To_Char(a.入院日期, 'yyyy-mm-dd hh24:mi:ss') As 入院时间, To_Char(a.出院日期, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.住院目的, a.登记人,
     To_Char(a.登记时间, 'yyyy-mm-dd') As 登记时间, a.住院医师, Null As 医疗付款方式名称, Null As 医疗付款方式编码, Null As 当前床号, Null As 病人类型,
     Null As 学历, Null As 职业, Null As 国籍, Null As 婚姻状况, Null As 当前病区id, Null As 当前病区名称, Null As 当前科室id, Null As 当前科室名称,
     Null As 险类, Null As 挂号id, Null As 编目日期, Null As 备注, Null As 门诊医师, Null As 责任护士, Null As 入院病况, Null As 当前病况,
     Null As 住院天数, Null As 入院科室, Null As 护理等级, a.家庭地址, a.户口地址, a.联系人地址
    From 病案主页 A, (Select Column_Value As 病人id From Table(f_Str2list(P病人信息))) B
    Where a.病人id = b.病人id And (n_最后一次住院 = 0 Or 主页id = (Select Max(主页id) From 病案主页 Where 病人id = b.病人id))
    Order By a.主页id Desc;

  --c_Pati_1_0  c_Pati_1_0(P病人信息
  --按病人id和主页id查询
  Cursor c_Pati_1_0(P病人信息 Varchar2) Is
    Select /*+cardinality(B,10) */
     a.病人id, a.主页id, a.姓名, a.性别, a.年龄, a.住院号, a.费别, a.病人性质, a.审核标志, a.状态 As 住院状态,
     To_Char(a.入院日期, 'yyyy-mm-dd hh24:mi:ss') As 入院时间, To_Char(a.出院日期, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.住院目的, a.登记人,
     To_Char(a.登记时间, 'yyyy-mm-dd') As 登记时间, a.住院医师, Null As 医疗付款方式名称, Null As 医疗付款方式编码, Null As 当前床号, Null As 病人类型,
     Null As 学历, Null As 职业, Null As 国籍, Null As 婚姻状况, Null As 当前病区id, Null As 当前病区名称, Null As 当前科室id, Null As 当前科室名称,
     Null As 险类, Null As 挂号id, Null As 编目日期, Null As 备注, Null As 门诊医师, Null As 责任护士, Null As 入院病况, Null As 当前病况,
     Null As 住院天数, Null As 入院科室, Null As 护理等级, a.家庭地址, a.户口地址, a.联系人地址
    From 病案主页 A, (Select C1 As 病人id, C2 As 主页id From Table(f_Str2list2(P病人信息, ',', ':'))) B
    Where a.病人id = b.病人id And a.主页id = b.主页id;

  Procedure Setreturnjson
  (
    n_是否婴儿   Number,
    n_转科       Number,
    n_包床位信息 Number
  ) Is
    --功能：根据记录集信息拼接json串
    --      游标一行记录调一次
    v_List_Baby Varchar2(32767);
    v_List_Tran Varchar2(32767);
    v_List_Bad  Varchar2(32767);
    v_医保号    病案主页从表.信息值%Type;
    n_颜色      病人类型.颜色%Type;
  Begin
    --主页信息
    Zljsonputvalue(v_List, 'pati_id', r_病案信息.病人id, 1, 1);
    Zljsonputvalue(v_List, 'pati_pageid', r_病案信息.主页id);
    Zljsonputvalue(v_List, 'pati_name', r_病案信息.姓名);
  
    --基本信息
    If n_类型 < 2 Then
      Zljsonputvalue(v_List, 'pati_sex', r_病案信息.性别);
      Zljsonputvalue(v_List, 'pati_age', r_病案信息.年龄);
      Zljsonputvalue(v_List, 'inpatient_num', r_病案信息.住院号, 0);
      Zljsonputvalue(v_List, 'fee_category', r_病案信息.费别);
      Zljsonputvalue(v_List, 'pati_nature', r_病案信息.病人性质, 1);
      Zljsonputvalue(v_List, 'audit_sign', r_病案信息.审核标志, 1);
      Zljsonputvalue(v_List, 'si_inp_status', r_病案信息.住院状态, 1);
      Zljsonputvalue(v_List, 'adta_time', r_病案信息.入院时间);
      Zljsonputvalue(v_List, 'adtd_time', r_病案信息.出院时间);
      Zljsonputvalue(v_List, 'in_objective', r_病案信息.住院目的);
      Zljsonputvalue(v_List, 'reg_name', r_病案信息.登记人);
      Zljsonputvalue(v_List, 'reg_date', r_病案信息.登记时间);
      Zljsonputvalue(v_List, 'pat_rsdpscn', r_病案信息.住院医师);
      Zljsonputvalue(v_List, 'pati_home_addr', r_病案信息.家庭地址);
      Zljsonputvalue(v_List, 'pati_house_addr', r_病案信息.户口地址);
      Zljsonputvalue(v_List, 'pati_contact_addr', r_病案信息.联系人地址);
    End If;
  
    --扩展信息
    If n_类型 = 1 Then
      --      mdlpay_mode_name    C 1 医疗付款方式名称
      --      mdlpay_mode_code    C 1 医疗付款方式编码
      Zljsonputvalue(v_List, 'mdlpay_mode_name', r_病案信息.医疗付款方式名称);
      Zljsonputvalue(v_List, 'mdlpay_mode_code', r_病案信息.医疗付款方式编码);
      --      pati_bed            C 1 当前床号
      --      pati_type           C 1 病人类型(普通，医保，留观)
      --      pati_show_color     N 1 病人显示颜色
      --      pati_education      C 1 学历
      --      ocpt_name           C 1 职业
      --      country_name        C 1 国籍
      --      pati_marital_cstatus  C 1 婚姻状况
      Zljsonputvalue(v_List, 'pati_bed', r_病案信息.当前床号);
      Zljsonputvalue(v_List, 'pati_type', r_病案信息.病人类型);
    
      If r_病案信息.病人类型 Is Not Null Then
        Select Max(颜色) Into n_颜色 From 病人类型 Where 名称 = Nvl(r_病案信息.病人类型, '');
      End If;
      Zljsonputvalue(v_List, 'pati_show_color', n_颜色, 1);
      Zljsonputvalue(v_List, 'pati_education', r_病案信息.学历);
      Zljsonputvalue(v_List, 'ocpt_name', r_病案信息.职业);
      Zljsonputvalue(v_List, 'country_name', r_病案信息.国籍);
      Zljsonputvalue(v_List, 'pati_marital_cstatus', r_病案信息.婚姻状况);
      --      pati_wardarea_id    N 1 当前病区id
      --      pati_wardarea_name  C 1 当前病区名称
      --      pati_dept_id        N 1 当前科室id
      --      pati_dept_name      C 1 当前科室名称
      Zljsonputvalue(v_List, 'pati_wardarea_id', r_病案信息.当前病区id, 1);
      Zljsonputvalue(v_List, 'pati_wardarea_name', r_病案信息.当前病区名称);
      Zljsonputvalue(v_List, 'pati_dept_id', r_病案信息.当前科室id, 1);
      Zljsonputvalue(v_List, 'pati_dept_name', r_病案信息.当前科室名称);
      --      insurance_type      N 1 险类
      --      rgst_id             N 1 挂号id
      --      catalog date        C 1 编目日期:yyyy-mm-dd hh24:mi:ss
      Zljsonputvalue(v_List, 'insurance_type', r_病案信息.险类, 1);
      Zljsonputvalue(v_List, 'rgst_id', r_病案信息.挂号id, 1);
      Zljsonputvalue(v_List, 'catalog_date', r_病案信息.编目日期);
      Zljsonputvalue(v_List, 'pati_desc', r_病案信息.备注);
      --      outpatient_doctor   C 1 门诊医师
      --      responsible_nurse   C 1 责任护士
      --      hospital_admissions C 1 入院病况
      --      current_conditions  C 1 当前病况
      --      hospital_days       N 1 住院天数
      --      hospital_dept       C 1 入院科室
      --      level_of_care       C 1 护理等级
      Zljsonputvalue(v_List, 'outpatient_doctor', r_病案信息.门诊医师);
      Zljsonputvalue(v_List, 'responsible_nurse', r_病案信息.责任护士);
      Zljsonputvalue(v_List, 'hospital_admissions', r_病案信息.入院病况);
      Zljsonputvalue(v_List, 'current_conditions', r_病案信息.当前病况);
      Zljsonputvalue(v_List, 'hospital_days', r_病案信息.住院天数, 1);
      Zljsonputvalue(v_List, 'hospital_dept', r_病案信息.入院科室);
      Zljsonputvalue(v_List, 'level_of_care', r_病案信息.护理等级);
    
      Select Max(Decode(a.信息名, '医保号', a.信息值, '')) As 医保号
      Into v_医保号
      From 病案主页从表 A
      Where a.病人id = r_病案信息.病人id And a.主页id = r_病案信息.主页id And a.信息名 In ('医保号');
    
      Zljsonputvalue(v_List, 'insurance_num', v_医保号);
    
      For r_等级 In (Select b.名称 As 床位等级
                   From 病人变动记录 A, 收费项目目录 B
                   Where r_病案信息.病人id = a.病人id And r_病案信息.主页id = a.主页id And
                         ((a.终止时间 Is Null And 附加床位 = 0) Or (a.终止时间 Is Not Null And 终止原因 = 1)) And a.床位等级id = b.Id(+)) Loop
        Zljsonputvalue(v_List, 'level_of_bed', r_等级.床位等级);
      End Loop;
      Select Count(1)
      Into n_已入科
      From 病人变动记录
      Where 病人id = r_病案信息.病人id And 主页id = r_病案信息.主页id And 开始原因 = 2 And 床号 Is Not Null And Rownum < 2;
      Zljsonputvalue(v_List, 'in_dept', n_已入科, 1);
    End If;
  
    If n_是否婴儿 = 1 Then
      For r_婴儿信息 In (Select 病人id, 主页id, 序号 As 婴儿序号, 婴儿姓名, 婴儿性别, To_Char(出生时间, 'yyyy-mm-dd hh24:mi:ss') As 出生时间
                     From 病人新生儿记录
                     Where 病人id = r_病案信息.病人id And 主页id = r_病案信息.主页id) Loop
        --      baby_list[]           1 婴儿信息，[数组]
        --        pati_id           N 1 病人id
        --        pati_pageid       N 1 主页id
        --        baby_num          N 1 婴儿序号
        --        baby_name         C 1 婴儿姓名
        --        baby_sex          C 1 婴儿性别
        --        baby_date         C 1 出生时间
        v_List_Baby := v_List_Baby || ',{';
        v_List_Baby := v_List_Baby || '"pati_id":' || r_婴儿信息.病人id || ',';
        v_List_Baby := v_List_Baby || '"pati_pageid":' || r_婴儿信息.主页id || ',';
        v_List_Baby := v_List_Baby || '"baby_num":' || r_婴儿信息.婴儿序号 || ',';
        v_List_Baby := v_List_Baby || '"baby_name":"' || Zljsonstr(r_婴儿信息.婴儿姓名) || '",';
        v_List_Baby := v_List_Baby || '"baby_sex":"' || r_婴儿信息.婴儿性别 || '",';
        v_List_Baby := v_List_Baby || '"baby_date":"' || r_婴儿信息.出生时间 || '"';
        v_List_Baby := v_List_Baby || '}';
      End Loop;
      v_List := v_List || ',"baby_list":[' || Substr(v_List_Baby, 2) || ']';
    End If;
  
    If n_转科 = 1 Then
      For r_转科 In (Select a.开始原因, To_Char(a.开始时间, 'yyyy-mm-dd hh24:mi:ss') As 开始时间, b.名称
                   From 病人变动记录 A, 部门表 B
                   Where a.科室id = b.Id And a.开始时间 Is Not Null And a.开始原因 = 3 And a.病人id = r_病案信息.病人id And
                         主页id = r_病案信息.主页id) Loop
      
        --      trans_list[]        C   转科列表信息
        --        start_reason      C 1 开始原因
        --        start_time        C 1 开始时间:yyyy-mm-dd hh24:mi:ss
        --        dept_name         C 1 科室名称
        v_List_Tran := v_List_Tran || ',{';
        v_List_Tran := v_List_Tran || '"start_reason":"' || r_转科.开始原因 || '",';
        v_List_Tran := v_List_Tran || '"start_time":"' || r_转科.开始时间 || '",';
        v_List_Tran := v_List_Tran || '"dept_name":"' || Zljsonstr(r_转科.名称) || '"';
        v_List_Tran := v_List_Tran || '}';
      End Loop;
      v_List := v_List || ',"trans_list":[' || Substr(v_List_Tran, 2) || ']';
    End If;
  
    If Nvl(n_包床位信息, 0) = 1 Then
      For r_床位 In (Select a.病区id, c.名称 As 病区名称, a.床号, b.编码, b.名称
                   From 床位状况记录 A, 床位编制分类 B, 部门表 C
                   Where a.床位编制 = b.名称(+) And a.病区id = c.Id And a.病人id = r_病案信息.病人id) Loop
      
        --      badinfo_list[]      床位信息，[数组]
        --        wardarea_id       N 1 病区id
        --        wardarea_name     C 1 病区名称
        --        bed_no            C 1 床号
        --        bed_class_code    C 1 分类编码
        --        bed_class_name    C 1 分类名称
        v_List_Bad := v_List_Bad || ',{';
        v_List_Bad := v_List_Bad || '"wardarea_id":' || r_床位.病区id || ',';
        v_List_Bad := v_List_Bad || '"wardarea_name":"' || Zljsonstr(r_床位.病区名称) || '",';
        v_List_Bad := v_List_Bad || '"bed_no":"' || Zljsonstr(r_床位.床号) || '",';
        v_List_Bad := v_List_Bad || '"bed_class_code":"' || Zljsonstr(r_床位.编码) || '",';
        v_List_Bad := v_List_Bad || '"bed_class_name":"' || Zljsonstr(r_床位.名称) || '"';
        v_List_Bad := v_List_Bad || '}';
      End Loop;
      v_List := v_List || ',"badinfo_list":[' || Substr(v_List_Bad, 2) || ']';
    End If;
    v_List := v_List || '}';
  
    If Length(v_List) > 20000 Then
      Json_Out_Tmp := Json_Out_Tmp || v_List;
      v_List       := ',';
    End If;
  End;
Begin
  --解析入参
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');
  --      query_type          C 1 查询类型:0-基本信息;1-基本信息的展
  --      pati_pageids        C 1 病人信息,格式两种:一种是:病人id:主页ID,…;一种：病人id,…
  --      is_babyinfo         N 1 是否包含婴儿信息:1-包含;0-不包含
  --      is_transdeptinfo    N 1 是否包含转科信息:1-包含;0-不包含
  --      is_lastpage         N 1 是否取最后一次住院
  --      pati_natures        C 1 病人性质:0-普通住院病人,1-门诊留观病人,2-住院留观病人；多个逗号分隔，不传为所有
  --      rgst_id             N 1 挂号ID,根据挂号ID查询
  --      is_badinfo          N 1 是否包含床位信息:1-包含;0-不包含
  n_类型         := Nvl(j_Json.Get_Number('query_type'), 0);
  v_病人信息     := j_Json.Get_String('pati_pageids');
  n_包含婴儿信息 := Nvl(j_Json.Get_Number('is_babyinfo'), 0);
  n_包含转科信息 := Nvl(j_Json.Get_Number('is_transdeptinfo'), 0);
  n_最后一次住院 := Nvl(j_Json.Get_Number('is_lastpage'), 0);
  v_病人性质     := j_Json.Get_String('pati_natures');
  n_挂号id       := Nvl(j_Json.Get_Number('rgst_id'), 0);
  n_包含床位信息 := Nvl(j_Json.Get_Number('is_badinfo'), 0);

  If Nvl(v_病人信息, '-') = '-' Then
    Json_Out := Zljsonout('未传入病人信息，请检查！');
    Return;
  End If;

  n_查询方式 := 0;
  If Instr(v_病人信息, ':') > 0 Then
    n_查询方式 := 1;
  End If;
  If Nvl(n_挂号id, 0) <> 0 Then
    n_查询方式 := 2;
  End If;

  --将 v_病人信息 串组装成不超过4000 的集合串，防止使用动态内存表查询时参数超长
  I := 0;
  While v_病人信息 Is Not Null Loop
    If Length(v_病人信息) <= 4000 Then
      Col_病人信息(I) := v_病人信息;
      v_病人信息 := Null;
    Else
      Col_病人信息(I) := Substr(v_病人信息, 1, Instr(v_病人信息, ',', 3980) - 1);
      v_病人信息 := Substr(v_病人信息, Instr(v_病人信息, ',', 3980) + 1);
    End If;
    I := I + 1;
  End Loop;

  --出参开始符串拼接
  Json_Out_Tmp := '{"output":{"code":1,"message":"成功","page_list":[';

  If n_类型 = 2 Then
    --仅查询主页
    If n_查询方式 = 0 Then
      -- c_Pati_0 
      --仅按病人id查询      
      For K In 0 .. Col_病人信息.Count - 1 Loop
        Open c_Pati_0(Col_病人信息(K));
        Loop
          Fetch c_Pati_0
            Into r_病案信息;
          Exit When c_Pati_0%NotFound;
          Setreturnjson(n_包含婴儿信息, n_包含转科信息, n_包含床位信息);
        End Loop;
        Close c_Pati_0;
      End Loop;
    Elsif n_查询方式 = 1 Then
      -- c_Pati_1  
      --按病人id和主页id查询
      For K In 0 .. Col_病人信息.Count - 1 Loop
        Open c_Pati_1(Col_病人信息(K));
        Loop
          Fetch c_Pati_1
            Into r_病案信息;
          Exit When c_Pati_1%NotFound;
          Setreturnjson(n_包含婴儿信息, n_包含转科信息, n_包含床位信息);
        End Loop;
        Close c_Pati_1;
      End Loop;
    Elsif n_查询方式 = 2 Then
      --按挂号ID查询
      Open c_Pati For
        Select a.病人id, a.主页id, a.姓名, Null As 性别, Null As 年龄, Null As 住院号, Null As 费别, Null As 病人性质, Null As 审核标志,
               Null As 住院状态, Null As 入院时间, Null As 出院时间, Null As 住院目的, Null As 登记人, Null As 登记时间, Null As 住院医师,
               Null As 医疗付款方式名称, Null As 医疗付款方式编码, Null As 当前床号, Null As 病人类型, Null As 学历, Null As 职业, Null As 国籍,
               Null As 婚姻状况, Null As 当前病区id, Null As 当前病区名称, Null As 当前科室id, Null As 当前科室名称, Null As 险类, Null As 挂号id,
               Null As 编目日期, Null As 备注, Null As 门诊医师, Null As 责任护士, Null As 入院病况, Null As 当前病况, Null As 住院天数,
               Null As 入院科室, Null As 护理等级, Null As 家庭地址, Null As 户口地址, Null As 户口地址
        From 病案主页 A
        Where a.挂号id = n_挂号id;
      Loop
        Fetch c_Pati
          Into r_病案信息;
        Exit When c_Pati%NotFound;
        Setreturnjson(n_包含婴儿信息, n_包含转科信息, n_包含床位信息);
      End Loop;
    End If;
  Elsif n_类型 = 1 Then
    --查询基本信息+扩展信息
    If n_查询方式 = 0 Then
      -- c_Pati_0_1  
      --仅按病人id查询
      For K In 0 .. Col_病人信息.Count - 1 Loop
        Open c_Pati_0_1(Col_病人信息(K));
        Loop
          Fetch c_Pati_0_1
            Into r_病案信息;
          Exit When c_Pati_0_1%NotFound;
          Setreturnjson(n_包含婴儿信息, n_包含转科信息, n_包含床位信息);
        End Loop;
        Close c_Pati_0_1;
      End Loop;
    Elsif n_查询方式 = 1 Then
      -- c_Pati_1_1  
      --按病人id和主页id查询
      For K In 0 .. Col_病人信息.Count - 1 Loop
        Open c_Pati_1_1(Col_病人信息(K));
        Loop
          Fetch c_Pati_1_1
            Into r_病案信息;
          Exit When c_Pati_1_1%NotFound;
          Setreturnjson(n_包含婴儿信息, n_包含转科信息, n_包含床位信息);
        End Loop;
        Close c_Pati_1_1;
      End Loop;
    Elsif n_查询方式 = 2 Then
      --按挂号ID查询
      Open c_Pati For
        Select a.病人id, a.主页id, a.姓名, a.性别, a.年龄, a.住院号, a.费别, a.病人性质, a.审核标志, a.状态 As 住院状态,
               To_Char(a.入院日期, 'yyyy-mm-dd hh24:mi:ss') As 入院时间, To_Char(a.出院日期, 'yyyy-mm-dd hh24:mi:ss') As 出院时间,
               a.住院目的, a.登记人, To_Char(a.登记时间, 'yyyy-mm-dd') As 登记时间, a.住院医师, c.名称 As 医疗付款方式名称, c.编码 医疗付款方式编码,
               a.出院病床 As 当前床号, a.病人类型, a.学历, a.职业, a.国籍, a.婚姻状况, a.当前病区id, D1.名称 As 当前病区名称, a.出院科室id As 当前科室id,
               D2.名称 As 当前科室名称, a.险类, a.挂号id, To_Char(a.编目日期, 'yyyy-mm-dd hh24:mi:ss') As 编目日期, a.备注, a.门诊医师, a.责任护士,
               a.入院病况, a.当前病况, a.住院天数, D3.名称 As 入院科室, e.名称 As 护理等级, a.家庭地址, a.户口地址, a.联系人地址
        From 病案主页 A, 医疗付款方式 C, 部门表 D1, 部门表 D2, 部门表 D3, 收费项目目录 E
        Where a.挂号id = n_挂号id And a.医疗付款方式 = c.名称 And a.当前病区id = D1.Id(+) And a.出院科室id = D2.Id(+) And
              a.入院科室id = D3.Id(+) And a.护理等级id = e.Id(+);
    
      Loop
        Fetch c_Pati
          Into r_病案信息;
        Exit When c_Pati%NotFound;
        Setreturnjson(n_包含婴儿信息, n_包含转科信息, n_包含床位信息);
      End Loop;
    End If;
  Else
    --只查询基本信息
    If n_查询方式 = 0 Then
      --c_Pati_0_0 
      --仅按病人id查询
      For K In 0 .. Col_病人信息.Count - 1 Loop
        Open c_Pati_0_0(Col_病人信息(K));
        Loop
          Fetch c_Pati_0_0
            Into r_病案信息;
          Exit When c_Pati_0_0%NotFound;
          Setreturnjson(n_包含婴儿信息, n_包含转科信息, n_包含床位信息);
        End Loop;
        Close c_Pati_0_0;
      End Loop;
    Elsif n_查询方式 = 1 Then
      --c_Pati_1_0 
      --按病人id和主页id查询
      For K In 0 .. Col_病人信息.Count - 1 Loop
        Open c_Pati_1_0(Col_病人信息(K));
        Loop
          Fetch c_Pati_1_0
            Into r_病案信息;
          Exit When c_Pati_1_0%NotFound;
          Setreturnjson(n_包含婴儿信息, n_包含转科信息, n_包含床位信息);
        End Loop;
        Close c_Pati_1_0;
      End Loop;
    Elsif n_查询方式 = 2 Then
      --按挂号ID查询
      Open c_Pati For
        Select /*+cardinality(B,10) */
         a.病人id, a.主页id, a.姓名, a.性别, a.年龄, a.住院号, a.费别, a.病人性质, a.审核标志, a.状态 As 住院状态,
         To_Char(a.入院日期, 'yyyy-mm-dd hh24:mi:ss') As 入院时间, To_Char(a.出院日期, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.住院目的,
         a.登记人, To_Char(a.登记时间, 'yyyy-mm-dd') As 登记时间, a.住院医师, Null As 医疗付款方式名称, Null As 医疗付款方式编码, Null As 当前床号,
         Null As 病人类型, Null As 学历, Null As 职业, Null As 国籍, Null As 婚姻状况, Null As 当前病区id, Null As 当前病区名称, Null As 当前科室id,
         Null As 当前科室名称, Null As 险类, Null As 挂号id, Null As 编目日期, Null As 备注, Null As 门诊医师, Null As 责任护士, Null As 入院病况,
         Null As 当前病况, Null As 住院天数, Null As 入院科室, Null As 护理等级, a.家庭地址, a.户口地址, a.联系人地址
        From 病案主页 A
        Where a.挂号id = n_挂号id;
      Loop
        Fetch c_Pati
          Into r_病案信息;
        Exit When c_Pati%NotFound;
        Setreturnjson(n_包含婴儿信息, n_包含转科信息, n_包含床位信息);
      End Loop;
    End If;
  End If;

  --出参结束符串拼接
  If v_List <> ',' Then
    Json_Out_Tmp := Json_Out_Tmp || v_List || ']}}';
  Else
    Json_Out_Tmp := Json_Out_Tmp || ']}}';
  End If;

  Json_Out := Json_Out_Tmp;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getpatipageinfo;
/
Create Or Replace Procedure Zl_Cissvr_Getpativisitid
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --功能：获取病人最后一次挂号的挂号记录
  --入参：Json_In:格式
  --input
  --  query_type        N    1 查询方式  0-按病人ID查询最近一次的就诊ID(门诊或住院) 1-获取每次住院的主页ID
  --  pati_id           N    1 病人id
  --  occasion          N    1 场合：0-不区分，1-门诊，2-住院
  --出参      json
  --output
  --    code                N   1 应答吗：0-失败；1-成功
  --    message             C   1 应答消息：失败时返回具体的错误信息
  --    visit_id            C  1 就诊id
  --    occasion            N  1 场合
  --    visit_list[]           query_type=1 时返回主页id和病人性质
  --    visit_id             N  1 病人id
  --    pati_type           N  1 病人性质
  -------------------------------------------
  j_In Pljson;

  j_Json Pljson;
  v_List Varchar2(32767);

  n_病人id Number(18);
  n_就诊id Number;
  n_场合   Number;
  n_主页id Number;
  n_Type   Number;
  n_Id     Number;
Begin
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_Type   := j_Json.Get_Number('query_type');
  n_病人id := j_Json.Get_Number('pati_id');
  n_场合   := j_Json.Get_Number('occasion');
  Json_Out := '{"output":{"code":1,"message":"成功"';
  If Nvl(n_病人id, 0) <> 0 Then
    If Nvl(n_Type, 0) = 0 Then
      If Nvl(n_场合, 0) = 0 Then
        Begin
          Select ID
          Into n_就诊id
          From (Select ID From 病人挂号记录 Where 病人id = n_病人id And Mod(记录状态, 2) <> 0 Order By 登记时间 Desc)
          Where Rownum < 2;
        Exception
          When Others Then
            Null;
        End;
        Begin
          Select Max(a.主页id)
          Into n_主页id
          From 病案主页 A, 在院病人 B
          Where a.病人id = n_病人id And Nvl(a.主页id, 0) <> 0 And a.病人id = b.病人id And a.主页id = b.主页id;
        Exception
          When Others Then
            Null;
        End;
        If Nvl(n_主页id, 0) = 0 Then
          n_Id   := Nvl(n_就诊id, 0);
          n_场合 := 1;
        Else
          n_Id   := n_主页id;
          n_场合 := 2;
        End If;
        Json_Out := Json_Out || ',"visit_id":' || n_Id || ',"occasion":' || n_场合 || '}}';
      Elsif Nvl(n_场合, 0) = 1 Then
        Begin
          Select ID
          Into n_就诊id
          From (Select ID From 病人挂号记录 Where 病人id = n_病人id And Mod(记录状态, 2) <> 0 Order By 登记时间 Desc)
          Where Rownum < 2;
        Exception
          When Others Then
            Null;
        End;
        n_Id     := Nvl(n_就诊id, 0);
        n_场合   := 1;
        Json_Out := Json_Out || ',"visit_id":' || n_Id || ',"occasion":' || n_场合 || '}}';
      Elsif Nvl(n_场合, 0) = 2 Then
        Begin
          Select Max(a.主页id)
          Into n_主页id
          From 病案主页 A, 在院病人 B
          Where a.病人id = n_病人id And Nvl(a.主页id, 0) <> 0 And a.病人id = b.病人id And a.主页id = b.主页id;
        Exception
          When Others Then
            Null;
        End;
        n_Id     := Nvl(n_主页id, 0);
        n_场合   := 2;
        Json_Out := Json_Out || ',"visit_id":' || n_Id || ',"occasion":' || n_场合 || '}}';
      End If;
    Elsif Nvl(n_Type, 0) = 1 Then
      For R In (Select 主页id, 病人性质 From 病案主页 Where 病人id = n_病人id And Nvl(主页id, 0) <> 0 Order By 主页id) Loop
        Zljsonputvalue(v_List, 'visit_id', r.主页id, 1, 1);
        Zljsonputvalue(v_List, 'pati_type', r.病人性质, 1, 2);
      End Loop;
      Json_Out := Json_Out || ',visit_list:[' || v_List || ']}}';
    End If;
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getpativisitid;
/
Create Or Replace Procedure Zl_Cissvr_Getpativisitinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  -------------------------------------------
  --功能：获取病人的就诊记录
  --入参：Json_In:格式
  --input
  --  pati_id                 N    1 病人id
  --出参      json
  --output
  --    code                N   1 应答吗：0-失败；1-成功
  --    message             C   1 应答消息：失败时返回具体的错误信息
  --    visit_list[]
  --      visit_id             N  1 就诊id
  --      pati_nature          N  1 病人性质
  --      occasion             N  1 场合 1-门诊 2-住院
  --      regist_no            C  1 挂号单
  --      create_time          C  1 登记时间
  --      pati_type            C  1 病人类型
  --      insurance_type       C  1 险类
  --      adta_time            C  1 入院时间
  -------------------------------------------
  j_In   Pljson;
  j_Json Pljson;
  v_List Varchar2(32767);

  n_病人id Number(18);
Begin
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');
  Json_Out := '{"output":{"code":1,"message":"成功","visit_list":[';
  For R In (Select *
            From (Select 1 性质, ID ID, NO, 0 病人性质, To_Char(登记时间, 'YYYY-MM-DD hh24:mi:ss') 登记时间, Null 病人类型, Null 险类,
                          Null As 入院日期
                   From 病人挂号记录
                   Where 病人id = n_病人id And Mod(记录状态, 2) <> 0
                   Union All
                   Select 2 性质, 主页id ID, '' || 主页id NO, 病人性质, To_Char(登记时间, 'YYYY-MM-DD hh24:mi:ss') 登记时间, 病人类型, 险类, 入院日期
                   From 病案主页
                   Where 病人id = n_病人id And Nvl(主页id, 0) <> 0)
            Order By NO Desc) Loop
    Zljsonputvalue(v_List, 'visit_id', r.Id, 1, 1);
    Zljsonputvalue(v_List, 'occasion', r.性质, 1);
    Zljsonputvalue(v_List, 'pati_nature', r.病人性质, 1);
    Zljsonputvalue(v_List, 'regist_no', r.No);
    Zljsonputvalue(v_List, 'create_time', r.登记时间);
    Zljsonputvalue(v_List, 'pati_type', r.病人类型);
    Zljsonputvalue(v_List, 'insurance_type', r.险类, 1);
    Zljsonputvalue(v_List, 'adta_time', r.入院日期, 0, 2);
    If Length(v_List) > 20000 Then
      Json_Out := Json_Out || v_List;
      v_List   := ',';
    End If;
  End Loop;
  If v_List <> ',' Then
    Json_Out := Json_Out || v_List || ']}}';
  Else
    Json_Out := Json_Out || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getpativisitinfo;
/
Create Or Replace Procedure Zl_Cissvr_Getpativitalsigns
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：获取病人生命体征信息
  --input      获取病人生命体征信息
  --  pati_id               N  1  病人ID
  --  visit_id              N  1  就诊id ，门诊病人为挂号ID;住院病人为主页ID;
  --  outpati_flag          N    门诊标志：1-门诊，2-住院
  --output
  --  code                  C  1  应答码：0-失败；1-成功
  --  message               C  1  应答消息：
  --  pativital_list[]      体征信息，包括项目，数值，单位。[数组]
  --     pativital_item     C  1  项目
  --     pativital_value    C  1  值
  --     pativital_unit     C  1  单位
  ---------------------------------------------------------------------------
  j_Input    Pljson;
  n_病人id   Number(18);
  n_就诊id   Number(18);
  n_门诊标志 Number(1) := 1; --1-门诊，2-住院

  v_体重 Varchar2(100);
  v_身高 Varchar2(100);

  j_In     Pljson;
  v_Output Varchar2(32767);
Begin
  j_In       := Pljson(Json_In);
  j_Input    := j_In.Get_Pljson('input');
  n_病人id   := j_Input.Get_Number('pati_id');
  n_就诊id   := j_Input.Get_Number('visit_id');
  n_门诊标志 := j_Input.Get_Number('outpati_flag');

  If n_门诊标志 = 1 Then
    For R In (Select a.中文名, a.中文名 || '<;>' || b.值 || '<;>' || a.单位 As 值, a.中文名 As 项目, b.值 As 项目值, a.单位
              From (Select 中文名, 单位
                     From 诊治所见项目
                     Where 分类id = 7 And 中文名 In ('体温', '脉搏', '收缩压', '舒张压', '体重', '身高', '呼吸', '血糖')) A,
                   (Select b.项目单位, b.项目名称 As 中文名, b.记录内容 As 值
                     From 病人护理记录 A, 病人护理内容 B
                     Where a.Id = b.记录id And a.病人id = n_病人id And a.主页id = n_就诊id And
                           b.项目名称 In ('体温', '脉搏', '收缩压', '舒张压', '体重', '身高', '呼吸', '血糖')) B
              Where a.中文名 = b.中文名(+)) Loop
    
      Zljsonputvalue(v_Output, 'pativital_item', Nvl(r.项目, ''), 0, 1);
      Zljsonputvalue(v_Output, 'pativital_value', Nvl(r.项目值, ''));
      Zljsonputvalue(v_Output, 'pativital_unit', Nvl(r.单位, ''), 0, 2);
    End Loop;
  Else
    --取各个项目最近一次的记录
    For R In (Select a.中文名, a.中文名 || '<;>' || b.值 || '<;>' || a.单位 As 值, b.值 As 记录内容, a.中文名 As 项目, b.值 As 项目值, a.单位
              From (Select 中文名, 单位
                     From 诊治所见项目
                     Where 分类id = 7 And 中文名 In ('体温', '脉搏', '收缩压', '舒张压', '体重', '身高', '呼吸', '血糖')) A,
                   (Select 项目名称 As 中文名, 记录内容 As 值
                     From (Select 项目名称, 记录内容, Row_Number() Over(Partition By 项目名称 Order By 记录时间 Desc) Rn
                            From (Select c.项目名称, c.记录内容, c.记录时间
                                   From 病人护理文件 A, 病人护理数据 B, 病人护理明细 C
                                   Where a.病人id = n_病人id And a.主页id = n_就诊id And Nvl(a.婴儿, 0) = 0 And a.Id = b.文件id And
                                         b.Id = c.记录id
                                   Union All
                                   Select b.项目名称, b.记录内容, b.修改时间 As 记录时间
                                   From 病人护理记录 A, 病人护理内容 B
                                   Where a.病人id = n_病人id And a.主页id = n_就诊id And Nvl(a.婴儿, 0) = 0 And b.记录类型 = 1 And
                                         a.Id = b.记录id))
                     Where Rn = 1) B
              Where a.中文名 = b.中文名(+)) Loop
    
      If r.中文名 = '体重' Then
        If r.记录内容 Is Null Then
          v_体重 := '';
        Else
          v_体重 := r.值;
        End If;
      Elsif r.中文名 = '身高' Then
        If r.记录内容 Is Null Then
          v_身高 := '';
        Else
          v_身高 := r.值;
        End If;
      End If;
    
      Zljsonputvalue(v_Output, 'pativital_item', Nvl(r.项目, ''), 0, 1);
      Zljsonputvalue(v_Output, 'pativital_value', Nvl(r.项目值, ''));
      Zljsonputvalue(v_Output, 'pativital_unit', Nvl(r.单位, ''), 0, 2);
    End Loop;
  End If;

  Json_Out := '{"output":{"code":1,"message": "成功","pativital_list":[' || v_Output || ']}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getpativitalsigns;
/
Create Or Replace Procedure Zl_Cissvr_Getpatpageinfbyrange
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:获取病案信息
  --入参：Json_In:格式
  --  input
  --    query_type          N 1 查询类型:0-基本;1-基本扩展
  --    wararea_ids         C   病区ids:多个用逗号分隔
  --    dept_ids            C   科室IDs:多个用逗号分隔
  --    pati_ids            C   病人ids:多个用逗号分隔
  --    pati_pageIds        C   主页IDs:病人id:主页id,…
  --    adta_start_time     C   入院开始时间:yyyy-mm-dd hh24:mi:ss
  --    adta_end_time       C   入院结束时间:yyyy-mm-dd hh24:mi:ss
  --    adtd_start_time     C   出院开始时间:yyyy-mm-dd hh24:mi:ss
  --    adtd_end_time       C   出院结束时间:yyyy-mm-dd hh24:mi:ss
  --    fee_category        C   费别:多个用逗号分隔
  --    inp_status          N   住院状态:0-在院病人;1-出院病人;2-在院或出院
  --    pati_natures        C   病人性质：多个用逗号分0-普通住院病人,1-门诊留观病人,2-住院留观病人，NULL-表示不区分
  --    pati_name           C   姓名:可以代%分号表表按姓名匹配
  --    dept_nodeno         C   科室站点编号
  --    change_dept_pati    N   是否查询转科病人
  --    is_lastpage         N   是否取最后一次住院 0-取所有 1-取最后一次
  --    insurance_type      N   根据险类查找:>0:指定险类医保病人,0:医保和普通病人,-1:普通病人,-2:医保病人
  --    wararea_nodeno      C   病区站点编号
  --    mdlpay_mode_name    C   医疗付款方式
  --    fee_type            C   费别
  --    is_babyinfo         N 1 是否包含婴儿信息:1-包含;0-不包含
  --出参      json
  --output
  -- code                   N 1 应答码：0-失败；1-成功
  -- message                C 1 应答消息： 失败时返回具体的错误信息
  --   page_list[]          数据组  √  √
  --    pati_id             N    病人id  √  √
  --    pati_pageid         N    主页id  √  √
  --    pati_name           C    姓名  √  √
  --    pati_sex            C    性别  √  √
  --    pati_age            C    年龄  √  √
  --    inpatient_num       C    住院号  √  √
  --    pati_bed            C    出院病床  √  √
  --    insurance_type      N    险类  √  √
  --    insurance_type_name    C 1 险类名称
  --    fee_category        C    费别  √  √
  --    pati_type           C    病人类型(普通,医保,留观)  √  √
  --    adta_time           C    入院时间:yyyy-mm-dd hh24:mi:ss  √  √
  --    adtd_time           C    出院时间:yyyy-mm-dd hh24:mi:ss  √  √
  --    si_inp_status       N    住院状态:病案主页.状态(0-正常住院；1-尚未入科；2-正在转科或正在转病区；3-已预出院)  √  √
  --    pati_nature         N    病人性质:0-普通住院病人,1-门诊留观病人,2-住院留观病人
  --    pati_wardarea_id    N    当前病区id
  --    pati_wardarea_name  C    当前病区名称
  --    pati_dept_id        N    当前科室id
  --    pati_dept_name      C    当前科室名称
  --    mdlpay_mode_name    C    医疗付款方式名称
  --    mdlpay_mode_code    C    医疗付款方式编码
  --    pat_rsdpscn         C    住院医师
  --    pati_desc           C    病人备注
  --    catalog_date        C    编目日期:yyyy-mm-dd hh24:mi:ss
  --    create_pati         C    登记人
  --    in_objective        C    住院目的
  --    insurance_num       C    医保号
  --    level_of_care       C    护理等级
  --    data_adto_sign      N    数据转出标志:0-未转出，1-已转出
  --    fee_auditor_sign    N    费用审核标志:0或空-未审核,1-已审核或开始审核;2-完成审核
  --    fee_auditor         C    费用审核人
  --    pre_dstat_time      C    预出院时间
  --    last_press_money    N    上次催款金额
  --    baby_list[]           1 婴儿信息，[数组]
  --      baby_num          N 1 婴儿序号
  --      baby_name         C 1 婴儿姓名
  --      baby_sex          C 1 婴儿性别
  --      baby_date         C 1 出生时间
  ---------------------------------------------------------------------------
  j_In       Pljson;
  j_Json     Pljson;
  n_查询类型 Number(2);
  v_病区ids  Varchar2(32680);
  v_科室ids  Varchar2(32680);
  v_姓名     Varchar2(200);
  v_科室站点 Zlnodelist.编号%Type;
  v_病区站点 Zlnodelist.编号%Type;
  n_险类     病案主页.险类%Type;

  n_Like Number(2);

  c_病人ids Clob;
  c_主页ids Clob;

  n_含婴儿信息   Number(2);
  n_转科病人     Number(2);
  d_入院开始时间 Date;
  d_入院结束时间 Date;
  d_出院开始时间 Date;
  d_出院结束时间 Date;
  v_费别         Varchar2(32680);
  n_住院状态     Number(2);
  v_病人性质     Varchar2(100);
  l_病人id       t_Strlist := t_Strlist();
  n_Last         Number;
  v_医疗付款方式 Varchar2(100);
  n_Firstitem    Number(1); --是否是第一个项目

  l_主页id t_Strlist := t_Strlist();

  Cursor c_病案基本信息 Is
    Select a.病人id, a.主页id, a.住院号, a.病人性质, a.费别, To_Char(a.入院日期, 'yyyy-mm-dd hh24:mi:ss') As 入院时间, a.当前病区id, a.出院科室id,
           a.出院病床, To_Char(a.出院日期, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.住院医师,
           To_Char(a.编目日期, 'yyyy-mm-dd hh24:mi:ss') As 编目日期, a.状态, a.姓名, a.性别, a.年龄, a.险类, a.登记人, a.登记时间, a.备注, a.病人类型,
           To_Char(a.预出院日期, 'yyyy-mm-dd hh24:mi:ss') As 预出院日期, c.名称 As 当前科室名称, d.名称 As 当前病区名称, e.编码 As 医疗付款方式编码,
           a.医疗付款方式 As 医疗付款方式名称, a.住院目的, f.名称 As 护理等级, a.数据转出, a.审核标志, a.审核人, x.名称 As 险类名称
    
    From 病案主页 A, 部门表 C, 部门表 D, 医疗付款方式 E, 收费项目目录 F, 保险类别 X
    Where 病人id = 0 And 主页id = 0 And a.出院科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.医疗付款方式 = e.名称(+) And
          a.护理等级id = f.Id(+) And a.险类 = x.序号(+) And Rownum < 1;
  Type Ty_病人信息 Is Ref Cursor;
  c_病人信息 Ty_病人信息; --动态游标变量
  --组装返回数据
  Procedure Get_Jsonliststr
  (
    病人信息_In  In Ty_病人信息,
    查询类型_In  In Number,
    含婴儿_In    In Number,
    Pati_Out     In Out Clob,
    Firstitem_In In Out Number
  ) As
    r_Pati   c_病案基本信息%RowType;
    v_医保号 病案主页从表.信息值%Type;
  
    n_上次催款金额 病案主页从表.信息值%Type;
  
    n_Firstitem Number(1);
    v_Temp      Varchar2(32767);
  
    n_Firstsubitem Number(1);
    v_Tempsub      Varchar2(32767);
  Begin
    n_Firstitem := Firstitem_In;
  
    Loop
      Fetch 病人信息_In
        Into r_Pati;
      Exit When 病人信息_In%NotFound;
    
      If Nvl(n_Firstitem, 0) = 1 Then
        v_Temp := v_Temp || ',';
      Else
        n_Firstitem := 1;
      End If;
      v_Temp := v_Temp || '{';
      v_Temp := v_Temp || '"pati_id":' || r_Pati.病人id;
      v_Temp := v_Temp || ',"pati_pageid":' || Nvl(r_Pati.主页id || '', 'null');
      v_Temp := v_Temp || ',"pati_name":"' || Zljsonstr(r_Pati.姓名) || '"';
      v_Temp := v_Temp || ',"pati_sex":"' || Zljsonstr(r_Pati.性别) || '"';
      v_Temp := v_Temp || ',"pati_age":"' || Zljsonstr(r_Pati.年龄) || '"';
    
      v_Temp := v_Temp || ',"inpatient_num":"' || Zljsonstr(r_Pati.住院号) || '"';
      v_Temp := v_Temp || ',"pati_bed":"' || Zljsonstr(r_Pati.出院病床) || '"';
      v_Temp := v_Temp || ',"insurance_type":' || Nvl(r_Pati.险类, 0);
      v_Temp := v_Temp || ',"insurance_type_name":"' || Zljsonstr(r_Pati.险类名称) || '"';
      v_Temp := v_Temp || ',"fee_category":"' || Zljsonstr(r_Pati.费别) || '"';
      v_Temp := v_Temp || ',"pati_type":"' || Zljsonstr(r_Pati.病人类型) || '"';
    
      v_Temp := v_Temp || ',"adta_time":"' || Zljsonstr(r_Pati.入院时间) || '"';
      v_Temp := v_Temp || ',"adtd_time":"' || Zljsonstr(r_Pati.出院时间) || '"';
      v_Temp := v_Temp || ',"create_pati":"' || Zljsonstr(r_Pati.登记人) || '"';
      v_Temp := v_Temp || ',"create_time":"' || Zljsonstr(r_Pati.登记时间) || '"';
      v_Temp := v_Temp || ',"si_inp_status":' || Nvl(r_Pati.状态, 0);
    
      v_Temp := v_Temp || ',"pati_nature":' || Nvl(r_Pati.病人性质, 0);
      v_Temp := v_Temp || ',"in_objective":"' || Zljsonstr(r_Pati.住院目的) || '"';
    
      If Nvl(查询类型_In, 0) = 1 Then
        v_Temp := v_Temp || ',"pati_wardarea_id":' || Nvl(r_Pati.当前病区id, 0);
        v_Temp := v_Temp || ',"pati_wardarea_name":"' || Zljsonstr(r_Pati.当前病区名称) || '"';
        v_Temp := v_Temp || ',"pati_dept_id":' || Nvl(r_Pati.出院科室id, 0);
        v_Temp := v_Temp || ',"pati_dept_name":"' || Zljsonstr(r_Pati.当前科室名称) || '"';
      
        v_Temp := v_Temp || ',"mdlpay_mode_name":"' || Zljsonstr(r_Pati.医疗付款方式名称) || '"';
        v_Temp := v_Temp || ',"mdlpay_mode_code":"' || Zljsonstr(r_Pati.医疗付款方式编码) || '"';
        v_Temp := v_Temp || ',"pat_rsdpscn":"' || Zljsonstr(r_Pati.住院医师) || '"';
        v_Temp := v_Temp || ',"pati_desc":"' || Zljsonstr(r_Pati.备注) || '"';
        v_Temp := v_Temp || ',"catalog_date":"' || Zljsonstr(r_Pati.编目日期) || '"';
      
        Select Max(Decode(a.信息名, '医保号', a.信息值, 0)), Max(Decode(a.信息名, '上次催款金额', a.信息值, 0))
        Into v_医保号, n_上次催款金额
        From 病案主页从表 A
        Where a.病人id = r_Pati.病人id And a.主页id = r_Pati.主页id And a.信息名 In ('医保号', '上次催款金额');
      
        v_Temp := v_Temp || ',"insurance_num":"' || Zljsonstr(v_医保号) || '"';
        v_Temp := v_Temp || ',"level_of_care":"' || Zljsonstr(r_Pati.护理等级) || '"';
        v_Temp := v_Temp || ',"data_adto_sign":' || Nvl(r_Pati.数据转出, 0);
        v_Temp := v_Temp || ',"fee_auditor_sign":' || Nvl(r_Pati.审核标志, 0);
        v_Temp := v_Temp || ',"fee_auditor":"' || Zljsonstr(r_Pati.审核人) || '"';
        v_Temp := v_Temp || ',"pre_dstat_time":"' || Zljsonstr(r_Pati.预出院日期) || '"';
        v_Temp := v_Temp || ',"last_press_money":' || Nvl(n_上次催款金额, 0);
      End If;
    
      If Nvl(含婴儿_In, 0) = 1 Then
        v_Tempsub      := '';
        n_Firstsubitem := 1;
        For r_婴儿 In (Select 病人id, 主页id, 序号 As 婴儿序号, 婴儿姓名, 婴儿性别, To_Char(出生时间, 'yyyy-mm-dd hh24:mi:ss') As 出生时间
                     From 病人新生儿记录
                     Where 病人id = r_Pati.病人id And 主页id = r_Pati.主页id) Loop
        
          If Nvl(n_Firstsubitem, 0) = 0 Then
            v_Tempsub := v_Tempsub || ',';
          Else
            n_Firstsubitem := 0;
          End If;
        
          v_Tempsub := v_Tempsub || '{';
          v_Tempsub := v_Tempsub || '"baby_num":' || Nvl(r_婴儿.婴儿序号, 0);
          v_Tempsub := v_Tempsub || ',"baby_name":"' || Zljsonstr(r_婴儿.婴儿姓名) || '"';
          v_Tempsub := v_Tempsub || ',"baby_sex":"' || Zljsonstr(r_婴儿.婴儿性别) || '"';
          v_Tempsub := v_Tempsub || ',"baby_date":"' || Zljsonstr(r_婴儿.出生时间) || '"';
          v_Tempsub := v_Tempsub || '}';
        End Loop;
        v_Temp := v_Temp || ',"baby_list":[' || v_Tempsub || ']';
      End If;
      v_Temp := v_Temp || '}';
    
      If Length(v_Temp) > 20000 Then
        Dbms_Lob.Append(Pati_Out, v_Temp);
        v_Temp := '';
      End If;
    End Loop;
    Firstitem_In := n_Firstitem;
    If v_Temp Is Not Null Then
      Dbms_Lob.Append(Pati_Out, v_Temp);
    End If;
  End Get_Jsonliststr;
Begin

  --解析入参
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  n_查询类型 := j_Json.Get_Number('query_type');
  v_病区ids  := j_Json.Get_String('wararea_ids');
  v_科室ids  := j_Json.Get_String('dept_ids');
  If j_Json.Exist('pati_ids') Is Not Null Then
    c_病人ids := j_Json.Get_Clob('pati_ids');
  End If;
  If j_Json.Exist('pati_pageids') Is Not Null Then
    c_主页ids := j_Json.Get_Clob('pati_pageids');
  End If;

  d_入院开始时间 := To_Date(j_Json.Get_String('adta_start_time'), 'YYYY-MM-DD hh24:mi:ss');
  d_入院结束时间 := To_Date(j_Json.Get_String('adta_end_time'), 'YYYY-MM-DD hh24:mi:ss');
  d_出院开始时间 := To_Date(j_Json.Get_String('adtd_start_time'), 'YYYY-MM-DD hh24:mi:ss');
  d_出院结束时间 := To_Date(j_Json.Get_String('adtd_end_time'), 'YYYY-MM-DD hh24:mi:ss');

  v_费别         := j_Json.Get_String('fee_category');
  n_住院状态     := Nvl(j_Json.Get_Number('inp_status'), 0);
  v_病人性质     := j_Json.Get_String('pati_natures');
  v_姓名         := j_Json.Get_String('pati_name');
  v_科室站点     := j_Json.Get_String('dept_nodeno');
  n_转科病人     := Nvl(j_Json.Get_Number('change_dept_pati'), 0);
  n_Last         := j_Json.Get_Number('is_lastpage');
  n_险类         := j_Json.Get_Number('insurance_type');
  v_病区站点     := j_Json.Get_String('wararea_nodeno');
  v_医疗付款方式 := j_Json.Get_String('mdlpay_mode_name');
  n_含婴儿信息   := j_Json.Get_Number('is_babyinfo');

  If v_病人性质 Is Null Then
    v_病人性质 := ',0,1,2,';
  Else
    v_病人性质 := ',' || v_病人性质 || ',';
  End If;

  If v_病区ids Is Not Null Then
    v_病区ids := ',' || v_病区ids || ',';
  End If;

  If v_费别 Is Not Null Then
    v_费别 := ',' || v_费别 || ',';
  End If;

  If v_科室ids Is Not Null Then
    v_科室ids := ',' || v_科室ids || ',';
  End If;

  n_Like := 0;
  If Instr(Nvl(v_姓名, '-'), '%') > 0 Then
    n_Like := 1;
  End If;

  While c_病人ids Is Not Null Loop
    If Length(c_病人ids) <= 4000 Then
      l_病人id.Extend;
      l_病人id(l_病人id.Count) := c_病人ids;
      c_病人ids := Null;
    Else
      l_病人id.Extend;
      l_病人id(l_病人id.Count) := Substr(c_病人ids, 1, Instr(c_病人ids, ',', 3950) - 1);
      c_病人ids := Substr(c_病人ids, Instr(c_病人ids, ',', 3950) + 1);
    End If;
  End Loop;

  While c_主页ids Is Not Null Loop
    If Length(c_主页ids) <= 4000 Then
      l_主页id.Extend;
      l_主页id(l_主页id.Count) := c_主页ids;
      c_主页ids := Null;
    Else
      l_主页id.Extend;
      l_主页id(l_主页id.Count) := Substr(c_主页ids, 1, Instr(c_主页ids, ',', 3950) - 1);
      c_主页ids := Substr(c_主页ids, Instr(c_主页ids, ',', 3950) + 1);
    End If;
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"成功","page_list":[';
  If Nvl(n_转科病人, 0) = 1 Then
    --转科病人
    If l_病人id.Count <> 0 Then
      For I In 1 .. l_病人id.Count Loop
        --可能存在同一病人一天或范围内的有两条以上的转科,则以最后一条为准.
        Open c_病人信息 For
          Select /*+cardinality(f,10)*/
           a.病人id, a.主页id, a.住院号, a.病人性质, a.费别, To_Char(a.入院日期, 'yyyy-mm-dd hh24:mi:ss') As 入院时间, a.当前病区id, a.出院科室id,
           a.出院病床, To_Char(a.出院日期, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.住院医师,
           To_Char(a.编目日期, 'yyyy-mm-dd hh24:mi:ss') As 编目日期, a.状态, a.姓名, a.性别, a.年龄, a.险类, a.登记人, a.登记时间, a.备注, a.病人类型,
           To_Char(a.预出院日期, 'yyyy-mm-dd hh24:mi:ss') As 预出院日期, c.名称 As 当前科室名称, d.名称 As 当前病区名称, e.编码 As 医疗付款方式编码,
           a.医疗付款方式 As 医疗付款方式名称, a.住院目的, Null As 护理等级, a.数据转出, a.审核标志, a.审核人, x.名称 As 险类名称
          
          From 病案主页 A, 病人变动记录 B, 部门表 C, 部门表 D, 医疗付款方式 E, Table(f_Num2list(l_病人id(I))) F, 保险类别 X
          Where a.病人id = b.病人id And a.主页id = b.主页id And b.科室id = c.Id(+) And b.病区id = d.Id(+) And a.医疗付款方式 = e.名称(+) And
                a.险类 = x.序号(+) And a.病人id = f.Column_Value
               
                And Instr(v_病人性质, ',' || a.病人性质 || ',') > 0 And Nvl(a.主页id, 0) <> 0 And Nvl(a.状态, 0) <> 2
               
                And (v_姓名 Is Null Or (n_Like = 1 And a.姓名 Like v_姓名) Or (n_Like = 0 And a.姓名 = v_姓名))
               
                And (v_病区ids Is Null Or Instr(v_病区ids, ',' || a.当前病区id || ',') = 0) And
                (v_病区ids Is Null Or Instr(v_病区ids, ',' || b.病区id || ',') > 0)
               
                And (v_科室ids Is Null Or Instr(v_科室ids, ',' || a.出院科室id || ',') = 0) And
                (v_科室ids Is Null Or Instr(v_科室ids, ',' || b.科室id || ',') > 0)
               
                And b.终止原因 = 3 And (b.终止时间 >= d_出院开始时间 Or d_出院开始时间 Is Null) And
                (b.终止时间 <= d_出院结束时间 Or d_出院结束时间 Is Null) And Nvl(b.附加床位, 0) = 0 And Nvl(a.病案状态, 0) <> 5 And
                a.封存时间 Is Null And
                b.终止时间 = (Select Max(终止时间)
                          From 病人变动记录
                          Where 病人id = b.病人id And 主页id = b.主页id And 终止原因 = 3 And Nvl(附加床位, 0) = 0 And
                                (v_病区ids Is Null Or Instr(v_病区ids, ',' || 病区id || ',') > 0) And
                                (v_科室ids Is Null Or Instr(v_科室ids, ',' || 科室id || ',') > 0));
      
        Get_Jsonliststr(c_病人信息, n_查询类型, n_含婴儿信息, Json_Out, n_Firstitem);
        If n_Firstitem = 1 Then
          n_Firstitem := 0;
        End If;
      End Loop;
    
    Else
      --可能存在同一病人一天或范围内的有两条以上的转科,则以最后一条为准.
      Open c_病人信息 For
        Select a.病人id, a.主页id, a.住院号, a.病人性质, a.费别, To_Char(a.入院日期, 'yyyy-mm-dd hh24:mi:ss') As 入院时间, a.当前病区id, a.出院科室id,
               a.出院病床, To_Char(a.出院日期, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.住院医师,
               To_Char(a.编目日期, 'yyyy-mm-dd hh24:mi:ss') As 编目日期, a.状态, a.姓名, a.性别, a.年龄, a.险类, a.登记人, a.登记时间, a.备注,
               a.病人类型, To_Char(a.预出院日期, 'yyyy-mm-dd hh24:mi:ss') As 预出院日期, c.名称 As 当前科室名称, d.名称 As 当前病区名称,
               e.编码 As 医疗付款方式编码, a.医疗付款方式 As 医疗付款方式名称, a.住院目的, Null As 护理等级, a.数据转出, a.审核标志, a.审核人, x.名称 As 险类名称
        From 病案主页 A, 病人变动记录 B, 部门表 C, 部门表 D, 医疗付款方式 E, 保险类别 X
        Where a.病人id = b.病人id And a.主页id = b.主页id And b.科室id = c.Id(+) And b.病区id = d.Id(+) And a.医疗付款方式 = e.名称(+) And
              a.险类 = x.序号(+)
             
              And Instr(v_病人性质, ',' || a.病人性质 || ',') > 0 And Nvl(a.主页id, 0) <> 0 And Nvl(a.状态, 0) <> 2
             
              And (v_姓名 Is Null Or (n_Like = 1 And a.姓名 Like v_姓名) Or (n_Like = 0 And a.姓名 = v_姓名))
             
              And (v_病区ids Is Null Or Instr(v_病区ids, ',' || a.当前病区id || ',') = 0) And
              (v_病区ids Is Null Or Instr(v_病区ids, ',' || b.病区id || ',') > 0)
             
              And (v_科室ids Is Null Or Instr(v_科室ids, ',' || a.出院科室id || ',') = 0) And
              (v_科室ids Is Null Or Instr(v_科室ids, ',' || b.科室id || ',') > 0)
             
              And b.终止原因 = 3 And (b.终止时间 >= d_出院开始时间 Or d_出院开始时间 Is Null) And (b.终止时间 <= d_出院结束时间 Or d_出院结束时间 Is Null) And
              Nvl(b.附加床位, 0) = 0 And Nvl(a.病案状态, 0) <> 5 And a.封存时间 Is Null And
              b.终止时间 = (Select Max(终止时间)
                        From 病人变动记录
                        Where 病人id = b.病人id And 主页id = b.主页id And 终止原因 = 3 And Nvl(附加床位, 0) = 0 And
                              (v_病区ids Is Null Or Instr(v_病区ids, ',' || 病区id || ',') > 0) And
                              (v_科室ids Is Null Or Instr(v_科室ids, ',' || 科室id || ',') > 0));
      Get_Jsonliststr(c_病人信息, n_查询类型, n_含婴儿信息, Json_Out, n_Firstitem);
    End If;
  Elsif l_病人id.Count <> 0 Then
    For I In 1 .. l_病人id.Count Loop
      Open c_病人信息 For
        Select /*+cardinality(B,10)*/
         a.病人id, a.主页id, a.住院号, a.病人性质, a.费别, To_Char(a.入院日期, 'yyyy-mm-dd hh24:mi:ss') As 入院时间, a.当前病区id, a.出院科室id,
         a.出院病床, To_Char(a.出院日期, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.住院医师,
         To_Char(a.编目日期, 'yyyy-mm-dd hh24:mi:ss') As 编目日期, a.状态, a.姓名, a.性别, a.年龄, a.险类, a.登记人, a.登记时间, a.备注, a.病人类型,
         To_Char(a.预出院日期, 'yyyy-mm-dd hh24:mi:ss') As 预出院日期, c.名称 As 当前科室名称, d.名称 As 当前病区名称, e.编码 As 医疗付款方式编码,
         a.医疗付款方式 As 医疗付款方式名称, a.住院目的, Null As 护理等级, a.数据转出, a.审核标志, a.审核人, x.名称 As 险类名称
        From 病案主页 A, (Select Column_Value As 病人id From Table(f_Num2list(l_病人id(I)))) B, 部门表 C, 部门表 D, 医疗付款方式 E, 在院病人 H,
             保险类别 X
        Where a.出院科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.病人id = b.病人id And a.险类 = x.序号(+) And
              (v_姓名 Is Null Or (n_Like = 1 And a.姓名 Like v_姓名) Or (n_Like = 0 And a.姓名 = v_姓名)) And
              (v_病区ids Is Null Or Instr(v_病区ids, ',' || a.当前病区id || ',') > 0) And
              (v_科室ids Is Null Or Instr(v_科室ids, ',' || a.出院科室id || ',') > 0) And
              (v_费别 Is Null Or Instr(v_费别, ',' || a.费别 || ',') > 0) And Instr(v_病人性质, ',' || a.病人性质 || ',') > 0 And
              a.病人id = h.病人id(+) And a.主页id = h.主页id(+) And
              (Nvl(v_医疗付款方式, '-') = '-' Or ((Nvl(n_住院状态, 0) = 0 Or Nvl(n_住院状态, 0) = 1) And a.医疗付款方式 = v_医疗付款方式)) And
             
              (n_住院状态 = 0 And h.病人id Is Not Null And (a.入院日期 >= d_入院开始时间 Or d_入院开始时间 Is Null) And
              (a.入院日期 <= d_入院结束时间 Or d_入院结束时间 Is Null) Or
              
              n_住院状态 = 1 And h.病人id Is Null And (a.出院日期 >= d_出院开始时间 Or d_出院开始时间 Is Null) And
              (a.出院日期 <= d_出院结束时间 Or d_出院结束时间 Is Null) Or
              
              n_住院状态 = 2 And
              (h.病人id Is Not Null And (a.入院日期 >= d_入院开始时间 Or d_入院开始时间 Is Null) And
              (a.入院日期 <= d_入院结束时间 Or d_入院结束时间 Is Null) Or h.病人id Is Null And (a.出院日期 >= d_出院开始时间 Or d_出院开始时间 Is Null) And
              (a.出院日期 <= d_出院结束时间 Or d_出院结束时间 Is Null)))
             
              And a.医疗付款方式 = e.名称(+) And
              (a.主页id = (Select Max(主页id) From 病案主页 Where 病人id = a.病人id) And Nvl(n_Last, 0) = 1 Or Nvl(n_Last, 0) = 0) And
              (Nvl(n_险类, 0) = 0 Or n_险类 > 0 And a.险类 = n_险类 Or n_险类 = -1 And Nvl(a.险类, 0) = 0 Or
              n_险类 = -2 And Nvl(a.险类, 0) <> 0) And (c.站点 Is Null Or c.站点 = v_科室站点 Or v_科室站点 Is Null) And
              (d.站点 Is Null Or d.站点 = v_病区站点 Or v_病区站点 Is Null);
    
      Get_Jsonliststr(c_病人信息, n_查询类型, n_含婴儿信息, Json_Out, n_Firstitem);
    End Loop;
  Elsif l_主页id.Count <> 0 Then
    For I In 1 .. l_主页id.Count Loop
      Open c_病人信息 For
        With c_病人 As
         (Select Distinct C1 As 病人id, C2 As 主页id From Table(f_Num2list2(l_主页id(I))))
        Select /*+cardinality(B,10)*/
         a.病人id, a.主页id, a.住院号, a.病人性质, a.费别, To_Char(a.入院日期, 'yyyy-mm-dd hh24:mi:ss') As 入院时间, a.当前病区id, a.出院科室id,
         a.出院病床, To_Char(a.出院日期, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.住院医师,
         To_Char(a.编目日期, 'yyyy-mm-dd hh24:mi:ss') As 编目日期, a.状态, a.姓名, a.性别, a.年龄, a.险类, a.登记人, a.登记时间, a.备注, a.病人类型,
         To_Char(a.预出院日期, 'yyyy-mm-dd hh24:mi:ss') As 预出院日期, c.名称 As 当前科室名称, d.名称 As 当前病区名称, e.编码 As 医疗付款方式编码,
         a.医疗付款方式 As 医疗付款方式名称, a.住院目的, Null As 护理等级, a.数据转出, a.审核标志, a.审核人, x.名称 As 险类名称
        From 病案主页 A, c_病人 B, 部门表 C, 部门表 D, 医疗付款方式 E, 在院病人 H, 保险类别 X
        Where a.出院科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.病人id = b.病人id And a.主页id = b.主页id And a.险类 = x.序号(+) And
              (v_姓名 Is Null Or (n_Like = 1 And a.姓名 Like v_姓名) Or (n_Like = 0 And a.姓名 = v_姓名)) And
              (v_病区ids Is Null Or Instr(v_病区ids, ',' || a.当前病区id || ',') > 0) And
              (v_科室ids Is Null Or Instr(v_科室ids, ',' || a.出院科室id || ',') > 0) And
              (v_费别 Is Null Or Instr(v_费别, ',' || a.费别 || ',') > 0) And Instr(v_病人性质, ',' || a.病人性质 || ',') > 0 And
              a.病人id = h.病人id(+) And a.主页id = h.主页id(+) And
              (Nvl(v_医疗付款方式, '-') = '-' Or ((Nvl(n_住院状态, 0) = 0 Or Nvl(n_住院状态, 0) = 1) And a.医疗付款方式 = v_医疗付款方式)) And
             
              (n_住院状态 = 0 And h.病人id Is Not Null And (a.入院日期 >= d_入院开始时间 Or d_入院开始时间 Is Null) And
              (a.入院日期 <= d_入院结束时间 Or d_入院结束时间 Is Null) Or
              
              n_住院状态 = 1 And h.病人id Is Null And (a.出院日期 >= d_出院开始时间 Or d_出院开始时间 Is Null) And
              (a.出院日期 <= d_出院结束时间 Or d_出院结束时间 Is Null) Or
              
              n_住院状态 = 2 And
              (h.病人id Is Not Null And (a.入院日期 >= d_入院开始时间 Or d_入院开始时间 Is Null) And
              (a.入院日期 <= d_入院结束时间 Or d_入院结束时间 Is Null) Or h.病人id Is Null And (a.出院日期 >= d_出院开始时间 Or d_出院开始时间 Is Null) And
              (a.出院日期 <= d_出院结束时间 Or d_出院结束时间 Is Null))) And a.医疗付款方式 = e.名称(+) And
              (Nvl(n_险类, 0) = 0 Or n_险类 > 0 And a.险类 = n_险类 Or n_险类 = -1 And Nvl(a.险类, 0) = 0 Or
              n_险类 = -2 And Nvl(a.险类, 0) <> 0) And (c.站点 Is Null Or c.站点 = v_科室站点 Or v_科室站点 Is Null) And
              (d.站点 Is Null Or d.站点 = v_病区站点 Or v_病区站点 Is Null);
    
      Get_Jsonliststr(c_病人信息, n_查询类型, n_含婴儿信息, Json_Out, n_Firstitem);
    End Loop;
  Elsif d_入院开始时间 Is Not Null Then
    Open c_病人信息 For
      Select a.病人id, a.主页id, a.住院号, a.病人性质, a.费别, To_Char(a.入院日期, 'yyyy-mm-dd hh24:mi:ss') As 入院时间, a.当前病区id, a.出院科室id,
             a.出院病床, To_Char(a.出院日期, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.住院医师,
             To_Char(a.编目日期, 'yyyy-mm-dd hh24:mi:ss') As 编目日期, a.状态, a.姓名, a.性别, a.年龄, a.险类, a.登记人, a.登记时间, a.备注, a.病人类型,
             To_Char(a.预出院日期, 'yyyy-mm-dd hh24:mi:ss') As 预出院日期, c.名称 As 当前科室名称, d.名称 As 当前病区名称, e.编码 As 医疗付款方式编码,
             a.医疗付款方式 As 医疗付款方式名称, a.住院目的, Null As 护理等级, a.数据转出, a.审核标志, a.审核人, x.名称 As 险类名称
      
      From 病案主页 A, 部门表 C, 部门表 D, 医疗付款方式 E, 在院病人 H, 保险类别 X
      Where a.出院科室id = c.Id(+) And a.当前病区id = d.Id(+) And
            (v_姓名 Is Null Or (n_Like = 1 And a.姓名 Like v_姓名) Or (n_Like = 0 And a.姓名 = v_姓名)) And a.险类 = x.序号(+) And
            (v_病区ids Is Null Or Instr(v_病区ids, ',' || a.当前病区id || ',') > 0) And
            (v_科室ids Is Null Or Instr(v_科室ids, ',' || a.出院科室id || ',') > 0) And Instr(v_病人性质, ',' || a.病人性质 || ',') > 0 And
            (v_费别 Is Null Or Instr(v_费别, ',' || a.费别 || ',') > 0) And a.病人id = h.病人id(+) And a.主页id = h.主页id(+) And
            (Nvl(v_医疗付款方式, '-') = '-' Or ((Nvl(n_住院状态, 0) = 0 Or Nvl(n_住院状态, 0) = 1) And a.医疗付款方式 = v_医疗付款方式)) And
            (n_住院状态 = 0 And h.病人id Is Not Null And (a.入院日期 >= d_入院开始时间 Or d_入院开始时间 Is Null) And
            (a.入院日期 <= d_入院结束时间 Or d_入院结束时间 Is Null) Or
            
            n_住院状态 = 1 And h.病人id Is Null And (a.出院日期 >= d_出院开始时间 Or d_出院开始时间 Is Null) And
            (a.出院日期 <= d_出院结束时间 Or d_出院结束时间 Is Null) Or
            
            n_住院状态 = 2 And
            (h.病人id Is Not Null And (a.入院日期 >= d_入院开始时间 Or d_入院开始时间 Is Null) And
            (a.入院日期 <= d_入院结束时间 Or d_入院结束时间 Is Null) Or
            h.病人id Is Null And (a.出院日期 >= d_出院开始时间 Or d_出院开始时间 Is Null) And (a.出院日期 <= d_出院结束时间 Or d_出院结束时间 Is Null))) And
            a.医疗付款方式 = e.名称(+) And
            (a.主页id = (Select Max(主页id) From 病案主页 Where 病人id = a.病人id) And Nvl(n_Last, 0) = 1 Or Nvl(n_Last, 0) = 0) And
            (c.站点 Is Null Or c.站点 = v_科室站点 Or v_科室站点 Is Null) And
            (Nvl(n_险类, 0) = 0 Or n_险类 > 0 And a.险类 = n_险类 Or n_险类 = -1 And Nvl(a.险类, 0) = 0 Or
            n_险类 = -2 And Nvl(a.险类, 0) <> 0) And (d.站点 Is Null Or d.站点 = v_病区站点 Or v_病区站点 Is Null);
    Get_Jsonliststr(c_病人信息, n_查询类型, n_含婴儿信息, Json_Out, n_Firstitem);
  
  Elsif d_出院开始时间 Is Not Null Then
    Open c_病人信息 For
      Select a.病人id, a.主页id, a.住院号, a.病人性质, a.费别, To_Char(a.入院日期, 'yyyy-mm-dd hh24:mi:ss') As 入院时间, a.当前病区id, a.出院科室id,
             a.出院病床, To_Char(a.出院日期, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.住院医师,
             To_Char(a.编目日期, 'yyyy-mm-dd hh24:mi:ss') As 编目日期, a.状态, a.姓名, a.性别, a.年龄, a.险类, a.登记人, a.登记时间, a.备注, a.病人类型,
             To_Char(a.预出院日期, 'yyyy-mm-dd hh24:mi:ss') As 预出院日期, c.名称 As 当前科室名称, d.名称 As 当前病区名称, e.编码 As 医疗付款方式编码,
             a.医疗付款方式 As 医疗付款方式名称, a.住院目的, '' As 护理等级, a.数据转出, a.审核标志, a.审核人, x.名称 As 险类名称
      
      From 病案主页 A, 部门表 C, 部门表 D, 医疗付款方式 E, 在院病人 H, 保险类别 X
      Where a.出院科室id = c.Id(+) And a.当前病区id = d.Id(+) And
            (v_姓名 Is Null Or (n_Like = 1 And a.姓名 Like v_姓名) Or (n_Like = 0 And a.姓名 = v_姓名)) And a.险类 = x.序号(+) And
            (v_病区ids Is Null Or Instr(v_病区ids, ',' || a.当前病区id || ',') > 0) And
            (v_科室ids Is Null Or Instr(v_科室ids, ',' || a.出院科室id || ',') > 0) And Instr(v_病人性质, ',' || a.病人性质 || ',') > 0 And
            (v_费别 Is Null Or Instr(v_费别, ',' || a.费别 || ',') > 0) And a.病人id = h.病人id(+) And a.主页id = h.主页id(+) And
            (Nvl(v_医疗付款方式, '-') = '-' Or ((Nvl(n_住院状态, 0) = 0 Or Nvl(n_住院状态, 0) = 1) And a.医疗付款方式 = v_医疗付款方式)) And
            (n_住院状态 = 0 And h.病人id Is Not Null And (a.入院日期 >= d_入院开始时间 Or d_入院开始时间 Is Null) And
            (a.入院日期 <= d_入院结束时间 Or d_入院结束时间 Is Null) Or
            
            n_住院状态 = 1 And h.病人id Is Null And (a.出院日期 >= d_出院开始时间 Or d_出院开始时间 Is Null) And
            (a.出院日期 <= d_出院结束时间 Or d_出院结束时间 Is Null) Or
            
            n_住院状态 = 2 And
            (h.病人id Is Not Null And (a.入院日期 >= d_入院开始时间 Or d_入院开始时间 Is Null) And
            (a.入院日期 <= d_入院结束时间 Or d_入院结束时间 Is Null) Or
            h.病人id Is Null And (a.出院日期 >= d_出院开始时间 Or d_出院开始时间 Is Null) And (a.出院日期 <= d_出院结束时间 Or d_出院结束时间 Is Null))) And
            a.医疗付款方式 = e.名称(+) And
            (a.主页id = (Select Max(主页id) From 病案主页 Where 病人id = a.病人id) And Nvl(n_Last, 0) = 1 Or Nvl(n_Last, 0) = 0) And
            (c.站点 Is Null Or c.站点 = v_科室站点 Or v_科室站点 Is Null) And
            (Nvl(n_险类, 0) = 0 Or n_险类 > 0 And a.险类 = n_险类 Or n_险类 = -1 And Nvl(a.险类, 0) = 0 Or
            n_险类 = -2 And Nvl(a.险类, 0) <> 0) And (d.站点 Is Null Or d.站点 = v_病区站点 Or v_病区站点 Is Null);
    Get_Jsonliststr(c_病人信息, n_查询类型, n_含婴儿信息, Json_Out, n_Firstitem);
  
  Elsif Nvl(n_住院状态, 0) = 0 Then
    --只取在院病人
    Open c_病人信息 For
      Select a.病人id, a.主页id, a.住院号, a.病人性质, a.费别, To_Char(a.入院日期, 'yyyy-mm-dd hh24:mi:ss') As 入院时间, a.当前病区id, a.出院科室id,
             a.出院病床, To_Char(a.出院日期, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.住院医师,
             To_Char(a.编目日期, 'yyyy-mm-dd hh24:mi:ss') As 编目日期, a.状态, a.姓名, a.性别, a.年龄, a.险类, a.登记人, a.登记时间, a.备注, a.病人类型,
             To_Char(a.预出院日期, 'yyyy-mm-dd hh24:mi:ss') As 预出院日期, c.名称 As 当前科室名称, d.名称 As 当前病区名称, e.编码 As 医疗付款方式编码,
             a.医疗付款方式 As 医疗付款方式名称, a.住院目的, f.名称 As 护理等级, a.数据转出, a.审核标志, a.审核人, x.名称 As 险类名称
      
      From 病案主页 A, 在院病人 B, 部门表 C, 部门表 D, 医疗付款方式 E, 收费项目目录 F, 保险类别 X
      Where a.出院科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.病人id = b.病人id And a.主页id = b.主页id And a.险类 = x.序号(+) And
            (v_姓名 Is Null Or (n_Like = 1 And a.姓名 Like v_姓名) Or (n_Like = 0 And a.姓名 = v_姓名)) And
            (v_费别 Is Null Or Instr(v_费别, ',' || a.费别 || ',') > 0) And Instr(v_病人性质, ',' || a.病人性质 || ',') > 0 And
            (Nvl(v_医疗付款方式, '-') = '-' Or ((Nvl(n_住院状态, 0) = 0 Or Nvl(n_住院状态, 0) = 1) And a.医疗付款方式 = v_医疗付款方式)) And
            (v_病区ids Is Null Or Instr(v_病区ids, ',' || a.当前病区id || ',') > 0 Or
            Nvl(n_含婴儿信息, 0) = 1 And Instr(v_病区ids, ',' || a.婴儿病区id || ',') > 0) And
            (v_科室ids Is Null Or Instr(v_科室ids, ',' || a.出院科室id || ',') > 0) And a.医疗付款方式 = e.名称(+) And
            (a.主页id = (Select Max(主页id) From 病案主页 Where 病人id = a.病人id) And Nvl(n_Last, 0) = 1 Or Nvl(n_Last, 0) = 0) And
            (c.站点 Is Null Or c.站点 = v_科室站点 Or v_科室站点 Is Null) And
            (Nvl(n_险类, 0) = 0 Or n_险类 > 0 And a.险类 = n_险类 Or n_险类 = -1 And Nvl(a.险类, 0) = 0 Or
            n_险类 = -2 And Nvl(a.险类, 0) <> 0) And a.护理等级id = f.Id(+) And
            (d.站点 Is Null Or d.站点 = v_病区站点 Or v_病区站点 Is Null);
    Get_Jsonliststr(c_病人信息, n_查询类型, n_含婴儿信息, Json_Out, n_Firstitem);
  
  Elsif v_病区ids Is Not Null Then
    Open c_病人信息 For
      With c_病区 As
       (Select Column_Value As 病区id From Table(f_Num2list(v_病区ids)))
      Select /*+cardinality(B,10)*/
       a.病人id, a.主页id, a.住院号, a.病人性质, a.费别, To_Char(a.入院日期, 'yyyy-mm-dd hh24:mi:ss') As 入院时间, a.当前病区id, a.出院科室id, a.出院病床,
       To_Char(a.出院日期, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.住院医师, To_Char(a.编目日期, 'yyyy-mm-dd hh24:mi:ss') As 编目日期, a.状态,
       a.姓名, a.性别, a.年龄, a.险类, a.登记人, a.登记时间, a.备注, a.病人类型, To_Char(a.预出院日期, 'yyyy-mm-dd hh24:mi:ss') As 预出院日期,
       c.名称 As 当前科室名称, d.名称 As 当前病区名称, e.编码 As 医疗付款方式编码, a.医疗付款方式 As 医疗付款方式名称, a.住院目的, Null As 护理等级, a.数据转出, a.审核标志,
       a.审核人, x.名称 As 险类名称
      From 病案主页 A, c_病区 B, 部门表 C, 部门表 D, 医疗付款方式 E, 在院病人 H, 保险类别 X
      Where a.出院科室id = c.Id(+) And a.当前病区id = d.Id(+) And
            (a.当前病区id = b.病区id Or Nvl(n_含婴儿信息, 0) = 1 And a.婴儿病区id = b.病区id) And a.险类 = x.序号(+) And
            (v_姓名 Is Null Or (n_Like = 1 And a.姓名 Like v_姓名) Or (n_Like = 0 And a.姓名 = v_姓名)) And
            (v_科室ids Is Null Or Instr(v_科室ids, ',' || a.出院科室id || ',') > 0) And Instr(v_病人性质, ',' || a.病人性质 || ',') > 0 And
            (v_费别 Is Null Or Instr(v_费别, ',' || a.费别 || ',') > 0) And a.病人id = h.病人id(+) And a.主页id = h.主页id(+) And
            (Nvl(v_医疗付款方式, '-') = '-' Or ((Nvl(n_住院状态, 0) = 0 Or Nvl(n_住院状态, 0) = 1) And a.医疗付款方式 = v_医疗付款方式)) And
            a.医疗付款方式 = e.名称(+) And
            (a.主页id = (Select Max(主页id) From 病案主页 Where 病人id = a.病人id) And Nvl(n_Last, 0) = 1 Or Nvl(n_Last, 0) = 0) And
            (c.站点 Is Null Or c.站点 = v_科室站点 Or v_科室站点 Is Null) And
            (Nvl(n_险类, 0) = 0 Or n_险类 > 0 And a.险类 = n_险类 Or n_险类 = -1 And Nvl(a.险类, 0) = 0 Or
            n_险类 = -2 And Nvl(a.险类, 0) <> 0) And (d.站点 Is Null Or d.站点 = v_病区站点 Or v_病区站点 Is Null);
    Get_Jsonliststr(c_病人信息, n_查询类型, n_含婴儿信息, Json_Out, n_Firstitem);
  
  Else
    Open c_病人信息 For
      Select a.病人id, a.主页id, a.住院号, a.病人性质, a.费别, To_Char(a.入院日期, 'yyyy-mm-dd hh24:mi:ss') As 入院时间, a.当前病区id, a.出院科室id,
             a.出院病床, To_Char(a.出院日期, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.住院医师,
             To_Char(a.编目日期, 'yyyy-mm-dd hh24:mi:ss') As 编目日期, a.状态, a.姓名, a.性别, a.年龄, a.险类, a.登记人, a.登记时间, a.备注, a.病人类型,
             To_Char(a.预出院日期, 'yyyy-mm-dd hh24:mi:ss') As 预出院日期, c.名称 As 当前科室名称, d.名称 As 当前病区名称, e.编码 As 医疗付款方式编码,
             a.医疗付款方式 As 医疗付款方式名称, a.住院目的, Null As 护理等级, a.数据转出, a.审核标志, a.审核人, x.名称 As 险类名称
      
      From 病案主页 A, 部门表 C, 部门表 D, 医疗付款方式 E, 在院病人 H, 保险类别 X
      Where a.出院科室id = c.Id(+) And a.当前病区id = d.Id(+) And
            (v_姓名 Is Null Or (n_Like = 1 And a.姓名 Like v_姓名) Or (n_Like = 0 And a.姓名 = v_姓名)) And a.险类 = x.序号(+) And
            (v_科室ids Is Null Or Instr(v_科室ids, ',' || a.出院科室id || ',') > 0) And Instr(v_病人性质, ',' || a.病人性质 || ',') > 0 And
            (v_费别 Is Null Or Instr(v_费别, ',' || a.费别 || ',') > 0) And a.病人id = h.病人id(+) And a.主页id = h.主页id(+) And
            a.医疗付款方式 = e.名称(+) And
            (Nvl(v_医疗付款方式, '-') = '-' Or ((Nvl(n_住院状态, 0) = 0 Or Nvl(n_住院状态, 0) = 1) And a.医疗付款方式 = v_医疗付款方式)) And
            (a.主页id = (Select Max(主页id) From 病案主页 Where 病人id = a.病人id) And Nvl(n_Last, 0) = 1 Or Nvl(n_Last, 0) = 0) And
            (c.站点 Is Null Or c.站点 = v_科室站点 Or v_科室站点 Is Null) And
            (Nvl(n_险类, 0) = 0 Or n_险类 > 0 And a.险类 = n_险类 Or n_险类 = -1 And Nvl(a.险类, 0) = 0 Or
            n_险类 = -2 And Nvl(a.险类, 0) <> 0) And (d.站点 Is Null Or d.站点 = v_病区站点 Or v_病区站点 Is Null);
    Get_Jsonliststr(c_病人信息, n_查询类型, n_含婴儿信息, Json_Out, n_Firstitem);
  
  End If;

  Dbms_Lob.Append(Json_Out, ']}}');
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getpatpageinfbyrange;
/
Create Or Replace Procedure Zl_Cissvr_Getpatpagewarnscheme
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:获取病人的适用类型
  --入参：Json_In:格式
  --    input
  --      pati_id           N 1 病人id
  --      pati_pageid       N 1 主页id
  --出参: Json_Out,格式如下
  --  output
  --    code                  N 1 应答码：0-失败；1-成功
  --    message               C 1 应答消息：失败时返回具体的错误信息
  --    pati_type             N 1 适用病人类型
  ---------------------------------------------------------------------------
  j_In       Pljson;
  j_Json     Pljson;
  n_病人id   病人变动记录.病人id%Type;
  n_主页id   病人变动记录.主页id%Type;
  v_适用类型 Varchar2(100);
Begin
  --解析入参
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');
  n_主页id := j_Json.Get_Number('pati_pageid');

  If Nvl(n_病人id, 0) = 0 Then
    Json_Out := Zljsonout('未传入病人id，请检查！');
    Return;
  End If;

  v_适用类型 := Zl_Patiwarnscheme(n_病人id, n_主页id);
  Json_Out   := '{"output":{"code":1,"message":"成功","pati_type":"' || Zljsonstr(v_适用类型) || '"}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getpatpagewarnscheme;
/
Create Or Replace Procedure Zl_Cissvr_Getstufferrdata
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：临床医嘱发送生成卫材数据同步
  --入参：Json_In:格式
  --  input
  --      pati_ids                        C 1 病人ids逗号拼串
  --出参: Json_Out,格式如下
  --   output:
  --     code: 1,
  --     message: 成功,
  --     pati_bill_list[]
  --         pati_id                      N 1 病人id
  --         pati_pageid                  N 0 主页id，住院病人传入，门诊传0
  --         rgst_id                      N 0 挂号id，门诊病人传入，住院病人传空
  --         rgst_no                      C 0 挂号单号
  --         send_no                      N 1 发送号
  --         order_list[]医嘱信息列表
  --             advice_id                N 1 医嘱id
  --             effectivetime            N 1 医嘱期效
  --             emergency_tag            N 1 紧急标志
  --             denominated              N 1 计价特性
  --             fee_source               N 0 费用来源：1-门诊；2-住院
  --             fee_billtype             N 0 费用单据类型：1-收费处方；2-记帐单处方
  --             fee_no                   C 0 费用单据号
  --             freq_name                C 0 频次名称
  --             single                   N 0 单量
  ------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  n_第1行   Number;
  c_Jtmp    Clob;
  n_Rgst_Id Number;

  j_Json       Pljson;
  j_Tmp        Pljson;
  v_Jtmp       Varchar2(32767);
  v_Stuff_List Varchar2(32767);
  c_Stuff_List Clob;
  v_Vals       Clob;
  l_Vals       t_Strlist;
  n_费用来源   Number;

  --门诊医嘱数据
  Cursor c_Out
  (
    挂号单_In 病人医嘱记录.挂号单%Type,
    发送号_In 病人医嘱发送.发送号%Type
  ) Is
    Select b.医嘱id As ID, b.No, b.发送号, a.医嘱期效, a.紧急标志, a.计价特性, b.记录性质, c.门诊记帐, a.执行频次, a.单次用量
    From 病人医嘱记录 A, 病人医嘱异常记录 B, 病人医嘱发送 C
    Where a.Id = b.医嘱id And a.挂号单 = 挂号单_In And b.产生环节 = 2 And b.发送号 = 发送号_In And b.医嘱id = c.医嘱id And b.发送号 = c.发送号;

  Type t_Order Is Table Of c_Out%RowType;
  r_Odr t_Order;

  Cursor c_In
  (
    病人id_In 病人医嘱记录.病人id%Type,
    主页id_In 病人医嘱记录.主页id%Type,
    发送号_In 病人医嘱发送.发送号%Type
  ) Is
    Select b.医嘱id As ID, b.No, b.发送号, a.医嘱期效, a.紧急标志, a.计价特性, b.记录性质, c.门诊记帐, a.执行频次, a.单次用量
    From 病人医嘱记录 A, 病人医嘱异常记录 B, 病人医嘱发送 C
    Where a.Id = b.医嘱id And a.病人id = 病人id_In And a.主页id = 主页id_In And b.产生环节 = 2 And b.发送号 = 发送号_In And b.医嘱id = c.医嘱id And
          b.发送号 = c.发送号;

Begin
  --解析入参
  If Json_In Is Null Then
    Select f_List2str(Cast(Collect(a.病人id || '') As t_Strlist), ',')
    Into v_Vals
    From (Select a.病人id From 病人医嘱异常记录 A Where a.产生环节 = 2 Group By a.病人id) A;
  Else
    j_Tmp  := Pljson(Json_In);
    j_Json := j_Tmp.Get_Pljson('input');
    v_Vals := j_Json.Get_Clob('pati_ids');
    If v_Vals Is Null Then
      Select f_List2str(Cast(Collect(a.病人id || '') As t_Strlist), ',')
      Into v_Vals
      From (Select a.病人id From 病人医嘱异常记录 A Where a.产生环节 = 2 Group By a.病人id) A;
    End If;
  End If;

  l_Vals := t_Strlist();
  While v_Vals Is Not Null Loop
    If Length(v_Vals) <= 4000 Then
      l_Vals.Extend;
      l_Vals(l_Vals.Count) := v_Vals;
      v_Vals := Null;
    Else
      l_Vals.Extend;
      l_Vals(l_Vals.Count) := Substr(v_Vals, 1, Instr(v_Vals, ',', 3980) - 1);
      v_Vals := Substr(v_Vals, Instr(v_Vals, ',', 3980) + 1);
    End If;
  End Loop;

  n_第1行 := 0;
  For Lp In 1 .. l_Vals.Count Loop
    For Cp In (Select /*+cardinality(j,10) */
                a.病人id, a.主页id, a.挂号单, b.发送号
               From 病人医嘱记录 A, 病人医嘱异常记录 B, Table(f_Num2list(l_Vals(Lp))) J
               Where a.Id = b.医嘱id And a.病人id = j.Column_Value And b.产生环节 = 2
               Group By a.病人id, a.主页id, a.挂号单, b.发送号) Loop
    
      n_Rgst_Id := 0;
      If Cp.挂号单 Is Not Null Then
        Select a.Id Into n_Rgst_Id From 病人挂号记录 A Where a.No = Cp.挂号单 And a.记录性质 = 1 And a.记录状态 = 1;
      End If;
    
      If Cp.挂号单 Is Null Then
        Open c_In(Cp.病人id, Cp.主页id, Cp.发送号);
        Fetch c_In Bulk Collect
          Into r_Odr;
        Close c_In;
      Else
        Open c_Out(Cp.挂号单, Cp.发送号);
        Fetch c_Out Bulk Collect
          Into r_Odr;
        Close c_Out;
      End If;
    
      n_第1行 := n_第1行 + 1;
      If r_Odr.Count > 0 Then
        v_Stuff_List := Null;
        c_Stuff_List := Null;
        For Ol In 1 .. r_Odr.Count Loop
          If r_Odr(Ol).记录性质 = 1 Or r_Odr(Ol).记录性质 = 2 And r_Odr(Ol).门诊记帐 = 1 Then
            n_费用来源 := 1;
          Else
            n_费用来源 := 2;
          End If;
        
          v_Stuff_List := v_Stuff_List || ',{"advice_id":' || r_Odr(Ol).Id;
          v_Stuff_List := v_Stuff_List || ',"effectivetime":' || r_Odr(Ol).医嘱期效;
          v_Stuff_List := v_Stuff_List || ',"emergency_tag":' || Nvl(r_Odr(Ol).紧急标志 || '', 'null');
          v_Stuff_List := v_Stuff_List || ',"denominated":' || Nvl(r_Odr(Ol).计价特性 || '', 'null');
          v_Stuff_List := v_Stuff_List || ',"fee_source":' || n_费用来源;
          v_Stuff_List := v_Stuff_List || ',"fee_billtype":' || r_Odr(Ol).记录性质;
          v_Stuff_List := v_Stuff_List || ',"fee_no":"' || r_Odr(Ol).No || '"';
          v_Stuff_List := v_Stuff_List || ',"freq_name":"' || Zljsonstr(r_Odr(Ol).执行频次) || '"';
          v_Stuff_List := v_Stuff_List || ',"single":' || Zljsonstr(r_Odr(Ol).单次用量, 1);
          v_Stuff_List := v_Stuff_List || '}';
        
          If Length(v_Stuff_List) > 30000 Then
            If c_Stuff_List Is Null Then
              c_Stuff_List := Substr(v_Stuff_List, 2);
            Else
              c_Stuff_List := c_Stuff_List || v_Stuff_List;
            End If;
            v_Stuff_List := Null;
          End If;
        End Loop;
      
        v_Jtmp := Null;
        v_Jtmp := v_Jtmp || ',{"pati_id":' || Cp.病人id;
        v_Jtmp := v_Jtmp || ',"pati_pageid":' || Nvl(Cp.主页id || '', 'null');
        v_Jtmp := v_Jtmp || ',"rgst_id":' || Nvl(n_Rgst_Id || '', 'null');
        v_Jtmp := v_Jtmp || ',"rgst_no":"' || Cp.挂号单 || '"';
        v_Jtmp := v_Jtmp || ',"send_no":' || Cp.发送号;
      
        If n_第1行 = 1 Then
          c_Jtmp := Substr(v_Jtmp, 2);
        Else
          c_Jtmp := c_Jtmp || v_Jtmp;
        End If;
      
        --医嘱列表结点
        If c_Stuff_List Is Null Then
          c_Jtmp := c_Jtmp || ',"order_list":[' || Substr(v_Stuff_List, 2) || ']';
        Else
          c_Jtmp := c_Jtmp || ',"order_list":[' || c_Stuff_List || v_Stuff_List || ']';
        End If;
      
        c_Jtmp := c_Jtmp || '}';
      End If;
    End Loop;
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"成功","pati_bill_list":[' || c_Jtmp || ']}}';
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Cissvr_Getstufferrdata;
/
Create Or Replace Procedure Zl_Cissvr_Getsurgandaneinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:获取病人的手术及麻醉信息
  --入参：Json_In:格式
  -- input
  --   pati_id              N 1 病人id
  --   pati_pageid          N 1 主页id
  --出参: Json_Out,格式如下
  --  output
  --    code                N  1 应答码：0-失败；1-成功
  --    message             C  1 应答消息：失败时返回具体的错误信息
  --    surg_list[]                [数组]
  --      dz_code           C 1 疾病编码
  --      dz_name           C 1 疾病名称
  --      sicstype          C 1 手术切口分类
  --      icshlv            C 1 手术切口愈合等级
  --      oper_date         C 1 手术日期:yyyy-mm-dd hh24:mi:ss
  --      aneitem_type      C 1 麻醉类型
  --      surgeon_name      C 1 主刀医生姓名
  --      first_assistant   C 1 第一助手
  --      second_assistant  C 1 第二助手
  --      surg_anst         C 1 麻醉医生
  --      recoder           C   记录人
  --      rec_time          C   记录时间:yyyy-mm-dd hh24:mi:ss
  ---------------------------------------------------------------------------
  j_In     Pljson;
  j_Json   Pljson;
  n_病人id 病人手麻记录.病人id%Type;
  n_主页id 病人手麻记录.主页id%Type;
  v_List   Varchar2(32767);
Begin
  --解析入参
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');
  n_主页id := j_Json.Get_Number('pati_pageid');

  If Nvl(n_病人id, 0) = 0 Or Nvl(n_主页id, 0) = 0 Then
    Json_Out := Zljsonout('未传入病人id或主页id，请检查！');
    Return;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功","surg_list":[';
  For r_手麻 In (Select b.编码, b.名称, a.切口, a.愈合, To_Char(a.手术日期, 'yyyy-mm-dd hh24:mi:ss') As 手术日期, a.麻醉类型, a.主刀医师, a.第一助手,
                      a.第二助手, a.麻醉医师, a.记录人, To_Char(Nvl(a.记录日期, Sysdate), 'yyyy-mm-dd hh24:mi:ss') As 记录日期
               From 病人手麻记录 A, 疾病编码目录 B
               Where a.手术操作id = b.Id And a.病人id = n_病人id And a.主页id = n_主页id) Loop
  
    --      dz_code           C 1 疾病编码
    --      dz_name           C 1 疾病名称
    --      sicstype          C 1 手术切口分类
    --      icshlv            C 1 手术切口愈合等级
    --      oper_date         C 1 手术日期:yyyy-mm-dd hh24:mi:ss
    --      aneitem_type      C 1 麻醉类型
  
    Zljsonputvalue(v_List, 'dz_code', r_手麻.编码, 0, 1);
    Zljsonputvalue(v_List, 'dz_name', r_手麻.名称);
    Zljsonputvalue(v_List, 'sicstype', r_手麻.切口);
    Zljsonputvalue(v_List, 'icshlv', r_手麻.愈合);
    Zljsonputvalue(v_List, 'oper_date', r_手麻.手术日期);
    Zljsonputvalue(v_List, 'aneitem_type', r_手麻.麻醉类型);
  
    --      surgeon_name      C 1 主刀医生姓名
    --      first_assistant   C 1 第一助手
    --      second_assistant  C 1 第二助手
    --      surg_anst         C 1 麻醉医生
    --      recoder           C   记录人
    --      rec_time          C   记录时间:yyyy-mm-dd hh24:mi:ss
    Zljsonputvalue(v_List, 'surgeon_name', r_手麻.主刀医师);
    Zljsonputvalue(v_List, 'first_assistant', r_手麻.第一助手);
    Zljsonputvalue(v_List, 'second_assistant', r_手麻.第二助手);
    Zljsonputvalue(v_List, 'surg_anst', r_手麻.麻醉医师);
    Zljsonputvalue(v_List, 'recoder', r_手麻.记录人);
    Zljsonputvalue(v_List, 'rec_time', r_手麻.记录日期, 0, 2);
    If Length(v_List) > 20000 Then
      Json_Out := Json_Out || v_List;
      v_List   := ',';
    End If;
  End Loop;
  If v_List <> ',' Then
    Json_Out := Json_Out || v_List || ']}}';
  Else
    Json_Out := Json_Out || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getsurgandaneinfo;
/
Create Or Replace Procedure Zl_Cissvr_Isouttakedrug
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：根据病人ID和主页ID,判断该病人是否出院带药
  --入参：Json_In:格式
  --    input
  --        pati_id                 N   1   病人ID
  --        pati_pageid             N   1   主页ID
  --    出参 json
  --    output
  --        code                    N   1   应答吗：0-失败；1-成功
  --        message                 C   1   应答消息：失败时返回具体的错误信息
  --        isexist                 N   1   是否存在: 1-存在;0-不存在
  ---------------------------------------------------------------------------
  j_In     Pljson;
  j_Json   Pljson;
  n_病人id 病人医嘱记录.病人id%Type;
  n_主页id 病人医嘱记录.主页id%Type;
  n_Count  Number(18);
Begin

  --解析入参
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');
  n_主页id := j_Json.Get_Number('pati_pageid');

  If n_病人id Is Null Then
    Json_Out := Zljsonout('未传入病人信息');
    Return;
  End If;

  Select Count(1)
  Into n_Count
  From 病人医嘱记录 A, 病人医嘱记录 B
  Where a.相关id = b.Id And Nvl(a.执行性质, 0) <> 5 And Nvl(b.执行性质, 0) = 5 And a.诊疗类别 In ('5', '6', '7') And a.病人id = n_病人id And
        a.主页id = n_主页id And Rownum < 2;

  Json_Out := '{"output":{"code":1,"message":"成功","isexist":' || n_Count || '}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Isouttakedrug;
/
Create Or Replace Procedure Zl_Cissvr_Newoutpativisitrec
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能: 挂号成功后生成临床就诊登记记录，  目前临床的就诊登记记录就是病人挂号记录，所以不用处理。
  --input
  --  rgst_no             C 1 挂号单号
  --  rgst_appt_sign      N 1 预约标志：0-挂号记录,1-预约记录
  --  rgst_code           C 1 号别
  --  rgst_rec_id         N 1 出诊记录id
  --  outptyp_name        C 1 号类
  --  pati_id             N 1 病人ID
  --  outpno              C 1 门诊号
  --  pati_name           C 1 姓名
  --  pati_sex            C 1 性别
  --  pati_age            C 1 年龄
  --  fee_category        C 1 费别
  --  revst_sign          N 1 复诊:0-否，1-是
  --  emg_sign            N 1 急诊
  --  outproom_name       C 1 诊室
  --  exe_deptid          N 1 执行部门ID
  --  rgst_exetr          C 1 执行人
  --  happen_time         C 1 发生时间
  --  close_account_type  N 1 结算模式
  --出参      json
  --output
  --  code                N 1 应答码：0-失败；1-成功
  --  message             C 1 应答消息：成功时返回成功信息,失败时返回具体的错误信息
  ----------------------------------------------------------------------------
  j_In         Pljson;
  j_Json       Pljson;
  v_挂号单号   病人挂号记录.No%Type;
  n_预约标志   病人挂号记录.记录性质%Type;
  v_号别       病人挂号记录.号别%Type;
  n_出诊记录id 病人挂号记录.出诊记录id%Type;
  v_号类       病人挂号记录.号类%Type;
  n_病人id     病人挂号记录.病人id%Type;
  n_门诊号     病人挂号记录.门诊号%Type;
  v_姓名       病人挂号记录.姓名%Type;
  v_性别       病人挂号记录.性别%Type;
  v_年龄       病人挂号记录.年龄%Type;
  n_复诊       病人挂号记录.复诊%Type;
  v_费别       病人挂号记录.费别%Type;
  n_急诊       病人挂号记录.急诊%Type;
  v_诊室       病人挂号记录.诊室%Type;
  n_执行部门id 病人挂号记录.执行部门id%Type;
  v_执行人     病人挂号记录.执行人%Type;
  d_发生时间   病人挂号记录.发生时间%Type;
  n_结算模式   病人挂号记录.结算模式%Type;

Begin
  j_In         := Pljson(Json_In);
  j_Json       := j_In.Get_Pljson('input');
  v_挂号单号   := j_Json.Get_String('rgst_no');
  n_预约标志   := j_Json.Get_Number('rgst_appt_sign');
  v_号别       := j_Json.Get_String('rgst_code');
  n_出诊记录id := j_Json.Get_Number('rgst_rec_id');
  v_号类       := j_Json.Get_String('outptyp_name');
  n_病人id     := j_Json.Get_Number('pati_id');
  n_门诊号     := To_Number(j_Json.Get_String('outpno'));
  v_姓名       := j_Json.Get_String('pati_name');
  v_性别       := j_Json.Get_String('pati_sex');
  v_年龄       := j_Json.Get_String('pati_age');
  n_复诊       := j_Json.Get_Number('revst_sign');
  v_费别       := j_Json.Get_String('fee_category');
  n_急诊       := j_Json.Get_Number('emg_sign');
  v_诊室       := j_Json.Get_String('outproom_name');
  n_执行部门id := j_Json.Get_Number('exe_deptid');
  v_执行人     := j_Json.Get_String('rgst_exetr');
  d_发生时间   := To_Date(j_Json.Get_String('happen_time'), 'YYYY-MM-DD hh24:mi:ss');
  n_结算模式   := j_Json.Get_Number('close_account_type');

  Json_Out := Zljsonout('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Newoutpativisitrec;
/
Create Or Replace Procedure Zl_Cissvr_Patiexistmemo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:根据病人id及主页id检查是否存在备注信息
  --入参：Json_In:格式
  -- input
  --   pati_id              N 1 病人id
  --   pati_pageid          N 1 主页id
  --出参: Json_Out,格式如下
  --  output
  --    code               C  1 应答码：0-失败；1-成功
  --    message            C  1 应答消息：失败时返回具体的错误信息
  --    exist              N  1 是否存在备注：1-是的;0-否
  ---------------------------------------------------------------------------
  j_In     Pljson;
  j_Json   Pljson;
  n_病人id 病人备注信息.病人id%Type;
  n_主页id 病人备注信息.主页id%Type;
  n_Tmp    Number(1);

Begin
  --解析入参
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');
  n_主页id := j_Json.Get_Number('pati_pageid');

  If Nvl(n_病人id, 0) = 0 Then
    Json_Out := Zljsonout('必须传入病人id！');
    Return;
  End If;

  Select Count(1)
  Into n_Tmp
  From 病人备注信息
  Where 病人id = n_病人id And Nvl(主页id, 0) = Nvl(n_主页id, 0) And Rownum <= 1;

  Json_Out := '{"output":{"code":1,"message":"成功","exist":' || n_Tmp || '}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Patiexistmemo;
/
Create Or Replace Procedure Zl_Cissvr_Patiisdead
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:根据病人id检查病人是否已经死亡
  --入参：Json_In:格式
  -- input
  --   pati_id              N 1 病人id
  --出参: Json_Out,格式如下
  --  output
  --    code               C  1 应答码：0-失败；1-成功
  --    message            C  1 应答消息：失败时返回具体的错误信息
  --    isdeath            N  1 是否已经死亡：1-是的;0-否
  ---------------------------------------------------------------------------
  j_In     Pljson;
  j_Json   Pljson;
  n_病人id 病人变动记录.病人id%Type;
  n_Tmp    Number(1);

Begin
  --解析入参
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');

  If Nvl(n_病人id, 0) = 0 Then
    Json_Out := Zljsonout('必须传入病人id！');
    Return;
  End If;

  Select Count(1) Into n_Tmp From 病案主页 Where 病人id = n_病人id And 出院方式 Like '%死亡%' And Rownum <= 1;

  Json_Out := '{"output":{"code":1,"message":"成功","isdeath":' || n_Tmp || '}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Patiisdead;
/
Create Or Replace Procedure Zl_Cissvr_Patiisinhospital
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  --------------------------------------------------------------------------------------------------------------------
  --功能:判断指定病人是否是处于在院就医状态
  --入参：Json_In,格式如下
  --  input
  --    pati_id            N 1 病人ID
  --    pati_pageid        N 1 主页id
  --出参: Json_Out,格式如下
  --  output
  --    code               N 1 应答吗：0-失败；1-成功
  --    message            C 1 应答消息：失败时返回具体的错误信息
  --    inhouspital        N 1 0-不是在院就医状态，1-是处于在院就医状态
  --------------------------------------------------------------------------------------------------------------------
  j_In     Pljson;
  j_Json   Pljson;
  n_病人id 病案主页.病人id%Type;
  n_主页id 病案主页.主页id%Type;
  n_状态   Number;

Begin
  --解析入参
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');
  n_主页id := j_Json.Get_Number('pati_pageid');

  n_状态 := Zl_Pati_Is_Inhospital(n_病人id, n_主页id);

  Json_Out := '{"output":{"code":1,"message":"成功","inhouspital":' || Nvl(n_状态, 0) || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Patiisinhospital;
/
Create Or Replace Procedure Zl_Cissvr_Patiisout
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：查询病人是否已经出院
  --input      查询病人是否已经出院
  --  pati_id               N  1  病人id
  --  pati_pageid           N  1  主页id
  --  query_type            N  1  查询类型：0-单个病人查询；1-多个病人批量查询
  --  pati_pageids          C  1  格式：病人ID:主页ID,病人ID:主页ID,...
  --output
  --  code                  C  1  应答码：0-失败；1-成功
  --  message               C  1  应答消息：失败时返回具体的错误信息
  --  pati_outsign          N  1  出院标记：0-未出院，1-出院；query_type=0时返回
  --  item_list[]           出院标记列表，query_type=1时返回
  --    pati_id             N  1  病人id
  --    pati_outsign        N  1  出院标记：0-未出院，1-出院
  ---------------------------------------------------------------------------
  j_In    PLJson;
  j_Input PLJson;

  n_查询类型 Number(1);
  n_病人id   Number(18);
  n_主页id   Number(5);
  n_出院标记 Number(1); --0-未出院，1-出院

  c_主页id Clob;
  l_主页id t_StrList := t_StrList();

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;
Begin
  j_In       := PLJson(Json_In);
  j_Input    := j_In.Get_Pljson('input');
  n_查询类型 := j_Input.Get_Number('query_type');

  --0-单个病人查询
  If Nvl(n_查询类型, 0) = 0 Then
    n_病人id := j_Input.Get_Number('pati_id');
    n_主页id := j_Input.Get_Number('pati_pageid');
  
    Select Count(1)
    Into n_出院标记
    From 病案主页
    Where 病人id = n_病人id And 主页id = n_主页id And 出院日期 Is Not Null And Rownum < 2;
  
    Json_Out := '{"output":{"code":1,"message": "成功","pati_outsign":' || n_出院标记 || '}}';
    Return;
  End If;

  --1 - 多个病人批量查询 
  If Nvl(n_查询类型, 0) = 1 Then
    c_主页id := j_Input.Get_Clob('pati_pageids');
  
    While c_主页id Is Not Null Loop
      If Length(c_主页id) <= 4000 Then
        l_主页id.Extend;
        l_主页id(l_主页id.Count) := c_主页id;
        c_主页id := Null;
      Else
        l_主页id.Extend;
        l_主页id(l_主页id.Count) := Substr(c_主页id, 1, Instr(c_主页id, ',', 3950) - 1);
        c_主页id := Substr(c_主页id, Instr(c_主页id, ',', 3950) + 1);
      End If;
    End Loop;
  
    For I In 1 .. l_主页id.Count Loop
      For r_病人 In (Select /*+Cardinality(j,10)*/
                    j.C1 As 病人id, Decode(a.出院日期, Null, 0, 1) As 出院标志
                   From 病案主页 A, Table(f_Num2List2(l_主页id(I), ',', ':')) J
                   Where a.病人id(+) = j.C1 And a.主页id(+) = j.C2) Loop
      
        v_Jtmp := v_Jtmp || ',{"pati_id":' || r_病人.病人id;
        v_Jtmp := v_Jtmp || ',"pati_outsign":' || r_病人.出院标志;
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
    Return;
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || zlJsonStr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Patiisout;
/
Create Or Replace Procedure Zl_Cissvr_Patiobsvtoinhos
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：将住院留观病人转为住院病人
  --入参:JSON格式
  --input
  --   pati_id              N 1 病人id
  --   pati_pageid          N 1 主页id
  --   inpatient_num        C 1 住院号
  --出参：JSON格式
  --output
  --    code                 N   1   应答吗：0-失败；1-成功
  --    message              C   1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  n_病人id 病案主页.病人id%Type;
  n_主页id 病案主页.主页id%Type;
  n_住院号 病案主页.住院号%Type;
  j_In     Pljson;
  j_Json   Pljson;
Begin
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');
  n_主页id := j_Json.Get_Number('pati_pageid');
  n_住院号 := To_Number(j_Json.Get_String('inpatient_num'));
  Zl_病人变动记录_转住院_s(n_病人id, n_主页id, n_住院号);
  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Patiobsvtoinhos;
/
Create Or Replace Procedure Zl_Cissvr_Patipageiscatalogue
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:检查指定的住院是否已经编码
  --入参：Json_In:格式
  --    input
  --      pati_id           N 1 病人id
  --      pati_pageid       N 1 主页id
  --出参: Json_Out,格式如下
  --  output
  --    code                  N 1 应答码：0-失败；1-成功
  --    message               C 1 应答消息：失败时返回具体的错误信息
  --    pati_Name             C 1 姓名
  --    iscatalogueed         N 1 是否已经编目:1-已经编目;0-未编目
  ---------------------------------------------------------------------------
  j_In     Pljson;
  j_Json   Pljson;
  n_病人id 病案主页.病人id%Type;
  n_主页id 病案主页.主页id%Type;
  v_姓名   病案主页.姓名%Type;
  n_Tmp    Number(1);
  --组装失败时返回的数据
  Function Get_Err_Message(Message_In Varchar2) Return Clob Is
    j_Out Clob;
  Begin
    j_Out := '{"output":{"code":0,"message":"' || Zltools.Zljsonstr(Message_In) || '"}}';
  
    Return j_Out;
  End Get_Err_Message;

Begin
  --解析入参
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');
  n_主页id := j_Json.Get_Number('pati_pageid');

  If Nvl(n_病人id, 0) = 0 Then
    Json_Out := Get_Err_Message('未传入病人id，请检查！');
    Return;
  End If;

  If Nvl(n_主页id, 0) = 0 Then
    Json_Out := Get_Err_Message('未传入主页id，请检查！');
    Return;
  End If;

  Select Count(1), Max(姓名)
  Into n_Tmp, v_姓名
  From 病案主页
  Where 病人id = n_病人id And 主页id = n_主页id And 编目日期 Is Not Null;

  Json_Out := '{"output":{"code":1,"message": "成功","pati_Name":"' || Zljsonstr(v_姓名) || '","iscatalogueed":' || n_Tmp || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Patipageiscatalogue;
/
Create Or Replace Procedure Zl_Cissvr_Patiregistblacklist
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  --------------------------------------------------------------------------------------------------
  --功能:实名认证前的检查 
  --入参 JSOM格式
  --input
  --  pati_id                N 1 病人id
  --  calcdate               C 1 计算日期
  --  operat_name            C 1 操作人员
  --  order_last_date        C 1 最后预约时间
  --  black_info             C 1 病人id,附加标志;病人id,附加标志....
  --出参 JSON格式
  --output
  --  code            N 1 应答码：0-失败；1-成功
  --  message         C 1 应答消息：失败时返回具体的错误信息
  --  black_list[]
  --      operat_type     C 1 行为类别
  --      pati_id         N 1 病人id
  --      order_date      C 1 预约时间
  --      in_reason       C 1 加入原因
  --      in_explain      C 1 加入说明
  --      sign            C 1 附加标志
  --      create_name     C 1 登记人
  --------------------------------------------------------------------------------------------------
  j_In           Pljson;
  j_Json         Pljson;
  n_病人id       Number;
  d_计算日期     Date;
  v_操作人员     Varchar2(100);
  d_最后预约时间 Date;
  n_Count        Number(18);
  n_预约接诊效期 Number(18);
  n_预约退号效期 Number(18);
  n_预约接收效期 Number(18);
  v_List         Varchar2(32676);
  v_Para         Varchar2(32767);
  c_不良信息     Clob;
  l_不良信息     t_Strlist := t_Strlist();
Begin
  j_In           := Pljson(Json_In);
  j_Json         := j_In.Get_Pljson('input');
  n_病人id       := j_Json.Get_Number('pati_id');
  v_操作人员     := j_Json.Get_String('create_name');
  d_计算日期     := To_Date(j_Json.Get_String('calc_date'), 'yyyy-mm-dd hh24:mi:ss');
  d_最后预约时间 := To_Date(j_Json.Get_String('order_last_date'), 'yyyy-mm-dd hh24:mi:ss');
  c_不良信息     := j_Json.Get_Clob('black_info');

  While c_不良信息 Is Not Null Loop
    If Length(c_不良信息) <= 4000 Then
      l_不良信息.Extend;
      l_不良信息(l_不良信息.Count) := c_不良信息;
      c_不良信息 := Null;
    Else
      l_不良信息.Extend;
      l_不良信息(l_不良信息.Count) := Substr(c_不良信息, 1, Instr(c_不良信息, ';', 3980) - 1);
      c_不良信息 := Substr(c_不良信息, Instr(c_不良信息, ';', 3980) + 1);
    End If;
  End Loop;

  If Nvl(n_病人id, 0) = 0 Then
    d_计算日期 := Trunc(Sysdate) - 1 / 24 / 60 / 60; -- 缺省计算头一天的数据
  End If;
  v_Para  := Nvl(zl_GetSysParameter('预约有效就诊期限', '1111'), '0|0|0');
  n_Count := Instr(v_Para, '|');
  If n_Count = 0 Then
    n_预约接收效期 := To_Number(Nvl(v_Para, '0'));
    v_Para         := Null;
  Else
    n_预约接收效期 := To_Number(Substr(v_Para, 1, n_Count - 1));
    v_Para         := Substr(v_Para, n_Count + 1);
  End If;

  n_预约接诊效期 := 0;
  If v_Para Is Not Null Then
    n_Count := Instr(v_Para, '|');
    If n_Count = 0 Then
      n_预约接诊效期 := To_Number(Nvl(v_Para, '0'));
      v_Para         := Null;
    Else
      n_预约接诊效期 := To_Number(Substr(v_Para, 1, n_Count - 1));
      v_Para         := Substr(v_Para, n_Count + 1);
    End If;
  End If;

  n_预约退号效期 := 0;
  If v_Para Is Not Null Then
    n_Count := Instr(v_Para, '|');
    If n_Count = 0 Then
      n_预约退号效期 := To_Number(Nvl(v_Para, '0'));
      v_Para         := Null;
    Else
      n_预约退号效期 := To_Number(Substr(v_Para, 1, n_Count - 1));
    End If;
  End If;

  n_预约接收效期 := -1 * Nvl(n_预约接收效期, 0);
  n_预约接诊效期 := Nvl(n_预约接诊效期, 0);
  n_预约退号效期 := Nvl(n_预约退号效期, 0);

  If n_预约接收效期 = 0 And n_预约接诊效期 = 0 And n_预约退号效期 = 0 Then
    Return;
  End If;

  Json_Out := '{"output":{"code":1,"message":"成功","black_list":[';
  For I In 1 .. l_不良信息.Count Loop
    For c_预约 In (Select Distinct a.No, a.病人id, a.记录性质, a.记录状态, Nvl(a.预约时间, a.发生时间) As 预约时间, c.名称 As 部门名称, 执行人, 接收时间
                 From 病人挂号记录 A, 部门表 C
                 Where a.执行部门id = c.Id(+) And ((a.病人id = n_病人id And Nvl(n_病人id, 0) <> 0) Or Nvl(n_病人id, 0) = 0) And
                       a.预约 = 1 And
                       ((a.记录性质 = 2 And Nvl(a.记录状态, 0) = 1 And
                       ((a.预约时间 + n_预约接收效期 * (1 / 24 / 60)) <= Sysdate And n_预约接收效期 <> 0)) Or
                       (a.记录性质 = 1 And Nvl(a.记录状态, 0) = 1 And
                       ((Nvl(a.执行时间, Sysdate) - Nvl(a.预约时间, Sysdate)) * 24 * 60 >= n_预约接诊效期) And n_预约接诊效期 <> 0) Or
                       (a.记录状态 = 2 And ((a.登记时间 - Nvl(a.发生时间, Sysdate)) * 24 * 60 >= n_预约退号效期) And n_预约退号效期 <> 0)) And
                       ((a.发生时间 + 0 >= d_最后预约时间 And d_最后预约时间 Is Not Null) Or d_最后预约时间 Is Null) And Not Exists
                  (Select 1
                        From (Select Distinct To_Number(C1) As 病人id, C2 As 附加标志
                               From Table(f_Str2list2(l_不良信息(I)))) B
                        Where a.No = b.附加标志 And a.病人id = b.病人id) And
                       ((a.预约时间 >= Trunc(d_计算日期) And a.预约时间 <= d_计算日期 And Nvl(n_病人id, 0) = 0) Or Nvl(n_病人id, 0) <> 0)) Loop
      If v_操作人员 Is Null Then
        v_操作人员 := Zl_Username;
      End If;
      v_Para := '在' || To_Char(c_预约.预约时间, 'yyyy-mm-dd hh24:mi:ss');
      v_Para := v_Para || '预约的"' || c_预约.部门名称 || '"科室';
    
      If c_预约.执行人 Is Not Null Then
        v_Para := v_Para || '、医生为"' || c_预约.执行人 || '"';
      End If;
      v_Para := v_Para || '(预约单:' || c_预约.No || Case
                  When c_预约.记录状态 = 2 Then
                   '发生退号'
                  When c_预约.记录性质 = 1 Then
                   ' 发生超期接诊'
                  Else
                   ''
                End || ')的号源，未按时就诊。';
      Zljsonputvalue(v_List, 'pati_id', c_预约.病人id, 1, 1);
      Zljsonputvalue(v_List, 'operat_type', '预约挂号');
      Zljsonputvalue(v_List, 'order_date', To_Char(c_预约.预约时间, 'yyyy-mm-dd hh24:mi:ss'));
      Zljsonputvalue(v_List, 'in_reason', '预约超期');
      Zljsonputvalue(v_List, 'in_explain', Zljsonstr(v_Para, 0), 0);
      Zljsonputvalue(v_List, 'sign', c_预约.No);
      Zljsonputvalue(v_List, 'create_name', v_操作人员, 0, 2);
      If Length(v_List) > 20000 Then
        Json_Out := Json_Out || v_List;
        v_List   := ',';
      End If;
    End Loop;
    If v_List <> ',' Then
      Json_Out := Json_Out || v_List || ']}}';
    Else
      Json_Out := Json_Out || ']}}';
    End If;
  End Loop;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Patiregistblacklist;
/

Create Or Replace Procedure Zl_Cissvr_Setautocalcsign
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  -------------------------------------------------------------------------------------------------
  --功能：更新自动记帐标记
  --入参：json格式
  --Input
  --   pati_id               N  1 病人id
  --   pati_pageids          C  1 主页id,多个用英文逗号分隔
  --   auto_account_sign     N  1 自动记帐标志：0-允许自动记帐,1-禁止自动记帐
  --出参：json格式
  --Json_Out
  --   code                  C  1  应答码：0-失败；1-成功
  --   message               C  1  应答消息： 失败时返回具体的错误信息
  -------------------------------------------------------------------------------------------------
  j_In   Pljson;
  j_Json Pljson;

  n_病人id       病案主页.病人id%Type;
  v_主页ids      Varchar2(32767);
  n_禁止自动记帐 病案主页.是否禁止自动记帐%Type;
Begin
  --解析入参
  j_In           := Pljson(Json_In);
  j_Json         := j_In.Get_Pljson('input');
  n_病人id       := j_Json.Get_Number('pati_id');
  v_主页ids      := j_Json.Get_String('pati_pageids');
  n_禁止自动记帐 := j_Json.Get_Number('auto_account_sign');

  Update 病案主页
  Set 是否禁止自动记帐 = Nvl(n_禁止自动记帐, 0)
  Where 病人id = n_病人id And 主页id In (Select Column_Value From Table(f_Num2list(v_主页ids))) And 病人性质 = 0;

  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Setautocalcsign;
/
Create Or Replace Procedure Zl_Cissvr_Setbedempty
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --功能：删除病人床位状况记录
  --入参：Json_In:格式
  --input
  --  pati_id           N    1 病人id
  --出参      json
  --output
  --    code                N   1 应答吗：0-失败；1-成功
  --    message             C   1 应答消息：失败时返回具体的错误信息
  -------------------------------------------
  j_In     Pljson;
  j_Json   Pljson;
  n_病人id Number(18);
Begin
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');
  Update 床位状况记录 Set 状态 = '空床', 病人id = Null, 科室id = Decode(共用, 1, Null, 科室id) Where 病人id = n_病人id;
  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Setbedempty;
/
Create Or Replace Procedure Zl_Cissvr_Updadvicechargetag
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：更改病人医嘱数据的计费状态及删除附费数据
  --入参：Json_In:格式
  --input
  --  item_list[]
  --    advice_id               N   1 医嘱ID
  --    send_no                 N   1 发送号
  --    bill_no                 C   1 单据号
  --    bill_prop               N   1 记录性质
  --    del_annex               N   0 是否删除对应的医嘱附费数据:1-删除;0-不删除
  --    charge_status           N   1 更新的计费状态:-1-无须计费(通常无执行和院外执行的都无须计费);0-未计费;1-已计费(或记帐)，2-部分收费/退费(记帐/销帐)，3-全部收费(仅门诊收费有此项)，4-全部退费(销帐)
  --  fee_detail_list[]                费用明细，仅门诊退费时传入
  --    advice_id               N   1 医嘱ID
  --    bill_no                 C   1 单据号
  --    bill_prop               N   1 记录性质
  --    fee_item_id             N   1 收费细目ID
  --出参: Json_Out,格式如下
  --output
  --   code                     N   1   应答吗：0-失败；1-成功
  --   message                  C   1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Json    Pljson;
  j_Json_In Pljson;

  j_Jsonlist_In Pljson_List;
  j_Detail      Pljson_List;
  v_No          病人医嘱发送.No%Type;
  n_发送号      病人医嘱发送.发送号%Type;
  n_记录性质    病人医嘱发送.记录性质%Type;
  n_医嘱id      病人医嘱发送.医嘱id%Type;
  n_计费状态    病人医嘱发送.计费状态%Type;
  n_删除附费    Number;
  n_收费细目id  医嘱执行计价.收费细目id%Type;
  n_主医嘱id    病人医嘱发送.医嘱id%Type;
Begin
  j_Json        := Pljson(Json_In);
  j_Json_In     := j_Json.Get_Pljson('input');
  j_Jsonlist_In := j_Json_In.Get_Pljson_List('item_list');
  j_Detail      := j_Json_In.Get_Pljson_List('fee_detail_list');

  For I In 1 .. j_Jsonlist_In.Count Loop
    j_Json     := Pljson();
    j_Json     := Pljson(j_Jsonlist_In.Get(I));
    n_医嘱id   := j_Json.Get_Number('advice_id');
    n_发送号   := j_Json.Get_Number('send_no');
    v_No       := j_Json.Get_String('bill_no');
    n_记录性质 := j_Json.Get_Number('bill_prop');
    n_删除附费 := j_Json.Get_Number('del_annex');
    n_计费状态 := j_Json.Get_Number('charge_status');
  
    If Nvl(n_删除附费, 0) = 1 Then
      Delete From 病人医嘱附费 Where 医嘱id = n_医嘱id And 记录性质 = n_记录性质 And NO = v_No;
    End If;
  
    If Nvl(n_发送号, 0) <> 0 And Nvl(n_医嘱id, 0) <> 0 And v_No Is Null Then
      --检查和手术医嘱需要同步更新主医嘱的计费状态
      Select Max(ID)
      Into n_主医嘱id
      From 病人医嘱记录
      Where ID In (Select Max(Nvl(相关id, 0)) From 病人医嘱记录 Where ID = n_医嘱id) And Instr(',D,F,', ',' || 诊疗类别 || ',') > 0;
    
      Update 病人医嘱发送
      Set 计费状态 = n_计费状态
      Where (医嘱id = n_医嘱id Or 医嘱id = Nvl(n_主医嘱id, 0)) And 发送号 = n_发送号;
    Else
      Update 病人医嘱发送 A
      Set a.计费状态 = n_计费状态
      Where 医嘱id = n_医嘱id And 记录性质 = n_记录性质 And NO = v_No;
    End If;
  End Loop;

  If j_Detail Is Not Null Then
    For I In 1 .. j_Detail.Count Loop
      j_Json       := Pljson();
      j_Json       := Pljson(j_Jsonlist_In.Get(I));
      n_医嘱id     := j_Json.Get_Number('advice_id');
      v_No         := j_Json.Get_String('bill_no');
      n_记录性质   := j_Json.Get_Number('bill_prop');
      n_收费细目id := j_Json.Get_Number('fee_item_id');
    
      Update 医嘱执行计价
      Set 执行状态 = 2
      Where (医嘱id, 发送号) In (Select 医嘱id, 发送号 From 病人医嘱发送 Where NO = v_No And 记录性质 = Nvl(n_记录性质, 0)) And
            收费细目id = n_收费细目id And 执行状态 = 0;
    End Loop;
  End If;

  Json_Out := Zljsonout('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Updadvicechargetag;
/
Create Or Replace Procedure Zl_Cissvr_Updadviceexestatus
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：更新医嘱发送的执行状态
  --入参：Json_In:格式
  --  input
  --    item_list[]
  --      advice_id              N  1 医嘱ID
  --      bill_no                C  1 单据号
  --      bill_prop              N  1 记录性质
  --      exe_status_old         C    执行状态:多个用逗号
  --      exe_status             N  1 更新的执行状态
  --      exetr                  C  1 更新的执行人
  --      exe_time               C  1 更新的执行时间:yyyy-mm-dd hh24:mi:ss
  --出参: Json_Out,格式如下
  --  output
  --    code                      N 1 应答吗：0-失败；1-成功
  --    message                   C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_In   Pljson;
  j_Json Pljson;
  j_List Pljson_List;

  n_医嘱id   病人医嘱发送.医嘱id%Type;
  n_执行状态 病人医嘱发送.执行状态%Type;
  n_记录性质 病人医嘱发送.记录性质%Type;
  v_执行状态 Varchar2(100);
  v_执行人   病人医嘱发送.完成人%Type;
  d_执行时间 病人医嘱发送.完成时间%Type;
  v_No       病人医嘱发送.No%Type;
  n_Count    Number(18);

Begin
  --解析入参
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');
  j_List := j_Json.Get_Pljson_List('item_list');

  If Not j_List Is Null Then
    n_Count := j_List.Count;
    For I In 1 .. n_Count Loop
      j_Json     := Pljson();
      j_Json     := Pljson(j_List.Get(I));
      n_医嘱id   := j_Json.Get_Number('advice_id');
      v_No       := j_Json.Get_String('bill_no');
      n_执行状态 := j_Json.Get_Number('exe_status');
      n_记录性质 := j_Json.Get_Number('bill_prop');
      v_执行状态 := j_Json.Get_String('exe_status_old');
      v_执行人   := j_Json.Get_String('exetr');
      d_执行时间 := To_Date(j_Json.Get_String('exe_time'), 'yyyy-mm-dd hh24:mi:ss');
    
      If Nvl(n_医嘱id, 0) = 0 Then
        Json_Out := Zljsonout('未传入医嘱id或发送号，请检查！');
        Return;
      End If;
    
      If Nvl(n_记录性质, 0) = 0 Then
        Json_Out := Zljsonout('未传入记录性质，请检查！');
        Return;
      End If;
    
      If Nvl(v_执行状态, '-') = '-' Then
        Json_Out := Zljsonout('未传入执行状态，请检查！');
        Return;
      End If;
    
      If Nvl(v_No, '-') = '-' Then
        Json_Out := Zljsonout('未传入NO，请检查！');
        Return;
      End If;
    
      Update 病人医嘱发送
      Set 执行状态 = Nvl(n_执行状态, 0), 完成人 = v_执行人, 完成时间 = d_执行时间
      Where 医嘱id = n_医嘱id And NO = v_No And 记录性质 = n_记录性质 And Instr(',' || v_执行状态 || ',', ',' || 执行状态 || ',') > 0;
    End Loop;
  End If;

  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Updadviceexestatus;
/
Create Or Replace Procedure Zl_Cissvr_Updatecritical
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is

  --------------------------------------------------------------------------------------------------------------------
  --功能：危急值相关操作，新增，修改，处理,删除
  --入参：Json_In,格式如下
  --  input
  --       business_type           N  1   业务类型:1-新增,2-修改,3-处理,4-删除

  --       cvalue_id               N  1    危急值id
  --       cvitem_source           C  1    数据来源
  --       pati_id                 N  1    病人id
  --       pati_pageid             N  1    主页id
  --       rgst_no                 C  1    挂号单
  --       baby_num                N  1    婴儿
  --       pat_name                C  1    病人姓名
  --       pat_sex                 C  1    病人性别
  --       pat_age                 C  1    病人年龄
  --       advice_id               N  1    医嘱id
  --       lspcm_id                N  1    标本id
  --       cvalue_rec_desc         C  1    危急值说明
  --       cvalue_rec_create_time  C  1    报告时间
  --       rpt_deptid              N  1    报告科室id
  --       rec_rptor               C  1    报告人

  --       proc_note               C  1    处理情况
  --       cvalue_cnfmtime         C  1    确认时间
  --       cvalue_cnfmer           C  1    确认人
  --       cvalue_deptid           N  1    确认科室id
  --       cvitem_result           N  1    是否危急值

  --出参: Json_Out,格式如下
  --  output
  --    code                            N 1 应答吗：0-失败；1-成功
  --    message                         C 1 应答消息：失败时返回具体的错误信息
  --------------------------------------------------------------------------------------------------------------------
  n_Type Number(5); --1-新增,2-修改,3-处理

  n_Id         病人危急值记录.Id%Type;
  v_数据来源   病人危急值记录.数据来源%Type;
  n_病人id     病人危急值记录.病人id%Type;
  n_主页id     病人危急值记录.主页id%Type;
  v_挂号单     病人危急值记录.挂号单%Type;
  n_婴儿       病人危急值记录.婴儿%Type;
  v_姓名       病人危急值记录.姓名%Type;
  v_性别       病人危急值记录.性别%Type;
  v_年龄       病人危急值记录.年龄%Type;
  n_医嘱id     病人危急值记录.医嘱id%Type;
  n_标本id     病人危急值记录.标本id%Type;
  v_危急值描述 病人危急值记录.危急值描述%Type;
  v_报告时间   病人危急值记录.报告时间%Type;
  n_报告科室id 病人危急值记录.报告科室id%Type;
  v_报告人     病人危急值记录.报告人%Type;

  v_处理情况   病人危急值记录.处理情况%Type;
  v_确认时间   病人危急值记录.确认时间%Type;
  v_确认人     病人危急值记录.确认人%Type;
  n_确认科室id 病人危急值记录.确认科室id%Type;
  n_是否危急值 病人危急值记录.是否危急值%Type;

  j_In   Pljson;
  j_Json Pljson;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin

  --解析入参
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');

  n_Type := Nvl(j_Json.Get_Number('business_type'), 0);
  If n_Type = 1 Or n_Type = 2 Then
    n_Id         := j_Json.Get_Number('cvalue_id');
    v_数据来源   := j_Json.Get_String('cvitem_source');
    n_病人id     := j_Json.Get_Number('pati_id');
    n_主页id     := j_Json.Get_Number('pati_pageid');
    v_挂号单     := j_Json.Get_String('rgst_no');
    n_婴儿       := j_Json.Get_Number('baby_num');
    v_姓名       := j_Json.Get_String('pat_name');
    v_性别       := j_Json.Get_String('pat_sex');
    v_年龄       := j_Json.Get_String('pat_age');
    n_医嘱id     := j_Json.Get_Number('advice_id');
    n_标本id     := j_Json.Get_Number('lspcm_id');
    v_危急值描述 := j_Json.Get_String('cvalue_rec_desc');
    v_报告时间   := To_Date(j_Json.Get_String('cvalue_rec_create_time'), 'yyyy-MM-dd HH24:MI:SS');
    n_报告科室id := j_Json.Get_Number('rpt_deptid');
    v_报告人     := j_Json.Get_String('rec_rptor');
  
    If n_Type = 1 Then
      Zl_病人危急值记录_Insert(n_Id, v_数据来源, n_病人id, n_主页id, v_挂号单, n_婴儿, v_姓名, v_性别, v_年龄, n_医嘱id, n_标本id, v_危急值描述, v_报告时间,
                        n_报告科室id, v_报告人);
    Else
      Zl_病人危急值记录_Update(n_Id, v_数据来源, n_病人id, n_主页id, v_挂号单, n_婴儿, v_姓名, v_性别, v_年龄, n_医嘱id, n_标本id, v_危急值描述, v_报告时间,
                        n_报告科室id, v_报告人);
    End If;
  Elsif n_Type = 3 Then
    n_Id         := j_Json.Get_Number('cvalue_id');
    v_处理情况   := j_Json.Get_String('proc_note');
    v_确认时间   := To_Date(j_Json.Get_String('cvalue_cnfmtime'), 'yyyy-MM-dd HH24:MI:SS');
    v_确认人     := j_Json.Get_String('cvalue_cnfmer');
    n_确认科室id := j_Json.Get_Number('cvalue_deptid');
    n_是否危急值 := j_Json.Get_Number('cvitem_result');
    Zl_病人危急值记录_处理(n_Id, v_处理情况, v_确认时间, v_确认人, n_确认科室id, n_是否危急值);
  Elsif n_Type = 4 Then
    n_Id := j_Json.Get_Number('cvalue_id');
    Zl_病人危急值记录_Delete(n_Id);
  End If;

  Json_Out := Zljsonout('成功', 1);
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Updatecritical;
/
Create Or Replace Procedure Zl_Cissvr_Updatediaginfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：更新病人诊断信息
  --入参：Json_In:格式
  --  input
  --     pati_id            N 1 病人id
  --     pati_pageid        N 1 主页id
  --     diag_types         C 1 诊断类型:0-所有类型,1-西医门诊诊断;2-西医入院诊断;3-西医出院诊断;5-院内感染;6-病理诊断;7-损伤中毒码,8-术前诊断;9-术后诊断;10-并发症;11-中医门诊诊断;12-中医入院诊断;
  --                                     13-中医出院诊断;21-病原学诊断;22-影像学诊断.可以为多个诊断类型，用逗号分离,如:2,12
  --     diag_num           N 1 诊断次序
  --     rec_source         N 1 记录来源:1-病历；2-入院登记；3-首页整理;4-病案
  --     diag_note          C 1 诊断描述
  --出参: Json_Out,格式如下
  --  output
  --    code          N 1 应答吗：0-失败；1-成功
  --    message       C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_In       Pljson;
  j_Json     Pljson;
  n_病人id   病人诊断记录.病人id%Type;
  n_主页id   病人诊断记录.主页id%Type;
  n_记录来源 病人诊断记录.记录来源%Type;
  n_诊断次序 病人诊断记录.诊断次序%Type;
  v_诊断类型 病人诊断记录.诊断类型%Type;
  v_诊断描述 病人诊断记录.诊断描述%Type;

Begin
  --解析入参
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  n_病人id   := j_Json.Get_Number('pati_id');
  n_主页id   := j_Json.Get_Number('pati_pageid');
  n_诊断次序 := j_Json.Get_Number('diag_num');
  n_记录来源 := j_Json.Get_Number('rec_source');
  v_诊断类型 := j_Json.Get_String('diag_types');
  v_诊断描述 := j_Json.Get_String('diag_note');

  Update 病人诊断记录
  Set 诊断描述 = v_诊断描述
  Where 记录来源 = n_记录来源 And (Instr(',' || v_诊断类型 || ',', ',' || 诊断类型 || ',') > 0 Or Nvl(v_诊断类型, 0) = 0) And 诊断次序 = n_诊断次序 And
        病人id = n_病人id And 主页id = n_主页id;
  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Updatediaginfo;
/
Create Or Replace Procedure Zl_Cissvr_Updateinfectinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  --------------------------------------------------------------------------------------------------------------------
  --功能：疾病阳性记录相关操作，新增，修改,删除
  --入参：Json_In,格式如下
  --  input
  --       business_type           N   1   业务类型:1-新增,2-修改,3-删除

  --       rec_id                   N   1  疾病id
  --       pati_id                 N   0  病人id
  --       pati_pageid             N   0  主页id
  --       reg_no                  C   0  挂号单
  --       advice_id               N   0  医嘱id
  --       spcm_send_time          C   0  送检时间
  --       spcm_send_deptid        N   0  送检科室ID
  --       spcm_send_dr            C   0  送检医生
  --       spcm_name               C   0  标本名称
  --       send_content            C   0  反馈结果
  --       infctdz_name            C   0  传染病名称
  --       eqpmtn_exetime          C   0  检查时间
  --       create_time             C   0  登记时间
  --       create_dr               C   0  登记医生
  --       create_dept_id          N   0  登记科室ID
  --       spcm_rec_status         N   0  记录状态

  --       operate_type            N   0  修改操作类型：1-设置处理说明 ,2-关联报告单和阳性结果反馈单,3-取消报告单和阳性结果反馈单的关联,4-修改阳性结果反馈单
  --       spcm_procor             C   0  处理人
  --       spcm_proctime           C   0  处理时间
  --       spcm_procdesc           C   0  处理情况说明
  --       emr_doc_id              N   0  文件ID

  --出参: Json_Out,格式如下
  --  output
  --    code                            N 1 应答吗：0-失败；1-成功
  --    message                         C 1 应答消息：失败时返回具体的错误信息
  --------------------------------------------------------------------------------------------------------------------
  n_Type Number(5); --1-新增,2-修改,3-删除

  n_疾病id     疾病阳性记录.Id%Type;
  n_病人id     疾病阳性记录.病人id%Type;
  n_主页id     疾病阳性记录.主页id%Type;
  v_挂号单     疾病阳性记录.挂号单%Type;
  n_医嘱id     疾病阳性记录.医嘱id%Type;
  d_送检时间   疾病阳性记录.送检时间%Type;
  n_送检科室id 疾病阳性记录.送检科室id%Type;
  v_送检医生   疾病阳性记录.送检医生%Type;
  v_标本名称   疾病阳性记录.标本名称%Type;
  v_反馈结果   疾病阳性记录.反馈结果%Type;
  v_传染病名称 疾病阳性记录.传染病名称%Type;
  d_检查时间   疾病阳性记录.检查时间%Type;
  d_登记时间   疾病阳性记录.登记时间%Type;
  v_登记医生   疾病阳性记录.登记人%Type;
  n_登记科室id 疾病阳性记录.登记科室id%Type;
  n_记录状态   疾病阳性记录.记录状态%Type;

  n_修改操作类型 Number(5);

  v_处理人       疾病阳性记录.处理人%Type;
  d_处理时间     疾病阳性记录.处理时间%Type;
  v_处理情况说明 疾病阳性记录.处理情况说明%Type;
  n_文件id       疾病阳性记录.文件id%Type;

  j_In    Pljson;
  j_Json  Pljson;
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin

  --解析入参
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');

  n_Type := Nvl(j_Json.Get_Number('business_type'), 0);
  If n_Type = 1 Then
    n_疾病id     := j_Json.Get_Number('rec_id');
    n_病人id     := j_Json.Get_Number('pati_id');
    n_主页id     := j_Json.Get_Number('pati_pageid');
    v_挂号单     := j_Json.Get_String('reg_no');
    n_医嘱id     := j_Json.Get_Number('advice_id');
    d_送检时间   := To_Date(j_Json.Get_String('spcm_send_time'), 'yyyy-MM-dd HH24:MI:SS');
    n_送检科室id := j_Json.Get_Number('spcm_send_deptid');
    v_送检医生   := j_Json.Get_String('spcm_send_dr');
    v_标本名称   := j_Json.Get_String('spcm_name');
    v_反馈结果   := j_Json.Get_String('send_content');
    v_传染病名称 := j_Json.Get_String('infctdz_name');
    d_检查时间   := To_Date(j_Json.Get_String('eqpmtn_exetime'), 'yyyy-MM-dd HH24:MI:SS');
    d_登记时间   := To_Date(j_Json.Get_String('create_time'), 'yyyy-MM-dd HH24:MI:SS');
    v_登记医生   := j_Json.Get_String('create_dr');
    n_登记科室id := j_Json.Get_Number('create_dept_id');
    n_记录状态   := j_Json.Get_Number('spcm_rec_status');
  
    Zl_疾病阳性检测记录_Insert(n_疾病id, n_病人id, n_主页id, v_挂号单, n_医嘱id, d_送检时间, n_送检科室id, v_送检医生, v_标本名称, v_反馈结果, v_传染病名称, d_检查时间,
                       d_登记时间, v_登记医生, n_登记科室id, n_记录状态);
  Elsif n_Type = 2 Then
    n_修改操作类型 := Nvl(j_Json.Get_Number('operate_type'), 0);
    n_疾病id       := j_Json.Get_Number('rec_id');
    n_文件id       := j_Json.Get_Number('emr_doc_id');
    n_记录状态     := j_Json.Get_Number('spcm_rec_status');
    v_处理人       := j_Json.Get_String('spcm_procor');
    d_处理时间     := To_Date(j_Json.Get_String('spcm_proctime'), 'yyyy-MM-dd HH24:MI:SS');
    v_处理情况说明 := j_Json.Get_String('spcm_procdesc');
  
    d_送检时间   := To_Date(j_Json.Get_String('spcm_send_time'), 'yyyy-MM-dd HH24:MI:SS');
    n_送检科室id := j_Json.Get_Number('spcm_send_deptid');
    v_送检医生   := j_Json.Get_String('spcm_send_dr');
    v_标本名称   := j_Json.Get_String('spcm_name');
    v_反馈结果   := j_Json.Get_String('send_content');
    v_传染病名称 := j_Json.Get_String('infctdz_name');
    d_检查时间   := To_Date(j_Json.Get_String('eqpmtn_exetime'), 'yyyy-MM-dd HH24:MI:SS');
    d_登记时间   := To_Date(j_Json.Get_String('create_time'), 'yyyy-MM-dd HH24:MI:SS');
    v_登记医生   := j_Json.Get_String('create_dr');
    n_登记科室id := j_Json.Get_Number('create_dept_id');
  
    Zl_疾病阳性检测记录_Update(n_修改操作类型, n_疾病id, n_文件id, n_记录状态, v_处理人, d_处理时间, v_处理情况说明, d_送检时间, n_送检科室id, v_送检医生, v_标本名称,
                       v_反馈结果, v_传染病名称, d_检查时间, d_登记时间, v_登记医生, n_登记科室id);
  Elsif n_Type = 3 Then
    n_疾病id := j_Json.Get_Number('rec_id');
    Zl_疾病阳性记录_Delete(n_疾病id);
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Updateinfectinfo;
/
Create Or Replace Procedure Zl_Cissvr_Updateinpatiextinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:修改病案主页从表相关信息
  --入参：Json_In:格式
  --    input
  --    pati_id             N  1  病人id
  --    pati_pageid         N  1  主页Id
  --    item_list[]               列表
  --      info_name         C  1  信息名
  --      info_value        C  1  修改的信息值
  --出参: Json_Out,格式如下
  --  output
  --    code                N   1   应答码：0-失败；1-成功
  --    message             C   1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------

  j_Json     Pljson;
  o_Json     Pljson;
  j_Jsonlist Pljson_List := Pljson_List();

  n_病人id 病案主页从表.病人id%Type;
  n_主页id 病案主页从表.主页id%Type;
  v_信息名 病案主页从表.信息名%Type;
  v_信息值 病案主页从表.信息值%Type;
  Err_Item Exception;
Begin
  --解析入参
  o_Json     := Pljson(Json_In);
  j_Json     := o_Json.Get_Pljson('input');
  n_病人id   := j_Json.Get_Number('pati_id');
  n_主页id   := j_Json.Get_Number('pati_pageid');
  j_Jsonlist := j_Json.Get_Pljson_List('item_list');
  If j_Jsonlist Is Null Then
    Json_Out := Zljsonout('不存在需要保存的从项项目信息', 0);
    Return;
  End If;
  For I In 1 .. j_Jsonlist.Count Loop
    o_Json   := Pljson();
    o_Json   := Pljson(j_Jsonlist.Get(I));
    v_信息名 := o_Json.Get_String('info_name');
    v_信息值 := o_Json.Get_String('info_value');
    Zl_病案主页从表_首页整理(n_病人id, n_主页id, v_信息名, v_信息值);
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Updateinpatiextinfo;
/
Create Or Replace Procedure Zl_Cissvr_Updateinpatipageinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------------------------------------------------------------
  --功能：更新病案主页和病案主页从表信息
  --入参：json格式
  --Input
  --   pati_id               N  1 病人id
  --   pati_pageid           N  1 主页id
  --   inpatient_num         C  1 住院号
  --   fee_type              C  1 费别
  --   inpati_fee_type       C  1 住院费别
  --   mdlpay_mode_name      C  1 医疗付款方式
  --   pati_area             C  1 区域
  --   remarkes              C  1 备注
  --   pati_marital_cstatus  C  1 婚姻状况
  --   pati_education        C  1 学历
  --   ocpt_name             C  1 职业
  --   emp_phno              C  1 单位电话
  --   emp_postcode          C  1 单位邮编
  --   pati_home_addr         C  1 家庭地址
  --   pati_home_phno         C  1 家庭电话
  --   pati_home_postcode     C  1 家庭地址邮编
  --   pati_hous_addr         C  1 户口地址
  --   pati_hous_postcode     C  1 户口地址邮编
  --   contacts_name         C  1 联系人姓名
  --   contacts_relation     C  1 联系人关系
  --   contacts_addr         C  1 联系人地址
  --   contacts_phno         C  1 联系人电话
  --   pati_type             C  1 病人类型
  --   insurance_num         C  1 医保号
  --   outpatient_num        N  1 门诊号
  --   item_list[]              1 病案主页从表信息
  --   regist_time           C  1 登记时间
  --   opr_type              N  1 执行方式 0-新增 1-修改
  --出参：json格式
  --Json_Out
  --   code                  C  1  应答码：0-失败；1-成功
  --   message               C  1  应答消息：失败时返回具体的错误信息
  -------------------------------------------------------------------------------------------------

  j_In       Pljson;
  j_Json     Pljson;
  j_Jsonlist Pljson_List := Pljson_List();
  o_Json     Pljson;

  n_病人id       病案主页.病人id%Type;
  n_主页id       病案主页.主页id%Type;
  n_住院号       病案主页.住院号%Type;
  v_费别         病案主页.费别%Type;
  v_住院费别     病案主页.费别%Type;
  v_医疗付款方式 病案主页.医疗付款方式%Type;
  v_区域         病案主页.区域%Type;
  v_备注         病案主页.备注%Type;
  v_婚姻状况     病案主页.婚姻状况%Type;
  v_学历         病案主页.学历%Type;
  v_职业         病案主页.职业%Type;
  v_单位电话     病案主页.单位电话%Type;
  v_单位邮编     病案主页.单位邮编%Type;
  v_家庭地址     病案主页.家庭地址%Type;
  v_家庭电话     病案主页.家庭电话%Type;
  v_家庭地址邮编 病案主页.家庭地址邮编%Type;
  v_户口地址     病案主页.户口地址%Type;
  v_户口地址邮编 病案主页.户口地址邮编%Type;
  v_联系人姓名   病案主页.联系人姓名%Type;
  v_联系人关系   病案主页.联系人关系%Type;
  v_联系人地址   病案主页.联系人地址%Type;
  v_联系人电话   病案主页.联系人电话%Type;
  v_病人类型     病案主页.病人类型%Type;
  v_出院日期     病案主页.出院日期%Type;
  v_医保号       病案主页从表.信息值%Type;
  n_门诊号       门诊病案记录.病案号%Type;
  v_信息名       病案主页从表.信息名%Type;
  v_信息值       病案主页从表.信息值%Type;
  d_登记时间     Date;
  n_Type         Number;
Begin
  --解析入参
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');

  n_病人id       := j_Json.Get_Number('pati_id');
  n_主页id       := j_Json.Get_Number('pati_pageid');
  n_住院号       := To_Number(j_Json.Get_String('inpatient_num'));
  v_费别         := j_Json.Get_String('fee_type');
  v_住院费别     := j_Json.Get_String('inpati_fee_type');
  v_医疗付款方式 := j_Json.Get_String('mdlpay_mode_name');
  v_区域         := j_Json.Get_String('pati_area');
  v_备注         := j_Json.Get_String('remarkes');
  v_婚姻状况     := j_Json.Get_String('pati_marital_cstatus');
  v_学历         := j_Json.Get_String('pati_education');
  v_职业         := j_Json.Get_String('ocpt_name');
  v_单位电话     := j_Json.Get_String('emp_phno');
  v_单位邮编     := j_Json.Get_String('emp_postcode');
  v_家庭地址     := j_Json.Get_String('pat_home_addr');
  v_家庭电话     := j_Json.Get_String('pat_home_phno');
  v_家庭地址邮编 := j_Json.Get_String('pat_home_postcode');
  v_户口地址     := j_Json.Get_String('pat_hous_addr');
  v_户口地址邮编 := j_Json.Get_String('pat_hous_postcode');
  v_联系人姓名   := j_Json.Get_String('contacts_name');
  v_联系人关系   := j_Json.Get_String('contacts_relation');
  v_联系人地址   := j_Json.Get_String('contacts_addr');
  v_联系人电话   := j_Json.Get_String('contacts_phno');
  v_病人类型     := j_Json.Get_String('pati_type');
  v_医保号       := j_Json.Get_String('insurance_num');
  n_门诊号       := To_Number(j_Json.Get_Number('outpatient_num'));
  j_Jsonlist     := j_Json.Get_Pljson_List('item_list');
  d_登记时间     := To_Date(j_Json.Get_String('regist_time'), 'yyyy-mm-dd hh24:mi:ss');
  n_Type         := j_Json.Get_Number('opr_type');
  If Nvl(n_Type, 0) = 1 Then
    If n_门诊号 Is Not Null Then
      Update 门诊病案记录 Set 病案号 = n_门诊号 Where 病人id = n_病人id;
      If Sql%RowCount = 0 Then
        Insert Into 门诊病案记录
          (病人id, 病案号, 建立日期, 病案类别, 存储状态, 存放位置)
        Values
          (n_病人id, n_门诊号, Sysdate, '一般', '正常', Null);
      End If;
    Else
      Delete From 门诊病案记录 Where 病人id = n_病人id;
    End If;
  Else
    If n_门诊号 Is Not Null Then
      Insert Into 门诊病案记录
        (病人id, 病案号, 建立日期, 病案类别, 存储状态, 存放位置)
      Values
        (n_病人id, n_门诊号, d_登记时间, '一般', '正常', Null);
    End If;
  End If;

  If n_主页id Is Not Null And n_主页id <> 0 Then
    If Nvl(n_Type, 0) = 1 Then
      Update 病案主页
      Set 住院号 = n_住院号, 费别 = Decode(Nvl(v_住院费别, 0), 1, v_费别, 费别), 医疗付款方式 = v_医疗付款方式, 区域 = Decode(v_区域, Null, 区域, v_区域),
          备注 = v_备注
      Where 病人id = n_病人id And 主页id = n_主页id;
    
      --对在院病人更新病案主页中的信息
      Begin
        Select 出院日期 Into v_出院日期 From 病案主页 Where 病人id = n_病人id And 主页id = n_主页id;
      Exception
        When Others Then
          Null;
      End;
      If v_出院日期 Is Null Then
        Update 病案主页
        Set 婚姻状况 = v_婚姻状况, 学历 = v_学历, 职业 = v_职业, 单位电话 = v_单位电话, 单位邮编 = v_单位邮编, 家庭地址 = v_家庭地址, 家庭电话 = v_家庭电话,
            家庭地址邮编 = v_家庭地址邮编, 户口地址 = Nvl(v_户口地址, v_户口地址), 户口地址邮编 = Nvl(v_户口地址邮编, 户口地址邮编), 联系人姓名 = v_联系人姓名,
            联系人关系 = v_联系人关系, 联系人地址 = v_联系人地址, 联系人电话 = v_联系人电话, 病人类型 = v_病人类型, 备注 = v_备注

        Where 病人id = n_病人id And 主页id = n_主页id;
      End If;
      If v_医保号 Is Not Null Then
        Update 病案主页从表 Set 信息值 = v_医保号 Where 病人id = n_病人id And 主页id = n_主页id And 信息名 = '医保号';
        If Sql%RowCount = 0 Then
          Insert Into 病案主页从表 (病人id, 主页id, 信息名, 信息值) Values (n_病人id, n_主页id, '医保号', v_医保号);
        End If;
      Else
        Delete From 病案主页从表 Where 病人id = n_病人id And 主页id = n_主页id And 信息名 = '医保号';
      End If;
    End If;
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json   := Pljson();
      o_Json   := Pljson(j_Jsonlist.Get(I));
      v_信息名 := o_Json.Get_String('info_name');
      v_信息值 := o_Json.Get_String('info_value');
      Zl_病案主页从表_首页整理(n_病人id, n_主页id, v_信息名, v_信息值);
    End Loop;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Updateinpatipageinfo;
/
Create Or Replace Procedure Zl_Cissvr_Updateoutmedrecord
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  -------------------------------------------------------------------------------------------------
  --功能：更新门诊档案信息
  --入参：json格式
  --Input
  --   pati_id               N  1 病人id
  --   mr_no                 C  1 病案号
  --   outpatient_num        C  1 门诊号
  --   create_date           C  1 建立日期
  --   mr_type               C  1 病案类别
  --   strg_status           C  1 存储状态
  --   strgloc               C  1 存放位置
  --出参：json格式
  --Json_Out
  --   code                  C  1  应答码：0-失败；1-成功
  --   message               C  1  应答消息：失败时返回具体的错误信息
  -------------------------------------------------------------------------------------------------

  n_病人id   门诊病案记录.病人id%Type;
  n_病案号   门诊病案记录.病案号%Type;
  n_门诊号   门诊病案记录.病案号%Type;
  d_建立日期 门诊病案记录.建立日期%Type;
  v_存放位置 门诊病案记录.存放位置%Type;
  v_病案类别 门诊病案记录.病案类别%Type;
  v_存储状态 门诊病案记录.存储状态%Type;
  j_Json     Pljson;
  j_In       Pljson;

Begin
  --解析入参
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');

  n_病人id   := j_Json.Get_Number('pati_id');
  n_病案号   := To_Number(j_Json.Get_String('mr_no'));
  d_建立日期 := To_Date(j_Json.Get_String('create_date'), 'yyyy-mm-dd hh24:mi:ss');
  n_门诊号   := To_Number(j_Json.Get_String('outpatient_num'));
  v_存放位置 := j_Json.Get_String('strgloc');
  v_病案类别 := j_Json.Get_String('mr_type');
  v_存储状态 := j_Json.Get_String('strg_status');
  If Nvl(n_病人id, 0) = 0 Then
    Json_Out := Zljsonout('未传入病人id，请检查！', 0);
    Return;
  End If;
  If n_病案号 Is Not Null Then
    If Nvl(n_病案号, 0) = 0 Then
      Json_Out := Zljsonout('未传入门诊号，请检查！', 0);
      Return;
    End If;
  
    Update 门诊病案记录 Set 病案号 = n_病案号 Where 病人id = n_病人id;
    If Sql%RowCount = 0 Then
      Insert Into 门诊病案记录
        (病人id, 病案号, 建立日期, 病案类别, 存储状态, 存放位置)
      Values
        (n_病人id, n_病案号, d_建立日期, v_病案类别, v_存储状态, v_存放位置);
    End If;
  Else
    If n_门诊号 Is Not Null Then
      Update 门诊病案记录 Set 病案号 = n_门诊号 Where 病人id = n_病人id;
      If Sql%RowCount = 0 Then
        Insert Into 门诊病案记录
          (病人id, 病案号, 建立日期, 病案类别, 存储状态, 存放位置)
        Values
          (n_病人id, n_门诊号, Sysdate, '一般', '正常', Null);
      End If;
    Else
      Delete From 门诊病案记录 Where 病人id = n_病人id;
    End If;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Updateoutmedrecord;
/
Create Or Replace Procedure Zl_Cissvr_Updateoutvitalsigns
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:修改病人的生命体征信息
  --入参：Json_In:格式
  --    input
  --      pati_id           N  1 病人ID
  --      rgst_id           N  1 挂号ID
  --      pat_vsign         N  1 体征。格式为：项目ID1|项目值1|单位1,项目ID2|项目值2|单位2
  --      operator_name     C    操作员姓名

  --出参: Json_Out,格式如下
  --  output
  --    code                N  1   应答码：0-失败；1-成功
  --    message             C  1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------

  j_In   Pljson;
  j_Json Pljson;

  n_病人id     病人护理记录.病人id%Type;
  n_挂号id     病人护理记录.主页id%Type;
  v_操作员姓名 病人护理记录.保存人%Type;
  v_体征       Varchar2(4000);

Begin
  --解析入参
  j_In   := Pljson(Json_In);
  j_Json := j_In.Get_Pljson('input');

  v_体征       := j_Json.Get_String('pat_vsign');
  n_病人id     := j_Json.Get_Number('pati_id');
  n_挂号id     := j_Json.Get_Number('rgst_id');
  v_操作员姓名 := j_Json.Get_String('operator_name');

  If Nvl(n_挂号id, 0) = 0 Or Nvl(n_病人id, 0) = 0 Then
    Json_Out := Zljsonout('未确定病人的就诊信息,请检查!');
    Return;
  End If;

  Zl_门诊生命体征_Update(n_病人id, n_挂号id, v_体征, v_操作员姓名);

  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Updateoutvitalsigns;
/
Create Or Replace Procedure Zl_Cissvr_Updatepatiauditinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------------------------------------------------------------
  --功能：更新病人审核信息
  --入参：json格式
  --Input
  --   pati_id               N  1 病人id
  --   pati_pageid           N  1 主页id
  --   audit_sign            N  1 审核标记：0或空-未审核,1-已审核或开始审核;2-完成审核
  --   auditor               C  0 审核人
  --   audit_desc            C  0 审核说明
  --   cancel_audit          N  0 是否取消审核：1-取消审核,0-审核
  --出参：json格式
  --Json_Out
  --   code                  C  1  应答码：0-失败；1-成功
  --   message               C  1  应答消息： 失败时返回具体的错误信息
  -------------------------------------------------------------------------------------------------
  j_In   Pljson;
  j_Json Pljson;

  n_病人id   病案主页.病人id%Type;
  n_主页id   病案主页.主页id%Type;
  n_审核标记 病案主页.审核标志%Type;
  v_审核人   病案主页.审核人%Type;
  v_审核说明 病案主页.审核说明%Type;
  n_取消审核 Number(2);
Begin
  --解析入参
  j_In       := Pljson(Json_In);
  j_Json     := j_In.Get_Pljson('input');
  n_病人id   := j_Json.Get_Number('pati_id');
  n_主页id   := j_Json.Get_Number('pati_pageid');
  n_审核标记 := j_Json.Get_Number('audit_sign');
  v_审核人   := j_Json.Get_String('auditor');
  v_审核说明 := j_Json.Get_String('audit_desc');
  n_取消审核 := j_Json.Get_Number('cancel_audit');

  Zl_病人审核_Execute(n_病人id, n_主页id, n_审核标记, v_审核人, Nvl(n_取消审核, 0), v_审核说明);

  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Updatepatiauditinfo;
/
Create Or Replace Procedure Zl_Cissvr_Updatepatibaseinfo
(
  Json_in    Varchar2,
  Json_out Out Varchar2
) Is
  --------------------------------------------------------------------------------------------------
  --功能:更新病人临床相关的病人基本信息
  --入参 JSOM格式
  --input
  --  pati_id               N 1 病人id
  --  visit_id              N 1 就诊id
  --  occasion              N 1 场合
  --  update_info[]  更新信息
  --      pati_name             C 1 姓名
  --      pati_age              C 1 年龄
  --      pati_sex              C 1 性别
  --出参 JSON格式
  --output
  --  code                      N 1 应答码：0-失败；1-成功
  --  message                   C 1 应答消息：失败时返回具体的错误信息
  --  adjust_explain            C 1 修改原因
  --------------------------------------------------------------------------------------------------
  o_Json   Pljson;
  j_Json   Pljson;
  j_In     Pljson;
  n_病人id 病案主页.病人id%Type;
  v_姓名   病案主页.姓名%Type;
  v_性别   病案主页.性别%Type;
  v_年龄   病案主页.年龄%Type; --更新前的年龄
  n_就诊id Number;
  n_场合   Number(1);
  v_说明   Varchar2(32676);
  说明_Out Clob;

  Procedure p_病历
  (
    病人id_In 病人挂号记录.病人id%Type,
    就诊id_In Number,
    姓名_In   病人挂号记录.姓名%Type,
    性别_In   病人挂号记录.性别%Type,
    年龄_In   病人挂号记录.年龄%Type,
    场合_In   Number, --1-门诊;2-住院
    说明_Out  Out Varchar2
  ) As
    Err_Custom Exception;
    v_Error Varchar2(2000);
  Begin
    --以要素方式存储的姓名、性别、年龄
    For r_Rec In (Select /*+ RULE */
                  Distinct d.名称 科室, a.Id, a.病历名称, a.病历种类, a.完成时间, a.保存人,
                           Nvl((Select Decode(a.编辑方式, 0, Decode(Substr(b.对象属性, 1, 1), '2', 1, 0),
                                               Decode(Substr(b.对象属性, Instr(b.对象属性, '|'), 1), '2', 1, 0))
                                From 电子病历内容 B
                                Where a.Id = b.文件id And
                                      ((b.对象类型 = 8 And a.编辑方式 = 0) Or (b.对象类型 In (6, 7, 8) And a.编辑方式 = 1)) And
                                      Rownum < 2), 0) As 电子签名
                  From 电子病历记录 A, 部门表 D
                  Where a.病人id = 病人id_In And a.主页id = 就诊id_In And a.科室id = d.Id And Exists
                   (Select 1 --包含
                         From 电子病历内容 C
                         Where a.Id = c.文件id And ((c.对象类型 = 4 And c.要素名称 = '姓名') Or (c.对象类型 = 4 And c.要素名称 = '性别') Or
                               (c.对象类型 = 4 And c.要素名称 = '年龄') Or (c.对象类型 = 2 And c.要素名称 = '姓名') Or
                               (c.对象类型 = 2 And c.要素名称 = '性别') Or (c.对象类型 = 2 And c.要素名称 = '年龄'))) And Rownum < 2) Loop
    
      --读取所有包含姓名、性别、年龄要素的病历
      If r_Rec.电子签名 <> 1 Then
        --返回病历名称串
        说明_Out := r_Rec.科室 || ':书写的病历中包含病人基本信息，需要手工调整。';
        If r_Rec.病历种类 = 5 Then
          --更新疾病申报记录
          Update 疾病申报记录 Set 姓名 = 姓名_In, 性别 = 性别_In, 年龄 = 年龄_In Where 文件id = r_Rec.Id;
        End If;
      End If;
    End Loop;
  
    --以自由录入的姓名
    If Nvl(场合_In, 0) = 1 Then
      For r_Rec In (Select /*+ RULE */
                    Distinct d.名称 科室, a.Id, a.病历名称, a.病历种类, a.完成时间, a.保存人,
                             Nvl((Select Decode(a.编辑方式, 0, Decode(Substr(b.对象属性, 1, 1), '2', 1, 0),
                                                 Decode(Substr(b.对象属性, Instr(b.对象属性, '|'), 1), '2', 1, 0))
                                  From 电子病历内容 B
                                  Where a.Id = b.文件id And
                                        ((b.对象类型 = 8 And 编辑方式 = 0) Or (b.对象类型 In (6, 7, 8) And 编辑方式 = 1)) And Rownum < 2),
                                  0) As 电子签名
                    From 电子病历记录 A, 部门表 D, 病人挂号记录 E
                    Where a.病人id = 病人id_In And a.主页id = 就诊id_In And a.科室id = d.Id And a.病人id = e.病人id And Exists
                     (Select 1 --包含
                           From 电子病历内容 C
                           Where a.Id = c.文件id And
                                 ((c.对象类型 = 2 And Instr(c.内容文本, e.姓名) > 0) Or (c.对象类型 = 1 And Instr(c.内容文本, e.姓名) > 0))) And
                          Rownum < 2) Loop
      
        --读取所有包含姓名的病历
        If r_Rec.电子签名 <> 1 Then
          --返回病历名称串
          说明_Out := r_Rec.科室 || ':书写的病历中包含病人基本信息，需要手工调整。';
          If r_Rec.病历种类 = 5 Then
            --更新疾病申报记录
            Update 疾病申报记录 Set 姓名 = 姓名_In, 性别 = 性别_In, 年龄 = 年龄_In Where 文件id = r_Rec.Id;
          End If;
        End If;
      End Loop;
    Else
      For r_Rec In (Select /*+ RULE */
                    Distinct d.名称 科室, a.Id, a.病历名称, a.病历种类, a.完成时间, a.保存人,
                             Nvl((Select Decode(a.编辑方式, 0, Decode(Substr(b.对象属性, 1, 1), '2', 1, 0),
                                                 Decode(Substr(b.对象属性, Instr(b.对象属性, '|'), 1), '2', 1, 0))
                                  From 电子病历内容 B
                                  Where a.Id = b.文件id And
                                        ((b.对象类型 = 8 And 编辑方式 = 0) Or (b.对象类型 In (6, 7, 8) And 编辑方式 = 1)) And Rownum < 2),
                                  0) As 电子签名
                    From 电子病历记录 A, 部门表 D, 病案主页 E
                    Where a.病人id = 病人id_In And a.主页id = 就诊id_In And a.科室id = d.Id And a.病人id = e.病人id And Exists
                     (Select 1 --包含
                           From 电子病历内容 C
                           Where a.Id = c.文件id And
                                 ((c.对象类型 = 2 And Instr(c.内容文本, e.姓名) > 0) Or (c.对象类型 = 1 And Instr(c.内容文本, e.姓名) > 0))) And
                          Rownum < 2) Loop
      
        --读取所有包含姓名的病历
        If r_Rec.电子签名 <> 1 Then
          --返回病历名称串
          说明_Out := r_Rec.科室 || ':书写的病历中包含病人基本信息，需要手工调整。';
          If r_Rec.病历种类 = 5 Then
            --更新疾病申报记录
            Update 疾病申报记录 Set 姓名 = 姓名_In, 性别 = 性别_In, 年龄 = 年龄_In Where 文件id = r_Rec.Id;
          End If;
        End If;
      End Loop;
    End If;
  
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_病历;

  Procedure p_医嘱
  (
    病人id_In 病人挂号记录.病人id%Type,
    就诊id_In Number,
    姓名_In   病人挂号记录.姓名%Type,
    性别_In   病人挂号记录.性别%Type,
    年龄_In   病人挂号记录.年龄%Type,
    场合_In   Number, --1-门诊;2-住院
    说明_Out  Out Varchar2
  ) As
    ------------------------------------------------------------------------------------------
    --功能:更新医嘱相关业务数据的病人基本信息
    --入参:病人id_In:病人ID
    --     就诊id_In:门诊病人为挂号ID;住院病人为主页ID;为0说明是外来或非挂号就诊的病人(就诊id_In为空时,将批量更改该病人的所有业务数据)
    --     姓名_In:需要更改的病人姓名
    --     性别_In:需要更改的病人性别
    --     年龄_In:需要更改的病人年龄
    --     场合_In:1-门诊;2-住院
    --出参:说明_Out:病人信息调整后的说明信息，用于提示操作员进行相关操作
    ------------------------------------------------------------------------------------------
    Err_Custom Exception;
    v_Error Varchar2(2000);
    n_Count Number(3);
    v_No    病人挂号记录.No%Type;
  Begin
    --外来人员，不处理
    If Nvl(就诊id_In, 0) = 0 Then
      Return;
    End If;
    --门诊取挂号单
    If Nvl(场合_In, 0) = 1 Then
      --更新病人本次就诊的医嘱中的病人基本信息
      Select NO Into v_No From 病人挂号记录 Where ID = 就诊id_In;
    
      Update 病人医嘱记录
      Set 姓名 = Nvl(姓名_In, 姓名), 性别 = Nvl(性别_In, 性别), 年龄 = Nvl(年龄_In, 年龄)
      Where 病人id = 病人id_In And 挂号单 = v_No;
    
      ---更新病人危急值记录
      Update 病人危急值记录
      Set 姓名 = Nvl(姓名_In, 姓名), 性别 = Nvl(性别_In, 性别), 年龄 = Nvl(年龄_In, 年龄)
      Where 病人id = 病人id_In And 挂号单 = v_No;
      Return;
    End If;
    --住院病人
    If Nvl(场合_In, 0) = 2 Then
      --已经打印了医嘱清单的提示重新打印
      Select Nvl(Count(1), 0)
      Into n_Count
      From 病人医嘱打印
      Where 病人id = 病人id_In And 主页id = 就诊id_In And Rownum < 2;
    
      If n_Count <> 0 Then
        If Not 说明_Out Is Null Then
          说明_Out := 说明_Out || Chr(13);
        End If;
        说明_Out := 说明_Out || '医嘱清单:已经打印需重新打印.';
      End If;
    
      --已经打印了首页的提示重新打印
      Select Nvl(Count(1), 0)
      Into n_Count
      From 电子病历打印
      Where 病人id = 病人id_In And 主页id = 就诊id_In And 文件id Is Null And 种类 = 9 And Rownum < 2;
      If n_Count <> 0 Then
        If Not 说明_Out Is Null Then
          说明_Out := 说明_Out || Chr(13);
        End If;
        说明_Out := 说明_Out || '病人首页:已经打印需重新打印.';
      End If;
    
      Update 病人医嘱记录
      Set 姓名 = Nvl(姓名_In, 姓名), 性别 = Nvl(性别_In, 性别), 年龄 = Nvl(年龄_In, 年龄)
      Where 病人id = 病人id_In And 主页id = 就诊id_In;
    
      ---更新病人危急值记录
      Update 病人危急值记录
      Set 姓名 = Nvl(姓名_In, 姓名), 性别 = Nvl(性别_In, 性别), 年龄 = Nvl(年龄_In, 年龄)
      Where 病人id = 病人id_In And 主页id = 就诊id_In;
      Return;
    End If;
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_医嘱;

Begin
  j_In     := Pljson(Json_in);
  j_Json   := j_In.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');
  n_就诊id := j_Json.Get_Number('visit_id');
  n_场合   := j_Json.Get_Number('occasion');
  o_Json   := j_Json.Get_Pljson('update_info');
  If o_Json Is Null Then
    Json_out := Zljsonout('未传入需要更新的信息，请检查！', 0);
    Return;
  End If;
  v_姓名 := o_Json.Get_String('pati_name');
  v_性别 := o_Json.Get_String('pati_sex');
  v_年龄 := o_Json.Get_String('pati_age');

  --医嘱部分
  p_医嘱(n_病人id, n_就诊id, v_姓名, v_性别, v_年龄, n_场合, v_说明);
  If v_说明 Is Not Null Then
    说明_Out := 说明_Out || Chr(13) || '医嘱部分:' || Chr(13) || v_说明;
  End If;
  --病历部分
  v_说明 := '';
  p_病历(n_病人id, n_就诊id, v_姓名, v_性别, v_年龄, n_场合, v_说明);
  If v_说明 Is Not Null Then
    说明_Out := 说明_Out || Chr(13) || '病历部分:' || Chr(13) || v_说明;
  End If;

  --最后完成主表的更新(病案主页、病人挂号记录、病人信息)
  If n_场合 = 1 And Nvl(n_就诊id, 0) <> 0 Then
    --门诊病人
    Update 病人挂号记录 A Set 姓名 = v_姓名, 性别 = v_性别, 年龄 = v_年龄 Where 病人id = n_病人id And a.Id = n_就诊id;
  End If;
  If n_场合 = 2 And Nvl(n_就诊id, 0) <> 0 Then
    --住院病人
    Update 病案主页 Set 姓名 = v_姓名, 性别 = v_性别, 年龄 = v_年龄 Where 病人id = n_病人id And 主页id = n_就诊id;
  End If;
  Json_out := '{"output":{"code":1,"message":"成功","adjust_explain":"' || Zljsonstr(说明_Out) || '"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Updatepatibaseinfo;
/
Create Or Replace Procedure Zl_Cissvr_Updatesyncstate
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：同步标记录更新
  --入参：Json_In:格式
  --  input
  --      order_list[]
  --          order_id          N 1 医嘱id
  --          send_no           N 1 发送号
  --          sign_type         N 1 设置标记录的类型，
  --                                  说明：1-清除静配标记录
  --                                        2-清除 生成药品同步标记
  --                                        3-清除 生成卫材同步标记
  --出参: Json_Out,格式如下
  --  output
  --    code          N 1 应答吗：0-失败；1-成功
  --    message       C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Json       Pljson;
  o_Json       Pljson;
  j_Order_List Pljson_List;
  n_医嘱id     Number;
  n_发送号     Number;
  n_类型       Number;
  n_Count      Number;
Begin
  --解析入参
  o_Json       := Pljson(Json_In);
  j_Json       := o_Json.Get_Pljson('input');
  j_Order_List := j_Json.Get_Pljson_List('order_list');
  n_Count      := j_Order_List.Count;
  If n_Count > 0 Then
    For I In 1 .. n_Count Loop
      o_Json   := Pljson();
      o_Json   := Pljson(j_Order_List.Get(I));
      n_医嘱id := o_Json.Get_Number('order_id');
      n_发送号 := o_Json.Get_Number('send_no');
      n_类型   := o_Json.Get_Number('sign_type');
      --产生环节: 1-生成药品，2-生成卫材，3-生成配液，4-收费确认药品，5-收费确认卫材，6-作废药品，7-作废卫材
      If n_类型 = 1 Then
        Delete 病人医嘱异常记录 Where 医嘱id = n_医嘱id And 发送号 = n_发送号 And 产生环节 In (3, 5);
      Elsif n_类型 = 2 Then
        Delete 病人医嘱异常记录 Where 医嘱id = n_医嘱id And 发送号 = n_发送号 And 产生环节 In (1, 4);
      Elsif n_类型 = 3 Then
        Delete 病人医嘱异常记录 Where 医嘱id = n_医嘱id And 发送号 = n_发送号 And 产生环节 In (2, 5);
      End If;
    End Loop;
  End If;
  Json_Out := '{"output":{"code":1,"message": "成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Updatesyncstate;
/
CREATE OR REPLACE Procedure Zl_Cissvr_Updoutpativisitrec
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------------------------------------------------------------
  --功能：更新病人就诊信息
  --入参：json格式
  --Input
  --   rgst_no	            C	1	挂号单号
  --   pati_id              N   病人ID
  --   outpatient_num       C   门诊号
  --   pati_name            C   姓名
  --   pati_sex             C   性别
  --   pati_age             C   年龄
  --   fee_category         C   费别
  --   exetr	              C	 	执行人
  --   outproom_name	      C	 	诊室
  --   exe_time	            C	 	执行时间
  --   exe_status           N   执行状态
  --   rgst_desc	          C	 	摘要
  --   pnurs_oprtr	        N	 	是否护士执行
  --   outp_recv_time_end   C	 	完成时间

  --出参：json格式
  --Json_Out
  --   code                  C  1  应答码：0-失败；1-成功
  --   message               C  1  应答消息：失败时返回具体的错误信息
  -------------------------------------------------------------------------------------------------

  v_挂号单号 病人挂号记录.No%Type;
  v_执行人   病人挂号记录.执行人%Type;
  v_诊室     病人挂号记录.诊室%Type;
  d_执行时间 病人挂号记录.执行时间%Type;
  v_摘要     病人挂号记录.摘要%Type;
  n_护士执行 Number;
  j_In       Pljson;
  j_Json     Pljson;

Begin
  --解析入参
  j_In := Pljson(Json_In);

  j_Json     := j_In.Get_Pljson('input');
  v_挂号单号 := j_Json.Get_String('rgst_no');
  v_执行人   := j_Json.Get_String('exetr');
  v_诊室     := j_Json.Get_String('outproom_name');
  d_执行时间 := To_Date(j_Json.Get_String('exe_time'), 'YYYY-MM-DD hh24:mi:ss');
  v_摘要     := j_Json.Get_String('rgst_desc');
  n_护士执行 := j_Json.Get_Number('pnurs_oprtr');

  Json_Out := '{"output":{"code":1,"message": "成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Updoutpativisitrec;
/
Create Or Replace Procedure Zl_Cissvr_Updpatbaseinfocheck
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ------------------------------------------------------------------------------------------
  --功能:更新医嘱相关业务数据的病人基本信息的检查
  --入参:JSON格式
  --input
  --   pati_id   N 1 病人id
  --   visit_id   N 1 就诊id ，门诊病人为挂号ID;住院病人为主页ID;为0说明是外来或非挂号就诊的病人(就诊id_In为空时,将批量更改该病人的所有业务数据)
  --   occasion   N 1 场合,场合_In:1-门诊;2-住院
  --出参:JSON格式
  --output
  --  code            N 1 应答码：0-失败；1-成功
  --  message        C 1 应答消息：失败时返回具体的错误信息
  ------------------------------------------------------------------------------------------
  j_In     Pljson;
  j_Json   Pljson;
  v_Error  Varchar2(2000);
  n_Count  Number(3);
  v_No     病人挂号记录.No%Type;
  v_Tmp    Varchar2(100);
  n_病人id Number;
  n_就诊id Number;
  n_场合   Number;
Begin
  j_In     := Pljson(Json_In);
  j_Json   := j_In.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');
  n_就诊id := j_Json.Get_Number('visit_id');
  n_场合   := j_Json.Get_Number('occasion');
  --门诊取挂号单
  If Nvl(n_就诊id, 0) = 0 Then
    Return;
  End If;
  If Nvl(n_场合, 0) = 1 Then
    Select NO Into v_No From 病人挂号记录 Where ID = n_就诊id;
    If v_No Is Null Then
      v_Error  := '未找到该病人的挂号记录,不能更新病人基本信息.';
      Json_Out := Zljsonout(v_Error, 1);
      Return;
    End If;
    --门诊医嘱签名,则不允许修改病人基本信息
    Select Nvl(Count(1), 0)
    Into n_Count
    From 病人医嘱记录
    Where 病人id = n_病人id And 挂号单 = v_No And 新开签名id Is Not Null And Rownum < 2;
    If n_Count <> 0 Then
      v_Error  := '病人医嘱已经签名,不能更新病人基本信息.';
      Json_Out := Zljsonout(v_Error, 1);
      Return;
    End If;
  End If;
  --住院病人
  If Nvl(n_场合, 0) = 2 Then
    --住院医嘱签名,则不允许修改病人基本信息
    Select Nvl(Count(1), 0)
    Into n_Count
    From 病人医嘱记录
    Where 病人id = n_病人id And 主页id = n_就诊id And 新开签名id Is Not Null And Rownum < 2;
  
    If n_Count <> 0 Then
      v_Error  := '该病人医嘱已经签名,不能更新病人基本信息.';
      Json_Out := Zljsonout(v_Error, 1);
      Return;
    End If;
    --病案处于锁定状态，则不允许修改病人基本信息
    Select Decode(病案状态, 1, '等待审查中', 3, '正在审查中', 5, '已经审查归档', 10, '接收待审中', Null)
    Into v_Tmp
    From 病案主页
    Where 病人id = n_病人id And 主页id = n_就诊id;
  
    If Not v_Tmp Is Null Then
      v_Error  := '该病人的病案' || v_Tmp || ',不能更新病人基本信息.';
      Json_Out := Zljsonout(v_Error, 1);
      Return;
    End If;
    --病案处于编目状态，则不允许修改病人基本信息
    Select Nvl(Count(1), 0)
    Into n_Count
    From 病案主页
    Where 病人id = n_病人id And 主页id = n_就诊id And 编目日期 Is Not Null;
    If n_Count <> 0 Then
      v_Error  := '该病人的病案已经编目,不能更新病人基本信息.';
      Json_Out := Zljsonout(v_Error, 1);
      Return;
    End If;
  End If;
  --以要素方式存储的姓名、性别、年龄
  For r_Rec In (Select /*+ RULE */
                Distinct d.名称 科室, a.Id, a.病历名称, a.病历种类, a.完成时间, a.保存人,
                         Nvl((Select Decode(a.编辑方式, 0, Decode(Substr(b.对象属性, 1, 1), '2', 1, 0),
                                             Decode(Substr(b.对象属性, Instr(b.对象属性, '|'), 1), '2', 1, 0))
                              From 电子病历内容 B
                              Where a.Id = b.文件id And
                                    ((b.对象类型 = 8 And a.编辑方式 = 0) Or (b.对象类型 In (6, 7, 8) And a.编辑方式 = 1)) And Rownum < 2),
                              0) As 电子签名
                From 电子病历记录 A, 部门表 D
                Where a.病人id = n_病人id And a.主页id = n_就诊id And a.科室id = d.Id And Exists
                 (Select 1 --包含
                       From 电子病历内容 C
                       Where a.Id = c.文件id And ((c.对象类型 = 4 And c.要素名称 = '姓名') Or (c.对象类型 = 4 And c.要素名称 = '性别') Or
                             (c.对象类型 = 4 And c.要素名称 = '年龄') Or (c.对象类型 = 2 And c.要素名称 = '姓名') Or
                             (c.对象类型 = 2 And c.要素名称 = '性别') Or (c.对象类型 = 2 And c.要素名称 = '年龄'))) And Rownum < 2) Loop
  
    --读取所有包含姓名、性别、年龄要素的病历
    If r_Rec.电子签名 = 1 Then
      --构建电子签名病历报错串
      v_Error  := '书写的病历已经进行过电子签名,不能进行病人信息修改操作！';
      Json_Out := Zljsonout(v_Error, 1);
      Return;
    End If;
  End Loop;

  --以自由录入的姓名
  If Nvl(n_场合, 0) = 0 Then
    For r_Rec In (Select /*+ RULE */
                  Distinct d.名称 科室, a.Id, a.病历名称, a.病历种类, a.完成时间, a.保存人,
                           Nvl((Select Decode(a.编辑方式, 0, Decode(Substr(b.对象属性, 1, 1), '2', 1, 0),
                                               Decode(Substr(b.对象属性, Instr(b.对象属性, '|'), 1), '2', 1, 0))
                                From 电子病历内容 B
                                Where a.Id = b.文件id And ((b.对象类型 = 8 And 编辑方式 = 0) Or (b.对象类型 In (6, 7, 8) And 编辑方式 = 1)) And
                                      Rownum < 2), 0) As 电子签名
                  From 电子病历记录 A, 部门表 D, 病人挂号记录 E
                  Where a.病人id = n_病人id And a.主页id = n_就诊id And a.科室id = d.Id And a.病人id = e.病人id And Exists
                   (Select 1 --包含
                         From 电子病历内容 C
                         Where a.Id = c.文件id And
                               ((c.对象类型 = 2 And Instr(c.内容文本, e.姓名) > 0) Or (c.对象类型 = 1 And Instr(c.内容文本, e.姓名) > 0))) And
                        Rownum < 2) Loop
    
      --读取所有包含姓名的病历
      If r_Rec.电子签名 = 1 Then
        --构建电子签名病历报错串
        v_Error  := '书写的病历已经进行过电子签名,不能进行病人信息修改操作！';
        Json_Out := Zljsonout(v_Error, 1);
        Return;
      End If;
    End Loop;
  Else
    For r_Rec In (Select /*+ RULE */
                  Distinct d.名称 科室, a.Id, a.病历名称, a.病历种类, a.完成时间, a.保存人,
                           Nvl((Select Decode(a.编辑方式, 0, Decode(Substr(b.对象属性, 1, 1), '2', 1, 0),
                                               Decode(Substr(b.对象属性, Instr(b.对象属性, '|'), 1), '2', 1, 0))
                                From 电子病历内容 B
                                Where a.Id = b.文件id And ((b.对象类型 = 8 And 编辑方式 = 0) Or (b.对象类型 In (6, 7, 8) And 编辑方式 = 1)) And
                                      Rownum < 2), 0) As 电子签名
                  From 电子病历记录 A, 部门表 D, 病案主页 E
                  Where a.病人id = n_病人id And a.主页id = n_就诊id And a.科室id = d.Id And a.病人id = e.病人id And Exists
                   (Select 1 --包含
                         From 电子病历内容 C
                         Where a.Id = c.文件id And
                               ((c.对象类型 = 2 And Instr(c.内容文本, e.姓名) > 0) Or (c.对象类型 = 1 And Instr(c.内容文本, e.姓名) > 0))) And
                        Rownum < 2) Loop
    
      --读取所有包含姓名的病历
      If r_Rec.电子签名 = 1 Then
        --构建电子签名病历报错串
        v_Error  := '书写的病历已经进行过电子签名,不能进行病人信息修改操作！';
        Json_Out := Zljsonout(v_Error, 1);
        Return;
      End If;
    End Loop;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Updpatbaseinfocheck;
/

Create Or Replace Procedure Zl_Cissvr_Getfeeitem
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：获取诊疗项目对应的收费细目明细
  --入参：Json_In:格式
  --  input
  --      nodeno                                     C 0 站点，可不传
  --      plcdept_id                                 N 1 开单科室id
  --      item_list[]                                项目列表，即新门诊的各类申请明细对应的诊疗项目信息和执行科室信息
  --          apply_id                               C 1 申请序号，新门诊的申请号，唯一标识一个申请
  --          apply_type                             N 1 申请类别：1-西药或成药，2-中药，3-检验，4-治疗(麻醉，皮试，针炙，换药等)，5-检查
  --          cure_info                              一般治疗项目信息，可以不传
  --              cure_item_id                       N 1 诊疗项目id,普通治疗项目对应的诊疗项目id
  --              cure_exedept_id                    N 1 执行科室id,普通治疗项目对应的诊疗项目id
  --          lis_info                               检验项目信息
  --              lis_items                          C 1 检验项目的诊疗项目id，可以传多个逗号拼串，当为多个时表示一并采集
  --              lis_exedept_id                     N 1 检验项目对应的执行科室id
  --              lis_collect_item_id                N 1 检验采集项目id
  --              lis_collect_exedept_id             N 1 检验采集项目对应的采集执行科室id
  --              lis_spcm                           C 1 检验采集标本
  --              emergency_tag                      N 1 紧急标识表明当前申请是否是紧急执行
  --          pacs_info                              检查项目信息
  --              pacs_item_id                       N 1 检查项目id
  --              pacs_exedept_id                    N 1 检查项目执行科室id
  --              pacs_part_list[]                   检查项目的部位方法列表，可不传，不传则只有一行数据
  --                  part_name                      C 1 检查部位名称【对应ZLHIS标本部件】
  --                  part_way                       C 1 检查方法名称【对应ZLHIS检查方法】
  --          drug_info                              药品或中药项目信息
  --              drug_use_item_id                   N 1 药品给药途径项目id,对于配方代表服法项目id
  --              drug_use_exedept_id                N 1 药品给药途径项目对应的执行科室id
  --              drug_decoction_id                  N 1 煎法项目id，中药配方时才传入
  --              drug_decoction_exedept_id          N 1 煎法项目对应的执行科室id可以不传，不传时认为是无需执行的叮嘱

  --出参: Json_Out,格式如下
  --  output
  --    code                      N 1 应答吗：0-失败；1-成功
  --    message                   C 1 应答消息：失败时返回具体的错误信息
  --    item_list[]     医嘱列表，即新门诊的各类申请明细
  --          apply_id            C 1 申请序号，新门诊的申请号，唯一标识一个新门诊申请
  --          cisitem_id          N 1 诊疗项目id
  --          fee_item_id         N 1 收费细目id
  --          exedept_id          N 0 执行部门id，当诊疗项目对应的是药品或卫材费用时需要外部获取，不传时以医嘱诊疗项目的执行科室id为准，当为药品和卫材费用时这个结点可能为空时不能指定，此时由调用方进行二次计算
  --          quantity            N 1 收费数量，诊疗收费关系中的对照数量
  --          part_name           C 0 检查部位名称，ZLHIS中费用对照存在按部位和项目对照
  --          part_way            C 0 检查方法名称
  --    reject_list[] 检项目排斥情况，主要是针对采集项目上的绑定费用，因为合管的原因，其它申请采集费用和试管费用不会收取，此列表记录合管情况
  --          apply_id            C 1 申请序号，正常收取的申请序号
  --          reject_id           C 1 被合并的申请序号，新门诊的申请号，唯一标识一个新门诊申请    
  ---------------------------------------------------------------------------

  v_Nodeno       Varchar2(2000);
  v_付款方式名称 Varchar2(2000);
  v_价格等级     Varchar2(100);
  v_普通等级     Varchar2(100);
  v_药品等级     Varchar2(100);
  v_卫材等级     Varchar2(100);
  v_Pricegrade   Varchar2(500);

  --LIS采集项目收费对照
  Type Rs_Collection Is Record(
    申请标识       Varchar2(4000),
    申请排斥       Varchar2(4000),
    项目id         Number(18),
    执行科室id     Number(18),
    管码           Varchar2(200),
    管码卫材id     Number(18),
    采集项目id     Number(18),
    采集执行科室id Number(18),
    采集标本       Varchar2(4000),
    收费细目id     Number(18),
    收费数量       Number(16, 5),
    单价           Number(16, 5),
    部位           Varchar2(4000),
    方法           Varchar2(4000),
    紧急标志       Number(1),
    收费类别       Varchar2(10),
    采集方法       Varchar2(300),
    收费项目名称   Varchar2(300),
    收费单位       Varchar2(300),
    收费方式       Number(1));
  Type t_Col Is Table Of Rs_Collection;
  r_Lis      t_Col; --采集项目收费对照缓存列表
  r_Item     t_Col; --诊疗收费对照 
  r_采集排斥 t_Col;

  --普通项目
  Cursor c_普通
  (
    P项目id Number,
    P科室id Number
  ) Is
    Select a.诊疗项目id, a.收费项目id, a.收费数量, a.收费方式, a.执行科室id, b.材料id, c.类别 收费类别, c.名称 收费项目名称, c.计算单位 收费单位
    From (Select c.诊疗项目id, c.收费项目id, c.收费数量, c.收费方式, c.执行科室id
           From (Select c.诊疗项目id, c.收费项目id, c.费用性质, c.收费数量, c.固有对照, c.从属项目, c.收费方式, c.适用科室id, P科室id 执行科室id,
                         Max(Nvl(c.适用科室id, 0)) Over(Partition By c.诊疗项目id, c.检查部位, c.检查方法, c.费用性质) As Top
                  From 诊疗收费关系 C
                  Where c.诊疗项目id = P项目id And c.检查部位 Is Null And c.检查方法 Is Null And
                        (c.适用科室id Is Null And Nvl(c.病人来源, 0) = 0 Or c.适用科室id = P科室id And c.病人来源 = 1)) C
           Where Nvl(c.适用科室id, 0) = c.Top) A, 收费项目目录 C, 材料特性 B
    Where a.收费项目id = c.Id And c.服务对象 In (1, 3) And (c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.撤档时间 Is Null) And
          (c.站点 = v_Nodeno Or c.站点 Is Null) And c.Id = b.材料id(+);

  --检查部位对照
  Cursor c_检查
  (
    P项目id Number,
    P科室id Number,
    P部位   Varchar2,
    P方法   Varchar2
  ) Is
    Select a.诊疗项目id, a.收费项目id, a.收费数量, a.收费方式, a.执行科室id, b.材料id, c.类别 收费类别, c.名称 收费项目名称, c.计算单位 收费单位
    From (Select c.诊疗项目id, c.收费项目id, c.收费数量, c.收费方式, c.执行科室id
           From (Select c.诊疗项目id, c.收费项目id, c.费用性质, c.收费数量, c.固有对照, c.从属项目, c.收费方式, c.适用科室id, P科室id 执行科室id,
                         Max(Nvl(c.适用科室id, 0)) Over(Partition By c.诊疗项目id, c.检查部位, c.检查方法, c.费用性质) As Top
                  From 诊疗收费关系 C
                  Where c.诊疗项目id = P项目id And c.检查部位 = P部位 And c.检查方法 = P方法 And
                        (c.适用科室id Is Null And Nvl(c.病人来源, 0) = 0 Or c.适用科室id = P科室id And c.病人来源 = 1)) C
           Where Nvl(c.适用科室id, 0) = c.Top) A, 收费项目目录 C, 材料特性 B
    Where a.收费项目id = c.Id And c.服务对象 In (1, 3) And (c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.撤档时间 Is Null) And
          (c.站点 = v_Nodeno Or c.站点 Is Null) And c.Id = b.材料id(+);

  j_Json       Pljson;
  j_In         Pljson;
  j_Json_Tmp   Pljson;
  j_Jsonlist   Pljson_List;
  n_开单科室id Number(18);
  n_Count      Number(6);
  n_申请类别   Number(2); --申请类别：1-西药或成药给药途径，2-中药，3-检验，4-治疗(麻醉，皮试，针炙，换药等)，5-检查
  v_申请标识   Varchar2(4000);
  v_Out_Tmp    Varchar2(32767);

  Function Getitem_Price
  (
    P项目id   Number,
    P收费类别 Varchar2
  ) Return Number As
    --获取收费项目的单价
    n_单价 Number(16, 5);
  Begin
    If Instr(',5,6,7,', ',' || P收费类别 || ',') > 0 Then
      v_价格等级 := v_药品等级;
    Elsif P收费类别 = '4' Then
      v_价格等级 := v_卫材等级;
    Else
      v_价格等级 := v_普通等级;
    End If;
    n_单价 := 0;
    For r_收入项目 In (Select a.Id As 收费细目id, b.收入项目id, c.名称, c.收据费目, b.现价, b.原价, b.加班加价率, b.附术收费率, b.缺省价格, a.计算单位, a.费用类型,
                          a.屏蔽费别, a.类别 As 收费类别
                   From 收费项目目录 A, 收费价目 B, 收入项目 C
                   Where b.收费细目id = a.Id And c.Id = b.收入项目id And Sysdate Between b.执行日期 And
                         Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And a.Id = P项目id And
                         ((b.价格等级 Is Null And Nvl(v_价格等级, '-') = '-') Or
                         (b.价格等级 = v_价格等级 Or
                         (b.价格等级 Is Null And Not Exists
                          (Select 1
                             From 收费价目
                             Where b.收费细目id = 收费细目id And 价格等级 = v_价格等级 And Sysdate Between 执行日期 And
                                   Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))))))
                   Order By 收费细目id, 收入项目id) Loop
      n_单价 := n_单价 + Nvl(r_收入项目.现价, 0);
    End Loop;
    Return n_单价;
  End;

  Function Get药品卫材缺省执行科室id
  (
    P项目id   Number,
    P收费类别 Varchar2 --4,5,6,7
  ) Return Number As
    --功能：获取药品卫材一个缺省的执行科室id
    v_工作性质   Varchar2(100);
    n_执行科室id Number(18);
  Begin
    If P收费类别 = '4' Then
      v_工作性质 := '发料部门';
    Elsif P收费类别 = '5' Then
      v_工作性质 := '西药房';
    Elsif P收费类别 = '6' Then
      v_工作性质 := '成药房';
    Elsif P收费类别 = '7' Then
      v_工作性质 := '中药房';
    End If;
  
    For R In (Select a.执行科室id
              From 收费执行科室 A, 部门性质说明 B, 部门表 C
              Where a.执行科室id + 0 = b.部门id And b.工作性质 = v_工作性质 And b.服务对象 In (1, 3) And b.部门id = c.Id And
                    (c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.撤档时间 Is Null) And (a.病人来源 Is Null Or a.病人来源 = 1) And
                    Nvl(a.开单科室id, 0) = n_开单科室id And (c.站点 = v_Nodeno Or c.站点 Is Null) And a.收费细目id = P项目id
              Order By b.服务对象, c.编码) Loop
      n_执行科室id := r.执行科室id;
      Exit;
    End Loop;
    Return n_执行科室id;
  End;

  Procedure Additem
  (
    P申请标识 Varchar2,
    P项目id   Number,
    P科室id   Number,
    P部位     Varchar2 := Null,
    P方法     Varchar2 := Null
  ) As
    --说明：一般诊疗收费对照
    N Number(3);
  Begin
    If P部位 Is Null Then
      For R In c_普通(P项目id, P科室id) Loop
        r_Item.Extend;
        N := r_Item.Count;
        r_Item(N).申请标识 := P申请标识;
        r_Item(N).项目id := r.诊疗项目id;
        r_Item(N).执行科室id := r.执行科室id;
        r_Item(N).管码 := Null;
        r_Item(N).管码卫材id := Null;
        r_Item(N).采集项目id := Null;
        r_Item(N).采集执行科室id := Null;
        r_Item(N).收费细目id := r.收费项目id;
        r_Item(N).收费数量 := r.收费数量;
        r_Item(N).部位 := Null;
        r_Item(N).方法 := Null;
        r_Item(N).收费方式 := r.收费方式;
      
        If r.材料id Is Not Null Then
          r_Item(N).执行科室id := Get药品卫材缺省执行科室id(r.收费项目id, '4');
        Elsif r.收费类别 In ('5', '6', '7') Then
          r_Item(N).执行科室id := Get药品卫材缺省执行科室id(r.收费项目id, r.收费类别);
        End If;
        r_Item(N).收费项目名称 := r.收费项目名称;
        r_Item(N).收费单位 := r.收费单位;
        r_Item(N).单价 := Getitem_Price(r.收费项目id, r.收费类别);
        r_Item(N).收费类别 := r.收费类别;
      End Loop;
    Else
      For R In c_检查(P项目id, P科室id, P部位, P方法) Loop
        r_Item.Extend;
        N := r_Item.Count;
        r_Item(N).申请标识 := P申请标识;
        r_Item(N).项目id := r.诊疗项目id;
        r_Item(N).执行科室id := r.执行科室id;
        r_Item(N).管码 := Null;
        r_Item(N).管码卫材id := Null;
        r_Item(N).采集项目id := Null;
        r_Item(N).采集执行科室id := Null;
        r_Item(N).收费细目id := r.收费项目id;
        r_Item(N).收费数量 := r.收费数量;
        r_Item(N).部位 := P部位;
        r_Item(N).方法 := P方法;
        r_Item(N).收费方式 := r.收费方式;
      
        If r.材料id Is Not Null Then
          r_Item(N).执行科室id := Get药品卫材缺省执行科室id(r.收费项目id, '4');
        Elsif r.收费类别 In ('5', '6', '7') Then
          r_Item(N).执行科室id := Get药品卫材缺省执行科室id(r.收费项目id, r.收费类别);
        End If;
        r_Item(N).收费项目名称 := r.收费项目名称;
        r_Item(N).收费单位 := r.收费单位;
        r_Item(N).单价 := Getitem_Price(r.收费项目id, r.收费类别);
        r_Item(N).收费类别 := r.收费类别;
      End Loop;
    End If;
  End;

  Procedure Additem_采集
  (
    P申请标识   Varchar2,
    P项目id     Number,
    P科室id     Number,
    P检验项目id Number,
    P检验科室id Number,
    P紧急       Number,
    P采集标本   Varchar2
  ) As
    --采集项目收费对照
    --说明：特殊收费方式 检验试管费用，要求基础设置中 检查项目设置管码，采血管绑定卫材 才生效，如果对照关系未设对都会当成正常收取可能引起收费不正确
    N                  Number(3);
    v_管码             Varchar2(3000);
    n_材料id           Number(18);
    n_检验试管费用已收 Number(1) := 0; --绑定费用中 收费方式为 1-检验试管费用 性质的费用已收取，但剩余的正常收取方式的费用需要继续添加
    n_要收费           Number(1);
    v_采集方法         Varchar2(300);
  Begin
    Select Max(a.名称) Into v_采集方法 From 诊疗项目目录 A Where a.Id = P项目id;
    Select Max(a.试管编码) Into v_管码 From 诊疗项目目录 A Where a.Id = P检验项目id;
    --只取一个
    For R In (Select 编码, 材料id From 采血管类型 Where 材料id Is Not Null And 编码 = v_管码) Loop
      n_材料id := r.材料id;
    End Loop;
  
    --先从已有费用查找，判断是否是已经采集过一次了，判断是否需要这条收费
    For N In 1 .. r_Lis.Count Loop
      If P检验项目id <> r_Lis(N).项目id And r_Lis(N).采集项目id = P项目id And r_Lis(N).管码 = v_管码 And r_Lis(N).执行科室id = P检验科室id And r_Lis(N)
        .紧急标志 = Nvl(P紧急, 0) And r_Lis(N).采集标本 = P采集标本 And r_Lis(N).采集执行科室id = P科室id Then
        n_检验试管费用已收 := 1; --已找到 已收过的 检验试管理费用 性质的费用
      
        r_采集排斥.Extend;
        r_采集排斥(r_采集排斥.Count).申请标识 := r_Lis(N).申请标识;
        r_采集排斥(r_采集排斥.Count).申请排斥 := P申请标识;
      
        Exit;
      End If;
    End Loop;
  
    For R In c_普通(P项目id, P科室id) Loop
      n_要收费 := 1;
    
      If n_检验试管费用已收 = 1 And r.收费方式 = 1 Then
        --特殊收费方式 检验试管费用 的费用明细排除掉
        n_要收费 := 0;
      End If;
    
      If n_要收费 = 1 Then
        --材料费排它性
        If r.收费方式 = 1 And n_材料id <> r.材料id Then
          n_要收费 := 0;
        End If;
      End If;
    
      If n_要收费 = 1 Then
      
        r_Lis.Extend;
        N := r_Lis.Count;
        r_Lis(N).申请标识 := P申请标识;
        r_Lis(N).项目id := P检验项目id;
        r_Lis(N).执行科室id := P检验科室id;
        r_Lis(N).管码 := v_管码;
        r_Lis(N).管码卫材id := n_材料id;
        r_Lis(N).采集项目id := r.诊疗项目id;
        r_Lis(N).采集执行科室id := r.执行科室id;
        r_Lis(N).收费细目id := r.收费项目id;
        r_Lis(N).收费数量 := r.收费数量;
        r_Lis(N).紧急标志 := Nvl(P紧急, 0);
        r_Lis(N).采集标本 := P采集标本;
        r_Lis(N).部位 := Null;
        r_Lis(N).方法 := Null;
        r_Lis(N).收费方式 := r.收费方式;
      
        --追到通用收费明细项目中
        r_Item.Extend;
        N := r_Item.Count;
        r_Item(N).申请标识 := P申请标识;
        r_Item(N).项目id := r.诊疗项目id;
        r_Item(N).执行科室id := r.执行科室id;
        r_Item(N).收费细目id := r.收费项目id;
        r_Item(N).收费数量 := r.收费数量;
      
        If r.材料id Is Not Null Then
          r_Item(N).执行科室id := Get药品卫材缺省执行科室id(r.收费项目id, '4');
        Elsif r.收费类别 In ('5', '6', '7') Then
          r_Item(N).执行科室id := Get药品卫材缺省执行科室id(r.收费项目id, r.收费类别);
        End If;
        r_Item(N).采集方法 := v_采集方法;
        r_Item(N).采集标本 := P采集标本;
        r_Item(N).收费项目名称 := r.收费项目名称;
        r_Item(N).收费单位 := r.收费单位;
        r_Item(N).单价 := Getitem_Price(r.收费项目id, r.收费类别);
        r_Item(N).收费类别 := r.收费类别;
        r_Item(N).采集项目id := P项目id;
      End If;
    End Loop;
  End;

  Procedure Getitem(j_Par_In Pljson) As
    --按json项目结点进行解析
    n_项目id      Number(18);
    n_科室id      Number(18);
    v_部位        Varchar2(4000);
    v_方法        Varchar2(4000);
    v_检验项目ids Varchar2(4000);
    j_List        Pljson_List;
    j_Json_Tmp    Pljson;
    j_Par         Pljson;
  Begin
    v_申请标识 := j_Par_In.Get_String('apply_id');
    n_申请类别 := j_Par_In.Get_Number('apply_type');
    If n_申请类别 = 1 Then
      j_Par    := j_Par_In.Get_Pljson('drug_info');
      n_项目id := j_Par.Get_Number('drug_use_item_id');
      n_科室id := j_Par.Get_Number('drug_use_exedept_id');
      Additem(v_申请标识, n_项目id, n_科室id, Null, Null);
    Elsif n_申请类别 = 2 Then
      j_Par    := j_Par_In.Get_Pljson('drug_info');
      n_项目id := j_Par.Get_Number('drug_use_item_id');
      n_科室id := j_Par.Get_Number('drug_use_exedept_id');
      Additem(v_申请标识, n_项目id, n_科室id, Null, Null);
      n_项目id := j_Par.Get_Number('drug_decoction_id');
      n_科室id := j_Par.Get_Number('drug_decoction_exedept_id');
      Additem(v_申请标识, n_项目id, n_科室id, Null, Null);
    Elsif n_申请类别 = 3 Then
      j_Par         := j_Par_In.Get_Pljson('lis_info');
      v_检验项目ids := j_Par.Get_String('lis_items');
      n_科室id      := j_Par.Get_Number('lis_exedept_id');
      For R In (Select /*+cardinality(j,10) */
                 j.Column_Value 检验项目id
                From Table(Cast(f_Num2list(v_检验项目ids) As Zltools.t_Numlist)) J) Loop
        If n_项目id Is Null Then
          --一并采集的首行检验项目
          n_项目id := r.检验项目id;
        End If;
        Additem(v_申请标识, r.检验项目id, n_科室id, Null, Null);
      End Loop;
      Additem_采集(v_申请标识, j_Par.Get_Number('lis_collect_item_id'), j_Par.Get_Number('lis_collect_exedept_id'), n_项目id,
                 n_科室id, j_Par.Get_Number('emergency_tag'), j_Par.Get_String('lis_spcm'));
    Elsif n_申请类别 = 4 Then
      j_Par    := j_Par_In.Get_Pljson('cure_info');
      j_Par    := j_Par_In.Get_Pljson('cure_info');
      n_项目id := j_Par.Get_Number('cure_item_id');
      n_科室id := j_Par.Get_Number('cure_exedept_id');
      Additem(v_申请标识, n_项目id, n_科室id, Null, Null);
    Elsif n_申请类别 = 5 Then
      j_Par    := j_Par_In.Get_Pljson('pacs_info');
      n_项目id := j_Par.Get_Number('pacs_item_id');
      n_科室id := j_Par.Get_Number('pacs_exedept_id');
      Additem(v_申请标识, n_项目id, n_科室id, Null, Null);
      j_List := j_Par.Get_Pljson_List('pacs_part_list');
      If j_List Is Not Null Then
        For I In 1 .. j_List.Count Loop
          j_Json_Tmp := Pljson();
          j_Json_Tmp := Pljson(j_List.Get(I));
          v_部位     := j_Json_Tmp.Get_String('part_name');
          v_方法     := j_Json_Tmp.Get_String('part_way');
          Additem(v_申请标识, n_项目id, n_科室id, v_部位, v_方法);
        End Loop;
      End If;
    End If;
  End;

  Function Get采集排斥 Return Varchar2 As
    --功能：获以被合并的费用项目，主要针对采集项目对照的费用合并情况
    --出参格：",reject_list":[{"apply_id":"444","reject_id":"2323"},{},{}...]  
    --          reject_list[] 检项目排斥情况，主要是针对采集项目上的绑定费用
    --               apply_id       C 1 申请序号，新门诊的申请号，唯一标识一个新门诊申请
    --               reject_id      C 1 被合并的申请序号，新门诊的申请号，唯一标识一个新门诊申请  
  
    v_Jtmp Varchar2(32767);
  Begin
    For I In 1 .. r_采集排斥.Count Loop
      v_Jtmp := v_Jtmp || ',{"apply_id":"' || r_采集排斥(I).申请标识 || '"';
      v_Jtmp := v_Jtmp || ',"reject_id":"' || r_采集排斥(I).申请排斥 || '"';
      v_Jtmp := v_Jtmp || '}';
    End Loop;
    If v_Jtmp Is Not Null Then
      v_Jtmp := ',"reject_list":[' || Substr(v_Jtmp, 2) || ']';
    End If;
    Return v_Jtmp;
  End;

Begin
  --解析入参
  j_In         := Pljson(Json_In);
  j_Json       := j_In.Get_Pljson('input');
  v_Nodeno     := j_Json.Get_String('nodeno');
  n_开单科室id := j_Json.Get_Number('plcdept_id');
  j_Jsonlist   := j_Json.Get_Pljson_List('item_list');

  v_Nodeno       := j_Json.Get_String('site_no');
  v_付款方式名称 := j_Json.Get_String('mdlpay_mode_name');
  If Nvl(v_Nodeno, '-') = '-' And Nvl(v_付款方式名称, '-') = '-' Then
    v_普通等级 := Null;
    v_药品等级 := Null;
    v_卫材等级 := Null;
  Else
    v_Pricegrade := Zl_Get_Pricegrade_s(v_Nodeno, v_付款方式名称);
    For c_价格等级 In (Select Rownum As 序号, Column_Value As 价格等级 From Table(f_Str2list(v_Pricegrade, '|'))) Loop
      If c_价格等级.序号 = 1 Then
        v_普通等级 := c_价格等级.价格等级;
      End If;
      If c_价格等级.序号 = 2 Then
        v_药品等级 := c_价格等级.价格等级;
      End If;
      If c_价格等级.序号 = 3 Then
        v_卫材等级 := c_价格等级.价格等级;
      End If;
    End Loop;
  End If;

  If Not j_Jsonlist Is Null Then
    n_Count    := j_Jsonlist.Count;
    r_Item     := t_Col();
    r_Lis      := t_Col();
    r_采集排斥 := t_Col();
    For I In 1 .. n_Count Loop
      j_Json_Tmp := Pljson();
      j_Json_Tmp := Pljson(j_Jsonlist.Get(I));
      Getitem(j_Json_Tmp);
    End Loop;
  End If;

  v_Out_Tmp := Null;
  For I In 1 .. r_Item.Count Loop
    v_Out_Tmp := v_Out_Tmp || ',{"apply_id":"' || r_Item(I).申请标识 || '"';
    v_Out_Tmp := v_Out_Tmp || ',"cisitem_id":' || Nvl(r_Item(I).项目id, 0);
    v_Out_Tmp := v_Out_Tmp || ',"fee_item_id":' || Nvl(r_Item(I).收费细目id, 0);
    v_Out_Tmp := v_Out_Tmp || ',"fee_item_type":"' || r_Item(I).收费类别 || '"';
    v_Out_Tmp := v_Out_Tmp || ',"fee_item_name":"' || r_Item(I).收费项目名称 || '"';
    v_Out_Tmp := v_Out_Tmp || ',"item_unit":"' || r_Item(I).收费单位 || '"';
    v_Out_Tmp := v_Out_Tmp || ',"price":' || Zljsonstr(r_Item(I).单价, 1);
    v_Out_Tmp := v_Out_Tmp || ',"exedept_id":' || Nvl(r_Item(I).执行科室id, 0);
    v_Out_Tmp := v_Out_Tmp || ',"quantity":' || Nvl(r_Item(I).收费数量, 0);
    v_Out_Tmp := v_Out_Tmp || ',"part_name":"' || r_Item(I).部位 || '"';
    v_Out_Tmp := v_Out_Tmp || ',"part_way":"' || r_Item(I).方法 || '"';
    v_Out_Tmp := v_Out_Tmp || ',"lis_spcm":"' || r_Item(I).采集标本 || '"';
    v_Out_Tmp := v_Out_Tmp || ',"lis_collect_way":"' || r_Item(I).采集方法 || '"';
    v_Out_Tmp := v_Out_Tmp || ',"lis_collect_item_id":' || Nvl(r_Item(I).采集项目id, 0);
    v_Out_Tmp := v_Out_Tmp || '}';
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || Substr(v_Out_Tmp, 2) || ']' || Get采集排斥 || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Cissvr_Getfeeitem;
/

Create Or Replace Procedure Zl_Cissvr_Outsendapply
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  ---------------------------------------------------------------------------
  --功能：接收新门诊药品(配方)/检验/一般治疗项目申请，生成【医嘱/计价/发送/执行】相关的临床域数据
  --入参：Json_In:格式
  --  input  

  --出参: Json_Out,格式如下
  --  output
  --    code                N 1 应答吗：0-失败；1-成功
  --    message             C 1 应答消息：失败时返回具体的错误信息
  --    advice_list[]       医嘱列表，即新门诊的各类申请明细
  --      apply_id          C 1 申请序号，新门诊的申请号，唯一标识一个新门诊申请
  --      fee_no            C 1 费用单据号
  --      lis_bar_code      C 1 检验项目条码，仅当检验项目才有此结点
  --      order_list[]      产生的医嘱列表
  --      order_id          N 1 医嘱id,zlhis中生成的医嘱id
  --      order_related_id  N 1 相关id,zlhis中生成的相关id
  --      cisitem_id        N 1 诊疗项目id
  ---------------------------------------------------------------------------
  --诊断列表 
  Type r_Diag Is Record(
    Diag_Num      Varchar2(200), --诊断序号     
    Csd_Code      Varchar2(200), --诊断编码，中医诊断编码 疾病编码目录
    Icd10_Code    Varchar2(200), --疾病编码 ICD10 西医
    Dz_Note       Varchar2(4000), --诊断描述
    Syndrome      Varchar2(4000), --中医证候
    Disease_Time  Varchar2(30), --发病时间 日期格式
    Diagnostician Varchar2(200), --诊断医生，记录人
    是否存在      Number(1), --     可能是已有的诊断记录,
    诊断类型      Number(1), --诊断类型，1-西医诊断，11-中医诊断       
    诊断描述      Varchar2(4000),
    疾病id        Number(18),
    诊断次序      Number(2),
    诊断ids       Varchar2(4000),
    医嘱id        Number(18),
    诊断id        Number(18));

  Type r_Apply Is Record(
    Apply_Id     Varchar2(200), --申请唯一标识，按此标志产生发送的NO号
    Apply_Type   Number(1), --申请类别：1-西药或成药，2-中药，3-检验，4-治疗(麻醉，皮试，针炙，换药等)，5-检查
    Diag_Nums    Varchar2(4000), --诊断序号
    Zlhis记录ids Varchar2(4000),
    Serial_Num   Number(18),
    Group_Sno    Number(18),
    ID           Number(18),
    相关id       Number(18),
    前提id       Number(18),
    病人来源     Number(1),
    病人id       Number(18),
    主页id       Number(5),
    挂号单       Varchar2(8),
    婴儿         Number(3),
    姓名         Varchar2(100),
    性别         Varchar2(4),
    年龄         Varchar2(20),
    病人科室id   Number(18),
    序号         Number(18),
    医嘱状态     Number(3),
    医嘱期效     Number(1),
    诊疗类别     Varchar2(1),
    诊疗项目id   Number(18),
    标本部位     Varchar2(60),
    检查方法     Varchar2(30),
    收费细目id   Number(18),
    天数         Number(16, 5),
    单次用量     Number(16, 5),
    首次用量     Number(16, 5),
    总给予量     Number(16, 5),
    医嘱内容     Varchar2(1000),
    医生嘱托     Varchar2(200),
    执行科室id   Number(18),
    皮试结果     Varchar2(10),
    执行频次     Varchar2(20),
    频率次数     Number(3),
    频率间隔     Number(3),
    间隔单位     Varchar2(4),
    执行时间方案 Varchar2(100),
    计价特性     Number(1),
    执行性质     Number(1),
    执行标记     Number(1),
    审核标记     Number(1),
    可否分零     Number(3),
    紧急标志     Number(1),
    开始执行时间 Date,
    开嘱科室id   Number(18),
    开嘱医生     Varchar2(41),
    开嘱时间     Date,
    手术时间     Date,
    是否上传     Number(1),
    审查结果     Number(1),
    屏蔽打印     Number(1),
    摘要         Varchar2(1000),
    零费记帐     Number(1),
    用药目的     Number(1),
    用药理由     Varchar2(1000),
    超量说明     Varchar2(1000),
    管码         Varchar2(100),
    配方id       Number(18),
    ----发送用
    NO       Varchar2(60),
    记录序号 Number(3),
    发送号   Number(18),
    发送数次 Number(16, 5), --如果是药品，则是按计算单位计算的，需要换算
    首次时间 Date,
    末次时间 Date,
    发送时间 Date,
    样本条码 Varchar2(600),
    
    Firstrow Number(1) --ZLHIS中一组医嘱中的第一行
    );

  Type r_Price Is Record(
    
    Apply_Id   Varchar2(200), --申请唯一标识，按此标志产生发送的NO号
    检验项目id Number(18),
    检验科室id Number(18),
    采集项目id Number(18),
    采集科室id Number(18),
    采集标本   Varchar(200),
    紧急标志   Number(1),
    婴儿       Number(3),
    管码       Varchar2(100), --二次加工
    管码卫材id Number(18), --二次加工
    样本条码   Varchar(200), --二次加工    
    医嘱id     Number(18),
    诊疗项目id Number(18),
    收费细目id Number(18),
    执行科室id Number(18),
    收费方式   Number(1));

  Type t_Apply Is Table Of r_Apply;
  Type t_Diag Is Table Of r_Diag;
  Type t_Price Is Table Of r_Price;

  Rsdiag         t_Diag := t_Diag();
  Rs诊断医嘱     t_Diag := t_Diag();
  Rsdiagnew      t_Diag := t_Diag();
  Rsap           t_Apply := t_Apply();
  r_Base         r_Apply;
  Rs条码         t_Price := t_Price();
  n_就诊id       Number(18); --病人挂号记录挂号id
  v_Nodeno       Varchar2(1000);
  d_发送时间     Date;
  n_是否生成条码 Number; --检验项目是否生成条码标志，null/0不生成，1-要生成，仅病人来源为体检病人时才传入，其它条件下以ZLHIS系统参数为准，参数名：医嘱发送生成条形码

  --v_执行时间方案 Varchar2(200); --可不传；原则上要和advice_frequency配合使用，主要用于计算医嘱执行时间点，当不传时可根据 频率编码 取出缺省执行时间方案
  --                              按周执行  每周三次 1/8-3/8-5/8 或 1/8:00-3/8:00-5/8:00 表示在每周星期一的8:00,星期三的8:00,星期五的8:00这几个时间执行
  --                              按天执行  每天三次 8-12-16 或 8:00-12:00-16:00 表示在每天8:00,12:00,16:00这几个时间执行 
  --                                    两天一次 1/8 或 1/8:00 表示在每两天中的第1天8:00这个时间执行
  --                              按时执行 每小时两次 1:20-1:40 表示在每小时内的20和40分钟这两个时间执行 
  --                                    两小时一次 2:30 或 1:30 或 1:00 表示在每两小时内的第2的个小时的30分钟这个时间执行 或在每两小时内的第1的个小时的30分钟这个时间执行 或在每两小时内的第1的个小时这个时间执行

  v_Input      Pljson;
  j_Tmp        Pljson;
  Jl_Tmp       Pljson_List;
  j_Advicelist Pljson_List;
  Jl_Diag      Pljson_List;
  Idx          Number(6);
  j_Adviceitem Pljson;
  l_医嘱id     t_Numlist := t_Numlist();
  n_医嘱id     Number(18);
  n_序列id     Number(18) := 0; --全局变量， 
  n_序号       Number;
  v_产地       Varchar2(2000);
  v_规格       Varchar2(2000);
  n_剂量系数   Number(16, 5);
  n_发送号     病人医嘱发送.发送号%Type;
  v_No         病人医嘱发送.No%Type;

  n_Drug_Rows Number(6);
  n_Pre序号   Number(6);

  v_Lisitems Varchar2(4000);

  v_Out  Varchar2(32767);
  v_Tmp1 Varchar2(32767);

  --普通项目  P项目id - 诊疗项目目录 ，P科室id - 开单科室（申请科室id）
  Cursor c_普通
  (
    P项目id Number,
    P科室id Number
  ) Is
    Select a.诊疗项目id, a.收费项目id, a.收费数量, a.收费方式, a.执行科室id, b.材料id, c.类别 收费类别
    From (Select c.诊疗项目id, c.收费项目id, c.收费数量, c.收费方式, c.执行科室id
           From (Select c.诊疗项目id, c.收费项目id, c.费用性质, c.收费数量, c.固有对照, c.从属项目, c.收费方式, c.适用科室id, P科室id 执行科室id,
                         Max(Nvl(c.适用科室id, 0)) Over(Partition By c.诊疗项目id, c.检查部位, c.检查方法, c.费用性质) As Top
                  From 诊疗收费关系 C
                  Where c.诊疗项目id = P项目id And c.检查部位 Is Null And c.检查方法 Is Null And
                        (c.适用科室id Is Null And Nvl(c.病人来源, 0) = 0 Or c.适用科室id = P科室id And c.病人来源 = 1)) C
           Where Nvl(c.适用科室id, 0) = c.Top) A, 收费项目目录 C, 材料特性 B
    Where a.收费项目id = c.Id And c.服务对象 In (1, 3) And (c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.撤档时间 Is Null) And
          (c.站点 = v_Nodeno Or c.站点 Is Null) And c.Id = b.材料id(+);

  --分解时间 
  Function f_Calc次数分解时间
  (
    次数_In     In 病人医嘱记录.总给予量%Type,
    开始时间_In In Date,
    终止时间_In In Date,
    执行时间_In In 病人医嘱记录.执行时间方案%Type,
    频率间隔_In In 病人医嘱记录.频率间隔%Type,
    间隔单位_In In 病人医嘱记录.间隔单位%Type
  ) Return Varchar2 Is
    v_Detailtime Varchar2(4000);
    n_First      Number(1);
    v_First      病人医嘱记录.执行时间方案%Type;
    v_Normal     病人医嘱记录.执行时间方案%Type;
    v_Mtime      Varchar(100);
    v_Rtime      Varchar(100);
    v_Curtime    Date;
    v_Tmptime    Date;
    n_Cnt        Number; --计数器 
    --求某时间所在周的星期一的日期 
    Function Getweekbase(v_Time Date) Return Date Is
      v_Week Number(1);
    Begin
      v_Week := To_Number(To_Char(v_Time, 'D'));
      v_Week := v_Week - 1;
      If v_Week = 0 Then
        v_Week := 7;
      End If;
      Return(Trunc(v_Time - (v_Week - 1)));
    End;
  Begin
  
    n_Cnt := 0;
  
    --执行时间方案基准 
    If Nvl(Instr(执行时间_In, ','), 0) > 0 Then
      v_First  := Substr(执行时间_In, 1, Instr(执行时间_In, ',') - 1);
      v_Normal := Substr(执行时间_In, Instr(执行时间_In, ',') + 1);
    Else
      v_First  := Null;
      v_Normal := 执行时间_In;
    End If;
  
    If 间隔单位_In = '周' Then
      v_Curtime := Getweekbase(开始时间_In); --按周执行时在医嘱开始那周的星期一作为基准 
    Else
      v_Curtime := 开始时间_In;
    End If;
    If 间隔单位_In = '周' Then
      If v_First Is Not Null Then
        If v_Curtime = Getweekbase(开始时间_In) Then
          n_First := 1;
        End If;
      End If;
      While v_Curtime <= 终止时间_In And n_Cnt < 次数_In Loop
        If Nvl(n_First, 0) = 1 Then
          v_Rtime := v_First || '-';
        Else
          v_Rtime := v_Normal || '-';
        End If;
        n_First := 0;
        --1/8:00-3/8:00-5/8:00 
        While v_Rtime Is Not Null Loop
          v_Mtime   := Substr(v_Rtime, 1, Instr(v_Rtime, '-') - 1);
          v_Tmptime := v_Curtime + To_Number(Substr(v_Mtime, 1, Instr(v_Mtime, '/') - 1)) - 1;
          v_Mtime   := Substr(v_Mtime, Instr(v_Mtime, '/') + 1);
          If Instr(v_Mtime, ':') = 0 Then
            v_Mtime := v_Mtime || ':00';
          End If;
          v_Tmptime := Trunc(v_Tmptime) + (To_Date(v_Mtime, 'HH24:MI:SS') - Trunc(To_Date(v_Mtime, 'HH24:MI:SS')));
          If v_Tmptime >= 开始时间_In And v_Tmptime <= 终止时间_In Then
            v_Detailtime := v_Detailtime || ',' || To_Char(v_Tmptime, 'YYYY-MM-DD HH24:MI:SS');
            n_Cnt        := n_Cnt + 1;
            If n_Cnt >= 次数_In Then
              Exit;
            End If;
          Elsif v_Tmptime > 终止时间_In Then
            Exit;
          End If;
          v_Rtime := Substr(v_Rtime, Instr(v_Rtime, '-') + 1);
        End Loop;
        v_Curtime := Trunc(v_Curtime + 7);
      End Loop;
    Elsif 间隔单位_In = '天' Then
      If v_First Is Not Null Then
        If Trunc(开始时间_In) = Trunc(开始时间_In) Then
          n_First := 1;
        End If;
      End If;
      While v_Curtime <= 终止时间_In And n_Cnt < 次数_In Loop
        If Nvl(n_First, 0) = 1 Then
          v_Rtime := v_First || '-';
        Else
          v_Rtime := v_Normal || '-';
        End If;
        n_First := 0;
        If 频率间隔_In = 1 Then
          --8:00-12:00-14:00；8-12-14 
          While v_Rtime Is Not Null Loop
            v_Mtime := Substr(v_Rtime, 1, Instr(v_Rtime, '-') - 1);
            If Instr(v_Mtime, ':') = 0 Then
              v_Mtime := v_Mtime || ':00';
            End If;
            v_Tmptime := Trunc(v_Curtime) + (To_Date(v_Mtime, 'HH24:MI:SS') - Trunc(To_Date(v_Mtime, 'HH24:MI:SS')));
            If v_Tmptime >= 开始时间_In And v_Tmptime <= 终止时间_In Then
            
              v_Detailtime := v_Detailtime || ',' || To_Char(v_Tmptime, 'YYYY-MM-DD HH24:MI:SS');
              n_Cnt        := n_Cnt + 1;
              If n_Cnt >= 次数_In Then
                Exit;
              
              End If;
            Elsif v_Tmptime > 终止时间_In Then
              Exit;
            End If;
            v_Rtime := Substr(v_Rtime, Instr(v_Rtime, '-') + 1);
          End Loop;
        Else
          --1/8:00-1/15:00-2/9:00 
          While v_Rtime Is Not Null Loop
            v_Mtime   := Substr(v_Rtime, 1, Instr(v_Rtime, '-') - 1);
            v_Tmptime := v_Curtime + To_Number(Substr(v_Mtime, 1, Instr(v_Mtime, '/') - 1)) - 1;
            v_Mtime   := Substr(v_Mtime, Instr(v_Mtime, '/') + 1);
            If Instr(v_Mtime, ':') = 0 Then
              v_Mtime := v_Mtime || ':00';
            End If;
            v_Tmptime := Trunc(v_Tmptime) + (To_Date(v_Mtime, 'HH24:MI:SS') - Trunc(To_Date(v_Mtime, 'HH24:MI:SS')));
            If v_Tmptime >= 开始时间_In And v_Tmptime <= 终止时间_In Then
            
              v_Detailtime := v_Detailtime || ',' || To_Char(v_Tmptime, 'YYYY-MM-DD HH24:MI:SS');
              n_Cnt        := n_Cnt + 1;
              If n_Cnt >= 次数_In Then
                Exit;
              
              End If;
            Elsif v_Tmptime > 终止时间_In Then
              Exit;
            End If;
            v_Rtime := Substr(v_Rtime, Instr(v_Rtime, '-') + 1);
          End Loop;
        End If;
        v_Curtime := Trunc(v_Curtime + 频率间隔_In); --因为Loop条件注意要取整 
      End Loop;
    Elsif 间隔单位_In = '小时' Then
      --10:00-20:00-40:00；10-20-40；02:30 
      While v_Curtime <= 终止时间_In And n_Cnt < 次数_In Loop
      
        v_Rtime := 执行时间_In || '-';
        While v_Rtime Is Not Null Loop
          v_Mtime := Substr(v_Rtime, 1, Instr(v_Rtime, '-') - 1);
          If Instr(v_Mtime, ':') = 0 Then
            v_Tmptime := v_Curtime + (To_Number(v_Mtime) - 1) / 24;
          Else
            v_Tmptime := v_Curtime + (To_Number(Substr(v_Mtime, 1, Instr(v_Mtime, ':') - 1)) - 1) / 24 +
                         To_Number(Substr(v_Mtime, Instr(v_Mtime, ':') + 1)) / 60 / 24;
          End If;
        
          If v_Tmptime >= 开始时间_In And v_Tmptime <= 终止时间_In Then
          
            v_Detailtime := v_Detailtime || ',' || To_Char(v_Tmptime, 'YYYY-MM-DD HH24:MI:SS');
            n_Cnt        := n_Cnt + 1;
            If n_Cnt >= 次数_In Then
              Exit;
            End If;
          
          Elsif v_Tmptime > 终止时间_In Then
            Exit;
          End If;
          v_Rtime := Substr(v_Rtime, Instr(v_Rtime, '-') + 1);
        End Loop;
        v_Curtime := v_Curtime + 频率间隔_In / 24;
      End Loop;
    Elsif 间隔单位_In = '分钟' Then
      --无执行时间 
      While v_Curtime <= 终止时间_In And n_Cnt < 次数_In Loop
        v_Tmptime := v_Curtime;
        If v_Tmptime >= 开始时间_In And v_Tmptime <= 终止时间_In Then
          v_Detailtime := v_Detailtime || ',' || To_Char(v_Tmptime, 'YYYY-MM-DD HH24:MI:SS');
          n_Cnt        := n_Cnt + 1;
          If n_Cnt >= 次数_In Then
            Exit;
          End If;
        Elsif v_Tmptime > 终止时间_In Then
          Exit;
        End If;
        v_Curtime := v_Curtime + 频率间隔_In / (24 * 60);
      End Loop;
    End If;
    If v_Detailtime Is Not Null Then
      v_Detailtime := Substr(v_Detailtime, 2);
    End If;
    Return(v_Detailtime);
  End;

  --大于指定数的整数 
  Function f_Intex(Num_In In Number) Return Number Is
    n_Num Number;
  Begin
    Select Round(Num_In) Into n_Num From Dual;
    If Num_In > n_Num Then
      n_Num := n_Num + 1;
    End If;
    Return(n_Num);
  End;

  Procedure Get医嘱诊断对应(Pd Number) As
    --  pd 为记录集 rsap 的下标索引值 
    v_序号串 Varchar2(4000);
    v_诊断串 Varchar2(4000);
  Begin
    v_序号串 := Rsap(Pd).Diag_Nums;
    If v_序号串 Is Not Null Then
      v_序号串 := ',' || v_序号串 || ',';
      For I In 1 .. Rsdiag.Count Loop
        If Instr(v_序号串, ',' || Rsdiag(I).Diag_Num || ',') > 0 Then
          v_诊断串 := v_诊断串 || ',' || Rsdiag(I).诊断id;
        End If;
      End Loop;
      If v_诊断串 Is Not Null Then
        Rs诊断医嘱.Extend;
        Rs诊断医嘱(Rs诊断医嘱.Count).诊断ids := Substr(v_诊断串, 2);
        Rs诊断医嘱(Rs诊断医嘱.Count).医嘱id := Rsap(Pd).Id;
      End If;
    End If;
  End;

  Procedure Get真实的诊断id(Pd Number) As
    --根据 rsdiag 处理诊断数据，pd 入参为记录的下标索引值
    --如果不存在，则直接保存一条新诊断
    n_记录id Number(18);
  
  Begin
    If Rsdiag(Pd).诊断类型 = 1 And Rsdiag(Pd).Icd10_Code Is Not Null Then
      Select Max(a.Id)
      Into n_记录id
      From 病人诊断记录 A, 疾病编码目录 B
      Where a.疾病id = b.Id And a.病人id = r_Base.病人id And a.主页id = n_就诊id And a.诊断类型 = 1 And b.编码 = Rsdiag(Pd).Icd10_Code And
            Nvl(a.录入次序, '01') = '01' And b.类别 = 'D';
    Elsif Rsdiag(Pd).诊断类型 = 11 And Rsdiag(Pd).Csd_Code Is Not Null Then
      Select Max(a.Id)
      Into n_记录id
      From 病人诊断记录 A, 疾病编码目录 B
      Where a.疾病id = b.Id And a.病人id = r_Base.病人id And a.主页id = n_就诊id And a.诊断类型 = 11 And b.编码 = Rsdiag(Pd).Csd_Code And
            Nvl(a.录入次序, '01') = '01' And b.类别 = 'B';
    Else
      --如果是自由录入则判断文本即可以
      Select Max(a.Id)
      Into n_记录id
      From 病人诊断记录 A
      Where a.病人id = r_Base.病人id And a.主页id = n_就诊id And Instr(a.诊断描述, Rsdiag(Pd).Dz_Note) > 0 And
            Nvl(a.录入次序, '01') = '01';
    End If;
  
    If n_记录id Is Not Null Then
      Rsdiag(Pd).诊断id := n_记录id;
      Rsdiag(Pd).是否存在 := 1;
    Else
      --产生序列备用
      Select 病人诊断记录_Id.Nextval Into Rsdiag(Pd).诊断id From Dual;
      Rsdiag(Pd).是否存在 := 0;
    End If;
  End;

  Procedure Get诊断新增 As
    --功能：获取需要新插入的诊断的准备数据，缓存到 Rsdiagnew 记录集中
    n_中医次序 Number(3);
    n_西医次序 Number(3);
  Begin
    For I In 1 .. Rsdiag.Count Loop
      If Nvl(Rsdiag(I).是否存在, 0) = 0 Then
        If n_中医次序 Is Null Then
          n_中医次序 := 0;
          n_西医次序 := 0;
          For R In (Select a.诊断类型, Nvl(Max(a.诊断次序), 0) 诊断次序
                    From 病人诊断记录 A
                    Where a.病人id = r_Base.病人id And a.主页id = n_就诊id And a.记录来源 = 3 And a.诊断类型 In (1, 11) And
                          Nvl(a.录入次序, '01') = '01'
                    Group By a.诊断类型) Loop
          
            If r.诊断类型 = 11 Then
              n_中医次序 := r.诊断次序;
            Else
              n_西医次序 := r.诊断次序;
            End If;
          End Loop;
        End If;
      
        If Rsdiag(I).诊断类型 = 1 And Rsdiag(I).Icd10_Code Is Not Null Then
          Select Max(b.Id)
          Into Rsdiag(I).疾病id
          From 疾病编码目录 B
          Where b.编码 = Rsdiag(I).Icd10_Code And b.类别 = 'D';
          n_西医次序 := n_西医次序 + 1;
          Rsdiag(I).诊断次序 := n_西医次序;
          Rsdiag(I).诊断描述 := '(' || Rsdiag(I).Icd10_Code || ')' || Rsdiag(I).Dz_Note;
        Elsif Rsdiag(I).诊断类型 = 11 And Rsdiag(I).Csd_Code Is Not Null Then
          Select Max(b.Id)
          Into Rsdiag(I).疾病id
          From 疾病编码目录 B
          Where b.编码 = Rsdiag(I).Csd_Code And b.类别 = 'B';
          n_中医次序 := n_中医次序 + 1;
          Rsdiag(I).诊断次序 := n_中医次序;
          Rsdiag(I).诊断描述 := '(' || Rsdiag(I).Csd_Code || ')' || Rsdiag(I).Dz_Note;
          If Rsdiag(I).Syndrome Is Not Null Then
            Rsdiag(I).诊断描述 := Rsdiag(I).诊断描述 || '(' || Rsdiag(I).Syndrome || ')';
          End If;
        Else
          --自由录入诊断
          If Rsdiag(I).诊断类型 = 1 Then
            n_西医次序 := n_西医次序 + 1;
            Rsdiag(I).诊断次序 := n_西医次序;
          Else
            n_中医次序 := n_中医次序 + 1;
            Rsdiag(I).诊断次序 := n_中医次序;
          End If;
          Rsdiag(I).诊断描述 := Rsdiag(I).Dz_Note;
        End If;
      
        -- Zl_病人诊断记录_Insert(n_病人id, n_就诊id, 3, Null, r_Dz(R).诊断类型, r_Dz(R).疾病id, Null, Null, r_Dz(R).诊断描述, Null, Null, 0,
        --r_Dz(R).记录日期, r_Dz(R).医嘱ids, r_Dz(R).诊断次序);
        Rsdiagnew.Extend;
        Rsdiagnew(Rsdiagnew.Count) := Rsdiag(I);
      End If;
    End Loop;
  End;

  Procedure Getmore处方申请(Pd Number) As
    --pd 当前给药行的行下标
    --更新相关id，收集诊断序号
    v_诊断序号s Varchar2(4000);
    v_诊断序号  Varchar2(4000);
  Begin
    For I In 1 .. Pd Loop
      If Rsap(Pd - I).Serial_Num = Rsap(Pd).Serial_Num Then
        Rsap(Pd - I).相关id := Rsap(Pd).Id;
        If Rsap(Pd - I).Diag_Nums Is Not Null Then
          --这里面可能有重复的，要去重复
          v_诊断序号s := v_诊断序号s || ',' || Rsap(Pd - I).Diag_Nums;
        End If;
      Else
        Rsap(Pd - I + 1).Firstrow := 1;
        Exit;
      End If;
    End Loop;
  
    If v_诊断序号s Is Not Null Then
      Select f_List2str(Cast(Collect(a.诊断序号 || '') As t_Strlist), ',') 诊断序号
      Into v_诊断序号
      From (Select a.诊断序号
             From (Select /*+cardinality(b,10) */
                     b.Column_Value 诊断序号
                    From Table(f_Str2list(v_诊断序号s)) B) A
             Where a.诊断序号 Is Not Null
             Group By a.诊断序号) A;
      Rsap(Pd).Diag_Nums := v_诊断序号;
    End If;
  
    Rsap(Pd).Apply_Id := r_Base.Apply_Id;
    Rsap(Pd).Apply_Type := r_Base.Apply_Type;
    Rsap(Pd).医生嘱托 := r_Base.医生嘱托;
    Rsap(Pd).诊疗项目id := r_Base.诊疗项目id; -- N
    Rsap(Pd).执行科室id := r_Base.执行科室id; --  N
    Rsap(Pd).总给予量 := r_Base.总给予量; -- N  
    Rsap(Pd).紧急标志 := Rsap(Pd - 1).紧急标志;
    Rsap(Pd).执行频次 := Rsap(Pd - 1).执行频次;
    Rsap(Pd).频率次数 := Rsap(Pd - 1).频率次数;
    Rsap(Pd).频率间隔 := Rsap(Pd - 1).频率间隔;
    Rsap(Pd).间隔单位 := Rsap(Pd - 1).间隔单位;
    Rsap(Pd).执行时间方案 := Rsap(Pd - 1).执行时间方案;
    Rsap(Pd).天数 := Rsap(Pd - 1).天数;
  
  End;

  Function Get医嘱内容(Pd Number) Return Varchar2 As
    --检查检验项目组织医嘱内容
    --Pd 主医嘱对应的行索引下标
    v_内容    Varchar2(4000);
    v_部位    Varchar2(4000);
    v_方法    Varchar2(4000);
    v_部位_前 Varchar2(4000);
    --n_材料id  Number(18);
    --v_管码    Varchar2(1000);
  Begin
    If Rsap(Pd).Apply_Type = 3 Then
      --检验  
      v_内容 := '(' || Rsap(Pd).标本部位 || ')';
      For I In 1 .. Pd Loop
        If Rsap(Pd).Id = Rsap(Pd - I).相关id Then
          If I = 1 Then
            v_内容 := Rsap(Pd - I).医嘱内容 || v_内容;
          Else
            v_内容 := Rsap(Pd - I).医嘱内容 || ',' || v_内容;
          End If;
        Else
          Exit;
        End If;
      End Loop;
      --Rsap(Pd).管码卫材id := n_材料id;
      --Rsap(Pd).管码 := v_管码;
    Elsif Rsap(Pd).Apply_Type = 5 Then
      --检查      
      For I In Pd + 1 .. Rsap.Count Loop
        If Rsap(Pd).Id = Rsap(I).相关id Then
          If v_部位_前 <> Rsap(I).标本部位 And v_部位_前 Is Not Null Then
            v_部位 := v_部位 || ',' || v_部位_前 || '(' || Substr(v_方法, 2) || ')';
            v_方法 := Null;
          End If;
          v_部位_前 := Rsap(I).标本部位;
          v_方法    := v_方法 || ',' || Rsap(I).检查方法;
        Else
          Exit;
        End If;
      End Loop;
      If v_部位_前 Is Not Null Then
        v_部位 := v_部位 || ',' || v_部位_前 || '(' || Substr(v_方法, 2) || ')';
      End If;
      v_内容 := Rsap(Pd).医嘱内容 || ':' || Substr(v_部位, 2);
    End If;
    Return v_内容;
  End;

  Procedure Get检验条码 As
    --根据 Rs条码 信息计算生成的 样本条码  
  
    Rs条码tmp t_Price := t_Price();
    n_存在    Number(1);
  
    Function 生成样本条码
    (
      P申请标识 Varchar2,
      P项目id   Number
    ) Return Varchar2 As
      n_Key医嘱id Number(18);
      v_样本条码  Varchar2(1000);
    Begin
      For I In 1 .. Rsap.Count Loop
        If P申请标识 = Rsap(I).Apply_Id Then
          n_Key医嘱id := Nvl(Rsap(I).相关id, Rsap(I).Id);
          Exit;
        End If;
      End Loop;
      v_样本条码 := Zl_Cis_Nextno('125', n_Key医嘱id, P项目id);
      Return v_样本条码;
    End;
  Begin
    If n_是否生成条码 = 1 Then
      For I In 1 .. Rs条码.Count Loop
        Select Max(a.试管编码) Into Rs条码(I).管码 From 诊疗项目目录 A Where a.Id = Rs条码(I).检验项目id;
        If Rs条码(I).管码 Is Not Null Then
          For R In (Select 材料id From 采血管类型 Where 材料id Is Not Null And 编码 = Rs条码(I).管码) Loop
            Rs条码(I).管码卫材id := r.材料id;
            Exit;
          End Loop;
        End If;
        Rs条码(I).紧急标志 := Nvl(Rs条码(I).紧急标志, 0);
        Rs条码(I).婴儿 := Nvl(Rs条码(I).婴儿, 0);
      End Loop;
    
      --对于检验项目一个申请元素就表示一组医嘱
      For I In 1 .. Rs条码.Count Loop
        n_存在 := 0;
        --先从已生成的条码中查找，未找到则重新提取
        For J In 1 .. Rs条码tmp.Count Loop
        
          If Rs条码(I)
           .检验项目id <> Rs条码tmp(J).检验项目id And Rs条码(I).检验科室id = Rs条码tmp(J).检验科室id And Rs条码(I).采集项目id = Rs条码tmp(J).采集项目id And Rs条码(I)
             .采集科室id = Rs条码tmp(J).采集科室id And Rs条码(I).采集标本 = Rs条码tmp(J).采集标本 And Rs条码(I).紧急标志 = Rs条码tmp(J).紧急标志 And Rs条码(I)
             .婴儿 = Rs条码tmp(J).婴儿 And Rs条码(I).管码 = Rs条码tmp(J).管码 Then
          
            n_存在 := 1;
            Rs条码(I).样本条码 := Rs条码tmp(J).样本条码;
            Exit;
          End If;
        End Loop;
      
        If n_存在 = 0 Then
          Rs条码(I).样本条码 := 生成样本条码(Rs条码(I).Apply_Id, Rs条码(I).检验项目id); ----此时是假条码，条码依赖医嘱id和项目id,后面再生成真实的条码
          Rs条码tmp.Extend;
          Rs条码tmp(Rs条码tmp.Count) := Rs条码(I);
        End If;
      End Loop;
    End If;
  End;

Begin
  --代码折叠，基本信息解析
  If '代码折叠' = '代码折叠' Then
    --解析入参
    j_Tmp             := Pljson(Json_In);
    v_Input           := j_Tmp.Get_Pljson('input');
    r_Base.病人id     := v_Input.Get_Number('pati_id');
    r_Base.挂号单     := v_Input.Get_String('visit_no');
    r_Base.病人来源   := v_Input.Get_Number('pati_source');
    r_Base.姓名       := v_Input.Get_String('pati_name');
    r_Base.性别       := v_Input.Get_String('pati_sex');
    r_Base.年龄       := v_Input.Get_String('pati_age');
    r_Base.病人科室id := v_Input.Get_Number('pati_deptid');
    r_Base.开嘱科室id := v_Input.Get_Number('apply_dept_id');
    r_Base.开嘱医生   := v_Input.Get_String('apply_doctor');
    r_Base.开嘱时间   := To_Date(v_Input.Get_String('apply_time'), 'yyyy-MM-dd HH24:MI:SS');
    v_Nodeno          := v_Input.Get_String('nodeno');
  
    j_Advicelist := v_Input.Get_Pljson_List('apply_list');
    Jl_Diag      := v_Input.Get_Pljson_List('diag_info');
  
    If r_Base.病人来源 = 4 Then
      n_是否生成条码 := Nvl(v_Input.Get_Number('lis_bar_code_tag'), 0);
    Else
      n_是否生成条码 := Nvl(zl_GetSysParameter(143), 0);
    End If;
    d_发送时间 := r_Base.开嘱时间 + 1 / 24 / 60 / 60;
  
    j_Tmp := Pljson();
  
    If Jl_Diag Is Not Null Then
      Rsdiag.Extend(Jl_Diag.Count);
      Select a.Id Into n_就诊id From 病人挂号记录 A Where a.No = r_Base.挂号单 And a.记录状态 = 1 And a.记录性质 = 1;
      For I In 1 .. Jl_Diag.Count Loop
        j_Tmp := Pljson(Jl_Diag.Get(I));
      
        Rsdiag(I).Diag_Num := j_Tmp.Get_String('diag_num'); --  C
        Rsdiag(I).诊断类型 := j_Tmp.Get_Number('diag_type'); -- N
        Rsdiag(I).Csd_Code := j_Tmp.Get_String('csd_code'); --  C
        Rsdiag(I).Icd10_Code := j_Tmp.Get_String('icd10_code'); --  C
        Rsdiag(I).Dz_Note := j_Tmp.Get_String('dz_note'); --  C
        Rsdiag(I).Syndrome := j_Tmp.Get_String('syndrome'); --  C
        Rsdiag(I).Disease_Time := j_Tmp.Get_String('disease_time'); --  C
        Rsdiag(I).Diagnostician := j_Tmp.Get_String('diagnostician'); --  C
      
        j_Tmp := Pljson();
      End Loop;
    End If;
  End If;

  For I In 1 .. j_Advicelist.Count Loop
    j_Adviceitem      := Pljson(j_Advicelist.Get(I));
    r_Base.Apply_Id   := j_Adviceitem.Get_String('apply_id');
    r_Base.Diag_Nums  := j_Adviceitem.Get_String('diag_nums');
    r_Base.Apply_Type := j_Adviceitem.Get_Number('apply_type');
    r_Base.紧急标志   := j_Adviceitem.Get_Number('emergency_tag');
  
    --治疗
    If r_Base.Apply_Type = 4 Then
      Rsap.Extend;
      Idx := Rsap.Count;
      n_序列id := n_序列id + 1;
      Rsap(Idx).Id := n_序列id;
      Rsap(Idx).Apply_Id := r_Base.Apply_Id;
      Rsap(Idx).Apply_Type := r_Base.Apply_Type;
      Rsap(Idx).Diag_Nums := r_Base.Diag_Nums;
      Rsap(Idx).紧急标志 := r_Base.紧急标志;
    
      j_Tmp := j_Adviceitem.Get_Pljson('cure_info');
      Rsap(Idx).执行频次 := j_Tmp.Get_String('frequency_name');
      Rsap(Idx).频率次数 := j_Tmp.Get_Number('frequency_times');
      Rsap(Idx).频率间隔 := j_Tmp.Get_Number('frequency_interval');
      Rsap(Idx).间隔单位 := j_Tmp.Get_String('interval_unit');
      Rsap(Idx).执行时间方案 := j_Tmp.Get_String('exetime_plane');
      Rsap(Idx).诊疗项目id := j_Tmp.Get_Number('cure_item_id');
      Rsap(Idx).执行科室id := j_Tmp.Get_Number('cure_exedept_id');
      Rsap(Idx).医生嘱托 := j_Tmp.Get_String('cure_doctor_note');
      Rsap(Idx).单次用量 := j_Tmp.Get_Number('cure_once_qunt');
      Rsap(Idx).总给予量 := j_Tmp.Get_Number('cure_total_qunt');
      Rsap(Idx).Firstrow := 1;
    
      j_Tmp := Pljson();
      --检验
    Elsif r_Base.Apply_Type = 3 Then
    
      j_Tmp             := j_Adviceitem.Get_Pljson('lis_info');
      v_Lisitems        := j_Tmp.Get_String('lis_items');
      n_序列id          := n_序列id + 1;
      r_Base.Id         := n_序列id;
      r_Base.执行科室id := j_Tmp.Get_Number('lis_exedept_id');
      r_Base.标本部位   := j_Tmp.Get_String('lis_spcm');
    
      --生成条码的数据准备
      Rs条码.Extend;
      Rs条码(Rs条码.Count).Apply_Id := r_Base.Apply_Id;
      Rs条码(Rs条码.Count).检验科室id := r_Base.执行科室id;
      Rs条码(Rs条码.Count).采集标本 := r_Base.标本部位;
      Rs条码(Rs条码.Count).婴儿 := 0;
      Rs条码(Rs条码.Count).紧急标志 := r_Base.紧急标志;
    
      For r_项目 In (Select /*+cardinality(b,10)*/
                    b.Column_Value 项目id
                   From Table(f_Str2list(v_Lisitems)) B) Loop
      
        Rsap.Extend;
        Idx := Rsap.Count;
        n_序列id := n_序列id + 1;
        Rsap(Idx).Id := n_序列id;
        Rsap(Idx).相关id := r_Base.Id;
        Rsap(Idx).诊疗项目id := r_项目.项目id;
        Rsap(Idx).执行科室id := r_Base.执行科室id;
        Rsap(Idx).标本部位 := r_Base.标本部位;
        Rsap(Idx).Apply_Id := r_Base.Apply_Id;
        Rsap(Idx).Apply_Type := r_Base.Apply_Type;
        Rsap(Idx).Diag_Nums := r_Base.Diag_Nums;
        Rsap(Idx).紧急标志 := r_Base.紧急标志;
        Rsap(Idx).执行频次 := '一次性';
      
        --一并采集用首行的检验项目
        If Rs条码(Rs条码.Count).检验项目id Is Null Then
          Rs条码(Rs条码.Count).检验项目id := r_项目.项目id;
        
          Rsap(Idx).Firstrow := 1;
        End If;
      End Loop;
    
      Rsap.Extend;
      Idx := Rsap.Count;
      Rsap(Idx).Id := r_Base.Id;
      Rsap(Idx).诊疗项目id := j_Tmp.Get_Number('lis_collect_item_id');
      Rsap(Idx).执行科室id := j_Tmp.Get_Number('lis_collect_exedept_id');
      Rsap(Idx).医生嘱托 := j_Tmp.Get_String('lis_doctor_note');
      Rsap(Idx).标本部位 := r_Base.标本部位;
      Rsap(Idx).Apply_Id := r_Base.Apply_Id;
      Rsap(Idx).Apply_Type := r_Base.Apply_Type;
      Rsap(Idx).Diag_Nums := r_Base.Diag_Nums;
      Rsap(Idx).紧急标志 := r_Base.紧急标志;
      Rsap(Idx).执行频次 := '一次性';
    
      Rs条码(Rs条码.Count).采集项目id := Rsap(Idx).诊疗项目id;
      Rs条码(Rs条码.Count).采集科室id := Rsap(Idx).执行科室id;
    
      j_Tmp := Pljson();
      --检查
    Elsif r_Base.Apply_Type = 5 Then
    
      j_Tmp := j_Adviceitem.Get_Pljson('pacs_info');
      Rsap.Extend;
      Idx := Rsap.Count;
      n_序列id := n_序列id + 1;
      Rsap(Idx).Id := n_序列id;
      r_Base.Id := n_序列id;
      Rsap(Idx).诊疗项目id := j_Tmp.Get_Number('pacs_item_id');
      Rsap(Idx).执行科室id := j_Tmp.Get_Number('pacs_exedept_id');
      Rsap(Idx).医生嘱托 := j_Tmp.Get_String('pacs_doctor_note');
    
      Rsap(Idx).Apply_Id := r_Base.Apply_Id;
      Rsap(Idx).Apply_Type := r_Base.Apply_Type;
      Rsap(Idx).Diag_Nums := r_Base.Diag_Nums;
      Rsap(Idx).紧急标志 := r_Base.紧急标志;
      Rsap(Idx).执行频次 := '一次性';
      Rsap(Idx).Firstrow := 1;
      --部位信息
    
      Jl_Tmp := j_Tmp.Get_Pljson_List('pacs_part_list');
    
      If Jl_Tmp Is Not Null Then
        For J In 1 .. Jl_Tmp.Count Loop
        
          j_Tmp := Pljson(Jl_Tmp.Get(J));
          Rsap.Extend;
          Idx := Rsap.Count;
          n_序列id := n_序列id + 1;
          Rsap(Idx).Id := n_序列id;
          Rsap(Idx).相关id := r_Base.Id;
          Rsap(Idx).诊疗项目id := Rsap(Idx - 1).诊疗项目id;
          Rsap(Idx).执行科室id := Rsap(Idx - 1).执行科室id;
          Rsap(Idx).标本部位 := j_Tmp.Get_String('part_name');
          Rsap(Idx).检查方法 := j_Tmp.Get_String('part_way');
        
          Rsap(Idx).Apply_Id := r_Base.Apply_Id;
          Rsap(Idx).Apply_Type := r_Base.Apply_Type;
          Rsap(Idx).Diag_Nums := r_Base.Diag_Nums;
          Rsap(Idx).紧急标志 := r_Base.紧急标志;
          Rsap(Idx).执行频次 := '一次性';
        
          j_Tmp := Pljson();
        End Loop;
      End If;
    
      j_Tmp  := Pljson();
      Jl_Tmp := Pljson_List();
      --处方
    Elsif r_Base.Apply_Type = 1 Then
      n_Pre序号   := Null;
      Jl_Tmp      := j_Adviceitem.Get_Pljson_List('drug_info');
      n_Drug_Rows := Jl_Tmp.Count;
      For J In 1 .. n_Drug_Rows Loop
        j_Tmp  := Pljson(Jl_Tmp.Get(J));
        n_序号 := j_Tmp.Get_Number('serial_num');
        If n_序号 <> n_Pre序号 And n_Pre序号 Is Not Null Then
          --追加一行给药途径行，同时需要更新前面行的药品行的相关id，收集诊断序号          
          Rsap.Extend;
          Idx := Rsap.Count;
          n_序列id := n_序列id + 1;
          Rsap(Idx).Id := n_序列id;
          Rsap(Idx).Serial_Num := n_Pre序号;
          Getmore处方申请(Idx);
        End If;
        n_Pre序号 := n_序号;
      
        Rsap.Extend;
        Idx := Rsap.Count;
        n_序列id := n_序列id + 1;
        Rsap(Idx).Id := n_序列id;
        Rsap(Idx).Apply_Id := r_Base.Apply_Id;
        Rsap(Idx).Apply_Type := r_Base.Apply_Type;
        Rsap(Idx).Serial_Num := n_序号;
        Rsap(Idx).Diag_Nums := j_Tmp.Get_String('diag_nums');
        Rsap(Idx).紧急标志 := j_Tmp.Get_Number('emergency_tag');
        Rsap(Idx).执行频次 := j_Tmp.Get_String('frequency_name');
        Rsap(Idx).频率次数 := j_Tmp.Get_Number('frequency_times');
        Rsap(Idx).频率间隔 := j_Tmp.Get_Number('frequency_interval');
        Rsap(Idx).间隔单位 := j_Tmp.Get_String('interval_unit');
        Rsap(Idx).执行时间方案 := j_Tmp.Get_String('exetime_plane');
        Rsap(Idx).天数 := j_Tmp.Get_Number('user_day');
        Rsap(Idx).收费细目id := j_Tmp.Get_Number('drug_id'); -- N
        Rsap(Idx).执行科室id := j_Tmp.Get_Number('pharmacy_id'); -- N
        Rsap(Idx).单次用量 := j_Tmp.Get_Number('drug_once_qunt'); --  N
        Rsap(Idx).总给予量 := j_Tmp.Get_Number('drug_total_qunt'); -- N
        Rsap(Idx).医生嘱托 := j_Tmp.Get_String('doctor_note'); -- C
        Rsap(Idx).用药目的 := j_Tmp.Get_Number('drug_purpose'); --  N
        Rsap(Idx).用药理由 := j_Tmp.Get_String('drug_reason'); -- C
        Rsap(Idx).超量说明 := j_Tmp.Get_String('excs_desc'); -- C      
      
        r_Base.医生嘱托   := j_Tmp.Get_String('dripping_speed');
        r_Base.诊疗项目id := j_Tmp.Get_Number('use_item_id'); -- N
        r_Base.执行科室id := j_Tmp.Get_Number('use_exedept_id'); --  N
        r_Base.总给予量   := j_Tmp.Get_Number('use_count'); -- N  
      
        j_Tmp := Pljson();
      End Loop;
      If n_Pre序号 Is Not Null Then
        --追加一行给药途径行，同时需要更新前面行的药品行的相关id，收集诊断序号          
        Rsap.Extend;
        Idx := Rsap.Count;
        n_序列id := n_序列id + 1;
        Rsap(Idx).Id := n_序列id;
        Rsap(Idx).Serial_Num := n_Pre序号;
        Getmore处方申请(Idx);
      End If;
      Jl_Tmp := Pljson_List();
    End If;
    j_Adviceitem := Pljson();
  End Loop;

  --更新本地列表中其它信息
  Select Nvl(Max(序号), 0) Into n_序号 From 病人医嘱记录 Where 挂号单 = r_Base.挂号单;
  For I In 1 .. Rsap.Count Loop
    If Nvl(Rsap(I).收费细目id, 0) <> 0 Then
      --药品（处方和配方）
      Select a.名称, a.Id, a.类别, a.执行科室, c.产地, c.规格, b.剂量系数
      Into Rsap(I).标本部位,Rsap(I).诊疗项目id,Rsap(I).诊疗类别,Rsap(I).执行性质, v_产地, v_规格, n_剂量系数
      From 诊疗项目目录 A, 药品规格 B, 收费项目目录 C
      Where a.Id = b.药名id And b.药品id = c.Id And c.Id = Rsap(I).收费细目id;
      Rsap(I).医嘱内容 := Rsap(I).标本部位;
      If v_产地 Is Not Null Then
        Rsap(I).医嘱内容 := Rsap(I).医嘱内容 || '(' || v_产地 || ')';
      End If;
      If v_规格 Is Not Null Then
        Rsap(I).医嘱内容 := Rsap(I).医嘱内容 || ' ' || v_规格;
      End If;
      Rsap(I).执行性质 := 4; --药品通常为指定科室执行，对新门诊而言离院带药不会发过来
      Rsap(I).发送数次 := n_剂量系数 * Rsap(I).总给予量;
    Else
      Select a.名称, a.类别, a.执行科室, a.试管编码
      Into Rsap(I).医嘱内容,Rsap(I).诊疗类别,Rsap(I).执行性质,Rsap(I).管码
      From 诊疗项目目录 A
      Where a.Id = Rsap(I).诊疗项目id;
      Rsap(I).发送数次 := Nvl(Rsap(I).总给予量, 1);
    End If;
  
    n_序号 := n_序号 + 1;
    Rsap(I).序号 := n_序号;
    Rsap(I).病人id := r_Base.病人id;
    Rsap(I).挂号单 := r_Base.挂号单;
    Rsap(I).病人来源 := r_Base.病人来源;
    Rsap(I).姓名 := r_Base.姓名;
    Rsap(I).性别 := r_Base.性别;
    Rsap(I).年龄 := r_Base.年龄;
    Rsap(I).病人科室id := r_Base.病人科室id;
    Rsap(I).开嘱科室id := r_Base.开嘱科室id;
    Rsap(I).开嘱医生 := r_Base.开嘱医生;
    Rsap(I).开嘱时间 := r_Base.开嘱时间;
    Rsap(I).开始执行时间 := r_Base.开嘱时间;
    --Rsap(I).婴儿 := 0; 为null 正常ZLHIS导航台中开出都是0目前可用此作区分，Rsap(I).执行标记 也让他为null
    Rsap(I).医嘱状态 := 1;
    Rsap(I).医嘱期效 := 1;
    Rsap(I).计价特性 := 0;
  
  End Loop;

  -------以上为基础数据提取部份---------
  -------数据加工-----------------------
  --产生医嘱序列
  l_医嘱id.Extend(n_序列id);
  --检查，检验项目构建医嘱内容，产生医嘱序列
  For I In 1 .. Rsap.Count Loop
    If Rsap(I).Apply_Type = 3 And Rsap(I).相关id Is Null Then
      Rsap(I).医嘱内容 := Get医嘱内容(I);
    Elsif Rsap(I).Apply_Type = 5 And Rsap(I).相关id Is Null Then
      Rsap(I).医嘱内容 := Get医嘱内容(I);
    End If;
    --生成序列
    n_医嘱id := Rsap(I).Id;
    Select 病人医嘱记录_Id.Nextval Into l_医嘱id(n_医嘱id) From Dual;
  End Loop;

  --替换为真实的医嘱id
  For I In 1 .. Rsap.Count Loop
    n_医嘱id := Rsap(I).Id;
    Rsap(I).Id := l_医嘱id(n_医嘱id);
    If Rsap(I).相关id Is Not Null Then
      n_医嘱id := Rsap(I).相关id;
      Rsap(I).相关id := l_医嘱id(n_医嘱id);
    End If;
  End Loop;

  --生成检验项目的条码，缓存在 rs条码 对象中，内部会用到真实的医嘱id，不能任意改顺序
  Get检验条码;

  --诊断数据处理
  For I In 1 .. Rsdiag.Count Loop
    Get真实的诊断id(I);
  End Loop;

  --医嘱保存数据插入调用ZLHIS的过程  
  --产生单据号，发送号，记录序号
  Select Zl_Cis_Nextno('10') Into n_发送号 From Dual;
  For I In 1 .. Rsap.Count Loop
    If I = 1 Then
      Select Zl_Cis_Nextno('13') Into v_No From Dual;
      Rsap(I).No := v_No;
      n_序号 := 1;
      Rsap(I).记录序号 := n_序号;
    Else
      If Rsap(I).Apply_Id = Rsap(I - 1).Apply_Id Then
        Rsap(I).No := Rsap(I - 1).No;
        n_序号 := n_序号 + 1;
        Rsap(I).记录序号 := n_序号;
      Else
        Select Zl_Cis_Nextno('13') Into v_No From Dual;
        Rsap(I).No := v_No;
        n_序号 := 1;
        Rsap(I).记录序号 := n_序号;
      End If;
    End If;
    Rsap(I).发送号 := n_发送号;
  
    --如果没得执行时间方案都作是执行一次
    If Rsap(I).执行时间方案 Is Null Then
      Rsap(I).首次时间 := Rsap(I).开始执行时间;
      Rsap(I).末次时间 := Rsap(I).首次时间;
    End If;
    --样本条码处理
    If Rsap(I).Apply_Type = 3 Then
      If Rsap(I).Firstrow = 1 Then
        For R In 1 .. Rs条码.Count Loop
          If Rs条码(R).Apply_Id = Rsap(I).Apply_Id And Rs条码(R).样本条码 Is Not Null Then
            Rsap(I).样本条码 := Rs条码(R).样本条码;
            Exit;
          End If;
        End Loop;
      Else
        Rsap(I).样本条码 := Rsap(I - 1).样本条码;
      End If;
    End If;
  
    --获取诊断对应
    Get医嘱诊断对应(I);
  End Loop;

  --获取需要新插入的诊断的准备数据，缓存到 Rsdiagnew 记录集中
  Get诊断新增;
  For I In 1 .. Rsdiagnew.Count Loop
    Zl_病人诊断记录_Insert(r_Base.病人id, n_就诊id, 3, Null, Rsdiagnew(I).诊断类型, Rsdiagnew(I).疾病id, Null, Null, Rsdiagnew(I).诊断描述,
                     Null, Null, 0, Sysdate, Null, Rsdiagnew(I).诊断次序, Null, Null,
                     To_Date(Rsdiagnew(I).Disease_Time, 'yyyy-mm-dd hh24:mi:ss'), Rsdiagnew(I).Diagnostician,
                     Rsdiagnew(I).诊断id);
  End Loop;

  For I In 1 .. Rsap.Count Loop
    Zl_病人医嘱记录_Insert_s(Rsap(I).Id, Rsap(I).相关id, Rsap(I).序号, Rsap(I).病人来源, Rsap(I).病人id, Null, 0, 1, 1, Rsap(I).诊疗类别,
                       Rsap(I).诊疗项目id, Rsap(I).收费细目id, Rsap(I).天数, Rsap(I).单次用量, Rsap(I).总给予量, Rsap(I).医嘱内容,
                       Rsap(I).医生嘱托, Rsap(I).标本部位, Rsap(I).执行频次, Rsap(I).频率次数, Rsap(I).频率间隔, Rsap(I).间隔单位,
                       Rsap(I).执行时间方案, Rsap(I).计价特性, Rsap(I).执行科室id, Rsap(I).执行性质, Rsap(I).紧急标志, Rsap(I).开始执行时间, Null,
                       Rsap(I).病人科室id, Rsap(I).开嘱科室id, Rsap(I).开嘱医生, Rsap(I).开嘱时间, Rsap(I).挂号单, Null, Rsap(I).检查方法,
                       Rsap(I).执行标记, Null, Null, Rsap(I).开嘱医生, Null, Rsap(I).用药目的, Rsap(I).用药理由, Null, Null,
                       Rsap(I).超量说明, Null, Rsap(I).配方id, Null, Null, Null, Null, Null, Rsap(I).姓名, Rsap(I).性别,
                       Rsap(I).年龄);
  End Loop;

  For I In 1 .. Rs诊断医嘱.Count Loop
    Zl_病人诊断医嘱_Insert(Rs诊断医嘱(I).医嘱id, Rs诊断医嘱(I).诊断ids);
  End Loop;

  For I In 1 .. Rsap.Count Loop
    Zl_门诊医嘱发送_Insert_s(Rsap(I).Id, n_发送号, 1, Rsap(I).No, Rsap(I).记录序号, Rsap(I).发送数次, Rsap(I).首次时间, Rsap(I).末次时间, d_发送时间,
                       Rsap(I).执行科室id, 0, Rsap(I).开嘱医生, Rsap(I).Firstrow, Rsap(I).样本条码, Null, 0, 0);
  End Loop;

  --获取出参相信息
  For I In 1 .. Rsap.Count Loop
    If I = 1 Then
      v_Out := v_Out || ',{"apply_id":"' || Rsap(I).Apply_Id || '"';
      v_Out := v_Out || ',"fee_no":"' || Rsap(I).No || '"';
      If Rsap(I).Apply_Type = 3 Then
        v_Out := v_Out || ',"lis_bar_code":"' || Rsap(I).样本条码 || '"';
      End If;
    Else
      If Rsap(I).Apply_Id <> Rsap(I - 1).Apply_Id Then
        v_Out  := v_Out || ',"order_list":[' || Substr(v_Tmp1, 2) || ']';
        v_Out  := v_Out || '}';
        v_Tmp1 := Null;
        v_Out  := v_Out || ',{"apply_id":"' || Rsap(I).Apply_Id || '"';
        v_Out  := v_Out || ',"fee_no":"' || Rsap(I).No || '"';
        If Rsap(I).Apply_Type = 3 Then
          v_Out := v_Out || ',"lis_bar_code":"' || Rsap(I).样本条码 || '"';
        End If;
      End If;
    End If;
    v_Tmp1 := v_Tmp1 || ',{"order_id":' || Rsap(I).Id;
    v_Tmp1 := v_Tmp1 || ',"order_related_id":' || Nvl(Rsap(I).相关id || '', 'null');
    v_Tmp1 := v_Tmp1 || ',"cisitem_id":' || Rsap(I).诊疗项目id;
    v_Tmp1 := v_Tmp1 || '}';
  End Loop;
  If v_Out Is Not Null Then
    v_Out := v_Out || ',"order_list":[' || Substr(v_Tmp1, 2) || ']';
    v_Out := v_Out || '}';
  End If;

  Json_Out := '{"output":{"code":1,"message":"成功","advice_list":[' || Substr(v_Out, 2) || ']}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Outsendapply;
/

Create Or Replace Procedure Zl_Cissvr_Revokeoutadvice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：新门诊病人取消申请（如注射输液类处方等），即ZLHIS作废医嘱
  --入参：Json_In:格式
  --  input
  --          operator_name     C 1 操作员姓名
  --          operator_code     C 1 操作员编号
  --          order_ids         C 1 主医嘱id,主医嘱id，逗号拼串，可支持一次传入多个批量取消

  --出参: Json_Out,格式如下
  --  output
  --    code                      N 1 应答吗：0-失败；1-成功
  --    message                   C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Json       Pljson;
  j_In         Pljson;
  v_医嘱ids    Varchar2(4000);
  v_操作员姓名 Varchar2(1000);
  v_操作员编号 Varchar2(1000);
  v_操作时间   Varchar2(300);
  n_Count      Number(2);
Begin
  --解析入参
  j_In         := Pljson(Json_In);
  j_Json       := j_In.Get_Pljson('input');
  v_医嘱ids    := j_Json.Get_String('order_ids');
  v_操作员姓名 := j_Json.Get_String('operator_name');
  v_操作员编号 := j_Json.Get_String('operator_code');

  Select To_Char(Sysdate, 'yyyy-mm-dd hh24:mi:ss') Into v_操作时间 From Dual;

  If Instr(v_医嘱ids, ',') > 0 Then
  
    Select Count(1)
    Into n_Count
    From 病人医嘱发送 A
    Where a.医嘱id In (Select /*+cardinality(j,10)*/
                      x.Id
                     From 病人医嘱记录 X, Table(f_Num2list(v_医嘱ids)) J
                     Where x.Id = j.Column_Value Or x.相关id = j.Column_Value) And a.执行状态 In (1, 3);
  
    If n_Count > 0 Then
      Json_Out := Zljsonout('当医嘱项目已执行或正在执行不能作废！');
      Return;
    End If;
  
    For R In (Select /*+cardinality(j,10) */
               j.Column_Value 医嘱id
              From Table(Cast(f_Num2list(v_医嘱ids) As Zltools.t_Numlist)) J) Loop
      Zl_病人医嘱记录_作废_s(Null, r.医嘱id, Null, Null, v_操作员姓名, v_操作员编号, v_操作时间);
    End Loop;
  
  Else
  
    Select Count(1)
    Into n_Count
    From 病人医嘱发送 A
    Where a.医嘱id In (Select x.Id From 病人医嘱记录 X Where x.Id = v_医嘱ids Or x.相关id = v_医嘱ids) And a.执行状态 In (1, 3);
    If n_Count > 0 Then
      Json_Out := Zljsonout('当医嘱项目已执行或正在执行不能作废！');
      Return;
    End If;
  
    Zl_病人医嘱记录_作废_s(Null, v_医嘱ids, Null, Null, v_操作员姓名, v_操作员编号, v_操作时间);
  End If;

  Json_Out := Zljsonout('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Cissvr_Revokeoutadvice;
/