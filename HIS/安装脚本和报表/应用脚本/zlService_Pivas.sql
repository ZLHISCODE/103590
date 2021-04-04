Create Or Replace Procedure Zl_Pivassvr_Checkorderroll
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：医嘱回退发送时静配中医嘱检查判断
  --入参：Json_In:格式
  --  input
  --     order_id           N 1 医嘱ID,主医嘱id
  --     send_no            N 1 发送号
  --     item_list[]
  --            order_id           N 1 医嘱ID,主医嘱id
  --            send_no            N 1 发送号

  --出参: Json_Out,格式如下
  --  output
  --    code                N 1 应答吗：0-失败；1-成功
  --    message             C 1 应答消息：失败时返回具体的错误信息
  --    pivas_ids           C 1 要销帐的输液记录id串
  ---------------------------------------------------------------------------
  n_医嘱id   Number;
  n_发送号   Number;
  j_Json     Pljson;
  j_Tmp      Pljson;
  n_Tmp      Number;
  v_配液ids  Varchar2(32767);
  n_List_Cnt Number;
  j_Jsonlist Pljson_List;
Begin
  --解析入参
  j_Tmp      := Pljson(Json_In);
  j_Json     := j_Tmp.Get_Pljson('input');
  j_Jsonlist := j_Json.Get_Pljson_List('item_list');

  If Not j_Jsonlist Is Null Then
    n_List_Cnt := j_Jsonlist.Count;
  Else
    n_List_Cnt := 1;
  End If;

  For I In 1 .. n_List_Cnt Loop
  
    If Not j_Jsonlist Is Null Then
      j_Json := Pljson();
      j_Json := Pljson(j_Jsonlist.Get(I));
    End If;
  
    n_医嘱id := j_Json.Get_Number('order_id');
    n_发送号 := j_Json.Get_Number('send_no');
  
    --检查是否是输液配液记录，并是否已经锁定
    Select Decode(Max(是否锁定), 1, 1, 0) Into n_Tmp From 输液配药记录 Where 医嘱id = n_医嘱id And 发送号 = n_发送号;
    If n_Tmp = 1 Then
      Json_Out := Zljsonout('当前处理的是输液药品医嘱，已经被输液配置中心锁定，不能回退发送。');
      Return;
    Elsif n_Tmp = 0 Then
      --只对状态=1(未配药)的记录处理，如果已经配药了，则通过销账方式处理
      Select Count(ID)
      Into n_Tmp
      From 输液配药记录
      Where 操作状态 In (1, 10) And 医嘱id = n_医嘱id And 发送号 = n_发送号;
      If n_Tmp > 0 Then
        For R In (Select ID From 输液配药记录 Where 医嘱id = n_医嘱id And 发送号 = n_发送号) Loop
          v_配液ids := v_配液ids || ',' || r.Id;
        End Loop;
      End If;
    End If;
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"成功","pivas_ids":"' || Substr(v_配液ids, 2) || '"}}';
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Pivassvr_Checkorderroll;
/
Create Or Replace Procedure Zl_Pivassvr_Infusion_Update
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：输液配药记录及输液配药状态修正
  --入参： Json_In:格式
  --  input
  --    dispensing_id       C 0 配药ID，如果传入空，则advice_list[]必传
  --    type                N 1 类型0-销帐申请;1-销帐申请取消(删除);2-销帐申请审核
  --    operator_name       C 1 操作员姓名
  --    operator_notes      C 0 操作说明
  --    operator_time       C 1 操作时间：yyyy-mm-dd hh24:mi:ss
  --    apply_time          C 1 申请时间(销帐申请审核时有效)：yyyy-mm-dd hh24:mi:ss

  --    advice_list[]医嘱信息列表:传入了配药ID的或type<>0时，此列表无效  表示：已申请的批次一起取消或审核
  --        advice_id        N 1 医嘱Id
  --        send_no          N 1 发送号
  --出参: Json_Out,格式如下
  --  output
  --    code                 N 1 应答吗：0-失败；1-成功
  --    message              C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------

  j_Json       Pljson;
  j_Jsonlist   Pljson_List;
  o_Json       Pljson;
  n_配药id     Number(18);
  v_Tmp        Varchar2(4000);
  n_Count      Number(18);
  n_Temp       Number(18);
  n_操作状态   Number(2);
  v_操作员姓名 输液配药记录.操作人员%Type;
  v_操作说明   输液配药状态.操作说明%Type;
  d_Date       Date;
  n_操作类型   输液配药记录.操作状态%Type;
  n_医嘱id     输液配药记录.医嘱id%Type;
  n_发送号     输液配药记录.发送号%Type;
  d_Apply      Date;
  j_Json_Tmp   Pljson;
Begin
  --解析入参
  j_Json_Tmp := Pljson(Json_In);
  j_Json     := j_Json_Tmp.Get_Pljson('input');
  v_Tmp      := j_Json.Get_String('dispensing_id');
  If Nvl(v_Tmp, '-') <> '-' Then
    n_配药id := To_Number(v_Tmp);
  End If;
  v_Tmp := j_Json.Get_String('operator_time');
  If Nvl(v_Tmp, '-') <> '-' Then
    d_Date := To_Date(v_Tmp, 'yyyy-mm-dd hh24:mi:ss');
  Else
    d_Date := Sysdate;
  End If;
  n_操作状态 := j_Json.Get_Number('type');

  If Nvl(n_操作状态, 0) = 0 Then
  
    v_操作员姓名 := j_Json.Get_String('operator_name');
    v_操作说明   := j_Json.Get_String('operator_notes');
  
    --销帐申请:
    If n_配药id = 0 Then
      --销帐申请时，配药ID必须传入 
      Json_Out := '{"output":{"code":0,"message":"配药ID未传入"}}';
      Return;
    End If;
    Select Count(1) Into n_Count From 输液配药状态 Where 配药id = n_配药id And 操作类型 = 9 And 操作时间 = d_Date;
    If n_Count = 0 Then
      Insert Into 输液配药状态
        (配药id, 操作类型, 操作人员, 操作时间, 操作说明)
      Values
        (n_配药id, 9, v_操作员姓名, d_Date, v_操作说明);
    End If;
    Update 输液配药记录 Set 操作人员 = v_操作员姓名, 操作时间 = d_Date, 操作状态 = 9 Where ID = n_配药id;
  End If;

  If Nvl(n_操作状态, 0) = 1 Then
    --销帐申请取消或删除
    If n_配药id Is Not Null Then
      Select 操作人员, 操作时间, 操作类型
      Into v_操作员姓名, d_Date, n_操作类型
      From (Select 操作人员, 操作时间, 操作类型
             From 输液配药状态
             Where 配药id = n_配药id And 操作类型 <> 9
             Order By 操作时间 Desc, 操作类型 Desc)
      Where Rownum < 2;
      Update 输液配药记录 Set 操作人员 = v_操作员姓名, 操作时间 = d_Date, 操作状态 = n_操作类型 Where ID = n_配药id;
    End If;
    j_Jsonlist := Pljson_List();
    j_Jsonlist := j_Json.Get_Pljson_List('advice_list');
    n_Count    := j_Jsonlist.Count;
    If n_Count = 0 Then
      Json_Out := '{"output":{"code":0,"message":"未传入医嘱信息或配方ID"}}';
      Return;
    End If;
    --暂未提供按配药批次取消的功能，所有已申请的批次一起取消
    For I In 1 .. n_Count Loop
      o_Json   := Pljson();
      o_Json   := Pljson(j_Jsonlist.Get(I));
      n_医嘱id := o_Json.Get_Number('advice_id');
      n_发送号 := o_Json.Get_Number('send_no');    
      For R In (Select d.Id From 输液配药记录 D Where d.医嘱id = Nvl(n_医嘱id, 0) And d.发送号 = Nvl(n_发送号, 0)) Loop
        Select 操作人员, 操作时间, 操作类型
        Into v_操作员姓名, d_Date, n_操作类型
        From (Select 操作人员, 操作时间, 操作类型
               From 输液配药状态
               Where 配药id = r.Id And 操作类型 <> 9
               Order By 操作时间 Desc, 操作类型 Desc)
        Where Rownum < 2;      
        Update 输液配药记录 Set 操作人员 = v_操作员姓名, 操作时间 = d_Date, 操作状态 = n_操作类型 Where ID = r.Id;
      End Loop;
    End Loop;
  End If;
  If Nvl(n_操作状态, 0) = 2 Then
    --销帐审核
    v_Tmp := j_Json.Get_String('apply_time');
    If Nvl(v_Tmp, '-') <> '-' Then
      d_Apply := To_Date(v_Tmp, 'yyyy-mm-dd hh24:mi:ss');
    Else
      d_Apply := Sysdate;
    End If;
    j_Jsonlist := Pljson_List();
    j_Jsonlist := j_Json.Get_Pljson_List('advice_list');
    n_Count    := j_Jsonlist.Count;
    If n_Count = 0 Then
      Json_Out := '{"output":{"code":0,"message":"未传入医嘱信息"}}';
      Return;
    End If;
  
    --暂未提供按配药批次取消的功能，所有已申请的批次一起取消
    For I In 1 .. n_Count Loop
      o_Json   := Pljson();
      o_Json   := Pljson(j_Jsonlist.Get(I));
      n_医嘱id := o_Json.Get_Number('advice_id');
      n_发送号 := o_Json.Get_Number('send_no');
      Select Nvl(Max(d.Id), 0)
      Into n_配药id
      From 输液配药记录 D
      Where d.Id = n_医嘱id And d.发送号 = n_发送号 And d.操作时间 = d_Apply And d.操作状态 = 9;
    
      If n_配药id <> 0 Then
        Select Count(1) Into n_Temp From 输液配药状态 Where 配药id = n_配药id And 操作类型 = 10 And 操作时间 = d_Date;
      
        If n_Temp = 0 Then
          Insert Into 输液配药状态 (配药id, 操作类型, 操作人员, 操作时间) Values (n_配药id, 10, v_操作员姓名, d_Date);
        End If;
        Update 输液配药记录 Set 操作人员 = v_操作员姓名, 操作时间 = d_Date, 操作状态 = 10 Where ID = n_配药id;
      End If;
    End Loop;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Pivassvr_Infusion_Update;
/

Create Or Replace Procedure Zl_Pivassvr_Isexsitinfusion
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：根据费用ID,判断对应的药品是否已经进入输液配药中心
  --入参：Json_In:格式
  --input
  --        advice_id   N 1 医嘱ID
  --        rcpdtl_ids  C 1 处方明细IDs（传入的费用IDs）,多个用逗号分离，如果传入了该节点以明细IDs为准，否则以医嘱ID为准。
  --        is_return   N 1 是否返回 进入配液中的费用id串
  --出参: Json_Out,格式如下
  --output
  --        code        N 1 应答吗：0-失败；1-成功
  --        message     C 1 应答消息：失败时返回具体的错误信息
  --        isexist     N 1 1-存在;0-不存在
  --        rcpdtl_ids  C 1 已经进入了配液中心的费用id串逗号分割
  ---------------------------------------------------------------------------
  j_Input      PLJson;
  j_Json       PLJson;
  n_Is_Return  Number(1);
  v_费用ids    Varchar2(4000);
  n_医嘱id     输液配药记录.医嘱id%Type;
  n_Count      Number(18);
  v_Rcpdtl_Ids Varchar2(32767);
Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_医嘱id    := j_Json.Get_Number('advice_id');
  v_费用ids   := j_Json.Get_String('rcpdtl_ids');
  n_Is_Return := j_Json.Get_Number('is_return');

  If v_费用ids Is Not Null Then
    If Nvl(n_Is_Return, 0) = 0 Then
      Select Count(1)
      Into n_Count
      From 输液配药内容 A, 药品收发记录 B
      Where a.收发id = b.Id And b.费用id In (Select Column_Value From Table(f_Num2List(v_费用ids))) And
            Instr(',8,9,10,', ',' || b.单据 || ',') > 0;
    End If;
  
    If Nvl(n_Is_Return, 0) = 1 Then
      For R In (Select Distinct b.费用id
                From 输液配药内容 A, 药品收发记录 B
                Where a.收发id = b.Id And b.费用id In (Select Column_Value From Table(f_Num2List(v_费用ids))) And
                      Instr(',8,9,10,', ',' || b.单据 || ',') > 0) Loop
        v_Rcpdtl_Ids := v_Rcpdtl_Ids || ',' || r.费用id;
      End Loop;
      If v_Rcpdtl_Ids Is Not Null Then
        n_Count := 1;
      Else
        n_Count := 0;
      End If;
    End If;
  Else
    Select Count(1)
    Into n_Count
    From 输液配药内容 A, 药品收发记录 B, 输液配药记录 C
    Where a.收发id = b.Id And b.医嘱id = n_医嘱id And a.记录id = c.Id And Rownum < 2;
  End If;

  If n_Count <> 0 Then
    n_Count := 1;
  End If;

  Json_Out := '{"output":{"code":1,"message":"成功","rcpdtl_ids":"' || Substr(v_Rcpdtl_Ids, 2) || '","isexist":' ||
              n_Count || '}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Pivassvr_Isexsitinfusion;
/


Create Or Replace Procedure Zl_Pivassvr_Getinfusion_Record
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：根据病人信息或处方号信息，获取进入输液配药记录中的处方ids
  --入参：Json_In:格式
  --    input
  --        pati_id                 N   1   病人ID
  --        pati_pageids            C   1   主页ID
  --        rcpdtl_ids              C   0   多个用逗号
  --        rcp_nos                 C   0   多个用逗号，rcpdtl_ids传入时，此参数无效
  --出参: Json_Out,格式如下
  --    output
  --        code                    N   1   应答吗：0-失败；1-成功
  --        message                 C   1   应答消息：失败时返回具体的错误信息
  --        rcpdtl_ids              C   1   处方明细id,目前传入的费用ID
  ---------------------------------------------------------------------------

  j_Json    PLJson;
  n_病人id  药品收发记录.病人id%Type;
  v_主页ids Varchar2(4000);
  n_费用id  药品收发记录.费用id%Type;
  Type t_输液数据 Is Ref Cursor;
  c_输液数据 t_输液数据;

  j_Json_Tmp PLJson;

  v_处方单  Varchar2(4000);
  v_Temp    Varchar2(32680);
  v_费用ids Varchar2(32680);
  c_费用ids Clob;
  n_Count   Number(18);
Begin
  --解析入参
  j_Json_Tmp := PLJson(Json_In);
  j_Json     := j_Json_Tmp.Get_Pljson('input');
  n_病人id   := j_Json.Get_Number('pati_id');
  v_主页ids  := j_Json.Get_String('pati_pageids');
  c_费用ids  := j_Json.Get_String('rcpdtl_ids');
  v_处方单   := j_Json.Get_Clob('rcp_nos');

  If v_处方单 Is Not Null Then
    If Instr(v_处方单, ',') > 0 Then
      Open c_输液数据 For
        Select Distinct 费用id
        From 药品收发记录 B1, 输液配药内容 C1
        Where B1.Id = C1.收发id And B1.No In (Select /*+cardinality(j,10) */
                                             Column_Value
                                            From Table(f_Str2List(v_处方单)) J) And
              Instr(',9,10,', ',' || B1.单据 || ',') > 0;
    Else
      Open c_输液数据 For
        Select Distinct 费用id
        From 药品收发记录 B1, 输液配药内容 C1
        Where B1.Id = C1.收发id And B1.No = v_处方单 And Instr(',9,10,', ',' || B1.单据 || ',') > 0;
    End If;
  Elsif c_费用ids Is Not Null Then
  
    If Length(c_费用ids) <= 4000 Then
    
      Open c_输液数据 For
        Select /*+cardinality(a,10) */
        Distinct B1.费用id
        From 药品收发记录 B1, 输液配药内容 C1, (Select Column_Value As 费用id From Table(f_Num2List(c_费用ids)) J) A
        Where a.费用id = B1.费用id And B1.Id = C1.收发id And Instr(',9,10,', ',' || B1.单据 || ',') > 0;
    Else
    
      v_费用ids := Null;
      Loop
        Exit When c_费用ids Is Null;
      
        If Length(c_费用ids) <= 4000 Then
          v_Temp    := Substr(c_费用ids, 1);
          c_费用ids := Null;
        Else
          n_Count   := Instr(c_费用ids, ',', 3900);
          v_Temp    := Substr(c_费用ids, 1, n_Count - 1);
          c_费用ids := Substr(c_费用ids, n_Count + 1);
        End If;
      
        For c_费用id In (Select /*+cardinality(a,10) */
                       Distinct B1.费用id
                       From 药品收发记录 B1, 输液配药内容 C1, (Select Column_Value As 费用id From Table(f_Num2List(v_Temp)) J) A
                       Where a.费用id = B1.费用id And B1.Id = C1.收发id And Instr(',9,10,', ',' || B1.单据 || ',') > 0) Loop
          If Instr(Nvl(v_费用ids, '') || ',', ',' || c_费用id.费用id || ',') = 0 Then
            v_费用ids := Nvl(v_费用ids, '') || ',' || c_费用id.费用id;
          End If;
        End Loop;
      End Loop;
      If Not v_费用ids Is Null Then
        v_费用ids := Substr(v_费用ids, 2);
      End If;
      Json_Out := '{"output":{"code":1,"message":"成功","rcpdtl_ids":"' || v_Temp || '"}}';
      Return;
    End If;
  Elsif Nvl(n_病人id, 0) <> 0 Then
    Open c_输液数据 For
      Select /*+cardinality(a,10) */
      Distinct B1.费用id
      From 药品收发记录 B1, 输液配药内容 C1
      Where B1.病人id = n_病人id And (Instr(',' || v_主页ids || ',', ',' || Nvl(B1.主页id, 0) || ',') > 0 Or v_主页ids Is Null) And
            B1.Id = C1.收发id And Instr(',9,10,', ',' || B1.单据 || ',') > 0;
  Else
    Json_Out := zlJsonOut('不能确定本次获取数据的条件，请检查！');
    Return;
  End If;
  v_费用ids := Null;
  Loop
    Fetch c_输液数据
      Into n_费用id;
    Exit When c_输液数据%NotFound;
  
    If v_费用ids Is Null Then
      v_费用ids := '' || n_费用id;
    Else
      v_费用ids := v_费用ids || ',' || n_费用id;
    End If;
  End Loop;
  Close c_输液数据;
  Json_Out := '{"output":{"code":1,"message":"成功","rcpdtl_ids":"' || v_费用ids || '"}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Pivassvr_Getinfusion_Record;
/

Create Or Replace Procedure Zl_Pivassvr_Newbill
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：医嘱发送后产生输液配药记录
  --入参：Json_In:格式
  --input      产生输液配药记录，按病人调用
  --  operator_name                   C 1 发送人(操作员姓名)
  --  operator_time                   C 1 发送时间
  --  pati_id                         N 1 病人id
  --  page_id                         N 1 主页ID
  --  pati_name                       C 1 姓名
  --  pati_sex                        C 1 性别
  --  pati_age                        C 1 年龄
  --  inpatient_num                   C 1 住院号
  --  pati_bed                        C 1 床号
  --  pati_wardarea_id                N 1 病人病区id
  --  pati_deptid                     N 1 病人科室id
  --  advice_list[]主医嘱，数组
  --    pivas_deptid                  N 1 静配中心id
  --    advice_id                     N 1 主医嘱ID(给药途径)
  --    advice_send_no                N 1 发送号
  --    effective_time                N 1 医嘱期效：0-长嘱，1-临嘱
  --    drug_method_id                N 1 给药途径id
  --    is_tpn                        N 1 是否tpn：0-不是，1-是
  --    advice_frequency              C 1 执行频次
  --    advice_drug_list[]给药途径对应的药嘱，数组
  --            advice_id             N 1 药嘱id
  --            advice_rcpno          C 1 药嘱发送产生的费用no
  --    advice_exetime_list[]医嘱执行时间，给药途径医嘱的执行时间，暂时提供该医嘱所有发送的时间，包括本次发送的执行时间。按发送号倒序组织数组数据


  --            advice_send_no        N 1 给药途径医嘱的发送号
  --            advice_require_time   C 1 要求时间

  --出参: Json_Out,格式如下
  --output
  --  code                          C 1 应答码：0-失败；1-成功
  --  message                       C 1 应答消息:失败时返回具体的错误信息
  ---------------------------------------------------------------------------

  j_Pivas        PLJson;
  j_Input        PLJson;
  Jl_Pid_Main    Pljson_List;
  j_Pid_Main     PLJson;
  Jl_Pid_Exetime Pljson_List;
  j_Pid_Exetime  PLJson;
  Jl_Pid_Drug    Pljson_List;
  j_Pid_Drug     PLJson;

  n_相关id_m     输液配药记录.医嘱id%Type;
  d_发送号_m     输液配药记录.发送号%Type;
  v_姓名_m       输液配药记录.姓名%Type;
  v_性别_m       输液配药记录.性别%Type;
  v_年龄_m       输液配药记录.年龄%Type;
  v_住院号_m     输液配药记录.住院号%Type;
  v_床号_m       输液配药记录.床号%Type;
  n_病人病区id_m 输液配药记录.病人病区id%Type;
  n_病人科室id_m 输液配药记录.病人科室id%Type;
  d_发送时间_m   输液配药记录.发送时间%Type;
  n_医嘱类型_m   Number(1);
  n_给药途径id_m 诊疗项目目录.Id%Type;
  n_病人id_m     输液配药记录.病人id%Type;
  n_是否tpn_m    Number(1);
  v_执行频次_m   Varchar2(100);
  n_主页id_m     输液配药记录.主页id%Type;
  v_执行时间s_m  Varchar2(4000);

  n_医嘱id_d  输液配药记录.医嘱id%Type;
  v_发送no_d  药品收发记录.No%Type;
  v_医嘱ids_d Varchar2(32767);

  n_配置中心id 输液配药记录.部门id%Type;
  v_核查人     输液配药记录.操作人员%Type;
  v_核查时间   输液配药记录.操作时间%Type;

  v_主医嘱   Varchar2(32767);
  v_药嘱     Varchar2(32767);
  v_医嘱信息 Varchar2(32767);
  c_医嘱信息 Clob;
Begin
  If Json_In Is Null Then
    Json_Out := zlJsonOut('未传入数据，请检查');
    Return;
  End If;

  j_Pivas := PLJson(Json_In);
  If j_Pivas Is Not Null Then
    --生成数据到静配中心
    j_Input := j_Pivas.Get_Pljson('input');
  
    v_核查人       := j_Input.Get_String('operator_name');
    v_核查时间     := To_Date(j_Input.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');
    d_发送时间_m   := To_Date(j_Input.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');
    n_病人id_m     := j_Input.Get_Number('pati_id');
    n_主页id_m     := j_Input.Get_Number('page_id');
    v_姓名_m       := j_Input.Get_String('pati_name');
    v_性别_m       := j_Input.Get_String('pati_sex');
    v_年龄_m       := j_Input.Get_String('pati_age');
    v_床号_m       := j_Input.Get_String('pati_bed');
    v_住院号_m     := j_Input.Get_Number('inpatient_num');
    n_病人病区id_m := j_Input.Get_Number('pati_wardarea_id');
    n_病人科室id_m := j_Input.Get_Number('pati_deptid');
  
    --医嘱信息：主医嘱，还包括其他数据节点
    Jl_Pid_Main := Pljson_List();
    Jl_Pid_Main := j_Input.Get_Pljson_List('advice_list');
    For I In 1 .. Jl_Pid_Main.Count Loop
    
      j_Pid_Main := PLJson();
      j_Pid_Main := PLJson(Jl_Pid_Main.Get(I));
    
      v_药嘱        := Null;
      v_执行时间s_m := Null;
    
      --医嘱相关信息
      n_配置中心id   := j_Pid_Main.Get_Number('pivas_deptid');
      n_相关id_m     := j_Pid_Main.Get_Number('advice_id');
      d_发送号_m     := j_Pid_Main.Get_Number('advice_send_no');
      n_医嘱类型_m   := j_Pid_Main.Get_Number('effective_time');
      n_给药途径id_m := j_Pid_Main.Get_Number('drug_method_id');
      n_是否tpn_m    := j_Pid_Main.Get_Number('is_tpn');
      v_执行频次_m   := j_Pid_Main.Get_String('advice_frequency');
    
      v_主医嘱 := n_配置中心id || ',' || n_相关id_m || ',' || d_发送号_m || ',' || n_医嘱类型_m || ',' || n_给药途径id_m || ',' ||
               n_是否tpn_m || ',' || v_执行频次_m;
    
      --医嘱执行时间分解，先组成字串，后面查询用到
      Jl_Pid_Exetime := Pljson_List();
      Jl_Pid_Exetime := j_Pid_Main.Get_Pljson_List('advice_exetime_list');
      For N In 1 .. Jl_Pid_Exetime.Count Loop
        j_Pid_Exetime := PLJson();
        j_Pid_Exetime := PLJson(Jl_Pid_Exetime.Get(N));
      
        --格式化执行时间串：要求时间,发送号|...
        If v_执行时间s_m Is Null Then
          v_执行时间s_m := j_Pid_Exetime.Get_String('advice_require_time') || ',' ||
                       j_Pid_Exetime.Get_Number('advice_send_no');
        Else
          If Length(v_执行时间s_m || '|' || j_Pid_Exetime.Get_String('advice_require_time') || ',') > 4000 Then
            --超过4K就不要后面的数据，理论上前面的足够了
            Exit;
          Else
            v_执行时间s_m := v_执行时间s_m || '|' || j_Pid_Exetime.Get_String('advice_require_time') || ',' ||
                         j_Pid_Exetime.Get_Number('advice_send_no');
          End If;
        End If;
      End Loop;
    
      --分解药嘱，先产生药嘱串和药嘱对应的发送费用NO
      v_医嘱ids_d := Null;
      Jl_Pid_Drug := Pljson_List();
      Jl_Pid_Drug := j_Pid_Main.Get_Pljson_List('advice_drug_list');
      For M In 1 .. Jl_Pid_Drug.Count Loop
        j_Pid_Drug := PLJson();
        j_Pid_Drug := PLJson(Jl_Pid_Drug.Get(M));
        n_医嘱id_d := j_Pid_Drug.Get_Number('advice_id');
      
        If v_医嘱ids_d Is Null Then
          v_医嘱ids_d := n_医嘱id_d;
        Else
          v_医嘱ids_d := v_医嘱ids_d || ',' || n_医嘱id_d;
        End If;
      
        v_发送no_d := j_Pid_Drug.Get_String('advice_rcpno');
      End Loop;
      v_药嘱 := v_发送no_d || '|' || v_医嘱ids_d;
    
      If v_医嘱信息 Is Null Then
        v_医嘱信息 := v_主医嘱 || ';' || v_药嘱 || ';' || v_执行时间s_m;
      Elsif Length(v_医嘱信息 || '||' || v_主医嘱 || ';' || v_药嘱 || ';' || v_执行时间s_m) > 4000 Then
      
        If c_医嘱信息 Is Null Then
          c_医嘱信息 := v_医嘱信息;
        Else
          c_医嘱信息 := c_医嘱信息 || '||' || v_医嘱信息;
        End If;
        v_医嘱信息 := v_主医嘱 || ';' || v_药嘱 || ';' || v_执行时间s_m;
      
      Else
        v_医嘱信息 := v_医嘱信息 || '||' || v_主医嘱 || ';' || v_药嘱 || ';' || v_执行时间s_m;
      End If;
    End Loop;
  
    If c_医嘱信息 Is Not Null Then
      c_医嘱信息 := c_医嘱信息 || '||' || v_医嘱信息;
      v_医嘱信息 := Null;
    End If;
  
    --主医嘱(静配中心ID,主医嘱ID,发送号,医嘱期效,给药途径id,是否tpn,执行频次);药嘱(费用NO|药嘱id,...);医嘱发送(发送时间,发送号|...)||...
    Zl_输液配药记录_Insert_s(v_核查人, v_核查时间, d_发送时间_m, n_病人id_m, n_主页id_m, v_姓名_m, v_性别_m, v_年龄_m, v_床号_m, v_住院号_m, n_病人病区id_m,
                       n_病人科室id_m, v_医嘱信息, c_医嘱信息);
  End If;

  Json_Out := '{"output":{"code":1,"message": "成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Pivassvr_Newbill;
/


Create Or Replace Procedure Zl_Pivassvr_Locked
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：获取输液配液记录是否被锁定
  --入参：Json_In:格式
  --  input
  --     pivas_ids      C  0   配液id拼串，逗号分割
  --     advice_ids     C  0   医嘱id拼串，逗号分割
  --     send_no        C  0   发送号拼串,逗号分割
  --     advice_info    C  0   格式:医嘱ID1,执行终止时间1;医嘱ID2,执行终止时间2;.......... 执行终止时间为日期格式
  --     query_type     N  1   查询方式，0-只判断是否存，不返配液id；1-判断是否存，并返回配液id ；2-只判断是否存，不返配液id(根据医嘱ID查询)；3-判断是否存，并返回配液id(根据医嘱ID查询)
  --                                     4-按医嘱id和发送号查询配液信息,此种方式时,医嘱id和发送号只传一个值不能有逗号,返回值列表值
  --                                     5-判断是否存已锁定的输液医嘱,传入信息  advice_info 出参判断是否存，并且返回列表 配液id,操作状态，是否打包

  --出参: Json_Out,格式如下

  --  output
  --    code            N 1 应答吗：0-失败；1-成功
  --    message         C 1 应答消息：失败时返回具体的错误信息
  --    isexist         N 1 是否锁定，1-锁定，0-未锁定
  --    pivas_ids       C 1 配液id拼串，逗号分割
  --    item_list
  --       pivas_id     N 1 配液ID
  --       status       N 1 操作状态
  --       is_package   N 0 是否打包
  --       order_id     N 0 医嘱id

  ---------------------------------------------------------------------------

  j_Json       PLJson;
  j_Json_Tmp   PLJson;
  n_Cnt        Number(2) := 0;
  v_配液ids    Varchar2(32767);
  v_医嘱ids    Varchar2(32767);
  v_发送号     Varchar2(32767);
  n_查询方式   Number(1);
  v_配液outids Varchar2(32767);
  v_医嘱信息   Varchar2(32767);
  v_Jtmp       Varchar2(32767);
  c_Jtmp       Clob;

Begin
  --解析入参
  j_Json_Tmp := PLJson(Json_In);
  j_Json     := j_Json_Tmp.Get_Pljson('input');
  n_查询方式 := Nvl(j_Json.Get_Number('query_type'), 0);
  If n_查询方式 > 1 Then
    v_医嘱ids := j_Json.Get_String('advice_ids');
    v_发送号  := j_Json.Get_String('send_no');
  Else
    v_配液ids := j_Json.Get_String('pivas_ids');
  End If;

  If n_查询方式 = 0 Then
    Select Count(1)
    Into n_Cnt
    From 输液配药记录 A
    Where a.Id In (Select /*+cardinality(E,10)*/
                    e.Column_Value
                   From Table(f_Num2List(v_配液ids)) E) And a.是否锁定 = 1;
  
    If n_Cnt > 0 Then
      n_Cnt := 1;
    End If;
    Json_Out := '{"output":{"code":1,"message":"成功","isexist":' || n_Cnt || '}}';
  Elsif n_查询方式 = 1 Then
  
    For R In (Select a.Id
              From 输液配药记录 A
              Where a.Id In (Select /*+cardinality(E,10)*/
                              e.Column_Value
                             From Table(f_Num2List(v_配液ids)) E) And a.是否锁定 = 1) Loop
    
      v_配液outids := v_配液outids || ',' || r.Id;
    
    End Loop;
    v_配液outids := Substr(v_配液outids, 2);
  
    If v_配液outids Is Not Null Then
      n_Cnt := 1;
    Else
      n_Cnt := 0;
    End If;
    Json_Out := '{"output":{"code":1,"message":"成功","isexist":' || n_Cnt || ',"pivas_ids":"' || v_配液outids || '"}}';
  Elsif n_查询方式 = 2 Then
    Select Count(1)
    Into n_Cnt
    From 输液配药记录 A
    Where a.医嘱id In (Select /*+cardinality(E,10)*/
                      e.Column_Value
                     From Table(f_Num2List(v_医嘱ids)) E) And
          (a.发送号 In (Select /*+cardinality(E,10)*/
                      g.Column_Value
                     From Table(f_Num2List(v_发送号)) G) Or Nvl(v_发送号, '空') = '空') And a.是否锁定 = 1;
  
    If n_Cnt > 0 Then
      n_Cnt := 1;
    End If;
    Json_Out := '{"output":{"code":1,"message":"成功","isexist":' || n_Cnt || '}}';
  
  Elsif n_查询方式 = 3 Then
  
    For R In (Select a.Id
              From 输液配药记录 A
              Where a.医嘱id In (Select /*+cardinality(E,10)*/
                                e.Column_Value
                               From Table(f_Num2List(v_医嘱ids)) E) And
                    (a.发送号 In (Select /*+cardinality(E,10)*/
                                g.Column_Value
                               From Table(f_Num2List(v_发送号)) G) Or Nvl(v_发送号, '空') = '空') And a.是否锁定 = 1) Loop
    
      v_配液outids := v_配液outids || ',' || r.Id;
    
    End Loop;
    v_配液outids := Substr(v_配液outids, 2);
  
    If v_配液outids Is Not Null Then
      n_Cnt := 1;
    Else
      n_Cnt := 0;
    End If;
    Json_Out := '{"output":{"code":1,"message":"成功","isexist":' || n_Cnt || ',"pivas_ids":"' || v_配液outids || '"}}';
  Elsif n_查询方式 = 4 Then
    For R In (Select a.Id, a.是否锁定, a.操作状态
              From 输液配药记录 A
              Where a.医嘱id = To_Number(v_医嘱ids) And a.发送号 = To_Number(v_发送号)) Loop
      If Nvl(r.是否锁定, 0) = 1 Then
        n_Cnt := 1;
      End If;
    
      v_Jtmp := v_Jtmp || ',{"pivas_id":' || r.Id;
      v_Jtmp := v_Jtmp || ',"status":' || r.操作状态;
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
  
    If n_Cnt > 0 Then
      n_Cnt := 1;
    End If;
    If c_Jtmp Is Null Then
      Json_Out := '{"output":{"code":1,"message":"成功","isexist":' || n_Cnt || ',"item_list":[' || Substr(v_Jtmp, 2) ||
                  ']}}';
    Else
      c_Jtmp   := c_Jtmp || v_Jtmp;
      Json_Out := '{"output":{"code":1,"message":"成功","isexist":' || n_Cnt || ',"item_list":[' || c_Jtmp || ']}}';
    End If;
  
  Elsif n_查询方式 = 5 Then
    v_医嘱信息 := j_Json.Get_String('advice_info');
    For R In (Select a.医嘱id, a.Id, a.是否锁定, a.操作状态, a.是否打包, a.发送号
              From 输液配药记录 A,
                   (Select /*+cardinality(b,10)*/
                      To_Number(C1) As 医嘱id, To_Date(C2, 'yyyy-mm-dd hh24:mi:ss') As 执行终止时间
                     From Table(Cast(f_Str2List2(v_医嘱信息, ';', ',') As t_StrList2)) B) X
              Where a.医嘱id = x.医嘱id And a.执行时间 > x.执行终止时间) Loop
      If Nvl(r.是否锁定, 0) = 1 Then
        n_Cnt := 1;
      End If;
      --将医嘱医嘱和 配药id一起返回出去方便下一步使用    
      v_Jtmp := v_Jtmp || ',{"pivas_id":' || r.Id;
      v_Jtmp := v_Jtmp || ',"status":' || r.操作状态;
      v_Jtmp := v_Jtmp || ',"is_package":' || Nvl(r.是否打包 || '', 'null');
      v_Jtmp := v_Jtmp || ',"order_id":' || r.医嘱id;
      v_Jtmp := v_Jtmp || ',"send_no":' || r.发送号;
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
    If n_Cnt > 0 Then
      n_Cnt := 1;
    End If;
    If c_Jtmp Is Null Then
      Json_Out := '{"output":{"code":1,"message":"成功","isexist":' || n_Cnt || ',"item_list":[' || Substr(v_Jtmp, 2) ||
                  ']}}';
    Else
      c_Jtmp   := c_Jtmp || v_Jtmp;
      Json_Out := '{"output":{"code":1,"message":"成功","isexist":' || n_Cnt || ',"item_list":[' || c_Jtmp || ']}}';
    End If;
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Pivassvr_Locked;
/

Create Or Replace Procedure Zl_Pivassvr_Getinfo_Batch
(
  Json_In  In Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：获取输液配液记录清单
  --入参：Json_In:格式
  --  input
  --    query_type                          N  1  查询方式   0-按输液配药记录.id查询 by_pivas_id必须传值
  --                                                         1-按医嘱id发送号查询，医嘱id和发送号 查询
  --                                                         2-按配液时间、打包状态、操作状态查询
  --                                                         3-按医嘱ID拼串查询，用于超期收回时医嘱查检已配液的医嘱不能收回 order_ids 需要传值
  --                                                         4-按医嘱id查询判断医嘱是否关联得有输液配药记录
  --                                                         5-按医嘱id查询判断医嘱和状态判断是否在医生站下方显示输液信息页卡
  --                                                         6-按医嘱ID 查询当前这条医嘱的配药记录信息，order_id 表示的是给药途径行的医嘱id
  --                                                                    出参为order_and_no 医嘱+发送号拼串，逗号和冒号分割
  --    by_pivas_id                         N  1   输液配药记录.id 当该结点传值后按单条查询
  --    order_id                            N  1   医嘱id
  --    send_no                             N  1   发送号
  --    begin_time                          C  1   开始时间
  --    end_time                            C  1   结束时间
  --    is_package                          N  1   是否打包
  --    operator_status                     C  1   操作状态，字符串，操作状态拼串
  --    order_ids                           C  1   医嘱ids

  --出参: Json_Out,格式如下

  --  output
  --    code                                N 1 应答吗：0-失败；1-成功
  --    message                             C 1 应答消息：失败时返回具体的错误信息
  --    isexist                             N 1 查询方式为4时是否存在记录
  --    order_and_no                        C 0 医嘱+发送号拼串，逗号和冒号分割
  --    item_list
  --       pivas_id                         N 1 id
  --       order_id                         N 1 医嘱id
  --       dept_id                          N 1 配液科室ID
  --       exe_time                         C 1 执行时间
  --       is_package                       N 1 是否打包
  --       is_locked                        N 1 是否锁定
  --       bottle_label                     C 1 瓶签号
  --       status                           C 1 状态，汉字类型
  --       name                             C 1 病人姓名
  --       inpatientnum                     C 1 住院号
  --       pati_bed                         C 1 床号
  --       pati_deptid                      N 1 病人科室id
  --       send_no                          N 1 发送号
  --       pivas_batchno                    N 1 配药批次
  --       pivas_work_time                  C 1 配药工作时间
  --       package_time                     C 1 打包时间
  --       oper_type                        N 1 操作状态，数字类型
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  n_医嘱id       Number(18);
  n_发送号       Number(18);
  n_配药id       Number(18);
  n_查询方式     Number(3);
  d_开始时间     Date;
  d_结束时间     Date;
  n_是否打包     Number(1);
  v_操作状态     Varchar2(30);
  v_医嘱ids      Varchar2(32767);
  v_Order_And_No Varchar2(32767);
  n_Cnt          Number;
  v_Jtmp         Varchar2(32767);
  c_Jtmp         Clob;
  Cursor c_Pivas Is
    Select b.Id As 配药id, b.发送号, b.医嘱id, To_Char(b.执行时间, 'YYYY-MM-DD HH24:MI') As 执行时间, b.配药批次, g.配药时间 As 配药工作时间, b.瓶签号,
           Decode(b.操作状态, 1, '待摆药', 2, '待配药', 3, '待配药', 4, '已配药', '已发送') As 状态,
           To_Char(b.打包时间, 'YYYY-MM-DD HH24:MI') As 打包时间
    From 输液配药记录 B, 配药工作批次 G
    Where b.配药批次 = g.批次(+) And b.部门id = g.配置中心id(+) And b.执行时间 Between d_开始时间 And d_结束时间 And Nvl(b.是否打包, 0) = n_是否打包 And
          (Instr(',' || v_操作状态 || ',', ',' || b.操作状态 || ',') > 0 Or v_操作状态 Is Null);

Begin
  --解析入参
  j_Jsonin := PLJson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');

  n_查询方式 := j_Json.Get_Number('query_type');
  n_查询方式 := Nvl(n_查询方式, 0);
  If n_查询方式 = 0 Then
    --只会有一条数据
    n_配药id := j_Json.Get_Number('by_pivas_id');
    For R In (Select a.Id, a.医嘱id, a.部门id As 配液科室id, To_Char(a.执行时间, 'YYYY-MM-DD HH24:MI') As 执行时间, a.是否打包, a.配药批次, a.瓶签号,
                     Decode(a.操作状态, 1, '待摆药', 2, '待配药', 3, '待配药', 4, '已配药', 5, '已发送', 6, '已签收', 7, '已拒绝签收', 8, '已确认拒收', 9,
                             '已销帐申请', 10, '已销帐审核', 11, '销帐拒绝', '已发送') As 状态, a.姓名, a.住院号, a.床号, a.病人科室id, a.是否锁定, a.操作状态
              From 输液配药记录 A
              Where a.Id = n_配药id) Loop
    
      v_Jtmp := v_Jtmp || '{"pivas_id":' || r.Id;
      v_Jtmp := v_Jtmp || ',"order_id":' || r.医嘱id;
      v_Jtmp := v_Jtmp || ',"dept_id":' || r.配液科室id;
      v_Jtmp := v_Jtmp || ',"exe_time":"' || r.执行时间 || '"';
      v_Jtmp := v_Jtmp || ',"is_package":' || Nvl(r.是否打包, 0);
      v_Jtmp := v_Jtmp || ',"is_locked":' || Nvl(r.是否锁定, 0);
      v_Jtmp := v_Jtmp || ',"pivas_batchno":' || Nvl(r.配药批次 || '', 'null');
      v_Jtmp := v_Jtmp || ',"bottle_label":"' || zlJsonStr(r.瓶签号) || '"';
      v_Jtmp := v_Jtmp || ',"status":"' || r.状态 || '"';
      v_Jtmp := v_Jtmp || ',"oper_type":' || Nvl(r.操作状态, 0);
      v_Jtmp := v_Jtmp || ',"name":"' || zlJsonStr(r.姓名) || '"';
      v_Jtmp := v_Jtmp || ',"inpatientnum":"' || r.住院号 || '"';
      v_Jtmp := v_Jtmp || ',"pati_bed":"' || zlJsonStr(r.床号) || '"';
      v_Jtmp := v_Jtmp || ',"pati_deptid":' || Nvl(r.病人科室id || '', 'null');
      v_Jtmp := v_Jtmp || '}';
    
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || v_Jtmp || ']}}';
  
  Elsif n_查询方式 = 1 Then
    n_医嘱id := j_Json.Get_Number('order_id');
    n_发送号 := j_Json.Get_Number('send_no');
    For R In (Select a.Id, a.医嘱id, a.部门id As 配液科室id, To_Char(a.执行时间, 'YYYY-MM-DD HH24:MI') As 执行时间, a.是否打包, a.配药批次, a.瓶签号,
                     Decode(操作状态, 1, '待摆药', 2, '待配药', 3, '待配药', 4, '已配药', 5, '已发送', 6, '已签收', 7, '已拒绝签收', 8, '已确认拒收', 9,
                             '已销帐申请', 10, '已销帐审核', 11, '销帐拒绝', '已发送') As 状态, a.姓名, a.住院号, a.床号, a.病人科室id, a.是否锁定, a.操作状态
              From 输液配药记录 A
              Where a.医嘱id = n_医嘱id And (a.发送号 = n_发送号 Or n_发送号 = 0) And a.操作状态 <> 8
              Order By a.执行时间) Loop
    
      v_Jtmp := v_Jtmp || ',{"pivas_id":' || r.Id;
      v_Jtmp := v_Jtmp || ',"order_id":' || r.医嘱id;
      v_Jtmp := v_Jtmp || ',"dept_id":' || r.配液科室id;
      v_Jtmp := v_Jtmp || ',"exe_time":"' || r.执行时间 || '"';
      v_Jtmp := v_Jtmp || ',"is_package":' || Nvl(r.是否打包, 0);
      v_Jtmp := v_Jtmp || ',"is_locked":' || Nvl(r.是否锁定, 0);
      v_Jtmp := v_Jtmp || ',"pivas_batchno":' || Nvl(r.配药批次 || '', 'null');
      v_Jtmp := v_Jtmp || ',"bottle_label":"' || zlJsonStr(r.瓶签号) || '"';
      v_Jtmp := v_Jtmp || ',"status":"' || r.状态 || '"';
      v_Jtmp := v_Jtmp || ',"oper_type":' || Nvl(r.操作状态, 0);
      v_Jtmp := v_Jtmp || ',"name":"' || zlJsonStr(r.姓名) || '"';
      v_Jtmp := v_Jtmp || ',"inpatientnum":"' || r.住院号 || '"';
      v_Jtmp := v_Jtmp || ',"pati_bed":"' || zlJsonStr(r.床号) || '"';
      v_Jtmp := v_Jtmp || ',"pati_deptid":' || Nvl(r.病人科室id || '', 'null');
      v_Jtmp := v_Jtmp || '}';
    
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Elsif n_查询方式 = 2 Then
  
    d_开始时间 := To_Date(j_Json.Get_String('begin_time'), 'YYYY-MM-DD HH24:MI:SS');
    d_结束时间 := To_Date(j_Json.Get_String('end_time'), 'YYYY-MM-DD HH24:MI:SS');
    n_是否打包 := j_Json.Get_Number('is_package');
    v_操作状态 := j_Json.Get_String('operator_status');
  
    For R In c_Pivas Loop
    
      v_Jtmp := v_Jtmp || '{"pivas_id":' || r.配药id;
      v_Jtmp := v_Jtmp || ',"order_id":' || r.医嘱id;
      v_Jtmp := v_Jtmp || ',"send_no":' || r.发送号;
      v_Jtmp := v_Jtmp || ',"exe_time":"' || r.执行时间 || '"';
      v_Jtmp := v_Jtmp || ',"pivas_work_time":"' || r.配药工作时间 || '"';
      v_Jtmp := v_Jtmp || ',"pivas_batchno":' || Nvl(r.配药批次 || '', 'null');
      v_Jtmp := v_Jtmp || ',"bottle_label":"' || zlJsonStr(r.瓶签号) || '"';
      v_Jtmp := v_Jtmp || ',"status":"' || r.状态 || '"';
      v_Jtmp := v_Jtmp || ',"package_time":"' || r.打包时间 || '"';
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
  Elsif n_查询方式 = 3 Then
    v_医嘱ids := j_Json.Get_String('order_ids');
    For R In (Select b.医嘱id, To_Char(Max(b.执行时间), 'YYYY-MM-DD HH24:MI:SS') As 执行时间
              From 输液配药记录 B
              Where (b.操作状态 In (4, 5, 6, 7, 8) And Nvl(b.是否打包, 0) = 0) And
                    b.医嘱id In (Select /*+cardinality(x,10)*/
                                x.Column_Value
                               From Table(Cast(f_Num2List(v_医嘱ids) As Zltools.t_Numlist)) X)
              Group By b.医嘱id) Loop
    
      v_Jtmp := v_Jtmp || ',{"order_id":' || r.医嘱id;
      v_Jtmp := v_Jtmp || ',"exe_time":"' || r.执行时间 || '"';
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
  
  Elsif n_查询方式 = 4 Then
    n_医嘱id := j_Json.Get_Number('order_id');
    Select Count(1) Into n_Cnt From 输液配药记录 A Where a.医嘱id = n_医嘱id;
    If n_Cnt > 0 Then
      n_Cnt := 1;
    End If;
    Json_Out := '{"output":{"code":1,"message":"成功","isexist":' || n_Cnt || '}}';
  Elsif n_查询方式 = 5 Then
    n_医嘱id   := j_Json.Get_Number('order_id');
    v_操作状态 := j_Json.Get_String('operator_status');
    Select Count(1)
    Into n_Cnt
    From 输液配药记录 A
    Where a.医嘱id = n_医嘱id And Instr(',' || v_操作状态 || ',', ',' || a.操作状态 || ',') = 0;
    If n_Cnt > 0 Then
      n_Cnt := 1;
    End If;
    Json_Out := '{"output":{"code":1,"message":"成功","isexist":' || n_Cnt || '}}';
  Elsif n_查询方式 = 6 Then
    n_医嘱id := j_Json.Get_Number('order_id');
    For R In (Select a.医嘱id || ':' || a.发送号 As 值 From 输液配药记录 A Where a.医嘱id = n_医嘱id) Loop
      v_Order_And_No := v_Order_And_No || ',' || r.值;
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"成功","order_and_no":"' || Substr(v_Order_And_No, 2) || '"}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Pivassvr_Getinfo_Batch;
/

Create Or Replace Procedure Zl_Pivassvr_Pivasupdate
(
  Json_In  Clob,
  Json_Out Out Clob
) Is

  --------------------------------------------------------------------------------------------------------------------
  --功能：输液配药记录修改
  --入参：Json_In,格式如下
  --  input
  --    item_list
  --         pivas_id                   N   1   配药id
  --         is_package                 N   1   是否打包
  --         batch_no                   N   1   配药批次
  --         checker                    C   1   核查人
  --         check_time                 D   1   核查时间
  --出参: Json_Out,格式如下
  --  output
  --    code                            N 1 应答吗：0-失败；1-成功
  --    message                         C 1 应答消息：失败时返回具体的错误信息
  --------------------------------------------------------------------------------------------------------------------

  n_配药id   Number(18);
  n_是否打包 输液配药记录.是否打包%Type;
  n_配药批次 输液配药记录.配药批次%Type;
  v_核查人   输液配药状态.操作人员%Type;
  d_核查时间 输液配药状态.操作时间%Type;

  j_Json  Pljson;
  Jl_Item Pljson_List;
  j_Item  Pljson;

  n_Ct    Number;
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --解析入参
  j_Item  := Pljson(Json_In);
  j_Json  := j_Item.Get_Pljson('input');
  Jl_Item := j_Json.Get_Pljson_List('item_list');
  For I In 1 .. Jl_Item.Count Loop
    j_Item     := Pljson();
    j_Item     := Pljson(Jl_Item.Get(I));
    n_配药id   := j_Item.Get_Number('pivas_id');
    n_是否打包 := j_Item.Get_Number('is_package');
    n_配药批次 := j_Item.Get_Number('batch_no');
    v_核查人   := j_Item.Get_String('checker');
    d_核查时间 := To_Date(j_Item.Get_String('check_time'), 'YYYY-MM-DD HH24:MI:SS');
    Select Count(1)
    Into n_Ct
    From 输液配药记录
    Where ID = n_配药id And 操作状态 In (1, 2, 3) And Nvl(配药批次, 0) <> Nvl(n_配药批次, 0);
    Update 输液配药记录
    Set 是否打包 = n_是否打包, 配药批次 = n_配药批次, 打包时间 = Decode(n_是否打包, 1, d_核查时间, Null), 批次标记 = Decode(n_Ct, 1, 2, 批次标记)
    Where ID = n_配药id And 操作状态 In (1, 2, 3);
    If Sql%RowCount = 0 Then
      v_Error := '由于并发操作,当前修改的配药记录已配药,操作失败.';
      Raise Err_Custom;
    End If;
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Pivassvr_Pivasupdate;
/

Create Or Replace Procedure Zl_Pivassvr_Getstatus_Info
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：获取输液配液记录状态信息
  --入参：Json_In:格式
  --  input
  --     pivas_ids    C  1   配液ids，逗号拼串
  --出参: Json_Out,格式如下
  --  output
  --    code            N 1 应答吗：0-失败；1-成功
  --    message         C 1 应答消息：失败时返回具体的错误信息
  --    item_list
  --       pivas_id      N 1 配药ID
  --       oper_type     N 1 操作类型
  --       operator_name C 1 操作员
  --       operator_time C 1 操作时间
  --       operator_notes C 1 操作说明
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  v_配液ids Varchar2(32767);
  v_Jtmp    Varchar2(32767);
  c_Jtmp    Clob;

  v_Vals Clob;
  l_Vals t_StrList;
Begin
  --解析入参
  j_Jsonin := PLJson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  v_Vals   := j_Json.Get_Clob('pivas_ids');

  l_Vals := t_StrList();
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

  For Lp In 1 .. l_Vals.Count Loop
    v_配液ids := l_Vals(Lp);
    For R In (Select a.配药id, a.操作类型, a.操作人员, To_Char(a.操作时间, 'YYYY-MM-DD HH24:MI:SS') As 操作时间, a.操作说明
              From 输液配药状态 A
              Where a.配药id In (Select /*+cardinality(x,10)*/
                                x.Column_Value
                               From Table(f_Num2List(v_配液ids)) X)) Loop
    
      v_Jtmp := v_Jtmp || ',{"pivas_id":' || r.配药id;
      v_Jtmp := v_Jtmp || ',"oper_type":' || r.操作类型;
      v_Jtmp := v_Jtmp || ',"operator_name":"' || zlJsonStr(r.操作人员) || '"';
      v_Jtmp := v_Jtmp || ',"operator_time":"' || r.操作时间 || '"';
      v_Jtmp := v_Jtmp || ',"operator_notes":"' || zlJsonStr(r.操作说明) || '"';
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
End Zl_Pivassvr_Getstatus_Info;
/

CREATE OR REPLACE Procedure Zl_Pivassvr_Pivascanstop
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：检查输液配液记录是否允许停止
  --入参：Json_In:格式
  --  input
  --     order_id    N  1   配液id
  --     stop_time   C  1   停嘱时间

  --出参: Json_Out,格式如下

  --  output
  --    code            N 1 应答吗：0-失败；1-成功
  --    message         C 1 应答消息：失败时返回具体的错误信息
  --    can_stop_time   C 1 允许停止的时间
  --    tip_time        C 1 实际执行时间
  --    is_package      N 1 是否打包

  ---------------------------------------------------------------------------

  j_Json         PLJson;
  j_Json_Tmp     PLJson;
  j_Item         PLJson;
  n_医嘱id       Number(18);
  n_打包         Number(2);
  d_Stoptime     Date;
  v_允许时间     Varchar2(30);
  v_提示执行时间 Varchar2(30);

Begin
  --解析入参
  j_Item     := PLJson(Json_In);
  j_Json     := j_Item.Get_Pljson('input');
  n_医嘱id   := j_Json.Get_Number('order_id');
  d_Stoptime := To_Date(j_Json.Get_String('stop_time'), 'YYYY-MM-DD HH24:MI:SS');

  For R In (Select Min(Decode(Instr(',4,5,6,7,8,', ',' || b.操作类型 || ','), 0, Null, To_Char(a.执行时间, 'yyyy-MM-dd HH24:MI'))) As 允许执行时间,
                   Min(Decode(a.操作状态, 1, Null, To_Char(a.执行时间, 'yyyy-MM-dd HH24:MI'))) As 提示执行时间, Min(a.是否打包) As 打包
            From 输液配药记录 A, 输液配药状态 B
            Where a.医嘱id = n_医嘱id And a.Id = b.配药id And a.执行时间 > d_Stoptime And a.操作状态 <> 10) Loop
    v_允许时间     := r.允许执行时间;
    v_提示执行时间 := r.提示执行时间;
    n_打包         := r.打包;
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"成功"';
  Json_Out := Json_Out || ',"can_stop_time":"' || v_允许时间 || '"';
  Json_Out := Json_Out || ',"tip_time":"' || v_提示执行时间 || '"';
  Json_Out := Json_Out || ',"is_package":' || Nvl(n_打包, 0);
  Json_Out := Json_Out || '}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Pivassvr_Pivascanstop;
/

Create Or Replace Procedure Zl_Pivassvr_Isselfdrug
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：获取输液自备药清单
  --入参：Json_In:格式
  --  input
  --    drug_ids    C   1   药品ID，多个用英文的逗号分隔
  --出参: Json_Out,格式如下
  --  output
  --    code          N 1 应答吗：0-失败；1-成功
  --    message       C 1 应答消息：失败时返回具体的错误信息
  --    drug_ids      C 1 药品ID，多个用英文的逗号分隔
  ---------------------------------------------------------------------------
  j_Jsonin PLJson;
  j_Json   PLJson;

  v_Drugids Varchar2(32767);

  v_Tmp Varchar2(32767);
Begin
  --解析入参
  j_Jsonin  := PLJson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  v_Drugids := j_Json.Get_String('drug_ids');

  For R In (Select a.药品id
            From 输液自备药清单 A
            Where a.药品id In (Select /*+cardinality(x,10)*/
                              x.Column_Value
                             From Table(Cast(f_Num2List(v_Drugids) As Zltools.t_Numlist)) X)) Loop
    v_Tmp := v_Tmp || ',' || r.药品id;
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"成功","drug_ids":"' || Substr(v_Tmp, 2) || '"}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Pivassvr_Isselfdrug;
/

Create Or Replace Procedure Zl_Pivassvr_Updatebatch
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  --------------------------------------------------------------------------------------------------------------------
  --功能：输液医嘱批次调整
  --入参：Json_In,格式如下
  --  input
  --    item_list
  --         order_id                   N   1   医嘱id,药品医嘱的主医嘱id
  --         operator_time              C   1   操作时间，病人医嘱状态中操作类型=8的记录里面的最大时间
  --出参: Json_Out,格式如下
  --  output
  --    code                            N 1 应答吗：0-失败；1-成功
  --    message                         C 1 应答消息：失败时返回具体的错误信息
  --------------------------------------------------------------------------------------------------------------------
  n_Count    Number(8);
  n_医嘱id   Varchar2(4000);
  d_操作时间 Date;
  v_批次     Varchar2(20);

  j_Json     Pljson; 
  Jl_Item    Pljson_List;
  j_Item     Pljson;

  Cursor c_输液记录 Is
    Select Distinct a.Id 配药id, a.执行时间 时间, a.部门id
    From 输液配药记录 A
    Where a.医嘱id = n_医嘱id And a.操作状态 = 1 And Trunc(d_操作时间) = Trunc(a.执行时间) And a.执行时间 < d_操作时间;

  v_输液记录 c_输液记录%RowType;

  Function Zl_Getpivaworkbatch
  (
    执行时间_In   In Date,
    配置中心id_In In 输液配药记录.部门id%Type
  ) Return Number As
    v_Exetime   Date;
    v_Starttime Date;
    v_Endtime   Date;
    v_Maxbatch  Number(2);
    v_Batch     Number;
    Cursor c_配药批次 Is
      Select 批次, 配药时间, 给药时间 From 配药工作批次 Where 启用 = 1 And 配置中心id = 配置中心id_In Order By 批次;
  
    v_配药批次 c_配药批次%RowType;
  Begin
    v_Exetime := To_Date(Substr(To_Char(执行时间_In, 'yyyy-mm-dd hh24:mi'), 12), 'hh24:mi');
  
    Select Max(Nvl(批次, 0)) + 1 Into v_Maxbatch From 配药工作批次 Where 启用 = 1 And 配置中心id = 配置中心id_In;
  
    For v_配药批次 In c_配药批次 Loop
      v_Batch     := 0;
      v_Starttime := To_Date(Substr(v_配药批次.给药时间, 1, Instr(v_配药批次.给药时间, '-') - 1), 'hh24:mi');
      v_Endtime   := To_Date(Substr(v_配药批次.给药时间, Instr(v_配药批次.给药时间, '-') + 1), 'hh24:mi');
    
      If v_Exetime >= v_Starttime And v_Exetime <= v_Endtime Then
        v_Batch := v_配药批次.批次;
      
        Exit When v_Batch > 0;
      End If;
    End Loop;
  
    If v_Batch = 0 Then
      v_Batch := v_Maxbatch;
    End If;
    Return(v_Batch);
  End;
Begin
  --解析入参
  j_Item := Pljson(Json_In);
  j_Json := j_Item.Get_Pljson('input');
  Jl_Item := j_Json.Get_Pljson_List('item_list');
  
  For I In 1 .. Jl_Item.Count Loop
    j_Item := Pljson();
    j_Item := Pljson(Jl_Item.Get(I));
  
    n_医嘱id   := j_Item.Get_Number('order_id');
    d_操作时间 := To_Date(j_Item.Get_String('operator_time'), 'YYYY-MM-DD HH24:MI:SS');
  
    Select Count(a.Id)
    Into n_Count
    From 输液配药记录 A
    Where a.医嘱id = n_医嘱id And a.操作状态 = 1 And a.执行时间 > d_操作时间;
  
    If n_Count > 0 Then
      For v_输液记录 In c_输液记录 Loop
        v_批次 := Zl_Getpivaworkbatch(v_输液记录.时间, v_输液记录.部门id);
        Update 输液配药记录 Set 配药批次 = v_批次, 是否调整批次 = 1 Where ID = v_输液记录.配药id;
      End Loop;
    End If;
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Pivassvr_Updatebatch;
/

Create Or Replace Procedure Zl_Pivassvr_Getpivascontent
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  -----------------------------------------------------------------------------
  --功能:获取输液配药内容
  --入参：Json_In:格式
  --  input
  --    pivas_id                    N      1   配液id
  --    order_id                    N      0   医嘱ID
  --    advice_endtime              C      0   医嘱的执行终止时间
  --    auto_aduit                  N      0   是否可以自动审核，0-不能自动审核，1-可以自动审核，主要用于区分是否已经发药
  --    query_type                  N      1   查询方式
  --                                           0-按照pivas_id单个，进行查询,应用场景护士工作站对输液医嘱进行销帐
  --                                           1-按照pivas_id+order_id+advice_endtime进行查询获取配药明细信息
  --                                           2-取消销帐申请时获取费用id明细
  --出参: Json_Out,格式如下
  --  output
  --    code                        N      1   应答吗：0-失败；1-成功
  --    message                     C      1   应答消息：失败时返回具体的错误信息
  --    fee_ids                     C      0   费用明细id，query_type=2时有此结点
  --    oper_time                   C      0   操作时间，销帐申请的时间query_type=2时有此结点
  --    item_list[]配液记录明细
  --       pivas_id                 N      1   配液id
  --       rcp_no                   C      1   单据号
  --       rcpdtl_id                N      1   药品收发记录明细id,处方明细id/费用id
  --       send_num                 N      1   发药品数量
  --       drug_id                  N      1   药品id
  --       si_drug_packg_qunt       N      1   住院包装,包装数量
  --       si_drug_packg_unit       C      1   包装单位
  --       drug_name                C      1   药品名称
  --       is_sended                N      1   是否发药
  --       drugstore_id             N      1   输液配药记录中的部门，返回后当做销帐时的审核部门
  --       status                   N      1   操作状态  （输液配药记录）
  --       order_id                 N      0   医嘱id
  --       send_no                  N      0   医嘱发送号
  -----------------------------------------------------------------------------
  n_配液id     Number(18);
  n_医嘱id     Number(18);
  d_终止时间   Date;
  n_Query_Type Number;
  v_Fee_Ids    Varchar2(32767);
  j_Json       Pljson;
  j_Json_Tmp   Pljson;
  n_Auto_Aduit Number;
  v_Jtmp       Varchar2(32767);
  c_Jtmp       Clob;

  Cursor c_Pivas_Ex Is
    Select b.费用id, b.药品id As 收费细目id, Sum(a.数量) As 数量, c.住院包装, c.住院单位, d.名称, b.No, e.医嘱id, e.发送号
    From 输液配药内容 A, 药品收发记录 B, 药品规格 C, 收费项目目录 D, 输液配药记录 E
    Where a.记录id = n_配液id And a.收发id = b.Id And b.药品id = c.药品id And c.药品id = d.Id And a.记录id = e.Id And
          Instr(',8,9,10,21,24,25,26,', ',' || b.单据 || ',') > 0 And
          (n_Auto_Aduit = 1 And b.审核人 Is Null Or n_Auto_Aduit = 0 And b.审核人 Is Not Null)
    Group By b.费用id, b.药品id, c.住院包装, c.住院单位, d.名称, b.No, e.医嘱id, e.发送号;

  Cursor c_Pivas2
  (
    相关id_In       Number,
    执行终止时间_In Date,
    配药id_In       Number
  ) Is
    Select a.记录id, a.收发id, Sum(a.数量) As 数量, b.费用id, b.药品id, 0 As 是否发药, c.住院包装, c.住院单位, d.名称, b.No, e.病人病区id As 部门id,
           e.操作状态
    From 输液配药内容 A, 药品收发记录 B, 药品规格 C, 收费项目目录 D, 输液配药记录 E
    Where a.收发id = b.Id And b.药品id = c.药品id And c.药品id = d.Id And e.Id = a.记录id And e.医嘱id = 相关id_In And
          e.执行时间 > 执行终止时间_In And e.Id = 配药id_In
    Group By b.费用id, b.药品id, c.住院包装, c.住院单位, d.名称, e.病人病区id, e.操作状态, e.Id, a.收发id, a.记录id, e.病人病区id, d.名称, b.No, e.操作状态;

  Type t_Pivas1 Is Table Of c_Pivas2%RowType;
  r_p t_Pivas1; --游标对其赋值后，可以修改数组中的值

Begin
  --解析入参
  j_Json_Tmp   := Pljson(Json_In);
  j_Json       := j_Json_Tmp.Get_Pljson('input');
  n_Query_Type := j_Json.Get_Number('query_type');
  n_配液id     := j_Json.Get_Number('pivas_id');
  n_Auto_Aduit := j_Json.Get_Number('auto_aduit');
  n_Auto_Aduit := Nvl(n_Auto_Aduit, 0);

  If Nvl(n_Query_Type, 0) = 0 Then
    For R In c_Pivas_Ex Loop
      v_Jtmp := v_Jtmp || ',{"rcp_no":"' || r.No || '"';
      v_Jtmp := v_Jtmp || ',"rcpdtl_id":' || r.费用id;
      v_Jtmp := v_Jtmp || ',"send_num":' || Zljsonstr(r.数量, 1);
      v_Jtmp := v_Jtmp || ',"drug_id":' || r.收费细目id;
      v_Jtmp := v_Jtmp || ',"si_drug_packg_qunt":' || Zljsonstr(r.住院包装, 1);
      v_Jtmp := v_Jtmp || ',"si_drug_packg_unit":"' || Zljsonstr(r.住院单位) || '"';
      v_Jtmp := v_Jtmp || ',"drug_name":"' || Zljsonstr(r.名称) || '"';
      v_Jtmp := v_Jtmp || ',"order_id":' || r.医嘱id;
      v_Jtmp := v_Jtmp || ',"send_no":' || r.发送号;
      v_Jtmp := v_Jtmp || '}';
      If Length(v_Jtmp) > 30000 Then
        If c_Jtmp Is Null Then
          c_Jtmp := Substr(v_Jtmp, 2);
        Else
          c_Jtmp := c_Jtmp || v_Jtmp;
        End If;
        v_Jtmp := Null;
      End If;
      If Instr(',' || v_Fee_Ids || ',', ',' || r.费用id || ',') = 0 Then
        v_Fee_Ids := v_Fee_Ids || ',' || r.费用id;
      End If;
    End Loop;
    v_Fee_Ids := Substr(v_Fee_Ids, 2);
  
    If c_Jtmp Is Null Then
      Json_Out := '{"output":{"code":1,"message":"成功","fee_ids":"' || v_Fee_Ids || '","item_list":[' ||
                  Substr(v_Jtmp, 2) || ']}}';
    Else
      c_Jtmp   := c_Jtmp || v_Jtmp;
      Json_Out := '{"output":{"code":1,"message":"成功","fee_ids":"' || v_Fee_Ids || '","item_list":[' || c_Jtmp || ']}}';
    End If;
  
  Elsif n_Query_Type = 1 Then
    n_医嘱id   := j_Json.Get_Number('order_id');
    d_终止时间 := To_Date(j_Json.Get_String('advice_endtime'), 'yyyy-mm-dd hh24:mi:ss');
    Open c_Pivas2(n_医嘱id, d_终止时间, n_配液id);
    Fetch c_Pivas2 Bulk Collect
      Into r_p;
    Close c_Pivas2;
  
    For I In 1 .. r_p.Count Loop
      v_Jtmp := v_Jtmp || ',{"pivas_id":' || r_p(I).记录id;
      v_Jtmp := v_Jtmp || ',"rcp_no":"' || r_p(I).No || '"';
      v_Jtmp := v_Jtmp || ',"rcpdtl_id":' || r_p(I).费用id;
      v_Jtmp := v_Jtmp || ',"send_num":' || Zljsonstr(r_p(I).数量, 1);
      v_Jtmp := v_Jtmp || ',"drug_id":' || r_p(I).药品id;
      v_Jtmp := v_Jtmp || ',"si_drug_packg_qunt":' || Zljsonstr(r_p(I).住院包装, 1);
      v_Jtmp := v_Jtmp || ',"si_drug_packg_unit":"' || Zljsonstr(r_p(I).住院单位) || '"';
      v_Jtmp := v_Jtmp || ',"drug_name":"' || Zljsonstr(r_p(I).名称) || '"';
      v_Jtmp := v_Jtmp || ',"is_sended":' || Nvl(r_p(I).是否发药, 0);
      v_Jtmp := v_Jtmp || ',"drugstore_id":' || Nvl(r_p(I).部门id || '', 'null');
      v_Jtmp := v_Jtmp || ',"status":' || Nvl(r_p(I).操作状态 || '', 'null');
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
  Elsif n_Query_Type = 2 Then
    v_Fee_Ids  := Null;
    d_终止时间 := Null;
    For R In (Select Distinct b.费用id From 输液配药内容 A, 药品收发记录 B Where a.收发id = b.Id And a.记录id = n_配液id) Loop
      v_Fee_Ids := v_Fee_Ids || ',' || r.费用id;
    End Loop;
    --返回下操作时间，按申请时来删除申请记录更准确些
    For R In (Select a.操作时间
              From (Select 操作时间
                     From 输液配药状态
                     Where 配药id = n_配液id And 操作类型 = 9
                     Order By 操作时间 Desc, 操作类型 Desc) A) Loop
      d_终止时间 := r.操作时间;
      Exit;
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"成功","fee_ids":"' || Substr(v_Fee_Ids, 2) || '","oper_time":"' ||
                To_Char(d_终止时间, 'yyyy-mm-dd hh24:mi:ss') || '"}}';
  End If;
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Pivassvr_Getpivascontent;
/

Create Or Replace Procedure Zl_Pivassvr_Statusupdate
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------------
  --功能：配液中心数据状态处理
  --入参
  -- input
  --   pivas_ids                               C   1   配药记录id逗号拼串方式,些结点不传则按列表值进行更新
  --   operator_status                         N   1   操作状态, 当操作状态为 -1 时表示为取消销帐申请
  --   operator_name                           C   1   操作员姓名
  --   operator_notes                          C   0   操作说明
  --   operator_time                           C   1   操作时间
  --   auto_aduit                              N   0   自动审核销帐申请，自动审核状态为10
  --   item_list[]此列表结点可以不传如果传入以列表方式更新
  --      pivas_id                                N   1   配药记录id
  --      operator_status                         N   1   操作状态, 当操作状态为 -1 时表示为取消销帐申请
  --      operator_name                           C   1   操作员姓名
  --      operator_notes                          C   0   操作说明
  --      operator_time                           C   1   操作时间
  --出参
  -- output
  --   code                                    N   1   应答吗：0-失败；1-成功
  --   message                                 C   1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------------

  j_Json       Pljson;
  j_List       Pljson_List;
  j_Tmp        Pljson;
  v_配液ids    Varchar2(32767);
  v_操作员姓名 Varchar2(40);
  d_操作时间   Date;
  d_审核时间   Date;
  n_操作状态   Number(2);
  v_操作说明   Varchar2(4000);
  n_配液id     Number(18);
  n_操作类型   Number(2);
  n_Auto_Aduit Number;
  v_Vals       Clob;
  l_Vals       t_Strlist;
Begin
  --解析入参
  j_Tmp  := Pljson(Json_In);
  j_Json := j_Tmp.Get_Pljson('input');
  v_Vals := j_Json.Get_Clob('pivas_ids');
  If v_Vals Is Not Null Then
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
  
    n_操作状态   := j_Json.Get_Number('operator_status');
    v_操作员姓名 := j_Json.Get_String('operator_name');
    d_操作时间   := To_Date(j_Json.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');
    v_操作说明   := j_Json.Get_String('operator_notes');
    n_Auto_Aduit := j_Json.Get_Number('auto_aduit');
  
    If Nvl(n_操作状态, 0) <> -1 Then
      If n_Auto_Aduit = 1 Then
        d_操作时间 := d_操作时间 + 1 / 24 / 60 / 60;
      End If;
      For Lp In 1 .. l_Vals.Count Loop
        v_配液ids := l_Vals(Lp);
        Insert Into 输液配药状态
          (配药id, 操作类型, 操作人员, 操作时间, 操作说明)
          Select /*+cardinality(y,10)*/
           To_Number(y.Column_Value) As 配药id, n_操作状态, v_操作员姓名, d_操作时间, v_操作说明
          From Table(f_Num2list(v_配液ids)) Y;
      
        Update 输液配药记录
        Set 操作人员 = v_操作员姓名, 操作时间 = d_操作时间, 操作状态 = n_操作状态
        Where ID In (Select /*+cardinality(y,10)*/
                      y.Column_Value
                     From Table(f_Num2list(v_配液ids)) Y);
      
        If n_Auto_Aduit = 1 Then
          --插入销帐申请审核状态，时间区分开来加一秒
          Insert Into 输液配药状态
            (配药id, 操作类型, 操作人员, 操作时间, 操作说明)
            Select /*+cardinality(y,10)*/
             To_Number(y.Column_Value) As 配药id, 10, v_操作员姓名, d_审核时间, v_操作说明
            From Table(f_Num2list(v_配液ids)) Y;
        
          Update 输液配药记录
          Set 操作人员 = v_操作员姓名, 操作时间 = d_审核时间, 操作状态 = 10
          Where ID In (Select /*+cardinality(y,10)*/
                        y.Column_Value
                       From Table(f_Num2list(v_配液ids)) Y);
        
        End If;
      End Loop;
    Else
      --特别注意，取消销帐只能是一个一个的取消
      v_配液ids := l_Vals(1);
      n_配液id  := To_Number(v_配液ids);
      Select 操作人员, 操作时间, 操作类型
      Into v_操作员姓名, d_操作时间, n_操作类型
      From (Select 操作人员, 操作时间, 操作类型
             From 输液配药状态
             Where 配药id = n_配液id And 操作类型 <> 9
             Order By 操作时间 Desc, 操作类型 Desc)
      Where Rownum = 1;
    
      Update 输液配药记录
      Set 操作人员 = v_操作员姓名, 操作时间 = d_操作时间, 操作状态 = n_操作类型
      Where ID = n_配液id;
    
    End If;
  Else
    j_List := j_Json.Get_Pljson_List('item_list');
    For I In 1 .. j_List.Count Loop
      j_Tmp        := Pljson();
      j_Tmp        := Pljson(j_List.Get(I));
      n_配液id     := j_Tmp.Get_Number('pivas_id');
      n_操作状态   := j_Tmp.Get_Number('operator_status');
      v_操作员姓名 := j_Tmp.Get_String('operator_name');
      d_操作时间   := To_Date(j_Tmp.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');
      v_操作说明   := j_Tmp.Get_String('operator_notes');
    
      Insert Into 输液配药状态
        (配药id, 操作类型, 操作人员, 操作时间, 操作说明)
      Values
        (n_配液id, n_操作状态, v_操作员姓名, d_操作时间, v_操作说明);
    
      Update 输液配药记录
      Set 操作人员 = v_操作员姓名, 操作时间 = d_操作时间, 操作状态 = n_操作状态
      Where ID = n_配液id;
    
    End Loop;
  End If;
  Json_Out := '{"output":{"code":1,"message": "成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Pivassvr_Statusupdate;
/

Create Or Replace Procedure Zl_Pivassvr_Patiinfoupdate
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：住院病人输液配药记录基本信息修改
  --入参：Json_In:格式
  --  input
  --   pati_id       N   1   病人id
  --   pati_name      C   1   患者姓名
  --   pati_sex       C   1   患者性别
  --   pati_age       C   1   患者年龄
  --   visit_id      N   1   就诊id

  --出参: Json_Out,格式如下
  --  output
  --    code                 N   1 应答吗：0-失败；1-成功
  --    message              C   1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------

  j_Json Pljson;

  v_姓名     Varchar2(100);
  v_性别     Varchar2(100);
  v_年龄     Varchar2(100);
  n_就诊id   Number;
  n_病人id   Number;
  j_Json_Tmp Pljson;
Begin
  --解析入参
  j_Json_Tmp := Pljson(Json_In);
  j_Json     := j_Json_Tmp.Get_Pljson('input');
  v_姓名     := j_Json.Get_String('pati_name');
  v_性别     := j_Json.Get_String('pati_sex');
  v_年龄     := j_Json.Get_String('pati_age');
  n_就诊id   := j_Json.Get_Number('visit_id');
  n_病人id   := j_Json.Get_Number('pati_id');
  Update 输液配药记录
  Set 姓名 = Nvl(v_姓名, 姓名), 性别 = Nvl(v_性别, 性别), 年龄 = Nvl(v_年龄, 年龄)
  Where 病人id = n_病人id And 主页id = n_就诊id;
  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Pivassvr_Patiinfoupdate;
/

Create Or Replace Procedure Zl_Pivassvr_Getworkbatch
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------------------------------------------------------------
  --功能：读取配药工作批次
  --入参：json
  --      可以不传入，暂时为空
  --出参：json
  --output
  --  code                      N 1 应答码：0-失败；1-成功
  --  message                   C 1 应答消息：失败时返回具体的错误信息
  --  item_list
  --        pivas_deptid        N 1     配置中心id
  --        batch               N 1     批次
  --        pivas_time          C 1     配药时间
  -------------------------------------------------------------------------------------------------
  v_Jtmp Varchar2(32767);
Begin
  For R In (Select 配置中心id, 批次, 配药时间 From 配药工作批次 Order By 批次) Loop
    v_Jtmp := v_Jtmp || ',{"pivas_deptid":' || r.配置中心id;
    v_Jtmp := v_Jtmp || ',"batch":' || r.批次;
    v_Jtmp := v_Jtmp || ',"pivas_time":"' || r.配药时间 || '"';
    v_Jtmp := v_Jtmp || '}';
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Pivassvr_Getworkbatch;
/

Create Or Replace Procedure Zl_Pivassvr_Checkordersendroll
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：医嘱回退发送时静配中医嘱检查判断
  --入参：Json_In:格式
  --  input
  --     order_id           N 1 医嘱ID,主医嘱id
  --     send_no            N 1 发送号
  --出参: Json_Out,格式如下
  --  output
  --    code                N 1 应答吗：0-失败；1-成功
  --    message             C 1 应答消息：失败时返回具体的错误信息
  --    pivas_ids           C 1 要销帐的输液记录id串
  ---------------------------------------------------------------------------
  医嘱id_In Number;
  n_发送号  Number;
  j_Json    PLJson;
  j_Tmp     PLJson;
  n_Tmp     Number;
  v_配液ids Varchar2(32767);
Begin
  --解析入参
  j_Tmp     := PLJson(Json_In);
  j_Json    := j_Tmp.Get_Pljson('input');
  医嘱id_In := j_Json.Get_Number('order_id');
  n_发送号  := j_Json.Get_Number('send_no');

  --检查是否是输液配液记录，并是否已经锁定
  Select Decode(Max(是否锁定), 1, 1, 0) Into n_Tmp From 输液配药记录 Where 医嘱id = 医嘱id_In And 发送号 = n_发送号;
  If n_Tmp = 1 Then
    Json_Out := zlJsonOut('当前处理的是输液药品医嘱，已经被输液配置中心锁定，不能回退发送。');
    Return;
  Elsif n_Tmp = 0 Then
    --Zl_输液配药记录_医嘱回退(医嘱id_In, n_发送号, Null, Null);
    --只对状态=1(未配药)的记录处理，如果已经配药了，则通过销账方式处理
    Select Count(ID)
    Into n_Tmp
    From 输液配药记录
    Where 操作状态 In (1, 10) And 医嘱id = 医嘱id_In And 发送号 = n_发送号;
    If n_Tmp > 0 Then
      For R In (Select ID From 输液配药记录 Where 医嘱id = 医嘱id_In And 发送号 = n_发送号) Loop
        v_配液ids := v_配液ids || ',' || r.Id;
      End Loop;
    End If;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功","pivas_ids":"' || Substr(v_配液ids, 2) || '"}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Pivassvr_Checkordersendroll;
/

Create Or Replace Procedure Zl_Pivassvr_Odr_Check
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：超期发送收回静配相关取数和检查
  --入参：Json_In:格式
  --  input
  --     chk_type                                  N 1 检查方式，0-获取列表，1-输液存在配药提示
  --     item_list[]要收回的医嘱列表
  --               order_id                        N 1 医嘱ID,主医嘱id
  --               exe_end_time                    C 1 执行终止时间
  --               advice_note                     C 0 医嘱内容，检查方式为1时此结点可以不传
  --出参: Json_Out,格式如下
  --  output
  --    code                          N 1 应答吗：0-失败；1-成功
  --    message                       C 1 应答消息：失败时返回具体的错误信息
  --    isexist                       N 1 是否存在已经配液的记录，1-存在;0-不存在
  --    pivas_list[]列表
  --               pivas_id           N 1 配液id
  --               order_id           N 1 医嘱id
  --               fee_id             N 1 费用id
  --               fee_item_id        N 1 收费细目id
  --               operator_status    N 1 操作状态
  --               quantity           N 1 数量
  ---------------------------------------------------------------------------
  j_Input          Pljson;
  j_Item           Pljson;
  j_List           Pljson_List := Pljson_List();
  n_相关id         Number(18); --主医嘱ID
  d_执行终止时间   Date;
  v_医嘱内容       Varchar2(30000);
  n_Count          Number;
  v_含配液记录     Varchar2(1000);
  v_配液药销帐申请 Varchar2(4000);
  v_Charge_List    Varchar2(32767);
  v_Json_Out       Varchar2(32767);
  v_医嘱ids        Varchar2(32767);
  n_检查方式       Number;
  v_Error          Varchar2(255);
  Err_Custom Exception;
Begin
  j_Item     := Pljson(Json_In);
  j_Input    := j_Item.Get_Pljson('input');
  n_检查方式 := j_Input.Get_Number('chk_type');
  j_List     := j_Input.Get_Pljson_List('item_list');
  If j_List Is Not Null Then
    For I In 1 .. j_List.Count Loop
      j_Item := Pljson();
      j_Item := Pljson(j_List.Get(I));
    
      n_相关id       := j_Item.Get_Number('order_id');
      d_执行终止时间 := To_Date(j_Item.Get_String('exe_end_time'), 'yyyy-mm-dd hh24:mi:ss');
      If d_执行终止时间 Is Null Then
        d_执行终止时间 := To_Date('1900-01-01', 'yyyy-mm-dd');
      End If;
    
      If n_检查方式 = 1 Then
        --判断是否含已经配液的记录，程序中会有个询提示的交互操作
        Select Count(1)
        Into n_Count
        From 输液配药记录 B
        Where (b.操作状态 In (4, 5, 6, 7, 8) And Nvl(b.是否打包, 0) = 0) And b.医嘱id = n_相关id And b.执行时间 > d_执行终止时间;
        If n_Count > 0 Then
          v_含配液记录 := ',"isexist":1';
          Exit;
        End If;
      Else
        v_配液药销帐申请 := zl_GetSysParameter('配液输液单配药后允许销帐申请', 1345);
        v_医嘱内容       := j_Item.Get_String('advice_note');
        --检查是否是输液配液记录，并是否已经锁定
        Select Count(1)
        Into n_Count
        From 输液配药记录 A
        Where a.医嘱id = n_相关id And a.执行时间 > d_执行终止时间 And a.是否锁定 = 1;
      
        If n_Count > 0 Then
          v_Error := '医嘱"' || v_医嘱内容 || '"是输液药品，已经被输液配置中心锁定，不能超期收回。';
          Raise Err_Custom;
        End If;
      
        For R In (Select b.费用id, b.药品id As 收费细目id, c.操作状态, b.医嘱id, c.Id 配药id, Sum(a.数量) As 数量
                  From 输液配药内容 A, 药品收发记录 B, 输液配药记录 C
                  Where a.收发id = b.Id And a.记录id = c.Id And c.医嘱id = n_相关id And c.执行时间 > d_执行终止时间 And
                        Nvl(c.操作状态, 0) In (1, 2, 3, 4, 5, 6, 7, 8) And
                        Not (c.操作状态 In (4, 5, 6, 7, 8) And Nvl(c.是否打包, 0) = 0 And Nvl(v_配液药销帐申请, '0') = '0')
                  Group By b.费用id, b.药品id, b.医嘱id, c.操作状态, b.医嘱id, c.Id, c.发送号, c.执行时间
                  Order By c.发送号, b.医嘱id, c.执行时间) Loop
        
          v_Charge_List := v_Charge_List || ',{"pivas_id":' || r.配药id;
          v_Charge_List := v_Charge_List || ',"order_id":' || r.医嘱id;
          v_Charge_List := v_Charge_List || ',"fee_id":' || r.费用id;
          v_Charge_List := v_Charge_List || ',"fee_item_id":' || r.收费细目id;
          v_Charge_List := v_Charge_List || ',"operator_status":' || r.操作状态;
          v_Charge_List := v_Charge_List || ',"quantity":' || Zljsonstr(r.数量, 1); --N 1 对照数量  
          v_Charge_List := v_Charge_List || '}';
        End Loop;
      End If;
    
    End Loop;
    If v_Charge_List Is Not Null Then
      v_Charge_List := ',"pivas_list":[' || Substr(v_Charge_List, 2) || ']';
    End If;
  End If;

  v_Json_Out := '{"code":1,"message":"成功"';
  v_Json_Out := v_Json_Out || v_含配液记录;
  v_Json_Out := v_Json_Out || v_Charge_List;
  v_Json_Out := v_Json_Out || '}';
  Json_Out   := '{"output":' || v_Json_Out || '}';

Exception
  When Err_Custom Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(v_Error) || '"}}';
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Pivassvr_Odr_Check;
/

CREATE OR REPLACE Procedure Zl_Pivassvr_Overdue_Recovery
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：超期发送收静配材相关处理
  --入参：Json_In:格式
  --  input
  --     operator_name                        C 1 操作员姓名
  --     pivas_list[]静配销帐列表
  --                  pivas_id               N 1 配液id
  --                  auto_aduit             N 1 是否自动审核 0-不审核,1-要自动审核
  --                  request_time           C 1 申请时间
  --                  reason                 C 1 销帐原因
  --出参: Json_Out,格式如下
  --  output
  --    code                          N 1 应答吗：0-失败；1-成功
  --    message                       C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Input      Pljson;
  j_Item       Pljson;
  j_List       Pljson_List := Pljson_List();
  v_操作员姓名 Varchar2(300);
  n_配液id     Number;
  n_自动审核   Number;
  d_申请时间   Date;
  v_销帐原因   Varchar2(4000);
Begin
  j_Item  := Pljson(Json_In);
  j_Input := j_Item.Get_Pljson('input');

  v_操作员姓名 := j_Input.Get_String('operator_name');

  j_List := j_Input.Get_Pljson_List('pivas_list');
  If j_List Is Not Null Then
    For I In 1 .. j_List.Count Loop
      j_Item     := Pljson();
      j_Item     := Pljson(j_List.Get(I));
      n_配液id   := j_Item.Get_Number('pivas_id');
      n_自动审核 := j_Item.Get_Number('auto_aduit');
      d_申请时间 := To_Date(j_Item.Get_String('request_time'), 'yyyy-mm-dd hh24:mi:ss');
      v_销帐原因 := j_Item.Get_String('reason');
    
      Insert Into 输液配药状态
        (配药id, 操作类型, 操作人员, 操作时间, 操作说明)
      Values
        (n_配液id, 9, v_操作员姓名, d_申请时间, v_销帐原因);
    
      Update 输液配药记录 Set 操作人员 = v_操作员姓名, 操作时间 = d_申请时间, 操作状态 = 9 Where ID = n_配液id;
    
      If n_自动审核 = 1 Then
        --加一秒,这里可能和费用销帐申请的审核时间会相差一秒,影响不大
        d_申请时间 := d_申请时间 + 1 / 24 / 60 / 60;
        Insert Into 输液配药状态
          (配药id, 操作类型, 操作人员, 操作时间,操作说明)
        Values
          (n_配液id, 10, v_操作员姓名, d_申请时间,v_销帐原因);
        Update 输液配药记录 Set 操作人员 = v_操作员姓名, 操作时间 = d_申请时间, 操作状态 = 10 Where ID = n_配液id;
      End If;
    End Loop;
  End If;

  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Pivassvr_Overdue_Recovery;
/

Create Or Replace Procedure Zl_Pivassvr_Getchargeerrdata
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：静配销帐申请自动审核时，费用域失败后修复时获取销帐数量
  --入参：Json_In:格式
  --  input
  --     order_id                  N 1 医嘱ID,主医嘱id
  --     send_no                   N 1 发送号
  --     pivas_id                  N 1 配液记录id，销帐申请自动审核后出现异常的配液记录id
  --     item_list[]
  --            rcpdtl_id          N 1 处方明细id
  --            rcp_no             C 1 处方单号
  --            drug_id            N 1 药品id
  --            quantity           N 1 医嘱发送后的总数量,剂量单位的总量，来源于医嘱发送记录，按理说应该在医嘱发送记录中冗余一个药品id字段因为长嘱按品种下达时取不到
  --            order_id           N 1 医嘱ID
  --出参: Json_Out,格式如下
  --  output
  --    code                N 1 应答吗：0-失败；1-成功
  --    message             C 1 应答消息：失败时返回具体的错误信息
  --    item_list[]
  --            rcpdtl_id          N 1 处方明细id
  --            rcp_no             C 1 处方单号
  --            drug_id            N 1 药品id
  --            quantity           N 1 数量销帐申请数量
  --            order_id           N 1 医嘱ID  
  --            request_time       C 1 申请时间
  --            audit_time         C 1 审核时间
  --            reason             C 1 销帐原因
  --            request_operator   C 1 销帐操作员
  --            request_code       C 1 销帐操作员编码
  ---------------------------------------------------------------------------
  n_医嘱id     Number(18);
  n_发送号     Number(18);
  n_配药id     Number(18);
  j_Json       Pljson;
  j_Tmp        Pljson;
  j_Item       Pljson;
  n_Tmp        Number;
  v_Jtmp       Varchar2(32767);
  n_List_Cnt   Number;
  j_Jsonlist   Pljson_List;
  v_申请时间   Varchar2(30);
  v_审核时间   Varchar2(30);
  n_次数       Number(3);
  n_最后一次   Number(3);
  n_数量       Number(16, 5);
  n_药品id     Number(18);
  n_销帐数量   Number(16, 5);
  v_销帐原因   Varchar2(1000);
  v_操作人员   Varchar2(1000);
  v_操作员编号 Varchar2(1000);
Begin
  --解析入参
  j_Tmp      := Pljson(Json_In);
  j_Json     := j_Tmp.Get_Pljson('input');
  n_医嘱id   := j_Json.Get_Number('order_id');
  n_发送号   := j_Json.Get_Number('send_no');
  n_配药id   := j_Json.Get_Number('pivas_id');
  j_Jsonlist := j_Json.Get_Pljson_List('item_list');
  If Not j_Jsonlist Is Null Then
    n_List_Cnt := j_Jsonlist.Count;
  Else
    n_List_Cnt := 1;
  End If;

  For R In (Select To_Char(a.操作时间, 'yyyy-mm-dd hh24:mi:ss') 操作时间, a.操作类型, a.操作说明, a.操作人员
            From 输液配药状态 A
            Where a.操作类型 In (9, 10) And a.配药id = n_配药id) Loop
    If r.操作类型 = 9 Then
      v_申请时间 := r.操作时间;
      v_销帐原因 := r.操作说明;
      v_操作人员 := r.操作人员;
      Select a.编号 Into v_操作员编号 From 人员表 A Where a.姓名 = v_操作人员;
    Else
      v_审核时间 := r.操作时间;
    End If;
  End Loop;
  n_次数 := 0;
  For R In (Select a.Id, a.执行时间
            From 输液配药记录 A
            Where a.医嘱id = n_医嘱id And a.发送号 = n_发送号
            Order By a.执行时间) Loop
    n_次数 := n_次数 + 1;
    If r.Id = n_配药id Then
      n_最后一次 := n_次数;
    End If;
  End Loop;

  If n_最后一次 = n_次数 Then
    n_最后一次 := 1;
  Else
    n_最后一次 := 0;
  End If;

  For I In 1 .. n_List_Cnt Loop
    j_Item     := Pljson(j_Jsonlist.Get(I));
    n_数量     := j_Item.Get_Number('quantity');
    n_药品id   := j_Item.Get_Number('drug_id');
    n_销帐数量 := n_数量 / n_次数;
    If n_最后一次 = 1 Then
      n_销帐数量 := n_数量 - (n_数量 / n_次数) * (n_次数 - 1);
    End If;
    Select (n_销帐数量 / a.剂量系数) Into n_销帐数量 From 药品规格 A Where a.药品id = n_药品id;
  
    v_Jtmp := v_Jtmp || ',{"rcpdtl_id":' || j_Item.Get_Number('rcpdtl_id');
    v_Jtmp := v_Jtmp || ',"rcp_no":"' || j_Item.Get_String('rcp_no') || '"';
    v_Jtmp := v_Jtmp || ',"drug_id":' || j_Item.Get_Number('drug_id');
    v_Jtmp := v_Jtmp || ',"quantity":' || Zljsonstr(n_销帐数量, 1);
    v_Jtmp := v_Jtmp || ',"order_id":' || j_Item.Get_Number('order_id');
    v_Jtmp := v_Jtmp || ',"request_time":"' || v_申请时间 || '"';
    v_Jtmp := v_Jtmp || ',"audit_time":"' || v_审核时间 || '"';
    v_Jtmp := v_Jtmp || ',"reason":"' || Zljsonstr(v_销帐原因) || '"';
    v_Jtmp := v_Jtmp || ',"request_operator":"' || Zljsonstr(v_操作人员) || '"';
    v_Jtmp := v_Jtmp || ',"request_code":"' || Zljsonstr(v_操作员编号) || '"';
    v_Jtmp := v_Jtmp || '}';
    j_Item := Pljson();
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Pivassvr_Getchargeerrdata;
/