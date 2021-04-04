Create Or Replace Procedure Zl_Patisvr_Batupdoutpativisit
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:批量更新病人的就诊状态和就诊诊室
  --入参：Json_In:格式
  --    input
  --      visit_status      N 1 就诊状态
  --      pati_list[]      数组
  --        pati_ids       C 1 病人id,多个用","分隔
  --        visit_room     C 1 诊室

  --出参: Json_Out,格式如下
  --  output
  --    code                N   1   应答码：0-失败；1-成功
  --    message             C   1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------

  j_Json     Pljson;
  j_Jsonin   Pljson;
  j_Json_Tmp Pljson;
  j_List     Pljson_List := Pljson_List();
  n_状态     病人信息.就诊状态%Type;
  v_病人ids  Varchar2(3000);
  v_诊室     病人信息.就诊诊室%Type;

Begin
  --解析入参
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_状态   := j_Json.Get_Number('visit_status');
  If n_状态 Is Null Then
    Json_Out := Zljsonout('未传入就诊状态，请检查！');
    Return;
  End If;
  j_List := j_Json.Get_Pljson_List('pati_list');
  If j_List Is Null Then
    Json_Out := Zljsonout('未传入病人id和诊室，请检查！');
    Return;
  End If;
  For I In 1 .. j_List.Count Loop
    j_Json_Tmp := Pljson();
    j_Json_Tmp := Pljson(j_List.Get(I));
    v_病人ids  := j_Json_Tmp.Get_String('pati_ids');
    v_诊室     := j_Json_Tmp.Get_String('visit_room');
    Update 病人信息
    Set 就诊状态 = n_状态, 就诊诊室 = v_诊室
    Where 病人id In (Select /*+cardinality(x,10)*/
                    x.Column_Value As 病人id
                   From Table(Cast(f_Num2list(v_病人ids) As Zltools.t_Numlist)) X) And 就诊状态 In (1, 2);
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Batupdoutpativisit;
/
Create Or Replace Procedure Zl_Patisvr_Calc_Age
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  --------------------------------------------------------------------------------------------------------------------
  --功能:根据出生日期计算年龄.当天登记病人,保持年龄不变.
  --入参：Json_In,格式如下
  --  input
  --    pati_id            N 1 病人ID
  --    birthdate          C 1 出生日期
  --    calc_date          C 1 计算日期
  --出参: Json_Out,格式如下
  --  output
  --    code               N 1 应答吗：0-失败；1-成功
  --    message            C 1 应答消息：失败时返回具体的错误信息
  --    age                C 1  返回:1天以内：X小时[X分钟],1天至1月以内：X天[X小时],1月至1岁以内：X月[X天],1岁至儿童年龄上限：X岁[X月],>=儿童年龄上限：X岁
  --                            说明:1天以内，是指按出生日期24小时算;1月以内，是指对天计算；比如7.8日出生，8.8日才算1月;1岁以内，也是对天计算。;“以内”都是指“<”。
  --------------------------------------------------------------------------------------------------------------------
  j_Json      Pljson;
  j_Jsonin    Pljson;
  n_Pati_Id   病人信息.病人id%Type;
  d_Birthdate Date;
  d_Calc_Date Date;

  v_Age Varchar2(20); --由于病人信息等相关表的年龄字段为10个字符，所以最大允许10个字符或5个汉字
Begin
  --解析入参
  j_Jsonin  := Pljson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  n_Pati_Id := j_Json.Get_Number('pati_id');

  d_Birthdate := To_Date(j_Json.Get_String('birthdate'), 'YYYY-MM-DD HH24:MI:SS');
  d_Calc_Date := To_Date(j_Json.Get_String('calc_date'), 'YYYY-MM-DD HH24:MI:SS');
  v_Age       := Zl_Age_Calc(n_Pati_Id, d_Birthdate, d_Calc_Date);
  Json_Out    := '{"output":{"code":1,"message":"成功","age":"' || Zljsonstr(v_Age) || '"}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Calc_Age;
/
Create Or Replace Procedure Zl_Patisvr_Checkcardexist
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:检查卡类别是否存在
  --入参：Json_In:格式
  --    input
  --      card_type_id      N 1 卡类别ID
  --      pati_id           N 1 病人id
  --      card_no           C 1 医疗卡号

  --出参: Json_Out,格式如下
  --  output
  --    code              N   1   应答码：0-失败；1-成功
  --    message           C   1   应答消息：失败时返回具体的错误信息
  --    pati_id           N   1   当前使用这张卡的病人id。针对传入卡号时有效
  --    exist             N   1   当前病人已经发过同类型的医疗卡。针对传入病人id时有效
  ---------------------------------------------------------------------------

  j_Json       Pljson;
  j_Jsonin     Pljson;
  n_病人id     病人医疗卡信息.病人id%Type;
  n_卡类别id   病人医疗卡信息.卡类别id%Type;
  v_卡号       病人医疗卡信息.卡号%Type;
  n_病人id_Out 病人医疗卡信息.病人id%Type;
  n_Exist      Number(2);

Begin
  --解析入参
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_病人id   := j_Json.Get_Number('pati_id');
  n_卡类别id := j_Json.Get_Number('card_type_id');
  v_卡号     := j_Json.Get_String('card_no');

  If Nvl(n_卡类别id, 0) = 0 Then
    Json_Out := Zljsonout('未传入卡类别id，请检查！');
    Return;
  End If;

  If v_卡号 Is Not Null Then
    Select Nvl(Max(病人id), 0) Into n_病人id_Out From 病人医疗卡信息 Where 卡号 = v_卡号 And 卡类别id = n_卡类别id;
  Else
    Select Count(1)
    Into n_Exist
    From 病人医疗卡信息
    Where 病人id = Nvl(n_病人id, 0) And 卡类别id = n_卡类别id And Nvl(状态, 0) = 0 And Rownum < 2;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功","exist":' || Nvl(n_Exist, 0) || ',"pati_id":' || Nvl(n_病人id_Out, 0) || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Checkcardexist;
/
Create Or Replace Procedure Zl_Patisvr_Checkdepositerrorno
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
  --    occasion          N 1 场合：1-医疗卡发放;2-病人信息登记（针对只传一个NO是有效）
  ---------------------------------------------------------------------------
  n_病人id  病人结算异常记录.病人id%Type;
  v_Nos     Varchar2(3000);
  j_Json    Pljson;
  j_Jsonin  Pljson;
  v_Nos_Out Varchar2(3000);
  n_场合    病人结算异常记录.操作场景%Type;

Begin
  --解析入参
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_病人id := Nvl(j_Json.Get_Number('pati_id'), 0);
  v_Nos    := j_Json.Get_String('bill_nos');

  If Nvl(v_Nos, '-') = '-' Then
    Json_Out := Zljsonout('未传入NO，请检查！');
    Return;
  End If;

  Select /*+cardinality(B,10)*/
   f_List2str(Cast(Collect(b.Column_Value) As t_Strlist)), Nvl(Max(a.操作场景), 0)
  Into v_Nos_Out, n_场合
  From 病人结算异常记录 A, Table(f_Str2list(v_Nos)) B
  Where a.操作场景 In (1, 2) And (a.预交单号 = b.Column_Value Or a.医疗卡单号 = b.Column_Value) And
        Decode(n_病人id, 0, 0, a.病人id) = n_病人id;

  Json_Out := '{"output":{"code":1,"message":"成功","bill_nos":"' || v_Nos_Out || '","occasion":' || n_场合 || '}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Checkdepositerrorno;
/
Create Or Replace Procedure Zl_Patisvr_Checkidcardunique
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：根据病人ID和身份证号检查同一身份证只能对应一个建档病人
  --入参：Json_In:格式
  --  input
  --   pati_idcard           C   1  身份证号
  --    pati_id              N   1  病人ID
  --出参: Json_Out,格式如下
  --  output
  --    code                 N   1   应答吗：0-失败；1-成功
  --    message              C   1   应答消息：失败时返回具体的错误信息
  --    isexist              N   1   是否存在（0-不存在 1-存在）
  ---------------------------------------------------------------------------
  n_病人id   病人信息.病人id%Type;
  v_身份证号 病人信息.身份证号%Type;
  n_Count    Number;
  j_Json     Pljson;
  j_Jsonin   Pljson;

  n_Isexist Number(1);
Begin

  --解析入参
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  v_身份证号 := j_Json.Get_String('pati_idcard');
  n_病人id   := j_Json.Get_Number('pati_id');
  Select Count(1) Into n_Count From 病人信息 A Where a.身份证号 = v_身份证号 And a.病人id <> n_病人id And Rownum < 2;
  If n_Count > 0 Then
    n_Isexist := 1;
  Else
    n_Isexist := 0;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功","isexist":' || n_Isexist || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Checkidcardunique;
/
Create Or Replace Procedure Zl_Patisvr_Checkinsnoisexist
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：检查指定的医保号是否存在
  --入参：Json_In:格式
  --  input
  --    insurance_num        C   1  医保号
  --出参: Json_Out,格式如下
  --  output
  --    code                 N   1   应答吗：0-失败；1-成功
  --    message              C   1   应答消息：失败时返回具体的错误信息
  --    isexist              N   1  1-存在;0-不存在
  v_医保号 病人信息.医保号%Type;
  n_Exist  Number(2);
  j_Json   Pljson;
  j_Jsonin Pljson;
Begin

  --解析入参
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  v_医保号 := j_Json.Get_String('insurance_num');
  If Nvl(v_医保号, '-') = '-' Then
    Json_Out := Zljsonout('未传入病人医保号');
    Return;
  End If;
  Select Count(1) Into n_Exist From 病人信息 Where 医保号 = v_医保号;
  Json_Out := '{"output":{"code":1,"message":"成功","isexist":' || n_Exist || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Checkinsnoisexist;

/
Create Or Replace Procedure Zl_Patisvr_Checkoutnoisexist
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：检查指定的门诊号是否已经被使用
  --入参：Json_In:格式
  --  input
  --    pati_id              N   1  病人ID:当前操作的病人
  --    outpatient_num       C   1  门诊号
  --出参: Json_Out,格式如下
  --  output
  --    code                 N   1   应答吗：0-失败；1-成功
  --    message              C   1   应答消息：失败时返回具体的错误信息
  --    isexist              N   1  1-存在;0-不存在
  ---------------------------------------------------------------------------
  n_病人id 病人信息.病人id%Type;
  n_门诊号 病人信息.门诊号%Type;
  n_Exist  Number(2);
  j_Json   Pljson;
  j_Jsonin Pljson;
Begin

  --解析入参
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');
  n_门诊号 := To_Number(j_Json.Get_String('outpatient_num'));
  Select Count(1) Into n_Exist From 病人信息 Where 门诊号 = n_门诊号 And 病人id <> n_病人id;
  Json_Out := '{"output":{"code":1,"message":"成功","isexist":' || n_Exist || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Checkoutnoisexist;
/
Create Or Replace Procedure Zl_Patisvr_Checkpatirealname
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：根据病人ID,判断该病人是否进行了实名认证
  --入参：Json_In:格式
  --  input
  --    pati_id              N   1  病人ID
  --出参: Json_Out,格式如下
  --  output
  --    code                 N   1   应答吗：0-失败；1-成功
  --    message              C   1   应答消息：失败时返回具体的错误信息
  --    isexist              N   1   是否存在（0-不存在 1-存在）
  ---------------------------------------------------------------------------
  n_Count   Number;
  n_Isexist Number;
  j_Json    Pljson;
  j_Jsonin  Pljson;
  n_病人id  Number;
Begin
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');
  Select Count(1) Into n_Count From 病人实名信息 Where 病人id = n_病人id And Rownum < 2;
  If n_Count > 0 Then
    n_Isexist := 1;
  Else
    n_Isexist := 0;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功","isexist":' || n_Isexist || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Checkpatirealname;
/
Create Or Replace Procedure Zl_Patisvr_Checkregisterinpati
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:入院登记检查
  --入参：Json_In:格式
  --   input
  --      type              N 1 调用类型  1-新增登记;2-修改登记
  --      pati_id           N 1 病人id
  --      pati_idcard       N 1 身份证号
  --      isnew             N  1-新病人;0-非新病人
  --出参  json
  --output
  --     code               N 1 应答码：0-失败；1-成功
  --     message            C 1 应答消息： 失败时返回具体的错误信息
  ---------------------------------------------------------------------------

  j_Json        Pljson;
  j_Jsonin      Pljson;
  n_Type        Number(5);
  n_Pati_Id     Number(18);
  v_Pati_Idcard Varchar2(18);
  n_Isnew       Number(1);
  n_Count       Number(5);
  n_Uniqueid    Number(1);
  v_Msg         Varchar2(200);
Begin
  --解析入参
  j_Jsonin      := Pljson(Json_In);
  j_Json        := j_Jsonin.Get_Pljson('input');
  n_Type        := j_Json.Get_Number('type');
  n_Pati_Id     := j_Json.Get_Number('pati_id');
  v_Pati_Idcard := j_Json.Get_String('pati_idcard');
  n_Isnew       := j_Json.Get_Number('isnew');

  --判断病人是否锁定
  Select Count(病人id) Into n_Count From 病人信息 Where 病人id = n_Pati_Id;

  If n_Count <> 0 Then
    Zl_病人信息_锁定检查(n_Pati_Id);
  End If;

  --身份证号不等于空,根据系统参数判读是否唯一建档病人
  If v_Pati_Idcard Is Not Null And ((n_Isnew = 1 And n_Type = 1) Or n_Type = 2) Then
    n_Uniqueid := Nvl(zl_GetSysParameter(279), 0);
    If n_Uniqueid = 1 Then
      Select Count(1) Into n_Count From 病人信息 Where 身份证号 = v_Pati_Idcard And 病人id <> Nvl(n_Pati_Id, 0);
      If n_Count <> 0 Then
        v_Msg    := '已经存在身份证号为' || v_Pati_Idcard || '的病人,不能再录入相同的身份证号!';
        Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(v_Msg) || '"}}';
        Return;
      End If;
    End If;
  End If;
  Json_Out := Zljsonout('成功', 1);
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Checkregisterinpati;
/
Create Or Replace Procedure Zl_Patisvr_Checkreturncard
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:判断当前卡是否允许退卡，如果不允许退卡，返回提示内容；否则返回NULL
  --入参：Json_In:格式
  -- input
  --   occasion             N 1 场合：1-医疗卡发放；2-门诊挂号；
  --   pati_id              N 1 当前病人id
  --   gvcard_type_id       N 1 卡类别ID
  --   gvcard_no            C 1 医疗卡号
  --出参: Json_Out,格式如下
  --  output
  --    code               C 1 应答码：0-失败；1-成功
  --    message            C 1 应答消息：失败时返回具体的错误信息
  --    isexist            N 1 是否存在，1-存在;0-不存在

  ---------------------------------------------------------------------------
  j_Json     Pljson;
  j_Jsonin   Pljson;
  n_病人id   病人医疗卡信息.病人id%Type;
  n_卡类别id 病人医疗卡信息.卡类别id%Type;
  v_卡号     病人医疗卡信息.卡号%Type;
  n_场合     Number(2);
  n_模块     Zlparameters.模块%Type;
  v_Msg      Varchar2(3000);

Begin
  --解析入参
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');

  n_病人id   := j_Json.Get_Number('pati_id');
  n_卡类别id := j_Json.Get_Number('gvcard_type_id');
  v_卡号     := j_Json.Get_String('gvcard_no');
  n_场合     := j_Json.Get_Number('occasion');
  If Nvl(n_场合, 0) = 0 Then
    Json_Out := Zljsonout('未传入场合，请检查！');
    Return;
  End If;
  If n_场合 = 1 Then
    n_模块 := 1107;
  Else
    n_模块 := 1111;
  End If;
  v_Msg := Zl1_Ex_Refundcard_Check(n_模块, n_病人id, n_卡类别id, v_卡号);

  If Nvl(v_Msg, '-') = '-' Then
    Json_Out := Zljsonout('成功', 1);
  Else
    Json_Out := Zljsonout(v_Msg);
  End If;

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Checkreturncard;
/
Create Or Replace Procedure Zl_Patisvr_Chkcardchangevalid
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:检查医疗卡变动前的合法性检查
  --入参：Json_In:格式
  --  input
  --    oper_state          N  1  操作状态:1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失),7-终止时间调整
  --    cardtype_id         N  1  卡类别ID
  --    cardno              C  1  卡号：发卡、补卡及换卡等的卡号或其他操作的原始卡号
  --    new_cardno          C     新卡号:换卡时的新卡号
  --    pati_id             N  1  病人ID
  --    err_id              N     异常业务ID
  --出参: Json_Out,格式如下
  --  output
  --    code                N   1   应答码：0-失败；1-成功
  --    message             C   1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Json   PLJson;
  j_Jsonin PLJson;

  n_操作状态 Number(3);
  n_卡类别id Number(18);
  v_卡号     Varchar2(100);
  v_新卡号   Varchar2(100);
  n_病人id   Number;
  v_应答信息 Varchar2(32767);
  n_应答码   Number(5);
  n_异常id   Number(18);
Begin
  --解析入参
  j_Jsonin := PLJson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');

  n_操作状态 := j_Json.Get_Number('oper_state');
  n_卡类别id := j_Json.Get_Number('cardtype_id');
  v_卡号     := j_Json.Get_String('cardno');
  v_新卡号   := j_Json.Get_String('new_cardno');
  n_病人id   := j_Json.Get_Number('pati_id');
  n_异常id   := Nvl(j_Json.Get_Number('err_id'), 0);

  Zl_医疗卡变动_Insert_Check(n_操作状态, n_卡类别id, v_卡号, v_新卡号, n_病人id, 0, n_应答码, v_应答信息,n_异常id);

  Json_Out := zlJsonOut(v_应答信息, n_应答码);
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Patisvr_Chkcardchangevalid;
/

Create Or Replace Procedure Zl_Patisvr_Confirmcardchange
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:医疗卡变动确认
  --入参：Json_In:格式
  --  input
  --    change_id           N  1  变动id
  --    pati_id             N  1  病人ID
  --    cardtype_id         N  1 卡类别ID
  --    card_no             C  1 医疗卡号
  --    card_notes          C  1 变动原因
  --    card_pwd            C  1 密码
  --    card_use_endtime    C  1  终止使用时间
  --出参: Json_Out,格式如下
  --  output
  --    code                N  1  应答码：0-失败；1-成功
  --    message             C  1  应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Json   Pljson;
  j_Jsonin Pljson;
  n_变动id Number(18);
  n_病人id Number(18);

  n_卡类别id 病人医疗卡变动.卡类别id%Type;
  v_卡号     病人医疗卡变动.卡号%Type;
  v_变动原因 病人医疗卡变动.变动原因%Type;
  v_密码     病人医疗卡变动.原密码%Type;
  d_终止时间 病人医疗卡变动.终止使用时间%Type;
Begin
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');

  n_变动id := j_Json.Get_Number('change_id');
  n_病人id := j_Json.Get_Number('pati_id');

  n_卡类别id := j_Json.Get_Number('cardtype_id');
  v_卡号     := j_Json.Get_String('card_no');
  v_变动原因 := j_Json.Get_String('card_notes');
  v_密码     := j_Json.Get_String('card_pwd');
  d_终止时间 := To_Date(j_Json.Get_String('card_use_endtime'), 'yyyy-mm-dd hh24:mi:ss');
  If Nvl(n_卡类别id, 0) = 0 Then
    n_卡类别id := Null;
  End If;

  Zl_病人医疗卡变动_Confirm(n_变动id, n_病人id, n_卡类别id, v_卡号, v_变动原因, v_密码, d_终止时间);

  Json_Out := Zljsonout('成功', 1);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Confirmcardchange;
/
Create Or Replace Procedure Zl_Patisvr_Delcardchangeinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:删除医疗卡变动信息
  --入参：Json_In:格式
  --  input
  --  change_id             N 1 变动id
  --  cardtype_id           C 1 卡类别id
  --  cardno                C 1 卡号
  --  pati_id               N 1 病人ID

  --出参: Json_Out,格式如下
  --  output
  --    code                N  1  应答码：0-失败；1-成功
  --    message             C  1  应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Json   Pljson;
  j_Jsonin Pljson;

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_变动id   Number(18);
  n_卡类别id Number(18);
  v_卡号     Varchar2(100);
  n_病人id   Number(18);

Begin
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');

  n_变动id   := j_Json.Get_Number('change_id');
  n_卡类别id := j_Json.Get_Number('cardtype_id');
  v_卡号     := j_Json.Get_String('cardno');
  n_病人id   := j_Json.Get_Number('pati_id');

  Zl_医疗卡变动记录_Delete(n_变动id, n_卡类别id, v_卡号, n_病人id);

  Json_Out := Zljsonout('成功', 1);

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Delcardchangeinfo;
/
Create Or Replace Procedure Zl_Patisvr_Deletepatiinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:删除指定病人
  --入参：Json_In:格式
  --    input
  --      pati_id           N  1  病人id
  --出参: Json_Out,格式如下
  --  output
  --    code              N   1   应答码：0-失败；1-成功
  --    message           C   1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Json   Pljson;
  j_Jsonin Pljson;
  n_病人id 病人信息.病人id%Type;

Begin
  --解析入参
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');

  n_病人id := j_Json.Get_Number('pati_id');

  Zl_病人信息_Delete_s(n_病人id);

  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Deletepatiinfo;
/
Create Or Replace Procedure Zl_Patisvr_Deletepatiphoto
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:删除指定病人相片
  --入参：Json_In:格式
  --    input
  --      pati_id           N  1  病人id
  --出参: Json_Out,格式如下
  --  output
  --    code              N   1   应答码：0-失败；1-成功
  --    message           C   1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Json   Pljson;
  j_Jsonin Pljson;
  n_病人id 病人照片.病人id%Type;

Begin
  --解析入参
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');

  Zl_病人照片_Delete(n_病人id);

  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Deletepatiphoto;

/
Create Or Replace Procedure Zl_Patisvr_Getblacklistbycons
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  --------------------------------------------------------------------------------------------------
  --功能:实名认证前的检查  
  --入参 JSOM格式
  --input
  --  pati_id        N 1 病人id
  --  operat_type    C 1 行为类别
  --出参 JSON格式
  --output
  --  code            N 1 应答码：0-失败；1-成功
  --  message         C 1 应答消息：失败时返回具体的错误信息
  --  black_list[]
  --     pati_id         N 1 病人id
  --     sign            C 1 附加信息
  --------------------------------------------------------------------------------------------------
  j_Json     Pljson;
  j_Jsonin   Pljson;
  n_病人id   Number;
  v_执行类别 Varchar2(200);
  v_Jtmp     Varchar2(32767);
  c_Jtmp     Clob;
Begin
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_病人id   := j_Json.Get_Number('pati_id');
  v_执行类别 := j_Json.Get_String('operat_type');

  For c_不良 In (Select 病人id, 附加信息
               From 病人不良记录
               Where 行为类别 = v_执行类别 And ((病人id = n_病人id And Nvl(n_病人id, 0) <> 0) Or Nvl(n_病人id, 0) = 0)) Loop
  
    v_Jtmp := v_Jtmp || ',{"pati_id":' || c_不良.病人id;
    v_Jtmp := v_Jtmp || ',"sign":"' || Zljsonstr(c_不良.附加信息) || '"';
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
    Json_Out := '{"output":{"code":1,"message":"成功","black_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","black_list":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getblacklistbycons;
/
Create Or Replace Procedure Zl_Patisvr_Getblackregnos
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:根据病人ID,获取进入黑名单的挂号单号
  --入参：Json_In:格式
  --    input
  --      pati_id  N  1  病人id
  --出参: Json_Out,格式如下
  --  output
  --    code              N   1   应答码：0-失败；1-成功
  --    message           C   1   应答消息：失败时返回具体的错误信息
  --    last_time         C  1  进入不良记录的最后一次时间：yyyy-mm-dd hh24:mi:ss
  --    regnos            C  1  进入黑名单的挂号单号,多个用逗号分离

  ---------------------------------------------------------------------------
  j_Json         Pljson;
  j_Jsonin       Pljson;
  n_病人id       病人不良记录.病人id%Type;
  d_计算日期     Date;
  v_最后预约时间 Varchar2(30);
  v_Nos          Varchar2(32680);

Begin
  --解析入参
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');

  If Nvl(n_病人id, 0) = 0 Then
    d_计算日期 := Trunc(Sysdate) - 1 / 24 / 60 / 60; -- 缺省计算头一天的数据
    For c_预约 In (Select Distinct a.附加信息
                 From 病人不良记录 A
                 Where 行为类别 = '预约挂号' And a.发生时间 >= Trunc(d_计算日期) And a.发生时间 <= d_计算日期) Loop
    
      v_Nos := Nvl(v_Nos, '') || ',' || c_预约.附加信息;
    End Loop;
  Else
    --主要是针对历史数据,可能还存在
    Select To_Char(Nvl(Max(加入时间), To_Date('2000-01-01', 'yyyy-mm-dd')), 'yyyy-mm-dd hh24:mi:ss')
    Into v_最后预约时间
    From 病人不良记录 A
    Where 病人id = n_病人id And (行为类别 = '预约超期' Or (加入原因 = '预约失约次数过多,自动进入黑名单' And 行为类别 = '其他'));
  
    For c_预约 In (Select Distinct 附加信息 From 病人不良记录 Where 病人id = n_病人id And 行为类别 = '预约挂号') Loop
      v_Nos := Nvl(v_Nos, '') || ',' || c_预约.附加信息;
    End Loop;
  
  End If;
  If v_Nos Is Not Null Then
    v_Nos := Substr(v_Nos, 2);
  End If;

  Json_Out := '{"output":{"code":1,"message":"成功","last_time":"' || v_最后预约时间 || '","regnos":"' || Zljsonstr(v_Nos) ||
              '"}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getblackregnos;
/
Create Or Replace Procedure Zl_Patisvr_Getcardlastchange
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:获取指定的医疗卡的最后一次变动信息
  --入参：Json_In:格式
  --    input
  --      pati_id           N 1 病人id
  --      cardtype_id       N 1 卡类别ID
  --      card_no             C 1 卡号
  --出参: Json_Out,格式如下
  --  output
  --    code              N   1   应答码：0-失败；1-成功
  --    message           C   1   应答消息：失败时返回具体的错误信息
  --    change_type       N   1   最后一次的变动类型
  ---------------------------------------------------------------------------

  j_Json   Pljson;
  j_Jsonin Pljson;

  n_病人id   病人医疗卡信息.病人id%Type;
  v_卡号     病人医疗卡信息.卡号%Type;
  n_卡类别id 病人医疗卡信息.卡类别id%Type;
  n_变动类型 病人医疗卡变动.变动类别%Type;

Begin
  --解析入参
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_病人id   := j_Json.Get_Number('pati_id');
  v_卡号     := j_Json.Get_String('card_no');
  n_卡类别id := j_Json.Get_Number('cardtype_id');

  If Nvl(n_病人id, 0) = 0 Then
    Json_Out := Zljsonout('未传入病人信息，请检查！');
    Return;
  End If;

  If Nvl(n_卡类别id, 0) = 0 Or Nvl(v_卡号, '-') = '-' Then
    Json_Out := Zljsonout('未传入卡类别或卡号，请检查！');
    Return;
  End If;

  Select Max(变动类别)
  Into n_变动类型
  From (With 医疗卡变动 As (Select 病人id, ID, 变动类别, 变动时间
                       From 病人医疗卡变动 Bd
                       Where Bd.卡号 = v_卡号 And 卡类别id = n_卡类别id And 病人id = n_病人id)
         Select a.变动类别
         From 医疗卡变动 A, (Select Max(变动时间) As 变动时间 From 医疗卡变动 C) B
         Where a.变动时间 = b.变动时间) A;


  Json_Out := '{"output":{"code":1,"message":"成功","change_type":' || Nvl(n_变动类型, 0) || '}}';

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Getcardlastchange;

/
Create Or Replace Procedure Zl_Patisvr_Getcardtypes
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:获取医疗卡类别数据
  --入参：Json_In:格式
  --    input
  --      cardtype_id          N 0 卡类别id:NULL表示不按卡类别ID查找
  --      query_type           N 1 查询类型:0-所有信息;1-基本信息(返回:id,编码，名称,卡号长度,前缀文本,是否启用,结算方式,是否全退,是否退现)
  --      cert_cardtype        N 0 只读取证件作为卡类别的医疗卡类别:1-只读取是否证件=1的医疗卡;0-全部读取
  --      dffective_cardtype   N 0 只读取有效的卡类别:1-只读取有效的卡类别;0-全部读取
  --      cardtype_name        C 0 卡名称，传入卡名称或特定项目名称进行过滤，默认传空
  --出参: Json_Out,格式如下
  --  output
  --    code                   N   1   应答码：0-失败；1-成功
  --    message                C   1   应答消息：失败时返回具体的错误信息
  --    type_list[]            C   1   支持的卡类别列表
  --        cardtype_id        N   1   ID
  --        cardtype_code      C   1   编码
  --        cardtype_name      C   1   名称
  --        cardtype_stname    C   1   短名
  --        prefix_text        C   1   前缀文本
  --        cardno_len         N   1   卡号长度
  --        default            N   1   缺省标志
  --        fixed              N   1   是否固定:1-是系统固定;0-不是系统固定
  --        strict             N   1   是否严格控制:1-是严格控制;0-不是严格控制
  --        self_make          N   1   是否自制:1-是的;0-不是
  --        exist_account      N   1   是否存在帐户:1-存在帐户;0-不存在账户
  --        allow_return_cash  N   1   是否退现:1-允许;0-不允许
  --        must_all_return    N   1   是否全退:1-必需全退;0-允许部分退
  --        component          C   1   部件
  --        memo               C   1   备注
  --        spec_item          C   1   特定项目
  --        blnc_mode          C   1   结算方式
  --        blnc_nature        N   1   结算性质
  --        cardno_pwdtxt      C   1   卡号密文:卡号从第几位至第几位显示密文,格式为:S-N:S表示从第几位开始,至第几位结束.比如:3-10,表示从3位到10位用密文*表示:12********3323主要是适应不同类别的医疗卡
  --        allow_repeat_use   N   1   是否重复使用:1-允许;0-不允许
  --        enabled            N   1   是否启用:1-已启用;0-未启用
  --        pwd_len            N   1   密码长度
  --        pwd_len_limit      N   1   密码长度限制:0-不作限制;1-固定密码长度;-n表示密码必须输入好多个位密码以上,但不能超过密码长度



  --        pwd_rule           N   1   密码规则:０-数字和字符组成;1-仅为数字组成
  --        allow_vaguefind    N   1   是否模糊查找:1-支持模糊查找;0-不支持
  --        pwd_require        N   1   密码输入限制:0-不限制;1-不输入,提醒;2-不输入禁止;缺省为不限制
  --        default_pwd        N   1   是否缺省密码:1-以身份证后N(以密码长度为准)位作为缺省密码;0-无缺省密码
  --        allow_makecard     N   1   是否制卡:1-是;0-否
  --        allow_sendcard     N   1   是否发卡:1-是;0-否
  --        allow_writcard     N   1   是否写卡:1-是;0-否
  --        insurance_type     N   1   险类
  --        insurance_name     C   1   险类名称
  --        sendcard_nature    N   1   发卡性质:0-不限制;1-同一病人只能发一张卡;2-同一病人允许发多张卡，但需提示;缺省为0
  --        allow_transfer     N   1   是否转帐及代扣:1-支持转帐及代扣;0-不支持
  --        readcard_nature    C   1   读卡性质,医疗卡读卡方式：第一位为:是否刷卡;第二位为是否扫描;第三位是否接触式读卡;第四位是否非接触式读卡。例如刷卡：'1000'
  --        keyboard_mode      N   1   键盘控制方式:：0-禁止使用软键盘;1-使用数字软键盘 ,2-使用字符软键盘
  --        advsend_buildqrcode N   1   是否医嘱发送调用条码生成接口:1-发送调用生成二维码接口;0-不调用
  --        holding_pay         N   1   是否持卡消费:1-是;0-否
  --        cert_cardtype       N   1   是否证件类型的医疗卡:0-不是；1-是
  --        verfycard           N   1   是否退款验卡
  --        sendcard_sign       N   1   发卡控制:0或NULL-发卡时，卡号必须达到卡号长度;1-发卡时，允许卡号小于等于卡号长度,发卡时，小于卡号长度时，不提示操作员;2-发卡时，允许卡号小于等于卡号长度,小于时，提示操作员。
  --        enterkey_enabled    N   1   设备是否启用回车:医疗卡对应的刷卡设备是否启用了回车，如果启用了回车，则卡号长度默认增加一位来屏蔽回车



  --        def_return_cash     N   1   是否缺省退现:允许退现时,默认是否退现
  --        balalone            N   1   是否独立结算:1-独立结算;0-非独立结算
  --        discern_rule        N   1   卡号识别规则:1-全部转换为大写;0-不区分大小写
  --        def_valid_time      C   1   缺省有效时间:NULL时，表示不限制;非空时，格式为:时间+单位(天，月),比如：3天,3月
  --        scanpay             N   1   是否支持扫码付:是否支持扫码付,支持时，会调用“zlReadQRCode部件”
  ---------------------------------------------------------------------------

  j_Json       Pljson;
  j_Jsonin     Pljson;
  v_Jvals      Varchar2(32767);
  n_卡类别id   医疗卡类别.Id%Type;
  n_查询类型   Number(2);
  n_是否证件   Number(2);
  n_是否有效卡 Number(2);
  v_卡名称     Varchar2(2000);
Begin
  --解析入参
  j_Jsonin     := Pljson(Json_In);
  j_Json       := j_Jsonin.Get_Pljson('input');
  n_卡类别id   := j_Json.Get_Number('cardtype_id');
  n_查询类型   := j_Json.Get_Number('query_type');
  n_是否证件   := j_Json.Get_String('cert_cardtype');
  n_是否有效卡 := j_Json.Get_Number('dffective_cardtype');
  v_卡名称     := j_Json.Get_String('cardtype_name');

  Json_Out := '{"output":{"code":1,"message":"成功","type_list":[';
  v_Jvals  := Null;

  For c_卡类别 In (Select a.Id, a.编码, a.名称, a.短名, a.前缀文本, a.卡号长度, a.缺省标志, a.是否固定, a.是否严格控制, a.是否自制, a.是否存在帐户, a.是否退现, a.是否全退,
                       a.部件, a.备注, a.特定项目, a.结算方式, a.卡号密文, a.是否重复使用, a.是否启用, a.密码长度, a.密码长度限制, a.密码规则, a.是否模糊查找, a.密码输入限制,
                       a.是否缺省密码, a.是否制卡, a.是否发卡, a.是否写卡, a.险类, a.发卡性质, a.是否转帐及代扣, a.读卡性质, a.键盘控制方式, a.发送调用接口, a.是否持卡消费,
                       a.是否证件, a.是否退款验卡, a.发卡控制, a.设备是否启用回车, a.是否缺省退现, a.是否独立结算, a.卡号识别规则, a.缺省有效时间, a.是否支持扫码付,
                       b.性质 As 结算性质, c.名称 As 险类名称
                From 医疗卡类别 A, 结算方式　b, 保险类别 C
                Where a.结算方式 = b.名称(+) And Decode(Nvl(n_卡类别id, 0), 0, 0, a.Id) = Nvl(n_卡类别id, 0) And
                      Decode(Nvl(n_是否证件, 0), 0, 0, Nvl(a.是否证件, 0)) = Nvl(n_是否证件, 0) And
                      Decode(Nvl(n_是否有效卡, 0), 0, 0, Nvl(a.是否启用, 0)) = Nvl(n_是否有效卡, 0) And Nvl(a.险类, 0) = c.序号(+) And
                      (v_卡名称 Is Null Or a.名称 = v_卡名称 Or a.特定项目 = v_卡名称)) Loop
  
    v_Jvals := v_Jvals || ',{"cardtype_id":' || c_卡类别.Id;
    v_Jvals := v_Jvals || ',"cardtype_code":"' || c_卡类别.编码 || '"';
    v_Jvals := v_Jvals || ',"cardtype_name":"' || c_卡类别.名称 || '"';
    v_Jvals := v_Jvals || ',"cardtype_stname":"' || c_卡类别.短名 || '"';
    v_Jvals := v_Jvals || ',"prefix_text":"' || c_卡类别.前缀文本 || '"';
    v_Jvals := v_Jvals || ',"cardno_len":' || Nvl(c_卡类别.卡号长度, 0);
    v_Jvals := v_Jvals || ',"default":' || Nvl(c_卡类别.缺省标志, 0);
    v_Jvals := v_Jvals || ',"fixed":' || Nvl(c_卡类别.是否固定, 0);
    v_Jvals := v_Jvals || ',"strict":' || Nvl(c_卡类别.是否严格控制, 0);
    v_Jvals := v_Jvals || ',"self_make":' || Nvl(c_卡类别.是否自制, 0);
    v_Jvals := v_Jvals || ',"exist_account":' || Nvl(c_卡类别.是否存在帐户, 0);
    v_Jvals := v_Jvals || ',"allow_return_cash":' || Nvl(c_卡类别.是否退现, 0);
    v_Jvals := v_Jvals || ',"must_all_return":' || Nvl(c_卡类别.是否全退, 0);
    v_Jvals := v_Jvals || ',"component":"' || c_卡类别.部件 || '"';
    v_Jvals := v_Jvals || ',"memo":"' || c_卡类别.备注 || '"';
    v_Jvals := v_Jvals || ',"spec_item":"' || c_卡类别.特定项目 || '"';
    v_Jvals := v_Jvals || ',"blnc_mode":"' || c_卡类别.结算方式 || '"';
    v_Jvals := v_Jvals || ',"blnc_nature":' || Nvl(c_卡类别.结算性质, 0);
    v_Jvals := v_Jvals || ',"cardno_pwdtxt":"' || c_卡类别.卡号密文 || '"';
    v_Jvals := v_Jvals || ',"allow_repeat_use":' || Nvl(c_卡类别.是否重复使用, 0);
    v_Jvals := v_Jvals || ',"enabled":' || Nvl(c_卡类别.是否启用, 0);
  
    If Nvl(n_查询类型, 0) = 0 Then
    
      --显示所有
      v_Jvals := v_Jvals || ',"pwd_len":' || Nvl(c_卡类别.密码长度, 0);
      v_Jvals := v_Jvals || ',"pwd_len_limit":' || Nvl(c_卡类别.密码长度限制, 0);
      v_Jvals := v_Jvals || ',"pwd_rule":' || Nvl(c_卡类别.密码规则, 0);
      v_Jvals := v_Jvals || ',"allow_vaguefind":' || Nvl(c_卡类别.是否模糊查找, 0);
      v_Jvals := v_Jvals || ',"pwd_require":' || Nvl(c_卡类别.密码输入限制, 0);
      v_Jvals := v_Jvals || ',"default_pwd":' || Nvl(c_卡类别.是否缺省密码, 0);
      v_Jvals := v_Jvals || ',"allow_makecard":' || Nvl(c_卡类别.是否制卡, 0);
      v_Jvals := v_Jvals || ',"allow_sendcard":' || Nvl(c_卡类别.是否发卡, 0);
      v_Jvals := v_Jvals || ',"allow_writecard":' || Nvl(c_卡类别.是否写卡, 0);
      v_Jvals := v_Jvals || ',"insurance_type":' || Nvl(c_卡类别.险类, 0);
      v_Jvals := v_Jvals || ',"insurance_name":"' || c_卡类别.险类名称 || '"';
      v_Jvals := v_Jvals || ',"sendcard_nature":' || Nvl(c_卡类别.发卡性质, 0);
      v_Jvals := v_Jvals || ',"allow_transfer":' || Nvl(c_卡类别.是否转帐及代扣, 0);
      v_Jvals := v_Jvals || ',"readcard_nature":"' || Nvl(c_卡类别.读卡性质, '1000') || '"';
      v_Jvals := v_Jvals || ',"keyboard_mode":' || Nvl(c_卡类别.键盘控制方式, 0);
      v_Jvals := v_Jvals || ',"advsend_buildqrcode":' || Nvl(c_卡类别.发送调用接口, 0);
      v_Jvals := v_Jvals || ',"holding_pay":' || Nvl(c_卡类别.是否持卡消费, 0);
      v_Jvals := v_Jvals || ',"cert_cardtype":' || Nvl(c_卡类别.是否证件, 0);
      v_Jvals := v_Jvals || ',"verfycard":' || Nvl(c_卡类别.是否退款验卡, 0);
      v_Jvals := v_Jvals || ',"sendcard_sign":' || Nvl(c_卡类别.发卡控制, 0);
      v_Jvals := v_Jvals || ',"enterkey_enabled":' || Nvl(c_卡类别.设备是否启用回车, 0);
      v_Jvals := v_Jvals || ',"def_return_cash":' || Nvl(c_卡类别.是否缺省退现, 0);
      v_Jvals := v_Jvals || ',"balalone":' || Nvl(c_卡类别.是否独立结算, 0);
      v_Jvals := v_Jvals || ',"discern_rule":' || Nvl(c_卡类别.卡号识别规则, 0);
      v_Jvals := v_Jvals || ',"def_valid_time":"' || c_卡类别.缺省有效时间 || '"';
      v_Jvals := v_Jvals || ',"scanpay":' || Nvl(c_卡类别.是否支持扫码付, 0);
    End If;
  
    v_Jvals := v_Jvals || '}';
    If Length(v_Jvals) > 30000 Then
      Json_Out := Json_Out || Substr(v_Jvals, 2);
      v_Jvals  := Null;
    End If;
  End Loop;
  If v_Jvals Is Not Null Then
    v_Jvals  := v_Jvals || ']}}';
    Json_Out := Json_Out || Substr(v_Jvals, 2);
  Else
    Json_Out := Json_Out || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getcardtypes;
/
Create Or Replace Procedure Zl_Patisvr_Getcommunityinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --功能：获取病人的社区信息
  --入参：Json_In:格式
  --input
  --  query_type        N    1 调用类型 ：1-通过病人id和社区名称获取社区号，2-获取社区和医保号[返回：community_code+insurance_num]
  --  pati_id           N    1 病人id
  --  community_id      N    1 社区id

  --出参      json
  --output
  --    code                N   1 应答吗：0-失败；1-成功
  --    message             C   1 应答消息：失败时返回具体的错误信息

  --    community_code      C   1 社区号【query_type=2】
  --    insurance_num       C   1 医保号【query_type=2】

  --    community_list[]社区信息列表，支持多个，[数组]【query_type=1】
  --       community_code    C   1 社区号
  -------------------------------------------
  j_Json   Pljson;
  j_Jsonin Pljson;
  v_List   Varchar2(32767);

  n_Type   Number(18);
  n_病人id Number(18);
  n_社区id Number(18);
  v_医保号 Varchar2(4000);
  v_社区号 Varchar2(4000);
Begin
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_Type   := j_Json.Get_Number('query_type');
  n_病人id := j_Json.Get_Number('pati_id');

  If n_Type = 2 Then
    n_社区id := j_Json.Get_String('community_id');
    Select Max(a.医保号) Into v_医保号 From 病人信息 A Where a.病人id = n_病人id;
    Select Max(a.社区号) Into v_社区号 From 病人社区信息 A Where a.病人id = n_病人id And a.社区 = n_社区id;
    Json_Out := '{"output":{"code":1,"message":"成功","community_code":"' || Zljsonstr(v_社区号) || '","insurance_num":"' ||
                Zljsonstr(v_医保号) || '"}}';
    Return;
  End If;

  If n_Type = 1 Then
  
    n_社区id := j_Json.Get_String('community_num');
  
    For c_社区信息 In (Select a.社区号 From 病人社区信息 A Where a.病人id = n_病人id And 社区 = n_社区id) Loop
    
      v_List := v_List || ',{"community_code":"' || Zljsonstr(c_社区信息.社区号) || '"}';
    
    End Loop;
  
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功","community_list":[' || Substr(v_List, 2) || ']}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getcommunityinfo;
/
Create Or Replace Procedure Zl_Patisvr_Getcustompatiinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:根据身份证号和基本信息来获取病人的详细信息(用户自定义返回)
  --入参：Json_In:格式
  --    input
  --      occasion          N 1 场合
  --      pati_idcard       C 1 身份证号
  --      pati_name         C   姓名
  --      pati_sex          C 1 性别
  --      query_type        N   查询类型：0-查询基本信息；1-只查询病人id
  --出参: Json_Out,格式如下
  --  output
  --    code              N   1   应答码：0-失败；1-成功
  --    message           C   1   应答消息：失败时返回具体的错误信息
  --    pati_ids          N   1   病人ids.查询类型为1时返回
  --    pati_count        N   1   病人信息条数.查询类型为0时返回
  --    pati_list         C       病人信息列表.查询类型为0时返回
  --      pati_id         N 1 病人id
  --      pati_pageid     N 1 主页id：病人信息.主页ID
  --      pati_name       C 1 姓名
  --      pati_sex        C 1 性别
  --      pati_age        C 1 年龄
  --      pati_birthdate  D 1 出生日期
  --      pati_nation     C 1 民族
  --      pati_idcard     C 1 身份证号
  --      pati_education  C   学历
  --      pati_identity   C   身份
  --      pati_marital_cstatus  C   婚姻状况
  --      pat_home_addr   C   家庭地址
  --      pati_area       C   区域
  --      pati_birthplace C   出生地点
  --      pati_emp_name   C   工作单位名称
  --      outpatient_num  C   门诊号
  --      inpatient_num   C   住院号
  --      insurance_num   C   医保号
  --      phone_number    C   联系电话(联系人电话，手机号，家庭电话三者取一)
  --      pati_bed        C   当前床号
  --      pati_type       C   病人类型(普通，医保，留观)
  --      out_date        D   出院日期
  --      create_time     C   登记时间:yyyy-mm-dd hh24:mi:ss
  ---------------------------------------------------------------------------
  j_Jsonput  Pljson;
  j_Json     Pljson;
  n_场合     Number(18);
  v_身份证号 病人信息.身份证号%Type;
  v_姓名     病人信息.姓名%Type;
  v_性别     病人信息.性别%Type;
  n_查询类型 Number(1);
  v_病人ids  Varchar2(32767);
  v_List     Varchar2(32767);
  n_Count    Number(10);

Begin
  --解析入参
  j_Jsonput  := Pljson(Json_In);
  j_Json     := j_Jsonput.Get_Pljson('input');
  n_场合     := j_Json.Get_Number('module');
  v_身份证号 := j_Json.Get_String('pati_idcard');
  v_姓名     := j_Json.Get_String('pati_name');
  v_性别     := j_Json.Get_String('pati_sex');
  n_Count    := 0;
  n_查询类型 := j_Json.Get_Number('query_type');
  If v_身份证号 Is Null And v_姓名 Is Null Then
    Json_Out := Zljsonout('未传入身份证号和姓名，请检查');
    Return;
  End If;
  v_病人ids := Zl_Custom_Patiids_Get(n_场合, v_身份证号, v_姓名, v_性别);
  If v_病人ids Is Null Then
    Json_Out := '{"output":{"code":1,"message":"成功","pati_count":0,"pati_list":[]}}';
    Return;
  End If;
  --n_查询类型：0-查询基本信息；1-只查询病人id
  If Nvl(n_查询类型, 0) = 0 Then
    Json_Out := '{"output":{"code":1,"message":"成功","pati_list":[';
    For c_病人信息 In (Select /*+cardinality(B,10)*/
                   Distinct a.病人id As ID, a.主页id, a.病人id, a.姓名, a.性别, a.年龄, a.出生日期, a.民族, a.身份证号, a.学历, a.身份, a.婚姻状况,
                            a.家庭地址, a.区域, a.出生地点, a.门诊号, a.住院号, a.医保号, Nvl(a.手机号, a.家庭电话) As 联系电话, a.工作单位, a.当前床号, a.病人类型,
                            a.出院时间, To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间
                   From 病人信息 A, Table(f_Str2list(v_病人ids)) B
                   Where a.病人id = b.Column_Value
                   Order By 姓名, 性别, 年龄) Loop
      n_Count := n_Count + 1;
      Zljsonputvalue(v_List, 'pati_id', c_病人信息.病人id, 1, 1);
      Zljsonputvalue(v_List, 'pati_pageid', c_病人信息.主页id, 1);
      Zljsonputvalue(v_List, 'pati_name', c_病人信息.姓名);
      Zljsonputvalue(v_List, 'pati_sex', c_病人信息.性别);
      Zljsonputvalue(v_List, 'pati_age', c_病人信息.年龄);
      Zljsonputvalue(v_List, 'pati_birthdate', To_Char(c_病人信息.出生日期, 'yyyy-mm-dd hh24:mi:ss'));
      Zljsonputvalue(v_List, 'pati_nation', c_病人信息.民族);
      Zljsonputvalue(v_List, 'pati_idcard', c_病人信息.身份证号);
      Zljsonputvalue(v_List, 'pati_education', c_病人信息.学历);
      Zljsonputvalue(v_List, 'pati_identity', c_病人信息.身份);
      Zljsonputvalue(v_List, 'pati_marital_cstatus', c_病人信息.婚姻状况);
      Zljsonputvalue(v_List, 'pat_home_addr', c_病人信息.家庭地址);
      Zljsonputvalue(v_List, 'pati_area', c_病人信息.区域);
      Zljsonputvalue(v_List, 'pati_birthplace', c_病人信息.出生地点);
      Zljsonputvalue(v_List, 'pati_emp_name', c_病人信息.工作单位);
      Zljsonputvalue(v_List, 'outpatient_num', c_病人信息.门诊号, 0);
      Zljsonputvalue(v_List, 'inpatient_num', c_病人信息.住院号, 0);
      Zljsonputvalue(v_List, 'insurance_num', c_病人信息.医保号);
      Zljsonputvalue(v_List, 'phone_number', c_病人信息.联系电话);
      Zljsonputvalue(v_List, 'pati_bed', c_病人信息.当前床号);
      Zljsonputvalue(v_List, 'pati_type', c_病人信息.病人类型);
      Zljsonputvalue(v_List, 'out_date', To_Char(c_病人信息.出院时间, 'yyyy-mm-dd hh24:mi:ss'));
      Zljsonputvalue(v_List, 'create_time', To_Char(c_病人信息.登记时间, 'yyyy-mm-dd hh24:mi:ss'), 0, 2);
      If Length(v_List) > 20000 Then
        Json_Out := Json_Out || v_List;
        v_List   := ',';
      End If;
    End Loop;
    If v_List <> ',' Then
      Json_Out := Json_Out || v_List || '],"pati_count":' || n_Count || '}}';
    End If;
    Return;
  Else
    Json_Out := '{"output":{"code":1,"message":"成功","pati_ids":"' || v_病人ids || '"}}';
    Return;
  End If;

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getcustompatiinfo;
/
Create Or Replace Procedure Zl_Patisvr_Getinputitemlength
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:获取输入项目的实际大小
  --入参：Json_In:格式
  --    input
  --    item_list[]
  --      table_name  C 1 表名
  --      column_name C 1 列名,多个用逗号
  --出参: Json_Out,格式如下
  --  output
  --    code              N   1   应答码：0-失败；1-成功
  --    message           C   1   应答消息：失败时返回具体的错误信息
  --    item_list[] C
  --      table_name  C 1 表名
  --      column_name C 1 列表
  --      column_size N 1 长度

  ---------------------------------------------------------------------------
  j_Jsonin   Pljson;
  j_Json     Pljson;
  o_Json     Pljson;
  j_Jsonlist Pljson_List := Pljson_List();

  v_表名 Varchar2(100);
  v_字段 Varchar2(32767);
  v_Jtmp Varchar2(32767);
Begin
  --解析入参
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  j_Jsonlist := j_Json.Get_Pljson_List('item_list');

  If j_Jsonlist Is Null Then
    Json_Out := Zljsonout('未传入需要查询的表信息，请检查!');
    Return;
  End If;

  Json_Out := '{"output":{"code":1,"message":"成功","item_list":[';
  For I In 1 .. j_Jsonlist.Count Loop
    o_Json := Pljson();
    o_Json := Pljson(j_Jsonlist.Get(I));
    v_表名 := o_Json.Get_String('table_name');
    v_字段 := o_Json.Get_String('column_name');
    For c_列信息 In (Select Column_Name As 列名, Max(Data_Length) As 长度
                  From User_Tab_Columns
                  Where Table_Name = v_表名 And Instr(',' || v_字段 || ',', ',' || Column_Name || ',') > 0
                  Group By Column_Name) Loop
      v_Jtmp := v_Jtmp || ',{"table_name":"' || Zljsonstr(v_表名) || '"';
      v_Jtmp := v_Jtmp || ',"column_name":"' || Zljsonstr(c_列信息.列名) || '"';
      v_Jtmp := v_Jtmp || ',"column_size":' || Zljsonstr(c_列信息.长度, 1);
      v_Jtmp := v_Jtmp || '}';
    
    End Loop;
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getinputitemlength;
/
Create Or Replace Procedure Zl_Patisvr_Getinsureaccbalance
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：获取病人帐户信息（医保帐户， 适用病人类型)
  --入参：Json_In:格式
  --input
  --  pati_id               N  1  病人ID
  --  insurance_type        N  0   险类
  --出参      json
  --output
  --  code                  C  1  应答码：0-失败；1-成功
  --  message               C  1  应答消息：
  --  pati_type             C  0  适用病人类型
  --  insure_srpls_chrg     N  0  医保账户余额
  ---------------------------------------------------------------------------
  j_Input  Pljson;
  j_Jsonin Pljson;

  n_病人id   Number(18);
  n_险类     Number(18);
  v_适用病人 Varchar2(200);
  n_帐户余额 医保病人档案.帐户余额%Type;

Begin
  j_Jsonin := Pljson(Json_In);
  j_Input  := j_Jsonin.Get_Pljson('input');
  n_病人id := j_Input.Get_Number('pati_id');
  n_险类   := j_Input.Get_Number('insurance_type');

  Select Decode(Nvl(n_险类, 0), 0, Max(险类), n_险类), Max(Zl_Patiwarnscheme(病人id))
  Into n_险类, v_适用病人
  From 病人信息
  Where 病人id = n_病人id;

  Select Nvl(Max(e.帐户余额), 0)
  Into n_帐户余额
  From 医保病人关联表 D, 医保病人档案 E
  Where d.病人id = n_病人id And d.险类 = Nvl(n_险类, 0) And d.险类 = e.险类(+) And d.医保号 = e.医保号(+) And d.标志(+) = 1;

  Json_Out := '{"output":{"code":1,"message":"成功","pati_type":"' || Zljsonstr(v_适用病人) || '","insure_srpls_chrg":' ||
              Zljsonstr(n_帐户余额, 1) || '}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getinsureaccbalance;
/
Create Or Replace Procedure Zl_Patisvr_Getinsureinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:根据病人id，获取病人的保险信息
  --入参：Json_In:格式
  --  input
  --    pati_id             N   1 病人id
  --    insure_type         N   1 险类
  --出参: Json_Out,格式如下
  --  output
  --    code                N   1 应答码：0-失败；1-成功
  --    message             C   1 应答消息：失败时返回具体的错误信息
  --    insure_type         C   1 险类
  --    insure_name         C   1 险类名称
  --    insure_no           C   1 医保号
  --    card_no             C   1 卡号
  --    pati_create_time    C   1 病人的登记时间:yyyy-mm-dd hh24:mi:ss
  --    insure_pwd          C   1 医保密码
  --    dz_type_id          N   1 病种id
  ---------------------------------------------------------------------------

  j_Json     Pljson;
  j_Jsonin   Pljson;
  n_病人id   病人信息.病人id%Type;
  n_险类     病人信息.险类%Type;
  v_保险名称 保险类别.名称%Type;
  v_医保号   医保病人档案.医保号%Type;
  v_卡号     医保病人档案.卡号%Type;
  d_登记时间 医保病人档案.就诊时间%Type;
  v_密码     医保病人档案.密码%Type;
  n_病种id   医保病人档案.病种id%Type;
  n_Type     Number(1);
  v_Jtmp     Varchar2(32767);
Begin
  --解析入参
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');
  n_险类   := j_Json.Get_Number('insure_type');
  If Nvl(n_病人id, 0) = 0 Then
    Json_Out := Zljsonout('失败，未传入病人id！');
    Return;
  End If;
  If Nvl(n_险类, 0) = 0 Then
    Select Max(险类) Into n_险类 From 病人信息 Where 病人id = n_病人id;
  Else
    n_Type := 1;
    Select Max(b.名称), Max(a.医保号), Max(a.就诊时间), Max(a.密码), Max(a.卡号), Max(a.病种id)
    Into v_保险名称, v_医保号, d_登记时间, v_密码, v_卡号, n_病种id
    From 保险帐户 A, 保险类别 B
    Where a.险类 = b.序号 And a.病人id = n_病人id And a.险类 = n_险类;
  End If;

  If Nvl(n_Type, 0) <> 0 Then
  
    v_Jtmp := v_Jtmp || ',"insure_type":' || Nvl(n_险类 || '', 'null');
    v_Jtmp := v_Jtmp || ',"insure_name":"' || Zljsonstr(v_保险名称) || '"';
    v_Jtmp := v_Jtmp || ',"insure_no":"' || Zljsonstr(v_医保号) || '"';
    v_Jtmp := v_Jtmp || ',"card_no":"' || Zljsonstr(v_卡号) || '"';
    v_Jtmp := v_Jtmp || ',"pati_create_time":"' || To_Char(d_登记时间, 'yyyy-mm-dd hh24:mi:ss') || '"';
    v_Jtmp := v_Jtmp || ',"insure_pwd":"' || Zljsonstr(v_密码) || '"';
    v_Jtmp := v_Jtmp || ',"dz_type_id":' || Nvl(n_病种id || '', 'null');
  
    Json_Out := '{"output":{"code":1,"message":"成功"' || v_Jtmp || '}}';
  Else
    Json_Out := '{"output":{"code":1,"message":"成功","insure_type":' || Nvl(n_险类 || '', 'null') || '}}';
  End If;

Exception
  When Others Then
    Json_Out := '{"output": {"code": 0,"message": "' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getinsureinfo;
/
Create Or Replace Procedure Zl_Patisvr_Getlastblackinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  --------------------------------------------------------------------------------------------------
  --功能:获取最后一次的不良记录的信息
  --入参 JSOM格式
  --input
  --  pati_id        N 1 病人id
  --出参 JSON格式
  --output
  --  code            N 1 应答码：0-失败；1-成功
  --  message         C 1 应答消息：失败时返回具体的错误信息
  --  pati_id         N 1 病人id
  --  last_time       C 1 最后预约时间
  --------------------------------------------------------------------------------------------------
  j_Json         Pljson;
  j_Jsonin       Pljson;
  n_病人id       Number;
  v_最后预约时间 Varchar2(30);
Begin
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');
  Begin
    Select Max(加入时间) As 最后预约时间
    Into v_最后预约时间
    From 病人不良记录
    Where 病人id = n_病人id And (行为类别 = '预约超期' Or (加入原因 = '预约失约次数过多,自动进入黑名单' And 行为类别 = '其他'));
  Exception
    When Others Then
      Null;
  End;
  Json_Out := '{"output":{"pati_id":' || n_病人id || ',"last_time":"' || To_Char(v_最后预约时间, 'yyyy-mm-dd hh24:mi:ss') ||
              '"}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getlastblackinfo;
/
Create Or Replace Procedure Zl_Patisvr_Getnextid
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

  v_Table     Varchar2(500);
  v_Col       Varchar2(500);
  n_Nextid    Number;
  j_Json      Pljson;
  j_Jsoninput Pljson;

Begin
  --解析入参
  j_Jsoninput := Pljson(Json_In);
  j_Json      := j_Jsoninput.Get_Pljson('input');
  v_Table     := j_Json.Get_String('table_name');
  v_Col       := Nvl(j_Json.Get_String('col_name'), 'ID');
  Execute Immediate 'select ' || v_Table || '_' || v_Col || '.nextval from dual'
    Into n_Nextid;
  Json_Out := '{"output":{"code":1,"message":"成功","next_id":' || n_Nextid || '}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Getnextid;
/
Create Or Replace Procedure Zl_Patisvr_Getnextno
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：功能：根据特定规则产生新的号码
  --入参：Json_In:格式
  --  input
  --    item_num            N   1   项目序号
  --    dept_id             N   0   科室ID
  --出参: Json_Out,格式如下
  --  output
  --    code                N   1   应答码：0-失败；1-成功
  --    message             C   1   应答消息：失败时返回具体的错误信息
  --    next_no             C   1   下一个号码
  ---------------------------------------------------------------------------
  j_Json      Pljson;
  j_Jsoninput Pljson;
  v_No        Varchar2(64);
  n_序号      Number(10);
  n_科室id    Number(18);
Begin
  j_Jsoninput := Pljson(Json_In);
  j_Json      := j_Jsoninput.Get_Pljson('input');
  n_序号      := j_Json.Get_Number('item_num');
  n_科室id    := j_Json.Get_Number('dept_id');

  Select Zl_Pati_Nextno(n_序号, n_科室id) Into v_No From Dual;
  Json_Out := '{"output":{"code":1,"message":"成功","next_no":"' || v_No || '"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Getnextno;

/
Create Or Replace Procedure Zl_Patisvr_Getpatallergicdrugs
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:获取病人信息的过敏药物信息
  --入参：Json_In:格式
  --  input
  --    pati_id             N   1 病人id
  --出参: Json_Out,格式如下
  --  output
  --    code                N   1   应答码：0-失败；1-成功
  --    message             C   1   应答消息：失败时返回具体的错误信息
  --    drug_list[]         C       过敏药物列表
  --      medicinal_id      N   1   过敏药品ID
  --      medicinal_name    C   1   过敏药物名称
  --      allergy_info      C   1   过每药物反应
  ---------------------------------------------------------------------------

  j_Json   Pljson;
  j_Jsonin Pljson;
  n_病人id 病人过敏药物.病人id%Type;
  v_List   Varchar2(32767);
  v_Jtmp   Varchar2(32767);
  c_Jtmp   Clob;
Begin
  --解析入参
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');

  If Nvl(n_病人id, 0) = 0 Then
    Json_Out := Zljsonout('未传入病人id，请检查!');
    Return;
  End If;

  For r_过敏记录 In (Select Distinct 过敏药物id, 过敏药物, 过敏反应 From 病人过敏药物 Where 病人id = n_病人id) Loop
  
    v_Jtmp := v_Jtmp || ',{"medicinal_id":' || Nvl(r_过敏记录.过敏药物id || '', 'null');
    v_Jtmp := v_Jtmp || ',"medicinal_name":"' || Zljsonstr(r_过敏记录.过敏药物) || '"';
    v_Jtmp := v_Jtmp || ',"allergy_info":"' || Zljsonstr(r_过敏记录.过敏反应) || '"';
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
    Json_Out := '{"output":{"code":1,"message":"成功","drug_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","drug_list":[' || c_Jtmp || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getpatallergicdrugs;
/
Create Or Replace Procedure Zl_Patisvr_Getpatiaddrssinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：根据病人ID,获取病人的地址信息
  --入参：Json_In:格式
  --  input
  --    pati_id              N   1  病人ID
  --    pati_pageid          N   0  主页id
  --    addr_type            N   0  地址类别:1-出生地，2-籍贯,3-现住址,4-户口地址,5-联系人地址，6-单位地址;为0时表示查询所有类型的地址信息
  --    addr_types           C   0  地址类别s:1-出生地，2-籍贯,3-现住址,4-户口地址,5-联系人地址，6-单位地址;多个类别,用","分隔.传来了该节点时则addr_type无效
  --出参: Json_Out,格式如下
  --  output
  --    code                 N   1   应答吗：0-失败；1-成功
  --    message              C   1   应答消息：失败时返回具体的错误信息
  --    addr_list[]          C       地址列表信息
  --      pat_addr_type      C   1   地址类别
  --      pat_addr_state     C   1   地址_省
  --      pat_addr_city      C   1   地址_市
  --      pat_addr_county    C   1   地址_县
  --      pat_addr_township  C   1   地址_乡
  --      pat_addr_other     C   1   地址_其他
  --      pat_region_code    C   1   区划代码
  ---------------------------------------------------------------------------
  n_病人id    病人信息.病人id%Type;
  n_主页id    病人信息.主页id%Type;
  n_地址类别  病人地址信息.地址类别%Type;
  v_地址类别s Varchar2(3000);
  v_List      Varchar2(32767);
  j_Json      Pljson;
  j_Jsonin    Pljson;
Begin

  --解析入参
  j_Jsonin    := Pljson(Json_In);
  j_Json      := j_Jsonin.Get_Pljson('input');
  n_病人id    := j_Json.Get_Number('pati_id');
  n_主页id    := Nvl(j_Json.Get_Number('pati_pageid'), 0);
  n_地址类别  := Nvl(j_Json.Get_Number('addr_type'), 0);
  v_地址类别s := Nvl(j_Json.Get_String('addr_type'), '-');
  If Nvl(n_病人id, 0) = 0 Then
    Json_Out := Zljsonout('未传入病人ID');
    Return;
  End If;
  If v_地址类别s <> '-' Then
    For r_地址 In (Select 地址类别, 省, 市, 县, 乡镇, 其他, 区划代码
                 From 病人地址信息
                 Where 病人id = n_病人id And Nvl(主页id, 0) = Nvl(n_主页id, 0) And
                       Instr(',' || v_地址类别s || ',', ',' || 地址类别 || ',') > 0) Loop
      --      pat_addr_type      C   1 地址类别
      --      pat_addr_state     C   1 地址_省
      --      pat_addr_city      C   1 地址_市
      --      pat_addr_county    C   1 地址_县
      --      pat_addr_township  C   1 地址_乡
      --      pat_addr_other     C   1 地址_其他
      --      pat_region_code    C   1 区划代码
      Zljsonputvalue(v_List, 'pat_addr_type', r_地址.地址类别, 0, 1);
      Zljsonputvalue(v_List, 'pat_addr_state', r_地址.省);
      Zljsonputvalue(v_List, 'pat_addr_city', r_地址.市);
      Zljsonputvalue(v_List, 'pat_addr_county', r_地址.县);
      Zljsonputvalue(v_List, 'pat_addr_township', r_地址.乡镇);
      Zljsonputvalue(v_List, 'pat_addr_other', r_地址.其他);
      Zljsonputvalue(v_List, 'pat_region_code', r_地址.区划代码, 0, 2);
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"成功","addr_list":[' || v_List || ']}}';
    Return;
  End If;
  For r_地址 In (Select 地址类别, 省, 市, 县, 乡镇, 其他, 区划代码
               From 病人地址信息
               Where 病人id = n_病人id And Nvl(主页id, 0) = Nvl(n_主页id, 0) And
                     ((地址类别 = n_地址类别 And Nvl(n_地址类别, 0) <> 0) Or Nvl(n_地址类别, 0) = 0)) Loop
    --      pat_addr_type      C   1 地址类别
    --      pat_addr_state     C   1 地址_省
    --      pat_addr_city      C   1 地址_市
    --      pat_addr_county    C   1 地址_县
    --      pat_addr_township  C   1 地址_乡
    --      pat_addr_other     C   1 地址_其他
    --      pat_region_code    C   1 区划代码
    Zljsonputvalue(v_List, 'pat_addr_type', r_地址.地址类别, 0, 1);
    Zljsonputvalue(v_List, 'pat_addr_state', r_地址.省);
    Zljsonputvalue(v_List, 'pat_addr_city', r_地址.市);
    Zljsonputvalue(v_List, 'pat_addr_county', r_地址.县);
    Zljsonputvalue(v_List, 'pat_addr_township', r_地址.乡镇);
    Zljsonputvalue(v_List, 'pat_addr_other', r_地址.其他);
    Zljsonputvalue(v_List, 'pat_region_code', r_地址.区划代码, 0, 2);
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"成功","addr_list":[' || v_List || ']}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getpatiaddrssinfo;
/
Create Or Replace Procedure Zl_Patisvr_Getpatiblackinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:获取病人黑名单信息
  --入参：Json_In:格式
  --    input
  --      pati_id           N 1 病人id
  --      occasion          C 1 应用场合:预约，挂号，结帐，入院，出院
  --      appt_mode_name    C 0 预约方式

  --出参: Json_Out,格式如下
  --  output
  --    code              N   1   应答码：0-失败；1-成功
  --    message           C   1   应答消息：失败时返回具体的错误信息
  --    tip_mode          N   1   控制方式：1-禁止;2-提示(或询问)
  --    tip_message       C   1   提示的信息
  ---------------------------------------------------------------------------

  j_Json        Pljson;
  j_Jsonin      Pljson;
  n_病人id      病人不良记录.病人id%Type;
  v_应用场合    不良行为控制.应用场合%Type;
  v_预约方式    不良行为控制.预约方式%Type;
  n_控制方式    Number(1);
  v_Message     Varchar2(30000);
  v_Black_Infor Varchar2(32767);
  v_Tmp         Varchar2(30000);

Begin
  --解析入参
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_病人id   := j_Json.Get_Number('pati_id');
  v_应用场合 := j_Json.Get_String('occasion');
  v_预约方式 := j_Json.Get_String('appt_mode_name');

  v_Black_Infor := Zl_Fun_Getblacklistinfor(n_病人id, v_应用场合, v_预约方式);

  If Nvl(v_Black_Infor, '-') <> '-' Then
    v_Tmp      := Substr(v_Black_Infor, 1, 1);
    n_控制方式 := To_Number(v_Tmp);
    v_Message  := Substr(v_Black_Infor, 3);
  End If;

  Json_Out := '{"output":{"code":1,"message":"成功","tip_mode":' || Nvl(n_控制方式, 0) || ',"tip_message":"' ||
              Zljsonstr(v_Message) || '"}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getpatiblackinfo;
/
Create Or Replace Procedure Zl_Patisvr_Getpaticardinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:获取病人信息
  --入参：Json_In:格式
  --    input
  --      pati_ids            C  1  病人ids,多个用","分隔
  --      cardtype_ids        C  1  卡类别IDs,多个用逗号分离
  --      card_no             C  0  卡号
  --      card_name           C  0  医疗卡名称
  --      cert_cardtype       N  1  只读取证件作为卡类别的医疗卡类别:1-只读取是否证件=1的医疗卡;0-全部读取
  --      query_type          N  1  查询基本类型:0-只获取病人ID,1-只获取卡类别ID;2-包含病人基本信息;3-所有
  --      dffective_cardtype  N  0  只读取有效的卡类别:1-只读取有效的卡类别;0-全部读取

  -- 出参：json
  --  output
  --  code                    N  1  应答码：0-失败；1-成功
  --  message                 C  1  应答消息：成功时返回成功信息，失败时返回具体的错误信息
  --  card_list[]             C     病人卡信息列表
  --    pati_id               N  1  病人id
  --    pati_name             C  1  姓名
  --    pati_sex              C  1  性别
  --    pati_age              C  1  年龄
  --    pati_birthdate        C  1  出生日期：yyyy-mm-dd hh24:mi:ss
  --    outpatient_num        C  1  门诊号
  --    pati_idcard           C  1  身份证号
  --    cardtype_id           N  1  卡类别ID
  --    card_no               C  1  卡号
  --    card_qrcode           C  1  二维码
  --    card_passwod          C  1  密码
  --    cardtype_name         C  1  卡类别名称
  --    cardtype_cardlen      N  1  卡号长度
  --    card_statu            N  1  状态:0-正常有效卡;1-已挂失; 2-补卡停用
  --    loscard_creator       C  1  挂失人
  --    loscard_time          C  1  挂失时间:yyyy-mm-dd hh24:mi:ss
  --    loscard_mode          C  1  挂失方式
  --    loscard_days          N  1  挂失天数
  --    sendcard_oper         C  1  发卡人
  --    end_time              C  1  终止使用时间:yyyy-mm-dd hh24:mi:ss
  ---------------------------------------------------------------------------
  v_Err_Msg Varchar2(500);
  j_Json    Pljson;
  j_Jsonin  Pljson;

  j_Json_Tmp   Pljson;
  v_病人ids    Varchar2(3000);
  v_List       Varchar2(32767);
  v_Tmp        Varchar2(32767);
  v_卡类别ids  Varchar2(32680);
  v_卡号       Varchar2(1000);
  v_卡名称     Varchar2(3000);
  n_查询类型   Number(2);
  n_是否证件   Number(2);
  n_是否有效卡 Number(2);
  n_卡类别id   病人医疗卡信息.卡类别id%Type;

  Cursor c_病人基本信息 Is
    Select a.病人id, a.卡类别id, a.卡类别名称, a.卡号, a.卡号长度, a.密码, a.状态, To_Char(a.挂失时间, 'yyyy-mm-dd hh24:mi:ss') As 挂失时间, a.挂失方式,
           a.挂失人, To_Char(a.发卡日期, 'yyyy-mm-dd hh24:mi:ss') As 发卡日期, a.发卡人,
           To_Char(a.终止使用时间, 'yyyy-mm-dd hh24:mi:ss') As 终止使用时间, a.二维码, b.姓名, b.性别, b.年龄,
           To_Char(b.出生日期, 'yyyy-mm-dd hh24:mi:ss') As 出生日期, b.身份证号, b.门诊号, a.有效天数
    From (Select a.病人id, a.卡类别id, a.卡号, a.密码, a.状态, a.挂失时间, a.挂失方式, a.挂失人, a.发卡日期, a.发卡人, a.终止使用时间, a.二维码, q.名称 As 卡类别名称,
                  q.卡号长度, m.有效天数
           From 病人医疗卡信息 A, 医疗卡挂失方式 M, 医疗卡类别 Q
           Where a.挂失方式 = m.名称(+) And a.卡类别id = q.Id And Nvl(q.是否启用, 0) = 1 And a.病人id = 0) A, 病人信息 B
    Where a.病人id = b.病人id And Rownum < 1;
  r_病人 c_病人基本信息%RowType;

  Type Ty_病人信息 Is Ref Cursor;
  c_病人信息 Ty_病人信息; --动态游标变量

Begin
  --解析入参
  j_Jsonin     := Pljson(Json_In);
  j_Json       := j_Jsonin.Get_Pljson('input');
  v_病人ids    := j_Json.Get_String('pati_ids');
  v_卡类别ids  := j_Json.Get_String('cardtype_ids');
  n_查询类型   := j_Json.Get_Number('query_type');
  n_是否证件   := j_Json.Get_Number('cert_cardtype');
  v_卡号       := j_Json.Get_String('card_no');
  n_是否有效卡 := j_Json.Get_Number(' dffective_cardtype');
  v_卡名称     := j_Json.Get_String('card_name');

  If Nvl(v_病人ids, '-') = '-' And v_卡号 Is Null Then
    v_Err_Msg := '未传入有效的查询条件，不能获取病人所持有的卡类别!';
    Json_Out  := Zljsonout(v_Err_Msg);
    Return;
  End If;

  If Not v_卡名称 Is Null Then
    Select Nvl(Max(ID), 0) Into n_卡类别id From 医疗卡类别 Where 名称 = v_卡名称;
  End If;

  If Nvl(v_病人ids, '-') <> '-' Then
    Open c_病人信息 For
      Select a.病人id, a.卡类别id, a.卡类别名称, a.卡号, a.卡号长度, a.密码, a.状态, To_Char(a.挂失时间, 'yyyy-mm-dd hh24:mi:ss') As 挂失时间,
             a.挂失方式, a.挂失人, To_Char(a.发卡日期, 'yyyy-mm-dd hh24:mi:ss') As 发卡日期, a.发卡人,
             To_Char(a.终止使用时间, 'yyyy-mm-dd hh24:mi:ss') As 终止使用时间, a.二维码, b.姓名, b.性别, b.年龄,
             To_Char(b.出生日期, 'yyyy-mm-dd hh24:mi:ss') As 出生日期, b.身份证号, b.门诊号, a.有效天数
      From (Select a.病人id, a.卡类别id, a.卡号, a.密码, a.状态, a.挂失时间, a.挂失方式, a.挂失人, a.发卡日期, a.发卡人, a.终止使用时间, a.二维码,
                    q.名称 As 卡类别名称, q.卡号长度, m.有效天数,
                    Case
                       When Nvl(a.状态, 0) = 1 And Nvl(n_是否有效卡, 0) = 1 And
                            (Nvl(m.有效天数, 0) = 0 Or Nvl(a.挂失时间, To_Date('3000-01-01', 'yyyy-mm-dd')) + Nvl(m.有效天数, 0) > Sysdate) Then
                        1
                       Else
                        Nvl(a.状态, 0)
                     End As 状态1
             From 病人医疗卡信息 A, 医疗卡挂失方式 M, 医疗卡类别 Q
             Where a.挂失方式 = m.名称(+) And a.卡类别id = q.Id And Nvl(q.是否启用, 0) = 1 And
                   a.病人id In (Select /*+cardinality(B,10) */
                               Column_Value As 病人id
                              From Table(f_Str2list(v_病人ids)) B) And
                   (v_卡类别ids Is Null Or Instr(',' || v_卡类别ids || ',', ',' || a.卡类别id || ',') > 0) And
                   Decode(Nvl(n_是否证件, 0), 0, 0, Nvl(q.是否证件, 0)) = Nvl(n_是否证件, 0) And
                   (v_卡号 Is Null Or a.卡号 = Nvl(v_卡号, '-'))) A, 病人信息 B
      Where a.病人id = b.病人id And a.状态1 = 0;
  Elsif v_卡号 Is Not Null Then
  
    Open c_病人信息 For
      Select a.病人id, a.卡类别id, a.卡类别名称, a.卡号, a.卡号长度, a.密码, a.状态, To_Char(a.挂失时间, 'yyyy-mm-dd hh24:mi:ss') As 挂失时间,
             a.挂失方式, a.挂失人, To_Char(a.发卡日期, 'yyyy-mm-dd hh24:mi:ss') As 发卡日期, a.发卡人,
             To_Char(a.终止使用时间, 'yyyy-mm-dd hh24:mi:ss') As 终止使用时间, a.二维码, b.姓名, b.性别, b.年龄,
             To_Char(b.出生日期, 'yyyy-mm-dd hh24:mi:ss') As 出生日期, b.身份证号, b.门诊号, a.有效天数
      From (Select a.病人id, a.卡类别id, a.卡号, a.密码, a.状态, a.挂失时间, a.挂失方式, a.挂失人, a.发卡日期, a.发卡人, a.终止使用时间, a.二维码,
                    q.名称 As 卡类别名称, q.卡号长度, m.有效天数,
                    Case
                       When Nvl(a.状态, 0) = 1 And Nvl(n_是否有效卡, 0) = 1 And
                            (Nvl(m.有效天数, 0) = 0 Or Nvl(a.挂失时间, To_Date('3000-01-01', 'yyyy-mm-dd')) + Nvl(m.有效天数, 0) > Sysdate) Then
                        1
                       Else
                        Nvl(a.状态, 0)
                     End As 状态1
             From 病人医疗卡信息 A, 医疗卡挂失方式 M, 医疗卡类别 Q
             Where a.挂失方式 = m.名称(+) And a.卡类别id = q.Id And Nvl(q.是否启用, 0) = 1 And a.卡号 = v_卡号 And
                   (v_卡类别ids Is Null Or Instr(',' || v_卡类别ids || ',', ',' || a.卡类别id || ',') > 0) And
                   (a.卡类别id = n_卡类别id Or n_卡类别id = 0) And Decode(Nvl(n_是否证件, 0), 0, 0, Nvl(q.是否证件, 0)) = Nvl(n_是否证件, 0)) A,
           病人信息 B
      Where a.病人id = b.病人id And a.状态1 = 0;
  
  Else
    v_Err_Msg := '未传入有效的查询条件，不能获取病人信息!';
    Json_Out  := Zljsonout(v_Err_Msg);
    Return;
  End If;

  Json_Out := '{"output":{"code":1,"message":"成功","card_list":[';

  Loop
    Fetch c_病人信息
      Into r_病人;
    Exit When c_病人信息%NotFound;
  
    j_Json_Tmp := Pljson();
    --1.取基本信息
    --0-只获取病人ID,1-只获取卡类别ID;2-包含病人基本信息;3-所有
    If Nvl(n_查询类型, 0) <> 1 Then
      v_Tmp := v_Tmp || ',{"pati_id":' || Nvl(r_病人.病人id, 0);
      If Nvl(n_查询类型, 0) = 0 Then
        v_Tmp := v_Tmp || '}';
      End If;
      If Nvl(n_查询类型, 0) <> 0 Then
        v_Tmp := v_Tmp || ',"pati_name":"' || Nvl(r_病人.姓名, '') || '"';
        v_Tmp := v_Tmp || ',"pati_sex":"' || Nvl(r_病人.性别, '') || '"';
        v_Tmp := v_Tmp || ',"pati_age":"' || Nvl(r_病人.年龄, '') || '"';
        v_Tmp := v_Tmp || ',"pati_birthdate":"' || Nvl(r_病人.出生日期, '') || '"';
        v_Tmp := v_Tmp || ',"outpatient_num":"' || Zljsonstr(r_病人.门诊号) || '"';
        v_Tmp := v_Tmp || ',"pati_idcard":"' || Nvl(r_病人.身份证号, '') || '"';
        v_Tmp := v_Tmp || ',"cardtype_id":' || Nvl(r_病人.卡类别id, 0);
        v_Tmp := v_Tmp || ',"card_no":"' || Nvl(r_病人.卡号, '') || '"';
        v_Tmp := v_Tmp || ',"card_qrcode":"' || Nvl(r_病人.二维码, '') || '"';
        v_Tmp := v_Tmp || ',"card_passwod":"' || Nvl(r_病人.密码, '') || '"';
        v_Tmp := v_Tmp || ',"cardtype_name":"' || Nvl(r_病人.卡类别名称, '') || '"';
        If Nvl(n_查询类型, 0) = 2 Then
          v_Tmp := v_Tmp || '}';
        End If;
        If Nvl(n_查询类型, 0) <> 2 Then
          v_Tmp := v_Tmp || ',"cardtype_cardlen":' || Nvl(r_病人.卡号长度, 0);
          v_Tmp := v_Tmp || ',"card_statu":' || Nvl(r_病人.状态, 0);
          v_Tmp := v_Tmp || ',"loscard_creator":"' || Nvl(r_病人.挂失人, '') || '"';
          v_Tmp := v_Tmp || ',"loscard_time":"' || Nvl(r_病人.挂失时间, '') || '"';
          v_Tmp := v_Tmp || ',"loscard_days":' || Nvl(r_病人.有效天数, 0);
          v_Tmp := v_Tmp || ',"loscard_mode":"' || Nvl(r_病人.挂失方式, '') || '"';
          v_Tmp := v_Tmp || ',"sendcard_oper":"' || Nvl(r_病人.发卡人, '') || '"';
          v_Tmp := v_Tmp || ',"end_time":"' || Nvl(r_病人.终止使用时间, '') || '"}';
        End If;
      End If;
      If Length(v_Tmp) > 20000 Then
        Json_Out := Json_Out || v_Tmp;
        v_Tmp    := ',';
      End If;
    Else
      v_List := v_List || ',{"cardtype_id":' || Nvl(r_病人.卡类别id, 0) || '}';
    End If;
  End Loop;
  If Nvl(n_查询类型, 0) = 1 Then
    v_List   := Substr(v_List, 2);
    Json_Out := Json_Out || v_List || ']}}';
  Else
    If v_Tmp = ',' Then
      Json_Out := Json_Out || ']}}';
    Else
      v_Tmp    := Substr(v_Tmp, 2);
      Json_Out := Json_Out || v_Tmp || ']}}';
    End If;
  End If;

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getpaticardinfo;
/
Create Or Replace Procedure Zl_Patisvr_Getpaticardno
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：根据病人id和缺省卡类别获取病人的卡信息
  --入参：Json_In:格式
  --  input
  --   pati_ids            C 1  病人ids,多个病人以逗号分割
  --   card_type_id        N 1 卡类别id
  --出参: Json_Out,格式如下
  --  output
  --    code              N 1 应答码：0-失败；1-成功
  --    message           C 1 应答消息：失败时返回具体的错误信息
  --    card_list[]
  --       vcard_no       C 1 就诊卡号
  --       pati_id        N 1 病人id
  ---------------------------------------------------------------------------
  l_病人id   t_Strlist := t_Strlist();
  c_病人ids  Clob;
  n_卡类别id Number;
  j_Json     Pljson;
  j_Jsonin   Pljson;
  v_List     Varchar2(32767);
Begin
  --解析入参
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  c_病人ids  := j_Json.Get_Clob('pati_ids');
  n_卡类别id := j_Json.Get_Number('card_type_id');

  Json_Out := '{"output":{"code":1,"message":"成功","card_list":[';
  While c_病人ids Is Not Null Loop
    If Length(c_病人ids) <= 4000 Then
      l_病人id.Extend;
      l_病人id(l_病人id.Count) := c_病人ids;
      c_病人ids := Null;
    Else
      l_病人id.Extend;
      l_病人id(l_病人id.Count) := Substr(c_病人ids, 1, Instr(c_病人ids, ',', 3980) - 1);
      c_病人ids := Substr(c_病人ids, Instr(c_病人ids, ',', 3980) + 1);
    End If;
  End Loop;
  For I In 1 .. l_病人id.Count Loop
    For R In (Select f_List2str(Cast(Collect(g.卡号) As t_Strlist)) As 卡号, g.病人id
              From 病人医疗卡信息 G, 医疗卡类别 H, (Select Column_Value As 病人id From Table(f_Num2list(l_病人id(I)))) A
              Where g.病人id = a.病人id And g.卡类别id = h.Id And g.状态 = 0 And h.Id = n_卡类别id
              Group By g.病人id) Loop
      Zljsonputvalue(v_List, 'pati_id', r.病人id, 1, 1);
      Zljsonputvalue(v_List, 'vcard_no', r.卡号, 0, 2);
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
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getpaticardno;
/
Create Or Replace Procedure Zl_Patisvr_Getpatiextendinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：获取病人信息从表
  --入参：Json_In:格式
  --  input
  --    pati_id             N 1 病人id
  --    info_names          C 1 信息名：多个用逗号
  --    visit_id            N 0 就诊id
  --出参: Json_Out,格式如下
  --  output
  --    code                N 1   应答码：0-失败；1-成功
  --    message             C 1   应答消息：失败时返回具体的错误信息
  --    slave_list[]        C     病人信息从表列表
  --     info_name          C 1   信息名
  --     info_value         C 1   信息值
  --     visit_id           N 1   就诊id
  ---------------------------------------------------------------------------

  j_Json    Pljson;
  j_Jsonin  Pljson;
  v_List    Varchar2(32767);
  n_病人id  病人信息从表.病人id%Type;
  v_信息名s Varchar2(32680);
  n_就诊id  病人信息从表.就诊id%Type;
Begin

  --解析入参
  j_Jsonin  := Pljson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  n_病人id  := j_Json.Get_Number('pati_id');
  v_信息名s := j_Json.Get_String('info_names');
  n_就诊id  := j_Json.Get_Number('visit_id');
  If Nvl(n_病人id, 0) = 0 Then
    Json_Out := Zljsonout('失败，未传入病人id！');
    Return;
  End If;

  Json_Out := '{"output":{"code":1,"message":"成功","slave_list":[';
  If Nvl(v_信息名s, '-') <> '-' Then
    For r_信息从表 In (Select a.信息名, a.信息值, a.就诊id
                   From 病人信息从表 A, Table(f_Str2list(v_信息名s)) B
                   Where a.病人id = n_病人id And a.信息名 = b.Column_Value And Nvl(a.就诊id, 0) = Nvl(n_就诊id, 0)) Loop
      Zljsonputvalue(v_List, 'info_name', r_信息从表.信息名, 0, 1);
      Zljsonputvalue(v_List, 'info_value', r_信息从表.信息值);
      Zljsonputvalue(v_List, 'visit_id', r_信息从表.就诊id, 1, 2);
      If Length(v_List) > 20000 Then
        Json_Out := Json_Out || v_List;
        v_List   := ',';
      End If;
    End Loop;
  Else
    For r_信息从表 In (Select Upper(信息名) 信息名, 信息值, 就诊id
                   From 病人信息从表
                   Where 病人id = n_病人id And (就诊id = n_就诊id Or 就诊id Is Null)
                   Order By Nvl(就诊id, 999999999)) Loop
      Zljsonputvalue(v_List, 'info_name', r_信息从表.信息名, 0, 1);
      Zljsonputvalue(v_List, 'info_value', r_信息从表.信息值);
      Zljsonputvalue(v_List, 'visit_id', r_信息从表.就诊id, 1, 2);
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
End Zl_Patisvr_Getpatiextendinfo;
/
Create Or Replace Procedure Zl_Patisvr_Getpatifamilymember
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：根据病人ID，获取该病人的家属成员信息
  --入参：Json_In:格式
  --  input
  --    pati_id              N   1  病人ID
  --    query_type           N   1  查询类型：0-只返回家属成员病人id；1-查询家属成员的基本信息

  --出参: Json_Out,格式如下
  --  output
  --    code                 N   1   应答吗：0-失败；1-成功
  --    message              C   1   应答消息：失败时返回具体的错误信息
  --    family_list[]        C       家属成员:病人家属
  --      pati_id            N   1   病人ID:家属id
  --      pati_relation      C   1   关系
  --      pati_name          C   1   姓名
  --      pati_sex           C   1   性别
  --      pati_age           C   1   年龄
  --      pati_birthdate     C   1   出生日期：yyyy-mm-dd hh24:mi:ss
  --      pati_nation        C   1   民族
  --      pati_idcard        C   1   身份证号
  --      family_id          N   1   家属id
  --      visit_cardno       C   1   就诊卡号
  --      state              N   1   状态
  ---------------------------------------------------------------------------

  n_查询类型 Number(1);
  n_病人id   病人信息.病人id%Type;
  v_List     Varchar2(32767);
  j_Json     Pljson;  
  j_Jsonin   Pljson;
Begin

  --解析入参
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_病人id   := j_Json.Get_Number('pati_id');
  n_查询类型 := Nvl(j_Json.Get_Number('query_type'), 0);

  If Nvl(n_病人id, 0) = 0 Then
    Json_Out := '{"output":{"code":0,"message":"未传入病人ID"}}';
    Return;
  End If;

  Json_Out := '{"output":{"code":1,"message":"成功","family_list":[';

  For r_病人信息 In (Select /*+cardinality(B,10)*/
                  b.家属id, a.就诊卡号, a.病人id, b.关系, a.姓名, a.性别, a.年龄, a.出生日期, a.民族, a.身份证号, 1 As 状态
                 From 病人信息 A, 病人家属 B
                 Where a.病人id = b.家属id And b.病人id = n_病人id And
                       (b.撤档时间 Is Null Or b.撤档时间 = To_Date('3000-01-01 00:00:00', 'yyyy-mm-dd hh24:mi:ss'))) Loop
    --      pati_id            N   1   病人ID:家属id
    --      pati_relation      C   1   关系
    --      pati_name          C   1   姓名
    --      pati_sex           C   1   性别
    --      pati_age           C   1   年龄
    --      pati_birthdate     C   1   出生日期：yyyy-mm-dd hh24:mi:ss
    --      pati_nation        C   1   民族
    --      pati_idcard        C   1   身份证号
  
    If n_查询类型 = 1 Then
      Zljsonputvalue(v_List, 'pati_id', r_病人信息.病人id, 1, 1);
      Zljsonputvalue(v_List, 'pati_relation', r_病人信息.关系);
      Zljsonputvalue(v_List, 'pati_name', r_病人信息.姓名);
      Zljsonputvalue(v_List, 'pati_sex', r_病人信息.性别);
      Zljsonputvalue(v_List, 'pati_age', r_病人信息.年龄);
      Zljsonputvalue(v_List, 'pati_birthdate', To_Char(r_病人信息.出生日期, 'yyyy-mm-dd hh24:mi:ss'));
      Zljsonputvalue(v_List, 'pati_nation', r_病人信息.民族);
      Zljsonputvalue(v_List, 'pati_idcard', r_病人信息.身份证号);
      Zljsonputvalue(v_List, 'family_id', r_病人信息.家属id, 1);
      Zljsonputvalue(v_List, 'visit_cardno', r_病人信息.就诊卡号);
      Zljsonputvalue(v_List, 'state', r_病人信息.状态, 1, 2);
    Else
      v_List := v_List || ',{"pati_id":' || Nvl(r_病人信息.病人id, 0) || '}';
    End If;
  End Loop;
  If n_查询类型 = 1 Then
    Json_Out := Json_Out || v_List || ']}}';
    Return;
  Else
    v_List   := Substr(v_List, 2);
    Json_Out := Json_Out || v_List || ']}}';
    Return;
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getpatifamilymember;
/
Create Or Replace Procedure Zl_Patisvr_Getpatiid
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:根据指定条件获取病人信息的病人ID
  --入参：Json_In:格式
  --  input
  --     card_find             C
  --         cardtype_id       N  1  医疗卡类别ID:=0时，表示模糊查找
  --         card_no           C  1  卡号
  --         qrcode            C     二维码
  --         is_check_usetime  N  1  是否检查使用时间:1-检查;0-不检查
  --         is_check_stop     N  1  是否检查停用或挂失:1-检查;0-不检查
  --     comminuty_find        C
  --        comminuty_num      N  1  社区序号
  --        comminuty_code     C     社区号
  --     other_cons_find       C
  --        find_name          C  1  查找的名称
  --        find_text          C  1  查找的文本
  --        pati_id            N     有此节点时过滤数据排开此病人ID
  --     is_stop               N     查找停用 0-不找停用的 1-找停用的

  --出参: Json_Out,格式如下
  --  output
  --    code                        N  1  应答码：0-失败；1-成功
  --    message                     C  1  应答消息：失败时返回具体的错误信息
  --  pati_list[]                   C  1  病人列表,模糊查找时，可能存在多个
  --         cardtype_id            N  1  卡类别ID
  --         pati_id                N  1  病人ID:未找到时也成功，返回0
  --         card_pwd               C  1  密码
  --         pati_pageid            N  1  主页ID
  --         enduse_time            C  1  终止使用时间:yyyy-mm-dd hh24mi:ss
  --         card_status            N  1  当前卡状态。0-正常有效卡;1-已挂失; 2-补卡停用;3-失效卡（病认医疗卡信息.终止使用时间到期时返回该状态，仅本服务使用）

  ---------------------------------------------------------------------------
  j_Json           Pljson;
  j_Jsonin         Pljson;
  j_Tmp            Pljson;
  n_病人id         病人信息.病人id%Type;
  n_卡类别id       病人医疗卡信息.卡类别id%Type;
  v_卡号           病人医疗卡信息.卡号%Type;
  n_主页id         病人信息.主页id%Type;
  v_二维码         Varchar2(500);
  n_社区           Number(5);
  v_社区号         Varchar2(500);
  v_查找名称       Varchar2(50);
  v_查找值         Varchar2(500);
  n_Find           Number(2);
  v_Err_Msg        Varchar2(500);
  n_检查期效       Number(2);
  n_检查停用及挂失 Number(2);
  n_停用           Number;
  n_排开病人id     Number;
  v_List           Varchar2(32767);
  --组装失败时返回的数据
  Function Get_Err_Message
  (
    Message_In    Varchar2,
    当前卡状态_In 病人医疗卡信息.状态%Type := 0
  ) Return Varchar2 Is
    j_Out Varchar2(32767);
  Begin
    j_Out := '{"output":{"code":0,"message":"' || Zljsonstr(Message_In) || '","card_status":' || 当前卡状态_In || '}}';
    Return j_Out;
  End Get_Err_Message;

  --组装成功信息
  Function Get_Succes_Message
  (
    卡类别id_In     医疗卡类别.Id%Type,
    病人id_In       病人信息.病人id%Type,
    主页id_In       病人信息.主页id%Type,
    密码_In         病人医疗卡信息.密码%Type := Null,
    终止使用时间_In Varchar2 := Null,
    当前卡状态_In   病人医疗卡信息.状态%Type := Null
  ) Return Varchar2 Is
    j_Out  Varchar2(32767);
    v_List Varchar2(32767);
  Begin
    v_List := '';
    If Nvl(病人id_In, 0) <> 0 Then
      v_List := '{"cardtype_id":' || Nvl(卡类别id_In, 0) || ',';
      v_List := v_List || '"pati_id":' || Nvl(病人id_In, 0) || ',';
      v_List := v_List || '"pati_pageid":' || Nvl(主页id_In, 0) || ',';
      v_List := v_List || '"card_pwd":"' || Nvl(密码_In, '') || '",';
      v_List := v_List || '"enduse_time":"' || Nvl(终止使用时间_In, '') || '",';
      v_List := v_List || '"card_status":' || Nvl(当前卡状态_In, 0) || '}';
    End If;
  
    j_Out := '{"output":{"code":1,"message":"成功","pati_list":[' || v_List || ']}}';
    Return j_Out;
  End Get_Succes_Message;

Begin
  --解析入参
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_停用   := j_Json.Get_Number('is_stop');
  --1.按医疗卡类信息查询
  If j_Json.Exist('card_find') Then
    --              cardtype_id       N  1  医疗卡类别ID:=0时，表示模糊查找
    --              card_no           C  1  卡号
    --              qrcode            C     二维码
    --              is_check_usetime  N  1  是否检查使用时间:1-检查;0-不检查
    --              is_check_stop     N  1  是否检查停用或挂失:1-检查;0-不检查
    j_Tmp            := Pljson();
    j_Tmp            := j_Json.Get_Pljson('card_find');
    n_卡类别id       := j_Tmp.Get_Number('cardtype_id');
    v_卡号           := j_Tmp.Get_String('card_no');
    v_二维码         := j_Tmp.Get_String('qrcode');
    n_检查期效       := j_Tmp.Get_Number('is_check_usetime');
    n_检查停用及挂失 := j_Tmp.Get_Number('is_check_stop');
  
    --1.1 按卡类别id查找
    If n_卡类别id <> 0 Then
      --1.1.1按医疗卡类别ID来查找
      If v_卡号 Is Not Null Then
      
        For c_病人 In (Select a.卡类别id, a.病人id, 密码, a.状态,
                            Nvl(挂失时间, To_Date('3000-01-01', 'yyyy-mm-dd')) + Nvl(b.有效天数, 0) As 挂失时间, Sysdate As 当前时间,
                            a.终止使用时间, c.主页id
                     From 病人医疗卡信息 A, 病人信息 C, 医疗卡挂失方式 B
                     Where a.病人id = c.病人id And a.卡类别id = n_卡类别id And a.卡号 = v_卡号 And a.挂失方式 = b.名称(+) And c.停用时间 Is Null)
        
         Loop
        
          If c_病人.终止使用时间 Is Not Null And Nvl(n_检查期效, 0) = 1 Then
            If c_病人.终止使用时间 <= c_病人.当前时间 Then
              v_Err_Msg := '卡号为' || v_卡号 || '已失效';
              Json_Out  := Get_Err_Message(v_Err_Msg, 3);
              Return;
            End If;
          End If;
        
          --0-正常有效卡;1-已挂失; 2-补卡停用
          If Nvl(c_病人.状态, 0) = 1 And Nvl(n_检查停用及挂失, 0) = 1 Then
            --挂失检查
            If Nvl(c_病人.挂失时间, c_病人.当前时间 - 1) < c_病人.当前时间 Then
              v_Err_Msg := '卡号为' || v_卡号 || '已挂失!';
              Json_Out  := Get_Err_Message(v_Err_Msg, c_病人.状态);
              Return;
            End If;
          End If;
        
          If Nvl(c_病人.状态, 0) = 2 And Nvl(n_检查停用及挂失, 0) = 1 Then
            --停用检查
            v_Err_Msg := '卡号为' || v_卡号 || '已停用!';
            Json_Out  := Get_Err_Message(v_Err_Msg, c_病人.状态);
            Return;
          End If;
          Json_Out := Get_Succes_Message(n_卡类别id, c_病人.病人id, c_病人.主页id, c_病人.密码,
                                         To_Char(c_病人.终止使用时间, 'yyyy-mm-dd hh24:mi:ss'), Nvl(c_病人.状态, 0));
          Return;
        End Loop;
      End If;
    
      --1.1.2 按二维码查询
      For c_病人 In (Select a.卡类别id, a.病人id, 密码, a.状态,
                          Nvl(挂失时间, To_Date('3000-01-01', 'yyyy-mm-dd')) + Nvl(b.有效天数, 0) As 挂失时间, Sysdate As 当前时间,
                          a.终止使用时间, c.主页id
                   From 病人医疗卡信息 A, 病人信息 C, 医疗卡挂失方式 B
                   Where a.病人id = c.病人id And a.卡类别id = n_卡类别id And a.二维码 = v_二维码 And a.挂失方式 = b.名称(+) And c.停用时间 Is Null)
      
       Loop
      
        If c_病人.终止使用时间 Is Not Null And Nvl(n_检查期效, 0) = 1 Then
          If c_病人.终止使用时间 <= c_病人.当前时间 Then
            v_Err_Msg := '卡号为' || v_卡号 || '已失效';
            Json_Out  := Get_Err_Message(v_Err_Msg, 3);
            Return;
          End If;
        End If;
      
        --0-正常有效卡;1-已挂失; 2-补卡停用
        If Nvl(c_病人.状态, 0) = 1 And Nvl(n_检查停用及挂失, 0) = 1 Then
          --挂失检查
          If Nvl(c_病人.挂失时间, c_病人.当前时间 - 1) < c_病人.当前时间 Then
            v_Err_Msg := '卡号为' || v_卡号 || '已挂失!';
            Json_Out  := Get_Err_Message(v_Err_Msg, c_病人.状态);
            Return;
          End If;
        End If;
      
        If Nvl(c_病人.状态, 0) = 2 And Nvl(n_检查停用及挂失, 0) = 1 Then
          --停用检查
          v_Err_Msg := '卡号为' || v_卡号 || '已停用!';
          Json_Out  := Get_Err_Message(v_Err_Msg, c_病人.状态);
          Return;
        End If;
        Json_Out := Get_Succes_Message(n_卡类别id, c_病人.病人id, c_病人.主页id, c_病人.密码,
                                       To_Char(c_病人.终止使用时间, 'yyyy-mm-dd hh24:mi:ss'), Nvl(c_病人.状态, 0));
        Return;
      
      End Loop;
    
      Json_Out := Get_Succes_Message(Null, Null, Null);
      Return;
    
    End If;
  
    --1.2 .模糊模找
    --1.2.1 按卡号模糊查找
    If v_卡号 Is Not Null Then
    
      v_Err_Msg := Null;
      For c_卡号 In (Select a.病人id, a.卡类别id, a.卡号, a.密码, a.状态, a.挂失时间, a.挂失方式, a.挂失人, a.发卡日期, a.发卡人, a.终止使用时间, a.二维码,
                          d.主页id
                   From 病人医疗卡信息 A, 医疗卡挂失方式 B, 医疗卡类别 C, 病人信息 D
                   Where a.卡类别id = c.Id And Nvl(c.是否模糊查找, 0) = 1 And a.病人id = d.病人id And a.卡号 = v_卡号 And
                         a.挂失方式 = b.名称(+) And Nvl(c.是否启用, 0) = 1 And d.停用时间 Is Null
                   Order By a.状态) Loop
        n_Find := 1;
        If c_卡号.终止使用时间 Is Not Null And Nvl(n_检查期效, 0) = 1 Then
          If c_卡号.终止使用时间 <= Sysdate Then
            If v_Err_Msg Is Null Then
              v_Err_Msg := '卡号为' || v_卡号 || '已失效';
            End If;
            n_Find := 0;
          End If;
        End If;
      
        --0-正常有效卡;1-已挂失; 2-补卡停用
        If Nvl(c_卡号.状态, 0) = 1 And Nvl(n_检查停用及挂失, 0) = 1 Then
          --挂失检查
          If Nvl(c_卡号.挂失时间, Sysdate - 1) < Sysdate Then
            If v_Err_Msg Is Null Then
              v_Err_Msg := '卡号为' || v_卡号 || '已挂失!';
            End If;
            n_Find := 0;
          End If;
        End If;
      
        If Nvl(c_卡号.状态, 0) = 2 And Nvl(n_检查停用及挂失, 0) = 1 Then
          --停用检查
          If v_Err_Msg Is Null Then
            v_Err_Msg := '卡号为' || v_卡号 || '已停用!';
          End If;
          n_Find := 0;
        End If;
      
        If n_Find = 1 Then
          Zljsonputvalue(v_List, 'cardtype_id', Nvl(c_卡号.卡类别id, 0), 1, 1);
          Zljsonputvalue(v_List, 'pati_id', Nvl(c_卡号.病人id, 0), 1);
          Zljsonputvalue(v_List, 'pati_pageid', Nvl(c_卡号.主页id, 0), 1);
          Zljsonputvalue(v_List, 'card_pwd', Nvl(c_卡号.密码, ''));
          Zljsonputvalue(v_List, 'enduse_time', Nvl(To_Char(c_卡号.终止使用时间, 'yyyy-mm-dd hh24:mi:ss'), ''));
          Zljsonputvalue(v_List, 'card_status', Nvl(c_卡号.状态, 0), 1, 2);
        End If;
      End Loop;
    
      If v_List Is Null Then
        Json_Out := Get_Err_Message(v_Err_Msg);
        Return;
      End If;
    
      Json_Out := '{"output":{"code":1,"message":"成功","pati_list":[' || v_List || ']}}';
      Return;
    End If;
  
    --1.2.2.按二维码进行糊糊查找
    If v_二维码 Is Not Null Then
    
      v_Err_Msg := Null;
      For c_卡号 In (Select a.病人id, a.卡类别id, a.卡号, a.密码, a.状态, a.挂失时间, a.挂失方式, a.挂失人, a.发卡日期, a.发卡人, a.终止使用时间, a.二维码,
                          d.主页id
                   From 病人医疗卡信息 A, 医疗卡挂失方式 B, 医疗卡类别 C, 病人信息 D
                   Where a.卡类别id = c.Id And Nvl(c.是否模糊查找, 0) = 1 And a.病人id = d.病人id And a.二维码 = v_二维码 And
                         a.挂失方式 = b.名称(+) And Nvl(c.是否启用, 0) = 1 And d.停用时间 Is Null
                   Order By a.状态) Loop
        n_Find := 1;
        If c_卡号.终止使用时间 Is Not Null And Nvl(n_检查期效, 0) = 1 Then
          If c_卡号.终止使用时间 <= Sysdate Then
            If v_Err_Msg Is Null Then
              v_Err_Msg := '卡号为' || v_卡号 || '已失效';
            End If;
            n_Find := 0;
          End If;
        End If;
      
        --0-正常有效卡;1-已挂失; 2-补卡停用
        If Nvl(c_卡号.状态, 0) = 1 And Nvl(n_检查停用及挂失, 0) = 1 Then
          --挂失检查
          If Nvl(c_卡号.挂失时间, Sysdate - 1) < Sysdate Then
            If v_Err_Msg Is Null Then
              v_Err_Msg := '卡号为' || v_卡号 || '已挂失!';
            End If;
            n_Find := 0;
          End If;
        End If;
      
        If Nvl(c_卡号.状态, 0) = 2 And Nvl(n_检查停用及挂失, 0) = 1 Then
          --停用检查
          If v_Err_Msg Is Null Then
            v_Err_Msg := '卡号为' || v_卡号 || '已停用!';
          End If;
          n_Find := 0;
        End If;
      
        If n_Find = 1 Then
          Zljsonputvalue(v_List, 'cardtype_id', Nvl(c_卡号.卡类别id, 0), 1, 1);
          Zljsonputvalue(v_List, 'pati_id', Nvl(c_卡号.病人id, 0), 1);
          Zljsonputvalue(v_List, 'pati_pageid', Nvl(c_卡号.主页id, 0), 1);
          Zljsonputvalue(v_List, 'card_pwd', Nvl(c_卡号.密码, ''));
          Zljsonputvalue(v_List, 'enduse_time', Nvl(To_Char(c_卡号.终止使用时间, 'yyyy-mm-dd hh24:mi:ss'), ''));
          Zljsonputvalue(v_List, 'card_status', Nvl(c_卡号.状态, 0), 1, 2);
        End If;
      End Loop;
    
      If v_List Is Null Then
        Json_Out := Get_Err_Message(v_Err_Msg);
        Return;
      End If;
    
      Json_Out := '{"output":{"code":1,"message":"成功","pati_list":[' || v_List || ']}}';
      Return;
    End If;
  
    Return;
  
    v_Err_Msg := '未传入医疗卡信息条件，请检查';
    Json_Out  := Get_Err_Message(v_Err_Msg);
    Return;
  End If;

  --2.按社区号来查找病人
  If j_Json.Exist('comminuty_find') Then
    --            comminuty_num       N  1  社区序号
    --            comminuty_code      C     社区号
    j_Tmp    := Pljson();
    j_Tmp    := j_Json.Get_Pljson('comminuty_find');
    n_社区   := j_Tmp.Get_Number('comminuty_num');
    v_社区号 := j_Tmp.Get_String('comminuty_code');
  
    If Nvl(n_社区, 0) = 0 Or Nvl(v_社区号, '-') = '-' Then
      v_Err_Msg := '未传入社区信息条件，请检查';
      Json_Out  := Get_Err_Message(v_Err_Msg);
      Return;
    End If;
  
    Select Max(a.病人id), Max(b.主页id)
    Into n_病人id, n_主页id
    From 病人社区信息 A, 病人信息 B
    Where a.病人id = b.病人id And a.社区 = n_社区 And a.社区号 = v_社区号 And
          (Nvl(n_停用, 0) = 0 Or (Nvl(n_停用, 0) = 1 And b.停用时间 Is Null));
  
    Json_Out := Get_Succes_Message(0, n_病人id, n_主页id, '');
    Return;
  
  End If;

  --3.按其他方式查找
  If Not j_Json.Exist('other_cons_find') Then
    v_Err_Msg := '未传入信息查询条件，请检查';
    Json_Out  := Get_Err_Message(v_Err_Msg);
    Return;
  End If;
  j_Tmp        := Pljson();
  j_Tmp        := j_Json.Get_Pljson('other_cons_find');
  v_查找名称   := j_Tmp.Get_String('find_name');
  v_查找值     := j_Tmp.Get_String('find_text');
  n_排开病人id := j_Tmp.Get_Number('pati_id');
  If Nvl(v_查找名称, '-') = '-' Or Nvl(v_查找值, '-') = '-' Then
    v_Err_Msg := '未传入信息查询条件，请检查';
    Json_Out  := Get_Err_Message(v_Err_Msg);
    Return;
  End If;
  v_查找名称 := Replace(v_查找名称, ' ', '');
  If v_查找名称 = 'IC卡' Or v_查找名称 = 'IC卡号' Then
    Select Max(病人id), Max(主页id)
    Into n_病人id, n_主页id
    From 病人信息
    Where Ic卡号 = v_查找值 And (Nvl(n_停用, 0) = 0 Or (Nvl(n_停用, 0) = 1 And 停用时间 Is Null));
  Elsif v_查找名称 = '身份证' Or v_查找名称 = '身份证号' Or v_查找名称 = '二代身份证' Then
    Select Max(病人id), Max(主页id)
    Into n_病人id, n_主页id
    From 病人信息
    Where 身份证号 = v_查找值 And (Nvl(n_停用, 0) = 0 Or (Nvl(n_停用, 0) = 1 And 停用时间 Is Null));
  Elsif v_查找名称 = '医保号' Or v_查找名称 = '医保证号' Then
    --医保号支持模糊查找，仅北京医保
    If Instr(v_查找值, '%') > 0 Then
      Select Max(病人id), Max(主页id)
      Into n_病人id, n_主页id
      From 病人信息
      Where 医保号 Like v_查找值 And (Nvl(n_停用, 0) = 0 Or (Nvl(n_停用, 0) = 1 And 停用时间 Is Null));
    Else
      Select Max(病人id), Max(主页id)
      Into n_病人id, n_主页id
      From 病人信息
      Where 医保号 = v_查找值 And (Nvl(n_停用, 0) = 0 Or (Nvl(n_停用, 0) = 1 And 停用时间 Is Null));
    End If;
  Elsif v_查找名称 = '手机号' Then
    Select Max(病人id), Max(主页id)
    Into n_病人id, n_主页id
    From 病人信息
    Where 手机号 = v_查找值 And (Nvl(n_停用, 0) = 0 Or (Nvl(n_停用, 0) = 1 And 停用时间 Is Null)) And
          (Nvl(n_排开病人id, 0) = 0 Or 病人id <> Nvl(n_排开病人id, 0));
  Elsif v_查找名称 = '门诊号' Then
    Select Max(病人id), Max(主页id)
    Into n_病人id, n_主页id
    From 病人信息
    Where 门诊号 = v_查找值 And (Nvl(n_停用, 0) = 0 Or (Nvl(n_停用, 0) = 1 And 停用时间 Is Null)) And
          (Nvl(n_排开病人id, 0) = 0 Or 病人id <> Nvl(n_排开病人id, 0));
  Elsif v_查找名称 = '住院号' Then
    Select Max(病人id), Max(主页id)
    Into n_病人id, n_主页id
    From 病人信息
    Where 住院号 = v_查找值 And (Nvl(n_停用, 0) = 0 Or (Nvl(n_停用, 0) = 1 And 停用时间 Is Null)) And
          (Nvl(n_排开病人id, 0) = 0 Or 病人id <> Nvl(n_排开病人id, 0));
  Elsif Upper(v_查找名称) = Upper('病人ID') Then
    Select Max(病人id), Max(主页id) Into n_病人id, n_主页id From 病人信息 Where 病人id = v_查找值;
  Elsif v_查找名称 = '健康号' Then
    Select Max(病人id), Max(主页id)
    Into n_病人id, n_主页id
    From 病人信息
    Where 健康号 = v_查找值 And (Nvl(n_停用, 0) = 0 Or (Nvl(n_停用, 0) = 1 And 停用时间 Is Null));
  Elsif v_查找名称 = '就诊卡号' Then
    Select Max(病人id), Max(主页id)
    Into n_病人id, n_主页id
    From 病人信息
    Where 就诊卡号 = v_查找值 And (Nvl(n_停用, 0) = 0 Or (Nvl(n_停用, 0) = 1 And 停用时间 Is Null));
  Elsif v_查找名称 = '姓名' Then
    If Instr(v_查找值, '%') > 0 Then
      Select Max(病人id), Max(主页id)
      Into n_病人id, n_主页id
      From 病人信息
      Where 姓名 Like v_查找值 And (Nvl(n_停用, 0) = 0 Or (Nvl(n_停用, 0) = 1 And 停用时间 Is Null));
    Else
      Select Max(病人id), Max(主页id)
      Into n_病人id, n_主页id
      From 病人信息
      Where 姓名 = v_查找值 And (Nvl(n_停用, 0) = 0 Or (Nvl(n_停用, 0) = 1 And 停用时间 Is Null));
    End If;
  Else
    --不支持的方式
    v_Err_Msg := '不支持' || v_查找名称 || '方式查找病人!';
    Json_Out  := Get_Err_Message(v_Err_Msg);
    Return;
  End If;
  If Nvl(n_病人id, 0) <> 0 Then
    Json_Out := Get_Succes_Message(0, n_病人id, n_主页id, '');
    Return;
  End If;

  Json_Out := Get_Succes_Message(0, n_病人id, n_主页id, '');

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getpatiid;
/
Create Or Replace Procedure Zl_Patisvr_Getpatiidsbyrange
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:根据指定条件获取病人信息的病人ID
  --入参：Json_In:格式
  --    input
  --      query_condition C 1 查询条件
  --      ctt_unit_id     N 1 合同单位ID，查询指定合同单位的门诊病人
  --出参: Json_Out,格式如下
  --  output
  --    code              N  1  应答码：0-失败；1-成功
  --    message           C  1  应答消息：失败时返回具体的错误信息
  --    pati_ids          C  1  病人IDs，逗号拼串
  ---------------------------------------------------------------------------
  j_Json    Pljson;
  j_Jsonin  Pljson;
  v_查找值  Varchar2(3000);
  v_Temp    Varchar2(500);
  v_病人ids Varchar2(32767);

  n_合同单位id 病人信息.合同单位id%Type;

Begin
  --解析入参
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');

  If j_Json.Exist('query_condition') Then
    v_查找值 := j_Json.Get_String('query_condition');
    If v_查找值 Is Null Then
      Json_Out := Zljsonout('未传入查询条件，请检查！');
      Return;
    End If;
  
    Select LTrim(v_查找值, '0123456789') Into v_Temp From Dual;
    If v_Temp Is Null Then
      Select f_List2str(Cast(Collect(To_Char(a.病人id)) As t_Strlist))
      Into v_病人ids
      From 病人信息 A
      Where a.门诊号 = To_Number(v_查找值) Or a.就诊卡号 = v_查找值 Or a.身份证号 = v_查找值 Or a.Ic卡号 = v_查找值;
    Else
      Select f_List2str(Cast(Collect(To_Char(a.病人id)) As t_Strlist))
      Into v_病人ids
      From 病人信息 A
      Where a.就诊卡号 = v_查找值 Or a.身份证号 = v_查找值 Or a.Ic卡号 = v_查找值;
    End If;
  
    Json_Out := '{"output":{"code":1,"message":"成功","pati_ids":"' || v_病人ids || '"}}';
    Return;
  End If;

  --按合同单位ID获取
  If j_Json.Exist('ctt_unit_id') Then
    n_合同单位id := j_Json.Get_Number('ctt_unit_id');
    If Nvl(n_合同单位id, 0) = 0 Then
      Json_Out := Zljsonout('未传入合同单位ID，请检查！');
      Return;
    End If;
    v_病人ids := Null;
  
    For c_病人 In (Select Distinct a.病人id From 病人信息 A Where a.合同单位id = n_合同单位id And a.当前科室id Is Null) Loop
      v_病人ids := v_病人ids || ',' || c_病人.病人id;
    End Loop;
  
    Json_Out := '{"output":{"code":1,"message":"成功","pati_ids":"' || Substr(v_病人ids, 2) || '"}}';
    Return;
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getpatiidsbyrange;
/
Create Or Replace Procedure Zl_Patisvr_Getpatiinfo
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:获取病人信息
  --入参：Json_In:格式
  --    input
  --      pati_id           N 1 病人id  病人ID<>0时，查询列表中的条件无效
  --      query_type        N 1 查询类型:如：0-基本;1-基本+联系人;2-所有
  --      query_card        N 1 是否包含卡信息:1-包含医疗卡;0-不包含医疗卡
  --      query_family      N 1 是否包含家属:1-包含家属信息，0-不包含家属信息
  --      query_drug        N 1 是否包含过敏药物:1-包含，0-不包含
  --      query_immune      N 1 是否包含免疫修:1-包含;0-不包含
  --      query_insurance_pwd C  是否包含医保密码:1-包含;0-不包含
  --      query_cons_list   C 1 查询条件:可以选择一定条件进行查询（是And关系),只有一行
  --        pati_ids        C   病人IDs:多个用逗号
  --        pati_name       C   姓名:可以代%分号表表按姓名匹配
  --        outpatient_num  C   门诊号
  --        inpatient_num   C   住院号
  --        pati_idcard     C   身份证号
  --        contacts_idcard C   联系人身份证号
  --        cardtype_id     N   医疗卡类别ID
  --        medc_card_name  N   医疗卡名称
  --        card_no         C   卡号
  --        qrcode          C   二维码
  --        iccard_no       C   Ic卡号
  --        visit_card      C   就诊卡号
  --        insurance_num   C   医保号
  --        qrspt_statu     C   查询住院状态:0-仅门诊;1-在院 ;2-门诊及在院
  --        phone_number    C   手机号
  --        pati_bed        C   当前床号
  --        dept_id         N   当前科室ID
  --        search_days     N   有此节点时按指定查找天数查找病人(用于姓名模糊查找)
  --出参      json
  --output
  --    code                N   1   应答码：0-失败；1-成功
  --    message             C   1   应答消息： 失败时返回具体的错误信息
  --    pati_list[]                 病人信息列表
  --    pati_id             N   1   病人id
  --    pati_pageid         N   1   主页id：病人信息.主页ID
  --    pati_name           C   1   姓名
  --    pati_sex            C   1   性别
  --    pati_age            C   1   年龄
  --    pati_birthdate      C   1   出生日期：yyyy-mm-dd hh24:mi:ss
  --    fee_category        C   1   费别
  --    outpatient_num      C   1   门诊号
  --    inpatient_num       C   1   住院号
  --    mdlpay_mode_name    C   1   医疗付款方式名称
  --    mdlpay_mode_code    C   1   医疗付款方式编码
  --    pati_nation         C   1   民族
  --    insurance_num       C   1   医保号
  --    pati_idcard         C   1   身份证号
  --    vcard_no            C   1   就诊卡号
  --    iccard_no           C   1   Ic卡号
  --    health_num          C   1   健康号
  --    inp_times           N   1   住院次数
  --    pati_education      C   1   学历
  --    ocpt_name           C   1   职业
  --    pati_identity       C   1   身份
  --    ntvplc_name         C   1   籍贯
  --    country_name        C   1   国籍
  --    pati_marital_cstatus    C   1   婚姻状况
  --    pat_home_addr           C   1   家庭地址
  --    pat_home_phno           C   1   家庭电话
  --    pat_home_postcode   C   1   家庭地址邮编
  --    pati_area           C   1   区域
  --    pati_birthplace     C   1   出生地点
  --    pat_hous_addr       C   1   户口地址
  --    pat_hous_postcode   C   1   户口地址邮编
  --    emp_name            C   1   工作单位名称
  --    emp_phno            C   1   单位电话
  --    emp_postcode        C   1   单位邮编
  --    emp_bank_name       C   1   单位开户行
  --    emp_bank_accnum     C   1   单位帐号
  --    emp_addr             C   1   单位地址
  --    ctt_unit_id         N   1   合同单位ID
  --    phone_number        C   1   手机号
  --    pati_bed            C   1   当前床号
  --    pati_type           C   1   病人类型(普通，医保，留观)
  --    insurance_type      C   1   险类
  --    insurance_name      C   1   险类名称
  --    pati_wardarea_id    N   1   当前病区id
  --    pati_wardarea_name  C   1   当前病区名称
  --    pati_dept_id        N   1   当前科室id
  --    pati_dept_name      C   1   当前科室名称
  --    adta_time           C   1   入院时间:yyyy-mm-dd hh24:mi:ss
  --    adtd_time           C   1   出院时间:yyyy-mm-dd hh24:mi:ss
  --    contacts_name       C   1   联系人姓名
  --    contacts_relation   C   1   联系人关系
  --    contacts_idcard     C   1   联系人身份证号
  --    contacts_addr       C   1   联系人地址
  --    contacts_phno       C   1   联系人电话
  --    pat_grdn_name       C   1   监护人
  --    cert_no_other       C   1   其他证件
  --    is_inhspt            C   1   是否在院:1-在院 ;0-不在院
  --    pati_show_color      N   1   病人显示颜色
  --    visit_room           C   1   就诊诊室
  --    visit_statu          N   1   就诊状态
  --    visit_time           C   1   就诊时间:yyyy-mm-dd hh24:mi:ss
  --    create_time          C   1   登记时间:yyyy-mm-dd hh24:mi:ss
  --    pati_email           C   1   email
  --    pati_qq              C   1   qq
  --    card_captcha         C   1  卡验证码
  --    insurance_pwd        C       医保密码
  --    family_list[]        C   1   家属成员:病人家属() query_family=1返回
  --        family_id        N   1   家属id  query_family=1
  --        family_relation  C   1   关系
  --    drug_list[]          C   1   过敏药物列表    query_drug=1时返回
  --        pat_algc_cadn_id N   1   过敏药品ID
  --        pat_algc_cadn    C   1   过敏药物名称
  --        allergy_info     C   1   过每药物反应
  --    immune_list[]        C   1   病人免疫列表    query_immune=1时返回
  --        vaccinate_time   C   1   接种时间:yyyy-mm-dd hh24:mi:ss
  --        vaccinate_name   C   1   接种名称
  --    card_list[]          C   1   病人医疗卡信息列表(如果条件中传入了卡类别ID的，则返回该卡类别的卡信息)  query_card=1时返回
  --        cardtype_id      N   1   医疗卡类别ID
  --        card_no          C   1   卡号
  --        card_pwd         C   1   密码
  ---------------------------------------------------------------------------
  v_Err_Msg Varchar2(500);
  j_Json    Pljson;
  o_Json    Pljson;
  j_Jsonin  Pljson;

  n_卡类别id   医疗卡类别.Id%Type;
  v_医疗卡名称 医疗卡类别.名称%Type;
  n_病人id     病人医疗卡信息.病人id%Type;
  n_门诊号     病人信息.门诊号%Type;
  n_住院号     病人信息.住院号%Type;
  v_身份证号   病人信息.身份证号%Type;
  n_颜色       病人类型.颜色%Type;

  v_姓名           Varchar2(100);
  v_卡号           Varchar2(1000);
  v_就诊卡号       病人信息.就诊卡号%Type;
  v_二维码         Varchar2(1000);
  v_手机号         Varchar2(50);
  v_联系人身份证号 Varchar2(50);
  v_医保号         Varchar2(30);
  v_床号           Varchar2(30);
  v_Ic卡号         Varchar2(100);
  n_查询类型       Number(2);
  n_查找天数       Number(10);
  n_查询住院状态   Number(2);
  n_科室id         病人信息.当前科室id%Type;

  n_是否包含卡信息   Number(2);
  n_是否包含家属     Number(2);
  n_是否包含过敏药物 Number(2);
  n_是否包含免疫信息 Number(2);
  n_包含医保密码     Number(2);

  l_病人ids  t_Strlist;
  c_病人ids  Clob;
  P          Number;
  v_医保密码 医保病人档案.密码%Type;

  Cursor c_病人基本信息 Is
    Select 病人id, a.姓名, a.性别, a.年龄, a.出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号, a.就诊卡号, a.卡验证码, a.费别, a.医疗付款方式 As 医疗付款方式名称,
           b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.婚姻状况, a.家庭地址, a.家庭电话, a.家庭地址邮编, a.联系人姓名, a.联系人关系,
           a.联系人地址, a.联系人电话, a.合同单位id, a.工作单位 As 工作单位名称, a.单位电话, a.单位邮编, a.单位开户行, a.单位帐号, a.就诊时间, a.就诊状态, a.就诊诊室, a.住院次数,
           a.当前科室id, a.当前病区id, a.入院时间, a.出院时间, a.Ic卡号, a.健康号, a.险类, a.登记时间, a.停用时间, a.当前床号, a.医保号, a.其他证件, a.监护人, a.查询密码,
           a.在院, a.锁定, a.户口地址, a.户口地址邮编, a.籍贯, a.Email, a.Qq, a.联系人身份证号, a.病人类型, a.主页id, a.手机号, c.名称 As 当前科室名称,
           d.名称 As 当前病区名称, a.单位地址 As 工作单位地址, f.名称 As 险类名称
    From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 保险类别 F
    Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.险类 = f.序号(+) And Rownum < 1;
  r_病人 c_病人基本信息%RowType;

  Type Ty_病人信息 Is Ref Cursor;
  c_病人信息 Ty_病人信息; --动态游标变量

  v_Json Varchar2(32767);
  v_Temp Varchar2(32767);

  n_Firstitem    Number;
  n_Firstsubitem Number;
Begin
  --解析入参
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_病人id   := j_Json.Get_Number('pati_id');
  n_查询类型 := j_Json.Get_Number('query_type');

  n_是否包含卡信息   := j_Json.Get_Number('query_card');
  n_是否包含家属     := j_Json.Get_Number('query_family');
  n_是否包含过敏药物 := j_Json.Get_Number('query_drug');
  n_是否包含免疫信息 := j_Json.Get_Number('query_immune');
  n_包含医保密码     := j_Json.Get_Number('query_insurance_pwd');

  o_Json := j_Json.Get_Pljson('query_cons_list');
  If Nvl(n_病人id, 0) = 0 And o_Json Is Null Then
    v_Err_Msg := '未传入有效的查询条件，不能获取病人信息!';
    Json_Out  := '{"output":{"code":0,"message":"' || Zljsonstr(v_Err_Msg) || '"}}';
  
    Return;
  End If;

  If o_Json Is Not Null Then
    Begin
      c_病人ids := o_Json.Get_Clob('pati_ids');
    Exception
      When Others Then
        c_病人ids := Null;
    End;
    If Not c_病人ids Is Null Then
      l_病人ids := t_Strlist();
      c_病人ids := c_病人ids || ',';
      Loop
        P := Instr(c_病人ids, ',');
        Exit When(Nvl(P, 0) = 0);
      
        l_病人ids.Extend;
        l_病人ids(l_病人ids.Count) := (Substr(c_病人ids, 1, P - 1));
        c_病人ids := Substr(c_病人ids, P + 1);
      End Loop;
    End If;
    v_姓名           := o_Json.Get_String('pati_name');
    n_门诊号         := To_Number(o_Json.Get_String('outpatient_num'));
    n_住院号         := To_Number(o_Json.Get_String('inpatient_num'));
    v_身份证号       := o_Json.Get_String('pati_idcard');
    v_联系人身份证号 := o_Json.Get_String('contacts_idcard');
    n_卡类别id       := o_Json.Get_Number('cardtype_id');
    v_医疗卡名称     := o_Json.Get_String('medc_card_name');
    v_卡号           := o_Json.Get_String('card_no');
    v_二维码         := o_Json.Get_String('qrcode');
    n_查询住院状态   := o_Json.Get_Number('qrspt_statu');
    v_手机号         := o_Json.Get_String('phone_number');
    v_Ic卡号         := o_Json.Get_String('iccard_no');
    v_就诊卡号       := o_Json.Get_String('visit_card');
    v_医保号         := o_Json.Get_String('insurance_num');
    v_床号           := o_Json.Get_String('pati_bed');
    n_科室id         := o_Json.Get_Number('dept_id');
    n_查找天数       := o_Json.Get_Number('search_days');
  End If;

  If Nvl(n_病人id, 0) <> 0 Then
    --按病人ID为主要查询条件进行查询
    Open c_病人信息 For
      Select a.病人id, a.姓名, a.性别, a.年龄, a.出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号, a.就诊卡号, a.卡验证码, a.费别, a.医疗付款方式 As 医疗付款方式名称,
             b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.婚姻状况, a.家庭地址, a.家庭电话, a.家庭地址邮编, a.联系人姓名, a.联系人关系,
             a.联系人地址, a.联系人电话, a.合同单位id, a.工作单位 As 工作单位名称, a.单位电话, a.单位邮编, a.单位开户行, a.单位帐号, a.就诊时间, a.就诊状态, a.就诊诊室,
             a.住院次数, a.当前科室id, a.当前病区id, a.入院时间, a.出院时间, a.Ic卡号, a.健康号, a.险类, a.登记时间, a.停用时间, a.当前床号, a.医保号, a.其他证件,
             a.监护人, a.查询密码, a.在院, a.锁定, a.户口地址, a.户口地址邮编, a.籍贯, a.Email, a.Qq, a.联系人身份证号, a.病人类型, a.主页id, a.手机号,
             c.名称 As 当前科室名称, d.名称 As 当前病区名称, a.单位地址 As 工作单位地址, f.名称 As 险类名称
      From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 保险类别 F
      Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.险类 = f.序号(+) And a.病人id = n_病人id And
            a.停用时间 Is Null;
  
  Elsif n_门诊号 <> 0 Then
    --按门诊号查询
    Open c_病人信息 For
      Select a.病人id, a.姓名, a.性别, a.年龄, a.出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号, a.就诊卡号, a.卡验证码, a.费别, a.医疗付款方式 As 医疗付款方式名称,
             b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.婚姻状况, a.家庭地址, a.家庭电话, a.家庭地址邮编, a.联系人姓名, a.联系人关系,
             a.联系人地址, a.联系人电话, a.合同单位id, a.工作单位 As 工作单位名称, a.单位电话, a.单位邮编, a.单位开户行, a.单位帐号, a.就诊时间, a.就诊状态, a.就诊诊室,
             a.住院次数, a.当前科室id, a.当前病区id, a.入院时间, a.出院时间, a.Ic卡号, a.健康号, a.险类, a.登记时间, a.停用时间, a.当前床号, a.医保号, a.其他证件,
             a.监护人, a.查询密码, a.在院, a.锁定, a.户口地址, a.户口地址邮编, a.籍贯, a.Email, a.Qq, a.联系人身份证号, a.病人类型, a.主页id, a.手机号,
             c.名称 As 当前科室名称, d.名称 As 当前病区名称, a.单位地址 As 工作单位地址, f.名称 As 险类名称
      From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 保险类别 F
      Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.险类 = f.序号(+) And a.门诊号 = n_门诊号 And
            Decode(Nvl(n_查询住院状态, 0), 2, 2, Nvl(a.在院, 0)) = Nvl(n_查询住院状态, 0) And a.停用时间 Is Null;
    --0-仅门诊;1-在院 ;2-门诊及在院
  Elsif n_住院号 <> 0 Then
    --按门诊号查询
    Open c_病人信息 For
      Select a.病人id, a.姓名, a.性别, a.年龄, a.出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号, a.就诊卡号, a.卡验证码, a.费别, a.医疗付款方式 As 医疗付款方式名称,
             b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.婚姻状况, a.家庭地址, a.家庭电话, a.家庭地址邮编, a.联系人姓名, a.联系人关系,
             a.联系人地址, a.联系人电话, a.合同单位id, a.工作单位 As 工作单位名称, a.单位电话, a.单位邮编, a.单位开户行, a.单位帐号, a.就诊时间, a.就诊状态, a.就诊诊室,
             a.住院次数, a.当前科室id, a.当前病区id, a.入院时间, a.出院时间, a.Ic卡号, a.健康号, a.险类, a.登记时间, a.停用时间, a.当前床号, a.医保号, a.其他证件,
             a.监护人, a.查询密码, a.在院, a.锁定, a.户口地址, a.户口地址邮编, a.籍贯, a.Email, a.Qq, a.联系人身份证号, a.病人类型, a.主页id, a.手机号,
             c.名称 As 当前科室名称, d.名称 As 当前病区名称, a.单位地址 As 工作单位地址, f.名称 As 险类名称
      From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 保险类别 F
      Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.险类 = f.序号(+) And a.住院号 = n_住院号 And
            Decode(Nvl(n_查询住院状态, 0), 2, 2, Nvl(a.在院, 0)) = Nvl(n_查询住院状态, 0) And a.停用时间 Is Null;
    --0-仅门诊;1-在院 ;2-门诊及在院
  Elsif v_Ic卡号 Is Not Null Then
    Open c_病人信息 For
      Select a.病人id, a.姓名, a.性别, a.年龄, a.出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号, a.就诊卡号, a.卡验证码, a.费别, a.医疗付款方式 As 医疗付款方式名称,
             b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.婚姻状况, a.家庭地址, a.家庭电话, a.家庭地址邮编, a.联系人姓名, a.联系人关系,
             a.联系人地址, a.联系人电话, a.合同单位id, a.工作单位 As 工作单位名称, a.单位电话, a.单位邮编, a.单位开户行, a.单位帐号, a.就诊时间, a.就诊状态, a.就诊诊室,
             a.住院次数, a.当前科室id, a.当前病区id, a.入院时间, a.出院时间, a.Ic卡号, a.健康号, a.险类, a.登记时间, a.停用时间, a.当前床号, a.医保号, a.其他证件,
             a.监护人, a.查询密码, a.在院, a.锁定, a.户口地址, a.户口地址邮编, a.籍贯, a.Email, a.Qq, a.联系人身份证号, a.病人类型, a.主页id, a.手机号,
             c.名称 As 当前科室名称, d.名称 As 当前病区名称, a.单位地址 As 工作单位地址, f.名称 As 险类名称
      From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 保险类别 F
      Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.险类 = f.序号(+) And a.Ic卡号 = v_Ic卡号 And
            Decode(Nvl(n_查询住院状态, 0), 2, 2, Nvl(a.在院, 0)) = Nvl(n_查询住院状态, 0) And a.停用时间 Is Null;
    --0-仅门诊;1-在院 ;2-门诊及在院
  Elsif v_身份证号 Is Not Null Then
    Select Max(病人id)
    Into n_病人id
    From 病人医疗卡信息 A, 医疗卡类别 B
    Where Nvl(a.状态, 0) = 0 And a.卡类别id = b.Id And b.名称 = '二代身份证' And b.是否启用 = 1 And a.卡号 = v_身份证号 And
          Nvl(a.终止使用时间, Sysdate + 1) > Sysdate;
    If Nvl(n_病人id, 0) <> 0 Then
      Open c_病人信息 For
        Select a.病人id, a.姓名, a.性别, a.年龄, a.出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号, a.就诊卡号, a.卡验证码, a.费别,
               a.医疗付款方式 As 医疗付款方式名称, b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.婚姻状况, a.家庭地址, a.家庭电话,
               a.家庭地址邮编, a.联系人姓名, a.联系人关系, a.联系人地址, a.联系人电话, a.合同单位id, a.工作单位 As 工作单位名称, a.单位电话, a.单位邮编, a.单位开户行, a.单位帐号,
               a.就诊时间, a.就诊状态, a.就诊诊室, a.住院次数, a.当前科室id, a.当前病区id, a.入院时间, a.出院时间, a.Ic卡号, a.健康号, a.险类, a.登记时间, a.停用时间,
               a.当前床号, a.医保号, a.其他证件, a.监护人, a.查询密码, a.在院, a.锁定, a.户口地址, a.户口地址邮编, a.籍贯, a.Email, a.Qq, a.联系人身份证号,
               a.病人类型, a.主页id, a.手机号, c.名称 As 当前科室名称, d.名称 As 当前病区名称, a.单位地址 As 工作单位地址, f.名称 As 险类名称
        From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 保险类别 F
        Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.险类 = f.序号(+) And a.身份证号 = v_身份证号 And
              Decode(Nvl(n_查询住院状态, 0), 2, 2, Nvl(a.在院, 0)) = Nvl(n_查询住院状态, 0) And a.停用时间 Is Null;
    Else
      Open c_病人信息 For
        Select a.病人id, a.姓名, a.性别, a.年龄, a.出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号, a.就诊卡号, a.卡验证码, a.费别,
               a.医疗付款方式 As 医疗付款方式名称, b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.婚姻状况, a.家庭地址, a.家庭电话,
               a.家庭地址邮编, a.联系人姓名, a.联系人关系, a.联系人地址, a.联系人电话, a.合同单位id, a.工作单位 As 工作单位名称, a.单位电话, a.单位邮编, a.单位开户行, a.单位帐号,
               a.就诊时间, a.就诊状态, a.就诊诊室, a.住院次数, a.当前科室id, a.当前病区id, a.入院时间, a.出院时间, a.Ic卡号, a.健康号, a.险类, a.登记时间, a.停用时间,
               a.当前床号, a.医保号, a.其他证件, a.监护人, a.查询密码, a.在院, a.锁定, a.户口地址, a.户口地址邮编, a.籍贯, a.Email, a.Qq, a.联系人身份证号,
               a.病人类型, a.主页id, a.手机号, c.名称 As 当前科室名称, d.名称 As 当前病区名称, a.单位地址 As 工作单位地址, f.名称 As 险类名称
        From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 保险类别 F
        Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.险类 = f.序号(+) And a.身份证号 = v_身份证号 And
              Decode(Nvl(n_查询住院状态, 0), 2, 2, Nvl(a.在院, 0)) = Nvl(n_查询住院状态, 0) And a.停用时间 Is Null;
    End If;
    --0-仅门诊;1-在院 ;2-门诊及在院
  Elsif v_医保号 Is Not Null Then
    If Instr(v_医保号, '%') > 0 Then
      Open c_病人信息 For
        Select a.病人id, a.姓名, a.性别, a.年龄, a.出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号, a.就诊卡号, a.卡验证码, a.费别,
               a.医疗付款方式 As 医疗付款方式名称, b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.婚姻状况, a.家庭地址, a.家庭电话,
               a.家庭地址邮编, a.联系人姓名, a.联系人关系, a.联系人地址, a.联系人电话, a.合同单位id, a.工作单位 As 工作单位名称, a.单位电话, a.单位邮编, a.单位开户行, a.单位帐号,
               a.就诊时间, a.就诊状态, a.就诊诊室, a.住院次数, a.当前科室id, a.当前病区id, a.入院时间, a.出院时间, a.Ic卡号, a.健康号, a.险类, a.登记时间, a.停用时间,
               a.当前床号, a.医保号, a.其他证件, a.监护人, a.查询密码, a.在院, a.锁定, a.户口地址, a.户口地址邮编, a.籍贯, a.Email, a.Qq, a.联系人身份证号,
               a.病人类型, a.主页id, a.手机号, c.名称 As 当前科室名称, d.名称 As 当前病区名称, a.单位地址 As 工作单位地址, f.名称 As 险类名称
        From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 保险类别 F
        Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.险类 = f.序号(+) And a.医保号 Like v_医保号 And
              Decode(Nvl(n_查询住院状态, 0), 2, 2, Nvl(a.在院, 0)) = Nvl(n_查询住院状态, 0) And a.停用时间 Is Null;
    Else
      Open c_病人信息 For
        Select a.病人id, a.姓名, a.性别, a.年龄, a.出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号, a.就诊卡号, a.卡验证码, a.费别,
               a.医疗付款方式 As 医疗付款方式名称, b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.婚姻状况, a.家庭地址, a.家庭电话,
               a.家庭地址邮编, a.联系人姓名, a.联系人关系, a.联系人地址, a.联系人电话, a.合同单位id, a.工作单位 As 工作单位名称, a.单位电话, a.单位邮编, a.单位开户行, a.单位帐号,
               a.就诊时间, a.就诊状态, a.就诊诊室, a.住院次数, a.当前科室id, a.当前病区id, a.入院时间, a.出院时间, a.Ic卡号, a.健康号, a.险类, a.登记时间, a.停用时间,
               a.当前床号, a.医保号, a.其他证件, a.监护人, a.查询密码, a.在院, a.锁定, a.户口地址, a.户口地址邮编, a.籍贯, a.Email, a.Qq, a.联系人身份证号,
               a.病人类型, a.主页id, a.手机号, c.名称 As 当前科室名称, d.名称 As 当前病区名称, a.单位地址 As 工作单位地址, f.名称 As 险类名称
        From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 保险类别 F
        Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.险类 = f.序号(+) And a.医保号 = v_医保号 And
              Decode(Nvl(n_查询住院状态, 0), 2, 2, Nvl(a.在院, 0)) = Nvl(n_查询住院状态, 0) And a.停用时间 Is Null;
    End If;
  Elsif v_床号 Is Not Null Then
    Open c_病人信息 For
      Select a.病人id, a.姓名, a.性别, a.年龄, a.出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号, a.就诊卡号, a.卡验证码, a.费别, a.医疗付款方式 As 医疗付款方式名称,
             b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.婚姻状况, a.家庭地址, a.家庭电话, a.家庭地址邮编, a.联系人姓名, a.联系人关系,
             a.联系人地址, a.联系人电话, a.合同单位id, a.工作单位 As 工作单位名称, a.单位电话, a.单位邮编, a.单位开户行, a.单位帐号, a.就诊时间, a.就诊状态, a.就诊诊室,
             a.住院次数, a.当前科室id, a.当前病区id, a.入院时间, a.出院时间, a.Ic卡号, a.健康号, a.险类, a.登记时间, a.停用时间, a.当前床号, a.医保号, a.其他证件,
             a.监护人, a.查询密码, a.在院, a.锁定, a.户口地址, a.户口地址邮编, a.籍贯, a.Email, a.Qq, a.联系人身份证号, a.病人类型, a.主页id, a.手机号,
             c.名称 As 当前科室名称, d.名称 As 当前病区名称, a.单位地址 As 工作单位地址, f.名称 As 险类名称
      From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 保险类别 F
      Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.险类 = f.序号(+) And
            a.当前床号 Like '%' || v_床号 || '%' And Decode(Nvl(n_查询住院状态, 0), 2, 2, Nvl(a.在院, 0)) = Nvl(n_查询住院状态, 0) And
            a.停用时间 Is Null;
    --0-仅门诊;1-在院 ;2-门诊及在院
  Elsif l_病人ids Is Not Null Then
    Open c_病人信息 For
      Select /*+cardinality(Q,10)*/
       a.病人id, a.姓名, a.性别, a.年龄, a.出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号, a.就诊卡号, a.卡验证码, a.费别, a.医疗付款方式 As 医疗付款方式名称,
       b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.婚姻状况, a.家庭地址, a.家庭电话, a.家庭地址邮编, a.联系人姓名, a.联系人关系, a.联系人地址,
       a.联系人电话, a.合同单位id, a.工作单位 As 工作单位名称, a.单位电话, a.单位邮编, a.单位开户行, a.单位帐号, a.就诊时间, a.就诊状态, a.就诊诊室, a.住院次数, a.当前科室id,
       a.当前病区id, a.入院时间, a.出院时间, a.Ic卡号, a.健康号, a.险类, a.登记时间, a.停用时间, a.当前床号, a.医保号, a.其他证件, a.监护人, a.查询密码, a.在院, a.锁定,
       a.户口地址, a.户口地址邮编, a.籍贯, a.Email, a.Qq, a.联系人身份证号, a.病人类型, a.主页id, a.手机号, c.名称 As 当前科室名称, d.名称 As 当前病区名称,
       a.单位地址 As 工作单位地址, f.名称 As 险类名称
      From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 保险类别 F, Table(l_病人ids) Q
      Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.险类 = f.序号(+) And
            a.病人id = q.Column_Value And Decode(Nvl(n_查询住院状态, 0), 2, 2, Nvl(a.在院, 0)) = Nvl(n_查询住院状态, 0) And
            a.停用时间 Is Null;
  Elsif v_手机号 Is Not Null Then
    Open c_病人信息 For
      Select /*+cardinality(c,10)*/
       a.病人id, a.姓名, a.性别, a.年龄, a.出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号, a.就诊卡号, a.卡验证码, a.费别, a.医疗付款方式 As 医疗付款方式名称,
       b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.婚姻状况, a.家庭地址, a.家庭电话, a.家庭地址邮编, a.联系人姓名, a.联系人关系, a.联系人地址,
       a.联系人电话, a.合同单位id, a.工作单位 As 工作单位名称, a.单位电话, a.单位邮编, a.单位开户行, a.单位帐号, a.就诊时间, a.就诊状态, a.就诊诊室, a.住院次数, a.当前科室id,
       a.当前病区id, a.入院时间, a.出院时间, a.Ic卡号, a.健康号, a.险类, a.登记时间, a.停用时间, a.当前床号, a.医保号, a.其他证件, a.监护人, a.查询密码, a.在院, a.锁定,
       a.户口地址, a.户口地址邮编, a.籍贯, a.Email, a.Qq, a.联系人身份证号, a.病人类型, a.主页id, a.手机号, c.名称 As 当前科室名称, d.名称 As 当前病区名称,
       a.单位地址 As 工作单位地址, f.名称 As 险类名称
      From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 保险类别 F
      Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.险类 = f.序号(+) And a.手机号 = v_手机号 And
            Decode(Nvl(n_查询住院状态, 0), 2, 2, Nvl(a.在院, 0)) = Nvl(n_查询住院状态, 0) And a.停用时间 Is Null;
  Elsif v_联系人身份证号 Is Not Null Then
    Open c_病人信息 For
      Select /*+cardinality(c,10)*/
       a.病人id, a.姓名, a.性别, a.年龄, a.出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号, a.就诊卡号, a.卡验证码, a.费别, a.医疗付款方式 As 医疗付款方式名称,
       b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.婚姻状况, a.家庭地址, a.家庭电话, a.家庭地址邮编, a.联系人姓名, a.联系人关系, a.联系人地址,
       a.联系人电话, a.合同单位id, a.工作单位 As 工作单位名称, a.单位电话, a.单位邮编, a.单位开户行, a.单位帐号, a.就诊时间, a.就诊状态, a.就诊诊室, a.住院次数, a.当前科室id,
       a.当前病区id, a.入院时间, a.出院时间, a.Ic卡号, a.健康号, a.险类, a.登记时间, a.停用时间, a.当前床号, a.医保号, a.其他证件, a.监护人, a.查询密码, a.在院, a.锁定,
       a.户口地址, a.户口地址邮编, a.籍贯, a.Email, a.Qq, a.联系人身份证号, a.病人类型, a.主页id, a.手机号, c.名称 As 当前科室名称, d.名称 As 当前病区名称,
       a.单位地址 As 工作单位地址, f.名称 As 险类名称
      From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 保险类别 F
      Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.险类 = f.序号(+) And
            a.联系人身份证号 = v_联系人身份证号 And Decode(Nvl(n_查询住院状态, 0), 2, 2, Nvl(a.在院, 0)) = Nvl(n_查询住院状态, 0) And
            a.停用时间 Is Null;
  Elsif v_就诊卡号 Is Not Null Then
    Open c_病人信息 For
      Select /*+cardinality(c,10)*/
       a.病人id, a.姓名, a.性别, a.年龄, a.出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号, a.就诊卡号, a.卡验证码, a.费别, a.医疗付款方式 As 医疗付款方式名称,
       b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.婚姻状况, a.家庭地址, a.家庭电话, a.家庭地址邮编, a.联系人姓名, a.联系人关系, a.联系人地址,
       a.联系人电话, a.合同单位id, a.工作单位 As 工作单位名称, a.单位电话, a.单位邮编, a.单位开户行, a.单位帐号, a.就诊时间, a.就诊状态, a.就诊诊室, a.住院次数, a.当前科室id,
       a.当前病区id, a.入院时间, a.出院时间, a.Ic卡号, a.健康号, a.险类, a.登记时间, a.停用时间, a.当前床号, a.医保号, a.其他证件, a.监护人, a.查询密码, a.在院, a.锁定,
       a.户口地址, a.户口地址邮编, a.籍贯, a.Email, a.Qq, a.联系人身份证号, a.病人类型, a.主页id, a.手机号, c.名称 As 当前科室名称, d.名称 As 当前病区名称,
       a.单位地址 As 工作单位地址, f.名称 As 险类名称
      From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 保险类别 F
      Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.险类 = f.序号(+) And a.就诊卡号 = v_就诊卡号 And
            Decode(Nvl(n_查询住院状态, 0), 2, 2, Nvl(a.在院, 0)) = Nvl(n_查询住院状态, 0) And a.停用时间 Is Null;
  
  Elsif v_姓名 Is Not Null Then
    --按姓名查找
    If Instr(v_姓名, '%') > 0 Then
      Open c_病人信息 For
        Select a.病人id, a.姓名, a.性别, a.年龄, a.出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号, a.就诊卡号, a.卡验证码, a.费别,
               a.医疗付款方式 As 医疗付款方式名称, b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.婚姻状况, a.家庭地址, a.家庭电话,
               a.家庭地址邮编, a.联系人姓名, a.联系人关系, a.联系人地址, a.联系人电话, a.合同单位id, a.工作单位 As 工作单位名称, a.单位电话, a.单位邮编, a.单位开户行, a.单位帐号,
               a.就诊时间, a.就诊状态, a.就诊诊室, a.住院次数, a.当前科室id, a.当前病区id, a.入院时间, a.出院时间, a.Ic卡号, a.健康号, a.险类, a.登记时间, a.停用时间,
               a.当前床号, a.医保号, a.其他证件, a.监护人, a.查询密码, a.在院, a.锁定, a.户口地址, a.户口地址邮编, a.籍贯, a.Email, a.Qq, a.联系人身份证号,
               a.病人类型, a.主页id, a.手机号, c.名称 As 当前科室名称, d.名称 As 当前病区名称, a.单位地址 As 工作单位地址, f.名称 As 险类名称
        From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 保险类别 F
        Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.险类 = f.序号(+) And a.姓名 Like v_姓名 And
              Decode(Nvl(n_查询住院状态, 0), 2, 2, Nvl(a.在院, 0)) = Nvl(n_查询住院状态, 0) And a.停用时间 Is Null And
              Decode(Nvl(n_查找天数, '0'), '0', '0', Nvl(a.就诊时间, a.登记时间) || '') >=
              Decode(Nvl(n_查找天数, '0'), '0', '0', Trunc(Sysdate - Nvl(n_查找天数, 0)) || '');
    
    Else
      Open c_病人信息 For
        Select a.病人id, a.姓名, a.性别, a.年龄, a.出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号, a.就诊卡号, a.卡验证码, a.费别,
               a.医疗付款方式 As 医疗付款方式名称, b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.婚姻状况, a.家庭地址, a.家庭电话,
               a.家庭地址邮编, a.联系人姓名, a.联系人关系, a.联系人地址, a.联系人电话, a.合同单位id, a.工作单位 As 工作单位名称, a.单位电话, a.单位邮编, a.单位开户行, a.单位帐号,
               a.就诊时间, a.就诊状态, a.就诊诊室, a.住院次数, a.当前科室id, a.当前病区id, a.入院时间, a.出院时间, a.Ic卡号, a.健康号, a.险类, a.登记时间, a.停用时间,
               a.当前床号, a.医保号, a.其他证件, a.监护人, a.查询密码, a.在院, a.锁定, a.户口地址, a.户口地址邮编, a.籍贯, a.Email, a.Qq, a.联系人身份证号,
               a.病人类型, a.主页id, a.手机号, c.名称 As 当前科室名称, d.名称 As 当前病区名称, a.单位地址 As 工作单位地址, f.名称 As 险类名称
        From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 保险类别 F
        Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.险类 = f.序号(+) And a.姓名 = v_姓名 And
              Decode(Nvl(n_查询住院状态, 0), 2, 2, Nvl(a.在院, 0)) = Nvl(n_查询住院状态, 0) And a.停用时间 Is Null;
    End If;
  Elsif v_卡号 Is Not Null Or v_二维码 Is Not Null Then
    If Nvl(n_卡类别id, 0) = 0 Then
      Select Max(ID) Into n_卡类别id From 医疗卡类别 Where 名称 = v_医疗卡名称;
    End If;
    If Nvl(n_卡类别id, 0) <> 0 And v_卡号 Is Not Null Then
      Open c_病人信息 For
        Select /*+cardinality(c,10)*/
         a.病人id, a.姓名, a.性别, a.年龄, a.出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号, a.就诊卡号, a.卡验证码, a.费别, a.医疗付款方式 As 医疗付款方式名称,
         b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.婚姻状况, a.家庭地址, a.家庭电话, a.家庭地址邮编, a.联系人姓名, a.联系人关系,
         a.联系人地址, a.联系人电话, a.合同单位id, a.工作单位 As 工作单位名称, a.单位电话, a.单位邮编, a.单位开户行, a.单位帐号, a.就诊时间, a.就诊状态, a.就诊诊室, a.住院次数,
         a.当前科室id, a.当前病区id, a.入院时间, a.出院时间, a.Ic卡号, a.健康号, a.险类, a.登记时间, a.停用时间, a.当前床号, a.医保号, a.其他证件, a.监护人, a.查询密码,
         a.在院, a.锁定, a.户口地址, a.户口地址邮编, a.籍贯, a.Email, a.Qq, a.联系人身份证号, a.病人类型, a.主页id, a.手机号, c.名称 As 当前科室名称,
         d.名称 As 当前病区名称, a.单位地址 As 工作单位地址, f.名称 As 险类名称, f.名称 As 险类名称
        From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 保险类别 F,
             (Select Distinct 病人id
               From (Select j.病人id,
                             Case
                                When Nvl(j.状态, 0) = 1 And
                                     (Nvl(m.有效天数, 0) = 0 Or Nvl(j.挂失时间, To_Date('3000-01-01', 'yyyy-mm-dd')) + Nvl(m.有效天数, 0) > Sysdate) Then
                                 1
                                Else
                                 Nvl(j.状态, 0)
                              End As 状态
                      From 病人医疗卡信息 J, 医疗卡挂失方式 M, 医疗卡类别 Q
                      Where j.挂失方式 = m.名称(+) And j.卡类别id = q.Id And Nvl(q.是否启用, 0) = 1 And j.卡类别id = n_卡类别id And
                            j.卡号 = v_卡号 And Sysdate < Nvl(j.终止使用时间, Sysdate + 1))
               Where 状态 = 0) Q
        Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.险类 = f.序号(+) And a.病人id = q.病人id And
              Decode(Nvl(n_查询住院状态, 0), 2, 2, Nvl(a.在院, 0)) = Nvl(n_查询住院状态, 0) And a.停用时间 Is Null;
      --0-正常有效卡;1-已挂失; 2-补卡停用
    Elsif Nvl(n_卡类别id, 0) <> 0 And v_二维码 Is Not Null Then
      Open c_病人信息 For
        Select /*+cardinality(c,10)*/
         a.病人id, a.姓名, a.性别, a.年龄, a.出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号, a.就诊卡号, a.卡验证码, a.费别, a.医疗付款方式 As 医疗付款方式名称,
         b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.婚姻状况, a.家庭地址, a.家庭电话, a.家庭地址邮编, a.联系人姓名, a.联系人关系,
         a.联系人地址, a.联系人电话, a.合同单位id, a.工作单位 As 工作单位名称, a.单位电话, a.单位邮编, a.单位开户行, a.单位帐号, a.就诊时间, a.就诊状态, a.就诊诊室, a.住院次数,
         a.当前科室id, a.当前病区id, a.入院时间, a.出院时间, a.Ic卡号, a.健康号, a.险类, a.登记时间, a.停用时间, a.当前床号, a.医保号, a.其他证件, a.监护人, a.查询密码,
         a.在院, a.锁定, a.户口地址, a.户口地址邮编, a.籍贯, a.Email, a.Qq, a.联系人身份证号, a.病人类型, a.主页id, a.手机号, c.名称 As 当前科室名称,
         d.名称 As 当前病区名称, a.单位地址 As 工作单位地址, f.名称 As 险类名称
        From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 保险类别 F,
             (Select Distinct 病人id
               From (Select j.病人id,
                             Case
                                When Nvl(j.状态, 0) = 1 And
                                     (Nvl(m.有效天数, 0) = 0 Or Nvl(j.挂失时间, To_Date('3000-01-01', 'yyyy-mm-dd')) + Nvl(m.有效天数, 0) > Sysdate) Then
                                 1
                                Else
                                 Nvl(j.状态, 0)
                              End As 状态
                      From 病人医疗卡信息 J, 医疗卡挂失方式 M, 医疗卡类别 Q
                      Where j.挂失方式 = m.名称(+) And j.卡类别id = q.Id And Nvl(q.是否启用, 0) = 1 And j.卡类别id = n_卡类别id And
                            j.二维码 = v_二维码 And Sysdate < Nvl(j.终止使用时间, Sysdate + 1))
               Where 状态 = 0) Q
        Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.险类 = f.序号(+) And a.病人id = q.病人id And
              Decode(Nvl(n_查询住院状态, 0), 2, 2, Nvl(a.在院, 0)) = Nvl(n_查询住院状态, 0) And a.停用时间 Is Null;
      --0-正常有效卡;1-已挂失; 2-补卡停用
    Elsif v_卡号 Is Not Null Then
      Open c_病人信息 For
        Select /*+cardinality(c,10)*/
         a.病人id, a.姓名, a.性别, a.年龄, a.出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号, a.就诊卡号, a.卡验证码, a.费别, a.医疗付款方式 As 医疗付款方式名称,
         b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.婚姻状况, a.家庭地址, a.家庭电话, a.家庭地址邮编, a.联系人姓名, a.联系人关系,
         a.联系人地址, a.联系人电话, a.合同单位id, a.工作单位 As 工作单位名称, a.单位电话, a.单位邮编, a.单位开户行, a.单位帐号, a.就诊时间, a.就诊状态, a.就诊诊室, a.住院次数,
         a.当前科室id, a.当前病区id, a.入院时间, a.出院时间, a.Ic卡号, a.健康号, a.险类, a.登记时间, a.停用时间, a.当前床号, a.医保号, a.其他证件, a.监护人, a.查询密码,
         a.在院, a.锁定, a.户口地址, a.户口地址邮编, a.籍贯, a.Email, a.Qq, a.联系人身份证号, a.病人类型, a.主页id, a.手机号, c.名称 As 当前科室名称,
         d.名称 As 当前病区名称, a.单位地址 As 工作单位地址, f.名称 As 险类名称
        From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 保险类别 F,
             (Select Distinct 病人id
               From (Select j.病人id,
                             Case
                                When Nvl(j.状态, 0) = 1 And
                                     (Nvl(m.有效天数, 0) = 0 Or Nvl(j.挂失时间, To_Date('3000-01-01', 'yyyy-mm-dd')) + Nvl(m.有效天数, 0) > Sysdate) Then
                                 1
                                Else
                                 Nvl(j.状态, 0)
                              End As 状态
                      From 病人医疗卡信息 J, 医疗卡挂失方式 M, 医疗卡类别 Q
                      Where j.挂失方式 = m.名称(+) And j.卡类别id = q.Id And Nvl(q.是否启用, 0) = 1 And j.卡号 = v_卡号 And
                            Sysdate < Nvl(j.终止使用时间, Sysdate + 1))
               Where 状态 = 0) Q
        Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.险类 = f.序号(+) And a.病人id = q.病人id And
              Decode(Nvl(n_查询住院状态, 0), 2, 2, Nvl(a.在院, 0)) = Nvl(n_查询住院状态, 0) And a.停用时间 Is Null;
    
    Else
      Open c_病人信息 For
        Select /*+cardinality(c,10)*/
         a.病人id, a.姓名, a.性别, a.年龄, a.出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号, a.就诊卡号, a.卡验证码, a.费别, a.医疗付款方式 As 医疗付款方式名称,
         b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.婚姻状况, a.家庭地址, a.家庭电话, a.家庭地址邮编, a.联系人姓名, a.联系人关系,
         a.联系人地址, a.联系人电话, a.合同单位id, a.工作单位 As 工作单位名称, a.单位电话, a.单位邮编, a.单位开户行, a.单位帐号, a.就诊时间, a.就诊状态, a.就诊诊室, a.住院次数,
         a.当前科室id, a.当前病区id, a.入院时间, a.出院时间, a.Ic卡号, a.健康号, a.险类, a.登记时间, a.停用时间, a.当前床号, a.医保号, a.其他证件, a.监护人, a.查询密码,
         a.在院, a.锁定, a.户口地址, a.户口地址邮编, a.籍贯, a.Email, a.Qq, a.联系人身份证号, a.病人类型, a.主页id, a.手机号, c.名称 As 当前科室名称,
         d.名称 As 当前病区名称, a.单位地址 As 工作单位地址, f.名称 As 险类名称
        From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 保险类别 F,
             
             (Select Distinct 病人id
               From (Select j.病人id,
                             Case
                               When Nvl(j.状态, 0) = 1 And
                                    (Nvl(m.有效天数, 0) = 0 Or
                                     Nvl(j.挂失时间, To_Date('3000-01-01', 'yyyy-mm-dd')) + Nvl(m.有效天数, 0) > Sysdate) Then
                                1
                               Else
                                Nvl(j.状态, 0)
                             End As 状态
                      From 病人医疗卡信息 J, 医疗卡挂失方式 M, 医疗卡类别 Q
                      Where j.挂失方式 = m.名称(+) And j.卡类别id = q.Id And Nvl(q.是否启用, 0) = 1 And j.二维码 = v_二维码 And
                            Sysdate < Nvl(j.终止使用时间, Sysdate + 1))
               Where 状态 = 0) Q
        Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.险类 = f.序号(+) And a.病人id = q.病人id And
              Decode(Nvl(n_查询住院状态, 0), 2, 2, Nvl(a.在院, 0)) = Nvl(n_查询住院状态, 0) And a.停用时间 Is Null;
    End If;
  Elsif Nvl(n_科室id, 0) <> 0 Then
    --按当前科室查找
    Open c_病人信息 For
      Select a.病人id, a.姓名, a.性别, a.年龄, a.出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号, a.就诊卡号, a.卡验证码, a.费别, a.医疗付款方式 As 医疗付款方式名称,
             b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.婚姻状况, a.家庭地址, a.家庭电话, a.家庭地址邮编, a.联系人姓名, a.联系人关系,
             a.联系人地址, a.联系人电话, a.合同单位id, a.工作单位 As 工作单位名称, a.单位电话, a.单位邮编, a.单位开户行, a.单位帐号, a.就诊时间, a.就诊状态, a.就诊诊室,
             a.住院次数, a.当前科室id, a.当前病区id, a.入院时间, a.出院时间, a.Ic卡号, a.健康号, a.险类, a.登记时间, a.停用时间, a.当前床号, a.医保号, a.其他证件,
             a.监护人, a.查询密码, a.在院, a.锁定, a.户口地址, a.户口地址邮编, a.籍贯, a.Email, a.Qq, a.联系人身份证号, a.病人类型, a.主页id, a.手机号,
             c.名称 As 当前科室名称, d.名称 As 当前病区名称, a.单位地址 As 工作单位地址, f.名称 As 险类名称
      From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 保险类别 F
      Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.险类 = f.序号(+) And
            Decode(Nvl(n_查询住院状态, 0), 2, 2, Nvl(a.在院, 0)) = Nvl(n_查询住院状态, 0) And a.当前科室id = n_科室id And a.停用时间 Is Null;
  Else
    v_Err_Msg := '未传入有效的查询条件，不能获取病人信息!';
    Json_Out  := '{"output":{"code":0,"message":"' || Zljsonstr(v_Err_Msg) || '"}}';
    Return;
  End If;

  Json_Out := '{"output":{"code":1,"message":"成功","pati_list":[';

  v_Json      := '';
  n_Firstitem := 1;
  Loop
    Fetch c_病人信息
      Into r_病人;
    Exit When c_病人信息 %NotFound;
    If Nvl(n_Firstitem, 0) = 0 Then
      v_Json := v_Json || ',';
    Else
      n_Firstitem := 0;
    End If;
  
    v_Json := v_Json || '{';
    --1.取基本信息
    v_Json := v_Json || '"pati_id":' || Nvl(r_病人.病人id, 0);
    v_Json := v_Json || ',"pati_pageid":' || Nvl(r_病人.主页id, 0);
    v_Json := v_Json || ',"pati_name":"' || Zljsonstr(r_病人.姓名) || '"';
    v_Json := v_Json || ',"pati_sex":"' || Zljsonstr(r_病人.性别) || '"';
    v_Json := v_Json || ',"pati_age":"' || Zljsonstr(r_病人.年龄) || '"';
  
    v_Json := v_Json || ',"pati_birthdate":"' || Zljsonstr(To_Char(r_病人.出生日期, 'yyyy-mm-dd hh24:mi:ss')) || '"';
    v_Json := v_Json || ',"fee_category":"' || Zljsonstr(r_病人.费别) || '"';
    If Nvl(r_病人.门诊号, 0) = 0 Then
      v_Json := v_Json || ',"outpatient_num":null';
    Else
      v_Json := v_Json || ',"outpatient_num":"' || r_病人.门诊号 || '"';
    End If;
    If Nvl(r_病人.住院号, 0) = 0 Then
      v_Json := v_Json || ',"inpatient_num":null';
    Else
      v_Json := v_Json || ',"inpatient_num":"' || r_病人.住院号 || '"';
    End If;
    v_Json := v_Json || ',"pati_nation":"' || Zljsonstr(r_病人.民族) || '"';
    v_Json := v_Json || ',"mdlpay_mode_name":"' || Zljsonstr(r_病人.医疗付款方式名称) || '"';
    v_Json := v_Json || ',"mdlpay_mode_code":"' || Zljsonstr(r_病人.医疗付款方式编码) || '"';
    v_Json := v_Json || ',"insurance_num":"' || Zljsonstr(r_病人.医保号) || '"';
    v_Json := v_Json || ',"pati_idcard":"' || Zljsonstr(r_病人.身份证号) || '"';
    v_Json := v_Json || ',"vcard_no":"' || Zljsonstr(r_病人.就诊卡号) || '"';
  
    v_Json := v_Json || ',"iccard_no":"' || Zljsonstr(r_病人.Ic卡号) || '"';
    v_Json := v_Json || ',"inp_times":' || Nvl(r_病人.住院次数, 0);
    v_Json := v_Json || ',"pati_education":"' || Zljsonstr(r_病人.学历) || '"';
    v_Json := v_Json || ',"ocpt_name":"' || Zljsonstr(r_病人.职业) || '"';
    v_Json := v_Json || ',"pati_marital_cstatus":"' || Zljsonstr(r_病人.婚姻状况) || '"';
  
    v_Json := v_Json || ',"phone_number":"' || Zljsonstr(r_病人.手机号) || '"';
    v_Json := v_Json || ',"pati_bed":"' || Zljsonstr(r_病人.当前床号) || '"';
    v_Json := v_Json || ',"pati_birthplace":"' || Zljsonstr(r_病人.出生地点) || '"';
    v_Json := v_Json || ',"pat_home_addr":"' || Zljsonstr(r_病人.家庭地址) || '"';
    v_Json := v_Json || ',"pat_home_phno":"' || Zljsonstr(r_病人.家庭电话) || '"';
  
    v_Json := v_Json || ',"insurance_type":' || Nvl(r_病人.险类, 0);
    v_Json := v_Json || ',"insurance_name":"' || Zljsonstr(r_病人.险类名称) || '"';
    v_Json := v_Json || ',"is_inhspt":' || Nvl(r_病人.在院, 0);
    v_Json := v_Json || ',"pati_type":"' || Zljsonstr(r_病人.病人类型) || '"';
  
    --2.查询基本信息+联系人信息
    If Nvl(n_查询类型, 0) >= 1 Then
      --查询类型:如：0-基本;1-基本+联系人;2-所有
      v_Json := v_Json || ',"contacts_name":"' || Zljsonstr(r_病人.联系人姓名) || '"';
      v_Json := v_Json || ',"contacts_relation":"' || Zljsonstr(r_病人.联系人关系) || '"';
      v_Json := v_Json || ',"contacts_idcard":"' || Zljsonstr(r_病人.联系人身份证号) || '"';
      v_Json := v_Json || ',"contacts_addr":"' || Zljsonstr(r_病人.联系人地址) || '"';
      v_Json := v_Json || ',"contacts_phno":"' || Zljsonstr(r_病人.联系人电话) || '"';
    End If;
  
    --3.查询类型所有:如：0-基本;1-基本+联系人;2-所有
    If Nvl(n_查询类型, 0) > 1 Then
      v_Json := v_Json || ',"pati_wardarea_id":' || Nvl(r_病人.当前病区id, 0);
      v_Json := v_Json || ',"pati_wardarea_name":"' || Zljsonstr(r_病人.当前病区名称) || '"';
      v_Json := v_Json || ',"pati_dept_id":' || Nvl(r_病人.当前科室id, 0);
      v_Json := v_Json || ',"pati_dept_name":"' || Zljsonstr(r_病人.当前科室名称) || '"';
      v_Json := v_Json || ',"health_num":"' || Zljsonstr(r_病人.健康号) || '"';
    
      v_Json := v_Json || ',"pati_identity":"' || Zljsonstr(r_病人.身份) || '"';
      v_Json := v_Json || ',"ntvplc_name":"' || Zljsonstr(r_病人.籍贯) || '"';
      v_Json := v_Json || ',"country_name":"' || Zljsonstr(r_病人.国籍) || '"';
      v_Json := v_Json || ',"pat_home_postcode":"' || Zljsonstr(r_病人.家庭地址邮编) || '"';
      v_Json := v_Json || ',"pati_area":"' || Zljsonstr(r_病人.区域) || '"';
    
      v_Json := v_Json || ',"pat_hous_addr":"' || Zljsonstr(r_病人.户口地址) || '"';
      v_Json := v_Json || ',"pat_hous_postcode":"' || Zljsonstr(r_病人.户口地址邮编) || '"';
      v_Json := v_Json || ',"emp_addr":"' || Zljsonstr(r_病人.工作单位地址) || '"';
      v_Json := v_Json || ',"emp_name":"' || Zljsonstr(r_病人.工作单位名称) || '"';
      v_Json := v_Json || ',"emp_phno":"' || Zljsonstr(r_病人.单位电话) || '"';
    
      v_Json := v_Json || ',"emp_postcode":"' || Zljsonstr(r_病人.单位邮编) || '"';
      v_Json := v_Json || ',"emp_bank_name":"' || Zljsonstr(r_病人.单位开户行) || '"';
      v_Json := v_Json || ',"emp_bank_accnum":"' || Zljsonstr(r_病人.单位帐号) || '"';
      v_Json := v_Json || ',"ctt_unit_id":' || Nvl(r_病人.合同单位id, 0);
      v_Json := v_Json || ',"adta_time":"' || Zljsonstr(To_Char(r_病人.入院时间, 'yyyy-mm-dd hh24:mi:ss')) || '"';
    
      v_Json := v_Json || ',"adtd_time":"' || Zljsonstr(To_Char(r_病人.出院时间, 'yyyy-mm-dd hh24:mi:ss')) || '"';
      v_Json := v_Json || ',"pat_grdn_name":"' || Zljsonstr(r_病人.监护人) || '"';
      v_Json := v_Json || ',"cert_no_other":"' || Zljsonstr(r_病人.其他证件) || '"';
      v_Json := v_Json || ',"pati_email":"' || Zljsonstr(r_病人.Email) || '"';
      v_Json := v_Json || ',"pati_qq":"' || Zljsonstr(r_病人.Qq) || '"';
      v_Json := v_Json || ',"card_captcha":"' || Zljsonstr(r_病人.卡验证码) || '"';
    
      n_颜色 := Null;
      If r_病人.病人类型 Is Not Null Then
        Select Max(颜色) Into n_颜色 From 病人类型 Where 名称 = r_病人.病人类型;
      End If;
      v_Json := v_Json || ',"pati_show_color":' || Nvl(n_颜色, 0);
      v_Json := v_Json || ',"visit_room":"' || Zljsonstr(r_病人.就诊诊室) || '"';
      v_Json := v_Json || ',"visit_statu":' || Nvl(r_病人.就诊状态, 0);
      v_Json := v_Json || ',"visit_time":"' || Zljsonstr(To_Char(r_病人.就诊时间, 'yyyy-mm-dd hh24:mi:ss')) || '"';
      v_Json := v_Json || ',"create_time":"' || Zljsonstr(To_Char(r_病人.登记时间, 'yyyy-mm-dd hh24:mi:ss')) || '"';
    
      If Nvl(r_病人.险类, 0) <> 0 And Nvl(n_包含医保密码, 0) = 1 Then
        Select Max(d.密码)
        Into v_医保密码
        From 医保病人档案 D, 医保病人关联表 E
        Where e.病人id = r_病人.病人id And e.险类 = r_病人.险类 And e.医保号 = r_病人.医保号 And e.标志 = 1 And e.医保号 = d.医保号(+) And
              e.险类 = d.险类(+) And e.中心 = d.中心(+);
      Else
        v_医保密码 := '';
      End If;
      v_Json := v_Json || ',"insurance_pwd":"' || Zljsonstr(v_医保密码) || '"';
    End If;
  
    --获取家属关系
    If Nvl(n_是否包含家属, 0) = 1 Then
      v_Temp         := '';
      n_Firstsubitem := 1;
      For c_家属 In (Select 家属id, 关系 From 病人家属 Where 病人id = r_病人.病人id And Nvl(撤档时间, Sysdate) <= Sysdate) Loop
        If Nvl(n_Firstsubitem, 0) = 0 Then
          v_Temp := v_Temp || ',';
        Else
          n_Firstsubitem := 0;
        End If;
      
        v_Temp := v_Temp || '{';
        v_Temp := v_Temp || '"family_id":' || c_家属.家属id;
        v_Temp := v_Temp || ',"family_relation":"' || Zljsonstr(c_家属.关系) || '"';
        v_Temp := v_Temp || '}';
      End Loop;
      v_Json := v_Json || ',"family_list":[' || v_Temp || ']';
    End If;
  
    --获取过敏药物
    If Nvl(n_是否包含过敏药物, 0) = 1 Then
      v_Temp         := '';
      n_Firstsubitem := 1;
      For c_过敏药物 In (Select 过敏药物id, 过敏药物, 过敏反应 From 病人过敏药物 Where 病人id = r_病人.病人id) Loop
        If Nvl(n_Firstsubitem, 0) = 0 Then
          v_Temp := v_Temp || ',';
        Else
          n_Firstsubitem := 0;
        End If;
      
        v_Temp := v_Temp || '{';
        v_Temp := v_Temp || '"pat_algc_cadn_id":' || Zljsonstr(c_过敏药物.过敏药物id, 1);
        v_Temp := v_Temp || ',"pat_algc_cadn":"' || Zljsonstr(c_过敏药物.过敏药物) || '"';
        v_Temp := v_Temp || ',"allergy_info":"' || Zljsonstr(c_过敏药物.过敏反应) || '"';
        v_Temp := v_Temp || '}';
      End Loop;
      v_Json := v_Json || ',"drug_list":[' || v_Temp || ']';
    End If;
  
    -- 获取病人免疫信息
    If Nvl(n_是否包含免疫信息, 0) = 1 Then
      v_Temp         := '';
      n_Firstsubitem := 1;
      For c_免疫记录 In (Select To_Char(接种时间, 'yyyy-mm-dd hh24:mi:ss') As 接种时间, 接种名称
                     From 病人免疫记录
                     Where 病人id = r_病人.病人id) Loop
        If Nvl(n_Firstsubitem, 0) = 0 Then
          v_Temp := v_Temp || ',';
        Else
          n_Firstsubitem := 0;
        End If;
      
        v_Temp := v_Temp || '{';
        v_Temp := v_Temp || '"vaccinate_time":"' || Zljsonstr(c_免疫记录.接种时间) || '"';
        v_Temp := v_Temp || ',"vaccinate_name":"' || Zljsonstr(c_免疫记录.接种名称) || '"';
        v_Temp := v_Temp || '}';
      End Loop;
      v_Json := v_Json || ',"immune_list":[' || v_Temp || ']';
    End If;
  
    --获取卡信息
    If Nvl(n_是否包含卡信息, 0) = 1 Then
      v_Temp         := '';
      n_Firstsubitem := 1;
      For c_医疗卡 In (Select Distinct 卡类别id, 卡号, 密码
                    From (Select j.卡类别id, j.卡号, j.密码,
                                  Case
                                    When Nvl(j.状态, 0) = 1 And
                                         (Nvl(m.有效天数, 0) = 0 Or
                                          Nvl(j.挂失时间, To_Date('3000-01-01', 'yyyy-mm-dd')) + Nvl(m.有效天数, 0) > Sysdate) Then
                                     1
                                    Else
                                     Nvl(j.状态, 0)
                                  End As 状态
                           From 病人医疗卡信息 J, 医疗卡挂失方式 M, 医疗卡类别 Q
                           Where j.挂失方式 = m.名称(+) And j.卡类别id = q.Id And Nvl(q.是否启用, 0) = 1 And j.病人id = r_病人.病人id And
                                 Decode(Nvl(n_卡类别id, 0), 0, j.卡类别id) = Nvl(n_卡类别id, 0) And
                                 Sysdate < Nvl(j.终止使用时间, Sysdate + 1))
                    Where 状态 = 0) Loop
        If Nvl(n_Firstsubitem, 0) = 0 Then
          v_Temp := v_Temp || ',';
        Else
          n_Firstsubitem := 0;
        End If;
      
        v_Temp := v_Temp || '{';
        v_Temp := v_Temp || '"cardtype_id":' || c_医疗卡.卡类别id;
        v_Temp := v_Temp || ',"card_no":' || Zljsonstr(c_医疗卡.卡号) || '"';
        v_Temp := v_Temp || ',"card_pwd":' || Zljsonstr(c_医疗卡.密码) || '"';
        v_Temp := v_Temp || '}';
      End Loop;
      v_Json := v_Json || ',"card_list":[' || v_Temp || ']';
    End If;
    v_Json := v_Json || '}';
  
    If Length(v_Json) > 20000 Then
      Json_Out := Json_Out || v_Json;
      v_Json   := '';
    End If;
  End Loop;
  Close c_病人信息;
  Json_Out := Json_Out || v_Json || ']}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getpatiinfo;
/
Create Or Replace Procedure Zl_Patisvr_Getpatiinfsbyrange
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:获取病人信息
  --入参：Json_In:格式
  --    input
  --      query_type          N 1 0：查询基本信息；1：查询基本信息+扩展信息
  --      pati_ids            C   病人IDs:多个用逗号
  --      pati_name           C   姓名:可以代%分号表表按姓名匹配
  --      pati_sex            C   性别
  --      pati_age            C   年龄
  --      birthdate_start     C   开始出生日期
  --      birthdate_end       C   终止出生日期
  --      outpatient_num      C   门诊号
  --      pati_idcard         C   身份证号
  --      fee_category        C   费别
  --      pati_area           C   区域
  --      insurance_num       C   医保号
  --      vcard_no            C   就诊卡号
  --      iccard_no           C   Ic卡号
  --      wardarea_ids        C   病区ids：多个用逗号
  --      qurey_max           N   查询的最大记录数，为0或NULL时表示不限制
  --      qrspt_statu         N   查询住院状态:0-仅门诊;1-在院 ;2-门诊及在院
  --      visit_star_time     C   就诊开始时间:yyyy-mm-dd hh24:mi:ss
  --      visit_end_time      C   就诊结束时间:yyyy-mm-dd hh24:mi:ss
  --      create_start_time   C   开始登记时间:yyyy-mm-dd hh24:mi:ss
  --      create_end_time     C   终止登记时间:yyyy-mm-dd hh24:mi:ss
  --      occasion            N   场合:调用Zl_Custom_Patiids_Get(根据身份证返回病人id)服务时需传入
  --      only_ctorg_pati     N   只查询合约单位的病人
  --      ctt_unit_id         N   合同单位id,只查询合约单位的病人时有效
  --      default_cardtype_id N   缺省卡类别id
  --      dept_ids            C   科室ids:多个用逗号分割
  --      mdlpay_mode_name    C   医疗付款方式
  --      phone_number        C   手机号
  --      is_stop             N   是否显示停用
  --      pati_similar        C   相似条件
  --        pati_name         C 1 姓名
  --        pati_sex          C 1 性别
  --        country_name      C 1 国籍
  --        pati_nation       C 1 民族
  --        pati_birthdate    C 1 出生日期：yyyy-mm-dd hh24:mi:ss
  --        pati_idcard       C 1 身份证号
  --出参      json
  --output
  -- code                     N 1 应答码：0-失败；1-成功
  -- message                  C 1 应答消息： 失败时返回具体的错误信息
  -- pati_list[]                  病人信息列表
  --   pati_id                N 1 病人id
  --   pati_pageid            N 1 主页id：病人信息.主页ID
  --   pati_name              C 1 姓名
  --   pati_sex               C 1 性别
  --   pati_age               C 1 年龄
  --   pati_birthdate         C 1 出生日期：yyyy-mm-dd hh24:mi:ss
  --   pati_birthplace        C 1 出生地点
  --   fee_category           C 1 费别
  --   outpatient_num         C 1 门诊号
  --   inpatient_num          C 1 住院号
  --   inp_times              N 1 住院次数
  --   pati_nation            C 1 民族
  --   pati_idcard            C 1 身份证号
  --   vcard_no               C 1 就诊卡号
  --   phone_number           C 1 手机号
  --   pat_home_phno          C 1 家庭电话
  --   pati_education         C 1 学历
  --   ocpt_name              C 1 职业
  --   pati_identity          C 1 身份
  --   country_name           C 1 国籍
  --   pat_home_addr          C 1 家庭地址
  --   pati_area              C 1 区域
  --   emp_name               C 1 工作单位名称
  --   pati_bed               C 1 当前床号
  --   is_inhspt              N 1 是否在院：1-在院；0-不在院
  --   pati_type              C 1 病人类型(普通，医保，留观)
  --   insurance_type         C 1 险类
  --   insurance_type_name    C 1 险类名称
  --   pati_wardarea_id       N 1 当前病区id
  --   pati_wardarea_name     C 1 当前病区名称
  --   pati_dept_id           N 1 当前科室id
  --   pati_dept_name         C 1 当前科室名称
  --   adta_time              C 1 入院时间:yyyy-mm-dd hh24:mi:ss
  --   adtd_time              C 1 出院时间:yyyy-mm-dd hh24:mi:ss
  --   create_time            C 1 登记时间:yyyy-mm-dd hh24:mi:ss
  --   medc_card_no           C   医疗卡号：当入参节点default_cardtype_id不为空时，才返回
  --   visit_time             C 1 就诊时间:yyyy-mm-dd hh24:mi:ss
  --   ctt_unit_id            N   合同单位id
  --   mdlpay_mode_name       C 1 医疗付款方式名称
  --   mdlpay_mode_code       C 1 医疗付款方式编码
  --   stop_time              C  停用时间
  --   insurance_num          C  医保号
  --   emp_addr               C  单位地址
  --   contacts               C   联系人信息节点
  --     name                 C 1 联系人姓名
  --     phone                C 1 联系人电话
  ---------------------------------------------------------------------------
  v_Err_Msg      Varchar2(500);
  j_Jsonin       Pljson;
  j_Json         Pljson;
  j_Json_Similar Pljson;
  c_病人ids      Clob;
  v_List         Varchar2(32767);
  v_Listtmp      Varchar2(32767);
  n_查询类型     Number(1);
  v_姓名         Varchar2(200);
  v_性别         Varchar2(50);
  v_病人ids      Varchar2(3000);
  d_开始出生日期 Date;
  d_终止出生日期 Date;
  d_出生日期     Date;
  n_门诊号       Number(18);
  v_身份证号     Varchar2(50);
  v_费别         Varchar2(50);

  v_区域         Varchar2(100);
  v_就诊卡号     Varchar2(200);
  v_Ic卡号       Varchar2(200);
  d_就诊开始时间 Date;
  d_就诊结束时间 Date;
  n_Like         Number(2);
  n_Max          Number(10);
  d_开始登记时间 Date;
  d_结束登记时间 Date;
  v_病区ids      Varchar2(32680);
  n_查询住院状态 Number(2);
  v_医保号       Varchar2(200);
  n_场合         Number(20);
  n_仅合约单位   Number(1);
  n_缺省卡类别id 病人医疗卡信息.卡类别id%Type;
  l_病人id       t_Strlist := t_Strlist();
  v_科室ids      Varchar2(32680);
  v_医疗付款方式 Varchar2(100);
  v_手机号       Varchar2(100);
  v_年龄         病人信息.年龄%Type;
  n_合同单位id   病人信息.合同单位id%Type;
  n_是否停用     Number;
  v_国籍         病人信息.国籍%Type;
  v_民族         病人信息.民族%Type;

  Cursor c_病人基本信息 Is
    Select a.病人id, a.姓名, a.性别, a.年龄, To_Char(a.出生日期, 'yyyy-mm-dd hh24:mi:ss') As 出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号,
           a.就诊卡号, a.费别, a.医疗付款方式 As 医疗付款方式名称, b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.家庭地址, a.工作单位,
           a.住院次数, a.当前科室id, a.当前病区id, To_Char(a.入院时间, 'yyyy-mm-dd hh24:mi:ss') As 入院时间,
           To_Char(a.出院时间, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.险类, To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间,
           a.在院, a.当前床号, a.病人类型, a.主页id, c.名称 As 当前科室名称, d.名称 As 当前病区名称, Nvl(e.名称, a.工作单位) As 工作单位名称, f.卡号,
           To_Char(a.就诊时间, 'yyyy-mm-dd hh24:mi:ss') As 就诊时间, a.手机号, a.家庭电话, e.Id As 合约单位id,
           To_Char(a.停用时间, 'yyyy-mm-dd hh24:mi:ss') As 停用时间, a.医保号, a.联系人姓名, a.联系人电话, a.单位地址, x.名称 As 险类名称
    From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 合约单位 E, 病人医疗卡信息 F, 保险类别 X
    Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.合同单位id = e.Id(+) And a.停用时间 Is Null And
          a.病人id = f.病人id(+) And a.险类 = x.序号(+) And Rownum < 1;
  r_病人 c_病人基本信息%RowType;

  Type Ty_病人信息 Is Ref Cursor;
  c_病人信息 Ty_病人信息; --动态游标变量

  --组装失败时返回的数据
  Function Get_Err_Message(Message_In Varchar2) Return Varchar2 Is
    j_Out Varchar2(32767);
  Begin
    j_Out := '{"output":{"code":0,"message":"' || Zljsonstr(Message_In) || '"}}';
    Return j_Out;
  End Get_Err_Message;

Begin
  --解析入参
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_查询类型 := Nvl(j_Json.Get_Number('query_type'), 0);
  Begin
    c_病人ids := j_Json.Get_Clob('pati_ids');
  Exception
    When Others Then
      c_病人ids := Null;
  End;

  d_开始出生日期 := To_Date(j_Json.Get_String('birthdate_start'), 'YYYY-MM-DD hh24:mi:ss');
  d_终止出生日期 := To_Date(j_Json.Get_String('birthdate_end'), 'YYYY-MM-DD hh24:mi:ss');
  n_门诊号       := To_Number(j_Json.Get_String('outpatient_num'));
  v_身份证号     := j_Json.Get_String('pati_idcard');
  v_费别         := j_Json.Get_String('fee_category');
  v_性别         := j_Json.Get_String('pati_sex');
  v_区域         := j_Json.Get_String('pati_area');
  v_就诊卡号     := j_Json.Get_String('vcard_no');
  v_Ic卡号       := j_Json.Get_String('iccard_no');
  d_就诊开始时间 := To_Date(j_Json.Get_String('visit_start_time'), 'yyyy-mm-dd hh24:mi:ss');
  d_就诊结束时间 := To_Date(j_Json.Get_String('visit_end_time'), 'yyyy-mm-dd hh24:mi:ss');

  d_开始登记时间 := To_Date(j_Json.Get_String('create_start_time'), 'yyyy-mm-dd hh24:mi:ss');
  d_结束登记时间 := To_Date(j_Json.Get_String('create_end_time'), 'yyyy-mm-dd hh24:mi:ss');
  v_姓名         := j_Json.Get_String('pati_name');
  v_病区ids      := j_Json.Get_String('wardarea_ids');
  n_查询住院状态 := Nvl(j_Json.Get_Number('qrspt_statu'), 0);
  n_Max          := j_Json.Get_Number('qurey_Max');
  v_医保号       := j_Json.Get_String('insurance_num');
  n_场合         := j_Json.Get_Number('occasion');
  n_缺省卡类别id := j_Json.Get_Number('default_cardtype_id');
  n_仅合约单位   := Nvl(j_Json.Get_Number('only_ctorg_pati'), 0);
  n_合同单位id   := j_Json.Get_Number('ctt_unit_id');
  v_科室ids      := j_Json.Get_Number('dept_ids');
  v_医疗付款方式 := j_Json.Get_String('mdlpay_mode_name');
  v_手机号       := j_Json.Get_String('phone_number');
  v_年龄         := j_Json.Get_String('pati_age');
  n_是否停用     := j_Json.Get_Number('is_stop');
  ---相似病人查询条件
  --      pati_similar        C   相似条件
  --        pati_name         C 1 姓名
  --        pati_sex          C 1 性别
  --        country_name      C 1 国籍
  --        pati_nation       C 1 民族
  --        pati_birthdate    C 1 出生日期：yyyy-mm-dd hh24:mi:ss
  --        pati_idcard       C 1 身份证号
  j_Json_Similar := j_Json.Get_Pljson('pati_similar');
  If Not j_Json_Similar Is Null Then
    v_姓名     := j_Json_Similar.Get_String('pati_name');
    v_性别     := j_Json_Similar.Get_String('pati_sex');
    v_国籍     := j_Json_Similar.Get_String('country_name');
    v_民族     := j_Json_Similar.Get_String('pati_nation');
    d_出生日期 := To_Date(j_Json_Similar.Get_String('birthdate_start'), 'YYYY-MM-DD hh24:mi:ss');
    v_身份证号 := j_Json_Similar.Get_String('pati_idcard');
  End If;

  If v_病区ids Is Not Null Then
    v_病区ids := ',' || v_病区ids || ',';
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
      l_病人id(l_病人id.Count) := Substr(c_病人ids, 1, Instr(c_病人ids, ',', 3980) - 1);
      c_病人ids := Substr(c_病人ids, Instr(c_病人ids, ',', 3980) + 1);
    End If;
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"成功","pati_list":[';
  If l_病人id.Count <> 0 Then
    --0-仅门诊;1-在院 ;2-门诊及在院
    For I In 1 .. l_病人id.Count Loop
      For r_病人 In (With c_病人 As
                      (Select Column_Value As 病人id From Table(f_Num2list(l_病人id(I))))
                     Select /*+cardinality(B,10)*/
                      a.病人id, a.姓名, a.性别, a.年龄, To_Char(a.出生日期, 'yyyy-mm-dd hh24:mi:ss') As 出生日期, a.出生地点, a.身份证号, a.门诊号,
                      a.住院号, a.就诊卡号, a.费别, a.医疗付款方式 As 医疗付款方式名称, b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历,
                      a.家庭地址, a.工作单位, a.住院次数, a.当前科室id, a.当前病区id, To_Char(a.入院时间, 'yyyy-mm-dd hh24:mi:ss') As 入院时间,
                      To_Char(a.出院时间, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.险类,
                      To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间, a.在院, a.当前床号, a.病人类型, a.主页id, c.名称 As 当前科室名称,
                      d.名称 As 当前病区名称, Nvl(e.名称, a.工作单位) As 工作单位名称, f.卡号, To_Char(a.就诊时间, 'yyyy-mm-dd hh24:mi:ss') As 就诊时间,
                      a.手机号, a.家庭电话, e.Id As 合约单位id, To_Char(a.停用时间, 'yyyy-mm-dd hh24:mi:ss') As 停用时间, a.医保号, a.联系人姓名,
                      a.联系人电话, a.单位地址, x.名称 As 险类名称
                     From 病人信息 A, 医疗付款方式 B, c_病人 B, 部门表 C, 部门表 D, 合约单位 E, 病人医疗卡信息 F, 保险类别 X
                     Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.合同单位id = e.Id(+) And
                           a.险类 = x.序号(+) And ((a.停用时间 Is Null And Nvl(n_是否停用, 0) = 0) Or Nvl(n_是否停用, 0) = 1) And
                           (n_仅合约单位 = 0 Or a.合同单位id Is Not Null) And a.病人id = f.病人id(+) And f.卡类别id(+) = n_缺省卡类别id And
                           a.病人id = b.病人id And Decode(d_开始出生日期, Null, Sysdate, a.出生日期) Between Nvl(d_开始出生日期, Sysdate) And
                           Nvl(d_终止出生日期, Sysdate + 1 - 1 / 24 / 60 / 60) And
                           Decode(d_开始登记时间, Null, Sysdate, a.登记时间) Between Nvl(d_开始登记时间, Sysdate) And
                           Nvl(d_结束登记时间, (Sysdate) + 1 - 1 / 24 / 60 / 60) And
                           Decode(d_就诊开始时间, Null, Sysdate, Nvl(a.就诊时间, a.登记时间)) Between Nvl(d_就诊开始时间, Sysdate) And
                           Nvl(d_就诊结束时间, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
                           (v_姓名 Is Null Or (n_Like = 1 And a.姓名 Like v_姓名) Or (n_Like = 0 And a.姓名 = v_姓名)) And
                           Decode(Nvl(n_门诊号, 0), 0, 0, a.门诊号) = Nvl(n_门诊号, 0) And
                           Decode(Nvl(v_身份证号, '-'), '-', '-', a.身份证号) = Nvl(v_身份证号, '-') And
                           Decode(Nvl(v_费别, '-'), '-', '-', a.费别) = Nvl(v_费别, '-') And
                           Decode(Nvl(v_性别, '-'), '-', '-', a.性别) = Nvl(v_性别, '-') And
                           Decode(Nvl(v_年龄, '-'), '-', '-', a.年龄) = Nvl(v_年龄, '-') And
                           Decode(Nvl(v_区域, '-'), '-', '-', a.区域) = Nvl(v_区域, '-') And
                           Decode(Nvl(v_医保号, '-'), '-', '-', a.医保号) = Nvl(v_医保号, '-') And
                           Decode(Nvl(v_就诊卡号, '-'), '-', '-', a.就诊卡号) = Nvl(v_就诊卡号, '-') And
                           Decode(Nvl(v_Ic卡号, '-'), '-', '-', a.Ic卡号) = Nvl(v_Ic卡号, '-') And
                           Decode(Nvl(v_医疗付款方式, '-'), '-', '-', a.医疗付款方式) = Nvl(v_医疗付款方式, '-') And
                           Decode(Nvl(v_手机号, '-'), '-', '-', a.手机号) = Nvl(v_手机号, '-') And
                           (v_科室ids Is Null Or a.当前科室id Is Null Or
                           a.当前科室id In (Select 科室id
                                         From 病区科室对应
                                         Where Instr(',' || v_科室ids || ',', ',' || 病区id || ',') > 0)) And
                           (v_病区ids Is Null Or n_查询住院状态 = 2 And a.当前病区id Is Null Or
                           Instr(v_病区ids, ',' || a.当前病区id || ',') > 0) And
                           (n_查询住院状态 = 2 Or Nvl(a.在院, 0) = Nvl(n_查询住院状态, 0)) And
                           (Nvl(n_Max, 0) = 0 Or Rownum < = Nvl(n_Max, Null))) Loop
      
        Zljsonputvalue(v_List, 'pati_id', r_病人.病人id, 1, 1);
        Zljsonputvalue(v_List, 'pati_pageid', r_病人.主页id, 1);
        Zljsonputvalue(v_List, 'pati_name', r_病人.姓名);
        Zljsonputvalue(v_List, 'pati_sex', r_病人.性别);
        Zljsonputvalue(v_List, 'pati_age', r_病人.年龄);
        Zljsonputvalue(v_List, 'pati_birthdate', r_病人.出生日期);
        Zljsonputvalue(v_List, 'pati_birthplace', r_病人.出生地点);
        Zljsonputvalue(v_List, 'fee_category', r_病人.费别);
        Zljsonputvalue(v_List, 'outpatient_num', r_病人.门诊号, 0);
        Zljsonputvalue(v_List, 'inpatient_num', r_病人.住院号, 0);
        Zljsonputvalue(v_List, 'inp_times', r_病人.住院次数, 1);
      
        Zljsonputvalue(v_List, 'pati_nation', r_病人.民族);
        Zljsonputvalue(v_List, 'pati_idcard', r_病人.身份证号);
        Zljsonputvalue(v_List, 'vcard_no', r_病人.就诊卡号);
      
        Zljsonputvalue(v_List, 'pati_education', r_病人.学历);
        Zljsonputvalue(v_List, 'ocpt_name', r_病人.职业);
      
        Zljsonputvalue(v_List, 'pati_identity', r_病人.身份);
        Zljsonputvalue(v_List, 'country_name', r_病人.国籍);
        Zljsonputvalue(v_List, 'pat_home_addr', r_病人.家庭地址);
        Zljsonputvalue(v_List, 'pati_area', r_病人.区域);
        Zljsonputvalue(v_List, 'emp_name', r_病人.工作单位名称);
        Zljsonputvalue(v_List, 'emp_addr', r_病人.单位地址);
      
        Zljsonputvalue(v_List, 'is_inhspt', r_病人.在院, 1);
        Zljsonputvalue(v_List, 'pati_bed', r_病人.当前床号);
        Zljsonputvalue(v_List, 'pati_type', r_病人.病人类型);
        Zljsonputvalue(v_List, 'insurance_type', r_病人.险类, 1);
        Zljsonputvalue(v_List, 'insurance_type_name', r_病人.险类名称);
        Zljsonputvalue(v_List, 'pati_wardarea_id', r_病人.当前病区id, 1);
        Zljsonputvalue(v_List, 'pati_wardarea_name', r_病人.当前病区名称);
        Zljsonputvalue(v_List, 'pati_dept_id', r_病人.当前科室id, 1);
        Zljsonputvalue(v_List, 'pati_dept_name', r_病人.当前科室名称);
      
        Zljsonputvalue(v_List, 'adta_time', r_病人.入院时间);
        Zljsonputvalue(v_List, 'adtd_time', r_病人.出院时间);
        Zljsonputvalue(v_List, 'create_time', r_病人.登记时间);
        Zljsonputvalue(v_List, 'phone_number', r_病人.手机号);
        Zljsonputvalue(v_List, 'pat_home_phno', r_病人.家庭电话);
        If n_查询类型 = 1 Then
          Zljsonputvalue(v_List, 'stop_time', r_病人.停用时间);
        Else
          Zljsonputvalue(v_List, 'stop_time', r_病人.停用时间, 0, 2);
        End If;
        If n_查询类型 = 1 Then
          Zljsonputvalue(v_List, 'medc_card_no', r_病人.卡号);
          Zljsonputvalue(v_List, 'visit_time', r_病人.就诊时间);
          Zljsonputvalue(v_List, 'ctt_unit_id', r_病人.合约单位id);
          Zljsonputvalue(v_List, 'mdlpay_mode_name', r_病人.医疗付款方式名称);
          Zljsonputvalue(v_List, 'mdlpay_mode_code', r_病人.医疗付款方式编码);
          Zljsonputvalue(v_List, 'insurance_num', r_病人.医保号, 0);
          v_Listtmp := '"contacts":{"name":"' || Nvl(r_病人.联系人姓名, '') || '","phone":"' || Nvl(r_病人.联系人电话, '') || '"}}';
          v_List    := v_List || ',' || v_Listtmp;
        End If;
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
  
    Return;
  
  Elsif v_手机号 Is Not Null Then
    --按手机号查询
    Open c_病人信息 For
      Select a.病人id, a.姓名, a.性别, a.年龄, To_Char(a.出生日期, 'yyyy-mm-dd hh24:mi:ss') As 出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号,
             a.就诊卡号, a.费别, a.医疗付款方式 As 医疗付款方式名称, b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.家庭地址, a.工作单位,
             a.住院次数, a.当前科室id, a.当前病区id, To_Char(a.入院时间, 'yyyy-mm-dd hh24:mi:ss') As 入院时间,
             To_Char(a.出院时间, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.险类, To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间,
             a.在院, a.当前床号, a.病人类型, a.主页id, c.名称 As 当前科室名称, d.名称 As 当前病区名称, Nvl(e.名称, a.工作单位) As 工作单位名称, f.卡号,
             To_Char(a.就诊时间, 'yyyy-mm-dd hh24:mi:ss') As 就诊时间, a.手机号, a.家庭电话, e.Id As 合约单位id,
             To_Char(a.停用时间, 'yyyy-mm-dd hh24:mi:ss') As 停用时间, a.医保号, a.联系人姓名, a.联系人电话, a.单位地址, x.名称 As 险类名称
      From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 合约单位 E, 病人医疗卡信息 F, 保险类别 X
      Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.合同单位id = e.Id(+) And a.险类 = x.序号(+) And
            ((a.停用时间 Is Null And Nvl(n_是否停用, 0) = 0) Or Nvl(n_是否停用, 0) = 1) And a.病人id = f.病人id(+) And
            (n_仅合约单位 = 0 Or a.合同单位id Is Not Null) And f.卡类别id(+) = n_缺省卡类别id And
            Decode(d_开始出生日期, Null, Sysdate, a.出生日期) Between Nvl(d_开始出生日期, Sysdate) And
            Nvl(d_终止出生日期, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            Decode(d_开始登记时间, Null, Sysdate, a.登记时间) Between Nvl(d_开始登记时间, Sysdate) And
            Nvl(d_结束登记时间, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            Decode(d_就诊开始时间, Null, Sysdate, Nvl(a.就诊时间, a.登记时间)) Between Nvl(d_就诊开始时间, Sysdate) And
            Nvl(d_就诊结束时间, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            (v_姓名 Is Null Or (n_Like = 1 And a.姓名 Like v_姓名) Or (n_Like = 0 And a.姓名 = v_姓名)) And a.门诊号 = n_门诊号 And
            Decode(Nvl(v_身份证号, '-'), '-', '-', a.身份证号) = Nvl(v_身份证号, '-') And
            Decode(Nvl(v_费别, '-'), '-', '-', a.费别) = Nvl(v_费别, '-') And
            Decode(Nvl(v_性别, '-'), '-', '-', a.性别) = Nvl(v_性别, '-') And
            Decode(Nvl(v_年龄, '-'), '-', '-', a.年龄) = Nvl(v_年龄, '-') And
            Decode(Nvl(v_区域, '-'), '-', '-', a.区域) = Nvl(v_区域, '-') And
            Decode(Nvl(v_医保号, '-'), '-', '-', a.医保号) = Nvl(v_医保号, '-') And
            Decode(Nvl(v_就诊卡号, '-'), '-', '-', a.就诊卡号) = Nvl(v_就诊卡号, '-') And
            Decode(Nvl(v_Ic卡号, '-'), '-', '-', a.Ic卡号) = Nvl(v_Ic卡号, '-') And
            Decode(Nvl(v_医疗付款方式, '-'), '-', '-', a.医疗付款方式) = Nvl(v_医疗付款方式, '-') And a.手机号 = v_手机号 And
            (v_病区ids Is Null Or n_查询住院状态 = 2 And a.当前病区id Is Null Or Instr(v_病区ids, ',' || a.当前病区id || ',') > 0) And
            (n_查询住院状态 = 2 Or Nvl(a.在院, 0) = Nvl(n_查询住院状态, 0));
  Elsif n_门诊号 <> 0 Then
    --按门诊号查询
    Open c_病人信息 For
      Select a.病人id, a.姓名, a.性别, a.年龄, To_Char(a.出生日期, 'yyyy-mm-dd hh24:mi:ss') As 出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号,
             a.就诊卡号, a.费别, a.医疗付款方式 As 医疗付款方式名称, b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.家庭地址, a.工作单位,
             a.住院次数, a.当前科室id, a.当前病区id, To_Char(a.入院时间, 'yyyy-mm-dd hh24:mi:ss') As 入院时间,
             To_Char(a.出院时间, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.险类, To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间,
             a.在院, a.当前床号, a.病人类型, a.主页id, c.名称 As 当前科室名称, d.名称 As 当前病区名称, Nvl(e.名称, a.工作单位) As 工作单位名称, f.卡号,
             To_Char(a.就诊时间, 'yyyy-mm-dd hh24:mi:ss') As 就诊时间, a.手机号, a.家庭电话, e.Id As 合约单位id,
             To_Char(a.停用时间, 'yyyy-mm-dd hh24:mi:ss') As 停用时间, a.医保号, a.联系人姓名, a.联系人电话, a.单位地址, x.名称 As 险类名称
      From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 合约单位 E, 病人医疗卡信息 F, 保险类别 X
      Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.合同单位id = e.Id(+) And
            a.病人id = f.病人id(+) And (n_仅合约单位 = 0 Or a.合同单位id Is Not Null) And
            ((a.停用时间 Is Null And Nvl(n_是否停用, 0) = 0) Or Nvl(n_是否停用, 0) = 1) And f.卡类别id(+) = n_缺省卡类别id And
            a.险类 = x.序号(+) And Decode(d_开始出生日期, Null, Sysdate, a.出生日期) Between Nvl(d_开始出生日期, Sysdate) And
            Nvl(d_终止出生日期, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            Decode(d_开始登记时间, Null, Sysdate, a.登记时间) Between Nvl(d_开始登记时间, Sysdate) And
            Nvl(d_结束登记时间, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            Decode(d_就诊开始时间, Null, Sysdate, Nvl(a.就诊时间, a.登记时间)) Between Nvl(d_就诊开始时间, Sysdate) And
            Nvl(d_就诊结束时间, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            (v_姓名 Is Null Or (n_Like = 1 And a.姓名 Like v_姓名) Or (n_Like = 0 And a.姓名 = v_姓名)) And a.门诊号 = n_门诊号 And
            Decode(Nvl(v_身份证号, '-'), '-', '-', a.身份证号) = Nvl(v_身份证号, '-') And
            Decode(Nvl(v_费别, '-'), '-', '-', a.费别) = Nvl(v_费别, '-') And
            Decode(Nvl(v_性别, '-'), '-', '-', a.性别) = Nvl(v_性别, '-') And
            Decode(Nvl(v_年龄, '-'), '-', '-', a.年龄) = Nvl(v_年龄, '-') And
            Decode(Nvl(v_区域, '-'), '-', '-', a.区域) = Nvl(v_区域, '-') And
            Decode(Nvl(v_医保号, '-'), '-', '-', a.医保号) = Nvl(v_医保号, '-') And
            Decode(Nvl(v_就诊卡号, '-'), '-', '-', a.就诊卡号) = Nvl(v_就诊卡号, '-') And
            Decode(Nvl(v_Ic卡号, '-'), '-', '-', a.Ic卡号) = Nvl(v_Ic卡号, '-') And
            Decode(Nvl(v_医疗付款方式, '-'), '-', '-', a.医疗付款方式) = Nvl(v_医疗付款方式, '-') And
            Decode(Nvl(v_手机号, '-'), '-', '-', a.手机号) = Nvl(v_手机号, '-') And
            (v_病区ids Is Null Or n_查询住院状态 = 2 And a.当前病区id Is Null Or Instr(v_病区ids, ',' || a.当前病区id || ',') > 0) And
            (n_查询住院状态 = 2 Or Nvl(a.在院, 0) = Nvl(n_查询住院状态, 0));
  Elsif v_身份证号 Is Not Null Then
    If Nvl(n_场合, 0) <> 0 Then
      v_病人ids := Zl_Custom_Patiids_Get(Nvl(n_场合, 0), v_身份证号);
    End If;
    If v_病人ids Is Null Then
      Open c_病人信息 For
        Select a.病人id, a.姓名, a.性别, a.年龄, To_Char(a.出生日期, 'yyyy-mm-dd hh24:mi:ss') As 出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号,
               a.就诊卡号, a.费别, a.医疗付款方式 As 医疗付款方式名称, b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.家庭地址, a.工作单位,
               a.住院次数, a.当前科室id, a.当前病区id, To_Char(a.入院时间, 'yyyy-mm-dd hh24:mi:ss') As 入院时间,
               To_Char(a.出院时间, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.险类, To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间,
               a.在院, a.当前床号, a.病人类型, a.主页id, c.名称 As 当前科室名称, d.名称 As 当前病区名称, Nvl(e.名称, a.工作单位) As 工作单位名称, f.卡号,
               To_Char(a.就诊时间, 'yyyy-mm-dd hh24:mi:ss') As 就诊时间, a.手机号, a.家庭电话, e.Id As 合约单位id,
               To_Char(a.停用时间, 'yyyy-mm-dd hh24:mi:ss') As 停用时间, a.医保号, a.联系人姓名, a.联系人电话, a.单位地址, x.名称 As 险类名称
        From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 合约单位 E, 病人医疗卡信息 F, 保险类别 X
        Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.合同单位id = e.Id(+) And
              a.险类 = x.序号(+) And (n_仅合约单位 = 0 Or a.合同单位id Is Not Null) And
              ((a.停用时间 Is Null And Nvl(n_是否停用, 0) = 0) Or Nvl(n_是否停用, 0) = 1) And a.病人id = f.病人id(+) And
              f.卡类别id(+) = n_缺省卡类别id And Decode(d_开始出生日期, Null, Sysdate, a.出生日期) Between Nvl(d_开始出生日期, Sysdate) And
              Nvl(d_终止出生日期, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
              Decode(d_开始登记时间, Null, Sysdate, a.登记时间) Between Nvl(d_开始登记时间, Sysdate) And
              Nvl(d_结束登记时间, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
              Decode(d_就诊开始时间, Null, Sysdate, Nvl(a.就诊时间, a.登记时间)) Between Nvl(d_就诊开始时间, Sysdate) And
              Nvl(d_就诊结束时间, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
              (v_姓名 Is Null Or (n_Like = 1 And a.姓名 Like v_姓名) Or (n_Like = 0 And a.姓名 = v_姓名)) And a.身份证号 = v_身份证号 And
              Decode(Nvl(v_费别, '-'), '-', '-', a.费别) = Nvl(v_费别, '-') And
              Decode(Nvl(v_性别, '-'), '-', '-', a.性别) = Nvl(v_性别, '-') And
              Decode(Nvl(v_年龄, '-'), '-', '-', a.年龄) = Nvl(v_年龄, '-') And
              Decode(Nvl(v_区域, '-'), '-', '-', a.区域) = Nvl(v_区域, '-') And
              Decode(Nvl(v_医保号, '-'), '-', '-', a.医保号) = Nvl(v_医保号, '-') And
              Decode(Nvl(v_就诊卡号, '-'), '-', '-', a.就诊卡号) = Nvl(v_就诊卡号, '-') And
              Decode(Nvl(v_Ic卡号, '-'), '-', '-', a.Ic卡号) = Nvl(v_Ic卡号, '-') And
              Decode(Nvl(v_医疗付款方式, '-'), '-', '-', a.医疗付款方式) = Nvl(v_医疗付款方式, '-') And
              Decode(Nvl(v_手机号, '-'), '-', '-', a.手机号) = Nvl(v_手机号, '-') And
              (v_科室ids Is Null Or a.当前科室id Is Null Or
              a.当前科室id In
              (Select 科室id From 病区科室对应 Where Instr(',' || v_科室ids || ',', ',' || 病区id || ',') > 0)) And
              (v_病区ids Is Null Or n_查询住院状态 = 2 And a.当前病区id Is Null Or Instr(v_病区ids, ',' || a.当前病区id || ',') > 0) And
              (n_查询住院状态 = 2 Or Nvl(a.在院, 0) = Nvl(n_查询住院状态, 0)) And (Nvl(n_Max, 0) = 0 Or Rownum < = Nvl(n_Max, Null));
    Else
      Open c_病人信息 For
        Select a.病人id, a.姓名, a.性别, a.年龄, To_Char(a.出生日期, 'yyyy-mm-dd hh24:mi:ss') As 出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号,
               a.就诊卡号, a.费别, a.医疗付款方式 As 医疗付款方式名称, b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.家庭地址, a.工作单位,
               a.住院次数, a.当前科室id, a.当前病区id, To_Char(a.入院时间, 'yyyy-mm-dd hh24:mi:ss') As 入院时间,
               To_Char(a.出院时间, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.险类, To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间,
               a.在院, a.当前床号, a.病人类型, a.主页id, c.名称 As 当前科室名称, d.名称 As 当前病区名称, Nvl(e.名称, a.工作单位) As 工作单位名称, f.卡号,
               To_Char(a.就诊时间, 'yyyy-mm-dd hh24:mi:ss') As 就诊时间, a.手机号, a.家庭电话, e.Id As 合约单位id,
               To_Char(a.停用时间, 'yyyy-mm-dd hh24:mi:ss') As 停用时间, a.医保号, a.联系人姓名, a.联系人电话, a.单位地址, x.名称 As 险类名称
        From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 合约单位 E, 病人医疗卡信息 F, 保险类别 X
        Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.合同单位id = e.Id(+) And
              a.险类 = x.序号(+) And (n_仅合约单位 = 0 Or a.合同单位id Is Not Null) And
              ((a.停用时间 Is Null And Nvl(n_是否停用, 0) = 0) Or Nvl(n_是否停用, 0) = 1) And a.病人id = f.病人id(+) And
              f.卡类别id(+) = n_缺省卡类别id And Decode(d_开始出生日期, Null, Sysdate, a.出生日期) Between Nvl(d_开始出生日期, Sysdate) And
              Nvl(d_终止出生日期, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
              Decode(d_开始登记时间, Null, Sysdate, a.登记时间) Between Nvl(d_开始登记时间, Sysdate) And
              Nvl(d_结束登记时间, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
              Decode(d_就诊开始时间, Null, Sysdate, Nvl(a.就诊时间, a.登记时间)) Between Nvl(d_就诊开始时间, Sysdate) And
              Nvl(d_就诊结束时间, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
              (v_姓名 Is Null Or (n_Like = 1 And a.姓名 Like v_姓名) Or (n_Like = 0 And a.姓名 = v_姓名)) And
              Instr(v_病人ids, ',' || a.病人id || ',') > 0 And Decode(Nvl(v_费别, '-'), '-', '-', a.费别) = Nvl(v_费别, '-') And
              Decode(Nvl(v_性别, '-'), '-', '-', a.性别) = Nvl(v_性别, '-') And
              Decode(Nvl(v_年龄, '-'), '-', '-', a.年龄) = Nvl(v_年龄, '-') And
              Decode(Nvl(v_区域, '-'), '-', '-', a.区域) = Nvl(v_区域, '-') And
              Decode(Nvl(v_医保号, '-'), '-', '-', a.医保号) = Nvl(v_医保号, '-') And
              Decode(Nvl(v_就诊卡号, '-'), '-', '-', a.就诊卡号) = Nvl(v_就诊卡号, '-') And
              Decode(Nvl(v_Ic卡号, '-'), '-', '-', a.Ic卡号) = Nvl(v_Ic卡号, '-') And
              Decode(Nvl(v_医疗付款方式, '-'), '-', '-', a.医疗付款方式) = Nvl(v_医疗付款方式, '-') And
              Decode(Nvl(v_手机号, '-'), '-', '-', a.手机号) = Nvl(v_手机号, '-') And
              (v_科室ids Is Null Or a.当前科室id Is Null Or
              a.当前科室id In
              (Select 科室id From 病区科室对应 Where Instr(',' || v_科室ids || ',', ',' || 病区id || ',') > 0)) And
              (v_病区ids Is Null Or n_查询住院状态 = 2 And a.当前病区id Is Null Or Instr(v_病区ids, ',' || a.当前病区id || ',') > 0) And
              (n_查询住院状态 = 2 Or Nvl(a.在院, 0) = Nvl(n_查询住院状态, 0)) And (Nvl(n_Max, 0) = 0 Or Rownum < = Nvl(n_Max, Null));
    End If;
  Elsif v_就诊卡号 Is Not Null Then
  
    Open c_病人信息 For
      Select a.病人id, a.姓名, a.性别, a.年龄, To_Char(a.出生日期, 'yyyy-mm-dd hh24:mi:ss') As 出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号,
             a.就诊卡号, a.费别, a.医疗付款方式 As 医疗付款方式名称, b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.家庭地址, a.工作单位,
             a.住院次数, a.当前科室id, a.当前病区id, To_Char(a.入院时间, 'yyyy-mm-dd hh24:mi:ss') As 入院时间,
             To_Char(a.出院时间, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.险类, To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间,
             a.在院, a.当前床号, a.病人类型, a.主页id, c.名称 As 当前科室名称, d.名称 As 当前病区名称, Nvl(e.名称, a.工作单位) As 工作单位名称, f.卡号,
             To_Char(a.就诊时间, 'yyyy-mm-dd hh24:mi:ss') As 就诊时间, a.手机号, a.家庭电话, e.Id As 合约单位id,
             To_Char(a.停用时间, 'yyyy-mm-dd hh24:mi:ss') As 停用时间, a.医保号, a.联系人姓名, a.联系人电话, a.单位地址, x.名称 As 险类名称
      From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 合约单位 E, 病人医疗卡信息 F, 保险类别 X
      Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.合同单位id = e.Id(+) And a.险类 = x.序号(+) And
            (n_仅合约单位 = 0 Or a.合同单位id Is Not Null) And ((a.停用时间 Is Null And Nvl(n_是否停用, 0) = 0) Or Nvl(n_是否停用, 0) = 1) And
            a.病人id = f.病人id(+) And f.卡类别id(+) = n_缺省卡类别id And
            Decode(d_开始出生日期, Null, Sysdate, a.出生日期) Between Nvl(d_开始出生日期, Sysdate) And
            Nvl(d_终止出生日期, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            Decode(d_开始登记时间, Null, Sysdate, a.登记时间) Between Nvl(d_开始登记时间, Sysdate) And
            Nvl(d_结束登记时间, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            Decode(d_就诊开始时间, Null, Sysdate, Nvl(a.就诊时间, a.登记时间)) Between Nvl(d_就诊开始时间, Sysdate) And
            Nvl(d_就诊结束时间, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            (v_姓名 Is Null Or (n_Like = 1 And a.姓名 Like v_姓名) Or (n_Like = 0 And a.姓名 = v_姓名)) And a.就诊卡号 = v_就诊卡号 And
            Decode(Nvl(v_费别, '-'), '-', '-', a.费别) = Nvl(v_费别, '-') And
            Decode(Nvl(v_性别, '-'), '-', '-', a.性别) = Nvl(v_性别, '-') And
            Decode(Nvl(v_年龄, '-'), '-', '-', a.年龄) = Nvl(v_年龄, '-') And
            Decode(Nvl(v_区域, '-'), '-', '-', a.区域) = Nvl(v_区域, '-') And
            Decode(Nvl(v_医保号, '-'), '-', '-', a.医保号) = Nvl(v_医保号, '-') And
            Decode(Nvl(v_Ic卡号, '-'), '-', '-', a.Ic卡号) = Nvl(v_Ic卡号, '-') And
            Decode(Nvl(v_医疗付款方式, '-'), '-', '-', a.医疗付款方式) = Nvl(v_医疗付款方式, '-') And
            Decode(Nvl(v_手机号, '-'), '-', '-', a.手机号) = Nvl(v_手机号, '-') And
            (v_科室ids Is Null Or a.当前科室id Is Null Or
            a.当前科室id In (Select 科室id From 病区科室对应 Where Instr(',' || v_科室ids || ',', ',' || 病区id || ',') > 0)) And
            (v_病区ids Is Null Or n_查询住院状态 = 2 And a.当前病区id Is Null Or Instr(v_病区ids, ',' || a.当前病区id || ',') > 0) And
            (n_查询住院状态 = 2 Or Nvl(a.在院, 0) = Nvl(n_查询住院状态, 0)) And (Nvl(n_Max, 0) = 0 Or Rownum < = Nvl(n_Max, Null));
  
  Elsif v_Ic卡号 Is Not Null Then
    Open c_病人信息 For
      Select a.病人id, a.姓名, a.性别, a.年龄, To_Char(a.出生日期, 'yyyy-mm-dd hh24:mi:ss') As 出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号,
             a.就诊卡号, a.费别, a.医疗付款方式 As 医疗付款方式名称, b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.家庭地址, a.工作单位,
             a.住院次数, a.当前科室id, a.当前病区id, To_Char(a.入院时间, 'yyyy-mm-dd hh24:mi:ss') As 入院时间,
             To_Char(a.出院时间, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.险类, To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间,
             a.在院, a.当前床号, a.病人类型, a.主页id, c.名称 As 当前科室名称, d.名称 As 当前病区名称, Nvl(e.名称, a.工作单位) As 工作单位名称, f.卡号,
             To_Char(a.就诊时间, 'yyyy-mm-dd hh24:mi:ss') As 就诊时间, a.手机号, a.家庭电话, e.Id As 合约单位id,
             To_Char(a.停用时间, 'yyyy-mm-dd hh24:mi:ss') As 停用时间, a.医保号, a.联系人姓名, a.联系人电话, a.单位地址, x.名称 As 险类名称
      From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 合约单位 E, 病人医疗卡信息 F, 保险类别 X
      Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.合同单位id = e.Id(+) And a.险类 = x.序号(+) And
            ((a.停用时间 Is Null And Nvl(n_是否停用, 0) = 0) Or Nvl(n_是否停用, 0) = 1) And a.病人id = f.病人id(+) And
            f.卡类别id(+) = n_缺省卡类别id And Decode(d_开始出生日期, Null, Sysdate, a.出生日期) Between Nvl(d_开始出生日期, Sysdate) And
            Nvl(d_终止出生日期, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            Decode(d_开始登记时间, Null, Sysdate, a.登记时间) Between Nvl(d_开始登记时间, Sysdate) And
            Nvl(d_结束登记时间, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            Decode(d_就诊开始时间, Null, Sysdate, Nvl(a.就诊时间, a.登记时间)) Between Nvl(d_就诊开始时间, Sysdate) And
            Nvl(d_就诊结束时间, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            (v_姓名 Is Null Or (n_Like = 1 And a.姓名 Like v_姓名) Or (n_Like = 0 And a.姓名 = v_姓名)) And a.Ic卡号 = v_Ic卡号 And
            Decode(Nvl(v_医疗付款方式, '-'), '-', '-', a.医疗付款方式) = Nvl(v_医疗付款方式, '-') And
            Decode(Nvl(v_手机号, '-'), '-', '-', a.手机号) = Nvl(v_手机号, '-') And (n_仅合约单位 = 0 Or a.合同单位id Is Not Null) And
            Decode(Nvl(v_费别, '-'), '-', '-', a.费别) = Nvl(v_费别, '-') And
            Decode(Nvl(v_性别, '-'), '-', '-', a.性别) = Nvl(v_性别, '-') And
            Decode(Nvl(v_年龄, '-'), '-', '-', a.年龄) = Nvl(v_年龄, '-') And
            Decode(Nvl(v_区域, '-'), '-', '-', a.区域) = Nvl(v_区域, '-') And
            Decode(Nvl(v_医保号, '-'), '-', '-', a.医保号) = Nvl(v_医保号, '-') And
            (v_科室ids Is Null Or a.当前科室id Is Null Or
            a.当前科室id In (Select 科室id From 病区科室对应 Where Instr(',' || v_科室ids || ',', ',' || 病区id || ',') > 0)) And
            (v_病区ids Is Null Or n_查询住院状态 = 2 And a.当前病区id Is Null Or Instr(v_病区ids, ',' || a.当前病区id || ',') > 0) And
            (n_查询住院状态 = 2 Or Nvl(a.在院, 0) = Nvl(n_查询住院状态, 0)) And (Nvl(n_Max, 0) = 0 Or Rownum < = Nvl(n_Max, Null));
  Elsif v_医保号 Is Not Null Then
    Open c_病人信息 For
      Select a.病人id, a.姓名, a.性别, a.年龄, To_Char(a.出生日期, 'yyyy-mm-dd hh24:mi:ss') As 出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号,
             a.就诊卡号, a.费别, a.医疗付款方式 As 医疗付款方式名称, b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.家庭地址, a.工作单位,
             a.住院次数, a.当前科室id, a.当前病区id, To_Char(a.入院时间, 'yyyy-mm-dd hh24:mi:ss') As 入院时间,
             To_Char(a.出院时间, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.险类, To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间,
             a.在院, a.当前床号, a.病人类型, a.主页id, c.名称 As 当前科室名称, d.名称 As 当前病区名称, Nvl(e.名称, a.工作单位) As 工作单位名称, f.卡号,
             To_Char(a.就诊时间, 'yyyy-mm-dd hh24:mi:ss') As 就诊时间, a.手机号, a.家庭电话, e.Id As 合约单位id,
             To_Char(a.停用时间, 'yyyy-mm-dd hh24:mi:ss') As 停用时间, a.医保号, a.联系人姓名, a.联系人电话, a.单位地址, x.名称 As 险类名称
      From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 合约单位 E, 病人医疗卡信息 F, 保险类别 X
      Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.合同单位id = e.Id(+) And a.险类 = x.序号(+) And
            ((a.停用时间 Is Null And Nvl(n_是否停用, 0) = 0) Or Nvl(n_是否停用, 0) = 1) And a.病人id = f.病人id(+) And
            f.卡类别id(+) = n_缺省卡类别id And Decode(d_开始出生日期, Null, Sysdate, a.出生日期) Between Nvl(d_开始出生日期, Sysdate) And
            Nvl(d_终止出生日期, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            Decode(d_开始登记时间, Null, Sysdate, a.登记时间) Between Nvl(d_开始登记时间, Sysdate) And
            Nvl(d_结束登记时间, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            Decode(d_就诊开始时间, Null, Sysdate, Nvl(a.就诊时间, a.登记时间)) Between Nvl(d_就诊开始时间, Sysdate) And
            Nvl(d_就诊结束时间, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            (v_姓名 Is Null Or (n_Like = 1 And a.姓名 Like v_姓名) Or (n_Like = 0 And a.姓名 = v_姓名)) And
            (n_仅合约单位 = 0 Or a.合同单位id Is Not Null) And a.医保号 = v_医保号 And
            Decode(Nvl(v_费别, '-'), '-', '-', a.费别) = Nvl(v_费别, '-') And
            Decode(Nvl(v_性别, '-'), '-', '-', a.性别) = Nvl(v_性别, '-') And
            Decode(Nvl(v_年龄, '-'), '-', '-', a.年龄) = Nvl(v_年龄, '-') And
            Decode(Nvl(v_区域, '-'), '-', '-', a.区域) = Nvl(v_区域, '-') And
            Decode(Nvl(v_医疗付款方式, '-'), '-', '-', a.医疗付款方式) = Nvl(v_医疗付款方式, '-') And
            Decode(Nvl(v_手机号, '-'), '-', '-', a.手机号) = Nvl(v_手机号, '-') And
            (v_科室ids Is Null Or a.当前科室id Is Null Or
            a.当前科室id In (Select 科室id From 病区科室对应 Where Instr(',' || v_科室ids || ',', ',' || 病区id || ',') > 0)) And
            (v_病区ids Is Null Or n_查询住院状态 = 2 And a.当前病区id Is Null Or Instr(v_病区ids, ',' || a.当前病区id || ',') > 0) And
            (n_查询住院状态 = 2 Or Nvl(a.在院, 0) = Nvl(n_查询住院状态, 0)) And (Nvl(n_Max, 0) = 0 Or Rownum < = Nvl(n_Max, Null));
  Elsif v_姓名 Is Not Null Then
  
    v_病人ids := Zl_Custom_Patiids_Get(Nvl(n_场合, 0), Null, v_姓名, v_性别);
  
    If v_病人ids Is Null Then
      Open c_病人信息 For
        Select a.病人id, a.姓名, a.性别, a.年龄, To_Char(a.出生日期, 'yyyy-mm-dd hh24:mi:ss') As 出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号,
               a.就诊卡号, a.费别, a.医疗付款方式 As 医疗付款方式名称, b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.家庭地址, a.工作单位,
               a.住院次数, a.当前科室id, a.当前病区id, To_Char(a.入院时间, 'yyyy-mm-dd hh24:mi:ss') As 入院时间,
               To_Char(a.出院时间, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.险类, To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间,
               a.在院, a.当前床号, a.病人类型, a.主页id, c.名称 As 当前科室名称, d.名称 As 当前病区名称, Nvl(e.名称, a.工作单位) As 工作单位名称, f.卡号,
               To_Char(a.就诊时间, 'yyyy-mm-dd hh24:mi:ss') As 就诊时间, a.手机号, a.家庭电话, e.Id As 合约单位id,
               To_Char(a.停用时间, 'yyyy-mm-dd hh24:mi:ss') As 停用时间, a.医保号, a.联系人姓名, a.联系人电话, a.单位地址, x.名称 As 险类名称
        From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 合约单位 E, 病人医疗卡信息 F, 保险类别 X
        Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.合同单位id = e.Id(+) And
              a.险类 = x.序号(+) And ((a.停用时间 Is Null And Nvl(n_是否停用, 0) = 0) Or Nvl(n_是否停用, 0) = 1) And a.病人id = f.病人id(+) And
              f.卡类别id(+) = n_缺省卡类别id And Decode(d_开始出生日期, Null, Sysdate, a.出生日期) Between Nvl(d_开始出生日期, Sysdate) And
              Nvl(d_终止出生日期, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
              Decode(d_开始登记时间, Null, Sysdate, a.登记时间) Between Nvl(d_开始登记时间, Sysdate) And
              Nvl(d_结束登记时间, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
              Decode(d_就诊开始时间, Null, Sysdate, Nvl(a.就诊时间, a.登记时间)) Between Nvl(d_就诊开始时间, Sysdate) And
              Nvl(d_就诊结束时间, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
              ((n_Like = 1 And a.姓名 Like v_姓名) Or (n_Like = 0 And a.姓名 = v_姓名)) And
              (n_仅合约单位 = 0 Or a.合同单位id Is Not Null) And Decode(Nvl(v_费别, '-'), '-', '-', a.费别) = Nvl(v_费别, '-') And
              Decode(Nvl(v_性别, '-'), '-', '-', a.性别) = Nvl(v_性别, '-') And
              Decode(Nvl(v_年龄, '-'), '-', '-', a.年龄) = Nvl(v_年龄, '-') And
              Decode(Nvl(v_区域, '-'), '-', '-', a.区域) = Nvl(v_区域, '-') And
              Decode(Nvl(v_医疗付款方式, '-'), '-', '-', a.医疗付款方式) = Nvl(v_医疗付款方式, '-') And
              Decode(Nvl(v_手机号, '-'), '-', '-', a.手机号) = Nvl(v_手机号, '-') And
              (v_科室ids Is Null Or a.当前科室id Is Null Or
              a.当前科室id In
              (Select 科室id From 病区科室对应 Where Instr(',' || v_科室ids || ',', ',' || 病区id || ',') > 0)) And
              (v_病区ids Is Null Or n_查询住院状态 = 2 And a.当前病区id Is Null Or Instr(v_病区ids, ',' || a.当前病区id || ',') > 0) And
              (n_查询住院状态 = 2 Or Nvl(a.在院, 0) = Nvl(n_查询住院状态, 0)) And (Nvl(n_Max, 0) = 0 Or Rownum < = Nvl(n_Max, Null));
    Else
      Open c_病人信息 For
        Select a.病人id, a.姓名, a.性别, a.年龄, To_Char(a.出生日期, 'yyyy-mm-dd hh24:mi:ss') As 出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号,
               a.就诊卡号, a.费别, a.医疗付款方式 As 医疗付款方式名称, b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.家庭地址, a.工作单位,
               a.住院次数, a.当前科室id, a.当前病区id, To_Char(a.入院时间, 'yyyy-mm-dd hh24:mi:ss') As 入院时间,
               To_Char(a.出院时间, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.险类, To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间,
               a.在院, a.当前床号, a.病人类型, a.主页id, c.名称 As 当前科室名称, d.名称 As 当前病区名称, Nvl(e.名称, a.工作单位) As 工作单位名称, f.卡号,
               To_Char(a.就诊时间, 'yyyy-mm-dd hh24:mi:ss') As 就诊时间, a.手机号, a.家庭电话, e.Id As 合约单位id,
               To_Char(a.停用时间, 'yyyy-mm-dd hh24:mi:ss') As 停用时间, a.医保号, a.联系人姓名, a.联系人电话, a.单位地址, x.名称 As 险类名称
        From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 合约单位 E, 病人医疗卡信息 F, 保险类别 X
        Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.合同单位id = e.Id(+) And
              a.险类 = x.序号(+) And (n_仅合约单位 = 0 Or a.合同单位id Is Not Null) And
              ((a.停用时间 Is Null And Nvl(n_是否停用, 0) = 0) Or Nvl(n_是否停用, 0) = 1) And a.病人id = f.病人id(+) And
              f.卡类别id(+) = n_缺省卡类别id And Decode(d_开始出生日期, Null, Sysdate, a.出生日期) Between Nvl(d_开始出生日期, Sysdate) And
              Nvl(d_终止出生日期, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
              Decode(d_开始登记时间, Null, Sysdate, a.登记时间) Between Nvl(d_开始登记时间, Sysdate) And
              Nvl(d_结束登记时间, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
              Decode(d_就诊开始时间, Null, Sysdate, Nvl(a.就诊时间, a.登记时间)) Between Nvl(d_就诊开始时间, Sysdate) And
              Nvl(d_就诊结束时间, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
              (v_姓名 Is Null Or (n_Like = 1 And a.姓名 Like v_姓名) Or (n_Like = 0 And a.姓名 = v_姓名)) And
              Instr(v_病人ids, ',' || a.病人id || ',') > 0 And Decode(Nvl(v_费别, '-'), '-', '-', a.费别) = Nvl(v_费别, '-') And
              Decode(Nvl(v_性别, '-'), '-', '-', a.性别) = Nvl(v_性别, '-') And
              Decode(Nvl(v_年龄, '-'), '-', '-', a.年龄) = Nvl(v_年龄, '-') And
              Decode(Nvl(v_区域, '-'), '-', '-', a.区域) = Nvl(v_区域, '-') And
              Decode(Nvl(v_医保号, '-'), '-', '-', a.医保号) = Nvl(v_医保号, '-') And
              Decode(Nvl(v_就诊卡号, '-'), '-', '-', a.就诊卡号) = Nvl(v_就诊卡号, '-') And
              Decode(Nvl(v_Ic卡号, '-'), '-', '-', a.Ic卡号) = Nvl(v_Ic卡号, '-') And
              Decode(Nvl(v_医疗付款方式, '-'), '-', '-', a.医疗付款方式) = Nvl(v_医疗付款方式, '-') And
              Decode(Nvl(v_手机号, '-'), '-', '-', a.手机号) = Nvl(v_手机号, '-') And
              (v_科室ids Is Null Or a.当前科室id Is Null Or
              a.当前科室id In
              (Select 科室id From 病区科室对应 Where Instr(',' || v_科室ids || ',', ',' || 病区id || ',') > 0)) And
              (v_病区ids Is Null Or n_查询住院状态 = 2 And a.当前病区id Is Null Or Instr(v_病区ids, ',' || a.当前病区id || ',') > 0) And
              (n_查询住院状态 = 2 Or Nvl(a.在院, 0) = Nvl(n_查询住院状态, 0)) And (Nvl(n_Max, 0) = 0 Or Rownum < = Nvl(n_Max, Null));
    End If;
  
  Elsif d_开始出生日期 Is Not Null Then
    Open c_病人信息 For
      Select a.病人id, a.姓名, a.性别, a.年龄, To_Char(a.出生日期, 'yyyy-mm-dd hh24:mi:ss') As 出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号,
             a.就诊卡号, a.费别, a.医疗付款方式 As 医疗付款方式名称, b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.家庭地址, a.工作单位,
             a.住院次数, a.当前科室id, a.当前病区id, To_Char(a.入院时间, 'yyyy-mm-dd hh24:mi:ss') As 入院时间,
             To_Char(a.出院时间, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.险类, To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间,
             a.在院, a.当前床号, a.病人类型, a.主页id, c.名称 As 当前科室名称, d.名称 As 当前病区名称, Nvl(e.名称, a.工作单位) As 工作单位名称, f.卡号,
             To_Char(a.就诊时间, 'yyyy-mm-dd hh24:mi:ss') As 就诊时间, a.手机号, a.家庭电话, e.Id As 合约单位id,
             To_Char(a.停用时间, 'yyyy-mm-dd hh24:mi:ss') As 停用时间, a.医保号, a.联系人姓名, a.联系人电话, a.单位地址, x.名称 As 险类名称
      From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 合约单位 E, 病人医疗卡信息 F, 保险类别 X
      Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.合同单位id = e.Id(+) And a.险类 = x.序号(+) And
            ((a.停用时间 Is Null And Nvl(n_是否停用, 0) = 0) Or Nvl(n_是否停用, 0) = 1) And a.病人id = f.病人id(+) And
            f.卡类别id(+) = n_缺省卡类别id And a.出生日期 Between d_开始出生日期 And Nvl(d_终止出生日期, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            Decode(d_开始登记时间, Null, Sysdate, a.登记时间) Between Nvl(d_开始登记时间, Sysdate) And
            Nvl(d_结束登记时间, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            Decode(d_就诊开始时间, Null, Sysdate, Nvl(a.就诊时间, a.登记时间)) Between Nvl(d_就诊开始时间, Sysdate) And
            Nvl(d_就诊结束时间, (Sysdate + 1 - 1 / 24 / 60 / 60)) And Decode(Nvl(v_费别, '-'), '-', '-', a.费别) = Nvl(v_费别, '-') And
            Decode(Nvl(v_性别, '-'), '-', '-', a.性别) = Nvl(v_性别, '-') And
            Decode(Nvl(v_年龄, '-'), '-', '-', a.年龄) = Nvl(v_年龄, '-') And
            Decode(Nvl(v_区域, '-'), '-', '-', a.区域) = Nvl(v_区域, '-') And
            Decode(Nvl(v_医疗付款方式, '-'), '-', '-', a.医疗付款方式) = Nvl(v_医疗付款方式, '-') And
            Decode(Nvl(v_手机号, '-'), '-', '-', a.手机号) = Nvl(v_手机号, '-') And
            (v_科室ids Is Null Or a.当前科室id Is Null Or
            a.当前科室id In (Select 科室id From 病区科室对应 Where Instr(',' || v_科室ids || ',', ',' || 病区id || ',') > 0)) And
            (v_病区ids Is Null Or n_查询住院状态 = 2 And a.当前病区id Is Null Or Instr(v_病区ids, ',' || a.当前病区id || ',') > 0) And
            (n_仅合约单位 = 0 Or a.合同单位id Is Not Null) And (n_查询住院状态 = 2 Or Nvl(a.在院, 0) = Nvl(n_查询住院状态, 0)) And
            (Nvl(n_Max, 0) = 0 Or Rownum < = Nvl(n_Max, Null));
  Elsif d_开始登记时间 Is Not Null Then
    Open c_病人信息 For
      Select a.病人id, a.姓名, a.性别, a.年龄, To_Char(a.出生日期, 'yyyy-mm-dd hh24:mi:ss') As 出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号,
             a.就诊卡号, a.费别, a.医疗付款方式 As 医疗付款方式名称, b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.家庭地址, a.工作单位,
             a.住院次数, a.当前科室id, a.当前病区id, To_Char(a.入院时间, 'yyyy-mm-dd hh24:mi:ss') As 入院时间,
             To_Char(a.出院时间, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.险类, To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间,
             a.在院, a.当前床号, a.病人类型, a.主页id, c.名称 As 当前科室名称, d.名称 As 当前病区名称, Nvl(e.名称, a.工作单位) As 工作单位名称, f.卡号,
             To_Char(a.就诊时间, 'yyyy-mm-dd hh24:mi:ss') As 就诊时间, a.手机号, a.家庭电话, e.Id As 合约单位id,
             To_Char(a.停用时间, 'yyyy-mm-dd hh24:mi:ss') As 停用时间, a.医保号, a.联系人姓名, a.联系人电话, a.单位地址, x.名称 As 险类名称
      From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 合约单位 E, 病人医疗卡信息 F, 保险类别 X
      Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.合同单位id = e.Id(+) And a.险类 = x.序号(+) And
            ((a.停用时间 Is Null And Nvl(n_是否停用, 0) = 0) Or Nvl(n_是否停用, 0) = 1) And a.病人id = f.病人id(+) And
            f.卡类别id(+) = n_缺省卡类别id And a.登记时间 Between d_开始登记时间 And Nvl(d_结束登记时间, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            Decode(d_就诊开始时间, Null, Sysdate, Nvl(a.就诊时间, a.登记时间)) Between Nvl(d_就诊开始时间, Sysdate) And
            Nvl(d_就诊结束时间, (Sysdate + 1 - 1 / 24 / 60 / 60)) And Decode(Nvl(v_费别, '-'), '-', '-', a.费别) = Nvl(v_费别, '-') And
            Decode(Nvl(v_性别, '-'), '-', '-', a.性别) = Nvl(v_性别, '-') And
            Decode(Nvl(v_年龄, '-'), '-', '-', a.年龄) = Nvl(v_年龄, '-') And
            Decode(Nvl(v_区域, '-'), '-', '-', a.区域) = Nvl(v_区域, '-') And
            Decode(Nvl(v_医疗付款方式, '-'), '-', '-', a.医疗付款方式) = Nvl(v_医疗付款方式, '-') And
            Decode(Nvl(v_手机号, '-'), '-', '-', a.手机号) = Nvl(v_手机号, '-') And
            (v_科室ids Is Null Or a.当前科室id Is Null Or
            a.当前科室id In (Select 科室id From 病区科室对应 Where Instr(',' || v_科室ids || ',', ',' || 病区id || ',') > 0)) And
            (v_病区ids Is Null Or n_查询住院状态 = 2 And a.当前病区id Is Null Or Instr(v_病区ids, ',' || a.当前病区id || ',') > 0) And
            (n_仅合约单位 = 0 Or a.合同单位id Is Not Null) And (n_查询住院状态 = 2 Or Nvl(a.在院, 0) = Nvl(n_查询住院状态, 0)) And
            (Nvl(n_Max, 0) = 0 Or Rownum < = Nvl(n_Max, Null));
  Elsif d_就诊开始时间 Is Not Null Then
    Open c_病人信息 For
      Select a.病人id, a.姓名, a.性别, a.年龄, To_Char(a.出生日期, 'yyyy-mm-dd hh24:mi:ss') As 出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号,
             a.就诊卡号, a.费别, a.医疗付款方式 As 医疗付款方式名称, b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.家庭地址, a.工作单位,
             a.住院次数, a.当前科室id, a.当前病区id, To_Char(a.入院时间, 'yyyy-mm-dd hh24:mi:ss') As 入院时间,
             To_Char(a.出院时间, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.险类, To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间,
             a.在院, a.当前床号, a.病人类型, a.主页id, c.名称 As 当前科室名称, d.名称 As 当前病区名称, Nvl(e.名称, a.工作单位) As 工作单位名称, f.卡号,
             To_Char(a.就诊时间, 'yyyy-mm-dd hh24:mi:ss') As 就诊时间, a.手机号, a.家庭电话, e.Id As 合约单位id,
             To_Char(a.停用时间, 'yyyy-mm-dd hh24:mi:ss') As 停用时间, a.医保号, a.联系人姓名, a.联系人电话, a.单位地址, x.名称 As 险类名称
      From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 合约单位 E, 病人医疗卡信息 F, 保险类别 X
      Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.合同单位id = e.Id(+) And a.险类 = x.序号(+) And
            ((a.停用时间 Is Null And Nvl(n_是否停用, 0) = 0) Or Nvl(n_是否停用, 0) = 1) And a.病人id = f.病人id(+) And
            f.卡类别id(+) = n_缺省卡类别id And Nvl(a.就诊时间, a.登记时间) Between d_就诊开始时间 And
            Nvl(d_就诊结束时间, (Sysdate + 1 - 1 / 24 / 60 / 60)) And Decode(Nvl(v_费别, '-'), '-', '-', a.费别) = Nvl(v_费别, '-') And
            Decode(Nvl(v_性别, '-'), '-', '-', a.性别) = Nvl(v_性别, '-') And
            Decode(Nvl(v_年龄, '-'), '-', '-', a.年龄) = Nvl(v_年龄, '-') And
            Decode(Nvl(v_区域, '-'), '-', '-', a.区域) = Nvl(v_区域, '-') And
            Decode(Nvl(v_医疗付款方式, '-'), '-', '-', a.医疗付款方式) = Nvl(v_医疗付款方式, '-') And
            Decode(Nvl(v_手机号, '-'), '-', '-', a.手机号) = Nvl(v_手机号, '-') And
            (v_科室ids Is Null Or a.当前科室id Is Null Or
            a.当前科室id In (Select 科室id From 病区科室对应 Where Instr(',' || v_科室ids || ',', ',' || 病区id || ',') > 0)) And
            (v_病区ids Is Null Or n_查询住院状态 = 2 And a.当前病区id Is Null Or Instr(v_病区ids, ',' || a.当前病区id || ',') > 0) And
            (n_仅合约单位 = 0 Or a.合同单位id Is Not Null) And (n_查询住院状态 = 2 Or Nvl(a.在院, 0) = Nvl(n_查询住院状态, 0)) And
            (Nvl(n_Max, 0) = 0 Or Rownum < = Nvl(n_Max, Null));
  
  Elsif Nvl(n_仅合约单位, 0) = 1 Then
    Open c_病人信息 For
      Select a.病人id, a.姓名, a.性别, a.年龄, To_Char(a.出生日期, 'yyyy-mm-dd hh24:mi:ss') As 出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号,
             a.就诊卡号, a.费别, a.医疗付款方式 As 医疗付款方式名称, b.编码 As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.家庭地址, a.工作单位,
             a.住院次数, a.当前科室id, a.当前病区id, To_Char(a.入院时间, 'yyyy-mm-dd hh24:mi:ss') As 入院时间,
             To_Char(a.出院时间, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.险类, To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间,
             a.在院, a.当前床号, a.病人类型, a.主页id, c.名称 As 当前科室名称, d.名称 As 当前病区名称, Nvl(e.名称, a.工作单位) As 工作单位名称, f.卡号,
             To_Char(a.就诊时间, 'yyyy-mm-dd hh24:mi:ss') As 就诊时间, a.手机号, a.家庭电话, e.Id As 合约单位id,
             To_Char(a.停用时间, 'yyyy-mm-dd hh24:mi:ss') As 停用时间, a.医保号, a.联系人姓名, a.联系人电话, a.单位地址, x.名称 As 险类名称
      From 病人信息 A, 医疗付款方式 B, 部门表 C, 部门表 D, 合约单位 E, 病人医疗卡信息 F, 保险类别 X
      Where a.医疗付款方式 = b.名称(+) And a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.合同单位id = e.Id(+) And a.险类 = x.序号(+) And
            ((a.停用时间 Is Null And Nvl(n_是否停用, 0) = 0) Or Nvl(n_是否停用, 0) = 1) And a.病人id = f.病人id(+) And
            f.卡类别id(+) = n_缺省卡类别id And Decode(Nvl(v_费别, '-'), '-', '-', a.费别) = Nvl(v_费别, '-') And
            Decode(Nvl(v_性别, '-'), '-', '-', a.性别) = Nvl(v_性别, '-') And
            Decode(Nvl(v_年龄, '-'), '-', '-', a.年龄) = Nvl(v_年龄, '-') And
            Decode(Nvl(v_区域, '-'), '-', '-', a.区域) = Nvl(v_区域, '-') And
            Decode(Nvl(v_医疗付款方式, '-'), '-', '-', a.医疗付款方式) = Nvl(v_医疗付款方式, '-') And
            Decode(Nvl(v_手机号, '-'), '-', '-', a.手机号) = Nvl(v_手机号, '-') And
            (v_科室ids Is Null Or a.当前科室id Is Null Or
            a.当前科室id In (Select 科室id From 病区科室对应 Where Instr(',' || v_科室ids || ',', ',' || 病区id || ',') > 0)) And
            (v_病区ids Is Null Or n_查询住院状态 = 2 And a.当前病区id Is Null Or Instr(v_病区ids, ',' || a.当前病区id || ',') > 0) And
            (Nvl(n_合同单位id, 0) = 0 And a.合同单位id Is Not Null Or a.合同单位id = n_合同单位id) And
            (n_查询住院状态 = 2 Or Nvl(a.在院, 0) = Nvl(n_查询住院状态, 0)) And (Nvl(n_Max, 0) = 0 Or Rownum < = Nvl(n_Max, Null));
  Elsif Not j_Json_Similar Is Null Then
    Open c_病人信息 For
      Select a.病人id, a.姓名, a.性别, a.年龄, To_Char(a.出生日期, 'yyyy-mm-dd hh24:mi:ss') As 出生日期, a.出生地点, a.身份证号, a.门诊号, a.住院号,
             a.就诊卡号, a.费别, a.医疗付款方式 As 医疗付款方式名称, Null As 医疗付款方式编码, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.家庭地址, a.工作单位,
             a.住院次数, a.当前科室id, a.当前病区id, To_Char(a.入院时间, 'yyyy-mm-dd hh24:mi:ss') As 入院时间,
             To_Char(a.出院时间, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.险类, To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间,
             a.在院, a.当前床号, a.病人类型, a.主页id, Null As 当前科室名称, Null As 当前病区名称, a.工作单位 As 工作单位名称, Null 卡号,
             To_Char(a.就诊时间, 'yyyy-mm-dd hh24:mi:ss') As 就诊时间, a.手机号, a.家庭电话, Null As 合约单位id,
             To_Char(a.停用时间, 'yyyy-mm-dd hh24:mi:ss') As 停用时间, a.医保号, a.联系人姓名, a.联系人电话, a.单位地址, x.名称 As 险类名称
      From 病人信息 A, 保险类别 X
      Where a.停用时间 Is Null And a.险类 = x.序号(+) And
            ((a.姓名 = v_姓名 And a.性别 = v_性别 And a.出生日期 = d_出生日期 And a.国籍 = v_国籍 And a.民族 = v_民族) Or a.身份证号 = v_身份证号)
      Order By Nvl(Nvl(a.就诊时间, a.入院时间), a.登记时间) Desc;
  Else
    v_Err_Msg := '未传入有效的查询条件，不能获取病人信息!';
    Json_Out  := Get_Err_Message(v_Err_Msg);
    Return;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功","pati_list":[';
  Loop
    Fetch c_病人信息
      Into r_病人;
    Exit When c_病人信息%NotFound;
    Zljsonputvalue(v_List, 'pati_id', r_病人.病人id, 1, 1);
    Zljsonputvalue(v_List, 'pati_pageid', r_病人.主页id, 1);
  
    Zljsonputvalue(v_List, 'pati_name', r_病人.姓名);
    Zljsonputvalue(v_List, 'pati_sex', r_病人.性别);
  
    Zljsonputvalue(v_List, 'pati_age', r_病人.年龄);
    Zljsonputvalue(v_List, 'pati_birthdate', r_病人.出生日期);
    Zljsonputvalue(v_List, 'pati_birthplace', r_病人.出生地点);
    Zljsonputvalue(v_List, 'fee_category', r_病人.费别);
  
    Zljsonputvalue(v_List, 'outpatient_num', r_病人.门诊号, 0);
    Zljsonputvalue(v_List, 'inpatient_num', r_病人.住院号, 0);
    Zljsonputvalue(v_List, 'inp_times', r_病人.住院次数, 1);
  
    Zljsonputvalue(v_List, 'pati_nation', r_病人.民族);
    Zljsonputvalue(v_List, 'pati_idcard', r_病人.身份证号);
    Zljsonputvalue(v_List, 'vcard_no', r_病人.就诊卡号);
  
    Zljsonputvalue(v_List, 'pati_education', r_病人.学历);
    Zljsonputvalue(v_List, 'ocpt_name', r_病人.职业);
  
    Zljsonputvalue(v_List, 'pati_identity', r_病人.身份);
    Zljsonputvalue(v_List, 'country_name', r_病人.国籍);
    Zljsonputvalue(v_List, 'pat_home_addr', r_病人.家庭地址);
    Zljsonputvalue(v_List, 'pati_area', r_病人.区域);
    Zljsonputvalue(v_List, 'emp_name', r_病人.工作单位名称);
    Zljsonputvalue(v_List, 'emp_addr', r_病人.单位地址);
  
    Zljsonputvalue(v_List, 'is_inhspt', r_病人.在院, 1);
    Zljsonputvalue(v_List, 'pati_bed', r_病人.当前床号);
    Zljsonputvalue(v_List, 'pati_type', r_病人.病人类型);
    Zljsonputvalue(v_List, 'insurance_type', r_病人.险类, 1);
    Zljsonputvalue(v_List, 'insurance_type_name', r_病人.险类名称);
    Zljsonputvalue(v_List, 'pati_wardarea_id', r_病人.当前病区id, 1);
    Zljsonputvalue(v_List, 'pati_wardarea_name', r_病人.当前病区名称);
    Zljsonputvalue(v_List, 'pati_dept_id', r_病人.当前科室id, 1);
    Zljsonputvalue(v_List, 'pati_dept_name', r_病人.当前科室名称);
  
    Zljsonputvalue(v_List, 'adta_time', r_病人.入院时间);
    Zljsonputvalue(v_List, 'adtd_time', r_病人.出院时间);
    Zljsonputvalue(v_List, 'create_time', r_病人.登记时间);
    Zljsonputvalue(v_List, 'phone_number', r_病人.手机号);
    Zljsonputvalue(v_List, 'pat_home_phno', r_病人.家庭电话);
    If n_查询类型 = 1 Then
      Zljsonputvalue(v_List, 'stop_time', r_病人.停用时间);
    Else
      Zljsonputvalue(v_List, 'stop_time', r_病人.停用时间, 0, 2);
    End If;
    If n_查询类型 = 1 Then
      Zljsonputvalue(v_List, 'medc_card_no', r_病人.卡号);
      Zljsonputvalue(v_List, 'visit_time', r_病人.就诊时间);
      Zljsonputvalue(v_List, 'ctt_unit_id', r_病人.合约单位id, 1);
      Zljsonputvalue(v_List, 'mdlpay_mode_name', r_病人.医疗付款方式名称);
      Zljsonputvalue(v_List, 'mdlpay_mode_code', r_病人.医疗付款方式编码);
      Zljsonputvalue(v_List, 'insurance_num', r_病人.医保号, 0);
      v_Listtmp := '"contacts":{"name":"' || Nvl(r_病人.联系人姓名, '') || '","phone":"' || Nvl(r_病人.联系人电话, '') || '"}}';
      v_List    := v_List || ',' || v_Listtmp;
    End If;
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
End Zl_Patisvr_Getpatiinfsbyrange;
/
Create Or Replace Procedure Zl_Patisvr_Getpatimergeinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:根据病人id获取病人的所有合并信息
  --入参：Json_In:格式
  --  input
  --    pati_id             N 1 病人id
  --出参: Json_Out,格式如下
  --  output
  --    code                N   1   应答码：0-失败；1-成功
  --    message             C   1   应答消息：失败时返回具体的错误信息
  --    merge_list[]        C       合并信息列表
  --      info_old          C   1   原信息
  --      merge_reason      C   1   合并原因
  --      operator_name     C   1   操作员
  --      merge_time        C   1   合并时间:yyyy-mm-dd hh24:mi:ss
  ---------------------------------------------------------------------------

  j_Json   Pljson;
  j_Jsonin Pljson;
  n_病人id 病人不良记录.病人id%Type;
  v_Jtmp   Varchar2(32767);
Begin
  --解析入参
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');

  If Nvl(n_病人id, 0) = 0 Then
    Json_Out := Zljsonout('失败，未传入病人id');
    Return;
  End If;
  For r_合并 In (Select 原信息, 合并原因, 操作员姓名, 合并时间
               From 病人合并记录
               Where 病人id = n_病人id
               Order By 合并时间 Desc) Loop
  
    v_Jtmp := v_Jtmp || ',{"info_old":"' || Zljsonstr(r_合并.原信息) || '"';
    v_Jtmp := v_Jtmp || ',"merge_reason":"' || Zljsonstr(r_合并.合并原因) || '"';
    v_Jtmp := v_Jtmp || ',"operator_name":"' || Zljsonstr(r_合并.操作员姓名) || '"';
    v_Jtmp := v_Jtmp || ',"merge_time":"' || To_Char(r_合并.合并时间, 'yyyy-mm-dd hh24:mi:ss') || '"';
  
    v_Jtmp := v_Jtmp || '}';
  
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"成功","merge_list":[' || Substr(v_Jtmp, 2) || ']}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getpatimergeinfo;
/
Create Or Replace Procedure Zl_Patisvr_Getpatimmuneinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:获取病人信息的免疫信息
  --入参：Json_In:格式
  --  input
  --    pati_id           N   1 病人id
  --出参: Json_Out,格式如下
  --  output
  --    code              N   1   应答码：0-失败；1-成功
  --    message           C   1   应答消息：失败时返回具体的错误信息
  --    immune_list[]     C       病人免疫列表
  --      vaccinate_time    C   1   接种时间:yyyy-mm-dd hh24:mi:ss
  --      vaccinate_name    C   1   接种名称
  ---------------------------------------------------------------------------

  j_Json   Pljson;
  j_Jsonin Pljson;
  n_病人id 病人免疫记录.病人id%Type;
  v_Jtmp   Varchar2(32767);
  c_Jtmp   Clob;

Begin
  --解析入参
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');

  If Nvl(n_病人id, 0) = 0 Then
    Json_Out := Zljsonout('未传入病人id，请检查!');
    Return;
  End If;
  For r_免疫记录 In (Select Distinct 接种时间, 接种名称 From 病人免疫记录 Where 病人id = n_病人id) Loop
    v_Jtmp := v_Jtmp || ',{"vaccinate_time":"' || To_Char(r_免疫记录.接种时间, 'yyyy-mm-dd hh24:mi:ss') || '"';
    v_Jtmp := v_Jtmp || ',"vaccinate_name":"' || Zljsonstr(r_免疫记录.接种名称) || '"';
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
    Json_Out := '{"output":{"code":1,"message":"成功","immune_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","immune_list":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getpatimmuneinfo;
/
Create Or Replace Procedure Zl_Patisvr_Getpatiphoto
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:获取病人相片
  --入参：Json_In:格式
  --    input
  --      pati_id           N 1 病人ID
  --出参      json
  --output
  --     code               N 1 应答码：0-失败；1-成功
  --     message            C 1 应答消息： 失败时返回具体的错误信息
  --     pati_photo         C 1 编码:base64
  ---------------------------------------------------------------------------
  j_Json     Pljson;
  j_Jsonin   Pljson;
  n_病人id   病人照片.病人id%Type;
  b_病人照片 病人照片.照片%Type;
  v_Clob     Clob;
Begin
  --解析入参
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');

  If Nvl(n_病人id, 0) = 0 Then
    Json_Out := '{"output":{"code":0,"message":"未传入病人id，不能保存病人照片!"}}';
    Return;
  End If;

  Begin
    Select 照片 Into b_病人照片 From 病人照片 Where 病人id = n_病人id;
    v_Clob := Zltools.Zlbase64.Encode(b_病人照片);
    v_Clob := Replace(v_Clob, Chr(13), '');
    v_Clob := Replace(v_Clob, Chr(10), '');
  Exception
    When Others Then
      v_Clob := Null;
  End;
  Json_Out := '{"output":{"code":1,"message":"成功","pati_photo":"' || v_Clob || '"}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getpatiphoto;
/
Create Or Replace Procedure Zl_Patisvr_Getpatirelate
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --功能：获取指定人与之关的联的病人ID串，串中不包含当前传入的病人ID
  --入参：Json_In:格式
  --input
  --  query_type        N    1 调用类型 ：1-通过身份证查询其他病人ID,2-通过病人身份关联表查询其他的病人id
  --  pati_id           N    1 病人id
  --  pati_idcard       C    1 身份证号

  --出参      json
  --output
  --    code                N   1 应答吗：0-失败；1-成功
  --    message             C   1 应答消息：失败时返回具体的错误信息
  --    pati_ids            C   1 病人id,逗号拼串
  -------------------------------------------

  j_Json   Pljson;
  j_Jsonin Pljson;
  v_List   Varchar2(32767);

  n_Type     Number(18);
  n_病人id   Number(18);
  v_身份证号 Varchar2(50);
Begin
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_Type     := j_Json.Get_Number('query_type');
  n_病人id   := j_Json.Get_Number('pati_id');
  v_身份证号 := j_Json.Get_String('pati_idcard');

  If n_Type = 1 Then
    For c_病人 In (Select a.病人id From 病人信息 A Where a.病人id <> n_病人id And a.身份证号 = v_身份证号) Loop
      v_List := v_List || ',' || c_病人.病人id;
    End Loop;
  Elsif n_Type = 2 Then
    For c_病人 In (Select b.病人id
                 From 病人身份关联 A, 病人身份关联 B, 病人信息 C
                 Where a.关联id = b.关联id And b.病人id = c.病人id And a.病人id = n_病人id And b.病人id + 0 <> n_病人id And
                       (Nvl(c.身份证号, '-') <> v_身份证号 Or v_身份证号 Is Null)) Loop
      v_List := v_List || ',' || c_病人.病人id;
    End Loop;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功","pati_ids":"' || Substr(v_List, 2) || '"}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getpatirelate;
/
Create Or Replace Procedure Zl_Patisvr_Getvisitpatis
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:获取就诊病人信息
  --入参：Json_In:格式
  --    input
  --      pati_ids          C   病人IDs:多个用逗号
  --      vcard_no          C   就诊卡号

  --出参      json
  --output
  -- code                   N 1 应答码：0-失败；1-成功
  -- message                C 1 应答消息： 失败时返回具体的错误信息
  -- pati_list[]                病人信息列表
  --   pati_id              N 1 病人id
  --   pati_name            C 1 姓名
  --   pati_sex             C 1 性别
  --   pati_age             C 1 年龄
  --   pati_birthdate       C 1 出生日期：yyyy-mm-dd hh24:mi:ss
  --   fee_category         C 1 费别
  --   outpatient_num       C 1 门诊号
  --   pati_nation          C 1 民族
  --   pati_idcard          C 1 身份证号
  --   vcard_no             C 1 就诊卡号
  --   pati_education       C 1 学历
  --   ocpt_name            C 1 职业
  --   pati_identity        C 1 身份
  --   country_name         C 1 国籍
  --   pat_home_addr        C 1 家庭地址
  --   pati_area            C 1 区域
  --   emp_name             C 1 工作单位名称
  --   pati_type            C 1 病人类型(普通，医保，留观)
  --   insurance_type       C 1 险类
  --   create_time          C 1 登记时间
  --   pati_dept_id         N 1 当前科室id
  --   pati_dept_name       C 1 当前科室名称
  --   iccard_no            C 1 Ic卡号
  ---------------------------------------------------------------------------
  v_Err_Msg Varchar2(500);
  j_Json    Pljson;
  j_Jsonin  Pljson;

  v_List Varchar2(32767);

  c_病人ids Clob;

  v_就诊卡号 Varchar2(200);

  l_病人id t_Strlist := t_Strlist();

  Cursor c_病人基本信息 Is
    Select a.病人id, a.姓名, a.性别, a.年龄, a.出生日期, a.身份证号, a.门诊号, a.住院号, a.就诊卡号, a.费别, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历,
           a.家庭地址, a.工作单位, a.住院次数, a.当前科室id, a.当前病区id, To_Char(a.入院时间, 'yyyy-mm-dd hh24:mi:ss') As 入院时间,
           To_Char(a.出院时间, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.险类, To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间,
           a.在院, a.当前床号, a.病人类型, a.主页id, c.名称 As 当前科室名称, d.名称 As 当前病区名称, Nvl(e.名称, a.工作单位) As 工作单位名称, a.Ic卡号
    From 病人信息 A, 部门表 C, 部门表 D, 合约单位 E
    Where a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.合同单位id = e.Id(+) And a.停用时间 Is Null And Rownum < 1;
  r_病人 c_病人基本信息%RowType;

  Type Ty_病人信息 Is Ref Cursor;
  c_病人信息 Ty_病人信息; --动态游标变量
Begin
  --解析入参
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');

  Begin
    c_病人ids := j_Json.Get_Clob('pati_ids');
  Exception
    When Others Then
      c_病人ids := Null;
  End;
  v_就诊卡号 := j_Json.Get_String('vcard_no');
  While c_病人ids Is Not Null Loop
    If Length(c_病人ids) <= 4000 Then
      l_病人id.Extend;
      l_病人id(l_病人id.Count) := c_病人ids;
      c_病人ids := Null;
    Else
      l_病人id.Extend;
      l_病人id(l_病人id.Count) := Substr(c_病人ids, 1, Instr(c_病人ids, ',', 3980) - 1);
      c_病人ids := Substr(c_病人ids, Instr(c_病人ids, ',', 3980) + 1);
    End If;
  End Loop;

  If l_病人id.Count <> 0 Then
    Json_Out := '{"output":{"code":1,"message":"成功","pati_list":[';
    For I In 1 .. l_病人id.Count Loop
      For c_病人查询 In (With c_病人 As
                        (Select Column_Value As 病人id From Table(f_Num2list(l_病人id(I))))
                       Select /*+cardinality(B,10)*/
                        a.病人id, a.姓名, a.性别, a.年龄, To_Char(a.出生日期, 'yyyy-mm-dd hh24:mi:ss') As 出生日期, a.身份证号, a.门诊号, a.住院号,
                        a.就诊卡号, a.费别, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.家庭地址, a.工作单位, a.住院次数, a.当前科室id, a.当前病区id,
                        To_Char(a.入院时间, 'yyyy-mm-dd hh24:mi:ss') As 入院时间,
                        To_Char(a.出院时间, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.险类,
                        To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间, a.在院, a.当前床号, a.病人类型, a.主页id, c.名称 As 当前科室名称,
                        d.名称 As 当前病区名称, Nvl(e.名称, a.工作单位) As 工作单位名称, a.Ic卡号
                       From 病人信息 A, c_病人 B, 部门表 C, 部门表 D, 合约单位 E
                       Where a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.合同单位id = e.Id(+) And a.停用时间 Is Null And
                             a.病人id = b.病人id) Loop
        Zljsonputvalue(v_List, 'pati_id', c_病人查询.病人id, 1, 1);
        Zljsonputvalue(v_List, 'pati_name', c_病人查询.姓名, 0);
        Zljsonputvalue(v_List, 'pati_sex', Nvl(c_病人查询.性别, ''), 0);
        Zljsonputvalue(v_List, 'pati_age', Nvl(c_病人查询.年龄, ''), 0);
        Zljsonputvalue(v_List, 'pati_birthdate', Nvl(c_病人查询.出生日期, ''), 0);
        Zljsonputvalue(v_List, 'fee_category', Nvl(c_病人查询.费别, ''), 0);
        Zljsonputvalue(v_List, 'outpatient_num', c_病人查询.门诊号, 0);
        Zljsonputvalue(v_List, 'pati_nation', Nvl(c_病人查询.民族, ''), 0);
        Zljsonputvalue(v_List, 'pati_idcard', Nvl(c_病人查询.身份证号, ''), 0);
        Zljsonputvalue(v_List, 'vcard_no', Nvl(c_病人查询.就诊卡号, ''), 0);
        Zljsonputvalue(v_List, 'pati_education', Nvl(c_病人查询.学历, ''), 0);
        Zljsonputvalue(v_List, 'ocpt_name', Nvl(c_病人查询.职业, ''), 0);
        Zljsonputvalue(v_List, 'pati_identity', Nvl(c_病人查询.身份, ''), 0);
        Zljsonputvalue(v_List, 'country_name', Nvl(c_病人查询.国籍, ''), 0);
        Zljsonputvalue(v_List, 'pat_home_addr', Nvl(c_病人查询.家庭地址, ''), 0);
        Zljsonputvalue(v_List, 'pati_area', Nvl(c_病人查询.区域, ''), 0);
        Zljsonputvalue(v_List, 'emp_name', Nvl(c_病人查询.工作单位名称, ''), 0);
        Zljsonputvalue(v_List, 'pati_type', Nvl(c_病人查询.病人类型, ''), 0);
        Zljsonputvalue(v_List, 'insurance_type', Nvl(c_病人查询.险类, ''), 0);
        Zljsonputvalue(v_List, 'pati_dept_id', c_病人查询.当前科室id, 1);
        Zljsonputvalue(v_List, 'pati_dept_name', Nvl(c_病人查询.当前科室名称, ''), 0);
        Zljsonputvalue(v_List, 'create_time', Nvl(c_病人查询.登记时间, ''), 0);
        Zljsonputvalue(v_List, 'iccard_no', Nvl(c_病人查询.Ic卡号, ''), 0, 2);
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
    Return;
  Elsif v_就诊卡号 Is Not Null Then
  
    Open c_病人信息 For
      Select a.病人id, a.姓名, a.性别, a.年龄, To_Char(a.出生日期, 'yyyy-mm-dd hh24:mi:ss') As 出生日期, a.身份证号, a.门诊号, a.住院号, a.就诊卡号,
             a.费别, a.身份, a.职业, a.民族, a.国籍, a.区域, a.学历, a.家庭地址, a.工作单位, a.住院次数, a.当前科室id, a.当前病区id,
             To_Char(a.入院时间, 'yyyy-mm-dd hh24:mi:ss') As 入院时间, To_Char(a.出院时间, 'yyyy-mm-dd hh24:mi:ss') As 出院时间, a.险类,
             To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间, a.在院, a.当前床号, a.病人类型, a.主页id, c.名称 As 当前科室名称,
             d.名称 As 当前病区名称, Nvl(e.名称, a.工作单位) As 工作单位名称, a.Ic卡号
      From 病人信息 A, 部门表 C, 部门表 D, 合约单位 E
      Where a.当前科室id = c.Id(+) And a.当前病区id = d.Id(+) And a.合同单位id = e.Id(+) And a.停用时间 Is Null And a.就诊卡号 = v_就诊卡号;
  Else
    v_Err_Msg := '未传入有效的查询条件，不能获取病人信息!';
    Json_Out  := Zljsonout(v_Err_Msg);
    Return;
  End If;

  Json_Out := '{"output":{"code":1,"message":"成功","pati_list":[';
  Loop
    Fetch c_病人信息
      Into r_病人;
    Exit When c_病人信息%NotFound;
  
    Zljsonputvalue(v_List, 'pati_id', r_病人.病人id, 1, 1);
    Zljsonputvalue(v_List, 'pati_name', r_病人.姓名, 0);
    Zljsonputvalue(v_List, 'pati_sex', Nvl(r_病人.性别, ''), 0);
    Zljsonputvalue(v_List, 'pati_age', Nvl(r_病人.年龄, ''), 0);
    Zljsonputvalue(v_List, 'pati_birthdate', Nvl(r_病人.出生日期, ''), 0);
    Zljsonputvalue(v_List, 'fee_category', Nvl(r_病人.费别, ''), 0);
    Zljsonputvalue(v_List, 'outpatient_num', r_病人.门诊号, 0);
    Zljsonputvalue(v_List, 'pati_nation', Nvl(r_病人.民族, ''), 0);
    Zljsonputvalue(v_List, 'pati_idcard', Nvl(r_病人.身份证号, ''), 0);
    Zljsonputvalue(v_List, 'vcard_no', Nvl(r_病人.就诊卡号, ''), 0);
    Zljsonputvalue(v_List, 'pati_education', Nvl(r_病人.学历, ''), 0);
    Zljsonputvalue(v_List, 'ocpt_name', Nvl(r_病人.职业, ''), 0);
    Zljsonputvalue(v_List, 'pati_identity', Nvl(r_病人.身份, ''), 0);
    Zljsonputvalue(v_List, 'country_name', Nvl(r_病人.国籍, ''), 0);
    Zljsonputvalue(v_List, 'pat_home_addr', Nvl(r_病人.家庭地址, ''), 0);
    Zljsonputvalue(v_List, 'pati_area', Nvl(r_病人.区域, ''), 0);
    Zljsonputvalue(v_List, 'emp_name', Nvl(r_病人.工作单位名称, ''), 0);
    Zljsonputvalue(v_List, 'pati_type', Nvl(r_病人.病人类型, ''), 0);
    Zljsonputvalue(v_List, 'insurance_type', Nvl(r_病人.险类, ''), 0);
    Zljsonputvalue(v_List, 'pati_dept_id', r_病人.当前科室id, 1);
    Zljsonputvalue(v_List, 'pati_dept_name', Nvl(r_病人.当前科室名称, ''), 0);
    Zljsonputvalue(v_List, 'create_time', Nvl(r_病人.登记时间, ''), 0);
    Zljsonputvalue(v_List, 'iccard_no', Nvl(r_病人.Ic卡号, ''), 0, 2);
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
End Zl_Patisvr_Getvisitpatis;
/
Create Or Replace Procedure Zl_Patisvr_Lockcheck
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：检查病人是否锁定
  --入参 JSON格式
  --input
  --  pati_id     N 1 病人id
  --出参：JSON格式
  --output
  --  code        N 1 应答吗：0-失败；1-成功
  --  message     C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  v_姓名      Varchar2(2550);
  v_锁定      Number;
  j_Json      Pljson;
  j_Jsoninput Pljson;
  n_病人id    Number;
Begin
  j_Jsoninput := Pljson(Json_In);
  j_Json      := j_Jsoninput.Get_Pljson('input');
  n_病人id    := j_Json.Get_Number('pati_id');
  Select 姓名, 锁定 Into v_姓名, v_锁定 From 病人信息 Where 病人id = n_病人id;
  If Nvl(v_锁定, 0) = 1 Then
    Json_Out := '{"output":{"code":0,"message":"病人【' || v_姓名 || '】当前已被锁定不允许进行任何操作，请等待一定时间后再试。"}}';
  Else
    Json_Out := '{"output":{"code":1,"message":"成功"}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Lockcheck;
/
Create Or Replace Procedure Zl_Patisvr_Newpatiarchives
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ------------------------------------------------------------------------
  --功能：新病人建档
  --入参      json
  --  input
  --    pati_id               N  1  病人id
  --    pati_name             C  1  姓名
  --    pati_sex              C  1  性别
  --    pati_age              C  1  年龄
  --    pati_birthdate        C  1  出生日期:yyyy-mm-dd hh24:mi:ss
  --    pati_idcard           C  1  身份证号
  --    pati_type             C  1  病人类型(普通，医保，留观)
  --    outpatient_num        C  1  门诊号
  --    vcard_no              C  1  就诊卡号
  --    vcard_pwd             C  1  卡验证码
  --    fee_category          C  1  费别
  --    mdlpay_mode_name      C  1  医疗付款方式名称
  --    native_place          C  1  籍贯
  --    country_name          C  1  国籍
  --    nation_name           C  1  民族
  --    mari_status           C  1  婚姻状况
  --    edu_name              C  1  学历
  --    ocpt_name             C  1  职业
  --    pati_identity         C  1  身份
  --    emp_name              C  1  工作单位
  --    emp_postcode          C  1  单位邮编
  --    emp_phno              C  1  单位电话
  --    emp_bank_name       C   1   单位开户行
  --    emp_bank_accnum     C   1   单位帐号
  --    ctt_unit_id           N  1  合同单位id
  --    pat_home_addr         C  1  家庭地址
  --    pat_home_phno         C  1  家庭电话
  --    pat_home_postcode     C  1  家庭地址邮编
  --    region                C  1  区域
  --    pat_baddr             C  1  出生地点
  --    pat_hous_addr         C  1  户口地址
  --    pat_hous_postcode     C  1  户口地址邮编
  --    pat_grdn_name         C  1  监护人
  --    phone_number          C  1  手机号
  --    insurance_num         C  1  医保号
  --    iccard_no             C  1  Ic卡号
  --    create_time           C  1  登记时间:yyyy-mm-dd hh24:mi:ss
  --    operator_name         C  1  操作员姓名
  --    idcard_sign           N     身份证签约
  --    idcard_sign_pwd       C     签约密码
  --    insurance_type      N   1   险类
  --    cert_no_other       C   1   其他证件
  --    contacts              C     更新联系人信息节点
  --      name                C  1  联系人姓名
  --      idcard              C  1  联系人身份证号
  --      phone               C  1  联系人电话
  --      relation            C  1  联系人关系
  --      address             C     联系人地址
  --    community_info        C     社区信息节点
  --      num                 N  1  社区序号
  --      code                C  1  社区号码
  --      oper_type           N  1  社区操作类型
  --    visit_info            C     就诊信息节点
  --      statu               N     更新的就诊状态
  --      room                C     更新的就诊诊室
  --      time                C     就诊时间:yyyy-mm-dd hh24:mi:ss
  --    addr_list[]           C     地址信息列表
  --      oper_fun            N  1  操作功能:1-新增,修改   2-删除
  --      type                C  1  地址类别
  --      state               C  1  地址_省
  --      city                C  1  地址_市
  --      county              C  1  地址_县
  --      township            C  1  地址_乡
  --      other               C  1  地址_其他
  --      code                C  1  区划代码
  --    ext_list[]            C     病人信息从项列表
  --      info_name           C  1  信息名
  --      upd_info_value      N  1  修改的信息值
  --    cert_list[]                 证件列表(主要是当成绑卡处理)
  --      cert_name           C  1  证件名称
  --      cert_no             C  1  证号号码
  --    allergic_drugs_list[]       病人过敏药物列表:有数据时，是先删除过敏药物插入的方式
  --      pat_algc_cadn_id    N  1  过敏药品ID
  --      pat_algc_cadn       C  1  过敏药物名称
  --      allergy_info        C  1  过每药物反应
  --    immune_list[]         C     病人免疫列表
  --      vaccinate_time      C  1  接种时间:yyyy-mm-dd hh24:mi:ss
  --      vaccinate_name      C  1  接种名称
  --    card_property_list[]  C     医疗卡属性列表
  --      cardtype_id         N  1  医疗卡类别ID
  --      card_no             C  1  卡号
  --      info_name           C  1  信息名
  --      info_value          N  1  信息值

  --出参      json
  --  output
  --    code  N 1 应答码：0-失败；1-成功
  --    message C 1应答消息：成功时返回成功信息
  -----------------------------------------------------------------------------------------------------
  n_病人id         病人信息.病人id%Type;
  v_姓名           病人信息.姓名%Type;
  v_身份证号       病人信息.身份证号%Type;
  v_病人类型       病人信息.病人类型%Type;
  v_年龄           病人信息.年龄%Type;
  v_性别           病人信息.性别%Type;
  d_出生日期       病人信息.出生日期%Type;
  v_年龄单位       Varchar2(20);
  v_手机号         病人信息.手机号%Type;
  v_家庭电话       病人信息.家庭电话%Type;
  n_门诊号         病人信息.门诊号%Type;
  v_费别           病人信息.费别%Type;
  v_医疗付款方式   病人信息.医疗付款方式%Type;
  v_国籍           病人信息.国籍%Type;
  v_籍贯           病人信息.籍贯%Type;
  v_民族           病人信息.民族%Type;
  v_婚姻           病人信息.婚姻状况%Type;
  v_职业           病人信息.职业%Type;
  v_学历           病人信息.学历%Type;
  v_工作单位       病人信息.工作单位%Type;
  n_合同单位id     病人信息.合同单位id%Type;
  v_单位电话       病人信息.单位电话%Type;
  v_单位邮编       病人信息.单位邮编%Type;
  v_单位开户行     病人信息.单位开户行%Type;
  v_单位帐号       病人信息.单位帐号%Type;
  v_家庭地址       病人信息.家庭地址%Type;
  v_家庭地址邮编   病人信息.家庭地址邮编%Type;
  v_户口地址       病人信息.户口地址%Type;
  v_户口地址邮编   病人信息.户口地址邮编%Type;
  d_登记时间       病人信息.登记时间%Type;
  v_医保号         病人信息.医保号%Type;
  v_区域           病人信息.区域%Type;
  v_联系人身份证号 病人信息.联系人身份证号%Type;
  v_联系人姓名     病人信息.联系人姓名%Type;
  v_联系人电话     病人信息.联系人电话%Type;
  v_联系人关系     病人信息.联系人关系%Type;
  v_联系人地址     病人信息.联系人地址%Type;
  v_监护人         病人信息.监护人%Type;
  v_出生地点       病人信息.出生地点%Type;
  v_身份           病人信息.身份%Type;
  v_操作员姓名     病人医疗卡信息.发卡人%Type;
  n_社区id         病人社区信息.社区%Type;
  v_社区号码       病人社区信息.社区号%Type;
  n_社区类型       病人社区信息.就诊类型%Type;
  n_地址类别       病人地址信息.地址类别%Type;
  v_地址_省        病人地址信息.省%Type;
  v_地址_市        病人地址信息.市%Type;
  v_地址_县        病人地址信息.县%Type;
  v_地址_乡        病人地址信息.乡镇%Type;
  v_地址_其他      病人地址信息.其他%Type;
  v_区划代码       病人地址信息.区划代码%Type;
  v_信息名         病人信息从表.信息名%Type;
  v_信息值         病人信息从表.信息值%Type;
  v_卡名称         医疗卡类别.名称%Type;
  n_卡类别id       医疗卡类别.Id%Type;
  v_卡号           病人医疗卡信息.卡号%Type;
  n_卡号长度       医疗卡类别.卡号长度%Type;
  v_编码           医疗卡类别.编码%Type;
  v_卡密码         病人医疗卡信息.密码%Type;
  v_变动原因       病人医疗卡变动.变动原因%Type;
  d_终止使用时间   病人医疗卡信息.终止使用时间%Type;
  n_过敏药品id     病人过敏药物.过敏药物id%Type;
  v_过敏药物名称   病人过敏药物.过敏药物%Type;
  v_过每药物反应   病人过敏药物.过敏反应%Type;
  d_接种时间       病人免疫记录.接种时间%Type;
  v_接种名称       病人免疫记录.接种名称%Type;
  v_就诊卡号       病人信息.就诊卡号%Type;
  v_卡验证码       病人信息.卡验证码%Type;
  v_Ic卡号         病人信息.Ic卡号%Type;
  n_操作功能       Number(2);
  n_唯一身份证     Number(2);
  n_就诊状态       病人信息.就诊状态%Type;
  v_就诊诊室       病人信息.就诊诊室%Type;
  d_就诊时间       病人信息.就诊时间%Type;
  n_险类           病人信息.险类%Type;
  v_其他证件       病人信息.其他证件%Type;
  n_最长值         Number(10);
  n_最大值         Number(10);

  n_Count    Number(2);
  j_Input    Pljson;
  o_Json     Pljson;
  o_Json1    Pljson;
  j_Jsonlist Pljson_List := Pljson_List();
  j_Jsonin   Pljson;
Begin
  j_Jsonin := Pljson(Json_In);
  j_Input  := j_Jsonin.Get_Pljson('input');
  --    pati_id               N  1  病人id
  --    pati_name             C  1  姓名
  --    pati_sex              C  1  性别
  --    pati_age              C  1  年龄
  --    pati_birthdate        C  1  出生日期:yyyy-mm-dd hh24:mi:ss
  --    pati_type             C   1   病人类型(普通，医保，留观)
  n_病人id   := j_Input.Get_Number('pati_id');
  v_姓名     := j_Input.Get_String('pati_name');
  v_性别     := j_Input.Get_String('pati_sex');
  v_年龄     := j_Input.Get_String('pati_age');
  d_出生日期 := To_Date(j_Input.Get_String('pati_birthdate'), 'YYYY-MM-DD hh24:mi:ss');
  v_病人类型 := j_Input.Get_String('pati_type');
  --    pati_idcard           C  1  身份证号
  --    outpno                N  1  门诊号
  --    vcard_no              C  1  就诊卡号
  --    vcard_pwd             C  1  卡验证码
  --    fee_category          C  1  费别
  --    mdlpay_mode_name      C  1  医疗付款方式名称
  v_身份证号     := j_Input.Get_String('pati_idcard');
  n_门诊号       := To_Number(j_Input.Get_String('outpatient_num'));
  v_就诊卡号     := j_Input.Get_String('vcard_no');
  v_卡验证码     := j_Input.Get_String('vcard_pwd');
  v_费别         := j_Input.Get_String('fee_category');
  v_医疗付款方式 := j_Input.Get_String('mdlpay_mode_name');

  --    native_place          C  1  籍贯
  --    country_name          C  1  国籍
  --    nation_name           C  1  民族
  --    mari_status           C  1  婚姻状况
  --    ocpt_name             C  1  职业
  --    edu_name              C  1  学历
  --    pati_identity         C  1  身份
  v_籍贯 := j_Input.Get_String('native_place');
  v_国籍 := j_Input.Get_String('country_name');
  v_民族 := j_Input.Get_String('nation_name');
  v_婚姻 := j_Input.Get_String('mari_status');
  v_职业 := j_Input.Get_String('ocpt_name');
  v_身份 := j_Input.Get_String('pati_identity');
  v_学历 := j_Input.Get_String('edu_name');
  --    emp_name              C  1  工作单位
  --    emp_postcode          C  1  单位邮编
  --    emp_phno              C  1  单位电话
  --    emp_bank_name       C   1   单位开户行
  --    emp_bank_accnum     C   1   单位帐号
  --    ctt_unit_id           N  1  合同单位id
  --    pat_home_addr         C  1  家庭地址
  --    pat_home_phno         C  1  家庭电话
  --    pat_home_postcode     C  1  家庭地址邮编
  v_工作单位     := j_Input.Get_String('emp_name');
  v_单位电话     := j_Input.Get_String('emp_phno');
  v_单位邮编     := j_Input.Get_String('emp_postcode');
  v_单位开户行   := j_Input.Get_String('emp_bank_name');
  v_单位帐号     := j_Input.Get_String('emp_bank_accnum');
  n_合同单位id   := j_Input.Get_Number('ctt_unit_id');
  v_家庭地址     := j_Input.Get_String('pat_home_addr');
  v_家庭电话     := j_Input.Get_String('pat_home_phno');
  v_家庭地址邮编 := j_Input.Get_String('pat_home_postcode');

  --    region                C  1  区域
  --    pat_baddr             C  1  出生地点
  --    pat_hous_addr         C  1  户口地址
  --    pat_hous_postcode     C  1  户口地址邮编

  v_区域         := j_Input.Get_String('region');
  v_出生地点     := j_Input.Get_String('pat_baddr');
  v_户口地址     := j_Input.Get_String('pat_hous_addr');
  v_户口地址邮编 := j_Input.Get_String('pat_hous_postcode');

  --    pat_grdn_name         C  1  监护人
  --    phone_number          C  1  手机号
  --    insurance_num         C  1  医保号
  --    iccard_no             C  1  Ic卡号
  --    insurance_type        N  1  险类
  --    create_time           C  1  登记时间:yyyy-mm-dd hh24:mi:ss
  --    operator_name         C  1  操作员姓名
  --    idcard_sign           N     身份证签约
  --    idcard_sign_pwd       C     签约密码

  v_监护人   := j_Input.Get_String('pat_grdn_name');
  v_手机号   := j_Input.Get_String('phone_number');
  v_医保号   := j_Input.Get_String('insurance_num');
  v_Ic卡号   := j_Input.Get_String('iccard_no');
  d_登记时间 := To_Date(j_Input.Get_String('create_time'), 'YYYY-MM-DD hh24:mi:ss');
  If d_登记时间 Is Null Then
    d_登记时间 := Sysdate;
  End If;
  v_操作员姓名 := j_Input.Get_String('operator_name');
  --    insurance_type      N   1   险类
  --    cert_no_other       C   1   其他证件
  n_险类     := j_Input.Get_Number('insurance_type');
  v_其他证件 := j_Input.Get_String('cert_no_other');

  If v_身份证号 Is Not Null Then
    n_唯一身份证 := Nvl(zl_GetSysParameter(279), 0);
    If n_唯一身份证 = 1 Then
      --检查身份证唯一性
      Select Count(1) Into n_Count From 病人信息 Where 身份证号 = v_身份证号 And 病人id <> n_病人id;
      If n_Count <> 0 Then
        Json_Out := Zljsonout('已经存在身份证号为' || v_身份证号 || '的病人,不能再录入相同的身份证号！');
        Return;
      End If;
    End If;
  End If;

  If d_出生日期 Is Null And v_年龄 Is Not Null Then
    --根据年龄求出生日期
    v_年龄单位 := Substr(v_年龄, Length(v_年龄), 1);
    If Instr('岁,月,天', v_年龄单位) <= 0 Then
      v_年龄单位 := Null;
    Else
      v_年龄 := Replace(v_年龄, v_年龄单位, '');
    End If;
    Begin
      v_年龄 := To_Number(v_年龄);
    Exception
      When Others Then
        v_年龄 := Null;
    End;
    If v_年龄 Is Not Null And v_年龄单位 Is Not Null Then
      Select Decode(v_年龄单位, '岁', Add_Months(Sysdate, -12 * v_年龄), '月', Add_Months(Sysdate, -1 * v_年龄), '天',
                     Sysdate - v_年龄)
      Into d_出生日期
      From Dual;
    End If;
  End If;

  If v_手机号 Is Null And v_家庭电话 Is Not Null Then
    Select Count(1)
    Into n_Count
    From 手机号常用号段表
    Where Length(v_家庭电话) = 号码长度 And v_家庭电话 Like 号段 || '%';
    If n_Count <> 0 Then
      v_手机号 := v_家庭电话;
    End If;
  End If;

  --联系人信息
  --    contacts              C     更新联系人信息节点
  --      name                C  1  联系人姓名
  --      idcard              C  1  联系人身份证号
  --      phone               C  1  联系人电话
  --      relation            C  1  联系人关系
  --      address             C     联系人地址
  o_Json1 := Pljson();
  o_Json1 := j_Input.Get_Pljson('contacts');
  If o_Json1 Is Not Null Then
    v_联系人姓名     := o_Json1.Get_String('name');
    v_联系人身份证号 := o_Json1.Get_String('idcard');
    v_联系人电话     := o_Json1.Get_String('phone');
    v_联系人关系     := o_Json1.Get_String('relation');
    v_联系人地址     := o_Json1.Get_String('address');
  End If;

  --就诊信息
  --    visit_info            C     就诊信息节点
  --      visit_statu         N     更新的就诊状态
  --      visit_room          C     更新的就诊诊室
  --      visit_time          C     就诊时间:yyyy-mm-dd hh24:mi:ss
  o_Json1 := Pljson();
  o_Json1 := j_Input.Get_Pljson('visit_info');
  If o_Json1 Is Not Null Then
    n_就诊状态 := o_Json1.Get_Number('visit_statu');
    v_就诊诊室 := o_Json1.Get_String('visit_room');
    d_就诊时间 := To_Date(o_Json1.Get_String('visit_time'), 'yyyy-mm-dd hh24:mi:ss');
  End If;

  --新病人信息
  Insert Into 病人信息
    (病人id, 门诊号, 姓名, 性别, 年龄, 出生日期, 费别, 医疗付款方式, 国籍, 民族, 籍贯, 婚姻状况, 职业, 学历, 病人类型, 身份证号, 工作单位, 单位开户行, 单位帐号, 合同单位id, 单位电话,
     单位邮编, 家庭地址, 家庭电话, 家庭地址邮编, 户口地址, 户口地址邮编, 登记时间, 医保号, 区域, 联系人身份证号, 联系人姓名, 联系人电话, 联系人关系, 联系人地址, 监护人, 出生地点, 手机号, 身份,
     就诊卡号, 卡验证码, Ic卡号, 就诊状态, 就诊时间, 就诊诊室, 险类, 其他证件)
  Values
    (n_病人id, n_门诊号, v_姓名, v_性别, v_年龄, d_出生日期, v_费别, v_医疗付款方式, v_国籍, v_民族, v_籍贯, v_婚姻, v_职业, v_学历, v_病人类型, v_身份证号,
     v_工作单位, v_单位开户行, v_单位帐号, Decode(n_合同单位id, 0, Null, n_合同单位id), v_单位电话, v_单位邮编, v_家庭地址, v_家庭电话, v_家庭地址邮编, v_户口地址,
     v_户口地址邮编, d_登记时间, v_医保号, v_区域, v_联系人身份证号, v_联系人姓名, v_联系人电话, v_联系人关系, v_联系人地址, v_监护人, v_出生地点, v_手机号, v_身份, v_就诊卡号,
     v_卡验证码, v_Ic卡号, n_就诊状态, d_就诊时间, v_就诊诊室, Decode(n_险类, 0, Null, n_险类), v_其他证件);

  --社区信息
  --    community_info        C     社区信息节点
  --      num                 N  1  社区序号
  --      code                C  1  社区号码
  --      oper_type           N  1  社区操作类型
  o_Json1 := Pljson();
  o_Json1 := j_Input.Get_Pljson('community_info');
  If o_Json1 Is Not Null Then
    n_社区id   := o_Json1.Get_Number('num');
    v_社区号码 := o_Json1.Get_String('code');
    n_社区类型 := o_Json1.Get_Number('oper_type');
  End If;
  --更新社区号
  If n_社区id <> 0 And v_社区号码 Is Not Null Then
    Zl_病人社区信息_Insert(n_病人id, n_社区id, v_社区号码, n_社区类型, d_登记时间);
  End If;

  --更新地址信息
  --    addr_list[]           C     地址信息列表
  --      oper_fun            N  1  操作功能:1-新增,修改   2-删除
  --      type                C  1  地址类别
  --      state               C  1  地址_省
  --      city                C  1  地址_市
  --      county              C  1  地址_县
  --      township            C  1  地址_乡
  --      other               C  1  地址_其他
  --      code                C  1  区划代码
  j_Jsonlist := Pljson_List();
  j_Jsonlist := j_Input.Get_Pljson_List('addr_list');
  If j_Jsonlist Is Not Null Then
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json     := Pljson();
      o_Json     := Pljson(j_Jsonlist.Get(I));
      n_操作功能  := o_Json.Get_Number('oper_fun');
      n_地址类别  := o_Json.Get_Number('type');
      v_地址_省   := o_Json.Get_String('state');
      v_地址_市   := o_Json.Get_String('city');
      v_地址_县   := o_Json.Get_String('county');
      v_地址_乡   := o_Json.Get_String('township');
      v_地址_其他 := o_Json.Get_String('other');
      v_区划代码  := o_Json.Get_String('code');
    
      Zl_病人地址信息_Update_s(n_操作功能, n_病人id, Null, n_地址类别, v_地址_省, v_地址_市, v_地址_县, v_地址_乡, v_地址_其他, v_区划代码);
    End Loop;
  End If;

  --更新病人从属信息
  --    ext_list[]            C     病人信息从项列表
  --      info_name           C  1  信息名
  --      upd_info_value      N  1  修改的信息值
  j_Jsonlist := Pljson_List();
  j_Jsonlist := j_Input.Get_Pljson_List('ext_list');
  If j_Jsonlist Is Not Null Then
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json   := Pljson();
      o_Json   := Pljson(j_Jsonlist.Get(I));
      v_信息名 := o_Json.Get_String('info_name');
      v_信息值 := o_Json.Get_String('upd_info_value');
    
      If v_信息名 Is Not Null And v_信息值 Is Not Null Then
        Zl_病人信息从表_Update(n_病人id, v_信息名, v_信息值);
      End If;
    End Loop;
  End If;

  --更新证件类型
  --    cert_list[]                 证件列表(主要是当成绑卡处理)
  --      cert_name           C  1  证件名称
  --      cert_no             C  1  证号号码
  j_Jsonlist := Pljson_List();
  j_Jsonlist := j_Input.Get_Pljson_List('cert_list');
  If j_Jsonlist Is Not Null Then
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json   := Pljson();
      o_Json   := Pljson(j_Jsonlist.Get(I));
      v_卡名称 := o_Json.Get_String('cert_name');
      v_卡号   := o_Json.Get_String('cert_no');
    
      If v_卡名称 Is Not Null Then
        If v_卡号 Is Not Null Then
          --检查卡号是否被他人使用
          Select Count(1)
          Into n_Count
          From 病人医疗卡信息 A, 医疗卡类别 B
          Where a.卡类别id = b.Id And b.名称 = v_卡名称 And b.是否证件 = 1 And a.卡号 = v_卡号 And a.病人id <> n_病人id;
          If n_Count <> 0 Then
            Json_Out := Zljsonout(v_卡名称 || ':' || v_卡号 || '正在被他人使用,请检查！');
            Return;
          End If;
        
          --不存在的就诊类型需要新增卡类别管理
          Select Nvl(Max(ID), 0), Nvl(Max(卡号长度), 0), Max(编码), Max(LPad(编码, 10)), Max(Length(编码))
          Into n_卡类别id, n_卡号长度, v_编码, n_最大值, n_最长值
          From 医疗卡类别
          Where 名称 = v_卡名称;
        
          Select Max(编码), Max(LPad(编码, 10)), Max(Length(编码)) Into v_编码, n_最大值, n_最长值 From 医疗卡类别;
        
          If v_编码 Is Null Then
            Select LPad(1, 10, '0') Into v_编码 From Dual;
          Else
            n_最大值 := n_最大值 + 1;
            Select LPad(n_最大值, n_最长值, '0') Into v_编码 From Dual;
          End If;
          If n_卡类别id = 0 Then
            --新增
            Select 医疗卡类别_Id.Nextval Into n_卡类别id From Dual;
            Zl_医疗卡类别_Update(n_卡类别id, v_编码, v_卡名称, Substr(v_卡名称, 1, 1), Null, Length(v_卡号), 0, 1, 0, 0, 0, 0, Null, Null,
                            v_编码, 0, Null, 1, Null, 1, 10, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 1000, 0, 1);
          Elsif Length(v_卡号) > n_卡号长度 Then
            --修改长度
            Zl_医疗卡类别_Update(n_卡类别id, v_编码, v_卡名称, Substr(v_卡名称, 1, 1), Null, Length(v_卡号), 0, 1, 0, 0, 0, 0, Null, Null,
                            v_编码, 0, Null, 1, Null, 1, 10, 0, 0, 1, 1, 0, 0, 0, 0, 0, 0, 0, 1, 0, 1000, 0, 1);
          End If;
        End If;
      
        --先清除病人卡信息
        n_Count := 0;
        For c_证件 In (Select a.卡类别id, a.卡号
                     From 病人医疗卡信息 A
                     Where a.卡类别id = n_卡类别id And a.病人id = n_病人id) Loop
          If c_证件.卡号 = Nvl(v_卡号, '_') Then
            n_Count := 1;
          Else
            Zl_医疗卡变动_Insert_s(14, n_病人id, c_证件.卡类别id, Null, c_证件.卡号, '证件卡取消绑定', Null, v_操作员姓名, d_登记时间);
          End If;
        End Loop;
        --新增病人卡信息
        If n_Count = 0 And v_卡号 Is Not Null Then
          Zl_医疗卡变动_Insert_s(11, n_病人id, n_卡类别id, Null, v_卡号, '证件卡绑定', Null, v_操作员姓名, d_登记时间);
        End If;
      End If;
    End Loop;
  End If;

  --更新过敏数据
  j_Jsonlist := Pljson_List();
  j_Jsonlist := j_Input.Get_Pljson_List('allergic_drugs_list');
  If j_Jsonlist Is Not Null Then
    --清除所有记录
    Zl_病人过敏药物_Delete(n_病人id);
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json         := Pljson();
      o_Json         := Pljson(j_Jsonlist.Get(I));
      n_过敏药品id   := o_Json.Get_Number('pat_algc_cadn_id');
      v_过敏药物名称 := o_Json.Get_String('pat_algc_cadn');
      v_过每药物反应 := o_Json.Get_String('allergy_info');
    
      If v_过敏药物名称 Is Not Null Then
        Zl_病人过敏药物_Update(n_病人id, n_过敏药品id, v_过敏药物名称, v_过每药物反应);
      End If;
    End Loop;
  End If;

  --更新免疫记录
  --    immune_list[]         C     病人免疫列表
  --      vaccinate_time      C  1  接种时间:yyyy-mm-dd hh24:mi:ss
  --      vaccinate_name      C  1  接种名称
  j_Jsonlist := Pljson_List();
  j_Jsonlist := j_Input.Get_Pljson_List('immune_list');
  If j_Jsonlist Is Not Null Then
    --清除所有记录
    Zl_病人免疫记录_Delete(n_病人id);
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json     := Pljson();
      o_Json     := Pljson(j_Jsonlist.Get(I));
      d_接种时间 := To_Date(o_Json.Get_String('vaccinate_time'), 'YYYY-MM-DD hh24:mi:ss');
      v_接种名称 := o_Json.Get_String('vaccinate_name');
    
      If v_接种名称 Is Not Null Then
        Zl_病人免疫记录_Update(n_病人id, d_接种时间, v_接种名称);
      End If;
    End Loop;
  End If;

  --更新医疗卡属性
  --    card_property_list[]  C     医疗卡属性列表
  --      cardtype_id         N  1  医疗卡类别ID
  --      card_no             C  1  卡号
  --      info_name           C  1  信息名
  --      info_value          N  1  信息值
  j_Jsonlist := Pljson_List();
  j_Jsonlist := j_Input.Get_Pljson_List('card_property_list');
  If j_Jsonlist Is Not Null Then
    --清除所有记录
    Zl_病人免疫记录_Delete(n_病人id);
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json     := Pljson();
      o_Json     := Pljson(j_Jsonlist.Get(I));
      n_卡类别id := o_Json.Get_Number('cardtype_id');
      v_卡号     := o_Json.Get_String('card_no');
      v_信息名   := o_Json.Get_String('info_name');
      v_信息值   := o_Json.Get_String('info_value');
    
      Zl_病人医疗卡属性_Update(n_病人id, n_卡类别id, v_卡号, v_信息名, v_信息值);
    End Loop;
  End If;

  --签约信息
  --    sign_info             C   签约信息
  --      card_type_id        N 1 卡类别ID
  --      card_no             C 1 卡号
  --      card_pwd            C   卡密码
  --      qrcode              C   二维码
  --      card_notes          C   变动原因
  --      card_use_endtime    C   终止使用时间
  o_Json1 := Pljson();
  o_Json1 := j_Input.Get_Pljson('sign_info');
  If o_Json1 Is Not Null Then
    n_卡类别id     := o_Json1.Get_Number('card_type_id');
    v_卡号         := o_Json1.Get_String('card_no');
    v_卡密码       := o_Json1.Get_String('card_pwd');
    v_变动原因     := o_Json1.Get_String('card_notes');
    d_终止使用时间 := To_Date(o_Json1.Get_String('card_use_endtime'), 'YYYY-MM-DD hh24:mi:ss');
    --签约
    Select Count(1) Into n_Count From 医疗卡类别 Where ID = n_卡类别id;
    If n_Count = 1 Then
      Select Count(1) Into n_Count From 病人医疗卡信息 Where 卡号 = v_卡号 And 卡类别id = n_卡类别id;
      If n_Count = 0 Then
        Zl_医疗卡变动_Insert_s(11, n_病人id, n_卡类别id, '', v_卡号, v_变动原因, v_卡密码, v_操作员姓名, d_登记时间, Null, d_终止使用时间);
      End If;
    End If;
  End If;

  Json_Out := Zljsonout('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Newpatiarchives;
/

Create Or Replace Procedure Zl_Patisvr_Patiagecheck
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ----------------------------------
  --功能：出生日期年龄检查
  --入参:Json格式
  --input
  --       pati_age         C 1 年龄
  --       pati_birthdate   C 1 年龄
  --       calcdate         C 1 计算日期
  --出参:json格式
  --output
  --       code             N 1   应答码：0-失败；1-成功
  --       message          C 1   应答消息：失败时返回具体的错误信息
  --       error_info       C 1 错误信息
  -----------------------------------
  j_Json     Pljson;
  j_Jsonin   Pljson;
  v_年龄     Varchar2(20);
  d_出生日期 Date;
  d_计算日期 Date;
  v_Info     Varchar2(32767);
Begin
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  v_年龄     := j_Json.Get_String('pati_age');
  d_出生日期 := To_Date(j_Json.Get_String('pati_birthdate'), 'YYYY-MM-DD HH24:MI:SS');
  d_计算日期 := To_Date(j_Json.Get_String('calcdate'), 'YYYY-MM-DD HH24:MI:SS');
  v_Info     := Zl_Age_Check(v_年龄, d_出生日期, d_计算日期);
  Json_Out   := '{"output":{"code":1,"message":"成功","error_info":"' || Zljsonstr(v_Info, 0) || '"}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Patiagecheck;
/
Create Or Replace Procedure Zl_Patisvr_Patiidcardcheck
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ----------------------------------
  --功能：身份证号检查
  --入参:Json格式
  --input
  --    pati_idcard           C 1 身份证号
  --    calcdate              C 1 计算日期
  --出参:json格式
  --output
  --    code                  N 1 应答码：0-失败；1-成功
  --    message               C 1 应答消息：失败时返回具体的错误信息
  --    info                  C 1 检查后通过返回的信息
  -----------------------------------
  j_Json     Pljson;
  j_Jsonin   Pljson;
  v_身份证号 Varchar2(20);
  d_计算日期 Date;
  v_Info     Varchar2(32767);
Begin
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  v_身份证号 := j_Json.Get_String('pati_idcard');
  d_计算日期 := To_Date(j_Json.Get_String('calcdate'), 'YYYY-MM-DD HH24:MI:SS');
  v_Info     := Zl_Fun_Checkidcard(v_身份证号, d_计算日期);
  Json_Out   := '{"output":{"code":1,"message":"成功","info":"' || Zljsonstr(v_Info) || '"}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Patiidcardcheck;
/
Create Or Replace Procedure Zl_Patisvr_Patirealnamecheck
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  --------------------------------------------------------------------------------------------------
  --功能:实名认证前的检查
  --入参 JSOM格式
  --input
  --  opr_fun               N 1  功能 0-新增实名信息时检查  1-修改实名信息时检查
  --  real_id               N 1  实名id  opr_fun=1 时传入
  --  pati_name             C 1 姓名
  --  pati_sex              C 1 性别
  --  pati_age              C 1 年龄
  --  pati_birthdate        C 1 出生日期
  --  pati_idcard           C 1 身份证号
  --  owner                 N 1 所有者
  --  grdn_name             C 1 陪诊人姓名
  --  grdn_sex              C 1 陪诊人性别
  --  grdn_birthdate        C 1 陪诊人出生日期
  --  grdn_idcard           C 1 陪诊人身份证号
  --  grdn_relation         C 1 陪诊人关系
  --  papers_info           C 1 证件信息拼串
  --出参 JSON格式
  --output
  --  code                  N 1 应答码：0-失败；1-成功
  --  message               C 1 应答消息：失败时返回具体的错误信息
  --  real_id               N 1 实名id
  --  pati_id               N 1 病人id
  --  new_pati              N 1 是否新病人
  --  pati_age              C 1 年龄
  --  pati_name             C 1 姓名
  --  pati_sex              C 1 性别
  --  pati_birthdate        C 1 出生日期
  --------------------------------------------------------------------------------------------------
  Type r_证件信息 Is Record(
    证件信息 Varchar2(4000));

  Type t_证件信息 Is Table Of r_证件信息;
  Rs_Sql证件信息 t_证件信息 := t_证件信息();

  Type r_证件 Is Record(
    实名id   病人实名证件.实名id%Type,
    证件id   病人实名证件.Id%Type,
    证件类型 病人实名证件.证件类型%Type,
    证件号码 病人实名证件.证件号码%Type,
    证件备注 病人实名证件.备注%Type,
    所有者   病人实名证件.所有者%Type,
    序号     Number(1));

  Type t_证件 Is Table Of r_证件;
  Rs_Sql证件 t_证件 := t_证件();

  j_Json     Pljson;
  j_Jsonin   Pljson;
  v_姓名     病人信息.姓名%Type;
  v_性别     病人信息.性别%Type;
  v_年龄     病人信息.年龄%Type; --更新前的年龄
  v_出生日期 病人信息.出生日期%Type;

  v_身份证号       病人实名信息.身份证号%Type;
  n_所有者         病人实名证件.所有者%Type;
  v_陪诊人姓名     病人实名信息.陪诊人姓名%Type;
  v_陪诊人身份证号 病人实名信息.陪诊人身份证号%Type;
  v_陪诊人性别     病人实名信息.陪诊人性别%Type;
  v_陪诊人出生日期 病人实名信息.陪诊人出生日期%Type;
  v_陪诊人关系     病人实名信息.陪诊人关系%Type;

  n_New    Number(1);
  n_Id     Number(18);
  n_病人id Number(18);
  n_实名id Number(18);
  n_Count  Number(5);

  v_证件所有者  Varchar2(200);
  v_证件号码    Varchar2(400);
  v_证件信息_In Varchar2(4000);
  v_证件信息    Varchar2(4000);
  n_检查        Number;
  n_Realid      Number;

  t_Key   t_Strlist;
  v_Error Varchar2(200);
Begin
  j_Jsonin         := Pljson(Json_In);
  j_Json           := j_Jsonin.Get_Pljson('input');
  n_检查           := j_Json.Get_Number('opr_fun');
  n_Realid         := j_Json.Get_Number('real_id');
  v_姓名           := j_Json.Get_String('pati_name');
  v_性别           := j_Json.Get_String('pati_sex');
  v_年龄           := j_Json.Get_String('pati_age');
  v_出生日期       := To_Date(j_Json.Get_String('pati_birthdate'), 'yyyy-mm-dd hh24:mi:ss');
  v_身份证号       := j_Json.Get_String('pati_idcard');
  n_所有者         := j_Json.Get_Number('owner');
  v_陪诊人姓名     := j_Json.Get_String('grdn_name');
  v_陪诊人身份证号 := j_Json.Get_String('grdn_idcard');
  v_陪诊人性别     := j_Json.Get_String('grdn_sex');
  v_陪诊人出生日期 := To_Date(j_Json.Get_String('grdn_birthdate'), 'yyyy-mm-dd hh24:mi:ss');
  v_陪诊人关系     := j_Json.Get_String('grdn_relation');
  v_证件信息_In    := j_Json.Get_String('papers_info');
  --检查必录信息,必须录入病人的姓名、性别、出生日期
  If v_姓名 Is Null Then
    v_Error := '必须录入病人姓名！';
  Elsif v_性别 Is Null Then
    v_Error := '必须录入病人性别！';
  Elsif v_出生日期 Is Null Then
    v_Error := '必须录入病人出生日期！';
  End If;
  If v_Error Is Not Null Then
    Json_Out := Zljsonout(v_Error, 0);
    Return;
  End If;
  --截取证件信息
  For X In (Select Column_Value As 证件信息 From Table(f_Str2list(v_证件信息_In, ','))) Loop
    Rs_Sql证件信息.Extend;
    Rs_Sql证件信息(Rs_Sql证件信息.Count).证件信息 := x.证件信息;
  End Loop;

  For I In 1 .. Rs_Sql证件信息.Count Loop
    Rs_Sql证件.Extend;
    Select Column_Value Bulk Collect Into t_Key From Table(f_Str2list(Rs_Sql证件信息(I).证件信息, '-'));
    Rs_Sql证件(Rs_Sql证件.Count).实名id := t_Key(1);
  
    Rs_Sql证件(Rs_Sql证件.Count).证件id := t_Key(2);
  
    Rs_Sql证件(Rs_Sql证件.Count).证件类型 := t_Key(3);
  
    Rs_Sql证件(Rs_Sql证件.Count).证件号码 := t_Key(4);
  
    Rs_Sql证件(Rs_Sql证件.Count).证件备注 := t_Key(5);
  
    Rs_Sql证件(Rs_Sql证件.Count).所有者 := To_Number(t_Key(6));
  
    Rs_Sql证件(Rs_Sql证件.Count).序号 := I;
  End Loop;

  For N In 1 .. Rs_Sql证件.Count Loop
    v_证件所有者 := v_证件所有者 || Rs_Sql证件(N).所有者;
    v_证件号码   := v_证件号码 || Rs_Sql证件(N).证件号码;
    v_证件信息   := v_证件信息 || ',' || Rs_Sql证件(N).证件类型 || ',' || Rs_Sql证件(N).证件号码 || ',' || Rs_Sql证件(N).所有者;
  End Loop;
  If (v_陪诊人姓名 Is Null And v_陪诊人身份证号 Is Null) And (v_陪诊人姓名 Is Null And v_证件号码 Is Null) Then
    --在没有输入陪诊人信息的情况下，病人身份证号和其他证件号必须录入一个
    If v_身份证号 Is Null And v_证件号码 Is Null Then
      v_Error := '必须录入病人身份证号或者其他证件号码！';
    End If;
  Else
    --录入了陪诊人信息后必须录入v_陪诊人性别、陪诊人出生日期、陪诊人关系
    If (Not v_陪诊人姓名 Is Null And Not v_陪诊人身份证号 Is Null) Or (n_所有者 = 2 And Not v_陪诊人姓名 Is Null And Not v_证件号码 Is Null) Then
      If v_陪诊人性别 Is Null Then
        v_Error := '必须录入陪诊人性别！';
      Elsif v_陪诊人出生日期 Is Null Then
        v_Error := '必须录入陪诊人出生日期！';
      Elsif v_陪诊人关系 Is Null Then
        v_Error := '必须录入陪诊人关系！';
      End If;
    End If;
  End If;
  If v_Error Is Not Null Then
    Json_Out := Zljsonout(v_Error, 0);
    Return;
  End If;
  If Nvl(n_检查, 0) = 0 Then
    --查询是否有重复的病人实名信息
    --第一种情况：姓名+身份证号
    If Nvl(v_身份证号, '|') <> '|' Then
      Select Count(1) Into n_Count From 病人实名信息 Where 身份证号 = v_身份证号;
      If n_Count > 0 Then
        v_Error := '已经存在身份证号为【' || v_身份证号 || '】的实名认证信息,请检查！';
      End If;
    End If;
    If v_Error Is Not Null Then
      Json_Out := Zljsonout(v_Error, 0);
      Return;
    End If;
    n_Count := 0;
    --第二种情况：姓名+其他证件类型+其他证件号码(本人的)
    If n_所有者 = 1 Then
      Select Count(1)
      Into n_Count
      From 病人实名信息 A, 病人实名证件 B
      Where a.实名id = b.实名id And a.姓名 = v_姓名 And
            Instr(v_证件信息 || ',', ',' || b.证件类型 || ',' || b.证件号码 || ',' || b.所有者 || ',') > 0;
      If n_Count > 0 Then
        v_Error := '该病人已经存在有效的实名认证信息，不需要再次认证！';
      End If;
    End If;
    If v_Error Is Not Null Then
      Json_Out := Zljsonout(v_Error, 0);
      Return;
    End If;
    n_Count := 0;
    --第三种情况：姓名+陪诊人姓名+陪诊人身份证号
    Select Count(1)
    Into n_Count
    From 病人实名信息
    Where 姓名 = v_姓名 And 陪诊人姓名 = v_陪诊人姓名 And 陪诊人身份证号 = v_陪诊人身份证号;
    If n_Count > 0 Then
      v_Error := '该病人已经存在有效的实名认证信息，不需要再次认证！';
    End If;
    If v_Error Is Not Null Then
      Json_Out := Zljsonout(v_Error, 0);
      Return;
    End If;
    n_Count := 0;
    --第四种情况：姓名+陪诊人姓名+其他证件类型+其他证件号码(陪诊人的)
    If n_所有者 = 2 Then
      Select Count(1)
      Into n_Count
      From 病人实名信息 A, 病人实名证件 B
      Where a.实名id = b.实名id And a.姓名 = v_姓名 And a.陪诊人姓名 = v_陪诊人姓名 And
            Instr(v_证件信息 || ',', ',' || b.证件类型 || ',' || b.证件号码 || ',' || b.所有者 || ',') > 0;
      If n_Count > 0 Then
        v_Error := '该病人已经存在有效的实名认证信息，不需要再次认证！';
      End If;
    End If;
    If v_Error Is Not Null Then
      Json_Out := Zljsonout(v_Error, 0);
      Return;
    End If;
    Select 病人实名信息_实名id.Nextval Into n_实名id From Dual;
    --新建指定病人的实名认证信息
    If v_身份证号 Is Null Then
      n_New := 1;
    Else
      Select Max(病人id) As 病人id
      Into n_Id
      From (Select Nvl(Nvl(就诊时间, 入院时间), 登记时间) As 时间, 病人id
             From 病人信息
             Where 姓名 = v_姓名 And 身份证号 = v_身份证号
             Order By 时间 Desc)
      Where Rownum = 1;
      If n_Id Is Null Then
        n_New := 1;
      Else
        n_New := 0;
      End If;
    End If;
  
    If n_New = 1 Then
      Select 病人信息_Id.Nextval Into n_病人id From Dual;
    Else
      n_病人id := n_Id;
    End If;
  Else
  
    --查询是否有重复的病人实名信息
    --第一种情况：身份证号
    If Nvl(v_身份证号, '|') <> '|' Then
      Select Count(1)
      Into n_Count
      From 病人实名信息
      Where 姓名 = v_姓名 And 身份证号 = v_身份证号 And 实名id <> n_Realid;
      If n_Count > 0 Then
        v_Error := '已经存在身份证号为【' || v_身份证号 || '】的实名认证信息,请检查！';
      End If;
    End If;
    If v_Error Is Not Null Then
      Json_Out := Zljsonout(v_Error, 0);
      Return;
    End If;
    n_Count := 0;
    --第二种情况：姓名+其他证件类型+其他证件号码(本人的)
    If n_所有者 = 1 Then
      Select Count(1)
      Into n_Count
      From 病人实名信息 A, 病人实名证件 B
      Where a.实名id = b.实名id And a.姓名 = v_姓名 And a.实名id <> n_Realid And
            Instr(v_证件信息 || ',', ',' || b.证件类型 || ',' || b.证件号码 || ',' || b.所有者 || ',') > 0;
      If n_Count > 0 Then
        v_Error := '该病人已经存在有效的实名认证信息，不需要再次认证！';
      End If;
    End If;
    If v_Error Is Not Null Then
      Json_Out := Zljsonout(v_Error, 0);
      Return;
    End If;
    n_Count := 0;
    --第三种情况：姓名+陪诊人姓名+陪诊人身份证号
    Select Count(1)
    Into n_Count
    From 病人实名信息
    Where 姓名 = v_姓名 And 陪诊人姓名 = v_陪诊人姓名 And 陪诊人身份证号 = v_陪诊人身份证号 And 实名id <> n_Realid;
    If n_Count > 0 Then
      v_Error := '该病人已经存在有效的实名认证信息，不需要再次认证！';
    End If;
    If v_Error Is Not Null Then
      Json_Out := Zljsonout(v_Error, 0);
      Return;
    End If;
    n_Count := 0;
    --第四种情况：姓名+陪诊人姓名+其他证件类型+其他证件号码(陪诊人的)
    If n_所有者 = 2 Then
      Select Count(1)
      Into n_Count
      From 病人实名信息 A, 病人实名证件 B
      Where a.实名id = b.实名id And a.姓名 = v_姓名 And a.陪诊人姓名 = v_陪诊人姓名 And a.实名id <> n_Realid And
            Instr(v_证件信息 || ',', ',' || b.证件类型 || ',' || b.证件号码 || ',' || b.所有者 || ',') > 0;
      If n_Count > 0 Then
        v_Error := '该病人已经存在有效的实名认证信息，不需要再次认证！';
      End If;
    End If;
    If v_Error Is Not Null Then
      Json_Out := Zljsonout(v_Error, 0);
      Return;
    End If;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功"';
  If Nvl(n_检查, 0) = 0 Then
    Json_Out := Json_Out || ',"real_id":' || n_实名id || ',"pati_id":' || n_病人id || ',"new_pati":' || n_New || '}}';
  Else
    Json_Out := Json_Out || '}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Patirealnamecheck;
/
Create Or Replace Procedure Zl_Patisvr_Phonenumberexist
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：检查指定的手机号是否已经被使用
  --入参：Json_In:格式
  --  input
  --    pati_id              N   1  病人ID:当前操作的病人
  --    phone_number         C   1  手机号
  --出参: Json_Out,格式如下
  --  output
  --    code                 N   1   应答吗：0-失败；1-成功
  --    message              C   1   应答消息：失败时返回具体的错误信息
  --    exist                N   1   1-存在;0-不存在
  ---------------------------------------------------------------------------
  n_病人id 病人信息.病人id%Type;
  v_手机号 病人信息.手机号%Type;
  n_Exist  Number(1);
  j_Json   Pljson;
  j_Jsonin Pljson;
Begin

  --解析入参
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');
  v_手机号 := j_Json.Get_String('phone_number');

  Select Count(1) Into n_Exist From 病人信息 Where 手机号 = v_手机号 And 病人id <> n_病人id And Rownum < 2;

  Json_Out := '{"output":{"code":1,"message":"成功","exist":' || Nvl(n_Exist, 0) || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Phonenumberexist;
/
Create Or Replace Procedure Zl_Patisvr_Recalcage
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：重算病人年龄
  --入参：Json_In:格式
  --input
  --    pati_ids                   C   1  病人IDs,多个用逗号分离(病人换床同时更新两个病人年龄)
  --出参: Json_Out,格式如下
  --    output
  --        code                    N   1   应答吗：0-失败；1-成功
  --        message                 C   1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Json     Pljson;
  j_Jsonin   Pljson;
  v_Pati_Ids Varchar2(2000);
  v_Age      Varchar2(20);
Begin
  --解析入参
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  v_Pati_Ids := j_Json.Get_String('pati_ids');
  If Nvl(v_Pati_Ids, '_') = '_' Then
    Json_Out := Zljsonout('未传入病人Id,请检查！');
    Return;
  End If;
  For R In (Select /*+cardinality(a,10)*/
             Column_Value As 病人id
            From Table(f_Num2list(v_Pati_Ids)) A) Loop
  
    v_Age := Zl_Age_Calc(r.病人id);
    Update 病人信息 Set 年龄 = v_Age Where 病人id = r.病人id;
  End Loop;
  Json_Out := Zljsonout('成功', 1);
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Recalcage;
/
CREATE OR REPLACE Procedure Zl_Patisvr_Savebadrecord
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:病人不良记录数据保存
  --入参：Json_In:格式
  --    input
  --      badrec_list[]          C   保存的不良记录列表
  --        pati_id            N 1 病人id
  --        behavior_category  C 1 行为类别:如预约挂号
  --        happen_time        C 1 发生时间:yyyy-mm-dd hh24:mi:ss
  --        add_time           C 1 加入时间:yyyy-mm-dd hh24:mi:ss
  --        add_Reason         C 1 加入原因：如预约超期
  --        add_memo           C 1 加入说明
  --        additional_info    C 1 附加信息
  --        creator            C 1 登记人
  --出参: Json_Out,格式如下
  --  output
  --    code                   N   1   应答码：0-失败；1-成功
  --    message                C   1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------

  j_Json     Pljson;
  o_Json     Pljson;
  j_Jsonin   Pljson;
  j_Jsonlist Pljson_List := Pljson_List();

  n_病人id     病人不良记录.病人id%Type;
  v_行为类别   病人不良记录.行为类别%Type;
  d_加入时间   病人不良记录.加入时间%Type;
  d_发生时间   病人不良记录.发生时间%Type;
  v_加入原因   病人不良记录.加入原因%Type;
  v_加入说明   病人不良记录.加入说明%Type;
  v_附加信息   病人不良记录.附加信息%Type;
  v_登记人     病人不良记录.登记人%Type;
  v_操作员姓名 病人不良记录.登记人%Type;

Begin
  --解析入参
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  j_Jsonlist := j_Json.Get_Pljson_List('badrec_list');

  If j_Jsonlist Is Null Then
    Json_Out := Zljsonout('未存在需要保存不良记录数据，不能保存');
    Return;
  End If;

  For I In 1 .. j_Jsonlist.Count Loop
    o_Json     := Pljson();
    o_Json     := Pljson(j_Jsonlist.Get(I));
    n_病人id   := o_Json.Get_Number('pati_id');
    v_行为类别 := o_Json.Get_String('behavior_category');
    d_加入时间 := To_Date(o_Json.Get_String('add_time'), 'yyyy-mm-dd hh24:mi:ss');
    d_发生时间 := To_Date(o_Json.Get_String('happen_time'), 'yyyy-mm-dd hh24:mi:ss');
    v_加入原因 := o_Json.Get_String('add_reason');
    v_加入说明 := o_Json.Get_String('add_memo');
    v_附加信息 := o_Json.Get_String('additional_info');
    v_登记人   := o_Json.Get_String('creator');

    If Nvl(n_病人id, 0) = 0 Then
      Json_Out := Zljsonout('不能确定病人信息，请检查!');
      Return;
    End If;
    If v_登记人 Is Null And v_操作员姓名 Is Null Then
      v_操作员姓名 := zl_UserName;
    End If;

    Insert Into 病人不良记录
      (ID, 行为类别, 病人id, 发生时间, 加入时间, 加入原因, 加入说明, 附加信息, 登记人)
      Select 病人不良记录_Id.Nextval, v_行为类别, n_病人id, d_发生时间, Nvl(d_加入时间, Sysdate), v_加入原因, v_加入说明, v_附加信息,
             Nvl(v_登记人, v_操作员姓名)
      From Dual;

  End Loop;

  Json_Out := Zljsonout('成功', 1);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Savebadrecord;
/

Create Or Replace Procedure Zl_Patisvr_Savemedccard
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------------------------------------------------------------
  --功能：对病人的医疗卡发放、绑定卡及发卡等相关操作进行医疗卡变动及发卡数据进行保存
  --入参：json格式
  --input
  --   oper_state            N  1  操作状态::0或NULL正常记录;1-产生异常数据;2-只产生变动记录
  --   oper_type             N  1 操作类型:1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定);5-密码调整(只记录);6-挂失(16取消挂失),7-终止时间调整
  --   change_id             N  1  变动ID
  --   pati_id               N  1 病人id
  --   card_type_id          N  1 卡类别ID
  --   card_no_old           C  1 原卡号
  --   card_no               C  1 医疗卡号
  --   card_notes            C  1 变动原因
  --   card_pwd              C  1 密码
  --   iccard_no             C  1 IC卡号
  --   loss_mode             C  1 挂失方式
  --   qrcode                C  1 二维码
  --   card_use_endtime      C  1 终止使用时间:yyyy-mm-dd hh24:mi:ss
  --   operator_time         C  1 操作时间:yyyy-mm-dd hh24:mi:ss
  --   operator_name         C  1 操作员姓名
  --   card_price            N  1 卡费
  --   fee_no                C  1 费用单号

  --出参：json格式
  --Json_Out
  --   code                  N  1  应答码：0-失败；1-成功
  --   message               C  1   应答消息：失败时返回具体的错误信息
  -------------------------------------------------------------------------------------------------
  n_操作状态     Number(2);
  n_操作类型     Number(2);
  n_病人id       病人医疗卡信息.病人id%Type;
  n_卡类别id     病人医疗卡信息.卡类别id%Type;
  v_原卡号       病人医疗卡信息.卡号%Type;
  v_医疗卡号     病人医疗卡信息.卡号%Type;
  v_变动原因     病人医疗卡变动.变动原因%Type;
  v_密码         病人信息.卡验证码%Type;
  v_操作员姓名   病人医疗卡变动.操作员姓名%Type;
  d_操作时间     Date;
  v_Ic卡号       病人信息.Ic卡号%Type := Null;
  v_挂失方式     病人医疗卡变动.挂失方式%Type := Null;
  d_终止使用时间 Date;
  v_二维码       病人医疗卡信息.二维码%Type;
  n_卡费         病人医疗卡变动.卡费%Type;
  v_费用单       病人医疗卡变动.费用单号%Type;
  n_变动id       病人医疗卡变动.Id%Type;

  j_Json   Pljson;
  j_Jsonin Pljson;
Begin
  --解析入参
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_操作类型 := j_Json.Get_Number('oper_type');
  n_病人id   := j_Json.Get_Number('pati_id');
  n_卡类别id := j_Json.Get_Number('card_type_id');
  If Nvl(n_病人id, 0) = 0 Or Nvl(n_卡类别id, 0) = 0 Then
    Json_Out := Zljsonout('失败，未传入病人ID或卡类别id');
    Return;
  End If;

  n_操作状态     := Nvl(j_Json.Get_Number('oper_state'), 0);
  v_原卡号       := j_Json.Get_String('card_no_old');
  v_医疗卡号     := j_Json.Get_String('card_no');
  v_变动原因     := j_Json.Get_String('card_notes');
  v_密码         := j_Json.Get_String('card_pwd');
  v_Ic卡号       := j_Json.Get_String('iccard_no');
  v_挂失方式     := j_Json.Get_String('loss_mode');
  v_二维码       := j_Json.Get_String('qrcode');
  d_终止使用时间 := To_Date(j_Json.Get_String('card_use_endtime'), 'yyyy-mm-dd hh24:mi:ss');
  d_操作时间     := To_Date(j_Json.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');
  v_操作员姓名   := j_Json.Get_String('operator_name');
  n_卡费         := j_Json.Get_Number('card_price');
  v_费用单       := j_Json.Get_String('fee_no');
  n_变动id       := j_Json.Get_Number('change_id');

  Zl_医疗卡变动_Insert_s(n_操作类型, n_病人id, n_卡类别id, v_原卡号, v_医疗卡号, v_变动原因, v_密码, v_操作员姓名, d_操作时间, v_挂失方式, d_终止使用时间, n_卡费, Null,
                    v_费用单, Null, n_变动id, n_操作状态);

  Json_Out := Zljsonout('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Savemedccard;
/
Create Or Replace Procedure Zl_Patisvr_Savepatiphoto
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:保存指定病人相片
  --入参：Json_In:格式
  --   input
  --      pati_id           N 1 病人ID
  --      pati_photo        C 1 编码:base64

  --出参      json
  --output
  --     code               N 1 应答码：0-失败；1-成功
  --     message            C 1 应答消息： 失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Json     Pljson;
  j_Jsonin   Pljson;
  n_病人id   病人照片.病人id%Type;
  b_病人照片 病人照片.照片%Type;
  c_病人照片 Clob;
Begin
  --解析入参
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_病人id := j_Json.Get_Number('pati_id');

  If Nvl(n_病人id, 0) = 0 Then
    Json_Out := '{"output":{"code":0,"message":"未传入病人id，不能保存病人照片!"}}';
    Return;
  End If;

  c_病人照片 := j_Json.Get_Clob('pati_photo');
  b_病人照片 := Zltools.Zlbase64.Decode(c_病人照片);

  Update 病人照片 Set 照片 = b_病人照片 Where 病人id = n_病人id;
  If Sql%RowCount = 0 Then
    Insert Into 病人照片 (病人id, 照片) Values (n_病人id, b_病人照片);
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Savepatiphoto;
/
Create Or Replace Procedure Zl_Patisvr_Updatecardtype
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:修改或增加医疗卡类别
  --入参：Json_In:格式
  --    input
  --      cardtype_id         N  1  ID
  --      cardtype_code       C  1  编码
  --      cardtype_name       C  1  名称
  --      cardtype_stname     C  1  短名
  --      prefix_text         C  1  前缀文本
  --      cardno_len          N  1  卡号长度
  --      default             N  1  缺省标志
  --      fixed               N  1  是否固定:1-是系统固定;0-不是系统固定
  --      strict              N  1  是否严格控制:1-是严格控制;0-不是严格控制
  --      self_make           N  1  是否自制:1-是的;0-不是
  --      exsit_account       N  1  是否存在帐户:1-存在帐户;0-不存在账户
  --      allow_return_cash   N  1  是否退现:1-允许;0-不允许
  --      must_all_return     N  1  是否全退:1-必需全退;0-允许部分退
  --      component           C  1  部件
  --      memo                C  1  备注
  --      spec_item           C  1  特定项目
  --      blnc_mode           C  1  结算方式
  --      cardno_pwdtxt       C  1  卡号密文:卡号从第几位至第几位显示密文,格式为:S-N:S表示从第几位开始,至第几位结束.比如:3-10;表示从3位到10位用密文*表示:12********3323主要是适应不同类别的医疗卡
  --      allow_repeat_use    N  1  是否重复使用:1-允许;0-不允许
  --      enabled             N  1  是否启用:1-已启用;0-未启用
  --      pwd_len             N  1  密码长度
  --      pwd_len_limit       N  1  密码长度限制:0-不作限制;1-固定密码长度;-n表示密码必须输入好多个位密码以上,但不能超过密码长度
  --      pwd_rule            N  1  密码规则:０-数字和字符组成;1-仅为数字组成
  --      allow_vaguefind     N  1  是否模糊查找:1-支持模糊查找;0-不支持
  --      pwd_require         N  1  密码输入限制:0-不限制;1-不输入,提醒;2-不输入禁止;缺省为不限制
  --      default_pwd         N  1  是否缺省密码:1-以身份证后N(以密码长度为准)位作为缺省密码;0-无缺省密码
  --      allow_makecard      N  1  是否制卡:1-是;0-否
  --      allow_sendcard      N  1  是否发卡:1-是;0-否
  --      allow_writcard      N  1  是否写卡:1-是;0-否
  --      insurance_type      N  1  险类
  --      sendcard_nature     N  1  发卡性质:0-不限制;1-同一病人只能发一张卡;2-同一病人允许发多张卡，但需提示;缺省为0
  --      allow_transfer      N  1  是否转帐及代扣:1-支持转帐及代扣;0-不支持
  --      readcard_nature     C  1  读卡性质,医疗卡读卡方式：第一位为:是否刷卡;第二位为是否扫描;第三位是否接触式读卡;第四位是否非接触式读卡。例如刷卡：'1000'
  --      keyboard_mode       N  1  键盘控制方式:：0-禁止使用软键盘;1-使用数字软键盘 ,2-使用字符软键盘
  --      advsend_buildqrcode N  1  是否医嘱发送调用条码生成接口:1-发送调用生成二维码接口;0-不调用
  --      holding_pay         N  1  是否持卡消费:1-是;0-否
  --      cert_cardtype       N  1  是否证件类型的医疗卡:0-不是；1-是
  --      verfycard           N  1  是否退款验卡
  --      sendcard_sign       N  1  发卡控制:0或NULL-发卡时，卡号必须达到卡号长度;1-发卡时，允许卡号小于等于卡号长度,发卡时，小于卡号长度时，不提示操作员;2-发卡时，允许卡号小于等于卡号长度,小于时，提示操作员。
  --      enterkey_enabled    N  1  设备是否启用回车:医疗卡对应的刷卡设备是否启用了回车，如果启用了回车，则卡号长度默认增加一位来屏蔽回车


  --      def_return_cash     N  1  是否缺省退现:允许退现时,默认是否退现
  --      balalone            N  1  是否独立结算:1-独立结算;0-非独立结算
  --      discern_rule        N  1  卡号识别规则:1-全部转换为大写;0-不区分大小写
  --      def_valid_time      C  1  缺省有效时间:NULL时，表示不限制;非空时，格式为:时间+单位(天，月),比如：3天,3月
  --      scanpay             N  1  是否支持扫码付:是否支持扫码付,支持时，会调用“zlReadQRCode部件”
  --出参: Json_Out,格式如下
  --   output
  --      code                N   1   应答码：0-失败；1-成功
  --      message             C   1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------

  j_Json             Pljson;
  j_Jsonin           Pljson;
  n_Id               医疗卡类别.Id%Type;
  v_编码             医疗卡类别.编码%Type;
  v_名称             医疗卡类别.名称%Type;
  v_短名             医疗卡类别.短名%Type;
  v_前缀文本         医疗卡类别.前缀文本%Type;
  n_卡号长度         医疗卡类别.卡号长度%Type;
  n_缺省标志         医疗卡类别.缺省标志%Type;
  n_是否固定         医疗卡类别.是否固定%Type;
  n_是否严格控制     医疗卡类别.是否严格控制%Type;
  n_是否自制         医疗卡类别.是否自制%Type;
  n_是否存在帐户     医疗卡类别.是否存在帐户%Type;
  n_是否全退         医疗卡类别.是否全退%Type;
  v_部件             医疗卡类别.部件%Type;
  v_备注             医疗卡类别.备注%Type;
  v_特定项目         医疗卡类别.特定项目%Type;
  v_结算方式         医疗卡类别.结算方式%Type;
  n_是否启用         医疗卡类别.是否启用%Type;
  v_卡号密文         医疗卡类别.卡号密文%Type;
  n_是否重复使用     医疗卡类别.是否重复使用%Type;
  n_密码长度         医疗卡类别.密码长度%Type;
  n_密码长度限制     医疗卡类别.密码长度限制%Type;
  n_密码规则         医疗卡类别.密码规则%Type;
  n_是否退现         医疗卡类别.是否退现%Type;
  n_操作方式         Integer := 0;
  n_是否模糊查找     医疗卡类别.是否模糊查找%Type := 0;
  n_密码输入限制     医疗卡类别.密码输入限制%Type := 0;
  n_是否缺省密码     医疗卡类别.是否缺省密码%Type := 0;
  n_是否制卡         医疗卡类别.是否制卡%Type := 0;
  n_是否发卡         医疗卡类别.是否发卡%Type := 0;
  n_是否写卡         医疗卡类别.是否写卡%Type := 0;
  n_险类             医疗卡类别.险类%Type := 0;
  n_发卡性质         医疗卡类别.发卡性质%Type := 0;
  n_是否转帐及代扣   医疗卡类别.是否转帐及代扣%Type := 0;
  v_读卡性质         医疗卡类别.读卡性质%Type := '1000';
  n_键盘控制方式     医疗卡类别.键盘控制方式%Type := 0;
  n_是否证件         医疗卡类别.是否证件%Type := 0;
  n_是否持卡消费     医疗卡类别.是否持卡消费%Type := 0;
  n_发送调用接口     医疗卡类别.发送调用接口%Type := 0;
  n_是否退款验卡     医疗卡类别.是否退款验卡%Type := 0;
  n_设备是否启用回车 医疗卡类别.设备是否启用回车%Type := 0;
  n_发卡卡号控制     医疗卡类别.发卡控制%Type := 0;
  n_是否缺省退现     医疗卡类别.是否缺省退现%Type := 0;
  n_是否独立结算     医疗卡类别.是否独立结算%Type := 0;
  d_缺省有效时间     医疗卡类别.缺省有效时间%Type := Null;
  n_卡号识别规则     医疗卡类别.卡号识别规则%Type := 0;
  n_是否支持扫码付   医疗卡类别.是否支持扫码付%Type := 0;

Begin
  --解析入参
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  --      cardtype_id         N  1 ID
  --      cardtype_code       C  1 编码
  --      cardtype_name       C  1 名称
  --      cardtype_stname     C  1 短名
  --      prefix_text         C  1 前缀文本
  n_Id       := j_Json.Get_Number('cardtype_id');
  v_编码     := j_Json.Get_String('cardtype_code');
  v_名称     := j_Json.Get_String('cardtype_name');
  v_短名     := j_Json.Get_String('cardtype_stname');
  v_前缀文本 := j_Json.Get_String('prefix_text');
  --      cardno_len          N  1 卡号长度
  --      default             N  1 缺省标志
  --      fixed               N  1 是否固定:1-是系统固定;0-不是系统固定
  --      strict              N  1 是否严格控制:1-是严格控制;0-不是严格控制
  --      self_make           N  1 是否自制:1-是的;0-不是
  n_卡号长度     := j_Json.Get_Number('cardno_len');
  n_缺省标志     := j_Json.Get_Number('default');
  n_是否固定     := j_Json.Get_Number('fixed');
  n_是否严格控制 := j_Json.Get_Number('strict');
  n_是否自制     := j_Json.Get_Number('self_make');
  --      exsit_account          N  1  是否存在帐户:1-存在帐户;0-不存在账户
  --      allow_return_cash      N  1  是否退现:1-允许;0-不允许
  --      must_all_return        N  1  是否全退:1-必需全退;0-允许部分退
  --      component              C  1  部件
  --      memo                   C  1  备注
  n_是否存在帐户 := j_Json.Get_Number('exsit_account');
  n_是否退现     := j_Json.Get_Number('allow_return_cash');
  n_是否全退     := j_Json.Get_Number('must_all_return');
  v_部件         := j_Json.Get_String('component');
  v_备注         := j_Json.Get_String('memo');
  --      spec_item           C  1  特定项目
  --      blnc_mode           C  1  结算方式
  --      cardno_pwdtxt       C  1  卡号密文:卡号从第几位至第几位显示密文,格式为:S-N:S表示从第几位开始,至第几位结束.比如:3-10;表示从3位到10位用密文*表示:12********3323主要是适应不同类别的医疗卡
  --      allow_repeat_use    N  1  是否重复使用:1-允许;0-不允许
  --      enabled             N  1  是否启用:1-已启用;0-未启用
  v_特定项目     := j_Json.Get_String('spec_item');
  v_结算方式     := j_Json.Get_String('blnc_mode');
  v_卡号密文     := j_Json.Get_String('cardno_pwdtxt');
  n_是否启用     := j_Json.Get_Number('enabled');
  n_是否重复使用 := j_Json.Get_Number('allow_repeat_use');
  --      pwd_len             N  1  密码长度
  --      pwd_len_limit       N  1  密码长度限制:0-不作限制;1-固定密码长度;-n表示密码必须输入好多个位密码以上,但不能超过密码长度
  --      pwd_rule            N  1  密码规则:０-数字和字符组成;1-仅为数字组成
  --      allow_vaguefind     N  1  是否模糊查找:1-支持模糊查找;0-不支持
  --      pwd_require         N  1  密码输入限制:0-不限制;1-不输入,提醒;2-不输入禁止;缺省为不限制

  n_密码长度     := j_Json.Get_Number('pwd_len');
  n_密码长度限制 := j_Json.Get_Number('pwd_len_limit');
  n_密码规则     := j_Json.Get_Number('pwd_rule');
  n_是否模糊查找 := j_Json.Get_Number('allow_vaguefind');
  n_密码输入限制 := j_Json.Get_Number('pwd_require');
  --      default_pwd            N  1  是否缺省密码:1-以身份证后N(以密码长度为准)位作为缺省密码;0-无缺省密码
  --      allow_makecard         N  1  是否制卡:1-是;0-否
  --      allow_sendcard         N  1  是否发卡:1-是;0-否
  --      allow_writcard         N  1  是否写卡:1-是;0-否
  --      insurance_type         N  1  险类
  n_是否缺省密码 := j_Json.Get_Number('default_pwd');
  n_是否制卡     := j_Json.Get_Number('allow_makecard');
  n_是否发卡     := j_Json.Get_Number('allow_sendcard');
  n_是否写卡     := j_Json.Get_Number('allow_writcard');
  n_险类         := j_Json.Get_Number('insurance_type');
  --      sendcard_nature     N  1  发卡性质:0-不限制;1-同一病人只能发一张卡;2-同一病人允许发多张卡，但需提示;缺省为0
  --      allow_transfer      N  1  是否转帐及代扣:1-支持转帐及代扣;0-不支持
  --      readcard_nature     C  1  读卡性质,医疗卡读卡方式：第一位为:是否刷卡;第二位为是否扫描;第三位是否接触式读卡;第四位是否非接触式读卡。例如刷卡：'1000'
  --      keyboard_mode       N  1  键盘控制方式:：0-禁止使用软键盘;1-使用数字软键盘 ,2-使用字符软键盘
  --      advsend_buildqrcode N  1  是否医嘱发送调用条码生成接口:1-发送调用生成二维码接口;0-不调用
  --      holding_pay         N  1  是否持卡消费:1-是;0-否
  --      cert_cardtype       N  1  是否证件类型的医疗卡:0-不是；1-是
  --      verfycard           N  1  是否退款验卡
  n_发卡性质         := j_Json.Get_Number('sendcard_nature');
  n_是否转帐及代扣   := j_Json.Get_Number('allow_transfer');
  v_读卡性质         := j_Json.Get_String('readcard_nature');
  n_键盘控制方式     := j_Json.Get_Number('keyboard_mode');
  n_发送调用接口     := j_Json.Get_Number('advsend_buildqrcode');
  n_是否持卡消费     := j_Json.Get_Number('holding_pay');
  n_是否证件         := j_Json.Get_Number('cert_cardtype');
  n_是否退款验卡     := j_Json.Get_Number('verfycard');
  n_设备是否启用回车 := j_Json.Get_Number('enterkey_enabled');
  n_发卡卡号控制     := j_Json.Get_Number('sendcard_sign');
  --      def_return_cash         N 1 是否缺省退现:允许退现时,默认是否退现
  --      balalone                N 1 是否独立结算:1-独立结算;0-非独立结算
  --      discern_rule            N 1 卡号识别规则:1-全部转换为大写;0-不区分大小写
  --      def_valid_time          C 1 缺省有效时间:NULL时，表示不限制;非空时，格式为:时间+单位(天，月),比如：3天,3月
  --      scanpay                 N 1 是否支持扫码付:是否支持扫码付,支持时，会调用“zlReadQRCode部件”
  n_是否缺省退现   := j_Json.Get_Number('def_return_cash');
  n_是否独立结算   := j_Json.Get_Number('balalone');
  d_缺省有效时间   := j_Json.Get_Number('def_valid_time');
  n_卡号识别规则   := j_Json.Get_String('discern_rule');
  n_是否支持扫码付 := j_Json.Get_Number('scanpay');

  Zl_医疗卡类别_Update(n_Id, v_编码, v_名称, v_短名, v_前缀文本, n_卡号长度, n_缺省标志, n_是否固定, n_是否严格控制, n_是否自制, n_是否存在帐户, n_是否全退, v_部件,
                  v_备注, v_特定项目, Null, v_结算方式, n_是否启用, v_卡号密文, n_是否重复使用, n_密码长度, n_密码长度限制, n_密码规则, n_是否退现, n_操作方式,
                  n_是否模糊查找, n_密码输入限制, n_是否缺省密码, n_是否制卡, n_是否发卡, n_是否写卡, n_险类, n_发卡性质, n_是否转帐及代扣, v_读卡性质, n_键盘控制方式,
                  n_是否证件, n_是否持卡消费, n_发送调用接口, n_是否退款验卡, n_设备是否启用回车, n_发卡卡号控制, n_是否缺省退现, n_是否独立结算, d_缺省有效时间, n_卡号识别规则,
                  n_是否支持扫码付);

  Json_Out := Zljsonout('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Updatecardtype;
/
Create Or Replace Procedure Zl_Patisvr_Updateinpatistate
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：更新住院病人就诊状态
  --入参：Json_In:格式
  --input
  --    pati_list[]              数组
  --      pati_id              N 1   病人id
  --      pati_pageid          N 1   主页id
  --      outpatient_num       C 1   门诊号
  --      inpatient_num        C 1   住院号
  --      in_time              C 1   入院时间
  --      adtd_time            C 1   出院时间
  --      pati_deptid          N 1   当前科室id
  --      wardarea_id          N 1   当前病区id
  --      pati_bed             C 1   当前床号
  --      inp_status           N 1   是否在院，0/1
  --      inp_times            N 1   住院次数
  --      inp_times_increment  N 1   =1时-住院次数自增， =-1时 住院次数自减
  --      insurance_type       N 1  险类
  --      addr_list[]           C     地址信息列表 
  --        oper_fun            N  1  操作功能:1-新增,修改   2-删除 
  --        type                C  1  地址类别 
  --        state               C  1  地址_省 
  --        city                C  1  地址_市 
  --        county              C  1  地址_县 
  --        township            C  1  地址_乡 
  --        other               C  1  地址_其他 
  --        code                C  1  区划代码
  --出参: Json_Out,格式如下
  --    output
  --        code                    N   1   应答吗：0-失败；1-成功
  --        message                 C   1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Json      Pljson;
  j_Temp      Pljson;
  j_Json_List Pljson_List;
  j_Addr_List Pljson_List;

  o_Json      Pljson;
  j_Jsoninput Pljson;

  n_病人id 病人信息.病人id%Type;

  n_主页id     病人信息.主页id%Type;
  n_门诊号     病人信息.门诊号%Type;
  n_住院号     病人信息.住院号%Type;
  n_当前科室id 病人信息.当前科室id%Type;
  n_当前病区id 病人信息.当前病区id%Type;
  n_住院次数   病人信息.住院次数%Type;
  n_在院       病人信息.在院%Type;
  n_险类       病人信息.险类%Type;
  v_当前床号   病人信息.当前床号%Type;

  d_入院时间 病人信息.入院时间%Type;
  d_出院时间 病人信息.出院时间%Type;
  --病人地址信息
  n_地址类别  病人地址信息.地址类别%Type;
  v_地址_省   病人地址信息.省%Type;
  v_地址_市   病人地址信息.市%Type;
  v_地址_县   病人地址信息.县%Type;
  v_地址_乡   病人地址信息.乡镇%Type;
  v_地址_其他 病人地址信息.其他%Type;
  v_区划代码  病人地址信息.区划代码%Type;

  n_操作功能     Number(2);
  n_主页id_b     Number(1);
  n_住院号_b     Number(1);
  n_门诊号_b     Number(1);
  n_当前科室id_b Number(1);
  n_当前病区id_b Number(1);
  n_住院次数_b   Number(1);
  n_入院时间_b   Number(1);
  n_出院时间_b   Number(1);
  n_在院_b       Number(1);
  n_当前床号_b   Number(1);
  n_险类_b       Number(1);

Begin
  --解析入参
  j_Jsoninput := Pljson(Json_In);
  j_Json      := j_Jsoninput.Get_Pljson('input');
  j_Json_List := j_Json.Get_Pljson_List('pati_list');

  If j_Json_List Is Null Then
    Json_Out := Zljsonout('传入值有误,请检查。');
    Return;
  End If;
  For I In 1 .. j_Json_List.Count Loop
    j_Temp := Pljson(j_Json_List.Get(I));
  
    n_主页id_b     := Null;
    n_住院号_b     := Null;
    n_门诊号_b     := Null;
    n_当前科室id_b := Null;
    n_当前病区id_b := Null;
    n_住院次数_b   := Null;
    n_入院时间_b   := Null;
    n_出院时间_b   := Null;
    n_在院_b       := Null;
    n_当前床号_b   := Null;
    n_险类_b       := Null;
    --病人ID
    If j_Temp.Exist('pati_id') Then
      n_病人id := j_Temp.Get_Number('pati_id');
    End If;
  
    --主页ID
    If j_Temp.Exist('pati_pageid') Then
      n_主页id   := j_Temp.Get_Number('pati_pageid');
      n_主页id_b := 1;
    End If;
    --门诊号
    If j_Temp.Exist('outpatient_num') Then
      n_门诊号   := To_Number(j_Temp.Get_String('outpatient_num'));
      n_门诊号_b := 1;
    End If;
  
    --住院号
    If j_Temp.Exist('inpatient_num') Then
      n_住院号   := To_Number(j_Temp.Get_String('inpatient_num'));
      n_住院号_b := 1;
    End If;
    --入院时间
    If j_Temp.Exist('in_time') Then
      d_入院时间   := To_Date(j_Temp.Get_String('in_time'), 'yyyy-mm-dd hh24:mi:ss');
      n_入院时间_b := 1;
    End If;
    --出院时间
    If j_Temp.Exist('adtd_time') Then
      d_出院时间   := To_Date(j_Temp.Get_String('adtd_time'), 'yyyy-mm-dd hh24:mi:ss');
      n_出院时间_b := 1;
    End If;
    --当前科室ID
    If j_Temp.Exist('pati_deptid') Then
      n_当前科室id := j_Temp.Get_Number('pati_deptid');
    
      n_当前科室id_b := 1;
    End If;
    --当前病区ID
    If j_Temp.Exist('wardarea_id') Then
      n_当前病区id   := j_Temp.Get_Number('wardarea_id');
      n_当前病区id_b := 1;
    End If;
    --当前床号
    If j_Temp.Exist('pati_bed') Then
      v_当前床号   := j_Temp.Get_String('pati_bed');
      n_当前床号_b := 1;
    End If;
    --是否在院
    If j_Temp.Exist('inp_status') Then
      n_在院   := j_Temp.Get_Number('inp_status');
      n_在院_b := 1;
    End If;
    --住院次数
    If j_Temp.Exist('inp_times') Then
      n_住院次数   := j_Temp.Get_Number('inp_times');
      n_住院次数_b := 1;
    End If;
    --险类
    If j_Temp.Exist('insurance_type') Then
      n_险类   := j_Temp.Get_Number('insurance_type');
      n_险类_b := 1;
    End If;
  
    Update 病人信息
    Set 主页id = Decode(n_主页id_b, 1, n_主页id, 主页id), 门诊号 = Decode(n_门诊号_b, 1, n_门诊号, 门诊号),
        住院号 = Decode(n_住院号_b, 1, n_住院号, 住院号), 入院时间 = Decode(n_入院时间_b, 1, d_入院时间, 入院时间),
        出院时间 = Decode(n_出院时间_b, 1, d_出院时间, 出院时间), 当前科室id = Decode(n_当前科室id_b, 1, n_当前科室id, 当前科室id),
        当前病区id = Decode(n_当前病区id_b, 1, n_当前病区id, 当前病区id), 当前床号 = Decode(n_当前床号_b, 1, v_当前床号, 当前床号),
        在院 = Decode(n_在院_b, 1, n_在院, 在院), 住院次数 = Decode(n_住院次数_b, 1, n_住院次数, 住院次数), 险类 = Decode(n_险类_b, 1, n_险类, 险类)
    Where 病人id = n_病人id;
  
    --住院次数自增
    If j_Temp.Exist('inp_times_increment') Then
      n_住院次数 := j_Temp.Get_Number('inp_times_increment');
      If Nvl(n_住院次数, 0) = 1 Then
        Update 病人信息 Set 住院次数 = Nvl(住院次数, 0) + 1 Where 病人id = n_病人id;
      Elsif Nvl(n_住院次数, 0) = -1 Then
        Update 病人信息 Set 住院次数 = Decode(Sign(住院次数 - 1), 1, 住院次数 - 1, Null) Where 病人id = n_病人id;
      End If;
    End If;
  
    --更新地址信息 
    --    addr_list[]           C     地址信息列表 
    --      oper_fun            N  1  操作功能:1-新增,修改   2-删除 
    --      type                C  1  地址类别 
    --      state               C  1  地址_省 
    --      city                C  1  地址_市 
    --      county              C  1  地址_县 
    --      township            C  1  地址_乡 
    --      other               C  1  地址_其他 
    --      code                C  1  区划代码 
    If j_Temp.Exist('addr_list') Then
      j_Addr_List := Pljson_List();
      j_Addr_List := j_Temp.Get_Pljson_List('addr_list');
      If j_Addr_List Is Not Null Then
        For I In 1 .. j_Addr_List.Count Loop
          o_Json      := Pljson();
          o_Json      := Pljson(j_Addr_List.Get(I));
          n_操作功能  := o_Json.Get_Number('oper_fun');
          n_地址类别  := o_Json.Get_Number('type');
          v_地址_省   := o_Json.Get_String('state');
          v_地址_市   := o_Json.Get_String('city');
          v_地址_县   := o_Json.Get_String('county');
          v_地址_乡   := o_Json.Get_String('township');
          v_地址_其他 := o_Json.Get_String('other');
          v_区划代码  := o_Json.Get_String('code');
          Zl_病人地址信息_Update_s(n_操作功能, n_病人id, n_主页id, n_地址类别, v_地址_省, v_地址_市, v_地址_县, v_地址_乡, v_地址_其他, v_区划代码);
        End Loop;
      End If;
    End If;
  End Loop;
  Json_Out := Zljsonout('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Updateinpatistate;
/
Create Or Replace Procedure Zl_Patisvr_Updateoutpatistate
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：更新门诊病人就诊状态
  --      可以判断结点是存在，如果不存在则不更新，或者更新为原值，目前暂时未用到后续可以基于这个进行扩展
  --入参：Json_In:格式
  --input
  --    pati_id            N 1 病人id
  --    pati_age           C 0 年龄
  --    phone_number       C 0 病人手机号
  --    fee_category       C 0 费别
  --    visit_room         C 0 更新的就诊诊室
  --    visit_status       N 0 更新的就诊状态
  --    visit_time         C 0 更新的就诊时间
  --    outpatient_num     C 0 门诊号

  --出参: Json_Out,格式如下
  --    output
  --        code                    N   1   应答吗：0-失败；1-成功
  --        message                 C   1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  o_Json   Pljson;
  j_Jsonin Pljson;

  n_病人id   病人信息.病人id%Type;
  v_年龄     病人信息.年龄%Type;
  n_就诊状态 病人信息.就诊状态%Type;
  v_就诊诊室 病人信息.就诊诊室%Type;
  d_就诊时间 病人信息.就诊时间%Type;
  v_费别     病人信息.费别%Type;
  v_手机号   病人信息.手机号%Type;
  n_门诊号   病人信息.门诊号%Type;
  n_费别性质 费别.属性%Type;

  n_费别_b     Number(1);
  n_就诊状态_b Number(1);
  n_就诊诊室_b Number(1);
  n_就诊时间_b Number(1);
  n_手机号_b   Number(1);
  n_门诊号_b   Number(1);
  n_年龄_b     Number(1);
Begin
  --解析入参
  j_Jsonin := Pljson(Json_In);
  o_Json   := j_Jsonin.Get_Pljson('input');

  n_病人id := o_Json.Get_Number('pati_id');

  If o_Json.Exist('phone_number') Then
    v_手机号   := o_Json.Get_String('phone_number');
    n_手机号_b := 1;
  End If;

  If o_Json.Exist('visit_status') Then
    n_就诊状态   := o_Json.Get_Number('visit_status');
    n_就诊状态_b := 1;
  End If;
  If o_Json.Exist('visit_room') Then
    v_就诊诊室   := o_Json.Get_String('visit_room');
    n_就诊诊室_b := 1;
  End If;
  If o_Json.Exist('visit_time') Then
    d_就诊时间   := To_Date(o_Json.Get_String('visit_time'), 'yyyy-mm-dd hh24:mi:ss');
    n_就诊时间_b := 1;
  End If;

  If o_Json.Exist('fee_category') Then
    v_费别   := o_Json.Get_String('fee_category');
    n_费别_b := 1;
  End If;

  If o_Json.Exist('outpatient_num') Then
    n_门诊号   := o_Json.Get_String('outpatient_num');
    n_门诊号_b := 1;
  End If;

  If o_Json.Exist('pati_age') Then
    v_年龄   := o_Json.Get_String('pati_age');
    n_年龄_b := 1;
  End If;

  If v_费别 Is Not Null Then
    Select Max(属性) Into n_费别性质 From 费别 Where 名称 = v_费别; --2-动态费别不更新
    If n_费别性质 = 2 Then
      n_费别_b := 0;
    End If;
  End If;

  Update 病人信息
  Set 费别 = Decode(n_费别_b, 1, v_费别, 费别), 手机号 = Decode(n_手机号_b, 1, v_手机号, 手机号), 就诊状态 = Decode(n_就诊状态_b, 1, n_就诊状态, 就诊状态),
      就诊诊室 = Decode(n_就诊诊室_b, 1, v_就诊诊室, 就诊诊室), 就诊时间 = Decode(n_就诊时间_b, 1, d_就诊时间, 就诊时间),
      门诊号 = Decode(n_门诊号_b, 1, n_门诊号, 门诊号), 年龄 = Decode(n_年龄_b, 1, v_年龄, 年龄)
  Where 病人id = n_病人id;

  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Updateoutpatistate;
/
Create Or Replace Procedure Zl_Patisvr_Updatepatiarchives
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  --------------------------------------------------------------------------- 
  --功能：修改病人档案信息 
  --入参：Json_In:格式 
  --input 
  --    oper_fun              N  1   0-要更新病人信息表 1-不更新病人信息表 
  --    is_realname_check     N  1  是否实名检查:1-实名检查;0-不检查 
  --    pati_id               N  1  病人id:更新条件 
  --    pati_pageid           N  1  主页ID 
  --    pati_name_old         N     病人姓名(未修改前的姓名):比如：新病人 
  --    pati_name             N  1  病人姓名 
  --    pati_sex              C  1  性别 
  --    pati_age              C  1  年龄 
  --    pati_type             C  1  病人类型(普通，医保，留观) 
  --    pati_birthdate        C  1  出生日期:yyyy-mm-dd hh24:mi:ss 
  --    phone_number          C  1  手机号 
  --    insurance_num         C  1  医保号 
  --    pati_idcard           C  1  身份证号 
  --    outpatient_num        C  1  门诊号 
  --    fee_category          C  1  费别 
  --    mdlpay_mode_name      C  1  医疗付款方式名称 
  --    country_name          C  1  国籍 
  --    native_place          C  1  籍贯 
  --    nation_name           C  1  民族 
  --    mari_status           C  1  婚姻状况 
  --    ocpt_name             C  1  职业 
  --    edu_name              C  1  学历 
  --    pati_identity         C  1  身份 
  --    insurance_type        N  1  险类 
  --    emp_name              C  1  工作单位 
  --    emp_postcode          C  1  单位邮编 
  --    emp_phno              C  1  单位电话 
  --    emp_bank_name         C   1   单位开户行 
  --    emp_bank_accnum       C   1   单位帐号 
  --    ctt_unit_id           N  1  合同单位id 
  --    pat_home_addr         C  1  家庭地址 
  --    pat_home_phno         C  1  家庭电话 
  --    pat_home_postcode     C  1  家庭地址邮编 
  --    region                C  1  区域 
  --    pat_baddr             C  1  出生地点 
  --    pat_hous_addr         C  1  户口地址 
  --    pat_hous_postcode     C  1  户口地址邮编 
  --    pat_grdn_name         C  1  监护人 
  --    vcard_no              C  1  就诊卡号 
  --    vcard_pwd             C  1  卡验证码 
  --    iccard_no             C  1  Ic卡号 
  --    create_time           C  1  登记时间:yyyy-mm-dd hh24:mi:ss 
  --    operator_name         C  1  操作员姓名 
  --    cardno_clear          N     清除就诊卡信息 
  --    pati_wardarea_id      N     当前病区id 
  --    pati_bed              C     当前床号 
  --    idcard_sign           N     身份证签约 
  --    idcard_sign_pwd       C     签约密码 
  --    cert_no_other         C  1 其他证件 
  --    qq                    C      qq 
  --    email                 C      email 
  --    emp_addr              C     单位地址 
  --    contacts              C     更新联系人信息节点 
  --      name                C  1  联系人姓名 
  --      idcard              C  1  联系人身份证号 
  --      phone               C  1  联系人电话 
  --      relation            C  1  联系人关系 
  --      address             C     联系人地址 
  --    community_info        C     社区信息节点 
  --      num                 N  1  社区序号 
  --      code                C  1  社区号码 
  --      oper_type           N  1  社区操作类型 
  --    visit_info            C     就诊信息节点 
  --      status              N     更新的就诊状态 
  --      room                C     更新的就诊诊室 
  --      time                C     就诊时间:yyyy-mm-dd hh24:mi:ss 
  --    addr_list[]           C     地址信息列表 
  --      oper_fun            N  1  操作功能:1-新增,修改   2-删除 
  --      type                C  1  地址类别 
  --      state               C  1  地址_省 
  --      city                C  1  地址_市 
  --      county              C  1  地址_县 
  --      township            C  1  地址_乡 
  --      other               C  1  地址_其他 
  --      code                C  1  区划代码 
  --      visit_or_in         N  1  是否存在就诊或者住院信息 
  --    ext_list[]            C     病人信息从项列表 
  --      info_name           C  1  信息名 
  --      upd_info_value      N  1  修改的信息值 
  --      visit_id            N     就诊id_In 
  --    cert_list[]                 证件列表(主要是当成绑卡处理) 
  --      cert_name           C  1  证件名称 
  --      cert_no             C  1  证号号码 
  --    oper_allergic_drugs N  1  过敏药物  0-删除后更新 1-单个记录更新 
  --    allergic_drugs_list[]       病人过敏药物列表:有数据时，是先删除过敏药物插入的方式 
  --      oper_type            N  1  0-更新 1-删除 
  --      pat_algc_cadn_id    N  1  过敏药品ID 
  --      pat_algc_cadn       C  1  过敏药物名称 
  --      allergy_info        C  1  过每药物反应 
  --      allergic_drugs      C  1  过敏药品ID:过敏药物名称拼串 
  --    immune_list[]         C     病人免疫列表 
  --      vaccinate_time      C  1  接种时间:yyyy-mm-dd hh24:mi:ss 
  --      vaccinate_name      C  1  接种名称 
  --    card_property_list[]  C     医疗卡属性列表 
  --      cardtype_id         N  1  医疗卡类别ID 
  --      card_no             C  1  卡号 
  --      info_name           C  1  信息名 
  --      info_value          N  1  信息值 
  --      item_list[]         更新病人信息某一个字段的值 
  --      item_name           C  1  字段名 
  --      item_value          C  1   字段值 

  --出参: Json_Out,格式如下 
  --  output 
  --    code                  N 1   应答码：0-失败；1-成功 
  --    message               C 1   应答消息：失败时返回具体的错误信息 
  --------------------------------------------------------------------------- 

  j_Json         Pljson;
  j_Jsonin       Pljson;
  j_Jsonlist     Pljson_List := Pljson_List();
  o_Json         Pljson;
  o_Json1        Pljson;
  n_实名检查     Number(1);
  n_病人id       病人信息.病人id%Type;
  n_主页id       病人信息.主页id%Type;
  v_姓名_Old     病人信息.姓名%Type;
  v_姓名         病人信息.姓名%Type;
  v_身份证号     病人信息.身份证号%Type;
  v_年龄         病人信息.年龄%Type;
  v_性别         病人信息.性别%Type;
  d_出生日期     病人信息.出生日期%Type;
  v_手机号       病人信息.手机号%Type;
  v_家庭电话     病人信息.家庭电话%Type;
  n_门诊号       病人信息.门诊号%Type;
  v_费别         病人信息.费别%Type;
  v_医疗付款方式 病人信息.医疗付款方式%Type;
  v_国籍         病人信息.国籍%Type;
  v_籍贯         病人信息.籍贯%Type;
  v_民族         病人信息.民族%Type;
  v_婚姻         病人信息.婚姻状况%Type;
  v_职业         病人信息.职业%Type;
  v_学历         病人信息.学历%Type;
  v_工作单位     病人信息.工作单位%Type;
  n_合同单位id   病人信息.合同单位id%Type;
  v_单位电话     病人信息.单位电话%Type;
  v_单位邮编     病人信息.单位邮编%Type;
  v_家庭地址     病人信息.家庭地址%Type;
  v_家庭地址邮编 病人信息.家庭地址邮编%Type;
  v_户口地址     病人信息.户口地址%Type;
  v_户口地址邮编 病人信息.户口地址邮编%Type;
  d_登记时间     病人信息.登记时间%Type;
  v_医保号       病人信息.医保号%Type;
  v_区域         病人信息.区域%Type;
  v_监护人       病人信息.监护人%Type;
  v_出生地点     病人信息.出生地点%Type;
  v_身份         病人信息.身份%Type;
  v_操作员姓名   病人医疗卡信息.发卡人%Type;
  n_社区id       病人社区信息.社区%Type;
  v_社区号码     病人社区信息.社区号%Type;
  n_社区类型     病人社区信息.就诊类型%Type;
  v_卡名称       医疗卡类别.名称%Type;
  n_卡类别id     医疗卡类别.Id%Type;
  v_卡号         病人医疗卡信息.卡号%Type;
  n_卡号长度     医疗卡类别.卡号长度%Type;
  v_编码         医疗卡类别.编码%Type;
  v_卡密码       病人医疗卡信息.密码%Type;
  v_变动原因     病人医疗卡变动.变动原因%Type;
  d_终止使用时间 病人医疗卡信息.终止使用时间%Type;
  n_过敏药品id   病人过敏药物.过敏药物id%Type;
  v_过敏药物名称 病人过敏药物.过敏药物%Type;
  v_过每药物反应 病人过敏药物.过敏反应%Type;
  d_接种时间     病人免疫记录.接种时间%Type;
  v_接种名称     病人免疫记录.接种名称%Type;
  v_就诊卡号     病人信息.就诊卡号%Type;
  v_卡验证码     病人信息.卡验证码%Type;
  v_Ic卡号       病人信息.Ic卡号%Type;
  v_当前床号     病人信息.当前床号%Type;
  n_当前病区id   病人信息.当前病区id %Type;
  n_就诊状态     病人信息.就诊状态%Type;
  v_就诊诊室     病人信息.就诊诊室%Type;
  d_就诊时间     病人信息.就诊时间%Type;
  v_其他证件     病人信息.其他证件%Type;
  v_单位帐号     病人信息.单位帐号%Type;
  v_单位开户行   病人信息.单位开户行%Type;
  v_病人类型     病人信息.病人类型%Type;
  v_Qq           病人信息.Qq%Type;
  v_Email        病人信息.Email%Type;
  n_是否就诊     Number;
  v_单位地址     病人信息.单位地址%Type;

  n_清除就诊卡信息 Number(1);
  n_Count          Number(10);
  --联系人 
  v_联系人姓名     病人信息.联系人姓名%Type;
  v_联系人关系     病人信息.联系人关系%Type;
  v_联系人电话     病人信息.联系人电话%Type;
  v_联系人身份证号 病人信息.联系人身份证号%Type;
  v_联系人地址     病人信息.联系人地址%Type;
  --病人信息从表 
  v_信息名 病人信息从表.信息名%Type;
  v_信息值 病人信息从表.信息值%Type;
  --病人地址信息 
  n_操作功能       Number(3);
  n_地址类型       病人地址信息.地址类别%Type;
  v_省             病人地址信息.省%Type;
  v_市             病人地址信息.市%Type;
  v_县             病人地址信息.县%Type;
  v_乡镇           病人地址信息.乡镇%Type;
  v_其他           病人地址信息.其他%Type;
  v_区划代码       病人地址信息.区划代码%Type;
  n_险类           病人信息.险类%Type;
  v_Msg            Varchar2(4000);
  v_Strtmpbefor    Varchar2(4000);
  v_字段名         Varchar2(1000);
  v_字段值         Varchar2(3682);
  v_Sql            Varchar2(3682);
  n_费别性质       Number(1);
  n_姓名_b         Number(1); --加后缀_b:为1时表示对应字段的json节点存在，为0时表示对应字段的json节点不存在 
  n_性别_b         Number(1);
  n_年龄_b         Number(1);
  n_出生日期_b     Number(1);
  n_门诊号_b       Number(1);
  n_费别_b         Number(1);
  n_医保号_b       Number(1);
  n_险类_b         Number(1);
  n_医疗付款方式_b Number(1);
  n_就诊_b         Number(1);
  n_手机号_b       Number(1);
  n_当前床号_b     Number(1);
  n_当前病区_b     Number(1);
  n_身份证号_b     Number(1);
  n_国籍_b         Number(1);
  n_籍贯_b         Number(1);
  n_婚姻状况_b     Number(1);
  n_出生地点_b     Number(1);
  n_学历_b         Number(1);
  n_职业_b         Number(1);
  n_区域_b         Number(1);
  n_工作单位_b     Number(1);
  n_合同单位id_b   Number(1);
  n_单位电话_b     Number(1);
  n_单位邮编_b     Number(1);
  n_家庭地址_b     Number(1);
  n_家庭电话_b     Number(1);
  n_家庭地址邮编_b Number(1);
  n_户口地址_b     Number(1);
  n_户口地址邮编_b Number(1);
  n_联系人_b       Number(1);
  n_民族_b         Number(1);
  n_身份_b         Number(1);
  n_监护人_b       Number(1);
  n_就诊卡号_b     Number(1);
  n_卡验证码_b     Number(1);
  n_Ic卡号_b       Number(1);
  n_其他证件_b     Number(1);
  n_病人类型_b     Number(1);
  n_单位开户行_b   Number(1);
  n_单位帐号_b     Number(1);
  n_Qq_b           Number(1);
  n_Email_b        Number(1);
  n_就诊id         病人信息从表.就诊id%Type;
  n_单位地址_b     Number(1);

  n_最长值       Number(10);
  n_最大值       Number(10);
  n_功能         Number(1);
  n_过敏药物更新 Number(1);
  n_方式         Number(1);
  c_过敏药物     Clob;
  l_过敏药物     t_Strlist := t_Strlist();

Begin
  --解析入参 
  j_Jsonin := Pljson(Json_In);
  If j_Jsonin Is Null Then
    Json_Out := Zljsonout('未传入任何信息，请检查');
    Return;
  Else
    o_Json := j_Jsonin.Get_Pljson('input');
  End If;
  --    is_realname_check     N  1  是否实名检查:1-实名检查;0-不检查 
  --    pati_id               N  1  病人id:更新条件 
  --    pati_pageid           N  1  主页ID 
  --    pati_name_old         N     病人姓名(未修改前的姓名):比如：新病人 
  --    pati_name             N  1  病人姓名 
  --    pati_sex              C  1  性别 
  --    pati_age              C  1  年龄 
  --    pati_birthdate        C  1  出生日期:yyyy-mm-dd hh24:mi:ss 
  n_病人id       := o_Json.Get_Number('pati_id');
  n_主页id       := o_Json.Get_Number('pati_pageid');
  n_功能         := o_Json.Get_Number('oper_fun');
  n_过敏药物更新 := o_Json.Get_Number('oper_allergic_drugs');
  If Nvl(n_功能, 0) = 0 Then
    n_实名检查 := o_Json.Get_Number('is_realname_check');
  
    If Nvl(n_病人id, 0) = 0 Then
      Json_Out := Zljsonout('未传入病人id，不能保存');
      Return;
    End If;
    v_姓名_Old := o_Json.Get_String('pati_name_old');
  
    If o_Json.Exist('pati_name') Then
      v_姓名   := o_Json.Get_String('pati_name');
      n_姓名_b := 1;
    End If;
  
    If o_Json.Exist('pati_sex') Then
      v_性别   := o_Json.Get_String('pati_sex');
      n_性别_b := 1;
    End If;
  
    If o_Json.Exist('pati_age') Then
      v_年龄   := o_Json.Get_String('pati_age');
      n_年龄_b := 1;
    End If;
  
    If o_Json.Exist('pati_birthdate') Then
      d_出生日期   := To_Date(o_Json.Get_String('pati_birthdate'), 'yyyy-mm-dd hh24:mi:ss');
      n_出生日期_b := 1;
    End If;
  
    --    phone_number          C  1  手机号 
    --    insurance_num         C  1  医保号 
    --    pati_idcard           C  1  身份证号 
    --    outpatient_num        C  1  门诊号 
    --    fee_category          C  1  费别 
    --    mdlpay_mode_name      C  1  医疗付款方式名称 
    If o_Json.Exist('phone_number') Then
      v_手机号   := o_Json.Get_String('phone_number');
      n_手机号_b := 1;
    End If;
  
    If o_Json.Exist('insurance_num') Then
      v_医保号   := o_Json.Get_String('insurance_num');
      n_医保号_b := 1;
    End If;
  
    If o_Json.Exist('pati_idcard') Then
      v_身份证号   := o_Json.Get_String('pati_idcard');
      n_身份证号_b := 1;
    End If;
  
    If o_Json.Exist('outpatient_num') Then
      n_门诊号   := To_Number(o_Json.Get_String('outpatient_num'));
      n_门诊号_b := 1;
    End If;
  
    If o_Json.Exist('fee_category') Then
      v_费别   := o_Json.Get_String('fee_category');
      n_费别_b := 1;
    End If;
  
    If o_Json.Exist('mdlpay_mode_name') Then
      v_医疗付款方式   := o_Json.Get_String('mdlpay_mode_name');
      n_医疗付款方式_b := 1;
    End If;
  
    --    country_name          C  1  国籍 
    --    native_place          C  1  籍贯 
    --    nation_name           C  1  民族 
    --    mari_status           C  1  婚姻状况 
    --    ocpt_name             C  1  职业 
    --    edu_name              C  1  学历 
    --    pati_identity         C  1  身份 
  
    If o_Json.Exist('country_name') Then
      v_国籍   := o_Json.Get_String('country_name');
      n_国籍_b := 1;
    End If;
  
    If o_Json.Exist('native_place') Then
      v_籍贯   := o_Json.Get_String('native_place');
      n_籍贯_b := 1;
    End If;
  
    If o_Json.Exist('nation_name') Then
      v_民族   := o_Json.Get_String('nation_name');
      n_民族_b := 1;
    End If;
  
    If o_Json.Exist('mari_status') Then
      v_婚姻       := o_Json.Get_String('mari_status');
      n_婚姻状况_b := 1;
    End If;
  
    If o_Json.Exist('ocpt_name') Then
      v_职业   := o_Json.Get_String('ocpt_name');
      n_职业_b := 1;
    End If;
  
    If o_Json.Exist('edu_name') Then
      v_学历   := o_Json.Get_String('edu_name');
      n_学历_b := 1;
    End If;
  
    If o_Json.Exist('pati_identity') Then
      v_身份   := o_Json.Get_String('pati_identity');
      n_身份_b := 1;
    End If;
  
    --    insurance_type        N  1  险类 
    --    emp_name              C  1  工作单位 
    --    emp_postcode          C  1  单位邮编 
    --    emp_phno              C  1  单位电话 
    --    ctt_unit_id           N  1  合同单位id 
    --    pat_home_addr         C  1  家庭地址 
    --    pat_home_phno         C  1  家庭电话 
    --    pat_home_postcode     C  1  家庭地址邮编 
    If o_Json.Exist('insurance_type') Then
      n_险类   := o_Json.Get_Number('insurance_type');
      n_险类_b := 1;
    End If;
  
    If o_Json.Exist('emp_name') Then
      v_工作单位   := o_Json.Get_String('emp_name');
      n_工作单位_b := 1;
    End If;
  
    If o_Json.Exist('emp_postcode') Then
      v_单位邮编   := o_Json.Get_String('emp_postcode');
      n_单位邮编_b := 1;
    End If;
  
    If o_Json.Exist('emp_phno') Then
      v_单位电话   := o_Json.Get_String('emp_phno');
      n_单位电话_b := 1;
    End If;
  
    If o_Json.Exist('ctt_unit_id') Then
      n_合同单位id := o_Json.Get_Number('ctt_unit_id');
      If n_合同单位id > 0 Then
        n_合同单位id_b := 1;
      End If;
    End If;
  
    If o_Json.Exist('pat_home_addr') Then
      v_家庭地址   := o_Json.Get_String('pat_home_addr');
      n_家庭地址_b := 1;
    End If;
  
    If o_Json.Exist('pat_home_phno') Then
      v_家庭电话   := o_Json.Get_String('pat_home_phno');
      n_家庭电话_b := 1;
    End If;
  
    If o_Json.Exist('pat_home_postcode') Then
      v_家庭地址邮编   := o_Json.Get_String('pat_home_postcode');
      n_家庭地址邮编_b := 1;
    End If;
  
    --    region                C  1  区域 
    --    pat_baddr             C  1  出生地点 
    --    pat_hous_addr         C  1  户口地址 
    --    pat_hous_postcode     C  1  户口地址邮编 
    --    pat_grdn_name         C  1  监护人 
    If o_Json.Exist('region') Then
      v_区域   := o_Json.Get_String('region');
      n_区域_b := 1;
    End If;
  
    If o_Json.Exist('pat_baddr') Then
      v_出生地点   := o_Json.Get_String('pat_baddr');
      n_出生地点_b := 1;
    End If;
  
    If o_Json.Exist('pat_hous_addr') Then
      v_户口地址   := o_Json.Get_String('pat_hous_addr');
      n_户口地址_b := 1;
    End If;
  
    If o_Json.Exist('pat_hous_postcode') Then
      v_户口地址邮编   := o_Json.Get_String('pat_hous_postcode');
      n_户口地址邮编_b := 1;
    End If;
  
    If o_Json.Exist('pat_grdn_name') Then
      v_监护人   := o_Json.Get_String('pat_grdn_name');
      n_监护人_b := 1;
    End If;
  
    --    vcard_no              C  1  就诊卡号 
    --    vcard_pwd             C  1  卡验证码 
    --    iccard_no             C  1  Ic卡号 
    --    create_time           C  1  登记时间:yyyy-mm-dd hh24:mi:ss 
    --    operator_name         C  1  操作员姓名 
    --    pati_wardarea_id      N     当前病区id 
    --    pati_bed              C     当前床号 
  
    If o_Json.Exist('vcard_no') Then
      v_就诊卡号   := o_Json.Get_String('vcard_no');
      n_就诊卡号_b := 1;
    End If;
  
    If o_Json.Exist('vcard_pwd') Then
      v_卡验证码   := o_Json.Get_String('vcard_pwd');
      n_卡验证码_b := 1;
    End If;
  
    If o_Json.Exist('iccard_no') Then
      v_Ic卡号   := o_Json.Get_String('iccard_no');
      n_Ic卡号_b := 1;
    End If;
  
    If o_Json.Exist('create_time') Then
      d_登记时间 := To_Date(o_Json.Get_String('create_time'), 'yyyy-mm-dd hh24:mi:ss');
    End If;
    If d_登记时间 Is Null Then
      d_登记时间 := Sysdate;
    End If;
  
    If o_Json.Exist('operator_name') Then
      v_操作员姓名 := o_Json.Get_String('operator_name');
    End If;
  
    If o_Json.Exist('pati_wardarea_id') Then
      n_当前病区id := o_Json.Get_Number('pati_wardarea_id');
      n_当前病区_b := 1;
    End If;
  
    If o_Json.Exist('pati_bed') Then
      v_当前床号   := o_Json.Get_String('pati_bed');
      n_当前床号_b := 1;
    End If;
  
    --    cert_no_other       C   1   其他证件 
    --    pati_type           C   1   病人类型(普通，医保，留观) 
    --    emp_bank_name       C   1   单位开户行 
    --    emp_bank_accnum     C   1   单位帐号 
  
    If o_Json.Exist('cert_no_other') Then
      v_其他证件   := o_Json.Get_String('cert_no_other');
      n_其他证件_b := 1;
    End If;
  
    If o_Json.Exist('pati_type') Then
      v_病人类型   := o_Json.Get_String('pati_type');
      n_病人类型_b := 1;
    End If;
  
    If o_Json.Exist('qq') Then
      v_Qq   := o_Json.Get_String('qq');
      n_Qq_b := 1;
    End If;
  
    If o_Json.Exist('email') Then
      v_Email   := o_Json.Get_String('email');
      n_Email_b := 1;
    End If;
  
    If o_Json.Exist('emp_bank_name') Then
      v_单位开户行   := o_Json.Get_String('emp_bank_name');
      n_单位开户行_b := 1;
    End If;
    If o_Json.Exist('emp_bank_accnum') Then
      v_单位帐号   := o_Json.Get_String('emp_bank_accnum');
      n_单位帐号_b := 1;
    End If;
  
    If o_Json.Exist('emp_addr') Then
      v_单位地址   := o_Json.Get_String('emp_addr');
      n_单位地址_b := 1;
    End If;
  
    --    contacts              C     更新联系人信息节点 
    --      name                C  1  联系人姓名 
    --      idcard              C  1  联系人身份证号 
    --      phone               C  1  联系人电话 
    --      relation            C  1  联系人关系 
    --      address             C     联系人地址 
    o_Json1 := Pljson();
    o_Json1 := o_Json.Get_Pljson('contacts');
    If Not o_Json1 Is Null Then
      v_联系人姓名     := o_Json1.Get_String('name');
      v_联系人关系     := o_Json1.Get_String('relation');
      v_联系人身份证号 := o_Json1.Get_String('idcard');
      v_联系人电话     := o_Json1.Get_String('phone');
      v_联系人地址     := o_Json1.Get_String('address');
      n_联系人_b       := 1;
    End If;
  
    --        visit_info          修正就诊登记信息 
    --          status        N 1 更新的就诊状态 
    --          room          C 1 更新的就诊诊室 
    --          time          C 1 更新的就诊时间 
    o_Json1 := Pljson();
    o_Json1 := o_Json.Get_Pljson('visit_info');
    If o_Json1 Is Not Null Then
      n_就诊状态 := o_Json1.Get_Number('status');
      v_就诊诊室 := o_Json1.Get_String('room');
      d_就诊时间 := To_Date(o_Json1.Get_String('time'), 'yyyy-mm-dd hh24:mi:ss');
      n_就诊_b   := 1;
    End If;
  
    If Nvl(n_实名检查, 0) = 1 Then
      Select Zl_Fun_Checkidentify(1, n_病人id, v_Strtmpbefor) Into v_Strtmpbefor From Dual;
    End If;
    If v_费别 Is Not Null Then
      Select Max(属性) Into n_费别性质 From 费别 Where 名称 = v_费别; --2-动态费别不更新 
      If n_费别性质 = 2 Then
        n_费别_b := 0;
      End If;
    End If;
  
    Update 病人信息
    Set 姓名 = Decode(n_姓名_b, 1, v_姓名, 姓名), 性别 = Decode(n_性别_b, 1, v_性别, 性别), 年龄 = Decode(n_年龄_b, 1, v_年龄, 年龄),
        出生日期 = Decode(n_出生日期_b, 1, d_出生日期, 出生日期), 门诊号 = Decode(n_门诊号_b, 1, n_门诊号, 门诊号), 费别 = Decode(n_费别_b, 1, v_费别, 费别),
        医保号 = Decode(n_医保号_b, 1, v_医保号, 医保号), 险类 = Decode(n_险类_b, 1, n_险类, 险类),
        医疗付款方式 = Decode(n_医疗付款方式_b, 1, v_医疗付款方式, 医疗付款方式), 手机号 = Decode(n_手机号_b, 1, v_手机号, 手机号),
        身份证号 = Decode(n_身份证号_b, 1, v_身份证号, 身份证号), 出生地点 = Decode(n_出生地点_b, 1, v_出生地点, 出生地点),
        婚姻状况 = Decode(n_婚姻状况_b, 1, v_婚姻, 婚姻状况), 国籍 = Decode(n_国籍_b, 1, v_国籍, 国籍), 学历 = Decode(n_学历_b, 1, v_学历, 学历),
        职业 = Decode(n_职业_b, 1, v_职业, 职业), 籍贯 = Decode(n_籍贯_b, 1, v_籍贯, 籍贯), 区域 = Decode(n_区域_b, 1, v_区域, 区域),
        工作单位 = Decode(n_工作单位_b, 1, v_工作单位, 工作单位), 合同单位id = Decode(n_合同单位id_b, 1, n_合同单位id, 合同单位id),
        单位电话 = Decode(n_单位电话_b, 1, v_单位电话, 单位电话), 单位邮编 = Decode(n_单位邮编_b, 1, v_单位邮编, 单位邮编),
        家庭地址 = Decode(n_家庭地址_b, 1, v_家庭地址, 家庭地址), 家庭电话 = Decode(n_家庭电话_b, 1, v_家庭电话, 家庭电话),
        家庭地址邮编 = Decode(n_家庭地址邮编_b, 1, v_家庭地址邮编, 家庭地址邮编), 户口地址 = Decode(n_户口地址_b, 1, v_户口地址, 户口地址),
        户口地址邮编 = Decode(n_户口地址邮编_b, 1, v_户口地址邮编, 户口地址邮编), 联系人姓名 = Decode(n_联系人_b, 1, v_联系人姓名, 联系人姓名),
        联系人关系 = Decode(n_联系人_b, 1, v_联系人关系, 联系人关系), 联系人身份证号 = Decode(n_联系人_b, 1, v_联系人身份证号, 联系人身份证号),
        联系人电话 = Decode(n_联系人_b, 1, v_联系人电话, 联系人电话), 联系人地址 = Decode(n_联系人_b, 1, v_联系人地址, 联系人地址),
        当前床号 = Decode(n_当前床号_b, 1, v_当前床号, 当前床号), 当前病区id = Decode(n_当前病区_b, 1, n_当前病区id, 当前病区id),
        就诊状态 = Decode(n_就诊_b, 1, n_就诊状态, 就诊状态), 就诊诊室 = Decode(n_就诊_b, 1, v_就诊诊室, 就诊诊室),
        就诊时间 = Decode(n_就诊_b, 1, d_就诊时间, 就诊时间), 民族 = Decode(n_民族_b, 1, v_民族, 民族), 身份 = Decode(n_身份_b, 1, v_身份, 身份),
        监护人 = Decode(n_监护人_b, 1, v_监护人, 监护人), 就诊卡号 = Decode(n_就诊卡号_b, 1, v_就诊卡号, 就诊卡号),
        卡验证码 = Decode(n_卡验证码_b, 1, v_卡验证码, 卡验证码), Ic卡号 = Decode(n_Ic卡号_b, 1, v_Ic卡号, Ic卡号),
        其他证件 = Decode(n_其他证件_b, 1, v_其他证件, 其他证件), 病人类型 = Decode(n_病人类型_b, 1, v_病人类型, 病人类型),
        单位开户行 = Decode(n_单位开户行_b, 1, v_单位开户行, 单位开户行), 单位帐号 = Decode(n_单位帐号_b, 1, v_单位帐号, 单位帐号),
        Qq = Decode(n_Qq_b, 1, v_Qq, Qq), Email = Decode(n_Email_b, 1, v_Email, Email),
        单位地址 = Decode(n_单位地址_b, 1, v_单位地址, 单位地址)
    Where 病人id = n_病人id And Decode(n_主页id, Null, 0, 主页id) = Decode(n_主页id, Null, 0, n_主页id) And
          Decode(v_姓名_Old, Null, '-', 姓名) = Decode(v_姓名_Old, Null, '-', v_姓名_Old);
  
    n_清除就诊卡信息 := o_Json.Get_Number('cardno_clear');
    If Nvl(n_清除就诊卡信息, 0) = 1 Then
      Update 病人信息 Set 就诊卡号 = Null, 卡验证码 = Null, Ic卡号 = Null Where 病人id = n_病人id;
    End If;
  
    If Nvl(n_实名检查, 0) = 1 Then
      Select Zl_Fun_Checkidentify(1, n_病人id, v_Strtmpbefor) Into v_Msg From Dual;
    End If;
  End If;
  --社区信息 
  --    community_info        C     社区信息节点 
  --      num                 N  1  社区序号 
  --      code                C  1  社区号码 
  --      oper_type           N  1  社区操作类型 
  o_Json1 := Pljson();
  o_Json1 := o_Json.Get_Pljson('community_info');
  If o_Json1 Is Not Null Then
    n_社区id   := o_Json1.Get_Number('num');
    v_社区号码 := o_Json1.Get_String('code');
    n_社区类型 := o_Json1.Get_Number('oper_type');
    --更新社区号 
    If n_社区id <> 0 And v_社区号码 Is Not Null Then
      Zl_病人社区信息_Insert(n_病人id, n_社区id, v_社区号码, n_社区类型, d_登记时间);
    End If;
  End If;

  --更新地址信息 
  --    addr_list[]           C     地址信息列表 
  --      oper_fun            N  1  操作功能:1-新增,修改   2-删除 
  --      type                C  1  地址类别 
  --      state               C  1  地址_省 
  --      city                C  1  地址_市 
  --      county              C  1  地址_县 
  --      township            C  1  地址_乡 
  --      other               C  1  地址_其他 
  --      code                C  1  区划代码 
  j_Jsonlist := Pljson_List();
  j_Jsonlist := o_Json.Get_Pljson_List('addr_list');
  If j_Jsonlist Is Not Null Then
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json1    := Pljson();
      o_Json1    := Pljson(j_Jsonlist.Get(I));
      n_操作功能 := o_Json1.Get_Number('oper_fun');
      n_地址类型 := o_Json1.Get_Number('type');
      v_省       := o_Json1.Get_String('state');
      v_市       := o_Json1.Get_String('city');
      v_县       := o_Json1.Get_String('county');
      v_乡镇     := o_Json1.Get_String('township');
      v_其他     := o_Json1.Get_String('other');
      v_区划代码 := o_Json1.Get_String('code');
      n_是否就诊 := o_Json1.Get_Number('visit_or_in');
    
      Zl_病人地址信息_Update_s(n_操作功能, n_病人id, n_主页id, n_地址类型, v_省, v_市, v_县, v_乡镇, v_其他, v_区划代码, n_是否就诊);
    End Loop;
  End If;
  --      item_list[]         更新病人信息某一个字段的值 
  --      item_name           C  1  字段名 
  --      item_value          C  1   字段值 
  j_Jsonlist := Pljson_List();
  j_Jsonlist := o_Json.Get_Pljson_List('item_list');
  If j_Jsonlist Is Not Null Then
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json1  := Pljson();
      o_Json1  := Pljson(j_Jsonlist.Get(I));
      v_字段名 := o_Json1.Get_String('item_name');
      v_字段值 := o_Json1.Get_String('item_value');
      If Nvl(n_病人id, 0) <> 0 And Nvl(v_字段名, '-') <> '-' Then
        If Nvl(v_字段值, '-') = 'Null' Then
          v_Sql := 'Update 病人信息 Set ' || v_字段名 || '=Null Where 病人ID=:1';
          Execute Immediate v_Sql
            Using n_病人id;
        Else
          v_Sql := 'Update 病人信息 Set ' || v_字段名 || '=:1 Where 病人ID=:2';
          Execute Immediate v_Sql
            Using v_字段值, n_病人id;
        End If;
      End If;
    End Loop;
  End If;
  --更新病人从属信息 
  --    ext_list[]            C     病人信息从项列表 
  --      info_name           C  1  信息名 
  --      upd_info_value      N  1  修改的信息值 
  j_Jsonlist := Pljson_List();
  j_Jsonlist := o_Json.Get_Pljson_List('ext_list');
  If j_Jsonlist Is Not Null Then
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json1  := Pljson();
      o_Json1  := Pljson(j_Jsonlist.Get(I));
      v_信息名 := o_Json1.Get_String('info_name');
      v_信息值 := o_Json1.Get_String('upd_info_value');
      n_就诊id := o_Json1.Get_Number('visit_id');
      If v_信息名 Is Not Null Then
        Zl_病人信息从表_Update(n_病人id, v_信息名, v_信息值, n_就诊id);
      End If;
    End Loop;
  End If;

  --更新证件类型 
  --    cert_list[]                 证件列表(主要是当成绑卡处理) 
  --      cert_name           C  1  证件名称 
  --      cert_no             C  1  证号号码 
  j_Jsonlist := Pljson_List();
  j_Jsonlist := o_Json.Get_Pljson_List('cert_list');
  If j_Jsonlist Is Not Null Then
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json1  := Pljson();
      o_Json1  := Pljson(j_Jsonlist.Get(I));
      v_卡名称 := o_Json1.Get_String('cert_name');
      v_卡号   := o_Json1.Get_String('cert_no');
    
      If v_卡名称 Is Not Null Then
        If v_卡号 Is Not Null Then
          --检查卡号是否被他人使用 
          Select Count(1)
          Into n_Count
          From 病人医疗卡信息 A, 医疗卡类别 B
          Where a.卡类别id = b.Id And b.名称 = v_卡名称 And b.是否证件 = 1 And a.卡号 = v_卡号 And a.病人id <> n_病人id;
          If n_Count <> 0 Then
            Json_Out := Zljsonout(v_卡名称 || ':' || v_卡号 || '正在被他人使用,请检查！');
            Return;
          End If;
        
          --不存在的就诊类型需要新增卡类别管理 
          Select Nvl(Max(ID), 0), Nvl(Max(卡号长度), 0), Max(编码), Max(LPad(编码, 10)), Max(Length(编码))
          Into n_卡类别id, n_卡号长度, v_编码, n_最大值, n_最长值
          From 医疗卡类别
          Where 名称 = v_卡名称;
        
          Select Max(编码), Max(LPad(编码, 10)), Max(Length(编码)) Into v_编码, n_最大值, n_最长值 From 医疗卡类别;
        
          If v_编码 Is Null Then
            Select LPad(1, 10, '0') Into v_编码 From Dual;
          Else
            n_最大值 := n_最大值 + 1;
            Select LPad(n_最大值, n_最长值, '0') Into v_编码 From Dual;
          End If;
        
          If n_卡类别id = 0 Then
            --新增 
            Select 医疗卡类别_Id.Nextval Into n_卡类别id From Dual;
          
            Zl_医疗卡类别_Update(n_卡类别id, v_编码, v_卡名称, Substr(v_卡名称, 1, 1), Null, Length(v_卡号), 0, 1, 0, 0, 0, 0, Null, Null,
                            v_编码, 0, Null, 1, Null, 1, 10, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 1000, 0, 1);
          Elsif Length(v_卡号) > n_卡号长度 Then
            --修改长度 
            Zl_医疗卡类别_Update(n_卡类别id, v_编码, v_卡名称, Substr(v_卡名称, 1, 1), Null, Length(v_卡号), 0, 1, 0, 0, 0, 0, Null, Null,
                            v_编码, 0, Null, 1, Null, 1, 10, 0, 0, 1, 1, 0, 0, 0, 0, 0, 0, 0, 1, 0, 1000, 0, 1);
          End If;
        End If;
      
        --先清除病人卡信息 
        n_Count := 0;
        For c_证件 In (Select a.卡类别id, a.卡号
                     From 病人医疗卡信息 A
                     Where a.卡类别id = n_卡类别id And a.病人id = n_病人id) Loop
          If c_证件.卡号 = Nvl(v_卡号, '_') Then
            n_Count := 1;
          Else
            Zl_医疗卡变动_Insert_s(14, n_病人id, c_证件.卡类别id, Null, c_证件.卡号, '证件卡取消绑定', Null, v_操作员姓名, d_登记时间);
          End If;
        End Loop;
        --新增病人卡信息 
        If n_Count = 0 And v_卡号 Is Not Null Then
          Zl_医疗卡变动_Insert_s(11, n_病人id, n_卡类别id, Null, v_卡号, '证件卡绑定', Null, v_操作员姓名, d_登记时间);
        End If;
      End If;
    End Loop;
  End If;

  --更新过敏数据 
  j_Jsonlist := Pljson_List();
  j_Jsonlist := o_Json.Get_Pljson_List('allergic_drugs_list');
  If j_Jsonlist Is Not Null Then
    --清除所有记录 
    If Nvl(n_过敏药物更新, 0) = 0 Then
      Zl_病人过敏药物_Delete(n_病人id);
    End If;
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json1        := Pljson();
      o_Json1        := Pljson(j_Jsonlist.Get(I));
      n_方式         := o_Json1.Get_Number('oper_type');
      n_过敏药品id   := o_Json1.Get_Number('pat_algc_cadn_id');
      v_过敏药物名称 := o_Json1.Get_String('pat_algc_cadn');
      v_过每药物反应 := o_Json1.Get_String('allergy_info');
      If o_Json1.Get_Clob('allergic_drugs') Is Not Null Then
        c_过敏药物 := o_Json1.Get_Clob('allergic_drugs');
      End If;
      If Nvl(n_方式, 0) = 0 Then
        If v_过敏药物名称 Is Not Null Then
          If n_过敏药品id = 0 Then
            n_过敏药品id := Null;
          End If;
          Zl_病人过敏药物_Update(n_病人id, n_过敏药品id, v_过敏药物名称, v_过每药物反应);
        End If;
      End If;
      If Nvl(n_方式, 0) = 1 Then
        While c_过敏药物 Is Not Null Loop
          If Length(c_过敏药物) <= 4000 Then
            l_过敏药物.Extend;
            l_过敏药物(l_过敏药物.Count) := c_过敏药物;
            c_过敏药物 := Null;
          Else
            l_过敏药物.Extend;
            l_过敏药物(l_过敏药物.Count) := Substr(c_过敏药物, 1, Instr(c_过敏药物, ',', 3980) - 1);
            c_过敏药物 := Substr(c_过敏药物, Instr(c_过敏药物, ',', 3980) + 1);
          End If;
        End Loop;
        For I In 1 .. l_过敏药物.Count Loop
          Delete From 病人过敏药物
          Where 病人id = n_病人id And
                (过敏药物id, 过敏药物) Not In
                (Select Distinct C1 As 过敏药物id, C2 As 过敏药物 From Table(f_Str2list2(l_过敏药物(I), ',')));
        End Loop;
        If l_过敏药物.Count = 0 Then
          Delete From 病人过敏药物 Where 病人id = n_病人id;
        End If;
      End If;
    End Loop;
  End If;

  --更新免疫记录 
  --    immune_list[]         C     病人免疫列表 
  --      vaccinate_time      C  1  接种时间:yyyy-mm-dd hh24:mi:ss 
  --      vaccinate_name      C  1  接种名称 
  j_Jsonlist := Pljson_List();
  j_Jsonlist := o_Json.Get_Pljson_List('immune_list');
  If j_Jsonlist Is Not Null Then
    --清除所有记录 
    Zl_病人免疫记录_Delete(n_病人id);
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json1    := Pljson();
      o_Json1    := Pljson(j_Jsonlist.Get(I));
      d_接种时间 := To_Date(o_Json1.Get_String('vaccinate_time'), 'YYYY-MM-DD hh24:mi:ss');
      v_接种名称 := o_Json1.Get_String('vaccinate_name');
    
      If v_接种名称 Is Not Null Then
        Zl_病人免疫记录_Update(n_病人id, d_接种时间, v_接种名称);
      End If;
    End Loop;
  End If;

  --更新医疗卡属性 
  --    card_property_list[]  C     医疗卡属性列表 
  --      cardtype_id         N  1  医疗卡类别ID 
  --      card_no             C  1  卡号 
  --      info_name           C  1  信息名 
  --      info_value          N  1  信息值 
  j_Jsonlist := Pljson_List();
  j_Jsonlist := o_Json.Get_Pljson_List('card_property_list');
  If j_Jsonlist Is Not Null Then
    --清除所有记录 
    Zl_病人免疫记录_Delete(n_病人id);
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json1    := Pljson();
      o_Json1    := Pljson(j_Jsonlist.Get(I));
      n_卡类别id := o_Json1.Get_Number('cardtype_id');
      v_卡号     := o_Json1.Get_String('card_no');
      v_信息名   := o_Json1.Get_String('info_name');
      v_信息值   := o_Json1.Get_String('info_value');
    
      Zl_病人医疗卡属性_Update(n_病人id, n_卡类别id, v_卡号, v_信息名, v_信息值);
    End Loop;
  End If;

  --签约信息 
  --    sign_info             C   签约信息 
  --      card_type_id        N 1 卡类别ID 
  --      card_no             C 1 卡号 
  --      card_pwd            C   卡密码 
  --      qrcode              C   二维码 
  --      card_notes          C   变动原因 
  --      card_use_endtime    C   终止使用时间 
  o_Json1 := Pljson();
  o_Json1 := o_Json.Get_Pljson('sign_info');
  If o_Json1 Is Not Null Then
    n_卡类别id     := o_Json1.Get_Number('card_type_id');
    v_卡号         := o_Json1.Get_String('card_no');
    v_卡密码       := o_Json1.Get_String('card_pwd');
    v_变动原因     := o_Json1.Get_String('card_notes');
    d_终止使用时间 := To_Date(o_Json1.Get_String('card_use_endtime'), 'YYYY-MM-DD hh24:mi:ss');
    --签约 
    Select Count(1) Into n_Count From 医疗卡类别 Where ID = n_卡类别id;
    If n_Count = 1 Then
      Select Count(1) Into n_Count From 病人医疗卡信息 Where 卡号 = v_卡号 And 卡类别id = n_卡类别id;
      If n_Count = 0 Then
        Zl_医疗卡变动_Insert_s(11, n_病人id, n_卡类别id, '', v_卡号, v_变动原因, v_卡密码, v_操作员姓名, d_登记时间, Null, d_终止使用时间);
      End If;
    End If;
  End If;
  b_Message.Zlhis_Patient_016(n_病人id);

  Json_Out := Zljsonout('成功', 1);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Updatepatiarchives;
/
Create Or Replace Procedure Zl_Patisvr_Updatepatibaseinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:更新病人基本信息
  --入参：JSON格式
  --input
  --  pati_id               N 1 病人id
  --  visit_id              N 1 就诊id
  --  model                 N 1 模块
  --  pati_name_n           C 1 姓名
  --  pati_sex_n            C 1 性别
  --  pati_age_n            C 1 年龄
  --  pati_birthdate_n      C 1 出生日期
  --  occasion              N 1 场合 1-门诊;2-住院
  --  pati_name_o           C 1 姓名
  --  pati_sex_o            C 1 性别
  --  pati_age_o            C 1 年龄
  --  pati_birthdate_o      C 1 出生日期
  --  explain               C 1 说明
  --出参：JSON格式
  --output
  --   code                 N  1  应答码：0-失败；1-成功
  --   message              C  1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Json        Pljson;
  j_Jsonin      Pljson;
  v_Username    人员表.姓名%Type;
  d_变动时间    病人信息变动.变动时间%Type;
  v_说明        病人信息变动.说明%Type;
  v_Strtmpbefor Varchar2(4000);
  v_Msg         Varchar2(4000);
  n_病人id      病人信息变动.病人id%Type;
  v_模块        病人信息变动.变动模块%Type;
  v_姓名_n      病人信息.姓名%Type;
  v_性别_n      病人信息.性别%Type;
  v_年龄_n      病人信息.年龄%Type;
  d_出生日期_n  病人信息.出生日期%Type;
  v_姓名_o      病人信息.姓名%Type;
  v_性别_o      病人信息.性别%Type;
  v_年龄_o      病人信息.年龄%Type;
  d_出生日期_o  病人信息.出生日期%Type;
Begin
  j_Jsonin     := Pljson(Json_In);
  j_Json       := j_Jsonin.Get_Pljson('input');
  n_病人id     := j_Json.Get_Number('pati_id');
  v_模块       := j_Json.Get_String('model');
  v_姓名_n     := j_Json.Get_String('pati_name_n');
  v_性别_n     := j_Json.Get_String('pati_sex_n');
  v_年龄_n     := j_Json.Get_String('pati_age_n');
  d_出生日期_n := To_Date(j_Json.Get_String('pati_birthdate_n'), 'yyyy-mm-dd hh24:mi:ss');
  v_姓名_o     := j_Json.Get_String('pati_name_o');
  v_性别_o     := j_Json.Get_String('pati_sex_o');
  v_年龄_o     := j_Json.Get_String('pati_age_o');
  d_出生日期_o := To_Date(j_Json.Get_String('pati_birthdate_o'), 'yyyy-mm-dd hh24:mi:ss');
  v_说明       := j_Json.Get_String('explain');
  v_Username   := zl_UserName;
  --3、体检部分
  --体检部分不调用子过程,因体检系统的就诊记录由体检系统自己产生,所以此处传入n_就诊id无法对体检系统进行修正。
  --体检系统提供单独的修改入口。
  --4、PACS部分
  Select Zl_Fun_Checkidentify(0, n_病人id, v_Strtmpbefor) Into v_Strtmpbefor From Dual;
  Update 病人信息
  Set 姓名 = v_姓名_n, 性别 = v_性别_n, 年龄 = v_年龄_n, 出生日期 = d_出生日期_n
  Where 病人id = n_病人id;
  Select Zl_Fun_Checkidentify(1, n_病人id, v_Strtmpbefor) Into v_Msg From Dual;

  d_变动时间 := Sysdate;
  If Nvl(v_姓名_n, '_') <> Nvl(v_姓名_o, '_') Then
    Insert Into 病人信息变动
      (病人id, 变动项目, 原信息, 新信息, 变动时间, 变动人, 变动模块, 说明)
    Values
      (n_病人id, '姓名', v_姓名_o, v_姓名_n, d_变动时间, v_Username, v_模块, v_说明);
  End If;
  If Nvl(v_性别_n, '_') <> Nvl(v_性别_o, '_') Then
    Insert Into 病人信息变动
      (病人id, 变动项目, 原信息, 新信息, 变动时间, 变动人, 变动模块, 说明)
    Values
      (n_病人id, '性别', v_性别_o, v_性别_n, d_变动时间, v_Username, v_模块, v_说明);
  End If;
  If Nvl(v_年龄_n, '_') <> Nvl(v_年龄_o, '_') Then
    Insert Into 病人信息变动
      (病人id, 变动项目, 原信息, 新信息, 变动时间, 变动人, 变动模块, 说明)
    Values
      (n_病人id, '年龄', v_年龄_o, v_年龄_n, d_变动时间, v_Username, v_模块, v_说明);
  End If;
  If Nvl(d_出生日期_n, Sysdate) <> Nvl(d_出生日期_o, Sysdate) Then
    Insert Into 病人信息变动
      (病人id, 变动项目, 原信息, 新信息, 变动时间, 变动人, 变动模块, 说明)
    Values
      (n_病人id, '出生日期', To_Char(d_出生日期_o, 'YYYY-MM-DD hh24:mi'), To_Char(d_出生日期_n, 'YYYY-MM-DD hh24:mi'), d_变动时间,
       v_Username, v_模块, v_说明);
  End If;
  b_Message.Zlhis_Patient_016(n_病人id);
  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Updatepatibaseinfo;
/
Create Or Replace Procedure Zl_Patisvr_Updatepatirelate
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：针对指定病人的信息进行关联
  --入参：Json_In:格式
  --  input
  --    oper_fun            N   1 操作功能:0-增加关联;1-取消关联;2-更新关联ID;3-入院登记自动关联
  --    relate_id           N     关联ID
  --    relate_pati_ids     C   1 需要关联的病人ids:多个用逗号
  --    operator_name       C   1 操作员姓名
  --    operator_time       C   1 操作时间:yyyy-mm-dd hh24:mi:ss
  --出参: Json_Out,格式如下
  --  output
  --    code                 N   1   应答吗：0-失败；1-成功
  --    message              C   1   应答消息：失败时返回具体的错误信息
  ----------------------------------------------------------------------------

  v_病人ids  Varchar2(32680);
  n_操作类型 Number(1);
  n_关联id   病人身份关联.关联id%Type;
  v_操作员   病人身份关联.操作人员%Type;
  d_操作时间 病人身份关联.操作时间%Type;
  j_Json     Pljson;
  j_Jsonin   Pljson;
Begin

  --解析入参
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_操作类型 := j_Json.Get_Number('oper_fun');
  n_关联id   := j_Json.Get_Number('relate_id');
  v_病人ids  := j_Json.Get_String('relate_pati_ids');
  v_操作员   := j_Json.Get_String('operator_name');
  d_操作时间 := To_Date(j_Json.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');

  Zl_病人身份关联_Update(n_操作类型, n_关联id, v_病人ids, v_操作员, d_操作时间);

  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Updatepatirelate;
/
Create Or Replace Procedure Zl_Patisvr_Updateproxy
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --功能：代办人信息更新
  --入参：Json_In:格式
  --input
  --  pati_id               N 1 病人id_In
  --  visit_id              N 1 就诊ID
  --  pati_idcard           C 1 身份证号
  --  proxy_name            C 1 代办人姓名
  --  proxy_idno            C 1 代办人身份证号
  --  proxy_sex             C 1 代办人性别
  --  pati_age              C 1 代办人年龄
  --  proxy_phno            C 1 代办人电话
  --  reason                C 1 用药理由
  --出参: Json_Out,格式如下
  --output
  --  code                  N 1 应答吗：0-失败；1-成功
  --  message               C 1 应答消息：失败时返回具体的错误信息
  -------------------------------------------
  j_Json   Pljson;
  j_Jsonin Pljson;
Begin
  --解析入参
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  Zl_代办人信息_Insert(j_Json.Get_Number('pati_id'), j_Json.Get_String('pati_idcard'), j_Json.Get_String('proxy_name'),
                  j_Json.Get_String('proxy_idno'), j_Json.Get_Number('visit_id'), j_Json.Get_String('proxy_sex'),
                  j_Json.Get_String('pati_age'), j_Json.Get_String('proxy_phno'), j_Json.Get_String('reason'));
  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Updateproxy;
/
Create Or Replace Procedure Zl_Patisvr_Updcommunityinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --功能：用于门诊医生站对社区病人进行补充身份验证时使用
  --入参:Json_In:格式
  --input
  --    pati_id               N  1 病人ID
  --    community_num         N  1  社区序号
  --    community_code        C  1  社区号码
  --    community_oper_type   N  1  社区操作类型
  --    visit_time            C  1  就诊时间:yyyy-mm-dd hh24:mi:ss
  --    pati_name             C  1  姓名
  --    pati_sex              C  1  性别
  --    pati_age              C  1  年龄
  --    pati_birthdate        C  1  出生日期:yyyy-mm-dd hh24:mi:ss
  --    pat_baddr             C  1  出生地点
  --    pati_idcard           C  1  身份证号
  --    nation_name           C  1  民族
  --    country_name          C  1  国籍
  --    mari_name             C  1  婚姻状况
  --    ocpt_name             C  1  职业
  --    pat_home_addr         C  1  家庭地址
  --    pat_home_phno         C  1  家庭电话
  --    pat_home_postcode     C  1  家庭地址邮编
  --    emp_name              C  1  工作单位
  --    emp_phno              C  1  单位电话
  --    emp_postcode          C  1  单位邮编
  --    contacts_name         C  1  联系人姓名
  --    contacts_relation     C  1  联系人关系
  --    ontacts_phno          C  1  联系人电话
  --    ontacts_addr          C  1  联系人地址
  --    pat_hous_addr         C  1  户口地址
  --    pat_hous_postcode     C  1  户口地址邮编

  -- 出参:
  --  output
  --    code                            N 1 应答吗:0-失败；1-成功
  --    message                         C 1 应答消息:失败时返回具体的错误信息
  -------------------------------------------

  n_病人id       Number;
  n_社区序号     Number;
  v_社区号码     Varchar2(20);
  n_社区操作类型 Number;
  d_就诊时间     Date;
  v_姓名         病人信息.姓名%Type;
  v_性别         病人信息.性别%Type;
  v_年龄         病人信息.年龄%Type;
  d_出生日期     病人信息.出生日期%Type;
  v_出生地点     病人信息.出生地点%Type;
  v_身份证号     病人信息.身份证号%Type;
  v_民族         病人信息.民族%Type;
  v_国籍         病人信息.国籍%Type;
  v_婚姻状况     病人信息.婚姻状况%Type;
  v_职业         病人信息.职业%Type;
  v_家庭地址     病人信息.家庭地址%Type;
  v_家庭电话     病人信息.家庭电话%Type;
  v_家庭地址邮编 病人信息.家庭地址邮编%Type;
  v_工作单位     病人信息.工作单位%Type;
  v_单位电话     病人信息.单位电话%Type;
  v_单位邮编     病人信息.单位邮编%Type;
  v_联系人姓名   病人信息.联系人姓名%Type;
  v_联系人关系   病人信息.联系人关系%Type;
  v_联系人电话   病人信息.联系人电话%Type;
  v_联系人地址   病人信息.联系人地址%Type;
  v_户口地址     病人信息.户口地址%Type;
  v_户口地址邮编 病人信息.户口地址邮编%Type;

  j_Json   Pljson;
  j_Jsonin Pljson;

  v_Strtmpbefor Varchar2(4000);
  v_Msg         Varchar2(4000);
Begin
  --解析入参
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');

  n_病人id       := j_Json.Get_Number('pati_id');
  n_社区序号     := j_Json.Get_Number('community_num');
  v_社区号码     := j_Json.Get_String('community_code');
  n_社区操作类型 := j_Json.Get_Number('community_oper_type');
  d_就诊时间     := To_Date(j_Json.Get_String('visit_time'), 'YYYY-MM-DD hh24:mi:ss');
  v_姓名         := j_Json.Get_String('pati_name');
  v_性别         := j_Json.Get_String('pati_sex');
  v_年龄         := j_Json.Get_String('pati_age');
  d_出生日期     := To_Date(j_Json.Get_String('pati_birthdate'), 'YYYY-MM-DD hh24:mi:ss');
  v_出生地点     := j_Json.Get_String('pat_baddr');
  v_身份证号     := j_Json.Get_String('pati_idcard');
  v_民族         := j_Json.Get_String('nation_name');
  v_国籍         := j_Json.Get_String('country_name');
  v_婚姻状况     := j_Json.Get_String('mari_name');
  v_职业         := j_Json.Get_String('ocpt_name');
  v_家庭地址     := j_Json.Get_String('pat_home_addr');
  v_家庭电话     := j_Json.Get_String('pat_home_phno');
  v_家庭地址邮编 := j_Json.Get_String('pat_home_postcode');
  v_工作单位     := j_Json.Get_String('emp_name');
  v_单位电话     := j_Json.Get_String('emp_phno');
  v_单位邮编     := j_Json.Get_String('emp_postcode');
  v_联系人姓名   := j_Json.Get_String('contacts_name');
  v_联系人关系   := j_Json.Get_String('contacts_relation');
  v_联系人电话   := j_Json.Get_String('ontacts_phno');
  v_联系人地址   := j_Json.Get_String('ontacts_addr');
  v_户口地址     := j_Json.Get_String('pat_hous_addr');
  v_户口地址邮编 := j_Json.Get_String('pat_hous_postcode');

  If d_就诊时间 Is Null Then
    d_就诊时间 := Sysdate;
  End If;

  Zl_病人社区信息_Insert(n_病人id, n_社区序号, v_社区号码, n_社区操作类型, d_就诊时间);

  Select Zl_Fun_Checkidentify(0, n_病人id, v_Strtmpbefor) Into v_Strtmpbefor From Dual;
  Update 病人信息
  Set 姓名 = Decode(v_姓名, Null, 姓名, v_姓名), 性别 = Decode(v_性别, Null, 性别, v_性别), 年龄 = Decode(v_年龄, Null, 年龄, v_年龄),
      出生日期 = Decode(d_出生日期, Null, 出生日期, d_出生日期), 出生地点 = Decode(v_出生地点, Null, 出生地点, v_出生地点),
      身份证号 = Decode(v_身份证号, Null, 身份证号, v_身份证号), 民族 = Decode(v_民族, Null, 民族, v_民族), 国籍 = Decode(v_国籍, Null, 国籍, v_国籍),
      婚姻状况 = Decode(v_婚姻状况, Null, 婚姻状况, v_婚姻状况), 职业 = Decode(v_职业, Null, 职业, v_职业),
      家庭地址 = Decode(v_家庭地址, Null, 家庭地址, v_家庭地址), 家庭电话 = Decode(v_家庭电话, Null, 家庭电话, v_家庭电话),
      家庭地址邮编 = Decode(v_家庭地址邮编, Null, 家庭地址邮编, v_家庭地址邮编), 工作单位 = Decode(v_工作单位, Null, 工作单位, v_工作单位),
      单位电话 = Decode(v_单位电话, Null, 单位电话, v_单位电话), 单位邮编 = Decode(v_单位邮编, Null, 单位邮编, v_单位邮编),
      联系人姓名 = Decode(v_联系人姓名, Null, 联系人姓名, v_联系人姓名), 联系人关系 = Decode(v_联系人姓名, Null, 联系人关系, v_联系人关系),
      联系人电话 = Decode(v_联系人姓名, Null, 联系人电话, v_联系人电话), 联系人地址 = Decode(v_联系人姓名, Null, 联系人地址, v_联系人地址),
      户口地址 = Decode(v_户口地址, Null, 户口地址, v_户口地址), 户口地址邮编 = Decode(v_户口地址邮编, Null, 户口地址邮编, v_户口地址邮编)
  Where 病人id = n_病人id;
  Select Zl_Fun_Checkidentify(1, n_病人id, v_Strtmpbefor) Into v_Msg From Dual;

  Json_Out := Zljsonout('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Updcommunityinfo;
/
Create Or Replace Procedure Zl_Patisvr_Updpatiaddressinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:修改病案主页从表相关信息
  --入参：Json_In:格式
  --    input
  --      oper_fun          N 1 操作功能:1-新增,修改   2-删除
  --      pati_id           N 1 病人id
  --      pati_pageid       N 1 主页Id
  --      pat_addr_type     C 1 地址类别
  --      pat_addr_state    C 1 地址_省
  --      pat_addr_city     C 1 地址_市
  --      pat_addr_county   C 1 地址_县
  --      pat_addr_township C 1 地址_乡
  --      pat_addr_other    C 1 地址_其他
  --      pat_region_code   C 1 区划代码

  --出参: Json_Out,格式如下
  --  output
  --    code              N   1   应答码：0-失败；1-成功
  --    message           C   1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------

  j_Json     Pljson;
  j_Jsonin   Pljson;
  n_病人id   病人信息.病人id%Type;
  n_主页id   病人信息.主页id%Type;
  n_功能     Number(3);
  v_地址类别 病人地址信息.地址类别%Type;
  v_省       病人地址信息.省%Type;
  v_市       病人地址信息.市%Type;
  v_县       病人地址信息.县%Type;
  v_乡镇     病人地址信息.乡镇%Type;
  v_其他     病人地址信息.其他%Type;
  v_区划代码 病人地址信息.区划代码%Type;

Begin
  --解析入参
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_功能     := j_Json.Get_Number('oper_fun');
  n_病人id   := j_Json.Get_Number('pati_id');
  n_主页id   := j_Json.Get_Number('pati_pageid');
  v_地址类别 := j_Json.Get_String('pat_addr_type');

  v_省       := j_Json.Get_String('pat_addr_state');
  v_市       := j_Json.Get_String('pat_addr_city');
  v_县       := j_Json.Get_String('pat_addr_county');
  v_乡镇     := j_Json.Get_String('pat_addr_township');
  v_其他     := j_Json.Get_String('pat_addr_other');
  v_区划代码 := j_Json.Get_String('pat_region_code');

  Zl_病人地址信息_Update_s(n_功能, n_病人id, n_主页id, v_地址类别, v_省, v_市, v_县, v_乡镇, v_其他, v_区划代码);

  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Updpatiaddressinfo;
/
Create Or Replace Procedure Zl_Patisvr_Updpatiallerinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --功能：用于更新病人信息的过敏药物信息
  --入参:Json_In:格式
  --input
  --    aller_list    病人过敏药物列表
  --      execute_type   N 1 执行方式 1-删除 2-新增或者更新
  --      pati_id        N 1 病人id
  --      drug_id        N 1 药物id
  --      drug_name      C 1 药物名
  --      aller_reflex   C 1 过敏反应

  -- 出参:
  --  output
  --    code                            N 1 应答吗:0-失败；1-成功
  --    message                         C 1 应答消息:失败时返回具体的错误信息
  -------------------------------------------

  n_病人id   Number;
  n_Type     Number;
  n_药物id   病人过敏药物.过敏药物id%Type;
  v_过敏药物 病人过敏药物.过敏药物%Type;
  v_过敏反应 病人过敏药物.过敏反应%Type;
  o_Json     Pljson;

  j_Json           Pljson;
  j_Jsonin         Pljson;
  j_Json_Allerlist Pljson_List := Pljson_List();

Begin
  --解析入参
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');

  j_Json_Allerlist := j_Json.Get_Pljson_List('aller_list');
  For I In 1 .. j_Json_Allerlist.Count Loop
    o_Json     := Pljson();
    o_Json     := Pljson(j_Json_Allerlist.Get(I));
    n_Type     := o_Json.Get_Number('execute_type');
    n_病人id   := o_Json.Get_Number('pati_id');
    n_药物id   := o_Json.Get_Number('drug_id');
    v_过敏药物 := o_Json.Get_String('drug_name');
    v_过敏反应 := o_Json.Get_String('aller_reflex');
    If Nvl(n_Type, 0) = 1 Then
      --如果没有过敏的记录就删除该药品的过敏记录
      Delete From 病人过敏药物 A Where a.病人id = n_病人id And a.过敏药物 = v_过敏药物 And a.过敏药物id = n_药物id;
    Else
      Update 病人过敏药物
      Set 过敏反应 = v_过敏反应, 过敏药物id = n_药物id
      Where 病人id = n_病人id And 过敏药物 = v_过敏药物;
      If Sql%RowCount = 0 Then
        Insert Into 病人过敏药物
          (病人id, 过敏药物id, 过敏药物, 过敏反应)
        Values
          (n_病人id, n_药物id, v_过敏药物, v_过敏反应);
      End If;
    End If;
  End Loop;

  Json_Out := Zljsonout('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Updpatiallerinfo;
/
Create Or Replace Procedure Zl_Patisvr_Updpatifamilyinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：更新病人家属信息
  --入参：Json_In:格式
  --  input
  --  opr_fun      N  1  操作方式  --1-新增,2-更新,3-假删除
  --  pati_id      N  1  病人ID
  --  family_id    N  1  家属id
  --  reg_name     C  1  登记人
  --  reg_time     C  1  登记时间
  --  relation     C  0  关系
  --  cancel_name  C  0  撤档人
  --  cancel_time  C  0  撤档时间
  --出参: Json_Out,格式如下
  --  output
  --    code                 N   1   应答吗：0-失败；1-成功
  --    message              C   1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------

  n_操作     Number;
  n_病人id   病人家属.病人id%Type;
  n_家属id   病人家属.家属id%Type;
  v_登记人   病人家属.登记人%Type;
  d_登记时间 病人家属.登记时间%Type;
  v_关系     病人家属.关系%Type;
  v_撤档人   病人家属.撤档人%Type;
  v_撤档时间 病人家属.撤档时间%Type;
  j_Json     Pljson;
  j_Jsonin   Pljson;
Begin

  --解析入参
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_操作     := j_Json.Get_Number('opr_fun');
  n_病人id   := j_Json.Get_Number('pati_id');
  n_家属id   := j_Json.Get_Number('family_id');
  v_登记人   := j_Json.Get_String('reg_name');
  d_登记时间 := To_Date(j_Json.Get_String('reg_time'), 'yyyy-mm-dd hh24:mi:ss');
  v_关系     := j_Json.Get_String('relation');
  v_撤档人   := j_Json.Get_String('cancel_name');
  v_撤档时间 := To_Date(j_Json.Get_String('cancel_time'), 'yyyy-mm-dd hh24:mi:ss');
  Zl_病人家属_Update(n_操作, n_病人id, n_家属id, v_登记人, d_登记时间, v_关系, v_撤档人, v_撤档时间);
  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Updpatifamilyinfo;
/