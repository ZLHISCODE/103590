Create Or Replace Procedure Zl_Exsesvr_Addeinvoice
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：增加电子票据信息
  --入参：Json_In:格式
  --  input
  --    balance_id          N  1  结算ID
  --    balance_delid       N     退款ID:退款开具红票时有效：目前只有预交款有效,填写的是退款预交ID
  --    einvoice_id         N  1  电子票据ID
  --    operator_code       C  1  操作员编号
  --    operator_name       C  1  操作员姓名
  --    happen_time         C  1  登记时间:yyyy-mm-dd hh24:mi:ss
  --    pati_info           C     病人信息
  --      pati_id           N  1  病人ID
  --      pati_pageid       N     主页ID
  --      pati_name         C  1  姓名
  --      pati_sex          C  1  性别
  --      pati_age          C  1  年龄
  --      outpatient_num    C  1  门诊号
  --      inpatient_num     C  1  住院号
  --    einvoce_info        C     电子票据信息
  --      invoice_type      N  1  票种：1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
  --      placeCode         C  1  开票点编码
  --      inv_total         N  1  开票金额
  --      inv_oldid         N     原票据ID
  --      sys_source        C  1  系统来源
  --      demo              C  1  备注
  --      einvoice_code     C  1  电子票据代码
  --      einvoice_no       C  1  电子票据号码
  --      einvoice_random   C  1  电子校验码
  --      voucher_code      C  1  预交金凭证代码
  --      voucher_no        C  1  预交金凭证号码
  --      voucher_random    C  1  预交金凭证校验码
  --      create_time       C  1  电子票据生成时间:yyyymmddhh24miss
  --      picture_url       C  1  电子票据H5页面URL
  --      picture_neturl    C  1  电子票据外网H5页面URL
  --      qrcode            C  1  电子票据二维码图片数据:该值已Base64编码，解析时需要Base64解码,图片格式为:PNG
  --    --出参: Json_Out,格式如下
  --  output
  --    code                N 1 应答吗：0-失败；1-成功
  --    message             C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  n_Id         电子票据使用记录.Id%Type;
  n_票种       电子票据使用记录.票种%Type;
  n_记录状态   电子票据使用记录.记录状态%Type;
  n_结算id     电子票据使用记录.结算id%Type;
  n_病人id     电子票据使用记录.病人id%Type;
  v_姓名       电子票据使用记录.姓名%Type;
  v_性别       电子票据使用记录.性别%Type;
  v_年龄       电子票据使用记录.年龄%Type;
  n_门诊号     电子票据使用记录.门诊号%Type;
  n_住院号     电子票据使用记录.住院号%Type;
  v_代码       电子票据使用记录.代码%Type;
  v_号码       电子票据使用记录.号码%Type;
  v_检验码     电子票据使用记录.检验码%Type;
  v_凭证代码   电子票据使用记录.凭证代码%Type;
  v_凭证号码   电子票据使用记录.凭证号码%Type;
  v_凭证检验码 电子票据使用记录.凭证检验码%Type;
  n_票据金额   电子票据使用记录.票据金额%Type;
  v_生成时间   电子票据使用记录.生成时间%Type;
  v_Url内网    电子票据使用记录.Url内网%Type;
  v_Url外网    电子票据使用记录.Url外网%Type;
  c_二维码     Clob;
  n_原票据id   电子票据使用记录.原票据id%Type;
  n_退款id     电子票据使用记录.退款id%Type;
  v_备注       电子票据使用记录.备注%Type;
  v_开票点     电子票据使用记录.开票点%Type;
  v_系统来源   电子票据使用记录.系统来源%Type;
  v_操作员编号 电子票据使用记录.操作员编号%Type;
  v_操作员姓名 电子票据使用记录.操作员姓名%Type;
  d_登记时间   电子票据使用记录.登记时间%Type;

  n_记录状态 Number(2);
  j_Input    PLJson;
  j_Json     PLJson;
  j_Temp     PLJson;
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_结算id     := j_Json.Get_Number('balance_id');
  n_退款id     := j_Json.Get_Number('balance_delid');
  n_Id         := j_Json.Get_Number('einvoice_id');
  v_操作员编号 := j_Json.Get_String('operator_code');
  v_操作员姓名 := j_Json.Get_String('operator_name');
  d_登记时间   := Nvl(To_Date(j_Json.Get_String('happen_time'), 'yyyy-mm-dd hh24:mi:ss'), Sysdate);

  --读取病人信息

  If Not j_Json.Exist('pati_info') Then
  
    Json_Out := zlJsonOut('无病人信息，不能增加电子票据。');
    Return;
  End If;

  j_Temp   := j_Json.Get_Pljson('pati_info');
  n_病人id := j_Temp.Get_Number('pati_id');
  --n_主页id := j_Temp.Get_Number('pati_pageid');

  v_姓名   := j_Temp.Get_String('pati_name');
  v_性别   := j_Temp.Get_String('pati_sex');
  v_年龄   := j_Temp.Get_String('pati_age');
  n_门诊号 := j_Temp.Get_Number('outpatient_num');
  n_住院号 := j_Temp.Get_Number('inpatient_num');

  If Not j_Json.Exist('einvoce_info') Then
  
    Json_Out := zlJsonOut('无电子票据信息,不能增加电子票据。');
    Return;
  End If;

  --读取电子票据信息
  j_Temp := PLJson();
  j_Temp := j_Json.Get_Pljson('einvoce_info');
  --票种:1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡
  n_票种       := Nvl(j_Temp.Get_Number('invoice_type'), 1);
  v_开票点     := j_Temp.Get_String('placeCode');
  n_原票据id   := j_Temp.Get_Number('inv_oldid');
  v_系统来源   := j_Temp.Get_String('sys_source');
  v_备注       := j_Temp.Get_String('demo');
  v_代码       := j_Temp.Get_String('einvoice_code');
  v_号码       := j_Temp.Get_String('einvoice_no');
  v_检验码     := j_Temp.Get_String('einvoice_random');
  v_凭证代码   := j_Temp.Get_String('voucher_code');
  v_凭证号码   := j_Temp.Get_String('voucher_no');
  v_凭证检验码 := j_Temp.Get_String('voucher_random');
  n_票据金额   := j_Temp.Get_Number('inv_total');
  v_生成时间   := j_Temp.Get_String('create_time');
  v_Url内网    := j_Temp.Get_String('picture_url');
  v_Url外网    := j_Temp.Get_String('picture_neturl');
  c_二维码     := j_Temp.Get_Clob('qrcode');

  --增加电子票据信息
  Zl_电子票据使用记录_Insert(n_Id, n_票种, n_结算id, n_病人id, v_姓名, v_性别, v_年龄, n_门诊号, n_住院号, n_票据金额, v_开票点, v_系统来源, v_生成时间, v_备注,
                     v_操作员编号, v_操作员姓名, d_登记时间, n_原票据id, n_退款id, v_代码, v_号码, v_检验码, v_凭证代码, v_凭证号码, v_凭证检验码, v_Url内网,
                     v_Url外网);
  --更新二维码
  Insert Into 电子票据二维码 (使用记录id, 二维码) Values (n_Id, c_二维码);
  Json_Out := zlJsonOut('成功', 1);
  Return;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Addeinvoice;
/
Create Or Replace Procedure Zl_Exsesvr_Deleinvoice
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：删除电子票据信息
  --入参：Json_In:格式
  -- input      
  --  einvoice_id  N  1  电子票据ID
  --  operator_code  C  1  操作员编号
  --  operator_name  C  1  操作员姓名
  --  create_time  C  1  登记时间:yyyy-mm-dd hh24:mi:ss
  --  einvoce_info  C    电子票据信息
  --    placeCode  C  1  开票点编码
  --    sys_source  C  1  系统来源
  --    demo  C  1  备注
  --    inv_oldid  N    原票据ID
  --    einvoice_code  C  1  电子票据代码
  --    einvoice_no  C  1  电子票据号码
  --    einvoice_random  C  1  电子校验码
  --    voucher_code  C  1  预交金凭证代码
  --    voucher_no  C  1  预交金凭证号码
  --    voucher_random  C  1  预交金凭证校验码
  --    happen_time  C  1  电子票据生成时间:yyyymmddhh24miss
  --    picture_url  C  1  电子票据H5页面URL
  --    picture_neturl  C  1  电子票据外网H5页面URL
  --    qrcode  C  1  电子票据二维码图片数据:该值已Base64编码，解析时需要Base64解码,图片格式为:PNG
  --    --出参: Json_Out,格式如下
  --  output
  --    code          N 1 应答吗：0-失败；1-成功
  --    message       C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  n_Id         电子票据使用记录.Id%Type;
  v_代码       电子票据使用记录.代码%Type;
  v_号码       电子票据使用记录.号码%Type;
  v_检验码     电子票据使用记录.检验码%Type;
  v_凭证代码   电子票据使用记录.凭证代码%Type;
  v_凭证号码   电子票据使用记录.凭证号码%Type;
  v_凭证检验码 电子票据使用记录.凭证检验码%Type;
  v_生成时间   电子票据使用记录.生成时间%Type;
  v_Url内网    电子票据使用记录.Url内网%Type;
  v_Url外网    电子票据使用记录.Url外网%Type;
  c_二维码     Clob;
  n_原票据id   电子票据使用记录.原票据id%Type;
  v_备注       电子票据使用记录.备注%Type;
  v_开票点     电子票据使用记录.开票点%Type;
  v_系统来源   电子票据使用记录.系统来源%Type;
  v_操作员编号 电子票据使用记录.操作员编号%Type;
  v_操作员姓名 电子票据使用记录.操作员姓名%Type;
  d_登记时间   电子票据使用记录.登记时间%Type;

  j_Input PLJson;
  j_Json  PLJson;
  j_Temp  PLJson;
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_Id         := j_Json.Get_Number('einvoice_id');
  v_操作员编号 := j_Json.Get_String('operator_code');
  v_操作员姓名 := j_Json.Get_String('operator_name');
  d_登记时间   := Nvl(To_Date(j_Json.Get_String('create_time'), 'yyyy-mm-dd hh24:mi:ss'), Sysdate);

  If Not j_Json.Exist('einvoce_info') Then
  
    Json_Out := zlJsonOut('无电子票据信息,不能增加电子票据。');
    Return;
  End If;

  --读取电子票据信息
  j_Temp     := PLJson();
  j_Temp     := j_Json.Get_Pljson('einvoce_info');
  v_开票点   := j_Temp.Get_String('placeCode');
  v_系统来源 := j_Temp.Get_String('sys_source');
  v_备注     := j_Temp.Get_String('demo');
  n_原票据id := j_Temp.Get_Number('inv_oldid');

  v_代码       := j_Temp.Get_String('einvoice_code');
  v_号码       := j_Temp.Get_String('einvoice_no');
  v_检验码     := j_Temp.Get_String('einvoice_random');
  v_凭证代码   := j_Temp.Get_String('voucher_code');
  v_凭证号码   := j_Temp.Get_String('voucher_no');
  v_凭证检验码 := j_Temp.Get_String('voucher_random');
  v_生成时间   := j_Temp.Get_String('happen_time');
  v_Url内网    := j_Temp.Get_String('picture_url');
  v_Url外网    := j_Temp.Get_String('picture_neturl');
  c_二维码     := j_Temp.Get_Clob('qrcode');

  Zl_电子票据使用记录_Delete(n_Id, v_开票点, v_系统来源, v_生成时间, v_备注, v_操作员编号, v_操作员姓名, d_登记时间, n_原票据id, v_代码, v_号码, v_检验码, v_凭证代码,
                     v_凭证号码, v_凭证检验码, v_Url内网, v_Url外网);
  --更新二维码
  Insert Into 电子票据二维码 (使用记录id, 二维码) Values (n_Id, c_二维码);
  Json_Out := zlJsonOut('成功', 1);
  Return;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Deleinvoice;
/
Create Or Replace Procedure Zl_Exsesvr_Savepaperinvoice
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：增加纸质票据使用信息
  --入参：Json_In:格式
  --   input
  --    oper_mode           N  1  操作方式:0-换开;1-重新换开;2-作废票据;3-回收票据
  --    einvoice_id         N  1  电子票据ID
  --    operator_code       C  1  操作员编号
  --    operator_name       C  1  操作员姓名
  --    create_time         C  1  登记时间:yyyy-mm-dd hh24:mi:ss
  --    paperinv_info       C     纸质票据信息:存在多条时，请按操作顺序上传(避免数据错误)
  --      inv_occasion      N  1  应用场合:1-收费,2-预交(包含押金),3-结帐,4-挂号,5-就诊卡
  --      invoice_type      N  1  票种:1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡
  --      inv_red           N     是否红票:1-红票;0-非红票
  --      invoice_no        C  1  发票号
  --      inv_total         N  1  发票金额
  --      recv_id           N     领用id
  --    einvoce_info        C     电子票据信息
  --      placeCode         C  1  开票点编码
  --      sys_source        C  1  系统来源
  --      demo              C  1  备注
  --      einvoice_id       N  1  电子票据ID(冲销)
  --      inv_oldid         N     原票据ID
  --      einvoice_code     C  1  电子票据代码
  --      einvoice_no       C  1  电子票据号码
  --      einvoice_random   C  1  电子校验码
  --      voucher_code      C  1  预交金凭证代码
  --      voucher_no        C  1  预交金凭证号码
  --      voucher_random    C  1  预交金凭证校验码
  --      happen_time       C  1  电子票据生成时间:yyyymmddhh24miss
  --      picture_url       C  1  电子票据H5页面URL
  --      picture_neturl    C  1  电子票据外网H5页面URL
  --      qrcode            C  1  电子票据二维码图片数据:该值已Base64编码，解析时需要Base64解码,图片格式为:PNG

  --    --出参: Json_Out,格式如下
  --  output
  --    code          N 1 应答吗：0-失败；1-成功
  --    message       C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  n_Id         电子票据使用记录.Id%Type;
  n_冲销id     电子票据使用记录.Id%Type;
  v_操作员编号 电子票据使用记录.操作员编号%Type;
  v_操作员姓名 电子票据使用记录.操作员姓名%Type;
  d_登记时间   电子票据使用记录.登记时间%Type;
  n_结算id     电子票据使用记录.结算id%Type;
  v_发票号     票据使用明细.号码%Type;
  n_发票金额   票据使用明细.票据金额%Type;
  n_领用id     票据使用明细.领用id%Type;
  n_操作方式   Number(2);
  n_应用场合   Number(2);
  n_票种       Number(2);
  n_是否红票   Number(2);
  v_代码       电子票据使用记录.代码%Type;
  v_号码       电子票据使用记录.号码%Type;
  v_检验码     电子票据使用记录.检验码%Type;
  v_凭证代码   电子票据使用记录.凭证代码%Type;
  v_凭证号码   电子票据使用记录.凭证号码%Type;
  v_凭证检验码 电子票据使用记录.凭证检验码%Type;
  v_生成时间   电子票据使用记录.生成时间%Type;
  v_Url内网    电子票据使用记录.Url内网%Type;
  v_Url外网    电子票据使用记录.Url外网%Type;
  c_二维码     Clob;
  n_原票据id   电子票据使用记录.原票据id%Type;
  v_备注       电子票据使用记录.备注%Type;
  v_开票点     电子票据使用记录.开票点%Type;
  v_系统来源   电子票据使用记录.系统来源%Type;
  j_Input      PLJson;
  j_Json       PLJson;
  j_Temp       PLJson;
  j_Temp1      PLJson;
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  --操作方式:0-换开;1-重新换开;2-作废票据;3-回收票据
  n_操作方式 := j_Json.Get_Number('oper_mode');
  n_Id       := j_Json.Get_Number('einvoice_id');
  v_操作员编号 := j_Json.Get_String('operator_code');
  v_操作员姓名 := j_Json.Get_String('operator_name');
  d_登记时间   := Nvl(To_Date(j_Json.Get_String('create_time'), 'yyyy-mm-dd hh24:mi:ss'), Sysdate);

  If Not j_Json.Exist('paperinv_info') Then
    Json_Out := zlJsonOut('无纸质票据信息。');
    Return;
  End If;
  Select Max(结算id) Into n_结算id From 电子票据使用记录 Where ID = n_Id;
  If Nvl(n_结算id, 0) = 0 Then
    Json_Out := zlJsonOut('传入的电子票据无效!');
    Return;
  End If;

  j_Temp := j_Json.Get_Pljson('paperinv_info');
  --应用场合:1-收费,2-预交(包含押金),3-结帐,4-挂号,5-就诊卡
  n_应用场合 := j_Temp.Get_Number('inv_occasion');
  --票种:1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡
  n_票种 := j_Temp.Get_Number('invoice_type');
  --是否红票:1-红票;0-非红票
  n_是否红票 := Nvl(j_Temp.Get_Number('inv_red'), 0);
  v_发票号   := j_Temp.Get_String('invoice_no');
  n_发票金额 := j_Temp.Get_Number('inv_total');
  n_领用id   := j_Temp.Get_Number('recv_id');

  --纸质票据处理
  --操作方式_In:0-换开;1-重新换开;2-作废票据;3-回收票据
  Zl_纸质票据使用_Update(n_应用场合, n_票种, n_结算id, n_Id, v_发票号, n_发票金额, n_领用id, v_操作员姓名, d_登记时间, n_操作方式, 0, n_是否红票);

  If j_Json.Exist('einvoce_info') Then
    --读取电子票据信息
    j_Temp1      := PLJson();
    j_Temp1      := j_Json.Get_Pljson('einvoce_info');
    v_开票点     := j_Temp1.Get_String('placeCode');
    v_系统来源   := j_Temp1.Get_String('sys_source');
    v_备注       := j_Temp1.Get_String('demo');
    n_原票据id   := j_Temp1.Get_Number('inv_oldid');
    n_冲销id     := j_Temp1.Get_Number('einvoice_id');
    v_代码       := j_Temp1.Get_String('einvoice_code');
    v_号码       := j_Temp1.Get_String('einvoice_no');
    v_检验码     := j_Temp1.Get_String('einvoice_random');
    v_凭证代码   := j_Temp1.Get_String('voucher_code');
    v_凭证号码   := j_Temp1.Get_String('voucher_no');
    v_凭证检验码 := j_Temp1.Get_String('voucher_random');
    v_生成时间   := j_Temp1.Get_String('happen_time');
    v_Url内网    := j_Temp1.Get_String('picture_url');
    v_Url外网    := j_Temp1.Get_String('picture_neturl');
    c_二维码     := j_Temp1.Get_Clob('qrcode');
  
    Zl_电子票据使用记录_Delete(n_冲销id, v_开票点, v_系统来源, v_生成时间, v_备注, v_操作员编号, v_操作员姓名, d_登记时间, n_原票据id, v_代码, v_号码, v_检验码,
                       v_凭证代码, v_凭证号码, v_凭证检验码, v_Url内网, v_Url外网);
    --更新二维码
    Insert Into 电子票据二维码 (使用记录id, 二维码) Values (n_冲销id, c_二维码);
    Json_Out := zlJsonOut('成功', 1);
    Return;
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Savepaperinvoice;
/
Create Or Replace Procedure Zl_Exsesvr_Getstarteinvoices
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:获取启用电子票据业务
  --入参：Json_In:NULL
  --     
  --出参: Json_Out,格式如下
  --output      
  --  code  C  1  应答码：0-失败；1-成功
  --  message  C  1  应答消息：成功时返回成功信息，失败时返回具体的错误信息
  --  data[]      启用站点列表
  --    occasion  N  1  场合:1-收费,2-预交,3-结帐,4-挂号
  --    client_name  C  1  站点名
  ---------------------------------------------------------------------------

  v_Output Varchar2(32767);
  c_Output Clob;
Begin
  --解析入参
  For c_卡类别 In (Select 场合, 站点 From 电子票据站点控制) Loop
  
    If Length(Nvl(v_Output, ' ')) > 30000 Then
      If c_Output Is Not Null Then
        c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
      Else
        c_Output := To_Clob(v_Output);
      End If;
      v_Output := Null;
    End If;
  
    zlJsonPutValue(v_Output, 'occasion', c_卡类别.场合, 1, 1);
    zlJsonPutValue(v_Output, 'client_name', c_卡类别.站点, 0, 2);
  End Loop;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"成功","data":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"成功","data":[' || v_Output || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getstarteinvoices;
/
Create Or Replace Procedure Zl_Exsesvr_Geteinvoicecode
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  -------------------------------------------------------------------------------------------------
  --功能：获取开票点编号
  --入参：json格式
  --input
  --   operator_id    N  1  操作员ID
  --   ssite          C  1  客户端
  --出参：json格式
  --Json_Out
  --  code            C  1  应答码：0-失败；1-成功
  --  message         C  1  应答消息： 成功时返回处方No，[数组] 失败时返回具体的错误信息
  --  einvoice_code   C  1  开票点编码
  --  is_exist        N  1  票据开票点对照是否存在数据:1-存在;0-不存在
  -------------------------------------------------------------------------------------------------
  n_操作员id   票据开票点对照.人员id%Type;
  v_客户端     票据开票点对照.客户端%Type;
  v_开票点编码 电子票据开票点.编码%Type;
  j_Input      PLJson;
  j_Json       PLJson;
  n_Count      Number(2);
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_操作员id := j_Json.Get_Number('operator_id');
  v_客户端   := j_Json.Get_String('ssite');

  Select Count(1) Into n_Count From 票据开票点对照 Where Rownum < 2;
  If n_Count = 0 Then
    Json_Out := '{"output":{"code":1,"message": "成功","einvoice_code":"","is_exist":0}}';
    Return;
  End If;

  --按收费员+客户端对码
  For r_开票点 In (Select b.编码 As 开票点编码
                From 票据开票点对照 A, 电子票据开票点 B
                Where a.开票点id = b.Id And Nvl(b.撤档时间, Sysdate + 1) >= Sysdate And a.人员id = n_操作员id And a.客户端 = v_客户端) Loop
    v_开票点编码 := r_开票点.开票点编码;
    Json_Out     := '{"output":{"code":1,"message": "成功","einvoice_code":"' || Nvl(v_开票点编码, '') || '","is_exist":1}}';
    Return;
  End Loop;

  --按收费员对码
  For r_开票点 In (Select Nvl(a.人员id, 0) As 人员id, Nvl(a.客户端, '-') As 客户端, b.编码 As 开票点编码
                From 票据开票点对照 A, 电子票据开票点 B
                Where a.开票点id = b.Id And Nvl(b.撤档时间, Sysdate + 1) >= Sysdate And a.人员id = n_操作员id And
                      Nvl(a.客户端, '-') = '-') Loop
    v_开票点编码 := r_开票点.开票点编码;
    Json_Out     := '{"output":{"code":1,"message": "成功","einvoice_code":"' || Nvl(v_开票点编码, '') || '","is_exist":1}}';
    Return;
  End Loop;

  --按客户端对码
  For r_开票点 In (Select Nvl(a.人员id, 0) As 人员id, Nvl(a.客户端, '-') As 客户端, b.编码 As 开票点编码
                From 票据开票点对照 A, 电子票据开票点 B
                Where a.开票点id = b.Id And Nvl(b.撤档时间, Sysdate + 1) >= Sysdate And Nvl(a.人员id, 0) = 0 And a.客户端 = v_客户端) Loop
    v_开票点编码 := r_开票点.开票点编码;
    Json_Out     := '{"output":{"code":1,"message": "成功","einvoice_code":"' || Nvl(v_开票点编码, '') || '","is_exist":1}}';
    Return;
  End Loop;

  Json_Out := '{"output":{"code":1,"message": "成功","einvoice_code":"' || Null || '","is_exist":1}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Geteinvoicecode;
/

Create Or Replace Procedure Zl_Exsesvr_Geteinvoicedata
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:根据结帐ID,获取有效的电子票据ID
  --入参：Json_In:
  --input
  --  fun_oper            N 1 操作类型：0-根据票种和结算id获取电子票据ID；1-根据电子票据ID获取 是否换开、纸质发票号、结算id
  --  blnc_id             N   结算ID(电子票据使用记录.结算id)
  --  inv_type            N   票种:1-收费, 2-预交, 3-结帐, 4-挂号;5-医疗发卡 
  --  einvoice_id         N   电子票据ID
  --出参: Json_Out,格式如下
  --output      
  --  code                C  1 应答码：0-失败；1-成功
  --  message             C  1 应答消息：成功时返回成功信息，失败时返回具体的错误信息
  --  einvoice_id         N    有效的电子票据ID(操作类型=0时返回)
  --  blnc_id             N    结算ID(操作类型=1时返回)
  --  is_turn             N    是否换开(操作类型=1时返回)
  --  inv_no              N    纸质票号(操作类型=1时返回)
  ---------------------------------------------------------------------------
  j_Input          PLJson;
  j_Json           PLJson;
  n_操作类型       Number(2);
  n_票种           电子票据使用记录.票种%Type;
  n_是否换开       电子票据使用记录.是否换开%Type;
  v_纸质发票号     电子票据使用记录.纸质发票号%Type;
  n_结算id         电子票据使用记录.结算id%Type;
  n_结算id_Out     电子票据使用记录.结算id%Type;
  n_电子票据id     电子票据使用记录.Id%Type;
  n_电子票据id_Out 电子票据使用记录.Id%Type;
Begin
  --解析入参

  j_Input    := PLJson(Json_In);
  j_Json     := j_Input.Get_Pljson('input');
  n_操作类型 := Nvl(Pljson_Ext.Get_Number(j_Json, 'fun_oper'), 0);

  If n_操作类型 = 0 Then
    --根据票种和结算id获取电子票据ID
    n_结算id := Pljson_Ext.Get_Number(j_Json, 'blnc_id');
    n_票种   := Pljson_Ext.Get_Number(j_Json, 'inv_type');
    If (Nvl(n_结算id, 0) = 0 Or Nvl(n_票种, 0) = 0) Then
      Json_Out := '{"output":{"code":0,"message": "失败,传入的结算id或场合为0"}}';
      Return;
    End If;
  
    Select Max(ID)
    Into n_电子票据id_Out
    From 电子票据使用记录
    Where 结算id = n_结算id And 票种 = n_票种 And 记录状态 = 1 And Nvl(原票据id, 0) = 0;
  
    Json_Out := '{"output":{"code":1,"message": "成功","einvoice_id":' || Nvl(n_电子票据id_Out, 0) || '}}';
  Else
    --根据电子票据ID获取 是否换开、纸质发票号、结算id
    n_电子票据id := Pljson_Ext.Get_Number(j_Json, 'einvoice_id');
    If Nvl(n_电子票据id, 0) = 0 Then
      Json_Out := '{"output":{"code":0,"message": "失败,传入的电子票据ID为0"}}';
      Return;
    End If;
  
    Select Max(是否换开), Max(纸质发票号), Max(结算id)
    Into n_是否换开, v_纸质发票号, n_结算id_Out
    From 电子票据使用记录
    Where ID = n_电子票据id;
  
    Json_Out := '{"output":{"code":1,"message": "成功","blnc_id":' || Nvl(n_结算id_Out, 0) || ',"is_turn":' ||
                Nvl(n_是否换开, 0) || ',"inv_no":"' || v_纸质发票号 || '"}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Geteinvoicedata;
/
Create Or Replace Procedure Zl_Exsesvr_Checkiseinvoice
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:根据传入的票种和结算ID,检查当前结算是否启用了有电子票据
  --入参：Json_In:
  --input
  --  blnc_id             N 1 结算ID(电子票据使用记录.id)
  --  inv_type            N 1 票种:1-收费, 2-预交, 3-结帐, 4-挂号;5-医疗发卡 
  --出参: Json_Out,格式如下
  --output      
  --  code                C  1 应答码：0-失败；1-成功
  --  message             C  1 应答消息：成功时返回成功信息，失败时返回具体的错误信息
  --  is_einvoice         N  1 是否启用电子票据:1-启用;0:未启用
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_票种     电子票据使用记录.票种%Type;
  n_结算id   电子票据使用记录.Id%Type;
  n_Einvoice Number(2);
Begin
  --解析入参

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_结算id := Pljson_Ext.Get_Number(j_Json, 'blnc_id');
  n_票种   := Pljson_Ext.Get_Number(j_Json, 'inv_type');

  If Nvl(n_结算id, 0) = 0 Or Nvl(n_票种, 0) = 0 Then
    Json_Out := '{"output":{"code":0,"message": "失败,传入的结算id或场合为0"}}';
    Return;
  End If;

  If n_票种 = 2 Then
    --预交记录
    Select Max(预交电子票据) Into n_Einvoice From 病人预交记录 Where Mod(记录性质, 10) = 1 And ID = n_结算id;
  Else
    Select Max(是否电子票据) Into n_Einvoice From 病人预交记录 Where 结帐id = n_结算id;
  End If;

  Json_Out := '{"output":{"code":1,"message": "成功","is_einvoice":' || Nvl(n_Einvoice, 0) || '}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkiseinvoice;
/
Create Or Replace Procedure Zl_Exsesvr_Geteinvoiceinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:获取电子票据信息
  --入参：Json_In:
  --input
  --err_id              N 1 异常ID
  --出参: Json_Out,格式如下
  --output      
  --code                C  1 应答码：0-失败；1-成功
  --message             C  1 应答消息：成功时返回成功信息，失败时返回具体的错误信息
  --记录标志=0 时 返回----------------------------------------
  --data
  --  input          
  --    balance_id        N  1  结算ID
  --    balance_delid     N     退款ID:退款开具红票时有效：目前只有预交款有效,填写的是退款预交ID
  --    einvoice_id       N  1  电子票据ID
  --    operator_code     C  1  操作员编号
  --    operator_name     C  1  操作员姓名
  --    create_time       C  1  登记时间:yyyy-mm-dd hh24:mi:ss
  --    pati_info         C     病人信息
  --      pati_id         N  1  病人ID
  --      pati_pageid     N     主页ID
  --      pati_name       C  1  姓名
  --      pati_sex        C  1  性别
  --      pati_age        C  1  年龄
  --      outpatient_num  C  1  门诊号
  --      inpatient_num   C  1  住院号
  --    einvoce_info      C     电子票据信息
  --      invoice_type    N  1  票种:1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡
  --      placeCode       C  1  开票点编码
  --      inv_total       N  1  开票金额
  --      inv_oldid       N    原票据ID
  --      sys_source      C  1  系统来源
  --      demo            C  1  备注
  --      einvoice_code   C  1  电子票据代码
  --      einvoice_no     C  1  电子票据号码
  --      einvoice_random C  1  电子校验码
  --      voucher_code    C  1  预交金凭证代码
  --      voucher_no      C  1  预交金凭证号码
  --      voucher_random  C  1  预交金凭证校验码
  --      happen_time     C  1  电子票据生成时间:yyyymmddhh24miss
  --      picture_url     C  1  电子票据H5页面URL
  --      picture_neturl  C  1  电子票据外网H5页面URL
  --      qrcode          C  1  电子票据二维码图片数据:该值已Base64编码，解析时需要Base64解码,图片格式为:PNG
  --记录标志=1 时 返回-----------------------------------------
  --data
  --  input         
  --    einvoice_id       N 1 电子票据ID
  --    operator_code     C 1 操作员编号
  --    operator_name     C 1 操作员姓名
  --    create_time       C 1 登记时间:yyyy-mm-dd hh24:mi:ss
  --    einvoce_info      C   电子票据信息
  --      placeCode       C 1 开票点编码
  --      sys_source      C 1 系统来源
  --      demo            C 1 备注
  --      inv_oldid       N   原票据ID
  --      einvoice_code   C 1 电子票据代码
  --      einvoice_no     C 1 电子票据号码
  --      einvoice_random C 1 电子校验码
  --      voucher_code    C 1 预交金凭证代码
  --      voucher_no      C 1 预交金凭证号码
  --      voucher_random  C 1 预交金凭证校验码
  --      happen_time     C 1 电子票据生成时间:yyyymmddhh24miss
  --      picture_url     C 1 电子票据H5页面URL
  --      picture_neturl  C 1 电子票据外网H5页面URL
  --      qrcode          C 1 电子票据二维码图片数据:该值已Base64编码，解析时需要Base64解码,图片格式为:PNG
  --记录标志=2,3 时 返回-------------------------------------------
  --data
  --  input         
  --    oper_mode         N 1 操作方式:0-换开;1-重新换开;2-作废票据;3-回收票据
  --    einvoice_id       N 1 电子票据ID
  --    operator_code     C 1 操作员编号
  --    operator_name     C 1 操作员姓名
  --    create_time       C 1 登记时间:yyyy-mm-dd hh24:mi:ss
  --    paperinv_info     C   纸质票据信息:存在多条时，请按操作顺序上传(避免数据错误)
  --      inv_occasion    N 1 应用场合:1-收费,2-预交(包含押金),3-结帐,4-挂号,5-就诊卡
  --      invoice_type    N 1 票种:1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡
  --      inv_red         N   是否红票:1-红票;0-非红票
  --      invoice_no      C 1 发票号
  --      inv_total       N 1 发票金额
  --      recv_id         N   领用id
  ---------------------------------------------------------------------------
  j_Input  PLJson;
  j_Json   PLJson;
  c_Output Clob;
  n_异常id 电子票据异常记录.Id%Type;

Begin
  --解析入参

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_异常id := Pljson_Ext.Get_Number(j_Json, 'err_id');

  If Nvl(n_异常id, 0) = 0 Then
    Json_Out := '{"output":{"code":0,"message": "失败,传入的异常id为0"}}';
    Return;
  End If;

  Begin
    Select 票据信息 Into c_Output From 电子票据异常记录 Where ID = n_异常id;
  Exception
    When Others Then
      Json_Out := '{"output":{"code":0,"message": "失败,根据传入的异常id未找到数据"}}';
      Return;
  End;

  If c_Output Is Null Then
    Json_Out := '{"output":{"code":0,"message": "失败,根据传入的异常id未找到数据"}}';
    Return;
  End If;

  Json_Out := '{"output":{"code":1,"message": "成功","data":' || c_Output || '}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Geteinvoiceinfo;
/
Create Or Replace Procedure Zl_Exsesvr_Drugwriteoff_Check
(
  Json_In  Clob,
  Json_Out Out Varchar2
) Is
  ---------------------------------------------------------------------------
  --功能：药品及卫材费用销帐审核检查
  --入参：Json_In:格式
  --input      药品费用销帐前检查
  --  part_ban_writeoffs    N  1  禁止部分销帐:0-允许;1-不允许部分销帐(含整张单据的部分或某笔的部份)
  --  fee_origin            N  1  费用来源:1-门诊，2-住院
  --  rcpdtl_list[]               本次销帐列表
  --    oper_type           N  1  操作类型:0-审核通过 1-审核不通过 2-审核拒绝 3-取消拒绝;
  --    rcpdtl_id           N  1  处方明细ID(费用ID)
  --    request_time        D     申请时间
  --    request_type        N     申请类别：缺省为1
  --    quantity            N  1  销帐数量：为零或null时,按费用ID申请数量直接销帐
  --    sended_num          N  1  已发数量
  --  pati_list[]                 病人信息
  --    pati_id             N     病人ID,为NULL或0时，表示整张单据
  --    fee_audit_status    N     费用审核标志:0或空-未审核;1-已审核或开始审核;2-完成审核,结合结帐权限
  --    si_inp_status       N     住院状态:0-正常住院;1-尚未入科;2-正在转科或正在转病区;3-已预出院
  --    catalog_date        C     病案编目日期：yyyy-mm-dd hh24:mi:ss
  --出参: Json_Out,格式如下
  --output
  --   code                          C 1 应答码：0-失败；1-成功
  --   message                       C 1 应答消息：失败时返回具体的错误信息
  --  tip_list[]  C  1  提示列表:主要是可能存在多个提示询问方式，所以用列表,禁止时，返回一条信息
  --    tip_mode  C  1  控制方式:1-提示询问;2-禁止
  --    tip_message  C  1  提示信息
  ---------------------------------------------------------------------------
  j_Input        PLJson;
  j_Json         PLJson;
  j_List         Pljson_List;
  j_Temp         PLJson;
  n_禁止部分销帐 Number(2);
  n_销帐方式     Number(2);
  n_操作类型     Number(2);
  n_申请类别     Number(2);
  d_申请时间     Date;
  n_销帐数量     病人费用销帐.数量%Type;
  n_已发数量     病人费用销帐.数量%Type;
  n_申请数量     病人费用销帐.数量%Type;
  n_审核部门id   病人费用销帐.审核部门id%Type;
  n_状态         Number(2);
  n_费用id       病人费用销帐.费用id%Type;
  n_Find         Number(2);
  n_已结单据操作 Number(3);
  v_Err_Msg      Varchar2(1000);

  l_Writeoffs  t_NumList2 := t_NumList2(); --费用本次销帐数量
  l_Excutes    t_NumList2 := t_NumList2(); --药品已发数量
  v_Patilist   Varchar2(32767);
  n_费用来源   Number;
  v_Json_In    Varchar2(32767);
  v_Itemlist   Varchar2(32767);
  v_Excutelist Varchar2(32767);

  v_已结序号 Varchar2(32767);
  n_Code     Number(2);
  Cursor c_费用信息 Is
    Select Distinct /*+cardinality(b,10)*/ a.Id As 费用id, a.收费类别, a.No, 序号, a.数次 As 销帐数量, a.数次 As 已发数量
    From 住院费用记录 A
    Where a.Id = 0;

  r_费用信息 c_费用信息%RowType;

  Type Ty_费用信息 Is Ref Cursor;
  c_销帐费用信息 Ty_费用信息; --动态游标变量

  v_No 门诊费用记录.No%Type;
Begin

  --取json节点的值（也是个json）
  j_Input        := PLJson(Json_In);
  j_Json         := j_Input.Get_Pljson('input');
  n_禁止部分销帐 := Nvl(j_Json.Get_Number('part_ban_writeoffs'), 0); --禁止部分销帐
  n_费用来源     := Nvl(j_Json.Get_Number('fee_origin'), 1);
  n_销帐方式     := 1; --药房使用，只有1:销帐方式：0-表示直接销帐;1-表示审核销帐(通过销帐申请-->销帐审核流程);2-表示转病区费用

  If Not j_Json.Exist('rcpdtl_list') Then
    Json_Out := zlJsonOut('未传入本次需要销帐的药品或卫材数据', 0);
    Return;
  End If;

  --0-允许 1-提示 2-禁止
  n_已结单据操作 := To_Number(Nvl(zl_GetSysParameter('已结帐单据操作'), '0'));

  If n_已结单据操作 = 1 Then
    n_已结单据操作 := 2;
  Elsif n_已结单据操作 = 2 Then
    n_已结单据操作 := 1;
  End If;

  --病人相关检查
  j_List     := Pljson_List();
  j_List     := j_Json.Get_Pljson_List('pati_list');
  v_Patilist := j_List.To_Char();
  v_Patilist := ',"pati_list":' || v_Patilist;

  j_List := Pljson_List();
  j_List := j_Json.Get_Pljson_List('rcpdtl_list');
  For J In 1 .. j_List.Count Loop
    j_Temp     := PLJson();
    j_Temp     := PLJson(j_List.Get(J));
    n_操作类型 := Nvl(j_Temp.Get_Number('oper_type'), 0);
    n_费用id   := Nvl(j_Temp.Get_Number('rcpdtl_id'), 0);
    n_申请类别 := Nvl(j_Temp.Get_Number('request_type'), 1);
    n_销帐数量 := Nvl(j_Temp.Get_Number('quantity'), 0);
    n_已发数量 := Nvl(j_Temp.Get_Number('sended_num'), 0);
    d_申请时间 := To_Date(j_Temp.Get_String('request_time'), 'yyyy-mm-dd hh24:mi:ss');
  
    --操作类型:0-审核通过 1-审核不通过 2-审核拒绝 3-取消拒绝;
    If d_申请时间 Is Not Null Then
      Begin
        --状态:0-申请,1-审核通过,2-审核未通过
        Select 状态, 数量, 审核部门id
        Into n_状态, n_申请数量, n_审核部门id
        From 病人费用销帐 A
        Where 费用id = n_费用id And 申请时间 = d_申请时间;
      Exception
        When Others Then
          n_状态 := -1;
      End;
    End If;
    If Nvl(n_操作类型, 0) In (2, 3) Then
      --重新审核拒绝 :
      --取消拒绝:主要是删除已经拒绝的审请
      If Nvl(n_状态, 0) = -1 Or d_申请时间 Is Null Then
        If d_申请时间 Is Null Then
          Json_Out := zlJsonOut('未传入指定的销帐申请数据，请检查!', 0);
        Else
          Json_Out := zlJsonOut('未发现申请时间为' || To_Char(d_申请时间, 'yyyy-mm-dd hh24:mi:ss') || '的费用申请记录，请检查!', 0);
        End If;
        Return;
      End If;
    
      If n_状态 <> 2 Then
        Begin
          If Nvl(n_费用来源, 0) = 1 Then
            Select '单据号:' || a.No || '中第' || a.序号 || '行(' || b.名称 || ')的' || Decode(a.收费类别, '4', '卫材', '药品') ||
                    ',不存在审核拒绝的记录,可能被他人取消。'
            Into v_Err_Msg
            From 门诊费用记录 A, 收费项目目录 B
            Where a.Id = n_费用id And a.收费细目id = b.Id(+);
          Else
            Select '单据号:' || a.No || '中第' || a.序号 || '行(' || b.名称 || ')的' || Decode(a.收费类别, '4', '卫材', '药品') ||
                    ',不存在审核拒绝的记录,可能被他人取消。'
            Into v_Err_Msg
            From 住院费用记录 A, 收费项目目录 B
            Where a.Id = n_费用id And a.收费细目id = b.Id(+);
          End If;
        Exception
          When Others Then
            v_Err_Msg := Null;
        End;
        If v_Err_Msg Is Null Then
          Json_Out := zlJsonOut('未找到费用ID=' || n_费用id || '的费用记录，请检查传入的费用ID是否正确!', 0);
          Return;
        End If;
        v_Err_Msg := '{"tip_mode":2,"tip_message":"' || zlJsonStr(v_Err_Msg) || '"}';
        Json_Out  := '{"output":{"code":1,"message":"成功","tip_list":[' || v_Err_Msg || ']}}';
        Return;
      End If;
    End If;
    --操作类型:0-审核通过 1-审核不通过 2-审核拒绝 3-取消拒绝;
    If Nvl(n_操作类型, 0) In (0, 2) Then
      --n_销帐方式:0-表示直接销帐;1-表示审核销帐(通过销帐申请-->销帐审核流程);2-表示转病区费用
      If Nvl(n_销帐方式, 0) = 1 Then
        If Nvl(n_申请数量, 0) < Nvl(n_销帐数量, 0) Then
          Begin
            If Nvl(n_费用来源, 0) = 1 Then
              Select '单据号:' || a.No || '中第' || a.序号 || '行(' || b.名称 || ')的' || Decode(a.收费类别, '4', '卫材', '药品') ||
                      '的本次销帐数量(' || Nvl(n_销帐数量, 0) || ')大于了本次申请数量(' || Nvl(n_申请数量, 0) || ')。'
              Into v_Err_Msg
              From 门诊费用记录 A, 收费项目目录 B
              Where a.Id = n_费用id And a.收费细目id = b.Id(+);
            Else
              Select '单据号:' || a.No || '中第' || a.序号 || '行(' || b.名称 || ')的' || Decode(a.收费类别, '4', '卫材', '药品') ||
                      '的本次销帐数量(' || Nvl(n_销帐数量, 0) || ')大于了本次申请数量(' || Nvl(n_申请数量, 0) || ')。'
              Into v_Err_Msg
              From 住院费用记录 A, 收费项目目录 B
              Where a.Id = n_费用id And a.收费细目id = b.Id(+);
            End If;
          Exception
            When Others Then
              v_Err_Msg := Null;
          End;
          If v_Err_Msg Is Null Then
            Json_Out := zlJsonOut('未找到费用ID=' || n_费用id || '的费用记录，请检查传入的费用ID是否正确!', 0);
            Return;
          End If;
          v_Err_Msg := '{"tip_mode":2,"tip_message":"' || zlJsonStr(v_Err_Msg) || '"}';
          Json_Out  := '{"output":{"code":1,"message":"成功","tip_list":[' || v_Err_Msg || ']}}';
          Return;
        End If;
      End If;
    
      n_Find := 0;
      For I In 1 .. l_Writeoffs.Count Loop
        If l_Writeoffs(I).C1 = n_费用id Then
          l_Writeoffs(I).C2 := n_销帐数量 + l_Writeoffs(I).C2;
          l_Excutes(I).C2 := Nvl(n_已发数量, 0) + l_Excutes(I).C2;
          n_Find := 1;
          Exit;
        End If;
      End Loop;
      If Nvl(n_Find, 0) = 0 Then
        l_Writeoffs.Extend;
        l_Writeoffs(l_Writeoffs.Count) := t_NumObj2(n_费用id, n_销帐数量);
        l_Excutes.Extend;
        l_Excutes(l_Excutes.Count) := t_NumObj2(n_费用id, n_已发数量);
      End If;
    
    End If;
  End Loop;

  If l_Writeoffs.Count = 0 Then
    --主要是取消操作
    Json_Out := zlJsonOut('成功', 1);
    Return;
  End If;

  If Nvl(n_费用来源, 0) = 1 Or n_费用来源 Is Null Then
    --门诊记帐
    Open c_销帐费用信息 For
      Select Distinct /*+cardinality(b,10)*/ a.Id As 费用id, a.收费类别, a.No, a.序号, b.C2 As 销帐数量, c.C2 As 已发数量
      From 门诊费用记录 A, Table(l_Writeoffs) B, Table(l_Excutes) C
      Where a.Id = b.C1 And a.Id = c.C1(+)
      Order By a.No, a.序号;
  Else
    Open c_销帐费用信息 For
      Select Distinct /*+cardinality(b,10)*/ a.Id As 费用id, a.收费类别, a.No, a.序号, b.C2 As 销帐数量, c.C2 As 已发数量
      From 住院费用记录 A, Table(l_Writeoffs) B, Table(l_Excutes) C
      Where a.Id = b.C1 And a.Id = c.C1(+)
      Order By a.No, a.序号;
  End If;

  v_No         := Null;
  v_Json_In    := Null;
  v_Itemlist   := Null;
  v_Excutelist := Null;
  Loop
    Fetch c_销帐费用信息
      Into r_费用信息;
    Exit When c_销帐费用信息%NotFound;
  
    If Nvl(v_No, '.') <> r_费用信息.No Then
    
      If v_Itemlist Is Not Null Then
      
        v_Itemlist := ',"item_list":[' || v_Itemlist || ']';
        If Not v_Excutelist Is Null Then
          --本次已发数量列表
          v_Excutelist := ',"excute_list":[' || v_Excutelist || ']';
        End If;
      
        v_Json_In := '"fee_no":"' || v_No || '"';
        v_Json_In := v_Json_In || ',"fee_bill_type":2';
        v_Json_In := v_Json_In || ',"balance_ban_writeoffs":' || Nvl(n_已结单据操作, 0);
        v_Json_In := v_Json_In || ',"part_ban_writeoffs":' || Nvl(n_禁止部分销帐, 0);
        v_Json_In := v_Json_In || ',"oper_type":' || Nvl(n_销帐方式, 0);
      
        --本次销帐列表
        v_Json_In := v_Json_In || v_Itemlist;
        --本次已发数量列表
        v_Json_In := v_Json_In || Nvl(v_Excutelist, '');
      
        v_Json_In := v_Json_In || Nvl(v_Patilist, '');
        v_Json_In := '{"input":{' || v_Json_In || '}}';
      
        If Nvl(n_费用来源, 1) = 1 Then
          Zl_门诊记帐记录_Delete_Check(v_Json_In, Json_Out);
        Else
          Zl_住院记帐记录_Delete_Check(v_Json_In, Json_Out);
        End If;
        --
        j_Input := PLJson(Json_Out);
        j_Json  := j_Input.Get_Pljson('output');
      
        n_Code := j_Json.Get_Number('code');
        If n_Code = 0 Then
          Json_Out := '{"output":{"code":1,"message":"成功","tip_list":[{"tip_mode":2,"tip_message":"' ||
                      zlJsonStr(j_Json.Get_String('message')) || '"}]}}';
          Return;
        End If;
        If j_Json.Exist('balance_serials') Then
          If v_已结序号 Is Not Null Then
            v_已结序号 := v_已结序号 || Chr(13);
          End If;
          v_已结序号 := v_已结序号 || v_No || ':' || j_Json.Get_String('balance_serials');
        End If;
      End If;
      v_No         := r_费用信息.No;
      v_Itemlist   := Null;
      v_Excutelist := Null;
    End If;
  
    --卫材及药品检查
    If Instr(',4,5,6,7,', ',' || r_费用信息.收费类别 || ',') = 0 Then
      v_Err_Msg := '在单据:' || v_No || '中的第' || r_费用信息.序号 || '行中存在非药品及卫材的收费项目';
      Json_Out  := '{"output":{"code":1,"message":"成功","tip_list":[{"tip_mode":2,"tip_message":"' ||
                   zlJsonStr(v_Err_Msg) || '"}]}}';
    
      Return;
    End If;
    --构建明细数据
    If v_Itemlist Is Not Null Then
      v_Itemlist := v_Itemlist || ',';
    End If;
    v_Itemlist := Nvl(v_Itemlist, '') || '{"serial_num":' || Nvl(r_费用信息.序号, 0);
    v_Itemlist := v_Itemlist || ',"quantity":' || zlJsonStr(r_费用信息.销帐数量, 1) || '}';
  
    --构建已发为数据
    If v_Excutelist Is Not Null Then
      v_Excutelist := v_Excutelist || ',';
    End If;
    v_Excutelist := Nvl(v_Excutelist, '') || '{"fee_id":' || Nvl(r_费用信息.费用id, 0);
    v_Excutelist := v_Excutelist || ',"sended_num":' || zlJsonStr(r_费用信息.已发数量, 1) || '}';
  End Loop;

  If v_Itemlist Is Not Null Then
  
    v_Itemlist := ',"item_list":[' || v_Itemlist || ']';
    If Not v_Excutelist Is Null Then
      --本次已发数量列表
      v_Excutelist := ',"excute_list":[' || v_Excutelist || ']';
    End If;
  
    v_Json_In := '"fee_no":"' || v_No || '"';
    v_Json_In := v_Json_In || ',"fee_bill_type":2';
    v_Json_In := v_Json_In || ',"balance_ban_writeoffs":' || Nvl(n_已结单据操作, 0);
    v_Json_In := v_Json_In || ',"part_ban_writeoffs":' || Nvl(n_禁止部分销帐, 0);
    v_Json_In := v_Json_In || ',"oper_type":' || Nvl(n_销帐方式, 0);
  
    --本次销帐列表
    v_Json_In := v_Json_In || v_Itemlist;
    --本次已发数量列表
    v_Json_In := v_Json_In || Nvl(v_Excutelist, '');
    v_Json_In := v_Json_In || Nvl(v_Patilist, '');
    v_Json_In := '{"input":{' || v_Json_In || '}}';
    If Nvl(n_费用来源, 1) = 1 Then
      Zl_门诊记帐记录_Delete_Check(v_Json_In, Json_Out);
    Else
      Zl_住院记帐记录_Delete_Check(v_Json_In, Json_Out);
    End If;
    j_Input := PLJson(Json_Out);
    j_Json  := j_Input.Get_Pljson('output');
  
    n_Code := j_Json.Get_Number('code');
    If n_Code = 0 Then
      Json_Out := '{"output":{"code":1,"message":"成功","tip_list":[{"tip_mode":2,"tip_message":"' ||
                  zlJsonStr(j_Json.Get_String('message')) || '"}]}}';
      Return;
    End If;
    If j_Json.Exist('balance_serials') Then
      If v_已结序号 Is Not Null Then
        v_已结序号 := v_已结序号 || Chr(13);
      End If;
      v_已结序号 := v_已结序号 || v_No || ':' || j_Json.Get_String('balance_serials');
    End If;
  End If;

  If v_已结序号 Is Not Null Then
    --返回：询问方式
    Json_Out := '{"output":{"code":1,"message":"成功","tip_list":[{"tip_mode":1,"tip_message":"' ||
                zlJsonStr('以下单据已经结帐:' || Chr(13) || v_已结序号) || '"}]}}';
    Return;
  End If;
  Json_Out := zlJsonOut('成功', 1);
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Drugwriteoff_Check;
/

Create Or Replace Procedure Zl_Exsesvr_Drugwriteoff
(
  Json_In  Clob,
  Json_Out Out Varchar2
) Is
  ---------------------------------------------------------------------------
  --功能：药品及卫材费用销帐(包含审核通过、重审拒绝，取消拒绝)
  --入参：Json_In:格式
  --input     
  --  fee_origin            N  1  费用来源（1-门诊，2-住院）
  --  operator_code         C  1  操作员编码
  --  operator_name         C  1  操作员姓名 
  --  operator_time         C     操作时间:yyyy-mm-dd hh24:mi:ss
  --  rcpdtl_list                 [数组]每个处方明细信息
  --    rcpdtl_id           N  1  处方明细id(费用id)
  --    request_time        D  1  申请时间
  --    oper_type           N  1  操作类型:0-审核通过;1-审核不通过 2-审核拒绝 3-取消拒绝;
  --    request_type        N  1  申请类别（默认传1）
  --    quantity            N  1  销帐数量
  --    sended_num          N  1  已发数量

  --出参: Json_Out,格式如下
  --output
  --   code                          C 1 应答码：0-失败；1-成功
  --   message                       C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  Cursor c_费用信息 Is
    Select Distinct /*+cardinality(b,10)*/ a.No, 序号, a.收费类别, a.数次 As 剩余数量, a.数次 As 销帐数量, a.数次 As 已发数量
    From 住院费用记录 A
    Where a.Id = 0;

  r_费用信息 c_费用信息%RowType;

  Type Ty_费用信息 Is Ref Cursor;
  c_销帐费用信息 Ty_费用信息; --动态游标变量

  j_Input PLJson;
  j_Json  PLJson;
  j_List  Pljson_List;
  j_Temp  PLJson;

  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  n_操作类型 Number(2);
  n_申请类别 Number(2);
  d_申请时间 Date;
  n_Find     Number(2);

  l_Writeoffs t_NumList2 := t_NumList2(); --费用本次销帐数量
  l_Excutes   t_NumList2 := t_NumList2(); --药品已发数量

  n_销帐方式 Number(2);
  v_序号     Varchar2(32767);
  n_Temp     Number(2);
  n_费用来源 Number(2);

  v_No         门诊费用记录.No%Type;
  v_操作员编号 门诊费用记录.操作员编号%Type;
  v_操作员姓名 门诊费用记录.操作员姓名%Type;
  n_销帐数量   病人费用销帐.数量%Type;
  n_已发数量   病人费用销帐.数量%Type;
  n_费用id     病人费用销帐.费用id%Type;
  d_操作时间   Date;
  n_执行状态   Number(2);
  n_Count      Number(2);
Begin

  --取json节点的值（也是个json）
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_操作员编号 := j_Json.Get_String('operator_code');
  v_操作员姓名 := j_Json.Get_String('operator_name');
  n_费用来源   := Nvl(j_Json.Get_Number('fee_origin'), 1);

  If j_Json.Exist('operator_time') Then
    d_操作时间 := To_Date(j_Json.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');
  Else
    d_操作时间 := Sysdate;
  End If;
  n_销帐方式 := 1; --药房使用，只有1:销帐方式：0-表示直接销帐;1-表示审核销帐(通过销帐申请-->销帐审核流程);2-表示转病区费用

  If Not j_Json.Exist('rcpdtl_list') Then
    Json_Out := zlJsonOut('未传入本次需要销帐的药品或卫材数据', 0);
    Return;
  End If;

  j_List := Pljson_List();
  j_List := j_Json.Get_Pljson_List('rcpdtl_list');
  For J In 1 .. j_List.Count Loop
    j_Temp := PLJson();
    j_Temp := PLJson(j_List.Get(J));
  
    n_操作类型 := Nvl(j_Temp.Get_Number('oper_type'), 0);
    n_费用id   := Nvl(j_Temp.Get_Number('rcpdtl_id'), 0);
    d_申请时间 := To_Date(j_Temp.Get_String('request_time'), 'yyyy-mm-dd hh24:mi:ss');
    n_申请类别 := Nvl(j_Temp.Get_Number('request_type'), 1);
  
    n_销帐数量 := Nvl(j_Temp.Get_Number('quantity'), 0);
    n_已发数量 := Nvl(j_Temp.Get_Number('sended_num'), 0);
  
    If Nvl(n_操作类型, 0) In (2, 3) Then
      --操作类型:0-审核通过;1-审核不通过 2-审核拒绝 3-取消拒绝;
      If Nvl(n_操作类型, 0) = 3 Then
        n_Temp := 1;
      Else
        n_Temp := 0;
      End If;
      -- n_Temp:0-审核拒绝 1-取消拒绝
      Zl_病人费用销帐_Cancel_s(n_费用id, d_申请时间, v_操作员姓名, d_操作时间, n_Temp, n_申请类别);
    Else
      If Nvl(n_操作类型, 0) = 1 Then
        --操作类型:0-审核通过; 1-审核不通过
        n_Temp := 2;
      Elsif Nvl(n_操作类型, 0) = 0 Then
        n_Temp := 1;
      Else
        v_Err_Msg := '操作状态传入错误(费用ID=' || n_费用id || ')，只能为四种状态:0-审核通过;1-审核不通过;2-审核拒绝;3-取消拒绝';
        Raise Err_Item;
      End If;
      Select Count(1)
      Into n_Count
      From 病人费用销帐
      Where 费用id = n_费用id And 申请时间 = d_申请时间 And 申请类别 = n_申请类别 And 状态 = n_Temp;
      If n_Count <> 0 Then
        v_Err_Msg := '该单据(费用ID=' || n_费用id || ')已审核，禁止重新审核';
        Raise Err_Item;
      End If;
      Zl_病人费用销帐_Audit_s(n_费用id, d_申请时间, v_操作员姓名, d_操作时间, n_Temp, n_申请类别);
    End If;
  
    --操作类型:0-审核通过;1-审核不通过 2-审核拒绝 3-取消拒绝;
    If Nvl(n_操作类型, 0) In (0, 2) Then
      --进行销帐处理
      n_Find := 0;
      For I In 1 .. l_Writeoffs.Count Loop
        If l_Writeoffs(I).C1 = n_费用id Then
          l_Writeoffs(I).C2 := n_销帐数量 + l_Writeoffs(I).C2;
          l_Excutes(I).C2 := n_已发数量;
          n_Find := 1;
          Exit;
        End If;
      End Loop;
      If Nvl(n_Find, 0) = 0 Then
        l_Writeoffs.Extend;
        l_Writeoffs(l_Writeoffs.Count) := t_NumObj2(n_费用id, n_销帐数量);
        l_Excutes.Extend;
        l_Excutes(l_Excutes.Count) := t_NumObj2(n_费用id, n_已发数量);
      End If;
    End If;
  End Loop;

  If l_Writeoffs.Count = 0 Then
    Json_Out := zlJsonOut('成功', 1);
    Return;
  
  End If;
  If Nvl(n_费用来源, 0) = 1 Or n_费用来源 Is Null Then
    --门诊记帐
    Open c_销帐费用信息 For
      Select a.No, a.序号, a.收费类别, Sum(Nvl(a.付数, 1) * Nvl(a.数次, 0)) As 剩余数量, Max(b.销帐数量) As 销帐数量, Max(b.已发数量) As 已发数量
      From 门诊费用记录 A,
           (Select Distinct /*+cardinality(b,10)*/ a.No, a.序号, b.C2 As 销帐数量, c.C2 As 已发数量
             From 门诊费用记录 A, Table(l_Writeoffs) B, Table(l_Excutes) C
             Where a.Id = b.C1 And a.Id = c.C1(+)) B
      Where a.No = b.No And a.序号 = b.序号 And a.记录性质 = 2 And 价格父号 Is Null
      Group By a.No, a.序号, a.收费类别
      Order By a.No, a.序号;
  Else
    Open c_销帐费用信息 For
      Select a.No, a.序号, a.收费类别, Sum(Nvl(a.付数, 1) * Nvl(a.数次, 0)) As 剩余数量, Max(b.销帐数量) As 销帐数量, Max(b.已发数量) As 已发数量
      From 住院费用记录 A,
           (Select Distinct /*+cardinality(b,10)*/ a.No, a.序号, b.C2 As 销帐数量, c.C2 As 已发数量
             From 住院费用记录 A, Table(l_Writeoffs) B, Table(l_Excutes) C
             Where a.Id = b.C1 And a.Id = c.C1(+)) B
      Where a.No = b.No And a.序号 = b.序号 And a.记录性质 = 2 And 价格父号 Is Null
      Group By a.No, a.序号, a.收费类别
      Order By a.No, a.序号;
  End If;

  v_No   := Null;
  v_序号 := Null;
  Loop
    Fetch c_销帐费用信息
      Into r_费用信息;
    Exit When c_销帐费用信息%NotFound;
  
    If Nvl(v_No, '.') <> r_费用信息.No Then
      If v_序号 Is Not Null Then
        If n_费用来源 = 1 Then
          Zl_门诊记帐记录_Delete_s(v_No, v_序号, v_操作员编号, v_操作员姓名, d_操作时间);
        Elsif n_费用来源 = 2 Then
          Zl_住院记帐记录_Delete_s(v_No, v_序号, v_操作员编号, v_操作员姓名, 2, 1, d_操作时间);
        End If;
      End If;
      v_No   := r_费用信息.No;
      v_序号 := Null;
    End If;
  
    n_执行状态 := 0;
    If Nvl(r_费用信息.已发数量, 0) <> 0 Then
      n_执行状态 := 2;
    End If;
    If Nvl(r_费用信息.已发数量, 0) = Nvl(r_费用信息.剩余数量, 0) - Nvl(r_费用信息.销帐数量, 0) Then
      n_执行状态 := 1;
    End If;
    If Instr(',4,5,6,7,', ',' || r_费用信息.收费类别 || ',') = 0 Then
      v_Err_Msg := '在单据:' || v_No || '中的第' || r_费用信息.序号 || '行中存在非药品及卫材的收费项目';
      Raise Err_Item;
    End If;
    If v_序号 Is Not Null Then
      v_序号 := v_序号 || ',';
    End If;
    v_序号 := v_序号 || r_费用信息.序号 || ':' || r_费用信息.销帐数量 || ':' || n_执行状态;
    --序号1:数量1:执行状态1,序号2:数量2:执行状态2,...序号n:数量n:执行状态n  如:"1:2:1,2:10:1,3:2:1"
  End Loop;
  If v_序号 Is Not Null Then
    If n_费用来源 = 1 Then
      Zl_门诊记帐记录_Delete_s(v_No, v_序号, v_操作员编号, v_操作员姓名, d_操作时间);
    Elsif n_费用来源 = 2 Then
      Zl_住院记帐记录_Delete_s(v_No, v_序号, v_操作员编号, v_操作员姓名, 2, n_销帐方式, d_操作时间);
    End If;
  End If;
  Json_Out := zlJsonOut('成功', 1);
Exception
  When Err_Item Then
    Json_Out := zlJsonOut(v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Drugwriteoff;
/
Create Or Replace Procedure Zl_Exsesvr_Updateexeinfo
(
  Json_In  Clob,
  Json_Out Out Varchar2
) Is
  ---------------------------------------------------------------------------
  --功能：更新费用执行部门、执行人、发药窗口及执行状态等信息
  --入参：Json_In:格式
  --input     
  --  fee_origin            N  1  费用来源(默认=2：1-门诊费用，2-住院费用)
  --  operator_code         C     操作员编码
  --  operator_name         C     操作员姓名
  --  operator_time         C     操作时间
  --  item_list                   按列表更新执行相关信息，传入列表时同时需要传入fee_origin
  --    fee_id              N     费用id,不传入时以费用单据号、医嘱id、收费细目id为准         
  --    fee_no              C     费用单据号
  --    advice_id           N     医嘱id(已经能确定是收费单还是记帐单了，所以不用再传入单据性质)
  --    fee_item_id         N     收费细目id     
  --                              注意：fee_id或(fee_no、advice_id、fee_item_id)必传其中一个条件.
  --    exe_nums            N  1  已执行数量:为0表示，未执行
  --    exe_people          C     执行人:部分执行或完全执行时，需要传入，不传入时，以operator_name为准
  --    exe_time            D     执行时间:yyyy-mm-dd hh24:mi:ss,:部分执行或完全执行时，需要传入，不传入时，以"create_time"为准
  --    pharmacy_window     C     发药窗口:药品及卫材有效,无此接点，不会更新发药窗口
  --  deptchange_list       C  1  执行科室变更信息列表
  --    fee_id              N  1  费用id
  --    exe_old_deptid      N     原执行科室ID 
  --    exe_deptid          N  1  执行部门id
  --  delrcp_list           C     [数组]自动销账时,需要同步销帐
  --    rcp_no              C  1  处方no
  --    serial_nums         C  1  格式: 序号1:数量:执行状态1,序号2:数量2:执行状态2,...
  --    operator_status     N     操作状态：0-表示直接销帐;1-表示审核销帐(通过销帐申请-->销帐审核流程);2-表示转病区费用
  --出参: Json_Out,格式如下
  --output
  --   code                 C  1  应答码：0-失败；1-成功
  --   message              C  1  应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  Cursor c_费用 Is
    Select a.No, a.序号, Mod(a.记录性质, 10) As 记录性质, a.收费类别, a.数次 As 剩余数量
    From 住院费用记录 A
    Where a.Id = 0;

  r_费用信息 c_费用%RowType;

  Cursor c_科室变更 Is
    Select a.病人id, a.No, Nvl(a.价格父号, a.序号) As 序号, Mod(a.记录性质, 10) As 记录性质, a.主页id, a.病人病区id, a.病人科室id, a.开单部门id,
           a.执行部门id, a.收入项目id, a.门诊标志, Sum(Nvl(a.实收金额, 0) - Nvl(a.结帐金额, 0)) As 未结金额, Min(a.登记时间) As 登记时间
    From 住院费用记录 A
    Where ID = 0
    Order By a.收费细目id;

  r_科室变更 c_科室变更%RowType;

  Type Ty_费用信息 Is Ref Cursor;
  c_费用信息 Ty_费用信息; --动态游标变量

  j_Input PLJson;
  j_Json  PLJson;
  j_List  Pljson_List;
  j_Temp  PLJson;

  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  v_序号 Varchar2(32767);

  v_No         门诊费用记录.No%Type;
  v_操作员编号 门诊费用记录.操作员编号%Type;
  v_操作员姓名 门诊费用记录.操作员姓名%Type;
  v_执行人     门诊费用记录.执行人%Type;
  d_执行时间   门诊费用记录.执行时间%Type;
  n_执行数量   门诊费用记录.数次%Type;
  n_返回值1    门诊费用记录.结帐金额%Type;
  n_返回值     门诊费用记录.结帐金额%Type;
  n_费用id     病人费用销帐.费用id%Type;
  v_单据号     门诊费用记录.No%Type;
  n_医嘱id     门诊费用记录.医嘱序号%Type;
  n_收费细目id 门诊费用记录.收费细目id%Type;

  v_发药窗口 门诊费用记录.发药窗口%Type;

  n_原执行部门id 门诊费用记录.执行部门id%Type;

  n_当前执行部门id 门诊费用记录.执行部门id%Type;

  d_操作时间     Date;
  n_费用来源     Number(2);
  n_执行状态     Number(2);
  n_附加标志     Number(2);
  n_操作状态     Number(2);
  n_更新执行信息 Number(2);
Begin

  --取json节点的值（也是个json）
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_操作员编号 := j_Json.Get_String('operator_code');
  v_操作员姓名 := j_Json.Get_String('operator_name');
  n_费用来源   := Nvl(j_Json.Get_Number('fee_origin'), 2);
  If j_Json.Exist('operator_time') Then
    d_操作时间 := To_Date(j_Json.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');
  Else
    d_操作时间 := Sysdate;
  End If;

  If j_Json.Exist('item_list') Then
    j_List := Pljson_List();
    j_List := j_Json.Get_Pljson_List('item_list');
    For J In 1 .. j_List.Count Loop
      j_Temp := PLJson();
      j_Temp := PLJson(j_List.Get(J));
    
      n_费用id     := Nvl(j_Temp.Get_Number('fee_id'), 0);
      n_执行数量   := Nvl(j_Temp.Get_Number('exe_nums'), 0);
      v_单据号     := Nvl(j_Temp.Get_String('fee_no'), '-');
      n_医嘱id     := Nvl(j_Temp.Get_Number('advice_id'), 0);
      n_收费细目id := Nvl(j_Temp.Get_Number('fee_item_id'), 0);
    
      If n_费用id = 0 Then
        If v_单据号 = '-' Or n_医嘱id = 0 Or n_收费细目id = 0 Then
          v_Err_Msg := '入参节点fee_id或(fee_no、advice_id、fee_item_id)必须有一个不为空，请检查!';
          Raise Err_Item;
        End If;
      End If;
    
      If j_Temp.Exist('exe_people') Then
        v_执行人       := j_Temp.Get_String('exe_people');
        d_执行时间     := To_Date(j_Temp.Get_String('exe_time'), 'yyyy-mm-dd hh24:mi:ss');
        n_更新执行信息 := 1;
      End If;
    
      v_发药窗口 := Null;
      If j_Temp.Exist('pharmacy_window') Then
        --传入节点，则按该节点更新
        v_发药窗口 := Nvl(j_Temp.Get_String('pharmacy_window'), ' ');
      End If;
    
      If Nvl(n_费用来源, 1) = 1 Then
        --门诊费用
        If n_费用id <> 0 Then
          Open c_费用信息 For
            Select Distinct /*+cardinality(b,10)*/ a.No, a.序号, Mod(a.记录性质, 10) As 记录性质, a.收费类别,
                            Sum(Nvl(a.付数, 1) * a.数次) As 剩余数量
            From 门诊费用记录 A, (Select NO, 序号, Mod(记录性质, 10) As 记录性质 From 门诊费用记录 Where ID = n_费用id) B
            Where a.No = b.No And a.序号 = b.序号 And Mod(a.记录性质, 10) = b.记录性质 And a.记录性质 <> 12
            Group By a.No, a.序号, Mod(a.记录性质, 10), a.收费类别;
        Else
          Open c_费用信息 For
            Select a.No, a.序号, Mod(a.记录性质, 10) As 记录性质, a.收费类别, Sum(Nvl(a.付数, 1) * a.数次) As 剩余数量
            From 门诊费用记录 A
            Where a.No = v_单据号 And a.医嘱序号 + 0 = n_医嘱id And a.收费细目id + 0 = n_收费细目id And a.价格父号 Is Null And a.记录性质 <> 12
            Group By a.No, a.序号, Mod(a.记录性质, 10), a.收费类别;
        End If;
      Else
        --住院费用
        If n_费用id <> 0 Then
          Open c_费用信息 For
            Select Distinct /*+cardinality(b,10)*/ a.No, a.序号, Mod(a.记录性质, 10) As 记录性质, a.收费类别,
                            Sum(Nvl(a.付数, 1) * a.数次) As 剩余数量
            From 住院费用记录 A, (Select NO, 序号, Mod(记录性质, 10) As 记录性质 From 住院费用记录 Where ID = n_费用id) B
            Where a.No = b.No And a.序号 = b.序号 And a.记录性质 = b.记录性质
            Group By a.No, a.序号, Mod(a.记录性质, 10), a.收费类别;
        Else
          Open c_费用信息 For
            Select a.No, a.序号, Mod(a.记录性质, 10) As 记录性质, a.收费类别, Sum(Nvl(a.付数, 1) * a.数次) As 剩余数量
            From 住院费用记录 A
            Where a.No = v_单据号 And a.医嘱序号 + 0 = n_医嘱id And a.收费细目id + 0 = n_收费细目id And a.价格父号 Is Null And a.记录性质 <> 12
            Group By a.No, a.序号, Mod(a.记录性质, 10), a.收费类别;
        End If;
      End If;
    
      Fetch c_费用信息
        Into r_费用信息;
    
      If c_费用信息%NotFound Then
        If n_费用id <> 0 Then
          v_Err_Msg := '未找到对应的费用记录(费用ID=' || n_费用id || ')';
        Else
          v_Err_Msg := '未找到对应的费用记录(单据号=' || v_单据号 || ')';
        End If;
        Raise Err_Item;
      End If;
    
      n_执行状态 := 0;
      If Nvl(n_执行数量, 0) <> 0 Then
        n_执行状态 := 2;
        If Nvl(r_费用信息.剩余数量, 0) = Nvl(n_执行数量, 0) Then
          n_执行状态 := 1;
        End If;
        If Abs(Nvl(r_费用信息.剩余数量, 0)) < Abs(Nvl(n_执行数量, 0)) Then
          v_Err_Msg := '单据号为' || r_费用信息.No || '中的第' || r_费用信息.序号 || '行的剩余数量于小了已执行数量，请检查!';
          Raise Err_Item;
        End If;
      End If;
    
      n_附加标志 := Null;
      If Instr(',5,6,7,', ',' || r_费用信息.收费类别 || ',') > 0 Then
        n_附加标志 := 0;
        If Nvl(v_执行人, '-') <> '-' Then
          n_附加标志 := 1;
        End If;
      End If;
    
      If Nvl(n_费用来源, 1) = 1 Then
        Update 门诊费用记录
        Set 执行状态 = n_执行状态, 附加标志 = Nvl(n_附加标志, 附加标志), 执行人 = Decode(n_更新执行信息, 1, v_执行人, 执行人),
            执行时间 = Decode(n_更新执行信息, 1, d_执行时间, 执行时间), 发药窗口 = LTrim(Nvl(v_发药窗口, 发药窗口))
        Where NO = r_费用信息.No And Nvl(价格父号, 序号) = r_费用信息.序号 And 记录状态 In (0, 1, 3) And 记录性质 = r_费用信息.记录性质;
      Else
        Update 住院费用记录
        Set 执行状态 = n_执行状态, 附加标志 = Nvl(n_附加标志, 附加标志), 执行人 = Decode(n_更新执行信息, 1, v_执行人, 执行人),
            执行时间 = Decode(n_更新执行信息, 1, d_执行时间, 执行时间)
        Where NO = r_费用信息.No And Nvl(价格父号, 序号) = r_费用信息.序号 And 记录状态 In (0, 1, 3) And 记录性质 = r_费用信息.记录性质;
      End If;
    
    End Loop;
    Close c_费用信息;
  End If;

  --执行科室变更
  --1 、更新费用执行部门及未接费用
  If j_Json.Exist('deptchange_list') Then
    j_List := Pljson_List();
    j_List := j_Json.Get_Pljson_List('deptchange_list');
    For J In 1 .. j_List.Count Loop
      j_Temp           := PLJson();
      j_Temp           := PLJson(j_List.Get(J));
      n_费用id         := Nvl(j_Temp.Get_Number('fee_id'), 0);
      n_原执行部门id   := Nvl(j_Temp.Get_Number('exe_old_deptid'), 0);
      n_当前执行部门id := j_Temp.Get_Number('exe_deptid');
      If Nvl(n_当前执行部门id, 0) = 0 Then
        v_Err_Msg := '传入更改的执行科室为0，请检查!';
        Raise Err_Item;
      End If;
    
      If Nvl(n_费用来源, 1) = 1 Then
        Open c_费用信息 For
          Select /*+cardinality(b,10)*/
           a.病人id, a.No, Nvl(a.价格父号, a.序号) As 序号, Mod(a.记录性质, 10) As 记录性质, 0 主页id, 0 病人病区id, a.病人科室id, a.开单部门id,
           a.执行部门id, a.收入项目id, a.门诊标志, Sum(Nvl(a.实收金额, 0) - Nvl(a.结帐金额, 0)) As 未结金额, Min(a.登记时间) As 登记时间
          From 门诊费用记录 A,
               (Select NO, 序号, Mod(记录性质, 10) As 记录性质
                 From 门诊费用记录
                 Where ID = n_费用id And (Nvl(执行部门id, 0) = n_原执行部门id Or 执行部门id Is Null) And Nvl(执行部门id, 0) <> n_当前执行部门id) B
          Where a.No = b.No And a.序号 = b.序号 And Mod(a.记录性质, 10) = b.记录性质
          Group By a.病人id, a.No, Nvl(a.价格父号, a.序号), Mod(a.记录性质, 10), a.病人科室id, a.开单部门id, a.执行部门id, a.收入项目id, a.门诊标志
          Order By a.No, Nvl(a.价格父号, a.序号);
      Else
        Open c_费用信息 For
          Select /*+cardinality(b,10)*/
           a.病人id, a.No, Mod(a.记录性质, 10) As 记录性质, a.主页id, a.病人病区id, a.病人科室id, a.开单部门id, a.执行部门id, a.收入项目id, a.门诊标志,
           Sum(Nvl(a.实收金额, 0) - Nvl(a.结帐金额, 0)) As 未结金额, Min(a.登记时间) As 登记时间
          From 住院费用记录 A,
               (Select NO, 序号, Mod(记录性质, 10) As 记录性质
                 From 住院费用记录
                 Where ID = n_费用id And (Nvl(执行部门id, 0) = n_原执行部门id Or 执行部门id Is Null) And Nvl(执行部门id, 0) <> n_当前执行部门id) B
          Where a.No = b.No And a.序号 = b.序号 And Mod(a.记录性质, 10) = b.记录性质
          Group By a.病人id, a.No, Nvl(a.价格父号, a.序号), Mod(a.记录性质, 10), a.主页id, a.病人病区id, a.病人科室id, a.开单部门id, a.执行部门id,
                   a.收入项目id, a.门诊标志
          Order By a.No, Nvl(a.价格父号, a.序号);
      End If;
      Loop
        Fetch c_费用信息
          Into r_科室变更;
        Exit When c_费用信息%NotFound;
      
        If r_科室变更.记录性质 = 2 Then
        
          If Trunc(Sysdate) > Trunc(r_科室变更.登记时间) Then
            --病人费用汇总可能已经计算，因此，禁卡更改
            v_Err_Msg := '单据号为' || r_费用信息.No || '的记帐单非当天的记帐单据，禁止更换执行科室!';
            Raise Err_Item;
          End If;
        
          --减原库房的未结费用
          Update 病人未结费用
          Set 金额 = Nvl(金额, 0) - Nvl(r_科室变更.未结金额, 0)
          Where 病人id = r_科室变更.病人id And Nvl(主页id, 0) = Nvl(r_科室变更.主页id, 0) And Nvl(病人病区id, 0) = Nvl(r_科室变更.病人病区id, 0) And
                Nvl(病人科室id, 0) = Nvl(r_科室变更.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(r_科室变更.开单部门id, 0) And
                Nvl(执行部门id, 0) = Nvl(r_科室变更.执行部门id, 0) And 收入项目id + 0 = r_科室变更.收入项目id And 来源途径 + 0 = r_科室变更.门诊标志
          Returning 金额 Into n_返回值1;
        
          If Sql%RowCount <> 0 Then
            --增加现库房的未结费用
            Update 病人未结费用
            Set 金额 = Nvl(金额, 0) + Nvl(r_科室变更.未结金额, 0)
            Where 病人id = r_科室变更.病人id And Nvl(主页id, 0) = Nvl(r_科室变更.主页id, 0) And Nvl(病人病区id, 0) = Nvl(r_科室变更.病人病区id, 0) And
                  Nvl(病人科室id, 0) = Nvl(r_科室变更.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(r_科室变更.开单部门id, 0) And
                  Nvl(执行部门id, 0) = n_当前执行部门id And 收入项目id + 0 = r_科室变更.收入项目id And 来源途径 + 0 = r_科室变更.门诊标志
            Returning 金额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 病人未结费用
                (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
              Values
                (r_科室变更.病人id, Decode(r_科室变更.主页id, 0, Null, r_科室变更.主页id), Decode(r_科室变更.病人病区id, 0, Null, r_科室变更.病人病区id),
                 Decode(r_科室变更.病人科室id, 0, Null, r_科室变更.病人科室id), Decode(r_科室变更.开单部门id, 0, Null, r_科室变更.开单部门id),
                 Decode(n_当前执行部门id, 0, Null, n_当前执行部门id), Decode(r_科室变更.收入项目id, 0, Null, r_科室变更.收入项目id), r_科室变更.门诊标志,
                 Nvl(r_科室变更.未结金额, 0));
              n_返回值 := Nvl(r_科室变更.未结金额, 0);
            End If;
          End If;
        
          If n_返回值 = 0 Or n_返回值1 = 0 Then
            Delete From 病人未结费用 Where 病人id = r_科室变更.病人id And Nvl(金额, 0) = 0;
          End If;
        
        End If;
      
        If Nvl(n_费用来源, 1) = 1 Then
          Update 门诊费用记录
          Set 执行部门id = n_当前执行部门id
          Where NO = r_科室变更.No And Mod(记录性质, 10) = r_科室变更.记录性质 And Nvl(价格父号, 序号) = r_科室变更.序号;
        Else
        
          Update 住院费用记录
          Set 执行部门id = n_当前执行部门id
          Where NO = r_科室变更.No And Mod(记录性质, 10) = r_科室变更.记录性质 And Nvl(价格父号, 序号) = r_科室变更.序号;
        End If;
      
      End Loop;
      Close c_费用信息;
    End Loop;
  End If;

  --费用销帐:取消输液时，需要同步销帐(取消发药时增加的项目)
  If j_Json.Exist('delrcp_list') Then
    j_List := Pljson_List();
    j_List := j_Json.Get_Pljson_List('delrcp_list');
    For J In 1 .. j_List.Count Loop
      j_Temp     := PLJson(j_List.Get(J));
      v_No       := j_Temp.Get_String('rcp_no');
      v_序号     := j_Temp.Get_String('serial_nums');
      n_操作状态 := j_Temp.Get_Number('operator_status');
      If n_操作状态 Is Null Then
        n_操作状态 := 0;
      End If;
      If Nvl(n_费用来源, 1) = 1 Then
        Zl_门诊记帐记录_Delete_s(v_No, v_序号, v_操作员编号, v_操作员姓名, d_操作时间);
      Elsif n_费用来源 = 2 Then
        Zl_住院记帐记录_Delete_s(v_No, v_序号, v_操作员编号, v_操作员姓名, 2, n_操作状态, d_操作时间);
      End If;
    End Loop;
  End If;

  Json_Out := zlJsonOut('成功', 1);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Err_Msg || '+[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Updateexeinfo;
/
Create Or Replace Procedure Zl_Exsesvr_Updatedepositinvinf
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:更新预交票据信息
  --入参：Json_In:格式
  --    input
  --      fun_oper          N 1 操作类型:1-发出;2-重打；3-补打;4-红票打印
  --      deposit_no        C 1 预交单号
  --      recv_id           N 1 领用id
  --      inv_no            C 1 当前发票号或开始使用发票号
  --      inv_usenums       N 1 发票使用数量
  --      use_time          C 1 票据使用时间:yyyy-mm-dd hh24:mi:ss
  --      inv_user          C 1 发票使用人
  --出参: Json_Out,格式如下
  --   output
  --     code               C 1 应答码：0-失败；1-成功
  --     message            C 1 应答消息： 成功时返回成功信息 失败时返回具体的错误信息
  ---------------------------------------------------------------------------

  j_Input PLJson;
  j_Json  PLJson;

  n_操作类型     Number(2);
  v_预交单号     Varchar2(20);
  n_领用id       Number(18);
  n_发票使用数量 Number(18);
  v_票据号     票据使用明细.号码%Type;
  v_使用人       票据使用明细.使用人%Type;
  d_使用时间     票据使用明细.使用时间%Type;
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

Begin
  --解析入参
  
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_操作类型     := j_Json.Get_Number('fun_oper');
  v_预交单号     := j_Json.Get_String('deposit_no');
  n_领用id       := j_Json.Get_Number('recv_id');
  v_票据号       := j_Json.Get_String('inv_no');
  n_发票使用数量 := j_Json.Get_Number('inv_usenums');
  v_使用人       := j_Json.Get_String('inv_user');
  d_使用时间     := To_Date(j_Json.Get_String('use_time'), 'yyyy-mm-dd hh24:mi:ss');
  If d_使用时间 Is Null Then
    d_使用时间 := Sysdate;
  End If;
  If v_预交单号 Is Null Then
    Json_Out := zlJsonOut('未传入需要打印的单据信息!');
    Return;
  End If;
  If n_操作类型=1 Then
     zl_病人预交票据_Insert(v_预交单号,v_票据号,n_领用id,v_使用人,d_使用时间,n_发票使用数量);
  Elsif n_操作类型=2 Then
     zl_病人预交记录_RePrint(v_预交单号,v_票据号,n_领用id,v_使用人);
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Updatedepositinvinf;
/

Create Or Replace Procedure Zl_Exsesvr_Getexsespec
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  -------------------------------------------------------------------------------------------------
  --功能：检查该规格是否产生过费用记录
  --input   根据材料id检查是否产生过费用记录
  --  item_id       N   1   收费细目id
  --output
  --  code          C   1   应答码：0-失败；1-成功
  --  message       C   1   应答消息：
  --  item_id       N   1   收费细目id
  -------------------------------------------------------------------------------------------------
  n_收费细目id Number(18);
  j_Input      PLJson;
  j_Json       PLJson;

Begin
  j_Input := PLJson(Json_In);
  j_Json       := j_Input.Get_Pljson('input');
  n_收费细目id := Nvl(j_Json.Get_Number('item_id'), 0);

  Select Nvl(Max(收费细目id), 0)
  Into n_收费细目id
  From (Select 收费细目id
         From 门诊费用记录
         Where 收费细目id = n_收费细目id And Rownum < 2
         Union All
         Select 收费细目id
         From 住院费用记录
         Where 收费细目id = n_收费细目id And Rownum < 2)
  Where Rownum < 2;
  If Nvl(n_收费细目id, 0) = 0 Then
    Json_Out := '{"output":{"code":1,"message": "成功","item_id":null}}';
  Else
    Json_Out := '{"output":{"code":1,"message": "成功","item_id":' || n_收费细目id || '}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getexsespec;
/


Create Or Replace Procedure Zl_Exsesvr_Delbill_Check
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：针对指定单据指定行行进行销帐
  --入参：Json_In:格式
  --  input
  --      fee_no                  C   1   费用单据号
  --      fee_bill_type           N   1   单据性质:2-门诊记帐单,3-自动记帐单
  --      balance_ban_writeoffs   N   1   已结禁止销帐:如果已结帐单据禁止销帐,或是医保记帐的单据。则在原始单据行中只取未结帐部分
  --      part_ban_writeoffs      N   1   禁止部分销帐:1-不允许；0-允许
  --      fee_origin              N   1   费用来源（1-门诊记帐，2-住院记帐）
  --      item_list[]             本次销帐列表
  --          serial_num          N   1   序号
  --          quantity            N   1   销帐数量(为零时，按序号直接销帐)
  --      excute_list[]           单据已执行列表(药品、卫材费用),即使已执行数为0也要传入
  --          fee_id              N   1   费用ID
  --          sended_num          N   1   已发数量
  --      advice_excute_list[]    单据已执行列表(医嘱费用),即使已执行数为0也要传入
  --          advice_id           N   1   医嘱ID
  --          fee_item_id         N   1   收费细目ID
  --          execute_num         N   1   已执行数
  --      pati_list[]             病人信息，仅审核这些病人的费用
  --          pati_id             N   1   病人ID
  --          fee_audit_status    N   1   费用审核标志:0或空-未审核;1-已审核或开始审核(结合参数:病人审核方式来控制);2-完成审核,结合结帐权限[禁止未审核病人结帐]进行管理控制
  --          si_inp_status       N   1   住院状态:0-正常住院;1-尚未入科;2-正在转科或正在转病区;3-已预出院
  --          catalog_date        C   0   病案编目日期：yyyy-mm-dd hh24:mi:ss
  --出参: Json_Out,格式如下
  --  output
  --      code                    N   1   应答吗：0-失败；1-成功
  --      message                 C   1   应答消息：失败时返回具体的错误信息
  --      item_list[]                         单据数据列表
  --          serial_num          N   1   序号
  --          quantity            N   1   销帐数量
  --          execute_tag         N   1   执行状态：0-未执行;1-已执行;2-部分执行

  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_费用来源 Number(1);
Begin

  j_Input    := PLJson(Json_In);
  j_Json     := j_Input.Get_Pljson('input');
  n_费用来源 := Nvl(j_Json.Get_Number('fee_origin'), 0);

  If n_费用来源 = 1 Or n_费用来源 Is Null Then
    Zl_门诊记帐记录_Delete_Check(Json_In, Json_Out);
  Else
    Zl_住院记帐记录_Delete_Check(Json_In, Json_Out);
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Delbill_Check;
/


Create Or Replace Procedure Zl_Exsesvr_Cancelacc_Reaudit
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --重审之前已拒绝的销账记录
  ---------------------------------------------------------------------------
  --input      重审之前已拒绝的销账记录
  --  rcpdtl_id  N  1  处方明细id(费用id)
  --  request_time  D  1  申请时间
  --  audit_operator  C  1  审核人
  --  fee_audit_time  D  1  审核时间
  --  oper_type  N  1  操作类型:0-审核拒绝 1-取消拒绝
  --  auto_stuff_return  N  1  自动退料
  --  request_type  N    申请类别
  ---------------------------------------------------------------------------
  j_Json  PLJson;
  j_Input PLJson;

  n_Id       Number(18);
  d_申请时间 Date;
  v_审核人   Varchar2(20);
  d_审核时间 Date;
  n_操作类型 Number;
  n_申请类别 Number(2);
Begin

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_Id       := j_Json.Get_Number('rcpdtl_id');
  d_申请时间 := To_Date(j_Json.Get_String('request_time'), 'yyyy-mm-dd hh24:mi:ss');
  v_审核人   := j_Json.Get_String('audit_operator');
  d_审核时间 := To_Date(j_Json.Get_String('fee_audit_time'), 'yyyy-mm-dd hh24:mi:ss');
  n_操作类型 := j_Json.Get_Number('oper_type');
  --Int自动退料 := j_Json.Get_Number('auto_stuff_return');
  n_申请类别 := j_Json.Get_Number('request_type');

  Zl_病人费用销帐_Cancel_s(n_Id, d_申请时间, v_审核人, d_审核时间, n_操作类型, n_申请类别);

  Json_Out := zlJsonOut('成功', 1);
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Cancelacc_Reaudit;
/


Create Or Replace Procedure Zl_Exsesvr_Getrequestcancel
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --查询销帐申请记录
  --入参      json
  --  input      查询是否存在销帐申请记录
  --    query_type          N 1 查询方式:0-根据费用ID查询,1-根据病人变动记录查询(转病区撤销)
  --    rcpdtl_id           C 0 单据明细id,[数组]：[1,2,3],查询方式=0时有效
  --    request_type        N 0 申请类别,查询方式=0时有效
  --    cancel_status       N 1 申请状态,查询方式=0时有效
  --    change_id_old       N 0 原病区的变动记录的ID,查询方式=1时有效
  --    change_id_new       N 0 目标病区的变动记录的ID,查询方式=1时有效
  --出参      json
  -- output
  --   code     C  1   应答码：0-失败；1-成功
  --   message  C  1   应答消息：
  --   fee_cancel_list      [数组]满足条件的每个费用销帐记录
  --     rcpdtl_id          N    处方明细id(费用id)
  --     apply_type         N    申请类别:对药品和卫材有效:0-未执行;1-已执行;非药品和卫材固定存为0
  --     apply_time         N    申请时间:yyyy-mm-dd hh24:mi:ss
  --     aplnt_name         N    申请人
  --     apply_dept_id      N    申请部门id
  --     apply_dept_name    N    申请部门名称
  --     audit_dept_id      N    审核部门id;
  --     audit_dept_name    N    审核部门名称
  --     bill_no            N    费用单据号
  --     item_id            N    收费细目id
  --     item_name          N    收费项目名称
  --     quantity           N    数量
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  j_Jsonlist Pljson_List;

  n_查询类型 Number(2);

  n_申请类别 Number(1);
  n_状态     Number(1);
  l_Feelist  t_NumList := t_NumList();

  n_原变动id   费用变动记录.原变动id%Type;
  n_目标变动id 费用变动记录.目标变动id%Type;

  n_Firstitem Number(1);
  v_Temp      Varchar2(32767);
  c_Temp      Clob;
Begin
  --解析入参

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_查询类型 := j_Json.Get_Number('query_type');

  n_Firstitem := 1;
  v_Temp      := '{"output":{"code":1,"message":"成功","fee_cancel_list":[';
  If Nvl(n_查询类型, 0) = 0 Then
    n_申请类别 := j_Json.Get_Number('request_type');
    n_状态     := j_Json.Get_Number('cancel_status');
    j_Jsonlist := j_Json.Get_Pljson_List('rcpdtl_id');
    If j_Jsonlist Is Not Null Then
      --按费用id查询
      For I In 1 .. j_Jsonlist.Count Loop
        l_Feelist.Extend;
        l_Feelist(l_Feelist.Count) := j_Jsonlist.Get_Number(I);
      End Loop;
    
      For r_费用 In (Select /*+cardinality(b,10)*/
                    a.费用id
                   From 病人费用销帐 A, Table(l_Feelist) B
                   Where a.费用id = b.Column_Value And a.申请类别 = n_申请类别 And a.状态 = n_状态) Loop
      
        If Nvl(n_Firstitem, 0) = 0 Then
          v_Temp := v_Temp || ',';
        Else
          n_Firstitem := 0;
        End If;
      
        v_Temp := v_Temp || '{';
        v_Temp := v_Temp || '"rcpdtl_id":' || r_费用.费用id;
        v_Temp := v_Temp || '}';
      
        If Length(v_Temp) > 30000 Then
          c_Temp := c_Temp || To_Clob(v_Temp);
          v_Temp := '';
        End If;
      End Loop;
    End If;
  
  Elsif Nvl(n_查询类型, 0) = 1 Then
    n_原变动id   := j_Json.Get_Number('change_id_old');
    n_目标变动id := j_Json.Get_Number('change_id_new');
  
    For r_费用 In (Select a.申请类别, To_Char(a.申请时间, 'yyyy-mm-dd hh24:mi:ss') As 申请时间, a.申请人, a.申请部门id, e.名称 As 申请部门, a.审核部门id,
                        f.名称 As 审核部门, b.No, a.收费细目id, c.名称 As 收费项目, Sum(a.数量) As 数量
                 From 病人费用销帐 A, 费用变动记录 B, 收费项目目录 C, 部门表 E, 部门表 F
                 Where a.费用id = b.费用id And a.收费细目id = c.Id And a.申请部门id = e.Id And a.审核部门id = f.Id And b.原变动id = n_原变动id And
                       b.目标变动id = n_目标变动id And b.状态 = 2 And a.状态 In (0, 2)
                 Group By a.申请类别, a.申请时间, a.申请人, a.申请部门id, e.名称, a.审核部门id, f.名称, b.No, a.收费细目id, c.名称
                 Order By NO, 收费细目id) Loop
    
      If Nvl(n_Firstitem, 0) = 0 Then
        v_Temp := v_Temp || ',';
      Else
        n_Firstitem := 0;
      End If;
    
      v_Temp := v_Temp || '{';
      v_Temp := v_Temp || '"apply_type":' || r_费用.申请类别;
      v_Temp := v_Temp || ',"apply_time":"' || zlJsonStr(r_费用.申请时间) || '"';
      v_Temp := v_Temp || ',"aplnt_name":"' || zlJsonStr(r_费用.申请人) || '"';
      v_Temp := v_Temp || ',"apply_dept_id":' || r_费用.申请部门id;
      v_Temp := v_Temp || ',"apply_dept_name":"' || r_费用.申请部门 || '"';
      v_Temp := v_Temp || ',"audit_dept_id":' || r_费用.审核部门id;
      v_Temp := v_Temp || ',"audit_dept_name":"' || r_费用.审核部门 || '"';
      v_Temp := v_Temp || ',"bill_no":"' || r_费用.No || '"';
      v_Temp := v_Temp || ',"item_id":' || r_费用.收费细目id;
      v_Temp := v_Temp || ',"item_name":"' || zlJsonStr(r_费用.收费项目) || '"';
      v_Temp := v_Temp || ',"quantity":' || zlJsonStr(r_费用.数量, 1);
      v_Temp := v_Temp || '}';
    
      If Length(v_Temp) > 30000 Then
        c_Temp := c_Temp || To_Clob(v_Temp);
        v_Temp := '';
      End If;
    End Loop;
  End If;
  v_Temp := v_Temp || ']}}';

  If c_Temp Is Not Null Then
    Json_Out := c_Temp || To_Clob(v_Temp);
  Else
    Json_Out := v_Temp;
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getrequestcancel;
/


Create Or Replace Procedure Zl_Exsesvr_Getremainmoney
(
  Json_In  In Varchar2,
  Json_Out Out Clob
) Is
  --获取病人费用余额
  ---------------------------------------------------------------------------
  --input      获取病人费用余额
  --  pati_id                 N  1  病人ID
  --  pati_pageid             N  1  主页ID
  --  insure_account_balance  N  1  医保账户余额
  --  query_type              N  0  查询方式；0-按单病人查询，1-批量查询病人余额，2-批量查询，担保额和适用病人信息
  --  pati_ids                C  0  批量查询病人关键信息拼串，两种格式：1-病人ID1:主页ID1,病人ID2:主页ID2,....；2-病人ID1,病人ID2,....
  --  fee_source              N  1  费用来源：0-不区分，1-门诊，2-住院；查询方式=1且仅按病人ID查询时有效
  --output
  --  code                    C  1  应答码：0-失败；1-成功
  --  message                 C  1  应答消息
  --  remain_money            N     剩余款
  --  guarantee_money         N     担保额
  --  expected_money          N     预结费用
  --  prepay_money            N  0  预交余额
  --  nobalance_money         N  0  未结费用金额
  --  item_list[]当传入批量病人信息时才返回，该列表可以不返回
  --       pati_id            N 1 病人id
  --       pati_pageid        N 1 主页id
  --       prepay_money       N 0 预交余额
  --       nobalance_money    N 0 未结费用金额
  --       remain_money       N 1 剩余款
  --       guarantee_money    N 1 担保额
  --       pati_type          C 1 适用病人
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;

  n_病人id   病人余额.病人id%Type;
  n_主页id   保险模拟结算.主页id%Type;
  n_担保额   病人余额.预交余额%Type;
  n_帐户余额 病人余额.预交余额%Type;
  n_剩余款   病人余额.预交余额%Type;
  n_预结费用 病人余额.预交余额%Type;
  n_预交余额 病人余额.预交余额%Type;
  n_费用余额 病人余额.预交余额%Type;

  n_Find     Number(2);
  n_查询方式 Number(2);
  l_病人ids  t_StrList;
  c_病人ids  Clob;
  n_费用来源 Number(2);
Begin
  j_Input    := PLJson(Json_In);
  j_Json     := j_Input.Get_Pljson('input');
  n_查询方式 := j_Json.Get_Number('query_type');

  --按单病人查询
  If Nvl(n_查询方式, 0) = 0 Then
    n_病人id   := j_Json.Get_Number('pati_id');
    n_主页id   := j_Json.Get_Number('pati_pageid');
    n_帐户余额 := j_Json.Get_Number('insure_account_balance');
    n_担保额   := Zl_Patientsurety(n_病人id, n_主页id);
  
    n_Find := 0;
    If n_主页id > 0 Then
      Select (Nvl(Sum(a.预交余额), 0) - Nvl(Sum(a.费用余额), 0) + Nvl(Sum(a.预结费用), 0)) As 剩余款, Nvl(Sum(a.预结费用), 0) As 预结费用,
             Nvl(Sum(a.预交余额), 0) As 预交余额, Nvl(Sum(a.费用余额), 0) As 费用余额
      Into n_剩余款, n_预结费用, n_预交余额, n_费用余额
      From (Select 病人id, 预交余额, 费用余额, 0 As 预结费用
             From 病人余额
             Where 性质 = 1 And 类型 = 2 And 病人id = n_病人id
             Union All
             Select a.病人id, 0, 0, Sum(金额)
             From 保险模拟结算 A
             Where a.病人id = n_病人id And a.主页id = n_主页id
             Group By a.病人id) A;
    
      n_Find := 1;
      v_Jtmp := v_Jtmp || ',"prepay_money":' || zlJsonStr(n_预交余额, 1);
      v_Jtmp := v_Jtmp || ',"nobalance_money":' || zlJsonStr(n_费用余额, 1);
      v_Jtmp := v_Jtmp || ',"expected_money":' || zlJsonStr(n_预结费用, 1);
      v_Jtmp := v_Jtmp || ',"guarantee_money":' || zlJsonStr(n_担保额, 1);
      v_Jtmp := v_Jtmp || ',"remain_money":' || zlJsonStr(n_剩余款, 1);
    Else
      Select (Nvl(Sum(预交余额), 0) - Nvl(Sum(费用余额), 0) + Nvl(n_帐户余额, 0)) As 剩余款, Nvl(Sum(预交余额), 0) As 预交余额,
             Nvl(Sum(费用余额), 0) As 费用余额
      Into n_剩余款, n_预交余额, n_费用余额
      From 病人余额
      Where 性质 = 1 And 类型 = 1 And 病人id = n_病人id;
    
      n_Find := 1;
      v_Jtmp := v_Jtmp || ',"prepay_money":' || zlJsonStr(n_预交余额, 1);
      v_Jtmp := v_Jtmp || ',"nobalance_money":' || zlJsonStr(n_费用余额, 1);
      v_Jtmp := v_Jtmp || ',"guarantee_money":' || zlJsonStr(n_担保额, 1);
      v_Jtmp := v_Jtmp || ',"remain_money":' || zlJsonStr(n_剩余款, 1);
    End If;
  
    If n_Find = 0 Then
      If n_主页id > 0 Then
        n_帐户余额 := 0;
      End If;
      v_Jtmp := v_Jtmp || ',"guarantee_money":' || zlJsonStr(n_帐户余额, 1);
    End If;
    Json_Out := '{"output":{"code":1,"message":"成功"' || v_Jtmp || '}}';
    Return;
  End If;

  --批量查询
  v_Jtmp     := Null;
  c_病人ids  := j_Json.Get_Clob('pati_ids');
  n_费用来源 := Nvl(j_Json.Get_Number('fee_source'), 0);
  l_病人ids  := t_StrList();
  While c_病人ids Is Not Null Loop
    If Length(c_病人ids) <= 4000 Then
      l_病人ids.Extend;
      l_病人ids(l_病人ids.Count) := c_病人ids;
      c_病人ids := Null;
    Else
      l_病人ids.Extend;
      l_病人ids(l_病人ids.Count) := Substr(c_病人ids, 1, Instr(c_病人ids, ',', 3980) - 1);
      c_病人ids := Substr(c_病人ids, Instr(c_病人ids, ',', 3980) + 1);
    End If;
  End Loop;

  If n_查询方式 = 1 Then
    For I In 1 .. l_病人ids.Count Loop
      If Instr(l_病人ids(I), ':') = 0 Then
        --格式：2-病人ID1,病人ID2,....
        For R In (Select /*+cardinality(b,10)*/
                   a.病人id, Nvl(Sum(a.预交余额), 0) As 预交余额, Nvl(Sum(a.费用余额), 0) As 费用余额,
                   Nvl(Sum(a.预交余额), 0) - Nvl(Sum(a.费用余额), 0) As 剩余款
                  From 病人余额 A, Table(f_Num2List(l_病人ids(I))) B
                  Where a.病人id = b.Column_Value And a.性质 = 1 And Decode(n_费用来源, 0, 0, a.类型) = n_费用来源
                  Group By a.病人id) Loop
        
          v_Jtmp := v_Jtmp || ',{"pati_id":' || r.病人id;
          v_Jtmp := v_Jtmp || ',"remain_money":' || zlJsonStr(r.剩余款, 1);
          v_Jtmp := v_Jtmp || ',"prepay_money":' || zlJsonStr(r.预交余额, 1);
          v_Jtmp := v_Jtmp || ',"nobalance_money":' || zlJsonStr(r.费用余额, 1);
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
        --格式：1-病人ID1:主页ID1,病人ID2:主页ID2,....
        For R In (Select a.病人id, a.主页id, Nvl(Sum(a.预交余额), 0) As 预交余额, Nvl(Sum(a.费用余额), 0) As 费用余额,
                         (Nvl(Sum(a.预交余额), 0) - Nvl(Sum(a.费用余额), 0) + Nvl(Sum(a.预结费用), 0)) As 剩余款
                  From (Select n.主页id, a.病人id, a.预交余额, a.费用余额, 0 As 预结费用
                         From 病人余额 A,
                              (Select /*+cardinality(f,10)*/
                                 To_Number(f.C1) As 病人id, To_Number(f.C2) As 主页id
                                From Table(f_Num2List2(l_病人ids(I))) F) N
                         Where a.性质 = 1 And a.类型 = 2 And a.病人id = n.病人id
                         Union All
                         Select a.主页id, a.病人id, 0, 0, Sum(a.金额)
                         From 保险模拟结算 A,
                              (Select /*+cardinality(f,10)*/
                                 To_Number(f.C1) As 病人id, To_Number(f.C2) As 主页id
                                From Table(f_Num2List2(l_病人ids(I))) F) N
                         Where a.病人id = n.病人id And a.主页id = n.主页id
                         Group By a.病人id, a.主页id) A
                  Group By a.病人id, a.主页id
                  Having Nvl(Sum(a.预交余额), 0) - Nvl(Sum(a.费用余额), 0) + Nvl(Sum(a.预结费用), 0) <> 0) Loop
        
          v_Jtmp := v_Jtmp || ',{"pati_id":' || r.病人id;
          v_Jtmp := v_Jtmp || ',"pati_pageid":' || r.主页id;
          v_Jtmp := v_Jtmp || ',"remain_money":' || zlJsonStr(r.剩余款, 1);
          v_Jtmp := v_Jtmp || ',"prepay_money":' || zlJsonStr(r.预交余额, 1);
          v_Jtmp := v_Jtmp || ',"nobalance_money":' || zlJsonStr(r.费用余额, 1);
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
    End Loop;
  End If;

  If n_查询方式 = 2 Then
    For I In 1 .. l_病人ids.Count Loop
      For R In (Select n.病人id, n.主页id, Zl_Patiwarnscheme(n.病人id) As 适用病人, a.担保额
                From (Select a.病人id, a.主页id, Sum(a.担保额) As 担保额
                       From 病人担保记录 A,
                            (Select /*+cardinality(f,10)*/
                               To_Number(f.C1) As 病人id, To_Number(f.C2) As 主页id
                              From Table(f_Num2List2(l_病人ids(I))) F) N
                       Where a.病人id = n.病人id And Nvl(a.主页id, 0) = n.主页id And (a.到期时间 Is Null Or a.到期时间 > Sysdate) And
                             a.删除标志 = 1
                       Group By a.病人id, a.主页id) A,
                     (Select /*+cardinality(f,10)*/
                        To_Number(f.C1) As 病人id, To_Number(f.C2) As 主页id
                       From Table(f_Num2List2(l_病人ids(I))) F) N
                Where n.病人id = a.病人id(+) And n.主页id = a.主页id(+)) Loop
      
        v_Jtmp := v_Jtmp || ',{"pati_id":' || r.病人id;
        v_Jtmp := v_Jtmp || ',"pati_pageid":' || r.主页id;
        v_Jtmp := v_Jtmp || ',"guarantee_money":' || zlJsonStr(r.担保额, 1);
        v_Jtmp := v_Jtmp || ',"pati_type":"' || zlJsonStr(r.适用病人) || '"';
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
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getremainmoney;
/
 
Create Or Replace Procedure Zl_Exsesvr_Getnextno
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
  --    quantity            N   0   所需no号的个数，如果只取一个该参不传或都传0 
  --出参: Json_Out,格式如下
  --  output
  --    code                N   1   应答码：0-失败；1-成功
  --    message             C   1   应答消息：失败时返回具体的错误信息
  --    next_no             C   1   下一个号码,quantity>1 时，表示取多个单号,多个时用逗号分离
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_No     Varchar2(64);
  n_序号   Number(10);
  n_科室id Number(18);
  n_数量   Number;
  v_Nos    Varchar2(32767);
Begin

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_序号   := j_Json.Get_Number('item_num');
  n_科室id := j_Json.Get_Number('dept_id');
  n_数量   := j_Json.Get_Number('quantity');

  If Nvl(n_数量, 0) > 1 Then
    For I In 1 .. n_数量 Loop
      Select Zl_Exse_Nextno(n_序号, n_科室id) Into v_No From Dual;
      v_Nos := v_Nos || ',' || v_No;
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"成功","next_no":"' || Substr(v_Nos, 2) || '"}}';
  Else
    Select Zl_Exse_Nextno(n_序号, n_科室id) Into v_No From Dual;
    Json_Out := '{"output":{"code":1,"message":"成功","next_no":"' || v_No || '"}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getnextno;
/


Create Or Replace Procedure Zl_Exsesvr_Cancelacc_Check
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --费用销帐记录核查
  ---------------------------------------------------------------------------
  --input     费用销帐记录核查
  --  check_people  C 1 核查人
  --  check_time    D 1 核查时间
  --  request_type  N   申请类别：0-未执行;1-已执行;非药品和卫材固定存为0
  --  rcpdtl_list     [数组]每个处方明细信息
  --    rcpdtl_id     N 1 处方明细id(费用id)
  --    request_time  D 1 申请时间
  --output
  --  code         C 1 应答码：0-失败；1-成功
  --  message      C 1 应答消息：
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  j_List     Pljson_List;
  j_Temp     PLJson;
  n_Id       Number(18);
  d_申请时间 Date;
  v_核查人   Varchar2(20);
  d_核查日期 Date;
  n_申请类别 Number(2); --0-未执行;1-已执行;非药品和卫材固定存为0
Begin
  j_Input    := PLJson(Json_In);
  j_Json     := j_Input.Get_Pljson('input');
  v_核查人   := j_Json.Get_String('check_people');
  d_核查日期 := To_Date(j_Json.Get_String('check_time'), 'yyyy-mm-dd hh24:mi:ss');
  n_申请类别 := j_Json.Get_Number('request_type');
  If n_申请类别 Is Null Then
    n_申请类别 := 1;
  End If;

  j_List := j_Json.Get_Pljson_List('rcpdtl_list');
  For J In 1 .. j_List.Count Loop
    j_Temp     := PLJson();
    j_Temp     := PLJson(j_List.Get(J));
    n_Id       := j_Temp.Get_Number('rcpdtl_id');
    d_申请时间 := To_Date(j_Temp.Get_String('request_time'), 'yyyy-mm-dd hh24:mi:ss');
  
    Zl_病人费用销帐_Check(n_Id, d_申请时间, v_核查人, d_核查日期, n_申请类别);
  End Loop;

  Json_Out := zlJsonOut('成功', 1);

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Cancelacc_Check;
/
 
Create Or Replace Procedure Zl_Exsesvr_Setsendwin
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：设置费用药品单据的发药窗口
  --入参：Json_In:格式
  --  input
  --    pharmacy_id              N   1  库房id
  --    pharmacy_window_old      C   1  旧发药窗口
  --    pharmacy_window_new      C   1  新发药窗口
  --    bill_list[]
  --      billtype               N   1 单据类型:1-收费处方;2-记帐处方
  --      rcp_no                 C   1 处方No
  --出参: Json_Out,格式如下
  --  output
  --     code                   N   1 应答吗：0-失败；1-成功
  --     message                C   1 应答消息：失败时返回具体的错误信息
  ------------------------------------------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  o_Json      PLJson;
  j_Bill_List Pljson_List;
  n_库房id    Number(18);
  v_旧窗口    Varchar2(50);
  v_新窗口    Varchar2(50);
  n_性质      Number(1);
  v_No        Varchar2(20);
  n_Count     Number;
Begin

  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_库房id    := j_Json.Get_Number('pharmacy_id');
  v_旧窗口    := j_Json.Get_String('pharmacy_window_old');
  v_新窗口    := j_Json.Get_String('pharmacy_window_new');
  j_Bill_List := j_Json.Get_Pljson_List('bill_list');
  n_Count     := j_Bill_List.Count;

  If n_Count > 0 Then
    For I In 1 .. n_Count Loop
      o_Json := PLJson(j_Bill_List.Get(I));
      n_性质 := o_Json.Get_Number('billtype');
      v_No   := o_Json.Get_String('rcp_no');
    
      Update 门诊费用记录
      Set 发药窗口 = v_新窗口
      Where 执行部门id = n_库房id And 记录性质 = n_性质 And NO = v_No And 发药窗口 = v_旧窗口;
    
      Update 住院费用记录
      Set 发药窗口 = v_新窗口
      Where 执行部门id = n_库房id And 记录性质 = n_性质 And NO = v_No And 发药窗口 = v_旧窗口;
    End Loop;
  End If;

  Json_Out := zlJsonOut('成功', 1);
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Setsendwin;
/

Create Or Replace Procedure Zl_Exsesvr_Getnobyinvoice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  -------------------------------------------------------------------------------------------------
  --功能：按票据号发药或退药中通过录入发票号获取对应的药品处方NO
  --入参：json格式
  --Input
  --   invc_no  C  1  票据号
  --出参：json格式
  --Json_Out
  --  code  C  1  应答码：0-失败；1-成功
  --  message  C  1  "应答消息： 成功时返回处方No，[数组] 失败时返回具体的错误信息"
  --  rcp_nos  C  1 处理方单据号：多个用逗号分隔 
  -------------------------------------------------------------------------------------------------
  v_票据号 票据使用明细.号码%Type;

  v_Tmp   Varchar2(32767);
  j_Input PLJson;
  j_Json  PLJson;

Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_票据号 := j_Json.Get_String('invc_no');
  For v_药品处方no In (Select Distinct a.No
                   From 票据打印内容 A, 票据使用明细 B
                   Where a.Id = b.打印id And a.数据性质 = 1 And b.票种 = 1 And b.号码 = v_票据号) Loop
  
    v_Tmp := Nvl(v_Tmp, '') || ',' || v_药品处方no.No;
  End Loop;
  If v_Tmp Is Not Null Then
    v_Tmp := Substr(v_Tmp, 2);
  End If;
  Json_Out := '{"output":{"code":1,"message": "成功","rcp_nos":"' || Nvl(v_Tmp, '') || '"}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getnobyinvoice;
/

Create Or Replace Procedure Zl_Exsesvr_Getchargeoffinfo
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --根据条件查询费用销帐信息
  --入参      json
  --  input      根据条件查询费用销帐信息
  --    audit_dept_id       N    审核部门ID(药房)
  --    request_begin_time  D    申请开始时间
  --    request_end_time    D    申请结束时间
  --    audit_begin_time    D    审核开始时间
  --    audit_end_time      D    审核结束时间
  --    cancel_status       N  1 状态
  --    request_dept_id     N    申请部门ID
  --    request_operator    C    申请人
  --    pati_id             N    病人ID
  --    cancel_condition    C    销账条件
  --    cancel_check        N    核查（选择参数【销账申请需要核查】时传入，0-未核查 1-已核查）
  --    rcpdtl_id          C     处方明细id,[数组]：[1,2,3]
  --    request_dept_ids   C     申请部门id串，用于批量查询
  --    item_ids           C     收费细目id串,用于批量查询
  --    request_type       N     申请类别：-1-不区分;0-未执行;1-已执行
  --出参      json
  -- output
  --   code     C  1   应答码：0-失败；1-成功
  --   message  C  1   应答消息：
  --   fee_cancel_list      [数组]满足条件的每个费用销帐记录
  --     rcpdtl_id          N    处方明细id(费用id)
  --     request_type       N    申请类别
  --     item_id            N    收费细目id
  --     request_dept_id    N    申请部门id
  --     request_dept       C    申请部门
  --     audit_dept_id      N    审核部门id
  --     quantity           N    数量
  --     request_operator   C    申请人
  --     request_time       D    申请时间
  --     auditor            C    审核人
  --     audit_time         D    审核时间
  --     cancel_status      N    状态
  --     cancel_reason      C    销帐原因
  --     checker            C    核查人
  --     price_retail       N    零售价
  --     advice_id          N    医嘱id
  --     pati_id            N    病人ID
  --     pati_name          C    病人姓名
  --     inpatient_num      C    住院号
  --     pati_pageid        N    主页id
  ---------------------------------------------------------------------------

  v_Sql Varchar2(4000);

  j_Input Pljson;
  j_Json  Pljson;

  j_Jsonlist Pljson_List := Pljson_List();

  n_审核部门id   Number(18);
  v_申请开始时间 Varchar2(50);
  v_申请结束时间 Varchar2(50);
  v_审核开始时间 Varchar2(50);
  v_审核结束时间 Varchar2(50);
  n_状态         Number(1);
  n_申请部门id   Number(18);
  v_申请人       Varchar2(20);
  n_病人id       Number(18);
  v_销账条件     Varchar2(32767); --申请时间,病人id|申请时间,病人id...
  n_核查         Number(1); --状态=0时使用
  v_申请部门ids  Varchar2(32767);
  v_收费项目ids  Varchar2(32767);
  n_申请类别     Number(2);

  n_Count   Number := 0;
  l_Feelist t_Numlist := t_Numlist();

  v_Output Varchar2(32767);
  c_Output Clob;

  Type t_费用信息 Is Ref Cursor;
  c_费用信息 t_费用信息; --动态游标变量

  Cursor c_销帐信息 Is
    Select a.费用id, a.申请类别, a.收费细目id, a.申请部门id, c.名称, a.审核部门id, a.数量, a.申请人, a.申请时间, a.审核人, a.审核时间, a.状态, a.销帐原因, a.核查人,
           b.标准单价, b.医嘱序号, b.病人id, b.姓名, b.标识号, b.主页id
    From 病人费用销帐 A, 住院费用记录 B, 部门表 C
    Where a.费用id = b.Id And a.申请部门id = c.Id And a.申请类别 = 1 And a.审核部门id = 0 And a.费用id = 0 And a.状态 = 0 And
          a.审核人 Is Null;
  r_销帐 c_销帐信息%RowType;

Begin
  --解析入参
  j_Input := Pljson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_审核部门id   := j_Json.Get_Number('audit_dept_id');
  v_申请开始时间 := j_Json.Get_String('request_begin_time');
  v_申请结束时间 := j_Json.Get_String('request_end_time');
  v_审核开始时间 := j_Json.Get_String('audit_begin_time');
  v_审核结束时间 := j_Json.Get_String('audit_end_time');
  n_状态         := j_Json.Get_Number('cancel_status');
  n_申请部门id   := j_Json.Get_Number('request_dept_id');
  v_申请人       := j_Json.Get_String('request_operator');
  n_病人id       := j_Json.Get_Number('pati_id');
  v_销账条件     := j_Json.Get_String('cancel_condition');
  n_核查         := j_Json.Get_Number('cancel_check');
  v_申请部门ids  := j_Json.Get_String('request_dept_ids');
  v_收费项目ids  := j_Json.Get_String('item_ids');
  n_申请类别     := Nvl(j_Json.Get_Number('request_type'), 1);

  j_Jsonlist := j_Json.Get_Pljson_List('rcpdtl_id');

  v_Output := Null;
  If j_Jsonlist Is Not Null Then
    --按费用id查询
    n_Count := j_Jsonlist.Count;
  
    For I In 1 .. n_Count Loop
      l_Feelist.Extend;
      l_Feelist(l_Feelist.Count) := j_Jsonlist.Get_Number(I);
    End Loop;
  
    For c_销帐申请 In (Select a.费用id, a.申请类别, a.收费细目id, a.申请部门id, c.名称, a.审核部门id, a.数量, a.申请人, a.申请时间, a.审核人, a.审核时间, a.状态,
                          a.销帐原因, a.核查人, b.标准单价, b.医嘱序号, b.病人id, b.姓名, b.标识号, b.主页id
                   From 病人费用销帐 A, 住院费用记录 B, 部门表 C
                   Where a.费用id = b.Id And a.申请部门id = c.Id And a.申请类别 = Decode(n_申请类别, -1, a.申请类别, n_申请类别) And
                         a.状态 = n_状态 And a.审核部门id = Nvl(n_审核部门id, a.审核部门id) And
                         a.费用id In (Select /*+cardinality(j,10)*/
                                     Column_Value
                                    From Table(l_Feelist) J)) Loop
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      
        v_Output := Null;
      End If;
    
      Zljsonputvalue(v_Output, 'rcpdtl_id', c_销帐申请.费用id, 1, 1);
      Zljsonputvalue(v_Output, 'request_type', c_销帐申请.申请类别, 1);
      Zljsonputvalue(v_Output, 'item_id', c_销帐申请.收费细目id, 1);
      Zljsonputvalue(v_Output, 'request_dept_id', c_销帐申请.申请部门id, 1);
      Zljsonputvalue(v_Output, 'request_dept', c_销帐申请.名称, 0);
      Zljsonputvalue(v_Output, 'audit_dept_id', c_销帐申请.审核部门id, 1);
      Zljsonputvalue(v_Output, 'quantity', c_销帐申请.数量, 1);
      Zljsonputvalue(v_Output, 'request_operator', c_销帐申请.申请人, 0);
      Zljsonputvalue(v_Output, 'request_time', To_Char(c_销帐申请.申请时间, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_Output, 'auditor', c_销帐申请.审核人, 0);
      Zljsonputvalue(v_Output, 'audit_time', To_Char(c_销帐申请.审核时间, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_Output, 'cancel_status', c_销帐申请.状态, 1);
      Zljsonputvalue(v_Output, 'cancel_reason', c_销帐申请.销帐原因, 0);
      Zljsonputvalue(v_Output, 'checker', c_销帐申请.核查人, 0);
      Zljsonputvalue(v_Output, 'price_retail', c_销帐申请.标准单价, 1);
      Zljsonputvalue(v_Output, 'advice_id', c_销帐申请.医嘱序号, 1);
      Zljsonputvalue(v_Output, 'pati_id', c_销帐申请.病人id, 1);
      Zljsonputvalue(v_Output, 'pati_name', c_销帐申请.姓名, 0);
      Zljsonputvalue(v_Output, 'inpatient_num', c_销帐申请.标识号, 0);
      Zljsonputvalue(v_Output, 'pati_pageid', c_销帐申请.主页id, 1, 2);
    
    End Loop;
  Else
  
    v_Sql := Nvl(v_Sql, '') || '   Select A.费用id, A.申请类别, A.收费细目id, A.申请部门id, C.名称, A.审核部门id, ' ||
             '   A.数量, A.申请人, A.申请时间 , A.审核人 , A.审核时间 , A.状态 , A.销帐原因 ,' ||
             '   A.核查人 , B.标准单价 , B.医嘱序号 , B.病人id , B.姓名 , B.标识号,b.主页id ';
    v_Sql := v_Sql || '   From 病人费用销帐 A, 住院费用记录 B,部门表 C';
    If v_销账条件 Is Not Null Then
      v_Sql := v_Sql || '   ,Table(f_Str2list2(''' || v_销账条件 || ''', ''| '', '','')) T';
    End If;
    v_Sql := v_Sql || '   Where A.费用id = B.Id And A.申请部门id = C.Id';
  
    If Nvl(n_审核部门id, 0) <> 0 Then
      v_Sql := v_Sql || ' And A.审核部门id = ' || n_审核部门id;
    End If;
  
    If n_申请类别 <> -1 Then
      v_Sql := v_Sql || ' And A.申请类别=' || n_申请类别;
    End If;
  
    If n_状态 = 0 Then
      v_Sql := v_Sql || '   And  A.状态 = 0 And A.审核人 Is Null';
      If n_核查 Is Not Null Then
        If n_核查 = 0 Then
          v_Sql := v_Sql || '   And A.核查人 Is Null ';
        Else
          v_Sql := v_Sql || '   And A.核查人 Is Not Null ';
        End If;
      End If;
    
      If v_申请开始时间 Is Not Null Then
        v_Sql := v_Sql || '   And A.申请时间 Between to_date(''' || v_申请开始时间 ||
                 ''',''yyyy-mm-dd hh24:mi:ss'') And to_date(''' || v_申请结束时间 || ''',''yyyy-mm-dd hh24:mi:ss'') ';
      End If;
    Else
      v_Sql := v_Sql || '   And  A.状态 <> 0 And A.审核人 Is Not Null ';
      If v_审核开始时间 Is Not Null Then
        v_Sql := v_Sql || '   And A.审核时间 Between to_date(''' || v_审核开始时间 ||
                 ''',''yyyy-mm-dd hh24:mi:ss'') And to_date(''' || v_审核结束时间 || ''',''yyyy-mm-dd hh24:mi:ss'') ';
      End If;
    End If;
  
    If n_申请部门id Is Not Null Then
      v_Sql := v_Sql || '   And A.申请部门id = ' || n_申请部门id;
    End If;
  
    If v_申请部门ids Is Not Null Then
      v_Sql := v_Sql || '   And Instr('',' || v_申请部门ids || ','', '','' || A.申请部门id || '','') > 0 ';
    End If;
  
    If v_收费项目ids Is Not Null Then
      v_Sql := v_Sql || '   And Instr('',' || v_收费项目ids || ','', '','' || A.收费细目id || '','') > 0 ';
    End If;
  
    If v_申请人 Is Not Null Then
      v_Sql := v_Sql || '   And A.申请人 = ''' || v_申请人 || '''';
    End If;
  
    If n_病人id Is Not Null Then
      v_Sql := v_Sql || '   And B.病人ID = ' || n_病人id;
    End If;
  
    If v_销账条件 Is Not Null Then
      v_Sql := v_Sql || '   And A.申请时间 = To_Date(t.C1,''yyyy-mm-dd hh24:mi:ss'') And B.病人ID = t.C2';
    End If;
  
    Open c_费用信息 For v_Sql;
    Loop
      Fetch c_费用信息
        Into r_销帐;
      Exit When c_费用信息 %NotFound;
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
        v_Output := Null;
      End If;
    
      Zljsonputvalue(v_Output, 'rcpdtl_id', r_销帐.费用id, 1, 1);
      Zljsonputvalue(v_Output, 'request_type', r_销帐.申请类别, 1);
      Zljsonputvalue(v_Output, 'item_id', r_销帐.收费细目id, 1);
      Zljsonputvalue(v_Output, 'request_dept_id', r_销帐.申请部门id, 1);
      Zljsonputvalue(v_Output, 'request_dept', r_销帐.名称, 0);
      Zljsonputvalue(v_Output, 'audit_dept_id', r_销帐.审核部门id, 1);
      Zljsonputvalue(v_Output, 'quantity', r_销帐.数量, 1);
      Zljsonputvalue(v_Output, 'request_operator', r_销帐.申请人, 0);
      Zljsonputvalue(v_Output, 'request_time', To_Char(r_销帐.申请时间, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_Output, 'auditor', r_销帐.审核人, 0);
      Zljsonputvalue(v_Output, 'audit_time', To_Char(r_销帐.审核时间, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_Output, 'cancel_status', r_销帐.状态, 1);
      Zljsonputvalue(v_Output, 'cancel_reason', r_销帐.销帐原因, 0);
      Zljsonputvalue(v_Output, 'checker', r_销帐.核查人, 0);
      Zljsonputvalue(v_Output, 'price_retail', r_销帐.标准单价, 1);
      Zljsonputvalue(v_Output, 'advice_id', r_销帐.医嘱序号, 1);
      Zljsonputvalue(v_Output, 'pati_id', r_销帐.病人id, 1);
      Zljsonputvalue(v_Output, 'pati_name', r_销帐.姓名, 0);
      Zljsonputvalue(v_Output, 'inpatient_num', r_销帐.标识号, 0);
      Zljsonputvalue(v_Output, 'pati_pageid', r_销帐.主页id, 1, 2);
    End Loop;
  End If;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := Null;
  End If;
  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"成功","fee_cancel_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"成功","fee_cancel_list":[' || v_Output || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getchargeoffinfo;
/


Create Or Replace Procedure Zl_Exsesvr_Getbilldetailinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) Is
  -------------------------------------------------------------------------------------------------
  --功能：获取和药品发药业务相关的费用信息，主要用于界面显示
  --入参：json格式
  --Input
  --   fee_ids    C     费用id，支持多个id，格式： 费用id,费用id,…
  --   bill_nos   C     费用no,记录性质，格式: no,记录性质|,...
  --出参：json格式
  --Json_Out
  --fee_list      [数组]每个费用ID信息
  --  bill_prop           N    记录性质:1-收费单;2-记帐单;3-自动记帐单;4-挂号单;5-就诊卡;6-预交单
  --  bill_no             C    单据号
  --  fee_id              N    处方明细id(费用id)
  --  fee_num             N    序号
  --  iden_id             N    标识号
  --  pati_bed            C    床号
  --  fee_ampaid          N    实收金额
  --  packages_num        N    付数
  --  quantity            N    数次
  --  placer              C    开单人
  --  operator_code       C    操作员编号
  --  operator_name       C    操作员姓名
  --  create_time         D    登记时间
  --  happen_time         D    发生时间
  --  rcp_type            N    处方类别(按整个NO来说，1-西药，2-中药，3-混合)
  --  fee_type            C    费别
  --  rec_status          N    记录状态
  --  register_id         N    挂号id
  --  register_no         C    挂号NO
  --  register_time       D    挂号登记时间
  --  income_item_id      N    收入项目id
  --  fee_origin          N    费用来源(1-门诊费用，2-住院费用)
  --  bill_deptid         N    开单部门id
  --  order_id            N    医嘱ID
  --  fee_item_id         N    收费细目id
  --  fee_status         N    费用状态
  -------------------------------------------------------------------------------------------------
  v_Output Varchar2(32767);
  c_Output Clob;

  v_费用id Varchar2(32767); --费用id
  v_Nos    Varchar2(32767);

  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  l_费用id Collection_Type;
  l_No     Collection_Type;

  j_Input Pljson;
  j_Json  Pljson;
Begin
  j_Input := Pljson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  If j_Json.Exist('bill_nos') Then
    v_Nos := j_Json.Get_String('bill_nos');
    --将 v_Nos 串组装成不超过4000 的集合串，防止使用 f_Str2list 参数超长
    While v_Nos Is Not Null Loop
      If Length(v_Nos) <= 4000 Then
        l_No(l_No.Count) := v_Nos;
        v_Nos := Null;
      Else
        l_No(l_No.Count) := Substr(v_Nos, 1, Instr(v_Nos, '|', 3980) - 1);
        v_Nos := Substr(v_Nos, Instr(v_Nos, '|', 3980) + 1);
      End If;
    End Loop;
  
    For I In 0 .. l_No.Count - 1 Loop
      For r_费用 In (Select a.记录性质, a.No, a.Id, a.序号, a.标识号, '' As 床号, a.实收金额, a.付数, a.数次, a.开单人, a.操作员编号, a.操作员姓名, a.登记时间,
                          a.发生时间, Zl_Get收费类别(a.记录性质, a.No, a.执行部门id) As 收费类别, a.费别, a.记录状态, a.收入项目id, a.挂号id,
                          b.No As 挂号no, b.登记时间 As 挂号登记时间, 1 As 费用来源, a.开单部门id, a.医嘱序号, a.收费细目id, a.费用状态
                   
                   From 门诊费用记录 A, 病人挂号记录 B,
                        (Select /*+cardinality(J,10)*/
                           C1 As NO, C2 As 记录性质
                          From Table(f_Str2list2(l_No(I), '|', ',')) J) C
                   Where a.挂号id = b.Id(+) And a.No = c.No And (a.记录性质 = c.记录性质 Or Nvl(c.记录性质, 0) = 0)
                   Union All
                   Select a.记录性质, a.No, a.Id, a.序号, a.标识号, 床号, a.实收金额, a.付数, a.数次, a.开单人, a.操作员编号, a.操作员姓名, a.登记时间,
                          a.发生时间, Zl_Get收费类别(a.记录性质, a.No, a.执行部门id) As 收费类别, a.费别, a.记录状态, a.收入项目id, 0 As 挂号id,
                          '' As 挂号no, Null As 挂号登记时间, 2 As 费用来源, a.开单部门id, a.医嘱序号, a.收费细目id, a.费用状态
                   From 住院费用记录 A,
                        (Select /*+cardinality(J,10)*/
                           C1 As NO, C2 As 记录性质
                          From Table(f_Str2list2(l_No(I), '|', ',')) J) C
                   Where a.No = c.No And (a.记录性质 = c.记录性质 Or Nvl(c.记录性质, 0) = 0)) Loop
      
        If Length(Nvl(v_Output, ' ')) > 30000 Then
          If c_Output Is Not Null Then
            c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
          Else
            c_Output := To_Clob(v_Output);
          End If;
          v_Output := Null;
        End If;
      
        Zljsonputvalue(v_Output, 'bill_prop', r_费用.记录性质, 1, 1);
        Zljsonputvalue(v_Output, 'bill_no', r_费用.No, 0);
        Zljsonputvalue(v_Output, 'fee_id', r_费用.Id, 1);
        Zljsonputvalue(v_Output, 'fee_num', r_费用.序号, 1);
        Zljsonputvalue(v_Output, 'iden_id', r_费用.标识号, 1);
        Zljsonputvalue(v_Output, 'pati_bed', r_费用.床号, 0);
        Zljsonputvalue(v_Output, 'fee_ampaid', r_费用.实收金额, 1);
        Zljsonputvalue(v_Output, 'packages_num', Nvl(r_费用.付数, 1), 1);
        Zljsonputvalue(v_Output, 'quantity', r_费用.数次, 1);
        Zljsonputvalue(v_Output, 'placer', r_费用.开单人, 0);
        Zljsonputvalue(v_Output, 'operator_code', r_费用.操作员编号, 0);
        Zljsonputvalue(v_Output, 'operator_name', r_费用.操作员姓名, 0);
        Zljsonputvalue(v_Output, 'create_time', To_Char(r_费用.登记时间, 'yyyy-mm-dd hh24:mi:ss'), 0);
        Zljsonputvalue(v_Output, 'happen_time', To_Char(r_费用.发生时间, 'yyyy-mm-dd hh24:mi:ss'), 0);
        Zljsonputvalue(v_Output, 'rcp_type', r_费用.收费类别, 1);
        Zljsonputvalue(v_Output, 'fee_type', r_费用.费别, 0);
        Zljsonputvalue(v_Output, 'rec_status', r_费用.记录状态, 1);
        Zljsonputvalue(v_Output, 'register_id', r_费用.挂号id, 1);
        Zljsonputvalue(v_Output, 'register_no', r_费用.挂号no, 0);
        Zljsonputvalue(v_Output, 'register_time', To_Char(r_费用.挂号登记时间, 'yyyy-mm-dd hh24:mi:ss'), 0);
        Zljsonputvalue(v_Output, 'income_item_id', r_费用.收入项目id, 1);
        Zljsonputvalue(v_Output, 'fee_origin', r_费用.费用来源, 1);
        Zljsonputvalue(v_Output, 'bill_deptid', r_费用.开单部门id, 1);
        Zljsonputvalue(v_Output, 'order_id', r_费用.医嘱序号, 1);
        Zljsonputvalue(v_Output, 'fee_item_id', r_费用.收费细目id, 1);
        Zljsonputvalue(v_Output, 'fee_status', r_费用.费用状态, 1, 2);
      
      End Loop;
    End Loop;
  
  Else
  
    v_费用id := j_Json.Get_String('fee_ids');
    --将 v_费用id 串组装成不超过4000 的集合串，防止使用 f_Num2list 参数超长
    While v_费用id Is Not Null Loop
      If Length(v_费用id) <= 4000 Then
        l_费用id(l_费用id.Count) := v_费用id;
        v_费用id := Null;
      Else
        l_费用id(l_费用id.Count) := Substr(v_费用id, 1, Instr(v_费用id, ',', 3980) - 1);
        v_费用id := Substr(v_费用id, Instr(v_费用id, ',', 3980) + 1);
      End If;
    End Loop;
  
    For I In 0 .. l_费用id.Count - 1 Loop
      For r_费用 In (Select /*+cardinality(J,10)*/
                    a.记录性质, a.No, a.Id, a.序号, a.标识号, '' As 床号, a.实收金额, a.付数, a.数次, a.开单人, a.操作员编号, a.操作员姓名, a.登记时间,
                    a.发生时间, Decode(a.收费类别, '7', 2, 1) As 收费类别, a.费别, a.记录状态, a.收入项目id, a.挂号id, b.No As 挂号no,
                    b.登记时间 As 挂号登记时间, 1 As 费用来源, a.开单部门id, a.医嘱序号, a.收费细目id, a.费用状态
                   From 门诊费用记录 A, 病人挂号记录 B, Table(f_Num2list(l_费用id(I))) J
                   Where a.挂号id = b.Id(+) And a.Id = j.Column_Value
                   Union All
                   Select /*+cardinality(J,10)*/
                    a.记录性质, a.No, a.Id, a.序号, a.标识号, 床号, a.实收金额, a.付数, a.数次, a.开单人, a.操作员编号, a.操作员姓名, a.登记时间, a.发生时间,
                    Decode(a.收费类别, '7', 2, 1) As 收费类别, a.费别, a.记录状态, a.收入项目id, 0 As 挂号id, '' As 挂号no, Null As 挂号登记时间,
                    2 As 费用来源, a.开单部门id, a.医嘱序号, a.收费细目id, a.费用状态
                   From 住院费用记录 A, Table(f_Num2list(l_费用id(I))) J
                   Where a.Id = j.Column_Value) Loop
      
        If Length(Nvl(v_Output, ' ')) > 30000 Then
          If c_Output Is Not Null Then
            c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
          Else
            c_Output := To_Clob(v_Output);
          End If;
          v_Output := Null;
        End If;
      
        Zljsonputvalue(v_Output, 'bill_prop', r_费用.记录性质, 1, 1);
        Zljsonputvalue(v_Output, 'bill_no', r_费用.No, 0);
        Zljsonputvalue(v_Output, 'fee_id', r_费用.Id, 1);
        Zljsonputvalue(v_Output, 'fee_num', r_费用.序号, 1);
        Zljsonputvalue(v_Output, 'iden_id', r_费用.标识号, 1);
        Zljsonputvalue(v_Output, 'pati_bed', r_费用.床号, 0);
        Zljsonputvalue(v_Output, 'fee_ampaid', r_费用.实收金额, 1);
        Zljsonputvalue(v_Output, 'packages_num', Nvl(r_费用.付数, 1), 1);
        Zljsonputvalue(v_Output, 'quantity', r_费用.数次, 1);
        Zljsonputvalue(v_Output, 'placer', r_费用.开单人, 0);
        Zljsonputvalue(v_Output, 'operator_code', r_费用.操作员编号, 0);
        Zljsonputvalue(v_Output, 'operator_name', r_费用.操作员姓名, 0);
        Zljsonputvalue(v_Output, 'create_time', To_Char(r_费用.登记时间, 'yyyy-mm-dd hh24:mi:ss'), 0);
        Zljsonputvalue(v_Output, 'happen_time', To_Char(r_费用.发生时间, 'yyyy-mm-dd hh24:mi:ss'), 0);
        Zljsonputvalue(v_Output, 'rcp_type', r_费用.收费类别, 1);
        Zljsonputvalue(v_Output, 'fee_type', r_费用.费别, 0);
        Zljsonputvalue(v_Output, 'rec_status', r_费用.记录状态, 1);
        Zljsonputvalue(v_Output, 'register_id', r_费用.挂号id, 1);
        Zljsonputvalue(v_Output, 'register_no', r_费用.挂号no, 0);
        Zljsonputvalue(v_Output, 'register_time', To_Char(r_费用.挂号登记时间, 'yyyy-mm-dd hh24:mi:ss'), 0);
        Zljsonputvalue(v_Output, 'income_item_id', r_费用.收入项目id, 1);
        Zljsonputvalue(v_Output, 'fee_origin', r_费用.费用来源, 1);
        Zljsonputvalue(v_Output, 'bill_deptid', r_费用.开单部门id, 1);
        Zljsonputvalue(v_Output, 'order_id', r_费用.医嘱序号, 1);
        Zljsonputvalue(v_Output, 'fee_item_id', r_费用.收费细目id, 1);
        Zljsonputvalue(v_Output, 'fee_status', r_费用.费用状态, 1, 2);
      End Loop;
    End Loop;
  
  End If;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"成功","fee_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"成功","fee_list":[' || v_Output || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getbilldetailinfo;
/

Create Or Replace Procedure Zl_Exsesvr_Getbillinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) Is
  -------------------------------------------------------------------------------------------------
  --功能：获取和药品发药业务相关的费用信息，主要用于界面显示
  --入参：json格式
  --Input
  --   pharmacy_id：库房id
  --   fee_nos：费用no，支持多个no，格式： 记录性质,no,…
  --出参：json格式
  --Json_Out
  --  code          C   1   应答码：0-失败；1-成功
  --  message       C   1   应答消息：
  --  fee_list      C       [数组]每个费用NO信息
  --    fee_properties      N 记录性质
  --    bill_no             C 费用no
  --    real_amount         N 实收金额
  --    rcp_type            N 收费类别(按整个NO来说，1-西药，2-中药，3-混合)
  --    iden_id             C 标识号
  --    placer              C 开单人
  --    bill_deptid         N 开单部门id
  --    create_time         D 登记时间
  --    pati_bed            C 当前床号
  --    operator_name       C 操作员姓名
  -------------------------------------------------------------------------------------------------
  n_库房id 门诊费用记录.执行部门id%Type;
  v_费用no Varchar2(32767);

  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_费用id Collection_Type;
  I          Number;

  v_Output Varchar2(32767);
  c_Output Clob;
  j_Input  PLJson;
  j_Json   PLJson;

Begin

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_库房id := j_Json.Get_Number('pharmacy_id');
  v_费用no := j_Json.Get_String('fee_nos');
  I        := 0;
  While v_费用no Is Not Null Loop
    If Length(v_费用no) <= 4000 Then
      Col_费用id(I) := v_费用no;
      v_费用no := Null;
    Else
      Col_费用id(I) := Substr(v_费用no, 1, Instr(v_费用no, '|', 3980) - 1);
      v_费用no := Substr(v_费用no, Instr(v_费用no, '|', 3980) + 1);
    End If;
    I := I + 1;
  End Loop;
  I := 0;

  For I In 0 .. Col_费用id.Count - 1 Loop
  
    For v_费用信息 In (Select /*+ rule*/
                    NO, 记录性质, Sum(实收金额) As 实收金额, 收费类别, Max(标识号) As 标识号, Max(开单人) As 开单人, Max(开单部门id) As 开单部门id,
                    Max(登记时间) As 登记时间, Max(床号) As 床号, Max(操作员姓名) As 操作员姓名
                   From (Select a.No, a.记录性质, a.实收金额, Zl_Get收费类别(a.记录性质, a.No, a.执行部门id) As 收费类别, a.标识号, a.开单人, a.开单部门id,
                                 Decode(Nvl(a.记录状态, 0), 2, To_Date(Null), a.登记时间) As 登记时间, '' As 床号,
                                 Decode(Nvl(a.记录状态, 0), 2, Null, a.操作员姓名) As 操作员姓名
                          From 门诊费用记录 A,
                               (Select /*+cardinality(c,10)*/
                                  C1 As 记录性质, C2 As NO
                                 From Table(f_Str2List2(Col_费用id(I), '|', ',')) C) C
                          Where a.执行部门id = Decode(Nvl(n_库房id, 0), 0, a.执行部门id, n_库房id) And Mod(a.记录性质, 10) = c.记录性质 And
                                a.No = c.No And a.记录性质 In (1, 2)
                          Union All
                          Select a.No, a.记录性质, a.实收金额, Zl_Get收费类别(a.记录性质, a.No, a.执行部门id) As 收费类别,
                                 Decode(Nvl(多病人单, 0), 1, -1 * Null, 标识号) As 标识号, a.开单人, a.开单部门id,
                                 Decode(Nvl(a.记录状态, 0), 2, To_Date(Null), a.登记时间) As 登记时间,
                                 Decode(Nvl(多病人单, 0), 1, '', 床号) As 床号, Decode(Nvl(a.记录状态, 0), 2, Null, a.操作员姓名) As 操作员姓名
                          From 住院费用记录 A,
                               (Select /*+cardinality(c,10)*/
                                  C1 As 记录性质, C2 As NO
                                 From Table(f_Str2List2(Col_费用id(I), '|', ',')) C) C
                          Where a.执行部门id = Decode(Nvl(n_库房id, 0), 0, a.执行部门id, n_库房id) And Mod(a.记录性质, 10) = c.记录性质 And
                                a.No = c.No And a.记录性质 In (1, 2))
                   Group By NO, 记录性质, 收费类别) Loop
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      
        v_Output := Null;
      End If;
      zlJsonPutValue(v_Output, 'fee_properties', v_费用信息.记录性质, 1, 1);
      zlJsonPutValue(v_Output, 'bill_no', v_费用信息.No, 0, 0);
      zlJsonPutValue(v_Output, 'real_amount', v_费用信息.实收金额, 1, 0);
      zlJsonPutValue(v_Output, 'rcp_type', v_费用信息.收费类别, 1, 0);
      zlJsonPutValue(v_Output, 'iden_id', v_费用信息.标识号, 0, 0);
      zlJsonPutValue(v_Output, 'placer', v_费用信息.开单人, 0, 0);
      zlJsonPutValue(v_Output, 'bill_deptid', v_费用信息.开单部门id, 1, 0);
      zlJsonPutValue(v_Output, 'create_time', To_Char(v_费用信息.登记时间, 'yyyy-mm-dd hh24:mi:ss'), 0, 0);
      zlJsonPutValue(v_Output, 'pati_bed', v_费用信息.床号, 0, 0);
      zlJsonPutValue(v_Output, 'operator_name', v_费用信息.操作员姓名, 0, 2);
    End Loop;
  
  End Loop;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"成功","fee_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"成功","fee_list":[' || v_Output || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getbillinfo;
/
Create Or Replace Procedure Zl_Exsesvr_Getinsureiteminfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：获取医保大类相关信息，如获取名称，是否设置了保险支付项目
  --入参：Json_In:格式
  --  input
  --    insurance_type          N 1 险类
  --    fee_item_id             N 1 收费细目ID
  --                                以下为批量获取的入参条件
  --    fee_item_ids            C 0 收费细目ids，取一批收费项目的医保大类名称
  --    insurance_types         C 0 险类逗号拼串
  --    query_type              N 0 查询方式
  --                                   0-传入fee_item_ids+insurance_type取一批收费项目的医保大类名称
  --                                   1-传入fee_item_ids+insurance_type返回设置了的fee_item_ids 返回一批设置了保险支付项目ids，
  --                                   2-传入fee_item_ids+insurance_types返回设置了的fee_item_ids和insurance_type列表
  --出参: Json_Out,格式如下
  --  output
  --    code                    N 1 应答吗：0-失败；1-成功
  --    message                 C 1 应答消息：失败时返回具体的错误信息
  --    insure_name             C 1 保险大类名称
  --    isexist                 N 1 否设置了保险支付项目,0-未设置，1-设置了
  --    fee_item_ids            C 1 设置了保险支付项目的收费细目id拼串
  --    item_list[]批量获取时才返回
  --           fee_item_id      N 1 收费细目ID
  --           insure_name      C 1 保险大类名称
  --           insure_name_ex   C 1 组合名称临床诊疗选择器用
  --    pay_list[]设置了保险支付项目列表query_type=2时返回
  --           insurance_type   N 1 险类
  --           fee_item_ids     C 1 设置了保险支付项目ids
  ---------------------------------------------------------------------------
  j_Json       Pljson;
  j_Input      Pljson;
  v_Tmp        Varchar2(32767);
  n_Tmp        Number;
  n_收费细目id Number;
  n_险类       Number;
  v_险类s      Varchar2(32767);
  v_Vals       Clob;
  l_Vals       t_Strlist;
  n_查询方式   Number;
  v_Jtmp       Varchar2(32767);
  c_Jtmp       Clob;

Begin
  --解析入参
  j_Input    := Pljson(Json_In);
  j_Json     := j_Input.Get_Pljson('input');
  n_险类     := j_Json.Get_Number('insurance_type');
  n_查询方式 := j_Json.Get_Number('query_type');

  If j_Json.Exist('fee_item_ids') Then
    v_Vals := j_Json.Get_Clob('fee_item_ids');
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
  
    If 0 = Nvl(n_查询方式, 0) Then
      For Lp In 1 .. l_Vals.Count Loop
        For R In (Select m.收费细目id, n.名称 || Decode(m.保险费用等级, Null, Null, '(' || m.保险费用等级 || ')') As 医保大类, n.名称
                  From 保险支付项目 M, 保险支付大类 N
                  Where m.项目编码 Is Not Null And m.大类id = n.Id And m.险类 = n_险类 And
                        m.收费细目id In (Select /*+cardinality(b,10)*/
                                      b.Column_Value As 收费细目id
                                     From Table(f_Num2list(l_Vals(Lp))) B)
                  Group By m.收费细目id, n.名称, m.保险费用等级) Loop
        
          v_Jtmp := v_Jtmp || ',{"fee_item_id":' || r.收费细目id;
          v_Jtmp := v_Jtmp || ',"insure_name":"' || Zljsonstr(r.名称) || '"';
          v_Jtmp := v_Jtmp || ',"insure_name_ex":"' || Zljsonstr(r.医保大类) || '"';
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
    
    Elsif 1 = n_查询方式 Then
    
      For Lp In 1 .. l_Vals.Count Loop
        For R In (Select Distinct 收费细目id
                  From 保险支付项目
                  Where 项目编码 Is Not Null And
                        收费细目id In (Select /*+cardinality(b,10)*/
                                    b.Column_Value As 收费细目id
                                   From Table(f_Num2list(l_Vals(Lp))) B) And 险类 = n_险类) Loop
          v_Jtmp := v_Jtmp || ',' || r.收费细目id;
        
          If Length(v_Jtmp) > 32000 Then
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
        Json_Out := '{"output":{"code":1,"message":"成功","fee_item_ids":"' || Substr(v_Jtmp, 2) || '"}}';
      Else
        c_Jtmp   := c_Jtmp || v_Jtmp;
        Json_Out := '{"output":{"code":1,"message":"成功","fee_item_ids":"' || c_Jtmp || '"}}';
      End If;
    
    Elsif 2 = n_查询方式 Then
    
      v_险类s := j_Json.Get_String('insurance_types');
      v_Tmp   := Null;
      n_险类  := Null;
    
      For Lp In 1 .. l_Vals.Count Loop
        For R In (Select a.险类, a.收费细目id
                  From 保险支付项目 A
                  Where a.项目编码 Is Not Null And Nvl(a.险类, 0) <> 0 And
                        a.收费细目id In (Select /*+cardinality(b,10)*/
                                      b.Column_Value As 收费细目id
                                     From Table(f_Num2list(l_Vals(Lp))) B) And
                        a.险类 In (Select /*+cardinality(x,10)*/
                                  x.Column_Value
                                 From Table(f_Num2list(v_险类s)) X)
                  Group By a.险类, a.收费细目id
                  Order By a.险类, a.收费细目id) Loop
        
          If n_险类 <> r.险类 And n_险类 Is Not Null Then
          
            v_Jtmp := v_Jtmp || ',{"insurance_type":' || n_险类;
            v_Jtmp := v_Jtmp || ',"fee_item_ids":"' || Substr(v_Tmp, 2) || '"';
            v_Jtmp := v_Jtmp || '}';
          
            If Length(v_Jtmp) > 30000 Then
              If c_Jtmp Is Null Then
                c_Jtmp := Substr(v_Jtmp, 2);
              Else
                c_Jtmp := c_Jtmp || v_Jtmp;
              End If;
              v_Jtmp := Null;
            End If;
          
            v_Tmp := Null;
          End If;
          n_险类 := r.险类;
          v_Tmp  := v_Tmp || ',' || r.收费细目id;
        End Loop;
      End Loop;
    
      --最末一次
      If n_险类 Is Not Null Then
        v_Jtmp := v_Jtmp || ',{"insurance_type":' || n_险类;
        v_Jtmp := v_Jtmp || ',"fee_item_ids":"' || Substr(v_Tmp, 2) || '"';
        v_Jtmp := v_Jtmp || '}';
      
        If Length(v_Jtmp) > 30000 Then
          If c_Jtmp Is Null Then
            c_Jtmp := Substr(v_Jtmp, 2);
          Else
            c_Jtmp := c_Jtmp || v_Jtmp;
          End If;
          v_Jtmp := Null;
        End If;
      End If;
      If c_Jtmp Is Null Then
        Json_Out := '{"output":{"code":1,"message":"成功","pay_list":[' || Substr(v_Jtmp, 2) || ']}}';
      Else
        c_Jtmp   := c_Jtmp || v_Jtmp;
        Json_Out := '{"output":{"code":1,"message":"成功","pay_list":[' || c_Jtmp || ']}}';
      End If;
    End If;
  Else
    n_收费细目id := j_Json.Get_Number('fee_item_id');
    Select Max(n.名称)
    Into v_Tmp
    From 保险支付项目 M, 保险支付大类 N
    Where m.项目编码 Is Not Null And m.收费细目id = n_收费细目id And m.大类id = n.Id And m.险类 = n_险类;
    Select Count(1) Into n_Tmp From 保险支付项目 Where 收费细目id = n_收费细目id And 险类 = n_险类;
    Json_Out := '{"output":{"code":1,"message":"成功","insure_name":"' || v_Tmp || '","isexist":' || n_Tmp || '}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Exsesvr_Getinsureiteminfo;
/
 
CREATE OR REPLACE Procedure Zl_Exsesvr_Delbill
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：门诊医嘱作废，住院医嘱回退发送，删除费用单据
  --入参：Json_In:格式
  --input
  --   operator_name         C  1 操作员姓名【记帐单删除时传入】
  --   operator_code         C  1 操作员编号【记帐单删除时传入】
  --   operator_time         C  1 操作时间:yyyy-mm-dd hh:mi:ss【记帐单删除时传入】
  --   del_list  直接删除的单据列表
  --             fee_source          N 1 费用来源:1-门诊费用记录;2-住院费用记录
  --             fee_bill_type       N 1 记录性质，1-收费单，2-记帐单
  --             fee_no              C 1 费用单据号
  --             del_type            N   退费方式:0-按序号串删除费用；1-按费用id串删除费用;2-全退
  --             serial_num          C   序号串,query_type=0时有效，
  --                                             记帐单格式: 序号1:数量:执行状态1,序号2:数量2:执行状态2,...
  --                                                 格式说明：执行状态:0-未执行;1-完全执行;2-部分执行
  --                                             收费单格式：序号1,序号2,序号3...
  --
  --             exe_sta_nums        C   需要先取消执行的项目，格式:序号1,序号2,序号3...
  --             fee_ids             C   费用id串，query_type=1时有效
  --                                               记帐单格式: id1:数量:执行状态1,id2:数量2:执行状态2,...
  --                                                   格式说明：执行状态:0-未执行;1-完全执行;2-部分执行
  --                                               收费单格式：id1,id2,id3...
  --             oper_status         N 1 操作状态，住院记帐单删时才传入，0-表示直接销帐;1-表示审核销帐(通过销帐申请-->销帐审核流程);2-表示转病区费用
  --出参: Json_Out,格式如下
  --output
  --    code          N 1 应答吗：0-失败；1-成功
  --    message       C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  --说明：入参确定为Clob原因：
  --        1.在静配中批量取消配药时删除附费时入参长度可能会超过32767；
  --        2.入参类型Clob调用一次比Varchar2平均固定只多10ms，且调用不频繁
  j_Input Pljson;
  j_Json  Pljson;

  j_List Pljson_List;
  j_Item Pljson;

  v_编号 住院费用记录.操作员编号%Type;
  v_人员 住院费用记录.操作员姓名%Type;
  d_时间 住院费用记录.登记时间%Type;

  n_来源 Number(2);
  n_性质 Number(2);

  v_No       住院费用记录.No%Type;
  v_序号销帐 Varchar2(32767);
  v_序号执行 Varchar2(32767);
  v_费用ID销帐 Varchar2(32767);
  v_费用ID   varchar2(4000);
  v_当前序号 varchar2(4000);
  n_操作状态 Number;
  n_退费方式 Number(2);

Begin
  --解析入参
  j_Input := Pljson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');
  v_编号  := j_Json.Get_String('operator_code');
  v_人员  := j_Json.Get_String('operator_name');
  d_时间  := To_Date(j_Json.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');

  j_List := j_Json.Get_Pljson_List('del_list');
  For J In 1 .. j_List.Count Loop
    j_Item     := Pljson();
    j_Item     := Pljson(j_List.Get(J));
    n_来源     := j_Item.Get_Number('fee_source');
    n_性质     := j_Item.Get_Number('fee_bill_type');
    v_No       := j_Item.Get_String('fee_no');
    v_序号执行 := j_Item.Get_String('exe_sta_nums');
    n_操作状态 := j_Item.Get_Number('oper_status');
    n_退费方式 := j_Item.Get_Number('del_type');
    If Nvl(n_退费方式, 0) = 0 Then
      v_序号销帐 := j_Item.Get_String('serial_num');
    Elsif Nvl(n_退费方式, 0) = 2 Then
      --全退,按费用no退
      v_序号销帐 := Null;
    Else
      --按费用ID退
      --将费用ID转换为序号
      v_费用id销帐 := j_Item.Get_String('fee_ids');
      If v_费用id销帐 Is Null Then
        Json_Out := zlJsonOut('未传入需要销帐的费用id');
        Return;
      End If;
      If n_性质 = 1 Then
        v_序号销帐 := v_费用id销帐;
      Else
        --记帐单
        v_费用ID销帐 := v_费用ID销帐 || ',';
        While v_费用ID销帐 Is Not Null Loop
          v_费用ID := Substr(v_费用ID销帐, 1, Instr(v_费用ID销帐, ',', 3940) - 1);
          If n_来源 = 1 Then
            Select /*+cardinality(b,10)*/
             f_List2str(Cast(Collect(a.序号 || ':' || B.C2) As t_Strlist))
            Into v_当前序号
            From 门诊费用记录 a, Table(f_Str2list2(v_费用ID)) b
            Where a.Id = b.C1 And a.No = v_No And a.记录性质 = 2;
          Else
            Select /*+cardinality(b,10)*/
             f_List2str(Cast(Collect(a.序号 || ':' || B.C2) As t_Strlist))
            Into v_当前序号
            From 住院费用记录 a, Table(f_Str2list2(v_费用ID)) b
            Where a.Id = b.C1 And a.No = v_No And a.记录性质 = 2;
          End If;

          v_序号销帐 := v_序号销帐 || ',' || v_当前序号;
          v_费用ID销帐 := Substr(v_费用ID销帐, Instr(v_费用ID销帐, ',', 3940) + 1);
        End Loop;
        If v_序号销帐 Is Not Null Then
          v_序号销帐 := substr(v_序号销帐, 2);
        End If;
      End If;
    End If;

    --因为费用退费过程中检查了执行状态，需要先修正费用执行状态
    --针对本科自动执行完成的医嘱：回退医嘱发送、门诊作废医嘱时需要先取消执行完成，再退费
    If v_序号执行 Is Not Null Or Nvl(n_退费方式, 0) = 2 Then
      If n_来源 = 2 Then
        Update 住院费用记录
        Set 执行状态 = 0, 执行时间 = Null, 执行人 = Null
        Where NO = v_No And (Instr(',' || v_序号执行 || ',', ',' || Nvl(价格父号, 序号) || ',') > 0 Or Nvl(n_退费方式, 0) = 2) And
              Mod(记录性质, 10) = n_性质 And 记录状态 In (0, 1, 3);
      Else
        Update 门诊费用记录
        Set 执行状态 = 0, 执行时间 = Null, 执行人 = Null
        Where NO = v_No And (Instr(',' || v_序号执行 || ',', ',' || Nvl(价格父号, 序号) || ',') > 0 Or Nvl(n_退费方式, 0) = 2) And
              Mod(记录性质, 10) = n_性质 And 记录状态 In (0, 1, 3);
      End If;
    End If;

    If n_来源 = 1 Then
      --门诊
      If n_性质 = 1 Then
        --门诊划价
        Zl_门诊划价记录_Delete_s(v_No, v_序号销帐, n_退费方式);
      Else
        --门诊记帐
        Zl_门诊记帐记录_Delete_s(v_No, v_序号销帐, v_编号, v_人员, d_时间, 2);
      End If;
    Else
      --住院
      Zl_住院记帐记录_Delete_s(v_No, v_序号销帐, v_编号, v_人员, 2, Nvl(n_操作状态, 0), d_时间);
    End If;
  End Loop;
  Json_Out := Zljsonout('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Delbill;
/

Create Or Replace Procedure Zl_Exsesvr_Billverify
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：费用单据审核
  --入参：Json_In:格式
  --  input
  --    operator_time         C 1 操作时间:yyyy-mm-dd hh24:mi:ss
  --    operator_name         C 1 操作员姓名
  --    operator_code         C 1 操作员编号
  --    item_list
  --        fee_source        N 1 费用来源:1-门诊;2-住院
  --        fee_no            C 1 费用单据号
  --        serial_nums       C 0 序号串，不传表示整张单据
  --        pharmacy_window   C 0 发药窗口，费用来源为门诊时传入，格式：库房ID1:发药窗口1,库房ID2:发药窗口2,....
  --        pati_id           N 0 病人id，费用来源为住院且按病人审核时传入(主要针对记帐表)
  --出参: Json_Out,格式如下
  --  output
  --    code          N 1 应答吗：0-失败；1-成功
  --    message       C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  j_Item PLJson;
  j_List Pljson_List;

  v_人员 Varchar2(300);
  v_编号 Varchar2(300);
  d_时间 Date;

  n_来源     Number(1); --1-门诊;2-住院
  v_No       门诊费用记录.No%Type;
  v_序号     Varchar2(32767);
  v_发药窗口 Varchar2(32767);
  n_病人id   门诊费用记录.病人id%Type;
Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_人员 := j_Json.Get_String('operator_name');
  v_编号 := j_Json.Get_String('operator_code');
  d_时间 := To_Date(j_Json.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');

  j_List := j_Json.Get_Pljson_List('item_list');
  For I In 1 .. j_List.Count Loop
    j_Item     := PLJson();
    j_Item     := PLJson(j_List.Get(I));
    n_来源     := j_Item.Get_Number('fee_source');
    v_No       := j_Item.Get_String('fee_no');
    v_序号     := j_Item.Get_String('serial_nums');
    v_发药窗口 := j_Item.Get_String('pharmacy_window');
    n_病人id   := j_Item.Get_Number('pati_id');
  
    If n_来源 = 1 Then
      Zl_门诊记帐记录_Verify_s(v_No, v_编号, v_人员, v_序号, d_时间, v_发药窗口, 0);
    Else
      Zl_住院记帐记录_Verify_s(v_No, v_编号, v_人员, v_序号, n_病人id, d_时间, 0);
    End If;
  End Loop;

  Json_Out := zlJsonOut('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Billverify;
/


  
Create Or Replace Procedure Zl_Exsesvr_Getnextid
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
  --  quantity      N  0 所需序列的个数，如果只取一个该参不传或都传0 
  -- 出参:
  --  output
  --  next_id      C   1  序列，quantity>1 时，返回多个序号，用逗号分离
  -------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_Table  Varchar2(500);
  v_Col    Varchar2(500);
  n_Nextid Number;
  n_数量   Number;
  v_Ids    Varchar2(32767);
  v_Sql    Varchar2(4000);
  --动态游标类型
  Type Rs_Recordset Is Ref Cursor;
  c_Tmp Rs_Recordset;
Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_Table := j_Json.Get_String('table_name');
  v_Col   := Nvl(j_Json.Get_String('col_name'), 'ID');
  n_数量  := j_Json.Get_Number('quantity');

  If Nvl(n_数量, 0) > 1 Then
    v_Sql := 'Select ' || v_Table || '_' || v_Col || '.Nextval as 序列 From Dual Connect By Level <= :1';
    Open c_Tmp For v_Sql
      Using In n_数量;
  
    v_Ids := Null;
    Loop
      Fetch c_Tmp
        Into n_Nextid;
      Exit When c_Tmp%NotFound;
      If c_Tmp%RowCount > 0 Then
        v_Ids := v_Ids || ',' || n_Nextid;
      End If;
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"成功","next_id":"' || Substr(v_Ids, 2) || '"}}';
    Return;
  End If;

  Execute Immediate 'select ' || v_Table || '_' || v_Col || '.nextval from dual'
    Into n_Nextid;
  Json_Out := '{"output":{"code":1,"message":"成功","next_id":"' || n_Nextid || '"}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getnextid;
/

CREATE OR REPLACE Procedure Zl_Exsesvr_Newbill_Check
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：记帐插入时，先进行相关数据的合法性检查
  --入参：Json_In:格式
  --input
  --        pati_id             N 1  病人ID
  --        pati_pageid         N 1  主页Id
  --        pati_deptid         N 1  病人科室 id
  --        pati_wardarea_id    N 1  病人病区iD  
  --        pati_name           C 1  病人姓名
  --        fee_audit_status    N 1  费用审核标志:0或空-未审核;1-已审核或开始审核(结合参数:病人审核方式来控制);2-完成审核,结合结帐权限[禁止未审核病人结帐]进行管理控制
  --        si_inp_status       N 1  住院状态:0-正常住院;1-尚未入科;2-正在转科或正在转病区;3-已预出院

  --        dept_list[]开单科室ID和领用部门ID
  --                            plcdept_id          N 1  开单科室ID
  --                            takedept_id         N 1  领用部门ID就是 药品的领药部门id
  --出参: Json_Out,格式如下
  --
  --    output
  --        code                    N   1   应答吗：0-失败；1-成功
  --        message                 C   1   应答消息：失败时返回具体的错误信息
  --        item_list[]                         数据组
  --            pati_id             N   1    病人ID
  --            takedept_id             N   1   领用部门ID
  ---------------------------------------------------------------------------
  j_Input   Pljson;
  j_Json    Pljson;
  j_Item    Pljson;
  j_List    Pljson_List := Pljson_List();
  v_Jtmp_In Varchar2(4000);
  v_Jpati   Varchar2(3000);
  n_病人id  住院费用记录.病人id%Type;
  n_主页id  住院费用记录.主页id%Type;
  v_姓名    住院费用记录.姓名%Type;

  n_病人科室id 部门表.Id%Type;
  n_病人病区id 部门表.Id%Type;
  n_开单科室id 住院费用记录.开单部门id%Type;

  n_领药部门id 部门表.Id%Type;

  n_费用审核标志 Number(2);
  n_住院状态     Number(2);
Begin

  --解析入参
  j_Input := Pljson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_病人id       := j_Json.Get_Number('pati_id');
  n_主页id       := j_Json.Get_Number('pati_pageid');
  n_费用审核标志 := j_Json.Get_Number('fee_audit_status');
  n_住院状态     := j_Json.Get_Number('si_inp_status');
  n_病人科室id   := j_Json.Get_Number('pati_deptid');
  n_病人病区id   := j_Json.Get_Number('pati_wardarea_id');
  v_姓名         := j_Json.Get_String('pati_name');

  j_List := j_Json.Get_Pljson_List('dept_list');
  For I In 1 .. j_List.Count Loop
    j_Item := Pljson();
    j_Item := Pljson(j_List.Get(I));
  
    n_开单科室id := j_Item.Get_Number('plcdept_id');
    n_领药部门id := j_Item.Get_Number('takedept_id');
  
    v_Jpati := v_Jpati || ',{"pati_id":' || n_病人id;
    v_Jpati := v_Jpati || ',"pati_pageid":' || n_主页id;
    v_Jpati := v_Jpati || ',"pati_deptid":' || Nvl(n_病人科室id, 0);
    v_Jpati := v_Jpati || ',"pati_wardarea_id":' || Nvl(n_病人病区id, 0);
    v_Jpati := v_Jpati || ',"plcdept_id":' || Nvl(n_开单科室id, 0);
    v_Jpati := v_Jpati || ',"takedept_id":' || Nvl(n_领药部门id, 0);
    v_Jpati := v_Jpati || ',"pati_name":"' || Zljsonstr(v_姓名) || '"';
    v_Jpati := v_Jpati || ',"fee_audit_status":' || Nvl(n_费用审核标志, 0);
    v_Jpati := v_Jpati || ',"si_inp_status":' || Nvl(n_住院状态, 0);
    v_Jpati := v_Jpati || '}';
  End Loop;
  v_Jtmp_In := '{"input":{"item_list":[' || Substr(v_Jpati, 2) || ']}}';

  Zl_住院记帐记录_Insert_Check(v_Jtmp_In, Json_Out);
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Newbill_Check;
/

Create Or Replace Procedure Zl_Exsesvr_Checkexcitemvalid
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:获取病人黑名单信息 
  --入参：Json_In:格式
  --    input
  --      module  N  1  模块号
  --      pati_id  N  1  病人id
  --      balance_mode  N  1  结算模式
  --      fitem_type  C  1  收费类别:多个用逗 号
  --      fitem_ids  C  1  收费细目ids:多个用逗号
  --      fee_nos  C    1 费用单据号
  --
  --出参: Json_Out,格式如下
  --  output
  --    code              N   1   应答码：0-失败；1-成功
  --    message           C   1   应答消息：失败时返回具体的错误信息
  --    black_infor C 1 "黑名单信息:该病人不是黑名单病人，返回NULl，否则返回格式:控制方式|提示的信息 ;控制方式：1-禁止;2-提示(或询问)"
  ---------------------------------------------------------------------------

  j_Input PLJson;
  j_Json  PLJson;

  n_病人id   门诊费用记录.病人id%Type;
  n_结算模式 Number(10);
  v_收费类别 Varchar2(32680);
  v_Nos      Varchar2(32680);
  n_模块号   Number(18);

  v_收费细目ids Varchar2(32680);
  v_Infor       Varchar2(32680);
Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_模块号   := j_Json.Get_Number('module');
  n_病人id   := j_Json.Get_Number('pati_id');
  n_结算模式 := j_Json.Get_Number('balance_mode');

  v_收费细目ids := j_Json.Get_String('fitem_type');
  v_收费类别    := j_Json.Get_String('fitem_ids');

  v_Nos := j_Json.Get_String('fee_nos');

  v_Infor := Zl_Get_Excuteitem_Infor_s(n_模块号, n_病人id, n_结算模式, v_收费类别, v_Nos, v_收费细目ids);

  Json_Out := '{"output":{"code":1,"message":"成功","ctrl_infor":"' || zlJsonStr(v_Infor) || '"}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkexcitemvalid;
/



CREATE OR REPLACE Procedure Zl_Exsesvr_Getnewblacklists
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:根据病人ID,获取需要保存的黑名单数据
  --入参：Json_In:格式
  --    input
  --      pati_id              N 1 病人id
  --      last_rgsappt_time    C 1 进入不良记录的最后一次时间：yyyy-mm-dd hh24:mi:ss
  --      operator_name        C 1 操作员姓名
  --      blackLst_regnos      C 1 进入黑名单的挂号单号,多个用逗号分离
  --
  --出参: Json_Out,格式如下
  --  output
  --    code                   N   1   应答码：0-失败；1-成功
  --    message                C   1   应答消息：失败时返回具体的错误信息
  --    badrec_list            C   保存的不良记录列表
  --      pati_id              N 1 病人id
  --      behavior_category    C 1 行为类别:如预约挂号
  --      happen_time          C 1 发生时间:yyyy-mm-dd hh24:mi:ss
  --      add_time             C 1 加入时间:yyyy-mm-dd hh24:mi:ss
  --      add_note             C 1 加入原因：如预约超期
  --      add_memo             C 1 加入说明
  --      additional_info      C 1 附加信息
  --      creator              C 1 登记人
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_Output Varchar2(32767);
  c_Output Clob;

  n_病人id     病人挂号记录.病人id%Type;
  v_操作员姓名 Varchar2(20);
  v_黑名单单号 Varchar2(32680);
  d_计算日期   Date;

  n_Count        Number(18);
  n_预约接诊效期 Number(18);
  n_预约退号效期 Number(18);
  n_预约接收效期 Number(18);
  d_最后预约时间 Date;
  v_Para         Varchar2(4000);

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_病人id       := j_Json.Get_Number('pati_id');
  d_最后预约时间 := To_Date(j_Json.Get_String('last_rgsappt_time'), 'yyyy-mm-dd hh24:mi:ss');

  v_操作员姓名 := j_Json.Get_String('operator_name');
  v_黑名单单号 := j_Json.Get_String('blackLst_regnos');
  v_黑名单单号 := ',' || Nvl(v_黑名单单号, '') || ',';

  If v_操作员姓名 Is Null Then
    v_操作员姓名 := zl_UserName;
  End If;

  --格式:预约未付款控制|预约接诊控制|预约退号控制
  --预约未付款控制：>0预约之后超过有效时间未接收预约的;<0表示预约之后，在超过延迟的有效时间未接收预约的
  --预约接诊控制：>0,预约之后超过有效时间就诊或未就诊视为爽约
  --预约退号控制:>0,预约之后超过有效时间未就诊且进行退号视为爽约

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
    Json_Out := '{"output":{"code":1,"message":"成功","badrec_list":[' || '' || ']}}';

    Return;
  End If;

  If Nvl(n_病人id, 0) = 0 Then
    d_计算日期 := Trunc(Sysdate) - 1 / 24 / 60 / 60; -- 缺省计算头一天的数据
    For c_预约 In (Select Distinct a.No, a.病人id, a.记录性质, a.记录状态, Nvl(a.预约时间, a.发生时间) As 预约时间, c.名称 As 部门名称, 执行人, 接收时间
                 From 病人挂号记录 A, 部门表 C
                 Where a.执行部门id = c.Id(+) And a.预约 = 1 And
                       ((a.记录性质 = 2 And Nvl(a.记录状态, 0) = 1 And
                       ((a.预约时间 + n_预约接收效期 * (1 / 24 / 60)) <= Sysdate And n_预约接收效期 <> 0)) Or
                       (a.记录性质 = 1 And Nvl(a.记录状态, 0) = 1 And
                       ((Nvl(a.执行时间, Sysdate) - Nvl(a.预约时间, Sysdate)) * 24 * 60 >= n_预约接诊效期) And n_预约接诊效期 <> 0) Or
                       (a.记录状态 = 2 And ((a.登记时间 - Nvl(a.发生时间, Sysdate)) * 24 * 60 >= n_预约退号效期) And n_预约退号效期 <> 0)) And
                       a.预约时间 >= Trunc(d_计算日期) And a.预约时间 <= d_计算日期 And Instr(v_黑名单单号, ',' || a.No || ',') = 0) Loop

      --预约超期
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

      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;

        v_Output := Null;
      End If;

      zlJsonPutValue(v_Output, 'pati_id', c_预约.病人id, 1, 1);
      zlJsonPutValue(v_Output, 'behavior_category', '预约挂号');
      zlJsonPutValue(v_Output, 'happen_time', To_Char(c_预约.预约时间, 'yyyy-mm-dd hh24:mi:ss'));
      zlJsonPutValue(v_Output, 'add_time', To_Char(Sysdate, 'yyyy-mm-dd hh24:mi:ss'));
      zlJsonPutValue(v_Output, 'add_note', '预约超期');
      zlJsonPutValue(v_Output, 'add_memo', v_Para);
      zlJsonPutValue(v_Output, 'additional_info', c_预约.No);
      zlJsonPutValue(v_Output, 'creator', v_操作员姓名, 0, 2);

    End Loop;
    If Not c_Output Is Null And Not v_Output Is Null Then
      c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
      v_Output := '';
    End If;
  Else

    For c_预约 In (Select Distinct a.No, a.病人id, a.记录性质, a.记录状态, Nvl(a.预约时间, a.发生时间) As 预约时间, c.名称 As 部门名称, 执行人, 接收时间
                 From 病人挂号记录 A, 部门表 C
                 Where a.病人id = n_病人id And a.执行部门id = c.Id(+) And a.预约 = 1 And
                       ((a.记录性质 = 2 And Nvl(a.记录状态, 0) = 1 And
                       ((a.预约时间 + n_预约接收效期 * (1 / 24 / 60)) <= Sysdate And n_预约接收效期 <> 0)) Or
                       (a.记录性质 = 1 And Nvl(a.记录状态, 0) = 1 And
                       ((Nvl(a.执行时间, Sysdate) - Nvl(a.预约时间, Sysdate)) * 24 * 60 >= n_预约接诊效期) And n_预约接诊效期 <> 0) Or
                       (a.记录状态 = 2 And ((a.登记时间 - Nvl(a.发生时间, Sysdate)) * 24 * 60 >= n_预约退号效期) And n_预约退号效期 <> 0)) And
                       a.发生时间 + 0 >= Nvl(d_最后预约时间, To_Date('1990-01-01', 'YYYY-MM-DD')) And Instr(v_黑名单单号, ',' || a.No || ',') = 0

                 ) Loop

      --预约超期
      v_Para := '在' || To_Char(c_预约.预约时间, 'yyyy-mm-dd hh24:mi:ss');
      v_Para := v_Para || '预约的"' || c_预约.部门名称 || '"科室';

      If c_预约.执行人 Is Not Null Then
        v_Para := v_Para || '、医生为"' || c_预约.执行人 || '"';
      End If;
      v_Para := v_Para || '(预约单:' || c_预约.No || Case
                  When c_预约.记录状态 = 2 Then
                   ' 发生超期退号'
                  When c_预约.记录性质 = 1 Then
                   ' 发生超期接诊'
                  Else
                   ''
                End || ')的号源，未按时就诊。';

      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;

        v_Output := Null;
      End If;
      zlJsonPutValue(v_Output, 'pati_id', Nvl(c_预约.病人id, 0), 1, 1);
      zlJsonPutValue(v_Output, 'behavior_category', '预约挂号');
      zlJsonPutValue(v_Output, 'happen_time', To_Char(c_预约.预约时间, 'yyyy-mm-dd hh24:mi:ss'));
      zlJsonPutValue(v_Output, 'add_time', To_Char(Sysdate, 'yyyy-mm-dd hh24:mi:ss'));
      zlJsonPutValue(v_Output, 'add_note', '预约超期');
      zlJsonPutValue(v_Output, 'add_memo', v_Para);
      zlJsonPutValue(v_Output, 'additional_info', c_预约.No);
      zlJsonPutValue(v_Output, 'creator', v_操作员姓名, 0, 2);

    End Loop;
    If Not c_Output Is Null And Not v_Output Is Null Then
      c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
      v_Output := '';
    End If;
  End If;
  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"成功","badrec_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"成功","badrec_list":[' || v_Output || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getnewblacklists;
/

Create Or Replace Procedure Zl_Exsesvr_Updateregpatiinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:修改病人挂号数据的病人信息
  --入参：Json_In:格式
  --    input
  --        reg_no C 1 单据号
  --        pati_name C 1 姓名
  --        pati_sex  C 1 性别
  --        pati_age  C 1 年龄
  --        outpatient_num  C   门诊号
  --出参: Json_Out,格式如下
  --  output
  --    code              N   1   应答码：0-失败；1-成功
  --    message           C   1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_No     病人挂号记录.No%Type;
  v_姓名   病人挂号记录.姓名%Type;
  v_性别   病人挂号记录.性别%Type;
  v_年龄   病人挂号记录.年龄%Type;
  n_门诊号 病人挂号记录.门诊号%Type;
Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_No := j_Json.Get_String('reg_no');

  v_姓名   := j_Json.Get_String('pati_name');
  v_性别   := j_Json.Get_String('pati_sex');
  v_年龄   := j_Json.Get_String('pati_age');
  n_门诊号 := To_Number(j_Json.Get_String('outpatient_num'));

  Zl_病人挂号基本信息_Update(v_No, n_门诊号, v_姓名, v_性别, v_年龄);

  Json_Out := zlJsonOut('成功', 1);
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Updateregpatiinfo;
/


CREATE OR REPLACE Procedure Zl_Exsesvr_Getpatisurplusinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：获取病人预交余额和费用余额
  --入参：Json_In:格式
  --  input
  --   pati_id            N   病人id
  --   pati_pageid        N   主页id
  --   pati_ids           C   病人IDs,多个用逗号分离
  --   use_type           N   0/null 返回病人预交余额和费用余额  =1时返回病人未结费用列表

  --   说明：根据病人id和主页id查询具体某一次的住院费用余额，返回infee_surplus
  --         根据病人ids串查询对应病人的所有余额信息，返回surplus_list[]
  --出参: Json_Out,格式如下
  --  output
  --    code                N 1 应答码：0-失败；1-成功
  --    message             C 1 应答消息：失败时返回具体的错误信息
  --    infee_surplus       N 1 总未结费用
  --    indpst_surplus      N 1 住院预交余额
  --    infee_surplusnew    N 1 本次未结费用
  --    surplus_list[]      C 1 余额列表
  --      pati_Id           N   病人ID
  --      outdpst_surplus   N 1 门诊预交余额
  --      indpst_surplus    N 1 住院预交余额
  --      outfee_surplus    N 1 门诊费用余额
  --      infee_surplus     N 1 住院费用余额
  --    unfinish_list[]     C 1 病人未结费用列表
  --      pati_Id           N   病人ID
  --      page_Id           N   主页ID
  --      infee_surplusnew  N 1 本次未结费用
  ---------------------------------------------------------------------------
  j_Input Pljson;
  j_Json  Pljson;

  v_Output Varchar2(32767);
  c_Output Clob;

  n_病人id 病人未结费用.病人id%Type;
  n_主页id 病人未结费用.主页id%Type;
  v_Ids    Varchar2(3000);
  n_Type   Number;

  n_本次费用余额 Number(16, 5);
  n_住院余额     Number(16, 5);
  n_总住院费用   Number(16, 5);

Begin

  --解析入参
  j_Input := Pljson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_病人id := j_Json.Get_Number('pati_id');
  n_主页id := j_Json.Get_Number('pati_pageid');
  n_Type   := Nvl(j_Json.Get_Number('use_type'), 0);

  v_Ids := j_Json.Get_String('pati_ids');

  If Nvl(v_Ids, '-') = '-' And Nvl(n_病人id, 0) = 0 Then
    Json_Out := Zljsonout('未传入任何查询条件，请检查!');
    Return;
  End If;

  If n_Type = 0 Then

    If Nvl(n_病人id, 0) <> 0 Then
      If Nvl(n_主页id, 0) = 0 Then
        Json_Out := Zljsonout('未传入主页id，请检查!');
        Return;
      End If;
      Select Nvl(Sum(金额), 0) Into n_本次费用余额 From 病人未结费用 Where 病人id = n_病人id And 主页id = n_主页id;

      Select Sum(Decode(a.类型, 2, a.预交余额, 0)) As 住院余额, Sum(Decode(a.类型, 2, a.费用余额, 0)) As 住院费用
      Into n_住院余额, n_总住院费用
      From 病人余额 A
      Where a.病人id = n_病人id And a.性质 = 1;

      Zljsonputvalue(v_Output, 'code', 1, 1, 1);
      Zljsonputvalue(v_Output, 'message', '成功');
      Zljsonputvalue(v_Output, 'infee_surplus', n_总住院费用, 1);
      Zljsonputvalue(v_Output, 'indpst_surplus', n_住院余额, 1);
      Zljsonputvalue(v_Output, 'infee_surplusnew', n_本次费用余额, 1, 2);
      Json_Out := '{"output":' || v_Output || '}';

      Return;
    Else
      For r_病人余额 In (Select a.病人id, Sum(Decode(a.类型, 1, a.预交余额, 0)) As 门诊余额, Sum(Decode(a.类型, 2, a.预交余额, 0)) As 住院余额,
                            Sum(Decode(a.类型, 1, a.费用余额, 0)) As 门诊费用, Sum(Decode(a.类型, 2, a.费用余额, 0)) As 住院费用
                     From 病人余额 A, Table(f_Num2list(v_Ids)) B
                     Where a.病人id = b.Column_Value And a.性质 = 1
                     Group By 病人id) Loop
        If Length(Nvl(v_Output, ' ')) > 30000 Then
          If c_Output Is Not Null Then
            c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
          Else
            c_Output := To_Clob(v_Output);
          End If;

          v_Output := Null;
        End If;

        Zljsonputvalue(v_Output, 'pati_id', r_病人余额.病人id, 1, 1);
        Zljsonputvalue(v_Output, 'outdpst_surplus', r_病人余额.门诊余额, 1);
        Zljsonputvalue(v_Output, 'indpst_surplus', r_病人余额.住院余额, 1);
        Zljsonputvalue(v_Output, 'outfee_surplus', r_病人余额.门诊费用, 1);
        Zljsonputvalue(v_Output, 'infee_surplus', r_病人余额.住院费用, 1, 2);
      End Loop;

      If Not c_Output Is Null And Not v_Output Is Null Then
        c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        v_Output := '';
      End If;

      If Not c_Output Is Null Then
        Json_Out := To_Clob('{"output":{"code":1,"message":"成功","surplus_list":[') || c_Output || To_Clob(']}}');
      Else
        Json_Out := '{"output":{"code":1,"message":"成功","surplus_list":[' || v_Output || ']}}';
      End If;

      Return;
    End If;
  Elsif n_Type = 1 Then
    For r_未结费用 In (Select a.病人id, a.主页id, Sum(Nvl(a.金额, 0)) As 未结费用
                   From 病人未结费用 A, Table(f_Num2list(v_Ids)) B
                   Where a.病人id = b.Column_Value And a.主页id Is Not Null
                   Group By a.病人id, a.主页id
                   Order By a.病人id, a.主页id) Loop
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
        v_Output := Null;
      End If;

      Zljsonputvalue(v_Output, 'pati_id', r_未结费用.病人id, 1, 1);
      Zljsonputvalue(v_Output, 'page_id', r_未结费用.主页id, 1);
      Zljsonputvalue(v_Output, 'infee_surplusnew', r_未结费用.未结费用, 1, 2);
    End Loop;

    If Not c_Output Is Null And Not v_Output Is Null Then
      c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
      v_Output := '';
    End If;

    If Not c_Output Is Null Then
      Json_Out := To_Clob('{"output":{"code":1,"message":"成功","unfinish_list":[') || c_Output || To_Clob(']}}');
    Else
      Json_Out := '{"output":{"code":1,"message":"成功","unfinish_list":[' || v_Output || ']}}';
    End If;

    Return;
  End If;

Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getpatisurplusinfo;
/


Create Or Replace Procedure Zl_Exsesvr_Getconsumercardtype
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:获取消费卡类别
  --入参：Json_In:格式
  --    input
  --      enabled                N    是否启用:1-已启用;0-所有
  --出参: Json_Out,格式如下
  --  output
  --    code                      N   1   应答码：0-失败；1-成功
  --    message                   C   1   应答消息：失败时返回具体的错误信息
  --    type_list[]               C  1  支持的卡类别列表
  --          cardtype_id       N  1  id
  --          cardtype_num      N  1  编号
  --          cardtype_name     C  1  名称
  --          cardtype_stname   C  1  短名
  --          prefix_text         C  1  前缀文本
  --          cardno_len          N  1  卡号长度
  --          default             N  1  缺省标志
  --          fixed               N  1  是否固定:1-是系统固定;0-不是系统固定
  --          strict              N  1  是否严格控制:1-是严格控制;0-不是严格控制
  --          self_make           N  1  是否自制:1-是的;0-不是
  --          allow_return_cash   N  1  是否退现:1-允许;0-不允许
  --          must_all_return     N   1   是否全退:1-必需全退;0-允许部分退
  --          specpati            N   1   特定病人
  --          component           C   1   部件
  --          memo                C   1   备注
  --          blnc_mode           C   1   结算方式
  --          blnc_nature         N   1   结算性质
  --          pwdtxt           N   1   是否密文
  --          enabled             N   1   是否启用:1-已启用;0-未启用
  --          pwd_len             N   1   密码长度
  --          pwd_len_limit       N   1   密码长度限制:0-不作限制;1-固定密码长度;-n表示密码必须输入好多个位密码以上,但不能超过密码长度
  --          pwd_rule            N   1   密码规则:０-数字和字符组成;1-仅为数字组成
  --          readcard_nature     C   1   读卡性质,医疗卡读卡方式：第一位为:是否刷卡;第二位为是否扫描;第三位是否接触式读卡;第四位是否非接触式读卡。例如刷卡：'1000'
  --          keyboard_mode       N   1   键盘控制方式:：0-禁止使用软键盘;1-使用数字软键盘 ,2-使用字符软键盘
  --          def_delcash         N   1   是否缺省退现:允许退现时,默认是否退现
  ---------------------------------------------------------------------------
  v_Output Varchar2(32767);
  c_Output Clob;

  j_Input PLJson;
  j_Json  PLJson;

  n_是否启用 Number(2);

Begin
  --解析入参

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_是否启用 := j_Json.Get_Number('enabled');
  For c_消费卡类别 In (Select a.编号 As ID, a.编号, a.名称, Nvl(a.系统, 0) As 是否固定, a.结算方式, a.部件, Nvl(a.启用, 0) As 是否启用,
                         Nvl(a.自制卡, 0) As 是否自制, a.前缀文本, a.卡号长度, a.是否密文, a.是否退现, a.是否全退, a.密码长度, a.密码长度限制, a.密码规则, a.读卡性质,
                         a.键盘控制方式, a.限制类别, a.是否严格控制, a.是否特定病人, a.是否允许换卡, a.是否允许补卡, a.是否允许余额退款, a.应用场合, 0 As 缺省标志,
                         Nvl(b.性质, 0) As 结算性质, 0 As 是否缺省退现, '' As 备注
                  From 消费卡类别目录 A, 结算方式 B
                  Where a.结算方式 = b.名称(+) And Decode(Nvl(n_是否启用, 0), 0, 0, Nvl(a.启用, 0)) = Nvl(n_是否启用, 0)
                  
                  ) Loop
  
    If Length(Nvl(v_Output, ' ')) > 30000 Then
      If c_Output Is Not Null Then
        c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
      Else
        c_Output := To_Clob(v_Output);
      End If;
    
      v_Output := Null;
    End If;
  
    zlJsonPutValue(v_Output, 'cardtype_id', c_消费卡类别.Id, 1, 1);
    zlJsonPutValue(v_Output, 'cardtype_num', c_消费卡类别.编号);
    zlJsonPutValue(v_Output, 'cardtype_name', c_消费卡类别.名称);
    zlJsonPutValue(v_Output, 'cardtype_stname', Substr(c_消费卡类别.名称, 1, 1));
  
    zlJsonPutValue(v_Output, 'prefix_text', Nvl(c_消费卡类别.前缀文本, ''));
    zlJsonPutValue(v_Output, 'cardno_len', Nvl(c_消费卡类别.卡号长度, 0), 1);
    zlJsonPutValue(v_Output, 'default', Nvl(c_消费卡类别.缺省标志, 0), 1);
  
    zlJsonPutValue(v_Output, 'fixed', Nvl(c_消费卡类别.是否固定, 0), 1);
    zlJsonPutValue(v_Output, 'strict', Nvl(c_消费卡类别.是否严格控制, 0), 1);
    zlJsonPutValue(v_Output, 'self_make', Nvl(c_消费卡类别.是否自制, 0), 1);
    zlJsonPutValue(v_Output, 'allow_return_cash', Nvl(c_消费卡类别.是否退现, 0), 1);
    zlJsonPutValue(v_Output, 'must_all_return', Nvl(c_消费卡类别.是否全退, 0), 1);
    zlJsonPutValue(v_Output, 'specpati', Nvl(c_消费卡类别.是否特定病人, 0), 1);
  
    zlJsonPutValue(v_Output, 'component', Nvl(c_消费卡类别.部件, ''));
    zlJsonPutValue(v_Output, 'memo', Nvl(c_消费卡类别.备注, ''));
  
    zlJsonPutValue(v_Output, 'blnc_mode', Nvl(c_消费卡类别.结算方式, ''));
    zlJsonPutValue(v_Output, 'blnc_nature', Nvl(c_消费卡类别.结算性质, 0), 1);
  
    zlJsonPutValue(v_Output, 'pwdtxt', Nvl(c_消费卡类别.是否密文, 0), 1);
    zlJsonPutValue(v_Output, 'enabled', Nvl(c_消费卡类别.是否启用, 0), 1);
  
    zlJsonPutValue(v_Output, 'pwd_len', Nvl(c_消费卡类别.密码长度, 0), 1);
    zlJsonPutValue(v_Output, 'pwd_len_limit', Nvl(c_消费卡类别.密码长度限制, 0), 1);
    zlJsonPutValue(v_Output, 'pwd_rule', Nvl(c_消费卡类别.密码规则, 0), 1);
  
    zlJsonPutValue(v_Output, 'readcard_nature', Nvl(c_消费卡类别.读卡性质, '1000'));
    zlJsonPutValue(v_Output, 'keyboard_mode', Nvl(c_消费卡类别.键盘控制方式, 0), 1);
    zlJsonPutValue(v_Output, 'def_return_cash', Nvl(c_消费卡类别.是否缺省退现, 0), 1, 2);
  
  End Loop;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"成功","type_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"成功","type_list":[' || v_Output || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getconsumercardtype;
/


Create Or Replace Procedure Zl_Exsesvr_Getpatisurety
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：获取病人担保额信息
  --入参：Json_In:格式
  --  input
  --     pati_id            N 1 病人Id
  --     pati_pageid        N 0 主页ID
  --     pati_ids           C 0 病案主页关键信息拼串，病人ID:主页ID,....
  --     surety_prop        N 0 是否获取担保性质，0-不获取，1-要获取，目前仅支持单个病人，
  --     query_type         N 0 1-获取存在病人担保记录的病人ID及主页ID
  --出参: Json_Out,格式如下
  --  output
  --    code                N   1   应答吗：0-失败；1-成功
  --    message             C   1   应答消息：失败时返回具体的错误信息
  --    guarantee_money     N       担保金额
  --    entsurety           C       担保人
  --    surety_prop         N       担保性质，0-长期担保，1-临时提保，以最近的一次担保记录的性质为准
  --    item_list[]
  --       pati_id            N 1 病人id
  --       pati_pageid        N 1 主页id
  --       guarantee_money    N 1 担保金额
  --       entsurety          C 1 担保人
  --       surety_prop        N 1 担保性质
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_主页id   病人担保记录.主页id%Type;
  n_病人id   病人担保记录.病人id%Type;
  v_担保人   病人担保记录.担保人%Type;
  n_担保额   病人担保记录.担保额%Type;
  n_担保性质 Number;
  n_查询类型 Number(3);
  l_病人     t_StrList := t_StrList();
  v_病人ids  Varchar2(32767);

  v_Output Varchar2(32767);
  c_Output Clob;
Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_病人id   := j_Json.Get_Number('pati_id');
  n_主页id   := j_Json.Get_Number('pati_pageid');
  n_查询类型 := j_Json.Get_Number('query_type');

  If j_Json.Exist('pati_ids') Then
    --没有该节点时,执行下面语句有误
    v_病人ids := j_Json.Get_String('pati_ids');
  End If;

  If Nvl(n_病人id, 0) = 0 And v_病人ids Is Null Then
    Json_Out := zlJsonOut('未传入病人id,请检查！', 0);
    Return;
  End If;

  --提取当前有效担保额及有效担保记录数
  n_担保额 := 0;
  v_担保人 := Null;
  If n_病人id <> 0 Then
    For r_提保信息 In (Select 担保人, 担保额, 担保性质
                   From 病人担保记录
                   Where 病人id = n_病人id And (主页id = n_主页id Or Nvl(n_主页id, 0) = 0) And (到期时间 Is Null Or 到期时间 > Sysdate) And
                         删除标志 = 1
                   Order By 登记时间 Desc) Loop
      If n_担保性质 Is Null Then
        n_担保性质 := Nvl(r_提保信息.担保性质, 0);
      End If;
      n_担保额 := n_担保额 + r_提保信息.担保额;
      v_担保人 := v_担保人 || ',' || r_提保信息.担保人;
    End Loop;
    v_担保人 := Substr(v_担保人, 2, 100);
    Json_Out := '{"output":{"code":1,"message":"成功","guarantee_money":' || zlJsonStr(Nvl(n_担保额, 0), 1) ||
                ',"entsurety":"' || zlJsonStr(v_担保人) || '","surety_prop":' || Nvl(n_担保性质, 0) || '}}';
    Return;
  
  End If;

  While v_病人ids Is Not Null Loop
    If Length(v_病人ids) <= 4000 Then
      l_病人.Extend;
      l_病人(l_病人.Count) := v_病人ids;
      v_病人ids := Null;
    Else
      l_病人.Extend;
      l_病人(l_病人.Count) := Substr(v_病人ids, 1, Instr(v_病人ids, ',', 3940) - 1);
      v_病人ids := Substr(v_病人ids, Instr(v_病人ids, ',', 3940) + 1);
    End If;
  End Loop;

  v_Output := Null;
  For I In 1 .. l_病人.Count Loop
    v_病人ids := l_病人(I);
    If Nvl(n_查询类型, 0) = 0 Then
    
      For R In (Select a.病人id, a.主页id, 担保人, a.担保性质, Sum(a.担保额) As 担保额
                From 病人担保记录 A,
                     (Select /*+cardinality(f,10)*/
                        To_Number(f.C1) As 病人id, To_Number(f.C2) As 主页id
                       From Table(f_Str2List2(v_病人ids)) F) N
                Where a.病人id = n.病人id And Nvl(a.主页id, 0) = n.主页id And (a.到期时间 Is Null Or a.到期时间 > Sysdate) And
                      a.删除标志 = 1
                Group By a.病人id, a.主页id, 担保人, a.担保性质) Loop
      
        If Length(Nvl(v_Output, ' ')) > 32000 Then
          If c_Output Is Not Null Then
            c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
          Else
            c_Output := To_Clob(v_Output);
          End If;
        
          v_Output := Null;
        End If;
        zlJsonPutValue(v_Output, 'pati_id', r.病人id, 1, 1);
        zlJsonPutValue(v_Output, 'pati_pageid', r.主页id, 1, 0);
        zlJsonPutValue(v_Output, 'guarantee_money', r.担保额, 1, 0);
        zlJsonPutValue(v_Output, 'entsurety', r.担保人, 0, 0);
        zlJsonPutValue(v_Output, 'surety_prop', Nvl(r.担保性质, 0), 1, 2);
      End Loop;
    
    Else
    
      For R In (Select a.病人id, a.主页id
                From 病人担保记录 A,
                     (Select /*+cardinality(f,10)*/
                        To_Number(f.C1) As 病人id, To_Number(f.C2) As 主页id
                       From Table(f_Str2List2(v_病人ids)) F) N
                Where a.病人id = n.病人id And Nvl(a.主页id, 0) = n.主页id
                Group By a.病人id, a.主页id) Loop
      
        If Length(Nvl(v_Output, ' ')) > 32700 Then
          If c_Output Is Not Null Then
            c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
          Else
            c_Output := To_Clob(v_Output);
          End If;
          v_Output := Null;
        End If;
      
        zlJsonPutValue(v_Output, 'pati_id', r.病人id, 1, 1);
        zlJsonPutValue(v_Output, 'pati_pageid', r.主页id, 1, 2);
      End Loop;
    End If;
  End Loop;
  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"成功","item_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || v_Output || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getpatisurety;
/


Create Or Replace Procedure Zl_Exsesvr_Getpatisuretylist
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：获取病人担保额信息
  --入参：Json_In:格式
  --  input
  --     pati_id            N 1 病人Id
  --     pati_pageid        N 0 主页ID
  --     expidate           N 1 1-查询有效的担保信息;0-所有担保信息
  --出参: Json_Out,格式如下
  --  output
  --    code                N   1   应答吗：0-失败；1-成功
  --    message             C   1   应答消息：失败时返回具体的错误信息
  --    item_list[]         数组
  --      type               C 1 类别
  --      guarantor          C 1 担保人
  --      garnt_amount       C 1 担保额
  --      garnt_prop         N 1 担保性质
  --      garnt_reason       C 1 担保原因
  --      create_time        C 1 登记时间
  --      due_time           C 1 到期时间
  --      is_del             C 1 删除标志
  --      operator_name      C 1 操作员姓名
  --      operator_code      C 1 操作员编号
  --      del_operator_name  C 1 删除操作员姓名
  --      del_operator_code  C 1 删除操作员编号
  --      del_time           C 1 删除时间
  ---------------------------------------------------------------------------
  j_Input  PLJson;
  j_Json   PLJson;
  v_Output Varchar2(32767);
  c_Output Clob;

  n_主页id 病人担保记录.主页id%Type;
  n_病人id 病人担保记录.病人id%Type;
  n_有效期 Number(3);
Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_病人id := j_Json.Get_Number('pati_id');
  n_主页id := j_Json.Get_Number('pati_pageid');
  n_有效期 := j_Json.Get_Number('expidate');

  If Nvl(n_病人id, 0) = 0 Then
    Json_Out := zlJsonOut('未传入病人id,请检查！');
    Return;
  End If;
 
  If Nvl(n_有效期, 0) = 0 Then
    v_Output := Null;
    For R In (Select Decode(主页id, Null, '门诊', '第' || 主页id || '次住院') 类别, 担保人,
                     Decode(担保额, 999999999, '不限', To_Char(担保额, '999999990.00')) As 担保额, 担保性质, 担保原因,
                     To_Char(登记时间, 'yyyy-mm-dd hh24:mi:ss') 登记时间, To_Char(到期时间, 'yyyy-mm-dd hh24:mi:ss') 到期时间,
                     Decode(删除标志, 1, '', -1, '删除', '') As 删除标志, 操作员姓名, 操作员编号, 删除操作员姓名, 删除操作员编号, 删除时间
              From 病人担保记录
              Where 病人id = n_病人id And (主页id = n_主页id Or Nvl(n_主页id, 0) = 0)
              Order By 登记时间 Desc) Loop
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      
        v_Output := Null;
      End If;
    
      zlJsonPutValue(v_Output, 'type', r.类别, 0, 1);
      zlJsonPutValue(v_Output, 'guarantor', r.担保人);
      zlJsonPutValue(v_Output, 'garnt_amount', r.担保额);
      zlJsonPutValue(v_Output, 'garnt_prop', r.担保性质, 1);
      zlJsonPutValue(v_Output, 'garnt_reason', r.担保原因);
      zlJsonPutValue(v_Output, 'create_time', r.登记时间);
      zlJsonPutValue(v_Output, 'due_time', r.到期时间);
      zlJsonPutValue(v_Output, 'is_del', r.删除标志);
      zlJsonPutValue(v_Output, 'operator_name', r.操作员姓名);
      zlJsonPutValue(v_Output, 'operator_code', r.操作员编号);
      zlJsonPutValue(v_Output, 'del_operator_name', r.删除操作员姓名);
      zlJsonPutValue(v_Output, 'del_operator_code', r.删除操作员编号);
      zlJsonPutValue(v_Output, 'del_time', r.删除时间, 0, 2);
    End Loop;
  Else
    For R In (Select 担保人, Decode(担保额, 999999999, '不限', To_Char(担保额, '999999990.00')) As 担保额, Nvl(担保性质, 0) As 担保性质, 担保原因,
                     To_Char(到期时间, 'yyyy-mm-dd hh24:mi:ss') 到期时间, To_Char(登记时间, 'yyyy-mm-dd hh24:mi:ss') 登记时间
              From 病人担保记录
              Where 病人id = n_病人id And 主页id = n_主页id And (到期时间 Is Null Or 到期时间 > Sysdate) And 删除标志 = 1) Loop
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      
        v_Output := Null;
      End If;
      zlJsonPutValue(v_Output, 'guarantor', r.担保人, 0, 1);
      zlJsonPutValue(v_Output, 'garnt_amount', r.担保额);
      zlJsonPutValue(v_Output, 'garnt_prop', r.担保性质, 1);
      zlJsonPutValue(v_Output, 'garnt_reason', r.担保原因);
      zlJsonPutValue(v_Output, 'create_time', r.登记时间);
      zlJsonPutValue(v_Output, 'due_time', r.到期时间, 0, 2);
    End Loop;
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"成功","item_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || v_Output || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getpatisuretylist;
/


Create Or Replace Procedure Zl_Exsesvr_Patisuretyexpire
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：更新病人担保额信息
  --入参：Json_In:格式
  --  input
  --     pati_id            N 1 病人Id
  --     pati_pageid        N 1 主页ID
  --出参: Json_Out,格式如下
  --  output
  --    code                N   1   应答吗：0-失败；1-成功
  --    message             C   1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  n_主页id Number(5);
  n_病人id Number(18);
  j_Input  PLJson;
  j_Json   PLJson;

Begin

  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_病人id := j_Json.Get_Number('pati_id');
  n_主页id := j_Json.Get_Number('pati_pageid');

  If Nvl(n_病人id, 0) = 0 Then
    Json_Out := zlJsonOut('未传入病人id,请检查！');
    Return;
  End If;
  --正常担保记录出院后自动失效 
  Update 病人担保记录
  Set 到期时间 = Sysdate
  Where 病人id = n_病人id And 主页id = n_主页id And Nvl(到期时间, Sysdate + 1) > Sysdate And 删除标志 = 1 And Nvl(担保性质, 0) = 0;
  Json_Out := zlJsonOut('成功', 1);
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Patisuretyexpire;
/

Create Or Replace Procedure Zl_Exsesvr_Getwarnline
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：获取记帐报警线，过滤病人用于排开欠费病人
  --入参：Json_In:格式
  --  input
  --     pati_scheme  C 1 适用病人
  --     wardarea_id  N 1 病区id
  --     query_type   N 1 查询方式
  --                     0-仅根据 病区id / 适用病人 查找，返回一个值
  --                     1-按病区id 查找，返回列表
  --                     2-获取所有报警线设置
  --                     3-根据病区id，适用病人查找，返回报警方法，报警值，报警标志
  --出参: Json_Out,格式如下
  --  output
  --    code                N 1 应答吗：0-失败；1-成功
  --    message             C 1 应答消息：失败时返回具体的错误信息
  --      alarm_value       N 1 报警值,-1，表示未到找数据
  --      item_list[]
  --        pati_scheme     C 1 适用病人
  --        alarm_way       N 1 报警方法
  --        alarm_value     N 1 报警值
  --        alarm_one       C 1 报警标志1
  --        alarm_two       C 1 报警标志2
  --        alarm_three     C 1 报警标志3
  --        wardarea_id     N 1 病区id
  ---------------------------------------------------------------------------
  j_Input    PLJson;
  j_Json     PLJson;
  n_查询方式 Number;
  v_适用病人 Varchar2(200);
  n_病区id   Number(18);
  n_报警值   Number;
  v_Temp     Varchar2(32767);
Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_查询方式 := j_Json.Get_Number('query_type');
  v_适用病人 := j_Json.Get_String('pati_scheme');
  n_病区id   := j_Json.Get_Number('wardarea_id');

  If Nvl(n_查询方式, 0) = 0 Then
    n_报警值 := -1;
    For R In (Select 报警值
              From 记帐报警线
              Where 报警方法 = 1 And Nvl(病区id, 0) = Nvl(n_病区id, 0) And 报警值 Is Not Null And 适用病人 = v_适用病人) Loop
      n_报警值 := r.报警值;
      Exit;
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"成功","alarm_value":' || zlJsonStr(n_报警值, 1) || '}}';
  Elsif n_查询方式 = 3 Then
    For r_报警线 In (Select 适用病人, Nvl(报警方法, 1) As 报警方法, 报警值, 报警标志1, 报警标志2, 报警标志3
                  From 记帐报警线
                  Where Nvl(病区id, 0) = Nvl(n_病区id, 0) And 适用病人 = v_适用病人) Loop
      --只取一条数据
      v_Temp := v_Temp || '{"pati_scheme":"' || r_报警线.适用病人 || '"';
      v_Temp := v_Temp || ',"alarm_way":' || Nvl(r_报警线.报警方法, 0);
      v_Temp := v_Temp || ',"alarm_value":' || zlJsonStr(r_报警线.报警值, 1);
      v_Temp := v_Temp || ',"alarm_one":"' || r_报警线.报警标志1 || '"';
      v_Temp := v_Temp || ',"alarm_two":"' || r_报警线.报警标志2 || '"';
      v_Temp := v_Temp || ',"alarm_three":"' || r_报警线.报警标志3 || '"';
      v_Temp := v_Temp || '}';
      Exit;
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || v_Temp || ']}}';
  Elsif n_查询方式 = 1 Then
    For r_报警线 In (Select 适用病人, Nvl(报警方法, 1) As 报警方法, 报警值, 报警标志1, 报警标志2, 报警标志3
                  From 记帐报警线
                  Where 病区id = Nvl(n_病区id, 0)) Loop
      v_Temp := v_Temp || ',{"pati_scheme":"' || zlJsonStr(r_报警线.适用病人) || '"';
      v_Temp := v_Temp || ',"alarm_way":' || Nvl(r_报警线.报警方法, 0);
      v_Temp := v_Temp || ',"alarm_value":' || zlJsonStr(r_报警线.报警值, 1);
      v_Temp := v_Temp || ',"alarm_one":"' || r_报警线.报警标志1 || '"';
      v_Temp := v_Temp || ',"alarm_two":"' || r_报警线.报警标志2 || '"';
      v_Temp := v_Temp || ',"alarm_three":"' || r_报警线.报警标志3 || '"';
      v_Temp := v_Temp || '}';
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || Substr(v_Temp, 2) || ']}}';
  Elsif n_查询方式 = 2 Then
    For r_报警线 In (Select Nvl(病区id, 0) As 病区id, 适用病人, Nvl(报警方法, 1) As 报警方法, 报警值, 报警标志1, 报警标志2, 报警标志3
                  From 记帐报警线) Loop
      v_Temp := v_Temp || ',{"wardarea_id":' || Nvl(r_报警线.病区id || '', 'null');
      v_Temp := v_Temp || ',"pati_scheme":"' || zlJsonStr(r_报警线.适用病人) || '"';
      v_Temp := v_Temp || ',"alarm_way":' || Nvl(r_报警线.报警方法, 0);
      v_Temp := v_Temp || ',"alarm_value":' || zlJsonStr(r_报警线.报警值, 1);
      v_Temp := v_Temp || ',"alarm_one":"' || r_报警线.报警标志1 || '"';
      v_Temp := v_Temp || ',"alarm_two":"' || r_报警线.报警标志2 || '"';
      v_Temp := v_Temp || ',"alarm_three":"' || r_报警线.报警标志3 || '"';
      v_Temp := v_Temp || '}';
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || Substr(v_Temp, 2) || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getwarnline;
/

Create Or Replace Procedure Zl_Exsesvr_Newbill
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：门诊病人住院病人发送医嘱生成费用单据
  --入参：Json_In:格式
  --input
  --  pati_list[] 病人列表，单个病人时可以无该节点
  --    billtype                                            N 1 类型,1-收费单，2-记帐单
  --    pati_source                                         N 1 来源，1-门诊，2-住院
  --    pati_id                                             N 1 病人id
  --    pati_pageid                                         N 1 主页id
  --    baby_num                                            N 1 婴儿费
  --    sgin_no                                             N 1 标识号，门诊号，住院号
  --    bed_num                                             C 1 床号
  --    pati_name                                           C 1 姓名
  --    pati_sex                                            C 1 性别
  --    pati_age                                            C 1 年龄
  --    fee_category                                        C 1 费别
  --    overtime_sign                                       N 1 加班标志
  --    pati_deptid                                         N 1 病人科室id
  --    pati_wardarea_id                                    N 1 病人病区id
  --    operator_name                                       C 1 操作员姓名
  --    operator_code                                       C 1 操作员编号
  --    outpati_tag                                         N 1 门诊标志
  --    rgst_id                                             N 1 就诊id
  --    emg_sign                                            N 1 是否急诊
  --    item_list[]  明细列表
  --        fee_id                                        N 1 费用id
  --        fee_no                                        C 1 No
  --        serial_num                                    N 1 序号
  --        charge_tag                                    N 1 划价
  --        placer                                        C 1 开单人
  --        plcdept_id                                    N 1 开单部门id
  --        sub_serial_num                                N 1 从属父号
  --        fitem_id                                      N 1 收费细目id
  --        item_type                                     C 1 收费类别
  --        unit                                          C 1 计算单位
  --        pharmacy_window                               C 1 发药窗口
  --        packages_num                                  N 1 付数
  --        send_num                                      N 1 数次
  --        ext_mark                                      N 1 附加标志
  --        exe_deptid                                    N 1 执行部门id
  --        price_ftrnum                                  N 1 价格父号
  --        income_item_id                                N 1 收入项目id
  --        receipt_name                                  C 1 收据费目
  --        price                                         N 1 标准单价
  --        fee_amrcvb                                    N 1 应收金额
  --        fee_ampaib                                    N 1 实收金额
  --        happen_time                                   C 1 发生时间
  --        create_time                                   C 1 登记时间
  --        memo                                          C 1 费用摘要
  --        order_id                                      N 1 医嘱序号
  --        baby_num                                      N 1 婴儿费
  --        exe_properties                                N 1 执行性质
  --        decoction_method                              C 1 煎法
  --        morphology                                    C 1 中药形态
  --        bakstuff_batch                                N 1 批次
  --        insurance                                     N 1 保险项目否
  --        insure_id                                     N 1 保险大类id
  --        insure_code                                   C 1 保险编码
  --        fee_type                                      C 1 费用类型
  --        si_manp_money                                 N 1 统筹金额
  --        synchro                                       N 1 更新同步标志
  --        effective_time                                N 1 期效
  --        receipt_issecret                              N 1 保密
  --        takedept_id                                   N 1 领药部门id
  --        group_id                                      N 0 医疗小组id
  --        auto_finish                                   N 0 自动完成，针对卫材自动发料
  --出参: Json_Out,格式如下
  --output
  --  code                                                N 1 应答吗：0-失败；1-成功
  --  message                                             C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;
  j_List  Pljson_List;

  n_类型       住院费用记录.附加标志%Type; --费用来源,1-门诊费用记录,2-住院费用记录
  n_来源       住院费用记录.附加标志%Type; --单据类型,1-收费,2-记帐
  n_病人id     住院费用记录.病人id%Type;
  n_主页id     住院费用记录.主页id%Type;
  n_婴儿费     住院费用记录.婴儿费%Type;
  n_标识号     住院费用记录.标识号%Type;
  v_床号       住院费用记录.床号%Type;
  v_姓名       住院费用记录.姓名%Type;
  v_性别       住院费用记录.性别%Type;
  v_年龄       住院费用记录.年龄%Type;
  v_费别       住院费用记录.费别%Type;
  n_加班标志   住院费用记录.加班标志%Type;
  n_病人科室id 住院费用记录.病人科室id%Type;
  n_病人病区id 住院费用记录.病人病区id%Type;
  v_操作员姓名 住院费用记录.操作员姓名%Type;
  v_操作员编号 住院费用记录.操作员编号%Type;
  n_门诊标志   住院费用记录.门诊标志%Type;
  n_就诊id     住院费用记录.病人id%Type;
  n_是否急诊   住院费用记录.是否急诊%Type;

  n_Id         住院费用记录.Id%Type;
  v_No         住院费用记录.No%Type;
  n_序号       住院费用记录.序号%Type;
  n_划价       住院费用记录.附加标志%Type;
  v_开单人     住院费用记录.开单人%Type;
  n_开单部门id 住院费用记录.开单部门id%Type;
  n_从属父号   住院费用记录.从属父号%Type;
  n_收费细目id 住院费用记录.收费细目id%Type;
  v_收费类别   住院费用记录.收费类别%Type;
  v_计算单位   住院费用记录.计算单位%Type;
  v_发药窗口   住院费用记录.发药窗口%Type;
  n_付数       住院费用记录.付数%Type;
  n_数次       住院费用记录.数次%Type;
  n_附加标志   住院费用记录.附加标志%Type;
  n_执行部门id 住院费用记录.执行部门id%Type;
  n_价格父号   住院费用记录.价格父号%Type;
  n_收入项目id 住院费用记录.收入项目id%Type;
  v_收据费目   住院费用记录.收据费目%Type;
  n_标准单价   住院费用记录.标准单价%Type;
  n_应收金额   住院费用记录.应收金额%Type;
  n_实收金额   住院费用记录.实收金额%Type;
  d_发生时间   住院费用记录.发生时间%Type;
  d_登记时间   住院费用记录.登记时间%Type;
  v_摘要       住院费用记录.摘要%Type;
  n_医嘱序号   住院费用记录.医嘱序号%Type;
  n_执行性质   住院费用记录.附加标志%Type;
  v_煎法       住院费用记录.结论%Type;
  v_中药形态   住院费用记录.结论%Type;
  n_批次       住院费用记录.批次%Type;
  n_保险项目否 住院费用记录.保险项目否%Type;
  n_保险大类id 住院费用记录.保险大类id%Type;
  v_保险编码   住院费用记录.保险编码%Type;
  v_费用类型   住院费用记录.费用类型%Type;
  n_统筹金额   住院费用记录.统筹金额%Type;
  n_同步标志   病人费用异常记录.同步标志%Type;
  n_医嘱期效   住院费用记录.医嘱期效%Type;
  n_是否保密   住院费用记录.是否保密%Type;
  n_领药部门id 住院费用记录.领药部门id%Type;
  n_医疗小组id 住院费用记录.医疗小组id%Type;
  n_自动行完成 Number;
  n_I婴儿费    住院费用记录.婴儿费%Type;

  v_执行完成ids Varchar2(32767);

  n_Billcount Number;
  j_Patilist  Pljson_List;
Begin
  --解析入参 
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  --判断是否存在 pati_list 节点
  n_Billcount := 1;
  If j_Json.Exist('pati_list') Then
    j_Patilist  := j_Json.Get_Pljson_List('pati_list');
    n_Billcount := j_Patilist.Count;
  End If;

  For I In 1 .. n_Billcount Loop
    If j_Patilist Is Not Null Then
      j_Json := PLJson();
      j_Json := PLJson(j_Patilist.Get(I));
    End If;
  
    n_类型       := j_Json.Get_Number('billtype');
    n_来源       := j_Json.Get_Number('pati_source');
    n_病人id     := j_Json.Get_Number('pati_id');
    n_主页id     := j_Json.Get_Number('pati_pageid');
    n_婴儿费     := j_Json.Get_Number('baby_num');
    n_标识号     := j_Json.Get_Number('sgin_no');
    v_床号       := j_Json.Get_String('bed_num');
    v_姓名       := j_Json.Get_String('pati_name');
    v_性别       := j_Json.Get_String('pati_sex');
    v_年龄       := j_Json.Get_String('pati_age');
    v_费别       := j_Json.Get_String('fee_category');
    n_加班标志   := j_Json.Get_Number('overtime_sign');
    n_病人科室id := j_Json.Get_Number('pati_deptid');
    n_病人病区id := j_Json.Get_Number('pati_wardarea_id');
    v_操作员姓名 := j_Json.Get_String('operator_name');
    v_操作员编号 := j_Json.Get_String('operator_code');
    n_门诊标志   := j_Json.Get_Number('outpati_tag');
    n_就诊id     := j_Json.Get_Number('rgst_id');
    n_是否急诊   := j_Json.Get_Number('emg_sign');
  
    j_List        := j_Json.Get_Pljson_List('item_list');
    v_执行完成ids := Null;
    For I In 1 .. j_List.Count Loop
      j_Json       := PLJson();
      j_Json       := PLJson(j_List.Get(I));
      n_Id         := j_Json.Get_Number('fee_id');
      v_No         := j_Json.Get_String('fee_no');
      n_序号       := j_Json.Get_Number('serial_num');
      n_划价       := j_Json.Get_Number('charge_tag');
      v_开单人     := j_Json.Get_String('placer');
      n_开单部门id := j_Json.Get_Number('plcdept_id');
      n_从属父号   := j_Json.Get_Number('sub_serial_num');
      n_收费细目id := j_Json.Get_Number('fitem_id');
      v_收费类别   := j_Json.Get_String('item_type');
      v_计算单位   := j_Json.Get_String('unit');
      v_发药窗口   := j_Json.Get_String('pharmacy_window');
      n_付数       := j_Json.Get_Number('packages_num');
      n_数次       := j_Json.Get_Number('send_num');
      n_附加标志   := j_Json.Get_Number('ext_mark');
      n_执行部门id := j_Json.Get_Number('exe_deptid');
      n_价格父号   := j_Json.Get_Number('price_ftrnum');
      n_收入项目id := j_Json.Get_Number('income_item_id');
      v_收据费目   := j_Json.Get_String('receipt_name');
      n_标准单价   := j_Json.Get_Number('price');
      n_应收金额   := j_Json.Get_Number('fee_amrcvb');
      n_实收金额   := j_Json.Get_Number('fee_ampaib');
      d_发生时间   := To_Date(j_Json.Get_String('happen_time'), 'yyyy-mm-dd hh24:mi:ss');
      d_登记时间   := To_Date(j_Json.Get_String('create_time'), 'yyyy-mm-dd hh24:mi:ss');
      v_摘要       := j_Json.Get_String('memo');
      n_医嘱序号   := j_Json.Get_Number('order_id');
      n_I婴儿费    := j_Json.Get_Number('baby_num');
      n_执行性质   := j_Json.Get_Number('exe_properties');
      v_煎法       := j_Json.Get_String('decoction_method');
      v_中药形态   := j_Json.Get_String('morphology');
      n_批次       := j_Json.Get_Number('bakstuff_batch');
      n_保险项目否 := j_Json.Get_Number('insurance');
      n_保险大类id := j_Json.Get_Number('insure_id');
      v_保险编码   := j_Json.Get_String('insure_code');
      v_费用类型   := j_Json.Get_String('fee_type');
      n_统筹金额   := j_Json.Get_Number('si_manp_money');
      n_同步标志   := j_Json.Get_Number('synchro');
      n_医嘱期效   := j_Json.Get_Number('effective_time');
      n_是否保密   := j_Json.Get_Number('receipt_issecret');
      n_领药部门id := j_Json.Get_Number('takedept_id');
      n_医疗小组id := j_Json.Get_Number('group_id');
      n_自动行完成 := j_Json.Get_Number('auto_finish');
    
      If n_I婴儿费 Is Not Null Then
        n_婴儿费 := n_I婴儿费;
      End If;
    
      If n_来源 = 1 And n_类型 = 2 Then
        Zl_门诊记帐记录_Insert_s(v_No, n_序号, n_病人id, n_标识号, v_姓名, v_性别, v_年龄, v_费别, n_加班标志, n_婴儿费, n_病人科室id, n_开单部门id, v_开单人,
                           n_从属父号, n_收费细目id, v_收费类别, v_计算单位, n_付数, n_数次, n_附加标志, n_执行部门id, n_价格父号, n_收入项目id, v_收据费目,
                           n_标准单价, n_应收金额, n_实收金额, d_发生时间, d_登记时间, n_划价, v_发药窗口, v_操作员编号, v_操作员姓名, n_Id, Null, v_摘要,
                           n_医嘱序号, n_门诊标志, v_中药形态, v_煎法, n_主页id, n_病人病区id, n_批次, n_同步标志, n_就诊id, n_是否急诊, n_医嘱期效, n_是否保密);
      Elsif n_来源 = 1 And n_类型 = 1 Then
        Zl_门诊划价记录_Insert_s(v_No, n_序号, n_病人id, n_主页id, n_标识号, Null, v_姓名, v_性别, v_年龄, v_费别, n_加班标志, n_病人科室id, n_开单部门id,
                           v_开单人, n_从属父号, n_收费细目id, v_收费类别, v_计算单位, v_发药窗口, n_付数, n_数次, n_附加标志, n_执行部门id, n_价格父号,
                           n_收入项目id, v_收据费目, n_标准单价, n_应收金额, n_实收金额, d_发生时间, d_登记时间, v_操作员姓名, n_Id, v_摘要, n_医嘱序号, v_煎法,
                           1, v_保险编码, v_费用类型, n_保险项目否, n_保险大类id, v_中药形态, Null, n_病人病区id, n_批次, n_同步标志, n_就诊id, n_是否急诊,
                           n_医嘱期效, n_是否保密);
      Elsif n_来源 = 2 And n_类型 = 2 Then
        Zl_住院记帐记录_Insert_s(v_No, n_序号, n_病人id, n_主页id, n_标识号, v_姓名, v_性别, v_年龄, v_床号, v_费别, n_病人病区id, n_病人科室id, n_加班标志,
                           n_婴儿费, n_开单部门id, v_开单人, n_从属父号, n_收费细目id, v_收费类别, v_计算单位, n_保险项目否, n_保险大类id, v_保险编码, n_付数,
                           n_数次, n_附加标志, n_执行部门id, n_价格父号, n_收入项目id, v_收据费目, n_标准单价, n_应收金额, n_实收金额, n_统筹金额, d_发生时间,
                           d_登记时间, n_划价, v_操作员编号, v_操作员姓名, n_Id, Null, Null, v_摘要, n_是否急诊, n_医嘱序号, Null, v_费用类型, Null,
                           v_中药形态, n_医疗小组id, v_煎法, n_执行性质, n_批次, n_领药部门id, n_同步标志, n_医嘱期效, n_是否保密);
      End If;
    
      If n_自动行完成 = 1 Then
        v_执行完成ids := v_执行完成ids || ',' || n_Id;
      End If;
    
    End Loop;
  
    If v_执行完成ids Is Not Null Then
      v_执行完成ids := Substr(v_执行完成ids, 2);
      Select Sysdate Into d_登记时间 From Dual;
      If n_来源 = 1 Then
        Update 门诊费用记录
        Set 执行状态 = 1, 执行人 = v_操作员姓名, 执行时间 = d_登记时间
        Where ID In (Select /*+cardinality(j,10)*/
                      j.Column_Value
                     From Table(f_Num2List(v_执行完成ids)) J);
      Else
        Update 住院费用记录
        Set 执行状态 = 1, 执行人 = v_操作员姓名, 执行时间 = d_登记时间
        Where ID In (Select /*+cardinality(j,10)*/
                      j.Column_Value
                     From Table(f_Num2List(v_执行完成ids)) J);
      End If;
    End If;
  End Loop;

  Json_Out := zlJsonOut('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Newbill;
/

Create Or Replace Procedure Zl_Exsesvr_Actualmoney
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -----------------------------------------------------------------------------
  --功能:根据应收金额按费别进行打折计算,获以实收金额 
  --入参json
  --input     获以实收金额
  --          fee_category        C 1 费别
  --          fee_item_id         N 1 收费细目id
  --          income_item_id      N 1 收入项目id
  --          fee_amrcvb          N 1 应收金额
  --          quantity            N 1 数量
  --          price_cost          N 1 成本价
  --          order_id            N 1 医嘱id
  --          item_list[]列表
  --                  fee_category        C 1 费别
  --                   fee_item_id         N 1 收费细目id
  --                   income_item_id      N 1 收入项目id
  --                   fee_amrcvb          N 1 应收金额
  --                   quantity            N 1 数量
  --                   price_cost          N 1 成本价
  --                   order_id            N 1 医嘱id
  --出参json
  --output      
  --        code                  C 1 应答码：0-失败；1-成功
  --        message               C 1 应答消息：成功时返回成功信息，失败时返回具体的错误信息
  --        fee_category          C 1 费别
  --        fee_ampaib            N 1 实收金额，
  --        fee_ampaibs           C 1 实收金额串，逗号分割，当传入为是列表时返回
  ---------------------------------------------------------------------------
  j_Input    Pljson;
  j_Json     Pljson;
  j_List     Pljson_List := Pljson_List();
  n_实收金额 门诊费用记录.实收金额%Type;
  v_费别     门诊费用记录.费别%Type;
  v_Tmp      Varchar2(1000);
  v_Jtmp     Varchar2(32767);
Begin
  --解析入参
  j_Input := Pljson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');
  If j_Json.Exist('item_list') Then
    j_List := j_Json.Get_Pljson_List('item_list');
    If j_List Is Not Null Then
      For I In 1 .. j_List.Count Loop
        j_Json := Pljson();
        j_Json := Pljson(j_List.Get(I));
        Select Zl_Actualmoney_s(j_Json.Get_String('fee_category'), j_Json.Get_Number('fee_item_id'),
                                 j_Json.Get_Number('income_item_id'), j_Json.Get_Number('fee_amrcvb'),
                                 j_Json.Get_Number('quantity'), j_Json.Get_Number('price_cost'),
                                 j_Json.Get_Number('order_id'))
        Into v_Tmp
        From Dual;
      
        Select To_Number(Substr(v_Tmp, Instr(v_Tmp, ':') + 1)) As 金额 Into n_实收金额 From Dual;
      
        v_Jtmp := v_Jtmp || ',' || n_实收金额;
      End Loop;
    End If;
    Json_Out := '{"output":{"code":1,"message":"成功","fee_ampaibs":"' || Substr(v_Jtmp, 2) || '"}}';
  Else
    Select Zl_Actualmoney_s(j_Json.Get_String('fee_category'), j_Json.Get_Number('fee_item_id'),
                             j_Json.Get_Number('income_item_id'), j_Json.Get_Number('fee_amrcvb'),
                             j_Json.Get_Number('quantity'), j_Json.Get_Number('price_cost'), j_Json.Get_Number('order_id'))
    Into v_Tmp
    From Dual;
    Select Substr(v_Tmp, 1, Instr(v_Tmp, ':') - 1) As 费别, To_Number(Substr(v_Tmp, Instr(v_Tmp, ':') + 1)) As 金额
    Into v_费别, n_实收金额
    From Dual;
    Json_Out := '{"output":{"code":1,"message":"成功","fee_category":"' || v_费别 || '","fee_ampaib":' || Nvl(n_实收金额, 0) || '}}';
  End If;
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Actualmoney;
/

Create Or Replace Procedure Zl_Exsesvr_Checkonetimefee
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  --------------------------------------------------------------------------------------------------------------------
  --功能：检查病人是否执行一次性费用
  --入参：Json_In,格式如下
  --  input
  --    pati_id            N 1 病人ID
  --    pati_pageid        N 1 主页ID
  --    pati_wardarea_id   N 1 入院病区ID
  --    in_date            C 1 入院日期  格式：YYYY-MM-DD HH:MM:SS
  --出参: Json_Out,格式如下
  --  output
  --    code               N 1 应答吗：0-失败；1-成功
  --    message            C 1 应答消息：失败时返回具体的错误信息
  --    is_exe             N 1 执行标记:0-不执行;1-执行
  --------------------------------------------------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_Count            Number;
  n_Pati_Id          Number(18);
  n_Pati_Pageid      Number(18);
  n_Pati_Wardarea_Id Number(18);

  d_In_Date Date;
Begin

  --解析入参
  j_Input            := PLJson(Json_In);
  j_Json             := j_Input.Get_Pljson('input');
  n_Pati_Id          := j_Json.Get_Number('pati_id');
  n_Pati_Pageid      := j_Json.Get_Number('pati_pageid');
  n_Pati_Wardarea_Id := j_Json.Get_Number('pati_wardarea_id');

  d_In_Date := To_Date(j_Json.Get_String('in_date'), 'yyyy-mm-dd hh24:mi:ss');

  Select Count(*)
  Into n_Count
  From 自动计价项目 B
  Where b.病区id = n_Pati_Wardarea_Id And b.计算标志 = 8 And Nvl(b.启用日期, To_Date('3000-01-01', 'YYYY-MM-DD')) <= d_In_Date;
  If n_Count = 0 Then
    Json_Out := '{"output":{"code":1,"message":"成功","is_exe":0}}';
    Return;
  End If;

  --检查该病人本次住院是否已经计算过
  Select Count(*)
  Into n_Count
  From 住院费用记录
  Where 病人id = n_Pati_Id And 主页id = n_Pati_Pageid And 记录性质 = 3 And 记录状态 = 1 And 附加标志 = 8;
  If n_Count > 0 Then
    Json_Out := '{"output":{"code":1,"message":"成功","is_exe":0}}';
    Return;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功","is_exe":1}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkonetimefee;
/
Create Or Replace Procedure Zl_Exsesvr_Calconetimefee
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  --------------------------------------------------------------------------------------------------------------------
  --功能：对住院病人计算一次性费用。
  --入参：Json_In,格式如下
  --  input
  --    pati_id            N 1 病人ID
  --    pati_pageid        N 1 主页ID
  --    inpatient_num      C 1 住院号
  --    pati_wardarea_id   N 1 入院病区ID
  --    pati_deptid        N 1 入院科室ID
  --    medical_team_id    N 1 医疗小组ID
  --    pati_name          C 1 病人姓名
  --    pati_Sex           C 1 病人性别
  --    pati_age           C 1 病人年龄
  --    pati_bed           C 1 出院病床
  --    fee_category       C 1 费别
  --    in_date            C 1 入院日期
  --    func_id            N 1 功能ID 0-新增,1-删除
  --    mdlpay_mode_name   C 1 医疗付款方式名称
  --    operator_name      C 1 操作员姓名
  --    operator_code      C 1 操作员编号
  --    operator_deptid    N 1 操作员部门ID    
  --出参: Json_Out,格式如下
  --  output
  --    code               N 1 应答吗：0-失败；1-成功
  --    message            C 1 应答消息：失败时返回具体的错误信息
  --------------------------------------------------------------------------------------------------------------------
  j_Input Pljson;
  j_Json  Pljson;

  n_Func_Id       Number(1);
  n_Pati_Id       Number(18);
  n_Pati_Pageid   Number(6);
  v_Operator_Name Varchar2(100);
  v_Operator_Code Varchar2(100);
Begin
  --解析入参
  j_Input   := Pljson(Json_In);
  j_Json    := j_Input.Get_Pljson('input');
  n_Func_Id := j_Json.Get_Number('func_id');

  If Nvl(n_Func_Id, 0) = 0 Then
    Zl_住院一次费用_Insert_s(Json_In, Json_Out);
  Else
    n_Pati_Id       := j_Json.Get_Number('pati_id');
    n_Pati_Pageid   := j_Json.Get_Number('pati_pageid');
    v_Operator_Code := j_Json.Get_String('operator_code');
    v_Operator_Name := j_Json.Get_String('operator_name');
    Zl_住院一次费用_Delete(n_Pati_Id, n_Pati_Pageid, v_Operator_Code, v_Operator_Name);
  End If;
  Json_Out := Zljsonout('成功', 1);
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Calconetimefee;
/

Create Or Replace Procedure Zl_Exsesvr_Checkmrbkfeeisdel
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:检查病历费是否已经退费
  --入参：Json_In:格式
  --   input
  --      fee_no            C 1 单据号
  --      pati_id           N 1 病人id
  --      fee_properties    N 1 记录性质
  --出参  json
  --output
  --     code               N 1 应答码：0-失败；1-成功
  --     message            C 1 应答消息： 失败时返回具体的错误信息
  --     isdel              N 1 是否已退费:1-已经退费;0-未退费
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_Err_Msg Varchar2(500);

  v_No       住院费用记录.No%Type;
  n_病人id   住院费用记录.病人id%Type;
  n_记录性质 住院费用记录.记录性质%Type;
  n_Exist    Number(1);

Begin
  --解析入参

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_病人id   := j_Json.Get_Number('pati_id');
  v_No       := j_Json.Get_String('fee_no');
  n_记录性质 := j_Json.Get_Number('fee_properties');

  If Nvl(v_No, '-') = '-' Then
    v_Err_Msg := '未传入单据号，请检查!';
    Json_Out  := zlJsonOut(v_Err_Msg);
    Return;
  End If;

  If Nvl(n_病人id, 0) = 0 Then
    v_Err_Msg := '未传入病人id，请检查!';
    Json_Out  := zlJsonOut(v_Err_Msg);
    Return;
  End If;

  If Nvl(n_记录性质, 0) = 0 Then
    v_Err_Msg := '未传入记录性质，请检查!';
    Json_Out  := zlJsonOut(v_Err_Msg);
    Return;
  End If;

  If n_记录性质 = 4 Then
    Select Max(1)
    Into n_Exist
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 = 1 And 附加标志 = 1 And 病人id = n_病人id And NO = v_No;
  Else
    Select Max(1)
    Into n_Exist
    From 住院费用记录
    Where 记录性质 = 5 And 记录状态 = 1 And 附加标志 = 8 And 病人id = n_病人id And NO = v_No;
  End If;
  If Nvl(n_Exist, 0) = 0 Then
    n_Exist := 1;
  Else
    n_Exist := 0;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功","isdel":' || n_Exist || '}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkmrbkfeeisdel;
/


Create Or Replace Procedure Zl_Exsesvr_Getdepositblncsign
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:获取预交单据结算的异常状态
  --入参：Json_In:格式
  --   input
  --      deposit_no        C 1 预交单据号
  --出参  json
  --output
  --     code               N 1 应答码：0-失败；1-成功
  --     message            C 1 应答消息： 失败时返回具体的错误信息
  --     blance_sign        N 0-正常状态,1-发卡异常状态,2-退卡异常状态
  ---------------------------------------------------------------------------
  v_Err_Msg Varchar2(500);
  v_Output  Varchar2(32767);

  j_Input PLJson;
  j_Json  PLJson;

  v_No    病人预交记录.No%Type;
  n_Count Number(1);
  n_State Number(1);

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_No := j_Json.Get_String('deposit_no');

  If Nvl(v_No, '-') = '-' Then
    v_Err_Msg := '未传入单据号，请检查!';
    Json_Out  := zlJsonOut(v_Err_Msg);
    Return;
  End If;
  Select Count(1), Max(记录状态)
  Into n_Count, n_State
  From 病人预交记录
  Where 记录性质 = 1 And NO = v_No And Nvl(校对标志, 0) <> 0;
  If n_Count = 0 Then
    n_State := 0;
  Else
    If n_State = 0 Then
      n_State := 1;
    End If;
  End If;
  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '成功');
  zlJsonPutValue(v_Output, 'isdelfee', n_State, 1, 2);
  Json_Out := '{"output":' || v_Output || '}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getdepositblncsign;
/



Create Or Replace Procedure Zl_Exsesvr_Getcardfeeblncsign
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:获取病人的异常发卡数据
  --入参：Json_In:格式
  --   input
  --      pati_id           N 1 病人id
  --      operator_name     C 1 操作员姓名
  --出参  json
  --output
  --     code               N 1 应答码：0-失败；1-成功
  --     message            C 1 应答消息： 失败时返回具体的错误信息
  --     item_list[]
  --       balance_id       N 1 结帐id
  --       operator_name    C 1 操作员姓名
  --       err_type         N 1 异常类型:1-发卡收费异常;2-退费异常
  --       is_mrbk          N 1 是否病历费:1-是病历费;0-不是病历费
  --       create_time      C 1 登记时间:yyyy-mm-dd hh24:mi:ss
  --       rec_sign         N 1 记录状态
  --       balance_num      N 1 结算序号
  ---------------------------------------------------------------------------
  v_Err_Msg Varchar2(500);
  j_Input   PLJson;
  j_Json    PLJson;

  v_操作员 住院费用记录.操作员姓名%Type;
  n_病人id 住院费用记录.病人id%Type;
  v_Output Varchar2(32767);

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_病人id := j_Json.Get_Number('pati_id');
  v_操作员 := j_Json.Get_String('operator_name');

  If Nvl(n_病人id, 0) = 0 Then
    v_Err_Msg := '未传入病人id，请检查!';
    Json_Out  := zlJsonOut(v_Err_Msg);
    Return;
  End If;

  If Nvl(v_操作员, '-') = '-' Then
    v_Err_Msg := '未传入操作员姓名，请检查!';
    Json_Out  := zlJsonOut(v_Err_Msg);
    Return;
  End If;

  For r_异常数据 In (Select Distinct a.No, a.结帐id, a.操作员姓名, a.异常类型, a.登记时间, b.结算序号, a.记录状态, Decode(a.附加标志, 8, 1, 0) As 病历费
                 From (Select a.No, a.结帐id, a.操作员姓名, 1 As 异常类型, a.登记时间, a.记录状态, a.附加标志
                        From 住院费用记录 A
                        Where Nvl(费用状态, 0) = 1 And 记录性质 = 5 And 病人id = n_病人id And 记录状态 = 1 And 结帐id Is Not Null
                        Union All
                        Select a.No, a.结帐id, a.操作员姓名, 2 As 异常类型, a.登记时间, a.记录状态, 0 As 附加标志
                        From 住院费用记录 A
                        Where Nvl(费用状态, 0) = 1 And 记录性质 = 5 And 病人id = n_病人id And 记录状态 = 2 And Not Exists
                         (Select 1 From 病人预交记录 Where 结帐id = a.结帐id And Nvl(校对标志, 0) = 0)) A, 病人预交记录 B
                 Where a.结帐id = b.结帐id
                 Order By Decode(a.操作员姓名, v_操作员, 0, 1), a.记录状态) Loop
  
    zlJsonPutValue(v_Output, 'balance_id', r_异常数据.结帐id, 1, 1);
    zlJsonPutValue(v_Output, 'operator_name', r_异常数据.操作员姓名);
    zlJsonPutValue(v_Output, 'err_type', r_异常数据.异常类型, 1);
    zlJsonPutValue(v_Output, 'is_mrbk', r_异常数据.病历费, 1);
    zlJsonPutValue(v_Output, 'create_time', r_异常数据.登记时间);
    zlJsonPutValue(v_Output, 'rec_sign', r_异常数据.记录状态, 1);
    zlJsonPutValue(v_Output, 'balance_num', r_异常数据.结算序号, 1, 2);
  
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"成功","surplus_list":[' || v_Output || ']}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getcardfeeblncsign;
/


Create Or Replace Procedure Zl_Exsesvr_Getdepositdetail
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:获取指定条件下的预交收支明细数据
  --入参：Json_In:格式
  --   input
  --      pati_id           N 1 病人id
  --      begin_time        C 1 开始时间:yyyy-mm-dd hh24:mi:ss
  --      end_time          C 1 终止时间:yyyy-mm-dd hh24:mi:ss
  --      type              N 1 类型标志:0-所有,1-门诊;2-住院
  --出参  json
  --output
  --     code               N 1 应答码：0-失败；1-成功
  --     message            C 1 应答消息： 失败时返回具体的错误信息
  --     item_list[]
  --       business_type    C 1 业务类型:期初，充值，收费用、结帐等
  --       happen_time      C 1 发生时间:yyyy-mm-dd hh24:mi:ss
  --       earlystage       N 1 期初余额
  --       recharge         N 1 本期充值
  --       consume          N 1 本期消费
  ---------------------------------------------------------------------------
  v_Err_Msg  Varchar2(500);
  j_Input    PLJson;
  j_Json     PLJson;
  n_病人id   住院费用记录.病人id%Type;
  d_开始时间 Date;
  d_终止时间 Date;
  n_类型     Number(1);
  n_已转出   Number(1);
  d_上次日期 Zldatamove.上次日期%Type;

  v_Output Varchar2(32767);
  c_Output Clob;

Begin
  --解析入参

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_病人id   := j_Json.Get_Number('pati_id');
  d_开始时间 := To_Date(j_Json.Get_String('begin_time'), 'yyyy-mm-dd hh24:mi:ss');
  d_终止时间 := To_Date(j_Json.Get_String('end_time'), 'yyyy-mm-dd hh24:mi:ss');
  n_类型     := Nvl(j_Json.Get_Number('type'), 0);

  If Nvl(n_病人id, 0) = 0 Then
    v_Err_Msg := '未传入病人id，请检查!';
    Json_Out  := zlJsonOut(v_Err_Msg);
    Return;
  End If;

  Select Max(上次日期) Into d_上次日期 From zlDataMove Where 组号 = 1 And 系统 = 100 And 上次日期 Is Not Null;
  n_已转出 := 0;
  If d_上次日期 Is Not Null Then
    If d_上次日期 > d_开始时间 Then
      n_已转出 := 1;
    End If;
  End If;

  If n_已转出 = 1 Then
    --包含历史数据
    For r_预交明细 In (Select /*+ RULE */
                    类别, 收款时间, 业务类型, Sum(期初余额) As 期初余额, Sum(本期充值) As 本期充值, Sum(本期消费) As 本期消费
                   From (With 预交 As (Select 病人id, 收款时间, 0 As 类型, 结帐id, Nvl(金额, 0) As 金额, 0 As 冲预交
                                     From 病人预交记录 A
                                     Where a.收款时间 >= d_开始时间 And a.记录性质 = 1 And
                                           ((a.记录状态 In (1, 3) And Nvl(a.校对标志, 0) = 0) Or a.记录状态 = 2) And a.病人id = n_病人id And
                                           Nvl(a.预交类别, 2) In (1, 2)
                                     Union All
                                     Select a.病人id, b.收费时间 As 收款时间, 2 As 类型, b.Id As 结帐id, 0 As 金额, Nvl(冲预交, 0) As 冲预交
                                     From 病人预交记录 A, 病人结帐记录 B
                                     Where b.收费时间 >= d_开始时间 And Mod(a.记录性质, 10) = 1 And Nvl(a.校对标志, 0) = 0 And
                                           a.结帐id = b.Id And a.病人id = n_病人id And (Nvl(a.预交类别, 2) = n_类型 Or n_类型 = 0)
                                     Union All
                                     Select 病人id, 收费时间, 1 As 类型, 结帐id, 0 As 金额, Nvl(Sum(冲预交), 0) As 冲预交
                                     From (Select a.病人id, Min(b.登记时间) As 收费时间, a.No As 充值单据号, b.结帐id, 0 As 金额,
                                                   Max(Nvl(a.冲预交, 0)) As 冲预交
                                            From 病人预交记录 A, 门诊费用记录 B
                                            Where b.登记时间 >= d_开始时间 And Mod(a.记录性质, 10) = 1 And Nvl(a.校对标志, 0) = 0 And
                                                  Nvl(b.记帐费用, 0) = 0 And a.结帐id = b.结帐id And b.病人id = n_病人id And
                                                  b.记录性质 In (1, 4) And Nvl(a.冲预交, 0) <> 0 And
                                                  (Nvl(a.预交类别, 2) = n_类型 Or n_类型 = 0)
                                            Group By a.病人id, a.No, b.结帐id)
                                     Group By 病人id, 收费时间, 结帐id
                                     Union All
                                     Select 病人id, 收费时间, 1 As 类型, 结帐id, 0 As 金额, Nvl(Sum(冲预交), 0) As 冲预交
                                     From (Select a.病人id, Min(b.登记时间) As 收费时间, a.No As 充值单据号, b.结帐id, 0 As 金额,
                                                   Max(Nvl(a.冲预交, 0)) As 冲预交
                                            From 病人预交记录 A, 住院费用记录 B
                                            Where b.登记时间 >= d_开始时间 And Mod(a.记录性质, 10) = 1 And Nvl(a.校对标志, 0) = 0 And
                                                  a.结帐id = b.结帐id And b.病人id = n_病人id And b.记录性质 = 5 And Nvl(b.记帐费用, 0) = 0 And
                                                  Nvl(a.冲预交, 0) <> 0
                                            Group By a.病人id, a.No, b.结帐id)
                                     Group By 病人id, 收费时间, 结帐id
                                     Union All
                                     --历史预交记录
                                     Select 病人id, 收款时间, 0 As 类型, 结帐id, Nvl(金额, 0) As 金额, 0 As 冲预交
                                     From H病人预交记录 A
                                     Where a.收款时间 >= d_开始时间 And a.记录性质 = 1 And
                                           ((a.记录状态 In (1, 3) And Nvl(a.校对标志, 0) = 0) Or a.记录状态 = 2) And a.病人id = n_病人id And
                                           Nvl(a.预交类别, 2) In (1, 2)
                                     Union All
                                     Select a.病人id, b.收费时间 As 收款时间, 2 As 类型, b.Id As 结帐id, 0 As 金额, Nvl(冲预交, 0) As 冲预交
                                     From H病人预交记录 A, 病人结帐记录 B
                                     Where b.收费时间 >= d_开始时间 And Mod(a.记录性质, 10) = 1 And Nvl(a.校对标志, 0) = 0 And
                                           a.结帐id = b.Id And a.病人id = n_病人id And (Nvl(a.预交类别, 2) = n_类型 Or n_类型 = 0)
                                     Union All
                                     Select 病人id, 收费时间, 1 As 类型, 结帐id, 0 As 金额, Nvl(Sum(冲预交), 0) As 冲预交
                                     From (Select a.病人id, Min(b.登记时间) As 收费时间, a.No As 充值单据号, b.结帐id, 0 As 金额,
                                                   Max(Nvl(a.冲预交, 0)) As 冲预交
                                            From H病人预交记录 A, H门诊费用记录 B
                                            Where b.登记时间 >= d_开始时间 And Mod(a.记录性质, 10) = 1 And Nvl(a.校对标志, 0) = 0 And
                                                  Nvl(b.记帐费用, 0) = 0 And a.结帐id = b.结帐id And b.病人id = n_病人id And
                                                  b.记录性质 In (1, 4) And Nvl(a.冲预交, 0) <> 0 And
                                                  (Nvl(a.预交类别, 2) = n_类型 Or n_类型 = 0)
                                            Group By a.病人id, a.No, b.结帐id)
                                     Group By 病人id, 收费时间, 结帐id
                                     Union All
                                     Select 病人id, 收费时间, 1 As 类型, 结帐id, 0 As 金额, Nvl(Sum(冲预交), 0) As 冲预交
                                     From (Select a.病人id, Min(b.登记时间) As 收费时间, a.No As 充值单据号, b.结帐id, 0 As 金额,
                                                   Max(Nvl(a.冲预交, 0)) As 冲预交
                                            From H病人预交记录 A, H住院费用记录 B
                                            Where b.登记时间 >= d_开始时间 And Mod(a.记录性质, 10) = 1 And Nvl(a.校对标志, 0) = 0 And
                                                  a.结帐id = b.结帐id And b.病人id = n_病人id And b.记录性质 = 5 And Nvl(b.记帐费用, 0) = 0 And
                                                  Nvl(a.冲预交, 0) <> 0
                                            Group By a.病人id, a.No, b.结帐id)
                                     Group By 病人id, 收费时间, 结帐id
                                     
                                     )
                          Select 0 As 类别, '' As 收款时间, '期初' As 业务类型, Sum(Nvl(预交余额, 0)) As 期初余额, 0 As 本期充值, 0 As 本期消费
                          From 病人余额 A
                          Where 病人id = n_病人id And 性质 = 1 And (Nvl(a.类型, 2) = n_类型 Or n_类型 = 0)
                          Union All
                          Select 0 As 类别, '' As 收款时间, '期初' As 业务类型, -1 * Sum(Nvl(金额, 0)) + Sum(Nvl(冲预交, 0)) As 期初余额, 0,
                                 0 As 本期消费
                          From 预交
                          Where 收款时间 >= d_开始时间
                          Group By To_Char(收款时间, 'yyyy-mm-dd')
                          Union All
                          Select 1 As 类别, To_Char(收款时间, 'yyyy-mm-dd') As 收款时间, '充值' As 业务类型, 0 As 期初余额,
                                 Sum(Nvl(金额, 0)) As 充值, 0 As 本期消费
                          From 预交
                          Where 收款时间 Between d_开始时间 And d_终止时间 And Nvl(金额, 0) > 0
                          Group By To_Char(收款时间, 'yyyy-mm-dd')
                          Union All
                          Select 1 As 类别, To_Char(收款时间, 'yyyy-mm-dd') As 收款时间, Decode(类型, 1, '收费', 2, '结帐', '消费') As 业务类型,
                                 0 As 期初余额, 0 As 充值, Sum(Nvl(冲预交, 0)) As 消费
                          From 预交
                          Where 收款时间 Between d_开始时间 And d_终止时间 Having Sum(Nvl(冲预交, 0)) <> 0
                          Group By To_Char(收款时间, 'yyyy-mm-dd'), Decode(类型, 1, '收费', 2, '结帐', '消费')
                          Union All
                          Select 2 As 类别, To_Char(收款时间, 'yyyy-mm-dd') As 收款时间, '充值' As 业务类型, 0 As 期初余额,
                                 Sum(Nvl(金额, 0)) As 充值, 0 As 本期消费
                          From 预交
                          Where 收款时间 Between d_开始时间 And d_终止时间 And Nvl(金额, 0) < 0
                          Group By To_Char(收款时间, 'yyyy-mm-dd'))
                          Group By 类别, 收款时间, 业务类型
                          Order By 类别, 收款时间
                   ) Loop
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      
        v_Output := Null;
      End If;
    
      zlJsonPutValue(v_Output, 'business_type', r_预交明细.业务类型, 0, 1);
      zlJsonPutValue(v_Output, 'happen_time', r_预交明细.收款时间);
      zlJsonPutValue(v_Output, 'earlystage', r_预交明细.期初余额, 1);
      zlJsonPutValue(v_Output, 'recharge', r_预交明细.本期充值, 1);
      zlJsonPutValue(v_Output, 'consume', r_预交明细.本期消费, 1, 2);
    
    End Loop;
  
  Else
    For r_预交明细 In (Select /*+ RULE */
                    类别, 收款时间, 业务类型, Sum(期初余额) As 期初余额, Sum(本期充值) As 本期充值, Sum(本期消费) As 本期消费
                   From (With 预交 As (Select 病人id, 收款时间, 0 As 类型, 结帐id, Nvl(金额, 0) As 金额, 0 As 冲预交
                                     From 病人预交记录 A
                                     Where a.收款时间 >= d_开始时间 And a.记录性质 = 1 And
                                           ((a.记录状态 In (1, 3) And Nvl(a.校对标志, 0) = 0) Or a.记录状态 = 2) And a.病人id = n_病人id And
                                           Nvl(a.预交类别, 2) In (1, 2)
                                     Union All
                                     Select a.病人id, b.收费时间 As 收款时间, 2 As 类型, b.Id As 结帐id, 0 As 金额, Nvl(冲预交, 0) As 冲预交
                                     From 病人预交记录 A, 病人结帐记录 B
                                     Where b.收费时间 >= d_开始时间 And Mod(a.记录性质, 10) = 1 And Nvl(a.校对标志, 0) = 0 And
                                           a.结帐id = b.Id And a.病人id = n_病人id And (Nvl(a.预交类别, 2) = n_类型 Or n_类型 = 0)
                                     Union All
                                     Select 病人id, 收费时间, 1 As 类型, 结帐id, 0 As 金额, Nvl(Sum(冲预交), 0) As 冲预交
                                     From (Select a.病人id, Min(b.登记时间) As 收费时间, a.No As 充值单据号, b.结帐id, 0 As 金额,
                                                   Max(Nvl(a.冲预交, 0)) As 冲预交
                                            From 病人预交记录 A, 门诊费用记录 B
                                            Where b.登记时间 >= d_开始时间 And Mod(a.记录性质, 10) = 1 And Nvl(a.校对标志, 0) = 0 And
                                                  Nvl(b.记帐费用, 0) = 0 And a.结帐id = b.结帐id And b.病人id = n_病人id And
                                                  b.记录性质 In (1, 4) And Nvl(a.冲预交, 0) <> 0 And
                                                  (Nvl(a.预交类别, 2) = n_类型 Or n_类型 = 0)
                                            Group By a.病人id, a.No, b.结帐id)
                                     Group By 病人id, 收费时间, 结帐id
                                     Union All
                                     Select 病人id, 收费时间, 1 As 类型, 结帐id, 0 As 金额, Nvl(Sum(冲预交), 0) As 冲预交
                                     From (Select a.病人id, Min(b.登记时间) As 收费时间, a.No As 充值单据号, b.结帐id, 0 As 金额,
                                                   Max(Nvl(a.冲预交, 0)) As 冲预交
                                            From 病人预交记录 A, 住院费用记录 B
                                            Where b.登记时间 >= d_开始时间 And Mod(a.记录性质, 10) = 1 And Nvl(a.校对标志, 0) = 0 And
                                                  a.结帐id = b.结帐id And b.病人id = n_病人id And b.记录性质 = 5 And Nvl(b.记帐费用, 0) = 0 And
                                                  Nvl(a.冲预交, 0) <> 0
                                            Group By a.病人id, a.No, b.结帐id)
                                     Group By 病人id, 收费时间, 结帐id)
                          Select 0 As 类别, '' As 收款时间, '期初' As 业务类型, Sum(Nvl(预交余额, 0)) As 期初余额, 0 As 本期充值, 0 As 本期消费
                          From 病人余额 A
                          Where 病人id = n_病人id And 性质 = 1 And (Nvl(a.类型, 2) = n_类型 Or n_类型 = 0)
                          Union All
                          Select 0 As 类别, '' As 收款时间, '期初' As 业务类型, -1 * Sum(Nvl(金额, 0)) + Sum(Nvl(冲预交, 0)) As 期初余额, 0,
                                 0 As 本期消费
                          From 预交
                          Where 收款时间 >= d_开始时间
                          Group By To_Char(收款时间, 'yyyy-mm-dd')
                          Union All
                          Select 1 As 类别, To_Char(收款时间, 'yyyy-mm-dd') As 收款时间, '充值' As 业务类型, 0 As 期初余额,
                                 Sum(Nvl(金额, 0)) As 充值, 0 As 本期消费
                          From 预交
                          Where 收款时间 Between d_开始时间 And d_终止时间 And Nvl(金额, 0) > 0
                          Group By To_Char(收款时间, 'yyyy-mm-dd')
                          Union All
                          Select 1 As 类别, To_Char(收款时间, 'yyyy-mm-dd') As 收款时间, Decode(类型, 1, '收费', 2, '结帐', '消费') As 业务类型,
                                 0 As 期初余额, 0 As 充值, Sum(Nvl(冲预交, 0)) As 消费
                          From 预交
                          Where 收款时间 Between d_开始时间 And d_终止时间 Having Sum(Nvl(冲预交, 0)) <> 0
                          Group By To_Char(收款时间, 'yyyy-mm-dd'), Decode(类型, 1, '收费', 2, '结帐', '消费')
                          Union All
                          Select 2 As 类别, To_Char(收款时间, 'yyyy-mm-dd') As 收款时间, '充值' As 业务类型, 0 As 期初余额,
                                 Sum(Nvl(金额, 0)) As 充值, 0 As 本期消费
                          From 预交
                          Where 收款时间 Between d_开始时间 And d_终止时间 And Nvl(金额, 0) < 0
                          Group By To_Char(收款时间, 'yyyy-mm-dd'))
                          Group By 类别, 收款时间, 业务类型
                          Order By 类别, 收款时间
                   ) Loop
    
      --     item_list[]
      --       business_type    C 1 业务类型:期初，充值，收费用、结帐等
      --       happen_time      C 1 发生时间:yyyy-mm-dd hh24:mi:ss
      --       earlystage       N 1 期初余额
      --       recharge         N 1 本期充值
      --       consume          N 1 本期消费
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      
        v_Output := Null;
      End If;
    
      zlJsonPutValue(v_Output, 'business_type', r_预交明细.业务类型, 0, 1);
      zlJsonPutValue(v_Output, 'happen_time', r_预交明细.收款时间);
      zlJsonPutValue(v_Output, 'earlystage', r_预交明细.期初余额, 1);
      zlJsonPutValue(v_Output, 'recharge', r_预交明细.本期充值, 1);
      zlJsonPutValue(v_Output, 'consume', r_预交明细.本期消费, 1, 2);
    
    End Loop;
  
  End If;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"成功","item_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || v_Output || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getdepositdetail;
/

Create Or Replace Procedure Zl_Exsesvr_Getdepositlist
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:获取指定病人的预交清单信息
  --入参：Json_In:格式
  --   input
  --      pati_id           N 1 病人id
  --      pati_pageid       N 1 主页ID
  --      type              N 1 预交类别 1-门诊;2-住院
  --出参  json
  --output
  --     code               N 1 应答码：0-失败；1-成功
  --     message            C 1 应答消息： 失败时返回具体的错误信息
  --     item_list[]        数组
  --       create_date      C 1 收款日期
  --       bill_no          C 1 单据号
  --       dept_name        C 1 科室名称
  --       money            N 1 金额
  --       blnc_mode        N 1 结算方式
  --       operator_name    C 1 操作员姓名
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_Output      Varchar2(32767);
  c_Output      Clob;
  n_Pati_Id     Number(18);
  n_Pati_Pageid Number(18);
  n_Type        Number(3);

  v_Err_Msg Varchar2(500);
Begin

  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_Pati_Id     := j_Json.Get_Number('pati_id');
  n_Pati_Pageid := j_Json.Get_Number('pati_pageid');
  n_Type        := j_Json.Get_Number('type');

  If Nvl(n_Pati_Id, 0) = 0 Then
    v_Err_Msg := '未传入病人id，请检查!';
    Json_Out  := zlJsonOut(v_Err_Msg);
    Return;
  
  End If;
  For R In (Select LTrim(To_Char(a.收款时间, 'YYYY-MM-DD')) As 日期, a.No, b.名称 As 科室, a.金额, a.结算方式, a.操作员姓名
            From 病人预交记录 A, 部门表 B
            Where a.科室id = b.Id(+) And a.记录性质 = 1 And a.病人id = n_Pati_Id And
                  (a.主页id = n_Pati_Pageid Or Nvl(n_Pati_Pageid, 0) = 0) And a.预交类别 = n_Type
            Order By a.收款时间 Desc) Loop
  
    If Length(Nvl(v_Output, ' ')) > 30000 Then
      If c_Output Is Not Null Then
        c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
      Else
        c_Output := To_Clob(v_Output);
      End If;
    
      v_Output := Null;
    End If;
  
    zlJsonPutValue(v_Output, 'create_date', r.日期, 0, 1);
    zlJsonPutValue(v_Output, 'bill_no', r.No);
    zlJsonPutValue(v_Output, 'dept_name', r.科室);
    zlJsonPutValue(v_Output, 'money', r.金额);
    zlJsonPutValue(v_Output, 'blnc_mode', r.结算方式);
    zlJsonPutValue(v_Output, 'operator_name', r.操作员姓名, 0, 2);
  
  End Loop;
  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"成功","item_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || v_Output || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getdepositlist;
/

Create Or Replace Procedure Zl_Exsesvr_Getwriteoffinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  --------------------------------------------------------------------------------------------------------------------
  --功能:获取本次住院病人销帐申请信息
  --入参：Json_In,格式如下
  --  input
  --    pati_id            N 1 病人ID
  --    pati_pageid        N 1 主页ID
  --    request_time       C 1 销帐申请时间

  --    type               N 1 0-检查是否存在未处理的销帐申请;1-获取本次住院病人未审核的销帐申请
  --出参: Json_Out,格式如下
  --  output
  --    code               N 1 应答吗：0-失败；1-成功
  --    message            C 1 应答消息：失败时返回具体的错误信息
  --    isexists           N 1   type=0或2 时返回 1-存在;0-不存在
  --    fee_list[]未审核销帐的单据信息 type=1时返回
  --      no               C  1   单据号
  --      fee_name         C  1   收费项目名称
  --      dept_name        C  1   部门名称

  --------------------------------------------------------------------------------------------------------------------
  j_Input Pljson;
  j_Json  Pljson;

  n_Pati_Id     住院费用记录.病人id%Type;
  n_Pati_Pageid 住院费用记录.主页id%Type;
  n_Type        Number;
  n_Count       Number;
  Vjtmp         Varchar2(32767);
  d_申请时间    Date;
Begin

  --解析入参
  j_Input := Pljson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_Pati_Id     := j_Json.Get_Number('pati_id');
  n_Pati_Pageid := j_Json.Get_Number('pati_pageid');
  n_Type        := Nvl(j_Json.Get_Number('type'), 0);

  If n_Type = 0 Then
    Select Count(1)
    Into n_Count
    From 住院费用记录 A, 病人费用销帐 B
    Where a.病人id = n_Pati_Id And a.主页id = n_Pati_Pageid And b.费用id = a.Id And b.状态 = 0;
    If n_Count > 1 Then
      n_Count := 1;
    End If;
    Json_Out := '{"output":{"code":1,"message": "成功","isexists":' || n_Count || '}}';
  Elsif n_Type = 1 Then
    Vjtmp := Null;
    For R In (Select Distinct a.No, d.名称 As 项目, c.名称 As 部门
              From 住院费用记录 A, 病人费用销帐 B, 部门表 C, 收费项目目录 D
              Where a.Id = b.费用id And a.收费细目id = d.Id And b.审核部门id = c.Id(+) And b.审核时间 Is Null And a.病人id = n_Pati_Id And
                    Nvl(a.主页id, 0) = n_Pati_Pageid) Loop
    
      Vjtmp := Vjtmp || ',{';
      Vjtmp := Vjtmp || '"no":"' || r.No || '"';
      Vjtmp := Vjtmp || ',"fee_name":"' || Zljsonstr(r.项目) || '"';
      Vjtmp := Vjtmp || ',"dept_name":"' || Zljsonstr(r.部门) || '"';
      Vjtmp := Vjtmp || '}';
    End Loop;
    Json_Out := '{"output":{"code":1,"message": "成功","fee_list":[' || Substr(Vjtmp, 2) || ']}}';
  Elsif n_Type = 2 Then
    d_申请时间 := To_Date(j_Json.Get_String('request_time'), 'yyyy-mm-dd hh24:mi:ss');
    Select Count(1) Into n_Count From 病人费用销帐 B Where b.申请时间 = d_申请时间;
    Json_Out := '{"output":{"code":1,"message": "成功","isexists":' || n_Count || '}}';
  End If;
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getwriteoffinfo;
/


Create Or Replace Procedure Zl_Exsesvr_Chargeissuccessed
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：功能：检查卡费或预交款是否已经结算成功
  --入参：Json_In:格式
  --  input
  --    cardfee_no       C  1 卡费对应的费用单据号
  --    deposit_no       C  1 预交单据号
  --出参: Json_Out,格式如下
  --  output
  --    code              N   1 应答码：0-失败；1-成功
  --    message           C   1 应答消息：失败时返回具体的错误信息
  --    is_successed      N   1 是否成功:1-成功;0-不成功
  ---------------------------------------------------------------------------

  j_Input PLJson;
  j_Json  PLJson;

  v_Card_No    住院费用记录.No%Type;
  v_Deposit_No 病人预交记录.No%Type;
  n_Exist      Number(1);
  v_Output     Varchar2(32767);

Begin

  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_Card_No    := j_Json.Get_String('cardfee_no');
  v_Deposit_No := j_Json.Get_String('deposit_no');

  If Nvl(v_Card_No, '-') = '-' Or Nvl(v_Deposit_No, '-') = '-' Then
    Json_Out := zlJsonOut('失败，必须传入预交NO和费用NO！');
    Return;
  End If;

  Select Count(1)
  Into n_Exist
  From 住院费用记录 A, 病人结帐记录 B
  Where a.结帐id = b.Id And a.记录性质 In (5, 15) And a.记录状态 = 1 And b.记录状态 = 1 And a.No = v_Card_No;

  If n_Exist = 0 Then
    Select Count(1)
    Into n_Exist
    From 病人预交记录
    Where NO = v_Deposit_No And 记录性质 = 5 And 记录状态 In (1, 3) And 校对标志 In (0, 2);
  End If;

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '成功');
  zlJsonPutValue(v_Output, 'is_successed', n_Exist, 1, 2);
  Json_Out := '{"output":' || v_Output || '}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Chargeissuccessed;
/

Create Or Replace Procedure Zl_Exsesvr_Getpatifee
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------------------------------------------------------------
  --功能：获取病人费用相关信息
  --入参：Json_In格式
  --  input
  --    pati_id            N 1 病人ID
  --    pati_pageid        N 1 主页ID
  --    fee_type           C 0 收费类别
  --    baby_num           N 0 婴儿费
  --出参：json_out格式
  --fee_list      [数组]  每条费用记录
  --  id               N    费用id
  --  no               N    单据号
  --  fee_id           N    收费细目id
  -------------------------------------------------------------------------------------------------
  j_Input       PLJson;
  j_Json        PLJson;
  v_Tmp         Varchar2(32767);
  n_Pati_Id     Number(18);
  n_Pati_Pageid Number(18);
  n_Baby_Num    Number(3);
  v_Fee_Type    住院费用记录.收费类别%Type;
Begin
  --解析入参

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_Pati_Id     := j_Json.Get_Number('pati_id');
  n_Pati_Pageid := j_Json.Get_Number('pati_pageid');
  v_Fee_Type    := j_Json.Get_String('fee_type');
  n_Baby_Num    := j_Json.Get_Number('baby_num');
  For R In (Select a.Id, a.No, a.收费细目id
            From 住院费用记录 A
            Where a.病人id = n_Pati_Id And a.主页id = n_Pati_Pageid And (a.收费类别 = v_Fee_Type Or '空' = Nvl(v_Fee_Type, '空')) And
                  (Nvl(a.婴儿费, 0) = n_Baby_Num Or - 1 = n_Baby_Num)) Loop
    v_Tmp := v_Tmp || ',' || '{"id":' || r.Id || ',"no": "' || r.No || '","fee_id":' || r.收费细目id || '}';
  End Loop;
  Json_Out := '{"output":{"code":1,"message": "成功","fee_list":[' || Substr(v_Tmp, 2) || ']}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getpatifee;
/

Create Or Replace Procedure Zl_Exsesvr_Getreceiveinvoice
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:获取票据领用信息
  --入参：Json_In:格式
  --    input
  --      oper_fun  N 1 0-获取票据领用信息 1-获取获取指定票种的共用票据批次
  --      recv_ids C 1 领用ids:票据领用id,多个用逗号
  --      inv_type  N 1 票种:1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡
  --      use_mode  N 1 使用方式:1-自用：该票据仅供领用者自己使用；2-共用：该票据由多个人员共同使用
  --      use_type C 1 票据使用类别:1,4: 票据使用类别.名称;2预见:1-门诊预交;2-住院预交;5:存储的是医疗卡类别.ID
  --      recvtr  C 1 领用人
  --      min_nums  N 1 发票最少数量
  --      nodeno  C  1  站点
  --出参: Json_Out,格式如下
  --    output
  --    code  C 1 应答码：0-失败；1-成功
  --    message C 1 "应答消息： 成功时返回成功信息,失败时返回具体的错误信息"
  --    item_list C
  --      recv_id N 1 领用ID
  --      use_mode  N 1 使用方式:1-自用：该票据仅供领用者自己使用；2-共用：该票据由多个人员共同使用
  --      use_type C 1 票据使用类别:1,4: 票据使用类别.名称;2预见:1-门诊预交;2-住院预交;5:存储的是医疗卡类别.ID
  --      prefix_text C 1 前缀文本
  --      start_no  C 1 开始号码
  --      end_no  C 1 终止号码
  --      inv_no_cur  C 1 当前号码
  --      surplus_num C 1 剩余数量
  --      create_time C 1 登记时间:yyyy-mm-dd hh24:mi:ss
  --      use_time  C 1 使用时间:yyyy-mm-dd hh24:mi:ss
  --      recvtr  C 1 领用人
  --      use_typecode      C 1 使用类别编码
  --      use_typeid        N 1 使用类别id

  ---------------------------------------------------------------------------
  j_Input  PLJson;
  j_Json   PLJson;
  v_Output Varchar2(32767);
  c_Output Clob;

  n_领用id       Number(18);
  n_票种         Number(2);
  n_使用方式     Number(2);
  v_票据使用类别 Varchar2(100);
  v_领用人       Varchar2(100);
  v_领用ids      Varchar2(4000);
  n_操作方式     Number;
  v_Node         Varchar2(100);

  n_最少数量 Number(18);

  Cursor c_票据领用信息 Is
    Select ID, 前缀文本, 当前号码, 开始号码, 终止号码, 剩余数量, 使用方式, 使用类别, 领用人, To_Char(登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间,
           To_Char(使用时间, 'yyyy-mm-dd hh24:mi:ss') As 使用时间
    From 票据领用记录
    Where Rownum < 1;

  Cursor c_票据批次信息 Is
    Select a.Id, '' As 使用类别编码, a.使用类别 As 使用类别id, a.类别名称 As 使用类别, a.领用人, a.登记时间, a.开始号码, a.终止号码, a.剩余数量
    From 票据领用记录 A, 人员表 B
    Where a.使用方式 = 2 And a.剩余数量 > 0 And a.领用人 = b.姓名 And Rownum < 1;

  r_票据领用 c_票据领用信息%RowType;

  Type Ty_票据领用 Is Ref Cursor;
  c_票据领用 Ty_票据领用; --动态游标变量

  r_票据批次 c_票据批次信息%RowType;

  Type Ty_票据批次 Is Ref Cursor;
  c_票据批次 Ty_票据批次; --动态游标变量

Begin
  --解析入参

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_领用ids      := j_Json.Get_String('recv_ids');
  n_票种         := j_Json.Get_Number('inv_type');
  n_使用方式     := j_Json.Get_Number('use_mode');
  v_票据使用类别 := j_Json.Get_String('use_type');
  v_领用人       := j_Json.Get_String('recvtr');
  n_最少数量     := j_Json.Get_Number('min_nums');
  n_操作方式     := j_Json.Get_Number('oper_fun');
  v_Node         := j_Json.Get_String('nodeno');

  If Nvl(n_票种, 0) = 0 And v_领用ids Is Null Then
    Json_Out := zlJsonOut('未传入票种信息!');
    Return;
  End If;
  If v_领用ids Is Not Null Then
    If Instr(v_领用ids, ',') = 0 Then
      n_领用id := To_Number(v_领用ids);
    End If;
  End If;
  If Nvl(n_操作方式, 0) = 0 Then
    If v_领用ids Is Not Null Then
      --按病人ID为主要查询条件进行查询
      If Nvl(n_领用id, 0) <> 0 Then
        Open c_票据领用 For
          Select ID, 前缀文本, 当前号码, 开始号码, 终止号码, 剩余数量, 使用方式, 使用类别, 领用人, To_Char(登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间,
                 To_Char(使用时间, 'yyyy-mm-dd hh24:mi:ss') As 使用时间
          From 票据领用记录
          Where ID = n_领用id And (Nvl(n_票种, 0) = 0 Or 票种 = n_票种) And 剩余数量 > 0 And
                (Nvl(使用类别, 'LXH') = v_票据使用类别 Or 使用类别 Is Null Or v_票据使用类别 Is Null) And Nvl(剩余数量, 0) >= Nvl(n_最少数量, 0);
      Else
        Open c_票据领用 For
          Select ID, 前缀文本, 当前号码, 开始号码, 终止号码, 剩余数量, 使用方式, 使用类别, 领用人, To_Char(登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间,
                 To_Char(使用时间, 'yyyy-mm-dd hh24:mi:ss') As 使用时间
          From 票据领用记录
          Where ID In (Select /*+cardinality(b,10) */
                        Column_Value
                       From Table(f_Num2List(v_领用ids)) B) And (Nvl(n_票种, 0) = 0 Or 票种 = n_票种) And 剩余数量 > 0 And
                (Nvl(使用类别, 'LXH') = v_票据使用类别 Or 使用类别 Is Null Or v_票据使用类别 Is Null) And Nvl(剩余数量, 0) >= Nvl(n_最少数量, 0)
          Order By Nvl(使用时间, To_Date('1900-01-01', 'YYYY-MM-DD')) Desc, 使用类别 Desc, 开始号码;
      End If;
    Else
      Open c_票据领用 For
        Select ID, 前缀文本, 当前号码, 开始号码, 终止号码, 剩余数量, 使用方式, 使用类别, 领用人, To_Char(登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间,
               To_Char(使用时间, 'yyyy-mm-dd hh24:mi:ss') As 使用时间
        From 票据领用记录
        Where 票种 = n_票种 And 使用方式 = Nvl(n_使用方式, 0) And 剩余数量 > 0 And 领用人 = v_领用人 And
              (Nvl(使用类别, 'LXH') = v_票据使用类别 Or 使用类别 Is Null) And Nvl(剩余数量, 0) >= Nvl(n_最少数量, 0)
        Order By Nvl(使用时间, To_Date('1900-01-01', 'YYYY-MM-DD')) Desc, 使用类别 Desc, 开始号码;
    End If;
  
    Loop
      Fetch c_票据领用
        Into r_票据领用;
      Exit When c_票据领用%NotFound;
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      
        v_Output := Null;
      End If;
    
      zlJsonPutValue(v_Output, 'recv_id', r_票据领用.Id, 1, 1);
      zlJsonPutValue(v_Output, 'use_mode', r_票据领用.使用方式);
      zlJsonPutValue(v_Output, 'use_type', Nvl(r_票据领用.使用类别, ''));
      zlJsonPutValue(v_Output, 'prefix_text', Nvl(r_票据领用.前缀文本, ''));
      zlJsonPutValue(v_Output, 'start_no', Nvl(r_票据领用.开始号码, ''));
      zlJsonPutValue(v_Output, 'end_no', Nvl(r_票据领用.终止号码, ''));
      zlJsonPutValue(v_Output, 'inv_no_cur', Nvl(r_票据领用.当前号码, ''));
      zlJsonPutValue(v_Output, 'surplus_num', r_票据领用.剩余数量, 1);
      zlJsonPutValue(v_Output, 'create_time', r_票据领用.登记时间);
      zlJsonPutValue(v_Output, 'use_time', r_票据领用.使用时间);
      zlJsonPutValue(v_Output, 'recvtr', r_票据领用.领用人, 0, 2);
    
    End Loop;
  
  Else
    If Nvl(n_票种, 0) = 1 Or Nvl(n_票种, 0) = 3 Then
      --收费和结帐
      Open c_票据批次 For
        Select a.Id, Nvl(m.编码, ' ') As 使用类别编码, Null 使用类别id, a.使用类别, a.领用人, a.登记时间, a.开始号码, a.终止号码, a.剩余数量
        From 票据领用记录 A, 人员表 B, 票据使用类别 M
        Where a.票种 = n_票种 And a.使用方式 = 2 And a.剩余数量 > 0 And a.领用人 = b.姓名 And a.使用类别 = m.名称(+) And
              (b.站点 = v_Node Or b.站点 Is Null)
        Order By 使用类别编码, 剩余数量 Desc;
    Elsif Nvl(n_票种, 0) = 5 Then
      --就诊卡
      Open c_票据批次 For
        Select a.Id, Null As 使用类别编码, a.使用类别 As 使用类别id, a.使用类别, a.领用人, a.登记时间, a.开始号码, a.终止号码, a.剩余数量
        From 票据领用记录 A, 人员表 B
        Where a.票种 = n_票种 And a.使用方式 = 2 And a.剩余数量 > 0 And a.领用人 = b.姓名 And (b.站点 = v_Node Or b.站点 Is Null)
        Order By 使用类别编码, 剩余数量 Desc;
    Elsif Nvl(n_票种, 0) = 2 Then
      --预交
      Open c_票据批次 For
        Select a.Id, Null 使用类别编码, Null 使用类别id, To_Number(Nvl(a.使用类别, '0')) As 使用类别, a.领用人, a.登记时间, a.开始号码, a.终止号码,
               a.剩余数量
        From 票据领用记录 A, 人员表 B
        Where a.票种 = n_票种 And a.使用方式 = 2 And a.剩余数量 > 0 And a.领用人 = b.姓名 And (b.站点 = v_Node Or b.站点 Is Null)
        Order By 使用类别, 剩余数量 Desc;
    Else
      Open c_票据批次 For
        Select a.Id, Null 使用类别编码, Null 使用类别id, a.使用类别, a.领用人, a.登记时间, a.开始号码, a.终止号码, a.剩余数量
        From 票据领用记录 A, 人员表 B
        Where a.票种 = n_票种 And a.使用方式 = 2 And a.剩余数量 > 0 And a.领用人 = b.姓名 And (b.站点 = v_Node Or b.站点 Is Null)
        Order By 使用类别, 剩余数量 Desc;
    End If;
    Loop
      Fetch c_票据批次
        Into r_票据批次;
      Exit When c_票据批次%NotFound;
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      
        v_Output := Null;
      End If;
    
      zlJsonPutValue(v_Output, 'recv_id', r_票据批次.Id, 1, 1);
      zlJsonPutValue(v_Output, 'use_typecode', r_票据批次.使用类别编码);
      zlJsonPutValue(v_Output, 'use_typeid', r_票据批次.使用类别id, 1);
      zlJsonPutValue(v_Output, 'use_type', r_票据批次.使用类别);
      zlJsonPutValue(v_Output, 'recvtr', r_票据批次.领用人);
      zlJsonPutValue(v_Output, 'create_time', r_票据批次.登记时间);
      zlJsonPutValue(v_Output, 'start_no', r_票据批次.开始号码);
      zlJsonPutValue(v_Output, 'end_no', r_票据批次.终止号码);
      zlJsonPutValue(v_Output, 'surplus_num', r_票据批次.剩余数量, 1, 2);
    
    End Loop;
  
  End If;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"成功","item_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || v_Output || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getreceiveinvoice;
/



Create Or Replace Procedure Zl_Exsesvr_Getnextinvoice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:根据当前发票号及票据使用明细，获取一个有效的发票号
  --入参：Json_In:格式
  -- input
  --   recv_id N 1 领用id:票据领用id
  --   inv_no  C 1 发票号

  --出参: Json_Out,格式如下
  -- output
  --   code  C 1 应答码：0-失败；1-成功
  --   message C 1 应答消息：  成功时返回成功信息 失败时返回具体的错误信息
  --   inv_no  C 1 下一个发票号
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_领用id     Number(18);
  v_当前发票号 Varchar2(100);
  n_Count      Number(2);

  v_前缀文本 票据领用记录.前缀文本%Type;
  v_开始号码 票据领用记录.开始号码%Type;
  v_终止号码 票据领用记录.终止号码%Type;

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_领用id     := j_Json.Get_Number('recv_id');
  v_当前发票号 := j_Json.Get_String('inv_no');
  If Nvl(n_领用id, 0) = 0 Then
    Json_Out := zlJsonOut('未传入票据领用信息!');
    Return;
  End If;
  Begin
    Select Upper(前缀文本) As 前缀文本, Upper(开始号码), Upper(终止号码)
    Into v_前缀文本, v_开始号码, v_终止号码
    From 票据领用记录
    Where ID = n_领用id;
    n_Count := 1;
  Exception
    When Others Then
      n_Count := 0;
  End;

  If n_Count <> 0 Then
  
    For c_票据 In (
                 
                 Select Upper(号码) As 号码
                 From 票据使用明细
                 Where 号码 || '' >= v_当前发票号 And 领用id = n_领用id
                 Order By 号码) Loop
      If Substr(v_当前发票号, 1, Length(Nvl(v_前缀文本, ''))) <> Nvl(v_前缀文本, '') Then
        v_当前发票号 := '';
        Exit;
      End If;
      If Not (v_当前发票号 >= v_开始号码 And v_当前发票号 <= v_终止号码) Then
        v_当前发票号 := '';
        Exit;
      End If;
    
      Select Nvl(Max(1), 0) Into n_Count From 票据使用明细 Where 号码 = v_当前发票号 And 领用id = n_领用id;
      If n_Count = 0 Then
        Exit;
      End If;
      v_当前发票号 := Zl_Incstr(v_当前发票号);
    End Loop;
  Else
    v_当前发票号 := Null;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功","inv_no":"' || Nvl(v_当前发票号, '') || '"}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getnextinvoice;
/


Create Or Replace Procedure Zl_Exsesvr_Updatecardinvoice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:卡费按门诊医疗票据使用时进行门诊医疗票据的相关更新操作
  --入参：Json_In:格式
  --    input
  --      fun_oper  N 1 操作类型:1-发卡；2-退卡；3-重打；4-补打；5-换卡
  --      fee_nos C 1 费用单号s:多个用逗号
  --      recv_id N 1 领用id
  --      inv_no  C 1 当前发票号或开始使用发票号
  --      inv_usenums N 1 发票使用数量
  --      use_time  C 1 票据使用时间:yyyy-mm-dd hh24:mi:ss
  --      inv_user  C 1 发票使用人
  --出参: Json_Out,格式如下
  --   output
  --     code  C 1 应答码：0-失败；1-成功
  --     message C 1 应答消息： 成功时返回成功信息 失败时返回具体的错误信息
  --     inv_outnos  C 1 门诊医疗票据:使用的门诊医疗票据,多个用逗号返回
  ---------------------------------------------------------------------------

  j_Input PLJson;
  j_Json  PLJson;

  n_操作类型     Number(2);
  v_费用单号     Varchar2(20);
  n_领用id       Number(18);
  v_开始发票号   Varchar2(100);
  n_发票使用数量 Number(18);

  Cursor c_Fact(n_领用id 票据领用记录.Id%Type) Is
    Select * From 票据领用记录 Where ID = Nvl(n_领用id, 0);
  r_Factrow c_Fact%RowType;

  v_收回id     票据打印内容.Id%Type;
  v_票据号     票据使用明细.号码%Type;
  v_当前票据号 票据使用明细.号码%Type;
  n_打印id     票据打印内容.Id%Type;

  n_票据金额 票据使用明细.票据金额%Type;

  v_使用票据信息 Varchar2(4000);
  v_使用人       票据使用明细.使用人%Type;
  d_使用时间     票据使用明细.使用时间%Type;
  v_使用时间     Varchar2(30);
  v_Err_Msg      Varchar2(255);
  Err_Item Exception;

Begin
  --解析入参

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_操作类型     := j_Json.Get_Number('fun_oper');
  v_费用单号     := j_Json.Get_String('fee_no');
  n_领用id       := j_Json.Get_Number('recv_id');
  v_票据号       := j_Json.Get_String('inv_no');
  n_发票使用数量 := j_Json.Get_Number('inv_usenums');
  v_使用人       := j_Json.Get_String('inv_user');
  v_使用时间     := j_Json.Get_String('use_time');

  If v_使用时间 Is Null Then
    d_使用时间 := Sysdate;
  Else
    d_使用时间 := To_Date(v_使用时间, 'yyyy-mm-dd hh24:mi:ss');
  End If;

  If v_费用单号 Is Null Then
    Json_Out := zlJsonOut('未传入需要打印的单据信息!');
    Return;
  End If;

  --无票据号时,不用处理票据
  If v_票据号 Is Null Then
    Return;
  End If;

  --退卡
  If n_操作类型 = 2 Then
    Begin
      --从最后一次打印的内容中取
      Select ID
      Into v_收回id
      From (Select b.Id
             From 票据使用明细 A, 票据打印内容 B
             Where a.打印id = b.Id And a.性质 = 1 And a.原因 In (1, 3) And a.票种 = 1 And b.数据性质 = 5 And b.No = v_费用单号
             Order By a.使用时间 Desc)
      Where Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
  
    If v_收回id Is Not Null Then
      Insert Into 票据使用明细
        (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
        Select 票据使用明细_Id.Nextval, 1, v_票据号, 2, 2, 领用id, 打印id, d_使用时间, v_使用人
        From 票据使用明细
        Where 打印id = v_收回id And 票种 = 1 And 性质 = 1;
    End If;
    Return;
  End If;

  --重打收回原始票据
  If n_操作类型 = 3 Or n_操作类型 = 5 Then
    Begin
      --从最后一次打印的内容中取
      Select ID
      Into v_收回id
      From (Select b.Id
             From 票据使用明细 A, 票据打印内容 B
             Where a.打印id = b.Id And a.性质 = 1 And a.原因 In (1, 3) And a.票种 = 1 And b.数据性质 = 5 And b.No = v_费用单号
             Order By a.使用时间 Desc)
      Where Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
  
    If v_收回id Is Not Null Then
      Begin
        Insert Into 票据使用明细
          (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
          Select 票据使用明细_Id.Nextval, 票种, 号码, 2, Decode(n_操作类型, 5, 2, 4), 领用id, 打印id, d_使用时间, v_使用人, 票据金额
          
          From 票据使用明细
          Where 打印id = v_收回id And 票种 = 1 And 性质 = 1;
      Exception
        When Others Then
          Null;
      End;
    End If;
  End If;

  --票据打印金额
  Select Nvl(Sum(实收金额), 0) Into n_票据金额 From 住院费用记录 Where 记录性质 = 5 And NO = v_费用单号;

  Select 票据打印内容_Id.Nextval Into n_打印id From Dual;
  --生成单据的票据打印内容
  Insert Into 票据打印内容 (ID, 数据性质, NO) Values (n_打印id, 5, v_费用单号);

  --并发出票据
  If Nvl(n_领用id, 0) <> 0 Then
    Open c_Fact(n_领用id);
    Fetch c_Fact
      Into r_Factrow;
    If c_Fact%RowCount = 0 Then
      v_Err_Msg := '无效的票据领用批次，无法完成挂号票据分配操作。';
      Close c_Fact;
      Raise Err_Item;
    Elsif Nvl(r_Factrow.剩余数量, 0) < n_发票使用数量 Then
      v_Err_Msg := '当前批次的剩余数量不足' || n_发票使用数量 || '张，无法完成挂号票据分配操作。';
      Close c_Fact;
      Raise Err_Item;
    End If;
  End If;
  v_使用票据信息 := Null;
  For I In 1 .. n_发票使用数量 Loop
    --检查票据范围是否正确
    If Nvl(n_领用id, 0) <> 0 Then
      If Not (Upper(v_票据号) >= Upper(r_Factrow.开始号码) And Upper(v_票据号) <= Upper(r_Factrow.终止号码) And
          Length(v_票据号) = Length(r_Factrow.终止号码)) Then
        v_Err_Msg := '该单据需要打印多张票据,但票据号"' || v_票据号 || '"超出票据领用的号码范围！';
        Close c_Fact;
        Raise Err_Item;
      End If;
    End If;
  
    --发出票据
    Insert Into 票据使用明细
      (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
    Values
      (票据使用明细_Id.Nextval, 1, v_票据号, 1, Decode(n_操作类型, 3, 3, 1), n_领用id, n_打印id, d_使用时间, v_使用人, n_票据金额);
  
    v_使用票据信息 := Nvl(v_使用票据信息, '') || ',' || v_票据号;
    v_当前票据号   := v_票据号;
    --下一个票据号
    v_票据号 := Zl_Incstr(v_票据号);
  End Loop;

  If Not v_使用票据信息 Is Null Then
    v_使用票据信息 := Substr(v_使用票据信息, 2);
  
  End If;
  If Nvl(n_领用id, 0) <> 0 Then
    Update 票据领用记录
    Set 使用时间 = d_使用时间, 当前号码 = v_当前票据号, 剩余数量 = Nvl(剩余数量, 0) - n_发票使用数量
    Where ID = n_领用id;
    Close c_Fact;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功","inv_outnos":"' || v_使用票据信息 || '"}}';
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Updatecardinvoice;
/

Create Or Replace Procedure Zl_Exsesvr_Getbillopercontrols
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:获取单据操作控制数据
  --入参：Json_In:格式
  --  input       
  --    bill_type  N  1  单据类型:1-挂号单据,2-收费单,3-划价单,4-门诊记帐,5-住院记帐,6-预交款,7-结帐单据,8-就诊卡
  --    operator_id  N  1  人员ID

  --出参: Json_Out,格式如下
  --   output      
  --    code  C  1  应答码：0-失败；1-成功
  --    message  C  1  应答消息：失败时返回具体的错误信息
  --      is_exist  N  1  存在控制数据:1-存在;0-不存在
  --    time_limit  N  1  0(NULL)-不限制,n-n天内
  --    other_bill  N  1  是否允许对其它单据进行操作
  --    uplimit_money  N  1  金额上线

  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_单据类型 Number(2);
  n_人员id   Number(18);
  n_Count    Number(5);

  n_时间限制 单据操作控制.时间限制%Type;
  n_他人单据 单据操作控制.他人单据%Type;
  n_金额上限 单据操作控制.金额上限%Type;

Begin
  --解析入参

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_单据类型 := j_Json.Get_Number('bill_type');
  n_人员id   := j_Json.Get_Number('operator_id');

  Select Max(1), Max(Nvl(时间限制, 0)), Max(Nvl(他人单据, 0)), Max(Nvl(金额上限, 0))
  Into n_Count, n_时间限制, n_他人单据, n_金额上限
  From 单据操作控制
  Where 人员id = n_人员id And 单据 = n_单据类型;

  Json_Out := '{"output":{"code":1,"message":"' || '成功' || '"';
  Json_Out := Json_Out || ',"is_exist":' || Nvl(n_Count, 0);
  Json_Out := Json_Out || ',"time_limit":' || Nvl(n_时间限制, 0) || '';
  Json_Out := Json_Out || ',"other_bill":' || Nvl(n_他人单据, 0);
  Json_Out := Json_Out || ',"uplimit_money":' || Nvl(n_金额上限, 0);
  Json_Out := Json_Out || '}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getbillopercontrols;
/


Create Or Replace Procedure Zl_Exsesvr_Getfullno
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:自动补齐单据号
  --入参：Json_In:格式
  --    input
  --      item_num  N 1 项目序号
  --      input_no  C 1 输入的单据号
  --      dept_id   N   科室ID

  --出参: Json_Out,格式如下
  --  output
  --    code              N   1   应答码：0-失败；1-成功
  --    message           C   1   应答消息：失败时返回具体的错误信息
  --    full_no           C       补齐后的单据号
  ---------------------------------------------------------------------------

  j_Input PLJson;
  j_Json  PLJson;

  n_序号   号码控制表.项目序号%Type;
  v_No     号码控制表.最大号码%Type;
  n_科室id 部门表.Id%Type := Null;
  v_No_Out 号码控制表.最大号码%Type;

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_序号   := j_Json.Get_Number('item_num');
  v_No     := j_Json.Get_String('input_no');
  n_科室id := j_Json.Get_Number('dept_id');
  If Nvl(n_序号, 0) = 0 Then
    Json_Out := zlJsonOut('未传入序号，请检查！');
    Return;
  End If;

  If Nvl(v_No, '-') = '-' Then
    Json_Out := zlJsonOut('未传入NO，请检查！');
    Return;
  End If;
  v_No_Out := Fullno(n_序号, v_No, n_科室id);
  Json_Out := '{"output":{"code":1,"message":"成功","full_no":"' || Nvl(v_No_Out, '') || '"}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getfullno;
/

Create Or Replace Procedure Zl_Exsesvr_Getpatitotalmoney
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:根据病人ID,主页ID或医嘱id，获取应收、实收总额
  --入参：Json_In:格式
  --  input
  --    pati_source N 1 病人来源:0-所有;1-门诊;2-住院
  --    pati_id N 1 病人ID
  --    visit_id  N   就诊ID:住院时，传入主页id,门诊暂传NULL
  --    advice_ids  C   医嘱ids:多个用逗号分离
  --    today_fee N   是否当日费用:1-是的;0-不限制
  --    price_tag N   划价标志:0-不限制;1-不含划价单;2-仅统计划价单
  --出参: Json_Out,格式如下
  --  output
  --    code        C  1  应答码：0-失败；1-成功
  --    message     C  1  应答消息：
  --    fee_amrcvb  N  1  应收金额
  --    fee_ampaib  N  1  实收金额
  ------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_Output Varchar2(32767);

  n_病人id   Number(18);
  n_就诊id   Number(18);
  n_病人来源 Number(2);

  v_医嘱ids  Varchar2(32767);
  n_当日费用 Number(2);
  n_划价标志 Number(2);
  v_记录状态 Varchar2(10);
  n_实收金额 门诊费用记录.实收金额%Type;
  n_应收金额 门诊费用记录.应收金额 %Type;

  Err_Item Exception;

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_病人来源 := Nvl(j_Json.Get_Number('pati_source'), 0);
  n_病人id   := Nvl(j_Json.Get_Number('pati_id'), 0);
  n_就诊id   := j_Json.Get_Number('visit_id');
  v_医嘱ids  := j_Json.Get_String('advice_ids');
  n_当日费用 := Nvl(j_Json.Get_Number('today_fee'), 0);
  n_划价标志 := Nvl(j_Json.Get_Number('only_price'), 0);

  --0-所有;1-门诊;2-住院
  v_记录状态 := ',0,1,2,3,';
  If n_划价标志 = 1 Then
    --0-不限制;1-不含划价单;2-仅统计划价单
    v_记录状态 := ',1,2,3,';
  Elsif n_划价标志 = 2 Then
    v_记录状态 := ',0,';
  End If;

  If v_医嘱ids Is Not Null Then
    Select Sum(应收金额), Sum(实收金额)
    Into n_实收金额, n_应收金额
    From (With 医嘱数据 As (Select Distinct Column_Value As 医嘱id From Table(f_Num2List(v_医嘱ids)))
           Select /*+cardinality(B,10) */
            Sum(a.实收金额) As 实收金额, Sum(a.应收金额) As 应收金额
           
           From 住院费用记录 A, 医嘱数据 B
           Where a.医嘱序号 = b.医嘱id
           Union All
           Select /*+cardinality(B,10) */
            Sum(a.实收金额) As 实收金额, Sum(a.应收金额) As 应收金额
           
           From 门诊费用记录 A, 医嘱数据 B
           Where a.医嘱序号 = b.医嘱id);
  
  
    zlJsonPutValue(v_Output, 'code', 1, 1, 1);
    zlJsonPutValue(v_Output, 'message', '成功');
    zlJsonPutValue(v_Output, 'fee_ampaib', Nvl(n_实收金额, 0), 1);
    zlJsonPutValue(v_Output, 'fee_amrcvb', Nvl(n_应收金额, 0), 1, 2);
    Json_Out := '{"output":' || v_Output || '}';
    Return;
  
  End If;

  If Nvl(n_病人来源, 0) = 0 Or Nvl(n_病人来源, 0) = 1 Then
    --查询所有费用及门诊
    If Nvl(n_当日费用, 0) = 1 Then
      --查当日费用
      Select Sum(实收金额), Sum(应收金额)
      Into n_实收金额, n_应收金额
      From (Select Sum(a.实收金额) As 实收金额, Sum(a.应收金额) As 应收金额
             From 住院费用记录 A
             Where 病人id = n_病人id And (n_就诊id Is Null Or Nvl(主页id, 0) = Nvl(n_就诊id, 0)) And
                   Instr(v_记录状态, ',' || a.记录状态 || ',') > 0 And (n_病人来源 = 0 Or n_病人来源 = 1 And 门诊标志 <> 2) And
                   发生时间 >= Trunc(Sysdate) And 发生时间 <= Trunc(Sysdate) + 1 - 1 / 24 / 60 / 60 And Nvl(a.记帐费用, 0) = 1
             Union All
             Select Sum(a.实收金额) As 实收金额, Sum(a.应收金额) As 应收金额
             From 门诊费用记录 A
             Where 病人id = n_病人id And (n_就诊id Is Null Or Nvl(主页id, 0) = Nvl(n_就诊id, 0)) And
                   Instr(v_记录状态, ',' || a.记录状态 || ',') > 0 And 发生时间 >= Trunc(Sysdate) And
                   发生时间 <= Trunc(Sysdate) + 1 - 1 / 24 / 60 / 60 And Nvl(a.记帐费用, 0) = 1);
    
    Else
      Select Sum(实收金额), Sum(应收金额)
      Into n_实收金额, n_应收金额
      From (Select Sum(a.实收金额) As 实收金额, Sum(a.应收金额) As 应收金额
             From 住院费用记录 A
             Where 病人id = n_病人id And (n_就诊id Is Null Or Nvl(主页id, 0) = Nvl(n_就诊id, 0)) And
                   Instr(v_记录状态, ',' || a.记录状态 || ',') > 0 And (n_病人来源 = 0 Or n_病人来源 = 1 And 门诊标志 <> 2) And
                   Nvl(a.记帐费用, 0) = 1
             Union All
             Select Sum(a.实收金额) As 实收金额, Sum(a.应收金额) As 应收金额
             From 门诊费用记录 A
             Where 病人id = n_病人id And (n_就诊id Is Null Or Nvl(主页id, 0) = Nvl(n_就诊id, 0)) And
                   Instr(v_记录状态, ',' || a.记录状态 || ',') > 0 And Nvl(a.记帐费用, 0) = 1);
    End If;
  Else
    --查住院 
    If Nvl(n_当日费用, 0) = 1 Then
      --查当日费用
      Select Sum(实收金额), Sum(应收金额)
      Into n_实收金额, n_应收金额
      From 住院费用记录 A
      Where 病人id = n_病人id And Nvl(a.记帐费用, 0) = 1 And a.门诊标志 = 2 And (n_就诊id Is Null Or Nvl(主页id, 0) = Nvl(n_就诊id, 0)) And
            Instr(v_记录状态, ',' || a.记录状态 || ',') > 0 And 发生时间 >= Trunc(Sysdate) And
            发生时间 <= Trunc(Sysdate) + 1 - 1 / 24 / 60 / 60 And Nvl(a.记帐费用, 0) = 1 And a.门诊标志 = 2;
    Else
      Select Sum(实收金额), Sum(应收金额)
      Into n_实收金额, n_应收金额
      From 住院费用记录 A
      Where 病人id = n_病人id And Nvl(a.记帐费用, 0) = 1 And a.门诊标志 = 2 And (n_就诊id Is Null Or Nvl(主页id, 0) = Nvl(n_就诊id, 0)) And
            Instr(v_记录状态, ',' || a.记录状态 || ',') > 0;
    End If;
  End If;
  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '成功');
  zlJsonPutValue(v_Output, 'fee_ampaib', Nvl(n_实收金额, 0), 1);
  zlJsonPutValue(v_Output, 'fee_amrcvb', Nvl(n_应收金额, 0), 1, 2);
  Json_Out := '{"output":' || v_Output || '}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getpatitotalmoney;
/


Create Or Replace Procedure Zl_Exsesvr_Getbilltotalmoney
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:根据指定的单据号，获取应收、实民总额
  --入参：Json_In:格式
  --  input       
  --    fee_origin  N  1  费用业源:1-门诊;2-住院
  --    bill_type  N  1  单据类型:1-收费单;3-自动记帐单；2 -记帐记录；4-挂号记录 ;5-就诊卡
  --    fee_no  C  1  费用单号
  --    pati_id  N  1  病人id
  --    rec_status  C    记录状态:可以多个状态,比如:0,1

  --出参: Json_Out,格式如下
  --  output      
  --    code  C  1  应答码：0-失败；1-成功
  --    message  C  1  应答消息：
  --    fee_amrcvb  N  1  应收金客
  --    fee_ampaib  N  1  实收金额
  ---------------------------------------------------------------------------
  j_Input     PLJson;
  j_Json      PLJson;
  n_费用来源  Number(2);
  n_单据类型  Number(2);
  v_单据号    Varchar2(100);
  n_病人id    Number(18);
  v_记录状态s Varchar2(100);

  n_实收金额 门诊费用记录.实收金额%Type;
  n_应收金额 门诊费用记录.应收金额 %Type;

  --组装失败时返回的数据
Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_费用来源  := j_Json.Get_Number('fee_origin');
  n_单据类型  := j_Json.Get_Number('bill_type');
  v_单据号    := j_Json.Get_String('fee_no');
  n_病人id    := j_Json.Get_Number('pati_id');
  v_记录状态s := j_Json.Get_String('rec_status');

  If v_记录状态s Is Null Then
    v_记录状态s := ',0,1,';
  Else
    v_记录状态s := ',' || v_记录状态s || ',';
  End If;

  If Nvl(n_费用来源, 0) <= 1 Then
    Select Sum(实收金额), Sum(应收金额)
    Into n_实收金额, n_应收金额
    From 门诊费用记录
    Where NO = v_单据号 And 记录性质 = n_单据类型 And Instr(v_记录状态s, ',' || 记录状态 || ',') > 0 And
          (Nvl(n_病人id, 0) = 0 Or 病人id = Nvl(n_病人id, 0));
  Else
    Select Sum(实收金额), Sum(应收金额)
    Into n_实收金额, n_应收金额
    From 住院费用记录
    Where NO = v_单据号 And 记录性质 = n_单据类型 And Instr(v_记录状态s, ',' || 记录状态 || ',') > 0 And
          (Nvl(n_病人id, 0) = 0 Or 病人id = Nvl(n_病人id, 0));
  End If;
  Json_Out := '{"output":{"code":1,"message":"' || '成功' || '"';
  Json_Out := Json_Out || ',"fee_ampaib":' || zlJsonStr(Nvl(n_实收金额, 0), 1);
  Json_Out := Json_Out || ',"fee_amrcvb":' || zlJsonStr(Nvl(n_应收金额, 0), 1);
  Json_Out := Json_Out || '}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getbilltotalmoney;
/

Create Or Replace Procedure Zl_Exsesvr_Getbillinfobyno
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:根据单据号获取单据的单据信息，如：登记人，登记时间
  -- input     
  --   fee_origin  N 1 费用来源:1-门诊;2-住院
  --   bill_type N 1 单据类型:1-收费单;3-自动记帐单；2 -记帐记录；4-挂号记录 ;5-就诊卡;-1-结帐单;-2-预交单;-3-补充结算
  --   bill_no C 1 单据号
  --出参: Json_Out,格式如下
  --  output     
  --   code  C 1 应答码：0-失败；1-成功
  --   message C 1 应答消息： 
  --   operator_name C 1 操作员姓名
  --   create_time C 1 登记时间:yyyy-mm-dd hh24:mi:ss
  --   pati_id N 1 病人id
  ---------------------------------------------------------------------------
  j_Input  PLJson;
  j_Json   PLJson;
  v_Output Varchar2(32767);

  n_费用来源 Number(2);
  n_单据类型 Number(2);
  v_单据号   Varchar2(100);
  n_病人id   Number(18);

  v_操作员姓名 门诊费用记录.操作员姓名%Type;
  v_登记时间   Varchar2(30);

Begin
  --解析入参

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_费用来源 := j_Json.Get_Number('fee_origin');
  n_单据类型 := Nvl(j_Json.Get_Number('bill_type'), 0);
  v_单据号   := j_Json.Get_String('bill_no');

  If n_单据类型 = -1 Then
    --1.结帐
    Select Max(操作员姓名), To_Char(Max(收费时间), 'yyyy-mm-dd hh24mi:ss'), Max(病人id)
    Into v_操作员姓名, v_登记时间, n_病人id
    From 病人结帐记录
    Where NO = v_单据号 And 记录状态 In (1, 3);
  Elsif n_单据类型 = -2 Then
    --2.预交
    Select Max(操作员姓名), To_Char(Max(收款时间), 'yyyy-mm-dd hh24:mi:ss'), Max(病人id)
    Into v_操作员姓名, v_登记时间, n_病人id
    From 病人预交记录
    Where NO = v_单据号 And 记录状态 In (1, 3);
  Elsif n_单据类型 = -3 Then
    --3.补充结算
    Select Max(操作员姓名), To_Char(Max(登记时间), 'yyyy-mm-dd hh24:mi:ss'), Max(病人id)
    Into v_操作员姓名, v_登记时间, n_病人id
    From 费用补充记录
    Where NO = v_单据号 And 记录状态 In (1, 3);
  Else
    --4.费用相关
    Begin
      If Nvl(n_费用来源, 0) <= 1 Then
        Select Nvl(操作员姓名, 划价人), To_Char(登记时间, 'yyyy-mm-dd hh24:mi:ss'), 病人id
        Into v_操作员姓名, v_登记时间, n_病人id
        From 门诊费用记录
        Where NO = v_单据号 And 记录性质 = n_单据类型 And 记录状态 In (0, 1, 3) And Rownum < 2;
      Else
        Select Nvl(操作员姓名, 划价人), To_Char(登记时间, 'yyyy-mm-dd hh24:mi:ss'), 病人id
        Into v_操作员姓名, v_登记时间, n_病人id
        From 住院费用记录
        Where NO = v_单据号 And 记录性质 = n_单据类型 And 记录状态 In (0, 1, 3) And Rownum < 2;
      End If;
    
    Exception
      When Others Then
        v_操作员姓名 := Null;
        v_登记时间   := '';
        n_病人id     := Null;
    End;
  End If;

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '成功');
  zlJsonPutValue(v_Output, 'operator_name', Nvl(v_操作员姓名, ''));
  zlJsonPutValue(v_Output, 'create_time', Nvl(v_登记时间, ''));
  zlJsonPutValue(v_Output, 'pati_id', Nvl(n_病人id, 0), 1, 2);

  v_Output := '{"output":' || v_Output || '}';
  Json_Out := v_Output;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getbillinfobyno;
/


Create Or Replace Procedure Zl_Exsesvr_Getpatiinvoiceclass
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:获取票据使用类别
  --入参：Json_In:格式
  -- input
  --   pati_id              N 1 病人id
  --   pati_pageid          N 1 主页id
  --   insure_type          N 1 险类

  --出参: Json_Out,格式如下
  --  output
  --    code                C 1 应答码：0-失败；1-成功
  --    message             C 1 应答消息： 失败时返回具体的错误信息
  --    use_type            C  1  票据使用类别
  ---------------------------------------------------------------------------
  j_Input    PLJson;
  j_Json     PLJson;
  n_病人id   Number(18);
  n_主页id   Number(18);
  n_险类     Number(18);
  v_使用类别 Varchar2(4000);

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_病人id   := j_Json.Get_Number('pati_id');
  n_主页id   := j_Json.Get_Number('pati_pageid');
  n_险类     := j_Json.Get_Number('insure_type');
  v_使用类别 := Zl_Billclass(n_病人id, n_主页id, n_险类);

  Json_Out := '{"output":{"code":1,"message":"' || '成功' || '","use_type":"' || Nvl(v_使用类别, '') || '"}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getpatiinvoiceclass;
/

Create Or Replace Procedure Zl_Exsesvr_Invoiceclassused
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:判断指定票种是否启用了分使用类别打印
  --入参：Json_In:格式
  --  input
  --    inv_type            N  1  票种:1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡

  --出参: Json_Out,格式如下
  --  output
  --    code                C  1  应答码：0-失败；1-成功
  --    message             C  1  应答消息： 失败时返回具体的错误信息
  --    is_start            N  1  是否启用:1-启用了的，0-未启用

  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_票种 Number(5);
  n_启用 Number(2);

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_票种 := j_Json.Get_Number('inv_type');

  Select Nvl(Max(1), 0)
  Into n_启用
  From 票据领用记录
  Where 票种 = n_票种 And Nvl(使用类别, 'LXH') <> 'LXH' And Rownum < 2;

  Json_Out := '{"output":{"code":1,"message":"' || '成功' || '","is_start":' || Nvl(n_启用, 0) || '}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Invoiceclassused;
/


Create Or Replace Procedure Zl_Exsesvr_Patimove_Check
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  -------------------------------------------------------------------------------------------------
  --功能：病人换床前费用相关检查
  --入参：Json_In格式
  --  input
  --    pati_id            N 1 病人ID
  --    pati_pageid        N 1 主页ID
  --    operator_time      C 0 操作时间
  --出参: Json_Out,格式如下
  --  output
  --    code               N 1 应答吗：0-失败；1-成功
  --    message            C 1 应答消息：失败时返回具体的错误信息
  -------------------------------------------------------------------------------------------------
  j_Input         PLJson;
  j_Json          PLJson;
  v_Tmp           Varchar2(20);
  n_Pati_Id       Number(18);
  n_Pati_Pageid   Number(18);
  d_Operator_Time Date;

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_Pati_Id       := j_Json.Get_Number('pati_id');
  n_Pati_Pageid   := j_Json.Get_Number('pati_pageid');
  v_Tmp           := j_Json.Get_String('operator_time');
  d_Operator_Time := To_Date(v_Tmp, 'YYYY-MM-DD HH24:MI:SS');

  For r_Fee In (Select NO
                From 住院费用记录
                Where 病人id = n_Pati_Id And 主页id = n_Pati_Pageid And Mod(记录性质, 10) = 3 And 登记时间 >= d_Operator_Time And
                      收费类别 = 'J'
                Group By NO, 序号, Mod(记录性质, 10)
                Having Sum(结帐金额) <> 0) Loop
    Json_Out := zlJsonOut('变动时间之后已有已结帐的自动记帐费用,不能进行换床操作！');
    Return;
  End Loop;

  Json_Out := zlJsonOut('成功', 1);
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Patimove_Check;
/

Create Or Replace Procedure Zl_Exsesvr_Billinhistory
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:查询指定的费用单据是否存在后备表空间中
  --入参：Json_In:格式
  --    input
  --      bill_no           C 1 单据号
  --      bill_type         C 1 单据类型:1-收费单,2-预交单,3-结帐单,4-挂号单,5-就诊卡单据,6-记帐单据;7-自动记帐单
  --      outpati_flag      N 1 门诊标志：1-门诊，2-住院
  --出参: Json_Out,格式如下
  --  output
  --    code              N   1   应答码：0-失败；1-成功
  --    message           C   1   应答消息：失败时返回具体的错误信息
  --    exits_history     C   1   存在历史后备表:1-存在;1-不存在
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_单据类型 Number(1);
  n_门诊标志 Number(1);
  v_单据号   Varchar2(100);

  v_Output  Varchar2(32767);
  n_Nomoved Number(2);
Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_单据号   := j_Json.Get_String('bill_no');
  n_单据类型 := j_Json.Get_Number('bill_type');
  n_门诊标志 := j_Json.Get_Number('outpati_flag');

  If Nvl(v_单据号, '-') = '-' Then
    Json_Out := zlJsonOut('未传入单据号，请检查！');
    Return;
  End If;

  If Nvl(n_单据类型, 0) = 0 Then
    Json_Out := zlJsonOut('未传入单据类型，请检查！');
    Return;
  End If;

  If Nvl(n_门诊标志, 0) = 0 And Nvl(n_单据类型, 0) = 6 Then
    Json_Out := zlJsonOut('未传入门诊标志，请检查！');
    Return;
  End If;

  n_Nomoved := Zl_Fun_Checkinhistory(n_单据类型, v_单据号, n_门诊标志);

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '成功');
  zlJsonPutValue(v_Output, 'exits_history', n_Nomoved, 1, 2);
  Json_Out := '{"output":' || v_Output || '}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Billinhistory;
/


Create Or Replace Procedure Zl_Exsesvr_Billisprintinvoice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:判断指定的单据是否打印了发票
  --入参：Json_In:格式
  -- input     
  --    bill_no  C 1 单据号
  --    bill_type N 1 单据类型:1-收费,2-预交(包含押金),3-结帐,4-挂号,5-就诊卡
  --    inv_type  N 1 票种:1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡

  --出参: Json_Out,格式如下
  --  output      
  --    code  C 1 应答码：0-失败；1-成功
  --    message C 1 应答消息： 失败时返回具体的错误信息
  --    printed N 1 是否打印:1-已打印;0-未打印
  --
  ---------------------------------------------------------------------------
  j_Input    PLJson;
  j_Json     PLJson;
  v_单据号   Varchar2(100);
  n_单据类型 Number(2);
  n_票种     Number(2);
  n_是否打印 Number(2);

  --组装失败时返回的数据
  v_Output Varchar2(32767);

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_单据号   := j_Json.Get_String('bill_no');
  n_单据类型 := j_Json.Get_Number('bill_type');
  n_票种     := j_Json.Get_Number('inv_type');

  Begin
    Select 1
    Into n_是否打印
    From 票据使用明细 A, 票据打印内容 B
    Where a.打印id = b.Id And a.票种 = n_票种 And b.No = v_单据号 And b.数据性质 = n_单据类型 And Rownum < 2;
  
  Exception
    When Others Then
      n_是否打印 := 0;
  End;

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '成功');
  zlJsonPutValue(v_Output, 'printed', n_是否打印, 1, 2);
  Json_Out := '{"output":' || v_Output || '}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Billisprintinvoice;
/


Create Or Replace Procedure Zl_Exsesvr_Getcardfeeinfobyno
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:根据卡费单据号，获取卡费及病历费及预交款的相关信息
  --入参：Json_In:格式
  --  input
  --    fee_no  C 1 单据号:费用单据号
  --    query_type  N   查询类型：0-读取正常单据:1-读取作废单据;2-剩余费用单据
  --    query_deposit N 1 是否包含预交:1-包含预交信息，0-不包含预交信息
  --出参: Json_Out,格式如下
  -- output
  --    code  C 1 应答码：0-失败；1-成功
  --    message C 1 应答消息：
  --    fee_list      [数组]每个费用ID信息
  --      fee_id  N 1 费用id
  --      fee_num N 1 序号
  --      pati_id N 1 病人id
  --      pati_name C 1 姓名
  --      pati_sex  C 1 性别
  --      pati_age  C 1 年龄
  --      fee_category  C 1 费别
  --      item_id N 1 收费项目id
  --      income_item_id  N 1 收入项目id
  --      quantity  N 1 数次
  --      fee_amrcvb  N 1 应收金额
  --      fee_ampaid  N 1 实收金额
  --      placer  C 1 开单人
  --      operator_code C 1 操作员编号
  --      operator_name C 1 操作员姓名
  --      create_time D 1 登记时间
  --      happen_time D 1 发生时间
  --      rec_status  N 1 记录状态
  --      mrbkfee_sign N 1 是否病历费:1-是病历费;0-不是病历费
  --      invoice_no  N 1 发票号
  --      kpbooks_sign N 1 记帐标志:1-是记帐;0-现收
  --      fee_status N 1 费用状态:1-异常状态;0-正常费用
  --      cardtype_id N 1 卡类别ID
  --      card_no C 1 卡号
  --      sendcard_reg  N 1 是否挂号发卡:1-是挂号同时发卡;0-非挂号同时发卡
  --    pricebill_info  C    卡费生成划价费用信息
  --      fee_no  C    划价单号
  --      cardfee_amrcvb  N 1 卡费应收金额
  --      cardfee_ampaid  N 1 卡费实收金额
  --      mrbkfee_amrcvb N 1 病历费应收
  --      mrbkfee_ampaid N 1 病历费实收
  --      charged_statu N 1 收费状态:0-未收费;1-已收费;2-已全退
  --    balance_list[]  C   结算信息
  --      blnc_mode C 1 结算方式名称
  --      balance_id  N 1 结帐ID： 查询作废的单据时为冲销ID
  --      blnc_money  N 1 结帐总额
  --      pay_cardno  N 1 支付卡号
  --      pay_swapno  C 1 交易流水号
  --      pay_swapmemo  C 1 交易说明
  --      relation_id N 1 关联交易id
  --      cardtype_id N 1 卡类别id
  --      consume_card  N 1 是否消费卡:1-是;0-不是
  --      blnc_nature N 1 结算性质:1-现金结算方式,2-其他非医保结算 , 8-结算卡结算 ,9-误差费
  --      blnc_statu  N 1 结算状态:1-未调用接口;2-接口调用成功,但还未收费完成,0-正常结算
  --      consume_card_id N 1 消费卡id
  --      blnc_no C 1 结算号码
  --      blnc_memo C 1 摘要
  --      original_id N 1 原结帐ID:冲销时返回
  --      original_money N,1 原始金额,求剩余款数时返回

  --   deposit_info  C   预交信息:query_deposit=1时有效，缺省包含
  --      deposit_id  N 1 预交id
  --      deposit_no  C 1 预交单据号
  --      deposit_money N 1 预交金额
  --      blnc_mode C 1 结算方式
  --      pay_cardno  N 1 支付卡号
  --      pay_swapno  C 1 交易流水号
  --      pay_swapmemo  C 1 交易说明
  --      relation_id N 1 关联交易id
  --      cardtype_id N 1 卡类别id
  --      consume_card  N 1 是否消费卡:1-是;0-不是
  --      blnc_mode C 1 结算方式
  --      blnc_nature N 1 结算性质
  --      blnc_statu  N 1 结算状态:1-未调用接口;2-接口调用成功,但还未收费完成,0-正常结算
  --      consume_card_id N 1 消费卡id
  --      blnc_no C 1 结算号码
  --      blnc_memo C 1 摘要
  ---------------------------------------------------------------------------
  j_Input        PLJson;
  j_Json         PLJson;
  v_单据号       Varchar2(100);
  n_查询类型     Number;
  v_记录状态     Varchar2(10);
  v_发票号       票据使用明细.号码%Type;
  v_划价单       Varchar2(100);
  n_实收金额     门诊费用记录.实收金额%Type;
  n_剩余金额     门诊费用记录.实收金额%Type;
  n_查询预交     Number(2);
  n_卡费应收     门诊费用记录.实收金额%Type;
  n_卡费实收     门诊费用记录.实收金额%Type;
  n_病历费应收   门诊费用记录.实收金额%Type;
  n_病历费实收   门诊费用记录.实收金额%Type;
  n_挂号同步发卡 Number(2);
  n_结帐id       Number(18);
  n_原结帐id     Number(18);
  n_Find         Number(2);
  n_Nomoved      Number(2);
  Cursor c_费用信息 Is
    Select a.Id, a.No, a.记录状态, a.序号, a.费别, a.姓名, a.性别, a.年龄, a.病人id, a.收费细目id, a.收入项目id, a.实际票号, a.数次,
           Decode(n_查询类型, 1, -1, 1) * a.应收金额 As 应收金额, Decode(n_查询类型, 1, -1, 1) * a.实收金额 As 实收金额, a.记帐费用,
           Nvl(a.加班标志, 0) As 变动类别, a.操作员姓名, a.操作员编号, To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间,
           To_Char(a.发生时间, 'yyyy-mm-dd hh24:mi:ss') As 发生时间, Decode(Nvl(a.附加标志, 0), 8, 1, 0) As 病历费, a.结论 As 卡类别id,
           a.结帐id, a.开单人, a.费用状态, a.摘要, a.实际票号 As 卡号
    From 住院费用记录 A
    Where a.记录性质 = 5 And NO = '-' And Rownum < 1;

  r_费用 c_费用信息%RowType;

  Cursor c_结算信息 Is
    Select a.No, a.结算方式, Nvl(a.冲预交, 0) As 冲预交, a.关联交易id, a.卡类别id, a.卡号, a.结算卡序号, a.交易流水号, a.交易说明, b.性质, a.校对标志,
           Max(c.消费卡id) As 消费卡id, Max(a.结算号码) As 结算号码, Max(a.摘要) As 摘要
    From 病人预交记录 A, 结算方式 B, 病人卡结算记录 C
    Where a.结算方式 = 名称(+) And a.结帐id = 1 And a.记录性质 = 1 And a.Id = c.结算id(+);

  r_结算信息 c_结算信息%RowType;

  Cursor c_预交信息 Is
    Select a.No, a.Id As 预交id, Nvl(Sum(a.金额), 0) As 金额, Max(a.结算方式) As 结算方式, Nvl(Sum(a.冲预交), 0) As 冲预交,
           Max(a.关联交易id) As 关联交易id, Max(a.卡类别id) As 卡类别id, Max(Decode(a.记录性质, 1, a.卡号, '')) As 卡号, Max(a.结算卡序号) As 结算卡序号,
           Max(Decode(a.记录性质, 1, a.交易流水号, '')) As 交易流水号, Max(Decode(a.记录性质, 1, a.交易说明, '')) 交易说明, Max(b.性质) As 性质,
           Max(Decode(a.记录性质, 1, a.校对标志, 0)) As 校对标志, Max(c.消费卡id) As 消费卡id, Max(a.结算号码) As 结算号码, Max(a.摘要) As 摘要
    From 病人预交记录 A, 结算方式 B, 病人卡结算记录 C
    Where a.结算方式 = 名称(+) And a.关联交易id = 0 And Mod(a.记录性质, 10) = 1 And a.Id = c.结算id(+)
    Group By NO;

  r_预交信息 c_预交信息%RowType;

  Type t_信息 Is Ref Cursor;

  c_信息 t_信息; --动态游标变量

  n_总额 住院费用记录.实收金额%Type;
  --组装失败时返回的数据
  v_Priebill   Varchar2(32767);
  v_Balanceinf Varchar2(32767);
  v_Deposit    Varchar2(32767);
  v_Output     Varchar2(32767);

  c_Output Clob;

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_单据号   := j_Json.Get_String('fee_no');
  n_查询类型 := Nvl(j_Json.Get_Number('query_type'), 0);

  n_查询预交 := Nvl(j_Json.Get_Number('query_deposit'), 1);

  v_记录状态 := ',1,3,';
  If Nvl(n_查询类型, 0) = 1 Then
    v_记录状态 := ',2,';
  Elsif n_查询类型 = 2 Then
    --只查询剩余费用单据
    v_记录状态 := ',1,2,3,';
  End If;

  Select Nvl(Max(1), 0) Into n_Nomoved From H住院费用记录 A Where NO = v_单据号 And 记录性质 = 5 And Rownum <= 1;

  If Nvl(n_Nomoved, 0) = 0 Then
    Select Max(m.号码)
    Into v_发票号
    From 票据打印内容 N, 票据使用明细 M
    Where n.数据性质 = 5 And n.Id = m.打印id And m.性质 = 1 And m.票种 = 1 And
          m.使用时间 = (Select Max(M2.使用时间)
                    From 票据打印内容 N2, 票据使用明细 M2
                    Where M2.打印id = N2.Id And n.数据性质 = 5 And M2.票种 = 1 And N2.No = v_单据号) And n.No = v_单据号;
  Else
    Select Max(m.号码)
    Into v_发票号
    From H票据打印内容 N, H票据使用明细 M
    Where n.数据性质 = 5 And n.Id = m.打印id And m.性质 = 1 And m.票种 = 1 And
          m.使用时间 = (Select Max(M2.使用时间)
                    From H票据打印内容 N2, H票据使用明细 M2
                    Where M2.打印id = N2.Id And n.数据性质 = 5 And M2.票种 = 1 And N2.No = v_单据号) And n.No = v_单据号;
  End If;

  --先读取费用
  If Nvl(n_Nomoved, 0) = 0 Then
  
    Select Max(1)
    Into n_挂号同步发卡
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 = 1 And
          (病人id, 登记时间) In (Select 病人id, 登记时间
                           From 住院费用记录
                           Where 记录性质 = 5 And NO = v_单据号 And 记录状态 In (0, 1, 3) And Rownum < 2);
  
    If Nvl(n_查询类型, 0) = 1 Then
      Select Max(结帐id) Into n_原结帐id From 住院费用记录 Where 记录性质 = 5 And 记录状态 In (1, 3) And NO = v_单据号;
    End If;
  
    If Nvl(n_查询类型, 0) = 1 Then
      --作废
      Open c_信息 For
        Select a.Id, a.No, a.记录状态, a.序号, a.费别, a.姓名, a.性别, a.年龄, a.病人id, a.收费细目id, a.收入项目id, a.实际票号, a.数次,
               Decode(n_查询类型, 1, -1, 1) * a.应收金额 As 应收金额, Decode(n_查询类型, 1, -1, 1) * a.实收金额 As 实收金额, a.记帐费用,
               Nvl(a.加班标志, 0) As 变动类别, a.操作员姓名, a.操作员编号, To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间,
               To_Char(a.发生时间, 'yyyy-mm-dd hh24:mi:ss') As 发生时间, Decode(Nvl(a.附加标志, 0), 8, 1, 0) As 病历费, a.结论 As 卡类别id,
               a.结帐id, a.开单人, a.费用状态, a.摘要, a.实际票号 As 卡号
        From 住院费用记录 A
        Where 结帐id In (Select Max(结帐id) As 结帐id
                       From 住院费用记录
                       Where 记录性质 = 5 And 记录状态 = 2 And NO = v_单据号 And Nvl(附加标志, 0) <> 8)
        Order By NO, 序号;
    Elsif n_查询类型 = 2 Then
      --剩余数量
      Open c_信息 For
        Select Max(Decode(a.记录状态, 2, 0, a.Id)) As ID, a.No, Null As 记录状态, a.序号, Max(a.费别) As 费别, Max(a.姓名) As 姓名,
               Max(a.性别) As 性别, Max(a.年龄) As 年龄, Max(a.病人id) As 病人id, a.收费细目id As 收费细目id, a.收入项目id, Max(a.实际票号) As 实际票号,
               Sum(a.数次) As 数次, Sum(a.应收金额) As 应收金额, Sum(a.实收金额) As 实收金额, Max(a.记帐费用) As 记帐费用,
               Max(Nvl(a.加班标志, 0)) As 变动类别, Max(Decode(a.记录状态, 2, Null, a.操作员姓名)) As 操作员姓名,
               Max(Decode(a.记录状态, 2, Null, a.操作员编号)) As 操作员编号,
               To_Char(Max(Decode(a.记录状态, 2, To_Date(Null), a.登记时间)), 'yyyy-mm-dd hh24:mi:ss') As 登记时间,
               To_Char(Max(Decode(a.记录状态, 2, To_Date(Null), a.发生时间)), 'yyyy-mm-dd hh24:mi:ss') As 发生时间,
               Max(Decode(Nvl(a.附加标志, 0), 8, 1, 0)) As 病历费, Max(a.结论) As 卡类别id,
               Max(Decode(a.记录状态, 2, Null, a.结帐id)) As 结帐id, Max(Decode(a.记录状态, 2, Null, a.开单人)) As 开单人,
               Max(a.费用状态) As 费用状态, Max(a.摘要) As 摘要, Max(a.实际票号) As 卡号
        From 住院费用记录 A
        Where a.记录性质 = 5 And NO = v_单据号
        Group By a.No, a.序号, a.收费细目id, a.收入项目id
        Having Sum(a.数次) <> 0
        Order By a.No, a.序号;
    Else
    
      Open c_信息 For
        Select a.Id, a.No, a.记录状态, a.序号, a.费别, a.姓名, a.性别, a.年龄, a.病人id, a.收费细目id, a.收入项目id, a.实际票号, a.数次,
               Decode(n_查询类型, 1, -1, 1) * a.应收金额 As 应收金额, Decode(n_查询类型, 1, -1, 1) * a.实收金额 As 实收金额, a.记帐费用,
               Nvl(a.加班标志, 0) As 变动类别, a.操作员姓名, a.操作员编号, To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间,
               To_Char(a.发生时间, 'yyyy-mm-dd hh24:mi:ss') As 发生时间, Decode(Nvl(a.附加标志, 0), 8, 1, 0) As 病历费, a.结论 As 卡类别id,
               a.结帐id, a.开单人, a.费用状态, a.摘要, a.实际票号 As 卡号
        From 住院费用记录 A
        Where a.记录性质 = 5 And Instr(v_记录状态, ',' || a.记录状态 || ',') > 0 And NO = v_单据号
        Order By NO, 序号;
    End If;
  Else
  
    Select Max(1)
    Into n_挂号同步发卡
    From H门诊费用记录
    Where 记录性质 = 4 And 记录状态 = 1 And
          (病人id, 登记时间) In (Select 病人id, 登记时间
                           From H住院费用记录
                           Where 记录性质 = 5 And NO = v_单据号 And Rownum < 2 And 记录状态 In (0, 1, 3));
  
    If Nvl(n_查询类型, 0) = 1 Then
      Select Max(结帐id)
      Into n_原结帐id
      From H住院费用记录
      Where 记录性质 = 5 And 记录状态 In (1, 3) And NO = v_单据号;
    
    End If;
  
    If Nvl(n_查询类型, 0) = 1 Then
      --作废
      Open c_信息 For
        Select a.Id, a.No, a.记录状态, a.序号, a.费别, a.姓名, a.性别, a.年龄, a.病人id, a.收费细目id, a.收入项目id, a.实际票号, a.数次,
               Decode(n_查询类型, 1, -1, 1) * a.应收金额 As 应收金额, Decode(n_查询类型, 1, -1, 1) * a.实收金额 As 实收金额, a.记帐费用,
               Nvl(a.加班标志, 0) As 变动类别, a.操作员姓名, a.操作员编号, To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间,
               To_Char(a.发生时间, 'yyyy-mm-dd hh24:mi:ss') As 发生时间, Decode(Nvl(a.附加标志, 0), 8, 1, 0) As 病历费, a.结论 As 卡类别id,
               a.结帐id, a.开单人, a.费用状态, a.摘要, a.实际票号 As 卡号
        From H住院费用记录 A
        Where 结帐id In (Select Max(结帐id) As 结帐id
                       From H住院费用记录
                       Where 记录性质 = 5 And 记录状态 = 2 And NO = v_单据号 And Nvl(附加标志, 0) <> 8)
        Order By NO, 序号;
    Elsif n_查询类型 = 2 Then
      --剩余数量
      Open c_信息 For
        Select Max(Decode(a.记录状态, 2, 0, a.Id)) As ID, a.No, Null As 记录状态, a.序号, Max(a.费别) As 费别, Max(a.姓名) As 姓名,
               Max(a.性别) As 性别, Max(a.年龄) As 年龄, Max(a.病人id) As 病人id, a.收费细目id As 收费细目id, a.收入项目id, Max(a.实际票号) As 实际票号,
               Sum(a.数次) As 数次, Sum(a.应收金额) As 应收金额, Sum(a.实收金额) As 实收金额, Max(a.记帐费用) As 记帐费用,
               Max(Nvl(a.加班标志, 0)) As 变动类别, Max(Decode(a.记录状态, 2, Null, a.操作员姓名)) As 操作员姓名,
               Max(Decode(a.记录状态, 2, Null, a.操作员编号)) As 操作员编号,
               To_Char(Max(Decode(a.记录状态, 2, To_Date(Null), a.登记时间)), 'yyyy-mm-dd hh24:mi:ss') As 登记时间,
               To_Char(Max(Decode(a.记录状态, 2, To_Date(Null), a.发生时间)), 'yyyy-mm-dd hh24:mi:ss') As 发生时间,
               Max(Decode(Nvl(a.附加标志, 0), 8, 1, 0)) As 病历费, Max(a.结论) As 卡类别id,
               Max(Decode(a.记录状态, 2, Null, a.结帐id)) As 结帐id, Max(Decode(a.记录状态, 2, Null, a.开单人)) As 开单人,
               Max(a.费用状态) As 费用状态, Max(a.摘要) As 摘要, Max(a.实际票号) As 卡号
        From H住院费用记录 A
        Where a.记录性质 = 5 And NO = v_单据号
        Group By a.No, a.序号, a.收费细目id, a.收入项目id
        Having Sum(a.数次) <> 0
        Order By a.No, a.序号;
    Else
      Open c_信息 For
        Select a.Id, a.No, a.记录状态, a.序号, a.费别, a.姓名, a.性别, a.年龄, a.病人id, a.收费细目id, a.收入项目id, a.实际票号, a.数次,
               Decode(n_查询类型, 1, -1, 1) * a.应收金额 As 应收金额, Decode(n_查询类型, 1, -1, 1) * a.实收金额 As 实收金额, a.记帐费用,
               Nvl(a.加班标志, 0) As 变动类别, a.操作员姓名, a.操作员编号, To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间,
               To_Char(a.发生时间, 'yyyy-mm-dd hh24:mi:ss') As 发生时间, Decode(Nvl(a.附加标志, 0), 8, 1, 0) As 病历费, a.结论 As 卡类别id,
               a.结帐id, a.开单人, a.费用状态, a.摘要, a.实际票号 As 卡号
        From H住院费用记录 A
        Where a.记录性质 = 5 And Instr(v_记录状态, ',' || a.记录状态 || ',') > 0 And NO = v_单据号
        Order By NO, 序号;
    End If;
  End If;

  v_划价单 := Null;
  Loop
    Fetch c_信息
      Into r_费用;
    Exit When c_信息%NotFound;
  
    If Length(Nvl(v_Output, ' ')) > 30000 Then
      If c_Output Is Not Null Then
        c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
      Else
        c_Output := To_Clob(v_Output);
      End If;
    
      v_Output := Null;
    End If;
    --1.取基本信息
    zlJsonPutValue(v_Output, 'fee_id', r_费用.Id, 1, 1);
    zlJsonPutValue(v_Output, 'fee_num', r_费用.序号, 1);
  
    zlJsonPutValue(v_Output, 'pati_id', r_费用.病人id, 1);
  
    zlJsonPutValue(v_Output, 'pati_name', r_费用.姓名);
    zlJsonPutValue(v_Output, 'pati_sex', Nvl(r_费用.性别, ''));
  
    zlJsonPutValue(v_Output, 'pati_age', Nvl(r_费用.年龄, ''));
  
    zlJsonPutValue(v_Output, 'fee_category', Nvl(r_费用.费别, ''));
  
    zlJsonPutValue(v_Output, 'item_id', r_费用.收费细目id, 1);
    zlJsonPutValue(v_Output, 'income_item_id', Nvl(r_费用.收入项目id, 0), 1);
  
    zlJsonPutValue(v_Output, 'quantity', Nvl(r_费用.数次, 0), 1);
  
    zlJsonPutValue(v_Output, 'fee_amrcvb', Nvl(r_费用.应收金额, 0), 1);
    zlJsonPutValue(v_Output, 'fee_ampaid', Nvl(r_费用.实收金额, 0), 1);
  
    zlJsonPutValue(v_Output, 'placer', Nvl(r_费用.开单人, ''));
    zlJsonPutValue(v_Output, 'operator_code', Nvl(r_费用.操作员编号, ''));
    zlJsonPutValue(v_Output, 'operator_name', Nvl(r_费用.操作员姓名, ''));
  
    zlJsonPutValue(v_Output, 'create_time', Nvl(r_费用.登记时间, ''));
    zlJsonPutValue(v_Output, 'happen_time', Nvl(r_费用.发生时间, ''));
  
    zlJsonPutValue(v_Output, 'rec_status', Nvl(r_费用.记录状态, 0), 1);
    zlJsonPutValue(v_Output, 'mrbkfee_sign', Nvl(r_费用.病历费, 0), 1);
  
    zlJsonPutValue(v_Output, 'invoice_no', Nvl(v_发票号, ''));
  
    zlJsonPutValue(v_Output, 'kpbooks_sign', Nvl(r_费用.记帐费用, 0), 1);
    zlJsonPutValue(v_Output, 'fee_status', Nvl(r_费用.费用状态, 0), 1);
  
    zlJsonPutValue(v_Output, 'cardtype_id', To_Number(Nvl(r_费用.卡类别id, '0')), 1);
    zlJsonPutValue(v_Output, 'card_no', r_费用.卡号);
  
    zlJsonPutValue(v_Output, 'sendcard_reg', Nvl(n_挂号同步发卡, 0), 1, 2);
  
    n_总额 := Nvl(n_总额, 0) + Nvl(r_费用.实收金额, 0);
  
    n_结帐id := Nvl(r_费用.结帐id, 0);
    If Nvl(r_费用.记帐费用, 0) = 1 Then
      n_原结帐id := Null;
    End If;
    If v_划价单 Is Null And r_费用.摘要 Is Not Null Then
      Select Max(No) Into v_划价单 From 门诊费用记录 Where no = r_费用.摘要 And 记录性质 = 1 And 病人ID = r_费用.病人id;
    End If;
  End Loop;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not v_划价单 Is Null Then
    If Nvl(n_Nomoved, 0) = 0 Then
      Select Sum(Decode(记录状态, 2, 0, 1) * Decode(a.附加标志, 8, 0, 1) * 应收金额),
             Sum(Decode(记录状态, 2, 0, 1) * Decode(a.附加标志, 8, 0, 1) * 实收金额),
             Sum(Decode(记录状态, 2, 0, 1) * Decode(a.附加标志, 8, 1, 0) * 应收金额),
             Sum(Decode(记录状态, 2, 0, 1) * Decode(a.附加标志, 8, 1, 0) * 实收金额), Sum(Decode(记录状态, 2, 0, 1) * 实收金额),
             Nvl(Sum(实收金额), 0) - Nvl(Sum(结帐金额), 0)
      Into n_卡费应收, n_卡费实收, n_病历费应收, n_病历费实收, n_实收金额, n_剩余金额
      From 门诊费用记录 A
      Where Mod(记录性质, 10) = 1 And NO = v_划价单;
    Else
      Select Sum(Decode(记录状态, 2, 0, 1) * Decode(a.附加标志, 8, 0, 1) * 应收金额),
             Sum(Decode(记录状态, 2, 0, 1) * Decode(a.附加标志, 8, 0, 1) * 实收金额),
             Sum(Decode(记录状态, 2, 0, 1) * Decode(a.附加标志, 8, 1, 0) * 应收金额),
             Sum(Decode(记录状态, 2, 0, 1) * Decode(a.附加标志, 8, 1, 0) * 实收金额), Sum(Decode(记录状态, 2, 0, 1) * 实收金额),
             Nvl(Sum(实收金额), 0) - Nvl(Sum(结帐金额), 0)
      Into n_卡费应收, n_卡费实收, n_病历费应收, n_病历费实收, n_实收金额, n_剩余金额
      From H门诊费用记录 A
      Where Mod(记录性质, 10) = 1 And NO = v_划价单;
    End If;
    zlJsonPutValue(v_Priebill, 'fee_no', v_划价单, 0, 1);
    zlJsonPutValue(v_Priebill, 'cardfee_amrcvb', Nvl(n_卡费应收, 0));
    zlJsonPutValue(v_Priebill, 'cardfee_ampaid', Nvl(n_卡费实收, 0));
    zlJsonPutValue(v_Priebill, 'mrbkfee_amrcvb', Nvl(n_病历费应收, 0));
    zlJsonPutValue(v_Priebill, 'mrbkfee_ampaid', Nvl(n_病历费实收, 0));
  
    If Nvl(n_剩余金额, 0) = Nvl(n_实收金额, 0) Then
      --收费状态:0-未收费;1-已收费;2-已全退
      zlJsonPutValue(v_Priebill, 'charged_statu', 0, 1, 2);
    Elsif Nvl(n_剩余金额, 0) = 0 Then
      zlJsonPutValue(v_Priebill, 'charged_statu', 2, 1, 2);
    Else
      zlJsonPutValue(v_Priebill, 'charged_statu', 1, 1, 2);
    End If;
  End If;

  If v_Priebill Is Not Null Then
    v_Priebill := ',"pricebill_info":' || v_Priebill;
  End If;

  If Not c_Output Is Null Then
    c_Output := To_Clob(',"fee_list":[') || c_Output || To_Clob(']') || To_Clob(v_Priebill);
  Elsif Length(Nvl(v_Output, '') || Nvl(v_Priebill, '')) > 32767 Then
    c_Output := To_Clob(',"fee_list":[') || To_Clob(v_Output) || To_Clob(']') || To_Clob(v_Priebill);
    v_Output := '';
  Else
    v_Output := ',"fee_list":[' || v_Output || ']' || Nvl(v_Priebill, '');
  End If;

  If Nvl(n_结帐id, 0) = 0 Then
    If Not c_Output Is Null Then
      Json_Out := To_Clob('{"output":{"code":1,"message":"成功"') || c_Output || '}}';
    Else
      Json_Out := '{"output":{"code":1,"message":"成功"' || v_Output || '}}';
    End If;
    Return;
  End If;

  Close c_信息;

  If Nvl(n_Nomoved, 0) = 0 Then
    Open c_信息 For
      Select a.No, a.结算方式, Decode(n_查询类型, 1, -1, 1) * Nvl(a.冲预交, 0) As 冲预交, a.关联交易id, a.卡类别id, a.卡号, a.结算卡序号, a.交易流水号,
             a.交易说明, b.性质, a.校对标志, c.消费卡id, a.结算号码, a.摘要
      From 病人预交记录 A, 结算方式 B, 病人卡结算记录 C
      Where a.结算方式 = 名称(+) And a.结帐id = n_结帐id And Mod(a.记录性质, 10) <> 1 And a.Id = c.结算id(+);
  Else
    Open c_信息 For
      Select a.No, a.结算方式, Decode(n_查询类型, 1, -1, 1) * Nvl(a.冲预交, 0) As 冲预交, a.关联交易id, a.卡类别id, a.卡号, a.结算卡序号, a.交易流水号,
             a.交易说明, b.性质, a.校对标志, c.消费卡id, a.结算号码, a.摘要
      From H病人预交记录 A, 结算方式 B, H病人卡结算记录 C
      Where a.结算方式 = 名称(+) And a.结帐id = n_结帐id And Mod(a.记录性质, 10) <> 1 And a.Id = c.结算id(+);
  End If;

  Loop
    Fetch c_信息
      Into r_结算信息;
    Exit When c_信息%NotFound;
  
    zlJsonPutValue(v_Balanceinf, 'blnc_mode', r_结算信息.结算方式, 0, 1);
    zlJsonPutValue(v_Balanceinf, 'balance_id', n_结帐id, 1);
    If n_查询类型 = 2 And Nvl(r_结算信息.性质, 0) <> 9 Then
      --只有一条数据,多条数量时
      zlJsonPutValue(v_Balanceinf, 'blnc_money', n_总额, 1);
    Else
      zlJsonPutValue(v_Balanceinf, 'blnc_money', Nvl(r_结算信息.冲预交, 0), 1);
    End If;
  
    zlJsonPutValue(v_Balanceinf, 'pay_cardno', Nvl(r_结算信息.卡号, ''));
    zlJsonPutValue(v_Balanceinf, 'pay_swapno', Nvl(r_结算信息.交易流水号, ''));
    zlJsonPutValue(v_Balanceinf, 'pay_swapmemo', Nvl(r_结算信息.交易说明, ''));
  
    zlJsonPutValue(v_Balanceinf, 'relation_id', Nvl(r_结算信息.关联交易id, 0), 1);
  
    If Nvl(r_结算信息.结算卡序号, 0) <> 0 Then
      zlJsonPutValue(v_Balanceinf, 'cardtype_id', Nvl(r_结算信息.结算卡序号, 0), 1);
      zlJsonPutValue(v_Balanceinf, 'consume_card', 1, 1);
    Else
      zlJsonPutValue(v_Balanceinf, 'cardtype_id', Nvl(r_结算信息.卡类别id, 0), 1);
      zlJsonPutValue(v_Balanceinf, 'consume_card', 0, 1);
    End If;
  
    zlJsonPutValue(v_Balanceinf, 'consume_card_id', To_Number(Nvl(r_结算信息.消费卡id, '0')), 1);
    zlJsonPutValue(v_Balanceinf, 'blnc_nature', Nvl(r_结算信息.性质, 0), 1);
    zlJsonPutValue(v_Balanceinf, 'blnc_statu', Nvl(r_结算信息.校对标志, 0), 1);
    zlJsonPutValue(v_Balanceinf, 'blnc_no', Nvl(r_结算信息.结算号码, ''));
    zlJsonPutValue(v_Balanceinf, 'blnc_memo', Nvl(r_结算信息.摘要, ''));
    If n_查询类型 = 2 And Nvl(r_结算信息.性质, 0) <> 9 Then
      zlJsonPutValue(v_Balanceinf, 'original_money', Nvl(r_结算信息.冲预交, 0), 1);
    Else
      zlJsonPutValue(v_Balanceinf, 'original_money', 0, 1);
    End If;
  
    zlJsonPutValue(v_Balanceinf, 'original_id', Nvl(n_原结帐id, 0), 1, 2);
  
  End Loop;

  If v_Balanceinf Is Not Null Then
    v_Balanceinf := ',"balance_list":[' || v_Balanceinf || ']';
  
    If Not c_Output Is Null Then
      c_Output := c_Output || To_Clob(v_Balanceinf);
    
    Elsif Length(Nvl(v_Output, '') || Nvl(v_Balanceinf, '')) > 32767 Then
      c_Output := To_Clob(v_Output) || To_Clob(Nvl(v_Balanceinf, ''));
      v_Output := '';
    Else
      v_Output := Nvl(v_Output, '') || Nvl(v_Balanceinf, '');
    End If;
  
  End If;

  --3.预交单据
  If Nvl(n_查询预交, 0) = 1 Then
  
    Close c_信息;
  
    If Nvl(n_Nomoved, 0) = 0 Then
      Open c_信息 For
        Select a.No, Max(a.Id) As 预交id, Decode(n_查询类型, 1, -1, 1) * Nvl(Sum(a.金额), 0) As 金额, Max(a.结算方式) As 结算方式,
               Nvl(Sum(a.金额), 0) As 冲预交, Max(a.关联交易id) As 关联交易id, Max(a.卡类别id) As 卡类别id,
               Max(Decode(a.记录性质, 1, a.卡号, '')) As 卡号, Max(a.结算卡序号) As 结算卡序号,
               Max(Decode(a.记录性质, 1, a.交易流水号, '')) As 交易流水号, Max(Decode(a.记录性质, 1, a.交易说明, '')) 交易说明, Max(b.性质) As 性质,
               Max(Decode(a.记录性质, 1, a.校对标志, 0)) As 校对标志, Max(c.消费卡id) As 消费卡id, Max(a.结算号码) As 结算号码, Max(a.摘要) As 摘要
        From 病人预交记录 A, 结算方式 B, 病人卡结算记录 C
        Where a.结算方式 = 名称(+) And a.关联交易id In (Select 关联交易id From 病人预交记录 Where 结帐id = n_结帐id) And Mod(a.记录性质, 10) = 1 And
              a.Id = c.结算id(+) And (Nvl(n_查询类型, 0) = 1 And a.记录状态 = 2 Or Nvl(n_查询类型, 0) <> 1)
        Group By NO;
    Else
      Open c_信息 For
        Select a.No, Max(a.Id) As 预交id, Decode(n_查询类型, 1, -1, 1) * Nvl(Sum(a.金额), 0) As 金额, Max(a.结算方式) As 结算方式,
               Nvl(Sum(a.金额), 0) As 冲预交, Max(a.关联交易id) As 关联交易id, Max(a.卡类别id) As 卡类别id,
               Max(Decode(a.记录性质, 1, a.卡号, '')) As 卡号, Max(a.结算卡序号) As 结算卡序号,
               Max(Decode(a.记录性质, 1, a.交易流水号, '')) As 交易流水号, Max(Decode(a.记录性质, 1, a.交易说明, '')) 交易说明, Max(b.性质) As 性质,
               Max(Decode(a.记录性质, 1, a.校对标志, 0)) As 校对标志, Max(c.消费卡id) As 消费卡id, Max(a.结算号码) As 结算号码, Max(a.摘要) As 摘要
        From H病人预交记录 A, 结算方式 B, H病人卡结算记录 C
        Where a.结算方式 = 名称(+) And a.关联交易id In (Select 关联交易id From H病人预交记录 Where 结帐id = n_结帐id) And Mod(a.记录性质, 10) = 1 And
              a.Id = c.结算id(+) And (Nvl(n_查询类型, 0) = 1 And a.记录状态 = 2 Or Nvl(n_查询类型, 0) <> 1)
        Group By NO;
    End If;
    Loop
    
      Fetch c_信息
        Into r_预交信息;
      Exit When c_信息%NotFound;
    
      --只有一条
    
      zlJsonPutValue(v_Deposit, 'deposit_id', r_预交信息.预交id, 1, 1);
      zlJsonPutValue(v_Deposit, 'deposit_no', r_预交信息.No);
      zlJsonPutValue(v_Deposit, 'deposit_money', Nvl(r_预交信息.金额, 0), 1);
    
      zlJsonPutValue(v_Deposit, 'blnc_mode', r_预交信息.结算方式);
      zlJsonPutValue(v_Deposit, 'balance_id', n_结帐id, 1);
      zlJsonPutValue(v_Deposit, 'pay_cardno', Nvl(r_预交信息.卡号, ''));
      zlJsonPutValue(v_Deposit, 'pay_swapno', Nvl(r_预交信息.交易流水号, ''));
      zlJsonPutValue(v_Deposit, 'pay_swapmemo', Nvl(r_预交信息.交易说明, ''));
    
      zlJsonPutValue(v_Deposit, 'relation_id', Nvl(r_预交信息.关联交易id, 0), 1);
    
      If Nvl(r_预交信息.结算卡序号, 0) <> 0 Then
        zlJsonPutValue(v_Deposit, 'cardtype_id', Nvl(r_预交信息.结算卡序号, 0), 1);
        zlJsonPutValue(v_Deposit, 'consume_card', 1, 1);
      Else
        zlJsonPutValue(v_Deposit, 'cardtype_id', Nvl(r_预交信息.卡类别id, 0), 1);
        zlJsonPutValue(v_Deposit, 'consume_card', 0, 1);
      End If;
      zlJsonPutValue(v_Deposit, 'consume_card_id', To_Number(Nvl(r_结算信息.消费卡id, '0')), 1);
    
      zlJsonPutValue(v_Deposit, 'blnc_nature', Nvl(r_预交信息.性质, 0), 1);
      zlJsonPutValue(v_Deposit, 'blnc_statu', Nvl(r_预交信息.校对标志, 0), 1);
      zlJsonPutValue(v_Deposit, 'blnc_no', Nvl(r_结算信息.结算号码, ''));
      zlJsonPutValue(v_Deposit, 'blnc_memo', Nvl(r_结算信息.摘要, ''), 0, 2);
    
      n_Find := 1;
      Exit;
    End Loop;
  
    If v_Deposit Is Not Null Then
      v_Deposit := ',"deposit_info":' || v_Deposit;
    
      If Not c_Output Is Null Then
        c_Output := c_Output || To_Clob(v_Deposit);
      Elsif Length(Nvl(v_Output, '') || Nvl(v_Deposit, '')) > 32767 Then
        c_Output := To_Clob(v_Output) || To_Clob(Nvl(v_Deposit, ''));
        v_Output := '';
      
      Else
        v_Output := Nvl(v_Output, '') || Nvl(v_Deposit, '');
      End If;
    End If;
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"成功"') || c_Output || '}}';
  Else
    Json_Out := '{"output":{"code":1,"message":"成功"' || v_Output || '}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getcardfeeinfobyno;
/

Create Or Replace Procedure Zl_Exsesvr_Chkfeecategorydept
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：检查费别适用于所有科室,或当前指定科室
  --入参：Json_In:格式
  --  input
  --    fee_category         N 1 费别
  --    pati_deptid         N 1 病人科室ID
  --出参: Json_Out,格式如下
  --  output
  --    code            N 1 应答吗：0-失败；1-成功
  --    message         C 1 应答消息：失败时返回具体的错误信息
  --    isexist          N 1 费别是否存在：0-不存在；1-存在
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_Count        Number(1);
  n_Pati_Deptid  Number(18);
  v_Fee_Category Varchar2(20);
Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_Fee_Category := j_Json.Get_String('fee_category');
  n_Pati_Deptid  := j_Json.Get_Number('pati_deptid');
  Select Count(1)
  Into n_Count
  From Dual
  Where Not Exists (Select 1 From 费别适用科室 Where 费别 = v_Fee_Category) Or Exists
   (Select 1 From 费别适用科室 Where 费别 = v_Fee_Category And 科室id = n_Pati_Deptid);
  If n_Count > 1 Then
    n_Count := 1;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功","isexist":' || n_Count || '}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Chkfeecategorydept;
/


Create Or Replace Procedure Zl_Exsesvr_Cardfeeisbalance
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  ---------------------------------------------------------------------------
  --功能：检查卡费是否已经结帐
  --入参      json
  --input     
  --  cardfee_no            C 1 卡费对应的费用单据号
  --  rdcardfee_sign        N 1 读取卡费标志:0-读取卡费,1-病历费;2-卡费或病历费用
  --出参      json
  --output      
  --  code                      C 1 应答码：0-失败；1-成功
  --  message                   C 1 应答消息：失败时返回具体的错误信息
  --  isbalanced                N 1 是否已经结帐:1-已结结帐;0-未结帐
  --  blnc_no                   C 1 结帐单据号
  ---------------------------------------------------------------------------
  j_Input    PLJson;
  j_Json     PLJson;
  v_Output   Varchar2(32767);
  v_No       住院费用记录.No%Type;
  n_标志     Number(1);
  n_已结帐   Number(2);
  v_结帐单号 病人结帐记录.No%Type;

Begin

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_No   := j_Json.Get_String('cardfee_no');
  n_标志 := Nvl(j_Json.Get_Number('rdcardfee_sign'), 0);

  If n_标志 = 0 Then
    Select Max(b.No), Count(1)
    Into v_结帐单号, n_已结帐
    From 住院费用记录 A, 病人结帐记录 B
    Where a.结帐id = b.Id And a.记录性质 In (5, 15) And a.记录状态 = 1 And b.记录状态 = 1 And a.No = v_No And Nvl(a.附加标志, 0) <> 8;
  Elsif n_标志 = 1 Then
    Select Max(b.No), Count(1)
    Into v_结帐单号, n_已结帐
    From 住院费用记录 A, 病人结帐记录 B
    Where a.结帐id = b.Id And a.记录性质 In (5, 15) And a.记录状态 = 1 And b.记录状态 = 1 And a.No = v_No And Nvl(a.附加标志, 0) = 8;
  Else
    Select Max(b.No), Count(1)
    Into v_结帐单号, n_已结帐
    From 住院费用记录 A, 病人结帐记录 B
    Where a.结帐id = b.Id And a.记录性质 In (5, 15) And a.记录状态 = 1 And b.记录状态 = 1 And a.No = v_No;
  End If;

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '成功');
  zlJsonPutValue(v_Output, 'isbalanced', n_已结帐, 1);
  zlJsonPutValue(v_Output, 'blnc_no', v_结帐单号, 0, 2);
  Json_Out := '{"output":' || v_Output || '}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Cardfeeisbalance;
/


Create Or Replace Procedure Zl_Exsesvr_Recalcfee
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：病人未结费用重算
  --入参：Json_In:格式
  --    input
  --      pati_id            N 1  病人id
  --      pati_pageid        N 1  主页ID
  --      pati_nature        N 1 病人性质:0-普通住院病人,1-门诊留观病人,2-住院留观病人
  --      outfee             N 1  是否门诊费别 
  --      fee_type           C 1  费别
  --出参: Json_Out,格式如下
  --    output
  --        code                    N   1   应答吗：0-失败；1-成功
  --        message                 C   1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_Pati_Id     Number(18);
  n_Pati_Pageid Number(18);
  n_Outfee      Number;
  n_Pati_Nature Number;
  v_Feetype     Varchar2(1000);
Begin
  --解析入参
  j_Input       := PLJson(Json_In);
  j_Json        := j_Input.Get_Pljson('input');
  n_Pati_Id     := j_Json.Get_Number('pati_id');
  n_Pati_Pageid := j_Json.Get_Number('pati_pageid');
  n_Outfee      := j_Json.Get_Number('outfee');
  n_Pati_Nature := j_Json.Get_Number('pati_nature');
  v_Feetype     := j_Json.Get_String('fee_type');
  If Nvl(n_Outfee, 0) = 0 Then
    Zl_病人未结费用_Recalc_s(n_Pati_Id, n_Pati_Pageid, n_Pati_Nature, v_Feetype);
  Else
    Zl_病人未结门诊费用_Recalc_s(n_Pati_Id, v_Feetype);
  End If;
  Json_Out := zlJsonOut('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Recalcfee;
/


CREATE OR REPLACE Procedure Zl_Exsesvr_Addcardfeeinfo
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:增加卡费及预交款数据
  --入参：Json_In:格式
  --  input
  --    oper_fun            C     操作状态:0-正常的预交款或卡费缴款;1-保存为未生效的预交款或异常的卡费;2-保存为记帐单;3-保存为划价单
  --    blnc_money          N  1  本次结算总计:预交+卡费
  --    balance_id          N     结帐id
  --    pati_info           C     病人信息
  --      pati_id           C  1  病人ID
  --      pati_pageid       N  1  主页id
  --      pati_name         C  1  病人姓名
  --      pati_sex          C  1  性别
  --      pati_age          C  1  年龄
  --      outpno  N         1     门诊号
  --      mdlmode_name      C  1  付款方式名称
  --      fee_category      C  1  费别
  --      insurance_type    N     险类
  --    card_info           C     医疗卡信息
  --      cardno            C  1  发卡卡号
  --      cardtype_id       N  1  发卡卡类别ID
  --      send_mode         N  1  发卡方式;0-发卡,1-补卡,2-换卡
  --      cardno_reusing    N  1  卡号重用:1-重用;0-不以许重用
  --      recv_id           N  1  领用id:领用Id
  --      cardno_old        C  1  原卡卡号:换卡时，需要传入原卡号
  --    deposit_info        C  1  预交款列表
  --      deposit_no        C  1  预交单据号
  --      deposit_id        N     预交ID
  --      fact_no           C  1  发票号
  --      deposit_type      N     预交类别:1-门诊;2-住院
  --      pati_id           N  1  病人id
  --      pati_pageid       N  1  主页id
  --      dept_id           N  1  缴款科室id
  --      money             N  1  缴款金额
  --      emp_name          C  1  缴款单位
  --      emp_bank_name     C  1  单位开户行
  --      emp_bank_actno    C  1  开户行账号
  --      memo              C  1  摘要
  --      recv_id           N  1  票据领用id
  --      start_einv        N  1  是否启用电子票据:1-启用;0-不启用
  --    cardfee_list[]      C  1  卡费列表
  --      fee_no            C  1  费用单据号
  --      serial_num        N  1  序号
  --      price_ftrnum      N  1  价格父号
  --      subde_ftrnum      N  1  从属父号
  --      receipt_type      C  1  收费类别
  --      fitem_id          N  1  收费细目id
  --      income_item_id    N  1  收入项目id
  --      price             N  1  标准单价
  --      receipt_fee       C  1  收据费目
  --      fee_amrcvb        N  1  应收金额
  --      fee_ampaib        N  1  实收金额
  --      pati_deptid       N  1  病人科室id
  --      pati_wardarea_id  N  1  病人病区id
  --      exedept_id        N  1  执行部门id
  --      bill_deptid       N  1  开单部门id
  --      mrbkfee_sign      N  1  是否病历费:1-是病历费;0-不是病历费
  --      insurance_code    C  1  保险编码
  --      insurance_type_id N  1  保险大类id
  --      insurance_sign    N  1  保险项目否:1-是保险项目;0-不是保险项目
  --      si_manp_money     N  1  统筹金额
  --      memo              C  1  摘要
  --      overtime_flag     N  1  加班标志
  --      cardno            C  1  发卡卡号
  --      cardtype_id       N  1  发卡卡类别ID
  --      send_mode         N  1  发卡方式;0-发卡,1-补卡,2-换卡
  --    balance_info        C     结算信息:目前只支持一种结算方式
  --      blnc_mode         C  1  结算方式
  --      blnc_no           C  1  结算号码
  --      cardtype_id       C  1  卡类别id
  --      consumer_no       C  1  结算卡序号，即卡消费接口目录.编号
  --      consume_card_id   N  1  消费卡ID
  --      cardno            C  1  支付卡号
  --      swapno            C  1  交易流水号
  --      swapmemo          C  1  交易说明
  --      cprtion_unit      C  1  合作单位
  --      start_einv        N  1  是否启用电子票据:1-启用;0-不启用
  --    operator_name       C  1  操作员姓名
  --    operator_code       C  1  操作员编号
  --    create_time         C  1  登记时间或收款时间:yyyy-mm-dd hh:mi:ss

  --出参: Json_Out,格式如下
  --  output
  --    code  C  1  应答码：0-失败；1-成功
  --    message  C  1  应答消息：失败时返回具体的错误信息
  --    cardfee_no  C  1  卡费的费用单据号
  --    deposit_no  C  1  预交单据号
  --    deposit_id  N 1 预交ID
  --    balance_id  N 1 结帐ID
  ---------------------------------------------------------------------------
  j_Input    PLJson;
  j_Json     PLJson;
  o_Json     PLJson;
  j_Billlist Pljson_List := Pljson_List();

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  --本次结算信息
  n_操作状态     Number(2);
  n_结帐id       门诊费用记录.结帐id%Type;
  n_本次结算总计 门诊费用记录.实收金额%Type;

  v_操作员姓名 门诊费用记录.操作员姓名%Type;
  v_操作员编号 门诊费用记录.操作员编号%Type;
  v_原单据号   门诊费用记录.No%Type;

  d_登记时间 门诊费用记录.登记时间%Type;
  n_卡号重用 Number(2);
  n_结算总额 Number(16, 5);
  n_是否划价 Number(2);
  n_性质     Number(5);
  n_消费卡id Number(18);
  n_计费方式 Number(18);
  n_返回值   病人预交记录.冲预交%Type;
  v_原卡号   票据使用明细.号码%Type;
  --病人信息相关定义

  n_病人id       门诊费用记录.病人id%Type;
  v_病人姓名     门诊费用记录.姓名%Type;
  v_性别         门诊费用记录.性别%Type;
  v_年龄         门诊费用记录.年龄%Type;
  n_门诊号       Number(18);
  n_住院号       Number(18);
  v_付款方式名称 医疗付款方式.名称%Type;
  v_费别         门诊费用记录.费别%Type;
  -- n_险类         保险结算记录.险类%Type;
  n_病人主页id 病人预交记录.主页id%Type;
  --费用相关定义
  v_费用单号     门诊费用记录.No%Type;
  v_划价单       门诊费用记录.No%Type;
  n_序号         门诊费用记录.序号%Type;
  n_价格父号     门诊费用记录.价格父号%Type;
  n_从属父号     门诊费用记录.从属父号%Type;
  v_收费类别     门诊费用记录.收费类别%Type;
  n_收费细目id   门诊费用记录.收费细目id%Type;
  n_收入项目id   门诊费用记录.收入项目id%Type;
  n_标准单价     门诊费用记录.标准单价%Type;
  v_收据费目     门诊费用记录.收据费目%Type;
  n_应收金额     门诊费用记录.应收金额%Type;
  n_实收金额     门诊费用记录.实收金额%Type;
  n_病人科室id   门诊费用记录.病人科室id%Type;
  n_开单部门id   门诊费用记录.开单部门id%Type;
  n_是否病历费   Number(2);
  n_是否记帐     Number(2);
  n_发卡方式     Number(2);
  v_保险编码     门诊费用记录.保险编码%Type;
  n_保险大类id   门诊费用记录.保险大类id%Type;
  n_保险项目否   门诊费用记录.保险项目否%Type;
  n_统筹金额     门诊费用记录.统筹金额%Type;
  v_费用摘要     门诊费用记录.摘要%Type;
  v_发卡卡号     病人预交记录.卡号%Type;
  n_发卡卡类别id 病人预交记录.卡类别id%Type;
  n_发卡领用id   票据使用明细.领用id%Type;

  n_病人病区id 门诊费用记录.病人病区id%Type;
  v_计算单位   门诊费用记录.计算单位%Type;
  n_执行部门id 门诊费用记录.执行部门id%Type;

  n_费用状态 门诊费用记录.费用状态%Type;
  n_校对标志 病人预交记录.校对标志%Type;
  n_附加标志 Number(2);
  n_加班标志 住院费用记录.加班标志%Type;
  --支付方式定义
  v_结算方式   病人预交记录.结算方式%Type;
  v_结算号码   病人预交记录.结算号码%Type;
  n_卡类别id   病人预交记录.卡类别id%Type;
  n_结算卡序号 病人预交记录.结算卡序号%Type;
  v_支付卡号   病人预交记录.卡号%Type;
  v_交易流水号 病人预交记录.交易流水号%Type;
  v_交易说明   病人预交记录.交易说明%Type;
  v_合作单位   病人预交记录.合作单位%Type;
  n_关联交易id 病人预交记录.关联交易id%Type;

  --预交相关变量定义
  n_预交id       病人预交记录.Id%Type;
  v_预交单号     病人预交记录.No%Type;
  v_发票号       票据使用明细.号码%Type;
  n_预交类别     病人预交记录.预交类别%Type;
  n_主页id       病人预交记录.主页id%Type;
  n_缴款科室id   病人预交记录.科室id%Type;
  n_缴款金额     病人预交记录.金额%Type;
  v_缴款单位     病人预交记录.缴款单位%Type;
  v_单位开户行   病人预交记录.单位开户行%Type;
  v_开户行账号   病人预交记录.单位帐号%Type;
  v_摘要         病人预交记录.摘要%Type;
  n_领用id       票据使用明细.领用id%Type;
  n_组id         Number(18);
  n_Count        Number(10);
  n_预交电子票据 病人预交记录.预交电子票据%Type;
  n_是否电子票据 病人预交记录.是否电子票据%Type;
  n_更新预交余额 Number(1);
Begin
  --解析入参

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_操作状态     := Nvl(j_Json.Get_Number('oper_fun'), 0);
  n_本次结算总计 := j_Json.Get_Number('blnc_total');
  v_操作员姓名   := j_Json.Get_String('operator_name');
  v_操作员编号   := j_Json.Get_String('operator_code');
  d_登记时间     := To_Date(j_Json.Get_String('create_time'), 'YYYY-MM-DD hh24:mi:ss');
  n_结帐id       := j_Json.Get_Number('balance_id');

  n_组id := Zl_Get组id(v_操作员姓名);

  If d_登记时间 Is Null Then
    d_登记时间 := Sysdate;
  End If;

  --1.获取病人信息
  o_Json := j_Json.Get_Pljson('pati_info');
  If o_Json Is Null Then
    v_Err_Msg := '不存在病人信息数据，请检查';
    Raise Err_Item;
  End If;

  n_病人id     := o_Json.Get_Number('pati_id');
  n_病人主页id := o_Json.Get_Number('pati_pageid');
  v_病人姓名   := o_Json.Get_String('pati_name');
  v_性别       := o_Json.Get_String('pati_sex');
  v_年龄       := o_Json.Get_String('pati_age');
  n_门诊号     := To_Number(o_Json.Get_String('outpatient_num'));
  n_住院号     := To_Number(o_Json.Get_String('inpatient_num'));

  v_付款方式名称 := o_Json.Get_String('mdlpay_name');
  v_费别         := o_Json.Get_String('fee_category');
  --n_险类         := o_Json.Get_Number('insurance_type');

  o_Json := PLJson();
  o_Json := j_Json.Get_Pljson('card_info');
  If o_Json Is Null Then
    v_Err_Msg := '不存在发卡或绑定卡等信息，请检查!';
    Raise Err_Item;
  End If;

  v_发卡卡号     := o_Json.Get_String('cardno');
  n_发卡卡类别id := o_Json.Get_Number('cardtype_id');
  n_发卡方式     := o_Json.Get_Number('send_mode');
  n_发卡领用id   := o_Json.Get_Number('recv_id');
  n_卡号重用     := o_Json.Get_Number('cardno_reusing');
  v_原卡号       := o_Json.Get_String('cardno_old');
  --2.处理费用
  j_Billlist := j_Json.Get_Pljson_List('cardfee_list');
  If j_Billlist Is Null Then
    v_Err_Msg := '不存在卡费所涉及的费用信息，请检查';
    Raise Err_Item;
  End If;
  If j_Billlist.Count = 0 Then
    v_Err_Msg := '不存在卡费所涉及的费用信息，请检查';
    Raise Err_Item;
  End If;

  n_校对标志 := Null;
  n_费用状态 := Null;
  n_是否划价 := 0;

  If Nvl(n_操作状态, 0) = 2 Then
    --2-保存为记帐单
    n_结帐id   := Null;
    n_是否记帐 := 1;
  Elsif n_操作状态 = 3 Then
    --3.保存为划价单
    v_划价单   := Nextno(13);
    n_是否划价 := 1;
  Elsif n_操作状态 = 1 Then
    --1-保存为未生效的预交款或异常的卡费
    n_校对标志 := 1;
    n_费用状态 := 1;

  Else
    --0-正常的预交款或卡费缴款
    n_更新预交余额 := 1;
  End If;

  If Nvl(n_操作状态, 0) <> 2 Then
    --不是记帐时，都存在结帐id
    If Nvl(n_结帐id, 0) = 0 Then
      Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
    End If;
  End If;

  For J In 1 .. j_Billlist.Count Loop
    o_Json       := PLJson(j_Billlist.Get(J));
    v_费用单号   := o_Json.Get_String('fee_no');
    n_序号       := o_Json.Get_Number('serial_num');
    n_价格父号   := o_Json.Get_Number('price_ftrnum');
    n_从属父号   := o_Json.Get_Number('subde_ftrnum');
    v_收费类别   := o_Json.Get_String('receipt_type');
    n_收费细目id := Nvl(o_Json.Get_Number('fitem_id'), 0);
    n_收入项目id := Nvl(o_Json.Get_Number('income_item_id'), 0);
    n_标准单价   := Nvl(o_Json.Get_Number('price'), 0);
    v_收据费目   := o_Json.Get_String('receipt_fee');
    n_应收金额   := Nvl(o_Json.Get_Number('fee_amrcvb'), 0);
    n_实收金额   := Nvl(o_Json.Get_Number('fee_ampaib'), 0);
    n_病人科室id := Nvl(o_Json.Get_Number('pati_deptid'), 0);
    n_病人病区id := Nvl(o_Json.Get_Number('pati_wardarea_id'), 0);
    n_开单部门id := Nvl(o_Json.Get_Number('bill_deptid'), 0);
    n_执行部门id := Nvl(o_Json.Get_Number('exedept_id'), 0);
    n_是否病历费 := Nvl(o_Json.Get_Number('mrbkfee_sign'), 0);
    v_保险编码   := o_Json.Get_String('insurance_code');
    n_保险大类id := o_Json.Get_Number('insurance_type_id');
    n_保险项目否 := o_Json.Get_Number('insurance_sign');
    n_统筹金额   := o_Json.Get_Number('si_manp_money');
    v_费用摘要   := o_Json.Get_String('memo');
    n_加班标志   := o_Json.Get_Number('overtime_flag');

    If v_费用单号 Is Null Then
      v_Err_Msg := '不存在指定的费用单据号，请检查';
      Raise Err_Item;
    End If;
    If Nvl(n_序号, 0) = 0 Then
      n_序号 := 1;
    End If;
    If Nvl(n_价格父号, 0) = 0 Then
      n_价格父号 := Null;
    End If;
    If Nvl(n_从属父号, 0) = 0 Then
      n_从属父号 := Null;
    End If;
    If n_是否划价 = 1 Then
      v_费用摘要 := v_划价单;
    End If;
    Select Max(计算单位) Into v_计算单位 From 收费项目目录 Where ID = n_收费细目id;

    --计费方式_In：0-收费;1-划价;2-记账
    If Nvl(n_是否划价, 0) = 1 Then
      n_计费方式 := 1;
    Elsif Nvl(n_是否记帐, 0) = 1 Then
      n_计费方式 := 2;
    Else
      n_计费方式 := 0;
    End If;
    n_附加标志 := n_发卡方式;

    If Nvl(n_是否病历费, 0) = 1 Then
      n_附加标志 := 8; --病历费
    End If;

    Zl_病人卡费_Insert_s(n_计费方式, v_费用单号, n_病人id, n_病人主页id, n_门诊号, v_费别, v_病人姓名, v_性别, v_年龄, v_付款方式名称, n_病人病区id, n_病人科室id,
                     n_收费细目id, v_收费类别, v_计算单位, n_收入项目id, v_收据费目, n_标准单价, n_应收金额, n_实收金额, n_执行部门id, n_开单部门id, v_操作员编号,
                     v_操作员姓名, n_加班标志, d_登记时间, n_发卡卡类别id, v_发卡卡号, v_费用摘要, n_结帐id, n_附加标志, v_划价单, n_序号, n_费用状态);

    n_结算总额 := Nvl(n_结算总额, 0) + Nvl(n_实收金额, 0);

  End Loop;

  --需要处理医疗卡的票据使用
  If n_卡号重用 = 0 Then
    --需要检查是否存在票据使用明细，如果存在，肯定会发生错误
    If Nvl(n_发卡领用id, 0) = 0 Then
      Select Nvl(Max(性质), 0)
      Into n_性质
      From 票据使用明细 A
      Where a.票种 = 5 And a.号码 = v_发卡卡号 And Nvl(a.领用id, 0) = 0;

    Else
      Select Nvl(Max(性质), 0)
      Into n_性质
      From 票据使用明细 A, 票据领用记录 B
      Where a.票种 = 5 And a.号码 = v_发卡卡号 And a.领用id = n_发卡领用id And a.领用id = b.Id;
    End If;
    If n_性质 <> 0 Then
      v_Err_Msg := '卡号:' || v_发卡卡号 || ' 已经使用，不能再进行发卡操作,请检查!';
      Raise Err_Item;
    End If;
  End If;

  --发卡方式;0-发卡,1-补卡,2-换卡
  If Nvl(n_操作状态, 0) <> 1 Then
    --变动类型_In=1-发卡 ;2-换卡;3-补卡 ;4-退卡
    n_Count := Case
                 When Nvl(n_发卡方式, 0) = 0 Then
                  1
                 When Nvl(n_发卡方式, 0) = 1 Then
                  3
                 When Nvl(n_发卡方式, 0) = 2 Then
                  2
                 Else
                  4
               End;

    If n_Count = 2 Then
      --需要获取原始发卡单据
      Select Max(NO)
      Into v_原单据号
      From 住院费用记录
      Where 记录性质 = 5 And 病人id = n_病人id And 实际票号 = v_原卡号 And To_Number(Nvl(结论, '0')) = Nvl(n_发卡卡类别id, 0) And 附加标志 <> 8;
      If v_原单据号 Is Null Then

        v_Err_Msg := '未找到原始卡号:' || v_原卡号 || '的费用单据 ，不能再进行换卡操作,请检查!';
        Raise Err_Item;
      End If;

      --附加标志:0-发卡，1-补卡，2-换卡
      Update 住院费用记录
      Set 附加标志 = Decode(Nvl(附加标志, 0), 8, 8, n_发卡方式)
      Where 记录性质 = 5 And NO = v_原单据号;

      Zl_病人医疗卡票据_Update_s(n_Count, v_发卡卡号, v_操作员姓名, d_登记时间, v_原单据号, n_发卡领用id, v_原卡号, n_卡号重用);
    Else
      Zl_病人医疗卡票据_Update_s(n_Count, v_发卡卡号, v_操作员姓名, d_登记时间, v_费用单号, n_发卡领用id, Null, n_卡号重用);
    End If;

  End If;

  If Nvl(n_是否记帐, 0) = 1 Then
    Json_Out := '{"output":{"code":1,"message":"' || '成功' || '","deposit_id":' || Nvl(n_预交id, 0) || ',"balance_id":' ||
                Nvl(n_结帐id, 0) || '}}';
    Return;
  End If;

  If Nvl(n_是否划价, 0) = 1 Then
    Json_Out := '{"output":{"code":1,"message":"' || '成功' || '","deposit_id":' || Nvl(n_预交id, 0) || ',"balance_id":' ||
                Nvl(n_结帐id, 0) || '}}';
    Return;
  End If;

  --3.处理结算信息
  o_Json := PLJson();
  o_Json := j_Json.Get_Pljson('balance_info');
  If Not o_Json Is Null Then
    v_结算方式     := o_Json.Get_String('blnc_mode');
    v_结算号码     := o_Json.Get_String('blnc_no');
    n_卡类别id     := o_Json.Get_Number('cardtype_id');
    n_结算卡序号   := o_Json.Get_Number('consumer_no');
    v_支付卡号     := o_Json.Get_String('cardno');
    v_交易流水号   := o_Json.Get_String('swapno');
    v_交易说明     := o_Json.Get_String('swapmemo');
    v_合作单位     := o_Json.Get_String('cprtion_unit');
    n_消费卡id     := o_Json.Get_Number('consume_card_id');
    n_是否电子票据 := o_Json.Get_Number('start_einv');

    If Nvl(n_结帐id, 0) = 0 Then
      v_Err_Msg := '结帐ID读取不正确,请检查!';
      Raise Err_Item;
    End If;
    If Nvl(n_卡类别id, 0) = 0 Then
      n_卡类别id := Null;
    End If;
    If Nvl(n_结算卡序号, 0) = 0 Then
      n_结算卡序号 := Null;
    End If;
    If n_是否电子票据 Is Null Then
      n_是否电子票据 := Zl_Fun_Isstarteinvoice(5, 0);
    End If;

    Update 病人预交记录
    Set 结算方式 = v_结算方式, 校对标志 = n_校对标志, 卡类别id = n_卡类别id, 结算卡序号 = n_结算卡序号, 卡号 = v_支付卡号, 交易流水号 = v_交易流水号, 交易说明 = v_交易说明,
        结算号码 = v_结算号码, 摘要 = '医疗卡费用', 交易人员 = v_操作员姓名, 交易时间 = d_登记时间, 结算性质 = 5, 收款时间 = d_登记时间, 操作员编号 = v_操作员编号,
        操作员姓名 = v_操作员姓名, 病人id = n_病人id, 主页id = Decode(n_病人主页id, 0, Null, n_病人主页id), 姓名 = v_病人姓名, 性别 = v_性别, 年龄 = v_年龄,
        门诊号 = n_门诊号, 住院号 = n_住院号, 是否电子票据 = Nvl(n_是否电子票据, 0)
    Where 结帐id = Nvl(n_结帐id, 0) And 结算方式 Is Null
    Returning ID, 关联交易id, 冲预交 Into n_预交id, n_关联交易id, n_返回值;

    If Sql%NotFound Then

      --插入结算数据
      Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      If Nvl(n_关联交易id, 0) = 0 Then
        n_关联交易id := n_预交id;

      End If;
      Insert Into 病人预交记录
        (ID, NO, 记录性质, 记录状态, 病人id, 主页id, 姓名, 性别, 年龄, 门诊号, 住院号, 付款方式名称, 科室id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 结算序号,
         摘要, 缴款组id, 卡类别id, 卡号, 结算卡序号, 交易流水号, 交易说明, 合作单位, 结算性质, 关联交易id, 校对标志, 交易人员, 交易时间, 是否电子票据)
      Values
        (n_预交id, v_费用单号, 5, 1, n_病人id, Decode(n_病人主页id, 0, Null, n_病人主页id), v_病人姓名, v_性别, v_年龄, n_门诊号, n_住院号, v_付款方式名称,
         Decode(n_病人科室id, 0, Null, n_病人科室id), v_结算方式, d_登记时间, v_操作员编号, v_操作员姓名, n_结算总额, n_结帐id, -1 * n_结帐id, '医疗卡费用',
         n_组id, n_卡类别id, v_支付卡号, n_结算卡序号, v_交易流水号, v_交易说明, v_合作单位, 5, n_关联交易id, n_校对标志, v_操作员姓名, d_登记时间,
         Nvl(n_是否电子票据, 0));
    Elsif Nvl(n_返回值, 0) <> Nvl(n_结算总额, 0) Then
      v_Err_Msg := '结算金额不正确,请检查!';
      Raise Err_Item;

    End If;

    If Nvl(n_结算卡序号, 0) <> 0 Then
      --消费卡
      Zl_病人卡结算记录_支付(n_结算卡序号, v_支付卡号, n_消费卡id, n_结算总额, n_预交id, v_操作员编号, v_操作员姓名, d_登记时间);

    End If;
    If Nvl(n_校对标志, 0) = 0 Then
      --完成时,需要更新人员缴款数据
      For c_缴款 In (Select 结算方式, 操作员姓名, Sum(Nvl(a.冲预交, 0)) As 冲预交
                   From 病人预交记录 A
                   Where a.结帐id = Nvl(n_结帐id, 0) And Mod(a.记录性质, 10) <> 1
                   Group By 结算方式, 操作员姓名) Loop

        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) + Nvl(c_缴款.冲预交, 0)
        Where 收款员 = c_缴款.操作员姓名 And 性质 = 1 And 结算方式 = c_缴款.结算方式;
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额
            (收款员, 结算方式, 性质, 余额)
          Values
            (c_缴款.操作员姓名, c_缴款.结算方式, 1, Nvl(c_缴款.冲预交, 0));
        End If;
      End Loop;
    End If;
  End If;

  --4.获取预交信息
  n_预交id := Null;
  o_Json   := PLJson();
  o_Json   := j_Json.Get_Pljson('deposit_info');
  If Not o_Json Is Null Then
    v_预交单号     := o_Json.Get_String('deposit_no');
    v_发票号       := o_Json.Get_String('fact_no');
    n_预交类别     := Nvl(o_Json.Get_Number('deposit_type'), 2);
    n_主页id       := o_Json.Get_Number('pati_pageid');
    n_缴款科室id   := o_Json.Get_Number('dept_id');
    n_缴款金额     := o_Json.Get_Number('money');
    v_缴款单位     := o_Json.Get_String('emp_name');
    v_单位开户行   := o_Json.Get_String('emp_bank_name');
    v_开户行账号   := o_Json.Get_String('emp_bank_actno');
    v_摘要         := o_Json.Get_String('memo');
    n_领用id       := o_Json.Get_Number('recv_id');
    n_预交id       := o_Json.Get_Number('deposit_id');
    n_预交电子票据 := o_Json.Get_Number('start_einv');

    If Nvl(n_预交id, 0) = 0 Then
      Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
    End If;
    If Nvl(n_关联交易id, 0) = 0 Then
      n_关联交易id := n_预交id;
    End If;
    If n_预交电子票据 Is Null Then
      n_预交电子票据 := Zl_Fun_Isstarteinvoice(2, 0, 1, n_预交类别);
    End If;

    --操作状态_In:0-正常结算，1-保存为异常单据或未生效的单据，2-完成异常结算
    Zl_病人预交记录_Insert_s(n_预交id, v_预交单号, v_发票号, n_病人id, n_主页id, v_病人姓名, v_性别, v_年龄, n_门诊号, n_住院号, v_付款方式名称, n_缴款科室id,
                       n_缴款金额, v_结算方式, v_结算号码, v_缴款单位, v_单位开户行, v_开户行账号, v_摘要, v_操作员编号, v_操作员姓名, n_领用id, n_预交类别, n_卡类别id,
                       n_结算卡序号, v_支付卡号, v_交易流水号, v_交易说明, v_合作单位, d_登记时间, n_结帐id, Null, Nvl(n_更新预交余额, 0), Nvl(n_费用状态, 0),
                       n_关联交易id, Null, Nvl(n_预交电子票据, 0));
    n_结算总额 := Nvl(n_结算总额, 0) + Nvl(n_缴款金额, 0);

  End If;
  If Nvl(n_结算总额, 0) <> Nvl(n_本次结算总计, 0) Then
    --本次传入的总额与结算总额不一致
    v_Err_Msg := '费用结算金额不正确,请检查!';
    Raise Err_Item;
  End If;

  Json_Out := '{"output":{"code":1,"message":"' || '成功' || '","deposit_id":' || Nvl(n_预交id, 0) || ',"balance_id":' ||
              Nvl(n_结帐id, 0) || '}}';
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Addcardfeeinfo;
/

Create Or Replace Procedure Zl_Exsesvr_Delcardfeeinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:对卡费及预交款进行销帐处理
  --入参：Json_In:格式
  --  input
  --    oper_fun  N 1 操作状态:0-正常的预交款或卡费的退款记录;1-保存为异常的退款记录;2-作废异常数据;3-删除所有异常记录
  --    cardfee_no  C 1 卡费对应的费用单据号
  --    deposit_no  C 1 预交单据号
  --    cardfee_sign  N 1 是否退卡费:1-是退卡费;0-不退卡费
  --    mrbkfee_sign N 1 是否退病历费:1-退病历费;0-不退病历费
  --    operator_name C 1 操作员姓名
  --    operator_code C 1 操作员编号
  --    del_time  C 1 退费时间:yyyy-mm-dd hh:mi:ss
  --    balance_info  C   只存在一条数据
  --      moeny N 1 退款金额
  --      blnc_mode C 1 结算方式
  --      blnc_no C 1 结算号码
  --      memo  C 1 摘要
  --      cardtype_id N 1 卡类别id
  --      consumer_no N 1 结算卡序号：即卡消费接口目录.编号
  --      consume_card_id N 1 消费卡ID
  --      cardno  C 1 卡号
  --      swapno  C 1 交易流水号
  --      swapmemo  C 1 交易说明
  --      cprtion_unit  C 1 合作单位
  --      relation_id N 1 关联交易ID
  --出参: Json_Out,格式如下
  -- output
  --  code  C 1 应答码：0-失败；1-成功
  --  message C 1 应答消息：失败时返回具体的错误信息
  --  deposit_id  N 1 预交ID:返回冲预交ID
  --  balance_id N 1 结帐ID：返回冲销ID

  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  o_Json PLJson;

  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  n_操作状态 Number(2);
  v_卡费单号 住院费用记录.No%Type;
  v_No       住院费用记录.No%Type;
  v_预交单号 病人预交记录.No%Type;
  v_划价单   住院费用记录.No%Type;
  n_记帐费用 住院费用记录.记帐费用%Type;

  n_是否退卡费   Number(2);
  n_是否退病历费 Number(2);
  v_操作员姓名   住院费用记录.操作员姓名%Type;
  v_操作员编号   住院费用记录.操作员编号%Type;

  n_病人id   住院费用记录.病人id%Type;
  d_退费时间 Date;

  n_退款金额 住院费用记录.实收金额%Type;

  v_结算方式   病人预交记录.结算方式%Type;
  v_结算号码   病人预交记录.结算号码%Type;
  n_卡类别id   病人预交记录.卡类别id%Type;
  n_结算卡序号 病人预交记录.结算卡序号%Type;
  v_卡号       病人预交记录.卡号%Type;
  v_交易流水号 病人预交记录.交易流水号%Type;
  v_交易说明   病人预交记录.交易说明%Type;
  v_合作单位   病人预交记录.合作单位%Type;
  n_关联交易id 病人预交记录.关联交易id%Type;

  n_退款合计 住院费用记录.实收金额%Type;

  n_冲销id   病人预交记录.结帐id%Type;
  n_原结帐id 病人预交记录.结帐id%Type;

  n_组id     病人预交记录.缴款组id%Type;
  n_Count    Number(18);
  n_返回值   住院费用记录.实收金额%Type;
  n_冲预交   病人预交记录.冲预交%Type;
  n_充值金额 病人预交记录.冲预交%Type;
  v_结算摘要 病人预交记录.摘要%Type;
  n_预交id   病人预交记录.Id%Type;
  n_原预交id 病人预交记录.Id%Type;

  n_门诊号       病人预交记录.门诊号%Type;
  n_住院号       病人预交记录.住院号%Type;
  v_付款方式名称 病人预交记录.付款方式名称%Type;
  n_是否电子票据 病人预交记录.是否电子票据%Type;
  v_Output       Varchar2(32767);
Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_操作状态 := Nvl(j_Json.Get_Number('oper_fun'), 0);

  v_卡费单号     := j_Json.Get_String('cardfee_no');
  v_预交单号     := j_Json.Get_String('deposit_no');
  n_是否退卡费   := j_Json.Get_Number('cardfee_sign');
  n_是否退病历费 := j_Json.Get_Number('mrbkfee_sign');
  v_操作员姓名   := j_Json.Get_String('operator_name');
  v_操作员编号   := j_Json.Get_String('operator_code');
  d_退费时间     := To_Date(j_Json.Get_String('del_time'), 'YYYY-MM-DD hh24:mi:ss');

  If Nvl(n_操作状态, 0) = 3 Then
    --删除异常记录
  
    Delete 病人预交记录
    Where 结帐id In (Select 结帐id
                   From 住院费用记录
                   Where NO = v_卡费单号 And 记录性质 = 5 And 记录状态 = 1 And Nvl(费用状态, 0) = 1) And Mod(记录性质, 10) <> 1 And
          Nvl(校对标志, 0) = 1;
  
    If Sql%NotFound Then
      v_Err_Msg := '单据可能因并发原因被他人删除或已经结算，不允许再进行删除操作!';
      Raise Err_Item;
    End If;
  
    Delete 住院费用记录 Where 记录性质 = 5 And 记录状态 = 1 And Nvl(费用状态, 0) = 1 And NO = v_卡费单号;
  
    --删除预交记录
    If v_预交单号 Is Not Null Then
      Delete 病人预交记录 Where NO = v_预交单号 And 记录性质 = 1 And Nvl(校对标志, 0) = 1;
      If Sql%NotFound Then
        v_Err_Msg := '单据可能因并发原因被他人删除或已经结算，不允许再进行删除操作!';
        Raise Err_Item;
      End If;
    End If;
  
    zlJsonPutValue(v_Output, 'code', 1, 1, 1);
    zlJsonPutValue(v_Output, 'message', '成功');
    zlJsonPutValue(v_Output, 'deposit_id', 0, 1);
    zlJsonPutValue(v_Output, 'balance_id', 0, 1, 2);
    Json_Out := '{"output":' || v_Output || '}';
  
    Return;
  End If;
  n_组id := Zl_Get组id(v_操作员姓名);

  If d_退费时间 Is Null Then
    d_退费时间 := Sysdate;
  End If;

  Select Max(NO), Nvl(Max(记帐费用), 0), -1 * Sum(实收金额), Max(摘要), Max(病人id), Max(结帐id)
  Into v_No, n_记帐费用, n_退款合计, v_划价单, n_病人id, n_原结帐id
  From 住院费用记录
  Where NO = v_卡费单号 And 记录性质 = 5 And 记录状态 = 1 And
        (Nvl(n_是否退卡费, 0) = 1 And Nvl(附加标志, 0) <> 8 Or Nvl(n_是否退病历费, 0) = 1 And Nvl(附加标志, 0) = 8);

  If v_No Is Null Then
    v_Err_Msg := '单据为' || v_卡费单号 || '不存在,可能该单据因并发原因被他人销帐或退费,不允许再进行销帐或退费处理!';
    Raise Err_Item;
  End If;

  --1.产生销帐费用记录
  n_冲销id := Null;
  If n_记帐费用 = 0 Then
    Select 病人结帐记录_Id.Nextval Into n_冲销id From Dual;
  End If;

  Insert Into 住院费用记录
    (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 病人id, 主页id, 病人病区id, 病人科室id, 门诊标志, 标识号, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位, 付数, 数次,
     加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 操作员编号, 操作员姓名, 发生时间, 登记时间, 结帐id, 结帐金额,
     缴款组id, 结论, 摘要, 费用状态)
    Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 病人id, 主页id, 病人病区id, 病人科室id, 门诊标志, 标识号, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id,
           计算单位, 付数, -数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, -应收金额, -实收金额, 开单部门id, 开单人, 执行部门id, v_操作员编号,
           v_操作员姓名, 发生时间, d_退费时间, n_冲销id, Decode(n_冲销id, Null, Null, -结帐金额), n_组id, 结论, 摘要,
           Decode(Nvl(n_操作状态, 0), 0, Null, 1)
    From 住院费用记录
    Where NO = v_卡费单号 And 记录性质 = 5 And 记录状态 = 1 And
          (Nvl(n_是否退卡费, 0) = 1 And Nvl(附加标志, 0) <> 8 Or Nvl(n_是否退病历费, 0) = 1 And Nvl(附加标志, 0) = 8);

  Update 住院费用记录
  Set 记录状态 = 3
  Where NO = v_卡费单号 And 记录性质 = 5 And 记录状态 = 1 And
        (Nvl(n_是否退卡费, 0) = 1 And Nvl(附加标志, 0) <> 8 Or Nvl(n_是否退病历费, 0) = 1 And Nvl(附加标志, 0) = 8);

  --处理发卡划价单，如果划价还未收费，直接删除
  If Not v_划价单 Is Null Then
    Select Count(1)
    Into n_Count
    From 门诊费用记录
    Where 病人id = n_病人id And 记录性质 = 1 And NO = v_划价单 And 记录状态 = 0;
  
    If n_Count <> 0 Then
      Zl_门诊划价记录_Delete_s(v_划价单);
    End If;
  End If;

  If Nvl(n_记帐费用, 0) = 1 Then
    --记帐单需要处理费用余额
    For c_退费 In (Select a.No, 序号, a.病人id, a.主页id, a.收费细目id, a.收入项目id, a.病人病区id, a.开单部门id, a.执行部门id, a.病人科室id,
                        Nvl(a.实收金额, 0) As 实收金额, Nvl(a.结帐金额, 0) As 结帐金额
                 From 住院费用记录 A
                 Where NO = v_卡费单号 And 记录性质 = 5 And 记录状态 = 2 And 登记时间 = d_退费时间) Loop
    
      Update 病人余额
      Set 费用余额 = Nvl(费用余额, 0) + c_退费.实收金额
      Where 性质 = 1 And 病人id = c_退费.病人id And Nvl(类型, 2) = Decode(Nvl(c_退费.主页id, 0), 0, 1, 2)
      Returning 费用余额 Into n_返回值;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 性质, 类型, 预交余额, 费用余额)
        Values
          (c_退费.病人id, 1, Decode(Nvl(c_退费.主页id, 0), 0, 1, 2), 0, c_退费.实收金额);
        n_返回值 := c_退费.实收金额;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete 病人余额 Where 性质 = 1 And 病人id = c_退费.病人id And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
      End If;
    
      --汇总'病人未结费用'
      Update 病人未结费用
      Set 金额 = Nvl(金额, 0) + c_退费.实收金额
      Where 病人id = Nvl(c_退费.病人id, 0) And Nvl(主页id, 0) = Nvl(c_退费.主页id, 0) And Nvl(病人病区id, 0) = Nvl(c_退费.病人病区id, 0) And
            Nvl(病人科室id, 0) = Nvl(c_退费.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(c_退费.开单部门id, 0) And
            Nvl(执行部门id, 0) = Nvl(c_退费.执行部门id, 0) And 收入项目id + 0 = Nvl(c_退费.收入项目id, 0) And 来源途径 = 3;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人未结费用
          (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
        Values
          (c_退费.病人id, Decode(c_退费.主页id, 0, Null, c_退费.主页id), Decode(c_退费.病人病区id, 0, Null, c_退费.病人病区id),
           Decode(c_退费.病人科室id, 0, Null, c_退费.病人科室id), Decode(c_退费.开单部门id, 0, Null, c_退费.开单部门id),
           Decode(c_退费.执行部门id, 0, Null, c_退费.执行部门id), c_退费.收入项目id, 3, c_退费.实收金额);
      End If;
    End Loop;
  End If;
  If n_操作状态 = 2 Then
  
    --作废异常
    If Nvl(n_记帐费用, 0) <> 1 Then
    
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 姓名, 性别, 年龄, 门诊号, 住院号, 付款方式名称, 科室id, 摘要, 金额, 结算方式, 结算号码, 收款时间, 操作员编号,
         操作员姓名, 缴款单位, 单位开户行, 单位帐号, 缴款组id, 预交类别, 卡类别id, 卡号, 交易流水号, 交易说明, 合作单位, 结算卡序号, 校对标志, 关联交易id, 交易时间, 交易人员, 结帐id, 冲预交,
         结算性质, 是否电子票据)
      
        Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 姓名, 性别, 年龄, 门诊号, 住院号, 付款方式名称, 科室id, 摘要, -1 * 金额, 结算方式,
               结算号码, 收款时间, 操作员编号, 操作员姓名, 缴款单位, 单位开户行, 单位帐号, 缴款组id, 预交类别, 卡类别id, 卡号, 交易流水号, 交易说明, 合作单位, 结算卡序号, Null,
               关联交易id, 交易时间, 交易人员, n_冲销id, -1 * 冲预交, 结算性质, 是否电子票据
        From 病人预交记录
        Where 结帐id = n_原结帐id;
    
      Update 病人预交记录 Set 记录状态 = 3 Where 结帐id = n_原结帐id And Mod(记录性质, 10) <> 1;
    End If;
  
    --作废异常数据
    If v_预交单号 Is Not Null Then
      --操作_In:0-删除异常充值单据，1-删除异常退款单据，2-删除异常余额退款单据
      Zl_病人预交异常记录_Delete(v_预交单号, 0);
    End If;
  
    zlJsonPutValue(v_Output, 'code', 1, 1, 1);
    zlJsonPutValue(v_Output, 'message', '成功');
    zlJsonPutValue(v_Output, 'deposit_id', Nvl(n_预交id, 0), 1);
    zlJsonPutValue(v_Output, 'balance_id', Nvl(n_冲销id, 0), 1, 2);
    Json_Out := '{"output":' || v_Output || '}';
  
    Return;
  End If;

  If n_操作状态 = 0 And Nvl(n_是否退卡费, 0) = 1 Then
    --回收票据处理
    Zl_病人医疗卡票据_Update_s(4, '', v_操作员姓名, d_退费时间, v_卡费单号, Null, Null, Null);
  End If;
  --结算信息
  o_Json := j_Json.Get_Pljson('balance_info');
  If o_Json Is Not Null Then
  
    n_退款金额   := o_Json.Get_Number('moeny');
    v_结算方式   := o_Json.Get_String('blnc_mode');
    v_结算号码   := o_Json.Get_String('blnc_no');
    n_卡类别id   := o_Json.Get_Number('cardtype_id');
    n_结算卡序号 := o_Json.Get_Number('consumer_no');
    v_卡号       := o_Json.Get_String('cardno');
    v_交易流水号 := o_Json.Get_String('swapno');
    v_交易说明   := o_Json.Get_String('swapmemo');
    v_合作单位   := o_Json.Get_String('cprtion_unit');
    n_关联交易id := o_Json.Get_Number('relation_id');
    v_结算摘要   := o_Json.Get_String('memo');
  
    If Nvl(n_卡类别id, 0) = 0 Then
      n_卡类别id := Null;
    End If;
    If Nvl(n_结算卡序号, 0) = 0 Then
      n_结算卡序号 := Null;
    End If;
  
    --1.处理销帐费用
    If n_记帐费用 = 0 Then
      --非记帐费用及非异常状态，都需要处理销帐数据
      For c_费用 In (Select NO, Max(病人id) As 病人id, Max(主页id) As 主页id, Max(姓名) As 姓名, Max(性别) As 性别, Max(年龄) As 年龄,
                          Max(病人科室id) As 病人科室id, Sum(结帐金额) As 结帐金额
                   From 住院费用记录
                   Where 结帐id = n_冲销id
                   Group By NO) Loop
        Select Max(门诊号), Max(住院号), Max(付款方式名称), Max(是否电子票据)
        Into n_门诊号, n_住院号, v_付款方式名称, n_是否电子票据
        From 病人预交记录
        Where 结帐id In (Select 结帐id From 住院费用记录 Where NO = c_费用.No And 记录性质 = 5 And 记录状态 In (0, 1, 3));
      
        Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
        If Nvl(n_关联交易id, 0) = 0 Then
          n_关联交易id := n_预交id;
        End If;
      
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 姓名, 性别, 年龄, 门诊号, 住院号, 付款方式名称, 科室id, 摘要, 结帐id, 冲预交, 结算方式, 结算号码, 收款时间,
           操作员编号, 操作员姓名, 缴款单位, 单位开户行, 单位帐号, 缴款组id, 预交类别, 卡类别id, 卡号, 交易流水号, 交易说明, 合作单位, 结算卡序号, 校对标志, 关联交易id, 交易时间, 交易人员,
           结算性质, 是否电子票据)
          Select n_预交id, c_费用.No, '' As 实际票号, 5, 2, c_费用.病人id, c_费用.主页id, c_费用.姓名, c_费用.性别, c_费用.年龄, n_门诊号, n_住院号,
                 v_付款方式名称, c_费用.病人科室id, v_结算摘要, n_冲销id, c_费用.结帐金额, v_结算方式, v_结算号码, d_退费时间, v_操作员编号, v_操作员姓名, '' As 缴款单位,
                 '' As 单位开户行, '' As 单位帐号, n_组id, Null, n_卡类别id, v_卡号, v_交易流水号, v_交易说明, v_合作单位, n_结算卡序号,
                 Decode(Nvl(n_操作状态, 0), 0, Null, 1), n_关联交易id, d_退费时间, v_操作员姓名, 5, n_是否电子票据
          From Dual;
      
        If n_操作状态 = 0 Then
          If Nvl(n_结算卡序号, 0) <> 0 Then
            Begin
              Select b.Id
              Into n_原预交id
              From 病人预交记录 B
              Where b.结帐id = n_原结帐id And b.结算卡序号 = n_结算卡序号;
            Exception
              When Others Then
                n_原预交id := -1;
            End;
          
            If n_原预交id = -1 Then
            
              v_Err_Msg := '没有发现' || v_结算方式 || '的原结算数据！';
              Raise Err_Item;
            
            End If;
            Zl_病人卡结算记录_退款(n_结算卡序号, v_卡号, Null, -1 * c_费用.结帐金额, n_原预交id, n_预交id, v_操作员编号, v_操作员姓名, d_退费时间);
          End If;
          Update 人员缴款余额
          Set 余额 = Nvl(余额, 0) + c_费用.结帐金额
          Where 性质 = 1 And 收款员 = v_操作员姓名 And 结算方式 = v_结算方式
          Returning 余额 Into n_返回值;
        
          If Sql%RowCount = 0 Then
            Insert Into 人员缴款余额
              (收款员, 结算方式, 性质, 余额)
            Values
              (v_操作员姓名, v_结算方式, 1, c_费用.结帐金额);
            n_返回值 := c_费用.结帐金额;
          
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 人员缴款余额
            Where 收款员 = v_操作员姓名 And 性质 = 1 And 结算方式 = v_结算方式 And Nvl(余额, 0) = 0;
          End If;
        End If;
      End Loop;
      Update 病人预交记录 Set 记录状态 = 3 Where 结帐id = n_原结帐id And 记录性质 = 5;
    End If;
    n_预交id := Null;
  
    --2.处理预交单据
    If v_预交单号 Is Not Null Then
    
      --作废预交记录
      Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
    
      Select Sum(冲预交), Max(Decode(记录性质, 1, ID, 0)), Max(病人id)
      Into n_冲预交, n_原预交id, n_病人id
      From 病人预交记录
      Where Mod(记录性质, 10) = 1 And NO = v_预交单号;
      If n_冲预交 <> 0 Then
        v_Err_Msg := '预交款已经发生消费数据，不允许再进行退款操作!';
        Raise Err_Item;
      End If;
    
      Update 病人预交记录 Set 记录状态 = 3 Where ID = n_原预交id;
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 姓名, 性别, 年龄, 门诊号, 住院号, 付款方式名称, 科室id, 摘要, 金额, 结算方式, 结算号码, 收款时间, 操作员编号,
         操作员姓名, 缴款单位, 单位开户行, 单位帐号, 缴款组id, 预交类别, 卡类别id, 卡号, 交易流水号, 交易说明, 合作单位, 结算卡序号, 校对标志, 关联交易id, 交易时间, 交易人员, 结算性质,
         结帐id, 预交电子票据)
        Select n_预交id, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 姓名, 性别, 年龄, 门诊号, 住院号, 付款方式名称, 科室id, Nvl(v_结算摘要, 摘要) As 摘要, -1 * 金额,
               v_结算方式, v_结算号码, d_退费时间, v_操作员编号, v_操作员姓名, 缴款单位, 单位开户行, 单位帐号, n_组id, 预交类别, n_卡类别id, v_卡号, v_交易流水号, v_交易说明,
               v_合作单位, n_结算卡序号, Decode(Nvl(n_操作状态, 0), 0, Null, 1), 关联交易id, d_退费时间, v_操作员姓名, 结算性质,
               Decode(Nvl(n_冲销id, 0), 0, Null, n_冲销id) As 结帐id, 预交电子票据
        From 病人预交记录
        Where ID = n_原预交id;
    
      Update 病人预交记录 Set 记录状态 = 3 Where ID = n_原预交id;
    
      If Nvl(n_操作状态, 0) = 0 Then
        --使用消费卡，不会产生预交款，因此不处理其他的
        For c_预交 In (Select ID, NO, 金额, 结算方式, 病人id, 预交类别 From 病人预交记录 Where ID = n_预交id) Loop
          --需要更新余额
          n_充值金额 := Nvl(c_预交.金额, 0);
        
          Update 病人余额
          Set 预交余额 = Nvl(预交余额, 0) + Nvl(c_预交.金额, 0)
          Where 性质 = 1 And 病人id = Nvl(c_预交.病人id, 0) And Nvl(类型, 2) = Nvl(c_预交.预交类别, 2)
          Returning 预交余额 Into n_返回值;
        
          If Sql%RowCount = 0 Then
            Insert Into 病人余额
              (病人id, 性质, 类型, 预交余额, 费用余额)
            Values
              (c_预交.病人id, 1, Nvl(c_预交.预交类别, 2), c_预交.金额, 0);
            n_返回值 := c_预交.金额;
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 病人余额
            Where 病人id = c_预交.病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
          End If;
        
          --预交单据余额
          Update 预交单据余额
          Set 预交余额 = Nvl(预交余额, 0) + c_预交.金额
          Where 预交id = n_原预交id
          Returning 预交余额 Into n_返回值;
          If Sql%RowCount = 0 Then
            Insert Into 预交单据余额
              (预交id, 病人id, 预交类别, 预交余额)
            Values
              (n_原预交id, c_预交.病人id, Nvl(c_预交.预交类别, 2), c_预交.金额);
            n_返回值 := c_预交.金额;
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 预交单据余额
            Where 预交id = n_原预交id And Nvl(预交类别, 2) = Nvl(c_预交.预交类别, 2) And Nvl(预交余额, 0) = 0;
          End If;
        
          --需要更新人员缴款余额
        
          Update 人员缴款余额
          Set 余额 = Nvl(余额, 0) + c_预交.金额
          Where 性质 = 1 And 收款员 = v_操作员姓名 And 结算方式 = c_预交.结算方式
          Returning 余额 Into n_返回值;
        
          If Sql%RowCount = 0 Then
            Insert Into 人员缴款余额
              (收款员, 结算方式, 性质, 余额)
            Values
              (v_操作员姓名, c_预交.结算方式, 1, c_预交.金额);
            n_返回值 := c_预交.金额;
          
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 人员缴款余额
            Where 收款员 = v_操作员姓名 And 性质 = 1 And 结算方式 = c_预交.结算方式 And Nvl(余额, 0) = 0;
          End If;
        
        End Loop;
      Else
        Select Sum(金额) Into n_充值金额 From 病人预交记录 Where ID = n_预交id;
      End If;
    
    End If;
  
    If Nvl(n_退款合计, 0) + Nvl(n_充值金额, 0) <> Nvl(n_退款金额, 0) Then
      v_Err_Msg := '当前退款金额(' || Trim(To_Char(Nvl(n_退款金额, 0), '9999999999.999')) || ')与本次销帐金额(' ||
                   Trim(To_Char(Nvl(n_退款合计, 0) + Nvl(n_充值金额, 0), '99999999999.99')) || ')不正确，请检查！';
      Raise Err_Item;
    End If;
  End If;

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '成功');
  zlJsonPutValue(v_Output, 'deposit_id', Nvl(n_预交id, 0), 1);
  zlJsonPutValue(v_Output, 'balance_id', Nvl(n_冲销id, 0), 1, 2);
  Json_Out := '{"output":' || v_Output || '}';

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Delcardfeeinfo;
/

Create Or Replace Procedure Zl_Exsesvr_Delcardfeecheck
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  ---------------------------------------------------------------------------
  --功能:检查退卡及病历费数据是否合法
  --入参      json
  --input
  --  cardfee_no  C 1 卡费单号
  --  deposit_no  C 1 预交单据号
  --  reretruned  N 1 是否异常重退:1-是异常重退;0-非异常重退
  --  delfee_sign N 1 退费标志：0-仅退卡费;1-仅退病历费;2-病历费及卡费
  --  balance_info  C   退款方式
  --    delmoney  N 1 本次退款金额
  --    pay_mode  C 1 结算方式
  --    cardtype_id N 1 卡类别id
  --    consumer_no N 1 结算卡序号，即卡消费接口目录.编号
  --    must_allreturn  N 1 是否全退:1-必须全退;0-允许部分退
  --出参      json
  --output
  --  code                      C 1 应答码：0-失败；1-成功
  --  message                   C 1 应答消息：失败时返回具体的错误信息
  --  tip_list[]  C 1 提示列表:主要是可能存在多个提示询问方式，所以用列表,禁止时，返回一条信息
  --    tip_mode  C 1 控制方式:1-提示询问;2-禁止
  --    tip_message C 1 提示信息
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  o_Json     PLJson;
  v_卡费单号 住院费用记录.No%Type;
  -- v_预交单据号   病人预交记录.No%Type;
  n_是否异常重退 Number(2);
  n_退费标志     Number(2);
  n_结账控制     Number(2);
  n_结帐id       Number(18);
  n_冲销id       Number(18);
  v_No           住院费用记录.No%Type;
  n_费用状态     Number(5);

  n_本次退款金额 Number(16, 5);
  -- v_结算方式     Varchar2(100);
  n_卡类别id   Number(18);
  n_结算卡序号 Number(18);
  n_是否全退   Number(18);
  n_Count      Number(18);
  n_记帐费用   Number(2);
  n_冲预交     Number(16, 5);

  v_Err_Msg    Varchar2(1000);
  n_关联交易id Number(16, 5);
  n_实收金额   住院费用记录.实收金额%Type;
  n_应收金额   住院费用记录.实收金额%Type;
  n_结帐金额   住院费用记录.实收金额%Type;
  n_充值金额   病人预交记录.金额%Type;
  Function Get_Success_Message
  (
    Tip_Mod_In     Integer,
    Tip_Message_In Varchar2
  ) Return Clob Is
  
  Begin
    Return '{"output":{"code":1,"message":"成功"' || ',"tip_list":[{"tip_mode":' || Nvl(Tip_Mod_In, 0) || ',"tip_Message":"' || Zltools.Zljsonstr(Tip_Message_In) || '"}]}}';
  End Get_Success_Message;

Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_卡费单号 := j_Json.Get_String('cardfee_no');
  --v_预交单据号   := j_Json.Get_String('deposit_no');
  n_是否异常重退 := j_Json.Get_Number('reretruned');
  n_退费标志     := Nvl(j_Json.Get_Number('delfee_sign'), 0);

  If v_卡费单号 Is Null Then
    v_Err_Msg := '未传入有效的费用单据号';
    Json_Out  := zlJsonOut(v_Err_Msg);
    Return;
  End If;

  Select Max(记帐费用), Max(费用状态), Max(结帐id)
  Into n_记帐费用, n_费用状态, n_结帐id
  From 住院费用记录
  Where NO = v_卡费单号 And 记录性质 = 5 And 记录状态 In (0, 1, 3) And Rownum < 2;

  If Nvl(n_是否异常重退, 0) = 1 Then
    If Nvl(n_记帐费用, 0) <> 1 Then
      --0-仅退卡费;1-仅退病历费;2-病历费及卡费
      If Nvl(n_退费标志, 0) = 0 Then
        Select Max(结帐id)
        Into n_冲销id
        From 住院费用记录
        Where NO = v_卡费单号 And 记录性质 = 5 And Nvl(附加标志, 0) <> 8 And 记录状态 = 2;
      Elsif Nvl(n_退费标志, 0) = 1 Then
        Select Max(结帐id)
        Into n_冲销id
        From 住院费用记录
        Where NO = v_卡费单号 And 记录性质 = 5 And Nvl(附加标志, 0) = 8 And 记录状态 = 2;
      Else
        Select Max(结帐id) Into n_冲销id From 住院费用记录 Where NO = v_卡费单号 And 记录性质 = 5 And 记录状态 = 2;
      End If;
    
      Select Max(1) Into n_Count From 病人预交记录 Where 结帐id = n_冲销id And Nvl(校对标志, 0) <> 0;
      If n_Count = 0 Then
        v_Err_Msg := '该单据可能已被他人作废或重退';
        Json_Out  := zlJsonOut(v_Err_Msg);
        Return;
      End If;
    End If;
    Json_Out := Get_Success_Message(0, '');
    Return;
  
  End If;

  If Nvl(n_费用状态, 0) = 1 Then
    --当前为异常单据
    v_Err_Msg := '单据【' || v_卡费单号 || '】为异常单据';
    Json_Out  := zlJsonOut(v_Err_Msg);
    Return;
  End If;

  o_Json := j_Json.Get_Pljson('balance_info');
  If o_Json Is Not Null Then
    n_本次退款金额 := o_Json.Get_Number('delmoney');
    --v_结算方式     := o_Json.Get_String('pay_mode');
    n_卡类别id   := o_Json.Get_Number('cardtype_id');
    n_结算卡序号 := o_Json.Get_Number('consumer_no');
    n_是否全退   := Nvl(o_Json.Get_Number('must_allreturn'), 0);
  
    If Nvl(n_是否全退, 0) = 1 Then
      --必须全退，则需要检查所有结算是否全退
      Select Max(关联交易id), Max(Decode(记录性质, 1, 1, 0))
      Into n_关联交易id, n_Count
      From 病人预交记录
      Where 结帐id = n_结帐id;
    
      --检查是否消费
      If Nvl(n_Count, 0) = 1 Then
        Select Sum(冲预交), Sum(Decode(记录性质, 1, 1, 0) * 金额)
        Into n_冲预交, n_充值金额
        From 病人预交记录
        Where 关联交易id = n_关联交易id And Mod(记录性质, 10) = 1;
        If Nvl(n_冲预交, 0) <> 0 Then
          v_Err_Msg := '发卡充值金额已经发生消费，当前结算又必须全退';
          Json_Out  := zlJsonOut(v_Err_Msg);
          Return;
        End If;
      End If;
    
      Select Sum(冲预交) + Nvl(n_充值金额, 0) Into n_冲预交 From 病人预交记录 Where 结帐id = n_结帐id;
    
      If Nvl(n_冲预交, 0) <> Nvl(n_本次退款金额, 0) Then
        v_Err_Msg := '本次结算(' || Nvl(n_冲预交, 0) || ')与当前退款金额(' || Nvl(n_本次退款金额, 0) || ')不符，必须全退';
        Json_Out  := zlJsonOut(v_Err_Msg);
        Return;
      End If;
    
      For c_预交 In (Select 卡类别id, 结算卡序号 From 病人预交记录 Where 结帐id = n_结帐id) Loop
        If Nvl(n_卡类别id, 0) <> Nvl(c_预交.卡类别id, 0) And n_卡类别id <> 0 Then
          v_Err_Msg := '不能使用其他三方卡进行退款';
          Json_Out  := zlJsonOut(v_Err_Msg);
          Return;
        Elsif Nvl(c_预交.结算卡序号, 0) <> 0 And Nvl(n_结算卡序号, 0) <> Nvl(c_预交.结算卡序号, 0) Then
          v_Err_Msg := '不能使用其他三方卡进行退款';
          Json_Out  := zlJsonOut(v_Err_Msg);
          Return;
        End If;
      End Loop;
    
    End If;
  End If;

  If Nvl(n_记帐费用, 0) = 1 Then
    --已结帐单据操作控制
    n_结账控制 := To_Number(Nvl(zl_GetSysParameter('已结帐单据操作'), '0'));
    --0-允许 1-提示 2-禁止7
    If Nvl(n_结账控制, 0) <> 0 Then
      --退费标志：0-仅退卡费;1-仅退病历费;2-病历费及卡费
      Select Max(NO), Nvl(Sum(实收金额), 0), Nvl(Sum(应收金额), 0), Nvl(Sum(结帐金额), 0), Nvl(Max(结帐id), 0)
      Into v_No, n_实收金额, n_应收金额, n_结帐金额, n_结帐id
      From 住院费用记录
      Where Mod(记录性质, 10) = 5 And 记帐费用 = 1 And NO = v_卡费单号 And
            ((n_退费标志 = 0 And Nvl(附加标志, 0) <> 8) Or (n_退费标志 = 1 And Nvl(附加标志, 0) = 8) Or n_退费标志 = 2);
    
      If v_No Is Not Null Then
        If (n_实收金额 - n_结帐金额 = 0 And n_应收金额 <> 0) Or (n_实收金额 = 0 And n_结帐金额 = 0 And n_结帐id <> 0) Then
          --肯定结帐
          v_Err_Msg := '记帐单【' || v_No || '】已经结帐';
          Json_Out  := Get_Success_Message(n_结账控制, v_Err_Msg);
          Return;
        End If;
      End If;
    End If;
  Else
    --退费标志：0-仅退卡费;1-仅退病历费;2-病历费及卡费
    Select Count(1)
    Into n_Count
    From 住院费用记录
    Where NO = v_卡费单号 And 记录状态 = 1 And 记录性质 = 5 And
          ((n_退费标志 = 0 And Nvl(附加标志, 0) <> 8) Or (n_退费标志 = 1 And Nvl(附加标志, 0) = 8) Or n_退费标志 = 2);
    If Nvl(n_Count, 0) = 0 Then
      If n_退费标志 = 1 Then
        Json_Out := Get_Success_Message(2, '当前病历费已被其他人员退费');
      Else
        Json_Out := Get_Success_Message(2, '当前卡费已被其他人员退费');
      End If;
      Return;
    End If;
  End If;
  Json_Out := Get_Success_Message(0, '');
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Delcardfeecheck;
/


Create Or Replace Procedure Zl_Exsesvr_Checkcardnoisused
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:检查卡号是否使用，使用后返回领用ID
  --入参：Json_In:格式
  --   input      
  --    cardtype_id  N  1  卡类别id
  --    cardno  C  1  卡号
  --出参: Json_Out,格式如下
  --  output      
  --    code  C  1  应答码：0-失败；1-成功
  --    message  C  1  应答消息：失败时返回具体的错误信息
  --    isexsit  N  1  是否存在:1-存在;0-不存在
  --    recv_id  N  1  领用id

  ---------------------------------------------------------------------------
  j_Input    PLJson;
  j_Json     PLJson;
  n_卡类别id Number(18);
  v_卡号     Varchar2(100);
  v_Output   Varchar2(32767);

  n_领用id 票据领用记录.Id%Type;

  n_Exist Number(2);
Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_卡类别id := j_Json.Get_Number('cardtype_id');
  v_卡号     := j_Json.Get_String('cardno');

  Select Max(b.领用id), Max(1)
  Into n_领用id, n_Exist
  From 票据领用记录 A, 票据使用明细 B
  Where a.Id = b.领用id And a.票种 = 5 And (Nvl(a.使用类别, 'LXH') = To_Char(n_卡类别id) Or a.使用类别 Is Null) And b.号码 = v_卡号;

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '成功');
  zlJsonPutValue(v_Output, 'isexist', Nvl(n_Exist, 0), 1);
  zlJsonPutValue(v_Output, 'isexist', Nvl(n_领用id, 0), 1, 2);
  Json_Out := '{"output":' || v_Output || '}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkcardnoisused;
/

Create Or Replace Procedure Zl_Exsesvr_Updcardfeeblncinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:修正改费及卡费同步缴预交的相关信息(发票、预交等信息)
  --入参：Json_In:格式
  --input
  --   oper_fun  N  1  操作状态:0-完成结算;1-支付接口调用前修正;2-支付接口调用后修正
  --   pati_id N 1 病人id
  --   fee_no  C 1 费用单号：发卡所涉及的费用单号
  --   balance_id  N 1 结帐ID
  --   operator_name C 1 操作员姓名
  --   operator_code C 1 操作员编号
  --   create_time C 1 操作时间:yyyy-mm-dd hh:mi:ss
  --   completioned N 1 完成标志: 1-完成结算;0-未完成结算  ,未传入本接点，默认为完成结算
  --   fee_einvoice  N  1  卡费或病历费是否启用电子票据:1-启用;0-不启用
  --   sendcard_info     发卡信息
  --     send_mode N 1 发卡方式;0-发卡,1-补卡,2-换卡
  --     cardtype_id C 1 卡类别id
  --     cardno  C 1 卡号:本次发放或绑定或补卡的卡号
  --     recv_id N 1 领用id:票据领用ID(卡号)
  --     cardno_reusing  N 1 卡号重用:1-卡号允许重复使用用;0-不允许重复使用
  --     cardno_old  C 1 原卡卡号:换卡时，需要传入原卡号
  --   balance_info  C   结算信息
  --     deposit_no  C   预交单号
  --     deposit_id  N   预交ID
  --     deposit_einvoice      预交启用电子票据:1-启用;0-不启用
  --     pay_mode  C 1 结算方式
  --     blnc_no C 1 结算号码
  --     cardtype_id N 1 卡类别id
  --     consumer_no N 1 结算卡序号，即卡消费接口目录.编号
  --     cardno  C 1 卡号
  --     swapno  C 1 交易流水号
  --     swapmemo  C 1 交易说明
  --     memo  C 1 摘要
  --     cprtion_unit  C 1 合作单位
  --     other_list[]  C 1 其他交易信息
  --       swap_name C 1 交易名称
  --       swap_note C 1 交易内容
  --出参: Json_Out,格式如下
  -- output
  --   code                  C 1 应答码：0-失败；1-成功
  --   message               C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  o_Json     PLJson;
  j_Jsonlist Pljson_List := Pljson_List();

  n_发卡方式   Number(2);
  n_病人id     Number(18);
  v_预交单号   病人预交记录.No%Type;
  n_预交id     病人预交记录.Id%Type;
  v_费用单号   门诊费用记录.No%Type;
  n_结帐id     门诊费用记录.结帐id%Type;
  n_领用id     票据领用记录.Id%Type;
  n_卡号重用   Number(2);
  v_原卡号     Varchar2(100);
  v_结算方式   病人预交记录.结算方式%Type;
  v_结算号码   病人预交记录.结算号码%Type;
  n_卡类别id   病人预交记录.卡类别id%Type;
  n_结算卡序号 病人预交记录.结算卡序号%Type;
  v_支付卡号   病人预交记录.卡号%Type;
  v_交易流水号 病人预交记录.交易流水号%Type;
  v_交易说明   病人预交记录.交易说明%Type;
  v_摘要       病人预交记录.摘要%Type;
  v_合作单位   病人预交记录.合作单位%Type;
  n_关联交易id 病人预交记录.关联交易id%Type;
  v_交易名称   三方结算交易.交易项目%Type;
  v_交易内容   三方结算交易.交易内容%Type;

  n_预交金额     病人预交记录.金额%Type;
  n_结算金额     病人预交记录.金额%Type;
  v_卡号         Varchar2(100);
  v_操作员姓名   病人预交记录.操作员姓名%Type;
  v_操作员编号   病人预交记录.操作员编号%Type;
  d_登记时间     病人预交记录.收款时间%Type;
  n_门诊号       病人预交记录.门诊号%Type;
  n_住院号       病人预交记录.住院号%Type;
  v_付款方式名称 病人预交记录.付款方式名称%Type;
  n_Count        Number(18);
  n_返回值       Number(16, 5);
  n_组id         Number(18);
  v_结算摘要     病人预交记录.摘要%Type;
  n_操作状态     Number(5);
  v_原单据号     住院费用记录.No%Type;
  n_是否电子票据 病人预交记录.是否电子票据%Type;
  n_预交电子票据 病人预交记录.是否电子票据%Type;
  n_Temp         Number(2);
  Err_Item Exception;
  v_Err_Msg Varchar2(500);
  l_预交id  t_NumList := t_NumList();

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_费用单号     := j_Json.Get_String('fee_no');
  n_结帐id       := j_Json.Get_Number('balance_id');
  v_操作员姓名   := j_Json.Get_String('operator_name');
  v_操作员编号   := j_Json.Get_String('operator_code');
  d_登记时间     := To_Date(j_Json.Get_String('create_time'), 'YYYY-MM-DD hh24:mi:ss');
  n_病人id       := Nvl(j_Json.Get_Number('pati_id'), 0);
  n_操作状态     := Nvl(j_Json.Get_Number('oper_fun'), 1);
  n_是否电子票据 := Nvl(j_Json.Get_Number('fee_einvoice'), 0);

  If d_登记时间 Is Null Then
    d_登记时间 := Sysdate;
  End If;

  If Nvl(n_病人id, 0) = 0 Then
  
    v_Err_Msg := '不能确定病人信息，请检查！';
    Raise Err_Item;
  End If;
  If v_费用单号 Is Null Then
    v_Err_Msg := '不能确定费用单据信息，请检查！';
    Raise Err_Item;
  End If;

  --读取发卡信息
  o_Json := j_Json.Get_Pljson('sendcard_info');
  If o_Json Is Not Null Then
    n_发卡方式 := Nvl(o_Json.Get_Number('send_mode'), 0); --发卡方式;0-发卡,1-补卡,2-换卡 ,3-退卡
    n_卡类别id := o_Json.Get_Number('cardtype_id');
    n_卡号重用 := Nvl(o_Json.Get_Number('cardno_reusing'), 0);
    v_卡号     := o_Json.Get_String('cardno');
    v_原卡号   := o_Json.Get_String('cardno_old');
    n_领用id   := o_Json.Get_Number('recv_id');
  
    --票据处理
    If Nvl(n_操作状态, 0) = 0 Then
      --完成时，需要处理票据
      If v_卡号 Is Null Then
        v_Err_Msg := '不能确定发卡的卡号信息，请检查！';
        Raise Err_Item;
      End If;
    
      --1-发卡 ;2-换卡;3-补卡 ;4-退卡
      n_Count := Case
                   When n_发卡方式 = 0 Then
                    1
                   When n_发卡方式 = 1 Then
                    3
                   When n_发卡方式 = 2 Then
                    2
                   Else
                    4
                 End;
    
      If n_发卡方式 = 2 Then
        --需要获取原始发卡单据
        Select Max(NO)
        Into v_原单据号
        From 住院费用记录
        Where 记录性质 = 5 And 病人id = n_病人id And 实际票号 = v_原卡号 And To_Number(Nvl(结论, '0')) = Nvl(n_卡类别id, 0) And 附加标志 <> 8;
        If v_原单据号 Is Null Then
          v_Err_Msg := '未找到原始卡号:' || v_原卡号 || '的费用单据 ，不能再进行换卡操作,请检查!';
          Raise Err_Item;
        End If;
        --附加标志:0-发卡，1-补卡，2-换卡
        Update 住院费用记录
        Set 附加标志 = Decode(Nvl(附加标志, 0), 8, 8, n_发卡方式)
        Where 记录性质 = 5 And NO = v_原单据号;
      Else
        v_原单据号 := v_费用单号;
      End If;
    
      Zl_病人医疗卡票据_Update_s(n_Count, v_卡号, v_操作员姓名, d_登记时间, v_原单据号, n_领用id, v_原卡号, n_卡号重用);
    End If;
  
    --非换卡，需要处理
    --如果是换卡且v_费用单号不为NULL时，表示只收病历费，所以不更新实际票号及卡号
    If Nvl(n_结帐id, 0) = 0 Then
      Update 住院费用记录
      Set 操作员姓名 = v_操作员姓名, 操作员编号 = v_操作员编号, 登记时间 = d_登记时间, 实际票号 = Nvl(v_卡号, 实际票号)
      Where Nvl(费用状态, 0) = 1 And NO = v_费用单号 And 记录性质 = 5 And 记录状态 In (1, 3);
    Else
      Update 住院费用记录
      Set 操作员姓名 = v_操作员姓名, 操作员编号 = v_操作员编号, 登记时间 = d_登记时间, 实际票号 = Nvl(v_卡号, 实际票号)
      Where Nvl(费用状态, 0) = 1 And 结帐id = Nvl(n_结帐id, 0);
    End If;
  
    If Sql%NotFound Then
      v_Err_Msg := '可能因并发原因补其他人重收或重退，请检查！';
      Raise Err_Item;
    End If;
  Elsif Nvl(n_操作状态, 0) = 0 Then
    --完成标志=1时，需要处理票据
    Select Count(1) Into n_Count From 住院费用记录 Where 结帐id = Nvl(n_结帐id, 0) And 附加标志 <> 8;
    If n_Count <> 0 Then
      v_Err_Msg := '未传入发卡数据信息，请检查';
      Raise Err_Item;
    End If;
  End If;

  o_Json := PLJson();

  o_Json := j_Json.Get_Pljson('balance_info');
  If o_Json Is Not Null Then
    n_预交电子票据 := Nvl(o_Json.Get_Number('deposit_einvoice'), 0);
    v_预交单号     := o_Json.Get_String('deposit_no');
    n_预交id       := o_Json.Get_Number('deposit_id');
  
    v_结算方式   := o_Json.Get_String('pay_mode');
    v_结算号码   := o_Json.Get_String('blnc_no');
    n_卡类别id   := o_Json.Get_Number('cardtype_id');
    n_结算卡序号 := o_Json.Get_Number('consumer_no');
    v_支付卡号   := o_Json.Get_String('cardno');
    v_交易流水号 := o_Json.Get_String('swapno');
    v_交易说明   := o_Json.Get_String('swapmemo');
    v_摘要       := o_Json.Get_String('memo');
    v_合作单位   := o_Json.Get_String('cprtion_unit');
  
    If Nvl(n_卡类别id, 0) = 0 Then
      n_卡类别id := Null;
    End If;
    If Nvl(n_结算卡序号, 0) = 0 Then
      n_结算卡序号 := Null;
    End If;
  
    If v_预交单号 Is Not Null Then
    
      Update 病人预交记录
      Set 结算方式 = Nvl(v_结算方式, 结算方式), 结算号码 = v_结算号码, 卡类别id = n_卡类别id, 结算卡序号 = Decode(Nvl(n_结算卡序号, 0), 0, Null, 结算卡序号),
          卡号 = v_支付卡号, 交易流水号 = v_交易流水号, 交易说明 = v_交易说明, 摘要 = Nvl(v_摘要, 摘要), 合作单位 = Nvl(v_合作单位, 合作单位), 收款时间 = d_登记时间,
          操作员姓名 = v_操作员姓名, 操作员编号 = v_操作员编号, 校对标志 = Nvl(n_操作状态, 校对标志), 预交电子票据 = Decode(记录性质, 5, Null, Nvl(n_预交电子票据, 0))
      Where ID = n_预交id And (记录状态 = 0 Or 记录状态 = 2);
    
      If Sql%NotFound Then
        v_Err_Msg := '未找到单据号为' || v_预交单号 || '的预交单据 ！';
        Raise Err_Item;
      End If;
    End If;
  
    If v_费用单号 Is Not Null Then
    
      Update 病人预交记录
      Set 结算方式 = Nvl(v_结算方式, 结算方式), 结算号码 = v_结算号码, 卡类别id = n_卡类别id, 结算卡序号 = Decode(Nvl(n_结算卡序号, 0), 0, Null, 结算卡序号),
          卡号 = v_支付卡号, 交易流水号 = v_交易流水号, 交易说明 = v_交易说明, 摘要 = Nvl(v_摘要, 摘要), 合作单位 = Nvl(v_合作单位, 合作单位),
          校对标志 = Nvl(n_操作状态, 校对标志), 收款时间 = d_登记时间, 操作员姓名 = v_操作员姓名, 操作员编号 = v_操作员编号,
          关联交易id = Decode(Nvl(关联交易id, 0), 0, ID, 关联交易id), 是否电子票据 = Decode(记录性质, 5, Nvl(n_是否电子票据, 0), Null)
      Where 结帐id = n_结帐id;
    
      If Sql%NotFound Then
      
        Select Max(关联交易id), Max(摘要), Max(门诊号), Max(住院号), Max(付款方式名称), Max(是否电子票据)
        Into n_关联交易id, v_结算摘要, n_门诊号, n_住院号, v_付款方式名称, n_是否电子票据
        From 病人预交记录
        Where 结帐id In (Select 结帐id From 住院费用记录 Where NO = v_费用单号 And 记录性质 = 5 And 记录状态 In (0, 1, 3));
      
        n_组id := Zl_Get组id(v_操作员姓名);
      
        For c_费用 In (Select NO, Max(病人id) As 病人id, Max(主页id) As 主页id, Max(姓名) As 姓名, Max(性别) As 性别, Max(年龄) As 年龄,
                            Max(病人科室id) As 病人科室id, Sum(结帐金额) As 结帐金额
                     From 住院费用记录
                     Where 结帐id = n_结帐id
                     Group By NO) Loop
          Insert Into 病人预交记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 姓名, 性别, 年龄, 门诊号, 住院号, 付款方式名称, 科室id, 摘要, 金额, 结帐id, 结算序号, 冲预交, 结算方式,
             结算号码, 收款时间, 操作员编号, 操作员姓名, 缴款单位, 单位开户行, 单位帐号, 缴款组id, 预交类别, 卡类别id, 卡号, 交易流水号, 交易说明, 合作单位, 结算卡序号, 校对标志, 关联交易id,
             交易时间, 交易人员, 是否电子票据)
          
            Select 病人预交记录_Id.Nextval, v_费用单号, '' As 实际票号, 5, 2, c_费用.病人id, c_费用.主页id, c_费用.姓名, c_费用.性别, c_费用.年龄, n_门诊号,
                   n_住院号, v_付款方式名称, c_费用.病人科室id, v_结算摘要, Null, n_结帐id, -1 * n_结帐id, c_费用.结帐金额, v_结算方式, v_结算号码, d_登记时间,
                   v_操作员编号, v_操作员姓名, '' As 缴款单位, '' As 单位开户行, '' As 单位帐号, n_组id, Null,
                   Decode(Nvl(n_卡类别id, 0), 0, Null, n_卡类别id), v_卡号, v_交易流水号, v_交易说明, v_合作单位,
                   Decode(Nvl(n_结算卡序号, 0), 0, Null, n_结算卡序号), n_操作状态, n_关联交易id, d_登记时间, v_操作员姓名, n_是否电子票据
            From Dual;
        End Loop;
      End If;
    
    End If;
  
    j_Jsonlist := o_Json.Get_Pljson_List('other_list');
    If Not j_Jsonlist Is Null Then
      --先删除，后增加
    
      If Nvl(n_预交id, 0) <> 0 Then
        l_预交id.Extend;
        l_预交id(l_预交id.Count) := n_预交id;
      End If;
      For c_预交 In (Select a.Id
                   From 病人预交记录 A
                   Where a.卡类别id = Nvl(n_卡类别id, 0) And a.结帐id = Nvl(n_结帐id, 0) And a.病人id = Nvl(n_病人id, 0) And
                         a.Id <> Nvl(n_预交id, 0)) Loop
        l_预交id.Extend;
        l_预交id(l_预交id.Count) := c_预交.Id;
      End Loop;
    
      Forall I In 1 .. l_预交id.Count
        Delete 三方结算交易 Where 交易id = l_预交id(I);
    
      For J In 1 .. j_Jsonlist.Count Loop
        o_Json     := PLJson();
        o_Json     := PLJson(j_Jsonlist.Get(J));
        v_交易名称 := o_Json.Get_String('swap_name');
        v_交易内容 := o_Json.Get_String('swap_note');
      
        --再插入
        Forall I In 1 .. l_预交id.Count
          Insert Into 三方结算交易
            (交易id, 交易项目, 交易内容, 原预交id, 性质)
            Select l_预交id(I) As 预交id, v_交易名称, v_交易内容, -1 * Null, -1 * Null As 性质 From Dual;
      End Loop;
    End If;
  
  End If;

  If Nvl(n_操作状态, 0) <> 0 Then
  
    Json_Out := zlJsonOut('成功', 1);
  
    Return;
  End If;
  -----------------------------------
  --完成处理

  Select Sum(预交金额), Sum(结帐金额)
  Into n_预交金额, n_结算金额
  From (Select Sum(冲预交) As 预交金额, 0 As 结帐金额
         From 病人预交记录
         Where 结帐id = n_结帐id And 病人id = Nvl(n_病人id, 0)
         Union All
         Select 0 As 结算金额, Sum(结帐金额) From 住院费用记录 Where 结帐id = n_结帐id And 病人id = Nvl(n_病人id, 0));

  If Nvl(n_预交金额, 0) <> Nvl(n_结算金额, 0) Then
    v_Err_Msg := '卡费结算合计(' || n_预交金额 || '与费用合计(' || n_结算金额 || ')不一致,不能继续操作！';
    Raise Err_Item;
  End If;

  If Nvl(n_预交id, 0) <> 0 Then
    --1.预交余额处理
    --冲销时，以原预交电子票据ID为准
    Begin
      Select Nvl(预交电子票据, 0)
      Into n_Count
      From 病人预交记录
      Where 记录状态 In (1, 3) And NO = v_预交单号 And 记录性质 = 1 And ID + 0 <> n_预交id;
    Exception
      When Others Then
        n_Count := Null;
    End;
    If n_Count Is Not Null Then
      n_预交电子票据 := n_Count;
    End If;
  
    Update 病人预交记录
    Set 记录状态 = Decode(记录状态, 0, 1, 记录状态), 校对标志 = Null, 预交电子票据 = n_预交电子票据
    Where ID = n_预交id And 病人id = n_病人id And 记录性质 = 1 And (Nvl(记录状态, 0) = 0 Or Nvl(记录状态, 0) = 2);
    If Sql%NotFound Then
      v_Err_Msg := '未找到预交单据号(' || v_预交单号 || ')的结算数据，可能因并发原因被他人收款或作废，请检查！';
      Raise Err_Item;
    End If;
  
    For c_预交 In (Select Max(Decode(a.记录状态, 2, 0, a.Id)) As ID, a.病人id, Max(a.预交类别) As 预交类别,
                        Sum(Decode(a.Id, n_预交id, a.金额, 0)) As 金额, Max(Decode(a.Id, n_预交id, b.性质, -1)) As 性质
                 From 病人预交记录 A, 结算方式 B
                 Where a.结算方式 = b.名称(+) And a.记录性质 = 1 And a.No = v_预交单号 And 病人id = Nvl(n_病人id, 0)
                 Group By a.病人id) Loop
    
      Update 预交单据余额
      Set 预交余额 = Nvl(预交余额, 0) + Nvl(c_预交.金额, 0)
      Where 预交id = c_预交.Id Return Nvl(预交余额, 0) Into n_返回值;
      If Sql%NotFound Then
        Insert Into 预交单据余额
          (预交id, 病人id, 预交类别, 预交余额)
        Values
          (c_预交.Id, c_预交.病人id, c_预交.预交类别, c_预交.金额);
        n_返回值 := c_预交.金额;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete 预交单据余额 Where 预交id = c_预交.Id And Nvl(预交余额, 0) = 0;
      End If;
    
      If Nvl(c_预交.性质, 1) <> 5 Then
        Update 病人余额
        Set 预交余额 = Nvl(预交余额, 0) + Nvl(c_预交.金额, 0)
        Where 性质 = 1 And 病人id = c_预交.病人id And Nvl(类型, 0) = Nvl(c_预交.预交类别, 0)
        Returning 预交余额 Into n_返回值;
        If Sql%RowCount = 0 Then
          Insert Into 病人余额
            (病人id, 性质, 类型, 预交余额, 费用余额)
          Values
            (c_预交.病人id, 1, Nvl(c_预交.预交类别, 0), Nvl(c_预交.金额, 0), 0);
          n_返回值 := Nvl(c_预交.金额, 0);
        End If;
        If Nvl(Nvl(c_预交.金额, 0), 0) = 0 Then
          Delete From 病人余额
          Where 病人id = c_预交.病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
        End If;
      End If;
    End Loop;
  End If;

  Update 住院费用记录
  Set 费用状态 = Null
  Where 结帐id = Nvl(n_结帐id, 0) And 病人id = n_病人id And Nvl(费用状态, 0) = 1;
  If Sql%NotFound Then
    v_Err_Msg := '未找到卡费费用单据(' || v_费用单号 || ')，可能因并发原因被他人收款或作废，请检查！';
    Raise Err_Item;
  End If;

  Select Max(是否电子票据), Count(*)
  Into n_Temp, n_Count
  From 病人预交记录
  Where 结帐id In (Select 结帐id From 住院费用记录 Where NO = v_费用单号 And 记录性质 = 5 And 记录状态 In (0, 1, 3)) And
        结帐id + 0 <> Nvl(n_结帐id, 0);

  If Nvl(n_Count, 0) <> 0 Then
    n_是否电子票据 := n_Temp;
  End If;

  Update 病人预交记录
  Set 校对标志 = Null, 是否电子票据 = Decode(记录性质, 5, Nvl(n_是否电子票据, 0), Null)
  Where 结帐id = Nvl(n_结帐id, 0) And Nvl(记录性质, 10) <> 1 And 病人id = n_病人id;
  If Sql%NotFound Then
    v_Err_Msg := '未找到卡费费用(单号为' || v_费用单号 || ')的结算信息，可能因并发原因被他人收款或作废，请检查！';
    Raise Err_Item;
  End If;

  --2.更新人员缴款余额
  For c_预交 In (Select 结算方式, 金额
               From 病人预交记录
               Where ID = Nvl(n_预交id, 0) And 记录性质 = 1 And 病人id = Nvl(n_病人id, 0)
               Union All
               Select 结算方式, 冲预交 From 病人预交记录 Where 结帐id = Nvl(n_结帐id, 0) And 病人id = Nvl(n_病人id, 0)) Loop
    If Nvl(c_预交.金额, 0) <> 0 Then
      Update 人员缴款余额
      Set 余额 = Nvl(余额, 0) + c_预交.金额
      Where 性质 = 1 And 收款员 = v_操作员姓名 And 结算方式 = c_预交.结算方式
      Returning 余额 Into n_返回值;
    
      If Sql%RowCount = 0 Then
        Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (v_操作员姓名, c_预交.结算方式, 1, c_预交.金额);
        n_返回值 := c_预交.金额;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 人员缴款余额
        Where 收款员 = v_操作员姓名 And 性质 = 1 And 结算方式 = c_预交.结算方式 And Nvl(余额, 0) = 0;
      End If;
    End If;
  End Loop;

  Json_Out := zlJsonOut('成功', 1);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Updcardfeeblncinfo;
/

Create Or Replace Procedure Zl_Exsesvr_Geteinvoicesinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) Is
  ------------------------------------------------------------------------------------------------- 
  --功能：获取电子票据信息
  --入参：json格式 
  --  input  
  --    query_type  N    查询范围:0-所有;1-只查询有效的电子票据;2-查询原始电子票据信息
  --    occasion  N  1  结算场合:1-收费,2-预交(包含押金),3-结帐,4-挂号,5-就诊卡,6-补充医保结算
  --    fee_nos  C    query_type=2时有效:单据号:结算场合=2时，为预交NO, 结算id未传入，该节点必传
  --    balance_id  N    结算ID：结算场合=2时，为预交ID
  --    read_oldbill  N  1  是否只读取原始单据的电子票据:1-是;2-否
  --    invoice_type  N    票种:1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡
  --出参：json格式 
  --  output 
  --    code  C  1  应答码：0-失败；1-成功
  --    message  C  1  应答消息：成功时返回成功信息，失败时返回具体的错误信息
  --    data  C    电子票据信息数据
  --    pati_info  C    病人信息
  --      pati_id  N  1  病人ID
  --      pati_pageid  N    主页ID
  --      pati_name  C  1  姓名
  --      pati_sex  C  1  性别
  --      pati_age  C  1  年龄
  --      outpatient_num  C  1  门诊号
  --      inpatient_num  C  1  住院号
  --    einvoice_info  C    电子票据信息:query_type=2时返回
  --      einv_id  N  1  电子票据ID
  --      paper_nos  C  1  未回收的纸质发票信息,多个用逗号返回
  --    einvoice_list[]  C    电子票据列表,query_type in (0,1)时返回
  --      einv_id  N  1  电子票据ID
  --      invoice_type  N  1  票种:1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡
  --      rec_state  N  1  记录状态
  --      placeCode  C  1  开票点编码
  --      inv_total  N  1  开票金额
  --      inv_oldid  N    原票据ID
  --      sys_source  C  1  系统来源
  --      demo  C  1  备注
  --      einvoice_code  C  1  电子票据代码
  --      einvoice_no  C  1  电子票据号码
  --      einvoice_random  C  1  电子校验码
  --      voucher_code  C  1  预交金凭证代码
  --      voucher_no  C  1  预交金凭证号码
  --      voucher_random  C  1  预交金凭证校验码
  --      happen_time  C  1  电子票据生成时间:yyyymmddhh24miss
  --      picture_url  C  1  电子票据H5页面URL
  --      picture_neturl  C  1  电子票据外网H5页面URL
  --      tran_paper  N  1  是否换开纸质发票
  --      trans_paperno  C  1  换开的纸质发票号
  --      trans_printid  N  1  换开的打印id
  --    operator_code  C  1  操作员编号
  --      operator_name  C  1  操作员姓名
  --    create_time  C  1  登记时间:yyyy-mm-dd hh24:mi:ss
  ------------------------------------------------------------------------------------------------- 
  j_Input PLJson;
  j_Json  PLJson;

  n_查询类型 Number(2);
  n_场合     Number(2);
  --n_返回二维码   Number(2);
  n_结算id       门诊费用记录.结帐id%Type;
  v_Nos          Varchar2(32767);
  n_票种         Number(2);
  v_Temp         Varchar2(32767);
  v_Output       Varchar2(32767);
  c_Output       Clob;
  n_仅读原始单据 Number(2);

  Cursor c_电子票据信息 Is(
    Select c.Id, Max(b.是否电子票据), Max(c.Id) As 电子票据id, Max(c.姓名) As 姓名, Max(c.性别) As 性别, Max(c.年龄) As 年龄,
           Max(c.病人id) As 病人id, Max(a.主页id) As 主页id, Max(c.门诊号) As 门诊号, Max(c.住院号) As 住院号, Max(a.No) As NO
    From 住院费用记录 A, 病人预交记录 B, 电子票据使用记录 C
    Where a.No = '-' And a.记录状态 In (1, 3) And a.结帐id = b.结帐id And a.结帐id = c.结算id(+) And c.票种(+) = 5 And c.记录状态(+) = 1
    Group By c.Id);

  r_电子票据信息 c_电子票据信息%RowType;

  Type Ty_Einvoce Is Ref Cursor;
  c_Einvoice Ty_Einvoce; --动态游标变量

  v_Pati Varchar2(32767);

Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  --0-所有;1-只查询有效的电子票据;2-查询原始电子票据信息
  n_查询类型     := Nvl(j_Json.Get_Number('query_type'), 0);
  n_场合         := Nvl(j_Json.Get_Number('occasion'), 0);
  n_票种         := j_Json.Get_Number('invoice_type');
  n_仅读原始单据 := j_Json.Get_Number('read_oldbill');

  --n_返回二维码 := Nvl(j_Json.Get_Number('return_qrcode'), 0);

  n_结算id := j_Json.Get_Number('balance_id');
  v_Nos    := j_Json.Get_String('fee_nos');

  If Nvl(n_结算id, 0) = 0 And v_Nos Is Null Then
    Json_Out := zlJsonOut('未传入需要查询结算id及费用单据!');
    Return;
  End If;

  If Nvl(n_查询类型, 0) = 2 Then
    --2-查询原始电子票据信息
    --n_场合:1-收费,2-预交(包含押金),3-结帐,4-挂号,5-就诊卡,6-补充医保结算
    If n_场合 = 1 Or n_场合 = 4 Then
      --收费或挂号
      If Nvl(n_结算id, 0) <> 0 Then
        Open c_Einvoice For
          Select c.Id, Max(b.是否电子票据), Max(c.Id) As 电子票据id, Max(c.姓名) As 姓名, Max(c.性别) As 性别, Max(c.年龄) As 年龄,
                 Max(c.病人id) As 病人id, Max(a.主页id) As 主页id, Max(c.门诊号) As 门诊号, Max(c.住院号) As 住院号, Max(a.No) As NO
          From 门诊费用记录 A, 病人预交记录 B, 电子票据使用记录 C
          Where a.No In (Select Distinct NO From 门诊费用记录 Where 结帐id = n_结算id And Mod(记录性质, 10) = 1) And a.记录性质 = n_场合 And
                a.记录状态 In (1, 3) And a.结帐id = b.结帐id And a.结帐id = c.结算id(+) And c.票种(+) = 1 And c.记录状态(+) = 1
          
          Group By c.Id;
      Elsif Instr(v_Nos, ',') > 0 Then
        Open c_Einvoice For
          Select c.Id, Max(b.是否电子票据), Max(c.Id) As 电子票据id, Max(c.姓名) As 姓名, Max(c.性别) As 性别, Max(c.年龄) As 年龄,
                 Max(c.病人id) As 病人id, Max(a.主页id) As 主页id, Max(c.门诊号) As 门诊号, Max(c.住院号) As 住院号, Max(a.No) As NO
          From 门诊费用记录 A, 病人预交记录 B, 电子票据使用记录 C
          Where a.No In (Select Column_Value From Table(f_Str2List(v_Nos))) And a.记录状态 In (1, 3) And a.记录性质 = n_场合 And
                a.结帐id = b.结帐id And a.结帐id = c.结算id(+) And c.票种(+) = 1 And c.记录状态(+) = 1
          
          Group By c.Id;
      
      Else
        Open c_Einvoice For
          Select c.Id, Max(b.是否电子票据), Max(c.Id) As 电子票据id, Max(c.姓名) As 姓名, Max(c.性别) As 性别, Max(c.年龄) As 年龄,
                 Max(c.病人id) As 病人id, Max(a.主页id) As 主页id, Max(c.门诊号) As 门诊号, Max(c.住院号) As 住院号, Max(a.No) As NO
          From 门诊费用记录 A, 病人预交记录 B, 电子票据使用记录 C
          Where a.No = v_Nos And a.记录状态 In (1, 3) And a.记录性质 = n_场合 And a.结帐id = b.结帐id And a.结帐id = c.结算id(+) And
                c.票种(+) = 1 And c.记录状态(+) = 1
          
          Group By c.Id;
      End If;
    Elsif n_场合 = 2 Then
      --预交
      If Nvl(n_结算id, 0) <> 0 Then
        Open c_Einvoice For
          Select c.Id, Max(b.是否电子票据), Max(c.Id) As 电子票据id, Max(c.姓名) As 姓名, Max(c.性别) As 性别, Max(c.年龄) As 年龄,
                 Max(c.病人id) As 病人id, Max(b.主页id) As 主页id, Max(c.门诊号) As 门诊号, Max(c.住院号) As 住院号, Max(b.No) As NO
          From 病人预交记录 B, 电子票据使用记录 C
          Where b.No In (Select Distinct NO From 病人预交记录 Where ID = n_结算id And 记录性质 = 1) And b.记录性质 = 1 And
                b.记录状态 In (1, 3) And b.Id = c.结算id(+) And c.票种(+) = 2 And c.记录状态(+) = 1
          Group By c.Id;
      Elsif Instr(v_Nos, ',') > 0 Then
        Open c_Einvoice For
          Select c.Id, Max(b.是否电子票据), Max(c.Id) As 电子票据id, Max(c.姓名) As 姓名, Max(c.性别) As 性别, Max(c.年龄) As 年龄,
                 Max(c.病人id) As 病人id, Max(b.主页id) As 主页id, Max(c.门诊号) As 门诊号, Max(c.住院号) As 住院号, Max(b.No) As NO
          From 病人预交记录 B, 电子票据使用记录 C
          Where b.No In (Select Column_Value From Table(f_Str2List(v_Nos))) And b.记录状态 In (1, 3) And b.记录性质 = 1 And
                b.Id = c.结算id(+) And c.票种(+) = 2 And c.记录状态(+) = 1
          Group By c.Id;
      Else
        Open c_Einvoice For
          Select c.Id, Max(b.是否电子票据), Max(c.Id) As 电子票据id, Max(c.姓名) As 姓名, Max(c.性别) As 性别, Max(c.年龄) As 年龄,
                 Max(c.病人id) As 病人id, Max(b.主页id) As 主页id, Max(c.门诊号) As 门诊号, Max(c.住院号) As 住院号, Max(b.No) As NO
          From 病人预交记录 B, 电子票据使用记录 C
          Where b.No = v_Nos And b.记录状态 In (1, 3) And b.记录性质 = 1 And b.Id = c.结算id(+) And c.票种(+) = 2 And c.记录状态(+) = 1
          Group By c.Id;
      End If;
    
    Elsif n_场合 = 5 Then
      --就诊卡
      If Nvl(n_结算id, 0) <> 0 Then
        Open c_Einvoice For
          Select c.Id, Max(b.是否电子票据), Max(c.Id) As 电子票据id, Max(c.姓名) As 姓名, Max(c.性别) As 性别, Max(c.年龄) As 年龄,
                 Max(c.病人id) As 病人id, Max(a.主页id) As 主页id, Max(c.门诊号) As 门诊号, Max(c.住院号) As 住院号, Max(a.No) As NO
          From 住院费用记录 A, 病人预交记录 B, 电子票据使用记录 C
          Where a.No In (Select Distinct NO From 住院费用记录 Where 结帐id = n_结算id And Mod(记录性质, 10) = 5) And a.记录性质 = 5 And
                a.记录状态 In (1, 3) And a.结帐id = b.结帐id And a.结帐id = c.结算id(+) And c.票种(+) = 5 And c.记录状态(+) = 1
          Group By c.Id;
      Elsif Instr(v_Nos, ',') > 0 Then
        Open c_Einvoice For
          Select c.Id, Max(b.是否电子票据), Max(c.Id) As 电子票据id, Max(c.姓名) As 姓名, Max(c.性别) As 性别, Max(c.年龄) As 年龄,
                 Max(c.病人id) As 病人id, Max(a.主页id) As 主页id, Max(c.门诊号) As 门诊号, Max(c.住院号) As 住院号, Max(a.No) As NO
          From 住院费用记录 A, 病人预交记录 B, 电子票据使用记录 C
          Where a.No In (Select Column_Value From Table(f_Str2List(v_Nos))) And a.记录状态 In (1, 3) And a.记录性质 = 5 And
                a.结帐id = b.结帐id And a.结帐id = c.结算id(+) And c.票种(+) = 5 And c.记录状态(+) = 1
          Group By c.Id;
      
      Else
        Open c_Einvoice For
          Select c.Id, Max(b.是否电子票据), Max(c.Id) As 电子票据id, Max(c.姓名) As 姓名, Max(c.性别) As 性别, Max(c.年龄) As 年龄,
                 Max(c.病人id) As 病人id, Max(a.主页id) As 主页id, Max(c.门诊号) As 门诊号, Max(c.住院号) As 住院号, Max(a.No) As NO
          From 住院费用记录 A, 病人预交记录 B, 电子票据使用记录 C
          Where a.No = v_Nos And a.记录状态 In (1, 3) And a.结帐id = b.结帐id And a.记录性质 = 5 And a.结帐id = c.结算id(+) And
                c.票种(+) = 5 And c.记录状态(+) = 1
          Group By c.Id;
      End If;
    Elsif n_场合 = 6 Then
      --费用补充记录
      If Nvl(n_结算id, 0) <> 0 Then
        Open c_Einvoice For
          Select c.Id, Max(b.是否电子票据), Max(c.Id) As 电子票据id, Max(c.姓名) As 姓名, Max(c.性别) As 性别, Max(c.年龄) As 年龄,
                 Max(c.病人id) As 病人id, 0 As 主页id, Max(c.门诊号) As 门诊号, Max(c.住院号) As 住院号, Max(a.No) As NO
          From 费用补充记录 A, 病人预交记录 B, 电子票据使用记录 C
          Where a.No In (Select Distinct NO From 费用补充记录 Where 结算id = n_结算id And Mod(记录性质, 10) = 1) And a.记录性质 = 1 And
                a.记录状态 In (1, 3) And a.结算id = b.结帐id And a.结算id = c.结算id(+) And c.票种(+) = 1 And c.记录状态(+) = 1
          
          Group By c.Id;
      Elsif Instr(v_Nos, ',') > 0 Then
        Open c_Einvoice For
          Select c.Id, Max(b.是否电子票据), Max(c.Id) As 电子票据id, Max(c.姓名) As 姓名, Max(c.性别) As 性别, Max(c.年龄) As 年龄,
                 Max(c.病人id) As 病人id, 0 As 主页id, Max(c.门诊号) As 门诊号, Max(c.住院号) As 住院号, Max(a.No) As NO
          From 费用补充记录 A, 病人预交记录 B, 电子票据使用记录 C
          Where a.No In (Select Column_Value From Table(f_Str2List(v_Nos))) And a.记录状态 In (1, 3) And a.记录性质 = 1 And
                a.结算id = b.结帐id And a.结算id = c.结算id(+) And c.票种(+) = 1 And c.记录状态(+) = 1
          Group By c.Id;
      Else
        Open c_Einvoice For
          Select c.Id, Max(b.是否电子票据), Max(c.Id) As 电子票据id, Max(c.姓名) As 姓名, Max(c.性别) As 性别, Max(c.年龄) As 年龄,
                 Max(c.病人id) As 病人id, 0 As 主页id, Max(c.门诊号) As 门诊号, Max(c.住院号) As 住院号, Max(a.No) As NO
          From 费用补充记录 A, 病人预交记录 B, 电子票据使用记录 C
          Where a.No = v_Nos And a.记录状态 In (1, 3) And a.记录性质 = 1 And a.结算id = b.结帐id And a.结算id = c.结算id(+) And
                c.票种(+) = 1 And c.记录状态(+) = 1
          Group By c.Id;
      End If;
    Else
      Json_Out := zlJsonOut('场合节点传入值不对!');
      Return;
    End If;
    Fetch c_Einvoice
      Into r_电子票据信息;
    If c_Einvoice %NotFound Then
      Close c_Einvoice;
      If Nvl(n_结算id, 0) = 0 Then
        Json_Out := zlJsonOut('未找到原始结算(NO=' || v_Nos || ')的电子票据，请检查!');
      Else
        Json_Out := zlJsonOut('未找到原始结算(结算id=' || n_结算id || ')的电子票据，请检查!');
      End If;
      Return;
    End If;
  
    --数据性质:1-收费,2-预交(包含押金),3-结帐,4-挂号,5-就诊卡
    --结算场合:1-收费,2-预交(包含押金),3-结帐,4-挂号,5-就诊卡,6-补充医保结算
    v_Temp := Null;
    For c_票据 In (Select Distinct 号码
                 From 票据使用明细
                 Where 票种 = n_票种 And 打印id In (Select ID As 打印id
                                              From (Select b.Id
                                                     From 票据使用明细 A, 票据打印内容 B
                                                     Where a.打印id = b.Id And a.性质 = 1 And a.原因 In (1, 3) And
                                                           b.数据性质 = Decode(Nvl(n_场合, 0), 6, 1, n_场合) And b.No = r_电子票据信息.No
                                                     Order By a.使用时间 Desc)
                                              Where Rownum < 2)) Loop
      If v_Temp Is Null Then
        v_Temp := c_票据.号码;
      Else
        v_Temp := v_Temp || ',' || c_票据.号码;
      End If;
    End Loop;
    v_Output := v_Output || '{"pati_id":' || zlJsonStr(r_电子票据信息.病人id, 1);
    v_Output := v_Output || ',"pati_pageid":' || zlJsonStr(r_电子票据信息.主页id, 1);
    v_Output := v_Output || ',"pati_name":"' || zlJsonStr(r_电子票据信息.姓名) || '"';
    v_Output := v_Output || ',"pati_sex":"' || zlJsonStr(r_电子票据信息.性别) || '"';
    v_Output := v_Output || ',"pati_age":"' || zlJsonStr(r_电子票据信息.年龄) || '"';
    v_Output := v_Output || ',"outpatient_num":"' || zlJsonStr(r_电子票据信息.门诊号) || '"';
    v_Output := v_Output || ',"inpatient_num":"' || zlJsonStr(r_电子票据信息.住院号) || '"';
    v_Output := v_Output || '}';
  
    v_Output := '"pati_info":' || v_Output;
    --电子票据信息
    v_Output := v_Output || ',"einvoice_info":';
    v_Output := v_Output || '{"einv_id":' || zlJsonStr(r_电子票据信息.Id, 1);
    v_Output := v_Output || ',"paper_nos":"' || zlJsonStr(v_Temp) || '"}';
    Json_Out := '{"output":{"code":1,"message":"成功","data":{' || v_Output || '}}}';
    Return;
  End If;

  --电子票据信息
  v_Output := Null;
  v_Pati   := Null;
  For c_电子票据 In (
                 
                 Select ID, 票种, 记录状态, 结算id, 病人id, 0 As 主页id, 姓名, 性别, 年龄, 门诊号, 住院号, 代码, 号码, 检验码, 凭证代码, 凭证号码, 凭证检验码, 票据金额,
                         生成时间, 原票据id, 是否换开, 纸质发票号, 打印id, 备注, 操作员编号, 操作员姓名, 登记时间, 开票点, 系统来源, Url内网, Url外网
                 From 电子票据使用记录
                 Where 结算id = n_结算id And 票种 = n_票种 And ((Nvl(n_仅读原始单据, 0) = 1 And 退款id Is Null) Or Nvl(n_仅读原始单据, 0) = 0) And
                       ((Nvl(n_查询类型, 0) = 1 And 记录状态 = 1) Or Nvl(n_查询类型, 0) = 0)
                 Order By 登记时间 Desc)
  
   Loop
  
    If v_Pati Is Null Then
      v_Pati := v_Pati || '{"pati_id":' || zlJsonStr(c_电子票据.病人id, 1);
      v_Pati := v_Pati || ',"pati_pageid":' || zlJsonStr(c_电子票据.主页id, 1);
      v_Pati := v_Pati || ',"pati_name":"' || zlJsonStr(c_电子票据.姓名) || '"';
      v_Pati := v_Pati || ',"pati_sex":"' || zlJsonStr(c_电子票据.性别) || '"';
      v_Pati := v_Pati || ',"pati_age":"' || zlJsonStr(c_电子票据.年龄) || '"';
      v_Pati := v_Pati || ',"outpatient_num":"' || zlJsonStr(c_电子票据.门诊号) || '"';
      v_Pati := v_Pati || ',"inpatient_num":":' || zlJsonStr(c_电子票据.住院号) || '"';
      v_Pati := v_Pati || '}';
    End If;
  
    If v_Output Is Not Null Then
      v_Output := v_Output || ',';
    End If;
    v_Output := v_Output || '{"einv_id":' || zlJsonStr(c_电子票据.Id, 1);
    v_Output := v_Output || ',"invoice_type":' || zlJsonStr(c_电子票据.票种, 1);
    v_Output := v_Output || ',"rec_state":' || zlJsonStr(c_电子票据.记录状态, 1);
    v_Output := v_Output || ',"placeCode":"' || zlJsonStr(c_电子票据.开票点) || '"';
    v_Output := v_Output || ',"inv_total":' || zlJsonStr(c_电子票据.票据金额, 1);
    v_Output := v_Output || ',"inv_oldid":' || zlJsonStr(c_电子票据.原票据id, 1);
    v_Output := v_Output || ',"sys_source":"' || zlJsonStr(c_电子票据.系统来源) || '"';
    v_Output := v_Output || ',"demo":"' || zlJsonStr(c_电子票据.备注) || '"';
    v_Output := v_Output || ',"einvoice_code":"' || zlJsonStr(c_电子票据.代码) || '"';
    v_Output := v_Output || ',"einvoice_no":"' || zlJsonStr(c_电子票据.号码) || '"';
    v_Output := v_Output || ',"einvoice_random":"' || zlJsonStr(c_电子票据.检验码) || '"';
    v_Output := v_Output || ',"voucher_code":"' || zlJsonStr(c_电子票据.凭证代码) || '"';
    v_Output := v_Output || ',"voucher_no":"' || zlJsonStr(c_电子票据.凭证号码) || '"';
    v_Output := v_Output || ',"voucher_random":"' || zlJsonStr(c_电子票据.凭证检验码) || '"';
    v_Output := v_Output || ',"happen_time":"' || zlJsonStr(c_电子票据.生成时间) || '"';
    v_Output := v_Output || ',"picture_url":"' || zlJsonStr(c_电子票据.Url内网) || '"';
    v_Output := v_Output || ',"picture_neturl":"' || zlJsonStr(c_电子票据.Url外网) || '"';
  
    v_Output := v_Output || ',"tran_paper":' || zlJsonStr(Nvl(c_电子票据.是否换开, 0), 1);
    v_Output := v_Output || ',"trans_paperno":"' || zlJsonStr(c_电子票据.纸质发票号) || '"';
    v_Output := v_Output || ',"trans_printid":' || zlJsonStr(Nvl(c_电子票据.打印id, 0), 1);
    v_Output := v_Output || ',"operator_code":"' || zlJsonStr(c_电子票据.操作员编号) || '"';
    v_Output := v_Output || ',"operator_name":"' || zlJsonStr(c_电子票据.操作员姓名) || '"';
    v_Output := v_Output || ',"create_time":"' || To_Char(c_电子票据.登记时间, 'yyyy-mm-dd hh24:mi:ss') || '"';
  
    v_Output := v_Output || '}';
    If Length(v_Output) > 30000 Then
      If c_Output Is Null Then
        c_Output := Substr(v_Output, 2);
      Else
        c_Output := c_Output || v_Output;
      End If;
      v_Output := Null;
    End If;
  End Loop;

  If v_Pati Is Null Then
    v_Pati := v_Pati || '{"pati_id":' || zlJsonStr(0, 1);
    v_Pati := v_Pati || ',"pati_pageid":' || zlJsonStr(0, 1);
    v_Pati := v_Pati || ',"pati_name":"' || zlJsonStr('') || '"';
    v_Pati := v_Pati || ',"pati_sex":"' || zlJsonStr('') || '"';
    v_Pati := v_Pati || ',"pati_age":"' || zlJsonStr('') || '"';
    v_Pati := v_Pati || ',"outpatient_num":"' || zlJsonStr('') || '"';
    v_Pati := v_Pati || ',"inpatient_num":":' || zlJsonStr('') || '"';
    v_Pati := v_Pati || '}';
  End If;
  v_Pati := '"pati_info":' || v_Pati;
  If Not c_Output Is Null And Not v_Output Is Null Then
  
    c_Output := To_Clob(',"einvoice_list":[') || c_Output || ',' || To_Clob(v_Output || ']');
    c_Output := To_Clob(v_Pati) || c_Output;
    v_Output := '';
  Elsif Not c_Output Is Null And v_Output Is Null Then
    c_Output := To_Clob(',"einvoice_list":[') || c_Output || To_Clob(']');
    c_Output := To_Clob(v_Pati) || c_Output;
    v_Output := '';
  Else
    If Length(v_Pati || ',"einvoice_list":[' || v_Output || ']') <= 30000 Then
      v_Output := v_Pati || ',"einvoice_list":[' || v_Output || ']';
    Else
      c_Output := To_Clob(',"einvoice_list":[') || To_Clob(v_Output || ']');
      c_Output := To_Clob(v_Pati) || c_Output;
      v_Output := '';
    End If;
  End If;
  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"成功","data":{') || c_Output || To_Clob('}}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"成功","data":{' || v_Output || '}}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Geteinvoicesinfo;
/


Create Or Replace Procedure Zl_Exsesvr_Checkpativisitstate
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --功能：更新病人就诊状态检查
  --入参：Json_In:格式
  --input
  --  reg_no             C  1 挂号单
  --  exe_status         N  1 执行状态 0:标记为待诊.-1:标记为不就诊

  -- 出参:
  --  output
  --    code                            N 1 应答吗：0-失败；1-成功
  --    message                         C 1 应答消息：失败时返回具体的错误信息
  --    msg_mode                        N 1 检查消息提示模式 0-禁止 1-询问
  --    msg_text                        C 1 检查提示内容
  -------------------------------------------
  v_挂号单   Varchar2(50);
  n_执行状态 Number;

  v_摘要     Varchar2(4000);
  v_划价no   Varchar2(4000);
  n_Count    Number;
  n_提示模式 Number;
  v_提示内容 Varchar2(4000);
  v_Output   Varchar2(1000);

  j_Input PLJson;
  j_Json  PLJson;

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_挂号单   := j_Json.Get_String('reg_no');
  n_执行状态 := Nvl(j_Json.Get_Number('exe_status'), 0);

  --获取挂号划价单信息
  Select Max(摘要) Into v_摘要 From 门诊费用记录 Where NO = v_挂号单 And 记录性质 = 4 And 记录状态 = 1 And Rownum < 2;

  If v_摘要 Is Not Null And Nvl(Instr(v_摘要 || '', '划价:'), 0) <> 0 Then
    --获取挂号划价单信息,判断挂号划价单是否存在，不存在，则不允许将病人状态设置为待诊
    v_划价no := Substr(v_摘要, Length('划价:') + 1);
    Select Count(1) Into n_Count From 门诊费用记录 Where NO = v_划价no And Mod(记录性质, 10) = 1 And 记录状态 = 0;
    If n_Count < 1 Then
      If n_执行状态 = 0 Then
        n_提示模式 := 0;
        v_提示内容 := '该挂号单的划价费用不存在，请退号后重新挂号!';
      End If;
    Else
      If n_执行状态 = -1 Then
        n_提示模式 := 1;
        v_提示内容 := '该病人存在挂号单的划价费用，设置为不就诊时将删除该挂号单的划价费用，' || Chr(13) || Chr(10) || '并且不能再恢复为待诊,是否继续?';
      End If;
    End If;
  End If;

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '成功');
  zlJsonPutValue(v_Output, 'msg_mode', Nvl(n_提示模式, 0), 1);
  zlJsonPutValue(v_Output, 'msg_text', v_提示内容, 0, 2);
  Json_Out := '{"output":' || v_Output || '}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkpativisitstate;
/


Create Or Replace Procedure Zl_Exsesvr_Outpatiforcereceive
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --功能：强制续诊接收
  --入参：Json_In:格式
  --input
  --  pati_id            N  1 病人id
  --  reg_no             C  1 挂号no
  --  exe_deptid         N  1 执行部门id
  --  outp_room_name     C  1 接诊诊室
  --  emg_sign           N  1 急诊标志
  --  operator_name      C  1 操作员姓名
  --  operator_code      C  1 操作员编号
  --  operator_id        N  1 操作员id
  --  rgst_appt_sign     N  1 预约标志
  --  recv_time          C  1 接收时间
  --  outpno             N  1 门诊号
  --  reg_id             N  1 挂号id

  -- 出参:
  --  output
  --    code                            N 1 应答吗：0-失败；1-成功
  --    message                         C 1 应答消息：失败时返回具体的错误信息
  -------------------------------------------

  v_挂号单 Varchar2(50);

  v_操作员姓名 Varchar2(50);
  v_操作员编号 Varchar2(50);
  n_操作员id   Number;
  n_执行部门id Number;

  n_预约标志 Number;
  n_急诊标志 Number;

  v_接诊诊室 Varchar2(50);
  d_接收时间 Date;
  n_病人id   Number;
  n_门诊号   Number;
  n_挂号id   Number;

  j_Input PLJson;
  j_Json  PLJson;

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_操作员姓名 := j_Json.Get_String('operator_name');
  v_操作员编号 := j_Json.Get_String('operator_code');
  n_操作员id   := j_Json.Get_Number('operator_id');

  n_执行部门id := j_Json.Get_Number('exe_deptid');

  n_预约标志 := Nvl(j_Json.Get_Number('rgst_appt_sign'), 0);
  n_急诊标志 := Nvl(j_Json.Get_Number('emg_sign'), 0);
  v_挂号单   := j_Json.Get_String('reg_no');
  v_接诊诊室 := j_Json.Get_String('outp_room_name');
  d_接收时间 := To_Date(j_Json.Get_String('recv_time'), 'YYYY-MM-DD HH24:MI:SS');
  n_病人id   := j_Json.Get_Number('pati_id');

  If d_接收时间 Is Null Then
    d_接收时间 := Sysdate;
  End If;
  If n_预约标志 = 1 Then
    n_门诊号 := j_Json.Get_Number('outpno');
    n_挂号id := j_Json.Get_Number('reg_id');
    Zl_病人预约挂号_接收_s(v_挂号单, v_接诊诊室, Null, Null, Null, Null, Null, d_接收时间, Null, n_病人id, n_门诊号, n_挂号id);
  End If;

  Zl_就诊变动记录_Insert(v_挂号单, 3, '强制续诊', v_操作员姓名, v_操作员编号, Null, n_执行部门id, Null, n_操作员id, v_操作员姓名);

  Zl_病人接诊_s(n_病人id, v_挂号单, n_执行部门id, v_操作员姓名, v_接诊诊室, n_急诊标志, Null, d_接收时间);

  Json_Out := zlJsonOut('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Outpatiforcereceive;
/

Create Or Replace Procedure Zl_Exsesvr_Getpatitriagemode
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --功能：获取挂号记录的分诊方式
  --入参：Json_In:格式
  --input 
  --  reg_no  C  1 挂号单

  -- 出参:
  --  output
  --    code                            N 1 应答吗：0-失败；1-成功
  --    message                         C 1 应答消息：失败时返回具体的错误信息
  --    triage_mode                     N 1 分诊方式
  -------------------------------------------

  v_挂号单   Varchar2(50);
  n_挂号模式 Number(3);
  n_分诊方式 Number(3);
  v_Output   Varchar2(1000);
  j_Input    PLJson;
  j_Json     PLJson;

Begin
  --解析入参
  j_Input    := PLJson(Json_In);
  j_Json     := j_Input.Get_Pljson('input');
  v_挂号单   := j_Json.Get_String('reg_no');
  n_挂号模式 := To_Number(Nvl(Substr(zl_GetSysParameter(256), 1, 1), 0));

  Begin
    If Nvl(n_挂号模式, 0) = 0 Then
      Select Nvl(Max(a.分诊方式), 0)
      Into n_分诊方式
      From 挂号安排 A, 病人挂号记录 B
      Where a.号码 = b.号别 And b.No = v_挂号单;
    Else
      Select Nvl(Max(a.分诊方式), 0)
      Into n_分诊方式
      From 临床出诊记录 A, 病人挂号记录 B
      Where a.Id = b.出诊记录id And b.No = v_挂号单;
    End If;
  Exception
    When Others Then
      Null;
  End;

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '成功');
  zlJsonPutValue(v_Output, 'triage_mode', n_分诊方式, 1, 2);
  Json_Out := '{"output":' || v_Output || '}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getpatitriagemode;
/


Create Or Replace Procedure Zl_Exsesvr_Getrgsapptpatilist
(
  Json_In  In Varchar2,
  Json_Out Out Clob
) As
  -------------------------------------------
  --功能：根据条件获取预约病人列表
  --入参：Json_In:格式
  --input
  --  operator_name          C    1 操作员姓名
  --  outp_recv_dept_id      C    1 门诊接诊科室ID
  --  outp_recv_Range        N    1 门诊接诊范围 1-挂本人号 2-本诊室（预约挂号的发药窗口填的是预号，没有诊室），本科室
  --  emg_sign               N    0 急诊标志
  --  err_sign               N    0 异常标志

  --出参      json
  --output
  --    code                N   1 应答吗：0-失败；1-成功
  --    message             C   1 应答消息：失败时返回具体的错误信息
  --    pati_list[]           病人列表，支持多个，[数组]
  --       reg_id           N   1 挂号ID
  --       reg_no           C   1 挂号单
  --       pati_id          N   1 病人id
  --       outpatient_num   C   1 门诊号
  --       pati_name        C   1 姓名
  --       pati_sex         C   1 性别
  --       pati_age         C   1 年龄
  --       emg_sign         N   1 急诊
  --       happen_time      C   1 发生时间
  --       exe_deptid       N   1 执行科室ID
  --       exetr            C   1 执行人
  --       outp_rfrl_status N   1 转诊状态
  --       record_sign      N   1 记录标志
  --       outptyp_name     C   1 号类
  --       pait_dept        C   1 病人科室
  --       exe_status       N   1 执行状态
  j_Input Pljson;
  j_Json  Pljson;

  v_操作员姓名 Varchar(50);
  n_接诊科室id Number(18);
  n_接诊范围   Number(5);
  n_急诊标志   Number(5);
  n_异常标志   Number(5);

  v_Para     Varchar(50);
  n_挂号安排 Number;
  d_启用时间 Date;
  v_Jtmp     Varchar2(32767);
  c_Jtmp     Clob;
Begin
  j_Input := Pljson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_操作员姓名 := j_Json.Get_String('operator_name');
  n_接诊科室id := Nvl(j_Json.Get_Number('outp_recv_dept_id'), 0);
  n_接诊范围   := Nvl(j_Json.Get_Number('outp_recv_Range'), 0);
  n_急诊标志   := Nvl(j_Json.Get_Number('emg_sign'), 0);
  n_异常标志   := Nvl(j_Json.Get_Number('err_sign'), 0);
  v_Para       := zl_GetSysParameter(256);
  If Nvl(Zl_To_Number(Substr(v_Para, 1, 1)), 0) <> 0 Then
    Begin
      d_启用时间 := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
    Exception
      When Others Then
        d_启用时间 := Null;
    End;
    If Sysdate >= d_启用时间 Then
      n_挂号安排 := 1;
    End If;
  End If;

  If n_异常标志 = 1 Then
    For c_病人列表 In (Select b.Id, b.No, b.病人id, b.门诊号, b.姓名, b.性别, b.年龄, b.复诊, b.急诊, b.社区, b.发生时间 As 时间, b.号类, e.名称 As 病人科室,
                          b.号序, b.诊室, b.分诊时间, b.发生时间, b.执行部门id, b.执行人, b.转诊状态, f.名称 As 转诊科室, b.转诊诊室, b.转诊医生, b.执行状态,
                          b.记录标志
                   From 病人挂号记录 B, 临床出诊记录 C, 部门表 E, 部门表 F
                   Where b.病人id Is Not Null And b.出诊记录id = c.Id And b.执行部门id = e.Id And b.转诊科室id = f.Id(+) And
                         b.记录性质 = 1 And b.记录状态 = 1 And ((n_急诊标志 = 1 And b.急诊 = 1) Or n_急诊标志 = 0) And Nvl(b.记录标志, 0) = -1 And
                         b.发生时间 Between Trunc(Sysdate) And Trunc(Sysdate + 1) - 1 / 24 / 60 / 60 And
                         Sysdate Between c.开始时间 And c.终止时间 And
                         (n_接诊范围 = 0 Or (n_接诊范围 = 1 And b.执行人 || '' = v_操作员姓名 || '') Or
                         ((n_接诊范围 = 2 or n_接诊范围 = 3) And b.执行部门id + 0 = n_接诊科室id And (b.执行人 || '' = v_操作员姓名 Or b.执行人 Is Null)))
                   Order By 发生时间) Loop
    
      v_Jtmp := v_Jtmp || ',{"reg_id":' || c_病人列表.Id;
      v_Jtmp := v_Jtmp || ',"reg_no":"' || c_病人列表.No || '"';
      v_Jtmp := v_Jtmp || ',"pati_id":' || c_病人列表.病人id;
      v_Jtmp := v_Jtmp || ',"outpatient_num":"' || c_病人列表.门诊号 || '"';
      v_Jtmp := v_Jtmp || ',"pati_name":"' || Zljsonstr(c_病人列表.姓名) || '"';
      v_Jtmp := v_Jtmp || ',"pati_sex":"' || c_病人列表.性别 || '"';
      v_Jtmp := v_Jtmp || ',"pati_age":"' || c_病人列表.年龄 || '"';
      v_Jtmp := v_Jtmp || ',"emg_sign":' || Nvl(c_病人列表.急诊, 0);
      v_Jtmp := v_Jtmp || ',"happen_time":"' || To_Char(c_病人列表.发生时间, 'YYYY-MM-DD HH24:MI') || '"';
      v_Jtmp := v_Jtmp || ',"exe_deptid":' || Nvl(c_病人列表.执行部门id || '', 'null');
      v_Jtmp := v_Jtmp || ',"exetr":"' || Zljsonstr(c_病人列表.执行人) || '"';
      v_Jtmp := v_Jtmp || ',"outp_rfrl_status":' || Nvl(c_病人列表.转诊状态 || '', 'null');
      v_Jtmp := v_Jtmp || ',"record_sign":' || Nvl(c_病人列表.记录标志 || '', 'null');
      v_Jtmp := v_Jtmp || ',"outptyp_name":"' || Zljsonstr(c_病人列表.号类) || '"';
      v_Jtmp := v_Jtmp || ',"pait_dept":"' || Zljsonstr(c_病人列表.病人科室) || '"';
      v_Jtmp := v_Jtmp || ',"exe_status":' || Nvl(c_病人列表.执行状态 || '', 'null');
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
    If n_挂号安排 = 1 Then
      For c_病人列表 In (Select b.Id, b.No, b.病人id, b.门诊号, b.姓名, b.性别, b.年龄, b.复诊, b.急诊, b.社区, b.发生时间 As 时间, b.号类,
                            e.名称 As 病人科室, b.号序, b.诊室, b.分诊时间, b.发生时间, b.执行部门id, b.执行人, b.转诊状态, f.名称 As 转诊科室, b.转诊诊室,
                            b.转诊医生, b.执行状态, b.记录标志
                     From 病人挂号记录 B, 临床出诊记录 C, 部门表 E, 部门表 F
                     Where b.病人id Is Not Null And b.出诊记录id = c.Id And b.执行部门id = e.Id And b.转诊科室id = f.Id(+) And
                           b.记录性质 = 2 And b.记录状态 = 1 And ((n_急诊标志 = 1 And b.急诊 = 1) Or n_急诊标志 = 0) And
                           Nvl(b.记录标志, 0) <> -1 And b.发生时间 Between Trunc(Sysdate) And
                           Trunc(Sysdate + 1) - 1 / 24 / 60 / 60 And Sysdate Between c.开始时间 And c.终止时间 And
                           (n_接诊范围 = 0 Or (n_接诊范围 = 1 And b.执行人 || '' = v_操作员姓名 || '') Or
                           ((n_接诊范围 = 2 or n_接诊范围 = 3) And b.执行部门id + 0 = n_接诊科室id And (b.执行人 || '' = v_操作员姓名 Or b.执行人 Is Null)))
                     Order By 发生时间) Loop
      
        v_Jtmp := v_Jtmp || ',{"reg_id":' || c_病人列表.Id;
        v_Jtmp := v_Jtmp || ',"reg_no":"' || c_病人列表.No || '"';
        v_Jtmp := v_Jtmp || ',"pati_id":' || c_病人列表.病人id;
        v_Jtmp := v_Jtmp || ',"outpatient_num":"' || c_病人列表.门诊号 || '"';
        v_Jtmp := v_Jtmp || ',"pati_name":"' || Zljsonstr(c_病人列表.姓名) || '"';
        v_Jtmp := v_Jtmp || ',"pati_sex":"' || c_病人列表.性别 || '"';
        v_Jtmp := v_Jtmp || ',"pati_age":"' || c_病人列表.年龄 || '"';
        v_Jtmp := v_Jtmp || ',"emg_sign":' || Nvl(c_病人列表.急诊, 0);
        v_Jtmp := v_Jtmp || ',"happen_time":"' || To_Char(c_病人列表.发生时间, 'YYYY-MM-DD HH24:MI') || '"';
        v_Jtmp := v_Jtmp || ',"exe_deptid":' || Nvl(c_病人列表.执行部门id || '', 'null');
        v_Jtmp := v_Jtmp || ',"exetr":"' || Zljsonstr(c_病人列表.执行人) || '"';
        v_Jtmp := v_Jtmp || ',"outp_rfrl_status":' || Nvl(c_病人列表.转诊状态 || '', 'null');
        v_Jtmp := v_Jtmp || ',"record_sign":' || Nvl(c_病人列表.记录标志 || '', 'null');
        v_Jtmp := v_Jtmp || ',"outptyp_name":"' || Zljsonstr(c_病人列表.号类) || '"';
        v_Jtmp := v_Jtmp || ',"pait_dept":"' || Zljsonstr(c_病人列表.病人科室) || '"';
        v_Jtmp := v_Jtmp || ',"exe_status":' || Nvl(c_病人列表.执行状态 || '', 'null');
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
      For c_病人列表 In (Select Null As ID, a.No, a.病人id, a.标识号 As 门诊号, a.姓名, a.性别, a.年龄, a.是否急诊 As 急诊, a.执行人, b.号类,
                            d.名称 As 病人科室, a.发生时间 As 时间, a.发生时间, a.执行部门id, 0 As 执行状态, 0 As 记录标志, Null As 转诊状态
                     From 门诊费用记录 A, 挂号安排 B, 部门表 D
                     Where a.计算单位 = b.号码 And a.执行部门id = d.Id And a.序号 = 1 And a.记录性质 = 4 And a.记录状态 = 0 And
                           ((n_急诊标志 = 1 And a.是否急诊 = 1) Or n_急诊标志 = 0) And
                           Decode(To_Char(Sysdate, 'D'), '1', b.周日, '2', b.周一, '3', b.周二, '4', b.周三, '5', b.周四, '6', b.周五,
                                  '7', b.周六, Null) In
                           (Select 时间段
                            From 时间段
                            Where ('3000-01-10 ' || To_Char(Sysdate, 'HH24:MI:SS') Between
                                  Decode(Sign(开始时间 - 终止时间), 1, '3000-01-09 ' || To_Char(开始时间, 'HH24:MI:SS'),
                                          '3000-01-10 ' || To_Char(开始时间, 'HH24:MI:SS')) And
                                  '3000-01-10 ' || To_Char(终止时间, 'HH24:MI:SS')) Or
                                  ('3000-01-10 ' || To_Char(Sysdate, 'HH24:MI:SS') Between
                                  '3000-01-10 ' || To_Char(开始时间, 'HH24:MI:SS') And
                                  Decode(Sign(开始时间 - 终止时间), 1, '3000-01-11 ' || To_Char(终止时间, 'HH24:MI:SS'),
                                          '3000-01-10 ' || To_Char(终止时间, 'HH24:MI:SS')))) And
                           (n_接诊范围 = 0 Or (n_接诊范围 = 1 And a.执行人 || '' = v_操作员姓名 || '') Or
                           ((n_接诊范围 = 2 or n_接诊范围 = 3) And a.执行部门id + 0 = n_接诊科室id And (a.执行人 || '' = v_操作员姓名 Or a.执行人 Is Null))) And
                           a.发生时间 Between Trunc(Sysdate) And Trunc(Sysdate) + 1 - 1 / 24 / 60 / 60
                     Order By 发生时间) Loop
      
        v_Jtmp := v_Jtmp || ',{"reg_id":' || Nvl(c_病人列表.Id || '', 'null');
        v_Jtmp := v_Jtmp || ',"reg_no":"' || c_病人列表.No || '"';
        v_Jtmp := v_Jtmp || ',"pati_id":' || c_病人列表.病人id;
        v_Jtmp := v_Jtmp || ',"outpatient_num":"' || c_病人列表.门诊号 || '"';
        v_Jtmp := v_Jtmp || ',"pati_name":"' || Zljsonstr(c_病人列表.姓名) || '"';
        v_Jtmp := v_Jtmp || ',"pati_sex":"' || c_病人列表.性别 || '"';
        v_Jtmp := v_Jtmp || ',"pati_age":"' || c_病人列表.年龄 || '"';
        v_Jtmp := v_Jtmp || ',"emg_sign":' || Nvl(c_病人列表.急诊, 0);
        v_Jtmp := v_Jtmp || ',"happen_time":"' || To_Char(c_病人列表.发生时间, 'YYYY-MM-DD HH24:MI') || '"';
        v_Jtmp := v_Jtmp || ',"exe_deptid":' || Nvl(c_病人列表.执行部门id || '', 'null');
        v_Jtmp := v_Jtmp || ',"exetr":"' || Zljsonstr(c_病人列表.执行人) || '"';
        v_Jtmp := v_Jtmp || ',"outp_rfrl_status":' || Nvl(c_病人列表.转诊状态 || '', 'null');
        v_Jtmp := v_Jtmp || ',"record_sign":' || Nvl(c_病人列表.记录标志 || '', 'null');
        v_Jtmp := v_Jtmp || ',"outptyp_name":"' || Zljsonstr(c_病人列表.号类) || '"';
        v_Jtmp := v_Jtmp || ',"pait_dept":"' || Zljsonstr(c_病人列表.病人科室) || '"';
        v_Jtmp := v_Jtmp || ',"exe_status":' || Nvl(c_病人列表.执行状态 || '', 'null');
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
    Json_Out := '{"output":{"code":1,"message":"成功","pati_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","pati_list":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getrgsapptpatilist;
/


Create Or Replace Procedure Zl_Exsesvr_Getvalidreglist
(
  Json_In  In Varchar2,
  Json_Out Out Clob
) As
  -------------------------------------------
  --功能：根据条件获取病人的有效挂号记录
  --入参：Json_In:格式
  --input
  --  query_type        N    1 调用类型 ：1-根据病人id获取病人的有效挂号记录，2-根据挂号单获取信息
  --  pati_id           N    1 病人id
  --  emg_sign          N    1 急诊标志 ：0-全部挂号  1-急诊挂号
  --  reg_no            C    0 挂号单，

  --出参      json
  --output
  --    code                N   1 应答吗：0-失败；1-成功
  --    message             C   1 应答消息：失败时返回具体的错误信息
  --    reg_list[]           挂号记录列表，支持多个，[数组]
  --       reg_id           N   1 挂号id
  --       reg_no           C   1 挂号单
  --       reg_properties   N   1 记录性质
  --       exe_deptid       N   1 执行部门id
  --       exe_dept         C   1 执行部门
  --       fitem_id         N   1 收费细目id
  --       fitem_name       C   1 收费细目
  --       exetr            C   1 执行人
  --       outp_room_name   C   1 诊室
  --       happen_time      C   1 发生时间
  --       exe_status       N   1 执行状态
  --       emg_sign         N   1 急诊

  j_Input PLJson;
  j_Json  PLJson;

  n_Type             Number(5);
  n_病人id           Number(18);
  n_是否急诊         Number(18);
  n_允许超过挂号日期 Number(18);
  v_挂号单           Varchar2(60);
  n_挂号有效天数     Number;
  n_急诊有效天数     Number;
  v_Jtmp             Varchar2(32767);
  c_Jtmp             Clob;

Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_Type := Nvl(j_Json.Get_Number('query_type'), 0);
  If n_Type = 1 Then
    n_病人id           := Nvl(j_Json.Get_Number('pati_id'), 0);
    n_允许超过挂号日期 := To_Number(Nvl(zl_GetSysParameter(210), '0'));
    n_是否急诊         := Nvl(j_Json.Get_Number('emg_sign'), 0);
  
    If n_允许超过挂号日期 = 1 Then
      For c_挂号列表 In (Select a.Id, a.No, a.记录性质, d.Id As 科室id, d.名称 As 科室, c.Id As 项目id, c.名称 As 项目, a.执行人, a.诊室, a.发生时间,
                            a.执行状态, a.急诊
                     From 病人挂号记录 A, 门诊费用记录 B, 收费项目目录 C, 部门表 D
                     Where a.No = b.No And b.记录性质 = 4 And b.记录状态 In (1, 0) And b.收费类别 = '1' And a.记录性质 In (1, 2) And
                           a.记录状态 = 1 And b.价格父号 Is Null And b.从属父号 Is Null And b.收费细目id = c.Id And a.执行部门id = d.Id And
                           a.发生时间 <= Trunc(Sysdate) + 1 - 1 / 24 / 60 / 60 And a.病人id = n_病人id And
                           (n_是否急诊 = 0 Or (n_是否急诊 = 1 And a.急诊 = 1))
                     Order By 发生时间 Desc) Loop
      
        v_Jtmp := v_Jtmp || ',{"reg_id":' || Nvl(c_挂号列表.Id || '', 'null');
        v_Jtmp := v_Jtmp || ',"reg_no":"' || c_挂号列表.No || '"';
        v_Jtmp := v_Jtmp || ',"reg_properties":' || Nvl(c_挂号列表.记录性质 || '', 'null');
        v_Jtmp := v_Jtmp || ',"exe_deptid":' || Nvl(c_挂号列表.科室id || '', 'null');
        v_Jtmp := v_Jtmp || ',"exe_dept":"' || zlJsonStr(c_挂号列表.科室) || '"';
        v_Jtmp := v_Jtmp || ',"fitem_id":' || Nvl(c_挂号列表.项目id || '', 'null');
        v_Jtmp := v_Jtmp || ',"fitem_name":"' || zlJsonStr(c_挂号列表.项目) || '"';
        v_Jtmp := v_Jtmp || ',"exetr":"' || zlJsonStr(c_挂号列表.执行人) || '"';
        v_Jtmp := v_Jtmp || ',"outp_room_name":"' || zlJsonStr(c_挂号列表.诊室) || '"';
        v_Jtmp := v_Jtmp || ',"happen_time":"' || To_Char(c_挂号列表.发生时间, 'YYYY-MM-DD HH24:MI') || '"';
        v_Jtmp := v_Jtmp || ',"exe_status":' || Nvl(c_挂号列表.执行状态 || '', 'null');
        v_Jtmp := v_Jtmp || ',"emg_sign":' || Nvl(c_挂号列表.急诊 || '', 'null');
        v_Jtmp := v_Jtmp || '}';
      
        If Length(v_Jtmp) > 20000 Then
          If c_Jtmp Is Null Then
            c_Jtmp := Substr(v_Jtmp, 2);
          Else
            c_Jtmp := c_Jtmp || v_Jtmp;
          End If;
          v_Jtmp := Null;
        End If;
      
      End Loop;
    Else
      n_挂号有效天数 := To_Number(Substr(Nvl(zl_GetSysParameter(21), '11') || '11', 1, 1));
      n_急诊有效天数 := To_Number(Substr(Nvl(zl_GetSysParameter(21), '11') || '11', 2, 1));
      If n_挂号有效天数 = 0 Then
        n_挂号有效天数 := 1;
      End If;
      If n_急诊有效天数 = 0 Then
        n_急诊有效天数 := 1;
      End If;
    
      For c_挂号列表 In (Select a.Id, a.No, a.记录性质, d.Id As 科室id, d.名称 As 科室, c.Id As 项目id, c.名称 As 项目, a.执行人, a.诊室, a.发生时间,
                            a.执行状态, a.急诊
                     From 病人挂号记录 A, 门诊费用记录 B, 收费项目目录 C, 部门表 D
                     Where a.No = b.No And b.记录性质 = 4 And b.记录状态 In (1, 0) And b.收费类别 = '1' And a.记录性质 In (1, 2) And
                           a.记录状态 = 1 And b.价格父号 Is Null And b.从属父号 Is Null And b.收费细目id = c.Id And a.执行部门id = d.Id And
                           a.发生时间 <= Trunc(Sysdate) + 1 - 1 / 24 / 60 / 60 And a.病人id = n_病人id And
                           a.发生时间 Between Sysdate - Decode(a.急诊, 1, n_急诊有效天数, n_挂号有效天数) And
                           Trunc(Sysdate) + 1 - 1 / 24 / 60 / 60 And (n_是否急诊 = 0 Or (n_是否急诊 = 1 And a.急诊 = 1))
                     Order By 发生时间 Desc) Loop
      
        v_Jtmp := v_Jtmp || ',{"reg_id":' || Nvl(c_挂号列表.Id || '', 'null');
        v_Jtmp := v_Jtmp || ',"reg_no":"' || c_挂号列表.No || '"';
        v_Jtmp := v_Jtmp || ',"reg_properties":' || Nvl(c_挂号列表.记录性质 || '', 'null');
        v_Jtmp := v_Jtmp || ',"exe_deptid":' || Nvl(c_挂号列表.科室id || '', 'null');
        v_Jtmp := v_Jtmp || ',"exe_dept":"' || zlJsonStr(c_挂号列表.科室) || '"';
        v_Jtmp := v_Jtmp || ',"fitem_id":' || Nvl(c_挂号列表.项目id || '', 'null');
        v_Jtmp := v_Jtmp || ',"fitem_name":"' || zlJsonStr(c_挂号列表.项目) || '"';
        v_Jtmp := v_Jtmp || ',"exetr":"' || zlJsonStr(c_挂号列表.执行人) || '"';
        v_Jtmp := v_Jtmp || ',"outp_room_name":"' || zlJsonStr(c_挂号列表.诊室) || '"';
        v_Jtmp := v_Jtmp || ',"happen_time":"' || To_Char(c_挂号列表.发生时间, 'YYYY-MM-DD HH24:MI') || '"';
        v_Jtmp := v_Jtmp || ',"exe_status":' || Nvl(c_挂号列表.执行状态 || '', 'null');
        v_Jtmp := v_Jtmp || ',"emg_sign":' || Nvl(c_挂号列表.急诊 || '', 'null');
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
  Elsif n_Type = 2 Then
    v_挂号单 := j_Json.Get_String('reg_no');
    For c_挂号列表 In (Select a.Id, a.No, a.记录性质, a.执行部门id As 科室id, Null As 科室, Null As 项目id, Null As 项目, a.执行人, a.诊室, a.发生时间,
                          a.执行状态, a.急诊
                   From 病人挂号记录 A
                   Where a.No = v_挂号单 And a.记录性质 = 1 And a.记录状态 = 1) Loop
    
      v_Jtmp := v_Jtmp || ',{"reg_id":' || Nvl(c_挂号列表.Id || '', 'null');
      v_Jtmp := v_Jtmp || ',"reg_no":"' || c_挂号列表.No || '"';
      v_Jtmp := v_Jtmp || ',"reg_properties":' || Nvl(c_挂号列表.记录性质 || '', 'null');
      v_Jtmp := v_Jtmp || ',"exe_deptid":' || Nvl(c_挂号列表.科室id || '', 'null');
      v_Jtmp := v_Jtmp || ',"exe_dept":"' || zlJsonStr(c_挂号列表.科室) || '"';
      v_Jtmp := v_Jtmp || ',"fitem_id":' || Nvl(c_挂号列表.项目id || '', 'null');
      v_Jtmp := v_Jtmp || ',"fitem_name":"' || zlJsonStr(c_挂号列表.项目) || '"';
      v_Jtmp := v_Jtmp || ',"exetr":"' || zlJsonStr(c_挂号列表.执行人) || '"';
      v_Jtmp := v_Jtmp || ',"outp_room_name":"' || zlJsonStr(c_挂号列表.诊室) || '"';
      v_Jtmp := v_Jtmp || ',"happen_time":"' || To_Char(c_挂号列表.发生时间, 'YYYY-MM-DD HH24:MI') || '"';
      v_Jtmp := v_Jtmp || ',"exe_status":' || Nvl(c_挂号列表.执行状态 || '', 'null');
      v_Jtmp := v_Jtmp || ',"emg_sign":' || Nvl(c_挂号列表.急诊 || '', 'null');
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
    Json_Out := '{"output":{"code":1,"message":"成功","reg_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","reg_list":[' || c_Jtmp || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getvalidreglist;
/

Create Or Replace Procedure Zl_Exsesvr_Outprevisit
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --功能：门诊病人回诊
  --入参：Json_In:格式
  --input
  --  reg_id             N  1 挂号id
  --  exe_deptid         N  1 执行部门id
  --  outp_room_name     C  1 接诊诊室
  --  outp_dr_name       C  1 医生
  --  revisit_sign       N  1 回诊标志
  --  appt_mode_name     C  1 预约方式

  -- 出参:
  --  output
  --    code                            N 1 应答吗：0-失败；1-成功
  --    message                         C 1 应答消息：失败时返回具体的错误信息
  -------------------------------------------
  n_挂号id     Number;
  n_执行部门id Number;
  v_接诊诊室   Varchar2(50);
  v_医生       Varchar2(50);
  n_回诊标志   Number;
  v_预约方式   Varchar2(50);

  j_Input PLJson;
  j_Json  PLJson;

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_挂号id     := j_Json.Get_Number('reg_id');
  n_执行部门id := j_Json.Get_Number('exe_deptid');
  v_接诊诊室   := j_Json.Get_String('outp_room_name');
  v_医生       := j_Json.Get_String('outp_dr_name');
  n_回诊标志   := j_Json.Get_Number('revisit_sign');
  v_预约方式   := j_Json.Get_String('appt_mode_name');

  Zl_病人挂号记录_回诊(n_挂号id, n_执行部门id, v_接诊诊室, v_医生, n_回诊标志, v_预约方式);
  Json_Out := zlJsonOut('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Outprevisit;
/

Create Or Replace Procedure Zl_Exsesvr_Outpfinish
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --功能：门诊病人完成接诊
  --入参：Json_In:格式
  --input
  --  pati_id            N  1 病人ID
  --  reg_no             C  1 挂号no
  --  outp_room_name     C  1 接诊诊室
  --  exetr              C  1 执行人
  --  fnsh_desc          C  1 完成摘要
  --  ext_mark           N  1 附加标志 为1时,表示护士完成就诊;2时表示其他系统补充登记的挂号数据(此种方式不产生费用记录和挂号汇总,虽登记);3-其他三方系统同步记录

  -- 出参:
  --  output
  --    code                            N 1 应答吗：0-失败；1-成功
  --    message                         C 1 应答消息：失败时返回具体的错误信息
  -------------------------------------------
  n_病人id   Number;
  v_挂号单   Varchar2(50);
  v_接诊科室 Varchar2(50);
  v_执行人   Varchar2(50);
  v_完成摘要 Varchar2(4000);
  n_附加标志 Number;
  j_Input    PLJson;
  j_Json     PLJson;

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_病人id   := j_Json.Get_Number('pati_id');
  v_挂号单   := j_Json.Get_String('reg_no');
  v_接诊科室 := j_Json.Get_String('outp_room_name');
  v_执行人   := j_Json.Get_String('exetr');
  v_完成摘要 := j_Json.Get_String('fnsh_desc');
  n_附加标志 := j_Json.Get_Number('ext_mark');

  If n_附加标志 = 0 Then
    n_附加标志 := Null;
  End If;

  Zl_病人接诊完成_s(n_病人id, v_挂号单, v_接诊科室, v_执行人, v_完成摘要, n_附加标志);

  Json_Out := zlJsonOut('成功', 1);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Outpfinish;
/


Create Or Replace Procedure Zl_Exsesvr_Outpfinishcancel
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --功能：门诊病人取消完成接诊
  --入参：Json_In:格式
  --input
  --  pati_id            N  1 病人ID
  --  reg_no             C  1 挂号no
  --  ext_mark           N  1 附加标志 为1时,表示护士完成就诊;2时表示其他系统补充登记的挂号数据(此种方式不产生费用记录和挂号汇总,虽登记);3-其他三方系统同步记录

  -- 出参:
  --  output
  --    code                            N 1 应答吗：0-失败；1-成功
  --    message                         C 1 应答消息：失败时返回具体的错误信息
  -------------------------------------------
  n_病人id   Number;
  v_挂号单   Varchar2(50);
  n_附加标志 Number;
  j_Input    PLJson;
  j_Json     PLJson;

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_病人id   := j_Json.Get_Number('pati_id');
  v_挂号单   := j_Json.Get_String('reg_no');
  n_附加标志 := j_Json.Get_Number('ext_mark');
  If n_附加标志 = 0 Then
    n_附加标志 := Null;
  End If;
  Zl_病人接诊完成_Cancel_s(n_病人id, v_挂号单, n_附加标志);
  Json_Out := zlJsonOut('成功', 1);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Outpfinishcancel;
/


Create Or Replace Procedure Zl_Exsesvr_Outpreceive
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --功能：门诊病人接诊
  --入参：Json_In:格式
  --input
  --  pati_id            N  1 病人ID
  --  reg_no             C  1 挂号no
  --  exe_deptid         N  1 执行部门id
  --  exetr              C  1 执行人
  --  outp_room_name     C  1 接诊诊室
  --  emg_sign           N  1 急诊标志
  --  revisit_sign       N  1 回诊标志
  --  exe_time           C  1 执行时间  -- 出参:
  --  output
  --    code                            N 1 应答吗：0-失败；1-成功
  --    message                         C 1 应答消息：失败时返回具体的错误信息
  -------------------------------------------
  n_病人id     Number;
  v_挂号单     Varchar2(50);
  n_执行部门id Number;
  v_执行人     Varchar2(50);
  v_接诊科室   Varchar2(50);
  n_急诊标志   Number;
  n_回诊标志   Number;
  d_执行时间   Date;
  j_Input      PLJson;
  j_Json       PLJson;

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_病人id     := j_Json.Get_Number('pati_id');
  v_挂号单     := j_Json.Get_String('reg_no');
  n_执行部门id := j_Json.Get_Number('exe_deptid');
  v_执行人     := j_Json.Get_String('exetr');
  v_接诊科室   := j_Json.Get_String('outp_room_name');
  n_急诊标志   := j_Json.Get_Number('emg_sign');
  n_回诊标志   := j_Json.Get_Number('revisit_sign');
  d_执行时间   := To_Date(j_Json.Get_String('exe_time'), 'YYYY-MM-DD HH24:MI:SS');

  If n_执行部门id = 0 Then
    n_执行部门id := Null;
  End If;

  Zl_病人接诊_s(n_病人id, v_挂号单, n_执行部门id, v_执行人, v_接诊科室, n_急诊标志, n_回诊标志, d_执行时间);

  Json_Out := zlJsonOut('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Outpreceive;
/

CREATE OR REPLACE Procedure Zl_Exsesvr_Outpreceivecancel
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --功能：门诊病人取消接诊
  --入参：Json_In:格式
  --input
  --  pati_id            N  1 病人ID
  --  reg_no             C  1 挂号no
  --  exe_deptid         N  1 执行部门id
  --  exetr              C  1 执行人
  --  referral_sign      N  1 是否转诊 0-未转诊  1-转诊
  --  referral_deptid    N  1 转诊科室id
  --  referral_doctor    C  1 转诊医生

  -- 出参:
  --  output
  --    code                            N 1 应答吗：0-失败；1-成功
  --    message                         C 1 应答消息：失败时返回具体的错误信息
  -------------------------------------------
  n_病人id     Number;
  v_挂号单     Varchar2(50);
  n_执行部门id Number;
  v_执行人     Varchar2(50);
  n_转诊标志   Number;
  n_转诊科室id Number;
  v_转诊医生   Varchar2(50);
  j_Input      PLJson;
  j_Json       PLJson;
Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_病人id     := j_Json.Get_Number('pati_id');
  v_挂号单     := j_Json.Get_String('reg_no');
  n_执行部门id := j_Json.Get_Number('exe_deptid');
  v_执行人     := j_Json.Get_String('exetr');
  n_转诊标志   := Nvl(j_Json.Get_Number('referral_sign'), 0);
  v_转诊医生   := j_Json.Get_String('referral_doctor');
  n_转诊科室id := j_Json.Get_Number('referral_deptid');

  If n_执行部门id = 0 Then
    n_执行部门id := Null;
  End If;
  If n_转诊科室id = 0 Then
    n_转诊科室id := Null;
  End If;

  Zl_病人接诊_Cancel_s(n_病人id, v_挂号单, n_执行部门id, v_执行人, n_转诊标志, n_转诊科室id, v_转诊医生);
  Json_Out := zlJsonOut('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Outpreceivecancel;
/

Create Or Replace Procedure Zl_Exsesvr_Outpreferral
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --功能：完成病人转诊，转诊接收，取消转诊，拒绝转诊功能
  --入参：Json_In:格式
  --input
  --  reg_no             C  1 挂号no
  --  referral_state     N  1 转诊状态 0:转诊(需要传入其他参数),1:接收,-1:拒绝,Null:取消转诊
  --  referral_deptid    N  1 转诊科室id
  --  referral_outproom  C  1 转诊诊室
  --  referral_doctor    C  1 转诊医生

  -- 出参:
  --  output
  --    code                            N 1 应答吗：0-失败；1-成功
  --    message                         C 1 应答消息：失败时返回具体的错误信息
  -------------------------------------------

  v_挂号单     Varchar2(50);
  n_转诊状态   Number;
  n_转诊科室id Number;
  v_转诊诊室   Varchar2(50);
  v_转诊医生   Varchar2(50);
  j_Input      PLJson;
  j_Json       PLJson;

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_挂号单     := j_Json.Get_String('reg_no');
  v_转诊诊室   := j_Json.Get_String('referral_outproom');
  v_转诊医生   := j_Json.Get_String('referral_doctor');
  n_转诊状态   := j_Json.Get_Number('referral_state');
  n_转诊科室id := j_Json.Get_Number('referral_deptid');

  Zl_病人挂号记录_转诊_s(v_挂号单, n_转诊状态, n_转诊科室id, v_转诊诊室, v_转诊医生);
  Json_Out := zlJsonOut('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Outpreferral;
/


Create Or Replace Procedure Zl_Exsesvr_Outprevisitcancel
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --功能：门诊病人取消回诊
  --入参：Json_In:格式
  --input
  --  reg_id             N  1 挂号id
  --  revisit_sign       N  1 回诊标志

  -- 出参:
  --  output
  --    code                            N 1 应答吗：0-失败；1-成功
  --    message                         C 1 应答消息：失败时返回具体的错误信息
  -------------------------------------------
  n_挂号id   Number;
  n_回诊标志 Number;
  j_Input    PLJson;
  j_Json     PLJson;

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_挂号id   := j_Json.Get_Number('reg_id');
  n_回诊标志 := j_Json.Get_Number('revisit_sign');

  Zl_病人挂号记录_取消回诊(n_挂号id, n_回诊标志);

  Json_Out := zlJsonOut('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Outprevisitcancel;
/


CREATE OR REPLACE Procedure Zl_Exsesvr_Getqueuecallcount
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --功能：获取当前排队叫号的呼叫人数
  --入参：Json_In:格式
  --input
  --  operator_name  C  1 操作员姓名
  --  emg_sign       N  0 急诊标志
  -- 出参:
  --  output
  --    code                            N 1 应答吗：0-失败；1-成功
  --    message                         C 1 应答消息：失败时返回具体的错误信息
  --    call_count                      N 1 呼叫人数
  -------------------------------------------

  v_操作员姓名   Varchar2(50);
  n_挂号有效天数 Number;
  n_急诊有效天数 Number;
  n_呼叫含回诊   Number;
  n_呼叫人数     Number;
  n_急诊标志     Number;

  v_Output Varchar2(1000);
  j_Input  PLJson;
  j_Json   PLJson;

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_操作员姓名 := j_Json.Get_String('operator_name');
  n_急诊标志   := Nvl(j_Json.Get_Number('emg_sign'), 0);

  n_挂号有效天数 := To_Number(Substr(Nvl(zl_GetSysParameter(21), '11') || '11', 1, 1));
  n_急诊有效天数 := To_Number(Substr(Nvl(zl_GetSysParameter(21), '11') || '11', 2, 1));

  If n_挂号有效天数 = 0 Then
    n_挂号有效天数 := 1;
  End If;
  If n_急诊有效天数 = 0 Then
    n_急诊有效天数 := 1;
  End If;
  n_呼叫含回诊 := To_Number(Nvl(zl_GetSysParameter('就诊人数含回诊', 1260), 0));

  If n_急诊标志 = 1 Then
    If n_呼叫含回诊 <> 1 Then
      Select Count(Distinct b.Id) As Count
      Into n_呼叫人数
      From 病人挂号记录 B, 排队叫号队列 A
      Where a.业务id = b.Id And a.业务类型 = 0 And Instr(',0,4,', ',' || a.排队状态 || ',') = 0 And b.记录性质 = 1 And b.记录状态 = 1 And
            a.医生姓名 || '' = v_操作员姓名 And Nvl(a.回诊序号, 0) = 0 And (Nvl(b.急诊, 0) = 1 And b.发生时间 >= Sysdate - n_急诊有效天数);
    Else
      Select Count(Distinct b.Id) As Count
      Into n_呼叫人数
      From 病人挂号记录 B, 排队叫号队列 A
      Where a.业务id = b.Id And a.业务类型 = 0 And Instr(',0,4,6,', ',' || a.排队状态 || ',') = 0 And b.记录性质 = 1 And b.记录状态 = 1 And
            a.医生姓名 || '' = v_操作员姓名 And Nvl(b.急诊, 0) = 1 And b.发生时间 >= Sysdate - n_急诊有效天数;
    End If;
  Else
    If n_呼叫含回诊 <> 1 Then
      Select Count(Distinct b.Id) As Count
      Into n_呼叫人数
      From 病人挂号记录 B, 排队叫号队列 A
      Where a.业务id = b.Id And a.业务类型 = 0 And Instr(',0,4,', ',' || a.排队状态 || ',') = 0 And b.记录性质 = 1 And b.记录状态 = 1 And
            a.医生姓名 || '' = v_操作员姓名 And Nvl(a.回诊序号, 0) = 0 And ((Nvl(b.急诊, 0) = 1 And b.发生时间 >= Sysdate - n_急诊有效天数) Or
            (Nvl(b.急诊, 0) <> 1 And b.发生时间 >= Sysdate - n_挂号有效天数));
    Else
      Select Count(Distinct b.Id) As Count
      Into n_呼叫人数
      From 病人挂号记录 B, 排队叫号队列 A
      Where a.业务id = b.Id And a.业务类型 = 0 And Instr(',0,4,6,', ',' || a.排队状态 || ',') = 0 And b.记录性质 = 1 And b.记录状态 = 1 And
            a.医生姓名 || '' = v_操作员姓名 And ((Nvl(b.急诊, 0) = 1 And b.发生时间 >= Sysdate - n_急诊有效天数) Or
            (Nvl(b.急诊, 0) <> 1 And b.发生时间 >= Sysdate - n_挂号有效天数));
    End If;
  End If;

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '成功');
  zlJsonPutValue(v_Output, 'call_count', n_呼叫人数, 1, 2);
  Json_Out := '{"output":' || v_Output || '}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getqueuecallcount;
/


Create Or Replace Procedure Zl_Exsesvr_Checkqueuedate
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --功能：对此挂号所在的队列进行有效检查
  --入参：Json_In:格式
  --input
  --  reg_id       N  1 挂号id
  -- 出参:
  --  output
  --    code                            N 1 应答吗：0-失败；1-成功
  --    message                         C 1 应答消息：失败时返回具体的错误信息
  --    result                          C 1 返回值:处理类型|提示信息
  -------------------------------------------
  j_Input  PLJson;
  j_Json   PLJson;
  n_挂号id Number;
  v_Out    Varchar2(500);
  v_Output Varchar2(32767);
Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_挂号id := j_Json.Get_Number('reg_id');

  v_Out := Zl_Queuedatecheck(n_挂号id);

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '成功');
  zlJsonPutValue(v_Output, 'result', v_Out, 0, 2);
  Json_Out := '{"output":' || v_Output || '}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkqueuedate;
/


Create Or Replace Procedure Zl_Exsesvr_Getqueuereginfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  -------------------------------------------
  --功能：根据条件获取排队叫号队列信息
  --入参：Json_In:格式
  --input
  --  query_type        N    1 调用类型 ：1-通过业务ids获取排队的挂号信息,2-通过挂号单获取排队的挂号信息，3-通过队列名称查询
  --  business_ids      C      业务ids
  --  reg_no            C      挂号单
  --  queue_name        C      队列名称
  --  queue_state       C      排队状态
  --出参      json
  --output
  --    code                N   1 应答吗：0-失败；1-成功
  --    message             C   1 应答消息：失败时返回具体的错误信息
  --    queue_list   排队叫号列表，支持多个，[数组]
  --       reg_id           N   1 挂号id
  --       reg_no           C   1 挂号单
  --       pati_id          N   1 病人id
  --       exec_deptid      N   1 执行部门id
  --       exec_state       N   1 执行状态
  --       outp_room        C   1 诊室
  --       queue_num        C   1 排队号码
  j_Input  PLJson;
  j_Json   PLJson;
  v_Output Varchar2(32767);
  c_Output Clob;

  n_Type     Number(18);
  v_业务ids  Varchar(32767);
  v_状态     Varchar(4000);
  v_挂号单   Varchar(50);
  v_队列名称 Varchar(200);
Begin

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_Type := j_Json.Get_Number('query_type');
  If n_Type = 1 Then
    v_业务ids := j_Json.Get_String('business_ids');
    v_状态    := j_Json.Get_String('queue_state');
  
    If v_业务ids Is Null Then
      Json_Out := zlJsonOut('未传入业务id');
      Return;
    End If;
  
    For c_队列信息 In (Select a.Id, a.No, a.病人id, a.执行部门id, a.执行状态, b.诊室, b.排队号码
                   From 病人挂号记录 A, 排队叫号队列 B
                   Where a.Id = b.业务id And b.业务类型 = 0 And a.Id In (Select Column_Value From Table(f_Str2List(v_业务ids))) And
                         (Instr(',' || v_状态 || ',', ',' || Nvl(b.排队状态, 0) || ',') > 0 Or v_状态 Is Null) And
                         a.记录性质 In (1, 2) And a.记录状态 = 1) Loop
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      
        v_Output := Null;
      End If;
    
      zlJsonPutValue(v_Output, 'reg_id', c_队列信息.Id, 1, 1);
      zlJsonPutValue(v_Output, 'reg_no', c_队列信息.No);
      zlJsonPutValue(v_Output, 'pati_id', c_队列信息.病人id, 1);
      zlJsonPutValue(v_Output, 'exec_deptid', c_队列信息.执行部门id, 1);
      zlJsonPutValue(v_Output, 'exec_state', c_队列信息.执行状态, 1);
      zlJsonPutValue(v_Output, 'outp_room', c_队列信息.诊室);
      zlJsonPutValue(v_Output, 'queue_num', c_队列信息.排队号码, 0, 2);
    
    End Loop;
  Elsif n_Type = 2 Then
    v_挂号单 := j_Json.Get_String('reg_no');
    v_状态   := j_Json.Get_String('queue_state');
  
    If v_挂号单 Is Null Then
      Json_Out := zlJsonOut('未传入挂号单');
      Return;
    End If;
  
    For c_队列信息 In (Select a.Id, a.No, a.病人id, a.执行部门id, a.执行状态, b.诊室, b.排队号码
                   From 病人挂号记录 A, 排队叫号队列 B
                   Where a.Id = b.业务id And b.业务类型 = 0 And a.No = v_挂号单 And
                         (Instr(',' || v_状态 || ',', ',' || Nvl(b.排队状态, 0) || ',') > 0 Or v_状态 Is Null) And
                         a.记录性质 In (1, 2) And a.记录状态 = 1) Loop
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      End If;
      zlJsonPutValue(v_Output, 'reg_id', c_队列信息.Id, 1, 1);
      zlJsonPutValue(v_Output, 'reg_no', c_队列信息.No);
      zlJsonPutValue(v_Output, 'pati_id', c_队列信息.病人id, 1);
      zlJsonPutValue(v_Output, 'exec_deptid', c_队列信息.执行部门id, 1);
      zlJsonPutValue(v_Output, 'exec_state', c_队列信息.执行状态, 1);
      zlJsonPutValue(v_Output, 'outp_room', c_队列信息.诊室);
      zlJsonPutValue(v_Output, 'queue_num', c_队列信息.排队号码, 0, 2);
    
    End Loop;
  Elsif n_Type = 3 Then
    v_队列名称 := j_Json.Get_String('queue_name');
    v_状态     := j_Json.Get_String('queue_state');
  
    For c_队列信息 In (Select Distinct /*+ Rule*/ a.Id, a.No, a.病人id, a.执行部门id, a.执行状态, b.诊室, b.排队号码
                   From 病人挂号记录 A, 排队叫号队列 B
                   Where a.Id = b.业务id And b.队列名称 = v_队列名称 And Nvl(b.业务类型, 0) = 0 And
                         (Instr(',' || v_状态 || ',', ',' || Nvl(b.排队状态, 0) || ',') > 0 Or v_状态 Is Null) And
                         Nvl(a.病人id, 0) = 0 And a.记录性质 = 1 And a.记录状态 = 1) Loop
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      End If;
      zlJsonPutValue(v_Output, 'reg_id', c_队列信息.Id, 1, 1);
      zlJsonPutValue(v_Output, 'reg_no', c_队列信息.No);
      zlJsonPutValue(v_Output, 'pati_id', c_队列信息.病人id, 1);
      zlJsonPutValue(v_Output, 'exec_deptid', c_队列信息.执行部门id, 1);
      zlJsonPutValue(v_Output, 'exec_state', c_队列信息.执行状态, 1);
      zlJsonPutValue(v_Output, 'outp_room', c_队列信息.诊室);
      zlJsonPutValue(v_Output, 'queue_num', c_队列信息.排队号码, 0, 2);
    End Loop;
  End If;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"成功","queue_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"成功","queue_list":[' || v_Output || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getqueuereginfo;
/


CREATE OR REPLACE Procedure Zl_Exsesvr_Getqueuereglist
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  -------------------------------------------
  --功能：根据条件获取排队叫号的候诊病人列表
  --入参：Json_In:格式
  --input
  --  outp_room_names        C    1 门诊诊室字符串       以逗号分隔
  --  outp_dr_names          C    1 门诊医生字符串       以逗号分隔
  --  recipe_exe_status      C    1 执行状态字符串       以逗号分隔
  --  outpque_names          C    1 排队队列名称字符串   以逗号分隔
  --  view_type              N    1 排队分组显示类型     1-按队列分组 2-按医生姓名分组  3-按诊室分组

  --出参      json
  --output
  --    code                N   1 应答吗：0-失败；1-成功
  --    message             C   1 应答消息：失败时返回具体的错误信息
  --    pati_list           病人列表，支持多个，[数组]
  --       outpque_id             N   1 排队ID
  --       pati_id                N   1 病人id
  --       outpque_name           C   1 排队队列名称
  --       outpque_sno            N   1 排队序号
  --       business_type          C   1 业务类型
  --       business_id            N   1 业务id
  --       dept_id                N   1 科室id
  --       dept_name              C   1 部门名称
  --       outpque_no             C   1 排队号码
  --       outpque_sign           N   1 排队标志
  --       pati_name              C   1 病人姓名
  --       pati_age               C   1 病人年龄
  --       outp_room_name         C   1 门诊诊室
  --       outp_dr_name           C   1 门诊医生
  --       call_dr_name           C   1 呼叫医生
  --       outpat_pri             N   1 优先
  --       revisit_sno            N   1 回诊序号
  --       outpque_time           C   1 排队时间
  --       call_time              C   1 呼叫时间
  --       outpque_state          N   1 排队状态
  --       outpque_revisit_num    N   1 回诊排序号

  j_Input  PLJson;
  j_Json   PLJson;
  v_Output Varchar2(32767);
  c_Output Clob;

  v_队列名称字符串 Varchar2(32767);
  n_队列过滤       Number;
  v_门诊医生字符串 Varchar(4000);
  v_门诊诊室字符串 Varchar(4000);
  v_执行状态字符串 Varchar(4000);
  n_显示类型       Number;
  n_回诊病人优先   Number;

  l_队列名称 t_StrList := t_StrList();

Begin

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_门诊医生字符串 := j_Json.Get_String('outp_dr_names');
  v_门诊诊室字符串 := j_Json.Get_String('outp_room_names');
  v_执行状态字符串 := j_Json.Get_String('recipe_exe_status');
  n_显示类型       := Nvl(j_Json.Get_Number('view_type'), 0);
  n_回诊病人优先   := Nvl(To_Number(zl_GetSysParameter('回诊病人是否优先', 1160)), 1);
  v_队列名称字符串 := j_Json.Get_String('outpque_names');

  If v_队列名称字符串 Is Not Null Then
    n_队列过滤 := 1;
  End If;

  While v_队列名称字符串 Is Not Null Loop
    If Length(v_队列名称字符串) <= 4000 Then
      l_队列名称.Extend;
      l_队列名称(l_队列名称.Count) := v_队列名称字符串;
      v_队列名称字符串 := Null;
    Else
      l_队列名称.Extend;
      l_队列名称(l_队列名称.Count) := Substr(v_队列名称字符串, 1, Instr(v_队列名称字符串, ',', 3940) - 1);
      v_队列名称字符串 := Substr(v_队列名称字符串, Instr(v_队列名称字符串, ',', 3940) + 1);
    End If;
  End Loop;

  For I In 1 .. l_队列名称.Count Loop

    If n_显示类型 = 1 Then
      For c_病人列表 In (Select /*+ Rule*/
                      To_Number(a.Id) As ID, To_Number(a.病人id) As 病人id, a.队列名称, a.排队序号, To_Number(a.业务类型) As 业务类型,
                      To_Number(a.业务id) As 业务id, To_Number(a.科室id) As 科室id, x.名称 As 部门名称, a.排队号码, a.排队标记,
                      a.患者姓名 || Decode(e.预约, 1, '(预)', Null) As 患者姓名, e.年龄, a.诊室, a.医生姓名,
                      (Select j.姓名 From 人员表 J, 上机人员表 K Where j.Id = k.人员id And k.用户名 = a.呼叫医生) As 呼叫医生,
                      To_Number(a.优先) As 优先, To_Number(a.回诊序号) As 回诊序号, To_Char(a.排队时间, 'yyyy-mm-dd hh24:mi:ss') As 排队时间,
                      To_Char(a.呼叫时间, 'yyyy-mm-dd hh24:mi:ss') As 呼叫时间, To_Number(a.排队状态) As 排队状态,
                      Decode(n_回诊病人优先, 1, To_Number(Nvl(a.回诊序号, 9999999999)), 0) As 回诊排序号
                     From 排队叫号队列 A, 部门表 X,
                          (Select /*+cardinality(f,10)*/
                             Column_Value As 队列名称
                            From Table(f_Str2List(l_队列名称(I))) F) B,
                          Table(Cast(f_Str2List(v_门诊诊室字符串) As Zltools.t_Strlist)) C,
                          Table(Cast(f_Str2List(v_门诊医生字符串) As Zltools.t_Strlist)) D, 病人挂号记录 E
                     Where To_Number(a.业务id) = e.Id And
                           (Nvl(a.是否分时点, 0) = 0 And a.排队时间 <= Trunc(Sysdate + 1) - 1 / 24 / 60 / 60 Or
                            Nvl(a.是否分时点, 0) = 1 And Sysdate > a.排队时间) And (a.队列名称 = b.队列名称 Or n_队列过滤 Is Null) And
                           (v_执行状态字符串 Is Null Or Instr(v_执行状态字符串, a.排队状态) = 0) And x.Id = a.科室id And
                           ((a.诊室 = c.Column_Value And a.医生姓名 Is Null) Or a.医生姓名 = d.Column_Value Or
                            (a.诊室 Is Null And a.医生姓名 Is Null)) And Nvl(a.排队状态, 0) <> 8
                     Order By 排队状态 Desc, 排队序号, 优先 Desc, 回诊排序号, 排队时间, 排队号码) Loop
        If Length(Nvl(v_Output, ' ')) > 30000 Then
          If c_Output Is Not Null Then
            c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
          Else
            c_Output := To_Clob(v_Output);
          End If;

          v_Output := Null;
        End If;
        zlJsonPutValue(v_Output, 'outpque_id', c_病人列表.Id, 1, 1);
        zlJsonPutValue(v_Output, 'pati_id', c_病人列表.病人id, 1);
        zlJsonPutValue(v_Output, 'outpque_name', c_病人列表.队列名称);
        zlJsonPutValue(v_Output, 'outpque_sno', c_病人列表.排队序号, 1);
        zlJsonPutValue(v_Output, 'business_type', c_病人列表.业务类型);
        zlJsonPutValue(v_Output, 'business_id', c_病人列表.业务id, 1);
        zlJsonPutValue(v_Output, 'dept_id', c_病人列表.科室id, 1);
        zlJsonPutValue(v_Output, 'dept_name', c_病人列表.部门名称);
        zlJsonPutValue(v_Output, 'outpque_no', c_病人列表.排队号码);
        zlJsonPutValue(v_Output, 'outpque_sign', c_病人列表.排队标记);
        zlJsonPutValue(v_Output, 'pati_name', c_病人列表.患者姓名);
        zlJsonPutValue(v_Output, 'pati_age', c_病人列表.年龄);
        zlJsonPutValue(v_Output, 'outp_room_name', c_病人列表.诊室);
        zlJsonPutValue(v_Output, 'outp_dr_name', c_病人列表.医生姓名);
        zlJsonPutValue(v_Output, 'call_dr_name', c_病人列表.呼叫医生);
        zlJsonPutValue(v_Output, 'outpat_pri', c_病人列表.优先, 1);
        zlJsonPutValue(v_Output, 'revisit_sno', c_病人列表.回诊序号, 1);
        zlJsonPutValue(v_Output, 'outpque_time', c_病人列表.排队时间);
        zlJsonPutValue(v_Output, 'call_time', c_病人列表.呼叫时间);
        zlJsonPutValue(v_Output, 'outpque_state', c_病人列表.排队状态, 1);
        zlJsonPutValue(v_Output, 'outpque_revisit_num', c_病人列表.回诊排序号, 1, 2);

      End Loop;
    Elsif n_显示类型 = 2 Then
      For c_病人列表 In (Select /*+ Rule*/
                      To_Number(a.Id) As ID, To_Number(a.病人id) As 病人id, a.队列名称, a.排队序号, To_Number(a.业务类型) As 业务类型,
                      To_Number(a.业务id) As 业务id, To_Number(a.科室id) As 科室id, x.名称 As 部门名称, a.排队号码, a.排队标记,
                      a.患者姓名 || Decode(e.预约, 1, '(预)', Null) As 患者姓名, e.年龄, a.诊室, a.医生姓名,
                      (Select j.姓名 From 人员表 J, 上机人员表 K Where j.Id = k.人员id And k.用户名 = a.呼叫医生) As 呼叫医生,
                      To_Number(a.优先) As 优先, To_Number(a.回诊序号) As 回诊序号, To_Char(a.排队时间, 'yyyy-mm-dd hh24:mi:ss') As 排队时间,
                      To_Char(a.呼叫时间, 'yyyy-mm-dd hh24:mi:ss') As 呼叫时间, To_Number(a.排队状态) As 排队状态,
                      Decode(n_回诊病人优先, 1, To_Number(Nvl(a.回诊序号, 9999999999)), 0) As 回诊排序号
                     From 排队叫号队列 A, 部门表 X,
                          (Select /*+cardinality(f,10)*/
                             Column_Value As 队列名称
                            From Table(f_Str2List(l_队列名称(I))) F) B,
                          Table(Cast(f_Str2List(v_门诊诊室字符串) As Zltools.t_Strlist)) C,
                          Table(Cast(f_Str2List(v_门诊医生字符串) As Zltools.t_Strlist)) D, 病人挂号记录 E
                     Where To_Number(a.业务id) = e.Id And
                           (Nvl(a.是否分时点, 0) = 0 And a.排队时间 <= Trunc(Sysdate + 1) - 1 / 24 / 60 / 60 Or
                            Nvl(a.是否分时点, 0) = 1 And Sysdate > a.排队时间) And (a.队列名称 = b.队列名称 Or n_队列过滤 Is Null) And
                           (v_执行状态字符串 Is Null Or Instr(v_执行状态字符串, a.排队状态) = 0) And x.Id = a.科室id And
                           (a.诊室 = c.Column_Value And (a.医生姓名 Is Null Or a.医生姓名 = d.Column_Value)) And
                           Nvl(a.排队状态, 0) <> 8
                     Order By 排队状态 Desc, 排队序号, 优先 Desc, 回诊排序号, 排队时间, 排队号码) Loop
        If Length(Nvl(v_Output, ' ')) > 30000 Then
          If c_Output Is Not Null Then
            c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
          Else
            c_Output := To_Clob(v_Output);
          End If;

          v_Output := Null;
        End If;

        zlJsonPutValue(v_Output, 'outpque_id', c_病人列表.Id, 1, 1);
        zlJsonPutValue(v_Output, 'pati_id', c_病人列表.病人id, 1);
        zlJsonPutValue(v_Output, 'outpque_name', c_病人列表.队列名称);
        zlJsonPutValue(v_Output, 'outpque_sno', c_病人列表.排队序号, 1);
        zlJsonPutValue(v_Output, 'business_type', c_病人列表.业务类型);
        zlJsonPutValue(v_Output, 'business_id', c_病人列表.业务id, 1);
        zlJsonPutValue(v_Output, 'dept_id', c_病人列表.科室id, 1);
        zlJsonPutValue(v_Output, 'dept_name', c_病人列表.部门名称);
        zlJsonPutValue(v_Output, 'outpque_no', c_病人列表.排队号码);
        zlJsonPutValue(v_Output, 'outpque_sign', c_病人列表.排队标记);
        zlJsonPutValue(v_Output, 'pati_name', c_病人列表.患者姓名);
        zlJsonPutValue(v_Output, 'pati_age', c_病人列表.年龄);
        zlJsonPutValue(v_Output, 'outp_room_name', c_病人列表.诊室);
        zlJsonPutValue(v_Output, 'outp_dr_name', c_病人列表.医生姓名);
        zlJsonPutValue(v_Output, 'call_dr_name', c_病人列表.呼叫医生);
        zlJsonPutValue(v_Output, 'outpat_pri', c_病人列表.优先, 1);
        zlJsonPutValue(v_Output, 'revisit_sno', c_病人列表.回诊序号, 1);
        zlJsonPutValue(v_Output, 'outpque_time', c_病人列表.排队时间);
        zlJsonPutValue(v_Output, 'call_time', c_病人列表.呼叫时间);
        zlJsonPutValue(v_Output, 'outpque_state', c_病人列表.排队状态, 1);
        zlJsonPutValue(v_Output, 'outpque_revisit_num', c_病人列表.回诊排序号, 1, 2);

      End Loop;
    Elsif n_显示类型 = 3 Then
      For c_病人列表 In (Select /*+ Rule*/
                      To_Number(a.Id) As ID, To_Number(a.病人id) As 病人id, a.队列名称, a.排队序号, To_Number(a.业务类型) As 业务类型,
                      To_Number(a.业务id) As 业务id, To_Number(a.科室id) As 科室id, x.名称 As 部门名称, a.排队号码, a.排队标记,
                      a.患者姓名 || Decode(e.预约, 1, '(预)', Null) As 患者姓名, e.年龄, a.诊室, a.医生姓名,
                      (Select j.姓名 From 人员表 J, 上机人员表 K Where j.Id = k.人员id And k.用户名 = a.呼叫医生) As 呼叫医生,
                      To_Number(a.优先) As 优先, To_Number(a.回诊序号) As 回诊序号, To_Char(a.排队时间, 'yyyy-mm-dd hh24:mi:ss') As 排队时间,
                      To_Char(a.呼叫时间, 'yyyy-mm-dd hh24:mi:ss') As 呼叫时间, To_Number(a.排队状态) As 排队状态,
                      Decode(n_回诊病人优先, 1, To_Number(Nvl(a.回诊序号, 9999999999)), 0) As 回诊排序号
                     From 排队叫号队列 A, 部门表 X,
                          (Select /*+cardinality(f,10)*/
                             Column_Value As 队列名称
                            From Table(f_Str2List(l_队列名称(I))) F) B, Table(f_Str2List(v_门诊医生字符串)) D, 病人挂号记录 E
                     Where To_Number(a.业务id) = e.Id And
                           (Nvl(a.是否分时点, 0) = 0 And a.排队时间 <= Trunc(Sysdate + 1) - 1 / 24 / 60 / 60 Or
                            Nvl(a.是否分时点, 0) = 1 And Sysdate > a.排队时间) And (a.队列名称 = b.队列名称 Or n_队列过滤 Is Null) And
                           (v_执行状态字符串 Is Null Or Instr(v_执行状态字符串, a.排队状态) = 0) And x.Id = a.科室id And
                           a.医生姓名 = d.Column_Value And Nvl(a.排队状态, 0) <> 8
                     Order By 排队状态 Desc, 排队序号, 优先 Desc, 回诊排序号, 排队时间, 排队号码) Loop

        If Length(Nvl(v_Output, ' ')) > 30000 Then
          If c_Output Is Not Null Then
            c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
          Else
            c_Output := To_Clob(v_Output);
          End If;

          v_Output := Null;
        End If;

        zlJsonPutValue(v_Output, 'outpque_id', c_病人列表.Id, 1, 1);
        zlJsonPutValue(v_Output, 'pati_id', c_病人列表.病人id, 1);
        zlJsonPutValue(v_Output, 'outpque_name', c_病人列表.队列名称);
        zlJsonPutValue(v_Output, 'outpque_sno', c_病人列表.排队序号, 1);
        zlJsonPutValue(v_Output, 'business_type', c_病人列表.业务类型);
        zlJsonPutValue(v_Output, 'business_id', c_病人列表.业务id, 1);
        zlJsonPutValue(v_Output, 'dept_id', c_病人列表.科室id, 1);
        zlJsonPutValue(v_Output, 'dept_name', c_病人列表.部门名称);
        zlJsonPutValue(v_Output, 'outpque_no', c_病人列表.排队号码);
        zlJsonPutValue(v_Output, 'outpque_sign', c_病人列表.排队标记);
        zlJsonPutValue(v_Output, 'pati_name', c_病人列表.患者姓名);
        zlJsonPutValue(v_Output, 'pati_age', c_病人列表.年龄);
        zlJsonPutValue(v_Output, 'outp_room_name', c_病人列表.诊室);
        zlJsonPutValue(v_Output, 'outp_dr_name', c_病人列表.医生姓名);
        zlJsonPutValue(v_Output, 'call_dr_name', c_病人列表.呼叫医生);
        zlJsonPutValue(v_Output, 'outpat_pri', c_病人列表.优先, 1);
        zlJsonPutValue(v_Output, 'revisit_sno', c_病人列表.回诊序号, 1);
        zlJsonPutValue(v_Output, 'outpque_time', c_病人列表.排队时间);
        zlJsonPutValue(v_Output, 'call_time', c_病人列表.呼叫时间);
        zlJsonPutValue(v_Output, 'outpque_state', c_病人列表.排队状态, 1);
        zlJsonPutValue(v_Output, 'outpque_revisit_num', c_病人列表.回诊排序号, 1, 2);

      End Loop;
    End If;
  End Loop;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"成功","pati_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"成功","pati_list":[' || v_Output || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getqueuereglist;
/

 

Create Or Replace Procedure Zl_Exsesvr_Updateregstate
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --功能：更新病人就诊状态
  --入参：Json_In:格式
  --input
  --  reg_no             C  1 挂号单
  --  exe_status         N  1 执行状态 0:标记为待诊.-1:标记为不就诊
  -- 出参:
  --  output
  --    code                            N 1 应答吗：0-失败；1-成功
  --    message                         C 1 应答消息：失败时返回具体的错误信息
  -------------------------------------------
  v_挂号单   Varchar2(50);
  n_执行状态 Number;

  j_Input PLJson;
  j_Json  PLJson;

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_挂号单   := j_Json.Get_String('reg_no');
  n_执行状态 := Nvl(j_Json.Get_Number('exe_status'), 0);

  Zl_病人挂号记录_状态(v_挂号单, n_执行状态);

  Json_Out := zlJsonOut('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Updateregstate;
/

Create Or Replace Procedure Zl_Exsesvr_Updatereginfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --功能：更新病人挂号信息
  --入参：Json_In:格式
  --input
  --  reg_no    C  1 挂号no  通过挂号no来更新信息
  --  reg_id    N  1 挂号id  通过挂号id来更新信息
  --      update_list            更新挂号信息列表：只有一条
  --        pay_method     C   医疗付款方式
  --        fee_category   C   费别
  --        community_num  N   社区序号
  --        pati_name      C   姓名
  --        pati_sex       C   性别
  --        pati_age       C   年龄

  -- 出参:
  --  output
  --    code                            N 1 应答吗：0-失败；1-成功
  --    message                         C 1 应答消息：失败时返回具体的错误信息
  -------------------------------------------

  v_挂号单 Varchar2(50);
  n_挂号id Number;

  v_医疗付款方式 Varchar2(50);
  v_费别         Varchar2(50);

  n_社区序号 Number;
  v_姓名     病人挂号记录.姓名%Type;
  v_性别     病人挂号记录.性别%Type;
  v_年龄     病人挂号记录.年龄%Type;

  n_医疗付款方式_b Number(1);
  n_费别_b         Number(1);

  n_社区序号_b Number(1);
  n_姓名_b     Number(1);
  n_性别_b     Number(1);
  n_年龄_b     Number(1);

  j_Input PLJson;
  j_Json  PLJson;
  o_Json  PLJson;

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_挂号id := Nvl(j_Json.Get_Number('reg_id'), 0);
  v_挂号单 := j_Json.Get_String('reg_no');
  o_Json   := j_Json.Get_Pljson('update_list');

  If o_Json.Exist('fee_category') Then
    v_费别   := o_Json.Get_String('fee_category');
    n_费别_b := 1;
  End If;

  If o_Json.Exist('pay_method') Then
    v_医疗付款方式   := o_Json.Get_String('pay_method');
    n_医疗付款方式_b := 1;
  End If;

  If o_Json.Exist('community_num') Then
    n_社区序号   := o_Json.Get_Number('community_num');
    n_社区序号_b := 1;
  End If;

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

  If n_挂号id <> 0 Then
    Update 病人挂号记录
    Set 医疗付款方式 = Decode(n_医疗付款方式_b, 1, v_医疗付款方式, 医疗付款方式), 费别 = Decode(n_费别_b, 1, v_费别, 费别),
        社区 = Decode(n_社区序号_b, 1, n_社区序号, 社区), 姓名 = Decode(n_姓名_b, 1, v_姓名, 姓名), 性别 = Decode(n_性别_b, 1, v_性别, 性别),
        年龄 = Decode(n_年龄_b, 1, v_年龄, 年龄)
    Where ID = n_挂号id;
  Else
    Update 病人挂号记录
    Set 医疗付款方式 = Decode(n_医疗付款方式_b, 1, v_医疗付款方式, 医疗付款方式), 费别 = Decode(n_费别_b, 1, v_费别, 费别),
        社区 = Decode(n_社区序号_b, 1, n_社区序号, 社区), 姓名 = Decode(n_姓名_b, 1, v_姓名, 姓名), 性别 = Decode(n_性别_b, 1, v_性别, 性别),
        年龄 = Decode(n_年龄_b, 1, v_年龄, 年龄)
    Where NO = v_挂号单;
  End If;
  Json_Out := zlJsonOut('成功', 1);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Updatereginfo;
/

Create Or Replace Procedure Zl_Exsesvr_Updateregroom
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --功能：病人挂号记录更新诊室
  --入参：Json_In:格式
  --input
  --  reg_no             C  1 挂号no
  --  pati_id            N  1 病人id
  --  outp_room          C  1 诊室
  --  outpat_dr          C  1 医生
  --  outpat_trg_time    C  1 分诊时间
  --  update_room        N  1 更新诊室
  --  appt_mode          C  1 预约方式
  -- 出参:
  --  output
  --    code                            N 1 应答吗：0-失败；1-成功
  --    message                         C 1 应答消息：失败时返回具体的错误信息
  -------------------------------------------

  v_挂号单   Varchar2(50);
  n_病人id   Number;
  v_诊室     Varchar2(50);
  v_医生     Varchar2(50);
  d_分诊时间 Date;
  n_更新诊室 Number;
  v_预约方式 Varchar2(50);

  j_Input PLJson;
  j_Json  PLJson;

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_挂号单   := j_Json.Get_String('reg_no');
  n_病人id   := j_Json.Get_Number('pati_id');
  v_诊室     := j_Json.Get_String('outp_room');
  v_医生     := j_Json.Get_String('outpat_dr');
  d_分诊时间 := To_Date(j_Json.Get_String('outpat_trg_time'), 'yyyy-mm-dd hh24:mi:ss');
  n_更新诊室 := j_Json.Get_Number('update_room');
  v_预约方式 := j_Json.Get_String('appt_mode');

  Zl_病人挂号记录_更新诊室_s(v_挂号单, n_病人id, v_诊室, v_医生, d_分诊时间, n_更新诊室, v_预约方式);

  Json_Out := zlJsonOut('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Updateregroom;
/

Create Or Replace Procedure Zl_Exsesvr_Getbillstatubyno
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:获取指定单据的收费、异常、结帐等状态
  --入参：Json_In:格式
  --   input      
  --    fee_no  C 1 单据号
  --    bill_prop N 1 记录性质:1-收费单;2-记帐单;3-自动记帐单;4-挂号单;5-就诊卡;6-预交单
  --出参: Json_Out,格式如下
  -- output      
  --   code  C 1 应答码：0-失败；1-成功
  --   message C 1 应答消息：失败时返回具体的错误信息
  --  statu N 1 收费状态:0-未收费或划价;1-已收费或已记帐;2-已全退或全销帐;3-部分退费或部分销帐
  --  err_sign  N 1 异常标志:0-正常数据;1-收款发生异常;2-退款发生异常
  --  blnc_sign N 1 结帐标志:针对记帐单有效;0-未结帐;1-已经结帐
  --  consumeed N 1 预交是否发生消费:1-发生了消费;0-未发生消费
  ---------------------------------------------------------------------------
  j_Input      PLJson;
  j_Json       PLJson;
  v_单据号     Varchar2(100);
  v_No         门诊费用记录.No%Type;
  n_记录性质   Number(2);
  n_记录状态   Number(5);
  n_存在剩余数 Number(2);
  n_收费异常   Number(2);
  n_退费异常   Number(2);
  n_异常id     Number(18);
  n_结帐id     Number(18);
  n_收费状态   Number(2);
  n_Count      Number(2);
  n_结帐标志   Number(2);
  n_记帐费用   Number(2);
  n_校对标志   Number(2);
  Err_Item Exception;
  v_Err_Msg Varchar2(500);

  v_Output Varchar2(32767);

  --组装成功时返回的数据
  Function Get_Success_Message
  (
    收费状态_In     Number,
    异常标志_In     Number,
    结帐标志_In     Number,
    预交消费标志_In Number
  ) Return Varchar2 Is
  
  Begin
  
    v_Output := '';
    zlJsonPutValue(v_Output, 'code', 1, 1, 1);
    zlJsonPutValue(v_Output, 'message', '成功');
    zlJsonPutValue(v_Output, 'statu', 收费状态_In, 1); --收费状态:0-未收费或划价;1-已收费或已记帐;2-已全退或全销帐;3-部分退费或部分销帐
  
    zlJsonPutValue(v_Output, 'err_sign', 异常标志_In, 1); --异常标志:0-正常数据;1-收款发生异常;2-退款发生异常
    zlJsonPutValue(v_Output, 'blnc_sign', 结帐标志_In, 1); --结帐标志:针对记帐单有效;0-未结帐;1-已经结帐
    zlJsonPutValue(v_Output, 'consumeed_sign', 预交消费标志_In, 1, 2); --预交消费标志:1-发生了消费;0-未发生消费
    v_Output := '{"output":' || v_Output || '}';
  
    Return v_Output;
  End Get_Success_Message;

Begin
  --解析入参

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_单据号   := j_Json.Get_String('fee_no');
  n_记录性质 := Nvl(j_Json.Get_Number('bill_prop'), 0);

  If n_记录性质 <= 0 Or v_单据号 Is Null Then
    v_Err_Msg := '未传入必要的查询条件，不能获取费用相关单据的状态，请检查!';
    Json_Out  := zlJsonOut(v_Err_Msg);
    Return;
  End If;

  If n_记录性质 = 1 Then
    --获取收费单;
    Select Max(NO), Nvl(Max(记录状态), 0), Max(Decode(Nvl(剩余数量, 0), 0, 0, 1)), Max(收费异常), Max(退费异常), Max(异常id)
    Into v_No, n_记录状态, n_存在剩余数, n_收费异常, n_退费异常, n_异常id
    From (Select NO, Max(记录状态) As 记录状态, 序号, Sum(Nvl(付数, 1) * Nvl(数次, 0)) As 剩余数量,
                  Max(Decode(a.记录状态, 2, a.费用状态, 0)) As 退费异常, Max(Decode(a.记录状态, 2, 0, a.费用状态)) As 收费异常,
                  Max(Decode(a.记录状态, 2, Decode(Nvl(费用状态, 0), 1, a.结帐id, 0), 0)) As 异常id
           From 门诊费用记录 A
           Where Mod(记录性质, 10) = 1 And NO = v_单据号 And 价格父号 Is Null
           Group By NO, 序号);
  
    If v_No Is Null Then
      --未找到单据
      v_Err_Msg := '未找到收费单据为' || v_单据号 || '的收费单据，不能正常获取单据的状态，请检查!';
      Json_Out  := zlJsonOut(v_Err_Msg);
      Return;
    End If;
  
    If Nvl(n_记录状态, 0) = 0 Then
      --划价单:收费状态_in number,异常标志_in number,结帐标志_in number,预交消费标志_in number 
      Json_Out := Get_Success_Message(0, 0, 0, 0);
      Return;
    End If;
  
    If n_记录状态 = 1 Then
      --已经收费
      Json_Out := Get_Success_Message(1, n_收费异常, 0, 1);
      Return;
    End If;
  
    If Nvl(n_存在剩余数, 0) = 0 Then
    
      n_收费状态 := 2;
    Else
      n_收费状态 := 3;
    End If;
    --存在退费
    If Nvl(n_退费异常, 0) = 1 Then
      Select Nvl(Max(1), 0) Into n_退费异常 From 病人预交记录 Where 结帐id = n_异常id And Nvl(校对标志, 0) <> 0;
      If Nvl(n_退费异常, 0) = 1 Then
        --存在异常，肯定就没有
        n_收费状态 := 3;
      End If;
    End If;
    Json_Out := Get_Success_Message(n_收费状态, n_退费异常, 1, 0);
    Return;
  End If;

  If n_记录性质 = 2 Or n_记录性质 = 3 Then
    --2-记帐单;3-自动记帐单    
    Select Max(NO), Nvl(Max(记录状态), 0), Max(Decode(Nvl(剩余数量, 0), 0, 0, 1)), Max(结帐id)
    Into v_No, n_记录状态, n_存在剩余数, n_异常id
    From (Select NO, Max(记录状态) As 记录状态, 序号, Sum(Nvl(付数, 1) * Nvl(数次, 0)) As 剩余数量, Max(结帐id) As 结帐id
           From 门诊费用记录 A
           Where 记录性质 = n_记录性质 And NO = v_单据号 And 价格父号 Is Null
           Group By NO, 序号
           Union All
           Select NO, Max(记录状态) As 记录状态, 序号, Sum(Nvl(付数, 1) * Nvl(数次, 0)) As 剩余数量, Max(结帐id) As 结帐id
           From 住院费用记录 A
           Where 记录性质 = n_记录性质 And NO = v_单据号 And 价格父号 Is Null
           Group By NO, 序号
           
           );
    If v_No Is Null Then
      --未找到单据
      v_Err_Msg := '未找到记帐单据为' || v_单据号 || '的记帐单据，不能正常获取单据的状态，请检查!';
      Json_Out  := zlJsonOut(v_Err_Msg);
      Return;
    End If;
    n_结帐标志 := 0;
    If Nvl(n_记录状态, 0) = 0 Then
      --划价单:收费状态_in number,异常标志_in number,结帐标志_in number,预交消费标志_in number 
      Json_Out := Get_Success_Message(0, 0, 0, 0);
      Return;
    End If;
  
    If Nvl(n_异常id, 0) <> 0 Then
    
      Select Count(1)
      Into n_Count
      From 门诊费用记录 A
      Where a.记录状态 <> 0 And a.记帐费用 = 1 Having
       Sum(Nvl(a.实收金额, 0)) - Sum(Nvl(a.结帐金额, 0)) <> 0 Or
            (Sum(Nvl(a.实收金额, 0)) = 0 And Sum(Nvl(a.应收金额, 0)) <> 0 And Sum(Nvl(a.结帐金额, 0)) = 0 And Mod(Count(*), 2) = 0) Or
            Sum(Nvl(a.结帐金额, 0)) = 0 And Sum(Nvl(a.应收金额, 0)) <> 0 And Mod(Count(*), 2) = 0
      Group By a.No, a.序号, Mod(a.记录性质, 10), Nvl(a.执行状态, 0);
      If Nvl(n_Count, 0) = 0 Then
        --已经结帐
        n_结帐标志 := 1;
      Else
        --部分结帐的，也算未结帐
        n_结帐标志 := 0;
      End If;
    
    End If;
    If n_记录状态 = 1 Then
      --已经记帐
      Json_Out := Get_Success_Message(1, 0, n_结帐标志, 0);
      Return;
    End If;
  
    If Nvl(n_存在剩余数, 0) = 0 Then
      n_收费状态 := 2;
    Else
      n_收费状态 := 3;
    End If;
  
    Json_Out := Get_Success_Message(n_收费状态, 0, n_结帐标志, 0);
    Return;
  End If;
  If n_记录性质 = 4 Then
    --挂号单
    Select Max(NO), Nvl(Max(记录状态), 0), Max(Decode(Nvl(剩余数量, 0), 0, 0, 1)), Max(收费异常), Max(退费异常), Max(异常id)
    Into v_No, n_记录状态, n_存在剩余数, n_收费异常, n_退费异常, n_异常id
    From (Select NO, Max(记录状态) As 记录状态, 序号, Sum(Nvl(付数, 1) * Nvl(数次, 0)) As 剩余数量,
                  Max(Decode(a.记录状态, 2, a.费用状态, 0)) As 退费异常, Max(Decode(a.记录状态, 2, 0, a.费用状态)) As 收费异常,
                  Max(Decode(a.记录状态, 2, Decode(Nvl(a.费用状态, 0), 1, a.结帐id, 0), 0)) As 异常id
           From 门诊费用记录 A
           Where Mod(记录性质, 10) = 4 And NO = v_单据号 And 价格父号 Is Null
           Group By NO, 序号);
  
    If v_No Is Null Then
      --未找到单据
      v_Err_Msg := '未找到挂号单据为' || v_单据号 || '的挂号单据，不能正常获取单据的状态，请检查!';
      Json_Out  := zlJsonOut(v_Err_Msg);
      Return;
    End If;
    If Nvl(n_记录状态, 0) = 0 Then
      --划价单:收费状态_in number,异常标志_in number,结帐标志_in number,预交消费标志_in number 
      Json_Out := Get_Success_Message(0, 0, 0, 0);
      Return;
    End If;
  
    If n_记录状态 = 1 Then
      --已经收费
      Json_Out := Get_Success_Message(1, n_收费异常, 0, 0);
      Return;
    End If;
  
    If Nvl(n_存在剩余数, 0) = 0 Then
    
      n_收费状态 := 2;
    Else
      n_收费状态 := 3;
    End If;
    --存在退费
    If Nvl(n_退费异常, 0) = 1 Then
      Select Nvl(Max(1), 0) Into n_退费异常 From 病人预交记录 Where 结帐id = n_异常id And Nvl(校对标志, 0) <> 0;
      If Nvl(n_退费异常, 0) = 1 Then
        --存在异常，肯定就没有
        n_收费状态 := 3;
        n_退费异常 := 2;
      End If;
    End If;
    Json_Out := Get_Success_Message(n_收费状态, n_退费异常, 1, 0);
    Return;
    Null;
  End If;

  If n_记录性质 = 5 Then
    --医疗卡
    Select Max(NO), Nvl(Max(记录状态), 0), Max(Decode(Nvl(剩余数量, 0), 0, 0, 1)), Max(收费异常), Max(退费异常), Max(异常id), Max(记帐费用),
           Max(结帐id)
    Into v_No, n_记录状态, n_存在剩余数, n_收费异常, n_退费异常, n_异常id, n_记帐费用, n_结帐id
    From (Select NO, Max(记录状态) As 记录状态, 序号, Sum(Nvl(付数, 1) * Nvl(数次, 0)) As 剩余数量,
                  Max(Decode(a.记录状态, 2, a.费用状态, 0)) As 退费异常, Max(Decode(a.记录状态, 2, 0, a.费用状态)) As 收费异常,
                  Max(a.记帐费用) As 记帐费用, Max(Decode(a.记录状态, 2, Decode(Nvl(a.费用状态, 0), 1, a.结帐id, 0), 0)) As 异常id,
                  Max(结帐id) As 结帐id
           From 住院费用记录 A
           Where Mod(记录性质, 10) = 5 And NO = v_单据号 And 价格父号 Is Null
           Group By NO, 序号);
  
    If v_No Is Null Then
      --未找到单据
      v_Err_Msg := '未找到单据为' || v_单据号 || '的就诊卡单据，不能正常获取单据的状态，请检查!';
      Json_Out  := zlJsonOut(v_Err_Msg);
      Return;
    End If;
  
    If Nvl(n_记录状态, 0) = 0 Then
      --划价单:收费状态_in number,异常标志_in number,结帐标志_in number,预交消费标志_in number 
      Json_Out := Get_Success_Message(0, 0, 0, 0);
      Return;
    End If;
    n_结帐标志 := 0;
    If Nvl(n_记帐费用, 0) = 1 And Nvl(n_结帐id, 0) <> 0 Then
      --只要已经结过帐的，就返回已结帐
      n_结帐标志 := 1;
    End If;
    If n_记录状态 = 1 Then
      --已经收费
      Json_Out := Get_Success_Message(1, n_收费异常, n_结帐标志, 0);
      Return;
    End If;
  
    If Nvl(n_存在剩余数, 0) = 0 Then
    
      n_收费状态 := 2;
    Else
      n_收费状态 := 3;
    End If;
  
    --存在退费
    If Nvl(n_退费异常, 0) = 1 Then
      Select Nvl(Max(1), 0) Into n_退费异常 From 病人预交记录 Where 结帐id = n_异常id And Nvl(校对标志, 0) <> 0;
      If Nvl(n_退费异常, 0) = 1 Then
        --存在异常，肯定就没有
        n_收费状态 := 3;
        n_退费异常 := 2;
      End If;
    End If;
    --异常标志:0-正常数据;1-收款发生异常;2-退款发生异常
    Json_Out := Get_Success_Message(n_收费状态, n_退费异常, 1, 0);
    Return;
  End If;

  If n_记录性质 = 6 Then
    --预蛟记录
    Select Max(NO), Nvl(Max(Decode(记录性质, 11, 0, 记录状态)), 0), Decode(Nvl(Sum(冲预交), 0), 0, 0, 1),
           Decode(Nvl(Max(Decode(记录性质, 11, 0, 校对标志)), 0), 0, 0, 1)
    Into v_No, n_记录状态, n_Count, n_校对标志
    From 病人预交记录
    Where Mod(记录性质, 10) = 1 And NO = v_单据号;
  
    If v_No Is Null Then
      --未找到单据
      v_Err_Msg := '未找到单据为' || v_单据号 || '的预交单据，不能正常获取单据的状态，请检查!';
      Json_Out  := zlJsonOut(v_Err_Msg);
      Return;
    End If;
  
    If n_记录状态 = 0 Or n_记录状态 = 1 And Nvl(n_校对标志, 0) <> 0 Then
      --未生效的异常单据
      Json_Out := Get_Success_Message(0, 1, 0, n_Count);
      Return;
    End If;
  
    If n_记录状态 = 1 Then
      --已生效
      Json_Out := Get_Success_Message(1, 0, 0, n_Count);
      Return;
    End If;
    n_退费异常 := 0;
    If Nvl(n_校对标志, 0) <> 0 Then
      n_退费异常 := 2;
    End If;
    Json_Out := Get_Success_Message(2, n_退费异常, 0, n_Count);
    Return;
  
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getbillstatubyno;
/


Create Or Replace Procedure Zl_Exsesvr_Updatecardfeeinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  ---------------------------------------------------------------------------
  --功能：修正卡号对应的收费单据信息(主要是更新卡号及卡类别ID及票据信息等等)
  --入参      json
  --input
  --    oper_fun  N  1  操作标志:0-只修改卡费记录;1-修改费用记录及票据使用明细;2-病历费修改
  --    fee_no  C  1  费用单号：本次要调整的费用单据
  --    operator_name  C  1  操作员姓名
  --    operator_code  C  1  操作员编号
  --    create_time C 1 登记时间:yyyy-mm-dd hh:mi:ss
  --    sendcard_info     发卡信息
  --      send_mode N 1 发卡方式;0-发卡,1-补卡,2-换卡;3-退卡
  --      cardtype_id C 1 卡类别id
  --      cardno  C 1 卡号:本次发放或绑定或补卡的卡号
  --      recv_id N 1 领用id:票据领用ID(卡号)
  --      cardno_reusing  N 1 卡号重用:1-卡号允许重复使用用;0-不允许重复使用
  --      cardno_old  C 1 原卡卡号:换卡时，需要传入原卡号  --出参      json
  --output
  --  code                      C 1 应答码：0-失败；1-成功
  --  message                   C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  o_Json PLJson;

  n_操作标志   Number(2);
  v_费用单号   住院费用记录.No%Type;
  n_卡类别id   病人预交记录.卡类别id%Type;
  n_发卡方式   Number(2);
  n_卡号重用   Number(2);
  v_操作员编号 病人预交记录.操作员编号%Type;
  v_操作员姓名 病人预交记录.操作员姓名%Type;
  v_卡号       住院费用记录.实际票号%Type;
  v_原卡号     住院费用记录.实际票号%Type;
  n_领用id     Number(18);

  d_登记时间 Date;
  v_Err_Msg  Varchar2(255);
  Err_Item Exception;

Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_操作标志   := j_Json.Get_Number('oper_fun');
  v_费用单号   := j_Json.Get_String('fee_no');
  v_操作员姓名 := j_Json.Get_String('operator_name');
  v_操作员编号 := j_Json.Get_String('operator_code');
  d_登记时间   := To_Date(j_Json.Get_String('create_time'), 'YYYY-MM-DD hh24:mi:ss');

  If d_登记时间 Is Null Then
    d_登记时间 := Sysdate;
  End If;

  If d_登记时间 Is Null Then
    d_登记时间 := Sysdate;
  End If;

  If Nvl(n_操作标志, 2) = 2 Then
    --病历费修改
    --只能针对异常的修正
    Update 住院费用记录
    Set 操作员姓名 = v_操作员姓名, 操作员编号 = v_操作员编号, 登记时间 = d_登记时间
    Where Nvl(费用状态, 0) = 1 And 记录性质 = 5 And 记录状态 In (1, 3) And NO = v_费用单号;
    If Sql%NotFound Then
      v_Err_Msg := '未找到需要更新的就诊卡费用所涉及的单据信息';
      Raise Err_Item;
    End If;
  
    Json_Out := zlJsonOut('成功', 1);
    Return;
  End If;
  o_Json := j_Json.Get_Pljson('sendcard_info');
  If o_Json Is Null Then
    v_Err_Msg := '不能确定本次修改的卡片信息，请检查！';
    Raise Err_Item;
  End If;

  n_发卡方式 := Nvl(o_Json.Get_Number('send_mode'), 0); --发卡方式;;0-发卡,1-补卡,2-换卡;3-退卡
  n_卡类别id := o_Json.Get_Number('cardtype_id');
  n_卡号重用 := Nvl(o_Json.Get_Number('cardno_reusing'), 0);
  v_卡号     := o_Json.Get_String('cardno');
  v_原卡号   := o_Json.Get_String('cardno_old');
  n_领用id   := o_Json.Get_Number('recv_id');

  If Nvl(n_操作标志, 0) = 1 Then
    --票据使用修正
    Update 住院费用记录
    Set 实际票号 = v_卡号, 结论 = Nvl(n_卡类别id, 结论), 附加标志 = Decode(Nvl(附加标志, 0), 8, 8, n_发卡方式)
    Where 记录性质 = 5 And 记录状态 In (1, 3) And NO = v_费用单号;
    If Sql%NotFound Then
      v_Err_Msg := '未找到需要更新的就诊卡费用所涉及的单据信息';
      Raise Err_Item;
    End If;
    --发卡类型=1-发卡 ;2-换卡;3-补卡 ;4-退卡
    n_发卡方式 := Case
                When Nvl(n_发卡方式, 0) = 0 Then
                 1
                When Nvl(n_发卡方式, 0) = 1 Then
                 3
                When Nvl(n_发卡方式, 0) = 2 Then
                 2
                Else
                 4
              End;
  
    Zl_病人医疗卡票据_Update_s(n_发卡方式, v_卡号, v_操作员姓名, d_登记时间, v_费用单号, n_领用id, v_原卡号, n_卡号重用);
  
  Else
    --只能针对异常的修正
    Update 住院费用记录
    Set 实际票号 = v_卡号, 结论 = Nvl(n_卡类别id, 结论), 操作员姓名 = v_操作员姓名, 操作员编号 = v_操作员编号, 登记时间 = d_登记时间
    Where Nvl(费用状态, 0) = 1 And 记录性质 = 5 And 记录状态 In (1, 3) And NO = v_费用单号;
    If Sql%NotFound Then
      v_Err_Msg := '未找到需要更新的就诊卡费用所涉及的单据信息';
      Raise Err_Item;
    End If;
  End If;
  Json_Out := zlJsonOut('成功', 1);

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Updatecardfeeinfo;
/

Create Or Replace Procedure Zl_Exsesvr_Updatepatibaseinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --------------------------------------------------------------------------------------------------
  --功能:更新病人费用相关的病人基本信息
  --------------------------------------------------------------------------------------------------
  --入参 JSOM格式
  --input
  --  pati_id               N 1 病人id
  --  visit_id              N   主页id
  --  occasion              N   场合
  --  update_info           N   需要更新的信息
  --    pati_name             C   姓名
  --    outpatient_num        C   门诊号
  --    pati_age              C   年龄
  --    pati_sex              C 1 性别
  --    explain               C 1 说明
  --    regist_no             C 1 挂号单
  --    remark                C 1 摘要
  --出参 JSON格式
  --output
  --  code                  N 1 应答码：0-失败；1-成功
  --  message               C 1 应答消息：失败时返回具体的错误信息
  --  adjust_explain        C 1 修改说明
  j_Input PLJson;
  j_Json  PLJson;

  o_Json    PLJson;
  n_病人id  门诊费用记录.病人id%Type;
  v_姓名    门诊费用记录.姓名%Type;
  n_门诊号  门诊费用记录.标识号%Type;
  v_性别    门诊费用记录.性别%Type;
  v_年龄    门诊费用记录.年龄%Type; --更新前的年龄
  n_就诊id  Number;
  n_场合    Number(1);
  v_说明    Varchar2(4000);
  v_说明_In Varchar2(4000);
  说明_Out  Clob;
  v_摘要    Varchar2(3682);
  v_挂号单  Varchar2(100);

  Procedure p_费用
  (
    病人id_In 门诊费用记录.病人id%Type,
    就诊id_In Number,
    姓名_In   门诊费用记录.姓名%Type,
    性别_In   门诊费用记录.性别%Type,
    年龄_In   门诊费用记录.年龄%Type,
    场合_In   Number, --1-门诊;2-住院
    说明_In   Varchar2,
    说明_Out  Out Varchar2
  ) As
    ------------------------------------------------------------------------------------------
    --功能:更新费用相关业务数据的病人基本信息
    --入参:病人id_In:病人ID
    --     就诊id_In:门诊病人为挂号ID;住院病人为主页ID;为0说明是外来或非挂号就诊的病人(就诊id_In为空时,不更改该病人的费用部分的业务数据)
    --     姓名_In:需要更改的病人姓名
    --     性别_In:需要更改的病人性别
    --     年龄_In:需要更改的病人年龄
    --     场合_In:1-门诊;2-住院
    --出参:说明_Out:病人信息调整后的说明信息，用于提示操作员进行相关操作
    ------------------------------------------------------------------------------------------
    Err_Custom Exception;
    v_Error   Varchar2(2000);
    v_说明    Varchar2(4000);
    v_No      门诊费用记录.No%Type;
    n_类别    Number(2);
    v_Temp    Varchar2(4000);
    n_Split   Number(2);
    d_Maxdate Date;
    v_科室    部门表.名称%Type;
  Begin
    --没有指定的就诊ID，不更新病人的费用业务数据
    If Nvl(就诊id_In, 0) = 0 Then
      Return;
    End If;
    v_说明 := 说明_In;
    If Nvl(场合_In, 0) <= 1 Then
      Begin
        Select NO, 名称, 登记时间
        Into v_No, v_科室, d_Maxdate
        From 病人挂号记录 A, 部门表 B
        Where a.执行部门id = b.Id(+) And a.Id = 就诊id_In;
      Exception
        When Others Then
          v_No := Null;
      End;
      If v_No Is Null Then
        v_No := '-';
      End If;
    
      n_类别 := 0;
      v_Temp := Null;
      For c_费用 In (Select Distinct 1 As 性质, '挂号:' As 类别, c.号码
                   From 门诊费用记录 A, 票据打印内容 B, 票据使用明细 C
                   Where a.No = b.No And b.数据性质 = 4 And b.Id = c.打印id And c.性质 = 1 And a.病人id = 病人id_In And a.记录性质 = 4 And
                         a.No = v_No
                   Union All
                   Select Distinct 2 As 性质, '收费:' As 类别, c.号码
                   From 门诊费用记录 A, 票据打印内容 B, 票据使用明细 C
                   Where a.No = v_No And b.数据性质 = 1 And b.Id = c.打印id And c.性质 = 1 And a.病人id = 病人id_In And a.记录性质 = 1 And
                         (a.挂号id = 就诊id_In Or 医嘱序号 Is Null)
                   Union All
                   Select Distinct 3 As 性质, '医疗卡:' As 类别, c.号码
                   From 住院费用记录 A, 票据打印内容 B, 票据使用明细 C
                   Where a.No = b.No And b.数据性质 = 5 And b.Id = c.打印id And c.性质 = 1 And Nvl(a.记帐费用, 0) = 0 And
                         a.病人id = 病人id_In And a.记录性质 = 5
                   Union All
                   Select Distinct 4 As 性质, '预交:' As 类别, c.号码
                   From 病人预交记录 A, 票据打印内容 B, 票据使用明细 C
                   Where a.No = b.No And b.数据性质 = 2 And b.Id = c.打印id And c.性质 = 1 And Nvl(a.预交类别, 0) = 1 And
                         a.病人id = 病人id_In And a.记录性质 = 1
                   Union All
                   Select Distinct 5 As 性质, '结帐:' As 类别, c.号码
                   From (Select Distinct b.Id, b.No
                          From 门诊费用记录 A, 病人结帐记录 B
                          Where a.结帐id = b.Id And a.记帐费用 = 1 And a.病人id = 病人id_In And a.记录性质 In (2, 12) And
                                (a.挂号id = 就诊id_In Or 医嘱序号 Is Null)
                          Union All
                          Select Distinct b.Id, b.No
                          From 住院费用记录 A, 病人结帐记录 B
                          Where a.结帐id = b.Id And a.记帐费用 = 1 And a.病人id = 病人id_In And a.记录性质 = 5) A, 票据打印内容 B, 票据使用明细 C
                   Where a.No = b.No And b.数据性质 = 3 And b.Id = c.打印id And c.性质 = 1
                   Order By 性质, 号码) Loop
      
        If Length(Nvl(v_说明, '-') || Nvl(v_Temp, '-')) > 3800 Then
          v_说明 := v_说明 || '等';
          Exit;
        End If;
      
        If n_类别 <> Nvl(c_费用.性质, 0) Then
          If Not v_Temp Is Null Then
            v_说明 := Nvl(v_说明, '') || v_Temp;
          End If;
        
          n_Split := 1;
          If v_Temp Is Null Then
            v_Temp := c_费用.类别;
          Else
            v_Temp := ';' || c_费用.类别;
          End If;
          n_类别 := Nvl(c_费用.性质, 0);
        End If;
      
        If n_Split = 1 Then
          v_Temp := Nvl(v_Temp, '') || c_费用.号码;
        Else
          v_Temp := Nvl(v_Temp, '') || ',' || c_费用.号码;
        End If;
        n_Split := 0;
      End Loop;
      If Not v_Temp Is Null Then
        If Length(Nvl(v_说明, '-') || Nvl(v_Temp, '-')) > 4000 Then
          v_说明 := v_说明 || '等';
        Else
          v_说明 := Nvl(v_说明, '') || v_Temp;
        End If;
      
      End If;
      说明_Out := v_说明;
    
      --有诊疗ID的,只更新这次就诊的或直接登记的病人信息
      Update 门诊费用记录 A
      Set 姓名 = Nvl(姓名_In, 姓名), 性别 = Nvl(性别_In, 性别), 年龄 = Nvl(年龄_In, 年龄)
      Where 病人id = 病人id_In And a.记录性质 <> 4 And (a.挂号id = 就诊id_In Or 医嘱序号 Is Null);
    
      Update 门诊费用记录 A
      Set 姓名 = Nvl(姓名_In, 姓名), 性别 = Nvl(性别_In, 性别), 年龄 = Nvl(年龄_In, 年龄)
      Where 病人id = 病人id_In And a.记录性质 = 4 And NO = v_No;
    
      Update 住院费用记录 A
      Set 姓名 = Nvl(姓名_In, 姓名), 性别 = Nvl(性别_In, 性别), 年龄 = Nvl(年龄_In, 年龄)
      Where 病人id = 病人id_In And 记录性质 = 5;
    
      Update 排队叫号队列
      Set 患者姓名 = Nvl(姓名_In, 患者姓名)
      Where 病人id = 病人id_In And 业务类型 = 0 And 业务id = 就诊id_In;
      Return;
    End If;
  
    --住院:
    --1.结了账且打印了发票的,不允许更改
    n_类别 := 0;
    v_Temp := Null;
  
    For c_费用 In (Select Distinct 4 As 性质, '预交:' As 类别, c.号码
                 From 病人预交记录 A, 票据打印内容 B, 票据使用明细 C
                 Where a.No = b.No And b.数据性质 = 2 And b.Id = c.打印id And c.性质 = 1 And Nvl(a.预交类别, 0) = 2 And
                       a.病人id = 病人id_In And a.记录性质 = 1 And a.主页id = 就诊id_In
                 Union All
                 Select Distinct 5 As 性质, '结帐:' As 类别, c.号码
                 From (Select Distinct b.Id, b.No
                        From 住院费用记录 A, 病人结帐记录 B
                        Where a.结帐id = b.Id And Nvl(a.记帐费用, 0) = 1 And a.病人id = 病人id_In And a.主页id = 就诊id_In And
                              a.记录性质 <> 5) A, 票据打印内容 B, 票据使用明细 C
                 Where a.No = b.No And b.数据性质 = 3 And b.Id = c.打印id And c.性质 = 1
                 Order By 性质, 号码) Loop
    
      If Length(Nvl(v_说明, '-') || Nvl(v_Temp, '-')) > 3800 Then
        v_说明 := v_说明 || '等';
        Exit;
      End If;
    
      If n_类别 <> Nvl(c_费用.性质, 0) Then
        If Not v_Temp Is Null Then
          v_说明 := Nvl(v_说明, '') || v_Temp;
        End If;
        n_Split := 1;
        If v_Temp Is Null Then
          v_Temp := c_费用.类别;
        Else
          v_Temp := ';' || c_费用.类别;
        End If;
        n_类别 := Nvl(c_费用.性质, 0);
      End If;
    
      If n_Split = 1 Then
        v_Temp := Nvl(v_Temp, '') || c_费用.号码;
      Else
        v_Temp := Nvl(v_Temp, '') || ',' || c_费用.号码;
      End If;
      n_Split := 0;
    End Loop;
    If Not v_Temp Is Null Then
      If Length(Nvl(v_说明, '-') || Nvl(v_Temp, '-')) > 4000 Then
        v_说明 := v_说明 || '等';
      Else
        v_说明 := Nvl(v_说明, '') || v_Temp;
      End If;
    End If;
    说明_Out := v_说明;
  
    Update 住院费用记录
    Set 姓名 = Nvl(姓名_In, 姓名), 性别 = Nvl(性别_In, 性别), 年龄 = Nvl(年龄_In, 年龄)
    Where 病人id = 病人id_In And 主页id = 就诊id_In And 记录性质 <> 5;
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_费用;

Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_病人id := j_Json.Get_Number('pati_id');
  n_场合   := j_Json.Get_Number('occasion');
  n_就诊id := j_Json.Get_Number('visit_id');
  o_Json   := j_Json.Get_Pljson('update_info');
  If o_Json Is Null Then
    Json_Out := zlJsonOut('未传入需要更新的信息，请检查！', 0);
    Return;
  End If;

  v_姓名    := o_Json.Get_String('pati_name');
  n_门诊号  := To_Number(o_Json.Get_String('outpatient_num'));
  v_性别    := o_Json.Get_String('pati_sex');
  v_年龄    := o_Json.Get_String('pati_age');
  v_挂号单  := o_Json.Get_String('regist_no');
  v_说明_In := o_Json.Get_String('explain');
  v_摘要    := o_Json.Get_String('remark');

  If Nvl(v_挂号单, '-') <> '-' Then
    Update 门诊费用记录
    Set 标识号 = n_门诊号, 姓名 = v_姓名, 性别 = v_性别, 年龄 = v_年龄, 结论 = v_摘要
    Where NO = v_挂号单 And 记录性质 = 4;
  End If;

  If Nvl(n_病人id, 0) = 0 Then
    Json_Out := zlJsonOut('未传入病人id，请检查！', 0);
    Return;
  End If;

  If Nvl(n_门诊号, 0) <> 0 Then
    Update 门诊费用记录 Set 标识号 = n_门诊号 Where 病人id = n_病人id;
    Update 病人挂号记录 Set 门诊号 = n_门诊号 Where 病人id = n_病人id;
  End If;
  If Nvl(n_就诊id, 0) <> 0 Then
    p_费用(n_病人id, n_就诊id, v_姓名, v_性别, v_年龄, n_场合, v_说明_In, v_说明);
    If v_说明 Is Not Null Then
      说明_Out := 说明_Out || Chr(13) || '费用部分:' || Chr(13) || v_说明;
    End If;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功","adjust_explain":"' || zlJsonStr(说明_Out, 0) || '"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Updatepatibaseinfo;
/
Create Or Replace Procedure Zl_Exsesvr_Getorderfeeexeinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ------------------------------------------------------------------------------------------------------
  --功能：门诊或住院医嘱[取消]执行完成检查同时获取费用等相关信息
  --入参：Json_In:格式
  --input
  --    is_finish           N 1 执行完成或取消完成，1-执行完成，2-取消执行完成
  --    fee_origin          N 1 费用来源(默认=2：1-门诊费用，2-住院费用)
  --    fee_nos             C 1 费用单据号拼串，格式：NO:记录性质... 如：T000001:1,T000002:2,T000003:3
  --    exe_deptid          N 1 费用执行科室ID,0-表示不区分科室,否则传执行科室id
  --    order_ids           C 1 医嘱IDs，费用单据下对应的医嘱id
  --    send_no             N 1 发送号
  --    wardarea_id         N 1 病人病区id，住院场合需要传入
  --    order_status        N 1 医嘱执行状态，取消执行完成时有此结点
  --出参: Json_Out,格式如下
  --  output
  --    code                  N 1 应答吗：0-失败；1-成功
  --    message               C 1 应答消息：失败时返回具体的错误信息
  --    fee_ids               C 1 需要执行完成的，费用ids，普通非药品卫材费用明细id，仅取行完成时有此结点
  --    stuffdtl_ids          C 1 卫材明细id,逗号分割，用于自动发[退]卫材
  --    rcpdtl_ids            C 1 药品处方明细id，住院可能会有，用于自动发[退]药品
  --    finish_list[]需要执行完成自动审核的费用明细，记帐划价，仅取行完成时有此结点列表
  --         pati_id          N 1 病人id
  --         fee_id           N 1 费用id
  --         fee_no           C 1 费用单据号
  --         serial_num       N 1 费用序号
  --         exe_deptid       N 1 执行科室id
  --         fee_type         N 1 收费类型，0-普通费用，1-药品费，2-跟踪在用卫材费
  --    order_list[]用于医嘱发送打标记，仅取行完成时有此结点列表
  --         order_id         N 1 医嘱ID
  --         send_no          N 1 发送号
  --         type             N 1 费用审核时的类型，用于打标，0-药品费，1-跟踪在用卫材费
  --    cancel_list[]取消执行完成时返回的费用状态更新列表，仅取消执行完成时有此结点列表
  --        fee_id            N 1 费用id
  --        exe_status        N 1 执行状态
  --        exe_people        C 1 执行人
  --        exe_time          C  执行时间
  --------------------------------------------------------------------------------------------------------
  n_审核费用      Number(1);
  v_医嘱ids       Varchar2(32767);
  v_Nos           Varchar2(32767);
  n_执行部门id    Number(18);
  v_执行前先结算  Varchar2(4000);
  v_Fee_Item      Varchar2(32767);
  v_Fee_Item_List Varchar2(32767);
  v_Affirm        Varchar2(3000);

  n_Finish    Number; --1-执行完成，2-取消执行完成
  n_Origin    Number; --1-门诊费用，2-住院费用
  j_Json      Pljson;
  j_Tmp       Pljson;
  j_Output    Pljson;
  v_Jtmp      Varchar2(32767);
  v_Jtmp1     Varchar2(32767);
  v_Jtmp2     Varchar2(32767);
  v_Pati_List Varchar2(32767);
  v_Json_In   Varchar2(32767);
  v_Json_Out  Varchar2(32767);
  v_No        住院费用记录.No%Type;
  v_序号      Varchar2(32767);
  v_Error     Varchar2(32767);
  Err_Custom Exception;

  Procedure p_Checkinfinish As
    ------------------------------------------------------------------------------------------------------
    --功能：住院医嘱执行完成检查
    --入参：Json_In:格式
    --input
    --    fee_nos             C 1 费用单据号拼串，格式：NO:记录性质... 如：T000001:1,T000002:2,T000003:3
    --    exe_deptid          N 1 执行科室ID,0-表示不区分科室,否则传执行科室id
    --    order_ids           C 1 医嘱IDs
    --    send_no             N 1 发送号
    --    wardarea_id         N 1 病区id
    --出参: Json_Out,格式如下
    --  output
    --    code                  N 1 应答吗：0-失败；1-成功
    --    message               C 1 应答消息：失败时返回具体的错误信息
    --    fee_ids               C 1 需要执行完成的，费用ids
    --    stuffdtl_ids          C 1 卫材明细id,逗号分割
    --    rcpdtl_ids            C 1 药品处方明细id
    --    item_list
    --         pati_id          N 1 病人id
    --         pati_pageid      N 1 主页id
    --         fee_id           N 1 费用id
    --         fee_no           C 1 费用单据号
    --         serial_num       N 1 费用序号
    --         exe_deptid       N 1 执行科室id
    --         fee_type         N 1 收费类型，0-普通费用，1-药品费，2-跟踪在用卫材费
    --    order_list
    --         order_id         N 1 医嘱ID
    --         send_no          N 1 发送号
    --         type             N 1 费用审核时的类型，用于打标，0-药品费，1-跟踪在用卫材费
    --------------------------------------------------------------------------------------------------------
    v_卫材明细ids  Varchar2(32767);
    v_药品明细ids  Varchar2(32767);
    v_Nos          Varchar2(32767);
    v_Orders       Varchar2(32767);
    n_执行部门id   Number;
    v_住院自动发料 Varchar2(300);
    --普通费用
    Cursor c_Finish Is
      Select a.Id
      From (Select Distinct a.Id, a.收费类别, a.收费细目id
             From 住院费用记录 A,
                  (Select /*+cardinality(f,10)*/
                     C1 As NO, To_Number(C2) As 记录性质
                    From Table(Cast(f_Str2list2(v_Nos) As t_Strlist2)) F) N
             Where Instr(',' || v_Orders || ',', ',' || a.医嘱序号 || ',') > 0 And a.No = n.No And a.记录性质 = n.记录性质 And
                   a.记录状态 In (0, 1, 3) And a.执行状态 <> 1 And (n_执行部门id = 0 Or a.执行部门id = n_执行部门id)) A, 材料特性 B
      Where a.收费细目id = b.材料id(+) And Not a.收费类别 In ('5', '6', '7') And Not (a.收费类别 = '4' And Nvl(b.跟踪在用, 0) = 1);
  
    --执行中包含跟踪在用的未发卫料时，根据参数设置是否自动发料
    --卫生材料医嘱目前不存在单独和组合执行的情况
    Cursor c_Stuff Is
      Select a.Id, a.医嘱序号 As 医嘱id
      From 住院费用记录 A, 材料特性 D,
           (Select /*+cardinality(f,10)*/
              C1 As NO, To_Number(C2) As 记录性质
             From Table(Cast(f_Str2list2(v_Nos) As t_Strlist2)) F) N
      Where d.材料id = a.收费细目id And a.收费类别 = '4' And d.跟踪在用 = 1 And a.记录状态 = 1 And
            Instr(',' || v_Orders || ',', ',' || a.医嘱序号 || ',') > 0 And a.No = n.No And Mod(a.记录性质, 10) = n.记录性质 And
            (n_执行部门id = 0 Or a.执行部门id = n_执行部门id Or v_住院自动发料 = '1');
  
    --未审核的费用行(包含药品和卫材)
    Cursor c_Verify Is
      Select a.Id, a.No, a.序号, a.医嘱序号 As 医嘱id, a.执行部门id, a.收费类别, a.收费细目id, b.跟踪在用, a.病人id, a.主页id
      From 住院费用记录 A, 材料特性 B,
           (Select /*+cardinality(f,10)*/
              C1 As NO, To_Number(C2) As 记录性质
             From Table(Cast(f_Str2list2(v_Nos) As t_Strlist2)) F) N
      Where a.记帐费用 = 1 And a.记录状态 = 0 And a.价格父号 Is Null And Instr(',' || v_Orders || ',', ',' || a.医嘱序号 || ',') > 0 And
            a.No = n.No And a.记录性质 = n.记录性质 And (n_执行部门id = 0 Or a.执行部门id = n_执行部门id) And a.收费细目id = b.材料id(+)
      Order By a.No, a.序号;
    v_Feeids   Varchar2(32767);
    n_发送号   Number;
    n_Cnt      Number;
    n_状态     Number;
    n_审核标志 Number;
  Begin
    v_Nos        := j_Json.Get_String('fee_nos');
    v_Orders     := j_Json.Get_String('order_ids');
    n_发送号     := j_Json.Get_Number('send_no');
    n_执行部门id := j_Json.Get_Number('exe_deptid');
    n_执行部门id := Nvl(n_执行部门id, 0);
    n_审核标志   := j_Json.Get_Number('fee_audit_status');
    n_状态       := j_Json.Get_Number('si_inp_status');
  
    Select zl_GetSysParameter(63) Into v_住院自动发料 From Dual;
    For R In c_Finish Loop
      v_Feeids := v_Feeids || ',' || r.Id;
    End Loop;
  
    v_Jtmp  := Null;
    v_Jtmp1 := Null;
    v_Jtmp2 := Null;
    --执行时自动审核对应的记帐划价单费用
    --包含医嘱对应的药品及卫材费用，因为医嘱已执行，费用应该生效
    For r_Verify In c_Verify Loop
      n_Cnt := 0;
      If r_Verify.收费类别 = '4' And r_Verify.跟踪在用 = 1 Then
        n_Cnt := 2;
      Elsif r_Verify.收费类别 In ('5', '6', '7') Then
        n_Cnt := 1;
      End If;
      If n_Cnt <> 0 Then
        v_Jtmp := v_Jtmp || ',{"order_id":' || r_Verify.医嘱id;
        v_Jtmp := v_Jtmp || ',"send_no":' || n_发送号;
        v_Jtmp := v_Jtmp || ',"type":' || (n_Cnt - 1);
        v_Jtmp := v_Jtmp || '}';
      End If;
    
      v_Jtmp2 := v_Jtmp2 || ',{"pati_id":' || r_Verify.病人id;
      v_Jtmp2 := v_Jtmp2 || ',"pati_pageid":' || r_Verify.主页id;
      v_Jtmp2 := v_Jtmp2 || ',"fee_id":' || r_Verify.Id;
      v_Jtmp2 := v_Jtmp2 || ',"fee_no":"' || r_Verify.No || '"';
      v_Jtmp2 := v_Jtmp2 || ',"serial_num":' || r_Verify.序号;
      v_Jtmp2 := v_Jtmp2 || ',"exe_deptid":' || r_Verify.执行部门id;
      v_Jtmp2 := v_Jtmp2 || ',"fee_type":' || n_Cnt; --N  1-表示药品，2-表示卫材
      v_Jtmp2 := v_Jtmp2 || '}';
    
      If r_Verify.收费类别 = '4' And r_Verify.跟踪在用 = 1 And (n_执行部门id = 0 Or r_Verify.执行部门id = n_执行部门id Or v_住院自动发料 = '1') Then
        v_卫材明细ids := v_卫材明细ids || ',' || r_Verify.Id;
        v_Feeids      := v_Feeids || ',' || r_Verify.Id;
      End If;
      If r_Verify.收费类别 In ('5', '6', '7') And r_Verify.执行部门id = n_执行部门id Then
        v_药品明细ids := v_药品明细ids || ',' || r_Verify.Id;
        v_Feeids      := v_Feeids || ',' || r_Verify.Id;
      End If;
    
      If v_Pati_List Is Null Then
        v_Pati_List := '{"pati_id":' || r_Verify.病人id;
        v_Pati_List := v_Pati_List || ',"fee_audit_status":' || Nvl(n_审核标志, 0);
        v_Pati_List := v_Pati_List || ',"si_inp_status":' || Nvl(n_状态, 0);
        v_Pati_List := v_Pati_List || '}';
        v_Pati_List := ',"pati_list":[' || v_Pati_List || ']';
      End If;
    
      If r_Verify.No <> v_No And v_序号 Is Not Null Then
        v_Json_In := '{"fee_nos":"' || v_No || '"';
        v_Json_In := v_Json_In || ',"":"' || v_序号 || '"';
        v_Json_In := v_Json_In || v_Pati_List;
        v_Json_In := v_Json_In || '}';
        v_Json_In := '{"input":' || v_Json_In || '}';
        Zl_住院记帐记录_Verify_Check(v_Json_In, v_Json_Out);
        v_序号   := Null;
        j_Tmp    := Pljson();
        j_Output := Pljson();
        j_Tmp    := Pljson(v_Json_Out);
        j_Output := j_Tmp.Get_Pljson('output');
        If j_Output.Get_Number('code') = 0 Then
          v_Error := j_Output.Get_String('message');
          Raise Err_Custom;
        End If;
      End If;
      v_No   := r_Verify.No;
      v_序号 := v_序号 || ',' || r_Verify.序号;
    End Loop;
  
    If v_序号 Is Not Null Then
      v_Json_In := '{"fee_nos":"' || v_No || '"';
      v_Json_In := v_Json_In || ',"":"' || v_序号 || '"';
      v_Json_In := v_Json_In || v_Pati_List;
      v_Json_In := v_Json_In || '}';
      v_Json_In := '{"input":' || v_Json_In || '}';
      Zl_住院记帐记录_Verify_Check(v_Json_In, v_Json_Out);
      j_Tmp    := Pljson();
      j_Output := Pljson();
      j_Tmp    := Pljson(v_Json_Out);
      j_Output := j_Tmp.Get_Pljson('output');
      If j_Output.Get_Number('code') = 0 Then
        v_Error := j_Output.Get_String('message');
        Raise Err_Custom;
      End If;
    End If;
  
    --处理跟踪在用卫材自动发料
    For r_Stuff In c_Stuff Loop
      --需要发卫材的明细
      v_卫材明细ids := v_卫材明细ids || ',' || r_Stuff.Id;
    End Loop;
  
    --处理跟踪在用卫材自动发料
    --根据传入的病区id来确定是否需要自动发药
    For r_Drug In (Select a.Id
                   From 住院费用记录 A,
                        (Select /*+cardinality(f,10)*/
                           C1 As NO, To_Number(C2) As 记录性质
                          From Table(Cast(f_Str2list2(v_Nos) As t_Strlist2)) F) N
                   Where a.收费类别 In ('5', '6', '7') And a.记录状态 = 1 And
                         Instr(',' || v_Orders || ',', ',' || a.医嘱序号 || ',') > 0 And a.No = n.No And a.记录性质 = n.记录性质 And
                         n_执行部门id = a.执行部门id) Loop
      --需要发药品的明细
      v_药品明细ids := v_药品明细ids || ',' || r_Drug.Id;
    End Loop;
  
    v_Jtmp1 := v_Jtmp1 || '{"code":1';
    v_Jtmp1 := v_Jtmp1 || ',"message":"成功"';
    v_Jtmp1 := v_Jtmp1 || ',"stuffdtl_ids":"' || Substr(v_卫材明细ids, 2) || '"';
    v_Jtmp1 := v_Jtmp1 || ',"rcpdtl_ids":"' || Substr(v_药品明细ids, 2) || '"';
    v_Jtmp1 := v_Jtmp1 || ',"fee_ids":"' || Substr(v_Feeids, 2) || '"';
    v_Jtmp1 := v_Jtmp1 || ',"finish_list":[' || Substr(v_Jtmp2, 2) || ']';
    v_Jtmp1 := v_Jtmp1 || ',"order_list":[' || Substr(v_Jtmp, 2) || ']';
    v_Jtmp1 := v_Jtmp1 || '}';
  
    Json_Out := '{"output":' || v_Jtmp1 || '}';
  
  End p_Checkinfinish;

  Procedure p_Checkoutfinish As
    ---------------------------------------------------------------------------
    --功能：门诊医嘱执行完成检查
    --入参：Json_In:格式
    --input
    --    fee_nos             C 1 费用单据号拼串，格式：NO:记录性质... 如：T000001:1,T000002:2,T000003:3
    --    exe_deptid          N 1 执行科室ID,0-表示不区分科室,否则传执行科室id
    --    order_ids           C 1 医嘱IDs
    --    send_no             N 1 发送号
    --出参: Json_Out,格式如下
    --  output
    --    code                  N 1 应答吗：0-失败；1-成功
    --    message               C 1 应答消息：失败时返回具体的错误信息
    --    fee_ids               C 1 需要执行完成的，费用ids
    --    stuffdtl_ids          C 1 卫材明细id,逗号分割
    --    item_list
    --         pati_id          N 1 病人id
    --         fee_id           N 1 费用id
    --         fee_no           C 1 费用单据号
    --         serial_num       N 1 费用序号
    --         exe_deptid       N 1 执行科室id
    --         fee_type         N 1 收费类型，0-普通费用，1-药品费，2-跟踪在用卫材费
    --    order_list
    --         order_id         N 1 医嘱ID
    --         send_no          N 1 发送号
    --         type             N 1 费用审核时的类型，用于打标，0-药品费，1-跟踪在用卫材费
    --------------------------------------------------------------------------------
    n_发送号       Number;
    v_Orders       Varchar2(32767);
    n_执行部门id   门诊费用记录.执行部门id%Type;
    n_Cnt          Number;
    v_Error        Varchar2(2000);
    v_门诊自动发料 Varchar2(300);
    Err_Custom Exception;
    v_执行前先结算 Varchar2(500);
    v_卫材明细ids  Varchar2(32767);
    v_Nos          Varchar2(32767);
    v_Feeids       Varchar2(32767);
    Cursor c_Finishone Is
      Select a.Id, a.医嘱序号 As 医嘱id
      From (Select a.Id, a.收费类别, a.收费细目id, a.医嘱序号
             From 门诊费用记录 A,
                  (Select /*+cardinality(f,10)*/
                     C1 As NO, To_Number(C2) As 记录性质
                    From Table(Cast(f_Str2list2(v_Nos) As t_Strlist2)) F) N
             Where Instr(',' || v_Orders || ',', ',' || a.医嘱序号 || ',') > 0 And a.No = n.No And Mod(a.记录性质, 10) = n.记录性质 And
                   a.记录状态 In (0, 1, 3) And a.执行状态 <> 1 And (n_执行部门id = 0 Or a.执行部门id = n_执行部门id)) A, 材料特性 B
      Where a.收费细目id = b.材料id(+) And Not a.收费类别 In ('5', '6', '7') And Not (a.收费类别 = '4' And Nvl(b.跟踪在用, 0) = 1);
  
    --执行中包含跟踪在用的未发卫料时，根据参数设置是否自动发料
    --卫生材料医嘱目前不存在单独和组合执行的情况
    Cursor c_Stuff Is
      Select a.Id, a.医嘱序号 As 医嘱id
      From 门诊费用记录 A, 材料特性 D,
           (Select /*+cardinality(f,10)*/
              C1 As NO, To_Number(C2) As 记录性质
             From Table(Cast(f_Str2list2(v_Nos) As t_Strlist2)) F) N
      Where d.材料id = a.收费细目id And a.收费类别 = '4' And d.跟踪在用 = 1 And a.记录状态 = 1 And
            Instr(',' || v_Orders || ',', ',' || a.医嘱序号 || ',') > 0 And a.No = n.No And Mod(a.记录性质, 10) = n.记录性质 And
            (n_执行部门id = 0 Or a.执行部门id = n_执行部门id Or v_门诊自动发料 = '1');
  
    --未审核的费用行(包含药品和卫材)
    Cursor c_Verifyone(P记帐费用 Number) Is
      Select a.Id, a.No, a.医嘱序号 As 医嘱id, a.序号, a.执行部门id, a.收费类别, a.收费细目id, b.跟踪在用, a.病人id
      From 门诊费用记录 A, 材料特性 B,
           (Select /*+cardinality(f,10)*/
              C1 As NO, To_Number(C2) As 记录性质
             From Table(Cast(f_Str2list2(v_Nos) As t_Strlist2)) F) N
      Where Nvl(a.记帐费用, 0) = P记帐费用 And a.记录状态 = 0 And a.价格父号 Is Null And
            Instr(',' || v_Orders || ',', ',' || a.医嘱序号 || ',') > 0 And a.No = n.No And Mod(a.记录性质, 10) = n.记录性质 And
            (n_执行部门id = 0 Or a.执行部门id = n_执行部门id) And a.收费细目id = b.材料id(+)
      Order By NO, 序号;
  Begin
    v_Nos        := j_Json.Get_String('fee_nos');
    v_Orders     := j_Json.Get_String('order_ids');
    n_发送号     := j_Json.Get_Number('send_no');
    n_执行部门id := j_Json.Get_Number('exe_deptid');
    n_执行部门id := Nvl(n_执行部门id, 0);
    Select zl_GetSysParameter(92) Into v_门诊自动发料 From Dual;
    For R In c_Finishone Loop
      v_Feeids := v_Feeids || ',' || r.Id;
    End Loop;
    v_Feeids := Substr(v_Feeids, 2);
    Select Count(1)
    Into n_Cnt
    From 门诊费用记录 A
    Where a.Id In (Select /*+cardinality(x,10)*/
                    x.Column_Value
                   From Table(Cast(f_Num2list(v_Feeids) As Zltools.t_Numlist)) X) And a.费用状态 = 1 And
          Nvl(a.结帐id, 0) <> 0;
    If n_Cnt > 0 Then
      v_Error := '当前执行的医嘱对应的费用单据中存在异常单据。';
      Raise Err_Custom;
    End If;
    Select zl_GetSysParameter(163) Into v_执行前先结算 From Dual;
    --执行时自动审核对应的记帐划价单费用
    --包含医嘱对应的药品及卫材费用，因为医嘱已执行，费用应该生效
    If Nvl(v_执行前先结算, '0') <> '0' Then
      For r_Verify In c_Verifyone(0) Loop
        v_Error := '当前执行的医嘱还存在未收取的费用。';
        Raise Err_Custom;
      End Loop;
    End If;
    v_Jtmp  := Null;
    v_Jtmp1 := Null;
    v_Jtmp2 := Null;
    For r_Verify In c_Verifyone(1) Loop
      n_Cnt := 0;
      If r_Verify.收费类别 = '4' And r_Verify.跟踪在用 = 1 Then
        n_Cnt := 2;
      Elsif r_Verify.收费类别 In ('5', '6', '7') Then
        n_Cnt := 1;
      End If;
    
      If n_Cnt <> 0 Then
        v_Jtmp := v_Jtmp || ',{"order_id":' || r_Verify.医嘱id;
        v_Jtmp := v_Jtmp || ',"send_no":' || n_发送号;
        v_Jtmp := v_Jtmp || ',"type":' || (n_Cnt - 1);
        v_Jtmp := v_Jtmp || '}';
      End If;
    
      v_Jtmp2 := v_Jtmp2 || ',{"pati_id":' || r_Verify.病人id;
      v_Jtmp2 := v_Jtmp2 || ',"fee_id":' || r_Verify.Id;
      v_Jtmp2 := v_Jtmp2 || ',"fee_no":"' || r_Verify.No || '"';
      v_Jtmp2 := v_Jtmp2 || ',"serial_num":' || r_Verify.序号;
      v_Jtmp2 := v_Jtmp2 || ',"exe_deptid":' || r_Verify.执行部门id;
      v_Jtmp2 := v_Jtmp2 || ',"fee_type":' || n_Cnt; --N  1-表示药品，2-表示卫材
      v_Jtmp2 := v_Jtmp2 || '}';
    
      If r_Verify.收费类别 = '4' And r_Verify.跟踪在用 = 1 And (n_执行部门id = 0 Or r_Verify.执行部门id = n_执行部门id Or v_门诊自动发料 = '1') Then
        v_卫材明细ids := v_卫材明细ids || ',' || r_Verify.Id;
      End If;
    End Loop;
    --处理跟踪在用卫材自动发料
    For r_Stuff In c_Stuff Loop
      --需要发卫材的明细
      v_卫材明细ids := v_卫材明细ids || ',' || r_Stuff.Id;
    End Loop;
  
    v_Jtmp1 := v_Jtmp1 || '{"code":1';
    v_Jtmp1 := v_Jtmp1 || ',"message":"成功"';
    v_Jtmp1 := v_Jtmp1 || ',"stuffdtl_ids":"' || Substr(v_卫材明细ids, 2) || '"';
    v_Jtmp1 := v_Jtmp1 || ',"fee_ids":"' || v_Feeids || '"';
    v_Jtmp1 := v_Jtmp1 || ',"finish_list":[' || Substr(v_Jtmp2, 2) || ']';
    v_Jtmp1 := v_Jtmp1 || ',"order_list":[' || Substr(v_Jtmp, 2) || ']';
    v_Jtmp1 := v_Jtmp1 || '}';
  
    Json_Out := '{"output":' || v_Jtmp1 || '}';
  
  Exception
    When Err_Custom Then
      Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(v_Error) || '"}}';
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Checkoutfinish;

  Procedure p_Checkincancel As
    ---------------------------------------------------------------------------
    --功能：住院医嘱取消执行完成检查
    --入参：Json_In:格式
    --input
    --    fee_nos             C 1 费用单据号拼串，格式：NO:记录性质... 如：T000001:1,T000002:2,T000003:3
    --    exe_deptid          N 1 执行科室ID,0-表示不区分科室,仅处理指定执行部门的费用，不传或传入0时不限制执行部门
    --    order_ids           C 1 医嘱IDs
    --    order_status        N 1 医嘱执行状态
    --    wardarea_id         N 1 病区id
    --出参: Json_Out,格式如下
    --  output
    --    code                  N 1 应答吗：0-失败；1-成功
    --    message               C 1 应答消息：失败时返回具体的错误信息
    --    stuffdtl_ids          C 1 卫材明细id,逗号分割
    --    rcpdtl_ids            C 1 处方明细id,逗号分割
    --    item_list       要更新的费用明细
    --        fee_id            N 1 费用id
    --        exe_status        N 1 执行状态
    --        exe_people        C 1 执行人
    --        exe_time          C  执行时间
    ---------------------------------------------------------------------------
  
    v_Orders      Varchar2(32767);
    n_执行部门id  门诊费用记录.执行部门id%Type;
    v_卫材明细ids Varchar2(32767);
    v_Nos         Varchar2(32767);
    v_药品明细ids Varchar2(32767);
    n_费用id      Number;
    d_执行时间    Date;
    v_执行人      Varchar2(100);
    n_执行状态    Number;
    --要取消执行的费用行(不包含药品和跟踪在用的卫材)
    Cursor c_Finishone Is
      Select a.Id, a.执行时间, a.执行人
      From (Select a.Id, a.收费类别, a.收费细目id, a.执行时间, a.执行人
             From 住院费用记录 A,
                  (Select /*+cardinality(f,10)*/
                     C1 As NO, To_Number(C2) As 记录性质
                    From Table(Cast(f_Str2list2(v_Nos) As t_Strlist2)) F) N
             Where Instr(',' || v_Orders || ',', ',' || a.医嘱序号 || ',') > 0 And a.No = n.No And
                   (a.记录性质 = n.记录性质 Or a.记录性质 = 11 And n.记录性质 = 1) And a.记录状态 In (0, 1, 3) And
                   (n_执行部门id = 0 Or a.执行部门id = n_执行部门id)) A, 材料特性 B
      Where a.收费细目id = b.材料id(+) And Not a.收费类别 In ('5', '6', '7') And Not (a.收费类别 = '4' And Nvl(b.跟踪在用, 0) = 1);
  
    --取消执行中包含跟踪在用的发料卫料时，根据参数设置是否自动退料
    --卫生材料医嘱目前不存在单独和组合执行的情况
    Cursor c_Stuff Is
      Select a.Id
      From 住院费用记录 A, 材料特性 D,
           (Select /*+cardinality(f,10)*/
              C1 As NO, To_Number(C2) As 记录性质
             From Table(Cast(f_Str2list2(v_Nos) As t_Strlist2)) F) N
      Where a.收费类别 = '4' And a.记录状态 = 1 And a.收费细目id = d.材料id And d.跟踪在用 = 1 And
            Instr(',' || v_Orders || ',', ',' || a.医嘱序号 || ',') > 0 And a.No = n.No And
            (a.记录性质 = n.记录性质 Or a.记录性质 = 11 And n.记录性质 = 1) And (n_执行部门id = 0 Or a.执行部门id = n_执行部门id);
  Begin
    v_Nos        := j_Json.Get_String('fee_nos');
    v_Orders     := j_Json.Get_String('order_ids');
    n_执行部门id := j_Json.Get_Number('exe_deptid');
    n_执行部门id := Nvl(n_执行部门id, 0);
    n_执行状态   := j_Json.Get_Number('order_status');
    v_Jtmp       := Null;
    v_Jtmp1      := Null;
    For R In c_Finishone Loop
      Select r.Id As 费用id, Decode(n_执行状态, 0, d_执行时间, r.执行时间) As 执行时间, Decode(n_执行状态, 0, Null, r.执行人) As 执行人
      Into n_费用id, d_执行时间, v_执行人
      From Dual;
    
      v_Jtmp := v_Jtmp || ',{"fee_id":' || n_费用id;
      v_Jtmp := v_Jtmp || ',"exe_status":' || Nvl(n_执行状态, 0);
      v_Jtmp := v_Jtmp || ',"exe_people":"' || Zljsonstr(v_执行人) || '"';
      v_Jtmp := v_Jtmp || ',"exe_time":"' || To_Char(d_执行时间, 'yyyy-mm-dd hh24:mi:ss') || '"';
      v_Jtmp := v_Jtmp || ',"fee_type":0';
      v_Jtmp := v_Jtmp || '}';
    
    End Loop;
  
    --处理跟踪在用卫材自动退料
    For r_Stuff In c_Stuff Loop
      --需要退的卫材明细
      v_卫材明细ids := v_卫材明细ids || ',' || r_Stuff.Id;
    
      v_Jtmp := v_Jtmp || ',{"fee_id":' || r_Stuff.Id;
      v_Jtmp := v_Jtmp || ',"exe_status":0';
      v_Jtmp := v_Jtmp || ',"exe_people":""';
      v_Jtmp := v_Jtmp || ',"exe_time":""';
      v_Jtmp := v_Jtmp || ',"fee_type":1';
      v_Jtmp := v_Jtmp || '}';
    
    End Loop;
  
    --根据传入的病区id来确定是否需要自退药
    For r_Drug In (Select a.Id
                   From 住院费用记录 A,
                        (Select /*+cardinality(f,10)*/
                           C1 As NO, To_Number(C2) As 记录性质
                          From Table(Cast(f_Str2list2(v_Nos) As t_Strlist2)) F) N
                   Where a.收费类别 In ('5', '6', '7') And a.记录状态 = 1 And
                         Instr(',' || v_Orders || ',', ',' || a.医嘱序号 || ',') > 0 And a.No = n.No And a.记录性质 = n.记录性质 And
                         n_执行部门id = a.执行部门id) Loop
      --需要退药品的明细
      v_药品明细ids := v_药品明细ids || ',' || r_Drug.Id;
    
      v_Jtmp := v_Jtmp || ',{"fee_id":' || r_Drug.Id;
      v_Jtmp := v_Jtmp || ',"exe_status":0';
      v_Jtmp := v_Jtmp || ',"exe_people":""';
      v_Jtmp := v_Jtmp || ',"exe_time":""';
      v_Jtmp := v_Jtmp || ',"fee_type":2';
      v_Jtmp := v_Jtmp || '}';
    
    End Loop;
  
    v_Jtmp1 := v_Jtmp1 || '{"code":1';
    v_Jtmp1 := v_Jtmp1 || ',"message":"成功"';
    v_Jtmp1 := v_Jtmp1 || ',"stuffdtl_ids":"' || Substr(v_卫材明细ids, 2) || '"';
    v_Jtmp1 := v_Jtmp1 || ',"rcpdtl_ids":"' || Substr(v_药品明细ids, 2) || '"';
    v_Jtmp1 := v_Jtmp1 || ',"cancel_list":[' || Substr(v_Jtmp, 2) || ']';
    v_Jtmp1 := v_Jtmp1 || '}';
  
    Json_Out := '{"output":' || v_Jtmp1 || '}';
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Checkincancel;

  Procedure p_Checkoutcancel As
    ---------------------------------------------------------------------------
    --功能：门诊医嘱取消执行完成检查
    --入参：Json_In:格式
    --input
    --    fee_nos             C 1 费用单据号拼串，格式：NO:记录性质... 如：T000001:1,T000002:2,T000003:3
    --    exe_deptid          N 1 执行科室ID,0-表示不区分科室,仅处理指定执行部门的费用，不传或传入0时不限制执行部门
    --    order_ids           C 1 医嘱IDs
    --    order_status        N 1 医嘱执行状态
    --出参: Json_Out,格式如下
    --  output
    --    code                  N 1 应答吗：0-失败；1-成功
    --    message               C 1 应答消息：失败时返回具体的错误信息
    --    stuffdtl_ids          C 1 卫材明细id,逗号分割
    --    item_list       要更新的费用明细
    --        fee_id            N 1 费用id
    --        exe_status        N 1 执行状态
    --        exe_people        C 1 执行人
    --        exe_time          C  执行时间
    ---------------------------------------------------------------------------
  
    v_Orders      Varchar2(32767);
    n_执行部门id  门诊费用记录.执行部门id%Type;
    v_卫材明细ids Varchar2(32767);
    v_Nos         Varchar2(32767);
    n_费用id      Number;
    d_执行时间    Date;
    v_执行人      Varchar2(100);
    n_执行状态    Number;
  
    --要取消执行的费用行(不包含药品和跟踪在用的卫材)
    Cursor c_Finishone Is
      Select a.Id, a.执行时间, a.执行人
      From (Select a.Id, a.收费类别, a.收费细目id, a.执行时间, a.执行人
             From 门诊费用记录 A,
                  (Select /*+cardinality(f,10)*/
                     C1 As NO, To_Number(C2) As 记录性质
                    From Table(Cast(f_Str2list2(v_Nos) As t_Strlist2)) F) N
             Where Instr(',' || v_Orders || ',', ',' || a.医嘱序号 || ',') > 0 And a.No = n.No And
                   (a.记录性质 = n.记录性质 Or a.记录性质 = 11 And n.记录性质 = 1) And a.记录状态 In (0, 1, 3) And
                   (n_执行部门id = 0 Or a.执行部门id = n_执行部门id)) A, 材料特性 B
      Where a.收费细目id = b.材料id(+) And Not a.收费类别 In ('5', '6', '7') And Not (a.收费类别 = '4' And Nvl(b.跟踪在用, 0) = 1);
  
    --取消执行中包含跟踪在用的发料卫料时，根据参数设置是否自动退料
    --卫生材料医嘱目前不存在单独和组合执行的情况
    Cursor c_Stuff Is
      Select a.Id
      From 门诊费用记录 A, 材料特性 D,
           (Select /*+cardinality(f,10)*/
              C1 As NO, To_Number(C2) As 记录性质
             From Table(Cast(f_Str2list2(v_Nos) As t_Strlist2)) F) N
      Where a.收费类别 = '4' And a.记录状态 = 1 And a.收费细目id = d.材料id And d.跟踪在用 = 1 And
            Instr(',' || v_Orders || ',', ',' || a.医嘱序号 || ',') > 0 And a.No = n.No And
            (a.记录性质 = n.记录性质 Or a.记录性质 = 11 And n.记录性质 = 1) And (n_执行部门id = 0 Or a.执行部门id = n_执行部门id);
  Begin
    v_Nos        := j_Json.Get_String('fee_nos');
    v_Orders     := j_Json.Get_String('order_ids');
    n_执行部门id := j_Json.Get_Number('exe_deptid');
    n_执行部门id := Nvl(n_执行部门id, 0);
    n_执行状态   := j_Json.Get_Number('order_status');
    v_Jtmp       := Null;
    v_Jtmp1      := Null;
    For R In c_Finishone Loop
      Select r.Id As 费用id, Decode(n_执行状态, 0, d_执行时间, r.执行时间) As 执行时间, Decode(n_执行状态, 0, Null, r.执行人) As 执行人
      Into n_费用id, d_执行时间, v_执行人
      From Dual;
      v_Jtmp := v_Jtmp || ',{"fee_id":' || n_费用id;
      v_Jtmp := v_Jtmp || ',"exe_status":' || Nvl(n_执行状态, 0);
      v_Jtmp := v_Jtmp || ',"exe_people":"' || Zljsonstr(v_执行人) || '"';
      v_Jtmp := v_Jtmp || ',"exe_time":"' || To_Char(d_执行时间, 'yyyy-mm-dd hh24:mi:ss') || '"';
      v_Jtmp := v_Jtmp || ',"fee_type":0';
      v_Jtmp := v_Jtmp || '}';
    End Loop;
    --处理跟踪在用卫材自动发料
    For r_Stuff In c_Stuff Loop
      --需要退的卫材明细
      v_卫材明细ids := v_卫材明细ids || ',' || r_Stuff.Id;
    
      v_Jtmp := v_Jtmp || ',{"fee_id":' || r_Stuff.Id;
      v_Jtmp := v_Jtmp || ',"exe_status":' || 0;
      v_Jtmp := v_Jtmp || ',"exe_people":""';
      v_Jtmp := v_Jtmp || ',"exe_time":""';
      v_Jtmp := v_Jtmp || ',"fee_type":1';
      v_Jtmp := v_Jtmp || '}';
    End Loop;
  
    v_Jtmp1 := v_Jtmp1 || '{"code":1';
    v_Jtmp1 := v_Jtmp1 || ',"message":"成功"';
    v_Jtmp1 := v_Jtmp1 || ',"stuffdtl_ids":"' || Substr(v_卫材明细ids, 2) || '"';
    v_Jtmp1 := v_Jtmp1 || ',"cancel_list":[' || Substr(v_Jtmp, 2) || ']';
    v_Jtmp1 := v_Jtmp1 || '}';
  
    Json_Out := '{"output":' || v_Jtmp1 || '}';
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Checkoutcancel;

  Procedure p_完成执行(Json_Out Out Varchar2) As
    n_自动发料 Number(1);
    Cursor c_Fee Is
      Select a.来源, a.Id, a.No, a.序号, a.记录性质, a.记录状态, a.执行状态, a.收费细目id, a.收费类别, a.执行部门id, a.费用状态, a.结帐id, a.记帐费用, a.价格父号,
             a.病人id, a.医嘱id, b.跟踪在用
      From (Select 2 来源, a.Id, a.No, a.序号, a.记录性质, a.记录状态, a.执行状态, a.收费细目id, a.收费类别, a.执行部门id, a.费用状态, a.结帐id, a.记帐费用,
                    a.价格父号, a.病人id, a.医嘱序号 医嘱id
             From 住院费用记录 A
             Where a.No In (Select /*+cardinality(X,10)*/
                             x.Column_Value NO
                            From Table(Cast(f_Str2list(v_Nos) As Zltools.t_Strlist)) X) And
                   Instr(',' || v_医嘱ids || ',', ',' || a.医嘱序号 || ',') > 0 And (n_执行部门id = 0 Or a.执行部门id = n_执行部门id)
             Union All
             Select 1 来源, a.Id, a.No, a.序号, a.记录性质, a.记录状态, a.执行状态, a.收费细目id, a.收费类别, a.执行部门id, a.费用状态, a.结帐id, a.记帐费用,
                    a.价格父号, a.病人id, a.医嘱序号 医嘱id
             From 门诊费用记录 A
             Where a.No In (Select /*+cardinality(X,10)*/
                             x.Column_Value NO
                            From Table(Cast(f_Str2list(v_Nos) As Zltools.t_Strlist)) X) And
                   Instr(',' || v_医嘱ids || ',', ',' || a.医嘱序号 || ',') > 0 And (n_执行部门id = 0 Or a.执行部门id = n_执行部门id)) A,
           材料特性 B
      Where a.收费细目id = b.材料id(+)
      Order By a.No, a.收费细目id;
  Begin
    Select zl_GetSysParameter(163) Into v_执行前先结算 From Dual;
    v_执行前先结算 := Nvl(v_执行前先结算, '0');
  
    For r_费用 In c_Fee Loop
      If r_费用.来源 = 1 Then
        --门诊费用特有检查
        If r_费用.费用状态 = 1 And Nvl(r_费用.结帐id, 0) <> 0 Then
          v_Error := '当前执行的医嘱对应的费用单据中存在异常单据。';
          Raise Err_Custom;
        End If;
        If v_执行前先结算 <> '0' Then
          If Nvl(r_费用.记帐费用, 0) = 0 And r_费用.记录状态 = 0 And r_费用.价格父号 Is Null Then
            v_Error := '当前执行的医嘱还存在未收取的费用。';
            Raise Err_Custom;
          End If;
        End If;
      End If;
      v_Fee_Item := Null;
      n_自动发料 := 0;
      n_审核费用 := 0;
      --待审核明细
      If Nvl(r_费用.记帐费用, 0) = 1 And r_费用.记录状态 = 0 And r_费用.价格父号 Is Null Then
        n_审核费用 := 1;
        v_Fee_Item := v_Fee_Item || '{"fee_origin":' || r_费用.来源;
        v_Fee_Item := v_Fee_Item || ',"fee_id":' || r_费用.Id;
        v_Fee_Item := v_Fee_Item || ',"fee_no":"' || r_费用.No || '"';
        v_Fee_Item := v_Fee_Item || ',"bill_prop":' || r_费用.记录性质;
        v_Fee_Item := v_Fee_Item || ',"rec_state":' || r_费用.记录状态;
        v_Fee_Item := v_Fee_Item || ',"serial_num":' || r_费用.序号;
        v_Fee_Item := v_Fee_Item || ',"fee_type":"' || r_费用.收费类别 || '"';
        v_Fee_Item := v_Fee_Item || ',"fee_item_id":' || r_费用.收费细目id;
        v_Fee_Item := v_Fee_Item || ',"order_id":' || r_费用.医嘱id;
        v_Fee_Item := v_Fee_Item || ',"stuff_used":' || Nvl(r_费用.跟踪在用, 0);
        v_Fee_Item := v_Fee_Item || ',"exe_dept_id":' || Nvl(r_费用.执行部门id, 0); --执行部门id
        v_Fee_Item := v_Fee_Item || ',"is_verify":' || n_审核费用;
      End If;
    
      --执行完成明细
      If r_费用.记录状态 In (0, 1, 3) And r_费用.执行状态 <> 1 Then
        If r_费用.记录状态 In (0, 1) And r_费用.跟踪在用 = 1 Then
          If r_费用.来源 = 1 And r_费用.记录性质 In (1, 11) Or r_费用.来源 = 2 And r_费用.记录性质 = 2 Then
            --门诊收费记录自动发料或住院记帐记录
            n_自动发料 := 1;
          End If;
        End If;
      
        If v_Fee_Item Is Null Then
          v_Fee_Item := v_Fee_Item || '{"fee_origin":' || r_费用.来源;
          v_Fee_Item := v_Fee_Item || ',"fee_id":' || r_费用.Id;
          v_Fee_Item := v_Fee_Item || ',"fee_no":"' || r_费用.No || '"';
          v_Fee_Item := v_Fee_Item || ',"bill_prop":' || r_费用.记录性质;
          v_Fee_Item := v_Fee_Item || ',"rec_state":' || r_费用.记录状态;
          v_Fee_Item := v_Fee_Item || ',"serial_num":' || r_费用.序号;
          v_Fee_Item := v_Fee_Item || ',"fee_type":"' || r_费用.收费类别 || '"';
          v_Fee_Item := v_Fee_Item || ',"fee_item_id":' || r_费用.收费细目id;
          v_Fee_Item := v_Fee_Item || ',"order_id":' || r_费用.医嘱id;
          v_Fee_Item := v_Fee_Item || ',"stuff_used":' || Nvl(r_费用.跟踪在用, 0);
          v_Fee_Item := v_Fee_Item || ',"exe_dept_id":' || Nvl(r_费用.执行部门id, 0); --执行部门id
          v_Fee_Item := v_Fee_Item || ',"is_verify":0';
        End If;
        v_Fee_Item := v_Fee_Item || ',"is_finish":1';
      End If;
    
      If v_Fee_Item Is Not Null Then
        v_Fee_Item := v_Fee_Item || '}';
      End If;
    
      If v_Fee_Item Is Not Null Then
        v_Fee_Item_List := v_Fee_Item_List || ',' || v_Fee_Item;
      End If;
    
    End Loop;
  
    If v_Fee_Item_List Is Not Null Then
      v_Fee_Item_List := ',"fee_list":[' || Substr(v_Fee_Item_List, 2) || ']';
    End If;
    v_Affirm := ',"is_affirm":' || Nvl(n_审核费用, 0);
    Json_Out := '{"output":{"code":1,"message":"成功"' || v_Affirm || v_Fee_Item_List || '}}';
  
  End;
  ---------------------------------------------------------------------------------------------------
Begin
  j_Tmp    := Pljson(Json_In);
  j_Json   := j_Tmp.Get_Pljson('input');
  n_Origin := j_Json.Get_Number('fee_origin');

  n_Finish     := j_Json.Get_Number('is_finish');
  v_医嘱ids    := j_Json.Get_String('fee_order_ids');
  v_Nos        := j_Json.Get_String('fee_nos');
  n_执行部门id := Nvl(j_Json.Get_Number('exe_deptid'), 0);

  If 1 = n_Finish Then
    p_完成执行(Json_Out);
    Return;
  End If;

  If 1 = n_Origin And 1 = n_Finish Then
    --门诊医嘱执行完成
    p_Checkoutfinish;
  Elsif 2 = n_Origin And 1 = n_Finish Then
    --住院医嘱执行完成
    p_Checkinfinish;
  End If;
  If 1 = n_Origin And 2 = n_Finish Then
    --门诊医嘱取消执行完成
    p_Checkoutcancel;
  Elsif 2 = n_Origin And 2 = n_Finish Then
    --住院医嘱取消执行完成
    p_Checkincancel;
  End If;
Exception
  When Err_Custom Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(v_Error) || '"}}';
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Exsesvr_Getorderfeeexeinfo;
/

Create Or Replace Procedure Zl_Exsesvr_Checkorderrevoke
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：门诊医嘱作废相关检查，门诊作废医嘱单组医嘱作废，按一次发送操作作废一批医嘱
  --入参：Json_In:格式
  --  input
  --     fee_nos                        C 1 格式：U0016921,U0016922,,,
  --     order_ids                      C 1 医嘱id串，本次处理的一批医嘱id逗号分割
  --     bill_prop                      N 1 记录性质,1-收费,2-记帐,门诊有医保退费重收,11和12
  --     after_order_ids                C 1 门诊参数先作废后退药品,已发药的医嘱行的医嘱id明细串,逗号分割
  --     exe_fee_ids                    C 1 已执行或正在执行的费用id拼串，主要应用于门诊先作废后退药的情况，费用为未收费，但是药品已发料的情况

  --出参: Json_Out,格式如下
  --  output
  --    code                          N 1 应答吗：0-失败；1-成功
  --    message                       C 1 应答消息：失败时返回具体的错误信息
  --    exist_balance                 N 1 是否存在已结帐的费用,0-不存在,1-存在
  --    exist_verify                  N 1 是否存在已审核的记帐单,0-不存在,1-存在
  --    del_list[]可以直接删除的记帐单
  --       fee_source                 N 1 费用来源:1-门诊费用记录，2-住院费用记录
  --       fee_bill_type              N 1 记录性质:1-收费单，2-记帐单
  --       fee_no                     C 1 费用单据号
  --       exe_sta_nums               C 1 序号串，用修正费用执行状态，格式：序号1,序号2,…
  --       serial_num                 C 1 两种格式，记录性质=1时格式：序号1,序号2,… 记录性质=2时格式：序号1:数量:执行状态1,序号2:数量2:执行状态2,…
  ---------------------------------------------------------------------------
  j_Input         Pljson;
  j_Json          Pljson;
  v_Nos           Varchar2(32767);
  v_先作废医嘱ids Varchar2(32767);
  v_修改执行状态  Varchar2(32767);
  v_Jtmp1         Varchar2(32767);
  v_Serial_Num    Varchar2(32767);
  v_Exe_Fee_Ids   Varchar2(32767);
  v_Order_Ids     Varchar2(32767);
  v_Count         Number;
  n_记录性质      Number(3);
  v_Error         Varchar2(255);
  Err_Custom Exception;
  n_有已结帐费用 Number;
  n_有记帐已审核 Number;
  v_Tmpno        Varchar2(30);
  v_Del_List     Varchar2(32767);

  --需要销帐的费用列表
  Cursor c_Billdel Is
    Select a.记录性质, a.No, f_List2str(Cast(Collect(a.序号 || '') As t_Strlist), ',') 序号串,
           f_List2str(Cast(Collect(a.销记帐单 || '') As t_Strlist), ',') 销记帐单
    From (Select Decode(a.记录性质, 11, 1, a.记录性质) As 记录性质, a.记录状态, a.No, a.序号, a.序号 || ':' || (Nvl(付数, 1) * 数次) || ':0' 销记帐单
           From 门诊费用记录 A
           Where a.医嘱序号 Is Not Null And Instr(',' || v_Order_Ids || ',', ',' || a.医嘱序号 || ',') > 0 And
                 Mod(a.记录性质, 10) = n_记录性质 And a.记录状态 In (0, 1) And Nvl(Nvl(付数, 1) * 数次, 0) <> 0 And
                 (v_Exe_Fee_Ids Is Null Or Instr(',' || v_Exe_Fee_Ids || ',', ',' || a.Id || ',') = 0) And
                 a.No In (Select /*+cardinality(X,10)*/
                           x.Column_Value
                          From Table(f_Str2list(v_Nos)) X)) A
    Group By a.记录性质, a.No;

Begin
  --解析入参
  j_Input         := Pljson(Json_In);
  j_Json          := j_Input.Get_Pljson('input');
  n_记录性质      := j_Json.Get_Number('bill_prop');
  v_Nos           := j_Json.Get_String('fee_nos');
  v_先作废医嘱ids := j_Json.Get_String('after_order_ids');
  v_Exe_Fee_Ids   := j_Json.Get_String('exe_fee_ids');
  v_Order_Ids     := j_Json.Get_String('order_ids');

  --费用转出判断
  Select Count(1)
  Into v_Count
  From H门诊费用记录 A
  Where a.医嘱序号 Is Not Null And (v_先作废医嘱ids Is Null Or Instr(',' || v_先作废医嘱ids || ',', ',' || a.医嘱序号 || ',') = 0) And
        Instr(',' || v_Order_Ids || ',', ',' || a.医嘱序号 || ',') > 0 And Mod(a.记录性质, 10) = n_记录性质 And
        a.No In (Select /*+cardinality(X,10)*/
                  x.Column_Value
                 From Table(f_Str2list(v_Nos)) X);
  If v_Count > 0 Then
    v_Error := '该医嘱的费用已经全部或部份转出到后备数据库，不允许操作。' || Chr(13) || '您可以与系统管理员联系，将相应数据抽选返回。';
    Raise Err_Custom;
  End If;

  --检查作废医嘱对应的费用是否存结帐情况
  Select Count(1)
  Into n_有已结帐费用
  From (Select Count(1) 记录行数
         From 门诊费用记录 A
         Where a.医嘱序号 Is Not Null And (v_先作废医嘱ids Is Null Or Instr(',' || v_先作废医嘱ids || ',', ',' || a.医嘱序号 || ',') = 0) And
               Instr(',' || v_Order_Ids || ',', ',' || a.医嘱序号 || ',') > 0 And Mod(a.记录性质, 10) = n_记录性质 And
               a.记录性质 In (2, 12) And a.记录状态 = 1 And
               Not (Nvl(a.费用状态, 0) = 1 And Nvl(a.结帐id, 0) = 0 And Nvl(a.记录状态, 0) = 1) And
               a.No In (Select /*+cardinality(X,10)*/
                         x.Column_Value
                        From Table(f_Str2list(v_Nos)) X)
         Group By a.No, Nvl(a.价格父号, a.序号)
         Having Sum(Nvl(a.结帐金额, 0)) <> 0);

  --已审核记帐费用检查审是否有已审核的记帐费用
  Select Count(1)
  Into n_有记帐已审核
  From 门诊费用记录 A
  Where a.医嘱序号 Is Not Null And (v_先作废医嘱ids Is Null Or Instr(',' || v_先作废医嘱ids || ',', ',' || a.医嘱序号 || ',') = 0) And
        Instr(',' || v_Order_Ids || ',', ',' || a.医嘱序号 || ',') > 0 And Mod(a.记录性质, 10) = n_记录性质 And a.记录性质 In (2, 12) And
        a.记录状态 = 1 And Not (Nvl(a.费用状态, 0) = 1 And Nvl(a.结帐id, 0) = 0 And Nvl(a.记录状态, 0) = 1) And a.划价人 Is Not Null And
        a.划价人 <> a.操作员姓名 And Not (a.费用状态 = 1 And a.结帐id Is Null And a.记录状态 = 1) And
        a.No In (Select /*+cardinality(X,10)*/
                  x.Column_Value
                 From Table(f_Str2list(v_Nos)) X);

  ----收费异常检查
  Select Max(a.No)
  Into v_Tmpno
  From 门诊费用记录 A
  Where a.医嘱序号 Is Not Null And (v_先作废医嘱ids Is Null Or Instr(',' || v_先作废医嘱ids || ',', ',' || a.医嘱序号 || ',') = 0) And
        Instr(',' || v_Order_Ids || ',', ',' || a.医嘱序号 || ',') > 0 And Mod(a.记录性质, 10) = n_记录性质 And a.记录状态 In (0, 1) And
        a.执行状态 = 9 And a.No In (Select /*+cardinality(X,10)*/
                                 x.Column_Value
                                From Table(f_Str2list(v_Nos)) X);

  If v_Tmpno Is Not Null Then
    v_Error := '医嘱费用单据"' || v_Tmpno || '"中的收费结算产生异常，不能作废。';
    Raise Err_Custom;
  End If;

  If n_记录性质 = 1 Then
    --已收费收费单
    --门诊收费单据判断是否已经收费，需排除自动取消，先作废后退药，药品卫材费用
    Select Max(a.No)
    Into v_Tmpno
    From 门诊费用记录 A, 材料特性 B
    Where a.收费细目id = b.材料id(+) And a.医嘱序号 Is Not Null And
          (v_先作废医嘱ids Is Null Or Instr(',' || v_先作废医嘱ids || ',', ',' || a.医嘱序号 || ',') = 0) And
          Instr(',' || v_Order_Ids || ',', ',' || a.医嘱序号 || ',') > 0 And Mod(a.记录性质, 10) = n_记录性质 And a.记录状态 = 1 And
          a.No In (Select /*+cardinality(X,10)*/
                    x.Column_Value
                   From Table(f_Str2list(v_Nos)) X);
  
    If v_Tmpno Is Not Null Then
      v_Error := '医嘱费用单据"' || v_Tmpno || '"已经收费，不能作废。';
      Raise Err_Custom;
    End If;
  End If;

  If v_先作废医嘱ids Is Not Null Then
    --先作废后退药的医嘱药品医嘱，要将这组医嘱中已收费的非药品卫材的费用执行状态改为未执行方便退费
    For R In (Select a.No, f_List2str(Cast(Collect(a.序号 || '') As t_Strlist), ',') 序号
              From 门诊费用记录 A, 材料特性 B
              Where a.收费细目id = b.材料id(+) And Instr(',' || v_先作废医嘱ids || ',', ',' || a.医嘱序号 || ',') > 0 And
                    Instr(',' || v_Order_Ids || ',', ',' || a.医嘱序号 || ',') > 0 And Mod(a.记录性质, 10) = n_记录性质 And
                    a.记录状态 = 1 And a.执行状态 = 1 And Not (a.收费类别 In ('5', '6', '7') Or a.收费类别 = '4' And b.跟踪在用 = 1) And
                    a.No In (Select /*+cardinality(X,10)*/
                              x.Column_Value
                             From Table(f_Str2list(v_Nos)) X)
              Group By a.No) Loop
    
      --修改执行状态
      v_修改执行状态 := v_修改执行状态 || ',{"fee_source":1';
      v_修改执行状态 := v_修改执行状态 || ',"fee_bill_type":1';
      v_修改执行状态 := v_修改执行状态 || ',"fee_no":"' || r.No || '"';
      v_修改执行状态 := v_修改执行状态 || ',"exe_sta_nums":"' || r.序号 || '"';
      v_修改执行状态 := v_修改执行状态 || '}';
    
    End Loop;
  End If;

  v_Del_List := Null;
  For R In c_Billdel Loop
    If r.记录性质 = 1 Then
      v_Serial_Num := r.序号串;
    Else
      v_Serial_Num := r.销记帐单;
    End If;
    --单据删除列表
    v_Del_List := v_Del_List || ',{"fee_source":1';
    v_Del_List := v_Del_List || ',"fee_bill_type":' || r.记录性质;
    v_Del_List := v_Del_List || ',"fee_no":"' || r.No || '"';
    v_Del_List := v_Del_List || ',"serial_num":"' || v_Serial_Num || '"';
    v_Del_List := v_Del_List || ',"exe_sta_nums":"' || r.序号串 || '"';
    v_Del_List := v_Del_List || '}';
  End Loop;

  If v_修改执行状态 Is Not Null Then
    v_Del_List := v_Del_List || v_修改执行状态;
  End If;

  v_Jtmp1 := Null;
  v_Jtmp1 := v_Jtmp1 || ',"exist_balance":' || Nvl(n_有已结帐费用, 0); --有已结帐的费用
  v_Jtmp1 := v_Jtmp1 || ',"exist_verify":' || Nvl(n_有记帐已审核, 0); --有已审核的记帐费用
  If Not v_Del_List Is Null Then
    v_Jtmp1 := v_Jtmp1 || ',"del_list":[' || Substr(v_Del_List, 2) || ']'; --待删除单据费用
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功"' || v_Jtmp1 || '}}';

Exception
  When Err_Custom Then
    Json_Out := Zljsonout(v_Error);
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkorderrevoke;
/


Create Or Replace Procedure Zl_Exsesvr_Checkbabyfee
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:检查婴儿是否已经产生费用
  --入参：Json_In:格式
  --   input
  --      pati_id           N 1 病人id
  --      pati_pageid       N 1 病人id
  --      baby_nums          C 1 婴儿序号,允许多个，用逗号分离;NULL表示查该病人的所有婴儿
  --出参  json
  --output
  --     code               N 1 应答码：0-失败；1-成功
  --     message            C 1 应答消息： 失败时返回具体的错误信息
  --     baby_nums          C 1 婴儿序号 :允许多个，用逗号分隔
  ---------------------------------------------------------------------------

  j_Input       PLJson;
  j_Json        PLJson;
  v_Output      Varchar2(4000);
  n_Pati_Id     Number(18);
  n_Pati_Pageid Number(5);
  v_Baby_Nums   Varchar2(2000);
  v_Babys       Varchar2(1000);
Begin
  --解析入参

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_Pati_Id     := j_Json.Get_Number('pati_id');
  n_Pati_Pageid := j_Json.Get_Number('pati_pageid');
  v_Baby_Nums   := j_Json.Get_Number('baby_nums');

  If n_Pati_Id Is Null Then
    Json_Out := zlJsonOut('未传入病人ID，请检查！');
    Return;
  End If;

  If n_Pati_Pageid Is Null Then
    Json_Out := zlJsonOut('未传入主页ID，请检查！');
    Return;
  End If;

  If v_Baby_Nums Is Null Then
    Json_Out := zlJsonOut('未传入婴儿序号，请检查！');
    Return;
  End If;
  v_Babys := Null;
  For R In (Select Distinct 婴儿费
            From 住院费用记录
            Where 病人id = n_Pati_Id And 主页id = n_Pati_Pageid And
                  ((Instr(',' || Nvl(v_Baby_Nums, '') || ',', ',' || Nvl(婴儿费, 0) || ',') > 0) Or
                  (v_Baby_Nums Is Null And Nvl(婴儿费, 0) > 0))) Loop
    v_Babys := Nvl(v_Babys, '') || ',' || r.婴儿费;
  End Loop;

  If v_Babys Is Not Null Then
    v_Babys := Substr(v_Babys, 2);
  End If;

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '成功');
  zlJsonPutValue(v_Output, 'baby_nums', v_Babys, 0, 2);
  Json_Out := '{"output":' || v_Output || '}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkbabyfee;
/


Create Or Replace Procedure Zl_Exsesvr_Updateoutprevstsign
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --功能：更新病人复诊标志
  --入参：Json_In:格式
  --input
  --  reg_id             N  1 挂号ID
  --  revst_sign         N  1 复诊标志 0:标记为初诊.1:标记为复诊 

  -- 出参:
  --  output
  --    code                            N 1 应答吗：0-失败；1-成功
  --    message                         C 1 应答消息：失败时返回具体的错误信息
  -------------------------------------------
  n_挂号id   Number;
  n_复诊标志 Number;

  j_Input PLJson;
  j_Json  PLJson;

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_挂号id   := j_Json.Get_Number('reg_id');
  n_复诊标志 := Nvl(j_Json.Get_Number('revst_sign'), 0);

  Zl_病人挂号记录_复诊(n_挂号id, n_复诊标志);
  Json_Out := zlJsonOut('成功', 1);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Updateoutprevstsign;
/
Create Or Replace Procedure Zl_Exsesvr_Checkorderroll
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：住院医嘱医嘱回退相关费用检查，同时获取费用相关信息
  --说明：费用单据都全销，部份销时不允操作
  --      处理费用执行状态，卫材药品取消执行，删除所有单据，异常处理
  --入参：Json_In:格式
  --  input
  --     outpati_account                N 1 门诊记帐
  --     bill_prop                      N 1 记录性质
  --     fee_nos                        C 1 格式：T000001,T000002,T000003...
  --     order_ids                      C 1 医嘱id串，本次处理的一批医嘱id逗号分割
  --     check_pacs                     N 0 检查类医嘱回退发送时生效，1-是否存在未审核的销帐申请

  --出参: Json_Out,格式如下
  --  output
  --    code                          N 1 应答吗：0-失败；1-成功
  --    message                       C 1 应答消息：失败时返回具体的错误信息
  --    isexist                       N 1 是否存在，0-不存，1-存在
  --    fee_nos                       C 1 逗号拼串，用于医保病人上传的no串
  --    del_list[]可以直接删除的记帐单
  --       fee_source                 N 1 费用来源:1-门诊费用记录，2-住院费用记录
  --       fee_bill_type              N 1 记录性质:1-收费单，2-记帐单
  --       fee_no                     C 1 费用单据号
  --       exe_sta_nums               C 1 序号串，用修正费用执行状态，格式：序号1,序号2,…
  --       serial_num                 C 1 两种格式，记录性质=1时格式：序号1,序号2,… 记录性质=2时格式：序号1:数量:执行状态1,序号2:数量2:执行状态2,…
  ---------------------------------------------------------------------------

  j_Input     Pljson;
  j_Json      Pljson;
  v_Nos       Varchar2(32767);
  v_Order_Ids Varchar2(32767);
  c_Order_Ids Clob;
  v_Del_List  Varchar2(32767);
  c_Del_List  Clob;

  c_Out_Tmp Clob;
  I         Number;
  v_Vals    Clob; --可能会长超的结点用此变量 v_vals , col_vals
  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_Vals     Collection_Type;
  n_记录性质   Number;
  n_门诊记帐   Number;
  n_Check_Pacs Number;
  v_Jtmp1      Varchar2(32767);
  v_Serial_Num Varchar2(32767);
  v_No_记帐    Varchar2(32767);
  v_Count      Number(5);
  v_No_已收费  Varchar2(300);
  v_Error      Varchar2(2000);
  Err_Custom Exception;

  --医保退费重收会有记录性质为两位的目前只有门诊病人发送的医嘱才会有，此处理不涉及
  Cursor c_Billdel Is
    Select a.记录性质, a.费用来源, a.No, Max(a.医保需要的单据) 医保需要的单据, f_List2str(Cast(Collect(a.序号 || '') As t_Strlist), ',') 序号串,
           f_List2str(Cast(Collect(a.销记帐单 || '') As t_Strlist), ',') 销记帐单序号串
    From (Select 1 费用来源, a.记录状态, a.No, a.序号, a.记录性质, Decode(a.记录性质 || a.记录状态, '21', a.No, Null) 医保需要的单据,
                  a.序号 || ':' || (Nvl(a.付数, 1) * a.数次) || ':0' 销记帐单
           From 门诊费用记录 A
           Where a.记录状态 In (0, 1) And a.记录性质 = n_记录性质 And Instr(',' || c_Order_Ids || ',', ',' || a.医嘱序号 || ',') > 0 And
                 Nvl(Nvl(a.付数, 1) * a.数次, 0) <> 0 And
                 a.No In (Select /*+cardinality(X,10)*/
                           x.Column_Value
                          From Table(f_Str2list(v_Nos)) X)
           Union All
           Select 2 费用来源, a.记录状态, a.No, a.序号, a.记录性质, Decode(a.记录性质 || a.记录状态, '21', a.No, Null) 医保需要的单据,
                  a.序号 || ':' || (Nvl(a.付数, 1) * a.数次) || ':0' 销记帐单
           From 住院费用记录 A
           Where a.记录状态 In (0, 1) And a.记录性质 = n_记录性质 And Instr(',' || c_Order_Ids || ',', ',' || a.医嘱序号 || ',') > 0 And
                 Nvl(Nvl(a.付数, 1) * a.数次, 0) <> 0 And
                 a.No In (Select /*+cardinality(X,10)*/
                           x.Column_Value
                          From Table(f_Str2list(v_Nos)) X)) A
    Group By a.记录性质, a.费用来源, a.No;

Begin

  --解析入参
  j_Input      := Pljson(Json_In);
  j_Json       := j_Input.Get_Pljson('input');
  n_记录性质   := j_Json.Get_Number('bill_prop');
  n_门诊记帐   := j_Json.Get_Number('outpati_account');
  n_Check_Pacs := j_Json.Get_Number('check_pacs');

  If n_Check_Pacs = 1 Then
    --主要是检查类型医嘱 D 绑定的药品要自动产生销帐申请
    v_Nos       := j_Json.Get_String('fee_nos');
    v_Order_Ids := j_Json.Get_String('order_ids');
    If n_门诊记帐 = 1 Then
      Select Count(1)
      Into v_Count
      From 门诊费用记录 C, 病人费用销帐 D
      Where c.Id = d.费用id And c.记录状态 In (0, 1, 3) And d.状态 = 0 And c.记录性质 = n_记录性质 And
            Instr(',' || v_Order_Ids || ',', ',' || c.医嘱序号 || ',') > 0 And
            c.No In (Select /*+cardinality(X,10)*/
                      x.Column_Value
                     From Table(f_Str2list(v_Nos)) X);
    Else
      Select Count(1)
      Into v_Count
      From 住院费用记录 C, 病人费用销帐 D
      Where c.Id = d.费用id And c.记录状态 In (0, 1, 3) And d.状态 = 0 And c.记录性质 = n_记录性质 And
            Instr(',' || v_Order_Ids || ',', ',' || c.医嘱序号 || ',') > 0 And
            c.No In (Select /*+cardinality(X,10)*/
                      x.Column_Value
                     From Table(f_Str2list(v_Nos)) X);
    End If;
  
    If v_Count > 0 Then
      v_Count := 1;
    End If;
    Json_Out := '{"output":{"code":1,"message":"成功","isexist":"' || v_Count || '"}}';
  Else
    c_Order_Ids := j_Json.Get_Clob('order_ids');
  
    --单据号分解----
    v_Vals := j_Json.Get_Clob('fee_nos');
    I      := 0;
    While v_Vals Is Not Null Loop
      If Length(v_Vals) <= 4000 Then
        Col_Vals(I) := v_Vals;
        v_Vals := Null;
      Else
        Col_Vals(I) := Substr(v_Vals, 1, Instr(v_Vals, ',', 3980) - 1);
        v_Vals := Substr(v_Vals, Instr(v_Vals, ',', 3980) + 1);
      End If;
      I := I + 1;
    End Loop;
    --单据号分解----
  
    For Lp In 0 .. Col_Vals.Count - 1 Loop
      v_Nos := Col_Vals(Lp);
      --判断数据转出时可以不加  医嘱序号  这个过滤条件也能达到相同效果
      --判断数据是否已经转出begin
      If n_门诊记帐 = 1 Then
        Select Count(1)
        Into v_Count
        From H门诊费用记录 A
        Where a.记录性质 = n_记录性质 And a.医嘱序号 Is Not Null And
              a.No In (Select /*+cardinality(X,10)*/
                        x.Column_Value
                       From Table(f_Str2list(v_Nos)) X);
      Else
        Select Count(1)
        Into v_Count
        From H住院费用记录 A
        Where a.记录性质 = n_记录性质 And a.医嘱序号 Is Not Null And
              a.No In (Select /*+cardinality(X,10)*/
                        x.Column_Value
                       From Table(f_Str2list(v_Nos)) X);
      End If;
      If v_Count > 0 Then
        v_Error := '该医嘱的费用已经全部或部份转出到后备数据库，不允许操作。' || Chr(13) || '您可以与系统管理员联系，将相应数据抽选返回。';
        Raise Err_Custom;
      End If;
      --判断数据是否已经转出end
    
      --未审核的销帐申请begin
      If n_门诊记帐 = 1 Then
        Select Count(1)
        Into v_Count
        From 门诊费用记录 C, 病人费用销帐 D
        Where c.Id = d.费用id And c.记录状态 In (0, 1, 3) And d.状态 = 0 And c.记录性质 = n_记录性质 And
              Instr(',' || c_Order_Ids || ',', ',' || c.医嘱序号 || ',') > 0 And
              c.No In (Select /*+cardinality(X,10)*/
                        x.Column_Value
                       From Table(f_Str2list(v_Nos)) X);
      Else
        Select Count(1)
        Into v_Count
        From 住院费用记录 C, 病人费用销帐 D
        Where c.Id = d.费用id And c.记录状态 In (0, 1, 3) And d.状态 = 0 And c.记录性质 = n_记录性质 And
              Instr(',' || c_Order_Ids || ',', ',' || c.医嘱序号 || ',') > 0 And
              c.No In (Select /*+cardinality(X,10)*/
                        x.Column_Value
                       From Table(f_Str2list(v_Nos)) X);
      End If;
      If v_Count > 0 Then
        v_Error := '该医嘱存在未审核的销帐申请，请取消或审核销帐申请后再回退发送。';
        Raise Err_Custom;
      End If;
      --未审核的销帐申请end
    
      --已收费的门诊收费单据判断begin
      Select Max(c.No)
      Into v_No_已收费
      From 门诊费用记录 C
      Where c.记录状态 = 1 And c.门诊标志 = 1 And c.记录性质 = 1 And Instr(',' || c_Order_Ids || ',', ',' || c.医嘱序号 || ',') > 0 And
            c.No In (Select /*+cardinality(X,10)*/
                      x.Column_Value
                     From Table(f_Str2list(v_Nos)) X);
      If v_No_已收费 Is Not Null Then
        v_Error := '该医嘱发送的门诊单据【' || v_No_已收费 || '】已收费，不能回退。';
        Raise Err_Custom;
      End If;
      --已收费的门诊收费单据判断end
    
      ----收集要删除的费用单据信息begin
      For R In c_Billdel Loop
      
        If r.记录性质 = 2 Then
          v_Serial_Num := r.销记帐单序号串;
        Else
          v_Serial_Num := r.序号串;
        End If;
      
        --收集医保需要的单据号串
        If r.医保需要的单据 Is Not Null Then
          If Instr(',' || v_No_记帐 || ',', ',' || r.医保需要的单据 || ',') = 0 Then
            v_No_记帐 := v_No_记帐 || ',' || r.医保需要的单据;
          End If;
        End If;
      
        --单据删除列表
        v_Del_List := v_Del_List || ',{"fee_source":' || r.费用来源;
        v_Del_List := v_Del_List || ',"fee_bill_type":' || r.记录性质;
        v_Del_List := v_Del_List || ',"fee_no":"' || r.No || '"';
        v_Del_List := v_Del_List || ',"serial_num":"' || v_Serial_Num || '"';
        v_Del_List := v_Del_List || ',"exe_sta_nums":"' || r.序号串 || '"';
        v_Del_List := v_Del_List || '}';
      
        If Length(v_Del_List) > 20000 Then
          If c_Del_List Is Null Then
            c_Del_List := v_Del_List;
          Else
            c_Del_List := c_Del_List || v_Del_List;
          End If;
          v_Del_List := Null;
        End If;
      End Loop;
      ----收集要删除的费用单据信息end    
    End Loop;
  
    v_Jtmp1 := Null;
    If Not v_No_记帐 Is Null Then
      v_Jtmp1 := v_Jtmp1 || ',"fee_nos":"' || Substr(v_No_记帐, 2) || '"'; --用于医嘱上传
    End If;
    c_Out_Tmp := v_Jtmp1;
  
    If Not v_Del_List Is Null Then
      c_Del_List := c_Del_List || v_Del_List;
      c_Del_List := ',"del_list":[' || Substr(c_Del_List, 2) || ']'; --待删除单据费用
    End If;
    c_Out_Tmp := c_Out_Tmp || c_Del_List;
    Json_Out  := '{"output":{"code":1,"message":"成功"' || c_Out_Tmp || '}}';
  End If;
Exception
  When Err_Custom Then
    Json_Out := Zljsonout(v_Error);
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkorderroll;
/


Create Or Replace Procedure Zl_Exsesvr_Billchargeoff
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：产生费用销帐申请数据
  --入参：Json_In:格式
  --  input
  --    request_operator                C 1 申请人  
  --    request_code                    C 0 申请人编号
  --    request_time                    C 1 申请时间
  --    request_type                    N 1 申请类别   
  --    del_tag                         N 1 删除标志
  --    reason                          C 1 销帐原因
  --    item_list[]用于生成销帐申请的列表
  --        fee_id                      N 1 费用ID
  --        request_dept_id             N 1 销帐申请科室ID
  --        fee_item_id                 N 1 收费细目ID
  --        quantity                    N 1 数次
  --        audit_dept_id               N 1 审核部门id
  --        auto_aduit                  N 0 是否自动审核
  --        outpati_account             N 0 是否门诊记帐        
  --        fee_no                      C 0 费用单据号
  --        serial_num                  N 0 序号：序号1:数量1:执行状态1,序号2:数量2:执行状态2,...序号n:数量n:执行状态n  如:"1:2:1,2:10:1,3:2:1"      
  --出参: Json_Out,格式如下
  --  output
  --    code          N 1 应答吗：0-失败；1-成功
  --    message       C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------

  Id_In         病人费用销帐.费用id%Type;
  收费细目id_In 病人费用销帐.收费细目id%Type;
  申请部门id_In 病人费用销帐.申请部门id%Type;
  数量_In       病人费用销帐.数量%Type;
  申请人_In     病人费用销帐.申请人%Type;
  申请人编号_In 病人费用销帐.申请人%Type;
  申请时间_In   病人费用销帐.申请时间%Type;
  申请类别_In   病人费用销帐.申请类别%Type;
  销帐原因_In   病人费用销帐.销帐原因%Type;
  审核部门id_In 病人费用销帐.审核部门id%Type;
  删除标志_In   Integer;
  审核时间_In   病人费用销帐.申请时间%Type;
  j_Input       PLJson;
  j_Json        PLJson;
  j_Item        PLJson;
  j_List        Pljson_List := Pljson_List();
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  申请人_In     := j_Json.Get_String('request_operator');
  申请人编号_In := j_Json.Get_String('request_code');
  申请时间_In   := To_Date(j_Json.Get_String('request_time'), 'yyyy-mm-dd hh24:mi:ss');
  申请类别_In   := j_Json.Get_Number('request_type');
  申请类别_In   := Nvl(申请类别_In, 0);
  删除标志_In   := j_Json.Get_Number('del_tag');
  删除标志_In   := Nvl(删除标志_In, 0);
  销帐原因_In   := j_Json.Get_String('reason');

  --如果是自动审核时这个时间会有用，为了将时间分开加一秒
  审核时间_In := 申请时间_In + 1 / 24 / 60 / 60;

  j_List := j_Json.Get_Pljson_List('item_list');
  For I In 1 .. j_List.Count Loop
    j_Item        := PLJson(j_List.Get(I));
    Id_In         := j_Item.Get_Number('fee_id');
    收费细目id_In := j_Item.Get_Number('fee_item_id');
    申请部门id_In := j_Item.Get_Number('request_dept_id');
    数量_In       := j_Item.Get_Number('quantity');
    审核部门id_In := j_Item.Get_Number('audit_dept_id');
    Zl_病人费用销帐_Insert_s(Id_In, 收费细目id_In, 申请部门id_In, 数量_In, 申请人_In, 申请时间_In, 申请类别_In, 销帐原因_In, 审核部门id_In, 删除标志_In);
    If 1 = j_Item.Get_Number('auto_aduit') Then
      Zl_病人费用销帐_Audit_s(Id_In, 申请时间_In, 申请人_In, 审核时间_In, 1, 申请类别_In);
      If 1 = j_Item.Get_Number('outpati_account') Then
        Zl_门诊记帐记录_Delete_s(j_Item.Get_String('fee_no'), j_Item.Get_String('serial_num'), 申请人编号_In, 申请人_In, 审核时间_In, 2);
      Else
        Zl_住院记帐记录_Delete_s(j_Item.Get_String('fee_no'), j_Item.Get_String('serial_num'), 申请人编号_In, 申请人_In, 2, 0,
                           审核时间_In);
      End If;
    End If;
  End Loop;
  Json_Out := zlJsonOut('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Billchargeoff;
/


Create Or Replace Procedure Zl_Exsesvr_Checkbillchargeoff
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：针对指定单据指定行行进行销帐申请检查，获取相关费用数据
  --入参：Json_In:格式
  --input
  --     oper_type                       N   1   操作类型，0-销帐申请列表检查传入item_list;
  --                                                       1-销帐申请表获取根据fee_ids+fee_source获取销帐费用明细
  --                                                       2-取消销帐申请获取有效的申请明细，传入fee_ids
  --     fee_ids                         C   1   费用IDs明细
  --     fee_source                      N   1   费用来源,1-门诊，2-住院
  --     item_list[]本次销帐列表
  --         fee_id                      N   1   费用ID
  --         request_dept_id             N   1   销帐申请科室ID
  --         item_id                     N   1   收费细目ID
  --         request_type                N   1   申请类别:对药品和卫材有效:0-未发药(料);1-已发药(料);其他为0
  --         request_num                 N   1   申请数量
  --         sended_num                  N   1   已发数量
  --    pati_list[]                 病人信息
  --         pati_id                     N   1   病人ID
  --         pati_name                   C   1   病人姓名
  --         pati_dept_id                N   1   出院科室id,病案主页.出院科室
  --         fee_audit_status            N   1   审核标志:病案主页.审核标志
  --         si_inp_status               N   1   住院状态:病案主页.状态(0-正常住院；1-尚未入科；2-正在转科或正在转病区；3-已预出院)
  --出参: Json_Out,格式如下
  --output
  --     code                            N   1   应答吗：0-失败；1-成功
  --     message                         C   1   应答消息：失败时返回具体的错误信息
  --     fee_ids                         C   1   费用明细id,oper_type=2返回
  --     item_list[]单据数据列表oper_type=0
  --         fee_id                      N   1   费用ID
  --         request_dept_id             N   1   销帐申请科室ID
  --         audit_dept_id               N   1   销帐审核科室ID
  --     charge_list[]销帐的费用明细列表oper_type=1
  --         pati_id                     N   1   病人id
  --         pati_pageid                 N   1   主页id
  --         fee_id                      N   1   费用id
  --         serial_num                  N   1   序号
  --         rec_status                  N   1   记录状态
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_操作类型   Number;
  v_Feeids     Varchar2(32767);
  v_Tmp        Varchar2(32767);
  n_Fee_Source Number;
Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_操作类型 := j_Json.Get_Number('oper_type');
  If Nvl(n_操作类型, 0) = 0 Then
    Zl_病人费用销帐_Insert_Check(Json_In, Json_Out);
  Elsif n_操作类型 = 1 Then
    v_Feeids     := j_Json.Get_String('fee_ids');
    n_Fee_Source := j_Json.Get_Number('fee_source');
    v_Tmp        := Null;
    If n_Fee_Source = 2 Then
      For R In (Select e.Id, e.序号, e.记录状态, e.病人id, e.主页id
                From 住院费用记录 E
                Where e.Id In (Select /*+cardinality(x,10)*/
                                x.Column_Value
                               From Table(Cast(f_Num2List(v_Feeids) As Zltools.t_Numlist)) X)) Loop
      
        v_Tmp := v_Tmp || ',{"pati_id":' || r.病人id;
        v_Tmp := v_Tmp || ',"pati_pageid":' || Nvl(r.主页id || '', 'null');
        v_Tmp := v_Tmp || ',"fee_id":' || r.Id;
        v_Tmp := v_Tmp || ',"serial_num":' || r.序号;
        v_Tmp := v_Tmp || ',"rec_status":' || Nvl(r.记录状态, 0);
        v_Tmp := v_Tmp || '}';
      
      End Loop;
    Else
      For R In (Select e.Id, e.序号, e.记录状态, e.病人id, e.主页id
                From 门诊费用记录 E
                Where e.Id In (Select /*+cardinality(x,10)*/
                                x.Column_Value
                               From Table(Cast(f_Num2List(v_Feeids) As Zltools.t_Numlist)) X)) Loop
        v_Tmp := v_Tmp || ',{"pati_id":' || r.病人id;
        v_Tmp := v_Tmp || ',"pati_pageid":' || Nvl(r.主页id || '', 'null');
        v_Tmp := v_Tmp || ',"fee_id":' || r.Id;
        v_Tmp := v_Tmp || ',"serial_num":' || r.序号;
        v_Tmp := v_Tmp || ',"rec_status":' || Nvl(r.记录状态, 0);
        v_Tmp := v_Tmp || '}';
      End Loop;
    End If;
    Json_Out := '{"output":{"code":1,"message":"成功","charge_list":[' || Substr(v_Tmp, 2) || ']}}';
  Elsif n_操作类型 = 2 Then
    v_Feeids := j_Json.Get_String('fee_ids');
    v_Tmp    := Null;
    For R In (Select Distinct e.费用id
              From 病人费用销帐 E
              Where e.审核人 Is Null And
                    e.费用id In (Select /*+cardinality(x,10)*/
                                x.Column_Value
                               From Table(Cast(f_Num2List(v_Feeids) As Zltools.t_Numlist)) X)) Loop
    
      v_Tmp := v_Tmp || ',' || r.费用id;
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"成功","fee_ids":"' || Substr(v_Tmp, 2) || '"}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkbillchargeoff;
/


Create Or Replace Procedure Zl_Exsesvr_Delbillchargeoff
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：删除销帐申请
  --入参：Json_In:格式
  --  input
  --       fee_ids               C 1 费用ids，费用id拼串
  --       request_time          C 1 销帐申请的时间                    
  --出参: Json_Out,格式如下
  --  output
  --       code                  N 1 应答吗：0-失败；1-成功
  --       message               C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  Zl_病人费用销帐_Delete_s(j_Json.Get_String('fee_ids'), To_Date(j_Json.Get_String('request_time'), 'yyyy-mm-dd hh24:mi:ss'));
  Json_Out := zlJsonOut('成功', 1);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Delbillchargeoff;
/


Create Or Replace Procedure Zl_Exsesvr_Checkshareinvoice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：判断是否存在指定的共享票据领用
  --入参：Json_In:格式
  --  input
  --   recv_id          N  1  领用id
  --   invc_type        N  1  票种
  --出参: Json_Out,格式如下
  --  output
  --    code                 N   1   应答吗：0-失败；1-成功
  --    message              C   1   应答消息：失败时返回具体的错误信息
  --    isexist              N   1   是否存在（0-不存在 1-存在）

  j_Input PLJson;
  j_Json  PLJson;
  n_Id    票据领用记录.Id%Type;
  n_Kind  票据领用记录.票种%Type;
  n_Count Number;

  v_Output  Varchar2(32767);
  n_Isexist Number(1);
Begin

  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_Kind := j_Json.Get_Number('invc_type');
  n_Id   := j_Json.Get_Number('recv_id');

  Select Count(1) Into n_Count From 票据领用记录 Where ID = n_Id And 票种 = n_Kind And 使用方式 = 2;
  If n_Count > 0 Then
    n_Isexist := 1;
  Else
    n_Isexist := 0;
  End If;

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '成功');
  zlJsonPutValue(v_Output, 'isexist', n_Isexist, 1, 2);
  Json_Out := '{"output":' || v_Output || '}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkshareinvoice;
/

Create Or Replace Procedure Zl_Exsesvr_Updpatbaseinfocheck
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ------------------------------------------------------------------------------------------
  --功能:更新费用相关业务数据的病人基本信息的检查
  --入参:JSON格式
  --input
  --    vist_id   N 1 就诊id串_ 门诊病人为挂号ID;住院病人为主页ID;为0说明是外来或非挂号就诊的病人(就诊id_In为空时,不更改该病人的费用部分的业务数据)
  --    occasion   N 1 场合,1-门诊;2-住院
  --出参:JSON格式
  --output
  --  code            N 1 应答码：0-失败；1-成功
  --  message        C 1 应答消息：失败时返回具体的错误信息
  --  explain   C 1 说明
  ------------------------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  d_Maxdate Date;
  v_No      门诊费用记录.No%Type;
  v_科室    部门表.名称%Type;
  v_说明    Clob;
  n_病人id  Number;
  n_就诊id  Number;
  n_场合    Number;
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_病人id := j_Json.Get_Number('pati_id');
  n_就诊id := j_Json.Get_Number('visit_id');
  n_场合   := j_Json.Get_Number('occasion');
  If Nvl(n_就诊id, 0) = 0 Then
    Return;
  End If;
  If Nvl(n_场合, 0) <= 1 Then
    --门诊
    If Nvl(n_就诊id, 0) <> 0 Then
      Begin
        Select a.No, b.名称, a.登记时间
        Into v_No, v_科室, d_Maxdate
        From 病人挂号记录 A, 部门表 B
        Where a.执行部门id = b.Id(+) And a.Id = n_就诊id;
      Exception
        When Others Then
          v_No := Null;
      End;
      If Not v_No Is Null Then
        v_说明 := '挂号单:' || v_No || LPad(' ', 4) || '挂号科室:' || v_科室 || ' 收回票据信息:';
      End If;
    Else
      v_说明 := '收回票据信息:';
    End If;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功","explain":"' || v_说明 || '"}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Updpatbaseinfocheck;
/


Create Or Replace Procedure Zl_Exsesvr_Getdynamiccost
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：读取动态费别
  --入参：Json_In:格式
  --  input
  --    dept_id             N 1 科室ID
  --出参: Json_Out,格式如下
  --  output
  --    code                N 1 应答吗：0-失败；1-成功
  --    message             C 1 应答消息：失败时返回具体的错误信息
  --    fee_category        C 1 费别拼串，逗号分割
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_科室id Number(18);
  v_费别   Varchar2(32767);
Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_科室id := j_Json.Get_Number('dept_id');
  For R In (Select 编码, 简码, 名称
            From 费别
            Where Nvl(属性, 1) = 2 And Nvl(适用科室, 1) = 1 And Nvl(服务对象, 3) In (1, 3) And
                  Trunc(Sysdate) Between Nvl(有效开始, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                  Nvl(有效结束, To_Date('3000-01-01', 'YYYY-MM-DD'))
            Union All
            Select Distinct a.编码, a.简码, a.名称
            From 费别 A, 费别适用科室 B
            Where a.名称 = b.费别 And b.科室id = n_科室id And Nvl(a.属性, 1) = 2 And Nvl(a.适用科室, 1) = 2 And
                  Nvl(a.服务对象, 3) In (1, 3) And Trunc(Sysdate) Between Nvl(a.有效开始, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                  Nvl(a.有效结束, To_Date('3000-01-01', 'YYYY-MM-DD'))
            Order By 编码) Loop
    v_费别 := v_费别 || ',' || r.名称;
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"成功","fee_category":"' || Substr(v_费别, 2) || '"}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getdynamiccost;
/

Create Or Replace Procedure Zl_Exsesvr_Getneedaudititems
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：获取指定病人的费用审批项目
  --入参：Json_In:格式
  --  input
  --       pati_id          N 1 病人ID
  --       pati_pageid      N 1 主页id
  --       fitem_id         N 1 项目ID
  --出参: Json_Out,格式如下
  --  output
  --       code             N 1 应答吗：0-失败；1-成功
  --       message          C 1 应答消息：失败时返回具体的错误信息
  --       item_list[]
  --          fitem_id         N 1 项目ID
  --          limit_quantity  N 1 使用限量
  --          used_quantity   N 1 已用数量
  --          avail_quantity  N 1 可用数量

  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_病人id Number(18);
  n_主页id Number(18);
  n_项目id Number(18);

  v_Output Varchar2(32767);
  c_Output Clob;

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_病人id := j_Json.Get_Number('pati_id');
  n_主页id := j_Json.Get_Number('pati_pageid');
  n_项目id := j_Json.Get_Number('fitem_id');

  For R In (Select 项目id, 使用限量, 已用数量, (使用限量 - 已用数量) As 可用数量
            From 病人审批项目
            Where 病人id = n_病人id And 主页id = n_主页id And (Nvl(n_项目id, 0) = 0 Or 项目id = Nvl(n_项目id, 0))) Loop
  
    If Length(Nvl(v_Output, ' ')) > 30000 Then
      If c_Output Is Not Null Then
        c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
      Else
        c_Output := To_Clob(v_Output);
      End If;
    
      v_Output := Null;
    End If;
  
    zlJsonPutValue(v_Output, 'item_id', r.项目id, 1, 1); --N
    zlJsonPutValue(v_Output, 'limit_quantity', r.使用限量, 1); --N
    zlJsonPutValue(v_Output, 'used_quantity', r.已用数量, 1); --N
    zlJsonPutValue(v_Output, 'avail_quantity', r.可用数量, 1, 2); --N
  
  End Loop;
  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"成功","item_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || v_Output || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getneedaudititems;
/



Create Or Replace Procedure Zl_Exsesvr_Adviceisexist
(
  Json_In  Varchar2,
  Json_Out Out Clob
) Is
  --根据医嘱ID查询是否在费用表存在记录
  ---------------------------------------------------------------------------
  --input      根据医嘱ID查询医嘱状态
  --  advice_ids  N  1  多个医嘱ID，用逗号分隔
  --output
  --  code              C  1  应答码：0-失败；1-成功
  --  message           C  1  应答消息：
  --  advice_list       医嘱列表[数组]
  --     advice_id      N    医嘱ID（存在费用的）
  --     fee_no         C    费用NO
  --     pati_id        N    病人id
  --     fee_properties N    记录性质
  --     fee_status     N    记录状态
  --     amount_id      N    结帐id
  --     nums           N    数次
  --     packages_num   N    付数
  --     parent_num     N    价格父号
  --     receipt_type   C    收费类别
  --     receipt_id     N    收费细目id
  --     cost_status    N    费用状态
  --     stdd_price     N    标准单价
  --     real_amount    N    实收金额

  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_医嘱id Clob; --记录医嘱id
  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_医嘱id Collection_Type;
  I          Number;

  v_Output Varchar2(32767);
  c_Output Clob;

Begin

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_医嘱id := j_Json.Get_String('advice_ids');

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
    For v_医嘱费用 In (Select Distinct 医嘱序号, 费用no, 病人id, 记录性质, 记录状态, 结帐id, 数次, 付数, 价格父号, 收费类别, 收费细目id, 费用状态, 标准单价, 实收金额
                   From (Select /*+cardinality(b,10)*/
                           医嘱序号, NO As 费用no, 病人id, 记录性质, 记录状态, 结帐id, 数次, 付数, 价格父号, 收费类别, 收费细目id, 费用状态, 标准单价, 实收金额
                          From 门诊费用记录 A, Table(f_Num2List(Col_医嘱id(I))) B
                          Where a.医嘱序号 = b.Column_Value
                          Union All
                          Select /*+cardinality(b,10)*/
                           医嘱序号, NO As 费用no, 病人id, 记录性质, 记录状态, 结帐id, 数次, 付数, 价格父号, 收费类别, 收费细目id, 费用状态, 标准单价, 实收金额
                          From 住院费用记录 A, Table(f_Num2List(Col_医嘱id(I))) B
                          Where a.医嘱序号 = b.Column_Value)) Loop
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      
        v_Output := Null;
      End If;
    
      zlJsonPutValue(v_Output, 'advice_id', v_医嘱费用.医嘱序号, 1, 1);
      zlJsonPutValue(v_Output, 'fee_no', v_医嘱费用.费用no);
      zlJsonPutValue(v_Output, 'pati_id', v_医嘱费用.病人id, 1);
      zlJsonPutValue(v_Output, 'fee_properties', v_医嘱费用.记录性质, 1);
      zlJsonPutValue(v_Output, 'fee_status', v_医嘱费用.记录状态, 1);
      zlJsonPutValue(v_Output, 'amount_id', v_医嘱费用.结帐id, 1);
      zlJsonPutValue(v_Output, 'nums', v_医嘱费用.数次, 1);
      zlJsonPutValue(v_Output, 'packages_num', v_医嘱费用.付数, 1);
      zlJsonPutValue(v_Output, 'parent_num', v_医嘱费用.价格父号, 1);
      zlJsonPutValue(v_Output, 'receipt_type', v_医嘱费用.收费类别);
      zlJsonPutValue(v_Output, 'receipt_id', v_医嘱费用.收费细目id, 1);
      zlJsonPutValue(v_Output, 'cost_status', v_医嘱费用.费用状态, 1);
      zlJsonPutValue(v_Output, 'stdd_price', v_医嘱费用.费用状态, 1);
      zlJsonPutValue(v_Output, 'real_amount', v_医嘱费用.实收金额, 1, 2);
    
    End Loop;
  End Loop;
  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"成功","advice_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"成功","advice_list":[' || v_Output || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Adviceisexist;
/

Create Or Replace Procedure Zl_Exsesvr_Getmrbkfeeinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:根据病人id获取指定病人所涉及的病历费清单信息
  --入参：Json_In:格式
  --  input
  --    pati_id  N  1  病人id
  --    fee_no  C    单据号:病历费所涉及的单据号
  --    rec_status  N  1  记录状态:1-原始记录;2-销帐数据
  --
  --出参: Json_Out,格式如下
  --output
  --   code  C  1  应答码：0-失败；1-成功
  --   message  C  1  应答消息：
  --   details_list[]  C    费用明细数据
  --    fee_no  C  1  单据号
  --    fee_num  N  1  序号
  --    pati_id  N  1  病人id
  --    pati_name  C  1  姓名
  --    pati_sex  C  1  性别
  --    pati_age  C  1  年龄
  --    fee_category  C  1  费别
  --    fee_status  N  1  费用状态:1-异常状态;0-正常费用
  --    rec_status  N  1  记录状态:1-正常记录;2-销帐记录;3-补销帐的记录
  --    charge_sign  N  1  收费标志:0-现收;1-记帐;2-划价单
  --    fee_ampaid  N  1  实收金额
  --    happen_time  C  1  发生时间:yyyy-mm-dd hh24:mi:ss
  --    operator_name  C  1  操作员姓名
  --    memo  C  1  摘要
  --    pricebill_no  C  1  划价单号
  --    price_charged  N  1  划价已收费:1-划价单已经在收费窗口收费;0-未收费
  --    balance_info  C    结算信息
  --      blnc_mode  C  1  结算方式名称
  --      balance_id  N  1  结帐ID：查询作废的单据时为冲销ID
  --      blnc_money  N  1  结帐金额
  --      pay_cardno  N  1  支付卡号
  --      pay_swapno  C  1  交易流水号
  --      pay_swapmemo  C  1  交易说明
  --      relation_id  N  1  关联交易id
  --      cardtype_id  N  1  卡类别id
  --      consume_card  N  1  是否消费卡:1-是;0-不是
  --      blnc_nature  N  1  结算性质:1-现金结算方式,2-其他非医保结算 , 8-结算卡结算
  --      error_moeny N 1 误差金额
  --      blnc_statu  N  1  结算状态:1-未调用接口;2-接口调用成功,但还未收费完成,0-正常结算
  --      consume_card_id  N  1  消费卡id
  --      blnc_no  C  1  结算号码
  --      blnc_memo  C  1  摘要
  --      original_id  N  1  原结帐ID:冲销时返回
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  o_Json PLJson;

  n_病人id   Number;
  v_单据号   门诊费用记录.No%Type;
  n_记录状态 门诊费用记录.记录状态%Type;
  n_实收金额 门诊费用记录.实收金额%Type;
  v_记录状态 Varchar2(6);
  v_划价单   Varchar2(100);
  n_划价已收 Number(2);
  n_收费标志 Number(2);
  n_结帐id   Number(18);
  n_原结帐id Number(18);
  n_Count    Number(18);
  v_Output   Varchar2(32767);
  v_Balance  Varchar2(32767);
  c_Output   Clob;

Begin

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_病人id   := j_Json.Get_Number('pati_id');
  v_单据号   := j_Json.Get_String('fee_no');
  n_记录状态 := j_Json.Get_Number('rec_status');

  v_记录状态 := ',1,';
  If Nvl(n_记录状态, 0) = 2 Then
    v_记录状态 := ',2,';
  End If;

  --先读取费用
  v_划价单 := Null;
  For r_费用 In (Select a.No, a.记录状态, Nvl(a.价格父号, a.序号) As 序号, Max(a.费别) As 费别, Max(a.姓名) As 姓名, Max(a.性别) As 性别,
                      Max(a.年龄) As 年龄, Max(a.病人id) As 病人id, a.收费细目id, Max(a.实际票号) As 实际票号, Avg(a.数次) As 数次,
                      Sum(Decode(n_记录状态, 2, -1, 1) * a.应收金额) As 应收金额, Sum(Decode(n_记录状态, 2, -1, 1) * a.实收金额) As 实收金额,
                      Nvl(Max(a.记帐费用), 0) As 记帐费用, Nvl(Max(Nvl(a.加班标志, 0)), 0) As 变动类别, Max(a.操作员姓名) As 操作员姓名,
                      Max(a.操作员编号) As 操作员编号, To_Char(Max(a.登记时间), 'yyyy-mm-dd hh24:mi:ss') As 登记时间,
                      To_Char(Max(a.发生时间), 'yyyy-mm-dd hh24:mi:ss') As 发生时间, Decode(Nvl(Max(a.附加标志), 0), 8, 1, 0) As 病历费,
                      Max(a.结论) As 卡类别id, Max(a.结帐id) As 结帐id, Max(a.开单人) As 开单人, Max(a.费用状态) As 费用状态, Max(a.摘要) As 摘要,
                      Max(a.实际票号) As 卡号
               From 住院费用记录 A
               Where a.记录性质 = 5 And a.附加标志 = 8 And Instr(v_记录状态, ',' || a.记录状态 || ',') > 0 And 病人id = n_病人id And
                     ((v_单据号 Is Not Null And NO = v_单据号) Or v_单据号 Is Null)
               Group By a.No, a.记录状态, Nvl(a.价格父号, a.序号), 收费细目id
               Order By a.No, 序号) Loop
    n_原结帐id := Null;
    v_划价单   := Null;
    n_结帐id   := r_费用.结帐id;
    If Nvl(n_记录状态, 0) = 2 And Nvl(r_费用.记帐费用, 0) = 0 Then
      Select Max(结帐id)
      Into n_原结帐id
      From 住院费用记录
      Where 记录性质 = 5 And 记录状态 In (1, 3) And NO = r_费用.No;
    End If;
  
    o_Json     := PLJson();
    n_实收金额 := Nvl(r_费用.实收金额, 0);
    n_收费标志 := Nvl(r_费用.记帐费用, 0);
    n_划价已收 := 0;
    If r_费用.摘要 Is Not Null And Nvl(n_收费标志, 0) <> 1 Then
      v_划价单 := r_费用.摘要;
      --一个病人应该不多，所以单独查询，性能影响不大
      Select Count(1), Sum(实收金额), Decode(Max(记录状态), 0, 0, 1)
      Into n_Count, n_实收金额, n_划价已收
      From 门诊费用记录
      Where NO = r_费用.No And 记录性质 = 1 And Nvl(附加标志, 0) = 8;
    
      If n_Count <> 0 Then
        n_收费标志 := 2; --划价
      
      Else
        n_实收金额 := Nvl(r_费用.实收金额, 0);
        v_划价单   := Null;
      End If;
    
    End If;
  
    If Length(Nvl(v_Output, ' ')) > 30000 Then
      c_Output := Nvl(c_Output, '') || To_Clob(v_Output);
      v_Output := '';
    End If;
  
    --1.取基本信息
    zlJsonPutValue(v_Output, 'fee_no', r_费用.No, 0, 1);
    zlJsonPutValue(v_Output, 'fee_num', r_费用.序号, 1);
  
    zlJsonPutValue(v_Output, 'pati_id', r_费用.病人id, 1);
  
    zlJsonPutValue(v_Output, 'pati_name', r_费用.姓名);
    zlJsonPutValue(v_Output, 'pati_sex', Nvl(r_费用.性别, ''));
  
    zlJsonPutValue(v_Output, 'pati_age', Nvl(r_费用.年龄, ''));
  
    zlJsonPutValue(v_Output, 'fee_category', Nvl(r_费用.费别, ''));
  
    zlJsonPutValue(v_Output, 'fee_status', Nvl(r_费用.费用状态, 0), 1);
    zlJsonPutValue(v_Output, 'rec_status', Nvl(r_费用.记录状态, 0), 1);
    zlJsonPutValue(v_Output, 'kpbooks_sign', Nvl(n_收费标志, 0), 1);
    zlJsonPutValue(v_Output, 'fee_ampaid', Nvl(n_实收金额, 0), 1);
    zlJsonPutValue(v_Output, 'happen_time', Nvl(r_费用.发生时间, ''));
    zlJsonPutValue(v_Output, 'operator_name', Nvl(r_费用.操作员姓名, ''));
    zlJsonPutValue(v_Output, 'memo', Nvl(r_费用.摘要, ''));
    zlJsonPutValue(v_Output, 'pricebill_no', Nvl(v_划价单, ''));
  
    zlJsonPutValue(v_Output, 'price_charged', Nvl(n_划价已收, 0), 1);
  
    --读取结算信息
    If Nvl(r_费用.记帐费用, 0) = 0 Then
      v_Balance := '';
      For r_结算信息 In (
                     
                     Select a.No, Max(Decode(Nvl(b.性质, 0), 9, '', a.结算方式)) As 结算方式,
                             Sum(Decode(Nvl(b.性质, 0), 9, 0, 1) * Decode(n_记录状态, 2, -1, 1) * Nvl(a.冲预交, 0)) As 冲预交,
                             Sum(Decode(Nvl(b.性质, 0), 9, 1, 0) * Decode(n_记录状态, 2, -1, 1) * Nvl(a.冲预交, 0)) As 误差费,
                             Max(Decode(Nvl(b.性质, 0), 9, 0, a.关联交易id)) As 关联交易id, Max(a.卡类别id) As 卡类别id, Max(a.卡号) As 卡号,
                             Max(a.结算卡序号) As 结算卡序号, Max(a.交易流水号) As 交易流水号, Max(a.交易说明) As 交易说明,
                             Max(Decode(Nvl(b.性质, 0), 9, -1, Nvl(b.性质, 0))) As 性质, Max(a.校对标志) As 校对标志,
                             Max(c.消费卡id) As 消费卡id, Max(a.结算号码) As 结算号码, Max(a.摘要) As 摘要
                     From 病人预交记录 A, 结算方式 B, 病人卡结算记录 C
                     Where a.结算方式 = 名称(+) And a.结帐id = n_结帐id And Mod(a.记录性质, 10) <> 1 And a.Id = c.结算id(+)
                     Group By NO) Loop
      
        zlJsonPutValue(v_Balance, 'blnc_mode', r_结算信息.结算方式, 0, 1);
        zlJsonPutValue(v_Balance, 'balance_id', n_结帐id, 1);
        zlJsonPutValue(v_Balance, 'blnc_money', Nvl(r_结算信息.冲预交, 0), 1);
        zlJsonPutValue(v_Balance, 'pay_cardno', Nvl(r_结算信息.卡号, ''));
        zlJsonPutValue(v_Balance, 'pay_swapno', Nvl(r_结算信息.交易流水号, ''));
        zlJsonPutValue(v_Balance, 'pay_swapmemo', Nvl(r_结算信息.交易说明, ''));
      
        zlJsonPutValue(v_Balance, 'relation_id', Nvl(r_结算信息.关联交易id, 0), 1);
      
        If Nvl(r_结算信息.结算卡序号, 0) <> 0 Then
          zlJsonPutValue(v_Balance, 'cardtype_id', Nvl(r_结算信息.结算卡序号, 0), 1);
          zlJsonPutValue(v_Balance, 'consume_card', 1, 1);
        Else
          zlJsonPutValue(v_Balance, 'cardtype_id', Nvl(r_结算信息.卡类别id, 0), 1);
          zlJsonPutValue(v_Balance, 'consume_card', 0, 1);
        End If;
      
        zlJsonPutValue(v_Balance, 'consume_card_id', Nvl(r_结算信息.消费卡id, 0), 1);
        zlJsonPutValue(v_Balance, 'error_moeny', Nvl(r_结算信息.误差费, 0), 1);
        zlJsonPutValue(v_Balance, 'blnc_nature', Nvl(r_结算信息.性质, 0), 1);
        zlJsonPutValue(v_Balance, 'blnc_statu', Nvl(r_结算信息.校对标志, 0), 1);
        zlJsonPutValue(v_Balance, 'blnc_no', Nvl(r_结算信息.结算号码, ''));
        zlJsonPutValue(v_Balance, 'blnc_memo', Nvl(r_结算信息.摘要, ''));
        zlJsonPutValue(v_Balance, 'original_id', Nvl(n_原结帐id, 0), 1, 2);
        v_Balance := ',"balance_info":' || v_Balance;
        Exit;
      End Loop;
    Else
      v_Balance := Null;
    End If;
    v_Output := v_Output || Nvl(v_Balance, '') || '}';
  
  End Loop;
  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"成功","fee_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"成功","fee_list":[' || v_Output || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getmrbkfeeinfo;
/

Create Or Replace Procedure Zl_Exsesvr_Checkunauditedfee
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  --------------------------------------------------------------------------------------------------------------------
  --功能：检查病人是否存在尚未生效的记账项目
  --入参：Json_In,格式如下
  --  input
  --    pati_id            N 1 病人ID
  --    pati_pageid        N 1 主页ID
  --出参: Json_Out,格式如下
  --  output
  --    code               N 1 应答吗：0-失败；1-成功
  --    message            C 1 应答消息：失败时返回具体的错误信息
  --    exist              N 1 执行标记:0-不存在;1-存在
  --------------------------------------------------------------------------------------------------------------------
  j_Input       PLJson;
  j_Json        PLJson;
  v_Output      Varchar2(400);
  n_Count       Number;
  n_Pati_Id     Number(18);
  n_Pati_Pageid Number(18);
Begin

  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_Pati_Id     := j_Json.Get_Number('pati_id');
  n_Pati_Pageid := j_Json.Get_Number('pati_pageid');

  --检查该病人本次住院是否已经计算过
  Select Count(*)
  Into n_Count
  From 住院费用记录
  Where 病人id = n_Pati_Id And 主页id = n_Pati_Pageid And 记帐费用 = 1 And 记录状态 = 0 And Rownum <= 2;

  If n_Count > 0 Then
    n_Count := 1;
  End If;

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '成功');
  zlJsonPutValue(v_Output, 'exist', n_Count, 1, 2);
  Json_Out := '{"output":' || v_Output || '}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkunauditedfee;
/


Create Or Replace Procedure Zl_Exsesvr_Getfeebillbycardno
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:根据卡号获取费用单据信息
  --入参：Json_In:格式
  --   input
  --    pati_id  N  1  病人id
  --    cardtype_id  N  1  卡类别id
  --    cardno  C  1  卡号

  --出参: Json_Out,格式如下
  --  output
  --    code  C  1  应答码：0-失败；1-成功
  --    message  C  1  应答消息：失败时返回具体的错误信息
  --    feeno  C  1  卡费单号
  --    charge_sign  N  1  收费标志:1-已经收费用;2-已经退费

  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_病人id   Number(18);
  n_卡类别id Number(18);
  v_卡号     Varchar2(100);
  v_单据号   住院费用记录.No%Type;
  n_记录状态 住院费用记录.记录状态%Type;
  v_摘要     住院费用记录.摘要%Type;
  n_记帐费用 住院费用记录.记帐费用%Type;
  v_划价单   住院费用记录.No%Type;
  n_Temp     Number(18);
  n_Count    Number(18);
  n_收费标志 Number(2);

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_病人id   := j_Json.Get_Number('pati_id');
  n_卡类别id := j_Json.Get_Number('cardtype_id');
  v_卡号     := j_Json.Get_String('cardno');

  n_收费标志 := 1;
  Select Max(NO), Max(记录状态), Max(摘要), Max(记帐费用)
  Into v_单据号, n_记录状态, v_摘要, n_记帐费用
  From 住院费用记录
  Where 记录性质 = 5 And 病人id = n_病人id And 实际票号 = v_卡号 And To_Number(Nvl(结论, '0')) = Nvl(n_卡类别id, 0) And Nvl(附加标志, 0) <> 8;

  If Nvl(n_记录状态, 0) = 0 Then
    n_收费标志 := 0;
  Elsif Nvl(n_记录状态, 0) = 1 Then
  
    n_收费标志 := 1;
  Else
    n_收费标志 := 2;
  End If;

  If Nvl(n_记帐费用, 0) <> 1 And v_摘要 Is Not Null And Nvl(n_记录状态, 0) = 1 Then
  
    Select Max(记录状态), Max(NO)
    Into n_Temp, v_划价单
    From 门诊费用记录
    Where 记录性质 = 1 And NO = v_摘要 And 价格父号 Is Null And Nvl(附加标志, 0) <> 8;
    If v_划价单 Is Not Null Then
      If Nvl(n_Temp, 0) = 0 Then
        n_收费标志 := 0; --未收费
      Elsif Nvl(n_Temp, 0) <> 1 Then
        Select Count(1)
        Into n_Count
        From (Select NO, 序号, Sum(Nvl(付数, 1) * 数次) As 剩余数, Max(记录状态) As 记录状态
               From 门诊费用记录
               Where Mod(记录性质, 10) = 1 And NO = v_摘要 And 价格父号 Is Null And Nvl(附加标志, 0) <> 8 Having
                Sum(Nvl(付数, 1) * 数次) <> 0
               Group By NO, 序号);
        If n_Count = 0 Then
          n_收费标志 := 2; --已退费
        End If;
      Else
        n_收费标志 := 1; --已收费
      End If;
    End If;
  End If;
  Json_Out := '{"output":{"code":1,"message":"' || '成功' || '"';
  Json_Out := Json_Out || ',"feeno":"' || Nvl(v_单据号, '') || '"';
  Json_Out := Json_Out || ',"priceno":"' || Nvl(v_划价单, '') || '"';
  Json_Out := Json_Out || ',"charge_sign":' || Nvl(n_收费标志, 0);
  Json_Out := Json_Out || '}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getfeebillbycardno;
/

Create Or Replace Procedure Zl_Exsesvr_Registerinpatient
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  ---------------------------------------------------------------------------
  --功能：病人入院登记费用相关处理
  --入参：Json_In:格式
  --    input
  --      pati_id            N 1  病人id
  --      pati_pageid        N 1  主页ID
  --      type               N 1  登记模式=0-正常登记,1-预约登记,2-接收预约
  --      pati_deptid        N 1 入院科室ID
  --出参: Json_Out,格式如下
  --    output
  --        code                    N   1   应答吗：0-失败；1-成功
  --        message                 C   1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_Pati_Id     Number(18);
  n_Pati_Pageid Number(18);
  n_Type        Number(3);
  n_Pati_Deptid Number(18);
Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_Pati_Id     := j_Json.Get_Number('pati_id');
  n_Pati_Pageid := j_Json.Get_Number('pati_pageid');
  n_Type        := j_Json.Get_Number('type');
  n_Pati_Deptid := j_Json.Get_Number('pati_deptid');

  If n_Type <> 1 Then
    --病人担保记录
    Update 病人担保记录
    Set 到期时间 = Sysdate
    Where 病人id = n_Pati_Id And 到期时间 Is Not Null And 到期时间 > Sysdate;
    --病人费用审批项目
    Delete From 病人审批项目 Where 病人id = n_Pati_Id;
  End If;

  If n_Type = 2 Then
    Update 病人预交记录
    Set 主页id = n_Pati_Pageid
    Where 病人id = n_Pati_Id And 主页id Is Null And 科室id = n_Pati_Deptid And 预交类别 = 2 And 冲预交 Is Null And
          Trunc(收款时间) = Trunc(Sysdate);
  End If;
  Json_Out := zlJsonOut('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Registerinpatient;
/

Create Or Replace Procedure Zl_Exsesvr_Unregisterinpatient
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --------------------------------------------------------------------------- 
  --功能：取消病人入院登记费用相关处理 
  --入参：Json_In:格式 
  --    input 
  --      pati_id            N 1  病人id 
  --      pati_pageid        N 1  主页ID 
  --出参: Json_Out,格式如下 
  --    output 
  --        code                    N   1   应答吗：0-失败；1-成功 
  --        message                 C   1   应答消息：失败时返回具体的错误信息 
  --------------------------------------------------------------------------- 
  j_Input Pljson;
  j_Json  Pljson;

  n_Pati_Id     Number(18);
  n_Pati_Pageid Number(18);
  n_Count       Number;
  n_Money       Number(16, 5);
Begin
  --解析入参 
  j_Input := Pljson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_Pati_Id     := j_Json.Get_Number('pati_id');
  n_Pati_Pageid := j_Json.Get_Number('pati_pageid');

  Select Sum(金额) Into n_Money From 病人预交记录 Where 病人id = n_Pati_Id And 主页id = n_Pati_Pageid;
  If n_Money <> 0 Then
    Json_Out := Zljsonout('病人本次住院有预交款未处理,请处理后再执行此操作。');
    Return;
  End If;
  Select Sum(金额) Into n_Money From 病人未结费用 Where 病人id = n_Pati_Id And 主页id = n_Pati_Pageid;
  If n_Money <> 0 Then
    Json_Out := Zljsonout('病人本次住院有未结费用,请处理后再执行此操作。');
    Return;
  End If;

  --本次住院如果交了预交款,改为当作门诊交的 
  Update 病人预交记录 Set 主页id = Null Where 病人id = n_Pati_Id And 主页id = n_Pati_Pageid;

  --本次发卡的,改变门诊发卡 
  Update 住院费用记录 Set 主页id = Null Where 病人id = n_Pati_Id And 主页id = n_Pati_Pageid And 记录性质 = 5;

  --本次住院的所有费用记录无结算且已全部冲销，则将对应费用记录中的"主页ID"清除。 
  n_Count := 0;
  Select Nvl(Count(*), 0)
  Into n_Count
  From 住院费用记录
  Where 病人id = n_Pati_Id And 主页id = n_Pati_Pageid And 记帐费用 = 1 And 结帐id Is Not Null;

  If n_Count = 0 Then
    Begin
      Select Nvl(Count(*), 0)
      Into n_Count
      From 住院费用记录
      Where 病人id = n_Pati_Id And 主页id = n_Pati_Pageid And 记帐费用 = 1
      Group By NO, 记录性质, 序号
      Having Nvl(Sum(实收金额), 0) <> 0;
    Exception
      When Others Then
        n_Count := 0;
    End;
  
    If n_Count = 0 Then
      Delete 病人未结费用 Where 病人id = n_Pati_Id And 主页id = n_Pati_Pageid And 金额 = 0;
      Update 住院费用记录 Set 主页id = Null Where 病人id = n_Pati_Id And 主页id = n_Pati_Pageid And 记帐费用 = 1;
    End If;
  End If;
  Json_Out := Zljsonout('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Unregisterinpatient;
/


Create Or Replace Procedure Zl_Exsesvr_Updatepatisurety
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  ---------------------------------------------------------------------------
  --功能：病人担保记录新增、更新、删除处理
  --入参：Json_In:格式
  --    input
  --      func_id          N 1 功能ID 1-新增;2-更新;3-删除
  --      pati_id          N 1 病人id
  --      pati_pageid      N 1 主页ID
  --      guarantor        c 1 担保人
  --      garnt_amount     N 1 担保额
  --      garnt_prop       N 1 担保性质
  --      garnt_reason     c 1 担保原因
  --      due_time         c 1 到期时间
  --      operator_code    c 1 操作员编号
  --      operator_name    c 1 操作员姓名
  --      create_time      C 0 登记时间   更新时传入此值
  --出参: Json_Out,格式如下
  --    output
  --        code                    N   1   应答吗：0-失败；1-成功
  --        message                 C   1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Input       PLJson;
  j_Json        PLJson;
  n_Func_Id     Number(3);
  n_Pati_Id     Number(18);
  n_Pati_Pageid Number(18);

  v_担保人   病人担保记录.担保人%Type;
  n_担保额   病人担保记录.担保额 %Type;
  n_担保性质 病人担保记录.担保性质%Type;
  v_担保原因 病人担保记录.担保原因%Type;

  d_到期时间   病人担保记录.到期时间%Type;
  v_操作员编号 病人担保记录.操作员编号%Type;
  v_操作员姓名 病人担保记录.操作员姓名%Type;
  d_登记时间   病人担保记录.登记时间%Type;

Begin
  --解析入参

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_Func_Id     := j_Json.Get_Number('func_id');
  n_Pati_Id     := j_Json.Get_Number('pati_id');
  n_Pati_Pageid := j_Json.Get_Number('pati_pageid');
  If n_Func_Id = 1 Or n_Func_Id = 2 Then
    v_担保人   := j_Json.Get_String('guarantor');
    n_担保额   := j_Json.Get_Number('garnt_amount');
    n_担保性质 := j_Json.Get_Number('garnt_prop');
    v_担保原因 := j_Json.Get_String('garnt_reason');
    d_到期时间 := To_Date(j_Json.Get_String('due_time'), 'YYYY-MM-DD HH24:MI:SS');
  End If;
  v_操作员编号 := j_Json.Get_String('operator_code');
  v_操作员姓名 := j_Json.Get_String('operator_name');
  If n_Func_Id = 2 Or n_Func_Id = 3 Then
    d_登记时间 := To_Date(j_Json.Get_String('create_time'), 'YYYY-MM-DD HH24:MI:SS');
  End If;
  If n_Func_Id = 1 Then
    Zl_病人担保记录_Insert(n_Pati_Id, n_Pati_Pageid, v_担保人, n_担保额, n_担保性质, v_担保原因, Null, d_到期时间, v_操作员编号, v_操作员姓名);
  Elsif n_Func_Id = 2 Then
    Zl_病人担保记录_Update(n_Pati_Id, n_Pati_Pageid, v_担保人, n_担保额, n_担保性质, v_担保原因, Null, d_到期时间, v_操作员编号, v_操作员姓名, d_登记时间);
  Elsif n_Func_Id = 3 Then
    Zl_病人担保记录_Delete(n_Pati_Id, n_Pati_Pageid, d_登记时间, v_操作员编号, v_操作员姓名);
  End If;

  Json_Out := zlJsonOut('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Updatepatisurety;
/

Create Or Replace Procedure Zl_Exsesvr_Getorderfeeinfo
(
  Json_In  In Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：获取费用相关信息
  --入参：Json_In:格式
  --  input
  --     query_type         N 1 查询方式必须传入
  --                               1-查询药占比,所需入参结点（pati_id+pati_pageid），返回 fee_ratio 所占比例：如 68.5%，保留一位小数，字符型
  --                               2-根据医嘱id获取未记帐费用的医嘱,所需入参结点（pati_id+pati_pageid+baby_num）
  --                               3-根据医嘱id获取门诊相关汇总金额,所需入参结点（order_ids）
  --                               4-根据医嘱id获取费用记录汇总信息医嘱清单下方的发送列表，,所需入参结点（fee_origin+order_ids）
  --                               5-根据药品医嘱ID获取最近一次发送的费用对应的收费项目ID(药品规格ID)，返回一个收费细目id,所需入参结点（fee_no+bill_prop+order_id+fee_origin）
  --                               6-获取医嘱对应未审核的记帐费用合计，所需入参结点（fee_origin+fee_no），此时fee_no为多个单据号拼串
  --                               7-根据费用id判断是否已已经收费，所需入参结点(fee_origin+fee_ids)返回已收费的费用id逗号拼串
  --                               8-根据费用来源和医嘱id获取费用明细列表，所需入参结点(fee_origin+fee_no+order_ids)，医技站执行科室变更时会调用
  --                               9-根据费用来源和医嘱id获取费用明细列表，所需入参结点(fee_origin+fee_no+order_ids)，护士站销帐后费用出现异常再次进行异常修复时获取数据                             
  --     pati_id            N 0 病人id
  --     pati_pageid        N 0 主页id
  --     baby_num           N 0 婴儿序号
  --     order_ids          C 0 医嘱ID拼串
  --     fee_origin         N 0 费用来源(默认=2：1-门诊费用，2-住院费用)
  --     fee_no             C 0 单据号
  --     bill_prop          N 0 记录性质
  --     order_id           N 0 医嘱id
  --     fee_ids            C 0 费用id逗号拼串
  --出参: Json_Out,格式如下
  --  output
  --    code                    N 1 应答吗：0-失败；1-成功
  --    message                 C 1 应答消息：失败时返回具体的错误信息
  --    fee_ratio               C 1 药品费所占比例：如 68.5%，保留一位小数，字符型[query_type=1时返回此结点]
  --    order_ids               C 1 医嘱ID拼串，[query_type=2时返回此结点]
  --    fee_ampaib              N 1 实收金额，[query_type=6时返回此结点]
  --    fee_ids                 C 0 费用id逗号拼串，[query_type=7时返回此结点]
  --    fee_am_list[]query_type=3时返回此列表
  --      fee_amrcvb            N  1 应收金额
  --      fee_ampaib            N  1 实收金额
  --      drug_amrcvb           N  1 药品应收金额
  --      drug_ampaib           N  1 药品实收金额
  --    fee_od_list[]query_type=4时有此列表
  --         order_id           N 1 医嘱id
  --         fee_no             C 1 费用单据号
  --         bill_prop          N 1 记录性质
  --         exe_state          N 1 费用记录执行状态
  --         rec_state          N 1 记录状态，费用记录的记录状态
  --         fee_state          N 1 费用状态，费用记录中
  --         exe_dept_id        N 1 执行部门id
  --         exe_dept_name      C 1 执行科室名称，费用记录执行部门id对应的名称
  --         fee_type           C 1 收费类别
  --         nums               N 1 数次
  --         fee_item_id        N 1 收费细目id
  --         unit               C 1 发送数次转换后的单位，药品就是 对应的包装单，普通的 收费项目目录 的 计算单位
  --         fee_name           C 1 收费项目的名称，按系统参获取了的【输入药品显示】
  --    fee_dept_list[]query_type=8时有此列表
  --       fee_id               N 1 费用id
  --       exe_dept_id          N 1 执行部门id
  --    fee_pivas_list[]静配的药品行医嘱的费用列表
  --       fee_id               N 1 费用id
  --       exe_dept_id          N 1 执行科室id
  --       drug_id              N 1 收费细目id
  --       quantity             N 1 剩余数量
  --       order_id             N 1 医嘱id,医嘱序号;
  --       fee_no               C 1 费用单据号
  --       fee_origin           N 1 费用源来，1-门诊费用表，2-住院费用表
  --       serial_num           N 1 序号
  ---------------------------------------------------------------------------------------------------------
  j_Input Pljson;
  j_Json  Pljson;

  n_病人id      Number;
  n_主页id      Number;
  n_查询方式    Number;
  n_药品显示    Number;
  v_Vals        Clob;
  l_Vals        t_Strlist;
  n_Baby        Number;
  v_Order_Ids   Varchar2(32767);
  n_Tmp         Number;
  n_Fee_Ampaib  Number;
  n_医嘱id      Number;
  n_来源        Number; --1-门诊，2-费用
  v_Fee_No      Varchar2(2000);
  n_记录性质    Number;
  v_Jtmp        Varchar2(32767);
  v_Fee_Ids     Varchar2(32767);
  v_Fee_Ids_Out Varchar2(32767);
  c_Jtmp        Clob;

  Cursor c_Feeoutone(P医嘱id Number) Is
    Select a.医嘱序号 As 医嘱id, a.记录性质, a.No, a.执行状态, a.记录状态, a.费用状态, a.结帐id, a.费用序号, a.执行部门id, c.名称 As 执行科室, a.收费类别,
           (a.数次 * a.付数) 剩余数量, (a.数次 * a.付数 / Nvl(a.门诊包装, 1)) As 发送数次, a.收费细目id,
           Decode(Nvl(Instr('567', a.收费类别), 0), 0, Decode(a.收费类别, '4', b.计算单位, b.计算单位), a.门诊单位) As 单位,
           Nvl(g.名称, b.名称) || Decode(b.产地, Null, Null, '(' || b.产地 || ')') || Decode(b.规格, Null, Null, ' ' || b.规格) As 收费项目
    From (Select a.医嘱序号, Min(a.记录性质) As 记录性质, a.No, a.执行状态, Min(a.记录状态) As 记录状态, Min(a.费用状态) As 费用状态, Min(a.结帐id) As 结帐id,
                  a.序号 As 费用序号, a.执行部门id, a.收费类别, a.数次, a.付数, a.收费细目id, b.门诊包装, b.门诊单位, b.剂量系数
           From 门诊费用记录 A, 药品规格 B
           Where a.记录状态 In (0, 1, 3) And a.价格父号 Is Null And a.收费细目id = b.药品id(+) And a.医嘱序号 = P医嘱id
           Group By a.医嘱序号, a.No, a.执行状态, a.序号, a.执行部门id, a.收费类别, a.数次, a.付数, a.收费细目id, b.门诊包装, b.门诊单位, b.剂量系数) A,
         收费项目目录 B, 收费项目别名 G, 部门表 C
    Where a.执行部门id = c.Id(+) And a.收费细目id = b.Id And a.收费细目id = g.收费细目id(+) And g.码类(+) = 1 And g.性质(+) = n_药品显示
    Order By a.费用序号;

  Type t_Fee Is Table Of c_Feeoutone%RowType;
  r_Fee t_Fee;

  Cursor c_Feeout(P医嘱ids Varchar2) Is
    Select a.医嘱序号 As 医嘱id, a.记录性质, a.No, a.执行状态, a.记录状态, a.费用状态, a.结帐id, a.费用序号, a.执行部门id, c.名称 As 执行科室, a.收费类别,
           (a.数次 * a.付数) 剩余数量, (a.数次 * a.付数 / Nvl(a.门诊包装, 1)) As 发送数次, a.收费细目id,
           Decode(Nvl(Instr('567', a.收费类别), 0), 0, Decode(a.收费类别, '4', b.计算单位, b.计算单位), a.门诊单位) As 单位,
           Nvl(g.名称, b.名称) || Decode(b.产地, Null, Null, '(' || b.产地 || ')') || Decode(b.规格, Null, Null, ' ' || b.规格) As 收费项目
    From (Select a.医嘱序号, Min(a.记录性质) As 记录性质, a.No, a.执行状态, Min(a.记录状态) As 记录状态, Min(a.费用状态) As 费用状态, Min(a.结帐id) As 结帐id,
                  a.序号 As 费用序号, a.执行部门id, a.收费类别, a.数次, a.付数, a.收费细目id, b.门诊包装, b.门诊单位, b.剂量系数
           From 门诊费用记录 A, 药品规格 B
           Where a.记录状态 In (0, 1, 3) And a.价格父号 Is Null And a.收费细目id = b.药品id(+) And
                 a.医嘱序号 In (Select /*+cardinality(x,10)*/
                             x.Column_Value
                            From Table(Cast(f_Num2list(P医嘱ids) As Zltools.t_Numlist)) X)
           Group By a.医嘱序号, a.No, a.执行状态, a.序号, a.执行部门id, a.收费类别, a.数次, a.付数, a.收费细目id, b.门诊包装, b.门诊单位, b.剂量系数) A,
         收费项目目录 B, 收费项目别名 G, 部门表 C
    Where a.执行部门id = c.Id(+) And a.收费细目id = b.Id And a.收费细目id = g.收费细目id(+) And g.码类(+) = 1 And g.性质(+) = n_药品显示
    Order By a.费用序号;

  Cursor c_Feeinone(P医嘱id Number) Is
    Select a.医嘱序号 As 医嘱id, a.记录性质, a.No, a.执行状态, a.记录状态, a.费用状态, a.结帐id, a.费用序号, a.执行部门id, c.名称 As 执行科室, a.收费类别,
           (a.数次 * a.付数) 剩余数量, (a.数次 * a.付数 / Nvl(a.住院包装, 1)) As 发送数次, a.收费细目id,
           Decode(Nvl(Instr('567', a.收费类别), 0), 0, Decode(a.收费类别, '4', b.计算单位, b.计算单位), a.住院单位) As 单位,
           Nvl(g.名称, b.名称) || Decode(b.产地, Null, Null, '(' || b.产地 || ')') || Decode(b.规格, Null, Null, ' ' || b.规格) As 收费项目
    From (Select a.医嘱序号, Min(a.记录性质) As 记录性质, a.No, a.执行状态, Min(a.记录状态) As 记录状态, Min(a.费用状态) As 费用状态, Min(a.结帐id) As 结帐id,
                  a.序号 As 费用序号, a.执行部门id, a.收费类别, a.数次, a.付数, a.收费细目id, b.住院包装, b.住院单位, b.剂量系数
           From 住院费用记录 A, 药品规格 B
           Where a.记录状态 In (0, 1, 3) And a.价格父号 Is Null And a.收费细目id = b.药品id(+) And a.医嘱序号 = P医嘱id
           Group By a.医嘱序号, a.No, a.执行状态, a.序号, a.执行部门id, a.收费类别, a.数次, a.付数, a.收费细目id, b.住院包装, b.住院单位, b.剂量系数) A,
         收费项目目录 B, 收费项目别名 G, 部门表 C
    Where a.执行部门id = c.Id(+) And a.收费细目id = b.Id And a.收费细目id = g.收费细目id(+) And g.码类(+) = 1 And g.性质(+) = n_药品显示
    Order By a.费用序号;

  Cursor c_Feein(P医嘱ids Varchar2) Is
    Select a.医嘱序号 As 医嘱id, a.记录性质, a.No, a.执行状态, a.记录状态, a.费用状态, a.结帐id, a.费用序号, a.执行部门id, c.名称 As 执行科室, a.收费类别,
           (a.数次 * a.付数) 剩余数量, (a.数次 * a.付数 / Nvl(a.住院包装, 1)) As 发送数次, a.收费细目id,
           Decode(Nvl(Instr('567', a.收费类别), 0), 0, Decode(a.收费类别, '4', b.计算单位, b.计算单位), a.住院单位) As 单位,
           Nvl(g.名称, b.名称) || Decode(b.产地, Null, Null, '(' || b.产地 || ')') || Decode(b.规格, Null, Null, ' ' || b.规格) As 收费项目
    From (Select a.医嘱序号, Min(a.记录性质) As 记录性质, a.No, a.执行状态, Min(a.记录状态) As 记录状态, Min(a.费用状态) As 费用状态, Min(a.结帐id) As 结帐id,
                  a.序号 As 费用序号, a.执行部门id, a.收费类别, a.数次, a.付数, a.收费细目id, b.住院包装, b.住院单位, b.剂量系数
           From 住院费用记录 A, 药品规格 B
           Where a.记录状态 In (0, 1, 3) And a.价格父号 Is Null And a.收费细目id = b.药品id(+) And
                 a.医嘱序号 In (Select /*+cardinality(x,10)*/
                             x.Column_Value
                            From Table(Cast(f_Num2list(P医嘱ids) As Zltools.t_Numlist)) X)
           Group By a.医嘱序号, a.No, a.执行状态, a.序号, a.执行部门id, a.收费类别, a.数次, a.付数, a.收费细目id, b.住院包装, b.住院单位, b.剂量系数) A,
         收费项目目录 B, 收费项目别名 G, 部门表 C
    Where a.执行部门id = c.Id(+) And a.收费细目id = b.Id And a.收费细目id = g.收费细目id(+) And g.码类(+) = 1 And g.性质(+) = n_药品显示
    Order By a.费用序号;

Begin
  --解析入参
  j_Input := Pljson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_查询方式  := j_Json.Get_Number('query_type');
  n_病人id    := j_Json.Get_Number('pati_id');
  n_主页id    := j_Json.Get_Number('pati_pageid');
  n_Baby      := j_Json.Get_Number('baby_num');
  n_医嘱id    := j_Json.Get_Number('order_id');
  n_来源      := j_Json.Get_Number('fee_origin');
  v_Fee_No    := j_Json.Get_String('fee_no');
  n_记录性质  := j_Json.Get_Number('bill_prop');
  v_Order_Ids := j_Json.Get_String('order_ids');

  If n_查询方式 = 1 Then
    v_Order_Ids := '0.0%';
    For R In (Select (100 * (a.所有费 - a.非药费) / Nvl(a.所有费, 1)) As 比例
              From (Select Sum(Decode(a.收费类别, '5', 0, '6', 0, '7', 0, a.实收金额)) As 非药费, Sum(a.实收金额) As 所有费
                     From 住院费用记录 A
                     Where a.病人id = n_病人id And a.主页id = n_主页id And a.记录状态 <> 0 Having Sum(a.实收金额) > 0) A) Loop
      If r.比例 > 0 Then
        v_Order_Ids := Round(r.比例, 1) || '%';
      End If;
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"成功","fee_ratio":"' || v_Order_Ids || '"}}';
  Elsif n_查询方式 = 2 Then
    For R In (Select a.医嘱序号 As 医嘱id
              From 住院费用记录 A
              Where a.病人id = n_病人id And a.主页id = n_主页id And a.记录状态 = 0 And (n_Baby Is Null Or Nvl(a.婴儿费, 0) = n_Baby) And
                    a.医嘱序号 Is Not Null
              Group By a.医嘱序号) Loop
      v_Order_Ids := v_Order_Ids || ',' || r.医嘱id;
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"成功","order_ids":"' || Substr(v_Order_Ids, 2) || '"}}';
  Elsif n_查询方式 = 3 Then
    For R In (Select Sum(a.应收金额) As 应收金额, Sum(a.实收金额) As 实收金额, Sum(Decode(Instr('567', a.收费类别), 0, 0, a.应收金额)) As 药品应收,
                     Sum(Decode(Instr('567', a.收费类别), 0, 0, a.实收金额)) As 药品实收
              From 门诊费用记录 A
              Where a.医嘱序号 In (Select /*+cardinality(x,10)*/
                                x.Column_Value
                               From Table(Cast(f_Num2list(v_Order_Ids) As Zltools.t_Numlist)) X)) Loop
    
      n_Tmp  := r.应收金额;
      v_Jtmp := v_Jtmp || '"fee_amrcvb":' || Zljsonstr(n_Tmp, 1);
    
      n_Tmp  := r.应收金额;
      v_Jtmp := v_Jtmp || ',"fee_ampaib":' || Zljsonstr(n_Tmp, 1);
    
      n_Tmp  := r.药品应收;
      v_Jtmp := v_Jtmp || ',"drug_amrcvb":' || Zljsonstr(n_Tmp, 1);
    
      n_Tmp  := r.药品实收;
      v_Jtmp := v_Jtmp || ',"drug_ampaib":' || Zljsonstr(n_Tmp, 1);
    
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"成功","fee_am_list":[{' || v_Jtmp || '}]}}';
  Elsif n_查询方式 = 4 Then
    n_药品显示 := zl_GetSysParameter('输入药品显示');
    n_药品显示 := Nvl(n_药品显示, 2);
    If Instr(v_Order_Ids, ',') = 0 Then
      n_医嘱id := v_Order_Ids;
    End If;
  
    If n_来源 = 1 Then
      If Nvl(n_医嘱id, 0) <> 0 Then
        Open c_Feeoutone(n_医嘱id);
        Fetch c_Feeoutone Bulk Collect
          Into r_Fee;
        Close c_Feeoutone;
      Else
        Open c_Feeout(v_Order_Ids);
        Fetch c_Feeout Bulk Collect
          Into r_Fee;
        Close c_Feeout;
      End If;
    Else
      If Nvl(n_医嘱id, 0) <> 0 Then
        Open c_Feeinone(n_医嘱id);
        Fetch c_Feeinone Bulk Collect
          Into r_Fee;
        Close c_Feeinone;
      Else
        Open c_Feein(v_Order_Ids);
        Fetch c_Feein Bulk Collect
          Into r_Fee;
        Close c_Feein;
      End If;
    End If;
    v_Jtmp := Null;
    For I In 1 .. r_Fee.Count Loop
      v_Jtmp := v_Jtmp || ',{"order_id":' || r_Fee(I).医嘱id;
      v_Jtmp := v_Jtmp || ',"fee_no":"' || r_Fee(I).No || '"';
      v_Jtmp := v_Jtmp || ',"bill_prop":' || Nvl(r_Fee(I).记录性质, 0);
      v_Jtmp := v_Jtmp || ',"exe_state":' || Nvl(r_Fee(I).执行状态, 0);
      v_Jtmp := v_Jtmp || ',"rec_state":' || Nvl(r_Fee(I).记录状态, 0);
      v_Jtmp := v_Jtmp || ',"fee_state":' || Nvl(r_Fee(I).费用状态, 0);
      v_Jtmp := v_Jtmp || ',"exe_dept_id":' || Nvl(r_Fee(I).执行部门id || '', 'null'); --执行部门id
      v_Jtmp := v_Jtmp || ',"exe_dept_name":"' || Zljsonstr(r_Fee(I).执行科室) || '"'; --C 1 执行科室名称，
      v_Jtmp := v_Jtmp || ',"fee_type":"' || r_Fee(I).收费类别 || '"'; --C 1 收费类别
      v_Jtmp := v_Jtmp || ',"nums":' || Zljsonstr(r_Fee(I).发送数次, 1); --N 1 数次，进行了换算的
      v_Jtmp := v_Jtmp || ',"quantity":' || Zljsonstr(r_Fee(I).剩余数量, 1); --N 1 剩余数量，费用中的剩余数理
      v_Jtmp := v_Jtmp || ',"fee_item_id":' || Nvl(r_Fee(I).收费细目id || '', 'null'); --N 1 收费细目id
      v_Jtmp := v_Jtmp || ',"unit":"' || Zljsonstr(r_Fee(I).单位) || '"';
      v_Jtmp := v_Jtmp || ',"fee_name":"' || Zljsonstr(r_Fee(I).收费项目) || '"';
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
      Json_Out := '{"output":{"code":1,"message":"成功","fee_od_list":[' || Substr(v_Jtmp, 2) || ']}}';
    Else
      c_Jtmp   := c_Jtmp || v_Jtmp;
      Json_Out := '{"output":{"code":1,"message":"成功","fee_od_list":[' || c_Jtmp || ']}}';
    End If;
  Elsif n_查询方式 = 5 Then
    If n_来源 = 1 Then
      Select Max(a.收费细目id)
      Into n_Tmp
      From 门诊费用记录 A
      Where a.No = v_Fee_No And a.记录性质 = n_记录性质 And a.医嘱序号 = n_医嘱id And a.记录状态 In (0, 1, 3);
    Else
      Select Max(a.收费细目id)
      Into n_Tmp
      From 住院费用记录 A
      Where a.No = v_Fee_No And a.记录性质 = n_记录性质 And a.医嘱序号 = n_医嘱id And a.记录状态 In (0, 1, 3);
    End If;
    Json_Out := '{"output":{"code":1,"message":"成功","fee_item_id":"' || Nvl(n_Tmp, 0) || '"}}';
  Elsif n_查询方式 = 6 Then
  
    v_Vals := v_Fee_No;
  
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
    n_Fee_Ampaib := 0;
    For I In 1 .. l_Vals.Count Loop
      If n_来源 = 1 Then
        Select Sum(a.实收金额) As 金额
        Into n_Tmp
        From 门诊费用记录 A,
             (Select /*+cardinality(f,10)*/
                f.C1 As NO, To_Number(f.C2) As 记录性质
               From Table(f_Str2list2(l_Vals(I), ',', ':')) F) N
        Where a.医嘱序号 Is Not Null And a.No = n.No And a.记录性质 = n.记录性质 And a.记帐费用 = 1 And a.记录状态 = 0;
      Else
        Select Sum(a.实收金额) As 金额
        Into n_Tmp
        From 住院费用记录 A,
             (Select /*+cardinality(f,10)*/
                f.C1 As NO, To_Number(f.C2) As 记录性质
               From Table(f_Str2list2(l_Vals(I), ',', ':')) F) N
        Where a.医嘱序号 Is Not Null And a.No = n.No And a.记录性质 = n.记录性质 And a.记帐费用 = 1 And a.记录状态 = 0;
      End If;
      n_Fee_Ampaib := n_Fee_Ampaib + Nvl(n_Tmp, 0);
    End Loop;
  
    Json_Out := '{"output":{"code":1,"message":"成功","fee_ampaib":' || Zljsonstr(n_Fee_Ampaib, 1) || '}}';
  Elsif n_查询方式 = 7 Then
    v_Fee_Ids := j_Json.Get_String('fee_ids');
    If n_来源 = 1 Then
      Select f_List2str(Cast(Collect(a.Id || '') As t_Strlist), ',') 费用ids
      Into v_Fee_Ids_Out
      From 门诊费用记录 A
      Where a.记录性质 In (2, 1, 11) And a.记录状态 = 1 And Nvl(a.费用状态, 0) = 0 And
            a.Id In (Select /*+cardinality(x,10)*/
                      x.Column_Value
                     From Table(Cast(f_Num2list(v_Fee_Ids) As Zltools.t_Numlist)) X);
    Else
      Select f_List2str(Cast(Collect(a.Id || '') As t_Strlist), ',') 费用ids
      Into v_Fee_Ids_Out
      From 住院费用记录 A
      Where a.记录性质 = 2 And a.记录状态 = 1 And Nvl(a.费用状态, 0) = 0 And
            a.Id In (Select /*+cardinality(x,10)*/
                      x.Column_Value
                     From Table(Cast(f_Num2list(v_Fee_Ids) As Zltools.t_Numlist)) X);
    End If;
    Json_Out := '{"output":{"code":1,"message":"成功","fee_ids":"' || v_Fee_Ids_Out || '"}}';
  Elsif n_查询方式 = 8 Then
    v_Jtmp := Null;
    --要变更执行的费用行，不包含药品和卫材
    If n_来源 = 1 Then
      For R In (Select ID, 执行部门id
                From 门诊费用记录
                Where 收费类别 Not In ('4', '5', '6', '7') And
                      医嘱序号 + 0 In (Select /*+cardinality(b,10) */
                                    Column_Value As 医嘱id
                                   From Table(Cast(f_Str2list(v_Order_Ids, ',') As t_Strlist)) B) And
                      NO In (Select /*+cardinality(b,10) */
                              Column_Value As NO
                             From Table(Cast(f_Str2list(v_Fee_No, ',') As t_Strlist)) B)) Loop
      
        v_Jtmp := v_Jtmp || ',{"fee_id":' || r.Id;
        v_Jtmp := v_Jtmp || ',"exe_dept_id":' || Nvl(r.执行部门id || '', 'null');
        v_Jtmp := v_Jtmp || '}';
      
      End Loop;
    Else
      For R In (Select ID, 执行部门id
                From 住院费用记录
                Where 收费类别 Not In ('4', '5', '6', '7') And
                      医嘱序号 + 0 In (Select /*+cardinality(b,10) */
                                    Column_Value As 医嘱id
                                   From Table(Cast(f_Str2list(v_Order_Ids, ',') As t_Strlist)) B) And
                      NO In (Select /*+cardinality(b,10) */
                              Column_Value As NO
                             From Table(Cast(f_Str2list(v_Fee_No, ',') As t_Strlist)) B)) Loop
        v_Jtmp := v_Jtmp || ',{"fee_id":' || r.Id;
        v_Jtmp := v_Jtmp || ',"exe_dept_id":' || Nvl(r.执行部门id || '', 'null');
        v_Jtmp := v_Jtmp || '}';
      End Loop;
    End If;
    Json_Out := '{"output":{"code":1,"message":"成功","fee_dept_list":[' || Substr(v_Jtmp, 2) || ']}}';
  
  Elsif n_查询方式 = 9 Then
    v_Jtmp := Null;
    --静配中的药品医嘱费用信息
    If n_来源 = 1 Then
      For R In (Select ID, 执行部门id, 收费细目id, (数次 * 付数) 剩余数量, 医嘱序号, NO, 序号
                From 门诊费用记录
                Where 收费类别 In ('5', '6') And 记录状态 In (0, 1, 3) And
                      医嘱序号 + 0 In (Select /*+cardinality(b,10) */
                                    Column_Value As 医嘱id
                                   From Table(Cast(f_Str2list(v_Order_Ids, ',') As t_Strlist)) B) And
                      NO In (Select /*+cardinality(b,10) */
                              Column_Value As NO
                             From Table(Cast(f_Str2list(v_Fee_No, ',') As t_Strlist)) B)
                Order By 序号 Desc) Loop
      
        v_Jtmp := v_Jtmp || ',{"fee_id":' || r.Id;
        v_Jtmp := v_Jtmp || ',"exe_dept_id":' || Nvl(r.执行部门id || '', 'null');
        v_Jtmp := v_Jtmp || ',"drug_id":' || Nvl(r.收费细目id || '', 'null');
        v_Jtmp := v_Jtmp || ',"quantity":' || Zljsonstr(r.剩余数量, 1);
        v_Jtmp := v_Jtmp || ',"order_id":' || r.医嘱序号;
        v_Jtmp := v_Jtmp || ',"fee_no":"' || r.No || '"';
        v_Jtmp := v_Jtmp || ',"fee_origin":1';
        v_Jtmp := v_Jtmp || ',"serial_num":' || r.序号;
        v_Jtmp := v_Jtmp || '}';
      
      End Loop;
    Else
      For R In (Select ID, 执行部门id, 收费细目id, (数次 * 付数) 剩余数量, 医嘱序号, NO, 序号
                From 住院费用记录
                Where 收费类别 In ('5', '6') And 记录状态 In (0, 1, 3) And
                      医嘱序号 + 0 In (Select /*+cardinality(b,10) */
                                    Column_Value As 医嘱id
                                   From Table(Cast(f_Str2list(v_Order_Ids, ',') As t_Strlist)) B) And
                      NO In (Select /*+cardinality(b,10) */
                              Column_Value As NO
                             From Table(Cast(f_Str2list(v_Fee_No, ',') As t_Strlist)) B)
                Order By 序号 Desc) Loop
        v_Jtmp := v_Jtmp || ',{"fee_id":' || r.Id;
        v_Jtmp := v_Jtmp || ',"exe_dept_id":' || Nvl(r.执行部门id || '', 'null');
        v_Jtmp := v_Jtmp || ',"drug_id":' || Nvl(r.收费细目id || '', 'null');
        v_Jtmp := v_Jtmp || ',"quantity":' || Zljsonstr(r.剩余数量, 1);
        v_Jtmp := v_Jtmp || ',"order_id":' || r.医嘱序号;
        v_Jtmp := v_Jtmp || ',"fee_no":"' || r.No || '"';
        v_Jtmp := v_Jtmp || ',"fee_origin":2';
        v_Jtmp := v_Jtmp || ',"serial_num":' || r.序号;
        v_Jtmp := v_Jtmp || '}';
      End Loop;
    End If;
    Json_Out := '{"output":{"code":1,"message":"成功","fee_pivas_list":[' || Substr(v_Jtmp, 2) || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getorderfeeinfo;
/

Create Or Replace Procedure Zl_Exsesvr_Getchargeofflist
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：获取未审核销帐的单
  --入参：Json_In:格式
  --  input
  --     pati_id                N 1 病人id
  --     pati_pageid            N 1 主页id 
  --出参: Json_Out,格式如下
  --  output
  --    code                    N 1 应答吗：0-失败；1-成功
  --    message                 C 1 应答消息：失败时返回具体的错误信息 
  --    charge_off_list[]列表
  --          fee_no            C 1 费用单据号
  --          item_name         C 1 收费项目名称
  --          dept_name         C 1 审核部门名称 
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_病人id Number;
  n_主页id Number;
  v_Output Varchar2(32767);

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_病人id := Nvl(j_Json.Get_Number('pati_id'), 0);
  n_主页id := Nvl(j_Json.Get_Number('pati_pageid'), 0);

  For R In (Select Distinct a.No, d.名称 As 项目, c.名称 As 部门
            From 住院费用记录 A, 病人费用销帐 B, 部门表 C, 收费项目目录 D
            Where a.Id = b.费用id And a.收费细目id = d.Id And b.审核部门id = c.Id(+) And b.审核时间 Is Null And a.病人id = n_病人id And
                  Nvl(a.主页id, 0) = n_主页id) Loop
  
    zlJsonPutValue(v_Output, 'fee_no', r.No, 0, 1);
    zlJsonPutValue(v_Output, 'item_name', r.项目);
    zlJsonPutValue(v_Output, 'dept_name', r.部门, 0, 2);
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"成功","charge_off_list":[' || v_Output || ']}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getchargeofflist;
/


Create Or Replace Procedure Zl_Exsesvr_Getbillsexestate
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：获取费用执行相关信息
  --入参：Json_In:格式
  --  input
  --     fee_nos               C 1 单据号逗号拼串，原则上来说是医嘱发送生成的费用no号不会有重复
  --     fee_origin            N 1 费用来源(默认=2：1-门诊费用，2-住院费用)
  --出参: Json_Out,格式如下
  --  output
  --    code                    N 1 应答吗：0-失败；1-成功
  --    message                 C 1 应答消息：失败时返回具体的错误信息
  --    fee_exe_list[]列表
  --          order_id          N 1 医嘱id
  --          fee_no            C 1 单据号
  --          fee_item_id       N 1 收费细目id
  --          exe_dept_id       N 1 执行部门id
  --          exe_state         N 1 执行状态
  ---------------------------------------------------------------------------
  j_Input Pljson;
  j_Json  Pljson;

  l_Nos  t_Strlist;
  v_Nos  Varchar2(32767);
  n_来源 Number; --1-门诊，2-费用
  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;

Begin
  --解析入参
  j_Input := Pljson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_来源 := j_Json.Get_Number('fee_origin');
  v_Nos  := j_Json.Get_String('fee_nos');
  l_Nos  := t_Strlist();

  While v_Nos Is Not Null Loop
    If Length(v_Nos) <= 4000 Then
      l_Nos.Extend;
      l_Nos(l_Nos.Count) := v_Nos;
      v_Nos := Null;
    Else
      l_Nos.Extend;
      l_Nos(l_Nos.Count) := Substr(v_Nos, 1, Instr(v_Nos, ',', 3980) - 1);
      v_Nos := Substr(v_Nos, Instr(v_Nos, ',', 3980) + 1);
    End If;
  End Loop;

  For I In 1 .. l_Nos.Count Loop
    If 1 = n_来源 Then
      --门诊
      For R In (Select a.医嘱序号 As 医嘱id, a.No, a.收费细目id, a.执行部门id, Max(a.执行状态) As 执行状态
                From 门诊费用记录 A
                Where a.No In (Select /*+cardinality(f,10)*/
                                f.Column_Value As NO
                               From Table(f_Str2list(l_Nos(I))) F) And a.医嘱序号 Is Not Null And a.记录状态 In (0, 1)
                Group By a.医嘱序号, a.No, a.收费细目id, a.执行部门id) Loop
      
        v_Jtmp := v_Jtmp || ',{"order_id":' || r.医嘱id;
        v_Jtmp := v_Jtmp || ',"fee_no":"' || r.No || '"';
        v_Jtmp := v_Jtmp || ',"fee_item_id":' || Nvl(r.收费细目id || '', 'null');
        v_Jtmp := v_Jtmp || ',"exe_dept_id":' || Nvl(r.执行部门id || '', 'null');
        v_Jtmp := v_Jtmp || ',"exe_state":' || Nvl(r.执行状态 || '', 'null');
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
      --住院
      For R In (Select a.医嘱序号 As 医嘱id, a.No, a.收费细目id, a.执行部门id, Max(a.执行状态) As 执行状态
                From 住院费用记录 A
                Where a.No In (Select /*+cardinality(f,10)*/
                                f.Column_Value As NO
                               From Table(f_Str2list(l_Nos(I))) F) And a.医嘱序号 Is Not Null And a.记录状态 In (0, 1)
                Group By a.医嘱序号, a.No, a.收费细目id, a.执行部门id
                Order By a.医嘱序号, a.No) Loop
      
        v_Jtmp := v_Jtmp || ',{"order_id":' || r.医嘱id;
        v_Jtmp := v_Jtmp || ',"fee_no":"' || r.No || '"';
        v_Jtmp := v_Jtmp || ',"fee_item_id":' || Nvl(r.收费细目id || '', 'null');
        v_Jtmp := v_Jtmp || ',"exe_dept_id":' || Nvl(r.执行部门id || '', 'null');
        v_Jtmp := v_Jtmp || ',"exe_state":' || Nvl(r.执行状态 || '', 'null');
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
  End Loop;

  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"成功","fee_exe_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"成功","fee_exe_list":[' || c_Jtmp || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Exsesvr_Getbillsexestate;
/

CREATE OR REPLACE Procedure Zl_Exsesvr_Getreturndruginfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：获取退药申请与审核信息，药疗收发查询退药信息时会用到
  --入参：Json_In:格式
  --  input
  --     begin_time            C 1 开始时间，
  --     end_time              C 1 结束时间
  --     request_dept_id       N 1 申请部门ID
  --     audit_dept_id         N 1 审核部门id

  --     type_query            N 0 查询方式，2-药疗收发查询时查退药明细【传出申请信息和退药的数量明细】，3-药疗收发查询时查退药汇总【返回汇总结果】，4-药疗收发查询发退汇总时用【退药的数量明细】
  --     effective_time        N 0 期效 0-长嘱，1-临嘱，2-长嘱和临嘱
  --     otherdept_id          N 0 对方部门id,领药部门id
  --     pati_ids              C 0 病人ID逗号拼串，如果不区分病人，传空
  --     rcp_no                C 0 处方单，单据号

  --出参: Json_Out,格式如下
  --  output
  --    code                    N 1 应答吗：0-失败；1-成功
  --    message                 C 1 应答消息：失败时返回具体的错误信息
  --    drug_list[]列表
  --          drug_id           N 1 药品id，即收费细目id
  --          request_num       N 1 申请数
  --          audit_num         N 1 审核数

  --    grp_list[]退药汇总 
  --          rcp_info                C 1 药品信息
  --          in_unit                 C 1 住院单位 
  --          drug_code               C 1 药品编码  
  --          quantity                N 1 应发数 
  --          back_number             N 1 退药数 
  --          reality_number          N 1 实发数 
  --          money                   N 1 金额
  --   detail_list[]退药申请明细
  --          rcpdtl_id               N 1 费用id，处方明细id
  --          quantity                N 1 数量,退药数量
  --          serial_num              N 1 序号
  --          order_id                N 1 医嘱id 
  --          charge_time             C 1 申请时间 
  --          rcp_no                  C 1 No 处方号
  --          charge_people           C 1 申请人 
  --          pati_id                 N 1 病人id
  --          pati_pageid             N 1 主页id

  --   quan_list[]数量金额
  --         drug_id                  N 1 药品id
  --         quantity                 N 1 数量,退药数量
  --         re_money                 N 1 金额 退药金额
  ---------------------------------------------------------------------------
  j_Input      Pljson;
  j_Json       Pljson;
  d_开始时间   Date;
  d_结束时间   Date;
  n_申请部门id 病人费用销帐.申请部门id%Type;
  n_审核部门id 病人费用销帐.审核部门id%Type;
  v_Output     Varchar2(32767);
  c_Output     Clob;
  n_Showtype   Number(3);
  n_Type       Number(1);
  v_Jtmp       Varchar2(32767); --不要随便使用此变量
  c_Jtmp       Clob; --不要随便使用此变量
  n_效期       Number(3);
  v_病人ids    Varchar2(4000);
  n_对方部门id Number(18);
  n_费用id     Number(18);
  v_退药信息   Varchar2(32767);
  v_No         Varchar2(30);

  Cursor c_Group_Type Is
    Select a.摘要 药品编码, a.摘要 药品信息, a.摘要 住院单位, a.数次 数量, a.数次 退药数, a.数次 实发数, a.数次 金额
    From 住院费用记录 A
    Where 0 = 1;
  r_Grp c_Group_Type%RowType;

  Procedure Get出参拼串汇总 As
  Begin
    v_Jtmp := v_Jtmp || ',';
  
    Zljsonputvalue(v_Jtmp, 'rcp_info', r_Grp.药品信息, 0, 1);
    Zljsonputvalue(v_Jtmp, 'drug_code', r_Grp.药品编码, 0);
    Zljsonputvalue(v_Jtmp, 'in_unit', r_Grp.住院单位, 0);
    Zljsonputvalue(v_Jtmp, 'money', r_Grp.金额, 1);
    Zljsonputvalue(v_Jtmp, 'quantity', r_Grp.数量, 1);
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
  j_Input      := Pljson(Json_In);
  j_Json       := j_Input.Get_Pljson('input');
  n_Type       := Nvl(j_Json.Get_Number('type_query'), 0);
  d_开始时间   := To_Date(j_Json.Get_String('begin_time'), 'YYYY-MM-DD hh24:mi:ss');
  d_结束时间   := To_Date(j_Json.Get_String('end_time'), 'YYYY-MM-DD hh24:mi:ss');
  n_申请部门id := Nvl(j_Json.Get_Number('request_dept_id'), 0);
  n_审核部门id := Nvl(j_Json.Get_Number('audit_dept_id'), 0);
  n_效期       := j_Json.Get_Number('effective_time');
  v_No         := j_Json.Get_String('rcp_no');

  If n_Type = 0 Then
    For R In (Select a.收费细目id As 药品id, Sum(a.数量 / Nvl(b.住院包装, 1)) As 申请数,
                     Sum(Decode(a.状态, 1, a.数量 / Nvl(b.住院包装, 1), 0)) As 审核数
              From 病人费用销帐 A, 药品规格 B
              Where a.收费细目id = b.药品id And a.申请时间 Between d_开始时间 And d_结束时间 And a.申请部门id = n_申请部门id And
                    a.审核部门id = n_审核部门id
              Group By a.收费细目id) Loop
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      
        v_Output := Null;
      End If;
    
      Zljsonputvalue(v_Output, 'drug_id', r.药品id, 1, 1);
      Zljsonputvalue(v_Output, 'request_num', r.申请数, 1);
      Zljsonputvalue(v_Output, 'audit_num', r.审核数, 1, 2);
    
    End Loop;
  
    If Not c_Output Is Null And Not v_Output Is Null Then
      c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
      v_Output := '';
    End If;
    If Not c_Output Is Null Then
      Json_Out := To_Clob('{"output":{"code":1,"message":"成功","drug_list":[') || c_Output || To_Clob(']}}');
    Else
      Json_Out := '{"output":{"code":1,"message":"成功","drug_list":[' || v_Output || ']}}';
    End If;
  Else
    If d_开始时间 Is Null Then
      Select Sysdate - 1, Sysdate Into d_开始时间, d_结束时间 From Dual;
    End If;
    If n_Type = 3 Then
      Select zl_GetSysParameter('药品名称显示', Null, 100) Into n_Showtype From Dual;
      If Nvl(n_Showtype, 0) = 0 Then
        n_Showtype := 1;
      End If;
      For R In (Select a.药品编码,
                       Nvl(b.名称, a.名称) || Decode(a.产地, Null, Null, '(' || a.产地 || ')') ||
                        Decode(a.规格, Null, Null, ' ' || a.规格) As 药品信息, a.住院单位, a.数量, 0 退药数, 0 实发数, a.金额
                From (Select b.药品id, c.编码 As 药品编码, c.名称, c.产地, c.规格, b.住院单位, Sum(a.数量 / Nvl(b.住院包装, 1)) As 数量,
                              Sum(a.金额) As 金额
                       From (Select a.费用id, a.数量, a.数量 * b.标准单价 金额, b.收费细目id 药品id
                              From 病人费用销帐 A, 住院费用记录 B
                              Where a.费用id = b.Id And a.申请时间 Between d_开始时间 And d_结束时间 And b.医嘱序号 Is Not Null And
                                    (b.No = v_No Or Nvl(v_No, 'NONE') = 'NONE') And a.审核部门id = n_审核部门id And Nvl(a.状态, 0) = 0 And
                                    b.收费类别 In ('5', '6', '7') And (b.领药部门id = n_对方部门id Or Nvl(n_对方部门id, 0) = 0) And
                                    a.申请部门id = n_申请部门id And
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
      If c_Jtmp Is Null Then
        Json_Out := '{"output":{"code":1,"message":"成功","grp_list":[' || Substr(v_Jtmp, 2) || ']}}';
      Else
        c_Jtmp   := c_Jtmp || v_Jtmp;
        Json_Out := '{"output":{"code":1,"message":"成功","grp_list":[' || c_Jtmp || ']}}';
      End If;
    Elsif n_Type = 2 Then
      --这一坨应该放到费用域的服务当中
      --退药部份 申请时间 为主索引   
      --出参明细带了 人员和时间，查明细的时候用到 2 时
      n_费用id := 1;
      For R In (Select a.费用id, a.病人id, a.主页id, a.No, a.医嘱序号 医嘱id, Sum(a.数量) 数量,
                       To_Char(a.申请时间, 'YYYY-MM-DD HH24:MI:SS') 申请时间, a.申请人
                From (Select a.费用id, a.数量, a.申请时间, a.申请人, b.医嘱序号, b.No, b.病人id, b.主页id
                       From 病人费用销帐 A, 住院费用记录 B
                       Where a.费用id = b.Id And a.申请时间 Between d_开始时间 And d_结束时间 And b.医嘱序号 Is Not Null And
                             a.审核部门id = n_审核部门id /* n_库房id*/
                             And Nvl(a.状态, 0) = 0 And b.收费类别 In ('5', '6', '7') And
                             (b.No = v_No Or Nvl(v_No, 'NONE') = 'NONE') And (b.领药部门id = n_对方部门id Or Nvl(n_对方部门id, 0) = 0) And
                             a.申请部门id = n_申请部门id /* n_病区id*/
                             And (b.病人id + 0 In (Select /*+cardinality(x,10)*/
                                                  x.Column_Value
                                                 From Table(f_Str2list(v_病人ids)) X) Or Nvl(v_病人ids, 'NONE') = 'NONE') And
                             (b.医嘱期效 = n_效期 Or Nvl(n_效期, 2) = 2)) A
                Group By a.费用id, a.病人id, a.主页id, a.No, a.医嘱序号, a.申请时间, a.申请人) Loop
      
        v_退药信息 := v_退药信息 || ',{"rcpdtl_id":' || r.费用id;
        v_退药信息 := v_退药信息 || ',"quantity":' || Zljsonstr(r.数量, 1);
        v_退药信息 := v_退药信息 || ',"serial_num":' || n_费用id;
        v_退药信息 := v_退药信息 || ',"order_id":' || r.医嘱id;
        v_退药信息 := v_退药信息 || ',"charge_time":"' || r.申请时间 || '"';
        v_退药信息 := v_退药信息 || ',"rcp_no":"' || r.No || '"';
        v_退药信息 := v_退药信息 || ',"charge_people":"' || Zljsonstr(r.申请人) || '"';
        v_退药信息 := v_退药信息 || ',"pati_pageid":' || r.主页id;
        v_退药信息 := v_退药信息 || ',"pati_id":' || r.病人id;
        v_退药信息 := v_退药信息 || '}';
        n_费用id   := n_费用id + 1;
      End Loop;
      Json_Out := '{"output":{"code":1,"message":"成功","detail_list":[' || Substr(v_退药信息, 2) || ']}}';
    Elsif n_Type = 4 Then
      --出参汇总明细，查 发退汇总的时候用到 4 时
      For R In (Select a.药品id, Sum(a.数量) 数量, Sum(a.数量 * a.标准单价) 金额
                From (Select a.数量, b.标准单价, b.收费细目id 药品id
                       From 病人费用销帐 A, 住院费用记录 B
                       Where a.费用id = b.Id And a.申请时间 Between d_开始时间 And d_结束时间 And b.医嘱序号 Is Not Null And
                             a.审核部门id = n_审核部门id And Nvl(a.状态, 0) = 0 And b.收费类别 In ('5', '6', '7') And
                             (b.领药部门id = n_对方部门id Or Nvl(n_对方部门id, 0) = 0) And a.申请部门id = n_申请部门id And
                             (b.No = v_No Or Nvl(v_No, 'NONE') = 'NONE') And
                             (b.病人id + 0 In (Select /*+cardinality(x,10)*/
                                              x.Column_Value
                                             From Table(f_Str2list(v_病人ids)) X) Or Nvl(v_病人ids, 'NONE') = 'NONE') And
                             (b.医嘱期效 = n_效期 Or Nvl(n_效期, 2) = 2)
                       -- 汇总发药号，转换成费用ID拼串再传到费用那边去，
                       -- 退药查询的时候让 给药途径过滤条件失效，在不冗余的情况下只有这个样子了
                       ) A
                Group By a.药品id, a.标准单价) Loop
        v_退药信息 := v_退药信息 || ',{"drug_id":' || r.药品id;
        v_退药信息 := v_退药信息 || ',"quantity":' || Zljsonstr(r.数量, 1);
        v_退药信息 := v_退药信息 || ',"re_money":' || Zljsonstr(r.金额, 1);
        v_退药信息 := v_退药信息 || '}';
      End Loop;
      Json_Out := '{"output":{"code":1,"message":"成功","quan_list":[' || Substr(v_退药信息, 2) || ']}}';
    End If;
  End If;

Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getreturndruginfo;
/

CREATE OR REPLACE Procedure Zl_Exsesvr_Adddepositinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:增加预交款数据
  --入参：Json_In:格式
  --  input
  --    oper_fun            C    操作状态:0-正常的预交款 ;1-保存为未生效的预交款
  --    deposit_info        C  1  预交款列表
  --      pati_id           N  1  病人id
  --      pati_pageid       N  1  主页id
  --      pati_name         C  1  病人姓名
  --      pati_sex          C  1  性别
  --      pati_age          C  1  年龄
  --      outpatient_num    C  1  门诊号
  --      inpatient_num     C  1  住院号
  --      mdlpay_name       C  1  付款方式名称
  --      deposit_id        N  1  预交ID
  --      deposit_no        C  1  预交单据号
  --      invc_no           C  1  发票号
  --      deposit_type      N     预交类别:1-门诊;2-住院
  --      dept_id           N  1  缴款科室id
  --      money             N  1  缴款金额
  --      emp_name          C  1  缴款单位
  --      emp_bank_name     C  1  单位开户行
  --      emp_bank_actno    C  1  开户行账号
  --      memo              C  1  摘要
  --      recv_id           N  1  票据领用id
  --    balance_info        C     结算信息:目前只支持一种结算方式
  --      blnc_mode         C  1  结算方式
  --      blnc_no           C  1  结算号码
  --      cardtype_id       C  1  卡类别id
  --      consumer_no       C  1  结算卡序号，即卡消费接口目录.编号
  --      consume_card_id   N  1  消费卡ID
  --      cardno            C  1  支付卡号
  --      swapno            C  1  交易流水号
  --      swapmemo          C  1  交易说明
  --      cprtion_unit      C  1  合作单位
  --      operator_name     C  1  操作员姓名
  --      operator_code     C  1  操作员编号
  --      create_time       C  1  登记时间或收款时间:yyyy-mm-dd hh:mi:ss
  --      insurance_type    N  1  险类
  --      insurance_num     C  1  医保号
  --      insurance_pwd     C  1  医保密码
  --      start_einv        N  1  是否启用电子票据:1-启用;0-不启用
  --出参: Json_Out,格式如下
  --  output
  --    code  C  1  应答码：0-失败；1-成功
  --    message  C  1  应答消息：失败时返回具体的错误信息
  --    deposit_id  N 1 预交ID
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;
  o_Json  PLJson;

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  --本次结算信息
  n_操作状态 Number(2);

  v_操作员姓名 门诊费用记录.操作员姓名%Type;
  v_操作员编号 门诊费用记录.操作员编号%Type;

  d_登记时间 门诊费用记录.登记时间%Type;
  n_消费卡id Number(18);

  --支付方式定义
  v_结算方式   病人预交记录.结算方式%Type;
  v_结算号码   病人预交记录.结算号码%Type;
  n_卡类别id   病人预交记录.卡类别id%Type;
  n_结算卡序号 病人预交记录.结算卡序号%Type;
  v_支付卡号   病人预交记录.卡号%Type;
  v_交易流水号 病人预交记录.交易流水号%Type;
  v_交易说明   病人预交记录.交易说明%Type;
  v_合作单位   病人预交记录.合作单位%Type;
  n_关联交易id 病人预交记录.关联交易id%Type;
  --预交相关变量定义
  n_病人id     病人预交记录.病人id%Type;
  n_预交id     病人预交记录.Id%Type;
  v_预交单号   病人预交记录.No%Type;
  n_预交类别   病人预交记录.预交类别%Type;
  n_主页id     病人预交记录.主页id%Type;
  n_缴款科室id 病人预交记录.科室id%Type;
  n_缴款金额   病人预交记录.金额%Type;
  v_缴款单位   病人预交记录.缴款单位%Type;
  v_单位开户行 病人预交记录.单位开户行%Type;
  v_开户行账号 病人预交记录.单位帐号%Type;
  v_摘要       病人预交记录.摘要%Type;
  n_领用id     票据使用明细.领用id%Type;
  v_发票号     病人预交记录.实际票号%Type;

  v_病人姓名     病人预交记录.姓名%Type;
  v_性别         病人预交记录.性别%Type;
  v_年龄         病人预交记录.年龄%Type;
  n_门诊号       病人预交记录.门诊号%Type;
  n_住院号       病人预交记录.住院号%Type;
  v_付款方式名称 病人预交记录.付款方式名称%Type;
  n_预交电子票据 病人预交记录.预交电子票据%Type;
  n_险类         Number(18);
  v_医保号       Varchar2(100);
  v_密码         Varchar2(100);
Begin
  --解析入参

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_操作状态   := Nvl(j_Json.Get_Number('oper_fun'), 0);
  v_操作员姓名 := j_Json.Get_String('operator_name');
  v_操作员编号 := j_Json.Get_String('operator_code');
  d_登记时间   := To_Date(j_Json.Get_String('create_time'), 'YYYY-MM-DD hh24:mi:ss');

  If d_登记时间 Is Null Then
    d_登记时间 := Sysdate;
  End If;

  --1处理结算信息
  o_Json := j_Json.Get_Pljson('balance_info');
  If o_Json Is Null Then
    v_Err_Msg := '不存在结算信息，请检查!';
    Raise Err_Item;
  End If;

  v_结算方式     := o_Json.Get_String('blnc_mode');
  v_结算号码     := o_Json.Get_String('blnc_no');
  n_卡类别id     := o_Json.Get_Number('cardtype_id');
  n_结算卡序号   := o_Json.Get_Number('consumer_no');
  v_支付卡号     := o_Json.Get_String('cardno');
  v_交易流水号   := o_Json.Get_String('swapno');
  v_交易说明     := o_Json.Get_String('swapmemo');
  v_合作单位     := o_Json.Get_String('cprtion_unit');
  n_消费卡id     := o_Json.Get_Number('consume_card_id');
  n_险类         := o_Json.Get_Number('insurance_type');
  v_医保号       := o_Json.Get_String('insurance_num');
  v_密码         := o_Json.Get_String('insurance_pwd');
  n_预交电子票据 := o_Json.Get_Number('start_einv');

  If Nvl(n_卡类别id, 0) = 0 Then
    n_卡类别id := Null;
  End If;
  If Nvl(n_结算卡序号, 0) = 0 Then
    n_结算卡序号 := Null;
  End If;

  --2.获取预交信息
  o_Json := PLJson();
  o_Json := j_Json.Get_Pljson('deposit_info');
  If o_Json Is Null Then
    v_Err_Msg := '不存在预交单据信息,请检查!';
    Raise Err_Item;
  End If;

  n_病人id     := o_Json.Get_Number('pati_id');
  n_主页id     := o_Json.Get_Number('pati_pageid');
  n_预交id     := o_Json.Get_Number('deposit_id');
  v_预交单号   := o_Json.Get_String('deposit_no');
  v_发票号     := o_Json.Get_String('invc_no');
  n_预交类别   := Nvl(o_Json.Get_Number('deposit_type'), 2);
  n_缴款科室id := o_Json.Get_Number('dept_id');
  n_缴款金额   := o_Json.Get_Number('money');
  v_缴款单位   := o_Json.Get_String('emp_name');
  v_单位开户行 := o_Json.Get_String('emp_bank_name');
  v_开户行账号 := o_Json.Get_String('emp_bank_actno');
  v_摘要       := o_Json.Get_String('memo');
  n_领用id     := o_Json.Get_Number('recv_id');

  v_病人姓名     := o_Json.Get_String('pati_name');
  v_性别         := o_Json.Get_String('pati_sex');
  v_年龄         := o_Json.Get_String('pati_age');
  n_门诊号       := To_Number(o_Json.Get_String('outpatient_num'));
  n_住院号       := To_Number(o_Json.Get_String('inpatient_num'));
  v_付款方式名称 := o_Json.Get_String('mdlpay_name');

  If Nvl(n_预交id, 0) = 0 Then
    Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
  End If;

  If Nvl(n_关联交易id, 0) = 0 Then
    n_关联交易id := n_预交id;
  End If;

  --操作状态_In:0-正常结算，1-保存为异常单据或未生效的单据，2-完成异常结算
  If Nvl(n_险类, 0) <> 0 Then
    v_缴款单位   := n_险类;
    v_开户行账号 := v_医保号;
    v_单位开户行 := v_密码;
  End If;

  If n_预交电子票据 Is Null Then
    n_预交电子票据 := Zl_Fun_Isstarteinvoice(2, Nvl(n_险类, 0), 1, n_预交类别);
  End If;

  Zl_病人预交记录_Insert_s(n_预交id, v_预交单号, v_发票号, n_病人id, n_主页id, v_病人姓名, v_性别, v_年龄, n_门诊号, n_住院号, v_付款方式名称, n_缴款科室id,
                     n_缴款金额, v_结算方式, v_结算号码, v_缴款单位, v_单位开户行, v_开户行账号, v_摘要, v_操作员编号, v_操作员姓名, n_领用id, n_预交类别, n_卡类别id,
                     n_结算卡序号, v_支付卡号, v_交易流水号, v_交易说明, v_合作单位, d_登记时间, Null, Null, 1, Nvl(n_操作状态, 0), n_关联交易id, Null,
                     Nvl(n_预交电子票据, 0), n_险类);

  If Nvl(n_结算卡序号, 0) <> 0 And Nvl(n_消费卡id, 0) <> 0 Then
    -- 消费卡处理
    Zl_病人卡结算记录_支付(n_消费卡id, v_支付卡号, 0, n_缴款金额, n_预交id, v_操作员编号, v_操作员姓名, d_登记时间);
  End If;

  Json_Out := zlJsonOut('成功', 1);

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, ' [ ZLSOFT ] ' || v_Err_Msg || ' [ ZLSOFT ] ');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Adddepositinfo;
/


Create Or Replace Procedure Zl_Exsesvr_Getorderchargedinfo
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：判断当前的执行医嘱对应的费用单是否已收费或记帐划价单是否已审核和单据的状态
  --入参：Json_In:格式
  --  input
  --     fee_origin         N 1 费用来源(默认=2：1-门诊费用，2-住院费用)
  --     order_ids          C 1 医嘱IDs，逗号分割
  --     fee_nos            C 1 费用单据拼串，逗号分割
  --     oper_type          N 1 判断方式 ：0-检查是否存在未收费记录，1-检查是否存在已收费记录

  --出参: Json_Out,格式如下
  --  output
  --    code                N 1 应答吗：0-失败；1-成功
  --    message             C 1 应答消息：失败时返回具体的错误信息
  --    isexist             N 1 返回值，真假：0-假，1-真
  --    blance_sign         N 1 是否有异常费用，0-正常，1-存在异常单据
  ---------------------------------------------------------------------------
  j_Input     PLJson;
  j_Json      PLJson;
  n_来源      Number;
  v_Order_Ids Varchar2(32767);
  v_Fee_Nos   Varchar2(32767);
  n_有异常费  Number;
  Int方式     Number;
  n_Cnt       Number;
  n_Blnout    Number;

  v_Output Varchar2(32767);
  Cursor c_Out Is
    Select Nvl(a.记录状态, 0) As 记录状态, a.医嘱序号 As 医嘱id, Nvl(a.执行状态, 0) As 执行状态, Nvl(a.结帐id, 0) As 结帐id, a.No,
           Nvl(a.费用状态, 0) As 费用状态, Nvl(a.记录性质, 0) As 记录性质
    From 门诊费用记录 A
    Where a.记录状态 In (0, 1, 3) And
          a.No In (Select /*+cardinality(x,10)*/
                    x.Column_Value
                   From Table(Cast(f_Str2List(v_Fee_Nos) As Zltools.t_Strlist)) X) And
          a.医嘱序号 + 0 In (Select /*+cardinality(x,10)*/
                          x.Column_Value
                         From Table(Cast(f_Num2List(v_Order_Ids) As Zltools.t_Numlist)) X);

  Type t_Fee Is Table Of c_Out%RowType;
  r_Fee t_Fee;

  Cursor c_In Is
    Select Nvl(a.记录状态, 0) As 记录状态, a.医嘱序号 As 医嘱id, Nvl(a.执行状态, 0) As 执行状态, Nvl(a.结帐id, 0) As 结帐id, a.No,
           Nvl(a.费用状态, 0) As 费用状态, Nvl(a.记录性质, 0) As 记录性质
    From 住院费用记录 A
    Where a.记录状态 In (0, 1, 3) And
          a.No In (Select /*+cardinality(x,10)*/
                    x.Column_Value
                   From Table(Cast(f_Str2List(v_Fee_Nos) As Zltools.t_Strlist)) X) And
          a.医嘱序号 + 0 In (Select /*+cardinality(x,10)*/
                          x.Column_Value
                         From Table(Cast(f_Num2List(v_Order_Ids) As Zltools.t_Numlist)) X);

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_来源      := j_Json.Get_Number('fee_origin');
  v_Order_Ids := j_Json.Get_String('order_ids');
  v_Fee_Nos   := j_Json.Get_String('fee_nos');
  Int方式     := j_Json.Get_Number('oper_type');

  If n_来源 = 1 Then
    Open c_Out;
    Fetch c_Out Bulk Collect
      Into r_Fee;
    Close c_Out;
  Else
    Open c_In;
    Fetch c_In Bulk Collect
      Into r_Fee;
    Close c_In;
  End If;

  n_Blnout   := 1;
  n_有异常费 := 0;
  n_Cnt      := r_Fee.Count;

  If Nvl(n_Cnt, 0) = 0 And Int方式 = 1 Then
    n_Blnout := 0;
  Else
    For I In 1 .. n_Cnt Loop
    
      --分支 int方式=0
      If Int方式 = 0 Then
        If r_Fee(I).记录性质 = 1 And r_Fee(I).记录状态 = 1 And r_Fee(I).费用状态 = 1 And r_Fee(I).结帐id <> 0 Then
          n_Blnout   := 0;
          n_有异常费 := 1;
          Exit;
        End If;
        If r_Fee(I).记录状态 = 0 Or r_Fee(I).记录性质 = 1 And r_Fee(I).记录状态 = 1 And r_Fee(I).结帐id = 0 Then
          n_Blnout := 0;
          Exit;
        End If;
      Else
        --分支 int方式=1
        If r_Fee(I).记录状态 <> 1 And r_Fee(I).费用状态 <> 1 Then
          n_Blnout := 0;
          Exit;
        End If;
      End If;
    
    End Loop;
  End If;
  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '成功');
  zlJsonPutValue(v_Output, 'isexist', n_Blnout, 1);
  zlJsonPutValue(v_Output, 'blance_sign', n_有异常费, 1, 2);
  Json_Out := '{"output":' || v_Output || '}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getorderchargedinfo;
/


Create Or Replace Procedure Zl_Exsesvr_Billhavebalance
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：判断一张记帐单/表是否已经结帐
  --入参：Json_In:格式
  --  input
  --       fee_origin         N 1 费用来源(默认=2：1-门诊费用，2-住院费用)
  --       fee_no             C 1 费用单据号，一个NO    
  --       order_id           N 1 医嘱id
  --出参: Json_Out,格式如下
  --  output
  --    code          N 1 应答吗：0-失败；1-成功
  --    message       C 1 应答消息：失败时返回具体的错误信息
  --    state         N 1 结帐情况，0-未结帐，1-已全部结帐，2-已部分结帐
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_Fee_No Varchar2(100);
  n_来源   Number;
  n_医嘱id Number;
  n_总行数 Number;
  n_结帐数 Number;
  n_结帐   Number;
Begin
  --解析入参 
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_Fee_No := j_Json.Get_String('fee_no');
  n_来源   := j_Json.Get_Number('fee_origin');
  n_医嘱id := j_Json.Get_Number('order_id');
  n_结帐   := 0;
  n_总行数 := 0;
  n_结帐数 := 0;
  If n_来源 = 1 Then
    For R In (Select Nvl(价格父号, 序号) As 序号, Sum(Nvl(结帐金额, 0)) As 结帐金额
              From 门诊费用记录
              Where NO = v_Fee_No And 记录性质 In (2, 12) And (医嘱序号 + 0 = n_医嘱id Or Nvl(n_医嘱id, 0) = 0)
              Group By Nvl(价格父号, 序号)) Loop
      n_总行数 := n_总行数 + 1;
      If Nvl(r.结帐金额, 0) <> 0 Then
        n_结帐数 := n_结帐数 + 1;
      End If;
    End Loop;
  Else
    For R In (Select Nvl(价格父号, 序号) As 序号, Sum(Nvl(结帐金额, 0)) As 结帐金额
              From 住院费用记录
              Where NO = v_Fee_No And 记录性质 In (2, 12) And (医嘱序号 + 0 = n_医嘱id Or Nvl(n_医嘱id, 0) = 0)
              Group By Nvl(价格父号, 序号)) Loop
      n_总行数 := n_总行数 + 1;
      If Nvl(r.结帐金额, 0) <> 0 Then
        n_结帐数 := n_结帐数 + 1;
      End If;
    End Loop;
  End If;
  --无结帐行,相当于未结帐
  If n_结帐数 = 0 Then
    n_结帐 := 0;
  Else
    If n_总行数 = n_结帐数 Then
      n_结帐 := 1;
    Else
      n_结帐 := 2;
    End If;
  End If;

  Json_Out := '{"output":{"code":1,"message":"成功","state":' || n_结帐 || '}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Billhavebalance;
/

CREATE OR REPLACE Procedure Zl_Exsesvr_Getdepositinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:获取预交信息
  --入参：Json_In:格式
  --   input
  --    deposit_no  C 1 单据号:预交单据号
  --    rec_state N 1 记录状态:1-原始充值的预交记录(不包含冲销记录);2-退款的预交记录,3-原始充值记录(包含销帐的及异常的)
  --出参  json
  --output
  --  code  C 1 应答码：0-失败；1-成功
  --  message C 1 应答消息：失败时返回具体的错误信息
  --  deposit_info  C 1 预交款信息
  --    pati_id N 1 病人ID
  --    pati_pageid N 1 主页id
  --    deposit_id  N 1 预交ID
  --    deposit_no  C 1 预交单据号
  --    invc_no C 1 发票号
  --    deposit_type  N   预交类别:1-门诊;2-住院
  --    dept_id N 1 缴款科室id
  --    money N 1 缴款金额
  --    emp_name  C 1 缴款单位
  --    emp_bank_name C 1 单位开户行
  --    emp_bank_actno  C 1 开户行账号
  --    memo  C 1 摘要
  --    operator_name C 1 操作员姓名
  --    operator_code C 1 操作员编号
  --    create_time C 1 收款时间:yyyy-mm-dd hh:mi:ss
  --  balance_info  C   结算信息:目前只支持一种结算方式
  --    blnc_mode C 1 结算方式
  --    blnc_no C 1 结算号码
  --    cardtype_id N 1 卡类别id
  --    consumer_no N 1 结算卡序号，即卡消费接口目录.编号
  --    consume_card_id N 1 消费卡ID
  --    cardno  C 1 支付卡号
  --    swapno  C 1 交易流水号
  --    swapmemo  C 1 交易说明
  --    cprtion_unit  C 1 合作单位
  --    blnc_state  N 1 结算状态(即校对标志):0或NULL正常缴款记录;1-未调用接口;2-接口调用完成
  --    insurance_type  N   险类
  --    insurance_num C   医保号
  --    insurance_pwd C   医保密码
  --    relation_id C 1 关联交易ID

  ---------------------------------------------------------------------------
  v_Err_Msg Varchar2(500);
  j_Input   PLJson;
  j_Json    PLJson;

  v_单据号   病人预交记录.No%Type;
  n_记录状态 Number(4);
  n_Find     Number(2);

  v_Output  Varchar2(32767);
  v_Balance Varchar2(32767);

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_单据号   := j_Json.Get_String('deposit_no');
  n_记录状态 := Nvl(j_Json.Get_Number('rec_state'), 1);

  If v_单据号 Is Null Then
    v_Err_Msg := '未传入预交单据号，请检查!';
    Json_Out  := zlJsonOut(v_Err_Msg);
    Return;
  End If;

  n_Find := 0;
  For c_预交 In (Select a.Id, a.No, a.实际票号, a.记录状态, a.病人id, a.主页id, a.科室id, a.缴款单位, a.单位开户行, a.单位帐号, a.摘要, a.金额, a.结算方式,
                      a.结算号码, To_Char(a.收款时间, 'yyyy-mm-dd hh24:mi:ss') As 收款时间, a.操作员编号, a.操作员姓名, a.预交类别, a.卡类别id,
                      a.结算卡序号, a.卡号, a.交易流水号, a.交易说明, a.合作单位, a.结算序号, a.校对标志, a.待转出, a.结算性质, a.会话号, a.关联交易id, a.交易时间,
                      a.交易人员, c.消费卡id, m.性质
               From 病人预交记录 A, 病人卡结算记录 C, 结算方式 M
               Where a.结算方式 = m.名称(+) And a.记录性质 = 1 And a.Id = c.结算id(+) And
                     (a.记录状态 = n_记录状态 Or n_记录状态 = 3 And a.记录状态 In (0, 1, 3)) And a.no= v_单据号) Loop
    --预交单据信息
    zlJsonPutValue(v_Output, 'pati_id', Nvl(c_预交.病人id, 0), 1, 1);
    zlJsonPutValue(v_Output, 'pati_pageid', Nvl(c_预交.主页id, 0), 1);
    zlJsonPutValue(v_Output, 'deposit_id', Nvl(c_预交.Id, 0), 1);
    zlJsonPutValue(v_Output, 'deposit_no', Nvl(c_预交.No, ''));
    zlJsonPutValue(v_Output, 'invc_no', Nvl(c_预交.实际票号, ''));
    zlJsonPutValue(v_Output, 'deposit_type', Nvl(c_预交.预交类别, ''));
    zlJsonPutValue(v_Output, 'dept_id', Nvl(c_预交.科室id, 0), 1);
    zlJsonPutValue(v_Output, 'money', Nvl(c_预交.金额, 0), 1);
    zlJsonPutValue(v_Output, 'emp_name', Nvl(c_预交.缴款单位, ''));
    zlJsonPutValue(v_Output, 'emp_bank_name', Nvl(c_预交.单位开户行, ''));
    zlJsonPutValue(v_Output, 'emp_bank_actno', Nvl(c_预交.单位帐号, ''));
    zlJsonPutValue(v_Output, 'memo', Nvl(c_预交.摘要, ''));
    zlJsonPutValue(v_Output, 'operator_code', Nvl(c_预交.操作员编号, ''));
    zlJsonPutValue(v_Output, 'operator_name', Nvl(c_预交.操作员姓名, ''));
    zlJsonPutValue(v_Output, 'create_time', Nvl(c_预交.收款时间, ''), 0, 2);
  
    v_Output := '"deposit_info": ' || v_Output;
  
    --结算信息
    zlJsonPutValue(v_Balance, 'blnc_mode', Nvl(c_预交.结算方式, ''), 0, 1);
    zlJsonPutValue(v_Balance, 'blnc_no', Nvl(c_预交.结算号码, ''));
  
    zlJsonPutValue(v_Balance, 'cardtype_id', Nvl(c_预交.卡类别id, 0), 1);
    zlJsonPutValue(v_Balance, 'consumer_no', Nvl(c_预交.结算卡序号, 0), 1);
    zlJsonPutValue(v_Balance, 'consume_card_id', Nvl(c_预交.消费卡id, 0), 1);
  
    zlJsonPutValue(v_Balance, 'cardno', Nvl(c_预交.卡号, ''));
    zlJsonPutValue(v_Balance, 'swapno', Nvl(c_预交.交易流水号, ''));
    zlJsonPutValue(v_Balance, 'swapmemo', Nvl(c_预交.交易说明, ''));
  
    zlJsonPutValue(v_Balance, 'cprtion_unit', Nvl(c_预交.合作单位, ''));
    zlJsonPutValue(v_Balance, 'relation_id', Nvl(c_预交.关联交易id, 0), 1);
  
    If Nvl(c_预交.性质, 0) = 3 Then
      zlJsonPutValue(v_Balance, 'insurance_type', To_Number(Nvl(c_预交.缴款单位, '0')));
      zlJsonPutValue(v_Balance, 'insurance_num', Nvl(c_预交.单位帐号, ''));
      zlJsonPutValue(v_Balance, 'insurance_pwd', Nvl(c_预交.单位开户行, ''));
	Else
      zlJsonPutValue(v_Balance, 'insurance_type', '0');
      zlJsonPutValue(v_Balance, 'insurance_num', '');
      zlJsonPutValue(v_Balance, 'insurance_pwd',  '');
    End If;
    zlJsonPutValue(v_Balance, 'blnc_state', Nvl(c_预交.校对标志, 0), 1, 2);
    v_Balance := '"balance_info":' || v_Balance || '';
    v_Output  := v_Output || ',' || v_Balance;
  
    n_Find := 1;
    Exit;
  End Loop;
  If n_Find = 0 Then
    v_Err_Msg := '未找到预交单据为' || v_单据号 || '的预交数据，请检查!';
    Json_Out  := zlJsonOut(v_Err_Msg);
    Return;
  End If;

  Json_Out := '{"output":{"code":1,"message":"成功",' || v_Output || '}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getdepositinfo;
/

Create Or Replace Procedure Zl_Exsesvr_Getoutproomlist
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  -------------------------------------------------------------------------------------
  --功能：根据条件获取有挂号安排的门诊诊室列表
  --入参：Json_In:格式
  --input
  --  query_type        N    1  查询方式 1-门诊医生站接诊诊室列表,2-病人转诊根据转诊科室加载接诊诊室,3-病人转诊根据临床出诊记录加载接诊诊室
  --  site_no           C    0 站点号
  --  outproom_name     C    0 诊室名称
  --  emg_sign          C    0 急诊标志
  --  dept_id           N    0 科室id
  --  outp_dr_name      C    0 门诊医生姓名
  --出参      json
  --output
  --    code                N   1 应答吗：0-失败；1-成功
  --    message             C   1 应答消息：失败时返回具体的错误信息
  --    outproom_list           门诊科室列表
  --       outproom_code    C   1 诊室编码
  --       outproom_name    C   1 诊室名称
  --       outproom_becode  C   1 诊室名称简码
  -------------------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_站点号   Varchar2(50);
  v_诊室名称 Varchar2(50);
  n_急诊标志 Number(5);

  n_查询方式 Number(5);

  n_科室id   Number(18);
  v_医生姓名 Varchar2(200);

  v_输入匹配 Varchar2(50);
  v_Output   Varchar2(32767);
  c_Output   Clob;
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_查询方式 := Nvl(j_Json.Get_Number('query_type'), 0);
  v_站点号   := j_Json.Get_String('site_no');
  v_诊室名称 := j_Json.Get_String('outproom_name');
  n_急诊标志 := Nvl(j_Json.Get_Number('outproom_name'), 0);
  v_医生姓名 := j_Json.Get_String('outp_dr_name');
  n_科室id   := Nvl(j_Json.Get_Number('dept_id'), 0);

  If n_查询方式 = 1 Then
  
    If Nvl(To_Number(zl_GetSysParameter('输入匹配')), 0) = 0 Then
      v_输入匹配 := '%';
    Else
      v_输入匹配 := '';
    End If;
  
    For c_门诊科室 In (Select Distinct e.编码, e.名称, e.简码
                   From 门诊诊室 E, 挂号安排诊室 D, 挂号安排 C, 部门人员 A, 上机人员表 B, 临床部门 F
                   Where a.人员id = b.人员id And b.用户名 = User And c.科室id = a.部门id And c.Id = d.号表id And e.名称 = d.门诊诊室 And
                         a.部门id = f.部门id And ((n_急诊标志 = 1 And f.工作性质 = '20') Or n_急诊标志 = 0) And
                         ((Upper(e.编码) Like v_诊室名称 || '%' Or Upper(e.简码) Like v_输入匹配 || v_诊室名称 || '%' Or
                         Upper(e.名称) Like v_输入匹配 || v_诊室名称 || '%') Or v_诊室名称 Is Null) And (e.站点 = v_站点号 Or e.站点 Is Null)) Loop
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      
        v_Output := Null;
      End If;
    
      zlJsonPutValue(v_Output, 'outproom_code', c_门诊科室.编码, 0, 1);
      zlJsonPutValue(v_Output, 'outproom_name', c_门诊科室.名称);
      zlJsonPutValue(v_Output, 'outproom_becode', c_门诊科室.简码, 0, 2);
    
    End Loop;
  Elsif n_查询方式 = 2 Then
    For c_门诊科室 In (Select Distinct e.编码, e.名称, e.简码
                   From 挂号安排诊室 A, 门诊诊室 E
                   Where a.号表id In (Select ID
                                    From 挂号安排
                                    Where 科室id = n_科室id And (医生姓名 = v_医生姓名 Or 医生姓名 Is Null Or v_医生姓名 Is Null)) And
                         a.门诊诊室 = e.名称
                   Order By 编码) Loop
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      
        v_Output := Null;
      End If;
    
      zlJsonPutValue(v_Output, 'outproom_code', c_门诊科室.编码, 0, 1);
      zlJsonPutValue(v_Output, 'outproom_name', c_门诊科室.名称);
      zlJsonPutValue(v_Output, 'outproom_becode', c_门诊科室.简码, 0, 2);
    
    End Loop;
  Elsif n_查询方式 = 3 Then
  
    For c_门诊科室 In (Select Distinct b.编码, b.名称, b.简码
                   From 临床出诊诊室记录 A, 门诊诊室 B
                   Where a.诊室id = b.Id And
                         a.记录id In
                         (Select a.Id
                          From 临床出诊记录 A
                          Where a.科室id = n_科室id And (a.医生姓名 = v_医生姓名 Or a.医生姓名 Is Null Or v_医生姓名 Is Null))
                   Order By b.名称) Loop
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      
        v_Output := Null;
      End If;
    
      zlJsonPutValue(v_Output, 'outproom_code', c_门诊科室.编码, 0, 1);
      zlJsonPutValue(v_Output, 'outproom_name', c_门诊科室.名称);
      zlJsonPutValue(v_Output, 'outproom_becode', c_门诊科室.简码, 0, 2);
    
    End Loop;
  End If;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"成功","outproom_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"成功","outproom_list":[' || v_Output || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getoutproomlist;
/

Create Or Replace Procedure Zl_Exsesvr_Getpatiantifee
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --功能：获取病人的抗菌药物费用汇总
  --入参：Json_In:格式
  --input
  --  pati_id        N    1  病人ID
  --  pati_pageid    N    1  主页ID
  --出参      json
  --output
  --    code                N   1 应答吗：0-失败；1-成功
  --    message             C   1 应答消息：失败时返回具体的错误信息
  --    fee_info            费用信息
  --       anti_fee            N   1 抗菌药费
  --       drug_fee            N   1 总药费
  --       inp_fee             N   1 住院费用

  j_Input  PLJson;
  j_Json   PLJson;
  v_Output Varchar2(32767);

  n_病人id Number(18);
  n_主页id Number(18);

Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_病人id := Nvl(j_Json.Get_Number('pati_id'), 0);
  n_主页id := Nvl(j_Json.Get_Number('pati_pageid'), 0);
  v_Output := Null;
  For c_费用列表 In (Select Sum(Decode(Nvl(e.抗生素, 0), 0, 0, a.结帐金额)) As 抗菌药费,
                        Sum(Decode(a.收费类别, '5', a.结帐金额, '6', a.结帐金额, '7', a.结帐金额, 0)) As 总药费, Sum(a.结帐金额) As 住院费用
                 From 住院费用记录 A, 药品规格 D, 药品特性 E
                 Where a.病人id = n_病人id And a.主页id = n_主页id And a.记录状态 <> 0 And a.收费细目id = d.药品id(+) And
                       d.药名id = e.药名id(+)) Loop
  
    zlJsonPutValue(v_Output, 'anti_fee', c_费用列表.抗菌药费, 1, 1);
    zlJsonPutValue(v_Output, 'drug_fee', c_费用列表.总药费, 1);
    zlJsonPutValue(v_Output, 'inp_fee', c_费用列表.住院费用, 1, 2);
  
  End Loop;

  If v_Output Is Null Then
    zlJsonPutValue(v_Output, 'anti_fee', 0, 1, 1);
    zlJsonPutValue(v_Output, 'drug_fee', 0, 1);
    zlJsonPutValue(v_Output, 'inp_fee', 0, 1, 2);
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功","fee_info":' || v_Output || '}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getpatiantifee;
/


Create Or Replace Procedure Zl_Exsesvr_Upddepositblncinfo
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:修改预交结算信息
  --入参：Json_In:格式
  --  input
  --    oper_state  N 1 操作状态:0-完成结算;1-接口调用前修正;2-接口调用后修正
  --    pati_id N 1 病人id
  --    deposit_no  C   预交单号
  --    deposit_id  N   预交ID
  --    operator_name C 1 操作员姓名
  --    operator_code C 1 操作员编号
  --    create_time C 1 操作时间:yyyy-mm-dd hh:mi:ss
  --    invc_no C 1 发票号
  --    recv_id N 1 领用id:领用Id
  --    balance_info  C   结算信息
  --      blnc_mode C 1 结算方式
  --      blnc_no C 1 结算号码
  --      cardtype_id N 1 卡类别id
  --      consumer_no N 1 结算卡序号，即卡消费接口目录.编号
  --      cardno  C 1 卡号
  --      swapno  C 1 交易流水号
  --      swapmemo  C 1 交易说明
  --      memo  C 1 摘要
  --      cprtion_unit  C 1 合作单位
  --      other_list[]  C 1 其他交易信息
  --        swap_name C 1 交易名称
  --        swap_note C 1 交易内容
  --出参: Json_Out,格式如下
  -- output
  --   code                  C 1 应答码：0-失败；1-成功
  --   message               C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Input    PLJson;
  j_Json     PLJson;
  o_Json     PLJson;
  j_Jsonlist Pljson_List := Pljson_List();

  n_病人id       Number(18);
  v_预交单号     病人预交记录.No%Type;
  n_预交id       病人预交记录.Id%Type;
  n_领用id       票据领用记录.Id%Type;
  v_结算方式     病人预交记录.结算方式%Type;
  v_结算号码     病人预交记录.结算号码%Type;
  n_卡类别id     病人预交记录.卡类别id%Type;
  n_结算卡序号   病人预交记录.结算卡序号%Type;
  v_支付卡号     病人预交记录.卡号%Type;
  v_交易流水号   病人预交记录.交易流水号%Type;
  v_交易说明     病人预交记录.交易说明%Type;
  v_摘要         病人预交记录.摘要%Type;
  v_合作单位     病人预交记录.合作单位%Type;
  v_交易名称     三方结算交易.交易项目%Type;
  v_交易内容     三方结算交易.交易内容%Type;
  v_发票号       病人预交记录.实际票号%Type;
  v_操作员姓名   病人预交记录.操作员姓名%Type;
  v_操作员编号   病人预交记录.操作员编号%Type;
  d_登记时间     病人预交记录.收款时间%Type;
  n_操作状态     Number(5);
  n_预交操作状态 Number(5);
  n_校对标志     Number(5);

  Err_Item Exception;
  v_Err_Msg Varchar2(500);
  n_Find    Number(2);

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_操作状态 := Nvl(j_Json.Get_Number('oper_state'), 0);
  n_病人id   := Nvl(j_Json.Get_Number('pati_id'), 0);
  v_预交单号 := j_Json.Get_String('deposit_no');
  n_预交id   := j_Json.Get_Number('deposit_id');

  v_操作员姓名 := j_Json.Get_String('operator_name');
  v_操作员编号 := j_Json.Get_String('operator_code');
  d_登记时间   := To_Date(j_Json.Get_String('create_time'), 'YYYY-MM-DD hh24:mi:ss');
  v_发票号     := j_Json.Get_String('invc_no');
  n_领用id     := j_Json.Get_Number('recv_id');

  If d_登记时间 Is Null Then
    d_登记时间 := Sysdate;
  End If;

  If Nvl(n_病人id, 0) = 0 Then
    v_Err_Msg := '不能确定病人信息，请检查！';
    Raise Err_Item;
  End If;
  If v_预交单号 Is Null Then
    v_Err_Msg := '不能确定预交单据信息，请检查！';
    Raise Err_Item;
  End If;

  o_Json := j_Json.Get_Pljson('balance_info');
  If o_Json Is Null Then
    v_Err_Msg := '不能确定预交单据为' || v_预交单号 || '的结算信息，请检查！';
    Raise Err_Item;
  End If;

  v_结算方式   := o_Json.Get_String('blnc_mode');
  v_结算号码   := o_Json.Get_String('blnc_no');
  n_卡类别id   := o_Json.Get_Number('cardtype_id');
  n_结算卡序号 := o_Json.Get_Number('consumer_no');
  v_支付卡号   := o_Json.Get_String('cardno');
  v_交易流水号 := o_Json.Get_String('swapno');
  v_交易说明   := o_Json.Get_String('swapmemo');
  v_摘要       := o_Json.Get_String('memo');
  v_合作单位   := o_Json.Get_String('cprtion_unit');
  n_Find       := 0;
  If Nvl(n_结算卡序号, 0) = 0 Then
    n_结算卡序号 := Null;
  
  End If;
  If Nvl(n_卡类别id, 0) = 0 Then
    n_卡类别id := Null;
  
  End If;
  For c_预交 In (Select ID, 记录性质, NO, 实际票号, 记录状态, 病人id, 主页id, 姓名, 性别, 年龄, 门诊号, 住院号, 付款方式名称, 科室id, 缴款单位, 单位开户行, 单位帐号, 摘要, 金额,
                      结算方式, 结算号码, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款, 找补, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
                      结算序号, 校对标志, 结算性质, 会话号, 附加标志, 关联交易id, 交易时间, 交易人员, 预交电子票据
               From 病人预交记录
               Where ID = n_预交id And 病人id = Nvl(n_病人id, 0)) Loop
    --操作状态_In:0-正常结算，1-保存为异常单据或未生效的单据，2-完成异常结算;3-修正数据
    --校对标志
  
    If n_操作状态 = 0 Then
      --操作状态:0-完成结算;1-接口调用前修正;2-接口调用后修正
      n_预交操作状态 := 2;
      n_校对标志     := Null;
    Elsif n_操作状态 = 1 Or n_操作状态 = 2 Then
      n_预交操作状态 := 3;
      n_校对标志     := n_操作状态;
    Else
      v_Err_Msg := '不能认别的操作功能，请检查！';
      Raise Err_Item;
    End If;
  
    Zl_病人预交记录_Insert_s(n_预交id, v_预交单号, v_发票号, n_病人id, c_预交.主页id, c_预交.姓名, c_预交.性别, c_预交.年龄, c_预交.门诊号, c_预交.住院号,
                       c_预交.付款方式名称, c_预交.科室id, c_预交.金额, v_结算方式, v_结算号码, c_预交.缴款单位, c_预交.单位开户行, c_预交.单位帐号, v_摘要, v_操作员编号,
                       v_操作员姓名, n_领用id, c_预交.预交类别, n_卡类别id, n_结算卡序号, v_支付卡号, v_交易流水号, v_交易说明, v_合作单位, d_登记时间, Null, Null,
                       1, n_预交操作状态, c_预交.关联交易id, n_校对标志, Nvl(c_预交.预交电子票据, 0));
  
    n_Find := 1;
  End Loop;
  If n_Find = 0 Then
    v_Err_Msg := '未找到单据号为' || v_预交单号 || '的预交单据信息，请检查！';
    Raise Err_Item;
  End If;
  j_Jsonlist := o_Json.Get_Pljson_List('other_list');
  If Not j_Jsonlist Is Null Then
  
    --先删除，后增加
    Delete 三方结算交易 Where 交易id = n_预交id;
    Delete 三方结算交易 Where 交易id = n_预交id;
  
    For J In 1 .. j_Jsonlist.Count Loop
      o_Json     := PLJson();
      o_Json     := PLJson(j_Jsonlist.Get(J));
      v_交易名称 := o_Json.Get_String('swap_name');
      v_交易内容 := o_Json.Get_String('swap_note');
      Insert Into 三方结算交易 (交易id, 交易项目, 交易内容) Values (n_预交id, v_交易名称, v_交易内容);
    End Loop;
  End If;
  Json_Out := zlJsonOut('成功', 1);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Upddepositblncinfo;
/
 
Create Or Replace Procedure Zl_Exsesvr_Deldepositerrorrec
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:删除异常的预交信息
  --入参：Json_In:格式
  --input
  --     deposit_no  C  1  预交单号
  --     oper_state  N  1  操作状态：0-删除异常充值单据，1-删除异常退款单据，2-删除异常余额退款单据 
  --出参: Json_Out,格式如下
  -- output
  --   code                  C 1 应答码：0-失败；1-成功
  --   message               C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_预交单号 病人预交记录.No%Type;
  n_操作状态 Number(5);
  Err_Item Exception;
  v_Err_Msg Varchar2(500);

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_预交单号 := j_Json.Get_String('deposit_no');
  n_操作状态 := Nvl(j_Json.Get_Number('oper_state'), 0);

  If v_预交单号 Is Null Then
    v_Err_Msg := '不能确定预交单据信息,不能进行作废操作！';
    Raise Err_Item;
  End If;

  --操作_In:0-删除异常充值单据，1-删除异常退款单据，2-删除异常余额退款单据 

  Zl_病人预交异常记录_Delete(v_预交单号, n_操作状态);

  Json_Out := zlJsonOut('成功', 1);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Deldepositerrorrec;
/

Create Or Replace Procedure Zl_Exsesvr_Checkexeitemvalied
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  -------------------------------------------------------------------------------------------------
  --功能：先诊疗后结算方式，检查执行项目的合法性
  --input   
  --  pati_id      N   1   病人id
  --  register_id   N   1   挂号id
  --  receipt_type  C   1   收费类别
  --output
  --  code          C   1   应答码：0-失败；1-成功
  --  message       C   1   应答消息：
  --  check_flag   N   0   检查标志：0-不检查或合法，1-提醒 ，2-拒绝
  --  check_msg    C   0   提醒或拒绝的内容提示
  -------------------------------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_病人id     Number;
  n_挂号id     Number;
  v_收费类别串 Varchar2(100);
  n_结算模式   Number(2);
  v_Check      Varchar2(1000);
  n_Count      Number;
Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_病人id     := Nvl(j_Json.Get_Number('pati_id'), 0);
  n_挂号id     := Nvl(j_Json.Get_Number('register_id'), 0);
  v_收费类别串 := j_Json.Get_String('receipt_type');

  If n_病人id = 0 Then
    Json_Out := '{"output":{"code":1,"message": "成功","check_flag":0,"check_msg":null}}';
    Return;
  Else
    If n_挂号id > 0 Then
      Select Nvl(Max(结算模式), 0) As 结算模式
      Into n_结算模式
      From 病人挂号记录
      Where 病人id = n_病人id And ID = n_挂号id;
    Else
      Select Nvl(Max(结算模式), 0) As 结算模式
      Into n_结算模式
      From 病人挂号记录
      Where 病人id = n_病人id And ID In (Select Max(ID) From 病人挂号记录 Where 病人id = n_病人id);
    End If;
  
    --未采用先诊疗后结算模式
    If n_结算模式 = 0 Then
      Json_Out := '{"output":{"code":1,"message": "成功","check_flag":0,"check_msg":null}}';
      Return;
    End If;
  End If;

  --发药时，必须先结账
  If Instr(',' || v_收费类别串 || ',', ',5,') <> 0 Or Instr(',' || v_收费类别串 || ',', ',6,') <> 0 Or
     Instr(',' || v_收费类别串 || ',', ',7,') <> 0 Then
    Select Count(1)
    Into n_Count
    From 病人未结费用
    Where 病人id = n_病人id And (来源途径 In (1, 4) Or 来源途径 = 3 And Nvl(主页id, 0) = 0) And Nvl(金额, 0) <> 0 And Rownum < 2;
    If Nvl(n_Count, 0) <> 0 Then
      --存在未结算数据，必须先结算后才允许执行
      v_Check := '2|在领药前，必须先结算后才能领药';
    End If;
  End If;

  If v_Check Is Null Then
    --检查通过时
    Json_Out := '{"output":{"code":1,"message": "成功","check_flag":0,"check_msg":null}}';
  Else
    --检查未通过，需要提醒或禁止时
    Json_Out := '{"output":{"code":1,"message": "成功","check_flag":' || Substr(v_Check, 1, 1) || ',"check_msg":"' ||
                Substr(v_Check, 3) || '"}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkexeitemvalied;
/


Create Or Replace Procedure Zl_Exsesvr_Updbillstartinvoice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:修改单据的启始发票号:预交，结帐，挂号等
  --入参：Json_In:格式
  --  input     
  --    bill_type N 1 单据类型:1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
  --    bill_nos C 1 单据号:费用单号或结算单号或预交单号
  --    inv_no  C   发票号:传空时，表示清除
  --出参: Json_Out,格式如下
  --  output
  --    code                C  1  应答码：0-失败；1-成功
  --    message             C  1  应答消息： 失败时返回具体的错误信息 
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_单据类型 Number(5);
  v_单据号   Varchar2(100);
  v_发票号   Varchar2(100);

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_单据类型 := j_Json.Get_Number('bill_type');
  v_单据号   := j_Json.Get_String('bill_nos');
  v_发票号   := j_Json.Get_String('inv_no');

  If v_单据号 Is Null Then
    v_Err_Msg := '未传入指定的单据信息!';
    Raise Err_Item;
  End If;
  If Instr(v_单据号, ',') > 0 Then
    For c_单据 In (Select Column_Value As NO From Table(f_Str2List(v_单据号))) Loop
      Zl_票据起始号_Update(c_单据.No, v_发票号, n_单据类型);
    End Loop;
  Else
  
    Zl_票据起始号_Update(v_单据号, v_发票号, n_单据类型);
  End If;

  Json_Out := zlJsonOut('成功', 1);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Updbillstartinvoice;
/

Create Or Replace Procedure Zl_Exsesvr_Reprintdepositinvc
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:重新打印预交发票
  --入参：Json_In:格式
  --input     
  -- deposit_nos C 1 单据号:预交单据号
  -- invc_no C 1 发票号
  -- invc_id N 1 领用ID
  -- user_name C 1 使用人姓名

  --出参: Json_Out,格式如下
  --  output
  --    code                C  1  应答码：0-失败；1-成功
  --    message             C  1  应答消息： 失败时返回具体的错误信息 
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_领用id     票据领用记录.Id%Type;
  v_单据号     票据打印内容.No%Type;
  v_发票号     票据使用明细.号码%Type;
  v_使用人姓名 票据使用明细.使用人%Type;
  v_Err_Msg    Varchar2(255);
  Err_Item Exception;

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_单据号 := j_Json.Get_String('deposit_nos');
  v_发票号 := j_Json.Get_String('invc_no');
  n_领用id := j_Json.Get_Number('invc_id');

  v_使用人姓名 := j_Json.Get_String('user_name');

  If v_单据号 Is Null Then
    v_Err_Msg := '未传入指定的预交单据信息!';
    Raise Err_Item;
  End If;
  Zl_病人预交记录_Reprint(v_单据号, v_发票号, n_领用id, v_使用人姓名);

  Json_Out := zlJsonOut('成功', 1);

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Reprintdepositinvc;
/


Create Or Replace Procedure Zl_Exsesvr_Checkerrordata
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：根据费用NO或费用ID检查收费记账异常信息
  --入参：Json_In:格式
  --  input
  --      fee_type              C   1 费用类别，'4'-卫材，'5,6,7'-药品
  --      rcpdtl_ids            C   1 处方明细ids,多个用逗号分隔
  --      bill_list[]                  数组，费用NO信息
  --         billtype               N   1 单据类型:1-收费处方;2-记帐处方
  --         rcp_nos                C   1 处方Nos,多个用逗号分隔
  --出参: Json_Out,格式如下
  --  output
  --     code                   N   1 应答吗：0-失败；1-成功
  --     message                C   1 应答消息：失败时返回具体的错误信息
  --     billid_list[]                   按费用ID传入时返回id列表
  --        rcpdtl_id           N   1 处方明细id
  --        fee_status          N   1 费用状态： 0-划价,1-记帐
  --        cancel_status       N   1 作废状态:0-正常状态,1-作废同步标志异常
  --        update_status       N   1 记费同步状态:0-正常状态,1-未更新药品/卫材记帐状态
  --     billno_list[]                 按NO传入时返回NO列表
  --         billtype               N   1 单据类型:1-收费处方;2-记帐处方
  --         rcp_no                 C   1 处方no
  --         fee_status             N   1 费用状态：针对收费时,0-未收费,1-已收费,2-异常收费;针对记帐时,0-划价,1-记帐
  --         cancel_status          N   1 作废状态:0-正常状态,1-作废同步标志异常
  --         update_drug_status     N   1 记费同步状态:0-正常状态,2-未更新药品/卫材收费状态
  --     expense_list[]               仅药品才有
  --         billtype               N   1 (原始)单据类型:1-收费处方;2-记帐处方
  --         rcp_no                 C   1 (原始)处方no
  --         rcpdtl_id              N   1 (原始)处方明细id
  --         rcp_no_new             C   1 新生成的处方NO
  --         rcpdtl_id_new          N   1 新生成处方明细id
  --         pati_pageid        N  1  主页ID
  ------------------------------------------------------------------------------------------------------------
  j_Input     PLJson;
  j_Json      PLJson;
  v_Output    Varchar2(32767);
  j_Bill_List Pljson_List := Pljson_List();
  v_Nos       Varchar2(4000);
  n_单据类型  Number(1); -- 1- 收费处方;2- 记帐单处方,3 - 记帐表处方 n_Count Number(3);
  v_收费类别  Varchar2(100);

  c_费用ids Clob; --费用id
  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  l_费用id Collection_Type;
  c_Output Clob;

  n_费用状态 Number(1);
  Err_Custom Exception;
  v_Err Varchar2(255);

  v_Billno_List  Varchar2(32767);
  v_Expense_List Varchar2(32767);
Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_收费类别  := j_Json.Get_String('fee_type');
  c_费用ids   := j_Json.Get_String('rcpdtl_ids');
  j_Bill_List := j_Json.Get_Pljson_List('bill_list');

  If c_费用ids Is Null And j_Bill_List Is Null Then
    v_Err := '未传入病人处方信息或处方明细ID，请检查！';
    Raise Err_Custom;
  End If;

  --1.按费用ID查找
  --将 c_费用ids 串组装成不超过4000 的集合串，防止使用 f_Num2list 参数超长
  If c_费用ids Is Not Null Then
    While c_费用ids Is Not Null Loop
      If Length(c_费用ids) <= 4000 Then
        l_费用id(l_费用id.Count) := c_费用ids;
        c_费用ids := Null;
      Else
        l_费用id(l_费用id.Count) := Substr(c_费用ids, 1, Instr(c_费用ids, ',', 3980) - 1);
        c_费用ids := Substr(c_费用ids, Instr(c_费用ids, ',', 3980) + 1);
      End If;
    End Loop;
  
    --根据病人ID查找异常单据
    v_Output := Null;
    For I In 0 .. l_费用id.Count - 1 Loop
      For r_费用 In (Select /*+cardinality(j,10)*/
                    a.Id, C1.同步标志 As 作废同步标志, c.同步标志 As 记费同步标志, Decode(a.记录状态, 0, 0, 1) As 费用状态
                   From 住院费用记录 A, Table(f_Num2List(l_费用id(I))) J, 病人费用异常记录 C, 病人费用异常记录 C1
                   Where a.Id = j.Column_Value And Instr(',' || v_收费类别 || ',', ',' || a.收费类别 || ',') > 0 And
                         a.Id = c.费用id(+) And c.产生环节(+) = 0 And a.Id = C1.费用id(+) And C1.产生环节(+) = 1
                   Union All
                   Select /*+cardinality(j,10)*/
                    a.Id, C1.同步标志 As 作废同步标志, c.同步标志 As 记费同步标志, Decode(a.记录状态, 0, 0, 1) As 费用状态
                   From 门诊费用记录 A, Table(f_Num2List(l_费用id(I))) J, 病人费用异常记录 C, 病人费用异常记录 C1
                   Where a.Id = j.Column_Value And Instr(',' || v_收费类别 || ',', ',' || a.收费类别 || ',') > 0 And
                         a.Id = c.费用id(+) And c.产生环节(+) = 0 And a.Id = C1.费用id(+) And C1.产生环节(+) = 1) Loop
      
        If Length(Nvl(v_Output, ' ')) > 30000 Then
          If c_Output Is Not Null Then
            c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
          Else
            c_Output := To_Clob(v_Output);
          End If;
          v_Output := Null;
        End If;
      
        zlJsonPutValue(v_Output, 'rcpdtl_id', r_费用.Id, 1, 1);
        zlJsonPutValue(v_Output, 'fee_status', r_费用.费用状态, 1);
        zlJsonPutValue(v_Output, 'update_status', r_费用.记费同步标志, 1);
        zlJsonPutValue(v_Output, 'cancel_status', r_费用.作废同步标志, 1, 2);
      End Loop;
    End Loop;
  
    If Not c_Output Is Null And Not v_Output Is Null Then
      c_Output := c_Output || ',' || To_Clob(v_Output);
      v_Output := Null;
    End If;
  
    If Not c_Output Is Null Then
      Json_Out := To_Clob('{"output":{"code":1,"message":"成功","billid_list":[') || c_Output || To_Clob(']}}');
    Else
      Json_Out := '{"output":{"code":1,"message":"成功","billid_list":[' || v_Output || ']}}';
    End If;
    Return;
  End If;

  --2.按处方NO查找
  For I In 1 .. j_Bill_List.Count Loop
    j_Json := PLJson();
  
    j_Json     := PLJson(j_Bill_List.Get(I));
    n_单据类型 := j_Json.Get_Number('billtype');
    v_Nos      := j_Json.Get_String('rcp_nos');
    If Nvl(n_单据类型, 0) = 0 Then
      v_Err := '未传入单据类型，请检查！';
      Raise Err_Custom;
    End If;
  
    If Nvl(v_Nos, '-') = '-' Then
      v_Err := '未传入处方NO，请检查！';
      Raise Err_Custom;
    End If;
  
    v_Output := Null;
    For c_费用信息 In (Select a.No, Max(c.同步标志) As 记费同步标志, Max(C1.同步标志) As 作废同步标志, Max(Decode(n_单据类型, 1, a.费用状态, 0)) As 费用状态,
                          Max(Decode(a.记录状态, 0, 0, 1)) As 记录状态
                   From 门诊费用记录 A, 病人费用异常记录 C, 病人费用异常记录 C1
                   Where a.记录性质 = n_单据类型 And a.Id = c.费用id(+) And c.产生环节(+) = 0 And a.Id = C1.费用id(+) And C1.产生环节(+) = 1 And
                         a.No In (Select /*+cardinality(B,10) */
                                   Column_Value
                                  From Table(f_Str2List(v_Nos)) B) And
                         Instr(',' || v_收费类别 || ',', ',' || a.收费类别 || ',') > 0
                   Group By a.No
                   Union All
                   Select a.No, Max(c.同步标志) As 记费同步标志, Max(C1.同步标志) As 作废同步标志, Max(Decode(n_单据类型, 1, a.费用状态, 0)) As 费用状态,
                          Max(Decode(a.记录状态, 0, 0, 1)) As 记录状态
                   From 住院费用记录 A, 病人费用异常记录 C, 病人费用异常记录 C1
                   Where a.记录性质 = n_单据类型 And a.Id = c.费用id(+) And c.产生环节(+) = 0 And a.Id = C1.费用id(+) And C1.产生环节(+) = 1 And
                         a.No In (Select /*+cardinality(B,10) */
                                   Column_Value
                                  From Table(f_Str2List(v_Nos)) B) And
                         Instr(',' || v_收费类别 || ',', ',' || a.收费类别 || ',') > 0
                   Group By a.No) Loop
    
      If Nvl(c_费用信息.费用状态, 0) = 1 Then
        n_费用状态 := 2;
      Else
        n_费用状态 := c_费用信息.记录状态;
      End If;
    
      zlJsonPutValue(v_Output, 'billtype', n_单据类型, 1, 1);
      zlJsonPutValue(v_Output, 'rcp_no', c_费用信息.No);
      zlJsonPutValue(v_Output, 'fee_status', n_费用状态, 1);
      zlJsonPutValue(v_Output, 'cancel_status', c_费用信息.作废同步标志, 1);
      zlJsonPutValue(v_Output, 'update_status', c_费用信息.记费同步标志, 1, 2);
    
    End Loop;
    If v_Output Is Not Null Then
      If v_Billno_List Is Null Then
        v_Billno_List := v_Output;
      Else
        v_Billno_List := v_Billno_List || ',' || v_Output;
      End If;
    End If;
  
    --获取门诊费用转住院的费用信息
    v_Output := Null;
    For c_转费用信息 In (Select b.Id As 原始id, b.No As 原始no, a.Id As 转入id, a.No As 转入no, a.主页id
                    From 住院费用记录 A, 门诊费用记录 B, 费用审核记录 C, 病人费用异常记录 D
                    Where a.Id = c.转出id And b.Id = c.费用id And b.记录性质 = n_单据类型 And a.Id = d.费用id And d.产生环节 = 2 And
                          b.No In (Select /*+cardinality(B,10) */
                                    Column_Value
                                   From Table(f_Str2List(v_Nos)) B) And
                          Instr(',' || v_收费类别 || ',', ',' || a.收费类别 || ',') > 0) Loop
    
      zlJsonPutValue(v_Output, 'billtype', n_单据类型, 1, 1);
      zlJsonPutValue(v_Output, 'rcpdtl_id', c_转费用信息.原始id, 1);
      zlJsonPutValue(v_Output, 'rcp_no', c_转费用信息.原始no);
      zlJsonPutValue(v_Output, 'rcpdtl_id_new', c_转费用信息.转入id, 1);
      zlJsonPutValue(v_Output, 'rcp_no_new', c_转费用信息.转入no, 0);
      zlJsonPutValue(v_Output, 'pati_pageid', c_转费用信息.主页id, 1, 2);
    End Loop;
    If v_Output Is Not Null Then
      If v_Expense_List Is Null Then
        v_Expense_List := v_Output;
      Else
        v_Expense_List := v_Expense_List || ',' || v_Output;
      End If;
    End If;
  
  End Loop;
  v_Billno_List  := ',"billno_list":[' || v_Billno_List || ']';
  v_Expense_List := ',"expense_list":[' || v_Expense_List || ']';

  Json_Out := To_Clob('{"output":{"code":1,"message":"成功"' || To_Clob(v_Billno_List) || To_Clob(v_Expense_List) || '}}');
Exception
  When Err_Custom Then
    Json_Out := zlJsonOut(v_Err);
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkerrordata;
/


Create Or Replace Procedure Zl_Exsesvr_Checkpatichangeundo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:撤销病人变动记录前检查
  --入参：Json_In:格式
  --input
  --    pati_list[]       数组 床位对换撤销时须同时检查两个病人
  --      pati_id           N 1 病人id 
  --      pati_pageid       N 1 主页ID
  --      undo_type         C 1 撤销类型
  --      create_time       C 1 登记时间
  --      fee_item_id       N 1 费用项目ID
  --出参  json
  --output
  --     code               N 1 应答码：0-失败；1-成功
  --     message            C 1 应答消息： 失败时返回具体的错误信息
  ---------------------------------------------------------------------------

  j_Input PLJson;
  j_Json  PLJson;

  j_Temp       PLJson;
  j_Json_List  Pljson_List;
  n_病人id     住院费用记录.病人id%Type;
  n_主页id     住院费用记录.主页id%Type;
  n_费用项目id Number(18);
  v_Undo_Type  Varchar2(100);

  d_开始时间 Date;

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  j_Json_List := j_Json.Get_Pljson_List('pati_list');

  If j_Json_List Is Null Then
    Json_Out := zlJsonOut('传入值有误,请检查。');
    Return;
  End If;
  For I In 1 .. j_Json_List.Count Loop
    j_Temp := PLJson(j_Json_List.Get(I));
  
    n_病人id     := j_Temp.Get_Number('pati_id');
    n_主页id     := j_Temp.Get_Number('pati_pageid');
    n_费用项目id := j_Temp.Get_Number('fee_item_id');
    v_Undo_Type  := j_Temp.Get_String('undo_type');
    d_开始时间   := To_Date(j_Temp.Get_String('create_time'), 'YYYY-MM-DD HH24:MI:SS');
    If Nvl(n_病人id, 0) = 0 Then
      Json_Out := zlJsonOut('未传入病人id，请检查!');
      Return;
    End If;
    If Instr(',入院入住,入住,转科入住,换床,床位对换,转病区入住,', ',' || v_Undo_Type || ',') > 0 Then
      For r_Fee In (Select NO, 姓名
                    From 住院费用记录
                    Where 病人id = n_病人id And 主页id = n_主页id And Mod(记录性质, 10) = 3 And 登记时间 >= d_开始时间
                    Group By NO, 序号, Mod(记录性质, 10), 姓名
                    Having Sum(结帐金额) <> 0) Loop
        If v_Undo_Type = '床位对换' Then
          Json_Out := zlJsonOut('病人 ' || r_Fee.姓名 || ' 的自动记帐费用已结帐,不能进行撤销操作！');
        Else
          Json_Out := zlJsonOut('该病人的自动记帐费用已结帐,不能进行撤销操作！');
        End If;
        Return;
      End Loop;
    Elsif Instr(',床位等级变动,护理等级变动,', ',' || v_Undo_Type || ',') > 0 Then
      -- v_Undo_Type = '床位等级变动' Or v_Undo_Type = '护理等级变动' 
      For r_Fee In (Select NO
                    From 住院费用记录
                    Where 病人id = n_病人id And 主页id = n_主页id And Mod(记录性质, 10) = 3 And 收费细目id = n_费用项目id And 登记时间 >= d_开始时间
                    Group By NO, 序号, Mod(记录性质, 10)
                    Having Sum(结帐金额) <> 0) Loop
        Json_Out := zlJsonOut('该病人的自动记帐费用已结帐,不能进行撤销操作！');
        Return;
      End Loop;
    End If;
  End Loop;
  Json_Out := zlJsonOut('成功', 1);
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkpatichangeundo;
/


Create Or Replace Procedure Zl_Exsesvr_Getconsumercardinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：获取消费卡信息
  --入参：Json_In:格式
  --  input
  --   cardno               C 1 卡号
  --   cardtype_num         N 1 接口编号
  --   check_valid          N   有效性检查：1-检查；0-不检查 
  --出参: Json_Out,格式如下
  --  output
  --    code              N 1 应答码：0-失败；1-成功
  --    message           C 1 应答消息：失败时返回具体的错误信息
  --    card_id           N 1 消费卡id
  --    card_pwd          C 1 密码
  --    surplus           N 1 余额
  --    limit_type        N 1 限制类别
  --    occasion          N 1 应用场合
  --    pati_id           N 1 病人ID
  --    specpati          N 1 是否特定病人
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_卡号     消费卡信息.卡号%Type;
  n_接口编号 消费卡信息.接口编号%Type;

  n_充值       消费卡信息.可否充值%Type;
  n_Id         消费卡信息.Id%Type;
  n_病人id     消费卡信息.病人id%Type;
  d_有效期     消费卡信息.有效期%Type;
  v_密码       消费卡信息.密码%Type;
  d_回收时间   消费卡信息.回收时间%Type;
  v_当前状态   Varchar2(20);
  n_余额       消费卡信息.余额%Type;
  d_停用日期   消费卡信息.停用日期%Type;
  v_限制类别   消费卡信息.限制类别%Type;
  v_应用场合   消费卡类别目录.应用场合%Type;
  n_特定病人   消费卡类别目录.是否特定病人%Type;
  n_失效面额   帐户缴款余额.余额%Type;
  n_交易序号   帐户缴款余额.交易序号%Type;
  n_Count      Number(5);
  v_Message    Varchar2(2000);
  v_Output     Varchar2(32767);
  n_有效性检查 Number(1);
Begin

  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_卡号       := j_Json.Get_String('cardno');
  n_接口编号   := j_Json.Get_Number('cardtype_num');
  n_有效性检查 := Nvl(j_Json.Get_Number('check_valid'), 0);

  If Nvl(v_卡号, '-') = '-' Or Nvl(n_接口编号, 0) = 0 Then
    Json_Out := zlJsonOut('未传入任何卡号或接口编号，请检查!');
    Return;
  End If;

  Select Count(1), Max(a.Id), Max(a.可否充值), Max(a.有效期), Max(a.密码), Max(a.回收时间),
         Max(Decode(a.当前状态, 2, '回收', 3, '退卡', '回收')), Max(a.余额), Max(a.停用日期), Max(a.限制类别), Max(b.应用场合), Max(a.病人id),
         Max(b.是否特定病人)
  Into n_Count, n_Id, n_充值, d_有效期, v_密码, d_回收时间, v_当前状态, n_余额, d_停用日期, v_限制类别, v_应用场合, n_病人id, n_特定病人
  From 消费卡信息 A, 消费卡类别目录 B
  Where a.接口编号 = b.编号 And a.卡号 = v_卡号 And a.接口编号 = n_接口编号 And
        序号 = (Select Max(序号) From 消费卡信息 B Where 卡号 = a.卡号 And 接口编号 = a.接口编号)
  Order By a.序号;

  If n_Count = 0 Then
    Json_Out := zlJsonOut('该卡不是有效卡!');
    Return;
  End If;

  --是否回收
  If n_有效性检查 = 1 And Nvl(d_回收时间, To_Date('3000-01-01 00:00:00', 'yyyy-mm-dd hh24:mi:ss')) <
     To_Date('3000-01-01 00:00:00', 'yyyy-mm-dd hh24:mi:ss') Then
    v_Message := '该卡已经被' || Nvl(v_当前状态, '回收') || ',不能刷卡消费!';
    Json_Out  := zlJsonOut(v_Message);
    Return;
  End If;

  --是否停用
  If n_有效性检查 = 1 And Nvl(d_停用日期, To_Date('3000-01-01 00:00:00', 'yyyy-mm-dd hh24:mi:ss')) <
     To_Date('3000-01-01 00:00:00', 'yyyy-mm-dd hh24:mi:ss') Then
    v_Message := '该卡已经被停止使用,不能刷卡消费!';
    Json_Out  := zlJsonOut(v_Message);
    Return;
  End If;

  --检查有效期
  If Nvl(d_有效期, To_Date('3000-01-01 00:00:00', 'yyyy-mm-dd hh24:mi:ss')) <
     To_Date('3000-01-01 00:00:00', 'yyyy-mm-dd hh24:mi:ss') Then
    --是否允许充值
    If n_有效性检查 = 1 And Nvl(n_充值, 0) <> 1 Then
      v_Message := '该卡已经失效,不能刷卡消费!';
      Json_Out  := zlJsonOut(v_Message);
      Return;
    End If;
    --获取实际可用余额(余额-失效面额)
    --升级后的发卡记录(交易序号>0)，直接取失效金额
    Select Count(1), Nvl(Max(b.交易序号), 0), Nvl(Max(b.余额), 0)
    Into n_失效面额, n_交易序号, n_Count
    From 病人卡结算记录 A, 帐户缴款余额 B
    Where a.交易序号 = b.交易序号 And a.消费卡id = b.消费卡id And a.记录性质 = 1 And a.消费卡id = n_Id;
  
    If n_Count > 0 And n_交易序号 = 0 Then
      --升级前的发卡记录(交易序号=0)，需要统计失效金额
      Select Sum(Nvl(失效金额, 0))
      Into n_失效面额
      From (Select 卡面金额 As 失效金额
             From 消费卡信息 A
             Where ID = n_Id And 有效期 < Sysdate
             Union All
             Select Nvl(Sum(a.应收金额), 0) As 失效金额
             From 病人卡结算记录 A, 消费卡信息 B
             Where a.消费卡id = b.Id And a.记录性质 = 4 And a.消费卡id = n_Id And
                   a.交易时间 <= Nvl(b.有效期, To_Date('3000-01-01', 'yyyy-mm-dd')));
    End If;
    n_余额 := n_余额 - n_失效面额;
  End If;

  --    card_id           N 1 消费卡id
  --    card_pwd          C 1 密码
  --    surplus           N 1 余额
  --    limit_type        N 1 限制类别
  --    occasion          N 1 应用场合
  --    pati_id           N 1 病人ID
  --    specpati          N 1 是否特定病人

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '成功');
  zlJsonPutValue(v_Output, 'card_id', n_Id, 1);
  zlJsonPutValue(v_Output, 'surplus', Nvl(n_余额, 0), 1);
  zlJsonPutValue(v_Output, 'card_pwd', Nvl(v_密码, ''));
  zlJsonPutValue(v_Output, 'limit_type', Nvl(v_限制类别, ''));
  zlJsonPutValue(v_Output, 'occasion', Nvl(v_应用场合, '000'));
  zlJsonPutValue(v_Output, 'pati_id', Nvl(n_病人id, 0), 1);
  zlJsonPutValue(v_Output, 'specpati', Nvl(n_特定病人, 0), 1, 2);
  Json_Out := '{"output":' || v_Output || '}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getconsumercardinfo;
/

CREATE OR REPLACE Procedure Zl_Exsesvr_Sync_Update
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  ---------------------------------------------------------------------------
  --功能：费用同步后清空记费同步标志（按NO或按费用ID）
  --入参：Json_In:格式
  --  input
  --    sign_type           N 1 标志类型：0-记费同步标志和作风同步标志,1-转费同步标志
  --    detail_ids          C  1  处方明细id串(费用id串),支持多个id，用“,”分隔
  --    bill_list[]
  --      billtype          N   1 单据类型:1-收费处方;2-记帐处方
  --      rcp_no            C   1 处方No
  --出参: Json_Out,格式如下
  --  output
  --    code                 N   1   应答吗：0-失败；1-成功
  --    message              C   1   应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  c_Detailids Clob;
  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  l_费用ids Collection_Type;

  I           Number;
  j_Bill_List Pljson_List;
  o_Json      PLJson;
  n_性质      Number(1);
  v_No        Varchar2(20);
  n_标志类型  Number(2);

  j_Input PLJson;
  j_Json  PLJson;
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_标志类型 := j_Json.Get_Number('sign_type');

  --1.按单据更新
  j_Bill_List := j_Json.Get_Pljson_List('bill_list');
  If j_Bill_List Is Not Null Then
    For I In 1 .. j_Bill_List.Count Loop
      o_Json := PLJson();
      o_Json := PLJson(j_Bill_List.Get(I));
      n_性质 := o_Json.Get_Number('billtype');
      v_No   := o_Json.Get_String('rcp_no');

      If Nvl(n_标志类型, 0) = 0 Then
        Delete From 病人费用异常记录 a
        Where (a.产生环节 = 0 Or a.产生环节 = 1) and a.费用ID In (Select ID From 住院费用记录 Where 记录状态 In (1, 3) And 记录性质 = n_性质 And NO = v_No);
        If Sql%NotFound Then
          Delete From 病人费用异常记录 a
          Where (a.产生环节 = 0 Or a.产生环节 = 1) and a.费用ID In (Select ID From 门诊费用记录 Where 记录状态 In (1, 3) And 记录性质 = n_性质 And NO = v_No);
        End If;
      Elsif n_标志类型 = 1 Then
        Delete From 病人费用异常记录 a
        Where a.产生环节 = 2 and a.费用ID In (Select ID From 住院费用记录 Where 记录状态 In (1, 3) And 记录性质 = n_性质 And NO = v_No);
      End If;
    End Loop;
  End If;

  --2.按费用ID更新
  c_Detailids := j_Json.Get_Clob('detail_ids');
  I           := 1;
  While c_Detailids Is Not Null Loop
    If Length(c_Detailids) <= 4000 Then
      l_费用ids(I) := c_Detailids;
      c_Detailids := Null;
    Else
      l_费用ids(I) := Substr(c_Detailids, 1, Instr(c_Detailids, ',', 3980) - 1);
      c_Detailids := Substr(c_Detailids, Instr(c_Detailids, ',', 3980) + 1);
    End If;
    I := I + 1;
  End Loop;

  If Nvl(n_标志类型, 0) = 0 Then
    Forall I In 1 .. l_费用ids.Count
      Delete 病人费用异常记录
      Where (产生环节 = 0 Or 产生环节 = 1) and 费用ID In (Select /*+Cardinality(j,10)*/
                                    j.Column_Value As ID
                                   From Table(f_Num2List(l_费用ids(I))) J);
  Elsif n_标志类型 = 1 Then
    Forall I In 1 .. l_费用ids.Count
      Delete 病人费用异常记录
        Where 产生环节 = 2 and 费用ID In (Select /*+Cardinality(j,10)*/
                                      j.Column_Value As ID
                                     From Table(f_Num2List(l_费用ids(I))) J);
  End If;

  Json_Out := zlJsonOut('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Sync_Update;
/


Create Or Replace Procedure Zl_Exsesvr_Getrelatedtransinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:根据关联交易id,获取交易信息s
  --入参：Json_In:格式
  --input
  --  related_ids  C 1 关联交易ID:多个用逗号分离

  --出参: Json_Out,格式如下
  --  output
  --    code                      N   1   应答码：0-失败；1-成功
  --    message                   C   1   应答消息：失败时返回具体的错误信息
  --     swap_list[]  C 1 交易信息列表
  --      related_id N 1 关联交易ID
  --      cardtype_id N 1 卡类别ID
  --      blnc_Mode C 1 结算方式
  --      swapno  C 1 交易流水号
  --      swapmemo  C 1 交易说明
  --      original_money  N 1 原始金额
  --      return_money  N 1 已退金额

  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_关联交易ids Varchar2(32680);

  v_Output Varchar2(32767);

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_关联交易ids := j_Json.Get_String('related_ids');
  If v_关联交易ids Is Null Then
    Json_Out := zlJsonOut('未传入关联交易信息!');
    Return;
  End If;
  For c_结算信息 In (With 关联交易 As
                    (Select Column_Value As 关联交易id From Table(f_Num2List(v_关联交易ids)))
                   Select /*+cardinality(B,10)*/
                    关联交易id, 卡类别id, a.结算方式, a.交易流水号, a.交易说明, Sum(原始金额) As 原始金额, Sum(已退金额) As 已退金额,
                    Sum(原始金额) - Sum(已退金额) As 剩余未退金额
                   From (Select a.关联交易id, a.卡类别id, a.结算方式, a.交易流水号, a.交易说明,
                                 Decode(a.记录性质, 1, Decode(a.记录状态, 2, 0, 1), 1) * Nvl(金额, 0) +
                                  Decode(Mod(记录性质, 10), 1, 0, 1) * Decode(Sign(Nvl(冲预交, 0)), 1, 1, 0) * Nvl(冲预交, 0) As 原始金额,
                                 (Decode(Sign(Nvl(金额, 0)), -1, 1, 0) * Nvl(金额, 0) +
                                  Decode(Sign(Nvl(冲预交, 0)), -1, 1, 0) * Nvl(冲预交, 0)) * Decode(Nvl(a.校对标志, 0), 1, 0) As 已退金额
                          From 病人预交记录 A, 关联交易 B
                          Where a.关联交易id = b.关联交易id
                          Union All
                          Select a.关联交易id, a.卡类别id, a.结算方式, a.交易流水号, a.交易说明, 0 As 原始金额, -1 * Nvl(b.金额, 0) As 已退金额
                          From 病人预交记录 A, 三方退款信息 B, 关联交易 C
                          Where a.Id = b.记录id And a.关联交易id = c.关联交易id And b.是否转帐 = 1) A
                   Group By a.关联交易id, a.卡类别id, a.结算方式, a.交易流水号, a.交易说明) Loop
  
    zlJsonPutValue(v_Output, 'related_id', Nvl(c_结算信息.关联交易id, 0), 1, 1);
    zlJsonPutValue(v_Output, 'cardtype_id', Nvl(c_结算信息.卡类别id, 0), 1);
    zlJsonPutValue(v_Output, 'blnc_mode', Nvl(c_结算信息.结算方式, ''));
    zlJsonPutValue(v_Output, 'swapno', Nvl(c_结算信息.交易流水号, ''));
    zlJsonPutValue(v_Output, 'swapmemo', Nvl(c_结算信息.交易说明, ''));
    zlJsonPutValue(v_Output, 'original_money', Nvl(c_结算信息.原始金额, 0), 1);
    zlJsonPutValue(v_Output, 'return_money', Nvl(c_结算信息.已退金额, 0), 1, 2);
  
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"成功","swap_list":[' || v_Output || ']}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getrelatedtransinfo;
/



Create Or Replace Procedure Zl_Exsesvr_Getspeccalcfeeitem
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  --------------------------------------------------------------------------------------
  --功能：获取打折费别明细列表
  --入参：Json_In:格式
  --input
  --出参      json
  --output
  --    code                N   1 应答吗：0-失败；1-成功
  --    message             C   1 应答消息：失败时返回具体的错误信息
  --    feecategory_list           费别明细列表
  --       fee_category      C   1 费别名称
  --       fee_item_id       N   1 收费项目ID
  --       detail_cacfml     N   1 计算方式
  --------------------------------------------------------------------------------------
  v_Output Varchar2(32767);
Begin
  For c_费别明细 In (Select Distinct 费别, 收费细目id, 计算方法
                 From 费别明细
                 Where 计算方法 = 1 And 收入项目id Is Null And 收费细目id Is Not Null) Loop
    zlJsonPutValue(v_Output, 'fee_category', c_费别明细.费别, 0, 1);
    zlJsonPutValue(v_Output, 'fee_item_id', c_费别明细.收费细目id, 1);
    zlJsonPutValue(v_Output, 'detail_cacfml', c_费别明细.计算方法, 1, 2);
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"成功","feecategory_list":[' || v_Output || ']}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getspeccalcfeeitem;
/


Create Or Replace Procedure Zl_Exsesvr_Executeturnwardfee
(
  Json_In  Clob,
  Json_Out Out Varchar2
) Is
  ---------------------------------------------------------------------------
  --功能:执行病人转病区费用的转入，转出处理
  --入参：Json_In:格式
  --input
  --  oper_type              N   1   操作类型：0-病区变动,1-撤消病区变动
  --  change_id_old          N   1   原病区的变动记录的ID
  --  change_id_new          N   1   目标病区的变动记录的ID
  --  ward_id_old            N   1   原病区ID
  --  ward_id_new            N   1   目标病区ID
  --  pat_visit_pnurs        C   1   责任护士姓名
  --  operator_code          C   1   操作员编号
  --  operator_name          C   1   操作员姓名
  --  pati_info              病人信息，仅审核这些病人的费用
  --    pati_id              N   1   病人ID
  --    pati_pageid          N   1   主页ID
  --    pati_name            C   1   病人姓名
  --    fee_audit_status     N   1   费用审核标志:0或空-未审核;1-已审核或开始审核;2-完成审核
  --    si_inp_status        N   1   住院状态:0-正常住院;1-尚未入科;2-正在转科或正在转病区;3-已预出院
  --    catalog_date         C   0   病案编目日期：yyyy-mm-dd hh24:mi:ss
  --  bill_list[]            转费用单据信息
  --    fee_no               C   1   费用单据号
  --    serial_num           N   1   序号
  --    quantity             N   1   转出数量
  --  excute_list[]          单据已执行列表(卫材费用),即使已执行数为0也要传入
  --    fee_id               N   1   费用ID
  --    sended_num           N   1   已发数量
  --出参: Json_Out,格式如下
  --  output
  --    code                C  1 应答码：0-失败；1-成功
  --    message             C  1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  --转入，转出规则:
  --1.病区执行的非药品和卫生材料，处理规则为
  --   1)将原记录进行销帐处理
  --   2)新增一条新病区的费用，病人科室，发生时间不变
  --2.病区执行的药品和卫生材料
  --   这个卫材退的处理在转病区时的界面中进行确认(可以打印核查清单)，在转病区发起的时候确认。
  --   a)卫材在原病区通过销帐申请来处理，新病区手工计卫材；
  --   b)撤消转病区时，自动撤消销帐申请，如果已经销帐审核了，则询问提示并且不作卫材费用处理，手工去处理。
  j_Input PLJson;
  j_Json  PLJson;

  j_List Pljson_List;
  j_Temp PLJson;

  Err_Item Exception;
  v_Err_Msg Varchar2(200);

  n_操作类型   Number(1);
  n_原变动id   费用变动记录.原变动id%Type;
  n_目标变动id 费用变动记录.目标变动id%Type;
  n_原病区id   部门表.Id%Type;
  n_目标病区id 部门表.Id%Type;
  v_责任护士   病人费用销帐.申请人%Type;
  v_操作员编号 住院费用记录.操作员编号%Type;
  v_操作员姓名 住院费用记录.操作员姓名%Type;

  n_病人id   住院费用记录.病人id%Type;
  n_主页id   住院费用记录.主页id%Type;
  v_病人姓名 住院费用记录.姓名%Type;
  n_审核标志 Number(2);
  n_住院状态 Number(2);
  v_编目日期 Varchar2(30);

  n_费用id   住院费用记录.Id%Type;
  n_已执行数 住院费用记录.数次%Type;
  v_No       住院费用记录.No%Type;
  n_Max序号  住院费用记录.序号%Type;
  n_Dec      Number;
  d_登记时间 Date;

  n_转出数量 住院费用记录.数次%Type;
  n_执行数量 住院费用记录.数次%Type;

  n_未发料数量   病人费用销帐.数量%Type;
  n_申请未发数量 病人费用销帐.数量%Type;
  n_申请已发数量 病人费用销帐.数量%Type;
  n_申请销账数量 病人费用销帐.数量%Type;
  n_申请取消数量 病人费用销帐.数量%Type;

  v_原病区名称   部门表.名称%Type;
  v_目标病区名称 部门表.名称%Type;
  n_应收金额     住院费用记录.应收金额%Type;
  n_实收金额     住院费用记录.实收金额%Type;

  Type t_Table Is Record(
    NO   门诊费用记录.No%Type,
    序号 门诊费用记录.序号%Type,
    数量 门诊费用记录.数次%Type);
  Type t_Fee_Table Is Table Of t_Table;

  l_Fee      t_Fee_Table;
  l_Feeno    t_StrList2;
  l_Executed t_NumList2;

  Procedure 销帐申请_Insert
  (
    费用id_In     病人费用销帐.费用id%Type,
    申请类别_In   病人费用销帐.申请类别%Type,
    收费细目id_In 病人费用销帐.收费细目id%Type,
    申请部门id_In 病人费用销帐.申请部门id%Type,
    审核部门id_In 病人费用销帐.审核部门id%Type,
    数量_In       病人费用销帐.数量%Type,
    申请人_In     病人费用销帐.申请人%Type,
    申请时间_In   病人费用销帐.申请时间%Type,
    状态_In       病人费用销帐.状态%Type,
    销帐原因_In   病人费用销帐.销帐原因%Type
  ) Is
  Begin
    --全部都执行了，肯定销帐数量为已执行的
    Insert Into 病人费用销帐
      (费用id, 申请类别, 收费细目id, 审核部门id, 申请部门id, 数量, 申请人, 申请时间, 状态, 销帐原因)
    Values
      (费用id_In, 申请类别_In, 收费细目id_In, 审核部门id_In, 申请部门id_In, 数量_In, 申请人_In, 申请时间_In, 状态_In, 销帐原因_In);
  End;
Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_操作类型   := j_Json.Get_Number('oper_type');
  n_原变动id   := j_Json.Get_Number('change_id_old');
  n_目标变动id := j_Json.Get_Number('change_id_new');
  n_原病区id   := j_Json.Get_Number('ward_id_old');
  n_目标病区id := j_Json.Get_Number('ward_id_new');
  v_责任护士   := j_Json.Get_String('pat_visit_pnurs');
  v_操作员编号 := j_Json.Get_String('operator_code');
  v_操作员姓名 := j_Json.Get_String('operator_name');

  n_病人id   := j_Json.Get_Number('pati_info.pati_id');
  n_主页id   := j_Json.Get_Number('pati_info.pati_pageid');
  v_病人姓名 := j_Json.Get_String('pati_info.pati_name');
  n_审核标志 := j_Json.Get_Number('pati_info.fee_audit_status');
  n_住院状态 := j_Json.Get_Number('pati_info.si_inp_status');
  v_编目日期 := j_Json.Get_String('pati_info.catalog_date');

  v_Err_Msg := Zl_Pati_Charge_Check(v_病人姓名, n_审核标志, n_住院状态, v_编目日期);
  If v_Err_Msg Is Not Null Then
    Json_Out := zlJsonOut(v_Err_Msg);
    Return;
  End If;

  --解析转出费用数据
  l_Fee   := t_Fee_Table();
  l_Feeno := t_StrList2();
  j_List  := j_Json.Get_Pljson_List('bill_list');
  If j_List Is Not Null Then
    For I In 1 .. j_List.Count Loop
      j_Temp := PLJson(j_List.Get(I));
      l_Fee.Extend;
      l_Fee(l_Fee.Count).No := j_Temp.Get_String('fee_no');
      l_Fee(l_Fee.Count).序号 := j_Temp.Get_Number('serial_num');
      l_Fee(l_Fee.Count).数量 := j_Temp.Get_Number('quantity');
    
      l_Feeno.Extend;
      l_Feeno(l_Feeno.Count) := t_StrObj2(l_Fee(l_Fee.Count).No, l_Fee(l_Fee.Count).序号);
    End Loop;
  End If;

  --解析卫生材料费用的已执行数
  l_Executed := t_NumList2();
  j_List     := Pljson_List();
  j_List     := j_Json.Get_Pljson_List('excute_list');
  If j_List Is Not Null Then
    For I In 1 .. j_List.Count Loop
      j_Temp     := PLJson();
      j_Temp     := PLJson(j_List.Get(I));
      n_费用id   := j_Temp.Get_Number('fee_id');
      n_已执行数 := j_Temp.Get_Number('sended_num');
    
      l_Executed.Extend;
      l_Executed(l_Executed.Count) := t_NumObj2(n_费用id, n_已执行数);
    End Loop;
  End If;

  Select Max(Decode(ID, n_原病区id, 名称, Null)), Max(Decode(ID, n_目标病区id, 名称, Null))
  Into v_原病区名称, v_目标病区名称
  From 部门表
  Where ID In (n_原病区id, n_目标病区id);

  d_登记时间 := Sysdate;
  --金额小数位数
  n_Dec := zl_To_Number(Nvl(zl_GetSysParameter(9), '2'));

  n_Max序号 := 0;
  v_No      := '-~';

  For r_费用 In (Select a.Id As 费用id, a.No, Nvl(a.价格父号, 序号) As 序号, a.收费细目id, a.医嘱序号 As 医嘱id, a.记录状态,
                      Nvl(a.付数, 1) * a.数次 As 数量, a.标准单价, a.应收金额, a.实收金额, a.收费类别 As 收费类别, Nvl(c.跟踪在用, 0) As 卫材, a.执行状态
               From 住院费用记录 A, Table(l_Feeno) B, 材料特性 C
               Where a.记录性质 = 2 And a.No = b.C1 And Nvl(a.价格父号, a.序号) = b.C2 And a.收费细目id = c.材料id(+) And
                     a.记录状态 In (0, 1, 3)
               Order By NO, 序号) Loop
  
    If v_No <> r_费用.No Then
      v_No := r_费用.No;
      Select Nvl(Max(序号), 0)
      Into n_Max序号
      From 住院费用记录
      Where NO = v_No And 记录性质 = 2 And 记录状态 In (0, 1, 3);
    End If;
  
    n_转出数量 := 0;
    For I In 1 .. l_Fee.Count Loop
      If l_Fee(I).No = r_费用.No And l_Fee(I).序号 = r_费用.序号 Then
        n_转出数量 := l_Fee(I).数量;
        Exit;
      End If;
    End Loop;
  
    --1.卫生材料在病区执行的，直接发出销帐申请
    If Nvl(r_费用.卫材, 0) = 1 Then
      If r_费用.记录状态 = 0 Then
        v_Err_Msg := '单据 ' || r_费用.No || ' 还未进行审核，禁止转病区操作！';
        Raise Err_Item;
      End If;
      If v_责任护士 Is Null Then
        v_Err_Msg := '原病区的责任护士不存在，不能进行卫材销帐申请！';
        Raise Err_Item;
      End If;
    
      Select Max(C2) Into n_执行数量 From Table(l_Executed) Where C1 = r_费用.费用id;
      n_未发料数量 := Nvl(n_转出数量, 0) - Nvl(n_执行数量, 0);
    
      Select Sum(Decode(申请类别, 0, 1, 0) * 数量), Sum(Decode(申请类别, 0, 0, 1) * 数量)
      Into n_申请未发数量, n_申请已发数量
      From 病人费用销帐
      Where 费用id = r_费用.费用id And Nvl(状态, 0) = 0;
    
      n_申请销账数量 := 0;
      If Nvl(n_未发料数量, 0) = Nvl(n_转出数量, 0) Then
        --都未执行
        n_申请未发数量 := Nvl(n_转出数量, 0) - Nvl(n_申请未发数量, 0);
        If n_申请未发数量 > 0 Then
          n_申请销账数量 := n_申请销账数量 + n_申请未发数量;
          销帐申请_Insert(r_费用.费用id, 0, r_费用.收费细目id, n_原病区id, n_原病区id, n_申请未发数量, v_责任护士, d_登记时间, 0,
                      '从' || v_原病区名称 || '转到' || v_目标病区名称);
        End If;
      Elsif Nvl(n_未发料数量, 0) = 0 Then
        --全部都执行了，肯定销帐数量为已执行的
        n_申请已发数量 := Nvl(n_转出数量, 0) - Nvl(n_申请已发数量, 0);
        If n_申请已发数量 > 0 Then
          n_申请销账数量 := n_申请销账数量 + n_申请已发数量;
          销帐申请_Insert(r_费用.费用id, 1, r_费用.收费细目id, n_原病区id, n_原病区id, n_申请已发数量, v_责任护士, d_登记时间, 0,
                      '从' || v_原病区名称 || '转到' || v_目标病区名称);
        End If;
      Else
        --可能有部分对执行的进行销帐，一部分对未执行的销帐
        n_申请未发数量 := Nvl(n_未发料数量, 0) - Nvl(n_申请未发数量, 0);
        If n_申请未发数量 > 0 Then
          n_申请销账数量 := n_申请销账数量 + n_申请未发数量;
          销帐申请_Insert(r_费用.费用id, 0, r_费用.收费细目id, n_原病区id, n_原病区id, n_申请未发数量, v_责任护士, d_登记时间, 0,
                      '从' || v_原病区名称 || '转到' || v_目标病区名称);
        End If;
        --已执行部分
        n_申请已发数量 := Nvl(n_转出数量, 0) - Nvl(n_未发料数量, 0) - Nvl(n_申请已发数量, 0);
        If n_申请已发数量 > 0 Then
          n_申请销账数量 := n_申请销账数量 + n_申请已发数量;
          销帐申请_Insert(r_费用.费用id, 1, r_费用.收费细目id, n_原病区id, n_原病区id, n_申请已发数量, v_责任护士, d_登记时间, 0,
                      '从' || v_原病区名称 || '转到' || v_目标病区名称);
        End If;
      End If;
    
      --增加变动记录
      If Nvl(n_申请销账数量, 0) > 0 Then
        --金额=剩余金额*(准退数/剩余数)
        n_应收金额 := Round(r_费用.应收金额 * (n_申请销账数量 / r_费用.数量), n_Dec);
        n_实收金额 := Round(r_费用.实收金额 * (n_申请销账数量 / r_费用.数量), n_Dec);
      
        Insert Into 费用变动记录
          (ID, 记录状态, 病人id, 主页id, 变动时间, 原变动id, 目标变动id, 原病区id, 目标病区id, 费用id, NO, 收费类别, 收费细目id, 医嘱序号, 数量, 单价, 应收金额, 实收金额,
           状态, 摘要, 操作员编号, 操作员姓名)
        Values
          (费用变动记录_Id.Nextval, Decode(Nvl(n_操作类型, 0), 0, 1, 2), n_病人id, n_主页id, d_登记时间, n_原变动id, n_目标变动id, n_原病区id,
           n_目标病区id, r_费用.费用id, r_费用.No, r_费用.收费类别, r_费用.收费细目id, r_费用.医嘱id, n_申请销账数量, r_费用.标准单价, n_应收金额, n_实收金额, 2,
           Decode(Nvl(n_操作类型, 0), 0, '病区变动', '病区变动撤销') || '产生的销帐申请', v_操作员编号, v_操作员姓名);
      End If;
    
      --2.其他收费项目(药品未包含)
    Else
      --处理规则:
      --1.对原始记录进行销帐
      --2.新增目标病区数据
      --3.如果是划价单，直接更改原记录病区id和执行部门
      If Nvl(r_费用.记录状态, 0) = 0 Then
        --直接修改(包含病人病区及目标病区)
        Update 住院费用记录
        Set 病人病区id = n_目标病区id, 执行部门id = n_目标病区id
        Where NO = r_费用.No And 记录性质 = 2 And 记录状态 = 0 And Nvl(价格父号, 序号) = r_费用.序号;
      
        Insert Into 费用变动记录
          (ID, 记录状态, 病人id, 主页id, 变动时间, 原变动id, 目标变动id, 原病区id, 目标病区id, 费用id, NO, 收费类别, 收费细目id, 医嘱序号, 数量, 单价, 应收金额, 实收金额,
           状态, 摘要, 操作员编号, 操作员姓名)
          Select 费用变动记录_Id.Nextval, Decode(Nvl(n_操作类型, 0), 0, 1, 2), n_病人id, n_主页id, d_登记时间, n_原变动id, n_目标变动id, n_原病区id,
                 n_目标病区id, 费用id, NO, 收费类别, 收费细目id, 医嘱序号, 数量, 标准单价, 应收金额, 实收金额, 0,
                 Decode(Nvl(n_操作类型, 0), 0, '病区变动', '病区变动撤销') || '修改记帐划价单', v_操作员编号, v_操作员姓名
          From (Select Max(Decode(价格父号, Null, ID, 0)) As 费用id, NO, 收费类别, 收费细目id, 医嘱序号, Avg(Nvl(付数, 1) * 数次) As 数量,
                        Sum(标准单价) As 标准单价, Sum(应收金额) As 应收金额, Sum(实收金额) As 实收金额
                 From 住院费用记录
                 Where NO = r_费用.No And 记录性质 = 2 And 记录状态 = 0 And Nvl(价格父号, 序号) = r_费用.序号
                 Group By NO, 收费类别, 收费细目id, 医嘱序号, Nvl(价格父号, 序号));
      
      Elsif Nvl(n_转出数量, 0) > 0 Then
        --直接销帐处理
        --序号：序号1:数量1:执行状态1,序号2:数量2:执行状态2,...序号n:数量n:执行状态n  如:"1:2:1,2:10:1,3:2:1"
        --1.先产生销帐记录
        Zl_住院记帐记录_Delete_s(r_费用.No, r_费用.序号 || ':' || n_转出数量 || ':0', v_操作员编号, v_操作员姓名, 2, 2, d_登记时间);
        --2.目标病区转入记录
        For c_明细 In (Select 病人费用记录_Id.Nextval As 费用id, NO, 记录性质, 1 As 记录状态, n_Max序号 + Rownum As 序号, 从属父号,
                            价格父号 + (n_Max序号 + Rownum - 序号) As 价格父号, 主页id, 病人id, 医嘱序号, 门诊标志, 多病人单, 婴儿费, 姓名, 性别, 年龄, 标识号,
                            床号, 费别, 病人科室id, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, -1 * 数次 As 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用,
                            标准单价, -1 * 应收金额 As 应收金额, -1 * 实收金额 As 实收金额, 开单部门id, 开单人, 划价人, 执行人, r_费用.执行状态 As 执行状态, 执行时间,
                            发生时间, 保险项目否, 保险大类id, -1 * 统筹金额 As 统筹金额, 保险编码, 记帐单id, 摘要, 费用类型, 是否急诊, 结论, 医疗小组id
                     From 住院费用记录
                     Where NO = r_费用.No And Nvl(价格父号, 序号) = r_费用.序号 And 记录状态 = 2 And 登记时间 = d_登记时间) Loop
        
          Insert Into 住院费用记录
            (ID, NO, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 主页id, 病人id, 医嘱序号, 门诊标志, 多病人单, 婴儿费, 姓名, 性别, 年龄, 标识号, 床号, 费别, 病人病区id,
             病人科室id, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人,
             执行部门id, 划价人, 执行人, 执行状态, 执行时间, 操作员编号, 操作员姓名, 发生时间, 登记时间, 保险项目否, 保险大类id, 统筹金额, 保险编码, 记帐单id, 摘要, 费用类型, 是否急诊,
             结论, 医疗小组id)
          Values
            (c_明细.费用id, c_明细.No, c_明细.记录性质, c_明细.记录状态, c_明细.序号, c_明细.从属父号, c_明细.价格父号, c_明细.主页id, c_明细.病人id, c_明细.医嘱序号,
             c_明细.门诊标志, c_明细.多病人单, c_明细.婴儿费, c_明细.姓名, c_明细.性别, c_明细.年龄, c_明细.标识号, c_明细.床号, c_明细.费别, n_目标病区id,
             c_明细.病人科室id, c_明细.收费类别, c_明细.收费细目id, c_明细.计算单位, c_明细.付数, c_明细.发药窗口, c_明细.数次, c_明细.加班标志, c_明细.附加标志,
             c_明细.收入项目id, c_明细.收据费目, c_明细.记帐费用, c_明细.标准单价, c_明细.应收金额, c_明细.实收金额, c_明细.开单部门id, c_明细.开单人, n_目标病区id,
             c_明细.划价人, c_明细.执行人, /*c_明细.执行状态*/ 0, c_明细.执行时间, v_操作员编号, v_操作员姓名, c_明细.发生时间, d_登记时间, c_明细.保险项目否,
             c_明细.保险大类id, c_明细.统筹金额, c_明细.保险编码, c_明细.记帐单id, c_明细.摘要, c_明细.费用类型, c_明细.是否急诊, c_明细.结论, c_明细.医疗小组id);
        
          Update 费用变动记录
          Set 单价 = Nvl(单价, 0) + Nvl(c_明细.标准单价, 0), 应收金额 = Nvl(应收金额, 0) + Nvl(c_明细.应收金额, 0),
              实收金额 = Nvl(实收金额, 0) + Nvl(c_明细.实收金额, 0)
          Where 费用id = r_费用.费用id And 变动时间 = d_登记时间 And 目标变动id = n_目标变动id And 收费细目id = r_费用.收费细目id And
                病人id + 0 = c_明细.病人id;
          If Sql%NotFound Then
            Insert Into 费用变动记录
              (ID, 记录状态, 病人id, 主页id, 变动时间, 原变动id, 目标变动id, 原病区id, 目标病区id, 费用id, NO, 收费类别, 收费细目id, 医嘱序号, 数量, 单价, 应收金额,
               实收金额, 状态, 摘要, 操作员编号, 操作员姓名)
            Values
              (费用变动记录_Id.Nextval, Decode(Nvl(n_操作类型, 0), 0, 1, 2), n_病人id, n_主页id, d_登记时间, n_原变动id, n_目标变动id, n_原病区id,
               n_目标病区id, r_费用.费用id, r_费用.No, r_费用.收费类别, r_费用.收费细目id, r_费用.医嘱id, Round(c_明细.数次 * Nvl(c_明细.付数, 1), 5),
               c_明细.标准单价, c_明细.应收金额, c_明细.实收金额, 1, Decode(Nvl(n_操作类型, 0), 0, '病区变动', '病区变动撤销') || '修改记帐单', v_操作员编号,
               v_操作员姓名);
          End If;
        
          Update 病人审批项目
          Set 已用数量 = Nvl(已用数量, 0) + Round(c_明细.数次 * Nvl(c_明细.付数, 1), 5)
          Where 病人id = n_病人id And 主页id = n_主页id And 项目id = c_明细.收费细目id And Nvl(使用限量, 0) <> 0;
        
          --病人余额
          Update 病人余额
          Set 费用余额 = Nvl(费用余额, 0) + c_明细.实收金额
          Where 病人id = c_明细.病人id And 类型 = 2 And 性质 = 1;
          If Sql%RowCount = 0 Then
            Insert Into 病人余额
              (病人id, 类型, 性质, 费用余额, 预交余额)
            Values
              (c_明细.病人id, 2, 1, c_明细.实收金额, 0);
          End If;
          --病人未结费用
          Update 病人未结费用
          Set 金额 = Nvl(金额, 0) + c_明细.实收金额
          Where 病人id = c_明细.病人id And Nvl(主页id, 0) = Nvl(c_明细.主页id, 0) And Nvl(病人病区id, 0) = Nvl(n_目标病区id, 0) And
                Nvl(病人科室id, 0) = Nvl(c_明细.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(c_明细.开单部门id, 0) And
                Nvl(执行部门id, 0) = Nvl(n_目标病区id, 0) And 收入项目id + 0 = c_明细.收入项目id And 来源途径 + 0 = 2;
          If Sql%RowCount = 0 Then
            Insert Into 病人未结费用
              (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
            Values
              (c_明细.病人id, c_明细.主页id, n_目标病区id, c_明细.病人科室id, c_明细.开单部门id, n_目标病区id, c_明细.收入项目id, 2, c_明细.实收金额);
          End If;
        
          n_Max序号 := c_明细.序号;
        End Loop;
      End If;
    End If;
  End Loop;

  If Nvl(n_操作类型, 0) = 1 Then
    --撤消操作,需要删除卫生材料部分中未审核部分
    For c_销账 In (Select a.费用id, a.病人id, a.主页id, a.No, a.收费类别, a.收费细目id, a.医嘱序号, a.数量, a.单价, a.应收金额, a.实收金额
                 From 费用变动记录 A
                 Where a.原变动id = n_目标变动id And a.目标变动id = n_原变动id And a.状态 = 2) Loop
    
      Select Sum(数量) Into n_申请取消数量 From 病人费用销帐 Where 费用id = c_销账.费用id And 状态 In (0, 2);
    
      If Nvl(n_申请取消数量, 0) > 0 Then
        n_应收金额 := Round(n_申请取消数量 * Nvl(c_销账.单价, 0), n_Dec);
        n_实收金额 := 0;
        If Nvl(c_销账.应收金额, 0) <> 0 Then
          n_实收金额 := Round(Nvl(n_应收金额, 0) * Nvl(c_销账.实收金额, 0) / c_销账.应收金额, n_Dec);
        End If;
      
        Insert Into 费用变动记录
          (ID, 记录状态, 病人id, 主页id, 变动时间, 原变动id, 目标变动id, 原病区id, 目标病区id, 费用id, NO, 收费类别, 收费细目id, 医嘱序号, 数量, 单价, 应收金额, 实收金额,
           状态, 摘要, 操作员编号, 操作员姓名)
        Values
          (费用变动记录_Id.Nextval, 2, c_销账.病人id, c_销账.主页id, d_登记时间, n_原变动id, n_目标变动id, n_原病区id, n_目标病区id, c_销账.费用id, c_销账.No,
           c_销账.收费类别, c_销账.收费细目id, c_销账.医嘱序号, n_申请取消数量, c_销账.单价, n_应收金额, n_实收金额, 3, '病区撤销后删除销帐申请', v_操作员编号, v_操作员姓名);
      End If;
    
      Delete 病人费用销帐 Where 费用id = c_销账.费用id And 状态 In (0, 2);
    End Loop;
  End If;
  Json_Out := zlJsonOut('成功', 1);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Executeturnwardfee;
/

Create Or Replace Procedure Zl_Exsesvr_Getfeechangerec
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:获取费用变动记录
  --入参：Json_In:格式
  -- input
  --   pati_id           N   1 病人ID
  --   pati_pageid       N   1 主页ID
  --出参: Json_Out,格式如下
  --  output
  --    code               C  1 应答码：0-失败；1-成功
  --    message            C  1 应答消息：失败时返回具体的错误信息
  --    item_list[]
  --      bill_no          C   费用单据号
  --      item_id          N   收费细目ID
  --      item_name        C   收费细目名称
  --      ward_id_old      N   原病区id
  --      ward_name_old    N   原病区名称
  --      ward_id_new      N   目标病区id
  --      ward_name_new    N   目标病区名称
  --      quantity         N   数量
  --      price            N   单价
  --      fee_ampaid       N   实收金额
  --      rec_type         N   记录类型:0-直接更改原单据；1-产生的正常转移数据；2-产生的销帐申请变动；3-取消销帐申请变动
  --      change_time      C   变动时间:yyyy-mm-dd hh24:mi:ss
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_Output Varchar2(32767);
  c_Output Clob;

  n_病人id 费用变动记录.病人id%Type;
  n_主页id 费用变动记录.主页id%Type;

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_病人id := j_Json.Get_Number('pati_id');
  n_主页id := j_Json.Get_Number('pati_pageid');

  For r_变动 In (Select a.No, a.收费细目id, Decode(j.是否保密, 1, '***', b.名称) As 项目名称, a.原病区id, c.名称 As 原病区, a.目标病区id,
                      d.名称 As 目标病区, a.数量, a.单价, a.实收金额, a.状态, To_Char(a.变动时间, 'yyyy-mm-dd hh24:mi:ss') As 变动时间
               From 费用变动记录 A, 收费项目目录 B, 部门表 C, 部门表 D, 住院费用记录 J
               Where a.收费细目id = b.Id And a.原病区id = c.Id And a.目标病区id = d.Id And a.病人id = n_病人id And a.主页id = n_主页id And
                     a.费用id = j.Id) Loop
  
    If Length(Nvl(v_Output, ' ')) > 30000 Then
      If c_Output Is Not Null Then
        c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
      Else
        c_Output := To_Clob(v_Output);
      End If;
    
      v_Output := Null;
    End If;
  
    zlJsonPutValue(v_Output, 'bill_no', r_变动.No, 0, 1);
    zlJsonPutValue(v_Output, 'item_id', r_变动.收费细目id, 1);
    zlJsonPutValue(v_Output, 'item_name', r_变动.项目名称);
    zlJsonPutValue(v_Output, 'ward_id_old', r_变动.原病区id, 1);
    zlJsonPutValue(v_Output, 'ward_name_old', r_变动.原病区);
    zlJsonPutValue(v_Output, 'ward_id_new', r_变动.目标病区id, 1);
    zlJsonPutValue(v_Output, 'ward_name_new', r_变动.目标病区);
    zlJsonPutValue(v_Output, 'quantity', r_变动.数量, 1);
    zlJsonPutValue(v_Output, 'price', r_变动.单价, 1);
    zlJsonPutValue(v_Output, 'fee_ampaid', r_变动.实收金额, 1);
    zlJsonPutValue(v_Output, 'rec_type', Nvl(r_变动.状态, 0), 1);
    zlJsonPutValue(v_Output, 'change_time', r_变动.变动时间, 1, 2);
  
  End Loop;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"成功","item_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || v_Output || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getfeechangerec;
/


Create Or Replace Procedure Zl_Exsesvr_Getturnwardfee
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:获取转病区费用
  --入参：Json_In:格式
  -- input
  --   pati_id           N   1 病人ID
  --   pati_pageid       N   1 主页ID
  --   exe_deptid        N   1 执行部门ID
  --出参: Json_Out,格式如下
  --  output
  --    code               C  1 应答码：0-失败；1-成功
  --    message            C  1 应答消息：失败时返回具体的错误信息
  --    item_list[]
  --      rec_type         N   记录类型:1-费用转移记录；2-销帐申请记录
  --      bill_prop        N   单据性质:0-记帐划价单,1-记帐单
  --      bill_no          C   费用单据号
  --      serial_num       N   费用单据序号
  --      item_id          N   收费细目ID
  --      item_name        C   收费细目名称
  --      advice_id        N   医嘱序号
  --      quantity         N   数量
  --      price            N   单价
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_病人id     住院费用记录.病人id%Type;
  n_主页id     住院费用记录.主页id%Type;
  n_执行部门id 住院费用记录.执行部门id%Type;

  v_Output Varchar2(32767);
  c_Output Clob;
Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_病人id     := j_Json.Get_Number('pati_id');
  n_主页id     := j_Json.Get_Number('pati_pageid');
  n_执行部门id := j_Json.Get_Number('exe_deptid');
  For r_变动 In (Select a.记录类型, Decode(a.记录状态, 0, 0, 1) As 单据性质, a.No, a.序号, a.收费细目id,
                      Max(Decode(a.是否保密, 1, '***', b.名称)) As 收费项目, a.医嘱序号, Sum(a.剩余数量) As 剩余数量, Max(a.标准单价) As 标准单价
               From (
                    
                    With 住院费用 As (Select a.No, a.序号, Max(a.收费细目id) As 收费细目id, Sum(数量) As 剩余数量,
                                         Max(Decode(a.记录状态, 2, 0, a.费用id)) As 费用id, Max(a.记录状态) As 记录状态,
                                         Max(a.医嘱序号) As 医嘱序号, Max(a.是否保密) As 是否保密, Max(a.标准单价) As 标准单价
                                  From (Select a.No, 记录状态, Nvl(a.价格父号, 序号) As 序号, 收费细目id, Avg(Nvl(a.付数, 1) * a.数次) As 数量,
                                                Max(Decode(a.价格父号, Null, a.Id, 0)) As 费用id, Max(a.医嘱序号) As 医嘱序号,
                                                Max(a.是否保密) As 是否保密, Sum(a.标准单价) As 标准单价
                                         From 住院费用记录 A, 材料特性 C
                                         Where a.记录性质 = 2 And a.执行部门id = n_执行部门id And a.医嘱序号 Is Not Null And
                                               Nvl(a.是否附费, 0) = 0 And a.病人id = n_病人id And a.主页id = n_主页id And
                                               Instr(',5,6,7,', ',' || a.收费类别 || ',') = 0 And a.收费细目id = c.材料id And
                                               Nvl(c.跟踪在用, 0) = 1
                                         Group By a.No, 记录状态, a.医嘱序号, Nvl(a.价格父号, 序号), 收费细目id, a.执行状态) A
                                  Group By a.No, a.序号)
                    --需要销账申请的卫材,减去已申请数量
                      Select 2 As 记录类型, a.No, a.序号, Max(a.记录状态) As 记录状态, a.收费细目id, Max(a.医嘱序号) As 医嘱序号,
                             Nvl(Sum(a.剩余数量), 0) - Nvl(Sum(b.数量), 0) As 剩余数量, Sum(a.标准单价) As 标准单价, Max(a.是否保密) As 是否保密
                      From 住院费用 A,
                           (Select b.费用id, Nvl(Sum(b.数量), 0) As 数量
                             From 住院费用 A, 病人费用销帐 B
                             Where a.费用id = b.费用id And Nvl(b.状态, 0) = 0
                             Group By b.费用id
                             Having Nvl(Sum(b.数量), 0) <> 0) B
                      Where a.费用id = b.费用id(+)
                      Group By a.No, a.序号, a.收费细目id
                      Having Nvl(Sum(a.剩余数量), 0) - Nvl(Sum(b.数量), 0) <> 0
                      Union All
                      Select 1 As 记录类型, a.No, Nvl(a.价格父号, a.序号) As 序号, a.记录状态, a.收费细目id, a.医嘱序号,
                             Avg(Nvl(a.付数, 1) * a.数次) As 剩余数量, Sum(a.标准单价) As 标准单价, Max(a.是否保密) As 是否保密
                      From 住院费用记录 A, 材料特性 C
                      Where a.收费细目id = c.材料id(+) And a.记录性质 = 2 And a.执行部门id = n_执行部门id And a.医嘱序号 Is Not Null And
                            Nvl(a.是否附费, 0) = 0 And a.病人id = n_病人id And a.主页id = n_主页id And
                            Instr(',5,6,7,', ',' || a.收费类别 || ',') = 0 And Nvl(c.跟踪在用, 0) = 0
                      Group By a.No, a.记录状态, a.医嘱序号, Nvl(a.价格父号, a.序号), a.收费细目id
                      
                      ) A, 收费项目目录 B
                      Where a.收费细目id = b.Id
                      Group By a.记录类型, Decode(a.记录状态, 0, 0, 1), a.No, a.序号, a.收费细目id, a.医嘱序号
               ) Loop
  
    If Length(Nvl(v_Output, ' ')) > 30000 Then
      If c_Output Is Not Null Then
        c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
      Else
        c_Output := To_Clob(v_Output);
      End If;
    
      v_Output := Null;
    End If;
  
    zlJsonPutValue(v_Output, 'rec_type', r_变动.记录类型, 1, 1);
    zlJsonPutValue(v_Output, 'bill_prop', r_变动.单据性质, 1);
    zlJsonPutValue(v_Output, 'bill_no', r_变动.No);
    zlJsonPutValue(v_Output, 'serial_num', r_变动.序号, 1);
    zlJsonPutValue(v_Output, 'item_id', r_变动.收费细目id, 1);
    zlJsonPutValue(v_Output, 'item_name', r_变动.收费项目);
    zlJsonPutValue(v_Output, 'advice_id', r_变动.医嘱序号, 1);
    zlJsonPutValue(v_Output, 'quantity', r_变动.剩余数量, 1);
    zlJsonPutValue(v_Output, 'price', r_变动.标准单价, 1, 2);
  
  End Loop;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"成功","item_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"成功","item_list":[' || v_Output || ']}}';
  End If;

Exception
  When Others Then
     Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getturnwardfee;
/


Create Or Replace Procedure Zl_Exsesvr_Getbillgrpbyfeetype
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：按收费类别分组获取费用单据信息
  --入参：Json_In:格式
  --  input
  --     query_type   N 1 查询方式：0-按单据号,1-按医嘱ID
  --     bill_nos     C 0 单据号，允许传入多个,用逗号分隔,如:A00001,A0002,...,A000n
  --     advice_ids   C 0 医嘱ID，允许传入多个,用逗号分隔,如:1,2,3,4
  --     pati_id      N 0 病人ID，记帐表时按病人获取
  --出参: Json_Out,格式如下
  --  output
  --    code                N 1 应答吗：0-失败；1-成功
  --    message             C 1 应答消息：失败时返回具体的错误信息
  --    item_list[]
  --      pati_id         N 1 病人ID
  --      pati_pageid     N 1 主页ID
  --      pati_name       C 1 病人姓名
  --      fee_type        C 1 费用类别
  --      fee_type_name   C 1 费用类别名称
  --      ward_id         N 1 病区ID
  --      fee_ampaid      N 1 实收金额合计
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_查询方式 Number;
  v_Nos      Varchar2(32767);
  v_医嘱ids  Varchar2(32767);
  n_病人id   门诊费用记录.病人id%Type;

  n_Firstitem Number(1);
  v_Temp      Varchar2(32767);
  c_Temp      Clob;
Begin
  --解析入参

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_查询方式 := j_Json.Get_Number('query_type');

  n_Firstitem := 1;
  v_Temp      := '{"output":{"code":1,"message":"成功","item_list":[';
  If Nvl(n_查询方式, 0) = 0 Then
    v_Nos    := j_Json.Get_String('bill_nos');
    n_病人id := j_Json.Get_Number('pati_id');
  
    For r_费用 In (Select m.类别, j.类别 As 名称, m.病区id, Sum(m.实收金额) As 金额, m.病人id, m.主页id, m.姓名
                 From (Select /*+cardinality(b,10)*/
                         a.病人id, 0 As 主页id, a.姓名, a.收费类别 As 类别, 0 As 病区id, Nvl(Sum(a.实收金额), 0) As 实收金额
                        From 门诊费用记录 A, Table(f_Str2List(v_Nos)) B
                        Where a.No = b.Column_Value And 记帐费用 = 1 And 记录状态 = 0 And 记录性质 = 2 And
                              (Nvl(n_病人id, 0) = 0 Or a.病人id = n_病人id)
                        Group By a.收费类别, a.病人id, a.姓名
                        Union All
                        Select /*+cardinality(b,10)*/
                         a.病人id, a.主页id, a.姓名, a.收费类别 As 类别, a.病人病区id As 病区id, Nvl(Sum(a.实收金额), 0) As 实收金额
                        From 住院费用记录 A, Table(f_Str2List(v_Nos)) B
                        Where a.No = b.Column_Value And 记帐费用 = 1 And 记录状态 = 0 And 记录性质 = 2 And
                              (Nvl(n_病人id, 0) = 0 Or a.病人id = n_病人id)
                        Group By a.收费类别, a.病人id, a.主页id, a.姓名, a.病人病区id) M, 收费类别 J
                 Where m.类别 = j.编码
                 Group By m.病人id, m.主页id, m.姓名, m.类别, j.类别, m.病区id) Loop
    
      If Nvl(n_Firstitem, 0) = 0 Then
        v_Temp := v_Temp || ',';
      Else
        n_Firstitem := 0;
      End If;
      v_Temp := v_Temp || '{';
      v_Temp := v_Temp || '"pati_id":' || r_费用.病人id;
      v_Temp := v_Temp || ',"pati_pageid":' || Nvl(r_费用.主页id, 0);
      v_Temp := v_Temp || ',"pati_name":"' || zlJsonStr(r_费用.姓名) || '"';
      v_Temp := v_Temp || ',"fee_type":"' || r_费用.类别 || '"';
      v_Temp := v_Temp || ',"fee_type_name":"' || r_费用.名称 || '"';
      v_Temp := v_Temp || ',"ward_id":' || Nvl(r_费用.病区id, 0);
      v_Temp := v_Temp || ',"fee_ampaid":' || zlJsonStr(r_费用.金额, 1);
      v_Temp := v_Temp || '}';
    
      If Length(v_Temp) > 20000 Then
        c_Temp := c_Temp || To_Clob(v_Temp);
        v_Temp := '';
      End If;
    End Loop;
  
  Else
    v_医嘱ids := j_Json.Get_String('advice_ids');
  
    For r_费用 In (Select m.类别, j.类别 As 名称, m.病区id, Sum(m.实收金额) As 金额, m.病人id, m.主页id, m.姓名
                 From (Select /*+cardinality(b,10)*/
                         a.病人id, 0 As 主页id, a.姓名, a.收费类别 As 类别, 0 As 病区id, Nvl(Sum(a.实收金额), 0) As 实收金额
                        From 门诊费用记录 A, Table(f_Num2List(v_医嘱ids)) B
                        Where a.医嘱序号 = b.Column_Value And 记帐费用 = 1 And 记录状态 = 0
                        Group By a.收费类别, a.病人id, a.姓名
                        Union All
                        Select /*+cardinality(b,10)*/
                         a.病人id, a.主页id, a.姓名, a.收费类别 As 类别, a.病人病区id As 病区id, Nvl(Sum(实收金额), 0) As 实收金额
                        From 住院费用记录 A, Table(f_Num2List(v_医嘱ids)) B
                        Where a.医嘱序号 = b.Column_Value And 记帐费用 = 1 And 记录状态 = 0
                        Group By a.收费类别, a.病人id, a.主页id, a.姓名, a.病人病区id) M, 收费类别 J
                 Where m.类别 = j.编码
                 Group By m.病人id, m.主页id, m.姓名, m.类别, j.类别, m.病区id) Loop
    
      If Nvl(n_Firstitem, 0) = 0 Then
        v_Temp := v_Temp || ',';
      Else
        n_Firstitem := 0;
      End If;
      v_Temp := v_Temp || '{';
      v_Temp := v_Temp || '"pati_id":' || r_费用.病人id;
      v_Temp := v_Temp || ',"pati_pageid":' || Nvl(r_费用.主页id, 0);
      v_Temp := v_Temp || ',"pati_name":"' || zlJsonStr(r_费用.姓名) || '"';
      v_Temp := v_Temp || ',"fee_type":"' || r_费用.类别 || '"';
      v_Temp := v_Temp || ',"fee_type_name":"' || r_费用.名称 || '"';
      v_Temp := v_Temp || ',"ward_id":' || Nvl(r_费用.病区id, 0);
      v_Temp := v_Temp || ',"fee_ampaid":' || zlJsonStr(r_费用.金额, 1);
      v_Temp := v_Temp || '}';
    
      If Length(v_Temp) > 30000 Then
        c_Temp := c_Temp || To_Clob(v_Temp);
        v_Temp := '';
      End If;
    End Loop;
  End If;
  v_Temp := v_Temp || ']}}';

  If c_Temp Is Not Null Then
    Json_Out := c_Temp || To_Clob(v_Temp);
  Else
    Json_Out := v_Temp;
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getbillgrpbyfeetype;
/

Create Or Replace Procedure Zl_Exsesvr_Getchargeoffapply
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：根据病人id和主页id,获取病人的销帐申请信息
  --入参：Json_In:格式
  --input
  -- pati_id N 1 病人id
  -- pati_pageid N 1 主页id
  --出参: Json_Out,格式如下
  --output
  --  code  C 1 应答码：0-失败；1-成功
  --  message C 1 应答消息：成功时返回成功信息，失败时返回具体的错误信息
  --  exists  N 1 是否存在:1-存在;0-不存在
  --    apply_list[]  C   申请列表
  --    fee_no  C 1 费用单据号
  --    fitem_name  C 1 收费项目名称
  --    audit_dept_name C 1 审核部分名称

  ---------------------------------------------------------------------------

  n_病人id Number(18);
  n_主页id Number(18);

  n_Count Number;
  j_Input PLJson;
  j_Json  PLJson;

  v_Output  Varchar2(32767);
  n_Isexist Number(1);
Begin

  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_病人id := j_Json.Get_Number('pati_id');
  n_主页id := j_Json.Get_Number('pati_pageid');

  n_Isexist := 0;
  For c_申请 In (Select Distinct a.No, d.名称 项目名称, c.名称 审核科室
               From 住院费用记录 A, 病人费用销帐 B, 部门表 C, 收费项目目录 D
               Where a.病人id = n_病人id And a.主页id = n_主页id And a.Id = b.费用id And b.状态 = 0 And b.审核部门id = c.Id And
                     b.收费细目id = d.Id
               Order By a.No, c.名称
               
               ) Loop
  
    zlJsonPutValue(v_Output, 'fee_no', c_申请.No, 0, 1);
    zlJsonPutValue(v_Output, 'fitem_name', c_申请.项目名称);
    zlJsonPutValue(v_Output, 'audit_dept_name', c_申请.审核科室, 0, 2);
  
    n_Isexist := 1;
  End Loop;

  If n_Count > 0 Then
    n_Isexist := 1;
  Else
    n_Isexist := 0;
  End If;
  Json_Out := '{"output":{"code":1,"message":"成功","exists":' || n_Isexist || ',"apply_list":[' || v_Output || ']}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getchargeoffapply;
/


Create Or Replace Procedure Zl_Exsesvr_Existspricebill
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：根据病人id或医嘱id判断是否存在对应的划价单
  --入参：Json_In:格式
  --input      
  -- pati_id N 1 病人id
  -- pati_pageid N 1 主页id
  -- advice_ids  C   医嘱id:多个用逗号
  -- billtype  N 1 单据类型:1-收费划价单;2-记帐划价单

  --出参: Json_Out,格式如下
  --  output
  --       code             N 1 应答吗：0-失败；1-成功
  --       message          C 1 应答消息：失败时返回具体的错误信息
  --     exists N 1 是否存在:1-存在;0-不存在
  ---------------------------------------------------------------------------
  j_Input    PLJson;
  j_Json     PLJson;
  n_病人id   Number(18);
  n_主页id   Number(18);
  v_医嘱ids  Varchar2(32767);
  n_单据类型 Number(2);
  n_Count    Number(5);
Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_病人id   := Nvl(j_Json.Get_Number('pati_id'), 0);
  n_主页id   := Nvl(j_Json.Get_Number('pati_pageid'), 0);
  v_医嘱ids  := j_Json.Get_String('advice_ids');
  n_单据类型 := Nvl(j_Json.Get_Number('billtype'), 1);

  If v_医嘱ids Is Null Then
    If n_单据类型 = 1 Then
      Select Max(1)
      Into n_Count
      From 门诊费用记录
      Where 记录性质 = 1 And (记录状态 = 0 Or 记录状态 = 1 And 结帐id Is Null) And 病人id = n_病人id And Rownum < 2;
    
    Else
      If Nvl(n_主页id, 0) = 0 Then
        Select Max(1)
        Into n_Count
        From (Select 1
               From 门诊费用记录
               Where 记录状态 = 0 And Nvl(记帐费用, 0) = 1 And 病人id = n_病人id And Rownum < 2
               Union All
               Select 1
               From 住院费用记录
               Where 记录状态 = 0 And Nvl(记帐费用, 0) = 1 And 门诊标志 <> 2 And 病人id = n_病人id And Rownum < 2);
      Else
        Select 1
        Into n_Count
        From 住院费用记录
        Where 记录状态 = 0 And Nvl(记帐费用, 0) = 1 And 病人id = n_病人id And 主页id = n_主页id And Rownum < 2;
      End If;
    End If;
  
  Else
  
    Select Max(1)
    Into n_Count
    From (With 医嘱数据 As (Select Column_Value As 医嘱id From Table(f_Num2List(v_医嘱ids)))
           Select /*+cardinality(B,10) */
            1
           From 门诊费用记录 A, 医嘱数据 B
           Where a.医嘱序号 = b.医嘱id And a.记录状态 = 0 And Nvl(a.记帐费用, 0) = 1 And Rownum < 2
           Union All
           Select /*+cardinality(B,10) */
            1
           From 住院费用记录 A, 医嘱数据 B
           Where a.医嘱序号 = b.医嘱id And a.记录状态 = 0 And Nvl(a.记帐费用, 0) = 1 And Rownum < 2);
  
  
  End If;

  Json_Out := '{"output":{"code":1,"message":"成功","exists":' || n_Count || '}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Existspricebill;
/
CREATE OR REPLACE Procedure Zl_Exsesvr_Getdrugerrdata
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：根据病人ID和医嘱信息返回病人费用信息
  --入参：Json_In:格式。当值为null时获取所有病人的异常信息
  --  input
  --    pati_list[]病人列表
  --       pati_id                    N 1 病人id
  --       bill_list[]                费用单据号列表，可以不传，不传时表示获取费用域同步异常的数据
  --         fee_source               N 0 费用来源：1-门诊；2-住院
  --         fee_billtype             N 0 费用单据类型：1-收费处方；2-记帐单处方
  --         fee_no                   C 0 费用单据号
  --出参: Json_Out,格式如下
  --  output
  --    code                          N   1 应答吗：0-失败；1-成功
  --    message                       C   1 应答消息：失败时返回具体的错误信息
  --    pati_bill_list[]
  --       billtype                   N   1 单据类型: 1 -收费处方  ;2- 记帐单处方;3- 记帐表处方
  --       pati_source                N   1 病人来源:1-门诊;2-住院;4-体检
  ---------------------------billtype = 3,记帐表处方（药品收发记录.单据=10）时，无以下节点--------------------------------------
  --       pati_id                    N   1 病人ID
  --       pati_pageid                N   1 主页ID
  --       pati_name                  C   1 病人姓名
  --       pati_sex_code              C   1 性别编号（新门诊)
  --       pati_sex                   C   1 性别
  --       pati_age                   C   1 年龄
  --       pati_deptid                N   1 病人科室ID
  --       pati_wardarea_id           N     病人病区ID
  ---------------------------billtype = 3,记帐表处方（药品收发记录.单据=10）时，无以上节点-----------------------------------------
  --       bill_list[]                      更新数据列表[数组]
  --         fee_source                N  0 费用来源
  --         rcp_no                    C  1 NO
  --         recipe_type               N  0 处方类型:0和空-普通,1-儿科,2-急诊,3-精二,4-精一,5-麻醉
  --         charge_tag                N  1 收费标志:0-未收费或记帐划价;1-已收费或记帐
  --         fee_acnter                C  0 划价人
  --         recipe_plcdept_id         C  0 开单科室id（新门诊)
  --         recipe_plcdept            C  0 开单科室名称（新门诊)
  --         recipe_placer_id          C  0 开单医师id（新门诊)
  --         recipe_placer             C  0 开单医师（新门诊) 增加
  --         operator_name             C  1 操作员姓名
  --         operator_code             C  1 操作员编号
  --         create_time               C  1 登记时间:yyyy-mm-dd hh:mi:ss
  --         item_list[]                    更新数据列表[数组]

  ---------------------------billtype = 3,记帐表处方（药品收发记录.单据=10）时，有以下节点----------------------------------------
  --           pati_id                 N  1 病人ID
  --           pati_pageid             N  0 主页ID
  --           pati_name               C  1 病人姓名
  --           pati_sex_code           C  1 性别编号（新门诊)
  --           pati_sex                C  1 性别
  --           pati_age                C  1 年龄
  --           pati_wardarea_id        N  0 病人病区ID
  --           pati_deptid             N  1 病人科室ID
  ---------------------------billtype = 3,记帐表处方（药品收发记录.单据=10）时，有以上节点-----------------------------------------

  --           rcpdtl_id               N  1 处方明细ID
  --           serial_num              N  1 序号:(变更(包括存储)：序号和组号，1、2、3、3、3、4…)
  --           pharmacy_id             N  1 药房ID
  --           pharmacy_name           C  1 药房名称(新门诊)
  --           takedept_id             N  1 领药部门ID:针对住院才传入
  --           drug_id                 N  1 药品ID
  --           baby_num                N  0  婴儿序号
  --           advice_id               N  0 医嘱ID
  --           decoction_method        C  0 煎法
  --           use_mode                N  0 取药特性：0-正常方式，1-离院带药，2-自取药
  --           packages_num            N  1 发药付数
  --           send_num                N  1 发药数量
  --           send_unit               C  1 发药单位：zlhis零售单位
  --           price                   N  0 售价
  --           money                   N  0 零售金额(新门诊)
  --           pharmacy_window         C  0 发药窗口
  --           memo                    C  0 摘要
  ------------------------------------------------------------------------------------------------------------
  j_Json      PLJson;
  j_Json_In   PLJson;
  j_Pati_List Pljson_List;
  j_Json_Out  PLJson;
  j_Bill_List Pljson_List;

  Json_Temp_Out Clob;
  c_Jtmp        Clob;

  j_Item   PLJson;
  n_病人id Number(18);

  v_Json Varchar2(4000);
  n_Code Number;

  n_费用来源 Number(1);
  n_记录性质 门诊费用记录.记录性质%Type;
  v_No       门诊费用记录.No%Type;

  l_Outnos t_StrList2 := t_StrList2();
  l_Innos  t_StrList2 := t_StrList2();
Begin
  If Json_In Is Null Then
    --费用系统中记费同步异常的数据
    For r_Fee In (Select Distinct 1 As 费用来源, Decode(Mod(a.记录性质, 10), 2, 2, 1) As 单据类型, a.No
                  From 门诊费用记录 A, 病人费用异常记录 B
                  Where a.收费类别 In ('5', '6', '7') And a.记录性质 In (1, 2) And
                        a.id = b.费用id And (b.产生环节 = 0 Or b.产生环节 = 1) And Nvl(b.同步标志, 0) = 1 And
                        Exists (Select 1
                         From 门诊费用记录
                         Where 记录性质 = a.记录性质 And NO = a.No And 序号 = a.序号
                         Group By 记录性质, NO, 序号
                         Having Nvl(Sum(Nvl(付数, 1) * 数次), 0) <> 0)
                  Union All
                  Select Distinct 2 As 费用来源, Decode(a.多病人单, 1, 3, 2) As 单据类型, a.No
                  From 住院费用记录 A, 病人费用异常记录 B
                  Where a.收费类别 In ('5', '6', '7') And a.记录性质 = 2 And
                        a.id = b.费用id And b.产生环节 = 0 And Nvl(b.同步标志, 0) = 1 And Exists
                   (Select 1
                         From 住院费用记录
                         Where 记录性质 = a.记录性质 And NO = a.No And 序号 = a.序号
                         Group By 记录性质, NO, 序号
                         Having Nvl(Sum(Nvl(付数, 1) * 数次), 0) <> 0)) Loop

      v_Json := Null;
      v_Json := v_Json || '{"fee_no":"' || r_Fee.No || '"';
      v_Json := v_Json || ',"billtype":' || r_Fee.单据类型;
      v_Json := v_Json || ',"fee_source":' || r_Fee.费用来源;
      v_Json := v_Json || '}';
      v_Json := '{"input":' || v_Json || '}';

      Json_Temp_Out := Null;
      Zl_Drugbill_Build(v_Json, Json_Temp_Out);

      --解析出参
      j_Json_Out := PLJson();
      j_Json_Out := PLJson(Json_Temp_Out);
      j_Json     := PLJson();
      j_Json     := j_Json_Out.Get_Pljson('output');

      n_Code := Nvl(j_Json.Get_Number('code'), '0');
      If n_Code = 0 Then
        Json_Out := zlJsonOut(j_Json.Get_String('message'));
        Return;
      End If;

      j_Json.Remove('code');
      j_Json.Remove('message');
      Json_Temp_Out := Empty_Clob();
      Dbms_Lob.Createtemporary(Json_Temp_Out, True);
      j_Json.To_Clob(Json_Temp_Out);

      If c_Jtmp Is Null Then
        c_Jtmp := Json_Temp_Out;
      Else
        c_Jtmp := c_Jtmp || ',' || Json_Temp_Out;
      End If;
    End Loop;
  Else
    --解析入参
    j_Json_In   := PLJson(Json_In);
    j_Json      := j_Json_In.Get_Pljson('input');
    j_Pati_List := j_Json.Get_Pljson_List('pati_list');
    
    For I In 1 .. j_Pati_List.Count Loop
    
      j_Item      := PLJson();
      j_Item      := PLJson(j_Pati_List.Get(I));
      n_病人id    := j_Item.Get_Number('pati_id');
      j_Bill_List := j_Item.Get_Pljson_List('bill_list');

      If j_Bill_List Is Not Null Then
        For J In 1 .. j_Bill_List.Count Loop
          j_Item     := PLJson();
          j_Item     := PLJson(j_Bill_List.Get(J));
          n_费用来源 := j_Item.Get_Number('fee_source');
          n_记录性质 := j_Item.Get_Number('fee_billtype');
          v_No       := j_Item.Get_String('fee_no');

          If n_费用来源 = 1 Then
            l_Outnos.Extend;
            l_Outnos(l_Outnos.Count) := t_StrObj2(n_记录性质, v_No);
          Else
            l_Innos.Extend;
            l_Innos(l_Innos.Count) := t_StrObj2(n_记录性质, v_No);
          End If;
        End Loop;
      End If;

      --费用系统中记费同步异常的数据
      For r_Fee In (Select Distinct 1 As 费用来源, Decode(Mod(a.记录性质, 10), 2, 2, 1) As 单据类型, a.No
                    From 门诊费用记录 A, 病人费用异常记录 B
                    Where a.病人id = n_病人id And a.收费类别 In ('5', '6', '7') And a.记录性质 In (1, 2) And
                          a.id = b.费用id And (b.产生环节 = 0 Or b.产生环节 = 1) And Nvl(b.同步标志, 0) = 1 And
                          Exists (Select 1
                           From 门诊费用记录
                           Where 记录性质 = a.记录性质 And NO = a.No And 序号 = a.序号
                           Group By 记录性质, NO, 序号
                           Having Nvl(Sum(Nvl(付数, 1) * 数次), 0) <> 0)
                    Union All
                    Select Distinct 2 As 费用来源, Decode(a.多病人单, 1, 3, 2) As 单据类型, a.No
                    From 住院费用记录 A, 病人费用异常记录 B
                    Where a.病人id = n_病人id And a.收费类别 In ('5', '6', '7') And a.记录性质 = 2 And
                          a.id = b.费用id And b.产生环节 = 0 And Nvl(b.同步标志, 0) = 1 And Exists
                     (Select 1
                           From 住院费用记录
                           Where 记录性质 = a.记录性质 And NO = a.No And 序号 = a.序号
                           Group By 记录性质, NO, 序号
                           Having Nvl(Sum(Nvl(付数, 1) * 数次), 0) <> 0)) Loop

        v_Json := Null;
        v_Json := v_Json || '{"fee_no":"' || r_Fee.No || '"';
        v_Json := v_Json || ',"billtype":' || r_Fee.单据类型;
        v_Json := v_Json || ',"fee_source":' || r_Fee.费用来源;
        v_Json := v_Json || '}';
        v_Json := '{"input":' || v_Json || '}';

        Json_Temp_Out := Null;
        Zl_Drugbill_Build(v_Json, Json_Temp_Out);

        --解析出参
        j_Json_Out := PLJson();
        j_Json_Out := PLJson(Json_Temp_Out);
        j_Json     := PLJson();
        j_Json     := j_Json_Out.Get_Pljson('output');

        n_Code := Nvl(j_Json.Get_Number('code'), '0');
        If n_Code = 0 Then
          Json_Out := zlJsonOut(j_Json.Get_String('message'));
          Return;
        End If;

        j_Json.Remove('code');
        j_Json.Remove('message');
        Json_Temp_Out := Empty_Clob();
        Dbms_Lob.Createtemporary(Json_Temp_Out, True);
        j_Json.To_Clob(Json_Temp_Out);

        If c_Jtmp Is Null Then
          c_Jtmp := Json_Temp_Out;
        Else
          c_Jtmp := c_Jtmp || ',' || Json_Temp_Out;
        End If;
      End Loop;

      --临床系统中记费同步异常的数据
      For r_Fee In (Select /*+Cardinality(j,10)*/
                    Distinct 1 As 费用来源, Decode(Mod(a.记录性质, 10), 2, 2, 1) As 单据类型, a.No
                    From 门诊费用记录 A, Table(l_Outnos) J
                    Where a.病人id = n_病人id And a.收费类别 In ('5', '6', '7') And a.记录性质 = j.C1 And a.No = j.C2 And Exists
                     (Select 1
                           From 门诊费用记录
                           Where 记录性质 = a.记录性质 And NO = a.No And 序号 = a.序号
                           Group By 记录性质, NO, 序号
                           Having Nvl(Sum(Nvl(付数, 1) * 数次), 0) <> 0)
                    Union All
                    Select /*+Cardinality(j,10)*/
                    Distinct 2 As 费用来源, Decode(a.多病人单, 1, 3, 2) As 单据类型, a.No
                    From 住院费用记录 A, Table(l_Innos) J
                    Where a.病人id = n_病人id And a.收费类别 In ('5', '6', '7') And a.记录性质 = j.C1 And a.No = j.C2 And Exists
                     (Select 1
                           From 住院费用记录
                           Where 记录性质 = a.记录性质 And NO = a.No And 序号 = a.序号
                           Group By 记录性质, NO, 序号
                           Having Nvl(Sum(Nvl(付数, 1) * 数次), 0) <> 0)) Loop

        v_Json := Null;
        v_Json := v_Json || '{"fee_no":"' || r_Fee.No || '"';
        v_Json := v_Json || ',"billtype":' || r_Fee.单据类型;
        v_Json := v_Json || ',"fee_source":' || r_Fee.费用来源;
        v_Json := v_Json || '}';
        v_Json := '{"input":' || v_Json || '}';

        Json_Temp_Out := Null;
        Zl_Drugbill_Build(v_Json, Json_Temp_Out);

        --解析出参
        j_Json_Out := PLJson();
        j_Json_Out := PLJson(Json_Temp_Out);
        j_Json     := PLJson();
        j_Json     := j_Json_Out.Get_Pljson('output');

        n_Code := Nvl(j_Json.Get_Number('code'), '0');
        If n_Code = 0 Then
          Json_Out := zlJsonOut(j_Json.Get_String('message'));
          Return;
        End If;

        j_Json.Remove('code');
        j_Json.Remove('message');
        Json_Temp_Out := Empty_Clob();
        Dbms_Lob.Createtemporary(Json_Temp_Out, True);
        j_Json.To_Clob(Json_Temp_Out);

        If c_Jtmp Is Null Then
          c_Jtmp := Json_Temp_Out;
        Else
          c_Jtmp := c_Jtmp || ',' || Json_Temp_Out;
        End If;
      End Loop;

    End Loop;
  End If;
  Json_Out := '{"output":{"code":1,"message": "成功","pati_bill_list":[' || c_Jtmp || ']}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getdrugerrdata;
/

CREATE OR REPLACE Procedure Zl_Exsesvr_Getstufferrdata
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：根据病人ID和医嘱信息返回病人费用信息
  --入参：Json_In:格式。当值为null时获取所有病人的异常信息
  --  input
  --    pati_list[]病人列表
  --       pati_id                    N 1 病人id
  --       bill_list[]                费用单据号列表，可以不传，不传时表示获取费用域同步异常的数据
  --         fee_source               N 0 费用来源：1-门诊；2-住院
  --         fee_billtype             N 0 费用单据类型：1-收费处方；2-记帐单处方
  --         fee_no                   C 0 费用单据号
  --出参: Json_Out,格式如下
  --  output
  --    code                          N   1 应答吗：0-失败；1-成功
  --    message                       C   1 应答消息：失败时返回具体的错误信息
  --    pati_bill_list[]
  --       billtype                   N   1 单据类型: 1 -收费处方  ;2- 记帐单处方;3- 记帐表处方
  --       pati_source                N   1 病人来源:1-门诊;2-住院;4-体检
  ---------------------------billtype = 3,记帐表处方（药品收发记录.单据=26）时，无以下节点--------------------------------------
  --       pati_id                    N   1 病人ID
  --       pati_pageid                N   1 主页ID
  --       pati_name                  C   1 病人姓名
  --       pati_sex_code              C   1 性别编号（新门诊)
  --       pati_sex                   C   1 性别
  --       pati_age                   C   1 年龄
  --       pati_deptid                N   1 病人科室ID
  --       pati_wardarea_id           N     病人病区ID
  ---------------------------billtype = 3,记帐表处方（药品收发记录.单据=26）时，无以上节点-----------------------------------------
  --       bill_list[]                      更新数据列表[数组]
  --         fee_source                N  0 费用来源
  --         stuff_no                  C  1 NO
  --         charge_tag                N  1 收费标志:0-未收费或记帐划价;1-已收费或记帐
  --         fee_acnter                C  0 划价人
  --         plcdept_id                C  0 开单科室id（新门诊)
  --         plcdept                   C  0 开单科室名称（新门诊)
  --         placer_id                 C  0 开单医师id（新门诊)
  --         placer                    C  0 开单医师（新门诊) 增加
  --         operator_name             C  1 操作员姓名
  --         operator_code             C  1 操作员编号
  --         create_time               C  1 登记时间:yyyy-mm-dd hh:mi:ss
  --         item_list[]                    更新数据列表[数组]
  ---------------------------billtype = 3,记帐表处方（药品收发记录.单据=26）时，有以下节点----------------------------------------
  --           pati_id                 N  1 病人ID
  --           pati_pageid             N  0 主页ID
  --           pati_name               C  1 病人姓名
  --           pati_sex_code           C  1 性别编号（新门诊)
  --           pati_sex                C  1 性别
  --           pati_age                C  1 年龄
  --           pati_wardarea_id        N    病人病区ID
  --           pati_deptid             N  1 病人科室ID
  ---------------------------billtype = 3,记帐表处方（药品收发记录.单据=26）时，有以上节点-----------------------------------------

  --           stuffdtl_id             N  1 处方明细ID(目前传入的是费用id)
  --           serial_num              N  1 序号:(变更(包括存储)：序号和组号，1、2、3、3、3、4…)
  --           warehouse_id            N  1 库房ID
  --           is_bakstuff             N  1 是否备货卫材:有高值卫材才需要传入，非0表示是高值卫材模式(如扫码时使用)
  --           bakstuff_batch          N  1 备货材料批次
  --           stuff_id                N  1 卫材ID
  --           baby_num                N  0 婴儿序号
  --           advice_id               N  0 医嘱ID
  --           packages_num            N  1 付数
  --           outbound_num            N  1 出库数量
  --           price                   N  0 售价
  --           money                   N  0 零售金额(新门诊)
  --           memo                    C  0 摘要
  ------------------------------------------------------------------------------------------------------------
  j_Json      PLJson;
  j_Json_In   PLJson;
  j_Pati_List Pljson_List;
  j_Json_Out  PLJson;
  j_Bill_List Pljson_List;

  Json_Temp_Out Clob;
  c_Jtmp        Clob;

  j_Item   PLJson;
  n_病人id Number(18);

  v_Json Varchar2(4000);
  n_Code Number;

  n_费用来源 Number(1);
  n_记录性质 门诊费用记录.记录性质%Type;
  v_No       门诊费用记录.No%Type;

  l_Outnos t_StrList2 := t_StrList2();
  l_Innos  t_StrList2 := t_StrList2();
Begin
  If Json_In Is Null Then
    --费用系统中记费同步异常的数据
    For r_Fee In (Select Distinct 1 As 费用来源, Decode(Mod(a.记录性质, 10), 2, 2, 1) As 单据类型, a.No
                  From 门诊费用记录 A, 病人费用异常记录 B
                  Where a.病人id = n_病人id And a.收费类别 = '4' And a.记录性质 In (1, 2) And 
                        a.id = b.费用id And (b.产生环节 = 0 Or b.产生环节 = 1) And Nvl(b.同步标志, 0) = 1 And Exists
                   (Select 1
                         From 门诊费用记录
                         Where 记录性质 = a.记录性质 And NO = a.No And 序号 = a.序号
                         Group By 记录性质, NO, 序号
                         Having Nvl(Sum(Nvl(付数, 1) * 数次), 0) <> 0)
                  Union All
                  Select Distinct 2 As 费用来源, Decode(a.多病人单, 1, 3, 2) As 单据类型, a.No
                  From 住院费用记录 A, 病人费用异常记录 B
                  Where a.病人id = n_病人id And a.收费类别 = '4' And a.记录性质 = 2 And 
                        a.id = b.费用id And b.产生环节 = 0 And Nvl(b.同步标志, 0) = 1 And Exists
                   (Select 1
                         From 住院费用记录
                         Where 记录性质 = a.记录性质 And NO = a.No And 序号 = a.序号
                         Group By 记录性质, NO, 序号
                         Having Nvl(Sum(Nvl(付数, 1) * 数次), 0) <> 0)) Loop
    
      v_Json := Null;
      v_Json := v_Json || '{"fee_no":"' || r_Fee.No || '"';
      v_Json := v_Json || ',"billtype":' || r_Fee.单据类型;
      v_Json := v_Json || ',"fee_source":' || r_Fee.费用来源;
      v_Json := v_Json || '}';
      v_Json := '{"input":' || v_Json || '}';
    
      Json_Temp_Out := Null;
      Zl_Stuffbill_Build(v_Json, Json_Temp_Out);
    
      --解析出参
      j_Json_Out := PLJson();
      j_Json_Out := PLJson(Json_Temp_Out);
      j_Json     := PLJson();
      j_Json     := j_Json_Out.Get_Pljson('output');
    
      n_Code := Nvl(j_Json.Get_Number('code'), '0');
      If n_Code = 0 Then
        Json_Out := zlJsonOut(j_Json.Get_String('message'));
        Return;
      End If;
    
      j_Json.Remove('code');
      j_Json.Remove('message');
      Json_Temp_Out := Empty_Clob();
      Dbms_Lob.Createtemporary(Json_Temp_Out, True);
      j_Json.To_Clob(Json_Temp_Out);
    
      If c_Jtmp Is Null Then
        c_Jtmp := Json_Temp_Out;
      Else
        c_Jtmp := c_Jtmp || ',' || Json_Temp_Out;
      End If;
    End Loop;
  Else
    --解析入参
    j_Json_In   := PLJson(Json_In);
    j_Json      := j_Json_In.Get_Pljson('input');
    j_Pati_List := j_Json.Get_Pljson_List('pati_list');
    
    For I In 1 .. j_Pati_List.Count Loop
      j_Item      := PLJson();
      j_Item      := PLJson(j_Pati_List.Get(I));
      n_病人id    := j_Item.Get_Number('pati_id');
      j_Bill_List := j_Item.Get_Pljson_List('bill_list');
    
      If j_Bill_List Is Not Null Then
        For J In 1 .. j_Bill_List.Count Loop
          j_Item     := PLJson();
          j_Item     := PLJson(j_Bill_List.Get(J));
          n_费用来源 := j_Item.Get_Number('fee_source');
          n_记录性质 := j_Item.Get_Number('fee_billtype');
          v_No       := j_Item.Get_String('fee_no');
        
          If n_费用来源 = 1 Then
            l_Outnos.Extend;
            l_Outnos(l_Outnos.Count) := t_StrObj2(n_记录性质, v_No);
          Else
            l_Innos.Extend;
            l_Innos(l_Innos.Count) := t_StrObj2(n_记录性质, v_No);
          End If;
        End Loop;
      End If;
    
      --费用系统中记费同步异常的数据
      For r_Fee In (Select Distinct 1 As 费用来源, Decode(Mod(a.记录性质, 10), 2, 2, 1) As 单据类型, a.No
                    From 门诊费用记录 A, 病人费用异常记录 B
                    Where a.病人id = n_病人id And a.收费类别 = '4' And a.记录性质 In (1, 2) And 
                          a.id = b.费用id And (b.产生环节 = 0 Or b.产生环节 = 1) And Nvl(b.同步标志, 0) = 1 And Exists
                     (Select 1
                           From 门诊费用记录
                           Where 记录性质 = a.记录性质 And NO = a.No And 序号 = a.序号
                           Group By 记录性质, NO, 序号
                           Having Nvl(Sum(Nvl(付数, 1) * 数次), 0) <> 0)
                    Union All
                    Select Distinct 2 As 费用来源, Decode(a.多病人单, 1, 3, 2) As 单据类型, a.No
                    From 住院费用记录 A, 病人费用异常记录 B
                    Where a.病人id = n_病人id And a.收费类别 = '4' And a.记录性质 = 2 And 
                          a.id = b.费用id And b.产生环节 = 0 And Nvl(b.同步标志, 0) = 1 And Exists
                     (Select 1
                           From 住院费用记录
                           Where 记录性质 = a.记录性质 And NO = a.No And 序号 = a.序号
                           Group By 记录性质, NO, 序号
                           Having Nvl(Sum(Nvl(付数, 1) * 数次), 0) <> 0)) Loop
      
        v_Json := Null;
        v_Json := v_Json || '{"fee_no":"' || r_Fee.No || '"';
        v_Json := v_Json || ',"billtype":' || r_Fee.单据类型;
        v_Json := v_Json || ',"fee_source":' || r_Fee.费用来源;
        v_Json := v_Json || '}';
        v_Json := '{"input":' || v_Json || '}';
      
        Json_Temp_Out := Null;
        Zl_Stuffbill_Build(v_Json, Json_Temp_Out);
      
        --解析出参
        j_Json_Out := PLJson();
        j_Json_Out := PLJson(Json_Temp_Out);
        j_Json     := PLJson();
        j_Json     := j_Json_Out.Get_Pljson('output');
      
        n_Code := Nvl(j_Json.Get_Number('code'), '0');
        If n_Code = 0 Then
          Json_Out := zlJsonOut(j_Json.Get_String('message'));
          Return;
        End If;
      
        j_Json.Remove('code');
        j_Json.Remove('message');
        Json_Temp_Out := Empty_Clob();
        Dbms_Lob.Createtemporary(Json_Temp_Out, True);
        j_Json.To_Clob(Json_Temp_Out);
      
        If c_Jtmp Is Null Then
          c_Jtmp := Json_Temp_Out;
        Else
          c_Jtmp := c_Jtmp || ',' || Json_Temp_Out;
        End If;
      End Loop;
    
      --临床系统中记费同步异常的数据 
      For r_Fee In (Select /*+Cardinality(j,10)*/
                    Distinct 1 As 费用来源, Decode(Mod(a.记录性质, 10), 2, 2, 1) As 单据类型, a.No
                    From 门诊费用记录 A, Table(l_Outnos) J
                    Where a.病人id = n_病人id And a.收费类别 = '4' And a.记录性质 = j.C1 And a.No = j.C2 And Exists
                     (Select 1
                           From 门诊费用记录
                           Where 记录性质 = a.记录性质 And NO = a.No And 序号 = a.序号
                           Group By 记录性质, NO, 序号
                           Having Nvl(Sum(Nvl(付数, 1) * 数次), 0) <> 0)
                    Union All
                    Select /*+Cardinality(j,10)*/
                    Distinct 2 As 费用来源, Decode(a.多病人单, 1, 3, 2) As 单据类型, a.No
                    From 住院费用记录 A, Table(l_Innos) J
                    Where a.病人id = n_病人id And a.收费类别 = '4' And a.记录性质 = j.C1 And a.No = j.C2 And Exists
                     (Select 1
                           From 住院费用记录
                           Where 记录性质 = a.记录性质 And NO = a.No And 序号 = a.序号
                           Group By 记录性质, NO, 序号
                           Having Nvl(Sum(Nvl(付数, 1) * 数次), 0) <> 0)) Loop
      
        v_Json := Null;
        v_Json := v_Json || '{"fee_no":"' || r_Fee.No || '"';
        v_Json := v_Json || ',"billtype":' || r_Fee.单据类型;
        v_Json := v_Json || ',"fee_source":' || r_Fee.费用来源;
        v_Json := v_Json || '}';
        v_Json := '{"input":' || v_Json || '}';
      
        Json_Temp_Out := Null;
        Zl_Stuffbill_Build(v_Json, Json_Temp_Out);
      
        --解析出参
        j_Json_Out := PLJson();
        j_Json_Out := PLJson(Json_Temp_Out);
        j_Json     := PLJson();
        j_Json     := j_Json_Out.Get_Pljson('output');
      
        n_Code := Nvl(j_Json.Get_Number('code'), '0');
        If n_Code = 0 Then
          Json_Out := zlJsonOut(j_Json.Get_String('message'));
          Return;
        End If;
      
        j_Json.Remove('code');
        j_Json.Remove('message');
        Json_Temp_Out := Empty_Clob();
        Dbms_Lob.Createtemporary(Json_Temp_Out, True);
        j_Json.To_Clob(Json_Temp_Out);
      
        If c_Jtmp Is Null Then
          c_Jtmp := Json_Temp_Out;
        Else
          c_Jtmp := c_Jtmp || ',' || Json_Temp_Out;
        End If;
      End Loop;
    End Loop;
  End If;
  Json_Out := '{"output":{"code":1,"message": "成功","pati_bill_list":[' || c_Jtmp || ']}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getstufferrdata;
/

Create Or Replace Procedure Zl_Exsesvr_Getorderfeestate
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:根据医嘱id,获取相关的收费状态
  --入参：Json_In:格式
  --input     
  --  advice_ids  C 1 医嘱id
  --  bill_nos  C 1 单据号
  --出参: Json_Out,格式如下
  -- output      
  --   code  C 1 应答码：0-失败；1-成功
  --   message C 1 应答消息：成功时返回成功信息，失败时返回具体的错误信息
  --   state N 1 状态:0-未收费,1-完全收费;2-部门收费
  --   billtype  N 1 单据类型:0-不存在任何单据;1-收费单;2-记帐单;3-收费和记帐都有
  --   advice_ids  C   未收费的医嘱ID:传入医嘱ids时有效

  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_Output Varchar2(32767);

  v_医嘱ids Varchar2(32767);

  v_单据号      Varchar2(32767);
  n_状态        Number(2);
  v_未收医嘱ids Varchar2(32767);
  n_单据类型    Number(2);

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_医嘱ids := j_Json.Get_String('advice_ids');
  v_单据号  := j_Json.Get_String('bill_nos');

  n_状态        := -1;
  v_未收医嘱ids := '';
  n_单据类型    := 0;
  If v_医嘱ids Is Not Null Then
  
    For c_医嘱 In (
                 
                 Select /*+ RULE */
                 Distinct 记录性质, 记录状态, 医嘱序号
                 From (With 医嘱数据 As (Select Column_Value As 医嘱id From Table(f_Num2List(v_医嘱ids)))
                         Select Distinct a.记录性质, a.记录状态, a.医嘱序号
                         From 门诊费用记录 A, 医嘱数据 B
                         Where a.医嘱序号 = b.医嘱id And a.记录性质 In (1, 2, 3) And a.记录状态 In (0, 1, 3)
                         Union All
                         Select Distinct a.记录性质, a.记录状态, a.医嘱序号
                         From 住院费用记录 A, 医嘱数据 B
                         Where a.医嘱序号 = b.医嘱id And a.记录性质 In (1, 2, 3) And a.记录状态 In (0, 1, 3))
                 ) Loop
    
      If c_医嘱.记录状态 = 0 Then
        --未收费
        If Nvl(c_医嘱.医嘱序号, 0) <> 0 Then
          v_未收医嘱ids := Nvl(v_未收医嘱ids, '') || ',' || Nvl(c_医嘱.医嘱序号, 0);
        
        End If;
      End If;
    
      If n_状态 = -1 Then
        If c_医嘱.记录状态 = 0 Then
          n_状态 := Case
                    When c_医嘱.记录状态 = 0 Then
                     0
                    Else
                     1
                  End;
        End If;
      Elsif n_状态 = 0 And (c_医嘱.记录状态 = 1 Or c_医嘱.记录状态 = 3) Then
        n_状态 := 2; --   部分收费
      Elsif n_状态 = 1 And c_医嘱.记录状态 = 0 Then
        n_状态 := 2; --部分收费
      End If;
    
      If n_单据类型 = 0 Then
        n_单据类型 := c_医嘱.记录性质;
      Elsif n_单据类型 <> c_医嘱.记录性质 Then
        --两都都有
        n_单据类型 := 3;
      End If;
    
    End Loop;
  
    zlJsonPutValue(v_Output, 'code', 1, 1, 1);
    zlJsonPutValue(v_Output, 'message', '成功');
    zlJsonPutValue(v_Output, 'state', Nvl(n_状态, 0), 1);
    zlJsonPutValue(v_Output, 'billtype', Nvl(n_单据类型, 0), 1);
    zlJsonPutValue(v_Output, 'advice_ids', Nvl(v_未收医嘱ids, ''), 0, 2);
    Json_Out := '{"output":' || v_Output || '}';
    Return;
  
  End If;

  For c_医嘱 In (Select /*+ RULE */
               Distinct 记录性质, 记录状态, 医嘱序号
               From (With 医嘱数据 As (Select Column_Value As NO From Table(f_Str2List(v_单据号)))
                      Select Distinct a.记录性质, a.记录状态, a.医嘱序号
                      From 门诊费用记录 A, 医嘱数据 B
                      Where a.No = b.No And a.记录性质 In (1, 2, 3) And a.记录状态 In (0, 1, 3)
                      Union All
                      Select Distinct a.记录性质, a.记录状态, a.医嘱序号
                      From 住院费用记录 A, 医嘱数据 B
                      Where a.No = b.No And a.记录性质 In (1, 2, 3) And a.记录状态 In (0, 1, 3))
               ) Loop
  
    If c_医嘱.记录状态 = 0 Then
      --未收费
      If Nvl(c_医嘱.医嘱序号, 0) <> 0 Then
        v_未收医嘱ids := Nvl(v_未收医嘱ids, '') || ',' || Nvl(c_医嘱.医嘱序号, 0);
      End If;
    End If;
  
    If n_状态 = -1 Then
      If c_医嘱.记录状态 = 0 Then
        n_状态 := Case
                  When c_医嘱.记录状态 = 0 Then
                   0
                  Else
                   1
                End;
      End If;
    Elsif n_状态 = 0 And (c_医嘱.记录状态 = 1 Or c_医嘱.记录状态 = 3) Then
      n_状态 := 2; --   部分收费
    Elsif n_状态 = 1 And c_医嘱.记录状态 = 0 Then
      n_状态 := 2; --部分收费
    End If;
  
    If n_单据类型 = 0 Then
      n_单据类型 := c_医嘱.记录性质;
    Elsif n_单据类型 <> c_医嘱.记录性质 Then
      --两都都有
      n_单据类型 := 3;
    End If;
  End Loop;

  --    state  N  1  状态:0-未收费,1-完全收费;2-部门收费
  --    billtype  N  1  单据类型:0-不存在任何单据;1-收费单;2-记帐单;3-收费和记帐都有
  --    advice_ids  C    未收费的医嘱ID:传入医嘱ids时有效

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '成功');
  zlJsonPutValue(v_Output, 'state', Nvl(n_状态, 0), 1);
  zlJsonPutValue(v_Output, 'billtype', Nvl(n_单据类型, 0), 1);
  zlJsonPutValue(v_Output, 'advice_ids', Nvl(v_未收医嘱ids, ''), 0, 2);
  Json_Out := '{"output":' || v_Output || '}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getorderfeestate;
/


Create Or Replace Procedure Zl_Exsesvr_Getfeechargestate
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:根据单据号信息，获取单据对应的收费状态
  --入参：Json_In:格式
  --input     
  --    query_mode  N 1 查询方式:0-查询收费状态;1-仅查是否存在未收费的
  --    bill_nos  C 1 单据号
  --    bill_type N 1 单据类型:1-收费单据;待以后扩展

  --出参: Json_Out,格式如下
  -- output      
  --   code  C 1 应答码：0-失败；1-成功
  --   message C 1 应答消息：成功时返回成功信息，失败时返回具体的错误信息
  --   state N 1 状态
  --       1.query_mode=0时
  --         状态:0-未收费;1-部分收费/退费;2-全部收费;3-全部退费
  --       2.query_mode=1时
  --         状态:1-存在未收费的;0-不存在未收费.
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_Output     Varchar2(32767);
  n_查询方式   Number(2);
  v_单据号     Varchar2(32767);
  n_单据类型   Number(2);
  n_状态       Number(2);
  n_是否全收   Number(2);
  n_是否全退   Number(2);
  n_是否部分退 Number(2);
  n_是否未收   Number(2);

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_查询方式 := Nvl(j_Json.Get_Number('query_mode'), 0);
  v_单据号   := Nvl(j_Json.Get_String('bill_nos'), '');
  n_单据类型 := Nvl(j_Json.Get_Number('bill_type'), 1);

  If Nvl(n_单据类型, 0) <> 1 Then
    Json_Out := zlJsonOut('暂不支持非收费单据。');
    Return;
  End If;

  If n_查询方式 = 1 Then
    Begin
      --判断费用状态，主要是异常的，可能需要重收
      If Instr(v_单据号, ',') > 0 Then
        Select /*+cardinality(b,10)*/
         1
        Into n_状态
        From 门诊费用记录 A, (Select Column_Value As NO From Table(f_Str2List(v_单据号))) B
        Where a.记录性质 = 1 And (a.结帐id Is Null Or Nvl(a.费用状态, 0) = 1 And a.记录状态 = 1) And a.No = b.No And Rownum < 2;
      Else
        Select 1
        Into n_状态
        From 门诊费用记录 A
        Where a.记录性质 = 1 And (a.结帐id Is Null Or Nvl(a.费用状态, 0) = 1 And a.记录状态 = 1) And a.No = v_单据号 And Rownum < 2;
      End If;
    Exception
      When Others Then
        n_状态 := 0;
    End;
    zlJsonPutValue(v_Output, 'code', 1, 1, 1);
    zlJsonPutValue(v_Output, 'message', '成功');
    zlJsonPutValue(v_Output, 'state', Nvl(n_状态, 0), 1, 2);
    Json_Out := '{"output":' || v_Output || '}';
    Return;
  End If;

  n_状态       := -1;
  n_是否全收   := -1;
  n_是否部分退 := -1;
  n_是否全退   := -1;
  n_是否未收   := -1;
  For c_费用 In (Select /*+cardinality(b,10)*/
                a.No, a.序号, Nvl(Sum(a.数次 * Nvl(a.付数, 1)), 0) As 剩余数量,
                Nvl(Sum(Decode(a.记录性质, 1, 1, 0) * Decode(a.记录状态, 2, 0, 1) * a.数次 * Nvl(a.付数, 1)), 0) As 原始数量,
                Nvl(Sum(Decode(a.记录性质, 1, 1, 0) *
                         Decode(a.结帐id, Null, 1, Decode(a.记录状态, 0, 1, 1, Decode(Nvl(a.费用状态, 0), 1, 1, 0), 0)) * a.数次 *
                         Nvl(a.付数, 1)), 0) As 未收数量
               From 门诊费用记录 A, (Select Column_Value As NO From Table(f_Str2List(v_单据号))) B
               Where Mod(a.记录性质, 10) = 1 And a.价格父号 Is Null And a.No = b.No
               Group By a.No, a.序号) Loop
  
    If c_费用.原始数量 <> 0 And c_费用.原始数量 = c_费用.未收数量 Then
      --未收费
      n_是否未收 := 1;
    Elsif c_费用.原始数量 = c_费用.剩余数量 And c_费用.未收数量 = 0 Then
      --全部收费
      n_是否全收 := 1;
    Elsif c_费用.剩余数量 = 0 Then
      --全退了
      n_是否全退 := 1;
    Else
      --部分收费或退费 
      n_是否部分退 := 1;
      Exit;
    End If;
    If n_是否未收 <> -1 And n_是否全收 <> -1 And n_是否部分退 <> -1 Then
      Exit;
    End If;
  End Loop;
  --1-不存在单据,0-未收费;1-部分收费或退费;2-全部收费;3-全部退费
  If n_是否部分退 = 1 Then
    n_状态 := 1;
  Elsif n_是否全收 = -1 And n_是否全退 = 1 And n_是否未收 = -1 Then
    --全退
    n_状态 := 3;
  Elsif n_是否全收 = 1 And n_是否全退 = -1 And n_是否未收 = -1 Then
    --全收
    n_状态 := 2;
  Elsif n_是否全收 = -1 And n_是否全退 = -1 And n_是否未收 = 1 Then
    n_状态 := 0;
  Elsif n_是否全收 = -1 And n_是否全退 = -1 And n_是否未收 = -1 Then
    n_状态 := -1;
  Else
    n_状态 := 1; --部分收或退
  End If;
  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '成功');
  zlJsonPutValue(v_Output, 'state', Nvl(n_状态, 0), 1, 2);
  Json_Out := '{"output":' || v_Output || '}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getfeechargestate;
/

Create Or Replace Procedure Zl_Exsesvr_Getfeebalancestate
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:根据单据号信息，获取单据对应的结帐状态
  --入参：Json_In:格式
  --input     
  --    query_mode  N 1 查询方式:0-门诊记帐;1-住院记帐
  --    bill_nos  C 1 单据号
  --出参: Json_Out,格式如下
  -- output      
  --    code  C 1 应答码：0-失败；1-成功
  --    message C 1 应答消息：成功时返回成功信息，失败时返回具体的错误信息
  --    state N 1 状态:-1-不存在记帐单据;0-未结帐;1-部分结帐;2-全部结帐

  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_Output   Varchar2(32767);
  n_查询方式 Number(2);
  v_单据号   Varchar2(32767);
  n_结帐标志 Number(18);
  n_存在     Number(18);
Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_查询方式 := Nvl(j_Json.Get_Number('query_mode'), 0);
  v_单据号   := Nvl(j_Json.Get_String('bill_nos'), '');

  If Nvl(n_查询方式, 0) = 0 Then
    If Instr(v_单据号, ',') > 0 Then
      Select Decode(Nvl(Sum(Nvl((Case
                                   When (未结金额 <> 0 And 结帐金额 = 0) Or (未结金额 = 0 And (实收金额 = 0 Or 结帐金额 = 0) And 统计数 = 0) Then
                                    0
                                   When 未结金额 <> 0 And 结帐金额 <> 0 Then
                                    1
                                   Else
                                    2
                                 End), 0)), 0), 0, 0, 2 * Count(1), 2, 1), Max(存在)
      Into n_结帐标志, n_存在
      From (Select /*+Cardinality(B,10)*/
              a.No, Nvl(a.价格父号, a.序号) As 序号, Nvl(Sum(Nvl(a.应收金额, 0)), 0) As 应收金额, Nvl(Sum(Nvl(a.实收金额, 0)), 0) As 实收金额,
              Nvl(Sum(Nvl(a.结帐金额, 0)), 0) As 结帐金额, Nvl(Sum(Nvl(a.实收金额, 0)) - Sum(Nvl(a.结帐金额, 0)), 0) As 未结金额,
              Mod(Sum(Decode(Nvl(a.结帐id, 0), 0, 0, 1)), 2) As 统计数, Max(1) As 存在
             From 门诊费用记录 A, Table(f_Str2List(v_单据号)) B
             Where a.No = b.Column_Value And a.记帐费用 = 1 And Mod(a.记录性质, 10) = 2
             Group By a.No, Nvl(a.价格父号, a.序号));
    
    Else
    
      Select Decode(Nvl(Sum(Nvl((Case
                                  When (未结金额 <> 0 And 结帐金额 = 0) Or (未结金额 = 0 And (实收金额 = 0 Or 结帐金额 = 0) And 统计数 = 0) Then
                                   0
                                  When 未结金额 <> 0 And 结帐金额 <> 0 Then
                                   1
                                  Else
                                   2
                                End), 0)), 0), 0, 0, 2 * Count(1), 2, 1), Max(存在)
      Into n_结帐标志, n_存在
      From (Select a.No, Nvl(a.价格父号, a.序号) As 序号, Nvl(Sum(Nvl(a.应收金额, 0)), 0) As 应收金额,
                    Nvl(Sum(Nvl(a.实收金额, 0)), 0) As 实收金额, Nvl(Sum(Nvl(a.结帐金额, 0)), 0) As 结帐金额,
                    Nvl(Sum(Nvl(a.实收金额, 0)) - Sum(Nvl(a.结帐金额, 0)), 0) As 未结金额,
                    Mod(Sum(Decode(Nvl(a.结帐id, 0), 0, 0, 1)), 2) As 统计数, Max(1) As 存在
             From 门诊费用记录 A
             Where a.No = v_单据号 And a.记帐费用 = 1 And Mod(a.记录性质, 10) = 2
             Group By a.No, Nvl(a.价格父号, a.序号));
    End If;
  Else
    If Instr(v_单据号, ',') > 0 Then
      Select Decode(Nvl(Sum(Nvl((Case
                                   When (未结金额 <> 0 And 结帐金额 = 0) Or (未结金额 = 0 And (实收金额 = 0 Or 结帐金额 = 0) And 统计数 = 0) Then
                                    0
                                   When 未结金额 <> 0 And 结帐金额 <> 0 Then
                                    1
                                   Else
                                    2
                                 End), 0)), 0), 0, 0, 2 * Count(1), 2, 1), Max(存在)
      Into n_结帐标志, n_存在
      From (Select /*+Cardinality(B,10)*/
              a.No, Nvl(a.价格父号, a.序号) As 序号, Nvl(Sum(Nvl(a.应收金额, 0)), 0) As 应收金额, Nvl(Sum(Nvl(a.实收金额, 0)), 0) As 实收金额,
              Nvl(Sum(Nvl(a.结帐金额, 0)), 0) As 结帐金额, Nvl(Sum(Nvl(a.实收金额, 0)) - Sum(Nvl(a.结帐金额, 0)), 0) As 未结金额,
              Mod(Sum(Decode(Nvl(a.结帐id, 0), 0, 0, 1)), 2) As 统计数, Max(1) As 存在
             From 住院费用记录 A, Table(f_Str2List(v_单据号)) B
             Where a.No = b.Column_Value And a.记帐费用 = 1 And Mod(a.记录性质, 10) = 2
             Group By a.No, Nvl(a.价格父号, a.序号));
    
    Else
    
      Select Decode(Nvl(Sum(Nvl((Case
                                  When (未结金额 <> 0 And 结帐金额 = 0) Or (未结金额 = 0 And (实收金额 = 0 Or 结帐金额 = 0) And 统计数 = 0) Then
                                   0
                                  When 未结金额 <> 0 And 结帐金额 <> 0 Then
                                   1
                                  Else
                                   2
                                End), 0)), 0), 0, 0, 2 * Count(1), 2, 1), Max(存在)
      Into n_结帐标志, n_存在
      From (Select a.No, Nvl(a.价格父号, a.序号) As 序号, Nvl(Sum(Nvl(a.应收金额, 0)), 0) As 应收金额,
                    Nvl(Sum(Nvl(a.实收金额, 0)), 0) As 实收金额, Nvl(Sum(Nvl(a.结帐金额, 0)), 0) As 结帐金额,
                    Nvl(Sum(Nvl(a.实收金额, 0)) - Sum(Nvl(a.结帐金额, 0)), 0) As 未结金额,
                    Mod(Sum(Decode(Nvl(a.结帐id, 0), 0, 0, 1)), 2) As 统计数, Max(1) As 存在
             From 住院费用记录 A
             Where a.No = v_单据号 And a.记帐费用 = 1 And Mod(a.记录性质, 10) = 2
             Group By a.No, Nvl(a.价格父号, a.序号));
    End If;
  End If;
  If Nvl(n_存在, 0) = 0 Then
    n_结帐标志 := -1;
  End If;
  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '成功');
  zlJsonPutValue(v_Output, 'state', Nvl(n_结帐标志, 0), 1, 2);
  Json_Out := '{"output":' || v_Output || '}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getfeebalancestate;
/


Create Or Replace Procedure Zl_Exsesvr_Getfeeinfobyblncid
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:根据结帐id获取对应的费用明细数据
  --入参：Json_In:格式
  --input      
  -- balance_id  N 1 结帐ID

  --出参: Json_Out,格式如下
  --output     
  -- code  C 1 应答码：0-失败；1-成功
  -- message C 1 应答消息：成功时返回成功信息，失败时返回具体的错误信息
  -- fee_list[]  C 1 费用明细列表
  --   fee_no  C 1 费用单据号
  --   serial_num  N 1 序号
  --   receipt_type  C 1 收费类别
  --   fitem_id  N 1 收费细目id
  --   fitem_name  C 1 收费名称
  --   nums  N 1 数次
  --   unit  C 1 计算单位
  --   price N 1 标准单价
  --   blnc_money  N 1 结帐金额
  --   exedept_name  C 1 执行科室
  --   happen_time     C 1 发生时间:yyyy-mm-dd HH:MM:SS

  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_结帐id Number(18);
  v_Output Varchar2(32767);
  c_Output Clob;

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_结帐id := Nvl(j_Json.Get_Number('balance_id'), 0);

  For c_结帐 In (Select a.No, a.序号, a.收费类别, Nvl(e.名称, d.名称) As 收费名称, a.数量 As 收费数量, a.结帐金额, a.收费单价, a.计算单位,
                      Nvl(b.名称, '未知') As 执行科室, To_Char(a.发生时间, 'yyyy-mm-dd hh24:mi:ss') As 发生时间, a.收费细目id
               From (
                      
                      Select a.发生时间, a.No, Nvl(价格父号, 序号) As 序号, a.收费类别, a.收费细目id, Avg(Nvl(付数, 1)) * Avg(数次) As 数量, a.计算单位,
                              Sum(a.结帐金额) As 结帐金额, Sum(a.标准单价) As 收费单价, a.执行部门id
                      From 门诊费用记录 A
                      Where a.结帐id = n_结帐id
                      Group By a.发生时间, a.No, Nvl(价格父号, 序号), a.收费类别, a.收费细目id, a.计算单位, a.执行部门id
                      Union All
                      Select a.发生时间, a.No, Nvl(价格父号, 序号) As 序号, a.收费类别, a.收费细目id, Avg(Nvl(付数, 1)) * Avg(数次) As 数量, a.计算单位,
                              Sum(a.结帐金额) As 结帐金额, Sum(a.标准单价) As 收费单价, a.执行部门id
                      From 住院费用记录 A
                      Where a.结帐id = n_结帐id
                      Group By a.发生时间, a.No, Nvl(价格父号, 序号), a.收费类别, a.收费细目id, a.计算单位, a.执行部门id) A, 部门表 B, 收费项目目录 D,
                    收费项目别名 E
               Where a.执行部门id = b.Id(+) And a.收费细目id = d.Id And a.收费细目id = e.收费细目id(+) And e.码类(+) = 1 And e.性质(+) = 3
               Order By 发生时间 Desc, NO Desc, 序号) Loop
    If Length(Nvl(v_Output, ' ')) > 30000 Then
      If c_Output Is Not Null Then
        c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
      Else
        c_Output := To_Clob(v_Output);
      End If;
    
      v_Output := Null;
    End If;
  
    zlJsonPutValue(v_Output, 'fee_no', c_结帐.No, 0, 1);
    zlJsonPutValue(v_Output, 'serial_num', Nvl(c_结帐.序号, 1), 1);
    zlJsonPutValue(v_Output, 'receipt_type', Nvl(c_结帐.收费类别, ''));
    zlJsonPutValue(v_Output, 'fitem_id', Nvl(c_结帐.收费细目id, 0), 1);
    zlJsonPutValue(v_Output, 'fitem_name', Nvl(c_结帐.收费名称, ''));
    zlJsonPutValue(v_Output, 'nums', Nvl(c_结帐.收费数量, 1), 1);
    zlJsonPutValue(v_Output, 'unit', Nvl(c_结帐.计算单位, ''));
    zlJsonPutValue(v_Output, 'price', Nvl(c_结帐.收费单价, 0), 1);
    zlJsonPutValue(v_Output, 'blnc_money', Nvl(c_结帐.结帐金额, 0), 1);
    zlJsonPutValue(v_Output, 'exedept_name', Nvl(c_结帐.执行科室, ''));
    zlJsonPutValue(v_Output, 'happen_time', Nvl(c_结帐.发生时间, ''), 0, 2);
  
  End Loop;
  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"成功","fee_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"成功","fee_list":[' || v_Output || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getfeeinfobyblncid;
/


Create Or Replace Procedure Zl_Exsesvr_Getbalanceinfobyid
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:根据结帐id获取对应的结算明细数据
  --入参：Json_In:格式
  --input      
  -- balance_id  N 1 结帐ID
  --出参: Json_Out,格式如下
  --  output      
  --    code  C 1 应答码：0-失败；1-成功
  --    message C 1 应答消息：成功时返回成功信息，失败时返回具体的错误信息
  --    blnc_list[] C 1 结算明细列表
  --      blnc_mode C 1 结算方式
  --      blnc_no C 1 结算号码
  --      blnc_money  N 1 结算金额
  --      cardtype_id N 1 卡类别id
  --      consumer_no N 1 结算卡序号，即卡消费接口目录.编号
  --      cardno  C 1 卡号
  --      swapno  C 1 交易流水号
  --      swapmemo  C 1 交易说明
  --      memo  C 1 摘要
  --      cprtion_unit  C 1 合作单位
  --      relation_id N 1 关联交易id
  --
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_结帐id Number(18);
  v_Output Varchar2(32767);
Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_结帐id := Nvl(j_Json.Get_Number('balance_id'), 0);

  For c_结帐 In (
               
               Select Decode(Mod(a.记录性质, 10), 1, '[冲预交]', a.结算方式) As 结算方式, 冲预交 As 结算金额, a.结算号码, a.摘要, a.卡类别id, a.结算卡序号,
                       a.交易流水号, a.交易说明, a.卡号, a.关联交易id, 合作单位
               From 病人预交记录 A
               Where a.结帐id = n_结帐id) Loop
  
    zlJsonPutValue(v_Output, 'blnc_mode', c_结帐.结算方式, 0, 1);
    zlJsonPutValue(v_Output, 'blnc_no', Nvl(c_结帐.结算号码, ''));
    zlJsonPutValue(v_Output, 'blnc_money', Nvl(c_结帐.结算金额, 0), 1);
    zlJsonPutValue(v_Output, 'cardtype_id', Nvl(c_结帐.卡类别id, 0), 1);
    zlJsonPutValue(v_Output, 'consumer_no', Nvl(c_结帐.结算卡序号, 0), 1);
    zlJsonPutValue(v_Output, 'cardno', Nvl(c_结帐.卡号, ''));
    zlJsonPutValue(v_Output, 'swapno', Nvl(c_结帐.交易流水号, ''));
    zlJsonPutValue(v_Output, 'swapmemo', Nvl(c_结帐.交易说明, ''));
    zlJsonPutValue(v_Output, 'memo', Nvl(c_结帐.摘要, ''));
    zlJsonPutValue(v_Output, 'cprtion_unit', Nvl(c_结帐.合作单位, ''));
    zlJsonPutValue(v_Output, 'relation_id', Nvl(c_结帐.关联交易id, 0), 1, 2);
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"成功","blnc_list":[' || v_Output || ']}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getbalanceinfobyid;
/

Create Or Replace Procedure Zl_Exsesvr_Getbalanceinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) Is
  ------------------------------------------------------------------------------------------------- 
  --功能：按时间范围获取费用单据 
  --入参：json格式 
  --  input  
  --  query_type  N  1  查询范围:0-返回剩余金额;1-仅返回原始结算信息
  --    occasion  N  1  结算场合:1-收费,2-预交(包含押金),3-结帐(暂无用),4-挂号,5-就诊卡,6-补充医保结算
  --    fee_nos  C    query_type=2时有效:单据号:结算场合=2时，为预交NO, 结算id未传入，该节点必传
  --出参：json格式 
  --  output 
  --    code  C  1  应答码：0-失败；1-成功
  --    message  C  1  应答消息：成功时返回成功信息，失败时返回具体的错误信息
  --    data  C    结算信息
  --      pati_info  C    病人信息
  --        pati_id  N  1  病人ID
  --        pati_pageid  N    主页ID
  --        pati_name  C  1  姓名
  --        pati_sex  C  1  性别
  --        pati_age  C  1  年龄
  --        outpatient_num  C  1  门诊号
  --        inpatient_num  C  1  住院号
  --        insurance_type  N  1  险类
  --      balance_info  C    结算信息
  --        invoice_no  C  1  发票号
  --        balance_oldid  N  1  原结算ID
  --        create_time  C  1  收费时间:yyyy-mm-dd hh:mi:ss
  --        total  N  1  结算总额
  --        balance_unit  N  1  是否合约单位结算
  --        balance_type  N  1  预交时，预交类别:1-门诊;2-住院 ;3-门诊和住院;结帐时：结帐类型:1-门诊;2-住院 ;3-门诊和住院;
  --        start_einv  N  1  是否启用电子票据
  ------------------------------------------------------------------------------------------------- 
  j_Input PLJson;
  j_Json  PLJson;

  n_场合     Number(2);
  n_查询类型 Number(2);
  v_Nos      Varchar2(32767);
  v_Output   Varchar2(32767);

  n_结帐id       门诊费用记录.结帐id%Type;
  n_结帐金额     门诊费用记录.结帐金额%Type;
  n_是否电子票据 病人预交记录.是否电子票据%Type;
  v_收款时间     Varchar2(30);

  Cursor c_结算信息 Is(
    Select Max(Decode(a.记录状态, 2, 0, a.Id)) As ID, a.病人id, a.主页id, Sum(a.金额) As 结帐金额, Max(a.预交电子票据) As 是否电子票据,
           Max(a.姓名) As 姓名, Max(a.性别) As 性别, Max(a.年龄) As 年龄, Max(a.住院号) As 住院号, Max(a.门诊号) As 门诊号, Max(m.险类) As 险类,
           To_Char(Max(Decode(a.记录状态, 2, To_Date(Null), a.收款时间)), 'yyyy-mm-dd hh24:mi:ss') As 收费时间, Max(a.预交类别) As 结帐类型
    From 病人预交记录 A, 保险结算记录 M
    Where a.No = '-' And a.记录性质 = 1 And a.Id = m.记录id(+) And m.性质(+) = 3);
  r_结算信息 c_结算信息%RowType;

  Type Ty_Einvoce Is Ref Cursor;
  c_Balanceinfo Ty_Einvoce; --动态游标变量

Begin
  j_Input    := PLJson(Json_In);
  j_Json     := j_Input.Get_Pljson('input');
  n_场合     := Nvl(j_Json.Get_Number('occasion'), 0);
  n_查询类型 := Nvl(j_Json.Get_Number('query_type'), 0);
  v_Nos      := j_Json.Get_String('fee_nos');

  If n_查询类型 = 1 Then
    --仅返回原始的结算信息
    If Nvl(n_场合, 0) = 2 Then
      Select a.Id As 结帐id, Max(a.金额) As 结帐金额, Max(a.预交电子票据) As 是否电子票据,
             To_Char(Max(a.收款时间), 'yyyyy-mm-dd hh24:mi:ss') As 收费时间
      Into n_结帐id, n_结帐金额, n_是否电子票据, v_收款时间
      From 病人预交记录 A
      Where a.No = v_Nos And a.记录性质 = 1 And a.记录状态 In (1, 3);
    Elsif Nvl(n_场合, 0) = 4 Or Nvl(n_场合, 0) = 1 Then
    
      Select Max(a.结帐id) As 结帐id, Max(a.结帐金额) As 结帐金额, Max(b.是否电子票据) As 是否电子票据,
             To_Char(Max(a.登记时间), 'yyyyy-mm-dd hh24:mi:ss') As 收费时间
      Into n_结帐id, n_结帐金额, n_是否电子票据, v_收款时间
      From (Select Max(结帐id) As 结帐id, Sum(结帐金额) As 结帐金额, Max(登记时间) As 登记时间
             From 门诊费用记录
             Where 记录性质 = n_场合 And NO = v_Nos And 记录状态 In (1, 3)) A, 病人预交记录 B
      Where a.结帐id = b.结帐id;
    Elsif Nvl(n_场合, 0) = 5 Then
      Select Max(a.结帐id) As 结帐id, Max(a.结帐金额) As 结帐金额, Max(b.是否电子票据) As 是否电子票据,
             To_Char(Max(a.登记时间), 'yyyyy-mm-dd hh24:mi:ss') As 收费时间
      Into n_结帐id, n_结帐金额, n_是否电子票据, v_收款时间
      From (Select Max(结帐id) As 结帐id, Sum(结帐金额) As 结帐金额, Max(登记时间) As 登记时间
             From 住院费用记录
             Where 记录性质 = 5 And NO = v_Nos And 记录状态 In (1, 3)) A, 病人预交记录 B
      Where a.结帐id = b.结帐id;
    Else
      Json_Out := zlJsonOut('场合节点传入值不对!');
      Return;
    End If;
    --结算信息
    v_Output := v_Output || '"balance_info":';
    v_Output := v_Output || '{"balance_oldid":' || zlJsonStr(n_结帐id, 1);
    v_Output := v_Output || ',"create_time":"' || zlJsonStr(v_收款时间) || '"';
    v_Output := v_Output || ',"start_einv":' || zlJsonStr(n_是否电子票据, 1);
    v_Output := v_Output || ',"total":' || zlJsonStr(n_结帐金额, 1);
    --以下暂不返回，无作用,需要时再加
    v_Output := v_Output || ',"balance_type":0';
    v_Output := v_Output || ',"balance_unit":0';
    v_Output := v_Output || '}';
  
    Json_Out := '{"output":{"code":1,"message":"成功","data":{' || v_Output || '}}}';
  
    Return;
  End If;
  If Nvl(n_场合, 0) = 2 Then
    Open c_Balanceinfo For
      Select Max(Decode(a.记录状态, 2, 0, a.Id)) As ID, a.病人id, a.主页id, Sum(a.金额) As 结帐金额, Max(a.预交电子票据) As 是否电子票据,
             Max(a.姓名) As 姓名, Max(a.性别) As 性别, Max(a.年龄) As 年龄, Max(a.住院号) As 住院号, Max(a.门诊号) As 门诊号, Max(m.险类) As 险类,
             To_Char(Max(Decode(a.记录状态, 2, To_Date(Null), a.收款时间)), 'yyyy-mm-dd hh24:mi:ss') As 收费时间,
             Max(a.预交类别) As 结帐类型
      From 病人预交记录 A, 保险结算记录 M
      Where a.No = v_Nos And a.记录性质 = 1 And a.Id = m.记录id(+) And m.性质(+) = 3
      Group By a.Id, a.No, a.病人id, a.主页id;
  
  Elsif Nvl(n_场合, 0) = 5 Then
    Open c_Balanceinfo For
      Select a.结帐id As ID, a.病人id, a.主页id, Max(a.结帐金额) As 结帐金额, Max(b.是否电子票据) As 是否电子票据, Max(a.姓名) As 姓名,
             Max(a.性别) As 性别, Max(a.年龄) As 年龄, Max(b.住院号) As 住院号, Max(Nvl(a.门诊号, b.门诊号)) As 门诊号, Max(m.险类) As 险类,
             Max(a.收费时间) As 收费时间, 1 As 结帐类型
      From (Select Max(Decode(a.记录状态, 2, 0, 11, 0, a.结帐id)) As 结帐id, Max(a.病人id) As 病人id, Max(a.主页id) As 主页id,
                    Max(a.姓名) As 姓名, Max(a.性别) As 性别, Max(a.年龄) As 年龄, Max(a.标识号) As 门诊号, Sum(a.结帐金额) As 结帐金额,
                    To_Char(Max(a.登记时间), 'yyyy-mm-dd hh24:mi:ss') As 收费时间
             From 住院费用记录 A
             Where a.No = v_Nos And 记录性质 = 5) A, 病人预交记录 B, 保险结算记录 M
      Where a.结帐id = b.结帐id And a.结帐id = m.记录id(+) And m.性质(+) = 1
      Group By a.结帐id, a.病人id, a.主页id;
  Elsif Nvl(n_场合, 0) = 4 Then
    --挂号
  
    Open c_Balanceinfo For
      Select a.结帐id As ID, a.病人id, a.主页id, Max(a.结帐金额) As 结帐金额, Max(b.是否电子票据) As 是否电子票据, Max(a.姓名) As 姓名,
             Max(a.性别) As 性别, Max(a.年龄) As 年龄, Max(b.住院号) As 住院号, Max(Nvl(a.门诊号, b.门诊号)) As 门诊号, Max(m.险类) As 险类,
             Max(a.收费时间) As 收费时间, 1 As 结帐类型
      From (Select Max(Decode(a.记录状态, 2, 0, 11, 0, a.结帐id)) As 结帐id, Max(a.病人id) As 病人id, Max(a.主页id) As 主页id,
                    Max(a.姓名) As 姓名, Max(a.性别) As 性别, Max(a.年龄) As 年龄, Max(a.标识号) As 门诊号, Sum(a.结帐金额) As 结帐金额,
                    To_Char(Max(a.登记时间), 'yyyy-mm-dd hh24:mi:ss') As 收费时间
             From 门诊费用记录 A
             Where a.No = v_Nos And 记录性质 = 4) A, 病人预交记录 B, 保险结算记录 M
      Where a.结帐id = b.结帐id And a.结帐id = m.记录id(+) And m.性质(+) = 1
      Group By a.结帐id, a.病人id, a.主页id;
  Elsif Nvl(n_场合, 0) = 1 Then
    --收费
    --注意：一次结算的单据号必须全传入
    Open c_Balanceinfo For
      Select a.结帐id As ID, a.病人id, a.主页id, Max(a.结帐金额) As 结帐金额, Max(b.是否电子票据) As 是否电子票据, Max(a.姓名) As 姓名,
             Max(a.性别) As 性别, Max(a.年龄) As 年龄, Max(b.住院号) As 住院号, Max(Nvl(a.门诊号, b.门诊号)) As 门诊号, Max(m.险类) As 险类,
             Max(a.收费时间) As 收费时间, 1 As 结帐类型
      From (Select /*+ cardinality(b, 10) */
              Max(Decode(a.记录状态, 2, 0, 11, 0, a.结帐id)) As 结帐id, Max(a.病人id) As 病人id, Max(a.主页id) As 主页id, Max(a.姓名) As 姓名,
              Max(a.性别) As 性别, Max(a.年龄) As 年龄, Max(a.标识号) As 门诊号, Sum(a.结帐金额) As 结帐金额,
              To_Char(Max(a.登记时间), 'yyyy-mm-dd hh24:mi:ss') As 收费时间
             From 门诊费用记录 A, Table(f_Str2List(v_Nos)) B
             Where a.No = b.Column_Value And Mod(记录性质, 10) = 1) A, 病人预交记录 B, 保险结算记录 M
      Where a.结帐id = b.结帐id And a.结帐id = m.记录id(+) And m.性质(+) = 1
      Group By a.结帐id, a.病人id, a.主页id;
    --ELSIF nvl(n_场合,0) = 3 THEN 
    --暂不支持
  Else
    Json_Out := zlJsonOut('场合节点传入值不对!');
    Return;
  End If;

  Fetch c_Balanceinfo
    Into r_结算信息;

  If c_Balanceinfo %NotFound Then
    Close c_Balanceinfo;
    Json_Out := zlJsonOut('未找到原始结算(NO=' || v_Nos || ')的电子票据，请检查!');
    Return;
  End If;

  v_Output := v_Output || '{"pati_id":' || zlJsonStr(r_结算信息.病人id, 1);
  v_Output := v_Output || ',"pati_pageid":' || zlJsonStr(r_结算信息.主页id, 1);
  v_Output := v_Output || ',"pati_name":"' || zlJsonStr(r_结算信息.姓名) || '"';
  v_Output := v_Output || ',"pati_sex":"' || zlJsonStr(r_结算信息.性别) || '"';
  v_Output := v_Output || ',"pati_age":"' || zlJsonStr(r_结算信息.年龄) || '"';
  v_Output := v_Output || ',"outpatient_num":"' || zlJsonStr(r_结算信息.门诊号) || '"';
  v_Output := v_Output || ',"inpatient_num":"' || zlJsonStr(r_结算信息.住院号) || '"';
  v_Output := v_Output || ',"insurance_type":' || zlJsonStr(r_结算信息.险类, 1);

  v_Output := v_Output || '}';

  v_Output := '"pati_info":' || v_Output;
  --结算信息
  v_Output := v_Output || ',"balance_info":';
  v_Output := v_Output || '{"balance_oldid":' || zlJsonStr(r_结算信息.Id, 1);
  v_Output := v_Output || ',"create_time":"' || zlJsonStr(r_结算信息.收费时间) || '"';
  v_Output := v_Output || ',"total":' || zlJsonStr(r_结算信息.结帐金额, 1);
  v_Output := v_Output || ',"start_einv":' || zlJsonStr(r_结算信息.是否电子票据, 1);
  v_Output := v_Output || ',"balance_type":' || zlJsonStr(r_结算信息.结帐类型, 1);
  v_Output := v_Output || ',"balance_unit":0';
  v_Output := v_Output || '}';

  Json_Out := '{"output":{"code":1,"message":"成功","data":{' || v_Output || '}}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getbalanceinfo;
/


Create Or Replace Procedure Zl_Exsesvr_Billverify_Check
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：审核一张住院记帐划价单的合法性检查
  --入参：Json_In:格式
  --  input
  --    fee_nos                       C   1   单据号，允许传入多张,多张时，用逗 \号分隔
  --    serials_num                   C   1   序号,多个用逗号分离,为空为所有,当fee_nos传入多张单据时，该项无效
  --    pati_list[]病人信息，仅审核这些病人的费用
  --      pati_id                     N   1   病人ID
  --      fee_audit_status            N   1   费用审核标志:0或空-未审核;1-已审核或开始审核(结合参数:病人审核方式来控制);2-完成审核,结合结帐权限[禁止未审核病人结帐]进行管理控制
  --      si_inp_status               N   1   住院状态:0-正常住院;1-尚未入科;2-正在转科或正在转病区;3-已预出院
  --出参: Json_Out,格式如下
  --  output
  --    code                          N   1   应答吗：0-失败；1-成功
  --    message                       C   1   应答消息：失败时返回具体的错误信息
  --    item_list[]
  --      rcp_no                      C   1   处方单号
  --      stuff_rcpdtl_ids            C   1   卫材处方明细IDs:卫生材料所涉及的费用ids
  --      drug_rcpdtl_ids             C   1   药品处方明细IDs:药品所涉及的费用ids
  --      autosendstuff_rcpdtl_ids    C   1   发料处方明细IDs:自动发放卫生材料所涉及的费用IDs
  ---------------------------------------------------------------------------
Begin
  Zl_住院记帐记录_Verify_Check(Json_In, Json_Out);
End Zl_Exsesvr_Billverify_Check;
/

Create Or Replace Procedure Zl_Exsesvr_Updrgstarrangement
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  --功能：调整号源、有效的安排、有效的出诊记录中的医生姓名。
  --入参：n_操作方式:1-修改姓名,2-停用人员,3-启用人员
  --      d_撤档时间:停用和启用时传入，启用时传入原撤档时间
  --说明：该过程供人员姓名调整或人员停用/启用时调用，同步调整挂号安排
  v_Para     Varchar2(200);
  n_挂号模式 Number(2);
  j_Input    PLJson;
  j_Json     PLJson;

  n_人员id   挂号安排.医生id%Type;
  n_操作方式 Number(2);
  v_人员姓名 人员表.姓名%Type;
  d_开始时间 临床出诊记录.开始时间%Type;
  d_撤档时间 人员表.撤档时间%Type;
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_人员id   := j_Json.Get_Number('rgst_dr_id');
  n_操作方式 := j_Json.Get_Number('oper_type');
  d_撤档时间 := To_Date(j_Json.Get_String('revoke_time'), 'YYYY-MM-DD hh24:mi:ss');
  If n_操作方式 = 2 And d_撤档时间 Is Null Then
    d_撤档时间 := Sysdate;
  End If;

  Begin
    Select a.姓名
    Into v_人员姓名
    From 人员表 A, 人员性质说明 B
    Where a.Id = n_人员id And a.Id = b.人员id And b.人员性质 = '医生';
  Exception
    When Others Then
      --无需处理,退出
      Json_Out := zlJsonOut('成功', 1);
      Return;
  End;

  v_Para     := zl_GetSysParameter(256);
  n_挂号模式 := Substr(v_Para, 1, 1);

  If n_挂号模式 = 1 Then
    --出诊表模式
    If n_操作方式 = 1 Then
      --修改
      Update 临床出诊号源 Set 医生姓名 = v_人员姓名 Where 医生id = n_人员id;
      Update 临床出诊安排 Set 医生姓名 = v_人员姓名 Where 医生id = n_人员id And 终止时间 > Sysdate;
      Update 临床出诊记录 Set 医生姓名 = v_人员姓名 Where 医生id = n_人员id And 出诊日期 >= Trunc(Sysdate);
      Update 临床出诊记录 Set 替诊医生姓名 = v_人员姓名 Where 替诊医生id = n_人员id And 出诊日期 >= Trunc(Sysdate);
    Elsif n_操作方式 = 2 Then
      --停用
      Update 临床出诊号源
      Set 撤档时间 = d_撤档时间
      Where 医生id = n_人员id And (撤档时间 Is Null Or 撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'));
    
      --将当前时间以后的所有出诊记录停诊
      For c_记录 In (Select ID, 开始时间, 终止时间
                   From 临床出诊记录
                   Where 医生id = n_人员id And 出诊日期 >= Trunc(d_撤档时间) - 1 And 终止时间 > d_撤档时间 And 上班时段 Is Not Null) Loop
      
        If c_记录.开始时间 < d_撤档时间 Then
          d_开始时间 := d_撤档时间;
        Else
          d_开始时间 := c_记录.开始时间;
        End If;
        Zl_临床出诊记录_Stopvisit(c_记录.Id, d_开始时间, c_记录.终止时间, '人员停用', zl_UserName, d_撤档时间, 0, 1);
      End Loop;
    Elsif n_操作方式 = 3 Then
      --启用
      For c_记录 In (Select a.Id
                   From 临床出诊记录 A, 临床出诊号源 B
                   Where a.号源id = b.Id And b.撤档时间 = d_撤档时间 And b.医生id = n_人员id And a.出诊日期 >= Trunc(Sysdate) - 1 And
                         a.终止时间 > Sysdate And a.上班时段 Is Not Null) Loop
      
        Zl_临床出诊记录_Stopvisit(c_记录.Id, Null, Null, Null, zl_UserName, Sysdate, 1, 1);
      End Loop;
    
      Update 临床出诊号源
      Set 撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd')
      Where 医生id = n_人员id And 撤档时间 = d_撤档时间;
    
      --生成出诊记录
      Zl1_Auto_Buildingregisterplan;
    End If;
  Else
    --挂号排班模式
    If n_操作方式 = 1 Then
      --修改
      Update 挂号安排 Set 医生姓名 = v_人员姓名 Where 医生id = n_人员id And (终止时间 Is Null Or 终止时间 > Sysdate);
      Update 挂号安排计划
      Set 医生姓名 = v_人员姓名
      Where 医生id = n_人员id And (失效时间 Is Null Or 失效时间 > Sysdate);
    Elsif n_操作方式 = 2 Then
      --停用
      Update 挂号安排 Set 停用日期 = d_撤档时间 Where 医生id = n_人员id And (终止时间 Is Null Or 终止时间 > Sysdate);
      Update 挂号安排计划
      Set 失效时间 = d_撤档时间
      Where 医生id = n_人员id And (失效时间 Is Null Or 失效时间 > Sysdate);
    Elsif n_操作方式 = 3 Then
      --启用，不处理
      Null;
    End If;
  End If;
  Json_Out := zlJsonOut('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Updrgstarrangement;
/

CREATE OR REPLACE Procedure Zl_Exsesvr_Bulidregistprice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --根据医嘱ID查询是否在费用表存在记录
  ---------------------------------------------------------------------------
  --input      根据挂号单构键收费划价单
  --  rgst_no         C  1  挂号单号
  --  pati_id         N     病人ID
  --  pati_name       C     姓名
  --  pati_sex        C     性别
  --  pati_age        C     年龄
  --  pati_idcard     C     身份证号
  --  birth_date      C     出生日期
  --  rgst_dept_id     N  1  挂号科室ID
  --  rgst_dr          C  1  医生姓名
  --  operator_name    C  1  操作员姓名
  --  site_no          C    站点
  --  rgst_visitinfo      病人就诊信息
  --    outp_room_name  C    接诊科室
  --    emg_sign        N    急诊标志
  --    revisit_sign    N    回诊标志
  --    exe_time        C    执行时间
  --  出参      json
  --output
  --  code    N  1  应答码：0-失败；1-成功
  --  message  C  1  应答消息：成功时返回成功信息，失败时返回具体的错误信息
  --  fee_no  C  1  划价单号
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  o_Json1 PLJson;

  n_病人id     病人挂号记录.病人id%Type;
  v_挂号单     病人挂号记录.No%Type;
  n_科室id     病人挂号记录.执行部门id%Type;
  v_医生姓名   病人挂号记录.执行人%Type;
  v_操作员姓名 病人挂号记录.操作员姓名%Type;
  v_诊室       病人挂号记录.诊室%Type;
  n_急诊标志   病人挂号记录.急诊%Type;
  n_回诊标志   Integer;
  d_执行时间   病人挂号记录.执行时间%Type;
  v_站点       Varchar2(100);
  v_划价单     病人挂号记录.收费单%Type;
  v_姓名       病人挂号记录.姓名%Type;
  v_性别       病人挂号记录.性别%Type;
  v_年龄       病人挂号记录.年龄%Type;
  d_出生日期   Date;
  v_身份证号   Varchar2(18);
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_挂号单     := j_Json.Get_String('rgst_no');
  n_科室id     := j_Json.Get_Number('rgst_dept_id');
  v_医生姓名   := j_Json.Get_String('rgst_dr');
  v_操作员姓名 := j_Json.Get_String('operator_name');
  v_站点       := j_Json.Get_String('site_no');
  n_病人id     := j_Json.Get_Number('pait_id');
  v_姓名       := j_Json.Get_String('pati_name');
  v_性别       := j_Json.Get_String('pati_sex');
  v_年龄       := j_Json.Get_String('pati_age');
  d_出生日期   := To_Date(j_Json.Get_String('birth_date'), 'YYYY-MM-DD hh24:mi:ss');
  v_身份证号   := j_Json.Get_String('pati_idcard');

  If Nvl(n_科室id, 0) = 0 Then
    n_科室id := Null;
  End If;
  Select Zl_Exse_Nextno(12, Null) Into v_划价单 From Dual;

  Zl_门诊划价记录_Buliding_s(v_挂号单, v_划价单, n_病人id, v_姓名, v_性别, v_年龄, d_出生日期, v_身份证号, n_科室id, v_医生姓名, v_操作员姓名, v_站点);

  o_Json1 := j_Json.Get_Pljson('rgst_visitinfo');
  If Not o_Json1 Is Null Then
    n_科室id   := j_Json.Get_Number('exe_deptid');
    v_诊室     := o_Json1.Get_String('outp_room_name');
    n_急诊标志 := o_Json1.Get_Number('emg_sign');
    n_回诊标志 := o_Json1.Get_Number('revisit_sign');
    d_执行时间 := To_Date(o_Json1.Get_String('exe_time'), 'yyyy-mm-dd hh24:mi:ss');

    Zl_病人接诊_s(n_病人id, v_挂号单, n_科室id, v_医生姓名, v_诊室, n_急诊标志, n_回诊标志, d_执行时间);
  End If;

  Json_Out := '{"output":{"code":1,"message":"成功","fee_no":"' || v_划价单 || '"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Bulidregistprice;
/

CREATE OR REPLACE Procedure Zl_Exsesvr_Checknoischarge
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --功能:检查单据是否已收费，目前用于挂号划价单检查
  ---------------------------------------------------------------------------
  --input
  --  fee_no          C  1  单据号
  --  checkCharge      N    检查划价单是否收费
  --  rgst_dept_id    N    执行部门ID
  --  rgst_dr          C    执行人
  --output
  --  code        N  1  应答码：0-失败；1-成功
  --  message      C  1  应答消息：成功时返回成功信息，失败时返回具体的错误信息
  --  fee_status  C  1  单据状态：-1-未找到对应的挂号单据,0-未收费;1-挂号单已收;2-还未产生划价记录; 3-挂号单对应的收费划价单已全收费(存在多张划价单时，必须全收的);4-挂号单对应的划价单存在部分收费)
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_单据号       门诊费用记录.No%Type;
  n_记录性质     病人挂号记录.记录性质%Type;
  v_收费单       病人挂号记录.收费单%Type;
  n_取号标志     病人挂号记录.取号标志%Type;
  n_结帐id       门诊费用记录.结帐id%Type;
  n_Min结帐id    门诊费用记录.结帐id%Type;
  n_Max结帐id    门诊费用记录.结帐id%Type;
  n_执行部门id   门诊费用记录.执行部门id%Type;
  v_执行人       门诊费用记录.执行人%Type;
  n_Checkcharge  Number(2); --检查单据是否收费
  n_Count        Number(2);
  n_执行部门id_b Number(2);
  n_执行人_b     Number(2);
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_单据号      := j_Json.Get_String('fee_no');
  n_Checkcharge := j_Json.Get_Number('checkcharge');

  If j_Json.Exist('input.rgst_dept_id') Then
    n_执行部门id   := j_Json.Get_Number('rgst_dept_id ');
    n_执行部门id_b := 1;
  End If;
  If j_Json.Exist('input.rgst_dr') Then
    v_执行人   := j_Json.Get_String('rgst_dr');
    n_执行人_b := 1;
  End If;

  Select Count(1), Max(a.收费单) As 收费单, Max(a.取号标志) As 取号标志, Max(b.结帐id) As 结帐id, Max(a.记录性质) As 记录性质
  Into n_Count, v_收费单, n_取号标志, n_结帐id, n_记录性质
  From 病人挂号记录 A, 门诊费用记录 B
  Where a.No = v_单据号 And a.No = b.No And b.记录性质 = 4 And b.记录状态 In (0, 1, 3);

  If n_Count = 0 Then
    Json_Out := '{"output":{"code":0,"message":"未找到对应的挂号单据","fee_status":-1}}';
    Return;
  End If;

  If v_收费单 Is Null Then
    If Nvl(n_取号标志, 0) = 1 Then
      --免挂号模式，未生成划价单
      Json_Out := '{"output":{"code":0,"message":"还未产生划价记录","fee_status":2}}';
    Elsif Nvl(n_结帐id, 0) = 0 Or Nvl(n_记录性质, 0) = 2 Then
      --未结账或预约未接收
      Json_Out := '{"output":{"code":1,"message":"成功","fee_status":0}}';
    Else
      --已收费
      Json_Out := '{"output":{"code":0,"message":"挂号单已收费","fee_status":1}}';
    End If;
    Return;
  End If;

  If Nvl(n_Checkcharge, 0) = 0 Then
    Json_Out := '{"output":{"code":1,"message":"成功","fee_status":0}}';
  End If;

  Select /* +cardinality(b,10) */
   Count(1), Min(Decode(结帐id, Null, 0, 结帐id)), Max(结帐id)
  Into n_Count, n_Min结帐id, n_Max结帐id
  From 门诊费用记录 A, Table(f_Str2List(v_收费单)) B
  Where a.记录性质 = 1 And a.记录状态 In (0, 1, 3) And a.No = b.Column_Value And
        a.执行部门id = Decode(n_执行部门id_b, 1, n_执行部门id, a.执行部门id) And a.执行人 = Decode(n_执行人_b, 1, v_执行人, a.执行人);

  If n_Count = 0 Then
    --没有划价单
    Json_Out := '{"output":{"code":0,"message":"还未产生划价记录","fee_status":2}}';

  Elsif Nvl(n_Max结帐id, 0) = 0 Then
    --未收费
    Json_Out := '{"output":{"code":1,"message":"成功","fee_status":0,"fee_status":"' || v_收费单 || '"}}';

  Elsif Nvl(n_Min结帐id, 0) = 0 And Nvl(n_Max结帐id, 0) > 0 Then
    --未全收费
    Json_Out := '{"output":{"code":0,"message":"挂号单对应的收费划价单已全收费","fee_status":3}}';

  Else
    --全收费
    Json_Out := '{"output":{"code":0,"message":"挂号单对应的划价单存在部分收费","fee_status":4}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checknoischarge;
/


CREATE OR REPLACE Procedure Zl_Exsesvr_Checkdelregistprice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --功能:删除挂号划价单前检查单据是否允许删除
  ---------------------------------------------------------------------------
  --input
  --  fee_no          C  1  单据号
  --  checkCharge     N    检查划价单是否收费
  --  rgst_dept_id    N    执行部门ID
  --  rgst_dr         C    执行人
  --output
  --  code        N  1  应答码：0-失败；1-成功
  --  message      C  1  应答消息：成功时返回成功信息，失败时返回具体的错误信息
  --  fee_status  C  1  单据状态：0-正常划价状态;1-未找到挂号单;2-未生成划价单;3-未找到符合条件的划价单;4-存在已经收费的单据
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_单据号       门诊费用记录.No%Type;
  v_收费单       病人挂号记录.收费单%Type;
  n_执行部门id   门诊费用记录.执行部门id%Type;
  v_执行人       门诊费用记录.执行人%Type;
  n_Count        Number(2);
  n_Code         Number(2);
  n_执行部门id_b Number(2);
  n_执行人_b     Number(2);
  v_Input        Varchar2(100);
  v_Output       Varchar2(100);
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_单据号 := j_Json.Get_String('fee_no');

  If j_Json.Exist('rgst_dept_id') Then
    n_执行部门id   := j_Json.Get_Number('rgst_dept_id');
    n_执行部门id_b := 1;
  End If;
  If j_Json.Exist('rgst_dr') Then
    v_执行人   := j_Json.Get_String('rgst_dr');
    n_执行人_b := 1;
  End If;

  Select Count(1), Max(a.收费单) As 收费单
  Into n_Count, v_收费单
  From 病人挂号记录 A
  Where a.No = v_单据号 And a.记录状态 In (0, 1, 3);

  If n_Count = 0 Then
    Json_Out := '{"output":{"code":0,"message":"未找到挂号单","fee_status":1}}';
    Return;
  End If;

  If v_收费单 Is Null Then
    Json_Out := '{"output":{"code":0,"message":"未生成划价单","fee_status":2}}';
    Return;
  End If;

  n_Count := 0;
  For c_Price In (Select /* +cardinality(b,10) */
                   a.No, Max(结帐id) As 结帐id
                  From 门诊费用记录 A, Table(f_Str2List(v_收费单)) B
                  Where a.记录性质 = 1 And a.记录状态 In (0, 1, 3) And a.No = b.Column_Value And
                        a.执行部门id = Decode(n_执行部门id_b, 1, n_执行部门id, a.执行部门id) And
                        a.执行人 = Decode(n_执行人_b, 1, v_执行人, a.执行人)
                  Group By a.No) Loop

    If Nvl(c_Price.结帐id, 0) > 0 Then
      Json_Out := '{"output":{"code":0,"message":"存在已经收费的单据","fee_status":4}}';
      Return;
    End If;

    v_Input := '{"input":{"fee_no":"' || c_Price.No || '"}}';
    Zl_门诊划价记录_Delete_Check(v_Input, v_Output);
    j_Json  := PLJson();
    j_Json  := PLJson(v_Output);
    j_Json  := j_Input.Get_Pljson('output');
    n_Code  := j_Json.Get_Number('code');
    If Nvl(n_Code, 0) = 0 Then
      Json_Out := '{"output":{"code":0,"message":"未找到符合条件的划价单","fee_status":3}}';
      Return;
    End If;
    n_Count := n_Count + 1;
  End Loop;
  If n_Count = 0 Then
    Json_Out := '{"output":{"code":0,"message":"未找到符合条件的划价单","fee_status":3}}';
  Else
    Json_Out := '{"output":{"code":1,"message":"成功","fee_status":0}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkdelregistprice;
/


CREATE OR REPLACE Procedure Zl_Exsesvr_Delregistprice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --功能:检查挂号划价单是否允许删除
  ---------------------------------------------------------------------------
  --input
  --  fee_no              C  1  单据号
  --  rgst_dept_id        N     科室ID
  --  rgst_dr             C     医生姓名
  --  rgst_visitinfo      N
  --     exe_deptid       N   执行部门ID
  --     exetr            C   执行人
  --     referral_sign    N   是否转诊: 0-未转诊  1-转诊
  --     referral_deptid  N   转诊科室ID
  --     referral_doctor  C   转诊医生
  --output
  --  code        N 1 应答码：0-失败；1-成功
  --  message     C 1 应答消息：成功时返回成功信息，失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  o_Json PLJson;

  v_单据号       门诊费用记录.No%Type;
  n_病人id       病人挂号记录.病人id%Type;
  v_收费单       病人挂号记录.收费单%Type;
  n_执行部门id   病人挂号记录.执行部门id%Type;
  v_执行人       病人挂号记录.执行人%Type;
  n_转诊科室id   病人挂号记录.转诊科室id%Type;
  v_转诊医生     病人挂号记录.转诊诊室%Type;
  n_转诊标志     Number(2);
  n_执行部门id_b Number(2);
  n_执行人_b     Number(2);
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_单据号 := j_Json.Get_String('fee_no');

  If j_Json.Exist('rgst_dept_id') Then
    n_执行部门id   := j_Json.Get_Number('rgst_dept_id');
    n_执行部门id_b := 1;
  End If;
  If j_Json.Exist('rgst_dr') Then
    v_执行人   := j_Json.Get_String('rgst_dr');
    n_执行人_b := 1;
  End If;

  Select Max(病人id), Max(a.收费单)
  Into n_病人id, v_收费单
  From 病人挂号记录 A
  Where a.No = v_单据号 And a.记录状态 In (0, 1, 3);
  If v_收费单 Is Null Then
    Json_Out := zlJsonOut('不存在挂号划价单！');
    Return;
  End If;

  For c_Price In (Select /* +cardinality(b,10) */
                  Distinct a.No
                  From 门诊费用记录 A, Table(f_Str2List(v_收费单)) B
                  Where a.记录性质 = 1 And a.记录状态 In (0, 1, 3) And a.No = b.Column_Value And
                        a.执行部门id = Decode(n_执行部门id_b, 1, n_执行部门id, a.执行部门id) And
                        a.执行人 = Decode(n_执行人_b, 1, v_执行人, a.执行人)) Loop
    Zl_门诊划价记录_Delete_s(c_Price.No);
  End Loop;

  o_Json := j_Json.Get_Pljson('rgst_visitinfo');
  If Not o_Json Is Null Then
    n_执行部门id := o_Json.Get_Number('exe_deptid');
    v_执行人     := o_Json.Get_String('exetr');
    n_转诊标志   := Nvl(o_Json.Get_Number('referral_sign'), 0);
    n_转诊科室id := o_Json.Get_Number('referral_dept_id');
    v_转诊医生   := o_Json.Get_String('referral_dr');

    If n_执行部门id = 0 Then
      n_执行部门id := Null;
    End If;
    If n_转诊科室id = 0 Then
      n_转诊科室id := Null;
    End If;

    Zl_病人接诊_Cancel_s(n_病人id, v_单据号, n_执行部门id, v_执行人, n_转诊标志, v_转诊医生, n_转诊科室id);
  End If;

  Json_Out := zlJsonOut('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Delregistprice;
/

Create Or Replace Procedure Zl_Exsesvr_Chkpatichangenurse
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能:护理等级调整检查
  --入参：Json_In:格式
  --input
  --      pati_id           N 1 病人id
  --      pati_pageid       N 1 主页ID
  --      create_time       C 1 登记时间
  --出参  json
  --output
  --     code               N 1 应答码：0-失败；1-成功
  --     message            C 1 应答消息： 失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_病人id   住院费用记录.病人id%Type;
  n_主页id   住院费用记录.主页id%Type;
  d_开始时间 Date;
Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_病人id   := j_Json.Get_Number('pati_id');
  n_主页id   := j_Json.Get_Number('pati_pageid');
  d_开始时间 := To_Date(j_Json.Get_String('create_time'), 'YYYY-MM-DD HH24:MI:SS');
  For r_Fee In (Select NO
                From 住院费用记录
                Where 病人id = n_病人id And 主页id = n_主页id And Mod(记录性质, 10) = 3 And 登记时间 >= d_开始时间 And 收费类别 = 'H'
                Group By NO, 序号, Mod(记录性质, 10)
                Having Sum(结帐金额) <> 0) Loop
    Json_Out := zlJsonOut('变动时间之后已有已结帐的自动记帐费用,不能更改护理等级！');
    Return;
  End Loop;
  Json_Out := zlJsonOut('成功', 1);
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Chkpatichangenurse;
/



Create Or Replace Procedure Zl_Exsesvr_Getpatireceivables
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：获取指定病人的应收款余额
  --入参：Json_In:格式
  --  input
  --   pati_id            N   病人id
  --出参: Json_Out,格式如下
  --  output
  --    code              N 1 应答码：0-失败；1-成功
  --    message           C 1 应答消息：失败时返回具体的错误信息
  --    fee_amrcvb        N 1 应收金额
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_Output   Varchar2(32767);
  n_病人id   病人预交记录.病人id%Type;
  n_应收余额 病人预交记录.冲预交%Type;
Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_病人id := j_Json.Get_Number('pati_id');

  If Nvl(n_病人id, 0) = 0 Then
    Json_Out := zlJsonOut('病人ID未传入。');
    Return;
  End If;

  Begin
    Select Nvl(a.应收款总额, 0) - Nvl(Sum(金额), 0)
    Into n_应收余额
    From (Select a.病人id, Sum(a.冲预交) 应收款总额
           From 病人预交记录 A, 结算方式 B
           Where a.病人id = n_病人id And a.结算方式 = b.名称 And b.应收款 = 1
           Group By 病人id) A, 病人缴款记录 B
    Where a.病人id = b.病人id(+) And b.记录状态(+) = 1
    Group By a.病人id, 应收款总额;
  Exception
    When Others Then
      n_应收余额 := 0;
  End;
  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '成功');
  zlJsonPutValue(v_Output, 'fee_amrcvb', n_应收余额, 1, 2);
  Json_Out := '{"output":' || v_Output || '}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getpatireceivables;
/

Create Or Replace Procedure Zl_Exsesvr_Getrgstinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --功能：获取预约挂号单据信息
  --入参：Json_In:格式
  --input
  --  rgst_no             C  1 挂号单
  --  appt_recv           N    预约接收
  -- 出参:
  --  output
  --    code                            N 1 应答吗：0-失败；1-成功
  --    message                         C 1 应答消息：失败时返回具体的错误信息
  --    close_account_type              N 1 挂号有效天数内的结算模式
  --    fee_list                        费用信息列表
  --       pati_id          N   1 病人ID
  --       outpatient_num   C   1 门诊号
  --       rgst_id          N   1 挂号id
  --       pati_name        C   1 姓名
  --       pati_sex         C   1 性别
  --       pati_age         C   1 年龄
  --       fee_category     C   1 费别
  --       num_category     C   1 号别
  --       mdlpay_mode_name C   1 付款方式
  --       overtime_sign    N   1 加班标志
  --       exe_deptid       N   1 执行部门id
  --       happen_time      C   1 发生时间
  --       appt_time        C   1 预约时间:yyyy-mm-dd hh24:mi:ss
  --       rgst_time        C   1 登记时间
  --       operator_code    C   1 操作员编号
  --       operator_name    C   1 操作员姓名
  --       appt_mode_name   C   1 预约方式
  --       fee_ampaid       N   1 实收金额
  --       fee_item_id      N   1 收费细目id
  --       outptyp_name     C   1 号类
  -------------------------------------------
  v_Output     Varchar2(32000);
  v_挂号单     病人挂号记录.No%Type;
  n_原结算模式 病人挂号记录.结算模式%Type;
  n_预约接收   Number(2);
  j_Input      Pljson;
  j_Json       Pljson;

Begin
  --解析入参
  j_Input := Pljson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_挂号单   := j_Json.Get_String('rgst_no');
  n_预约接收 := j_Json.Get_Number('appt_recv');

  --检查当前是否挂号信息是否存在
  For c_费用信息 In (Select Max(a.病人id) As 病人id, Max(c.Id) As 挂号id, Max(a.标识号) As 门诊号, Max(a.姓名) As 姓名, Max(a.性别) As 性别,
                        Max(a.年龄) As 年龄, Max(a.费别) As 费别, Max(Nvl(c.号别, Decode(a.序号, 1, a.计算单位, ''))) As 号别,
                        Max(a.加班标志) As 加班标志, Max(Decode(a.序号, 1, a.执行部门id, 0)) As 执行部门id,
                        To_Char(Max(a.发生时间), 'yyyy-mm-dd hh24:mi:ss') As 发生时间,
                        To_Char(Max(a.登记时间), 'yyyy-mm-dd hh24:mi:ss') As 登记时间, Max(a.操作员编号) As 操作员编号,
                        Max(a.操作员姓名) As 操作员姓名, Max(c.预约方式) As 预约方式,
                        To_Char(Max(Nvl(c.预约时间, a.发生时间)), 'yyyy-mm-dd hh24:mi:ss') As 预约时间,
                        Max(Nvl(c.挂号项目id, Decode(a.序号, 1, a.收费细目id, 0))) As 收费细目id, Sum(a.实收金额) As 挂号金额,
                        Max(b.名称) As 医疗付款名称, Max(c.号类) As 号类, Max(c.复诊) As 复诊, Max(c.急诊) As 急诊
                 From 门诊费用记录 A, 医疗付款方式 B, 病人挂号记录 C
                 Where a.No = v_挂号单 And a.No = c.No And a.记录性质 = 4 And a.记录状态 In (0, 1) And a.付款方式 = b.编码(+)) Loop
    Zljsonputvalue(v_Output, 'pati_id', c_费用信息.病人id, 1, 1);
    Zljsonputvalue(v_Output, 'outpatient_num', c_费用信息.门诊号);
    Zljsonputvalue(v_Output, 'rgst_id', c_费用信息.挂号id, 1);
    Zljsonputvalue(v_Output, 'pati_name', c_费用信息.姓名);
    Zljsonputvalue(v_Output, 'pati_sex', c_费用信息.性别);
    Zljsonputvalue(v_Output, 'pati_age', c_费用信息.年龄);
    Zljsonputvalue(v_Output, 'fee_category', c_费用信息.费别);
    Zljsonputvalue(v_Output, 'mdlpay_mode_name', c_费用信息.医疗付款名称);
    Zljsonputvalue(v_Output, 'num_category', c_费用信息.号别);
    Zljsonputvalue(v_Output, 'overtime_sign', c_费用信息.加班标志, 1);
    Zljsonputvalue(v_Output, 'exe_deptid', c_费用信息.执行部门id, 1);
    Zljsonputvalue(v_Output, 'happen_time', c_费用信息.发生时间);
    Zljsonputvalue(v_Output, 'rgst_time', c_费用信息.登记时间);
    Zljsonputvalue(v_Output, 'operator_code', c_费用信息.操作员编号);
    Zljsonputvalue(v_Output, 'operator_name', c_费用信息.操作员姓名);
    Zljsonputvalue(v_Output, 'appt_mode_name', c_费用信息.预约方式);
    Zljsonputvalue(v_Output, 'fee_item_id', c_费用信息.收费细目id, 1);
    Zljsonputvalue(v_Output, 'appt_time', Nvl(c_费用信息.预约时间, ''));
    Zljsonputvalue(v_Output, 'fee_ampaid', Nvl(c_费用信息.挂号金额, 0), 1);
    Zljsonputvalue(v_Output, 'outptyp_name', c_费用信息.号类);
    Zljsonputvalue(v_Output, 'revst_sign', c_费用信息.复诊, 1);
    Zljsonputvalue(v_Output, 'emg_sign', c_费用信息.急诊, 1, 2);
  
    If Nvl(c_费用信息.病人id, 0) <> 0 And Nvl(n_预约接收, 0) = 1 Then
      Zl_预约挂号接收_Check_s(v_挂号单, c_费用信息.病人id, Sysdate, c_费用信息.操作员姓名, c_费用信息.急诊, 0, 0, 0, n_原结算模式);
    End If;
  End Loop;

  If v_Output Is Null Then
    Json_Out := Zljsonout('挂号信息不存在，可能该挂号单已被他人处理。');
    Return;
  End If;

  Json_Out := '{"output":{"code":1,"message":"成功","close_account_type":' || Nvl(n_原结算模式, 0) || ',"fee_list":[' ||
              v_Output || ']}}';

Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getrgstinfo;
/

CREATE OR REPLACE Procedure Zl_Exsesvr_Rgstapptreceive
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --功能：预约挂号接收
  --入参：Json_In:格式
  --input
  --  rgst_no             C  1 挂号no
  --  outp_room_name     C  1 接诊诊室
  --  recv_time          C  1 接收时间
  --  prepay_pati_ids    C  1 冲预交病人ids
  --  pati_inhospital    N  1 冲预交病人ids

  --  checkout_id        N  1 结帐id
  --  cardtype_id        N  1 卡类别id
  --  pat_card_no        C  1 卡号
  --  trans_no           C  1 交易流水号
  --  trans_desc         C  1 交易说明
  --  recv_time          C  1 接收时间
  --  prepay_pati_ids    C  1 冲预交病人ids
  --  pati_id            N  1 病人id
  --  outpatient_num     C  1 门诊号
  --  reg_id             N  1 挂号id
  --  blnc_id            N  1 结帐ID
  --  relation_id        N  1 关联交易ID

  -- 出参:
  --  output
  --    code             N 1 应答吗：0-失败；1-成功
  --    message          C 1 应答消息：失败时返回具体的错误信息

  -------------------------------------------
  n_挂号id        病人挂号记录.Id%Type;
  v_挂号单        病人挂号记录.No%Type;
  n_病人id        病人挂号记录.病人id%Type;
  n_门诊号        病人挂号记录.门诊号%Type;
  v_姓名          病人挂号记录.姓名%Type;
  v_性别          病人挂号记录.性别%Type;
  v_年龄          病人挂号记录.年龄%Type;
  v_付款方式编码  医疗付款方式.编码%Type;
  v_费别          病人挂号记录.费别%Type;
  n_出诊记录id    病人挂号记录.出诊记录id%Type;
  n_结帐金额      门诊费用记录.结帐金额%Type;
  n_结帐id        门诊费用记录.结帐id%Type;
  v_划价单        门诊费用记录.No%Type;
  v_操作员编码    病人挂号记录.操作员编号%Type;
  v_操作员姓名    病人挂号记录.操作员姓名%Type;
  n_挂号项目id    病人挂号记录.挂号项目id%Type;
  n_接诊诊室      病人挂号记录.诊室%Type;
  d_接收时间      病人挂号记录.接收时间%Type;
  n_执行部门id    病人挂号记录.执行部门id%Type;
  v_号别          病人挂号记录.号别%Type;
  v_医生姓名      病人挂号记录.执行人%Type;
  n_医生id        人员表.Id%Type;
  n_预交id        病人预交记录.Id%Type;
  v_冲预交病人ids Varchar2(1000);
  v_结算方式      病人预交记录.结算方式%Type;
  n_卡类别id      病人预交记录.卡类别id%Type;
  v_支付卡号      病人预交记录.卡号%Type;
  n_挂号划价      Number(2);
  n_在院          Number(2);
  n_挂号生成队列  Number(2);
  o_Json          PLJson;
  j_Input         PLJson;
  j_Json          PLJson;

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_挂号单        := j_Json.Get_String('reg_no');
  n_接诊诊室      := j_Json.Get_String('outp_room_name');
  n_病人id        := j_Json.Get_Number('pati_id');
  n_门诊号        := j_Json.Get_String('outpatient_num');
  d_接收时间      := To_Date(j_Json.Get_String('recv_time'), 'YYYY-MM-DD hh24:mi:ss');
  v_冲预交病人ids := j_Json.Get_String('prepay_pati_ids');
  n_在院          := j_Json.Get_Number('pati_inhospital');
  n_结帐id        := j_Json.Get_Number('blnc_id');
  n_预交id        := j_Json.Get_Number('relation_id');
  v_操作员编码    := j_Json.Get_String('operator_code');
  v_操作员姓名    := j_Json.Get_String('operator_name');
  n_挂号划价      := j_Json.Get_Number('pricing_sign');
  If Nvl(n_挂号划价, 0) <> 1 Then
    n_挂号划价 := 0;
  End If;

  If d_接收时间 Is Null Then
    d_接收时间 := Sysdate;
  End If;

  Select Max(a.Id), Max(a.No), Max(a.姓名), Max(a.性别), Max(a.年龄), Max(c.编码), Max(a.费别), Max(a.出诊记录id), Sum(b.实收金额),
         Max(a.挂号项目id), Max(a.号别), Max(a.执行部门id), Max(a.执行人), Max(d.Id) As 医生id
  Into n_挂号id, v_挂号单, v_姓名, v_性别, v_年龄, v_付款方式编码, v_费别, n_出诊记录id, n_结帐金额, n_挂号项目id, v_号别, n_执行部门id, v_医生姓名, n_医生id
  From 病人挂号记录 A, 门诊费用记录 B, 医疗付款方式 C, 人员表 D
  Where a.No = v_挂号单 And a.记录性质 = 2 And a.记录状态 = 1 And a.No = b.No And b.记录性质 = 4 And a.医疗付款方式 = c.名称(+) And
        a.执行人 = d.姓名(+);

  If v_挂号单 Is Null Then
    Json_Out := zlJsonOut('预约挂号信息不存在，可能该预约挂号已被接收。');
    Return;
  End If;

  If n_挂号划价 = 1 Then
    n_结帐金额 := 0;
    Select Nextno(13) Into v_划价单 From Dual;
  
    Insert Into 门诊费用记录
      (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 门诊标志, 病人id, 标识号, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 数次, 收入项目id,
       收据费目, 标准单价, 应收金额, 实收金额, 记帐费用, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 摘要, 是否急诊, 挂号id, 主页id, 付款方式)
      Select 病人费用记录_Id.Nextval, 1, v_划价单, 0, a.序号, a.从属父号, a.价格父号, a.门诊标志, n_病人id, Decode(n_门诊号, 0, Null, n_门诊号), a.姓名,
             a.性别, a.年龄, a.病人科室id, a.费别, a.收费类别, a.收费细目id, b.计算单位, a.付数, a.数次, a.收入项目id, a.收据费目, a.标准单价, a.应收金额, a.实收金额,
             0, v_操作员姓名, n_执行部门id, v_操作员姓名, a.发生时间, d_接收时间, a.执行部门id, '挂号:' || v_挂号单, a.是否急诊, a.挂号id, a.主页id, a.付款方式
      From 门诊费用记录 A, 收费项目目录 B
      Where a.No = v_挂号单 And a.记录性质 = 4 And a.记录状态 = 0 And a.收费细目id = b.Id;
  
    Update 门诊费用记录
    Set 应收金额 = 0, 实收金额 = 0, 摘要 = '划价' || v_划价单
    Where NO = v_挂号单 And 记录性质 = 4 And 记录状态 = 0;
  End If;

  If Nvl(n_出诊记录id, 0) <> 0 Then
    Zl_预约挂号接收_出诊_Insert_s(v_挂号单, Null, n_病人id, n_门诊号, v_姓名, v_性别, v_年龄, v_付款方式编码, v_费别, n_接诊诊室, n_结帐id, n_结帐金额, d_接收时间,
                          d_接收时间, v_操作员编码, v_操作员姓名, n_挂号划价, Null, 0, Null, v_划价单, n_挂号项目id);
  Else
    Zl_预约挂号接收_Insert_s(v_挂号单, Null, n_病人id, n_门诊号, v_姓名, v_性别, v_年龄, v_付款方式编码, v_费别, n_接诊诊室, n_结帐id, n_结帐金额, d_接收时间,
                       d_接收时间, v_操作员编码, v_操作员姓名, n_挂号划价, Null, 0, Null, v_划价单, n_挂号项目id);
  End If;

  n_挂号生成队列 := zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113));
  If Nvl(n_挂号生成队列, 0) <> 0 Then
    n_挂号生成队列 := 1;
  End If;
  Update 门诊费用记录 Set 发生时间 = d_接收时间 Where 记录性质 = 4 And 记录状态 = 1 And NO = v_挂号单;
  Update 病人挂号记录 Set 发生时间 = d_接收时间 Where 记录状态 = 1 And NO = v_挂号单;

  If Nvl(n_挂号划价, 0) = 1 Or v_冲预交病人ids Is Not Null Then
    If v_冲预交病人ids Is Not Null Then
      Zl_病人挂号收费_Modify_s(v_挂号单, n_结帐id, n_结帐金额 || '|' || v_冲预交病人ids, 3, 1);
    End If;
    Zl_病人挂号记录_完成挂号_s(v_挂号单, n_在院, 2, n_挂号生成队列);
  
    --医生站自动接诊
    Update 门诊费用记录 Set 执行人 = v_操作员姓名, 执行时间 = d_接收时间, 执行状态 = 2 Where 记录性质 = 4 And 记录状态 = 1 And NO = v_挂号单;
    Update 病人挂号记录 Set 执行人 = v_操作员姓名, 执行时间 = d_接收时间, 执行状态 = 2 Where NO = v_挂号单;
    If Nvl(n_挂号生成队列, 0) = 1 Then
      --接收后,变成弃号
      Update 排队叫号队列 Set 排队状态 = 2 Where 业务类型 = 0 And 业务id = n_挂号id;
    End If;
  
  Else
    o_Json := j_Json.Get_Pljson('balance_info');
    If o_Json Is Not Null Then
      n_结帐金额 := o_Json.Get_Number('blnc_money');
      v_结算方式 := o_Json.Get_String('blnc_mode');
      n_卡类别id := o_Json.Get_Number('cardtype_id');
      v_支付卡号 := o_Json.Get_String('pay_cardno');
    
      Zl_病人挂号收费_Modify_s(v_挂号单, n_结帐id, v_结算方式 || ',' || n_结帐金额 || ', , ', 1, 0, 0, n_预交id, n_卡类别id, v_支付卡号, Null, Null,
                         0, 1);
    End If;
  End If;

  Zl_病人挂号汇总_Update(v_医生姓名, n_医生id, n_挂号项目id, n_执行部门id, d_接收时间, 2, v_号别, 0, n_出诊记录id);

  Json_Out := '{"output":{"code":1,"message":"成功","blnc_id":' || Nvl(n_结帐id || '', 'null') || ',"relation_id":' ||
              Nvl(n_预交id || '', 'null') || '}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Rgstapptreceive;
/

Create Or Replace Procedure Zl_Exsesvr_Updrgstbalanceinfo
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --功能：医生站预约接收更新挂号支付信息
  --入参：Json_In:格式
  --input
  --  rgst_no            C  1 挂号no
  --  blnc_id            N  1 结帐ID
  --  relation_id        N  1 关联交易ID
  --  pati_inhospital    N  1 在院病人
  --  totalmoney         N  1 支付总金额
  --  cardtype_id        N  1 支付卡类别ID
  --  rgst_recv_time     C  1 接收时间
  --  recharge           C    异常重新接诊
  --  operator_code      C    操作员编码
  --  operator_name      C    操作员姓名
  --  balance_list[]
  --     blnc_mode       N  1 结算方式
  --     swapmoney       C  1 结算金额
  --     swapno          C  1 交易流水号
  --     swapmemo        C  1 交易说明
  --     blnc_no         C  1 结算号码
  --     blnc_memo       C  1 结算摘要
  --     card_no         C  1 支付卡号
  --     cardtype_id     N    卡类别ID
  --  otherswap_list[]   C    其他交易信息
  --     swap_name       C  1  交易名称
  --     swap_note       C  1  交易内容
  -- 出参:
  --  output
  --    code                            N 1 应答吗：0-失败；1-成功
  --    message                         C 1 应答消息：失败时返回具体的错误信息
  -------------------------------------------
  n_挂号id     病人挂号记录.Id%Type;
  v_挂号单     病人挂号记录.No%Type;
  d_接收时间   病人挂号记录.接收时间%Type;
  v_操作员编码 病人挂号记录.操作员编号%Type;
  v_操作员姓名 病人挂号记录.操作员姓名%Type;
  n_结算金额   门诊费用记录.结帐金额%Type;
  n_结帐id     门诊费用记录.结帐id%Type;
  v_交易流水号 病人预交记录.交易流水号%Type;
  v_交易说明   病人预交记录.交易说明%Type;
  v_结算号码   病人预交记录.结算号码%Type;
  v_结算摘要   病人预交记录.摘要%Type;
  n_关联交易id 病人预交记录.关联交易id%Type;
  v_结算方式   病人预交记录.结算方式%Type;
  n_卡类别id   病人预交记录.卡类别id%Type;
  v_支付卡号   病人预交记录.卡号%Type;

  n_支付总金额   门诊费用记录.结帐金额%Type;
  n_合计金额     门诊费用记录.结帐金额%Type;
  v_交易名称     三方结算交易.交易项目%Type;
  v_交易内容     三方结算交易.交易内容%Type;
  v_扩展信息     Varchar2(4000);
  n_重收         Number(2);
  n_普通结算     Number(2);
  n_在院         Number(2);
  n_挂号生成队列 Number(2);
  n_连续更新     Number(2);
  n_完成挂号     Number(2);
  o_Json         PLJson;
  j_Input        PLJson;
  j_Json         PLJson;

  j_Jsonlist Pljson_List := Pljson_List();

Begin
  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_挂号单     := j_Json.Get_String('rgst_no');
  n_结帐id     := j_Json.Get_Number('blnc_id');
  n_关联交易id := j_Json.Get_Number('relation_id');
  n_在院       := j_Json.Get_Number('pati_inhospital');
  n_支付总金额 := j_Json.Get_Number('totalmoney');
  n_重收       := j_Json.Get_Number('recharge');
  d_接收时间   := To_date(j_Json.Get_String('rgst_recv_time'), 'YYYY-MM-DD hh24:mi:ss');
  v_操作员编码 := j_Json.Get_String('operator_code');
  v_操作员姓名 := j_Json.Get_String('operator_name');
  
  If Nvl(n_重收, 0) = 1 Then
    zl_病人挂号记录_异常重收_s(v_挂号单, Null, d_接收时间, v_操作员编码, v_操作员姓名);
  End If;
  
  n_连续更新 := 0;
  n_完成挂号 := 0;
  j_Jsonlist := j_Json.Get_Pljson_List('balance_list');
  If j_Jsonlist Is Not Null Then
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json       := PLJson();
      o_Json       := PLJson(j_Jsonlist.Get(I));
      v_结算方式   := o_Json.Get_String('blnc_mode');
      n_结算金额   := o_Json.Get_Number('swapmoney');
      v_交易流水号 := o_Json.Get_String('swapno');
      v_交易说明   := o_Json.Get_String('swapmemo');
      v_结算号码   := o_Json.Get_String('blnc_no');
      v_结算摘要   := o_Json.Get_String('blnc_memo');
      v_支付卡号   := o_Json.Get_String('card_no');
      n_卡类别id   := o_Json.Get_Number('cardtype_id');
    
      n_合计金额 := Nvl(n_合计金额, 0) + n_结算金额;
      If I > 1 Then
        n_连续更新 := 1;
      End If;
      If I = j_Jsonlist.Count And n_支付总金额 = n_合计金额 Then
        n_完成挂号 := 1;
      End If;
      If Nvl(n_卡类别id, 0) = 0 Then
        n_普通结算 := 1;
      End If;
      Zl_病人挂号收费_Modify_s(v_挂号单, n_结帐id, v_结算方式 || ',' || n_结算金额 || ',' || v_结算号码 || ',' || v_结算摘要, 1, n_完成挂号, n_连续更新,
                         n_关联交易id, n_卡类别id, v_支付卡号, v_交易流水号, v_交易说明, Nvl(n_普通结算, 0), 2);
    End Loop;
  End If;

  Begin
    n_卡类别id := j_Json.Get_Number('cardtype_id');
    j_Jsonlist := Pljson_List();
    j_Jsonlist := j_Json.Get_Pljson_List('otherswap_list');
    If j_Jsonlist Is Not Null Then
      For I In 1 .. j_Jsonlist.Count Loop
        o_Json     := PLJson();
        o_Json     := PLJson(j_Jsonlist.Get(I));
        v_交易名称 := o_Json.Get_String('swap_name');
        v_交易内容 := o_Json.Get_String('swap_note');
        v_扩展信息 := v_扩展信息 || '||' || v_交易名称 || '|' || v_交易名称 || '|' || v_交易内容;
      
        If Lengthb(v_扩展信息) > 2000 Then
          Zl_三方结算交易_Insert(n_卡类别id, 0, v_支付卡号, n_结帐id, v_扩展信息);
          v_扩展信息 := Null;
        End If;
      End Loop;
    End If;
  
    If v_扩展信息 Is Not Null Then
      Zl_三方结算交易_Insert(n_卡类别id, 0, v_支付卡号, n_结帐id, v_扩展信息);
    End If;
  Exception
    When Others Then
      Null;
  End;

  If n_完成挂号 = 1 Then
    n_挂号生成队列 := zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113));
    If Nvl(n_挂号生成队列, 0) <> 0 Then
      n_挂号生成队列 := 1;
    End If;
  
    Zl_病人挂号记录_完成挂号_s(v_挂号单, n_在院, 2, n_挂号生成队列);
  
    Update 门诊费用记录 Set 执行人 = v_操作员姓名, 执行时间 = d_接收时间, 执行状态 = 2 Where 记录性质 = 4 And 记录状态 = 1 And NO = v_挂号单;
    Update 病人挂号记录 Set 执行人 = v_操作员姓名, 执行时间 = d_接收时间, 执行状态 = 2 Where NO = v_挂号单;
    If Nvl(n_挂号生成队列, 0) = 1 Then
      --接收后,变成弃号
      Select Max(ID) Into n_挂号id From 病人挂号记录 Where NO = v_挂号单;
      Update 排队叫号队列 Set 排队状态 = 2 Where 业务类型 = 0 And 业务id = n_挂号id;
    End If;
  End If;

  Json_Out := zlJsonOut('成功', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Updrgstbalanceinfo;
/

Create Or Replace Procedure Zl_Exsesvr_Odr_Check
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：医嘱超期发送收费用回检查和数据获取
  --入参：Json_In:格式
  --  input
  --    check_type                  N 1 检查类型，
  --                                说明：1-检查要收回的医嘱对应的费用结帐情况，根据 order_ids查询
  --                                      10-正式超期收回此时order_ids传入一个医嘱id,配合[收回的其它参数]
  --    order_ids                   C 1 医嘱ID拼串
  -----------[收回的其它参数]----check_type=10时传入------------------------------------------------------------------------------------------
  --    roll_num                    N 1 本次要收回的量，check_type=10时传入
  --    fee_no                      N 1 费用单据信息,主要用于区分当前收回模:null-销帐申请; 调整划价单-表示是调整划价单,具体单据号-表是负数冲销
  --    fee_nos                     C 1 传入的医嘱 order_ids 对应的单据号,逗号拼串,此时的 order_ids 只有一个值
  --    advice_dosage               N 1 单次用量
  --    advice_note                 N 1 医嘱内容
  --    clinic_type                 C 1 诊疗类别
  --    is_stuff_order                    N 1 跟踪在用卫材医嘱
  --    price_list[]医嘱计价列表
  --          order_id               N 1 医嘱id
  --          fee_item_id            N 1 收费细目id
  --          refer_num              N 1 对照数量
  --          fee_way                N 1 收费方式，普通收费方式为0正常收取
  --    price_exe_list[]医嘱执行计价数量列表
  --          fee_item_id            N 1 收费细目id
  --          roll_num               N 1 收回数量
  --    excute_list[]           单据已执行列表(药品、卫材费用),即使已执行数为0也要传入
  --          fee_id              N   1   费用ID
  --          sended_num          N   1   已发数量
  --    advice_excute_list[]    单据已执行列表(医嘱费用),即使已执行数为0也要传入
  --          advice_id           N   1   医嘱ID
  --          fee_item_id         N   1   收费细目ID
  --          execute_num         N   1   已执行数
  --    pati_list[]             病人信息，仅审核这些病人的费用
  --          pati_id             N   1   病人ID
  --          pati_name           C   1   病人姓名
  --          fee_audit_status    N   1   费用审核标志:0或空-未审核;1-已审核或开始审核(结合参数:病人审核方式来控制);2-完成审核,结合结帐权限[禁止未审核病人结帐]进行管理控制
  --          si_inp_status       N   1   住院状态:0-正常住院;1-尚未入科;2-正在转科或正在转病区;3-已预出院
  --          catalog_date        C   0   病案编目日期：yyyy-mm-dd hh24:mi:ss

  --出参: Json_Out,格式如下
  --  output
  --    code          N 1 应答吗：0-失败；1-成功
  --    message       C 1 应答消息：失败时返回具体的错误信息

  --    charge_list[]销帐申请列表
  --                  outpati_account        N 1 门诊记帐0-住院记账,1-门诊记帐
  --                  fee_id                 N 1 费用id
  --                  fee_item_id            N 1 收费细目id
  --                  request_dept_id        N 1 申请科室id
  --                  audit_dept_id          N 1 审核科室id
  --                  request_num            N 1 申请数量

  --    del_list[]单据删除列表
  --                  outpati_account        N 1 门诊记帐0-住院记账,1-门诊记帐
  --                  fee_no                 C 1 费用单据号
  --                  serial_num             C 1 删除序号,序号格式:数量:执行状态

  --   del_drug[]药品删除列表
  --                  rcpdtl_id              N 1 处方明细id,目前传入的费用ID
  --                  chargeoffs_num         N 1 销帐数量

  --   del_stuff[]卫材删除列表
  --                  stuffdtl_id            N 1 处方明细id,目前传入的费用ID
  --                  return_num             N 1 销帐数量

  --   pivas_list[]静配销帐列表
  --                  pivas_id               N 1 配液id
  --                  auto_aduit             N 1 是否自动审核 0-不审核,1-要自动审核
  --                  request_time           C 1 申请时间
  --                  reason                 C 1 销帐原因

  --    roll_list[]负数单据列表费用
  --                  outpati_account        N 1 门诊记帐0-住院记账,1-门诊记帐
  --                  clinic_type            C 1 医嘱诊疗类别
  --                  fee_no                 C 1 单据号
  --                  item_type              C 1 收费细目类别
  --                  fee_id                 N 1 费用id
  --                  fee_id_old             N 1 费用id,原始费用id
  --                  packages_num           N 1 付数
  --                  send_num               N 1 数次
  --                  is_stuff_order         N 1 区分是否是绑定的卫材费用0-非卫材医嘱,1-卫材医嘱
  --                  stuff_used             N 1 是否是跟踪在在卫才费用
  --                  exe_status             N 1 执行状态

  --     roll_drug_list[]负数收回列表药品
  --                  clinic_type            C 1 医嘱诊疗类别
  --                  rcp_no                 C 1 处方号,费用单号
  --                  rcpdtl_id              N 1 处方明细ID
  --                  rcpdtl_id_old          N 1 处方明细ID,原始明细id
  --                  packages_num           N 1 发药付数
  --                  send_num               N 1 发药数量
  --                  item_type              C 1 收费项目类别

  --     roll_stuff_list[]负数收回列表卫材
  --                  clinic_type            C 1 医嘱诊疗类别
  --                  stuff_no               C 1 负数记帐的单据号
  --                  stuffdtl_id            N 1 卫材明细id,费用记录id
  --                  stuffdtl_id_old        N 1 原始卫材明细id,费用记录id
  --                  packages_num           N 1 付数
  --                  outbound_num           N 1 数量
  --                  is_stuff_order         N 1 区分是否是绑定的卫材费用0-非卫材医嘱,1-卫材医嘱

  ---------------------------------------------------------------------------

  --    price_list[]医嘱计价列表
  Type t_Rs_计价 Is Record(
    医嘱id     Number,
    收费细目id Number,
    对照数量   Number,
    收回数量   Number,
    收费方式   Number);
  Type t_计价 Is Table Of t_Rs_计价;
  Rs_计价 t_计价;
  Rs_收回 t_计价;

  Type t_Rs_执行 Is Record(
    医嘱id   Number,
    费用id   Number,
    收发id   Number,
    未执行量 Number,
    已执行量 Number);
  Type t_执行 Is Table Of t_Rs_执行;
  Rs_执行 t_执行; --药品卫材已经执行数量

  Type t_Rs_静配 Is Record(
    医嘱id     Number,
    费用id     Number,
    收费细目id Number,
    数量       Number,
    配液id     Number,
    操作状态   Number);
  Type t_静配 Is Table Of t_Rs_静配;
  Rs_静配 t_静配;

  Type t_Rs_单据 Is Record(
    费用id       Number(18),
    序号         Number,
    数量         Number,
    门诊记帐     Number,
    销帐         Number,
    收费细目id   Number,
    申请科室id   Number,
    审核科室id   Number,
    已执行数     Number,
    申请时间     Date,
    申请类别     Number,
    静配自动审核 Number,
    配液id       Number,
    NO           Varchar2(60),
    类别         Number, --  0-普通,1-药品,2-卫材
    医嘱id       Number);
  Type t_单据 Is Table Of t_Rs_单据;
  Rs_单据    t_单据;
  Rs_其它    t_单据;
  Rs_Dellist t_单据;

  j_Json       PLJson;
  j_Tmp        PLJson;
  j_Output     PLJson;
  j_Item       PLJson;
  j_List       Pljson_List := Pljson_List();
  j_List_Order Pljson_List := Pljson_List();

  Lngtmp Number;
  --v_Dec                Number;
  v_Json_In   Varchar2(32767);
  v_Json_Out  Varchar2(32767);
  v_Item_List Varchar2(32767); --序号+数量
  --v_划价类别           Varchar2(322);
  v_Orderfeenos        Varchar2(32767); --医嘱对应的费用单据号,逗号拼串,医嘱ID+单据号 可以唯一确定费用了
  v_Pati_List          Varchar2(32767);
  v_Chk销帐申请        Varchar2(32767);
  v_Excute_List        Varchar2(32767);
  v_Advice_Excute_List Varchar2(32767);
  v_Charge_List        Varchar2(32767);
  v_Del_List           Varchar2(32767);
  v_医嘱内容           Varchar2(4000);
  v_Del_Drug           Varchar2(32767);
  v_Del_Stuff          Varchar2(32767);
  v_自动发料           Varchar2(4000);

  收回量_In     Number;
  v_收费细目id  Number;
  Nt_收费细目id Number;
  Nt_跟踪在用   Number;
  Nt_销帐原因   Varchar2(4000);
  Nt_发送单据   Varchar2(200);
  Nt_发送数次   Number;
  Nt_自动审核   Number;
  Nt_收费方式   Number;
  Nt_单次用量   Number;
  Nt_诊疗类别   Varchar2(30);
  Nt_收费标志   Number;
  Nt_医嘱内容   Varchar2(4000);
  v_收费内容    Varchar2(4000);
  v_费用ids     Varchar2(32767);
  v_划价类别    Varchar2(32767);
  v_费用id      Number(18);
  v_当前数量    Number;
  v_剩余数量    Number;
  v_结帐金额    Number;
  v_收回剩余    Number;
  v_收回数量    Number;
  v_收回数量tmp Number;
  v_收回量      Number;
  v_对照数量    Number;
  v_剂量系数    药品规格.剂量系数%Type;
  v_住院包装    药品规格.住院包装%Type;
  v_结帐参数    Varchar2(4000);
  No_In         Varchar2(2000);
  v_自审        Number;
  v_Tmp         Varchar2(32767);
  v_Pivas_Ids   Varchar2(32767);
  v_Pivas_Out   Varchar2(32767);
  v_Roll_List   Varchar2(32767);
  v_Roll_List_d Varchar2(32767); --负数收回的药品
  v_Roll_List_s Varchar2(32767); --负数收回的卫材
  n_检查方式    Number;
  l_医嘱ids     t_StrList;
  c_医嘱ids     Clob;
  Vo_Vals       Clob;
  v_Error       Varchar2(255);
  b_In          Clob;
  b_Out         Clob;
  收回时间_In   Date;
  Err_Custom Exception;
  医嘱id_In Number;

  n_静配销帐操作 Number;

  --销帐申请时间,申请类别,销帐数量,静配自动审核,销帐模式

  --定义结构
  Cursor c_Fee_List_Type Is
    Select a.Id 费用id, a.No, a.序号, a.收费细目id, a.病人病区id, a.收费类别, a.执行状态 跟踪在用, a.收费类别 诊疗类别, a.摘要 医嘱内容, a.数次 单次用量, a.数次 剩余数量,
           a.数次 已执行量, a.数次 未执行量, a.执行状态 执行标志, a.记录状态, a.发生时间 登记时间, a.执行状态 收费方式, a.执行状态 门诊记帐, a.发生时间 销帐申请时间, a.执行状态 申请类别,
           a.数次 销帐数量, a.执行状态 静配自动审核, a.执行状态 销帐模式, a.执行部门id, a.Id 配液id
    From 住院费用记录 A
    Where 0 = 1;
  r_Detail c_Fee_List_Type%RowType;

  --直接销帐的情况
  Cursor c_Detail Is
    Select a.费用id, a.No, a.序号, a.收费细目id, a.病人病区id, a.收费类别, a.跟踪在用, Null 诊疗类别, Null 医嘱内容, -null 单次用量, a.剩余数量, -null 已执行量,
           -null 未执行量, a.执行标志, a.记录状态, a.登记时间, -null 收费方式, a.门诊记帐
    From (Select 0 As 门诊记帐, Max(Decode(b.记录状态, 2, 0, b.Id)) As 费用id, b.No, Nvl(b.价格父号, b.序号) As 序号, b.收费细目id, b.病人病区id,
                  Sum(Nvl(b.付数, 1) * b.数次) As 剩余数量, b.收费类别, Max(Nvl(b.执行状态, 0)) As 执行标志, d.跟踪在用, Max(b.记录状态) As 记录状态,
                  Max(b.登记时间) As 登记时间
           From 住院费用记录 B, 材料特性 D
           Where b.医嘱序号 = 医嘱id_In And b.价格父号 Is Null And b.收费细目id = d.材料id(+) And
                 Instr(',' || v_Orderfeenos || ',', ',' || b.No || ',') > 0
           Group By b.No, b.记录性质, Nvl(b.价格父号, b.序号), b.收费细目id, b.病人病区id, b.收费类别, d.跟踪在用
           Having Sum(Nvl(b.付数, 1) * b.数次) > 0
           Union All
           Select 1 As 门诊记帐, Max(Decode(b.记录状态, 2, 0, b.Id)) As 费用id, b.No, Nvl(b.价格父号, b.序号) As 序号, b.收费细目id, b.病人病区id,
                  Sum(Nvl(b.付数, 1) * b.数次) As 剩余数量, b.收费类别, Max(Nvl(b.执行状态, 0)) As 执行标志, d.跟踪在用, Max(b.记录状态) As 记录状态,
                  Max(b.登记时间) As 登记时间
           From 门诊费用记录 B, 材料特性 D
           Where b.医嘱序号 = 医嘱id_In And b.价格父号 Is Null And b.收费细目id = d.材料id(+) And
                 Instr(',' || v_Orderfeenos || ',', ',' || b.No || ',') > 0
           Group By b.No, b.记录性质, Nvl(b.价格父号, b.序号), b.收费细目id, b.病人病区id, b.收费类别, d.跟踪在用
           Having Sum(Nvl(b.付数, 1) * b.数次) > 0) A
    Order By a.收费细目id, a.执行标志, a.登记时间 Desc;

  --未生效单计帐划价单
  Cursor c_Del Is
    Select b.Id As 费用id, b.No, b.序号, b.收费类别, b.收费细目id, Nvl(b.付数, 1) * b.数次 As 剩余数量, d.跟踪在用, b.门诊记帐
    From (Select 0 As 门诊记帐, a.Id, a.No, a.序号, a.收费细目id, a.付数, a.数次, a.价格父号, a.医嘱序号, a.收费类别
           From 住院费用记录 A
           Where a.医嘱序号 = 医嘱id_In And a.记录状态 = 0
           Union All
           Select 1 As 门诊记帐, a.Id, a.No, a.序号, a.收费细目id, a.付数, a.数次, a.价格父号, a.医嘱序号, a.收费类别
           From 门诊费用记录 A
           Where a.医嘱序号 = 医嘱id_In And a.记录状态 = 0) B, 材料特性 D
    Where b.医嘱序号 = 医嘱id_In And b.价格父号 Is Null And b.收费细目id = d.材料id(+) And
          Instr(',' || v_Orderfeenos || ',', ',' || b.No || ',') > 0
    Order By b.收费细目id, b.No Desc;

  --负数冲销，不考虑静配
  Cursor c_Negdrug Is
    Select b.Id As 费用id, b.No, b.序号, b.收费类别, b.收费细目id, b.病人病区id, Nvl(b.付数, 1) * b.数次 As 剩余数量, b.门诊记帐, b.记录状态,
           b.执行状态 As 执行标志, d.跟踪在用, b.执行部门id
    From (Select 0 As 门诊记帐, a.Id, a.No, a.序号, a.收费细目id, a.病人病区id, a.付数, a.数次, a.价格父号, a.医嘱序号, a.执行状态, a.记录状态, a.收费类别,
                  a.执行部门id
           From 住院费用记录 A
           Where a.医嘱序号 = 医嘱id_In And a.记录状态 In (0, 1, 3)
           Union All
           Select 1 As 门诊记帐, a.Id, a.No, a.序号, a.收费细目id, a.病人病区id, a.付数, a.数次, a.价格父号, a.医嘱序号, a.执行状态, a.记录状态, a.收费类别,
                  a.执行部门id
           From 门诊费用记录 A
           Where a.医嘱序号 = 医嘱id_In And a.记录状态 In (0, 1, 3)) B, 材料特性 D
    Where Instr(',' || v_Orderfeenos || ',', ',' || b.No || ',') > 0 And b.收费细目id = d.材料id(+)
    Order By b.收费细目id, b.No Desc;

  --包含非药长嘱(含给药途径)发送时所产生的费用(因多个收入而有多条记录)
  --对非药医嘱,直接收回指定量,不管多次发送(如果多次发送价格不同,则收回的价格是以最后次的；不然就要根据多个收入依次减收回量)。
  --卫材本身是售价单位，无需住院单位转换
  --非药长嘱都填写了发送记录(除开了叮嘱及护理等级)
  --一天只收一次或一次发送只收一次的项目暂时不支持负数申请
  Cursor c_Other Is
    Select a.门诊记帐, a.No, a.序号, a.费用id, a.剩余数量, a.收费细目id, a.病人病区id, a.记录状态, a.执行标志, a.对照数量, a.收费方式, a.收费类别, d.跟踪在用,
           a.执行部门id
    From (Select 0 As 门诊记帐, a.No, a.序号, a.记录状态, a.收费细目id, a.病人病区id, a.Id As 费用id, a.数次 As 剩余数量, Nvl(a.执行状态, 0) As 执行标志,
                  a.医嘱序号, Null 发送号, Null 对照数量, Null 收费方式, a.收费类别, a.执行部门id
           From 住院费用记录 A
           Where a.No = Nt_发送单据 And a.记录状态 In (0, 1, 3) And a.医嘱序号 + 0 = 医嘱id_In
           Union All
           Select 1 As 门诊记帐, a.No, a.序号, a.记录状态, a.收费细目id, a.病人病区id, a.Id As 费用id, a.数次 As 剩余数量, Nvl(a.执行状态, 0) As 执行状态,
                  a.医嘱序号, Null 发送号, Null 对照数量, Null 收费方式, a.收费类别, a.执行部门id
           From 门诊费用记录 A
           Where a.No = Nt_发送单据 And a.记录状态 In (0, 1, 3) And a.医嘱序号 + 0 = 医嘱id_In) A, 材料特性 D
    Where a.收费细目id = d.材料id(+)
    Order By a.收费细目id, a.序号, a.记录状态;

  Procedure p_Add_Negbill As
    P付数     Number;
    P自动发料 Number;
  Begin
  
    If Nt_诊疗类别 = '7' Then
      -- 中药医嘱，配方式付数还原;
      P付数             := 收回量_In;
      r_Detail.销帐数量 := r_Detail.销帐数量 / P付数;
    Else
      P付数 := 1;
    End If;
  
    --对于非药品卫材医嘱行,要对费用进行自动审核,药品卫材需要自动发料
    --产生负数单据
    Select 病人费用记录_Id.Nextval Into v_费用id From Dual;
  
    v_Roll_List := v_Roll_List || ',{"outpati_account":' || r_Detail.门诊记帐; --0-住院记帐,1-门诊记帐
    v_Roll_List := v_Roll_List || ',"clinic_type":"' || Nt_诊疗类别 || '"'; --医嘱的诊疗类别
    v_Roll_List := v_Roll_List || ',"fee_no":"' || No_In || '"';
    v_Roll_List := v_Roll_List || ',"item_type":"' || r_Detail.收费类别 || '"';
    v_Roll_List := v_Roll_List || ',"fee_id":' || v_费用id;
    v_Roll_List := v_Roll_List || ',"fee_id_old":' || r_Detail.费用id;
    v_Roll_List := v_Roll_List || ',"packages_num":' || zlJsonStr(P付数, 1);
    v_Roll_List := v_Roll_List || ',"send_num":' || zlJsonStr(r_Detail.销帐数量, 1);
    v_Roll_List := v_Roll_List || ',"is_stuff_order":' || Nvl(Nt_跟踪在用, 0); --是否是跟踪在用的医嘱
    v_Roll_List := v_Roll_List || ',"stuff_used":' || Nvl(r_Detail.跟踪在用, 0); --是否是跟踪在用卫材费用
    v_Roll_List := v_Roll_List || ',"exe_status":' || Nvl(r_Detail.执行标志, 0);
    v_Roll_List := v_Roll_List || ',"exe_deptid":' || r_Detail.执行部门id; --这几个结点将来可用于计算成本价,暂时未用
    v_Roll_List := v_Roll_List || ',"fee_item_id":' || r_Detail.收费细目id;
    v_Roll_List := v_Roll_List || ',"charge_tag":' || Nvl(Nt_收费标志, 0);
    v_Roll_List := v_Roll_List || '}';
  
    If r_Detail.收费类别 In ('5', '6', '7') Then
      --是否收费
      v_Roll_List_d := v_Roll_List_d || ',{"clinic_type":"' || Nt_诊疗类别 || '"'; --医嘱的诊疗类别
      v_Roll_List_d := v_Roll_List_d || ',"rcp_no":"' || No_In || '"';
      v_Roll_List_d := v_Roll_List_d || ',"rcpdtl_id":' || v_费用id;
      v_Roll_List_d := v_Roll_List_d || ',"rcpdtl_id_old":' || r_Detail.费用id;
      v_Roll_List_d := v_Roll_List_d || ',"packages_num":' || zlJsonStr(P付数, 1);
      v_Roll_List_d := v_Roll_List_d || ',"send_num":' || zlJsonStr(r_Detail.销帐数量, 1);
      v_Roll_List_d := v_Roll_List_d || ',"charge_tag":' || Nvl(Nt_收费标志, 0);
      v_Roll_List_d := v_Roll_List_d || '}';
    End If;
  
    If r_Detail.跟踪在用 = 1 Then
      If v_自动发料 = '1' And r_Detail.执行标志 = 1 Then
        P自动发料 := 1;
      End If;
      --是否收费
      v_Roll_List_s := v_Roll_List_s || ',{"clinic_type":"' || Nt_诊疗类别 || '"'; --医嘱的诊疗类别
      v_Roll_List_s := v_Roll_List_s || ',"stuff_no":"' || No_In || '"';
      v_Roll_List_s := v_Roll_List_s || ',"stuffdtl_id":' || v_费用id;
      v_Roll_List_s := v_Roll_List_s || ',"stuffdtl_id_old":' || r_Detail.费用id;
      v_Roll_List_s := v_Roll_List_s || ',"packages_num":' || zlJsonStr(P付数, 1);
      v_Roll_List_s := v_Roll_List_s || ',"outbound_num":' || zlJsonStr(r_Detail.销帐数量, 1);
      v_Roll_List_s := v_Roll_List_s || ',"is_stuff_order":' || Nvl(Nt_跟踪在用, 0);
      v_Roll_List_s := v_Roll_List_s || ',"stuff_auto_send":' || Nvl(P自动发料, 0);
      v_Roll_List_s := v_Roll_List_s || ',"charge_tag":' || Nvl(Nt_收费标志, 0);
      v_Roll_List_s := v_Roll_List_s || '}';
    End If;
  End;

  Procedure p_Getoutlist As
    --出参列表串组装
    P序号    Varchar2(32767);
    Pno      Varchar2(30);
    p_Deltmp Varchar2(32767);
  Begin
    If v_Charge_List Is Not Null Then
      v_Json_Out := v_Json_Out || ',"charge_list":[' || Substr(v_Charge_List, 2) || ']';
    End If;
  
    If v_Del_List Is Not Null Then
      v_Del_List := Null;
      --需要汇总下
      For I In 1 .. Rs_Dellist.Count Loop
        If Pno Is Not Null And Pno <> Rs_Dellist(I).No Then
          p_Deltmp := '{"outpati_account":' || Rs_Dellist(I - 1).门诊记帐;
          p_Deltmp := p_Deltmp || ',"fee_no":"' || Rs_Dellist(I - 1).No || '"';
          p_Deltmp := p_Deltmp || ',"serial_num":"' || Substr(P序号, 2) || '"';
          p_Deltmp := p_Deltmp || '}';
        
          If v_Del_List Is Null Then
            v_Del_List := p_Deltmp;
          Else
            v_Del_List := p_Deltmp || ',' || v_Del_List;
          End If;
          P序号 := Null;
        End If;
        Pno   := Rs_Dellist(I).No;
        P序号 := P序号 || ',' || Rs_Dellist(I).序号 || ':' || Rs_Dellist(I).数量 || ':0';
      End Loop;
    
      p_Deltmp := '{"outpati_account":' || Rs_Dellist(Rs_Dellist.Count).门诊记帐;
      p_Deltmp := p_Deltmp || ',"fee_no":"' || Rs_Dellist(Rs_Dellist.Count).No || '"';
      p_Deltmp := p_Deltmp || ',"serial_num":"' || Substr(P序号, 2) || '"';
      p_Deltmp := p_Deltmp || '}';
    
      If v_Del_List Is Null Then
        v_Del_List := ',' || p_Deltmp;
      Else
        v_Del_List := ',' || p_Deltmp || ',' || v_Del_List;
      End If;
    
      v_Json_Out := v_Json_Out || ',"del_list":[' || Substr(v_Del_List, 2) || ']';
    
    End If;
    If v_Del_Drug Is Not Null Then
      v_Json_Out := v_Json_Out || ',"del_drug":[' || Substr(v_Del_Drug, 2) || ']';
    End If;
    If v_Del_Stuff Is Not Null Then
      v_Json_Out := v_Json_Out || ',"del_stuff":[' || Substr(v_Del_Stuff, 2) || ']';
    End If;
    If v_Roll_List Is Not Null Then
      v_Json_Out := v_Json_Out || ',"roll_list":[' || Substr(v_Roll_List, 2) || ']';
    End If;
  
    If v_Roll_List_d Is Not Null Then
      v_Json_Out := v_Json_Out || ',"roll_drug_list":[' || Substr(v_Roll_List_d, 2) || ']';
    End If;
  
    If v_Roll_List_s Is Not Null Then
      v_Json_Out := v_Json_Out || ',"roll_stuff_list":[' || Substr(v_Roll_List_s, 2) || ']';
    End If;
  
    If v_Pivas_Out Is Not Null Then
      v_Json_Out := v_Json_Out || ',"pivas_list":[' || Substr(v_Pivas_Out, 2) || ']';
    End If;
  End;

  Procedure p_Add_Delitem As
    --添加要删除的费用列和销帐申请列表元素
  Begin
    Rs_单据.Extend;
    Lngtmp := Rs_单据.Count;
    Rs_单据(Lngtmp).费用id := r_Detail.费用id;
    Rs_单据(Lngtmp).No := r_Detail.No;
    Rs_单据(Lngtmp).序号 := r_Detail.序号;
    Rs_单据(Lngtmp).数量 := r_Detail.销帐数量;
    Rs_单据(Lngtmp).门诊记帐 := r_Detail.门诊记帐;
    Rs_单据(Lngtmp).配液id := r_Detail.配液id;
    If r_Detail.收费类别 In ('5', '6', '7') Then
      Rs_单据(Lngtmp).类别 := 1;
    Elsif r_Detail.收费类别 = '4' And r_Detail.跟踪在用 = 1 Then
      Rs_单据(Lngtmp).类别 := 2;
    Else
      Rs_单据(Lngtmp).类别 := 0;
    End If;
    Rs_单据(Lngtmp).销帐 := r_Detail.销帐模式;
    If Rs_单据(Lngtmp).销帐 = 1 Then
      Rs_单据(Lngtmp).收费细目id := r_Detail.收费细目id;
      Rs_单据(Lngtmp).申请科室id := r_Detail.病人病区id;
      Rs_单据(Lngtmp).静配自动审核 := r_Detail.静配自动审核;
      Rs_单据(Lngtmp).申请类别 := r_Detail.申请类别;
      Rs_单据(Lngtmp).已执行数 := r_Detail.已执行量;
      Rs_单据(Lngtmp).申请时间 := r_Detail.销帐申请时间;
    End If;
  End;

  Procedure p_Delbill_Check(Prownum Number) As
    --直接可以删除的单据检查
    Rp    Number;
    Phave Number := 0;
  Begin
    Rp          := Prownum;
    v_Item_List := '{"serial_num":' || Rs_单据(Rp).序号;
    v_Item_List := v_Item_List || ',"quantity":' || zlJsonStr(Rs_单据(Rp).数量, 1);
    v_Item_List := v_Item_List || '}';
    v_Json_Out  := Null;
    If Rs_单据(Rp).门诊记帐 = 1 Then
      v_Json_In := '{"fee_no":"' || Rs_单据(Rp).No || '"';
      v_Json_In := v_Json_In || ',"fee_bill_type":2';
      v_Json_In := v_Json_In || ',"item_list":[' || v_Item_List || ']';
      v_Json_In := v_Json_In || v_Excute_List;
      v_Json_In := v_Json_In || v_Advice_Excute_List;
      v_Json_In := v_Json_In || '}';
      v_Json_In := '{"input":' || v_Json_In || '}';
      Zl_门诊记帐记录_Delete_Check(v_Json_In, v_Json_Out);
    Else
      v_Json_In := '{"fee_no":"' || Rs_单据(Rp).No || '"';
      v_Json_In := v_Json_In || ',"fee_bill_type":2';
      v_Json_In := v_Json_In || ',"balance_ban_writeoffs":0';
      v_Json_In := v_Json_In || ',"part_ban_writeoffs":0';
      v_Json_In := v_Json_In || ',"item_list":[' || v_Item_List || ']';
      v_Json_In := v_Json_In || v_Excute_List;
      v_Json_In := v_Json_In || v_Advice_Excute_List;
      v_Json_In := v_Json_In || v_Pati_List;
      v_Json_In := v_Json_In || '}';
      v_Json_In := '{"input":' || v_Json_In || '}';
      Zl_住院记帐记录_Delete_Check(v_Json_In, v_Json_Out);
    End If;
  
    j_Tmp    := PLJson();
    j_Output := PLJson();
    j_Tmp    := PLJson(v_Json_Out);
    j_Output := j_Tmp.Get_Pljson('output');
    If j_Output.Get_Number('code') = 0 Then
      v_Error := j_Output.Get_String('message');
      Raise Err_Custom;
    End If;
    j_List := Pljson_List();
    j_List := j_Output.Get_Pljson_List('item_list');
    j_Tmp  := PLJson();
    j_Tmp  := PLJson(j_List.Get(1));
  
    --数量
    Lngtmp := j_Tmp.Get_Number('quantity');
  
    For I In 1 .. Rs_Dellist.Count Loop
      If Rs_Dellist(I).No = Rs_单据(Rp).No And Rs_Dellist(I).门诊记帐 = Nvl(Rs_单据(Rp).门诊记帐, 0) And Rs_Dellist(I)
         .序号 = j_Tmp.Get_Number('serial_num') Then
        Rs_Dellist(I).数量 := Rs_Dellist(I).数量 + j_Tmp.Get_Number('quantity');
        Phave := 1;
      End If;
    End Loop;
  
    If Phave <> 1 Then
      Rs_Dellist.Extend;
      Rs_Dellist(Rs_Dellist.Count).No := Rs_单据(Rp).No;
      Rs_Dellist(Rs_Dellist.Count).门诊记帐 := Nvl(Rs_单据(Rp).门诊记帐, 0);
      Rs_Dellist(Rs_Dellist.Count).序号 := j_Tmp.Get_Number('serial_num');
      Rs_Dellist(Rs_Dellist.Count).数量 := j_Tmp.Get_Number('quantity');
    End If;
  
    v_Tmp := j_Tmp.Get_Number('serial_num');
    v_Tmp := v_Tmp || ':' || j_Tmp.Get_Number('quantity');
    v_Tmp := v_Tmp || ':' || j_Tmp.Get_Number('execute_tag');
  
    --单据删除列表，要汇总
    v_Del_List := v_Del_List || ',{"outpati_account":' || Nvl(Rs_单据(Rp).门诊记帐, 0);
    v_Del_List := v_Del_List || ',"fee_no":"' || Rs_单据(Rp).No || '"';
    v_Del_List := v_Del_List || ',"serial_num":"' || v_Tmp || '"';
    v_Del_List := v_Del_List || '}';
    v_Del_List := '有删除';
    --药品列表
    If Rs_单据(Rp).类别 = 1 Then
      v_Del_Drug := v_Del_Drug || ',{"rcpdtl_id":' || Rs_单据(Rp).费用id;
      v_Del_Drug := v_Del_Drug || ',"chargeoffs_num":' || zlJsonStr(Lngtmp, 1);
      If Rs_单据(Rp).配液id Is Not Null Then
        v_Del_Drug := v_Del_Drug || ',"pivas_id":' || Rs_单据(Rp).配液id;
      End If;
      v_Del_Drug := v_Del_Drug || '}';
    End If;
  
    --卫材列表
    If Rs_单据(Rp).类别 = 2 Then
      v_Del_Stuff := v_Del_Stuff || ',{"stuffdtl_id":' || Rs_单据(Rp).费用id;
      v_Del_Stuff := v_Del_Stuff || ',"return_num":' || zlJsonStr(Lngtmp, 1);
      v_Del_Stuff := v_Del_Stuff || '}';
    End If;
  End;

  Procedure p_Charge_Check(Prownum Number) As
    --销帐申请检查
  Begin
    Lngtmp        := Prownum;
    v_Chk销帐申请 := ',{"fee_id":' || Rs_单据(Lngtmp).费用id;
    v_Chk销帐申请 := v_Chk销帐申请 || ',"fee_item_id":' || Rs_单据(Lngtmp).收费细目id;
    v_Chk销帐申请 := v_Chk销帐申请 || ',"request_dept_id":' || Rs_单据(Lngtmp).申请科室id;
    v_Chk销帐申请 := v_Chk销帐申请 || ',"audit_dept_id":0'; --审核部门不确定传0通过检查方法来确定
    v_Chk销帐申请 := v_Chk销帐申请 || ',"request_type":' || Nvl(Rs_单据(Lngtmp).申请类别, 0);
    v_Chk销帐申请 := v_Chk销帐申请 || ',"request_num":' || zlJsonStr(Rs_单据(Lngtmp).数量, 1);
    v_Chk销帐申请 := v_Chk销帐申请 || ',"sended_num":' || zlJsonStr(Rs_单据(Lngtmp).已执行数, 1);
    v_Chk销帐申请 := v_Chk销帐申请 || '}';
  
    v_Tmp := '{"input":{';
    v_Tmp := v_Tmp || '"item_list":[' || Substr(v_Chk销帐申请, 2) || ']';
    v_Tmp := v_Tmp || v_Pati_List;
    v_Tmp := v_Tmp || '}}';
    b_In  := v_Tmp;
    Zl_病人费用销帐_Insert_Check(b_In, b_Out);
    j_Tmp    := PLJson();
    j_Output := PLJson();
    j_Tmp    := PLJson(b_Out);
    j_Output := j_Tmp.Get_Pljson('output');
    If j_Output.Get_Number('code') = 0 Then
      v_Error := j_Output.Get_String('message');
      Raise Err_Custom;
    End If;
    j_List := Pljson_List();
    j_List := j_Output.Get_Pljson_List('item_list');
    j_Tmp  := PLJson();
    j_Tmp  := PLJson(j_List.Get(1));
  
    Rs_单据(Lngtmp).审核科室id := j_Tmp.Get_Number('audit_dept_id');
    Nt_自动审核 := 0;
    --v_自审:参数判断
    If Rs_单据(Lngtmp).审核科室id = Rs_单据(Lngtmp).申请科室id And (v_自审 = 1 Or Rs_单据(Lngtmp).静配自动审核 = 1) Then
    
      v_Tmp := '{"fee_id":' || Rs_单据(Lngtmp).费用id;
      v_Tmp := v_Tmp || ',"stuff_auto_return":' || 0;
      v_Tmp := v_Tmp || ',"request_time":""';
      v_Tmp := v_Tmp || ',"request_type":' || Nvl(Rs_单据(Lngtmp).申请类别, 0);
      v_Tmp := v_Tmp || ',"sended_num":' || zlJsonStr(Rs_单据(Lngtmp).已执行数, 1);
      v_Tmp := v_Tmp || '}';
    
      v_Json_In := '{"input":{"no_consistence":1';
      v_Json_In := v_Json_In || ',"item_list":[' || v_Tmp || ']';
      v_Json_In := v_Json_In || v_Pati_List;
      v_Json_In := v_Json_In || '}}';
    
      Zl_病人费用销帐_Audit_Check(v_Json_In, v_Json_Out);
    
      j_Tmp    := PLJson();
      j_Output := PLJson();
      j_Tmp    := PLJson(v_Json_Out);
      j_Output := j_Tmp.Get_Pljson('output');
      If j_Output.Get_Number('code') = 0 Then
        v_Error := j_Output.Get_String('message');
        Raise Err_Custom;
      End If;
      Nt_自动审核 := 1;
      Rs_单据(Lngtmp).销帐 := 0;
    End If;
  
    --销帐申请列表
    v_Charge_List := v_Charge_List || ',{"outpati_account":' || Rs_单据(Lngtmp).门诊记帐;
    v_Charge_List := v_Charge_List || ',"fee_id":' || Rs_单据(Lngtmp).费用id;
    v_Charge_List := v_Charge_List || ',"fee_item_id":' || Rs_单据(Lngtmp).收费细目id;
    v_Charge_List := v_Charge_List || ',"request_dept_id":' || Rs_单据(Lngtmp).申请科室id;
    v_Charge_List := v_Charge_List || ',"audit_dept_id":' || Nvl(Rs_单据(Lngtmp).审核科室id || '', 'null');
    v_Charge_List := v_Charge_List || ',"request_type":' || Nvl(Rs_单据(Lngtmp).申请类别, 0);
    v_Charge_List := v_Charge_List || ',"request_num":' || zlJsonStr(Rs_单据(Lngtmp).数量, 1);
    v_Charge_List := v_Charge_List || ',"auto_aduit":' || Nt_自动审核;
    v_Charge_List := v_Charge_List || ',"request_time":"' || To_Char(Rs_单据(Lngtmp).申请时间, 'yyyy-mm-dd hh24:mi:ss') || '"';
    v_Charge_List := v_Charge_List || ',"reason":"' || zlJsonStr(Nt_销帐原因) || '"';
    v_Charge_List := v_Charge_List || '}';
  End;

  Procedure p_Get销帐信息(P收回数量 Out Number) As
    --需要进一步确定的值:v_剂量系数,v_住院包装,r_Detail.收费方式,r_Detail.诊疗类别,r_Detail.单次用量,p收回数量,v_对照数量
    --该方法内部会对这些变量进行一次赋值 v_剂量系数,v_住院包装,收费方式,诊疗类别,p收回数量,v_对照数量
  Begin
    r_Detail.诊疗类别 := Nt_诊疗类别;
    r_Detail.单次用量 := Nt_单次用量;
    r_Detail.医嘱内容 := Nt_医嘱内容;
    --药品收回总量是以最后发送规格为准计算的，以此计算出收回售价数量
    Begin
      Select 剂量系数, 住院包装 Into v_剂量系数, v_住院包装 From 药品规格 Where 药品id = r_Detail.收费细目id;
    Exception
      When Others Then
        v_剂量系数 := 1;
        v_住院包装 := 1;
    End;
  
    --从医嘱计价中获取收费方式和对照数量
    v_对照数量        := Null;
    r_Detail.收费方式 := Null;
    For Lngtmp In 1 .. Rs_计价.Count Loop
      If Rs_计价(Lngtmp).收费细目id = r_Detail.收费细目id And Rs_计价(Lngtmp).医嘱id = 医嘱id_In Then
        v_对照数量        := Rs_计价(Lngtmp).对照数量;
        r_Detail.收费方式 := Rs_计价(Lngtmp).收费方式;
      End If;
    End Loop;
    v_对照数量        := Nvl(v_对照数量, 1);
    r_Detail.收费方式 := Nvl(r_Detail.收费方式, 0);
    r_Detail.已执行量 := 0;
    If r_Detail.收费类别 In ('5', '6', '7') Or r_Detail.收费类别 = '4' And r_Detail.跟踪在用 = 1 Then
      For Rl In 1 .. Rs_执行.Count Loop
        If Rs_执行(Rl).费用id = r_Detail.费用id And Rs_执行(Rl).医嘱id = 医嘱id_In Then
          r_Detail.已执行量 := Rs_执行(Rl).已执行量;
        End If;
      End Loop;
    End If;
  
    --计算收回数量，如果收费方式不是0正常收取，则使用特殊的方法进行计算
    If r_Detail.收费方式 = 0 Then
      If r_Detail.诊疗类别 = '7' Then
        --中药配方药品：付数*单量
        P收回数量 := Round(收回量_In * r_Detail.单次用量 / Nvl(v_剂量系数, 1), 5);
      Else
        P收回数量 := Round(收回量_In * Nvl(v_住院包装, 1), 5) * v_对照数量;
      End If;
    Else
      P收回数量 := 0;
      For Lngtmp In 1 .. Rs_收回.Count Loop
        If Rs_收回(Lngtmp).收费细目id = r_Detail.收费细目id And Rs_收回(Lngtmp).医嘱id = 医嘱id_In Then
          P收回数量 := Rs_收回(Lngtmp).收回数量;
        End If;
      End Loop;
      P收回数量 := Round(P收回数量, 5);
    End If;
  End;

  Procedure p_产生销帐中间数据
  (
    P执行标志 Number,
    P数量     Number
  ) As
  Begin
    --其它中间数据可通过 r_Detail 类型获取
    --系统参数决定执行后是否审核划价单，所以，已执行的仍然可能是划价单
    If P执行标志 = 0 And r_Detail.记录状态 = 0 Then
      r_Detail.销帐数量 := P数量;
      r_Detail.销帐模式 := 0;
      --销帐申请时间,申请类别,销帐数量,静配自动审核,销帐模式
      p_Add_Delitem;
    Else
      If Not (r_Detail.收费类别 = '7' And P执行标志 <> 0) Then
        r_Detail.销帐数量 := P数量;
        r_Detail.销帐模式 := 1;
        r_Detail.执行标志 := P执行标志;
      
        If Nvl(No_In, '调整划价单') <> '调整划价单' And Nvl(n_静配销帐操作, 0) = 0 Then
          p_Add_Negbill;
        Else
          p_Add_Delitem;
        End If;
      End If;
    End If;
  End;

  Procedure p_部分执行拆分销帐(P收回量 Number) As
    Lng执行数量 Number;
    Lng未执行量 Number;
  Begin
  
    Lng执行数量 := Nvl(r_Detail.已执行量, 0);
    --销帐申请时间,申请类别,销帐数量,静配自动审核,销帐模式
    If Nvl(Lng执行数量, 0) <= 0 Then
      --都是未执行,则可以销
      r_Detail.申请类别 := 0;
      p_产生销帐中间数据(0, P收回量);
    Else
      Lng未执行量 := Nvl(r_Detail.剩余数量, 0) - Nvl(Lng执行数量, 0);
      If Lng未执行量 <= 0 Then
        r_Detail.申请类别 := 1;
        p_产生销帐中间数据(1, P收回量);
      Else
        --已执行数不会比费用的剩余数量大,如果超过当成全部已经执行,已经执行数也不会为负数为负数也当成是未执行
        If Lng未执行量 >= P收回量 Then
          --都是未执行,则可以销
          r_Detail.申请类别 := 0;
          p_产生销帐中间数据(0, P收回量);
        Else
          p_产生销帐中间数据(0, Lng未执行量);
          r_Detail.申请类别 := 1;
          p_产生销帐中间数据(1, P收回量 - Lng未执行量);
        End If;
      End If;
    End If;
  End;

  Procedure p_静配销帐
  (
    P收回量 Number,
    P是     Out Number
  ) As
    p_收回时间 Date;
    P静配销量  Number;
    P静配余量  Number;
  Begin
    P静配销量  := 0;
    P是        := 0;
    p_收回时间 := 收回时间_In;
    For R In 1 .. Rs_静配.Count Loop
      If Rs_静配(R).费用id = r_Detail.费用id And Rs_静配(R).收费细目id = r_Detail.收费细目id And Rs_静配(R).医嘱id = 医嘱id_In Then
        n_静配销帐操作        := 1;
        P静配销量             := P静配销量 + Rs_静配(R).数量;
        r_Detail.配液id       := Rs_静配(R).配液id;
        r_Detail.销帐申请时间 := p_收回时间;
        If Rs_静配(R).操作状态 = 1 Then
          r_Detail.申请类别     := 0;
          r_Detail.静配自动审核 := 1;
          p_产生销帐中间数据(0, Rs_静配(R).数量);
        Else
          r_Detail.申请类别     := 1;
          r_Detail.静配自动审核 := 0;
          p_产生销帐中间数据(1, Rs_静配(R).数量);
        End If;
        r_Detail.配液id := Null;
        If Instr(',' || v_Pivas_Ids || ',', ',' || Rs_静配(R).配液id || ',') = 0 Then
          v_Pivas_Out := v_Pivas_Out || ',{"pivas_id":' || Rs_静配(R).配液id;
          v_Pivas_Out := v_Pivas_Out || ',"auto_aduit":' || r_Detail.静配自动审核;
          v_Pivas_Out := v_Pivas_Out || ',"request_time":"' || To_Char(p_收回时间, 'yyyy-mm-dd hh24:mi:ss') || '"';
          v_Pivas_Out := v_Pivas_Out || ',"reason":"' || zlJsonStr(Nt_销帐原因) || '"';
          v_Pivas_Out := v_Pivas_Out || '}';
          v_Pivas_Ids := v_Pivas_Ids || ',' || Rs_静配(R).配液id;
        End If;
        p_收回时间     := p_收回时间 + 1 / 24 / 60 / 60;
        n_静配销帐操作 := 0;
      End If;
    End Loop;
  
    If P静配销量 <> 0 Then
      --剩下部分都采用销申请方式
      P静配余量 := P收回量 - P静配销量;
      If P静配余量 > 0 Then
        r_Detail.静配自动审核 := 0;
        r_Detail.申请类别     := 1;
        r_Detail.销帐申请时间 := 收回时间_In;
        p_产生销帐中间数据(1, P静配余量);
      End If;
      P是 := 1;
    End If;
    --变量还原
    r_Detail.静配自动审核 := Null;
    r_Detail.申请类别     := Null;
    r_Detail.销帐申请时间 := 收回时间_In;
  End;

  Procedure p_Get_Json_Out As
  Begin
    If Rs_单据.Count > 0 Then
      For Rp In 1 .. Rs_单据.Count Loop
        If Nvl(Rs_单据(Rp).销帐, 0) = 1 Then
          p_Charge_Check(Rp);
        
        End If;
        If Nvl(Rs_单据(Rp).销帐, 0) = 0 Then
          p_Delbill_Check(Rp);
        End If;
      End Loop;
    End If;
    v_Json_Out := '{"code":1,"message":"成功"';
    p_Getoutlist;
    v_Json_Out := v_Json_Out || '}';
    Json_Out   := '{"output":' || v_Json_Out || '}';
  End;

Begin

  --解析入参
  j_Tmp      := PLJson(Json_In);
  j_Json     := j_Tmp.Get_Pljson('input');
  n_检查方式 := j_Json.Get_Number('check_type');

  If 1 = n_检查方式 Then
    c_医嘱ids := j_Json.Get_Clob('order_ids');
    l_医嘱ids := t_StrList();
    While c_医嘱ids Is Not Null Loop
      If Length(c_医嘱ids) <= 4000 Then
        l_医嘱ids.Extend;
        l_医嘱ids(l_医嘱ids.Count) := c_医嘱ids;
        c_医嘱ids := Null;
      Else
        l_医嘱ids.Extend;
        l_医嘱ids(l_医嘱ids.Count) := Substr(c_医嘱ids, 1, Instr(c_医嘱ids, ',', 3980) - 1);
        c_医嘱ids := Substr(c_医嘱ids, Instr(c_医嘱ids, ',', 3980) + 1);
      End If;
    End Loop;
    For I In 1 .. l_医嘱ids.Count Loop
      For R In (Select a.医嘱序号 As 医嘱id
                From 住院费用记录 A
                Where a.记录性质 In (2, 12) And a.记录状态 = 1 And
                      a.医嘱序号 In (Select /*+cardinality(b,10) */
                                  Column_Value
                                 From Table(f_Num2List(l_医嘱ids(I))) B)
                
                Group By a.医嘱序号
                Having Sum(Nvl(a.结帐金额, 0)) <> 0) Loop
        Vo_Vals := Vo_Vals || ',' || r.医嘱id;
      End Loop;
      For R In (Select a.医嘱序号 As 医嘱id
                From 门诊费用记录 A
                Where a.记录性质 In (2, 12) And a.记录状态 = 1 And
                      a.医嘱序号 In (Select /*+cardinality(b,10) */
                                  Column_Value
                                 From Table(f_Num2List(l_医嘱ids(I))) B)
                
                Group By a.医嘱序号
                Having Sum(Nvl(a.结帐金额, 0)) <> 0) Loop
        Vo_Vals := Vo_Vals || ',' || r.医嘱id;
      End Loop;
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"成功","order_ids":"' || Substr(Vo_Vals, 2) || '"}}';
  Elsif 2 = n_检查方式 Then
    --检查收回次数的医嘱对应费用是否全是未审核的划价单，以便确定直接修改划价单，无需取新的单据号
    收回量_In := j_Json.Get_Number('roll_num');
    医嘱id_In := j_Json.Get_String('order_ids'); --目前是一个医嘱传一次，后续优化
    Select Sum(a.剩余数量) 剩余数量
    Into v_当前数量
    From (Select Nvl(a.付数, 1) * a.数次 As 剩余数量
           From 住院费用记录 A
           Where a.医嘱序号 = 医嘱id_In And a.记录状态 = 0
           Union All
           Select Nvl(a.付数, 1) * a.数次 As 剩余数量
           From 门诊费用记录 A
           Where a.医嘱序号 = 医嘱id_In And a.记录状态 = 0) A;
    If Nvl(收回量_In, 0) > Nvl(v_当前数量, 0) Then
      医嘱id_In := Null;
    End If;
    --说明：order_ids 返回的结点为空值说明不能改划价单，要产生负数
    Json_Out := '{"output":{"code":1,"message":"成功","order_ids":"' || 医嘱id_In || '"}}';
  Else
  
    Rs_计价    := t_计价();
    Rs_收回    := t_计价();
    Rs_执行    := t_执行();
    Rs_单据    := t_单据();
    Rs_静配    := t_静配();
    Rs_其它    := t_单据();
    Rs_Dellist := t_单据();
    j_List     := j_Json.Get_Pljson_List('price_list');
    If j_List Is Not Null Then
      For I In 1 .. j_List.Count Loop
        j_Item := PLJson();
        j_Item := PLJson(j_List.Get(I));
        Rs_计价.Extend;
        Lngtmp := Rs_计价.Count;
        Rs_计价(Lngtmp).收费细目id := j_Item.Get_Number('fee_item_id');
        Rs_计价(Lngtmp).对照数量 := j_Item.Get_Number('refer_num');
        Rs_计价(Lngtmp).收费方式 := j_Item.Get_Number('fee_way');
        Rs_计价(Lngtmp).医嘱id := j_Item.Get_Number('order_id');
      End Loop;
    End If;
  
    --绑定对照费用的收回数量
    j_List := Pljson_List();
    j_List := j_Json.Get_Pljson_List('price_exe_list');
    If j_List Is Not Null Then
      For I In 1 .. j_List.Count Loop
        j_Item := PLJson();
        j_Item := PLJson(j_List.Get(I));
        Rs_收回.Extend;
        Lngtmp := Rs_收回.Count;
        Rs_收回(Lngtmp).收费细目id := j_Item.Get_Number('fee_item_id');
        Rs_收回(Lngtmp).收回数量 := j_Item.Get_Number('roll_num');
        Rs_收回(Lngtmp).医嘱id := j_Item.Get_Number('order_id');
      End Loop;
    End If;
  
    --静配列表
    j_List := Pljson_List();
    j_List := j_Json.Get_Pljson_List('pivas_list');
    If j_List Is Not Null Then
      For I In 1 .. j_List.Count Loop
        j_Item := PLJson();
        j_Item := PLJson(j_List.Get(I));
        Rs_静配.Extend;
        Lngtmp := Rs_静配.Count;
        Rs_静配(Lngtmp).费用id := j_Item.Get_Number('fee_id');
        Rs_静配(Lngtmp).数量 := j_Item.Get_Number('quantity');
        Rs_静配(Lngtmp).收费细目id := j_Item.Get_Number('fee_item_id');
        Rs_静配(Lngtmp).配液id := j_Item.Get_Number('pivas_id');
        Rs_静配(Lngtmp).操作状态 := j_Item.Get_Number('operator_status');
        Rs_静配(Lngtmp).医嘱id := j_Item.Get_Number('order_id');
      End Loop;
    End If;
  
    --病人列表
    j_List := Pljson_List();
    j_List := j_Json.Get_Pljson_List('pati_list');
    If j_List Is Not Null Then
      v_Pati_List := j_List.To_Char(False);
      If v_Pati_List Is Not Null Then
        v_Pati_List := ',"pati_list":' || v_Pati_List;
      End If;
    End If;
  
    --药品卫材执行列表
    j_List := Pljson_List();
    j_List := j_Json.Get_Pljson_List('excute_list');
    If j_List Is Not Null Then
      For I In 1 .. j_List.Count Loop
        j_Item := PLJson();
        j_Item := PLJson(j_List.Get(I));
        Rs_执行.Extend;
        Lngtmp := Rs_执行.Count;
        Rs_执行(Lngtmp).费用id := j_Item.Get_Number('fee_id');
        Rs_执行(Lngtmp).已执行量 := j_Item.Get_Number('sended_num');
        Rs_执行(Lngtmp).医嘱id := j_Item.Get_Number('order_id');
      End Loop;
      v_Excute_List := j_List.To_Char(False);
      If v_Excute_List Is Not Null Then
        v_Excute_List := ',"excute_list":' || v_Excute_List;
      End If;
    End If;
  
    --医嘱执行列表
    j_List := Pljson_List();
    j_List := j_Json.Get_Pljson_List('advice_excute_list');
    If j_List Is Not Null Then
      v_Advice_Excute_List := j_List.To_Char(False);
      If v_Advice_Excute_List Is Not Null Then
        v_Advice_Excute_List := ',"advice_excute_list":' || v_Advice_Excute_List;
      End If;
    End If;
  
    --其它发送列表
    j_List := Pljson_List();
    j_List := j_Json.Get_Pljson_List('other_send_list');
    If j_List Is Not Null Then
      For I In 1 .. j_List.Count Loop
        j_Item := PLJson();
        j_Item := PLJson(j_List.Get(I));
        Rs_其它.Extend;
        Lngtmp := Rs_其它.Count;
        Rs_其它(Lngtmp).No := j_Item.Get_String('fee_no');
        Rs_其它(Lngtmp).数量 := j_Item.Get_Number('send_num');
        Rs_其它(Lngtmp).医嘱id := j_Item.Get_Number('order_id');
      End Loop;
    End If;
    v_自审 := zl_GetSysParameter('超期收回费用本科自动审核', 1254);
    --开始按医嘱循环收回
    j_List_Order := j_Json.Get_Pljson_List('order_list');
    For Lp_Order In 1 .. j_List_Order.Count Loop
      j_Item := PLJson();
      j_Item := PLJson(j_List_Order.Get(Lp_Order));
    
      收回量_In := j_Item.Get_Number('roll_num');
      If 收回量_In <= 0 Then
        v_Error := '要收回的数量为零请检查';
        Raise Err_Custom;
      End If;
    
      No_In         := j_Item.Get_String('fee_no'); --负数方式冲销的单据号
      v_医嘱内容    := j_Item.Get_String('advice_note');
      医嘱id_In     := j_Item.Get_Number('order_id'); --只传一个医嘱id
      收回时间_In   := To_Date(j_Item.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');
      v_Orderfeenos := j_Item.Get_String('fee_nos');
      Nt_销帐原因   := j_Item.Get_String('reason');
    
      Nt_单次用量  := j_Item.Get_Number('advice_dosage');
      Nt_医嘱内容  := j_Item.Get_String('advice_note');
      Nt_诊疗类别  := j_Item.Get_String('clinic_type');
      Nt_跟踪在用  := j_Item.Get_Number('is_stuff_order');
      v_收费细目id := Null;
    
      If No_In Is Null Then
        --a.销帐申请收回模式
        --输液配药记录单独进行销帐
        v_结帐参数 := zl_GetSysParameter(23);
        --根据收回数量对照原始费用进行分摊申请
      
        For r_Fee In c_Detail Loop
          --赋值
          Select r_Fee.费用id, r_Fee.No, r_Fee.序号, r_Fee.收费细目id, r_Fee.病人病区id, r_Fee.收费类别, r_Fee.跟踪在用, r_Fee.诊疗类别,
                 r_Fee.医嘱内容, r_Fee.单次用量, r_Fee.剩余数量, r_Fee.已执行量, r_Fee.未执行量, r_Fee.执行标志, r_Fee.记录状态, r_Fee.登记时间,
                 r_Fee.收费方式, r_Fee.门诊记帐, 收回时间_In, Null, Null, Null, Null
          Into r_Detail.费用id, r_Detail.No, r_Detail.序号, r_Detail.收费细目id, r_Detail.病人病区id, r_Detail.收费类别, r_Detail.跟踪在用,
               r_Detail.诊疗类别, r_Detail.医嘱内容, r_Detail.单次用量, r_Detail.剩余数量, r_Detail.已执行量, r_Detail.未执行量, r_Detail.执行标志,
               r_Detail.记录状态, r_Detail.登记时间, r_Detail.收费方式, r_Detail.门诊记帐, r_Detail.销帐申请时间, r_Detail.申请类别, r_Detail.销帐数量,
               r_Detail.静配自动审核, r_Detail.销帐模式
          From Dual;
        
          --需要进一步确定的值:v_剂量系数,v_住院包装,收费方式,诊疗类别,v_收回数量,v_对照数量
          v_收回数量tmp := 0;
          p_Get销帐信息(v_收回数量tmp);
        
          --确定该收费细目id的收回总数量
          If Nvl(v_收费细目id, 0) <> r_Detail.收费细目id And (r_Detail.诊疗类别 Not In ('5', '6', '7') Or Nvl(v_收费细目id, 0) = 0) Then
            --数量未分摊完成
            If v_收费细目id Is Not Null And v_收回数量 > 0 Then
              v_Error := '医嘱"' || v_医嘱内容 || '"对应的费用剩余数量不足收回数量，可能存在手工销帐或已结帐的费用，或相应的划价单已被删除。';
              Raise Err_Custom;
            End If;
            v_收回数量 := v_收回数量tmp;
            v_医嘱内容 := r_Detail.医嘱内容;
          End If;
          --该收费细目的每个费用明细分摊收回
          If v_收回数量 > 0 Then
            --检查对应费用是否已结帐，当禁止时
            v_结帐金额 := 0;
            If v_结帐参数 = '2' And r_Detail.记录状态 <> 0 Then
              Select Sum(结帐金额)
              Into v_结帐金额
              From 住院费用记录
              Where NO = r_Detail.No And 记录性质 In (2, 12) And Nvl(价格父号, 序号) = r_Detail.序号;
            End If;
          
            If Nvl(v_结帐金额, 0) = 0 Then
              v_剩余数量 := r_Detail.剩余数量;
              If v_收回数量 > v_剩余数量 Then
                v_当前数量 := v_剩余数量;
              Else
                v_当前数量 := v_收回数量;
              End If;
              v_收回数量 := v_收回数量 - v_当前数量;
            
              If r_Detail.收费类别 In ('5', '6', '7') Or r_Detail.收费类别 = '4' And r_Detail.跟踪在用 = 1 Then
                p_静配销帐(v_当前数量, Lngtmp);
                If Lngtmp = 0 Then
                  --因为药品卫材存在部分执行的情况,这里可能就要被分为两次来调用,算数量时都当成统一看待
                  p_部分执行拆分销帐(v_当前数量);
                End If;
              Else
                p_产生销帐中间数据(r_Detail.执行标志, v_当前数量);
              End If;
              v_费用ids := v_费用ids || ',' || r_Detail.费用id;
            End If;
          End If;
          v_收费细目id := r_Detail.收费细目id;
        End Loop;
      
        --数量未分摊完成
        If v_收回数量 > 0 Then
          v_Error := '医嘱"' || v_医嘱内容 || '"对应的费用剩余数量不足收回数量，可能存在手工销帐或已结帐的费用，或相应的划价单已被删除。';
          Raise Err_Custom;
        End If;
      
      Elsif No_In = '调整划价单' Then
        --直接调整数量
        For r_Fee In c_Del Loop
          --赋值
          Select r_Fee.费用id, r_Fee.No, r_Fee.序号, r_Fee.收费细目id, Null, r_Fee.收费类别, r_Fee.跟踪在用, Null, Null, 0, r_Fee.剩余数量,
                 0, 0, 0, 0, Null, Null, r_Fee.门诊记帐, 收回时间_In, Null, Null, Null, Null
          Into r_Detail.费用id, r_Detail.No, r_Detail.序号, r_Detail.收费细目id, r_Detail.病人病区id, r_Detail.收费类别, r_Detail.跟踪在用,
               r_Detail.诊疗类别, r_Detail.医嘱内容, r_Detail.单次用量, r_Detail.剩余数量, r_Detail.已执行量, r_Detail.未执行量, r_Detail.执行标志,
               r_Detail.记录状态, r_Detail.登记时间, r_Detail.收费方式, r_Detail.门诊记帐, r_Detail.销帐申请时间, r_Detail.申请类别, r_Detail.销帐数量,
               r_Detail.静配自动审核, r_Detail.销帐模式
          From Dual;
          v_收回数量tmp := 0;
          p_Get销帐信息(v_收回数量tmp);
          r_Detail.已执行量 := 0; --是划价单都是未执行
          If Nvl(v_收费细目id, 0) <> r_Fee.收费细目id Then
            --数量未分摊完成
            If v_收费细目id Is Not Null And v_收回数量 > 0 Then
              v_Error := '医嘱"' || v_医嘱内容 || '"对应的费用剩余数量不足收回数量，可能相关划价单已被删除或审核。';
              Raise Err_Custom;
            End If;
            v_收回数量 := v_收回数量tmp;
            v_医嘱内容 := r_Detail.医嘱内容;
          End If;
          If v_收回数量 > 0 Then
            v_剩余数量 := r_Detail.剩余数量;
            If v_收回数量 > v_剩余数量 Then
              v_当前数量 := v_剩余数量;
            Else
              v_当前数量 := v_收回数量;
            End If;
            v_收回数量 := v_收回数量 - v_当前数量;
            If r_Detail.收费类别 In ('5', '6', '7') Or r_Detail.收费类别 = '4' And r_Detail.跟踪在用 = 1 Then
              p_静配销帐(v_当前数量, Lngtmp);
              If Lngtmp = 0 Then
                p_部分执行拆分销帐(v_当前数量);
              End If;
            Else
              p_产生销帐中间数据(r_Detail.执行标志, v_当前数量);
            End If;
          End If;
          v_收费细目id := r_Fee.收费细目id;
        End Loop;
        --数量未分摊完成
        If v_收回数量 > 0 Then
          v_Error := '医嘱"' || v_医嘱内容 || '"对应的费用剩余数量不足收回数量，可能相关划价单已被删除或审核。';
          Raise Err_Custom;
        End If;
      
      Elsif Nvl(No_In, '调整划价单') <> '调整划价单' Then
        Select zl_GetSysParameter(63) Into v_自动发料 From Dual;
        Select zl_GetSysParameter(80) Into v_划价类别 From Dual;
        --产生负数单据--负数冲销，可能存在划价单与记帐单混合的情况
        --负数冲销的时候可能要两边分开冲,负数冲销的时候没得数量限制，有多少就冲多少即使有剩余也不中断
        --费用记录和收发记录是一对一的关系,静配相关是已经限制不能负数冲销的.
        --负数冲销就是把销帐申请那部分采用负单据方式来冲掉,整体模没变,销帐申请部分改为生成单据,还有一点差别就是不分摊数量
        --费用序号，收发序号，顺序产生，
        --一条医嘱的药品只有一行，这里的循环是为了处理多次发送的情况，分批药品在界面已禁用负数收回
        If Nt_诊疗类别 In ('5', '6', '7') Or (Nt_诊疗类别 = '4' And Nvl(Nt_跟踪在用, 0) = 1) Then
        
          Select Decode(Nvl(Instr(v_划价类别, Decode(Nt_诊疗类别, '4', '4', '5')), 0), 0, 1, 0)
          Into Nt_收费标志
          From Dual;
        
          For r_Drug In c_Negdrug Loop
            --赋值
            Select r_Drug.费用id, r_Drug.No, r_Drug.序号, r_Drug.收费细目id, r_Drug.病人病区id, r_Drug.收费类别, r_Drug.跟踪在用, Null, Null,
                   0, r_Drug.剩余数量, 0, 0, 0, r_Drug.记录状态, Null, Null, r_Drug.门诊记帐, 收回时间_In, Null, Null, Null, Null,
                   r_Drug.执行部门id
            Into r_Detail.费用id, r_Detail.No, r_Detail.序号, r_Detail.收费细目id, r_Detail.病人病区id, r_Detail.收费类别, r_Detail.跟踪在用,
                 r_Detail.诊疗类别, r_Detail.医嘱内容, r_Detail.单次用量, r_Detail.剩余数量, r_Detail.已执行量, r_Detail.未执行量, r_Detail.执行标志,
                 r_Detail.记录状态, r_Detail.登记时间, r_Detail.收费方式, r_Detail.门诊记帐, r_Detail.销帐申请时间, r_Detail.申请类别,
                 r_Detail.销帐数量, r_Detail.静配自动审核, r_Detail.销帐模式, r_Detail.执行部门id
            From Dual;
          
            v_收回数量tmp := 0;
            p_Get销帐信息(v_收回数量tmp);
            v_收回数量 := v_收回数量tmp;
          
            If v_收回数量 > 0 Then
              v_剩余数量 := r_Detail.剩余数量;
              If v_收回数量 > v_剩余数量 Then
                v_当前数量 := v_剩余数量;
              Else
                v_当前数量 := v_收回数量;
              End If;
              v_收回数量 := v_收回数量 - v_当前数量;
              If r_Detail.收费类别 In ('5', '6', '7') Or r_Detail.收费类别 = '4' And r_Detail.跟踪在用 = 1 Then
                p_静配销帐(v_当前数量, Lngtmp);
                If Lngtmp = 0 Then
                  p_部分执行拆分销帐(v_当前数量);
                End If;
              Else
                p_产生销帐中间数据(r_Detail.执行标志, v_当前数量);
              End If;
            End If;
            If v_收回数量 <= 0 Then
              Exit;
            End If;
          End Loop;
        
          If v_收回数量 <> 0 Then
            --没有收回所有数量,收发记录本身有问题(如记录不全或数量为负)
            Null;
          End If;
        Else
          --非药品卫材
          --药品卫材执行列表
          v_收回剩余 := 收回量_In;
          For I其它 In 1 .. Rs_其它.Count Loop
            If Rs_其它(I其它).医嘱id = 医嘱id_In Then
              Nt_发送单据 := Rs_其它(I其它).No;
              Nt_发送数次 := Rs_其它(I其它).数量;
              If Nt_发送数次 < v_收回剩余 Then
                --一次收回多次发送，但是每次发送费用有所变动（计价）
                v_收回剩余 := v_收回剩余 - Nt_发送数次;
                v_收回量   := Nt_发送数次;
              Else
                --一次发送中收回剩余；
                v_收回量   := v_收回剩余;
                v_收回剩余 := 0;
              End If;
            
              v_收费内容 := '';
              For r_Other In c_Other Loop
                Nt_收费细目id := r_Other.收费细目id;
                If Nvl(v_收费内容, '0') <> r_Other.收费细目id || ',' || r_Other.序号 Then
                  --和普通的区别,不分药品剂量系数关系,无已经执行数和未执行数区分
                  --从医嘱计价中获取收费方式和对照数量
                  v_对照数量  := Null;
                  Nt_收费方式 := Null;
                  For Lngtmp In 1 .. Rs_计价.Count Loop
                    If Rs_计价(Lngtmp).收费细目id = Nt_收费细目id And Rs_计价(Lngtmp).医嘱id = 医嘱id_In Then
                      v_对照数量  := Rs_计价(Lngtmp).对照数量;
                      Nt_收费方式 := Rs_计价(Lngtmp).收费方式;
                    End If;
                  End Loop;
                  v_对照数量  := Nvl(v_对照数量, 1);
                  Nt_收费方式 := Nvl(Nt_收费方式, 0);
                  --根据最近一次发送的费用记录，按需要收回的数量全部收回
                  --计算收回数量，如果收费方式不是0正常收取，则使用特殊的方法进行计算
                  If Nt_收费方式 = 0 Then
                    --重新获取收费方式和对照数量
                    v_收回数量 := v_收回量 * Nvl(v_对照数量, 1);
                  Else
                    v_收回数量 := 0;
                    For Lngtmp In 1 .. Rs_收回.Count Loop
                      If Rs_收回(Lngtmp).收费细目id = Nt_收费细目id And Rs_收回(Lngtmp).医嘱id = 医嘱id_In Then
                        v_收回数量 := Rs_收回(Lngtmp).收回数量;
                      End If;
                    End Loop;
                    v_收回数量 := Round(v_收回数量, 5);
                  End If;
                End If;
              
                If v_收回数量 > 0 Then
                  If r_Other.记录状态 = 0 Then
                    If v_收回数量 > r_Other.剩余数量 Then
                      v_当前数量 := r_Other.剩余数量;
                    Else
                      v_当前数量 := v_收回数量;
                    End If;
                  Else
                    v_当前数量 := v_收回数量;
                  End If;
                  v_收回数量 := v_收回数量 - v_当前数量;
                
                  --赋值
                  Select r_Other.费用id, r_Other.No, r_Other.序号, r_Other.收费细目id, r_Other.病人病区id, r_Other.收费类别,
                         r_Other.跟踪在用, Null, Null, 0, 0, 0, 0, r_Other.执行标志, 0, Null, Null, r_Other.门诊记帐, 收回时间_In, Null,
                         Null, Null, Null, r_Other.执行部门id
                  Into r_Detail.费用id, r_Detail.No, r_Detail.序号, r_Detail.收费细目id, r_Detail.病人病区id, r_Detail.收费类别,
                       r_Detail.跟踪在用, r_Detail.诊疗类别, r_Detail.医嘱内容, r_Detail.单次用量, r_Detail.剩余数量, r_Detail.已执行量,
                       r_Detail.未执行量, r_Detail.执行标志, r_Detail.记录状态, r_Detail.登记时间, r_Detail.收费方式, r_Detail.门诊记帐,
                       r_Detail.销帐申请时间, r_Detail.申请类别, r_Detail.销帐数量, r_Detail.静配自动审核, r_Detail.销帐模式, r_Detail.执行部门id
                  From Dual;
                
                  Select Decode(Nvl(Instr(v_划价类别, r_Detail.收费类别), 0), 0, 1, 0) Into Nt_收费标志 From Dual;
                
                  If r_Detail.执行标志 = 1 Then
                    Nt_收费标志 := 1;
                  End If;
                  If r_Other.记录状态 = 0 Then
                    p_产生销帐中间数据(r_Other.记录状态, v_当前数量);
                  Else
                    p_产生销帐中间数据(1, v_当前数量);
                  End If;
                  v_收费内容 := r_Other.收费细目id || ',' || r_Other.序号;
                End If;
              End Loop;
              If v_收回剩余 <= 0 Then
                Exit;
              End If;
            End If;
          End Loop;
        
        End If;
      End If;
    End Loop;
    p_Get_Json_Out;
  End If;
Exception
  When Err_Custom Then
    Json_Out := '{"output":{"code":0,"message":"' || zlJsonStr(v_Error) || '"}}';
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Odr_Check;
/

Create Or Replace Procedure Zl_Exsesvr_Overdue_Recovery
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：超期发送收回费用相关处理
  --入参：Json_In:格式
  --  input
  --     operator_name                      C 1 操作员姓名
  --     operator_code                      C 1 操作员编号
  --     operator_time                      C 1 操作时间
  --     charge_list[]销帐申请列表
  --                  outpati_account        N 1 门诊记帐0-住院记账,1-门诊记帐
  --                  fee_id                 N 1 费用id
  --                  fee_item_id            N 1 收费细目id
  --                  request_dept_id        N 1 申请科室id
  --                  audit_dept_id          N 1 审核科室id
  --                  request_num            N 1 申请数量
  --    del_list[]单据删除列表
  --                  outpati_account        N 1 门诊记帐0-住院记账,1-门诊记帐
  --                  fee_no                 C 1 费用单据号
  --                  serial_num             C 1 删除序号,序号格式:数量:执行状态
  --    roll_list[]负数单据列表
  --                  outpati_account        N 1 门诊记帐0-住院记账,1-门诊记帐
  --                  clinic_type            C 1 医嘱诊疗类别
  --                  fee_no                 C 1 单据号
  --                  item_type              C 1 收费细目类别
  --                  fee_id                 N 1 费用id
  --                  fee_id_old             N 1 费用id,原始费用id
  --                  packages_num           N 1 付数
  --                  send_num               N 1 数次
  --                  is_stuff_order         N 1 区分是否是绑定的卫材费用0-非卫材医嘱,1-卫材医嘱
  --                  stuff_used             N 1 是否是跟踪在在卫才费用
  --                  exe_status             N 1 执行状态

  --出参: Json_Out,格式如下
  --  output
  --    code                          N 1 应答吗：0-失败；1-成功
  --    message                       C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  j_Input        Pljson;
  j_Item         Pljson;
  j_List         Pljson_List := Pljson_List();
  No_In          住院费用记录.No%Type;
  v_No           住院费用记录.No%Type;
  n_费用id       住院费用记录.Id%Type;
  n_收费细目id   住院费用记录.Id%Type;
  n_申请部门id   住院费用记录.Id%Type;
  n_审核部门id   住院费用记录.Id%Type;
  n_数量         住院费用记录.数次%Type;
  v_申请人       Varchar2(300);
  d_申请时间     病人费用销帐.申请时间%Type;
  n_申请类别     病人费用销帐.申请类别%Type;
  v_销帐原因     病人费用销帐.销帐原因%Type;
  n_自动审核     Number;
  n_门诊记帐     Number;
  v_序号         Varchar2(30000);
  v_操作员编号   Varchar2(300);
  v_操作员姓名   Varchar2(300);
  v_人员编号     Varchar2(300);
  v_人员姓名     Varchar2(300);
  d_登记时间     Date;
  d_操作时间     Date;
  v_费用序号     Number;
  v_费用id       住院费用记录.Id%Type;
  Old_费用id     住院费用记录.Id%Type;
  v_Dec          Number;
  v_划价类别     Varchar2(3000);
  v_自动发料     Varchar2(4000);
  v_开始序号     Number;
  n_类型         Number;
  v_结束序号     Number;
  收回时间_In    Date;
  v_Temp         Varchar2(4000);
  v_诊疗类别     Varchar2(4000);
  v_收费类别     Varchar2(300);
  v_当前付数     Number;
  v_当前数量     Number;
  v_实收金额     Number;
  n_跟踪在用医嘱 Number;
  n_跟踪在用费用 Number;
  n_执行状态     Number;
  v_医嘱执行     Number;
  n_记录状态     Number;
  n_划价         Number;

  n_执行状态费用 Number(2);
  d_执行时间     Date;
  v_执行人       Varchar2(300);

  --该游标用于处理费用相关汇总表
  Cursor c_Money
  (
    v_Start 住院费用记录.序号%Type,
    v_End   住院费用记录.序号%Type
  ) Is
    Select 病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, Sum(Nvl(应收金额, 0)) As 应收金额, Sum(Nvl(实收金额, 0)) As 实收金额
    From 住院费用记录
    Where 记录性质 = 2 And 记录状态 = 1 And NO = No_In And 序号 Between v_Start And v_End
    Group By 病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id
    Union All
    Select 病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, Sum(Nvl(应收金额, 0)) As 应收金额, Sum(Nvl(实收金额, 0)) As 实收金额
    From 门诊费用记录
    Where 记录性质 = 2 And 记录状态 = 1 And NO = No_In And 序号 Between v_Start And v_End
    Group By 病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id;

Begin
  j_Item  := Pljson(Json_In);
  j_Input := j_Item.Get_Pljson('input');

  v_操作员编号 := j_Input.Get_String('operator_code');
  v_操作员姓名 := j_Input.Get_String('operator_name');
  d_操作时间   := To_Date(j_Input.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');

  v_人员编号 := v_操作员编号;
  v_人员姓名 := v_操作员姓名;

  v_申请人    := v_操作员姓名;
  d_登记时间  := d_操作时间;
  收回时间_In := d_登记时间;

  --销帐列表
  j_List := j_Input.Get_Pljson_List('charge_list');
  If j_List Is Not Null Then
    For I In 1 .. j_List.Count Loop
      j_Item       := Pljson();
      j_Item       := Pljson(j_List.Get(I));
      n_费用id     := j_Item.Get_Number('fee_id');
      n_收费细目id := j_Item.Get_Number('fee_item_id');
      n_申请部门id := j_Item.Get_Number('request_dept_id');
      n_审核部门id := j_Item.Get_Number('audit_dept_id');
      n_申请类别   := j_Item.Get_Number('request_type');
      n_数量       := j_Item.Get_Number('request_num');
      n_自动审核   := j_Item.Get_Number('auto_aduit');
      d_申请时间   := To_Date(j_Item.Get_String('request_time'), 'yyyy-mm-dd hh24:mi:ss');
      v_销帐原因   := j_Item.Get_String('reason'); --销帐原因
      Zl_病人费用销帐_Insert_s(n_费用id, n_收费细目id, n_申请部门id, n_数量, v_申请人, d_申请时间, n_申请类别, v_销帐原因, n_审核部门id, 2);
      If n_自动审核 = 1 Then
        Zl_病人费用销帐_Audit_s(n_费用id, d_申请时间, v_申请人, d_申请时间, 1, n_申请类别);
      End If;
    End Loop;
  End If;

  --费用删除列表
  j_List := Pljson_List();
  j_List := j_Input.Get_Pljson_List('del_list');
  If j_List Is Not Null Then
    For I In 1 .. j_List.Count Loop
      j_Item     := Pljson();
      j_Item     := Pljson(j_List.Get(I));
      n_门诊记帐 := j_Item.Get_Number('outpati_account');
      v_No       := j_Item.Get_String('fee_no');
      v_序号     := j_Item.Get_String('serial_num');
      If n_门诊记帐 = 1 Then
        Zl_门诊记帐记录_Delete_s(v_No, v_序号, v_操作员编号, v_操作员姓名, d_登记时间, 2);
      Else
        Zl_住院记帐记录_Delete_s(v_No, v_序号, v_操作员编号, v_操作员姓名, 2, 0, d_登记时间);
      End If;
    End Loop;
  End If;

  --费用负数冲销列表
  j_List := Pljson_List();
  j_List := j_Input.Get_Pljson_List('roll_list');
  If j_List Is Not Null Then
    --负数冲销，可能存在划价单与记帐单混合的情况
    --金额小数位数
    Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into v_Dec From Dual;
    --生成划价单系统参数
    Select zl_GetSysParameter(80) Into v_划价类别 From Dual;
    Select zl_GetSysParameter(63) Into v_自动发料 From Dual;
  
    For I In 1 .. j_List.Count Loop
      j_Item := Pljson();
      j_Item := Pljson(j_List.Get(I));
    
      --这里还需要按NO_In进行排序,按NO来提交数据
      v_诊疗类别     := j_Item.Get_String('clinic_type');
      No_In          := j_Item.Get_String('fee_no');
      v_收费类别     := j_Item.Get_String('item_type');
      v_费用id       := j_Item.Get_Number('fee_id');
      Old_费用id     := j_Item.Get_Number('fee_id_old');
      v_当前付数     := j_Item.Get_Number('packages_num');
      v_当前数量     := j_Item.Get_Number('send_num');
      n_门诊记帐     := j_Item.Get_Number('outpati_account');
      n_跟踪在用医嘱 := j_Item.Get_Number('is_stuff_order');
      n_跟踪在用费用 := j_Item.Get_Number('stuff_used');
      v_医嘱执行     := j_Item.Get_Number('exe_status');
    
      Select Decode(n_执行状态, 1, Decode(v_收费类别, '4', Decode(n_跟踪在用费用, 1, 0, 1), Decode(Instr(',5,6,7,', v_收费类别), 0, 1, 0)),
                     0)
      Into n_执行状态费用
      From Dual;
    
      If v_收费类别 = '4' And n_跟踪在用费用 = 1 Then
        If v_自动发料 = '1' Then
          n_执行状态费用 := 1;
        End If;
      End If;
    
      If n_执行状态费用 = 1 Then
        d_执行时间 := d_操作时间;
        v_执行人   := v_操作员姓名;
      Else
        d_执行时间 := Null;
        v_执行人   := Null;
      End If;
    
      If n_门诊记帐 = 1 Then
        --门诊费用记录
        -------------------------------------------------------------------------------------
      
        Select Nvl(Max(序号), 0) + 1
        Into v_费用序号
        From 门诊费用记录
        Where 记录性质 = 2 And 记录状态 In (0, 1) And NO = No_In;
      
        --记录序号范围以处理汇总表
        If v_开始序号 Is Null Then
          v_开始序号 := v_费用序号;
        End If;
        v_结束序号 := v_费用序号;
      
        If v_诊疗类别 In ('5', '6', '7') Or n_跟踪在用医嘱 = 1 Then
          Select Nvl(Instr(v_划价类别, Decode(v_诊疗类别, '4', '4', '5')), 0) Into n_划价 From Dual;
          Insert Into 门诊费用记录
            (是否急诊, 结论, 记帐单id, 医嘱期效, 是否保密, 批次, ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 门诊标志, 病人id, 主页id, 标识号, 姓名, 性别, 年龄,
             病人病区id, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id, 付数, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额,
             实收金额, 统筹金额, 记帐费用, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 执行时间, 医嘱序号, 划价人, 操作员编号, 操作员姓名)
            Select 是否急诊, 结论, 记帐单id, 医嘱期效, 是否保密, 批次, v_费用id, 2, No_In, Decode(n_划价, 1, 0, 1), v_费用序号, Null, Null, 1, 病人id,
                   主页id, 标识号, 姓名, 性别, 年龄, 病人病区id, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id, v_当前付数, -1 * v_当前数量,
                   加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, Round(-1 * v_当前付数 * v_当前数量 * 标准单价, v_Dec),
                   Round(-1 * v_当前付数 * v_当前数量 * 标准单价, v_Dec), Null, 1, 开单部门id, 开单人, 收回时间_In, 收回时间_In, 执行部门id, v_执行人,
                   n_执行状态费用, d_执行时间, 医嘱序号, Decode(n_划价, 1, v_人员姓名, Null), Decode(n_划价, 1, Null, v_人员编号),
                   Decode(n_划价, 1, Null, v_人员姓名)
            From 门诊费用记录
            Where ID = Old_费用id;
        Else
          --非药品卫材医嘱
          --医嘱已执行，收回的费用也填为已执行：不包含药品和跟踪在用的卫材，因为实际发放表示执行
          --根据执行状态直接更新为记帐单，不用再单独审核记帐划价单
          n_执行状态 := v_医嘱执行;
          Select Decode(n_执行状态, 1, 1, Decode(Nvl(Instr(v_划价类别, v_收费类别), 0), 0, 1, 0))
          Into n_记录状态
          From Dual;
        
          Insert Into 门诊费用记录
            (是否急诊, 结论, 记帐单id, 医嘱期效, 是否保密, 批次, ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 门诊标志, 病人id, 主页id, 标识号, 姓名, 性别, 年龄,
             病人病区id, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id, 付数, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额,
             实收金额, 统筹金额, 记帐费用, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 执行时间, 执行人, 医嘱序号, 划价人, 操作员编号, 操作员姓名)
            Select 是否急诊, 结论, 记帐单id, 医嘱期效, 是否保密, 批次, v_费用id, 2, No_In, n_记录状态, v_费用序号, Null,
                   Decode(a.价格父号, Null, Null, v_费用序号 + a.价格父号 - a.序号), 1, a.病人id, a.主页id, a.标识号, a.姓名, a.性别, a.年龄,
                   a.病人病区id, a.病人科室id, a.费别, a.收费类别, a.收费细目id, a.计算单位, a.保险项目否, a.保险大类id, 1, -1 * v_当前数量, a.加班标志, a.附加标志,
                   a.婴儿费, a.收入项目id, a.收据费目, a.标准单价, Round(-1 * v_当前数量 * a.标准单价, v_Dec),
                   Round(-1 * v_当前数量 * a.标准单价, v_Dec), Null, 1, a.开单部门id, a.开单人, 收回时间_In, 收回时间_In, a.执行部门id, n_执行状态费用,
                   d_执行时间, v_执行人, a.医嘱序号, Decode(n_划价, 1, v_人员姓名, Null), Decode(n_划价, 1, Null, v_人员编号),
                   Decode(n_划价, 1, Null, v_人员姓名)
            From 门诊费用记录 A
            Where a.Id = Old_费用id;
        End If;
      
        --按理说这里应该调药品卫材服务计算 成本价,此处用 标准单价 做为成本价,影响不大
        Select Zl_Actualmoney_s(费别, 收费细目id, 收入项目id, 应收金额, 数次, 标准单价, 医嘱序号)
        Into v_Temp
        From 门诊费用记录
        Where ID = v_费用id;
        v_实收金额 := Round(Substr(v_Temp, Instr(v_Temp, ':') + 1), v_Dec);
        Update 门诊费用记录 A Set 实收金额 = v_实收金额 Where ID = v_费用id;
        v_结束序号 := v_费用序号;
        v_费用序号 := v_费用序号 + 1;
        n_类型     := 1;
      Else
      
        --住院费用记录
        -------------------------------------------------------------------------------------
        Select Nvl(Max(序号), 0) + 1
        Into v_费用序号
        From 住院费用记录
        Where 记录性质 = 2 And 记录状态 In (0, 1) And NO = No_In;
      
        --记录序号范围以处理汇总表
        If v_开始序号 Is Null Then
          v_开始序号 := v_费用序号;
        End If;
        v_结束序号 := v_费用序号;
      
        If v_诊疗类别 In ('5', '6', '7') Or n_跟踪在用医嘱 = 1 Then
          Select Nvl(Instr(v_划价类别, Decode(v_诊疗类别, '4', '4', '5')), 0) Into n_划价 From Dual;
          Insert Into 住院费用记录
            (是否急诊, 结论, 记帐单id, 医嘱期效, 是否保密, 批次, 领药部门id, ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 门诊标志, 病人id, 主页id, 标识号,
             姓名, 性别, 年龄, 床号, 病人病区id, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id, 付数, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id,
             收据费目, 标准单价, 应收金额, 实收金额, 统筹金额, 记帐费用, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 执行时间, 医嘱序号, 划价人, 操作员编号,
             操作员姓名)
            Select 是否急诊, 结论, 记帐单id, 医嘱期效, 是否保密, 批次, 领药部门id, v_费用id, 2, No_In, Decode(n_划价, 1, 0, 1), v_费用序号, Null, Null,
                   多病人单, 2, 病人id, 主页id, 标识号, 姓名, 性别, 年龄, 床号, 病人病区id, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id,
                   v_当前付数, -1 * v_当前数量, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, Round(-1 * v_当前付数 * v_当前数量 * 标准单价, v_Dec),
                   Round(-1 * v_当前付数 * v_当前数量 * 标准单价, v_Dec), Null, 1, 开单部门id, 开单人, 收回时间_In, 收回时间_In, 执行部门id, v_执行人,
                   n_执行状态费用, d_执行时间, 医嘱序号, Decode(n_划价, 1, v_人员姓名, Null), Decode(n_划价, 1, Null, v_人员编号),
                   Decode(n_划价, 1, Null, v_人员姓名)
            From 住院费用记录
            Where ID = Old_费用id;
        Else
          --非药品卫材医嘱
          --医嘱已执行，收回的费用也填为已执行：不包含药品和跟踪在用的卫材，因为实际发放表示执行
          --根据执行状态直接更新为记帐单，不用再单独审核记帐划价单
          n_执行状态 := v_医嘱执行;
          Select Decode(n_执行状态, 1, 1, Decode(Nvl(Instr(v_划价类别, v_收费类别), 0), 0, 1, 0))
          Into n_记录状态
          From Dual;
          Insert Into 住院费用记录
            (是否急诊, 结论, 记帐单id, 医嘱期效, 是否保密, 批次, 领药部门id, ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 门诊标志, 病人id, 主页id, 标识号,
             姓名, 性别, 年龄, 床号, 病人病区id, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id, 付数, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id,
             收据费目, 标准单价, 应收金额, 实收金额, 统筹金额, 记帐费用, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 执行时间, 执行人, 医嘱序号, 划价人, 操作员编号,
             操作员姓名)
            Select 是否急诊, 结论, 记帐单id, 医嘱期效, 是否保密, 批次, 领药部门id, v_费用id, 2, No_In, n_记录状态, v_费用序号, Null,
                   Decode(a.价格父号, Null, Null, v_费用序号 + a.价格父号 - a.序号), a.多病人单, 2, a.病人id, a.主页id, a.标识号, a.姓名, a.性别,
                   a.年龄, a.床号, a.病人病区id, a.病人科室id, a.费别, a.收费类别, a.收费细目id, a.计算单位, a.保险项目否, a.保险大类id, 1, -1 * v_当前数量,
                   a.加班标志, a.附加标志, a.婴儿费, a.收入项目id, a.收据费目, a.标准单价, Round(-1 * v_当前数量 * a.标准单价, v_Dec),
                   Round(-1 * v_当前数量 * a.标准单价, v_Dec), Null, 1, a.开单部门id, a.开单人, 收回时间_In, 收回时间_In, a.执行部门id, n_执行状态费用,
                   d_执行时间, v_执行人, a.医嘱序号, Decode(n_划价, 1, v_人员姓名, Null), Decode(n_划价, 1, Null, v_人员编号),
                   Decode(n_划价, 1, Null, v_人员姓名)
            From 住院费用记录 A
            Where a.Id = Old_费用id;
        End If;
      
        --按理说这里应该调药品卫材服务计算 成本价,此处用 标准单价 做为成本价,影响不大
        Select Zl_Actualmoney_s(费别, 收费细目id, 收入项目id, 应收金额, 数次, 标准单价, 医嘱序号)
        Into v_Temp
        From 住院费用记录
        Where ID = v_费用id;
        v_实收金额 := Round(Substr(v_Temp, Instr(v_Temp, ':') + 1), v_Dec);
        Update 住院费用记录 A Set 实收金额 = v_实收金额 Where ID = v_费用id;
        v_结束序号 := v_费用序号;
        v_费用序号 := v_费用序号 + 1;
        n_类型     := 2;
      End If;
    
      --处理费用相关汇总表
      For r_Money In c_Money(v_开始序号, v_结束序号) Loop
        --病人余额
        Update 病人余额
        Set 费用余额 = Nvl(费用余额, 0) + r_Money.实收金额
        Where 病人id = r_Money.病人id And 性质 = 1 And 类型 = n_类型;
      
        If Sql%RowCount = 0 Then
          Insert Into 病人余额
            (病人id, 性质, 类型, 费用余额, 预交余额)
          Values
            (r_Money.病人id, 1, n_类型, r_Money.实收金额, 0);
        End If;
      
        --病人未结费用
        Update 病人未结费用
        Set 金额 = Nvl(金额, 0) + r_Money.实收金额
        Where 病人id = r_Money.病人id And 主页id = r_Money.主页id And Nvl(病人病区id, 0) = Nvl(r_Money.病人病区id, 0) And
              Nvl(病人科室id, 0) = Nvl(r_Money.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(r_Money.开单部门id, 0) And
              Nvl(执行部门id, 0) = Nvl(r_Money.执行部门id, 0) And 收入项目id + 0 = r_Money.收入项目id And 来源途径 + 0 = n_类型;
      
        If Sql%RowCount = 0 Then
          Insert Into 病人未结费用
            (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
          Values
            (r_Money.病人id, r_Money.主页id, r_Money.病人病区id, r_Money.病人科室id, r_Money.开单部门id, r_Money.执行部门id, r_Money.收入项目id,
             n_类型, r_Money.实收金额);
        End If;
      End Loop;
    End Loop;
  End If;

  Json_Out := '{"output":{"code":1,"message":"成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Overdue_Recovery;
/
Create Or Replace Procedure Zl_Exsesvr_Getbillbytime
(
  Json_In  Varchar2,
  Json_Out Out Clob
) Is
  -------------------------------------------------------------------------------------------------
  --功能：按时间范围获取费用单据
  --入参：json格式
  --  input
  --    query_type          N 0 查询方式:0-获取药品医嘱费用单据，1-获取卫材医嘱费用单据
  --    fee_source          N 1 费用来源:0-不区分;1-门诊;2-住院
  --    start_time          C 1 开始时间，格式：yyyy-mm-dd hh24:mi:ss
  --    end_time            C 1 结束时间，格式：yyyy-mm-dd hh24:mi:ss
  --    exe_deptids         C 0 执行部门ID，多个用逗英文号分隔
  --    excp_exe_deptids    C 0 不包含的执行部门ID，多个用逗英文号分隔
  --出参：json格式
  --  output
  --    code                C 1 应答码：0-失败；1-成功
  --    message             C 1 应答消息：成功时返回成功信息，失败时返回具体的错误信息
  --    bill_nos            C 1 单据信息:格式：单据类型1:NO1,单据类型2:NO2,...
  --                            其中，单据类型: 1-收费处方;2-记帐单处方;3-记帐表处方
  -------------------------------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_查询类型 Number(2);
  n_费用来源 Number(2);
  d_开始时间 门诊费用记录.发生时间%Type;
  d_结束时间 门诊费用记录.发生时间%Type;

  v_执行部门id     Varchar2(32767);
  v_不含执行部门id Varchar2(32767);

  v_Output Varchar2(32767);
  c_Output Clob;
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_查询类型 := j_Json.Get_Number('query_type');
  n_费用来源 := j_Json.Get_Number('fee_source');
  d_开始时间 := To_Date(j_Json.Get_String('start_time'), 'yyyy-mm-dd hh24:mi:ss');
  d_结束时间 := To_Date(j_Json.Get_String('end_time'), 'yyyy-mm-dd hh24:mi:ss');

  v_执行部门id     := j_Json.Get_String('exe_deptids');
  v_不含执行部门id := j_Json.Get_String('excp_exe_deptids');

  If d_开始时间 Is Null Or d_结束时间 Is Null Then
    Json_Out := zlJsonOut('查询时间范围无效！');
    Return;
  End If;

  --0-获取药品医嘱费用单据
  If Nvl(n_查询类型, 0) = 0 Then
    --门诊
    If Nvl(n_费用来源, 0) = 0 Or n_费用来源 = 1 Then
      For r_费用 In (Select Distinct a.No, a.记录性质 As 单据类型
                   From 门诊费用记录 A
                   Where a.记录性质 In (1, 2) And a.医嘱序号 Is Not Null And a.发生时间 Between d_开始时间 And d_结束时间 And
                         a.收费类别 In ('5', '6', '7') And
                         (v_执行部门id Is Null Or Instr(',' || v_执行部门id || ',', ',' || Nvl(a.执行部门id, 0) || ',') > 0) And
                         (v_不含执行部门id Is Null Or Instr(',' || v_不含执行部门id || ',', ',' || Nvl(a.执行部门id, 0) || ',') = 0)) Loop
      
        If Length(v_Output) > 30000 Then
          If c_Output Is Not Null Then
            c_Output := c_Output || ',' || To_Clob(v_Output);
          Else
            c_Output := To_Clob(v_Output);
          End If;
          v_Output := Null;
        End If;
      
        If v_Output Is Null Then
          v_Output := r_费用.单据类型 || ':' || r_费用.No;
        Else
          v_Output := v_Output || ',' || r_费用.单据类型 || ':' || r_费用.No;
        End If;
      End Loop;
    End If;
  
    --住院
    If Nvl(n_费用来源, 0) = 0 Or n_费用来源 = 2 Then
      For r_费用 In (Select Distinct a.No, Decode(Nvl(a.多病人单, 0), 1, 3, 2) As 单据类型
                   From 住院费用记录 A
                   Where a.记录性质 In (1, 2) And a.医嘱序号 Is Not Null And a.发生时间 Between d_开始时间 And d_结束时间 And
                         a.收费类别 In ('5', '6', '7') And
                         (v_执行部门id Is Null Or Instr(',' || v_执行部门id || ',', ',' || Nvl(a.执行部门id, 0) || ',') > 0) And
                         (v_不含执行部门id Is Null Or Instr(',' || v_不含执行部门id || ',', ',' || Nvl(a.执行部门id, 0) || ',') = 0)) Loop
      
        If Length(v_Output) > 30000 Then
          If c_Output Is Not Null Then
            c_Output := c_Output || ',' || To_Clob(v_Output);
          Else
            c_Output := To_Clob(v_Output);
          End If;
          v_Output := Null;
        End If;
      
        If v_Output Is Null Then
          v_Output := r_费用.单据类型 || ':' || r_费用.No;
        Else
          v_Output := v_Output || ',' || r_费用.单据类型 || ':' || r_费用.No;
        End If;
      End Loop;
    End If;
  End If;

  --1-获取卫材医嘱费用单据
  If Nvl(n_查询类型, 0) = 1 Then
    --门诊
    If Nvl(n_费用来源, 0) = 0 Or n_费用来源 = 1 Then
      For r_费用 In (Select Distinct a.No, a.记录性质 As 单据类型
                   From 门诊费用记录 A
                   Where a.记录性质 In (1, 2) And a.医嘱序号 Is Not Null And a.发生时间 Between d_开始时间 And d_结束时间 And a.收费类别 = '4' And
                         (v_执行部门id Is Null Or Instr(',' || v_执行部门id || ',', ',' || Nvl(a.执行部门id, 0) || ',') > 0) And
                         (v_不含执行部门id Is Null Or Instr(',' || v_不含执行部门id || ',', ',' || Nvl(a.执行部门id, 0) || ',') = 0)) Loop
      
        If Length(v_Output) > 30000 Then
          If c_Output Is Not Null Then
            c_Output := c_Output || ',' || To_Clob(v_Output);
          Else
            c_Output := To_Clob(v_Output);
          End If;
          v_Output := Null;
        End If;
      
        If v_Output Is Null Then
          v_Output := r_费用.单据类型 || ':' || r_费用.No;
        Else
          v_Output := v_Output || ',' || r_费用.单据类型 || ':' || r_费用.No;
        End If;
      End Loop;
    End If;
  
    --住院
    If Nvl(n_费用来源, 0) = 0 Or n_费用来源 = 2 Then
      For r_费用 In (Select Distinct a.No, Decode(Nvl(a.多病人单, 0), 1, 3, 2) As 单据类型
                   From 住院费用记录 A
                   Where a.记录性质 In (1, 2) And a.医嘱序号 Is Not Null And a.发生时间 Between d_开始时间 And d_结束时间 And a.收费类别 = '4' And
                         (v_执行部门id Is Null Or Instr(',' || v_执行部门id || ',', ',' || Nvl(a.执行部门id, 0) || ',') > 0) And
                         (v_不含执行部门id Is Null Or Instr(',' || v_不含执行部门id || ',', ',' || Nvl(a.执行部门id, 0) || ',') = 0)) Loop
      
        If Length(v_Output) > 30000 Then
          If c_Output Is Not Null Then
            c_Output := c_Output || ',' || To_Clob(v_Output);
          Else
            c_Output := To_Clob(v_Output);
          End If;
          v_Output := Null;
        End If;
      
        If v_Output Is Null Then
          v_Output := r_费用.单据类型 || ':' || r_费用.No;
        Else
          v_Output := v_Output || ',' || r_费用.单据类型 || ':' || r_费用.No;
        End If;
      End Loop;
    End If;
  End If;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := c_Output || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"成功","bill_nos":"') || c_Output || To_Clob('"}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"成功","bill_nos":"' || v_Output || '"}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getbillbytime;
/


Create Or Replace Procedure Zl_Exsesvr_Outnewbill
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：医嘱发送生成费用单据，门诊收费单，门诊帐单，住院记帐单
  --入参：Json_In:格式
  --  input
  --    billtype              N  1  1-收费单，2-记帐单
  --    pati_id               N  1  病人id
  --    pati_pageid           N  0  主页id:主要是门诊留观病人记帐或住院病人门诊记帐时传入,其他可以不传入该接点
  --    sgin_no               C  1  门诊号
  --    pati_name             C  1  病人姓名
  --    pati_sex              C  1  性别
  --    pati_age              C  1  年龄
  --    fee_category          C  1  费别
  --    pati_deptid           N  1  病人科室id
  --    operator_name         C  1  操作员姓名 
  --    operator_code         C  1  操作员编码
  --    outpati_tag           N  0  门诊标识:1-门诊;3-就诊卡;4-体检不传时，缺省为1
  --    rgst_id               N  1  挂号id
  --    emg_sign              N  0  是否急诊
  --    charge_tag            N  1  是否划价:门诊记帐时传入，1-表示门诊记帐划价单;0-表示门诊记帐单
  --    placer                C  1  开单人
  --    plcdept_id            N  1  开单部门id
  --    happen_time           C  0  发生时间:不传时，以当前时间为准,格式为yyyy-mm-dd hh24:mi:ss
  --    create_time           C  0  登记时间:不传时，以当前时间为准,格式为yyyy-mm-dd hh24:mi:ss
  --    site_no               C  0  站点号:院区
  --    mdlpay_mode_name      C  0  医疗付款方式名称
  --    bill_list[]           C     单据列表
  --      fee_no              C     未传入该接点时，由系统自动生成。
  --      apply_id            C  1  申请ID:外部临床系统的申请ID,目前未存储，用于返回信息
  --      item_list[]      
  --        fitem_id          N  1  收费细目id
  --        packages_num      N  1  付数
  --        send_num          N  1  数次
  --        drug_price        N  1  实价药品或卫材价格:实价必传;其他不传入时以收费价目.现价为准.
  --        exe_deptid        N  1  执行部门id
  --        memo              C  1  摘要
  --        order_id          N  1  医嘱ID:ZLHIS内部的医嘱ID
  --        decoction_method  C  1  煎法
  --        morphology        C  1  中药形态
  --        bakstuff_batch    N     批次:针对高质卫材时有效。
  --        receipt_issecret  N  1  是否保密，0-不保密，1-保密
  --        excute_tag        N  1  执行状态:0-未执行;1-已执行    在发送申请后，自动发料的情况下，才有执行状态=1的情况
  --出参: Json_Out,格式如下
  --  output
  --    code                N 1 应答吗：0-失败；1-成功
  --    message             C 1 应答消息：失败时返回具体的错误信息
  --    bill_list[]             单据列表
  --      fee_no            C 1 NO
  --      apply_id          N 1 外部临床系统的申请ID，HIS系统目前未保存
  --      item_list[]           项目列表
  --      fee_id            N 1 费用ID
  --      order_id          N 1 医嘱ID
  --      fitem_id          N 1 收费细目id
  --      fee_amrcvb        N 1 应收金额
  --      fee_ampaib        N 1 实收金额
  ---------------------------------------------------------------------------
  j_Input    PLJson;
  j_Json     PLJson;
  j_Billlist Pljson_List;
  j_Jsonbill PLJson;
  j_Itemlist Pljson_List;
  j_Jsonitem PLJson;
  v_Output   Varchar2(32767);
  c_Output   Clob;
  v_Bill     Varchar2(32767);
  c_Bill     Clob;
  v_Err_Msg  Varchar2(255);
  Err_Item Exception;
  --   input
  n_单据类型     Number(2); --1-收费单，2-记帐单
  n_病人id       门诊费用记录.病人id%Type;
  n_主页id       门诊费用记录.主页id%Type;
  n_门诊号       门诊费用记录.标识号%Type;
  v_姓名         门诊费用记录.姓名%Type;
  v_性别         门诊费用记录.性别%Type;
  v_年龄         门诊费用记录.年龄%Type;
  v_费别         门诊费用记录.费别%Type;
  n_病人科室id   门诊费用记录.病人科室id%Type;
  v_操作员姓名   门诊费用记录.操作员姓名%Type;
  v_操作员编号   门诊费用记录.操作员编号%Type;
  n_门诊标识     Number(2); --1-门诊;3-就诊卡;4-体检不传时，缺省为1
  n_挂号id       Number(18);
  n_急诊         Number(2);
  n_划价         Number(2);
  v_开单人       门诊费用记录.开单人%Type;
  n_开单部门id   门诊费用记录.开单部门id%Type;
  v_发生时间     Varchar2(100);
  d_发生时间     门诊费用记录.发生时间%Type;
  v_登记时间     Varchar2(100);
  d_登记时间     门诊费用记录.登记时间%Type;
  v_院区         Varchar2(50);
  v_付款方式名称 门诊费用记录.付款方式%Type;
  --    bill_list[] 
  v_No     门诊费用记录.No%Type;
  v_申请id varchar2(100);
  --    item_list[]      
  n_收费细目id 门诊费用记录.收费细目id%Type;
  n_付数       门诊费用记录.付数%Type;
  n_数次       门诊费用记录.数次%Type;
  n_标准单价   门诊费用记录.标准单价%Type;
  n_单价       门诊费用记录.标准单价%Type;
  n_执行科室id 门诊费用记录.执行部门id%Type;
  v_摘要       门诊费用记录.摘要%Type;
  n_医嘱序号   门诊费用记录.医嘱序号%Type;
  v_煎法       门诊费用记录.结论%Type;
  v_中药形态   门诊费用记录.结论%Type;
  n_批次       门诊费用记录.批次%Type;
  n_保密       Number(2);
  n_执行状态   Number(2); --0-未执行;1-已执行    在发送申请后，自动发料的情况下，才有执行状态=1的情况
  n_序号       门诊费用记录.序号%Type;
  n_价格父号   门诊费用记录.价格父号%Type;
  v_收费类别   门诊费用记录.收费类别%Type;
  v_价格等级   Varchar2(100);
  v_普通等级   Varchar2(100);
  v_药品等级   Varchar2(100);
  v_卫材等级   Varchar2(100);
  v_Pricegrade Varchar2(500);
  n_应收金额   门诊费用记录.应收金额%Type;
  n_实收金额   门诊费用记录.实收金额%Type;
  v_Tmp        Varchar2(500);
  n_费用id     门诊费用记录.Id%Type;
  n_Count      Number(5);
  n_Money_Dec  Number(2); --金额小数
  n_Price_Dec  Number(2); --单价小数
Begin

  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_单据类型 := j_Json.Get_Number('billtype');
  n_病人id   := j_Json.Get_Number('pati_id');
  n_主页id   := j_Json.Get_Number('pati_pageid');
  n_门诊号   := j_Json.Get_String('sgin_no');

  v_姓名 := j_Json.Get_String('pati_name');
  v_性别 := j_Json.Get_String('pati_sex');
  v_年龄 := j_Json.Get_String('pati_age');
  v_费别 := j_Json.Get_String('fee_category');

  n_病人科室id := j_Json.Get_Number('pati_deptid');
  v_操作员姓名 := j_Json.Get_String('operator_name');
  v_操作员编号 := j_Json.Get_String('operator_code');

  n_门诊标识 := j_Json.Get_Number('outpati_tag');
  n_挂号id   := j_Json.Get_Number('rgst_id');
  n_急诊     := j_Json.Get_Number('emg_sign');
  n_划价     := j_Json.Get_Number('charge_tag');

  v_开单人 := j_Json.Get_String('placer');

  If v_开单人 Is Null Then
    v_Err_Msg := '没有传入开单人，请检查！';
    Raise Err_Item;
  End If;

  n_开单部门id := j_Json.Get_Number('plcdept_id');

  If Nvl(n_开单部门id, 0) = 0 Then
    v_Err_Msg := '没有传入开单部门id，请检查！';
    Raise Err_Item;
  End If;

  v_发生时间 := j_Json.Get_String('happen_time');
  If v_发生时间 Is Not Null Then
    d_发生时间 := To_Date(v_发生时间, 'yyyy-mm-dd hh24:mi:ss');
  Else
    d_发生时间 := Sysdate;
  End If;

  v_登记时间 := j_Json.Get_String('create_time');
  If v_登记时间 Is Not Null Then
    d_登记时间 := To_Date(v_登记时间, 'yyyy-mm-dd hh24:mi:ss');
  Else
    d_登记时间 := Sysdate;
  End If;

  v_院区         := j_Json.Get_String('site_no');
  v_付款方式名称 := j_Json.Get_String('mdlpay_mode_name');
  If Nvl(v_院区, '-') = '-' And Nvl(v_付款方式名称, '-') = '-' Then
    v_普通等级 := Null;
    v_药品等级 := Null;
    v_卫材等级 := Null;
  Else
    v_Pricegrade := Zl_Get_Pricegrade_s(v_院区, v_付款方式名称);
    For c_价格等级 In (Select Rownum As 序号, Column_Value As 价格等级 From Table(f_Str2List(v_Pricegrade, '|'))) Loop
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

  n_Count := 0;

  n_Money_Dec := zl_To_Number(Nvl(zl_GetSysParameter(9), '2'));
  n_Price_Dec := zl_To_Number(Nvl(zl_GetSysParameter(157), '5'));

  If Not j_Json.Exist('bill_list') Then
    v_Err_Msg := '未传入的【bill_list】节点，请检查！';
    Raise Err_Item;
  End If;

  j_Billlist := j_Json.Get_Pljson_List('bill_list');
  If j_Billlist.Count = 0 Then
    v_Err_Msg := '未传入的【bill_list】节点，请检查！';
    Raise Err_Item;
  End If;
  For I In 1 .. j_Billlist.Count Loop
  
    j_Jsonbill := PLJson();
    j_Jsonbill := PLJson(j_Billlist.Get(I));
  
    v_No := j_Jsonbill.Get_String('fee_no');
    If v_No Is Null Then
      If n_单据类型 = 1 Then
        v_No := Zl_Exse_Nextno(13, 0);
      Else
        v_No := Zl_Exse_Nextno(14, 0);
      End If;
    End If;
  
    v_申请id := j_Jsonbill.Get_String('apply_id');
	 
    n_序号   := 1;
    n_费用id := Null;
    n_Count  := n_Count + 1;
    v_Bill   := Null;
  
    If Not j_Jsonbill.Exist('item_list') Then
      v_Err_Msg := '未传入的【item_list】节点，请检查！';
      Raise Err_Item;
    End If;
  
    j_Itemlist := j_Jsonbill.Get_Pljson_List('item_list');
    If j_Itemlist.Count = 0 Then
      v_Err_Msg := '未传入的【item_list】节点，请检查！';
      Raise Err_Item;
    End If;
    For J In 1 .. j_Itemlist.Count Loop
      j_Jsonitem   := PLJson();
      j_Jsonitem   := PLJson(j_Itemlist.Get(J));
      n_收费细目id := j_Jsonitem.Get_Number('fitem_id');
      n_价格父号   := Null;
      Select Max(类别) Into v_收费类别 From 收费项目目录 Where ID = n_收费细目id;
      If v_收费类别 Is Null Then
        v_Err_Msg := '当前传入的收费细目ID无对应的收费项目目录记录，请检查！';
        Raise Err_Item;
      End If;
    
      If Instr(',5,6,7,', ',' || v_收费类别 || ',') > 0 Then
        v_价格等级 := v_药品等级;
      Elsif v_收费类别 = '4' Then
        v_价格等级 := v_卫材等级;
      Else
        v_价格等级 := v_普通等级;
      End If;
    
      n_付数 := j_Jsonitem.Get_Number('packages_num');
      If Nvl(n_付数, 0) = 0 Then
        n_付数 := 1;
      End If;
    
      n_数次 := j_Jsonitem.Get_Number('send_num');
    
      If Nvl(n_数次, 0) = 0 Then
        v_Err_Msg := '当前传入的数次为0，请检查！';
        Raise Err_Item;
      End If;
    
      n_标准单价   := j_Jsonitem.Get_Number('drug_price');
      n_执行科室id := j_Jsonitem.Get_Number('exe_deptid');
    
      If Nvl(n_执行科室id, 0) = 0 Then
        v_Err_Msg := '没有传入执行科室id，请检查！';
        Raise Err_Item;
      End If;
    
      v_摘要     := j_Jsonitem.Get_String('memo');
      n_医嘱序号 := j_Jsonitem.Get_Number('order_id');
      v_煎法     := j_Jsonitem.Get_String('decoction_method');
      v_中药形态 := j_Jsonitem.Get_String('morphology');
      n_批次     := j_Jsonitem.Get_Number('bakstuff_batch');
      n_保密     := j_Jsonitem.Get_Number('receipt_issecret');
      n_执行状态 := j_Jsonitem.Get_Number('excute_tag');
    
      If Nvl(n_医嘱序号, 0) = 0 Then
        n_医嘱序号 := Null;
      End If;
    
      For r_收入项目 In (Select a.Id As 收费细目id, b.收入项目id, c.名称, c.收据费目, b.现价, b.原价, b.加班加价率, b.附术收费率, b.缺省价格, a.计算单位, a.费用类型,
                            a.屏蔽费别, a.类别 As 收费类别
                     From 收费项目目录 A, 收费价目 B, 收入项目 C
                     Where b.收费细目id = a.Id And c.Id = b.收入项目id And Sysdate Between b.执行日期 And
                           Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And a.Id = n_收费细目id And
                           ((b.价格等级 Is Null And Nvl(v_价格等级, '-') = '-') Or
                           (b.价格等级 = v_价格等级 Or
                           (b.价格等级 Is Null And Not Exists
                            (Select 1
                               From 收费价目
                               Where b.收费细目id = 收费细目id And 价格等级 = v_价格等级 And Sysdate Between 执行日期 And
                                     Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))))))
                     Order By 收费细目id, 收入项目id) Loop
        If Nvl(n_标准单价, 0) = 0 Then
          n_单价 := r_收入项目.现价;
        Else
          n_单价 := n_标准单价;
        End If;
        n_单价     := Round(n_单价, n_Price_Dec);
        n_应收金额 := Round(Nvl(n_单价, 0) * Nvl(n_付数, 1) * n_数次, n_Money_Dec);
      
        If Nvl(r_收入项目.屏蔽费别, 0) = 1 Then
          n_实收金额 := n_应收金额;
        Else
          --获取实收金额
          v_Tmp      := Zl_Actualmoney_s(v_费别, n_收费细目id, r_收入项目.收入项目id, n_应收金额, n_数次, n_单价, n_医嘱序号);
          n_实收金额 := Round(zl_To_Number(Substr(v_Tmp, Instr(v_Tmp, ':') + 1)), n_Money_Dec);
        End If;
      
        Select 病人费用记录_Id.Nextval Into n_费用id From Dual;
      
        --保存门诊记帐单据或收费单据
        If n_单据类型 = 2 Then
          Zl_门诊记帐记录_Insert_s(v_No, n_序号, n_病人id, n_门诊号, v_姓名, v_性别, v_年龄, v_费别, 0, 0, n_病人科室id, n_开单部门id, v_开单人, Null,
                             r_收入项目.收费细目id, r_收入项目.收费类别, r_收入项目.计算单位, n_付数, n_数次, 0, n_执行科室id, n_价格父号, r_收入项目.收入项目id,
                             r_收入项目.收据费目, n_单价, n_应收金额, n_实收金额, d_发生时间, d_登记时间, n_划价, Null, v_操作员编号, v_操作员姓名, n_费用id,
                             Null, v_摘要, n_医嘱序号, n_门诊标识, v_中药形态, v_煎法, n_主页id, Null, n_批次, Null, n_挂号id, n_急诊, 1, n_保密);
        Else
          Zl_门诊划价记录_Insert_s(v_No, n_序号, n_病人id, n_主页id, n_门诊号, Null, v_姓名, v_性别, v_年龄, v_费别, 0, n_病人科室id, n_开单部门id,
                             v_开单人, Null, r_收入项目.收费细目id, r_收入项目.收费类别, r_收入项目.计算单位, Null, n_付数, n_数次, 0, n_执行科室id, n_价格父号,
                             r_收入项目.收入项目id, r_收入项目.收据费目, n_单价, n_应收金额, n_实收金额, d_发生时间, d_登记时间, v_操作员姓名, n_费用id, v_摘要,
                             n_医嘱序号, v_煎法, 1, Null, r_收入项目.费用类型, Null, Null, v_中药形态, Null, Null, n_批次, Null, n_挂号id,
                             n_急诊, 1, n_保密);
        
        End If;
      
        If n_价格父号 Is Null Then
          n_价格父号 := n_序号;
        End If;
        n_序号 := n_序号 + 1;
      
        v_Bill := v_Bill || ',{"fee_id":' || n_费用id;
        v_Bill := v_Bill || ',"order_id":' || zlJsonStr(n_医嘱序号, 1);
        v_Bill := v_Bill || ',"fitem_id":' || zlJsonStr(r_收入项目.收费细目id, 1);
        v_Bill := v_Bill || ',"fee_amrcvb":' || zlJsonStr(n_应收金额, 1);
        v_Bill := v_Bill || ',"fee_ampaib":' || zlJsonStr(n_实收金额, 1);
        v_Bill := v_Bill || '}';
      
        If Length(v_Bill) > 30000 Then
          If c_Bill Is Null Then
            c_Bill := Substr(v_Bill, 2);
          Else
            c_Bill := c_Bill || v_Bill;
          End If;
          v_Bill := Null;
        End If;
      End Loop;
    
      If n_费用id Is Null Then
        v_Err_Msg := '根据传入的【item_list】节点未找到有效收费项目，请检查！';
        Raise Err_Item;
      End If;
    
    End Loop;
  
    v_Output := Null;
    v_Output := v_Output || ',{"fee_no":"' || v_No || '"';
    v_Output := v_Output || ',"apply_id":' || zlJsonStr(v_申请id);
  
    If n_Count = 1 Then
      c_Output := Substr(v_Output, 2);
    Else
      c_Output := c_Output || v_Output;
    End If;
  
    If c_Bill Is Null Then
      c_Output := c_Output || ',"item_list":[' || Substr(v_Bill, 2) || ']';
    Else
      c_Output := c_Output || ',"item_list":[' || c_Bill || v_Bill || ']';
    End If;
  
    c_Output := c_Output || '}';
  
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"成功","bill_list":[' || c_Output || ']}}';

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zltools.Zljsonstr(SQLCode || SQLErrM) || '"}}';
End Zl_Exsesvr_Outnewbill;
/


Create Or Replace Procedure Zl_Exsesvr_Getwaitingpati
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能：获取待诊病人信息
  --入参：Json_In:格式
  --  input
  --    pati_id             N    病人ID:传入时，表示按该病人ID获取待诊信息
  --    exe_deptid          N    执行科室id:传入时，表示按执行部门Id获取待诊信息
  --出参: Json_Out,格式如下
  --  output
  --    code                N 1 应答吗：0-失败；1-成功
  --    message             C 1 应答消息：失败时返回具体的错误信息
  --    reg_list[]          C   挂号信息列表
  --      pati_id           N 1 病人ID
  --      pati_name         C 1 姓名
  --      pati_sex          C 1 性别
  --      pati_age          C 1 年龄
  --      insurance_type    C 1 险类
  --      insurance_name    C 1 险类名称
  --      reg_no            C 1 挂号no
  --      reg_id            C 1 挂号ID
  --      exe_deptid        N 1 执行部门id
  --      exer_id           N 1 医生ID
  --      exetr             C 1 医生
  --      outp_room_name    C   接诊诊室
  --      emg_sign          N   急诊标志
  ---------------------------------------------------------------------------
  j_Input        PLJson;
  j_Json         PLJson;
  v_Output       Varchar2(32767);
  c_Output       Clob;
  v_Para         Varchar2(100);
  n_病人id       门诊费用记录.病人id%Type;
  n_执行科室id   门诊费用记录.执行部门id%Type;
  n_普通有效天数 Number(2);
  n_急诊有效天数 Number(2);
Begin

  v_Para         := Nvl(zl_GetSysParameter(21), '11');
  n_普通有效天数 := To_Number(Substr(v_Para, 1, 1));
  n_急诊有效天数 := To_Number(Substr(v_Para, 2));

  --解析入参
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_病人id     := Nvl(j_Json.Get_Number('pati_id'), 0);
  n_执行科室id := Nvl(j_Json.Get_Number('exe_deptid'), 0);
  For r_待诊 In (Select a.病人id, a.姓名, a.性别, a.年龄, a.险类, b.名称 As 险类名称, a.No As 挂号no, a.Id As 挂号id, a.执行部门id, a.执行人 As 医生,
                      c.Id As 医生id, a.诊室, Nvl(a.急诊, 0) As 急诊
               From 病人挂号记录 A, 保险类别 B, 人员表 C
               Where Nvl(a.执行状态, 0) = 0 And a.险类 = b.序号(+) And (a.病人id = n_病人id Or n_病人id = 0) And
                     (a.执行部门id = n_执行科室id Or n_执行科室id = 0) And
                     ((a.登记时间 >= Trunc(Sysdate) - n_普通有效天数 And Nvl(a.急诊, 0) = 0) Or
                      (a.登记时间 >= Trunc(Sysdate) - n_急诊有效天数 And Nvl(a.急诊, 0) = 1)) And a.执行人 = c.姓名(+)) Loop
  
    zlJsonPutValue(v_Output, 'pati_id', r_待诊.病人id, 1, 1);
    zlJsonPutValue(v_Output, 'pati_name', r_待诊.姓名);
    zlJsonPutValue(v_Output, 'pati_sex', r_待诊.性别);
    zlJsonPutValue(v_Output, 'pati_age', r_待诊.年龄);
    zlJsonPutValue(v_Output, 'insurance_type', r_待诊.险类, 1);
    zlJsonPutValue(v_Output, 'insurance_name', r_待诊.险类名称);
    zlJsonPutValue(v_Output, 'reg_no', r_待诊.挂号no);
    zlJsonPutValue(v_Output, 'reg_id', r_待诊.挂号id, 1);
    zlJsonPutValue(v_Output, 'exe_deptid', r_待诊.执行部门id, 1);
    zlJsonPutValue(v_Output, 'exer_id', r_待诊.医生id, 1);
    zlJsonPutValue(v_Output, 'exetr', r_待诊.医生);
    zlJsonPutValue(v_Output, 'outp_room_name', r_待诊.诊室);
    zlJsonPutValue(v_Output, 'emg_sign', r_待诊.急诊, 1, 2);
  
    If Length(Nvl(v_Output, ' ')) > 30000 Then
      If c_Output Is Not Null Then
        c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
      Else
        c_Output := To_Clob(v_Output);
      End If;
      v_Output := Null;
    End If;
  
  End Loop;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"成功","reg_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"成功","reg_list":[' || v_Output || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zltools.Zljsonstr(SQLCode || SQLErrM) || '"}}';
End Zl_Exsesvr_Getwaitingpati;
/

Create Or Replace Procedure Zl_Exsesvr_Getusebillinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) Is
  ------------------------------------------------------------------------------------------------- 
  --功能：获取票据使用明细数据 
  --入参：json格式 
  --  input      
  --    occasion  N  1  业务场合:1-收费,2-预交(包含押金),3-结帐,4-挂号,5-就诊卡 
  --    inv_type  N  1  票种:1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡
  --    fee_nos  C  1  费用单据号,多个用逗号分离
  --      exits_history C
  --出参：json格式 
  --   output      
  --    code  C  1  应答码：0-失败；1-成功
  --    message  C  1  "应答消息：  成功时返回成功信息  失败时返回具体的错误信息"
  --    data[]  C  1  使用明细数据
  --      use_id  N  1  使用id
  --      invoice_no  C  1  发票号
  --      use_note  C  1  使用原因
  --      use_time  C  1  使用时间:yyyy-mm-dd hh24:mi:ss
  --      inv_user  C  1  发票使用人
  --      recv_id  C  1  领用ID
  ------------------------------------------------------------------------------------------------- 
  j_Input PLJson;
  j_Json  PLJson;

  n_业务场合 Number(2);
  v_Nos      Varchar2(32767);
  n_票种     Number(2);
  v_Output   Varchar2(32767);
  c_Output   Clob;
  n_Nomoved  Number(2);

  Cursor c_票据信息 Is(
    Select b.Id, b.号码 As 票据号, Decode(b.原因, 1, '正常发出', 2, '作废收回', 3, '重打发出', 4, '重打收回', 6, '红票发出') As 使用原因,
           To_Char(b.使用时间, 'YYYY-MM-DD HH24:MI') As 使用时间, b.使用人, b.领用id
    From 票据打印内容 A, 票据使用明细 B
    Where a.数据性质 = 5 And a.Id = b.打印id And a.No = '-' And b.票种 = 1);

  r_票据信息 c_票据信息%RowType;

  Type Ty_Invoce Is Ref Cursor;
  c_Invoice Ty_Invoce; --动态游标变量

  v_No 门诊费用记录.No%Type;
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_业务场合 := Nvl(j_Json.Get_Number('occasion'), 0);
  n_票种     := j_Json.Get_Number('inv_type');
  v_Nos      := j_Json.Get_String('fee_nos');

  If v_Nos Is Null Then
    Json_Out := zlJsonOut('未传入需要查询的费用单据!');
    Return;
  End If;

  --门诊标志：只有记帐才有，所以缺省为NULL
  If Instr(v_Nos, ',') > 0 Then
    v_Nos := Substr(v_Nos, 1, Instr(v_Nos, ',') - 1);
  Else
    v_No := v_Nos;
  End If;
  n_Nomoved := Zl_Fun_Checkinhistory(n_业务场合, v_No, Null);

  If Instr(v_Nos, ',') > 0 Then
    If Nvl(n_Nomoved, 0) = 1 Then
      Open c_Invoice For
        Select b.Id, b.号码 As 票据号, Decode(b.原因, 1, '正常发出', 2, '作废收回', 3, '重打发出', 4, '重打收回', 6, '红票发出') As 使用原因,
               To_Char(b.使用时间, 'YYYY-MM-DD HH24:MI') As 使用时间, b.使用人, b.领用id
        From H票据打印内容 A, H票据使用明细 B
        Where a.数据性质 = n_业务场合 And a.Id = b.打印id And a.No In (Select Column_Value From Table(f_Str2List(v_Nos))) And
              b.票种 = n_票种;
    
    Else
      Open c_Invoice For
        Select b.Id, b.号码 As 票据号, Decode(b.原因, 1, '正常发出', 2, '作废收回', 3, '重打发出', 4, '重打收回', 6, '红票发出') As 使用原因,
               To_Char(b.使用时间, 'YYYY-MM-DD HH24:MI') As 使用时间, b.使用人, b.领用id
        From 票据打印内容 A, 票据使用明细 B
        Where a.数据性质 = n_业务场合 And a.Id = b.打印id And a.No In (Select Column_Value From Table(f_Str2List(v_Nos))) And
              b.票种 = n_票种;
    End If;
  Else
    If Nvl(n_Nomoved, 0) = 1 Then
      Open c_Invoice For
        Select b.Id, b.号码 As 票据号, Decode(b.原因, 1, '正常发出', 2, '作废收回', 3, '重打发出', 4, '重打收回', 6, '红票发出') As 使用原因,
               To_Char(b.使用时间, 'YYYY-MM-DD HH24:MI') As 使用时间, b.使用人, b.领用id
        From H票据打印内容 A, H票据使用明细 B
        Where a.数据性质 = n_业务场合 And a.Id = b.打印id And a.No = v_Nos And b.票种 = n_票种;
    Else
      Open c_Invoice For
        Select b.Id, b.号码 As 票据号, Decode(b.原因, 1, '正常发出', 2, '作废收回', 3, '重打发出', 4, '重打收回', 6, '红票发出') As 使用原因,
               To_Char(b.使用时间, 'YYYY-MM-DD HH24:MI') As 使用时间, b.使用人, b.领用id
        From 票据打印内容 A, 票据使用明细 B
        Where a.数据性质 = n_业务场合 And a.Id = b.打印id And a.No = v_Nos And b.票种 = n_票种;
    End If;
  End If;

  --电子票据信息
  v_Output := Null;

  Loop
    Fetch c_Invoice
      Into r_票据信息;
    Exit When c_Invoice %NotFound;
  
    If v_Output Is Not Null Then
      v_Output := v_Output || ',';
    End If;
  
    --      use_id  N  1  使用id
    --      invoice_no  C  1  发票号
    --      use_note  C  1  使用原因
    --      use_time  C  1  使用时间:yyyy-mm-dd hh24:mi:ss
    --      inv_user  C  1  发票使用人
    --      recv_id  C  1  领用ID
    v_Output := v_Output || '{"use_id":' || zlJsonStr(r_票据信息.Id, 1);
    v_Output := v_Output || ',"invoice_no":"' || zlJsonStr(r_票据信息.票据号) || '"';
    v_Output := v_Output || ',"use_note":"' || zlJsonStr(r_票据信息.使用原因) || '"';
    v_Output := v_Output || ',"use_time":"' || zlJsonStr(r_票据信息.使用时间) || '"';
    v_Output := v_Output || ',"inv_user":"' || zlJsonStr(r_票据信息.使用人) || '"';
    v_Output := v_Output || ',"recv_id":' || zlJsonStr(Nvl(r_票据信息.领用id, 0), 1);
    v_Output := v_Output || '}';
    If Length(v_Output) > 30000 Then
      If c_Output Is Null Then
        c_Output := Substr(v_Output, 2);
      Else
        c_Output := c_Output || v_Output;
      End If;
      v_Output := Null;
    End If;
  End Loop;
  Close c_Invoice;
  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := c_Output || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;
  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"成功","data":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"成功","data":[' || v_Output || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getusebillinfo;
/